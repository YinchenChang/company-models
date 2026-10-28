"""Wattonomics API 工具共用函式（2026-10-28 起）。

- 呼叫 Anthropic API：Batch（五折）為主；批次逾 STALE_BATCH_H（預設 6）小時未完成就取消，未完成的請求改走一般 API。
- 用量：每次執行寫 usage.json（input／output tokens、web search 次數、估計 US$）。
- 預算：本批工具（evidence_audit、site_pipeline、weekly_monitor）累計上限 BATCH_BUDGET_USD（預設 85）；
  另以 PRIOR_SPEND_USD（本批以外已用，repo variable API_PRIOR_SPEND_USD）＋本批累計，接近 HARD_STOP_USD（預設 140）即停。
  累計值由所有遠端分支上的 tools/wattonomics_api/*/out/**/usage*.json 加總（checkout 需 fetch-depth: 0）。
- 任何請求都不帶使用者 email；SEC 一律用 GitHub 免回覆地址當聯絡方式。

只用標準函式庫。
"""
import datetime as dt
import json
import os
import re
import subprocess
import time
import urllib.error
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
MODEL = os.environ.get("WATT_MODEL", "claude-sonnet-5-5")
SEC_UA = "company-models research tool YinchenChang@users.noreply.github.com"   # SEC 要求聯絡方式；不用個人信箱
WEB_UA = "company-models wattonomics tool (github.com/YinchenChang/company-models)"
STALE_BATCH_H = float(os.environ.get("STALE_BATCH_H", "6"))
POLL_MAX_MIN = int(os.environ.get("POLL_MAX_MIN", "200"))
# 美元／百萬 token：一般 API（Batch 為一半）。Sonnet 5.5 單價沿用 tools/consistency/README.md 的 [Verified] 價目（2026-10-27）。
PRICE_STD = {"opus": (4, 20), "sonnet": (2, 10), "haiku": (0.1, 0.5)}
SEARCH_USD = 10 / 1000          # web search 每次 US$0.01（Batch 同價）
BATCH_BUDGET_USD = float(os.environ.get("BATCH_BUDGET_USD", "85"))
HARD_STOP_USD = float(os.environ.get("HARD_STOP_USD", "140"))
PRIOR_SPEND_USD = float(os.environ.get("PRIOR_SPEND_USD") or "0")
SAFETY_USD = 5.0                # 距 HARD_STOP 留的緩衝


# ---------- 用量與預算 ----------
class Usage:
    def __init__(self, tool, run):
        self.d = {"tool": tool, "run": run, "model": MODEL, "input_tokens": 0, "output_tokens": 0,
                  "cache_creation_input_tokens": 0, "cache_read_input_tokens": 0, "web_search_requests": 0,
                  "requests_batch": 0, "requests_standard": 0, "est_usd": 0.0,
                  "started_utc": dt.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")}

    def add(self, u, via="batch", model=MODEL):
        if not u:
            return
        pi, po = next(v for k, v in PRICE_STD.items() if k in model)
        f = 0.5 if via == "batch" else 1.0
        it = u.get("input_tokens", 0) or 0
        ot = u.get("output_tokens", 0) or 0
        cw = u.get("cache_creation_input_tokens", 0) or 0
        cr = u.get("cache_read_input_tokens", 0) or 0
        ws = (u.get("server_tool_use") or {}).get("web_search_requests", 0) or 0
        for k, v in (("input_tokens", it), ("output_tokens", ot), ("cache_creation_input_tokens", cw),
                     ("cache_read_input_tokens", cr), ("web_search_requests", ws)):
            self.d[k] += v
        self.d["requests_" + ("batch" if via == "batch" else "standard")] += 1
        self.d["est_usd"] += f * ((it + cw * 1.25 + cr * 0.1) * pi + ot * po) / 1e6 + ws * SEARCH_USD

    @property
    def usd(self):
        return self.d["est_usd"]

    def write(self, path, extra=None):
        self.d["est_usd"] = round(self.d["est_usd"], 4)
        self.d["finished_utc"] = dt.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
        if extra:
            self.d.update(extra)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        json.dump(self.d, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)


def spent_so_far():
    """所有遠端分支（含本機工作樹）上本批工具 usage*.json 的 est_usd 合計（同一 blob 只算一次）。"""
    seen, total = set(), 0.0

    def git(*a):
        return subprocess.run(["git", "-C", ROOT, *a], capture_output=True, text=True).stdout

    refs = [r for r in git("for-each-ref", "--format=%(refname)", "refs/remotes/origin").split() if not r.endswith("/HEAD")]
    for ref in refs + ["HEAD"]:
        for line in git("ls-tree", "-r", ref, "--", "tools/wattonomics_api").splitlines():
            meta, _, path = line.partition("\t")
            if not re.search(r"/out/.*usage[^/]*\.json$", path):
                continue
            blob = meta.split()[2]
            if blob in seen:
                continue
            seen.add(blob)
            try:
                total += float(json.loads(git("cat-file", "-p", blob)).get("est_usd", 0))
            except Exception:
                pass
    for dirpath, _, files in os.walk(f"{ROOT}/tools/wattonomics_api"):   # 本次工作樹尚未 commit 的
        for fn in files:
            p = os.path.join(dirpath, fn)
            if "/out/" in p and re.search(r"usage[^/]*\.json$", fn):
                blob = subprocess.run(["git", "hash-object", p], capture_output=True, text=True).stdout.strip()
                if blob and blob not in seen:
                    seen.add(blob)
                    try:
                        total += float(json.load(open(p)).get("est_usd", 0))
                    except Exception:
                        pass
    return total


def run_cap(requested_cap):
    """本次可用上限（美元）＝ min(請求上限, 本批剩餘, 距 140 停止線剩餘)。"""
    spent = spent_so_far()
    rem_batch = BATCH_BUDGET_USD - spent
    rem_hard = HARD_STOP_USD - SAFETY_USD - PRIOR_SPEND_USD - spent
    cap = min(requested_cap, rem_batch, rem_hard)
    print(f"預算：本批已用 US${spent:.2f}／上限 {BATCH_BUDGET_USD}；本批以外已用 {PRIOR_SPEND_USD}；本次上限 US${cap:.2f}")
    return cap, spent


# ---------- API ----------
def api(path, body=None, method="GET", timeout=300):
    req = urllib.request.Request(f"https://api.anthropic.com/v1/{path}",
                                 data=json.dumps(body).encode() if body is not None else None, method=method,
                                 headers={"x-api-key": API_KEY, "anthropic-version": "2023-06-01",
                                          "content-type": "application/json"})
    for attempt in range(4):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read().decode()
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503, 529) and attempt < 3:
                time.sleep(20 * (attempt + 1))
                continue
            raise RuntimeError(f"API HTTP {e.code}: {e.read()[:400].decode('utf-8', 'replace')}") from None


def text_of(msg):
    return "".join(c.get("text", "") for c in msg.get("content", []) if c.get("type") == "text")


def standard_message(params, usage, max_continue=3):
    """一般 API（非 Batch）。遇 pause_turn（長搜尋暫停）把助理訊息原樣送回續跑。回傳最後的 message。"""
    p = dict(params)
    msgs = list(p["messages"])
    for _ in range(max_continue + 1):
        p["messages"] = msgs
        m = json.loads(api("messages", p, "POST", timeout=900))
        usage.add(m.get("usage"), via="standard", model=p["model"])
        if m.get("stop_reason") != "pause_turn":
            return m
        msgs = msgs + [{"role": "assistant", "content": m["content"]}]
    return m


def continue_paused(params, msg, usage, max_continue=2):
    """Batch 結果為 pause_turn 時，以一般 API 接續。最終文字＝各段文字相接。"""
    p = dict(params)
    msgs = list(p["messages"]) + [{"role": "assistant", "content": msg["content"]}]
    texts = [text_of(msg)]
    for _ in range(max_continue):
        p["messages"] = msgs
        m = json.loads(api("messages", p, "POST", timeout=900))
        usage.add(m.get("usage"), via="standard", model=p["model"])
        texts.append(text_of(m))
        if m.get("stop_reason") != "pause_turn":
            break
        msgs = msgs + [{"role": "assistant", "content": m["content"]}]
    return "".join(texts)


def batch_age_h(b):
    created = dt.datetime.strptime(b["created_at"][:19], "%Y-%m-%dT%H:%M:%S")
    return (dt.datetime.utcnow() - created).total_seconds() / 3600


def collect_batch(b, params_by_id, usage, out):
    if not b.get("results_url"):
        return
    for line in api(f"messages/batches/{b['id']}/results", timeout=600).splitlines():
        r = json.loads(line)
        res = r["result"]
        cid = r["custom_id"]
        if res["type"] != "succeeded" or cid not in params_by_id:
            continue                       # 取消／過期／錯誤：由呼叫端改走一般 API
        m = res["message"]
        usage.add(m.get("usage"), via="batch", model=params_by_id[cid]["model"])
        if m.get("stop_reason") == "pause_turn":
            out[cid] = continue_paused(params_by_id[cid], m, usage)
        else:
            out[cid] = text_of(m)


def run_requests(params_by_id, usage, state=None, poll_max_min=None, stale_h=None):
    """params_by_id：{custom_id: Messages API 參數}。回傳 ({custom_id: 文字}, state)。
    state={"batch_id":…} 時續接該批次（不重送）。輪詢 poll_max_min 分鐘仍未完成：
      - 批次已逾 stale_h 小時 → 取消，已完成的取回，其餘改走一般 API；
      - 否則回傳 state（含 batch_id）讓呼叫端保留請求，下次執行續接。"""
    poll_max_min = POLL_MAX_MIN if poll_max_min is None else poll_max_min
    stale_h = STALE_BATCH_H if stale_h is None else stale_h
    out = {}
    if not params_by_id:
        return out, {"status": "ended"}
    state = dict(state or {})
    if state.get("batch_id"):
        b = json.loads(api(f"messages/batches/{state['batch_id']}"))
    else:
        reqs = [{"custom_id": cid, "params": p} for cid, p in params_by_id.items()]
        b = json.loads(api("messages/batches", {"requests": reqs}, "POST", timeout=600))
        state["batch_id"] = b["id"]
    print("batch", b["id"], b["processing_status"], len(params_by_id))
    t0 = time.time()
    while b["processing_status"] != "ended":
        if batch_age_h(b) > stale_h:
            print(f"批次 {b['id']} 已 {batch_age_h(b):.1f} 小時未完成：取消，未完成的改走一般 API")
            api(f"messages/batches/{b['id']}/cancel", {}, "POST")
            for _ in range(60):            # 等取消完成（已處理完的結果仍可取回）
                time.sleep(20)
                b = json.loads(api(f"messages/batches/{b['id']}"))
                if b["processing_status"] == "ended":
                    break
            break
        if time.time() - t0 > poll_max_min * 60:
            state["status"] = "pending"
            return out, state
        time.sleep(30)
        b = json.loads(api(f"messages/batches/{b['id']}"))
    collect_batch(b, params_by_id, usage, out)
    for cid, p in params_by_id.items():    # 批次失敗或被取消的請求：一般 API（約 2 倍價）
        if cid not in out:
            try:
                out[cid] = text_of(standard_message(p, usage))
            except Exception as e:
                out[cid] = f"__error__ {e}"
    state["status"] = "ended"
    return out, state


def web_search_tool(max_uses):
    return {"type": "web_search_20250305", "name": "web_search", "max_uses": int(max_uses)}


# ---------- 解析與輸出 ----------
def parse_jsonl(text):
    out = []
    for line in (text or "").splitlines():
        line = line.strip().strip("`").rstrip(",")
        if line.startswith("{") and line.endswith("}"):
            try:
                out.append(json.loads(line))
            except json.JSONDecodeError:
                pass
    if not out:
        m = re.search(r"\[\s*\{.*\}\s*\]", text or "", re.S)
        if m:
            try:
                out = json.loads(m.group(0))
            except json.JSONDecodeError:
                pass
    return [x for x in out if isinstance(x, dict)]


def evidence_dir(arg=None):
    """證據資料夾：參數 → wattonomics/evidence_raw → gwledger/evidence_raw（改名過渡期）。"""
    for d in [arg, "wattonomics/evidence_raw", "gwledger/evidence_raw"]:
        if d and os.path.isdir(f"{ROOT}/{d}"):
            return d
    raise RuntimeError("找不到 evidence_raw 資料夾（wattonomics/ 或 gwledger/）")


def pr_comment(body):
    """若本次執行的分支有開啟中的 PR，留言摘要（失敗不影響執行）。"""
    tok, repo, br = os.environ.get("GH_TOKEN"), os.environ.get("GITHUB_REPOSITORY"), os.environ.get("GITHUB_REF_NAME")
    if not (tok and repo and br):
        return
    h = {"Authorization": f"Bearer {tok}", "Accept": "application/vnd.github+json"}
    try:
        owner = repo.split("/")[0]
        req = urllib.request.Request(f"https://api.github.com/repos/{repo}/pulls?state=open&head={owner}:{br}", headers=h)
        prs = json.load(urllib.request.urlopen(req, timeout=60))
        for pr in prs[:1]:
            req = urllib.request.Request(f"https://api.github.com/repos/{repo}/issues/{pr['number']}/comments",
                                         data=json.dumps({"body": body[:60000]}).encode(), method="POST", headers=h)
            urllib.request.urlopen(req, timeout=60).read()
    except Exception as e:
        print("PR 留言失敗（不影響結果）：", e)

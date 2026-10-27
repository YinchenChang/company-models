"""跨模型數字一致性檢查（API 版，2026-10-27 起）。

只產出檢查結果，不改任何模型、公司檔案或網站；修正由各模型的維護串依工作單處理。

A. Tokenomics 快照落差（純程式，不用 API）：
   找出 repo 內所有 tokenomics_snapshot_v*.json，以版本最新者為基準，列出各資料夾快照的版本與「基準情境數值已改變」的名稱。
B. 公司檔案 vs 模型儲存格（API，Batch，Sonnet）：
   markets/invest_map/profiles/*.md 每句話後的出處標記「[本專案：<檔名>.xlsx <工作表>!<儲存格>…]」，
   由程式打開被引用的 Excel 讀出實際值（openpyxl，讀 LibreOffice 存下的計算值），
   再請模型逐句判定文字中的數字是否與儲存格值（含標記中寫明的推導，如 ÷、−）一致。

輸出：tools/consistency/out/<日期>.xlsx 與 .md（commit 回分支）。
"""
import datetime as dt
import glob
import json
import os
import re
import sys
import time
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
HERE = os.path.dirname(os.path.abspath(__file__))
API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
MODEL = os.environ.get("CHECK_MODEL", "claude-sonnet-5-5")
POLL_MAX_MIN = int(os.environ.get("POLL_MAX_MIN", "100"))


# ---------- A. 快照落差 ----------
def ver_key(p):
    m = re.search(r"_v(\d+)\.(\d+)", p)
    return (int(m.group(1)), int(m.group(2))) if m else (0, 0)


def base_values(snap):
    out = {}
    for name, it in snap.get("items", {}).items():
        v = it.get("values")
        if isinstance(v, dict):
            for gen, d in v.items():
                out[f"{name}｜{gen}"] = d.get("基準") if isinstance(d, dict) else d
        elif isinstance(v, list):
            out[name] = v[1] if len(v) == 3 else v
        else:
            out[name] = v
    return out


def snapshot_drift():
    files = [f for f in glob.glob(f"{ROOT}/**/tokenomics_snapshot_v*.json", recursive=True) if "/.git/" not in f]
    if not files:
        return [], None
    latest = max(files, key=ver_key)
    ref = base_values(json.load(open(latest, encoding="utf-8")))
    rows = []
    for f in sorted(files):
        vals = base_values(json.load(open(f, encoding="utf-8")))
        common = [k for k in vals if k in ref]
        changed = [k for k in common if vals[k] != ref[k]]
        rows.append({"folder": os.path.relpath(os.path.dirname(os.path.dirname(f)), ROOT), "file": os.path.basename(f),
                     "names": len(vals), "common": len(common), "changed": len(changed) if common else "格式不同，無法比對", "examples": "；".join(
                         f"{k}: {vals[k]} → {ref[k]}" for k in changed[:6])})
    return rows, os.path.basename(latest)


# ---------- B. 公司檔案 vs 儲存格 ----------
HOW = {}        # 儲存格值的來源統計：原檔／原檔（git 歷史）／同模型最新版／找不到
STALE = set()   # (公司, 引用檔名, 同模型最新版)：公司檔案引用的模型版本已落後
TAG = re.compile(r"\[本專案：([^\]]+)\]")
XLSX = re.compile(r"([^\s：；，、（）\[\]]+\.xlsx)")
CELL = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)")
_index, _wb = None, {}


STEM = re.compile(r"^(?:\d{8}_)?(.+?)_v[\d._]+\.xlsx$")


def _from_history(name):
    """工作樹沒有此檔時，從 git 歷史取出最後一次存在的版本（需完整歷史；Actions 以 fetch-depth: 0 取得）。"""
    import subprocess
    os.makedirs(f"{HERE}/_hist", exist_ok=True)
    dest = f"{HERE}/_hist/{name}"
    if os.path.exists(dest):
        return dest
    r = subprocess.run(["git", "-c", "core.quotepath=off", "-C", ROOT, "log", "--all", "--format=%H", "--name-only", "--diff-filter=AMR", "--", f"*{name}"],
                       capture_output=True, text=True)
    sha, path = None, None
    for line in r.stdout.splitlines():
        line = line.strip()
        if re.fullmatch(r"[0-9a-f]{40}", line):
            sha = line
        elif line.endswith(name) and sha:
            path = line
            break
    if not path:
        return None
    blob = subprocess.run(["git", "-C", ROOT, "show", f"{sha}:{path}"], capture_output=True)
    if blob.returncode != 0:
        return None
    open(dest, "wb").write(blob.stdout)
    return dest




def ver_tuple(n):
    m = re.search(r"_v([\d._]+)\.xlsx$", n)
    nums = tuple(int(x) for x in re.findall(r"\d+", m.group(1))) if m else ()
    d = re.match(r"^(\d{8})_", n)
    return nums, d.group(1) if d else ""


def find_xlsx(name):
    """回傳 (路徑, 來源說明, 最新版檔名)。優先用工作樹中的原檔；否則從 git 歷史取原檔；都沒有才用同模型最新版（儲存格位置可能已移動）。"""
    global _index
    name = os.path.basename(name)   # 出處標記有時寫成含路徑（例 skhynix/dist/…xlsx）
    if _index is None:
        _index = {}
        for f in glob.glob(f"{ROOT}/**/*.xlsx", recursive=True):
            if "/archive/" in f or "_待刪除" in f or "/_hist/" in f:
                continue
            _index.setdefault(os.path.basename(f), []).append(f)
    m = STEM.match(name)
    cands = [n for n in _index if m and (STEM.match(n) or [None, None])[1] == m.group(1)]
    newest = max(cands, key=ver_tuple) if cands else None   # 依版本數值比較（v10 > v9），不是字串比較
    if name in _index:
        return _index[name][0], "原檔", newest
    h = _from_history(name)
    if h:
        return h, "原檔（git 歷史）", newest
    if newest:
        return _index[newest][0], f"同模型最新版 {newest}（儲存格位置可能已移動）", newest
    return None, "找不到", None


def read_cells(path, sheet, ref):
    from openpyxl import load_workbook
    if path not in _wb:
        _wb[path] = load_workbook(path, data_only=True, read_only=True)
    wb = _wb[path]
    if sheet not in wb.sheetnames:
        return f"（無工作表 {sheet}）"
    ws = wb[sheet]
    ref = ref.replace("$", "")
    try:
        if ":" in ref:
            vals = [c.value for row in ws[ref] for c in row]
            return vals if len(vals) <= 40 else vals[:40] + ["…"]
        return ws[ref].value
    except Exception as e:
        return f"（讀取失敗 {type(e).__name__}）"


def profile_items(path):
    """回傳 [(行號, 原句, {引用: 值})]；只收出處標記中引用到 xlsx 儲存格的句子。"""
    items, last_file = [], None
    for ln, line in enumerate(open(path, encoding="utf-8"), 1):
        tags = TAG.findall(line)
        if not tags:
            continue
        cells = {}
        for t in tags:
            m = XLSX.search(t)
            if m:
                last_file = m.group(1)
            elif "同檔" not in t:
                continue
            if not last_file:
                continue
            fp, how, newest = find_xlsx(last_file)
            found_cells = CELL.findall(t)
            if found_cells:
                HOW[how.split("（")[0].split(" ")[0]] = HOW.get(how.split("（")[0].split(" ")[0], 0) + len(found_cells)
            if newest and newest != last_file and found_cells and ver_tuple(newest) > ver_tuple(last_file):
                STALE.add((os.path.basename(path)[:-3], last_file, newest))
            for sh, ref in CELL.findall(t):
                key = f"{last_file} {sh}!{ref}" + ("" if how.startswith("原檔") else f"〔{how}〕")
                cells[key] = read_cells(fp, sh, ref) if fp else "（repo 內找不到此檔或同模型任何版本）"
        if cells:
            items.append((ln, line.strip()[:600], cells))
    return items


SYS = """你是獨立核對者。給你一家公司「公司檔案」中的若干句子，每句後面附有出處標記，以及程式從被引用的 Excel 儲存格讀出的實際值。
逐句判定句中的數字是否與儲存格值一致：可以照出處標記寫明的推導（例如 ÷、−、＋、加總、換算百分比）驗算；四捨五入與單位換算（十億／百萬、小數／百分比）差異算一致。
判定只能是：一致／不一致／無法判定（儲存格是空值、找不到檔、或推導方式不明）。
輸出 JSON Lines，每句一行：{"line":行號,"verdict":判定,"text_value":句中數字,"cell_value":儲存格值或驗算結果,"note":理由（繁體中文，60 字內，寫出算式）}；不要其他文字。"""


def api(path, body=None, method="GET"):
    req = urllib.request.Request(f"https://api.anthropic.com/v1/{path}",
                                 data=json.dumps(body).encode() if body is not None else None, method=method,
                                 headers={"x-api-key": API_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json"})
    with urllib.request.urlopen(req, timeout=300) as r:
        return r.read().decode()


def run_batch(items):
    reqs = [{"custom_id": k, "params": {"model": MODEL, "max_tokens": 16000, "system": SYS,
                                       "messages": [{"role": "user", "content": v}]}} for k, v in items.items()]
    b = json.loads(api("messages/batches", {"requests": reqs}, "POST"))
    print("batch", b["id"], len(reqs))
    t0 = time.time()
    while b["processing_status"] != "ended" and time.time() - t0 < POLL_MAX_MIN * 60:
        time.sleep(30)
        b = json.loads(api(f"messages/batches/{b['id']}"))
    out, cost = {}, 0.0
    if b["processing_status"] == "ended":
        for line in api(f"messages/batches/{b['id']}/results").splitlines():
            r = json.loads(line)
            if r["result"]["type"] == "succeeded":
                m = r["result"]["message"]
                out[r["custom_id"]] = "".join(c.get("text", "") for c in m["content"])
                u = m.get("usage", {})
                cost += (u.get("input_tokens", 0) * 1 + u.get("output_tokens", 0) * 5) / 1e6
    return out, cost, b


def parse_lines(text):
    out = []
    for line in (text or "").splitlines():
        line = line.strip().rstrip(",")
        if line.startswith("{") and line.endswith("}"):
            try:
                out.append(json.loads(line))
            except json.JSONDecodeError:
                pass
    return out


def main():
    day = dt.datetime.utcnow().strftime("%Y%m%d")
    drift, latest = snapshot_drift()
    profiles = sorted(p for p in glob.glob(f"{ROOT}/markets/invest_map/profiles/*.md") if not os.path.basename(p).startswith("00_"))
    per = {}
    for p in profiles:
        it = profile_items(p)
        if it:
            per[os.path.basename(p)[:-3]] = it
    n_items = sum(len(v) for v in per.values())
    print("快照", len(drift), "公司檔案", len(per), "句子", n_items)
    results, cost, binfo = ({}, 0.0, {"processing_status": "skipped"})
    if API_KEY and per:
        payload = {}
        for tk, items in per.items():
            payload[tk] = f"公司代號 {tk}\n\n" + "\n\n".join(
                f"行 {ln}：{txt}\n儲存格值：{json.dumps(cells, ensure_ascii=False, default=str)[:3000]}" for ln, txt, cells in items)
        results, cost, binfo = run_batch(payload)

    rows = []
    for tk, items in per.items():
        verdicts = {v.get("line"): v for v in parse_lines(results.get(tk, ""))}
        for ln, txt, cells in items:
            v = verdicts.get(ln, {})
            rows.append([v.get("verdict", "未判定"), tk, ln, txt[:300], json.dumps(cells, ensure_ascii=False, default=str)[:300],
                         v.get("text_value"), v.get("cell_value"), v.get("note")])
    order = {"不一致": 0, "無法判定": 1, "未判定": 2, "一致": 3}
    rows.sort(key=lambda r: (order.get(r[0], 2), r[1], r[2]))

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "公司檔案vs模型"
    ws.append(["判定", "公司", "行", "公司檔案原句", "引用儲存格與值", "文字數字", "儲存格值／驗算", "理由"])
    for r in rows:
        ws.append([str(x) if isinstance(x, (list, dict)) else x for x in r])
        if r[0] == "不一致":
            for c in ws[ws.max_row]:
                c.fill = PatternFill("solid", fgColor="F8D7DA")
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.freeze_panes = "C2"
    for col, w in zip("ABCDEFGH", (10, 10, 6, 60, 50, 16, 20, 50)):
        ws.column_dimensions[col].width = w
    s2 = wb.create_sheet("TK快照落差")
    s2.append([f"基準：最新快照 {latest}（基準情境數值）"])
    s2.append(["資料夾", "快照檔", "名稱數", "可比對名稱數", "數值已變的名稱數", "例子（舊 → 新）"])
    for d in drift:
        s2.append([d["folder"], d["file"], d["names"], d["common"], d["changed"], d["examples"]])
    s4 = wb.create_sheet("引用版本落後")
    s4.append(["公司", "公司檔案引用的模型檔", "repo 中同模型最新版"])
    for row in sorted(STALE):
        s4.append(list(row))
    s3 = wb.create_sheet("說明與用量")
    for row in [["模型", MODEL], ["批次狀態", binfo.get("processing_status")], ["估計費用（美元）", round(cost, 2)],
                ["說明", "判定由模型產生，只作人工複核的優先順序；儲存格值讀自 repo 內 xlsx 存下的計算值。修正由各模型維護串依工作單處理。"]]:
        s3.append(row)
    os.makedirs(f"{HERE}/out", exist_ok=True)
    wb.save(f"{HERE}/out/{day}.xlsx")
    tally = {}
    for r in rows:
        tally[r[0]] = tally.get(r[0], 0) + 1
    lag = [d for d in drift if isinstance(d["changed"], int) and d["changed"] > 0]
    nocmp = [d for d in drift if isinstance(d["changed"], str)]
    md = [f"# 跨模型數字一致性檢查｜{day}", "",
          f"- A. Tokenomics 快照：{len(drift)} 份，基準 {latest}；數值已落後 {len(lag)} 份：" + ("、".join(f"{d['folder']}（{d['file'][20:-5]}：{d['changed']} 項）" for d in lag) or "無")
          + (f"；格式不同無法比對 {len(nocmp)} 份：" + "、".join(f"{d['folder']}（{d['file'][20:-5]}）" for d in nocmp) if nocmp else ""),
          f"- B. 公司檔案 {len(per)} 份、引用儲存格的句子 {n_items} 句；判定：" + "、".join(f"{k} {v}" for k, v in sorted(tally.items(), key=lambda x: order.get(x[0], 2))),
          f"- 引用版本落後：上述 {len(per)} 份中有 {len({x[0] for x in STALE} & set(per))} 份引用的模型檔已有新版（明細見「引用版本落後」頁）",
          "- 儲存格值來源：" + "、".join(f"{k} {v} 格" for k, v in sorted(HOW.items())),
          f"- 估計費用約 US${cost:.2f}（{MODEL}，Batch 五折）", "", "詳見同名 xlsx。判定只作人工複核的優先順序。"]
    open(f"{HERE}/out/{day}.md", "w", encoding="utf-8").write("\n".join(md) + "\n")
    print("\n".join(md))


if __name__ == "__main__":
    main()

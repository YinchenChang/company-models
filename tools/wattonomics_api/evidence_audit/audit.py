"""Wattonomics 證據逐列回查（API 版，2026-10-28 起；範本 Tokenomics tools/evidence_audit）。

逐列讀 evidence_raw/*.csv（A_us_cloud、B_labs_china_totals、B_delay、F_funding、P_facility_power…），
抓每列 url 的原文，請模型比對該列 value／unit／as_of／basis／tag（及其他欄）與原文，判定：
  屬實／小誤／不符／口徑不同／原文找不到；抓不到原文者記「無法取得」（不重試、不繞過；SEC 403 不重試）。
只產出核對結果，不改證據檔或 Excel。

請求檔 tools/wattonomics_api/evidence_audit/requests/<名>.json：
  {"evidence_dir": null, "files": [], "max_rows": null, "cost_cap_usd": 15}
  - evidence_dir：預設 wattonomics/evidence_raw，找不到用 gwledger/evidence_raw
  - files：只跑這些 csv（檔名）；空＝全部
  - max_rows：只取前 N 列（冒煙測試用，例 5）
輸出 out/<名>.csv、out/<名>.md、out/<名>_usage_<時間>.json；請求檔移到 requests/done/。
"""
import csv
import glob
import html
import html.parser
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common as C  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
TOOL = "evidence_audit"
MAX_SRC_CHARS = int(os.environ.get("AUDIT_MAX_SRC_CHARS", "120000"))   # 單一來源送模型的上限（約 3–4 萬 token）
DEFAULT_CAP = 15.0
VERDICTS = ["不符", "口徑不同", "小誤", "原文找不到", "無法取得", "未判定", "屬實"]


# ---------- 抓原文 ----------
class _Text(html.parser.HTMLParser):
    def __init__(self):
        super().__init__()
        self.out, self.skip = [], 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "noscript", "svg"):
            self.skip += 1
        elif tag in ("br", "p", "div", "tr", "li", "h1", "h2", "h3", "h4"):
            self.out.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style", "noscript", "svg"):
            self.skip = max(0, self.skip - 1)
        elif tag in ("td", "th"):
            self.out.append(" | ")

    def handle_data(self, d):
        if not self.skip:
            self.out.append(d)


def to_text(data, ctype):
    if data[:5] == b"%PDF-" or "pdf" in ctype:
        p = subprocess.run(["pdftotext", "-layout", "-", "-"], input=data, capture_output=True)
        return p.stdout.decode("utf-8", "replace")
    s = data.decode("utf-8", "replace")
    if "html" in ctype or s.lstrip()[:300].lower().startswith(("<!doctype", "<html")):
        p = _Text()
        p.feed(s)
        s = html.unescape("".join(p.out))
    return re.sub(r"\n\s*\n+", "\n", re.sub(r"[ \t\xa0]+", " ", s))


def fetch(url):
    """回傳 (狀態, 文字)。狀態 ok 或「無法取得（原因）」。一律不重試、不改用其他管道。"""
    u = (url or "").strip()
    if not u.lower().startswith(("http://", "https://")):
        # repo 內部來源（如 /root/company-models/microsoft/data/x.json 或相對路徑）：改讀 checkout 中的檔案
        rel = re.sub(r"^/?root/(company-models|Tokenomics)/", "", u).lstrip("/")
        p = os.path.join(os.environ.get("GITHUB_WORKSPACE", "."), rel)
        if rel and os.path.isfile(p):
            with open(p, "rb") as fh:
                return "ok", to_text(fh.read(25_000_000), "text/plain")
        return "無法取得（非網址，repo 內找不到）", ""
    url = u
    host = urllib.parse.urlparse(url).netloc.lower()
    ua = C.SEC_UA if host.endswith("sec.gov") else C.WEB_UA
    req = urllib.request.Request(url, headers={"User-Agent": ua, "Accept": "*/*"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            t = to_text(r.read(25_000_000), r.headers.get("Content-Type", ""))
            if len(t.strip()) < 300:
                return "無法取得（內容過短，可能需 JavaScript 或登入）", ""
            return "ok", t
    except urllib.error.HTTPError as e:
        return f"無法取得（HTTP {e.code}）", ""
    except Exception as e:
        return f"無法取得（{type(e).__name__}）", ""
    finally:
        time.sleep(0.3)


def window(text, row, limit):
    """原文過長時，保留開頭與含該列數值關鍵字附近的段落（降低 token）。"""
    if len(text) <= limit:
        return text
    keys = set()
    for k in ("value", "low", "base", "high", "low_months", "mid_months", "high_months"):
        v = str(row.get(k) or "").strip()
        if v and v not in ("-", "—"):
            keys.add(v)
            try:
                f = float(v.replace(",", ""))
                keys |= {f"{f:,.0f}", f"{f:g}", f"{f:,.1f}"}
            except ValueError:
                pass
    keys = {k for k in keys if len(k) >= 2}
    spans = []
    for k in keys:
        for m in re.finditer(re.escape(k), text):
            spans.append((max(0, m.start() - 1500), m.end() + 1500))
            if len(spans) > 40:
                break
    spans.sort()
    head = text[: limit // 4]
    parts, used = [head], len(head)
    for a, b in spans:
        if used >= limit:
            break
        seg = text[a:b]
        parts.append("\n…\n" + seg)
        used += len(seg)
    return "".join(parts)[:limit]


# ---------- 讀證據列 ----------
def load_rows(ev_dir, files, max_rows):
    rows = []
    paths = sorted(glob.glob(f"{C.ROOT}/{ev_dir}/*.csv"))
    if files:
        paths = [p for p in paths if os.path.basename(p) in files]
    for p in paths:
        with open(p, encoding="utf-8-sig", newline="") as f:
            for i, r in enumerate(csv.DictReader(f), start=2):
                r = {k: (v or "").strip() for k, v in r.items() if k}
                r["_file"], r["_line"] = os.path.basename(p), i
                u = r.get("url", "").split()[0].rstrip(".,;）)」") if r.get("url") else ""
                if u and not u.startswith("http") and re.match(r"^[\w.-]+\.[a-z]{2,}(/|$)", u):
                    u = "https://" + u
                r["_url"] = u
                rows.append(r)
    return rows[:max_rows] if max_rows else rows


SYS = """你是獨立核對者。給你一份來源原文，以及研究證據表（CSV）中引用這份來源的若干列。逐列判定該列是否得到原文支持。
逐欄比對：value（或 low／base／high、*_months）、unit、as_of（時點）、basis（IT／設施／未說明）、tag（[Verified]／[Interested-party]／[Derived]／[Assumed]／[Analogy]），以及 quote_or_paraphrase 中「本人說」的部分。
判定只能是下列之一：
- 屬實：數字、單位、時點、口徑都與原文一致（四捨五入差異算屬實，在 note 說明）。
- 小誤：主旨正確但有小錯（日期差幾天、四捨五入方式不同、標題抄錯、tag 稍偏），不影響模型數值。
- 不符：數字、單位或時點與原文不同，會影響模型數值。
- 口徑不同：數字有出現但口徑不同（IT vs 設施、年化 vs 全年、簽約 vs 上線、記者推論寫成本人說、[Verified] 但其實是利害關係人說法）。
- 原文找不到：原文中找不到該數字或說法（若可能在圖表圖片中，note 寫「可能在圖片中」）。
若列中數值是 [Derived]（由多個數字推算），只核對推算所用的原文數字是否存在。
輸出 JSON Lines，每列一行，不要其他文字：
{"id":列編號,"verdict":判定,"fields":不一致的欄位（逗號分隔，無則空字串）,"source_snippet":原文最相關片段（20 字以內）,"note":理由（繁體中文，60 字內）}"""


def build_items(rows, texts):
    by_url = {}
    for i, r in enumerate(rows):
        by_url.setdefault(r["_url"], []).append(i)
    items = {}
    for j, (url, idxs) in enumerate(by_url.items()):
        st, t = texts.get(url, ("", ""))
        if st != "ok":
            continue
        lst = "\n".join("- id=%d｜%s" % (i, json.dumps({k: v for k, v in rows[i].items() if not k.startswith("_") and v},
                                                         ensure_ascii=False)) for i in idxs)
        src = window(t, rows[idxs[0]], MAX_SRC_CHARS)
        items[f"u{j}"] = (f"來源網址：{url}\n\n## 證據表列（{len(idxs)} 列）\n{lst}\n\n## 來源原文\n{src}")
    return items


def est_cost(items):
    """事前估計（Batch 價）：輸入以 3 字元／token 計，輸出每則 1,500 token。"""
    pi, po = next(v for k, v in C.PRICE_STD.items() if k in C.MODEL)
    return sum((len(t) / 3 * pi + 1500 * po) / 2 / 1e6 for t in items.values())


def write_outputs(name, rows, texts, usage, req, skipped):
    os.makedirs(f"{HERE}/out", exist_ok=True)
    for r in rows:
        st = texts.get(r["_url"], ("無法取得（無網址）", ""))[0] if r["_url"] else "無法取得（無網址）"
        if st != "ok":
            r["verdict"], r["note"] = "無法取得", st
        r.setdefault("verdict", "未判定")
    rows.sort(key=lambda r: (VERDICTS.index(r["verdict"]) if r["verdict"] in VERDICTS else 5, r["_file"], r["_line"]))
    head = ["verdict", "file", "line", "player", "item", "value", "unit", "as_of", "basis", "tag", "url",
            "fields", "source_snippet", "note"]
    with open(f"{HERE}/out/{name}.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(head)
        for r in rows:
            item = r.get("site_or_scope") or r.get("site_or_contract") or r.get("item") or r.get("bottleneck") or ""
            sub = r.get("metric") or r.get("sub_item") or ""
            val = r.get("value") or "／".join(x for x in (r.get("low", "") or r.get("low_months", ""),
                                                          r.get("base", "") or r.get("mid_months", ""),
                                                          r.get("high", "") or r.get("high_months", "")) if x)
            w.writerow([r["verdict"], r["_file"], r["_line"], r.get("player") or r.get("company") or r.get("region", ""),
                        f"{item}｜{sub}" if sub else item, val, r.get("unit", ""), r.get("as_of", ""), r.get("basis", ""),
                        r.get("tag", ""), r["_url"], r.get("fields", ""), r.get("source_snippet", ""), r.get("note", "")])
    tally, by_file = {}, {}
    for r in rows:
        tally[r["verdict"]] = tally.get(r["verdict"], 0) + 1
        by_file.setdefault(r["_file"], {}).setdefault(r["verdict"], 0)
        by_file[r["_file"]][r["verdict"]] += 1
    ok = sum(1 for st, _ in texts.values() if st == "ok")
    md = [f"# Wattonomics 證據逐列回查｜{name}", "",
          f"- 資料夾：{req['_dir']}；列數 {len(rows)}；來源網址 {len(texts)} 個（可自動取得 {ok}）",
          "- 判定：" + "、".join(f"{k} {v}" for k, v in sorted(tally.items(), key=lambda x: VERDICTS.index(x[0]) if x[0] in VERDICTS else 9)),
          f"- 用量：input {usage.d['input_tokens']:,}／output {usage.d['output_tokens']:,} tokens；估計 US${usage.usd:.2f}（{C.MODEL}，Batch 五折）",
          ]
    if skipped:
        md.append(f"- **超出本次費用上限未送核對：{skipped} 個網址**（這些列判「未判定」）")
    md += ["", "| 檔案 | " + " | ".join(VERDICTS) + " |", "|---" * (len(VERDICTS) + 1) + "|"]
    md += [f"| {f} | " + " | ".join(str(t.get(v, 0)) for v in VERDICTS) + " |" for f, t in sorted(by_file.items())]
    md += ["", "## 需處理（不符／口徑不同，前 60 列）", "", "| 判定 | 檔案:列 | 項目 | 值 | 理由 |", "|---|---|---|---|---|"]
    for r in [r for r in rows if r["verdict"] in ("不符", "口徑不同")][:60]:
        item = r.get("site_or_scope") or r.get("site_or_contract") or r.get("item") or r.get("bottleneck") or ""
        md.append(f"| {r['verdict']} | {r['_file']}:{r['_line']} | {item[:40]} | {r.get('value', '')[:20]} {r.get('unit', '')[:12]} | {str(r.get('note', ''))[:60]} |")
    md += ["", "判定由模型產生，只作人工複核的優先順序；「無法取得」依規則不繞過，須人工開原文。詳見同名 csv。"]
    open(f"{HERE}/out/{name}.md", "w", encoding="utf-8").write("\n".join(md) + "\n")
    return tally, "\n".join(md[:6])


def main():
    if not C.API_KEY:
        print("未設定 ANTHROPIC_API_KEY，結束。")
        return
    for path in sorted(glob.glob(f"{HERE}/requests/*.json")):
        name = os.path.splitext(os.path.basename(path))[0]
        req = json.load(open(path, encoding="utf-8"))
        stamp = time.strftime("%Y%m%dT%H%M%SZ", time.gmtime())
        usage = C.Usage(TOOL, name)
        try:
            cap, _ = C.run_cap(float(req.get("cost_cap_usd", DEFAULT_CAP)))
            if cap <= 0.2:
                raise RuntimeError(f"預算已用盡（本次可用 US${cap:.2f}），未執行")
            req["_dir"] = C.evidence_dir(req.get("evidence_dir"))
            rows = load_rows(req["_dir"], req.get("files") or [], req.get("max_rows"))
            urls = list(dict.fromkeys(r["_url"] for r in rows if r["_url"]))
            print(name, "列", len(rows), "網址", len(urls))
            texts = {u: fetch(u) for u in urls}
            items = build_items(rows, texts)
            skipped = 0
            while items and est_cost(items) > cap:         # 超出上限：從最後的網址往前砍
                items.pop(list(items)[-1])
                skipped += 1
            print("送核對", len(items), "則；事前估計 US$%.2f；上限 %.2f" % (est_cost(items), cap))
            params = {cid: {"model": C.MODEL, "max_tokens": 8000, "system": SYS,
                            "messages": [{"role": "user", "content": t}]} for cid, t in items.items()}
            res, state = C.run_requests(params, usage, req.get("_state"))
            if state.get("status") == "pending":
                req["_state"] = state
                json.dump(req, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
                open(f"{HERE}/out/{name}.md", "w", encoding="utf-8").write(
                    f"# 回查處理中｜{name}\n\n批次 {state['batch_id']} 尚未完成；請求檔已記下批次編號，重新觸發會續接，不重送。\n")
                usage.write(f"{HERE}/out/{name}_usage_{stamp}.json", {"status": "pending"})
                continue
            os.makedirs(f"{HERE}/out/raw", exist_ok=True)
            json.dump(res, open(f"{HERE}/out/raw/{name}.json", "w", encoding="utf-8"), ensure_ascii=False)
            for txt in res.values():
                for v in C.parse_jsonl(txt):
                    try:
                        rows[int(v["id"])].update({"verdict": v.get("verdict"), "fields": v.get("fields", ""),
                                                   "source_snippet": v.get("source_snippet", ""), "note": v.get("note", "")})
                    except (KeyError, ValueError, IndexError, TypeError):
                        pass
            tally, summary = write_outputs(name, rows, texts, usage, req, skipped)
            usage.write(f"{HERE}/out/{name}_usage_{stamp}.json", {"status": "done", "rows": len(rows), "urls": len(urls),
                                                                  "verdicts": tally})
            print(name, tally, f"US${usage.usd:.2f}")
            C.pr_comment(summary + f"\n\n結果：`tools/wattonomics_api/evidence_audit/out/{name}.csv`（參考用，不阻擋合併）")
            os.makedirs(f"{HERE}/requests/done", exist_ok=True)
            shutil.move(path, f"{HERE}/requests/done/{name}.json")
        except Exception as e:
            os.makedirs(f"{HERE}/out", exist_ok=True)
            open(f"{HERE}/out/{name}.md", "w", encoding="utf-8").write(
                f"# 回查失敗｜{name}\n\n{type(e).__name__}: {e}\n\n請求檔保留在 requests/，修正後重推即重跑。\n")
            usage.write(f"{HERE}/out/{name}_usage_{stamp}.json", {"status": "failed", "error": str(e)[:300]})
            print("失敗", name, e)


if __name__ == "__main__":
    main()

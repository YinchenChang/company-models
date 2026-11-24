# 外部稽核優化 W3-6（2026-11-23）：治理項目盤點。
# 對每個 <資料夾>/dist/ 內的現行成品跑 tools/model_governance/check_governance.py（含 --html）與
# tools/sens_coverage/check_sens_coverage.py（--require-keycol），把失敗項目歸類成一張表。只讀檔、不改任何成品。
# 用法（repo 根目錄）：python3 tools/audit_w3_agent/inventory.py [--md 報告.md] [--json 明細.json] [--jobs 8]
#   check_governance.py 每頁只列前 5 個標記不合格的格，所以「標記」欄的格數由本檔另外算（同一套判斷）。
import glob, json, os, re, subprocess, sys
from concurrent.futures import ProcessPoolExecutor
import openpyxl

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
sys.path.insert(0, os.path.join(ROOT, "tools", "sens_coverage"))
GOV = os.path.join(ROOT, "tools", "model_governance", "check_governance.py")
TAGS = {"Verified", "Interested-party", "Analogy", "Assumed", "Derived", "—"}


def tag_count(ws):
    """回傳（標記不是單一主要標記的格數, 有無「來源誘因」欄, Interested-party 列缺誘因文字的格數）；找不到「標記」欄回傳 None"""
    for r in range(1, 13):
        vals = [str(ws.cell(r, c).value or "") for c in range(1, ws.max_column + 1)]
        if "標記" in vals:
            tc = vals.index("標記") + 1
            ic = next((i + 1 for i, v in enumerate(vals) if "來源誘因" in v), None)
            bad = noinc = 0
            for rr in range(r + 1, ws.max_row + 1):
                t = ws.cell(rr, tc).value
                if t is None or str(t) == "標記": continue
                if str(t) not in TAGS: bad += 1
                elif str(t) == "Interested-party" and ic and not ws.cell(rr, ic).value: noinc += 1
            return bad, ic is not None, noinc
    return None


def one(co):
    import check_sens_coverage as SC
    xs = sorted(glob.glob(os.path.join(ROOT, co, "dist", "*.xlsx")))
    x = xs[-1]; hs = sorted(glob.glob(os.path.join(ROOT, co, "dist", "*.html")))
    o = {"co": co, "file": os.path.basename(x), "n_xlsx": len(xs)}
    p = subprocess.run([sys.executable, GOV, x] + (["--html", hs[-1]] if hs else []), capture_output=True, text=True)
    out = p.stdout + p.stderr; o["gov_rc"] = p.returncode
    o["gov_na"] = "不適用" in out
    o["gov_err"] = "" if p.returncode in (0, 1) and "治理核對" in out else out.strip().splitlines()[-1][:120] if out.strip() else "無輸出"
    o["no_bid"] = "沒有 build_id 列" in out or "沒有 Sens 頁" in out
    o["bid_readme"] = "README!A1 未含" in out; o["bid_html"] = "HTML 未含 build_id" in out
    m = re.search(r"Sens 模型狀態＝(.+?)（應為", out); o["status"] = m.group(1) if m else None
    o["verified_conflict"] = out.count("標 Verified，但來源事實總帳")
    wb = openpyxl.load_workbook(x, data_only=True)
    for sh in ("In", "Q"):
        o[sh] = tag_count(wb[sh]) if sh in wb.sheetnames else "無此頁"
    r = SC.check(x, None, require_keycol=True)
    o["sens_app"] = r["applicable"]; o["sens_err"] = r["error"]; o["keys"] = r["keys"]; o["keycol"] = r["keycol"]
    o["missing"] = [m_[0] for m_ in r["missing"]]; o["exempt"] = len(r["exempt"])
    o["he_missing"] = sum(m_[5].startswith("IF_HoldEcon") for m_ in r["missing"])
    o["has_sens"] = "Sens" in wb.sheetnames
    o["posts"] = sorted(os.path.basename(f) for f in glob.glob(os.path.join(ROOT, co, "*_post.py")))
    return o


def cell_tag(v):
    if v == "無此頁": return "無此頁", "—", "—"
    if v is None: return "找不到「標記」欄", "—", "—"
    bad, hasinc, noinc = v
    return (f"{bad} 格" if bad else "通過"), ("有" if hasinc else "缺"), (f"{noinc} 格" if noinc else ("通過" if hasinc else "—"))


def main(argv):
    jobs = int(argv[argv.index("--jobs") + 1]) if "--jobs" in argv else 8
    cos = sorted(os.path.basename(os.path.dirname(d)) for d in glob.glob(os.path.join(ROOT, "*", "dist")) if glob.glob(os.path.join(d, "*.xlsx")))
    with ProcessPoolExecutor(jobs) as ex:
        res = list(ex.map(one, cos))
    L = ["| 資料夾 | 成品 | build_id 列 | 模型狀態 | In 標記 | In 來源誘因欄 | Q 標記 | Q 來源誘因欄 | 誘因文字空白 | 治理結果 | Sens 表頭 | Sens 鍵欄 | 未涵蓋鍵／待涵蓋 | 覆蓋結果 | *_post.py |",
         "|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|"]
    T = {}
    def bump(k, co): T.setdefault(k, []).append(co)
    for o in res:
        co = o["co"]
        if o["gov_na"]:
            gov = "不適用（無 In 頁）"; bid = st = "—"; it = iq = ("—", "—", "—")
            bump("治理：不適用（模型商結構，無 In 頁）", co)
        else:
            bid = "缺" if o["no_bid"] else ("README 不同" if o["bid_readme"] else ("HTML 不同" if o["bid_html"] else "有"))
            st = "通過" if o["status"] is None else f"{o['status']}"
            it, iq = cell_tag(o["In"]), cell_tag(o["Q"])
            gov = "通過" if o["gov_rc"] == 0 else ("執行錯誤：" + o["gov_err"] if o["gov_err"] else "失敗")
            if o["gov_rc"] == 0: bump("治理：全部通過", co)
            if o["no_bid"]: bump("G1 Sens 頁沒有 build_id 列", co)
            if o["bid_readme"] or o["bid_html"]: bump("G1b build_id 在 README 或 HTML 對不上", co)
            if o["status"] is not None: bump("G2 Sens 沒有模型狀態列或不是 CURRENT", co)
            if it[0] not in ("通過", "無此頁") or iq[0] not in ("通過", "無此頁"): bump("G3 In 或 Q 頁標記不是單一主要標記（含找不到標記欄）", co)
            if it[1] == "缺" or iq[1] == "缺": bump("G4 In 或 Q 頁沒有「來源誘因」欄", co)
            if (it[2].endswith("格")) or (iq[2].endswith("格")) or o["verified_conflict"]: bump("G5 有誘因欄但 Interested-party 列空白，或 Verified 與來源總帳不符", co)
        if not o["sens_app"]:
            sh = sk = "—"; cov = "不適用"; mk = "—"; bump("覆蓋：不適用（In 頁沒有低／基準／高三欄）", co)
        else:
            e = o["sens_err"] or ""
            sh = "無 Sens 頁" if "無 Sens 頁" in e else ("找不到" if "找不到表頭" in e else "有")
            sk = "—" if sh != "有" else ("有" if o["keycol"] else "缺")
            mk = f"{len(o['missing'])}／{o['keys']}"
            cov = "通過" if not o["missing"] and not e else "失敗"
            if cov == "通過": bump("覆蓋：通過（含鍵欄）", co)
            if sh == "無 Sens 頁": bump("S0 適用但沒有 Sens 頁", co)
            if sh == "找不到": bump("S1 Sens 頁前 40 列找不到含「目標價」欄名的表頭", co)
            if sk == "缺": bump("S2 有表頭但沒有「鍵」欄", co)
            if sh == "有" and o["missing"]: bump("S3 有表頭，仍有鍵沒有案例也沒有合格豁免", co)
            if o["he_missing"]: bump("S4 IF_HoldEcon 群組缺「整組同時」案例（不得豁免）", co)
        L.append(f"| `{co}` | {o['file']} | {bid} | {st} | {it[0]} | {it[1]} | {iq[0]} | {iq[1]} | "
                 f"{'／'.join(x for x in (it[2], iq[2]) if x != '—') or '—'} | {gov} | {sh} | {sk} | {mk} | {cov} | {len(o['posts'])} |")
    S = ["| 失敗類型 | 家數 | 資料夾 |", "|---|---|---|"] + [f"| {k} | {len(v)} | {'、'.join(v)} |" for k, v in sorted(T.items())]
    txt = "\n".join(S) + "\n\n" + "\n".join(L) + "\n"
    print(txt)
    if "--md" in argv: open(argv[argv.index("--md") + 1], "w", encoding="utf-8").write(txt)
    if "--json" in argv: json.dump(res, open(argv[argv.index("--json") + 1], "w", encoding="utf-8"), ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main(sys.argv[1:])

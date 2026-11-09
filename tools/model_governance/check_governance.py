# 個股模型治理核對（2026-11-09 SpaceX v0.2.5 驗收衍生原則：docs/plan/個股模型原則_命題加營收分部.md 第 10 節）
#   1. build_id：Sens 頁有「build_id」列，且 README 頁 A1 含同一字串（HTML 由 html.json 讀同一格，另由 --html 檢查）。
#   2. 模型狀態：Sens 頁「模型狀態」列＝CURRENT（STALE＝In 頁改過但 Sens 未重跑）。
#   3. 標記：In 頁 G 欄、Q 頁「標記」欄每格只能是單一主要標記（Verified／Interested-party／Analogy／Assumed／Derived／—）；
#      來源字串以「…；Derived）」或「…；Interested-party）」結尾（事實總帳標記）時，主要標記不得為 Verified。
#   4. Interested-party 列須有「來源誘因」欄文字（欄名含「來源誘因」者；沒有此欄＝失敗）。
# 用法：python3 tools/model_governance/check_governance.py <重算後 xlsx> [--html <成品 html>]；失敗時結束碼 1。
import re, sys
import openpyxl

TAGS = {"Verified", "Interested-party", "Analogy", "Assumed", "Derived", "—"}
args = sys.argv[1:]
X = args[0]; HTML = args[args.index("--html") + 1] if "--html" in args else None
wb = openpyxl.load_workbook(X, data_only=True)
fail = []


def row_of(ws, prefix):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith(prefix): return r
    return None


se = wb["Sens"] if "Sens" in wb.sheetnames else None
bid = None
if se is None: fail.append("沒有 Sens 頁")
else:
    rb = row_of(se, "build_id"); rs = row_of(se, "模型狀態")
    bid = se.cell(rb, 4).value if rb else None
    st = se.cell(rs, 4).value if rs else None
    if not bid: fail.append("Sens 頁沒有 build_id 列（D 欄）")
    if st != "CURRENT": fail.append(f"Sens 模型狀態＝{st!r}（應為 CURRENT；STALE 表示 In 頁改過、須重跑建置與敏感度）")
    if bid and bid not in str(wb["README"]["A1"].value): fail.append("README!A1 未含 Sens 的 build_id")
    if bid and HTML and bid not in open(HTML, encoding="utf-8").read(): fail.append("HTML 未含 build_id")


def tagcheck(ws, tagcol_name, srccol_name):
    hr = None
    for r in range(1, 13):   # 2026-11-09 Arm v0.3.4：apply_in_layout 在 In 頁表頭前加說明列（表頭在第 9 列）
        vals = [str(ws.cell(r, c).value or "") for c in range(1, ws.max_column + 1)]
        if tagcol_name in vals: hr = r; break
    if hr is None: fail.append(f"{ws.title} 頁找不到「{tagcol_name}」欄"); return
    vals = [str(ws.cell(hr, c).value or "") for c in range(1, ws.max_column + 1)]
    tc = vals.index(tagcol_name) + 1; sc = vals.index(srccol_name) + 1 if srccol_name in vals else None
    ic = next((i + 1 for i, v in enumerate(vals) if "來源誘因" in v), None)
    if ic is None: fail.append(f"{ws.title} 頁沒有「來源誘因」欄")
    n = 0
    for r in range(hr + 1, ws.max_row + 1):
        t = ws.cell(r, tc).value
        if t is None or str(t) in ("標記",): continue
        t = str(t)
        if t not in TAGS:
            n += 1
            if n <= 5: fail.append(f"{ws.title}!{ws.cell(r, tc).coordinate} 標記「{t}」不是單一主要標記")
            continue
        src = str(ws.cell(r, sc).value or "") if sc else ""
        m = re.search(r"；(Derived|Interested-party)）$", src)
        if m and t == "Verified": fail.append(f"{ws.title}!{ws.cell(r, tc).coordinate} 標 Verified，但來源事實總帳為 {m.group(1)}")
        if t == "Interested-party" and ic and not ws.cell(r, ic).value: fail.append(f"{ws.title}!{ws.cell(r, tc).coordinate} Interested-party 列缺來源誘因")


tagcheck(wb["In"], "標記", "來源")
if "Q" in wb.sheetnames: tagcheck(wb["Q"], "標記", "來源")
print(f"治理核對：build_id {bid}；" + ("全部通過" if not fail else f"失敗 {len(fail)} 項"))
for f in fail[:30]: print("  -", f)
sys.exit(1 if fail else 0)

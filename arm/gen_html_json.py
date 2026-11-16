# 產生 arm/html.json（HTML 展示層設定；v0.3 起先組舊版設定〔寫 arm/html_src.json，含標準版 DCF 區塊〕再以 tools/company_html/relayout.py 轉成個股 HTML 版面規格 v1.0）。用法：ARM_ROWMAP=rowmap.json python3 arm/build_arm.py raw.xlsx（取得列號），重算成 dist xlsx 後：python3 arm/gen_html_json.py <dist xlsx> arm/html.json rowmap.json
import json, os, re, sys, openpyxl
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "company_html"))
import relayout as RLY  # noqa: E402
rm = json.load(open(sys.argv[3]))
xl = sys.argv[1]
wb = openpyxl.load_workbook(xl, data_only=True)
IR = rm["INROW"]; PB = rm["PB"]; SUMR = rm["SUMR"]; T = rm["T"]; XR = rm["XR"]; XH = rm["XH"]
S = lambda s, k: SUMR[f"{s}|{k}"]
# 情境驅動對照
INS = wb["In"]
scn = next(r for r in range(1, INS.max_row + 1) if str(INS.cell(r, 1).value or "").startswith("情境驅動對照"))
scn_rows = []   # 第三輪修正：只取「情境驅動對照」節本身（到第一個空白鍵為止）；原本取到工作表末尾，混入只進敏感度的 Wattonomics 錨定係數列
for r in range(scn + 2, INS.max_row + 1):
    if not INS.cell(r, 1).value: break
    scn_rows.append(r)
# 第三輪修正：情境驅動表數值格式（年度不加千分位、百分比／pt 以 % 顯示、其他依大小取小數位）
def _scen_fmt(ws, r):
    u = str(ws.cell(r, 3).value or "").strip()
    vs = [ws.cell(r, c).value for c in (4, 5, 6)]
    vs = [v for v in vs if isinstance(v, (int, float)) and not isinstance(v, bool)]
    if not vs:
        return ""
    if all(float(v).is_integer() and 1990 <= v <= 2100 for v in vs) and ("年" in u or u in ("", "—")):
        return "|yr"
    if u in ("%", "pt", "百分點") or u.startswith("%"):
        return "|p2" if any(0 < abs(v) < 0.001 for v in vs) or any(abs(round(v * 1000) - v * 1000) > 1e-6 for v in vs) else "|p1"
    if all(float(v).is_integer() for v in vs):
        return "|n0"
    return "|n1" if max(abs(v) for v in vs) >= 100 else "|n2"
def drv(key, fmt, dlabel, dref, dfmt, why, unit=True):
    r = IR[key]
    base = [f"@In!E{r}|{fmt}"] + ([" ", f"@In!C{r}"] if unit else [])
    return [f"@In!B{r}", base, f"@In!G{r}|tag", [dlabel, f"{dref}|{dfmt}"], why]
A = PB["B"]
drivers = [
    drv("q2_lg", "p0", "授權及其他（FY2027 錨點）：", "@Q!H4", "n3", "第二季實績推得；錨點＝第一、二季實績＋第三季公司指引＋第四季假設", unit=False),
    drv("ph_s", "p0", "智慧型手機權利金（FY2026）：", f"@PL!D{A['ph']}", "n3", "公司揭露的唯一市場比例；其餘市場為假設", unit=False),
    drv("ph_u26", "n2", "手機每支權利金（FY2026）：", f"@PL!D{A['phrt']}", "n3", "出貨持平到下滑，權利金成長靠費率"),
    drv("cpu_gpu", "n2", "輝達 Grace／Vera 處理器（FY2028，情境 B）：", f"@PL!F{A['nvcpu']}", "n2", "加速器顆數讀 Siliconomics，同字母情境"),
    drv("g_hc", "p0", "資料中心處理器顆數指數（FY2028，情境 B）：", f"@PL!F{A['dcx']}", "n2", "雲端自研處理器，乘同字母顆數係數；待併入 Siliconomics", unit=False),
    drv("dcr", "p0", "資料中心權利金（FY2028，情境 B）：", f"@PL!F{A['dc']}", "n3", "核心數增加使每顆權利金上升", unit=False),
    drv("g_lic", "p0", "授權（不含晶片，FY2028）：", f"@PL!F{A['lic']}", "n3", "年度合約價值成長；含軟銀協議", unit=False),
    drv("chip_gm1", "p0", "毛利率（FY2028，情境 B）：", f"@PL!F{A['gm']}", "p1", "首代晶片毛利率取公司揭露區間，遠低於矽智財；晶片占比上升使毛利率下降", unit=False),
    drv("g_opx", "p0", "營業利益率（FY2028，情境 B）：", f"@PL!F{A['oim']}", "p1", "晶片與運算子系統研發持續投入", unit=False),
]
fm = lambda c: c
html = {
 "company": "Arm", "ticker": "ARM", "thesis": "@README!A3",
 "question": "現價是否已把資料中心權利金倍增、運算子系統費率上升與自研晶片全部算進去；高通判決與手機出貨下修會讓權利金少多少",
 "scenarios": [{"key": "A", "name": "情境 A"}, {"key": "B", "name": "情境 B"}, {"key": "C", "name": "情境 C"}],
 "sections": [
  {"id": "thesis", "title": "命題與核心問題", "intro": "@README!A4", "blocks": [
     {"type": "kv", "items": [["情境 A", f"@PL!A{PB['A']['nvcy'] - 3}"], ["情境 B", f"@PL!A{PB['B']['nvcy'] - 3}"], ["情境 C", f"@PL!A{PB['C']['nvcy'] - 3}"], ["會計年度與口徑", "@README!A5"]]}]},
  {"id": "drivers", "title": "驅動因子 → 推導量", "intro": "標記沿用 Excel：Verified、Interested-party、Analogy、Assumed、Derived。輸入取自 In 頁，推導量取自 PL 頁情境 B 與 Units 頁。", "blocks": [
     {"type": "table", "caption": "單一參數（In 頁基準值；敏感度低／高只用於敏感度頁、不進目標價）", "head": ["輸入", "基準值", "標記", "推導量", "因果方向"], "rows": drivers},
     {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；取自 In 頁最後一節「情境驅動對照」）",
      "head": [f"@In!B{scn + 1}", f"@In!D{scn + 1}", f"@In!E{scn + 1}", f"@In!F{scn + 1}", f"@In!C{scn + 1}"],
      "rows": [[f"@In!B{r}", f"@In!D{r}{_scen_fmt(INS, r)}", f"@In!E{r}{_scen_fmt(INS, r)}", f"@In!F{r}{_scen_fmt(INS, r)}", f"@In!C{r}"] for r in scn_rows]},
     {"type": "table", "caption": "權利金分市場與資料中心處理器（情境 B）", "head": ["@Units!A3", "@Units!B3", "@Units!C3", "@Units!D3", "@Units!E3", "@Units!F3"],
      "rowgen": {"sheet": "Units", "rows": f"4-{rm['UR']['chip']}", "cols": ["A", "B", "C", "D", "E", "F"], "fmts": ["", "", "n3", "n3", "n3", "n3"],
                 "row_fmt": {str(rm["UR"][k]): {c: f for c in "CDEF"} for k, f in (("nvcy", "n2"), ("nvcpu", "n2"), ("dcx", "n2"), ("dcrx", "n2"), ("dcsh", "p1"))}}},
     {"type": "note", "parts": [f"@Units!A{rm['UNOTE']}"]}]},
  {"id": "pl", "title": "三情境損益", "intro": "FY2025、FY2026 為實績，FY2027 為錨點（第一、二季實績＋第三季公司指引＋第四季假設），FY2028 起為模型。會計年度三月底結束；上游日曆年依月份加權。", "blocks": [
     {"type": "pl", "years": "@PL!C3:G3", "metrics": [
        {"name": ["營收（", f"@PL!B{S('A', 'rev')}", "）"], "fmt": "n3", "chart": True, "rows": {s: f"@PL!C{S(s, 'rev')}:G{S(s, 'rev')}" for s in "ABC"}},
        {"name": ["權利金（", f"@PL!B{S('A', 'ro')}", "）"], "fmt": "n3", "chart": True, "rows": {s: f"@PL!C{S(s, 'ro')}:G{S(s, 'ro')}" for s in "ABC"}},
        {"name": ["其中資料中心（", f"@PL!B{S('A', 'dc')}", "）"], "fmt": "n3", "rows": {s: f"@PL!C{S(s, 'dc')}:G{S(s, 'dc')}" for s in "ABC"}},
        {"name": ["授權及其他（", f"@PL!B{S('A', 'lo')}", "）"], "fmt": "n3", "rows": {s: f"@PL!C{S(s, 'lo')}:G{S(s, 'lo')}" for s in "ABC"}},
        {"name": ["其中晶片（", f"@PL!B{S('A', 'chip')}", "）"], "fmt": "n3", "rows": {s: f"@PL!C{S(s, 'chip')}:G{S(s, 'chip')}" for s in "ABC"}},
        {"name": "毛利率", "fmt": "p1", "chart": True, "ymin": "auto", "rows": {s: f"@PL!C{S(s, 'gm')}:G{S(s, 'gm')}" for s in "ABC"}},
        {"name": "營業利益率", "fmt": "p1", "rows": {s: f"@PL!C{S(s, 'oim')}:G{S(s, 'oim')}" for s in "ABC"}},
        {"name": ["EPS（", f"@PL!B{S('A', 'eps')}", "）"], "fmt": "n2", "chart": True, "rows": {s: f"@PL!C{S(s, 'eps')}:G{S(s, 'eps')}" for s in "ABC"}}]},
     {"type": "kv", "items": [[f"@PL!A{rm['CHK'] + i}", f"@PL!D{rm['CHK'] + i}|n3"] for i in range(4)]},
     {"type": "h3", "text": "延伸到 FY2032（FY2030 起為收斂法）"},
     {"type": "note", "parts": [f"@PL!A{rm['XT']}"]},
     {"type": "pl", "years": f"@PL!C{XH}:J{XH}", "metrics": [
        {"name": ["營收（", f"@PL!B{XR['A']['rev']}", "）"], "fmt": "n3", "chart": True, "rows": {s: f"@PL!C{XR[s]['rev']}:J{XR[s]['rev']}" for s in "ABC"}},
        {"name": "淨利率", "fmt": "p1", "rows": {s: f"@PL!C{XR[s]['nim']}:J{XR[s]['nim']}" for s in "ABC"}},
        {"name": ["EPS（", f"@PL!B{XR['A']['eps']}", "）"], "fmt": "n2", "chart": True, "rows": {s: f"@PL!C{XR[s]['eps']}:J{XR[s]['eps']}" for s in "ABC"}}]}]},
 ]}
V = wb["Val"]
DW = wb["DCF"]
FB = next(r for r in range(next(r for r in range(1, DW.max_row + 1) if str(DW.cell(r, 1).value or "").startswith("情境 B（基準）")), DW.max_row + 1) if DW.cell(r, 1).value == "營收")
_dfm = lambda lab: ("n3" if lab.startswith("評價用稀釋股數") else "p1" if "終值現值占 EV" in lab else "x1" if "隱含終值" in lab else "n2")
assert str(V.cell(T["dcftp"], 1).value).startswith("DCF 目標價（情境加權）")
def vrow(k, fmt, three=False):
    r = T[k]
    return [f"@Val!A{r}", f"@Val!B{r}|{fmt}"] + ([f"@Val!C{r}|{fmt}", f"@Val!D{r}|{fmt}"] if three else ["", ""])
rmrow = {str(V.cell(r, 1).value): r for r in range(1, V.max_row + 1) if V.cell(r, 1).value}
rmk = [rmrow[x] for x in ("本公司 10 年中位數（納入值）", "同業 10 年中位數的中位數", "原始比值（本公司 ÷ 同業）", "相對倍數調整係數", "Comparable 目標價（相對倍數調整後）")]
OB = rm["OBSROW"]
val_blocks = [
  {"type": "table", "caption": "目標價（DCF 與 Comparable 加權）", "head": [f"@Val!A{T['h3']}", f"@Val!B{T['h3']}", "", ""],
   "rows": [vrow("tpw", "n2"), vrow("comp2", "n2"), vrow("cons", "n2"), vrow("upw", "sp1"), vrow("upc", "sp1"), [f"@Val!A{T['grade']}", f"@Val!B{T['grade']}", "", ""],
            vrow("geps", "sp1"), vrow("gmul", "sp1"), [f"@Val!A{T['band']}", f"@Val!B{T['band']}", "", ""], vrow("pr", "p0", True), [f"@Val!A{T['m']}", f"@Val!B{T['m']}", "", ""]]},
  {"type": "note", "parts": [f"@Val!A{T['rec']}", "：", f"@Val!B{T['rec']}", "（依上檔空間固定規則）"]},
  {"type": "h3", "text": "相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司相對同業的長期倍數比）"},
  {"type": "kv", "items": [[f"@Val!A{r}", f"@Val!B{r}|n2"] for i, r in enumerate(rmk)]},
  {"type": "h3", "text": "重點觀察"},
  {"type": "list", "items": [f"@README!A{OB + 1}", f"@README!A{OB + 2}", f"@README!A{OB + 3}"]},
  {"type": "table", "caption": "DCF 折現率（WACC、股權成本、稅後債務成本、RONIC、永續成長率；直接顯示）", "head": ["項目", "數值", "說明"],
   "rows": [[f"@Val!A{r}", f"@Val!B{r}|p2", f"@Val!E{r}"] for r in range(T["rk"], T["rk"] + 6)]},
  {"type": "table", "caption": "DCF 彙總（標準版：無槓桿自由現金流 FCFF＋WACC；三情境；十億美元、每股美元）", "head": ["項目", "A（樂觀）", "B（基準）", "C（保守）", "說明"],
   "rows": [[f"@Val!A{r}", f"@Val!B{r}|{fm}", f"@Val!C{r}|{fm}" if r != T["dcftp"] else "", f"@Val!D{r}|{fm}" if r != T["dcftp"] else "", f"@Val!E{r}"]
            for r, fm in ((r, _dfm(str(V.cell(r, 1).value).strip())) for r in range(T["rk"] + 6, T["dcftp"] + 1))]},   # 第二輪修正：依 DCF 目標價列定位（模組 v2 stub 期多兩列）
  {"type": "table", "collapse": True, "summary": "展開：情境 B（基準）FCFF 逐年明細（DCF 頁；十億美元；FY2027 只計評價日後一季）", "head": ["@DCF!A5", "@DCF!C5", "@DCF!D5", "@DCF!E5", "@DCF!F5", "@DCF!G5", "@DCF!H5"],
   "rowgen": {"sheet": "DCF", "rows": f"{FB}-{FB + 10}", "cols": ["A", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n3", "n3", "n3", "n3", "n3", "n3"],
              "row_fmt": {str(FB + 3): {c: "p1" for c in "CDEFGH"}, str(FB + 5): {c: "p1" for c in "CDEFGH"}, str(FB + 6): {c: "p1" for c in "CDEFGH"}}}},
  {"type": "bars", "title": "DCF 情境價值（推到十二個月後）與現價", "items": [{"label": f"情境 {s}", "value": f"@Val!{c}{T['dcf12']}", "series": i} for i, (s, c) in enumerate(zip("ABC", "BCD"))],
   "ref": {"label": "現價", "value": f"@Val!B{T['px']}"}, "fmt": "n2"},
  {"type": "table", "caption": f"@Val!A{T['ph'] - 1}", "head": [f"@Val!A{T['ph']}", f"@Val!B{T['ph']}", f"@Val!C{T['ph']}", f"@Val!D{T['ph']}", f"@Val!F{T['ph']}"],
   "rows": [[f"@Val!A{r}", f"@Val!B{r}", f"@Val!C{r}|x1", f"@Val!D{r}", f"@Val!F{r}"] for r in range(T["p1"], T["p2"] + 1)]
           + [[f"@Val!A{T['pmed']}", f"@Val!B{T['pmed']}|x1", "", "", ""], [f"@Val!A{T['feps']}", f"@Val!B{T['feps']}|n2", "", "", ""], [f"@Val!A{T['comptp']}", f"@Val!B{T['comptp']}|n2", "", "", ""]]
           + [[f"@Val!A{T[k]}", f"@Val!B{T[k]}|{f}", "", "", ""] for k, f in (("pm_ip", "x1"), ("pm_chip", "x1"), ("ct_ip", "n2"), ("ct_chip", "n2"))]},
  {"type": "note", "parts": [f"@Val!E{T['pmed']}"]},
  {"type": "table", "caption": f"@Val!A{T['pe10h']}", "collapse": True, "summary": "展開：十年年底市盈率（分位檢查用）",
   "head": [f"@Val!{c}{T['pe10h']}" for c in "ABCDEFGHIJK"], "rows": [[f"@Val!A{T['pe10']}"] + [f"@Val!{c}{T['pe10']}|x1" for c in "BCDEFGHIJK"]]},
  {"type": "table", "caption": "現價與市場隱含倍數", "head": [f"@Val!A{T['s1'] + 1}", f"@Val!B{T['s1'] + 1}", f"@Val!C{T['s1'] + 1}", f"@Val!D{T['s1'] + 1}"],
   "rows": [vrow("px", "n2")] + [vrow(f"pe{y}", "x1", True) for y in range(2026, 2030)] + [vrow("ttm", "n2"), vrow("pettm", "x1"), vrow("pecons", "x1"), vrow("mcap", "n0"), vrow("float", "n0")]},
  {"type": "note", "parts": ["@Val!A2"]}]
html["sections"].append({"id": "valuation", "title": "估值", "intro": "@Val!A1", "blocks": val_blocks})
RL = rm["RECON_LAST"]
RC = wb["Recon"]
rf = {}
for r in range(4, RL + 1):
    lab = str(RC.cell(r, 1).value or "")
    if "÷" in lab: rf[str(r)] = {c: "p1" for c in "BCEG"}
    if "EPS" in lab or "CPU" in lab: rf[str(r)] = {c: "n2" for c in "BCEG"}
    if "營收" in lab and "EPS" not in lab: rf[str(r)] = {c: "n3" for c in "BCEG"}
    if "每支" in lab: rf[str(r)] = {c: "n3" for c in "BCEG"}
    if lab.startswith(("支持現價所需", "評等翻轉")): rf[str(r)] = {"B": "n2" if "EPS" in lab else "x1", "C": "n2" if "EPS" in lab else "x1", "D": "sp1"}
html["sections"].append({"id": "recon", "title": "對帳：公司指引、共識與權利金結構", "intro": "@Recon!A1", "blocks": [
  {"type": "table", "head": [f"@Recon!{c}3" for c in "ABCDEFGHIJ"], "rowgen": {"sheet": "Recon", "rows": f"4-{RL}", "cols": list("ABCDEFGHIJ"),
   "fmts": ["", "n3", "n3", "sp1", "n3", "sp1", "n3", "sp1", "", ""], "row_fmt": rf}}]})
SL = rm["SENS_LAST"]
html["sections"].append({"id": "sens", "title": "敏感度", "intro": "@Sens!A1", "blocks": [
  {"type": "tornado", "title": "加權目標價相對基準的變動（一次改一個輸入的敏感度低／高值）", "sheet": "Sens", "rows": f"5-{SL}", "label_col": "A", "value_col": "E", "fmt": "sp1"},
  {"type": "note", "parts": ["@Sens!A2"]},
  {"type": "table", "collapse": True, "summary": "展開：全部案例的加權目標價與情境 B 的 FY2028 EPS", "head": [f"@Sens!{c}3" for c in "ABCDEFGH"],
   "rowgen": {"sheet": "Sens", "rows": f"4-{SL}", "cols": list("ABCDEFGH"), "fmts": ["", "", "", "n2", "sp1", "", "n2", "sp1"]}},
  {"type": "note", "parts": [f"@Sens!A{rm['SENS_NOTE']}"]},
  {"type": "table", "caption": f"@Sens!A{rm['TVS0']}", "head": [f"@Sens!{c}{rm['TVS0'] + 1}" for c in "ABCDEFG"],
   "rowgen": {"sheet": "Sens", "rows": f"{rm['TVS0'] + 2}-{rm['TVS1']}", "cols": list("ABCDEFG"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2"]}}]})
V34 = rm["V34"]["G"]
RDW = wb["README"]
html["sections"].append({"id": "status", "title": "資料狀態", "blocks": [
  {"type": "note", "parts": [f"@README!A{V34['state']}"]},
  {"type": "h3", "text": "主要失效條件與估值框架並列"},
  {"type": "note", "parts": [f"@README!A{V34['fail_h']}"]},
  {"type": "list", "items": [f"@README!A{r}" for r in V34["fail"]] + [f"@README!A{V34['frame']}"]},
  {"type": "h3", "text": "資料基準日與逾期標示"},
  {"type": "note", "parts": [f"@README!A{V34['asof_h']}"]},
  {"type": "table", "head": [f"@README!{c}{V34['asof_t']}" for c in "ABC"], "rows": [[f"@README!A{r}", f"@README!B{r}", f"@README!C{r}"] for r in V34["asof_rows"]]},
  {"type": "note", "parts": [f"@README!A{V34['asof_n']}"]},
  {"type": "h3", "text": "風險登錄（是否建模、理由、觀察指標）"},
  {"type": "list", "items": [f"@README!A{r}" for r in V34["risk"]]},
  {"type": "h3", "text": "主要假設驅動與可證偽的公開指標"},
  {"type": "list", "items": [f"@README!A{r}" for r in V34["ind"]]},
  {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋（取代建置 commit）", "$sha"],
                           ["行業變數來源", "@In!A2"]]},
  {"type": "h3", "text": "上游成品（市場層與其他公司模型）"},
  {"type": "upstream", "sheets": ["In"]},
  {"type": "h3", "text": "待複核項目"},
  {"type": "list", "items": ["財報數字取自公司股東信（美國證管會六之 K 附件）；分市場權利金除手機外為假設，待公司揭露後複核。", "@Q!A2", "標為「未找到」的項目（見下方清單）待找到來源後補上。"]},
  {"type": "h3", "text": "已知限制"},
  {"type": "note", "parts": ["@README!A8"]},
  {"type": "h3", "text": "「未找到」清單"},
  {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的輸入與來源", "sheets": {"In": ["B", "H", "I"], "Q": ["A", "I"], "Val": ["A", "E"], "Sources": ["A", "B"]}, "empty": "沒有標為未找到的項目。"}]})
json.dump(html, open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "html_src.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
# ---------- v0.3：個股 HTML 版面規格 v1.0（relayout）----------
for b in html["sections"][0]["blocks"]: pass
for sec_ in html["sections"]:
    if sec_["id"] == "valuation":
        sec_["blocks"] = [b for b in sec_["blocks"] if not (b["type"] == "h3" and b.get("text") == "重點觀察") and not (b["type"] == "list" and str(b.get("items", [""])[0]).startswith("@README"))]
    if sec_["id"] == "drivers":
        for b in sec_["blocks"]:
            if str(b.get("caption", "")).startswith("單一參數"): b["caption"] = "驅動 → 推導對照（單一參數：輸入、基準值、推導量與因果方向）"
IK = {str(INS.cell(r, 1).value) for r in range(4, INS.max_row + 1) if INS.cell(r, 1).value}
SPEC = {
 "in_skip": ["F", next(re.split(r"[.\s]", str(INS.cell(r, 1).value))[0] for r in range(4, INS.max_row + 1) if str(INS.cell(r, 1).value or "").startswith("情境驅動對照"))],
 "in_groups": [
  ["權利金（智慧型手機、資料中心 Neoverse、車用、物聯網：出貨 × 每顆權利金率）", [
    ["FY2026 基期比例與 FY2027 年增", ["ph_s", "dc_s", "au_s", "ph_u26", "g_ph27", "g_dc27", "g_au27"]],
    ["智慧型手機（出貨、費率、Qualcomm）", ["g_phu28", "g_phu29", "phr_A", "phr_B", "phr_C", "qc_adj"]],
    ["事件型尾端風險（只進敏感度；Qualcomm 停付、Arm China、SoftBank 授權）", ["qc_sh", "qc_stop", "cn_s", "cn_adj", "sb_lic"]],
    ["資料中心（NVIDIA Grace／Vera 顆數讀 Siliconomics、雲端自研 CPU、每顆權利金率）", ["cpu_gpu"] + [k for k in sorted(IK) if k.startswith(("nvbw_", "nvrb_", "nvru_", "uc28_", "uc29_"))] + ["nv_cy25", "w_nv", "g_hc", "dcr"]],
    ["車用、物聯網與其他", ["g_au", "g_iot"]],
    ["Wattonomics 錨定係數（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  ["授權及其他（含 SoftBank 協議）與 Arm AGI CPU 晶片", [
    ["授權成長與晶片營收、毛利率", ["g_lic", "chip28_A", "chip28_B", "chip28_C", "chip29_A", "chip29_B", "chip29_C", "chip_gm1", "chip_gm", "chip_27"]]]],
  ["全公司（FY2027 錨點、費用與稅、股價、目標價規則、折現率與自由現金流）", [
    ["FY2027 錨點（第一、二季實績、第三季指引、第四季假設）", ["q2_lg", "q2_rg", "q3_rev", "q3_rg", "h2_lg", "h2_rg", "q2_opx", "q3_opx", "q4_opx", "cogs_r", "q_oth", "q_tx", "q_sh"]],
    ["營業費用、業外、稅與股數（FY2028 起）", ["g_opx", "oth_f", "tx_f", "sh_g"]],
    ["股價、股權與共識", ["px", "sb_own", "cons_tp", "cons_feps"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_wA", "tp_wB", "tp_wC", "tp_w_dcf"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta_u", "beta_de", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（股票薪酬、折舊、資本支出、營運資金）", ["sbc_rev", "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["收斂法（FY2030–FY2032）", ["tp_g", "fade_gA", "fade_gB", "fade_gC", "nim10_f"]]]]],
 "blocks_to": {},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"]},
 "n_recon": 5,
}
out = RLY.relayout(html, xl, SPEC)
ht = out["hero"]
ht["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{T['comp2']}|n2", ["上檔空間 ", f"@Val!B{T['upc']}|sp1"]]
for _t in ht["tiles"]:   # 第三輪修正：橫幅金額單位（美元）
    if isinstance(_t[0], str): _t[0] = [_t[0], "（美元）"]
# ================= v0.3.5（2026-11-15）：個股 HTML 版面規格 v1.3＋三層協調（樣板 avgo v0.5.6／v0.5.7、amd v0.4.7、crdo v0.3.6） =================
cfg, X, OUT = out, xl, sys.argv[2]
HERE = os.path.dirname(os.path.abspath(__file__))


def find(sh, p):
    W = wb[sh]
    return next(r for r in range(1, W.max_row + 1) if str(W.cell(r, 1).value or "").strip().startswith(p))


# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含 C0 實績）→ 資料狀態。
# 數字與文字取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末、Coord 頁；arm/v035_post.py 產生）。列號依標籤找；設定文字除年份與日期外不含數字。
PCODE = re.compile(r"[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+")


def _v13():
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RV, CW, SW = wb["README"], wb["PL"], wb["RevDCF"], wb["Coord"], wb["Sens"]
    r13 = find("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "coord", "ver"], range(r13 + 1, r13 + 11)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.3.6"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    old_note = list(cfg["hero"]["note"])
    assert old_note[0] == "投資建議依加權目標價判定：", old_note
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"]
    # ---- 1. 命題：置頂對照圖（FY2027 營收、EPS 對共識）、投資建議條件式、重點觀察 5 點 ----
    th = S["thesis"]["blocks"]
    rr = lambda p: find("Recon", p)
    r27, rep27 = rr("FY2027 營收 vs 共識"), rr("FY2027 non-GAAP EPS vs 共識")
    grp = lambda nm, r, fmt, ref: {"name": [nm], "fmt": fmt, "items": [{"name": [ref], "v": f"@Recon!B{r}", "ref": True},
                                                                      {"name": ["模型 情境 A"], "v": f"@Recon!C{r}"}, {"name": ["模型 情境 B"], "v": f"@Recon!E{r}"}, {"name": ["模型 情境 C"], "v": f"@Recon!G{r}"}]}
    cmp_ = {"type": "compare", "title": "回答命題：模型的 FY2027 營收與 EPS 對市場共識（三情境在錨點年相同；差異在倍數）",
            "groups": [grp("FY2027 營收（US$B）", r27, "n2", "市場共識"), grp("FY2027 non-GAAP EPS（US$）", rep27, "n2", "市場共識")],
            "caption": "灰色＝市場共識；青綠＝本模型三情境（FY2027 為錨點：上半年實績＋第三季指引＋第四季假設，三情境相同）。盈餘與共識相近，目標價差距來自倍數。對照只作事後檢查，校準時未使用。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th.insert(0, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
    th.insert(0, cmp_)
    # ---- 2. 估值：四個目標價置頂（G 位置）；原橫幅的評等規則與相對倍數說明放圖下 ----
    va = S["valuation"]["blocks"]
    g0 = find("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = g0 + 7
    four = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
            "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                      ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
            "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
    tbl4 = {"type": "table", "caption": "四個目標價並列（美元；相對現價）", "head": ["目標價", "美元", "相對現價", "說明"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_px}", f"@Val!B{r_px}|n2", "", f"@Val!D{r_px}"]]}
    va[0:0] = [four, tbl4, {"type": "note", "parts": old_note}]
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").startswith(p))
    ra = rv("A. DCF"); rb_ = rv("(a)"); rbb = rv("(b)"); rc = rv("(c)"); rcmp = rv("B. Comparable"); rsum = rv("C. 摘要"); yh = rv("項目（情境 B）")
    m1, m0 = rv("現價隱含的 EBIT 率", rb_), rv("模型：基準的等效 EBIT 率", rb_)
    s1, rv1, rv0 = rv("現價隱含的營收倍數", rbb), rv("現價隱含的 FY2028 營收", rbb), rv("模型：FY2028 營收", rbb)
    gg1, gg0 = rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
    w1, w0, wok = rv("現價隱含的 WACC", rc + 1), rv("模型 WACC", rc), rv("判讀（隱含 WACC）", rc)
    a_rows = [rv(p, ra) for p in ("現價（美元）", "支持現價所需的評價日每股價值", "淨現金滾到評價日", "評價用稀釋股數", "支持現價所需的企業價值", "模型企業價值", "EV 需放大", "模型：情境 B 評價日每股價值")]
    afmt = ["n2", "n2", "n1", "n3", "n1", "n1", "n2", "n2"]
    rev_ = {
        "id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
        "blocks": [
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
            {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
             "rows": [["預測期 EBIT 率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                      ["營收規模（FY2028 起各年同乘的倍數）", f"@RevDCF!B{s1}|n2", "—"],
                      ["FY2028 營收（十億美元）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                      ["FY2028 營收年增率", f"@RevDCF!B{gg1}|p1", f"@RevDCF!B{gg0}|p1"],
                      ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"],
                      ["隱含 WACC 判讀", f"@RevDCF!B{wok}", ""]]},
            {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(a_rows, afmt)]},
            {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
             "rows": [[f"@RevDCF!A{rcmp + 2}", "—", f"@RevDCF!C{rcmp + 2}|n2", "—"]]
                     + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!C{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!D{r}|sp1"] for k, r in enumerate(range(rcmp + 3, rcmp + 6))]
                     + [[f"@RevDCF!A{rcmp + 6}", "—", f"@RevDCF!C{rcmp + 6}|n1", "—"]]},
            {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
             "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）",
             "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|" + ("p1" if r in (yh + 3, yh + 4) else "n3" if r in (yh + 5, yh + 6, yh + 7, yh + 8) else "n1") for c in "BCDEFG"]
                      for r in range(yh + 1, rb_ - 1)]},
        ]}
    # ---- 5. 損益（標準損益表）----
    p0 = find("PL", "標準損益表（v0.3.5")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
    COLS_ = "CDEFGHIJK"
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（研發＋銷管）", "營業利益", "營業利益率",
            "其他收入（淨額）＝利息收入＋其他業外", "稅前利益", "所得稅", "稅後淨利", "淨利率", "稀釋股數", "EPS（non-GAAP 稀釋）"]
    fmt = lambda u: {"%": "p1", "US$": "n2", "十億股": "n3"}.get(u, "n1")
    rowx = lambda r, cols=COLS_: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in cols]
    r_ratio = plr("比率"); r_gaap = next(r for r in range(ph, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("GAAP 對照"))
    r_chk = plr("核對（應為 0）"); r_c2b = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("差額（FY2030E"))
    is_tbl = {"type": "table", "caption": "標準損益表（Non-GAAP；十億美元，EPS 為美元，稀釋）：FY2024A–FY2026A 實績、FY2027E–FY2032E 情境 B（基準）",
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in COLS_],
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_gaap)]
                      + [{"sub": "核對：FY2030E–FY2032E 兩種推法的營業利益差額（毛利率與營業費用率沿用 FY2029E）"}] + [rowx(r) for r in (r_c2b, r_c2b + 1)]}
    gaap_tbl = {"type": "table", "caption": f"@PL!A{r_gaap}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in "CDE"], "collapse": True,
                "summary": "展開：GAAP 對照（只列實績）與核對列",
                "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" if PLW[f"{c}{r}"].value is not None else "—" for c in "CDE"] for r in range(r_gaap + 1, r_chk)]
                        + [{"sub": f"@PL!A{r_chk}"}] + [[f"@PL!A{r}"] + [f"@PL!{c}{r}|n3" if PLW[f"{c}{r}"].value is not None else "—" for c in "CDE"] for r in range(r_chk + 1, r_chk + 9) if r not in (r_c2b, r_c2b + 1)]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率、淨利率：FY2024A–FY2032E（情境 B）", "x": f"@PL!C{ph}:K{ph}",
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:K{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:K{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:K{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:K{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，Non-GAAP 模型口徑）。FY2027E 為錨點；FY2028E 起毛利率下降來自 Arm 自研晶片（毛利率遠低於矽智財）；FY2030E–FY2032E 為收斂期：營收與淨利率由收斂法推得，毛利率與營業費用率沿用 FY2029E。"}
    xh = 3
    assert PLW.cell(xh, 1).value == "情境／項目"
    xs_ = {k: next(r for r in range(xh, PLW.max_row + 1) if PLW.cell(r, 1).value == f"情境 {k}") for k in "ABC"}
    xr = lambda base, lab_: next(r for r in range(base, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
    sc_chart = {"type": "theme", "title": "三情境並列：營收（長條，十億美元）與 EPS（折線，美元，右軸）", "x": f"@PL!C{xh}:G{xh}",
                "fmt": "n1", "fmt_r": "n2", "stack": False, "pad": 1.1,
                "bars": [{"name": [f"@PL!A{xs_[k]}", " 營收"], "y": f"@PL!C{xr(xs_[k], '營收')}:G{xr(xs_[k], '營收')}", "color": c} for k, c in (("A", "#b7e3e4"), ("B", "#10a0a7"), ("C", "#8a8c8c"))],
                "lines": [{"name": [f"@PL!A{xs_[k]}", " EPS"], "y": f"@PL!C{xr(xs_[k], 'EPS（non-GAAP）')}:G{xr(xs_[k], 'EPS（non-GAAP）')}", "axis": "r", "color": c, **kw}
                          for k, c, kw in (("A", "#0b7a80", {"dash": True}), ("B", "#3b3838", {"labels": True}), ("C", "#b4472f", {"dash": True}))],
                "caption": "情境 A（樂觀）、B（基準）、C（保守）在 FY2027 相同（錨點）；差異自 FY2028 起出現（NVIDIA 與雲端自研處理器顆數係數、手機費率、自研晶片營收）；FY2030 起的收斂期見下方摺疊表。"}
    old_pl = S["pl"]["blocks"]
    keep = []
    for b in old_pl:
        if b["type"] == "pl":
            keep.append(dict(b, collapse=True, summary="展開：三情境摘要表與折線" if not keep else "展開：三情境延伸到 FY2032（收斂法）"))
        elif b["type"] == "note":
            keep.append(dict(b, collapse=True, summary="展開：收斂期（FY2030–FY2032）的說明"))
        elif b["type"] == "kv":
            keep.append(dict(b, collapse=True, summary="展開：損益口徑與期間說明"))
    pl_intro = S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, sc_chart, gaap_tbl]
                   + ([{"type": "note", "parts": [pl_intro], "collapse": True, "summary": "展開：三情境損益摘要的期間說明"}] if pl_intro else []) + keep)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    rbB = find("PL", "情境 B："); hdrB = rbB + 1
    assert PLW.cell(hdrB, 1).value == "項目"
    pr = lambda lab_, start=rbB: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip().startswith(lab_))
    r_ph, r_dc, r_au, r_iot = pr("權利金：智慧型手機"), pr("權利金：資料中心"), pr("權利金：車用"), pr("權利金：物聯網與其他")
    r_chip, r_lic = pr("授權及其他：Arm AGI CPU 晶片營收"), pr("授權及其他：授權（不含晶片")
    r_rev = next(r for r in range(pr("C. 授權及其他"), PLW.max_row + 1) if PLW.cell(r, 1).value == "營收")
    r_dsh, r_end = pr("資料中心 ÷ 權利金"), pr("智慧型手機 ÷ 權利金")

    def seg_row(r):
        a = str(PLW.cell(r, 1).value or "")
        vals = [PLW.cell(r, c).value for c in range(3, 8)]
        if all(v is None for v in vals):
            return {"sub": f"@PL!A{r}"}
        f = {"%": "p1", "x": "n2", "百萬顆": "n2", "十億支": "n2", "US$": "n2"}.get(PLW.cell(r, 2).value, "n3")
        return [f"@PL!A{r}", f"@PL!B{r}"] + [(f"@PL!{c}{r}|{f}" if PLW[f"{c}{r}"].value is not None else "—") for c in "CDEFG"]
    seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；十億美元）：權利金四個市場（出貨 × 每顆權利金率）＋授權及其他（授權＋自研晶片）＝營收",
               "head": [f"@PL!A{hdrB}", f"@PL!B{hdrB}"] + [f"@PL!{c}{hdrB}" for c in "CDEFG"],
               "rows": [seg_row(r) for r in range(hdrB + 1, r_end + 1) if not str(PLW.cell(r, 1).value or "").startswith(("檢查", "核對"))]}
    seg_chart = {"type": "theme", "title": "分部營收（情境 B，十億美元）與資料中心占權利金比重（右軸）", "x": f"@PL!C{hdrB}:G{hdrB}",
                 "fmt": "n2", "fmt_r": "p0", "stack": True, "pad": 1.2, "total": f"@PL!C{r_rev}:G{r_rev}",
                 "bars": [{"name": ["權利金：智慧型手機"], "y": f"@PL!C{r_ph}:G{r_ph}", "color": "#0b7a80"},
                          {"name": ["權利金：資料中心"], "y": f"@PL!C{r_dc}:G{r_dc}", "color": "#10a0a7"},
                          {"name": ["權利金：車用"], "y": f"@PL!C{r_au}:G{r_au}", "color": "#6fc7cb"},
                          {"name": ["權利金：物聯網與其他"], "y": f"@PL!C{r_iot}:G{r_iot}", "color": "#b7e3e4"},
                          {"name": ["授權（不含晶片）"], "y": f"@PL!C{r_lic}:G{r_lic}", "color": "#c29a2b"},
                          {"name": ["Arm 自研晶片"], "y": f"@PL!C{r_chip}:G{r_chip}", "color": "#6f6bae"}],
                 "lines": [{"name": ["資料中心 ÷ 權利金"], "y": f"@PL!C{r_dsh}:G{r_dsh}", "axis": "r", "color": "#3b3838", "labels": True}],
                 "caption": "堆疊＝權利金四個市場、授權與自研晶片；頂端數字＝營收；折線＝資料中心占權利金比重。FY2025A、FY2026A 的市場拆分只有智慧型手機為公司揭露，其餘為假設；FY2027 為錨點年。"}
    dr_ = S["drivers"]["blocks"]
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + dr_
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    S["drivers"].pop("intro", None)
    # ---- 9. 資料狀態 ----
    st = S["status"]["blocks"]
    st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]},
               {"type": "h3", "text": "三層協調：讀哪一層哪一版、由下往上對帳"}, {"type": "note", "parts": [f"@README!A{RM['coord']}"]},
               {"type": "h3", "text": "會計年度"}, {"type": "note", "parts": [f"@README!A{find('README', '會計年度：')}"]}]
    S["revdcf"] = rev_
    # ---- 7. 對帳章：C0（實績）、讀取層級核對（同源）、層級之間參考 ----
    fc = lambda p: next(r for r in range(1, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
    jr = fc("J. C0（實績）")
    jh = next(r for r in range(jr, CW.max_row + 1) if CW.cell(r, 1).value == "項目" and CW.cell(r, 2).value == "公司（口述）")
    je = jh
    while CW.cell(je + 1, 1).value: je += 1
    ih = next(r for r in range(jr, jh) if CW.cell(r, 1).value == "輸入")
    ie = ih
    while CW.cell(ie + 1, 1).value: ie += 1
    nf = lambda r, c: "n0" if CW.cell(r, 3).value == "月" else "p1" if CW.cell(r, 3).value == "%" else "n3"
    c0a_tbl = {"type": "table", "caption": "實績對帳（對帳判定）：資料中心權利金沒有同口徑揭露，判定無可比實績；下兩列只列參考（公司口述年增對層級隱含與本模型假設）",
               "head": ["項目", "公司（口述）", "層級／模型隱含", "差距", "判定", "說明"],
               "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|p0" if CW.cell(r, 2).value is not None else "", f"@Coord!C{r}|p1" if CW.cell(r, 3).value is not None else "",
                         f"@Coord!D{r}|sp1" if CW.cell(r, 4).value is not None else "", f"@Coord!E{r}", f"@Coord!I{r}"] for r in range(jh + 1, je + 1)]}
    c0i_tbl = {"type": "table", "caption": "實績對帳的輸入（出處與等級）", "collapse": True, "summary": "展開：實績與層級輸入（期間、出處、等級）",
               "head": [f"@Coord!{c}{ih}" for c in "ABCDEF"],
               "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|{nf(r, 2)}", f"@Coord!C{r}", f"@Coord!D{r}", f"@Coord!E{r}", f"@Coord!F{r}"] for r in range(ih + 1, ie + 1)]}
    gr0 = fc("G. 讀取層級核對"); gh = gr0 + 1; ge = gh
    while CW.cell(ge + 1, 1).value: ge += 1
    g_tbl = {"type": "table", "caption": "讀取層級核對（同源，不判定）：個股讀入的層級值與層級值必然相同，不進實績對帳判定、不計對帳中", "collapse": True, "summary": "展開：讀取層級核對（同源，不判定）",
             "head": ["項目", "本模型值", "層級對照值", "差距", "狀態", "說明"],
             "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|n2", f"@Coord!C{r}|n2", f"@Coord!D{r}|sp1", f"@Coord!E{r}", f"@Coord!I{r}"] for r in range(gh + 1, ge + 1)]}
    lr = fc("H. 層級之間"); lh = lr + 1; le = lh
    while CW.cell(le + 1, 1).value: le += 1
    lay_tbl = {"type": "table", "caption": "層級之間（參考；Siliconomics 對帳單處理，不計入本公司對帳中項目）：產業部署 GW，Siliconomics 基準對另一側讀數",
               "head": ["項目", "Siliconomics 基準", "另一側讀數", "差距", "狀態"],
               "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|n2", f"@Coord!C{r}|n2", f"@Coord!D{r}|sp1", f"@Coord!E{r}"] for r in range(lh + 1, le + 1)]}
    cnt = fc("對帳中的 C0 項目數")
    S["recon"]["blocks"][0:0] = [{"type": "note", "parts": [f"@Coord!A{jr + 1}"]}, c0a_tbl, {"type": "note", "parts": [f"@Coord!A{cnt}", "：", f"@Coord!B{cnt}|n0", "。", f"@Coord!A{cnt + 1}"]},
                                 c0i_tbl, g_tbl, lay_tbl]
    cfg["sections"] = [S[k] for k in ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]]


_v13()


# 敏感度章：兩張「加權目標價相對基準」龍捲風主圖（低案、高案；變動達兩位數百分比者直接顯示，其餘摺疊）；資料取 Sens 頁末表
def _sens_tor():
    SW = wb["Sens"]; t0 = find("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SW.cell(r2 + 1, 1).value and not str(SW.cell(r2 + 1, 1).value).startswith("快照狀態"): r2 += 1
    out = []
    for col, side in (("B", "低"), ("C", "高")):
        big = [r for r in range(r1, r2 + 1) if isinstance(SW[f"{col}{r}"].value, (int, float)) and abs(SW[f"{col}{r}"].value) >= 0.1]
        small = [r for r in range(r1, r2 + 1) if isinstance(SW[f"{col}{r}"].value, (int, float)) and 0.0005 <= abs(SW[f"{col}{r}"].value) < 0.1]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        out.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
        out.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                        summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例（幾乎為零者見全部案例表）"))
    out.append({"type": "table", "caption": "全部案例的加權目標價相對基準（低案、高案；營運輸入為整本重算快照）", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
                "head": [f"@Sens!{c}{t0 + 1}" for c in "ABCDE"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1" if SW[f"B{r}"].value is not None else "—", f"@Sens!C{r}|sp1" if SW[f"C{r}"].value is not None else "—",
                                                                         f"@Sens!D{r}|n2" if SW[f"D{r}"].value is not None else "—", f"@Sens!E{r}|n2" if SW[f"E{r}"].value is not None else "—"] for r in range(r1, r2 + 1)]})
    out.append({"type": "note", "parts": [f"@Sens!A{r2 + 1}", "：", f"@Sens!B{r2 + 1}", "。", f"@Sens!C{r2 + 1}"], "collapse": True, "summary": "展開：營運輸入案例的快照狀態"})
    return out


_SS = [x for x in cfg["sections"] if x["id"] == "sens"][0]
for _b in _SS["blocks"]:
    if _b.get("type") in ("table", "note") and not _b.get("collapse"):
        _b["collapse"] = True
        _b["summary"] = ("展開：" + ("目標價敏感度表（業務主因、折現參數）" if _b["type"] == "table" else "敏感度頁原文說明"))
    if _b.get("type") == "tornado":
        _b["summary"] = "展開：對照：EPS 相對基準的變動（情境 A，一次動一個輸入）"
        _b["collapse"] = True
_SS["blocks"] = _sens_tor() + _SS["blocks"]
_SS["intro"] = ("主圖為加權目標價相對基準的變動（低案、高案分開；變動達兩位數百分比者直接顯示，其餘摺疊）。營運輸入案例為整本重算（三情境同時套用）；"
                "估值框架案例（相對倍數係數、情境權重、DCF 權重）為封閉式；折現參數案例引用目標價參數表。下方摺疊為全部案例與明細。")


# 全頁可見文字不出現參數代號（版面規格 v1.3）：xlsx 原文不改，展示層含代號的格改用白話
def _plain_pass():
    PLAIN = {}
    V = wb["Val"]
    for r in range(1, V.max_row + 1):
        a = V.cell(r, 1).value
        if isinstance(a, str) and a.startswith("永續期 RONIC＝WACC＋"): PLAIN[f"Val!A{r}"] = "永續期新投資報酬率 RONIC（＝WACC＋超額報酬差）"
        if isinstance(a, str) and a.startswith("永續成長率 g＝MIN"): PLAIN[f"Val!A{r}"] = "永續成長率（長期成長上限與無風險利率取小者）"
        if isinstance(a, str) and a.startswith("DCF 情境價值推到 12 個月後"): PLAIN[f"Val!A{r}"] = "DCF 情境價值推到目標時點（以股權成本複利、扣除期間股利）"
    def p(spec):
        if isinstance(spec, list):
            return [p(x) for x in spec]
        if isinstance(spec, str) and spec.startswith("@"):
            ref = spec[1:].split("|")[0]
            if ref in PLAIN:
                return PLAIN[ref]
        return spec
    for sec in cfg["sections"]:
        if sec["id"] == "status":
            continue
        for b in sec["blocks"]:
            if b.get("type") == "table":
                b["rows"] = [({"sub": p(r["sub"])} if isinstance(r, dict) else p(r)) for r in b.get("rows", [])]
_plain_pass()

# 三層協調 C4（規則 v1.4；tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(cfg, X)


# 損益表預測年度欄標「對帳中」（Coord 頁 I 節「損益表年度標示」列，FY2027E–FY2032E＝標準損益表 F–K 欄）
def _coord_pl():
    CW = wb["Coord"]
    r_pl = next(r for r in range(1, CW.max_row + 1) if str(CW.cell(r, 1).value or "") == "損益表年度標示（FY2027E–FY2032E）")
    n = 0
    for b in [x for x in cfg["sections"] if x["id"] == "pl"][0]["blocks"]:
        if b.get("type") == "table" and isinstance(b.get("head"), list) and len(b["head"]) == 10 and str(b["head"][4]).startswith("@PL!F"):
            b["head"][4:10] = [f"@Coord!{c}{r_pl}" for c in "BCDEFG"]; n += 1
    assert n == 1, n
_coord_pl()


# 主圖置頂：估值章第一個區塊＝四個目標價與現價圖（章節說明、四個目標價表與說明放圖下，其後才是對帳中兩讀數區）；
# 敏感度章前兩個區塊＝加權目標價相對基準低案、高案主圖（說明與其他內容放下方）。只改區塊順序。
def _top_first():
    S = {x["id"]: x for x in cfg["sections"]}
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b.get("type") == "bars" and str(b.get("title", "")).startswith("四個目標價"))
    i_tbl, i_nt = i_bar + 1, i_bar + 2
    assert va[i_tbl].get("type") == "table" and str(va[i_tbl].get("caption", "")).startswith("四個目標價並列"), va[i_tbl]
    assert va[i_nt].get("type") == "note" and "投資建議依加權目標價判定" in str(va[i_nt].get("parts")), va[i_nt]
    top = [va[i_bar]]
    if S["valuation"].get("intro"):
        top.append({"type": "note", "parts": [S["valuation"].pop("intro")]})
    top += [va[i_tbl], va[i_nt]]
    S["valuation"]["blocks"] = top + [b for i, b in enumerate(va) if i not in (i_tbl, i_bar, i_nt)]
    sb = S["sens"]["blocks"]
    main = [i for i, b in enumerate(sb) if b.get("type") == "tornado" and not b.get("collapse") and str(b.get("title", "")).startswith("加權目標價相對基準")]
    assert len(main) == 2, main
    top = [sb[i] for i in main]
    if S["sens"].get("intro"):
        top.append({"type": "note", "parts": [S["sens"].pop("intro")]})
    S["sens"]["blocks"] = top + [b for i, b in enumerate(sb) if i not in main]
_top_first()
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

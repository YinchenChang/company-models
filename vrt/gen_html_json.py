# Vertiv html.json 產生器：讀 vrt/rowmap.json（build_vrt.py 輸出的列位置）與 dist 內的 xlsx（只用來判斷哪些敏感度案例變動達兩位數百分比），寫 vrt/html.json（勿手改）。
# 設定檔文字除年度與日期外不含阿拉伯數字；所有數字與標籤取自 xlsx 儲存格（@工作表!儲存格|格式）。
# 用法：python3 vrt/gen_html_json.py [成品 xlsx]（省略時用 vrt/dist/ 內唯一的 xlsx）
import glob, json, os, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
X = sys.argv[1] if len(sys.argv) > 1 else (lambda fs: fs[0] if len(fs) == 1 else sys.exit(f"dist 內應剛好一個 xlsx：{fs}"))(sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx"))))
wb = openpyxl.load_workbook(X, data_only=True)
M = json.load(open(os.path.join(HERE, "rowmap.json"), encoding="utf-8"))
T, RM, SD, RDM, PB, SRW, RD, SUMR = M["T"], M["RM"], M["SDCF"], M["RD_"], M["PB"], M["SRW"], M["README"], M["SUMR"]
YC = "CDEFGHIJ"
c = lambda sh, ref, f="": f"@{sh}!{ref}" + (f"|{f}" if f else "")
R_ = lambda k: c("README", f"A{RD[k]}")
VA = lambda r: c("Val", f"A{r}")
VB = lambda r, f="": c("Val", f"B{r}", f)

# ── 命題與觀察 ──
sec_thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "資料中心設施產品營收與其餘營收（情境 B，十億美元）與可及市場（右軸，十億美元）", "x": c("PL", f"C{PB['B']['mkt'] - 1}:G{PB['B']['mkt'] - 1}"), "fmt": "n1", "fmt_r": "n0", "stack": True, "pad": 1.25,
     "bars": [{"name": ["資料中心設施產品營收"], "y": c("PL", f"C{PB['B']['dcp']}:G{PB['B']['dcp']}"), "color": "#0b7a80"}, {"name": ["其餘營收（服務、備品、非資料中心）"], "y": c("PL", f"C{PB['B']['oth']}:G{PB['B']['oth']}"), "color": "#c29a2b"}],
     "lines": [{"name": ["可及市場（機櫃外配電、UPS、冷卻）"], "y": c("PL", f"C{PB['B']['mkt']}:G{PB['B']['mkt']}"), "axis": "r", "color": "#6fc7cb", "labels": True}],
     "caption": "兩個驅動：可及市場（Wattonomics 基準）與 Vertiv 份額；乘積得資料中心設施產品營收，加其餘營收得合併營收。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [R_("p1"), R_("p2"), R_("p3")]},
    {"type": "note", "parts": [R_("mkt")]},
    {"type": "list", "ordered": True, "collapse": True, "summary": "展開：命題回覆（依本模型計算的答案；想先看結論可展開，或讀完全文再看）", "items": [R_("r1"), R_("r2"), R_("r3"), R_("r4")]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [R_(f"o{i}") for i in range(1, 6)]},
    {"type": "h3", "text": "情境定義"},
    {"type": "kv", "items": [["情境 A", c("PL", f"A{PB['A']['mkt'] - 2}")], ["情境 B", c("PL", f"A{PB['B']['mkt'] - 2}")], ["情境 C", c("PL", f"A{PB['C']['mkt'] - 2}")], ["會計年度與口徑", R_("basis")]]}]}

# ── 估值 ──
rm = RM
LBL = {"pvx": "明確期 FCFF 現值", "pvtv": "終值現值", "ev": "企業價值", "ncr": "淨現金滾到評價日", "dcfv": "DCF 情境每股價值（評價日）", "v12": "DCF 情境價值推到十二個月後"}
pxr = T["pxr"]
val_blocks = [
    {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n0", "items": [
        {"label": "DCF 目標價", "value": VB(T["dcftp"]), "series": 1}, {"label": "Comparable（調整後，進加權）", "value": VB(T["compadj"]), "series": 2},
        {"label": "Comparable（未調整，不進加權）", "value": VB(T["comp2"]), "series": 2}, {"label": "加權目標價（橫幅主值）", "value": VB(T["tpw"]), "series": 0},
        {"label": "市場共識目標價", "value": VB(T["cons"]), "series": 1}], "ref": {"label": VA(pxr), "value": VB(pxr)}},
    {"type": "note", "parts": [R_("valn")]},
    {"type": "table", "caption": c("Val", f"A{T['h3'] - 1}"), "head": [VA(T["h3"]), c("Val", f"B{T['h3']}"), c("Val", f"C{T['h3']}"), c("Val", f"E{T['h3']}")],
     "rows": [[VA(r), VB(r, "n2"), c("Val", f"C{r}", "sp1") if r != pxr else "", c("Val", f"E{r}")] for r in range(T["c0"], T["c1"] + 1)]},
    {"type": "note", "parts": ["投資建議依加權目標價判定：", c("Val", f"E{T['rec']}"), "；進加權的 Comparable 為相對倍數調整後 ", VB(T["compadj"], "n2"), "（上檔空間 ", c("Val", f"C{T['compadj']}", "sp1"), "）"]},
    {"type": "table", "caption": "目標價、現價與投資建議（其餘項目）", "collapse": True, "summary": "展開：上檔空間級距、共識差距與價格帶", "head": ["項目", "數值"],
     "rows": [[VA(T[k]), VB(T[k], f)] for k, f in (("upw", "sp1"), ("upc", "sp1"), ("grade", ""), ("rec", ""), ("geps", "sp1"), ("band", ""), ("m", ""))]},
    {"type": "table", "caption": "DCF 折現率：WACC、股權成本、稅後債務成本、RONIC、永續成長率", "head": ["項目", "數值", "說明"],
     "rows": [[VA(SD["wacc"]), VB(SD["wacc"], "p2"), "股權成本與稅後債務成本以市值權重加權；WACC 調整只用於敏感度（基準為零）"], [VA(SD["ke"]), VB(SD["ke"], "p2"), c("Val", f"E{SD['ke']}")], [VA(SD["kd_at"]), VB(SD["kd_at"], "p2"), c("Val", f"E{SD['kd_at']}")],
              [VA(SD["w_e"]), VB(SD["w_e"], "p1"), c("Val", f"E{SD['w_e']}")], [VA(SD["g"]), VB(SD["g"], "p2"), c("Val", f"E{SD['g']}")], [VA(SD["ronic"]), VB(SD["ronic"], "p2"), c("Val", f"E{SD['ronic']}")]]},
    {"type": "table", "caption": c("Val", f"A{SD['title']}"), "head": [VA(SD["h"]), c("Val", f"B{SD['h']}"), c("Val", f"C{SD['h']}"), c("Val", f"D{SD['h']}")],
     "rows": [[LBL[k], c("Val", f"B{SD[k]}", "n2"), c("Val", f"C{SD[k]}", "n2"), c("Val", f"D{SD[k]}", "n2")] for k in ("pvx", "pvtv", "ev", "ncr", "dcfv", "v12")]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 彙總的其他項目", "head": [VA(SD["h"]), c("Val", f"B{SD['h']}"), c("Val", f"C{SD['h']}"), c("Val", f"D{SD['h']}")],
     "rows": [[VA(SD["nc"]), c("Val", f"B{SD['nc']}", "n2"), c("Val", f"C{SD['nc']}", "n2"), c("Val", f"D{SD['nc']}", "n2")], [VA(SD["stub"]), c("Val", f"B{SD['stub']}", "n2"), c("Val", f"C{SD['stub']}", "n2"), c("Val", f"D{SD['stub']}", "n2")],
              [VA(SD["sh"]), c("Val", f"B{SD['sh']}", "n3"), c("Val", f"C{SD['sh']}", "n3"), c("Val", f"D{SD['sh']}", "n3")], [VA(SD["tvsh"]), c("Val", f"B{SD['tvsh']}", "p1"), c("Val", f"C{SD['tvsh']}", "p1"), c("Val", f"D{SD['tvsh']}", "p1")],
              [VA(SD["mult"]), c("Val", f"B{SD['mult']}", "x1"), c("Val", f"C{SD['mult']}", "x1"), c("Val", f"D{SD['mult']}", "x1")], [VA(SD["dcf"]), c("Val", f"B{SD['dcf']}", "n2"), "", ""]]}]
D = wb["DCF"]
for s in "ABC":
    t = next(r for r in range(1, D.max_row + 1) if str(D.cell(r, 1).value or "").startswith(f"情境 {s}（"))
    val_blocks.append({"type": "table", "caption": c("DCF", f"A{t}"), "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細", "head": [c("DCF", "A5")] + [c("DCF", f"{x}5") for x in "CDEFGH"],
                       "rowgen": {"sheet": "DCF", "rows": f"{t + 1}-{t + 13}", "cols": ["A", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n3", "n3", "n3", "n3", "n3", "n3"],
                                  "row_fmt": {str(r): {x: "p1" for x in "CDEFGH"} for r in (t + 4, t + 6, t + 7)} | {str(t + 12): {x: "n3" for x in "CDEFGH"}}}})
val_blocks += [
    {"type": "note", "parts": [c("DCF", "A2")], "collapse": True, "summary": "展開：DCF 的營業利益與自由現金流口徑（含 Non-GAAP 橋接）"},
    {"type": "table", "caption": c("Val", f"A{T['t0'] + 23}"), "head": [VA(T["ph"]), c("Val", f"B{T['ph']}"), c("Val", f"C{T['ph']}"), c("Val", f"D{T['ph']}")],
     "rows": [[VA(r), c("Val", f"B{r}"), c("Val", f"C{r}", "x1"), c("Val", f"D{r}")] for r in range(T["p1"], T["p2"] + 1)] + [[VA(T["pmed"]), "", c("Val", f"B{T['pmed']}", "x1"), ""], [VA(T["feps"]), "", c("Val", f"B{T['feps']}", "n2"), ""], [VA(T["comptp"]), "", c("Val", f"B{T['comptp']}", "n2"), ""]]},
    {"type": "table", "caption": "相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司相對同業的長期倍數比）", "head": ["項目", "數值"],
     "rows": [[VA(rm["own"]), VB(rm["own"], "n2")], [VA(rm["pmed"]), VB(rm["pmed"], "n2")], [VA(rm["fac"]), VB(rm["fac"], "n2")], [VA(rm["adj"]), VB(rm["adj"], "n2")]]},
    {"type": "table", "collapse": True, "summary": "展開：相對倍數調整的其他項目", "head": ["項目", "數值"], "rows": [[VA(rm["npeer"]), VB(rm["npeer"], "n0")], [VA(rm["raw"]), VB(rm["raw"], "n2")]]},
    {"type": "table", "caption": "現價與市場隱含倍數", "head": [c("Val", "A5"), c("Val", "B5"), c("Val", "C5"), c("Val", "D5")],
     "rows": [[c("Val", "A6"), c("Val", "B6", "n2"), "", ""]] + [[c("Val", f"A{r}"), c("Val", f"B{r}", "x1"), c("Val", f"C{r}", "x1"), c("Val", f"D{r}", "x1")] for r in (7, 8, 9, 10)]},
    {"type": "note", "parts": [c("Val", "A2")]}]
sec_val = {"id": "valuation", "title": "估值", "blocks": val_blocks}

# ── 敏感度 ──
S = wb["Sens"]
cases = []   # (列, 低相對, 高相對)
for r in list(range(M["FD_FIRST"], M["FD_LAST"] + 1)) + list(range(M["BR0"] + 2, M["BIZ_LAST"] + 1)):
    if not S.cell(r, 1).value or str(S.cell(r, 1).value).startswith("檢查"): continue
    lo, hi = S.cell(r, 4).value, S.cell(r, 6).value
    cases.append((r, lo if isinstance(lo, (int, float)) else None, hi if isinstance(hi, (int, float)) else None))
big_lo = [r for r, lo, hi in cases if lo is not None and abs(lo) >= 0.10]
big_hi = [r for r, lo, hi in cases if hi is not None and abs(hi) >= 0.10]
sm_lo = [r for r, lo, hi in cases if lo is not None and abs(lo) < 0.10]
sm_hi = [r for r, lo, hi in cases if hi is not None and abs(hi) < 0.10]
tn = lambda col, rows, title, **kw: {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1", "title": title, "rows": ",".join(map(str, rows)), **kw}
sens_blocks = [tn("D", big_lo or sm_lo[:3], "加權目標價相對基準：輸入取低值" + ("，變動達兩位數百分比的案例" if big_lo else "（無達兩位數百分比者，列前三項）")),
               tn("F", big_hi or sm_hi[:3], "加權目標價相對基準：輸入取高值" + ("，變動達兩位數百分比的案例" if big_hi else "（無達兩位數百分比者，列前三項）")),
               {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入；變動達兩位數百分比者直接顯示，其餘摺疊）；下方摺疊為全部案例與說明。"]}]
if big_lo and sm_lo: sens_blocks.append(tn("D", sm_lo, "加權目標價相對基準：輸入取低值", collapse=True, summary="展開：加權目標價相對基準：輸入取低值，變動未達兩位數百分比的案例"))
if big_hi and sm_hi: sens_blocks.append(tn("F", sm_hi, "加權目標價相對基準：輸入取高值", collapse=True, summary="展開：加權目標價相對基準：輸入取高值，變動未達兩位數百分比的案例"))
sens_blocks += [
    {"type": "table", "caption": c("Sens", f"A{M['TPS0']}"), "collapse": True, "summary": "展開：目標價參數表（折現率、永續成長率、DCF 權重等；閉式公式）", "head": [c("Sens", f"{x}{M['TPS0'] + 1}") for x in "ABCDEFG"],
     "rowgen": {"sheet": "Sens", "rows": f"{M['FD_FIRST']}-{M['FD_LAST']}", "cols": list("ABCDEFG"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2"]}},
    {"type": "table", "caption": c("Sens", f"A{M['BR0']}"), "collapse": True, "summary": "展開：業務主因表（整本重建快照）", "head": [c("Sens", f"{x}{M['BR0'] + 1}") for x in "ABCDEFG"],
     "rowgen": {"sheet": "Sens", "rows": f"{M['BR0'] + 2}-{M['BIZ_LAST']}", "cols": list("ABCDEFG"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2"]}},
    {"type": "note", "parts": [c("Sens", "A1")], "collapse": True, "summary": "展開：敏感度頁原文說明"}]
sec_sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}

# ── 反向 DCF ──
sec_rev = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "blocks": [
    {"type": "note", "parts": [c("RevDCF", "A14")]}, {"type": "note", "parts": [c("RevDCF", "A2")]},
    {"type": "table", "caption": c("RevDCF", "A4"), "head": ["項目", "現價隱含", "模型基準"],
     "rows": [[c("RevDCF", "A6"), c("RevDCF", "B6", "n2"), c("RevDCF", "C6", "n2")], [c("RevDCF", "A7"), c("RevDCF", "B7", "n2"), c("RevDCF", "C7", "n2")], [c("RevDCF", "A8"), c("RevDCF", "B8", "n1"), c("RevDCF", "C8", "n1")],
              [c("RevDCF", "A9"), c("RevDCF", "B9", "n1"), c("RevDCF", "C9", "n1")], [c("RevDCF", "A10"), c("RevDCF", "B10", "n1"), c("RevDCF", "C10", "n1")], [c("RevDCF", "A11"), c("RevDCF", "B11", "p1"), c("RevDCF", "C11", "p1")],
              [c("RevDCF", "A12"), c("RevDCF", "B12", "p1"), c("RevDCF", "C12", "p1")]]}]}

# ── 損益 ──
ph = lambda k: PB[k]
hdrB = PB["B"]["mkt"] - 1
sumA = lambda k: SUMR[f"A|{k}"]; sumB = lambda k: SUMR[f"B|{k}"]; sumC = lambda k: SUMR[f"C|{k}"]
trip = lambda k, cols="C{r}:J{r}": {s: c("PL", f"C{SUMR[f'{s}|{k}']}:J{SUMR[f'{s}|{k}']}") for s in "ABC"}
SHR = M["SRW"]
r0 = SHR["營收"] - 1
pl_blocks = [
    {"type": "note", "parts": [c("PL", "A1")]},
    {"type": "theme", "title": "三情境並列：營收（長條，十億美元）與 EPS（折線，美元，右軸）", "x": c("PL", "C3:J3"), "fmt": "n1", "fmt_r": "n2", "stack": False, "pad": 1.15,
     "bars": [{"name": ["情境 A（樂觀）營收"], "y": c("PL", f"C{sumA('rev')}:J{sumA('rev')}"), "color": "#b7e3e4"}, {"name": ["情境 B（基準）營收"], "y": c("PL", f"C{sumB('rev')}:J{sumB('rev')}"), "color": "#10a0a7"}, {"name": ["情境 C（保守）營收"], "y": c("PL", f"C{sumC('rev')}:J{sumC('rev')}"), "color": "#8a8c8c"}],
     "lines": [{"name": ["情境 A（樂觀）EPS"], "y": c("PL", f"C{sumA('eps')}:J{sumA('eps')}"), "axis": "r", "color": "#0b7a80", "labels": False}, {"name": ["情境 B（基準）EPS"], "y": c("PL", f"C{sumB('eps')}:J{sumB('eps')}"), "axis": "r", "color": "#c29a2b", "labels": True},
               {"name": ["情境 C（保守）EPS"], "y": c("PL", f"C{sumC('eps')}:J{sumC('eps')}"), "axis": "r", "color": "#6f6f6f", "labels": False}]},
    {"type": "table", "caption": c("PL", f"A{r0 - 1}"), "head": [c("PL", f"{x}{r0}") for x in "ABCDEFGHIJKL"],
     "rows": [[c("PL", f"A{SHR[l]}")] + [c("PL", f"B{SHR[l]}")] + [c("PL", f"{x}{SHR[l]}", f) for x in "CDEFGHIJKL"] for l, f in (
         ("營收", "n2"), ("營收年增率", "p1"), ("營業成本", "n2"), ("毛利", "n2"), ("毛利率", "p1"), ("營業費用", "n2"), ("營業利益", "n2"), ("營業利益率", "p1"), ("其他收入（業外淨額）", "n2"), ("稅前利益", "n2"),
         ("所得稅", "n2"), ("稅後淨利", "n2"), ("淨利率", "p1"), ("EPS（GAAP，稀釋）", "n2"), ("淨利年增率", "p1"), ("Non-GAAP：調整後營業利益（公司口徑，對照）", "n2"))]},
    {"type": "table", "caption": "標準損益表核對列（恆等式與分部合計，應為零）", "collapse": True, "summary": "展開：標準損益表的核對列", "head": [c("PL", f"{x}{r0}") for x in "ABCDEFGHIJKL"],
     "rows": [[c("PL", f"A{M['SCHK']}"), ""] + [c("PL", f"{x}{M['SCHK']}", "n3") for x in "CDEFGHIJKL"], [c("PL", f"A{M['SCHK'] + 1}"), ""] + [""] * 4 + [c("PL", f"G{M['SCHK'] + 1}", "n3")] + [""] * 5,
              [c("PL", f"A{M['SEGR'] + 3}"), ""] + ["", ""] + [c("PL", f"E{M['SEGR'] + 3}", "n3")] + [""] * 6]},
    {"type": "pl", "years": c("PL", "C3:J3"), "metrics": [
        {"name": ["營收（", c("PL", f"B{sumA('rev')}"), "）"], "fmt": "n1", "chart": True, "rows": trip("rev")},
        {"name": ["資料中心設施產品營收（", c("PL", f"B{sumA('dcp')}"), "）"], "fmt": "n1", "rows": {s: c("PL", f"C{SUMR[f'{s}|dcp']}:G{SUMR[f'{s}|dcp']}") for s in "ABC"}},
        {"name": "毛利率", "fmt": "p1", "rows": trip("gm")}, {"name": "營業利益率", "fmt": "p1", "chart": True, "ymin": "auto", "rows": trip("opm")},
        {"name": ["EPS（", c("PL", f"B{sumA('eps')}"), "）"], "fmt": "n2", "chart": True, "rows": trip("eps")}]}]
sec_pl = {"id": "pl", "title": "損益（標準損益表）", "blocks": pl_blocks}

# ── 驅動因子 ──
IW = wb["In"]
secs = []; cur = None
for r in range(4, IW.max_row + 1):
    a, b = IW.cell(r, 1).value, IW.cell(r, 2).value
    if a and not b: cur = {"title": str(a), "rows": []}; secs.append(cur)
    elif a and cur is not None: cur["rows"].append(r)
fm = lambda u: {"%": "p1", "pt": "p1", "x": "n2", "US$B": "n3", "US$": "n2", "十億股": "n3", "年": "n3", "日期": ""}.get(u, "n2")
SEC_NAMES = ["一、機房設施市場（Wattonomics 讀取）", "二、情境（個股自訂）", "三、營收結構：份額與其餘營收", "四、損益", "五、FY2026 錨點（實績與第四季假設）", "六、長期收斂與目標價規則", "七、折現率與自由現金流"]
def in_tbl(rows, collapse, title, summary=None):
    b = {"type": "table", "caption": title, "head": [c("In", "B3"), c("In", "D3"), c("In", "E3"), c("In", "F3"), c("In", "C3"), c("In", "G3")],
         "rows": [[c("In", f"B{r}"), c("In", f"D{r}", fm(IW.cell(r, 3).value)), c("In", f"E{r}", fm(IW.cell(r, 3).value)), c("In", f"F{r}", fm(IW.cell(r, 3).value)), c("In", f"C{r}"), c("In", f"G{r}", "tag")] for r in rows]}
    if collapse: b["collapse"] = True; b["summary"] = summary
    return b
drv = [
    {"type": "note", "parts": [R_("mkt")]},
    {"type": "theme", "title": "可及市場（長條，十億美元）與 Vertiv 資料中心設施產品營收（折線，十億美元，右軸）：情境 B", "x": c("PL", f"C{hdrB}:G{hdrB}"), "fmt": "n0", "fmt_r": "n1", "stack": False, "pad": 1.3,
     "bars": [{"name": ["可及市場（Wattonomics 配電、UPS、冷卻）"], "y": c("PL", f"C{PB['B']['mkt']}:G{PB['B']['mkt']}"), "color": "#6fc7cb"}],
     "lines": [{"name": ["Vertiv 資料中心設施產品營收"], "y": c("PL", f"C{PB['B']['dcp']}:G{PB['B']['dcp']}"), "axis": "r", "color": "#0b7a80", "labels": True}],
     "caption": "營收＝市場 × 份額：份額 FY2026 由實績校準，FY2027 起乘以份額相對值；市場讀 Wattonomics，情境以 FY2028–FY2029 倍數表達。"},
    {"type": "table", "caption": "機房設施市場與 Vertiv 份額（Wattonomics 讀取；本專案推導份額）", "head": [c("Units", f"{x}3") for x in "ABCDEFG"],
     "rowgen": {"sheet": "Units", "rows": "4-13", "cols": list("ABCDEFG"), "fmts": ["", "", "n2", "n2", "n2", "n2", "n2"], "row_fmt": {"10": {x: "p1" for x in "CDEFG"}, "12": {x: "p1" for x in "CDEFG"}, "13": {x: "p1" for x in "CDEFG"}}}},
    {"type": "note", "parts": [c("Units", "A2")], "collapse": True, "summary": "展開：機房設施邊界與 Wattonomics 後期下降的處理"}]
for s_ in secs:
    rngs = [r for r in s_["rows"] if isinstance(IW.cell(r, 4).value, (int, float)) and isinstance(IW.cell(r, 6).value, (int, float)) and IW.cell(r, 4).value != IW.cell(r, 6).value]
    fixed = [r for r in s_["rows"] if r not in rngs]
    drv.append({"type": "h3", "text": SEC_NAMES[secs.index(s_)]})
    if rngs: drv.append(in_tbl(rngs, False, "有敏感度區間的參數（基準進目標價；低／高只用於敏感度頁）"))
    if fixed: drv.append(in_tbl(fixed, True, "其他固定輸入（實績、公司揭露、換算值）", "展開：其他固定輸入（實績、公司揭露、換算值）"))
sec_drv = {"id": "drivers", "title": "驅動因子：設施市場 × 份額", "blocks": drv}

# ── 對帳 ──
RC = wb["Recon"]
ry = next(r for r in range(1, RC.max_row + 1) if str(RC.cell(r, 1).value or "").startswith("Yahoo 與公司原文的差異"))
rec_blocks = [
    {"type": "h3", "text": "三層協調：讀 Wattonomics 層"},
    {"type": "note", "parts": [c("Coord", "A2")]}, {"type": "note", "parts": [c("Coord", "A18")]}, {"type": "note", "parts": [c("Coord", "A19")]},
    {"type": "table", "caption": c("Coord", "A10"), "head": [c("Coord", "A10"), c("Coord", "B10"), c("Coord", "C10")],
     "rows": [[c("Coord", f"A{r}"), c("Coord", f"B{r}", f), c("Coord", f"C{r}", f)] for r, f in ((11, "n2"), (12, "n2"), (13, "p1"), (14, "sp1"), (15, "p1"), (16, "p1")) if r != 11 or True]},
    {"type": "table", "caption": "展開：讀取的層級、版本與用途", "collapse": True, "summary": "展開：本模型讀哪一層、哪一版、用在哪裡", "head": [c("Coord", f"{x}4") for x in "ABCD"],
     "rows": [[c("Coord", f"{x}{r}") for x in "ABCD"] for r in (5, 6, 7)]},
    {"type": "h3", "text": "公司指引、市場共識與 Yahoo 差異"},
    {"type": "table", "caption": "本模型與公司指引、共識的對照（共識為調整後口徑，本模型為 GAAP）", "head": [c("Recon", f"{x}3") for x in "ABCD"] + [c("Recon", "F3")],
     "rows": [[c("Recon", f"A{r}"), c("Recon", f"B{r}", "n2"), c("Recon", f"C{r}", "n2"), c("Recon", f"D{r}", "sp1"), c("Recon", f"F{r}")] for r in range(4, ry - 2)]},
    {"type": "table", "caption": "展開：Yahoo 與公司原文的差異（公司原文優先）", "collapse": True, "summary": "展開：Yahoo 與公司原文的差異", "head": [c("Recon", f"{x}{ry + 1}") for x in "ABCD"] + [c("Recon", f"J{ry + 1}")],
     "rows": [[c("Recon", f"A{r}"), c("Recon", f"B{r}", "n3"), c("Recon", f"C{r}", "n3"), c("Recon", f"D{r}", "sp1"), c("Recon", f"J{r}")] for r in range(ry + 2, RC.max_row + 1)]}]
sec_rec = {"id": "recon", "title": "對帳：公司指引、共識與設施市場", "intro": c("Recon", "A1"), "blocks": rec_blocks}

# ── 資料狀態 ──
sec_status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [R_("ver")]},
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋（取代建置 commit）", "$sha"], ["敏感度建置識別", c("Sens", f"D{M['BID_ROW']}")],
                             ["機房設施市場來源（Wattonomics 快照）", c("Coord", "B6")], ["快照雜湊", c("Coord", "C6")]]},
    {"type": "h3", "text": "待判斷（已先決定並標範圍）"}, {"type": "list", "ordered": True, "items": [R_(f"todo{i}") for i in range(1, 12)]},
    {"type": "h3", "text": "建議補進 Siliconomics／Wattonomics"}, {"type": "list", "ordered": True, "items": [R_(f"sug{i}") for i in range(1, 5)]},
    {"type": "h3", "text": "已知限制"}, {"type": "note", "parts": [R_("lim")]},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的輸入與來源", "sheets": {"In": ["B", "I"], "Q": ["A", "J"], "Sources": ["A", "B"]}, "empty": "沒有標為未找到的項目。"}]}

cfg = {"company": "Vertiv", "ticker": "VRT", "thesis": R_("p1"), "question": "設施市場在 Wattonomics 基準下成長放緩後，Vertiv 的份額與利潤率能否撐住現價隱含的成長",
       "scenarios": [{"key": "A", "name": "情境 A"}, {"key": "B", "name": "情境 B"}, {"key": "C", "name": "情境 C"}],
       "sections": [sec_thesis, sec_val, sec_sens, sec_rev, sec_pl, sec_drv, sec_rec, sec_status], "wide_text": True,
       "hero": {"rec": c("Val", f"B{T['rec']}"), "tiles": [[VA(T["tpw"]), VB(T["tpw"], "n2"), ["上檔空間 ", VB(T["upw"], "sp1")]],
                                                   ["只看 Comparable 目標價（市場主流同業法）", VB(T["comp2"], "n2"), ["上檔空間 ", VB(T["upc"], "sp1")]], [c("Val", "A6"), c("Val", "B6", "n2"), ""],
                                                   [VA(T["cons"]), VB(T["cons"], "n2"), ["相對現價 ", c("Val", f"C{T['cons']}", "sp1")]]], "note": [R_("banner")]},
       "layout": "v1.3"}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("wrote html.json", len(json.dumps(cfg, ensure_ascii=False)))

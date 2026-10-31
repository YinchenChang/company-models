# 產生 tuc/html.json（HTML 展示層設定）。用法：TUC_ROWMAP=rowmap.json python3 tuc/build_tuc.py raw.xlsx（取得列號），重算成 dist xlsx 後：python3 tuc/gen_html_json.py <dist xlsx> tuc/html.json rowmap.json
import json, sys, openpyxl
rm = json.load(open(sys.argv[3]))
xl = sys.argv[1]
wb = openpyxl.load_workbook(xl, data_only=True)
IR = rm["INROW"]; PB = rm["PB"]; SUMR = rm["SUMR"]; T = rm["T"]; XR = rm["XR"]; XH = rm["XH"]
S = lambda s, k: SUMR[f"{s}|{k}"]
# 情境驅動對照
INS = wb["In"]
scn = next(r for r in range(1, INS.max_row + 1) if "情境驅動對照" in str(INS.cell(r, 1).value or ""))
scn_rows = [r for r in range(scn + 2, INS.max_row + 1) if INS.cell(r, 2).value]
def drv(key, fmt, dlabel, dref, dfmt, why, unit=True):
    r = IR[key]
    base = [f"@In!E{r}|{fmt}"] + ([" ", f"@In!C{r}"] if unit else [])
    return [f"@In!B{r}", base, f"@In!G{r}|tag", [dlabel, f"{dref}|{dfmt}"], why]
A = PB["A"]
drivers = [
    drv("gw_27", "n1", "部署 GW（FY2027）：", f"@Units!D4", "n1", "讀 Siliconomics 部署路徑，不在本模型另訂；情境 B、C 從 CY2028 起縮放"),
    drv("pgw_27", "n0", "AI 高速 CCL 市場（FY2027）：", f"@Units!D6", "n2", "每 GW 板數、層數與材料等級上升；讀 Siliconomics，絕對水準只決定校準出的份額"),
    ["@PL!A" + str(A["sh0"]), [f"@PL!D{A['sh0']}|p1"], "Derived", ["高速材料營收（FY2026）：", f"@PL!D{A['ai']}|n1"], "以公司損耗等級組合（極低與超低損耗占比）校準"],
    drv("rel_28", "n2", "台燿有效份額（FY2028）：", "@Units!E7", "p1", "新廠開出後份額上升；放公司 In 頁，待併入 Siliconomics", unit=False),
    drv("pp_28", "n2", "高速材料毛利率（FY2028）：", f"@PL!F{A['gma']}", "p1", "每張成本不變，價格回落全部落在毛利", unit=False),
    drv("cap_27", "n2", "產能上限營收（FY2027）：", f"@PL!E{A['rcap']}", "n1", "全產全銷：需求面高於產能時，營收由月產能決定"),
    drv("mix_28", "n2", "每張營收（FY2028）：", "@Units!E13", "n0", "高階材料占比上升，每張營收上升", unit=False),
    drv("g_nai27", "p0", "其他等級營收（FY2027）：", f"@PL!E{A['nai']}", "n1", "一般伺服器、四百 G 以下網通、厚銅與電源；與 AI 部署無關", unit=False),
    drv("d_ai", "p0", "毛利率（FY2027）：", f"@PL!E{A['gm']}", "p1", "高速材料占比上升使毛利率上升", unit=False),
    drv("phi", "n2", "營業利益率（FY2027）：", f"@PL!E{A['oim']}", "p1", "營業費用成長慢於營收", unit=False),
]
fm = lambda c: c
html = {
 "company": "台燿", "ticker": "6274", "thesis": "@README!A3",
 "question": "新產能開出時，台燿在人工智慧伺服器與交換器的高階銅箔基板份額與單價能撐多久；份額與價格回落時獲利還剩多少",
 "scenarios": [{"key": "A", "name": "情境 A"}, {"key": "B", "name": "情境 B"}, {"key": "C", "name": "情境 C"}],
 "sections": [
  {"id": "thesis", "title": "命題與核心問題", "intro": "@README!A4", "blocks": [
     {"type": "kv", "items": [["情境 A", f"@PL!A{PB['A']['gw'] - 3}"], ["情境 B", f"@PL!A{PB['B']['gw'] - 3}"], ["情境 C", f"@PL!A{PB['C']['gw'] - 3}"], ["會計年度與口徑", "@README!A5"]]}]},
  {"id": "drivers", "title": "驅動因子 → 推導量", "intro": "標記沿用 Excel：Verified、Interested-party、Analogy、Assumed、Derived。輸入取自 In 頁，推導量取自 PL 頁情境 A 與 Units 頁。", "blocks": [
     {"type": "table", "caption": "單一參數（In 頁基準值；敏感度低／高只用於敏感度頁、不進目標價）", "head": ["輸入", "基準值", "標記", "推導量", "因果方向"], "rows": drivers},
     {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；取自 In 頁最後一節「情境驅動對照」）",
      "head": [f"@In!B{scn + 1}", f"@In!D{scn + 1}", f"@In!E{scn + 1}", f"@In!F{scn + 1}", f"@In!C{scn + 1}"],
      "rows": [[f"@In!B{r}", f"@In!D{r}|n2", f"@In!E{r}|n2", f"@In!F{r}|n2", f"@In!C{r}"] for r in scn_rows]},
     {"type": "table", "caption": "AI 高速 CCL 市場、台燿份額與產能；與台光電的份額一致性", "head": ["@Units!A3", "@Units!B3", "@Units!C3", "@Units!D3", "@Units!E3", "@Units!F3"],
      "rowgen": {"sheet": "Units", "rows": f"4-{rm['UR']['oth']}", "cols": ["A", "B", "C", "D", "E", "F"], "fmts": ["", "", "n2", "n2", "n2", "n2"],
                 "row_fmt": {str(rm["UR"][k]): {c: f for c in "CDEF"} for k, f in (("gw", "n1"), ("pgw", "n0"), ("sh", "p1"), ("sha", "p1"), ("pgwe", "n0"), ("rps", "n0"), ("util", "p1"), ("esh", "p1"), ("sum", "p1"), ("suma", "p1"), ("oth", "p1"))}}},
     {"type": "note", "parts": [f"@Units!A{rm['UNOTE']}"]}]},
  {"id": "pl", "title": "三情境損益", "intro": "FY2025 為實績，FY2026 為錨點（上半年與第三季財報＋第四季假設），FY2027 起為模型。三情境在 FY2027 相同，差異從 CY2028 的顆數情境開始。", "blocks": [
     {"type": "pl", "years": "@PL!C3:G3", "metrics": [
        {"name": ["營收（", f"@PL!B{S('A', 'rev')}", "）"], "fmt": "n1", "chart": True, "rows": {s: f"@PL!C{S(s, 'rev')}:G{S(s, 'rev')}" for s in "ABC"}},
        {"name": ["高速材料營收（", f"@PL!B{S('A', 'ai')}", "）"], "fmt": "n1", "chart": True, "rows": {s: f"@PL!C{S(s, 'ai')}:G{S(s, 'ai')}" for s in "ABC"}},
        {"name": "需求面 ÷ 產能上限", "fmt": "p1", "rows": {s: f"@PL!C{S(s, 'util')}:G{S(s, 'util')}" for s in "ABC"}},
        {"name": "毛利率", "fmt": "p1", "chart": True, "ymin": "auto", "rows": {s: f"@PL!C{S(s, 'gm')}:G{S(s, 'gm')}" for s in "ABC"}},
        {"name": "營業利益率", "fmt": "p1", "rows": {s: f"@PL!C{S(s, 'oim')}:G{S(s, 'oim')}" for s in "ABC"}},
        {"name": ["EPS（", f"@PL!B{S('A', 'eps')}", "）"], "fmt": "n2", "chart": True, "rows": {s: f"@PL!C{S(s, 'eps')}:G{S(s, 'eps')}" for s in "ABC"}}]},
     {"type": "kv", "items": [[f"@PL!A{rm['CHK'] + i}", f"@PL!D{rm['CHK'] + i}|n3"] for i in range(4)]},
     {"type": "h3", "text": "延伸到 FY2032（FY2030 起為收斂法）"},
     {"type": "note", "parts": [f"@PL!A{rm['XT']}"]},
     {"type": "pl", "years": f"@PL!C{XH}:J{XH}", "metrics": [
        {"name": ["營收（", f"@PL!B{XR['A']['rev']}", "）"], "fmt": "n1", "chart": True, "rows": {s: f"@PL!C{XR[s]['rev']}:J{XR[s]['rev']}" for s in "ABC"}},
        {"name": "淨利率", "fmt": "p1", "rows": {s: f"@PL!C{XR[s]['nim']}:J{XR[s]['nim']}" for s in "ABC"}},
        {"name": ["EPS（", f"@PL!B{XR['A']['eps']}", "）"], "fmt": "n2", "chart": True, "rows": {s: f"@PL!C{XR[s]['eps']}:J{XR[s]['eps']}" for s in "ABC"}}]}]},
 ]}
V = wb["Val"]
def vrow(k, fmt, three=False):
    r = T[k]
    return [f"@Val!A{r}", f"@Val!B{r}|{fmt}"] + ([f"@Val!C{r}|{fmt}", f"@Val!D{r}|{fmt}"] if three else ["", ""])
rmrow = {str(V.cell(r, 1).value): r for r in range(1, V.max_row + 1) if V.cell(r, 1).value}
rmk = [rmrow[x] for x in ("本公司 10 年中位數（納入值）", "同業 10 年中位數的中位數", "原始比值（本公司 ÷ 同業）", "相對倍數調整係數", "Comparable 目標價（相對倍數調整後）")]
OB = rm["OBSROW"]
val_blocks = [
  {"type": "table", "caption": "目標價（DCF 與 Comparable 加權）", "head": [f"@Val!A{T['h3']}", f"@Val!B{T['h3']}", "", ""],
   "rows": [vrow("tpw", "n0"), vrow("comp2", "n0"), vrow("cons", "n0"), vrow("upw", "sp1"), vrow("upc", "sp1"), [f"@Val!A{T['grade']}", f"@Val!B{T['grade']}", "", ""],
            vrow("geps", "sp1"), vrow("gmul", "sp1"), [f"@Val!A{T['band']}", f"@Val!B{T['band']}", "", ""], vrow("pr", "p0", True), [f"@Val!A{T['m']}", f"@Val!B{T['m']}", "", ""]]},
  {"type": "note", "parts": [f"@Val!A{T['rec']}", "：", f"@Val!B{T['rec']}", "（依上檔空間固定規則）"]},
  {"type": "h3", "text": "相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司相對同業的長期倍數比）"},
  {"type": "kv", "items": [[f"@Val!A{r}", f"@Val!B{r}|{'n0' if i == 4 else 'n2'}"] for i, r in enumerate(rmk)]},
  {"type": "h3", "text": "重點觀察"},
  {"type": "list", "items": [f"@README!A{OB + 1}", f"@README!A{OB + 2}", f"@README!A{OB + 3}"]},
  {"type": "table", "caption": "@Val!A5", "head": [f"@Val!A{T['h1']}", f"@Val!B{T['h1']}", f"@Val!C{T['h1']}", f"@Val!D{T['h1']}"],
   "rows": [[f"@Val!A{T['eps_h']}", "", "", ""]] + [vrow(f"eps{y}", "n2", True) for y in range(2027, 2033)] + [[f"@Val!A{T['cf_h']}", "", "", ""], vrow("reinv", "p1")]
           + [vrow(f"cf{y}", "n2", True) for y in range(2027, 2033)] + [vrow("dcfv", "n0", True), vrow("dcf12", "n0", True), vrow("dcftp", "n0")]},
  {"type": "bars", "title": "DCF 情境價值（推到十二個月後）與現價", "items": [{"label": f"情境 {s}", "value": f"@Val!{c}{T['dcf12']}", "series": i} for i, (s, c) in enumerate(zip("ABC", "BCD"))],
   "ref": {"label": "現價", "value": f"@Val!B{T['px']}"}, "fmt": "n0"},
  {"type": "table", "caption": f"@Val!A{T['ph'] - 1}", "head": [f"@Val!A{T['ph']}", f"@Val!B{T['ph']}", f"@Val!C{T['ph']}", f"@Val!D{T['ph']}", f"@Val!F{T['ph']}"],
   "rows": [[f"@Val!A{r}", f"@Val!B{r}", f"@Val!C{r}|x1", f"@Val!D{r}", f"@Val!F{r}"] for r in range(T["p1"], T["p2"] + 1)]
           + [[f"@Val!A{T['pmed']}", f"@Val!B{T['pmed']}|x1", "", "", ""], [f"@Val!A{T['feps']}", f"@Val!B{T['feps']}|n2", "", "", ""], [f"@Val!A{T['comptp']}", f"@Val!B{T['comptp']}|n0", "", "", ""]]},
  {"type": "note", "parts": [f"@Val!E{T['pmed']}"]},
  {"type": "table", "caption": f"@Val!A{T['pe10h']}", "collapse": True, "summary": "展開：十年年底市盈率（分位檢查用）",
   "head": [f"@Val!{c}{T['pe10h']}" for c in "ABCDEFGHIJK"], "rows": [[f"@Val!A{T['pe10']}"] + [f"@Val!{c}{T['pe10']}|x1" for c in "BCDEFGHIJK"]]},
  {"type": "table", "caption": "現價與市場隱含倍數", "head": [f"@Val!A{T['s1'] + 1}", f"@Val!B{T['s1'] + 1}", f"@Val!C{T['s1'] + 1}", f"@Val!D{T['s1'] + 1}"],
   "rows": [vrow("px", "n0")] + [vrow(f"pe{y}", "x1", True) for y in range(2026, 2030)] + [vrow("ttm", "n2"), vrow("pettm", "x1")]},
  {"type": "note", "parts": ["@Val!A2"]}]
html["sections"].append({"id": "valuation", "title": "估值", "intro": "@Val!A1", "blocks": val_blocks})
RL = rm["RECON_LAST"]
RC = wb["Recon"]
rf = {}
for r in range(4, RL + 1):
    lab = str(RC.cell(r, 1).value or "")
    if "份額" in lab or "產能" in lab or "毛利率" in lab: rf[str(r)] = {c: "p1" for c in "BCEG"}
    if "每張營收" in lab: rf[str(r)] = {c: "n0" for c in "BCEG"}
    if "Siliconomics" in lab: rf[str(r)] = {"C": "n2", "D": "n3", "E": "n2", "F": "n3", "G": "n2", "H": "n3"}
html["sections"].append({"id": "recon", "title": "對帳：共識、券商與 Siliconomics 市場", "intro": "@Recon!A1", "blocks": [
  {"type": "table", "head": [f"@Recon!{c}3" for c in "ABCDEFGHIJ"], "rowgen": {"sheet": "Recon", "rows": f"4-{RL}", "cols": list("ABCDEFGHIJ"),
   "fmts": ["", "n1", "n1", "sp1", "n1", "sp1", "n1", "sp1", "", ""], "row_fmt": rf}}]})
SL = rm["SENS_LAST"]
html["sections"].append({"id": "sens", "title": "敏感度", "intro": "@Sens!A1", "blocks": [
  {"type": "tornado", "title": "加權目標價相對基準的變動（一次改一個輸入的敏感度低／高值）", "sheet": "Sens", "rows": f"5-{SL}", "label_col": "A", "value_col": "E", "fmt": "sp1"},
  {"type": "note", "parts": ["@Sens!A2"]},
  {"type": "table", "collapse": True, "summary": "展開：全部案例的加權目標價與情境 A 的 FY2028 EPS", "head": [f"@Sens!{c}3" for c in "ABCDEFGH"],
   "rowgen": {"sheet": "Sens", "rows": f"4-{SL}", "cols": list("ABCDEFGH"), "fmts": ["", "", "", "n0", "sp1", "", "n2", "sp1"]}},
  {"type": "note", "parts": [f"@Sens!A{rm['SENS_NOTE']}"]}]})
html["sections"].append({"id": "status", "title": "資料狀態", "blocks": [
  {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋（取代建置 commit）", "$sha"],
                           ["行業變數來源", "@In!A2"]]},
  {"type": "h3", "text": "上游成品（市場層與其他公司模型）"},
  {"type": "upstream", "sheets": ["In", "Units"]},
  {"type": "h3", "text": "待複核項目"},
  {"type": "list", "items": ["財報數字經 HiStock、nstock 轉載公開資訊觀測站，原文待人工複核；法說內容經富果轉述。", "@Q!A2", "標為「未找到」的項目（見下方清單）待找到來源後補上。"]},
  {"type": "h3", "text": "已知限制"},
  {"type": "note", "parts": ["@README!A8"]},
  {"type": "h3", "text": "「未找到」清單"},
  {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的輸入與來源", "sheets": {"In": ["B", "H", "I"], "Q": ["A", "I"], "Val": ["A", "E"], "Sources": ["A", "B"]}, "empty": "沒有標為未找到的項目。"}]})
json.dump(html, open(sys.argv[2], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

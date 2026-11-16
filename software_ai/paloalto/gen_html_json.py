# 由重算後的成品 xlsx 以「列標籤」找列號，產生 html.json（列號隨模型變動，勿手改 html.json）。
# 用法：python3 gen_html_json.py <成品 xlsx> html.json
import json, sys, openpyxl
X, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(X, data_only=True)

def find(sh, text, col=1, start=1, exact=False):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        v = str(ws.cell(r, col).value or "")
        if (v == text) if exact else v.startswith(text): return r
    raise KeyError((sh, text))

V = {k: find("Val", t) for k, t in dict(
    wacc_h="折現率", rf="無風險利率", ke="股權成本", kd="稅後債務成本", wacc="WACC", g="永續成長率", ronic="RONIC", px="現價", tp="加權目標價", dcf="DCF 目標價（三情境", cmp="Comparable 目標價（同業", cons="市場共識", up="加權目標價上檔",
    rec="投資建議", scn_tp="情境目標價 A", gate_scn="收支關卡 A", dcf_h="DCF 彙總", pvx="明確期 FCFF", dcfv="DCF 目標價", tvsh="永續價值占 EV", peh="同業本益比（Comparable）", pe1="同業前瞻本益比：Fortinet", pe_rel="PANW 資料商前瞻", pxpe="PANW 現價 ÷", per_h="期間與評等邊界", att2="上檔空間歸因：Comparable", pemed="同業中位數", rel="相對倍數調整係數", eps="FY2028 Non-GAAP EPS", cmpabc="Comparable 目標價（A／B／C）", lvl="級距", valr="估值評等", gate="入不敷出判定用", gtxt="入不敷出判定", wts="情境權重", sdcf="情境 DCF 目標價", scmp="情境 Comparable 目標價").items()}
SC = {k: find("Sens", t) for k, t in dict(hdr="案例", rel="加權目標價相對基準", note="未涵蓋說明").items()}
rel_hdr = SC["rel"] + 1
rel_rows = [r for r in range(rel_hdr + 1, SC["note"]) if wb["Sens"].cell(r, 1).value]
case_first, case_last = SC["hdr"] + 1, SC["rel"] - 2
R = {k: find("RevDCF", t) for k, t in dict(px="現價", mrow="模型核心", imp="現價隱含核心", interp="解讀", ev0="現價隱含企業價值").items()}
RC = {"h1": find("Recon", "項目", start=3), "s2": find("Recon", "2. 逐分部"), "s3": find("Recon", "3. 一次性"), "s4": find("Recon", "4. 同業"), "s5": find("Recon", "5. Tokenomics"), "s6": find("Recon", "6. 風險")}
P = {k: find("PL", t) for k, t in dict(prod="Product", ss="Subscription and support", rev="合併營收", chk1="檢查：Product＋", n="受保護代理身分數", rho="每單位營收", arr="NGS ARR（期末）", incr="AI 營收增量（代理", eros="定價侵蝕（相對", inf="推論成本增量（市場價", net="AI 淨效果", cpt="每次成功分析", fcff="FCFF（無槓桿", is0="標準損益表年度", io="營收", yoy_n="淨利年增率", isnote="標準損益表說明", ck1="檢查：標準損益表", ck2="檢查：營收 ＝", opm="營業利益率").items()}
IS0 = find("PL", "營收", start=P["is0"] + 1, exact=True)
IS_END = find("PL", "EPS（完全稀釋", start=IS0)
NYOY = find("PL", "淨利年增率", start=IS_END)
SCN_H = find("In", "鍵", start=100)
YC = "CDEFGHIJK"
def yrs(sh, r, c1="C", c2="K"): return f"@{sh}!{c1}{r}:{c2}{r}"
rec = V["rec"]
tornAll = ",".join(map(str, rel_rows))

cfg = {
 "company": "Palo Alto Networks", "ticker": "PANW", "thesis": "@README!A6",
 "question": "AI 進入資安後，代理身分與模組加購的增量，能否抵過價格侵蝕與推論成本，並撐起現價",
 "scenarios": [{"key": "A", "name": "情境 A（代理身分多、加購高、侵蝕低）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（代理身分少、加購低、侵蝕高）"}],
 "wide_text": True, "layout": "v1.3",
 "hero": {"rec": f"@Val!B{rec}", "tiles": [[f"@Val!A{V['tp']}", f"@Val!B{V['tp']}|n2", ["上檔空間 ", f"@Val!B{V['up']}|sp1"]], [f"@Val!A{V['cmp']}", f"@Val!B{V['cmp']}|n2", ["上檔空間 ", f"@Val!B{V['cmp']+4}|sp1"]], [f"@Val!A{V['px']}", f"@Val!B{V['px']}|n2", ""]],
          "note": ["@README!A5"]},
 "sections": [
  {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "note", "parts": ["@README!A6"]},
    {"type": "theme", "title": "營收（十億美元，依官方營收線堆疊）：FY2024A–FY2032E（情境 B）", "x": yrs("PL", 3), "fmt": "n0", "stack": True, "pad": 1.15, "total": yrs("PL", P["rev"]),
     "bars": [{"name": [f"@PL!A{P['prod']}"], "y": yrs("PL", P["prod"]), "color": "#6fa8dc"}, {"name": [f"@PL!A{P['ss']}"], "y": yrs("PL", P["ss"]), "color": "#0b7a80"}],
     "caption": "堆疊＝Product 與 Subscription and support（單一應報導分部下的兩條官方營收線）；百萬美元，圖軸以十億計；FY2027E 起為情境 B；含 CyberArk 併購貢獻。"},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [f"@README!A{r}" for r in range(12, 17)]},
    {"type": "h3", "text": "命題回覆（依模型計算）"},
    {"type": "list", "items": [f"@README!A{r}" for r in (9, 10, 11)]},
    {"type": "h3", "text": "本版要點"},
    {"type": "list", "items": ["@README!A8", "@README!A18"]}]},
  {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "目標價與現價（美元；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{V[k]}", "value": f"@Val!B{V[k]}", "series": sr} for k, sr in (("dcf", 1), ("cmp", 2), ("tp", 0))],
     "ref": {"label": f"@Val!A{V['px']}", "value": f"@Val!B{V['px']}"}},
    {"type": "kv", "items": [[f"@Val!A{V[k]}", f"@Val!B{V[k]}|p2"] for k in ("ke", "kd", "wacc", "ronic", "g")]},
    {"type": "table", "caption": "結論（三情境加權）", "head": ["項目", "值／情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{V['px']}-{V['wts']}", "cols": list("ABCD"), "fmts": ["", "n2", "n2", "n2"],
                "row_fmt": {**{str(V[k]): {"B": "sp1"} for k in ("up",)}, str(V["up"] + 1): {"B": "sp1"}, str(V["up"] + 2): {"B": "sp1"}, str(V["lvl"]): {"B": ""}, str(V["valr"]): {"B": ""}, str(V["gate"]): {"B": "n0"}, str(V["gate"] + 1): {c: "n0" for c in "BCD"}, str(V["gtxt"]): {"B": ""}, str(V["rec"]): {"B": ""}, str(V["cons"]): {"C": ""}, str(V["wts"]): {c: "p0" for c in "BCD"}}}},
    {"type": "table", "caption": "DCF 彙總（三情境；百萬美元，股數為百萬股，每股為美元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{V['pvx']}-{V['tvsh']}", "cols": list("ABCD")[:1] + list("BCD"), "fmts": ["", "n1", "n1", "n1"], "row_fmt": {str(V["tvsh"]): {c: "p1" for c in "BCD"}, str(V["dcfv"]): {c: "n2" for c in "BCD"}, str(V["dcfv"] - 1): {c: "n2" for c in "BCD"}, str(V["dcfv"] - 2): {c: "n0" for c in "BCD"}}}},
    {"type": "table", "caption": "Comparable（同業前瞻本益比中位數；倍數分母為 FY2028 Non-GAAP EPS）", "head": ["項目", "值／情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{V['peh']+1}-{V['pxpe']}", "cols": list("ABCD"), "fmts": ["", "n2", "n2", "n2"], "row_fmt": {str(V["rel"]): {"C": ""}, str(V["pe_rel"]): {"C": ""}, str(V["pemed"]): {"B": "n2"}}}},
    {"type": "table", "collapse": True, "summary": "展開：折現率組成與期間、評等邊界", "head": ["項目", "數值"],
     "rowgen": {"sheet": "Val", "rows": f"{V['rf']}-{V['ke']-1},{V['kd']+1}-{V['wacc']-1},{V['per_h']+1}-{V['att2']}", "cols": ["A", "B"], "fmts": ["", "n4"], "row_fmt": {str(V["att2"] - 1): {"B": "p1"}, str(V["att2"]): {"B": "p1"}, str(V["att2"] - 2): {"B": ""}, str(V["per_h"] + 4): {"B": ""}}}}]},
  {"id": "sens", "title": "敏感度", "blocks": [
    {"type": "tornado", "title": "加權目標價相對基準：輸入取低值（一次改一個，整本重算）", "sheet": "Sens", "rows": tornAll, "label_col": "A", "value_col": "B", "fmt": "sp1"},
    {"type": "tornado", "title": "加權目標價相對基準：輸入取高值（一次改一個，整本重算）", "sheet": "Sens", "rows": tornAll, "label_col": "A", "value_col": "C", "fmt": "sp1"},
    {"type": "note", "parts": ["@Sens!A2"]}, {"type": "note", "parts": ["@Sens!A9"]},
    {"type": "table", "collapse": True, "summary": "展開：全部案例（加權、DCF、Comparable 目標價）", "head": ["案例", "設定", "鍵", "低案加權", "高案加權", "低案 DCF", "高案 DCF", "低案 Comparable", "高案 Comparable"],
     "rowgen": {"sheet": "Sens", "rows": f"{case_first}-{case_last}", "cols": list("ABCDEFGHI"), "fmts": ["", "", "", "n2", "n2", "n2", "n2", "n2", "n2"]}},
    {"type": "kv", "items": [["build_id", "@Sens!D3"], ["模型狀態", "@Sens!D4"], ["@Sens!A5", "@Sens!D5|n2"], ["@Sens!A6", "@Sens!D6|n2"]]}]},
  {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "blocks": [
    {"type": "note", "parts": ["@RevDCF!A1"]}, {"type": "note", "parts": ["@RevDCF!A2"]},
    {"type": "table", "caption": "現價隱含值與模型基準（情境 B；DCF 單法）", "head": ["項目", "數值"],
     "rowgen": {"sheet": "RevDCF", "rows": f"{R['px']}-{R['interp']-1}", "cols": ["A", "B"], "fmts": ["", "n2"],
                "row_fmt": {**{str(r): {"B": f} for r, f in ((R["px"] + 2, "n0"), (R["px"] + 3, "n0"), (R["px"] + 4, "n0"), (R["px"] + 5, "n0"), (R["px"] + 6, "n0"), (R["px"] + 7, "p1"), (R["px"] + 8, "p1"), (R["px"] + 9, "p1"), (R["px"] + 10, "x1"), (R["px"] + 11, "x1"), (R["px"] + 12, "p1"), (R["px"] + 13, "p1"))}}}},
    {"type": "note", "parts": [f"@RevDCF!B{R['interp']}"]}]},
  {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "note", "parts": ["@PL!A1"]}, {"type": "note", "parts": ["@PL!A2"]},
    {"type": "theme", "title": "營收（十億美元）與營業利益率：FY2024A–FY2032E（情境 B）", "x": yrs("PL", 3), "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.3,
     "bars": [{"name": [f"@PL!A{IS0}"], "y": yrs("PL", IS0), "color": "#6fa8dc"}],
     "lines": [{"name": [f"@PL!A{IS0+7}"], "y": yrs("PL", IS0 + 7), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "長條＝營收（百萬美元，軸以十億計）；折線＝GAAP 營業利益率（右軸）；GAAP 口徑含股票薪酬與收購無形資產攤銷。FY2027E 起為情境 B。"},
    {"type": "table", "caption": "標準損益表（GAAP；百萬美元，EPS 為美元）：FY2024A–FY2026A 實績、FY2027E–FY2032E 情境 B", "head": ["項目"] + [f"@PL!{c}3" for c in YC],
     "rowgen": {"sheet": "PL", "rows": f"{IS0}-{IS_END},{NYOY}", "cols": ["A"] + list(YC), "fmts": [""] + ["n0"] * 9,
                "row_fmt": {**{str(IS0 + o): {c: "p1" for c in YC} for o in (1, 4, 7, 12)}, str(NYOY): {c: "p1" for c in YC}, str(IS_END): {c: "n2" for c in YC}}}},
    {"type": "note", "parts": [f"@PL!A{P['isnote']}"]},
    {"type": "theme", "title": "FCFF：情境 A／B／C（十億美元）", "x": "@PL!F3:K3", "fmt": "n0", "stack": False, "pad": 1.2,
     "bars": [{"name": ["情境 ", s, "：FCFF"], "y": f"@{sh}!F{P['fcff']}:K{P['fcff']}", "color": c} for s, sh, c in (("A", "PL_A", "#7fb8bb"), ("B", "PL", "#3b3838"), ("C", "PL_C", "#d9a08f"))],
     "caption": "三情境的無槓桿自由現金流（百萬美元，軸以十億計）；標準 FCFF 以 GAAP EBIT 計稅後利益，股票薪酬不加回。"},
    {"type": "table", "caption": "營收分部（情境 B；官方兩條營收線與其下的子線）", "head": ["項目"] + [f"@PL!{c}3" for c in YC],
     "rowgen": {"sheet": "PL", "rows": f"{P['prod']}-{P['chk1']+2}", "cols": ["A"] + list(YC), "fmts": [""] + ["n0"] * 9}}]},
  {"id": "drivers", "title": "驅動因子：受保護單位、代理身分與推論成本", "blocks": [
    {"type": "note", "parts": ["@README!A17"]},
    {"type": "theme", "title": "AI 淨效果的三個分項（情境 B；十億美元）", "x": "@PL!F3:K3", "fmt": "n1", "stack": False, "pad": 1.2,
     "bars": [{"name": [f"@PL!A{P[k]}"], "y": f"@PL!F{P[k]}:K{P[k]}", "color": c} for k, c in (("incr", "#0b7a80"), ("eros", "#c29a2b"), ("inf", "#b4472f"))],
     "lines": [{"name": [f"@PL!A{P['net']}"], "y": f"@PL!F{P['net']}:K{P['net']}", "color": "#3b3838", "labels": True}],
     "caption": "長條＝AI 營收增量、定價侵蝕、推論成本增量（皆相對 FY2026，百萬美元，軸以十億計）；折線＝AI 淨效果。推論成本取 Tokenomics 市場價。"},
    {"type": "table", "caption": "命題模組推導（情境 B；單位依列）", "head": ["項目"] + [f"@PL!{c}3" for c in "EFGHIJK"],
     "rowgen": {"sheet": "PL", "rows": f"{P['n']}-{P['net']+3}", "cols": ["A"] + list("EFGHIJK"), "fmts": [""] + ["n2"] * 7}},
    {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；In 頁情境驅動對照）", "head": [f"@In!{c}{SCN_H}" for c in "BDEF"],
     "rowgen": {"sheet": "In", "rows": f"{SCN_H+1}-{SCN_H+5}", "cols": ["B", "D", "E", "F"], "fmts": ["", "n2", "n2", "n2"], "row_fmt": {str(SCN_H + 5): {c: "p0" for c in "DEF"}, str(SCN_H + 2): {c: "p0" for c in "DEF"}, str(SCN_H + 3): {c: "p0" for c in "DEF"}}}},
    {"type": "table", "collapse": True, "summary": "展開：單一參數區間（基準進主模型，敏感度低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rowgen": {"sheet": "In", "rows": "8-38,51-103", "cols": ["B", "D", "E", "F", "G"], "fmts": ["", "n3", "n3", "n3", ""]}}]},
  {"id": "recon", "title": "對帳：公司指引與共識", "blocks": [
    {"type": "note", "parts": ["@Recon!A1"]},
    {"type": "table", "caption": "指引與共識對照（FY2027）", "head": ["項目", "本模型", "指引低", "指引高", "共識", "判定", "說明"], "rowgen": {"sheet": "Recon", "rows": f"{RC['h1']+1}-{RC['s2']-2}", "cols": list("ABCDEFG"), "fmts": ["", "n3", "n3", "n3", "n3", "", ""]}},
    {"type": "table", "collapse": True, "summary": "展開：逐分部、有機成長、同業可比性與 Tokenomics 讀值", "head": ["項目", "欄一", "欄二", "欄三", "欄四", "欄五", "欄六"], "rowgen": {"sheet": "Recon", "rows": f"{RC['s2']+2}-{RC['s6']-2}", "cols": list("ABCDEFG"), "fmts": ["", "n3", "n3", "n3", "n3", "", ""]}},
    {"type": "table", "caption": "風險登錄", "head": ["風險", "是否建模", "理由", "觀察指標"], "rowgen": {"sheet": "Recon", "rows": f"{RC['s6']+2}-{wb['Recon'].max_row}", "cols": list("ABCD"), "fmts": [""] * 4}}]},
  {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["build_id", "@Sens!D3"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"], ["Tokenomics 快照", "@Sources!B3"]]},
    {"type": "h3", "text": "市場資料日期與失效條件"},
    {"type": "list", "items": ["@README!A2", "@README!A29", "@README!A30"]},
    {"type": "h3", "text": "不適用與未做項目"},
    {"type": "list", "items": [f"@README!A{r}" for r in range(21, 29)]},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A", "F"], "In": ["B"], "Q": ["A"]}, "empty": "沒有標為未找到的項目。"}]}
 ]}
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("html.json 寫出")

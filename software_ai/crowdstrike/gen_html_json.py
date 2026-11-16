# 由成品 xlsx 與 build_crowdstrike.py 輸出的列號對照（layout.json）產生 html.json（列號隨模型變動，勿手改 html.json）。
# 用法：python3 gen_html_json.py <成品 xlsx> <layout.json> html.json
import json, sys, openpyxl
X, LJ, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
L = json.load(open(LJ, encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)
TB, PB, CB = L["TB"], L["PB"], L["CB"]
tb, pb, cb = TB["B"], PB["B"], CB["B"]
bid = L["reg"]["Sens_bid"]
YC = list("CDEFGHIJK")
sens = wb["Sens"]
case_rows = [r for r in range(4, bid - 1) if sens.cell(r, 6).value and isinstance(sens.cell(r, 5).value, (int, float))]
tornL, tornH = [], []
for r in case_rows:
    side, v = sens.cell(r, 3).value, sens.cell(r, 5).value
    if abs(v) >= 0.10:
        (tornL if side == "低" else tornH).append(r)
if len(tornL) < 2:   # 達兩位數百分比的低值案例過少時，補入絕對值最大者，使圖有內容
    extra = sorted([r for r in case_rows if sens.cell(r, 3).value == "低" and r not in tornL], key=lambda r: -abs(sens.cell(r, 5).value))
    tornL += extra[:3 - len(tornL)]
def yrs(sh, r, c1, c2): return f"@{sh}!{c1}{r}:{c2}{r}"
cfg = {
 "company": "CrowdStrike", "ticker": "CRWD", "thesis": "@README!A3",
 "question": "AI 進入資安後，AI 代理防護的增量能否抵過價格侵蝕與推論成本",
 "scenarios": [{"key": "A", "name": "情境 A（代理防護放量、價格侵蝕小）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（放量有限、價格侵蝕大）"}],
 "wide_text": True, "layout": "v1.3",
 "hero": {"rec": "@Val!B14", "tiles": [["@Val!A8", "@Val!B8|n2", ["上檔空間 ", "@Val!B11|sp1"]], ["@Val!A7", "@Val!B7|n2", ["上檔空間 ", "@Val!B12|sp1"]], ["@Val!A9", "@Val!B9|n2", ""]],
          "note": ["@README!A9"]},
 "sections": [
  {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "note", "parts": ["@README!A3"]},
    {"type": "theme", "title": "營收（十億美元，依營收線堆疊）：FY2024A–FY2032E（情境 B）", "x": yrs("PL", pb["yr"], "C", "K"), "fmt": "n0", "stack": True, "pad": 1.15, "total": yrs("PL", pb["rev"], "C", "K"),
     "bars": [{"name": [f"@PL!A{pb['sub']}"], "y": yrs("PL", pb["sub"], "C", "K"), "color": "#6fa8dc"}, {"name": [f"@PL!A{pb['ps']}"], "y": yrs("PL", pb["ps"], "C", "K"), "color": "#999999"}],
     "caption": "堆疊＝訂閱與專業服務（官方只有一個應報導分部）；單位為百萬美元，圖軸以十億計；FY2027E 起為情境 B。"},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [f"@README!A{r}" for r in (6, 7, 8)]},
    {"type": "table", "collapse": True, "summary": "展開：命題回覆", "head": ["命題回覆"], "rowgen": {"sheet": "README", "rows": "44-46", "cols": ["A"], "fmts": [""]}},
    {"type": "h3", "text": "本版要點"},
    {"type": "list", "items": [f"@README!A{r}" for r in (20, 21, 22, 23)]}]},
  {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "目標價與現價（美元；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{r}", "value": f"@Val!B{r}", "series": sr} for r, sr in ((6, 1), (7, 2), (8, 0))],
     "ref": {"label": "@Val!A9", "value": "@Val!B9"}},
    {"type": "note", "parts": ["@Val!A2"]}, {"type": "note", "parts": ["@Val!A3"]},
    {"type": "note", "parts": ["@README!A60"]},
    {"type": "kv", "items": [[f"@Val!A{r}", f"@Val!B{r}|p2"] for r in (33, 34, 39, 40)]},
    {"type": "table", "caption": "結論（情境加權）", "head": ["項目", "數值", "對照"], "rowgen": {"sheet": "Val", "rows": "6-28", "cols": ["A", "B", "C"], "fmts": ["", "n2", "n2"],
      "row_fmt": {"11": {"B": "sp1"}, "12": {"B": "sp1"}, "13": {"B": ""}, "14": {"B": ""}, "15": {"B": "n2", "C": "n2"}, "16": {"B": "n0"}, "17": {"B": "sp1"}, "18": {"B": "p1"}, "20": {"B": "sp1"}, "21": {"B": "sp1"}, "22": {"B": "sp1"},
                  "25": {"B": ""}, "26": {"B": "n0"}, "27": {"B": ""}, "28": {"B": ""}}}},
    {"type": "table", "caption": "DCF 彙總（三情境）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rowgen": {"sheet": "Val", "rows": "46-53", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"], "row_fmt": {"51": {c: "n2" for c in "BCD"}, "52": {c: "n2" for c in "BCD"}, "53": {c: "p1" for c in "BCD"}}}},
    {"type": "table", "caption": "Comparable（同業 EV／營收，TTM）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rowgen": {"sheet": "Val", "rows": "60-64", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"], "row_fmt": {"64": {c: "n2" for c in "BCD"}}}},
    {"type": "table", "collapse": True, "summary": "展開：折現率組成與資本結構", "head": ["項目", "數值"], "rowgen": {"sheet": "Val", "rows": "30-41", "cols": ["A", "B"], "fmts": ["", "n3"]}}]},
  {"id": "sens", "title": "敏感度", "blocks": [
    {"type": "tornado", "title": "加權目標價相對基準：輸入取低值（一次改一個，整本重算）", "sheet": "Sens", "rows": ",".join(map(str, tornL)), "label_col": "A", "value_col": "E", "fmt": "sp1"},
    {"type": "tornado", "title": "加權目標價相對基準：輸入取高值與聯動案例（一次改一個，整本重算）", "sheet": "Sens", "rows": ",".join(map(str, tornH)), "label_col": "A", "value_col": "E", "fmt": "sp1"},
    {"type": "note", "parts": ["@Sens!A2"]},
    {"type": "table", "collapse": True, "summary": "展開：全部案例（含 DCF 與 Comparable 單法）", "head": ["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "說明", "DCF 目標價", "Comparable 目標價"],
     "rowgen": {"sheet": "Sens", "rows": f"4-{max(case_rows)}", "cols": list("ABCDEFGHI"), "fmts": ["", "", "", "n2", "sp1", "", "", "n2", "n2"]}},
    {"type": "kv", "items": [["build_id", f"@Sens!D{bid}"], ["模型狀態", f"@Sens!D{bid+1}"]]}]},
  {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "blocks": [
    {"type": "note", "parts": ["@RevDCF!A1"]}, {"type": "note", "parts": ["@RevDCF!A2"]},
    {"type": "table", "caption": "現價隱含值與模型基準（情境 B）", "head": ["項目", "數值"], "rowgen": {"sheet": "RevDCF", "rows": "6-21", "cols": ["A", "B"], "fmts": ["", "n2"], "row_fmt": {"8": {"B": "n0"}, "9": {"B": "n0"}, "10": {"B": "n0"}, "20": {"B": "n0"}}}},
    {"type": "note", "parts": ["@RevDCF!A24"]}, {"type": "note", "parts": ["@RevDCF!A25"]},
    {"type": "table", "caption": "整本重算反解（其餘輸入不變，使加權目標價＝現價）", "head": ["輸入", "基準值", "反解值", "是否在敏感度範圍內", "說明"], "rowgen": {"sheet": "RevDCF", "rows": "29-30", "cols": list("ABCDE"), "fmts": ["", "n3", "n3", "", ""]}},
    {"type": "table", "caption": "DCF 與 Comparable 落差分解（每股，美元）", "head": ["項目", "數值"], "rowgen": {"sheet": "RevDCF", "rows": "33-37", "cols": ["A", "B"], "fmts": ["", "n2"]}}]},
  {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "note", "parts": ["@PL!A1"]}, {"type": "note", "parts": ["@PL!A2"]},
    {"type": "theme", "title": "營收（十億美元）與 GAAP 營業利益率：FY2024A–FY2032E（情境 B）", "x": yrs("PL", pb["yr"], "C", "K"), "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.3,
     "bars": [{"name": [f"@PL!A{pb['rev']}"], "y": yrs("PL", pb["rev"], "C", "K"), "color": "#6fa8dc"}],
     "lines": [{"name": [f"@PL!A{pb['ebitm']}"], "y": yrs("PL", pb["ebitm"], "C", "K"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "長條＝營收（百萬美元，軸以十億計）；折線＝GAAP 營業利益率（右軸）。FY2027E 起為情境 B。"},
    {"type": "table", "caption": "標準損益表（GAAP；百萬美元，EPS 為美元）：FY2024A–FY2026A 實績、FY2027E–FY2032E 情境 B", "head": ["項目"] + [f"@PL!{c}{pb['yr']}" for c in YC],
     "rowgen": {"sheet": "PL", "rows": f"{pb['sub']}-{pb['nig']}", "cols": ["A"] + YC, "fmts": [""] + ["n0"] * 9,
                "row_fmt": {**{str(pb[n]): {c: "p1" for c in YC} for n in ("g", "m", "gm", "ebitm", "nim")}, str(pb["m"]): {c: "n2" for c in YC}, str(pb["p"]): {c: "n2" for c in YC}, str(pb["eps"]): {c: "n2" for c in YC}, str(pb["chk"]): {c: "n3" for c in YC}, str(pb["chk2"]): {c: "n3" for c in YC}}}},
    {"type": "table", "collapse": True, "summary": "展開：Non-GAAP（公司自訂口徑）與增量推論成本", "head": ["項目"] + [f"@PL!{c}{pb['yr']}" for c in YC],
     "rowgen": {"sheet": "PL", "rows": f"{pb['sh']}-{pb['inf']}", "cols": ["A"] + YC, "fmts": [""] + ["n0"] * 9, "row_fmt": {str(pb["nggm"]): {c: "p1" for c in YC}, str(pb["ngopm"]): {c: "p1" for c in YC}, str(pb["ngeps"]): {c: "n2" for c in YC}}}},
    {"type": "theme", "title": "FCFF：情境 A／B／C（十億美元，FY2026A–FY2037E）", "x": yrs("CF", CB["B"]["yr"], "E", "P"), "fmt": "n0", "stack": False, "pad": 1.2,
     "bars": [{"name": ["情境 ", s, "：FCFF"], "y": yrs("CF", CB[s]["fcff"], "E", "P"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
     "caption": "三情境的無槓桿自由現金流（百萬美元，軸以十億計；股票薪酬已費用化、不加回）；情境差異來自 AI 代理防護放量、每模組價格與推論成本。"}]},
  {"id": "drivers", "title": "驅動因子：受保護單位、模組與 AI 淨效果", "blocks": [
    {"type": "note", "parts": ["@Thesis!A1"]}, {"type": "note", "parts": ["@Thesis!A2"]},
    {"type": "theme", "title": "AI 淨效果的分項（情境 B；十億美元）", "x": yrs("Thesis", tb["yr"], "F", "K"), "fmt": "n1", "stack": False, "pad": 1.2,
     "bars": [{"name": [f"@Thesis!A{tb[k]}"], "y": yrs("Thesis", tb[k], "F", "K"), "color": c} for k, c in (("inc", "#0b7a80"), ("ero", "#c29a2b"), ("cost", "#b4472f"))],
     "lines": [{"name": [f"@Thesis!A{tb['net']}"], "y": yrs("Thesis", tb["net"], "F", "K"), "color": "#3b3838", "labels": True}],
     "caption": "長條＝AI 代理防護增量營收、定價侵蝕（負值）、增量推論成本；折線＝AI 淨效果。百萬美元，軸以十億計。推論成本取 Tokenomics 市場價。"},
    {"type": "table", "caption": "命題推導（情境 B；ARR 與營收為百萬美元）", "head": ["項目"] + [f"@Thesis!{c}{tb['yr']}" for c in "EFGHIJK"],
     "rowgen": {"sheet": "Thesis", "rows": f"{tb['gu']}-{tb['chk']}", "cols": ["A"] + list("EFGHIJK"), "fmts": [""] + ["n1"] * 7,
                "row_fmt": {**{str(tb[k]): {c: "p1" for c in "EFGHIJK"} for k in ("gu", "dm", "r", "arrg", "ero", "share")}, str(tb["P"]): {c: "n2" for c in "EFGHIJK"}, str(tb["M"]): {c: "n2" for c in "EFGHIJK"}}}},
    {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；In 頁情境驅動對照）", "head": ["項目", "單位", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "In", "rows": "108-112", "cols": ["B", "C", "D", "E", "F"], "fmts": ["", "", "n2", "n2", "n2"], "row_fmt": {"108": {c: "p1" for c in "DEF"}, "112": {c: "p0" for c in "DEF"}}}}]},
  {"id": "recon", "title": "對帳：公司指引與共識", "blocks": [
    {"type": "note", "parts": ["@Recon!A1"]},
    {"type": "table", "caption": "指引與共識對照", "head": ["項目", "本模型", "指引或共識", "差距", "說明", "來源"], "rowgen": {"sheet": "Recon", "rows": "5-12", "cols": list("ABCDEF"), "fmts": ["", "n2", "n2", "p1", "", ""]}},
    {"type": "table", "collapse": True, "summary": "展開：指引差距分解與 AI 淨效果對照", "head": ["項目", "A", "B", "C"], "rowgen": {"sheet": "Recon", "rows": "15-25", "cols": list("ABCD"), "fmts": ["", "n2", "n2", "n2"]}},
    {"type": "table", "caption": "口徑與未找到", "head": ["項目", "說明"], "rowgen": {"sheet": "Recon", "rows": f"29-{wb['Recon'].max_row}", "cols": ["A", "C"], "fmts": ["", ""]}}]},
  {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["build_id", f"@Sens!D{bid}"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"], ["Tokenomics 快照", f"@Sources!B{L['reg']['Sources_tk']}"]]},
    {"type": "h3", "text": "合理性關卡診斷"},
    {"type": "list", "items": [f"@README!A{r}" for r in (56, 57, 58, 59, 60)]},
    {"type": "h3", "text": "市場資料日期與失效條件"},
    {"type": "list", "items": [f"@README!A{r}" for r in (26, 27, 32)]},
    {"type": "table", "caption": "風險登錄", "head": ["風險", "是否建模", "理由", "觀察指標"], "rowgen": {"sheet": "README", "rows": "37-41", "cols": list("ABCD")}},
    {"type": "table", "collapse": True, "summary": "展開：每頁的輸入、推導量與因果方向", "head": ["頁", "輸入", "推導量", "因果方向"], "rowgen": {"sheet": "README", "rows": "12-17", "cols": list("ABCD")}},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A", "C"], "In": ["B", "H"], "Q": ["A"]}, "empty": "沒有標為未找到的項目。"}]}
 ]}
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("html.json 寫出")

# 由 build_trendmicro.py 的列號對照（rowmap）產生 html.json（個股 HTML 版面規格 v1.3；數字與文字全取自 xlsx）。
# 用法：python3 gen_html_json.py <成品 xlsx> <rowmap.json> html.json
import json, sys, openpyxl
X, RM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
rm = json.load(open(RM, encoding="utf-8")); wb = openpyxl.load_workbook(X, data_only=True)
R, IX, VR, RVR, RCR, RDR, SEM, INR = rm["R"], rm["IX"], rm["VR"], rm["RVR"], rm["RCR"], rm["RDR"], rm["SE"], rm["INR"]
ph = rm["HD"]; GATE = rm["GATE"]
def rg(sh, r, c1="C", c2="L"): return f"@{sh}!{c1}{r}:{c2}{r}"
fmtu = lambda u: {"比率": "p1", "日圓": "n0", "百萬股": "n2", "人": "n0"}.get(u, "n1")
PLW = wb["PL"]
def rowx(r, cols="CDEFGHIJKL"): return [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmtu(PLW.cell(r, 2).value)}" for c in cols]
head = [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in "CDEFGHIJKL"]
th_chart = {"type": "theme", "title": "營收（十億日圓，依所在地別分部堆疊）與營業利益率：FY2023A–FY2032E（情境 B）", "x": rg("PL", ph), "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
            "total": rg("PL", R["B"]["rev"]),
            "bars": [{"name": [f"@PL!A{R['B']['r_' + k]}"], "y": rg("PL", R["B"]["r_" + k]), "color": c} for k, c in (("jp", "#6fa8dc"), ("am", "#c29a2b"), ("eu", "#0b7a80"), ("ap", "#9a8fc9"))],
            "lines": [{"name": [f"@PL!A{IX['opm']}"], "y": rg("PL", IX["opm"]), "axis": "r", "color": "#3b3838", "labels": True}],
            "caption": "堆疊＝四個官方報告分部（Japan、Americas、Europe、Asia Pacific）；折線＝合併營業利益率（右軸）。FY2026E 起為情境 B（基準）。"}
sens_ws = wb["Sens"]
def num(c, r): return isinstance(sens_ws[f"{c}{r}"].value, (int, float))
t1, t2 = SEM["T1"], SEM["T2"]
tors = []
for col, side in (("B", "低"), ("C", "高")):
    rows = [r for r in range(t1, t2 + 1) if num(col, r)]
    big = [r for r in rows if abs(sens_ws[f"{col}{r}"].value) >= 0.1] or rows
    small = [r for r in rows if r not in big and abs(sens_ws[f"{col}{r}"].value) >= 0.0005]
    tors.append({"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1", "rows": ",".join(map(str, big)),
                 "title": f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例（未達者摺疊於下）"})
    if small:
        tors.append({"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1", "rows": ",".join(map(str, small)), "collapse": True,
                     "summary": f"展開：加權目標價相對基準：輸入取{side}值，未達兩位數的案例", "title": f"加權目標價相對基準：輸入取{side}值（其餘案例）"})
case_tbl = {"type": "table", "collapse": True, "summary": "展開：全部案例（含 DCF 單法、評等與收支關卡）", "head": ["驅動", "設定", "方向", "加權目標價", "相對基準", "DCF 目標價", "DCF 變動（日圓）", "投資建議", "收支關卡 B"],
            "rowgen": {"sheet": "Sens", "rows": f"{SEM['HDR'] + 1}-{SEM['END']}", "cols": ["A", "B", "C", "D", "E", "F", "G", "H", "I"], "fmts": ["", "", "", "n0", "sp1", "n0", "n0", "", "n1"]}}
kreq_tbl = {"type": "table", "caption": "反解（整本重算）：支持現價與評等邊界所需值", "head": [f"@Sens!{c}{SEM['K0'] + 1}" for c in "ABCDE"],
            "rows": [[f"@Sens!A{r}"] + [f"@Sens!{c}{r}|n3" for c in "BCDE"] for r in SEM["KR"].values()]}
vrows = lambda a, b: f"{VR[a]}-{VR[b]}"
cfg = {
 "company": "趨勢科技 Trend Micro", "ticker": "4704", "thesis": f"@README!A{RDR['thesis']}",
 "question": "AI 進入資安軟體後，Vision One 的受保護範圍、模組數與單價成長，能否蓋過單價侵蝕與 AI 推論成本；現價隱含的利益率是否合理",
 "scenarios": [{"key": "A", "name": "情境 A（AI 增量大、侵蝕小）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（增量小、侵蝕大）"}],
 "wide_text": True, "layout": "v1.3",
 "hero": {"rec": f"@Val!B{VR['rec']}", "tiles": [[f"@Val!A{VR['tp']}", f"@Val!B{VR['tp']}|n0", ["上檔空間 ", f"@Val!B{VR['up']}|sp1"]],
                                               [f"@Val!A{VR['tpdcf']}", f"@Val!B{VR['tpdcf']}|n0", ["上檔空間 ", f"@Val!B{VR['upd']}|sp1"]],
                                               [f"@Val!A{VR['tpcomp']}", f"@Val!B{VR['tpcomp']}|n0", ["上檔空間 ", f"@Val!B{VR['upc']}|sp1"]],
                                               [f"@Val!A{VR['px']}", f"@Val!B{VR['px']}|n0", ""]],
          "note": [f"@README!A{RDR['hero']}"]},
 "sections": [
  {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "note", "parts": [f"@README!A{RDR['rec']}"]}, th_chart,
    {"type": "h3", "text": "命題"}, {"type": "list", "ordered": True, "items": [f"@README!A{RDR['p' + str(i)]}" for i in (1, 2, 3, 4)]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆", "parts": [f"@README!A{RDR['ans1']}", f"@README!A{RDR['ans2']}", f"@README!A{RDR['ans3']}"]},
    {"type": "h3", "text": "重點觀察"}, {"type": "list", "ordered": True, "items": [f"@README!A{RDR['o' + str(i)]}" for i in range(1, 6)]}]},
  {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（日圓；豎線＝現價）", "fmt": "n0",
     "items": [{"label": f"@Val!A{VR[k]}", "value": f"@Val!B{VR[k]}", "series": sr} for k, sr in (("tpdcf", 1), ("tpcomp", 2), ("tp", 0), ("cons", 1))],
     "ref": {"label": f"@Val!A{VR['px']}", "value": f"@Val!B{VR['px']}"}},
    {"type": "note", "parts": [f"@Val!A1", f"@Val!A2", f"@README!A{RDR['dist']}"]},
    {"type": "table", "caption": "目標價、上檔空間與評等", "head": ["項目", "數值"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['tpdcf']}-{VR['rec']}", "cols": ["A", "B"], "fmts": ["", "n0"],
                "row_fmt": {str(VR[k]): {"B": f} for k, f in (("up", "sp1"), ("upd", "sp1"), ("upc", "sp1"), ("vrate", ""), ("gate", "n1"), ("gatej", ""), ("rec", ""), ("wts", "p0"))}}},
    {"type": "kv", "items": [[f"@Val!A{VR[k]}", f"@Val!B{VR[k]}|{f}"] for k, f in (("rf", "p2"), ("ke", "p2"), ("wacc", "p2"), ("ronic", "p2"))]},
    {"type": "table", "caption": "DCF 彙總（三情境）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['pvx']}-{VR['comp']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n0", "n0", "n0"],
                "row_fmt": {str(VR["tvsh"]): {c: "p1" for c in "BCD"}, str(VR["pm"]): {"B": "n2", "C": "", "D": ""}}}},
    {"type": "table", "caption": "DCF 與 Comparable 落差分解（情境 B）與現行倍數", "head": ["項目", "數值"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['gap']}-{VR['rel']}", "cols": ["A", "B"], "fmts": ["", "n0"], "row_fmt": {str(VR["mdcf"]): {"B": "n2"}, str(VR["gap_c"]): {"B": "n3"}, str(VR["cur"]): {"B": "n2"}, str(VR["rel"]): {"B": "n2"}}}},
    {"type": "note", "parts": [f"@Val!B{VR['stale']}"]}]},
  {"id": "sens", "title": "敏感度", "blocks": tors + [{"type": "note", "parts": ["@Sens!A1", "@Sens!A2"]}, case_tbl, kreq_tbl,
    {"type": "kv", "items": [[f"@Sens!A{SEM['ST']}", f"@Sens!D{SEM['ST']}"], ["build_id", f"@Sens!D{SEM['BID']}"]]}]},
  {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "blocks": [
    {"type": "note", "parts": ["@RevDCF!A2", f"@RevDCF!B{RVR['sum']}"]},
    {"type": "table", "caption": "現價隱含值與模型", "head": ["項目", "數值"],
     "rows": [[f"@RevDCF!A{RVR[k]}", f"@RevDCF!B{RVR[k]}|{f}"] for k, f in (("evreq", "n0"), ("m1", "p1"), ("m0", "p1"), ("gapm", "n1"), ("mdc", "n2"))]}]},
  {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "note", "parts": ["@PL!A1"]},
    {"type": "theme", "title": "營收（十億日圓）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": rg("PL", ph), "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
     "bars": [{"name": [f"@PL!A{IX['rev']}"], "y": rg("PL", IX["rev"]), "color": "#6fc7cb"}],
     "lines": [{"name": [f"@PL!A{IX[k]}"], "y": rg("PL", IX[k]), "axis": "r", "color": c, "labels": k == "gm"} for k, c in (("gm", "#0b7a80"), ("opm", "#c29a2b"), ("nm", "#6f6bae"))],
     "caption": "長條＝營收；線＝毛利率、營業利益率、淨利率（右軸）。毛利率含 AI 推論成本，營業成本率其餘持平於 FY2025 實績。"},
    {"type": "table", "caption": "標準損益表（GAAP：日本會計準則，財報口徑，為公司主要使用的口徑；十億日圓，EPS 為日圓）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B", "head": head,
     "rows": [rowx(IX[k]) for k in ("rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "non", "pti", "tax", "ni", "nm", "eps")] + [{"sub": f"@PL!A{IX['ratio']}"}, rowx(IX["niy"])]},
    {"type": "theme", "title": "營業利益（十億日圓）× 情境 A／B／C", "x": rg("PL", ph), "fmt": "n0", "stack": False,
     "bars": [{"name": ["情境 ", s, "：營業利益"], "y": rg("PL", R[s]["oi"], "F", "L"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
     "caption": "三情境只在 Vision One 的 AI 增量、侵蝕、推論成本與營業費用年增上不同。"} if False else {"type": "note", "parts": ["@PL!A2"]},
    {"type": "table", "caption": "分部對帳（應為零）", "head": ["項目"] + [f"@PL!{c}{ph}" for c in "CDEFGHIJKL"],
     "rowgen": {"sheet": "PL", "rows": f"{rm['CK1']}-{rm['CK2']}", "cols": ["A", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"], "fmts": [""] + ["n3"] * 10}},
    {"type": "table", "caption": "收支關卡", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "CF", "rows": f"{GATE}-{GATE + 4}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"], "row_fmt": {str(GATE + 4): {c: "" for c in "BCD"}}}}]},
  {"id": "drivers", "title": "驅動因子：分部營收推導", "blocks": [
    {"type": "note", "parts": [f"@README!A{RDR['drv']}"]},
    {"type": "theme", "title": "Vision One 營收（十億日圓）× 情境 A／B／C", "x": rg("PL", ph, "F", "L"), "fmt": "n0", "stack": False,
     "bars": [{"name": ["情境 ", s, "：Vision One"], "y": rg("PL", R[s]["v1"], "F", "L"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
     "caption": "Vision One 營收＝平均 ARR × 匯率；三情境在 AI 增量倍數與單價侵蝕上不同。"},
    {"type": "table", "caption": "分部營收推導（情境 B；十億日圓）：各分部＝Vision One（命題子列）＋其他，加總＝合併營收", "head": head,
     "rows": [rowx(R["B"][k]) for k in ("r_jp", "v_jp", "o_jp", "r_am", "v_am", "o_am", "r_eu", "v_eu", "o_eu", "r_ap", "v_ap", "o_ap", "rev")]},
    {"type": "table", "caption": "AI 淨效果（情境 B；十億日圓）：AI 營收增量 − 單價侵蝕 − 推論成本", "head": head,
     "rows": [rowx(R["B"][k]) for k in ("a2", "v1", "inc", "ero", "inf", "ain")]},
    {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價）", "head": [f"@In!{c}{rm['SCN_H']}" for c in "BDEF"],
     "rowgen": {"sheet": "In", "rows": f"{rm['SCN_H'] + 1}-{rm['SCN_H'] + 7}", "cols": ["B", "D", "E", "F"], "fmts": ["", "n3", "n3", "n3"]}},
    {"type": "table", "caption": "單一參數區間（基準進主模型；敏感度低／高只用於敏感度頁、不進目標價）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[f"@In!B{r}", f"@In!D{r}|g", f"@In!E{r}|g", f"@In!F{r}|g", f"@In!G{r}|tag"] for k, r in INR.items()
              if wb["In"].cell(r, 4).value is not None and not k.endswith(("_A", "_B", "_C"))]}]},
  {"id": "recon", "title": "對帳：指引、共識、分部與單位經濟", "blocks": [
    {"type": "table", "caption": "合併層：模型 vs 公司指引與共識（財報前）", "head": ["項目", "模型情境 B", "對照值", "差", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{RCR['r26']}-{RCR['c27']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""],
                "row_fmt": {str(RCR["m28"]): {"B": "p1", "C": "p1", "D": "n1"}, str(RCR["gm"]): {"B": "p1", "C": "p1", "D": "n1"}, str(RCR["c26"]): {"D": "sp1"}, str(RCR["c27"]): {"D": "sp1"}}}},
    {"type": "table", "caption": "分部：上半年實績 vs 模型", "head": ["分部", "上半年實績", "上半年年增", "模型 FY2026E", "隱含下半年", "隱含下半年年增", "FY2025 實績"],
     "rowgen": {"sheet": "Recon", "rows": f"{RCR['h1jp']}-{RCR['h1t']}", "cols": list("ABCDEFG"), "fmts": ["", "n1", "p1", "n1", "n1", "p1", "n1"]}},
    {"type": "table", "caption": "原型 P 單位經濟：每百萬美元平均 ARR（千美元）", "head": ["項目", "情境 A", "情境 B", "情境 C", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{RCR['uinc2027']}-{RCR['uain']}", "cols": list("ABCDE"), "fmts": ["", "n0", "n0", "n0", ""],
                "row_fmt": {str(RCR["uinfr"]): {c: "p1" for c in "BCD"}, str(RCR["uain"]): {c: "n1" for c in "BCD"}}}}]},
  {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RDR['ver']}"]},
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期（資料截止日、成品日、股價基準日）", f"@README!A{RDR['a1']}"], ["build_id", f"@Sens!D{SEM['BID']}"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"], ["Tokenomics 快照", "@Sources!B2"]]},
    {"type": "h3", "text": "資料與口徑"}, {"type": "list", "items": ["@In!A2", "@Sources!B3", "@Sources!B4"]},
    {"type": "h3", "text": "市場資料日期、失效條件與承重假設"},
    {"type": "list", "items": [f"@README!A{RDR[k]}" for k in ("md1", "diag", "f1", "f2", "f3", "l1", "l2", "l3", "na1", "na2", "na3", "na4")]},
    {"type": "table", "caption": "風險登錄", "head": ["風險", "是否建模", "理由", "觀察指標"], "rowgen": {"sheet": "README", "rows": f"{rm['RISK_HDR'] + 1}-{rm['RISK_HDR'] + 5}", "cols": list("ABCD")}},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "list", "items": [f"@README!A{RDR['nf']}"]}]}]}
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

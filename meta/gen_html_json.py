# 由 build_meta.py 的列號對照（ROWMAP）產生 meta/html.json（列號隨模型變動，勿手改 html.json）。
# 用法：python3 meta/gen_html_json.py <成品 xlsx> <rowmap.json> meta/html.json
import json, sys, openpyxl
X, RM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
rm = json.load(open(RM, encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)
T, T0, PR, CR, VR = rm["T"], rm["T0"], rm["PR"], rm["CR"], rm["VR"]
def find(sh, text, col=1, start=1):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        if str(ws.cell(r, col).value or "").startswith(text): return r
    raise KeyError((sh, text))
PB = PR["B"]
hA = PR["A"]["ads"] - 1
def rng(sh, r, c1="G", c2="M"): return f"@{sh}!{c1}{r}:{c2}{r}"
def met(name, key, fmt, chart=False, sheet="PL", d=PR, ymin=None):
    m = {"name": name, "fmt": fmt, "rows": {s: rng(sheet, d[s][key]) for s in "ABC"}}
    if chart: m["chart"] = True
    if ymin: m["ymin"] = ymin
    return m
conc0 = VR["tpdcf"]; conc1 = VR["chk_rd"]
rowfmt_conc = {}
for k in ("up", "upc", "scenup", "tvsh2"): rowfmt_conc[str(VR[k])] = {"B": "sp1", "C": "sp1", "D": "sp1"}
rowfmt_conc[str(VR["per"])] = {"B": "n3", "C": "n3"}; rowfmt_conc[str(VR["bdist"])] = {"B": "n1", "C": "n2", "D": "n2"}; rowfmt_conc[str(VR["uattr"])] = {"B": "sp1", "C": "sp1", "D": "sp1"}
for k in ("dep_dcf", "dep_sotp"): rowfmt_conc[str(VR[k])] = {"B": "n2", "C": "p1", "D": "n1"}
rowfmt_conc[str(VR["dep_tot"])] = {"B": "n2", "C": "p1"}; rowfmt_conc[str(VR["banner"])] = {"B": ""}
for k in ("liq55", "liq55c"): rowfmt_conc[str(VR[k])] = {"B": "n1", "C": "n1", "D": "n1"}
rowfmt_conc[str(VR["liq55l"])] = {c: "n2" for c in "BCD"}; rowfmt_conc[str(VR["liq55i"])] = {c: "n1" for c in "BCD"}
rowfmt_conc[str(VR["tier"])] = {"B": ""}; rowfmt_conc[str(VR["rec"])] = {"B": ""}; rowfmt_conc[str(VR["meth"])] = {"B": ""}
for k in ("tp_r2", "tp_r1"): rowfmt_conc[str(VR[k])] = {"B": "n2", "C": "sp1", "D": ""}
for k in ("tp_dcf12", "tp_sotp12"): rowfmt_conc[str(VR[k])] = {"B": "n2", "C": "n2"}
rowfmt_conc[str(VR["wsum"])] = {"B": "n2"}; rowfmt_conc[str(VR["chk_rd"])] = {"B": "n3"}
dsum0, dsum1 = VR["pvx"], VR["dcf12"]
sotp0 = VR["s_foa"]; sotp1 = VR["simp"]
rf_s = {str(VR["tvsh"]): {"B": "p1", "C": "p1", "D": "p1"}, str(VR["shv"]): {"B": "n3", "C": "n3", "D": "n3"}, str(VR["ppref"]): {"B": "n2", "C": "n2", "D": "n2"}}
rf_t = {str(VR["ssh"]): {"B": "n3", "C": "n3", "D": "n3"}, str(VR["simp"]): {"B": "x1", "C": "x1", "D": "x1"}}
gate = rm["GATE"]
r1 = find("Recon", "1. 合併層"); r2 = find("Recon", "2. 逐分部"); r3 = find("Recon", "3. M4"); r4 = find("Recon", "4. 一次性"); r5 = find("Recon", "5. v0.1.1")
r5b = find("Recon", "5b. 指引差距"); r6 = find("Recon", "6. 自用 AI 讀法 2"); r6b = find("Recon", "6b. 反解"); r6c = find("Recon", "6c. 隱含"); r6d = find("Recon", "6d. 可觀察")
r7 = find("Recon", "7. 自用在役容量"); r8 = find("Recon", "8. 同業可比性")
RC6, RC7, RC8 = rm["RC6"], rm["RC7"], rm["RC8"]
rend = wb["Recon"].max_row
RC6_q = find("Recon", "反解：達到指引中點")
se_end = max(x[3] for x in rm["SE_ROWS"])
TOR_ROWS = ",".join(str(x[3]) for x in rm["SE_ROWS"] if abs(wb["Sens"].cell(x[3], 5).value or 0) > 1e-4)
TB = T["B"]
thh = TB["adds"] - 2
th_rows = ",".join(str(TB[k]) for k in ("adds", "end", "avg_inc", "cx_ai", "cx_tot", "da_ai", "int_opex", "int_cost", "da_b", "shadow", "sh_vs"))
th_fmt = {str(TB[k]): {c: "n0" for c in "EFGHIJKLM"} for k in ("adds", "end", "avg_inc")} | {str(TB["sh_vs"]): {c: "n2" for c in "EFGHIJKLM"}}
seg_rows = ",".join(str(PB[k]) for k in ("ads", "other", "foa", "rl", "rev", "e_foa", "oi_foa", "oi_rl", "oi", "om", "da", "da_ai", "shadow", "oiv", "ni", "eps", "epsx", "epsv"))
seg_fmt = {str(PB["om"]): {c: "p1" for c in "CDEFGHIJKLM"}} | {str(PB[k]): {c: "n2" for c in "CDEFGHIJKLM"} for k in ("eps", "epsx", "epsv")}
cf_rows = ",".join(str(CR["B"][k]) for k in ("cfo", "ag", "cx_ai", "cx_na", "cx", "fcf", "div", "pref", "fcfd", "eq", "bb", "plug", "cend", "dend", "nc"))
INK = {str(wb["In"].cell(r, 1).value): r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}
SCN_H = find("In", "J. 情境驅動對照") + 1
SCN = {k: INK[k] for k in ("scn_p", "scn_cx27", "scn_w")}
t0_rows = f"{T0['pmw_h']}-{T0['opx0']}"
t0_fmt = {str(T0[k]): {"C": "n2"} for k in ("pmw_h", "pmw_b2", "pmw_b3", "pmw_26", "pmw_27", "pmw_28", "aicx_24", "aicx_25", "aicx_9m", "dpi0", "he0", "opx0")} | {str(T0["ai_sh26"]): {"C": "p1"}}
cfg = {
 "company": "Meta Platforms",
 "ticker": "META",
 "thesis": "@README!A3",
 "question": "FY2028 AI 增量營業利益（主值讀法：自用 AI 影子效益＝自用增量在役 MW × Tokenomics 每 MW 持有成本 × k）能否覆蓋 AI 資本支出的資金成本；收支關卡＝FY2026 到 FY2030 自由現金流扣股利合計",
 "scenarios": [{"key": "A", "name": "情境 A（AI 提高每則廣告價格）"}, {"key": "B", "name": "情境 B（近四季成長收斂；基準）"}, {"key": "C", "name": "情境 C（AI 價格提升消退）"}],
 "sections": [
  {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "kv", "items": [["投資建議", "@README!A5"], ["重點觀察一", "@README!A6"], ["重點觀察二", "@README!A7"], ["重點觀察三", "@README!A8"]]},
    {"type": "h3", "text": "本版要點"},
    {"type": "list", "items": [f"@README!A{r}" for r in range(find("README", "本版要點") + 1, find("README", "本版要點") + 6)]}]},
  {"id": "drivers", "title": "驅動因子 → 推導量", "intro": "@Thesis!A1", "blocks": [
    {"type": "table", "caption": "驅動 → 推導對照表", "head": ["頁", "輸入", "推導量", "因果方向"],
     "rowgen": {"sheet": "README", "rows": f"{find('README', '頁') + 1}-{find('README', 'Sens', 1, find('README', '頁'))}", "cols": ["A", "B", "C", "D"]}},
    {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；In 頁 J 節）", "head": [f"@In!{c}{SCN_H}" for c in "BDEFCG"],
     "rows": [[f"@In!B{SCN[k]}", f"@In!D{SCN[k]}|{f}", f"@In!E{SCN[k]}|{f}", f"@In!F{SCN[k]}|{f}", f"@In!C{SCN[k]}", f"@In!G{SCN[k]}|tag"] for k, f in (("scn_p", "p0"), ("scn_cx27", "n1"), ("scn_w", "p0"))]},
    {"type": "table", "caption": "單一參數區間（In 頁；基準進主模型，敏感度低／高只用於敏感度頁、不進目標價；三情境共用）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!C3", "@In!G3"],
     "rows": [[f"@In!B{INK[k]}", f"@In!D{INK[k]}|{f}", f"@In!E{INK[k]}|{f}", f"@In!F{INK[k]}|{f}", f"@In!C{INK[k]}", f"@In!G{INK[k]}|tag"]
              for k, f in (("capex_g", "n0"), ("cx27_k", "n2"), ("a28", "n2"), ("cx_na", "p0"), ("cip_ai", "p0"), ("selfb", "p0"), ("life_it", "n0"), ("rent_l", "n1"),
                           ("gl_impr", "p0"), ("gl_price", "p0"), ("gl_other", "p0"), ("rl_loss32", "n1"), ("cip_q", "n0"), ("internal_ai_reading", "n0"), ("beta", "n3"), ("g_tv", "p0"), ("w_dcf", "p0"), ("contg", "n1"))]},
    {"type": "table", "caption": "命題模組：期初機隊（資本支出換算 MW）", "head": ["項目", "單位", "數值"],
     "rowgen": {"sheet": "Thesis", "rows": t0_rows, "cols": ["A", "B", "C"], "fmts": ["", "", "n0"], "row_fmt": t0_fmt}},
    {"type": "table", "caption": "命題模組：自用 AI 容量、資本支出、折舊與影子效益（情境 B）", "head": ["項目", "單位"] + [f"@Thesis!{c}{thh}" for c in "EFGHIJKLM"],
     "rowgen": {"sheet": "Thesis", "rows": th_rows, "cols": ["A", "B", "E", "F", "G", "H", "I", "J", "K", "L", "M"], "fmts": ["", "", "n1", "n1", "n1", "n1", "n1", "n1", "n1", "n1", "n1"], "row_fmt": th_fmt}},
    {"type": "note", "parts": ["@Thesis!A2"]},
    {"type": "note", "parts": ["@Seg!A2"]}]},
  {"id": "pl", "title": "三情境損益", "intro": "@PL!A2", "blocks": [
    {"type": "pl", "years": f"@PL!G{hA}:M{hA}", "metrics": [
       met("合併營收（十億美元）", "rev", "n0", True), met("Advertising（十億美元）", "ads", "n0", True),
       met("合併營業利益率（財報口徑）", "om", "p1", True, ymin="auto"), met("Family of Apps 營業利益率", "om_foa", "p1"),
       met("EPS（GAAP 稀釋，美元）", "eps", "n2", True, ymin="auto"), met("EPS（評價口徑，美元）", "epsv", "n2"),
       {"name": "自由現金流（十億美元）", "fmt": "n0", "rows": {s: rng("CF", CR[s]["fcf"]) for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 的分部與合併損益（含實績年度）", "head": ["項目"] + [f"@PL!{c}{hA}" for c in "CDEFGHIJKLM"],
     "rowgen": {"sheet": "PL", "rows": seg_rows, "cols": ["A", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"], "fmts": [""] + ["n1"] * 11, "row_fmt": seg_fmt}},
    {"type": "table", "caption": "分部對帳（三情境 × 全部年度取絕對值最大者，應為零）", "head": ["項目"] + [f"@PL!{c}5" for c in "CDEFGHIJKLM"],
     "rowgen": {"sheet": "PL", "rows": f"{rm['SUMROW']}-{rm['SUMROW'] + 2}", "cols": ["A", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"], "fmts": [""] + ["n3"] * 11}},
    {"type": "table", "caption": "收支關卡", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "CF", "rows": f"{gate}-{gate + 3},{gate + 5}-{gate + 13}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"],
                "row_fmt": {**{str(gate + i): {"B": "", "C": "", "D": ""} for i in (3, 6, 8, 10, 12)}, str(gate + 9): {c: "n2" for c in "BCD"}, str(gate + 13): {c: "n0" for c in "BCD"}}}},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 的現金流、股東回饋與融資", "head": ["項目"] + [f"@CF!{c}{CR['B']['ni'] - 1}" for c in "EFGHIJKLM"],
     "rowgen": {"sheet": "CF", "rows": cf_rows, "cols": ["A", "E", "F", "G", "H", "I", "J", "K", "L", "M"], "fmts": [""] + ["n1"] * 9}}]},
  {"id": "valuation", "title": "估值", "intro": "@Val!A1", "blocks": [
    {"type": "table", "caption": "@Val!A4", "head": ["項目", "數值", "對照", "對照", "說明"],
     "rowgen": {"sheet": "Val", "rows": f"{conc0}-{conc1}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""], "row_fmt": rowfmt_conc}},
    {"type": "bars", "title": "目標價與現價", "fmt": "n0", "items": [
       {"label": "@Val!A" + str(VR["tpdcf"]), "value": "@Val!B" + str(VR["tpdcf"]), "series": 0},
       {"label": "@Val!A" + str(VR["tpcomp"]), "value": "@Val!B" + str(VR["tpcomp"]), "series": 1},
       {"label": "@Val!A" + str(VR["tp"]), "value": "@Val!B" + str(VR["tp"]), "series": 2},
       {"label": "@Val!A" + str(VR["tp_r1"]), "value": "@Val!B" + str(VR["tp_r1"]), "series": 1},
       {"label": "@Val!A" + str(VR["cons"]), "value": "@Val!B" + str(VR["cons"]), "series": 0}],
     "ref": {"label": "@Val!A" + str(VR["px"]), "value": "@Val!B" + str(VR["px"])}},
    {"type": "kv", "items": [["@Val!A" + str(VR[k]), "@Val!B" + str(VR[k]) + "|" + f] for k, f in (("emkt", "n0"), ("dmkt", "n1"), ("ke", "p2"), ("kdat", "p2"), ("wacc", "p2"), ("ronic", "p2"))]},
    {"type": "table", "caption": "DCF 彙總（三情境）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{dsum0}-{dsum1}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"], "row_fmt": rf_s}},
    {"type": "table", "caption": "分部加總（目標時點 NTM EBITDA × 同業倍數）", "head": ["分部", "情境 A", "情境 B", "情境 C", "倍數", "EV：A", "EV：B", "EV：C"],
     "rowgen": {"sheet": "Val", "rows": f"{sotp0}-{sotp1}", "cols": ["A", "B", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n1", "n1", "n1", "", "n0", "n0", "n0"], "row_fmt": rf_t}},
    {"type": "table", "collapse": True, "summary": "展開：同業倍數", "head": ["同業組", "倍數", "家數"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['mult_ads']}-{VR['mult_x']}", "cols": ["A", "B", "C"], "fmts": ["", "x1", "n0"]}}]},
  {"id": "recon", "title": "對帳：共識、公司指引與未入基準的算力", "intro": "@Recon!A1", "blocks": [
    {"type": "table", "caption": "@Recon!A" + str(r1), "head": ["項目", "年度", "情境 A", "情境 B", "情境 C", "共識", "B 對共識", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r1 + 2}-{r2 - 2}", "cols": ["A", "B", "C", "D", "E", "F", "G", "H"], "fmts": ["", "", "n1", "n1", "n1", "n1", "sp1", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r2), "head": ["項目", "本模型", "指引低", "指引高", "共識", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r2 + 2}-{r3 - 2}", "cols": ["A", "B", "C", "D", "E", "F"], "fmts": ["", "n1", "n1", "n1", "n1", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r3), "head": ["項目", "數值", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r3 + 2}-{r4 - 2}", "cols": ["A", "B", "C"], "fmts": ["", "n1", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r4), "head": ["項目", "財報口徑", "排除一次性", "差額", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r4 + 2}-{r5 - 2}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""]}},
    {"type": "table", "collapse": True, "summary": "展開：季報新事實與法說方向（只作對照）", "head": ["項目", "數值", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r5 + 2}-{r5b - 2}", "cols": ["A", "B", "C"], "fmts": ["", "n2", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r5b), "head": ["項目", "數值", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r5b + 2}-{r6 - 2}", "cols": ["A", "B", "C"], "fmts": ["", "n2", ""], "row_fmt": {str(RC6_q): {"B": "p1"}}}},
    {"type": "table", "caption": "@Recon!A" + str(r6 + 1), "head": ["@Recon!A" + str(r6 + 2), "@Recon!B" + str(r6 + 2), "@Recon!C" + str(r6 + 2), "@Recon!D" + str(r6 + 2), "@Recon!E" + str(r6 + 2), "@Recon!F" + str(r6 + 2), "@Recon!G" + str(r6 + 2)],
     "rowgen": {"sheet": "Recon", "rows": f"{r6 + 3}-{r6b - 1}", "cols": ["A", "B", "C", "D", "E", "F", "G"], "fmts": ["n2", "n2", "sp1", "", "n2", "n2", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r6b), "head": ["@Recon!A" + str(r6b + 1), "@Recon!B" + str(r6b + 1), "@Recon!C" + str(r6b + 1), "@Recon!D" + str(r6b + 1)],
     "rowgen": {"sheet": "Recon", "rows": f"{r6b + 2}-{RC6['kchk']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n3", "n2", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r6c), "head": ["@Recon!A" + str(r6c + 1), "@Recon!B" + str(r6c + 1), "@Recon!C" + str(r6c + 1), "@Recon!D" + str(r6c + 1), "@Recon!E" + str(r6c + 1), "@Recon!F" + str(r6c + 1), "@Recon!G" + str(r6c + 1)],
     "rowgen": {"sheet": "Recon", "rows": f"{r6c + 2}-{r6d - 1}", "cols": ["A", "B", "C", "D", "E", "F", "G"], "fmts": ["", "", "n1", "n1", "n1", "n1", ""],
                "row_fmt": {str(RC6["att_" + s_][k]): {c: "p0" for c in "CDEF"} for s_ in "ABC" for k in ("ratio", "ratio2")}}},
    {"type": "table", "caption": "@Recon!A" + str(r6d), "head": ["@Recon!A" + str(r6d + 1), "@Recon!B" + str(r6d + 1), "@Recon!C" + str(r6d + 1), "@Recon!D" + str(r6d + 1), "@Recon!E" + str(r6d + 1), "@Recon!F" + str(r6d + 1), "@Recon!G" + str(r6d + 1)],
     "rowgen": {"sheet": "Recon", "rows": f"{r6d + 2}-{r7 - 2}", "cols": ["A", "B", "C", "D", "E", "F", "G"], "fmts": ["", "p0", "p0", "p0", "p0", "", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r7), "head": ["@Recon!A" + str(r7 + 1), "@Recon!B" + str(r7 + 1), "@Recon!C" + str(r7 + 1), "@Recon!D" + str(r7 + 1), "@Recon!E" + str(r7 + 1), "@Recon!F" + str(r7 + 1), "@Recon!G" + str(r7 + 1), "@Recon!H" + str(r7 + 1), "@Recon!I" + str(r7 + 1)],
     "rowgen": {"sheet": "Recon", "rows": f"{r7 + 2}-{RC7['n']}", "cols": ["A", "B", "C", "D", "E", "F", "G", "H", "I"], "fmts": ["", "n2", "n2", "n2", "n2", "n2", "n2", "n2", ""],
                "row_fmt": {str(RC7["gap"]): {c: "sp1" for c in "BCDEFGH"}, str(RC7["flag"]): {c: "" for c in "BCDEFGH"}, str(RC7["n"]): {"B": "n0"}}}},
    {"type": "table", "collapse": True, "summary": "展開：同業可比性與商業模式子群中位數（只作展示）", "head": ["@Recon!A" + str(r8 + 1), "@Recon!B" + str(r8 + 1), "@Recon!C" + str(r8 + 1), "@Recon!D" + str(r8 + 1), "@Recon!E" + str(r8 + 1), "@Recon!F" + str(r8 + 1), "@Recon!G" + str(r8 + 1)],
     "rowgen": {"sheet": "Recon", "rows": f"{r8 + 2}-{rend}", "cols": ["A", "B", "C", "D", "E", "F", "G"], "fmts": ["", "", "x2", "", "", "", ""]}}]},
  {"id": "sens", "title": "敏感度", "intro": "@Sens!A1", "blocks": [
    {"type": "tornado", "title": "加權目標價相對基準的變動（一次改一個輸入，整本重算；|變動| ≤ 0.01% 的案例只列在下方）", "sheet": "Sens", "rows": TOR_ROWS, "label_col": "A", "value_col": "E", "fmt": "sp1"},
    {"type": "table", "collapse": True, "summary": "展開：全部案例（含 DCF 單法、評等與收支關卡 B）", "head": ["驅動", "設定", "方向", "加權目標價", "相對基準", "DCF 目標價", "DCF 變動（美元）", "投資建議", "收支關卡 B"],
     "rowgen": {"sheet": "Sens", "rows": f"4-{se_end}", "cols": ["A", "B", "C", "D", "E", "J", "K", "L", "H"], "fmts": ["", "", "", "n2", "sp1", "n2", "n2", "", "n1"]}},
    {"type": "kv", "items": [["@Sens!A" + str(rm["SE_BASE"]), "@Sens!D" + str(rm["SE_BASE"]) + "|n2"], ["@Sens!A" + str(rm["SE_STALE"]), "@Sens!D" + str(rm["SE_STALE"]) + "|n2"],
                             ["模型狀態", "@Sens!D" + str(rm["SE_STATUS"])], ["build_id", "@Sens!D" + str(rm["SE_BID"])]]}]},
  {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期（資料截止日、成品日、股價基準日）", "@README!A1"], ["build_id", "@Sens!D" + str(rm["SE_BID"])], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"],
                             ["Tokenomics 快照", "@Sources!B" + str(find("Sources", "Tokenomics 快照", 2))]]},
    {"type": "h3", "text": "資料與口徑"},
    {"type": "list", "items": ["@Q!A1", "@Q!A2", "@In!A2", "@Sens!A2"]},
    {"type": "h3", "text": "市場資料日期、失效條件與風險登錄"},
    {"type": "list", "items": [f"@README!A{r}" for r in list(range(find("README", "市場資料日期") + 1, find("README", "主要失效條件"))) + list(range(find("README", "主要失效條件") + 1, find("README", "風險登錄"))) if wb["README"].cell(r, 1).value]},
    {"type": "table", "caption": "風險登錄", "head": ["風險", "是否建模", "理由", "觀察指標"],
     "rowgen": {"sheet": "README", "rows": f"{find('README', '風險登錄') + 2}-{find('README', '工作表：') - 2}", "cols": ["A", "B", "C", "D"]}},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A", "C", "F"], "In": ["B", "H", "I"]}, "empty": "沒有標為未找到的項目。"}]}
 ]
}
# v0.1.5：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）
import os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "company_html"))
from relayout import relayout
SPEC = {
 "in_skip": ["H", "J"],
 "in_groups": [
  ["廣告本業（Family of Apps）", [
    ["廣告曝光與價格", ["conv_n", "g0_q", "gl_impr", "gl_price", "pA", "pC", "inc_m", "mb_q3"]],
    ["其他營收", ["gl_other"]],
    ["一次性項目（評價口徑排除）", ["ot_legal_q2", "ot_sev_q2", "ot_legal_q3", "ot_tax_q1", "eps_x_q3", "ot_legal_at"]],
    ["估值倍數（數位廣告同業 NTM EV/EBITDA）", ["pm_ads_googl", "pm_ads_pins", "pm_ads_ttd", "pm_ads_app", "pm_ads_rddt", "pm_ads_snap", "pm_meta_own", "ads_mode"]]]],
  ["Reality Labs", [
    ["營收與虧損", ["gl_rl", "rl_loss26", "rl_loss27", "rl_loss32", "rl_da_sh", "rl_mode"]]]],
  ["自用 AI 算力", [
    ["資本支出與 AI 占比", ["cx_na", "cx_23", "ai_23", "cx_24", "cx_25", "capex_9m", "capex_g", "cx27_k", "cx27_up", "cx27_B", "cx27_A", "cx27_C", "a28", "conv_mw", "bd_delay"]],
    ["機隊組成與機房", ["selfb", "am_26", "am_27", "am_28"]],
    ["自用在役容量（Wattonomics 上線時程）", ["mw_src", "wif_2026", "wif_2027", "wif_2028", "wif_2029", "wif_2030", "wif_2031", "wif_2032"]],
    ["在建工程與折舊", ["cip_0", "cip_q", "cipd_q4", "dep_lag", "cip_ai", "cip_it", "life_it", "ec_press", "life_fac", "life_na", "phi_0", "phi_q4", "da_base"]],
    ["租用算力（只進敏感度）", ["rent_l"]],
    ["每 MW 成本（Tokenomics 快照）", ["he_h", "he_b2", "he_b3", "he_vr", "opx_h", "opx_b2", "opx_b3", "opx_vr", "cit_h", "cit_b2", "cit_b3", "cit_vr", "dpi_h", "dpi_b2", "dpi_b3", "cfac"]],
    ["自用 AI 讀法", ["internal_ai_reading", "k_ai"]]]],
  ["全公司：資產負債、折現率、目標價規則、稅與現金流", [
    ["日期（資料截止、股價基準、評價、目標時點）", ["data_date", "px_date", "val_date", "tp_date"]],
    ["股價、股數與資產負債（評價日）", ["px", "px_val", "sh_out", "sh_dil_q3", "sh_basic_q3", "rsu_unv", "opt_n", "opt_k", "dil_inc", "dil_base", "cash_0", "cash_ye25", "debt_0", "debt_ye25", "fin_lease", "legal_oth", "ag_acc_x", "esc_in"]],
    ["特別股與遠期出售", ["mcps", "conv_sh", "pref_2H", "pref_27", "pref_28", "pref_29", "fwd_on", "fwd_sh", "fwd_cash"]],
    ["折現率", ["rf", "beta_cr", "beta_cd", "blume_a", "beta", "erp", "kd_sp", "kd", "wacc_adj", "g_tv", "ronic_sp"]],
    ["目標價規則", ["mid", "t_val", "t_q4", "tp_wA", "tp_wB", "tp_wC", "w_dcf", "t_tgt", "t_hold", "w_ntm27", "cons_tp", "dps_tgt", "sbc_add", "tv_alt"]],
    ["稅、股利、回購與現金", ["tax", "r_cash", "r_debt", "sbc_r", "dil", "wc_k", "dps", "dps_g", "bb_t", "min_cash", "mat_27", "lev_max", "icr_min"]],
    ["多州檢察長協議", ["ag_pay", "ag_y1", "ag_rest", "contg"]]]]],
 "blocks_to": {"命題模組": "自用 AI 算力", "@Thesis!A1": "自用 AI 算力", "@Thesis!A2": "自用 AI 算力", "@Seg!A2": "廣告本業（Family of Apps）"},
 "main": {"DCF 彙總": ["明確期現值", "終值現值", "企業價值 EV（主值", "＋淨現金", "DCF 情境價值推到 12 個月後"],
          "分部加總": ["Family of Apps", "Reality Labs", "分部加總 EV（主值）", "分部加總每股價值（主值", "  隱含 EV"],
          "1. 合併層": ["營收|FY2027", "營收|FY2028", "營業利益（財報口徑）|FY2028", "自由現金流|FY2027"]},
}
cfg = relayout(cfg, X, SPEC)
for _sec in cfg["sections"]:   # 龍捲風排除門檻說明（取自 Sens 頁；relayout 會改寫龍捲風標題，故另加說明區塊）
    if _sec["id"] == "sens":
        _i = next(i for i, b in enumerate(_sec["blocks"]) if b["type"] == "tornado")
        _sec["blocks"].insert(_i, {"type": "note", "parts": [f"@Sens!A{rm['SE_TORN']}"]})


# HTML 版面規格 v1.2：分區主題圖（照 nvidia/gen_html_json.py 的 _theme_blocks()；列號依標籤找、數字全取 xlsx）
def _theme_blocks():
    h7 = r7 + 1
    drv = {"type": "theme", "title": "自用 AI 在役容量（＝Wattonomics）vs 原口徑（資本支出推得），GW 右軸；AI 資本支出（情境 B，US$B）", "x": f"@Recon!B{h7}:H{h7}",
           "fmt": "n0", "fmt_r": "n1", "stack": False, "pad": 1.25,
           "bars": [{"name": f"@Recon!A{RC7['cxai']}", "y": f"@Recon!B{RC7['cxai']}:H{RC7['cxai']}", "color": "#c29a2b"}],
           "lines": [{"name": f"@Recon!A{RC7['mod']}", "y": f"@Recon!B{RC7['mod']}:H{RC7['mod']}", "axis": "r", "color": "#3b3838", "labels": True},
                     {"name": f"@Recon!A{RC7['cx']}", "y": f"@Recon!B{RC7['cx']}:H{RC7['cx']}", "axis": "r", "color": "#b4472f", "dash": True}],
           "caption": "長條＝情境 B 的 AI 資本支出；實線＝本模型在役容量（讀 Wattonomics 上線時程）；虛線＝資本支出當年轉在役的原口徑（只作敏感度與對照）。"}
    hp = PR["A"]["ads"] - 1
    plc = {"type": "theme", "title": "Family of Apps 營收（長條）與營業利益（線，右軸）× 情境 A／B／C（US$B）", "x": f"@PL!G{hp}:M{hp}",
           "fmt": "n0", "fmt_r": "n0", "stack": False, "pad": 1.2,
           "bars": [{"name": ["情境 ", s, "：營收"], "y": f"@PL!G{PR[s]['foa']}:M{PR[s]['foa']}", "color": c}
                    for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
           "lines": [{"name": ["情境 ", s, "：營業利益"], "y": f"@PL!G{PR[s]['oi_foa']}:M{PR[s]['oi_foa']}", "axis": "r", "color": c, "dash": s != "B", "labels": s == "B"}
                     for s, c in (("A", "#0b7a80"), ("B", "#3b3838"), ("C", "#b4472f"))],
           "caption": "長條＝三情境的 Family of Apps 營收；線＝Family of Apps 營業利益（財報口徑，已扣自用 AI 增量成本與新增 AI 折舊）。"}
    val = {"type": "theme", "title": "k 梯度：加權目標價 vs 現價（US$／股）", "x": f"@Recon!B{RC6['hk']}:G{RC6['hk']}", "fmt": "n0", "stack": False, "pad": 1.1,
           "bars": [{"name": f"@Recon!A{RC6['htp']}", "y": f"@Recon!B{RC6['htp']}:G{RC6['htp']}", "color": "#0b7a80"}],
           "lines": [{"name": f"@Recon!A{RC6['hpx']}", "y": f"@Recon!B{RC6['hpx']}:G{RC6['hpx']}", "color": "#3b3838"},
                     {"name": f"@Recon!A{RC6['hbuy']}", "y": f"@Recon!B{RC6['hbuy']}:G{RC6['hbuy']}", "color": "#b4472f", "dash": True}],
           "caption": f"@Recon!B{RC6['hcap']}"}
    hc = CR["B"]["ni"] - 1
    gate_ = {"type": "theme", "title": "收支關卡：逐年自由現金流扣股利（長條）與新增舉債（線）× 情境 A／B／C（US$B）", "x": f"@CF!G{hc}:K{hc}", "fmt": "n0", "stack": False,
             "bars": [{"name": ["情境 ", s, "：自由現金流扣股利"], "y": f"@CF!G{CR[s]['fcfd']}:K{CR[s]['fcfd']}", "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
             "lines": [{"name": ["情境 ", s, "：新增舉債"], "y": f"@CF!G{CR[s]['plug']}:K{CR[s]['plug']}", "color": c, "dash": s != "B", "labels": s == "B"} for s, c in (("A", "#0b7a80"), ("B", "#c29a2b"), ("C", "#b4472f"))],
             "caption": "長條＝各情境逐年自由現金流扣股利（收支關卡為其五年合計）；線＝現金低於最低現金時補足的新增舉債。"}
    S = {s_["id"]: s_ for s_ in cfg["sections"]}
    S["drivers"]["blocks"].insert(0, drv)
    S["pl"]["blocks"][0:0] = [plc, gate_]
    S["valuation"]["blocks"].insert(0, val)


_theme_blocks()
# ================= v0.2.1（2026-11-15）：個股 HTML 版面規格 v1.3 全套＋Coord 三層協調（樣板 Alphabet v0.3.9 gen_html_json.py；數字與文字全取自 xlsx） =================
V21 = rm["V"]
PLA, RVA, CRD, RMA = V21["PL"], V21["RV"], V21["CRD"], V21["RM"]


def _v13():
    S = {x["id"]: x for x in cfg["sections"]}
    PLW = wb["PL"]
    r13 = find("README", "v1.3 展示層")
    RM_ = {k: r13 + 1 + i for i, k in enumerate(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"])}
    assert str(wb["README"][f"A{RM_['ver']}"].value).startswith("v0.2.2"), wb["README"][f"A{RM_['ver']}"].value
    cfg["layout"] = "v1.3"
    old_note = cfg["hero"]["note"]
    cfg["hero"]["note"] = [f"@README!A{RM_['hero']}"]
    ph = PLA["hdr"]; ix = PLA["idx"]; sx = PLA["sidx"]
    # ---- 1. 命題 ----
    th = S["thesis"]["blocks"]
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM_[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th_chart = {"type": "theme", "title": "營收（十億美元，依分部堆疊）與營業利益率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}",
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3, "total": f"@PL!C{ix['rev']}:L{ix['rev']}",
                "bars": [{"name": [f"@PL!A{sx['s_ads']}"], "y": f"@PL!C{sx['s_ads']}:L{sx['s_ads']}", "color": "#6fa8dc"},
                         {"name": [f"@PL!A{sx['s_oth']}"], "y": f"@PL!C{sx['s_oth']}:L{sx['s_oth']}", "color": "#0b7a80"},
                         {"name": [f"@PL!A{sx['s_rl']}"], "y": f"@PL!C{sx['s_rl']}:L{sx['s_rl']}", "color": "#c29a2b"}],
                "lines": [{"name": [f"@PL!A{ix['opm']}"], "y": f"@PL!C{ix['opm']}:L{ix['opm']}", "axis": "r", "color": "#3b3838", "labels": True}],
                "caption": "堆疊＝Advertising、Other revenue（兩者合為 Family of Apps）與 Reality Labs 三個分部；折線＝合併營業利益率（右軸）。FY2026E 起為情境 B（基準）；新增 AI 折舊使營業利益率逐年下降。"}
    th[0:0] = [{"type": "note", "parts": [f"@README!A{RM_['rec']}"]}, th_chart]
    # ---- 2. 估值：四個目標價與現價置頂 ----
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b["type"] == "bars")
    va[i_bar] = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
                 "items": [{"label": f"@Val!A{VR[k]}", "value": f"@Val!B{VR[k]}", "series": sr} for k, sr in (("tpdcf", 1), ("tpcomp", 2), ("tp", 0), ("cons", 1))],
                 "ref": {"label": f"@Val!A{VR['px']}", "value": f"@Val!B{VR['px']}"}}
    bar = va.pop(i_bar)
    top = [bar, {"type": "note", "parts": [S["valuation"].pop("intro")]}, {"type": "note", "parts": list(old_note)}]
    S["valuation"]["blocks"] = top + va
    # ---- 3. 敏感度：加權目標價相對基準龍捲風圖置頂 ----
    SW = wb["Sens"]; t0 = find("Sens", "加權目標價相對基準（HTML"); r1_ = t0 + 2; r2_ = r1_
    while SW.cell(r2_ + 1, 1).value: r2_ += 1
    tors = []
    num = lambda col, r: isinstance(SW[f"{col}{r}"].value, (int, float))
    for col, side in (("B", "低"), ("C", "高")):
        big = [r for r in range(r1_, r2_ + 1) if num(col, r) and abs(SW[f"{col}{r}"].value) >= 0.1]
        small = [r for r in range(r1_, r2_ + 1) if num(col, r) and 0.0005 <= abs(SW[f"{col}{r}"].value) < 0.1]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        if big:
            tors.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
        if small:
            tors.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=big != [],
                             summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例"))
    tors.append({"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
                 "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1_, r2_ + 1)]})
    sb = S["sens"]["blocks"]
    for b in sb:
        if b["type"] == "tornado":
            b["collapse"] = True; b["summary"] = "展開：對照：加權目標價相對基準（原圖，標籤為模型輸入名稱）"
    first2 = [t for t in tors if not t.get("collapse")]
    S["sens"]["blocks"] = first2 + [{"type": "note", "parts": [S["sens"].pop("intro")]}] + [t for t in tors if t.get("collapse")] + sb
    # ---- 4. 反向 DCF（新）----
    A_ = RVA
    yh = A_["yh"]
    yr_fmt = {"y_rev": "n1", "y_ebit": "n1", "y_m": "p1", "y_df": "n3", "y_tvf": "n3", "y_k": "n3"}
    rev_ = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2", "blocks": [
        {"type": "note", "parts": [f"@RevDCF!A{A_['sum_dcf']}"]}, {"type": "note", "parts": [f"@RevDCF!A{A_['sum_cmp']}"]},
        {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
         "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{A_['m1']}|p1", f"@RevDCF!B{A_['m0']}|p1"],
                  ["營收規模（FY2028 起各年同乘的倍數）", f"@RevDCF!B{A_['s1']}|n2", "—"],
                  ["FY2028 營收（十億美元）", f"@RevDCF!B{A_['rv1']}|n0", f"@RevDCF!B{A_['rv0']}|n0"],
                  ["FY2028 營收年增率", f"@RevDCF!B{A_['g1']}|p1", f"@RevDCF!B{A_['g0']}|p1"],
                  ["折現率（WACC）", f"@RevDCF!B{A_['w1']}|p2", f"@RevDCF!B{A_['w0']}|p2"]]},
        {"type": "table", "caption": "分部加總法：現價隱含的倍數與模型", "head": ["項目", "現價隱含", "模型", "差距（現價隱含相對模型）"],
         "rows": [[f"@RevDCF!A{A_[k]}", f"@RevDCF!B{A_[k]}|{f}", f"@RevDCF!C{A_[k]}|{f}", f"@RevDCF!D{A_[k]}|sp1"] for k, f in (("sps", "n2"), ("sev", "n0"), ("smul", "n1"), ("scmp", "n2"))]},
        {"type": "table", "caption": "現金流折現法：現價隱含的企業價值與模型", "head": ["項目", "值"],
         "rows": [[f"@RevDCF!A{A_[k]}", f"@RevDCF!B{A_[k]}|{f}"] for k, f in (("px", "n2"), ("v0", "n2"), ("evr", "n0"), ("ev", "n0"), ("evx", "n2"), ("v_b", "n2"))]},
        {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFGH"], "collapse": True,
         "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）",
         "rows": [[f"@RevDCF!A{A_[k]}"] + [f"@RevDCF!{c}{A_[k]}|{f}" for c in "BCDEFGH"] for k, f in yr_fmt.items()]}]}
    # ---- 5. 損益：標準損益表 ----
    fmtu = lambda u: {"%": "p1", "US$": "n2", "十億股": "n2"}.get(u, "n1")
    rowx = lambda r, cols="CDEFGHIJKLM": [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmtu(PLW.cell(r, 2).value)}" for c in cols]
    head = [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in "CDEFGHIJKLM"]
    keys = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pti", "tax", "ni", "nm", "sh", "eps"]
    rid = PLA["ridx"]
    is_tbl = {"type": "table", "caption": "標準損益表（GAAP；十億美元，EPS 為美元）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B（基準）、最右欄為 2026 年前三季累計實績", "head": head,
              "rows": [rowx(ix[k]) for k in keys] + [{"sub": f"@PL!A{PLA['ratio']}"}] + [rowx(rid[k]) for k in ("r_niy",)]}
    chk_tbl = {"type": "table", "caption": "核對（應為零）", "collapse": True, "summary": "展開：標準損益表核對列", "head": head,
               "rows": [rowx(r) for r in range(PLA["chk"] + 1, PLA["chk"] + 6)]}
    rv_, gm_, opm_, nm_ = ix["rev"], ix["gm"], ix["opm"], ix["nm"]
    is_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": [f"@PL!A{rv_}"], "y": f"@PL!C{rv_}:L{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": [f"@PL!A{gm_}"], "y": f"@PL!C{gm_}:L{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": [f"@PL!A{opm_}"], "y": f"@PL!C{opm_}:L{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": [f"@PL!A{nm_}"], "y": f"@PL!C{nm_}:L{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP）。FY2025A 淨利率含一次性稅務費用，FY2026E 含前三季一次性稅務利益，預測期不估；毛利率逐年下降是新增 AI 折舊上升（模型假設新增折舊全數記入營業成本）。"}
    old = S["pl"]["blocks"]
    seg_old = next(b for b in old if b.get("type") == "table" and str(b.get("summary", "")).startswith("展開：情境 B 的分部"))
    rest = [b for b in old if b is not seg_old]
    for b in rest:
        if b.get("type") == "pl":
            b["collapse"] = True; b["summary"] = "展開：三情境損益（A／B／C 並列，含折線）"
    pl_intro = S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{PLA['note']}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{PLA['note2']}"]}, chk_tbl]
                   + ([{"type": "note", "parts": [pl_intro], "collapse": True, "summary": "展開：三情境損益頁說明"}] if pl_intro else []) + rest)
    # ---- 6. 驅動因子：分部營收推導 ----
    seg_old = dict(seg_old, collapse=False, caption="分部營收推導（情境 B；十億美元）：Advertising、Other revenue 與 Reality Labs 加總＝合併營收，另列營業利益與每股")
    seg_old.pop("summary", None)
    seg_chart = {"type": "theme", "title": "Advertising 營收（長條）× 情境 A／B／C（十億美元）", "x": f"@PL!G{hA}:M{hA}", "fmt": "n0", "stack": False, "pad": 1.2,
                 "bars": [{"name": ["情境 ", s, "：Advertising"], "y": f"@PL!G{PR[s]['ads']}:M{PR[s]['ads']}", "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
                 "caption": "Advertising＝廣告曝光數 × 每則廣告平均價格；三情境只調價格路徑（曝光路徑相同），Other revenue 與 Reality Labs 另依近四季年增收斂。"}
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM_['drv']}"]}, seg_chart, seg_old] + S["drivers"]["blocks"]
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    # ---- 7. 對帳：三層協調 ----
    CW = wb["Coord"]
    cr = lambda p_, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p_))
    f1h, f1b, f2h = cr("F1. 本模型讀取"), cr("F1b. 追公式"), cr("F2. C0（實績）")
    F1a = CRD["F1"][0]; F1e = CRD["wrows"]["act"]; r_tr = CRD["rtr"]
    rd1, rd2, rj, rjr, rg, rl = CRD["rd1"], CRD["rd2"], CRD["rj"], CRD["rjr"], CRD["rg"], CRD["rl"]
    rai, rmw = CRD["rai"], CRD["rmw"]
    rq = CRD["rq"]
    capx = {"type": "table", "caption": "資本支出（公司已公布實績 vs Wattonomics 隱含值；十億美元）", "head": ["讀數", "期間／口徑", "公司實績", "層級隱含", "差距", "狀態／判定"],
            "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in (rd1, rd2, rg)]
                    + [[f"@Coord!A{rj}", f"@Coord!B{rj}", "", "", "", f"@Coord!G{rj}"], [f"@Coord!A{rjr}", "", "", "", "", f"@Coord!G{rjr}"]]}
    nocmp = {"type": "table", "caption": "沒有可比公司實績的項目（參考；不判定）", "head": ["項目", "期間／口徑", "狀態"],
             "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!G{r}"] for r in (rai, rmw)]}
    cnt = {"type": "table", "caption": "實績對帳的項目數", "head": ["項目", "數"],
           "rows": [[f"@Coord!A{CRD['rn']}", f"@Coord!D{CRD['rn']}|n0"], [f"@Coord!A{CRD['rnb']}", f"@Coord!D{CRD['rnb']}|n0"], [f"@Coord!A{CRD['rnc']}", f"@Coord!D{CRD['rnc']}|n0"],
                   [f"@Coord!A{CRD['rsrc']}", f"@Coord!D{CRD['rsrc']}"]]}
    d3 = {"type": "table", "caption": "展開：實績對帳的輸入、口徑、出處與證據等級", "collapse": True, "summary": "展開：實績對帳的輸入、出處與證據等級",
          "head": ["項目", "期間／口徑", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級"],
          "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!H{r}", f"@Coord!I{r}", f"@Coord!J{r}", f"@Coord!K{r}"] for r in list(rq) + [rd1, rd2, rg, rl, rai, rmw]]}
    trc = {"type": "table", "caption": f"@Coord!A{f1b}", "head": ["輸出名稱 ← 公式鏈", "同源判定", "讀到什麼"],
           "rows": [[f"@Coord!A{r}", f"@Coord!G{r}", f"@Coord!H{r}"] for r in range(f1b + 2, r_tr)]}
    t1 = {"type": "table", "caption": f"@Coord!A{f1h}", "collapse": True, "summary": "展開：本模型讀取的 Tokenomics 名稱與 Wattonomics 對照名稱、值與版本",
          "head": ["名稱", "年度", "層級值", "本模型採用值", "單位", "層級與版本"],
          "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}|n3", f"@Coord!D{r}|n3", f"@Coord!E{r}", f"@Coord!F{r}"] for r in range(F1a, F1e + 1)
                   if CW.cell(r, 3).value is not None]}
    d0, d1 = CRD["rdump"]
    dump_t = {"type": "table", "caption": "展開：追公式傾印（腳本輸出：每格公式與值）", "collapse": True, "summary": "展開：追公式傾印（每格公式與值；腳本 trace_wattonomics.py 的輸出）",
              "head": ["儲存格（縮排＝依賴深度）", "值", "標記", "公式"],
              "rows": [[f"@Coord!A{r}", f"@Coord!D{r}|n3", f"@Coord!G{r}", f"@Coord!H{r}"] for r in range(d0, d1 + 1)]}
    S["recon"]["blocks"].extend([
        {"type": "h3", "text": "三層協調：讀取層級與實績對帳"},
        {"type": "note", "parts": [f"@Coord!A{F1a}"]},
        {"type": "note", "parts": [f"@Coord!D{CRD['rs']}"]}, capx, nocmp, cnt, d3,
        {"type": "note", "parts": [f"@Coord!B{CRD['rr']}"]}, {"type": "note", "parts": [f"@Coord!B{CRD['rsame']}"]}, {"type": "note", "parts": [f"@Coord!B{CRD['rk']}"]},
        {"type": "h3", "text": "追公式：層級值是否讀公司實績或目標"}, trc, {"type": "note", "parts": [f"@Coord!A{r_tr}"]}, dump_t, t1])
    # ---- 8. 資料狀態：本版要點 ----
    S["status"]["blocks"][0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM_['ver']}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    S["revdcf"] = rev_
    cfg["sections"] = [S[k] for k in order]


_v13()

json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

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
rowfmt_conc[str(VR["tier"])] = {"B": ""}; rowfmt_conc[str(VR["rec"])] = {"B": ""}; rowfmt_conc[str(VR["meth"])] = {"B": ""}
for k in ("tp_r2", "tp_r1"): rowfmt_conc[str(VR[k])] = {"B": "n2", "C": "sp1", "D": ""}
for k in ("tp_dcf12", "tp_sotp12"): rowfmt_conc[str(VR[k])] = {"B": "n2", "C": "n2"}
rowfmt_conc[str(VR["wsum"])] = {"B": "n2"}; rowfmt_conc[str(VR["chk_rd"])] = {"B": "n3"}
dsum0, dsum1 = VR["pvx"], VR["dcf12"]
sotp0 = VR["s_foa"]; sotp1 = VR["simp"]
rf_s = {str(VR["tvsh"]): {"B": "p1", "C": "p1", "D": "p1"}, str(VR["shv"]): {"B": "n3", "C": "n3", "D": "n3"}, str(VR["ppref"]): {"B": "n2", "C": "n2", "D": "n2"}}
rf_t = {str(VR["ssh"]): {"B": "n3", "C": "n3", "D": "n3"}, str(VR["simp"]): {"B": "x1", "C": "x1", "D": "x1"}}
gate = rm["GATE"]
r1 = find("Recon", "1. 合併層"); r2 = find("Recon", "2. 逐分部"); r3 = find("Recon", "3. M4"); r4 = find("Recon", "4. 一次性")
rend = wb["Recon"].max_row
se_end = max(x[3] for x in rm["SE_ROWS"])
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
 "question": "FY2028 AI 增量營業利益（主值讀法：自用 AI 影子效益＝自用在役 MW × Tokenomics 每 MW 持有成本）能否覆蓋 AI 資本支出的資金成本；收支關卡＝FY2026 到 FY2030 自由現金流扣股利合計",
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
                           ("gl_impr", "p0"), ("gl_price", "p0"), ("gl_other", "p0"), ("rl_loss32", "n1"), ("internal_ai_reading", "n0"), ("beta", "n3"), ("g_tv", "p0"), ("w_dcf", "p0"), ("contg", "n1"))]},
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
     "rowgen": {"sheet": "CF", "rows": f"{gate}-{gate + 3}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"], "row_fmt": {str(gate + 3): {"B": "", "C": "", "D": ""}}}},
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
     "rowgen": {"sheet": "Recon", "rows": f"{r4 + 2}-{rend}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""]}}]},
  {"id": "sens", "title": "敏感度", "intro": "@Sens!A1", "blocks": [
    {"type": "tornado", "title": "加權目標價相對基準的變動（一次改一個輸入，整本重算）", "sheet": "Sens", "rows": f"4-{se_end}", "label_col": "A", "value_col": "E", "fmt": "sp1"},
    {"type": "table", "collapse": True, "summary": "展開：全部案例（含收支關卡 B）", "head": ["驅動", "設定", "方向", "加權目標價", "相對基準", "收支關卡 B"],
     "rowgen": {"sheet": "Sens", "rows": f"4-{se_end}", "cols": ["A", "B", "C", "D", "E", "H"], "fmts": ["", "", "", "n2", "sp1", "n1"]}},
    {"type": "kv", "items": [["@Sens!A" + str(rm["SE_BASE"]), "@Sens!D" + str(rm["SE_BASE"]) + "|n2"], ["@Sens!A" + str(rm["SE_STALE"]), "@Sens!D" + str(rm["SE_STALE"]) + "|n2"]]}]},
  {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"],
                             ["Tokenomics 快照", "@Sources!B" + str(find("Sources", "Tokenomics 快照", 2))]]},
    {"type": "h3", "text": "資料與口徑"},
    {"type": "list", "items": ["@Q!A1", "@Q!A2", "@In!A2", "@Sens!A2"]},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A", "C", "F"], "In": ["B", "H", "I"]}, "empty": "沒有標為未找到的項目。"}]}
 ]
}
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

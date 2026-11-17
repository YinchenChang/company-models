# 由 build_applovin.py 的列號對照（ROWMAP）產生 software_ai/applovin/html.json（列號隨模型變動，勿手改 html.json）。
# 版面：個股 HTML 版面規格 v1.3（thesis→valuation→sens→revdcf→pl→drivers→recon→status）；數字與文字全取自 xlsx（@工作表!儲存格|格式）。
# 用法：python3 software_ai/applovin/gen_html_json.py <成品 xlsx> <rowmap.json> software_ai/applovin/html.json
import json, sys, openpyxl
X, RMF, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
rm = json.load(open(RMF, encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)
cr_ = rm["RCR"]
VR, PLR, CFR, RM, THR, RVR, RCR, TKR, INR = rm["VR"], rm["PLR"], rm["CFR"], rm["RM"], rm["THR"], rm["RVR"], rm["RCR"], rm["TKR"], rm["INR"]
THS = rm["THS"]
PB = PLR["B"]
hp = PB["hdr"]
YC = "CDEFGHIJKL"


def fmtu(u):
    return {"%": "p1", "US$": "n2", "百萬股": "n1", "百萬件": "n0"}.get(u, "n1")


def plrow(r, cols=YC + "M"):
    u = wb["PL"].cell(r, 2).value
    return [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmtu(u)}" for c in cols]


def find(sh, text, col=1, start=1):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        if str(ws.cell(r, col).value or "").startswith(text): return r
    raise KeyError((sh, text))


hero = {"rec": f"@Val!B{VR['rec']}",
        "tiles": [[f"@Val!A{VR['tp']}", f"@Val!B{VR['tp']}|n2", ["上檔空間 ", f"@Val!B{VR['up']}|sp1"]],
                  [f"@Val!A{VR['tpdcf']}", f"@Val!B{VR['tpdcf']}|n2", ""], [f"@Val!A{VR['tpcomp']}", f"@Val!B{VR['tpcomp']}|n2", ""],
                  [f"@Val!A{VR['px']}", f"@Val!B{VR['px']}|n2", ""]],
        "note": [f"@README!A{RM['hero']}"]}
hdrs = f"@PL!C{hp}:L{hp}"
# ---- thesis ----
thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "note", "parts": [f"@README!A{RM['rec']}"]},
    {"type": "theme", "title": "營收（百萬美元）與營業利益率：FY2023A–FY2032E（情境 B）", "x": hdrs, "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "bars": [{"name": [f"@PL!A{PB['s_ads']}"], "y": f"@PL!C{PB['s_ads']}:L{PB['s_ads']}", "color": "#0b7a80"}],
     "lines": [{"name": [f"@PL!A{PB['opm']}"], "y": f"@PL!C{PB['opm']}:L{PB['opm']}", "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "長條＝單一官方分部 Advertising 營收（百萬美元）；折線＝合併營業利益率（右軸，GAAP）。FY2026E 起為情境 B（基準）。"},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]},
    {"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}, {"type": "note", "parts": [f"@README!A{RM['prior']}"]}]}
# ---- valuation ----
dcf_rows = ["pvx", "pvtv", "ev", "cash", "debt", "fl", "legal", "eqinv", "eqv", "shs", "ps0", "ps12", "tvsh"]
dcf_fmt = {str(VR["tvsh"]): {c: "p1" for c in "BCD"}, str(VR["shs"]): {c: "n3" for c in "BCD"}, str(VR["ps0"]): {c: "n2" for c in "BCD"}, str(VR["ps12"]): {c: "n2" for c in "BCD"}}
cmp_fmt = {str(VR["fw"]): {c: "n3" for c in "BCD"}, str(VR["mult"]): {c: "n2" for c in "BCD"}, str(VR["rev26"]): {c: "n0" for c in "BCD"}, str(VR["rev27"]): {c: "n0" for c in "BCD"},
           str(VR["ttm_t"]): {c: "n0" for c in "BCD"}, str(VR["tpc_s"]): {c: "n2" for c in "BCD"}}
pe_fmt = {str(VR[k]): {c: "n2" for c in "BCD"} for k in ("eps27", "eps28", "ntm", "pe_v1", "pe_v2", "pe_w1", "pe_w2", "pe_tp1", "pe_tp2")}
wt_fmt = {str(VR["w"]): {c: "p0" for c in "BCD"}, str(VR["tp_s"]): {c: "n2" for c in "BCD"}, str(VR["up_s"]): {c: "sp1" for c in "BCD"}}
valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{VR[k]}", "value": f"@Val!B{VR[k]}", "series": sr} for k, sr in (("tpdcf", 1), ("tpcomp", 2), ("tp", 0), ("cons", 1))],
     "ref": {"label": f"@Val!A{VR['px']}", "value": f"@Val!B{VR['px']}"}},
    {"type": "note", "parts": [f"@Val!A1"]}, {"type": "note", "parts": [f"@Val!A2"]},
    {"type": "kv", "items": [[f"@Val!A{VR[k]}", f"@Val!B{VR[k]}|{f}"] for k, f in (("up", "sp1"), ("vrate", ""), ("gate", "n0"), ("rec", ""), ("rf", "p2"), ("beta", "n3"), ("ke", "p2"), ("kdat", "p2"), ("wacc", "p2"), ("ronic", "p2"), ("gtv", "p1"), ("shval", "n1"))]},
    {"type": "table", "caption": "DCF 彙總（三情境；百萬美元，每股為美元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['pvx']}-{VR['ps12']},{VR['tvsh']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n0", "n0", "n0"], "row_fmt": dcf_fmt}},
    {"type": "note", "parts": [f"@README!A{RM['g6_2']}"]},
    {"type": "note", "parts": [f"@README!A{RM['g7_1']}"]},
    {"type": "table", "caption": "Comparable 主腿：同業中位 P/S × 相對倍數 × 目標時點 TTM 營收 ÷ 稀釋股數（三情境）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['fw']}-{VR['tpc_s']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n2", "n2", "n2"], "row_fmt": cmp_fmt}},
    {"type": "table", "caption": "前瞻本益比對照（不進加權目標價；固定相對倍數與現行相對值兩讀）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['eps27']}-{VR['pe_tp2']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n2", "n2", "n2"], "row_fmt": pe_fmt}},
    {"type": "note", "parts": [f"@README!A{RM['g6_3']}"]},
    {"type": "table", "caption": "情境加權", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['w']}-{VR['up_s']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n2", "n2", "n2"], "row_fmt": wt_fmt}},
    {"type": "note", "parts": [f"@README!A{RM['g5_div']}"]},
    {"type": "table", "caption": "DCF（股票薪酬稅後加回）對照讀數：只作對照，不進加權目標價", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['sb_pvx']}-{VR['sb_ps12']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n0", "n0", "n0"], "row_fmt": {str(VR["sb_ps12"]): {c: "n2" for c in "BCD"}}}},
    {"type": "kv", "items": [[f"@Val!A{VR[k]}", f"@Val!B{VR[k]}|{f}"] for k, f in (("sb_dcf", "n2"), ("sb_tp", "n2"), ("sb_up", "sp1"), ("sb_dd", "n2"))]},
    {"type": "note", "parts": [f"@Recon!A{cr_['sb_txt']}"]},
    {"type": "table", "caption": "同業 P/S（Comparable 主腿；資料日與逾期標示）", "head": ["同業", "P/S", "資料日", "逾期？", "是否納入中位數"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['psp_ps_pins']}-{VR['ps_rerate']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "", "", ""],
                "row_fmt": {str(VR["ps_rerate"]): {"B": "sp1"}, str(VR["ps_ratio"]): {"B": "n2"}, str(VR["ps_q"]): {"B": ""}, str(VR["ps_db"]): {"B": "n2"}}}},
    {"type": "table", "caption": "同業預估本益比（對照列用；不進目標價）", "head": ["同業", "預估本益比", "資料日", "逾期？", "口徑"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['peer_pe_pins']}-{VR['mult_med']},{VR['app_pe']}-{VR['rerate']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "", "", ""],
                "row_fmt": {str(VR["rerate"]): {"B": "sp1"}, str(VR["app_ratio"]): {"B": "n2"}}}}]}
# ---- sens ----
t1, t2 = rm["SE_T1"], rm["SE_T2"]
sens = {"id": "sens", "title": "敏感度", "blocks": [
    {"type": "tornado", "title": "加權目標價相對基準：輸入取低值", "sheet": "Sens", "rows": f"{t1}-{t2}", "label_col": "A", "value_col": "B", "fmt": "sp1"},
    {"type": "tornado", "title": "加權目標價相對基準：輸入取高值", "sheet": "Sens", "rows": f"{t1}-{t2}", "label_col": "A", "value_col": "C", "fmt": "sp1"},
    {"type": "note", "parts": ["@Sens!A1"]}, {"type": "note", "parts": ["@Sens!A2"]},
    {"type": "table", "collapse": True, "summary": "展開：全部案例（含 DCF 單法、評等與收支關卡）", "head": ["驅動", "設定", "方向", "加權目標價", "相對基準", "DCF 目標價", "投資建議", "收支關卡"],
     "rowgen": {"sheet": "Sens", "rows": f"4-{rm['SE_END']}", "cols": ["A", "B", "C", "D", "E", "J", "L", "H"], "fmts": ["", "", "", "n2", "sp1", "n2", "", "n0"]}},
    {"type": "kv", "items": [[f"@Sens!A{rm['SE_BASE']}", f"@Sens!D{rm['SE_BASE']}|n2"], [f"@Sens!A{rm['SE_STALE']}", f"@Sens!D{rm['SE_STALE']}|n2"],
                             ["模型狀態", f"@Sens!D{rm['SE_STATUS']}"], ["build_id", f"@Sens!D{rm['SE_BID']}"]]}]}
# ---- revdcf ----
revdcf = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2", "blocks": [
    {"type": "note", "parts": [f"@RevDCF!A{RVR['sum_dcf']}"]}, {"type": "note", "parts": [f"@RevDCF!A{RVR['sum_cmp']}"]},
    {"type": "table", "caption": "現價隱含值與模型（情境 B）", "head": ["項目", "數值"],
     "rows": [[f"@RevDCF!A{RVR[k]}", f"@RevDCF!B{RVR[k]}|{f}"] for k, f in (("px", "n2"), ("mcap", "n0"), ("nc", "n0"), ("evm", "n0"), ("pvx", "n0"), ("tvn", "n0"), ("tvm", "n0"), ("nopn", "n0"), ("nopm", "n0"), ("scl", "n2"),
                                                                            ("rv0", "n0"), ("rv1", "n0"), ("g0", "p1"), ("g1", "p1"), ("w0", "p2"), ("pe_px", "n1"), ("pe_m", "n1"), ("pe_gap", "sp1"))]}]}
# ---- pl ----
rows_is = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "int", "oth", "pti", "tax", "ni", "nm", "shd", "eps", "niy"]
scn_rows = []
for X in "ABC":
    for k in ("rev", "oi", "eps"):
        scn_rows.append(plrow(PLR[X][k], "GHIJKL"))
pl = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "note", "parts": ["@PL!A1"]},
    {"type": "theme", "title": "營收（百萬美元）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": hdrs, "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
     "bars": [{"name": [f"@PL!A{PB['rev']}"], "y": f"@PL!C{PB['rev']}:L{PB['rev']}", "color": "#6fc7cb"}],
     "lines": [{"name": [f"@PL!A{PB[k]}"], "y": f"@PL!C{PB[k]}:L{PB[k]}", "axis": "r", "color": c, "labels": k == "gm", "dash": k == "nm"} for k, c in (("gm", "#0b7a80"), ("opm", "#c29a2b"), ("nm", "#6f6bae"))],
     "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP）。"},
    {"type": "table", "caption": "標準損益表（GAAP；百萬美元，EPS 為美元）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B（基準）、最右欄為 2026 年前三季累計實績",
     "head": [f"@PL!A{hp}"] + [f"@PL!{c}{hp}" for c in YC + "M"], "rows": [plrow(PB[k]) for k in rows_is]},
    {"type": "note", "parts": ["@PL!A2"]},
    {"type": "table", "caption": "Non-GAAP：調整後 EBITDA（本模型口徑）", "head": [f"@PL!A{hp}"] + [f"@PL!{c}{hp}" for c in YC + "M"], "rows": [plrow(PB["aeb"]), plrow(PB["aebm"])]},
    {"type": "table", "collapse": True, "summary": "展開：三情境營收、營業利益與 EPS", "head": [f"@PL!A{hp}"] + [f"@PL!{c}{hp}" for c in "GHIJKL"], "rows": scn_rows},
    {"type": "table", "collapse": True, "summary": "展開：核對列（應為零）", "head": [f"@PL!A{hp}"] + [f"@PL!{c}{hp}" for c in YC + "M"],
     "rows": [plrow(PB[k]) for k in ("chk_seg", "chk_oi", "chk_pti", "chk_ni")]},
    {"type": "table", "caption": "現金流與收支關卡（情境 A／B／C）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "CF", "rows": f"{rm['GATE']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n0", "n0", "n0"]}}]}
# ---- drivers ----
cmp_r = [THR[k] for k in ("c_inc", "c_ero", "c_llm", "c_net", "c_cum")]
ue_rows = [THR[k] for k in ("ue_rev", "ue_dc", "ue_llm", "ue_oc", "ue_gp", "ue_opx", "ue_oi", "ue_sbc")]
SCN = {k: INR for k in ()}
scn_in = {}
ws_in = wb["In"]
for r in range(1, ws_in.max_row + 1):
    if str(ws_in.cell(r, 1).value) in ("scn_u", "scn_e", "scn_cdiv", "scn_dcr", "scn_w"): scn_in[ws_in.cell(r, 1).value] = r
H_IN = rm["SCN_SEC"] + 1
drivers = {"id": "drivers", "title": "驅動因子：單位經濟與 AI 淨效果", "blocks": [
    {"type": "note", "parts": [f"@README!A{RM['drv']}"]},
    {"type": "theme", "title": "AI 淨效果（營業利益口徑，百萬美元）× 情境 A／B／C", "x": f"@Thesis!G{THS['B|hdr']}:L{THS['B|hdr']}", "fmt": "n0", "stack": False, "pad": 1.2,
     "bars": [{"name": ["情境 ", s, "：AI 淨效果"], "y": f"@Thesis!G{THS[s + '|net']}:L{THS[s + '|net']}", "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
     "caption": "AI 淨效果＝(AI 增量營收＋分流侵蝕)×貢獻利潤率（一減資料中心成本率減其他營業成本率）減 LLM 推論成本；三情境差距來自 AI 提升、新垂直、分流與資料中心成本率。"},
    {"type": "note", "parts": ["@Thesis!A1"]}, {"type": "note", "parts": ["@Thesis!A2"]},
    {"type": "table", "caption": "單位經濟：每一千美元營收（情境 B）", "head": ["項目"] + [f"@Thesis!{c}{THR['hdr_ue']}" for c in YC],
     "rows": [[f"@Thesis!A{r}"] + [f"@Thesis!{c}{r}|n0" for c in YC] for r in ue_rows]},
    {"type": "table", "caption": "AI 淨效果三情境對照", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@Thesis!A{r}", f"@Thesis!C{r}|n0", f"@Thesis!D{r}|n0", f"@Thesis!E{r}|n0"] for r in cmp_r]},
    {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；In 頁最後一節）", "head": [f"@In!{c}{H_IN}" for c in "BDEF"],
     "rows": [[f"@In!B{scn_in[k]}"] + [f"@In!{c}{scn_in[k]}|p1" for c in "DEF"] for k in ("scn_u", "scn_e", "scn_cdiv", "scn_dcr", "scn_w")]},
    {"type": "table", "caption": "單一參數區間（In 頁；基準進主模型，敏感度低／高只用於敏感度頁；三情境共用）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!C3", "@In!G3"],
     "rows": [[f"@In!B{INR[k]}"] + [f"@In!{c}{INR[k]}|{f}" for c in "DEF"] + [f"@In!C{INR[k]}", f"@In!G{INR[k]}|tag"]
              for k, f in (("beta", "n3"), ("erp", "p1"), ("g_tv", "p1"), ("ginst0", "p1"), ("nrpi_base", "p1"), ("u_decay", "n2"), ("e_decay", "n2"), ("tk_px", "n3"), ("tk_mult", "n1"), ("llm_tasks27", "n0"),
                           ("sm_r", "p2"), ("rd_r", "p2"), ("ga_r", "p2"), ("sbc_r", "p2"), ("tax", "p1"), ("capex_r", "p2"), ("wc_k", "p2"), ("bb_pct", "p0"), ("w_dcf", "p0"), ("pe_rel", "n2"))]}]}
# ---- recon ----
def rrow(r, fm):
    return [f"@Recon!{c}{r}|{f}" if f != "T" else f"@Recon!{c}{r}" for c, f in zip("ABCDEFGH", fm)]
cr = RCR
s1 = cr["s1"] + 2
recon = {"id": "recon", "title": "對帳：共識、指引、DCF 與 Comparable 落差、反解", "intro": "@Recon!A1", "blocks": [
    {"type": "table", "caption": f"@Recon!A{cr['s1']}", "head": ["項目", "年度", "情境 A", "情境 B", "情境 C", "共識", "B 對共識", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{s1}-{s1 + 2}", "cols": list("ABCDEFGH"), "fmts": ["", "", "n0", "n0", "n0", "n0", "sp1", ""]}},
    {"type": "note", "parts": [f"@Recon!A{cr['eps_note']}"]},
    {"type": "table", "caption": f"@Recon!A{cr['s2']}", "head": ["項目", "本模型（情境 B）", "指引低", "指引高", "情境 A", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{cr['g_rev']}-{cr['g_eb']}", "cols": list("ABCDEF"), "fmts": ["", "n0", "n0", "n0", "n0", ""], "row_fmt": {str(cr["g_gap"]): {"B": "sp1", "E": "sp1"}}}},
    {"type": "note", "parts": [f"@Recon!A{cr['gap_txt']}"]},
    {"type": "table", "caption": f"@Recon!A{cr['s3']}", "head": ["項目", "數值", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{cr['gap_comp']}-{cr['gap_ke']}", "cols": list("ABC"), "fmts": ["", "n2", ""],
                "row_fmt": {str(cr["gap_pe"]): {"B": "n1"}, str(cr["gap_ke"]): {"B": "p1"}}}},
    {"type": "table", "caption": f"@Recon!A{cr['s4']}", "head": ["主導變數", "基準值", "使加權目標價＝現價", "使上檔空間為買進邊界", "狀態"],
     "rows": [[f"@Recon!A{rm['SE_KREQ'][k]}"] + [f"@Recon!{c}{rm['SE_KREQ'][k]}|{f}" for c, f in zip("BCD", (("n3",) * 3 if k in ("beta", "pe_rel") else ("p1",) * 3))] + [f"@Recon!E{rm['SE_KREQ'][k]}"] for k in ("pe_rel", "beta", "u0_B", "e0_B")]},
    {"type": "table", "caption": f"@Recon!A{cr['s8']}", "head": ["項目", "數值", "說明"],
     "rows": [[f"@Recon!A{rm['RL'][k]}", f"@Recon!B{rm['RL'][k]}|{f}", f"@Recon!C{rm['RL'][k]}"] for k, f in (("ps", "n3"), ("psg", "n3"), ("geo", "n3"), ("gr", "n2"), ("cg", "p1"), ("pg", "p1"), ("div", "n2"), ("lo", "n3"), ("hi", "n3"), ("pso", "n3"), ("pgo", "n3"))]},
    {"type": "note", "parts": [f"@Recon!A{cr['rel_txt']}"]},
    {"type": "note", "parts": [f"@README!A{RM['gate3_d']}"]},
    {"type": "note", "parts": [f"@Recon!A{cr['ver_txt']}"]},
    {"type": "table", "caption": f"@Recon!A{cr['s5']}", "head": ["項目", "數值", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{TKR['t_terra']}-{TKR['share27']}", "cols": list("ABC"), "fmts": ["", "n3", ""], "row_fmt": {str(TKR["cost27"]): {"B": "n1"}, str(TKR["share27"]): {"B": "p1"}}}}]}
# ---- status ----
rk0 = find("README", "市場資料日期"); rf0 = find("README", "主要失效條件"); rr0 = find("README", "風險登錄"); sh0 = find("README", "工作表：")
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期（資料截止日、成品日、股價基準日）", "@README!A1"], ["build_id", f"@Sens!D{rm['SE_BID']}"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "h3", "text": "資料與口徑"}, {"type": "list", "items": ["@In!A2", "@Hist!A1", "@Hist!A2", f"@README!A{RM['na']}", f"@README!A{RM['pr_note']}"]},
    {"type": "h3", "text": "市場資料日期、失效條件與風險登錄"},
    {"type": "list", "items": [f"@README!A{r}" for r in list(range(rk0 + 1, rf0 - 1)) + list(range(rf0 + 1, rr0 - 1)) if wb["README"].cell(r, 1).value]},
    {"type": "table", "caption": "風險登錄", "head": ["風險", "是否建模", "理由", "觀察指標"], "rowgen": {"sheet": "README", "rows": f"{rr0 + 2}-{sh0 - 2}", "cols": ["A", "B", "C", "D"]}},
    {"type": "h3", "text": "共通必修對照"}, {"type": "list", "items": [f"@README!A{r}" for r in range(find("README", "G2-1"), find("README", "G2-10") + 1)]},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A"], "In": ["B", "H", "I"]}, "empty": "沒有標為未找到的項目。"}]}
cfg = {"company": "AppLovin", "ticker": "APP", "thesis": f"@README!A{RM['thesis']}",
       "question": "AI 提升單次安裝淨收入與新垂直的增量，是否足以抵銷 AI 聊天介面的分流與運算成本，並支撐現價隱含的利潤路徑；收支關卡為五年自由現金流扣股利合計。",
       "scenarios": [{"key": "A", "name": "情境 A（AI 效益充分釋放）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（AI 效益遞延）"}],
       "wide_text": True, "hero": hero, "layout": "v1.3", "sections": [thesis, valuation, sens, revdcf, pl, drivers, recon, status]}
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

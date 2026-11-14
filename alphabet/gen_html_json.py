# 由 build_alphabet.py 的列號對照（ROWMAP）產生 alphabet/html.json（列號隨模型變動，勿手改 html.json）。
# 用法：python3 alphabet/gen_html_json.py <成品 xlsx> <rowmap.json> alphabet/html.json
import json, sys, openpyxl
X, RM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
rm = json.load(open(RM, encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)
T, PR, CR, VR, SR = rm["T"], rm["PR"], rm["CR"], rm["VR"], rm["SR"]
def find(sh, text, col=1, start=1):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        if str(ws.cell(r, col).value or "").startswith(text): return r
    raise KeyError((sh, text))
PA, PB, PCc = PR["A"], PR["B"], PR["C"]
hA = PA["search"] - 1
def rng(sh, r, c1="G", c2="M"): return f"@{sh}!{c1}{r}:{c2}{r}"
def met(name, key, fmt, chart=False, sheet="PL", d=PR, ymin=None):
    m = {"name": name, "fmt": fmt, "rows": {s: rng(sheet, d[s][key]) for s in "ABC"}}
    if chart: m["chart"] = True
    if ymin: m["ymin"] = ymin
    return m
conc0 = VR["tpdcf"]; conc1 = VR["chk_rd"]
rowfmt_conc = {}
for k in ("up", "upc", "scenup", "tvsh2"):
    rowfmt_conc[str(VR[k])] = {"B": "sp1", "C": "sp1", "D": "sp1"}
rowfmt_conc[str(VR["tier"])] = {"B": ""}; rowfmt_conc[str(VR["rec"])] = {"B": ""}; rowfmt_conc[str(VR["meth"])] = {"B": ""}
for k in ("tp_r2", "tp_r1"): rowfmt_conc[str(VR[k])] = {"B": "n2", "C": "sp1", "D": ""}
for k in ("tp_dcf12", "tp_sotp12"): rowfmt_conc[str(VR[k])] = {"B": "n2", "C": "n2"}
rowfmt_conc[str(VR["wsum"])] = {"B": "n2"}; rowfmt_conc[str(VR["chk_rd"])] = {"B": "n3"}
dsum0, dsum1 = VR["pvx"], VR["dcf12"]
sotp0 = VR["s_yt"]; sotp1 = VR["simp"]
rf_s = {str(VR["tvsh"]): {"B": "p1", "C": "p1", "D": "p1"}, str(VR["shv"]): {"B": "n3", "C": "n3", "D": "n3"}, str(VR["ppref"]): {"B": "n2", "C": "n2", "D": "n2"}}
rf_t = {str(VR["ssh"]): {"B": "n3", "C": "n3", "D": "n3"}, str(VR["simp"]): {"B": "x1", "C": "x1", "D": "x1"}}
gate = rm["GATE"]
r1 = find("Recon", "1. 合併層"); r2 = find("Recon", "2. 逐分部"); r3 = find("Recon", "3. v0.3"); rend = find("Recon", "檢查：(a)")
r4 = find("Recon", "4. v0.3.4"); rend4 = find("Recon", "檢查：10-Q")
se_end = max(x[3] for x in rm["SE_ROWS"])
thh = T["adds"] - 1
th_rows = ",".join(str(T[k]) for k in ("end_ext", "avg_ext", "ai_rev", "permw", "ai_ebitda", "ai_em", "capchk", "cx_ai", "cx_rev", "da_ai", "da_ext", "da_int", "int_cost"))
th_fmt = {str(T["permw"]): {c: "n2" for c in "EFGHIJKLM"}, str(T["ai_em"]): {c: "p1" for c in "EFGHIJKLM"}, str(T["capchk"]): {c: "p1" for c in "EFGHIJKLM"},
          str(T["cx_rev"]): {c: "n2" for c in "EFGHIJKLM"}, str(T["end_ext"]): {c: "n0" for c in "EFGHIJKLM"}, str(T["avg_ext"]): {c: "n0" for c in "EFGHIJKLM"}}
seg_rows = ",".join(str(PB[k]) for k in ("search", "yt", "net", "subs", "gs", "ai", "tpu", "co", "cloud", "ob", "hedge", "rev", "oi_gs", "oi_cl", "oi_ob", "oi_abc", "oi", "om", "da", "da_ai", "ni", "eps", "epsx"))
seg_fmt = {str(PB["om"]): {c: "p1" for c in "CDEFGHIJKLM"}, str(PB["eps"]): {c: "n2" for c in "CDEFGHIJKLM"}, str(PB["epsx"]): {c: "n2" for c in "CDEFGHIJKLM"}}
cf_rows = ",".join(str(CR["B"][k]) for k in ("cfo", "cx_ai", "cx_na", "cx", "fcf", "div", "pref", "fcfd", "anth", "bb", "plug", "cend", "dend", "nc"))
INK = {str(wb["In"].cell(r, 1).value): r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}
SCN_H = find("In", "J. 情境驅動對照") + 1
SCN = {k: INK[k] for k in ("scn_d", "scn_m", "scn_w")}
cfg = {
 "company": "Alphabet",
 "ticker": "GOOGL",
 "thesis": "@README!A3",
 "question": "Search & other 在 C 情境（AI 助理分流、每次查詢變現下降）下 FY2028 營收與利益率會掉多少，Google Cloud 的 AI 營收能否抵銷；加計 AI 折舊與股權發行後的每股價值",
 "scenarios": [{"key": "A", "name": "情境 A（AI 搜尋擴大變現）"}, {"key": "B", "name": "情境 B（近四季成長收斂；基準）"}, {"key": "C", "name": "情境 C（查詢分流、變現下降）"}],
 "sections": [
  {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "kv", "items": [["投資建議", "@README!A5"], ["重點觀察一", "@README!A6"], ["重點觀察二", "@README!A7"], ["重點觀察三", "@README!A8"]]},
    {"type": "h3", "text": "與上一版的差異"},
    {"type": "list", "items": [f"@README!A{r}" for r in range(find("README", "與上一版的差異") + 1, find("README", "與上一版的差異") + 6)]}]},
  {"id": "drivers", "title": "驅動因子 → 推導量", "intro": "@Thesis!A1", "blocks": [
    {"type": "table", "caption": "驅動 → 推導對照表", "head": ["頁", "輸入", "推導量", "因果方向"],
     "rowgen": {"sheet": "README", "rows": f"{find('README', '頁') + 1}-{find('README', 'Sens', 1, find('README', '頁'))}", "cols": ["A", "B", "C", "D"]}},
    {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；In 頁 J 節）", "head": [f"@In!{c}{SCN_H}" for c in "BDEFCG"],
     "rows": [[f"@In!B{SCN[k]}", f"@In!D{SCN[k]}|{f}", f"@In!E{SCN[k]}|{f}", f"@In!F{SCN[k]}|{f}", f"@In!C{SCN[k]}", f"@In!G{SCN[k]}|tag"] for k, f in (("scn_d", "p0"), ("scn_m", "p0"), ("scn_w", "p0"))]},
    {"type": "table", "caption": "單一參數區間（In 頁；基準進主模型，敏感度低／高只用於敏感度頁、不進目標價；三情境共用）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!C3", "@In!G3"],
     "rows": [[f"@In!B{INK[k]}", f"@In!D{INK[k]}|{f}", f"@In!E{INK[k]}|{f}", f"@In!F{INK[k]}|{f}", f"@In!C{INK[k]}", f"@In!G{INK[k]}|tag"]
              for k, f in (("mw_q2", "n0"), ("pace", "n0"), ("ext_sh", "p0"), ("long_sh", "p0"), ("tpu_f", "n2"), ("capex_g", "n0"), ("life_it", "n0"),
                           ("gl_search", "p0"), ("gl_yt", "p0"), ("gl_co", "p0"), ("internal_ai_reading", "n0"), ("beta", "n3"), ("g_tv", "p0"), ("w_dcf", "p0"))]},
    {"type": "table", "caption": "命題模組的價格與起始值", "head": ["項目", "單位", "數值"],
     "rowgen": {"sheet": "Thesis", "rows": f"{T['k_long']}-{T['m_co']}", "cols": ["A", "B", "C"], "fmts": ["", "", "n3"],
                "row_fmt": {str(T[k]): {"C": "p1"} for k in ("g0_co", "m_co_ttm", "m_co_1h", "m_co")} | {str(T[k]): {"C": "n2"} for k in ("permw0", "opex0")} | {str(T[k]): {"C": "n1"} for k in ("ai_ttm", "ai_pttm", "ai_ttm_e")}}},
    {"type": "table", "caption": "命題模組：AI 容量、Cloud AI、AI 資本支出與折舊（三情境相同）", "head": ["項目", "單位"] + [f"@Thesis!{c}{thh}" for c in "EFGHIJKLM"],
     "rowgen": {"sheet": "Thesis", "rows": th_rows, "cols": ["A", "B", "E", "F", "G", "H", "I", "J", "K", "L", "M"], "fmts": ["", "", "n1", "n1", "n1", "n1", "n1", "n1", "n1", "n1", "n1"], "row_fmt": th_fmt}},
    {"type": "note", "parts": ["@Seg!A2"]}]},
  {"id": "pl", "title": "三情境損益", "intro": "@PL!A2", "blocks": [
    {"type": "pl", "years": f"@PL!G{hA}:M{hA}", "metrics": [
       met("合併營收（十億美元）", "rev", "n0", True), met("Google Search & other（十億美元）", "search", "n0", True),
       met("Google Cloud（十億美元）", "cloud", "n0"), met("合併營業利益率", "om", "p1", True, ymin="auto"),
       met("Google Services 營業利益率", "om_gs", "p1"), met("EPS（GAAP 稀釋，美元）", "eps", "n2", True, ymin="auto"),
       {"name": "自由現金流（十億美元）", "fmt": "n0", "rows": {s: rng("CF", CR[s]["fcf"]) for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 的分部、產品線與合併損益（含實績年度）", "head": ["項目", "@PL!C12", "@PL!D12", "@PL!E12", "@PL!F12", "@PL!G12", "@PL!H12", "@PL!I12", "@PL!J12", "@PL!K12", "@PL!L12", "@PL!M12"],
     "rowgen": {"sheet": "PL", "rows": seg_rows, "cols": ["A", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"], "fmts": [""] + ["n1"] * 11, "row_fmt": seg_fmt}},
    {"type": "table", "caption": "分部對帳（三情境 × 全部年度取絕對值最大者，應為零）", "head": ["項目", "@PL!C5", "@PL!D5", "@PL!E5", "@PL!F5", "@PL!G5", "@PL!H5", "@PL!I5", "@PL!J5", "@PL!K5", "@PL!L5", "@PL!M5"],
     "rowgen": {"sheet": "PL", "rows": f"{rm['SUMROW']}-{rm['SUMROW'] + 3}", "cols": ["A", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"], "fmts": [""] + ["n3"] * 11}},
    {"type": "table", "caption": "收支關卡", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "CF", "rows": f"{gate}-{gate + 3}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"], "row_fmt": {str(gate + 3): {"B": "", "C": "", "D": ""}}}},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 的現金流、股東回饋與融資", "head": ["項目", "@CF!E12", "@CF!F12", "@CF!G12", "@CF!H12", "@CF!I12", "@CF!J12", "@CF!K12", "@CF!L12", "@CF!M12"],
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
    {"type": "kv", "items": [["@Val!A" + str(VR[k]), "@Val!B" + str(VR[k]) + "|" + f] for k, f in (("emkt", "n0"), ("dmkt", "n1"), ("ke", "p2"), ("kdat", "p2"), ("wacc", "p2"), ("ronic", "p2"), ("ai_pv", "n1"), ("ai_tv", "n1"), ("ai_v", "n1"))]},
    {"type": "table", "caption": "DCF 彙總（三情境）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{dsum0}-{dsum1}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"], "row_fmt": rf_s}},
    {"type": "table", "caption": "分部加總（目標時點 NTM EBITDA × 同業倍數）", "head": ["分部／產品線", "情境 A", "情境 B", "情境 C", "倍數", "EV：A", "EV：B", "EV：C"],
     "rowgen": {"sheet": "Val", "rows": f"{sotp0}-{sotp1}", "cols": ["A", "B", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n1", "n1", "n1", "", "n0", "n0", "n0"], "row_fmt": rf_t}},
    {"type": "table", "collapse": True, "summary": "展開：同業倍數中位數", "head": ["同業組", "中位數", "家數"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['mult_ads']}-{VR['mult_adsx']}", "cols": ["A", "B", "C"], "fmts": ["", "x1", "n0"]}}]},
  {"id": "recon", "title": "對帳：共識、公司指引與舊版", "intro": "@Recon!A1", "blocks": [
    {"type": "table", "caption": "@Recon!A" + str(r1), "head": ["項目", "年度", "情境 A", "情境 B", "情境 C", "共識", "B 對共識", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r1 + 2}-{r2 - 2}", "cols": ["A", "B", "C", "D", "E", "F", "G", "H"], "fmts": ["", "", "n1", "n1", "n1", "n1", "sp1", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r2), "head": ["分部／產品線", "FY2026E", "FY2027E", "FY2028E", "共識", "公司指引", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r2 + 2}-{r3 - 2}", "cols": ["A", "B", "C", "D", "E", "F", "G"], "fmts": ["", "n1", "n1", "n1", "", "", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r3), "head": ["項目", "數值", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r3 + 2}-{rend}", "cols": ["A", "B", "C"], "fmts": ["", "n2", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r4), "head": ["項目", "數值", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r4 + 2}-{rend4}", "cols": ["A", "B", "C"], "fmts": ["", "n2", ""]}}]},
  {"id": "sens", "title": "敏感度", "intro": "@Sens!A1", "blocks": [
    {"type": "tornado", "title": "加權目標價相對基準的變動（一次改一個輸入，整本重算）", "sheet": "Sens", "rows": f"4-{se_end}", "label_col": "A", "value_col": "E", "fmt": "sp1"},
    {"type": "table", "collapse": True, "summary": "展開：全部案例", "head": ["驅動", "設定", "方向", "加權目標價", "相對基準"],
     "rowgen": {"sheet": "Sens", "rows": f"4-{se_end}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "", "", "n2", "sp1"]}},
    {"type": "kv", "items": [["@Sens!A" + str(rm["SE_BASE"]), "@Sens!D" + str(rm["SE_BASE"]) + "|n2"], ["@Sens!A" + str(rm["SE_STALE"]), "@Sens!D" + str(rm["SE_STALE"]) + "|n2"]]}]},
  {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"],
                             ["Tokenomics 快照", "@Sources!B" + str(find("Sources", "Tokenomics 快照", 2))]]},
    {"type": "h3", "text": "資料與口徑"},
    {"type": "list", "items": ["@Q!A1", "@Q!A2", "@In!A2", "@Sens!A2"]},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A", "E", "G"], "In": ["B", "H", "I"]}, "empty": "沒有標為未找到的項目。"}]}
 ]
}
# v0.3.7：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）
import os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "company_html"))
from relayout import relayout
_IW = wb["In"]
_keys = {}
_sec = None
for _r in range(4, _IW.max_row + 1):
    _a = str(_IW.cell(_r, 1).value or "")
    if _a and _IW.cell(_r, 2).value in (None, ""): _sec = _a.split(".")[0]; continue
    if _a: _keys.setdefault(_sec, []).append(_a)
_sel = lambda sec, pred=lambda k: True: [k for k in _keys[sec] if pred(k)]
SVC = "搜尋、YouTube 與訂閱（Google Services）"; CLD = "Google Cloud（含 TPU 與 AI 算力）"; OB = "Other Bets（Waymo 等）"; CO = "全公司：持股、資產負債、折現率、目標價規則、稅與現金流"
SPEC = {
 "in_skip": ["H", "J"],
 "in_groups": [
  [SVC, [["成長路徑與搜尋情境軸", ["conv_n", "g0s_q", "gl_search", "gl_yt", "gl_net", "gl_subs", "dA", "mA", "dC", "mC", "inc_m"]],
         ["利益率、非 AI 資本支出與 Alphabet 層級費用", ["m_yt", "m_net", "m_subs", "cx_svc", "abc_rate", "mb_ttm"]],
         ["估值倍數（數位廣告、訂閱影音同業）", _sel("C2", lambda k: k.startswith(("pm_ads", "pm_subs"))) + ["mult_ads_own", "ads_mode"]]]],
  [CLD, [["AI 容量與擴產速度", _sel("D", lambda k: not k.startswith("tpu"))],
         ["TPU", _sel("D", lambda k: k.startswith("tpu")) + ["tpu_2h", "tpu_yr", "tpu_m"]],
         ["每 MW 價格（k）與 Tokenomics 成本", ["kl_px", "ks_px", "long_sh"] + _sel("E", lambda k: (k[:3] in ("he_", "opx", "cit", "dpi") or k.startswith("cap_")) or k in ("gpuhr_h", "cfac"))],
         ["機房、折舊與在建工程", ["selfb", "life_it", "life_fac", "life_na", "phi_0", "da_base", "cip_0", "cip_it"]],
         ["資本支出指引、自用 AI 與租用算力", ["capex_g", "capex_1h", "svc_sh", "ai_mw_ttm", "ai_mw_pttm"] + _sel("E", lambda k: k.startswith("rent_")) + ["internal_ai_reading"]],
         ["Cloud 其他與估值倍數", ["gl_co", "cx_co"] + _sel("C2", lambda k: k.startswith("pm_cloud")) + ["mult_ai"]]]],
  [OB, [["營收、利益與資本支出", _sel("F", lambda k: k.startswith("ob_")) + ["cx_ob"]],
        ["估值", ["waymo_val", "waymo_stake", "ob_val"]]]],
  [CO, [["股價、股數與資產負債（評價日）", _keys["A"]],
        ["持股", _keys["B"]],
        ["折現率", ["rf", "beta_sa", "beta_gf", "blume_a", "beta", "erp", "kd_sp", "kd", "fin_lease", "wacc_adj", "tax_w", "g_tv", "ronic_sp"]],
        ["目標價規則", ["mid", "t_val", "t_q4", "tp_wA", "tp_wB", "tp_wC", "w_dcf", "t_tgt", "w_ntm27", "cons_tp", "dps_tgt", "t_px0", "sbc_add", "tv_alt"]],
        ["損益與現金流", _keys["G"]],
        ["今年以來實績", _keys["G2"]]]]],
 "blocks_to": {"命題模組": CLD, "@Thesis!A1": CLD, "@Seg!A2": SVC},
 "main": {"DCF 彙總": ["明確期現值", "終值現值", "企業價值 EV（主值", "＋淨現金", "DCF 情境價值推到 12 個月後"],
          "分部加總": ["Google Search & other", "YouTube ads", "Cloud AI（EV", "分部加總 EV（主值）", "分部加總每股價值（主值"],
          "1. 合併層": ["營收|FY2027", "營收|FY2028", "營業利益|FY2028", "自由現金流|FY2027"],
          "2. 逐分部": ["Google Search & other", "Google Services", "Google Cloud", "  Cloud AI"]},
}
cfg = relayout(cfg, X, SPEC)

# ================= v0.3.9（2026-11-14）：個股 HTML 版面規格 v1.3 全套＋Coord 三層協調（樣板 NVIDIA v0.3.10 gen_html_json.py；數字與文字全取自 xlsx） =================
V39 = rm["V39"]
PLA, RVA, CRD, RMA = V39["PL"], V39["RV"], V39["CRD"], V39["RM"]


def _v13():
    S = {x["id"]: x for x in cfg["sections"]}
    PLW = wb["PL"]
    r13 = find("README", "v1.3 展示層")
    RM_ = {k: r13 + 1 + i for i, k in enumerate(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"])}
    assert str(wb["README"][f"A{RM_['ver']}"].value).startswith("v0.3.9"), wb["README"][f"A{RM_['ver']}"].value
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
                "bars": [{"name": [f"@PL!A{sx['s_gs']}"], "y": f"@PL!C{sx['s_gs']}:L{sx['s_gs']}", "color": "#6fa8dc"},
                         {"name": [f"@PL!A{sx['s_cl']}"], "y": f"@PL!C{sx['s_cl']}:L{sx['s_cl']}", "color": "#0b7a80"},
                         {"name": [f"@PL!A{sx['s_ob']}"], "y": f"@PL!C{sx['s_ob']}:L{sx['s_ob']}", "color": "#c29a2b"}],
                "lines": [{"name": [f"@PL!A{ix['opm']}"], "y": f"@PL!C{ix['opm']}:L{ix['opm']}", "axis": "r", "color": "#3b3838", "labels": True}],
                "caption": "堆疊＝Google Services、Google Cloud、Other Bets（含避險損益）三個分部；折線＝合併營業利益率（右軸）。FY2026E 起為情境 B（基準）；AI 折舊上升使營業利益率逐年下降。"}
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
    for col, side in (("B", "低"), ("C", "高")):
        big = [r for r in range(r1_, r2_ + 1) if isinstance(SW[f"{col}{r}"].value, (int, float)) and abs(SW[f"{col}{r}"].value) >= 0.1]
        small = [r for r in range(r1_, r2_ + 1) if isinstance(SW[f"{col}{r}"].value, (int, float)) and 0.0005 <= abs(SW[f"{col}{r}"].value) < 0.1]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        assert big, (col, "無達兩位數的案例")
        tors.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
        if small:
            tors.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
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
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP）。FY2026E 淨利率含 2026 年前三季權益證券未實現損益，預測期不估；毛利率逐年下降是 AI 折舊上升。"}
    old = S["pl"]["blocks"]
    seg_old = next(b for b in old if b.get("type") == "table" and str(b.get("summary", "")).startswith("展開：情境 B 的分部"))
    rest = [b for b in old if b is not seg_old]
    for b in rest:
        if b.get("type") == "pl":
            b["collapse"] = True; b["summary"] = "展開：三情境損益（A／B／C 並列，含折線）"
    S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{PLA['note']}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{PLA['note2']}"]}, chk_tbl] + rest)
    S["pl"].pop("intro", None)
    S["pl"]["blocks"].insert(5, {"type": "note", "parts": [f"@PL!A2"], "collapse": True, "summary": "展開：三情境損益頁說明"})
    # ---- 6. 驅動因子：分部營收推導 ----
    seg_old = dict(seg_old, collapse=False, caption="分部營收推導（情境 B；十億美元）：四條產品線＋Cloud AI／TPU／其他＝三個分部，加 Other Bets 與避險損益＝合併營收", summary=None)
    seg_old.pop("summary")
    pbk = PR["B"]
    segs = [("search", "#6fa8dc"), ("yt", "#cc4125"), ("net", "#999999"), ("subs", "#e69138"), ("ai", "#0b7a80"), ("tpu", "#6fc7cb"), ("co", "#9fc5e8")]
    xr = f"@PL!G{hA}:M{hA}"
    seg_chart = {"type": "theme", "title": "分部營收堆疊（情境 B，十億美元）：Google Services 四條產品線與 Google Cloud 三部分", "x": xr, "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.2,
                 "total": f"@PL!G{pbk['rev']}:M{pbk['rev']}",
                 "bars": [{"name": [f"@PL!A{pbk[k]}"], "y": f"@PL!G{pbk[k]}:M{pbk[k]}", "color": c} for k, c in segs],
                 "caption": "堆疊＝Search & other、YouTube ads、Google Network、訂閱與裝置（Google Services）與 Cloud AI、TPU、Cloud 其他（Google Cloud）；Other Bets 與避險損益金額小，未畫；頂端數字＝合併營收。"}
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM_['drv']}"]}, seg_chart, seg_old] + S["drivers"]["blocks"]
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    # ---- 7. 對帳：三層協調 ----
    CW = wb["Coord"]
    cr = lambda p_, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p_))
    f1h, f1b, f2h = cr("F1. 本模型讀取"), cr("F1b. 追公式"), cr("F2. C0（實績）")
    F1a, F1e = CRD["F1"]; r_tr = CRD["rtr"]
    rd1, rd2, rj, rjr, rg, rl = CRD["rd1"], CRD["rd2"], CRD["rj"], CRD["rjr"], CRD["rg"], CRD["rl"]
    rc1, rc2, rcu, rtp, rmw = CRD["rc1"], CRD["rc2"], CRD["rcu"], CRD["rtp"], CRD["rmw"]
    ra = CRD["ra"]; rq = CRD["rq"]; F2a, F2b = CRD["F2"]
    capx = {"type": "table", "caption": "資本支出（公司已公布實績 vs Wattonomics 隱含值；US$B）", "head": ["讀數", "期間／口徑", "公司實績", "層級隱含", "差距", "狀態／判定"],
            "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in (rd1, rd2, rg)]
                    + [[f"@Coord!A{rj}", f"@Coord!B{rj}", "", "", "", f"@Coord!G{rj}"], [f"@Coord!A{rjr}", "", "", "", "", f"@Coord!G{rjr}"]]}
    nocmp = {"type": "table", "caption": "沒有可比公司實績的項目（參考；不判定）", "head": ["項目", "期間／口徑", "公司實績", "層級隱含", "層級占公司實績", "狀態"],
             "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", f"@Coord!F{r}|p0", f"@Coord!G{r}"] for r in (rc1, rc2)]
                     + [[f"@Coord!A{r}", f"@Coord!B{r}", "", "", "", f"@Coord!G{r}"] for r in (rcu, rtp, rmw)]}
    d3 = {"type": "table", "caption": "展開：實績對帳的輸入、口徑、出處與證據等級", "collapse": True, "summary": "展開：實績對帳的輸入、出處與證據等級",
          "head": ["項目", "期間／口徑", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級"],
          "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!H{r}", f"@Coord!I{r}", f"@Coord!J{r}", f"@Coord!K{r}"] for r in list(rq) + [rd1, rd2, rg, rl] + list(ra) + [rc1, rc2, rtp, rmw]]}
    trc = {"type": "table", "caption": f"@Coord!A{f1b}", "head": ["輸出名稱 ← 公式鏈", "同源判定", "讀到什麼"],
           "rows": [[f"@Coord!A{r}", f"@Coord!G{r}", f"@Coord!H{r}"] for r in range(f1b + 2, r_tr)]}
    t1 = {"type": "table", "caption": f"@Coord!A{f1h}", "collapse": True, "summary": "展開：本模型讀取的 Tokenomics 名稱與 Wattonomics 對照名稱、值與版本",
          "head": ["名稱", "年度", "層級值", "本模型採用值", "單位", "層級與版本"],
          "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}|n3", f"@Coord!D{r}|n3", f"@Coord!E{r}", f"@Coord!F{r}"] for r in range(F1a + 1, F1e + 1)]}
    d0, d1 = CRD["rdump"]
    dump_t = {"type": "table", "caption": "展開：追公式傾印（腳本輸出：每格公式與值）", "collapse": True, "summary": "展開：追公式傾印（每格公式與值；腳本 trace_wattonomics.py 的輸出）",
              "head": ["儲存格（縮排＝依賴深度）", "值", "標記", "公式"],
              "rows": [[f"@Coord!A{r}", f"@Coord!D{r}|n3", f"@Coord!G{r}", f"@Coord!H{r}"] for r in range(d0, d1 + 1)]}
    S["recon"]["blocks"].extend([
        {"type": "h3", "text": "三層協調：讀取層級與實績對帳"},
        {"type": "note", "parts": [f"@Coord!D{CRD['rs']}"]}, capx, nocmp, d3,
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

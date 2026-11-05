# 由 build_spacex.py 的列號對照（ROWMAP）產生 spacex/html.json（列號隨模型變動，勿手改 html.json）。
# 用法：python3 spacex/gen_html_json.py <成品 xlsx> <rowmap.json> spacex/html.json
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
PB = PR["B"]
hA = PR["A"]["lsv"] - 1
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
for k in ("tier", "rec", "meth"): rowfmt_conc[str(VR[k])] = {"B": ""}
for k in ("tp_r2", "tp_r1"): rowfmt_conc[str(VR[k])] = {"B": "n2", "C": "sp1", "D": ""}
for k in ("tp_dcf12", "tp_sotp12"): rowfmt_conc[str(VR[k])] = {"B": "n2", "C": "n2"}
rowfmt_conc[str(VR["s2"])] = {"B": "n2", "C": "n2", "D": "n2"}
rowfmt_conc[str(VR["gate"])] = {"B": "n1", "C": "n1", "D": "n1"}
rowfmt_conc[str(VR["wsum"])] = {"B": "n2"}; rowfmt_conc[str(VR["chk_rd"])] = {"B": "n3"}
dsum0, dsum1 = VR["pvx"], VR["r_v"]
sotp0 = VR["s_cn"]; sotp1 = VR["mkt_imp"]
rf_s = {str(VR["tvsh"]): {"B": "p1", "C": "p1", "D": "p1"}, str(VR["shv"]): {"B": "n3", "C": "n3", "D": "n3"}}
rf_t = {str(VR["ssh"]): {"B": "n3", "C": "n3", "D": "n3"}, str(VR["simp"]): {"B": "x1", "C": "x1", "D": "x1"}, str(VR["mkt_imp"]): {"B": "x1"}}
gate = rm["GATE"]
r1 = find("Recon", "1. 合併層"); r2 = find("Recon", "2. 逐分部"); r3 = find("Recon", "3. S2"); r4 = find("Recon", "4. 每 GW"); r5 = find("Recon", "5. 法說")
r5_end = max(r for r in range(r5, wb["Recon"].max_row + 1) if wb["Recon"].cell(r, 1).value)
se_end = max(x[3] for x in rm["SE_ROWS"])
_se = wb["Sens"]
_tv = [abs(_se.cell(x[3], 5).value) for x in rm["SE_ROWS"] if isinstance(_se.cell(x[3], 5).value, (int, float))]
_thr = max(0.0005, max(_tv) / 5000)   # 長條寬度＝變動 ÷ 最大 × 50%，以兩位小數輸出；太小的案例畫不出寬度，只列在全表
tor_rows = ",".join(str(x[3]) for x in rm["SE_ROWS"] if isinstance(_se.cell(x[3], 5).value, (int, float)) and abs(_se.cell(x[3], 5).value) >= _thr)
thh = rm["TH_HDR"]
th_keys = ("end", "avg_ext", "avg_int", "he_fl", "si_inv", "si_cash", "si_fin", "da_ai", "da_b", "int_opex", "shadow", "int_cost")
th_rows = ",".join(str(T[k]) for k in th_keys)
th_fmt = {str(T["he_fl"]): {c: "n2" for c in "EFGHIJKLM"}, str(T["end"]): {c: "n2" for c in "EFGHIJKLM"}, str(T["avg_ext"]): {c: "n2" for c in "EFGHIJKLM"}, str(T["avg_int"]): {c: "n2" for c in "EFGHIJKLM"}}
SB_ = SR["B"]
sg_keys = ("k", "rev_gw", "s2", "subs", "arpu", "nl")
sg_rows = ",".join(str(SB_[k]) for k in sg_keys)
sg_fmt = {str(SB_["k"]): {c: "n2" for c in "FGHIJKLM"}, str(SB_["s2"]): {c: "n2" for c in "FGHIJKLM"}, str(SB_["rev_gw"]): {c: "n2" for c in "FGHIJKLM"},
          str(SB_["nl"]): {c: "n0" for c in "EFGHIJKLM"}, str(SB_["arpu"]): {c: "n1" for c in "FGHIJKLM"}}
seg_rows = ",".join(str(PB[k]) for k in ("lsv", "ldv", "space", "cons", "eg", "conn", "ads", "rent", "nr", "ssi", "si", "rev", "e_sp", "e_cn", "e_si", "ebitda", "da", "sbc",
                                          "oi_sp", "oi_cn", "oi_si", "oi", "om", "ni", "eps"))
seg_fmt = {str(PB["om"]): {c: "p1" for c in "CDEFGHIJKLM"}, str(PB["eps"]): {c: "n2" for c in "CDEFGHIJKLM"}}
cf_rows = ",".join(str(CR["B"][k]) for k in ("cfo", "cx_sp", "cx_cn", "cx_si", "cx", "fin", "lease", "slb", "fcf_co", "fcf", "spec", "plug", "cend", "dend", "nc"))
INK = {str(wb["In"].cell(r, 1).value): r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}
SCN_H = find("In", "J. 情境驅動對照") + 1
SCN = {k: INK[k] for k in ("scn_k", "scn_n", "scn_w")}
mh = PB["lsv"] - 1
# ─── v0.2.4 版面（Andy 2026-11-05）：參數依業務分「發射／Starlink／xAI 算力／全公司」四塊 ───
IW = wb["In"]
def in_unit_fmt(u):
    u = str(u or "")
    if u == "%": return "p1"
    if u in ("x",): return "n2"
    if u in ("次", "年", "季", "US$/月"): return "n1"
    if u.startswith("US$m/") or u in ("GW", "GW/年", "US$"): return "n2"
    if u in ("十億股",): return "n3"
    return "n1"
def in_row(k):
    r = INK[k]; f = in_unit_fmt(IW.cell(r, 3).value)
    return [f"@In!B{r}", f"@In!D{r}|{f}", f"@In!E{r}|{f}", f"@In!F{r}|{f}", f"@In!C{r}", f"@In!G{r}|tag"]
def has_rng(k):
    r = INK[k]; return IW.cell(r, 4).value not in (None, "") or IW.cell(r, 6).value not in (None, "")
IN_GROUPS = {
 "space": [("發射次數與單價", ["nl_q4", "nl_27", "g_nl", "g_rpl"]), ("發射與開發合約", ["gl_ld"]),
           ("利益率、資本支出與折舊", ["m_spl", "cx_spl", "life_sp", "rem_sp"]), ("估值倍數", ["evs_space"])],
 "conn": [("消費訂戶與 ARPU", ["nadd_q4", "nadd_A", "nadd_B", "nadd_C", "nadd_dec", "arpu0", "g_arpu"]), ("企業與政府", ["g0_eg_adj", "gl_eg"]),
          ("利益率、資本支出與折舊", ["m_cn", "cx_cn", "life_cn", "rem_cn"]), ("EchoStar 頻譜付款", ["spec_2h", "spec_27"]),
          ("估值倍數（同業 EV/EBITDA）", ["pm_conn_irdm", "pm_conn_tmus", "pm_conn_t", "pm_conn_cmcsa", "pm_conn_sats", "mult_cn_adj", "mult_cn_pm"])],
 "si": [("算力規模（名目 GW）與投資速度", ["gw_0", "gw_q2", "gw_ye25", "gw_ye24", "gw_ye26", "cx_q4", "si_cx_rr", "pace", "conv_mw", "mgmt_gw27", "gw27_mgmt", "fin_sh"]),
        ("機隊組成與機房", ["mix0_h", "mix0_b2", "mix0_b3", "am_27", "selfb"]),
        ("每 GW 成本（Tokenomics 快照）", ["he_h", "he_b2", "he_b3", "he_vr", "opx_h", "opx_b2", "opx_b3", "opx_vr", "cit_b3", "cit_vr", "cfac"]),
        ("折舊與在建工程", ["dpi_h", "dpi_b2", "dpi_b3", "life_it", "life_fac", "phi_0", "cip_0", "cip_si", "cip_it", "da_si0", "it_sh_da"]),
        ("對外出租與定價（k＝每 GW 年營收 ÷ 持有成本）", ["ext_sh", "k30_A", "k30_B", "k30_C", "conv_k", "k0_ovr", "k0_cal"]),
        ("Grok／Cursor／X 其他收入", ["cur_pf", "g27_nr", "gl_nr", "gl_ads", "conv_n", "m_so_adj", "sbc_cur_left", "sbc_cur_yrs"]),
        ("自用 AI 與估值倍數", ["internal_ai_reading", "evs_si", "pm_si_crwv", "pm_si_nbis", "pm_si_orcl", "mult_si_adj"])],
 "co": [("股價、股數與資產負債（評價日）", ["px", "px_val", "sh_out", "sh_dil_9m", "dil_perf", "dil_opt", "dil_rsu", "dil_echo", "dil_x", "cash_0", "cash_ye25", "debt_0", "debt_ye25", "slb_0", "fin_lease", "btc", "legal"]),
        ("折現率", ["rf", "beta", "erp", "kd", "tax_w", "wacc_adj", "g_tv", "ronic_sp", "tax_tv"]),
        ("目標價規則", ["mid", "t_val", "t_q4", "tp_wA", "tp_wB", "tp_wC", "w_dcf", "t_tgt", "w_ntm27", "cons_tp", "sbc_add", "tv_alt"]),
        ("稅、利率與股數稀釋", ["tax", "r_cash", "r_debt", "dil", "wc_k", "min_cash"]),
        ("債務與融資租賃到期", ["mat_2H", "mat_27", "mat_28", "mat_29", "mat_30", "mat_31", "mat_32", "lease_2H", "lease_27", "lease_28", "lease_29", "lease_30", "lease_31", "lease_32"])],
}
# 每個 In 參數（A–I 節，不含 I2 實績、J0 對照、J 情境對照）須剛好歸入一組，避免漏列
_sec, _need = None, set()
for _r in range(4, IW.max_row + 1):
    _a = str(IW.cell(_r, 1).value or "")
    if _a and IW.cell(_r, 2).value in (None, ""): _sec = _a.split(".")[0]; continue
    if _a and _sec in ("A", "B", "C", "D", "E", "F", "G", "H", "I"): _need.add(_a)
_got = [k for g in IN_GROUPS.values() for _, ks in g for k in ks]
assert len(_got) == len(set(_got)), "In 參數重複歸組"
assert set(_got) == _need, ("In 參數歸組不完整", sorted(_need - set(_got)), sorted(set(_got) - _need))
IN_HEAD = ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!C3", "@In!G3"]
def in_tables(g, label):
    rng_rows, fix_rows = [], []
    for sub, ks in IN_GROUPS[g]:
        a = [in_row(k) for k in ks if has_rng(k)]; b = [in_row(k) for k in ks if not has_rng(k)]
        if a: rng_rows += [{"sub": sub}] + a
        if b: fix_rows += [{"sub": sub}] + b
    out = []
    if rng_rows:
        out.append({"type": "table", "caption": f"{label}：有敏感度區間的參數（基準進目標價；低／高只用於敏感度頁）", "head": IN_HEAD, "rows": rng_rows})
    if fix_rows:
        out.append({"type": "table", "collapse": True, "summary": f"展開：{label}的其他固定輸入（實績、公司揭露、換算值）", "head": IN_HEAD, "rows": fix_rows})
    return out
# Seg 頁情境 B 區塊：依業務取營運驅動列
SW = wb["Seg"]
segB0 = find("Seg", "情境 B"); segB1 = find("Seg", "情境 C") - 1
def segr(prefix):
    for r in range(segB0, segB1 + 1):   # 先找完全相同的列名（分部合計列），再找開頭相同者
        if str(SW.cell(r, 1).value or "").strip() == prefix: return r
    for r in range(segB0, segB1 + 1):
        if str(SW.cell(r, 1).value or "").strip().startswith(prefix): return r
    raise KeyError(prefix)
def seg_unit_fmt(u):
    u = str(u or "")
    return {"x": "n2", "%": "p1", "次": "n1", "US$B／次": "n3", "US$": "n1", "百萬": "n1"}.get(u, "n1")
SEG_YC = ["E", "F", "G", "H", "I", "J", "K", "L", "M"]
def seg_table(caption, prefixes):
    rows = []
    for p_ in prefixes:
        r = segr(p_); f = seg_unit_fmt(SW.cell(r, 2).value)
        rows.append([f"@Seg!A{r}", f"@Seg!B{r}"] + [f"@Seg!{c}{r}|{f}" for c in SEG_YC])
    return {"type": "table", "caption": caption, "head": ["項目", "單位"] + [f"@Seg!{c}{segB0 + 1}" for c in SEG_YC], "rows": rows}
SEG_SPACE = ["客戶發射次數", "每次客戶發射收入", "Launch Services", "Launch & Development 年增", "Launch & Development", "Space", "Space 資本支出占營收", "資本支出：Space", "新增折舊：Space"]
SEG_CONN = ["Starlink 消費訂戶", "平均消費訂戶", "消費 ARPU", "Consumer", "Enterprise & Government 年增", "Enterprise & Government", "Connectivity", "資本支出：Connectivity", "新增折舊：Connectivity"]
SEG_SI = ["k＝每 GW", "對外每 GW 年營收", "SI 對外雲端服務", "SI 其他收入年增", "Grok／Cursor／X", "SI Solutions", "Advertising 年增", "Advertising", "SI (Super", "S2 核心問題"]
# Thesis A 節（校準與起始值）依項目分組
TW = wb["Thesis"]
def thr(prefix, start):
    for r in range(start, start + 20):
        if str(TW.cell(r, 1).value or "").startswith(prefix): return r
    raise KeyError(prefix)
TA = T["he0"]
CAL_GROUPS = [("每 GW 持有成本與營運費用（評價日機隊）", ["評價日機隊每 GW 年持有成本", "評價日機隊每 GW 年營運費用"]),
              ("每 GW 年營收與 k 校準（最近一季）", ["校準季每 GW 年營收", "校準 k0"]),
              ("對外出租與其他收入（實績）", ["2026Q3 對外雲端服務營收", "9M26 對外雲端服務營收", "2026Q3 SI 其他收入", "SI 其他 EBITDA 率（2026Q3", "SI 其他 EBITDA 率（預測期"]),
              ("每 GW 投資", ["每 GW 投資（2026", "每 GW 投資（FY2027", "對照：PP&E"])]
cal_rows = []
_cal_used = []
for sub, ps in CAL_GROUPS:
    cal_rows.append({"sub": sub})
    for p_ in ps:
        r = thr(p_, TA); _cal_used.append(r)
        f = "p1" if TW.cell(r, 2).value == "%" else "n2"
        cal_rows.append([f"@Thesis!A{r}", f"@Thesis!B{r}", f"@Thesis!C{r}|{f}"])
assert sorted(_cal_used) == list(range(TA, T["ppe_gw"] + 1)), "Thesis A 節校準列分組不完整"
RDW = wb["README"]
rd_show = find("README", "展示層")

cfg = {
 "company": "SpaceX",
 "ticker": "SPCX",
 "wide_text": True,
 "thesis": f"@README!A{rd_show + 1}",
 "question": "FY2028 SI 每 GW 年營收 ÷ Tokenomics 每 GW 年持有成本（與 OpenAI PxQ 頁同口徑）；收支關卡＝FY2026–FY2030 自由現金流合計（扣融資租賃本金與以其他融資取得之不動產設備；Valor 失敗售後租回本金列融資、不扣）",
 "scenarios": [{"key": "A", "name": "情境 A（算力租金維持高檔、Starlink 增長快）"}, {"key": "B", "name": "情境 B（租金向持有成本部分收斂；基準）"}, {"key": "C", "name": "情境 C（租金收斂到持有成本、Starlink 增長放緩）"}],
 "sections": [
  {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [f"@README!A{rd_show + 2}", f"@README!A{rd_show + 3}", f"@README!A{rd_show + 4}"]},
    {"type": "note", "parts": ["@README!A5"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [f"@README!A{rd_show + 5}", f"@README!A{rd_show + 6}", f"@README!A{rd_show + 7}"]}]},
  {"id": "drivers", "title": "驅動因子 → 推導量", "blocks": [
    {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；In 頁 J 節）", "head": [f"@In!{c}{SCN_H}" for c in "BDEFCG"],
     "rows": [[f"@In!B{SCN[k]}", f"@In!D{SCN[k]}|{f}", f"@In!E{SCN[k]}|{f}", f"@In!F{SCN[k]}|{f}", f"@In!C{SCN[k]}", f"@In!G{SCN[k]}|tag"] for k, f in (("scn_k", "n2"), ("scn_n", "n1"), ("scn_w", "p0"))]},
    {"type": "note", "parts": ["@Seg!A2"]},
    {"type": "h3", "text": "一、發射（Space）"}]
    + in_tables("space", "發射")
    + [seg_table("發射：營運驅動與分部營收（情境 B）", SEG_SPACE),
    {"type": "h3", "text": "二、星鏈（Starlink／Connectivity）"}]
    + in_tables("conn", "星鏈")
    + [seg_table("星鏈：營運驅動與分部營收（情境 B）", SEG_CONN),
    {"type": "h3", "text": "三、xAI 算力（SI 分部）"},
    {"type": "note", "parts": ["@Thesis!A1"]}]
    + in_tables("si", "xAI 算力")
    + [{"type": "table", "caption": "xAI 算力：命題模組的校準與起始值（依項目分組）", "head": ["項目", "單位", "數值"], "rows": cal_rows},
    {"type": "table", "caption": "xAI 算力：名目算力、投資、折舊與自用 AI（三情境相同）", "head": ["項目", "單位"] + [f"@Thesis!{c}{thh}" for c in "EFGHIJKLM"],
     "rowgen": {"sheet": "Thesis", "rows": th_rows, "cols": ["A", "B", "E", "F", "G", "H", "I", "J", "K", "L", "M"], "fmts": ["", "", "n1", "n1", "n1", "n1", "n1", "n1", "n1", "n1", "n1"], "row_fmt": th_fmt}},
    seg_table("xAI 算力：k、每 GW 年營收、核心問題比值與 SI 營收（情境 B）", SEG_SI),
    {"type": "h3", "text": "四、全公司：折現率、目標價規則、稅與融資"}]
    + in_tables("co", "全公司")},
  {"id": "pl", "title": "三情境損益", "intro": "@PL!A2", "blocks": [
    {"type": "pl", "years": f"@PL!G{hA}:M{hA}", "metrics": [
       met("合併營收（十億美元）", "rev", "n0", True), met("SI 分部營收（十億美元）", "si", "n0", True),
       met("SI 對外雲端服務（十億美元）", "rent", "n0"), met("Connectivity 分部營收（十億美元）", "conn", "n0"),
       met("合併營業利益率", "om", "p1", True, ymin="auto"), met("EPS（GAAP 稀釋，美元）", "eps", "n2", True, ymin="auto"),
       {"name": "自由現金流（收支關卡口徑，十億美元）", "fmt": "n0", "rows": {s: rng("CF", CR[s]["fcf"]) for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 的分部、子列與合併損益（含實績年度）", "head": ["項目"] + [f"@PL!{c}{mh}" for c in "CDEFGHIJKLM"],
     "rowgen": {"sheet": "PL", "rows": seg_rows, "cols": ["A", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"], "fmts": [""] + ["n1"] * 11, "row_fmt": seg_fmt}},
    {"type": "table", "caption": "分部對帳（三情境 × 全部年度取絕對值最大者，應為零）", "head": ["項目"] + [f"@PL!{c}5" for c in "CDEFGHIJKLM"],
     "rowgen": {"sheet": "PL", "rows": f"{rm['SUMROW']}-{rm['SUMROW'] + 5}", "cols": ["A", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"], "fmts": [""] + ["n3"] * 11}},
    {"type": "table", "caption": "收支關卡", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "CF", "rows": f"{gate}-{gate + 4}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"], "row_fmt": {str(gate + 4): {"B": "", "C": "", "D": ""}}}},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 的現金流與融資", "head": ["項目"] + [f"@CF!{c}{CR['B']['ni'] - 1}" for c in "EFGHIJKLM"],
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
    {"type": "table", "caption": "DCF 彙總與 SI 對外出租模組（三情境）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{dsum0}-{dsum1}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"], "row_fmt": rf_s}},
    {"type": "table", "caption": "分部加總（目標時點 NTM 分部調整後 EBITDA × 同業倍數；Space 為營收 × EV／營收）", "head": ["分部", "情境 A", "情境 B", "情境 C", "倍數", "EV：A", "EV：B", "EV：C"],
     "rowgen": {"sheet": "Val", "rows": f"{sotp0}-{sotp1}", "cols": ["A", "B", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n1", "n1", "n1", "", "n0", "n0", "n0"], "row_fmt": rf_t}},
    {"type": "table", "collapse": True, "summary": "展開：同業倍數中位數", "head": ["同業組", "中位數", "家數"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['mult_conn']}-{VR['mult_si']}", "cols": ["A", "B", "C"], "fmts": ["", "x1", "n0"]}}]},
  {"id": "recon", "title": "對帳：共識、公司數字與口徑", "intro": "@Recon!A1", "blocks": [
    {"type": "table", "caption": "@Recon!A" + str(r1), "head": ["項目", "年度", "情境 A", "情境 B", "情境 C", "過期共識", "B 對共識", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r1 + 2}-{r2 - 3}", "cols": ["A", "B", "C", "D", "E", "F", "G", "H"], "fmts": ["", "", "n1", "n1", "n1", "n1", "sp1", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r2), "head": ["分部／項目", "FY2026E", "FY2027E", "FY2028E", "共識", "公司指引或原文", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r2 + 2}-{r3 - 2}", "cols": ["A", "B", "C", "D", "E", "F", "G"], "fmts": ["", "n1", "n1", "n1", "", "", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r3), "head": ["項目", "情境 A", "情境 B", "情境 C", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r3 + 2}-{r4 - 2}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r4), "head": ["項目", "數值", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r4 + 1}-{r4 + 4}", "cols": ["A", "B", "E"], "fmts": ["", "n1", ""]}},
    {"type": "table", "caption": "@Recon!A" + str(r5), "head": ["項目", "情境 A", "情境 B", "情境 C", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{r5 + 2}-{r5_end}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""],
                "row_fmt": {str(r): {"B": "", "C": "", "D": ""} for r in range(r5 + 2, r5_end + 1) if isinstance(wb["Recon"].cell(r, 2).value, str)}
                           | {str(r): {"B": "sp1", "C": "sp1", "D": "sp1"} for r in range(r5 + 2, r5_end + 1) if "÷ 換算中值 − 1" in str(wb["Recon"].cell(r, 1).value)}}}]},
  {"id": "sens", "title": "敏感度", "intro": "@Sens!A1", "blocks": [
    {"type": "tornado", "title": "加權目標價相對基準的變動（一次改一個輸入，整本重算；變動極小的案例只列在下方全表）", "sheet": "Sens", "rows": tor_rows, "label_col": "A", "value_col": "E", "fmt": "sp1"},
    {"type": "table", "collapse": True, "summary": "展開：全部案例", "head": ["驅動", "設定", "方向", "加權目標價", "相對基準"],
     "rowgen": {"sheet": "Sens", "rows": f"4-{se_end}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "", "", "n2", "sp1"]}},
    {"type": "kv", "items": [["@Sens!A" + str(rm["SE_BASE"]), "@Sens!D" + str(rm["SE_BASE"]) + "|n2"], ["@Sens!A" + str(rm["SE_STALE"]), "@Sens!D" + str(rm["SE_STALE"]) + "|n2"]]}]},
  {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"],
                             ["Tokenomics 快照", "@Sources!B" + str(find("Sources", "Tokenomics 快照", 2))]]},
    {"type": "h3", "text": "資料與口徑"},
    {"type": "list", "items": ["@Q!A1", "@Q!A2", "@In!A2", "@CF!A2", "@Sens!A2"]},
    {"type": "h3", "text": "附錄：驅動 → 推導對照表（查核用）"},
    {"type": "table", "collapse": True, "summary": "展開：每頁的輸入、推導量與因果方向", "head": ["頁", "輸入", "推導量", "因果方向"],
     "rowgen": {"sheet": "README", "rows": f"{find('README', '頁') + 1}-{find('README', 'Sens', 1, find('README', '頁'))}", "cols": ["A", "B", "C", "D"]}},
    {"type": "h3", "text": "附錄：版本要點"},
    {"type": "table", "collapse": True, "summary": "展開：本版與前版的修改要點", "head": ["要點"],
     "rowgen": {"sheet": "README", "rows": f"{find('README', '本版要點') + 1}-{find('README', '工作表：') - 2}", "cols": ["A"]}},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A", "E", "G"], "In": ["B", "H", "I"]}, "empty": "沒有標為未找到的項目。"}]}
 ]
}
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

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
r6 = find("Recon", "6. k 分解")   # v0.2.6
r5_end = max(r for r in range(r5, r6) if wb["Recon"].cell(r, 1).value)
r7 = find("Recon", "7. 同業可比性")   # v0.2.7
r6_end = max(r for r in range(r6, r7) if wb["Recon"].cell(r, 1).value)
r7_end = max(r for r in range(r7, wb["Recon"].max_row + 1) if wb["Recon"].cell(r, 1).value)
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
 "co": [("股價、股數與資產負債（評價日）", ["px", "px_val", "sh_out", "sh_dil_9m", "dil_perf", "dil_opt", "dil_rsu", "dil_echo", "opt_cla", "opt_clb", "opt_cur", "rsu_cur", "optk_cla", "optk_clb", "dil_cur_net", "dil_inc", "dil_base", "dil_x", "cash_0", "cash_ye25", "debt_0", "debt_ye25", "slb_0", "fin_lease", "btc", "legal"]),
        ("折現率", ["rf", "beta", "erp", "kd", "tax_w", "wacc_adj", "g_tv", "ronic_sp", "tax_tv"]),
        ("目標價規則", ["mid", "t_val", "t_q4", "tp_wA", "tp_wB", "tp_wC", "w_dcf", "t_tgt", "w_ntm27", "cons_tp", "sbc_add", "tv_alt"]),
        ("稅、利率與股數稀釋", ["tax", "r_cash", "r_debt", "dil", "wc_k", "min_cash", "lev_max", "icr_min"]),
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
rd_ans = find("README", "命題回覆")
rd_inv = find("README", "主要失效條件")   # v0.2.6
rd_risk = find("README", "風險登錄")
rd_stat = find("README", "模型狀態（README")

# ─── 估值、對帳、敏感度：主要項目直接列，其餘摺疊（Andy 2026-11-05 第二版） ───
def vrows(keys, cols, fmts, rf=None):
    rf = rf or {}
    out = []
    for k in keys:
        r = VR[k]; ov = rf.get(str(r), {})
        out.append([f"@Val!{c}{r}|{ov.get(c, f)}" for c, f in zip(cols, fmts)])
    return out
CONC_MAIN = ["tp", "tpdcf", "tpcomp", "px", "up", "upc", "rec", "scen", "cons"]
CONC_ALL = [k for k in VR if VR[k] and conc0 <= VR[k] <= conc1]
CONC_ALL = sorted(set(CONC_ALL), key=lambda k: VR[k])
CONC_REST = [k for k in CONC_ALL if k not in CONC_MAIN]
CONC_COLS, CONC_FMT = ["A", "B", "C", "D", "E"], ["", "n2", "n2", "n2", ""]
DSUM_MAIN = ["pvx", "pvtv", "ev", "nc0", "dcf12", "r_v"]
DSUM_ALL = sorted({k for k in VR if dsum0 <= VR[k] <= dsum1}, key=lambda k: VR[k])
DSUM_REST = [k for k in DSUM_ALL if k not in DSUM_MAIN]
SOTP_MAIN = ["s_cn", "s_so", "s_sp", "s_rent", "sotp"]
SOTP_ALL = sorted({k for k in VR if sotp0 <= VR[k] <= sotp1}, key=lambda k: VR[k])
SOTP_REST = [k for k in SOTP_ALL if k not in SOTP_MAIN]
for _a, _m in ((CONC_ALL, CONC_MAIN), (DSUM_ALL, DSUM_MAIN), (SOTP_ALL, SOTP_MAIN)):
    assert set(_m) <= set(_a), set(_m) - set(_a)
RW = wb["Recon"]
def R6_FMT(r):   # v0.2.6：第 6 節逐列格式（文字列不套數字格式）
    v = [RW.cell(r, c).value for c in (2, 3, 4)]
    if any(isinstance(x, str) and not x.startswith("=") for x in v if x is not None): return ""
    lab = str(RW.cell(r, 1).value)
    return "n2" if any(k in lab for k in ("目標價", "EBITDA", "影子效益")) and "倍" not in lab else "n2"
def rrow(start, end, label, b=None):
    for r in range(start, end + 1):
        if str(RW.cell(r, 1).value or "").strip().startswith(label.strip()) and (b is None or RW.cell(r, 2).value == b): return r
    raise KeyError((label, b))
def rec_tbl(rows_main, rows_all, cols, fmts, head, caption, summary, rf=None):
    rf = rf or {}
    mk = lambda rr: [[f"@Recon!{c}{r}|{rf.get(str(r), {}).get(c, f)}" for c, f in zip(cols, fmts)] for r in rr]
    rest = [r for r in rows_all if r not in rows_main]
    out = [{"type": "table", "caption": caption, "head": head, "rows": mk(rows_main)}]
    if rest: out.append({"type": "table", "collapse": True, "summary": summary, "head": head, "rows": mk(rest)})
    return out
R1_ALL = [r for r in range(r1 + 2, r2 - 1) if RW.cell(r, 1).value]
R1_MAIN = [rrow(r1, r2, "營收", "FY2027"), rrow(r1, r2, "營收", "FY2028"), rrow(r1, r2, "淨利", "FY2028"), rrow(r1, r2, "公司口徑 FCF", "FY2027")]
R1_TP = rrow(r1, r2, "市場共識目標價")
R1_ALL = [r for r in R1_ALL if r != R1_TP]
R2_ALL = [r for r in range(r2 + 2, r3 - 1) if RW.cell(r, 1).value]
R2_MAIN = [rrow(r2, r3, "Space"), rrow(r2, r3, "Connectivity"), rrow(r2, r3, "SI (Super")]
R3_ALL = [r for r in range(r3 + 2, r4 - 1) if RW.cell(r, 1).value]
R3_MAIN = [rrow(r3, r4, "FY2028 對外每 GW"), rrow(r3, r4, "FY2028 在役機隊"), rrow(r3, r4, "比值（讀法 A")]
R4_ALL = [r for r in range(r4 + 1, r4 + 5) if RW.cell(r, 1).value]
R4_MAIN = R4_ALL[:2]
R5_ALL = [r for r in range(r5 + 2, r5_end + 1) if RW.cell(r, 1).value]
R5_MAIN = [rrow(r5, r5_end, "  換算 2026Q4 營收：中值"), rrow(r5, r5_end, "本模型 2026Q4 合併營收"), rrow(r5, r5_end, "  本模型 ÷ 換算中值"),
           rrow(r5, r5_end, "2027Q1 合併營收指引"), rrow(r5, r5_end, "本模型 2027Q1 合併營收"), rrow(r5, r5_end, "  是否落在指引區間"), rrow(r5, r5_end, "G3：FY2027 年底")]
R5_RF = ({str(r): {"B": "", "C": "", "D": ""} for r in R5_ALL if isinstance(RW.cell(r, 2).value, str)}
         | {str(r): {"B": "sp1", "C": "sp1", "D": "sp1"} for r in R5_ALL if "÷ 換算中值 − 1" in str(RW.cell(r, 1).value)})
_tv_all = [(x[3], _se.cell(x[3], 5).value) for x in rm["SE_ROWS"] if isinstance(_se.cell(x[3], 5).value, (int, float))]
tor_big = ",".join(str(r) for r, v in _tv_all if abs(v) >= 0.1)
tor_small = ",".join(str(r) for r, v in _tv_all if _thr <= abs(v) < 0.1)

cfg = {
 "company": "SpaceX",
 "ticker": "SPCX",
 "wide_text": True,
 "hero": {"rec": f"@Val!B{VR['rec']}",
          "tiles": [[f"@Val!A{VR['tp']}", f"@Val!B{VR['tp']}|n2", ["上檔空間 ", f"@Val!B{VR['up']}|sp1"]],
                    [f"@Val!A{VR['tpcomp']}", f"@Val!B{VR['tpcomp']}|n2", ["上檔空間 ", f"@Val!B{VR['upc']}|sp1"]],
                    [f"@Val!A{VR['px']}", f"@Val!B{VR['px']}|n2", ""]],
          "note": [f"@Val!B{VR['rec_cond']}", "。投資建議依加權目標價判定：", f"@Val!E{VR['rec']}", "。模型狀態：", "@Sens!D" + str(rm["SE_STATUS"]), "（", "@Sens!D" + str(rm["SE_BID"]), "）"]},
 "thesis": f"@README!A{rd_show + 1}",
 "question": "FY2028 SI 每 GW 年營收 ÷ Tokenomics 每 GW 年持有成本（與 OpenAI PxQ 頁同口徑）；收支關卡＝FY2026–FY2030 自由現金流合計（扣融資租賃本金與以其他融資取得之不動產設備；Valor 失敗售後租回本金列融資、不扣）",
 "scenarios": [{"key": "A", "name": "情境 A（算力租金維持高檔、Starlink 增長快）"}, {"key": "B", "name": "情境 B（租金向持有成本部分收斂；基準）"}, {"key": "C", "name": "情境 C（租金收斂到持有成本、Starlink 增長放緩）"}],
 "sections": [
  {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [f"@README!A{rd_show + 2}", f"@README!A{rd_show + 3}", f"@README!A{rd_show + 4}"]},
    {"type": "list", "ordered": True, "collapse": True, "summary": "展開：命題回覆（依本模型計算的答案；想先看結論可展開，或讀完全文再看）",
     "items": [f"@README!A{rd_ans + 1}", f"@README!A{rd_ans + 2}", f"@README!A{rd_ans + 3}"]},
    {"type": "list", "collapse": True, "summary": "展開：主要失效條件（出現任一可觀察證據，結論須重估）",
     "items": [f"@README!A{rd_inv + 1}", f"@README!A{rd_inv + 2}", f"@README!A{rd_inv + 3}"]},
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
    {"type": "bars", "title": "目標價與現價", "fmt": "n0", "items": [
       {"label": "@Val!A" + str(VR["tpdcf"]), "value": "@Val!B" + str(VR["tpdcf"]), "series": 0},
       {"label": "@Val!A" + str(VR["tpcomp"]), "value": "@Val!B" + str(VR["tpcomp"]), "series": 1},
       {"label": "@Val!A" + str(VR["tp"]), "value": "@Val!B" + str(VR["tp"]), "series": 2},
       {"label": "@Val!A" + str(VR["tp_r1"]), "value": "@Val!B" + str(VR["tp_r1"]), "series": 1},
       {"label": "@Val!A" + str(VR["cons"]), "value": "@Val!B" + str(VR["cons"]), "series": 0}],
     "ref": {"label": "@Val!A" + str(VR["px"]), "value": "@Val!B" + str(VR["px"])}},
    {"type": "table", "caption": "結論（主要項目）", "head": ["項目", "數值", "對照", "對照", "說明"], "rows": vrows(CONC_MAIN, CONC_COLS, CONC_FMT, rowfmt_conc)},
    {"type": "table", "collapse": True, "summary": "展開：結論的其他項目（價格帶、讀法一／二、核心問題比值、收支關卡、檢查列）", "head": ["項目", "數值", "對照", "對照", "說明"],
     "rows": vrows(CONC_REST, CONC_COLS, CONC_FMT, rowfmt_conc)},
    {"type": "table", "caption": "DCF 折現率（主要項目）", "head": ["項目", "數值"],
     "rows": [["@Val!A" + str(VR[k]), "@Val!B" + str(VR[k]) + "|p2"] for k in ("wacc", "ke", "kdat", "ronic")]
             + [[f"@In!B{INK['g_tv']}", f"@In!E{INK['g_tv']}|p1"]]},
    {"type": "table", "collapse": True, "summary": "展開：折現率的組成（股權市值、負債、無風險利率、β、股權風險溢酬、稅前債務成本）", "head": ["項目", "數值"],
     "rows": [["@Val!A" + str(VR["emkt"]), "@Val!B" + str(VR["emkt"]) + "|n0"], ["@Val!A" + str(VR["dmkt"]), "@Val!B" + str(VR["dmkt"]) + "|n1"]]
             + [[f"@In!B{INK[k]}", f"@In!E{INK[k]}|{f}"] for k, f in (("rf", "p2"), ("beta", "n2"), ("erp", "p1"), ("kd", "p2"), ("tax_w", "p0"))]},
    {"type": "table", "caption": "DCF 彙總與 SI 對外出租模組（三情境；主要項目）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": vrows(DSUM_MAIN, ["A", "B", "C", "D"], ["", "n1", "n1", "n1"], rf_s)},
    {"type": "table", "collapse": True, "summary": "展開：DCF 彙總的其他項目（讀法一／二、自用 AI 影子效益、股數等）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": vrows(DSUM_REST, ["A", "B", "C", "D"], ["", "n1", "n1", "n1"], rf_s)},
    {"type": "table", "caption": "分部加總（主要項目；目標時點 NTM 分部調整後 EBITDA × 同業倍數，Space 為營收 × EV／營收）", "head": ["分部", "情境 A", "情境 B", "情境 C", "倍數", "EV：A", "EV：B", "EV：C"],
     "rows": vrows(SOTP_MAIN, ["A", "B", "C", "D", "E", "F", "G", "H"], ["", "n1", "n1", "n1", "", "n0", "n0", "n0"], rf_t)},
    {"type": "table", "collapse": True, "summary": "展開：分部加總的其他項目（EV 小計、淨現金、股數、隱含倍數）", "head": ["分部", "情境 A", "情境 B", "情境 C", "倍數", "EV：A", "EV：B", "EV：C"],
     "rows": vrows(SOTP_REST, ["A", "B", "C", "D", "E", "F", "G", "H"], ["", "n1", "n1", "n1", "", "n0", "n0", "n0"], rf_t)},
    {"type": "table", "collapse": True, "summary": "展開：同業倍數中位數", "head": ["同業組", "中位數", "家數"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['mult_conn']}-{VR['mult_si']}", "cols": ["A", "B", "C"], "fmts": ["", "x1", "n0"]}}]},
  {"id": "recon", "title": "對帳：共識、公司數字與口徑", "intro": "@Recon!A1", "blocks":
    rec_tbl(R1_MAIN, R1_ALL, ["A", "B", "C", "D", "E", "F", "G", "H"], ["", "", "n1", "n1", "n1", "n1", "sp1", ""],
            ["項目", "年度", "情境 A", "情境 B", "情境 C", "過期共識", "B 對共識", "說明"], "@Recon!A" + str(r1), "展開：合併層其他年度")
    + [{"type": "kv", "items": [[f"@Recon!A{R1_TP}", [f"@Recon!C{R1_TP}|n2", "（本模型 ", f"@Recon!D{R1_TP}|n2", "，差距 ", f"@Recon!E{R1_TP}|sp1", "）"]]]}]
    + rec_tbl(R2_MAIN, R2_ALL, ["A", "B", "C", "D", "E", "F", "G"], ["", "n1", "n1", "n1", "", "", ""],
              ["分部／項目", "FY2026E", "FY2027E", "FY2028E", "共識", "公司指引或原文", "說明"], "@Recon!A" + str(r2), "展開：子列與公司營運數字對照（算力、Backlog、雲端合約、承諾、Cursor）")
    + rec_tbl(R3_MAIN, R3_ALL, ["A", "B", "C", "D", "E"], ["", "n2", "n2", "n2", ""], ["項目", "情境 A", "情境 B", "情境 C", "說明"],
              "@Recon!A" + str(r3), "展開：讀法 B 與最近一季校準值")
    + rec_tbl(R4_MAIN, R4_ALL, ["A", "B", "E"], ["", "n1", ""], ["項目", "數值", "說明"], "@Recon!A" + str(r4), "展開：近四季與最近一季的每 GW 投資")
    + rec_tbl(R5_MAIN, R5_ALL, ["A", "B", "C", "D", "E"], ["", "n2", "n2", "n2", ""], ["項目", "情境 A", "情境 B", "情境 C", "說明"],
              "@Recon!A" + str(r5), "展開：法說指引的換算細節、須為真條件、定價、Backlog、Cursor 對照", R5_RF)
    + [{"type": "table", "caption": "@Recon!A" + str(r6), "head": ["項目", "B", "C", "D", "說明"],
        "rows": [[f"@Recon!A{r}", f"@Recon!B{r}|{R6_FMT(r)}", f"@Recon!C{r}|{R6_FMT(r)}", f"@Recon!D{r}|{R6_FMT(r)}", f"@Recon!E{r}"]
                 if RW.cell(r, 2).value is not None or RW.cell(r, 5).value else [f"@Recon!A{r}", "", "", "", ""]
                 for r in range(r6 + 2, r6_end + 1) if RW.cell(r, 1).value]},
       {"type": "table", "collapse": True, "summary": "展開：同業可比性（同一估值日、NTM 口徑）", "caption": "@Recon!A" + str(r7),
        "head": ["同業", "EV／NTM EBITDA", "營收成長", "EBITDA 率", "淨負債 ÷ EBITDA", "FCF ÷ EBITDA", "說明"],
        "rows": [[f"@Recon!A{r}", f"@Recon!B{r}|x1", f"@Recon!C{r}|p1", f"@Recon!D{r}|p1", f"@Recon!E{r}|n2", f"@Recon!F{r}|n2", f"@Recon!G{r}"]
                 for r in range(r7 + 2, r7_end + 1) if RW.cell(r, 1).value]}]},
  {"id": "sens", "title": "敏感度", "intro": "@Sens!A1", "blocks": [
    {"type": "tornado", "title": "加權目標價相對基準的變動（變動達兩位數百分比的案例；一次改一個輸入，整本重算）", "sheet": "Sens", "rows": tor_big, "label_col": "A", "value_col": "E", "fmt": "sp1"},
    {"type": "tornado", "collapse": True, "summary": "展開：變動未達兩位數百分比的案例", "title": "加權目標價相對基準的變動（未達兩位數百分比）", "sheet": "Sens", "rows": tor_small, "label_col": "A", "value_col": "E", "fmt": "sp1"},
    {"type": "table", "collapse": True, "summary": "展開：全部案例", "head": ["驅動", "設定", "方向", "加權目標價", "相對基準"],
     "rowgen": {"sheet": "Sens", "rows": f"4-{se_end}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "", "", "n2", "sp1"]}},
    {"type": "kv", "items": [["@Sens!A" + str(rm["SE_BASE"]), "@Sens!D" + str(rm["SE_BASE"]) + "|n2"], ["@Sens!A" + str(rm["SE_STALE"]), "@Sens!D" + str(rm["SE_STALE"]) + "|n2"]]}]},
  {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["build_id", "@Sens!D" + str(rm["SE_BID"])], ["模型狀態", "@Sens!D" + str(rm["SE_STATUS"])],
                             ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"],
                             ["Tokenomics 快照", "@Sources!B" + str(find("Sources", "Tokenomics 快照", 2))]]},
    {"type": "h3", "text": "資料與口徑"},
    {"type": "list", "items": ["@Q!A1", "@Q!A2", "@In!A2", "@CF!A2", "@Sens!A2"]},
    {"type": "h3", "text": "附錄：驅動 → 推導對照表（查核用）"},
    {"type": "table", "collapse": True, "summary": "展開：每頁的輸入、推導量與因果方向", "head": ["頁", "輸入", "推導量", "因果方向"],
     "rowgen": {"sheet": "README", "rows": f"{find('README', '頁') + 1}-{find('README', 'Sens', 1, find('README', '頁'))}", "cols": ["A", "B", "C", "D"]}},
    {"type": "h3", "text": "附錄：版本要點"},
    {"type": "table", "collapse": True, "summary": "展開：本版與前版的修改要點", "head": ["要點"],
     "rowgen": {"sheet": "README", "rows": f"{find('README', '本版要點') + 1}-{find('README', '工作表：') - 2}", "cols": ["A"]}},
    {"type": "h3", "text": "風險登錄（未進主模型的尾部風險）"},
    {"type": "table", "collapse": True, "summary": "展開：風險、是否建模、理由、觀察指標", "head": ["風險", "是否建模", "理由", "觀察指標"],
     "rowgen": {"sheet": "README", "rows": f"{rd_risk + 2}-{rd_stat - 2}", "cols": ["A", "B", "C", "D"]}},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A", "E", "G"], "In": ["B", "H", "I"]}, "empty": "沒有標為未找到的項目。"}]}
 ]
}

def _v13():
    """v0.2.8：個股 HTML 版面規格 v1.3（章節順序、命題圖、估值首圖、敏感度龍捲風、反向 DCF、標準損益表、驅動因子圖、資金模組、三層協調）。"""
    V28 = rm["V28"]; PLA, RVA, CRD, RMA = V28["PL"], V28["RV"], V28["CRD"], V28["RM"]
    S = {x["id"]: x for x in cfg["sections"]}
    PLW = wb["PL"]
    assert str(wb["README"][f"A{RMA['ver']}"].value).startswith("v0.2.9"), wb["README"][f"A{RMA['ver']}"].value
    cfg["layout"] = "v1.3"
    old_note = cfg["hero"]["note"]
    cfg["hero"]["note"] = [f"@README!A{RMA['hero']}"]
    ph = PLA["hdr"]; ix = PLA["idx"]; sx = PLA["sidx"]
    # ---- 1. 命題 ----
    th = S["thesis"]["blocks"]
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RMA[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th_chart = {"type": "theme", "title": "營收（十億美元，依分部堆疊）與營業利益率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}",
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3, "total": f"@PL!C{ix['rev']}:L{ix['rev']}",
                "bars": [{"name": [f"@PL!A{sx['s_sp']}"], "y": f"@PL!C{sx['s_sp']}:L{sx['s_sp']}", "color": "#6fa8dc"},
                         {"name": [f"@PL!A{sx['s_cn']}"], "y": f"@PL!C{sx['s_cn']}:L{sx['s_cn']}", "color": "#0b7a80"},
                         {"name": [f"@PL!A{sx['s_si']}"], "y": f"@PL!C{sx['s_si']}:L{sx['s_si']}", "color": "#c29a2b"}],
                "lines": [{"name": [f"@PL!A{ix['opm']}"], "y": f"@PL!C{ix['opm']}:L{ix['opm']}", "axis": "r", "color": "#3b3838", "labels": True}],
                "caption": "堆疊＝Space、Connectivity、SI（xAI 算力出租與 Grok、Cursor、X）三個分部；折線＝合併營業利益率（右軸，財報口徑）。FY2026E 起為情境 B（基準）；FY2023A–FY2025A 為招股書實績（已追溯重編納入 xAI 與 X）。"}
    th[0:0] = [{"type": "note", "parts": [f"@README!A{RMA['rec']}"]}, th_chart]
    # ---- 2. 估值：四個目標價與現價置頂 ----
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b["type"] == "bars")
    va[i_bar] = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
                 "items": [{"label": f"@Val!A{VR[k]}", "value": f"@Val!B{VR[k]}", "series": sr} for k, sr in (("tpdcf", 1), ("tpcomp", 2), ("tp", 0), ("cons", 1))],
                 "ref": {"label": f"@Val!A{VR['px']}", "value": f"@Val!B{VR['px']}"}}
    bar = va.pop(i_bar)
    S["valuation"]["blocks"] = [bar, {"type": "note", "parts": [S["valuation"].pop("intro")]}, {"type": "note", "parts": list(old_note)}] + va
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
    A_ = RVA; yh = A_["yh"]
    yr_fmt = {"y_rev": "n1", "y_ebit": "n1", "y_m": "p1", "y_df": "n3", "y_tvf": "n3", "y_k": "n3"}
    rev_ = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2", "blocks": [
        {"type": "note", "parts": [f"@RevDCF!A{A_['sum_dcf']}"]}, {"type": "note", "parts": [f"@RevDCF!A{A_['sum_cmp']}"]},
        {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
         "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{A_['m1']}|p1", f"@RevDCF!B{A_['m0']}|p1"],
                  ["營收規模（FY2028 起各年同乘的倍數）", f"@RevDCF!B{A_['s1']}|n2", "—"],
                  ["FY2028 營收（十億美元）", f"@RevDCF!B{A_['rv1']}|n0", f"@RevDCF!B{A_['rv0']}|n0"],
                  ["FY2028 營收年增率", f"@RevDCF!B{A_['g1']}|p1", f"@RevDCF!B{A_['g0']}|p1"],
                  ["折現率（WACC）", f"@RevDCF!B{A_['w1']}|p2", f"@RevDCF!B{A_['w0']}|p2"],
                  ["折現率反解結果", f"@RevDCF!B{A_['w_ok']}", ""]]},
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
               "rows": [rowx(r) for r in range(PLA["chk"] + 1, PLA["chk"] + 8)]}
    rv_, gm_, opm_, nm_ = ix["rev"], ix["gm"], ix["opm"], ix["nm"]
    is_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": [f"@PL!A{rv_}"], "y": f"@PL!C{rv_}:L{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": [f"@PL!A{gm_}"], "y": f"@PL!C{gm_}:L{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": [f"@PL!A{opm_}"], "y": f"@PL!C{opm_}:L{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": [f"@PL!A{nm_}"], "y": f"@PL!C{nm_}:L{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP）。預測期毛利率是營業成本占成本費用合計之比固定的結果（模型未分列營業成本），營業利益率才是模型推導量；FY2023A 含大額減損，FY2026E 起 AI 折舊上升壓低營業利益率。"}
    old = S["pl"]["blocks"]
    seg_old = next(b for b in old if b.get("type") == "table" and str(b.get("summary", "")).startswith("展開：情境 B 的分部"))
    cf_old = next(b for b in old if b.get("type") == "table" and str(b.get("summary", "")).startswith("展開：情境 B 的現金流"))
    gate_old = next(b for b in old if b.get("type") == "table" and b.get("caption") == "收支關卡")
    rest = [b for b in old if b not in (seg_old, cf_old, gate_old)]
    for b in rest:
        if b.get("type") == "pl":
            b["collapse"] = True; b["summary"] = "展開：三情境損益（A／B／C 並列，含折線）"
    S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{PLA['note']}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{PLA['note2']}"]}, chk_tbl,
                                                          {"type": "note", "parts": ["@PL!A2"], "collapse": True, "summary": "展開：三情境損益頁說明"}] + rest + [dict(seg_old)])
    S["pl"].pop("intro", None)
    # ---- 6. 驅動因子：分部營收推導圖 ----
    pbk = PR["B"]
    segs = [("space", "#6fa8dc"), ("cons", "#0b7a80"), ("eg", "#45b0b5"), ("rent", "#c29a2b"), ("nr", "#e69138"), ("ads", "#999999")]
    seg_chart = {"type": "theme", "title": "營收推導堆疊（情境 B，十億美元）：發射、Starlink 消費與企業政府、SI 對外雲端服務、Grok／Cursor／X 與廣告", "x": f"@PL!G{hA}:M{hA}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.2,
                 "total": f"@PL!G{pbk['rev']}:M{pbk['rev']}",
                 "bars": [{"name": [f"@PL!A{pbk[k]}"], "y": f"@PL!G{pbk[k]}:M{pbk[k]}", "color": c} for k, c in segs],
                 "caption": "堆疊＝Space、Consumer、Enterprise & Government、SI 對外雲端服務、Grok／Cursor／X 訂閱與資料授權、Advertising；頂端數字＝合併營收。SI 對外雲端服務是營收規模與目標價最大的驅動。"}
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RMA['drv']}"]}, seg_chart] + S["drivers"]["blocks"]
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    # ---- 6b. 資金模組（由損益章移出：收支關卡與情境 B 現金流）----
    cb = CR["B"]
    fund_chart = {"type": "theme", "title": "情境 B 營運現金流、資本支出與自由現金流（十億美元）：FY2026E–FY2032E", "x": f"@CF!G{cb['ni'] - 1}:M{cb['ni'] - 1}", "fmt": "n0", "fmt_r": "n0", "stack": False, "pad": 1.2,
                  "bars": [{"name": [f"@CF!A{cb['cfo']}"], "y": f"@CF!G{cb['cfo']}:M{cb['cfo']}", "color": "#0b7a80"},
                           {"name": [f"@CF!A{cb['cx']}"], "y": f"@CF!G{cb['cx']}:M{cb['cx']}", "color": "#c29a2b"}],
                  "lines": [{"name": [f"@CF!A{cb['fcf']}"], "y": f"@CF!G{cb['fcf']}:M{cb['fcf']}", "axis": "l", "color": "#3b3838", "labels": True}],
                  "caption": "長條＝營運現金流與現金資本支出；折線＝自由現金流（收支關卡口徑，扣融資租賃本金與以其他融資取得之不動產設備）。資本支出高於營運現金流的缺口由融資補足（見下方現金流與融資表）。"}
    funding = {"id": "funding", "title": "資金：收支關卡與融資", "blocks": [fund_chart, gate_old, cf_old]}
    # ---- 7. 對帳：三層協調 ----
    CW = wb["Coord"]
    cr = lambda p_, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p_))
    f1h, f1b, f2h = cr("F1. 本模型讀取"), cr("F1b. 追公式"), cr("F2. C0（實績）")
    F1a, F1e = CRD["F1"]; r_tr = CRD["rtr"]
    rg1, rg2, rgj = CRD["rg"]; rc1, rc2, rcj, rcjr, rcn = CRD["rc"]; rn1, rn2, rn3, rn4, rn5, rn6 = CRD["rn"]; rl = CRD["rl"]
    gw_t = {"type": "table", "caption": "名目算力（公司揭露 vs Wattonomics 隱含 2025 年底上線 GW；GW）", "head": ["讀數", "期間／口徑", "公司實績", "層級隱含", "差距", "狀態／判定"],
            "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!D{r}|n2", f"@Coord!E{r}|n2", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in (rg1, rg2)] + [[f"@Coord!A{rgj}", f"@Coord!B{rgj}", "", "", "", f"@Coord!G{rgj}"]]
                    + [[f"@Coord!A{CRD['rgr']}", f"@Coord!B{CRD['rgr']}", f"@Coord!D{CRD['rgr']}|n2", f"@Coord!E{CRD['rgr']}|n2", f"@Coord!F{CRD['rgr']}|sp1", f"@Coord!G{CRD['rgr']}"]]}
    capx = {"type": "table", "caption": "資本支出（公司已公布實績 vs Wattonomics 隱含值；US$B；層級值同源，參考）", "head": ["讀數", "期間／口徑", "公司實績", "層級隱含", "差距", "狀態／判定"],
            "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in (rc1, rc2, rcn)]
                    + [[f"@Coord!A{rcj}", f"@Coord!B{rcj}", "", "", "", f"@Coord!G{rcj}"], [f"@Coord!A{rcjr}", "", "", "", "", f"@Coord!G{rcjr}"]]}
    nocmp = {"type": "table", "caption": "沒有可比公司實績的項目與並列讀法（參考；不判定）", "head": ["項目", "期間／口徑", "公司值", "層級隱含", "差距", "狀態"],
             "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!D{r}|n2", f"@Coord!E{r}|n2", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in (rn2, rn4, rn5, rn6)]
                     + [[f"@Coord!A{r}", f"@Coord!B{r}", "", "", "", f"@Coord!G{r}"] for r in (rn1, rn3)]}
    d3 = {"type": "table", "caption": "展開：實績對帳的輸入、出處與證據等級", "collapse": True, "summary": "展開：實績對帳的輸入、出處與證據等級",
          "head": ["項目", "期間／口徑", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級"],
          "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!H{r}", f"@Coord!I{r}", f"@Coord!J{r}", f"@Coord!K{r}"] for r in [rg1, rg2, rgj, CRD["rgr"], rl, rc1, rc2, rcn, rn1, rn2, rn3, rn4, rn5, rn6]]}
    trc = {"type": "table", "caption": f"@Coord!A{f1b}", "head": ["輸出名稱 ← 公式鏈", "同源判定", "讀到什麼"],
           "rows": [[f"@Coord!A{r}", f"@Coord!G{r}", f"@Coord!H{r}"] for r in range(f1b + 2, r_tr)]}
    t1 = {"type": "table", "caption": f"@Coord!A{f1h}", "collapse": True, "summary": "展開：本模型讀取的 Tokenomics 名稱與 Wattonomics 對照名稱、值與版本",
          "head": ["名稱", "年度", "層級值", "本模型採用值", "單位", "層級與版本"],
          "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}|n3", f"@Coord!D{r}|n3", f"@Coord!E{r}", f"@Coord!F{r}"] for r in range(F1a + 1, F1e + 1)]}
    d0, d1 = CRD["rdump"]
    dump_t = {"type": "table", "caption": "展開：追公式傾印（腳本輸出：每格公式與值）", "collapse": True, "summary": "展開：追公式傾印（每格公式與值；腳本 trace_wattonomics.py 的輸出）",
              "head": ["儲存格（縮排＝依賴深度）", "值", "狀態", "標記", "公式"],
              "rows": [[f"@Coord!A{r}", f"@Coord!D{r}|n3", f"@Coord!B{r}", f"@Coord!G{r}", f"@Coord!H{r}"] for r in range(d0, d1 + 1)]}
    S["recon"]["blocks"].extend([
        {"type": "h3", "text": "三層協調：讀取層級與實績對帳"},
        {"type": "note", "parts": [f"@Coord!D{CRD['rs']}"]}, gw_t, {"type": "note", "parts": ["口徑：", f"@Coord!B{CRD['rkq']}"]}, capx, nocmp, d3,
        {"type": "note", "parts": [f"@Coord!B{CRD['rr']}"]}, {"type": "note", "parts": [f"@Coord!B{CRD['rsame']}"]}, {"type": "note", "parts": [f"@Coord!B{CRD['rk']}"]},
        {"type": "h3", "text": "追公式：層級值是否讀公司實績或目標"}, trc, {"type": "note", "parts": [f"@Coord!A{r_tr}"]}, dump_t, t1])
    # ---- 8. 資料狀態：本版要點 ----
    S["status"]["blocks"][0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RMA['ver']}"]}]
    S["revdcf"] = rev_; S["funding"] = funding
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "funding", "recon", "status"]
    cfg["sections"] = [S[k] for k in order]


_v13()
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

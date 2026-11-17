# 由 build_palantir.py 的列號對照（ROWMAP）產生 software_ai/palantir/html.json（列號隨模型變動，勿手改 html.json；版面規格 v1.3）。
# 用法：python3 software_ai/palantir/gen_html_json.py <成品 xlsx> <rowmap.json> software_ai/palantir/html.json
import json, sys, openpyxl
X, RM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
rm = json.load(open(RM, encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)
VR, SEGR, SB, PLR, PLS, CFR, GR, DR, RVR, RCK, IR, RDR = (rm[k] for k in ("VR", "SEGR", "SB", "PLR", "PLS", "CFR", "GR", "DR", "RVR", "RCK", "IR", "RDR"))
PB, SB_ = PLR["B"], SEGR["B"]
hdr = 3   # 年度表頭列（各頁第 3 列）
def rowspec(sh, r, cols, fmt):
    return [f"@{sh}!A{r}"] + [f"@{sh}!{c}{r}|{fmt}" for c in cols]
YC = "CDEFGHIJKL"
def rng(sh, r, c1="C", c2="L"): return f"@{sh}!{c1}{r}:{c2}{r}"
def rows_tbl(sh, keys, regs, cols, fmts, head_row=hdr, caption=None, collapse=False, summary=None):
    t = {"type": "table", "head": [f"@{sh}!A{head_row}"] + [f"@{sh}!{c}{head_row}" for c in cols],
         "rows": [[f"@{sh}!A{regs[k]}"] + [f"@{sh}!{c}{regs[k]}|{fmts.get(k, 'n1')}" for c in cols] for k in keys]}
    if caption: t["caption"] = caption
    if collapse: t["collapse"] = True; t["summary"] = summary
    return t
def val_rows(keys, cols="BCD", fmts=None):
    return [[f"@Val!A{VR[k]}"] + [f"@Val!{c}{VR[k]}|{(fmts or {}).get(k, 'n2')}" for c in cols] for k in keys]
def isnum(v): return isinstance(v, (int, float)) and not isinstance(v, bool)

cfg = {
 "company": "Palantir Technologies", "ticker": "PLTR", "layout": "v1.3", "wide_text": True,
 "thesis": f"@README!A{RDR['thesis']}",
 "question": "AI 讓軟體公司按用量計價的營收中，AI 增量是否大於單價侵蝕與推論成本；現價隱含什麼",
 "scenarios": [{"key": "A", "name": "情境 A（工作負載成長維持、單價侵蝕小）"}, {"key": "B", "name": "情境 B（逐季放緩；基準）"}, {"key": "C", "name": "情境 C（客戶趨飽和、單價侵蝕深）"}],
 "hero": {"rec": f"@Val!B{VR['rec']}",
          "tiles": [[f"@Val!A{VR['tp']}", f"@Val!B{VR['tp']}|n2", ["上檔空間 ", f"@Val!B{VR['up']}|sp1"]],
                    [f"@Val!A{VR['tpdcf']}", f"@Val!B{VR['tpdcf']}|n2", ["上檔空間 ", f"@Val!B{VR['updcf']}|sp1"]],
                    [f"@Val!A{VR['tpcomp']}", f"@Val!B{VR['tpcomp']}|n2", ["上檔空間 ", f"@Val!B{VR['upcomp']}|sp1"]]],
          "note": [f"@README!A{RDR['hero']}"]},
 "sections": []}
S = cfg["sections"]
# ---- thesis ----
th_chart = {"type": "theme", "title": "營收（十億美元，依官方分部堆疊）與營業利益率：FY2023A–FY2032E（情境 B）", "x": rng("PL", hdr), "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
            "total": rng("PL", PB["rev"]),
            "bars": [{"name": [f"@PL!A{PB['gov']}"], "y": rng("PL", PB["gov"]), "color": "#6fa8dc"}, {"name": [f"@PL!A{PB['com']}"], "y": rng("PL", PB["com"]), "color": "#0b7a80"}],
            "lines": [{"name": [f"@PL!A{PLS['opm']}"], "y": rng("PL", PLS["opm"]), "axis": "r", "color": "#3b3838", "labels": True}],
            "caption": "堆疊＝Government 與 Commercial 兩個官方分部；折線＝合併 GAAP 營業利益率（右軸）。FY2026E 起為情境 B（基準）。"}
S.append({"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "note", "parts": [f"@README!A{RDR['rec']}"]}, th_chart,
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [f"@README!A{RDR[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]},
    {"type": "h3", "text": "命題回覆"},
    {"type": "list", "items": [f"@README!A{RDR[k]}" for k in ("qa1", "qa2", "qa3")]}]})
# ---- valuation ----
V_ = VR
val_blocks = [
    {"type": "bars", "title": "目標價與現價（美元；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{V_[k]}", "value": f"@Val!B{V_[k]}", "series": sr} for k, sr in (("tpdcf", 1), ("tpcomp", 2), ("tp", 0), ("cons", 1))],
     "ref": {"label": f"@Val!A{V_['px']}", "value": f"@Val!B{V_['px']}"}},
    {"type": "note", "parts": [f"@Val!A1"]}, {"type": "note", "parts": [f"@README!A{RDR['dt1']}"]}, {"type": "note", "parts": [f"@README!A{RDR['rel1']}"]}, {"type": "note", "parts": [f"@README!A{RDR['rel2']}"]}, {"type": "note", "parts": [f"@README!A{RDR['dt2']}"]},
    {"type": "table", "caption": "結論主要項目", "head": ["項目", "數值", "對照", "對照"],
     "rows": [[f"@Val!A{V_[k]}", f"@Val!B{V_[k]}|{f}", "", ""] for k, f in (("px", "n2"), ("tpdcf", "n2"), ("tpcomp", "n2"), ("tp", "n2"), ("up", "sp1"), ("est", ""), ("gate", "n1"), ("rec", ""))]
              + [[f"@Val!A{V_['bdist']}", f"@Val!B{V_['bdist']}|n1", f"@Val!C{V_['bdist']}|n2", f"@Val!D{V_['bdist']}|n2"]]},
    {"type": "table", "caption": "上檔空間歸因", "head": ["項目", "貢獻"], "rows": [[f"@Val!A{V_[k]}", f"@Val!B{V_[k]}|sp1"] for k in ("att_dcf", "att_cmp", "att_adj")]},
    {"type": "kv", "items": [[f"@Val!A{V_[k]}", f"@Val!B{V_[k]}|{f}"] for k, f in (("ke", "p2"), ("kdat", "p2"), ("wacc", "p2"), ("ronic", "p2"), ("g", "p2"))]},
    {"type": "table", "caption": "DCF 彙總（三情境）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rows": val_rows(["d_pv", "d_pvtv", "d_ev", "d_nc", "d_eq", "d_ps", "d_pstp", "d_tvsh"], fmts={"d_tvsh": "p1"})},
    {"type": "table", "caption": "DCF（股票薪酬加回）對照讀數（只作對照，不進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": val_rows(["d_pstp", "d_pstp_sbc"]) + [[f"@Val!A{V_['tpdcf_sbc']}", f"@Val!B{V_['tpdcf_sbc']}|n2", "", ""], [f"@Val!A{V_['updcf_sbc']}", f"@Val!B{V_['updcf_sbc']}|sp1", "", ""]]},
    {"type": "table", "caption": "Comparable 主腿：同業 P/S × 相對倍數（三情境）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@Val!A{V_['m_c']}", f"@Val!B{V_['m_c']}|n2", "", ""], [f"@Val!A{V_['m_g']}", f"@Val!B{V_['m_g']}|n2", "", ""], [f"@Val!A{V_['pg_c']}", f"@Val!B{V_['pg_c']}|p1", "", ""], [f"@Val!A{V_['pg_g']}", f"@Val!B{V_['pg_g']}|p1", "", ""]]
             + val_rows(["sg_c", "sg_g"], fmts={"sg_c": "p1", "sg_g": "p1"}) + val_rows(["rl_c", "rl_g"]) + val_rows(["rr_c", "rc_c", "rr_g", "rc_g", "ma_c", "ma_g", "ttm_c", "ttm_g", "ev_t", "nc_t", "ps_c", "ps_u", "imp_m"], fmts={"ttm_c": "n0", "ttm_g": "n0", "ev_t": "n0", "nc_t": "n0"})
             + val_rows(["clamp"], fmts={"clamp": ""})},
    {"type": "table", "collapse": True, "summary": "展開：評價日稀釋股數與折現率組成", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{V_[k]}", f"@Val!B{V_[k]}|{f}"] for k, f in (("sh_basic", "n1"), ("sh_rsu", "n1"), ("sh_opt", "n1"), ("sh_sar", "n1"), ("sh_dil", "n1"), ("nc0", "n1"), ("rf", "p2"), ("beta", "n3"), ("erp", "p2"), ("mcap", "n0"), ("we", "p1"))]}]
S.append({"id": "valuation", "title": "估值", "blocks": val_blocks})
# ---- sens ----
SW = wb["Sens"]; r1, r2 = rm["SE_REL1"], rm["SE_REL2"]
blocks = []
for col, side in (("B", "低"), ("C", "高")):
    rows = [r for r in range(r1, r2 + 1) if isnum(SW[f"{col}{r}"].value) and abs(SW[f"{col}{r}"].value) >= 0.0005]
    if rows: blocks.append({"type": "tornado", "title": f"加權目標價相對基準：輸入取{side}值的案例（一次改一個輸入，整本重算）", "sheet": "Sens", "rows": ",".join(map(str, rows)), "label_col": "A", "value_col": col, "fmt": "sp1"})
assert len(blocks) == 2, "Sens 低案與高案都需有變動案例"
se_end = max(x[3] for x in rm["SE_ROWS"])
blocks += [{"type": "note", "parts": ["@Sens!A1"]}, {"type": "note", "parts": ["@Sens!A2"]},
           {"type": "table", "collapse": True, "summary": "展開：全部案例（含 DCF 單法、投資建議與收支關卡 B）", "head": ["驅動", "設定", "方向", "加權目標價", "相對基準", "DCF 目標價", "DCF 變動（美元）", "投資建議", "收支關卡 B"],
            "rowgen": {"sheet": "Sens", "rows": f"4-{se_end}", "cols": ["A", "B", "C", "D", "E", "J", "K", "L", "H"], "fmts": ["", "", "", "n2", "sp1", "n2", "n2", "", "n1"]}},
           {"type": "kv", "items": [[f"@Sens!A{rm['SE_BASE']}", f"@Sens!D{rm['SE_BASE']}|n2"], [f"@Sens!A{rm['SE_STALE']}", f"@Sens!D{rm['SE_STALE']}|n2"], ["模型狀態", f"@Sens!D{rm['SE_STATUS']}"], ["build_id", f"@Sens!D{rm['SE_BID']}"]]}]
S.append({"id": "sens", "title": "敏感度", "blocks": blocks})
# ---- revdcf ----
A_ = RVR
S.append({"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2", "blocks": [
    {"type": "note", "parts": [f"@RevDCF!A{A_['sum_dcf']}"]}, {"type": "note", "parts": [f"@RevDCF!A{A_['sum_cmp']}"]},
    {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含與模型", "head": ["項目", "值"],
     "rows": [[f"@RevDCF!A{A_[k]}", f"@RevDCF!B{A_[k]}|{f}"] for k, f in (("px", "n2"), ("mkt_ev", "n0"), ("mod_ev", "n0"), ("scale", "n2"), ("rv28_0", "n0"), ("rv28_1", "n0"), ("cagr0", "p1"), ("cagr1", "p1"), ("w0", "p2"), ("w1", "p2"))]},
    {"type": "table", "caption": "Comparable：本公司現行倍數與同業", "head": ["項目", "值"],
     "rows": [[f"@RevDCF!A{A_[k]}", f"@RevDCF!B{A_[k]}|{f}"] for k, f in (("cur_mult", "n1"), ("peer_mult", "n1"), ("cur_rel", "n1"), ("imp_up", "sp1"))]}]})
# ---- pl ----
pl_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": rng("PL", hdr), "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
            "bars": [{"name": [f"@PL!A{PLS['rev']}"], "y": rng("PL", PLS["rev"]), "color": "#6fc7cb"}],
            "lines": [{"name": [f"@PL!A{PLS['gm']}"], "y": rng("PL", PLS["gm"]), "axis": "r", "color": "#0b7a80", "labels": True},
                      {"name": [f"@PL!A{PLS['opm']}"], "y": rng("PL", PLS["opm"]), "axis": "r", "color": "#c29a2b"},
                      {"name": [f"@PL!A{PLS['nm']}"], "y": rng("PL", PLS["nm"]), "axis": "r", "color": "#6f6bae", "dash": True}],
            "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP）。預測期營業成本與營業費用的拆分是殘差，營業利益不受拆分影響。"}
std_keys = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pbt", "tax", "ni", "nm", "eps", "r_niy"]
std_fmt = {"yoy": "p1", "gm": "p1", "opm": "p1", "nm": "p1", "r_niy": "p1", "eps": "n2"}
cols_all = "CDEFGHIJKLM"
is_tbl = {"type": "table", "caption": "標準損益表（GAAP；百萬美元，EPS 為美元）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B（基準）、最右欄為 2026 年前三季累計實績（Non-GAAP 調整不入本表）",
          "head": [f"@PL!A{PLS['hd']}"] + [f"@PL!{c}{PLS['hd']}" for c in cols_all],
          "rows": [[f"@PL!A{PLS[k]}"] + [f"@PL!{c}{PLS[k]}|{std_fmt.get(k, 'n1')}" for c in cols_all] for k in std_keys]}
chk_tbl = {"type": "table", "caption": "核對（應為零）", "collapse": True, "summary": "展開：標準損益表核對列", "head": [f"@PL!A{PLS['hd']}"] + [f"@PL!{c}{PLS['hd']}" for c in cols_all],
           "rows": [[f"@PL!A{PLS[k]}"] + [f"@PL!{c}{PLS[k]}|n3" for c in cols_all] for k in ("chk1", "chk2")]}
plb = {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線）", "years": rng("PL", hdr, "C", "L"),
       "metrics": [{"name": nm, "fmt": f, "rows": {s: rng("PL", PLR[s][k]) for s in "ABC"}} | ({"chart": True} if ch else {}) for nm, k, f, ch in (("合併營收（百萬美元）", "rev", "n0", True), ("GAAP 營業利益（百萬美元）", "oi", "n0", True), ("EPS（GAAP 稀釋，美元）", "eps", "n2", False))]}
S.append({"id": "pl", "title": "損益（標準損益表）", "blocks": [{"type": "note", "parts": ["@PL!A2"]}, pl_chart, is_tbl, chk_tbl, plb]})
# ---- drivers ----
Sg = SEGR
ai_chart = {"type": "theme", "title": "AI 淨效果的分量：增量、單價侵蝕、增量推論成本（長條）與淨效果（折線）：FY2027E–FY2032E（情境 B；百萬美元）", "x": rng("Seg", hdr, "G", "L"), "fmt": "n0", "stack": False, "pad": 1.2,
            "bars": [{"name": [f"@Seg!A{SB_['inc']}"], "y": rng("Seg", SB_["inc"], "G", "L"), "color": "#0b7a80"}, {"name": [f"@Seg!A{SB_['ero']}"], "y": rng("Seg", SB_["ero"], "G", "L"), "color": "#b4472f"},
                     {"name": [f"@Seg!A{SB_['incinf']}"], "y": rng("Seg", SB_["incinf"], "G", "L"), "color": "#c29a2b"}],
            "lines": [{"name": [f"@Seg!A{SB_['net']}"], "y": rng("Seg", SB_["net"], "G", "L"), "color": "#3b3838", "labels": True}],
            "caption": "增量＝客戶數×前一年每客戶營收×工作負載成長；侵蝕＝同底數×單價變動（負值）；增量推論成本相對 FY2026 隱含占比（只讀 Tokenomics 快照價格）。"}
scn_chart = {"type": "theme", "title": "AI 淨效果 × 情境 A／B／C（百萬美元）", "x": rng("Seg", hdr, "G", "L"), "fmt": "n0", "stack": False, "pad": 1.2,
             "bars": [{"name": [f"@Seg!A{Sg[s]['hd']}"], "y": rng("Seg", Sg[s]["net"], "G", "L"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
             "caption": "三情境的 AI 淨效果差異＝客戶數成長、工作負載成長、單價侵蝕三個輸入的差別，分化依據見 README 情境分化依據。"}
seg_keys = ["cg", "wl", "pe", "gg", "cust", "arpc", "usc", "ic", "com", "gov", "rev", "yoy", "inc", "ero", "tasks", "cost", "share", "incinf", "net", "be"]
seg_fmt = {"cg": "p1", "wl": "p1", "pe": "p1", "gg": "p1", "yoy": "p1", "share": "p2", "be": "p1", "arpc": "n2", "tasks": "n1", "cust": "n0"}
SCN_H, SCN_ROWS = rm["SCN_H"], rm["SCN_ROWS"]
INK = {k: v for k, v in IR.items()}
S.append({"id": "drivers", "title": "驅動因子：分部營收推導", "blocks": [
    {"type": "note", "parts": [f"@README!A{RDR['drv']}"]}, ai_chart, scn_chart,
    {"type": "table", "caption": "分部營收與命題模組（情境 B）", "head": [f"@Seg!A{hdr}"] + [f"@Seg!{c}{hdr}" for c in "FGHIJKL"],
     "rows": [[f"@Seg!A{SB_[k]}"] + [f"@Seg!{c}{SB_[k]}|{seg_fmt.get(k, 'n1')}" for c in "FGHIJKL"] for k in seg_keys]},
    {"type": "table", "caption": "情境驅動（三情境並排；進加權目標價）", "head": [f"@In!{c}{SCN_H}" for c in "BDEF"],
     "rows": [[f"@In!B{SCN_ROWS[k]}"] + [f"@In!{c}{SCN_ROWS[k]}|p0" for c in "DEF"] for k in ("scn_cg", "scn_wl", "scn_pe", "scn_gg", "scn_w")]},
    {"type": "table", "caption": "單一參數區間（In 頁；基準進主模型，敏感度低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[f"@In!B{IR[k]}", f"@In!D{IR[k]}|{f}", f"@In!E{IR[k]}|{f}", f"@In!F{IR[k]}|{f}", f"@In!G{IR[k]}|tag"]
              for k, f in (("tk_p", "n3"), ("tk_mult", "n1"), ("tk_n", "n1"), ("tk_pg", "p0"), ("cm_g", "p1"), ("cm_c", "p1"), ("ua_r", "p1"), ("ic27", "p0"), ("gov_shock", "p0"), ("wc_k", "p0"), ("beta", "n3"), ("g_tv", "p1"), ("w_dcf", "p0"), ("gth", "n1"), ("gcap", "n1"))]}]})
# ---- recon ----
RC = RCK
S.append({"id": "recon", "title": "對帳：共識、公司指引、落差分解與逾期標示", "intro": "@Recon!A1", "blocks": [
    {"type": "table", "caption": f"@Recon!A{RC['s1']}", "head": [f"@Recon!{c}{RC['h1']}" for c in "ABCDEFGH"],
     "rowgen": {"sheet": "Recon", "rows": f"{RC['rev26']}-{RC['eps27']}", "cols": list("ABCDEFGH"), "fmts": ["", "", "n1", "n1", "n1", "n1", "sp1", ""]}},
    {"type": "table", "caption": f"@Recon!A{RC['s2']}", "head": [f"@Recon!{c}{RC['h2']}" for c in "ABCDEF"],
     "rowgen": {"sheet": "Recon", "rows": f"{RC['q4']}-{RC['usc']}", "cols": list("ABCDEF"), "fmts": ["", "n1", "n1", "n1", "", ""]}},
    {"type": "table", "caption": f"@Recon!A{RC['s3']}", "head": [f"@Recon!{c}{RC['h3']}" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{RC['d_ps']}-{RC['d_sbcw']}", "cols": list("ABCDE"), "fmts": ["", "n2", "n2", "n2", ""]}},
    {"type": "table", "caption": f"@Recon!A{RC['s4']}", "head": [f"@Recon!{c}{RC['h4']}" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{RC['k_wacc']}-{RC['k_chk']}", "cols": list("ABCDE"), "fmts": ["", "n3", "n3", "n2", ""], "row_fmt": {str(RC["k_wl"]): {"B": "p1", "C": "p1"}, str(RC["k_wacc"]): {"B": "p2", "C": "p2"}}}},
    {"type": "table", "caption": f"@Recon!A{RC['s5']}", "collapse": True, "summary": "展開：市場資料日期與逾期標示（逐筆）", "head": [f"@Recon!{c}{RC['h5']}" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{RC['dt0']}-{RC['s6'] - 1}", "cols": list("ABCDE"), "fmts": ["", "", "n0", "", ""]}},
    {"type": "table", "caption": f"@Recon!A{RC['s6']}", "head": [f"@Recon!{c}{RC['h6']}" for c in "ABCD"],
     "rowgen": {"sheet": "Recon", "rows": f"{RC['fc']}-{RC['joint']}", "cols": list("ABCD"), "fmts": ["", "n3", "", ""], "row_fmt": {str(RC["ratio"]): {"B": "n1"}, str(RC["joint"]): {"B": "p2"}}}}]})
# ---- status ----
rk0 = RDR["risk_h"]
S.append({"id": "status", "title": "資料狀態", "blocks": [
    {"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RDR['ver']}"]},
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期（資料截止日、成品日、股價基準日）", "@README!A1"], ["build_id", f"@Sens!D{rm['SE_BID']}"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "h3", "text": "資料與口徑"},
    {"type": "list", "items": ["@Q!A1" if False else f"@In!A2", "@Sens!A2", f"@README!A{RDR['na1']}"]},
    {"type": "h3", "text": "失效條件與承重假設"},
    {"type": "list", "items": [f"@README!A{RDR[k]}" for k in ("f1", "f2", "f3", "ld1", "ld2", "ld3", "sc1")]},
    {"type": "table", "caption": "風險登錄", "head": ["風險", "是否建模", "理由", "觀察指標"],
     "rowgen": {"sheet": "README", "rows": f"{RDR['risk0']}-{RDR['risk4']}", "cols": ["A", "B", "C", "D"]}},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A", "E"], "In": ["B", "H", "I"]}, "empty": "沒有標為未找到的項目。"}]})
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

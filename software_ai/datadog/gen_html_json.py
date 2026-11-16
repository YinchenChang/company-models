# 由 build_datadog.py 的列號對照（ROWMAP）產生 software_ai/datadog/html.json（版面規格 v1.3；列號隨模型變動，勿手改 html.json）。
# 用法：python3 software_ai/datadog/gen_html_json.py <成品 xlsx> <rowmap.json> software_ai/datadog/html.json
import json, sys
import openpyxl

X, RM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
rm = json.load(open(RM, encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)
PR, TR, IX, VR, RVR, RCR, INR, SCN = rm["PR"], rm["TR"], rm["IX"], rm["VR"], rm["RVR"], rm["RCR"], rm["IN"], rm["SCN"]
PB, TB = PR["B"], TR["B"]
HR = rm["PL_HDR"]; CHR = rm["CHR"]
STD = ["C", "D", "E", "H", "I", "J", "K", "L", "M", "N"]
FC = "HIJKLMN"


def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")


def met(name, key, fmt, chart=False):
    m = {"name": name, "fmt": fmt, "rows": {s: f"@PL!H{PR[s][key]}:N{PR[s][key]}" for s in "ABC"}}
    if chart: m["chart"] = True
    return m


# ---- 敏感度 ----
SE = wb["Sens"]
case_rows = [r for r in rm["SE_ROWS"] if isinstance(SE.cell(r, 5).value, (int, float))]
sens_blocks = []
for side in ("低", "高", "聯動"):
    rs = [r for r in case_rows if SE.cell(r, 3).value == side]
    if not rs: continue
    big = [r for r in rs if abs(SE.cell(r, 5).value) >= 0.1]
    small = [r for r in rs if 0.0005 <= abs(SE.cell(r, 5).value) < 0.1]
    base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": "E", "fmt": "sp1"}
    sens_blocks.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big or rs[:1]))))
    if small:
        sens_blocks.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，其餘案例", rows=",".join(map(str, small)), collapse=True, summary=f"展開：輸入取{side}值，變動未達兩位數百分比的案例"))
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：全部案例（含收支關卡、DCF 目標價與投資建議）",
                    "head": ["驅動", "設定", "方向", "加權目標價", "相對基準", "收支關卡 B", "DCF 目標價", "投資建議"],
                    "rowgen": {"sheet": "Sens", "rows": f"4-{rm['SE_END']}", "cols": ["A", "B", "C", "D", "E", "H", "I", "J"], "fmts": ["", "", "", "n2", "sp1", "n1", "n2", ""]}})
sens_blocks.append({"type": "kv", "items": [[f"@Sens!A{rm['SE_BASE']}", f"@Sens!D{rm['SE_BASE']}|n2"], ["模型狀態", f"@Sens!D{rm['SE_STATUS']}"], ["build_id", f"@Sens!D{rm['SE_BID']}"]]})
_main = [b for b in sens_blocks if not b.get("collapse") and b["type"] == "tornado"]
sens = {"id": "sens", "title": "敏感度", "blocks": _main[:2] + [{"type": "note", "parts": ["@Sens!A2"]}] + _main[2:] + [b for b in sens_blocks if b not in _main]}

std_keys = ["rev", "yoy", "cogs", "gp", "gm", "opex", "oi", "om", "oth", "pti", "tax", "ni", "nm", "eps"]
def row_std(k):
    r = IX[k]; u = wb["PL"].cell(r, 2).value
    f = "p1" if u == "%" else ("n2" if k == "eps" else "n1")
    return [f"@PL!A{r}"] + [f"@PL!{c}{r}|{f}" for c in STD]
head_std = [f"@PL!A{HR}"] + [f"@PL!{c}{HR}" for c in STD]
yr_h = PR["A"]["yr"] - 1
cfg = {
    "company": "Datadog", "ticker": "DDOG", "layout": "v1.3", "wide_text": True,
    "thesis": "@README!A3",
    "question": "AI 淨效果（增量減價格讓利減推論成本）在三情境是否為正，且是否足以支持現價",
    "scenarios": [{"key": "A", "name": "情境 A（代理工作負載放大、讓利較小）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（最大客戶停留最低承諾、讓利較大）"}],
    "hero": {"rec": vv("rec"), "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n2"), ["上檔空間 ", vv("up", "B", "sp1")]],
                                       [f"@Val!A{VR['tpcomp']}", vv("tpcomp", "B", "n2"), ["上檔空間 ", vv("upc", "B", "sp1")]],
                                       [f"@Val!A{VR['px']}", vv("px", "B", "n2"), ""]],
             "note": ["@README!A11"]},
    "sections": [],
}
thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "note", "parts": ["@README!A5"]},
    {"type": "theme", "title": "營收（百萬美元，依命題子列堆疊）與 GAAP 營業利益率：FY2026E–FY2032E（情境 B）", "x": f"@PL!H{HR}:N{HR}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": f"@PL!H{PB['rev']}:N{PB['rev']}",
     "bars": [{"name": [f"@PL!A{PB['sN']}"], "y": f"@PL!H{PB['sN']}:N{PB['sN']}", "color": "#6fa8dc"},
              {"name": [f"@PL!A{PB['sA']}"], "y": f"@PL!H{PB['sA']}:N{PB['sA']}", "color": "#0b7a80"},
              {"name": [f"@PL!A{PB['sI']}"], "y": f"@PL!H{PB['sI']}:N{PB['sI']}", "color": "#c29a2b"}],
     "lines": [{"name": [f"@PL!A{PB['om']}"], "y": f"@PL!H{PB['om']}:N{PB['om']}", "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝命題拆分的三條子列（官方只有單一應報導部門，三條合計等於合併營收）；折線＝GAAP 營業利益率（右軸，已扣股票薪酬）。"},
    {"type": "h3", "text": "命題與核心問題"},
    {"type": "list", "items": ["@README!A3", "@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [f"@README!A{r}" for r in (6, 7, 8, 9, 10)]},
    {"type": "h3", "text": "情境分化依據"},
    {"type": "note", "parts": ["@README!A13"]},
    {"type": "h3", "text": "發布前合理性關卡診斷"},
    {"type": "note", "parts": ["@README!A18"]}]}
valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n0",
     "items": [{"label": f"@Val!A{VR[k]}", "value": f"@Val!B{VR[k]}", "series": sr} for k, sr in (("tpdcf", 1), ("tpcomp", 2), ("tp", 0), ("cons", 1))],
     "ref": {"label": f"@Val!A{VR['px']}", "value": f"@Val!B{VR['px']}"}},
    {"type": "note", "parts": ["@Val!A2"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值", "說明"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f), f"@Val!E{VR[k]}"] for k, f in (("tpdcf", "n2"), ("psv_s", "n2"), ("tpdcf_s", "n2"), ("tpcomp", "n2"), ("tp", "n2"), ("px", "n2"), ("cons", "n2"), ("up", "sp1"), ("upc", "sp1"), ("tier", ""), ("gate_chk", ""), ("rec", ""))]},
    {"type": "kv", "items": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("rf", "p2"), ("beta", "n2"), ("erp", "p1"), ("ke", "p2"), ("kdat", "p2"), ("wacc", "p2"), ("ronic", "p2"), ("g", "p1"))]},
    {"type": "table", "caption": "DCF 彙總（三情境；百萬美元，每股為美元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['pvx']}-{VR['dcf12_s']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"],
                "row_fmt": {str(VR["tvsh"]): {c: "p1" for c in "BCD"}, str(VR["shv"]): {c: "n2" for c in "BCD"}, str(VR["psv"]): {c: "n2" for c in "BCD"}, str(VR["dcf12"]): {c: "n2" for c in "BCD"}, str(VR["dcf12_s"]): {c: "n2" for c in "BCD"}}}},
    {"type": "table", "caption": "Comparable（同業 Forward 倍數 × 目標時點 NTM）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['fmix']}-{VR['comp']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n2", "n2", "n2"],
                "row_fmt": {str(VR["fmix"]): {c: "n3" for c in "BCD"}, str(VR["ntm_rev"]): {c: "n0" for c in "BCD"}, str(VR["ntm_sh"]): {c: "n1" for c in "BCD"}, str(VR["m_ps"]): {"B": "n2"}, str(VR["m_pe"]): {"B": "n2"}}}},
    {"type": "table", "caption": f"@Val!A{VR['g_tot'] - 1}", "head": ["項目", "每股（美元）"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("g_tot", "sp2"), ("g_sbc", "sp2"), ("g_mult", "sp2"), ("g_lev", "n2"), ("g_time", "sp2"), ("g_chk", "n3"))]},
    {"type": "table", "caption": "同業倍數、本公司現行倍數與逾期核對", "head": ["項目", "Forward P/S", "Forward P/E", "資料日", "狀態"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['p_snow']}-{VR['p_ddog']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "", ""]}},
    {"type": "table", "caption": "現行倍數 ÷ 同業中位數與隱含重估", "head": ["項目", "本公司", "同業中位數", "比值", "隱含重估"],
     "rows": [[f"@Val!A{VR[k]}", f"@Val!B{VR[k]}|n2", f"@Val!C{VR[k]}|n2", f"@Val!D{VR[k]}|n2", f"@Val!E{VR[k]}|sp1"] for k in ("c_ps", "c_pe")]},
    {"type": "table", "caption": "情境目標價與收支關卡", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("t_A", "n2"), ("t_B", "n2"), ("t_C", "n2"), ("wsum", "n3"), ("fcf_gate", "n0"), ("gate_A", "n0"), ("gate_C", "n0"))]}]}
revdcf = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "blocks": [
    {"type": "note", "parts": ["@RevDCF!A2"]}, {"type": "note", "parts": [f"@RevDCF!A{RVR['sum']}"]}, {"type": "note", "parts": [f"@RevDCF!A{RVR['sum2']}"]},
    {"type": "table", "caption": "現金流折現法（情境 B 營收路徑）：現價隱含值與模型", "head": ["項目", "現價隱含", "模型情境 B"],
     "rows": [[f"@RevDCF!A{RVR['m1']}", f"@RevDCF!B{RVR['m1']}|p1", f"@RevDCF!B{RVR['m0']}|p1"], [f"@RevDCF!A{RVR['evr']}", f"@RevDCF!B{RVR['evr']}|n0", "—"],
              [f"@RevDCF!A{RVR['rv0']}", "—", f"@RevDCF!B{RVR['rv0']}|n0"], [f"@RevDCF!A{RVR['w0']}", "—", f"@RevDCF!B{RVR['w0']}|p2"]]},
    {"type": "table", "caption": "同業倍數法：現行倍數相對同業中位數", "head": ["項目", "數值"],
     "rows": [[f"@RevDCF!A{RVR[k]}", f"@RevDCF!B{RVR[k]}|{f}"] for k, f in (("mps", "n2"), ("mpe", "n2"), ("rrps", "sp1"), ("rrpe", "sp1"), ("sps", "n2"))]}]}
pl = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "note", "parts": ["損益表採美國會計準則（GAAP）口徑，股票薪酬為費用；公司指引與市場多用 non-GAAP，見上方情境區塊與對帳章。", f"@PL!A{rm['PL_NOTE']}"]},
    {"type": "theme", "title": "營收（百萬美元）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{CHR['hdr']}:L{CHR['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
     "bars": [{"name": [f"@PL!A{CHR['rev']}"], "y": f"@PL!C{CHR['rev']}:L{CHR['rev']}", "color": "#6fc7cb"}],
     "lines": [{"name": [f"@PL!A{CHR['gm']}"], "y": f"@PL!C{CHR['gm']}:L{CHR['gm']}", "axis": "r", "color": "#0b7a80", "labels": True},
               {"name": [f"@PL!A{CHR['om']}"], "y": f"@PL!C{CHR['om']}:L{CHR['om']}", "axis": "r", "color": "#c29a2b"},
               {"name": [f"@PL!A{CHR['nm']}"], "y": f"@PL!C{CHR['nm']}:L{CHR['nm']}", "axis": "r", "color": "#6f6bae", "dash": True}],
     "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP）。毛利率走低反映 AI 原生客戶價格讓利與推論成本；淨利率含利息收入。"},
    {"type": "table", "caption": "標準損益表（GAAP；百萬美元，EPS 為美元）：前三個年度實績、後七個年度情境 B", "head": head_std,
     "rows": [row_std(k) for k in std_keys] + [{"sub": f"@PL!A{IX['r_yoy'] - 1}"}] + [[f"@PL!A{IX[k]}"] + [f"@PL!{c}{IX[k]}|p1" for c in STD] for k in ("r_yoy", "r_gm", "r_om", "r_nm", "r_niy")]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線）", "years": f"@PL!H{yr_h}:N{yr_h}",
     "metrics": [met("合併營收（百萬美元）", "rev", "n0", True), met("營業利益（GAAP，百萬美元）", "oi", "n0", True), met("營業利益率（GAAP）", "om", "p1"), met("EPS（GAAP 稀釋，美元）", "eps", "n2", True),
                 met("EPS（non-GAAP 稀釋，美元）", "ngeps", "n2")]},
    {"type": "table", "caption": "分部對帳（三情境 × 全部年度取絕對值最大者，應為零）", "head": ["項目"] + [f"@PL!{c}{HR}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{rm['PL_SUM']}-{rm['PL_SUM3']}", "cols": ["A"] + STD, "fmts": [""] + ["n3"] * 10}}]}
cols_s = ["A"] + list(FC)
drivers = {"id": "drivers", "title": "驅動因子：命題推導（用量 × 單價）", "blocks": [
    {"type": "note", "parts": ["@Thesis!A1"]}, {"type": "note", "parts": ["@Thesis!A2"]},
    {"type": "theme", "title": "AI 淨效果（百萬美元）× 情境 A／B／C", "x": f"@PL!H{yr_h}:N{yr_h}", "fmt": "n0", "stack": False, "pad": 1.2,
     "bars": [{"name": ["情境 ", s, "：AI 淨效果"], "y": f"@Thesis!H{TR[s]['net']}:N{TR[s]['net']}", "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
     "caption": "AI 淨效果＝AI 營收增量 − AI 原生客戶價格讓利 − 推論成本；三情境在用量成長、讓利幅度與 AI credits 營收分化。"},
    {"type": "table", "caption": "命題推導（情境 B）：用量指數、單價指數、AI 淨效果與推論成本", "head": ["項目"] + [f"@PL!{c}{yr_h}" for c in FC],
     "rowgen": {"sheet": "Thesis", "rows": f"{TB['rev']}-{TB['pidx']}", "cols": cols_s, "fmts": [""] + ["n2"] * 7,
                "row_fmt": {**{str(TB[k]): {c: "p1" for c in FC} for k in ("vN", "vA", "pA", "gai", "rm")}, **{str(TB[k]): {c: "n3" for c in FC} for k in ("volN", "priN", "volA", "priA", "pxt", "crp", "ctask", "pbe", "pact", "pbe_task", "pidx")}, str(TB["yr"]): {c: "yr" for c in FC}}}},
    {"type": "table", "caption": "AI 淨效果與其組成（三情境 FY2032E；百萬美元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@Thesis!A{TR['B'][k]}"] + [f"@Thesis!N{TR[s][k]}|n1" for s in "ABC"] for k in ("inc", "eros", "inf2", "net", "pbe")]},
    {"type": "table", "caption": "三情境驅動（進加權目標價；In 頁情境驅動對照）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{SCN[k]}", f"@In!D{SCN[k]}|{f}", f"@In!E{SCN[k]}|{f}", f"@In!F{SCN[k]}|{f}"] for k, f in (("scn_vn", "p0"), ("scn_va", "p0"), ("scn_pa", "p0"), ("scn_ai", "n0"), ("scn_w", "p0"))]},
    {"type": "table", "caption": "單一參數區間（基準進主模型；敏感度低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[f"@In!B{INR[k]}", f"@In!D{INR[k]}|{f}", f"@In!E{INR[k]}|{f}", f"@In!F{INR[k]}|{f}", f"@In!G{INR[k]}|tag"]
              for k, f in (("sa_26", "p0"), ("vn_32", "p0"), ("va_32", "p0"), ("pn_g", "p1"), ("cost_sh", "p0"), ("tk_px", "n3"), ("route_sh", "p0"), ("pt_k", "p0"), ("gm_core", "p1"), ("sbc_st", "p2"), ("sbc_floor", "p0"),
                           ("wc_k", "p1"), ("beta", "n2"), ("erp", "p1"), ("g_tv", "p1"), ("w_dcf", "p0"), ("w_pe", "p0"), ("rel_m", "n2"))]}]}
recon = {"id": "recon", "title": "對帳：公司指引、共識、校準與合理性關卡", "blocks": [
    {"type": "note", "parts": ["@Recon!A2"]},
    {"type": "table", "caption": "公司指引 vs 模型情境 B（公司指引＝Interested-party）", "head": ["項目", "期間", "模型", "指引下緣", "指引上緣", "模型相對指引中點"],
     "rowgen": {"sheet": "Recon", "rows": f"{RCR['q4rev']}-{RCR['fyeps']}", "cols": ["A", "B", "C", "D", "E", "F"], "fmts": ["", "", "n1", "n1", "n1", "sp1"], "row_fmt": {str(RCR["fyeps"]): {c: "n2" for c in "CDE"}}}},
    {"type": "table", "caption": "市場共識對模型（只作對照）", "head": ["項目", "期間", "情境 A", "情境 B", "情境 C", "共識"],
     "rowgen": {"sheet": "Recon", "rows": f"{RCR['g27']}-{RCR['tpc']}", "cols": ["A", "B", "C", "D", "E", "F"], "fmts": ["", "", "n1", "n1", "n1", "n1"],
                "row_fmt": {str(RCR["g27"]): {c: "p1" for c in "CDEF"}, str(RCR["eps27"]): {c: "n2" for c in "CDEF"}, str(RCR["tpc"]): {"C": "n2", "D": "n2", "E": "sp1"}}}},
    {"type": "table", "caption": "每任務 token 校準（公司揭露 vs 聊天基數類推）", "head": ["項目", "單位", "數值"],
     "rowgen": {"sheet": "Recon", "rows": f"{RCR['tk_d']}-{RCR['tk_rev']}", "cols": ["A", "B", "C"], "fmts": ["", "", "n3"]}},
    {"type": "table", "caption": "反解：主導變數整本重算", "head": ["項目", "單位", "數值", "對照"],
     "rowgen": {"sheet": "Recon", "rows": f"{RCR['sv_var']}-{RCR['sv_sbcfloor']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "", "p2", "n2"], "row_fmt": {str(RCR["sv_tp"]): {"C": "n2", "D": "n2"}}}},
    {"type": "table", "caption": "合理性關卡（發布前必查）", "head": ["項目", "判定值", "門檻", "是否觸發"],
     "rowgen": {"sheet": "Recon", "rows": f"{RCR['gt1']}-{RCR['gt_n']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n2", "", ""],
                "row_fmt": {str(RCR["gt2"]): {"B": "n0"}, str(RCR["gt4"]): {"B": "sp1"}, str(RCR["gt5"]): {"B": "sp1"}, str(RCR["gt6"]): {"B": "n1"}, str(RCR["gt_n"]): {"B": "n0"}, str(RCR["gt3"]): {"B": ""}}}},
    {"type": "table", "caption": "DCF 與公司自由現金流對帳", "head": ["項目", "期間", "百萬美元", "終值折現年數", "終值占 EV"],
     "rowgen": {"sheet": "Recon", "rows": f"{RCR['rc_fcf']}-{RCR['rc_txt']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "", "n1", "n2", "p0"], "row_fmt": {str(RCR["rc_chk"]): {"C": "n3"}, str(RCR["rc_wacc"]): {"C": "p2"}, str(RCR["rc_psv_s"]): {"C": "n2"}, str(RCR["rc_tp_s"]): {"C": "n2", "D": "n2", "E": "n2"}}}}]}
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["build_id", f"@Sens!D{rm['SE_BID']}"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "h3", "text": "本版要點與限制"},
    {"type": "list", "items": ["@README!A2", "@README!A14", "@README!A15", "@README!A16", "@README!A17", "@README!A43"]},
    {"type": "h3", "text": "市場資料日期與失效條件"},
    {"type": "list", "items": ["@README!A30", "@README!A32"]},
    {"type": "table", "caption": "風險登錄", "head": ["風險", "是否建模", "理由", "觀察指標"], "rowgen": {"sheet": "README", "rows": "35-39", "cols": ["A", "B", "C", "D"]}},
    {"type": "table", "collapse": True, "summary": "展開：驅動 → 推導對照表", "head": ["頁", "輸入", "推導量", "因果方向"], "rowgen": {"sheet": "README", "rows": "21-27", "cols": ["A", "B", "C", "D"]}},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A"], "In": ["B", "I"]}, "empty": "沒有標為未找到的項目。"}]}
cfg["sections"] = [thesis, valuation, sens, revdcf, pl, drivers, recon, status]
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

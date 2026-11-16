# 由 build_naver.py 的列號對照（ROWMAP）產生 software_ai/naver/html.json（版面規格 v1.3；列號隨模型變動，勿手改 html.json）。
# 用法：python3 software_ai/naver/gen_html_json.py <成品 xlsx> <rowmap.json> software_ai/naver/html.json
import json, sys
import openpyxl

X, RM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
rm = json.load(open(RM, encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)
PR, IX, RID, VR, RVR, RCR, DR, INR = rm["PR"], rm["IX"], rm["RID"], rm["VR"], rm["RVR"], rm["RCR"], rm["DR"], rm["IN"]
HR, STDH = rm["HR"], rm["STDH"]
PB = PR["B"]
C = list("CDEFGHIJKL")


def rng(sh, r, c1="F", c2="L"):
    return f"@{sh}!{c1}{r}:{c2}{r}"


def vv(k, col="B", f=""):
    return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")


SE = wb["Sens"]
case_rows = [r for r in rm["SE_ROWS"] if isinstance(SE.cell(r, 5).value, (int, float))]
sens_blocks = []
for side in ("低", "高"):
    rs = [r for r in case_rows if SE.cell(r, 3).value == side]
    big = [r for r in rs if abs(SE.cell(r, 5).value) >= 0.1]
    small = [r for r in rs if 0.0005 <= abs(SE.cell(r, 5).value) < 0.1]
    base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": "E", "fmt": "sp1"}
    pick = sorted(big or rs, key=lambda r: -abs(SE.cell(r, 5).value))[:1] if not big else sorted(big, key=lambda r: -abs(SE.cell(r, 5).value))
    sens_blocks.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, pick))))
    if small:
        sens_blocks.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，其餘案例", rows=",".join(map(str, sorted(small, key=lambda r: -abs(SE.cell(r, 5).value)))), collapse=True,
                                summary=f"展開：輸入取{side}值，變動未達兩位數百分比的案例"))
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：全部案例（含收支關卡 B、DCF 目標價與投資建議）",
                    "head": ["驅動", "設定", "方向", "加權目標價", "相對基準", "收支關卡 B", "DCF 目標價", "投資建議"],
                    "rowgen": {"sheet": "Sens", "rows": f"4-{rm['SE_END']}", "cols": ["A", "B", "C", "D", "E", "H", "I", "J"], "fmts": ["", "", "", "n0", "sp1", "n1", "n0", ""]}})
sens_blocks.append({"type": "kv", "items": [[f"@Sens!A{rm['SE_BASE']}", f"@Sens!D{rm['SE_BASE']}|n0"], ["模型狀態", f"@Sens!D{rm['SE_STATUS']}"], ["build_id", f"@Sens!D{rm['SE_BID']}"]]})


def met(name, key, fmt, chart=False):
    m = {"name": name, "fmt": fmt, "rows": {s: rng("PL", PR[s][key]) for s in "ABC"}}
    if chart: m["chart"] = True
    return m


def scn_table(keys):
    return {"type": "table", "caption": "AI 淨效果與其組成（三情境 FY2032E；十億韓元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
            "rows": [[f"@PL!A{PB[k]}"] + [f"@PL!L{PR[s][k]}|n1" for s in "ABC"] for k in keys]}


std_keys = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pti", "tax", "ni", "nm", "eps"]


def row_std(k):
    r = IX[k]
    u = wb["PL"].cell(r, 2).value
    f = "p1" if u == "%" else ("n0" if k == "eps" else "n1")
    return [f"@PL!A{r}"] + [f"@PL!{c}{r}|{f}" for c in C]


head_std = [f"@PL!A{STDH}"] + [f"@PL!{c}{STDH}" for c in C]
cfg = {
    "company": "Naver", "ticker": "035420", "layout": "v1.3", "wide_text": True,
    "thesis": "@README!A3",
    "question": "AI 簡報進入搜尋後，廣告單位營收的增量能否抵過外部 AI 分流與自建模型的推論成本",
    "scenarios": [{"key": "A", "name": "情境 A（AI 簡報變現佳、無外部分流）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（AI 回答變現差、外部分流大）"}],
    "hero": {"rec": vv("rec"), "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n0"), ["上檔空間 ", vv("up", "B", "sp1")]],
                                       [f"@Val!A{VR['tpdcf']}", vv("tpdcf", "B", "n0"), ["上檔空間 ", f"@Val!B{VR['tpdcf']}/@Val!B{VR['px']}-1|sp1"] if False else ""],
                                       [f"@Val!A{VR['px']}", vv("px", "B", "n0"), ""]],
             "note": ["@README!A11"]},
    "sections": [],
}
thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "note", "parts": ["@README!A5"]},
    {"type": "theme", "title": "營收（十億韓元，依分部堆疊）與營業利益率：FY2025A–FY2032E（情境 B）", "x": f"@PL!E{HR}:L{HR}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": f"@PL!E{PB['rev']}:L{PB['rev']}",
     "bars": [{"name": [f"@PL!A{PB['np']}"], "y": f"@PL!E{PB['np']}:L{PB['np']}", "color": "#6fa8dc"},
              {"name": [f"@PL!A{PB['fin']}"], "y": f"@PL!E{PB['fin']}:L{PB['fin']}", "color": "#0b7a80"},
              {"name": [f"@PL!A{PB['glob']}"], "y": f"@PL!E{PB['glob']}:L{PB['glob']}", "color": "#c29a2b"}],
     "lines": [{"name": [f"@PL!A{PB['om']}"], "y": f"@PL!E{PB['om']}:L{PB['om']}", "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝三個官方分部（2026 年起新口徑，FY2025A 為重編後的季度加總）；折線＝合併營業利益率（右軸）。FY2026E 為九個月實績加第四季預測，其後為情境 B。Global Opportunities 含 AI Factory 模組營收。"},
    {"type": "h3", "text": "命題與核心問題"},
    {"type": "list", "items": ["@README!A3", "@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [f"@README!A{r}" for r in (6, 7, 8, 9, 10)]},
    {"type": "h3", "text": "情境分化依據"},
    {"type": "note", "parts": ["@README!A13"]}]}
valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（韓元；豎線＝現價）", "fmt": "n0",
     "items": [{"label": f"@Val!A{VR[k]}", "value": f"@Val!B{VR[k]}", "series": sr} for k, sr in (("tpdcf", 1), ("tpcomp", 2), ("tp", 0), ("cons", 1))],
     "ref": {"label": f"@Val!A{VR['px']}", "value": f"@Val!B{VR['px']}"}},
    {"type": "note", "parts": ["@Val!A2"]},
    {"type": "note", "parts": ["@README!A12"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值", "說明"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f), f"@Val!E{VR[k]}"] for k, f in (("tpdcf", "n0"), ("tpcomp", "n0"), ("tp", "n0"), ("px", "n0"), ("cons", "n0"), ("up", "sp1"), ("upc", "sp1"), ("tier", ""), ("val_rating", ""), ("gate_val", "n0"), ("rec", ""))]},
    {"type": "kv", "items": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("rf", "p2"), ("rf_flag", ""), ("beta", "n2"), ("erp", "p1"), ("ke", "p2"), ("kdat", "p2"), ("wacc", "p2"), ("ronic", "p2"), ("g", "p1"))]},
    {"type": "table", "caption": "DCF 彙總（三情境；十億韓元，每股為韓元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['pvx']}-{VR['dcf12']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"],
                "row_fmt": {str(VR["tvsh"]): {c: "p1" for c in "BCD"}, str(VR["shv"]): {c: "n0" for c in "BCD"}, str(VR["nonop_ps"]): {c: "n0" for c in "BCD"}, str(VR["dcf12"]): {c: "n0" for c in "BCD"}}}},
    {"type": "table", "caption": "Comparable：分部 NTM NOPAT 乘同業前瞻本益比（目標時點）", "head": ["項目", "情境 A", "情境 B", "情境 C", "同業群組倍數"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['n_np']}-{VR['ntm_mult']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1", "x1"],
                "row_fmt": {str(VR["s_ps"]): {c: "n0" for c in "BCD"}, str(VR["ntm_mult"]): {c: "x1" for c in "BCD"}, str(VR["v_np"]): {"E": "x1"}, str(VR["v_fin"]): {"E": "x1"}, str(VR["v_glob"]): {"E": "x1"}}}},
    {"type": "table", "caption": "同業前瞻本益比逐家（會計準則、基準日與逾期旗標）", "head": ["同業", "前瞻本益比", "基準日", "會計準則與租賃處理", "逾期", "群組"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['peh'] + 1}-{VR['peer_end']}", "cols": list("ABCDEF"), "fmts": ["", "n2", "date", "", "", ""]}},
    {"type": "table", "caption": "本公司現行倍數與 DCF／Comparable 落差分解（情境 B）", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("own_pe", "n1"), ("own_ratio", "n2"), ("own_rerate", "sp1"), ("comp_alt", "n0"), ("tp_alt", "n0"), ("gp_gap", "n0"), ("gp_sbc", "n0"), ("gp_lev", "n0"), ("gp_time", "n0"), ("gp_mult", "n0"), ("gp_chk", "n0"), ("gp_dcfmult", "n1"))]},
    {"type": "table", "caption": "情境目標價與收支關卡", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("t_A", "n0"), ("t_B", "n0"), ("t_C", "n0"), ("wsum", "p0"), ("per", "n1"), ("gate_A", "n0"), ("gate_B", "n0"), ("gate_C", "n0"))]}]}
_main = [b for b in sens_blocks if not b.get("collapse") and b["type"] == "tornado"]
sens = {"id": "sens", "title": "敏感度", "blocks": _main + [{"type": "note", "parts": ["@Sens!A2"]}] + [b for b in sens_blocks if b not in _main]}
revdcf = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "blocks": [
    {"type": "note", "parts": ["@RevDCF!A2"]}, {"type": "note", "parts": [f"@RevDCF!A{RVR['sum']}"]},
    {"type": "table", "caption": "現金流折現法（情境 B 營收路徑）：現價隱含值與模型", "head": ["項目", "現價隱含", "模型情境 B"],
     "rows": [[f"@RevDCF!A{RVR['m1']}", f"@RevDCF!B{RVR['m1']}|p1", f"@RevDCF!B{RVR['m0']}|p1"], [f"@RevDCF!A{RVR['evr']}", f"@RevDCF!B{RVR['evr']}|n0", "—"],
              [f"@RevDCF!A{RVR['rv0']}", "—", f"@RevDCF!B{RVR['rv0']}|n0"], [f"@RevDCF!A{RVR['w0']}", "—", f"@RevDCF!B{RVR['w0']}|p2"]]},
    {"type": "table", "caption": "Comparable：現價隱含的營運倍數與模型", "head": ["項目", "數值"],
     "rows": [[f"@RevDCF!A{RVR[k]}", f"@RevDCF!B{RVR[k]}|{f}"] for k, f in (("smul", "n1"), ("smod", "n1"), ("sps", "n2"))]}]}
pl = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "note", "parts": ["損益表採 K-IFRS 口徑（GAAP 口徑，非 Non-GAAP 調整後數字）。", f"@PL!A{rm['STD_NOTE']}"]},
    {"type": "theme", "title": "營收（十億韓元）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{STDH}:L{STDH}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
     "bars": [{"name": [f"@PL!A{IX['rev']}"], "y": f"@PL!C{IX['rev']}:L{IX['rev']}", "color": "#6fc7cb"}],
     "lines": [{"name": [f"@PL!A{IX['gm']}"], "y": f"@PL!C{IX['gm']}:L{IX['gm']}", "axis": "r", "color": "#0b7a80", "labels": True},
               {"name": [f"@PL!A{IX['opm']}"], "y": f"@PL!C{IX['opm']}:L{IX['opm']}", "axis": "r", "color": "#c29a2b"},
               {"name": [f"@PL!A{IX['nm']}"], "y": f"@PL!C{IX['nm']}:L{IX['nm']}", "axis": "r", "color": "#6f6bae", "dash": True}],
     "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，K-IFRS）。Naver 不揭露毛利，營業成本取 Partner 加 Infrastructure（模型分類）；AI Factory 營收自 FY2027E 起墊高營收、拉低利益率。"},
    {"type": "table", "caption": "標準損益表（K-IFRS；十億韓元，EPS 為韓元）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B", "head": head_std,
     "rows": [row_std(k) for k in std_keys] + [[f"@PL!A{RID[k]}"] + [f"@PL!{c}{RID[k]}|p1" for c in C] for k in ("r_rev", "r_gm", "r_opm", "r_nm", "r_niy")]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線）", "years": f"@PL!F{HR}:L{HR}",
     "metrics": [met("合併營收（十億韓元）", "rev", "n0", True), met("營業利益（十億韓元）", "oi", "n0", True), met("營業利益率", "om", "p1"), met("EPS（韓元）", "eps", "n0", True),
                 met("AI 淨效果：營收層（十億韓元）", "ai_net", "n1", True)]},
    {"type": "table", "caption": "分部對帳（差額應為零）", "head": ["項目", "最大絕對差（十億韓元）"],
     "rowgen": {"sheet": "PL", "rows": f"{rm['PLSUM'] + 1}-{rm['PLSUM'] + 9}", "cols": ["A", "C"], "fmts": ["", "n3"]}}]}
pct = {"lam", "sh", "ctr"}
rf_ = {str(PB[k]): {c: ("p1" if k in pct else "n2" if k in ("reff", "rstar", "al") else "n0" if k in ("u", "cpc") else "n1") for c in C} for k in
       ("q", "lam", "sh", "reff", "al", "ctr", "cpc", "u", "base", "ero", "inc", "ads_exp", "chk_id", "oth_ads", "ads", "svc", "np", "fin", "c2c", "cont", "ent_core", "af_mw", "af", "ent", "glob", "rev", "chk_np", "chk_seg")}
drivers = {"id": "drivers", "title": "驅動因子：廣告單位經濟與分部營收", "blocks": [
    {"type": "note", "parts": ["@PL!A2"]},
    {"type": "theme", "title": "AI 暴露廣告營收（長條）× 情境 A／B／C（十億韓元）", "x": f"@PL!F{HR}:L{HR}", "fmt": "n0", "stack": False, "pad": 1.2,
     "bars": [{"name": ["情境 ", s, "：AI 暴露廣告"], "y": f"@PL!F{PR[s]['ads_exp']}:L{PR[s]['ads_exp']}", "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
     "caption": "AI 暴露廣告＝查詢量 x（扣除外部分流）x 每查詢營收 x（AI 回答占比帶來的相對值調整）；FY2026E 為校準年，三情境相同，其後只在 AI 回答占比、R 與外部分流分化。"},
    {"type": "table", "caption": "廣告單位經濟與分部營收推導（情境 B）：子線加總＝分部，分部加總＝合併營收", "head": ["項目"] + [f"@PL!{c}{HR}" for c in C],
     "rowgen": {"sheet": "PL", "rows": f"{PB['q']}-{PB['chk_seg']}", "cols": ["A"] + C, "fmts": [""] + ["n1"] * 10, "row_fmt": rf_}},
    scn_table(["inc", "ero", "inf", "ai_net", "rstar", "ai_oi"]),
    {"type": "table", "caption": "三情境驅動（AI 驅動並排，進加權目標價；In 頁情境驅動對照）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{rm['SCN_ROWS'][k]}", f"@In!D{rm['SCN_ROWS'][k]}|{f}", f"@In!E{rm['SCN_ROWS'][k]}|{f}", f"@In!F{rm['SCN_ROWS'][k]}|{f}"] for k, f in (("scn_sai", "p0"), ("scn_rai", "n2"), ("scn_lam", "p0"), ("scn_w", "p0"))]},
    {"type": "table", "caption": "單一參數區間（基準進主模型；敏感度低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[f"@In!B{INR[k]}", f"@In!D{INR[k]}|{f}", f"@In!E{INR[k]}|{f}", f"@In!F{INR[k]}|{f}", f"@In!G{INR[k]}|tag"]
              for k, f in (("q0", "n0"), ("gq", "p1"), ("s26", "p0"), ("al_R", "n2"), ("ctr_R", "n2"), ("cpc_R", "n2"), ("tok_mult", "n2"), ("ptok_idx", "n0"), ("fx", "n0"), ("m_np", "p1"), ("m_fin", "p1"), ("m_glob", "p1"),
                           ("capex26", "n0"), ("capex_pct", "p1"), ("beta", "n2"), ("erp", "p1"), ("g_tv", "p1"), ("rel_adj", "n2"), ("w_dcf", "p0"))]}]}
aif = {"id": "aifactory", "title": "AI Factory：算力出租模組", "blocks": [
    {"type": "note", "parts": ["AI Factory 併入 Global Opportunities 的 Enterprise；營收＝年度平均上線容量 x 每容量年營收，營業利益率為 Assumed；資產由 Asset Co 持有，其資本支出與負債未入模型。"]},
    {"type": "theme", "title": "AI Factory 營收與營業利益（十億韓元，情境 B）", "x": f"@PL!G{HR}:L{HR}", "fmt": "n0", "stack": False, "pad": 1.2,
     "bars": [{"name": [f"@PL!A{PB['af']}"], "y": f"@PL!G{PB['af']}:L{PB['af']}", "color": "#0b7a80"}, {"name": [f"@PL!A{PB['oi_af']}"], "y": f"@PL!G{PB['oi_af']}:L{PB['oi_af']}", "color": "#c29a2b"}],
     "caption": "容量與每容量營收取自公司宣布的合約與規劃（Interested-party）；2027 年起營收，利益率為判斷。"},
    {"type": "table", "caption": "AI Factory 推導（情境 B）", "head": ["項目"] + [f"@PL!{c}{HR}" for c in "GHIJKL"],
     "rows": [[f"@PL!A{PB[k]}"] + [f"@PL!{c}{PB[k]}|n1" for c in "GHIJKL"] for k in ("af_mw", "af", "oi_af")]},
    {"type": "table", "caption": "AI Factory 輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[f"@In!B{INR[k]}", f"@In!D{INR[k]}|{f}", f"@In!E{INR[k]}|{f}", f"@In!F{INR[k]}|{f}", f"@In!G{INR[k]}|tag"] for k, f in (("af_mw27", "n0"), ("af_mw28", "n0"), ("af_ext", "n0"), ("af_rev", "n1"), ("af_decl", "p1"), ("af_m27", "p1"), ("af_m28", "p1"))]}]}
recon = {"id": "recon", "title": "對帳：公司指引、市場共識與反解", "blocks": [
    {"type": "note", "parts": ["@Recon!A2"]},
    {"type": "table", "caption": "共識（財報前，Interested-party）vs 模型情境 B", "head": ["項目", "年度", "模型情境 B", "共識", "差距", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{RCR['rev26']}-{RCR['cx27']}", "cols": list("ABCDEF"), "fmts": ["", "", "n1", "n1", "sp1", ""]}},
    {"type": "table", "caption": "九個月實績與財報後券商目標價", "head": ["項目", "數值", "對照", "差距", "說明"],
     "rows": [[f"@Recon!A{RCR['h1v']}", f"@Recon!C{RCR['h1v']}|n1", "", "", f"@Recon!F{RCR['h1v']}"], [f"@Recon!A{RCR['h1v'] + 1}", f"@Recon!C{RCR['h1v'] + 1}|p1", "", "", f"@Recon!F{RCR['h1v'] + 1}"]]
     + [[f"@Recon!A{RCR['tpc'] + i}", f"@Recon!C{RCR['tpc'] + i}|n0", f"@Recon!D{RCR['tpc'] + i}|n0", f"@Recon!E{RCR['tpc'] + i}|sp1", f"@Recon!F{RCR['tpc'] + i}"] for i in range(1, 6)]},
    {"type": "table", "caption": "整本重算反解：加權目標價等於現價所需的主導變數值", "head": ["項目", "反解值", "基準", "相對基準", "說明"],
     "rows": [[f"@Recon!A{RCR['solve_key']}", f"@Recon!C{RCR['solve_key']}", "", "", f"@Recon!F{RCR['solve_key']}"], [f"@Recon!A{RCR['solve_val']}", f"@Recon!C{RCR['solve_val']}|p2", f"@Recon!D{RCR['solve_val']}|p2", f"@Recon!E{RCR['solve_val']}|sp1", f"@Recon!F{RCR['solve_val']}"],
              [f"@Recon!A{RCR['solve_chk']}", f"@Recon!C{RCR['solve_chk']}|n0", f"@Recon!D{RCR['solve_chk']}|n0", "", f"@Recon!F{RCR['solve_chk']}"]]}]}
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["build_id", f"@Sens!D{rm['SE_BID']}"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "h3", "text": "本版要點、診斷與限制"},
    {"type": "list", "items": ["@README!A2", "@README!A30", "@README!A14", "@README!A15", "@README!A16", "@README!A17"]},
    {"type": "h3", "text": "市場資料日期與失效條件"},
    {"type": "list", "items": ["@README!A29", "@README!A31"]},
    {"type": "table", "caption": "風險登錄", "head": ["風險", "是否建模", "理由", "觀察指標"], "rowgen": {"sheet": "README", "rows": "34-38", "cols": ["A", "B", "C", "D"]}},
    {"type": "table", "collapse": True, "summary": "展開：驅動 → 推導對照表", "head": ["頁", "輸入", "推導量", "因果方向"], "rowgen": {"sheet": "README", "rows": "20-26", "cols": ["A", "B", "C", "D"]}},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A", "F"], "In": ["B", "H", "I"], "Q": ["A"]}, "empty": "沒有標為未找到的項目。"}]}
cfg["sections"] = [thesis, valuation, sens, revdcf, pl, drivers, aif, recon, status]
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

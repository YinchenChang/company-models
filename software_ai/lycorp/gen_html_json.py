# 由 build_lycorp.py 的列號對照（ROWMAP）產生 software_ai/lycorp/html.json（版面規格 v1.3；列號隨模型變動，勿手改 html.json）。
# 用法：python3 software_ai/lycorp/gen_html_json.py <成品 xlsx> <rowmap.json> software_ai/lycorp/html.json
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


# ---- 敏感度案例（依相對基準絕對值排序由產生器處理；這裡只列有數值的列） ----
SE = wb["Sens"]
case_rows = [r for r in rm["SE_ROWS"] if isinstance(SE.cell(r, 5).value, (int, float))]
sens_blocks = []
for side in ("低", "高"):
    rs = [r for r in case_rows if SE.cell(r, 3).value == side]
    big = [r for r in rs if abs(SE.cell(r, 5).value) >= 0.1]
    small = [r for r in rs if 0.0005 <= abs(SE.cell(r, 5).value) < 0.1]
    base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": "E", "fmt": "sp1"}
    sens_blocks.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big or rs[:1]))))
    if small:
        sens_blocks.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，其餘案例", rows=",".join(map(str, small)), collapse=True,
                                summary=f"展開：輸入取{side}值，變動未達兩位數百分比的案例"))
_link = [r for r in case_rows if SE.cell(r, 3).value == "聯動"]
if _link:
    sens_blocks.append({"type": "tornado", "title": "加權目標價相對基準：價格與成本聯動、模型路由組合案例", "sheet": "Sens", "rows": ",".join(map(str, _link)), "label_col": "A", "value_col": "E", "fmt": "sp1", "collapse": True,
                        "summary": "展開：價格與成本聯動、模型路由組合案例"})
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：全部案例（含收支關卡 B、DCF 目標價與投資建議）",
                    "head": ["驅動", "設定", "方向", "加權目標價", "相對基準", "收支關卡 B", "DCF 目標價", "投資建議"],
                    "rowgen": {"sheet": "Sens", "rows": f"4-{rm['SE_END']}", "cols": ["A", "B", "C", "D", "E", "H", "I", "J"], "fmts": ["", "", "", "n2", "sp1", "n1", "n2", ""]}})
sens_blocks.append({"type": "kv", "items": [[f"@Sens!A{rm['SE_BASE']}", f"@Sens!D{rm['SE_BASE']}|n2"], ["模型狀態", f"@Sens!D{rm['SE_STATUS']}"], ["build_id", f"@Sens!D{rm['SE_BID']}"]]})


def met(name, key, fmt, chart=False):
    m = {"name": name, "fmt": fmt, "rows": {s: rng("PL", PR[s][key]) for s in "ABC"}}
    if chart: m["chart"] = True
    return m


def scn_table(keys):
    return {"type": "table", "caption": "AI 淨效果與其組成（三情境 FY2032E；十億日圓）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
            "rows": [[f"@PL!A{PB[k]}"] + [f"@PL!L{PR[s][k]}|n1" for s in "ABC"] for k in keys]}


std_keys = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pti", "tax", "ni", "nm", "eps"]
unit_fmt = {"%": "p1", "日圓": "n2"}


def row_std(k):
    r = IX[k]
    u = wb["PL"].cell(r, 2).value
    f = "p1" if u == "%" else ("n2" if k == "eps" else "n1")
    return [f"@PL!A{r}"] + [f"@PL!{c}{r}|{f}" for c in C]


head_std = [f"@PL!A{STDH}"] + [f"@PL!{c}{STDH}" for c in C]
sec = {}
cfg = {
    "company": "LY Corporation", "ticker": "4689", "layout": "v1.3", "wide_text": True,
    "thesis": "@README!A3",
    "question": "FY2032E 的 AI 淨效果（營業利益）在三情境是否為正；基準下 AI 是否足以抵銷傳統搜尋與展示廣告的衰退",
    "scenarios": [{"key": "A", "name": "情境 A（分流小、價格提升大）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（聊天介面大量接走查詢）"}],
    "hero": {"rec": vv("rec"), "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n2"), ["上檔空間 ", vv("up", "B", "sp1")]],
                                       [f"@Val!A{VR['tpcomp']}", vv("tpcomp", "B", "n2"), ["上檔空間 ", vv("upc", "B", "sp1")]],
                                       [f"@Val!A{VR['px']}", vv("px", "B", "n2"), ""]],
             "note": ["@README!A11"]},
    "sections": [],
}
# ---------- 1. 命題 ----------
thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "note", "parts": ["@README!A5"]},
    {"type": "theme", "title": "營收（十億日圓，依分部堆疊）與營業利益率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{HR}:L{HR}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": f"@PL!C{PB['rev']}:L{PB['rev']}",
     "bars": [{"name": [f"@PL!A{PB['media']}"], "y": f"@PL!C{PB['media']}:L{PB['media']}", "color": "#6fa8dc"},
              {"name": [f"@PL!A{PB['com']}"], "y": f"@PL!C{PB['com']}:L{PB['com']}", "color": "#0b7a80"},
              {"name": [f"@PL!A{PB['str']}"], "y": f"@PL!C{PB['str']}:L{PB['str']}", "color": "#c29a2b"}],
     "lines": [{"name": [f"@PL!A{PB['om']}"], "y": f"@PL!C{PB['om']}:L{PB['om']}", "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝三個官方分部（Other 與 Adjustments 金額小，未畫出，合計以標準損益表為準）；折線＝合併營業利益率（右軸）。FY2026E 起為情境 B。"},
    {"type": "h3", "text": "命題與核心問題"},
    {"type": "list", "items": ["@README!A3", "@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [f"@README!A{r}" for r in (6, 7, 8, 9, 10)]},
    {"type": "h3", "text": "情境分化依據"},
    {"type": "note", "parts": ["@README!A13"]}]}
# ---------- 2. 估值 ----------
conc = [VR[k] for k in ("tpdcf", "tpcomp", "tp", "px", "cons", "up", "upc", "tier", "gate_chk", "rec")]
valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（日圓；豎線＝現價）", "fmt": "n0",
     "items": [{"label": f"@Val!A{VR[k]}", "value": f"@Val!B{VR[k]}", "series": sr} for k, sr in (("tpdcf", 1), ("tpcomp", 2), ("tp", 0), ("cons", 1))],
     "ref": {"label": f"@Val!A{VR['px']}", "value": f"@Val!B{VR['px']}"}},
    {"type": "note", "parts": ["@Val!A2"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值", "說明"],
     "rows": [[f"@Val!A{VR['tpdcf']}", vv("tpdcf", "B", "n2"), f"@Val!E{VR['tpdcf']}"], [f"@Val!A{VR['tpcomp']}", vv("tpcomp", "B", "n2"), f"@Val!E{VR['tpcomp']}"],
              [f"@Val!A{VR['tp']}", vv("tp", "B", "n2"), f"@Val!E{VR['tp']}"], [f"@Val!A{VR['px']}", vv("px", "B", "n2"), f"@Val!E{VR['px']}"],
              [f"@Val!A{VR['cons']}", vv("cons", "B", "n0"), f"@Val!E{VR['cons']}"], [f"@Val!A{VR['up']}", vv("up", "B", "sp1"), ""], [f"@Val!A{VR['upc']}", vv("upc", "B", "sp1"), ""],
              [f"@Val!A{VR['tier']}", vv("tier"), f"@Val!E{VR['tier']}"], [f"@Val!A{VR['valrate']}", vv("valrate"), f"@Val!E{VR['valrate']}"], [f"@Val!A{VR['gate_chk']}", vv("gate_chk"), f"@Val!E{VR['gate_chk']}"],
              [f"@Val!A{VR['gate_in']}", vv("gate_in", "B", "n1"), f"@Val!E{VR['gate_in']}"], [f"@Val!A{VR['rec']}", vv("rec"), f"@Val!E{VR['rec']}"]]},
    {"type": "kv", "items": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("rf", "p2"), ("beta", "n2"), ("erp", "p1"), ("ke", "p2"), ("kdat", "p2"), ("wacc", "p2"), ("ronic", "p2"), ("g", "p1"))]},
    {"type": "table", "caption": "DCF 彙總（三情境；十億日圓，每股為日圓）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['pvx']}-{VR['dcf12']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"],
                "row_fmt": {str(VR["tvsh"]): {c: "p1" for c in "BCD"}, str(VR["shv"]): {c: "n2" for c in "BCD"}, str(VR["dcf12"]): {c: "n2" for c in "BCD"}}}},
    {"type": "table", "caption": "分部加總（FY2027E 分部 EBITDA × 倍數；目標時點）", "head": ["分部", "情境 A", "情境 B", "情境 C", "倍數", "EBITDA 率"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['s_media']}-{VR['s_ps']}", "cols": ["A", "B", "C", "D", "E", "F"], "fmts": ["", "n1", "n1", "n1", "x1", "p1"],
                "row_fmt": {str(VR["s_ps"]): {c: "n2" for c in "BCD"}}}},
    {"type": "table", "caption": "DCF 與 Comparable 落差分解（情境 B，每股日圓）", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("g_tot", "n2"), ("g_mult", "n2"), ("g_lev", "n2"), ("g_time", "n2"), ("g_sbc", ""), ("g_chk", "n2"), ("g_evd", "n0"), ("g_imp", "n1"), ("g_sum", "n1"))]},
    {"type": "table", "caption": "本公司現行倍數與同業中位數（隱含重估）", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("c_ev", "n0"), ("c_eb", "n1"), ("c_mult", "n1"), ("c_peer", "n1"), ("c_ratio", "n2"), ("c_rel", "n2"), ("c_rr", "sp1"))]},
    {"type": "table", "caption": "同業 EV／EBITDA（TTM；Analogy；逾期＝資料日與現價基準日相差超過兩個交易日）", "head": ["公司", "對應分部", "會計準則", "租賃處理", "倍數", "資料日", "逾期"],
     "rowgen": {"sheet": "Peers", "rows": "4-14", "cols": ["A", "C", "D", "E", "H", "I", "J"], "fmts": ["", "", "", "", "x1", "", ""]}},
    {"type": "table", "caption": "歷史相對倍數 rel_mult（本公司 ÷ 同業中位數；自算，財報年度值）", "head": ["項目", "@Peers!B24", "@Peers!C24", "@Peers!D24", "@Peers!E24", "@Peers!F24", "中位數", "採用（夾值後）"],
     "rowgen": {"sheet": "Peers", "rows": "41-43", "cols": ["A", "B", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n2", "n2", "n2", "n2", "n2", "n2", "n2"]}},
    {"type": "table", "caption": "合理性關卡（發布前必查）", "head": ["項目", "數值", "狀態"],
     "rows": [[f"@Val!A{VR['z_dcf']}", vv("z_dcf", "B", "n2"), f"@Val!C{VR['z_dcf']}"], [f"@Val!A{VR['z_ev']}", vv("z_ev", "B", "n0"), f"@Val!C{VR['z_ev']}"], [f"@Val!A{VR['z_rec']}", vv("z_rec"), f"@Val!C{VR['z_rec']}"],
              [f"@Val!A{VR['z_rev']}", vv("z_rev", "B", "sp1"), f"@Val!C{VR['z_rev']}"], [f"@Val!A{VR['z_ebt']}", vv("z_ebt", "B", "sp1"), f"@Val!C{VR['z_ebt']}"], [f"@Val!A{VR['z_any']}", vv("z_any"), ""]]},
    {"type": "table", "caption": "逾期核對", "head": ["項目", "資料日", "狀態"],
     "rows": [[f"@Val!A{VR[k]}", f"@Val!B{VR[k]}", f"@Val!C{VR[k]}"] for k in ("o_rf", "o_peer", "o_cons", "o_ly")]},
    {"type": "table", "caption": "情境目標價與收支關卡", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("t_A", "n2"), ("t_B", "n2"), ("t_C", "n2"), ("wsum", "p0"), ("per", "n1"), ("gate_A", "n0"), ("fcf_gate", "n0"), ("gate_C", "n0"))]}]}
# ---------- 3. 敏感度 ----------
_main = [b for b in sens_blocks if not b.get("collapse") and b["type"] == "tornado"]
sens = {"id": "sens", "title": "敏感度", "blocks": _main + [{"type": "note", "parts": ["@Sens!A2"]}] + [b for b in sens_blocks if b not in _main]}
# ---------- 4. 反向 DCF ----------
revdcf = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "blocks": [
    {"type": "note", "parts": ["@RevDCF!A2"]}, {"type": "note", "parts": [f"@RevDCF!A{RVR['sum']}"]},
    {"type": "table", "caption": "現金流折現法（情境 B 營收路徑）：現價隱含值與模型", "head": ["項目", "現價隱含", "模型情境 B"],
     "rows": [[f"@RevDCF!A{RVR['m1']}", f"@RevDCF!B{RVR['m1']}|p1", f"@RevDCF!B{RVR['m0']}|p1"], [f"@RevDCF!A{RVR['evr']}", f"@RevDCF!B{RVR['evr']}|n0", "—"],
              [f"@RevDCF!A{RVR['rv0']}", "—", f"@RevDCF!B{RVR['rv0']}|n0"], [f"@RevDCF!A{RVR['w0']}", "—", f"@RevDCF!B{RVR['w0']}|p2"]]},
    {"type": "table", "caption": "分部加總法：現價隱含的倍數與模型", "head": ["項目", "數值"],
     "rows": [[f"@RevDCF!A{RVR[k]}", f"@RevDCF!B{RVR[k]}|{f}"] for k, f in (("smul", "n1"), ("smod", "n1"), ("sps", "n2"))]}]}
# ---------- 5. 損益 ----------
pl = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "note", "parts": ["損益表採 IFRS 口徑（GAAP 口徑，非 Non-GAAP 調整後數字）。", f"@PL!A{rm['STD_NOTE']}"]},
    {"type": "theme", "title": "營收（十億日圓）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{STDH}:L{STDH}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
     "bars": [{"name": [f"@PL!A{IX['rev']}"], "y": f"@PL!C{IX['rev']}:L{IX['rev']}", "color": "#6fc7cb"}],
     "lines": [{"name": [f"@PL!A{IX['gm']}"], "y": f"@PL!C{IX['gm']}:L{IX['gm']}", "axis": "r", "color": "#0b7a80", "labels": True},
               {"name": [f"@PL!A{IX['opm']}"], "y": f"@PL!C{IX['opm']}:L{IX['opm']}", "axis": "r", "color": "#c29a2b"},
               {"name": [f"@PL!A{IX['nm']}"], "y": f"@PL!C{IX['nm']}:L{IX['nm']}", "axis": "r", "color": "#6f6bae", "dash": True}],
     "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，IFRS）。LY 的營業成本只含部分成本，多數費用在營業費用，毛利率偏高；FY2025A 淨利率含一次性稅務與併購再衡量利益。"},
    {"type": "table", "caption": "標準損益表（IFRS；十億日圓，EPS 為日圓）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B", "head": head_std,
     "rows": [row_std(k) for k in std_keys] + [{"sub": f"@PL!A{rm['STD_NOTE'] - 6}"}] + [[f"@PL!A{RID[k]}"] + [f"@PL!{c}{RID[k]}|p1" for c in C] for k in ("r_rev", "r_gm", "r_opm", "r_nm", "r_niy")]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線）", "years": f"@PL!F{HR}:L{HR}",
     "metrics": [met("合併營收（十億日圓）", "rev", "n0", True), met("營業利益（十億日圓）", "oi", "n0", True), met("營業利益率", "om", "p1"), met("EPS（日圓）", "eps", "n2", True),
                 met("AI 淨效果：營業利益（十億日圓）", "ai_oi", "n1", True)]},
    {"type": "table", "caption": "分部對帳（差額應為零）", "head": ["項目", "最大絕對差（十億日圓）"],
     "rowgen": {"sheet": "PL", "rows": f"{rm['PLSUM'] + 1}-{rm['PLSUM'] + 6}", "cols": ["A", "C"], "fmts": ["", "n3"]}}]}
# ---------- 6. 驅動因子 ----------
drivers = {"id": "drivers", "title": "驅動因子：分部營收推導", "blocks": [
    {"type": "note", "parts": ["@PL!A2"]},
    {"type": "theme", "title": "搜尋廣告營收（長條）× 情境 A／B／C（十億日圓）", "x": f"@PL!F{HR}:L{HR}", "fmt": "n0", "stack": False, "pad": 1.2,
     "bars": [{"name": ["情境 ", s, "：搜尋廣告"], "y": f"@PL!F{PR[s]['srch']}:L{PR[s]['srch']}", "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
     "caption": "搜尋廣告＝基礎營收 ×（扣除查詢分流占比）×（加上價格提升）；三情境只改分流與價格提升，基礎成長相同。"},
    {"type": "table", "caption": "分部營收推導（情境 B；十億日圓）：子線加總＝分部，分部加總＝合併營收", "head": ["項目"] + [f"@PL!{c}{HR}" for c in C],
     "rowgen": {"sheet": "PL", "rows": f"{PB['s_base']}-{PB['chk_seg']}", "cols": ["A"] + C, "fmts": [""] + ["n1"] * 10}},
    scn_table(["ai_inc", "ai_ero", "inf", "ai_net", "ai_oi"]),
    {"type": "table", "caption": "三情境驅動（AI 驅動並排，進加權目標價；In 頁情境驅動對照）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{rm['SCN_ROWS'][k]}", f"@In!D{rm['SCN_ROWS'][k]}|p0", f"@In!E{rm['SCN_ROWS'][k]}|p0", f"@In!F{rm['SCN_ROWS'][k]}|p0"] for k in ("scn_div", "scn_lift", "scn_aim", "scn_agp", "scn_w")]},
    {"type": "table", "caption": "單一參數區間（基準進主模型；敏感度低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[f"@In!B{INR[k]}", f"@In!D{INR[k]}|{f}", f"@In!E{INR[k]}|{f}", f"@In!F{INR[k]}|{f}", f"@In!G{INR[k]}|tag"]
              for k, f in (("qg26", "p1"), ("qg32", "p1"), ("ag26", "p1"), ("ag32", "p1"), ("dg26", "p1"), ("cg26", "p1"), ("sg26", "p1"), ("ag_price", "n0"), ("tok_chat", "n0"), ("ptok", "n3"), ("fx", "n0"), ("ai_share", "p0"),
                           ("opm26", "p1"), ("opm32", "p1"), ("m_inc", "p0"), ("m_lost", "p0"), ("beta", "n2"), ("erp", "p1"), ("g_tv", "p1"), ("m_media", "x1"), ("m_com", "x1"), ("m_str", "x1"), ("w_dcf", "p0"))]}]}
# ---------- 7. 對帳 ----------
recon = {"id": "recon", "title": "對帳：公司指引與市場共識", "blocks": [
    {"type": "note", "parts": ["@Recon!A2"]},
    {"type": "table", "caption": "公司指引 vs 模型情境 B（公司指引＝Interested-party）", "head": ["項目", "年度", "模型情境 B", "公司指引", "差距", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{RCR['rev']}-{RCR['eps']}", "cols": ["A", "B", "C", "D", "E", "F"], "fmts": ["", "", "n1", "n1", "sp1", ""]}},
    {"type": "table", "caption": "反解：整本重算，使加權目標價等於現價所需的主導變數", "head": ["項目", "數值", "對照", "差距", "說明"],
     "rows": [[f"@Recon!A{RCR['sv_base']}", f"@Recon!C{RCR['sv_base']}|p2", "", "", f"@Recon!F{RCR['sv_base']}"],
              [f"@Recon!A{RCR['sv_x']}", f"@Recon!C{RCR['sv_x']}|p2", "", "", f"@Recon!F{RCR['sv_x']}"],
              [f"@Recon!A{RCR['sv_d']}", f"@Recon!C{RCR['sv_d']}|n2", "", "", ""],
              [f"@Recon!A{RCR['sv_tp']}", f"@Recon!C{RCR['sv_tp']}|n2", f"@Recon!D{RCR['sv_tp']}|n2", f"@Recon!E{RCR['sv_tp']}|sp2", f"@Recon!F{RCR['sv_tp']}"]]},
    {"type": "table", "caption": "上半年實績與共識", "head": ["項目", "數值", "對照", "差距", "說明"],
     "rows": [[f"@Recon!A{RCR['h1v']}", f"@Recon!C{RCR['h1v']}|n1", "", "", f"@Recon!F{RCR['h1v']}"], [f"@Recon!A{RCR['h1v'] + 1}", f"@Recon!C{RCR['h1v'] + 1}|p1", "", "", f"@Recon!F{RCR['h1v'] + 1}"],
              [f"@Recon!A{RCR['tpc']}", f"@Recon!C{RCR['tpc']}|n2", f"@Recon!D{RCR['tpc']}|n0", f"@Recon!E{RCR['tpc']}|sp1", f"@Recon!F{RCR['tpc']}"]]}]}
# ---------- 8. 資料狀態 ----------
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["build_id", f"@Sens!D{rm['SE_BID']}"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "h3", "text": "本版要點與限制"},
    {"type": "list", "items": ["@README!A2", "@README!A14", "@README!A15", "@README!A16", "@README!A17", "@README!A41", "@README!A42"]},
    {"type": "h3", "text": "市場資料日期與失效條件"},
    {"type": "list", "items": ["@README!A29", "@README!A31"]},
    {"type": "table", "caption": "風險登錄", "head": ["風險", "是否建模", "理由", "觀察指標"], "rowgen": {"sheet": "README", "rows": "34-38", "cols": ["A", "B", "C", "D"]}},
    {"type": "table", "collapse": True, "summary": "展開：驅動 → 推導對照表", "head": ["頁", "輸入", "推導量", "因果方向"], "rowgen": {"sheet": "README", "rows": "20-26", "cols": ["A", "B", "C", "D"]}},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A", "F"], "In": ["B", "H", "I"]}, "empty": "沒有標為未找到的項目。"}]}
cfg["sections"] = [thesis, valuation, sens, revdcf, pl, drivers, recon, status]
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

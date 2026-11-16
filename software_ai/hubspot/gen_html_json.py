#!/usr/bin/env python3
"""由成品 xlsx 與列號對照（rowmap.json）產生 html.json（個股 HTML 版面規格 v1.3；列號隨模型變動，勿手改 html.json）。
用法：python3 gen_html_json.py <成品 xlsx> <rowmap.json> html.json
文字中不含數字（年份與日期除外）：所有數字取自儲存格。
"""
import json
import sys

import openpyxl

X, RM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
rm = json.load(open(RM, encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)
VR, RVA, RCA, SENS, PB, CB, TB, SIS, Q4 = (rm[k] for k in ("VR", "RVA", "RCA", "SENS", "PB", "CB", "TB", "SIS", "Q4"))
SCN, SCN_HDR, KEYROW = rm["SCN"], rm["SCN_HDR"], rm["keyrow"]
YC = {y: openpyxl.utils.get_column_letter(3 + i) for i, y in enumerate(range(2023, 2033))}


def find(sh, text, col=1, start=1):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        if str(ws.cell(r, col).value or "").startswith(text):
            return r
    raise KeyError((sh, text))


def rng(sh, r, c1, c2):
    return f"@{sh}!{c1}{r}:{c2}{r}"


def trow(sh, r, cols, fmts):
    return [f"@{sh}!{c}{r}" + (f"|{f}" if f else "") for c, f in zip(cols, fmts)]


YR = SIS["yr"]
hdr = [f"@PL!{c}{YR}" for c in "ACDEFGHIJKL"]
sis_rows = [("rev", "n1"), ("g", "p1"), ("cogs", "n1"), ("gp", "n1"), ("gm", "p1"), ("opx", "n1"), ("ebit", "n1"), ("ebitm", "p1"), ("oth", "n1"), ("pt", "n1"), ("tax", "n1"),
            ("ni", "n1"), ("nim", "p1"), ("eps", "n2"), ("nig", "p1"), ("ngeps", "n2"), ("ngop", "n1"), ("ngopm", "p1")]
cols10 = list("CDEFGHIJKL")


def met(name, key, fmt, chart=False, ymin=None, sheet="PL", d=PB):
    m = {"name": name, "fmt": fmt, "rows": {s: rng(sheet, d[s][key], "F", "L") for s in "ABC"}}
    if chart:
        m["chart"] = True
    if ymin:
        m["ymin"] = ymin
    return m


# ───────── Sens 龍捲風 ─────────
SW = wb["Sens"]
t0 = SENS["t0"]
r1_, r2_ = t0 + 2, t0 + 1
while SW.cell(r2_ + 1, 1).value:
    r2_ += 1
tors = []
for col, side in (("B", "低"), ("C", "高")):
    rows = [r for r in range(r1_, r2_ + 1) if isinstance(SW[f"{col}{r}"].value, (int, float)) and abs(SW[f"{col}{r}"].value) >= 0.0005]
    tors.append({"type": "tornado", "title": f"加權目標價相對基準：輸入取{side}值（一次改一個輸入，整本重算）", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1",
                 "rows": ",".join(map(str, rows))})
rd = wb["README"]
r_obs = [6, 7, 8, 9]
assert sum(len(str(rd.cell(r, 1).value or "")) for r in r_obs) <= 740, "重點觀察超過字數"

cfg = {
    "company": "HubSpot",
    "ticker": "HUBS",
    "thesis": "@README!A3",
    "question": "席位被 AI 代理取代的侵蝕是否大於按任務計價的增量；現價隱含什麼；收支關卡通過與否",
    "scenarios": [{"key": "A", "name": "情境 A（席位取代慢、agent 放量）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（席位取代快、agent 慢）"}],
    "wide_text": True,
    "layout": "v1.3",
    "hero": {"rec": f"@Val!B{VR['rec']}",
             "tiles": [[f"@Val!A{VR['tp']}", f"@Val!B{VR['tp']}|n2", ["上檔空間 ", f"@Val!B{VR['up']}|sp1"]],
                       [f"@Val!A{VR['tpcomp']}", f"@Val!B{VR['tpcomp']}|n2", ["上檔空間 ", f"@Val!B{VR['upc']}|sp1"]],
                       [f"@Val!A{VR['px']}", f"@Val!B{VR['px']}|n2", ""]],
             "note": ["@README!A56"]},
    "sections": [],
}
S = cfg["sections"]
# ───────── 1. 命題 ─────────
th_chart = {"type": "theme", "title": "營收（百萬美元，依營收線堆疊）與非 GAAP 營業利益率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{PB['B']['yr']}:L{PB['B']['yr']}",
            "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3, "total": f"@PL!C{PB['B']['rev']}:L{PB['B']['rev']}",
            "bars": [{"name": [f"@PL!A{PB['B']['sub']}"], "y": f"@PL!C{PB['B']['sub']}:L{PB['B']['sub']}", "color": "#6fa8dc"},
                     {"name": [f"@PL!A{PB['B']['ps']}"], "y": f"@PL!C{PB['B']['ps']}:L{PB['B']['ps']}", "color": "#c29a2b"}],
            "lines": [{"name": [f"@PL!A{PB['B']['ngopm']}"], "y": f"@PL!G{PB['B']['ngopm']}:L{PB['B']['ngopm']}", "axis": "r", "color": "#3b3838", "labels": True}],
            "caption": "長條＝訂閱與專業服務營收；折線＝非 GAAP 營業利益率（右軸，公司自訂口徑；歷史年度公司只揭露營業利益，未列率）。FY2026E 起為情境 B。"}
S.append({"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "note", "parts": ["@README!A5"]}, th_chart,
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [f"@README!A{r}" for r in r_obs]},
    {"type": "h3", "text": "命題回覆"},
    {"type": "list", "items": [f"@README!A{r}" for r in (45, 46, 47)]},
    {"type": "h3", "text": "本版要點"},
    {"type": "list", "items": ["@README!A2"]}]})
# ───────── 2. 估值 ─────────
rf_f = {str(VR[k]): {c: f for c in "BCD"} for k, f in (("up", "sp1"), ("upc", "sp1"), ("updcf", "sp1"), ("tvsh2", "p1"), ("san_dcf", "n2"), ("san_g", "p1"), ("san_m27", "sp1"), ("bdist", "n1"),
                                                      ("gap", "n2"), ("gap_sbc", "n2"), ("ab_ps0", "n2"), ("ab_tp", "n2"), ("ab_up", "sp1"), ("gap_time", "n2"), ("gap_lev", "n2"), ("gap_mult", "n2"), ("pe_hubs_now", "n1"), ("gatev", "n0"),
                                                      ("gate_all", "n0"), ("per", "n2"), ("perel", "n2"), ("uattr", "sp1"), ("dep", "n2"), ("san_ev", "n0"), ("tpdcf", "n2"), ("tpcomp", "n2"), ("tp", "n2"),
                                                      ("px", "n2"), ("cons", "n2"))}
for k in ("tier", "rec_val", "rec", "san_rec", "san_diag"):
    rf_f[str(VR[k])] = {c: "" for c in "BCD"}
rf_f[str(VR["ab_up"])] = {"B": "sp1", "C": "", "D": ""}
rf_f[str(VR["ab_ps0"])] = {"B": "n2", "C": "", "D": ""}
rf_f[str(VR["ab_tp"])] = {"B": "n2", "C": "", "D": ""}
rf_f[str(VR["perel"])] = {"B": "n2", "C": "sp1", "D": ""}
rf_f[str(VR["bdist"])] = {"B": "n1", "C": "", "D": ""}
rf_f[str(VR["san_dcf"])] = {"B": "n2", "C": "", "D": ""}
rf_f[str(VR["san_ev"])] = {"B": "n0", "C": "", "D": ""}
rf_f[str(VR["san_g"])] = {"B": "p1", "C": "", "D": ""}
rf_f[str(VR["per"])] = {"B": "n2", "C": "n2", "D": ""}
rf_f[str(VR["dep"])] = {"B": "n2", "C": "n2", "D": ""}
rf_f[str(VR["uattr"])] = {"B": "sp1", "C": "sp1", "D": "sp1"}
rf_f[str(VR["tvsh2"])] = {c: "p1" for c in "BCD"}
rf_f[str(VR["gate_all"])] = {c: "n0" for c in "BCD"}
vw = rm["vw"]
S.append({"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{VR[k]}", "value": f"@Val!B{VR[k]}", "series": sr} for k, sr in (("tpdcf", 1), ("tpcomp", 2), ("tp", 0), ("cons", 1))],
     "ref": {"label": f"@Val!A{VR['px']}", "value": f"@Val!B{VR['px']}"}},
    {"type": "note", "parts": ["@Val!A1"]}, {"type": "note", "parts": ["@Val!A2"]}, {"type": "note", "parts": ["@Val!A3"]},
    {"type": "table", "caption": "@Val!A4", "head": ["項目", "數值", "對照", "對照", "說明"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['tpdcf']}-{VR['san_diag']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""], "row_fmt": rf_f}},
    {"type": "note", "parts": [f"@Recon!A{RCA['ab_txt']}"]},
    {"type": "table", "caption": "DCF 讀數並列：股票薪酬費用化（進加權目標價）與加回（只作對照）", "head": ["項目", "費用化（基準）", "加回（對照）", "差額"],
     "rows": [[f"@Recon!A{RCA[k]}", f"@Recon!B{RCA[k]}|{f}", f"@Recon!C{RCA[k]}|{f}", f"@Recon!D{RCA[k]}|{f}"] for k, f in (("ab_ps", "n2"), ("ab_tp", "n2"), ("ab_up", "sp1"))]},
    {"type": "table", "caption": "折現率與股數", "head": ["項目", "數值"],
     "rowgen": {"sheet": "Val", "rows": f"{vw['rf']}-{vw['sh_val']}", "cols": ["A", "B"], "fmts": ["", "n3"],
                "row_fmt": {str(vw[k]): {"B": "p2"} for k in ("rf", "erp", "ke", "kd", "wacc", "ronic")} | {str(vw["tax"]): {"B": "p1"}, str(vw["g"]): {"B": "p1"}, str(vw["wd"]): {"B": "p2"},
                                                                    str(vw["mc"]): {"B": "n0"}}}},
    {"type": "table", "caption": "DCF 彙總（三情境）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dsum0']}-{VR['dsum0'] + 8}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"],
                "row_fmt": {str(VR["dsum0"] + 4 + 1): {c: "n3" for c in "BCD"}, str(VR["dsum0"] + 6): {c: "n2" for c in "BCD"}, str(VR["dsum0"] + 7): {c: "n2" for c in "BCD"},
                            str(VR["dsum0"] + 8): {c: "p1" for c in "BCD"}}}},
    {"type": "table", "caption": "Comparable（同業前瞻 P/E 中位數 × 目標時點 NTM 非 GAAP EPS）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['cmp0']}-{VR['cmp0'] + 5}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n2", "n2", "n2"],
                "row_fmt": {str(VR["cmp0"] + 2): {c: "n3" for c in "BCD"}, str(VR["cmp0"] + 4): {c: "n1" for c in "BCD"}}}},
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻 P/E（逐家）", "head": ["同業", "前瞻 P/E（倍）"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['peer0']}-{VR['peer_hubs']}", "cols": ["A", "B"], "fmts": ["", "x1"]}}]})
# ───────── 3. 敏感度 ─────────
S.append({"id": "sens", "title": "敏感度", "blocks": tors + [
    {"type": "note", "parts": ["@Sens!A1"]},
    {"type": "note", "parts": ["@Sens!A2"]},
    {"type": "table", "collapse": True, "summary": "展開：全部案例（含 DCF 單法與 Comparable 單法）", "head": ["驅動", "設定", "方向", "加權目標價", "相對基準", "DCF 單法", "Comparable 單法"],
     "rowgen": {"sheet": "Sens", "rows": f"{SENS['first']}-{SENS['last']}", "cols": ["A", "B", "C", "D", "E", "H", "I"], "fmts": ["", "", "", "n2", "sp1", "n2", "n2"]}},
    {"type": "kv", "items": [["模型狀態", f"@Sens!D{SENS['status']}"], ["建置識別碼", f"@Sens!D{SENS['bid']}"]]}]})
# ───────── 4. 反向 DCF ─────────
RV = RVA
S.append({"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2", "blocks": [
    {"type": "note", "parts": [f"@RevDCF!A{RV['sum_dcf']}"]}, {"type": "note", "parts": [f"@RevDCF!A{RV['sum_cmp']}"]},
    {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
     "rows": [[f"@RevDCF!A{RV['m1']}", f"@RevDCF!B{RV['m1']}|p1", f"@RevDCF!B{RV['m0']}|p1"], ["營收規模（FY2028 起各年同乘的倍數；模型為一）", f"@RevDCF!B{RV['s1']}|n2", "—"],
              ["FY2028 營收（百萬美元）", f"@RevDCF!B{RV['rv1']}|n0", f"@RevDCF!B{RV['rv0']}|n0"], ["FY2028 營收年增率", f"@RevDCF!B{RV['g1']}|p1", f"@RevDCF!B{RV['g0']}|p1"],
              ["折現率（WACC；無解時為搜尋下界）", f"@RevDCF!B{RV['w1']}|p2", f"@RevDCF!B{RV['w0']}|p2"], [f"@RevDCF!A{RV['w_flag']}", f"@RevDCF!B{RV['w_flag']}", ""]]},
    {"type": "table", "caption": "Comparable：現價隱含的倍數與模型", "head": ["項目", "現價隱含", "模型", "差距（現價隱含相對模型）"],
     "rows": [[f"@RevDCF!A{RV[k]}", f"@RevDCF!B{RV[k]}|{f}", f"@RevDCF!C{RV[k]}|{f}", f"@RevDCF!D{RV[k]}|sp1"] for k, f in (("sps", "n2"), ("spe", "n1"), ("scmp", "n2"))]},
    {"type": "table", "caption": "現金流折現法：現價隱含的企業價值與模型", "head": ["項目", "值"],
     "rows": [[f"@RevDCF!A{RV[k]}", f"@RevDCF!B{RV[k]}|{f}"] for k, f in (("px", "n2"), ("v0", "n2"), ("evr", "n0"), ("ev", "n0"), ("evx", "n2"), ("v_b", "n2"))]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）", "head": [f"@RevDCF!{c}{RV['yh']}" for c in "ABCDEFGH"],
     "rows": [[f"@RevDCF!A{RV['y_' + k]}"] + [f"@RevDCF!{c}{RV['y_' + k]}|{f}" for c in "BCDEFGH"] for k, f in (("rev", "n1"), ("ebit", "n1"), ("m", "p1"), ("df", "n3"), ("tvf", "n3"), ("k", "n3"))]}]})
# ───────── 5. 損益 ─────────
pl_chart = {"type": "theme", "title": "營收（百萬美元）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B，GAAP）", "x": f"@PL!C{SIS['yr']}:L{SIS['yr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
            "bars": [{"name": [f"@PL!A{SIS['rev']}"], "y": f"@PL!C{SIS['rev']}:L{SIS['rev']}", "color": "#6fc7cb"}],
            "lines": [{"name": [f"@PL!A{SIS['gm']}"], "y": f"@PL!C{SIS['gm']}:L{SIS['gm']}", "axis": "r", "color": "#0b7a80", "labels": True},
                      {"name": [f"@PL!A{SIS['ebitm']}"], "y": f"@PL!C{SIS['ebitm']}:L{SIS['ebitm']}", "axis": "r", "color": "#c29a2b"},
                      {"name": [f"@PL!A{SIS['nim']}"], "y": f"@PL!C{SIS['nim']}:L{SIS['nim']}", "axis": "r", "color": "#6f6bae", "dash": True}],
            "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP；股票薪酬費用化）。FY2026E 含第四季重組費用。"}
chk_r = [PB[s]["chk"] for s in SC] if (SC := "ABC") else []
S.append({"id": "pl", "title": "損益（標準損益表）", "intro": "@PL!A2", "blocks": [
    pl_chart,
    {"type": "table", "caption": f"@PL!A{SIS['title']}", "head": hdr, "rows": [[f"@PL!A{SIS[k]}"] + [f"@PL!{c}{SIS[k]}|{f}" for c in cols10] for k, f in sis_rows]},
    {"type": "note", "parts": [f"@PL!A{SIS['title']}"]},
    {"type": "pl", "years": rng("PL", PB["B"]["yr"], "F", "L"), "metrics": [
        met("營收（百萬美元）", "rev", "n0", True), met("訂閱營收（百萬美元）", "sub", "n0", True), met("營業利益率（GAAP）", "ebitm", "p1", True, ymin="auto"),
        met("非 GAAP 營業利益率", "ngopm", "p1", False), met("EPS（GAAP 稀釋，美元）", "eps", "n2", True, ymin="auto"), met("非 GAAP EPS（公司口徑，美元）", "ngeps", "n2"),
        {"name": "自由現金流（百萬美元）", "fmt": "n0", "rows": {s: rng("CF", CB[s]["fcf"], "F", "L") for s in "ABC"}}]},
    {"type": "table", "caption": "分部對帳（三情境；訂閱＋專業服務 − 營收合計，應為零）", "head": ["項目"] + [f"@PL!{c}{PB['B']['yr']}" for c in cols10],
     "rows": [[f"@PL!A{PB[s]['chk']}"] + [f"@PL!{c}{PB[s]['chk']}|n3" for c in cols10] for s in "ABC"]},
    {"type": "table", "caption": "收支關卡（FY2026–FY2030 自由現金流減股利；三情境）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [["@CF!A3", "@CF!B3|n0", "@CF!C3|n0", "@CF!D3|n0"], ["@CF!A4", "@CF!B4|n0", "@CF!C4|n0", "@CF!D4|n0"]]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 的現金流與 DCF 明細", "head": ["項目"] + [f"@CF!{c}{CB['B']['yr']}" for c in "FGHIJKL"],
     "rows": [[f"@CF!A{CB['B'][k]}"] + [f"@CF!{c}{CB['B'][k]}|{f}" for c in "FGHIJKL"] for k, f in (("rev", "n1"), ("ebit", "n1"), ("nopat", "n1"), ("da", "n1"), ("cx", "n1"), ("wc", "n1"), ("fcff", "n1"), ("fcf", "n1"),
                                                                                                      ("bb", "n1"), ("cash", "n1"), ("t", "n2"), ("df", "n3"), ("pv", "n1"))]}]})
# ───────── 6. 驅動 ─────────
TBB = TB["B"]
net_chart = {"type": "theme", "title": "AI 淨效果（長條，百萬美元）：按任務計價營收減席位取代損失與增量推論成本 × 情境 A／B／C", "x": f"@Thesis!G{TBB['yr']}:L{TBB['yr']}", "fmt": "n0", "stack": False, "pad": 1.2,
             "bars": [{"name": ["情境 ", s, "：AI 淨效果"], "y": f"@Thesis!G{TB[s]['net']}:L{TB[s]['net']}", "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
             "lines": [{"name": ["情境 ", s, "：按任務計價營收"], "y": f"@Thesis!G{TB[s]['aiRev']}:L{TB[s]['aiRev']}", "color": c, "dash": s != "B", "labels": s == "B"} for s, c in (("A", "#0b7a80"), ("B", "#c29a2b"), ("C", "#b4472f"))],
             "caption": "長條＝三情境逐年 AI 淨效果（負值＝侵蝕大於增量）；線＝按任務計價營收。席位被取代的損失是淨效果的主要來源。"}
inkey = lambda k: KEYROW[k]
param_keys = [k for k in KEYROW if wb["In"].cell(KEYROW[k], 4).value is not None or wb["In"].cell(KEYROW[k], 6).value is not None]
pfmt = lambda k: "p1" if wb["In"].cell(KEYROW[k], 3).value == "%" else ("n3" if wb["In"].cell(KEYROW[k], 5).value is not None and isinstance(wb["In"].cell(KEYROW[k], 5).value, float) and abs(wb["In"].cell(KEYROW[k], 5).value) < 10 else "n1")
S.append({"id": "drivers", "title": "驅動因子 → 推導量", "intro": "@Thesis!A1", "blocks": [
    net_chart,
    {"type": "table", "caption": "命題模組（情境 B）：客戶數、每客戶營收、AI 增量與侵蝕", "head": ["項目", "FY2025A"] + [f"@Thesis!{c}{TBB['yr']}" for c in "FGHIJKL"],
     "rows": [[f"@Thesis!A{TBB[k]}", f"@Thesis!E{TBB[k]}|{f}" if k == "cust" else ""] + [f"@Thesis!{c}{TBB[k]}|{f}" for c in "FGHIJKL"]
              for k, f in (("cust", "n0"), ("na", "p1"), ("arpc", "n0"), ("adopt", "p1"), ("tasks", "n0"), ("price", "n2"), ("aiRev", "n1"), ("loss", "n1"), ("inc", "n1"), ("net", "n1"), ("brk", "n2"), ("share", "p1"))]},
    {"type": "note", "parts": ["@Thesis!A2"]},
    {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；In 頁最後一節）", "head": [f"@In!{c}{SCN_HDR}" for c in "BDEFCG"],
     "rows": [[f"@In!B{SCN[k]}", f"@In!D{SCN[k]}|p1", f"@In!E{SCN[k]}|p1", f"@In!F{SCN[k]}|p1", f"@In!C{SCN[k]}", f"@In!G{SCN[k]}|tag"] for k in ("scn_repl", "scn_adopt", "scn_dp", "scn_w")]},
    {"type": "table", "caption": "單一參數區間（基準進主模型；低／高只用於敏感度頁、不進目標價；三情境共用）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!C3", "@In!G3"],
     "rows": [[f"@In!B{inkey(k)}", f"@In!D{inkey(k)}|{pfmt(k)}", f"@In!E{inkey(k)}|{pfmt(k)}", f"@In!F{inkey(k)}|{pfmt(k)}", f"@In!C{inkey(k)}", f"@In!G{inkey(k)}|tag"] for k in param_keys]},
    {"type": "table", "collapse": True, "summary": "展開：第四季 2026 預測（第一期現金流）", "head": ["項目", "第四季 2026E"],
     "rows": [[f"@Q4E!A{Q4[k]}", f"@Q4E!B{Q4[k]}|{f}"] for k, f in (("cust", "n0"), ("asrpc", "n0"), ("rev", "n1"), ("ngop", "n1"), ("sbc", "n1"), ("restr", "n1"), ("ebit", "n1"), ("fcff", "n1"), ("fcf", "n1"), ("ngeps", "n2"))]}]})
# ───────── 7. 對帳 ─────────
R = RCA
S.append({"id": "recon", "title": "對帳：指引、共識與反解", "intro": "@Recon!A1", "blocks": [
    {"type": "table", "caption": f"@Recon!A{R['q4rev'] - 2}", "head": ["項目", "模型（情境 B）", "指引或共識", "差距", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{R['q4rev']}-{R['tpc']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "sp1", ""]}},
    {"type": "table", "caption": f"@Recon!A{R['ab_ps'] - 2}", "head": ["項目", "費用化（基準）", "加回（對照）", "差額", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{R['ab_ps']}-{R['ab_up']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""], "row_fmt": {str(R["ab_up"]): {"B": "sp1", "C": "sp1", "D": "sp1"}}}},
    {"type": "table", "caption": f"@Recon!A{R['q4gap'] - 2}", "head": ["項目", "數值", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{R['q4gap']}-{R['m27']}", "cols": ["A", "B", "C"], "fmts": ["", "n2", "p1"], "row_fmt": {str(R["q4_implied"]): {"B": "p1"}, str(R["q4_pct"]): {"B": "sp1"}, str(R["m27"]): {"B": "p1", "C": "p1"}}}},
    {"type": "table", "caption": f"@Recon!A{R['k_key'] - 2}", "head": ["項目", "數值", "說明"],
     "rowgen": {"sheet": "Recon", "rows": f"{R['k_key']}-{R['k_up']}", "cols": ["A", "B", "C"], "fmts": ["", "p1", ""], "row_fmt": {str(R["k_tp"]): {"B": "n2"}}}},
    {"type": "table", "caption": f"@Recon!A{R['d_rf'] - 2}", "head": ["項目", "資料日", "與股價基準日相差（交易日）", "狀態"],
     "rowgen": {"sheet": "Recon", "rows": f"{R['d_rf']}-{R['d_px']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "date", "n0", ""]}},
    {"type": "table", "caption": f"@Recon!A{R['peer0'] - 2}", "head": ["同業", "前瞻 P/E（倍）", "會計準則與租賃處理", "備註"],
     "rowgen": {"sheet": "Recon", "rows": f"{R['peer0']}-{R['pe_rr']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "x1", "", ""], "row_fmt": {str(R["pe_ratio"]): {"B": "n2"}, str(R["pe_rr"]): {"B": "sp1"}}}},
    {"type": "table", "caption": f"@Recon!A{R['ai_rev'] - 2}", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rowgen": {"sheet": "Recon", "rows": f"{R['ai_rev']}-{R['ai_share']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n1", "n1", "n1"],
                "row_fmt": {str(R["ai_brk"]): {c: "n2" for c in "BCD"}, str(R["ai_price"]): {c: "n2" for c in "BCD"}, str(R["ai_share"]): {c: "p1" for c in "BCD"}}}}]})
# ───────── 8. 資料狀態 ─────────
S.append({"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期（資料截止日、成品日、股價基準日）", "@README!A1"], ["建置識別碼", f"@Sens!D{SENS['bid']}"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "h3", "text": "資料與口徑"},
    {"type": "list", "items": ["@In!A2", "@Hist!A1", "@Q4E!A1", "@Sens!A2"]},
    {"type": "h3", "text": "市場資料日期、失效條件與風險登錄"},
    {"type": "list", "items": [f"@README!A{r}" for r in (29, 30, 31, 32, 33, 34, 57, 58)]},
    {"type": "table", "caption": "風險登錄", "head": ["風險", "是否建模", "理由", "觀察指標"], "rowgen": {"sheet": "README", "rows": "38-42", "cols": ["A", "B", "C", "D"]}},
    {"type": "h3", "text": "未做與不適用"},
    {"type": "list", "items": [f"@README!A{r}" for r in (49, 50, 51, 52)]},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A", "B"], "Hist": ["A", "B"]}, "empty": "沒有標為未找到的項目。"}]})
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("html.json", OUT)

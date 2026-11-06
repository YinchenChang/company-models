# v0.4.2：DCF 彙總表格式依標籤（模組 v2 stub 期）。v0.4（2026-11-05）：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md），並配合 DCF 改標準版（FCFF＋WACC）：估值章 DCF 區塊由本檔依 xlsx 列號產生。
# 讀 innolight/html_src.json（v0.3 的 html.json，原樣保留）→ 少量前處理 → tools/company_html/relayout.py → 寫 innolight/html.json。
# 用法：python3 innolight/gen_html_json.py [成品 xlsx（預設 innolight/dist/ 內唯一的 xlsx）]
import glob, json, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
from relayout import relayout  # noqa: E402
X = sys.argv[1] if len(sys.argv) > 1 else (lambda f: f[0] if len(f) == 1 else sys.exit(f"dist 應恰有一個 xlsx：{f}"))(sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx"))))
cfg = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)
IN, VA = wb["In"], wb["Val"]
def row(ws, prefix, col=1):
    return next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, col).value or "").startswith(prefix))
ink = lambda k: next(r for r in range(1, IN.max_row + 1) if IN.cell(r, 1).value == k)
secs = {s["id"]: s for s in cfg["sections"]}
DC = wb["DCF"]
r_dtp = row(VA, "DCF 目標價（情境加權）")
_rk = row(VA, "股權成本 ke"); _rr = row(VA, "永續期 RONIC"); _rp = row(VA, "明確期 FCFF 現值")
_dh = row(DC, "項目")
# v0.4.2：DCF 彙總表格式依標籤（模組 v2 stub 期在「＋淨現金與投資」後多兩列，不寫死列偏移）
_fm = lambda r: (lambda lab: "n3" if lab.startswith("評價用稀釋股數") else "p1" if "終值現值占 EV" in lab else "x1" if "隱含終值" in lab else "n2")(str(VA.cell(r, 1).value or "").strip())
DCF_BLOCKS = [
    {"type": "note", "parts": [f"@Val!A{_rk - 2}"]},
    {"type": "table", "caption": "DCF 折現率：股權成本、稅後債務成本、WACC、永續成長率、RONIC", "head": ["項目", "數值", "說明"],
     "rows": [[f"@Val!A{r}", f"@Val!B{r}|" + ("p1" if r == _rk + 2 else "p2"), f"@Val!E{r}"] for r in range(_rk, _rr + 1)]},
    {"type": "table", "caption": "DCF 彙總：明確期現值、終值現值、企業價值、淨現金與每股價值", "head": [f"@Val!{c}{_rk - 1}" for c in "ABCD"],
     "rows": [[f"@Val!A{r}"] + [f"@Val!{c}{r}|" + _fm(r) for c in "BCD"] for r in range(_rp, r_dtp)]
             + [[f"@Val!A{r_dtp}", f"@Val!B{r_dtp}|n2", "", ""]]}]
for s_ in "ABC":
    _a = row(DC, f"情境 {s_}（"); _b = next(r for r in range(_a, DC.max_row + 1) if str(DC.cell(r, 1).value or "").startswith("現值（FCFF"))
    _cols = ["A"] + [chr(ord("C") + i) for i in range(6)]
    _pct = {str(r): {c: "p1" for c in _cols[1:]} for r in range(_a + 1, _b + 1) if str(DC.cell(r, 1).value).startswith(("稅率", "折舊攤銷 ÷", "資本支出 ÷"))}
    _pct.update({str(r): {c: "n3" for c in _cols[1:]} for r in range(_a + 1, _b + 1) if str(DC.cell(r, 1).value).startswith("折現因子")})
    DCF_BLOCKS.append({"type": "table", "caption": f"@DCF!A{_a}", "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細",
                       "head": [f"@DCF!{c}{_dh}" for c in _cols],
                       "rowgen": {"sheet": "DCF", "rows": f"{_a + 1}-{_b}", "cols": _cols, "fmts": [""] + ["n2"] * 6, "row_fmt": _pct}})
# v0.4：Sens 目標價參數表改為標準版 DCF 的十個案例（tools/fcff_dcf.py write_sens；起始列不變）
SE = wb["Sens"]
_st = row(SE, "目標價參數敏感度")
_se = next(r for r in range(_st + 2, SE.max_row + 1) if str(SE.cell(r, 1).value or "").startswith("檢查：閉式公式"))
for b in secs["sens"]["blocks"]:
    if b["type"] == "table" and b.get("caption") == f"@Sens!A{_st}":
        b["head"] = [f"@Sens!{c}{_st + 1}" for c in "ABCDEFG"]
        b["rowgen"] = {"sheet": "Sens", "rows": f"{_st + 2}-{_se - 1}", "cols": list("ABCDEFG"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2"]}

# 前處理（只動版面，不加數字）
# 0. 驅動：「情境驅動」表依 xlsx 的「情境驅動對照」節重建列號（v0.3 的 html.json 列號已與 In 頁錯位）
_s0 = row(IN, "情境驅動對照") + 1
_s1 = max(r for r in range(_s0 + 1, IN.max_row + 1) if str(IN.cell(r, 1).value or "").startswith("scn_"))
for b in secs["drivers"]["blocks"]:
    if str(b.get("caption", "")).startswith("情境驅動"):
        b["head"] = [f"@In!{c}{_s0}" for c in "BDEFC"]
        _uf = lambda r: {"%": "p1", "x": "n2"}.get(str(IN.cell(r, 3).value), "n2")
        b["rows"] = [[f"@In!B{r}"] + [f"@In!{c}{r}|{_uf(r)}" for c in "DEF"] + [f"@In!C{r}"] for r in range(_s0 + 1, _s1 + 1)]
# 1. 驅動：原「單一參數」表實為輸入 → 推導量對照，交給轉換器放到附錄
for b in secs["drivers"]["blocks"]:
    if str(b.get("caption", "")).startswith("單一參數"):
        b["caption"] = "驅動 → 推導量對照"
# 2. 估值：重點觀察已移到第一章；兩個 kv（相對倍數調整、Comparable 計算）改為表格（轉換器會把 kv 當成折現率表）；
#    DCF 折現率直接顯示、緊接在 DCF 計算表之前
va = secs["valuation"]["blocks"]
out, skip = [], False
r_rm, r_cmp, r_dcf = row(VA, "0.5 相對倍數調整"), row(VA, "同業前瞻 P/E 中位數"), row(VA, "0.1 DCF")
for b in va:
    if b["type"] == "h3" and b.get("text") == "重點觀察":
        skip = True
        continue
    if skip and b["type"] == "list":
        skip = False
        continue
    if b["type"] == "kv":
        first = b["items"][0][0]
        cap = "Comparable 計算" if first == f"@Val!A{r_cmp}" else "相對倍數調整"
        b = {"type": "table", "caption": cap, "head": ["項目", "數值"], "rows": [list(i) for i in b["items"]]}
    if b["type"] == "table" and b.get("caption") == f"@Val!A{r_dcf}":   # v0.4：舊 0.1 DCF（股東現金流）表 → 標準版 DCF 折現率、彙總與各情境 FCFF 明細
        out += DCF_BLOCKS
        continue
    if b["type"] == "bars":
        b = json.loads(json.dumps(b))
        for it in b["items"]:
            if it["label"] == "@Val!A24":
                it["label"], it["value"] = f"@Val!A{r_dtp}", f"@Val!B{r_dtp}"
    out.append(b)
secs["valuation"]["blocks"] = out

scen_tok = re.split(r"[.\s]", IN.cell(row(IN, "情境驅動對照"), 1).value)[0]
SPEC = {
 "in_skip": ["H", scen_tok],
 "in_groups": [
  ["AI 光模組（可插拔與 CPO）", [
    ["AI 光模組市場（讀網通市場模型；情境 A）", ["mpl_26", "mpl_27", "mpl_28", "mpl_29", "mcp_26", "mcp_27", "mcp_28", "mcp_29"]],
    ["網通 CPO 情境 B、C 的市場", ["mplB_26", "mplB_27", "mplB_28", "mplB_29", "mplC_26", "mplC_27", "mplC_28", "mplC_29",
                                  "mcpB_26", "mcpB_27", "mcpB_28", "mcpB_29", "mcpC_26", "mcpC_27", "mcpC_28", "mcpC_29"]],
    ["情境顆數係數", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["中際旭創份額、美國限制、CPO 份額與單價", ["rel_27", "rel_28", "rel_29", "us_27", "us_28", "us_29", "cpo_27", "cpo_28", "cpo_29", "pxr_27", "pxr_28", "pxr_29"]],
    ["匯率", ["fx"]],
    ["對帳參考（AI IT 資本支出）", ["itc_26", "itc_27", "itc_28", "itc_29"]],
    ["市場量敏感度係數（只進敏感度）", ["mf", "wta_27", "wta_28", "wta_29"]]]],
  ["其他營收（國內、電信等）", [
    ["其他營收占比與成長", ["oth_s", "g_oth"]]]],
  ["全公司營運：實績錨點、毛利率與費用", [
    ["FY2026 錨點（上半年與第三季實績、第四季假設）", ["h1_rev", "h1_gm", "h1_ni", "h1_eps", "q1_rev", "q_g", "q3_rev", "q3_gp", "q3_ni", "q3_eps", "sh_post", "q_sh_h1"]],
    ["FY2025 實績", ["a25_rev_in", "a25_gm_in", "a25_ni_in", "a25_sh_in"]],
    ["毛利率路徑", ["gm_h2", "gm_27", "gm_28", "gm_29"]],
    ["營業費用、稅、歸母比例與股數", ["phi", "opx_adj", "tx", "attr", "sh_g"]]]],
  ["全公司：股價、折現率、目標價規則與收斂法", [
    ["股價", ["px", "px_hi"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_wA", "tp_wB", "tp_wC", "tp_w_dcf"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "rf_us", "beta_u", "beta_u_med", "beta_du", "beta_de", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc", "fc_nonop", "fcf_da", "fcf_capex"]],
    ["收斂法（FY2030 起）", ["tp_g", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["共識目標價與前瞻 EPS（只作對照）", ["cons_tp", "cons_feps"]]]]],
 "blocks_to": {"中際旭創份額": "AI 光模組（可插拔與 CPO）", "@Units!A12": "AI 光模組（可插拔與 CPO）"},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "Comparable 計算": ["同業前瞻 P/E 中位數", "12 個月後的前瞻 EPS", "Comparable 目標價"],
          "sheet:Recon": ["FY2026 營收 vs 共識", "FY2027 EPS vs 共識", "FY2028 EPS vs 共識", "FY2027 歸母淨利 vs 高盛"]},
}
# 第三輪修正 15／21：情境驅動表依 In 頁 C 欄單位格式化（% 與 pt 用百分比，避免 0.06 顯示成小數）；DCF 彙總主表直接顯示 stub 兩列
def _uf3(r_):
    u = IN.cell(r_, 3).value
    return "p1" if u in ("%", "pt") else "n3" if u == "十億股" else "n2"
for b in secs["drivers"]["blocks"]:
    if b["type"] == "table" and str(b.get("caption", "")).startswith("情境驅動"):
        b["rows"] = [[x[0]] + [(f"{c}|{_uf3(int(c.split('!')[1][1:]))}" if "|" not in c and isinstance(IN[c.split('!')[1]].value, (int, float)) else c) for c in x[1:4]] + x[4:] for x in b["rows"]]
_m3 = SPEC["main"]["DCF 彙總"]
if "＋Stub 期 FCFF" not in _m3:
    _m3.insert(_m3.index("＋淨現金與投資") + 1, "＋Stub 期 FCFF")
cfg = relayout(cfg, X, SPEC)
secs = {s["id"]: s for s in cfg["sections"]}
# 後處理：轉換器把估值章所有長條圖移到最前面；舊法（規則值）長條圖放回舊法段落，避免與目標價並列
_vb = secs["valuation"]["blocks"]
_old = next(b for b in _vb if b["type"] == "bars" and b.get("title") == "規則值隱含股價與現價")
_vb.remove(_old)
_h = next(i for i, b in enumerate(_vb) if b["type"] == "h3" and str(b.get("text", "")).startswith("規則值隱含股價"))
_vb.insert(_h + 2, _old)
# 後處理：本模型的龍捲風圖是 FY2028 EPS（不是加權目標價），轉換器的通用標題改回原意
for b in secs["sens"]["blocks"]:
    if b["type"] == "tornado":
        b["title"] = "FY2028 EPS 相對基準的變動（情境 A，一次動一個輸入；" + ("未達兩位數百分比）" if b.get("collapse") else "變動達兩位數百分比的案例）")
# 後處理（補充規則 A）：橫幅中間格＝只看 Comparable（市場主流同業法、未經相對倍數調整）；附註列進加權的調整後值
_rc, _ru, _ra = row(VA, "Comparable 目標價"), row(VA, "上檔空間：Comparable"), row(VA, "Comparable 目標價（相對倍數調整後）")
assert VA.cell(_rc, 1).value == "Comparable 目標價" and str(VA.cell(_ra - 1, 1).value).startswith("相對倍數調整係數")
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{_rc}|n2", ["上檔空間 ", f"@Val!B{_ru}|sp1"]]
cfg["hero"]["note"] = list(cfg["hero"].get("note") or []) + ["。進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{_ra}|n2", "（調整係數 ", f"@Val!B{_ra - 1}|n2", "）"]
# 第三輪修正 17：橫幅金額標明幣別
_CUR3 = '人民幣'
def _lab3(x):
    t = str(wb[x[1:].split("!")[0]][x.split("!")[1]].value) if isinstance(x, str) and x.startswith("@") else str(x)
    return x if _CUR3 in t else ([x, "（" + _CUR3 + "）"] if isinstance(x, str) and x.startswith("@") else t + "（" + _CUR3 + "）")
for _t in cfg["hero"]["tiles"]:
    _t[0] = _lab3(_t[0])
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

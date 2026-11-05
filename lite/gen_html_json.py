# v0.4.1（2026-11-05）：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）——只改展示層。v0.5：DCF 改標準版（FCFF＋WACC），估值章 DCF 區塊改由本檔依 xlsx 列號產生；html_src.json 只重對 In 列號。
# 讀 lite/html_src.json（v0.4 的版面，內容不動）→ 少量前處理 → tools/company_html/relayout.py → 寫 lite/html.json。
# 用法：python3 lite/gen_html_json.py [成品 xlsx（預設 lite/dist/ 內唯一的 xlsx）]
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
DCF_BLOCKS = [
    {"type": "note", "parts": [f"@Val!A{_rk - 2}"]},
    {"type": "table", "caption": "DCF 折現率：股權成本、稅後債務成本、WACC、永續成長率、RONIC", "head": ["項目", "數值", "說明"],
     "rows": [[f"@Val!A{r}", f"@Val!B{r}|" + ("p1" if r == _rk + 2 else "p2"), f"@Val!E{r}"] for r in range(_rk, _rr + 1)]},
    {"type": "table", "caption": "DCF 彙總：明確期現值、終值現值、企業價值、淨現金與每股價值", "head": [f"@Val!{c}{_rk - 1}" for c in "ABCD"],
     "rows": [[f"@Val!A{r}"] + [f"@Val!{c}{r}|" + {_rp + 4: "n3", _rp + 7: "p1", _rp + 8: "x1"}.get(r, "n2") for c in "BCD"] for r in range(_rp, r_dtp)]
             + [[f"@Val!A{r_dtp}", f"@Val!B{r_dtp}|n2", "", ""]]}]
for s_ in "ABC":
    _a = row(DC, f"情境 {s_}（"); _b = next(r for r in range(_a, DC.max_row + 1) if str(DC.cell(r, 1).value or "").startswith("現值（FCFF"))
    _cols = ["A"] + [chr(ord("C") + i) for i in range(6)]
    _pct = {str(r): {c: "p1" for c in _cols[1:]} for r in range(_a + 1, _b + 1) if str(DC.cell(r, 1).value).startswith(("稅率", "折舊攤銷 ÷", "資本支出 ÷"))}
    _pct.update({str(r): {c: "n3" for c in _cols[1:]} for r in range(_a + 1, _b + 1) if str(DC.cell(r, 1).value).startswith("折現因子")})
    DCF_BLOCKS.append({"type": "table", "caption": f"@DCF!A{_a}", "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細",
                       "head": [f"@DCF!{c}{_dh}" for c in _cols],
                       "rowgen": {"sheet": "DCF", "rows": f"{_a + 1}-{_b}", "cols": _cols, "fmts": [""] + ["n2"] * 6, "row_fmt": _pct}})
# v0.5：Sens 目標價參數表改為標準版 DCF 的十個案例（tools/fcff_dcf.py write_sens；起始列比 v0.4.1 早一列）
SE = wb["Sens"]
_st = row(SE, "目標價參數敏感度")
_se = next(r for r in range(_st + 2, SE.max_row + 1) if str(SE.cell(r, 1).value or "").startswith("檢查：閉式公式"))
_n = 0
for b in secs["sens"]["blocks"]:
    if b["type"] == "table" and str(b.get("caption", "")).startswith("@Sens!A") and b.get("rowgen", {}).get("cols") == list("ABCDEF"):
        b["caption"] = f"@Sens!A{_st}"
        b["head"] = [f"@Sens!{c}{_st + 1}" for c in "ABCDEFG"]
        b["rowgen"] = {"sheet": "Sens", "rows": f"{_st + 2}-{_se - 1}", "cols": list("ABCDEFG"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2"]}
        _n += 1
assert _n == 1, _n

# 前處理（只動版面，不加數字）
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
    if b["type"] == "table" and b.get("caption") == f"@Val!A{r_dcf}":   # v0.5：舊 0.1 DCF（股東現金流）表 → 標準版 DCF 折現率、彙總與各情境 FCFF 明細
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
  ["AI 光學（雷射晶片、光模組、OCS、CPO）", [
    ["AI 光學市場與 TPU 顆數（讀網通市場模型；情境 A）", ["mpl_26", "mpl_27", "mpl_28", "mpl_29", "mcp_26", "mcp_27", "mcp_28", "mcp_29", "tpu_26", "tpu_27", "tpu_28", "tpu_29"]],
    ["網通 CPO 情境 B、C 的市場", ["mplB_26", "mplB_27", "mplB_28", "mplB_29", "mplC_26", "mplC_27", "mplC_28", "mplC_29",
                                  "mcpB_26", "mcpB_27", "mcpB_28", "mcpB_29", "mcpC_26", "mcpC_27", "mcpC_28", "mcpC_29"]],
    ["情境顆數係數", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["雷射晶片：模組單價、每模組雷射數、雷射單價與份額", ["asp_26", "asp_27", "asp_28", "asp_29", "lpm_26", "lpm_27", "lpm_28", "lpm_29",
                                                       "lasp_26", "lasp_27", "lasp_28", "lasp_29", "lrel_27", "lrel_28", "lrel_29"]],
    ["光模組、OCS、CPO", ["mrel_27", "mrel_28", "mrel_29", "ocsk_27", "ocsk_28", "ocsk_29", "cpo_27", "cpo_28", "cpo_29"]],
    ["錨點的 AI 產品拆分（雷射晶片＝殘差）", ["ocs_1h", "ocs_2h", "cpo_a26", "xcvr_a26"]],
    ["對帳參考（AI IT 資本支出）", ["itc_26", "itc_27", "itc_28", "itc_29"]],
    ["市場量敏感度係數（只進敏感度）", ["mf", "wta_27", "wta_28", "wta_29"]]]],
  ["電信與 DCI、工業雷射與三維感測", [
    ["錨點營收", ["tel_a26", "ind_a26"]],
    ["營收成長", ["g_tel_27", "g_tel_28", "g_tel_29", "g_ind"]]]],
  ["全公司營運：實績錨點、毛利率與費用", [
    ["季度實績與指引", ["q3_26", "q4_26", "q1g", "q_g", "om_g", "eps_g", "sh_q1"]],
    ["會計年度換算", ["h_1h"]],
    ["FY2025、FY2026 全年實績", ["a25_rev_in", "a25_gm_in", "a25_opx_in", "a25_ni_in", "a25_eps_in", "a26_rev_in", "a25_cmp", "a25_sys", "a26_cmp", "a26_sys",
                               "a26_gm_in", "a26_opx_in", "a26_ni_in", "a26_eps_in"]],
    ["毛利率路徑（FY2027 起）", ["gm_27", "gm_28", "gm_29"]],
    ["營業費用、營業外收入、稅與股數", ["phi", "opx_adj", "oi_oth", "tx", "sh_g"]]]],
  ["全公司：股價、折現率、目標價規則與收斂法", [
    ["股價與股數", ["px", "px_hi", "sh_basic"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_wA", "tp_wB", "tp_wC", "tp_w_dcf"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta_u", "beta_u_med", "beta_de", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "sh_v"]],
    ["自由現金流（股票薪酬、折舊、資本支出、營運資金）", ["fc_sbc", "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["收斂法（FY2030 起）", ["tp_g", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["共識目標價與前瞻 EPS（只作對照）", ["cons_tp", "cons_feps"]]]]],
 "blocks_to": {"數量、每模組雷射數與份額": "AI 光學（雷射晶片、光模組、OCS、CPO）", "@Units!A15": "AI 光學（雷射晶片、光模組、OCS、CPO）"},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "Comparable 計算": ["同業前瞻 P/E 中位數", "12 個月後的前瞻 EPS", "Comparable 目標價"],
          "sheet:Recon": ["FY2027 EPS vs 共識", "FY2028 EPS vs 共識", "CY2026 雷射有效份額 vs", "CY2027 AI 相關營收 ÷", "FY2026 EPS vs 公司揭露"]},
}
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
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

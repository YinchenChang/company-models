# v0.4.1（2026-11-05）：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）——只改展示層。
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
    if b["type"] == "table" and b.get("caption") == f"@Val!A{r_dcf}":
        out.append({"type": "table", "caption": "DCF 折現率與永續假設（股東權益現金流折現）", "head": ["@In!B3", "@In!E3", "@In!C3", "@In!G3"],
                    "rows": [[f"@In!B{ink(k)}", f"@In!E{ink(k)}|p1", f"@In!C{ink(k)}", f"@In!G{ink(k)}|tag"] for k in ("tp_r", "tp_g", "tp_roe")]})
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
    ["折現率與永續假設", ["tp_r", "tp_g", "tp_roe"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_wA", "tp_wB", "tp_wC", "tp_w_dcf"]],
    ["收斂法（FY2030 起）", ["fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["共識目標價與前瞻 EPS（只作對照）", ["cons_tp", "cons_feps"]]]]],
 "blocks_to": {"數量、每模組雷射數與份額": "AI 光學（雷射晶片、光模組、OCS、CPO）", "@Units!A15": "AI 光學（雷射晶片、光模組、OCS、CPO）"},
 "main": {"0.1 DCF": ["DCF 情境每股價值", "  DCF 情境價值推到 12 個月後", "DCF 目標價", "  FY2028", "  FY2032"],
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

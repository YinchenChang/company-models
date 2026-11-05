# v0.4.1（2026-11-05）：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）——只改展示層。
# 讀 arista/html_src.json（v0.4 的版面，內容不動）→ 少量前處理 → tools/company_html/relayout.py → 寫 arista/html.json。
# 用法：python3 arista/gen_html_json.py [成品 xlsx（預設 arista/dist/ 內唯一的 xlsx）]
import copy, glob, json, os, re, sys
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
# 2. 估值：重點觀察已移到第一章；相對倍數調整由 kv 改為表格（轉換器會把 kv 當成折現率表）；
#    DCF 折現率直接顯示、緊接在 DCF 彙總之前；無標題的 DCF、Comparable 計算表補標題以便指定主要列
va = secs["valuation"]["blocks"]
out, skip = [], False
for b in va:
    if b["type"] == "h3" and b.get("text") == "重點觀察":
        skip = True
        continue
    if skip and b["type"] == "list":
        skip = False
        continue
    if b["type"] == "kv":
        b = {"type": "table", "caption": "相對倍數調整", "head": ["項目", "數值"], "rows": [list(i) for i in b["items"]]}
    if b["type"] == "table" and not b.get("caption") and b.get("rows") and b["rows"][0][0] == f"@Val!A{row(VA, '股東現金流＝')}":
        out.append({"type": "table", "caption": "DCF 折現率與永續假設（股東權益現金流折現）", "head": ["@In!B3", "@In!E3", "@In!C3", "@In!G3"],
                    "rows": [[f"@In!B{ink(k)}", f"@In!E{ink(k)}|p1", f"@In!C{ink(k)}", f"@In!G{ink(k)}|tag"] for k in ("tp_r", "tp_g", "tp_roe")]})
        b = dict(b, caption="DCF 彙總（股東權益現金流折現）")
    if b["type"] == "table" and not b.get("caption") and b.get("rows") and b["rows"][0][0] == f"@Val!A{row(VA, '同業前瞻 P/E 中位數')}":
        b = dict(b, caption="Comparable 計算")
    out.append(b)
secs["valuation"]["blocks"] = out
# 3. 對帳：單一大表補標題以便指定主要列
for b in secs["recon"]["blocks"]:
    if b["type"] == "table" and not b.get("caption"):
        b["caption"] = "對帳要點"
tor_title = {b["title"]: b for b in secs["sens"]["blocks"] if b["type"] == "tornado"}

scen_tok = re.split(r"[.\s]", IN.cell(row(IN, "情境驅動對照"), 1).value)[0]
SPEC = {
 "in_skip": [scen_tok],
 "in_groups": [
  ["AI 後端乙太網路（AI 營收）", [
    ["AI 後端 scale-out 交換器市場（讀網通市場模型）", ["mkt_26", "mkt_27", "mkt_28", "mkt_29"]],
    ["AI IT 資本支出（對帳參考）", ["itc_26", "itc_27", "itc_28", "itc_29"]],
    ["情境顆數係數", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["乙太網路占比", ["eth_26", "eth_27", "eth_28", "eth_29"]],
    ["Arista 份額相對值", ["rel_27", "rel_28", "rel_29"]],
    ["AI 營收校準", ["ai25", "ai26"]],
    ["AI 毛利率", ["d_ai"]],
    ["市場量敏感度係數（只進敏感度）", ["mf", "wta_27", "wta_28", "wta_29"]]]],
  ["非 AI 產品與服務", [
    ["非 AI 產品營收成長", ["g_nai27", "g_nai28", "g_nai29"]],
    ["服務營收", ["svc25", "g_svc26", "g_svc27", "g_svc28", "g_svc29"]]]],
  ["全公司營運：FY2026 錨點與營業費用", [
    ["FY2026 錨點（第三季實績與第四季指引）", ["q3_rev", "q3_gp", "q3_oi", "q4_rev", "q4_gm", "q4_oim", "q_oth", "q_tx", "q_sh"]],
    ["營業費用、其他收益、稅與股數（FY2027 起）", ["phi", "opx_adj", "oth_f", "tx_f", "sh_g"]]]],
  ["全公司：股價、折現率、目標價規則與現金流", [
    ["股價與市場倍數", ["px", "pe_ttm_gaap", "pe_fwd_sa", "g2n"]],
    ["共識與先驗（只作對帳）", ["cons_rev26", "cons_eps26", "erste_eps27", "dell_top2", "idc_anet", "pe_fair_andy", "cons_tp"]],
    ["折現率與永續假設", ["tp_r", "tp_g", "tp_gd", "tp_roe", "nm10_k"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_wA", "tp_wB", "tp_wC", "tp_w_dcf"]],
    ["自由現金流比率", ["fcf_da", "fcf_capex"]]]]],
 "blocks_to": {"Arista 份額": "AI 後端乙太網路（AI 營收）", "@Units!A13": "AI 後端乙太網路（AI 營收）"},
 "main": {"0.1 DCF": ["FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"],
          "DCF 彙總": ["股東現金流＝", "永續價值", "DCF 情境每股價值", "其中：永續價值現值占比", "DCF 目標價"],
          "對帳要點": ["FY2026 營收 vs 共識", "FY2026 EPS vs 共識", "FY2027 EPS vs 前瞻", "FY2026 Arista 有效份額 vs IDC", "3Q26 EPS"]},
}
cfg = relayout(cfg, X, SPEC)
# 後處理：轉換器把估值章所有長條圖移到最前面；舊法（規則值）長條圖放回舊法段落，避免與目標價並列
_vb = cfg["sections"][[s["id"] for s in cfg["sections"]].index("valuation")]["blocks"]
_old = next(b for b in _vb if b["type"] == "bars" and b.get("title") == "規則值隱含股價與現價")
_vb.remove(_old)
_h = next(i for i, b in enumerate(_vb) if b["type"] == "h3" and str(b.get("text", "")).startswith("規則值隱含股價"))
_vb.insert(_h + 2, _old)
# 後處理：單位為 pt（百分點）的參數，轉換器預設一位小數；改用百分比格式，避免 0.06 顯示成 0.1
_pt = {r for r in range(4, IN.max_row + 1) if IN.cell(r, 3).value == "pt"}
def _fix_pt(x):
    m = re.match(r"^@In!([DEF])(\d+)\|n1$", x) if isinstance(x, str) else None
    return f"@In!{m.group(1)}{m.group(2)}|p1" if m and int(m.group(2)) in _pt else x
for b in cfg["sections"][[s["id"] for s in cfg["sections"]].index("drivers")]["blocks"]:
    if b["type"] == "table" and isinstance(b.get("rows"), list):
        b["rows"] = [[_fix_pt(c) for c in r] if isinstance(r, list) else r for r in b["rows"]]
# 後處理：本模型的龍捲風圖是 FY2028 EPS（不是加權目標價），轉換器的通用標題改回原意
for b in cfg["sections"][[s["id"] for s in cfg["sections"]].index("sens")]["blocks"]:
    if b["type"] == "tornado":
        b["title"] = "FY2028 EPS 相對基準的變動（情境 A，一次動一個輸入；" + ("未達兩位數百分比）" if b.get("collapse") else "變動達兩位數百分比的案例）")
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

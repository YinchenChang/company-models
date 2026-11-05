# v0.2.4：套用個股 HTML 版面規格 v1.0；v0.3：DCF 改標準版（FCFF＋WACC），html_src.json 的 DCF 區塊與 In 列號已隨 v0.3 xlsx 更新（docs/plan/個股HTML版面規格.md）
# 讀 nvidia/html_src.json（v0.2.3 的 html.json，原樣保留）→ 本檔前處理 → tools/company_html/relayout.py → 本檔後處理 → 寫 nvidia/html.json。
# 用法：python3 nvidia/gen_html_json.py（讀 nvidia/dist 內唯一的 xlsx）；之後 python3 tools/company_html/build_company_html.py nvidia
import copy, glob, json, os, re, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
from relayout import relayout  # noqa: E402

xs = glob.glob(os.path.join(HERE, "dist", "*.xlsx"))
assert len(xs) == 1, xs
X = xs[0]
cfg = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)


def find(sh, prefix, col="A"):
    ws = wb[sh]
    for r in range(1, ws.max_row + 1):
        v = ws[f"{col}{r}"].value
        if isinstance(v, str) and v.strip().startswith(prefix):
            return r
    raise KeyError((sh, prefix))


IN = wb["In"]
KEY = {str(IN[f"A{r}"].value): r for r in range(4, IN.max_row + 1) if IN[f"A{r}"].value}
SEC = {s["id"]: s for s in cfg["sections"]}

# ---------- 前處理（只改展示設定，不動 html_src.json） ----------
# 驅動因子：輸入 → 推導量的對照表移到附錄（轉換器依 caption「驅動 → 推導」辨識）
dr = SEC["drivers"]
drivers_intro = dr.pop("intro", None)
for b in dr["blocks"]:
    if b.get("caption", "").startswith("單一參數"):
        b["caption"] = "驅動 → 推導對照（In 頁基準值與推導量）"

# 估值：相對倍數 kv 改成有標題的表；移除舊「重點觀察」（已移到第一章）；長條圖加入相對倍數調整後的 Comparable；
# v0.3：DCF 折現率表（WACC 等）與 DCF 彙總、各情境 FCFF 明細已寫在 html_src.json（折現率緊接 DCF 彙總之前）；上一版估值內容摺疊
va = SEC["valuation"]
nb, legacy = [], False
r_adj = find("Val", "Comparable 目標價（相對倍數調整後）")
LEG_SUM = iter(["展開：舊法規則值情境加權（參考，非目標價）", "展開：現價與市場隱含倍數", "展開：要達到現價所需的倍數",
                "展開：循環位置判定", "展開：以營收路徑判斷循環位置（對照）", "展開：各分位倍數"])
skip_next_list = False
for b in va["blocks"]:
    t = b["type"]
    if t == "h3" and b.get("text") == "重點觀察":
        skip_next_list = True
        continue
    if t == "list" and skip_next_list:
        skip_next_list = False
        continue
    if t == "h3" and "相對倍數調整" in b.get("text", ""):
        continue
    if t == "kv" and not legacy:
        nb.append({"type": "table", "caption": "相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司相對同業的長期倍數比）",
                   "head": ["項目", "數值"], "rows": [[k, v] for k, v in b["items"]]})
        continue
    if t == "bars":
        b = copy.deepcopy(b)
        b["items"].insert(3, {"label": f"@Val!A{r_adj}", "value": f"@Val!B{r_adj}", "series": 2})
    if t == "h3" and b.get("text", "").startswith("以下為上一版"):
        legacy = True
    if t == "table" and b.get("caption") == "@Val!A25":
        b = dict(b, collapse=True, summary="展開：同業前瞻本益比")
    if t == "table" and "caption" not in b and b.get("rows") and b["rows"][0][0] == f"@Val!A{find('Val', '同業前瞻 P/E 中位數')}":
        b = dict(b, caption="Comparable 彙總：同業倍數 × 前瞻 EPS")
    if legacy and t == "table" and not b.get("collapse"):
        b = dict(b, collapse=True, summary=next(LEG_SUM))
    nb.append(b)
va["blocks"] = nb

# 對帳：單一大表加標題，指定主要對帳要點
rc = SEC["recon"]
for b in rc["blocks"]:
    if b["type"] == "table" and "caption" not in b:
        b["caption"] = "對帳要點：本模型與公司指引、共識、命題"

skip = [re.split(r"[.\s]", str(IN[f"A{r}"].value))[0] for r in range(4, IN.max_row + 1)
        if str(IN[f"A{r}"].value or "").startswith("情境驅動對照")] + ["I"]

SPEC = {
 "in_skip": skip,
 "in_groups": [
  ["資料中心加速器（GPU 與整櫃系統）", [
    ["NVIDIA 各世代顆數", ["u_bw_26", "u_bw_27", "u_bw_28", "g29_bw", "u_rb_26", "u_rb_27", "u_rb_28", "g29_rb", "u_ru_26", "u_ru_27", "u_ru_28", "g29_ru", "u_hp_26", "u_hp_27", "u_hp_28", "g29_hp"]],
    ["情境顆數係數與顆數敏感度", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C", "uk"]],
    ["每顆售價（GPU 模組＋CPU＋LPDDR＋匣級加價）", ["pm_bw", "pm_rb", "pm_ru", "pm_hp", "pcpu_bw", "plp_bw", "pres_bw", "pcpu_rb", "plp_rb", "pres_rb", "pcpu_ru", "plp_ru", "pres_ru", "pcpu_hp", "plp_hp", "pres_hp", "k_tray"]],
    ["HBM 容量、單價與吸收比例", ["gb_bw_26", "gb_bw_27", "gb_bw_28", "gb_rb_26", "gb_rb_27", "gb_rb_28", "gb_ru_26", "gb_ru_27", "gb_ru_28", "gb_hp_26", "gb_hp_27", "gb_hp_28", "hpx_bw", "hpx_rb", "hpx_ru", "hpx_hp", "abs_mem"]],
    ["記憶體價格路徑（讀 SK hynix 模型）", ["hbm_A_26", "hbm_A_27", "hbm_A_28", "hbm_A_29", "hbm_B_26", "hbm_B_27", "hbm_B_28", "hbm_B_29", "hbm_C_26", "hbm_C_27", "hbm_C_28", "hbm_C_29",
                                     "cdx_A_26", "cdx_A_27", "cdx_A_28", "cdx_A_29", "cdx_B_26", "cdx_B_27", "cdx_B_28", "cdx_B_29", "cdx_C_26", "cdx_C_27", "cdx_C_28", "cdx_C_29"]],
    ["競爭對手顆數與每顆金額（份額計算）", ["u_amd_26", "u_amd_27", "u_amd_28", "g29_amd", "u_tpu_26", "u_tpu_27", "u_tpu_28", "g29_tpu", "u_trn_26", "u_trn_27", "u_trn_28", "g29_trn",
                                     "u_oth_26", "u_oth_27", "u_oth_28", "g29_oth", "pa_amd", "pa_tpu", "pa_trn", "pa_oth", "dsh"]],
    ["中國資料中心運算（只進敏感度）", ["cn_amt"]],
    ["Wattonomics 錨定係數（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  ["網通（InfiniBand、Spectrum-X、NVLink）", [
    ["網通營收與成本", ["net_26", "net_27", "net_28", "net_29", "rep_nv26", "net_q1", "c_net"]]]],
  ["Edge Computing（遊戲、PC、工作站、車用與機器人）", [
    ["營收與毛利率", ["gm_edge", "eg", "edge_q"]]]],
  ["全公司：錨點、費用、稅與股數、股價與目標價規則", [
    ["FY2027 錨點（第三季指引＋第四季假設）", ["q3_rev", "q3_gm", "q3_opx", "q4_g", "q4_gm", "q4_opx_g", "q_oth", "tx_27", "q_sh"]],
    ["營業費用、稅與股數", ["phi", "g_opx_min", "opx_adj", "tx_f", "oth_f", "sh_g"]],
    ["股價與循環位置", ["px", "r_disc", "pk_A", "pk_B", "pk_C"]],
    ["目標價規則（情境權重、DCF 權重、延伸期收斂與共識）", ["tp_wA", "tp_wB", "tp_wC", "tp_g", "tp_gd", "tp_w_dcf", "nm10", "cons_tp"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "blume_a", "beta_1", "beta_2", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["da_rev", "cx_rev", "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]]]]],
 "blocks_to": {},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "Comparable 彙總": ["同業前瞻 P/E 中位數", "各情境 12 個月後的前瞻 EPS", "12 個月後的前瞻 EPS（情境加權）", "Comparable 目標價"],
          "對帳要點": ["FY2028 營收 vs 指引推算", "FY2028 毛利率 vs 指引", "FY2028 EPS vs 共識", "2027 年 NVIDIA 金額份額（加速器晶片", "FY2028 毛利率 − FY2027 毛利率", "現價 ÷ FY2028 EPS vs 命題"]},
}
old_thesis_kv = [b for b in SEC["thesis"]["blocks"] if b["type"] == "kv"]
cfg = relayout(cfg, X, SPEC)
SEC = {s["id"]: s for s in cfg["sections"]}

# ---------- 後處理 ----------
# 橫幅中間格（工作單補充規則 A）：只看 Comparable＝市場主流同業法、未經相對倍數調整的 Val!B36（上檔空間 Val!B44＝B41/px−1，B41＝B36）；
# 進加權的是相對倍數調整後的值，寫在附註
r_c, r_cu = find("Val", "Comparable 目標價"), find("Val", "上檔空間：Comparable")
assert wb["Val"][f"B{find('Val', 'Comparable 目標價（市場主流同業法）')}"].value == wb["Val"][f"B{r_c}"].value
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{r_c}|n2", ["上檔空間 ", f"@Val!B{r_cu}|sp1"]]
cfg["hero"]["note"] = list(cfg["hero"].get("note") or []) + [
    "；進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{r_adj}|n2", "（市場主流同業法 × 相對倍數調整係數 ", f"@Val!B{r_adj - 1}|n2", "）"]
# 第一章：保留情境定義與會計年度（原第一章內容）
if old_thesis_kv:
    SEC["thesis"]["blocks"] += [{"type": "h3", "text": "情境定義與會計年度"}] + old_thesis_kv
# 敏感度：本模型龍捲風圖的值是 FY2028 EPS（情境 A）相對基準的變動，不是加權目標價
for b in SEC["sens"]["blocks"]:
    if b["type"] == "tornado":
        b["title"] = ("FY2028 EPS 相對基準的變動（情境 A，一次動一個輸入；未達兩位數百分比）" if b.get("collapse")
                      else "FY2028 EPS 相對基準的變動（情境 A，一次動一個輸入；變動達兩位數百分比的案例）")
        if b.get("collapse"):
            b["summary"] = "展開：FY2028 EPS 變動未達兩位數百分比的案例"
# 附錄：驅動 → 推導對照表的說明
st = SEC["status"]["blocks"]
for i, b in enumerate(st):
    if b["type"] == "h3" and b.get("text", "").startswith("附錄：驅動 → 推導") and drivers_intro:
        st.insert(i + 1, {"type": "note", "parts": [drivers_intro]})
        break
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

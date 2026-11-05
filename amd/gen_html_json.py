# v0.3.1：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）
# 讀 amd/html_src.json（v0.3 的 html.json，原樣保留）→ 本檔前處理 → tools/company_html/relayout.py → 本檔後處理 → 寫 amd/html.json。
# 用法：python3 amd/gen_html_json.py（讀 amd/dist 內唯一的 xlsx）；之後 python3 tools/company_html/build_company_html.py nvidia
import copy, glob, itertools, json, os, re, sys
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
# DCF 折現率表放在 DCF 彙總之前；上一版估值內容摺疊
va = SEC["valuation"]
nb, legacy = [], False
r_adj = find("Val", "Comparable 目標價（相對倍數調整後）")
LEG_SUM = itertools.repeat("展開：上一版估值內容（保留供對照）")
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
    if t == "table" and "caption" not in b and b.get("rows") and b["rows"][0][0] == f"@Val!A{find('Val', '股東現金流＝EPS')}":
        rows = [[f"@In!B{KEY[k]}", f"@In!E{KEY[k]}|p1"] for k in ("tp_r", "tp_g", "tp_roe")]
        rows.append([f"@Val!A{find('Val', '股東現金流＝EPS')}", f"@Val!B{find('Val', '股東現金流＝EPS')}|p1"])
        nb.append({"type": "table", "caption": "DCF 折現率與永續假設（股東權益成本、永續成長率、再投資報酬率與可發放比例）", "head": ["項目", "數值"], "rows": rows})
    if t == "table" and "caption" not in b and b.get("rows") and b["rows"][0][0] == f"@Val!A{find('Val', '股東現金流＝EPS')}":
        b = dict(b, caption="DCF 彙總：股東現金流、永續價值與每股價值")
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
        b["caption"] = "對帳要點：本模型與公司說法、大單、共識"

skip = [re.split(r"[.\s]", str(IN[f"A{r}"].value))[0] for r in range(4, IN.max_row + 1)
        if str(IN[f"A{r}"].value or "").startswith("情境驅動對照")] + ["H"]

SPEC = {
 "in_skip": skip,
 "in_groups": [
  ["資料中心 GPU（Instinct）", [
    ["AMD 顆數與每顆 HBM 容量", ["u_amd_26", "u_amd_27", "u_amd_28", "g29_amd", "gb_amd_26", "gb_amd_27", "gb_amd_28"]],
    ["其他平台顆數（市場金額與份額用）", ["u_bw_26", "u_bw_27", "u_bw_28", "g29_bw", "u_rb_26", "u_rb_27", "u_rb_28", "g29_rb", "u_ru_26", "u_ru_27", "u_ru_28", "g29_ru", "u_hp_26", "u_hp_27", "u_hp_28", "g29_hp",
                               "u_tpu_26", "u_tpu_27", "u_tpu_28", "g29_tpu", "u_trn_26", "u_trn_27", "u_trn_28", "g29_trn", "u_oth_26", "u_oth_27", "u_oth_28", "g29_oth"]],
    ["情境顆數係數與顆數敏感度", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C", "uf"]],
    ["各平台每顆售價（讀 NVIDIA 模型）", ["pm_bw", "pm_rb", "pm_ru", "pm_hp", "pa_amd", "pa_tpu", "pa_trn", "pa_oth"]],
    ["GPU 錨點、毛利率與世代售價係數", ["gpu25", "gs_q1", "gs_q2", "ng_q3", "ng_q4", "gm_gpu", "pg_27", "pg_28", "pg_29"]],
    ["HBM 價格路徑與吸收比例", ["hbm_A_26", "hbm_A_27", "hbm_A_28", "hbm_A_29", "hbm_B_26", "hbm_B_27", "hbm_B_28", "hbm_B_29", "hbm_C_26", "hbm_C_27", "hbm_C_28", "hbm_C_29", "abs_mem", "abs_mem_A", "abs_mem_B", "abs_mem_C"]],
    ["客戶大單認股權證", ["w_tot", "w_gw", "gw_26", "gw_27", "gw_28", "gw_29", "wf"]],
    ["Wattonomics 錨定係數（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  ["資料中心 CPU 與其他業務（Client、Gaming、Embedded）", [
    ["成長率（FY2027 起）", ["g_cpu27", "g_cpu28", "g_cpu29", "g_cli", "g_gam", "g_emb"]]]],
  ["全公司：錨點、費用、稅與股數、股價與目標價規則", [
    ["FY2026 錨點（第三季實績＋第四季指引）", ["q3_rev", "q3_gm", "q3_opx", "q4_oth", "tx_26", "q3_sh", "q3_cli", "q3_gam", "q3_emb", "q4_rev", "q4_cli", "q4_gam", "q4_emb", "q4_gm", "q4_opx", "q4_sh"]],
    ["營業費用、稅與股數", ["phi", "g_opx_min", "opx_adj", "tx_f", "oth_f", "sh_g"]],
    ["股價與循環位置", ["px", "r_disc", "pk_A", "pk_B", "pk_C"]],
    ["目標價規則（情境權重、DCF 與 Comparable）", ["tp_wA", "tp_wB", "tp_wC", "tp_r", "tp_g", "tp_gd", "tp_roe", "tp_w_dcf", "nm10_k", "cons_tp"]]]]],
 "blocks_to": {},
 "main": {"0.1 DCF": ["FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"],
          "DCF 彙總": ["股東現金流＝EPS", "永續價值", "DCF 情境每股價值", "其中：永續價值現值占比", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["DCF 折現率", "長期成長率", "ROE", "股東現金流＝EPS"],
          "Comparable 彙總": ["同業前瞻 P/E 中位數", "各情境 12 個月後的前瞻 EPS", "12 個月後的前瞻 EPS（情境加權）", "Comparable 目標價"],
          "對帳要點": ["FY2027 資料中心營收 vs 公司", "FY2027 毛利率 vs FY2026 錨點", "FY2027 EPS vs 共識（stockanalysis", "FY2028 EPS vs 共識約", "FY2029 EPS vs 公司策略期間", "認股權證稀釋占 FY2029 股數"]},
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
    "；進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{r_adj}|n2", "（市場主流同業法 × 相對倍數調整係數 ", f"@Val!B{r_adj - 1}|n2",
    "，已達夾值上限）" if wb["Val"][f"B{r_adj - 1}"].value >= wb["Val"][f"B{r_adj - 2}"].value else "）"]
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

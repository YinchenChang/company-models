# v0.4：DCF 改為標準版（FCFF＋WACC；html_src.json 的 DCF 區塊同 NVIDIA v0.3）。v0.3.1：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）
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

# 情境驅動表：依 In 頁「情境驅動對照」節的實際位置重建（html_src.json 寫死的 In 列號在 M2 節插入後錯位；不寫死列號）
# 第三輪修正：情境驅動表數值格式（年度不加千分位、百分比／pt 以 % 顯示、其他依大小取小數位）
def _scen_fmt(ws, r):
    u = str(ws.cell(r, 3).value or "").strip()
    vs = [ws.cell(r, c).value for c in (4, 5, 6)]
    vs = [v for v in vs if isinstance(v, (int, float)) and not isinstance(v, bool)]
    if not vs:
        return ""
    if all(float(v).is_integer() and 1990 <= v <= 2100 for v in vs) and ("年" in u or u in ("", "—")):
        return "|yr"
    if u in ("%", "pt", "百分點") or u.startswith("%"):
        return "|p2" if any(0 < abs(v) < 0.001 for v in vs) or any(abs(round(v * 1000) - v * 1000) > 1e-6 for v in vs) else "|p1"
    if all(float(v).is_integer() for v in vs):
        return "|n0"
    return "|n1" if max(abs(v) for v in vs) >= 100 else "|n2"
def _scen_fix():
    ws = wb["In"]
    s0 = next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value or "").startswith("情境驅動對照"))
    h0 = s0 + 1
    assert ws.cell(h0, 1).value == "鍵", ws.cell(h0, 1).value
    r1 = h0 + 1
    while ws.cell(r1, 1).value: r1 += 1
    n = 0
    for b in SEC["drivers"]["blocks"]:
        if b["type"] == "table" and str(b.get("caption", "")).startswith("情境驅動"):
            b["head"] = [f"@In!{c}{h0}" for c in "BDEFC"]
            b["rows"] = [[f"@In!{c}{r}" + (_scen_fmt(ws, r) if c in "DEF" else "") for c in "BDEFC"] for r in range(h0 + 1, r1)]
            n += 1
    assert n == 1, n
_scen_fix()

# 第二輪修正：DCF 彙總表與「目標價與現價」長條圖的 Val 列號依標籤重建（模組 v2 stub 期在「＋淨現金與投資」後多兩列，不寫死列號；同 NVIDIA v0.3.2）
def _dcf_sum_fix():
    r0, r1 = find("Val", "明確期 FCFF 現值"), find("Val", "DCF 目標價（情境加權）")
    fm = lambda lab: ("n3" if lab.startswith("評價用稀釋股數") else "p1" if "終值現值占 EV" in lab else "x1" if "隱含終值" in lab else
                      "n2" if lab.startswith(("DCF 情境", "DCF 目標價")) else "n1")
    n = 0
    for b in SEC["valuation"]["blocks"]:
        if b["type"] == "table" and str(b.get("caption", "")).startswith("DCF 彙總"):
            rows = []
            for r in range(r0, r1 + 1):
                lab = str(wb["Val"][f"A{r}"].value).strip(); f = fm(lab)
                rows.append([f"@Val!A{r}", f"@Val!B{r}|{f}", "", ""] if r == r1 else [f"@Val!A{r}"] + [f"@Val!{c}{r}|{f}" for c in "BCD"])
            b["rows"] = rows; n += 1
        if b["type"] == "bars":
            for it in b["items"]:
                if it["value"] == "@Val!B21":   # html_src.json 的 DCF 目標價列
                    it["label"], it["value"] = f"@Val!A{r1}", f"@Val!B{r1}"; n += 1
    assert n == 2, n
_dcf_sum_fix()

# ---------- 前處理（只改展示設定，不動 html_src.json） ----------
# 驅動因子：輸入 → 推導量的對照表移到附錄（轉換器依 caption「驅動 → 推導」辨識）
dr = SEC["drivers"]
drivers_intro = dr.pop("intro", None)
for b in dr["blocks"]:
    if b.get("caption", "").startswith("單一參數"):
        b["caption"] = "驅動 → 推導對照（In 頁基準值與推導量）"

# 估值：相對倍數 kv 改成有標題的表；移除舊「重點觀察」（已移到第一章）；長條圖加入相對倍數調整後的 Comparable；
# v0.4：DCF 折現率表（WACC 等）與 DCF 彙總、各情境 FCFF 明細已寫在 html_src.json（折現率緊接 DCF 彙總之前）；上一版估值內容摺疊
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
    ["目標價規則（情境權重、DCF 權重、延伸期收斂與共識）", ["tp_wA", "tp_wB", "tp_wC", "tp_g", "tp_gd", "tp_w_dcf", "nm10_k", "cons_tp"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta_u", "beta_de", "beta_du", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（股票薪酬、折舊、資本支出、營運資金）", ["sbc_rev", "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]]]]],
 "blocks_to": {},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
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
# 第三輪修正：橫幅金額單位（美元）
for _t in cfg["hero"]["tiles"]:
    if isinstance(_t[0], str):
        _t[0] = [_t[0], "（美元）"]
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

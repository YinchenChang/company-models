# v0.2.4：套用個股 HTML 版面規格 v1.0；v0.3：DCF 改標準版（FCFF＋WACC），html_src.json 的 DCF 區塊與 In 列號已隨 v0.3 xlsx 更新（docs/plan/個股HTML版面規格.md）；v0.3.2：DCF 彙總表列號依標籤重建（模組 v2 stub 期）
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
# 情境驅動表：依 In 頁「情境驅動對照」節的實際位置重建（html_src.json 寫死的 In 列號在 M2 節〔標準版 DCF〕插入後錯位；不寫死列號）
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

# v0.3.2：DCF 彙總表與「目標價與現價」長條圖的 Val 列號依標籤重建（模組 v2 stub 期在「＋淨現金與投資」後多兩列，不寫死列號）
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
                if it["value"] == "@Val!B21":   # html_src.json（v0.3）的 DCF 目標價列
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
    ["三層協調：產業部署 GW 兩讀數與對帳中旗標（只進兩讀數目標價）", ["co_gwb_26", "co_gwb_27", "co_gwb_28", "co_gwb_29", "co_gwa_26", "co_gwa_27", "co_gwa_28", "co_gwa_29",
                                     "co_gws_26", "co_gws_27", "co_gws_28", "co_gws_29", "co_flag_2026", "co_flag_2027", "co_flag_2028", "co_flag_2029", "co_ns_27", "co_ns_28", "co_ns_29", "co_r_26", "co_r_27", "co_r_28", "co_r_29", "co_ucoef_28", "co_ucoef_29", "co_chk_r"]]]],
  ["網通（InfiniBand、Spectrum-X、NVLink）", [
    ["網通營收與成本", ["net_26", "net_27", "net_28", "net_29", "rep_nv26", "net_q1", "c_net"]]]],
  ["Edge Computing（遊戲、PC、工作站、車用與機器人）", [
    ["營收與毛利率", ["gm_edge", "eg", "edge_q"]]]],
  ["全公司：錨點、費用、稅與股數、股價與目標價規則", [
    ["FY2027 錨點（第三季指引＋第四季假設）", ["q3_rev", "q3_gm", "q3_opx", "q4_g", "q4_gm", "q4_opx_g", "q_oth", "tx_27", "q_sh"]],
    ["營業費用、稅與股數", ["phi", "g_opx_min", "opx_adj", "tx_f", "oth_f", "sh_g"]],
    ["股價與循環位置", ["px", "r_disc", "pk_A", "pk_B", "pk_C"]],
    ["目標價規則（情境權重、DCF 權重、延伸期收斂與共識）", ["tp_wA", "tp_wB", "tp_wC", "tp_g", "tp_gd", "tp_w_dcf", "nm10", "cons_tp"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta_u", "beta_de", "blume_a", "beta_du", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["da_rev", "cx_rev", "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]]]]],
 "blocks_to": {},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
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
# 第三輪修正（v0.3.3）：橫幅金額單位（美元）
for _t in cfg["hero"]["tiles"]:
    if isinstance(_t[0], str) and not _t[0].endswith("（美元）"):
        _t[0] = [_t[0], "（美元）"]
# ================= v0.3.5／v0.3.6（2026-11-11）：NVIDIA v0.3.4 驗收採用規劃書（nvidia/v035_post.py；做法同 amd/gen_html_json.py v0.4.5） =================
SEC = {s["id"]: s for s in cfg["sections"]}
RDW, SEW, RCW = wb["README"], wb["Sens"], wb["Recon"]
def _rows_after(head, stop_prefixes):
    r = find("README", head) + 1; out = []
    while RDW.cell(r, 1).value and not str(RDW.cell(r, 1).value).startswith(stop_prefixes):
        out.append(r); r += 1
    return out
r_state, r_edge = find("README", "build_id "), find("README", "評等")
r_a47, r_r47, r_g50 = find("README", "上檔空間三段歸因"), find("README", "現行相對倍數與 10 年比值"), find("README", "基準與公司指引")
r_asof_h = find("README", "資料基準日（as_of）"); r_asof_t = r_asof_h + 1; asof_rows = _rows_after("項目", ("說明：",))
r_asof_n = find("README", "說明：rf、同業")
fail_rows = _rows_after("主要失效條件", ("估值框架並列",)); r_frame = find("README", "估值框架並列")
risk_rows = _rows_after("風險登錄", ("主要假設驅動",)); ind_rows = _rows_after("主要假設驅動與可證偽", ("\u0000",))
st_add = [{"type": "note", "parts": [f"@README!A{r_state}"]}, {"type": "note", "parts": [f"@README!A{r_edge}"]},
          {"type": "h3", "text": "主要失效條件與估值框架並列"}, {"type": "list", "items": [f"@README!A{r}" for r in fail_rows] + [f"@README!A{r_frame}"]},
          {"type": "h3", "text": "資料基準日與逾期標示"}, {"type": "note", "parts": [f"@README!A{r_asof_h}"]},
          {"type": "table", "head": [f"@README!{c}{r_asof_t}" for c in "ABC"], "rows": [[f"@README!A{r}", f"@README!B{r}", f"@README!C{r}"] for r in asof_rows]},
          {"type": "note", "parts": [f"@README!A{r_asof_n}"]},
          {"type": "h3", "text": "風險登錄（是否建模、理由、觀察指標）"}, {"type": "list", "items": [f"@README!A{r}" for r in risk_rows]},
          {"type": "h3", "text": "主要假設驅動與可證偽的公開指標"}, {"type": "list", "items": [f"@README!A{r}" for r in ind_rows]}]
SEC["status"]["blocks"] = st_add + SEC["status"]["blocks"]
# 橫幅：原則 36 評等邊界、原則 47 三段歸因與現行相對倍數、原則 50 指引差距
cfg["hero"]["note"] = list(cfg["hero"]["note"]) + ["", f"@README!A{r_edge}", "", f"@README!A{r_a47}", "", f"@README!A{r_r47}", "", f"@README!A{r_g50}"]
# 第一章：投資建議條件式（README「投資建議：」列）
r_rec = find("README", "投資建議：")
SEC["thesis"]["blocks"].insert(0, {"type": "note", "parts": [f"@README!A{r_rec}"]})
# 估值章：估值框架並列與期間並列、方法占比、子群中位數
va = SEC["valuation"]
r_p1, r_p2, r_sh, r_s1, r_s2 = find("Val", "折現期間："), find("Val", "持有期間："), find("Val", "DCF 在加權目標價中的占比"), find("Val", "子群中位數（純展示）：GPU"), find("Val", "子群中位數（純展示）：代工")
i_c = next((i for i, b in enumerate(va["blocks"]) if b.get("type") == "bars"), 0)
va["blocks"][i_c + 1:i_c + 1] = [{"type": "note", "parts": [f"@README!A{r_frame}"]},
                                 {"type": "table", "caption": "期間並列與方法占比", "head": ["項目", "數值", "說明"],
                                  "rows": [[f"@Val!A{r_p1}", f"@Val!B{r_p1}|n2", f"@Val!E{r_p1}"], [f"@Val!A{r_p2}", f"@Val!B{r_p2}|n2", f"@Val!E{r_p2}"],
                                           [f"@Val!A{r_sh}", f"@Val!B{r_sh}|p1", f"@Val!E{r_sh}"], [f"@Val!A{r_s1}", f"@Val!B{r_s1}|x1", ""], [f"@Val!A{r_s2}", f"@Val!B{r_s2}|x1", ""]]}]
# 敏感度章：目標價表加 DCF 單法美元欄；業務主因（整本重算快照）
se = SEC["sens"]
r_tp = find("Sens", "目標價參數敏感度"); r_end = r_tp + 2
while SEW.cell(r_end + 1, 1).value and not str(SEW.cell(r_end + 1, 1).value).startswith("檢查"): r_end += 1
for b in se["blocks"]:
    if b.get("type") == "table" and b.get("rowgen", {}).get("sheet") == "Sens" and str(b["rowgen"].get("rows", "")).startswith(f"{r_tp + 2}-"):
        b["head"] = [f"@Sens!{c}{r_tp + 1}" for c in "ABCDEFGIJ"]
        b["rowgen"] = {"sheet": "Sens", "rows": f"{r_tp + 2}-{r_end}", "cols": list("ABCDEFGIJ"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2", "n2", "n2"]}
r_bz = find("Sens", "目標價敏感度：業務主因")
r_bh = r_bz + 2; r_be = r_bh
while SEW.cell(r_be + 1, 1).value: r_be += 1
se["blocks"].insert(0, {"type": "table", "caption": f"@Sens!A{r_bz}", "head": [f"@Sens!{c}{r_bh}" for c in "ABCDEFGIJ"],
                        "rowgen": {"sheet": "Sens", "rows": f"{r_bh + 1}-{r_be}", "cols": list("ABCDEFGIJ"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2", "n2", "n2"]}})
# 對帳章：反解（F 節）、毛利率差距分解（G 節）、供應承諾（H 節）
def _sect(prefix):
    try:
        r0 = find("Recon", prefix)
    except KeyError:
        return None
    r1 = r0
    while RCW.cell(r1 + 1, 1).value and not str(RCW.cell(r1 + 1, 1).value)[:2] in ("G.", "H.", "I."): r1 += 1
    return r0, r1
s_f = _sect("F. 反解")
r_e, r_ee = s_f; r_eh = next(r for r in range(r_e + 1, r_ee + 1) if RCW.cell(r, 1).value == "項目")
def _rf(r):
    t = str(RCW.cell(r, 1).value)
    if t.startswith("評等級距邊界"): return [f"@Recon!A{r}", f"@Recon!B{r}|sp0", "", ""]
    f = "x1" if "倍數（" in t else "n2"
    return [f"@Recon!A{r}", f"@Recon!B{r}|{f}", f"@Recon!C{r}|{f}", f"@Recon!D{r}|sp1"]
SEC["recon"]["blocks"].append({"type": "table", "caption": f"@Recon!A{r_e}", "head": [f"@Recon!{c}{r_eh}" for c in "ABCD"], "rows": [_rf(r) for r in range(r_e + 1, r_ee + 1) if r != r_eh]})
s_g = _sect("G. FY2028 毛利率")
if s_g:
    g0, g1 = s_g
    rows = []
    for r in range(g0 + 2, g1 + 1):
        t = str(RCW.cell(r, 1).value)
        if t.startswith("k_rec 持續性"): continue
        f = "n2" if t.startswith(("反解", "檢查")) else "n2"
        rows.append([f"@Recon!A{r}", f"@Recon!B{r}|{f}", f"@Recon!J{r}"])
    SEC["recon"]["blocks"].append({"type": "table", "caption": f"@Recon!A{g0}", "head": ["項目", "百分點／值", "說明"], "rows": rows})
    SEC["recon"]["blocks"].append({"type": "note", "parts": [f"@Recon!A{find('Recon', 'k_rec 持續性')}"]})
s_h = _sect("H. 供應預付款")
if s_h:
    h0, h1 = s_h
    SEC["recon"]["blocks"].append({"type": "table", "caption": f"@Recon!A{h0}", "head": ["項目", "值", "說明"],
                                   "rows": [[f"@Recon!A{r}", f"@Recon!B{r}" + ("|n0" if isinstance(RCW.cell(r, 2).value, (int, float)) else ""), f"@Recon!J{r}"] for r in range(h0 + 1, h1 + 1)]})
# 估值章：EBIT 口徑與現金稅率（摺疊）
DW = wb["DCF"]
r_bh = find("DCF", "EBIT 口徑與現金稅率")
va["blocks"].append({"type": "table", "caption": f"@DCF!A{r_bh}", "head": ["項目", "數值", "對照", "判定"],
                     "rows": [[f"@DCF!A{r_bh + 2}", f"@DCF!B{r_bh + 2}|p1", f"@DCF!C{r_bh + 2}", f"@DCF!D{r_bh + 2}"]], "collapse": True, "summary": "展開：EBIT 口徑與現金稅率"})

# v0.3.6（2026-11-12）：分區主題圖（Andy 2026-11-11 選定四張、11-12 08:38 認可樣板）；列號依標籤找，不寫死
def _find_after(sh, start, prefix, col="A"):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        v = ws[f"{col}{r}"].value
        if isinstance(v, str) and v.strip().startswith(prefix):
            return r
    raise KeyError((sh, start, prefix))


def _theme_blocks():
    ua = find("Units", "情境 A")
    amt = _find_after("Units", ua, "金額（顆數")
    tot = _find_after("Units", amt + 1, "合計")
    shr = _find_after("Units", tot, "金額份額")
    units = {"type": "theme", "title": "加速器市場金額（情境 A，US$B）與 NVIDIA 金額份額", "x": "@Units!C3:F3",
             "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.35, "total": f"@Units!C{tot}:F{tot}",
             "bars": [{"name": f"@Units!A{r}", "y": f"@Units!C{r}:F{r}"} for r in range(amt + 1, tot)],
             "lines": [{"name": ["NVIDIA 金額份額"], "y": f"@Units!C{shr + 1}:F{shr + 1}", "axis": "r", "labels": True}],
             "caption": "堆疊＝各平台加速器金額（顆數 × 每顆售價，模組層級）；折線＝NVIDIA 占加速器總金額的比例（右軸）。"}
    pa, pb, pc = find("PL", "情境 A"), find("PL", "情境 B"), find("PL", "情境 C")
    rev = {k: _find_after("PL", v, "營收") for k, v in (("A", pa), ("B", pb), ("C", pc))}
    gm = {k: _find_after("PL", v, "毛利率") for k, v in (("A", pa), ("B", pb), ("C", pc))}
    dc, edge = _find_after("PL", pb, "Data Center"), _find_after("PL", pb, "Edge Computing")
    hbm = _find_after("PL", find("PL", "情境 B："), "HBM 成本占由下而上營收")
    pl = {"type": "theme", "title": "營收分部（情境 B）與三情境營收（US$B）", "x": "@PL!C3:G3", "fmt": "n0", "stack": True,
          "bars": [{"name": f"@PL!A{dc}", "y": f"@PL!C{dc}:G{dc}"}, {"name": f"@PL!A{edge}", "y": f"@PL!C{edge}:G{edge}", "color": "#c29a2b"}],
          "lines": [{"name": [f"@PL!A{pa}", "營收"], "y": f"@PL!C{rev['A']}:G{rev['A']}", "color": "#0b7a80", "dash": True},
                    {"name": [f"@PL!A{pb}", "營收"], "y": f"@PL!C{rev['B']}:G{rev['B']}", "color": "#3b3838", "labels": True},
                    {"name": [f"@PL!A{pc}", "營收"], "y": f"@PL!C{rev['C']}:G{rev['C']}", "color": "#b4472f", "dash": True}],
          "caption": "堆疊長條＝情境 B 的 Data Center 與 Edge Computing；三條線＝情境 A／B／C 的合併營收。FY2026A 為實績、FY2027 為錨點。"}
    gmc = {"type": "theme", "title": "毛利率（三情境）與 HBM 成本占營收（情境 B）", "x": "@PL!D3:G3", "fmt": "p0", "stack": False,
           "bars": [{"name": [f"@PL!A{hbm}", "（情境 B）"], "y": f"@PL!D{hbm}:G{hbm}", "color": "#c29a2b"}],
           "lines": [{"name": [f"@PL!A{pa}", "毛利率"], "y": f"@PL!D{gm['A']}:G{gm['A']}", "color": "#0b7a80", "dash": True},
                     {"name": [f"@PL!A{pb}", "毛利率"], "y": f"@PL!D{gm['B']}:G{gm['B']}", "color": "#3b3838", "labels": "below"},
                     {"name": [f"@PL!A{pc}", "毛利率"], "y": f"@PL!D{gm['C']}:G{gm['C']}", "color": "#b4472f", "dash": True}],
           "caption": "HBM 成本占營收是毛利率的主要變動來源之一；兩者並列，可看記憶體漲價對毛利的壓力在哪一年最大。"}
    rg, rc = find("Recon", "FY2028 營收 vs 指引"), find("Recon", "FY2028 營收 vs 共識")
    mg, ep = find("Recon", "FY2028 毛利率 vs 指引"), find("Recon", "FY2028 EPS vs 共識")
    cmp = {"type": "compare", "title": "FY2028：模型 vs 公司指引 vs 市場共識", "groups": [
        {"name": ["營收（US$B）"], "fmt": "n0", "items": [{"name": ["公司指引推算"], "v": f"@Recon!B{rg}", "ref": True}, {"name": ["市場共識推算"], "v": f"@Recon!B{rc}", "ref": True},
                                                       {"name": ["模型 情境 A"], "v": f"@Recon!C{rg}"}, {"name": ["模型 情境 B"], "v": f"@Recon!E{rg}"}, {"name": ["模型 情境 C"], "v": f"@Recon!G{rg}"}]},
        {"name": ["毛利率"], "fmt": "p1", "items": [{"name": ["公司指引中點"], "v": f"@Recon!B{mg}", "ref": True},
                                                 {"name": ["模型 情境 A"], "v": f"@Recon!C{mg}"}, {"name": ["模型 情境 B"], "v": f"@Recon!E{mg}"}, {"name": ["模型 情境 C"], "v": f"@Recon!G{mg}"}]},
        {"name": ["EPS（US$）"], "fmt": "n2", "items": [{"name": ["市場共識"], "v": f"@Recon!B{ep}", "ref": True},
                                                      {"name": ["模型 情境 A"], "v": f"@Recon!C{ep}"}, {"name": ["模型 情境 B"], "v": f"@Recon!E{ep}"}, {"name": ["模型 情境 C"], "v": f"@Recon!G{ep}"}]}],
        "caption": "灰色＝公司指引或市場共識；青綠＝本模型三情境。對照只作事後檢查，校準時未使用。"}
    S = {s["id"]: s for s in cfg["sections"]}
    S["drivers"]["blocks"].insert(0, units)
    S["pl"]["blocks"][0:0] = [pl, gmc]
    S["recon"]["blocks"].insert(0, cmp)


_theme_blocks()


# ================= v0.3.7（2026-11-14）：個股 HTML 版面規格 v1.3 試作（Andy 2026-11-14 08:11；docs/plan/個股HTML版面規格.md「v1.3 規格」） =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳 → 資料狀態；NVIDIA 無個別模組。
# 數字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁；nvidia/v037_post.py 產生）。列號依標籤找。
def _v13():
    S = {s["id"]: s for s in cfg["sections"]}
    RD, PLW, RV = wb["README"], wb["PL"], wb["RevDCF"]
    r13 = find("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.3.7"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    # ---- 橫幅：只留 ≤200 字白話（規則與歸因移到估值章）----
    old_note = cfg["hero"]["note"]
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"]
    # ---- 1. 命題與核心問題 ----
    th = S["thesis"]["blocks"]
    units = next(b for b in S["drivers"]["blocks"] if b["type"] == "theme")
    S["drivers"]["blocks"].remove(units)
    units["caption"] = ("堆疊＝各平台加速器金額（顆數 × 每顆售價，模組層級）；折線＝NVIDIA 占加速器總金額的比例（右軸）。"
                        "回答命題一（ASIC 會把 NVIDIA 的金額份額壓到多少）；情境 A 為上游原始顆數，基準情境 B 的顆數在 CY2028 起較低。")
    kv = next(b for b in th if b["type"] == "kv")
    kv["items"] = [it for it in kv["items"] if it[0] != "會計年度"]
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th[0] = {"type": "note", "parts": [f"@README!A{RM['rec']}"]}           # 投資建議條件式（白話；原 README「投資建議：」列移到資料狀態）
    th.insert(0, units)
    for b in th:
        if b["type"] == "h3" and b.get("text") == "情境定義與會計年度":
            b["text"] = "情境定義"
    # ---- 2. 估值：橫幅原有的評等規則、相對倍數調整與歸因說明移到這裡（摺疊）----
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b["type"] == "bars")
    # 版面規格 v1.3 G 節（Andy 2026-11-14 09:03）：估值章開頭並排四個目標價（表＋長條圖，現價線）；數字取自 Val 頁 0.7 節
    g0 = find("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = find("Val", "現價（2026")
    va[i_bar] = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
                 "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                           ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
                 "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
    va.insert(i_bar, {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
                      "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr]
                              + [[f"@Val!A{g0 + 7}", f"@Val!B{g0 + 7}|n2", "", f"@Val!D{g0 + 7}"]]})
    i_bar += 1
    va.insert(i_bar + 1, {"type": "note", "parts": old_note[:7]})
    va.insert(i_bar + 2, {"type": "list", "items": [p for p in old_note[7:] if p], "collapse": True,
                          "summary": "展開：評等邊界、上檔空間三段歸因、現行相對倍數與公司指引對照"})
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").startswith(p))
    ra = rv("A. DCF"); rb_ = rv("(a)"); rbb = rv("(b)"); rc = rv("(c)"); rcmp = rv("B. Comparable"); rsum = rv("C. 摘要")
    yh = rv("項目（情境 B）")
    R = lambda p, start: rv(p, start)
    m1, m0 = R("現價隱含的營業利益率", rb_), R("模型：基準的等效營業利益率", rb_)
    s1, rv1, rv0, g1, g0 = R("現價隱含的營收倍數", rbb), R("現價隱含的 FY2028 營收", rbb), R("模型：FY2028 營收", rbb), R("現價隱含的 FY2028 營收年增率", rbb), R("模型：FY2028 營收年增率", rbb)
    w1, w0 = R("現價隱含的 WACC", rc), R("模型 WACC", rc)
    rev_ = {
        "id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": f"@RevDCF!A2",
        "blocks": [
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
            {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
             "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                      ["營收規模（FY2028 起各年同乘的倍數）", f"@RevDCF!B{s1}|n2", "—"],
                      ["FY2028 營收（十億美元）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                      ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0}|p1"],
                      ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
            {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"],
             "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if r in (ra + 2, ra + 3, ra + 9) else "n3" if r == ra + 5 else "x1" if r == ra + 8 else "n1")]
                      for r in range(ra + 2, ra + 10)]},
            {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
             "rows": [[f"@RevDCF!A{rcmp + 2}", "—", f"@RevDCF!C{rcmp + 2}|n2", "—"]]
                     + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!C{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!D{r}|sp1"] for k, r in enumerate(range(rcmp + 3, rcmp + 6))]
                     + [[f"@RevDCF!A{rcmp + 6}", "—", f"@RevDCF!C{rcmp + 6}|n1", "—"]]},
            {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
             "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）",
             "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|" + ("p1" if r in (yh + 3, yh + 4) else "n3" if r in (yh + 5, yh + 6, yh + 7, yh + 8) else "n1") for c in "BCDEFG"]
                      for r in range(yh + 1, rb_ - 1)]
                     + [{"sub": f"@RevDCF!A{rb_}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "利益率" in str(RV.cell(r, 1).value) else "n1"), "", "", "", "", ""] for r in range(rb_ + 1, rb_ + 5)]
                     + [{"sub": f"@RevDCF!A{rbb}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "年增率" in str(RV.cell(r, 1).value) else "n2" if "倍數" in str(RV.cell(r, 1).value) else "n3" if "核對" in str(RV.cell(r, 1).value) else "n1"), "", "", "", "", ""] for r in range(rbb + 1, rc - 1)]
                     + [{"sub": f"@RevDCF!A{rc}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p2" if str(RV.cell(r, 1).value).startswith(("現價隱含的 WACC", "模型 WACC")) else "n2"), "", "", "", "", ""] for r in range(w1, w1 + 4)]},
        ]}
    # ---- 5. 損益（標準損益表）----
    p0 = find("PL", "標準損益表（v0.3.7")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（研發＋銷管）", "營業利益", "營業利益率",
            "其他收入（淨額，不含權益證券損益）", "稅前利益", "所得稅", "稅後淨利", "淨利率", "稀釋股數", "EPS（完全稀釋）"]
    fmt = lambda u: {"%": "p1", "US$": "n2", "十億股": "n2"}.get(u, "n1")
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in "CDEFGHIJ"]
    r_ratio = plr("比率"); r_gaap = next(r for r in range(ph, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("GAAP 對照"))
    r_chk = plr("核對（應為 0）")
    r_d31 = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("差額（FY2031E）：毛利"))
    is_tbl = {"type": "table", "caption": "標準損益表（十億美元；EPS 為美元）：FY2024A–FY2026A 實績、FY2027E–FY2031E 情境 B（基準）",
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in "CDEFGHIJ"],
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_gaap)]
                      + [{"sub": "核對：FY2031E 兩種推法的營業利益差額（毛利率與營業費用率沿用 FY2030，Andy 決定）"}]
                      + [rowx(r) for r in (r_d31, r_d31 + 1)]}
    gaap_tbl = {"type": "table", "caption": f"@PL!A{r_gaap}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in "CDE"], "collapse": True,
                "summary": "展開：GAAP 對照（含權益證券損益；只列實績）與核對列",
                "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in "CDE"] for r in range(r_gaap + 1, r_chk)]
                        + [{"sub": f"@PL!A{r_chk}"}] + [[f"@PL!A{r}"] + [f"@PL!{c}{r}|n3" for c in "CDE"] for r in range(r_chk + 1, r_chk + 7) if r not in (r_d31, r_d31 + 1)]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率、淨利率：FY2024A–FY2031E（情境 B）", "x": f"@PL!C{ph}:J{ph}",
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:J{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:J{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:J{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:J{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，Non-GAAP 模型口徑）。FY2031E 為延伸期：營收與淨利率由收斂法推得，毛利率與營業費用率沿用 FY2030（Andy 決定）。"}
    x0 = find("PL", "三情境延伸到 FY2032"); xh = x0 + 2
    xa, xb, xc = (next(r for r in range(x0, PLW.max_row + 1) if PLW.cell(r, 1).value == f"情境 {k}") for k in "ABC")
    xr = lambda base, lab_: next(r for r in range(base, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
    sc_chart = {"type": "theme", "title": "三情境並列：營收（長條，十億美元）與 EPS（折線，美元，右軸）", "x": f"@PL!C{xh}:H{xh}",
                "fmt": "n0", "fmt_r": "n2", "stack": False, "pad": 1.1,
                "bars": [{"name": [f"@PL!A{xa}", " 營收"], "y": f"@PL!C{xr(xa, '營收')}:H{xr(xa, '營收')}", "color": "#b7e3e4"},
                         {"name": [f"@PL!A{xb}", " 營收"], "y": f"@PL!C{xr(xb, '營收')}:H{xr(xb, '營收')}", "color": "#10a0a7"},
                         {"name": [f"@PL!A{xc}", " 營收"], "y": f"@PL!C{xr(xc, '營收')}:H{xr(xc, '營收')}", "color": "#8a8c8c"}],
                "lines": [{"name": [f"@PL!A{xa}", " EPS"], "y": f"@PL!C{xr(xa, 'EPS（延伸表）')}:H{xr(xa, 'EPS（延伸表）')}", "axis": "r", "color": "#0b7a80", "dash": True},
                          {"name": [f"@PL!A{xb}", " EPS"], "y": f"@PL!C{xr(xb, 'EPS（延伸表）')}:H{xr(xb, 'EPS（延伸表）')}", "axis": "r", "color": "#3b3838", "labels": True},
                          {"name": [f"@PL!A{xc}", " EPS"], "y": f"@PL!C{xr(xc, 'EPS（延伸表）')}:H{xr(xc, 'EPS（延伸表）')}", "axis": "r", "color": "#b4472f", "dash": True}],
                "caption": "情境 A（樂觀）、B（基準）、C（保守）在 FY2027 相同（錨點），差異自 FY2028 起出現、主要從 FY2029 起；FY2031 起為延伸期（收斂法）。"}
    old_pl = S["pl"]["blocks"]
    gmc = next(b for b in old_pl if b["type"] == "theme" and "毛利率" in str(b.get("title")))
    old_rest = [b for b in old_pl if b["type"] != "theme"]
    pls = [b for b in old_rest if b["type"] == "pl"]
    notes = [p_ for b in old_rest if b["type"] == "note" for p_ in b["parts"]]
    assert len(pls) == 2 and len(notes) == 2, (len(pls), notes)
    old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境摘要表與折線（FY2026–FY2030；營收、毛利率、營業利益、EPS）"),
                  {"type": "list", "items": notes, "collapse": True, "summary": "展開：延伸期（FY2031–FY2032）的收斂法說明"},
                  dict(pls[1], collapse=True, summary="展開：三情境延伸到 FY2032（收斂法）")]
    pl_blocks = ([{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, sc_chart, gaap_tbl]
                 + old_blocks)
    S["pl"].update(title="損益（標準損益表）", blocks=pl_blocks)
    S["pl"].pop("intro", None)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    rbB = find("PL", "情境 B："); pr = lambda lab_: next(r for r in range(rbB, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith(lab_))
    hdrB = rbB + 1
    gens = [r for r in range(rbB, rbB + 30) if str(PLW.cell(r, 1).value or "").endswith(" 認列顆數") and str(PLW.cell(r, 1).value).startswith("NVIDIA ")]
    prices = [r for r in range(rbB, rbB + 30) if str(PLW.cell(r, 1).value or "").endswith(" 每顆售價")]
    comp = [r for r in range(rbB, rbB + 30) if str(PLW.cell(r, 1).value or "").endswith(" 運算營收")]
    r_ut, r_cm, r_cn, r_nt, r_ed, r_bu = pr("NVIDIA 認列顆數合計"), pr("運算營收（Σ"), pr("中國資料中心運算營收"), pr("NVIDIA 網通營收"), pr("Edge Computing 營收"), pr("由下而上營收合計")
    r_rep = next(r for r in range(r_bu, PLW.max_row + 1) if PLW.cell(r, 1).value == "營收")
    r_dc, r_eg = r_rep + 1, r_rep + 2
    r_sg = find("PL", "差額：由下而上營收合計")
    seg_rows = lambda rs, f: [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{f}" for c in "DEFG"] for r in rs]
    seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；十億美元）：資料中心運算＋網通＋中國＋Edge Computing＝由下而上合計（與報告值的差額見末列）",
               "head": [f"@PL!A{hdrB}", f"@PL!B{hdrB}"] + [f"@PL!{c}{hdrB}" for c in "DEFG"],
               "rows": [{"sub": "一、資料中心運算＝各世代認列顆數 × 每顆售價"}] + seg_rows(gens + [r_ut], "n2") + seg_rows(prices, "n1") + seg_rows(comp + [r_cm], "n1")
                       + [{"sub": "二、網通（隨 NVIDIA 顆數同比例縮放）"}] + seg_rows([r_nt], "n1")
                       + [{"sub": "三、中國資料中心（基準為零，只進敏感度）"}] + seg_rows([r_cn], "n1")
                       + [{"sub": "四、Edge Computing（遊戲、PC、工作站、車用與機器人）"}] + seg_rows([r_ed], "n1")
                       + [{"sub": "合計：由下而上與報告值（FY2027 報告值為錨點）"}] + seg_rows([r_bu, r_rep, r_dc, r_eg], "n1")
                       + [[f"@PL!A{r_sg}", "@PL!B{}".format(r_rep)] + [f"@PL!{c}{r_sg}|n3" for c in "DEFG"]]}
    seg_chart = {"type": "theme", "title": "分部營收（情境 B，十億美元）與 NVIDIA 認列顆數（百萬顆，右軸）", "x": f"@PL!D{hdrB}:G{hdrB}",
                 "fmt": "n0", "fmt_r": "n1", "stack": True, "pad": 1.2, "total": f"@PL!D{r_bu}:G{r_bu}",
                 "bars": [{"name": ["資料中心運算"], "y": f"@PL!D{r_cm}:G{r_cm}", "color": "#0b7a80"},
                          {"name": ["網通"], "y": f"@PL!D{r_nt}:G{r_nt}", "color": "#6fc7cb"},
                          {"name": ["Edge Computing"], "y": f"@PL!D{r_ed}:G{r_ed}", "color": "#c29a2b"}],
                 "lines": [{"name": [f"@PL!A{r_ut}"], "y": f"@PL!D{r_ut}:G{r_ut}", "axis": "r", "color": "#3b3838", "labels": True}],
                 "caption": "堆疊＝資料中心運算、網通、Edge Computing 三個分部（中國基準為零，未畫）；頂端數字＝由下而上合計；折線＝NVIDIA 認列顆數。FY2027 為錨點年（由下而上校準到公司實績與指引）。"}
    gmc["caption"] = "HBM 成本占營收是毛利率的主要變動來源之一；兩者並列，可看記憶體漲價對毛利率的壓力在哪一年最大（模型假設漲價有一半由 NVIDIA 吸收）。"
    dr = S["drivers"]["blocks"]
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "毛利率的驅動：HBM 與 LPDDR 成本"}, gmc,
                              {"type": "h3", "text": "情境驅動與單一參數"}] + dr
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    # ---- 9. 資料狀態：會計年度、原投資建議條件式、本版要點 ----
    st = S["status"]["blocks"]
    st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]},
               {"type": "h3", "text": "會計年度與原投資建議條件式（模型原文）"}, {"type": "note", "parts": [f"@README!A{find('README', '會計年度：')}"]},
               {"type": "note", "parts": [f"@README!A{find('README', '投資建議：')}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    S["revdcf"] = rev_
    cfg["sections"] = [S[k] for k in order]


_v13()


# v0.3.7：全頁可見文字（摺疊與資料狀態以外）不出現參數代號（版面規格 v1.3 B 節原則）：
#   含代號的儲存格改用白話（設定檔文字，不含數字）；含代號的整段說明改為摺疊。xlsx 原文不改（逐格比對 0 差異）。
PLAIN = {
    "Val!A10": "永續成長率（長期成長上限與無風險利率取小者）",
    "Val!A11": "永續期新投資報酬率 RONIC（＝WACC＋超額報酬差）",
    "Val!A20": "DCF 情境價值推到目標時點（以股權成本複利、扣除期間股利）",
    "Val!E9": "股權成本與稅後債務成本以市值權重加權；WACC 調整只用於敏感度（基準為零）",
    "Val!E40": "DCF 與 Comparable 加權（權重見 In 頁目標價規則，Andy 可改）",
    "Val!E202": "DCF 每股價值以股權成本由評價日推到目標時點",
    "In!B176": "永續成長率（最後一年之後；長期成長上限與無風險利率取小者）",
    "Sens!B37": "DCF 權重 低／高",
    "Sens!B991": "所有 NVIDIA 世代顆數同乘 低／高",
    "Sens!A1002": "營業費用彈性（營業費用年增 ÷ 營收年增）低／高",
    "Sens!B992": "售價改變後，FY2027 校準係數整本重算",
    "Sens!A996": "匣級加價保留比例（全不保留／全保留）",
    "Sens!B996": "校準係數隨之重算",
    "Sens!A997": "HBM／LPDDR 漲價吸收比例（全額轉嫁／全額吸收；三情境）",
    "Sens!B1001": "資料中心錨點＝總營收 − Edge；Edge 誤差傳到錨點與校準係數",
    "Sens!B1004": "營業費用額外增減占營收的比例 低／高",
    "Sens!A1005": "有效稅率 低／高（公司 FY2027 指引區間）",
    "Sens!B1005": "公司 FY2027 指引區間",
    "Sens!B1008": "錨點改變，校準係數整本重算",
    "Sens!A1016": "FY2027 校準缺口自 FY2028 起逐年收斂、FY2030 消失（基準＝缺口持續）",
    "Sens!B1016": "只有高（整本重算）",
    "Recon!A59": "FY2028 毛利率（情境 B）− 公司 FY2028 毛利率指引中點",
    "Recon!A60": "① HBM／LPDDR 吸收比例（基準吸收一半 vs 全額轉嫁）",
    "Recon!J63": "Rubin／Rubin Ultra 占比與非 HBM 成本係數校準",
    "Recon!A64": "⑤ FY2027 錨點毛利率 − 指引中點",
    "Recon!A67": "反解：FY2028 毛利率（情境 B）達指引下緣所需的漲價吸收比例（只展示，基準不變）",
    "Recon!J72": "公司未單獨揭露供應預付，含在預付及其他流動資產與其他資產內；營運資金比率因此無法剔除供應預付",
}
FOLD = {"DCF!A2": "展開：DCF 的營業利益與自由現金流口徑", "Sens!A24": "展開：敏感度讀法（月數加權與售價案例的口徑）",
        "Recon!A68": "展開：FY2027 校準缺口的持續性假設"}


def _plain(spec):
    if isinstance(spec, list):
        return [_plain(x) for x in spec]
    if isinstance(spec, str) and spec.startswith("@"):
        ref = spec[1:].split("|")[0]
        if ref in PLAIN:
            return PLAIN[ref]
    return spec


def _rowgen_rows(b):
    g = b.pop("rowgen"); ws = wb[g["sheet"]]; fm = g.get("fmts", [""] * len(g["cols"])); rows = []
    a, z = (int(x) for x in str(g["rows"]).split("-"))
    for r in range(a, z + 1):
        vals = [ws[f"{c}{r}"].value for c in g["cols"]]
        cells = [f"@{g['sheet']}!{c}{r}|{f}" for c, f in zip(g["cols"], fm)]
        rows.append({"sub": cells[0]} if (len(vals) > 1 and vals[0] is not None and all(v is None for v in vals[1:])) else cells)
    b["rows"] = rows


for sec in cfg["sections"]:
    if sec["id"] == "status":
        continue
    nb = []
    for b in sec["blocks"]:
        if b.get("type") == "table" and "rowgen" in b and b["rowgen"].get("sheet") == "Sens":
            _rowgen_rows(b)
        if b.get("type") == "table":
            b["rows"] = [({"sub": _plain(r["sub"])} if isinstance(r, dict) else _plain(r)) for r in b.get("rows", [])]
            if "head" in b: b["head"] = _plain(b["head"])
        if b.get("type") == "note" and not b.get("collapse") and len(b["parts"]) == 1 and str(b["parts"][0])[1:] in FOLD:
            b = dict(b, collapse=True, summary=FOLD[str(b["parts"][0])[1:]])
        nb.append(b)
    sec["blocks"] = nb
_S = {s["id"]: s for s in cfg["sections"]}
_S["sens"]["intro"] = "上方表為加權目標價的敏感度（業務主因整本重算、目標價參數）；下方龍捲風圖為 FY2028 EPS 的敏感度（以情境 A 為底，一次動一個輸入）。"
_S["sens"]["blocks"].insert(0, {"type": "note", "parts": ["@Sens!A1"], "collapse": True, "summary": "展開：敏感度頁原文說明（含模型參數代號）"})
# 敏感度章的主圖：加權目標價在主要案例下的值（長條）與基準（豎線）；案例取業務主因表與目標價參數表中相對基準變動達門檻者
def _sens_bars(th=0.02, top=10):
    SW = wb["Sens"]; cand = []
    for t0 in ("目標價敏感度：業務主因", "目標價參數敏感度"):
        r0 = find("Sens", t0); hr = next(r for r in range(r0, r0 + 4) if SW.cell(r, 1).value in ("案例", "參數"))
        r = hr + 1
        while SW.cell(r, 1).value and not str(SW.cell(r, 1).value).startswith("檢查"):
            lo, hi = SW.cell(r, 4).value, SW.cell(r, 6).value
            mags = [abs(v) for v in (lo, hi) if isinstance(v, (int, float))]
            if mags and max(mags) >= th:
                cand.append((max(mags), r, lo, hi))
            r += 1
    cand.sort(key=lambda t: (-t[0], t[1]))
    items = []
    SHORT = {"相對倍數係數＝1": ("同業比較法不加歷史溢價", None), "DCF 權重": ("DCF 權重調低", "DCF 權重調高"),
             "k_rec 自 FY2028 起": (None, "FY2027 校準缺口逐年收斂"), "NVIDIA 顆數係數": ("NVIDIA 顆數少一成", "NVIDIA 顆數多一成"),
             "中國恢復": ("中國恢復（低）", "中國恢復（高）"), "GPU 模組售價：Blackwell": ("Blackwell 售價低", "Blackwell 售價高"),
             "GPU 模組售價：Rubin Ultra": ("Rubin Ultra 售價低", "Rubin Ultra 售價高"), "GPU 模組售價：Rubin 低": ("Rubin 售價低", "Rubin 售價高"),
             "匣級加價保留比例": ("匣級加價全不保留", "匣級加價全保留"), "營業費用率": ("營業費用率調低", "營業費用率調高"),
             "ERP": ("股權風險溢酬低", "股權風險溢酬高"), "WACC": ("WACC 調低", "WACC 調高"), "β": ("β 低", "β 高"),
             "HBM／LPDDR 吸收比例": ("漲價全額轉嫁", "漲價全額吸收"), "永續成長率": ("永續成長率低", "永續成長率高")}
    for _, r, lo, hi in cand[:top]:
        t = str(SW.cell(r, 1).value)
        lh = next((v for k, v in SHORT.items() if t.startswith(k)), None)
        for i, (col, v) in enumerate((("C", lo), ("E", hi))):
            if isinstance(v, (int, float)):
                labl = lh[i] if (lh and lh[i]) else [PLAIN.get(f"Sens!A{r}", f"@Sens!A{r}"), ("（低）", "（高）")[i]]
                items.append({"label": labl, "value": f"@Sens!{col}{r}", "series": 0 if v > 0 else 1})
    return {"type": "bars", "title": "加權目標價的主要敏感度（長條＝該案例下的加權目標價；豎線＝基準加權目標價；依影響大小排序）",
            "fmt": "n2", "items": items, "ref": {"label": "基準加權目標價", "value": f"@Val!B{find('Val', '目標價（DCF 與 Comparable 加權）')}"}}


# v0.3.7（Andy 2026-11-14 09:16，樣板 CoreWeave v5.0.2；版面規格 v1.3 H 節）：敏感度主圖改為「加權目標價相對基準」龍捲風圖（低案、高案各一，
# |變動| ≥ 10% 直接顯示，未達者摺疊；全部案例表摺疊）。資料取 Sens 頁末「加權目標價相對基準」表（v037_post.sens_tor）。原 top 10 長條圖刪除，FY2028 EPS 龍捲風圖改摺疊對照。
def _sens_tor():
    SW = wb["Sens"]; t0 = find("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SW.cell(r2 + 1, 1).value: r2 += 1
    out = []
    for col, side in (("B", "低"), ("C", "高")):
        big = [r for r in range(r1, r2 + 1) if isinstance(SW[f"{col}{r}"].value, (int, float)) and abs(SW[f"{col}{r}"].value) >= 0.1]
        small = [r for r in range(r1, r2 + 1) if isinstance(SW[f"{col}{r}"].value, (int, float)) and 0.0005 <= abs(SW[f"{col}{r}"].value) < 0.1]   # 未達 0.05% 的案例不畫（長條寬度為零），只在全部案例表
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        out.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
        out.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                        summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例（幾乎為零者見全部案例表）"))
    out.append({"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
                "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1, r2 + 1)]})
    return out


_sb = _S["sens"]["blocks"]
for _b in _sb:
    if _b.get("type") == "table" and not _b.get("collapse") and str(_b.get("caption", "")).startswith(("@Sens!A987", "@Sens!A26")):
        _b["collapse"] = True
        _b["summary"] = "展開：業務主因表（整本重算快照；加權目標價與 DCF 單法變化）" if _b["caption"] == "@Sens!A987" else "展開：目標價參數表（WACC、β、ERP、永續成長率等）"
    if _b.get("type") == "tornado":
        _b["summary"] = "展開：對照：FY2028 EPS 相對基準的變動（情境 A，一次動一個輸入）"
_S["sens"]["blocks"] = _sb[:1] + _sens_tor() + _sb[1:]
_S["sens"]["intro"] = "主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入；變動達兩位數百分比者直接顯示，其餘摺疊）；下方摺疊為全部案例、業務主因與目標價參數明細，以及 FY2028 EPS 敏感度對照。"

# ================= v0.3.8 試作（2026-11-14）：三層協調——對帳中標示與兩讀數目標價（三層基準協調規則 v1.1 C4；版面規格 v1.3 I 節） =================
# 數字與文字全部取自 xlsx 的 Coord 頁（coord_post.py）：橫幅加「另一側讀數」一格與揭露句；估值章開頭並列兩讀數（表＋長條圖）；標準損益表預測年度欄標「對帳中」。
def _coord():
    if "Coord" not in wb.sheetnames:
        return
    CW = wb["Coord"]
    cr = lambda p, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
    if CW[f"B{cr('是否有估值依賴的對帳中年份')}"].value != "是":
        return
    S = {x["id"]: x for x in cfg["sections"]}
    r_tp, r_up, r_rec = cr("加權目標價"), cr("上檔空間（加權目標價"), cr("投資建議（五級")
    r_hero, r_cur, r_alt, r_note, r_pl, r_yrs = cr("橫幅揭露"), cr("目標價區標籤：現行"), cr("目標價區標籤：另一側"), cr("目標價區說明"), cr("損益表年度標示"), cr("對帳中年份")
    # 橫幅：另一側讀數一格＋揭露句（≤200 字合計）
    cfg["hero"]["tiles"].append([["另一側讀數目標價（對帳中；美元）"], f"@Coord!C{r_tp}|n2", ["上檔空間 ", f"@Coord!C{r_up}|sp1", "；", f"@Coord!C{r_rec}"]])
    cfg["hero"]["note"] = list(cfg["hero"]["note"]) + [" ", f"@Coord!B{r_hero}"]
    # 估值章開頭：兩讀數目標價
    rows_ = [cr("NVIDIA 認列顆數（情境 B）FY2028"), cr("營收（情境 B）FY2028"), cr("EPS 情境 B（基準）FY2028"), cr("EPS 情境 B（基準）FY2029"), cr("EPS 情境 B（基準）FY2030"),
             cr("12 個月後前瞻 EPS"), cr("DCF 目標價（情境加權）"), cr("Comparable 目標價（進加權"), r_tp, r_up]
    fm = lambda r: "p1" if r == r_up else "n2"
    tbl = {"type": "table", "caption": ["兩讀數目標價（對帳中：", f"@Coord!B{r_yrs}", "）——現行基準讀數與另一側讀數並列"],
           "head": ["項目", f"@Coord!B{r_cur}", f"@Coord!B{r_alt}", "另一側相對現行"],
           "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|{fm(r)}", f"@Coord!C{r}|{fm(r)}", f"@Coord!D{r}|sp1"] for r in rows_]
                   + [[f"@Coord!A{r_rec}", f"@Coord!B{r_rec}", f"@Coord!C{r_rec}", ""]]}
    r_px = find("Val", "現價（2026")
    bars = {"type": "bars", "title": "兩讀數的加權目標價與現價（美元；豎線＝現價）", "fmt": "n2",
            "items": [{"label": f"@Coord!B{r_cur}", "value": f"@Coord!B{r_tp}", "series": 0}, {"label": f"@Coord!B{r_alt}", "value": f"@Coord!C{r_tp}", "series": 2}],
            "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
    full = {"type": "table", "caption": "展開：兩讀數的三情境 EPS 與情境 B 顆數、營收（Coord 頁 B 節）", "collapse": True, "summary": "展開：兩讀數的三情境 EPS 與情境 B 顆數、營收",
            "head": ["項目", f"@Coord!B{r_cur}", f"@Coord!C{r_cur - 1}" if False else f"@Coord!B{r_alt}", "另一側相對現行"],
            "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|n2", f"@Coord!C{r}|n2", f"@Coord!D{r}|sp1"] for r in range(cr("NVIDIA 認列顆數（情境 B）FY2028"), cr("12 個月後前瞻 EPS"))]}
    S["valuation"]["blocks"][0:0] = [{"type": "h3", "text": "對帳中：兩讀數目標價"}, {"type": "note", "parts": [f"@Coord!B{r_note}"]}, tbl, bars, full]
    # 敏感度章：總量缺口全由 NVIDIA 吸收（Coord 頁 E 節；兩讀數各一）
    e0 = cr("E. 敏感度：總量缺口全由 NVIDIA 吸收"); eh = e0 + 1; r_feas = cr("全吸收是否可行")
    fe = lambda r: "sp1" if "全吸收相對" in str(CW.cell(r, 1).value) else ("p1" if ("上檔空間" in str(CW.cell(r, 1).value) or "缺口 ÷" in str(CW.cell(r, 1).value)) else "n2")
    S["sens"]["blocks"].insert(1, {"type": "table", "caption": f"@Coord!A{e0}", "head": [f"@Coord!{c}{eh}" for c in "ABCDE"],
                                   "rows": [[f"@Coord!A{r}"] + [f"@Coord!{c}{r}|{fe(r)}" for c in "BCDE"] for r in range(eh + 1, r_feas + 1)]})
    S["sens"]["blocks"].insert(2, {"type": "note", "parts": [f"@Coord!A{r_feas + 1}"]})
    # 標準損益表：預測年度欄標「對帳中」（FY2027E–FY2031E＝PL F–J 欄）
    n = 0
    for b in S["pl"]["blocks"]:
        if b.get("type") == "table" and isinstance(b.get("head"), list) and len(b["head"]) == 9 and str(b["head"][4]).startswith("@PL!F"):
            b["head"][4:9] = [f"@Coord!{c}{r_pl}" for c in "BCDEF"]; n += 1
    assert n == 1, n
    # v0.3.9：損益章在標準損益表後加 FY2029 下滑說明（Coord 頁 C 節；Andy 11-14 17:47 選項 A）
    r_pl29 = cr("損益章說明：FY2029 營收下滑")
    pb = S["pl"]["blocks"]
    i_is = next(i for i, b in enumerate(pb) if b.get("type") == "table" and str(b.get("caption", "")).startswith("標準損益表（十億美元"))
    pb.insert(i_is + 1, {"type": "note", "parts": [f"@Coord!B{r_pl29}"]})


_coord()


# ================= v0.3.9（2026-11-14）：對帳章加「三層協調：讀取層級與 C0 由下往上對帳」（Coord 頁 F 節；數字與文字全取自 xlsx） =================
def _coord_c0():
    if "Coord" not in wb.sheetnames:
        return
    CW = wb["Coord"]
    cr = lambda p, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
    f0, f1h = cr("F. 三層協調"), cr("F1. 本模型讀取")
    f2h = cr("F2. C0 由下往上對帳"); r_tol = cr("對帳門檻"); rh = r_tol + 2
    r_sum = cr("對帳摘要（HTML"); r_read = cr("讀法：F2 各列"); r_n = cr("對帳中項目數"); r_chk = cr("F1 核對")
    S = {x["id"]: x for x in cfg["sections"]}
    # F2 主表：項目、年度、單位、本模型值、層級值、差距、狀態（數值格式依列：顆數 n2、營收 n1、單價 n1；無可比列為文字）
    rows = []
    for r in range(rh + 1, r_n):
        u = str(CW.cell(r, 3).value or "")
        f = "n2" if "顆" in u else ("n3" if u.startswith("x") else "n1")
        rows.append([f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|{f}", f"@Coord!E{r}|{f}", f"@Coord!F{r}|sp1", f"@Coord!G{r}"])
    t2 = {"type": "table", "caption": f"@Coord!A{f2h}", "head": ["項目", "年度", "單位", "本模型", "Siliconomics", "差距", "狀態"], "rows": rows}
    d2 = {"type": "table", "caption": "展開：由下往上對帳每列的口徑、來源與證據等級（兩邊原值見上表）", "collapse": True, "summary": "展開：由下往上對帳的口徑、來源與證據等級",
          "head": ["項目", "本模型口徑與來源", "本模型證據等級", "Siliconomics 口徑與來源", "Siliconomics 證據等級"],
          "rows": [[f"@Coord!A{r}", f"@Coord!H{r}", f"@Coord!I{r}", f"@Coord!J{r}", f"@Coord!K{r}"] for r in range(rh + 1, r_n)]}
    # F1：讀取層級（名稱、年度、層級值、採用值、單位、層級與版本）
    F1a = f1h + 2; F1b = next(r for r in range(F1a, r_chk) if r == r_chk - 1)
    t1 = {"type": "table", "caption": f"@Coord!A{f1h}", "collapse": True, "summary": "展開：本模型讀取的 Siliconomics 名稱、值與版本（含 Tokenomics 鏡像）",
          "head": ["名稱", "年度", "層級值", "本模型採用值", "單位", "層級與版本"],
          "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}|n3", f"@Coord!D{r}|n3", f"@Coord!F{r}", f"@Coord!G{r}"] for r in range(F1a, F1b + 1)]}
    # v0.3.10：C0（實績）表（Coord 頁 F3）放在最前；其後為 F2 既有列（讀取層級、參考；不判定）
    f3h, f3a = cr("F3. C0（實績）"), cr("資料中心營收 FY2027 Q1（公司實績）")
    r_d1, r_d2, r_j = cr("主讀數：已公布期間"), cr("第二讀數：最近一季"), cr("C0（實績）判定"); r_rel, r_org = cr("F3 與 F2 既有"), cr("F3 同源檢查"); r_k, r_s1 = cr("口徑換算"), cr("n2 敏感度")
    t3 = {"type": "table", "caption": f"@Coord!A{f3h}", "head": ["讀數", "期間／口徑", "公司實績（US$B）", "層級隱含（US$B）", "差距", "狀態／判定"],
          "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in (r_d1, r_d2)] + [[f"@Coord!A{r_j}", f"@Coord!B{r_j}", "", "", "", f"@Coord!G{r_j}"]]}
    d3 = {"type": "table", "caption": "展開：實績對帳的輸入、口徑、出處與證據等級", "collapse": True, "summary": "展開：實績對帳的輸入、出處與證據等級",
          "head": ["項目", "期間／口徑", "數值", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級"],
          "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!D{r}|n1", f"@Coord!H{r}", f"@Coord!I{r}", f"@Coord!J{r}", f"@Coord!K{r}"] for r in range(f3a, r_d1 + 2)]}
    blocks = [{"type": "h3", "text": "實績對帳：公司已公布實績對層級隱含值"},
              {"type": "note", "parts": [f"@Coord!D{r_sum}"]}, t3, d3,
              {"type": "note", "parts": [f"@Coord!B{r_k}"]},
              {"type": "table", "caption": "實績對帳敏感度：只換層級每 IT GW 內容，出貨 GW 與實績不變", "head": ["內容取值", "每 IT GW 內容（US$B／GW）", "層級隱含（US$B）", "主讀數差距", "第二讀數差距", "判定"],
               "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|n3", f"@Coord!C{r}|n1", f"@Coord!D{r}|sp1", f"@Coord!E{r}|sp1", f"@Coord!F{r}"] for r in range(r_s1 + 2, r_s1 + 5)]},
              {"type": "note", "parts": [f"@Coord!B{r_org}"]}, {"type": "note", "parts": [f"@Coord!B{r_rel}"]},
              {"type": "h3", "text": "三層協調：讀取層級與由下往上對帳（參考列；不判定）"}, t2, d2,
              {"type": "note", "parts": [f"@Coord!A{r_read}"]}, t1]
    S["recon"]["blocks"].extend(blocks)


_coord_c0()


# ================= v0.3.8 試作 r5（Andy 2026-11-14 12:47；版面規格 v1.3 G、H 節「位置」）：主圖置頂、說明放圖下 =================
# 估值章：第一個區塊＝四個目標價與現價長條圖；章節說明（原 intro）、四個目標價表與說明放在圖下；其後才是對帳中兩讀數目標價區與其餘內容。
# 敏感度章：前兩個區塊＝加權目標價相對基準的低案、高案主圖（變動達兩位數者）；章節說明、未達兩位數的摺疊圖、全吸收表與其他內容放在下方。
# 只改區塊順序；不改任何區塊內容與數字。
def _top_first():
    S = {x["id"]: x for x in cfg["sections"]}
    # 估值章
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b.get("type") == "bars" and str(b.get("title", "")).startswith("四個目標價"))
    i_tbl, i_nt = i_bar - 1, i_bar + 1
    assert va[i_tbl].get("type") == "table" and str(va[i_tbl].get("caption", "")).startswith("@Val!A"), va[i_tbl]
    assert va[i_nt].get("type") == "note" and "投資建議依加權目標價判定" in str(va[i_nt].get("parts")), va[i_nt]
    top = [va[i_bar]]
    if S["valuation"].get("intro"):
        top.append({"type": "note", "parts": [S["valuation"].pop("intro")]})
    top += [va[i_tbl], va[i_nt]]
    S["valuation"]["blocks"] = top + [b for i, b in enumerate(va) if i not in (i_tbl, i_bar, i_nt)]
    # 敏感度章
    sb = S["sens"]["blocks"]
    main = [i for i, b in enumerate(sb) if b.get("type") == "tornado" and not b.get("collapse") and str(b.get("title", "")).startswith("加權目標價相對基準")]
    assert len(main) == 2, main
    top = [sb[i] for i in main]
    if S["sens"].get("intro"):
        top.append({"type": "note", "parts": [S["sens"].pop("intro")]})
    S["sens"]["blocks"] = top + [b for i, b in enumerate(sb) if i not in main]


_top_first()

json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

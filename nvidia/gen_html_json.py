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

json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

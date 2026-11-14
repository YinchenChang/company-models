# v0.4.1：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）
# 讀 avgo/html_src.json（v0.4 的 html.json，原樣保留）→ 本檔前處理 → tools/company_html/relayout.py → 本檔後處理 → 寫 avgo/html.json。
# 用法：python3 avgo/gen_html_json.py（讀 avgo/dist 內唯一的 xlsx）；之後 python3 tools/company_html/build_company_html.py avgo
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

# 情境驅動表：依 In 頁「情境驅動對照」節的實際位置重建（html_src.json 寫死的 In 列號在 M2 節〔標準版 DCF〕插入後錯位；不寫死列號）
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

# ---------- 前處理（只改展示設定，不動 html_src.json） ----------
# 驅動因子：輸入 → 推導量的對照表移到附錄（轉換器依 caption「驅動 → 推導」辨識）
dr = SEC["drivers"]
drivers_intro = dr.pop("intro", None)
for b in dr["blocks"]:
    if b.get("caption", "").startswith("單一參數"):
        b["caption"] = "驅動 → 推導對照（In 頁基準值與推導量）"

# 估值：相對倍數 kv 改成有標題的表；移除舊「重點觀察」（已移到第一章）；長條圖加入相對倍數調整後的 Comparable；
# v0.5：DCF 折現率表（WACC 等）與 DCF 彙總、各情境 FCFF 明細已寫在 html_src.json（同 NVIDIA v0.3）；上一版估值內容摺疊
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
        b["caption"] = "對帳要點：本模型與公司展望、共識、市場層"

skip = [re.split(r"[.\s]", str(IN[f"A{r}"].value))[0] for r in range(4, IN.max_row + 1)
        if str(IN[f"A{r}"].value or "").startswith("情境驅動對照")] + ["H"]

SPEC = {
 "in_skip": skip,
 "in_groups": [
  ["客製加速器 XPU（Google TPU 與其他 ASIC）", [
    ["ASIC 平台顆數與每顆 HBM 容量（市場層）", ["u_tpu_26", "u_tpu_27", "u_tpu_28", "g29_tpu", "u_oth_26", "u_oth_27", "u_oth_28", "g29_oth",
                                       "gb_tpu_26", "gb_tpu_27", "gb_tpu_28", "gb_oth_26", "gb_oth_27", "gb_oth_28"]],
    ["情境顆數係數與顆數敏感度", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C", "uk"]],
    ["ASIC 每顆規格（Broadcom 份額、核心售價、HBM 代購）", ["sh_tpu_26", "sh_tpu_27", "sh_tpu_28", "sh_oth_26", "sh_oth_27", "sh_oth_28",
                                                 "core_tpu", "core_oth", "pfc", "f_hbm", "m_hbm"]],
    ["HBM 價格路徑（讀 SK 海力士模型）", ["hbm_A_26", "hbm_A_27", "hbm_A_28", "hbm_A_29", "hbm_B_26", "hbm_B_27", "hbm_B_28", "hbm_B_29",
                                "hbm_C_26", "hbm_C_27", "hbm_C_28", "hbm_C_29"]],
    ["其他平台顆數與各平台每顆價格（金額份額計算用）", ["u_bw_26", "u_bw_27", "u_bw_28", "g29_bw", "u_rb_26", "u_rb_27", "u_rb_28", "g29_rb",
                                         "u_ru_26", "u_ru_27", "u_ru_28", "g29_ru", "u_hp_26", "u_hp_27", "u_hp_28", "g29_hp",
                                         "u_amd_26", "u_amd_27", "u_amd_28", "g29_amd", "u_trn_26", "u_trn_27", "u_trn_28", "g29_trn",
                                         "pm_bw", "pm_rb", "pm_ru", "pm_hp", "pa_amd", "pa_trn", "pa_tpu", "pa_oth"]],
    ["客戶融資曝險（只作資訊，不進 EPS）", ["fin_anth", "fin_oai", "fin_lease"]],
    ["Wattonomics 錨定係數（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  ["AI 網通（交換器、網卡等）", [
    ["網通營收（讀 markets/network）與毛利率", ["net_26", "net_27", "net_28", "net_29", "rep_avgo26", "gm_net"]]]],
  ["非 AI 半導體與基礎架構軟體（VMware 等）", [
    ["成長率與毛利率", ["g_nai", "gm_nai", "g_sw27", "g_sw", "gm_sw"]]]],
  ["全公司：錨點、費用、稅與股數、股價與目標價規則", [
    ["FY2026 錨點（第一至第三季實績＋第四季指引）", ["q4_rev", "q4_ai", "q4_sw", "q4_gm", "q4_oim", "q4_int", "q4_oth", "q4_tx", "q4_sh"]],
    ["營業費用、利息、稅與股數", ["phi", "g_opx_min", "opx_adj", "int_f", "oth_f", "tx_f", "sh_g"]],
    ["股價、市場倍數與循環位置", ["px", "pe_ttm_gaap", "pe_fwd_sa", "eps_q4_25", "pe_fair_andy", "pk_A", "pk_B", "pk_C"]],
    ["目標價規則（情境權重、DCF 權重、延伸期收斂與共識）", ["tp_wA", "tp_wB", "tp_wC", "tp_g", "tp_gd", "tp_w_dcf", "nm10_k", "cons_tp"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta_u", "beta_de", "beta_du", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（股票薪酬、折舊、資本支出、營運資金）", ["sbc_rev", "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["自由現金流", ["da_rev", "cx_rev"]]]]],
 "blocks_to": {},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "Comparable 彙總": ["同業前瞻 P/E 中位數", "各情境 12 個月後的前瞻 EPS", "12 個月後的前瞻 EPS（情境加權）", "Comparable 目標價"],
          "對帳要點": ["FY2027 AI 營收 vs 公司展望", "FY2028 AI 營收 vs 公司展望", "FY2028：所需 XPU 顆數倍數", "FY2027 EPS vs 共識", "兩家合計 ÷ FY2027 AI 營收"]},
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

# ================= v0.5.6（2026-11-14）：個股 HTML 版面規格 v1.3（Andy 2026-11-14 08:11、09:03、09:16、12:47；樣板 nvidia/gen_html_json.py v0.3.7／v0.3.8） =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳 → 資料狀態。
# 數字與文字取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末；avgo/v056_post.py 產生）。列號依標籤找。
def _v13():
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RV, RCW, CW = wb["README"], wb["PL"], wb["RevDCF"], wb["Recon"], wb["Coord"]
    r13 = find("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "coord", "ver"], range(r13 + 1, r13 + 11)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.5.6"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    old_note = list(cfg["hero"]["note"])
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"]
    # ---- 1. 命題與核心問題：置頂圖（AI 營收：公司展望 vs 模型三情境）、投資建議條件式、重點觀察 5 點 ----
    th = S["thesis"]["blocks"]
    rr = lambda p: find("Recon", p)
    r27, r28, rj, re27 = rr("FY2027 AI 營收 vs 公司展望"), rr("FY2028 AI 營收 vs 公司展望"), rr("FY2027 AI 營收 vs JPMorgan"), rr("FY2027 EPS vs 共識")
    cmp_ = {"type": "compare", "title": "回答命題：模型的 AI 營收與 EPS 對公司展望與市場共識", "groups": [
        {"name": ["FY2027 AI 營收（US$B）"], "fmt": "n0", "items": [{"name": ["公司展望"], "v": f"@Recon!B{r27}", "ref": True}, {"name": ["JPMorgan"], "v": f"@Recon!B{rj}", "ref": True},
                                                            {"name": ["模型 情境 A"], "v": f"@Recon!C{r27}"}, {"name": ["模型 情境 B"], "v": f"@Recon!E{r27}"}, {"name": ["模型 情境 C"], "v": f"@Recon!G{r27}"}]},
        {"name": ["FY2028 AI 營收（US$B）"], "fmt": "n0", "items": [{"name": ["公司展望"], "v": f"@Recon!B{r28}", "ref": True},
                                                            {"name": ["模型 情境 A"], "v": f"@Recon!C{r28}"}, {"name": ["模型 情境 B"], "v": f"@Recon!E{r28}"}, {"name": ["模型 情境 C"], "v": f"@Recon!G{r28}"}]},
        {"name": ["FY2027 EPS（US$）"], "fmt": "n2", "items": [{"name": ["市場共識"], "v": f"@Recon!B{re27}", "ref": True},
                                                          {"name": ["模型 情境 A"], "v": f"@Recon!C{re27}"}, {"name": ["模型 情境 B"], "v": f"@Recon!E{re27}"}, {"name": ["模型 情境 C"], "v": f"@Recon!G{re27}"}]}],
        "caption": "灰色＝公司展望或市場共識；青綠＝本模型三情境。FY2027 為錨點年、三情境相同；FY2028 起情境分歧，基準情境的 AI 營收遠低於公司展望。對照只作事後檢查，校準時未使用。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th.insert(1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
    th.insert(0, cmp_)
    for b in th:
        if b["type"] == "h3" and b.get("text") == "情境定義與會計年度":
            b["text"] = "情境定義"
        if b["type"] == "kv":
            b["items"] = [it for it in b["items"] if it[0] != "會計年度"]
    # ---- 2. 估值：四個目標價置頂；原橫幅的評等規則與相對倍數調整說明放圖下 ----
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b["type"] == "bars")
    g0 = find("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = find("Val", "現價（2026")
    va[i_bar] = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
                 "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                           ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
                 "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
    va.insert(i_bar, {"type": "table", "caption": "四個目標價並列（美元；相對現價）",
                      "head": ["目標價", "美元", "相對現價", "說明"],
                      "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr]
                              + [[f"@Val!A{g0 + 7}", f"@Val!B{g0 + 7}|n2", "", f"@Val!D{g0 + 7}"]]})
    va.insert(i_bar + 2, {"type": "note", "parts": old_note[:7]})
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").startswith(p))
    ra = rv("A. DCF"); rb_ = rv("(a)"); rbb = rv("(b)"); rc = rv("(c)"); rcmp = rv("B. Comparable"); rsum = rv("C. 摘要"); yh = rv("項目（情境 B）")
    m1, m0 = rv("現價隱含的 EBIT 率", rb_), rv("模型：基準的等效 EBIT 率", rb_)
    s1, rv1, rv0 = rv("現價隱含的營收倍數", rbb), rv("現價隱含的 FY2028 營收", rbb), rv("模型：FY2028 營收", rbb)
    g1, g0_ = rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
    w1, w0 = rv("現價隱含的 WACC", rc + 1), rv("模型 WACC", rc)
    rev_ = {
        "id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
        "blocks": [
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
            {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
             "rows": [["預測期 EBIT 率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                      ["營收規模（FY2028 起各年同乘的倍數）", f"@RevDCF!B{s1}|n2", "—"],
                      ["FY2028 營收（十億美元）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                      ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0_}|p1"],
                      ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
            {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"],
             "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if r in (ra + 2, ra + 3, ra + 9) else "n3" if r == ra + 5 else "x1" if r == ra + 8 else "n1")] for r in range(ra + 2, ra + 10)]},
            {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
             "rows": [[f"@RevDCF!A{rcmp + 2}", "—", f"@RevDCF!C{rcmp + 2}|n2", "—"]]
                     + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!C{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!D{r}|sp1"] for k, r in enumerate(range(rcmp + 3, rcmp + 6))]
                     + [[f"@RevDCF!A{rcmp + 6}", "—", f"@RevDCF!C{rcmp + 6}|n1", "—"]]},
            {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
             "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）",
             "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|" + ("p1" if r in (yh + 3, yh + 4) else "n3" if r in (yh + 5, yh + 6, yh + 7, yh + 8) else "n1") for c in "BCDEFG"]
                      for r in range(yh + 1, rb_ - 1)]
                     + [{"sub": f"@RevDCF!A{rb_}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "EBIT 率" in str(RV.cell(r, 1).value) else "n1"), "", "", "", "", ""] for r in range(rb_ + 1, rb_ + 5)]
                     + [{"sub": f"@RevDCF!A{rbb}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "年增率" in str(RV.cell(r, 1).value) else "n2" if "倍數" in str(RV.cell(r, 1).value) else "n3" if "核對" in str(RV.cell(r, 1).value) else "n1"), "", "", "", "", ""] for r in range(rbb + 1, rc - 1)]
                     + [{"sub": f"@RevDCF!A{rc}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p2" if str(RV.cell(r, 1).value).startswith(("現價隱含的 WACC", "模型 WACC")) else "n2"), "", "", "", "", ""] for r in range(w1, w1 + 4)]},
        ]}
    # ---- 5. 損益（標準損益表）----
    p0 = find("PL", "標準損益表（v0.5.6")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
    COLS_ = "CDEFGHIJKL"
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（研發＋銷管）", "營業利益", "營業利益率",
            "其他收入（淨額）＝其他收益 − 利息費用", "稅前利益", "所得稅", "稅後淨利", "淨利率", "稀釋股數", "EPS（完全稀釋）"]
    fmt = lambda u: {"%": "p1", "US$": "n2", "十億股": "n2"}.get(u, "n1")
    rowx = lambda r, cols=COLS_: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in cols]
    r_ratio = plr("比率"); r_gaap = next(r for r in range(ph, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("GAAP 對照"))
    r_chk = plr("核對（應為 0）"); r_c2b = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("差額（FY2030E"))
    is_tbl = {"type": "table", "caption": "標準損益表（Non-GAAP；十億美元，EPS 為美元）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B（基準）",
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in COLS_],
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_gaap)]
                      + [{"sub": "核對：FY2030E–FY2032E 兩種推法的營業利益差額（毛利率與營業費用率沿用 FY2029E）"}] + [rowx(r) for r in (r_c2b, r_c2b + 1)]}
    gaap_tbl = {"type": "table", "caption": f"@PL!A{r_gaap}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in "CDE"], "collapse": True,
                "summary": "展開：GAAP 對照（含股票薪酬與無形資產攤銷；只列實績）與核對列",
                "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in "CDE"] for r in range(r_gaap + 1, r_chk)]
                        + [{"sub": f"@PL!A{r_chk}"}] + [[f"@PL!A{r}"] + [f"@PL!{c}{r}|n3" for c in "CDE"] for r in range(r_chk + 1, r_chk + 8) if r not in (r_c2b, r_c2b + 1)]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}",
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:L{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:L{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:L{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:L{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，Non-GAAP 模型口徑）。FY2026E 為錨點；FY2030E–FY2032E 為延伸期：營收與淨利率由收斂法推得，毛利率與營業費用率沿用 FY2029E。"}
    x0 = find("PL", "三情境延伸到 FY2032"); xh = next(r for r in range(x0, x0 + 4) if PLW.cell(r, 1).value == "情境／延伸項目")
    xs_ = {k: next(r for r in range(x0, PLW.max_row + 1) if PLW.cell(r, 1).value == f"情境 {k}") for k in "ABC"}
    xr = lambda base, lab_: next(r for r in range(base, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
    sc_chart = {"type": "theme", "title": "三情境並列：營收（長條，十億美元）與 EPS（折線，美元，右軸）", "x": f"@PL!C{xh}:J{xh}",
                "fmt": "n0", "fmt_r": "n2", "stack": False, "pad": 1.1,
                "bars": [{"name": [f"@PL!A{xs_[k]}", " 營收"], "y": f"@PL!C{xr(xs_[k], '營收')}:J{xr(xs_[k], '營收')}", "color": c} for k, c in (("A", "#b7e3e4"), ("B", "#10a0a7"), ("C", "#8a8c8c"))],
                "lines": [{"name": [f"@PL!A{xs_[k]}", " EPS"], "y": f"@PL!C{xr(xs_[k], 'EPS（延伸表）')}:J{xr(xs_[k], 'EPS（延伸表）')}", "axis": "r", "color": c, **kw}
                          for k, c, kw in (("A", "#0b7a80", {"dash": True}), ("B", "#3b3838", {"labels": True}), ("C", "#b4472f", {"dash": True}))],
                "caption": "情境 A（樂觀）、B（基準）、C（保守）在 FY2027 相同（錨點），差異自 FY2028 起出現；FY2030 起為延伸期（收斂法）。"}
    old_pl = S["pl"]["blocks"]
    pls = [b for b in old_pl if b["type"] == "pl"]
    notes = [p_ for b in old_pl if b["type"] == "note" for p_ in b["parts"]]
    assert len(pls) == 2 and len(notes) == 2, (len(pls), notes)
    old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境摘要表與折線（FY2025–FY2029；營收、毛利率、營業利益、EPS）"),
                  {"type": "list", "items": notes, "collapse": True, "summary": "展開：延伸期（FY2030–FY2032）的收斂法說明"},
                  dict(pls[1], collapse=True, summary="展開：三情境延伸到 FY2032（收斂法）")]
    S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, sc_chart, gaap_tbl] + old_blocks)
    S["pl"].pop("intro", None)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    rbB = find("PL", "情境 B："); hdrB = rbB + 1
    pr = lambda lab_, start=rbB: next(r for r in range(start, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_ or str(PLW.cell(r, 1).value or "").startswith(lab_))
    r_xpu, r_net, r_nai, r_sw, r_ai = pr("XPU 營收", rbB + 15), pr("AI 網通營收"), pr("非 AI 半導體"), pr("基礎架構軟體"), pr("AI 半導體營收")
    r_bu, r_rep, r_sg = pr("  由下而上營收"), pr("營收", pr("C. 報告值")), pr("檢查：分部合計")
    r_u = pr("Broadcom XPU 顆數（未乘")
    plain = lambda t: t.replace("k_x", "XPU 校準係數").replace("k_n", "網通校準係數").replace("m_core", "XPU 核心毛利率")
    def seg_row(r):
        a = str(PLW.cell(r, 1).value or "")
        vals = [PLW.cell(r, c).value for c in range(4, 8)]
        if all(v is None for v in vals):
            return {"sub": re.sub(r"[（(].*$", "", plain(a))}
        lab_ = f"@PL!A{r}" if not re.search(r"[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+", a) else plain(a)
        f = {"%": "p1", "百萬顆": "n2", "GB": "n0", "US$/GB": "n2", "US$K": "n2", "x": "n2"}.get(PLW.cell(r, 2).value, "n1")
        return [lab_, f"@PL!B{r}"] + [f"@PL!{c}{r}|{f}" for c in "DEFG"]
    seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；十億美元）：XPU＋AI 網通＝AI 半導體；加非 AI 半導體與基礎架構軟體＝由下而上合計",
               "head": [f"@PL!A{hdrB}", f"@PL!B{hdrB}"] + [f"@PL!{c}{hdrB}" for c in "DEFG"],
               "rows": [seg_row(r) for r in range(hdrB + 1, r_sg + 1) if r not in (hdrB + 1,) or True]}
    seg_chart = {"type": "theme", "title": "分部營收（情境 B，十億美元）與 Broadcom XPU 顆數（百萬顆，右軸）", "x": f"@PL!D{hdrB}:G{hdrB}",
                 "fmt": "n0", "fmt_r": "n1", "stack": True, "pad": 1.2, "total": f"@PL!D{r_bu}:G{r_bu}",
                 "bars": [{"name": ["XPU（客製加速器）"], "y": f"@PL!D{r_xpu}:G{r_xpu}", "color": "#0b7a80"},
                          {"name": ["AI 網通"], "y": f"@PL!D{r_net}:G{r_net}", "color": "#6fc7cb"},
                          {"name": ["非 AI 半導體"], "y": f"@PL!D{r_nai}:G{r_nai}", "color": "#c29a2b"},
                          {"name": ["基礎架構軟體"], "y": f"@PL!D{r_sw}:G{r_sw}", "color": "#9a97c9"}],
                 "lines": [{"name": ["Broadcom XPU 顆數（未乘校準係數）"], "y": f"@PL!D{r_u}:G{r_u}", "axis": "r", "color": "#3b3838", "labels": True}],
                 "caption": "堆疊＝XPU、AI 網通、非 AI 半導體、基礎架構軟體四個分部；頂端數字＝由下而上合計；折線＝Broadcom XPU 顆數。FY2026 為錨點年（XPU 與網通校準到公司實績與指引）。"}
    dr_ = S["drivers"]["blocks"]
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + dr_
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    S["drivers"].pop("intro", None)
    # ---- 9. 資料狀態 ----
    st = S["status"]["blocks"]
    st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]},
               {"type": "h3", "text": "三層協調：讀哪一層哪一版、由下往上對帳"}, {"type": "note", "parts": [f"@README!A{RM['coord']}"]},
               {"type": "h3", "text": "會計年度"}, {"type": "note", "parts": [f"@README!A{find('README', '會計年度：')}"]}]
    S["revdcf"] = rev_
    # ---- 7. 對帳章：C0 由下往上對帳（Broadcom 自己的量對它讀的層級值）與層級之間參考 ----
    c0r = next(r for r in range(1, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith("G. C0"))
    c0h = c0r + 1; c0e = c0h
    while CW.cell(c0e + 1, 1).value and not str(CW.cell(c0e + 1, 1).value).startswith("層級之間"): c0e += 1
    lr = next(r for r in range(c0e, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith("層級之間"))
    lh = lr + 1; le = lh
    while CW.cell(le + 1, 1).value: le += 1
    def _fm(r, c):
        nf = CW.cell(r, c).number_format
        return "p1" if "%" in nf else "n3" if nf == "0.000" else "n2" if nf == "#,##0.00" else "n1"
    c0_tbl = {"type": "table", "caption": "由下往上對帳（只算 Broadcom 自己的量對它讀的層級值；差距超過門檻標「對帳中」，不改上游值）",
              "head": ["項目", "本模型值", "層級對照值", "差距", "狀態", "說明"],
              "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|{_fm(r, 2)}", f"@Coord!C{r}|{_fm(r, 3)}", f"@Coord!D{r}|sp1", f"@Coord!E{r}", f"@Coord!I{r}"] for r in range(c0h + 1, c0e + 1)]}
    lay_tbl = {"type": "table", "caption": "層級之間（參考；Siliconomics 對帳單處理，不計入本公司對帳中項目）：產業部署 GW，Siliconomics 基準對另一側讀數",
               "head": ["項目", "Siliconomics 基準", "另一側讀數", "差距", "狀態"],
               "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|n2", f"@Coord!C{r}|n2", f"@Coord!D{r}|sp1", f"@Coord!E{r}"] for r in range(lh + 1, le + 1)]}
    S["recon"]["blocks"][1:1] = [c0_tbl, lay_tbl]
    cfg["sections"] = [S[k] for k in ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]]


_v13()


# 敏感度章：兩張「加權目標價相對基準」龍捲風主圖（低案、高案；變動達兩位數百分比者直接顯示，其餘摺疊）；資料取 Sens 頁末表
def _sens_tor():
    SW = wb["Sens"]; t0 = find("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SW.cell(r2 + 1, 1).value and not str(SW.cell(r2 + 1, 1).value).startswith("快照狀態"): r2 += 1
    out = []
    for col, side in (("B", "低"), ("C", "高")):
        big = [r for r in range(r1, r2 + 1) if isinstance(SW[f"{col}{r}"].value, (int, float)) and abs(SW[f"{col}{r}"].value) >= 0.1]
        small = [r for r in range(r1, r2 + 1) if isinstance(SW[f"{col}{r}"].value, (int, float)) and 0.0005 <= abs(SW[f"{col}{r}"].value) < 0.1]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        out.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
        out.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                        summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例（幾乎為零者見全部案例表）"))
    out.append({"type": "table", "caption": "全部案例的加權目標價相對基準（低案、高案；營運輸入為整本重算快照）", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
                "head": [f"@Sens!{c}{t0 + 1}" for c in "ABCDE"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1", f"@Sens!D{r}|n2", f"@Sens!E{r}|n2"] for r in range(r1, r2 + 1)]})
    out.append({"type": "note", "parts": [f"@Sens!A{r2 + 1}", "：", f"@Sens!B{r2 + 1}", "。", f"@Sens!C{r2 + 1}"], "collapse": True, "summary": "展開：營運輸入案例的快照狀態"})
    return out


_sb = SEC_S = [x for x in cfg["sections"] if x["id"] == "sens"][0]["blocks"]
for _b in _sb:
    if _b.get("type") == "table" and not _b.get("collapse"):
        _b["collapse"] = True; _b["summary"] = "展開：目標價參數表（WACC、β、ERP、永續成長率等）"
    if _b.get("type") == "tornado":
        _b["summary"] = "展開：對照：FY2028 EPS 相對基準的變動（情境 A，一次動一個輸入；未達兩位數百分比）"
    if _b.get("type") == "note" and not _b.get("collapse"):
        _b["collapse"] = True; _b["summary"] = "展開：敏感度頁原文說明"
[x for x in cfg["sections"] if x["id"] == "sens"][0]["blocks"] = _sens_tor() + _sb
[x for x in cfg["sections"] if x["id"] == "sens"][0]["intro"] = ("主圖為加權目標價相對基準的變動（低案、高案分開；變動達兩位數百分比者直接顯示，其餘摺疊）。營運輸入案例為整本重算（三情境同時套用）；"
                                                              "估值框架案例（歷史折價係數、情境權重、DCF 權重）為封閉式；折現參數案例引用目標價參數表。下方摺疊為全部案例與明細。")


# 全頁可見文字不出現參數代號（版面規格 v1.3 B 節）：xlsx 原文不改，展示層含代號的格改用白話
PLAIN = {
    "Val!A10": "永續成長率（長期成長上限與無風險利率取小者）",
    "Val!A11": "永續期新投資報酬率 RONIC（＝WACC＋超額報酬差）",
    "Val!A20": "DCF 情境價值推到目標時點（以股權成本複利、扣除期間股利）",
    "Val!E40": "DCF 與 Comparable 加權（權重見輸入頁目標價規則，可調）",
    "Val!E9": "股權成本與稅後債務成本以市值權重加權；WACC 調整只用於敏感度（基準為零）",
    "In!B157": "永續成長率（最後一年之後；長期成長上限與無風險利率取小者）",
}
FOLD = {"Val!A2": "展開：循環位置判定的說明", "DCF!A2": "展開：DCF 的營業利益與自由現金流口徑"}


def _plain(spec):
    if isinstance(spec, list):
        return [_plain(x) for x in spec]
    if isinstance(spec, str) and spec.startswith("@"):
        ref = spec[1:].split("|")[0]
        if ref in PLAIN:
            return PLAIN[ref]
    return spec


for _sec in cfg["sections"]:
    if _sec["id"] == "status":
        continue
    for _b in _sec["blocks"]:
        if _b.get("type") == "table":
            _b["rows"] = [({"sub": _plain(r["sub"])} if isinstance(r, dict) else _plain(r)) for r in _b.get("rows", [])]
            if "head" in _b: _b["head"] = _plain(_b["head"])
        if _b.get("type") == "note" and not _b.get("collapse") and len(_b["parts"]) == 1 and str(_b["parts"][0])[1:] in FOLD:
            _b["collapse"] = True; _b["summary"] = FOLD[str(_b["parts"][0])[1:]]
# 三層協調 C4（規則 v1.4；tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(cfg, xs[0])


# ================= 三層協調（版面規格 v1.3 I 節）與位置（G、H 節，Andy 2026-11-14 12:47） =================
# 損益表預測年度欄標「對帳中」（Coord 頁 H 節「損益表年度標示」列，FY2026E–FY2032E＝標準損益表 F–L 欄）
def _coord_pl():
    CW = wb["Coord"]
    r_pl = next(r for r in range(1, CW.max_row + 1) if str(CW.cell(r, 1).value or "") == "損益表年度標示（FY2026E–FY2032E）")
    n = 0
    for b in [x for x in cfg["sections"] if x["id"] == "pl"][0]["blocks"]:
        if b.get("type") == "table" and isinstance(b.get("head"), list) and len(b["head"]) == 11 and str(b["head"][4]).startswith("@PL!F"):
            b["head"][4:11] = [f"@Coord!{c}{r_pl}" for c in "BCDEFGH"]; n += 1
    assert n == 1, n
_coord_pl()


# 主圖置頂：估值章第一個區塊＝四個目標價與現價圖（章節說明、四個目標價表與說明放圖下，其後才是對帳中兩讀數區）；
# 敏感度章前兩個區塊＝加權目標價相對基準低案、高案主圖（說明與其他內容放下方）。只改區塊順序。
def _top_first():
    S = {x["id"]: x for x in cfg["sections"]}
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b.get("type") == "bars" and str(b.get("title", "")).startswith("四個目標價"))
    i_tbl, i_nt = i_bar - 1, i_bar + 1
    assert va[i_tbl].get("type") == "table" and str(va[i_tbl].get("caption", "")).startswith("四個目標價並列"), va[i_tbl]
    assert va[i_nt].get("type") == "note" and "投資建議依加權目標價判定" in str(va[i_nt].get("parts")), va[i_nt]
    top = [va[i_bar]]
    if S["valuation"].get("intro"):
        top.append({"type": "note", "parts": [S["valuation"].pop("intro")]})
    top += [va[i_tbl], va[i_nt]]
    S["valuation"]["blocks"] = top + [b for i, b in enumerate(va) if i not in (i_tbl, i_bar, i_nt)]
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

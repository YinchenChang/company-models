# v0.3：DCF 改為標準版（FCFF＋WACC；html_src.json 的 DCF 區塊同 NVIDIA v0.3）並套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）
# 讀 mrvl/html_src.json（v0.2.4 的 html.json，DCF 區塊與列號已改為 v0.3）→ 本檔前處理 → tools/company_html/relayout.py → 本檔後處理 → 寫 amd/html.json。
# 用法：python3 mrvl/gen_html_json.py（讀 mrvl/dist 內唯一的 xlsx）；之後 python3 tools/company_html/build_company_html.py mrvl
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
        b["caption"] = "對帳要點：本模型與公司展望、投資人日、共識"

skip = [re.split(r"[.\s]", str(IN[f"A{r}"].value))[0] for r in range(4, IN.max_row + 1)
        if str(IN[f"A{r}"].value or "").startswith("情境驅動對照")] + ["H"]

SPEC = {
 "in_skip": skip,
 "in_groups": [
  ["客製 XPU（ASIC）", [
    ["加速器顆數與情境係數（市場層）", ["u_trn_26", "u_trn_27", "u_trn_28", "g29_trn", "u_oth_26", "u_oth_27", "u_oth_28", "g29_oth", "u_tpu_26", "u_tpu_27", "u_tpu_28", "g29_tpu", "um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["每顆 HBM 容量與 ASIC 每顆規格（讀 Broadcom 模型）", ["gb_trn_26", "gb_trn_27", "gb_trn_28", "gb_oth_26", "gb_oth_27", "gb_oth_28", "gb_tpu_26", "gb_tpu_27", "gb_tpu_28", "core_tpu", "core_oth", "f_hbm", "m_hbm", "pa_trn", "core_trn",
                                     "av_tpu_26", "av_tpu_27", "av_tpu_28", "av_oth_26", "av_oth_27", "av_oth_28"]],
    ["Marvell 各平台份額與顆數敏感度", ["msh_trn_26", "msh_trn_27", "msh_trn_28", "msh_oth_26", "msh_oth_27", "msh_oth_28", "msh_tpu_26", "msh_tpu_27", "msh_tpu_28", "pfc", "uk"]],
    ["HBM 價格路徑（讀 Broadcom 模型）", ["hbm_A_26", "hbm_A_27", "hbm_A_28", "hbm_A_29", "hbm_B_26", "hbm_B_27", "hbm_B_28", "hbm_B_29", "hbm_C_26", "hbm_C_27", "hbm_C_28", "hbm_C_29"]],
    ["客製營收錨點與 XPU attach", ["cus26", "g_cus27", "att26", "att27", "att28", "g_att", "gm_att"]],
    ["Wattonomics 錨定係數（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  ["光互連（DSP、相干 DCI）", [
    ["AI 後端可插拔光模組（讀 markets/network；CPO 三情境）", ["plug_26", "plug_27", "plug_28", "plug_29", "plugB_26", "plugB_27", "plugB_28", "plugB_29", "plugC_26", "plugC_27", "plugC_28", "plugC_29", "rep_opt26", "asp_26", "asp_27", "asp_28", "asp_29"]],
    ["線性直驅、每模組晶片組金額與份額", ["lpo_26", "lpo_27", "lpo_28", "lpo_29", "dsp_26", "dsp_27", "dsp_28", "dsp_29", "msh_dsp_26", "msh_dsp_27", "msh_dsp_28", "msh_dsp_29", "of"]],
    ["相干 DSP／DCI 與毛利率", ["coh27", "g_coh", "gm_opt"]]]],
  ["交換器、AEC、儲存與通訊", [
    ["交換器、AEC／retimer、儲存", ["sw26", "sw27", "g_sw28", "g_sw", "aec26", "aec27", "g_aec", "sto26", "g_sto", "gm_oth"]],
    ["通訊與其他", ["g_co28", "g_co", "gm_co"]]]],
  ["全公司：錨點、費用、稅與股數、股價與目標價規則", [
    ["FY2027 錨點（前兩季實績、第三季指引、第四季推得）", ["q3_rev", "q3_cog", "q3_gm", "q3_opx", "q3_oth", "q3_tx", "q3_sh", "co_rev27", "co_dcg27", "q4_gm", "co_opx27", "q4_oth", "q4_tx", "q4_sh"]],
    ["營業費用、利息、稅與股數", ["phi", "g_opx_min", "opx_adj", "int_f", "tx_f", "sh_g"]],
    ["股價、循環位置與近一年倍數", ["px", "pe_ttm_gaap", "eps_q3_26", "eps_q4_26", "pk_A", "pk_B", "pk_C"]],
    ["目標價規則（情境權重、DCF 權重、延伸期收斂與共識）", ["tp_wA", "tp_wB", "tp_wC", "tp_g", "tp_gd", "tp_w_dcf", "nm10_k", "cons_tp"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta_u", "beta_de", "beta_du", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（股票薪酬、折舊、資本支出、營運資金）", ["sbc_rev", "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]]]]],
 "blocks_to": {},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "Comparable 彙總": ["同業前瞻 P/E 中位數", "各情境 12 個月後的前瞻 EPS", "12 個月後的前瞻 EPS（情境加權）", "Comparable 目標價"],
          "對帳要點": ["FY2028 營收 vs 投資人日", "FY2028 資料中心 vs 投資人日", "FY2029 客製 vs 投資人日", "FY2028 EPS vs 共識", "FY2028：所需資料中心倍數", "FY2030 毛利率 vs 公司 FY2031 目標"]},
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
        _y = "FY2029" if b.get("value_col") == "L" else "FY2028"
        b["title"] = (f"{_y} EPS 相對基準的變動（情境 A，一次動一個輸入；未達兩位數百分比）" if b.get("collapse")
                      else f"{_y} EPS 相對基準的變動（情境 A，一次動一個輸入；變動達兩位數百分比的案例）")
        if b.get("collapse"):
            b["summary"] = f"展開：{_y} EPS 變動未達兩位數百分比的案例"
# 附錄：驅動 → 推導對照表的說明
st = SEC["status"]["blocks"]
for i, b in enumerate(st):
    if b["type"] == "h3" and b.get("text", "").startswith("附錄：驅動 → 推導") and drivers_intro:
        st.insert(i + 1, {"type": "note", "parts": [drivers_intro]})
        break
# 第三輪修正（v0.3.3）：橫幅附註加併購協議（第三輪清單 28）
cfg["hero"]["note"] = list(cfg["hero"].get("note") or []) + ["。", f"@README!A{find('README', '併購協議（')}"]
# 第三輪修正：橫幅金額單位（美元）
for _t in cfg["hero"]["tiles"]:
    if isinstance(_t[0], str):
        _t[0] = [_t[0], "（美元）"]


# ================= v0.3.5（2026-11-15）：個股 HTML 版面規格 v1.3＋三層協調（樣板 avgo/gen_html_json.py v0.5.6／v0.5.7、amd v0.4.7） =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含 C0 實績）→ 資料狀態。
# 數字與文字取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末、Coord 頁；mrvl/v035_post.py 產生）。列號依標籤找；設定文字除年份與日期外不含數字。
PCODE = re.compile(r"[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+")


def _v13():
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RV, CW, SW = wb["README"], wb["PL"], wb["RevDCF"], wb["Coord"], wb["Sens"]
    r13 = find("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "coord", "ver"], range(r13 + 1, r13 + 11)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.3.7"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    old_note = list(cfg["hero"]["note"])
    assert old_note[0] == "投資建議依加權目標價判定：", old_note
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"]
    # ---- 1. 命題：置頂對照圖（FY2028 營收、資料中心對投資人日展望，EPS 對共識）、投資建議條件式、重點觀察 5 點 ----
    th = S["thesis"]["blocks"]
    rr = lambda p: find("Recon", p)
    r28, rdc, rep = rr("FY2028 營收 vs 投資人日"), rr("FY2028 資料中心 vs 投資人日"), rr("FY2028 EPS vs 共識")
    grp = lambda nm, r, fmt, ref: {"name": [nm], "fmt": fmt, "items": [{"name": [ref], "v": f"@Recon!B{r}", "ref": True},
                                                                      {"name": ["模型 情境 A"], "v": f"@Recon!C{r}"}, {"name": ["模型 情境 B"], "v": f"@Recon!E{r}"}, {"name": ["模型 情境 C"], "v": f"@Recon!G{r}"}]}
    cmp_ = {"type": "compare", "title": "回答命題：模型的 FY2028 營收、資料中心與 EPS 對公司展望與市場共識",
            "groups": [grp("FY2028 營收（US$B）", r28, "n1", "投資人日展望"), grp("FY2028 資料中心（US$B）", rdc, "n1", "投資人日展望"), grp("FY2028 EPS（US$）", rep, "n2", "市場共識")],
            "caption": "灰色＝公司投資人日展望或市場共識；青綠＝本模型三情境（FY2028 只差在網通 CPO 情境，顆數情境自 FY2029 起分歧）。對照只作事後檢查，校準時未使用。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th.insert(0, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
    th.insert(0, cmp_)
    for b in th:
        if b["type"] == "h3" and b.get("text") == "情境定義與會計年度":
            b["text"] = "情境定義"
        if b["type"] == "kv":
            b["items"] = [it for it in b["items"] if it[0] != "會計年度"]
    # ---- 2. 估值：四個目標價置頂（G 位置）；原橫幅的評等規則與相對倍數說明放圖下 ----
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b["type"] == "bars")
    g0 = find("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = g0 + 7
    four = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
            "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                      ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
            "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
    tbl4 = {"type": "table", "caption": "四個目標價並列（美元；相對現價）", "head": ["目標價", "美元", "相對現價", "說明"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_px}", f"@Val!B{r_px}|n2", "", f"@Val!D{r_px}"]]}
    va[i_bar] = four
    va.insert(i_bar + 1, {"type": "note", "parts": old_note})
    va.insert(i_bar + 1, tbl4)
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").startswith(p))
    ra = rv("A. DCF"); rb_ = rv("(a)"); rbb = rv("(b)"); rc = rv("(c)"); rcmp = rv("B. Comparable"); rsum = rv("C. 摘要"); yh = rv("項目（情境 B）")
    m1, m0 = rv("現價隱含的 EBIT 率", rb_), rv("模型：基準的等效 EBIT 率", rb_)
    s1, rv1, rv0 = rv("現價隱含的營收倍數", rbb), rv("現價隱含的 FY2028 營收", rbb), rv("模型：FY2028 營收", rbb)
    gg1, gg0 = rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
    w1, w0, wok = rv("現價隱含的 WACC", rc + 1), rv("模型 WACC", rc), rv("判讀（隱含 WACC）", rc)
    a_rows = [rv(p, ra) for p in ("現價（美元）", "支持現價所需的評價日每股價值", "淨現金滾到評價日", "評價用稀釋股數", "支持現價所需的企業價值", "模型企業價值", "EV 需放大", "模型：情境 B 評價日每股價值")]
    afmt = ["n2", "n2", "n1", "n3", "n1", "n1", "n2", "n2"]
    rev_ = {
        "id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
        "blocks": [
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
            {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
             "rows": [["預測期 EBIT 率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                      ["營收規模（FY2028 起各年同乘的倍數）", f"@RevDCF!B{s1}|n2", "—"],
                      ["FY2028 營收（十億美元）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                      ["FY2028 營收年增率", f"@RevDCF!B{gg1}|p1", f"@RevDCF!B{gg0}|p1"],
                      ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"],
                      ["隱含 WACC 判讀", f"@RevDCF!B{wok}", ""]]},
            {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(a_rows, afmt)]},
            {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
             "rows": [[f"@RevDCF!A{rcmp + 2}", "—", f"@RevDCF!C{rcmp + 2}|n2", "—"]]
                     + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!C{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!D{r}|sp1"] for k, r in enumerate(range(rcmp + 3, rcmp + 6))]
                     + [[f"@RevDCF!A{rcmp + 6}", "—", f"@RevDCF!C{rcmp + 6}|n1", "—"]]},
            {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
             "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）",
             "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|" + ("p1" if r in (yh + 3, yh + 4) else "n3" if r in (yh + 5, yh + 6, yh + 7, yh + 8) else "n1") for c in "BCDEFG"]
                      for r in range(yh + 1, rb_ - 1)]},
        ]}
    # ---- 5. 損益（標準損益表）----
    p0 = find("PL", "標準損益表（v0.3.5")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
    COLS_ = "CDEFGHIJK"
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（研發＋銷管）", "營業利益", "營業利益率",
            "其他收入（淨額）＝−（利息與其他淨損）", "稅前利益", "所得稅", "稅後淨利", "淨利率", "稀釋股數", "EPS（non-GAAP 稀釋）"]
    fmt = lambda u: {"%": "p1", "US$": "n2", "十億股": "n3"}.get(u, "n1")
    rowx = lambda r, cols=COLS_: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in cols]
    r_ratio = plr("比率"); r_gaap = next(r for r in range(ph, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("GAAP 對照"))
    r_chk = plr("核對（應為 0）"); r_c2b = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("差額（FY2031E"))
    is_tbl = {"type": "table", "caption": "標準損益表（Non-GAAP；十億美元，EPS 為美元，稀釋）：FY2024A–FY2026A 實績、FY2027E–FY2032E 情境 B（基準）",
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in COLS_],
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_gaap)]
                      + [{"sub": "核對：FY2031E–FY2032E 兩種推法的營業利益差額（毛利率與營業費用率沿用 FY2030E）"}] + [rowx(r) for r in (r_c2b, r_c2b + 1)]}
    gaap_tbl = {"type": "table", "caption": f"@PL!A{r_gaap}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in "CDE"], "collapse": True,
                "summary": "展開：GAAP 對照（只列實績）與核對列",
                "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" if PLW[f"{c}{r}"].value is not None else "—" for c in "CDE"] for r in range(r_gaap + 1, r_chk)]
                        + [{"sub": f"@PL!A{r_chk}"}] + [[f"@PL!A{r}"] + [f"@PL!{c}{r}|n3" if PLW[f"{c}{r}"].value is not None else "—" for c in "CDE"] for r in range(r_chk + 1, r_chk + 8) if r not in (r_c2b, r_c2b + 1)]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率、淨利率：FY2024A–FY2032E（情境 B）", "x": f"@PL!C{ph}:K{ph}",
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:K{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:K{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:K{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:K{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，Non-GAAP 模型口徑）。FY2027E 為錨點；FY2031E–FY2032E 為延伸期：營收與淨利率由收斂法推得，毛利率與營業費用率沿用 FY2030E。"}
    x0 = find("PL", "三情境延伸到 FY2032"); xh = next(r for r in range(x0, x0 + 4) if PLW.cell(r, 1).value == "情境／延伸項目")
    xs_ = {k: next(r for r in range(x0, PLW.max_row + 1) if PLW.cell(r, 1).value == f"情境 {k}") for k in "ABC"}
    xr = lambda base, lab_: next(r for r in range(base, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
    sc_chart = {"type": "theme", "title": "三情境並列：營收（長條，十億美元）與 EPS（折線，美元，右軸）", "x": f"@PL!C{xh}:I{xh}",
                "fmt": "n0", "fmt_r": "n2", "stack": False, "pad": 1.1,
                "bars": [{"name": [f"@PL!A{xs_[k]}", " 營收"], "y": f"@PL!C{xr(xs_[k], '營收')}:I{xr(xs_[k], '營收')}", "color": c} for k, c in (("A", "#b7e3e4"), ("B", "#10a0a7"), ("C", "#8a8c8c"))],
                "lines": [{"name": [f"@PL!A{xs_[k]}", " EPS"], "y": f"@PL!C{xr(xs_[k], 'EPS（延伸表）')}:I{xr(xs_[k], 'EPS（延伸表）')}", "axis": "r", "color": c, **kw}
                          for k, c, kw in (("A", "#0b7a80", {"dash": True}), ("B", "#3b3838", {"labels": True}), ("C", "#b4472f", {"dash": True}))],
                "caption": "情境 A（樂觀）、B（基準）、C（保守）在 FY2027 相同（錨點）；FY2028 只差在網通 CPO 情境，顆數情境自 FY2029 起分歧；FY2031 起為延伸期（收斂法）。"}
    old_pl = S["pl"]["blocks"]
    keep = []
    for b in old_pl:
        if b["type"] == "pl":
            keep.append(dict(b, collapse=True, summary="展開：三情境摘要表與折線" if not keep else "展開：三情境延伸到 FY2032（收斂法）"))
        elif b["type"] == "note":
            keep.append(dict(b, collapse=True, summary="展開：延伸期（FY2031–FY2032）的收斂法說明"))
    pl_intro = S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, sc_chart, gaap_tbl]
                   + ([{"type": "note", "parts": [pl_intro], "collapse": True, "summary": "展開：三情境損益摘要的期間說明"}] if pl_intro else []) + keep)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    rbB = find("PL", "情境 B："); hdrB = rbB + 1
    assert PLW.cell(hdrB, 1).value == "項目"
    pr = lambda lab_, start=rbB: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    rcB = next(r for r in range(rbB, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("D. 報告值"))
    r_cus, r_ic, r_sw, r_aec, r_sto, r_co, r_rev = pr("客製合計（XPU＋attach）"), pr("光互連營收"), pr("交換器"), pr("AEC 與 retimer"), pr("儲存與其他資料中心"), pr("通訊與其他"), pr("營收", rcB)
    r_xu = next(r for r in range(rbB, rcB) if str(PLW.cell(r, 1).value or "").startswith("Marvell XPU 顆數（未乘"))
    PLN = {"XPU 校準係數": "XPU 校準係數（FY2027 錨點 ÷ 由下而上未校準；FY2028 起沿用）", "光互連校準係數": "光互連校準係數（（錨點光互連 − 相干）÷ 未校準；FY2028 起沿用）",
           "Marvell XPU 顆數（未乘": "Marvell XPU 顆數（未乘校準係數）", "每顆 Marvell 售價（核心＋代購 HBM）": None}

    def seg_row(r):
        a = str(PLW.cell(r, 1).value or "")
        vals = [PLW.cell(r, c).value for c in range(4, 8)]
        if all(v is None for v in vals):
            return {"sub": re.sub(r"[（(].*$", "", a)}
        lab_ = f"@PL!A{r}"
        if PCODE.search(a):
            lab_ = next(v for k, v in PLN.items() if a.startswith(k) and v)
        f = {"%": "p1", "百萬顆": "n3", "百萬個": "n1", "x": "n3", "GB": "n0", "US$/GB": "n2", "US$K": "n2", "US$": "n0"}.get(PLW.cell(r, 2).value, "n2")
        return [lab_, f"@PL!B{r}"] + [f"@PL!{c}{r}|{f}" for c in "DEFG"]
    seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；十億美元）：客製（XPU＋attach）＋光互連＋交換器＋AEC＋儲存＝資料中心；加通訊與其他＝營收",
               "head": [f"@PL!A{hdrB}", f"@PL!B{hdrB}"] + [f"@PL!{c}{hdrB}" for c in "DEFG"],
               "rows": [seg_row(r) for r in range(hdrB + 1, r_co + 1)]}
    seg_chart = {"type": "theme", "title": "分部營收（情境 B，十億美元）與 Marvell XPU 顆數（百萬顆，右軸）", "x": f"@PL!D{hdrB}:G{hdrB}",
                 "fmt": "n1", "fmt_r": "n2", "stack": True, "pad": 1.2, "total": f"@PL!D{r_rev}:G{r_rev}",
                 "bars": [{"name": ["客製（XPU＋attach）"], "y": f"@PL!D{r_cus}:G{r_cus}", "color": "#0b7a80"},
                          {"name": ["光互連"], "y": f"@PL!D{r_ic}:G{r_ic}", "color": "#6fc7cb"},
                          {"name": ["交換器"], "y": f"@PL!D{r_sw}:G{r_sw}", "color": "#c29a2b"},
                          {"name": ["AEC 與 retimer"], "y": f"@PL!D{r_aec}:G{r_aec}", "color": "#e3c77a"},
                          {"name": ["儲存與其他資料中心"], "y": f"@PL!D{r_sto}:G{r_sto}", "color": "#9a97c9"},
                          {"name": ["通訊與其他"], "y": f"@PL!D{r_co}:G{r_co}", "color": "#8a8c8c"}],
                 "lines": [{"name": ["Marvell XPU 顆數（未乘校準係數）"], "y": f"@PL!D{r_xu}:G{r_xu}", "axis": "r", "color": "#3b3838", "labels": True}],
                 "caption": "堆疊＝客製、光互連、交換器、AEC、儲存與其他資料中心、通訊與其他；頂端數字＝營收；折線＝Marvell XPU 顆數。FY2027 為錨點年（客製與光互連拆分由法說口述與差額推得）。"}
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
    # ---- 7. 對帳章：C0（實績）、讀取層級核對（同源）、層級之間參考 ----
    fc = lambda p: next(r for r in range(1, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
    jr = fc("J. C0（實績）")
    jh = next(r for r in range(jr, CW.max_row + 1) if CW.cell(r, 1).value == "項目" and CW.cell(r, 2).value == "公司實績")
    je = jh
    while CW.cell(je + 1, 1).value: je += 1
    ih = next(r for r in range(jr, jh) if CW.cell(r, 1).value == "輸入")
    ie = ih
    while CW.cell(ie + 1, 1).value: ie += 1
    nf = lambda r, c: "n0" if CW.cell(r, 3).value == "月" else "n1" if CW.cell(r, 3).value == "US$K" else "n3"
    c0a_tbl = {"type": "table", "caption": "實績對帳（對帳判定）：公司實績 vs 層級隱含值（Marvell 資料中心營收 vs 層級客製 XPU 隱含；口徑不同，只列參考）",
               "head": ["項目", "公司實績", "層級隱含", "差距", "判定", "說明"],
               "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|n2" if CW.cell(r, 2).value is not None else "", f"@Coord!C{r}|n2" if CW.cell(r, 3).value is not None else "",
                         f"@Coord!D{r}|sp1" if CW.cell(r, 4).value is not None else "", f"@Coord!E{r}", f"@Coord!I{r}"] for r in range(jh + 1, je + 1)]}
    c0i_tbl = {"type": "table", "caption": "實績對帳的輸入（出處與等級）", "collapse": True, "summary": "展開：實績與層級輸入（期間、出處、等級）",
               "head": [f"@Coord!{c}{ih}" for c in "ABCDEF"],
               "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|{nf(r, 2)}", f"@Coord!C{r}", f"@Coord!D{r}", f"@Coord!E{r}", f"@Coord!F{r}"] for r in range(ih + 1, ie + 1)]}
    gr0 = fc("G. 讀取層級核對"); gh = gr0 + 1; ge = gh
    while CW.cell(ge + 1, 1).value: ge += 1
    g_tbl = {"type": "table", "caption": "讀取層級核對（同源，不判定）：個股讀入的層級值與層級值必然相同，不進實績對帳判定、不計對帳中", "collapse": True, "summary": "展開：讀取層級核對（同源，不判定）",
             "head": ["項目", "本模型值", "層級對照值", "差距", "狀態", "說明"],
             "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|n3", f"@Coord!C{r}|n3", f"@Coord!D{r}|sp1", f"@Coord!E{r}", f"@Coord!I{r}"] for r in range(gh + 1, ge + 1)]}
    lr = fc("H. 層級之間"); lh = lr + 1; le = lh
    while CW.cell(le + 1, 1).value: le += 1
    lay_tbl = {"type": "table", "caption": "層級之間（參考；Siliconomics 對帳單處理，不計入本公司對帳中項目）：產業部署 GW，Siliconomics 基準對另一側讀數",
               "head": ["項目", "Siliconomics 基準", "另一側讀數", "差距", "狀態"],
               "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|n2", f"@Coord!C{r}|n2", f"@Coord!D{r}|sp1", f"@Coord!E{r}"] for r in range(lh + 1, le + 1)]}
    cnt = fc("對帳中的 C0 項目數")
    S["recon"]["blocks"][0:0] = [{"type": "note", "parts": [f"@Coord!A{jr + 1}"]}, c0a_tbl, {"type": "note", "parts": [f"@Coord!A{cnt}", "：", f"@Coord!B{cnt}|n0", "。", f"@Coord!A{cnt + 1}"]},
                                 c0i_tbl, g_tbl, lay_tbl]
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
                "head": [f"@Sens!{c}{t0 + 1}" for c in "ABCDE"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1" if SW[f"B{r}"].value is not None else "—", f"@Sens!C{r}|sp1" if SW[f"C{r}"].value is not None else "—",
                                                                         f"@Sens!D{r}|n2" if SW[f"D{r}"].value is not None else "—", f"@Sens!E{r}|n2" if SW[f"E{r}"].value is not None else "—"] for r in range(r1, r2 + 1)]})
    out.append({"type": "note", "parts": [f"@Sens!A{r2 + 1}", "：", f"@Sens!B{r2 + 1}", "。", f"@Sens!C{r2 + 1}"], "collapse": True, "summary": "展開：營運輸入案例的快照狀態"})
    return out


_SS = [x for x in cfg["sections"] if x["id"] == "sens"][0]
for _b in _SS["blocks"]:
    if _b.get("type") in ("table", "note") and not _b.get("collapse"):
        _b["collapse"] = True
        _b["summary"] = ("展開：" + ("目標價敏感度表（業務主因、折現參數）" if _b["type"] == "table" else "敏感度頁原文說明"))
    if _b.get("type") == "tornado":
        _b["summary"] = "展開：對照：EPS 相對基準的變動（情境 A，一次動一個輸入）"
        _b["collapse"] = True
_SS["blocks"] = _sens_tor() + _SS["blocks"]
_SS["intro"] = ("主圖為加權目標價相對基準的變動（低案、高案分開；變動達兩位數百分比者直接顯示，其餘摺疊）。營運輸入案例為整本重算（三情境同時套用）；"
                "估值框架案例（相對倍數係數、情境權重、DCF 權重）為封閉式；折現參數案例引用目標價參數表。下方摺疊為全部案例與明細。")


# 全頁可見文字不出現參數代號（版面規格 v1.3）：xlsx 原文不改，展示層含代號的格改用白話
def _plain_pass():
    PLAIN = {}
    V = wb["Val"]
    for r in range(1, V.max_row + 1):
        a = V.cell(r, 1).value
        if isinstance(a, str) and a.startswith("永續期 RONIC＝WACC＋"): PLAIN[f"Val!A{r}"] = "永續期新投資報酬率 RONIC（＝WACC＋超額報酬差）"
        if isinstance(a, str) and a.startswith("永續成長率 g＝MIN"): PLAIN[f"Val!A{r}"] = "永續成長率（長期成長上限與無風險利率取小者）"
        if isinstance(a, str) and a.startswith("DCF 情境價值推到 12 個月後"): PLAIN[f"Val!A{r}"] = "DCF 情境價值推到目標時點（以股權成本複利、扣除期間股利）"
    def p(spec):
        if isinstance(spec, list):
            return [p(x) for x in spec]
        if isinstance(spec, str) and spec.startswith("@"):
            ref = spec[1:].split("|")[0]
            if ref in PLAIN:
                return PLAIN[ref]
        return spec
    for sec in cfg["sections"]:
        if sec["id"] == "status":
            continue
        for b in sec["blocks"]:
            if b.get("type") == "table":
                b["rows"] = [({"sub": p(r["sub"])} if isinstance(r, dict) else p(r)) for r in b.get("rows", [])]
_plain_pass()

# 三層協調 C4（規則 v1.4；tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(cfg, X)


# 損益表預測年度欄標「對帳中」（Coord 頁 I 節「損益表年度標示」列，FY2027E–FY2032E＝標準損益表 F–K 欄）
def _coord_pl():
    CW = wb["Coord"]
    r_pl = next(r for r in range(1, CW.max_row + 1) if str(CW.cell(r, 1).value or "") == "損益表年度標示（FY2027E–FY2032E）")
    n = 0
    for b in [x for x in cfg["sections"] if x["id"] == "pl"][0]["blocks"]:
        if b.get("type") == "table" and isinstance(b.get("head"), list) and len(b["head"]) == 10 and str(b["head"][4]).startswith("@PL!F"):
            b["head"][4:10] = [f"@Coord!{c}{r_pl}" for c in "BCDEFG"]; n += 1
    assert n == 1, n
_coord_pl()


# 主圖置頂：估值章第一個區塊＝四個目標價與現價圖（章節說明、四個目標價表與說明放圖下，其後才是對帳中兩讀數區）；
# 敏感度章前兩個區塊＝加權目標價相對基準低案、高案主圖（說明與其他內容放下方）。只改區塊順序。
def _top_first():
    S = {x["id"]: x for x in cfg["sections"]}
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b.get("type") == "bars" and str(b.get("title", "")).startswith("四個目標價"))
    i_tbl, i_nt = i_bar + 1, i_bar + 2
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

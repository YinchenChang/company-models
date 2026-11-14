# Astera Labs HTML 設定：讀 alab/html_src.json（v0.2 的 html.json，內容不動）→ 套用個股 HTML 版面規格 v1.0 → 寫 alab/html.json。
# 用法：python3 alab/gen_html_json.py <成品 xlsx> alab/html_src.json alab/html.json
# 規格：docs/plan/個股HTML版面規格.md；共用轉換器 tools/company_html/relayout.py（原樣使用）。
# 本模型屬伺服器組裝樣板（v0.3 起 DCF 為標準版 FCFF＋WACC、Comparable 為同業前瞻本益比 × 相對倍數調整），與轉換器預設的估值頁不同，
# 所以 relayout 之後在本檔重組：投資建議橫幅（補充規則 A：中間格為市場主流同業法，調整後值寫在附註）、估值章、敏感度章（龍捲風為 EPS，不是目標價）。
import copy
import json
import os
import re
import sys

import openpyxl

X, SRC, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "company_html"))
from relayout import relayout  # noqa: E402

cfg = json.load(open(SRC, encoding="utf-8"))
src = copy.deepcopy(cfg)
wb = openpyxl.load_workbook(X, data_only=True)


def lab(sh, prefix, start=1, col=1):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip().startswith(prefix.strip()):
            return r
    raise KeyError((sh, prefix))


INK = {str(wb["In"].cell(r, 1).value): r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}
sec = lambda c, i: next(s for s in c["sections"] if s["id"] == i)

# 第三輪修正：情境驅動表的數值依單位給格式（% → 百分比、pt → 帶正負號百分點、情境權重 → 整數百分比）
def _r3_scn_row(r):
    k = str(wb["In"].cell(r, 1).value or ""); u = str(wb["In"].cell(r, 3).value or "")
    f = "p0" if k == "scn_w" else {"%": "p2" if "fade_m" in k else "p1", "pt": "sp1"}.get(u, "")
    unit = "百分點" if u == "pt" else ("權重" if k == "scn_w" else f"@In!C{r}")
    return [f"@In!B{r}"] + [f"@In!{c}{r}" + (f"|{f}" if f else "") for c in "DEF"] + [unit]


# 情境驅動表：依 In 頁「情境驅動對照」節的實際列號重寫參照（v0.4 起 In 頁新增折現率一節，html_src 的固定列號錯位；欄序同 html_src：B、D、E、F、C）
_scn_r0 = next(r for r in range(4, wb["In"].max_row + 1) if str(wb["In"].cell(r, 1).value or "").startswith("情境驅動對照")) + 1
_scn_rs, _r = [], _scn_r0 + 1
while wb["In"].cell(_r, 1).value:
    _scn_rs.append(_r); _r += 1
_n = 0
for b in sec(cfg, "drivers")["blocks"]:
    if str(b.get("caption", "")).startswith("情境驅動"):
        b["head"] = [f"@In!{c}{_scn_r0}" for c in "BDEFC"]
        b["rows"] = [_r3_scn_row(r) for r in _scn_rs]
        _n += 1
assert _n == 1, "情境驅動表"

# ---------- relayout 前：驅動 → 推導對照表移到附錄；對帳表給標題 ----------
for b in sec(cfg, "drivers")["blocks"]:
    if str(b.get("caption", "")).startswith("單一參數（"):
        b["caption"] = "驅動 → 推導對照（主要輸入的基準值與推導量）"
rc = sec(cfg, "recon")["blocks"][0]
rc["caption"] = "本模型與共識、指引及份額口徑的對照"

_scn = next(re.split(r"[.\s]", str(wb["In"].cell(r, 1).value))[0] for r in range(4, wb["In"].max_row + 1)
            if str(wb["In"].cell(r, 1).value or "").startswith("情境驅動對照"))
AI = "Scorpio X 與 PCIe 等（scale-up 交換與加速器市場、Astera 份額、客戶集中）"
GS = "（保留）"
CO = "全公司（錨點與實績、費用與稅率、股價、目標價規則、折現率與自由現金流）"
SPEC = {
 "in_skip": ["H", _scn],
 "in_groups": [
  [AI, [
    ["scale-up 交換晶片與加速器市場（讀本專案網通模型）", ["msx_26", "msx_27", "msx_28", "msx_29", "msu_26", "msu_27", "msu_28", "msu_29", "u_26", "u_27", "u_28", "u_29", "nnu_26", "nnu_27", "nnu_28", "nnu_29", "itc_26", "itc_27", "itc_28", "itc_29"]],
    ["情境顆數係數（2028 年起三情境不同）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["Scorpio X（本專案權威）", ["x_25", "x_26", "rel_x_27", "rel_x_28", "rel_x_29"]],
    ["PCIe 與訊號調節、客戶集中", ["c_pcie", "rel_p_27", "rel_p_28", "rel_p_29", "cust_top", "cl_loss", "cl_27", "cl_28", "cl_29"]],
    ["Wattonomics 錨定係數（只用於敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  [CO, [
    ["FY2026 錨點（前三季實績＋第四季公司指引）", ["q1_rev", "q1_gp", "q1_opx", "q1_ninc", "q1_ni", "q1_sh", "q2_rev", "q2_gp", "q2_opx", "q2_ninc", "q2_ni", "q2_sh", "q2_eps", "q3_rev", "q3_gp", "q3_opx", "q3_eps", "q3_sh", "tx_26", "q_oth", "q4_rev", "gm_q4", "q4_opx", "q4_sh"]],
    ["FY2024、FY2025 實績（non-GAAP）與股東權益", ["a24_rev_in", "a24_gp_in", "a24_opx_in", "a24_ninc_in", "a24_ni_in", "a24_sh_in", "a24_eps_in", "a25_rev_in", "a25_gp_in", "a25_opx_in", "a25_ninc_in", "a25_ni_in", "a25_sh_in", "a25_eps_in", "a25_ni_h1", "eq_23", "eq_24", "eq_25", "eq_q2"]],
    ["營業費用、毛利率、稅率、股數（FY2027 起）", ["phi", "opx_adj", "gm_27", "gm_28", "gm_29", "oth_f", "tx_27", "tx_28", "tx_29", "sh_g", "mf", "g2n"]],
    ["股價", ["px", "px_hi", "sh_out"]],
    ["目標價規則與收斂法", ["tp_g", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC", "cons_tp", "cons_feps"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta", "beta_u", "beta_de", "beta_du", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（GAAP 換算、折舊、資本支出、營運資金）", ["gaap_k", "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]]]]],
 "blocks_to": {"Astera Labs 份額": AI, "@Units!A16": AI},
 "main": {"本模型與": ["FY2026 EPS vs 共識", "FY2027 營收 vs 共識", "FY2027 EPS vs 共識", "FY2028 EPS vs 共識", "FY2027 毛利率 vs 公司長期目標"]},
}
cfg = relayout(cfg, X, SPEC)

# ---------- relayout 後（本模型樣板專用） ----------
V = lambda p, start=1: lab("Val", p, start)
r_tpw, r_dcf, r_cmp, r_px = V("目標價（DCF 與 Comparable 加權）"), V("DCF 目標價（情境加權）"), V("Comparable 目標價（相對倍數調整後）"), V("現價（")
r_up, r_upc, r_rec, r_cons = V("上檔空間：加權目標價"), V("上檔空間：Comparable 目標價"), V("投資建議"), V("市場共識目標價")
r_cmp0, r_grade = V("Comparable 目標價（市場主流同業法）"), V("上檔空間級距")
h_upc = lab("README", "上檔空間：Comparable 目標價（相對倍數調整後）")
h_val = lab("README", "估值說明：")

# 投資建議橫幅（補充規則 A）：中間格＝只看 Comparable（市場主流同業法、未經相對倍數調整）＋自己的上檔空間；
# 進加權目標價的是相對倍數調整後的 Comparable，數值寫在橫幅附註（公式帶數字）
cfg["hero"]["tiles"][0][1] = f"@Val!B{r_tpw}|n2"
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{r_cmp0}|n2", ["上檔空間 ", f"@Val!B{r_upc}|sp1"]]
cfg["hero"]["tiles"][2] = [f"@Val!A{r_px}", f"@Val!B{r_px}|n2", ""]
_note = cfg["hero"].get("note") or []
cfg["hero"]["note"] = (_note if isinstance(_note, list) else [_note]) + ["；進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{r_cmp}|n2", " 美元（上檔空間 ", f"@README!B{h_upc}|sp1", "）。"]

# 命題章：原有的情境定義與口徑摺疊在最後
th_kv = next(b for b in sec(src, "thesis")["blocks"] if b["type"] == "kv")
sec(cfg, "thesis")["blocks"].append(dict(th_kv, collapse=True, summary="展開：三情境的定義與會計口徑"))

# 估值章：說明 → 目標價長條圖 → 結論主要項目 → 折現率 → DCF → Comparable；舊法（十年本益比規則值）全部摺疊
vs = sec(src, "valuation")["blocks"]
va = sec(cfg, "valuation")
va["intro"] = f"@README!A{h_val}"


def vrow(r, f="n0"):
    return [f"@Val!A{r}", f"@Val!B{r}|{f}", f"@Val!E{r}"]


def vrow3(r, f="n1"):
    return [f"@Val!A{r}", f"@Val!B{r}|{f}", f"@Val!C{r}|{f}", f"@Val!D{r}|{f}", f"@Val!E{r}"]


ink = lambda k, f: [f"@In!B{INK[k]}", f"@In!E{INK[k]}|{f}"]
r_hist, r_meth, r_band = V("歷史倍數檢查"), V("目標價方法"), V("價格帶（DCF")
conc = [vrow(r_tpw), vrow(r_dcf), vrow(r_cmp), vrow(r_px), vrow(r_up, "sp1"),
        [f"@README!A{h_upc}", f"@README!B{h_upc}|sp1", "上檔空間以相對倍數調整後的 Comparable 目標價計算（進加權目標價的值）"],
        [f"@Val!A{r_rec}", f"@Val!B{r_rec}", ""], vrow(r_cons)]
conc_rest = [vrow(r_cmp0), vrow(r_upc, "sp1"), vrow(V("與市場共識目標價的差距：EPS 因素"), "sp1"), vrow(V("與市場共識目標價的差距：倍數因素"), "sp1"),
             vrow(r_band, ""), [f"@Val!A{r_hist}", [f"@Val!B{r_hist}|p0", "／", f"@Val!C{r_hist}|p0", "／", f"@Val!D{r_hist}|p0"], f"@Val!E{r_hist}"],
             [f"@Val!A{r_meth}", f"@Val!B{r_meth}", ""], [f"@Val!A{r_grade}", f"@Val!B{r_grade}", f"@Val!E{r_grade}"]]
# v0.2：標準版 DCF——折現率（緊接在 DCF 彙總之前；規格 v1.0 直接顯示 WACC、ke、稅後 kd、RONIC、永續成長率）
r_d0 = V("0.1 DCF")
vr = lambda p, f: [f"@Val!A{V(p, r_d0)}", f"@Val!B{V(p, r_d0)}|{f}", f"@Val!E{V(p, r_d0)}"]
disc = {"type": "table", "caption": "DCF 折現率（WACC）、永續成長率與權重", "head": ["項目", "數值", "說明"],
        "rows": [vr("股權成本", "p2"), vr("稅後債務成本", "p2"), vr("權益占資本", "p1"), vr("WACC", "p2"), vr("永續成長率", "p2"), vr("永續期 RONIC", "p2"),
                 [f"@In!B{INK['tp_w_dcf']}", f"@In!E{INK['tp_w_dcf']}|p0", ""],
                 ["情境權重（A／B／C）", [f"@In!E{INK['tp_wA']}|p0", "／", f"@In!E{INK['tp_wB']}|p0", "／", f"@In!E{INK['tp_wC']}|p0"], ""]]}
d3 = lambda p, f: [f"@Val!A{V(p, r_d0)}"] + [f"@Val!{c}{V(p, r_d0)}|{f}" for c in "BCD"] + [f"@Val!E{V(p, r_d0)}"]
dhead = [f"@Val!A{r_d0 + 1}", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "說明"]
dcf_main = [d3("明確期 FCFF 現值", "n1"), d3("終值現值", "n1"), d3("企業價值 EV", "n1"), d3("＋淨現金與投資", "n1"), d3("＋Stub 期 FCFF", "n1"), d3("淨現金滾到評價日", "n1"), d3("DCF 情境價值推到 12 個月後", "n0"),
            [f"@Val!A{r_dcf}", f"@Val!B{r_dcf}|n0", "", "", f"@Val!E{r_dcf}"]]
dcf_rest = [d3("評價用稀釋股數", "n4"), d3("DCF 情境每股價值（評價日）", "n1"), d3("  其中：終值現值占 EV", "p1"), d3("  隱含終值 EV／EBITDA", "x1")]
r_c0, r_med, r_feps, r_c = V("0.2 Comparable"), V("同業前瞻 P/E 中位數"), V("12 個月後的前瞻 EPS"), V("Comparable 目標價", V("12 個月後的前瞻 EPS"))
r_coef = V("相對倍數調整係數")
cmp_src = next(b for b in vs if b.get("caption") == f"@Val!A{r_c0}")
rm_kv = next(b for b in vs if b["type"] == "kv" and b["items"][0][0] == f"@Val!A{V('本公司 10 年中位數')}")
nb = [{"type": "bars", "title": "目標價與現價", "fmt": "n0", "items": [
        {"label": f"@Val!A{r_dcf}", "value": f"@Val!B{r_dcf}", "series": 0},
        {"label": f"@Val!A{r_cmp}", "value": f"@Val!B{r_cmp}", "series": 1},
        {"label": f"@Val!A{r_tpw}", "value": f"@Val!B{r_tpw}", "series": 2},
        {"label": f"@Val!A{r_cons}", "value": f"@Val!B{r_cons}", "series": 0}],
       "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}},
      {"type": "table", "caption": "結論（主要項目）", "head": ["項目", "數值", "說明"], "rows": conc},
      {"type": "table", "collapse": True, "summary": "展開：結論的其他項目（市場主流同業法、與共識的差距拆解、價格帶、歷史倍數檢查、目標價方法、級距）", "head": ["項目", "數值", "說明"], "rows": conc_rest},
      disc,
      {"type": "table", "caption": f"@Val!A{r_d0}", "head": dhead, "rows": dcf_main},
      {"type": "table", "collapse": True, "summary": "展開：DCF 的其他項目（股數、評價日每股價值、終值占比、隱含終值倍數）", "head": dhead, "rows": dcf_rest},
      {"type": "table", "caption": f"@Val!A{r_c0}", "head": ["項目", "數值"],
       "rows": [[f"@Val!A{r_med}", f"@Val!B{r_med}|x1"], [f"@Val!A{r_feps}", f"@Val!B{r_feps}|n1"], [f"@Val!A{r_c}", f"@Val!B{r_c}|n0"],
                [(f"@Val!A{r_coef}" if isinstance(wb["Val"].cell(r_coef - 3, 2).value, (int, float)) else "相對倍數調整係數：未調整（資料不足）"), f"@Val!B{r_coef}|n2"], [f"@Val!A{r_cmp}", f"@Val!B{r_cmp}|n0"]]},
      dict({k: v for k, v in cmp_src.items() if k != "caption"}, collapse=True, summary="展開：同業前瞻本益比"),
      {"type": "table", "collapse": True, "summary": "展開：相對倍數調整的組成（本公司與同業的長期本益比）", "head": ["項目", "數值"],
       "rows": [[k, v] for k, v in rm_kv["items"] if not v.startswith((f"@Val!B{r_coef}", f"@Val!B{r_cmp}"))]},
      {"type": "note", "parts": [f"@Val!E{r_med}"]},
      {"type": "h3", "text": "參考：舊法（十年本益比規則值）與市場隱含倍數（不進目標價）"},
      {"type": "note", "parts": ["@Val!A1"], "collapse": True, "summary": "展開：本頁原說明"}]
# 舊法各塊（既有估值、市場隱含 EPS、循環位置、規則值）：依原順序全部摺疊
i_old = next(i for i, b in enumerate(vs) if b["type"] == "h3" and str(b.get("text", "")).startswith("既有估值"))
cap_ = "舊法"
for b in vs[i_old + 1:]:
    if b["type"] == "h3":
        cap_ = re.sub(r"（.*?）", "", str(b.get("text", ""))).strip() or cap_
        continue
    if b.get("collapse"):
        nb.append(b)
        continue
    if b["type"] == "note":
        nb.append(dict(b, collapse=True, summary=f"展開：{cap_}的說明"))
        continue
    t_ = b.get("caption") or b.get("title") or cap_
    if isinstance(t_, str) and t_.startswith("@"):
        _m = re.match(r"^@([^!]+)!([A-Z]+\d+)", t_)
        t_ = str(wb[_m.group(1)][_m.group(2)].value or cap_) if _m else cap_
    nb.append(dict(b, collapse=True, summary="展開：" + re.sub(r"^[\d.\s]+", "", re.sub(r"（.*?）", "", str(t_))).strip()[:30]))
va["blocks"] = nb

# 敏感度章：龍捲風是 EPS（情境 A），標題改回原標題
se = sec(cfg, "sens")
tor_src = {b["value_col"]: b["title"] for b in sec(src, "sens")["blocks"] if b["type"] == "tornado"}
for b in se["blocks"]:
    if b["type"] == "tornado":
        t = tor_src[b["value_col"]]
        base, _, rest = t.partition("（")
        if b.get("collapse"):
            b["title"] = f"{base}（未達兩位數百分比）"
        else:
            b["title"] = f"{base}（變動達兩位數百分比的案例" + (f"；{rest}" if rest else "）")
# v0.2：目標價參數表（標準版 DCF 案例；表格移到案例區塊之後）
r_tp = lab("Sens", "目標價參數敏感度")
r_end = r_tp + 2
while wb["Sens"].cell(r_end + 1, 1).value:
    r_end += 1
for b in se["blocks"]:
    if b.get("caption") == "@Sens!A23":
        b["caption"] = f"@Sens!A{r_tp}"
        b["head"] = [f"@Sens!{c}{r_tp + 1}" for c in "ABCDEFG"]
        b["rowgen"] = {"sheet": "Sens", "rows": f"{r_tp + 2}-{r_end}", "cols": list("ABCDEFG"), "fmts": ["", "", "n0", "sp1", "n0", "sp1", "n0"]}
# Wattonomics 錨定案例（Sens 頁最後一節；原 HTML 未列）
r_w = lab("Sens", "Wattonomics 錨定（三層校準原則）")
r_wh = lab("Sens", "案例", r_w)
r_we = r_wh
while wb["Sens"].cell(r_we + 1, 1).value:
    r_we += 1
full = next(b for b in sec(src, "sens")["blocks"] if b["type"] == "table" and b.get("collapse"))
se["blocks"] += [{"type": "note", "parts": [f"@Sens!A{r_w + 1}"]},
                 {"type": "table", "collapse": True, "summary": "展開：Wattonomics 錨定案例（情境 B 乘上錨定係數；不重算目標價）",
                  "head": [f"@Sens!{c}{r_wh}" for c in full["rowgen"]["cols"]],
                  "rowgen": dict(full["rowgen"], rows=f"{r_wh + 1}-{r_we}")}]

# 第三輪修正：In 頁單位為 pt（百分點，存為小數）的列，relayout 預設 n1 會把 0.02 顯示成 0.0 → 改 sp1（帶正負號百分比）
_PT = {r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 3).value == "pt"}
def _r3_pt(o):
    if isinstance(o, dict):
        return {k: _r3_pt(v) for k, v in o.items()}
    if isinstance(o, list):
        return [_r3_pt(v) for v in o]
    if isinstance(o, str):
        m = re.fullmatch(r"@In!([DEF])(\d+)\|n1", o)
        if m and int(m.group(2)) in _PT:
            return f"@In!{m.group(1)}{m.group(2)}|sp1"
    return o
cfg = _r3_pt(cfg)

# ================= v0.3.6／v0.3.7（2026-11-09）：Astera Labs v0.3.5 驗收報告採用項（alab/v036_post.py） =================
RDW = wb["README"]
R_ = lambda p, start=1: lab("README", p, start)
def _rows_after(head, stop_prefixes):
    r = R_(head) + 1; out = []
    while RDW.cell(r, 1).value and not str(RDW.cell(r, 1).value).startswith(stop_prefixes):
        out.append(r); r += 1
    return out
r_state, r_edge, r_asof_h = R_("build_id "), R_("評等"), R_("資料基準日（as_of）")
r_asof_t = r_asof_h + 1; asof_rows = _rows_after("項目", ("說明：",)) if RDW.cell(r_asof_t, 1).value == "項目" else []
r_asof_n = R_("說明：rf 與同業")
fail_rows = _rows_after("主要失效條件", ("估值框架並列",)); r_frame = R_("估值框架並列")
risk_rows = _rows_after("風險登錄", ("主要假設驅動",)); ind_rows = _rows_after("主要假設驅動與可證偽", ("\u0000",))
st_add = [{"type": "note", "parts": [f"@README!A{r_state}"]}, {"type": "note", "parts": [f"@README!A{r_edge}"]},
          {"type": "h3", "text": "主要失效條件與估值框架並列"}, {"type": "list", "items": [f"@README!A{r}" for r in fail_rows] + [f"@README!A{r_frame}"]},
          {"type": "h3", "text": "資料基準日與逾期標示"}, {"type": "note", "parts": [f"@README!A{r_asof_h}"]},
          {"type": "table", "head": [f"@README!{c}{r_asof_t}" for c in "ABC"], "rows": [[f"@README!A{r}", f"@README!B{r}", f"@README!C{r}"] for r in asof_rows]},
          {"type": "note", "parts": [f"@README!A{r_asof_n}"]},
          {"type": "h3", "text": "風險登錄（是否建模、理由、觀察指標）"}, {"type": "list", "items": [f"@README!A{r}" for r in risk_rows]},
          {"type": "h3", "text": "主要假設驅動與可證偽的公開指標"}, {"type": "list", "items": [f"@README!A{r}" for r in ind_rows]}]
sec(cfg, "status")["blocks"] = st_add + sec(cfg, "status")["blocks"]
# 橫幅附註：評等邊界（J4）
cfg["hero"]["note"] = cfg["hero"]["note"] + ["", f"@README!A{r_edge}"]
# 估值章：舊法（十年本益比規則值）不顯示（J3 選項 a）；市場隱含 EPS 讀法展開；期間並列與估值框架
va = sec(cfg, "valuation")
def _txt(b):
    t = str(b.get("summary") or b.get("caption") or b.get("title") or b.get("text") or "")
    m = re.match(r"^@([^!]+)!([A-Z]+\d+)", t)
    return str(wb[m.group(1)][m.group(2)].value or "") if m else t
r_read = V("讀法（公式組字")
keep = []
for b in va["blocks"]:
    t = _txt(b)
    if b["type"] == "h3" and t.startswith("參考："):
        keep.append(dict(b, text="參考：市場隱含倍數（不進目標價）")); continue
    if any(k in t for k in ("規則值", "舊法", "十年", "循環位置")) and not (b["type"] == "note" and b.get("parts") == [f"@Val!A{r_read}"]):
        continue
    if b["type"] == "note" and b.get("parts") == [f"@Val!A{r_read}"]:
        b = {k: v for k, v in b.items() if k not in ("collapse", "summary")}
    keep.append(b)
if not any(b.get("parts") == [f"@Val!A{r_read}"] for b in keep):
    keep.append({"type": "note", "parts": [f"@Val!A{r_read}"]})
r_p1, r_p2 = V("折現期間："), V("持有期間：")
i_conc = next(i for i, b in enumerate(keep) if b.get("caption") == "結論（主要項目）")
keep[i_conc + 1:i_conc + 1] = [{"type": "note", "parts": [f"@README!A{r_frame}"]},
                               {"type": "table", "caption": "期間並列", "head": ["項目", "數值", "說明"], "rows": [vrow(r_p1, "n2"), vrow(r_p2, "n2")]}]
va["blocks"] = keep
# 敏感度章：目標價表加 DCF 單法美元欄；業務主因（整本重算快照）；龍捲風標題的基準情境
SEW = wb["Sens"]
se = sec(cfg, "sens")
sb0 = "B" if "情境 B；" in str(SEW["A1"].value) else "A"
for b in se["blocks"]:
    if b["type"] == "tornado" and sb0 == "B":
        b["title"] = b["title"].replace("情境 A", "情境 B")
    if b.get("caption") == f"@Sens!A{r_tp}":
        b["head"] = [f"@Sens!{c}{r_tp + 1}" for c in "ABCDEFGIJ"]
        b["rowgen"] = {"sheet": "Sens", "rows": f"{r_tp + 2}-{r_end}", "cols": list("ABCDEFGIJ"), "fmts": ["", "", "n0", "sp1", "n0", "sp1", "n0", "n2", "n2"]}
try:
    r_bz = lab("Sens", "目標價敏感度：業務主因")
except KeyError:
    r_bz = None
if r_bz:
    r_bh = r_bz + 2; r_be = r_bh
    while SEW.cell(r_be + 1, 1).value: r_be += 1
    i_tp = next(i for i, b in enumerate(se["blocks"]) if b.get("caption") == f"@Sens!A{r_tp}")
    se["blocks"][i_tp + 1:i_tp + 1] = [{"type": "table", "caption": f"@Sens!A{r_bz}", "head": [f"@Sens!{c}{r_bh}" for c in "ABCDEFGIJ"],
                                        "rowgen": {"sheet": "Sens", "rows": f"{r_bh + 1}-{r_be}", "cols": list("ABCDEFGIJ"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2", "n2", "n2"]}},
                                       {"type": "note", "parts": [f"@Sens!A{r_bz + 1}"]}]
# 對帳章：反解（E 節）
RCW = wb["Recon"]
r_e = lab("Recon", "E. 反解")
r_ee = r_e + 1
while RCW.cell(r_ee + 1, 1).value: r_ee += 1
r_eh = next(r for r in range(r_e + 1, r_ee + 1) if RCW.cell(r, 1).value == "項目")
inv_rows = [r for r in range(r_e + 1, r_ee + 1) if r != r_eh]
def _rf(r):
    t = str(RCW.cell(r, 1).value)
    if t.startswith("評等級距邊界"): return [f"@Recon!A{r}", f"@Recon!B{r}|sp0", "", ""]
    f = "n2" if ("EPS" in t or "乘數" in t) else "x1"
    return [f"@Recon!A{r}", f"@Recon!B{r}|{f}", f"@Recon!C{r}|{f}", f"@Recon!D{r}|sp1"]
sec(cfg, "recon")["blocks"].append({"type": "table", "caption": f"@Recon!A{r_e}", "head": [f"@Recon!{c}{r_eh}" for c in "ABCD"], "rows": [_rf(r) for r in inv_rows]})


# ================= v0.3.8（2026-11-14）：個股 HTML 版面規格 v1.3（Andy 2026-11-14 08:11、09:03、09:16、12:47；樣板 avgo/gen_html_json.py v0.5.6／v0.5.7） =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳 → 資料狀態。
# 數字與文字取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末、Coord 頁；alab/v038_post.py 產生）。列號依標籤找。
HERE = os.path.dirname(os.path.abspath(__file__))


def _v13():
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RV, CW = wb["README"], wb["PL"], wb["RevDCF"], wb["Coord"]
    r13 = lab("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "coord", "ver"], range(r13 + 1, r13 + 11)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.3.8"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    old_note = list(cfg["hero"]["note"])
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"]
    # ---- 1. 命題：置頂對照圖（模型三情境 vs 市場共識）、投資建議條件式、重點觀察 5 點 ----
    th = S["thesis"]["blocks"]
    rr = lambda p: lab("Recon", p)
    r27, r28, re27, re28 = rr("FY2027 營收 vs 共識"), rr("FY2028 營收 vs 共識"), rr("FY2027 EPS vs 共識"), rr("FY2028 EPS vs 共識")
    grp = lambda nm, r, f: {"name": [nm], "fmt": f, "items": [{"name": ["市場共識"], "v": f"@Recon!B{r}", "ref": True},
                                                          {"name": ["模型 情境 A"], "v": f"@Recon!C{r}"}, {"name": ["模型 情境 B"], "v": f"@Recon!E{r}"}, {"name": ["模型 情境 C"], "v": f"@Recon!G{r}"}]}
    cmp_ = {"type": "compare", "title": "回答命題：模型三情境的營收與 EPS 對市場共識",
            "groups": [grp("FY2027 營收（US$B）", r27, "n2"), grp("FY2028 營收（US$B）", r28, "n2"), grp("FY2027 EPS（US$）", re27, "n2"), grp("FY2028 EPS（US$；共識為推導值）", re28, "n2")],
            "caption": "灰色＝市場共識；青綠＝本模型三情境。FY2026 為錨點年、三情境相同；FY2027 起情境分歧，基準情境的 FY2028 營收與 EPS 低於共識。對照只作事後檢查，校準時未使用。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th.insert(1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
    th.insert(0, cmp_)
    # ---- 2. 估值：四個目標價置頂 ----
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b["type"] == "bars")
    g0 = lab("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 6)); r_px = lab("Val", "現價（", g0)
    va[i_bar] = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
                 "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 0, 1),
                           ("DCF 目標價", "Comparable（情境加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
                 "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
    va.insert(i_bar, {"type": "table", "caption": "四個目標價並列（美元；相對現價）", "head": ["目標價", "美元", "相對現價", "說明"],
                      "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr]})
    va.insert(i_bar + 2, {"type": "note", "parts": old_note[:7]})
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").startswith(p))
    ra = rv("A. DCF"); rb_ = rv("(a)"); rbb = rv("(b)"); rc = rv("(c)"); rcmp = rv("B. Comparable"); rsum = rv("C. 摘要"); yh = rv("項目（情境 B）")
    m1, m0 = rv("現價隱含的 EBIT 率", rb_), rv("模型：基準的等效 EBIT 率", rb_)
    s1, rv1, rv0 = rv("現價隱含的營收倍數", rbb), rv("現價隱含的 FY2028 營收", rbb), rv("模型：FY2028 營收", rbb)
    g1, g0_ = rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
    w1, w0, w1t = rv("現價隱含的 WACC", rc + 1), rv("模型 WACC", rc), rv("現價隱含的 WACC（顯示）", rc)
    a_f = {0: "n2", 1: "n2", 2: "n2", 3: "n2", 4: "n2", 5: "n3", 6: "n1", 7: "n2", 8: "x1", 9: "n2"}
    y_f = lambda r: "p1" if r in (yh + 3, yh + 4) else "n3" if r in (yh + 5, yh + 6, yh + 7, yh + 8) else "n2"
    lbl = lambda r: str(RV.cell(r, 1).value or "")
    rev_ = {
        "id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
        "blocks": [
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
            {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
             "rows": [["預測期 EBIT 率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                      ["營收規模（FY2028 起各年同乘的倍數）", f"@RevDCF!B{s1}|n2", "—"],
                      ["FY2028 營收（十億美元）", f"@RevDCF!B{rv1}|n1", f"@RevDCF!B{rv0}|n1"],
                      ["FY2028 營收年增率", f"@RevDCF!B{g1}|p0", f"@RevDCF!B{g0_}|p1"],
                      ["折現率（WACC；到搜尋下限仍不足時標示無解）", f"@RevDCF!B{w1t}", f"@RevDCF!B{w0}|p2"]]},
            {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"],
             "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{a_f[r - ra]}"] for r in range(ra + 2, ra + 10)]},
            {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
             "rows": [[f"@RevDCF!A{rcmp + 2}", "—", f"@RevDCF!C{rcmp + 2}|n2", "—"]]
                     + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!C{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!D{r}|sp1"] for k, r in enumerate(range(rcmp + 3, rcmp + 6))]
                     + [[f"@RevDCF!A{rcmp + 6}", "—", f"@RevDCF!C{rcmp + 6}|n1", "—"]]},
            {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
             "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）",
             "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{y_f(r)}" for c in "BCDEFG"] for r in range(yh + 1, rb_ - 1)]
                     + [{"sub": f"@RevDCF!A{rb_}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "EBIT 率" in lbl(r) else "n2"), "", "", "", "", ""] for r in range(rb_ + 1, rb_ + 5)]
                     + [{"sub": f"@RevDCF!A{rbb}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "年增率" in lbl(r) else "n2" if "倍數" in lbl(r) else "n3" if "核對" in lbl(r) else "n2"), "", "", "", "", ""] for r in range(rbb + 1, rc - 1)]
                     + [{"sub": f"@RevDCF!A{rc}"}] + [[f"@RevDCF!A{r}", (f"@RevDCF!B{r}" if r == w1t else f"@RevDCF!B{r}|" + ("p2" if lbl(r).startswith(("現價隱含的 WACC", "模型 WACC")) else "n2")), "", "", "", "", ""] for r in range(w1, w1 + 5)]},
        ]}
    # ---- 5. 損益（標準損益表）----
    p0 = lab("PL", "標準損益表（v0.3.8")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
    COLS_ = "CDEFGHIJKL"
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（研發＋銷管）", "營業利益", "營業利益率",
            "其他收入（淨額）＝利息與其他收益", "稅前利益", "所得稅", "稅後淨利", "淨利率", "稀釋股數", "EPS（完全稀釋，non-GAAP）"]
    fmt = lambda u: {"%": "p1", "US$": "n2", "十億股": "n3"}.get(u, "n3")
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
                "fmt": "n1", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:L{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:L{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:L{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:L{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，Non-GAAP 模型口徑）。FY2026E 為錨點；FY2030E–FY2032E 為延伸期：營收與淨利率由收斂法推得，毛利率與營業費用率沿用 FY2029E。"}
    sc_chart = {"type": "theme", "title": "三情境並列：營收（長條，十億美元）與 EPS（折線，美元，右軸）", "x": "@PL!C3:J3",
                "fmt": "n1", "fmt_r": "n2", "stack": False, "pad": 1.1,
                "bars": [{"name": [f"@PL!A{a}", " 營收"], "y": f"@PL!C{b}:J{b}", "color": c} for a, b, c in ((4, 5, "#b7e3e4"), (16, 17, "#10a0a7"), (28, 29, "#8a8c8c"))],
                "lines": [{"name": [f"@PL!A{a}", " EPS"], "y": f"@PL!C{b}:J{b}", "axis": "r", "color": c, **kw}
                          for a, b, c, kw in ((4, 13, "#0b7a80", {"dash": True}), (16, 25, "#3b3838", {"labels": True}), (28, 37, "#b4472f", {"dash": True}))],
                "caption": "情境 A（樂觀）、B（基準）、C（保守）在 FY2026（錨點）相同，FY2027 起情境分歧；FY2030 起為延伸期（收斂法）。"}
    old_pl = S["pl"]["blocks"]
    pls = [b for b in old_pl if b["type"] == "pl"]
    kvs = [b for b in old_pl if b["type"] == "kv"]
    notes = [p_ for b in old_pl if b["type"] == "note" for p_ in b["parts"]]
    assert len(pls) == 2 and len(kvs) == 1 and len(notes) == 1, (len(pls), len(kvs), notes)
    old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境摘要表與折線（FY2025–FY2029；營收、毛利率、淨利、EPS）"),
                  dict(kvs[0], collapse=True, summary="展開：FY2026 錨點的核對"),
                  {"type": "list", "items": notes, "collapse": True, "summary": "展開：延伸期（FY2030–FY2032）的收斂法說明"},
                  dict(pls[1], collapse=True, summary="展開：三情境延伸到 FY2032（收斂法）")]
    S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, sc_chart, gaap_tbl] + old_blocks)
    S["pl"].pop("intro", None)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    hdrB = lab("PL", "項目", lab("PL", "情境 B：基準"))
    rA = lab("PL", "A. Scorpio X", hdrB); rD = lab("PL", "D. 損益", hdrB)
    r_sx, r_pc, r_bu, r_u = lab("PL", "Scorpio X 營收", rA), lab("PL", "PCIe 與訊號調節營收", rA), lab("PL", "營收", lab("PL", "C. 合計", rA)), lab("PL", "AI 加速器顆數", rA)

    def plain(t):
        return t.replace("（c_pcie）", "")

    def seg_row(r):
        a = str(PLW.cell(r, 1).value or "")
        if all(PLW.cell(r, c).value is None for c in range(4, 8)):
            return {"sub": re.sub(r"[（(].*$", "", plain(a))}
        lab_ = f"@PL!A{r}" if not re.search(r"[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+", a) else plain(a)
        f = {"%": "p1", "百萬顆": "n1", "x": "n2"}.get(PLW.cell(r, 2).value, "n2")
        return [lab_, f"@PL!B{r}"] + [(f"@PL!{c}{r}|{f}" if PLW.cell(r, " ABCDEFG".index(c)).value is not None else "") for c in "DEFG"]
    seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；十億美元）：Scorpio X＋PCIe 與訊號調節＝營收",
               "head": [f"@PL!A{hdrB}", f"@PL!B{hdrB}"] + [f"@PL!{c}{hdrB}" for c in "DEFG"],
               "rows": [seg_row(r) for r in range(rA, rD)]}
    seg_chart = {"type": "theme", "title": "分部營收（情境 B，十億美元）與 AI 加速器顆數（百萬顆，右軸）", "x": f"@PL!D{hdrB}:G{hdrB}",
                 "fmt": "n1", "fmt_r": "n1", "stack": True, "pad": 1.2, "total": f"@PL!D{r_bu}:G{r_bu}",
                 "bars": [{"name": ["Scorpio X（scale-up 交換晶片）"], "y": f"@PL!D{r_sx}:G{r_sx}", "color": "#0b7a80"},
                          {"name": ["PCIe 與訊號調節"], "y": f"@PL!D{r_pc}:G{r_pc}", "color": "#6fc7cb"}],
                 "lines": [{"name": ["AI 加速器顆數（乘情境顆數係數）"], "y": f"@PL!D{r_u}:G{r_u}", "axis": "r", "color": "#3b3838", "labels": True}],
                 "caption": "堆疊＝Scorpio X 與 PCIe 與訊號調節兩個分部；頂端數字＝營收合計；折線＝AI 加速器顆數（右軸）。FY2026 為錨點年（有效份額由錨點校準），FY2027 起為模型。"}
    dr_ = S["drivers"]["blocks"]
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + dr_
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    S["drivers"].pop("intro", None)
    # ---- 9. 資料狀態 ----
    st = S["status"]["blocks"]
    st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]},
               {"type": "h3", "text": "三層協調：讀哪一層哪一版、由下往上對帳"}, {"type": "note", "parts": [f"@README!A{RM['coord']}"]}]
    S["revdcf"] = rev_
    # ---- 7. 對帳章：Coord F（讀取層級）、G（層級之間）、I（C0 實績）----
    cr = lambda p, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
    fr, gr_, ir = cr("F. 讀取層級"), cr("G. 層級之間"), cr("I. C0（實績）")
    fh = fr + 1; fe = fh
    while CW.cell(fe + 1, 1).value and not str(CW.cell(fe + 1, 1).value).startswith("下游只註明"): fe += 1
    gh = gr_ + 1; ge = gh
    while CW.cell(ge + 1, 1).value: ge += 1
    ih = cr("輸入", ir); ie = ih
    while CW.cell(ie + 1, 1).value: ie += 1
    rh = next(r for r in range(ie, CW.max_row + 1) if CW.cell(r, 1).value == "項目" and CW.cell(r, 2).value == "公司實績"); re_ = rh
    while CW.cell(re_ + 1, 1).value: re_ += 1
    cnt = cr("對帳中的 C0 項目數")
    lay_read = {"type": "table", "caption": "讀取層級（同源，不判定）：本模型讀哪一層、哪一版、用在哪裡", "collapse": True, "summary": "展開：讀取層級（本模型讀哪一層、哪一版、哪些名稱）",
                "head": [f"@Coord!{c}{fh}" for c in "ABCDE"], "rows": [[f"@Coord!{c}{r}" for c in "ABCDE"] for r in range(fh + 1, fe + 1)]}
    lay_tbl = {"type": "table", "caption": "層級之間（參考；Siliconomics 對帳單處理，不計入本公司對帳中項目）：產業部署 GW，Siliconomics 基準對另一側讀數",
               "head": ["項目", "Siliconomics 基準", "另一側讀數", "差距", "狀態"],
               "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|n1", f"@Coord!C{r}|n1", f"@Coord!D{r}|sp1", f"@Coord!E{r}"] for r in range(gh + 1, ge + 1)]}
    def _cf(r, c):
        v = CW.cell(r, c).value
        if v is None: return ""
        nf = CW.cell(r, c).number_format
        return f"@Coord!{'ABCDEFGHI'[c - 1]}{r}" + ("" if isinstance(v, str) else "|p1" if "%" in nf else "|n3")
    c0_in = {"type": "table", "caption": "實績對帳輸入：公司實績（公司當期財報新聞稿）與層級值、同源校準的證明", "collapse": True, "summary": "展開：實績對帳輸入與同源校準的核對（實績、層級、份額、乘積）",
             "head": [f"@Coord!{c}{ih}" for c in "ABCDF"],
             "rows": [[f"@Coord!A{r}", _cf(r, 2), f"@Coord!C{r}", f"@Coord!D{r}", f"@Coord!F{r}"] for r in range(ih + 1, ie + 1)]}
    c0_res = {"type": "table", "caption": "實績對帳（參考；同源校準）：本年已公布三季的公司營收實績 vs 層級隱含值（兩讀數）",
              "head": ["項目", "公司實績", "層級隱含", "差距", "判定", "說明"],
              "rows": [[f"@Coord!A{r}", (f"@Coord!B{r}|n3" if CW.cell(r, 2).value is not None else ""), (f"@Coord!C{r}|n3" if CW.cell(r, 3).value is not None else ""),
                        (f"@Coord!D{r}|sp1" if CW.cell(r, 4).value is not None else ""), f"@Coord!E{r}", f"@Coord!I{r}"] for r in range(rh + 1, re_ + 1)]}
    S["recon"]["blocks"][1:1] = [{"type": "note", "parts": [f"@Coord!A{ir + 2}"]}, c0_res, {"type": "note", "parts": [f"@Coord!A{ir + 1}"], "collapse": True, "summary": "展開：實績對帳的方法（期間、口徑、兩讀數與判定規則）"},
                                 c0_in, {"type": "note", "parts": [f"@Coord!A{cnt}", "：", f"@Coord!B{cnt}"]}, lay_tbl, lay_read]
    cfg["sections"] = [S[k] for k in ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]]


_v13()


# 敏感度章：兩張「加權目標價相對基準」龍捲風主圖（低案、高案；變動達兩位數百分比者直接顯示，其餘摺疊）；資料取 Sens 頁末表
def _sens_tor():
    SW = wb["Sens"]; t0 = lab("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
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


_se = [x for x in cfg["sections"] if x["id"] == "sens"][0]
for _b in _se["blocks"]:
    if _b.get("type") == "table" and not _b.get("collapse"):
        _b["collapse"] = True; _b["summary"] = "展開：目標價參數表與業務主因表（原敏感度頁明細）"
    if _b.get("type") == "tornado":
        _b["collapse"] = True
        _b["summary"] = "展開：對照：FY2028 EPS 相對基準的變動（情境 B，一次動一個輸入）" + ("；未達兩位數百分比" if "未達" in str(_b.get("title")) else "；達兩位數百分比")
    if _b.get("type") == "note" and not _b.get("collapse"):
        _b["collapse"] = True; _b["summary"] = "展開：敏感度頁原文說明"
_se["blocks"] = _sens_tor() + _se["blocks"]
_se["intro"] = ("主圖為加權目標價相對基準的變動（低案、高案分開；變動達兩位數百分比者直接顯示，其餘摺疊）。營運輸入案例為整本重算（三情境同時套用）；"
                "估值框架案例（情境權重、DCF 權重）為封閉式；折現參數案例引用目標價參數表。下方摺疊為全部案例與明細。")


# 全頁可見文字不出現參數代號（版面規格 v1.3 B 節）：xlsx 原文不改，展示層含代號的格改用白話
PLAIN = {
    "Val!D162": "賣方分析師平均目標價（來源與資料日見輸入頁）",
    "Val!E41": "賣方分析師平均目標價（來源與資料日見輸入頁）",
    "Val!E39": "DCF 與 Comparable 加權（權重見輸入頁目標價規則，可調）",
    "Val!E153": "DCF 每股價值以股權成本推到目標時點（評價日見估值頁原文）",
    "Val!E10": "股權成本與稅後債務成本以市值權重加權；折現率調整只用於敏感度（基準為零）",
    "Val!A11": "永續成長率 g（長期成長上限與無風險利率取小者）",
    "Val!A12": "永續期新投資報酬率 RONIC（＝WACC＋超額報酬差）",
    "Val!E17": "資產負債表日之後的 stub 期：該會計年度 FCFF 乘 stub 期占年比例",
    "Val!A21": "DCF 情境價值推到目標時點（以股權成本複利、扣除期間股利）",
    "In!B172": "收斂法：FY2032 營收年增收斂目標（A 比基準高、B 為基準、C 比基準低，各差兩個百分點）",
    "In!B143": "永續成長率 g（最後一年之後；長期成長上限與無風險利率取小者）",
    "Units!A13": "PCIe 等有效份額（÷ 顆數 × 每顆內容）",
}
FOLD = {}


def _plain(spec):
    if isinstance(spec, list):
        return [_plain(x) for x in spec]
    if isinstance(spec, str) and spec.startswith("@"):
        ref = spec[1:].split("|")[0]
        if ref in PLAIN:
            return PLAIN[ref]
    return spec


def _expand(b):
    """rowgen 表若有格要換成白話，展開成明列（格式與原 rowgen 相同）。"""
    rg = b.get("rowgen")
    if not rg or "-" not in rg["rows"]:
        return b
    r0, r1 = (int(x) for x in rg["rows"].split("-"))
    if not any(f"{rg['sheet']}!{c}{r}" in PLAIN for r in range(r0, r1 + 1) for c in rg["cols"]):
        return b
    rf = rg.get("row_fmt", {})
    rows = [[f"@{rg['sheet']}!{c}{r}" + (lambda f: f"|{f}" if f else "")(rf.get(str(r), {}).get(c, rg["fmts"][i])) for i, c in enumerate(rg["cols"])] for r in range(r0, r1 + 1)]
    return {k: v for k, v in b.items() if k != "rowgen"} | {"rows": rows}


for _sec in cfg["sections"]:
    if _sec["id"] == "status":
        continue
    _sec["blocks"] = [_expand(_b) if _b.get("type") == "table" else _b for _b in _sec["blocks"]]
    for _b in _sec["blocks"]:
        if _b.get("type") == "table":
            _b["rows"] = [({"sub": _plain(r["sub"])} if isinstance(r, dict) else _plain(r)) for r in _b.get("rows", [])]
            if "head" in _b: _b["head"] = _plain(_b["head"])
        if _b.get("type") == "note" and not _b.get("collapse") and len(_b["parts"]) == 1 and str(_b["parts"][0])[1:] in FOLD:
            _b["collapse"] = True; _b["summary"] = FOLD[str(_b["parts"][0])[1:]]
# 三層協調 C4（規則 v1.4；tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(cfg, X)


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


# 主圖置頂：估值章第一個區塊＝四個目標價與現價圖；敏感度章前兩個區塊＝加權目標價相對基準低案、高案主圖。只改區塊順序。
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

json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

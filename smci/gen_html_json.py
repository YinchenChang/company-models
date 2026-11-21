# 美超微 HTML 設定：讀 smci/html_src.json（v0.2 的 html.json，內容不動）→ 套用個股 HTML 版面規格 v1.0 → 寫 smci/html.json。
# 用法：python3 smci/gen_html_json.py <成品 xlsx> smci/html_src.json smci/html.json
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
rc["caption"] = "本模型與共識、公司指引及份額口徑的對照"

_scn = next(re.split(r"[.\s]", str(wb["In"].cell(r, 1).value))[0] for r in range(4, wb["In"].max_row + 1)
            if str(wb["In"].cell(r, 1).value or "").startswith("情境驅動對照"))
AI = "AI GPU 平台（商用 GPU 顆數、每顆系統價值、美超微份額與毛利率）"
GS = "非 AI 營收"
CO = "全公司（第一季實績、費用與淨利、特別股與股數、股價、目標價規則、折現率與自由現金流）"
_ss = next(re.split(r"[.\s]", str(wb["In"].cell(r, 1).value))[0] for r in range(4, wb["In"].max_row + 1)
           if str(wb["In"].cell(r, 1).value or "").startswith("情境專屬列"))
SPEC = {
 "in_skip": ["F", _scn],
 "in_groups": [
  [AI, [
    ["商用 GPU 顆數（讀台積電模型）", ["uo_26", "uo_27", "uo_28", "uo_29", "un_26", "un_27", "un_28", "un_29", "ut_26", "ut_27", "ut_28", "ut_29", "mf"]],
    ["情境顆數係數（2028 年起三情境不同）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["每顆系統價值（讀橋接表）", ["v_gb", "v_vr", "vpg_28", "vpg_29"]],
    ["營收結構與份額（FY2027 以公司指引校準）", ["ais_25", "ais_1h", "ais_q3", "ais_q4", "g27", "rel_28", "rel_29", "g27_A", "rel_28_A", "rel_29_A", "g27_C", "rel_28_C", "rel_29_C"]],
    ["毛利率（A 最有利、C 最不利）", ["gm_nai", "dgm_27", "dgm_28", "dgm_27_A", "dgm_28_A", "dgm_27_C", "dgm_28_C"]],
    ["Wattonomics 錨定係數（只用於敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  [GS, [
    ["成長率", ["g_nai27", "g_nai28", "g_nai29"]]]],
  [CO, [
    ["第一季實績（對帳用）", ["q1_rev", "q1_gm"]],
    ["營業費用、淨利係數、特別股與股數（FY2027 起）", ["phi", "opx_adj", "nio_f", "pref_div", "sh27", "sh_g", "conv_sh", "sbc_ps"]],
    ["股價、市場倍數與共識", ["px", "pe_ttm", "fpe", "cons_feps", "cons_tp", "nim10"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_wA", "tp_wB", "tp_wC", "tp_w_dcf"]],
    ["收斂法（FY2030 起）", ["tp_g", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta", "beta_u", "beta_de", "beta_du", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨負債與股數（評價用；含特別股轉換）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]]]]],
 "blocks_to": {"商用 GPU 系統市場": AI, "@Units!A16": AI},
 "main": {"本模型與共識": ["FY2027 營收 vs 共識", "FY2027 EPS（模型＋股票報酬調整）vs 共識", "FY2028 EPS（模型＋股票報酬調整）vs 共識", "FY2027 營收 vs 指引中點", "FY2027 有效份額"]},
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
                [f"@Val!A{r_coef}", f"@Val!B{r_coef}|n2"], [f"@Val!A{r_cmp}", f"@Val!B{r_cmp}|n0"]]},
      dict({k: v for k, v in cmp_src.items() if k != "caption"}, collapse=True, summary="展開：同業前瞻本益比"),
      {"type": "table", "collapse": True, "summary": "展開：相對倍數調整的組成（本公司與同業的長期本益比）", "head": ["項目", "數值"],
       "rows": [[k, v] for k, v in rm_kv["items"] if not v.startswith((f"@Val!B{r_coef}", f"@Val!B{r_cmp}"))]},
      {"type": "note", "parts": [f"@Val!E{r_med}"]},
      {"type": "h3", "text": "參考：舊法（十年本益比規則值）與市場隱含倍數（不進目標價）"},
      {"type": "note", "parts": ["@Val!A1"], "collapse": True, "summary": "展開：本頁原說明"}]
# 舊法各塊：依原順序全部摺疊
old_kv = next(b for b in vs if b["type"] == "kv" and b["items"][0][0] == f"@Val!A{r_px}")
nb.append({"type": "table", "collapse": True, "summary": "展開：現價與市場隱含倍數", "head": ["項目", "數值"],
           "rows": [[k, v] for k, v in old_kv["items"]]})
i10 = next(i for i, b in enumerate(vs) if b["type"] == "h3" and str(b.get("text", "")).startswith("十年"))
NOTE_SUM = {"@Val!A2": "循環位置的判定方式", "@Val!A82": "What must be true 表的讀法與先驗"}
for b in vs[i10 + 1:]:
    if b.get("collapse"):
        nb.append(b)
    elif b["type"] == "note":
        nb.append(dict(b, collapse=True, summary="展開：" + NOTE_SUM.get(b["parts"][0], "舊法的補充說明")))
    elif b["type"] == "table":
        hd = str(wb["Val"][b["head"][0][5:]].value or "") if b["head"] else ""
        nm = "規則值隱含股價" if hd.startswith("情境") and b["rowgen"]["cols"][1] == "B" and "E" in b["rowgen"]["cols"] else "循環位置"
        nb.append(dict(b, collapse=True, summary=f"展開：{nm}"))
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
    if b.get("caption") == "@Sens!A28":
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
# ================= v0.3.7（2026-11-15）：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10；同 lineage 緯穎 v0.2.6）＋三層協調 C0 =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含 C0）→ 資料狀態；美超微無個別模組。
# 數字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末、Coord 頁 C0；smci/v037_post.py 產生）。列號依標籤找。
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "coord"))
import coord as COORD  # noqa: E402

WBF = openpyxl.load_workbook(X, data_only=True)


def find(sh, prefix, start=1, col=1, exact=False):
    ws = WBF[sh]
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and ((v.strip() == prefix) if exact else v.strip().startswith(prefix)):
            return r
    raise KeyError((sh, prefix))


S = {s_["id"]: s_ for s_ in cfg["sections"]}
RD, PLW, RV = WBF["README"], WBF["PL"], WBF["RevDCF"]
r13 = find("README", "v1.3 展示層")
RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
assert str(RD[f"A{RM['ver']}"].value).startswith("v0.3.9"), RD[f"A{RM['ver']}"].value
cfg["layout"] = "v1.3"
old_note = cfg["hero"]["note"]
cfg["hero"]["note"] = [f"@README!A{RM['hero']}"]            # 橫幅只留 ≤200 字白話；評等規則與相對倍數調整說明移到估值章

# ---- 1. 命題與核心問題 ----
rbB = find("PL", "情境 B（基準）：")
prB = lambda p, ex=False: find("PL", p, rbB, exact=ex)
hdB = rbB + 1
r_ai, r_oiB, r_epsB = prB("AI GPU 平台營收", True), prB("營業利益", True), prB("EPS（GAAP 稀釋，兩類法）", True)
core_chart = {"type": "theme", "title": "AI GPU 平台放量撐不撐得住 EPS：AI GPU 平台營收、營業利益（長條，十億美元）與 GAAP EPS（折線，美元，右軸；情境 B）",
              "x": f"@PL!C{hdB}:G{hdB}", "fmt": "n0", "fmt_r": "n2", "stack": False, "pad": 1.25,
              "bars": [{"name": ["AI GPU 平台營收"], "y": f"@PL!C{r_ai}:G{r_ai}", "color": "#6fc7cb"},
                       {"name": ["營業利益"], "y": f"@PL!C{r_oiB}:G{r_oiB}", "color": "#0b7a80"}],
              "lines": [{"name": ["EPS（GAAP 稀釋，美元）"], "y": f"@PL!C{r_epsB}:G{r_epsB}", "axis": "r", "color": "#3b3838", "labels": True}],
              "caption": "回答核心問題：AI GPU 平台營收＝商用 GPU 系統市場（會計年度各取兩個日曆年一半）× 美超微份額；FY2027 份額以公司營收指引校準，之後份額相對值下修，情境 B 的顆數自 CY2028 起縮減。FY2025、FY2026 為實績，FY2027 起依賴對帳中年份。"}
th = S["thesis"]["blocks"]
i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
th.insert(0, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
th.insert(0, core_chart)

# ---- 2. 估值：四個目標價與現價圖置頂（G 節位置）；原橫幅評等規則與相對倍數調整說明放在圖下 ----
va = S["valuation"]["blocks"]
i_bar = next(i for i, b in enumerate(va) if b["type"] == "bars" and b.get("title") == "目標價與現價")
g0 = find("Val", "0.7 四個目標價並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = find("Val", "現價（", g0)
bars4 = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n0",
         "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                   ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
         "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
        "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{g0 + 7}", f"@Val!B{g0 + 7}|n2", "", f"@Val!D{g0 + 7}"]]}
note_rule = {"type": "note", "parts": list(old_note)}
conc = va[i_bar + 1]
assert conc.get("caption") == "結論（主要項目）", conc.get("caption")
conc = dict(conc, collapse=True, summary="展開：結論的主要項目（四個目標價之外的上檔空間、投資建議、市場共識等）")
va[i_bar + 1] = conc
del va[i_bar]
_vi = S["valuation"].pop("intro", None)
va[0:0] = [bars4, {"type": "note", "parts": [_vi]} if _vi else {"type": "note", "parts": [""]}, tbl4, note_rule]

# ---- 4. 反向 DCF ----
rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").startswith(p))
ra = rv("A. DCF"); rb_ = rv("(a)"); rbb = rv("(b)"); rc = rv("(c)"); rcmp = rv("B. Comparable"); rsum = rv("C. 摘要"); yh = rv("項目（情境 B）")
m1, m0 = rv("現價隱含的營業利益率", rb_), rv("模型：基準的等效營業利益率", rb_)
s1, rv1, rv0, g1, g0_ = rv("現價隱含的營收倍數", rbb), rv("現價隱含的 FY2028 營收", rbb), rv("模型：FY2028 營收", rbb), rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
w1, w0 = rv("現價隱含的 WACC", rc), rv("模型 WACC", rc)
rev_ = {
    "id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
    "blocks": [
        {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
        {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
        {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
         "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                  ["營收規模（FY2028 起各年同乘的倍數）", f"@RevDCF!B{s1}|n2", "—"],
                  ["FY2028 營收（十億元）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
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
         "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|" + ("p1" if r in (yh + 3, yh + 4) else "n3" if r in (yh + 5, yh + 6, yh + 7, yh + 8) else "n1") for c in "BCDEFG"] for r in range(yh + 1, rb_ - 1)]
                 + [{"sub": f"@RevDCF!A{rb_}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "利益率" in str(RV.cell(r, 1).value) else "n1"), "", "", "", "", ""] for r in range(rb_ + 1, rb_ + 5)]
                 + [{"sub": f"@RevDCF!A{rbb}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "年增率" in str(RV.cell(r, 1).value) else "n2" if "倍數" in str(RV.cell(r, 1).value) else "n3" if "核對" in str(RV.cell(r, 1).value) else "n1"), "", "", "", "", ""] for r in range(rbb + 1, rc - 1)]
                 + [{"sub": f"@RevDCF!A{rc}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p2" if str(RV.cell(r, 1).value).startswith(("現價隱含的 WACC", "模型 WACC")) else "n2"), "", "", "", "", ""] for r in range(w1, w1 + 4)]},
    ]}

# ---- 5. 損益（標準損益表） ----
p0 = find("PL", "標準損益表（v0.3.7")
ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（銷管＋研發）", "營業利益", "營業利益率", "其他收入（淨額：利息、匯兌、投資收益等）",
        "稅前利益", "所得稅", "特別股股利", "稅後淨利（歸屬母公司）", "淨利率", "股數（十億股）", "EPS（GAAP 稀釋；美元）"]
fmt = lambda u: {"%": "p1", "US$": "n2", "十億股": "n3"}.get(u, "n1")
PC9 = "CDEFGHIJK"
rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in PC9]
r_ratio = plr("比率"); r_dil = plr("基本 EPS 對照（歷史；預測不估基本）"); r_chk = plr("核對（應為 0）")
r_d31 = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("差額（FY2030E、FY2031E）：毛利"))
is_tbl = {"type": "table", "caption": "標準損益表（GAAP 口徑，十億美元；EPS 為美元、兩類法）：FY2023A–FY2026A 實績、FY2027E–FY2031E 情境 B（基準）；標「對帳中」的年度依賴 Siliconomics 對帳中年份（會計年度七月至隔年六月）",
          "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC9],
          "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_dil)]
                  + [{"sub": "核對：FY2030E、FY2031E 兩種推法的營業利益差額（毛利率與營業費用率沿用前一年）"}] + [rowx(r) for r in (r_d31, r_d31 + 1)]}
dil_tbl = {"type": "table", "caption": f"@PL!A{r_dil}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in "CDE"], "collapse": True,
           "summary": "展開：基本 EPS 對照（歷史）與核對列",
           "rows": [[f"@PL!A{r_dil + 1}"] + [f"@PL!{c}{r_dil + 1}|n2" for c in "CDE"]] + [{"sub": f"@PL!A{r_chk}"}]
                   + [[f"@PL!A{r}"] + [f"@PL!{c}{r}|n3" for c in "CDE"] for r in range(r_chk + 1, PLW.max_row + 1) if r not in (r_d31, r_d31 + 1)]}
r_lbl = ph
rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
is_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率、淨利率：FY2023A–FY2031E（情境 B）", "x": f"@PL!C{r_lbl}:K{r_lbl}",
            "fmt": "n0", "fmt_r": "p1", "stack": True, "pad": 1.6,
            "bars": [{"name": [f"@PL!A{rv_}"], "y": f"@PL!C{rv_}:K{rv_}", "color": "#6fc7cb"}],
            "lines": [{"name": [f"@PL!A{gm_}"], "y": f"@PL!C{gm_}:K{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                      {"name": [f"@PL!A{opm_}"], "y": f"@PL!C{opm_}:K{opm_}", "axis": "r", "color": "#c29a2b"},
                      {"name": [f"@PL!A{nm_}"], "y": f"@PL!C{nm_}:K{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
            "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP 口徑）。FY2026A 為實績錨點；FY2027E 起依賴對帳中年份（FY2027E 營收以公司指引校準）；FY2030E–FY2031E 為延伸期：營收與淨利由收斂法推得，毛利率與營業費用率沿用前一年。"}
rbA, rbC = find("PL", "情境 A（樂觀）："), find("PL", "情境 C（保守）：")
xh = hdB
xr = lambda base, lab_: next(r for r in range(base, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
R_REV, R_EPS = "營收", "EPS（GAAP 稀釋，兩類法）"
sc_chart = {"type": "theme", "title": "三情境並列：營收（長條，十億美元）與 GAAP EPS（折線，美元，右軸）：FY2025A–FY2032", "x": f"@PL!C{xh}:J{xh}",
            "fmt": "n0", "fmt_r": "n2", "stack": False, "pad": 1.1,
            "bars": [{"name": ["情境 A（樂觀）營收"], "y": f"@PL!C{xr(rbA, R_REV)}:J{xr(rbA, R_REV)}", "color": "#b7e3e4"},
                     {"name": ["情境 B（基準）營收"], "y": f"@PL!C{xr(rbB, R_REV)}:J{xr(rbB, R_REV)}", "color": "#10a0a7"},
                     {"name": ["情境 C（保守）營收"], "y": f"@PL!C{xr(rbC, R_REV)}:J{xr(rbC, R_REV)}", "color": "#8a8c8c"}],
            "lines": [{"name": ["情境 A EPS"], "y": f"@PL!C{xr(rbA, R_EPS)}:J{xr(rbA, R_EPS)}", "axis": "r", "color": "#0b7a80", "dash": True},
                      {"name": ["情境 B EPS"], "y": f"@PL!C{xr(rbB, R_EPS)}:J{xr(rbB, R_EPS)}", "axis": "r", "color": "#3b3838", "labels": True},
                      {"name": ["情境 C EPS"], "y": f"@PL!C{xr(rbC, R_EPS)}:J{xr(rbC, R_EPS)}", "axis": "r", "color": "#b4472f", "dash": True}],
            "caption": "三情境自 FY2027 即不同（營收指引上緣、中點、下緣），FY2028 起再加上顆數係數、份額與毛利率差異；FY2030 起為延伸期（收斂法）。"}
old_pl = S["pl"]["blocks"]
pls = [b for b in old_pl if b["type"] == "pl"]
notes = [p_ for b in old_pl if b["type"] == "note" for p_ in b["parts"]]
kvs = [b for b in old_pl if b["type"] == "kv"]
assert len(pls) == 2 and len(kvs) == 1, (len(pls), len(kvs))
old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境摘要表與折線（FY2025–FY2032；營收、AI GPU 平台、毛利率、EPS）"), dict(kvs[0], collapse=True, summary="展開：三情境錨點與驅動核對"),
              dict(pls[1], collapse=True, summary="展開：三情境摘要（另一張）")] + ([{"type": "note", "parts": notes, "collapse": True, "summary": "展開：延伸期的收斂法說明"}] if notes else [])
S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, sc_chart, dil_tbl] + old_blocks)

# ---- 6. 驅動因子：分部營收推導（情境 B） ----
ps = lambda p: prB(p)
rows_seg = [("AI GPU 平台營收＝商用 GPU 系統市場（會計年度＝前後兩個日曆年各半）× 美超微有效份額", [ps("GB300 世代商用 GPU 顆數（CY）"), ps("Rubin 世代商用 GPU 顆數（CY）"), ps("商用 GPU 系統市場（CY）"), ps("商用 GPU 系統市場（FY"), ps("美超微有效份額（2H FY26 校準）"), ps("FY2027 營收（公司指引"), ps("份額相對值"), ps("美超微有效份額（FY2027 以指引校準）"), prB("AI GPU 平台營收", True)]),
            ("非 AI 營收＝FY2026 實績 × 成長率", [prB("非 AI（企業、儲存、邊緣等）", True)]),
            ("合計：合併營收", [prB("營收", True)])]
fm_seg = lambda u: {"百萬顆": "n2", "US$B": "n1", "%": "p1", "x": "n2"}.get(u, "n1")
seg_rows = []
for t_, rs in rows_seg:
    seg_rows.append({"sub": t_})
    seg_rows += [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{fm_seg(PLW.cell(r, 2).value)}" for c in "DEFG"] for r in rs]
seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；FY2026 為實績錨點）：驅動因子 → 各分部營收 → 合併營收", "head": [f"@PL!A{hdB}", f"@PL!B{hdB}"] + [f"@PL!{c}{hdB}" for c in "DEFG"], "rows": seg_rows}
r_nai, r_tot = prB("非 AI（企業、儲存、邊緣等）", True), prB("營收", True)
seg_chart = {"type": "theme", "title": "分部營收（情境 B，十億美元）：AI GPU 平台、非 AI（企業、儲存、邊緣）", "x": f"@PL!C{hdB}:G{hdB}", "fmt": "n0", "fmt_r": "n0", "stack": True, "pad": 1.15, "total": f"@PL!C{r_tot}:G{r_tot}",
             "bars": [{"name": ["AI GPU 平台"], "y": f"@PL!C{r_ai}:G{r_ai}", "color": "#0b7a80"}, {"name": ["非 AI"], "y": f"@PL!C{r_nai}:G{r_nai}", "color": "#6fc7cb"}],
             "lines": [], "caption": "堆疊＝兩個分部營收（FY2025A、FY2026A 為營收 × 公司口述 AI 占比推得）；頂端數字＝合併營收（取 PL 頁合計列，不在產生器加總）。AI GPU 平台營收由商用 GPU 系統市場與美超微份額推導，見下表。"}
S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + S["drivers"]["blocks"]
S["drivers"]["title"] = "驅動因子：分部營收推導"

# ---- 8. 對帳：C0（實績）＋讀取層級（同源，不判定） ----
if "Coord" in WBF.sheetnames:
    CW = WBF["Coord"]
    cr = lambda p, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
    c0 = cr("C0 由下往上對帳"); c1 = cr("1. 讀哪一層", c0); c2 = cr("2. 讀取層級", c0); c3 = cr("3. 結論", c0); r_cc = cr("C0 結論（白話）", c3)
    c4 = cr("4. C0（實績）", c3); r_m = cr("4.1 實績營收", c4); r_p = cr("4.2 口徑換算", c4); r_cmp = cr("4.3 對帳", c4)
    r_meth = cr("方法：", c4); r_rem = cr("為什麼不判定", c4); r_end = CW.max_row
    nfmt = lambda r: next((CW[f"{c}{r}"].number_format for c in "BCDE" if isinstance(CW[f"{c}{r}"].value, (int, float)) or (isinstance(CW[f"{c}{r}"].value, str) and CW[f"{c}{r}"].value.startswith("=") and c in "BC")), "General")
    fcode = lambda a, nf: "" if ("判定：" in a or "是否都在" in a) else ("sp1" if nf.startswith("+") else "p1" if "%" in nf else "n3" if nf == "0.000" else "n2" if nf == "0.00" else "n1")
    lay_rows = []
    for r in range(c1 + 2, c2 - 1):
        if CW.cell(r, 1).value:
            lay_rows.append([f"@Coord!{c}{r}" for c in "ABCDE"])
    c0_rows, blk = [], ""
    for r in range(c2 + 2, c3 - 1):
        a = str(CW.cell(r, 1).value or "")
        if not a:
            continue
        if not a.startswith("\u3000"):
            c0_rows.append({"sub": f"@Coord!A{r}"}); blk = a
            continue
        nf = next((CW[f"{c}{r}"].number_format for c in "BCDE" if CW[f"{c}{r}"].value is not None), "General")
        f_ = "" if "判定" in a else ("sp1" if nf.startswith("+") else "p1" if "%" in nf else "n4" if nf == "0.0000" else "n2" if nf == "0.00" else "n1")
        cx = lambda c: f"@Coord!{c}{r}" if CW[f"{c}{r}"].value else ""
        c0_rows.append([f"@Coord!A{r}"] + [f"@Coord!{c}{r}|{f_}" for c in "BCDE"] + [cx("F"), cx("G")])
    act_rows = []
    for r in range(r_cmp + 1, r_end + 1):
        a = str(CW.cell(r, 1).value or "")
        if not a:
            continue
        f_ = fcode(a, nfmt(r))
        act_rows.append([f"@Coord!A{r}", f"@Coord!B{r}|{f_}", (f"@Coord!C{r}|{f_}" if CW[f"C{r}"].value is not None and not (isinstance(CW[f"C{r}"].value, str) and not CW[f"C{r}"].value.startswith("=")) else ""),
                         (f"@Coord!D{r}" if CW[f"D{r}"].value else ""), (f"@Coord!E{r}" if CW[f"E{r}"].value else "")])
    in_rows = []
    for r in range(r_m + 1, r_cmp):
        a = str(CW.cell(r, 1).value or "")
        if not a:
            continue
        if not a.startswith("\u3000"):
            in_rows.append({"sub": f"@Coord!A{r}"})
            continue
        f_ = fcode(a, nfmt(r))
        if "核對" in a:
            f_ = "n3"
        in_rows.append([f"@Coord!A{r}", f"@Coord!B{r}|{f_}", (f"@Coord!D{r}" if CW[f"D{r}"].value else ""), (f"@Coord!E{r}" if CW[f"E{r}"].value else "")])
    c0_blocks = [{"type": "h3", "text": "三層協調：實績營收 × 口述占比對照（無可比實績，不判定）"},
                 {"type": "note", "parts": [f"@Coord!B{r_cc}"]},
                 {"type": "note", "parts": [f"@Coord!A{r_meth}"]},
                 {"type": "table", "caption": "實績營收 × 公司口述 AI solutions 占比（換算，Assumed）對照（參考，不判定：缺 AI 營收金額揭露，美超微份額由本公司上半年換算值反推，無可比實績）：2026 年已公布三季 AI GPU 平台營收對層級隱含值的差（金額十億美元）",
                  "head": ["項目", "主讀數（已公布季度累計）", "第二讀數（最近一季年化）", "單位與口徑", "來源與證據等級"], "rows": act_rows},
                 {"type": "note", "parts": [f"@Coord!A{r_rem}"]},
                 {"type": "table", "caption": "換算輸入（逐季實績營收 × 口述占比、核對列）與參考值", "head": ["項目", "值", "單位與口徑", "來源與證據等級"], "rows": in_rows,
                  "collapse": True, "summary": "展開：逐季換算值、核對列與出處"},
                 {"type": "table", "caption": "讀取層級（同源，不判定）：本模型自 Siliconomics 層讀入的值與層級值必然相同，差距恆為零，無檢驗力，不計入對帳中（情境 B 基準對 Siliconomics 基準；金額十億美元）", "head": ["項目", "CY2026", "CY2027", "CY2028", "CY2029", "單位與口徑", "來源與證據等級"], "rows": c0_rows,
                  "collapse": True, "summary": "展開：讀取層級（同源，不判定）與層級之間的參考差距"},
                 {"type": "table", "caption": "讀哪一層、哪一版（下游只註明版本、不自行重估）", "head": ["項目", "值", "版本／檔案", "本模型用途", "備註"], "rows": lay_rows,
                  "collapse": True, "summary": "展開：本模型讀哪一層、哪一版（名稱、值、版本）"},
                 {"type": "h3", "text": "共識、公司指引與份額口徑的對照"}]
    S["recon"]["blocks"] = c0_blocks + S["recon"]["blocks"]

# ---- 9. 資料狀態：本版要點、歷史資料來源 ----
st = S["status"]["blocks"]
r_h1 = find("Hist", "Hist — 美超微損益表歷史")
st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]},
           {"type": "h3", "text": "標準損益表歷史資料來源"}, {"type": "note", "parts": [f"@Hist!A{r_h1}"]}]

# ---- 3. 敏感度：加權目標價相對基準的低案、高案主圖置頂（H 節位置） ----
SEW = WBF["Sens"]
t0 = find("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
while SEW.cell(r2 + 1, 1).value:
    r2 += 1
tors = []
for col, side in (("B", "低"), ("C", "高")):
    big = [r for r in range(r1, r2 + 1) if isinstance(SEW[f"{col}{r}"].value, (int, float)) and abs(SEW[f"{col}{r}"].value) >= 0.1]
    small = [r for r in range(r1, r2 + 1) if isinstance(SEW[f"{col}{r}"].value, (int, float)) and 0.0005 <= abs(SEW[f"{col}{r}"].value) < 0.1]
    base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
    tors.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
    tors.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                     summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例（幾乎為零者見全部案例表）"))
tors.append({"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
             "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1, r2 + 1)]})
tb0 = find("Sens", "目標價敏感度：業務主因")
n_biz = 0
while SEW.cell(tb0 + 2 + n_biz, 1).value and isinstance(SEW.cell(tb0 + 2 + n_biz, 3).value, (int, float)):
    n_biz += 1
tors.append({"type": "table", "caption": f"@Sens!A{tb0}", "collapse": True, "summary": "展開：業務主因案例的加權目標價（整本重算快照；含快照狀態）",
             "head": [f"@Sens!{c}{tb0 + 1}" for c in "ACDEFGH"],
             "rows": [[f"@Sens!A{r}", f"@Sens!C{r}|n2", f"@Sens!D{r}|sp1", f"@Sens!E{r}|n2", f"@Sens!F{r}|sp1", f"@Sens!G{r}|n2", f"@Sens!H{r}"] for r in range(tb0 + 2, tb0 + 2 + n_biz)]})
sb = S["sens"]["blocks"]
for b in sb:
    if b.get("type") == "tornado":
        b["summary"] = "展開：對照：FY2028 EPS 相對基準的變動（情境 A，一次動一個輸入）" if not b.get("collapse") else b.get("summary", "展開：FY2028 EPS 相對基準的變動（未達兩位數百分比）")
        b["collapse"] = True
    if b.get("type") == "table" and not b.get("collapse") and str(b.get("caption", "")) == f"@Sens!A{r_tp}":
        b["collapse"] = True; b["summary"] = "展開：目標價參數表（WACC、β、ERP、永續成長率等；加權目標價與相對基準）"
    if b.get("type") == "h3" and b.get("text") == "目標價參數":
        b["text"] = "目標價參數與 EPS 對照"
_it = {k: S[k].pop("intro", None) for k in ("valuation", "sens")}
tors = [tors[0], tors[2], tors[1], tors[3]] + tors[4:]
S["sens"]["blocks"] = tors[:2] + \
    [{"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案＝輸入取敏感度低值、高案＝取高值；一次動一個輸入，其他不變）：業務主因為整本重算快照，目標價參數為閉式公式；變動達兩位數百分比者直接顯示，其餘摺疊。下方摺疊為全部案例與 FY2028 EPS 對照。"]}] + tors[2:] + sb

# ---- 順序與對帳中兩讀數（tools/coord/coord.py）：橫幅一格與揭露句、估值章兩讀數表與長條圖放在四個目標價圖與表之後 ----
cfg["sections"] = [S[k] for k in ("thesis", "valuation", "sens")] + [rev_] + [S[k] for k in ("pl", "drivers", "recon", "status")]
n_before = len(S["valuation"]["blocks"])
COORD.inject_html(cfg, X)
va = S["valuation"]["blocks"]
if len(va) > n_before:                                           # inject_html 在最前面插入 4 個區塊（h3、note、table、bars）
    inj, rest = va[:4], va[4:]
    assert inj[0].get("type") == "h3" and inj[3].get("type") == "bars", inj
    S["valuation"]["blocks"] = rest[:3] + inj + rest[3:]
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

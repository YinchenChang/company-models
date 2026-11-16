# 台積電：把 html_src.json（v1.0 之前的頁面設定，內容不動）依「個股 HTML 版面規格 v1.0」（docs/plan/個股HTML版面規格.md）轉成 html.json。
# 用法：python3 tsmc/gen_html_json.py <成品 xlsx>（讀 tsmc/html_src.json，寫 tsmc/html.json；勿手改 html.json）
# ─── 共通段（各公司 gen_html_json.py 相同；只依賴 tools/company_html/relayout.py，不修改它） ───
import copy, json, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
import relayout as RL  # noqa: E402

X = sys.argv[1]
WB = openpyxl.load_workbook(X, data_only=True)
SRC = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))
cfg = copy.deepcopy(SRC)
SEC = {s["id"]: s for s in cfg["sections"]}


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def row_of(sh, prefix, col=1, need_num=False, exact=False, start=1):
    """第一個 A 欄（或 col 欄）以 prefix 開頭的列；need_num＝B 欄須為數字。"""
    ws = WB[sh]
    for r in range(start, ws.max_row + 1):
        t = str(ws.cell(r, col).value or "").strip()
        if (t == prefix if exact else t.startswith(prefix)) and (not need_num or num(ws.cell(r, 2).value)):
            return r
    raise KeyError((sh, prefix))


IN_ROW = {str(WB["In"].cell(r, 1).value): r for r in range(4, WB["In"].max_row + 1) if WB["In"].cell(r, 1).value}


def in_skip_token(prefix="情境驅動對照"):
    """In 頁「情境驅動對照」節標題在轉換器內的節代號（依 relayout._check_in_coverage 的切法）。"""
    ws = WB["In"]
    for r in range(4, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "")
        if a.startswith(prefix) and ws.cell(r, 2).value in (None, ""):
            return re.split(r"[.\s]", a)[0]
    return None


def make_hero(px_ref=None):
    """投資建議橫幅：加權目標價與「加權目標價實際使用的 Comparable」（相對倍數調整後）各附上檔空間。"""
    tp = row_of("Val", "目標價（DCF 與 Comparable 加權）", need_num=True)
    up = row_of("Val", "上檔空間：加權目標價", need_num=True)
    rec = row_of("Val", "投資建議", exact=True)
    try:
        tpc = row_of("Val", "Comparable 目標價（相對倍數調整後）", need_num=True)
        upc_ref = f"@README!B{row_of('README', '展示層輔助：Comparable 目標價（相對倍數調整後）的上檔空間')}"
    except KeyError:
        tpc = row_of("Val", "Comparable 目標價", need_num=True)
        upc_ref = f"@Val!B{row_of('Val', '上檔空間：Comparable', need_num=True)}"
    if px_ref is None:
        try:
            p = row_of("Val", "現價", need_num=True)
            px_ref = (f"@Val!A{p}", f"@Val!B{p}")
        except KeyError:
            p = IN_ROW["px"]
            px_ref = (f"@In!B{p}", f"@In!E{p}")
    h = {"rec": f"@Val!B{rec}",
         "tiles": [[f"@Val!A{tp}", f"@Val!B{tp}|n2", ["上檔空間 ", f"@Val!B{up}|sp1"]],
                   [f"@Val!A{tpc}", f"@Val!B{tpc}|n2", ["上檔空間 ", upc_ref + "|sp1"]],
                   [px_ref[0], px_ref[1] + "|n2", ""]]}
    if WB["Val"].cell(rec, 5).value:
        h["note"] = ["投資建議依加權目標價判定：", f"@Val!E{rec}"]
    return h


def run(spec, hero=None):
    """套用轉換器；橫幅改用 make_hero()（轉換器以第一個「Comparable 目標價」列為準，與加權目標價實際使用的相對倍數調整後 Comparable 不同）。"""
    tor_titles = {(b.get("sheet"), b.get("value_col"), b.get("label_col")): b.get("title") for s in cfg["sections"] for b in s["blocks"] if b["type"] == "tornado"}
    RL._hero = lambda w: hero or make_hero()
    out = RL.relayout(cfg, X, spec)
    # 龍捲風標題：轉換器一律改寫為「加權目標價…」；原圖若不是加權目標價，還原原標題（避免把 EPS 敏感度寫成目標價敏感度）
    for s in out["sections"]:
        for b in s["blocks"]:
            if b["type"] != "tornado":
                continue
            t0 = tor_titles.get((b.get("sheet"), b.get("value_col"), b.get("label_col")))
            if t0 and not t0.startswith("加權目標價"):
                b["title"] = t0 + ("（未達兩位數百分比的案例）" if b.get("collapse") else "（變動達兩位數百分比的案例）")
    # 附錄「版本要點」只收版本說明列（v0.／日期開頭），不收舊的投資建議與重點觀察列
    st = next(s for s in out["sections"] if s["id"] == "status")
    for b in st["blocks"]:
        if b.get("summary") == "展開：本版與前版的修改要點":
            ws = WB["README"]
            b["rows"] = [[f"@README!A{r}"] for r in range(1, ws.max_row + 1)
                         if re.match(r"^(v\d+\.\d|\d{4}-\d{2}-\d{2})", str(ws.cell(r, 1).value or ""))]
    return out


def hero_rule_a(out):
    """補充規則 A（2026-11-05 20:20）：橫幅中間格＝只看 Comparable（市場主流同業法、未經相對倍數調整）＋自己的上檔空間；
    有相對倍數調整且數值不同時，附註寫出進加權的調整後 Comparable。"""
    try:
        tpm = row_of("Val", "Comparable 目標價（市場主流同業法）", need_num=True)
    except KeyError:
        tpm = row_of("Val", "Comparable 目標價", need_num=True, exact=True)
    upm = row_of("Val", "上檔空間：Comparable", need_num=True)
    h = out["hero"]
    h["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{tpm}|n2", ["上檔空間 ", f"@Val!B{upm}|sp1"]]
    try:
        adj = row_of("Val", "Comparable 目標價（相對倍數調整後）", need_num=True)
    except KeyError:
        adj = None
    if adj and abs(WB["Val"].cell(adj, 2).value - WB["Val"].cell(tpm, 2).value) > 1e-9:
        note = h.get("note") or []
        note = (note if isinstance(note, list) else [note]) + ["；進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{adj}|n2", " 元。"]
        h["note"] = note
    return out


def dump(out):
    json.dump(out, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("ok")
# ─── 共通段結束 ───

# ─── 台積電專屬：依業務性質分塊與估值、對帳版面 ───
I = IN_ROW
inr = lambda k, f: [f"@In!B{I[k]}", f"@In!E{I[k]}|{f}"]

# 驅動因子：舊「單一參數」表（含推導量欄）移到附錄當「驅動 → 推導對照」；CoWoS 供需表加標題以便歸塊
dr = SEC["drivers"]
nb = []
for b in dr["blocks"]:
    cap = str(b.get("caption") or "")
    if cap.startswith("單一參數"):
        nb.append(dict(b, caption="驅動 → 推導對照（單一參數：輸入、基準值、推導量與因果方向）"))
    elif b["type"] == "h3":
        continue
    elif b["type"] == "table" and not cap:
        nb.append(dict(b, caption="CoWoS 供需（瓶頸檢查；千片）"))
    else:
        nb.append(b)
dr["blocks"] = nb


def scn_table_fix(sec):
    """v0.5.2（第三輪修正第 32 點）：情境驅動表依 In 頁「情境驅動對照」節標題找列範圍（不寫死列號）；百分比列以 % 顯示。"""
    ws = WB["In"]
    t = next(r for r in range(4, ws.max_row + 1) if str(ws.cell(r, 1).value or "").startswith("情境驅動對照"))
    h = t + 1; rows = []; r = h + 1
    while ws.cell(r, 1).value not in (None, ""):
        rows.append(r); r += 1
    assert ws.cell(rows[-1], 1).value == "scn_w", ws.cell(rows[-1], 1).value
    def fm(rr):
        u = str(ws.cell(rr, 3).value or "")
        vals = [ws.cell(rr, c).value for c in (4, 5, 6)]
        if u in ("%", "pt"): return "p1"
        big = max(abs(v) for v in vals if isinstance(v, (int, float))) if any(isinstance(v, (int, float)) for v in vals) else 0
        return "n0" if big >= 100 else "n2"
    for b in sec["blocks"]:
        if b.get("type") == "table" and str(b.get("caption", "")).startswith("情境驅動"):
            b["head"] = [f"@In!B{h}", f"@In!D{h}", f"@In!E{h}", f"@In!F{h}", f"@In!C{h}"]
            b["rows"] = [[f"@In!B{rr}", f"@In!D{rr}|{fm(rr)}", f"@In!E{rr}|{fm(rr)}", f"@In!F{rr}|{fm(rr)}", f"@In!C{rr}"] for rr in rows]
            return
    raise KeyError("情境驅動表")


scn_table_fix(dr)

# 估值：說明 → 長條圖 → 結論主要項目 → DCF 折現率 → DCF 彙總 → Comparable → 其餘摺疊
V = lambda k: row_of("Val", k, need_num=True)
tpw, dcft, compadj, comp, cons = V("目標價（DCF 與 Comparable 加權）"), V("DCF 目標價（情境加權）"), V("Comparable 目標價（相對倍數調整後）"), V("Comparable 目標價"), V("市場共識目標價")
upw, rec = V("上檔空間：加權目標價"), row_of("Val", "投資建議", exact=True)
dcf12 = row_of("Val", "DCF 情境價值推到 12 個月後")
hc = row_of("README", "展示層輔助：Comparable 目標價（相對倍數調整後）的上檔空間")
c0 = row_of("Val", "0.3 結論")
rk = row_of("Val", "股權成本 ke")                       # 0.1 節：ke、kd_at、w_e、WACC、g、RONIC，接著 DCF 彙總（v0.5.1 起依標籤定位：模組 v2 stub 期多兩列）
r_pvx, r_dcf = row_of("Val", "明確期 FCFF 現值"), row_of("Val", "DCF 目標價（情境加權）")
assert r_pvx == rk + 6, (r_pvx, rk)
def _fm(lab):
    return ("n2" if lab.startswith("評價用") else "p1" if "終值現值占 EV" in lab else "x1" if "隱含終值" in lab else "n0")
fb = row_of("DCF", "營收", start=row_of("DCF", "情境 B（基準）"))   # 情境 B 明細首列（營收）
va = SEC["valuation"]
old = va["blocks"]
dcf_tbl = next(b for b in old if b.get("caption") == "@Val!A60")
peer_tbl = next(b for b in old if b.get("caption") == "@Val!A81")
oldm = next(b for b in old if b.get("caption") == "@Val!A105")
rule_tbl = next(b for b in old if str(b.get("caption", "")).startswith("估值規則"))
hist_tbl = next(b for b in old if str(b.get("summary", "")).startswith("展開：十年 P/E"))
grid_tbl = next(b for b in old if str(b.get("summary", "")).startswith("展開：固定倍數格"))
H5 = ["項目", "數值", "", "", "說明"]
va["blocks"] = [
 {"type": "bars", "title": "目標價與現價", "fmt": "n0", "items": [
   {"label": f"@Val!A{tpw}", "value": f"@Val!B{tpw}", "series": 0},
   {"label": f"@Val!A{dcft}", "value": f"@Val!B{dcft}", "series": 1},
   {"label": f"@Val!A{compadj}", "value": f"@Val!B{compadj}", "series": 2},
   {"label": f"@Val!A{comp}", "value": f"@Val!B{comp}", "series": 1},
   {"label": f"@Val!A{cons}", "value": f"@Val!B{cons}", "series": 0}],
  "ref": {"label": f"@In!B{I['px']}", "value": f"@In!E{I['px']}"}},
 {"type": "table", "caption": "目標價結論（主要項目）", "head": H5, "rows": [
   [f"@Val!A{tpw}", f"@Val!B{tpw}|n0", "", "", f"@Val!E{tpw}"],
   [f"@Val!A{dcft}", f"@Val!B{dcft}|n0", "", "", f"@Val!E{dcft}"],
   [f"@Val!A{compadj}", f"@Val!B{compadj}|n0", "", "", f"@Val!E{compadj}"],
   [f"@In!B{I['px']}", f"@In!E{I['px']}|n0", "", "", ""],
   [f"@Val!A{upw}", f"@Val!B{upw}|sp1", "", "", f"@Val!E{upw}"],
   ["Comparable（相對倍數調整後）的上檔空間", f"@README!B{hc}|sp1", "", "", ""],
   [f"@Val!A{rec}", f"@Val!B{rec}", "", "", f"@Val!E{rec}"],
   [f"@Val!A{dcf12}", f"@Val!B{dcf12}|n0", f"@Val!C{dcf12}|n0", f"@Val!D{dcf12}|n0", ""],
   [f"@Val!A{cons}", f"@Val!B{cons}|n0", "", "", f"@Val!E{cons}"]]},
 {"type": "table", "collapse": True, "summary": "展開：結論的其他項目（市場主流同業法、與共識的差距拆解、價格帶、十年分位、方法）", "head": H5,
  "rowgen": {"sheet": "Val", "rows": f"{c0 + 3},{c0 + 6}-{c0 + 11}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n0", "", "", ""],
             "row_fmt": {str(c0 + 6): {"B": "sp1"}, str(c0 + 7): {"B": "sp1"}, str(c0 + 8): {"B": "sp1"}, str(c0 + 9): {"B": ""}, str(c0 + 10): {"B": "p0", "C": "p0", "D": "p0"}, str(c0 + 11): {"B": ""}}}},
 {"type": "table", "caption": "DCF 折現率（WACC、股權成本、稅後債務成本、RONIC、永續成長率；直接顯示）", "head": ["項目", "數值", "說明"], "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|p2", f"@Val!E{r}"] for r in range(rk, rk + 6)]},
 {"type": "table", "caption": "DCF 彙總（標準版：無槓桿自由現金流 FCFF＋WACC；三情境）", "head": ["項目", "A（樂觀）", "B（基準）", "C（保守）", "說明"], "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|{fm}", f"@Val!C{r}|{fm}" if r != r_dcf else "", f"@Val!D{r}|{fm}" if r != r_dcf else "", f"@Val!E{r}"]
   for r, fm in ((r, _fm(str(WB["Val"].cell(r, 1).value).strip())) for r in range(r_pvx, r_dcf + 1))]},
 {"type": "table", "collapse": True, "summary": "展開：情境 B（基準）FCFF 逐年明細（DCF 頁；新台幣十億元）", "head": ["@DCF!A5", "@DCF!C5", "@DCF!D5", "@DCF!E5", "@DCF!F5", "@DCF!G5", "@DCF!H5"],
  "rowgen": {"sheet": "DCF", "rows": f"{fb}-{fb + 10}", "cols": ["A", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n0", "n0", "n0", "n0", "n0", "n0"],
             "row_fmt": {str(fb + 3): {c: "p1" for c in "CDEFGH"}, str(fb + 5): {c: "p1" for c in "CDEFGH"}, str(fb + 6): {c: "p1" for c in "CDEFGH"}}}},
 {"type": "table", "caption": "Comparable（同業前瞻 P/E × 前瞻 EPS × 相對倍數調整）", "head": ["項目", "數值"], "rows": [
   [f"@Val!A{comp - 2}", f"@Val!B{comp - 2}|x1"], [f"@Val!A{comp - 1}", f"@Val!B{comp - 1}|n2"], [f"@Val!A{comp}", f"@Val!B{comp}|n0"],
   [f"@Val!A{compadj - 1}", f"@Val!B{compadj - 1}|n2"], [f"@Val!A{compadj}", f"@Val!B{compadj}|n0"]]},
 dict(peer_tbl, collapse=True, summary="展開：同業倍數"),
 {"type": "table", "collapse": True, "summary": "展開：相對倍數調整的組成（本公司與同業十年前瞻倍數中位數、夾值）", "head": ["項目", "數值"], "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|n2"] for r in range(compadj - 7, compadj - 1)]},
 oldm,
 {"type": "note", "parts": ["@Val!A27"]},
 dict(rule_tbl, collapse=True, summary="展開：十年 P/E 規則值（舊法參考，不進目標價）"),
 hist_tbl, grid_tbl,
 {"type": "note", "parts": ["@Val!A23"]}]
SEC["recon"]["blocks"][0]["rowgen"]["rows"] = "4-14"
# 敏感度：全部案例表摺疊（龍捲風由轉換器依 |變動| ≥ 10% 分顯示／摺疊）
_sum = {"@Sens!A3": "展開：全部案例（2027 EPS 與 2027 CoWoS 需求 ÷ 產能）", "@Sens!A14": "展開：淨利 ÷ 營業利益的三種設定",
        "@Sens!A22": "展開：目標價參數案例（WACC、β、ERP、永續成長率、RONIC、資本支出、營運資金、權益投資、終值替代法、美元口徑折現、DCF 權重）"}
_ts = row_of("Sens", "WACC ±1pt"); _te = row_of("Sens", "DCF 權重 0.3／0.7")
for b in SEC["sens"]["blocks"]:
    if b["type"] == "table" and b.get("head", [None])[0] == "@Sens!A22":   # v0.5：標準版 DCF 敏感度表（tools/fcff_dcf.write_sens；8 欄）
        b.update(head=[f"@Sens!{c}22" for c in "ABCDEFG"],
                 rowgen={"sheet": "Sens", "rows": f"{_ts}-{_te}", "cols": list("ABCDEFG"), "fmts": ["", "", "n0", "sp1", "n0", "sp1", "n0"]})
    if b["type"] == "table" and b.get("head", [None])[0] in _sum:
        b.update(collapse=True, summary=_sum[b["head"][0]])
SEC["recon"]["blocks"][0]["caption"] = "本模型與公司口徑、共識的對照"

SPEC = {
 "in_skip": [in_skip_token()],
 "in_groups": [
  ["AI 加速器（HPC 的 AI 子列）", [
    ["加速器顆數（markets/memory，唯讀）", [k for k in I if re.match(r"^(u_(bw|rb|ru|hp|amd|tpu|trn|oth)_\d\d|g29_\w+)$", k)]],
    ["情境顆數係數（個股自訂）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["每顆台積電內容、價格調升與其他 AI", [k for k in I if k.startswith("ct_")] + ["px27", "px28", "px29", "ai_oth"]],
    ["AI 實績與公司口徑（對帳用）", ["ai25_share", "ai_cagr"]]]],
  ["CoWoS 先進封裝（瓶頸檢查）", [
    ["每片 CoWoS 封裝數", [k for k in I if k.startswith("pp_")]],
    ["其他 CoWoS 用途", ["cow_oth26", "cow_oth_g"]],
    ["產能與委外", ["cap_end25", "cap_end26", "cap_end27", "cap_end28", "cap_end29", "osat"]],
    ["外部對照（Morgan Stanley、HBM 缺口）", ["ms_dem26", "ms_dem27", "hbm_gap27"]]]],
  ["非 AI 平台（HPC 非 AI、智慧型手機、IoT、車用、DCE、其他）", [
    ["非 AI 基準成長（三情境）", [k for k in I if re.match(r"^na2\d_[ABC]$", k)]],
    ["平台差距（相對非 AI 基準成長）", [k for k in I if k.startswith("dp_")]],
    ["平台營收比重（公司揭露）", [k for k in I if k.startswith("pf2")] + ["hpc_3q"]]]],
  ["全公司（2026 錨點、利潤率、資本支出、股本與目標價規則）", [
    ["2026 實績與第四季指引", ["fx", "rev25", "rev26_1q", "rev26_2q", "rev26_3q", "rev26_4q", "fx26_4q", "ntd26_1q", "ntd26_2q", "ntd26_3q", "ntd26_4q",
                         "gm26_1q", "gm26_2q", "gm26_3q", "gm26_4q", "opm26_1q", "opm26_2q", "opm26_3q", "opm26_4q", "ni26_1q", "ni26_2q", "ni26_3q",
                         "eps26_1to3q", "g26_guid", "rev26", "ntd26", "fx_26", "gm26", "opm26", "opex_26", "ni_op_26"]],
    ["毛利率、營業費用與淨利轉換（FY2027 起）", [k for k in I if re.match(r"^gm2[789]_[ABC]$", k)] + ["opex", "ni_op"]],
    ["資本支出與自由現金流", ["capex26", "capex27", "capex28", "capex29", "da_rev", "cx_rev"]],
    ["股價與股數", ["px", "shares"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "rf_us", "beta_u", "beta_u_med", "beta_de", "beta_du", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["非控制權益（上市子公司市值、其餘帳面）", ["nci_bk", "nci_bk_l", "own_guc", "sh_guc", "eq_guc", "px_guc", "nci_mv_guc", "own_vis", "sh_vis", "eq_vis", "px_vis", "nci_mv_vis"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["收斂法（FY2030–FY2032）", ["tp_g", "nim10", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["固定倍數與市場共識（對帳用）", ["pe_1", "pe_2", "pe_3", "cons_rev27", "cons_eps26", "cons_eps27", "cons_tp", "cons_feps27", "cons_feps28"]],
    ["Wattonomics 錨定（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]]],
 "blocks_to": {"CoWoS 供需": "CoWoS 先進封裝（瓶頸檢查）", "@CoWoS!F17": "CoWoS 先進封裝（瓶頸檢查）", "標記沿用": "全公司（2026 錨點、利潤率、資本支出、股本與目標價規則）"},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "目標價結論": ["目標價（DCF", "DCF 目標價", "Comparable 目標價（相對", "股價", "上檔空間：加權", "Comparable（相對倍數調整後）的上檔空間", "投資建議", "DCF 情境價值推到", "市場共識"],
          "本模型與公司口徑、共識的對照": ["2026 AI 相關營收", "2026 營收 vs 公司", "2027 EPS（A）vs 共識（", "2028 EPS（A）vs 共識", "2027 CoWoS 需求"]},
 "n_recon": 4,
}

cfg = hero_rule_a(run(SPEC))
# 三層協調 C4（規則 v1.4；tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(cfg, X)


# ================= v0.5.5（2026-11-15）：個股 HTML 版面規格 v1.3（樣板 NVIDIA、泰瑞達、日月光；docs/plan/個股HTML版面規格.md「v1.3 規格」）＋三層協調 C0 =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導；CoWoS 瓶頸檢查併入）→ 對帳（含三層協調 C0）→ 資料狀態；台積電無個別模組。
# 數字與白話文字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末相對基準表、Coord 頁；tsmc/v055_post.py 產生）。列號依標籤找。
def _v13():
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RVW, CW, SEW = WB["README"], WB["PL"], WB["RevDCF"], WB["Coord"], WB["Sens"]
    r13 = row_of("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert "v0.5.5（" in str(RD[f"A{RM['ver']}"].value), RD[f"A{RM['ver']}"].value   # v0.5.6：本版要點列前置 px1113 說明，v0.5.5 說明接在其後
    cfg["layout"] = "v1.3"
    cr = lambda p, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
    # ---- 橫幅：≤200 字白話與對帳中揭露；規則與歸因移到估值章 ----
    old_note = list(cfg["hero"]["note"])
    assert old_note[-2] == " " and str(old_note[-1]).startswith("@Coord!B"), old_note
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"] + old_note[-2:]
    for t in cfg["hero"]["tiles"]:
        t[0] = [t[0], "（新台幣）"] if isinstance(t[0], str) else list(t[0]) + ["（新台幣）"]
    # ---- 1. 命題：主圖＝分部營收推導（情境 B；AI 相關營收與非 AI 平台堆疊）----
    sg = row_of("PL", "分部營收推導（情境 B")
    hS = sg + 1
    ex = lambda label, start=sg: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == label)
    r_ai, r_hn, r_sp, r_iot, r_au, r_dce, r_oth, r_rs, r_sh, r_u = (ex(t) for t in ("AI 相關營收（HPC 子列）", "HPC：非 AI", "智慧型手機", "物聯網", "車用", "消費性電子（DCE）", "其他", "營收（美元）", "AI 相關營收 ÷ 營收", "AI 加速器顆數（情境 B，含情境係數）"))
    th = S["thesis"]["blocks"]
    chart1 = {"type": "theme", "title": "營收分部（情境 B，十億美元）與 AI 相關營收占比：FY2026E–FY2029", "x": f"@PL!F{hS}:I{hS}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.35,
              "total": f"@PL!F{r_rs}:I{r_rs}",
              "bars": [{"name": [f"@PL!A{r_ai}"], "y": f"@PL!F{r_ai}:I{r_ai}", "color": "#0b7a80"}, {"name": [f"@PL!A{r_hn}"], "y": f"@PL!F{r_hn}:I{r_hn}", "color": "#6fc7cb"},
                       {"name": [f"@PL!A{r_sp}"], "y": f"@PL!F{r_sp}:I{r_sp}", "color": "#c29a2b"}, {"name": [f"@PL!A{r_iot}"], "y": f"@PL!F{r_iot}:I{r_iot}", "color": "#8a8c8c"},
                       {"name": [f"@PL!A{r_au}"], "y": f"@PL!F{r_au}:I{r_au}", "color": "#b4472f"}, {"name": [f"@PL!A{r_dce}"], "y": f"@PL!F{r_dce}:I{r_dce}", "color": "#6f6bae"},
                       {"name": [f"@PL!A{r_oth}"], "y": f"@PL!F{r_oth}:I{r_oth}", "color": "#c9ccd0"}],
              "lines": [{"name": [f"@PL!A{r_sh}"], "y": f"@PL!F{r_sh}:I{r_sh}", "axis": "r", "color": "#3b3838", "labels": True}],
              "caption": "堆疊＝情境 B（基準）的 AI 相關營收（HPC 子列）與五個非 AI 平台；頂端數字＝合併營收（美元）；折線＝AI 相關營收占營收（右軸）。FY2026 為前三季實績加第四季指引的錨點；AI 相關營收由層級顆數 × 每顆台積電內容推導。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    S["thesis"]["blocks"] = [chart1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]}] + th
    # ---- 2. 估值：四個目標價圖＋表置頂（G）；章說明、評等規則、對帳中兩讀數區其後 ----
    va = S["valuation"]["blocks"]
    g0 = row_of("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = g0 + 7
    bars4 = {"type": "bars", "title": "四個目標價與現價（新台幣；豎線＝現價）", "fmt": "n0",
             "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                       ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
             "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
    tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n0", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_px}", f"@Val!B{r_px}|n0", "", f"@Val!D{r_px}"]]}
    rules = {"type": "note", "parts": old_note[:-2]}
    h_i = next(i for i, b in enumerate(va) if b.get("type") == "h3" and str(b.get("text", "")).startswith("對帳中"))
    coord_blocks = va[h_i:h_i + 4]
    assert [b["type"] for b in coord_blocks] == ["h3", "note", "table", "bars"], [b["type"] for b in coord_blocks]
    old_bars = next(b for b in va if b.get("type") == "bars" and b.get("title") == "目標價與現價")
    rest = [b for b in va if b not in coord_blocks and b is not old_bars]
    top = [bars4]
    if S["valuation"].get("intro"):
        top.append({"type": "note", "parts": [S["valuation"].pop("intro")]})
    S["valuation"]["blocks"] = top + [tbl4, rules] + coord_blocks + rest
    # ---- 3. 敏感度：加權目標價相對基準的低案、高案主圖置頂（H）----
    t0 = row_of("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SEW.cell(r2 + 1, 1).value and not str(SEW.cell(r2 + 1, 1).value).startswith("快照狀態"):
        r2 += 1
    snap = row_of("Sens", "快照狀態")
    tor = []
    for col, side in (("B", "低"), ("C", "高")):
        big = [r for r in range(r1, r2 + 1) if isinstance(SEW[f"{col}{r}"].value, (int, float)) and abs(SEW[f"{col}{r}"].value) >= 0.1]
        small = [r for r in range(r1, r2 + 1) if isinstance(SEW[f"{col}{r}"].value, (int, float)) and 0.0005 <= abs(SEW[f"{col}{r}"].value) < 0.1]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        tor.append((dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))),
                    dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                         summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例（幾乎為零者見全部案例表）")))
    all_tbl = {"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
               "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1, r2 + 1)]}
    others = []
    for b in S["sens"]["blocks"]:
        if b.get("type") == "tornado":
            b = dict(b, collapse=True, summary=("展開：對照：2027 EPS 相對基準的變動未達兩位數百分比的案例（情境 A）" if b.get("collapse") else "展開：對照：2027 EPS 相對基準的變動達兩位數百分比的案例（情境 A，一次動一個輸入）"))
        elif b.get("type") == "table" and not b.get("collapse"):
            b = dict(b, collapse=True, summary="展開：" + (str(b.get("caption") or "表")[:40]))
        others.append(b)
    intro = S["sens"].pop("intro", None)
    lead = [tor[0][0], tor[1][0], {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入；變動達兩位數百分比者直接顯示，其餘摺疊）。業務主因為整本重算快照，快照狀態：", f"@Sens!B{snap}", "；目標價參數取 Sens 頁「目標價參數敏感度」表。下方摺疊為全部案例與 EPS 敏感度對照。"]}]
    if intro: lead.append({"type": "note", "parts": [intro], "collapse": True, "summary": "展開：EPS 敏感度的說明"})
    S["sens"]["blocks"] = lead + [tor[0][1], tor[1][1], all_tbl] + others
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RVW.max_row + 1) if str(RVW.cell(r, 1).value or "").startswith(p))
    ra, rb_, rbb, rc, rcmp, rsum, yh = rv("A. DCF"), rv("(a)"), rv("(b)"), rv("(c)"), rv("B. Comparable"), rv("C. 摘要"), rv("項目（情境 B）")
    m1, m0 = rv("現價隱含的營業利益率", rb_), rv("模型：基準的等效營業利益率", rb_)
    s1, rv1, rv0, g1, g0_ = rv("現價隱含的放大倍數", rbb), rv("現價隱含的 FY2028 營收（", rbb), rv("模型：FY2028 營收（", rbb), rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
    w1, w0 = rv("現價隱含的 WACC", rc), rv("模型 WACC", rc)
    kvr = [rv(p, ra) for p in ("現價（", "支持現價所需的評價日", "淨現金滾到評價日", "評價用稀釋股數", "支持現價所需的企業價值", "模型企業價值", "EV 需放大", "模型：情境 B 評價日")]
    kvf = ["n2", "n2", "n1", "n2", "n1", "n1", "n2", "n2"]
    cm = [rv(p, rcmp) for p in ("12 個月後前瞻 EPS", "只看 Comparable：支持現價所需的前瞻", "只看 Comparable：所需的相對", "加權（DCF 不變）", "同業前瞻本益比中位數")]
    yrr = [rv(p, yh) for p in ("營收（", "營業利益 EBIT", "營業利益率", "稅率", "NOPAT＋折舊", "營運資金增加", "現值（")]
    yfm = ["n1", "n1", "p1", "p1", "n1", "n1", "n1"]
    S["revdcf"] = {
        "id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
        "blocks": [
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
            {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
             "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                      ["現金流放大倍數（FY2028 起各年同乘；模型基準為一倍）", f"@RevDCF!B{s1}|n2", "—"],
                      ["FY2028 營收（十億新台幣）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                      ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0_}|p1"],
                      ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
            {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(kvr, kvf)]},
            {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
             "rows": [[f"@RevDCF!A{cm[0]}", "—", f"@RevDCF!C{cm[0]}|n2", "—"]]
                     + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!C{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!D{r}|sp1"] for k, r in enumerate(cm[1:4])]
                     + [[f"@RevDCF!A{cm[4]}", "—", f"@RevDCF!C{cm[4]}|n1", "—"]]},
            {"type": "table", "caption": f"@RevDCF!A{yh}", "collapse": True, "summary": "展開：情境 B 各年營收、營業利益與現值（反解的計算輸入；十億新台幣）",
             "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"],
             "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{f}" for c in "BCDEFG"] for r, f in zip(yrr, yfm)]}]}
    # ---- 5. 損益（標準損益表）----
    p0 = row_of("PL", "標準損益表（v0.5.5")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（含其他營業損益）", "營業利益", "營業利益率", "其他收入（淨額；營業外）", "稅前利益", "所得稅",
            "稅後淨利（合併，含非控制權益）", "淨利率", "股數（十億股；歷史為稀釋加權平均）", "EPS（新台幣；歷史為公司公告稀釋 EPS）"]
    fmt = lambda u: {"%": "p1", "NT$": "n2", "十億股": "n2"}.get(str(u).strip(), "n1")
    PC = "CDEFGHIJKL"
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in PC]
    r_ratio = plr("比率"); r_chk = next(r for r in range(r_ratio, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("核對（應為 0"))
    r_ycav = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("年度與對帳中"))
    head = [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC]
    is_tbl = {"type": "table", "caption": "標準損益表（十億新台幣；EPS 為新台幣）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B（基準）；IFRS 合併（GAAP 口徑）", "head": head,
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]}
    chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": head, "collapse": True, "summary": "展開：核對列（歷史捨入差、預測期對 PL 情境 B 與延伸表、FY2030E–FY2032E 的推法差額）",
               "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{'p1' if '占營收' in str(PLW.cell(r, 1).value) else 'n3'}" for c in PC] for r in range(r_chk + 1, r_ycav) if PLW.cell(r, 1).value]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（十億新台幣）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:L{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:L{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:L{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:L{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸；IFRS 合併，淨利含非控制權益）。FY2026E 為前三季實績加第四季指引的錨點；FY2030E–FY2032E 為收斂法延伸期（營業利益由淨利還原）。"}
    old_pl = S["pl"]["blocks"]
    legacy = []
    for b in old_pl:
        if b["type"] == "pl":
            legacy.append(dict(b, collapse=True, summary="展開：三情境損益摘要與折線（美元口徑）"))
        elif b["type"] in ("note", "kv", "table"):
            legacy.append(dict(b, collapse=True, summary="展開：" + ("延伸期（FY2030 起）說明" if b["type"] == "note" else "三情境錨點核對")))
    pl_intro = S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, {"type": "note", "parts": [f"@PL!A{r_ycav}"]}, chk_tbl]
                   + ([{"type": "note", "parts": [pl_intro], "collapse": True, "summary": "展開：三情境損益摘要的期間說明"}] if pl_intro else []) + legacy)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    sf = {"百萬顆": "n2", "%": "p1", "US$B": "n2"}
    seg_rows = [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{sf.get(str(PLW.cell(r, 2).value), 'n2')}" for c in "FGHI"] for r in range(r_u, r_sh + 1)]
    seg_tbl = {"type": "table", "caption": f"@PL!A{sg}", "head": [f"@PL!A{hS}", f"@PL!B{hS}"] + [f"@PL!{c}{hS}" for c in "FGHI"], "rows": seg_rows}
    chart2 = {"type": "theme", "title": "AI 加速器顆數（情境 B，百萬顆）與 AI 相關營收（十億美元，右軸）：FY2026E–FY2029", "x": f"@PL!F{hS}:I{hS}", "fmt": "n1", "fmt_r": "n1", "stack": False, "pad": 1.3,
              "bars": [{"name": [f"@PL!A{r_u}"], "y": f"@PL!F{r_u}:I{r_u}", "color": "#6fc7cb"}],
              "lines": [{"name": [f"@PL!A{r_ai}"], "y": f"@PL!F{r_ai}:I{r_ai}", "axis": "r", "color": "#0b7a80", "labels": True}],
              "caption": "長條＝各平台 AI 加速器顆數加總（Siliconomics 現行基準；CY2028 起乘情境 B 係數）；折線＝AI 相關營收（顆數 × 每顆台積電內容 × 價格調升，再加其他 AI 加成）。"}
    dr = S["drivers"]["blocks"]
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, chart2, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + dr
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    S["drivers"].pop("intro", None)
    # ---- 7. 對帳：三層協調 C0（Coord 頁）----
    c1, c2, c3 = cr("1. 讀哪一層"), cr("2. C0 讀取層級核對"), cr("3. 對帳中年份")
    h1 = c1 + 1; h2 = next(r for r in range(c2, CW.max_row + 1) if CW.cell(r, 1).value == "項目")
    cnt = cr("對帳中項目數（讀取層級核對")
    ucf = lambda r: {"%": "p1", "x": "n3", "百萬顆": "n3", "NT$/US$": "n2"}.get(str(CW.cell(r, 2).value), "n1")
    c0_tbl = {"type": "table", "caption": "讀取層級核對：個股讀入的層級值 vs 層級值（同源，不判定、不計入對帳中；有檢驗力的對帳見上表「公司實績 vs 層級隱含值」）", "collapse": True, "summary": "展開：讀取層級核對（同源，不判定）",
              "head": [f"@Coord!{c}{h2}" for c in "ABCDEFG"],
              "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|{ucf(r)}", f"@Coord!E{r}|{ucf(r)}", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in range(h2 + 1, cnt) if CW.cell(r, 4).value is not None]}
    c0_src = {"type": "table", "caption": "層級名稱、版本、口徑與證據等級（每列的來源；展開看）", "collapse": True, "summary": "展開：每列的層級名稱、版本、口徑與證據等級",
              "head": ["項目", "年度", "層級名稱與版本／口徑／證據等級／備註"], "rows": [[f"@Coord!A{r}", f"@Coord!C{r}", f"@Coord!H{r}"] for r in range(h2 + 1, cnt) if CW.cell(r, 8).value]}
    rd_tbl = {"type": "table", "caption": f"@Coord!A{c1}", "head": [f"@Coord!{c}{h1}" for c in "ABCDE"], "collapse": True, "summary": "展開：本模型讀哪一層、哪一版（下游只註明版本、不自行重估）",
              "rows": [[f"@Coord!A{r}", f"@Coord!B{r}" + ("|n3" if isinstance(CW.cell(r, 2).value, (int, float)) and not isinstance(CW.cell(r, 2).value, bool) else ""), f"@Coord!C{r}", f"@Coord!D{r}", f"@Coord!E{r}"]
                       for r in range(h1 + 1, c2 - 1) if CW.cell(r, 1).value]}
    rf0 = cr("層級之間"); rfh = next(r for r in range(rf0, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); rfe = next(r for r in range(rfh, CW.max_row + 1) if not CW.cell(r + 1, 1).value)
    ref_blocks = [{"type": "h3", "text": f"@Coord!A{rf0}"}, {"type": "note", "parts": [f"@Coord!A{rf0 + 1}"]},
                  {"type": "table", "head": [f"@Coord!{c}{rfh}" for c in "ABCDEFG"], "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n3", f"@Coord!E{r}|n3", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in range(rfh + 1, rfe + 1)]}]
    a4 = cr("4. C0（實績）"); ih = next(r for r in range(a4, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); mh = next(r for r in range(ih, CW.max_row + 1) if CW.cell(r, 1).value == "讀數")
    jr0 = cr("C0（實績）判定：合併營收"); a4e = cr("對帳中項目數（C0 實績表", mh)
    ifm = lambda r: "p0" if str(CW.cell(r, 2).value) == "%" and "占營收" in str(CW.cell(r, 1).value) else "p1" if str(CW.cell(r, 2).value) == "%" else "n2" if str(CW.cell(r, 2).value) == "US$B" else "n1"
    in_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|{ifm(r)}", f"@Coord!E{r}", f"@Coord!H{r}"] for r in range(ih + 1, mh - 1) if CW.cell(r, 1).value]
    rd2 = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", f"@Coord!F{r}|p1", f"@Coord!G{r}"] for r in range(mh + 1, jr0)]
    jr_rows = [[f"@Coord!A{r}", f"@Coord!G{r}", f"@Coord!H{r}"] for r in range(jr0, a4e)]
    act_blocks = [{"type": "h3", "text": "公司實績對層級隱含值（實績對帳）"}, {"type": "note", "parts": [f"@Coord!A{a4 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 2}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 3}"]},
                  {"type": "table", "caption": "HPC 平台營收（公司實績，占比 × 營收換算）與層級隱含 AI 相關營收（參考；同源校準，不判定）", "head": [f"@Coord!{c}{mh}" for c in "ABCDEFG"], "rows": rd2},
                  {"type": "table", "caption": "公司實績對層級的可比性判定", "head": ["項目", "判定", "結論與缺漏揭露"], "rows": jr_rows},
                  {"type": "note", "parts": [f"@Coord!A{a4e + 1}"]}, {"type": "note", "parts": ["對帳中項目數：", f"@Coord!G{a4e}|n0", "。"]},
                  {"type": "table", "caption": "實績與層級隱含值的輸入、等級與出處（展開看）", "collapse": True, "summary": "展開：公司實績、層級值、換算參數與出處", "head": [f"@Coord!{c}{ih}" for c in "ABCDE"] + [f"@Coord!H{ih}"], "rows": in_rows}]
    c0_note = [{"type": "h3", "text": "讀取層級核對（同源，不判定）"}, {"type": "note", "parts": [f"@Coord!A{c2}"]}, {"type": "note", "parts": [f"@Coord!A{c2 + 1}"]}, c0_tbl]
    c3_notes = [{"type": "note", "parts": [f"@Coord!A{c3 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{c3 + 2}"]}]
    S["recon"]["blocks"] = [{"type": "h3", "text": "三層協調：由下往上對帳（公司實績對各層）"}] + act_blocks + c0_note + ref_blocks + c3_notes + [c0_src, rd_tbl] + [{"type": "h3", "text": "本模型與公司口徑、共識的對照"}] + S["recon"]["blocks"]
    S["recon"]["title"] = "對帳：三層協調、公司實績與共識對照"
    # ---- 8. 資料狀態：本版要點 ----
    S["status"]["blocks"][0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    cfg["sections"] = [S[k] for k in order]


_v13()
dump(cfg)

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
fb = row_of("DCF", "營收", start=row_of("DCF", "情境 B（平平）"))   # 情境 B 明細首列（營收）
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
 {"type": "table", "caption": "DCF 彙總（標準版：無槓桿自由現金流 FCFF＋WACC；三情境）", "head": ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"], "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|{fm}", f"@Val!C{r}|{fm}" if r != r_dcf else "", f"@Val!D{r}|{fm}" if r != r_dcf else "", f"@Val!E{r}"]
   for r, fm in ((r, _fm(str(WB["Val"].cell(r, 1).value).strip())) for r in range(r_pvx, r_dcf + 1))]},
 {"type": "table", "collapse": True, "summary": "展開：情境 B（平平）FCFF 逐年明細（DCF 頁；新台幣十億元）", "head": ["@DCF!A5", "@DCF!C5", "@DCF!D5", "@DCF!E5", "@DCF!F5", "@DCF!G5", "@DCF!H5"],
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
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["收斂法（FY2030–FY2032）", ["tp_g", "nim10", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["固定倍數與市場共識（對帳用）", ["pe_1", "pe_2", "pe_3", "cons_rev27", "cons_eps26", "cons_eps27", "cons_tp", "cons_feps27", "cons_feps28"]],
    ["Wattonomics 錨定（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]]],
 "blocks_to": {"CoWoS 供需": "CoWoS 先進封裝（瓶頸檢查）", "@CoWoS!F17": "CoWoS 先進封裝（瓶頸檢查）", "標記沿用": "全公司（2026 錨點、利潤率、資本支出、股本與目標價規則）"},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "目標價結論": ["目標價（DCF", "DCF 目標價", "Comparable 目標價（相對", "股價", "上檔空間：加權", "Comparable（相對倍數調整後）的上檔空間", "投資建議", "DCF 情境價值推到", "市場共識"],
          "本模型與公司口徑、共識的對照": ["2026 AI 相關營收", "2026 營收 vs 公司", "2027 EPS（A）vs 共識（", "2028 EPS（A）vs 共識", "2027 CoWoS 需求"]},
 "n_recon": 4,
}

dump(hero_rule_a(run(SPEC)))

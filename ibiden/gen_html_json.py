# 由 html_src.json（v0.4 以前的 HTML 設定，內容不動）套用「個股 HTML 版面規格 v1.0」（docs/plan/個股HTML版面規格.md），產生 ibiden/html.json。
# 用法：python3 ibiden/gen_html_json.py [成品 xlsx]（省略時取 ibiden/dist/ 內唯一的 xlsx）；勿手改 html.json。
# 轉換器 tools/company_html/relayout.py 之外，本檔只做 Ibiden 專屬的前處理（估值章重排）與後處理（橫幅、龍捲風標題、附錄）。
import copy, glob, json, os, re, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
from relayout import relayout  # noqa: E402

X = sys.argv[1] if len(sys.argv) > 1 else (lambda f: f[0] if len(f) == 1 else sys.exit(f"dist 內應剛好一個 xlsx：{f}"))(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))
OUT = os.path.join(HERE, "html.json")
cfg = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)


def find(sh, prefix, start=1, col=1):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        if str(ws.cell(r, col).value or "").strip().startswith(prefix):
            return r
    raise KeyError((sh, prefix))


def sec(cid):
    return next(s for s in cfg["sections"] if s["id"] == cid)


V = lambda r, f="": f"@Val!B{r}" + (f"|{f}" if f else "")
VA = lambda r: f"@Val!A{r}"
INK = {str(wb["In"].cell(r, 1).value): r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}
# ---- v0.5：標準版 DCF（FCFF＋WACC）的估值區塊——折現率表緊接 DCF 彙總，各情境 FCFF 明細摺疊（docs/plan/個股DCF標準規格.md、個股HTML版面規格 v1.0）----
def fcff_blocks(pxf="n0", amt="n1"):
    Vw, Dw = wb["Val"], wb["DCF"]
    rk = find("Val", "股權成本 ke")
    rate = {"type": "table", "caption": "DCF 折現率：WACC、股權成本、稅後債務成本、RONIC、永續成長率（直接顯示）", "head": ["項目", "數值", "說明"],
            "rows": [[VA(r), V(r, "p1" if r == rk + 2 else "p2"), f"@Val!E{r}"] for r in range(rk, rk + 6)]}
    rows = []
    r_end = find("Val", "DCF 目標價（情境加權）", rk)   # v0.5.1：依標籤定位（模組 v2 stub 期在「＋淨現金與投資」後多兩列）
    for r in range(rk + 6, r_end):
        lab = str(Vw.cell(r, 1).value)
        f = ("n3" if lab.startswith("評價用") else "p1" if "占 EV" in lab else "x1" if "EV／EBITDA" in lab
             else pxf if ("每股" in lab or "推到" in lab) else amt)
        rows.append([VA(r)] + [f"@Val!{c}{r}|{f}" for c in "BCD"])
    rows.append([VA(r_end), V(r_end, pxf), "", ""])
    summ = {"type": "table", "caption": "DCF 彙總：明確期現值、終值現值、企業價值、淨現金與每股價值", "head": [f"@Val!{c}{rk - 1}" for c in "ABCD"], "rows": rows}
    hd = next(r for r in range(1, Dw.max_row + 1) if Dw.cell(r, 1).value == "項目")
    ncol = next(c for c in range(3, 30) if not Dw.cell(hd, c).value) - 3
    cols = [chr(ord("C") + i) for i in range(ncol)]
    det = []
    for r in range(hd + 1, Dw.max_row + 1):
        if str(Dw.cell(r, 1).value or "").startswith("情境 ") and Dw.cell(r + 1, 1).value == "營收" and len(det) < 3:
            rf = {str(r + k): {c: f for c in cols} for k, f in ((4, "p1"), (6, "p1"), (7, "p1"), (12, "n3"))}
            det.append({"type": "table", "caption": f"@DCF!A{r}", "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細",
                        "head": [f"@DCF!A{hd}"] + [f"@DCF!{c}{hd}" for c in cols],
                        "rowgen": {"sheet": "DCF", "rows": f"{r + 1}-{r + 13}", "cols": ["A"] + cols, "fmts": [""] + [amt] * ncol, "row_fmt": rf}})
    return [rate, summ] + det


FCFF_MAIN = {"DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
             "DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"]}
FCFF_KEYS = {"rate": ["rf", "rf_us", "beta_u", "beta_de", "beta_du", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"],
             "bs": ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"],
             "fcf": ["cxs_26", "cxs_3y", "cxs_add", "cxs_27", "cxs_28", "das_26", "life_s", "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]}


def sens_tp_block(pxf="n0"):
    t0 = find("Sens", "目標價參數敏感度")
    return {"type": "table", "caption": f"@Sens!A{t0}", "head": [f"@Sens!{c}{t0 + 1}" for c in "ABCDEFG"],
            "rowgen": {"sheet": "Sens", "rows": f"{t0 + 2}-{find('Sens', '檢查：閉式公式', t0) - 1}", "cols": list("ABCDEFG"), "fmts": ["", "", pxf, "sp1", pxf, "sp1", pxf]}}

v_tp, v_dcf, v_cmp, v_cmp0 = find("Val", "目標價（DCF 與 Comparable 加權）"), find("Val", "DCF 目標價（情境加權）"), find("Val", "Comparable 目標價（相對倍數調整後）"), find("Val", "Comparable 目標價（市場主流同業法）")
v_px, v_up, v_up0, v_cons = find("Val", "現價（"), find("Val", "上檔空間：加權目標價"), find("Val", "上檔空間：Comparable 目標價"), find("Val", "市場共識目標價")
v_rec, v_grade, v_band, v_hist, v_meth = find("Val", "投資建議"), find("Val", "上檔空間級距"), find("Val", "價格帶"), find("Val", "歷史倍數檢查"), find("Val", "目標價方法")
v_eps_f, v_eps_m = find("Val", "與市場共識目標價的差距：EPS"), find("Val", "與市場共識目標價的差距：倍數")
v_s12 = find("Val", "DCF 情境價值推到")
rd_upc = find("README", "上檔空間：Comparable 目標價（相對倍數調整後）", find("README", "HTML 輔助數字"))

# ---- 前處理一：估值章依規格重排（長條圖 → 結論主要項目 → 折現率 → DCF → Comparable → 其他）----
va = sec("valuation")
old = va["blocks"]
cap = lambda b: b.get("caption", "")
dcf_tbl = next(b for b in old if cap(b) == "@Val!A62" or (cap(b).startswith("@Val!A") and str(wb["Val"][cap(b)[5:]].value).startswith("0.1 DCF")))
cmp_tbl = next(b for b in old if cap(b).startswith("@Val!A") and str(wb["Val"][cap(b)[5:]].value).startswith("0.2 Comparable"))
rm_kv = next(b for b in old if b["type"] == "kv")
rest = [b for b in old if b is not dcf_tbl and b is not cmp_tbl and b is not rm_kv
        and cap(b) != "目標價（DCF 與 Comparable 加權）"
        and not (b["type"] == "h3" and b.get("text") in ("重點觀察",) or b["type"] == "h3" and str(b.get("text", "")).startswith("相對倍數調整"))
        and not (b["type"] == "list" and any(str(i).startswith("@README!") for i in b.get("items", [])))
        and not (b["type"] == "note" and b.get("parts", [""])[0] in (VA(v_grade), VA(v_rec)))]
va["blocks"] = [
    {"type": "bars", "title": "目標價與現價", "fmt": "n0", "items": [
        {"label": VA(v_dcf), "value": V(v_dcf), "series": 0},
        {"label": VA(v_cmp), "value": V(v_cmp), "series": 1},
        {"label": VA(v_tp), "value": V(v_tp), "series": 2},
        {"label": VA(v_cons), "value": V(v_cons), "series": 0}],
     "ref": {"label": VA(v_px), "value": V(v_px)}},
    {"type": "table", "caption": "目標價結論", "head": ["項目", "數值"], "rows": [
        [VA(v_tp), V(v_tp, "n0")], [VA(v_dcf), V(v_dcf, "n0")], [VA(v_cmp), V(v_cmp, "n0")], [VA(v_px), V(v_px, "n0")],
        [VA(v_up), V(v_up, "sp1")], [f"@README!A{rd_upc}", f"@README!B{rd_upc}|sp1"], [VA(v_rec), V(v_rec)], [VA(v_band), V(v_band)], [VA(v_cons), V(v_cons, "n0")],
        [VA(v_cmp0), V(v_cmp0, "n0")], [VA(v_up0), V(v_up0, "sp1")], [VA(v_grade), V(v_grade)], [VA(v_eps_f), V(v_eps_f, "sp1")], [VA(v_eps_m), V(v_eps_m, "sp1")],
        [VA(v_hist), [V(v_hist, "p0"), "／", f"@Val!C{v_hist}|p0", "／", f"@Val!D{v_hist}|p0"]], [VA(v_meth), V(v_meth)]]},
    *fcff_blocks(),
    cmp_tbl,
    {"type": "table", "caption": "相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司相對同業的長期倍數比）", "head": ["項目", "數值"], "rows": [[k, v] for k, v in rm_kv["items"]]},
] + rest

# ---- 前處理二：敏感度章「目標價參數」表的列號改為 Sens 頁實際位置（舊設定的列號已位移）----
se = sec("sens")
for i_, b in enumerate(se["blocks"]):
    if b["type"] == "table" and cap(b).startswith("@Sens!A") and str(b.get("head", [""])[0]).startswith("@Sens!A") and b.get("rowgen", {}).get("cols") == list("ABCDEF"):
        se["blocks"][i_] = sens_tp_block()   # v0.5：目標價參數表改為標準版 DCF 案例（tools/fcff_dcf.write_sens）

# ---- 轉換：版面規格 v1.0 ----
scn_tok = re.split(r"[.\s]", str(wb["In"].cell(find("In", "情境驅動對照"), 1).value))[0]
SPEC = {
 "in_skip": [scn_tok],
 "in_groups": [
  ["電子事業：AI 伺服器載板", [
    ["加速器顆數（讀網通市場模型）", [k for k in INK if re.match(r"^(u_|g29_|itc_)", k)]],
    ["情境顆數係數（個股自訂）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["份額與每顆載板金額", ["w_ot", "ai_e", "g_asp", "rel_27", "rel_28", "rel_29"]],
    ["短缺溢價（價格指數）", ["pp_27", "pp_28", "pp_29"]],
    ["Wattonomics 錨定係數（只用於敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  ["電子事業：非 AI 部分與電子事業起始值", [
    ["非 AI 部分（通用伺服器、PC 載板等）", ["m_ne", "g_ne"]],
    ["電子事業的公司預測與實績（校準起點）", ["f_e", "f_eoi", "a_e", "a_eoi"]]]],
  ["陶瓷與其他事業", [
    ["營收成長與營益率", ["g_cer", "m_cer", "g_oth", "m_oth"]],
    ["公司預測與實績（校準起點）", ["f_cer", "f_ceroi", "f_oth", "f_othoi", "a_cer", "a_ceroi", "a_oth", "a_othoi"]]]],
  ["全公司：股價、折現率、目標價規則、稅與現金流", [
    ["全公司的公司預測與實績（校準起點）", ["f_rev", "f_oi", "f_ni", "f_eps", "a_rev", "a_oi", "a_gp", "a_ni", "a_eps"]],
    ["淨利、股數與匯率；標準損益表的營業費用彈性", ["ni_oi", "sh_g", "fx", "opx_el"]],
    ["股價、市場倍數與公司目標", ["px", "pe_fwd_sa", "pe_ttm", "pe_fair_andy", "tgt30_rev", "tgt30_oi", "cons_tp", "cons_feps"]],
    ["折現率（標準版 DCF：無風險利率、β、股權風險溢酬、WACC、永續成長率、RONIC）", FCFF_KEYS["rate"]],
    ["資本結構與淨現金（WACC 權重、EV 到每股價值）", FCFF_KEYS["bs"]],
    ["自由現金流假設（DCF：擴產排程、折舊、資本支出、營運資金；da_rev、cx_rev 只用於 PL 頁自由現金流列）", FCFF_KEYS["fcf"] + ["da_rev", "cx_rev"]],
    ["目標價規則、情境權重與延伸期收斂", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC", "tp_g", "fade_gA", "fade_gB", "fade_gC", "nim10_f"]]]]],
 "blocks_to": {"單一參數": "電子事業：AI 伺服器載板", "Ibiden 在 AI 加速器載板": "電子事業：AI 伺服器載板", "@Units!A": "電子事業：AI 伺服器載板"},
 "main": {
  "目標價結論": ["目標價（DCF 與 Comparable 加權）", "DCF 目標價", "Comparable 目標價（相對倍數調整後）", "現價", "上檔空間：加權目標價", "上檔空間：Comparable 目標價（相對倍數調整後）", "投資建議", "價格帶", "市場共識目標價"],
  **FCFF_MAIN,
  "0.2 Comparable": ["同業前瞻 P/E 中位數", "12 個月後的前瞻 EPS", "Comparable 目標價"],
  "相對倍數調整": ["本公司 10 年中位數", "同業 10 年中位數的中位數", "相對倍數調整係數", "Comparable 目標價（相對倍數調整後）"],
  "現價與市場隱含倍數": ["現價", "現價 ÷ FY2027 EPS", "現價 ÷ FY2028 EPS", "前瞻 P/E"],
 },
}
if sec("drivers").get("intro"):   # 驅動因子章的說明文字（轉換器轉成 note）放第一塊
    SPEC["blocks_to"][sec("drivers")["intro"]] = "電子事業：AI 伺服器載板"
# blocks_to 的 note 以 parts[0] 開頭比對：Units 頁註解
for b in sec("drivers")["blocks"]:
    if b["type"] == "note" and str(b.get("parts", [""])[0]).startswith("@Units!A"):
        SPEC["blocks_to"][b["parts"][0]] = "電子事業：AI 伺服器載板"
cfg = relayout(cfg, X, SPEC)

# ---- 後處理一：橫幅中間格＝只看 Comparable（市場主流同業法、未經相對倍數調整；工作單補充規則 A），附它自己的上檔空間；
#      進加權的是相對倍數調整後的值，寫在橫幅附註 ----
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", V(v_cmp0, "n2"), ["上檔空間 ", V(v_up0, "sp1")]]
cfg["hero"]["note"] = list(cfg["hero"]["note"]) + ["；進加權的 Comparable 為相對倍數調整後 ", V(v_cmp, "n0"), "（調整係數 ", f"@Val!B{find('Val', '相對倍數調整係數')}|n2", "）。"]

# ---- 後處理二：龍捲風圖是「FY2028 EPS（情境 A）」的變動，不是加權目標價；標題改回正確口徑 ----
for b in sec("sens")["blocks"]:
    if b["type"] == "tornado":
        b["title"] = ("FY2028 EPS 相對基準的變動（情境 A；變動未達兩位數百分比的案例）" if b.get("collapse")
                      else "FY2028 EPS 相對基準的變動（情境 A；變動達兩位數百分比的案例；一次動一個輸入）")

# ---- 後處理三：估值章的「以 Andy 合理市盈率計算的隱含價」長條圖放回該表之後（轉換器會把所有長條圖移到章首）----
vb = sec("valuation")["blocks"]
andy = [b for b in vb if b["type"] == "bars" and "Andy" in str(b.get("title"))]
for b in andy:
    vb.remove(b)
vb += andy

# ---- 後處理四：命題章補情境定義（原命題章的 kv，轉換器不保留）；附錄的版本要點改列 README 各版說明列 ----
th = sec("thesis")
src_th = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))
old_kv = next(b for b in next(s for s in src_th["sections"] if s["id"] == "thesis")["blocks"] if b["type"] == "kv")
th["blocks"] += [{"type": "h3", "text": "情境定義與口徑"}, copy.deepcopy(old_kv)]
ws = wb["README"]
r_end = find("README", "展示層")
ver_rows = [[f"@README!A{r}"] for r in range(2, r_end) if re.match(r"^(v\d|版本：|\d{4}-\d{2}-\d{2}：)", str(ws.cell(r, 1).value or ""))]
st = sec("status")
for i, b in enumerate(st["blocks"]):
    if b["type"] == "h3" and b.get("text") == "附錄：版本要點":
        st["blocks"][i + 1]["rows"] = ver_rows
        break
else:
    ins = next(i for i, b in enumerate(st["blocks"]) if b["type"] == "h3" and "未找到" in str(b.get("text")))
    st["blocks"][ins:ins] = [{"type": "h3", "text": "附錄：版本要點"},
                             {"type": "table", "collapse": True, "summary": "展開：各版的修改要點", "head": ["要點"], "rows": ver_rows}]

# 第三輪修正（2026-11-06）：情境驅動表依節標題取列並自檢、情境定義 kv 自檢、DCF 明細比例列格式、橫幅幣別（r3_html.py）
import r3_html  # noqa: E402
r3_html.post(cfg, wb, '日圓')
# 三層協調 C4（規則 v1.4；tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(cfg, X)


# ================= v0.5.5（2026-11-14）：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.7；docs/plan/個股HTML版面規格.md「v1.3 規格」） =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含 C0）→ 資料狀態；Ibiden 無個別模組。
# 數字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末相對基準表、Coord 頁 C0；ibiden/v055_post.py 產生）。列號依標籤找。
def _v13():
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RV, SW, CW, VW = wb["README"], wb["PL"], wb["RevDCF"], wb["Sens"], wb["Coord"], wb["Val"]
    fnd = lambda ws, p, start=1: next(r for r in range(start, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(p))
    r13 = find("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.5.6"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    fmt_u = lambda u: {"%": "p1", "¥": "n2", "十億股": "n3", "x": "n3", "百萬顆": "n2", "千日圓／顆": "n1"}.get(str(u).strip(), "n1")
    # ---- 橫幅：≤200 字白話＋對帳中揭露句；評等規則與歸因移到估值章 ----
    old_note = list(cfg["hero"]["note"])
    i_cd = len(old_note) - 2   # 末兩項＝" "、Coord 橫幅揭露句
    assert old_note[i_cd] == " " and str(old_note[i_cd + 1]).startswith("@Coord!B"), old_note
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"] + old_note[i_cd:]
    # ---- 1. 命題：主圖＝三情境營收與情境 B 分部（堆疊）----
    pa, pb, pc = find("PL", "情境 A"), find("PL", "情境 B"), find("PL", "情境 C")
    fa = lambda base, lab_: next(r for r in range(base, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    rb = find("PL", "情境 B：")
    hdrB = rb + 1
    prow = lambda lab_: next(r for r in range(rb, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip().startswith(lab_))
    r_ai, r_ne, r_cer, r_oth, r_rev = prow("AI 伺服器載板營收"), prow("電子事業非 AI 營收"), prow("陶瓷事業營收"), prow("其他事業營收"), next(r for r in range(prow("C. 陶瓷"), PLW.max_row + 1) if PLW.cell(r, 1).value == "營收")
    r_opm = next(r for r in range(r_rev, PLW.max_row + 1) if PLW.cell(r, 1).value == "營業利益率")
    revA, revC = fa(pa, "營收"), fa(pc, "營收")
    th = S["thesis"]["blocks"]
    chart1 = {"type": "theme", "title": "營收分部（情境 B）與三情境營收（十億日圓）", "x": f"@PL!D{hdrB}:G{hdrB}", "fmt": "n0", "stack": True, "pad": 1.35,
              "total": f"@PL!D{r_rev}:G{r_rev}",
              "bars": [{"name": f"@PL!A{r_ai}", "y": f"@PL!D{r_ai}:G{r_ai}", "color": "#0b7a80"},
                       {"name": f"@PL!A{r_ne}", "y": f"@PL!D{r_ne}:G{r_ne}", "color": "#6fc7cb"},
                       {"name": f"@PL!A{r_cer}", "y": f"@PL!D{r_cer}:G{r_cer}", "color": "#c29a2b"},
                       {"name": f"@PL!A{r_oth}", "y": f"@PL!D{r_oth}:G{r_oth}", "color": "#8a8c8c"}],
              "lines": [{"name": [f"@PL!A{pa}", "營收"], "y": f"@PL!D{revA}:G{revA}", "color": "#0b7a80", "dash": True},
                        {"name": [f"@PL!A{pc}", "營收"], "y": f"@PL!D{revC}:G{revC}", "color": "#b4472f", "dash": True}],
              "caption": "堆疊＝情境 B（基準）的四個分部，頂端數字＝情境 B 合併營收；虛線＝情境 A、情境 C 的合併營收。FY2026 為公司上修後預測的錨點；三情境差異來自加速器顆數與短缺溢價。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th.insert(0, chart1)
    th.insert(1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
    for b in th:
        if b["type"] == "h3" and b.get("text") == "情境定義與口徑":
            b["text"] = "情境定義"
    # ---- 2. 估值：第一個區塊＝四個目標價與現價（G 位置）；規則與歸因說明放圖下 ----
    va = S["valuation"]["blocks"]
    g0 = find("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_pxg = g0 + 7
    four = {"type": "bars", "title": "四個目標價與現價（日圓；豎線＝現價）", "fmt": "n2",
            "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                      ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
            "ref": {"label": f"@Val!A{r_pxg}", "value": f"@Val!B{r_pxg}"}}
    tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_pxg}", f"@Val!B{r_pxg}|n2", "", f"@Val!D{r_pxg}"]]}
    keep = [b for b in va if not (b.get("type") == "bars" and b.get("title") == "目標價與現價")]
    i_co = next(i for i, b in enumerate(keep) if b.get("type") == "h3" and b.get("text") == "對帳中：兩讀數目標價")
    head_note = {"type": "note", "parts": old_note[:i_cd]}
    S["valuation"]["blocks"] = [four, head_note, tbl4] + keep
    intro = S["valuation"].pop("intro", None)
    if intro:
        S["valuation"]["blocks"].insert(2, {"type": "note", "parts": [intro]})
    # ---- 3. 敏感度：主圖＝加權目標價相對基準（H 位置）----
    se = S["sens"]["blocks"]
    t0 = find("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SW.cell(r2 + 1, 1).value:
        r2 += 1
    tor = []
    for col, side in (("B", "低"), ("C", "高")):
        vals = {r: SW[f"{col}{r}"].value for r in range(r1, r2 + 1) if isinstance(SW[f"{col}{r}"].value, (int, float))}
        big = [r for r, v in vals.items() if abs(v) >= 0.1]
        small = [r for r, v in vals.items() if 0.0005 <= abs(v) < 0.1]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        tor.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
        tor.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                        summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例（幾乎為零者見全部案例表）"))
    full = {"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
            "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1, r2 + 1)]}
    snap = fnd(SW, "快照狀態")
    for b in se:
        if b.get("type") == "tornado":
            b["summary"] = ("展開：對照：FY2028 EPS 相對基準的變動未達兩位數百分比的案例（情境 A）" if b.get("collapse") else "展開：對照：FY2028 EPS 相對基準的變動達兩位數百分比的案例（情境 A，一次動一個輸入）")
            b["collapse"] = True
        if b.get("type") == "table" and not b.get("collapse"):
            b["collapse"] = True
            b["summary"] = "展開：目標價參數表（WACC、β、ERP、永續成長率等）"
    intro_s = {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入；變動達兩位數百分比者直接顯示，其餘摺疊）。業務主因為整本重算快照，快照狀態："
                                          , f"@Sens!B{snap}", "；目標價參數取 Sens 頁「目標價參數敏感度」表。下方摺疊為全部案例與 FY2028 EPS 敏感度對照。"]}
    sec_intro = S["sens"].pop("intro", None)
    S["sens"]["blocks"] = tor[:1] + tor[2:3] + [intro_s, tor[1], tor[3], full] + se
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").strip().startswith(p))
    ra, rcmp, rsum, yh = rv("A. DCF"), rv("B. Comparable"), rv("C. 摘要"), rv("項目（情境 B）")
    rb_, rbb, rcc = rv("(a)"), rv("(b)"), rv("(c)")
    m1, m0 = rv("現價隱含的營業利益率", rb_), rv("模型：基準的等效營業利益率", rb_)
    s1, rv1, rv0, g1, g0b = rv("現價隱含的放大倍數", rbb), rv("現價隱含的 FY2028 營收", rbb), rv("模型：FY2028 營收", rbb), rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
    w1, w0 = rv("現價隱含的 WACC", rcc), rv("模型 WACC", rcc)
    a_rows = [rv(p, ra) for p in ("支持現價所需的評價日每股價值", "淨現金滾到評價日", "評價用稀釋股數", "支持現價所需的企業價值", "模型企業價值", "EV 需放大", "模型：情境 B 評價日每股價值")]
    afmt = ["n2", "n1", "n3", "n1", "n1", "x1", "n2"]
    pe1, rc1, pe2, pm = rv("只看 Comparable：支持現價所需的前瞻本益比", rcmp), rv("只看 Comparable：所需的相對倍數係數", rcmp), rv("加權（DCF 不變）", rcmp), rv("同業前瞻本益比中位數", rcmp)
    yrs_rows = list(range(yh + 1, yh + 12))
    yfmt = ["n1", "n1", "p1", "p1", "n3", "n3", "n3", "n3", "n1", "n1", "n1"]
    rev_ = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
            "blocks": [
                {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
                {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
                {"type": "table", "caption": "現金流折現法（情境 B 路徑）：一次只動一個因素，反解到現價", "head": ["項目", "現價隱含", "模型基準"],
                 "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                          ["現金流放大倍數（FY2028 起各年同乘；模型基準為一倍）", f"@RevDCF!B{s1}|n2", "—"],
                          ["FY2028 營收（十億日圓）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                          ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0b}|p1"],
                          ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
                {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(a_rows, afmt)]},
                {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
                 "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}", f"@RevDCF!C{r}|{f}", f"@RevDCF!D{r}|sp1"] for r, f in ((pe1, "n1"), (rc1, "n2"), (pe2, "n1"))]},
                {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFGH"], "collapse": True,
                 "summary": "展開：情境 B 各年數字（反解的計算基礎；演算過程見 RevDCF 頁）",
                 "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{f}" for c in "BCDEFGH"] for r, f in zip(yrs_rows, yfmt)]}]}
    # ---- 5. 損益：標準損益表 ----
    p0 = find("PL", "標準損益表（v0.5.5")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（含研發）", "營業利益", "營業利益率", "其他收入（淨額）", "稅前利益", "所得稅",
            "稅後淨利（歸屬母公司）", "淨利率", "股數（十億股，分割後）", "EPS（基本；分割後）"]
    PC = "CDEFGHIJK"
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in PC]
    r_ratio = plr("比率"); r_chk = next(r for r in range(r_ratio, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("核對（應為 0"))
    r_end = next(r for r in range(r_chk, PLW.max_row + 1) if not PLW.cell(r, 1).value) - 1
    is_tbl = {"type": "table", "caption": "標準損益表（十億日圓；EPS 為日圓）：FY2023A–FY2025A 實績、FY2026E–FY2031E 情境 B（基準；GAAP＝日本會計基準）",
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC],
              "rows": [rowx(plr(k)) for k in keys] + [[f"@PL!A{plr('（備註）EPS 完全稀釋（公司公告，分割後換算；只列實績）')}"] + [f"@PL!{c}{plr('（備註）EPS 完全稀釋（公司公告，分割後換算；只列實績）')}|n2" for c in "CDE"] + [""] * 6]
                      + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]}
    chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC], "collapse": True,
               "summary": "展開：核對列（歷史捨入差、預測期對 PL 情境 B 與延伸表、FY2030E–FY2031E 的推法差額）",
               "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{'p1' if '占營收' in str(PLW.cell(r, 1).value) else 'n3'}" for c in PC] for r in range(r_chk + 1, r_end)]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（十億日圓）與毛利率、營業利益率、淨利率：FY2023A–FY2031E（情境 B）", "x": f"@PL!C{ph}:K{ph}",
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:K{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:K{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:K{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:K{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP）。FY2024A–FY2025A 淨利含政策持股出售利益等特別利益；FY2026E 為公司上修後預測；FY2030E–FY2031E 為收斂法延伸期（毛利率與營業費用率沿用前一年）。"}
    old_pl = S["pl"]["blocks"]
    pls = [b for b in old_pl if b["type"] == "pl"]
    assert len(pls) == 2, [b["type"] for b in old_pl]
    ext_notes = [b for b in old_pl if b["type"] == "note"]
    old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境損益摘要與折線（FY2026–FY2029；營收、電子事業營益率、營業利益率、EPS）")]
    old_blocks += [{"type": "kv", "items": next(b["items"] for b in old_pl if b["type"] == "kv"), "collapse": True, "summary": "展開：情境摘要補充列"}] if False else []
    old_blocks += [dict(b, collapse=True, summary="展開：延伸期（FY2030 起）的收斂法說明") for b in ext_notes]
    old_blocks += [dict(pls[1], collapse=True, summary="展開：三情境延伸到 FY2032（收斂法）")]
    S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）",
                   blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, {"type": "note", "parts": [f"@PL!A{r_end + 2}"]}, chk_tbl] + old_blocks)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    r_eff, r_nv = prow("有效顆數"), prow("NVIDIA 加速器顆數")
    gen_rows = [prow(p_) for p_ in ("NVIDIA 加速器顆數", "非 NVIDIA 加速器顆數", "有效顆數")]
    r_idx = [prow(p_) for p_ in ("每顆載板金額指數", "Ibiden 份額相對值", "短缺溢價指數")]
    r_pr = prow("AI 載板營業利益率 ＝")
    seg = lambda rs: [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in "DEFG"] for r in rs]
    r_ne2, r_nem, r_nea = prow("電子事業營收"), prow("電子事業營益率"), prow("電子事業營業利益")
    r_tot = [prow("營收年增"), prow("AI 伺服器載板 ÷ 營收"), prow("營業利益"), prow("營業利益率"), prow("淨利（歸屬母公司）"), prow("EPS（日圓）")]
    seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；十億日圓）：AI 伺服器載板＋電子事業非 AI＋陶瓷＋其他事業＝合併營收", "head": [f"@PL!A{hdrB}", f"@PL!B{hdrB}"] + [f"@PL!{c}{hdrB}" for c in "DEFG"],
               "rows": [{"sub": "一、AI 伺服器載板＝有效加速器顆數 × 每顆有效金額 × 金額指數 × 份額相對值 × 短缺溢價指數"}] + seg(gen_rows + r_idx + [r_ai])
                       + [{"sub": "二、電子事業＝AI 載板＋非 AI 營收；營業利益率隨短缺溢價回落"}] + seg([r_ne, r_ne2, r_pr, r_nea, r_nem])
                       + [{"sub": "三、陶瓷、其他事業與合併"}] + seg([r_cer, r_oth, r_rev] + r_tot)}
    seg_chart = {"type": "theme", "title": "AI 伺服器載板營收（情境 B，十億日圓）與有效加速器顆數（百萬顆，右軸）", "x": f"@PL!D{hdrB}:G{hdrB}", "fmt": "n0", "fmt_r": "n1", "stack": True, "pad": 1.2,
                 "bars": [{"name": f"@PL!A{r_ai}", "y": f"@PL!D{r_ai}:G{r_ai}", "color": "#0b7a80"}, {"name": f"@PL!A{r_ne}", "y": f"@PL!D{r_ne}:G{r_ne}", "color": "#6fc7cb"}],
                 "lines": [{"name": [f"@PL!A{r_eff}"], "y": f"@PL!D{r_eff}:G{r_eff}", "axis": "r", "color": "#3b3838", "labels": True}],
                 "caption": "長條＝電子事業的 AI 伺服器載板與非 AI 部分；折線＝有效加速器顆數（NVIDIA ＋ 權重 × 非 NVIDIA；情境 B 在 CY2028 起顆數低於原始值）。FY2026 為校準年。"}
    dr = S["drivers"]["blocks"]
    S["drivers"].update(title="驅動因子：分部營收推導")
    S["drivers"].pop("intro", None)
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + dr
    # ---- 7. 對帳：C0 ----
    c1 = fnd(CW, "1. 讀哪一層"); c2 = fnd(CW, "2. C0"); c3 = fnd(CW, "3. 對帳中年份")
    h1, h2 = c1 + 1, next(r for r in range(c2, CW.max_row + 1) if CW.cell(r, 1).value == "項目")
    cnt = fnd(CW, "對帳中項目數")
    cunit = lambda r: str(CW.cell(r, 2).value or "")
    cf = lambda r: "p1" if cunit(r) in ("%",) else "n3" if cunit(r) in ("百萬顆", "x", "x（FY2026＝1）", "US$B") else "n1"
    c0_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|{cf(r)}", f"@Coord!E{r}|{cf(r)}", f"@Coord!F{r}|sp1", f"@Coord!G{r}"]
               for r in range(h2 + 1, cnt) if CW.cell(r, 4).value is not None]
    c0_tbl = {"type": "table", "caption": "讀取層級核對：個股讀入的層級值 vs 層級值（同源，不判定、不計入對帳中；有檢驗力的對帳見上表「公司實績 vs 層級隱含值」）", "collapse": True, "summary": "展開：讀取層級核對（同源，不判定）", "head": [f"@Coord!{c}{h2}" for c in "ABCDEFG"], "rows": c0_rows}
    # C0（實績）：公司實績 vs 層級隱含值（Coord 頁「4. C0（實績）」；v0.5.6）
    a4 = fnd(CW, "4. C0（實績）"); ih = next(r for r in range(a4, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); mh = next(r for r in range(ih, CW.max_row + 1) if CW.cell(r, 1).value == "讀數")
    a4e = next(r for r in range(mh, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith("對帳中項目數"))
    fm = lambda r: "p1" if str(CW.cell(r, 2).value) == "%" else "n3" if str(CW.cell(r, 2).value) == "US$B" else "n1"
    act_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", (f"@Coord!F{r}|p1" if r == a4e - 1 else f"@Coord!F{r}|sp1"), f"@Coord!G{r}"]
                for r in range(mh + 1, a4e)]
    in_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|" + ("n0" if CW.cell(r, 2).value == "月" else fm(r)), f"@Coord!E{r}", f"@Coord!H{r}"] for r in range(ih + 1, mh - 1) if CW.cell(r, 1).value]
    act_blocks = [{"type": "h3", "text": "公司實績對層級隱含值（實績對帳）"}, {"type": "note", "parts": [f"@Coord!A{a4 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 2}"]},
                  {"type": "table", "caption": "電子事業營收：決算短信實績 vs Siliconomics 基準隱含值（主讀數、第二讀數；差距超過正負十五個百分點且同號標「對帳中」）", "head": [f"@Coord!{c}{mh}" for c in "ABCDEFG"], "rows": act_rows},
                  {"type": "note", "parts": ["對帳中項目數：", f"@Coord!G{a4e}|n0", "。"]},
                  {"type": "table", "caption": "實績與層級隱含值的輸入、等級與出處（展開看）", "collapse": True, "summary": "展開：實績、層級值、換算參數與出處", "head": [f"@Coord!{c}{ih}" for c in "ABCDE"] + [f"@Coord!H{ih}"], "rows": in_rows}]
    c0_note = [{"type": "h3", "text": "讀取層級核對（同源，不判定）"}, {"type": "note", "parts": [f"@Coord!A{c2}"]}, {"type": "note", "parts": [f"@Coord!A{c2 + 1}"]}, c0_tbl, {"type": "note", "parts": [f"@Coord!A{cnt + 1}"]}]
    c0_src = {"type": "table", "caption": "層級名稱、版本、口徑與證據等級（每列的來源；展開看）", "collapse": True, "summary": "展開：每列的層級名稱、版本、口徑與證據等級",
              "head": ["項目", "年度", "層級名稱與版本／口徑／證據等級／備註"], "rows": [[f"@Coord!A{r}", f"@Coord!C{r}", f"@Coord!H{r}"] for r in range(h2 + 1, cnt) if CW.cell(r, 8).value]}
    def _rd(r):
        v = CW.cell(r, 2).value
        return [f"@Coord!A{r}", f"@Coord!B{r}" + ("|n3" if isinstance(v, (int, float)) and not isinstance(v, bool) else ""), f"@Coord!C{r}", f"@Coord!D{r}", f"@Coord!E{r}"]
    rd_tbl = {"type": "table", "caption": f"@Coord!A{c1}", "head": [f"@Coord!{c}{h1}" for c in "ABCDE"], "collapse": True, "summary": "展開：本模型讀哪一層、哪一版（下游只註明版本、不自行重估）",
              "rows": [_rd(r) for r in range(h1 + 1, c2 - 1) if CW.cell(r, 1).value]}
    rf0 = fnd(CW, "層級之間"); rfh = next(r for r in range(rf0, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); rfe = next(r for r in range(rfh, CW.max_row + 1) if not CW.cell(r + 1, 1).value)
    ref_blocks = [{"type": "h3", "text": f"@Coord!A{rf0}"}, {"type": "note", "parts": [f"@Coord!A{rf0 + 1}"]},
                  {"type": "table", "head": [f"@Coord!{c}{rfh}" for c in "ABCDEFG"],
                   "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n3", f"@Coord!E{r}|n3", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in range(rfh + 1, rfe + 1)]}]
    c3_notes = [{"type": "note", "parts": [f"@Coord!A{c3 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{c3 + 2}"]}]
    S["recon"]["blocks"] = [{"type": "h3", "text": "三層協調：由下往上對帳（公司實績對各層）"}] + act_blocks + c0_note + ref_blocks + c3_notes + [c0_src, rd_tbl] + [{"type": "h3", "text": "公司目標、先驗比對與校準"}] + S["recon"]["blocks"]
    # ---- 8. 資料狀態 ----
    st = S["status"]["blocks"]
    st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    S["revdcf"] = rev_
    cfg["sections"] = [S[k] for k in order]


_v13()
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)

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
    ["淨利、股數與匯率", ["ni_oi", "sh_g", "fx"]],
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
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)

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
v_tp, v_dcf, v_cmp, v_cmp0 = find("Val", "目標價（DCF 與 Comparable 加權）"), find("Val", "DCF 目標價（情境加權）"), find("Val", "Comparable 目標價（相對倍數調整後）"), find("Val", "Comparable 目標價（市場主流同業法）")
v_px, v_up, v_up0, v_cons = find("Val", "現價（"), find("Val", "上檔空間：加權目標價"), find("Val", "上檔空間：Comparable 目標價"), find("Val", "市場共識目標價")
v_rec, v_grade, v_band, v_hist, v_meth = find("Val", "投資建議"), find("Val", "上檔空間級距"), find("Val", "價格帶"), find("Val", "歷史倍數檢查"), find("Val", "目標價方法")
v_eps_f, v_eps_m = find("Val", "與市場共識目標價的差距：EPS"), find("Val", "與市場共識目標價的差距：倍數")
v_s12, v_reinv = find("Val", "DCF 情境價值推到"), find("Val", "再投資率", find("Val", "0.1 DCF"))
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
    {"type": "table", "caption": "DCF 折現率（股東權益現金流折現；直接顯示）", "head": ["項目", "數值"], "rows": [
        [f"@In!B{INK['tp_r']}", f"@In!E{INK['tp_r']}|p1"], [f"@In!B{INK['tp_g']}", f"@In!E{INK['tp_g']}|p1"],
        [f"@In!B{INK['tp_roe']}", f"@In!E{INK['tp_roe']}|p1"], [f"@Val!A{v_reinv}", V(v_reinv, "p1")]]},
    dcf_tbl, cmp_tbl,
    {"type": "table", "caption": "相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司相對同業的長期倍數比）", "head": ["項目", "數值"], "rows": [[k, v] for k, v in rm_kv["items"]]},
] + rest

# ---- 前處理二：敏感度章「目標價參數」表的列號改為 Sens 頁實際位置（舊設定的列號已位移）----
se = sec("sens")
for b in se["blocks"]:
    if b["type"] == "table" and cap(b).startswith("@Sens!A"):
        t0 = int(cap(b)[7:])
        assert str(wb["Sens"][f"A{t0}"].value).startswith("目標價參數"), cap(b)
        t0 = find("Sens", "目標價參數")
        r1 = t0 + 2
        while wb["Sens"][f"A{r1 + 1}"].value:
            r1 += 1
        b["caption"] = f"@Sens!A{t0}"
        b["head"] = [f"@Sens!{c}{t0 + 1}" for c in "ABCDEF"]
        b["rowgen"]["rows"] = f"{t0 + 2}-{r1}"

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
    ["折現率（DCF）", ["tp_r", "tp_g", "tp_roe"]],
    ["目標價規則與情境權重", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC", "fade_gA", "fade_gB", "fade_gC", "nim10_f"]],
    ["自由現金流假設（不進目標價）", ["da_rev", "cx_rev"]]]]],
 "blocks_to": {"單一參數": "電子事業：AI 伺服器載板", "Ibiden 在 AI 加速器載板": "電子事業：AI 伺服器載板", "@Units!A": "電子事業：AI 伺服器載板"},
 "main": {
  "目標價結論": ["目標價（DCF 與 Comparable 加權）", "DCF 目標價", "Comparable 目標價（相對倍數調整後）", "現價", "上檔空間：加權目標價", "上檔空間：Comparable 目標價（相對倍數調整後）", "投資建議", "價格帶", "市場共識目標價"],
  "DCF 折現率": ["折現率", "長期成長率", "ROE", "再投資率"],
  "0.1 DCF": ["DCF 情境每股價值", "DCF 情境價值推到", "DCF 目標價", "再投資率"],
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

# ---- 後處理一：橫幅的 Comparable 目標價改用「相對倍數調整後」（加權目標價實際使用的值；轉換器預設抓到未調整的列）----
cfg["hero"]["tiles"][1] = [VA(v_cmp), V(v_cmp, "n2"), ["上檔空間 ", f"@README!B{rd_upc}|sp1"]]

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

json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)

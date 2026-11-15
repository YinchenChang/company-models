# 由 html_src.json（v0.1 的 HTML 設定，原由舊 gen_html_json.py＋rowmap 產生；v0.2 起內容不動，舊產生器已移除）套用「個股 HTML 版面規格 v1.0」（docs/plan/個股HTML版面規格.md），產生 tuc/html.json。
# 用法：python3 tuc/gen_html_json.py [成品 xlsx]（省略時取 tuc/dist/ 內唯一的 xlsx）；勿手改 html.json。
# 轉換器 tools/company_html/relayout.py 之外，本檔只做 台燿專屬的前處理（估值章重排）與後處理（橫幅、龍捲風標題、附錄）。
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
    r_end = find("Val", "DCF 目標價（情境加權）", rk)   # 第二輪修正：依標籤定位（模組 v2 stub 期在「＋淨現金與投資」後多兩列）
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
             "fcf": ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]}


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
for b in old:   # v0.2：舊「DCF 情境價值」長條圖改讀 0.1 節新列
    if b["type"] == "bars" and "DCF 情境價值" in str(b.get("title")):
        for it, c in zip(b["items"], "BCD"):
            it["value"] = f"@Val!{c}{v_s12}"
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
_sl = find("Sens", "讀法：", 5) - 1
while not wb["Sens"].cell(_sl, 1).value:
    _sl -= 1
for b in se["blocks"]:
    if b["type"] == "tornado":
        b["rows"] = f"5-{_sl}"
    if b["type"] == "table" and b.get("rowgen", {}).get("sheet") == "Sens":
        b["rowgen"]["rows"] = f"4-{_sl}"
    if b["type"] == "note" and str(b.get("parts", [""])[0]).startswith("@Sens!A") and b["parts"][0] != "@Sens!A2":
        b["parts"] = [f"@Sens!A{find('Sens', '讀法：', 5)}"]
se["blocks"] += [{"type": "h3", "text": "目標價參數（標準版 DCF）"}, sens_tp_block()]   # v0.2：tools/fcff_dcf.write_sens 的閉式案例

# ---- 轉換：版面規格 v1.0 ----
scn_tok = re.split(r"[.\s]", str(wb["In"].cell(find("In", "情境驅動對照"), 1).value))[0]
# In 頁參數依「節」歸到業務分塊（每個參數剛好一組；轉換器會檢查）
B_AI, B_CAP, B_CO = "AI 高速 CCL：市場、台燿份額與價格", "產能、營收結構與毛利率", "全公司：股價、折現率、目標價規則、稅與現金流"
SECMAP = [("A.", B_AI, "AI 部署與每 GW 高速 CCL 金額（讀 Siliconomics）"), ("B.", B_AI, "台燿份額與價格（公司變數）"), ("Z.", B_AI, "Wattonomics 錨定係數（只用於敏感度）"),
          ("C.", B_CAP, "產能（全產全銷上限）"), ("D.", B_CAP, "營收結構錨點（損耗等級組合）"), ("E.", B_CAP, "毛利率"),
          ("F.", B_CO, "FY2026 錨點（季報實績與第四季假設）"), ("G.", B_CO, "營業費用、業外、稅與股數"),
          ("H.", B_CO, "共識與券商（只作對帳）"), ("I.", B_CO, "股價"), ("J.", B_CO, "目標價規則、情境權重與延伸期收斂"),
          ("M2.", B_CO, "折現率（標準版 DCF：無風險利率、β、股權風險溢酬、WACC、永續成長率、RONIC）")]
EXTRA = {k: "資本結構與淨現金（WACC 權重、EV 到每股價值）" for k in FCFF_KEYS["bs"]} | {k: "自由現金流假設（DCF：折舊、資本支出、營運資金）" for k in FCFF_KEYS["fcf"]}
def _in_groups():
    ws = wb["In"]; cur = None; G = {}
    for r in range(4, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if a and not b:
            hit = next((m for m in SECMAP if str(a).startswith(m[0])), None)
            cur = hit if hit else ("SKIP",)
            continue
        if not a or cur is None or cur[0] == "SKIP":
            continue
        blk, grp = cur[1], EXTRA.get(str(a), cur[2])
        G.setdefault(blk, {}).setdefault(grp, []).append(str(a))
    return [[blk, [[g, ks] for g, ks in G[blk].items()]] for blk in (B_AI, B_CAP, B_CO) if blk in G]
IN_GROUPS = _in_groups()
SPEC = {
 "in_skip": [scn_tok],
 "in_groups": IN_GROUPS,
 "blocks_to": {"單一參數": B_AI, "AI 高速 CCL 市場": B_AI, "@Units!A": B_AI},
 "main": {
  "目標價結論": ["目標價（DCF 與 Comparable 加權）", "DCF 目標價", "Comparable 目標價（相對倍數調整後）", "現價", "上檔空間：加權目標價", "上檔空間：Comparable 目標價（相對倍數調整後）", "投資建議", "價格帶", "市場共識目標價"],
  **FCFF_MAIN,
  "0.2 Comparable": ["同業前瞻 P/E 中位數", "12 個月後的前瞻 EPS", "Comparable 目標價"],
  "相對倍數調整": ["本公司 10 年中位數", "同業 10 年中位數的中位數", "相對倍數調整係數", "Comparable 目標價（相對倍數調整後）"],
  "現價與市場隱含倍數": ["現價", "現價 ÷ FY2027 EPS", "現價 ÷ FY2028 EPS", "  現價 ÷ 近 12 月 EPS"],
 },
}
if sec("drivers").get("intro"):   # 驅動因子章的說明文字（轉換器轉成 note）放第一塊
    SPEC["blocks_to"][sec("drivers")["intro"]] = B_AI
# blocks_to 的 note 以 parts[0] 開頭比對：Units 頁註解
for b in sec("drivers")["blocks"]:
    if b["type"] == "note" and str(b.get("parts", [""])[0]).startswith("@Units!A"):
        SPEC["blocks_to"][b["parts"][0]] = B_AI
cfg = relayout(cfg, X, SPEC)

# ---- 後處理一：橫幅中間格＝只看 Comparable（市場主流同業法、未經相對倍數調整；工作單補充規則 A），附它自己的上檔空間；
#      進加權的是相對倍數調整後的值，寫在橫幅附註 ----
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", V(v_cmp0, "n2"), ["上檔空間 ", V(v_up0, "sp1")]]
cfg["hero"]["note"] = list(cfg["hero"]["note"]) + ["；進加權的 Comparable 為相對倍數調整後 ", V(v_cmp, "n0"), "（調整係數 ", f"@Val!B{find('Val', '相對倍數調整係數')}|n2", "）。"]

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
r3_html.post(cfg, wb, '新台幣')
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)
# ───────── v0.2.4（2026-11-15）：個股 HTML 版面規格 v1.3＋三層協調（兩讀數、C0（實績））；共通步驟 tuc/html13.py（高速 CCL／PCB 四家同一份；樣板 delta、auras、emc） ─────────
import html13  # noqa: E402
AIR_ = "高速材料營收（ELL／SLL，M7 以上；AI 相關）＝ MIN（需求面、產能上限 − 其他等級）"
NAIR_ = "其他等級營收（VLL／LL、非低損耗、特殊應用）"
REVR_ = "營收（單一應報導部門：銅箔基板與膠片）"
H = {
    "ver": "v0.2.4",
    "units_chart": lambda pB, hB: {
        "type": "theme", "title": "高速材料與其他等級營收（情境 B，十億元）與台燿在 AI 高速 CCL 市場的有效份額（右軸）", "x": f"@PL!D{hB}:G{hB}",
        "fmt": "n1", "fmt_r": "p1", "stack": True, "pad": 1.25, "total": f"@PL!D{pB(REVR_)}:G{pB(REVR_)}",
        "bars": [{"name": ["高速材料（ELL／SLL，AI 相關）"], "y": f"@PL!D{pB(AIR_)}:G{pB(AIR_)}", "color": "#0b7a80"},
                 {"name": ["其他等級（VLL／LL、非低損耗、特殊應用）"], "y": f"@PL!D{pB(NAIR_)}:G{pB(NAIR_)}", "color": "#b9b4a6"}],
        "lines": [{"name": ["台燿在 AI 高速 CCL 市場的有效份額"], "y": f"@PL!D{pB('台燿有效份額')}:G{pB('台燿有效份額')}", "axis": "r", "color": "#3b3838", "labels": True}],
        "caption": "堆疊＝兩塊營收，頂端數字＝合併營收；折線＝台燿有效份額（右軸）。回答命題：AI 高速 CCL 市場擴大時，泰國、常熟、中山擴產與組合升級能讓高速材料營收成長多少。FY2026 為校準年。"},
    "sc_caption": "三情境 FY2027 相同（部署量情境差異從 CY2028 起）；FY2030 起為延伸期（收斂法）。",
    "seg": [("一、AI 高速 CCL 市場與台燿份額", "A. AI 高速 CCL 市場", "B. 產能上限"), ("二、產能上限、其他等級與合併營收", "B. 產能上限", "C. 損益")],
    "seg_caption": "分部營收推導（情境 B；十億元新台幣）：高速材料＝部署 GW × 每 GW 高速 CCL 金額 × 台燿有效份額 × 售價指數，以產能上限減其他等級為上限；加其他等級得合併營收",
    "gw_chart": lambda pB, hB: {
        "type": "theme", "title": "部署 GW（IT 電力，長條）與 AI 高速 CCL 市場（折線，十億美元，右軸）：情境 B", "x": f"@PL!D{hB}:G{hB}",
        "fmt": "n0", "fmt_r": "n1", "stack": False, "pad": 1.3,
        "bars": [{"name": ["部署 GW（IT 電力）"], "y": f"@PL!D{pB('部署 GW（IT 電力）')}:G{pB('部署 GW（IT 電力）')}", "color": "#0b7a80"}],
        "lines": [{"name": ["AI 高速 CCL 市場"], "y": f"@PL!D{pB('AI 高速 CCL 市場（美元；不含價格指數）')}:G{pB('AI 高速 CCL 市場（美元；不含價格指數）')}", "axis": "r", "color": "#c29a2b", "labels": True}],
        "caption": "高速材料營收的產業輸入：部署量與每 GW 高速 CCL 金額（皆讀 Siliconomics）＝市場；乘上台燿有效份額與售價指數即為需求面營收。"},
    "recon_h3": "券商、研調與同業份額對照（原有）",
    "plain": {f"Val!E{find('Val', 'WACC')}": "股權成本與稅後債務成本以市值權重加權；WACC 調整只用於敏感度（基準為零）",
              f"Val!A{find('Val', '永續成長率 g')}": "永續成長率 g（長期成長上限與無風險利率取小者）",
              f"Val!A{find('Val', '永續期 RONIC')}": "永續期 RONIC（＝WACC＋超額報酬差）",
              f"Val!A{find('Val', 'DCF 情境價值推到 12 個月後')}": "DCF 情境價值推到目標時點（以股權成本複利，扣除期間股利）",
              f"In!B{INK['g_tv']}": "永續成長率 g（最後一年之後；長期成長上限與無風險利率取小者）",
              f"In!B{INK['scn_fade_g']}": "FY2032 營收年增收斂目標（A 為基準加兩個百分點、B 為基準、C 為基準減兩個百分點）",
              f"In!B{INK['tp_g']}": "PL 延伸期營收年增的收斂基準（情境 B 的 FY2032 營收年增；不用於 DCF，DCF 永續成長率見折現率一節）"},
    "fold": {"DCF!A2": "展開：DCF 的營業利益與自由現金流口徑"},
}
html13.apply(cfg, wb, X, H, os.path.join(HERE, ".."))
for b in next(s for s in cfg["sections"] if s["id"] == "sens")["blocks"]:   # 台燿原有龍捲風圖是整本重算的加權目標價（sens_tuc.py），不是 FY2028 EPS：摺疊說明改正
    if b.get("type") == "tornado" and b.get("collapse") and str(b.get("summary", "")).startswith("展開：對照：FY2028 EPS"):
        b["summary"] = "展開：對照：整本重算的加權目標價敏感度（In 頁輸入一次改一個；sens_tuc.py 快照）" + ("，未達兩位數百分比的案例" if "未達" in str(b.get("title")) else "")
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("v1.3 ok", os.path.relpath(X, HERE))

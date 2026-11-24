# AutoStore（AUTO.OL）個股模型 v0.1 — AI實體應用 Project（線 5 倉儲物流自動化：立方格 AS/RS）
# 用法：python3 physical_ai/autostore/build_autostore.py → 建置、整本重算 Sens 與反解（LibreOffice 批次）、輸出 dist/ 成品、README.md、sens.json、segments.json
# 會計年度＝曆年（12 月決算，IFRS）。金額：百萬美元（US$M，公司以美元報告）；股數：百萬股；股價與每股目標價：挪威克朗（NOK）；EPS 另列 NOK。
# 結構：README、In、Hist、Q、PL（三情境：AS/RS 市場 × 份額 → 營收 → 標準損益表 → 現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md。
import re, sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261124"
NAME = f"{DATE}_AutoStore個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032

# ---------------- 出處 ----------------
Q3R = "AutoStore Q3 2026 financial result（公司公告，2026-11-05）https://mfn.se/all/a/autostore/autostore-q3-2026-financial-result-a513e386"
Q4R = "AutoStore Q4 2025 financial results（公司公告，2026-02-12）https://www.autostoresystem.com/investors-press-releases/autostore-q4-2025-financial-results（經搜尋摘要讀取，非全文）"
SAI = "stockanalysis.com AutoStore 損益表（資料商；讀取 2026-11-24）https://stockanalysis.com/quote/osl/AUTO/financials/income-statement/"
CALL = "AutoStore Q3 2026 財報電話會議摘要（Yahoo Finance／GuruFocus 轉載；二手，未對照原始報告）https://finance.yahoo.com/markets/stocks/articles/autostore-holdings-ltd-autsf-q3-150936179.html"
PRO = "physical_ai/map/profiles/AUTO.md（本專案公司檔案，股數與 Logiqs 收購引自公司公告）"
YCL = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
SAS = "stockanalysis.com 各公司 statistics 頁（physical_ai/symbotic 讀取 2026-11-23；頁面報價日見各列）"
UST = "美國 10 年期公債殖利率 5.10%（FRED DGS10 2026-11-12；與 physical_ai/symbotic In 頁同一取值，跨模型一致；逾期）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"
WH = "physical_ai/lines/warehouse/dist/20261116_倉儲物流單位經濟_v0.1.xlsx"
SYMX = "physical_ai/symbotic/build_symbotic.py（In 頁同業前瞻 P/E）"

ROWS = []
URLRE = re.compile(r"https?://[A-Za-z0-9._~:/?#@!$&*+,;=%\-]+")
def nu(t): return URLRE.sub("（網址見來源1欄）", t) if isinstance(t, str) else t
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. FY2026 錨點：公司指引與已揭露實績（Interested-party）")
inp("g26_rev", "公司指引：FY2026 營收（約）", "US$M", None, 750.0, None, "Interested-party", Q3R + "：指引上調至約 7.5 億美元（Q2 時約 7.0 億）", "TTM 729.6（" + SAI + "）減 Q4'25 179.7 → 9M'26 約 549.9；Q4 隱含約 200（Derived）", CO)
inp("g_anc", "FY2026 實績相對公司指引", "%", -0.03, 0.0, 0.03, "Assumed", "基準＝公司指引；偏離只來自 Q4（9M 已實績）", "低／高 ±3%（Q4 單季佔全年約 27%）")
inp("gm26", "FY2026 毛利率", "%", 0.71, 0.72, 0.73, "Derived", Q3R + "：Q3'26 毛利率 71.3%；TTM 528.8 ÷ 729.6 ＝ 72.5%；FY2025 72.4%（" + SAI + "）", "")
inp("em26", "FY2026 營業利益率（EBIT 占營收）", "%", 0.27, 0.30, 0.33, "Assumed", Q3R + "：Q3'26 EBIT 54.1 ÷ 192.1 ＝ 28.2%（Q3'25 35.0%）；TTM EBIT 232.3 ÷ 729.6 ＝ 31.8%（" + SAI + "）", "公司揭露調整後 EBIT 率 Q3'26 34.9%；本模型用報告口徑 EBIT（未加回股權激勵等）")
inp("fin26", "FY2026 業外：財務收益 − 財務費用（含匯兌）", "US$M", -30.0, -20.0, -12.0, "Assumed", "TTM 利息費用 16.2（" + SAI + "）；FY2025 稅前 − EBIT ＝ 105.8 − 142.1 ＝ −36.3（Derived）", "淨負債已由 FY2025 年底 180 降到 Q3'26 約 48（" + CALL + "），故費用低於 FY2025")
inp("fx", "美元兌挪威克朗", "NOK/US$", 9.0, 9.77018, 10.5, "Verified", YCL + "：NOK=X 2026-11-20 收盤 9.77018", "每股 NOK 值＝美元每股值 × 此匯率（假設目標時點匯率不變）；低／高為 Assumed")

sec("B. 命題模組：立方格 AS/RS 市場（線 5 下游取數）× AutoStore 份額 → 營收")
MK = {"m26": ("2026", 8, {"A": 10.9294503319294, "B": 10.6943372354646, "C": 10.4524071462717}),
      "m28": ("2028", 9, {"A": 13.4289214189686, "B": 12.5807865845056, "C": 11.7461386987719}),
      "m30": ("2030", 10, {"A": 16.5, "B": 14.8, "C": 13.2}),
      "m32": ("2032", 11, {"A": 19.60365, "B": 16.94452, "C": 14.27712})}
for k, (yy, rr, d) in MK.items():
    for s, col in (("A", "E"), ("B", "D"), ("C", "C")):
        inp(f"{k}_{s}", f"情境 {s}：AS/RS 市場 {yy}（線 5；{'高' if s == 'A' else '基準' if s == 'B' else '低'}）", "US$B", None, d[s], None, "Derived",
            f"{WH} 下游取數!{col}{rr}（WH_ASRS_MKT_{yy}；MarketsandMarkets 口徑，含製造、汽車）", "年間以幾何插值（Derived）；2031 起為線模型假設成長")
for s, v, lo, hi in (("A", 0.003, None, None), ("B", 0.0, -0.002, 0.002), ("C", -0.004, None, None)):
    inp(f"dshr_{s}", f"情境 {s}：AutoStore 占 AS/RS 市場份額每年變動（FY2027 起）", "pt/年", lo, v, hi, "Assumed",
        "FY2026 份額＝營收 ÷ 線 5 市場（基準 7.0%，與線模型 WH_AUTOSTORE_SHARE 0.0701 同口徑：出廠價，低估系統占比）",
        "A：接單 Q3'26 年增 16.5% 快於市場；C：AMR 貨到人（每行成本與 AS/RS 接近，線 5 下游取數 D16、D17）與整合商自製替代；歷史份額變動資料未取得")

sec("C. 損益表結構與現金流")
for s, v in (("A", 0.73), ("B", 0.72), ("C", 0.70)):
    inp(f"gm_{s}", f"情境 {s}：FY2027 起毛利率", "%", 0.70 if s == "B" else None, v, 0.74 if s == "B" else None, "Assumed",
        "FY2024 73.1%、FY2025 72.4%、TTM 72.5%、Q3'26 71.3%（" + SAI + "；" + Q3R + "）", "C：關稅與價格壓力；A：回到 FY2024 水準")
inp("ox_inf", "營業費用基本年增（通膨與人力）", "%", 0.0, 0.03, 0.06, "Assumed", "FY2025 營業費用 247.8（毛利 389.9 − EBIT 142.1；Derived）、FY2024 216.1", "營業費用含研發、銷管與攤銷；公司未在可取得摘要中拆分")
inp("ox_el", "營業費用對營收增幅的彈性", "x", 0.3, 0.5, 0.8, "Assumed", "FY2025 營收 −10.4% 而營業費用 +14.7%（Derived：" + SAI + "），顯示費用剛性高、與營收弱相關", "基準 0.5")
inp("nonop", "業外：財務收益 − 財務費用（FY2027 起）", "US$M", -25.0, -12.0, -5.0, "Assumed", "FY2026 假設 −20 並隨去槓桿收斂；淨負債 48（" + CALL + "）", "")
inp("tax_w", "有效稅率", "%", 0.20, 0.22, 0.25, "Derived", "TTM 47.0 ÷ 216.7 ＝ 21.7%；FY2025 24.0 ÷ 105.8 ＝ 22.7%；FY2024 22.4%（" + SAI + "）", "")
inp("fc_da", "折舊攤銷 ÷ 營收", "%", 0.05, 0.075, 0.09, "Derived", Q3R + "：Q3'26 調整後 EBITDA 81.3 − 調整後 EBIT 67.1 ＝ 14.2 ÷ 營收 192.1 ＝ 7.4%", "資料商 D&A（TTM 23.9）口徑較窄，未採用")
inp("fc_cx", "資本支出（含資本化研發）÷ 營收", "%", 0.07, 0.095, 0.12, "Derived", Q3R + "：現金流轉換率 78.0%（＝(調整後 EBITDA − 資本支出) ÷ 調整後 EBITDA）→ 資本支出約 17.9 ÷ 192.1 ＝ 9.3%", "FY2025 轉換率 76%、調整後 EBITDA 228.1 → 約 54.7 ÷ 538.6 ＝ 10.2%（" + Q4R + "）；實際資本支出金額未取得")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.05, 0.15, 0.30, "Assumed", "Q3'26 營運現金流 101.4 受正向營運資金支持（" + CALL + "）；AutoStore 預收款與整合商結構細節未取得", "")
inp("dps", "每股股利（美元）", "US$", None, 0.0, None, "Assumed", "未找到股利政策；公司 2026 年以庫藏股註銷與去槓桿為主（" + PRO + "）", "收支關卡的股利列＝0；若實際配息，關卡結果下降")

sec("D. 估值輸入（評價日 2026-12-31；資產負債表 2026-09-30）")
inp("px", "股價（2026-11-20 收盤）", "NOK", None, 14.79, None, "Verified", YCL + "：AUTO.OL 2026-11-20 收盤 14.79", "未還原股利。2026-11-16 至 20 全球 AI 與半導體股大跌，收盤檔的下跌不是資料錯誤；當週 AUTO.OL −3.6%（11-13 收盤 15.93）")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("sh_v", "評價日股數（2026-09-30 註銷庫藏股後約 34.12 億股）", "百萬股", None, 3412.0, None, "Derived", PRO + "：2026-09-30 註銷庫藏股後總股數 34.12 億股", "FY2025 稀釋 3,441（" + SAI + "）；股票報酬與選擇權稀釋未取得，視為 0")
inp("nd_v", "淨負債（含租賃；2026-09-30）", "US$M", 30.0, 48.0, 120.0, "Interested-party", CALL + "：Q3'26 淨負債 48、淨槓桿 0.2 倍", "二手來源未對照原始報告；公司自己 Q1'26 揭露 0.5 倍（約 140）、FY2025 年底 180（" + Q4R + "）；高案 120 為 Assumed；Logiqs 收購最高 3,000 萬歐元（約 3,300 萬美元）尚未入帳", CO)
inp("stub_frac", "Stub 期（2026-09-30 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,9,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（美國 10 年期公債；現金流以美元計）", "%", None, 0.051, None, "Verified", UST, "與股價基準日相差逾 2 個交易日（逾期），維持跨模型一致")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "國家風險溢酬", "%", None, 0.0, None, "Assumed", "現金流與報告幣別為美元、挪威 AAA 級，取 0", "")
inp("beta", "β（Blume 調整後中位數）", "x", "=E{r}-0.3", "=MEDIAN(0.67*1.98+0.33,0.67*1.03+0.33,0.67*2.05+0.33)", "=E{r}+0.3", "Derived",
    SAS + "：AutoStore 5 年 β 1.98、Daifuku 1.03、KION 2.05（physical_ai/symbotic In 頁、physical_ai/geekplus In 頁同一取值）", "Blume 0.67β＋0.33；敏感度 ±0.3（Assumed）；高 β 使 ke 達 13% 以上，是目標價偏低的主因之一")
inp("kd", "稅前負債成本", "%", None, "=rf+0.02", None, "Assumed", "rf＋2pt（負債占比低，對 WACC 影響小）", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt（AutoStore 資本密集度低，可能偏保守）", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable：同業前瞻 P/E）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺 AutoStore 對同業 10 年前瞻 P/E 比值歷史）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("E. 同業前瞻 P/E（報價一律以 2026-11-20 收盤等比例換算；分母 EPS 不變；Derived）")
PEERS = [("pe_daif", "Daifuku 大福 6383.T（物料搬運系統）", "=21.33*4129/5429", "2026-09-11 頁面 21.33 倍、收盤 5,429 日圓 → 2026-11-20 收盤 4,129"),
         ("pe_kion", "KION KGX.DE（Dematic 倉儲系統整合＋堆高機）", "=9.85*29.2/40.58", "2026-09-18 頁面 9.85 倍、收盤 40.58 歐元 → 2026-11-20 收盤 29.20"),
         ("pe_geek", "極智嘉 Geek+ 2590.HK（AMR 貨到人）", "=56.32*5.405/9.22", "2026-09-04 頁面 56.32 倍、收盤 9.22 港元 → 2026-11-20 收盤 5.405")]
for k, lab, v, note in PEERS:
    inp(k, f"前瞻 P/E：{lab}（換算至 2026-11-20）", "x", None, v, None, "Derived", f"{SAS}：{note}（{YCL}）；與 {SYMX} 同一取值",
        "原值 Interested-party（賣方共識 EPS）；P/E 為股權口徑不受租賃處理影響；會計口徑：KION、Geek+ IFRS、Daifuku 日本基準", DV)
inp("pe_med", "同業前瞻 P/E 中位數（3 家；進 Comparable）", "x", 7.09, "=MEDIAN(pe_daif,pe_kion,pe_geek)", 33.0, "Derived",
    "3 家中位數；少於約 8 家，不做迴歸（第 10 節第 7 條）。未納入：Symbotic（收盤檔有價，但無前瞻 EPS 可換算）、Ocado（虧損）、Kardex（收盤檔無報價，報價日未取得）", "低＝KION 7.09、高＝極智嘉約 33.0（同業最低與最高；Analogy，家數少、離散大）")
inp("pe_self", "AutoStore 現行前瞻 P/E（換算至 2026-11-20）", "x", None, "=26.20*14.79/14.68", None, "Derived", SAS + "：2026-10-22 頁面 26.20 倍、收盤 14.68 挪威克朗 → 2026-11-20 收盤 14.79（physical_ai/symbotic 取值）", "只作對照，不進模型", DV)


def build(path, ov=None, sens=None, rev=None, bid="{BID}"):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 US$M、股數百萬股、股價與每股目標價 NOK"
    IN["A2"] = "標記：Verified（附出處）／Interested-party（寫明是誰、誘因欄）／Analogy（寫明類比對象）／Assumed（給區間）／Derived（寫明換算）；查不到寫「未取得」"
    hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源1", "來源2／說明", "來源誘因"])
    r = 4; KR = {}
    for row in ROWS:
        if row[0] == "sec":
            IN.cell(r, 1, row[1]).font = BOLD; r += 1; continue
        k, lab, u, lo, b, hi, tag, s1, s2, inc = row
        if k in ov: b = ov[k]
        IN.cell(r, 1, k); IN.cell(r, 2, lab); IN.cell(r, 3, u)
        for col, v in ((4, lo), (5, b), (6, hi)):
            if v is None: continue
            if isinstance(v, str): v = v.replace("{r}", str(r))
            c = IN.cell(r, col, v)
            if col == 5: c.font = BLUE
        IN.cell(r, 7, tag); IN.cell(r, 8, nu(s1) if "未取得" in s1 else s1); IN.cell(r, 9, nu(s2)); IN.cell(r, 10, inc)
        dn(k, f"In!$E${r}"); KR[k] = r; r += 1
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("m26", "AS/RS 市場 2026（線 5，US$B）"), ("m28", "AS/RS 市場 2028（US$B）"), ("m30", "AS/RS 市場 2030（US$B）"), ("m32", "AS/RS 市場 2032（US$B）"),
           ("dshr", "AutoStore 份額每年變動"), ("gm", "FY2027 起毛利率")]
    for f, lab in FAM:
        IN.cell(r, 1, f + "_scn"); IN.cell(r, 2, lab)
        for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR[f + '_' + s]}")
        IN.cell(r, 8, f"引用 In 頁 {f}_A／{f}_B／{f}_C；{lab}"); r += 1
    IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重")
    for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR['w_' + s]}")
    IN.cell(r, 8, "引用 In 頁 w_A／w_B／w_C；合計＝1"); r += 1
    IN.column_dimensions["B"].width = 56; IN.column_dimensions["H"].width = 60

    # ---------------- Hist ----------------
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（年度；US$M；資料商 stockanalysis.com 彙整公司財報，Derived；FY2025 營收與 EBITDA 另對公司公告）。公式只引用本頁（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "標記", "來源"])
    HR = [("h_rev", "營收", 645.7, 601.4, 538.6, "Verified", Q4R + "（FY2025 538.6，−10.4%）；" + SAI),
          ("h_cogs", "營業成本", 207.6, 161.6, 148.7, "Derived", SAI),
          ("h_gp", "毛利", "=B3-B4", "=C3-C4", "=D3-D4", "Derived", "營收 − 營業成本（資料商值 438.1／439.8／389.9）"),
          ("h_op", "營業利益（EBIT）", 225.5, 223.7, 142.1, "Derived", SAI + "；公司公告 FY2025 EBIT 140.1（差 2.0，原因未查核）"),
          ("h_opex", "營業費用（毛利 − 營業利益）", "=B5-B6", "=C5-C6", "=D5-D6", "Derived", "計算"),
          ("h_pbt", "稅前利益", -46.3, 176.1, 105.8, "Derived", SAI + "（FY2023 含 2.39 億美元和解費用，列在營業利益以下）"),
          ("h_fin", "業外（含和解費用）＝稅前 − 營業利益", "=B8-B6", "=C8-C6", "=D8-D6", "Derived", "計算；FY2023 含一次性和解費用"),
          ("h_tax", "所得稅", 13.7, 39.5, 24.0, "Derived", SAI),
          ("h_ni", "淨利", -32.6, 136.6, 81.8, "Derived", SAI),
          ("h_sh", "稀釋股數（百萬股）", 3419.0, 3432.0, 3441.0, "Derived", SAI),
          ("h_ebitda", "調整後 EBITDA（公司定義）", "未取得", "未取得", 228.1, "Verified", Q4R + "（42.3% 利潤率；搜尋摘要）"),
          ("h_ord", "接單", "未取得", "未取得", "未取得", "—", "FY2025 全年接單未取得；Q4'25 為歷史新高（" + Q4R + "）")]
    for i, (k, lab, a, b, c, tag, src) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b, c)):
            cc = H.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): cc.font = BLUE
        H.cell(i, 5, tag); H.cell(i, 6, nu(src) if "未取得" in src else src)
        dn(k + "_23", f"Hist!$B${i}"); dn(k + "_24", f"Hist!$C${i}"); dn(k + "_25", f"Hist!$D${i}")
    H.column_dimensions["A"].width = 44

    # ---------------- Q（季度揭露） ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 季度揭露（公司單一報告分部；US$M）：Q3 2026 與 Q3 2025，及 AS/RS 市場份額對照。公司在可取得的公告中未拆分系統與軟體服務營收，故營收以「市場 × 份額」推導"
    hdr(Q, 2, ["項目", "Q3'25", "Q3'26", "年增", "標記", "來源", "來源誘因"])
    qs = [("營收", 139.0, 192.1), ("毛利", 101.7, 137.1), ("營業利益（EBIT）", 48.6, 54.1), ("調整後 EBITDA", 65.5, 81.3), ("調整後 EBIT", 53.8, 67.1), ("接單", 152.4, 177.6)]
    for i, (lab, a, b) in enumerate(qs, 3):
        Q.cell(i, 1, lab); Q.cell(i, 2, a).font = BLUE; Q.cell(i, 3, b).font = BLUE; Q.cell(i, 4, f"=C{i}/B{i}-1"); Q.cell(i, 5, "Verified"); Q.cell(i, 6, Q3R); Q.cell(i, 7, CO)
        Q.cell(i, 4).number_format = "0.0%"
    Q.cell(9, 1, "毛利率"); Q.cell(9, 2, "=B4/B3"); Q.cell(9, 3, "=C4/C3"); Q.cell(9, 5, "Derived"); Q.cell(9, 6, "計算（公司揭露 73.1%／71.3%）")
    Q.cell(10, 1, "接單 ÷ 營收"); Q.cell(10, 2, "=B8/B3"); Q.cell(10, 3, "=C8/C3"); Q.cell(10, 5, "Derived"); Q.cell(10, 6, "計算；Q3'26 低於 1，營收已超前接單")
    for rr_ in (9, 10):
        for cc_ in "BC": Q[f"{cc_}{rr_}"].number_format = "0.0%" if rr_ == 9 else "0.00"
    Q.cell(11, 1, "營收：全年 FY2025（參考）"); Q.cell(11, 2, 538.6).font = BLUE; Q.cell(11, 5, "Verified"); Q.cell(11, 6, Q4R)
    Q.cell(12, 1, "營收：Q4'25"); Q.cell(12, 2, 179.7).font = BLUE; Q.cell(12, 5, "Verified"); Q.cell(12, 6, Q4R)
    Q.cell(13, 1, "營收：9M'26（TTM 729.6 − Q4'25；Derived）"); Q.cell(13, 3, "=729.6-B12"); Q.cell(13, 5, "Derived"); Q.cell(13, 6, SAI)
    Q.cell(14, 1, "注：公司揭露系統與軟體服務、地區拆分的金額，本輪未取得（搜尋摘要未含）；未建模分部")
    Q.column_dimensions["A"].width = 46

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：線 5 AS/RS 市場 × AutoStore 份額 → 營收 → 標準損益表（IFRS）→ 現金流（US$M；FY2023–25 實績、FY2026 公司指引 ×(1＋g_anc)、FY2027 起模型）"
    PR = {}
    rr = 3
    KEYS = ["mkt", "shr", "rev", "chk", "gr", "cogs", "gp", "gm", "opex", "op", "opm", "fin", "pbt", "tax", "ni", "nim", "sh", "eps", "epsn", "nig", "da", "cx", "dwc", "fcff", "div", "fcfad"]
    LAB = ["參考：AS/RS 市場（US$M，線 5 下游取數）", "AutoStore 占 AS/RS 市場份額（出廠價口徑）", "營收（合併）", "檢查：市場 × 份額 − 營收（FY2026 起）",
           "營收年增率", "營業成本", "毛利", "毛利率", "營業費用（毛利 − 營業利益）", "營業利益", "營業利益率", "業外：財務收益 − 財務費用", "稅前利益",
           "所得稅", "稅後淨利", "淨利率", "平均股數（百萬股）", "EPS（美元）", "EPS（挪威克朗）", "淨利年增率", "折舊攤銷", "資本支出（含資本化研發）", "營運資金增加",
           "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "股利", "自由現金流扣股利（收支關卡用）"]
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, k in enumerate(KEYS)}
        for i, l in enumerate(LAB): P.cell(b + i, 1, l)
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1); cell = lambda k: f"{c}{R[k]}"; prev = lambda k: f"{p}{R[k]}"
            def w(k, f): P[cell(k)] = f
            if y <= 2025:
                yy = str(y)[2:]
                for k, hk in (("rev", "rev"), ("cogs", "cogs"), ("gp", "gp"), ("opex", "opex"), ("op", "op"), ("fin", "fin"), ("pbt", "pbt"), ("tax", "tax"), ("ni", "ni"), ("sh", "sh")):
                    w(k, f"=h_{hk}_{yy}")
                w("gm", f"={cell('gp')}/{cell('rev')}"); w("opm", f"={cell('op')}/{cell('rev')}"); w("nim", f"={cell('ni')}/{cell('rev')}")
                w("eps", f"={cell('ni')}/{cell('sh')}"); w("epsn", f"={cell('eps')}*fx")
                if y == 2023: w("gr", '="未取得（FY2022 未收入）"'); w("nig", '="N/M（FY2023 為虧損）"')
                else:
                    w("gr", f"={cell('rev')}/{prev('rev')}-1")
                    w("nig", f"=IF({prev('ni')}>0,{cell('ni')}/{prev('ni')}-1,\"N/M（前一年為虧損）\")")
                continue
            if y == 2026:
                w("rev", "=g26_rev*(1+g_anc)"); w("gp", f"={cell('rev')}*gm26"); w("op", f"={cell('rev')}*em26"); w("fin", "=fin26")
                w("sh", "=sh_v"); w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}")
                w("mkt", f"=m26_{s}*1000"); w("shr", f"={cell('rev')}/{cell('mkt')}")
            else:
                gk = None
                if y <= 2028: mk = f"=m26_{s}*(m28_{s}/m26_{s})^(({y}-2026)/2)*1000"
                elif y <= 2030: mk = f"=m28_{s}*(m30_{s}/m28_{s})^(({y}-2028)/2)*1000"
                else: mk = f"=m30_{s}*(m32_{s}/m30_{s})^(({y}-2030)/2)*1000"
                w("mkt", mk); w("shr", f"={prev('shr')}+dshr_{s}"); w("rev", f"={cell('mkt')}*{cell('shr')}")
                w("gp", f"={cell('rev')}*gm_{s}"); w("opex", f"={prev('opex')}*(1+ox_inf+ox_el*{cell('gr')})")
                w("op", f"={cell('gp')}-{cell('opex')}"); w("fin", "=nonop"); w("sh", "=sh_v")
                w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}")
            if y == 2026: w("opex", f"={cell('gp')}-{cell('op')}"); w("cogs", f"={cell('rev')}-{cell('gp')}")
            else: w("cogs", f"={cell('rev')}-{cell('gp')}")
            w("chk", f"=ROUND({cell('mkt')}*{cell('shr')}-{cell('rev')},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            w("gm", f"={cell('gp')}/{cell('rev')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            w("pbt", f"={cell('op')}+{cell('fin')}"); w("tax", f"={cell('pbt')}*tax_w"); w("ni", f"={cell('pbt')}-{cell('tax')}")
            w("nim", f"={cell('ni')}/{cell('rev')}"); w("eps", f"={cell('ni')}/{cell('sh')}"); w("epsn", f"={cell('eps')}*fx")
            w("nig", f"=IF({prev('ni')}>0,{cell('ni')}/{prev('ni')}-1,\"N/M（前一年為虧損）\")")
            w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
            w("fcff", f"={cell('op')}*(1-tax_w)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("div", f"=dps*{cell('sh')}"); w("fcfad", f"={cell('fcff')}-{cell('div')}")
        P.cell(R["rev"], 2, "FY2026＝公司指引；FY2027 起＝市場 × 份額")
        P.cell(R["shr"], 2, "FY2026＝營收 ÷ 市場；之後每年 ＋ dshr")
        P.cell(R["gp"], 2, "FY2026＝營收 × gm26；FY2027 起＝營收 × 情境毛利率")
        P.cell(R["opex"], 2, "FY2027 起＝前一年 ×(1＋基本年增＋彈性 × 營收年增)")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 ×(1 − tax_w)；股權激勵未取得（不加回）")
        for k in ("gr", "gm", "opm", "nim", "nig", "shr"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(LAB) + 2
    P.column_dimensions["A"].width = 52; P.column_dimensions["B"].width = 30
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "fcfad"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現（美元），評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)；每股美元值 × 匯率 ＝ 挪威克朗"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（US$M）"; D["B4"] = "=px/fx*sh_v"
    D["A5"] = "負債 D（淨負債）"; D["B5"] = "=nd_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金（負為淨負債；2026-09-30）＝ − 淨負債"; D["B10"] = "=-nd_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND((((px/fx*sh_v)*(rf+beta*(erp+crp))+nd_v*kd*(1-tax_w))/((px/fx*sh_v)+nd_v))+wacc_adj-wacc,9)"
    for rr_ in range(3, 11): D[f"B{rr_}"].number_format = "0.00%" if rr_ in (3, 6, 7, 8, 9) else "#,##0.0"
    DR = {}; rr = 13
    for s in SC:
        R = PR[s]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}" for y in range(2027, 2033)]); b = rr + 2
        rows = ["EBIT（營業利益）", "NOPAT", "＋折舊攤銷", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF（應為 0）", "折現年數（期中）", "折現因子", "FCFF 現值"]
        for i, l in enumerate(rows): D.cell(b + i, 1, l)
        for j, y in enumerate(range(2027, 2033)):
            c = "BCDEFG"[j]; pc = COL[y]
            D[f"{c}{b}"] = f"=PL!{pc}{R['op']}"; D[f"{c}{b+1}"] = f"={c}{b}*(1-tax_w)"
            D[f"{c}{b+2}"] = f"=PL!{pc}{R['da']}"; D[f"{c}{b+3}"] = f"=-PL!{pc}{R['cx']}"; D[f"{c}{b+4}"] = f"=-PL!{pc}{R['dwc']}"
            D[f"{c}{b+5}"] = f"=PL!{pc}{R['fcff']}"; D[f"{c}{b+6}"] = f"=ROUND(SUM({c}{b+1}:{c}{b+4})-{c}{b+5},9)"
            D[f"{c}{b+7}"] = j + 0.5; D[f"{c}{b+8}"] = f"=1/(1+wacc)^{c}{b+7}"; D[f"{c}{b+9}"] = f"={c}{b+5}*{c}{b+8}"
        t = b + 11
        items = [("明確期現值合計", f"=SUM(B{b+9}:G{b+9})"),
                 ("NOPAT FY2033＝NOPAT FY2032 ×(1＋g)", f"=G{b+1}*(1+g_tv)"),
                 ("終值（FY2032 年底）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^6"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金（2026-09-30）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!F{R['fcff']}*stub_frac"),
                 ("− 尚待買回金額（無）", 0),
                 ("股權價值", f"=B{t+4}+B{t+5}+B{t+6}+B{t+7}"),
                 ("每股價值（評價日，NOK）", f"=MAX(0,B{t+8})/sh_v*fx"),
                 ("DCF 情境價值推到 12 個月後（NOK；×(1＋ke)^t − 期間股利）", f"=(B{t+9}/fx*(1+ke)^t_tgt-dps*t_tgt)*fx"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+11}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=B{t+2}/G{b+1}"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（NOK，推到 12 個月後）", f"=(((B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6)+B{t+5}+B{t+6}+B{t+7})/sh_v*(1+ke)^t_tgt-dps*t_tgt)*fx")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 10, ps=t + 9, ev=t + 4, tvs=t + 11, alt=t + 14, chk=b + 6)
        dn(f"dcf_{s}", f"DCF!$B${t+10}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 60

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價（NOK）：DCF（權重 w_dcf）＋ Comparable（同業前瞻 P/E 中位數 × 情境 FY2028 EPS，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}")
        V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "DCF 頁")
    r_cu = vrow("Comparable（未調整）：同業前瞻 P/E 中位數 × 情境 FY2028 EPS", [f"=pe_med*PL!H{PR[s]['epsn']}" for s in SC],
                "目標時點的 NTM 以 FY2028 代；分母：模型各情境 FY2028 EPS（NOK）")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（挪威克朗）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    r_px = one("現價（In 頁 px）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（US$M；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=IF(B{r_up}>0.5,"已在最高級",TEXT(px*(1+IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5)))),"#,##0.00"))&"／"&IF(B{r_up}<-0.5,"已在最低級",TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,-0.5)))),"#,##0.00"))', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 EPS（NOK）", f"=B{r_tp}/PL!H{PR['B']['epsn']}", "分母：模型情境 B FY2028 EPS")
    r_pe1 = one("現價 ÷ 情境 B FY2027 EPS（NOK）", f"=px/PL!G{PR['B']['epsn']}", "分母：模型情境 B FY2027")
    r_pe2 = one("現行前瞻 P/E（資料商 NTM，換算至 2026-11-20）÷ 同業中位數", "=pe_self/pe_med", "只作對照；相對重估需要的倍數＝1 ÷ 本值")
    r_shr = one("情境 B FY2030 AutoStore 占 AS/RS 市場份額", f"=PL!J{PR['B']['shr']}")
    for rr_ in (r_up, r_a1, r_a2, r_a3, r_shr): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0.00"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "#,##0.00"
    V.column_dimensions["A"].width = 70; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, ca=r_ca, pd2=r_pd2, pe=r_pe, pe1=r_pe1, pe2=r_pe2, gj=r_gj, wk=r_wk, shr=r_shr)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司指引、已揭露實績對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或實績", "差", "說明"])
    rc = [("FY2026 營收", f"=PL!F{PR['B']['rev']}", "=g26_rev", "公司指引（Interested-party）"),
          ("FY2026 營收 vs 9M'26 ＋ Q3 年化外推（9M 549.9 ＋ Q3 192.1）", f"=PL!F{PR['B']['rev']}", "=Q!C13+Q!C3", "Q4 若與 Q3 持平的全年營收；指引高於此表示公司預期 Q4 續增"),
          ("FY2026 營業利益 vs TTM EBIT（資料商）", f"=PL!F{PR['B']['op']}", 232.3, "TTM 至 2026-09-30，" + SAI),
          ("FY2027 EPS（模型 B，NOK） vs 現價 ÷ 現行前瞻 P/E 隱含 NTM EPS", f"=PL!G{PR['B']['epsn']}", "=px/pe_self", "口徑：資料商 NTM（2026-10-22 頁面）vs 模型 FY2027；口徑不同，只看方向"),
          ("FY2026 份額（營收 ÷ 線 5 市場）vs 線模型 WH_AUTOSTORE_SHARE", f"=PL!F{PR['B']['shr']}", 0.0701305731703361, "線 5 下游取數!D27（0.0701）；口徑一致"),
          ("加權目標價 vs 賣方平均目標價", "=tp", "未取得", "本輪未找到 AutoStore 賣方共識目標價，不作對照")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=IF(ISNUMBER(C{i}),B{i}-C{i},\"未取得\")"); RC.cell(i, 5, nu(n))
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_autostore.py 批次格點＋線性內插；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 60; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_autostore.py 以 LibreOffice 批次重算每一案例）；結果為常數快照"
    base_tp = sens["base"]["tp"] if sens else None
    SE["A2"] = "快照基準值（加權，NOK）"; SE["B2"] = base_tp
    SE["C2"] = "本檔即時值（加權，NOK）"; SE["D2"] = "=tp"
    SE["E2"] = f"快照輸入指紋 {sens['fp'] if sens else ''}"
    SE["A3"] = "build_id"; SE["D3"] = bid
    SE["A4"] = "模型狀態"; SE["D4"] = '=IF(ABS(D2-B2)>0.005,"STALE","CURRENT")'
    hdr(SE, 6, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    sr = 7
    if sens:
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]); SE.cell(sr, 5, cs["tp"] / base_tp - 1)
            SE.cell(sr, 6, cs["dcf"]); SE.cell(sr, 7, cs["dcf"] / sens["base"]["dcf"] - 1); SE.cell(sr, 8, cs["rt"])
            SE[f"E{sr}"].number_format = "0.0%"; SE[f"G{sr}"].number_format = "0.0%"; SE[f"D{sr}"].number_format = "#,##0.00"; SE[f"F{sr}"].number_format = "#,##0.00"
            sr += 1
    SE.cell(sr + 1, 1, "說明：相對倍數係數 rel_k 基準即為 1.00，故不列案例。市場規模（m26–m32）的低／高已作情境 C／A，不另設區間；沒有區間的鍵不列案例。")
    SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 34

    # ---------------- 對帳 ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = "對帳 — 本模型使用的線層級變數（讀線 5 模型成品「下游取數」頁；與線模型不一致的自設值標明理由，不靜默改口徑）"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型一致？"])
    rec = [("AS/RS 市場 2026（低／基準／高）", "5", "10.45／10.69／10.93", "US$B", "MarketsandMarkets 口徑（含製造、汽車）", WH + " 下游取數!C8:E8", "Derived", "In m26_C／m26_B／m26_A → PL 市場列", "一致（照取）"),
           ("AS/RS 市場 2028（低／基準／高）", "5", "11.75／12.58／13.43", "US$B", "同上", WH + " 下游取數!C9:E9", "Derived", "In m28_*", "一致"),
           ("AS/RS 市場 2030（低／基準／高）", "5", "13.2／14.8／16.5", "US$B", "同上", WH + " 下游取數!C10:E10", "Derived", "In m30_*", "一致"),
           ("AS/RS 市場 2032（低／基準／高）", "5", "14.28／16.94／19.60", "US$B", "同上", WH + " 下游取數!C11:E11", "Derived", "In m32_*", "一致"),
           ("AutoStore 占 AS/RS 市場 2026", "5", "7.0%（FY2026 營收 750 ÷ 10,694）", "比例", "出廠價口徑，低估系統占比", WH + " 下游取數!D27", "Derived", "PL 份額列 FY2026；Recon", "一致（線模型 0.0701）"),
           ("AS/RS 立方格回本期（穩態；悲觀／基準／樂觀）", "5", "20.05／2.16／0.71", "年", "美國 2026，每埠＝每座倉", WH + " 下游取數!C14:E14", "Derived", "dshr_A／dshr_C 說明（需求面證據，未進公式）", "一致；區間極寬，承重變數為假設"),
           ("每行成本：AS/RS 0.197、AMR 0.205 美元", "5", "0.197／0.205", "USD/訂單行", "基準；兩者接近", WH + " 下游取數!D17、D16", "Derived", "dshr_C 理由（AMR 替代）", "一致"),
           ("極智嘉占行動機器人營收 2025", "5", "7.2%", "比例", "量級對照", WH + " 下游取數!D26", "Derived", "未進公式", "未使用"),
           ("線 5 追蹤指標：Q3'26 營收 1.921 億、接單 1.776 億", "5", "192.1／177.6", "US$M", "公司公告", WH + " 追蹤指標!D9（本專案公司檔案引用）", "Verified", "Q 頁", "一致（本模型另自公告核對）")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 30; RE.column_dimensions["F"].width = 60; RE.column_dimensions["I"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("倉儲自動化資本支出循環回落（接單先行）", "是（情境 C：低市場＋份額下滑）", "營收由市場 × 份額推導", "季度接單、接單 ÷ 營收、整合商（KION、Swisslog、Kardex）訂單"),
             ("AMR 貨到人與其他方案取代立方格", "是（dshr_C）", "線 5 每行成本兩者接近", "AMR 廠商價格、競品份額"),
             ("大股東（軟銀 38%、THL 25%）減持造成股價壓力", "否", "不影響內在價值，但影響短期股價", "大額持股公告"),
             ("關稅與毛利率壓力", "部分（gm_C 70%）", "Q3'26 毛利率年減 1.8 個百分點", "季度毛利率"),
             ("Logiqs 收購整合與其他併購", "否", "金額最高 3,000 萬歐元，相對市值小", "公司公告"),
             ("美元兌挪威克朗", "部分（fx 區間）", "目標價以美元估值乘匯率換算", "NOK=X"),
             ("股權激勵與稀釋", "否", "金額未取得，FCFF 未加回（偏保守）", "年報薪酬揭露")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 50

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([Q3R, Q4R, SAI, CALL, PRO, YCL, UST, SAS, SYMX, WH + "（下游取數、追蹤指標）"], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"AutoStore（AUTO.OL）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-20（第 3 季公告 2026-11-05）；成品日 2026-11-24；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"#,##0.00")&" 挪威克朗）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：AutoStore 是立方格 AS/RS 的核心設備與軟體供應商，毛利率 70% 以上、資本支出輕；股價由倉儲自動化資本支出循環（市場規模）與其市場份額，加上高折現率（高 β）下的現金流決定"
    RD["A4"] = (f'="在線 5 AS/RS 市場基準成長、份額持平、毛利率 "&TEXT(gm_B,"0%")&"、同業前瞻 P/E 中位數 "&TEXT(pe_med,"0.0")&" 倍的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"#,##0.00")&" 挪威克朗（DCF "&TEXT(Val!B{r_dcft},"#,##0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"#,##0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0")&" 百萬美元）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 挪威克朗。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：季度接單連兩季低於營收的 80%；季度毛利率跌破 68%；FY2027 公司指引營收低於 FY2026 指引；線 5 AS/RS 市場基準下修"
    RD["A7"] = "範圍：營收只用市場 × 份額推導（公司未在可取得公告中揭露系統與軟體服務拆分，未建分部）；不讀 Siliconomics／Tokenomics／Wattonomics（算力線）；自用 AI 讀法等規格條文不適用。HTML 展示層見 dist/ 同名 .html"
    RD["A8"] = f'="命題回覆：問：市場成長加高毛利能否支持高於現價的目標價？答：情境 B FY2028 EPS "&TEXT(PL!H{PR["B"]["epsn"]},"0.00")&" 挪威克朗，現價為其 "&TEXT(px/PL!H{PR["B"]["epsn"]},"0.0")&" 倍，高於同業中位數 "&TEXT(pe_med,"0.0")&" 倍；加權目標價 "&TEXT(Val!B{r_tp},"#,##0.00")&" 挪威克朗。問：份額多重要？答：情境 B 份額每年持平；情境 A（每年 +0.3 個百分點）與情境 C（每年 −0.4 個百分點）的差距見 Sens 頁（模型輸出，非獨立證據）。"'
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
def lo_profile():
    d = os.path.join(os.environ.get("HOME", tempfile.gettempdir()), "lo_autostore")
    os.makedirs(d, exist_ok=True)
    return "file://" + d

def recalc_many(srcs, outdir):
    os.makedirs(outdir, exist_ok=True)
    for i in range(0, len(srcs), 40):
        chunk = srcs[i:i + 40]
        for attempt in range(2):
            try:
                subprocess.run(["soffice", f"-env:UserInstallation={lo_profile()}", "--headless", "--convert-to", "xlsx", "--outdir", outdir] + chunk,
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
                break
            except subprocess.TimeoutExpired:
                subprocess.run(["pkill", "-f", "soffice"]); continue
    return [os.path.join(outdir, os.path.basename(s)) for s in srcs]

def read_vals(path, VR, DR=None):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]
    d = dict(tp=V[f"B{VR['tp']}"].value, dcf=V[f"B{VR['dcf']}"].value, cmp=V[f"B{VR['cmp']}"].value, up=V[f"B{VR['up']}"].value, rt=V[f"B{VR['rt']}"].value)
    if DR:
        I = wb["In"]; w = {}; wd = None
        for row in I.iter_rows(values_only=True):
            if row[0] in ("w_A", "w_B", "w_C"): w[row[0][-1]] = row[4]
            if row[0] == "w_dcf": wd = row[4]
        alt = sum(w[s] * wb["DCF"][f"B{DR[s]['alt']}"].value for s in SC)
        d["alt_dcf"] = alt; d["alt_tp"] = wd * alt + (1 - wd) * d["cmp"]
    return d

def run_cases(ovs):
    d = tempfile.mkdtemp(prefix="as_"); paths = []; VR = DR = None
    for tag, ov in ovs:
        p = os.path.join(d, f"c_{tag}.xlsx"); VR, _, DR = build(p, ov); paths.append(p)
    outs = recalc_many(paths, os.path.join(d, "o"))
    res = {tag: read_vals(o, VR, DR if tag == "base" else None) for (tag, _), o in zip(ovs, outs)}
    shutil.rmtree(d, ignore_errors=True); return res

def in_val(k, col):
    for row in ROWS:
        if row[0] == k:
            return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)

def fingerprint():
    return hashlib.sha1(json.dumps([r for r in ROWS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]

def rating(u):
    return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"

CASES = [("g_anc", "FY2026 實績相對公司指引"), ("gm26", "FY2026 毛利率"), ("em26", "FY2026 營業利益率"), ("fin26", "FY2026 業外"), ("fx", "美元兌挪威克朗"),
         ("dshr_B", "AutoStore 份額每年變動（情境 B）"), ("gm_B", "FY2027 起毛利率（情境 B）"), ("ox_inf", "營業費用基本年增"), ("ox_el", "營業費用彈性"), ("nonop", "業外財務損益"),
         ("tax_w", "有效稅率"), ("fc_da", "折舊 ÷ 營收"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("nd_v", "淨負債"),
         ("beta", "β"), ("erp", "股權風險溢酬"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重"), ("pe_med", "同業前瞻 P/E 中位數")]
BETA_B = statistics.median([0.67 * 1.98 + 0.33, 0.67 * 1.03 + 0.33, 0.67 * 2.05 + 0.33])
FXD = {"pe_med": statistics.median([21.33 * 4129 / 5429, 9.85 * 29.2 / 40.58, 56.32 * 5.405 / 9.22])}


def solve_batch(key, lo, hi, targets, px0):
    import numpy as np
    xs = list(np.linspace(lo, hi, 13))
    res = run_cases([(f"g{i}", {key: x}) for i, x in enumerate(xs)])
    ys = [res[f"g{i}"]["up"] for i in range(len(xs))]
    out = {}
    for tgt in targets:
        if tgt is None: continue
        br = next((i for i in range(len(xs) - 1) if (ys[i] - tgt) * (ys[i + 1] - tgt) <= 0), None)
        if br is None: out[tgt] = None; continue
        a, b_ = xs[br], xs[br + 1]
        xs2 = list(np.linspace(a, b_, 7)); r2 = run_cases([(f"f{i}", {key: x}) for i, x in enumerate(xs2)]); y2 = [r2[f"f{i}"]["up"] for i in range(7)]
        j = next(i for i in range(6) if (y2[i] - tgt) * (y2[i + 1] - tgt) <= 0)
        x0, x1, y0, y1 = xs2[j], xs2[j + 1], y2[j], y2[j + 1]
        out[tgt] = x0 if y1 == y0 else x0 + (tgt - y0) * (x1 - x0) / (y1 - y0)
    return out


def main():
    px0 = in_val("px", "base")
    jobs = [("base", {})]; meta = []
    for k, drv in CASES:
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if k == "beta": v = BETA_B + (-0.3 if side == "lo" else 0.3)
            if v is None: continue
            b_ = in_val(k, "base")
            if not isinstance(b_, (int, float)): b_ = FXD.get(k, BETA_B if k == "beta" else None)
            if b_ is not None and abs(v - b_) < 1e-12: continue
            tag = f"{k}_{side}"; jobs.append((tag, {k: v})); meta.append((tag, drv + ("（低）" if side == "lo" else "（高）"), f"{k}＝{v:.4g}", k))
    for s in SC:
        tag = f"only_{s}"; jobs.append((tag, {f"w_{x}": (1.0 if x == s else 0.0) for x in SC})); meta.append((tag, f"只看情境 {s}", f"w_{s}＝1，其餘 0", f"w_{s}"))
    res = run_cases(jobs)
    base = res["base"]; print("基準", base)
    cases = [dict(drv=d, set=st, key=k, **res[t]) for t, d, st, k in meta]
    up_ = base["alt_tp"] / px0 - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rating(up_)))
    for c in cases: print(c["key"], c["drv"], round(c["tp"], 2), c["rt"])
    up0 = base["up"]; tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    revs = []
    for key, b0, lo_, hi_, note in (("dshr_B", 0.0, -0.02, 0.05, "情境 B 的 AutoStore 份額每年變動（其他不變，整本重算；情境 A／C 不動）"),
                                    ("beta", BETA_B, 0.3, 3.0, "β（Blume 調整後；三情境同動；影響 DCF 的 ke 與 WACC）"),
                                    ("pe_med", FXD["pe_med"], 0.0, 60.0, "同業前瞻 P/E 中位數（Comparable；DCF 不受影響）")):
        sol = solve_batch(key, lo_, hi_, [0.0, upper, lower], px0)
        for lab, tgt in (("目標價＝現價（上檔 0%）", 0.0), (f"升一級邊界（上檔 {upper:+.0%}）" if upper is not None else None, upper), (f"降一級邊界（上檔 {lower:+.0%}）" if lower is not None else None, lower)):
            if lab is None: continue
            x = sol.get(tgt)
            revs.append(dict(label=lab, key=key, val=x if x is not None else f"區間 {lo_:+.2f}～{hi_:+.2f} 內無解", base=b0, note=note))
            print("反解", key, lab, x)
    sens = dict(base=base, cases=cases, fp=fingerprint())
    json.dump(sens, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="asf_"); raw = os.path.join(tmp, NAME)
    bid = f"{VER}-{fingerprint()}"
    VR, PR, DR = build(raw, {}, sens, revs, bid)
    out = recalc_many([raw], os.path.join(tmp, "o"))[0]
    shutil.copy(out, os.path.join(dist, NAME))
    wbv = openpyxl.load_workbook(os.path.join(dist, NAME), data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10]); assert not errs, errs
    print("Sens 狀態", wbv["Sens"]["D4"].value, "tp", wbv["Val"][f"B{VR['tp']}"].value)
    write_readme(os.path.join(dist, NAME), VR, PR, DR, bid, sens, revs)
    json.dump(dict(company="autostore", ticker="AUTO.OL", model_sheet="PL",
                   official_segments=[dict(name="合併營收（單一報告分部）", label="營收（合併）", period="FY2025", value=0.5386, unit="US$B", source=Q4R)],
                   sublines=[], check_labels=["檢查：市場 × 份額 − 營收（FY2026 起）"],
                   note="公司為單一報告分部；可取得公告未拆分系統與軟體服務營收，故以合併營收為唯一分部列，PL 頁另以「市場 × 份額 − 營收」核對推導一致（FY2026 起）。"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)


def write_readme(path, VR, PR, DR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f2 = lambda x: f"{x:,.2f}"; pc = lambda x: f"{x*100:+.1f}%"
    inv = lambda k: I['E' + str([r for r in range(1, I.max_row + 1) if I['A' + str(r)].value == k][0])].value
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:8]
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {r['val'] if isinstance(r['val'], str) else format(r['val'], '.4f')} | {r['base']:.4g} | {r['note']} |" for r in revs)
    def yr(s, k, fmt="{:,.1f}", ys=range(2023, 2033)):
        out = []
        for y in ys:
            v = P[f"{COL[y]}{PR[s][k]}"].value
            out.append(fmt.format(v) if isinstance(v, (int, float)) else "—")
        return " | ".join(out)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f2(g(chr(66+i)+str(VR['dcfr'])))} | {f2(g(chr(66+i)+str(VR['cu'])))} | {f2(g(chr(66+i)+str(VR['tps'])))} | {g(chr(66+i)+str(VR['gate'])):,.0f} |" for i, s in enumerate(SC))
    rts = sorted(set(c['rt'] for c in sens['cases']))
    YS = range(2023, 2033)
    txt = f"""# physical_ai/autostore — AutoStore（AUTO.OL）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

奧斯陸證券交易所上市，以美元報告、挪威克朗交易。本 README 的數字由 `build_autostore.py` 從重算後的成品讀出產生，不手寫。

## 未完成與限制（先讀）
- **未建分部**：公司為單一報告分部；可取得的公告摘要沒有系統與軟體服務、地區的營收金額，營收只用「線 5 AS/RS 市場 × AutoStore 份額」推導。
- **未取得**：FY2025 全年與 9M'26 的完整損益、現金流量表與資本支出金額（資本支出由公司「現金流轉換率」反推，Derived）；Q3'26 淨負債 48 百萬美元只有二手財報會議摘要；賣方共識目標價與共識 EPS；股權激勵金額；Kardex 同業報價（收盤檔無）。
- 同業前瞻 P/E 只有 3 家（Daifuku、KION、極智嘉；報價一律換算到 2026-11-20 收盤），家數少、離散大（7.1–33 倍）；不做迴歸。
- 2026-11-16 至 20 全球 AI 與半導體股大跌，收盤檔的下跌不是資料錯誤。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A8'].value}

**讀法：** 估值偏低主要來自兩件事：(1) 折現率高：β 取 Blume 調整後 {inv('beta'):.2f}（AutoStore、KION 5 年 β 約 2）加美國 10 年期公債 5.10%，ke 約 {g(f"B{VR['wk']}").split('／')[1]}，WACC {g(f"B{VR['wk']}").split('／')[0]}；(2) Comparable 的同業中位數只有 {inv('pe_med'):.1f} 倍，低於 AutoStore 現行 {inv('pe_self'):.1f} 倍。現價為模型情境 B FY2027 EPS 的 {g(f"B{VR['pe1']}"):.1f} 倍。結論主要建立在 Derived 的市場規模（線 5）、Assumed 的份額與毛利率、以及 Analogy 的高 β 與少數同業倍數上。

## 估值（評價日 2026-12-31；目標時點約 12 個月；每股為挪威克朗）
| 情境 | DCF（推到 12 個月後） | Comparable（P/E × FY2028 EPS） | 情境目標價 | 收支關卡（FY2026–30 FCF 扣股利，US$M） |
|---|---|---|---|---|
{scn}

- 加權目標價 **{f2(tp)} 挪威克朗**（DCF {f2(dcf)}、Comparable {f2(cmp_)}），現價 {inv('px'):.2f}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")} → 投資建議 **{g(f"B{VR['rec']}")}**。
- 方法：DCF 標準版 FCFF（{g(f"B{VR['wk']}")}，為 WACC／ke／g／RONIC）50%＋Comparable 50%（同業 3 家前瞻 P/E 中位數 {inv('pe_med'):.2f} 倍 × 各情境 FY2028 EPS）。相對倍數係數 1.00（未調整）。
- 隱含前瞻 P/E（加權目標價 ÷ 情境 B FY2028 EPS）{g(f"B{VR['pe']}"):.1f} 倍；AutoStore 現行前瞻 P/E {inv('pe_self'):.1f} 倍 ＝ 同業中位數的 {g(f"B{VR['pe2']}"):.2f} 倍。

## 營收與標準損益（情境 B，US$M；FY2023–25 實績、FY2026 公司指引、FY2027 起模型）
| 列 | {" | ".join(f"FY{y}" for y in YS)} |
|---|{"---|" * len(YS)}
| AS/RS 市場（線 5） | {yr('B','mkt','{:,.0f}')} |
| AutoStore 份額 | {yr('B','shr','{:.2%}')} |
| 營收 | {yr('B','rev')} |
| 毛利 | {yr('B','gp')} |
| 營業利益 | {yr('B','op')} |
| 淨利 | {yr('B','ni')} |
| EPS（挪威克朗） | {yr('B','epsn','{:.2f}')} |

- FY2023 稅前含 2.39 億美元和解費用（營業利益以下），淨利為負；FY2025 EBIT 資料商 142.1、公司公告 140.1，差 2.0 未查核。

## 主要驅動因子與標記
1. AS/RS 市場規模（線 5 下游取數，Derived）：2026 年 10.69 十億美元、2030 年 14.8、2032 年 16.9（情境 A／C 取線模型高／低）。
2. AutoStore 份額與毛利率（Assumed）：FY2026 份額 7.0%（營收 ÷ 市場）；情境每年 +0.3／0／−0.4 個百分點；毛利率 73%／72%／70%。
3. 折現率與同業倍數（Analogy／Derived）：β {inv('beta'):.2f}、同業 P/E 中位數 {inv('pe_med'):.1f} 倍。

## 失效條件
{RD['A6'].value.split('：', 1)[1]}。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f2(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁；各案例估值評等出現「{'、'.join(rts)}」。

## 反解（整本重算；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
- 線 5：AS/RS 市場 2026／2028／2030／2032 低／基準／高（下游取數!C8:E11）照取；FY2026 份額 7.0% 與線模型 WH_AUTOSTORE_SHARE（下游取數!D27，0.0701）一致。
- 線模型的回本期、每行成本只作份額情境（dshr）的方向證據，未進公式。無與線模型不一致之處。

## 檔案
- `build_autostore.py`：建置＋整本重算 Sens／反解（LibreOffice 批次）＋產生本 README、`segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-24）：初版；現價 {inv('px'):.2f}（2026-11-20 收盤）；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


if __name__ == "__main__":
    main()

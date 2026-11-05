# Arista Networks（ANET）收支與估值模型 v0.2.1（AI 半導體 Project：網通系統端樣板）
# 用法：python3 arista/build_arista.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不做平移。金額：十億美元（US$B）；EPS：美元；口徑：公司 non-GAAP（排除股票薪酬、攤銷等；
# 2026 年起 non-GAAP 稅額改用新方法，公司已重編 2025 年各季，本模型 FY2025 用 2025 年度新聞稿的原報值，見 Q 頁說明）。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "siliconomics")); import reader as SIL  # v0.12 個股自訂情境說明
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))   # 本資料夾的 tp_v2.py
import cycle_position as CP
import tp_v2 as TP
import fcff_dcf as FD   # v0.5：標準版 DCF（無槓桿自由現金流＋WACC）共用模組
import valuation_rule as VR

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")

# ---------------- 讀入其他模型的成品（唯讀；不修改它們） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
NET_F = only("markets/network/dist/*.xlsx"); TSMC_F = only("tsmc/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
wsn = openpyxl.load_workbook(NET_F, data_only=True)
NETR = {}
for r in wsn["Mkt"].iter_rows(min_row=4):
    lab = r[0].value
    if lab in ("scale-out 交換器", "AI IT 資本支出（加速器 × 每顆 IT capex）", "NVIDIA 網通（本模型歸屬）", "AI 網通市場合計"):
        NETR[lab] = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
assert len(NETR) == 4 and all(v is not None for v in NETR["scale-out 交換器"][0]), NETR.keys()
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and str(k).startswith("um"):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
assert all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億美元（US$B）；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文＝Arista 8-K 新聞稿（SEC EDGAR）；公司的 AI 營收目標與全年指引為 Interested-party"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
EDG = "https://www.sec.gov/Archives/edgar/data/0001596532/"
PR_Q226 = EDG + "000159653226000174/ex991q226-earningsrelease.htm"
PR_Q326 = EDG + "000159653226000252/ex991q326-earningsrelease.htm"   # v0.4：3Q26 實績與 4Q26 指引（2026-11-03）
PR_Q425 = "https://www.businesswire.com/news/home/20260212962784/en/Arista-Networks-Inc.-Reports-Fourth-Quarter-and-Year-End-2025-Financial-Results"
CALL = "法說轉述：BigGo Finance 摘要（2026-08-04；二手）"

sec("A. AI 後端 scale-out 交換器市場（讀 markets/network 成品；情境 A（原始顆數），B／C 依顆數係數縮放）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"mkt_{y}", f"AI 後端 scale-out 交換器市場 20{y}", "US$B", None, NETR["scale-out 交換器"][0][j], None, "Derived", NETR["scale-out 交換器"][1],
        "加速器顆數 × 每顆 scale-out 交換器內容（不分乙太網路與 InfiniBand）；只含 AI 後端" if y == 26 else "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}", "US$B", None, NETR["AI IT 資本支出（加速器 × 每顆 IT capex）"][0][j], None, "Derived", NETR["AI IT 資本支出（加速器 × 每顆 IT capex）"][1], "Recon 用：Arista AI 營收 ÷ AI IT 資本支出" if y == 26 else "同上")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}（個股自訂）：2028 顆數 ÷ 原始顆數", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], SIL.scen_note(s, 28, TS[f"um28_{s}"][0]))
    inp(f"um29_{s}", f"情境 {s}（個股自訂）：2029 顆數 ÷ 原始顆數", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], SIL.scen_note(s, 29, TS[f"um29_{s}"][0]))

sec("B. 乙太網路占比與 Arista 份額（本專案為權威：晶片商與網通廠份額，規則第 13 條）")
inp("eth_26", "乙太網路 ÷ AI 後端交換器（金額）2026", "%", 0.70, 0.75, 0.80, "Assumed", "Dell'Oro 2026-03-10：2025 年乙太網路占 AI 叢集交換器三分之二以上、為 InfiniBand 兩倍以上（Interested-party）", "IDC 2Q26：NVIDIA 資料中心乙太網路年增 181%（乙太網路取代 InfiniBand 的方向）")
inp("eth_27", "乙太網路 ÷ AI 後端交換器 2027", "%", 0.72, 0.80, 0.88, "Assumed", "同上；2027 年無公開預測：未找到", "外推（Assumed）")
inp("eth_28", "乙太網路 ÷ AI 後端交換器 2028", "%", 0.74, 0.83, 0.90, "Assumed", "同上", "—")
inp("eth_29", "乙太網路 ÷ AI 後端交換器 2029", "%", 0.75, 0.85, 0.92, "Assumed", "同上", "—")
inp("rel_27", "Arista 份額 2027 ÷ 2026（相對值）", "x", 0.80, 0.95, 1.05, "Assumed", "IDC 2Q26：NVIDIA 20.4% 已超過 Arista 18.7%（資料中心乙太網路，全部速率；Interested-party）", "Dell'Oro 2025 AI 後端乙太網路：Celestica（白牌）與 NVIDIA 並列第一、合計約 50%，Arista 第三（Interested-party）")
inp("rel_28", "Arista 份額 2028 ÷ 2026", "x", 0.70, 0.90, 1.05, "Assumed", "同上", "—")
inp("rel_29", "Arista 份額 2029 ÷ 2026", "x", 0.65, 0.88, 1.05, "Assumed", "同上", "—")

sec("C. 非 AI 營收（雲端前端、企業、園區、服務）")
inp("g_nai27", "非 AI 營收年增 2027", "%", 0.03, 0.10, 0.15, "Assumed", "2026 錨點非 AI 約 9.1（營收指引 12.6 − AI 目標 3.5）；2025 約 7.4（9.0 − AI 約 1.63）→ 2026 約 +23%（Derived，含記憶體缺貨下的提前拉貨可能）", "園區目標 2026 至少 1.25（公司，Interested-party）")
inp("svc25", "FY2025 服務營收", "US$B", None, 1.4288, None, "Verified", "Arista 4Q25 新聞稿（2026-02-12）：FY2025 Product 7,576.9、Service 1,428.8（US$M）", "分部原則：單一分部，公司固定揭露產品／服務兩條營收線")
for y_, (lo_, b_, hi_) in ((26, (0.15, 0.22, 0.28)), (27, (0.10, 0.15, 0.20)), (28, (0.06, 0.10, 0.15)), (29, (0.05, 0.08, 0.12))):
    inp(f"g_svc{y_}", f"服務營收年增 20{y_}", "%", lo_, b_, hi_, "Assumed", "FY2025 +27.7%（1,119.1 → 1,428.8；Arista 新聞稿）；服務隨裝機量與維護合約遞延認列", "非 AI 產品＝非 AI − 服務（差額）")
inp("g_nai28", "非 AI 營收年增 2028", "%", 0.02, 0.08, 0.12, "Assumed", "同上", "—")
inp("g_nai29", "非 AI 營收年增 2029", "%", 0.02, 0.06, 0.10, "Assumed", "同上", "—")
inp("ai25", "2025 年 AI 營收（參考）", "US$B", 1.5, 1.63, 1.8, "Derived", "Next Platform 2026-02-13：以 2026 目標倒推約 1.63（公司稱超過 1.5 目標）", "公司：2025 年 AI 目標 1.5 已超越（法說）")

sec("D. 毛利率")
inp("d_ai", "AI 營收毛利率低於非 AI 的幅度", "pt", 0.03, 0.06, 0.10, "Assumed", "公司：毛利率下滑來自雲端與 AI 大客戶占比（2Q26 毛利率較去年同期 −2.2 個百分點；法說，二手）", "產品別毛利率未揭露：未找到")

sec("E. FY2026 錨點（v0.4：1Q–3Q 實績＋4Q26 公司指引；原「全年指引 − 前三季」改為公司 4Q26 指引，同列改寫不位移）")
inp("q4_rev", "4Q26 營收指引（中點）", "US$B", 3.9, 3.95, 4.0, "Interested-party", PR_Q326 + "：$3.9–4.0B", "v0.3.1 本列為 FY2026 營收指引 12.6（fy_rev；4Q＝全年 − 前三季＝3.56）；新指引隱含 FY2026 約 13.35")
inp("q4_gm", "4Q26 non-GAAP 毛利率（公司新聞稿未給；假設）", "%", 0.62, 0.63, 0.64, "Assumed", "公司 4Q26 新聞稿只給營收、營業利益率、EPS；8 月全年毛利率指引 62–64%（" + CALL + "）", "3Q26 實績 64.0%、前 9 月 63.4%（公司）；法說逐字稿：未取得。只影響毛利與營業費用拆分（營業利益依指引）")
inp("q4_oim", "4Q26 non-GAAP 營業利益率指引（中點）", "%", 0.48, 0.485, 0.49, "Interested-party", PR_Q326 + "：48–49%", "v0.3.1 為 FY2026 營業利益率指引 48.5%（fy_oim）")
inp("q3_rev", "3Q26 營收（實績）", "US$B", None, 3.6576, None, "Verified", PR_Q326 + "：3,657.6 百萬美元（Product 3,190.0、Service 467.6；年增 58.5%）", "v0.3.1 為指引約 3.3（Interested-party）")
inp("q3_gp", "3Q26 non-GAAP 毛利（實績）", "US$B", None, 2.3411, None, "Verified", PR_Q326 + "：non-GAAP 毛利 2,341.1 百萬美元（毛利率 64.0%）", "v0.3.1 本列為毛利率指引約 63%（q3_gm）")
inp("q3_oi", "3Q26 non-GAAP 營業利益（實績）", "US$B", None, 1.7958, None, "Verified", PR_Q326 + "：non-GAAP 營業利益 1,795.8 百萬美元（營業利益率 49.1%）", "v0.3.1 本列為營業利益率指引 48.5%（q3_oim）")
inp("q_oth", "4Q26 non-GAAP 其他收益", "US$B", 0.13, 0.14, 0.15, "Derived", "3Q26 non-GAAP 0.1386＝GAAP 0.2731 − 策略投資利得 0.1345（公司）；1Q26 0.1136、2Q26 0.1263", "現金與有價證券 14.4（2026-09，公司）")
inp("q_tx", "4Q26 non-GAAP 有效稅率", "%", 0.20, 0.21, 0.22, "Derived", "3Q26：1 − 1530.8 ÷（1795.8＋138.6）＝20.9%；2Q26 20.6%（推導）", "4Q26 EPS 指引 1.26–1.30；本列與 q_oth、q_sh 推得約 1.27")
inp("q_sh", "4Q26 稀釋股數", "十億股", None, 1.279, None, "Derived", "3Q26 1.2782（公司）", "—")
inp("ai26", "FY2026 AI 營收（公司目標「至少」）", "US$B", 3.5, 3.5, 4.0, "Interested-party", CALL + "：AI 至少 $3.5B（2 月 3.25、8 月 3.5）；其中 scale-across 約 1.2", "sec-api 彙整：至少 $3.50B（2026-08-04 法說）")

sec("F. 營業費用、其他收益、稅、股數（FY2027 起）")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.40, 0.55, 0.80, "Derived", "FY2026 錨點營業費用年增 ÷ 營收年增（Q 頁）", "低／高 Assumed")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.02, 0.0, 0.02, "Assumed", "無", "Sens 用 ±2 個百分點")
inp("oth_f", "non-GAAP 其他收益（每年）", "US$B", 0.40, 0.55, 0.70, "Assumed", "FY2026 錨點約 0.50（Q 頁）；現金 13.3 × 約 4%", "—")
inp("tx_f", "non-GAAP 有效稅率", "%", 0.19, 0.21, 0.22, "Derived", "FY2025 19.5%、1H26 約 21%（推導）", "—")
inp("sh_g", "稀釋股數年變動", "%", -0.01, 0.0, 0.005, "Derived", "FY2025 1.2757 → 2Q26 1.2760（公司）：買回大致抵銷股票薪酬", "低／高 Assumed")
inp("mf", "AI 市場係數（Sens 用；FY2027 起）", "x", 0.8, 1.0, 1.2, "Assumed", "網通模型 2027 年 scale-out 交換器 +90%；公司稱供給吃緊是「兩年的問題」、2028 年才解除（法說，二手）", "±20%")

sec("G. 共識、市場資料（只作 Recon 對帳，不作校準）")
inp("cons_rev26", "共識 FY2026 營收", "US$B", None, 12.68, None, "Interested-party", "StockAnalysis（27 位，2026-10-23）", "MarketBeat：FY2026 EPS 共識 3.79（口徑可能為 GAAP 或舊值）")
inp("cons_eps26", "共識 FY2026 non-GAAP EPS", "US$", None, 4.12, None, "Interested-party", "StockAnalysis（區間 4.05–4.22）", "—")
inp("erste_eps27", "Erste Group FY2027 EPS（單一分析師）", "US$", None, 4.79, None, "Interested-party", "MarketBeat 2026-10-23（10-21 由 4.74 上調）", "2027 共識（StockAnalysis 付費）：未找到")
inp("dell_top2", "2025 AI 後端乙太網路：Celestica＋NVIDIA 合計份額", "%", None, 0.50, None, "Interested-party", "Dell'Oro 2026-03-10：兩家並列第一、合計約 50%；Arista 第三", "—")
inp("idc_anet", "2Q26 資料中心乙太網路交換器：Arista 份額（全部速率）", "%", None, 0.187, None, "Interested-party", "IDC 2026-09-11", "NVIDIA 20.4%、Cisco 18.2%")
inp("pe_fair_andy", "Andy 先驗：合理 P/E", "x", None, 25, None, "Assumed", "先驗_Andy_Arista_20261024 A9b（Drive；不在本 repo）", "只作對照")

sec("H. 價格與估值")
inp("px", "股價（2026-10-23 收盤）", "US$", None, 172.37, None, "Interested-party", "StockAnalysis（2026-10-23 收盤）", "markets/invest_map/data/prices.csv 同值（stockscan；bullfincher 兩來源核對）；MarketBeat：10-23 開盤 170.57")
inp("pe_ttm_gaap", "近 12 月 GAAP P/E", "x", None, 54.55, None, "Interested-party", "StockAnalysis", "MarketBeat 53.81")
inp("pe_fwd_sa", "前瞻 P/E（StockAnalysis）", "x", None, 37.06, None, "Interested-party", "StockAnalysis（口徑：未來 12 個月 non-GAAP，頁面未明示）", "第二來源：未找到")
inp("g2n", "GAAP EPS ÷ non-GAAP EPS（規則值換算用）", "x", None, "=1.75/1.89", None, "Derived", PR_Q226 + "：1H26 GAAP 稀釋 EPS 1.75、non-GAAP 1.89", "FY2025：2.75 ÷ 2.98＝0.923（公司）")
sec("I. 目標價情境權重（v0.1.1 新增；Val 頁第 0 節用）")
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")

sec("J. 目標價參數（v0.2 新增：DCF 與 Comparable 加權、FY2029 後收斂法；v0.5 起 DCF 改為標準版，折現率見 M2 節；黃底＝Andy 可改）")
inp("tp_g", "PL 延伸期營收年增的收斂目標（情境 B；FY2032；v0.5 起不再用於 DCF，DCF 永續成長率見 M2 節 g_tv）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單 2026-10-25）", "v0.5 刪除 tp_r、tp_roe（舊股東現金流法 DCF 參數）；本值只決定 FY2030–FY2032 營收收斂路徑")
inp("tp_gd", "延伸期 A／C 情境長期成長相對 tp_g 的差（A＋、C−）", "pt", None, 0.02, None, "Assumed", "工作單 2026-10-25：A＝tp_g＋2pt、C＝tp_g−2pt", "—")
inp("tp_w_dcf", "目標價中 DCF 的權重（其餘為 Comparable）", "%", None, 0.5, None, "Assumed", "Andy 可改（Andy 2026-10-25 11:53 定 50／50）", "Sens 頁 0.3／0.7")
inp("nm10_k", "延伸期 C 情境淨利率目標＝FY2029 淨利率 × 本係數（同口徑十年平均未找到）", "x", None, 0.8, None, "Assumed", "工作單：找不到十年平均用最後一年 × 0.8。本模型為 non-GAAP 口徑；同口徑十年淨利率：未找到", "參考：GAAP 十年平均 29.51%（valueinvesting.io 2016–2025：16.31%、25.71%、15.25%、35.67%、27.38%、28.52%、30.87%、35.62%、40.73%、38.99%）")
inp("cons_tp", "市場共識目標價（賣方平均）", "US$", None, 243.59, None, "Interested-party", "stockanalysis.com（32 位分析師平均，2026-10-23）", "賣方有目標價誘因；只作對照，不進目標價")
# v0.5：標準版 DCF（無槓桿自由現金流＋WACC；tools/fcff_dcf.py）。評價日（折現基準日）沿用 v0.2 的 2026 年底；股價、rf 為 2026-10-23；資產負債為 10-Q 2026-09-30
Q10 = "Arista 10-Q（季度結束 2026-09-30；accession 0001596532-26-000254；SEC XBRL companyfacts https://data.sec.gov/api/xbrl/companyfacts/CIK0001596532.json，2026-11-05 讀取）"
K10 = "Arista 10-K FY2023–FY2025（SEC XBRL companyfacts，同上；FY2025 accession 0001596532-26-000013）"
for _row in FD.in_rows(dict(
        date="2026-12-31（同 v0.2 DCF 基準日）", amt="US$B", shu="十億股",
        rf=(0.0500, "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；與 NVIDIA v0.3 同一來源）",
            "FRED DGS10（同一序列，https://fred.stlouisfed.org/series/DGS10）；與股價同日（2026-10-23）", "無風險利率（美國 10 年期公債殖利率，2026-10-23）"),
        beta_fixed=("=(0.67*0.8*beta_u_med+0.33)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=beta_u_med*(1+(1-tax_w)*beta_de)", "Derived",
                    "自算產業 β：US 網通／光通訊 同業 7 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 S&P 500 ^GSPC 迴歸，Hamada 去槓桿、現金修正）；Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；以本公司市值 D/E 與稅率重新加槓桿（Hamada：βL＝βU ×（1＋(1 − tax_w) × D/E）；Andy 2026-11-05 21:21）",
                    "低＝本組 Blume 前中位數 × 0.8 再 Blume、高＝本組 Blume 前中位數（未調整），重新加槓桿（Analogy）；個股迴歸 β 未取得"),
        kd=("=rf+0.01", "Assumed", "無金融負債（10-Q 2026-09-30 無借款）；kd＝rf＋1pt 只用於租賃負債", "負債（租賃）占資本 < 0.1%，對 WACC 影響可忽略"),
        tax_w=("=(334.7/2422+413/3265.1+738.3/4249.7)/3", "Derived",
               K10 + "：GAAP 有效稅率 FY2023 334.7 ÷ 2,422.0＝13.8%、FY2024 413.0 ÷ 3,265.1＝12.6%、FY2025 738.3 ÷ 4,249.7＝17.4%（US$M），平均",
               "只用於債務稅盾與 β 重新加槓桿；NOPAT 用 PL 頁各年 non-GAAP 有效稅率（FY2027 起 tx_f 21%）"),
        debt=(0.0, "Verified", Q10 + "：資產負債表無借款", "—"),
        lease=(0.0905, "Derived", K10 + "：FY2025 營業租賃未來付款合計 90.5（US$M，未折現；現值標籤未揭露，以未折現值作上限）", "FY2024 營業租賃負債現值 59.6（流動 22.1＋非流動 37.5）"),
        cash=("=2.2881+12.0709", "Verified", Q10 + "：現金及約當現金 2,288.1＋可供出售債務證券 12,070.9（US$M）", "8-K 2026-11-03 新聞稿：現金與有價證券約 14.4"),
        inv=(0.0813, "Verified", K10 + "：無公開市價權益證券帳面 81.3（FY2024 年底，US$M；FY2025 與 10-Q 未單獨標註）", "3Q26 策略投資利得 134.5 未反映在帳面值（保守）；Sens 測不計"),
        mi=(0.0, "Verified", Q10 + "：無少數股權", "—"),
        dps=(0, "Verified", "Arista 不配息（10-K：未曾宣告現金股利）", "—"),
        sh=("=Q!E14", "Verified", "3Q26 稀釋加權股數 1.2782 十億股（Q 頁 E14；公司 8-K）", "回購由 FCFF 支應，評價用最近一季稀釋股數，不另推未來股數"),
        fc_da=(None, "=((70.6-33.4)/5860.2+(62-26.8)/7003.1+(72.6-41.6)/9005.7)/3", None, "Derived",
               K10 + "：折舊攤銷扣除收購無形資產攤銷 ÷ 營收，FY2023 (70.6 − 33.4) ÷ 5,860.2＝0.63%、FY2024 (62.0 − 26.8) ÷ 7,003.1＝0.50%、FY2025 (72.6 − 41.6) ÷ 9,005.7＝0.34%，平均",
               "本模型 EBIT 為 non-GAAP（已排除收購無形資產攤銷），故折舊只加回非無形資產部分，避免重複"),
        fc_cx=(None, "=(34.4/5860.2+32/7003.1+119.5/9005.7)/3", None, "Derived",
               K10 + "：購置不動產、設備與無形資產 ÷ 營收，FY2023 34.4 ÷ 5,860.2＝0.59%、FY2024 32.0 ÷ 7,003.1＝0.46%、FY2025 119.5 ÷ 9,005.7＝1.33%，平均",
               Q10 + "：2026 年前三季 184.9（約 1.8% × 營收）；Sens 用 fc_cx_d ±1pt"),
        fc_cx_ss=("=MAX(fc_cx,fc_da_ss*(1+g_tv))", "Derived", "長期穩態：資本支出 ≥ 折舊 ×（1＋g）（工作單一致性檢查）", "近期比例已高於折舊時＝近期比例"),
        fc_wc=(-0.10, "=((1034.4+1945.2+412.5-435.1-407.3-161.9-915.2-591)/5860.2+(1140.5+1834.6+632.3-381.1-435.3-188.5-1727.3-1064.1)/7003.1+(1886.9+2247.1+1510-651.7-475.4-246.8-4002.6-1369.8)/9005.7)/3", 0.10, "Derived",
               K10 + "：營運資金＝應收＋存貨＋預付及其他流動資產 − 應付 − 應計負債 − 其他流動負債 − 遞延收入（流動＋非流動）；÷ 營收 FY2023 15.0%、FY2024 −2.7%、FY2025 −12.2%，平均",
               "遞延收入大增使近兩年營運資金轉負；Sens 低／高 −10%／＋10%（Assumed）"),
        )):
    rows.append(_row)
inp("fc_sbc", "股票薪酬 ÷ 營收（non-GAAP 營業利益扣除，SBC 視為費用）", "%", None, "=(296.8/5860.2+355.4/7003.1+439.2/9005.7)/3", None, "Derived",
    K10 + "：股票薪酬 FY2023 296.8 ÷ 5,860.2＝5.06%、FY2024 355.4 ÷ 7,003.1＝5.07%、FY2025 439.2 ÷ 9,005.7＝4.88%，平均",
    "個股DCF標準規格第 1 節：股票薪酬不加回。本模型 PL 為 non-GAAP（排除股票薪酬），故 DCF 的 EBIT＝non-GAAP 營業利益 − fc_sbc × 營收")
inp("beta_u", "產業無槓桿 β（自算 US 網通／光通訊，Blume 調整後；採用值）", "x", None, "=0.67*beta_u_med+0.33", None, "Derived",
    "自算產業 β：US 網通／光通訊 同業 7 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 S&P 500 ^GSPC 迴歸，Hamada 去槓桿、現金修正）；＝0.67 × 中位數 beta_u_med ＋ 0.33", "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；用現金修正值：淨現金另計，β 只對營業資產")
inp("beta_u_med", "產業無槓桿 β 中位數（自算 US 網通／光通訊，Blume 調整前；β 低／高案例用）", "x", None, 1.67, None, "Derived",
    "自算產業 β：US 網通／光通訊 同業 7 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 S&P 500 ^GSPC 迴歸，Hamada 去槓桿、現金修正）", "β 低案例＝中位數 × 0.8 再 Blume 調整；高案例＝中位數未調整（Analogy）")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務本金 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 稀釋股數 sh_v)", "與 WACC 權重同一口徑")

r = 4
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, s1, s2 = srow
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): IN.cell(r, j).font = BLUE
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    r += 1
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績（十億美元；EPS 為美元；股數為十億股）。口徑：公司 non-GAAP。藍字＝輸入；黑字＝公式。實績取自 Arista 8-K 新聞稿（SEC EDGAR、Business Wire）"
Q["A2"] = "FY2026 錨點（v0.4）＝1Q26–3Q26 實績＋4Q26 公司指引（營收 3.9–4.0、營業利益率 48–49%；毛利率公司未給，假設 63%）。1Q26＝上半年合計 − 2Q26。FY2025 non-GAAP 淨利與 EPS 為 2025 年度新聞稿原報值；公司 2026 年起改 non-GAAP 稅額方法並重編 2025 年各季，重編後的全年值：未找到"
hdr(Q, 3, ["項目", "FY2025A", "1Q26", "2Q26", "3Q26（實績）", "4Q26（指引）", "FY2026 錨點", "標記", "來源1（公司）", "來源2／說明"])
S1 = "FY2025：" + PR_Q425 + "；1Q26、2Q26：" + PR_Q226 + "（2Q26 與上半年合計）；3Q26：" + PR_Q326
S2 = "1Q26＝上半年 − 2Q26；交叉：sec-api 彙整的 GAAP、non-GAAP 毛利率與營業利益率同值"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2=S2, ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"G{r}"] = ann or f"=SUM(C{r}:F{r})"; Q[f"G{r}"].number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, s1); Q.cell(r, 10, s2)
qrow(4, "營收", {"B": 9.0057, "C": "=5.7447-D4", "D": 3.0357, "E": "=q3_rev", "F": "=q4_rev"}, tag="Verified／Interested-party（4Q 指引）")
qrow(5, "non-GAAP 毛利", {"B": 5.8154, "C": "=3.6155-D5", "D": 1.9247, "E": "=q3_gp", "F": "=F4*q4_gm"}, tag="Verified／Assumed（4Q 毛利率）")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEFG"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="FY2025 64.6%；FY2026 指引 62–64%")
qrow(7, "non-GAAP 營業費用", {c: f"={c}5-{c}8" for c in "BCDEF"}, tag="Derived（毛利 − 營業利益）")
qrow(8, "non-GAAP 營業利益", {"B": 4.3370, "C": "=2.8077-D8", "D": 1.5137, "E": "=q3_oi", "F": "=F4*q4_oim"}, tag="Verified／Interested-party（4Q 指引）")
qrow(9, "營業利益率", {c: f"={c}8/{c}4" for c in "BCDEFG"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(10, "其他收益（GAAP，non-GAAP 沿用；3Q26 起扣除策略投資利得）", {"B": 0.3936, "C": "=0.2399-D10", "D": 0.1263, "E": "=0.2731-0.1345", "F": "=q_oth"}, tag="Verified／Derived（4Q）", s2=S2 + "；3Q26＝GAAP 273.1 − 策略投資利得 134.5（公司 non-GAAP 調節）")
qrow(11, "稅前利益", {c: f"={c}8+{c}10" for c in "BCDEF"}, tag="Derived")
qrow(12, "non-GAAP 淨利", {"B": 3.8060, "C": "=2.4099-D12", "D": 1.3017, "E": 1.5308, "F": "=F11*(1-q_tx)"}, tag="Verified／Derived（4Q）")
qrow(13, "有效稅率（推導）", {c: f"=1-{c}12/{c}11" for c in "BCDEFG"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="non-GAAP 稅額未直接揭露：1 − 淨利 ÷ 稅前")
qrow(14, "稀釋股數（十億股）", {"B": 1.2757, "C": "=(1.2749*2-D14)", "D": 1.2760, "E": 1.2782, "F": "=q_sh"}, ann="=AVERAGE(C14:F14)", tag="Verified／Derived", s2="1Q26＝上半年加權 × 2 − 2Q26（近似）")
qrow(15, "non-GAAP EPS", {c: f"={c}12/{c}14" for c in "BCDEF"}, fmt="0.000", ann="=G12/G14", tag="Derived", s1="淨利 ÷ 稀釋股數", s2="—")
qrow(16, "  公司揭露 non-GAAP EPS", {"B": 2.98, "D": 1.02, "E": 1.20}, fmt="0.00", ann="", tag="Verified", s2="1Q26：1.89 − 1.02＝0.87；3Q26 1.20（舊指引 1.06–1.08）；4Q26 指引 1.26–1.30")
qrow(17, "AI 營收", {"B": "=ai25"}, ann="=ai26", tag="Derived／Interested-party", s1="FY2025：In ai25；FY2026：公司目標「至少 3.5」", s2="各季 AI 營收未揭露：未找到")
qrow(18, "非 AI 營收（＝營收 − AI）", {"B": "=B4-B17"}, ann="=G4-G17", tag="Derived", s1="—", s2="—")
qrow(19, "AI ÷ 營收", {"B": "=B17/B4"}, fmt="0.0%", ann="=G17/G4", tag="Derived", s1="—", s2="Andy 先驗 A2：2026 年 50%")
Q.cell(21, 1, "FY2026 錨點與共識的距離（參考；校準未使用共識）").font = BOLD
Q.cell(22, 1, "FY2026 營收 ÷ 共識 − 1"); Q["B22"] = "=G4/cons_rev26-1"
Q.cell(23, 1, "FY2026 EPS ÷ 共識 − 1"); Q["B23"] = "=G15/cons_eps26-1"
Q.cell(24, 1, "FY2026 營收年增"); Q["B24"] = "=G4/B4-1"
Q.cell(25, 1, "FY2026 營業費用年增"); Q["B25"] = "=G7/B7-1"
Q.cell(26, 1, "營業費用彈性（營業費用年增 ÷ 營收年增；In 頁 phi 的依據）"); Q["B26"] = "=B25/B24"
Q.cell(27, 1, "3Q26 EPS（模型）vs 公司揭露 1.20"); Q["B27"] = "=E15/1.20-1"
for rr in range(22, 28): Q[f"B{rr}"].number_format = "0.0%" if rr != 26 else "0.00"
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 7), ("oi", 8), ("oth", 10), ("ni", 12), ("sh", 14), ("eps", 15), ("ai", 17), ("nai", 18)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$G${rr}")
Q.column_dimensions["A"].width = 52
for c in "BCDEFG": Q.column_dimensions[c].width = 15
Q.column_dimensions["H"].width = 26; Q.column_dimensions["I"].width = 90; Q.column_dimensions["J"].width = 70
Q.freeze_panes = "B4"

# ---------------- 引擎 ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
BU = (2026, 2027, 2028, 2029)
umx = lambda y, s: {2026: "1", 2027: "1", 2028: f"um28_{s}", 2029: f"um29_{s}"}[y]
DEFP = dict(scn="A", rel="rel", eth="eth", mf="1", gn="g_nai", d="d_ai", opxa="0")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.00", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 8, note); R[key] = r; r += 1
    def sec_(t):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
    c_ = lambda y: FC[y]
    sec_("A. AI 營收（D 欄＝FY2026 校準年；E–G＝FY2027–FY2029）")
    row("mkt", "AI 後端 scale-out 交換器市場 × 情境顆數係數", "US$B", {c_(y): f"=mkt_{y % 100}*{umx(y, s)}" + ("" if y == 2026 else f"*{(P['mf'][y] if isinstance(P['mf'], dict) else P['mf'])}") for y in BU})
    row("eth", "乙太網路占比", "%", {c_(y): f"={P['eth']}_{y % 100}" for y in BU}, fmt="0.0%")
    row("emk", "AI 後端乙太網路交換器市場", "US$B", {c_(y): f"={c_(y)}{R['mkt']}*{c_(y)}{R['eth']}" for y in BU})
    row("sh0", "Arista 有效份額（FY2026 校準）＝ AI 目標 ÷ 乙太網路市場", "%", {"D": f"=a26_ai/D{R['emk']}"}, fmt="0.0%",
        note="Derived：分子含前端與 scale-across，所以高於純後端份額；FY2027 起乘相對值")
    row("rel", "份額相對值（÷ FY2026）", "x", {"D": "=1", **{c_(y): f"={P['rel']}_{y % 100}" for y in BU[1:]}}, fmt="0.00")
    row("sh", "Arista 有效份額", "%", {c_(y): f"=$D${R['sh0']}*{c_(y)}{R['rel']}" for y in BU}, fmt="0.0%")
    row("ai", "AI 營收", "US$B", {"C": "=a25_ai", **{c_(y): f"={c_(y)}{R['emk']}*{c_(y)}{R['sh']}" for y in BU}})
    sec_("B. 非 AI 與合計")
    row("nai", "非 AI 營收", "US$B", {"C": "=a25_nai", "D": "=a26_nai", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gn']}{y % 100})" for y in BU[1:]}})
    row("rev", "營收", "US$B", {"C": "=a25_rev", **{c_(y): f"={c_(y)}{R['ai']}+{c_(y)}{R['nai']}" for y in BU}})
    row("svc", "服務營收", "US$B", {"C": "=svc25", **{c_(y): f"={c_(y - 1)}{r}*(1+g_svc{y % 100})" for y in BU}}, note="公司營收線：Service")
    row("npd", "  產品：非 AI（園區、企業、路由等；＝非 AI − 服務）", "US$B", {c_(y): f"={c_(y)}{R['nai']}-{c_(y)}{R['svc']}" for y in FYS})
    row("prd", "產品營收（AI＋非 AI 產品）", "US$B", {c_(y): f"={c_(y)}{R['ai']}+{c_(y)}{R['npd']}" for y in FYS}, note="公司營收線：Product")
    row("segchk", "檢查：產品＋服務 − 合併營收", "US$B", {c_(y): f"={c_(y)}{R['prd']}+{c_(y)}{R['svc']}-{c_(y)}{R['rev']}" for y in FYS}, fmt="0.000")
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("aish", "AI ÷ 營收", "%", {c_(y): f"={c_(y)}{R['ai']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    sec_("C. 損益")
    row("gmn", "非 AI 毛利率 ＝ 錨點毛利率 ＋ 差距 × AI 占比（FY2026 校準）", "%", {"D": f"=a26_gp/a26_rev+{P['d']}*D{R['aish']}"}, fmt="0.0%",
        note="Derived：讓 FY2026 毛利率等於錨點；AI 毛利率＝非 AI − 差距。FY2027 起毛利率隨 AI 占比變動")
    row("gp", "毛利", "US$B", {"C": "=a25_gp", "D": "=a26_gp", **{c_(y): f"=$D${R['gmn']}*{c_(y)}{R['nai']}+($D${R['gmn']}-{P['d']})*{c_(y)}{R['ai']}" for y in BU[1:]}})
    row("gm", "毛利率", "%", {c_(y): f"={c_(y)}{R['gp']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("opb", "營業費用（基準路徑）", "US$B", {"C": "=a25_opx", "D": "=a26_opx", **{c_(y): f"={c_(y - 1)}{r}*(1+MAX(0,phi*{c_(y)}{R['g']}))" for y in BU[1:]}})
    row("opx", "營業費用", "US$B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", **{c_(y): f"={c_(y)}{R['opb']}+{P['opxa']}*{c_(y)}{R['rev']}" for y in BU[1:]}})
    row("oi", "營業利益", "US$B", {c_(y): f"={c_(y)}{R['gp']}-{c_(y)}{R['opx']}" for y in FYS})
    row("oim", "營業利益率", "%", {c_(y): f"={c_(y)}{R['oi']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("oth", "其他收益", "US$B", {"C": "=a25_oth", "D": "=a26_oth", **{c_(y): "=oth_f" for y in BU[1:]}})
    row("etr", "有效稅率", "%", {"C": "=1-a25_ni/(a25_oi+a25_oth)", "D": "=1-a26_ni/(a26_oi+a26_oth)", **{c_(y): "=tx_f" for y in BU[1:]}}, fmt="0.0%")
    row("ni", "淨利（non-GAAP）", "US$B", {"C": "=a25_ni", "D": "=a26_ni", **{c_(y): f"=({c_(y)}{R['oi']}+{c_(y)}{R['oth']})*(1-{c_(y)}{R['etr']})" for y in BU[1:]}})
    row("sh_", "稀釋股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", **{c_(y): f"={c_(y - 1)}{r}*(1+sh_g)" for y in BU[1:]}}, fmt="0.000")
    row("eps", "EPS（non-GAAP）", "US$", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['sh_']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — Arista 損益（US$B；non-GAAP）。FY2025A 實績、FY2026 錨點（1Q–2Q 實績＋3Q 指引＋全年指引），FY2027 起為模型。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要。FY2026–FY2027 三情境相同（顆數係數從 CY2028 開始）"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A：原始顆數（顆數係數 1.0；v0.12 起本庫基準＝情境 B）", "B": "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85", "C": "情境 C：CY2028–29 顆數 ×0.6"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "US$B", "#,##0.00"), ("g", "  營收年增", "%", "0.0%"), ("ai", "  AI 營收", "US$B", "#,##0.00"), ("nai", "  非 AI 營收", "US$B", "#,##0.00"),
                                ("aish", "  AI ÷ 營收", "%", "0.0%"), ("gm", "毛利率", "%", "0.0%"), ("oim", "營業利益率", "%", "0.0%"), ("ni", "淨利（non-GAAP）", "US$B", "#,##0.00"),
                                ("eps", "EPS（non-GAAP）", "US$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS:
            if key in ("g", "epsg") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型營收 − 錨點；FY2026 模型 AI − 公司目標（校準恆等式）").font = BOLD; sr += 1
PL.cell(sr, 1, "  情境 A：營收差（US$B）"); PL[f"D{sr}"] = f"=D{PB['A']['rev']}-a26_rev"; PL[f"D{sr}"].number_format = "0.0000"; CHK_ROW = sr; sr += 1
PL.cell(sr, 1, "  情境 A：AI 差（US$B）"); PL[f"D{sr}"] = f"=D{PB['A']['ai']}-a26_ai"; PL[f"D{sr}"].number_format = "0.0000"; sr += 1
PL.cell(sr, 1, "  情境 A：毛利差（US$B）"); PL[f"D{sr}"] = f"=D{PB['A']['gmn']}*D{PB['A']['nai']}+(D{PB['A']['gmn']}-d_ai)*D{PB['A']['ai']}-a26_gp"; PL[f"D{sr}"].number_format = "0.0000"
assert sr < 60, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 12; PL.column_dimensions["H"].width = 70
for c in "CDEFG": PL.column_dimensions[c].width = 14

# ---------------- Units：份額（權威） ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — Arista 在 AI 後端乙太網路交換器的份額（本專案權威，規則第 13 條）。市場取自 markets/network（只含 AI 後端 scale-out 交換器，不分協定）"
A_ = PB["A"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("mkt", "AI 後端 scale-out 交換器市場（情境 A）", "US$B", lambda y: f"=PL!{FC[y]}{A_['mkt']}", "#,##0.0", "markets/network Mkt 頁"),
        ("eth", "乙太網路占比", "%", lambda y: f"=PL!{FC[y]}{A_['eth']}", "0.0%", "In 頁 B 節（Assumed）"),
        ("emk", "AI 後端乙太網路交換器市場", "US$B", lambda y: f"=PL!{FC[y]}{A_['emk']}", "#,##0.0", ""),
        ("ib", "InfiniBand 等其他協定（＝市場 − 乙太網路）", "US$B", lambda y: f"=PL!{FC[y]}{A_['mkt']}-PL!{FC[y]}{A_['emk']}", "#,##0.0", "幾乎全為 NVIDIA"),
        ("sh", "Arista 有效份額（含前端與 scale-across）", "%", lambda y: f"=PL!{FC[y]}{A_['sh']}", "0.0%", "FY2026 以公司 AI 目標校準"),
        ("ai", "Arista AI 營收（情境 A）", "US$B", lambda y: f"=PL!{FC[y]}{A_['ai']}", "#,##0.00", ""),
        ("oth", "其他廠商（NVIDIA Spectrum-X、白牌、Cisco 等）＝ 乙太網路市場 − Arista", "US$B", lambda y: f"=PL!{FC[y]}{A_['emk']}-PL!{FC[y]}{A_['ai']}", "#,##0.0", "上限：Arista AI 含前端，所以此列偏低"),
        ("itr", "Arista AI 營收 ÷ AI IT 資本支出", "%", lambda y: f"=PL!{FC[y]}{A_['ai']}/itc_{y % 100}", "0.00%", "markets/network 的 AI IT 資本支出（加速器 × 每顆 IT capex）")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); UR[key] = ur; ur += 1
ur += 1
UN.cell(ur, 1, "讀法：FY2026 有效份額約 15%，與 Dell'Oro「Celestica＋NVIDIA 合計約 50%、Arista 第三」相容；本模型的份額路徑（相對值）是 Assumed，是最需要每季檢查的輸入。").font = BOLD
UN.column_dimensions["A"].width = 64; UN.column_dimensions["B"].width = 10
for c in "CDEF": UN.column_dimensions[c].width = 14
UN.column_dimensions["G"].width = 60

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（美元）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」與十年 P/E 規則值。Arista 的 GAAP 與 non-GAAP 差距小（約 7%），十年 GAAP P/E 適用；規則值以 GAAP 換算 EPS（non-GAAP × g2n）計算"
V["A2"] = "循環位置：Arista 不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照 NVIDIA；FY(t)＝CY(t)，不平移）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
V.cell(4, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, 5, ["項目", "A", "B", "C", "說明"])
V.cell(6, 1, "現價（2026-10-23 收盤）"); V["B6"] = "=px"; V["B6"].number_format = "0.00"
V.cell(7, 1, "現價 ÷ FY2026 EPS（錨點）"); V.cell(8, 1, "現價 ÷ FY2027 EPS（模型）"); V.cell(9, 1, "現價 ÷ FY2028 EPS（模型）"); V.cell(10, 1, "現價 ÷ FY2029 EPS（模型）")
for j, s in enumerate(SC):
    c = "BCD"[j]
    for rr, y in ((7, 2026), (8, 2027), (9, 2028), (10, 2029)):
        V[f"{c}{rr}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{rr}"].number_format = "0.0"
V.cell(11, 1, "前瞻 P/E（StockAnalysis）"); V["B11"] = "=pe_fwd_sa"; V["B11"].number_format = "0.0"
V.cell(12, 1, "  隱含未來 12 月 EPS ＝ 現價 ÷ 前瞻 P/E"); V["B12"] = "=px/pe_fwd_sa"; V["B12"].number_format = "0.00"
V.cell(13, 1, "近 12 月 GAAP P/E（StockAnalysis）"); V["B13"] = "=pe_ttm_gaap"; V["B13"].number_format = "0.0"
V["E8"] = "FY2027 三情境相同（顆數係數從 CY2028 開始）"; V["E11"] = "Andy 先驗 A9a 答 30x（口徑未註明）"
V.cell(15, 1, "2. 循環位置（本模型營收最高年；FY(t)＝CY(t)）").font = BOLD
hdr(V, 16, ["情境", "營收最高年（FY2026–FY2029）", "FY2027 判定", "規則倍數（GAAP P/E）", "見頂年是否未知", "說明"])
for i, s in enumerate(SC, 17):
    rr = f"PL!D{PB[s]['rev']}:G{PB[s]['rev']}"
    V.cell(i, 1, s)
    V.cell(i, 2, f"=INDEX({{2026,2027,2028,2029}},MATCH(MAX({rr}),{rr},0))")
    V.cell(i, 3, '=IF(B{0}=2029,"中段",'.format(i) + CP.excel_rule("2027", f"B{i}") + ")")
    V.cell(i, 4, f'=IF(C{i}="見頂",pe_peak,IF(C{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 5, f'=IF(B{i}=2029,"是：營收到觀測期末仍在增加","否")')
    V.cell(i, 6, "tools/cycle_position.py excel_rule；營收最高年取代設備市場見頂年" if s == "A" else "同上")
V.cell(21, 1, "3. 規則值隱含股價（美元；約 FY2027 年底，未折現）＝ FY2027 GAAP 換算 EPS × 規則倍數").font = BOLD
hdr(V, 22, ["情境", "FY2027 non-GAAP EPS", "GAAP 換算 EPS", "規則倍數", "隱含股價", "相對現價", "現價 ÷ GAAP 換算 EPS"])
VAL = {}
for j, s in enumerate(SC):
    i = 23 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 3, f"=B{i}*g2n").number_format = "0.00"
    V.cell(i, 4, f"=D{17 + j}").number_format = "0.0"
    V.cell(i, 5, f"=C{i}*D{i}").number_format = "#,##0"
    V.cell(i, 6, f"=E{i}/px-1").number_format = "0%"
    V.cell(i, 7, f"=px/C{i}").number_format = "0.0"
    for c in range(1, 8): V.cell(i, c).fill = KEY
    VAL[s] = i
V.cell(26, 1, "4. What must be true：FY2027 年底價 ＝ 現價 ×（1＋報酬）所需的倍數（情境 A 的 FY2028 EPS，non-GAAP；約 2027 年底的價格看下一年 EPS）").font = BOLD
hdr(V, 27, ["FY2027 年底價 ÷ 現價", "所需倍數（÷ FY2028 EPS，A）", "所需倍數（÷ FY2028 EPS，C）", "對照：Andy 合理 P/E", "對照：十年中位數（GAAP）"])
vr = 28
for x in (1.0, 1.1, 1.21, 1.5, 2.0):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    V.cell(vr, 2, f"=px*A{vr}/PL!F{SUMR[('A', 'eps')]}").number_format = "0.0"
    V.cell(vr, 3, f"=px*A{vr}/PL!F{SUMR[('C', 'eps')]}").number_format = "0.0"
    V.cell(vr, 4, "=pe_fair_andy").number_format = "0.0"; V.cell(vr, 5, "=pe_mid").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數低於歷史中位數時，現價相對便宜；倍數是否合理，另看前瞻 P/E 歷史（Bloomberg）。")
vr += 2
V.cell(vr, 1, "5. 對照：以 Andy 合理 P/E 25x 計算的隱含價（non-GAAP EPS；只作對照，非規則值）").font = BOLD; vr += 1
hdr(V, vr, ["情境", "FY2027 EPS", "× 25x", "相對現價", "FY2028 EPS", "× 25x", "相對現價"]); vr += 1
IMP = {}
for s in SC:
    V.cell(vr, 1, s)
    V.cell(vr, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"; V.cell(vr, 3, f"=B{vr}*pe_fair_andy").number_format = "#,##0"; V.cell(vr, 4, f"=C{vr}/px-1").number_format = "0%"
    V.cell(vr, 5, f"=PL!F{SUMR[(s, 'eps')]}").number_format = "0.00"; V.cell(vr, 6, f"=E{vr}*pe_fair_andy").number_format = "#,##0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
    IMP[s] = vr; vr += 1
assert vr < 50, vr
PE_CMC = [36.2, 40.3, 24.8, 18.1, 34.8, 52.4, 27.5, 34.9, 48.6, 46.8]          # companiesmarketcap 年底 P/E 2016–2025
PE_MT = [38.86, 46.56, 26.91, 19.12, 36.37, 54.50, 28.42, 35.85, None, None]    # Macrotrends 12-31 P/E 2016–2023（頁面只到 2024-04）
block_end = VR.add_rule_block(V, 50, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "companiesmarketcap（年底 P/E）", "vals": PE_CMC, "tag": "Interested-party（資料商；2016–2025 共 10 年）"},
    {"label": "Macrotrends（12 月 31 日 ÷ TTM GAAP EPS）", "vals": PE_MT, "tag": "Interested-party（資料商；2016–2023 共 8 年，2024 起頁面無資料：未找到）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 GAAP 換算 EPS（A）", "basis": f"=C{VAL['A']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 GAAP 換算 EPS（A）", "basis": f"=C{VAL['A']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 GAAP 換算 EPS（A）", "basis": f"=C{VAL['A']}"}]}],
    "0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位；兩個來源各取分位再平均。歷史倍數為 GAAP 落後 P/E，所以基準用 GAAP 換算 EPS。2016–2025 含高成長期（2021 年 52x），規則值可能高估成長放緩後的倍數。")
rule_rows = [rw for rw in range(50, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 18

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；有效份額、非 AI 毛利率在每個案例內依 FY2026 錨點重算，FY2026 恆等於錨點）：對 FY2027、FY2028 的營收、AI 營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 AI 營收", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 AI 營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("份額流失（低）", "份額相對值 2027 0.80、2028 0.70、2029 0.65（NVIDIA、白牌加速）", dict(rel="RLO")),
         ("份額持平（高）", "份額相對值 1.05", dict(rel="RHI")),
         ("乙太網路占比 低", "2026 70% → 2029 75%", dict(eth="ETLO")), ("乙太網路占比 高", "2026 80% → 2029 92%", dict(eth="ETHI")),
         ("AI 市場 ×0.8", "FY2027 起 scale-out 交換器市場 ×0.8（供給受限或內容不如預期）", dict(mf="mf_lo")), ("AI 市場 ×1.2", "×1.2", dict(mf="mf_hi")),
         ("非 AI 成長 低", "2027 +3%、2028–29 +2%", dict(gn="GNLO")), ("非 AI 成長 高", "2027 +15%、2028 +12%、2029 +10%", dict(gn="GNHI")),
         ("AI 毛利率差距 3pt", "AI 毛利率只低 3 個百分點", dict(d="d_ai_lo")), ("AI 毛利率差距 10pt", "低 10 個百分點", dict(d="d_ai_hi")),
         ("營業費用率 +2pt", "營業費用額外 +2% × 營收", dict(opxa="opx_adj_hi"))]
# 低／高路徑的具名範圍（逐年）：以 _lo／_hi 名稱組合
for nm, base, suf in (("RLO", "rel", "_lo"), ("RHI", "rel", "_hi"), ("ETLO", "eth", "_lo"), ("ETHI", "eth", "_hi"), ("GNLO", "g_nai", "_lo"), ("GNHI", "g_nai", "_hi")):
    for y in (26, 27, 28, 29):
        src = f"{base}_{y}{suf}" if base != "g_nai" else f"{base}{y}{suf}"
        if base == "rel" and y == 26: continue
        if base == "g_nai" and y == 26: continue
        target = f"{nm}_{y}" if base != "g_nai" else f"{nm}{y}"
        wb.defined_names[target] = DefinedName(target, attr_text=wb.defined_names[src].attr_text)
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.00"), ("D", "ai", 2027, "#,##0.00"), ("E", "gm", 2027, "0.0%"), ("F", "eps", 2027, "0.00"),
                           ("H", "rev", 2028, "#,##0.00"), ("I", "ai", 2028, "#,##0.00"), ("J", "gm", 2028, "0.0%"), ("K", "eps", 2028, "0.00")):
        SE[f"{c}{i}"] = f"={pre}{FC[y]}{R[key]}"; SE[f"{c}{i}"].number_format = fmt
    SE[f"G{i}"] = f"=F{i}/F$4-1"; SE[f"L{i}"] = f"=K{i}/K$4-1"
    for c in "GL": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
base_row = 40; SB = {}
for nm, setting, ov in cases:
    SB[nm] = engine(SE, base_row, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    base_row = SB[nm]["end"] + 2
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm], "")
nn = 5 + len(cases)
put_case(nn, "顆數情境 B", "CY2028 ×0.8、CY2029 ×0.85", PB["B"], "PL!")
put_case(nn + 1, "顆數情境 C", "CY2028–29 ×0.6", PB["C"], "PL!")
assert nn + 3 < 40
SENS_N = len(cases)
SE.cell(nn + 3, 1, "讀法：FY2027 三情境相同；顆數情境 B、C 的差異從 FY2028 起。份額與乙太網路占比的案例中，FY2026 有效份額依新的乙太網路占比重算，所以 FY2026 恆等於錨點。")
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 56
for c in "CDEFGHIJKL": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司指引、共識、研調份額。只作事後對照，不為對上數字而調參數。2027 年共識（StockAnalysis 付費）未找到，只有單一分析師與前瞻 P/E 隱含值"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.00", pt=False, gap=True):
    global rr_
    RC.cell(rr_, 1, label)
    if ref is not None: RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        if ref is not None and gap:
            L = chr(64 + cv)
            RC.cell(rr_, cd, f"={L}{rr_}-B{rr_}" if pt else f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y: (lambda s: f"=PL!{FC[y]}{PB[s][key]}")
RC.cell(rr_, 1, "A. 共識與分析師（校準未使用）").font = BOLD; rr_ += 1
rec("FY2026 營收 vs 共識", "=cons_rev26", pl("rev", 2026), "StockAnalysis（3Q 財報前）", "錨點＝1Q–3Q 實績＋4Q 指引中點（v0.4）")
rec("FY2026 EPS vs 共識", "=cons_eps26", pl("eps", 2026), "StockAnalysis", "錨點（3Q、4Q 依指引與 In 頁假設）")
rec("FY2027 EPS vs Erste Group（單一分析師）", "=erste_eps27", pl("eps", 2027), "MarketBeat", "2027 共識：未找到（付費）")
rec("FY2027 EPS vs 前瞻 P/E 隱含未來 12 月 EPS", "=px/pe_fwd_sa", pl("eps", 2027), "StockAnalysis（推導）", "未來 12 月 ≈ 4Q26＋¾ FY2027，口徑不同，只作量級比較")
rec("FY2027 營收年增（模型）", None, pl("g", 2027), "—", "", fmt="0.0%")
RC.cell(rr_, 1, "B. 份額與市場（先驗比對的機制檢查）").font = BOLD; rr_ += 1
rec("FY2026 Arista 有效份額 vs IDC 資料中心乙太網路份額（全部速率）", "=idc_anet", lambda s: f"=PL!D{PB[s]['sh0']}", "IDC 2Q26", "口徑不同：IDC 含前端與非 AI；本模型分母只含 AI 後端", fmt="0.0%", pt=True)
rec("FY2026 Arista 有效份額 vs 1 − Celestica、NVIDIA 合計（上限）", "=1-dell_top2", lambda s: f"=PL!D{PB[s]['sh0']}", "Dell'Oro 2025", "Arista 第三，有效份額應低於上限", fmt="0.0%", pt=True)
rec("FY2027 AI ÷ 營收 vs Andy 先驗（54%）", 0.54, pl("aish", 2027), "先驗_Andy_Arista A2（Drive）", "", fmt="0.0%", pt=True)
rec("FY2027 Arista AI 營收 ÷ AI IT 資本支出", None, lambda s: f"=PL!E{PB[s]['ai']}/itc_27", "markets/network", "Andy 先驗 A7：每 GW 4 億美元", fmt="0.00%")
rec("FY2028 Arista AI 營收 ÷ AI 後端乙太網路市場（有效份額）", None, pl("sh", 2028), "—", "", fmt="0.0%")
RC.cell(rr_, 1, "C. 校準").font = BOLD; rr_ += 1
rec("非 AI 毛利率（Derived）", None, lambda s: f"=PL!D{PB[s]['gmn']}", "FY2026 錨點", "AI 毛利率＝此值 − d_ai", fmt="0.0%")
rec("FY2027 毛利率 vs FY2026（百分點）", "=a26_gp/a26_rev", pl("gm", 2027), "錨點", "AI 占比上升使毛利率下滑", fmt="0.0%", pt=True)
rec("3Q26 EPS（模型）vs 公司揭露 1.20", 1.20, lambda s: "=Q!E15", "公司（3Q26 新聞稿）", "舊指引 1.06–1.08", fmt="0.000")
RC.cell(rr_, 1, "D. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以 Arista AI 營收 ÷ 部署 GW 計算每 GW 營收，對照 Andy 先驗 A7"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝Arista 8-K 新聞稿，已逐份抓取；法說內容只經二手轉述）"; SO["A1"].font = BOLD
srcs = [("公司：3Q26 新聞稿（2026-11-03；3Q26 實績、4Q26 指引；v0.4 滾入）", PR_Q326),
        ("公司：2Q26 新聞稿（含上半年合計、3Q26 指引）", PR_Q226),
        ("公司：4Q25 與 FY2025 新聞稿", PR_Q425),
        ("二手：BigGo Finance 2Q26 法說摘要（全年指引 12.6、AI 至少 3.5、scale-across 約 1.2、毛利率與營業利益率指引）", "https://finance.biggo.com/news/US_ANET_2026-08-04"),
        ("二手：sec-api 彙整（客戶集中、遞延收入、GAAP／non-GAAP 毛利率與營業利益率）", "https://sec-api.io/insights/arista-networks-anet-stock-analysis-q2-2026-ai-networking-revenue-deferred-revenue-customer-concentration-insider-sales-and-scenarios-to-fy2029"),
        ("二手：Next Platform（2025 營收、2025 AI 約 1.63、微軟 26%、Meta 16%）", "https://www.nextplatform.com/2026/02/13/the-current-ai-networking-wave-will-be-a-tsunami-of-money-by-2027/"),
        ("研調：IDC 2Q26 乙太網路交換器（NVIDIA 20.4%、Arista 18.7%、Cisco 18.2%）", "https://www.idc.com/resource-center/blog/ethernet-switch-market-surges-43-4-to-18-9b-in-2q26-as-ai-infrastructure-demand-drives-record-datacenter-spending/"),
        ("研調：Dell'Oro 2025 AI 後端（乙太網路超過 InfiniBand 兩倍；Celestica、NVIDIA 並列第一，Arista 第三）", "https://www.delloro.com/news/ethernet-more-than-doubles-size-of-infiniband-as-the-leading-fabric-for-ai-scale-out-networks-in-2025/"),
        ("二手：StockAnalysis 現價、前瞻 P/E、2026 共識", "https://stockanalysis.com/stocks/anet/forecast/"),
        ("二手：MarketBeat（Erste FY2027 EPS 4.79）", "https://www.marketbeat.com/instant-alerts/estimates-fy2027-eps-forecast-for-arista-networks-increased-by-analyst-2026-10-23/"),
        ("資料商：companiesmarketcap 年底 P/E", "https://companiesmarketcap.com/arista-networks/pe-ratio/"),
        ("資料商：Macrotrends 季度 P/E（至 2024-04）", "https://www.macrotrends.net/stocks/charts/ANET/arista-networks/pe-ratio"),
        ("市場層：markets/network、tsmc 成品（本 repo，唯讀）", ""),
        ("待補（人工）：2027 共識（付費）；各季 AI 營收；產品別毛利率；前瞻 P/E 十年歷史（Bloomberg）；重編後 FY2025 non-GAAP 淨利", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 90; SO.column_dimensions["B"].width = 140

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("Arista Networks（ANET）收支與估值模型 v0.4（AI 半導體 Project：網通系統端樣板）　2026-11-04", True),
 ("版本：v0.2.1（2026-10-25：補上檔空間級距、投資建議、重點觀察，數值不變）；v0.1（初版）。Arista 在 AI 後端乙太網路交換器的份額在本模型 Units 頁建立，為本專案權威資料（規則第 13 條）。", False),
 ("模型命題：Arista 的 AI 營收＝AI 後端 scale-out 交換器市場（本專案網通模型）× 乙太網路占比 × Arista 份額（含前端與 scale-across 附掛）；非 AI（雲端前端、企業、園區）另依成長率推。股價約 37 倍前瞻 EPS：在 NVIDIA（Spectrum-X）與白牌夾擊下，Arista 的份額路徑能否支撐這個倍數？", False),
 ("核心問題：2027–28 年 Arista 在 AI 乙太網路的份額每掉一截，EPS 少多少；現價隱含的份額路徑合不合理。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029（三年）。口徑：公司 non-GAAP；十年 P/E 規則值用 GAAP 換算 EPS。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("AI 營收：markets/network 的 AI 後端 scale-out 交換器市場 × 情境顆數係數 × 乙太網路占比（Assumed）× Arista 有效份額（FY2026 以公司 AI 目標 3.5 校準）× 份額相對值（Assumed）。", False),
 ("非 AI 營收：FY2026 錨點（營收指引 − AI 目標）× Assumed 成長率。毛利率：AI 比非 AI 低 d_ai（Assumed），非 AI 毛利率為 FY2026 殘差（Derived）。營業費用：彈性 phi（Derived）。", False),
 ("校準（v0.4）：FY2026 ＝ 1Q–3Q 實績＋4Q26 公司指引（營收 3.9–4.0、營業利益率 48–49%；毛利率公司未給，假設 63%）。不用共識（只在 Recon 對帳）。", False),
 ("估值：現價 ÷ FY2027、FY2028 EPS；十年 GAAP P/E 規則值（兩個資料商來源）× FY2027 GAAP 換算 EPS；循環位置以本模型營收最高年判定（Arista 不套設備股見頂年）。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司原文＝SEC EDGAR 8-K；法說內容只經二手轉述（標二手）。", False),
 ("", False),
 ("已知限制：(1) 份額相對值與乙太網路占比皆為 Assumed，是 FY2028 EPS 最敏感的輸入（Sens）；(2) FY2026 AI 用公司目標「至少 3.5」，實際可能較高；(3) 網通模型 2027 年 scale-out 交換器 +90%，公司稱供給吃緊到 2028 年，模型未設供給上限（Sens 用市場 ×0.8）；(4) 2027 共識未找到，只有單一分析師；(5) 十年 P/E 含高成長期溢價；(6) R4 產業部署路徑未建（Recon 預留列）。", False)]
lines.append(("v0.1.1（2026-10-25）：Val 頁新增「0. 目標價與上檔空間」（放在 Val 既有內容之後，不位移既有儲存格）與 In 頁目標價情境權重 tp_wA／tp_wB／tp_wC（0.3／0.5／0.2，Assumed）。各情境隱含價＝第 3 節規則值；其餘儲存格不變。", False))
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ---------------- 舊法：規則值情境加權（v0.1.1 的「0. 目標價與上檔空間」；v0.2 改名保留，不再叫目標價） ----------------
def add_target_price(V, imp, note, fmt, method, px_note):
    t0 = V.max_row + 2
    V.cell(t0, 1, "舊法：規則值情境加權（參考）").font = BOLD
    hdr(V, t0 + 1, ["項目", "數值", "", "", "說明"])
    V.cell(t0 + 2, 1, "規則值情境加權價（參考，非目標價）"); V[f"B{t0 + 2}"] = f"=tp_wA*B{t0 + 4}+tp_wB*B{t0 + 5}+tp_wC*B{t0 + 6}"; V[f"B{t0 + 2}"].number_format = fmt
    V[f"E{t0 + 2}"] = "權重 × 各情境隱含價"
    V.cell(t0 + 3, 1, "相對現價（參考）"); V[f"B{t0 + 3}"] = f"=B{t0 + 2}/px-1"; V[f"B{t0 + 3}"].number_format = "0.0%"; V[f"E{t0 + 3}"] = px_note
    for i, (s, lab) in enumerate((("A", "情境 A 隱含價（樂觀）"), ("B", "情境 B 隱含價（基準）"), ("C", "情境 C 隱含價（下檔）"))):
        rr = t0 + 4 + i
        V.cell(rr, 1, lab); V[f"B{rr}"] = imp[s]; V[f"B{rr}"].number_format = fmt; V[f"E{rr}"] = note[s]
    rr = t0 + 7
    V.cell(rr, 1, "情境權重 A／B／C"); V[f"B{rr}"] = "=tp_wA"; V[f"C{rr}"] = "=tp_wB"; V[f"D{rr}"] = "=tp_wC"
    for c in "BCD": V[f"{c}{rr}"].number_format = "0%"
    V[f"E{rr}"] = f'="合計 "&TEXT(B{rr}+C{rr}+D{rr},"0%")&"；In 頁 tp_wA／tp_wB／tp_wC（Assumed，Andy 可改）"'
    V.cell(t0 + 8, 1, "舊法說明"); V[f"B{t0 + 8}"] = method
    return t0
TP0 = add_target_price(V, {s: f"=E{VAL[s]}" for s in SC},
    {s: f'="第 3 節規則值：FY2027 GAAP 換算 EPS × "&TEXT(D{VAL[s]},"0.0")&"x（"&C{17 + SC.index(s)}&"）"' for s in SC}, "#,##0.00",
    "P/E：各情境 FY2027 GAAP 換算 EPS（non-GAAP EPS × g2n）× 十年（2016–2025 年底）GAAP P/E 依循環位置取分位（以本模型營收最高年判定：A、B 中段＝中位數，C 見頂＝25 分位；companiesmarketcap 與 Macrotrends 兩來源平均）。"
    "十年倍數含 2021 年與 2024–2025 年高成長期的高倍數，偏誤方向為高估；FY2027 三情境 EPS 相同，A、B 隱含價相同、C 只差在倍數；隱含價約為 FY2027 年底的價格、未折現，當作 12 個月目標價時偏高。",
    "現價＝In 頁 px（2026-10-23 收盤）")

# ---------------- v0.2：三情境延伸到 FY2032、Val 第 0 節（DCF 與 Comparable 加權）、Sens 目標價參數 ----------------
XH, XR, XC = TP.extend_pl(PL, PB, FC, FYS, 2029, {"rev": "rev", "ni": "ni", "sh": "sh_", "eps": "eps"}, "US$",
    "三情境延伸到 FY2032（v0.2；FY2025A–FY2029 引用上方推導區塊，FY2030 起為收斂法，非驅動因子推導；收斂年度為橘底）",
    "上游市場層（markets/network、台積電顆數）只到 CY2029＝FY2029；FY2030 起：營收年增由 FY2029 年增線性收斂到 FY2032 的長期成長（A＝tp_g＋tp_gd、B＝tp_g、C＝tp_g−tp_gd），"
    "淨利率由 FY2029 線性收斂到中段利潤率（A＝FY2029、B＝FY2029 與 C 目標的中點、C＝FY2029 × nm10_k；同口徑十年平均未找到），股數沿用 FY2029 年變動")
_h = next(r for r in range(1, V.max_row + 1) if V.cell(r, 1).value == "P/E：十年歷史倍數")
TPL, TPN = TP.add_tp_section(wb, V,
    eps={s: [f"PL!{XC[y]}{XR[(s, 'eps')]}" for y in range(2027, 2033)] for s in SC}, e_mon=12,
    fwd={s: [(f"PL!{XC[2027]}{XR[(s, 'eps')]}", "2/12"), (f"PL!{XC[2028]}{XR[(s, 'eps')]}", "10/12")] for s in SC},
    fwd_note="會計年度＝日曆年：約 12 個月後（2027-10 底）的未來 12 個月＝2027-11～2028-10（2/12 × FY2027＋10/12 × FY2028）",
    peers=[("Cisco", "CSCO", 20.70, "2026-10-23", "stockanalysis.com 統計頁（Forward PE）", "Interested-party（資料商彙整賣方預估）"),
           ("Ciena", "CIEN", 30.77, "2026-10-23", "同上", "Interested-party（同上）"),
           ("Celestica", "CLS", 16.65, "2026-10-23", "同上", "Interested-party（同上）"),
           ("Hewlett Packard Enterprise（含 Juniper）", "HPE", 12.19, "2026-10-23", "同上", "Interested-party（同上）"),
           ("Credo", "CRDO", 20.19, "2026-10-23", "同上", "Interested-party（同上）")],
    peer_note="券商常用的資料中心網通同業（交換器：Cisco、HPE＋Juniper；光傳輸：Ciena；白牌交換器 ODM：Celestica；AI 互連：Credo）；不含 A 股（Andy 11:53）。前瞻 P/E 口徑為資料商的未來 12 個月共識 EPS，各家 GAAP／non-GAAP 口徑不一",
    cons_tp="cons_tp", cons_note="stockanalysis.com：32 位分析師平均目標價（2026-10-23）；Interested-party",
    cons_fwd="px/pe_fwd_sa", cons_fwd_note="共識未來 12 個月 EPS＝現價 ÷ StockAnalysis 前瞻 P/E（FY2027 以後共識未找到）；比較窗為 2026-11～2027-10，故 B 用同窗的情境加權 2/12 × FY2026＋10/12 × FY2027",
    ours=lambda L: "(" + "+".join(f"tp_w{s}*(2/12*PL!{XC[2026]}{XR[(s, 'eps')]}+10/12*PL!{XC[2027]}{XR[(s, 'eps')]})" for s in SC) + ")",
    eps28={s: f"PL!{XC[2028]}{XR[(s, 'eps')]}" for s in SC}, hist=f"B{_h + 2}:B{_h + 11}", fmt="#,##0.00",
    method="DCF（v0.5 標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加（EBIT＝non-GAAP 營業利益 − 股票薪酬；股票薪酬視為費用），以 WACC 折現到 2026 年底（期中慣例），"
           "終值＝NOPAT(FY2033) ×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, rf)、RONIC＝WACC；EV 加淨現金 ÷ 稀釋股數，以 ke 推到 12 個月後（DCF 頁、tools/fcff_dcf.py；規格 docs/plan/個股DCF標準規格.md）。"
           "Comparable：Cisco、Ciena、Celestica、HPE、Credo 前瞻 P/E 中位數 × 12 個月後前瞻 EPS。偏誤方向：同業成長低於 Arista，中位數偏低（Comparable 偏保守）；non-GAAP EPS 排除股票薪酬（偏高；DCF 已扣回）；"
           "C 情境 FY2030 起淨利率收斂到 FY2029 × 0.8（Assumed）；歷史分位用十年 GAAP 落後 P/E（companiesmarketcap）；β 取自算 US 網通／光通訊產業 β（同業 7 家中位數 Blume 調整後重新加槓桿；非個股迴歸），DCF 對 β、WACC 敏感（Sens 頁）。")
# ---------------- v0.5：標準版 DCF（FCFF＋WACC）——DCF 明細頁＋Val 0.1 節摘要（取代 v0.2 股東現金流法；tools/fcff_dcf.py） ----------------
DCFW = wb.create_sheet("DCF")
DCFW["A1"] = "DCF — 標準版（v0.5）：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例）；終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py"
DCFW["A2"] = ("EBIT＝non-GAAP 營業利益 − 股票薪酬（fc_sbc × 營收；規格：股票薪酬視為費用、不加回）。FY2027–FY2029 取 PL 推導區塊「營業利益」；FY2030–FY2032（PL 延伸表只有淨利）由淨利還原：non-GAAP 營業利益＝淨利 ÷（1 − tx_f）− 其他收益 oth_f（與 PL 同一稅率、同一其他收益口徑）。"
              "其他收益（主要是現金利息）不進 NOPAT，現金另在評價時加入。折舊（不含收購無形資產攤銷）、資本支出、營運資金：近三年（FY2023–FY2025）占營收平均（In 頁 M2 節）。資產負債為 2026-09-30，評價日 2026-12-31 之間的現金流不另計。")
FY_X = list(range(2027, 2033))
def _rev(s_, y):
    return f"PL!{XC[y]}{XR[(s_, 'rev')]}"
def _ebit(s_, y):
    oi = f"PL!{FC[y]}{PB[s_]['oi']}" if y <= 2029 else f"(PL!{XC[y]}{XR[(s_, 'ni')]}/(1-tx_f)-oth_f)"
    return f"={oi}-fc_sbc*{_rev(s_, y)}"
def _tax(s_, y):
    return f"PL!{FC[y]}{PB[s_]['etr']}" if y <= 2029 else "tx_f"
DPATHS = [dict(name=s, label=f"情境 {s}（{ {'A': '順利', 'B': '平平', 'C': '不如預期'}[s] }）", rev={y: _rev(s, y) for y in FY_X}, rev0=_rev(s, 2026),
              ebit={y: _ebit(s, y) for y in FY_X}, tax={y: _tax(s, y) for y in FY_X}, ebit_note="non-GAAP 營業利益 − fc_sbc × 營收；FY2030 起 non-GAAP 營業利益＝淨利 ÷（1 − tx_f）− oth_f") for s in SC]
RD_ = FD.write_detail(DCFW, 4, DPATHS, FY_X, 1.0, title="三情境 FCFF 明細（US$B；FY(t)＝CY(t)；評價日 2026-12-31，FY2027 全年在評價日後）")
RD_["end"] = FD.write_checks(DCFW, RD_["end"], nc_bs="2.2881+12.0709+0.0813-0-0.0905",
                             nc_note="2026-09-30 資產負債表（US$B）：現金 2.2881＋可供出售債務證券 12.0709＋權益證券 0.0813（FY2024 帳面）− 借款 0 − 營業租賃 0.0905（FY2025 未折現付款）") + 1
DCFW.column_dimensions["A"].width = 58; DCFW.column_dimensions["B"].width = 40
for _c in "CDEFGH": DCFW.column_dimensions[_c].width = 13
_r1, _r2 = TPL["eps_h"] - 2, TPL["dcf"]            # 舊 0.1 節（標題列～DCF 目標價列）整段改寫；列數不變，0.2、0.3 節位置不動
assert V.cell(_r1, 1).value.startswith("0.1 DCF"), V.cell(_r1, 1).value
for _rr in range(_r1, _r2 + 1):
    for _cc in range(1, 8):
        V.cell(_rr, _cc).value = None; V.cell(_rr, _cc).fill = PatternFill(); V.cell(_rr, _cc).font = Font()
assert _r2 - _r1 + 1 >= FD.SUMMARY_ROWS
SDCF = FD.write_summary(wb, V, _r1, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；v0.5；明細見 DCF 頁）")
V.cell(SDCF["ke"], 5).value = '="rf "&TEXT(rf,"0.00%")&"、β "&TEXT(beta,"0.000")&"（自算 US 網通／光通訊產業 β 依 D/E 重新加槓桿）、ERP "&TEXT(erp,"0.0%")&"、CRP "&TEXT(crp,"0.0%")'
for _k in [k for k in TPL if k.startswith(("eps", "cf")) or k in ("po", "tv", "h1")]:
    del TPL[_k]
TPL.update(dcf=SDCF["dcf"], dcfv=SDCF["v12"], tvsh=SDCF["tvsh"])
V.cell(TPL["tp"], 2).value = f"=tp_w_dcf*B{SDCF['dcf']}+(1-tp_w_dcf)*B{TPL['comp']}"
V.cell(TPL["band"], 2).value = f"=D{SDCF['v12']}"; V.cell(TPL["band"], 3).value = f"=B{SDCF['v12']}"
V.cell(TPL["band"], 5).value = "B＝C 情境 DCF 價值、C＝A 情境 DCF 價值（皆已推到 12 個月後）"
SENS_TP0 = nn + 5
_tvs = FD.write_sens(SE, SENS_TP0, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref=f"Val!$B${TPL['comp']}", base_ref=f"Val!$B${TPL['tp']}",
                     title="目標價參數敏感度（v0.5 標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）")
assert _tvs[2] <= 40, _tvs
TP.dump_rowmap(TPL, TPN, XH, XR, XC, nn + 5)
TP.yellow_inputs(IN)
FD.yellow(IN, FD.JUDG + ("beta_u_med", "fc_sbc"))
SO.cell(SO.max_row + 1, 1, "[Analogy] v0.5 標準版 DCF（v0.5.1 起不再使用）：Damodaran Betas by Sector（US、Global，2026 年 1 月）Telecom. Equipment 現金修正無槓桿 β 0.89（US）、1.30（Global）"); SO.cell(SO.max_row, 2, "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/Betas.html；https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/BetasGlobal.html")
SO.cell(SO.max_row + 1, 1, "[Verified] v0.5 標準版 DCF：10-Q 2026-09-30 資產負債（現金、可供出售債務證券）與 10-K FY2023–FY2025 折舊、無形資產攤銷、資本支出、營運資金科目、股票薪酬、有效稅率（SEC XBRL companyfacts）"); SO.cell(SO.max_row, 2, "https://data.sec.gov/api/xbrl/companyfacts/CIK0001596532.json")
SO.cell(SO.max_row + 1, 1, "[Verified] v0.5 標準版 DCF：美國 10 年期公債殖利率 2026-10-23 5.00%（財政部每日殖利率曲線）"); SO.cell(SO.max_row, 2, "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610")
RD.cell(len(lines) + 1, 1, "v0.2（2026-10-25）：(1) 目標價改為 DCF 與 Comparable 加權（Val 頁最前面「0. 目標價」：0.1 DCF 股東權益現金流、0.2 Comparable 同業前瞻 P/E、0.3 結論；Val 既有內容整體下移，公式一併改寫）；"
        "v0.1.1 的「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」保留在 Val 頁末。(2) PL 頁末新增三情境延伸到 FY2032：FY2030 起為收斂法，非驅動因子推導（上游市場層只到 CY2029＝FY2029）。"
        "(3) 現價維持 2026-10-23 收盤 172.37（補第二來源 prices.csv）。(4) In 頁 J 節新增目標價參數（黃底，Andy 可改）；Sens 頁新增 tp_r、tp_g、tp_w_dcf 對加權目標價的影響。FY2025–FY2029 既有數字不變。")
RD.cell(len(lines) + 2, 1, "情境對應（v0.2 延伸期）：A＝順利（基準顆數）、B＝平平（CY2028 ×0.8、CY2029 ×0.85）、C＝不如預期（CY2028–29 ×0.6）；FY2027 三情境相同，情境差從 FY2028 起。")

# ---------------- v0.2.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 14:00、14:17；既有數值與儲存格位置不變） ----------------
_UPW = [r_ for r_ in range(1, 60) if V.cell(r_, 1).value == "上檔空間：加權目標價"]
assert _UPW == [43], _UPW
GR_ROW, RC_ROW = 52, 53   # 0.3 結論後的空白預留列（不移動其他列）
for _r in (GR_ROW, RC_ROW):
    assert all(V.cell(_r, _c).value is None for _c in range(1, V.max_column + 1)), _r
V.cell(GR_ROW, 1, "上檔空間級距")
V.cell(GR_ROW, 2, '=IF(B43>0.5,"大於 50%",IF(B43>0.25,"25%–50%",IF(B43>=0,"0%–25%",IF(B43>=-0.25,"0%～−25%",IF(B43>=-0.5,"−25%～−50%","小於 −50%")))))').number_format = "@"
V.cell(GR_ROW, 5, "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）")
V.cell(RC_ROW, 1, "投資建議")
V.cell(RC_ROW, 2, '=IF(B43>0.5,"強力買進",IF(B43>0.25,"買進",IF(B43>=-0.25,"中立",IF(B43>=-0.5,"放空","強力放空"))))').number_format = "@"
V.cell(RC_ROW, 5, "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
# 自由現金流（算力收支 2026-10-25 請求；地圖抽數依列標籤讀取）：In 頁、PL 頁、Sources 頁皆在現有最後一列之後新增，不移動既有列
_ri = IN.max_row + 1
for _k, _lab, _f, _s1 in (("fcf_da", "折舊 ÷ 營收（自由現金流用；各年營收 × 比率）", "=72.6/9006", "FY2025 10-K：折舊與攤銷 72.6 百萬美元 ÷ 營收 9,006 百萬美元（Verified）"),
                         ("fcf_capex", "資本支出 ÷ 營收（自由現金流用；各年營收 × 比率）", "=119.5/9006", "FY2025 10-K：購置不動產、設備與無形資產 119.5 百萬美元 ÷ 營收 9,006 百萬美元（Verified）")):
    IN.cell(_ri, 1, _k); IN.cell(_ri, 2, _lab); IN.cell(_ri, 3, "%"); IN.cell(_ri, 5, _f).number_format = "0.00%"
    IN.cell(_ri, 7, "Assumed"); IN.cell(_ri, 8, _s1); IN.cell(_ri, 9, "FY2026 起沿用最近一個完整年度（FY2025）年報比率（Assumed）；只用於 PL 頁自由現金流列，不進 EPS 與目標價")
    defname(_k, f"In!$E${_ri}"); _ri += 1
_fr = PL.max_row + 2
PL.cell(_fr, 1, "自由現金流（v0.2.1；算力收支 2026-10-25 請求）：淨利（non-GAAP，延伸表）＋折舊 − 資本支出；折舊、資本支出＝各年營收 × In 頁 fcf_da、fcf_capex（FY2025 年報比率）；不含營運資金變動、不進 EPS 與目標價").font = BOLD
_yrs = sorted(XC)
hdr(PL, _fr + 1, ["項目", "單位"] + [f"FY{y}" + ("A" if y == 2025 else "") for y in _yrs] + ["說明"])
for _i, s in enumerate(SC):
    _rr = _fr + 2 + _i
    PL.cell(_rr, 1, f"自由現金流（淨利＋折舊 − 資本支出；不含營運資金變動）：情境 {s}"); PL.cell(_rr, 2, "US$B")
    for y in _yrs:
        PL[f"{XC[y]}{_rr}"] = f"={XC[y]}{XR[(s, 'ni')]}+{XC[y]}{XR[(s, 'rev')]}*(fcf_da-fcf_capex)"; PL[f"{XC[y]}{_rr}"].number_format = "#,##0.00"
    PL.cell(_rr, 3 + len(_yrs), "FY2025A 為實績營收 × 年報比率；FY2030 起營收、淨利為收斂法" if _i == 0 else None)
SO.cell(SO.max_row + 1, 1, "[Verified] 公司 FY2025 10-K（SEC XBRL）：折舊與攤銷 72.6、購置不動產設備與無形資產 119.5（百萬美元）；營收 9,006 百萬美元——自由現金流比率錨點"); SO.cell(SO.max_row, 2, "https://data.sec.gov/api/xbrl/companyconcept/CIK0001596532/us-gaap/DepreciationDepletionAndAmortization.json；https://data.sec.gov/api/xbrl/companyconcept/CIK0001596532/us-gaap/PaymentsToAcquireProductiveAssets.json（10-K accession 0001596532-26-000013）")
SO.cell(SO.max_row + 1, 1, "[Assumed] 自由現金流：FY2026 起折舊 ÷ 營收、資本支出 ÷ 營收沿用 FY2025 比率（In 頁 fcf_da、fcf_capex）；不含營運資金變動"); SO.cell(SO.max_row, 2, "本模型假設")
OBS = [
       "目標價 US$164（DCF 92、Comparable 經相對倍數調整後 236 各半），比現價 172.37（2026-10-23，沿用）低 5.1%，級距 0%～−25%，投資建議維持中立；上一版 155（−9.9%）。上調來自 3Q26 實績與 4Q26 指引：FY2026 錨點營收 12.60 → 13.35、EPS 4.09 → 4.36，FY2027 EPS 5.47 → 5.75（三情境相同）；12 個月後前瞻 EPS 5.83 比共識前瞻 EPS 4.65 高 25.4%，與共識的差距主要在倍數（共識目標價隱含 52.4 倍 vs 同業 20.2 倍）。",
       "3Q26 實績營收 3.658（模型原用指引 3.30，+10.8%）、non-GAAP EPS 1.20（原估 1.07，+12.1%），毛利率 64.0%、營業利益率 49.1% 都高於指引；4Q26 指引營收 3.9–4.0（原「全年指引 − 前三季」3.56，+11.1%）、營業利益率 48–49%、EPS 1.26–1.30。公司新聞稿沒有更新 AI 營收目標，模型 FY2026 AI 仍用「至少 3.5」，上修全部落在非 AI。",
       "最敏感的是 AI 市場量與份額：AI scale-out 交換器市場 ×0.8 時 FY2027 EPS −7.4%、份額流失 −5.8%（Sens 頁）；情境 C FY2028 EPS 5.27 比 A 的 6.29 低 16.2%。下一個檢驗點：4Q26 財報（2027-02）與 2027 年營收、AI 目標指引。",
      ]   # 重點觀察（與 README.md 相同文字）
_n0 = RD.max_row
RD.cell(_n0 + 1, 1, "v0.2.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（Val 頁 0.3 結論加「上檔空間級距」「投資建議」兩列、README 頁加重點觀察三點；數值不變，既有儲存格位置不變）。另在 PL 頁末新增「自由現金流（淨利＋折舊 − 資本支出；不含營運資金變動）」三情境列（FY2025A–FY2032；算力收支 2026-10-25 請求），In 頁末新增 fcf_da、fcf_capex（FY2025 10-K 比率 0.81%、1.33%）。")
RD.cell(_n0 + 2, 1, f'="投資建議："&Val!B{RC_ROW}&"（依上檔空間固定規則）"').font = BOLD
for _i, _t in enumerate(OBS, 1): RD.cell(_n0 + 2 + _i, 1, f"重點觀察 {_i}：" + _t)
# v0.4（2026-11-04）：滾入 3Q26 實績與 4Q26 指引；README 頁只在最後一列之後追加
RD.cell(RD.max_row + 1, 1, "v0.4（2026-11-04）：滾入 3Q26 實績（8-K 2026-11-03：營收 3.6576、non-GAAP 毛利 2.3411（64.0%）、營業利益 1.7958（49.1%）、其他收益 0.1386（扣策略投資利得 0.1345）、淨利 1.5308、稀釋股數 1.2782、EPS 1.20）與 4Q26 指引（營收 3.9–4.0、營業利益率 48–49%、EPS 1.26–1.30；毛利率公司未給，假設 63%）。原 4Q26＝全年指引 − 前三季，改為公司 4Q26 指引；In 頁 E 節同列改寫（fy_rev → q4_rev 等），不位移。方法、情境結構、估值規則不變；現價沿用 2026-10-23。")
RD.cell(RD.max_row + 1, 1, "模型原估 vs 實績：3Q26 營收 3.300（指引）vs 3.658（+10.8%）、毛利率 63.0% vs 64.0%、營業利益率 48.5% vs 49.1%、EPS 1.071 vs 1.20（+12.1%）。4Q26：原（全年指引 − 前三季）營收 3.555、EPS 1.134 → 公司指引營收 3.95、模型推得 EPS 1.270（指引 1.26–1.30）。FY2026 錨點：營收 12.60 → 13.35、EPS 4.09 → 4.36。")

order =["README", "In", "Q", "Units", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "arista")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.3.1", "2026-10-30")  # v0.3.1：In 頁版面（個股模型原則第 8 節；數值不變）

# ---------------- Wattonomics 錨定（決策 S37；Siliconomics v0.13）：In 頁最後三列（wta_27／28／29，數值取 SIL.WTA）＋Sens 頁新案例；既有儲存格不變 ----------------
WTA_NOTE = "Wattonomics 錨定：以情境 B 為底，產業量 CY2027 起 × wta（Siliconomics SIF_WTAnchor_Base＝R4 比值）；FY2026 維持錨點。三層校準原則（Andy 2026-11-04）；個股基準暫不改，2027-01～02 雲端資本支出指引後再決定。"
def _wta_in(wb):
    from openpyxl.styles import Font as _F
    from openpyxl.workbook.defined_name import DefinedName as _DN
    ws = wb["In"]; r = ws.max_row + 2
    ws.cell(r, 1, "Z. Wattonomics 錨定係數（敏感度用；Siliconomics v0.13）").font = _F(bold=True); r += 1
    rows = []; SIL.wta_rows(lambda *a: rows.append(a))
    for a in rows:
        for j, v in enumerate(a, 1):
            if v is not None: ws.cell(r, j, v)
        ws.cell(r, 5).font = _F(color="0000FF")
        wb.defined_names[a[0]] = _DN(a[0], attr_text=f"In!$E${r}"); r += 1
_wta_in(wb)
def _wta_tp(SE, r0, Rw, years, last_up, keys, L, XC, XR, fwd, m10, fmt, unit, mTf=None):
    """Wattonomics 錨定案例的單一路徑目標價（Sens 頁）：延伸到 FY2032 的收斂法與 PL 延伸表情境 B 同式（tp_v2.extend_pl），
    DCF 與 Val 第 0 節同式（折現年數、期內占比取 Val F／G 欄），Comparable＝tp_comp_adj × 本路徑前瞻 EPS ÷ 情境加權前瞻 EPS。
    同時以 PL 延伸表情境 B 列算情境 B 單一路徑目標價供對照。回傳 {鍵: 儲存格}。"""
    from openpyxl.styles import Font as _F
    allys = list(years) + list(range(last_up + 1, 2033)); conv = [y for y in allys if y > last_up]; N = len(conv)
    SE.cell(r0, 1, "Wattonomics 錨定：延伸到 FY2032（收斂法同 PL 延伸表情境 B）與單一路徑目標價").font = _F(bold=True)
    SE.cell(r0 + 1, 1, "項目"); SE.cell(r0 + 1, 2, "單位")
    for y in allys: SE[f"{XC[y]}{r0 + 1}"] = f"FY{y}"
    rr = {k: r0 + 2 + i for i, k in enumerate(("rev", "g", "nm", "ni", "sh", "eps"))}
    for k, lab in (("rev", "營收"), ("g", "  營收年增"), ("nm", "淨利率"), ("ni", "淨利"), ("sh", "稀釋股數"), ("eps", f"EPS（{unit}）")): SE.cell(rr[k], 1, lab)
    lc = XC[last_up]
    M10 = "nm10" if m10 == "nm10" else f"${lc}${rr['nm']}*nm10_k"
    mT = mTf(f"${lc}${rr['nm']}") if mTf else f"(${lc}${rr['nm']}+{M10})/2"
    for y in allys:
        c = XC[y]; p = XC.get(y - 1)
        if y <= last_up:
            for k in ("rev", "ni", "sh", "eps"): SE[f"{c}{rr[k]}"] = f"={c}{Rw[keys[k]]}"
            if p: SE[f"{c}{rr['g']}"] = f"={c}{rr['rev']}/{p}{rr['rev']}-1"
            SE[f"{c}{rr['nm']}"] = f"={c}{rr['ni']}/{c}{rr['rev']}"
        else:
            k_ = y - last_up
            SE[f"{c}{rr['g']}"] = f"=${lc}${rr['g']}+(tp_g-${lc}${rr['g']})*{k_}/{N}"
            SE[f"{c}{rr['rev']}"] = f"={p}{rr['rev']}*(1+{c}{rr['g']})"
            SE[f"{c}{rr['nm']}"] = f"=${lc}${rr['nm']}+({mT}-${lc}${rr['nm']})*{k_}/{N}"
            SE[f"{c}{rr['ni']}"] = f"={c}{rr['rev']}*{c}{rr['nm']}"
            SE[f"{c}{rr['sh']}"] = f"={p}{rr['sh']}*(1+${lc}${rr['sh']}/${XC[last_up - 1]}${rr['sh']}-1)"
            SE[f"{c}{rr['eps']}"] = f"={c}{rr['ni']}/{c}{rr['sh']}"
        for k, f_ in (("rev", "#,##0.0"), ("g", "0.0%"), ("nm", "0.0%"), ("ni", "#,##0.0"), ("sh", "0.000"), ("eps", "#,##0.00")): SE[f"{c}{rr[k]}"].number_format = f_
    # v0.5：DCF 改為標準版（FCFF＋WACC）：本路徑在 DCF 頁另寫一個明細區塊（tools/fcff_dcf.py，與三情境同式）
    _P = dict(name="W", label="Wattonomics 錨定（情境 B × wta；Sens 案例，未進基準）", rev={y: f"Sens!{XC[y]}{rr['rev']}" for y in range(2027, 2033)}, rev0=f"Sens!{XC[2026]}{rr['rev']}",
              ebit={y: (f"=Sens!{XC[y]}{Rw['oi']}" if y <= last_up else f"=(Sens!{XC[y]}{rr['ni']}/(1-tx_f)-oth_f)") + f"-fc_sbc*Sens!{XC[y]}{rr['rev']}" for y in range(2027, 2033)},
              tax={y: (f"Sens!{XC[y]}{Rw['etr']}" if y <= last_up else "tx_f") for y in range(2027, 2033)}, ebit_note="non-GAAP 營業利益 − fc_sbc × 營收；FY2030 起由淨利還原")
    RW = FD.write_detail(wb["DCF"], RD_["end"] + 1, [_P], list(range(2027, 2033)), 1.0, title="Sens 案例路徑：Wattonomics 錨定（情境 B × wta）FCFF 明細")
    v12 = lambda R_, n: f"DCF!$C${R_['paths'][n]['v12']}"
    ew = lambda y: f"{XC[y]}{rr['eps']}"; eb = lambda y: f"PL!{XC[y]}{XR[('B', 'eps')]}"
    h = rr["eps"] + 2
    for j, lab in enumerate(["目標價（單一路徑）", "Wattonomics 錨定", "情境 B（PL 延伸表）", "基準（情境加權，Val 第 0 節）"]): SE.cell(h, 1 + j, lab).font = _F(bold=True)
    T = {"dcf": h + 1, "fwd": h + 2, "comp": h + 3, "tp": h + 4, "chk": h + 5}
    SE.cell(T["dcf"], 1, "DCF 價值推到 12 個月後（標準版 FCFF＋WACC；×（1＋ke）^t_tgt）"); SE.cell(T["dcf"], 2, "=" + v12(RW, "W")); SE.cell(T["dcf"], 3, "=" + v12(RD_, "B")); SE.cell(T["dcf"], 4, f"=Val!B{L['dcf']}")
    SE.cell(T["fwd"], 1, "12 個月後前瞻 EPS（權重同 Val 第 0 節）")
    SE.cell(T["fwd"], 2, "=" + "+".join(f"{w}*{ew(y)}" for y, w in fwd)); SE.cell(T["fwd"], 3, "=" + "+".join(f"{w}*{eb(y)}" for y, w in fwd)); SE.cell(T["fwd"], 4, f"=Val!B{L['fwd']}")
    SE.cell(T["comp"], 1, "Comparable（相對倍數調整後；tp_comp_adj × 本路徑前瞻 EPS ÷ 情境加權前瞻 EPS）")
    for c in "BC": SE[f"{c}{T['comp']}"] = f"=tp_comp_adj*{c}{T['fwd']}/Val!B{L['fwd']}"
    SE[f"D{T['comp']}"] = "=tp_comp_adj"
    SE.cell(T["tp"], 1, "目標價（tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable）")
    for c in "BC": SE[f"{c}{T['tp']}"] = f"=tp_w_dcf*{c}{T['dcf']}+(1-tp_w_dcf)*{c}{T['comp']}"
    SE[f"D{T['tp']}"] = f"=Val!B{L['tp']}"
    for k in ("dcf", "fwd", "comp", "tp"):
        for c in "BCD": SE[f"{c}{T[k]}"].number_format = "#,##0.00" if k == "fwd" else fmt
    SE.cell(T["chk"], 1, "核對：情境 B 欄 DCF − Val 第 0.1 節情境 B 價值推到 12 個月後（應為 0）"); SE.cell(T["chk"], 3, f"=C{T['dcf']}-Val!C{L['dcfv']}").number_format = "0.000000"
    return {"tp_w": f"B{T['tp']}", "tp_b": f"C{T['tp']}", "tp_base": f"D{T['tp']}", "eps": ew, "end": T["chk"]}


def _wta_table(SE, r0, cols, rows):
    """小表：cols＝[(欄名, fmt)]；rows＝[(名稱, [儲存格…])]，第一列＝基準、第二列＝情境 B；每個指標後接「相對基準」「相對情境 B」。"""
    from openpyxl.styles import Font as _F
    from openpyxl.utils import get_column_letter as _gl
    SE.cell(r0, 1, "案例").font = _F(bold=True)
    for j, (nm, fmt) in enumerate(cols):
        for k, t in enumerate((nm, "相對基準", "相對情境 B")): SE.cell(r0, 2 + 3 * j + k, t).font = _F(bold=True)
    for i, (nm, cells) in enumerate(rows):
        r = r0 + 1 + i; SE.cell(r, 1, nm)
        for j, (cell, (cn, fmt)) in enumerate(zip(cells, cols)):
            v, d1, d2 = (_gl(2 + 3 * j + k) for k in range(3))
            SE[f"{v}{r}"] = f"={cell}"; SE[f"{v}{r}"].number_format = fmt
            SE[f"{d1}{r}"] = f"={v}{r}/{v}${r0 + 1}-1"; SE[f"{d2}{r}"] = f"={v}{r}/{v}${r0 + 2}-1"
            SE[f"{d1}{r}"].number_format = SE[f"{d2}{r}"].number_format = "0.0%"
    return r0 + len(rows)
WTA_D = {2027: "wta_27", 2028: "wta_28", 2029: "wta_29"}
_w0 = SE.max_row + 3
SE.cell(_w0, 1, "Wattonomics 錨定（三層校準原則）").font = BOLD
SE.cell(_w0 + 1, 1, WTA_NOTE)
SE.cell(_w0 + 2, 1, '做法：FY(t)＝CY(t)。AI 後端 scale-out 交換器市場的 AI 市場係數 mf 改為年別（FY2027＝wta_27 … FY2029＝wta_29）；有效份額、非 AI 毛利率由 FY2026 校準不受影響；非 AI、服務不乘。FY2030 起沿用 PL 延伸表情境 B 的收斂法。')
_t0 = _w0 + 4
SB_W = engine(SE, _t0 + 6, "案例區塊：Wattonomics 錨定（情境 B × wta）", scn="B", mf=WTA_D)
_W = _wta_tp(SE, SB_W["end"] + 1, SB_W, FYS, 2029, {'rev': 'rev', 'ni': 'ni', 'sh': 'sh_', 'eps': 'eps'}, TPL, XC, XR, [(2027, '2/12'), (2028, '10/12')], 'nm10_k', '#,##0.00', 'US$', mTf=None)
_e = lambda s_, y: f"PL!{XC[y]}{XR[(s_, 'eps')]}"
_wta_table(SE, _t0, [("FY2027 EPS", "0.00"), ("FY2028 EPS", "0.00"), ("目標價", '#,##0.00')],
           [("基準（Sens 主表基準＝情境 A；目標價＝情境加權）", [_e("A", 2027), _e("A", 2028), _W["tp_base"]]),
            ("情境 B（案例底）", [_e("B", 2027), _e("B", 2028), _W["tp_b"]]),
            ("Wattonomics 錨定（情境 B × wta）", [_W["eps"](2027), _W["eps"](2028), _W["tp_w"]])])

# ---------------- v0.4.1（2026-11-05）：只改展示層——README 頁「展示層」與「命題回覆」兩段（個股 HTML 版面規格 v1.0；公式帶數字，原列保留） ----------------
RD["A1"] = "Arista Networks（ANET）收支與估值模型 v0.5.1（AI 半導體 Project：網通系統端樣板）　2026-11-05；v0.5.1：β 改用自算產業 β（US 網通／光通訊；Andy 2026-11-05 22:55），其他不變；v0.4.1 只改展示層（README 頁新增展示層與命題回覆兩段、HTML 改用個股 HTML 版面規格 v1.0）；v0.5 模型改版：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數），Comparable、情境權重、DCF 權重 50% 與投資建議規則不變"
def _lr(ws, prefix, col=1):
    """A 欄（或指定欄）以 prefix 開頭的第一列。"""
    return next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, col).value or "").startswith(prefix))
_ik = lambda k: _lr(IN, k) if IN.cell(_lr(IN, k), 1).value == k else None
_V = {k: _lr(V, p) for k, p in (("tp", "目標價（DCF 與 Comparable 加權）"), ("dcf", "DCF 目標價（情境加權）"), ("comp", "Comparable 目標價"),
                                ("compadj", "Comparable 目標價（相對倍數調整後）"), ("up", "上檔空間：加權目標價"), ("upc", "上檔空間：Comparable"),
                                ("rec", "投資建議"), ("px", "現價（"), ("pe28", "  檢查用：現價 ÷ 各情境 FY2028 EPS"), ("pct", "歷史倍數檢查"),
                                ("med", "同業前瞻 P/E 中位數"), ("fwd", "12 個月後的前瞻 EPS（情境加權）"), ("epsgap", "與市場共識目標價的差距：EPS 因素"), ("gap", "與市場共識目標價的差距：倍數因素"), ("coef", "相對倍數調整係數"), ("raw", "原始比值"), ("tvsh", "  其中：終值現值占 EV"), ("mult", "  隱含終值 EV／EBITDA"))}
_v = lambda k, c="B": f"Val!{c}{_V[k]}"
_S = {k: _lr(SE, p) for k, p in (("share", "份額流失（低）"), ("hold", "份額持平（高）"), ("mkt", "AI 市場 ×0.8"), ("scC", "顆數情境 C"), ("nai", "非 AI 成長 低"), ("wdcf", "DCF 權重 0.3／0.7"), ("wacc", "WACC ±1pt"), ("beta", "β 低／高"), ("erp", "ERP 低／高"), ("gtv", "永續成長率 g ±1pt"))}
_s = lambda k, c: f"Sens!{c}{_S[k]}"
_pl = lambda s, k, y: f"PL!{FC[y]}{PB[s][k]}"
assert PL.cell(XR[("B", "rev")] + 1, 1).value.strip() == "營收年增" and PL.cell(XR[("B", "rev")] + 2, 1).value == "淨利率" and PL.cell(XR[("B", "eps")] + 1, 1).value.strip() == "EPS 年增" and XC[2032] == "J"
_xb = lambda k, off, c: f"PL!{c}{XR[('B', k)] + off}"   # 延伸表情境 B（FY2029＝G、FY2032＝J）
_u_sh = _lr(UN, "Arista 有效份額")
_P = lambda x: f'TEXT({x},"+0.0%;-0.0%")'
_N = lambda x, f="0.00": f'TEXT({x},"{f}")'
_r = RD.max_row + 2
RD.cell(_r, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；A3、A4、重點觀察原列保留）").font = BOLD; _r += 1
RD.cell(_r, 1, "命題：Arista 的 AI 營收取決於 AI 後端交換器市場、乙太網路占比與 Arista 份額；在 NVIDIA（Spectrum-X）與白牌夾擊下，份額路徑與市場量能否支撐目前的高倍數。三個問題如下。"); _r += 1
RD.cell(_r, 1, "命題 1｜份額掉一截，EPS 少多少：在 NVIDIA 與白牌夾擊下，Arista 在 AI 後端乙太網路的份額若比基準再往下掉，2027–28 年的 EPS 會少多少。"); _r += 1
RD.cell(_r, 1, "命題 2｜成長靠什麼：AI 營收＝AI 後端 scale-out 交換器市場 × 乙太網路占比 × Arista 份額；加速器顆數放緩（情境 B、C）時，2028 年 EPS 還能不能成長。"); _r += 1
RD.cell(_r, 1, f'="命題 3｜現價合不合理：現價約 "&{_N("pe_fwd_sa", "0")}&" 倍前瞻 EPS（StockAnalysis），遠高於網通同業；這個倍數合不合理、每股值多少。"'); _r += 1
RD.cell(_r, 1, f'="重點 1｜合理價值"&IF({_v("up")}<0,"低於","高於")&"現價，投資建議"&{_v("rec")}&"：加權目標價每股約 "&{_N(_v("tp"))}&" 美元（DCF "&{_N(_v("dcf"))}&"、Comparable 經相對倍數調整後 "&{_N(_v("compadj"))}&"，各占一半），相對現價 "&{_N(_v("px"))}&" 為 "&{_P(_v("up"))}&"。只看同業倍數、不做相對倍數調整的 Comparable 目標價是 "&{_N(_v("comp"))}&"（"&{_P(_v("upc"))}&"），這個數不進加權目標價。"'); _r += 1
RD.cell(_r, 1, f'="重點 2｜兩種算法差很多，差在倍數：Comparable 用網通同業前瞻 P/E 中位數 "&{_N(_v("med"), "0.0")}&" 倍，再乘相對倍數調整係數 "&{_N(_v("coef"), "0.0")}&"（Arista 過去十年前瞻 P/E 約為同業的 "&{_N(_v("raw"), "0.0")}&" 倍"&IF({_v("raw")}>{_v("coef")}+1E-9,"，夾到上限","")&"）；DCF（v0.5 起為標準版：無槓桿自由現金流以 WACC 折現）只有 "&{_N(_v("dcf"))}&"：WACC "&{_N("wacc", "0.0%")}&"（rf "&{_N("rf", "0.0%")}&" ＋ β "&{_N("beta", "0.00")}&" × ERP "&{_N("erp", "0%")}&"；β 取自算 US 網通／光通訊產業 β（同業中位數 Blume 調整後依 D/E 重新加槓桿），ERP 與 RONIC＝WACC 為假設），EPS 成長在 FY2028 起放緩（情境 B 從 FY2027 的 "&{_N(_pl("B", "eps", 2027))}&" 到 FY2028 的 "&{_N(_pl("B", "eps", 2028))}&"），終值現值占情境 B 企業價值 "&{_N(_v("tvsh", "C"), "0%")}&"、隱含終值 EV／EBITDA 約 "&{_N(_v("mult", "C"), "0.0")}&" 倍；FY2030 起為收斂法：情境 B 營收年增逐年收斂到 FY2032 的 "&{_N(_xb("rev", 1, "J"), "0.0%")}&"（＝長期成長率 tp_g），淨利率由 FY2029 的 "&{_N(_xb("rev", 2, "G"), "0.0%")}&" 下修到 FY2032 的 "&{_N(_xb("rev", 2, "J"), "0.0%")}&"，FY2032 EPS 年增只剩 "&{_P(_xb("eps", 1, "J"))}&"，趨近於零。市場共識目標價隱含約 "&{_N(_v("gap", "C"), "0.0")}&" 倍。"'); _r += 1
RD.cell(_r, 1, f'="重點 3｜EPS 對加速器顆數最敏感，份額流失測試案例的影響在一成以內：以情境 A 為基準、一次改一個輸入，顆數情境 C（"&{_s("scC", "B")}&"）讓 FY2028 EPS 變動 "&{_P(_s("scC", "L"))}&"，是唯一達兩位數百分比的測試案例；份額流失 "&{_P(_s("share", "L"))}&"、非 AI 成長低 "&{_P(_s("nai", "L"))}&"、AI 市場 ×0.8 "&{_P(_s("mkt", "L"))}&"。目標價參數（以情境加權目標價為基準，一次動一個）：DCF 權重 0.3／0.7 時加權目標價 "&{_N(_s("wdcf", "C"))}&"／"&{_N(_s("wdcf", "E"))}&"；DCF 參數案例：β 低／高（"&{_s("beta", "B")}&"）相對基準 "&{_P(_s("beta", "D"))}&"／"&{_P(_s("beta", "F"))}&"，WACC ±1pt 為 "&{_P(_s("wacc", "D"))}&"／"&{_P(_s("wacc", "F"))}&"。"'); _r += 1
_r += 1
RD.cell(_r, 1, "命題回覆（HTML 用；依本模型計算回答開頭三個命題）").font = BOLD; _r += 1
_rl = {k: _ik(k) for k in ("rel_27", "rel_28", "rel_29")}
RD.cell(_r, 1, f'="份額掉一截，EPS 少多少？答：影響有限。以情境 A 為基準，份額相對值（相對 FY2026）由基準 "&{_N("In!E" + str(_rl["rel_27"]))}&"／"&{_N("In!E" + str(_rl["rel_28"]))}&"／"&{_N("In!E" + str(_rl["rel_29"]))}&" 改為 "&{_N("In!D" + str(_rl["rel_27"]))}&"／"&{_N("In!D" + str(_rl["rel_28"]))}&"／"&{_N("In!D" + str(_rl["rel_29"]))}&"（2027／28／29）的測試案例，FY2027 EPS 變動 "&{_P(_s("share", "G"))}&"、FY2028 EPS 變動 "&{_P(_s("share", "L"))}&"；份額持平（"&{_N("In!F" + str(_rl["rel_27"]))}&"）時 FY2028 EPS 變動 "&{_P(_s("hold", "L"))}&"。影響不大的原因是 AI 只占 FY2027 營收約 "&{_N(_pl("A", "aish", 2027), "0%")}&"，其餘是非 AI 產品與服務。情境 A 的有效份額由 FY2026 的 "&{_N(f"Units!C{_u_sh}", "0.0%")}&" 降到 FY2028 的 "&{_N(f"Units!E{_u_sh}", "0.0%")}&"（基準已含份額逐年下滑）。"'); _r += 1
RD.cell(_r, 1, f'="成長靠什麼、顆數放緩時 EPS 還能成長嗎？答：FY2027 靠 AI 市場量：AI 營收由 FY2026 的 "&{_N(_pl("A", "ai", 2026), "0.0")}&"（公司目標「至少」）增至 "&{_N(_pl("A", "ai", 2027), "0.0")}&" 十億美元，三情境相同，EPS "&{_N(_pl("A", "eps", 2027))}&"。差異從 CY2028 的顆數係數開始：FY2028 EPS 情境 A／B／C 為 "&{_N(_pl("A", "eps", 2028))}&"／"&{_N(_pl("B", "eps", 2028))}&"／"&{_N(_pl("C", "eps", 2028))}&"，相對 FY2027 為 "&{_P(f'{_pl("A", "eps", 2028)}/{_pl("A", "eps", 2027)}-1')}&"／"&{_P(f'{_pl("B", "eps", 2028)}/{_pl("B", "eps", 2027)}-1')}&"／"&{_P(f'{_pl("C", "eps", 2028)}/{_pl("C", "eps", 2027)}-1')}&"；情境 C 的 FY2028 營收 "&{_N(_pl("C", "rev", 2028), "0.0")}&" 低於 FY2027 的 "&{_N(_pl("C", "rev", 2027), "0.0")}&"，營收在 FY2027 見頂。"'); _r += 1
RD.cell(_r, 1, f'="現價合不合理、每股值多少？答：每股約 "&{_N(_v("tp"))}&" 美元（加權目標價；DCF "&{_N(_v("dcf"))}&"、Comparable 經相對倍數調整後 "&{_N(_v("compadj"))}&" 各半），相對現價 "&{_N(_v("px"))}&" 為 "&{_P(_v("up"))}&"，投資建議"&{_v("rec")}&"。和同業比要用同一時間窗：網通同業前瞻 P/E 中位數 "&{_N(_v("med"), "0.0")}&" 倍是各同業現價 ÷ 各自未來 12 個月（約 2026-11～2027-10）的共識 EPS；同窗下，Arista 現價 ÷ 本模型同期 EPS（情境加權 "&{_N(_v("epsgap", "C") + "*(1+" + _v("epsgap") + ")")}&"）約 "&{_N(_v("px") + "/(" + _v("epsgap", "C") + "*(1+" + _v("epsgap") + "))", "0.0")}&" 倍，是同業中位數的 "&{_N(_v("px") + "/(" + _v("epsgap", "C") + "*(1+" + _v("epsgap") + "))/" + _v("med"), "0.0")}&" 倍（用共識 EPS 則是 "&{_N(_v("px") + "/" + _v("epsgap", "C"), "0.0")}&" 倍）；本模型的 Comparable 已按 Arista 長期相對同業的溢價乘上調整係數 "&{_N(_v("coef"), "0.0")}&"，不做這個調整時 Comparable 目標價只有 "&{_N(_v("comp"))}&"（"&{_P(_v("upc"))}&"）。對照 Arista 自己的歷史：現價 ÷ FY2028 EPS 為 "&{_N(_v("pe28"), "0.0")}&"／"&{_N(_v("pe28", "C"), "0.0")}&"／"&{_N(_v("pe28", "D"), "0.0")}&" 倍（情境 A／B／C），落在 2016–2025 年底落後 GAAP P/E 的 "&{_N(_v("pct"), "0%")}&"～"&{_N(_v("pct", "D"), "0%")}&" 分位（口徑不同：兩年後 non-GAAP EPS 對落後 GAAP P/E，只作參考）。"'); _r += 1
_r += 1
RD.cell(_r, 1, "本版要點與版本沿革（HTML 附錄用）").font = BOLD; _r += 1
RD.cell(_r, 1, '="v0.5.1（2026-11-05）：β 改用自算產業 β（Andy 2026-11-05 22:55）：US 網通／光通訊同業 7 家（ANET、CSCO、CIEN、COHR、LITE、FN、AAOI）無槓桿 β 中位數 "&TEXT(beta_u_med,"0.00")&"（Yahoo Finance 月報酬 2021-11～2026-10 對 S&P 500 迴歸，Hamada 去槓桿、現金修正），Blume 調整後 "&TEXT(beta_u,"0.000")&"，依市值 D/E 重新加槓桿 β＝"&TEXT(beta,"0.000")&"，WACC "&TEXT(wacc,"0.00%")&"；取代 Damodaran 美國 Telecom. Equipment 0.89。β 低／高案例改為中位數 × 0.8 再 Blume／中位數未調整。評價日維持 2026-12-31。其他輸入不變。"'); _r += 1
RD.cell(_r, 1, "v0.5（2026-11-05）：模型改版——DCF 改為標準版（Andy 2026-11-05 20:42、20:45；docs/plan/個股DCF標準規格.md v1.0）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加（EBIT＝non-GAAP 營業利益 − 股票薪酬，股票薪酬視為費用），以 WACC 折現（ke＝rf＋β ×（ERP＋CRP），rf 5.00%、β＝Damodaran 美國 Telecom. Equipment 現金修正無槓桿 β 0.89 依市值 D/E 重新加槓桿、ERP 5%、CRP 0），終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, rf)、RONIC＝WACC；EV 加淨現金 ÷ 稀釋股數，以 ke 推到 12 個月後。新增 DCF 頁（共用模組 tools/fcff_dcf.py），Val 0.1 節改為 WACC 與 DCF 彙總，Sens 目標價參數改為 WACC、β、ERP、g、RONIC、資本支出、營運資金、權益投資、終值替代法與 DCF 權重；刪除 tp_r、tp_roe，tp_g 只留給 PL 延伸期收斂。PL、Comparable、情境權重不變。"); _r += 1
RD.cell(_r, 1, "v0.4.1（2026-11-05）：只改展示層——README 頁新增「展示層」「命題回覆」兩段（公式帶數字），HTML 改用個股 HTML 版面規格 v1.0（tools/company_html/relayout.py）；README 以外逐格與 v0.4 相同。"); _r += 1
RD.cell(_r, 1, f"=A{_lr(RD, 'v0.4（2026-11-04）')}"); _r += 1
SO.cell(SO.max_row + 1, 1, "[Derived] 自算產業 β（Andy 2026-11-05 22:55）：US 網通／光通訊同業 7 家（ANET、CSCO、CIEN、COHR、LITE、FN、AAOI）Yahoo Finance 含股利調整月收盤價 2021-11～2026-10，對 S&P 500 ^GSPC 迴歸；Hamada 去槓桿、現金修正，中位數 1.67；Blume 調整 0.67β＋0.33＝1.449"); SO.cell(SO.max_row, 2, "https://finance.yahoo.com（chart API 月資料；2026-11-05 抓取）")
wb.save(sys.argv[1])
if os.environ.get("ARISTA_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK_ROW, "SENSN": SENS_N, "IMP": IMP, "VAL": VAL, "SB": SB},
              open(os.environ["ARISTA_ROWMAP"], "w"), ensure_ascii=False, indent=1)

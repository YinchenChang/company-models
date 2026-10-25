# Coherent（COHR，NYSE）收支與估值模型 v0.1（AI 半導體 Project：網通次產業內複製——光模組，樣板 innolight/）
# 用法：python3 cohr/build_cohr.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度：FY(t)＝t−1 年 7 月～t 年 6 月（6 月底結束）。FY(t) 一半落在 CY(t−1) 下半年、一半落在 CY(t) 上半年。
#   市場模型是日曆年：FY2027 資料中心＝CY2026 下半年（FY2027 Q1 指引＋Q2 假設）＋ hw × CY2027；FY2028＝（1 − hw）× CY2027 ＋ hw × CY2028；FY2029 同理。
# 金額：十億美元（US$B）；EPS：美元；口徑：公司 non-GAAP（排除股票薪酬、併購攤銷、重組、處分損益等）；淨利＝歸屬 Coherent 淨利。
# FY2026（2026-06 結束）為完整實績，即錨點。已出售事業（航太國防 A&D、慕尼黑雷射加工工具）以 FY2026 Q4 工業營收年化排除。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP
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
NET_F = only("markets/network/dist/*.xlsx"); TSMC_F = only("tsmc/dist/*.xlsx"); INL_F = only("innolight/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
NETR = {}
for r in openpyxl.load_workbook(NET_F, data_only=True)["Mkt"].iter_rows(min_row=4):
    lab = r[0].value
    if lab in ("  其中 CPO", "  其中可插拔", "AI IT 資本支出（加速器 × 每顆 IT capex）"):
        NETR[lab.strip()] = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
assert len(NETR) == 3 and all(v is not None for k in NETR for v in NETR[k][0]), NETR.keys()
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and str(k).startswith("um"):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
assert all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()
INL = None
for r in openpyxl.load_workbook(INL_F, data_only=True)["Units"].iter_rows(min_row=4):
    if r[0].value == "中際旭創可插拔有效份額":
        INL = ([r[2 + j].value for j in range(4)], f"{rel(INL_F)}｜Units!C{r[0].row}:F{r[0].row}")
assert INL and all(isinstance(v, float) for v in INL[0]), INL

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億美元（US$B）；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文＝SEC 8-K 新聞稿與 10-K；資料中心與通訊的金額拆分公司未揭露，以法說會季增、年增率推導"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
PR = "https://www.sec.gov/Archives/edgar/data/0000820318/000119312526346860/d128030dex991.htm"
F4 = "https://futurumgroup.com/insights/coherent-q4-fy-2026-earnings-1-6t-transceivers-ramp-cpo-revenue-approaches/"
WB4 = "https://www.webull.com/news/15393555909010432"
F1 = "https://futurumgroup.com/insights/coherent-q1-fy-2026-shows-strong-datacenter-comms-momentum"
F2 = "https://futurumgroup.com/insights/coherent-q2-fy-2026-ai-datacenter-demand-lifts-revenue-and-margins"
F3 = "https://futurumgroup.com/insights/coherent-q3-fy-2026-ai-data-center-demand-accelerates-optical-growth/"
TENK = "https://www.sec.gov/Archives/edgar/data/820318/000082031826000020/iivi-20260630.htm"
AB = "https://www.autobound.ai/companies/coherent/10-k"
INV = "https://www.investing.com/news/stock-market-news/coherent-posts-a-blowout-q4-and-aggressive-guidance-yet-shares-take-a-5-haircut-4856014"
MB = "https://www.marketbeat.com/stocks/NYSE/COHR"
BZ = "https://www.benzinga.com/quote/COHR/earnings-forecasts"
TK = "https://www.tikr.com/blog/coherent-stock-fell-nearly-10-in-a-day-as-rising-yields-hit-optics-stocks-heres-where-the-stock-could-go-by-2031"
LC = "https://lightcounting.com/newsletter/en/may-2026-optical-vendor-landscape-376"
CMC = "https://companiesmarketcap.com/coherent/pe-ratio/"
MT = "https://www.macrotrends.net/stocks/charts/COHR/coherent/pe-ratio"
DG = "https://www.digrin.com/stocks/detail/COHR/price"
GF = "https://www.google.com/finance/labs/quote/COHR:NYSE"
FV = "https://finviz.com/news/402637/coherent-corp-announces-timing-of-fy2027-first-quarter-earnings-release"

sec("A. AI 光模組市場（讀 markets/network 成品；日曆年；情境 A 基準，B／C 依顆數係數縮放）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"mpl_{y}", f"AI 可插拔光模組市場 CY20{y}", "US$B", None, NETR["其中可插拔"][0][j], None, "Derived", NETR["其中可插拔"][1],
        "加速器顆數 × 每顆光學內容（scale-out 為主；2029 含 scale-up 光學）− CPO" if y == 26 else "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"mcp_{y}", f"AI CPO 光學市場 CY20{y}", "US$B", None, NETR["其中 CPO"][0][j], None, "Derived", NETR["其中 CPO"][1], "共同封裝光學；網通模型的 CPO 滲透路徑" if y == 26 else "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 CY20{y}", "US$B", None, NETR["AI IT 資本支出（加速器 × 每顆 IT capex）"][0][j], None, "Derived", NETR["AI IT 資本支出（加速器 × 每顆 IT capex）"][1], "Recon 用：Coherent 資料中心營收 ÷ AI IT 資本支出" if y == 26 else "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"inl_{y}", f"中際旭創可插拔有效份額 CY20{y}（份額口徑一致；加總檢查用）", "%", None, INL[0][j], None, "Derived", INL[1], "中際旭創模型為光模組份額權威；本模型 Coherent 份額與它相加不得超過 100%（Units 頁）" if y == 26 else "同上")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝基準；B＝AI 資本支出消化；C＝明顯下修（與台積電、NVIDIA、中際旭創模型相同）")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")

sec("B. Coherent 份額與單價（資料中心＝光模組＋自製 EML／VCSEL／CW 雷射＋矽光元件；有效份額＝資料中心營收 ÷ 可插拔市場）")
inp("rel_27", "可插拔有效份額 CY2027 ÷ CY2026（相對值）", "x", 0.90, 1.10, 1.25, "Assumed", F4 + "：磷化銦雷射內部產能 CY2026 底較前一年倍增、CY2027 底再倍增以上；FY2027 訂單近乎滿載、訂單延伸到 CY2028（公司，Interested-party）",
    LC + "：LightCounting 2025 光模組排名新易盛首次超過 Coherent 升到第二（Interested-party）")
inp("rel_28", "可插拔有效份額 CY2028 ÷ CY2026", "x", 0.85, 1.15, 1.35, "Assumed", "同上；1.6T 放量較三個月前預期提前（公司）", "中際旭創模型同期份額相對值 0.90（新易盛、Coherent 搶 1.6T）")
inp("rel_29", "可插拔有效份額 CY2029 ÷ CY2026", "x", 0.80, 1.15, 1.40, "Assumed", "同上", "—")
inp("cpo_27", "Coherent 占 CPO 光學金額 CY2027", "%", 0.0, 0.10, 0.25, "Assumed", F4 + "：CPO 營收自 2026 年 12 月季開始、scale-up CPO 自 CY2027 下半年；與 NVIDIA 合作超高功率 CW 雷射（公司）", "CPO、NPO 等整合光學增量市場逾 150 億美元（公司說法，Interested-party）")
inp("cpo_28", "Coherent 占 CPO 光學金額 CY2028", "%", 0.0, 0.15, 0.30, "Assumed", "同上", "—")
inp("cpo_29", "Coherent 占 CPO 光學金額 CY2029", "%", 0.0, 0.20, 0.35, "Assumed", "同上", "—")
inp("pxr_27", "單價係數 CY2027（相對網通模型每顆內容；基準 1）", "x", 0.90, 1.0, 1.05, "Assumed", "網通模型每顆光學內容已含世代單價；本係數只在 Sens 動（降價快於模型）", F4 + "：長約含固定價格與最低採購承諾（公司）")
inp("pxr_28", "單價係數 CY2028", "x", 0.80, 1.0, 1.05, "Assumed", "同上（與中際旭創模型相同）", "—")
inp("pxr_29", "單價係數 CY2029", "x", 0.72, 1.0, 1.05, "Assumed", "同上", "—")
inp("hw", "日曆年上半年占全年比重（資料中心；會計年度換算用）", "%", 0.42, 0.46, 0.50, "Assumed", "CY2026 推導：上半年（FY2026 Q3＋Q4）÷ 全年約 41%（高成長年下半年較重）", "CY2027 起成長放緩，比重向 50% 靠近；Sens 動 42%／50%")
inp("q2_dg", "FY2027 Q2（2026-12 季）資料中心季增", "%", 0.05, 0.12, 0.20, "Assumed", BZ + "：共識 Q2 營收 25.4 億美元（約 +10%，Interested-party）", "CPO 與 PhotonLink 初期營收自 12 月季開始（公司）")

sec("C. FY2026 Q3／Q4 實績與資料中心、通訊拆分（公司只揭露 Datacenter & Communications 合計）")
inp("q3_rev", "FY2026 Q3（2026-03 季）營收", "US$B", None, 1.8056, None, "Verified", PR + "：1,805.6 百萬美元", F3 + "：18.1 億美元")
inp("q4_rev", "FY2026 Q4（2026-06 季）營收", "US$B", None, 2.0455, None, "Verified", PR + "：2,045.5 百萬美元（年增 34%）", F4 + "：20.5 億美元")
inp("q3_dcc", "FY2026 Q3 資料中心與通訊（DCC）營收", "US$B", None, 1.3616, None, "Verified", PR + "：1,361.6 百萬美元", F3 + "：14 億美元（+37%）")
inp("q4_dcc", "FY2026 Q4 DCC 營收", "US$B", None, 1.615, None, "Verified", PR + "：1,615.0 百萬美元", F4 + "：16.2 億美元（+58.6%，占 79%）")
inp("q4_dcc25", "FY2025 Q4 DCC 營收（年增核對用）", "US$B", None, 1.0183, None, "Verified", PR + "：1,018.3 百萬美元", "—")
inp("q3_ind", "FY2026 Q3 工業營收", "US$B", None, 0.444, None, "Verified", PR + "：444.0 百萬美元", F3 + "：4.44 億美元")
inp("q4_ind", "FY2026 Q4 工業營收", "US$B", None, 0.4305, None, "Verified", PR + "：430.5 百萬美元（年減 15.8%；剔除出售事業約持平）", F4 + "：同值")
inp("q4_dg", "FY2026 Q4 資料中心季增", "%", None, 0.24, None, "Interested-party", F4 + "：資料中心季增 24%、年增 66%（法說會，公司）", WB4 + "：同值（逐字稿轉述）")
inp("q4_cg", "FY2026 Q4 通訊季增", "%", None, 0.11, None, "Interested-party", F4 + "：通訊（DCI、ZR／ZR+、泵浦雷射）季增 11%、年增 56%（公司）", WB4 + "：同值")
inp("q4_dy", "FY2026 Q4 資料中心年增（核對用）", "%", None, 0.66, None, "Interested-party", F4, WB4)
inp("q4_cy", "FY2026 Q4 通訊年增（核對用）", "%", None, 0.56, None, "Interested-party", F4, WB4)
inp("h1r26", "FY2026 上半年資料中心 ÷ 下半年資料中心", "x", 0.72, 0.79, 0.86, "Assumed", "Futurum FY2026 Q1（網址見 Sources）：FY2026 Q1 資料中心年增 23%；Q2 指引約季增 10%（公司）", "FY2027 Q1 資料中心年增逾 80%（公司指引）回推 Q1 約 0.67、Q2 約 0.74，對下半年 1.79 約 0.79（推導；資料中心金額公司未揭露：未找到）")

sec("D. FY2027 Q1 指引與 Q2 假設（＝CY2026 下半年）")
inp("q1_rev", "FY2027 Q1（2026-09 季）營收指引中點", "US$B", 2.2, 2.3, 2.4, "Interested-party", PR + "：22–24 億美元（公司指引）", INV + "：高於共識 21.3 億美元")
inp("q1_gm", "FY2027 Q1 non-GAAP 毛利率指引中點", "%", 0.395, 0.405, 0.415, "Interested-party", PR + "：39.5%–41.5%（公司指引）", "—")
inp("q1_opx", "FY2027 Q1 non-GAAP 營業費用指引中點", "US$B", 0.40, 0.41, 0.42, "Interested-party", PR + "：4.0–4.2 億美元（公司指引）", "—")
inp("q1_eps", "FY2027 Q1 non-GAAP EPS 指引中點", "US$", 1.85, 1.95, 2.05, "Interested-party", PR + "：1.85–2.05 美元（公司指引）", BZ + "：共識 1.93")
inp("q_cg", "通訊季增（FY2027 Q1、Q2）", "%", 0.0, 0.05, 0.10, "Assumed", "FY2026 Q4 通訊季增 11%（公司）", "—")
inp("q_ig", "工業季增（FY2027 Q1、Q2）", "%", -0.03, 0.0, 0.03, "Assumed", F4 + "：工業剔除出售事業約持平；半導體與顯示器設備季增（公司）", "—")

sec("E. 通訊與工業（年增；FY2028 起）——已出售事業排除")
inp("g_com_27", "通訊營收年增 FY2027", "%", 0.10, 0.20, 0.30, "Assumed", F4 + "：通訊 FY2026 Q4 年增 56%（DCI、scale-across、ZR／ZR+）", "OCS 光路交換潛在市場逾 40 億美元（約 2030，公司說法）")
inp("g_com_28", "通訊營收年增 FY2028", "%", 0.0, 0.12, 0.20, "Assumed", "同上", "—")
inp("g_com_29", "通訊營收年增 FY2029", "%", -0.05, 0.08, 0.15, "Assumed", "同上", "—")
inp("ind_run", "工業營收基期（FY2026 Q4 × 4；已排除出售事業）", "US$B", None, "=q4_ind*4", None, "Derived", F2 + "：慕尼黑工具事業 2026-01 出售（每季約 2,500 萬美元）；A&D 於 FY2026 出售（約 4 億美元，10-K）", AB + "：A&D 出售約 4 億美元（10-K 摘要）")
inp("g_ind_27", "工業營收 FY2027 ÷ 基期 − 1", "%", -0.05, 0.02, 0.08, "Assumed", "半導體設備回升；Thermadite（XPU 散熱材料）CY2027 下半年起量（公司）", "工業含雷射、工程材料與 SiC 基板（10-K）；SiC 單獨營收：未找到")
inp("g_ind_28", "工業營收年增 FY2028", "%", -0.05, 0.05, 0.10, "Assumed", "同上", "SiC 子公司 25% 由 DENSO、三菱電機持有（2023 年投資 10 億美元）")
inp("g_ind_29", "工業營收年增 FY2029", "%", -0.05, 0.05, 0.10, "Assumed", "同上", "—")

sec("F. 毛利率、費用、稅率、股數（non-GAAP）")
inp("gm_27", "non-GAAP 毛利率 FY2027", "%", 0.385, 0.410, 0.425, "Assumed", PR + "：FY2026 39.4%、Q4 40.2%；Q1 指引 39.5%–41.5%", INV + "：Barclays 指出低毛利光模組占比上升壓毛利（Interested-party）")
inp("gm_28", "non-GAAP 毛利率 FY2028", "%", 0.375, 0.420, 0.440, "Assumed", F4 + "：長期目標 42% 以上；六吋磷化銦產出 4 倍、成本約一半（公司）", F2 + "：1.6T 毛利率高於 800G（公司）")
inp("gm_29", "non-GAAP 毛利率 FY2029", "%", 0.365, 0.420, 0.450, "Assumed", "同上", "—")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.40, "=Q!B24", 0.90, "Derived", "FY2026 non-GAAP 營業費用年增 ÷ 營收年增（Q 頁）", WB4 + "：目標營業費用率 18%；FY2026 19.0%、Q4 18.4%（公司）")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.01, 0.0, 0.02, "Assumed", "無", "Sens 用 +2 個百分點")
inp("int_f", "non-GAAP 利息與其他淨支出（FY2027 起，每年）", "US$B", 0.04, 0.08, 0.12, "Assumed", PR + "：FY2026 1.181 億美元；長期借款 32.1 億、現金與短期投資 19.9 億美元（2026-06）", WB4 + "：FY2026 還債 5.13 億美元、槓桿 0.7x")
inp("tx", "non-GAAP 稅率", "%", 0.18, 0.19, 0.20, "Interested-party", PR + "：FY2027 Q1 指引 18%–20%；FY2026 19%（公司）", "—")
inp("nci", "少數股權調整（加回；每年）", "US$B", None, "=Q!C15", None, "Derived", "FY2026 non-GAAP 歸屬淨利 − 稅後稅前利益（Q 頁）；SiC 子公司少數股權分攤虧損", "—")
inp("sh_q4", "FY2026 Q4 non-GAAP 稀釋股數", "十億股", None, 0.2022, None, "Verified", PR + "：202.2 百萬股", TENK + "：普通股 195,832,246 股（2026-08-10）")
inp("sh_g", "股數年變動", "%", 0.0, 0.01, 0.03, "Assumed", "10-K 摘要（autobound，網址見 Sources）：NVIDIA 投資 20 億美元；FY2026 發行新股 19.99 億美元（8-K）", "股票薪酬稀釋；回購：未找到")
inp("mf", "AI 市場係數（Sens 用；CY2027 起）", "x", 0.8, 1.0, 1.2, "Assumed", "網通模型 CY2027 可插拔光模組 +84%", "±20%")

sec("G. FY2025、FY2026 實績（non-GAAP；US$B）")
for k, lab, v25, v26, n in (
        ("rev", "營收", 5.8101, 7.1182, "5,810.1／7,118.2 百萬美元"),
        ("dcc", "資料中心與通訊（DCC）營收", 3.7552, 5.2746, "3,755.2／5,274.6（FY2026 起新分部，FY2025 重編）"),
        ("ind", "工業營收（含已出售事業）", 2.0549, 1.8436, "2,054.9／1,843.6"),
        ("gp", "non-GAAP 毛利", 2.2023, 2.8065, "2,202.3／2,806.5"),
        ("opx", "non-GAAP 營業費用（研發＋銷管）", "=0.5543+0.611", "=0.6933+0.6563", "研發 554.3＋銷管 611.0／研發 693.3＋銷管 656.3"),
        ("int", "non-GAAP 利息與其他淨支出", 0.1673, 0.1181, "167.3／118.1"),
        ("tax", "non-GAAP 所得稅", 0.184, 0.2544, "184.0／254.4"),
        ("ni", "non-GAAP 歸屬 Coherent 淨利", 0.6926, 1.0966, "692.6／1,096.6"),
        ("eps", "non-GAAP 稀釋 EPS", 3.53, 5.61, "3.53／5.61 美元"),
        ("sh", "non-GAAP 稀釋股數", 0.1592, 0.1954, "159.2／195.4 百萬股")):
    inp(f"a25_{k}_in", f"FY2025 {lab}", "US$" if k == "eps" else ("十億股" if k == "sh" else "US$B"), None, v25, None, "Verified", PR + "：" + n.split("／")[0], "FY2025 EPS 已扣優先股股利（Q 頁推導）" if k == "eps" else "—")
    inp(f"a26_{k}_in", f"FY2026 {lab}", "US$" if k == "eps" else ("十億股" if k == "sh" else "US$B"), None, v26, None, "Verified", PR + "：" + n.split("／")[-1], WB4 + "：同值（法說會轉述）" if k in ("rev", "eps", "dcc") else "—")

sec("H. 共識、市場資料（只作 Recon 對帳，不作校準）")
inp("cons_rev27", "共識 FY2027 營收", "US$B", None, 10.63, None, "Interested-party", TK + "：106.3 億美元（2026-09-30，TIKR 彙整）", "06-30 時為 95.7")
inp("cons_rev28", "共識 FY2028 營收", "US$B", None, 14.76, None, "Interested-party", TK + "：147.6 億美元（2026-09-30）", "06-30 時為 129.2")
inp("cons_eps27", "共識 FY2027 EPS", "US$", None, 8.67, None, "Interested-party", MB + "：8.67（2026-10-23，MarketBeat 未標年度，以下一年度判讀）", BZ + "：四季加總 1.93＋2.16＋2.26＋2.50＝8.85")
inp("cons_eps27b", "共識 FY2027 EPS（四季加總）", "US$", None, "=1.93+2.16+2.26+2.5", None, "Interested-party", BZ + "：各季共識（2026-10-23）", "—")
inp("cons_eps28", "共識 FY2028 EPS", "US$", None, 13.07, None, "Interested-party", "MarketBeat（網址見 Sources）：13.07（2026-10-23）", "第二來源：未找到")
inp("cons_q1eps", "共識 FY2027 Q1 EPS", "US$", None, 1.93, None, "Interested-party", BZ, "—")
inp("lc_eo25", "新易盛 2025 光模組營收（LightCounting；Coherent 低於此值）", "US$B", None, 3.5, None, "Interested-party", "LightCounting（網址見 Sources）：新易盛 35 億美元，2025 年首次超過 Coherent", "Coherent 光模組營收金額：未找到")

sec("I. 價格與估值")
inp("px", "股價（2026-10-23 收盤）", "US$", None, 209.17, None, "Interested-party", MB + "：209.17（+0.65%）", BZ + "：209.29（10-23）")
inp("px_hi", "52 週高點", "US$", None, 440.0, None, "Interested-party", MB + "：52 週區間 123.75–440.00", GF + "：52 週高點 440.00")
for s, w in zip(SC, (0.3, 0.5, 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")

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
INROW = {}
rr = 4
for srow in rows:
    if srow[0] != "sec": INROW[srow[0]] = rr
    rr += 1
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績、FY2026 季度拆分與 FY2027 上半年 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與季度拆分（十億美元；EPS 為美元；股數為十億股；non-GAAP）。藍字＝輸入（In 頁）；黑字＝公式。FY(t)＝t−1 年 7 月～t 年 6 月"
Q["A2"] = "資料中心與通訊的金額公司未揭露：FY2026 Q4 資料中心＝由 Q3、Q4 的 DCC 合計與兩者季增率聯立解出；Q3＝Q4 ÷（1＋季增）。FY2027 Q1 資料中心＝營收指引中點 − 通訊 − 工業；Q2 依季增假設。FY2027 Q1 季報（2026-11-04）公布後以實績取代"
hdr(Q, 3, ["項目", "FY2025A", "FY2026A", "FY2026 Q3", "FY2026 Q4", "FY2027 Q1（指引）", "FY2027 Q2（假設）", "標記", "說明"])
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", note=""):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, note)
qrow(4, "營收", {"B": "=a25_rev_in", "C": "=a26_rev_in", "D": "=q3_rev", "E": "=q4_rev", "F": "=q1_rev", "G": "=G5+G8"}, tag="Verified／Interested-party（指引）")
qrow(5, "資料中心與通訊（DCC）", {"B": "=a25_dcc_in", "C": "=a26_dcc_in", "D": "=q3_dcc", "E": "=q4_dcc", "F": "=F6+F7", "G": "=G6+G7"}, tag="Verified")
qrow(6, "  資料中心（AI 光模組、雷射與光元件、OCS）", {"B": "=B5*C6/C5", "C": "=(D6+E6)*(1+h1r26)", "D": "=E6/(1+q4_dg)", "E": "=(D5-E5/(1+q4_cg))/(1/(1+q4_dg)-1/(1+q4_cg))", "F": "=F4-F7-F8", "G": "=F6*(1+q2_dg)"},
     tag="Derived／Assumed", note="Q4：聯立 Q3 DCC＝Q4 資料中心 ÷（1＋24%）＋Q4 通訊 ÷（1＋11%）；FY2026 全年＝下半年 ×（1＋上半年 ÷ 下半年比）；FY2025 假設與 FY2026 相同結構")
qrow(7, "  通訊（電信、DCI、ZR、泵浦雷射）", {"B": "=B5-B6", "C": "=C5-C6", "D": "=D5-D6", "E": "=E5-E6", "F": "=E7*(1+q_cg)", "G": "=F7*(1+q_cg)"}, tag="Derived")
qrow(8, "工業（雷射、工程材料、SiC）", {"B": "=a25_ind_in", "C": "=a26_ind_in", "D": "=q3_ind", "E": "=q4_ind", "F": "=E8*(1+q_ig)", "G": "=F8*(1+q_ig)"}, tag="Verified／Assumed", note="FY2025、FY2026 含已出售的 A&D 與慕尼黑工具事業；Q4 起不含")
qrow(9, "non-GAAP 毛利", {"B": "=a25_gp_in", "C": "=a26_gp_in", "F": "=F4*q1_gm"}, tag="Verified")
qrow(10, "non-GAAP 毛利率", {c: f"={c}9/{c}4" for c in "BCF"}, fmt="0.0%", tag="Derived")
qrow(11, "non-GAAP 營業費用", {"B": "=a25_opx_in", "C": "=a26_opx_in", "F": "=q1_opx"}, tag="Verified")
qrow(12, "non-GAAP 營業利益", {c: f"={c}9-{c}11" for c in "BCF"}, tag="Derived", note="FY2026 公司揭露 1,456.9 百萬美元")
qrow(13, "non-GAAP 利息與其他淨支出", {"B": "=a25_int_in", "C": "=a26_int_in", "F": "=int_f/4"}, tag="Verified")
qrow(14, "non-GAAP 稅率", {"B": "=a25_tax_in/(B12-B13)", "C": "=a26_tax_in/(C12-C13)", "F": "=tx"}, fmt="0.0%", tag="Derived")
qrow(15, "少數股權調整（歸屬淨利 −（營業利益 − 利息）×（1 − 稅率））", {"B": "=a25_ni_in-(B12-B13)*(1-B14)", "C": "=a26_ni_in-(C12-C13)*(1-C14)", "F": "=nci/4"}, tag="Derived")
qrow(16, "non-GAAP 歸屬 Coherent 淨利", {"B": "=a25_ni_in", "C": "=a26_ni_in", "F": "=(F12-F13)*(1-F14)+F15"}, tag="Verified／Derived")
qrow(17, "優先股股利（EPS 計算扣除）", {"B": "=a25_ni_in-a25_eps_in*a25_sh_in", "C": "=0", "F": "=0"}, tag="Derived", note="FY2025 由淨利 − EPS × 股數推得；FY2026 起優先股已轉換，無扣除")
qrow(18, "non-GAAP 稀釋股數", {"B": "=a25_sh_in", "C": "=a26_sh_in", "E": "=sh_q4", "F": "=sh_q4"}, fmt="0.0000", tag="Verified")
qrow(19, "non-GAAP 稀釋 EPS", {c: f"=({c}16-{c}17)/{c}18" for c in "BCF"}, fmt="0.00", tag="Derived", note="FY2027 Q1 指引中點 1.95")
Q.cell(21, 1, "推導量").font = BOLD
Q.cell(22, 1, "FY2026 營收年增"); Q["B22"] = "=C4/B4-1"
Q.cell(23, 1, "FY2026 營業費用年增"); Q["B23"] = "=C11/B11-1"
Q.cell(24, 1, "營業費用彈性（營業費用年增 ÷ 營收年增；In 頁 phi 的依據）"); Q["B24"] = "=B23/B22"
Q.cell(25, 1, "CY2026 資料中心（FY2026 Q3＋Q4＋FY2027 Q1＋Q2；份額校準基礎）"); Q["B25"] = "=D6+E6+F6+G6"
Q.cell(26, 1, "CY2026 下半年資料中心（＝FY2027 上半年）"); Q["B26"] = "=F6+G6"
Q.cell(27, 1, "CY2026 上半年資料中心 ÷ 全年（hw 的依據）"); Q["B27"] = "=(D6+E6)/B25"
Q.cell(28, 1, "核對：FY2025 Q4 DCC 由年增率回推 ÷ 實績 − 1"); Q["B28"] = "=(E6/(1+q4_dy)+E7/(1+q4_cy))/q4_dcc25-1"
Q.cell(29, 1, "FY2027 Q1 EPS（模型，指引中點）÷ 公司 EPS 指引中點 − 1"); Q["B29"] = "=F19/q1_eps-1"
Q.cell(30, 1, "FY2026 資料中心 ÷ DCC"); Q["B30"] = "=C6/C5"
for rr_, f in ((22, "0.0%"), (23, "0.0%"), (24, "0.00"), (25, "#,##0.000"), (26, "#,##0.000"), (27, "0.0%"), (28, "0.0%"), (29, "0.0%"), (30, "0.0%")): Q[f"B{rr_}"].number_format = f
for nm, rr_ in (("rev", 4), ("dc", 6), ("com", 7), ("ind", 8), ("gp", 9), ("opx", 11), ("int", 13), ("ni", 16), ("pref", 17), ("sh", 18), ("eps", 19)):
    defname(f"a25_{nm}", f"Q!$B${rr_}"); defname(f"a26_{nm}", f"Q!$C${rr_}")
defname("dc_cy26", "Q!$B$25"); defname("h1_27", "Q!$B$26")
Q.column_dimensions["A"].width = 60
for c in "BCDEFG": Q.column_dimensions[c].width = 15
Q.column_dimensions["H"].width = 26; Q.column_dimensions["I"].width = 90
Q.freeze_panes = "B4"

# ---------------- 引擎 ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
BU = (2026, 2027, 2028, 2029)
umx = lambda y, s: {2026: "1", 2027: "1", 2028: f"um28_{s}", 2029: f"um29_{s}"}[y]
DEFP = dict(scn="A", rel="rel", cpo="cpo", pxr="pxr", gm="gm", mf="1", hw="hw", gc="g_com", opxa="0")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]; hw = P["hw"]
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026A（錨點）", "FY2027", "FY2028", "FY2029", "說明"]); r += 1
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
    F = BU[1:]
    sec_("A. 日曆年 AI 光學（本節 D–G 欄＝CY2026–CY2029；D 欄＝CY2026 校準年）")
    row("mpl", "AI 可插拔光模組市場 × 情境顆數係數（日曆年）", "US$B", {c_(y): f"=mpl_{y % 100}*{umx(y, s)}" + ("" if y == 2026 else f"*{P['mf']}") for y in BU}, note="本節欄位為日曆年")
    row("mcp", "AI CPO 光學市場 × 情境顆數係數（日曆年）", "US$B", {c_(y): f"=mcp_{y % 100}*{umx(y, s)}" for y in BU})
    row("sh0", "Coherent 可插拔有效份額（CY2026 校準）＝ CY2026 資料中心 ÷ 可插拔市場", "%", {"D": f"=dc_cy26/D{R['mpl']}"}, fmt="0.0%",
        note="Derived：CY2026 CPO 營收視為 0；含賣給其他模組廠的雷射元件，屬「有效」份額")
    row("rel", "份額相對值（÷ CY2026）", "x", {"D": "=1", **{c_(y): f"={P['rel']}_{y % 100}" for y in F}}, fmt="0.00")
    row("sh", "可插拔有效份額（日曆年）", "%", {c_(y): f"=$D${R['sh0']}*{c_(y)}{R['rel']}" for y in BU}, fmt="0.0%")
    row("pxr", "單價係數（相對網通模型）", "x", {"D": "=1", **{c_(y): f"={P['pxr']}_{y % 100}" for y in F}}, fmt="0.00")
    row("cps", "Coherent 占 CPO 光學金額", "%", {"D": "=0", **{c_(y): f"={P['cpo']}_{y % 100}" for y in F}}, fmt="0.0%")
    row("dcy", "資料中心營收（日曆年）", "US$B", {"D": "=dc_cy26", **{c_(y): f"={c_(y)}{R['mpl']}*{c_(y)}{R['sh']}*{c_(y)}{R['pxr']}+{c_(y)}{R['mcp']}*{c_(y)}{R['cps']}" for y in F}})
    sec_("B. 會計年度營收（本節起 C–G 欄＝FY2025–FY2029）")
    row("dc", "資料中心", "US$B", {"C": "=a25_dc", "D": "=a26_dc",
        "E": f"=h1_27+{hw}*E{R['dcy']}", "F": f"=(1-{hw})*E{R['dcy']}+{hw}*F{R['dcy']}", "G": f"=(1-{hw})*F{R['dcy']}+{hw}*G{R['dcy']}"},
        note="FY2027＝CY2026 下半年（Q 頁）＋ hw × CY2027；FY2028＝（1 − hw）× CY2027 ＋ hw × CY2028")
    row("com", "通訊", "US$B", {"C": "=a25_com", "D": "=a26_com", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gc']}_{y % 100})" for y in F}})
    row("ind", "工業（FY2027 起排除已出售事業）", "US$B", {"C": "=a25_ind", "D": "=a26_ind", "E": "=ind_run*(1+g_ind_27)", **{c_(y): f"={c_(y - 1)}{r}*(1+g_ind_{y % 100})" for y in (2028, 2029)}})
    row("rev", "營收", "US$B", {c_(y): f"={c_(y)}{R['dc']}+{c_(y)}{R['com']}+{c_(y)}{R['ind']}" for y in FYS})
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("dcs", "資料中心 ÷ 營收", "%", {c_(y): f"={c_(y)}{R['dc']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    sec_("C. 損益（non-GAAP）")
    row("gm", "毛利率", "%", {"C": "=a25_gp/a25_rev", "D": "=a26_gp/a26_rev", **{c_(y): f"={P['gm']}_{y % 100}" for y in F}}, fmt="0.0%",
        note="FY2027 起為輸入路徑（六吋磷化銦降本、1.6T 上推 vs 光模組占比上升下壓）")
    row("gp", "毛利", "US$B", {c_(y): f"={c_(y)}{R['rev']}*{c_(y)}{R['gm']}" for y in FYS})
    row("opb", "營業費用（基準路徑）", "US$B", {"C": "=a25_opx", "D": "=a26_opx", **{c_(y): f"={c_(y - 1)}{r}*(1+MAX(0,phi*{c_(y)}{R['g']}))" for y in F}})
    row("opx", "營業費用", "US$B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", **{c_(y): f"={c_(y)}{R['opb']}+{P['opxa']}*{c_(y)}{R['rev']}" for y in F}})
    row("op", "營業利益", "US$B", {c_(y): f"={c_(y)}{R['gp']}-{c_(y)}{R['opx']}" for y in FYS})
    row("opm", "營業利益率", "%", {c_(y): f"={c_(y)}{R['op']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("int", "利息與其他淨支出", "US$B", {"C": "=a25_int", "D": "=a26_int", **{c_(y): "=int_f" for y in F}})
    row("pre", "稅前利益", "US$B", {c_(y): f"={c_(y)}{R['op']}-{c_(y)}{R['int']}" for y in FYS})
    row("ni", "歸屬 Coherent 淨利", "US$B", {"C": "=a25_ni", "D": "=a26_ni", **{c_(y): f"={c_(y)}{R['pre']}*(1-tx)+nci" for y in F}})
    row("nim", "淨利率", "%", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("pref", "優先股股利", "US$B", {"C": "=a25_pref", **{c_(y): "=0" for y in BU}}, fmt="#,##0.000")
    row("sh_", "稀釋股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", "E": "=sh_q4*(1+sh_g)", **{c_(y): f"={c_(y - 1)}{r}*(1+sh_g)" for y in (2028, 2029)}}, fmt="0.0000")
    row("eps", "EPS", "US$", {c_(y): f"=({c_(y)}{R['ni']}-{c_(y)}{R['pref']})/{c_(y)}{R['sh_']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — Coherent 損益（US$B；non-GAAP）。FY2025A、FY2026A（錨點）為實績，FY2027 起為模型。FY(t)＝t−1 年 7 月～t 年 6 月。三情境各一個完整推導區塊（本頁下方），頂端為摘要。FY2027 三情境相同（顆數係數從 CY2028 開始）"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A：基準（顆數係數 1.0）", "B": "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85", "C": "情境 C：CY2028–29 顆數 ×0.6"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026A（錨點）", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "US$B", "#,##0.00"), ("g", "  營收年增", "%", "0.0%"), ("dc", "  資料中心", "US$B", "#,##0.00"), ("com", "  通訊", "US$B", "#,##0.00"),
                                ("ind", "  工業", "US$B", "#,##0.00"), ("dcs", "  資料中心 ÷ 營收", "%", "0.0%"), ("gm", "毛利率", "%", "0.0%"), ("opm", "營業利益率", "%", "0.0%"),
                                ("ni", "歸屬淨利", "US$B", "#,##0.00"), ("eps", "EPS", "US$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS:
            if key in ("g", "epsg") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型與實績（校準恆等式）").font = BOLD; sr += 1
PL.cell(sr, 1, "  情境 A：營收差（US$B）"); PL[f"D{sr}"] = f"=D{PB['A']['rev']}-a26_rev_in"; PL[f"D{sr}"].number_format = "0.0000"; CHK_ROW = sr; sr += 1
PL.cell(sr, 1, "  情境 A：營業利益差（US$B；公司揭露 1.4569）"); PL[f"D{sr}"] = f"=D{PB['A']['op']}-1.4569"; PL[f"D{sr}"].number_format = "0.0000"; sr += 1
PL.cell(sr, 1, "  情境 A：EPS 差（公司揭露 5.61）"); PL[f"D{sr}"] = f"=D{PB['A']['eps']}-a26_eps_in"; PL[f"D{sr}"].number_format = "0.0000"; sr += 1
PL.cell(sr, 1, "  情境 A：CY2026 資料中心 − 可插拔市場 × 有效份額"); PL[f"D{sr}"] = f"=D{PB['A']['dcy']}-D{PB['A']['mpl']}*D{PB['A']['sh']}"; PL[f"D{sr}"].number_format = "0.0000"
assert sr < 60, sr
PL.column_dimensions["A"].width = 66; PL.column_dimensions["B"].width = 12; PL.column_dimensions["H"].width = 80
for c in "CDEFG": PL.column_dimensions[c].width = 14

# ---------------- Units：市場與份額 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — Coherent 在 AI 光模組的有效份額（日曆年）。市場取自 markets/network（加速器顆數 × 每顆光學內容；美元）；份額口徑與中際旭創模型一致（資料中心營收 ÷ 可插拔市場），兩者相加須低於 100%"
A_ = PB["A"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("mpl", "AI 可插拔光模組市場（情境 A）", "US$B", lambda y: f"=PL!{FC[y]}{A_['mpl']}", "#,##0.0", "markets/network Mkt 頁"),
        ("mcp", "AI CPO 光學市場（情境 A）", "US$B", lambda y: f"=PL!{FC[y]}{A_['mcp']}", "#,##0.0", "同上"),
        ("sh", "Coherent 可插拔有效份額", "%", lambda y: f"=PL!{FC[y]}{A_['sh']}", "0.0%", "CY2026 以 FY2026 下半年實績（推導）＋FY2027 Q1 指引＋Q2 假設校準"),
        ("inl", "中際旭創可插拔有效份額（innolight/dist）", "%", lambda y: f"=inl_{y % 100}", "0.0%", "唯讀上游"),
        ("tot", "兩者合計", "%", lambda y: f"={cyc(y)}{ur - 2}+{cyc(y)}{ur - 1}", "0.0%", "須低於 100%；其餘由新易盛、Lumentum、Fabrinet 代工客戶等分配"),
        ("dcy", "Coherent 資料中心營收（情境 A；日曆年）", "US$B", lambda y: f"=PL!{FC[y]}{A_['dcy']}", "#,##0.00", ""),
        ("oth", "其他廠商＝ 可插拔＋CPO − 中際旭創 − Coherent", "US$B", lambda y: f"={cyc(y)}4+{cyc(y)}5-{cyc(y)}7*{cyc(y)}4-{cyc(y)}9", "#,##0.0", "中際旭創以可插拔份額近似（其 CPO 份額不扣）"),
        ("itr", "Coherent 資料中心營收 ÷ AI IT 資本支出", "%", lambda y: f"=PL!{FC[y]}{A_['dcy']}/itc_{y % 100}", "0.00%", "markets/network 的 AI IT 資本支出")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); UR[key] = ur; ur += 1
assert UR["sh"] == 6 and UR["inl"] == 7 and UR["tot"] == 8 and UR["dcy"] == 9
UN.cell(ur, 1, "核對：兩者合計最大值 < 100%（1＝通過）"); UN[f"C{ur}"] = f"=IF(MAX(C{UR['tot']}:F{UR['tot']})<1,1,0)"; UR["chk"] = ur; ur += 2
UN.cell(ur, 1, "讀法：Coherent 有效份額含賣給其他模組廠的雷射與光元件（與中際旭創的採購成本部分重疊），所以是「有效」份額而非出貨份額；份額相對值與 CPO 份額是 Assumed，是最需要每季檢查的輸入。").font = BOLD
UR["note"] = ur
UN.column_dimensions["A"].width = 64; UN.column_dimensions["B"].width = 10
for c in "CDEF": UN.column_dimensions[c].width = 14
UN.column_dimensions["G"].width = 70

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（美元；non-GAAP EPS）。結論用目標價（第 0 節）、「現價 ÷ FY2027、FY2028 EPS（模型）」、市場隱含 EPS 與十年 P/E 規則值。十年 P/E 為 GAAP 年底 P/E（II-VI 2022 年改名 Coherent，同一法人）；GAAP 虧損年度不列入"
V["A2"] = "循環位置：Coherent 不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照中際旭創、Arista）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
# 第 0 節置於第 4–12 列（在下方 VAL 決定後填入）
V.cell(15, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, 16, ["項目", "A", "B", "C", "說明"])
V.cell(17, 1, "現價（2026-10-23 收盤）"); V["B17"] = "=px"; V["B17"].number_format = "0.00"
V.cell(18, 1, "現價 ÷ FY2026 EPS（實績）"); V.cell(19, 1, "現價 ÷ FY2027 EPS（模型）"); V.cell(20, 1, "現價 ÷ FY2028 EPS（模型）"); V.cell(21, 1, "現價 ÷ FY2029 EPS（模型）")
for j, s in enumerate(SC):
    c = "BCD"[j]
    for rr_, y in ((18, 2026), (19, 2027), (20, 2028), (21, 2029)):
        V[f"{c}{rr_}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{rr_}"].number_format = "0.0"
V.cell(22, 1, "現價 ÷ 共識未來 12 月 EPS（9/12 × FY2027＋3/12 × FY2028）"); V["B22"] = "=px/(cons_eps27*9/12+cons_eps28*3/12)"; V["B22"].number_format = "0.0"
V.cell(23, 1, "現價 ÷ 52 週高點 − 1"); V["B23"] = "=px/px_hi-1"; V["B23"].number_format = "0%"
V.cell(24, 1, "現價 ÷ FY2027 Q1 EPS 指引中點年化"); V["B24"] = "=px/(q1_eps*4)"; V["B24"].number_format = "0.0"
V["E19"] = "FY2027 三情境相同（顆數係數從 CY2028 開始）"; V["E24"] = "1.95 × 4"
V.cell(26, 1, "2. 循環位置（本模型營收最高的會計年度）").font = BOLD
hdr(V, 27, ["情境", "營收最高年（FY2026–FY2029）", "FY2027 判定", "規則倍數（P/E）", "見頂年是否未知", "說明"])
for i, s in enumerate(SC, 28):
    rr_ = f"PL!D{PB[s]['rev']}:G{PB[s]['rev']}"
    V.cell(i, 1, s)
    V.cell(i, 2, f"=INDEX({{2026,2027,2028,2029}},MATCH(MAX({rr_}),{rr_},0))")
    V.cell(i, 3, '=IF(B{0}=2029,"中段",'.format(i) + CP.excel_rule("2027", f"B{i}") + ")")
    V.cell(i, 4, f'=IF(C{i}="見頂",pe_peak,IF(C{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 5, f'=IF(B{i}=2029,"是：營收到觀測期末仍在增加","否")')
    V.cell(i, 6, "tools/cycle_position.py excel_rule；營收最高年取代設備市場見頂年" if s == "A" else "同上")
V.cell(32, 1, "3. 規則值隱含股價（美元；約 FY2027 年底＝2027-06，未折現）＝ FY2027 EPS × 規則倍數").font = BOLD
hdr(V, 33, ["情境", "FY2027 EPS", "規則倍數", "隱含股價", "相對現價", "現價 ÷ FY2027 EPS"])
VAL = {}
for j, s in enumerate(SC):
    i = 34 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 3, f"=D{28 + j}").number_format = "0.0"
    V.cell(i, 4, f"=B{i}*C{i}").number_format = "#,##0"
    V.cell(i, 5, f"=D{i}/px-1").number_format = "0%"
    V.cell(i, 6, f"=px/B{i}").number_format = "0.0"
    for c in range(1, 7): V.cell(i, c).fill = KEY
    VAL[s] = i
V.cell(38, 1, "4. 核心問題：市場隱含的 FY2028 EPS（＝現價 ÷ 假設倍數）與模型的距離").font = BOLD
hdr(V, 39, ["假設倍數（÷ FY2028 EPS）", "市場隱含 FY2028 EPS", "模型 FY2028 EPS（A）", "隱含 ÷ 模型 A − 1", "模型 FY2028 EPS（C）", "隱含 ÷ 模型 C − 1"])
vr = 40; IMPL = {}
for m in (10, 15, 20, "=pe_peak"):
    V.cell(vr, 1, m).number_format = "0.0"
    if not isinstance(m, str): V.cell(vr, 1).font = BLUE
    V.cell(vr, 2, f"=px/A{vr}").number_format = "0.00"
    V.cell(vr, 3, f"=PL!F{SUMR[('A', 'eps')]}").number_format = "0.00"
    V.cell(vr, 4, f"=B{vr}/C{vr}-1").number_format = "0%"
    V.cell(vr, 5, f"=PL!F{SUMR[('C', 'eps')]}").number_format = "0.00"
    V.cell(vr, 6, f"=B{vr}/E{vr}-1").number_format = "0%"
    IMPL[m] = vr; vr += 1
V.cell(vr, 1, "讀法：第四列＝十年 P/E 25 分位（見頂倍數）。現價約等於 15 倍模型 FY2028 EPS（A）；若市場對 FY2028 給見頂倍數，隱含 FY2028 EPS 比模型 A 低四成以上——比 Sens 任何單一不利案例都低，要成立需要份額、單價、毛利率與 AI 市場同時不利。"); READ_ROW = vr
vr += 2
V.cell(vr, 1, "本版未做先驗：無 Andy 合理 P/E 對照。"); NOPRIOR_ROW = vr
assert vr < 50, vr
# 十年 GAAP 年底 P/E（2016–2025）；負值（GAAP 虧損）不列入
PE_CMC = [26.7, 34.3, 18.3, None, 42.2, 29.1, None, None, None, 158.0]
PE_MT = [26.95, 35.57, 18.98, None, 44.68, 31.2, None, None, None, None]
block_end = VR.add_rule_block(V, 50, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "companiesmarketcap（年底 P/E）", "vals": PE_CMC, "tag": "Interested-party（資料商；GAAP TTM；2019、2022–2024 為負值不列入）"},
    {"label": "macrotrends（12 月底 P/E）", "vals": PE_MT, "tag": "Interested-party（資料商；GAAP TTM；2025 年底未列，未找到）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（A）", "basis": f"=B{VAL['A']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（A）", "basis": f"=B{VAL['A']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（A）", "basis": f"=B{VAL['A']}"}]}],
    "#,##0.00", "規則（Andy 2026-10-13）：合理倍數＝年底 P/E 依循環位置取分位。兩個資料商來源（GAAP 落後 P/E）；GAAP 虧損年度（2019、2022–2024）不列入，有效 5–6 年。GAAP 倍數套在 non-GAAP EPS 上（non-GAAP EPS 高於 GAAP），且 2025 年 158x 為 GAAP 盈餘壓低所致，偏誤方向為高估；前瞻 P/E 歷史待 Bloomberg。")
rule_rows = [rw for rw in range(50, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")
RULE_INTRO = 51
# 0. 目標價與上檔空間（Andy 2026-10-25 要求；列標籤依工作單逐字）
V.cell(4, 1, "0. 目標價與上檔空間").font = BOLD
hdr(V, 5, ["項目", "數值", "", "", "說明"])
V.cell(6, 1, "目標價（12 個月，情境加權）"); V["B6"] = f"=tp_wA*D{VAL['A']}+tp_wB*D{VAL['B']}+tp_wC*D{VAL['C']}"; V["B6"].number_format = "#,##0"
V["E6"] = "權重 × 各情境隱含價（第 3 節規則值）"
for c in "AB": V[f"{c}6"].fill = KEY
V.cell(7, 1, "上檔空間（目標價 ÷ 現價 − 1）"); V["B7"] = "=B6/px-1"; V["B7"].number_format = "0.0%"; V["E7"] = "現價＝In 頁 px（2026-10-23 收盤）"
for rr_, s, lab in ((8, "A", "情境 A 隱含價（樂觀）"), (9, "B", "情境 B 隱含價（基準）"), (10, "C", "情境 C 隱含價（下檔）")):
    V.cell(rr_, 1, lab); V[f"B{rr_}"] = f"=D{VAL[s]}"; V[f"B{rr_}"].number_format = "#,##0"
    V[f"E{rr_}"] = f'="FY2027 EPS "&TEXT(B{VAL[s]},"0.00")&" × "&C{28 + SC.index(s)}&" 判定倍數 "&TEXT(C{VAL[s]},"0.0")&"x"'
V.cell(11, 1, "情境權重 A／B／C"); V["B11"] = "=tp_wA"; V["C11"] = "=tp_wB"; V["D11"] = "=tp_wC"
for c in "BCD": V[f"{c}11"].number_format = "0%"
V["E11"] = '="合計 "&TEXT(B11+C11+D11,"0%")&"；In 頁 tp_wA／tp_wB／tp_wC（Assumed，Andy 可改）"'
V.cell(12, 1, "目標價方法")
V["B12"] = ("P/E：各情境 FY2027（2027-06 結束）non-GAAP EPS × 十年（2016–2025 年底）GAAP P/E 依循環位置取分位（營收最高年規則：仍成長＝中段＝中位數；見頂＝25 分位）。"
            "偏誤方向為高估：GAAP 倍數套在較高的 non-GAAP EPS 上，且 2025 年 158x 為 GAAP 盈餘壓低的離群值；隱含價約為 2027 年中的價格、未折現。FY2027 EPS 三情境相同（顆數係數從 CY2028 起），且三情境營收最高年都在 FY2027 之後（皆判中段），所以三個隱含價相同；情境差異見第 1、4 節的 FY2028。")
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 18

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；有效份額在每個案例內依 CY2026 重算，FY2026 恆等於實績）：對 FY2027、FY2028 的營收、資料中心營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 資料中心", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 資料中心", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("份額流失（低）", "份額相對值 CY2027 0.90、CY2028 0.85、CY2029 0.80（新易盛、中際旭創搶 1.6T）", dict(rel="RLO")),
         ("份額大增（高）", "份額相對值 1.25／1.35／1.40（磷化銦產能倍增）", dict(rel="RHI")),
         ("降價快於模型", "單價係數 CY2027 0.90、CY2028 0.80、CY2029 0.72", dict(pxr="PXLO")),
         ("毛利率 低（光模組占比上升）", "毛利率 FY2027 38.5%、FY2028 37.5%、FY2029 36.5%", dict(gm="GMLO")), ("毛利率 高（六吋磷化銦）", "毛利率 42.5%／44%／45%", dict(gm="GMHI")),
         ("CPO 份額 0", "Coherent 不參與 CPO 光學", dict(cpo="CPLO")), ("CPO 份額 高", "CY2027 25%、CY2028 30%、CY2029 35%", dict(cpo="CPHI")),
         ("AI 市場 ×0.8", "CY2027 起可插拔光模組市場 ×0.8", dict(mf="mf_lo")), ("AI 市場 ×1.2", "×1.2", dict(mf="mf_hi")),
         ("上半年比重 42%", "日曆年上半年占 42%（FY 認列較晚）", dict(hw="hw_lo")), ("上半年比重 50%", "50%", dict(hw="hw_hi")),
         ("通訊成長 低", "通訊年增 10%／0%／−5%", dict(gc="GCLO")),
         ("營業費用率 +2pt", "營業費用額外 +2% × 營收", dict(opxa="opx_adj_hi"))]
for nm, base, suf, ys in (("RLO", "rel", "_lo", (27, 28, 29)), ("RHI", "rel", "_hi", (27, 28, 29)), ("PXLO", "pxr", "_lo", (27, 28, 29)), ("GMLO", "gm", "_lo", (27, 28, 29)), ("GMHI", "gm", "_hi", (27, 28, 29)),
                          ("CPLO", "cpo", "_lo", (27, 28, 29)), ("CPHI", "cpo", "_hi", (27, 28, 29)), ("GCLO", "g_com", "_lo", (27, 28, 29))):
    for y in ys:
        wb.defined_names[f"{nm}_{y}"] = DefinedName(f"{nm}_{y}", attr_text=wb.defined_names[f"{base}_{y}{suf}"].attr_text)
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.00"), ("D", "dc", 2027, "#,##0.00"), ("E", "gm", 2027, "0.0%"), ("F", "eps", 2027, "0.00"),
                           ("H", "rev", 2028, "#,##0.00"), ("I", "dc", 2028, "#,##0.00"), ("J", "gm", 2028, "0.0%"), ("K", "eps", 2028, "0.00")):
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
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：FY2027 三顆數情境相同；B、C 的差異從 FY2028 起（FY2028 只含半年 CY2028）。FY2027 上半年＝指引＋假設，所以份額、市場等案例對 FY2027 只影響下半年；「上半年比重」案例顯示日曆年換算會計年度的口徑影響。")
SENS_NOTE = nn + 3
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 56
for c in "CDEFGHIJKL": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 共識、研調與公司指引。只作事後對照，不為對上數字而調參數。共識來自彙整網站（MarketBeat、Benzinga、TIKR），家數與區間未找到"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4; RCR = {}
def rec(label, ref, getter, src, note="", fmt="#,##0.00", pt=False, gap=True, key=None):
    global rr_
    RC.cell(rr_, 1, label)
    if ref is not None: RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        if ref is not None and gap:
            L = chr(64 + cv)
            RC.cell(rr_, cd, f"={L}{rr_}-B{rr_}" if pt else f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note)
    if key: RCR[key] = rr_
    rr_ += 1
pl = lambda key, y: (lambda s: f"=PL!{FC[y]}{PB[s][key]}")
RC.cell(rr_, 1, "A. 共識（校準未使用）").font = BOLD; rr_ += 1
rec("FY2027 營收 vs 共識", "=cons_rev27", pl("rev", 2027), "TIKR（2026-09-30）", "", key="rev27")
rec("FY2027 EPS vs 共識", "=cons_eps27", pl("eps", 2027), "MarketBeat（2026-10-23）", "", key="eps27")
rec("FY2027 EPS vs 共識（四季加總）", "=cons_eps27b", pl("eps", 2027), "Benzinga 各季共識", "", key="eps27b")
rec("FY2028 營收 vs 共識", "=cons_rev28", pl("rev", 2028), "TIKR（2026-09-30）", "")
rec("FY2028 EPS vs 共識", "=cons_eps28", pl("eps", 2028), "MarketBeat（2026-10-23）", "", key="eps28")
RC.cell(rr_, 1, "B. 公司指引與校準").font = BOLD; rr_ += 1
rec("FY2027 Q1 EPS（模型，指引中點的營收、毛利率、費用）vs 公司 EPS 指引中點", "=q1_eps", lambda s: "=Q!F19", "公司指引（2026-08-12）", "費用與稅率皆取指引；股數用 Q4 稀釋股數", key="q1")
rec("FY2027 Q1 EPS（模型）vs 共識", "=cons_q1eps", lambda s: "=Q!F19", "Benzinga", "")
rec("FY2025 Q4 DCC 由年增率回推 vs 實績（季增拆分的一致性）", "=q4_dcc25", lambda s: "=E6/(1+q4_dy)+E7/(1+q4_cy)".replace("E6", "Q!E6").replace("E7", "Q!E7"), "公司", "差距小＝資料中心、通訊拆分與兩組成長率一致", fmt="#,##0.000")
RC.cell(rr_, 1, "C. 份額與市場").font = BOLD; rr_ += 1
rec("CY2026 Coherent＋中際旭創可插拔有效份額", None, lambda s: f"=Units!C{UR['tot']}", "Units 頁", "須低於 100%", fmt="0.0%")
rec("CY2028 Coherent＋中際旭創可插拔有效份額", None, lambda s: f"=Units!E{UR['tot']}", "Units 頁（情境 A）", "", fmt="0.0%")
rec("FY2026 資料中心（含元件）vs 新易盛 2025 光模組營收", "=lc_eo25", pl("dc", 2026), "LightCounting", "LightCounting 只計光模組；Coherent 2025 光模組低於新易盛，資料中心含雷射元件與 OCS", fmt="#,##0.00")
rec("FY2027 資料中心 ÷ CY2027 AI IT 資本支出", None, lambda s: f"=PL!E{PB[s]['dc']}/itc_27", "markets/network", "每 1 美元 AI IT 資本支出分到的 Coherent 資料中心營收", fmt="0.00%")
RC.cell(rr_, 1, "D. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以資料中心營收 ÷ 部署 GW 計算每 GW 營收"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.cell(rr_, 1, "FY2027 Q1 季報（2026-11-04 盤後）：公布後以實績取代 Q 頁 Q1 指引，並更新資料中心拆分"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC_LAST = rr_ - 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝SEC 8-K 新聞稿與 10-K；法說會內容經第三方轉述，待人工複核）"; SO["A1"].font = BOLD
srcs = [("公司：FY2026 第四季與全年財報新聞稿（8-K Exhibit 99.1；分部營收、non-GAAP 調節、Q1 指引、股數）", PR),
        ("公司：FY2026 10-K（分部定義：工業含 SiC；普通股股數）", TENK),
        ("二手：10-K 摘要（NVIDIA 投資 20 億美元、A&D 出售約 4 億美元、兩大客戶各逾 10%）", AB),
        ("二手：Futurum FY2026 Q4（資料中心、通訊季增與年增；磷化銦產能；CPO 時程；FY2027 展望）", F4),
        ("二手：Webull 法說會摘要（同上；營業費用率；FY2027 訂單）", WB4),
        ("二手：Futurum FY2026 Q1、Q2、Q3（資料中心年增、慕尼黑工具出售）", F1 + "；" + F2 + "；" + F3),
        ("媒體：Investing.com（Q1 指引 vs 共識；Barclays 毛利率評論）", INV),
        ("共識：MarketBeat（FY2027／FY2028 EPS、股價 10-23、52 週區間）", MB),
        ("共識：Benzinga（各季 EPS 共識、股價）", BZ),
        ("共識：TIKR（FY2027／FY2028 營收共識，2026-09-30）", TK),
        ("研調：LightCounting 2025 光模組廠商排名", LC),
        ("資料商：companiesmarketcap 年底 P/E", CMC),
        ("資料商：macrotrends 12 月底 P/E", MT),
        ("價格：digrin 月底價（10 月 209.29）、Google 財經（52 週區間）", DG + "；" + GF),
        ("公司：FY2027 Q1 財報日期（2026-11-04 盤後）", FV),
        ("市場層：markets/network、tsmc、innolight 成品（本 repo，唯讀）", ""),
        ("待補（人工）：資料中心與通訊金額拆分；光模組營收；SiC 營收與少數股權明細；共識家數與區間；前瞻 P/E 歷史（Bloomberg）；回購", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 90; SO.column_dimensions["B"].width = 140

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("Coherent（COHR，NYSE）收支與估值模型 v0.1（AI 半導體 Project：網通次產業內複製——光模組，樣板中際旭創）　2026-10-25", True),
 ("版本：v0.1（初版）。Coherent 的有效份額口徑與中際旭創模型一致（資料中心營收 ÷ 可插拔市場），Units 頁檢查兩者相加低於 100%。", False),
 ("模型命題：Coherent 的資料中心營收＝AI 可插拔光模組市場（本專案網通模型）× Coherent 有效份額（光模組＋自製 EML／VCSEL／CW 雷射）× 單價係數＋CPO 光學 × CPO 份額；通訊、工業另依成長率推。股價自 52 週高點 440 跌到約 209（−52%），約 21 倍共識未來 12 月 EPS：市場在問磷化銦產能倍增換來的成長能否撐到 FY2028。", False),
 ("核心問題：FY2027 訂單已近滿載，股價卻腰斬——市場隱含的 FY2028 EPS 要多低，份額（新易盛、中際旭創）、單價、AI 市場三條路徑各要不利到什麼程度？", False), ("", False),
 ("會計年度：FY(t)＝t−1 年 7 月～t 年 6 月（6 月底結束）；FY2026＝2025-07～2026-06（實績）、FY2027＝2026-07～2027-06。日曆年市場以「上半年比重 hw」換算：FY(t)＝（1 − hw）× CY(t−1)＋hw × CY(t)，FY2027 的 CY2026 下半年用 Q1 指引＋Q2 假設。口徑：non-GAAP；淨利＝歸屬 Coherent。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("資料中心：markets/network 的 AI 可插拔光模組市場 × 情境顆數係數 × Coherent 有效份額（CY2026 校準）× 份額相對值（Assumed）× 單價係數（基準 1）＋ CPO 光學市場 × CPO 份額（Assumed）。", False),
 ("通訊（電信、DCI、ZR、泵浦雷射）：FY2026 推導值 × 年增（Assumed）。工業（雷射、工程材料、SiC）：FY2026 Q4 × 4（已排除出售的 A&D 與慕尼黑工具）× 成長（Assumed）。毛利率：FY2027 起為輸入路徑。營業費用：彈性 phi（Derived）。稅率 19%（公司指引）。", False),
 ("校準：FY2026 全年實績；資料中心與通訊的金額公司未揭露，由 Q3、Q4 合計與兩者季增率聯立推導（Q 頁）。不用共識（只在 Recon 對帳）。", False),
 ("估值：目標價（Val 第 0 節）＝情境權重 × 規則值隱含價；現價 ÷ FY2027、FY2028 EPS；市場隱含 FY2028 EPS（Val 第 4 節）；十年 GAAP P/E 規則值（兩個資料商，虧損年不列入）× FY2027 EPS。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。", False),
 ("", False),
 ("已知限制：(1) 份額相對值、CPO 份額、毛利率路徑皆為 Assumed，是 FY2028 EPS 最敏感的輸入（Sens）；(2) 資料中心金額為推導值，FY2027 Q1 季報（2026-11-04）後更新；(3) 有效份額含賣給其他模組廠的雷射元件，與中際旭創份額有部分重疊；(4) 十年 P/E 為 GAAP 倍數、有效年數少，套在 non-GAAP EPS 偏高估；(5) SiC 單獨營收未揭露，包含在工業；(6) R4 產業部署路徑未建（Recon 預留列）。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(sys.argv[1])
if os.environ.get("COHR_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK_ROW, "SENSN": SENS_N, "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE,
               "VAL": VAL, "IMPL": {str(k): v for k, v in IMPL.items()}, "READ": READ_ROW, "RCR": RCR, "RC_LAST": RC_LAST, "INROW": INROW, "RULE": rule_rows},
              open(os.environ["COHR_ROWMAP"], "w"), ensure_ascii=False, indent=1)

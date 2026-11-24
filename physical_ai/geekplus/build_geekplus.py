# 極智嘉 Geek+（2590.HK）個股模型 v0.1 — AI實體應用 Project（線 5 倉儲物流自動化：AMR 貨到人）
# 用法：python3 physical_ai/geekplus/build_geekplus.py   → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、segments.json、sens.json
# 會計年度＝曆年（12 月決算，IFRS）。金額：百萬人民幣（RMB m）；股數：百萬股；EPS：人民幣元；股價與目標價：港元（HK$）。
# 結構：README、In、Hist、Q（半年度產品線，含分部對帳）、PL（三情境：命題模組→產品線→標準損益表→現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 線層級變數讀 physical_ai/lines/warehouse/dist/20261116_倉儲物流單位經濟_v0.1.xlsx「下游取數」頁（建置時讀值，寫明儲存格）。
# 不讀 Siliconomics／Tokenomics／Wattonomics（實體 AI 線）。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1.1"
DATE = "20261123"
NAME = f"{DATE}_極智嘉個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032
HISTY = (2023, 2024, 2025)

# ---------------- 線模型（下游取數） ----------------
LINE_REL = "physical_ai/lines/warehouse/dist/20261116_倉儲物流單位經濟_v0.1.xlsx"
_lw = openpyxl.load_workbook(os.path.join(ROOT, LINE_REL), data_only=True)
_ld = _lw["下游取數"]
LINE = {}
for _r in range(4, _ld.max_row + 1):
    k = _ld.cell(_r, 1).value
    if k: LINE[k] = dict(row=_r, lo=_ld.cell(_r, 3).value, b=_ld.cell(_r, 4).value, hi=_ld.cell(_r, 5).value, name=_ld.cell(_r, 2).value)
MK25 = _lw["情境"]["B6"].value          # 行動機器人營收 2025（十億美元；線模型 情境!B6）
LFX = _lw["輸入"]["E44"].value           # 線模型份額換算用匯率（Assumed 7.15）
def lref(k, col="D"): return f"{LINE_REL} 下游取數!{col}{LINE[k]['row']}（{k}）"

# ---------------- 出處 ----------------
IR26 = "公司 2026 年中期業績公告（港交所 2590，2026-08-28）https://www1.hkexnews.hk/listedco/listconews/sehk/2026/0828/2026082803684.pdf"
AR25 = "公司 2025 年報（港交所 2590，2026-04-23）https://www.hkexnews.hk/listedco/listconews/sehk/2026/0423/2026042303218.pdf"
SA = "stockanalysis.com（資料商；讀取 2026-11-23）"
SAIS = "stockanalysis.com HKG:2590 income statement（資料商彙整；讀取 2026-11-23）https://stockanalysis.com/quote/hkg/2590/financials/income-statement/"
SAFC = "stockanalysis.com HKG:2590 forecast（目標價 2026-09-05 更新、9 位分析師；讀取 2026-11-23）https://stockanalysis.com/quote/hkg/2590/forecast/"
INV = "hk.investing.com 2590 頁（讀取 2026-11-23）https://hk.investing.com/equities/beijing-geekplus-tech"
SINA = "新浪財經 2026-08-29 中期業績報導（具身智能研發 4,450 萬元）https://news.sina.cn/sx/2026-08-29/detail-inipymqr8737793.d.html"
YCL = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
CNRF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與中際旭創 v0.4 同一取值，跨模型一致；逾期：與股價基準日相差逾 2 個交易日）"
DAMO = "Damodaran Country Risk Premium（2026-01-05）：China 0.91%（與中際旭創 v0.4 同值）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整（立場：中立彙整，來源為賣方或交易所資料）"
SELL = "賣方分析師（立場：券商研究，與承銷、交易業務有關）"

# ---------------- In 列 ----------------
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. FY2026 錨點（2026 上半年實績 ＋ 下半年假設；公司未給量化指引）")
for s, v, lo, hi, note in (("A", 0.31, None, None, "樂觀：資料商共識 FY2026 營收 42 億（+31%，9 位分析師；財報前，只作歷史對照）"),
                           ("B", 0.25, 0.20, 0.31, "基準：上半年營收 +25.3%、新簽訂單 +35.5%；下半年維持上半年增速"),
                           ("C", 0.20, None, None, "保守：下半年放緩到約 +17%（上半年 +25.3%）")):
    inp(f"g26_{s}", f"情境 {s}：FY2026 營收年增", "%", lo, v, hi, "Assumed", note, IR26 + "：上半年營收 1,283.9（+25.3%）；新簽訂單 2,384.8（+35.5%）；公司未給全年量化指引")
inp("gimt26", "FY2026 工業搬運營收年增", "%", 0.10, 0.25, 0.40, "Assumed", IR26 + "：上半年工業搬運 82.2（+34.4%）", "FY2025 −34.6%（" + AR25 + "）；波動大")
inp("oi26", "FY2026 其他收入淨額（含匯兌、利息、補助、財務成本、聯營）", "RMB m", -60, -15, 20, "Assumed",
    IR26 + "：上半年 其他收入及虧損淨額 −55.0（匯兌損失 −104.6、利息收入 42.2）、財務成本 −2.7、聯營 5.2 → −52.5", "下半年假設 +37.5（利息收入約 35、聯營、無匯兌）")
inp("sh_26", "FY2026 平均股數", "百萬股", None, "=(1337.287-2.62+sh_v)/2", None, "Derived", IR26 + "：已發行 1,337,287 千股、庫存股 2,620 千股（6-30）", "與評價日股數平均（7 月再買回 16.2 百萬股）")

sec("B. 命題模組：倉儲履約營收＝行動機器人市場（線 5 下游取數）× 極智嘉份額（相對變化）")
MKY = (2026, 2028, 2030, 2032)
for y in MKY:
    k = f"WH_AMR_MKT_{y}"
    for s, col, lab in (("A", "E", "高"), ("B", "D", "基準"), ("C", "C", "低")):
        v = LINE[k][{"A": "hi", "B": "b", "C": "lo"}[s]]
        lo_, hi_ = (LINE[k]["lo"], LINE[k]["hi"]) if s == "B" else (None, None)
        inp(f"mk{y % 100}_{s}", f"情境 {s}：行動機器人營收 {y}（線 5 {lab}）", "十億 USD", lo_, v, hi_, "Derived",
            lref(k, col) + "：Interact 口徑（製造＋物流物料搬運，含 AGV、自主堆高機）", "敏感度低／高＝線模型低／高（四年同動，一案）" if s == "B" else "情境 A＝線模型高、C＝線模型低")
inp("mk25", "行動機器人營收 2025（線 5）", "十億 USD", None, MK25, None, "Derived", f"{LINE_REL} 情境!B6（2024 × (1＋2025 成長)）", "只用於份額對照")
inp("lfx", "線模型份額換算匯率（線 5 輸入!E44）", "RMB/USD", None, LFX, None, "Assumed", f"{LINE_REL} 輸入!E44（線模型自標 Assumed，本輪未查證匯率）", "只作份額對照；本模型營收以人民幣、以市場相對成長推導，不用匯率")
for s, v, lo, hi in (("A", 0.05, None, None), ("B", 0.0, -0.04, 0.04), ("C", -0.06, None, None)):
    inp(f"shr_{s}", f"情境 {s}：倉儲履約營收相對行動機器人市場的年變化（份額＋價格）", "%/年", lo, v, hi, "Assumed",
        "FY2025 極智嘉營收約占行動機器人營收 7.2%（" + lref("WH_GEEKPLUS_SHARE") + "；口徑不同，量級對照）", "A：訂閱與海外擴張（上半年訂閱訂單 +75%）；C：中國廠商價格競爭（線 5：年降價基準 5%、區間 3–8%）")
for s, v, lo, hi in (("A", 0.20, None, None), ("B", 0.12, 0.05, 0.20), ("C", 0.03, None, None)):
    inp(f"gimt_{s}", f"情境 {s}：工業搬運營收年增 FY2027 起", "%/年", lo, v, hi, "Assumed", "FY2024 226.1 → FY2025 147.9（−34.6%）→ 上半年 +34.4%（" + AR25 + "、" + IR26 + "）", "線模型未涵蓋工廠內搬運，無線變數可讀")
inp("oth_rev", "其他營收（FY2026 起持平）", "RMB m", None, 2.5, None, "Assumed", "FY2025 2.5、FY2024 6.7、上半年 0.1（" + AR25 + "）", "金額極小")
inp("us_sh", "美國市場占 FY2025 營收", "%", None, "=839.187/3171.013", None, "Derived", AR25 + "：美國 839.2 ÷ 合併 3,171.0", "尾端 Sens 用（第 10 節第 19 條）")
inp("us_cut", "尾端：FY2027 起美國營收流失比例（關稅或進口限制）", "%", None, 0.0, 1.0, "Assumed", "基準 0（事件型尾端風險不進基準）；高＝美國營收全數流失", "只動營收，不同步調整毛利率（美洲毛利率未取得；海外毛利率 46.6% 高於整體，故本案為下界的反面＝損失偏小）")
inp("cust_sh", "最大客戶占 FY2025 營收", "%", None, 0.106, None, "Verified", AR25 + "：最大客戶約 10.6%（2024 年 15.5%）；前五大 42.9%", "")
inp("cust_cut", "尾端：FY2027 起最大客戶流失比例", "%", None, 0.0, 1.0, "Assumed", "基準 0；高＝最大客戶全數流失", "")

sec("C. 毛利與費用（比率以最近實績為起點）")
inp("gm0", "毛利率起點（FY2026；最近四季 TTM）", "%", 0.33, "=(1125.037-359.914+459.893)/(3171.013-1024.722+1283.899)", 0.38, "Derived",
    "TTM＝(FY2025 毛利 − 上半年 2025 ＋ 上半年 2026) ÷ 同口徑營收＝35.7%（" + AR25 + "、" + IR26 + "）", "FY2025 35.5%、FY2024 34.8%、上半年 35.8%；海外毛利率 46.2–46.6%、中國較低；低／高 Assumed")
for s, v, lo, hi in (("A", 0.005, None, None), ("B", 0.0, -0.005, 0.005), ("C", -0.005, None, None)):
    inp(f"gmd_{s}", f"情境 {s}：FY2027 起毛利率每年變化", "pt/年", lo, v, hi, "Assumed", "A：訂閱與軟體占比提高；C：中國 AMR 價格競爭", "基準持平（原則第 10 節第 35 條：無收斂證據）")
inp("e_ox", "營業費用（研發、銷售、行政）對營收成長的彈性", "x", 0.45, "=((335.076+538.525+270.642)/(282.048+445.637+220.289)-1)/(3171.013/2409.011-1)", 0.85, "Derived",
    "FY2024→FY2025：營業費用 +20.7% ÷ 營收 +31.6%＝0.65（" + AR25 + "）", "低＝TTM 對 FY2025（費用 +3.7% ÷ 營收 +8.2%＝0.45，Derived）；高 0.85 Assumed（具身智能投入加大）")
inp("emb_rd", "具身智能額外研發（FY2027 起每年，另加於研發）", "RMB m", None, 0.0, 150.0, "Assumed", SINA + "：上半年具身智能研發 44.5（已在費用基期內）", "基準 0＝不另加（已含在費用基期）；高＝每年另加 150")
inp("imp_r", "應收款減值 ÷ 營收", "%", 0.006, 0.010, 0.014, "Derived", "FY2025 35.4 ÷ 3,171.0＝1.1%；FY2024 0.6%；上半年 1.2%（" + AR25 + "、" + IR26 + "）", "")
inp("oi_lt", "FY2027 起其他收入淨額（利息收入、補助、聯營、財務成本；不預測匯兌）", "RMB m", 0.0, 30.0, 60.0, "Assumed", "上半年利息收入 42.2、財務成本 −2.7、聯營 5.2；現金 2,200 減少中", "匯兌損益不預測（上半年 −104.6、FY2025 −20.5）")
inp("tax_w", "所得稅率（只對正的稅前利益；營業利益同口徑）", "%", 0.10, 0.15, 0.25, "Analogy", "中國高新技術企業 15%（類比法定優惠稅率）；累計虧損可抵減未建模", "FY2025 所得稅 5.1（稅前 −5.3；海外子公司）")
inp("pay_r", "配息率", "%", None, 0.0, None, "Interested-party", AR25 + "：FY2025、FY2024 均未派末期股息", "只用於收支關卡", CO)

sec("D. 估值輸入（評價日 2026-12-31；資產負債表 2026-06-30）")
inp("px", "股價（2026-11-20 收盤，港元）", "HK$", None, 5.405, None, "Verified", YCL + "：2590.HK 2026-11-20 收盤 5.405", "11-11 收盤 7.22 → 11-20 5.405（−25.1%，同檔）；52 週區間 5.17–33.50（" + INV + "）；跌幅主因未找到")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("fx_hk", "美元兌港元（2026-11-20）", "HKD/USD", None, 7.8426, None, "Verified", "Yahoo Finance HKD=X 2026-11-20（physical_ai/data/closes2_20261109_20261120.csv）", "")
inp("fx_cn", "美元兌人民幣（2026-11-20）", "CNY/USD", None, 6.6665, None, "Verified", "Yahoo Finance CNY=X 2026-11-20（同檔）", "")
inp("hkcny", "港元 ÷ 人民幣（每 1 人民幣兌港元）", "HKD/CNY", None, "=fx_hk/fx_cn", None, "Derived", "＝HKD=X ÷ CNY=X", "目標價＝人民幣每股價值 × 本列")
inp("sh_v", "評價日股數（已發行 − 庫存股 − 7 月買回）", "百萬股", 1289.0, "=1337.287-2.62-16.2166", None, "Derived",
    IR26 + "：已發行 1,337,287 千股、庫存股 2,620 千股（6-30）、7 月再買回 16,216,600 股", "低＝資料商流通股 1,289 百萬股（" + INV + "；Interested-party，未核對 8–11 月逐筆買回）；RSU／選擇權未取得")
inp("bb_post", "6-30 後買回付出現金（7 月 16.2 百萬股）", "RMB m", 0.0, "=16.2166*10/hkcny", None, "Assumed", "均價未取得：假設每股 10 港元（7 月股價約 10–11.7，" + "longbridge 2026-07-09 報導 11.20–11.67 買回、收 10.92）", "低＝0")
inp("cash_v", "現金及約當現金＋以公允價值計量金融資產（2026-06-30）", "RMB m", None, "=2199.917+1.224", None, "Verified", IR26, "定期存款 0（2025 年底 180.6）")
inp("debt_v", "銀行借款（2026-06-30，全部流動）", "RMB m", None, 65.4, None, "Verified", IR26, "")
inp("lease_v", "租賃負債（2026-06-30；流動 16.39＋非流動 16.88）", "RMB m", None, "=16.39+16.876", None, "Verified", IR26, "")
inp("h1fcf", "2026 上半年實績自由現金流（營業現金流 − 資本支出）", "RMB m", None, "=-285.167-14.939", None, "Verified", IR26 + "：營業活動淨現金 −285.2、購置資產 −14.9", "Stub（下半年）FCFF＝FY2026 模型 FCFF − 本列")
inp("rf", "無風險利率（人民幣 10 年期國債）", "%", None, 0.0161, None, "Verified", CNRF, "")
inp("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Interested-party", DAMO, "", "Damodaran（學者整理；評等機構違約利差推得）")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("beta", "β（同業 5 家 5 年 β 中位數，Blume 調整）", "x", "=E{r}-0.3", "=0.67*MEDIAN(2.05,1.03,1.98,2.34,2.11)+0.33", "=E{r}+0.3", "Analogy",
    SA + "：KION 2.05、大福 1.03、AutoStore 1.98、Ocado 2.34、Symbotic 2.11（5 年月報酬）", "本公司上市未滿 2 年，個股 β 未取得；第二來源未取得；敏感度 ±0.3（Assumed）")
inp("kd", "稅前負債成本", "%", None, "=rf+0.01", None, "Assumed", "rf ＋ 1pt（借款 65.4，占資本很小）", "")
inp("g_cap", "永續成長率上限", "%", 0.01, 0.03, 0.04, "Assumed", "DCF 標準規格：3%，且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("fc_da", "折舊攤銷 ÷ 營收", "%", 0.008, "=(17.347+19.485+3.566)/3171.013", 0.02, "Derived", AR25 + "：FY2025 折舊 17.3＋使用權資產 19.5＋攤銷 3.6 ＝ 40.4 ÷ 3,171.0", "FY2024 62.5 ÷ 2,409.0＝2.6%")
inp("fc_cx", "資本支出 ÷ 營收", "%", 0.012, 0.025, 0.041, "Assumed", "FY2025 130.1 ÷ 3,171.0＝4.1%（含新廠）；上半年 14.9 ÷ 1,283.9＝1.2%（" + AR25 + "、" + IR26 + "）", "取兩者之間；輕資產（委外製造為主未核對）")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.07, "=((958.208+803.151-1137.519-392.904)-(713.556+1029.457-999.76-610.674))/(3171.013-2409.011)", 0.25, "Derived",
    "（應收＋存貨 − 應付 − 合約負債）的變動 ÷ 營收增量：FY2024→FY2025 (230.9 − 132.6) ÷ 762.0＝12.9%（" + AR25 + "）", "低 7%（FY2024 營運資金 ÷ 營收 5.5%）、高 25%（上半年季節性墊高到 470.2）；Assumed")
inp("dps_tgt", "評價日到目標時點之每股股利", "RMB", None, 0.0, None, "Assumed", "未派息", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable：EV ÷ 營收）", "x", 0.5, 0.7, 0.9, "Assumed", "個股模型原則第 4 節預設各半；v0.1.1 依獨立核對調降 Comparable 權重：Comparable 的上檔幾乎全部來自「倍數向同業中位數收斂」的假設（Analogy），證據不足以與 DCF 等權", "低＝原則預設各半；高＝幾乎只看 DCF")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("E. Comparable：同業 EV ÷ 近四季營收（以 2026-11-20 收盤重算市值；EV 其餘項目、股數與營收取資料商頁面）")
# (鍵, 名稱, 股數(百萬或十億，與市值同單位), 11-20 收盤, 淨負債（正＝負債）, TTM 營收, 群組, 頁面日期, 毛利率)
PEERS = [("ev_sym", "Symbotic SYM（美國倉儲系統）", "591.52*31.1/1000", 1.21 * -1, 2.25, "系統", "頁面日期未標示（TTM 為 FY2025，舊）", "未取得"),
         ("ev_kgx", "KION KGX（Dematic 倉儲系統、堆高機）", "131.12*29.2/1000", 7.33, 11.49, "系統", "2026-09-18", "26.3%"),
         ("ev_dfk", "大福 6383（物流自動化系統）", "372.21*4129/1000000", -0.22395, 0.68975, "系統", "2026-09-11", "25.3%"),
         ("ev_ocdo", "Ocado OCDO（電商倉儲技術）", "832.86*3.102/1000", 0.9695, 1.74, "系統", "2026-10-23", "23.0%"),
         ("ev_auto", "AutoStore AUTO（立方格 AS/RS，授權式）", "3.36*14.79", 0.90144, 6.70, "標準化產品", "2026-10-23", "未取得（約 7 成，線 5 追蹤指標 71.3%）"),
         ("ev_dobot", "越疆 2432（港股協作機器人）", "439.96*14.4/1000", -2.31, 0.75669, "標準化產品", "2026-09-21", "46.5%")]
for k, lab, mc, nd, rv, grp, d, gm in PEERS:
    inp(k, f"EV ÷ 近四季營收：{lab}", "x", None, f"=({mc}+({nd}))/{rv}", None, "Derived",
        f"{SA}（{d} 頁面）：股數、淨負債、近四季營收；× 2026-11-20 收盤（{YCL}）", f"群組 {grp}；毛利率 {gm}；會計口徑：IFRS（KION、Ocado、AutoStore、越疆）、US GAAP（Symbotic）、日本基準（大福）；EV 含租賃與否依資料商口徑，未逐家核對")
inp("ev_med", "同業 EV ÷ 近四季營收中位數（6 家，進 Comparable）", "x", "=MEDIAN(ev_sym,ev_kgx,ev_dfk,ev_ocdo)", "=MEDIAN(ev_sym,ev_kgx,ev_dfk,ev_ocdo,ev_auto,ev_dobot)", "=MEDIAN(ev_auto,ev_dobot)", "Derived",
    "少於 8 家，不做迴歸（第 10 節第 7 條）", "敏感度低＝系統群組 4 家中位數；高＝標準化產品群組 2 家中位數（第 18 條：依商業模式並列子群中位數）")
inp("ev_self", "極智嘉現行 EV ÷ 近四季營收（2026-11-20；只作對照，不進模型）", "x", None,
    "=(px*sh_v/hkcny-(cash_v-debt_v-lease_v-bb_post))/(3171.013-1024.722+1283.899)", None, "Derived", "市值（現價 × 評價日股數 ÷ 匯率）− 淨現金，÷ TTM 營收 3,430.2", "")
inp("rel_k", "相對倍數調整係數", "x", 0.31, 1.0, None, "Assumed", "1.00（未調整：缺極智嘉對同業的歷史相對倍數；上市未滿 2 年）", "低＝現行相對倍數（極智嘉 ÷ 同業中位數，約 0.31；建置時寫入常數）：市場維持現行折價")
inp("cons_tp", "賣方平均目標價（只作對照，不進模型）", "HK$", None, 21.92, None, "Interested-party", SAFC + "：平均 21.92、中位數 20.31、區間 12.56–38.26（10 位）", "財報前（2026-09-05 更新），只作歷史對照", SELL)
inp("cons_rev26", "共識 FY2026 營收（只作對照，不進模型）", "RMB m", None, 4200.0, None, "Interested-party", SAFC + "：平均 4.2B（區間 4.0–4.2B，9 位）", "日期未標示", SELL)
inp("cons_eps26", "共識 FY2026 EPS（只作對照，不進模型）", "RMB", None, 0.14, None, "Interested-party", SAFC + "：平均 0.14（0.06–0.26）", "日期未標示；上半年 EPS −0.13", SELL)


def build(path, ov=None, sens=None, rev=None):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額百萬人民幣、股數百萬股、EPS 人民幣元、股價港元"
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
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2); IN.cell(r, 10, inc)
        dn(k, f"In!$E${r}"); KR[k] = r; r += 1
    LAST_IN = r - 1
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("g26", "FY2026 營收年增"), ("mk26", "行動機器人營收 2026（線 5）"), ("mk28", "行動機器人營收 2028（線 5）"), ("mk30", "行動機器人營收 2030（線 5）"),
           ("mk32", "行動機器人營收 2032（線 5）"), ("shr", "倉儲履約相對市場年變化"), ("gimt", "工業搬運年增"), ("gmd", "毛利率每年變化")]
    for f, lab in FAM:
        IN.cell(r, 1, f + "_scn"); IN.cell(r, 2, lab)
        for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR[f + '_' + s]}")
        IN.cell(r, 8, f"引用 In 頁 {f}_A／{f}_B／{f}_C；{lab}"); r += 1
    IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重")
    for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR['w_' + s]}")
    IN.cell(r, 8, "引用 In 頁 w_A／w_B／w_C；合計＝1"); r += 1
    IN.column_dimensions["B"].width = 56; IN.column_dimensions["H"].width = 70

    # ---------------- Hist ----------------
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（年度；IFRS；百萬人民幣）。公式只引用本頁儲存格（定義名稱 h_*）。FY2023 取資料商（招股書原文未逐項核對）"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "標記", "來源"])
    HR = [("h_rev", "營收", 2143.0, 2409.011, 3171.013), ("h_wf", "倉儲履約營收", None, 2176.174, 3020.670), ("h_imt", "工業搬運營收", None, 226.140, 147.880),
          ("h_oth", "其他營收", None, 6.697, 2.463), ("h_cogs", "營業成本", "=h_rev_23-h_gp_23", 1571.844, 2045.976), ("h_gp", "毛利", 659.27, 837.167, 1125.037),
          ("h_rd", "研發費用", 379.87, 282.048, 335.076), ("h_sm", "銷售及營銷費用", None, 445.637, 538.525), ("h_ga", "行政費用", None, 220.289, 270.642),
          ("h_sga", "銷售＋行政（FY2023 只有合計）", 772.76, "=h_sm_24+h_ga_24", "=h_sm_25+h_ga_25"), ("h_imp", "應收款減值", None, 15.173, 35.427),
          ("h_oi", "其他收入及虧損淨額（公司列在經營虧損內）", None, -1.595, 25.855), ("h_fin", "財務成本", None, -13.923, -12.730),
          ("h_red", "贖回負債帳面值變動（上市前優先股）", None, -685.807, 21.163), ("h_aso", "應佔聯營公司利潤", None, 0.024, 15.003),
          ("h_pbt", "稅前利益", -1123.0, -827.281, -5.342), ("h_tax", "所得稅", 4.16, 4.220, 5.065), ("h_ni", "年度利潤（歸屬母公司）", -1127.0, -831.501, -10.407),
          ("h_eps", "每股盈餘（基本＝稀釋，人民幣元；公司公告值）", -0.97, -0.72, 0.03), ("h_sh", "加權平均股數（百萬股）", 292.0, 292.289, 792.036),
          ("h_adj", "經調整淨利（非 IFRS；公司口徑）", None, -92.237, 43.822), ("h_da", "折舊攤銷（含使用權資產）", 43.06, "=26.635+33.854+1.992", "=17.347+19.485+3.566"),
          ("h_cx", "購置物業廠房及無形資產", None, 102.234, 130.092), ("h_ocf", "營業活動淨現金", None, -108.101, 85.663), ("h_sbc", "股份支付費用", None, 29.494, 42.107),
          ("h_rec", "應收帳款及票據", None, 713.556, 958.208), ("h_inv", "存貨", None, 1029.457, 803.151), ("h_ap", "應付帳款", None, 999.760, 1137.519),
          ("h_cl", "合約負債", None, 610.674, 392.904)]
    for i, (k, lab, a, b, c) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b, c)):
            if v is None: H.cell(i, 2 + j, "未取得"); continue
            H.cell(i, 2 + j, v).font = BLUE
        H.cell(i, 5, "Interested-party（FY2023）／Verified" if a is not None else "Verified")
        H.cell(i, 6, (SAIS + "；" if a is not None else "") + AR25)
        for j, y in enumerate(("23", "24", "25")): dn(f"{k}_{y}", f"Hist!${'BCD'[j]}${i}")
    i = 3 + len(HR)
    H.cell(i, 1, "營運資金（應收＋存貨 − 應付 − 合約負債；Derived）"); H.cell(i, 3, "=h_rec_24+h_inv_24-h_ap_24-h_cl_24"); H.cell(i, 4, "=h_rec_25+h_inv_25-h_ap_25-h_cl_25")
    H.cell(i + 1, 1, "檢查：產品線合計 − 營收（應為 0）"); H.cell(i + 1, 3, "=ROUND(h_wf_24+h_imt_24+h_oth_24-h_rev_24,3)"); H.cell(i + 1, 4, "=ROUND(h_wf_25+h_imt_25+h_oth_25-h_rev_25,3)")
    H.cell(i + 2, 1, "FY2023 產品線拆分、銷售與行政分列、減值與業外明細未取得（資料商只給合計）；FY2023 營收與稅前、淨利為資料商取整到百萬")
    H.cell(i + 3, 1, "歷史 EPS 用公司公告值（FY2025 0.03 為正、年度利潤 −10.4：差異來自上市前贖回負債調整歸屬，資料商列「優先股股利及其他調整」；FY2023–FY2024 的公告 EPS 與年度利潤 ÷ 加權股數不一致，分母口徑未取得）；預測期 EPS＝稅後淨利 ÷ 評價日股數")
    H.column_dimensions["A"].width = 46

    # ---------------- Q（半年度） ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 半年度產品線營收與損益（百萬人民幣）。H1 為公告原文（Verified）；H2'25＝FY2025 − H1'25（Derived）。公司只有一個經營分部，以產品線營收為分部列"
    hdr(Q, 2, ["項目", "H1'25", "H2'25", "H1'26", "標記", "來源", "來源誘因"])
    QV = [("倉儲履約營收", 962.455, "=h_wf_25-B3", 1201.553), ("工業搬運營收", 61.195, "=h_imt_25-B4", 82.221), ("其他營收", 1.072, "=h_oth_25-B5", 0.125),
          ("營收", 1024.722, "=h_rev_25-B6", 1283.899), ("檢查：產品線合計 − 營收（應為 0）", "=ROUND(SUM(B3:B5)-B6,3)", "=ROUND(SUM(C3:C5)-C6,3)", "=ROUND(SUM(D3:D5)-D6,3)"),
          ("毛利", 359.914, "=h_gp_25-B8", 459.893), ("毛利率", "=B8/B6", "=C8/C6", "=D8/D6"), ("研發費用", 147.154, "=h_rd_25-B10", 187.869),
          ("銷售及營銷費用", 240.216, "=h_sm_25-B11", 277.274), ("行政費用", 134.754, "=h_ga_25-B12", 99.394), ("應收款減值", 5.713, "=h_imp_25-B13", 16.037),
          ("營業利益（本模型口徑：毛利 − 營業費用）", "=B8-SUM(B10:B13)", "=C8-SUM(C10:C13)", "=D8-SUM(D10:D13)"),
          ("其他收入淨額（含財務成本、贖回負債、聯營）", "=B16-B14", "=C16-C14", "=D16-D14"), ("稅前利益", -46.577, "=h_pbt_25-B16", -173.168),
          ("期間利潤", -47.956, "=h_ni_25-B17", -176.534), ("經調整淨利（非 IFRS）", -89.279, "=h_adj_25-B18", -60.611), ("新簽訂單", None, None, 2384.8)]
    for i, (lab, a, b, c) in enumerate(QV, 3):
        Q.cell(i, 1, lab)
        for j, v in enumerate((a, b, c)):
            if v is None: Q.cell(i, 2 + j, "未取得"); continue
            cc = Q.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): cc.font = BLUE
        Q.cell(i, 5, "Verified" if isinstance(a, (int, float)) or isinstance(c, (int, float)) and not isinstance(a, str) else "Derived")
        Q.cell(i, 6, IR26 + "（含前年同期欄）；H2'25＝FY2025（" + AR25 + "）− H1'25")
    for c in "BCD": Q[f"{c}9"].number_format = "0.0%"
    Q.cell(3 + len(QV) + 1, 1, "註：上半年新簽訂單 2,384.8（+35.5%）、訂閱訂單 156（+75%）、具身智能研發 44.5；營收 75% 以上來自中國以外（" + IR26 + "、" + SINA + "）")
    Q.column_dimensions["A"].width = 46

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（行動機器人市場 × 份額）→ 產品線營收 → 標準損益表（IFRS 口徑）→ 現金流（百萬人民幣；FY2023–FY2025 實績、FY2026 起模型）"
    PR = {}
    rr = 3
    KEYS = ["mkt", "shp", "core", "tail", "wf", "imt", "oth", "rev", "chk", "gr", "cogs", "gp", "gm", "rd", "sm", "ga", "imp", "ox", "op", "opm", "oi", "pbt", "tax", "ni", "nm", "nig",
            "sh", "eps", "da", "cx", "dwc", "fcff", "div", "fcfad"]
    LABS = ["命題模組：行動機器人營收（十億美元；線 5 下游取數）", "極智嘉隱含份額（倉儲履約 ÷ 市場；線 5 匯率換算，只作對照）", "倉儲履約營收（尾端調整前）",
            "尾端調整係數（美國、最大客戶；基準 1）", "倉儲履約（Warehouse fulfilment）營收", "工業搬運（Industrial material transport）營收", "其他營收", "營收（合併）",
            "檢查：分部合計 − 營收（應為 0）", "營收年增率", "營業成本", "毛利", "毛利率", "研發費用", "銷售及營銷費用", "行政費用", "應收款減值",
            "營業費用（研發＋銷售＋行政＋減值）", "營業利益（本模型口徑：毛利 − 營業費用）", "營業利益率", "其他收入淨額（業外：利息、匯兌、補助、財務成本、聯營、贖回負債）",
            "稅前利益", "所得稅", "稅後淨利", "淨利率", "淨利年增率", "平均股數（百萬股）", "EPS（基本＝稀釋，人民幣元）",
            "折舊攤銷", "資本支出", "營運資金增加", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "股利", "自由現金流扣股利（收支關卡用）"]
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" + ("" if y in HISTY else "E") for y in YRS]); b = rr + 2
        R = {k: b + i for i, k in enumerate(KEYS)}
        R["hdr"] = rr + 1
        for i, l in enumerate(LABS): P.cell(b + i, 1, l)
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1)
            cell = lambda k: f"{c}{R[k]}"; prev = lambda k: f"{p}{R[k]}"
            def w(k, f): P[cell(k)] = f
            yy = str(y)[2:]
            if y in HISTY:
                if y >= 2024:
                    w("wf", f"=h_wf_{yy}"); w("imt", f"=h_imt_{yy}"); w("oth", f"=h_oth_{yy}")
                    w("chk", f"=ROUND({cell('wf')}+{cell('imt')}+{cell('oth')}-{cell('rev')},6)")
                    w("sm", f"=h_sm_{yy}"); w("ga", f"=h_ga_{yy}"); w("imp", f"=h_imp_{yy}")
                    w("ox", f"={cell('rd')}+{cell('sm')}+{cell('ga')}+{cell('imp')}")
                    w("cx", f"=h_cx_{yy}")
                else:
                    for k in ("wf", "imt", "oth", "sm", "ga", "imp", "cx"): w(k, "未取得")
                    w("ox", f"={cell('rd')}+h_sga_{yy}")
                w("rev", f"=h_rev_{yy}"); w("cogs", f"=h_cogs_{yy}"); w("gp", f"=h_gp_{yy}"); w("rd", f"=h_rd_{yy}")
                w("op", f"={cell('gp')}-{cell('ox')}"); w("pbt", f"=h_pbt_{yy}"); w("oi", f"={cell('pbt')}-{cell('op')}")
                w("tax", f"=h_tax_{yy}"); w("ni", f"=h_ni_{yy}"); w("sh", f"=h_sh_{yy}"); w("eps", f"=h_eps_{yy}"); w("da", f"=h_da_{yy}")
                if y == 2025:
                    w("mkt", "=mk25"); w("shp", f"={cell('wf')}/lfx/{cell('mkt')}/1000")
                    w("dwc", "=(h_rec_25+h_inv_25-h_ap_25-h_cl_25)-(h_rec_24+h_inv_24-h_ap_24-h_cl_24)")
                if y >= 2024:
                    w("gr", f"={cell('rev')}/{prev('rev')}-1"); w("nig", "N/M（前一年虧損）" if y >= 2024 else "")
            else:
                if y in MKY: w("mkt", f"=mk{y % 100}_{s}")
                else:
                    w("mkt", f"=SQRT(mk{(y - 1) % 100}_{s}*mk{(y + 1) % 100}_{s})")
                w("shp", f"={cell('wf')}/lfx/{cell('mkt')}/1000")
                if y == 2026:
                    w("rev", f"=h_rev_25*(1+g26_{s})"); w("imt", "=h_imt_25*(1+gimt26)"); w("oth", "=oth_rev")
                    w("core", f"={cell('rev')}-{cell('imt')}-{cell('oth')}"); w("tail", 1); w("wf", f"={cell('core')}")
                    w("gm", "=gm0"); w("oi", "=oi26"); w("sh", "=sh_26")
                    w("rd", f"={prev('rd')}*(1+e_ox*{cell('gr')})")
                else:
                    w("core", f"={prev('core')}*{cell('mkt')}/{prev('mkt')}*(1+shr_{s})")
                    w("tail", "=1-us_cut*us_sh-cust_cut*cust_sh")
                    w("wf", f"={cell('core')}*{cell('tail')}"); w("imt", f"={prev('imt')}*(1+gimt_{s})*{cell('tail')}/{prev('tail')}"); w("oth", f"={prev('oth')}")
                    w("rev", f"={cell('wf')}+{cell('imt')}+{cell('oth')}")
                    w("gm", f"={prev('gm')}+gmd_{s}"); w("oi", "=oi_lt"); w("sh", "=sh_v")
                w("chk", f"=ROUND({cell('wf')}+{cell('imt')}+{cell('oth')}-{cell('rev')},6)")
                w("gr", f"={cell('rev')}/{prev('rev')}-1")
                w("cogs", f"={cell('rev')}*(1-{cell('gm')})"); w("gp", f"={cell('rev')}-{cell('cogs')}")
                w("sm", f"={prev('sm')}*(1+e_ox*{cell('gr')})"); w("ga", f"={prev('ga')}*(1+e_ox*{cell('gr')})"); w("imp", f"=imp_r*{cell('rev')}")
                w("ox", f"={cell('rd')}+{cell('sm')}+{cell('ga')}+{cell('imp')}")
                w("op", f"={cell('gp')}-{cell('ox')}"); w("pbt", f"={cell('op')}+{cell('oi')}"); w("tax", f"=MAX(0,{cell('pbt')})*tax_w"); w("ni", f"={cell('pbt')}-{cell('tax')}")
                w("nig", f'=IF({prev("ni")}>0,{cell("ni")}/{prev("ni")}-1,"N/M（前一年虧損）")')
                w("eps", f"={cell('ni')}/{cell('sh')}")
                w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}"); w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
                w("fcff", f"={cell('op')}-MAX(0,{cell('op')})*tax_w+{cell('da')}-{cell('cx')}-{cell('dwc')}")
                w("div", f"=pay_r*MAX(0,{cell('ni')})"); w("fcfad", f"={cell('fcff')}-{cell('div')}")
            if y in HISTY:
                w("gm", f"={cell('gp')}/{cell('rev')}")
            w("opm", f"={cell('op')}/{cell('rev')}"); w("nm", f"={cell('ni')}/{cell('rev')}")
            if y == 2023:
                w("gr", "未取得（FY2022 只作資料商參考）"); w("nig", "N/M（前一年虧損）")
        # 研發公式簡化（上面 2027 起含 emb_rd）
        for y in YRS:
            if y >= 2027:
                c = COL[y]; p = COL[y - 1]
                P[f"{c}{R['rd']}"] = f"=({p}{R['rd']}-IF({y}=2027,0,emb_rd))*(1+e_ox*{c}{R['gr']})+emb_rd"
        P.cell(R["mkt"], 2, "FY2026、2028、2030、2032 讀線 5；奇數年取相鄰兩年幾何平均（線模型同為幾何路徑）")
        P.cell(R["core"], 2, "FY2026＝營收 − 工業搬運 − 其他；FY2027 起＝前一年 × 市場成長 ×(1＋shr)")
        P.cell(R["rd"], 2, "前一年 ×(1＋彈性 e_ox × 營收年增)；FY2027 起另加 emb_rd（不複利）")
        P.cell(R["op"], 2, "公司「經營虧損」＝本列＋其他收入及虧損淨額（見 Recon）")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 − MAX(0, 營業利益) × tax_w；股份支付不加回")
        for k in ("gr", "gm", "opm", "nm", "nig", "shp"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(LABS) + 2
    P.column_dimensions["A"].width = 52; P.column_dimensions["B"].width = 30
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "wf", "fcfad"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），人民幣現金流以人民幣 WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)；每股價值以現匯換成港元"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（百萬人民幣）"; D["B4"] = "=px*sh_v/hkcny"
    D["A5"] = "負債 D（借款＋租賃）"; D["B5"] = "=debt_v+lease_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g（上限 3% 與 rf 取小）"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金（2026-06-30）＝現金 − 借款 − 租賃 − 6-30 後買回"; D["B10"] = "=cash_v-debt_v-lease_v-bb_post"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND(((px*sh_v/hkcny)*(rf+beta*(erp+crp))+(debt_v+lease_v)*kd*(1-tax_w))/((px*sh_v/hkcny)+debt_v+lease_v)+wacc_adj-wacc,9)"
    D["A12"] = "檢查：淨現金 − 資產負債表重算（應為 0）"; D["B12"] = "=ROUND(dcf_nc-(2199.917+1.224-65.4-16.39-16.876-bb_post),6)"
    for rr_ in range(3, 11): D[f"B{rr_}"].number_format = "0.00%" if rr_ in (3, 6, 7, 8, 9) else "#,##0.0"
    DR = {}; rr = 14
    for s in SC:
        R = PR[s]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}E" for y in range(2027, 2033)]); b = rr + 2
        rows = ["EBIT（營業利益）", "NOPAT", "＋折舊攤銷", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF（應為 0）", "折現年數（期中）", "折現因子", "FCFF 現值"]
        for i, l in enumerate(rows): D.cell(b + i, 1, l)
        for j, y in enumerate(range(2027, 2033)):
            c = "BCDEFG"[j]; pc = COL[y]
            D[f"{c}{b}"] = f"=PL!{pc}{R['op']}"; D[f"{c}{b+1}"] = f"={c}{b}-MAX(0,{c}{b})*tax_w"
            D[f"{c}{b+2}"] = f"=PL!{pc}{R['da']}"; D[f"{c}{b+3}"] = f"=-PL!{pc}{R['cx']}"; D[f"{c}{b+4}"] = f"=-PL!{pc}{R['dwc']}"
            D[f"{c}{b+5}"] = f"=PL!{pc}{R['fcff']}"; D[f"{c}{b+6}"] = f"=ROUND(SUM({c}{b+1}:{c}{b+4})-{c}{b+5},6)"
            D[f"{c}{b+7}"] = j + 0.5; D[f"{c}{b+8}"] = f"=1/(1+wacc)^{c}{b+7}"; D[f"{c}{b+9}"] = f"={c}{b+5}*{c}{b+8}"
        t = b + 11
        items = [("明確期現值合計", f"=SUM(B{b+9}:G{b+9})"),
                 ("NOPAT FY2033＝NOPAT FY2032 ×(1＋g)", f"=G{b+1}*(1+g_tv)"),
                 ("終值（FY2032 年底）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^6"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金（2026-06-30，扣 6-30 後買回）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（2026 下半年＝FY2026 模型 FCFF − 上半年實績）", f"=PL!F{R['fcff']}-h1fcf"),
                 ("股權價值（百萬人民幣）", f"=B{t+4}+B{t+5}+B{t+6}"),
                 ("每股價值（評價日，人民幣）", f"=MAX(0,B{t+7})/sh_v"),
                 ("每股價值（評價日，港元）", f"=B{t+8}*hkcny"),
                 ("DCF 情境價值推到 12 個月後（港元；×(1＋ke)^t − 期間股利）", f"=B{t+9}*(1+ke)^t_tgt-dps_tgt*hkcny"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+3}>=0,B{t+11}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=IF(G{b+1}>0,B{t+2}/G{b+1},\"N/M\")"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（港元，推到 12 個月後）", f"=MAX(0,(B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6)+B{t+5}+B{t+6})/sh_v*hkcny*(1+ke)^t_tgt-dps_tgt*hkcny")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 10, ps=t + 9, ev=t + 4, tvs=t + 11, alt=t + 14, chk=b + 6, stub=t + 6)
        dn(f"dcf_{s}", f"DCF!$B${t+10}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 64

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價（港元）：DCF（權重見 In 頁）＋ Comparable（EV ÷ 營收；其餘權重）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}")
        V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後（港元）", [f"=dcf_{s}" for s in SC], "DCF 頁")
    cf = lambda s, k: f"PL!G{PR[s][k]}"
    comp_f = [f"=(ev_med*{cf(s,'rev')}+dcf_nc+DCF!B{DR[s]['stub']})/sh_v*hkcny" for s in SC]
    r_cu = vrow("Comparable（未調整）：同業 EV ÷ 營收中位數 × FY2027 營收 ＋ 淨現金（港元）", comp_f,
                "目標時點（2027-11）的近四季營收約等於 FY2027（1/12 FY2026＋11/12 FY2027 以 FY2027 代）；同業倍數為 2026-11-20 的近四季口徑")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"=(ev_med*rel_k*{cf(s,'rev')}+dcf_nc+DCF!B{DR[s]['stub']})/sh_v*hkcny" for s in SC], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（DCF 權重 × DCF ＋ (1 − DCF 權重) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note="", fmt=None):
        nonlocal vr
        V.cell(vr, 1, lab); c = V.cell(vr, 2, f); V.cell(vr, 6, note)
        if fmt: c.number_format = fmt
        vr += 1; return vr - 1
    r_tp = one("加權目標價（港元）", f"=E{r_tps}", "", "#,##0.00"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權，港元）", f"=E{r_dcf}", "", "#,##0.00")
    r_cmpt = one("Comparable 目標價（EV ÷ 營收，情境加權，進加權，港元）", f"=E{r_ca}", "", "#,##0.00")
    r_cmpu = one("Comparable 目標價（未調整，情境加權，港元）", f"=E{r_cu}", "", "#,##0.00")
    r_px = one("現價（In 頁 px，港元）", "=px", "", "#,##0.000")
    r_up = one("上檔空間", f"=B{r_tp}/px-1", "", "0.0%"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（百萬人民幣；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_lq = one("期間流動性：情境 B 最低年末淨現金（FY2026–FY2030，百萬人民幣；不含新增舉債）", f"=dcf_nc+DCF!B{DR['B']['stub']}+MIN(0,PL!G{PR['B']['fcfad']},PL!G{PR['B']['fcfad']}+PL!H{PR['B']['fcfad']},PL!G{PR['B']['fcfad']}+PL!H{PR['B']['fcfad']}+PL!I{PR['B']['fcfad']},PL!G{PR['B']['fcfad']}+PL!H{PR['B']['fcfad']}+PL!I{PR['B']['fcfad']}+PL!J{PR['B']['fcfad']})", "第 10 節第 55 條；無借款需求時期間最大新增舉債為 0")
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100", "", "0.0")
    r_bnd = one("翻轉目標價：升一級／降一級（港元）", f'=IF(B{r_up}>0.5,"無（已是最高級）",TEXT(px*(1+IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5)))),"0.00"))&"／"&TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,IF(B{r_up}>=-0.5,-0.5,-9))))),"0.00")', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)", "", "+0.0%;-0.0%")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)", "", "+0.0%;-0.0%")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px", "", "+0.0%;-0.0%")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt", "", "0.00")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬", "0.00")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_evs = one("現行 EV ÷ 近四季營收：極智嘉", "=ev_self", "分母：TTM 營收 3,430.2（2025 下半年＋2026 上半年）", "0.00")
    r_evm = one("同業 EV ÷ 近四季營收中位數（6 家）", "=ev_med", "", "0.00")
    r_rel = one("現行相對倍數（極智嘉 ÷ 同業中位數）", "=ev_self/ev_med", "上市未滿 2 年，無 10 年比值；調整後 Comparable 需要的相對重估倍數＝rel_k ÷ 本列", "0.00")
    r_sg1 = one("子群中位數（純展示）：系統（Symbotic、KION、大福、Ocado）", "=MEDIAN(ev_sym,ev_kgx,ev_dfk,ev_ocdo)", "第 10 節第 18 條", "0.00")
    r_sg2 = one("子群中位數（純展示）：標準化產品（AutoStore、越疆）", "=MEDIAN(ev_auto,ev_dobot)", "", "0.00")
    r_ipe = one("隱含 EV ÷ FY2027 營收（加權目標價，情境 B 營收）", f"=(B{r_tp}/hkcny*sh_v-dcf_nc)/PL!G{PR['B']['rev']}", "分母：模型情境 B FY2027 營收", "0.00")
    r_pe = one("隱含 P/E：加權目標價 ÷ 情境 B FY2028 EPS（港元換算）", f'=IF(PL!H{PR["B"]["eps"]}>0,B{r_tp}/(PL!H{PR["B"]["eps"]}*hkcny),"N/M")', "分母：模型情境 B FY2028 EPS × 匯率", "0.0")
    V.cell(r_up, 2).number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0.00"
    for L in "BCD": V[f"{L}{r_gate}"].number_format = "#,##0.0"
    V.column_dimensions["A"].width = 74; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, cmpu=r_cmpu, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, ca=r_ca, pd2=r_pd2, pe=r_pe, gj=r_gj, wk=r_wk, evs=r_evs, evm=r_evm, rel=r_rel, sg1=r_sg1, sg2=r_sg2, ipe=r_ipe, lq=r_lq, px=r_px, vg=r_vg)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司實績、市場共識同口徑對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或共識", "差", "說明"])
    B_ = PR["B"]
    rc = [("FY2026 營收 vs 共識（資料商，財報前，只作歷史對照）", f"=PL!F{B_['rev']}", "=cons_rev26", "共識 9 位、日期未標示；口徑同（IFRS 營收）"),
          ("FY2026 EPS vs 共識（人民幣元）", f"=PL!F{B_['eps']}", "=cons_eps26", "共識 0.14 隱含下半年淨利約 3.6 億（上半年 −1.77 億）；模型遠低於共識"),
          ("FY2026 下半年隱含營收 vs 2025 下半年 ×(1＋上半年增速)", f"=PL!F{B_['rev']}-Q!D6", "=Q!C6*(Q!D6/Q!B6)", "對照＝H2'25 × (H1'26 ÷ H1'25)"),
          ("FY2026 下半年隱含營業利益（本模型口徑）vs 2025 下半年", f"=PL!F{B_['op']}-Q!D14", "=Q!C14", "下半年營收季節性高（FY2025 下半年占 68%）"),
          ("FY2025 公司經營虧損 ＝ 本模型營業利益 ＋ 其他收入及虧損淨額（應為 0）", f"=ROUND(PL!E{B_['op']}+h_oi_25-(-28.778),3)", 0, "公司 FY2025 經營虧損 −28.8（2025 年報）"),
          ("FY2025 隱含份額（倉儲履約 ÷ 行動機器人市場；線 5 匯率）vs 線 5 份額（只動營收口徑）", f"=PL!E{B_['shp']}", f"={LINE['WH_GEEKPLUS_SHARE']['b']}", "線 5 用合併營收（含工業搬運），本模型只用倉儲履約；匯率 7.15 為線模型 Assumed（csv 現匯 6.67）"),
          ("FY2032 隱含份額（情境 B）", f"=PL!L{B_['shp']}", f"={LINE['WH_GEEKPLUS_SHARE']['b']}", "情境 B 份額年變化為 0 時份額持平（口徑同 FY2025）"),
          ("加權目標價 vs 賣方平均目標價（港元）", "=tp", "=cons_tp", "賣方 10 位平均（Interested-party；2026-09-05 更新，財報前）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，建置程式分批格點搜尋；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 64; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高或指定值），整本重建並以 LibreOffice 重算後讀加權目標價（不是閉式近似）；結果為常數快照"
    SE["A2"] = "build_id"; SE["A3"] = "模型狀態"; SE["A4"] = "In 頁輸入指紋（Sens 計算時）"; SE["A5"] = "In 頁輸入指紋（本檔即時）"
    SE["D5"] = f"=SUMPRODUCT(In!E4:E{LAST_IN},ROW(In!E4:E{LAST_IN}))"
    hdr(SE, 7, ["驅動", "設定", "鍵", "輸入值", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    if sens:
        SE["D2"] = sens["build_id"]; SE["D4"] = sens["fp"]
        SE["D3"] = f'=IF(AND(ABS(D5-D4)<1E-6,ABS(Val!B{r_tp}-{sens["base"]["tp"]})<0.0001),"CURRENT","STALE")'
        sr = 8
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            if isinstance(cs.get("val"), (int, float)): SE.cell(sr, 4, cs["val"])
            SE.cell(sr, 5, cs["tp"]).number_format = "0.00"; SE.cell(sr, 6, cs["tp"] / sens["base"]["tp"] - 1).number_format = "+0.0%;-0.0%"
            SE.cell(sr, 7, cs["dcf"]).number_format = "0.00"; SE.cell(sr, 8, cs["dcf"] / sens["base"]["dcf"] - 1).number_format = "+0.0%;-0.0%"
            SE.cell(sr, 9, cs["rt"]); sr += 1
        sr += 1
        SE.cell(sr, 1, "讀法：相對基準＝案例加權目標價 ÷ 基準加權目標價 − 1；同一案例 DCF 單法與加權目標價並列（第 10 節第 32 條）。情境專屬驅動以「情境 B 改為敏感度低／高」呈現，情境 A／C 不動。")
        SE.cell(sr + 1, 1, "說明：折舊、資本支出、營運資金隨營收等比例（第 44 條）；相對倍數係數 rel_k 基準即 1.00，故「相對倍數係數＝1」案例＝基準。尾端案例（美國、最大客戶）基準為 0、只動營收（上界的反面：未同步降毛利率）。pay_r、oth_rev 等無區間的鍵不列案例。")
    SE.column_dimensions["A"].width = 52; SE.column_dimensions["B"].width = 40

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = f"對帳 — 本模型使用的線層級變數（讀 {LINE_REL}「下游取數」頁；與線模型不一致的自設值標明理由）"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型是否一致"])
    def lv(k): return f"{LINE[k]['lo']:.3f}／{LINE[k]['b']:.3f}／{LINE[k]['hi']:.3f}" if LINE[k]["lo"] is not None else f"{LINE[k]['b']:.4f}"
    rec = [(f"行動機器人營收 {y}（低／基準／高）", "5", lv(f"WH_AMR_MKT_{y}"), "十億 USD", "Interact 口徑；含製造業與自主堆高機", f"{LINE_REL} 下游取數!C{LINE[f'WH_AMR_MKT_{y}']['row']}:E{LINE[f'WH_AMR_MKT_{y}']['row']}",
            "Derived", f"In mk{y % 100}_A／B／C → PL 命題模組列與倉儲履約營收", "一致（情境 A＝高、B＝基準、C＝低）") for y in MKY]
    rec += [("行動機器人營收 2025", "5", f"{MK25:.3f}", "十億 USD", "同上", f"{LINE_REL} 情境!B6", "Derived", "In mk25 → PL FY2025 隱含份額（只作對照）", "一致"),
            ("極智嘉占行動機器人營收 2025", "5", f"{LINE['WH_GEEKPLUS_SHARE']['b']:.4f}", "比例", "線 5 用合併營收 ÷ 市場、匯率 7.15（Assumed）", lref("WH_GEEKPLUS_SHARE"), "Derived", "Recon 份額對照；shr_* 說明", "口徑不同：本模型用倉儲履約（不含工業搬運），FY2025 隱含份額見 PL；未進公式"),
            ("線模型份額換算匯率", "5", f"{LFX}", "RMB/USD", "線模型自標 Assumed", f"{LINE_REL} 輸入!E44", "Assumed", "In lfx → PL 隱含份額列（只作對照）", "不一致：csv 2026-11-20 現匯 6.6665；為與線模型份額同口徑，對照列沿用 7.15，不在本模型靜默改線值"),
            ("AMR 年降價（線 5 輸入）", "5", "基準 5%（3–8%）", "%/年", "AMR 單價", f"{LINE_REL} 輸入!D18:F18", "Assumed", "shr_C、gmd_C 說明（價格競爭）", "未直接進公式：本模型以營收相對市場的變化（份額＋價格）合併表達"),
            ("每台 AMR 回本期 2026（穩態）／2030", "5", f"{LINE['WH_AMR_PAYBACK']['b']:.2f}／{LINE['WH_AMR_PAYBACK_2030']['b']:.2f}", "年", "美國、貨到人、兩班", lref("WH_AMR_PAYBACK") + "；" + lref("WH_AMR_PAYBACK_2030"), "Derived", "命題說明（需求面：回本期 1–2 年支持市場成長）；未進公式", "一致（只作需求面背景）"),
            ("AMR 貨到人每訂單行成本 vs 人工", "5", f"{LINE['WH_COST_AMR']['b']:.3f} vs {LINE['WH_COST_MANUAL']['b']:.3f}", "USD/訂單行", "美國", lref("WH_COST_AMR") + "；" + lref("WH_COST_MANUAL"), "Derived", "命題說明；未進公式", "一致（只作需求面背景）"),
            ("每台 AMR 全含資本支出", "5", f"{LINE['WH_AMR_CAPEX_UNIT']['b']:,.0f}", "USD/台", "單價＋整合＋工作站分攤", lref("WH_AMR_CAPEX_UNIT"), "Derived", "未進公式（公司未揭露台數與單價）", "—"),
            ("人形每人工當量小時成本", "5", f"{LINE['WH_HUMANOID_COST_HR']['b']:.1f}", "USD/時", "Agility RaaS 示意", lref("WH_HUMANOID_COST_HR"), "Derived", "emb_rd 說明（具身智能無營收，不建營收列）", "一致（只作背景）")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 30; RE.column_dimensions["F"].width = 60; RE.column_dimensions["I"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("美國關稅或對中國製機器人的進口限制（美國占營收 26.5%）", "是（尾端 Sens us_cut）", "事件型尾端，基準 0", "美國關稅清單、301／232 調查、美國營收占比"),
             ("最大客戶流失（占營收 10.6%；前五大 42.9%）", "是（尾端 Sens cust_cut）", "事件型尾端，基準 0", "年報主要客戶占比、大型電商客戶資本支出"),
             ("中國 AMR 廠價格戰", "是（shr_C、gmd_C）", "情境 C 與 Sens", "毛利率（半年報）、海外與中國毛利率差"),
             ("匯兌損益（營收 75% 以上在海外，功能貨幣人民幣）", "否", "不預測匯兌；上半年匯兌損失 104.6", "人民幣兌美元、歐元"),
             ("股價自高點下跌約 84%、主因未找到", "否", "未找到下跌主因（解禁、股東減持、上半年虧損擴大為可能因素，未證實）", "港交所權益披露、賣空比率"),
             ("具身智能投入擴大、再度虧損", "部分（emb_rd 高案）", "人形與移動操作尚無營收", "研發費用、具身智能出貨（目標三年 1 萬台）"),
             ("RSU／選擇權稀釋", "否", "未取得流通在外數量", "年報股份激勵附註")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 54; RK.column_dimensions["C"].width = 60

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([IR26, AR25, SAIS, SAFC, SA + " statistics 頁（SYM、AUTO、KGX、6383、2432、OCDO、2590）", INV, SINA, YCL, CNRF, DAMO, LINE_REL + "（下游取數、情境、輸入頁）",
                            "longbridge 2026-07-09 報導（買回 11.20–11.67 港元）https://longbridge.com/news/292135057"], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f'="極智嘉 Geek+（2590.HK）個股模型 {VER}｜build_id "&Sens!D2&"｜資料截止日 2026-11-23｜成品日 2026-11-23｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'
    RD["A2"] = '="資料截止日 2026-11-23（財報為 2026 中期業績，2026-08-28）；成品日 2026-11-23；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.000")&" 港元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：極智嘉的股價由倉儲 AMR（貨到人）營收能否跟上行動機器人市場成長、並把毛利率 35% 左右的營收轉成正的營業利益決定——倉儲履約營收＝行動機器人市場（線 5）× 份額；工業搬運與其他以成長路徑建模；具身智能無營收，只在費用與 Sens 呈現"
    RD["A4"] = (f'="在情境 B FY2026 營收年增 "&TEXT(g26_B,"0%")&"、倉儲履約跟隨行動機器人市場（份額年變化 "&TEXT(shr_B,"0.0%")&"）、毛利率 "&TEXT(gm0,"0.0%")&" 持平、營業費用彈性 "&TEXT(e_ox,"0.00")&" 的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 港元（DCF "&TEXT(Val!B{r_dcft},"0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0")&" 百萬人民幣）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 港元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")&"；現行 EV ÷ 營收 "&TEXT(Val!B{r_evs},"0.00")&" 倍、同業中位數 "&TEXT(Val!B{r_evm},"0.00")&" 倍（相對 "&TEXT(Val!B{r_rel},"0.00")&"）"'
    RD["A6"] = "失效條件（可觀察）：半年營收年增跌破 10%；新簽訂單年增轉負；毛利率跌破 32%；下半年（季節性高）營業利益仍為負；美國營收占比因關稅明顯下降"
    RD["A7"] = "範圍：不讀 Siliconomics／Tokenomics／Wattonomics（算力線）；自用 AI 讀法、Wattonomics 上線時程等規格條文不適用。具身智能（人形 Gino、移動操作）無營收揭露，未建營收列。HTML 展示層見 dist/ 同名 .html"
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
LO = "file://" + os.path.join(tempfile.gettempdir(), "lo_geekplus_profile")
def recalc_many(srcs, outdir):
    os.makedirs(outdir, exist_ok=True)
    for attempt in range(2):
        try:
            subprocess.run(["soffice", f"-env:UserInstallation={LO}", "--headless", "--convert-to", "xlsx", "--outdir", outdir] + srcs,
                           check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=max(180, 6 * len(srcs)))
            break
        except subprocess.TimeoutExpired:
            subprocess.run(["pkill", "-f", "lo_geekplus_profile"]);
            if attempt: raise
    return [os.path.join(outdir, os.path.basename(s)) for s in srcs]

def read_vals(path, VR, DR=None):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]
    d = dict(tp=V[f"B{VR['tp']}"].value, dcf=V[f"B{VR['dcf']}"].value, cmp=V[f"B{VR['cmp']}"].value, up=V[f"B{VR['up']}"].value, rt=V[f"B{VR['rt']}"].value)
    if DR:
        I = wb["In"]; w = {}; wd = None
        for row in I.iter_rows(values_only=True):
            if row[0] in ("w_A", "w_B", "w_C"): w[row[0][-1]] = row[4]
            if row[0] == "w_dcf": wd = row[4]
        d["alt_dcf"] = sum(w[s] * wb["DCF"][f"B{DR[s]['alt']}"].value for s in SC)
        d["alt_tp"] = wd * d["alt_dcf"] + (1 - wd) * d["cmp"]
        d["fp"] = wb["Sens"]["D5"].value
    return d

def run_cases(ovs):
    """ovs：[(tag, override dict)] → {tag: vals}；一次 LibreOffice 轉換全部"""
    d = tempfile.mkdtemp(prefix="gkp_"); paths = []; VRs = {}
    for tag, ov in ovs:
        p = os.path.join(d, f"c_{tag}.xlsx"); VR, _, DR = build(p, ov); paths.append(p); VRs[tag] = (VR, DR)
    outs = []
    for i in range(0, len(paths), 40):
        outs += recalc_many(paths[i:i + 40], os.path.join(d, "o"))
    res = {}
    for (tag, _), o in zip(ovs, outs):
        VR, DR = VRs[tag]; res[tag] = read_vals(o, VR, DR if tag == "base" else None)
    shutil.rmtree(d, ignore_errors=True); return res

def in_val(k, col):
    for row in ROWS:
        if row[0] == k:
            return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)

def rating(u): return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"

CASES = [("g26_B", "FY2026 營收年增（情境 B）"), ("gimt26", "FY2026 工業搬運年增"), ("oi26", "FY2026 其他收入淨額"), ("shr_B", "倉儲履約相對市場年變化（情境 B）"),
         ("gimt_B", "工業搬運年增（情境 B）"), ("gm0", "毛利率起點"), ("gmd_B", "毛利率每年變化（情境 B）"), ("e_ox", "營業費用彈性"), ("emb_rd", "具身智能額外研發"),
         ("imp_r", "應收款減值率"), ("oi_lt", "FY2027 起其他收入淨額"), ("tax_w", "所得稅率"), ("us_cut", "尾端：美國營收流失"), ("cust_cut", "尾端：最大客戶流失"),
         ("beta", "β"), ("erp", "ERP"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("fc_da", "折舊 ÷ 營收"), ("fc_cx", "資本支出 ÷ 營收"),
         ("fc_wc", "營運資金 ÷ 營收增量"), ("sh_v", "評價日股數"), ("bb_post", "6-30 後買回現金"), ("ev_med", "同業 EV ÷ 營收倍數（子群）"), ("rel_k", "相對倍數係數"), ("w_dcf", "DCF 權重")]
MKK = [f"mk{y % 100}_B" for y in MKY]

def write_readme(path, VR, PR, DR, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    IV = {I.cell(r, 1).value: I.cell(r, 5).value for r in range(4, I.max_row + 1) if I.cell(r, 1).value}
    f2 = lambda x: f"{x:,.2f}"; f0 = lambda x: f"{x:,.0f}"; pc = lambda x: f"{x*100:+.1f}%"
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f2(g(chr(66+i)+str(VR['dcfr'])))} | {f2(g(chr(66+i)+str(VR['ca'])))} | {f2(g(chr(66+i)+str(VR['tps'])))} |" for i, s in enumerate(SC))
    def yr(s, k, fmt="{:,.1f}"):
        out = []
        for y in YRS:
            v = P[f"{COL[y]}{PR[s][k]}"].value
            out.append(fmt.format(v) if isinstance(v, (int, float)) else "—")
        return " | ".join(out)
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:8]
    rts = sorted(set(c["rt"] for c in sens["cases"]))
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {r['val'] if isinstance(r['val'], str) else f'{r[chr(118)+chr(97)+chr(108)]:.4f}'} | {r['base']} | {r['note']} |" for r in revs)
    gate = [g(f"{c}{VR['gate']}") for c in "BCD"]
    sens_by = {c["key"] + ("_lo" if c["drv"].endswith("（低）") else "_hi" if c["drv"].endswith("（高）") else ""): c for c in sens["cases"]}
    def sc(k): c = sens_by.get(k); return f"{f2(c['tp'])}（{pc(c['tp']/sens['base']['tp']-1)}，{c['rt']}）" if c else "—"
    ends = [sens_by[k]["rt"] for k in ("ev_med_lo", "rel_k_lo", "w_C") if k in sens_by]
    base_rt = g(f"B{VR['rt']}")
    same = all(e == base_rt for e in ends)
    order = ["強力放空", "放空", "中立", "買進", "強力買進"]
    lowest = min(ends + [base_rt], key=order.index)
    robust = (f"三個端點案例的估值評等都是{base_rt}（區間穩健性：穩健）。" if same else
              f"三個端點案例的估值評等為{'、'.join(ends)}，最低為{lowest}；基準評等{base_rt}只在點值成立（區間穩健性：{lowest}以上穩健、{base_rt}不穩健）。")
    txt = f"""# physical_ai/geekplus — 極智嘉 Geek+（2590.HK）個股模型 {VER}

{RD['A2'].value}｜build_id `{sens['build_id']}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

港交所主板上市（2025-07，發行價 16.80 港元）；報表人民幣、股價港元。本 README 的數字由 `build_geekplus.py` 從重算後的成品讀出產生，不手寫。

**本版沒有做的部分（高不確定性原則：證據到哪裡判斷到哪裡）：** 公司只有一個經營分部、未給全年量化指引；具身智能（人形、移動操作）沒有營收揭露，不建營收列，只在費用（emb_rd）與 Sens 呈現；出貨台數、單價、訂閱營收、地區別毛利未揭露，命題模組以「營收相對行動機器人市場的變化」合併表達份額與價格，不拆量價；FY2023 只有資料商合計值；RSU／選擇權數量、聯營投資帳面值未取得（不進股數與淨現金）。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

**結論的證據邊界：** 加權目標價的上檔主要來自 {'Comparable' if g(f"B{VR['a2']}") > g(f"B{VR['a1']}") else 'DCF'}（Comparable 用同業 EV ÷ 營收中位數 {g(f"B{VR['evm']}"):.2f} 倍，Analogy；極智嘉現行 {g(f"B{VR['evs']}"):.2f} 倍）。只看 DCF 為 {f2(dcf)} 港元（{pc(dcf/IV['px']-1)}）。改用系統子群中位數 {g(f"B{VR['sg1']}"):.2f} 倍：{sc('ev_med_lo')}；維持現行相對折價（相對倍數係數 {sens['rel_now']:.2f}）：{sc('rel_k_lo')}；只看情境 C：{sc('w_C')}。{robust}股價自 52 週高點 33.50 港元跌到 {IV['px']:.3f}，跌幅主因未找到。

## 估值（評價日 2026-12-31；目標時點約 12 個月；港元）
| 情境 | DCF（推到 12 個月後） | Comparable（EV ÷ 營收） | 情境目標價 |
|---|---|---|---|
{scn}

- 加權目標價 **{f2(tp)} 港元**（DCF {f2(dcf)}、Comparable {f2(cmp_)}），現價 {IV['px']:.3f}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")}（情境 A／B／C FY2026–FY2030 自由現金流扣股利 {f0(gate[0])}／{f0(gate[1])}／{f0(gate[2])} 百萬人民幣）→ 投資建議 **{g(f"B{VR['rec']}")}**。情境 B 期間最低年末淨現金 {f0(g(f"B{VR['lq']}"))} 百萬人民幣（不需新增舉債）。
- 方法：DCF 標準版 FCFF（WACC／ke／g／RONIC＝{g(f"B{VR['wk']}")}）50%＋Comparable 50%（同業 6 家 EV ÷ 近四季營收中位數 × FY2027 營收＋淨現金；目標時點的近四季營收以 FY2027 代）。相對倍數係數 1.00（未調整：缺歷史相對倍數）。
- 隱含 EV ÷ FY2027 營收（加權目標價）{g(f"B{VR['ipe']}"):.2f} 倍；隱含 P/E（÷ 情境 B FY2028 EPS）{g(f"B{VR['pe']}"):.1f} 倍。賣方平均目標價 {IV['cons_tp']:.2f} 港元（10 位，2026-09-05 更新，財報前，只作對照）。

## 營收與損益（情境 B，百萬人民幣；FY2023–FY2025 實績，FY2026 起模型）
| 列 | {" | ".join(f"FY{y}" for y in YRS)} |
|---|{"---|" * len(YRS)}
| 行動機器人市場（十億美元，線 5） | {yr('B','mkt','{:.2f}')} |
| 倉儲履約營收 | {yr('B','wf')} |
| 工業搬運營收 | {yr('B','imt')} |
| 營收（合併） | {yr('B','rev')} |
| 毛利率 | {yr('B','gm','{:.1%}')} |
| 營業利益（毛利 − 營業費用） | {yr('B','op')} |
| 稅後淨利 | {yr('B','ni')} |
| EPS（人民幣元） | {yr('B','eps','{:.3f}')} |
| FCFF | {yr('B','fcff')} |

- 命題模組：倉儲履約營收 FY2027 起＝前一年 × 行動機器人市場成長（線 5 下游取數，情境 A＝高、B＝基準、C＝低）×(1＋相對變化 shr)；FY2026＝營收（上半年實績推估年增）− 工業搬運 − 其他。
- 毛利率起點＝最近四季 35.7%（Derived），情境 B 持平、A／C 每年 ±0.5pt；營業費用＝前一年 ×(1＋彈性 × 營收年增)，彈性 0.65（FY2024→FY2025 實績，Derived）。
- 營業利益為本模型口徑（毛利 − 研發 − 銷售 − 行政 − 應收減值）；公司「經營虧損」另含其他收入及虧損淨額（匯兌、利息、補助），見 Recon。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- FY2026 營收年增 31%／25%／20%（A／B／C；上半年 +25.3%、新簽訂單 +35.5%）；倉儲履約相對市場 +5%／0%／−6% 每年；工業搬運 +20%／+12%／+3%。
- 行動機器人市場（十億美元）2026／2028／2030／2032：基準 {IV['mk26_B']:.2f}／{IV['mk28_B']:.2f}／{IV['mk30_B']:.2f}／{IV['mk32_B']:.2f}（線 5）。
- 折現：rf 1.61%（人民幣 10 年期，與中際旭創同值，逾期）、β {IV['beta']:.2f}（同業 5 家 Blume，Analogy）、ERP 5%、CRP 0.91%、g＝MIN(3%, rf)、RONIC＝WACC；所得稅 15%（只對正的稅前）；折舊 1.3%、資本支出 2.5%、營運資金 12.9% × 營收增量。
- 評價日股數 {IV['sh_v']:.1f} 百萬股（已發行 − 庫存股 − 7 月買回）；淨現金 {f0(wb['DCF']['B10'].value)} 百萬人民幣（6-30 現金 2,201 − 借款 − 租賃 − 7 月買回假設 138）。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| FY2026 營收年增（g26） | 2026 年度業績（約 2027-03）；下半年營收 ≥ 上半年 × 2.09 |
| 倉儲履約相對市場（shr） | 半年報倉儲履約營收年增 vs Interact 行動機器人營收年增；新簽訂單年增 |
| 毛利率（gm0、gmd） | 半年報毛利率、海外與中國毛利率差 |
| 營業費用彈性（e_ox） | 半年報研發、銷售、行政費用年增 ÷ 營收年增；具身智能研發金額 |
| Comparable 倍數（ev_med，Analogy） | 同業 EV ÷ 營收；極智嘉自身倍數是否向同業收斂 |

失效條件：{RD['A6'].value.replace('失效條件（可觀察）：', '')}。評等翻轉需要的條件見下方反解（Sens 各案例評等：{'、'.join(rts)}）。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f2(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁；各案例估值評等出現「{'、'.join(rts)}」。

## 反解（整本重算，分批格點搜尋；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
讀 `{LINE_REL}`「下游取數」：行動機器人營收 2026／2028／2030／2032（D4:D7 基準、C 欄低、E 欄高）→ mk*_A／B／C → 命題模組；2025 市場（情境!B6）與極智嘉份額（D26）只作對照。與線模型不一致：線模型份額換算匯率 7.15（Assumed）vs csv 現匯 6.67，本模型沿用 7.15 只作份額對照、營收推導不用匯率；線 5 份額用合併營收，本模型用倉儲履約。AMR 回本期、每訂單行成本只作需求面背景，未進公式。

## 限制與未取得資料
- 公司未給全年量化指引；共識 FY2026 營收 42 億、EPS 0.14（日期未標示，財報前）只作對照；共識 FY2027 起未取得（付費牆）。
- 同業資料取自資料商 statistics 頁（頁面日期 2026-09-11～10-23，Symbotic 頁面日期未標示、近四季營收停在 FY2025）；市值以 2026-11-20 收盤重算，EV 其餘項目與營收未更新；各群組少於 8 家，不做迴歸；EV 的租賃口徑未逐家核對。
- β 只有一個來源（同業資料商 5 年 β），個股 β 未取得（上市未滿 2 年）。
- FY2023 為資料商取整值（招股書原文未逐項核對）；FY2023 產品線拆分、銷售與行政分列未取得。
- 歷史 EPS 用公司公告值，FY2023–FY2024 與年度利潤 ÷ 加權股數不一致（上市前優先股口徑），分母未取得。
- 7 月買回 16.2 百萬股的金額未取得（假設 10 港元）；8–11 月是否續買未核對（資料商流通股 1,289 百萬股，Sens 低案）。
- 股價 52 週下跌約 84% 的主因未找到（「沒找到」，不是「分不出來」）。
- 獨立核對（2026-11-24，未參與製作者）：數字一致性通過；結論穩健性見「結論的證據邊界」。未做：check_display（無 display_check.json）。

## 檔案
- `build_geekplus.py`：建置＋整本重算 Sens／反解＋產生本 README、`segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-23）：初版；現價 5.405 港元（2026-11-20 收盤）；加權目標價 12.17、上檔 +125.2%、估值評等 強力買進、投資建議 強力買進。
- v0.1.1（2026-11-24）：獨立核對後 DCF 權重由一半改為七成（Comparable 三成），其餘輸入不變；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


def main():
    beta_b = 0.67 * statistics.median([2.05, 1.03, 1.98, 2.34, 2.11]) + 0.33
    # 先算基準，取得 ev_self／ev_med 以寫入 rel_k 低案（現行相對倍數）
    base0 = run_cases([("base", {})])["base"]
    tmpd = tempfile.mkdtemp(prefix="gkp0_"); p0 = os.path.join(tmpd, "b.xlsx"); VR0, _, _ = build(p0); o0 = recalc_many([p0], os.path.join(tmpd, "o"))[0]
    wv = openpyxl.load_workbook(o0, data_only=True); rel_now = wv["Val"][f"B{VR0['rel']}"].value
    evmed = wv["Val"][f"B{VR0['evm']}"].value; sg1 = wv["Val"][f"B{VR0['sg1']}"].value; sg2 = wv["Val"][f"B{VR0['sg2']}"].value
    shutil.rmtree(tmpd, ignore_errors=True)
    for i, row in enumerate(ROWS):
        if row[0] == "rel_k": ROWS[i] = row[:3] + (round(rel_now, 4),) + row[4:]
    base = run_cases([("base", {})])["base"]; px0 = in_val("px", "base")
    print("基準", base)
    jobs = []; meta = {}
    for k, drv in CASES:
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if k == "beta": v = beta_b + (-0.3 if side == "lo" else 0.3)
            if k == "ev_med": v = sg1 if side == "lo" else sg2
            if v is None: continue
            b_ = in_val(k, "base")
            if isinstance(b_, (int, float)) and abs(v - b_) < 1e-12: continue
            tag = f"{k}_{side}"; jobs.append((tag, {k: v}))
            meta[tag] = dict(drv=drv + ("（低）" if side == "lo" else "（高）"), set=f"{k}＝{v:.4g}", key=k, val=v)
    for side in ("lo", "hi"):
        ov = {k: in_val(k, side) for k in MKK}
        tag = f"mk_{side}"; jobs.append((tag, ov))
        meta[tag] = dict(drv="行動機器人市場 2026–2032（情境 B，四年同動，線 5 " + ("低" if side == "lo" else "高") + "）" + ("（低）" if side == "lo" else "（高）"),
                         set="線 5 下游取數 " + ("C" if side == "lo" else "E") + " 欄", key=", ".join(MKK), val=None)
    for s in SC:
        tag = f"only_{s}"; jobs.append((tag, {f"w_{x}": (1.0 if x == s else 0.0) for x in SC}))
        meta[tag] = dict(drv=f"只看情境 {s}", set=f"w_{s}＝1，其餘 0", key=f"w_{s}", val=None)
    res = run_cases(jobs)
    cases = []
    for tag, _ in jobs:
        r = res[tag]; cases.append(dict(**meta[tag], tp=r["tp"], dcf=r["dcf"], cmp=r["cmp"], up=r["up"], rt=r["rt"]))
        print(tag, round(r["tp"], 3), r["rt"])
    up_ = base["alt_tp"] / px0 - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", val=None, tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rating(up_)))
    # 反解：分批格點搜尋（每輪 9 點、3 輪）
    tiers = [0.5, 0.25, -0.25, -0.5]; up0 = base["up"]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    targets = [("目標價＝現價（上檔 0%）", 0.0)] + ([(f"升一級邊界（上檔 {upper:+.0%}）", upper)] if upper is not None else []) + ([(f"降一級邊界（上檔 {lower:+.0%}）", lower)] if lower is not None else [])
    SOLVE = [("shr_B", -0.40, 0.30, "情境 B 倉儲履約相對行動機器人市場的年變化（FY2027–FY2032 每年；情境 A／C 不動）"),
             ("gm0", 0.05, 0.60, "毛利率起點（三情境同動；A／C 另加各自的年變化）"),
             ("ev_med", 0.0, 8.0, "同業 EV ÷ 營收倍數（Comparable；DCF 不動）")]
    revs = []
    for key, lo, hi, note in SOLVE:
        brackets = {lab: (lo, hi) for lab, _ in targets}
        for rnd in range(3):
            jobs = []
            for lab, tgt in targets:
                a, b2 = brackets[lab]
                if a is None: continue
                for j in range(9):
                    x = a + (b2 - a) * j / 8; jobs.append((f"{key}|{lab}|{j}", {key: x}))
            rs = run_cases(jobs)
            for lab, tgt in targets:
                a, b2 = brackets[lab]
                if a is None: continue
                xs = [a + (b2 - a) * j / 8 for j in range(9)]; fs = [rs[f"{key}|{lab}|{j}"]["up"] - tgt for j in range(9)]
                br = None
                for j in range(8):
                    if fs[j] == 0 or fs[j] * fs[j + 1] < 0: br = (xs[j], xs[j + 1]); break
                brackets[lab] = br if br else (None, None)
        for lab, tgt in targets:
            a, b2 = brackets[lab]
            revs.append(dict(label=lab, key=key, val=(a + b2) / 2 if a is not None else f"區間 {lo:g}～{hi:g} 內無解", base=round(float(in_val(key, 'base')) if isinstance(in_val(key, 'base'), (int, float)) else (base0.get('gm0') or 0), 4), note=note))
            print("反解", key, lab, revs[-1]["val"])
    # gm0、ev_med 基準為公式：基準值以重算值填
    sens = dict(base=base, cases=cases, fp=base["fp"], rel_now=rel_now, evmed=evmed, sg1=sg1, sg2=sg2)
    # 成品
    fpr = hashlib.sha1(json.dumps(ROWS, ensure_ascii=False, default=str).encode()).hexdigest()[:10]
    sens["build_id"] = f"{VER}-{fpr}"
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="gkpf_"); raw = os.path.join(tmp, NAME)
    VR, PR, DR = build(raw, {}, sens, revs)
    out = recalc_many([raw], os.path.join(tmp, "o"))[0]
    # 反解基準值：gm0、ev_med 為公式，改填重算值後再重算一次
    wv = openpyxl.load_workbook(out, data_only=True)
    gm0v = [r_ for r_ in wv["In"].iter_rows(values_only=True) if r_[0] == "gm0"][0][4]
    for rv_ in revs:
        if rv_["key"] == "gm0": rv_["base"] = round(gm0v, 4)
        if rv_["key"] == "ev_med": rv_["base"] = round(evmed, 4)
    VR, PR, DR = build(raw, {}, sens, revs); out = recalc_many([raw], os.path.join(tmp, "o"))[0]
    shutil.copy(out, os.path.join(dist, NAME))
    wbv = openpyxl.load_workbook(os.path.join(dist, NAME), data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10]); assert not errs, errs
    print("模型狀態", wbv["Sens"]["D3"].value)
    json.dump(sens, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1, default=str)
    write_readme(os.path.join(dist, NAME), VR, PR, DR, sens, revs)
    json.dump(dict(company="geekplus", ticker="2590.HK", model_sheet="Q",
                   official_segments=[dict(name=n, label=l, period="H1'26", value=v, unit="RMB m", source=IR26)
                                      for n, l, v in (("倉儲履約（Warehouse fulfilment）", "倉儲履約營收", 1201.553), ("工業搬運（Industrial material transport）", "工業搬運營收", 82.221),
                                                      ("其他（Others）", "其他營收", 0.125))],
                   sublines=[], check_labels=["檢查：產品線合計 − 營收（應為 0）"],
                   note="公司只有一個經營分部（年報分部附註：One）；以公司揭露的產品線營收作分部列。Q 頁半年度（H1'25、H2'25、H1'26）；PL 頁三情境另有「檢查：分部合計 − 營收（應為 0）」（FY2024 起）；Hist 頁 FY2024、FY2025 年度對帳。FY2023 產品線未取得"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)

if __name__ == "__main__":
    sys.path.insert(0, HERE)
    main()

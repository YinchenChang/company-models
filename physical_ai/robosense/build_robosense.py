# 速騰聚創 RoboSense（2498.HK）個股模型 v0.1 — AI實體應用 Project（線 1 Robotaxi／線 2 智慧車；LiDAR 另見線 6 共用零組件）
# 用法：python3 physical_ai/robosense/build_robosense.py   → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、segments.json、sens.json
# 會計年度＝曆年（12 月決算，IFRS）。金額：百萬人民幣（RMB m）；股數：百萬股；EPS：人民幣元；股價與目標價：港元（HK$）。出貨量：萬顆；單價：人民幣元／顆。
# 結構：README、In、Hist、Q（半年度分部與出貨）、PL（三情境：命題模組→分部營收→標準損益表→現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 線層級變數讀 physical_ai/lines/smart_car/dist/20261116_智慧車單位經濟_v0.1.xlsx「下游取數」（中國城市 NOA 新車數與滲透率），建置時讀值並寫明儲存格。
# 不讀 Siliconomics／Tokenomics／Wattonomics（實體 AI 線）。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261124"
NAME = f"{DATE}_速騰聚創個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032
HISTY = (2023, 2024, 2025)

# ---------------- 線模型（下游取數） ----------------
LINE_REL = "physical_ai/lines/smart_car/dist/20261116_智慧車單位經濟_v0.1.xlsx"
_lw = openpyxl.load_workbook(os.path.join(ROOT, LINE_REL), data_only=True)
_ld = _lw["下游取數"]
LINE = {}
for _r in range(3, _ld.max_row + 1):
    k, y = _ld.cell(_r, 1).value, _ld.cell(_r, 2).value
    if k and isinstance(y, int): LINE[f"{k}|{y}"] = dict(row=_r, lo=_ld.cell(_r, 3).value, b=_ld.cell(_r, 4).value, hi=_ld.cell(_r, 5).value)
NOA = "中國城市 NOA 新車數"; PEN = "中國城市 NOA 滲透率"
LATT = _lw["輸入"]["E22"].value            # 城市 NOA 車搭載 LiDAR 比例（線模型 輸入!E22，Assumed 0.65）
LNUM = _lw["輸入"]["E23"].value            # 每車 LiDAR 數（輸入!E23）
LASP = _lw["輸入"]["E24"].value            # ADAS LiDAR 單價 USD（輸入!E24，Analogy）
def lref(k, y, col="D"): return f"{LINE_REL} 下游取數!{col}{LINE[f'{k}|{y}']['row']}（{k} {y}）"
def lv3(k, y): d = LINE[f"{k}|{y}"]; return d["lo"], d["b"], d["hi"]

# ---------------- 出處 ----------------
IR26 = "公司 2026 年中期業績公告（港交所 2498，2026-08-26）https://www1.hkexnews.hk/listedco/listconews/sehk/2026/0826/2026082601087.pdf"
FY25 = "公司 2025 年度業績（2026-03-25；經 AAP、北京商報、36氪報導）https://www.aap.com.au/aapreleases/cision20260325ae18999"
GAS = "Gasgoo 2026-10-09：速騰聚創 Q3 LiDAR 出貨 59.56 萬顆（ADAS 45.04 萬、機器人 14.52 萬）https://autonews.gasgoo.com/articles/market-industry/robosense-quarterly-lidar-sales-hit-new-high-in-q3-2026-2108495209121140737"
SAIS = "stockanalysis.com HKG:2498 income statement（資料商彙整；讀取 2026-11-24）https://stockanalysis.com/quote/hkg/2498/financials/income-statement/"
SAST = "stockanalysis.com HKG:2498 statistics（資料商；財務更新 2026-11-23）https://stockanalysis.com/quote/hkg/2498/statistics/"
YCL = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
CNRF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與極智嘉、中際旭創模型同一取值，跨模型一致；逾期：與股價基準日相差逾 2 個交易日）"
DAMO = "Damodaran Country Risk Premium（2026-01-05）：China 0.91%（與其他中國模型同值）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整（立場：中立彙整，來源為交易所資料）"

# ---------------- In 列 ----------------
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. FY2026 錨點（2026 上半年實績＋Q3 出貨公告＋Q4 假設；公司未給全年量化指引）")
inp("u_a_h1", "2026 上半年 ADAS LiDAR 出貨", "萬顆", None, 43.66, None, "Verified", IR26 + "：上半年 ADAS 出貨 436,600 顆（+98.0%）", "")
inp("u_a_q3", "2026 Q3 ADAS LiDAR 出貨", "萬顆", None, 45.04, None, "Interested-party", GAS, "公司季度公告經媒體報導；財報未出", CO)
inp("u_r_h1", "2026 上半年機器人等 LiDAR 出貨", "萬顆", None, 28.26, None, "Verified", IR26 + "：機器人 LiDAR 282,600 顆（+510.4%）", "")
inp("u_r_q3", "2026 Q3 機器人等 LiDAR 出貨", "萬顆", None, 14.52, None, "Interested-party", GAS, "Q3 機器人出貨低於 Q2（H1 28.26 萬）；原因未找到", CO)
for s, v, lo, hi, note in (("A", 60.0, None, None, "樂觀：Q4 ADAS 出貨較 Q3 +33%"), ("B", 52.0, 45.0, 60.0, "基準：Q4 較 Q3 +15%（Q2 29.18 → Q3 45.04 的成長放緩）"), ("C", 45.0, None, None, "保守：Q4 持平 Q3")):
    inp(f"u_a_q4_{s}", f"情境 {s}：2026 Q4 ADAS 出貨", "萬顆", lo, v, hi, "Assumed", note, "公司未給出貨指引；Q3 年增 200%（Gasgoo）")
for s, v, lo, hi, note in (("A", 26.0, None, None, "樂觀：Q4 機器人出貨較 Q3 +80%（SPAD-SoC E2 交付，公司稱下半年超過 20 萬顆含車用）"), ("B", 19.0, 14.0, 26.0, "基準：Q4 較 Q3 +30%"), ("C", 14.5, None, None, "保守：Q4 持平 Q3")):
    inp(f"u_r_q4_{s}", f"情境 {s}：2026 Q4 機器人等 LiDAR 出貨", "萬顆", lo, v, hi, "Assumed", note, "2025 全年機器人 30.3 萬顆；2026 前三季 42.78 萬顆")
inp("rho", "機器人占產品營收比重（2026 上半年）", "%", 0.44, 0.47, 0.50, "Assumed", IR26 + "：機器人營收占產品銷售營收「近一半」（未給金額）", "低 44%、高 50% 為對「近一半」的讀法；用於拆出 ADAS 與機器人單價")
inp("prod_h1", "2026 上半年產品營收", "RMB m", None, 958.003, None, "Verified", IR26, "解決方案 17.2、服務與其他 44.5（技術開發服務專案）")
inp("asp_a0", "ADAS LiDAR 平均單價（2026 上半年，推導）", "RMB/顆", None, "=prod_h1*(1-rho)/u_a_h1*100", None, "Derived", "（產品營收 × (1 − 機器人比重) ÷ ADAS 出貨）；公司未揭露單價", "對照線模型 ADAS LiDAR 單價 200 美元（約 1,300 人民幣；Analogy 150–300 美元）")
inp("asp_r0", "機器人等 LiDAR 平均單價（2026 上半年，推導）", "RMB/顆", None, "=prod_h1*rho/u_r_h1*100", None, "Derived", "（產品營收 × 機器人比重 ÷ 機器人出貨）；含割草機、配送、人形等混合", "公司未揭露單價")
for s, v, lo, hi in (("A", -0.01, None, None), ("B", -0.04, -0.08, 0.0), ("C", -0.08, None, None)):
    inp(f"dasp26_{s}", f"情境 {s}：FY2026 全年 ADAS 單價相對上半年", "%", lo, v, hi, "Assumed", IR26 + "：上半年毛利率下降主因為單價下降與原料成本上升", "下半年單價續降、新品 E2 單價未取得")
for s, v, lo, hi in (("A", -0.05, None, None), ("B", -0.10, -0.18, -0.03), ("C", -0.16, None, None)):
    inp(f"dasp_{s}", f"情境 {s}：ADAS 單價年變化（FY2027 起）", "%/年", lo, v, hi, "Assumed", "線 2 智慧車與共用零組件線：ADAS LiDAR 基準年降 10%；同業 ASP 年減約 35% 為上界風險（共用零組件單位經濟 輸入!H112、追蹤指標!D8）", "中國價格競爭；2026 全球產能稼動約 56%（供給偏鬆）")
for s, v, lo, hi in (("A", -0.04, None, None), ("B", -0.08, -0.14, -0.02), ("C", -0.12, None, None)):
    inp(f"dasp_r_{s}", f"情境 {s}：機器人 LiDAR 單價年變化", "%/年", lo, v, hi, "Assumed", "機器人 LiDAR 毛利率較高（2025Q4 41.5%，媒體引述）；價格壓力小於車用", "線模型未涵蓋割草機等，無線變數可讀")
inp("oth26", "FY2026 解決方案、服務與其他營收（上半年 61.7）", "RMB m", 80.0, 130.0, 200.0, "Assumed", IR26 + "：上半年解決方案 17.2＋服務與其他 44.5＝61.7（含一次性技術開發服務專案）", "下半年假設 68；觸覺、關節模組 Q4 開始交付")
for s, v, lo, hi in (("A", 0.40, None, None), ("B", 0.20, 0.0, 0.40), ("C", 0.0, None, None)):
    inp(f"goth_{s}", f"情境 {s}：解決方案與其他營收年增（FY2027 起）", "%/年", lo, v, hi, "Assumed", "觸覺感測、關節模組、空間相機規劃 2026Q4 量產（" + IR26 + "）；金額未揭露", "營收占比小（約 5%），對結論影響有限")

sec("B. 命題模組：ADAS LiDAR 出貨＝中國城市 NOA 新車（線 2 下游取數）× 每台 NOA 車對應速騰 ADAS 出貨（含搭載率、每車顆數、份額）")
for y in (2026, 2027, 2028, 2030, 2032):
    for s, ci, col in (("A", "hi", "E"), ("B", "b", "D"), ("C", "lo", "C")):
        v = LINE[f"{PEN}|{y}"][ci]
        lo_, hi_ = (LINE[f"{PEN}|{y}"]["lo"], LINE[f"{PEN}|{y}"]["hi"]) if s == "B" else (None, None)
        inp(f"pen{y % 100}_{s}", f"情境 {s}：中國城市 NOA 滲透率 {y}（線 2 {'高' if s=='A' else '基準' if s=='B' else '低'}）", "%", lo_, v, hi_, "Derived",
            lref(PEN, y, col) + "：國內乘用車零售口徑", "敏感度低／高＝線模型低／高（五年同動，一案）" if s == "B" else "情境 A＝線模型高、C＝線模型低")
for s, ci, col in (("A", "hi", "E"), ("B", "b", "D"), ("C", "lo", "C")):
    inp(f"noa26_{s}", f"情境 {s}：中國城市 NOA 新車數 2026", "萬台", None, LINE[f"{NOA}|2026"][ci], None, "Derived", lref(NOA, 2026, col), "零售量＝新車數 ÷ 滲透率（2026 與 2030 皆約 1,830 萬台，線模型為常數）")
inp("lidar_att", "線模型：城市 NOA 車搭載 LiDAR 比例（只作對照）", "%", None, LATT, None, "Assumed", f"{LINE_REL} 輸入!E22（線模型自標 Assumed，區間 50–80%；線模型引用 Gasgoo 速騰 Q3 出貨）", "只作對照，不進模型：本模型把搭載率、每車顆數、份額合併為一個係數；見對帳頁")
inp("lidar_n", "線模型：每車 LiDAR 顆數（只作對照）", "顆", None, LNUM, None, "Assumed", f"{LINE_REL} 輸入!E23（區間 1–1.2）", "只作對照，不進模型：同上")
for s, v, lo, hi in (("A", 0.03, None, None), ("B", -0.02, -0.06, 0.02), ("C", -0.08, None, None)):
    inp(f"shr_{s}", f"情境 {s}：每台 NOA 車對應的速騰 ADAS 出貨，年變化（FY2027 起）", "%/年", lo, v, hi, "Assumed",
        "合併搭載率、每車顆數與份額；2025 年速騰與禾賽兩家出貨比 36%／64%（共用零組件單位經濟 份額頁）；禾賽、華為、圖達通競爭", "A：新車型與海外客戶放量；C：份額流失與搭載率下降")
for s, v, lo, hi in (("A", 0.70, None, None), ("B", 0.45, 0.20, 0.70), ("C", 0.20, None, None)):
    inp(f"gr_{s}", f"情境 {s}：機器人等 LiDAR 出貨年增（FY2027）", "%", lo, v, hi, "Assumed", "FY2026 約 +100%（61.8 萬 vs 30.3 萬）；2026 上半年 +510%、Q3 年增放緩；割草機、配送、人形需求", "線模型未涵蓋割草機與配送，無線變數可讀；共用零組件線 2030 全部 LiDAR 需求基準 1,749 萬顆（五線歸因）供 Recon 對照")
inp("fade_r", "機器人出貨年增每年衰減係數（FY2028 起 ×本值）", "x", 0.55, 0.75, 0.90, "Assumed", "基數放大後增速遞減；高＝增速衰減慢", "")

sec("C. 毛利與費用（比率以最近實績為起點）")
inp("gm0", "毛利率起點（FY2026）", "%", 0.19, 0.22, 0.25, "Assumed", "上半年 21.8%、Q2 21.9%（IFRS）；FY2025 26.5%、Q4'25 28.5%、TTM 24.5%（" + IR26 + "、" + FY25 + "）", "下半年機器人占比提高（機器人毛利率較高）與 ADAS 單價下降、原料漲價互抵；低／高 Assumed")
for s, v, lo, hi in (("A", 0.015, None, None), ("B", 0.005, -0.005, 0.015), ("C", -0.010, None, None)):
    inp(f"gmd_{s}", f"情境 {s}：FY2027 起毛利率每年變化", "pt/年", lo, v, hi, "Assumed", "A：自研 SPAD-SoC 晶片降成本、機器人占比；C：價格戰與原料", "基準小幅改善（0.5pt）：晶片自製成本節省 vs 單價年降 10%（原則第 10 節第 35 條：證據不足不給大幅收斂）")
inp("e_ox", "營業費用（研發、銷售、行政）對營收成長的彈性", "x", 0.20, "=((646.674+128.073+168.021)/(615.43+274.24)-1)/(1941.008/1649-1)", 0.55, "Derived",
    "FY2024→FY2025：營業費用 942.8 ÷ 889.7＝+6.0%，營收 +17.7%（" + FY25 + "；" + SAIS + "）", "上半年費用 +13.6%（514.0 vs 443.9）÷ 營收 +30.2%＝0.45；低／高 Assumed")
inp("oi26", "FY2026 其他收入淨額（其他收入、其他損益、財務、聯營；公司列在營業虧損內的其他項目）", "RMB m", 100.0, 170.0, 200.0, "Assumed",
    IR26 + "：上半年 其他收入 45.7＋其他損益 73.5（含出售聯營部分股權利得、匯兌收益）＋財務淨收入 19.9＋聯營 −8.1＋減值迴轉 0.8＝131.8", "下半年假設 +38（不含一次性利得）")
inp("oi_lt", "FY2027 起其他收入淨額（其他收入、利息、聯營；不預測匯兌與處分利得）", "RMB m", 30.0, 90.0, 150.0, "Assumed", "上半年其他收入 45.7（利息、增值稅退稅、補助）＋財務淨收入 19.9（含利息收入，現金下降中）", "FY2025 業外合計 +222（資料商：稅前 −125.7 減營業利益 −347.7），含一次性項目")
inp("sh_26", "FY2026 平均股數", "百萬股", None, 467.5, None, "Derived", IR26 + "：上半年加權 466.67 百萬股（不含未歸屬限制性股份）", "下半年略增")
inp("tax_w", "所得稅率（只對正的稅前利益）", "%", 0.10, 0.15, 0.25, "Analogy", "中國高新技術企業 15%（類比法定優惠稅率）；累計虧損與研發加計扣除未建模", "FY2025 所得稅 19.3（資料商）；上半年 2026 稅利益 0.04")
inp("pay_r", "配息率", "%", None, 0.0, None, "Interested-party", FY25 + "：未見派息；公司虧損", "只用於收支關卡", CO)

sec("D. 估值輸入（評價日 2026-12-31；資產負債表 2026-06-30）")
inp("px", "股價（2026-11-20 收盤，港元）", "HK$", None, 13.32, None, "Verified", YCL + "：2498.HK 2026-11-20 收盤 13.32", "11-11 收盤 16.47 → 11-20 13.32（−19.1%，同檔；全球 AI 與半導體股同週大跌，非資料錯誤）")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("fx_hk", "美元兌港元（2026-11-20）", "HKD/USD", None, 7.8426, None, "Verified", "Yahoo Finance HKD=X 2026-11-20（physical_ai/data/closes2_20261109_20261120.csv）", "")
inp("fx_cn", "美元兌人民幣（2026-11-20）", "CNY/USD", None, 6.6665, None, "Verified", "Yahoo Finance CNY=X 2026-11-20（同檔）", "")
inp("hkcny", "港元 ÷ 人民幣（每 1 人民幣兌港元）", "HKD/CNY", None, "=fx_hk/fx_cn", None, "Derived", "＝HKD=X ÷ CNY=X", "目標價＝人民幣每股價值 × 本列")
inp("sh_v", "評價日股數", "百萬股", 466.7, 471.24, None, "Derived", SAST + "：471.24 百萬股", "低＝上半年加權基本股數 466.67（公司）；庫藏股 12.48 百萬股（2026-06-30，公司公告）是否已含在內未核對；RSU／選擇權未取得")
inp("bb_post", "6-30 後買回付出現金（未取得）", "RMB m", 0.0, 0.0, None, "Assumed", "未取得 6-30 後買回金額；上半年僅買回 21.5 萬股（公司公告）", "")
inp("cash_v", "現金及流動金融資產（2026-06-30）＝現金 940.3＋定存 1,109.1＋受限現金 8.0＋攤銷成本金融資產 442.9＋公允價值計入損益（流動）149.7", "RMB m", 2057.3, "=2057.3+442.855+149.689", None, "Verified",
    IR26 + "：公司口徑現金、定存、受限款合計 2,057.3；加攤銷成本金融資產（結構性存款等，性質假設為流動）與流動 FVTPL", "低＝公司口徑 2,057.3；不含非流動定存 20.4 與非流動 FVTPL 191.9")
inp("debt_v", "銀行借款（2026-06-30；流動 371.0＋非流動 154.4）", "RMB m", None, 525.4, None, "Verified", IR26, "")
inp("lease_v", "租賃負債（2026-06-30）", "RMB m", None, 144.2, None, "Verified", IR26, "上半年因新增租賃由 31.3 增至 144.2")
inp("h1fcf", "2026 上半年自由現金流（營業現金流 − 資本支出）", "RMB m", None, "=-419.7-172.2", None, "Derived", IR26 + "：營業現金流 −419.7；資本支出以物業廠房淨額增加 172.2（464.9 − 292.7）代，未取得現金流量表資本支出", "Stub（下半年）FCFF＝FY2026 模型 FCFF − 本列")
inp("rf", "無風險利率（人民幣 10 年期國債）", "%", None, 0.0161, None, "Verified", CNRF, "")
inp("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Interested-party", DAMO, "", "Damodaran（學者整理；評等機構違約利差推得）")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("beta", "β（同業 4 家中位數＋本公司，Blume 調整）", "x", "=E{r}-0.3", 1.60, "=E{r}+0.3", "Analogy",
    "資料商顯示本公司 5 年 β 0.24（" + SAST + "，上市未滿 2 年，不採用）；類比上市虧損硬體同業：極智嘉模型採 1.70（Blume，同業 5 家）", "取 1.60（較極智嘉略低，因毛利與現金部位）；第二來源未取得；敏感度 ±0.3（Assumed）")
inp("kd", "稅前負債成本", "%", None, "=rf+0.01", None, "Assumed", "rf ＋ 1pt", "")
inp("g_cap", "永續成長率上限", "%", 0.01, 0.03, 0.04, "Assumed", "DCF 標準規格：3%，且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("fc_da", "折舊攤銷 ÷ 營收", "%", 0.03, 0.045, 0.07, "Assumed", "物業廠房淨額 292.7→464.9（上半年新增 172.2）、使用權資產 28.9→135.7；折舊攤銷金額未取得", "上限含新廠投產後的折舊")
inp("fc_cx", "資本支出 ÷ 營收", "%", 0.05, 0.08, 0.14, "Assumed", "上半年物業廠房淨額增加 172.2 ÷ 營收 1,019.7＝17%（擴廠高峰）；穩態假設 8%", "FY2025 資本支出未取得")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.15, 0.30, 0.40, "Derived", "（應收＋存貨 − 應付 − 合約負債）÷ 營收：FY2025 (800.3＋395.2−545.7−20.5)＝629.3 ÷ 1,941.0＝32%；2026-06-30 686.8（" + IR26 + "）", "以水準比例代增量比例（FY2024 資產負債表未取得）")
inp("dps_tgt", "評價日到目標時點之每股股利", "RMB", None, 0.0, None, "Assumed", "未派息", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable：EV ÷ 營收）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("E. Comparable：同業 EV ÷ 近四季營收（一律以 2026-11-20 收盤重算市值；其餘取資料商頁面，財務更新日 2026-11-24 前後）")
# (鍵, 名稱, 股數×收盤（市值，原幣）, 淨負債（正＝負債，原幣）, TTM 營收, 群組, 毛利率)
PEERS = [("ev_hsai", "禾賽 HSAI（LiDAR；美股 ADS）", "156.44*11.83/1000", -0.89723, 0.5382, "LiDAR", "40.4%"),
         ("ev_oust", "Ouster OUST（LiDAR；美股）", "73.0*23.8/1000", -0.40443, 0.22471, "LiDAR", "46.3%"),
         ("ev_mbly", "Mobileye MBLY（ADAS 晶片與系統；美股）", "850.17*4.57/1000", -1.41, 2.00, "感知晶片與系統", "48.6%"),
         ("ev_hor", "地平線 9660（智駕晶片；港股，港元）", "15870*3.12/1000", -7.59, 4.93, "感知晶片與系統", "64.7%")]
for k, lab, mc, nd, rv, grp, gm in PEERS:
    inp(k, f"EV ÷ 近四季營收：{lab}", "x", None, f"=({mc}+({nd}))/{rv}", None, "Derived",
        f"stockanalysis.com 各股 statistics 頁（資料商；讀取 2026-11-24）：股數、淨現金、近四季營收；× 2026-11-20 收盤（{YCL}）", f"群組 {grp}；毛利率 {gm}；單位為各公司原幣（十億或百萬同比例，比值不受影響）；禾賽財務幣別資料商未標明（顯示為美元），未核對；EV 租賃口徑未逐家核對")
inp("ev_med", "同業 EV ÷ 近四季營收中位數（4 家，進 Comparable）", "x", "=MEDIAN(ev_hsai,ev_mbly)", "=MEDIAN(ev_hsai,ev_oust,ev_mbly,ev_hor)", "=MEDIAN(ev_oust,ev_hor)", "Derived",
    "少於 8 家，不做迴歸（第 10 節第 7 條）；另有奧比中光 688322、Innoviz 等未取得報價日同日資料，不納入", "敏感度低＝禾賽與 Mobileye（兩家倍數最低）中位數；高＝Ouster 與地平線（兩家最高）中位數")
inp("ev_self", "速騰現行 EV ÷ 近四季營收（2026-11-20；只作對照）", "x", None,
    "=(px*sh_v/hkcny-(cash_v-debt_v-lease_v-bb_post))/(1941.008-783.207+1019.749)", None, "Derived", "市值（現價 × 股數 ÷ 匯率）− 淨現金，÷ TTM 營收 2,177.55（FY2025 − 上半年 2025 ＋ 上半年 2026）", "")
inp("rel_k", "相對倍數調整係數（速騰 TTM 毛利率 ÷ 同業毛利率中位數）", "x", 0.35, "=0.245/MEDIAN(0.404,0.463,0.486,0.647)", 1.0, "Assumed",
    "EV ÷ 營收隨毛利率縮放：速騰 TTM 毛利率 24.5% ÷ 同業中位數 47.5%（資料商）＝0.52", "低＝速騰現行相對倍數（約 0.35；建置時寫入常數）；高＝1.00（不調整）")


def build(path, ov=None, sens=None, rev=None):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額百萬人民幣、股數百萬股、出貨萬顆、單價人民幣元／顆、EPS 人民幣元、股價港元"
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
    FAM = [("u_a_q4", "2026 Q4 ADAS 出貨（萬顆）"), ("u_r_q4", "2026 Q4 機器人等出貨（萬顆）"), ("dasp26", "FY2026 ADAS 單價相對上半年"), ("dasp", "ADAS 單價年變化"),
           ("dasp_r", "機器人單價年變化"), ("noa26", "中國城市 NOA 新車數 2026（萬台，線 2）"), ("pen26", "中國城市 NOA 滲透率 2026（線 2）"), ("pen27", "中國城市 NOA 滲透率 2027（線 2）"), ("pen28", "中國城市 NOA 滲透率 2028（線 2）"), ("pen30", "中國城市 NOA 滲透率 2030（線 2）"),
           ("pen32", "中國城市 NOA 滲透率 2032（線 2）"), ("shr", "每台 NOA 車對應速騰 ADAS 出貨年變化"), ("gr", "機器人出貨年增 FY2027"), ("goth", "解決方案與其他營收年增"), ("gmd", "毛利率每年變化")]
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
    H["A1"] = "Hist — 歷史財務（年度；IFRS；百萬人民幣）。公式只引用本頁儲存格（定義名稱 h_*）。FY2023–FY2024 取資料商（年報原文未逐項核對）；FY2025 營收、毛利、費用分項取公司年度業績報導"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "標記", "來源"])
    HR = [("h_rev", "營收", 1120.0, 1649.0, 1941.008), ("h_cogs", "營業成本", "=h_rev_23-h_gp_23", "=h_rev_24-h_gp_24", "=h_rev_25-h_gp_25"),
          ("h_gp", "毛利", 93.64, 283.55, 514.152), ("h_rd", "研發費用", 635.11, 615.43, 646.674),
          ("h_sga", "銷售＋行政（FY2023–FY2024 只有合計）", 431.95, 274.24, "=h_sm_25+h_ga_25"), ("h_sm", "銷售及營銷費用", None, None, 128.073), ("h_ga", "行政費用", None, None, 168.021),
          ("h_pbt", "稅前利益", -4329.0, -476.83, -125.67), ("h_tax", "所得稅", 1.87, 4.98, 19.30), ("h_ni", "年度淨利（稅前 − 所得稅）", "=h_pbt_23-h_tax_23", "=h_pbt_24-h_tax_24", "=h_pbt_25-h_tax_25"),
          ("h_eps", "每股盈餘（資料商，稀釋；FY2023 為上市前股數）", -44.67, -1.11, -0.32), ("h_sh", "股數（資料商，稀釋，百萬股）", 97.0, 435.0, 462.0),
          ("h_au", "ADAS LiDAR 出貨（萬顆）", None, None, 60.9), ("h_ru", "機器人及其他 LiDAR 出貨（萬顆）", None, None, 30.3),
          ("h_rec", "應收帳款及票據", None, None, 800.267), ("h_inv", "存貨", None, None, 395.171), ("h_ap", "應付帳款", None, None, 545.673), ("h_cl", "合約負債", None, None, 20.535)]
    for i, (k, lab, a, b, c) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b, c)):
            if v is None: H.cell(i, 2 + j, "未取得"); continue
            H.cell(i, 2 + j, v).font = BLUE
        H.cell(i, 5, "Verified（FY2025）／Interested-party（資料商）")
        H.cell(i, 6, SAIS + "；" + FY25 + "；" + IR26 + "（2025-12-31 資產負債表欄）")
        for j, y in enumerate(("23", "24", "25")): dn(f"{k}_{y}", f"Hist!${'BCD'[j]}${i}")
    i = 3 + len(HR)
    H.cell(i, 1, "營運資金（應收＋存貨 − 應付 − 合約負債；Derived；FY2025）"); H.cell(i, 4, "=h_rec_25+h_inv_25-h_ap_25-h_cl_25"); dn("h_nwc_25", f"Hist!$D${i}")
    H.cell(i + 1, 1, "檢查：FY2025 銷售＋行政 − 資料商合計 296.09（應為 0）"); H.cell(i + 1, 4, "=ROUND(h_sga_25-296.09,2)")
    H.cell(i + 2, 1, "稅前與淨利：資料商 FY2023 淨利 −4,337（歸屬）與稅前 − 稅 −4,331 差 6；FY2023 的巨額虧損主因為上市前可贖回優先股公允價值變動（資料商業外 −3,471），不具持續性，不用於推估")
    H.cell(i + 3, 1, "FY2025 業績報導：淨虧損 144.97、經調整淨虧損 53.52、毛利率 26.5%（FY2024 17.2%）；Q4 單季淨利 104（含一次性項目，細項未取得）")
    H.column_dimensions["A"].width = 52

    # ---------------- Q（半年度） ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 半年度營收分類、出貨與損益（百萬人民幣）。H1 為公告原文（Verified）；公司以產品／解決方案／服務與其他分類營收（單一營運分部未核對）；ADAS 與機器人營收公司只稱「機器人近一半」，拆分為本模型推導（Derived）"
    hdr(Q, 2, ["項目", "H1'25", "H1'26", "標記", "來源", "來源誘因"])
    QV = [("產品營收", 721.026, 958.003, "Verified"), ("解決方案營收", 51.174, 17.220, "Verified"), ("服務與其他營收", 11.007, 44.526, "Verified"), ("營收", 783.207, 1019.749, "Verified"),
          ("檢查：分類合計 − 營收（應為 0）", "=ROUND(SUM(B3:B5)-B6,3)", "=ROUND(SUM(C3:C5)-C6,3)", "Derived"),
          ("毛利", 203.136, 222.189, "Verified"), ("毛利率", "=B8/B6", "=C8/C6", "Derived"), ("研發費用", 308.712, 346.452, "Verified"), ("銷售及營銷費用", 56.894, 73.827, "Verified"),
          ("行政費用", 78.317, 93.686, "Verified"), ("營業利益（本模型口徑：毛利 − 三項費用）", "=B8-SUM(B10:B12)", "=C8-SUM(C10:C12)", "Derived"),
          ("其他收入淨額（其他收入、其他損益、財務、聯營、減值）", "=B15-B13", "=C15-C13", "Derived"), ("稅前利益", -146.420, -159.982, "Verified"), ("淨利", -148.606, -159.947, "Verified"),
          ("ADAS LiDAR 出貨（萬顆）", "=43.66/(1+0.980)", "=u_a_h1", "Derived"), ("機器人等 LiDAR 出貨（萬顆）", "=28.26/(1+5.104)", "=u_r_h1", "Derived"),
          ("LiDAR 出貨合計（萬顆）", "=B17+B18", "=C17+C18", "Derived"),
          ("機器人營收（推導＝產品營收 × 機器人比重）", None, "=C3*rho", "Derived"), ("ADAS 營收（推導）", None, "=C3-C20", "Derived")]
    for i, (lab, a, b, tg) in enumerate(QV, 3):
        Q.cell(i, 1, lab)
        for j, v in enumerate((a, b)):
            if v is None: Q.cell(i, 2 + j, "未取得"); continue
            cc = Q.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): cc.font = BLUE
        Q.cell(i, 4, tg); Q.cell(i, 5, IR26 + ("；H1'25 出貨由公司公告年增率反推；機器人營收＝產品營收 × 機器人比重（H1'25 無資料）" if tg == "Derived" else "")); Q.cell(i, 6, CO)
    for c in "BC": Q[f"{c}9"].number_format = "0.0%"
    Q.cell(3 + len(QV) + 1, 1, "註：Q2 2026 營收 561.0（+23.2%）、毛利 122.6（毛利率 21.9%）、淨虧損 96.6；ADAS Q2 出貨 29.18 萬（+135.7%）；2026 Q3 出貨 59.56 萬顆（" + GAS + "）；上半年營運現金流 −419.7（" + IR26 + "）")
    Q.column_dimensions["A"].width = 52

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（中國城市 NOA 新車 × 每台對應速騰 ADAS 出貨）→ 分部營收（ADAS、機器人、其他）→ 標準損益表（IFRS 口徑）→ 現金流（百萬人民幣；FY2023–FY2025 實績、FY2026 起模型）"
    PR = {}
    rr = 3
    KEYS = ["mkt", "pen", "kf", "au", "asp_a", "ru", "asp_r", "adas", "rob", "oth", "rev", "chk", "gr", "cogs", "gp", "gm", "rd", "sm", "ga", "ox", "op", "opm", "oi", "pbt", "tax", "ni", "nm", "nig",
            "sh", "eps", "da", "cx", "dwc", "fcff", "div", "fcfad"]
    LABS = ["命題模組：中國城市 NOA 新車數（萬台；線 2 下游取數）", "中國城市 NOA 滲透率（線 2 下游取數）", "每台 NOA 車對應速騰 ADAS 出貨（顆／台；含搭載率、每車顆數、份額）", "ADAS LiDAR 出貨（萬顆）",
            "ADAS LiDAR 平均單價（人民幣元／顆）", "機器人及其他 LiDAR 出貨（萬顆）", "機器人等 LiDAR 平均單價（人民幣元／顆）", "ADAS LiDAR 營收", "機器人等 LiDAR 營收", "解決方案、服務與其他營收",
            "營收（合併）", "檢查：分部合計 − 營收（應為 0）", "營收年增率", "營業成本", "毛利", "毛利率", "研發費用", "銷售及營銷費用", "行政費用", "營業費用（研發＋銷售＋行政）",
            "營業利益（本模型口徑：毛利 − 營業費用）", "營業利益率", "其他收入淨額（業外：其他收入、其他損益、財務、聯營、減值；FY2023 含優先股公允價值變動）",
            "稅前利益", "所得稅", "稅後淨利", "淨利率", "淨利年增率", "平均股數（百萬股）", "EPS（基本，人民幣元）",
            "折舊攤銷", "資本支出", "營運資金增加", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "股利", "自由現金流扣股利（收支關卡用）"]
    PY = {2027: "pen27", 2028: "pen28", 2030: "pen30", 2032: "pen32"}
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
                for k in ("mkt", "pen", "kf", "asp_a", "asp_r", "adas", "rob", "oth", "da", "cx", "dwc", "div", "fcff", "fcfad"): w(k, "未取得")
                w("au", "=h_au_25" if y == 2025 else "未取得"); w("ru", "=h_ru_25" if y == 2025 else "未取得")
                w("chk", "未揭露（分部營收未取得）")
                w("rev", f"=h_rev_{yy}"); w("cogs", f"=h_cogs_{yy}"); w("gp", f"=h_gp_{yy}"); w("rd", f"=h_rd_{yy}")
                if y == 2025: w("sm", "=h_sm_25"); w("ga", "=h_ga_25")
                else: w("sm", "未取得（只有銷售＋行政合計）"); w("ga", "未取得（只有銷售＋行政合計）")
                w("ox", f"={cell('rd')}+h_sga_{yy}")
                w("op", f"={cell('gp')}-{cell('ox')}"); w("pbt", f"=h_pbt_{yy}"); w("oi", f"={cell('pbt')}-{cell('op')}")
                w("tax", f"=h_tax_{yy}"); w("ni", f"=h_ni_{yy}"); w("sh", f"=h_sh_{yy}"); w("eps", f"=h_eps_{yy}")
                w("gm", f"={cell('gp')}/{cell('rev')}")
                if y >= 2024: w("gr", f"={cell('rev')}/{prev('rev')}-1"); w("nig", "N/M（前一年虧損）")
            else:
                if y == 2026:
                    w("mkt", f"=noa26_{s}"); w("pen", f"=pen26_{s}")
                    w("au", f"=u_a_h1+u_a_q3+u_a_q4_{s}"); w("kf", f"={cell('au')}/{cell('mkt')}")
                    w("asp_a", f"=asp_a0*(1+dasp26_{s})"); w("ru", f"=u_r_h1+u_r_q3+u_r_q4_{s}"); w("asp_r", f"=asp_r0*(1+dasp_r_{s}/2)")
                    w("oth", "=oth26"); w("oi", "=oi26"); w("sh", "=sh_26"); w("gm", "=gm0")
                    w("rd", f"={prev('rd')}*(1+e_ox*{cell('gr')})")
                    w("sm", f"={prev('sm')}*(1+e_ox*{cell('gr')})"); w("ga", f"={prev('ga')}*(1+e_ox*{cell('gr')})")
                else:
                    if y in PY: w("pen", f"=pen{y % 100}_{s}")
                    else: w("pen", f"=SQRT(pen{(y - 1) % 100}_{s}*pen{(y + 1) % 100}_{s})")
                    w("mkt", f"={cell('pen')}*noa26_{s}/pen26_{s}")
                    w("kf", f"={prev('kf')}*(1+shr_{s})"); w("au", f"={cell('mkt')}*{cell('kf')}")
                    w("asp_a", f"={prev('asp_a')}*(1+dasp_{s})")
                    w("ru", f"={prev('ru')}*(1+gr_{s}*fade_r^{y - 2027})"); w("asp_r", f"={prev('asp_r')}*(1+dasp_r_{s})")
                    w("oth", f"={prev('oth')}*(1+goth_{s})"); w("oi", "=oi_lt"); w("sh", "=sh_v"); w("gm", f"={prev('gm')}+gmd_{s}")
                    w("rd", f"={prev('rd')}*(1+e_ox*{cell('gr')})")
                    w("sm", f"={prev('sm')}*(1+e_ox*{cell('gr')})"); w("ga", f"={prev('ga')}*(1+e_ox*{cell('gr')})")
                w("adas", f"={cell('au')}*{cell('asp_a')}/100"); w("rob", f"={cell('ru')}*{cell('asp_r')}/100")
                w("rev", f"={cell('adas')}+{cell('rob')}+{cell('oth')}")
                w("chk", f"=ROUND({cell('adas')}+{cell('rob')}+{cell('oth')}-{cell('rev')},6)")
                w("gr", f"={cell('rev')}/{prev('rev')}-1")
                w("cogs", f"={cell('rev')}*(1-{cell('gm')})"); w("gp", f"={cell('rev')}-{cell('cogs')}")
                w("ox", f"={cell('rd')}+{cell('sm')}+{cell('ga')}")
                w("op", f"={cell('gp')}-{cell('ox')}"); w("pbt", f"={cell('op')}+{cell('oi')}"); w("tax", f"=MAX(0,{cell('pbt')})*tax_w"); w("ni", f"={cell('pbt')}-{cell('tax')}")
                w("nig", f'=IF({prev("ni")}>0,{cell("ni")}/{prev("ni")}-1,"N/M（前一年虧損）")')
                w("eps", f"={cell('ni')}/{cell('sh')}")
                w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}")
                w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
                w("fcff", f"={cell('op')}-MAX(0,{cell('op')})*tax_w+{cell('da')}-{cell('cx')}-{cell('dwc')}")
                w("div", f"=pay_r*MAX(0,{cell('ni')})"); w("fcfad", f"={cell('fcff')}-{cell('div')}")
            w("opm", f"={cell('op')}/{cell('rev')}"); w("nm", f"={cell('ni')}/{cell('rev')}")
            if y == 2023: w("gr", "未取得（FY2022 未取得）"); w("nig", "N/M（前一年虧損）")
        P.cell(R["mkt"], 2, "線 2 下游取數：滲透率 × 國內零售（零售＝2026 新車數 ÷ 2026 滲透率，線模型為常數）")
        P.cell(R["pen"], 2, "2026、2027、2028、2030、2032 讀線 2；2029、2031 取相鄰兩年幾何平均")
        P.cell(R["kf"], 2, "FY2026＝ADAS 出貨 ÷ NOA 新車（Derived）；FY2027 起 ×(1＋shr)")
        P.cell(R["au"], 2, "FY2026＝上半年實績＋Q3 公告＋Q4 假設；FY2027 起＝NOA 新車 × 每台出貨")
        P.cell(R["ru"], 2, "FY2026＝上半年實績＋Q3 公告＋Q4 假設；之後年增 gr × 衰減^(年−2027)")
        P.cell(R["rd"], 2, "前一年 ×(1＋彈性 e_ox × 營收年增)")
        P.cell(R["op"], 2, "公司「營業虧損」另含其他收入與其他損益（見 Recon）")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 − MAX(0, 營業利益) × tax_w；股份支付不加回")
        for k in ("gr", "gm", "opm", "nm", "nig", "pen"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(LABS) + 2
    P.column_dimensions["A"].width = 52; P.column_dimensions["B"].width = 30
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "adas", "rob", "fcfad"):
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
    D["A12"] = "檢查：淨現金 − 資產負債表重算（應為 0）"; D["B12"] = "=ROUND(dcf_nc-(cash_v-525.4-144.2-bb_post),6)"
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
    V["A1"] = "Val — 目標價（港元）：DCF（權重見 In 頁，基準一半）＋ Comparable（EV ÷ 營收；其餘權重）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
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
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"=(ev_med*rel_k*{cf(s,'rev')}+dcf_nc+DCF!B{DR[s]['stub']})/sh_v*hkcny" for s in SC], "rel_k＝毛利率相對縮放（約 0.52；Assumed）")
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
    r_evs = one("現行 EV ÷ 近四季營收：速騰聚創", "=ev_self", "分母：TTM 營收 2,177.5（2025 下半年＋2026 上半年）", "0.00")
    r_evm = one("同業 EV ÷ 近四季營收中位數（4 家）", "=ev_med", "", "0.00")
    r_rel = one("現行相對倍數（速騰聚創 ÷ 同業中位數）", "=ev_self/ev_med", "上市未滿 2 年，無 10 年比值；調整後 Comparable 需要的相對重估倍數＝rel_k ÷ 本列", "0.00")
    r_sg1 = one("子群中位數（純展示）：LiDAR（禾賽、Ouster）", "=MEDIAN(ev_hsai,ev_oust)", "第 10 節第 18 條", "0.00")
    r_sg2 = one("子群中位數（純展示）：感知晶片與系統（Mobileye、地平線）", "=MEDIAN(ev_mbly,ev_hor)", "", "0.00")
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
    RC["A1"] = "Recon — 與公司實績、線模型同口徑對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或對照", "差", "說明"])
    B_ = PR["B"]
    rc = [("FY2026 下半年隱含營收 vs 2026 上半年", f"=PL!F{B_['rev']}-Q!C6", "=Q!C6", "對照＝上半年營收 1,019.7；下半年（Q3 出貨 59.56 萬顆、Q4 假設）預期高於上半年"),
          ("FY2026 營收年增（模型）vs 上半年年增 30.2%", f"=PL!F{B_['gr']}", "=Q!C6/Q!B6-1", "上半年營收年增 30.2%；全年高於上半年反映下半年出貨放量（LiDAR 出貨上半年 +169.6%、Q3 +220.9%）"),
          ("FY2026 全年 LiDAR 出貨（萬顆）vs 2025 全年 91.2", f"=PL!F{B_['au']}+PL!F{B_['ru']}", "=h_au_25+h_ru_25", "2026 前三季 ADAS 88.7 萬＋機器人 42.78 萬＝131.5 萬顆；已超過 2025 全年"),
          ("FY2026 營業利益（本模型口徑）vs 上半年實績 ×2", f"=PL!F{B_['op']}", "=Q!C13*2", "上半年本模型口徑營業利益 −291.8（毛利 222.2 − 費用 514.0）；下半年毛利隨出貨放量增加"),
          ("FY2026 稅前利益（模型）vs 上半年實績 ×2", f"=PL!F{B_['pbt']}", "=Q!C15*2", "上半年稅前 −160.0 含一次性處分聯營利得約 73"),
          ("上半年公司營業虧損 −171.8 ＝ 本模型營業利益 ＋ 其他收入 45.7 ＋ 其他損益 73.5 ＋ 減值迴轉 0.8（應為 0）", "=ROUND(Q!C13+45.691+73.463+0.848-(-171.774),3)", 0, "公司公告營業虧損 171.774（2026 上半年）"),
          ("2030 ADAS＋機器人 LiDAR 出貨（萬顆）vs 線 6 共用零組件 2030 全部 LiDAR 需求基準 1,749 萬顆（五線歸因，不含割草機）", f"=PL!I{B_['au']}+PL!I{B_['ru']}", 1749.0, f"共用零組件單位經濟 下游取數!C21:E21（716／1,749／3,983.5 萬顆）；速騰占比見差欄，口徑不同（本模型含割草機與全球車用）"),
          ("FY2026 ADAS 出貨 ÷ 線 2 NOA 新車（顆／台；含搭載率 65%、每車 1.0 顆）", f"=PL!F{B_['kf']}", f"={LATT}*{LNUM}", "線模型搭載率 × 每車顆數＝0.65；本模型係數>0.65 表示速騰出貨含非 NOA 車型、出口與庫存，口徑較寬（Derived，不調整線值）")]
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
        SE.cell(sr + 1, 1, "說明：折舊、資本支出、營運資金隨營收等比例（第 44 條）；相對倍數係數 rel_k 基準約 0.52，低案為現行相對倍數，高案為 1.00。出貨量與單價以情境 A／B／C 表達，Sens 中各情境專屬驅動只動情境 B。pay_r 等無區間的鍵不列案例。")
    SE.column_dimensions["A"].width = 52; SE.column_dimensions["B"].width = 40


    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = f"對帳 — 本模型使用的線層級變數（讀 {LINE_REL}「下游取數」頁；與線模型不一致的自設值標明理由）"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（低／基準／高）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型是否一致"])
    def lv(k, y): a, b, c = lv3(k, y); return f"{a:.3f}／{b:.3f}／{c:.3f}"
    rec = [(f"中國城市 NOA 滲透率 {y}", "2 智慧車", lv(PEN, y), "%", "國內乘用車零售口徑", lref(PEN, y, "C") + "；" + f"下游取數!D{LINE[f'{PEN}|{y}']['row']}:E{LINE[f'{PEN}|{y}']['row']}", "Derived", f"In pen{y % 100}_A／B／C → PL 命題模組（NOA 新車數）", "一致（A＝高、B＝基準、C＝低）") for y in (2026, 2027, 2028, 2030, 2032)]
    rec += [("中國城市 NOA 新車數 2026（萬台）", "2 智慧車", lv(NOA, 2026), "萬台", "滲透率 × 國內零售", lref(NOA, 2026, "C"), "Derived", "In noa26_A／B／C → PL FY2026 起各年 NOA 新車數（零售＝本值 ÷ 2026 滲透率）", "一致"),
            ("中國城市 NOA 新車數 2030（萬台，只作對照）", "2 智慧車", lv(NOA, 2030), "萬台", "同上", lref(NOA, 2030, "C"), "Derived", "Recon／對照：本模型 2030 NOA＝滲透率 × 零售，應與本列相同", "一致"),
            ("城市 NOA 車搭載 LiDAR 比例", "2 智慧車", f"0.50／{LATT}／0.80", "%", "線模型 Assumed", f"{LINE_REL} 輸入!D22:F22", "Assumed", "In lidar_att（只作對照）→ Recon 係數對照列", "不一致（刻意）：本模型把搭載率、每車顆數、速騰份額合併為 PL「每台 NOA 車對應速騰 ADAS 出貨」係數，由 2026 出貨反推，避免與線模型重複計算；不是線值被改"),
            ("每車 LiDAR 顆數", "2 智慧車", f"1.0／{LNUM}／1.2", "顆", "線模型 Assumed", f"{LINE_REL} 輸入!D23:F23", "Assumed", "In lidar_n（只作對照）", "同上"),
            ("ADAS LiDAR 單價（美元）", "2 智慧車", f"150／{LASP}／300", "USD/顆", "類比速騰 MX、禾賽 ATX", f"{LINE_REL} 輸入!D24:F24", "Analogy", "不進公式；對照速騰推導單價 asp_a0（約 1,160 人民幣）", "不一致（刻意）：本模型單價由公司產品營收與出貨推導（Derived）；200 美元約 1,330 人民幣，高於推導值約 15%，因 H1 含低價與出口混合；不靜默改線值"),
            ("LiDAR 市場規模 2030／需求量 2030", "6 共用零組件", "850／2,141／5,059 百萬美元；716／1,749／3,984 萬顆", "百萬美元；萬顆", "五線歸因（不含割草機）", "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx 下游取數!C20:E21", "Derived", "Recon 第 7 列（只作對照）；機器人出貨路徑無線變數可讀", "口徑不同：線 6 不含割草機等；僅對照量級"),
            ("線 1 Robotaxi", "1 Robotaxi", "未讀", "—", "線 1 下游取數無 LiDAR 欄位", "physical_ai/lines/robotaxi/dist/20261116_Robotaxi單位經濟_v0.2.xlsx", "—", "無；速騰 Robotaxi 營收未揭露，不建模", "—")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 34; RE.column_dimensions["F"].width = 60; RE.column_dimensions["I"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("車用 LiDAR 單價持續下降（同業 ASP 年減約 35%）", "是（dasp_*、Sens）", "基準年降 10%；情境 C 年降 16%", "半年報毛利率、每顆單價"),
             ("中國城市 NOA 滲透率不如預期", "是（線 2 低／高；Sens 的 pen*）", "讀線 2 下游取數", "城市 NOA 車銷量與 LiDAR 搭載比例"),
             ("份額流失（禾賽、華為、圖達通）", "是（shr_*）", "情境 C 每年 −8%", "季度 ADAS 出貨相對禾賽"),
             ("客戶集中度", "否", "財報揭露的前五大客戶占比未找到（查過 2026 中期公告、2025 年度業績新聞稿）；比亞迪、吉利等為公司新聞稿所列客戶", "年報主要客戶附註"),
             ("機器人出貨放量能否持續（Q3 低於 Q2）", "是（gr_*、fade_r）", "Q3 機器人出貨 14.52 萬，與上半年季均 14.13 萬（28.26 ÷ 2）約持平，Q2 起未再加速；原因未找到", "Q4 與 2027 上半年機器人出貨"),
             ("擴廠資本支出與折舊（物業廠房 6 個月增 59%）", "部分（fc_cx、fc_da）", "上半年資本支出約 172；現金流量表資本支出未取得", "年報現金流量表"),
             ("原料與晶片成本上升（AI 帶動零件缺貨）", "是（gm0、gmd_C）", "公司稱上半年毛利率下降部分來自原料成本上升", "毛利率"),
             ("RSU／選擇權稀釋與庫藏股處理", "否", "RSU 與選擇權數量未取得；庫藏股 12.48 百萬股是否含於股數未核對", "年報股份激勵附註"),
             ("出售聯營股權的一次性利得（上半年約 73）", "否", "不預測處分利得", "下半年其他損益")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 54; RK.column_dimensions["C"].width = 60

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([IR26, FY25, GAS, SAIS, SAST, "stockanalysis.com HSAI、OUST、MBLY、9660 statistics 頁（讀取 2026-11-24）", YCL, CNRF, DAMO,
                            LINE_REL + "（下游取數、輸入頁）", "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx（下游取數）"], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f'="速騰聚創 RoboSense（2498.HK）個股模型 {VER}｜build_id "&Sens!D2&"｜資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'
    RD["A2"] = '="資料截止日 2026-11-24（財報為 2026 中期業績，2026-08-26；Q3 出貨公告 2026-10-09）；成品日 2026-11-24；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.00")&" 港元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：速騰聚創的股價由車用 ADAS LiDAR 出貨量（中國城市 NOA 新車數 × 每台對應出貨）與單價年降的賽跑、以及機器人 LiDAR 出貨能否把營收規模推到攤平固定研發費用決定——ADAS 營收＝中國城市 NOA 新車（線 2）× 每台對應出貨 × 單價；機器人營收＝出貨成長路徑 × 單價；其他以成長路徑建模"
    RD["A4"] = (f'="在情境 B FY2026 營收 "&TEXT(PL!F{PR["B"]["rev"]},"#,##0")&" 百萬人民幣（年增 "&TEXT(PL!F{PR["B"]["gr"]},"0%")&"）、ADAS 單價年降 "&TEXT(-dasp_B,"0%")&"、每台 NOA 車對應出貨年變化 "&TEXT(shr_B,"0%")&"、毛利率 "&TEXT(gm0,"0.0%")&" 起每年 +"&TEXT(gmd_B*100,"0.0")&"pt、營業費用彈性 "&TEXT(e_ox,"0.00")&" 的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 港元（DCF "&TEXT(Val!B{r_dcft},"0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0")&" 百萬人民幣）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 港元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")&"；現行 EV ÷ 營收 "&TEXT(Val!B{r_evs},"0.00")&" 倍、同業中位數 "&TEXT(Val!B{r_evm},"0.00")&" 倍（相對 "&TEXT(Val!B{r_rel},"0.00")&"）"'
    RD["A6"] = "失效條件（可觀察）：ADAS LiDAR 季出貨連兩季年增低於 50%；每顆單價（營收 ÷ 出貨）年降超過 20%；季度毛利率跌破 18%；機器人 LiDAR 季出貨低於 14 萬顆；現金（含定存）跌破 15 億人民幣"
    RD["A7"] = "範圍：不讀 Siliconomics／Tokenomics／Wattonomics（算力線）；ADAS 與機器人營收拆分公司未揭露，由「機器人近一半」推導（Derived）；觸覺、關節模組、空間相機（2026Q4 起量產）無金額，只併入「解決方案與其他」成長路徑；線 1 Robotaxi 無速騰營收揭露，未建模。HTML 展示層見 dist/ 同名 .html"
    wb.save(path)
    return VR, PR, DR

# ---------------- 重算與讀值 ----------------
LO = "file://" + os.path.join(tempfile.gettempdir(), "lo_robosense_profile")
def recalc_many(srcs, outdir):
    os.makedirs(outdir, exist_ok=True)
    for attempt in range(2):
        try:
            subprocess.run(["soffice", f"-env:UserInstallation={LO}", "--headless", "--convert-to", "xlsx", "--outdir", outdir] + srcs,
                           check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=max(180, 6 * len(srcs)))
            break
        except subprocess.TimeoutExpired:
            subprocess.run(["pkill", "-f", "lo_robosense_profile"]);
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
    d = tempfile.mkdtemp(prefix="rbs_"); paths = []; VRs = {}
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

CASES = [("u_a_q4_B", "2026 Q4 ADAS 出貨（情境 B）"), ("u_r_q4_B", "2026 Q4 機器人出貨（情境 B）"), ("rho", "機器人占產品營收比重"), ("dasp26_B", "FY2026 ADAS 單價相對上半年（情境 B）"),
         ("dasp_B", "ADAS 單價年變化（情境 B）"), ("dasp_r_B", "機器人單價年變化（情境 B）"), ("shr_B", "每台 NOA 車對應速騰 ADAS 出貨年變化（情境 B）"), ("gr_B", "機器人出貨年增 FY2027（情境 B）"),
         ("fade_r", "機器人出貨年增衰減係數"), ("oth26", "FY2026 解決方案與其他營收"), ("goth_B", "解決方案與其他年增（情境 B）"), ("gm0", "毛利率起點"), ("gmd_B", "毛利率每年變化（情境 B）"),
         ("e_ox", "營業費用彈性"), ("oi26", "FY2026 其他收入淨額"), ("oi_lt", "FY2027 起其他收入淨額"), ("tax_w", "所得稅率"), ("beta", "β"), ("erp", "ERP"), ("g_cap", "永續成長率上限"),
         ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("fc_da", "折舊 ÷ 營收"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("sh_v", "評價日股數"),
         ("cash_v", "現金及流動金融資產"), ("ev_med", "同業 EV ÷ 營收倍數（子群）"), ("rel_k", "相對倍數係數"), ("w_dcf", "DCF 權重")]
MKK = [f"pen{y % 100}_B" for y in (2026, 2027, 2028, 2030, 2032)]

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
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {r['val'] if isinstance(r['val'], str) else format(r['val'], '.4f')} | {r['base']} | {r['note']} |" for r in revs)
    gate = [g(f"{c}{VR['gate']}") for c in "BCD"]
    sens_by = {c["key"] + ("_lo" if c["drv"].endswith("（低）") else "_hi" if c["drv"].endswith("（高）") else ""): c for c in sens["cases"]}
    def sc(k): c = sens_by.get(k); return f"{f2(c['tp'])}（{pc(c['tp']/sens['base']['tp']-1)}，{c['rt']}）" if c else "—"
    ends = [sens_by[k]["rt"] for k in ("ev_med_lo", "rel_k_lo", "w_C") if k in sens_by]
    base_rt = g(f"B{VR['rt']}")
    order = ["強力放空", "放空", "中立", "買進", "強力買進"]
    lowest = min(ends + [base_rt], key=order.index)
    robust = (f"三個端點案例的估值評等都是{base_rt}（區間穩健性：穩健）。" if all(e == base_rt for e in ends) else
              f"三個端點案例的估值評等為{'、'.join(ends)}，最低為{lowest}；基準評等{base_rt}只在點值成立（區間穩健性：{lowest}以上穩健、{base_rt}不穩健）。")
    tpl = lambda k: g(f"B{VR[k]}")
    txt = f"""# physical_ai/robosense — 速騰聚創 RoboSense（2498.HK）個股模型 {VER}

{RD['A2'].value}｜build_id `{sens['build_id']}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

港交所主板上市；報表人民幣、股價港元。本 README 的數字由 `build_robosense.py` 從重算後的成品讀出產生，不手寫。

**本版沒有做的部分（高不確定性原則：證據到哪裡判斷到哪裡）：** ADAS 與機器人營收公司只稱「機器人近一半」，拆分為推導值；單價未揭露（由產品營收與出貨推導）；財報揭露的客戶集中度未找到；觸覺、關節模組、空間相機沒有金額，只併入「解決方案與其他」；線 1 Robotaxi 無速騰營收揭露，未建模；FY2023–FY2024 只有資料商彙整；現金流量表資本支出、折舊攤銷、RSU／選擇權、6-30 後買回金額未取得；賣方目標價與共識未查，Recon 不含賣方對照；同業只有 4 家有 2026-11-20 同日報價（奧比中光等未納入）。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

**結論的證據邊界：** 加權目標價的上檔主要來自 {'Comparable' if g(f"B{VR['a2']}") > g(f"B{VR['a1']}") else 'DCF'}（Comparable 用同業 EV ÷ 營收中位數 {tpl('evm'):.2f} 倍並按毛利率縮放，Analogy；速騰現行 {tpl('evs'):.2f} 倍）。只看 DCF 為 {f2(dcf)} 港元（{pc(dcf/IV['px']-1)}）。同業取最低兩家倍數：{sc('ev_med_lo')}；維持現行相對折價（係數 {sens['rel_now']:.2f}）：{sc('rel_k_lo')}；只看情境 C：{sc('w_C')}。{robust}

## 估值（評價日 2026-12-31；目標時點約 12 個月；港元）
| 情境 | DCF（推到 12 個月後） | Comparable（EV ÷ 營收） | 情境目標價 |
|---|---|---|---|
{scn}

- 加權目標價 **{f2(tp)} 港元**（DCF {f2(dcf)}、Comparable {f2(cmp_)}），現價 {IV['px']:.2f}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{tpl('rt')}**；收支關卡 {tpl('gj')}（情境 A／B／C FY2026–FY2030 自由現金流扣股利 {f0(gate[0])}／{f0(gate[1])}／{f0(gate[2])} 百萬人民幣）→ 投資建議 **{tpl('rec')}**。情境 B 期間最低年末淨現金 {f0(tpl('lq'))} 百萬人民幣（不含新增舉債）。
- 方法：DCF 標準版 FCFF（WACC／ke／g／RONIC＝{tpl('wk')}）50%＋Comparable 50%（同業 4 家 EV ÷ 近四季營收中位數 × 相對倍數係數 × FY2027 營收＋淨現金）。
- 隱含 EV ÷ FY2027 營收（加權目標價）{tpl('ipe'):.2f} 倍；隱含 P/E（÷ 情境 B FY2028 EPS）{tpl('pe') if isinstance(tpl('pe'), str) else format(tpl('pe'), '.1f')}。

## 營收與損益（情境 B，百萬人民幣；FY2023–FY2025 實績，FY2026 起模型）
| 列 | {" | ".join(f"FY{y}" for y in YRS)} |
|---|{"---|" * len(YRS)}
| 中國城市 NOA 新車（萬台，線 2） | {yr('B','mkt','{:,.0f}')} |
| ADAS LiDAR 出貨（萬顆） | {yr('B','au')} |
| 機器人等 LiDAR 出貨（萬顆） | {yr('B','ru')} |
| ADAS 營收 | {yr('B','adas')} |
| 機器人等營收 | {yr('B','rob')} |
| 營收（合併） | {yr('B','rev')} |
| 毛利率 | {yr('B','gm','{:.1%}')} |
| 營業利益（毛利 − 營業費用） | {yr('B','op')} |
| 稅後淨利 | {yr('B','ni')} |
| EPS（人民幣元） | {yr('B','eps','{:.2f}')} |
| FCFF | {yr('B','fcff')} |

- 命題模組：ADAS 出貨 FY2027 起＝中國城市 NOA 新車（線 2 下游取數；A＝高、B＝基準、C＝低）× 每台對應出貨（FY2026 由出貨反推，之後年變化 shr）；ADAS 營收＝出貨 × 單價（上半年產品營收扣掉機器人近一半的推導值，年降 10%）。機器人出貨為成長路徑（線模型未涵蓋割草機與配送，無線變數可讀）。
- 毛利率起點 22%（上半年 21.8%、FY2025 26.5%），情境 B 每年 +0.5pt；營業費用＝前一年 ×(1＋彈性 × 營收年增)，彈性 {IV['e_ox']:.2f}（FY2024→FY2025，Derived）。
- 營業利益為本模型口徑（毛利 − 研發 − 銷售 − 行政）；公司「營業虧損」另含其他收入與其他損益，見 Recon；FY2025 業外 +303 含一次性項目。

## 主要驅動與標記
- 每台 NOA 車對應速騰 ADAS 出貨（shr、kf）：Assumed（區間 −6%～+2% 每年）；NOA 新車數：Derived（線 2）。
- ADAS 單價年降（dasp）：Assumed（−3%～−18%）；機器人出貨增速（gr、fade_r）：Assumed。
- 結論主要建立在 Assumed 與 Analogy（同業倍數）輸入上；實績（Verified）只錨定 2026 上半年與 FY2025。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| Q4 出貨與 FY2026 營收 | 2026 年度業績（約 2027-03）；Q4 ADAS 出貨是否達 {IV['u_a_q4_B']:.0f} 萬顆 |
| 單價年降（dasp） | 年報產品營收 ÷ 出貨；同業 ASP |
| 每台 NOA 車對應出貨（shr） | 季度 ADAS 出貨相對中國城市 NOA 新車數 |
| 毛利率（gm0、gmd） | 季報毛利率、機器人占比 |
| Comparable 倍數（ev_med，Analogy） | 同業 EV ÷ 營收；速騰自身倍數是否向同業收斂 |

失效條件：{RD['A6'].value.replace('失效條件（可觀察）：', '')}。

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
讀 `{LINE_REL}`「下游取數」：中國城市 NOA 滲透率 2026／2027／2028／2030／2032（C、D、E 欄＝低、基準、高）、NOA 新車數 2026 → 命題模組。與線模型不一致（刻意）：線模型的搭載率 65%、每車 1.0 顆、ADAS 單價 200 美元不直接進公式——本模型把搭載率、每車顆數與速騰份額合併為「每台 NOA 車對應速騰出貨」，由 2026 出貨反推；單價由公司產品營收與出貨推導（約 1,160 人民幣，低於 200 美元約 1,330 人民幣）。線 6 的 2030 LiDAR 需求只作量級對照。線 1 Robotaxi 無可讀變數。

## 限制與未取得資料
- 公司未給全年量化指引；Q3 出貨取自媒體引述的公司公告（Interested-party），Q3 財報未出。
- 同業取自資料商 statistics 頁（財務更新日 2026-11-24 前後），市值以 2026-11-20 收盤重算；禾賽財務幣別資料商未標明；各群組少於 8 家，不做迴歸。
- 資料商顯示速騰 5 年 β 0.24（上市未滿 2 年）不採用，β 用類比值 1.60（區間 1.3–1.9）。
- 股數 471.24 百萬（資料商）與庫藏股 12.48 百萬股的關係未核對。
- 未做：獨立核對（待外層流程）；check_display（無 display_check.json）。

## 檔案
- `build_robosense.py`：建置＋整本重算 Sens／反解＋產生本 README、`segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-24）：初版；現價 {IV['px']:.2f} 港元（2026-11-20 收盤）；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {tpl('rt')}、投資建議 {tpl('rec')}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


def main():
    tmpd = tempfile.mkdtemp(prefix="rbs0_"); p0 = os.path.join(tmpd, "b.xlsx"); VR0, _, _ = build(p0); o0 = recalc_many([p0], os.path.join(tmpd, "o"))[0]
    wv = openpyxl.load_workbook(o0, data_only=True); rel_now = wv["Val"][f"B{VR0['rel']}"].value
    evmed = wv["Val"][f"B{VR0['evm']}"].value; sg1 = wv["Val"][f"B{VR0['sg1']}"].value; sg2 = wv["Val"][f"B{VR0['sg2']}"].value
    evrow = [r_ for r_ in wv["In"].iter_rows(values_only=True) if r_[0] == "ev_med"][0]; ev_lo, ev_hi = evrow[3], evrow[5]
    shutil.rmtree(tmpd, ignore_errors=True)
    for i, row in enumerate(ROWS):
        if row[0] == "rel_k": ROWS[i] = row[:3] + (round(rel_now, 4),) + row[4:]
    base = run_cases([("base", {})])["base"]; px0 = in_val("px", "base")
    print("基準", base)
    jobs = []; meta = {}
    for k, drv in CASES:
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if k == "beta": v = in_val("beta", "base") + (-0.3 if side == "lo" else 0.3)
            if k == "ev_med": v = ev_lo if side == "lo" else ev_hi
            if v is None: continue
            b_ = in_val(k, "base")
            if isinstance(b_, (int, float)) and abs(v - b_) < 1e-12: continue
            tag = f"{k}_{side}"; jobs.append((tag, {k: v}))
            meta[tag] = dict(drv=drv + ("（低）" if side == "lo" else "（高）"), set=f"{k}＝{v:.4g}", key=k, val=v)
    for side in ("lo", "hi"):
        ov = {k: in_val(k, side) for k in MKK}
        tag = f"pen_{side}"; jobs.append((tag, ov))
        meta[tag] = dict(drv="中國城市 NOA 滲透率 2026–2032（情境 B，五年同動，線 2 " + ("低" if side == "lo" else "高") + "）" + ("（低）" if side == "lo" else "（高）"),
                         set="線 2 下游取數 " + ("C" if side == "lo" else "E") + " 欄", key=", ".join(MKK), val=None)
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
    tiers = [0.5, 0.25, -0.25, -0.5]; up0 = base["up"]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    targets = [("目標價＝現價（上檔 0%）", 0.0)] + ([(f"升一級邊界（上檔 {upper:+.0%}）", upper)] if upper is not None else []) + ([(f"降一級邊界（上檔 {lower:+.0%}）", lower)] if lower is not None else [])
    SOLVE = [("shr_B", -0.30, 0.30, "情境 B 每台 NOA 車對應速騰 ADAS 出貨的年變化（FY2027–FY2032 每年；情境 A／C 不動）"),
             ("gm0", 0.05, 0.60, "毛利率起點（三情境同動；A／C 另加各自的年變化）"),
             ("ev_med", 0.0, 10.0, "同業 EV ÷ 營收倍數（Comparable；DCF 不動）")]
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
            bv = in_val(key, 'base')
            revs.append(dict(label=lab, key=key, val=(a + b2) / 2 if a is not None else f"區間 {lo:g}～{hi:g} 內無解", base=round(float(bv), 4) if isinstance(bv, (int, float)) else 0, note=note))
            print("反解", key, lab, revs[-1]["val"])
    sens = dict(base=base, cases=cases, fp=base["fp"], rel_now=rel_now, evmed=evmed, sg1=sg1, sg2=sg2)
    fpr = hashlib.sha1(json.dumps(ROWS, ensure_ascii=False, default=str).encode()).hexdigest()[:10]
    sens["build_id"] = f"{VER}-{fpr}"
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="rbsf_"); raw = os.path.join(tmp, NAME)
    VR, PR, DR = build(raw, {}, sens, revs)
    out = recalc_many([raw], os.path.join(tmp, "o"))[0]
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
    json.dump(dict(company="robosense", ticker="2498.HK", model_sheet="Q",
                   official_segments=[dict(name=n, label=l, period="H1'26", value=v, unit="RMB m", source=IR26)
                                      for n, l, v in (("產品（Products）", "產品營收", 958.003), ("解決方案（Solutions）", "解決方案營收", 17.220), ("服務與其他（Services and others）", "服務與其他營收", 44.526))],
                   sublines=[], check_labels=["檢查：分類合計 − 營收（應為 0）"],
                   note="公司披露營收分類為產品、解決方案、服務與其他（中期業績公告）；Q 頁 H1'25、H1'26 對帳。ADAS 與機器人營收分拆公司未揭露金額（只稱機器人近一半），PL 頁為推導值；PL 頁另有「檢查：分部合計 − 營收（應為 0）」（FY2026 起）"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)

if __name__ == "__main__":
    sys.path.insert(0, HERE)
    main()

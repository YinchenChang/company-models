# Mobileye（MBLY，Nasdaq）個股模型 v0.1 — AI實體應用 Project（線 2 智慧車：ADAS／自駕系統方案；兼讀線 1 Robotaxi）
# 用法：python3 physical_ai/mobileye/build_mobileye.py            → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、segments.json、sens.json
#       python3 physical_ai/mobileye/build_mobileye.py --raw OUT  → 只寫一本（讀環境變數 MBLY_OVERRIDE 的 JSON 覆寫 In 頁基準值），供重算
# 會計年度＝曆年（52／53 週，12 月最後一個週六結算；US GAAP）。金額：百萬美元（US$M）；股數：百萬股；EPS、股價：美元。
# 結構：README、In、Hist、Q（季度分部與分部對帳）、PL（三情境：命題模組 → 官方分部 → Non-GAAP → GAAP 標準損益表 → 現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加；NOPAT 以調整後營業利益扣股票報酬為 EBIT）。
import sys, os, json, subprocess, shutil, hashlib, tempfile
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
VER = "v0.1"
DATE = "20261123"
NAME = f"{DATE}_Mobileye個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032

# ---------------- 出處 ----------------
K3Q = "Mobileye 8-K 附件 99.1 第 3 季財報新聞稿（2026-10-29）https://www.sec.gov/Archives/edgar/data/0001910139/000110465926119830/tm2628197d1_ex99-1.htm"
Q3Q = "Mobileye 10-Q（2026-09-26 季，Note 9 分部資訊、封面股數、R&D Law 段）https://www.sec.gov/Archives/edgar/data/0001910139/000110465926120226/mbly-20260926x10q.htm"
Q2Q = "Mobileye 10-Q（2026-06-27 季，分部資訊、Note 13 Mentee、商譽減損）https://www.sec.gov/Archives/edgar/data/1910139/000110465926086264/mbly-20260627x10q.htm"
K4Q = "Mobileye 8-K 附件 99.1 第 4 季及 2025 全年財報新聞稿（2026-01-22）https://www.sec.gov/Archives/edgar/data/1910139/000110465926005578/tm263599d1_ex99-1.htm"
K4QH = "Mobileye 8-K（2026-01-22，ir.mobileye.com 版，季度出貨表）https://ir.mobileye.com/node/9076/html"
NQ2 = "Nasdaq／Zacks（2026-07-24）Q2 財報 R&D 補助報導 https://www.nasdaq.com/articles/mobileye-q2-earnings-surpass-expectations-rd-incentive-grant"
MB = "Nasdaq／MarketBeat（2026-08-15）轉述公司於投資人會議的說法 https://www.nasdaq.com/articles/mobileye-eyes-higher-asps-surround-adas-volkswagen-programs-ramp"
SA = "stockanalysis.com（資料商）"
SAI = SA + " MBLY 損益表頁 https://stockanalysis.com/stocks/mbly/financials/income-statement/（讀取 2026-11-23）"
SAS = SA + " MBLY 統計頁（2026-10-26 收盤、10-27 更新）https://stockanalysis.com/stocks/mbly/statistics/"
YCL = "Yahoo Finance 日線（physical_ai/data/closes2_20261109_20261120.csv）"
RF = "美國 10 年期公債殖利率 2026-10-23 5.00%（美國財政部每日殖利率曲線；與 physical_ai/tesla v0.3.1、AI半導體 NVIDIA／AMD 同一取值，跨模型一致）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"
MBI = "媒體轉述公司投資人會議說法（立場：公司對自家產品價格與量的說法）"
LSC = "physical_ai/lines/smart_car/dist/20261116_智慧車單位經濟_v0.1.xlsx"
LRT = "physical_ai/lines/robotaxi/dist/20261116_Robotaxi單位經濟_v0.2.xlsx"

# ---------------- In 列 ----------------
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. 2026 年前三季實績（2025-12-28～2026-09-26；Verified）")
inp("n9_rev", "營收 9M 2026", "US$M", None, 1558, None, "Verified", K3Q + "：Revenue 1,558（9M 2025 1,448）", "")
inp("n9_agp", "調整後毛利 9M 2026（Non-GAAP）", "US$M", None, 1017, None, "Verified", K3Q + "：Adjusted gross profit 1,017（65%）", "GAAP 毛利 771；差額＝無形資產攤銷 245＋股票報酬 1")
inp("n9_aoi", "調整後營業利益 9M 2026（＝分部績效合計）", "US$M", None, 339, None, "Verified", K3Q + "：Adjusted operating income 339", Q3Q + "：Segment performance 合計 339（Mobileye 346、Other −7）")
inp("n9_inc", "R&D Law 補助 9M 2026（Non-GAAP 口徑，已含在調整後營業利益）", "US$M", None, 144, None, "Verified", K3Q + "：R&D Law incentive 對調整後營業利益 144（GAAP 162）", "Q2 一次認列上半年；以色列 2026 年研發獎勵法（可退還的研發租稅抵減，2026-03-29 立法、自 2026-01-01 起適用）")
inp("n9_fin", "財務收入淨額 9M 2026", "US$M", None, 41, None, "Verified", K3Q, "")
inp("n9_tax", "所得稅 9M 2026（負值＝利益）", "US$M", None, -61, None, "Verified", K3Q + "：Benefit for income taxes 61", "")
inp("n9_ani", "調整後淨利 9M 2026（Non-GAAP）", "US$M", None, 346, None, "Verified", K3Q, "")
inp("n9_mbly", "Mobileye 分部營收 9M 2026", "US$M", None, 1530, None, "Verified", Q3Q + "：Note 9 Mobileye revenue 1,530", "")
inp("n9_oth", "Other（Moovit）分部營收 9M 2026", "US$M", None, 28, None, "Verified", Q3Q + "：Note 9 Other revenue 28", "")
inp("n9_sys", "EyeQ＋SuperVision 系統營收 9M 2026（Q1 535＋Q2 485＋Q3 468）", "US$M", None, "=535+485+468", None, "Derived", K3Q + "：補充系統資料表逐季（Verified 原值）相加", "")
inp("n9_u", "系統出貨 9M 2026（Q1 10.8＋Q2 10.0＋Q3 9.5）", "百萬套", None, "=10.8+10+9.5", None, "Derived", K3Q + "：Systems shipped 逐季相加", "前三季 EyeQ 出貨量年增 11%（公司）")
inp("q3_asp", "平均系統售價 Q3 2026", "US$", None, 49.1, None, "Verified", K3Q + "：Average system price 49.1（Q3 2025 51.7；Q4 2025 50.8；Q1 2026 49.3；Q2 48.5）", "Q3 營收年減 2% 主因單價 −5%")
inp("n9_sbc", "股票報酬 9M 2026", "US$M", None, 275, None, "Verified", K3Q + "：營業利益調節 SBC 275", "")
inp("n9_am", "併購無形資產攤銷 9M 2026", "US$M", None, 296, None, "Verified", K3Q + "：營業利益調節 296（其中毛利調節 245）", "")
inp("n9_amc", "其中列在營業成本的攤銷 9M 2026", "US$M", None, 245, None, "Verified", K3Q, "")
inp("int_bal", "併購無形資產淨額（2026-09-26）", "US$M", None, 998, None, "Verified", K3Q + "：Intangible assets, net 998（2025-12-27 為 1,166）", "")

sec("B. 2026 全年錨點：公司指引（2026-10-29 更新；Interested-party）")
inp("g26_rev", "公司指引：FY2026 營收（區間中點）", "US$M", None, "=(1995+2020)/2", None, "Interested-party", K3Q + "：Revenue 1,995–2,020（前次 1,970–2,020）", "年增 5–7%", CO)
inp("g26_aoi", "公司指引：FY2026 調整後營業利益（中點）", "US$M", None, "=(380+420)/2", None, "Interested-party", K3Q + "：Adjusted operating income 380–420", "中點含 R&D Law 補助約 190", CO)
inp("g26_inc", "公司指引：FY2026 R&D Law 補助（Non-GAAP，中點）", "US$M", None, 190, None, "Interested-party", K3Q + "：中點約 190", NQ2 + "：公司 7 月預估全年 180–200", CO)
inp("g26_sbc", "公司指引：FY2026 股票報酬", "US$M", None, 385, None, "Interested-party", K3Q + "：Share-based compensation 385（前次 396）", "", CO)
inp("g26_am", "公司指引：FY2026 併購無形資產攤銷", "US$M", None, 346, None, "Interested-party", K3Q, "", CO)
inp("g26_oth", "公司指引：FY2026 商譽減損＋併購費用 − R&D Law RSU 項", "US$M", None, "=3788+6-18", None, "Interested-party", K3Q + "：商譽減損 3,788（Q1 已認列）、併購費用 6、R&D Law RSU 項 (18)", "GAAP 營業虧損指引 (4,127)–(4,087)，中點與本列及上列相符（Derived 檢查見 Recon）", CO)
inp("g_anc", "FY2026 實績相對公司指引（營收、調整後營業利益同比例）", "%", -0.02, 0.0, 0.02, "Assumed", "基準＝指引中點（前三季已實績，偏離只來自 Q4）", "低／高 ±2%：指引區間寬 25（±0.6%），另含 Tier 1 年末去庫存風險（公司第 2 季說法）")
inp("q4_oth", "Q4 2026 Other（Moovit）營收", "US$M", None, 9, None, "Assumed", "＝Q3 2026 單季 9（" + Q3Q + "）", "")
inp("q4_rev", "Q4 2026 隱含營收", "US$M", None, "=g26_rev*(1+g_anc)-n9_rev", None, "Derived", "＝全年指引中點 ×(1＋g_anc) − 9M 實績", "")
inp("q4_mo", "Q4 2026 Mobileye 分部中非系統營收（REM、NRE、記憶體轉嫁等）", "US$M", None, "=(n9_mbly-n9_sys)/3", None, "Derived", "＝(9M Mobileye 分部 − 9M 系統營收) ÷ 3", "Q3 含記憶體成本轉嫁約 7（零毛利，" + K3Q + "）")
inp("q4_sys", "Q4 2026 隱含系統營收", "US$M", None, "=q4_rev-q4_oth-q4_mo", None, "Derived", "", "")
inp("q4_u", "Q4 2026 隱含系統出貨", "百萬套", None, "=q4_sys/q3_asp", None, "Derived", "＝Q4 系統營收 ÷ Q3 平均售價", "Q4 2025 實績 8.3 百萬套")
inp("sh_26", "FY2026 平均股數", "百萬股", None, 822, None, "Derived", K3Q + "：9M 加權 820、Q3 824", "取 9M 與 Q3 之間；GAAP 虧損，基本＝稀釋")

sec("C. 命題模組：EyeQ 出貨 × 單價（基礎 ADAS）＋ Surround ADAS ＋ SuperVision／Chauffeur ＋ Drive（Robotaxi）")
inp("lvp", "全球輕型車產量年增（FY2027 起）", "%/年", 0.0, 0.01, 0.025, "Assumed", "公司：前三季全球汽車產量約 −1%（" + K3Q + "）", "低 0%／高 2.5%；IHS／S&P 長期約 1–2%（未查核，不作依據）")
for s, v, lo, hi, note in (("A", 0.04, None, None, "樂觀：歐盟 GSR、NCAP 新規推升基礎 ADAS 標配，日系與亞洲新定點（2026 年取得前五大亞洲車廠、前十大日本車廠雲端增強 ADAS 定點）"),
                           ("B", 0.01, -0.01, 0.03, "基準：2026 前十大客戶出貨 +7% 而其產量 −5%（含安全庫存，" + K3Q + "），扣除安全庫存回補後，長期每年 +1%"),
                           ("C", -0.04, None, None, "保守：中國車廠自研與本土方案（地平線、華為、NVIDIA）擠壓，Tesla 式自研擴散；2026 Q4 2025 出貨年減 11% 為類比")):
    inp(f"pen_{s}", f"情境 {s}：EyeQ 出貨相對車產的滲透／份額年變化", "%/年", lo, v, hi, "Assumed", note, "區間見情境 A／B／C 三列")
for s, v, lo, hi, note in (("A", -0.005, None, None, "樂觀：基本款占比不再上升"), ("B", -0.02, -0.035, -0.005, "基準：2025→2026 平均售價約 −3%（Q3 年減 5%），每年 −2%"),
                           ("C", -0.04, None, None, "保守：中國出口車與低價基本款占比續升（Seeking Alpha 2026-07-23，經公司檔案轉述）")):
    inp(f"asp_{s}", f"情境 {s}：基礎 ADAS 平均售價年變化", "%/年", lo, v, hi, "Assumed", note, "Q3 2026 49.1 美元（Verified）為起點")
for s, v, lo, hi, note in (("A", 10.0, None, None, "樂觀：三個定點全數轉換，另有新定點"), ("B", 5.0, 2.5, 7.5, "基準：三個定點各約占該車廠 30–40% 銷量、2028 年上市（公司說法）；以約 1,500 萬台 × 35% ≈ 5 百萬套（Assumed）"),
                           ("C", 2.0, None, None, "保守：上市延後、轉換慢")):
    inp(f"srd_{s}", f"情境 {s}：Surround ADAS 出貨 FY2032（2028 年起線性）", "百萬套", lo, v, hi, "Assumed", note, MB)
inp("srd_asp", "Surround ADAS 平均售價", "US$", 100, 125, 140, "Interested-party", MB + "：約為基礎款的 2.5–3 倍、約 125 美元；毛利率與基礎款相近（中高 60%）", "低／高 100／140（Assumed）", MBI)
for s, v, lo, hi, note in (("A", 0.6, None, None, "樂觀：公司稱第三年約 30–35 萬台，之後再增"), ("B", 0.3, 0.15, 0.45, "基準：公司稱 SuperVision 第三年約 30–35 萬台（2027 Q1 Porsche 首款）；FY2032 取 30 萬"),
                           ("C", 0.08, None, None, "保守：延後、只有單一車廠")):
    inp(f"sv_{s}", f"情境 {s}：SuperVision／Chauffeur 出貨 FY2032（2027 年起線性）", "百萬套", lo, v, hi, "Assumed", note, MB)
inp("sv_u26", "SuperVision 出貨 FY2026", "百萬套", 0.005, 0.015, 0.03, "Assumed", "未揭露（公司只揭露 EyeQ＋SuperVision 合計）", "")
inp("sv_asp", "SuperVision／Chauffeur 混合平均售價", "US$", 1300, 1500, 2500, "Assumed", MB + "：SuperVision 約 1,300 美元、Chauffeur 2,500–3,000 美元（Interested-party）", "混合比例未揭露，基準 1,500")
inp("gm_sv", "SuperVision／Chauffeur 調整後毛利率", "%", 0.30, 0.45, 0.60, "Assumed", "未揭露；含外購感測與運算板", "")
inp("gm_oth", "Mobileye 其他與 Moovit 調整後毛利率", "%", 0.30, 0.50, 0.70, "Assumed", "未揭露；記憶體轉嫁為零毛利（公司）", "")
inp("gm_drv", "Drive（Robotaxi 系統）調整後毛利率", "%", 0.15, 0.30, 0.45, "Assumed", "未揭露；系統含外購感測器", "")
inp("core26", "FY2026 基礎 ADAS 營收（系統營收 − SuperVision）", "US$M", None, "=n9_sys+q4_sys-sv_u26*sv_asp", None, "Derived", "", "")
inp("oth26", "FY2026 Mobileye 其他＋Moovit 營收", "US$M", None, "=(n9_mbly-n9_sys)+q4_mo+n9_oth+q4_oth", None, "Derived", "", "")
inp("agp26", "FY2026 調整後毛利（9M 實績＋Q4 以 9M 毛利率）", "US$M", None, "=n9_agp+q4_rev*n9_agp/n9_rev", None, "Derived", "", "Q4 毛利率未指引")
inp("gm_core", "基礎 ADAS／Surround ADAS 調整後毛利率（FY2026 校準）", "%", None, "=(agp26-sv_u26*sv_asp*gm_sv-oth26*gm_oth)/core26", None, "Derived",
    "＝(FY2026 調整後毛利 − SuperVision × gm_sv − 其他 × gm_oth) ÷ 基礎 ADAS 營收", "Surround ADAS 毛利率與基礎款相近（公司說法，" + MB + "）；校準值隨 gm_sv、gm_oth、sv_u26 的 Sens 重算")
inp("drv_asp", "Drive（Robotaxi）每車系統價格", "US$", 25000, 40000, 60000, "Interested-party", MB + "：Mobileye Drive 約 4 萬美元／套（初期數百台約 10 萬美元／台含車）", "低／高 25,000／60,000（Assumed）", MBI)
for s, a27, a30, a35 in (("A", 7503.12160666208, 58536.3687437831, 115734.46517585), ("B", 5332.09477649166, 30729.5790760588, 68688.4164712717), ("C", 2864.16915962494, 7904.24250736704, 17372.040191682)):
    col = {"A": "E（高）", "B": "D（基準）", "C": "C（低）"}[s]
    for k, y, v, r in (("rt27", 2027, a27, 55), ("rt30", 2030, a30, 56), ("rt35", 2035, a35, 57)):
        inp(f"{k}_{s}", f"情境 {s}：線 1 美國 Robotaxi 年新增車輛需求 {y}", "台", None, v, None, "Derived",
            f"讀 {LRT}「下游取數」第 {r} 列 RT_VEH_ADDS_{y} 欄 {col}（線模型值；情境 A 讀高、B 讀基準、C 讀低）", "線模型口徑：擴車＋汰換；中間年線性內插")
for s, v, lo, hi, note in (("A", 0.15, None, None, "樂觀：Lyft／Uber／VW MOIA 美國擴點順利"), ("B", 0.08, 0.03, 0.12, "基準：美國以 Waymo、Tesla 自有系統為主，Mobileye 經 VW ID. Buzz（MOIA、Uber、Lyft）與 Holon 取得少數份額"),
                           ("C", 0.02, None, None, "保守：無人駕駛（driver-out）里程碑延後")):
    inp(f"dshr_{s}", f"情境 {s}：Mobileye 占美國 Robotaxi 新增車輛比例", "%", lo, v, hi, "Assumed", note, "公司稱德國 L4 測試許可車輛中 Mobileye Drive 占 60% 以上（" + K3Q + "，Interested-party）")
inp("geo_m", "美國以外（歐洲等）Drive 車輛倍數（全球 ÷ 美國）", "x", 1.0, 2.0, 3.0, "Assumed", "漢堡、奧斯陸（Holo 挪威許可）、德鐵 HOLON；線 1 只有美國口徑", "未取得歐洲 Robotaxi 車輛需求的線模型")
inp("mo_g", "Mobileye 其他營收（REM、NRE、記憶體轉嫁）年增", "%/年", 0.0, 0.03, 0.06, "Assumed", "", "")
inp("oth_g", "Other（Moovit）營收年增", "%/年", -0.03, 0.02, 0.06, "Assumed", "9M 2026 28 vs 9M 2025 29（" + Q3Q + "）", "")
inp("cust_loss", "最大客戶部分轉單（FY2028 起基礎 ADAS 出貨減少比例；尾端事件，基準 0）", "%", None, 0.0, 0.15, "Assumed", Q3Q + "：客戶 A 占 9M 營收 28%、客戶 D 16%、C 15%、B 14%", "高案 15%＝客戶 A 約一半量轉自研或他家（原則第 10 節第 19 條）")

sec("D. 營業費用、補助、股票報酬與攤銷")
inp("aop26", "FY2026 調整後營業費用（未扣 R&D Law 補助）", "US$M", None, "=agp26-g26_aoi*(1+g_anc)+g26_inc", None, "Derived", "＝調整後毛利 − 調整後營業利益指引 ＋ 補助", "")
inp("opex_g", "調整後營業費用年增（FY2027 起）", "%/年", 0.02, 0.04, 0.06, "Assumed", "9M 2026 研發淨額 814 vs 9M 2025 861（含補助）；毛額約 958 vs 861（+11%，含 Mentee 併入）", "人形機器人（Mentee）研發另計在本列")
inp("inc_r", "R&D Law 補助 ÷ 調整後營業費用（FY2026）", "%", None, "=g26_inc/aop26", None, "Derived", "", "補助＝以色列合格研發支出 × 比率（比率未揭露）")
inp("inc_k", "R&D Law 補助延續係數（FY2027 起）", "x", 0.0, 1.0, None, "Assumed", "法律無到期日（" + NQ2 + "）；執行長稱可持續提高利潤率基準（" + Q2Q + "）", "低案 0＝取消或改制；補助以租稅抵減入帳、現金收取有時差（10-Q：可選擇現金給付）")
inp("sbc_r", "股票報酬 ÷ 營收（FY2027 起，全期持平）", "%", 0.14, "=(71+n9_sbc)/(446+n9_rev)", 0.20, "Derived", "最近四季實績：(Q4 2025 71＋9M 2026 275) ÷ (446＋1,558)（" + K4Q + "；" + K3Q + "）", "FY2025 14.6%（277÷1,894）、FY2024 16.9%；2026-07 授予 3,502 萬股 RSU（3.34 億美元，" + Q2Q + "）；低／高 Assumed")
inp("am_rate", "併購無形資產年攤銷（FY2027 起，Q3 2026 × 4）", "US$M", None, "=69*4", None, "Derived", K3Q + "：Q3 2026 攤銷 69", "攤銷時程表（10-K 未來五年攤銷）未取得；以 Q3 年化、攤完為止")
inp("am_cg", "攤銷列在營業成本的比例", "%", None, "=n9_amc/n9_am", None, "Derived", "＝245 ÷ 296", "")

sec("E. 損益表下半段與現金流")
inp("fin_y", "財務收入 ÷ 現金與有價證券", "%", 0.025, "=n9_fin/0.75/((1891+cash_v+sec_v)/2)", 0.045, "Derived", "＝9M 財務收入年化 ÷ 期初期末平均現金與證券（2025 年底 1,891；" + K3Q + "）", "FY2027 起以評價日現金固定（未建現金累積）")
inp("tax_w", "有效稅率（Non-GAAP 與 FY2027 起 GAAP）", "%", 0.05, "=(n9_aoi+n9_fin-n9_ani)/(n9_aoi+n9_fin)", 0.15, "Derived", "＝9M 調整後稅前 380 − 調整後淨利 346，÷ 380（" + K3Q + "）", "以色列優先科技事業稅率 7.5%／12%（未查核，不作依據）；低／高 Assumed")
inp("da_r", "折舊 ÷ 營收", "%", 0.035, "=64/n9_rev", 0.05, "Derived", K3Q + "：9M 折舊 64 ÷ 營收 1,558", "FY2025 74 ÷ 1,894 ＝ 3.9%（" + K4Q + "）")
inp("cx_r", "資本支出 ÷ 營收", "%", 0.04, "=72/n9_rev", 0.06, "Derived", K3Q + "：9M 資本支出 72 ÷ 1,558", "FY2025 79 ÷ 1,894 ＝ 4.2%、FY2024 81 ÷ 1,654 ＝ 4.9%")
inp("wc_r", "營運資金增加 ÷ 營收增量", "%", 0.05, "=(202+295-272)/(g26_rev)", 0.20, "Derived", K3Q + "：(應收 202＋存貨 295 − 應付及應計 272) ÷ FY2026 營收", "2025 年底 (131＋327 − 228) ÷ 1,894 ＝ 12.1%")
inp("pay_r", "配息率（無股利）", "%", None, 0.0, None, "Verified", "公司未發放現金股利", "")

sec("F. 估值輸入（評價日 2026-12-31；資產負債表 2026-09-26）")
inp("px", "股價（2026-11-20 收盤）", "US$", None, 4.57, None, "Verified", YCL + "：MBLY 2026-11-20 收盤 4.57", "同檔 11-09 5.86、11-19 5.09；11-20 單日 −10.2%，原因未找到")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("sh_a", "A 類流通股（2026-10-21）", "百萬股", None, 252.403844, None, "Verified", Q3Q + "：封面 252,403,844 股", "含 Mentee 創辦人限制股 2,365 萬股")
inp("sh_b", "B 類股（Intel 全數持有；2025-12-27）", "百萬股", None, 597.768015, None, "Verified", Q2Q + "：資產負債表 597,768,015 股", "2026 年後 B 類股數未取得（假設不變）")
inp("rsu", "未歸屬 RSU（2026-09-26）", "百萬股", None, 61.123, None, "Verified", Q3Q + "：Unvested RSUs 61,123 千股、未認列成本 5.54 億美元", "選擇權：未取得（Mentee 承接選擇權公允價值 7 百萬美元，量小不計）")
inp("mentee_rs", "Mentee 創辦人未歸屬限制股（收購後服務期獎酬）", "百萬股", None, 23.652, None, "Verified", Q3Q + "：23,652 千股、未認列成本 1.73 億美元；24／48 個月歸屬", "DCF 股數扣除（剩餘成本由預測期股票報酬扣除；DCF 標準規格第 3 節）")
inp("sh_dcf", "DCF 評價日股數（流通＋RSU − Mentee 限制股）", "百萬股", None, "=sh_a+sh_b+rsu-mentee_rs", None, "Derived", "DCF 標準規格第 3 節", "")
inp("sh_cmp", "Comparable／EPS 股數（流通＋RSU）", "百萬股", None, "=sh_a+sh_b+rsu", None, "Derived", "調整後 EPS 不扣股票報酬，股數含全部授予（DCF 標準規格第 3 節）", "")
inp("cash_v", "現金及約當現金（2026-09-26）", "US$M", None, 1346, None, "Verified", K3Q, "")
inp("sec_v", "有價證券及存款（2026-09-26）", "US$M", None, 153, None, "Verified", K3Q, "無借款（資產負債表無借款科目）；租賃負債未取得（量小，不計）")
inp("recv_v", "R&D Law 補助應收（其他長期資產，2026-09-26）", "US$M", None, 162, None, "Verified", Q3Q + "：已認列資產 162", "")
inp("recv_k", "R&D Law 應收計入比例", "x", 0.5, 1.0, None, "Assumed", "可抵以色列所得稅或選擇現金給付；時差與折現未建", "")
inp("stub_frac", "Stub 期（2026-09-26 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,9,26))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（美國 10 年期公債殖利率，2026-10-23）", "%", None, 0.05, None, "Verified", RF, "日期早於股價日 2 個交易日以上：逾期（跨模型一致，下次全面改版更新）")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "國家風險溢酬（以色列營運；中東地緣風險）", "%", 0.008, 0.012, 0.018, "Assumed", "Damodaran 以色列值未查核，不作依據；公司 2026 Q1 商譽減損說明折現率因中東地緣與競爭風險溢酬上升（" + Q2Q + "）", "")
inp("beta", "β（兩來源 Blume 調整後中位數）", "x", "=E{r}-0.3", "=MEDIAN(0.67*1.24+0.33,0.67*MEDIAN(1.76,1.47,2.38,2.35,0.61)+0.33)", "=E{r}+0.3", "Derived",
    SA + "：MBLY 5 年 β 1.24；同業 5 家 5 年 β 中位數 1.76（Qualcomm 1.76、禾賽 1.47、Ambarella 2.38、地平線 2.35、DENSO 0.61）", "Blume 0.67β＋0.33；敏感度 ±0.3（Assumed）")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺 Mobileye 對同業的 10 年前瞻 P/E 比值歷史）", "第 10 節第 18、47 條：列為承重輸入；Sens 另列「改用現行相對倍數」")
inp("cmp_sbc", "Comparable 的 EPS 是否扣股票報酬（0＝調整後 EPS，同業倍數口徑；1＝扣稅後股票報酬）", "x", None, 0, 1, "Assumed", "同業前瞻 P/E 為賣方 non-GAAP 口徑（多數不含股票報酬）", "Mobileye 股票報酬占營收約 17%，高於多數同業（同業比率未取得）；高案只作 Sens")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業前瞻 P/E（資料商頁面；Interested-party：賣方共識 EPS）")
PEERS = [("pe_qcom", "Qualcomm QCOM（車用與手機 SoC）", 15.58, "2026-11-20 收盤", True),
         ("pe_hsai", "禾賽 HSAI（車用光達）", 17.86, "2026-11-20 收盤", True),
         ("pe_amba", "Ambarella AMBA（車用視覺 SoC）", 74.89, "2026-10-23 收盤", False),
         ("pe_denso", "DENSO 6902.T（Tier 1，ADAS 模組）", 11.58, "2026-10-20 收盤", False)]
for k, lab, v, d, upd in PEERS:
    note = "報價日 2026-11-20，與本模型股價日相同" if upd else "報價日未更新（csv 有 2026-11-20 收盤，但舊報價日不在 csv 期間 2026-11-09～11-20，無舊收盤可等比例換算）"
    inp(k, f"前瞻 P/E：{lab}", "x", None, v, None, "Interested-party", f"{SA} 統計頁：{d} Forward PE {v}；{note}",
        "會計口徑：美國公司 US GAAP、禾賽與地平線 IFRS／US GAAP 申報（未逐家核對）、DENSO 日本基準或 IFRS；P/E 為股權口徑不受租賃處理影響。地平線（9660.HK）虧損、無前瞻 P/E，不列", DV)
inp("pe_med", "同業前瞻 P/E 中位數（4 家）", "x", None, "=MEDIAN(pe_qcom,pe_hsai,pe_amba,pe_denso)", None, "Derived", "少於 8 家，不做迴歸（第 10 節第 7 條）；Ambarella 為離群值（中位數不受影響）", "")
inp("pe_self", "Mobileye 現行前瞻 P/E（換算至 2026-11-20；只作對照，不進模型）", "x", None, "=15.77*px/6.66", None, "Derived", SAS + "：Forward PE 15.77 於收盤 6.66（Interested-party）× 2026-11-20 收盤 ÷ 6.66（分母 EPS 不變）", "隱含 NTM 調整後 EPS ＝ 6.66 ÷ 15.77（第 3 季財報前）", DV)
inp("cons_tp", "賣方平均目標價（只作對照，不進模型）", "US$", None, 11.11, None, "Interested-party", SAS + "：24 位分析師平均 11.11、評等 Buy（第 3 季財報前，只作歷史對照）", "", DV)
inp("pipe", "公司 8 年預期車用營收管線（2025 年底；只作對照，不進模型）", "US$M", None, 24500, None, "Interested-party", K4Q + "：8-year future expected automotive revenue pipeline 245 億美元", "年均約 30.6 億美元（Derived）", CO)

HIST = {  # 年: (營收, 營業成本, 毛利, 營業費用, 營業利益, 財務收入, 稅前, 所得稅(+費用), 淨利, EPS, 調整後毛利, 調整後營業利益, 調整後淨利, SBC, 攤銷, 出貨)
    2023: (2079, 1032, 1047, 1080, -33, 49, 16, 43, -27, -0.03, None, None, None, None, None, None),
    2024: (1654, 913, 741, 3966, -3225, 62, -3163, -73, -3090, -3.82, 1119, 193, 205, 279, 444, None),
    2025: (1894, 990, 904, 1344, -440, 63, -377, 15, -392, -0.48, 1283, 280, 286, 277, 443, "=8.5+9.7+9.2+8.3")}
HLAB = ["營收", "營業成本", "毛利", "營業費用（含商譽減損）", "營業利益", "財務收入淨額", "稅前利益", "所得稅（負值＝利益）", "淨利", "EPS（GAAP，基本＝稀釋）",
        "調整後毛利（Non-GAAP）", "調整後營業利益（Non-GAAP）", "調整後淨利（Non-GAAP）", "股票報酬", "併購無形資產攤銷", "系統出貨（百萬套）"]
HKEY = ["rev", "cogs", "gp", "opex", "op", "fin", "pbt", "tax", "ni", "eps", "agp", "aoi", "ani", "sbc", "am", "u"]


def build(path, ov=None, sens=None, rev=None):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 US$M、股數百萬股、EPS 與股價美元"
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
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("pen", "EyeQ 出貨相對車產的滲透／份額年變化"), ("asp", "基礎 ADAS 平均售價年變化"), ("srd", "Surround ADAS 出貨 FY2032（百萬套）"),
           ("sv", "SuperVision／Chauffeur 出貨 FY2032（百萬套）"), ("rt27", "線 1 美國 Robotaxi 年新增車輛 2027"), ("rt30", "線 1 美國 Robotaxi 年新增車輛 2030"),
           ("rt35", "線 1 美國 Robotaxi 年新增車輛 2035"), ("dshr", "Mobileye 占美國 Robotaxi 新增車輛比例")]
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
    H["A1"] = "Hist — 歷史財務（年度；US GAAP；US$M）。公式只引用本頁儲存格（定義名稱 h_*）。FY2024、FY2025：公司 8-K（Verified）；FY2023：資料商（法定原文未逐項核對）"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "標記", "來源"])
    for i, (lab, k) in enumerate(zip(HLAB, HKEY)):
        rr = 3 + i; H.cell(rr, 1, lab)
        for j, y in enumerate((2023, 2024, 2025)):
            v = HIST[y][i]
            if v is None: H.cell(rr, 2 + j, "未取得"); continue
            c = H.cell(rr, 2 + j, v); c.font = BLUE
            dn(f"h_{k}_{y % 100}", f"Hist!${'BCD'[j]}${rr}")
        H.cell(rr, 5, "Derived" if k == "u" else "Verified")
        H.cell(rr, 6, (K4QH + "：Q1 8.5＋Q2 9.7＋Q3 9.2＋Q4 8.3（逐季相加）") if k == "u" else (K4Q + "（FY2024、FY2025）；" + SAI + "（FY2023；營業費用＝研發 889＋銷管 191）"))
    H.cell(20, 1, "註：FY2024 營業費用含商譽減損 2,695；FY2026 第 1 季另認列商譽減損 3,788（2026-06-27 季 10-Q，Sources 頁）。FY2023 調整後各列未取得。")
    H.column_dimensions["A"].width = 36

    # ---------------- Q（季度分部、分部對帳） ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 官方報告分部（Mobileye、Other＝Moovit；10-Q Note 9）與分部對帳；US$M"
    hdr(Q, 2, ["項目", "Q3'25", "Q3'26", "9M'25", "9M'26", "", "標記", "來源", "來源誘因"])
    qd = [("Mobileye 分部營收", (494, 483, 1419, "=n9_mbly")), ("Other（Moovit）分部營收", (10, 9, 29, "=n9_oth")), ("合併營收", (504, 492, 1448, "=n9_rev")),
          ("Mobileye 分部績效", (76, 91, 246, 346)), ("Other 分部績效", (-2, -2, -7, -7)), ("分部績效合計（＝調整後營業利益）", (74, 89, 239, "=n9_aoi"))]
    for i, (lab, vals) in enumerate(qd):
        rr = 3 + i; Q.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            c = Q.cell(rr, 2 + j, v)
            if not (isinstance(v, str)): c.font = BLUE
        Q.cell(rr, 7, "Verified"); Q.cell(rr, 8, Q3Q)
    Q.cell(9, 1, "檢查：分部合計 − 合併營收（應為 0）")
    Q.cell(10, 1, "檢查：分部績效合計 − 兩分部（應為 0）")
    for j in range(4):
        L = "BCDE"[j]
        Q.cell(9, 2 + j, f"=ROUND({L}3+{L}4-{L}5,6)"); Q.cell(10, 2 + j, f"=ROUND({L}6+{L}7-{L}8,6)")
    Q.cell(9, 7, "Derived"); Q.cell(10, 7, "Derived")
    hdr(Q, 12, ["系統（EyeQ＋SuperVision）", "Q3'25", "Q4'25", "Q1'26", "Q2'26", "Q3'26", "標記", "來源", "來源誘因"])
    for i, (lab, vals) in enumerate((("系統營收", (478, 420, 535, 485, 468)), ("出貨（百萬套）", (9.2, 8.3, 10.8, 10.0, 9.5)), ("平均系統售價（美元）", (51.7, 50.8, 49.3, 48.5, 49.1)))):
        rr = 13 + i; Q.cell(rr, 1, lab)
        for j, v in enumerate(vals): Q.cell(rr, 2 + j, v).font = BLUE
        Q.cell(rr, 7, "Verified"); Q.cell(rr, 8, K3Q)
    Q.cell(16, 1, "主要客戶占營收（9M 2026）：客戶 A 28%、D 16%、C 15%、B 14%（" + Q3Q + "）")
    Q.column_dimensions["A"].width = 40

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（出貨 × 單價）→ 官方分部（Mobileye、Other）→ Non-GAAP → GAAP 標準損益表 → 現金流（US$M；FY2023–FY2025 實績、FY2026＝9M 實績＋Q4 指引隱含、FY2027 起模型）"
    PR = {}
    rr = 3
    KEYS = [("u_g", "EyeQ 系統出貨（未扣客戶轉單；百萬套）"), ("u", "EyeQ 系統出貨（百萬套）"), ("u_srd", "Surround ADAS 出貨（百萬套）"), ("u_sv", "SuperVision／Chauffeur 出貨（百萬套）"),
            ("u_base", "基礎 ADAS 出貨（百萬套）"), ("asp", "基礎 ADAS 平均售價（美元）"), ("rt", "線 1 美國 Robotaxi 年新增車輛（台）"), ("u_drv", "Drive（Robotaxi）出貨（台）"),
            ("base", "基礎 ADAS（EyeQ 前視；子列）"), ("srd", "Surround ADAS（子列）"), ("sv", "SuperVision／Chauffeur（子列）"), ("drv", "Drive（Robotaxi；子列）"),
            ("mo", "Mobileye 其他（REM、NRE、記憶體轉嫁；子列）"), ("mbly", "Mobileye 分部營收"), ("oth", "Other（Moovit）分部營收"), ("rev", "營收（合併）"),
            ("chk1", "檢查：Mobileye 子列合計 − Mobileye 分部（應為 0）"), ("chk2", "檢查：分部合計 − 合併營收（應為 0）"), ("gr", "營收年增率"),
            ("agp", "調整後毛利（Non-GAAP）"), ("agm", "調整後毛利率（Non-GAAP）"), ("aop", "調整後營業費用（未扣 R&D Law 補助；Non-GAAP）"), ("inc", "R&D Law 補助（Non-GAAP 口徑）"),
            ("aoi", "調整後營業利益（Non-GAAP；分部績效合計）"), ("aom", "調整後營業利益率（Non-GAAP）"), ("sbc", "股票報酬"), ("am", "併購無形資產攤銷"),
            ("ambal", "併購無形資產淨額（年末）"), ("oit", "商譽減損與併購費用（減 R&D Law RSU 項）"),
            ("cogs", "營業成本（GAAP）"), ("gp", "毛利"), ("gm", "毛利率"), ("opex", "營業費用（GAAP，含研發淨額、行銷、管理、商譽減損）"), ("op", "營業利益（GAAP）"),
            ("opm", "營業利益率（GAAP）"), ("fin", "業外收入（財務收入淨額）"), ("pbt", "稅前利益"), ("tax", "所得稅（負值＝利益）"), ("ni", "稅後淨利"), ("nim", "淨利率"),
            ("sh", "平均股數（百萬股）"), ("eps", "EPS（GAAP，美元）"), ("nig", "淨利年增率"), ("ani", "調整後淨利（Non-GAAP）"), ("aeps", "調整後 EPS（Non-GAAP，美元）"),
            ("ebit", "DCF EBIT（調整後營業利益 − 股票報酬）"), ("da", "折舊"), ("cx", "資本支出"), ("dwc", "營運資金增加"),
            ("fcff", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）"), ("fat", "稅後財務收入"), ("div", "股利"), ("fcfad", "自由現金流扣股利（收支關卡用；FCFF＋稅後財務收入 − 股利）")]
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, (k, _) in enumerate(KEYS)}
        for i, (_, l) in enumerate(KEYS): P.cell(b + i, 1, l)
        R["hdr"] = rr + 1
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1)
            cell = lambda k: f"{c}{R[k]}"; prev = lambda k: f"{p}{R[k]}"
            def w(k, f): P[cell(k)] = f
            if y <= 2025:
                yy = y % 100
                for k in ("rev", "cogs", "gp", "opex", "op", "fin", "pbt", "tax", "ni", "eps"):
                    w(k, f"=h_{k}_{yy}")
                if y >= 2024:
                    for k in ("agp", "aoi", "ani", "sbc", "am"): w(k, f"=h_{k}_{yy}")
                    w("agm", f"={cell('agp')}/{cell('rev')}"); w("aom", f"={cell('aoi')}/{cell('rev')}")
                if y == 2025: w("u", "=h_u_25")
                w("gm", f"={cell('gp')}/{cell('rev')}"); w("opm", f"={cell('op')}/{cell('rev')}"); w("nim", f"={cell('ni')}/{cell('rev')}")
                if y >= 2024:
                    w("gr", f"={cell('rev')}/{prev('rev')}-1")
                    w("nig", f'=IF({prev("ni")}>0,{cell("ni")}/{prev("ni")}-1,"N/M（前一年虧損）")')
                continue
            if y == 2026:
                w("u", "=n9_u+q4_u"); w("u_g", f"={cell('u')}"); w("u_srd", "=0"); w("u_sv", "=sv_u26")
                w("asp", f"={cell('base')}/{cell('u_base')}")
                w("base", "=core26"); w("srd", "=0"); w("sv", "=sv_u26*sv_asp"); w("drv", "=0"); w("rt", "=0"); w("u_drv", "=0")
                w("mo", "=(n9_mbly-n9_sys)+q4_mo"); w("oth", "=n9_oth+q4_oth")
                w("agp", "=agp26"); w("aop", "=aop26"); w("inc", "=g26_inc")
                w("sbc", "=g26_sbc"); w("am", "=g26_am"); w("ambal", "=int_bal-(g26_am-n9_am)"); w("oit", "=g26_oth")
                w("cogs", f"={cell('rev')}-{cell('agp')}+n9_amc+(g26_am-n9_am)*am_cg")
                w("fin", "=n9_fin+fin_y*(cash_v+sec_v)/4"); w("tax", "=n9_tax"); w("sh", "=sh_26")
                w("ani", f"=n9_ani+({cell('aoi')}-n9_aoi+{cell('fin')}-n9_fin)*(1-tax_w)")
            else:
                pen, asp = f"pen_{s}", f"asp_{s}"
                w("u_g", f"={prev('u_g')}*(1+lvp)*(1+{pen})")
                w("u_srd", f"=srd_{s}*MAX(0,{y}-2027)/5")
                w("u_sv", f"=sv_u26+(sv_{s}-sv_u26)*({y}-2026)/6")
                loss = "*(1-cust_loss)" if y >= 2028 else ""
                w("u", f"={cell('u_g')}{loss}")
                w("asp", f"={prev('asp')}*(1+{asp})")
                if y == 2027: rtf = f"=rt27_{s}"
                elif y <= 2030: rtf = f"=rt27_{s}+(rt30_{s}-rt27_{s})*({y}-2027)/3"
                else: rtf = f"=rt30_{s}+(rt35_{s}-rt30_{s})*({y}-2030)/5"
                w("rt", rtf); w("u_drv", f"={cell('rt')}*dshr_{s}*geo_m")
                w("base", f"={cell('u_base')}*{cell('asp')}"); w("srd", f"={cell('u_srd')}*srd_asp"); w("sv", f"={cell('u_sv')}*sv_asp")
                w("drv", f"={cell('u_drv')}*drv_asp/1000000"); w("mo", f"={prev('mo')}*(1+mo_g)"); w("oth", f"={prev('oth')}*(1+oth_g)")
                w("agp", f"=({cell('base')}+{cell('srd')})*gm_core+{cell('sv')}*gm_sv+{cell('drv')}*gm_drv+({cell('mo')}+{cell('oth')})*gm_oth")
                w("aop", f"={prev('aop')}*(1+opex_g)"); w("inc", f"=inc_r*{cell('aop')}*inc_k")
                w("sbc", f"=sbc_r*{cell('rev')}"); w("am", f"=MIN(am_rate,{prev('ambal')})"); w("ambal", f"={prev('ambal')}-{cell('am')}"); w("oit", "=0")
                w("cogs", f"={cell('rev')}-{cell('agp')}+{cell('am')}*am_cg")
                w("fin", "=fin_y*(cash_v+sec_v)"); w("tax", f"=tax_w*{cell('pbt')}"); w("sh", "=sh_cmp")
                w("ani", f"=({cell('aoi')}+{cell('fin')})*(1-tax_w)")
            w("u_base", f"={cell('u')}-{cell('u_srd')}-{cell('u_sv')}")
            w("mbly", f"={cell('base')}+{cell('srd')}+{cell('sv')}+{cell('drv')}+{cell('mo')}")
            w("rev", f"={cell('mbly')}+{cell('oth')}")
            w("chk1", f"=ROUND({cell('base')}+{cell('srd')}+{cell('sv')}+{cell('drv')}+{cell('mo')}-{cell('mbly')},6)")
            w("chk2", f"=ROUND({cell('mbly')}+{cell('oth')}-{cell('rev')},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            w("agm", f"={cell('agp')}/{cell('rev')}")
            w("aoi", f"={cell('agp')}-{cell('aop')}+{cell('inc')}"); w("aom", f"={cell('aoi')}/{cell('rev')}")
            w("op", f"={cell('aoi')}-{cell('sbc')}-{cell('am')}-{cell('oit')}")
            w("gp", f"={cell('rev')}-{cell('cogs')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
            w("opex", f"={cell('gp')}-{cell('op')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            w("pbt", f"={cell('op')}+{cell('fin')}"); w("ni", f"={cell('pbt')}-{cell('tax')}"); w("nim", f"={cell('ni')}/{cell('rev')}")
            w("eps", f"={cell('ni')}/{cell('sh')}"); w("nig", f'=IF({prev("ni")}>0,{cell("ni")}/{prev("ni")}-1,"N/M（前一年虧損）")')
            w("aeps", f"={cell('ani')}/{cell('sh')}")
            w("ebit", f"={cell('aoi')}-{cell('sbc')}")
            w("da", f"=da_r*{cell('rev')}"); w("cx", f"=cx_r*{cell('rev')}"); w("dwc", f"=wc_r*({cell('rev')}-{prev('rev')})")
            w("fcff", f"={cell('ebit')}*(1-tax_w)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("fat", f"={cell('fin')}*(1-tax_w)"); w("div", f"=pay_r*MAX(0,{cell('ni')})")
            w("fcfad", f"={cell('fcff')}+{cell('fat')}-{cell('div')}")
        for y in (2024, 2025): P[f"{COL[y]}{R['aeps']}"] = f"={COL[y]}{R['ani']}/(h_ni_{y % 100}/h_eps_{y % 100})"
        P.cell(R["u_g"], 2, "前一年 ×(1＋輕型車產量年增)×(1＋滲透／份額變化)")
        P.cell(R["base"], 2, "FY2026＝系統營收 − SuperVision；FY2027 起＝基礎出貨 × 平均售價")
        P.cell(R["drv"], 2, "線 1 新增車輛 × Mobileye 比例 × 美國以外倍數 × 每車系統價格")
        P.cell(R["agp"], 2, "FY2027 起：各子列 × 各自調整後毛利率")
        P.cell(R["op"], 2, "＝調整後營業利益 − 股票報酬 − 攤銷 − 商譽減損等")
        P.cell(R["fcff"], 2, "NOPAT＝(調整後營業利益 − 股票報酬)×(1 − tax_w)；股票報酬不加回")
        P.cell(R["tax"], 2, "FY2026＝9M 實績（Q4 設 0）；FY2027 起 tax_w × 稅前（虧損時為利益）")
        for k in ("gr", "agm", "aom", "gm", "opm", "nim", "nig"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(KEYS) + 2
    P.column_dimensions["A"].width = 52; P.column_dimensions["B"].width = 30
    for s in SC:
        for k in ("rev", "aoi", "ni", "eps", "fcff", "mbly", "fcfad", "aeps"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)；EBIT＝調整後營業利益 − 股票報酬（攤銷、商譽減損非現金，不扣）"
    D["A3"] = "ke（股權成本）＝rf＋β×(ERP＋CRP)"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（US$M；現價 × DCF 股數）"; D["B4"] = "=px*sh_dcf"
    D["A5"] = "負債 D（無借款；租賃負債未取得，不計）"; D["B5"] = 0
    D["A6"] = "稅後 kd（無負債，不適用）"; D["B6"] = 0
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金與非營業資產（2026-09-26）＝現金＋有價證券 − 借款 ＋ R&D Law 應收 × recv_k"; D["B10"] = "=cash_v+sec_v-0+recv_v*recv_k"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND(rf+beta*(erp+crp)+wacc_adj-wacc,9)"
    for rr_ in range(3, 11): D[f"B{rr_}"].number_format = "0.00%" if rr_ in (3, 6, 7, 8, 9) else "#,##0.0"
    DR = {}; rr = 13
    for s in SC:
        R = PR[s]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}" for y in range(2027, 2033)]); b = rr + 2
        rows = ["EBIT（調整後營業利益 − 股票報酬）", "NOPAT", "＋折舊", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF（應為 0）", "折現年數（期中）", "折現因子", "FCFF 現值"]
        for i, l in enumerate(rows): D.cell(b + i, 1, l)
        for j, y in enumerate(range(2027, 2033)):
            c = "BCDEFG"[j]; pc = COL[y]
            D[f"{c}{b}"] = f"=PL!{pc}{R['ebit']}"; D[f"{c}{b+1}"] = f"={c}{b}*(1-tax_w)"
            D[f"{c}{b+2}"] = f"=PL!{pc}{R['da']}"; D[f"{c}{b+3}"] = f"=-PL!{pc}{R['cx']}"; D[f"{c}{b+4}"] = f"=-PL!{pc}{R['dwc']}"
            D[f"{c}{b+5}"] = f"=PL!{pc}{R['fcff']}"; D[f"{c}{b+6}"] = f"=ROUND(SUM({c}{b+1}:{c}{b+4})-{c}{b+5},9)"
            D[f"{c}{b+7}"] = j + 0.5; D[f"{c}{b+8}"] = f"=1/(1+wacc)^{c}{b+7}"; D[f"{c}{b+9}"] = f"={c}{b+5}*{c}{b+8}"
        t = b + 11
        items = [("明確期現值合計", f"=SUM(B{b+9}:G{b+9})"),
                 ("NOPAT FY2033＝NOPAT FY2032 ×(1＋g)", f"=G{b+1}*(1+g_tv)"),
                 ("終值（FY2032 年底）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^6"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金與非營業資產（2026-09-26）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!F{R['fcff']}*stub_frac"),
                 ("（保留列：無待執行的買回承諾）", "=0"),
                 ("股權價值", f"=B{t+4}+B{t+5}+B{t+6}+B{t+7}"),
                 ("每股價值（評價日）", f"=MAX(0,B{t+8})/sh_dcf"),
                 ("DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）", f"=B{t+9}*(1+ke)^t_tgt"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+11}>=0,B{t+11}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=IF(G{b+1}>0,B{t+2}/G{b+1},\"N/M\")"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）", f"=MAX(0,(B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6)+B{t+5}+B{t+6}+B{t+7})/sh_dcf*(1+ke)^t_tgt")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 10, ps=t + 9, ev=t + 4, tvs=t + 11, alt=t + 14, chk=b + 6)
        dn(f"dcf_{s}", f"DCF!$B${t+10}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 64

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價：DCF（權重 w_dcf）＋ Comparable 同業前瞻本益比（1 − w_dcf；FY2028 調整後淨利 × 同業中位數）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}")
        V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "DCF 頁")
    cf = lambda s, k: f"PL!H{PR[s][k]}"
    r_cu = vrow("Comparable（未調整）：FY2028 調整後淨利 × 同業前瞻 P/E 中位數 ÷ 股數", [f"=({cf(s,'ani')}-cmp_sbc*{cf(s,'sbc')}*(1-tax_w))*pe_med/sh_cmp" for s in SC],
                "目標時點的 NTM 以 FY2028 代；cmp_sbc＝1 時扣稅後股票報酬（只作 Sens）")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（美元）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
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
    r_bnd = one("翻轉目標價：升一級／降一級", f'=TEXT(px*(1+IF(B{r_up}>0.5,9,IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5))))),"0.00")&"／"&TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,IF(B{r_up}>=-0.5,-0.5,-9))))),"0.00")', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 調整後 EPS", f"=B{r_tp}/PL!H{PR['B']['aeps']}", "分母：模型情境 B FY2028 調整後 EPS（Non-GAAP）")
    r_pe2 = one("現行前瞻 P/E（資料商 NTM 調整後，換算至 2026-11-20）÷ 同業中位數", "=pe_self/pe_med", "只作對照；第 47 條現行相對倍數")
    r_ev = one("現行 EV ÷ FY2026 調整後營業利益（倍；EV＝現價 × Comparable 股數 − 淨現金）", f"=(px*sh_cmp-cash_v-sec_v)/PL!F{PR['B']['aoi']}", "只作對照")
    r_ev2 = one("現行 EV ÷ FY2026（調整後營業利益 − 股票報酬）（倍）", f"=IF(PL!F{PR['B']['ebit']}>0,(px*sh_cmp-cash_v-sec_v)/PL!F{PR['B']['ebit']},\"N/M\")", "只作對照：扣股票報酬後的獲利口徑")
    for rr_ in (r_up, r_a1, r_a2, r_a3): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "0.00"
    for rr_ in (r_tp, r_dcft, r_cmpt, r_px): V[f"B{rr_}"].number_format = "0.00"
    V.column_dimensions["A"].width = 70; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, pd2=r_pd2, pe=r_pe, gj=r_gj, ev=r_ev, ev2=r_ev2, pe2=r_pe2)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司指引、市場共識、公司管線對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或共識", "差", "說明"])
    B = PR["B"]
    rc = [("FY2026 營收", f"=PL!F{B['rev']}", "=g26_rev", "公司指引中點（Interested-party）"),
          ("FY2026 調整後營業利益", f"=PL!F{B['aoi']}", "=g26_aoi", "公司指引中點"),
          ("FY2026 GAAP 營業利益", f"=PL!F{B['op']}", "=-(4127+4087)/2", "公司指引中點（GAAP 營業虧損 4,087–4,127）"),
          ("FY2026 股票報酬", f"=PL!F{B['sbc']}", "=g26_sbc", "公司指引"),
          ("FY2027–FY2033 營收合計（FY2033 以 FY2032 × (1＋FY2032 年增) 外推）vs 公司 8 年車用營收管線", f"=SUM(PL!G{B['rev']}:PL!L{B['rev']})+PL!L{B['rev']}*(1+PL!L{B['gr']})+PL!F{B['rev']}", "=pipe",
           "口徑不同：管線為 2025 年底起 8 年（FY2026–FY2033）預期車用營收、含未量產定點；模型含 FY2026 與 Moovit，只看方向"),
          ("FY2027 調整後 EPS（模型 B） vs 現價 ÷ 現行前瞻 P/E 隱含 NTM 調整後 EPS", f"=PL!G{B['aeps']}", "=px/pe_self", "資料商 NTM（第 3 季財報前，2026-10-26）；口徑：NTM 約 FY2027，只看方向"),
          ("加權目標價 vs 賣方平均目標價", "=tp", "=cons_tp", "賣方 24 位平均（第 3 季財報前，只作歷史對照）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_mobileye.py 二分法；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 60; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_mobileye.py 以 LibreOffice 重算每一案例）；結果為常數快照"
    base_tp = sens["base"]["tp"] if sens else None
    SE["A2"] = "快照基準值（加權，美元）"; SE["B2"] = base_tp
    SE["C2"] = "本檔即時值（加權，美元）"; SE["D2"] = "=tp"
    SE["E2"] = '=IF(ABS(D2-B2)>0.0005,"STALE：In 頁已改，Sens 須重跑","OK：快照與本檔一致")'
    SE["A3"] = "In 頁輸入指紋：快照"; SE["B3"] = sens["insum"] if sens else None
    SE["C3"] = "In 頁輸入指紋：即時（In!E 欄合計）"; SE["D3"] = "=SUM(In!E4:E400)"
    hdr(SE, 4, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    sr = 5
    if sens:
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]); SE.cell(sr, 5, cs["tp"] / base_tp - 1)
            SE.cell(sr, 6, cs["dcf"]); SE.cell(sr, 7, (cs["dcf"] / sens["base"]["dcf"] - 1) if sens["base"]["dcf"] else 0); SE.cell(sr, 8, cs["rt"])
            SE[f"E{sr}"].number_format = "0.0%"; SE[f"G{sr}"].number_format = "0.0%"; SE[f"D{sr}"].number_format = "0.00"; SE[f"F{sr}"].number_format = "0.00"
            sr += 1
    SE.cell(sr + 1, 1, "說明：相對倍數係數 rel_k 基準即為 1.00（未調整），「相對倍數係數＝1」即基準；另列「改用現行相對倍數」案例。pay_r（無股利）與 sh_26 等無區間的鍵不列案例。")
    SE.cell(sr + 2, 1, "build_id"); SE.cell(sr + 2, 4, "{BID}")
    SE.cell(sr + 3, 1, "模型狀態"); SE.cell(sr + 3, 4, '=IF(AND(ABS(D2-B2)<0.0005,ABS(D3-B3)<0.000001),"CURRENT","STALE")')
    SE.column_dimensions["A"].width = 50; SE.column_dimensions["B"].width = 34

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = "對帳 — 本模型使用或對照的「線層級」變數（取自線模型「下游取數」頁；與線模型不一致的自設值標明理由，不靜默改口徑）"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格／與線模型是否一致"])
    rec = [("美國 Robotaxi 年新增車輛需求 2027／2030／2035", "1 Robotaxi", "低 2,864／7,904／17,372；基準 5,332／30,730／68,688；高 7,503／58,536／115,734", "台", "擴車＋汰換（線模型）",
            LRT + "「下游取數」第 55–57 列 C／D／E 欄（RT_VEH_ADDS_2027／2030／2035）", "Derived", "In rt27_*／rt30_*／rt35_* → PL Drive 子列；一致（情境 A 讀高、B 讀基準、C 讀低，中間年線性內插）"),
           ("美國以外 Robotaxi 車輛", "1 Robotaxi", "以美國 × geo_m（2.0；1–3）", "倍", "線模型只有美國口徑", "線模型未建", "Assumed", "In geo_m；線模型無此變數（自設，理由：Mobileye 首批商轉在德國、挪威）"),
           ("Robotaxi 每車自駕系統價格", "1 Robotaxi", "4 萬美元（2.5–6 萬）", "US$/車", "公司說法", MB, "Interested-party", "In drv_asp；線模型的成本堆疊以營運商為主，未拆外購自駕系統價格（未比對）"),
           ("中國城市 NOA 滲透率 2027／2028／2030", "2 智慧車", "基準 29.8%／38.2%／55.0%（低 24.1%／29.4%／40.0%）", "%", "中國國內乘用車零售",
            LSC + "「下游取數」第 28–30 列 D 欄（低：C 欄）", "Derived", "只作對照、未進公式：Mobileye 主力在中國以外；高階產品（Surround＋SuperVision）占 EyeQ 出貨比見 PL，與中國滲透率的落差為判斷題"),
           ("中國 L3 滲透率 2028／2030", "2 智慧車", "基準 2%／4%", "%", "車廠擔責 L3", LSC + "「下游取數」第 34–35 列 D 欄", "Derived", "只作對照：Chauffeur（L3）2028 上半年上市（公司說法），本模型併入 SuperVision／Chauffeur 子列"),
           ("高速 NOA 方案／授權費", "2 智慧車", "線模型 25 美元（10–50），以 Mobileye 平均售價 49.1 為上緣", "US$/車", "中國高速 NOA 授權費",
            LSC + "「輸入」第 29 列 D–F 欄（lic_hwy；輸入值，非下游取數）", "Assumed", "口徑不同：本模型基礎 ADAS 平均售價 49.1 美元為晶片＋演算法整套（Verified），非中國高速 NOA 授權費；不改線模型"),
           ("Mobileye 平均系統售價與出貨（追蹤）", "2 智慧車", "Q3 2026 950 萬套、49.1 美元", "—", "公司 8-K", LSC + "「追蹤指標」第 15 列", "Verified", "In q3_asp、n9_u；一致"),
           ("全球輕型車產量年增", "2 智慧車", "1%（0–2.5%）", "%/年", "全球", "線模型未建（只有中國零售與 Tesla 交付）", "Assumed", "In lvp；線模型無此變數（自設）"),
           ("海外（歐美日）非 Tesla ADAS 每車內容", "2 智慧車", "未建（線模型限制欄）", "—", "—", "physical_ai/lines/smart_car/README.md「限制與未取得」", "—", "本模型以公司自揭出貨與單價代替；建議線 2 下一版補海外 ADAS 每車內容後回頭對帳")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 44; RE.column_dimensions["F"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("最大客戶轉自研或轉單（客戶 A 占營收約 28–30%）", "是（cust_loss，尾端 Sens，基準 0）", "原則第 10 節第 19 條", "10-Q 主要客戶占比、車廠自研公告"),
             ("Intel 持股（B 類股、投票權約 97%）出售或二次發行", "否", "影響股價供需，不影響營運現金流", "Intel 13D／424B 申報"),
             ("以色列地緣政治（營運、人員、折現率）", "部分（crp）", "國家風險溢酬 Sens", "公司營運公告、以色列主權評等"),
             ("R&D Law 補助取消或改制", "是（inc_k 低案 0）", "法律無到期日，但為新法", "以色列稅務當局公告、公司季報補助金額"),
             ("中國車廠自研與本土方案擠壓", "是（pen_C、asp_C）", "情境 C", "中國出口車占比、地平線出貨"),
             ("Robotaxi 無人駕駛（driver-out）里程碑延後", "是（dshr_C）", "情境 C；Drive 營收占比小", "漢堡、奧蘭多、奧斯陸無安全員營運時間表"),
             ("執行長交接（創辦人 2026-07 宣布卸任）", "否", "無可量化機制", "新執行長人選公告"),
             ("人形機器人（Mentee）研發支出", "部分（opex_g）", "營收未建（無量產時程）", "公司法說、人形機器人出貨")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 50

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([K3Q, Q3Q, Q2Q, K4Q, K4QH, NQ2, MB, SAI, SAS, SA + " Qualcomm、禾賽、Ambarella、地平線、DENSO 統計頁（2026-11-23 讀取）", YCL, RF, LSC, LRT], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"Mobileye（MBLY）個股模型 {VER}｜build_id {{BID}}"
    RD["A2"] = '="資料截止日 2026-11-23（財報與資料商）；成品日 2026-11-23；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.00")&" 美元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：Mobileye 的股價由三件事決定——基礎 ADAS（EyeQ 前視）的出貨與單價能否守住、Surround ADAS 與 SuperVision／Chauffeur 能否自 2027–2028 年拉高每車內容、以及扣除股票報酬後的利潤率能否擴張；Drive（Robotaxi）讀線 1 新增車輛，營收占比小"
    RD["A4"] = (f'="在 EyeQ 出貨每年相對車產 "&TEXT(pen_B,"+0.0%;-0.0%")&"、基礎 ADAS 單價每年 "&TEXT(asp_B,"+0.0%;-0.0%")&"、FY2032 Surround ADAS "&TEXT(srd_B,"0.0")&" 百萬套與股票報酬占營收 "&TEXT(sbc_r,"0.0%")&" 的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 美元（DCF "&TEXT(Val!B{r_dcft},"0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0")&" 百萬美元）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 美元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")&"。DCF（扣股票報酬）"&IF(Val!B{r_dcft}<px,"低於","高於")&"現價、Comparable（同業 non-GAAP 倍數）"&IF(Val!B{r_cmpt}<px,"低於","高於")&"現價：結論依賴 Comparable 採調整後 EPS 的口徑"'
    RD["A6"] = "失效條件（可觀察）：平均系統售價連兩季年減超過 5%；EyeQ 出貨年減（扣除安全庫存）；Surround ADAS 2028 年上市延後；R&D Law 補助取消；股票報酬占營收升破 20%"
    RD["A7"] = "範圍：不讀 Siliconomics／Tokenomics／Wattonomics（算力線）；晶片規格與份額屬 AI半導體（只讀）。人形機器人（Mentee）營收未建（無量產時程）。HTML 展示層見 dist/ 同名 .html"
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
LO = "file:///tmp/h_mobileye/lo_profile"
def recalc(src, outdir):
    os.makedirs(outdir, exist_ok=True)
    for attempt in range(2):
        try:
            subprocess.run(["soffice", f"-env:UserInstallation={LO}", "--headless", "--convert-to", "xlsx", "--outdir", outdir, src],
                           check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
            break
        except subprocess.TimeoutExpired:
            subprocess.run(["pkill", "-f", "lo_profile"], check=False)
            if attempt: raise
    return os.path.join(outdir, os.path.basename(src))

def read_vals(path, VR, DR=None):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]
    d = dict(tp=V[f"B{VR['tp']}"].value, dcf=V[f"B{VR['dcf']}"].value, cmp=V[f"B{VR['cmp']}"].value, up=V[f"B{VR['up']}"].value, rt=V[f"B{VR['rt']}"].value)
    d["insum"] = wb["Sens"]["D3"].value
    if DR:
        I = wb["In"]; w = {}; wd = None
        for row in I.iter_rows(values_only=True):
            if row[0] in ("w_A", "w_B", "w_C"): w[row[0][-1]] = row[4]
            if row[0] == "w_dcf": wd = row[4]
        alt = sum(w[s] * wb["DCF"][f"B{DR[s]['alt']}"].value for s in SC)
        d["alt_dcf"] = alt; d["alt_tp"] = wd * alt + (1 - wd) * d["cmp"]
    return d

def run_case(ov, tag):
    d = tempfile.mkdtemp(prefix="mbly_"); p = os.path.join(d, f"c_{tag}.xlsx")
    VR, _, DR = build(p, ov)
    out = recalc(p, os.path.join(d, "o"))
    v = read_vals(out, VR, DR if tag == "base" else None); shutil.rmtree(d, ignore_errors=True); return v

def in_val(k, col):
    for row in ROWS:
        if row[0] == k:
            return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)

def fingerprint():
    return hashlib.sha1(json.dumps([r for r in ROWS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]

def rating(u):
    return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"

CASES = [  # (鍵, 驅動)
    ("g_anc", "FY2026 實績相對指引"), ("lvp", "全球輕型車產量年增"), ("pen_B", "EyeQ 滲透／份額年變化（情境 B）"), ("asp_B", "基礎 ADAS 單價年變化（情境 B）"),
    ("srd_B", "Surround ADAS 出貨 FY2032（情境 B）"), ("srd_asp", "Surround ADAS 單價"), ("sv_B", "SuperVision／Chauffeur 出貨 FY2032（情境 B）"), ("sv_u26", "SuperVision 出貨 FY2026"),
    ("sv_asp", "SuperVision／Chauffeur 單價"), ("gm_sv", "SuperVision 毛利率"), ("gm_oth", "其他營收毛利率"), ("gm_drv", "Drive 毛利率"), ("drv_asp", "Drive 每車系統價格"),
    ("dshr_B", "Mobileye 占美國 Robotaxi 新增車輛比例（情境 B）"), ("geo_m", "美國以外 Drive 倍數"), ("mo_g", "Mobileye 其他營收年增"), ("oth_g", "Moovit 營收年增"),
    ("cust_loss", "最大客戶部分轉單（尾端）"), ("opex_g", "調整後營業費用年增"), ("inc_k", "R&D Law 補助延續"), ("sbc_r", "股票報酬 ÷ 營收"),
    ("fin_y", "財務收入率"), ("tax_w", "有效稅率"), ("da_r", "折舊 ÷ 營收"), ("cx_r", "資本支出 ÷ 營收"), ("wc_r", "營運資金 ÷ 營收增量"), ("recv_k", "R&D Law 應收計入比例"),
    ("crp", "國家風險溢酬"), ("beta", "β"), ("erp", "ERP"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重"),
    ("cmp_sbc", "Comparable 改用扣股票報酬 EPS")]

def write_readme(path, VR, PR, DR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f2 = lambda x: f"{x:,.2f}"; f0 = lambda x: f"{x:,.0f}"; pc = lambda x: f"{x*100:+.1f}%"
    inv = lambda k: I['E' + str([r for r in range(1, 400) if I['A' + str(r)].value == k][0])].value
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:8]
    def rvf(r): return r['val'] if isinstance(r['val'], str) else f"{r['val']:.4f}"
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {rvf(r)} | {r['base']:.4f} | {r['note']} |" for r in revs)
    def yr(s, k, fmt="{:,.0f}"):
        out = []
        for y in YRS:
            v = P[f"{COL[y]}{PR[s][k]}"].value
            out.append(fmt.format(v) if isinstance(v, (int, float)) else "—")
        return " | ".join(out)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f2(g(chr(66+i)+str(VR['dcfr'])))} | {f2(g(chr(66+i)+str(VR['cu'])))} | {f2(g(chr(66+i)+str(VR['tps'])))} |" for i, s in enumerate(SC))
    rts = sorted(set(c['rt'] for c in sens['cases']))
    B = PR["B"]
    txt = f"""# physical_ai/mobileye — Mobileye（MBLY，Nasdaq）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

本 README 的數字由 `build_mobileye.py` 從重算後的成品讀出產生，不手寫。所屬線：線 2 智慧車（ADAS／自駕系統方案）；Drive（Robotaxi）讀線 1。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

**一句話：** 以調整後（Non-GAAP）口徑，Mobileye FY2026 調整後營業利益約 {f0(P[f"F{B['aoi']}"].value)} 百萬美元、現行 EV 約為其 {g(f"B{VR['ev']}"):.1f} 倍；但股票報酬約 {f0(P[f"F{B['sbc']}"].value)} 百萬美元，扣除後的營業利益只有 {f0(P[f"F{B['ebit']}"].value)} 百萬美元（EV 倍數 {g(f"B{VR['ev2']}") if isinstance(g(f"B{VR['ev2']}"), str) else f"{g(f'B{VR[chr(101)+chr(118)+chr(50)]}'):.0f}"} 倍）。市場價格接近「股票報酬是真成本、成長有限」的讀法（DCF），賣方與同業倍數接近「調整後 EPS」的讀法（Comparable）；兩法差距就是這檔股票的爭議。調整後營業利益中約 {f0(P[f"F{B['inc']}"].value)} 百萬美元是 2026 年新的以色列 R&D Law 補助。

## 估值（評價日 2026-12-31；目標時點約 12 個月）
| 情境 | DCF（推到 12 個月後） | Comparable（同業前瞻 P/E） | 情境目標價 |
|---|---|---|---|
{scn}

- 加權目標價 **{f2(tp)} 美元**（DCF {f2(dcf)}、Comparable {f2(cmp_)}），現價 {f2(inv('px'))}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")}（情境 B FY2026–30 合計 {f0(g(f"C{VR['gate']}"))} 百萬美元）→ 投資建議 **{g(f"B{VR['rec']}")}**。
- 方法：DCF 標準版 FCFF（WACC／ke／g／RONIC {V['B' + str(VR['pe'] - 1)].value}；無負債，WACC＝ke；EBIT＝調整後營業利益 − 股票報酬）50%＋Comparable 50%（FY2028 調整後淨利 × 同業前瞻 P/E 中位數 {inv('pe_med'):.2f} 倍：Qualcomm、禾賽、Ambarella、DENSO）。相對倍數係數 1.00（未調整：缺 Mobileye 對同業的歷史相對倍數）；Mobileye 現行前瞻 P/E {inv('pe_self'):.2f} 倍、為同業中位數的 {g(f"B{VR['pe2']}"):.2f} 倍。
- 隱含前瞻 P/E（加權目標價 ÷ 情境 B FY2028 調整後 EPS）{g(f"B{VR['pe']}"):.1f} 倍。

## 營收分部（官方報告分部 Mobileye、Other；PL 頁三情境，情境 B 摘要，US$M）
| 列 | {" | ".join(f"FY{y}" for y in YRS)} |
|---|{"---|" * len(YRS)}
| 基礎 ADAS（子列） | {yr('B','base')} |
| Surround ADAS（子列） | {yr('B','srd')} |
| SuperVision／Chauffeur（子列） | {yr('B','sv')} |
| Drive（Robotaxi；子列） | {yr('B','drv')} |
| Mobileye 其他（子列） | {yr('B','mo')} |
| Mobileye 分部 | {yr('B','mbly')} |
| Other（Moovit） | {yr('B','oth')} |
| 合併營收 | {yr('B','rev')} |
| 調整後營業利益（Non-GAAP） | {yr('B','aoi')} |
| 股票報酬 | {yr('B','sbc')} |
| GAAP 營業利益 | {yr('B','op')} |
| GAAP 淨利 | {yr('B','ni')} |
| 調整後 EPS（美元） | {yr('B','aeps','{:,.2f}')} |
| EyeQ 系統出貨（百萬套） | {yr('B','u','{:,.1f}')} |

- FY2023–FY2025 實績（分部與子列未取得者為「—」）；FY2026＝前三季實績＋公司 2026-10-29 指引中點隱含的 Q4；FY2027 起為模型。
- 命題模組：EyeQ 出貨＝前一年 ×(1＋輕型車產量)×(1＋滲透／份額)；基礎 ADAS＝(出貨 − Surround − SuperVision) × 單價；Surround ADAS 2028 年起、SuperVision／Chauffeur 2027 年起線性爬到 FY2032 設定值；Drive＝線 1 美國 Robotaxi 新增車輛 × Mobileye 比例 × 美國以外倍數 × 4 萬美元。子列合計＝Mobileye 分部、分部合計＝合併營收，核對列為 0。
- 利潤：調整後毛利依子列毛利率（基礎／Surround 以 FY2026 校準 {inv('gm_core')*100:.1f}%）；調整後營業費用每年 +{inv('opex_g')*100:.0f}%；R&D Law 補助＝費用 × {inv('inc_r')*100:.1f}%；股票報酬＝營收 × {inv('sbc_r')*100:.1f}%（最近四季）。

## 承重假設與可證偽指標
| 驅動 | 標記 | 公開可觀察指標 |
|---|---|---|
| EyeQ 滲透／份額（pen） | Assumed | 季度系統出貨、前十大客戶出貨相對其產量 |
| 基礎 ADAS 單價（asp） | Assumed | 季度平均系統售價（Q3 2026 49.1 美元、年減 5%） |
| Surround ADAS 出貨（srd） | Assumed（公司說法為依據） | 2028 年三個定點上市時程、車廠銷量 |
| 股票報酬 ÷ 營收（sbc_r） | Derived（最近四季） | 季報股票報酬、RSU 授予量 |
| R&D Law 補助（inc_k） | Assumed | 季報補助金額、以色列法規 |
| 同業前瞻 P/E（pe_med） | Interested-party | 資料商前瞻 P/E（Ambarella、DENSO 報價日未更新） |

失效條件：{RD['A6'].value.split('：', 1)[1]}。Sens 各案例估值評等：{'、'.join(rts)}。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f2(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁。

## 反解（整本重算，二分法；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
- 線 1 Robotaxi：讀 `{LRT}`「下游取數」第 55–57 列 RT_VEH_ADDS_2027／2030／2035（情境 A 讀高、B 讀基準、C 讀低）→ Drive 子列。美國以外倍數 geo_m、每車系統價格 drv_asp 為本模型自設（線模型無此變數）。
- 線 2 智慧車：中國城市 NOA 滲透率（第 28–30 列）、中國 L3 滲透率（第 34–35 列）只作對照、未進公式（Mobileye 主力在中國以外，線模型未建海外非 Tesla ADAS 每車內容）；線模型 lic_hwy（25 美元）以 Mobileye 平均售價為上緣，與本模型 49.1 美元口徑不同（整套晶片＋演算法 vs 中國高速 NOA 授權費），不改線模型。
- 與線模型不一致的自設值：lvp（全球車產）、geo_m、drv_asp、pen／asp（Mobileye 自身份額與單價），理由見對帳頁。

## 限制與未取得資料
- FY2023 損益取自資料商（法定原文未逐項核對）；FY2023 調整後各列、FY2023–FY2025 分部營收與子列（基礎／SuperVision／REM）未取得；FY2025 Q1、Q2 平均系統售價未取得。
- Surround ADAS、SuperVision、Chauffeur、Drive 的單價只有公司在投資人會議的說法（Interested-party，經媒體轉述）；各產品毛利率、R&D Law 補助比率、無形資產未來攤銷時程、租賃負債、選擇權、2026 年 B 類股數未取得。
- 同業只有 4 家有前瞻 P/E（地平線虧損不列），Ambarella（2026-10-23）與 DENSO（2026-10-20）報價日未更新到 2026-11-20（csv 無舊報價日收盤，無法等比例換算），違反第 10 節第 7 條同日要求，列為限制；同業股票報酬占營收比未取得，Comparable 未就 Mobileye 股票報酬偏高做調整（另列 Sens）。
- 賣方共識目標價與前瞻 P/E 為第 3 季財報前（2026-10-26）資料，只作歷史對照；共識逐年 EPS 未取得。2026-11-20 股價單日 −10.2% 的原因未找到。
- 財務收入以評價日現金固定（未建現金累積與買回）；R&D Law 補助以費用減項入帳、現金收取有時差，未建時差。
- 情境 C 的營收下滑與費用年增（opex_g 三情境共用）並存、未建管理層降費反應，扣股票報酬後永續虧損，DCF 股權價值為負、每股取 0（含淨現金在內仍為負）；情境 C 的 DCF 0 為公式下限，不是清算價值。
- 人形機器人（Mentee Robotics，2026-02 以 6.37 億美元收購）只在費用中，營收未建。
- 未做：未參與製作的代理獨立核對（Project 規則第 10 條，待主代理安排）。

## 檔案
- `build_mobileye.py`：建置＋整本重算 Sens／反解＋產生本 README 與 `segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-23）：初版；現價 {f2(inv('px'))}（2026-11-20 收盤）；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


def main():
    if "--raw" in sys.argv:
        ov = json.loads(os.environ.get("MBLY_OVERRIDE", "{}"))
        build(sys.argv[sys.argv.index("--raw") + 1], ov); return
    base = run_case({}, "base"); print("基準", base); px0 = in_val("px", "base")
    cases = []
    beta_b = sorted([0.67 * 1.24 + 0.33, 0.67 * 1.76 + 0.33])
    beta_b = sum(beta_b) / 2
    only = os.environ.get("MBLY_QUICK")
    for k, drv in ([] if only else CASES):
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if k == "beta": v = beta_b + (-0.3 if side == "lo" else 0.3)
            if v is None: continue
            b_ = in_val(k, "base")
            if isinstance(b_, (int, float)) and abs(v - b_) < 1e-12: continue
            r = run_case({k: v}, f"{k}_{side}")
            cases.append(dict(drv=drv + ("（低）" if side == "lo" else "（高）"), set=f"{k}＝{v:.4g}", key=k, **r))
            print(k, side, round(r["tp"], 3), r["rt"])
    for s in SC:
        w = {f"w_{x}": (1.0 if x == s else 0.0) for x in SC}
        r = run_case(w, f"only_{s}")
        cases.append(dict(drv=f"只看情境 {s}", set=f"w_{s}＝1，其餘 0", key=f"w_{s}", **r))
    r = run_case({"rel_k": round(in_val("pe_self", "base") if False else 15.77 * px0 / 6.66 / 16.72, 4)}, "relk")
    cases.append(dict(drv="改用現行相對倍數（Mobileye 現行前瞻 P/E ÷ 同業中位數）", set="rel_k＝現行比值", key="rel_k", **r))
    up_ = base["alt_tp"] / px0 - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rating(up_), insum=base["insum"]))
    def solve(target_up, key, lo, hi, n=14):
        f = lambda x: run_case({key: x}, "rev")["up"] - target_up
        flo, fhi = f(lo), f(hi)
        if flo * fhi > 0: return None
        for _ in range(n):
            mid = (lo + hi) / 2; fm = f(mid)
            if flo * fm <= 0: hi, fhi = mid, fm
            else: lo, flo = mid, fm
        return (lo + hi) / 2
    revs = []
    up0 = base["up"]; tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    targets = [("目標價＝現價（上檔 0%）", 0.0)]
    if upper is not None: targets.append((f"升一級邊界（上檔 {upper:+.0%}）", upper))
    if lower is not None: targets.append((f"降一級邊界（上檔 {lower:+.0%}）", lower))
    for key, b0, lo_, hi_, note in (("pen_B", 0.01, -0.15, 0.15, "情境 B 的 EyeQ 出貨滲透／份額年變化（FY2027–FY2032 每年；情境 A／C 不動）"),
                                    ("sbc_r", None, 0.0, 0.40, "股票報酬 ÷ 營收（FY2027 起；三情境同動，影響 DCF；Comparable 用調整後 EPS 不受影響）"),
                                    ("pe_med", None, 2.0, 40.0, "同業前瞻 P/E 中位數（Comparable 倍數；其他不變）")):
        if b0 is None:
            b0 = {"sbc_r": (71 + 275) / (446 + 1558), "pe_med": 16.72}[key]
        for lab, tgt in targets:
            if os.environ.get("MBLY_NOREV"): break
            x = solve(tgt, key, lo_, hi_)
            revs.append(dict(label=lab, key=key, val=x if x is not None else f"區間 {lo_:+.2f}～{hi_:+.2f} 內無解", base=b0, note=note))
            print("反解", key, lab, x)
    sens = dict(base=base, cases=cases, fp=fingerprint(), insum=base["insum"])
    json.dump(sens, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="mblyf_"); raw = os.path.join(tmp, NAME)
    VR, PR, DR = build(raw, {}, sens, revs)
    bid = f"{VER}-{fingerprint()}"
    wb = openpyxl.load_workbook(raw)
    wb["README"]["A1"] = wb["README"]["A1"].value.replace("{BID}", bid)
    for row in wb["Sens"].iter_rows():
        for c in row:
            if c.value == "{BID}": c.value = bid
    wb.save(raw)
    out = recalc(raw, os.path.join(tmp, "o"))
    shutil.copy(out, os.path.join(dist, NAME))
    wbv = openpyxl.load_workbook(os.path.join(dist, NAME), data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10])
    assert not errs, errs
    write_readme(os.path.join(dist, NAME), VR, PR, DR, bid, sens, revs)
    json.dump(dict(company="mobileye", ticker="MBLY", model_sheet="Q",
                   official_segments=[dict(name=n, label=l, period="9M'26（2025-12-28～2026-09-26）", value=v, unit="US$M", source=Q3Q)
                                      for n, l, v in (("Mobileye", "Mobileye 分部營收", 1530), ("Other（Moovit）", "Other（Moovit）分部營收", 28))],
                   sublines=[], check_labels=["檢查：分部合計 − 合併營收（應為 0）", "檢查：分部績效合計 − 兩分部（應為 0）"],
                   note="Q 頁為 10-Q Note 9 官方分部（Q3'25、Q3'26、9M'25、9M'26）；PL 頁三情境另有「檢查：Mobileye 子列合計 − Mobileye 分部」與「檢查：分部合計 − 合併營收」（FY2026 起）。FY2023–FY2025 年度分部值未取得"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)

if __name__ == "__main__":
    main()

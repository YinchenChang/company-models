# -*- coding: utf-8 -*-
"""線 1a Robotaxi 單位經濟模型產生器（v0.2：以 v0.1 為底改版）。

python3 build_robotaxi.py  ->  dist/20261116_Robotaxi單位經濟_v0.2.xlsx
之後以 LibreOffice 重算，再跑 check_robotaxi.py。
所有輸入集中在 INPUTS；計算頁、情境頁、敏感度頁、下游取數頁全部是公式。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "dist", "20261116_Robotaxi單位經濟_v0.2.xlsx")
VERSION = "v0.2"
YEARS = list(range(2026, 2036))  # 2026..2035

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDR_FILL = PatternFill("solid", fgColor="D9E1F2")
YEL = PatternFill("solid", fgColor="FFF2CC")
SEC_FILL = PatternFill("solid", fgColor="E2EFDA")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------------------------------------------------------------------
# 來源
S = {
    "LYFT8K": "Lyft 8-K Q3 2026 新聞稿 https://www.sec.gov/Archives/edgar/data/0001759509/000162828026071043/lyft-20260930xpressrelease.htm",
    "LYFTQ1": "Lyft Q1 2026 法說逐字稿（Fool）https://www.fool.com/earnings/call-transcripts/2026/05/07/lyft-lyft-q1-2026-earnings-call-transcript/",
    "SM": "Bloomberg Second Measure（2024-03）https://secondmeasure.com/datapoints/rideshare-industry-overview/；William Blair 2026 Q2 調查 Uber 72%（最常用）https://www.williamblair.com/-/media/downloads/eqr/2026/williamblair-the-mobility-stack_second-quarter-2026-consumer-survey-insights.pdf",
    "GRID": "Gridwise Q3 2026 叫車與外送報告 https://gridwise.io/analytics/q3-2026-rideshare-delivery-report",
    "UBER": "Uber Q3 2026 新聞稿 https://seekingalpha.com/pr/20712391",
    "CPUC": "CPUC 季報（Waymo 加州月資料，Platform Aeronaut 整理）https://avrideshare.platformaeronaut.com/tracker/waymo/california",
    "FIND": "Abdelhalim (2026) Findings，CPUC 未遮蔽月資料 https://findingspress.org/article/161870-millions-of-trips-waymo-empty-miles-california-s-first-thousand-days-of-commercial-robotaxi-service",
    "WAYMO500": "Waymo 官方 500k+ 每週付費趟（2026-03，TechCrunch；Alphabet Q3 法說未更新）https://futuresearch.ai/app/p/a/waymo-weekly-paid-rides-q4-2026",
    "ELEC": "Electrek 2026-07-08 https://electrek.co/2026/07/08/waymo-driverless-las-vegas-four-new-cities/",
    "GOOG": "Alphabet Q3 2026 CEO 談話 https://blog.google/company-news/inside-google/message-ceo/alphabet-earnings-q3-2026/",
    "YIPIT": "YipitData（Business Insider 轉述，2026-06 訂單金額份額）https://letsdatascience.com/news/waymo-reaches-mid-teen-rideshare-share-but-driver-impact-rem-ae4d57e9",
    "OBI": "Obi 9.4 萬次叫車模擬（2025-11～2026-01，灣區）https://www.automotiveworld.com/news/obi-study-compares-tesla-waymo-uber-pricing/ ；https://techcrunch.com/2026/01/27/the-price-gap-between-waymo-and-uber-is-narrowing",
    "GS": "Goldman Sachs Research 2025-07-03 https://goldmansachs.com/insights/articles/autonomous-vehicle-market-forecast-to-grow-ridesharing-presence",
    "MS32": "Morgan Stanley（Sherwood 2025-12-09 轉述）https://www.sherwood.news/tech/where-morgan-stanley-thinks-autonomous-taxis-will-be-in-2032",
    "MS30": "Morgan Stanley（NewsBytes 2026-05-02 轉述）https://www.newsbytesapp.com/news/auto/us-robotaxis-set-to-expand-morgan-stanley-forecasts-750-million-rides/tldr",
    "TNW": "TNW（Morgan Stanley 估 Ojai 成本）https://thenextweb.com/news/waymo-ojai-robotaxi-geely-sixth-generation-fleet-cost",
    "TSLAQ3": "Tesla Q3 2026 法說整理 https://notateslaapp.com/news/4867/tesla-q3-2026-earnings-call-recap-everything-you-need-to-know",
    "BASENOR": "Basenor（Cybercab、奧斯汀牌價）https://www.basenor.com/blogs/news/tesla-cybercab-launches-in-austin-purpose-built-robotaxi-hits-the-road",
    "TXREG": "Tech Insider 彙整德州 AV 登記（TxDMV、EVwire）https://tech-insider.org/tesla-robotaxi-tally-739-trails-waymo-2026/",
    "ZOOX": "Axios 2026-08-05 https://axios.com/2026/08/05/zooxs-amazon-robotaxis-vegas ；TechCrunch 豁免 https://techcrunch.com/2026/07/30/zoox-clears-final-federal-hurdle-to-launch-paid-robotaxi-service/",
    "AAA": "AAA Your Driving Costs 2025（2025-09）https://newsroom.aaa.com/2025/09/aaa-new-vehicle-costs-drop-to-11577/",
    "FP": "Fehr & Peers（2018-09 研究、2019-08 發布；Uber、Lyft 委託）https://www.fehrandpeers.com/?p=823",
    "NBER": "Christensen, Nino & Osman, NBER w31330（2023-06，2025-09 修訂）https://www.nber.org/papers/w31330 ；全文 https://nber.org/system/files/working_papers/w31330/w31330.pdf",
    "FRED": "FHWA 經 FRED M12MTVUSM227NFWA（12 個月移動合計，2026-08）https://fred.stlouisfed.org/series/M12MTVUSM227NFWA",
    "MARKEY": "Markey 參議員遠端協助調查 https://www.markey.senate.gov/imo/media/doc/remote_assistance_investigation_report.pdf",
    # v0.2 新增
    "TCMARKEY": "TechCrunch 2026-03-31（Markey 調查：7 家皆未揭露遠端協助頻率；Waymo 約半數在菲律賓；Tesla 遠端可接手 ≤10 mph）https://techcrunch.com/2026/03/31/robotaxi-companies-refuse-to-say-how-often-their-avs-need-remote-help/",
    "TFARE": "Tesla 奧斯汀票價調整（2026-03-12 起 起跳 3.00＋每英里 1.40；之前 3.25＋1.00）https://www.tparts.com/blogs/tesla-latest-news/tesla-raises-robotaxi-fares-in-austin-second-price-hike-in-weeks-sees-mileage-fees-jump-40",
    "TCC": "Not a Tesla App 2026-09（Cybercab 開放叫車；同路線 Model Y 12.10 vs Cybercab 7.77 美元）https://notateslaapp.com/news/4651/cybercab-joins-teslas-robotaxi-fleet-with-cheaper-fares",
    "TSAFE": "Tesla Robotaxi 安全報告（2026-01-22～10-18 無監督 310 萬英里；0 受傷）Not a Tesla App 2026-10-21 https://notateslaapp.com/news/4860/tesla-names-next-robotaxi-cities-shares-miles-driven-safety-data",
    "ELECMAR": "Electrek 2026-03-31（奧斯汀車隊約 37–42 台，無監督 4–8 台；灣區約 130 台有安全員）https://electrek.co/2026/03/31/tesla-expands-unsupervised-robotaxi-service-area-still-only-handful-vehicles/",
    "ELECFEB": "Electrek 2026-02-17（NHTSA SGO：Tesla 14 起事故／約 80 萬英里，約每 5.7 萬英里 1 起，皆有安全員；Tesla 自報人類基準每 22.9 萬英里 1 起輕微碰撞）https://electrek.co/2026/02/17/tesla-robotaxi-adds-5-more-crashes-austin-month-4x-worse-than-humans/",
    "AW25": "Automotive World（Robotaxi Tracker：無監督 25 台，奧斯汀 19；Musk：大規模無監督待 FSD v15）https://www.automotiveworld.com/news/tesla-robotaxi-fleet-hits-25-as-musk-defers-scale-to-fsd-v15/",
    "TNW8": "TNW 2026-08-20（Robotaxi Tracker：奧斯汀兩週 170 趟 100% 無安全員、54 台車；達拉斯＋休士頓約 30 台；邁阿密 7 月移除安全員）https://thenextweb.com/news/tesla-robotaxi-austin-fully-driverless-bollards",
    "FSDT": "FSD Community Tracker（社群自願填報）彙整：Dominik Roth 2026-06 https://blog.dominik-roth.eu/posts/fsd-14.html ；原站 https://sites.google.com/view/fsdbetacommunitytracker",
    "CADMV": "CA DMV 2025 disengagement 報告（2024-12～2025-12，僅含有安全員的測試）Consumer Watchdog 2026-02-25 整理 https://consumerwatchdog.org/accountability/new-ca-dmv-numbers-show-uber-robotaxi-partner-nuro-could-not-go-700-miles-without-human-intervention/",
    "IIHS": "IIHS 2026-07-23（Waymo 警方可報事故率較人類低 68%，傷害事故低 81%；約 5,000 萬英里）https://iihs.org/news/detail/waymos-driverless-cars-crash-less-often-than-people",
    "SWISSRE": "Waymo × Swiss Re（2024-12）：2,530 萬英里 財損理賠 9 件（人類預期 78）、體傷 2 件（人類預期 26）https://waymo.com/blog/2024/12/new-swiss-re-study-waymo/",
    "CNCRT6": "CarNewsChina 2024-07（Apollo Go RT6 單車 204,600 元；武漢約 5 公里 7.5 元、Didi 11–15 元；尖峰每車日約 20 單；價格有補貼）https://carnewschina.com/2024/07/18/robotaxis-make-friends-and-enemies-on-the-streets-of-wuhan/",
    "CNCAPOLLO": "CarNewsChina 2025-11（Apollo Go 每週 25 萬+ 全無人單；李彥宏預估 2030 美國營運成本約 0.25 美元/英里）https://carnewschina.com/2025/11/13/baidus-apollo-go-robotaxi-leads-global-autonomous-driving-with-17m-orders-targets-profit-this-year/",
    "GASGOO": "Gasgoo 2026-05-21（百度六代單車略高於 20 萬元；小馬 2027 車型整車 <23 萬元、七代套件 BOM −70%；文遠 GXR 硬體約 4 萬美元；2026-03 武漢系統故障停運）https://autonews.gasgoo.com/articles/news/robotaxi-in-2026-real-momentum-but-profitability-remains-a-hurdle-2057454858058256385",
    "PONYUE": "CleanTechnica 2026-03-27（小馬七代深圳、廣州全城 UE 打平；含車輛與套件折舊、充電、維修）https://cleantechnica.com/2026/03/27/pony-ai-on-a-roll-uber-collab-breakeven-in-shenzhen-3000-robotaxi-target/",
    "SSENI": "SoftwareSeni（北京標準遠端安全員 1:3；Waymo 現行約 1:6 為未註出處估計）https://www.softwareseni.com/the-hidden-human-economy-inside-robotaxi-per-mile-costs",
    "OTHERBETS": "Fortune 2026-10-30（Alphabet Other Bets Q3 營收 3.72 億、營業虧損 20.5 億美元）https://fortune.com/2026/10/30/waymo-126-billion-alphabet-372-million-2-05-billion-google",
}

# ---------------------------------------------------------------------------
# 輸入：id, 變數, 單位, 低, 基準, 高, 標記, 用法, 出處, 日期, 備註
# 用法：共用＝三情境都用基準欄（低／高只進敏感度）；依情境＝低／基準／高情境各用一欄；參考＝不進計算，只供對照
INPUTS = [
    ("SEC", "A. 美國叫車市場基期"),
    ("M_LYFT_Q", "Lyft 單季總趟數（Q3 2026，含 FreeNow、單車）", "百萬趟", None, 273.5, None, "[Verified]", "共用", S["LYFT8K"], "2026-11", "Q3 2025 為 248.8，年增 10%"),
    ("M_LYFT_NA", "Lyft 趟數中北美汽車叫車占比", "%", 0.80, 0.85, 0.90, "[Assumed]", "共用", S["LYFTQ1"], "2026-05", "公司未揭露 FreeNow 與單車趟數；只說全年「超過 10 億趟」"),
    ("M_UBER_SPEND", "Uber 美國叫車消費份額（Uber＋Lyft＝100%）", "%", 0.72, 0.76, 0.80, "[Analogy]", "共用", S["SM"], "2024-03", "以 2024 消費份額類比 2026；調查最常用 Uber 72%、Lyft 16%"),
    ("M_FARE_UBER", "Uber 平均每趟車資（司機端觀察）", "US$/趟", None, 24.30, None, "[Verified]", "共用", S["GRID"], "2026 Q3", "Gridwise 司機樣本，含平台費"),
    ("M_FARE_LYFT", "Lyft 平均每趟車資", "US$/趟", None, 19.76, None, "[Verified]", "共用", S["GRID"], "2026 Q3", ""),
    ("M_TRIP_MI", "美國叫車平均每趟付費英里", "英里", 4.5, 6.0, 7.1, "[Assumed]", "共用", S["GRID"] + "；" + S["CPUC"], "—", "上緣 7.1＝Gridwise（車資−平台費）÷每英里司機收入 推得；下緣參考 Waymo 加州每趟約 3.7 載客英里（舊金山短程）"),
    ("M_GROWTH", "美國叫車趟數年成長（不含 Robotaxi 誘發）", "%/年", 0.02, 0.04, 0.06, "[Assumed]", "依情境", S["LYFT8K"] + "；" + S["UBER"], "—", "Lyft 趟數年增 10%（含 FreeNow）；Uber 全球趟數年增 16%（含外送）"),
    ("M_DRIVER_MI", "人類司機每英里毛收入（Uber）", "US$/英里", None, 2.65, None, "[Verified]", "共用", S["GRID"], "2026 Q3", "Robotaxi 可吃掉的人工成本上限；Lyft 2.87"),
    ("M_TAKE", "Uber 平台抽成中位數", "%", None, 0.2244, None, "[Verified]", "共用", S["GRID"], "2026 Q3", "用於經平台趟數的抽成"),

    ("SEC", "B. Robotaxi 現況校準（2026 年中）"),
    ("R_CA_SHARE", "Waymo 加州趟數占 Waymo 全美比例（2026 Q2）", "%", 0.45, 0.52, 0.60, "[Assumed]", "共用", S["CPUC"] + "；" + S["WAYMO500"], "2026 Q2", "上限 60%＝Q1 2026 加州每週約 30 萬 ÷ 官方 50 萬+（2026-03）；新城市成長較快，Q2 占比應更低"),
    ("R_OTHERS", "非 Waymo 營運商占美國 Robotaxi 趟數（2026）", "%", 0.03, 0.05, 0.08, "[Assumed]", "共用", S["TSLAQ3"] + "；" + S["ZOOX"], "2026", "Tesla 6 城無人、Zoox 拉斯維加斯收費、Avride 經 Uber；均未揭露趟數"),
    ("R_TESLA_2026", "其中 Tesla 占美國 Robotaxi 趟數（2026）", "%", 0.02, 0.03, 0.05, "[Assumed]", "共用", S["TXREG"], "2026-11", "德州登記 1,302 台（2026-11-10）但登記≠營運；須 ≤ R_OTHERS"),
    ("R_FLEET", "Waymo 車隊（2026 年中）", "台", 3000, 3500, 4000, "[Interested-party]", "共用", S["ELEC"], "2026-07", "媒體轉述，內文 3,000 與 3,500 不一致；含未載客車"),
    ("R_COV_2026", "商轉都會區占美國叫車趟數比例（2026）", "%", 0.25, 0.30, 0.35, "[Assumed]", "共用", S["GOOG"], "2026-10", "Waymo 15 城（缺紐約、芝加哥、波士頓、華府、西雅圖）；都會區層級，非服務區面積"),

    ("SEC", "C. 滲透路徑（依情境）"),
    ("S_COV_2028", "商轉都會覆蓋率 2028", "%", 0.35, 0.45, 0.55, "[Assumed]", "依情境", "—", "—", "2026→2028→2030→2032→2035 線性內插"),
    ("S_COV_2030", "商轉都會覆蓋率 2030", "%", 0.40, 0.60, 0.75, "[Assumed]", "依情境", "—", "—", ""),
    ("S_COV_2032", "商轉都會覆蓋率 2032", "%", 0.45, 0.70, 0.85, "[Assumed]", "依情境", "—", "—", ""),
    ("S_COV_2035", "商轉都會覆蓋率 2035", "%", 0.55, 0.80, 0.90, "[Assumed]", "依情境", "—", "—", ""),
    ("S_SMAX", "都會區內 Robotaxi 長期份額上限（趟數）", "%", 0.25, 0.50, 0.70, "[Assumed]", "依情境", S["YIPIT"], "2026-06", "對照：Waymo 舊金山、洛杉磯 15%、鳳凰城 16%（訂單金額，營運 2–3 年後）"),
    ("S_T0", "都會區份額 S 曲線中點年", "年", 2033, 2031, 2030.5, "[Assumed]", "依情境", "—", "—", "斜率 k 由 2026 校準點反推"),
    ("S_TPVD_2030", "每車每日付費趟數 2030", "趟/日", 18, 22, 26, "[Assumed]", "依情境", "—", "—", "2026 由 Waymo 每週趟數 ÷ 車隊推得；中國文遠 21、小馬 23（公司揭露）"),
    ("S_WAYMO_2030", "Waymo 占美國 Robotaxi 趟數 2030", "%", 0.65, 0.50, 0.40, "[Assumed]", "依情境", S["MS32"], "—", "對照 MS 2032：Waymo 38%、Tesla 29%（自駕里程）"),
    ("S_WAYMO_2035", "Waymo 占美國 Robotaxi 趟數 2035", "%", 0.55, 0.40, 0.30, "[Assumed]", "依情境", "—", "—", ""),
    ("S_TESLA_2030", "Tesla 占美國 Robotaxi 趟數 2030", "%", 0.15, 0.30, 0.45, "[Assumed]", "依情境", S["MS32"], "—", ""),
    ("S_TESLA_2035", "Tesla 占美國 Robotaxi 趟數 2035", "%", 0.20, 0.35, 0.50, "[Assumed]", "依情境", "—", "—", ""),
    ("R_PLAT_2026", "經 Uber／Lyft 平台的 Robotaxi 趟數占比 2026", "%", 0.05, 0.10, 0.15, "[Assumed]", "共用", S["UBER"], "2026", "Waymo 經 Uber（奧斯汀、亞特蘭大）、經 Lyft（納許維爾）"),
    ("S_PLAT_2030", "經平台占比 2030", "%", 0.20, 0.30, 0.40, "[Assumed]", "依情境", S["MS32"], "—", "MS：Uber／Lyft 約拿 30% AV 趟數；本欄低／高為數值區間，不代表滲透方向"),

    ("SEC", "D. 每英里實收價格與價格錨（核心指標）"),
    ("P_PREM_2026", "Robotaxi 相對人類叫車價格溢價 2026（Waymo vs Uber）", "%", None, 0.127, None, "[Verified]", "共用", S["OBI"], "2025-11～2026-01", "中程段 Waymo 每英里只貴約 2%（媒體寫「4.3–9.3 公里、每公里」；v0.2 判定 Obi 原始單位為英里，見 X_WAYMO_SF_MI 與 README 更正紀錄）；Obi 原文未給 Uber 每英里絕對值"),
    ("P_REL_2030", "Robotaxi 每英里價格 ÷ 人類叫車每英里價格 2030", "倍", 1.00, 0.85, 0.65, "[Assumed]", "依情境", S["TFARE"], "—", "Tesla 奧斯汀牌價 起跳 3.00＋每英里 1.40（2026-03-12 起；公司定價，補貼檢驗見「每英里價格」與「Tesla成本落差診斷」頁）"),
    ("P_REL_2035", "Robotaxi 每英里價格 ÷ 人類叫車每英里價格 2035", "倍", 0.90, 0.60, 0.35, "[Assumed]", "依情境", "—", "—", "2026→2030→2035 線性內插"),
    ("P_OWN_AVG", "自有新車每英里全成本（AAA 平均）", "US$/英里", None, 0.7718, None, "[Verified]", "共用", S["AAA"], "2025-09", "US$11,577/年 ÷ 15,000 英里（AAA 口徑，5 年持有、含折舊與融資）"),
    ("P_OWN_SEDAN", "自有小轎車每英里全成本（AAA）", "US$/英里", None, 0.5587, None, "[Verified]", "共用", S["AAA"], "2025-09", ""),
    ("P_ELAST", "叫車需求的總移動價格彈性", "—", -0.6, -1.0, -1.8, "[Analogy]", "依情境", S["NBER"], "2023-06／2025-09", "埃及 Uber 隨機折扣 3 個月：25% 折扣使用量翻倍以上、50% 折扣四倍以上；總移動彈性男 −0.60、女 −1.47（論文引言；5.1 節另報 −0.38／−1.32）；Uber 本身彈性 −4.4～−9.5 含替代其他運具，不用"),

    ("SEC", "E. 市場池（價格跨錨就打開的池）"),
    ("V_US_VMT", "美國全車種 12 個月 VMT", "十億英里", None, 3335.695, None, "[Verified]", "共用", S["FRED"], "2026-08", ""),
    ("V_LDV", "輕型乘用車占 VMT", "%", 0.85, 0.88, 0.91, "[Assumed]", "共用", "—", "—", "未取得 2025 FHWA 車種分拆；以經驗區間"),
    ("V_URBAN", "都會區占輕型車 VMT", "%", 0.65, 0.70, 0.75, "[Assumed]", "共用", "—", "—", "未取得最新 FHWA 都會／鄉村分拆"),
    ("V_TNC_VMT", "叫車占都會 VMT（六都會區）", "%", 0.01, 0.02, 0.03, "[Interested-party]", "參考", S["FP"], "2018-09", "Uber、Lyft 委託；核心郡 2–13%（舊金山最高）"),

    ("SEC", "F. 每英里成本（v0.2 起改讀「成本輸入」與「每英里成本堆疊」頁；此處只留 2032 後外推）"),
    ("S_DECL", "2032 後每英里成本年降幅", "%/年", 0.00, 0.03, 0.05, "[Assumed]", "依情境", "—", "—", "情境頁 2033–2035 由堆疊頁 2032 值外推"),

    ("SEC", "H. 參考值（不進計算，供對照）"),
    ("X_OJAI", "Waymo Ojai 單車成本", "US$", None, 125000, None, "[Interested-party]", "參考", S["TNW"], "2026", "Morgan Stanley 估計"),
    ("X_CYBERCAB", "Cybercab 目標售價", "US$", None, 30000, None, "[Interested-party]", "參考", S["BASENOR"], "2026", "Musk 口頭「低於 3 萬」；個人版 29,990"),
    ("X_TSLA_COST", "Tesla 每英里營運成本目標", "US$/英里", None, 0.30, None, "[Interested-party]", "參考", S["TSLAQ3"], "2026-10", "口徑未明"),
    ("X_TSLA_PRICE", "Tesla 奧斯汀每英里牌價（2026-03-12 起）", "US$/英里", None, 1.40, None, "[Interested-party]", "參考", S["TFARE"], "2026-03", "另加起跳 3.00；之前 1.00＋3.25"),
    ("X_TSLA_BASE", "Tesla 奧斯汀起跳價（2026-03-12 起）", "US$/趟", None, 3.00, None, "[Interested-party]", "參考", S["TFARE"], "2026-03", ""),
    ("X_WAYMO_SF_MI", "Waymo 灣區平均每英里價格（v0.1 誤植為每公里，v0.2 更正）", "US$/英里", None, 5.72, None, "[Verified]", "參考", S["OBI"], "2025-11～2026-01", "19.69÷5.72＝3.44 英里/趟，對應 CPUC 每趟約 3.4–3.7 付費英里；若為每公里則每趟僅 2.1 英里，與 CPUC 不符。Tesla 1.99 同理為每英里（8.17÷1.99＝4.1 英里）"),
    ("X_WAYMO_SF_TRIP", "Waymo 灣區平均每趟價格", "US$/趟", None, 19.69, None, "[Verified]", "參考", S["OBI"], "2025-11～2026-01", "Uber 17.47、Lyft 15.47"),
    ("X_TSLA_SF_MI", "Tesla 灣區平均每英里價格（有安全員）", "US$/英里", None, 1.99, None, "[Verified]", "參考", S["OBI"], "2025-11～2026-01", "每趟 8.17"),
    ("X_TSLA_SF_TRIP", "Tesla 灣區平均每趟價格", "US$/趟", None, 8.17, None, "[Verified]", "參考", S["OBI"], "2025-11～2026-01", ""),
    ("X_GS_INS", "GS 估 Robotaxi 保險每英里（2025）", "US$/英里", None, 0.50, None, "[Interested-party]", "參考", S["GS"], "2025-07", "對照本模型「預期理賠＋保費下限」口徑"),
    ("X_MS_COST", "MS 估 Waymo 每英里成本 2028", "US$/英里", None, 0.99, None, "[Interested-party]", "參考", S["MS32"], "2025-12", "由 1.43 降至 0.99"),
    ("X_GS_2030", "GS 估 2030 Robotaxi 占美國叫車份額", "%", None, 0.08, None, "[Interested-party]", "參考", S["GS"], "2025-07", "車隊約 3.5 萬台"),
    ("X_GS_FLEET", "GS 估 2030 美國 Robotaxi 車隊", "台", None, 35000, None, "[Interested-party]", "參考", S["GS"], "2025-07", ""),
    ("X_MS_RIDES", "MS 估 2030 美國 Robotaxi 年趟數", "百萬趟", None, 750, None, "[Interested-party]", "參考", S["MS30"], "2026-05", ""),
    ("X_MS_2032", "MS 估自駕占美國叫車里程（約 2032）", "%", None, 0.30, None, "[Interested-party]", "參考", S["MS32"], "2025-12", "年份為推定"),
]

# CPUC 月資料（Waymo 加州；+ 表示部分欄位遮蔽，為下限）
CPUC = [
    ("2026-06", 1348926, 2144678, 9023791.29, "+"), ("2026-05", 1454991, 2309049, 9704968.28, "+"),
    ("2026-04", 1416158, 2207679, 9414525.13, "+"), ("2026-03", 1363067, 1717304, 9105108.29, "+"),
    ("2026-02", 1220222, 1530644, 8191719.63, "+"), ("2026-01", 1322645, 1657530, 9092739.35, "+"),
    ("2025-12", 1257187, 1601574, 8420126.85, ""), ("2025-11", 1273473, 1650645, 8414101.98, ""),
    ("2025-10", 1217872, 1600676, 7902019.29, ""), ("2025-09", 1016546, 1369203, 6567950.55, ""),
    ("2025-08", 880643, 1265004, 5571862.80, ""), ("2025-07", 778812, 1122471, 4875510.04, ""),
    ("2025-06", 583262, 828029, 3567131.71, ""), ("2025-05", 875592, 1260448, 5336125.62, ""),
    ("2025-04", 776058, 1118391, 4784485.32, ""), ("2025-03", 708180, 1023028, 4265981.50, ""),
    ("2025-02", 559569, 779358, 3485921.73, ""), ("2025-01", 550457, 783182, 3481697.93, ""),
    ("2024-12", 541378, 814602, 3325573.16, ""), ("2024-06", 188847, 291644, 1117399.53, ""),
    ("2024-01", 77242, 98240, 426807.62, ""),
]

# ---------------------------------------------------------------------------
# v0.2 成本輸入（每英里成本堆疊的權威輸入）
# 欄位順序＝(2026, 2030, 2032) 三個錨點；每個錨點三值＝(低成本, 基準, 高成本)。
# 注意：低／高指「成本」高低，不是數值高低（例如壽命里程：低成本欄＝較長壽命）。
# 2028 由 2026 與 2030 平均（公式）。純量＝三值相同。v30／v32 為 None＝沿用前一錨點。
YRS4 = [2026, 2028, 2030, 2032]
SCS = ["L", "B", "H"]
OPS = [("W", "Waymo"), ("TY", "Tesla Model Y（現行）"), ("TC", "Tesla Cybercab"), ("CN", "中國（百度／小馬平均）"),
       ("TSLA", "Tesla 共用（學習曲線、研發）"), ("ALL", "全營運商共用")]
OPNAME = dict(OPS)
PNAME = {
    "VEH": ("車價（含感測器與自駕套件，營運商取得成本）", "US$/台"),
    "LIFE": ("單車壽命里程", "英里"),
    "PMD": ("每車每日付費英里（稼動率，承重）", "英里/日"),
    "PAID": ("付費英里占總行駛英里（1−空駛比）", "%"),
    "HRS": ("每車每日營運小時", "時/日"),
    "RATE": ("融資年利率（平均餘額＝車價×0.5）", "%/年"),
    "ENERGY": ("能源每車英里", "US$/英里"),
    "MAINT": ("維修清潔每車英里", "US$/英里"),
    "CLAIMS": ("責任理賠件數（每百萬車英里）", "件/百萬英里"),
    "SEV": ("每件平均理賠成本（含深口袋效應）", "US$/件"),
    "LOSSR": ("損失率（理賠÷保費）", "%"),
    "INS_FLOOR": ("商業保險保費下限（每車年）", "US$/車年"),
    "PHYS": ("自身車損、破壞與免賠額每車英里", "US$/英里"),
    "R": ("遠端協助次數（每千車英里）", "次/千英里"),
    "RA_MIN": ("每次遠端協助處理分鐘", "分鐘"),
    "RA_WAGE": ("遠端人員全負擔時薪", "US$/時"),
    "RA_CONC": ("每位遠端人員同時監看車數（concurrency，政策／保守值）", "台/人"),
    "MON": ("隨車安全員占營運車時比例", "%"),
    "MON_WAGE": ("隨車安全員全負擔時薪", "US$/時"),
    "CHASE": ("追車占營運車時比例", "%"),
    "DEPOT_FIX": ("每場站年固定成本（租金、24 小時人力、充電樁折舊）", "百萬美元/年"),
    "DEPOT_VEH": ("每場站服務車數", "台"),
    "CHG": ("充電與停車變動成本每車英里", "US$/英里"),
    "FLEET": ("營運車隊（年平均）", "台"),
    "RND": ("軟體、研發與地圖年費用（分攤給 Robotaxi）", "百萬美元/年"),
    "PLAT_SH": ("經第三方平台的付費英里占比", "%"),
    "TAKE": ("第三方平台抽成率", "%"),
    "PRICE": ("每付費英里實收價格（低欄＝低價、高欄＝高價）", "US$/英里"),
    "MPI0": ("Tesla 城市關鍵接管里程（FSD v14.3，社群）", "英里/次"),
    "T0": ("MPI0 對應時點", "年（小數）"),
    "G": ("關鍵接管里程年改善倍數（學習曲線斜率）", "倍/年"),
    "K_RA": ("遠端協助次數÷關鍵接管次數", "倍"),
    "P_CLAIM": ("每次關鍵接管（無人接手時）轉成責任理賠的機率", "件/次"),
    "CONC0": ("Tesla 2026 每位遠端人員監看車數", "台/人"),
    "BETA": ("監看車數對接管里程的彈性 β", "—"),
    "CONCMAX": ("每位遠端人員監看車數上限", "台/人"),
    "ATFAULT": ("SGO 申報事故中屬自身責任比例（校準用）", "%"),
    "OCC": ("遠端人員處理事件的目標占用率", "%"),
    "CHASE_HR": ("追車每小時車輛成本（不含人力）", "US$/時"),
}
PCT = {"PAID", "RATE", "LOSSR", "MON", "CHASE", "PLAT_SH", "TAKE", "ATFAULT", "OCC"}
A_ = "[Assumed]"; AN = "[Analogy]"; IP = "[Interested-party]"; V_ = "[Verified]"; D_ = "[Derived]"
COST = [
    # ---------------- Waymo
    ("W", "VEH", (120000, 140000, 170000), (60000, 80000, 110000), (50000, 70000, 100000), AN, S["TNW"], "2026", "Ojai 約 12.5 萬（MS 估）；I-PACE 約 15–20 萬；2030 後假設第六／七代與現代 Ioniq 5 平台降本"),
    ("W", "LIFE", (450000, 350000, 250000), (500000, 400000, 300000), None, A_, "—", "—", "未揭露；類比計程車車隊 30–50 萬英里"),
    ("W", "PMD", (115, 95, 75), (180, 150, 110), (200, 165, 120), D_, S["CPUC"] + "；" + S["ELEC"], "2026 Q2", "2026 基準＝每車日 25.5 趟（計算頁）×CPUC 每趟付費英里約 3.7；2030 後為假設"),
    ("W", "PAID", (0.57, 0.56, 0.55), (0.66, 0.62, 0.57), (0.68, 0.64, 0.58), A_, S["FIND"], "2025 H2", "2026 為 CPUC 實測（空駛 43–45%）；之後為假設"),
    ("W", "HRS", (18, 20, 22), None, None, A_, "—", "—", "24 小時營運扣充電與清潔；低成本欄＝同付費英里下較少人工時數"),
    ("W", "RATE", (0.05, 0.07, 0.09), None, None, A_, "—", "—", "Alphabet 資金成本＋車隊租賃利差"),
    ("W", "ENERGY", (0.05, 0.07, 0.10), (0.04, 0.06, 0.08), None, A_, S["AAA"], "—", "約 0.33–0.40 kWh/英里 × 商用電價 0.15–0.25"),
    ("W", "MAINT", (0.15, 0.22, 0.30), (0.10, 0.14, 0.20), (0.09, 0.12, 0.18), AN, "—", "—", "類比電動車隊＋感測器校正與每日清潔"),
    ("W", "CLAIMS", (0.30, 0.45, 1.00), (0.25, 0.40, 0.90), None, IP, S["SWISSRE"] + "；" + S["IIHS"], "2024-12／2026-07", "Swiss Re：(9+2)÷2,530 萬英里＝0.43 件/百萬英里（人類預期 4.1）；Waymo 委託研究"),
    ("W", "SEV", (15000, 25000, 45000), None, None, A_, "—", "—", "人類平均財損約 5–6 千、體傷約 2.5–3 萬美元，加深口袋倍數"),
    ("W", "LOSSR", (0.75, 0.65, 0.55), None, None, A_, "—", "—", "自保或專屬保險；0.65＝商業車險常見損失率"),
    ("W", "INS_FLOOR", (3000, 6000, 12000), None, None, A_, S["GS"], "—", "商業責任險法定與再保下限；GS 估每英里 0.50 為市場報價口徑"),
    ("W", "PHYS", (0.02, 0.04, 0.08), None, None, A_, "—", "—", "非自身責任事故未追回部分、破壞、感測器損壞"),
    ("W", "R", (0.5, 1.5, 4.0), (0.2, 0.6, 2.0), (0.1, 0.4, 1.5), A_, S["TCMARKEY"], "2026-03", "Waymo 拒絕揭露（商業機密），只說每英里求助次數「大幅下降」；CA DMV 測試 19,234 英里/次為有安全員測試的安全性接管，不含遠端協助"),
    ("W", "RA_MIN", (1.5, 3, 5), (1, 2, 4), None, A_, "—", "—", "未揭露"),
    ("W", "RA_WAGE", (20, 30, 45), None, None, A_, S["TCMARKEY"], "2026-03", "約半數在菲律賓"),
    ("W", "RA_CONC", (20, 10, 5), (60, 30, 15), (80, 40, 20), A_, S["GS"] + "；" + S["SSENI"], "—", "GS 估 2025 約 3:1、2030 約 10:1；SoftwareSeni 稱現行約 6:1（未註出處）"),
    ("W", "MON", (0, 0, 0.02), None, None, A_, S["GOOG"], "—", "商轉城市全無人；高欄保留新城市付費前導期"),
    ("W", "MON_WAGE", (25, 30, 40), None, None, A_, "—", "—", ""),
    ("W", "CHASE", (0, 0, 0), None, None, A_, "—", "—", "未觀察到"),
    ("W", "DEPOT_FIX", (2.0, 3.0, 4.5), None, None, A_, "—", "—", "灣區、洛杉磯租金與 24 小時清潔、充電人力"),
    ("W", "DEPOT_VEH", (400, 250, 150), (600, 400, 250), (700, 450, 280), A_, "—", "—", ""),
    ("W", "CHG", (0.02, 0.04, 0.06), None, None, A_, "—", "—", ""),
    ("W", "FLEET", (4000, 3500, 3000), (40000, 25000, 12000), (80000, 50000, 20000), A_, S["ELEC"] + "；" + S["MS32"], "2026-07", "2026 為媒體轉述；含海外"),
    ("W", "RND", (3000, 4000, 5500), (3500, 5000, 7000), None, A_, S["OTHERBETS"], "2026 Q3", "Other Bets 季虧 20.5 億（年化約 80 億，含營運與車輛折舊）；研發與地圖分攤數為假設"),
    ("W", "PLAT_SH", (0.05, 0.10, 0.15), (0.20, 0.30, 0.40), None, A_, S["UBER"] + "；" + S["MS32"], "—", "同輸入頁 R_PLAT_2026／S_PLAT_2030"),
    ("W", "TAKE", (0.20, 0.2244, 0.25), None, None, AN, S["GRID"], "2026 Q3", "類比人類司機抽成；Waymo 與 Uber 實際分潤未揭露"),
    ("W", "PRICE", (3.8, 4.3, 5.7), (2.0, 3.0, 4.0), (1.5, 2.6, 3.6), A_, S["OBI"], "2025-11～2026-01", "2026 高欄＝灣區 Obi 5.72/英里；基準＝人類叫車每英里×1.127（計算頁）"),
    # ---------------- Tesla Model Y
    ("TY", "VEH", (38000, 42000, 48000), (35000, 40000, 46000), None, A_, "—", "—", "Model Y 製造成本＋Robotaxi 改裝；純視覺、無光達"),
    ("TY", "LIFE", (400000, 300000, 250000), None, None, A_, "—", "—", ""),
    ("TY", "PMD", (70, 45, 30), (160, 130, 90), (180, 145, 100), A_, S["TSAFE"] + "；" + S["ELECMAR"], "2026-10", "車隊小、服務區小、營運至午夜；2026 推估：310 萬無監督英里÷270 日÷約 100–300 台＝每車日 38–115 車英里"),
    ("TY", "PAID", (0.55, 0.50, 0.40), (0.65, 0.60, 0.55), (0.66, 0.62, 0.56), A_, "—", "—", "未揭露；服務區小、車少時空駛比高"),
    ("TY", "HRS", (14, 16, 18), (18, 20, 22), None, A_, S["TSLAQ3"], "2026-10", "奧斯汀營運至午夜；2026 底目標 24 小時"),
    ("TY", "RATE", (0.05, 0.07, 0.09), None, None, A_, "—", "—", ""),
    ("TY", "ENERGY", (0.04, 0.05, 0.07), None, None, A_, "—", "—", "約 0.27 kWh/英里×0.15–0.25"),
    ("TY", "MAINT", (0.10, 0.15, 0.22), (0.08, 0.11, 0.16), None, AN, "—", "—", "類比電動車隊"),
    ("TY", "SEV", (8000, 15000, 30000), (12000, 20000, 35000), None, A_, S["ELECFEB"], "—", "2026 多為低速輕微事故"),
    ("TY", "LOSSR", (0.75, 0.65, 0.55), None, None, A_, "—", "—", ""),
    ("TY", "INS_FLOOR", (3000, 6000, 12000), None, None, A_, "—", "—", ""),
    ("TY", "PHYS", (0.02, 0.03, 0.06), None, None, A_, "—", "—", ""),
    ("TY", "RA_MIN", (2, 3, 5), (1, 2, 4), None, A_, "—", "—", "未揭露"),
    ("TY", "RA_WAGE", (30, 40, 50), None, None, A_, S["TCMARKEY"], "2026-03", "美國本地；遠端可直接接手低速駕駛"),
    ("TY", "MON", (0.30, 0.50, 0.70), (0, 0.02, 0.10), (0, 0, 0.05), A_, S["ELECMAR"] + "；" + S["TNW8"], "2026-03／2026-08", "2026 年平均：3 月奧斯汀 4–8／約 40 台無監督；8 月奧斯汀追蹤趟次 100% 無安全員；灣區約 130 台全有安全員"),
    ("TY", "MON_WAGE", (25, 30, 40), None, None, A_, "—", "—", ""),
    ("TY", "CHASE", (0, 0.02, 0.10), (0, 0, 0.02), None, A_, S["ELECMAR"], "2026-03", "早期使用追車，2026 看來已改遠端監督"),
    ("TY", "DEPOT_FIX", (1.5, 2.5, 4.0), None, None, A_, "—", "—", "Tesla 自有超充可降低部分成本"),
    ("TY", "DEPOT_VEH", (150, 60, 25), (500, 350, 200), (700, 450, 250), A_, S["AW25"], "2026-04", "2026 單城數十台，場站規模不經濟"),
    ("TY", "CHG", (0.01, 0.02, 0.04), None, None, A_, "—", "—", ""),
    ("TY", "FLEET", (500, 350, 200), (5000, 2000, 500), (3000, 1000, 300), A_, S["TXREG"] + "；" + S["ELECMAR"], "2026-11", "德州登記 Model Y 729 台（登記≠營運）；之後由 Cybercab 取代"),
    ("TY", "PLAT_SH", (0, 0, 0.05), (0, 0.05, 0.15), None, A_, "—", "—", "自有 App"),
    ("TY", "TAKE", (0.20, 0.2244, 0.25), None, None, AN, S["GRID"], "—", ""),
    ("TY", "PRICE", (1.70, 2.00, 2.60), (0.90, 1.30, 1.80), (0.80, 1.10, 1.60), A_, S["TFARE"], "2026-03", "2026 基準＝(3.00＋1.40×5)÷5 英里；低＝10 英里行程；高＝社群推測起跳 7 美元"),
    # ---------------- Tesla Cybercab
    ("TC", "VEH", (35000, 50000, 80000), (22000, 28000, 35000), (20000, 25000, 32000), A_, S["BASENOR"] + "；" + S["TSLAQ3"], "2026-10", "Musk「低於 3 萬」為量產目標；2026 小量爬坡成本較高"),
    ("TC", "LIFE", (500000, 400000, 300000), None, None, A_, "—", "—", "專用車設計壽命較長（未揭露）"),
    ("TC", "PMD", (60, 35, 25), (180, 150, 100), (200, 165, 110), A_, S["TCC"], "2026-09", "2026-09 才開放叫車"),
    ("TC", "PAID", (0.55, 0.50, 0.40), (0.68, 0.62, 0.55), (0.70, 0.64, 0.56), A_, "—", "—", ""),
    ("TC", "HRS", (14, 16, 18), (18, 20, 22), None, A_, S["TSLAQ3"], "—", ""),
    ("TC", "RATE", (0.05, 0.07, 0.09), None, None, A_, "—", "—", ""),
    ("TC", "ENERGY", (0.025, 0.035, 0.05), None, None, A_, "—", "—", "約 0.18–0.22 kWh/英里（未驗證）"),
    ("TC", "MAINT", (0.08, 0.12, 0.18), (0.05, 0.08, 0.12), (0.045, 0.07, 0.11), A_, "—", "—", ""),
    ("TC", "SEV", (8000, 15000, 30000), (12000, 20000, 35000), None, A_, "—", "—", ""),
    ("TC", "LOSSR", (0.75, 0.65, 0.55), None, None, A_, "—", "—", ""),
    ("TC", "INS_FLOOR", (3000, 6000, 12000), None, None, A_, "—", "—", ""),
    ("TC", "PHYS", (0.02, 0.03, 0.06), None, None, A_, "—", "—", ""),
    ("TC", "RA_MIN", (2, 3, 5), (1, 2, 4), None, A_, "—", "—", ""),
    ("TC", "RA_WAGE", (30, 40, 50), None, None, A_, "—", "—", ""),
    ("TC", "MON", (0, 0, 0.05), (0, 0, 0.03), (0, 0, 0.02), IP, S["TCC"], "2026-09", "無方向盤與踏板"),
    ("TC", "MON_WAGE", (25, 30, 40), None, None, A_, "—", "—", ""),
    ("TC", "CHASE", (0, 0.02, 0.10), (0, 0, 0.02), None, A_, "—", "—", ""),
    ("TC", "DEPOT_FIX", (1.5, 2.5, 4.0), None, None, A_, "—", "—", ""),
    ("TC", "DEPOT_VEH", (200, 80, 30), (600, 350, 200), (700, 450, 250), A_, S["TSLAQ3"], "2026-10", "奧斯汀 300+ 台，另有數百台在各城場站待命"),
    ("TC", "CHG", (0.01, 0.02, 0.04), None, None, A_, "—", "—", ""),
    ("TC", "FLEET", (300, 150, 80), (100000, 30000, 8000), (250000, 80000, 20000), A_, S["TSLAQ3"], "2026-10", "2026 年平均（9 月起營運）；2030 後為假設，對照情境頁 Tesla 所需車隊"),
    ("TC", "PLAT_SH", (0, 0, 0.05), (0, 0.05, 0.15), None, A_, "—", "—", ""),
    ("TC", "TAKE", (0.20, 0.2244, 0.25), None, None, AN, S["GRID"], "—", ""),
    ("TC", "PRICE", (1.10, 1.30, 1.60), (0.80, 1.10, 1.50), (0.60, 0.90, 1.30), A_, S["TCC"], "2026-09", "單一路線 Cybercab 較 Model Y 便宜約 35%"),
    # ---------------- 中國（百度／小馬平均，美元計，匯率 7.2）
    ("CN", "VEH", (27000, 30000, 40000), (16000, 20000, 26000), (14000, 18000, 24000), IP, S["CNCRT6"] + "；" + S["GASGOO"], "2024-07／2026-05", "RT6 204,600 元≈2.84 萬美元；小馬 2027 車型 <23 萬元；文遠 GXR 硬體約 4 萬美元"),
    ("CN", "LIFE", (450000, 350000, 250000), None, None, A_, "—", "—", ""),
    ("CN", "PMD", (130, 110, 80), (170, 140, 100), (190, 155, 110), A_, S["CNCRT6"], "2024-07", "尖峰每車日約 20 單；文遠 21、小馬 23 趟（v0.1 輸入頁）"),
    ("CN", "PAID", (0.65, 0.60, 0.50), (0.70, 0.65, 0.58), (0.72, 0.66, 0.60), A_, "—", "—", "未揭露"),
    ("CN", "HRS", (18, 20, 22), None, None, A_, "—", "—", ""),
    ("CN", "RATE", (0.03, 0.04, 0.06), None, None, A_, "—", "—", "人民幣融資"),
    ("CN", "ENERGY", (0.015, 0.025, 0.04), None, None, A_, "—", "—", ""),
    ("CN", "MAINT", (0.03, 0.05, 0.08), (0.02, 0.04, 0.06), None, A_, "—", "—", ""),
    ("CN", "CLAIMS", (0.5, 1.5, 4.0), (0.3, 1.0, 3.0), None, A_, S["CNCAPOLLO"], "2025-11", "百度只揭露每 1,014 萬公里 1 次氣囊彈出，無理賠率"),
    ("CN", "SEV", (3000, 6000, 12000), None, None, A_, "—", "—", ""),
    ("CN", "LOSSR", (0.75, 0.65, 0.55), None, None, A_, "—", "—", ""),
    ("CN", "INS_FLOOR", (500, 1000, 2000), None, None, A_, "—", "—", ""),
    ("CN", "PHYS", (0.005, 0.01, 0.02), None, None, A_, "—", "—", ""),
    ("CN", "R", (1.0, 3.0, 8.0), (0.4, 1.2, 4.0), (0.2, 0.8, 3.0), A_, "—", "—", "未揭露"),
    ("CN", "RA_MIN", (1.5, 3, 5), (1, 2, 4), None, A_, "—", "—", ""),
    ("CN", "RA_WAGE", (4, 6, 9), None, None, A_, "—", "—", ""),
    ("CN", "RA_CONC", (15, 8, 3), (40, 20, 6), (60, 30, 10), AN, S["SSENI"] + "；" + S["PONYUE"], "—", "北京標準 1:3（高成本欄）；小馬稱已精簡遠端比"),
    ("CN", "MON", (0, 0.10, 0.30), (0, 0.02, 0.10), (0, 0, 0.05), A_, S["CNCAPOLLO"], "2025-11", "核心區全無人，部分城市仍有安全員"),
    ("CN", "MON_WAGE", (5, 7, 10), None, None, A_, "—", "—", ""),
    ("CN", "CHASE", (0, 0, 0.05), None, None, A_, "—", "—", ""),
    ("CN", "DEPOT_FIX", (0.3, 0.5, 0.9), None, None, A_, "—", "—", ""),
    ("CN", "DEPOT_VEH", (400, 250, 150), (600, 400, 250), None, A_, "—", "—", ""),
    ("CN", "CHG", (0.005, 0.01, 0.02), None, None, A_, "—", "—", ""),
    ("CN", "FLEET", (6000, 5000, 4000), (60000, 30000, 12000), (120000, 60000, 25000), A_, S["GASGOO"], "2026", "小馬 2026-06 為 1,975 台（目標 3,500+）；GS 估 2026 中國約 1.4 萬台"),
    ("CN", "RND", (300, 500, 800), (400, 600, 900), None, A_, "—", "—", "百度＋小馬 Robotaxi 分攤"),
    ("CN", "PLAT_SH", (0.10, 0.20, 0.30), None, None, A_, "—", "—", ""),
    ("CN", "TAKE", (0.15, 0.20, 0.25), None, None, A_, "—", "—", ""),
    ("CN", "PRICE", (0.35, 0.50, 0.65), (0.30, 0.42, 0.55), (0.28, 0.38, 0.50), A_, S["CNCRT6"], "2024-07", "武漢 2024 例：約 1.5 元/公里（有補貼，≈0.34 美元/英里）；基準約 2.2 元/公里"),
    # ---------------- Tesla 共用
    ("TSLA", "MPI0", 853, None, None, IP, S["FSDT"], "2026-06", "社群自願填報，非隨機樣本；v14.3 城市 853（95% CI 750–955），全部 1,750"),
    ("TSLA", "T0", 2026.4, None, None, D_, S["FSDT"], "2026-06", "彙整截至 2026-05"),
    ("TSLA", "G", (2.6, 1.8, 1.25), None, None, A_, S["FSDT"], "—", "低成本欄＝歷史斜率（接管轉成本頁 v12.3→v14.3 推得約 2.6 倍/年）；v14.1→v14.3 曾退步"),
    ("TSLA", "K_RA", (2, 4, 8), None, None, A_, S["TCMARKEY"], "—", "未揭露；遠端協助含卡住、上下車點等非安全事件"),
    ("TSLA", "CONC0", (6, 3, 2), None, None, A_, S["ELECMAR"] + "；" + S["TCMARKEY"], "2026-03", "2026 每位遠端人員監看車數。Electrek：監督「看來已改為遠端」；Tesla 未揭露比例。最不確定的承重假設"),
    ("TSLA", "BETA", (0.8, 0.5, 0.3), None, None, A_, "—", "—", "監看比對接管里程的彈性：監看車數＝CONC0×(MPI_t÷MPI_2026)^β＝CONC0×G^(β×(年−2026))；把學習曲線接到遠端成本"),
    ("TSLA", "CONCMAX", (80, 40, 20), None, None, A_, S["GS"], "—", "監管或營運上限；GS 估 2030 約 10:1"),
    ("TSLA", "P_CLAIM", (0.004, 0.0075, 0.015), None, None, A_, S["ELECFEB"], "—", "以 2026 SGO 事故率校準（見接管轉成本頁）"),
    ("TSLA", "ATFAULT", (0.3, 0.5, 0.7), None, None, A_, "—", "—", "只用於校準檢核，不進成本"),
    ("TSLA", "RND", (300, 600, 1200), (400, 800, 1500), None, A_, "—", "—", "只含 Robotaxi 營運軟體、車隊管理、地圖；FSD 核心模型訓練算力歸算力收支，不重複計入"),
    # ---------------- 共用
    ("ALL", "OCC", (0.8, 0.6, 0.4), None, None, A_, "—", "—", "事件處理之外需留監看餘裕"),
    ("ALL", "CHASE_HR", (10, 15, 20), None, None, A_, "—", "—", ""),
]

wb = Workbook()
ws0 = wb.active
ws0.title = "說明"
wsI = wb.create_sheet("輸入")
wsK = wb.create_sheet("成本輸入")
wsD = wb.create_sheet("CPUC資料")
wsC = wb.create_sheet("計算")
wsX = wb.create_sheet("接管轉成本")
wsK2 = wb.create_sheet("每英里成本堆疊")
wsG = wb.create_sheet("Tesla成本落差診斷")
wsP = wb.create_sheet("每英里價格")
wsS = wb.create_sheet("敏感度")
wsN = wb.create_sheet("情境")
wsO = wb.create_sheet("下游取數")
wsT = wb.create_sheet("追蹤指標")


def hdr(ws, row, vals, widths=None):
    for j, v in enumerate(vals, 1):
        c = ws.cell(row=row, column=j, value=v)
        c.font = BOLD
        c.fill = HDR_FILL
        c.alignment = WRAP
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[CL(j)].width = w


# ---------------------------------------------------------------- 輸入
REF = {}  # id -> {"L": addr, "B": addr, "H": addr}
hdr(wsI, 1, ["ID", "變數", "單位", "低", "基準", "高", "標記", "用法", "出處", "日期", "備註"],
    [16, 44, 11, 11, 11, 11, 17, 8, 60, 14, 60])
r = 2
for it in INPUTS:
    if it[0] == "SEC":
        c = wsI.cell(row=r, column=1, value=it[1]); c.font = BOLD
        for j in range(1, 12):
            wsI.cell(row=r, column=j).fill = SEC_FILL
        r += 1
        continue
    iid, name, unit, lo, ba, hi, tag, use, src, date, note = it
    vals = [iid, name, unit, lo, ba, hi, tag, use, src, date, note]
    for j, v in enumerate(vals, 1):
        c = wsI.cell(row=r, column=j, value=v)
        c.alignment = WRAP
    judg = tag in ("[Assumed]", "[Analogy]")
    for j in (4, 5, 6):
        c = wsI.cell(row=r, column=j)
        if c.value is not None:
            c.font = BLUE
            if judg:
                c.fill = YEL
            if unit and unit.startswith("%"):
                c.number_format = "0.0%"
            elif isinstance(c.value, float) and abs(c.value) < 10:
                c.number_format = "0.00"
            else:
                c.number_format = "#,##0.0##"
    REF[iid] = {"L": f"'輸入'!$D${r}", "B": f"'輸入'!$E${r}", "H": f"'輸入'!$F${r}"}
    r += 1
wsI.freeze_panes = "C2"


def I(iid, col="B"):
    return REF[iid][col]


SC = {"低": "L", "基準": "B", "高": "H"}

# ---------------------------------------------------------------- 成本輸入（v0.2）
def dcol(y, s, base):
    return CL(base + YRS4.index(y) * 3 + SCS.index(s))


hdr(wsK, 1, ["ID", "營運商", "變數", "單位"] + [f"{y} {n}" for y in YRS4 for n in ("低成本", "基準", "高成本")]
    + ["標記", "出處", "日期", "備註"], [14, 20, 40, 12] + [10] * 12 + [16, 60, 14, 60])
wsK.cell(row=2, column=1, value="讀法：低／高＝成本低／高（不是數值低／高）。2028 欄為 2026 與 2030 平均（公式）。藍字＝輸入；黃底＝判斷。PRICE 列例外：低欄＝低價。").font = BOLD
KROW = {}
r = 3
cur_op = None
for (op, pid, v26, v30, v32, tag, src, date, note) in COST:
    if op != cur_op:
        c = wsK.cell(row=r, column=1, value=f"{op}｜{OPNAME[op]}"); c.font = BOLD
        for j in range(1, 21):
            wsK.cell(row=r, column=j).fill = SEC_FILL
        r += 1
        cur_op = op
    tri = lambda v: v if isinstance(v, tuple) else (v, v, v)
    a26 = tri(v26); a30 = tri(v30) if v30 is not None else a26; a32 = tri(v32) if v32 is not None else a30
    nm, unit = PNAME[pid]
    for j, v in enumerate([f"{op}_{pid}", OPNAME[op], nm, unit], 1):
        wsK.cell(row=r, column=j, value=v).alignment = WRAP
    judg = tag in (A_, AN)
    fmt = "0.0%" if pid in PCT else ("#,##0" if max(abs(x) for x in a26 + a30) >= 100 else "0.000")
    if pid in ("T0",):
        fmt = "0.0"
    for y, arr in ((2026, a26), (2030, a30), (2032, a32)):
        for s_, v in zip(SCS, arr):
            c = wsK.cell(row=r, column=5 + YRS4.index(y) * 3 + SCS.index(s_), value=v)
            c.font = BLUE; c.number_format = fmt
            if judg:
                c.fill = YEL
    for s_ in SCS:
        c26 = dcol(2026, s_, 5); c30 = dcol(2030, s_, 5)
        c = wsK.cell(row=r, column=5 + YRS4.index(2028) * 3 + SCS.index(s_), value=f"=({c26}{r}+{c30}{r})/2")
        c.number_format = fmt
    for j, v in zip((17, 18, 19, 20), (tag, src, date, note)):
        wsK.cell(row=r, column=j, value=v).alignment = WRAP
    KROW[(op, pid)] = r
    r += 1
wsK.freeze_panes = "E3"


def P(op, pid, y, s):
    """成本輸入儲存格參照。TY／TC 缺項時讀 TSLA，再讀 ALL。"""
    for o in (op, "TSLA", "ALL"):
        if (o, pid) in KROW:
            return f"'成本輸入'!${dcol(y, s, 5)}${KROW[(o, pid)]}"
    raise KeyError((op, pid))


OPS4 = ["W", "TY", "TC", "CN"]
OPID = {"W": "WAYMO", "TY": "TSLA_MY", "TC": "TSLA_CC", "CN": "CN"}
ITEMS = [("veh", "1 車輛（含感測器）攤提"), ("fin", "2 融資"), ("energy", "3 能源"), ("maint", "4 維修清潔"),
         ("ins", "5 保險與事故理賠"), ("remote", "6 遠端營運"), ("monitor", "7 安全員／追車（過渡）"), ("depot", "8 場站與充電")]
FULL_ITEMS = [("rnd", "9 軟體與研發攤提"), ("plat", "10 平台抽成")]


def tesla_mpi(p, y):
    return f"({p['MPI0']}*{p['G']}^({y}-{p['T0']}))"


def params(op, y, s):
    """回傳某營運商、年份、成本情境的參數字典（值為儲存格參照或運算式）。"""
    keys = ["VEH", "LIFE", "PMD", "PAID", "HRS", "RATE", "ENERGY", "MAINT", "SEV", "LOSSR", "INS_FLOOR", "PHYS",
            "RA_MIN", "RA_WAGE", "RA_CONC", "MON", "MON_WAGE", "CHASE", "DEPOT_FIX", "DEPOT_VEH", "CHG",
            "FLEET", "RND", "PLAT_SH", "TAKE", "OCC", "CHASE_HR"]
    if op in ("TY", "TC"):
        keys = [k for k in keys if k != "RA_CONC"]
    p = {k: P(op, k, y, s) for k in keys}
    p["PRICE"] = P(op, "PRICE", y, "B")  # 平台抽成一律以基準價格計
    if op in ("TY", "TC"):
        for k in ("MPI0", "T0", "G", "K_RA", "P_CLAIM", "CONC0", "BETA", "CONCMAX"):
            p[k] = P("TSLA", k, y, s)
        p["Y"] = str(y)
        p["FLEET_TY"] = P("TY", "FLEET", y, s); p["PMD_TY"] = P("TY", "PMD", y, s)
        p["FLEET_TC"] = P("TC", "FLEET", y, s); p["PMD_TC"] = P("TC", "PMD", y, s)
        if op == "TY":
            p["PMD_TY"] = None
        else:
            p["PMD_TC"] = None
    else:
        p["R"] = P(op, "R", y, s)
        p["CLAIMS"] = P(op, "CLAIMS", y, s)
    return p


def derived(p):
    """依參數字典補齊 R、CLAIMS、FLEETMI 運算式（敏感度與診斷換值後重算用）。"""
    q = dict(p)
    if "MPI0" in q:
        mpi = f"({q['MPI0']}*{q['G']}^({q['Y']}-{q['T0']}))"
        q.setdefault("R", f"({q['K_RA']}*1000/{mpi})")
        q.setdefault("CLAIMS", f"(1000000/{mpi}*{q['P_CLAIM']})")
        q.setdefault("RA_CONC", f"MIN({q['CONCMAX']},{q['CONC0']}*{q['G']}^({q['BETA']}*({q['Y']}-2026)))")
        pty = q["PMD_TY"] or q["PMD"]; ptc = q["PMD_TC"] or q["PMD"]
        q.setdefault("FLEETMI", f"(({q['FLEET_TY']}*{pty}+{q['FLEET_TC']}*{ptc})*365)")
    else:
        q.setdefault("FLEETMI", f"({q['FLEET']}*{q['PMD']}*365)")
    return q


def items(p):
    q = derived(p)
    vmd = f"({q['PMD']}/{q['PAID']})"
    evmin = f"({q['R']}/1000*{vmd}/{q['HRS']}*{q['RA_MIN']})"
    d = {
        "veh": f"{q['VEH']}/({q['LIFE']}*{q['PAID']})",
        "fin": f"{q['VEH']}*{q['RATE']}*0.5/({q['PMD']}*365)",
        "energy": f"{q['ENERGY']}/{q['PAID']}",
        "maint": f"{q['MAINT']}/{q['PAID']}",
        "ins": f"(MAX({q['CLAIMS']}*{q['SEV']}/1000000/{q['LOSSR']},{q['INS_FLOOR']}/({vmd}*365))+{q['PHYS']})/{q['PAID']}",
        "remote": f"{q['RA_WAGE']}*{q['HRS']}/(MIN({q['RA_CONC']},60*{q['OCC']}/{evmin})*{q['PMD']})",
        "monitor": f"({q['MON']}*{q['MON_WAGE']}+{q['CHASE']}*({q['MON_WAGE']}+{q['CHASE_HR']}))*{q['HRS']}/{q['PMD']}",
        "depot": f"{q['DEPOT_FIX']}*1000000/({q['DEPOT_VEH']}*{q['PMD']}*365)+{q['CHG']}/{q['PAID']}",
        "rnd": f"{q['RND']}*1000000/{q['FLEETMI']}",
        "plat": f"{q['PRICE']}*{q['PLAT_SH']}*{q['TAKE']}",
    }
    return d


def unit_expr(p):
    d = items(p)
    return "(" + "+".join(f"({d[k]})" for k, _ in ITEMS) + ")"


def full_expr(p):
    d = items(p)
    return "(" + "+".join(f"({d[k]})" for k, _ in ITEMS + FULL_ITEMS) + ")"

# ---------------------------------------------------------------- CPUC
hdr(wsD, 1, ["月份", "完成趟數", "乘客數", "總行駛英里", "遮蔽註記", "標記"], [10, 14, 14, 16, 10, 12])
for i, (m, t, p, mi, f) in enumerate(CPUC, 2):
    for j, v in enumerate([m, t, p, mi, f, "[Verified]"], 1):
        c = wsD.cell(row=i, column=j, value=v)
        if j in (2, 3, 4):
            c.font = BLUE
            c.number_format = "#,##0"
last = 1 + len(CPUC)
wsD.cell(row=last + 2, column=1, value="出處：" + S["CPUC"])
wsD.cell(row=last + 3, column=1, value="「+」＝部分欄位遮蔽，為下限。只含 Waymo 加州（舊金山灣區＋洛杉磯），不含其他州。")
q = last + 5
wsD.cell(row=q, column=1, value="季合計").font = BOLD
wsD.cell(row=q + 1, column=1, value="2026 Q2"); wsD.cell(row=q + 1, column=2, value="=SUM(B2:B4)"); wsD.cell(row=q + 1, column=4, value="=SUM(D2:D4)")
wsD.cell(row=q + 2, column=1, value="2026 Q1"); wsD.cell(row=q + 2, column=2, value="=SUM(B5:B7)"); wsD.cell(row=q + 2, column=4, value="=SUM(D5:D7)")
wsD.cell(row=q + 3, column=1, value="2025 Q2"); wsD.cell(row=q + 3, column=2, value="=SUM(B15:B17)"); wsD.cell(row=q + 3, column=4, value="=SUM(D15:D17)")
wsD.cell(row=q + 4, column=1, value="Q2 年增率"); wsD.cell(row=q + 4, column=2, value=f"=B{q+1}/B{q+3}-1")
for rr in range(q + 1, q + 4):
    for cc in (2, 4):
        wsD.cell(row=rr, column=cc).number_format = "#,##0"
wsD.cell(row=q + 4, column=2).number_format = "0%"
CPUC_Q2 = f"'CPUC資料'!$B${q+1}"
CPUC_Q1 = f"'CPUC資料'!$B${q+2}"
CPUC_Q2MI = f"'CPUC資料'!$D${q+1}"

# ---------------------------------------------------------------- 計算
hdr(wsC, 1, ["項目", "低", "基準", "高", "單位", "等式／說明"], [46, 14, 14, 14, 12, 80])
CR = {}
row = [2]


def sec(title):
    c = wsC.cell(row=row[0], column=1, value=title); c.font = BOLD
    for j in range(1, 7):
        wsC.cell(row=row[0], column=j).fill = SEC_FILL
    row[0] += 1


def line(key, label, fL, fB, fH, unit, note, fmt="#,##0.00"):
    r_ = row[0]
    wsC.cell(row=r_, column=1, value=label)
    for j, f in zip((2, 3, 4), (fL, fB, fH)):
        if f is not None:
            c = wsC.cell(row=r_, column=j, value=f)
            c.number_format = fmt
    wsC.cell(row=r_, column=5, value=unit)
    wsC.cell(row=r_, column=6, value=note).alignment = WRAP
    CR[key] = {"L": f"'計算'!$B${r_}", "B": f"'計算'!$C${r_}", "H": f"'計算'!$D${r_}"}
    row[0] += 1


def same(key, label, f, unit, note, fmt="#,##0.00"):
    line(key, label, None, f, None, unit, note, fmt)
    CR[key]["L"] = CR[key]["B"]
    CR[key]["H"] = CR[key]["B"]


def C(key, col="B"):
    return CR[key][col]


sec("A. 美國叫車市場基期（2026 年化）")
same("lyft_trip_share", "Lyft 美國趟數份額（由消費份額與車資換算）",
     f"=((1-{I('M_UBER_SPEND')})/{I('M_FARE_LYFT')})/((1-{I('M_UBER_SPEND')})/{I('M_FARE_LYFT')}+{I('M_UBER_SPEND')}/{I('M_FARE_UBER')})",
     "%", "趟數份額＝(消費份額÷每趟車資) 正規化；Lyft 車資較低，趟數份額高於消費份額", "0.0%")
same("us_trips", "美國叫車年趟數 2026", f"={I('M_LYFT_Q')}*4*{I('M_LYFT_NA')}/{C('lyft_trip_share')}",
     "百萬趟", "＝Lyft 單季趟數×4×北美汽車占比÷Lyft 趟數份額（含加拿大，未扣除）", "#,##0")
same("fare_w", "加權平均每趟車資", f"={C('lyft_trip_share')}*{I('M_FARE_LYFT')}+(1-{C('lyft_trip_share')})*{I('M_FARE_UBER')}",
     "US$/趟", "趟數份額加權")
same("p_human", "人類叫車每付費英里價格（價格錨 1）", f"={C('fare_w')}/{I('M_TRIP_MI')}", "US$/英里", "＝加權車資÷每趟付費英里")
same("us_miles", "美國叫車年付費英里 2026", f"={C('us_trips')}*{I('M_TRIP_MI')}", "百萬英里", "", "#,##0")
same("us_gb", "美國叫車年訂單金額 2026", f"={C('us_trips')}*{C('fare_w')}", "百萬美元", "", "#,##0")

sec("B. Robotaxi 現況校準（2026 年中）")
same("ca_weekly", "Waymo 加州每週趟數（2026 Q2）", f"={CPUC_Q2}/13", "趟/週", "CPUC Q2 合計÷13 週（下限）", "#,##0")
same("ca_weekly_q1", "Waymo 加州每週趟數（2026 Q1）", f"={CPUC_Q1}/13", "趟/週", "對照官方 50 萬+（2026-03）", "#,##0")
same("ca_share_cap", "加州占比上限（Q1 加州÷官方 50 萬）", f"={C('ca_weekly_q1')}/500000", "%", "官方數字為下限，故此為占比上限", "0.0%")
same("waymo_weekly", "Waymo 全美每週付費趟數（2026 Q2 推估）", f"={C('ca_weekly')}/{I('R_CA_SHARE')}", "趟/週", "＝加州每週÷加州占比", "#,##0")
same("rt_trips_2026", "美國 Robotaxi 年趟數（2026 年中年化）", f"={C('waymo_weekly')}*52/(1-{I('R_OTHERS')})/1000000", "百萬趟", "＝Waymo 每週×52÷(1−其他營運商占比)", "#,##0.0")
same("pen_2026", "美國 Robotaxi 滲透率（趟數）2026", f"={C('rt_trips_2026')}/{C('us_trips')}", "%", "＝Robotaxi 趟數÷美國叫車趟數", "0.00%")
same("s_2026", "商轉都會區內份額 2026", f"={C('pen_2026')}/{I('R_COV_2026')}", "%", "＝滲透率÷覆蓋率；對照 YipitData 舊金山／洛杉磯 15%（訂單金額）", "0.00%")
same("tpvd_2026", "Waymo 每車每日付費趟數 2026", f"={C('waymo_weekly')}/7/{I('R_FLEET')}", "趟/日", "＝每週趟數÷7÷車隊（車隊含非營運車，偏低）", "0.0")
same("ca_vmt_trip", "Waymo 加州每趟總行駛英里（含空駛）", f"={CPUC_Q2MI}/{CPUC_Q2}", "英里", "CPUC Q2；×付費占比≈每趟載客英里", "0.00")

sec("C. 滲透 S 曲線參數（依情境）")
line("k", "S 曲線斜率 k", *[f"=LN({I('S_SMAX',x)}/{C('s_2026')}-1)/({I('S_T0',x)}-2026)" for x in "LBH"],
     "1/年", "＝ln(Smax÷S2026−1)÷(t0−2026)，使 2026 曲線值等於校準點")

sec("D. 每英里實收價格（成本改讀「每英里成本堆疊」頁，見 F 節）")
same("p26", "2026 Robotaxi 每付費英里實收價格", f"={C('p_human')}*(1+{I('P_PREM_2026')})", "US$/英里", "＝人類叫車每英里×(1+溢價)")
line("p30", "2030 Robotaxi 每付費英里實收價格", *[f"={C('p_human')}*{I('P_REL_2030',x)}" for x in "LBH"], "US$/英里", "")
line("p35", "2035 Robotaxi 每付費英里實收價格", *[f"={C('p_human')}*{I('P_REL_2035',x)}" for x in "LBH"], "US$/英里", "")

sec("E. 傑文斯檢驗：價格錨、市場池、需求彈性")
same("pool_rh", "叫車池：美國叫車年付費英里", f"={C('us_miles')}/1000", "十億英里", "價格錨 1（人類叫車每英里）以下即在此池內替代", "#,##0.0")
same("pool_own", "自有車池：美國都會輕型車 VMT", f"={I('V_US_VMT')}*{I('V_LDV')}*{I('V_URBAN')}", "十億英里", "價格錨 2（自有車每英里全成本）以下才打開", "#,##0")
same("pool_ratio", "自有車池÷叫車池", f"={C('pool_own')}/{C('pool_rh')}", "倍", "", "#,##0")
same("rh_vmt_share", "叫車付費英里占都會輕型車 VMT", f"={C('pool_rh')}/{C('pool_own')}", "%", "對照 Fehr & Peers 六都會區 1–3%（含空駛、Uber/Lyft 委託）", "0.00%")
line("j_double", "傑文斯門檻價格：叫車英里翻倍所需每英里價格", *[f"={C('p_human')}*2^(1/{I('P_ELAST',x)})" for x in "LBH"],
     "US$/英里", "＝人類叫車每英里×2^(1/ε)；常數彈性下，價格降到此值時叫車池英里數為今日兩倍")
line("j_rev30", "2030 價格下總營收倍數（相對 2026 價格、同一批需求）", *[f"=({C('p30',x)}/{C('p26')})^(1+{I('P_ELAST',x)})" for x in "LBH"],
     "倍", "＝(P2030÷P2026)^(1+ε)；>1 表示降價使總營收擴大（|ε|>1 才會）", "0.00")
line("j_rev_flag", "降價是否擴大總營收", *[f"=IF({I('P_ELAST',x)}<-1,\"是（彈性大於 1）\",IF({I('P_ELAST',x)}=-1,\"持平（單位彈性）\",\"否（營收縮、英里增）\"))" for x in "LBH"],
     "—", "", "@")
same("own_gap_2026", "2026 Robotaxi 價格÷自有車每英里成本", f"={C('p26')}/{I('P_OWN_AVG')}", "倍", "", "0.0")

wsC.freeze_panes = "B2"

CONCL = [
    "以下文字依 v0.2 基準數字撰寫；改輸入後以本頁上方公式與排名為準。",
    "主因：稼動率低（車隊小、服務區小、營運時數短）。2026 Model Y 基準每車日付費英里約 45、空駛約 50%；安全員、遠端監看、場站都按「車時」或「車年」計費，被很少的付費英里分攤。只把稼動換成 2030 成熟值（每車日 150、空駛 38%）單位營運成本就從約 14.3 降到約 5.7 美元/英里；與其他因素平均分攤後仍是最大一項（約 5.1）。",
    "次因：隨車安全員（約 3.9）與遠端監看比低（約 2.3）。場站規模不經濟約 1.4。車型（Model Y→Cybercab）只約 0.3、保險與事故約 0.2：Tesla 低成本論點中最常被引用的 Cybercab 車價，不是目前成本高的原因。研發攤提只影響全成本（2026 每英里約 78 美元，來自極小車隊），隨車隊擴大消失。",
    "Tesla 0.30 美元的口徑：本模型 Cybercab 2032 低成本情境單位營運成本約 0.31，扣除車輛攤提與融資約 0.26。0.30 只有在 Cybercab、成熟稼動、低成本假設、且不含車輛與研發時才成立；基準情境 2030 約 1.19、2032 約 0.82。",
    "價格：Model Y 2026 每英里約 2.00（起跳 3.00＋1.40/英里，5 英里行程），只覆蓋單位營運成本約 14%，每車每日虧損約 550 美元。低價反映的是目標成本與試營運補貼，不是現行成本；2026-03 漲價 40% 是方向一致的訊號。",
    "何時可能消失：安全員在德州、佛州 2026 下半年已大致移除（灣區仍需 CPUC 無人許可）；稼動率要等車隊與服務區擴大，Musk 稱大規模無監督待 FSD v15（2026 底至 2027 初）；遠端監看比隨接管里程改善（基準每年 1.8 倍）在 2030 約每人 10 台。",
    "證偽：(1) 若 Tesla 或社群追蹤顯示每車日付費英里 >120 且無監督比例 >90%，單位成本仍 >3 美元/英里，則主因不是稼動，應改看遠端監看比與場站；(2) 若 2027 年底每車日付費英里仍 <60，稼動假說成立但消失時點延後，2030 跨 1 美元的判斷要下修；(3) 若社群關鍵接管里程連續兩版不升（如 v14.1→v14.3），學習曲線斜率改用高成本欄。",
]

# ---------------------------------------------------------------- 接管轉成本（v0.2）
for j, w in enumerate([46, 12] + [11] * 12 + [16, 50, 14, 50], 1):
    wsX.column_dimensions[CL(j)].width = w
wsX.cell(row=1, column=1, value="接管→成本：接管／遠端協助率 → 遠端營運成本；隨車安全員 → 人工成本；事故率 → 理賠。本頁是堆疊頁第 5、6、7 項的來源。").font = BOLD
XR = {}
xr = [3]


def xsec(t):
    c = wsX.cell(row=xr[0], column=1, value=t); c.font = BOLD
    for j in range(1, 19):
        wsX.cell(row=xr[0], column=j).fill = SEC_FILL
    xr[0] += 1


xsec("A1. FSD 社群接管里程（關鍵接管；監督版 FSD，非 Robotaxi 版）")
hdr(wsX, xr[0], ["版本", "約略時點（年）", "全部（英里/次）", "城市（英里/次）", "標記", "出處／備註"])
xr[0] += 1
FSD = [("12.3.x", 2024.25, 177, 121), ("12.5.x", 2024.65, 273, 129), ("13.2.x", 2024.95, 444, 217),
       ("14.1.x", 2025.80, 9263, 4109), ("14.2.x", 2025.95, 1642, 696), ("14.3.x", 2026.30, 1750, 853)]
FSDR = {}
for v, t, a, b in FSD:
    for j, val in enumerate([v, t, a, b, "[Interested-party]（社群）", S["FSDT"] + "；時點為依推送月份估計（[Assumed]）；14.1 僅搶先體驗、樣本小，視為離群值"], 1):
        c = wsX.cell(row=xr[0], column=j, value=val)
        if j in (2, 3, 4):
            c.font = BLUE; c.number_format = "0.00" if j == 2 else "#,##0"
    FSDR[v] = xr[0]
    xr[0] += 1
xr[0] += 1

xsec("A2. 接管、事故與遠端協助證據（依來源）")
hdr(wsX, xr[0], ["項目", "值", "單位", "標記", "出處", "日期", "備註"])
xr[0] += 1
EVID = [
    ("dmv_waymo", "Waymo CA DMV 每次接管英里（2025 報告年）", 19234, "英里/次", "[Verified]", S["CADMV"], "2026-02", "只含有安全員的測試里程（334.7 萬英里、174 次）；無人商轉不在此報告"),
    ("dmv_zoox", "Zoox CA DMV 每次接管英里", 60682, "英里/次", "[Verified]", S["CADMV"], "2026-02", ""),
    ("dmv_nuro", "Nuro CA DMV 每次接管英里", 646, "英里/次", "[Verified]", S["CADMV"], "2026-02", ""),
    ("t_crash_n", "Tesla Robotaxi NHTSA SGO 事故件數（2025-06～2026-01）", 14, "件", "[Verified]", S["ELECFEB"], "2026-02", "皆有安全員在車上"),
    ("t_crash_mi", "同期累計付費英里（Electrek 外推）", 800000, "英里", "[Interested-party]", S["ELECFEB"], "2026-01", ""),
    ("t_human", "Tesla 自報人類基準：每次輕微碰撞英里", 229000, "英里/次", "[Interested-party]", S["ELECFEB"], "—", ""),
    ("nhtsa_pr", "NHTSA 警方報案事故平均間隔（Electrek 引用）", 500000, "英里/次", "[Interested-party]", S["ELECFEB"], "—", ""),
    ("sr_claims", "Waymo 責任理賠件數（Swiss Re，財損 9＋體傷 2）", 11, "件", "[Interested-party]", S["SWISSRE"], "2024-12", "Waymo 委託"),
    ("sr_miles", "Swiss Re 研究里程", 25.3, "百萬英里", "[Interested-party]", S["SWISSRE"], "2024-12", ""),
    ("sr_human", "同里程人類預期理賠（78＋26）", 104, "件", "[Interested-party]", S["SWISSRE"], "2024-12", ""),
    ("iihs", "IIHS：Waymo 警方可報事故率相對人類", -0.68, "%", "[Verified]", S["IIHS"], "2026-07", "傷害事故 −81%；奧斯汀樣本小 +4%"),
    ("t_unsup", "Tesla 無監督里程（2026-01-22～10-18），0 受傷", 3.1, "百萬英里", "[Interested-party]", S["TSAFE"], "2026-10", "Tesla 未揭露 SGO 輕微事故件數"),
    ("bj_ratio", "北京標準：遠端安全員對車輛", 3, "台/人", "[Analogy]", S["SSENI"], "—", "二手轉述"),
]
EVR = {}
for k, nm, v, u, tag, src, dt, note in EVID:
    for j, val in enumerate([nm, v, u, tag, src, dt, note], 1):
        c = wsX.cell(row=xr[0], column=j, value=val); c.alignment = WRAP
        if j == 2:
            c.font = BLUE; c.number_format = "0%" if u == "%" else "#,##0.0##"
    EVR[k] = f"'接管轉成本'!$B${xr[0]}"
    xr[0] += 1
wsX.cell(row=xr[0], column=1, value="遠端協助頻率：Markey 調查 7 家（Aurora、May、Motional、Nuro、Tesla、Waymo、Zoox）皆未直接回答；Waymo、May 稱商業機密，Tesla 未作答。故 R（每千英里遠端協助）全為 [Assumed]。" + S["TCMARKEY"]).alignment = WRAP
xr[0] += 2

xsec("B. 學習曲線：把 FSD 版本別接管里程視為學習曲線")


def xline(key, label, f, unit, note, fmt="#,##0.00"):
    wsX.cell(row=xr[0], column=1, value=label)
    c = wsX.cell(row=xr[0], column=2, value=f); c.number_format = fmt
    wsX.cell(row=xr[0], column=3, value=unit)
    wsX.cell(row=xr[0], column=5, value=note)
    XR[key] = f"'接管轉成本'!$B${xr[0]}"
    xr[0] += 1


xline("hist_g_city", "歷史年改善倍數（城市，v12.3→v14.3）",
      f"=(D{FSDR['14.3.x']}/D{FSDR['12.3.x']})^(1/(B{FSDR['14.3.x']}-B{FSDR['12.3.x']}))", "倍/年", "對照成本輸入 TSLA_G 低成本欄")
xline("hist_g_all", "歷史年改善倍數（全部，v12.3→v14.3）",
      f"=(C{FSDR['14.3.x']}/C{FSDR['12.3.x']})^(1/(B{FSDR['14.3.x']}-B{FSDR['12.3.x']}))", "倍/年", "")
xline("recent_g", "近期年改善倍數（城市，v13.2→v14.3）",
      f"=(D{FSDR['14.3.x']}/D{FSDR['13.2.x']})^(1/(B{FSDR['14.3.x']}-B{FSDR['13.2.x']}))", "倍/年", "")
xline("t_sgo_rate", "Tesla Robotaxi SGO 事故率（有安全員）", f"={EVR['t_crash_n']}/{EVR['t_crash_mi']}*1000000", "件/百萬英里", "校準 P_CLAIM 用", "0.0")
xline("sr_rate", "Waymo 責任理賠率（Swiss Re）", f"={EVR['sr_claims']}/{EVR['sr_miles']}", "件/百萬英里", "成本輸入 W_CLAIMS 基準 0.45", "0.00")
xline("sr_human_rate", "人類預期責任理賠率（同研究）", f"={EVR['sr_human']}/{EVR['sr_miles']}", "件/百萬英里", "", "0.00")
for s_, nm in zip(SCS, ("低成本（快）", "基準", "高成本（慢）")):
    xline(f"x_dmv_{s_}", f"Tesla 城市關鍵接管里程達 Waymo CA DMV 水準的年份｜{nm}",
          f"={P('TSLA','T0',2026,s_)}+LN({EVR['dmv_waymo']}/{P('TSLA','MPI0',2026,s_)})/LN({P('TSLA','G',2026,s_)})", "年",
          "定義不同（社群關鍵接管 vs DMV 測試接管），只作量級對照", "0.0")
xr[0] += 1

xsec("C. Tesla 接管里程投影 → 遠端協助率、理賠率（欄＝年份×成本情境）")
hdr(wsX, xr[0], ["項目", "單位"] + [f"{y} {n}" for y in YRS4 for n in ("低成本", "基準", "高成本")] + ["說明"])
xr[0] += 1


def xrow(key, label, unit, fn, note="", fmt="#,##0.00"):
    wsX.cell(row=xr[0], column=1, value=label)
    wsX.cell(row=xr[0], column=2, value=unit)
    for y in YRS4:
        for s_ in SCS:
            col = dcol(y, s_, 3)
            c = wsX.cell(row=xr[0], column=3 + YRS4.index(y) * 3 + SCS.index(s_), value="=" + fn(y, s_))
            c.number_format = fmt
            XR[(key, y, s_)] = f"'接管轉成本'!${col}${xr[0]}"
    wsX.cell(row=xr[0], column=15, value=note).alignment = WRAP
    xr[0] += 1


xrow("t_mpi", "Tesla 城市關鍵接管里程", "英里/次", lambda y, s_: f"{P('TSLA','MPI0',y,s_)}*{P('TSLA','G',y,s_)}^({y}-{P('TSLA','T0',y,s_)})",
     "＝MPI0×G^(年−T0)", "#,##0")
xrow("t_r", "Tesla 遠端協助次數（每千車英里）", "次/千英里", lambda y, s_: f"{P('TSLA','K_RA',y,s_)}*1000/{XR[('t_mpi',y,s_)]}", "＝K×1000÷MPI")
xrow("t_claims", "Tesla 責任理賠件數（每百萬車英里）", "件/百萬英里", lambda y, s_: f"1000000/{XR[('t_mpi',y,s_)]}*{P('TSLA','P_CLAIM',y,s_)}", "＝關鍵接管次數×P_CLAIM（無人接手時）")
xrow("t_sgo_impl", "隱含 SGO 事故率（＝理賠率÷自身責任比例）", "件/百萬英里", lambda y, s_: f"{XR[('t_claims',y,s_)]}/{P('TSLA','ATFAULT',y,s_)}",
     "2026 基準對照觀察值（B 節 Tesla SGO 事故率）", "0.0")
xr[0] += 1

xsec("D. 每營運商：接管 → 遠端營運成本、安全員成本、理賠成本（每付費英里）")
hdr(wsX, xr[0], ["項目", "單位"] + [f"{y} {n}" for y in YRS4 for n in ("低成本", "基準", "高成本")] + ["說明"])
xr[0] += 1
for op in OPS4:
    c = wsX.cell(row=xr[0], column=1, value=OPNAME[op]); c.font = BOLD
    xr[0] += 1
    k = lambda nm: f"{op}_{nm}"
    if op in ("TY", "TC"):
        xrow(k("r"), "遠端協助次數（每千車英里）", "次/千英里", lambda y, s_: XR[("t_r", y, s_)], "讀 C 節學習曲線")
        xrow(k("claims"), "責任理賠件數（每百萬車英里）", "件/百萬英里", lambda y, s_: XR[("t_claims", y, s_)], "讀 C 節")
    else:
        xrow(k("r"), "遠端協助次數（每千車英里）", "次/千英里", lambda y, s_: P(op, "R", y, s_), "成本輸入（未揭露，假設）")
        xrow(k("claims"), "責任理賠件數（每百萬車英里）", "件/百萬英里", lambda y, s_: P(op, "CLAIMS", y, s_), "成本輸入")
    xrow(k("mph"), "每車每營運小時行駛英里", "英里/時", lambda y, s_: f"{P(op,'PMD',y,s_)}/{P(op,'PAID',y,s_)}/{P(op,'HRS',y,s_)}", "＝(付費英里÷付費占比)÷營運小時；稼動率低＝每小時英里少")
    xrow(k("evmin"), "每車每小時事件處理分鐘", "分鐘/時", lambda y, s_: f"{XR[(k('r'),y,s_)]}/1000*{XR[(k('mph'),y,s_)]}*{P(op,'RA_MIN',y,s_)}", "", "0.000")
    xrow(k("cap"), "事件負載下每人可監看車數上限", "台/人", lambda y, s_: f"60*{P('ALL','OCC',y,s_)}/{XR[(k('evmin'),y,s_)]}", "＝60×占用率÷每車每時事件分鐘", "#,##0")
    xrow(k("pol"), "政策／保守監看車數" + ("（Tesla＝CONC0×G^(β×(年−2026))，上限 CONCMAX）" if op in ("TY", "TC") else "（成本輸入）"), "台/人",
         lambda y, s_: derived(params(op, y, s_))["RA_CONC"], "Tesla 由接管學習曲線推得：接管越少、可同時監看越多車", "0.0")
    xrow(k("conc"), "實際每人監看車數＝MIN(政策值, 事件上限)", "台/人", lambda y, s_: f"MIN({XR[(k('pol'),y,s_)]},{XR[(k('cap'),y,s_)]})", "政策值通常小於事件上限＝成本由監看比決定", "0.0")
    xrow(k("remote"), "遠端營運成本／付費英里", "US$/英里", lambda y, s_: f"{P(op,'RA_WAGE',y,s_)}*{P(op,'HRS',y,s_)}/({XR[(k('conc'),y,s_)]}*{P(op,'PMD',y,s_)})",
         "＝時薪×營運小時÷(每人車數×每車每日付費英里)")
    xrow(k("monitor"), "隨車安全員＋追車成本／付費英里", "US$/英里",
         lambda y, s_: f"({P(op,'MON',y,s_)}*{P(op,'MON_WAGE',y,s_)}+{P(op,'CHASE',y,s_)}*({P(op,'MON_WAGE',y,s_)}+{P('ALL','CHASE_HR',y,s_)}))*{P(op,'HRS',y,s_)}/{P(op,'PMD',y,s_)}",
         "安全員在車上＝人類司機成本（過渡）")
    xrow(k("eloss"), "預期理賠／車英里", "US$/英里", lambda y, s_: f"{XR[(k('claims'),y,s_)]}*{P(op,'SEV',y,s_)}/1000000/{P(op,'LOSSR',y,s_)}", "＝理賠率×每件成本÷損失率", "0.000")
    xrow(k("floor"), "保費下限／車英里", "US$/英里", lambda y, s_: f"{P(op,'INS_FLOOR',y,s_)}/({P(op,'PMD',y,s_)}/{P(op,'PAID',y,s_)}*365)", "", "0.000")
    xrow(k("ins"), "保險與事故理賠／付費英里", "US$/英里", lambda y, s_: f"(MAX({XR[(k('eloss'),y,s_)]},{XR[(k('floor'),y,s_)]})+{P(op,'PHYS',y,s_)})/{P(op,'PAID',y,s_)}",
         "＝(MAX(預期理賠, 保費下限)＋自身車損)÷付費占比")
    xr[0] += 1
wsX.freeze_panes = "C3"

# ---------------------------------------------------------------- 每英里成本堆疊（v0.2，本線權威）
for j, w in enumerate([40, 11] + [10] * 12 + [60], 1):
    wsK2.column_dimensions[CL(j)].width = w
wsK2.cell(row=1, column=1, value="每付費英里成本堆疊（本線權威；Waymo、Tesla Robotaxi 個股模型讀此頁）。欄＝年份×成本情境（低成本／基準／高成本）。").font = BOLD
KS = {}
kr = 3
for op in OPS4:
    c = wsK2.cell(row=kr, column=1, value=OPNAME[op]); c.font = BOLD
    for j in range(1, 16):
        wsK2.cell(row=kr, column=j).fill = SEC_FILL
    kr += 1
    hdr(wsK2, kr, ["項目", "單位"] + [f"{y} {n}" for y in YRS4 for n in ("低成本", "基準", "高成本")] + ["等式"])
    kr += 1
    rows_ = []
    for key, lab in ITEMS + [("unit", "＝單位營運成本（1–8，不含研發、平台）")] + FULL_ITEMS + \
            [("full", "＝每付費英里全成本（1–10）"), ("gap1", "全成本 − 1 美元門檻"), ("gapaaa", "全成本 − AAA 自有車 0.77"),
             ("gap1u", "單位營運成本 − 1 美元門檻"), ("m_pmd", "備忘：每車每日付費英里（稼動率）"), ("m_dead", "備忘：空駛比（deadhead）"),
             ("m_vmy", "備忘：每車年行駛英里"), ("m_life", "備忘：車輛壽命（年）"), ("m_fleetmi", "備忘：車隊年付費英里（研發分母）")]:
        wsK2.cell(row=kr, column=1, value=lab)
        wsK2.cell(row=kr, column=2, value="US$/英里" if not key.startswith("m_") else "")
        if key in ("unit", "full"):
            wsK2.cell(row=kr, column=1).font = BOLD
        for y in YRS4:
            for s_ in SCS:
                col = dcol(y, s_, 3)
                p = params(op, y, s_)
                d = items(p)
                q = derived(p)
                if key in ("ins", "remote", "monitor"):
                    f = XR[(f"{op}_{key}", y, s_)]
                elif key in d:
                    f = d[key]
                elif key == "unit":
                    f = "SUM(" + ",".join(f"{col}{KS[(op,k_,'row')]}" for k_, _ in ITEMS) + ")"
                elif key == "full":
                    f = f"{col}{KS[(op,'unit','row')]}+{col}{KS[(op,'rnd','row')]}+{col}{KS[(op,'plat','row')]}"
                elif key == "gap1":
                    f = f"{col}{KS[(op,'full','row')]}-1"
                elif key == "gapaaa":
                    f = f"{col}{KS[(op,'full','row')]}-{I('P_OWN_AVG')}"
                elif key == "gap1u":
                    f = f"{col}{KS[(op,'unit','row')]}-1"
                elif key == "m_pmd":
                    f = p["PMD"]
                elif key == "m_dead":
                    f = f"1-{p['PAID']}"
                elif key == "m_vmy":
                    f = f"{p['PMD']}/{p['PAID']}*365"
                elif key == "m_life":
                    f = f"{p['LIFE']}/{col}{kr-1}"
                elif key == "m_fleetmi":
                    f = q["FLEETMI"]
                c = wsK2.cell(row=kr, column=3 + YRS4.index(y) * 3 + SCS.index(s_), value="=" + f)
                c.number_format = "0%" if key == "m_dead" else ("#,##0" if key in ("m_pmd", "m_vmy", "m_fleetmi") else "0.00")
                KS[(op, key, y, s_)] = f"'每英里成本堆疊'!${col}${kr}"
        KS[(op, key, "row")] = kr
        eq = {"veh": "車價÷(壽命里程×付費占比)", "fin": "車價×利率×0.5÷(每車日付費英里×365)", "energy": "每車英里能源÷付費占比",
              "maint": "每車英里維修清潔÷付費占比", "ins": "讀接管轉成本頁（理賠率×每件成本÷損失率，與保費下限取大）",
              "remote": "讀接管轉成本頁（時薪×營運小時÷(每人車數×每車日付費英里)）", "monitor": "讀接管轉成本頁（安全員占比×時薪×營運小時÷每車日付費英里）",
              "depot": "每場站年固定成本÷(每場站車數×每車日付費英里×365)＋充電變動÷付費占比",
              "rnd": "年研發÷車隊年付費英里（Tesla＝Model Y＋Cybercab 合計）", "plat": "基準價格×平台占比×抽成率",
              "m_fleetmi": "英里/年"}.get(key, "")
        wsK2.cell(row=kr, column=15, value=eq).alignment = WRAP
        kr += 1
    kr += 1

# 年度路徑（2026–2032）：錨點年取堆疊值，奇數年取相鄰錨點幾何平均
PYEARS = list(range(2026, 2033))
PC = {y: CL(3 + i) for i, y in enumerate(PYEARS)}
c = wsK2.cell(row=kr, column=1, value="年度路徑 2026–2032（奇數年＝相鄰錨點幾何平均）與跨門檻年份"); c.font = BOLD
for j in range(1, 16):
    wsK2.cell(row=kr, column=j).fill = SEC_FILL
kr += 1
hdr(wsK2, kr, ["營運商｜口徑｜情境", "單位"] + PYEARS + ["跨 1 美元年份", "跨 AAA 0.77 年份"])
kr += 1
PATH = {}


def path_row(label, anchor_fn, key):
    global kr
    wsK2.cell(row=kr, column=1, value=label)
    wsK2.cell(row=kr, column=2, value="US$/英里")
    for y in PYEARS:
        if y in YRS4:
            f = anchor_fn(y)
        else:
            f = f"SQRT({PC[y-1]}{kr}*{PC[y+1]}{kr})"
        c = wsK2.cell(row=kr, column=3 + PYEARS.index(y), value="=" + f); c.number_format = "0.00"
        PATH[key + (y,)] = f"'每英里成本堆疊'!${PC[y]}${kr}"
    for j, thr in ((10, "1"), (11, I('P_OWN_AVG'))):
        f = "\"2032 前未跨過\""
        for yy in reversed(PYEARS):
            f = f"IF({PC[yy]}{kr}<={thr},{yy},{f})"
        f = "=" + f
        wsK2.cell(row=kr, column=j, value=f)
        PATH[key + (("x1" if j == 10 else "xaaa"),)] = f"'每英里成本堆疊'!${CL(j)}${kr}"
    kr += 1


KR_HDR = kr - 1
SNM = {"L": "低成本", "B": "基準", "H": "高成本"}
for op in OPS4:
    for met, mname in (("full", "全成本"), ("unit", "單位營運成本")):
        for s_ in SCS:
            path_row(f"{OPNAME[op]}｜{mname}｜{SNM[s_]}", lambda y, op=op, met=met, s_=s_: KS[(op, met, y, s_)], (op, met, s_))
# Tesla 合計（依車隊付費英里加權 Model Y 與 Cybercab）
for met, mname in (("full", "全成本"), ("unit", "單位營運成本")):
    for s_ in SCS:
        def tfn(y, met=met, s_=s_):
            wy = f"{P('TY','FLEET',y,s_)}*{P('TY','PMD',y,s_)}"
            wc = f"{P('TC','FLEET',y,s_)}*{P('TC','PMD',y,s_)}"
            return f"({KS[('TY',met,y,s_)]}*{wy}+{KS[('TC',met,y,s_)]}*{wc})/({wy}+{wc})"
        path_row(f"Tesla 合計（車隊英里加權）｜{mname}｜{SNM[s_]}", tfn, ("T", met, s_))
wsK2.freeze_panes = "C3"

# ---------------------------------------------------------------- 每英里價格（v0.2）
for j, w in enumerate([46, 13, 13, 13, 13, 13, 13, 13, 18, 60], 1):
    wsP.column_dimensions[CL(j)].width = w
wsP.cell(row=1, column=1, value="每英里價格：觀察值（依城市、行程長度）、營運商價格路徑、價格是否覆蓋成本").font = BOLD
pr = 3
c = wsP.cell(row=pr, column=1, value="A. 觀察價格"); c.font = BOLD
for j in range(1, 11):
    wsP.cell(row=pr, column=j).fill = SEC_FILL
pr += 1
hdr(wsP, pr, ["項目", "值", "單位", "城市", "日期", "", "", "", "標記", "出處／備註"])
pr += 1
PX = {}
OBS = [
    ("w_sf_mi", "Waymo 每英里（Obi）", f"={I('X_WAYMO_SF_MI')}", "US$/英里", "舊金山灣區", "2025-11～2026-01", "[Verified]", "v0.1 誤植為每公里；v0.2 更正"),
    ("w_sf_trip", "Waymo 每趟（Obi）", f"={I('X_WAYMO_SF_TRIP')}", "US$/趟", "舊金山灣區", "2025-11～2026-01", "[Verified]", ""),
    ("w_sf_len", "隱含每趟英里＝每趟÷每英里", "=B6/B5", "英里", "舊金山灣區", "—", "[Derived]", "若單位為公里則每趟僅約 2.1 英里，與 CPUC 不符"),
    ("cpuc_len", "CPUC 每趟付費英里＝Q2 總英里÷趟數×付費占比", f"={CPUC_Q2MI}/{CPUC_Q2}*{P('W','PAID',2026,'B')}", "英里", "加州", "2026 Q2", "[Derived]", "交叉驗證：與 Obi 隱含 3.44 同量級 → 5.72 為每英里"),
    ("w_old_km", "v0.1 錯誤換算（5.72 美元/公里 × 1.609）", f"={I('X_WAYMO_SF_MI')}*1.609", "US$/英里", "—", "—", "[Derived]", "已作廢，僅留作更正紀錄（框架 v0.1 第 60 列「≈9.2」）"),
    ("t_sf_mi", "Tesla 每英里（Obi，有安全員）", f"={I('X_TSLA_SF_MI')}", "US$/英里", "舊金山灣區", "2025-11～2026-01", "[Verified]", ""),
    ("t_sf_len", "Tesla 隱含每趟英里", f"={I('X_TSLA_SF_TRIP')}/{I('X_TSLA_SF_MI')}", "英里", "舊金山灣區", "—", "[Derived]", ""),
    ("prem", "Waymo 對 Uber 溢價（2025-11～2026-01）", f"={I('P_PREM_2026')}", "%", "舊金山灣區", "2026-01", "[Verified]", "2025-06 Obi 研究為 30–40%：溢價半年內收斂"),
    ("human", "人類叫車每付費英里（全美推估）", f"={C('p_human')}", "US$/英里", "全美", "2026 Q3", "[Derived]", "計算頁價格錨 1"),
    ("cn_wuhan", "Apollo Go 武漢約 5 公里 7.5 元（換算每英里，匯率 7.2）", "=7.5/(5/1.609)/7.2", "US$/英里", "武漢", "2024-07", "[Derived]", S["CNCRT6"] + "；有補貼"),
    ("cc_vs_my", "Cybercab 相對 Model Y 同路線價格", "=7.77/12.1-1", "%", "奧斯汀", "2026-09", "[Interested-party]", S["TCC"] + "；單一樣本"),
]
for k, nm, f, u, city, dt, tag, note in OBS:
    for j, v in ((1, nm), (2, f), (3, u), (4, city), (5, dt), (9, tag), (10, note)):
        c = wsP.cell(row=pr, column=j, value=v); c.alignment = WRAP
    wsP.cell(row=pr, column=2).number_format = "0.0%" if u == "%" else "0.00"
    PX[k] = f"'每英里價格'!$B${pr}"
    pr += 1
assert PX["w_sf_trip"].endswith("$6") and PX["w_sf_mi"].endswith("$5"), (PX["w_sf_mi"], PX["w_sf_trip"])
pr += 1
c = wsP.cell(row=pr, column=1, value="B. Tesla 奧斯汀票價：每英里實收價格隨行程長度變化"); c.font = BOLD
for j in range(1, 11):
    wsP.cell(row=pr, column=j).fill = SEC_FILL
pr += 1
TRIPS = [1, 2, 3, 5, 8, 10, 15]
hdr(wsP, pr, ["票價表＼行程英里"] + TRIPS + ["", "說明"])
pr += 1
for nm, base, rate, note in (("現行（2026-03-12 起）起跳 3.00＋1.40/英里", I('X_TSLA_BASE'), I('X_TSLA_PRICE'), S["TFARE"]),
                             ("前一版 起跳 3.25＋1.00/英里", "3.25", "1.00", S["TFARE"])):
    wsP.cell(row=pr, column=1, value=nm)
    for j, d in enumerate(TRIPS, 2):
        c = wsP.cell(row=pr, column=j, value=f"=({base}+{rate}*{d})/{d}"); c.number_format = "0.00"
    wsP.cell(row=pr, column=10, value=note)
    pr += 1
wsP.cell(row=pr, column=1, value="人類叫車每英里（全美，對照）")
for j in range(2, 2 + len(TRIPS)):
    c = wsP.cell(row=pr, column=j, value=f"={C('p_human')}"); c.number_format = "0.00"
pr += 2
c = wsP.cell(row=pr, column=1, value="C. 營運商價格路徑 vs 成本（價格低／高＝低價／高價；成本取基準）"); c.font = BOLD
for j in range(1, 11):
    wsP.cell(row=pr, column=j).fill = SEC_FILL
pr += 1
hdr(wsP, pr, ["營運商｜年", "價格 低", "價格 基準", "價格 高", "單位營運成本 基準", "全成本 基準", "基準價−單位成本", "基準價−全成本", "價格覆蓋單位成本？", "說明"])
pr += 1
PPATH = {}
for op in OPS4:
    for y in YRS4:
        wsP.cell(row=pr, column=1, value=f"{OPNAME[op]}｜{y}")
        for j, s_ in zip((2, 3, 4), SCS):
            c = wsP.cell(row=pr, column=j, value=f"={P(op,'PRICE',y,s_)}"); c.number_format = "0.00"
            PPATH[(op, y, s_)] = f"'每英里價格'!${CL(j)}${pr}"
        for j, f in ((5, KS[(op, 'unit', y, 'B')]), (6, KS[(op, 'full', y, 'B')]), (7, f"C{pr}-E{pr}"), (8, f"C{pr}-F{pr}")):
            c = wsP.cell(row=pr, column=j, value=f"={f}"); c.number_format = "0.00"
        wsP.cell(row=pr, column=9, value=f"=IF(G{pr}>=0,\"是\",\"否（補貼）\")")
        pr += 1

# ---------------------------------------------------------------- Tesla 成本落差診斷（v0.2）
for j, w in enumerate([8, 34, 30, 12, 12, 12, 8, 12, 50, 36, 40], 1):
    wsG.column_dimensions[CL(j)].width = w
wsG.cell(row=1, column=1, value="Tesla 成本落差診斷：Tesla 稱營運成本可到每英里 0.30 美元（Q3 2026 法說），但 2026 Model Y 車隊看起來遠高於此。以下把 2026 Model Y 基準逐項換成 2030 Cybercab 基準，量化每個假說。").font = BOLD
gr = 3
BASEP = derived(params("TY", 2026, "B"))
MATP = derived(params("TC", 2030, "B"))
GROUPS = [
    ("H1", "安全員或追車尚未移除", ["MON", "CHASE", "MON_WAGE"],
     "2026-03 奧斯汀約 40 台中無監督僅 4–8 台；8 月追蹤 100% 無安全員；灣區約 130 台全有安全員（" + S["ELECMAR"] + "；" + S["TNW8"] + "）",
     "德州、佛州已大致移除；灣區需 CPUC 無人許可。Cybercab 無方向盤，結構上不可能有安全員",
     "Tesla 揭露的無監督里程占總 Robotaxi 里程比例；灣區取得 CPUC 無人載客許可"),
    ("H2", "車隊小、服務區小 → 稼動率低（每車日付費英里、空駛）", ["PMD", "PAID", "HRS"],
     "Tesla 無監督 310 萬英里÷270 日，平均每日約 1.1 萬車英里，分到 100–300 台約每車日 38–115 車英里；Waymo 約 170 車英里（" + S["TSAFE"] + "）",
     "車隊與服務區擴大、24 小時營運；FSD v15（Musk 稱 2026 底至 2027 初）後才大規模無監督",
     "每車每日付費趟數／英里（Robotaxi Tracker、TxDMV 登記÷營運）；等候時間"),
    ("H3", "仍用 Model Y 而非 Cybercab（車價、壽命、能耗、維修）", ["VEH", "LIFE", "ENERGY", "MAINT"],
     "Cybercab 2026-09 才開放；同路線票價低約 35%（" + S["TCC"] + "）",
     "Cybercab 量產爬坡（Giga Texas）；車價需降到 3 萬以下才兌現",
     "Cybercab 占 Tesla Robotaxi 車隊比例；Cybercab 單位成本揭露"),
    ("H4", "遠端營運比例（每位遠端人員監看車數低）", ["RA_CONC", "R", "RA_MIN", "RA_WAGE", "OCC"],
     "Tesla 未揭露遠端比例與頻率，遠端可直接接手低速駕駛（" + S["TCMARKEY"] + "）；FSD 社群關鍵接管 v14.3 城市約 853 英里/次（" + S["FSDT"] + "）",
     "接管里程每年改善（基準 1.8 倍/年）且監管接受較高監看比之後",
     "Tesla 或監管揭露遠端人員數÷車隊；社群 tracker 關鍵接管里程；NHTSA 遠端操作相關規範"),
    ("H5", "保險與法規（事故率高、保費下限）", ["CLAIMS", "SEV", "PHYS", "LOSSR", "INS_FLOOR"],
     "NHTSA SGO：14 起／約 80 萬英里，約為 Tesla 自報人類基準 4 倍（" + S["ELECFEB"] + "）；2026-10 安全報告 310 萬無監督英里 0 受傷（" + S["TSAFE"] + "）",
     "理賠率隨接管里程改善下降",
     "NHTSA SGO 每百萬英里事故件數；Tesla 安全報告；保險揭露"),
    ("H6", "場站規模不經濟（每場站車數少）", ["DEPOT_FIX", "DEPOT_VEH", "CHG"],
     "每城數十至數百台；奧斯汀 300+ Cybercab，各城另有數百台在場站待命（" + S["TSLAQ3"] + "）",
     "單城車隊達數百台以上",
     "各城車隊規模；場站數"),
    ("H8", "其他（融資、平台、追車車輛成本）", ["RATE", "PLAT_SH", "TAKE", "PRICE", "CHASE_HR"], "—", "—", "—"),
    ("H7", "研發攤提（車隊規模，僅影響全成本）", ["RND", "FLEETMI"],
     "研發分攤假設 6 億美元/年 ÷ 2026 車隊年付費英里約 770 萬",
     "車隊年付費英里達數十億英里（2030 基準約 3 萬台 Cybercab）",
     "Tesla Robotaxi 車隊規模、付費里程"),
]
# 完整性：所有 items 用到的鍵都被某群組覆蓋
_used = set(BASEP.keys()) - {"Y", "MPI0", "T0", "G", "K_RA", "P_CLAIM", "CONC0", "BETA", "CONCMAX", "FLEET", "FLEET_TY", "FLEET_TC", "PMD_TY", "PMD_TC"}
_cov = set(k for g in GROUPS for k in g[2])
assert _used <= _cov, _used - _cov


def swap(src, dst, keys):
    q = dict(src)
    for k in keys:
        q[k] = dst[k]
    return q


c = wsG.cell(row=gr, column=1, value="A. 起點與終點"); c.font = BOLD
for j in range(1, 12):
    wsG.cell(row=gr, column=j).fill = SEC_FILL
gr += 1
GD = {}
for k, nm, f in (("b_unit", "起點：Model Y 2026 基準 單位營運成本", unit_expr(BASEP)),
                 ("b_full", "起點：Model Y 2026 基準 全成本", full_expr(BASEP)),
                 ("m_unit", "終點：Cybercab 2030 基準 單位營運成本", unit_expr(MATP)),
                 ("m_full", "終點：Cybercab 2030 基準 全成本", full_expr(MATP)),
                 ("stack_chk", "核對：起點單位成本 − 堆疊頁 Model Y 2026 基準", f"{unit_expr(BASEP)}-{KS[('TY','unit',2026,'B')]}"),
                 ("tsla_claim", "Tesla 目標：營運成本每英里", f"{I('X_TSLA_COST')}"),
                 ("cc32L", "Cybercab 2032 低成本情境 單位營運成本", f"{KS[('TC','unit',2032,'L')]}"),
                 ("cc32L_x", "同上，扣除車輛攤提與融資（推測 Tesla 口徑）", f"{KS[('TC','unit',2032,'L')]}-{KS[('TC','veh',2032,'L')]}-{KS[('TC','fin',2032,'L')]}")):
    assert len(f) < 7900, (k, len(f))
    wsG.cell(row=gr, column=2, value=nm)
    c = wsG.cell(row=gr, column=4, value="=" + f); c.number_format = "0.000"
    wsG.cell(row=gr, column=5, value="US$/英里")
    GD[k] = f"'Tesla成本落差診斷'!$D${gr}"
    gr += 1
gr += 1
c = wsG.cell(row=gr, column=1, value="B. 假說逐項量化（改善＝成本下降，US$/付費英里）"); c.font = BOLD
for j in range(1, 12):
    wsG.cell(row=gr, column=j).fill = SEC_FILL
gr += 1
hdr(wsG, gr, ["ID", "假說", "參數群", "單獨改善（單位成本）", "最後改善（單位成本）", "平均（排序依據）", "排名", "平均改善（全成本）", "證據", "何時可能消失", "證偽指標"])
gr += 1
h_first = gr
for hid, nm, keys, ev, when, fals in GROUPS:
    alone_u = f"{GD['b_unit']}-{unit_expr(swap(BASEP, MATP, keys))}"
    last_u = f"{unit_expr(swap(MATP, BASEP, keys))}-{GD['m_unit']}"
    alone_f = f"{GD['b_full']}-{full_expr(swap(BASEP, MATP, keys))}"
    last_f = f"{full_expr(swap(MATP, BASEP, keys))}-{GD['m_full']}"
    for f in (alone_u, last_u, alone_f, last_f):
        assert len(f) < 7900, (hid, len(f))
    wsG.cell(row=gr, column=1, value=hid)
    wsG.cell(row=gr, column=2, value=nm).alignment = WRAP
    wsG.cell(row=gr, column=3, value="、".join(keys)).alignment = WRAP
    c = wsG.cell(row=gr, column=4, value="=" + alone_u); c.number_format = "0.00"
    c = wsG.cell(row=gr, column=5, value="=" + last_u); c.number_format = "0.00"
    c = wsG.cell(row=gr, column=6, value=f"=(D{gr}+E{gr})/2+ROW()*1E-12"); c.number_format = "0.00"
    c = wsG.cell(row=gr, column=8, value=f"=(({alone_f})+({last_f}))/2"); c.number_format = "0.00"
    for j, v in ((9, ev), (10, when), (11, fals)):
        wsG.cell(row=gr, column=j, value=v).alignment = WRAP
    gr += 1
h_last = gr - 1
for rr_ in range(h_first, h_last + 1):
    wsG.cell(row=rr_, column=7, value=f"=RANK(F{rr_},$F${h_first}:$F${h_last})")
GD["h_first"], GD["h_last"] = h_first, h_last
wsG.cell(row=gr, column=2, value="讀法：「單獨改善」＝只把這一群參數從 2026 Model Y 換成 2030 Cybercab；「最後改善」＝其他都已成熟、最後才換這一群。兩者平均為交互作用下的近似 Shapley 分攤。稼動率與人工類（安全員、遠端）相乘，所以單獨改善加總大於總落差。").alignment = WRAP
gr += 2

c = wsG.cell(row=gr, column=1, value="C. 依序拆解（瀑布；順序：安全員→遠端→稼動→車型→保險→場站→其他→研發）"); c.font = BOLD
for j in range(1, 12):
    wsG.cell(row=gr, column=j).fill = SEC_FILL
gr += 1
hdr(wsG, gr, ["步驟", "換入", "", "單位成本", "全成本", "本步改善（單位）", "", "本步改善（全）"])
gr += 1
ORDER = ["H1", "H4", "H2", "H3", "H5", "H6", "H8", "H7"]
cur = dict(BASEP)
wsG.cell(row=gr, column=1, value=0); wsG.cell(row=gr, column=2, value="起點：Model Y 2026 基準")
c = wsG.cell(row=gr, column=4, value=f"={GD['b_unit']}"); c.number_format = "0.00"
c = wsG.cell(row=gr, column=5, value=f"={GD['b_full']}"); c.number_format = "0.00"
w_first = gr
gr += 1
gmap = {g[0]: g for g in GROUPS}
for n, hid in enumerate(ORDER, 1):
    cur = swap(cur, MATP, gmap[hid][2])
    wsG.cell(row=gr, column=1, value=n)
    wsG.cell(row=gr, column=2, value=f"{hid} {gmap[hid][1]}")
    c = wsG.cell(row=gr, column=4, value="=" + unit_expr(cur)); c.number_format = "0.00"
    c = wsG.cell(row=gr, column=5, value="=" + full_expr(cur)); c.number_format = "0.00"
    c = wsG.cell(row=gr, column=6, value=f"=D{gr-1}-D{gr}"); c.number_format = "0.00"
    c = wsG.cell(row=gr, column=8, value=f"=E{gr-1}-E{gr}"); c.number_format = "0.00"
    gr += 1
GD["w_end_unit"] = f"'Tesla成本落差診斷'!$D${gr-1}"
GD["w_end_full"] = f"'Tesla成本落差診斷'!$E${gr-1}"
wsG.cell(row=gr, column=2, value="核對：瀑布終點 − Cybercab 2030 基準（應為 0）")
c = wsG.cell(row=gr, column=4, value=f"=D{gr-1}-{GD['m_unit']}"); c.number_format = "0.000000"
c = wsG.cell(row=gr, column=5, value=f"=E{gr-1}-{GD['m_full']}"); c.number_format = "0.000000"
GD["w_chk"] = gr
gr += 2

c = wsG.cell(row=gr, column=1, value="D. 價格低是反映成本還是補貼？"); c.font = BOLD
for j in range(1, 12):
    wsG.cell(row=gr, column=j).fill = SEC_FILL
gr += 1
for k, nm, f, fmt in (
        ("sub_price", "Model Y 2026 基準每付費英里價格（3.00＋1.40/英里，5 英里行程）", P('TY', 'PRICE', 2026, 'B'), "0.00"),
        ("sub_unit", "每付費英里補貼（單位營運成本 − 價格）", f"{KS[('TY','unit',2026,'B')]}-{P('TY','PRICE',2026,'B')}", "0.00"),
        ("sub_full", "每付費英里補貼（全成本 − 價格）", f"{KS[('TY','full',2026,'B')]}-{P('TY','PRICE',2026,'B')}", "0.00"),
        ("sub_day", "每車每日營運虧損（單位口徑）", f"({KS[('TY','unit',2026,'B')]}-{P('TY','PRICE',2026,'B')})*{P('TY','PMD',2026,'B')}", "#,##0"),
        ("sub_cov", "價格÷單位營運成本（覆蓋率）", f"{P('TY','PRICE',2026,'B')}/{KS[('TY','unit',2026,'B')]}", "0%"),
        ("sub_mature", "Cybercab 2030 基準：價格 − 單位營運成本", f"{P('TC','PRICE',2030,'B')}-{KS[('TC','unit',2030,'B')]}", "0.00"),
        ("sub_mature_f", "Cybercab 2030 基準：價格 − 全成本", f"{P('TC','PRICE',2030,'B')}-{KS[('TC','full',2030,'B')]}", "0.00"),
        ("sub_human", "Model Y 2026 價格÷人類叫車每英里", f"{P('TY','PRICE',2026,'B')}/{C('p_human')}", "0%")):
    wsG.cell(row=gr, column=2, value=nm)
    c = wsG.cell(row=gr, column=4, value="=" + f); c.number_format = fmt
    GD[k] = f"'Tesla成本落差診斷'!$D${gr}"
    gr += 1
gr += 1
c = wsG.cell(row=gr, column=1, value="E. 結論"); c.font = BOLD
for j in range(1, 12):
    wsG.cell(row=gr, column=j).fill = SEC_FILL
gr += 1
rng_f = f"$F${h_first}:$F${h_last}"
wsG.cell(row=gr, column=2, value="主因（單位營運成本，平均改善最大）")
wsG.cell(row=gr, column=4, value=f"=INDEX($B${h_first}:$B${h_last},MATCH(LARGE({rng_f},1),{rng_f},0))")
GD["main1"] = f"'Tesla成本落差診斷'!$D${gr}"
gr += 1
wsG.cell(row=gr, column=2, value="次因")
wsG.cell(row=gr, column=4, value=f"=INDEX($B${h_first}:$B${h_last},MATCH(LARGE({rng_f},2),{rng_f},0))")
GD["main2"] = f"'Tesla成本落差診斷'!$D${gr}"
gr += 1
wsG.cell(row=gr, column=2, value="第三")
wsG.cell(row=gr, column=4, value=f"=INDEX($B${h_first}:$B${h_last},MATCH(LARGE({rng_f},3),{rng_f},0))")
gr += 1
for t in CONCL:
    c = wsG.cell(row=gr, column=2, value=t); c.alignment = WRAP
    wsG.merge_cells(start_row=gr, start_column=2, end_row=gr, end_column=11)
    wsG.row_dimensions[gr].height = 48
    gr += 1

# ---------------------------------------------------------------- 情境
YC = {y: CL(3 + i) for i, y in enumerate(YEARS)}  # C..L
CMAP = {"低": "H", "基準": "B", "高": "L"}  # 低滲透情境配高成本、高滲透配低成本
NR = {}  # (scn, key) -> row


def interp(y, anchors):
    """anchors: list of (year, ref). 回傳線性內插公式（不含 =）。"""
    for (y1, a), (y2, b) in zip(anchors, anchors[1:]):
        if y1 <= y <= y2:
            if y == y1:
                return f"{a}"
            if y == y2:
                return f"{b}"
            return f"{a}+({b}-{a})*({y}-{y1})/({y2}-{y1})"
    return f"{anchors[-1][1]}"


ROWS = [
    ("year", "年", "", "0"),
    ("us_trips", "美國叫車趟數（不含誘發）", "百萬趟", "#,##0"),
    ("us_miles", "美國叫車付費英里", "百萬英里", "#,##0"),
    ("cov", "商轉都會覆蓋率", "%", "0.0%"),
    ("s", "都會區內 Robotaxi 份額", "%", "0.0%"),
    ("pen", "Robotaxi 滲透率（替代趟數÷叫車趟數）", "%", "0.00%"),
    ("rt_trips", "Robotaxi 替代趟數", "百萬趟", "#,##0.0"),
    ("rt_weekly", "Robotaxi 每週付費趟數（替代）", "千趟/週", "#,##0"),
    ("price", "Robotaxi 每付費英里實收價格", "US$/英里", "0.00"),
    ("price_rel", "價格÷人類叫車每英里", "倍", "0.00"),
    ("cost", "Robotaxi 每付費英里全成本", "US$/英里", "0.00"),
    ("margin", "每付費英里毛利", "US$/英里", "0.00"),
    ("rt_miles", "Robotaxi 付費英里（替代）", "百萬英里", "#,##0"),
    ("rt_gb", "Robotaxi 訂單金額（替代）", "百萬美元", "#,##0"),
    ("induce", "誘發需求倍數 (P/P2026)^ε", "倍", "0.00"),
    ("rt_miles_j", "Robotaxi 付費英里（含誘發）", "百萬英里", "#,##0"),
    ("rt_gb_j", "Robotaxi 訂單金額（含誘發）", "百萬美元", "#,##0"),
    ("pool_rh_share", "含誘發英里÷原叫車池英里", "%", "0.0%"),
    ("flag_double", "價格≤傑文斯門檻（叫車英里翻倍）", "年或 9999", "0"),
    ("flag_own", "價格≤自有車每英里成本（打開自有車池）", "年或 9999", "0"),
    ("tpvd", "每車每日付費趟數", "趟/日", "0.0"),
    ("fleet", "所需車隊（替代趟數口徑）", "台", "#,##0"),
    ("paid", "付費英里占比", "%", "0.0%"),
    ("vmi_yr", "每車年行駛英里", "英里", "#,##0"),
    ("life_yr", "車輛壽命", "年", "0.0"),
    ("adds", "年新增車輛需求（擴車＋汰換）", "台", "#,##0"),
    ("w_sh", "Waymo 份額", "%", "0.0%"),
    ("t_sh", "Tesla 份額", "%", "0.0%"),
    ("o_sh", "其他份額", "%", "0.0%"),
    ("w_trips", "Waymo 趟數", "百萬趟", "#,##0.0"),
    ("t_trips", "Tesla 趟數", "百萬趟", "#,##0.0"),
    ("o_trips", "其他營運商趟數", "百萬趟", "#,##0.0"),
    ("w_weekly", "Waymo 每週趟數", "千趟/週", "#,##0"),
    ("plat_sh", "經 Uber／Lyft 平台占比", "%", "0.0%"),
    ("plat_trips", "經平台 Robotaxi 趟數", "百萬趟", "#,##0.0"),
    ("plat_rev", "平台抽成收入（替代口徑）", "百萬美元", "#,##0"),
]

hdr(wsN, 1, ["項目", "單位"] + YEARS + ["說明"], [40, 11] + [11] * len(YEARS) + [50])
r0 = 3
for scn in ("低", "基準", "高"):
    x = SC[scn]
    t = wsN.cell(row=r0, column=1, value=f"情境：{scn}（{'低滲透' if scn=='低' else '高滲透' if scn=='高' else '基準'}）")
    t.font = BOLD
    for j in range(1, 3 + len(YEARS) + 1):
        wsN.cell(row=r0, column=j).fill = SEC_FILL
    base_r = r0 + 1
    for i, (k, lab, unit, fmt) in enumerate(ROWS):
        NR[(scn, k)] = base_r + i
        wsN.cell(row=base_r + i, column=1, value=lab)
        wsN.cell(row=base_r + i, column=2, value=unit)

    def A(k, y):
        return f"{YC[y]}{NR[(scn, k)]}"

    for y in YEARS:
        col = YC[y]
        F = {}
        F["year"] = y
        F["us_trips"] = f"={C('us_trips')}*(1+{I('M_GROWTH',x)})^({A('year',y)}-2026)"
        F["us_miles"] = f"={A('us_trips',y)}*{I('M_TRIP_MI')}"
        F["cov"] = "=" + interp(y, [(2026, I('R_COV_2026')), (2028, I('S_COV_2028', x)), (2030, I('S_COV_2030', x)),
                                    (2032, I('S_COV_2032', x)), (2035, I('S_COV_2035', x))])
        F["s"] = f"={I('S_SMAX',x)}/(1+EXP(-{C('k',x)}*({A('year',y)}-{I('S_T0',x)})))"
        F["pen"] = f"={A('cov',y)}*{A('s',y)}"
        F["rt_trips"] = f"={A('us_trips',y)}*{A('pen',y)}"
        F["rt_weekly"] = f"={A('rt_trips',y)}*1000/52"
        F["price"] = "=" + interp(y, [(2026, C('p26')), (2030, C('p30', x)), (2035, C('p35', x))])
        F["price_rel"] = f"={A('price',y)}/{C('p_human')}"
        cm_ = CMAP[scn]
        if y == 2026:
            F["cost"] = f"={A('w_sh',y)}*{PATH[('W','unit','B',y)]}+{A('t_sh',y)}*{PATH[('T','unit','B',y)]}+{A('o_sh',y)}*{PATH[('W','unit','B',y)]}"
        elif y < 2030:
            F["cost"] = f"={YC[2026]}{NR[(scn,'cost')]}+({YC[2030]}{NR[(scn,'cost')]}-{YC[2026]}{NR[(scn,'cost')]})*({y}-2026)/4"
        elif y <= 2032:
            F["cost"] = f"={A('w_sh',y)}*{PATH[('W','unit',cm_,y)]}+{A('t_sh',y)}*{PATH[('T','unit',cm_,y)]}+{A('o_sh',y)}*{PATH[('W','unit',cm_,y)]}"
        else:
            F["cost"] = f"={YC[2032]}{NR[(scn,'cost')]}*(1-{I('S_DECL',x)})^({A('year',y)}-2032)"
        F["margin"] = f"={A('price',y)}-{A('cost',y)}"
        F["rt_miles"] = f"={A('rt_trips',y)}*{I('M_TRIP_MI')}"
        F["rt_gb"] = f"={A('rt_miles',y)}*{A('price',y)}"
        F["induce"] = f"=({A('price',y)}/{C('p26')})^{I('P_ELAST',x)}"
        F["rt_miles_j"] = f"={A('rt_miles',y)}*{A('induce',y)}"
        F["rt_gb_j"] = f"={A('rt_miles_j',y)}*{A('price',y)}"
        F["pool_rh_share"] = f"={A('rt_miles_j',y)}/{A('us_miles',y)}"
        F["flag_double"] = f"=IF({A('price',y)}<={C('j_double',x)},{A('year',y)},9999)"
        F["flag_own"] = f"=IF({A('price',y)}<={I('P_OWN_AVG')},{A('year',y)},9999)"
        F["tpvd"] = "=" + interp(y, [(2026, C('tpvd_2026')), (2030, I('S_TPVD_2030', x))])
        F["fleet"] = f"={A('rt_trips',y)}*1000000/({A('tpvd',y)}*365)"
        F["paid"] = "=" + interp(y, [(2026, P('W', 'PAID', 2026, cm_)), (2030, P('W', 'PAID', 2030, cm_)), (2032, P('W', 'PAID', 2032, cm_))])
        F["vmi_yr"] = f"={A('tpvd',y)}*{I('M_TRIP_MI')}/{A('paid',y)}*365"
        F["life_yr"] = f"={P('W', 'LIFE', 2026, 'B')}/{A('vmi_yr',y)}"
        if y == 2026:
            F["adds"] = None
        else:
            F["adds"] = f"=MAX(0,{A('fleet',y)}-{A('fleet',y-1)})+{A('fleet',y-1)}/{A('life_yr',y-1)}"
        F["w_sh"] = "=" + interp(y, [(2026, f"(1-{I('R_OTHERS')})"), (2030, I('S_WAYMO_2030', x)), (2035, I('S_WAYMO_2035', x))])
        F["t_sh"] = "=" + interp(y, [(2026, I('R_TESLA_2026')), (2030, I('S_TESLA_2030', x)), (2035, I('S_TESLA_2035', x))])
        F["o_sh"] = f"=1-{A('w_sh',y)}-{A('t_sh',y)}"
        F["w_trips"] = f"={A('rt_trips',y)}*{A('w_sh',y)}"
        F["t_trips"] = f"={A('rt_trips',y)}*{A('t_sh',y)}"
        F["o_trips"] = f"={A('rt_trips',y)}*{A('o_sh',y)}"
        F["w_weekly"] = f"={A('w_trips',y)}*1000/52"
        F["plat_sh"] = "=" + interp(y, [(2026, I('R_PLAT_2026')), (2030, I('S_PLAT_2030', x))])
        F["plat_trips"] = f"={A('rt_trips',y)}*{A('plat_sh',y)}"
        F["plat_rev"] = f"={A('rt_gb',y)}*{A('plat_sh',y)}*{I('M_TAKE')}"
        for k, lab, unit, fmt in ROWS:
            v = F[k]
            if v is None:
                continue
            c = wsN.cell(row=NR[(scn, k)], column=3 + YEARS.index(y), value=v)
            c.number_format = fmt
    # 跨錨年份
    rr = base_r + len(ROWS)
    lab = [("x_double", "傑文斯門檻跨過年份（叫車英里翻倍價格）", "flag_double"),
           ("x_own", "自有車池打開年份（價格≤AAA 平均 0.77）", "flag_own")]
    for j, (k, l, fk) in enumerate(lab):
        NR[(scn, k)] = rr + j
        wsN.cell(row=rr + j, column=1, value=l)
        f = f"MIN({YC[2026]}{NR[(scn, fk)]}:{YC[2035]}{NR[(scn, fk)]})"
        wsN.cell(row=rr + j, column=3, value=f"=IF({f}=9999,\"2035 前未跨過\",{f})")
    r0 = rr + len(lab) + 2

# 說明欄
notes = {
    "s": "S(t)＝Smax/(1+exp(−k(t−t0)))",
    "pen": "＝覆蓋率×都會區內份額；2026 等於校準值",
    "price": "2026→2030→2035 線性內插（人類叫車每英里×倍數）",
    "cost": "讀每英里成本堆疊年度路徑：營運商份額加權的單位營運成本（不含研發、平台；其他營運商以 Waymo 代）。2026 三情境皆用基準（校準年），2027–2029 線性內插至 2030；低滲透情境配高成本、高滲透配低成本；2032 後每年降 S_DECL",
    "induce": "相對 2026 實收價格的常數彈性；2026＝1（已觀察趟數含當時價格）",
    "rt_miles_j": "傑文斯口徑：替代英里×誘發倍數；與滲透模型可能部分重複，下游擇一使用",
    "adds": "＝車隊淨增＋前一年車隊÷壽命",
    "flag_double": "符合者填年份，否則 9999（取最小值即跨過年份）",
}
for scn in ("低", "基準", "高"):
    for k, v in notes.items():
        wsN.cell(row=NR[(scn, k)], column=3 + len(YEARS), value=v)

# 外部預測對照
rr = r0
wsN.cell(row=rr, column=1, value="外部預測對照（2030 欄＝模型值）").font = BOLD
for j in range(1, 3 + len(YEARS) + 1):
    wsN.cell(row=rr, column=j).fill = SEC_FILL
hdr(wsN, rr + 1, ["項目", "外部值", "低", "基準", "高", "出處"])
cmp_rows = [
    ("GS 2030 Robotaxi 份額 8%", I('X_GS_2030'), "pen", 2030, "0.0%", "Goldman Sachs 2025-07"),
    ("GS 2030 車隊 3.5 萬台", I('X_GS_FLEET'), "fleet", 2030, "#,##0", "Goldman Sachs 2025-07"),
    ("MS 2030 年趟數 7.5 億", I('X_MS_RIDES'), "rt_trips", 2030, "#,##0", "Morgan Stanley 2026-05"),
    ("MS 約 2032 自駕占叫車里程 30%（以趟數滲透率對照）", I('X_MS_2032'), "pen", 2032, "0.0%", "Morgan Stanley 2025-12"),
]
for j, (lab, ext, k, y, fmt, src) in enumerate(cmp_rows):
    R_ = rr + 2 + j
    wsN.cell(row=R_, column=1, value=lab)
    c = wsN.cell(row=R_, column=2, value=f"={ext}"); c.number_format = fmt
    for m, scn in enumerate(("低", "基準", "高")):
        c = wsN.cell(row=R_, column=3 + m, value=f"={YC[y]}{NR[(scn, k)]}")
        c.number_format = fmt
    wsN.cell(row=R_, column=6, value=src)
wsN.freeze_panes = "C2"

# ---------------------------------------------------------------- 計算 F 節（v0.2：成本連結，寫在情境之後以取得列號）
sec("F. 每英里成本連結（讀每英里成本堆疊與情境頁）")
same("c26_pmi", "2026 Waymo 單位營運成本／付費英里（基準）", f"={KS[('W','unit',2026,'B')]}", "US$/英里", "讀堆疊頁；不含研發、平台")
same("c26_full", "2026 Waymo 全成本／付費英里（基準）", f"={KS[('W','full',2026,'B')]}", "US$/英里", "含研發攤提與平台抽成")
same("m26", "2026 每付費英里毛利（價格−Waymo 單位營運成本）", f"={C('p26')}-{C('c26_pmi')}", "US$/英里", "")
same("gap_driver", "2026 成本 vs 人類司機每英里收入", f"={C('c26_pmi')}-{I('M_DRIVER_MI')}", "US$/英里", "負值＝Waymo 單位營運成本已低於只付司機的人工成本")
line("c30_pmi", "2030 全美加權單位營運成本／付費英里", *[f"='情境'!${YC[2030]}${NR[(scn,'cost')]}" for scn in ("低", "基準", "高")], "US$/英里", "讀情境頁（低滲透配高成本）")
line("m30", "2030 每付費英里毛利", *[f"={C('p30',x)}-{C('c30_pmi',x)}" for x in "LBH"], "US$/英里", "")
line("rev_vday30", "2030 單車每日營收", *[f"={I('S_TPVD_2030',x)}*{I('M_TRIP_MI')}*{C('p30',x)}" for x in "LBH"], "US$/日", "＝每日趟數×每趟英里×每英里價格", "#,##0")
line("mar_vday30", "2030 單車每日毛利", *[f"={I('S_TPVD_2030',x)}*{I('M_TRIP_MI')}*{C('m30',x)}" for x in "LBH"], "US$/日", "", "#,##0")
line("j_floor_margin", "傑文斯門檻價格下每付費英里毛利（以 2030 成本）", *[f"={C('j_double',x)}-{C('c30_pmi',x)}" for x in "LBH"], "US$/英里", "負值＝在該情境成本下無法以門檻價格獲利")


def N(scn, k, y):
    return f"'情境'!${YC[y]}${NR[(scn, k)]}"


# ---------------------------------------------------------------- 敏感度
# 封閉式公式模板：{名} 代入輸入儲存格
def pen30_tpl(v):
    lyft = f"(((1-{v['M_UBER_SPEND']})/{I('M_FARE_LYFT')})/((1-{v['M_UBER_SPEND']})/{I('M_FARE_LYFT')}+{v['M_UBER_SPEND']}/{I('M_FARE_UBER')}))"
    us = f"({I('M_LYFT_Q')}*4*{v['M_LYFT_NA']}/{lyft})"
    rt = f"({CPUC_Q2}/13/{v['R_CA_SHARE']}*52/(1-{v['R_OTHERS']})/1000000)"
    s26 = f"({rt}/{us}/{v['R_COV_2026']})"
    k = f"(LN({v['S_SMAX']}/{s26}-1)/({v['S_T0']}-2026))"
    pen = f"{v['S_COV_2030']}*{v['S_SMAX']}/(1+EXP(-{k}*(2030-{v['S_T0']})))"
    return pen, us


def gb30_tpl(v):
    pen, us = pen30_tpl(v)
    lyft = f"(((1-{v['M_UBER_SPEND']})/{I('M_FARE_LYFT')})/((1-{v['M_UBER_SPEND']})/{I('M_FARE_LYFT')}+{v['M_UBER_SPEND']}/{I('M_FARE_UBER')}))"
    fare = f"({lyft}*{I('M_FARE_LYFT')}+(1-{lyft})*{I('M_FARE_UBER')})"
    return f"{us}*(1+{v['M_GROWTH']})^4*{pen}*{fare}*{v['P_REL_2030']}"


def gbj30_tpl(v):
    # 含誘發口徑：替代口徑 × (P2030/P2026)^ε；價格皆以人類叫車每英里定錨，P2030/P2026＝倍數÷(1+溢價)
    return f"{gb30_tpl(v)}*({v['P_REL_2030']}/(1+{v['P_PREM_2026']}))^{v['P_ELAST']}"


TARGETS = [
    ("T1", "2030 美國 Robotaxi 滲透率（趟數）", lambda v: pen30_tpl(v)[0], "0.00%",
     ["S_COV_2030", "S_SMAX", "S_T0", "R_COV_2026", "R_CA_SHARE", "R_OTHERS", "M_LYFT_NA", "M_UBER_SPEND"]),
    ("T2", "2030 Robotaxi 訂單金額（替代口徑，百萬美元）", gb30_tpl, "#,##0",
     ["S_COV_2030", "S_SMAX", "S_T0", "R_COV_2026", "R_CA_SHARE", "R_OTHERS", "M_LYFT_NA", "M_UBER_SPEND", "M_GROWTH", "P_REL_2030"]),
    ("T4", "2030 Robotaxi 訂單金額（含誘發，百萬美元）", gbj30_tpl, "#,##0",
     ["S_COV_2030", "S_SMAX", "S_T0", "R_COV_2026", "R_CA_SHARE", "M_UBER_SPEND", "M_GROWTH", "P_REL_2030", "P_ELAST"]),
]
ALLV = set()
for t in TARGETS:
    ALLV.update(t[4])
ALLV.update(["P_PREM_2026"])
NAME = {it[0]: it[1] for it in INPUTS if it[0] != "SEC"}

hdr(wsS, 1, ["承重變數排序（單變數擺動；其他維持基準）", "", "", "", "", "", "", ""], [44, 14, 14, 14, 14, 10, 10, 50])
wsS.cell(row=2, column=1, value="讀法：每列只把一個變數換成其低值與高值；擺動＝|高−低|。依情境的變數，低／高＝低／高情境值。先看排序再看情境頁的基準結果。")
r = 4
SENS_TOP = {}
for tid, tname, tpl, fmt, vars_ in TARGETS:
    c = wsS.cell(row=r, column=1, value=f"{tid}｜{tname}"); c.font = BOLD
    for j in range(1, 9):
        wsS.cell(row=r, column=j).fill = SEC_FILL
    r += 1
    base = {k: I(k) for k in ALLV}
    wsS.cell(row=r, column=1, value="基準值")
    c = wsS.cell(row=r, column=2, value="=" + tpl(base)); c.number_format = fmt
    base_row = r
    r += 1
    hdr(wsS, r, ["變數", "低值時", "高值時", "擺動", "排名", "ID", "", "說明"])
    r += 1
    first = r
    for vid in vars_:
        lo = dict(base); lo[vid] = I(vid, "L")
        hi = dict(base); hi[vid] = I(vid, "H")
        wsS.cell(row=r, column=1, value=NAME[vid])
        c = wsS.cell(row=r, column=2, value="=" + tpl(lo)); c.number_format = fmt
        c = wsS.cell(row=r, column=3, value="=" + tpl(hi)); c.number_format = fmt
        c = wsS.cell(row=r, column=4, value=f"=ABS(C{r}-B{r})+ROW()*1E-12"); c.number_format = fmt
        wsS.cell(row=r, column=6, value=vid)
        r += 1
    last_ = r - 1
    for rr_ in range(first, last_ + 1):
        wsS.cell(row=rr_, column=5, value=f"=RANK(D{rr_},$D${first}:$D${last_})")
    # 排序後前三
    wsS.cell(row=r, column=1, value="前三名（依擺動排序）").font = BOLD
    r += 1
    SENS_TOP[tid] = []
    for n in range(1, min(3, len(vars_)) + 1):
        wsS.cell(row=r, column=1, value=f"=INDEX($A${first}:$A${last_},MATCH(LARGE($D${first}:$D${last_},{n}),$D${first}:$D${last_},0))")
        c = wsS.cell(row=r, column=4, value=f"=LARGE($D${first}:$D${last_},{n})"); c.number_format = fmt
        wsS.cell(row=r, column=5, value=n)
        SENS_TOP[tid].append(r)
        r += 1
    r += 1

# 成本目標（v0.2）：2030 Cybercab、2030 Waymo 每付費英里全成本；變數讀成本輸入 2030 欄（低成本／高成本）
CTARGETS = [("T3", "2030 Tesla Cybercab 每付費英里全成本（US$）", "TC"), ("T5", "2030 Waymo 每付費英里全成本（US$）", "W")]
for tid, tname, op in CTARGETS:
    c = wsS.cell(row=r, column=1, value=f"{tid}｜{tname}"); c.font = BOLD
    for j in range(1, 9):
        wsS.cell(row=r, column=j).fill = SEC_FILL
    r += 1
    base_p = params(op, 2030, "B")
    wsS.cell(row=r, column=1, value="基準值")
    c = wsS.cell(row=r, column=2, value="=" + full_expr(base_p)); c.number_format = "0.00"
    r += 1
    hdr(wsS, r, ["變數", "低成本值時", "高成本值時", "擺動", "排名", "ID", "", "說明"])
    r += 1
    first = r
    vlist = []
    for k in base_p:
        if base_p[k] is None or k in ("Y", "PRICE"):
            continue
        ref = base_p[k]
        if ref in [v for _, v in vlist]:
            continue
        vlist.append((k, ref))
    for k, ref in vlist:
        def at(col):
            q = dict(base_p)
            for kk, vv in base_p.items():
                if vv == ref:
                    q[kk] = vv.replace(f"${dcol(2030, 'B', 5)}$", f"${dcol(2030, col, 5)}$")
            return q
        fL = full_expr(at("L")); fH = full_expr(at("H"))
        assert len(fL) < 7900
        oid = [o for (o, pid), rr2 in KROW.items() if ref.endswith(f"${rr2}")][0]
        pid = [pid for (o, pid), rr2 in KROW.items() if ref.endswith(f"${rr2}")][0]
        wsS.cell(row=r, column=1, value=f"{OPNAME[oid]}｜{PNAME[pid][0]}")
        c = wsS.cell(row=r, column=2, value="=" + fL); c.number_format = "0.00"
        c = wsS.cell(row=r, column=3, value="=" + fH); c.number_format = "0.00"
        c = wsS.cell(row=r, column=4, value=f"=ABS(C{r}-B{r})+ROW()*1E-12"); c.number_format = "0.00"
        wsS.cell(row=r, column=6, value=f"{oid}_{pid}")
        r += 1
    last_ = r - 1
    for rr_ in range(first, last_ + 1):
        wsS.cell(row=rr_, column=5, value=f"=RANK(D{rr_},$D${first}:$D${last_})")
    wsS.cell(row=r, column=1, value="前三名（依擺動排序）").font = BOLD
    r += 1
    SENS_TOP[tid] = []
    for n in range(1, 4):
        wsS.cell(row=r, column=1, value=f"=INDEX($A${first}:$A${last_},MATCH(LARGE($D${first}:$D${last_},{n}),$D${first}:$D${last_},0))")
        c = wsS.cell(row=r, column=4, value=f"=LARGE($D${first}:$D${last_},{n})"); c.number_format = "0.00"
        wsS.cell(row=r, column=5, value=n)
        SENS_TOP[tid].append(r)
        r += 1
    r += 1

# ---------------------------------------------------------------- 下游取數
hdr(wsO, 1, ["ID", "名稱", "低", "基準", "高", "單位", "口徑", "來源儲存格（基準）"], [30, 46, 13, 13, 13, 12, 60, 22])
OUTR = []


def out(oid, name, refs, unit, scope, fmt):
    OUTR.append((oid, name, refs, unit, scope, fmt))


out("RT_US_TRIPS_2026", "美國叫車年趟數 2026", [C('us_trips')] * 3, "百萬趟", "Uber＋Lyft（推估，含加拿大）；由 Lyft 趟數÷趟數份額", "#,##0")
out("RT_US_MILES_2026", "美國叫車年付費英里 2026", [C('us_miles')] * 3, "百萬英里", "趟數×每趟付費英里（基準 6.0）", "#,##0")
out("RT_US_GB_2026", "美國叫車年訂單金額 2026", [C('us_gb')] * 3, "百萬美元", "趟數×加權車資（司機端觀察車資）", "#,##0")
for y in (2028, 2030, 2032, 2035):
    out(f"RT_US_TRIPS_{y}", f"美國叫車年趟數 {y}", [N(s, 'us_trips', y) for s in ("低", "基準", "高")], "百萬趟", "不含 Robotaxi 誘發需求", "#,##0")
for y in YEARS:
    out(f"RT_PEN_{y}", f"Robotaxi 滲透率（趟數）{y}", [N(s, 'pen', y) for s in ("低", "基準", "高")], "%", "替代趟數÷美國叫車趟數；2026＝年中年化", "0.00%")
for y in YEARS:
    out(f"RT_PRICE_{y}", f"Robotaxi 每付費英里實收價格 {y}", [N(s, 'price', y) for s in ("低", "基準", "高")], "US$/英里", "全美營運商平均；名目、不含小費", "0.00")
for y in (2026, 2028, 2030, 2032, 2035):
    out(f"RT_COST_{y}", f"Robotaxi 全美加權單位營運成本 {y}", [N(s, 'cost', y) for s in ("低", "基準", "高")], "US$/英里", "v0.2 起讀堆疊頁：營運商份額加權的單位營運成本（不含研發、平台）；低／高＝低／高滲透情境（配高／低成本）", "0.00")
for y in (2026, 2030, 2035):
    out(f"RT_TRIPS_{y}", f"美國 Robotaxi 替代趟數 {y}", [N(s, 'rt_trips', y) for s in ("低", "基準", "高")], "百萬趟", "滲透模型口徑（不含誘發）", "#,##0.0")
    out(f"RT_GB_{y}", f"美國 Robotaxi 訂單金額（替代）{y}", [N(s, 'rt_gb', y) for s in ("低", "基準", "高")], "百萬美元", "替代英里×實收價格", "#,##0")
    out(f"RT_GB_J_{y}", f"美國 Robotaxi 訂單金額（含誘發）{y}", [N(s, 'rt_gb_j', y) for s in ("低", "基準", "高")], "百萬美元", "傑文斯口徑；與替代口徑擇一", "#,##0")
    out(f"RT_WAYMO_TRIPS_{y}", f"Waymo 美國趟數 {y}", [N(s, 'w_trips', y) for s in ("低", "基準", "高")], "百萬趟", "替代口徑×Waymo 份額", "#,##0.0")
    out(f"RT_TESLA_TRIPS_{y}", f"Tesla 美國 Robotaxi 趟數 {y}", [N(s, 't_trips', y) for s in ("低", "基準", "高")], "百萬趟", "替代口徑×Tesla 份額", "#,##0.0")
    out(f"RT_FLEET_{y}", f"美國 Robotaxi 所需車隊 {y}", [N(s, 'fleet', y) for s in ("低", "基準", "高")], "台", "替代趟數÷(每車每日趟數×365)", "#,##0")
    out(f"RT_PLAT_REV_{y}", f"Uber／Lyft 平台 Robotaxi 抽成收入 {y}", [N(s, 'plat_rev', y) for s in ("低", "基準", "高")], "百萬美元", "訂單金額×經平台占比×22.44% 抽成", "#,##0")
for y in (2027, 2030, 2035):
    out(f"RT_VEH_ADDS_{y}", f"美國 Robotaxi 年新增車輛需求 {y}", [N(s, 'adds', y) for s in ("低", "基準", "高")], "台", "擴車＋汰換", "#,##0")
out("RT_WAYMO_WEEKLY_2026", "Waymo 每週付費趟數（2026 Q2 推估）", [C('waymo_weekly')] * 3, "趟/週", "CPUC 加州÷加州占比（官方最新 50 萬+，2026-03）", "#,##0")
out("RT_PRICE_HUMAN", "價格錨 1：人類叫車每付費英里價格", [C('p_human')] * 3, "US$/英里", "加權車資÷每趟付費英里", "0.00")
out("RT_PRICE_OWN_AVG", "價格錨 2：自有新車每英里全成本（AAA 平均）", [I('P_OWN_AVG')] * 3, "US$/英里", "AAA 2025", "0.00")
out("RT_PRICE_OWN_SEDAN", "價格錨 2b：自有小轎車每英里全成本", [I('P_OWN_SEDAN')] * 3, "US$/英里", "AAA 2025", "0.00")
out("RT_JEVONS_PRICE", "傑文斯門檻價格（叫車英里翻倍）", [C('j_double', x) for x in "LBH"], "US$/英里", "人類叫車每英里×2^(1/ε)；ε＝−0.6／−1.0／−1.8", "0.00")
out("RT_JEVONS_REV_2030", "2030 降價後總營收倍數", [C('j_rev30', x) for x in "LBH"], "倍", "(P2030÷P2026)^(1+ε)", "0.00")
out("RT_POOL_RH", "叫車池（美國叫車年付費英里）", [C('pool_rh')] * 3, "十億英里", "2026", "#,##0.0")
out("RT_POOL_OWN", "自有車池（美國都會輕型車 VMT）", [C('pool_own')] * 3, "十億英里", "FHWA×輕型車×都會占比", "#,##0")
# v0.2：各營運商每英里成本路徑（本線權威；低／高＝低成本／高成本）
for op in OPS4 + ["T"]:
    oid = OPID.get(op, "TSLA_FLEET")
    onm = OPNAME.get(op, "Tesla 合計（車隊英里加權）")
    for y in PYEARS:
        out(f"RT_CPM_{oid}_{y}", f"{onm} 每付費英里全成本 {y}", [PATH[(op, 'full', s_, y)] for s_ in SCS], "US$/英里", "含研發攤提與平台抽成；奇數年為錨點幾何平均", "0.00")
    for y in PYEARS:
        out(f"RT_OPEX_{oid}_{y}", f"{onm} 每付費英里單位營運成本 {y}", [PATH[(op, 'unit', s_, y)] for s_ in SCS], "US$/英里", "不含研發、平台", "0.00")
for op in OPS4:
    for y in YRS4:
        out(f"RT_REMOTE_{OPID[op]}_{y}", f"{OPNAME[op]} 遠端營運成本每付費英里 {y}", [XR[(f'{op}_remote', y, s_)] for s_ in SCS], "US$/英里", "接管轉成本頁", "0.000")
        out(f"RT_MONITOR_{OPID[op]}_{y}", f"{OPNAME[op]} 安全員／追車成本每付費英里 {y}", [XR[(f'{op}_monitor', y, s_)] for s_ in SCS], "US$/英里", "過渡成本", "0.000")
    for y in YRS4:
        out(f"RT_PMD_{OPID[op]}_{y}", f"{OPNAME[op]} 每車每日付費英里 {y}", [P(op, 'PMD', y, s_) for s_ in SCS], "英里/日", "稼動率（承重）；低欄＝低成本＝高稼動", "#,##0")
        out(f"RT_OPPRICE_{OPID[op]}_{y}", f"{OPNAME[op]} 每付費英里價格 {y}", [P(op, 'PRICE', y, s_) for s_ in SCS], "US$/英里", "低／高＝低價／高價", "0.00")
for y in YRS4:
    out(f"RT_TSLA_MPI_{y}", f"Tesla 城市關鍵接管里程 {y}", [XR[('t_mpi', y, s_)] for s_ in SCS], "英里/次", "社群 tracker 學習曲線外推；低欄＝快速改善", "#,##0")
for op in OPS4 + ["T"]:
    oid = OPID.get(op, "TSLA_FLEET")
    onm = OPNAME.get(op, "Tesla 合計")
    out(f"RT_YEAR_1USD_{oid}", f"{onm} 全成本跨 1 美元年份", [PATH[(op, 'full', s_, 'x1')] for s_ in SCS], "年", "2026–2032；未跨過顯示文字", "0")
    out(f"RT_YEAR_1USD_OPEX_{oid}", f"{onm} 單位營運成本跨 1 美元年份", [PATH[(op, 'unit', s_, 'x1')] for s_ in SCS], "年", "", "0")
    out(f"RT_YEAR_AAA_OPEX_{oid}", f"{onm} 單位營運成本跨 AAA 0.77 年份", [PATH[(op, 'unit', s_, 'xaaa')] for s_ in SCS], "年", "", "0")
for i, (oid, name, refs, unit, scope, fmt) in enumerate(OUTR, 2):
    wsO.cell(row=i, column=1, value=oid)
    wsO.cell(row=i, column=2, value=name)
    for j, ref in enumerate(refs):
        c = wsO.cell(row=i, column=3 + j, value=f"={ref}"); c.number_format = fmt
    wsO.cell(row=i, column=6, value=unit)
    wsO.cell(row=i, column=7, value=scope).alignment = WRAP
    wsO.cell(row=i, column=8, value=refs[1])
n_out = len(OUTR) + 1
wsO.cell(row=n_out + 2, column=1, value="跨錨年份").font = BOLD
for j, (k, lab) in enumerate([("x_double", "傑文斯門檻跨過年份"), ("x_own", "自有車池打開年份")]):
    R_ = n_out + 3 + j
    wsO.cell(row=R_, column=1, value=f"RT_YEAR_{'JEVONS' if k=='x_double' else 'OWNPOOL'}")
    wsO.cell(row=R_, column=2, value=lab)
    for m, scn in enumerate(("低", "基準", "高")):
        wsO.cell(row=R_, column=3 + m, value=f"='情境'!$C${NR[(scn, k)]}")
    wsO.cell(row=R_, column=6, value="年")
wsO.freeze_panes = "C2"

# ---------------------------------------------------------------- 追蹤指標
hdr(wsT, 1, ["指標", "最新值", "日期", "頻率", "資料來源", "標記", "為何重要"], [36, 34, 14, 10, 60, 18, 40])
TRACK = [
    ("Waymo 每週付費趟數（官方）", "50 萬+", "2026-03", "里程碑", S["WAYMO500"], "[Interested-party]", "滲透主指標；2026 年底目標每週 100 萬"),
    ("Waymo 加州月趟數（CPUC）", "1,348,926（下限）", "2026-06", "季（落後約 1 季）", S["CPUC"], "[Verified]", "唯一法定揭露的趟數與里程；可推全美"),
    ("Waymo 加州請求趟數（CPUC）", "390 萬（Q4 2025，季增 40%）", "2025 Q4", "季", "https://robotaxitracker.com/cpuc", "[Verified]", "需求面；履約率 95%"),
    ("Waymo 空駛比例（加州）", "43–45%", "2025 H2", "季", S["FIND"], "[Verified]", "付費英里占比，直接影響每付費英里成本"),
    ("Waymo 商轉城市數", "15 城", "2026-10-28", "季", S["GOOG"], "[Interested-party]", "覆蓋率；下一步加拿大、德國、新加坡、西班牙"),
    ("Waymo 車隊", "約 3,500（媒體）；德州登記 1,297", "2026-07／2026-11-10", "不定期", S["ELEC"] + "；" + S["TXREG"], "[Interested-party]", "供給面、每車日趟數"),
    ("Waymo 叫車訂單金額份額", "舊金山 15%、洛杉磯 15%、鳳凰城 16%", "2026-06", "月", S["YIPIT"], "[Verified]（第三方信用卡／電子收據樣本）", "都會區內份額 S 的直接觀察"),
    ("Waymo 累計全自駕里程", "3 億英里+", "2026-10-28", "季", S["GOOG"], "[Interested-party]", ""),
    ("Alphabet Other Bets 營業虧損", "20.5 億美元（營收 3.72 億）", "2026 Q3", "季", "https://fortune.com/2026/10/30/waymo-126-billion-alphabet-372-million-2-05-billion-google", "[Verified]", "Waymo 燒錢代理"),
    ("Tesla 無人營運城市", "6 城", "2026-10-22", "季", S["TSLAQ3"], "[Interested-party]", "Tesla 份額路徑"),
    ("Tesla 德州 Robotaxi 登記數", "1,302（Model Y 729、Cybercab 573）", "2026-11-10", "週", S["TXREG"], "[Verified]（登記≠營運）", "供給面；與 Waymo 1,297 相當"),
    ("Tesla 無安全員里程", "310 萬英里（2026-01-22～10-18）", "2026-10-18", "週", S["TXREG"], "[Interested-party]（第三方追蹤）", ""),
    ("Zoox 車隊與城市", "約 100+ 台；拉斯維加斯收費，4 城", "2026-08-05", "不定期", S["ZOOX"], "[Interested-party]", "豁免上限每年 2,500 台"),
    ("Uber 平台上無人車", "年底約 700 台；2027 年 2,000+", "2026 Q3", "季", S["UBER"], "[Interested-party]", "平台 vs 直營分潤"),
    ("Uber 全球趟數", "40.61 億（年增 16%）", "2026 Q3", "季", S["UBER"], "[Verified]", "叫車池（含外送）"),
    ("Lyft 趟數", "2.735 億（年增 10%）", "2026 Q3", "季", S["LYFT8K"], "[Verified]", "叫車池基期"),
    ("人類叫車車資與抽成", "Uber 24.30／Lyft 19.76 美元；抽成 22.44%；司機 2.65 美元/英里", "2026 Q3", "季", S["GRID"], "[Verified]（司機樣本）", "價格錨 1"),
    ("灣區叫車價格比較", "Waymo 19.69、Uber 17.47、Lyft 15.47 美元/趟；Waymo 5.72 美元/英里（v0.1 誤植每公里）；Tesla 1.99 美元/英里（有安全員）", "2025-11～2026-01", "半年", S["OBI"], "[Verified]（模擬叫車）", "每英里實收價格（核心指標）"),
    ("自有車每英里成本", "0.77（平均）、0.56（小轎車）美元", "2025-09", "年", S["AAA"], "[Verified]", "價格錨 2：打開自有車池"),
    ("Apollo Go（百度）", "28 城；全無人累計 2.4 億公里+", "2026 Q2", "季", "https://baidu.gcs-web.com/news-releases/news-release-details/baidu-announces-second-quarter-2026-results", "[Interested-party]", "中國滲透（本版未建模）"),
    ("小馬智行車隊", "1,975 台（目標年底 3,500+）；廣州、深圳全城 UE 平衡", "2026-06-30", "半年", "https://www.sec.gov/Archives/edgar/data/0001969302/000110465926109772/tm2625374d1_ex99-1.pdf", "[Interested-party]", "中國滲透；亞太掛牌"),
    ("文遠知行", "Q3 財報 2026-11-20 公布（未取得）", "2026-11-20", "季", "https://www.panabee.com/news/weride-earnings-q3-2026", "—", "亞太掛牌"),
    # v0.2 新增：接管、安全員、稼動、價格、事故
    ("【v0.2】FSD 社群關鍵接管里程（城市／全部）", "v14.3：853／1,750 英里/次；v13.2：217／444；v12.3：121／177", "2026-06", "月", S["FSDT"], "[Interested-party]（社群自願填報）", "學習曲線斜率→遠端協助率與理賠率（接管轉成本頁）"),
    ("【v0.2】Tesla Robotaxi 安全報告", "無監督 310 萬英里（2026-01-22～10-18），0 受傷、0 氣囊；分城：奧斯汀 212 萬、達拉斯 61 萬、休士頓 23 萬", "2026-10-21", "季", S["TSAFE"], "[Interested-party]", "無監督里程規模；未含 SGO 輕微事故件數"),
    ("【v0.2】CA DMV disengagement（每次接管英里）", "Waymo 19,234；Zoox 60,682；Nuro 646（只含有安全員測試）", "2026-02（2025 報告年）", "年", S["CADMV"], "[Verified]", "公開唯一的法定接管資料；不含無人商轉與遠端協助"),
    ("【v0.2】CPUC 季報（無人載客）", "Waymo 加州趟數、乘客、總英里（見 CPUC資料頁）；不含遠端協助次數", "2026-06", "季", S["CPUC"], "[Verified]", "稼動率與空駛比"),
    ("【v0.2】NHTSA SGO 事故（Tesla Robotaxi）", "14 起／約 80 萬英里≈每 5.7 萬英里 1 起（皆有安全員）", "2026-01", "月", S["ELECFEB"], "[Verified]（件數）／[Interested-party]（里程外推）", "事故率→理賠成本"),
    ("【v0.2】Waymo 事故與理賠", "IIHS：警方可報事故 −68%、傷害 −81%；Swiss Re：責任理賠 0.43 件/百萬英里（人類 4.1）", "2026-07／2024-12", "半年", S["IIHS"] + "；" + S["SWISSRE"], "[Verified]／[Interested-party]", "保險成本"),
    ("【v0.2】遠端協助頻率與監看比", "7 家皆未揭露頻率；Waymo 約半數遠端人員在菲律賓；Tesla 遠端可接手 ≤10 mph", "2026-03", "不定期", S["TCMARKEY"], "[Verified]（回函）", "遠端營運成本最大不確定；等待聯邦規範或公司揭露"),
    ("【v0.2】安全員移除城市數", "Tesla：6 城無監督（奧斯汀、達拉斯、休士頓、奧蘭多、坦帕、邁阿密）；灣區仍有安全員；Waymo 商轉 15 城全無人", "2026-10", "月", S["TSAFE"] + "；" + S["TNW8"] + "；" + S["GOOG"], "[Interested-party]", "過渡成本何時消失"),
    ("【v0.2】Tesla 無監督車數（社群）", "2026-04：25 台（奧斯汀 19）；2026-08：奧斯汀追蹤 54 台、德州兩城約 30 台", "2026-08", "週", S["AW25"] + "；" + S["TNW8"], "[Interested-party]（社群）", "車隊規模→稼動率與場站規模"),
    ("【v0.2】每車每日付費英里", "Waymo 約 95（每車日 25.5 趟×CPUC 每趟 3.7 付費英里，推估）；Tesla 未揭露（推估每車日 38–115 車英里）", "2026 Q2／2026-10", "季", S["CPUC"] + "；" + S["TSAFE"], "[Derived]", "稼動率（承重變數）"),
    ("【v0.2】空駛比", "Waymo 加州 43–45%；Tesla 未揭露", "2025 H2", "季", S["FIND"], "[Verified]", "付費占比"),
    ("【v0.2】每英里價格（依城市、行程長度）", "Tesla 奧斯汀 3.00＋1.40/英里（5 英里≈2.00/英里，10 英里≈1.70）；Waymo 灣區 5.72/英里；Cybercab 同路線較 Model Y 低約 35%；武漢 Apollo 約 0.34 美元/英里（2024，有補貼）", "2026", "月", S["TFARE"] + "；" + S["OBI"] + "；" + S["TCC"], "[Interested-party]／[Verified]", "價格路徑與補貼檢驗（每英里價格頁）"),
    ("【v0.2】中國單車成本", "百度 RT6 204,600 元；小馬 2027 車型 <23 萬元、七代套件 BOM −70%；北京遠端 1:3 標準", "2026-05", "半年", S["GASGOO"] + "；" + S["SSENI"], "[Interested-party]", "中國每英里成本"),
    ("社群追蹤（robotaxitracker.com）", "車輛目擊、等候時間、價格、政府資料彙整", "每日", "日", "https://robotaxitracker.com/?city=all", "[Interested-party]（社群提交，方法未完整公開）", "最即時的早期訊號；數字須以 CPUC、TxDMV、公司揭露核對"),
    ("NHTSA 召回與事故", "Waymo 校車事件升級召回調查（2,895 台）", "2026-11", "即時", "https://driveteslacanada.ca/news/nhtsa-waymo-rq26006-school-bus-recall-query/", "[Verified]", "停駛風險"),
]
for i, rowv in enumerate(TRACK, 2):
    for j, v in enumerate(rowv, 1):
        wsT.cell(row=i, column=j, value=v).alignment = WRAP
k0 = len(TRACK) + 3
wsT.cell(row=k0, column=1, value="資料來源評比（滲透指標）").font = BOLD
SRCR = [
    ("1. CPUC 季報（加州）", "法定、逐月趟數／乘客／里程；部分遮蔽、落後約一季、只含加州"),
    ("2. TxDMV AV 報告（德州）", "法定登記車數；登記≠營運"),
    ("3. 公司揭露", "Waymo 里程碑、Alphabet 法說、Tesla 法說、Uber／Lyft 季報；有利害關係、口徑不一"),
    ("4. YipitData／Bloomberg Second Measure", "第三方收據或刷卡樣本的訂單金額份額；最接近「叫車市場份額」"),
    ("5. Obi、Gridwise", "價格與司機收入；每英里實收價格的主要來源"),
    ("6. robotaxitracker.com 等社群追蹤", "最即時，但方法不完整公開；作為早期訊號，數字以 1–3 核對"),
]
for j, (a, b) in enumerate(SRCR, 1):
    wsT.cell(row=k0 + j, column=1, value=a)
    wsT.cell(row=k0 + j, column=2, value=b)

# ---------------------------------------------------------------- 說明
lines = [
    ("線 1a Robotaxi 單位經濟模型 " + VERSION + "（2026-11-16）", True),
    ("核心問題：美國 Robotaxi 在叫車市場的滲透狀況（城市、車隊、每週付費趟數、份額）與每英里實收價格路徑；價格降到哪個錨會打開哪個市場池（傑文斯檢驗）。收支平衡只作輔助。", False),
    ("範圍：滲透與傑文斯為美國口徑；v0.2 起每英里成本堆疊另含中國（百度／小馬平均，美元計）。框架：線1a_Robotaxi_框架_v0.1.md。", False),
    ("v0.2（Andy 2026-11-16 22:57）：比照 Tokenomics「每 GW 成本與營收」，算每英里成本與每英里價格；預估 2030–2032 每英里成本能否降到約 1 美元（才能滲透自有車市場）；診斷 Tesla 號稱低成本但目前不是的原因；把接管／遠端介入轉成成本；追蹤每英里價格。", False),
    ("", False),
    ("頁面", True),
    ("輸入：藍字＝輸入；黃底＝判斷（[Assumed]／[Analogy]），可改。用法欄：共用＝三情境都用基準、低／高只進敏感度；依情境＝低／基準／高情境各用一欄；參考＝不進計算。", False),
    ("成本輸入（v0.2）：四個營運商（Waymo、Tesla Model Y、Tesla Cybercab、中國）× 2026／2028／2030／2032 × 低成本／基準／高成本。低／高指成本，不是數值。", False),
    ("接管轉成本（v0.2）：FSD 社群接管里程學習曲線 → 遠端協助率、理賠率；遠端監看比、安全員、事故率 → 每付費英里成本。", False),
    ("每英里成本堆疊（v0.2，本線權威）：十項成本、單位營運成本（1–8）、全成本（1–10）、與 1 美元及 AAA 0.77 的差距、2026–2032 年度路徑與跨門檻年份。", False),
    ("Tesla成本落差診斷（v0.2）：七個假說逐項量化（單獨／最後／平均改善）、瀑布、價格補貼檢驗、結論。", False),
    ("每英里價格（v0.2）：觀察價格（含 Obi 單位更正）、Tesla 票價依行程長度、營運商價格路徑 vs 成本。", False),
    ("CPUC資料：Waymo 加州月資料（法定揭露），用來把官方「每週 50 萬+」更新到 2026 Q2。", False),
    ("計算：市場基期、現況校準、S 曲線參數、每英里成本與價格（框架第 3 節等式）、傑文斯檢驗。", False),
    ("敏感度：先看承重變數排序（T1 2030 滲透率、T2 2030 訂單金額替代口徑、T3 2030 Cybercab 每英里全成本、T4 2030 訂單金額含誘發、T5 2030 Waymo 每英里全成本），再看情境頁。", False),
    ("情境：2026–2035，低（低滲透）／基準／高（高滲透）三組；含外部預測（GS、MS）對照。", False),
    ("下游取數：給 Waymo（Alphabet）、Tesla Robotaxi、Uber、1b 智慧車等模型讀的輸出；ID 固定。", False),
    ("追蹤指標：最新值與日期、資料來源評比。", False),
    ("", False),
    ("方法要點", True),
    ("1. 滲透率＝商轉都會覆蓋率×都會區內份額 S(t)；S(t) 為 logistic，斜率由 2026 校準點反推，所以三情境 2026 值相同。", False),
    ("2. 2026 校準：CPUC 加州 Q2 趟數÷13 週÷加州占比 → Waymo 全美每週趟數；÷(1−其他營運商占比) → 全美 Robotaxi；÷美國叫車趟數 → 滲透率。", False),
    ("3. 美國叫車趟數：Lyft 單季趟數×4×北美汽車占比÷Lyft 趟數份額（消費份額 76/24 依車資換算）。未取得官方全美總趟數。", False),
    ("4. 每英里實收價格：人類叫車每付費英里×倍數（2026 用 Obi 溢價 12.7%；2030、2035 為情境假設）。", False),
    ("5. 傑文斯：常數彈性 ε（−0.6／−1.0／−1.8，類比埃及 Uber 實驗）。誘發倍數＝(P/P2026)^ε；營收倍數＝(P/P2026)^(1+ε)，|ε|>1 才擴大營收。傑文斯門檻價格＝叫車英里翻倍所需價格＝人類叫車每英里×2^(1/ε)。自有車池在價格≤AAA 每英里全成本時才打開，本版只標年份、不把自有車池需求加進營收。常數彈性外推超出實驗範圍（實驗最大折扣 50%、為期 3 個月、埃及），高情境 2035 價格較 2026 低約七成，誘發倍數應視為上緣。", False),
    ("6. 替代口徑與含誘發口徑不可相加；下游擇一並註明。", False),
    ("", False),
    ("已知限制", True),
    ("Waymo 全美每週趟數是推估（官方 2026-03 後未更新）；車隊數媒體轉述不一致；保險、遠端協助、場站成本無公司揭露；美國叫車總趟數為推估；價格為名目、未計通膨；未建模中國；未建模叫車市場因 Robotaxi 擴大以外的結構變化（如政策禁令）。", False),
]
ws0.column_dimensions["A"].width = 130
for i, (t, b) in enumerate(lines, 1):
    c = ws0.cell(row=i, column=1, value=t)
    c.alignment = WRAP
    if b:
        c.font = BOLD
r = len(lines) + 2
ws0.cell(row=r, column=1, value="敏感度前三名（2030 滲透率）").font = BOLD
for n, rr_ in enumerate(SENS_TOP["T1"], 1):
    ws0.cell(row=r + n, column=1, value=f"=\"{n}. \"&'敏感度'!A{rr_}")
r += 5
ws0.cell(row=r, column=1, value="敏感度前三名（2030 Cybercab 每付費英里全成本）").font = BOLD
for n, rr_ in enumerate(SENS_TOP["T3"], 1):
    ws0.cell(row=r + n, column=1, value=f"=\"{n}. \"&'敏感度'!A{rr_}")

r += 5
ws0.cell(row=r, column=1, value="敏感度前三名（2030 Waymo 每付費英里全成本）").font = BOLD
for n, rr_ in enumerate(SENS_TOP["T5"], 1):
    ws0.cell(row=r + n, column=1, value=f"=\"{n}. \"&'敏感度'!A{rr_}")
r += 5
ws0.cell(row=r, column=1, value="Tesla 成本落差主因").font = BOLD
ws0.cell(row=r + 1, column=1, value=f"=\"1. \"&{GD['main1']}")
ws0.cell(row=r + 2, column=1, value=f"=\"2. \"&{GD['main2']}")

os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)
print("wrote", OUT)

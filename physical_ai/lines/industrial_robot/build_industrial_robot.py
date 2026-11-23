# 線 2a 工業機器人（含協作）單位經濟模型 v0.1
# 用法：python3 physical_ai/lines/industrial_robot/build_industrial_robot.py
#       之後跑 check_industrial_robot.py（LibreOffice headless 重算、0 錯誤值、對帳檢查）
# 依據：線2a_工業機器人_框架_v0.1.md（2026-11-16）第 3 節等式、第 4 節門檻、第 5 節追蹤指標
# 頁面：說明｜輸入｜計算｜情境｜敏感度｜下游取數｜追蹤指標
# 單位：台數＝台；減速機＝顆；市場值＝百萬美元（USD m）；單台回本＝人民幣（RMB）
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "dist", "20261116_工業機器人單位經濟_v0.1.xlsx")
VERSION = "v0.1"
BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=13)
HDR = PatternFill("solid", fgColor="DDEBF7")
JUDG = PatternFill("solid", fgColor="FFF2CC")
SEC = PatternFill("solid", fgColor="E2EFDA")
WRAP = Alignment(wrap_text=True, vertical="top")
YEARS = list(range(2025, 2033))          # 2025 為實績（IFR），2026–2032 為情境
YCOL = {y: CL(4 + i) for i, y in enumerate(YEARS)}   # D=2025 … K=2032
SCN = [("低", "D"), ("基準", "E"), ("高", "F")]

# ---------------- 來源 ----------------
S = {
 "ifr_es": "IFR World Robotics 2026 Executive Summary（全球裝機 603,307 台、市場值 186 億美元、中國 354,226 台、2026 預測 65.5 萬、2029 預測 80.6 萬）https://ifr.org/img/worldrobotics/Executive_Summary_WR_2026_-_Industrial_Robots.pdf",
 "ifr_pr": "IFR 新聞稿 2026-09-24（國產 55%、日本 36,219、美國約 38,500、韓國 30,000）https://ifr.org/ifr-press-releases/news/five-million-robots-now-operate-in-factories-globally",
 "ifr_pres": "IFR President's Report 2026（協作 6.6 萬台、占 11%、+7%）https://ifr.org/ifr-press-releases/news/presidents-report-by-jane-heffner-3-2026",
 "fxg_red": "faxiangongchang《China Reducer Deep Dive 2026》（6 軸：RV 3–4、諧波 2–4；協作 6–7 諧波；減速機約占本體成本 30–35%；Nabtesco 全球 RV >35%；HDS 全球諧波 >50%）https://faxiangongchang.com/en/reports/china-reducer-deep-dive-2026",
 "fxg_ir": "faxiangongchang《China Industrial Robot 2026》（2024 中國產量 556,400 台，6 軸 60–65%、SCARA 15–18%、協作 10–15%、Delta 等 5–8%；國產 6 軸 20kg 10–15 萬元、外資 20–35 萬元；國產協作 8–15 萬元）https://faxiangongchang.com/en/reports/china-industrial-robot-2026",
 "fxg_cobot": "faxiangongchang《China Collaborative Robot 2026》（協作價格帶、BOM 占比、回本期案例）https://faxiangongchang.com/en/reports/china-collaborative-robot-2026",
 "dgzq": "東莞證券《機器人系列報告（五）之減速器》2025-01-27（引智研諮詢：RV 均價約 4,530 元、諧波約 2,048 元，2022；引華經：2023 全球 RV Nabtesco 61.41%）https://pdf.dfcfw.com/pdf/H3_AP202501271642597473_1.pdf",
 "gytx": "觀研天下（chinabaogao，引 IFR）：每台協作 6.5 顆諧波 https://www.chinabaogao.com/data/202304/631736.html",
 "ggii": "界面新聞引高工機器人 GGII：中國 6 軸 6kg 本體略高於 2 萬元、SCARA 約 1 萬元（2025）https://en.jiemian.com/article/13632808.html",
 "nbs_wage": "國家統計局 2026-05-15：2025 年城鎮私營單位製造業年平均工資 76,055 元（+6.4%）；非私營 113,594 元（+5.2%）https://www.stats.gov.cn/xxgk/sjfb/zxfb2020/202605/t20260515_1963707.html",
 "nbs_sep": "國家統計局 2026-10-19：2026 年 9 月工業機器人產量 96,752 台（+18.9%）；1–9 月 794,781 台（+29.1%）https://www.stats.gov.cn/xxgk/sjfb/zxfb2020/202610/t20261019_1965617.html",
 "nbs_2025": "同花順財經 2026-01-19 引國家統計局：2025 年工業機器人產量 +28.0%、機器人減速器 +63.9%（台數未列）https://news.10jqka.com.cn/20260119/c674113409.shtml",
 "jara_q3": "JARA 2026 年 7–9 月（會員）https://www.jara.jp/data/dl/quarter/IR2026-07-09.pdf",
 "jara_q2": "JARA 2026 年 4–6 月（會員）https://www.jara.jp/data/dl/quarter/IR2026-04-06.pdf",
 "jara_q1": "JARA 2026 年 1–3 月（會員）https://www.jara.jp/data/dl/quarter/IR2026-01-03.pdf",
 "jara_q4": "JARA 2025 年 10–12 月與 2025 全年（會員）https://jara.jp/data/dl/quarter/IR2025-10-12.pdf",
 "jara_fc": "mreport 轉述 JARA：2025 受注（含非會員）1.0456 兆日圓 +25.7%；2026 預測 1.22 兆 +16.7%（上修 1,900 億）https://www.mreport.co.th/en/news/trend-and-innovation/138-jara-industrial-robot-orders-2026-physical-ai",
 "hds_h1": "Harmonic Drive Systems 2027 年 3 月期第 2 季決算短信（2026-11-11）：上半年營收 36,124 百萬日圓 +29.8%、營業利益 4,205 百萬 https://finance-frontend-pc-dist.west.edge.storage-yahoo.jp/disclosure/20261111/20261111573783.pdf",
 "hds_rev": "HDS 業績預測修正（2026-11-11）：FY3/2027 營收 78,000 百萬日圓（原 74,500）；FY3/2026 實績 59,557 百萬；假設 157 日圓/美元、23 日圓/人民幣；Q2 受注「高水準」未揭露數字 https://finance-frontend-pc-dist.west.edge.storage-yahoo.jp/disclosure/20261111/20261111572884.pdf",
 "hds_q1": "HDS 2027 年 3 月期第 1 季決算短信（受注 241 億日圓 +55.7%）https://www2.jpx.co.jp/disc/63240/140120260804507706.pdf",
 "nab_cmp": "envalith 彙整 Nabtesco 分部：Component Solutions 2026 年 1–3 月外部營收 22,259 百萬日圓、受注 22,796 百萬、受注殘 20,203 百萬 https://envalith.com/en/companies/6268/segments",
 "nab_h1": "quartr 摘要 Nabtesco 2026 上半年：集團受注 1,985 億日圓 +33.1% https://quartr.com/events/nabtesco-corporation-6268-q2-2026_F9b29Dfj",
 "fanuc": "Fanuc 2026 年 9 月補充資料（Robot 部門受注 1,034 億日圓 +12.6%）https://www.fanuc.co.jp/en/ir/announce/pdf/2026/reference202609_e.pdf",
 "yaskawa": "technews 2026-10-12：安川 2026 年 6–8 月 Robotics 受注 776 億日圓 +33% https://cdnfinance.technews.tw/2026/10/12/yaskawa-orders-hit-record-raises-profit-forecast",
 "ter": "Zacks：Teradyne 2026Q3 Robotics 營收 9,600 萬美元 https://www.zacks.com/stock/news/3017640/teradyne-q3-earnings-beat-estimates,-revenues-increase-year-over-year",
 "a3": "A3 北美 2026Q3 機器人訂單 +24.8%（僅標題）https://www.automate.org/robotics/news/north-american-robot-orders-increase-24-8-in-third-quarter-of-2026/anc",
 "frame": "線2a_工業機器人_框架_v0.1.md（2026-11-16）",
}

# ---------------- 輸入定義 ----------------
# (ID, 項目, 單位, 低, 基準, 高, 標記, 類別, 出處, 說明)
# 低／高為 None → 等於基準（非判斷值）。低／高為數值區間；情境頁使用的變數（成長、占比、均價）低＝規模較小
INP = [
 ("sec", "A. 產業層：裝機實績與 IFR 預測"),
 ("G_INST_2025", "全球工業機器人裝機 2025（含協作）", "台", None, 603307, None, "Verified", "A", S["ifr_es"], "+11%"),
 ("CN_INST_2025", "中國裝機 2025", "台", None, 354226, None, "Verified", "A", S["ifr_es"], "+20%；占全球 59%"),
 ("CN_DOM_2025", "中國裝機中國產品牌占比 2025", "%", None, 0.55, None, "Verified", "A", S["ifr_pr"], "2024 年 57%"),
 ("COBOT_INST_2025", "全球協作機器人裝機 2025", "台", None, 66000, None, "Verified", "A", S["ifr_pres"], "IFR 報 6.6 萬台（取整），占 11%、+7%"),
 ("G_MV_2025", "全球裝機市場值 2025（本體）", "USD bn", None, 18.6, None, "Verified", "A", S["ifr_es"], "+11%；IFR 口徑為本體價值，不含整合與周邊"),
 ("IFR_F_2026", "IFR 預測全球裝機 2026", "台", None, 655000, None, "Interested-party", "A", S["ifr_es"], "IFR 為產業協會（會員為機器人廠），預測有推廣立場"),
 ("IFR_F_2029", "IFR 預測全球裝機 2029", "台", None, 806000, None, "Interested-party", "A", S["ifr_es"], "同上；2026–2029 年均約 7%"),
 ("sec", "B. 產業層：情境驅動（判斷）"),
 ("G26", "全球裝機年增 2026", "%", 0.03, 0.0857, 0.14, "Assumed", "A", S["ifr_es"] + "；" + S["jara_q3"] + "；" + S["nbs_sep"], "基準＝IFR 預測 655,000÷603,307−1；高＝訂單領先（JARA 受注 +46.6%、中國產量 1–9 月 +29.1%）；低＝循環轉弱"),
 ("G2729", "全球裝機年複合成長 2027–2029", "%", 0.03, 0.0716, 0.11, "Assumed", "A", S["ifr_es"], "基準＝(806÷655)^(1/3)−1"),
 ("G3032", "全球裝機年複合成長 2030–2032", "%", 0.02, 0.06, 0.10, "Assumed", "A", "—", "IFR 無 2030 後預測；高＝AI 降低整合成本擴大一般工業"),
 ("CN_SH", "中國占全球裝機比 2026–2032", "%", 0.55, 0.58, 0.62, "Assumed", "A", S["ifr_es"], "2025 實績 58.7%"),
 ("COBOT_SH_2032", "協作占全球裝機比 2032（由 2025 實績線性過渡）", "%", 0.11, 0.15, 0.20, "Assumed", "A", S["ifr_pres"] + "；" + S["fxg_ir"], "2025 實績 10.9%；產業報告預期中國 2030 協作占新裝 22–25%（Interested-party）"),
 ("MIX_ART", "非協作裝機中多關節（含 6 軸）占比", "%", 0.65, 0.71, 0.75, "Analogy", "P", S["fxg_ir"], "類比：中國 2024 產量結構 6 軸 60–65%、協作 10–15% → 扣除協作後約 71%"),
 ("MIX_SCARA", "非協作裝機中 SCARA 占比", "%", 0.15, 0.19, 0.22, "Analogy", "P", S["fxg_ir"], "類比：中國 2024 產量 SCARA 15–18% → 扣除協作後約 19%；其餘為直角座標、Delta 等"),
 ("ROBOT_ASP_G", "本體平均單價年變動", "%", -0.04, -0.01, 0.01, "Assumed", "I", S["ggii"], "中國價格戰（SCARA 一年腰斬）對比日系漲價；全球混合"),
 ("sec", "C. 每台減速機用量（物理）"),
 ("RVTYPE_SH", "多關節中採 RV 底座（中大負載）占比", "%", 0.50, 0.65, 0.80, "Assumed", "P", "—", "未取得公開統計；小負載 6 軸全用諧波"),
 ("N_RV_ART", "中大負載 6 軸：每台 RV 顆數", "顆/台", 3, 3.5, 4, "Interested-party", "P", S["fxg_red"], "底座、大臂、小臂關節；產業報告 3–4 顆"),
 ("N_HD_ART_RV", "中大負載 6 軸：每台諧波顆數", "顆/台", 2, 3, 4, "Interested-party", "P", S["fxg_red"], "腕部；產業報告 2–4 顆"),
 ("N_HD_SMALL", "小負載 6 軸：每台諧波顆數", "顆/台", 5, 6, 6, "Analogy", "P", S["fxg_red"], "類比協作（6 關節全諧波）；部分腕軸改用皮帶或行星"),
 ("N_HD_SCARA", "SCARA：每台諧波顆數", "顆/台", 2, 2, 3, "Assumed", "P", "—", "J1、J2 用諧波，J3/J4 多用滾珠螺桿花鍵（上銀產品）；未取得公開統計"),
 ("N_HD_COBOT", "協作：每台諧波顆數", "顆/台", 6, 6.5, 7, "Interested-party", "P", S["gytx"] + "；" + S["fxg_red"], "6–7 軸全諧波；觀研天下 6.5"),
 ("AFTER", "替換、維修與通路庫存係數", "倍", 1.00, 1.05, 1.15, "Assumed", "A", "—", "裝機口徑 → 減速機出貨口徑；不含半導體設備等非機器人用途"),
 ("sec", "D. 減速機單價（全球混合，2025 年價）"),
 ("ASP_RV", "RV 減速機全球混合均價 2025", "USD/顆", 500, 750, 1000, "Assumed", "I", S["dgzq"], "錨：中國 RV 均價約 4,530 元（2022，智研）≈ 660 美元；Nabtesco 價格高於陸製"),
 ("ASP_HD", "諧波減速機全球混合均價 2025", "USD/顆", 200, 300, 450, "Assumed", "I", S["dgzq"] + "；" + S["fxg_cobot"], "錨：中國諧波均價約 2,048 元（2022）≈ 300 美元；陸製比 HDS 低 40–60%"),
 ("ASP_DECL", "減速機均價年變動", "%", -0.05, -0.02, 0.0, "Assumed", "I", S["fxg_cobot"], "國產替代與價格戰壓價；HDS 高階持穩"),
 ("sec", "E. 單台回本（中國一般工業）"),
 ("BODY_6AX", "國產 6 軸本體價（中負載）", "RMB/台", 20000, 100000, 150000, "Interested-party", "I", S["ggii"] + "；" + S["fxg_ir"], "低＝GGII 6kg 小負載略高於 2 萬；基準與高＝產業報告 20kg 國產 10–15 萬"),
 ("BODY_6AX_F", "外資 6 軸本體價（中負載）", "RMB/台", 200000, 275000, 350000, "Interested-party", "I", S["fxg_ir"], "產業報告 20kg 外資 20–35 萬"),
 ("BODY_COBOT", "國產協作本體價", "RMB/台", 20000, 80000, 150000, "Interested-party", "I", S["fxg_cobot"] + "；" + S["ggii"], "5–15 萬元帶占八成以上量；低端已見 2 萬以下"),
 ("INTEG", "6 軸：整合與周邊 ÷ 本體價", "倍", 1.0, 1.5, 3.0, "Assumed", "A", S["frame"], "框架 1–3 倍；未取得公開統計"),
 ("INTEG_C", "協作：整合與周邊 ÷ 本體價", "倍", 0.3, 0.6, 1.0, "Assumed", "A", "—", "少圍欄、部署較簡單；未取得公開統計"),
 ("WAGE", "製造業每人年工資（2025）", "RMB/年", 76055, 76055, 113594, "Verified", "I", S["nbs_wage"], "基準與低＝私營單位（中小企業一般工業）；高＝非私營單位；口徑選擇為判斷"),
 ("LOAD", "雇主社保與福利附加倍數", "倍", 1.2, 1.3, 1.4, "Assumed", "I", "—", "五險一金雇主負擔與福利"),
 ("FTE", "6 軸：每班替代人數", "人/班", 0.5, 1.0, 1.5, "Assumed", "P", "—", ""),
 ("FTE_C", "協作：每班替代人數", "人/班", 0.3, 0.6, 1.0, "Assumed", "P", "—", "協作多為人機共線，替代比例較低"),
 ("SHIFTS", "每日班次", "班", 1, 2, 3, "Assumed", "P", "—", ""),
 ("MAINT", "年維護費 ÷ 系統投資", "%", 0.02, 0.04, 0.08, "Assumed", "A", "—", ""),
 ("KW", "6 軸平均耗電", "kW", 1.0, 1.5, 3.0, "Assumed", "P", "—", ""),
 ("KW_C", "協作平均耗電", "kW", 0.2, 0.4, 0.8, "Assumed", "P", "—", ""),
 ("HRS", "每班年運轉小時", "小時/年", 1800, 2000, 2400, "Assumed", "P", "—", ""),
 ("ELEC", "工業電價", "RMB/kWh", 0.6, 0.7, 0.9, "Assumed", "I", "—", ""),
 ("PB_TH", "商業化門檻：一般工業可接受回本期", "年", 2.0, 2.5, 3.0, "Assumed", "I", S["frame"], "框架第 4 節門檻 1"),
 ("sec", "F. 本體成本結構（BOM）"),
 ("RED_SH_6AX", "6 軸：減速機占本體成本", "%", 0.25, 0.30, 0.35, "Analogy", "P", S["fxg_red"] + "；" + S["frame"], "產業報告約 35%（另處 30–35%）；框架 25–35%"),
 ("HD_SH_COBOT", "協作：諧波占本體成本", "%", 0.15, 0.175, 0.20, "Interested-party", "P", S["fxg_cobot"], ""),
 ("TQ_SH_COBOT", "協作：關節力矩感測占本體成本（有配置時）", "%", 0.08, 0.115, 0.15, "Interested-party", "P", S["fxg_cobot"], "每關節約 500–2,000 美元"),
 ("sec", "G. 對帳用參考值"),
 ("FX_CNY", "人民幣兌美元", "RMB/USD", None, 6.83, None, "Derived", "A", S["hds_rev"], "HDS 財測假設 157 日圓/美元 ÷ 23 日圓/人民幣"),
 ("FX_JPY", "日圓兌美元", "JPY/USD", None, 157, None, "Interested-party", "A", S["hds_rev"], "HDS 財測匯率假設"),
 ("NAB_CMP_Q", "Nabtesco Component Solutions 外部營收 2026 年 1–3 月", "JPY m", None, 22259, None, "Verified", "A", S["nab_cmp"], "二手彙整（envalith），原始決算資料未讀取；分部含 RV 以外產品"),
 ("NAB_RV_SH", "Nabtesco 全球 RV 份額（2023）", "%", 0.35, 0.6141, 0.65, "Interested-party", "I", S["dgzq"] + "；" + S["fxg_red"], "華經 61.41%；另一報告稱 >35%（消費量口徑）"),
 ("HDS_REV", "HDS 營收 FY3/2026 實績", "JPY m", None, 59557, None, "Verified", "A", S["hds_rev"], "含半導體設備、機電一體產品"),
 ("HDS_SH", "HDS 全球諧波份額（金額）", "%", 0.40, 0.50, 0.60, "Interested-party", "I", S["fxg_red"], "產業報告 >50%"),
 ("NBS_2024", "中國工業機器人產量 2024", "台", None, 556400, None, "Verified", "A", S["fxg_ir"], "二手轉引國家統計局，原始公報未讀取"),
 ("NBS_G2025", "中國工業機器人產量年增 2025", "%", None, 0.28, None, "Verified", "A", S["nbs_2025"], "二手轉引國家統計局"),
]

# 追蹤：JARA 季度（會員口徑，百萬日圓、台）
JARA = [  # (期間, 受注額, 受注YoY, 生產額, 生產YoY, 受注台, 生產台, 出處鍵)
 ("2025Q4", 271373, 0.291, 228607, 0.260, 54740, 51797, "jara_q4"),
 ("2026Q1", 294799, 0.410, 244195, 0.229, 60412, 52983, "jara_q1"),
 ("2026Q2", 312663, 0.400, 252008, 0.243, 57695, 51437, "jara_q2"),
 ("2026Q3", 325419, 0.466, 261548, 0.262, 67003, 51414, "jara_q3"),
]

wb = openpyxl.Workbook()
wsD = wb.active; wsD.title = "說明"
wsI = wb.create_sheet("輸入")
wsC = wb.create_sheet("計算")
wsS = wb.create_sheet("情境")
wsT = wb.create_sheet("敏感度")
wsO = wb.create_sheet("下游取數")
wsK = wb.create_sheet("追蹤指標")


def hdr(ws, r, vals, fill=HDR):
    for i, v in enumerate(vals):
        c = ws.cell(r, 1 + i, v); c.font = BOLD; c.fill = fill; c.alignment = WRAP


# ================= 輸入 =================
wsI["A1"] = "輸入（藍字＝輸入；黃底＝判斷值，可改；計算頁一律取「基準」欄；情境頁的成長、中國占比、協作占比、減速機均價依情境取低／高欄；敏感度逐一取低／高）"; wsI["A1"].font = TITLE
wsI["A2"] = "標記：[Verified] 可公開驗證｜[Interested-party] 利害關係方｜[Analogy] 類比｜[Assumed] 假設｜[Derived] 推得。類別：[P] 物理｜[A] 會計恆等式｜[I] 誘因資訊"
hdr(wsI, 4, ["ID", "項目", "單位", "低", "基準", "高", "標記", "類別", "出處", "說明"])
IROW = {}
r = 5
for row in INP:
    if row[0] == "sec":
        c = wsI.cell(r, 1, row[1]); c.font = BOLD; c.fill = SEC
        r += 1; continue
    k, name, unit, lo, base, hi, tag, cat, src, note = row
    lo = base if lo is None else lo
    hi = base if hi is None else hi
    vals = [k, name, unit, lo, base, hi, f"[{tag}]", f"[{cat}]", src, note]
    for i, v in enumerate(vals):
        c = wsI.cell(r, 1 + i, v); c.alignment = WRAP
    judg = tag in ("Assumed", "Analogy", "Interested-party") or row[3] is not None
    for col in (4, 5, 6):
        c = wsI.cell(r, col); c.font = BLUE
        if judg: c.fill = JUDG
        if unit == "%": c.number_format = "0.0%"
        elif isinstance(base, float) and base < 100: c.number_format = "0.00"
        else: c.number_format = "#,##0"
    IROW[k] = r
    r += 1
for col, w in zip("ABCDEFGHIJ", [16, 36, 10, 11, 11, 11, 18, 7, 70, 50]):
    wsI.column_dimensions[col].width = w
wsI.freeze_panes = "D5"

VAL = {k: {"D": None, "E": None, "F": None} for k in IROW}
for row in INP:
    if row[0] == "sec": continue
    k = row[0]
    VAL[k] = {"D": row[3] if row[3] is not None else row[4], "E": row[4], "F": row[5] if row[5] is not None else row[4]}


def I(k, col="E"):
    return f"'輸入'!${col}${IROW[k]}"


# ================= 計算 =================
ws = wsC
ws["A1"] = "計算（框架第 3 節等式；全部引用「輸入」基準欄）"; ws["A1"].font = TITLE
CROW = {}


def crow(r, key, label, formula, unit="", fmt="#,##0", note=""):
    ws.cell(r, 1, key); ws.cell(r, 2, label)
    c = ws.cell(r, 3, formula); c.number_format = fmt
    ws.cell(r, 4, unit); ws.cell(r, 5, note).alignment = WRAP
    CROW[key] = r


# A. 單台回本：三欄（國產 6 軸、外資 6 軸、國產協作）
r = 3
ws.cell(r, 1, "A. 單台回本（中國一般工業）：系統投資＝本體×(1＋整合倍數)；年淨節省＝人工×附加×替代人數×班次−電力−維護；回本期＝系統投資÷年淨節省").font = BOLD
ws.cell(r, 1).fill = SEC
r += 1
hdr(ws, r, ["ID", "項目", "國產 6 軸", "外資 6 軸", "國產協作", "單位", "說明"])
PB = {}
cols = {"C": ("BODY_6AX", "INTEG", "FTE", "KW"), "D": ("BODY_6AX_F", "INTEG", "FTE", "KW"), "E": ("BODY_COBOT", "INTEG_C", "FTE_C", "KW_C")}
lines = [
 ("PB_BODY", "本體價", lambda b, m, f, k: f"={I(b)}", "RMB", "#,##0"),
 ("PB_MULT", "整合與周邊倍數", lambda b, m, f, k: f"={I(m)}", "倍", "0.00"),
 ("PB_SYS", "系統投資", None, "RMB", "#,##0"),
 ("PB_LAB1", "每人年人工成本（含附加）", lambda b, m, f, k: f"={I('WAGE')}*{I('LOAD')}", "RMB/年", "#,##0"),
 ("PB_FTE", "每班替代人數", lambda b, m, f, k: f"={I(f)}", "人", "0.00"),
 ("PB_SH", "班次", lambda b, m, f, k: f"={I('SHIFTS')}", "班", "0"),
 ("PB_LAB", "年人工節省", None, "RMB/年", "#,##0"),
 ("PB_PWR", "年電費", None, "RMB/年", "#,##0"),
 ("PB_MNT", "年維護", None, "RMB/年", "#,##0"),
 ("PB_NET", "年淨節省", None, "RMB/年", "#,##0"),
 ("PB_YR", "回本期", None, "年", "0.00"),
 ("PB_TH", "門檻", lambda b, m, f, k: f"={I('PB_TH')}", "年", "0.0"),
 ("PB_OK", "過門檻？", None, "", "@"),
]
r += 1
rr = {}
for key, label, fn, unit, fmt in lines:
    rr[key] = r; r += 1
for key, label, fn, unit, fmt in lines:
    rw = rr[key]
    ws.cell(rw, 1, key); ws.cell(rw, 2, label); ws.cell(rw, 6, unit)
    for colL, (b, m, f, k) in cols.items():
        if fn is not None:
            v = fn(b, m, f, k)
        elif key == "PB_SYS":
            v = f"={colL}{rr['PB_BODY']}*(1+{colL}{rr['PB_MULT']})"
        elif key == "PB_LAB":
            v = f"={colL}{rr['PB_LAB1']}*{colL}{rr['PB_FTE']}*{colL}{rr['PB_SH']}"
        elif key == "PB_PWR":
            v = f"={I(k)}*{I('HRS')}*{colL}{rr['PB_SH']}*{I('ELEC')}"
        elif key == "PB_MNT":
            v = f"={colL}{rr['PB_SYS']}*{I('MAINT')}"
        elif key == "PB_NET":
            v = f"={colL}{rr['PB_LAB']}-{colL}{rr['PB_PWR']}-{colL}{rr['PB_MNT']}"
        elif key == "PB_YR":
            v = f"=IF({colL}{rr['PB_NET']}>0,{colL}{rr['PB_SYS']}/{colL}{rr['PB_NET']},999)"
        elif key == "PB_OK":
            v = f'=IF({colL}{rr["PB_YR"]}<={colL}{rr["PB_TH"]},"是","否")'
        c = ws.cell(rw, {"C": 3, "D": 4, "E": 5}[colL], v); c.number_format = fmt
ws.cell(rr["PB_YR"], 7, "回本期 999＝年淨節省 ≤ 0")
for key in rr: CROW[key] = rr[key]
PB_ROW = rr

# B. BOM
r += 1
ws.cell(r, 1, "B. 本體成本結構：6 軸減速機價值 由上而下（本體×占比）對 由下而上（顆數×單價）").font = BOLD; ws.cell(r, 1).fill = SEC
r += 1; hdr(ws, r, ["ID", "項目", "值", "單位", "說明"]); r += 1
crow(r, "BOM_TD", "6 軸減速機價值：由上而下（國產本體×占比）", f"={I('BODY_6AX')}*{I('RED_SH_6AX')}", "RMB/台"); r += 1
crow(r, "BOM_BU", "6 軸減速機價值：由下而上（中大負載，全球混合單價）", f"=({I('N_RV_ART')}*{I('ASP_RV')}+{I('N_HD_ART_RV')}*{I('ASP_HD')})*{I('FX_CNY')}", "RMB/台", note="用全球混合單價（含日系），高於國產 6 軸實際採購價屬預期"); r += 1
crow(r, "BOM_RATIO", "由下而上 ÷ 由上而下", f"=C{CROW['BOM_BU']}/C{CROW['BOM_TD']}", "倍", "0.00", "≫1 表示國產 6 軸用的是低價陸製減速機，或 10 萬元本體價偏低"); r += 1
crow(r, "BOM_RED_F", "外資 6 軸：減速機占本體（由下而上÷外資本體）", f"=C{CROW['BOM_BU']}/{I('BODY_6AX_F')}", "%", "0.0%"); r += 1
crow(r, "BOM_HD_C", "協作：諧波價值（本體×占比）", f"={I('BODY_COBOT')}*{I('HD_SH_COBOT')}", "RMB/台"); r += 1
crow(r, "BOM_HD_C_U", "協作：隱含每顆諧波單價", f"=C{CROW['BOM_HD_C']}/{I('N_HD_COBOT')}", "RMB/顆", note="對照國產諧波均價約 2,048 元（2022）"); r += 1
crow(r, "BOM_TQ_C", "協作：力矩感測價值（有配置時）", f"={I('BODY_COBOT')}*{I('TQ_SH_COBOT')}", "RMB/台"); r += 1

# C. 每台用量（下游）
r += 1
ws.cell(r, 1, "C. 每台減速機用量（混合口徑，下游取數）").font = BOLD; ws.cell(r, 1).fill = SEC
r += 1; hdr(ws, r, ["ID", "項目", "值", "單位", "說明"]); r += 1
crow(r, "U_RV_TRAD", "每台非協作機器人平均 RV 顆數", f"={I('MIX_ART')}*{I('RVTYPE_SH')}*{I('N_RV_ART')}", "顆/台", "0.00", "＝多關節占比×RV 型占比×每台 RV"); r += 1
crow(r, "U_HD_TRAD", "每台非協作機器人平均諧波顆數", f"={I('MIX_ART')}*({I('RVTYPE_SH')}*{I('N_HD_ART_RV')}+(1-{I('RVTYPE_SH')})*{I('N_HD_SMALL')})+{I('MIX_SCARA')}*{I('N_HD_SCARA')}", "顆/台", "0.00"); r += 1
crow(r, "U_HD_COBOT", "每台協作諧波顆數", f"={I('N_HD_COBOT')}", "顆/台", "0.00"); r += 1
crow(r, "U_RV_ALL25", "2025 每台（全部裝機）平均 RV", f"=C{CROW['U_RV_TRAD']}*(1-{I('COBOT_INST_2025')}/{I('G_INST_2025')})", "顆/台", "0.00"); r += 1
crow(r, "U_HD_ALL25", "2025 每台（全部裝機）平均諧波", f"=C{CROW['U_HD_TRAD']}*(1-{I('COBOT_INST_2025')}/{I('G_INST_2025')})+{I('N_HD_COBOT')}*{I('COBOT_INST_2025')}/{I('G_INST_2025')}", "顆/台", "0.00"); r += 1
crow(r, "MIX_OTHER", "非協作中其他（直角座標、Delta 等）占比", f"=1-{I('MIX_ART')}-{I('MIX_SCARA')}", "%", "0.0%", "須 ≥ 0"); r += 1

# D. 產業層對帳（2025）——情境頁算完後引用
r += 1
ws.cell(r, 1, "D. 產業層對帳（2025，基準情境）").font = BOLD; ws.cell(r, 1).fill = SEC
r += 1; hdr(ws, r, ["ID", "項目", "值", "單位", "說明"]); r += 1
D_START = r
r += 12  # 預留，情境頁建好後填
# E. 循環層
E_START = r + 1

# ================= 情境 =================
ws = wsS
ws["A1"] = "情境（2025＝IFR 實績；2026–2032 低／基準／高）。成長、中國占比、協作占比、減速機單價依情境取值；機型結構與每台顆數三情境都用基準（其不確定性見敏感度）"; ws["A1"].font = BOLD
ROWS = [
 ("g", "全球裝機年增", "%", "0.0%"),
 ("inst", "全球裝機（含協作）", "台", "#,##0"),
 ("cnsh", "中國占比", "%", "0.0%"),
 ("cn", "中國裝機", "台", "#,##0"),
 ("cobsh", "協作占比", "%", "0.0%"),
 ("cob", "協作裝機", "台", "#,##0"),
 ("trad", "非協作裝機", "台", "#,##0"),
 ("art", "  多關節", "台", "#,##0"),
 ("rvt", "    中大負載（RV 型）", "台", "#,##0"),
 ("sml", "    小負載（全諧波）", "台", "#,##0"),
 ("scara", "  SCARA", "台", "#,##0"),
 ("oth", "  其他（直角座標、Delta 等）", "台", "#,##0"),
 ("rvu", "全球 RV 需求", "顆", "#,##0"),
 ("hdu", "全球諧波需求", "顆", "#,##0"),
 ("hdcob", "  其中協作拉動", "顆", "#,##0"),
 ("asprv", "RV 均價", "USD/顆", "#,##0"),
 ("asphd", "諧波均價", "USD/顆", "#,##0"),
 ("mrv", "RV 市場", "USD m", "#,##0"),
 ("mhd", "諧波市場", "USD m", "#,##0"),
 ("mred", "機器人減速機市場合計（RV＋諧波）", "USD m", "#,##0"),
 ("cnrv", "中國 RV 需求（假設與全球同結構）", "顆", "#,##0"),
 ("cnhd", "中國諧波需求（假設與全球同結構）", "顆", "#,##0"),
 ("mv", "全球本體市場值", "USD bn", "#,##0.0"),
 ("redmv", "減速機市場 ÷ 本體市場值", "%", "0.0%"),
]
SROW = {}  # (scen, key) -> row
r = 3
for sname, scol in SCN:
    c = ws.cell(r, 1, f"情境：{sname}"); c.font = BOLD; c.fill = SEC
    r += 1
    hdr(ws, r, ["ID", "項目", "單位"] + [f"{y}{'A' if y == 2025 else 'E'}" for y in YEARS])
    r += 1
    for key, label, unit, fmt in ROWS:
        SROW[(sname, key)] = r
        ws.cell(r, 1, f"{key}"); ws.cell(r, 2, label); ws.cell(r, 3, unit)
        r += 1
    r += 1

for sname, scol in SCN:
    R = lambda k: SROW[(sname, k)]
    for y in YEARS:
        L = YCOL[y]
        P = YCOL.get(y - 1)
        t = y - 2025
        f = {}
        if y == 2025:
            f["g"] = None
            f["inst"] = f"={I('G_INST_2025')}"
            f["cnsh"] = f"={I('CN_INST_2025')}/{I('G_INST_2025')}"
            f["cobsh"] = f"={I('COBOT_INST_2025')}/{I('G_INST_2025')}"
        else:
            gk = "G26" if y == 2026 else ("G2729" if y <= 2029 else "G3032")
            f["g"] = f"={I(gk, scol)}"
            f["inst"] = f"={P}{R('inst')}*(1+{L}{R('g')})"
            f["cnsh"] = f"={I('CN_SH', scol)}"
            f["cobsh"] = f"=$D{R('cobsh')}+({I('COBOT_SH_2032', scol)}-$D{R('cobsh')})*{t}/7"
        f["cn"] = f"={L}{R('inst')}*{L}{R('cnsh')}"
        f["cob"] = f"={L}{R('inst')}*{L}{R('cobsh')}"
        f["trad"] = f"={L}{R('inst')}-{L}{R('cob')}"
        f["art"] = f"={L}{R('trad')}*{I('MIX_ART')}"
        f["rvt"] = f"={L}{R('art')}*{I('RVTYPE_SH')}"
        f["sml"] = f"={L}{R('art')}-{L}{R('rvt')}"
        f["scara"] = f"={L}{R('trad')}*{I('MIX_SCARA')}"
        f["oth"] = f"={L}{R('trad')}-{L}{R('art')}-{L}{R('scara')}"
        f["rvu"] = f"={L}{R('rvt')}*{I('N_RV_ART')}*{I('AFTER')}"
        f["hdu"] = f"=({L}{R('rvt')}*{I('N_HD_ART_RV')}+{L}{R('sml')}*{I('N_HD_SMALL')}+{L}{R('scara')}*{I('N_HD_SCARA')}+{L}{R('cob')}*{I('N_HD_COBOT')})*{I('AFTER')}"
        f["hdcob"] = f"={L}{R('cob')}*{I('N_HD_COBOT')}*{I('AFTER')}"
        f["asprv"] = f"={I('ASP_RV', scol)}*(1+{I('ASP_DECL')})^{t}"
        f["asphd"] = f"={I('ASP_HD', scol)}*(1+{I('ASP_DECL')})^{t}"
        f["mrv"] = f"={L}{R('rvu')}*{L}{R('asprv')}/1000000"
        f["mhd"] = f"={L}{R('hdu')}*{L}{R('asphd')}/1000000"
        f["mred"] = f"={L}{R('mrv')}+{L}{R('mhd')}"
        f["cnrv"] = f"={L}{R('rvu')}*{L}{R('cnsh')}"
        f["cnhd"] = f"={L}{R('hdu')}*{L}{R('cnsh')}"
        f["mv"] = f"={I('G_MV_2025')}*{L}{R('inst')}/{I('G_INST_2025')}*(1+{I('ROBOT_ASP_G')})^{t}"
        f["redmv"] = f"={L}{R('mred')}/({L}{R('mv')}*1000)"
        for key, label, unit, fmt in ROWS:
            v = f.get(key)
            if v is None: continue
            c = ws.cell(R(key), 4 + YEARS.index(y), v); c.number_format = fmt
    # CAGR 欄
    ws.cell(SROW[(sname, "inst")] - 1 - ROWS.index(("inst", "全球裝機（含協作）", "台", "#,##0")), 12, "CAGR 25–32").font = BOLD
    for key in ("inst", "cn", "cob", "rvu", "hdu", "mred", "mv"):
        rw = SROW[(sname, key)]
        c = ws.cell(rw, 12, f"=(K{rw}/D{rw})^(1/7)-1"); c.number_format = "0.0%"
ws.column_dimensions["A"].width = 8; ws.column_dimensions["B"].width = 34
for y in YEARS: ws.column_dimensions[YCOL[y]].width = 12
ws.freeze_panes = "D3"

# 回填 計算 D：對帳
ws = wsC
B = lambda k, y=2025: f"'情境'!{YCOL[y]}{SROW[('基準', k)]}"
r = D_START
crow(r, "RC_RV_M", "模型 2025 RV 市場（基準）", f"={B('mrv')}", "USD m"); r += 1
crow(r, "RC_NAB_IMP", "Nabtesco CMP 年化營收 ÷ RV 份額（上限參考）", f"={I('NAB_CMP_Q')}*4/{I('FX_JPY')}/{I('NAB_RV_SH')}", "USD m", note="CMP 含 RV 以外產品，故為上限；口徑差異未調整"); r += 1
crow(r, "RC_RV_RATIO", "模型 ÷ Nabtesco 推得上限", f"=C{CROW['RC_RV_M']}/C{CROW['RC_NAB_IMP']}", "倍", "0.00", "應 < 1；落在 0.5–1.0 視為合理"); r += 1
crow(r, "RC_HD_M", "模型 2025 諧波市場（基準，僅機器人）", f"={B('mhd')}", "USD m"); r += 1
crow(r, "RC_HDS_IMP", "HDS FY3/2026 營收 ÷ 諧波份額（含非機器人，上限參考）", f"={I('HDS_REV')}/{I('FX_JPY')}/{I('HDS_SH')}", "USD m", note="HDS 營收含半導體設備等非機器人用途與機電產品"); r += 1
crow(r, "RC_HD_RATIO", "模型 ÷ HDS 推得上限", f"=C{CROW['RC_HD_M']}/C{CROW['RC_HDS_IMP']}", "倍", "0.00", "應 < 1"); r += 1
crow(r, "RC_AVGV", "IFR 隱含每台本體均價 2025", f"={I('G_MV_2025')}*1000000000/{I('G_INST_2025')}", "USD/台"); r += 1
crow(r, "RC_RED_PER", "每台（全部裝機）減速機價值 2025", f"={B('mred')}*1000000/{I('G_INST_2025')}", "USD/台"); r += 1
crow(r, "RC_RED_SH", "減速機價值 ÷ IFR 本體均價", f"=C{CROW['RC_RED_PER']}/C{CROW['RC_AVGV']}", "%", "0.0%", "遠低於 BOM 占比 25–35% 屬預期：IFR 均價含日歐高價機與 SCARA、直角座標；BOM 占比是國產 6 軸成本口徑"); r += 1
crow(r, "RC_NBS25", "中國產量 2025（推得）", f"={I('NBS_2024')}*(1+{I('NBS_G2025')})", "台"); r += 1
crow(r, "RC_NBS_IFR", "中國產量 ÷ IFR 中國裝機 2025", f"=C{CROW['RC_NBS25']}/{I('CN_INST_2025')}", "倍", "0.00", "口徑差：產量含出口、可能以套計含部件；只看增速，不看絕對值"); r += 1

# ================= 追蹤指標 =================
ws = wsK
ws["A1"] = "追蹤指標（框架第 5 節）：傳導順序 減速機與整機訂單（領先）→ 整機生產出貨 → 中國產量與出口 → IFR 裝機（落後）→ 零組件營收"; ws["A1"].font = BOLD
r = 3
ws.cell(r, 1, "A. JARA 季度（會員口徑，百萬日圓）").font = BOLD; ws.cell(r, 1).fill = SEC
r += 1
hdr(ws, r, ["期間", "受注額", "受注 YoY", "生產額", "生產 YoY", "受注台數", "生產台數", "訂單÷生產（金額）", "訂單÷生產（台數）", "受注額 QoQ", "出處"])
r += 1
JROW = {}
for i, (p, o, oy, pr, py, ou, pu, sk) in enumerate(JARA):
    vals = [p, o, oy, pr, py, ou, pu]
    for j, v in enumerate(vals):
        c = ws.cell(r, 1 + j, v)
        if j > 0: c.font = BLUE
        c.number_format = "0.0%" if j in (2, 4) else "#,##0"
    ws.cell(r, 8, f"=B{r}/D{r}").number_format = "0.00"
    ws.cell(r, 9, f"=F{r}/G{r}").number_format = "0.00"
    if i > 0:
        ws.cell(r, 10, f"=B{r}/B{r-1}-1").number_format = "0.0%"
    ws.cell(r, 11, S[sk])
    JROW[p] = r
    r += 1
ws.cell(r, 1, "2025 全年");
for j, v in enumerate([925845, 0.278, 837301, 0.210, 197738, 189500]):
    c = ws.cell(r, 2 + j, v); c.font = BLUE; c.number_format = "0.0%" if j in (1, 3) else "#,##0"
ws.cell(r, 8, f"=B{r}/D{r}").number_format = "0.00"
ws.cell(r, 9, f"=F{r}/G{r}").number_format = "0.00"
ws.cell(r, 11, S["jara_q4"])
JROW["2025"] = r; r += 1
ws.cell(r, 1, "2026 前三季合計")
ws.cell(r, 2, f"=SUM(B{JROW['2026Q1']}:B{JROW['2026Q3']})").number_format = "#,##0"
ws.cell(r, 4, f"=SUM(D{JROW['2026Q1']}:D{JROW['2026Q3']})").number_format = "#,##0"
ws.cell(r, 8, f"=B{r}/D{r}").number_format = "0.00"
ws.cell(r, 11, "[Derived]")
JROW["2026YTD"] = r; r += 1
ws.cell(r, 1, "近四季受注合計")
ws.cell(r, 2, f"=SUM(B{JROW['2025Q4']}:B{JROW['2026Q3']})").number_format = "#,##0"
ws.cell(r, 3, f"=B{r}/B{JROW['2025']}-1").number_format = "0.0%"
ws.cell(r, 11, "[Derived]；YoY 欄＝對 2025 全年")
JROW["LTM"] = r; r += 1
ws.cell(r, 1, "JARA 2026 全年受注預測（含非會員）")
c = ws.cell(r, 2, 1220000); c.font = BLUE; c.fill = JUDG; c.number_format = "#,##0"
c = ws.cell(r, 3, 0.167); c.font = BLUE; c.number_format = "0.0%"
ws.cell(r, 11, "[Interested-party] 產業協會預測；" + S["jara_fc"] + "；2025 含非會員實績 1.0456 兆")
JROW["FC"] = r; r += 2

ws.cell(r, 1, "B. 最新讀數（2026-11-16 查）").font = BOLD; ws.cell(r, 1).fill = SEC
r += 1
hdr(ws, r, ["指標", "最新期", "讀數", "單位", "YoY", "公布／查詢日", "標記", "出處", "備註"])
r += 1
LATEST = [
 ("JARA 受注額", "2026 年 7–9 月", f"=B{JROW['2026Q3']}", "百萬日圓", f"=C{JROW['2026Q3']}", "2026-10", "[Verified]", S["jara_q3"], "受注台數 67,003（+36.1%）"),
 ("JARA 生產額", "2026 年 7–9 月", f"=D{JROW['2026Q3']}", "百萬日圓", f"=E{JROW['2026Q3']}", "2026-10", "[Verified]", S["jara_q3"], "出口 216,478（+23.5%）；對中國 88,141（+29.2%）；中國以外亞洲 +41.3%"),
 ("JARA 訂單÷生產", "2026 年 7–9 月", f"=H{JROW['2026Q3']}", "倍", None, "2026-10", "[Derived]", "由上兩列", "連續四季 > 1.1"),
 ("中國工業機器人產量", "2026 年 9 月", 96752, "台", 0.189, "2026-10-19", "[Verified]", S["nbs_sep"], "1–9 月 794,781 台（+29.1%）"),
 ("中國工業機器人產量（累計）", "2026 年 1–9 月", 794781, "台", 0.291, "2026-10-19", "[Verified]", S["nbs_sep"], "2026 年 10 月數據本輪未取得"),
 ("中國機器人減速器產量", "2025 全年", "未取得（台數）", "", 0.639, "2026-01-19", "[Verified]", S["nbs_2025"], "二手轉引國家統計局"),
 ("哈默納科受注", "2026 年 4–6 月", 24100, "百萬日圓", 0.557, "2026-08", "[Verified]", S["hds_q1"], "受注殘 +45.6%；7–9 月受注數字未揭露（僅稱高水準）"),
 ("哈默納科營收", "2026 年 4–9 月", 36124, "百萬日圓", 0.298, "2026-11-11", "[Verified]", S["hds_h1"], "營業利益 4,205（前年 465）；全年財測上修至 78,000"),
 ("Nabtesco CMP 受注", "2026 年 1–3 月", 22796, "百萬日圓", "未取得", "—", "[Verified]", S["nab_cmp"], "二手彙整；精密減速機單獨數字與 7–9 月未取得"),
 ("Nabtesco 集團受注", "2026 年 1–6 月", 198500, "百萬日圓", 0.331, "2026-07", "[Verified]", S["nab_h1"], ""),
 ("Fanuc Robot 部門受注", "2026 年 7–9 月", 103400, "百萬日圓", 0.126, "2026-10", "[Verified]", S["fanuc"], "QoQ +5.6%；FA 部門 +130%（框架）"),
 ("安川 Robotics 受注", "2026 年 6–8 月", 77600, "百萬日圓", 0.33, "2026-10", "[Verified]", S["yaskawa"], "美洲 +60%、中國 +35%"),
 ("Teradyne Robotics（UR）營收", "2026 年 7–9 月", 96, "百萬美元", "未取得", "2026-10", "[Verified]", S["ter"], "部門轉盈"),
 ("北美機器人訂單（A3）", "2026 年 7–9 月", "未取得（台數）", "", 0.248, "2026-11", "[Verified]", S["a3"], "僅讀標題"),
 ("IFR 全球裝機", "2025", f"={I('G_INST_2025')}", "台", 0.11, "2026-09-24", "[Verified]", S["ifr_es"], "協作 6.6 萬（11%）；中國國產比 55%"),
]
KROW = {}
for row in LATEST:
    for j, v in enumerate(row):
        c = ws.cell(r, 1 + j, v)
        if j == 2 and isinstance(v, (int, float)): c.font = BLUE
        if j == 4 and isinstance(v, (int, float)): c.font = BLUE
        if j == 2: c.number_format = "#,##0.00" if row[3] == "倍" else "#,##0"
        if j == 4: c.number_format = "0.0%"
    KROW[row[0]] = r
    r += 1
r += 1
ws.cell(r, 1, "C. 追蹤表（頻率與來源）").font = BOLD; ws.cell(r, 1).fill = SEC
r += 1
hdr(ws, r, ["指標", "頻率", "資料來源", "為何重要"])
r += 1
TRACK = [
 ("哈默納科受注、受注殘、訂單÷營收", "季（5、8、11、2 月）", "公司決算短信", "最上游、最早反映機器人與半導體設備備貨；諧波也連到協作與人形"),
 ("Nabtesco 精密減速機訂單", "季", "公司決算資料（分部）", "RV 對應傳統重負載機器人，與哈默納科並看可分辨傳統與協作"),
 ("安川 Robotics 訂單（分地區）", "季；2 月制年度，最早公布", "安川決算說明", "主要整機廠中最早出數，中國、美洲分地區訂單"),
 ("Fanuc Robot 部門訂單與 FA 訂單", "季", "Fanuc 法說補充資料（reference）", "Robot 看整機，FA（CNC）看更廣的機械資本支出"),
 ("JARA 受注、生產、出口（分地區）", "季", "日本機器人工業會", "日系全體訂單與出口目的地；訂單÷生產是循環位置"),
 ("中國工業機器人產量", "月（約 15–19 日）", "國家統計局規模以上工業生產", "最高頻；口徑含出口、與 IFR 裝機不同，看增速不看絕對值"),
 ("台灣零組件月營收：上銀、和大、大銀微、達明", "月（10 日前）", "公開資訊觀測站", "亞太掛牌標的的最高頻讀數"),
 ("綠的諧波、匯川、埃斯頓營收與毛利率", "季", "公司財報", "國產替代與價格戰的利潤端"),
 ("北美機器人訂單", "季", "A3", "美國製造回流需求"),
 ("IFR 全球裝機、協作占比、中國國產比", "年（9 月）", "IFR World Robotics", "落後指標，用來校準本模型 2025 錨點與情境"),
 ("中國協作機器人出貨與均價", "季或半年", "GGII、MIR 睿工業（付費）", "協作價格曲線"),
 ("日本工具機受注", "月", "日本工作機械工業會", "與機器人同屬設備資本支出，交叉驗證循環"),
]
for row in TRACK:
    for j, v in enumerate(row): ws.cell(r, 1 + j, v).alignment = WRAP
    r += 1
for col, w in zip("ABCDEFGHIJK", [30, 16, 14, 12, 10, 12, 14, 60, 40, 10, 60]):
    ws.column_dimensions[col].width = w

# 計算 E：循環層
ws = wsC
r = E_START
ws.cell(r, 1, "E. 循環層（JARA，引用追蹤指標頁）").font = BOLD; ws.cell(r, 1).fill = SEC
r += 1; hdr(ws, r, ["ID", "項目", "值", "單位", "說明"]); r += 1
crow(r, "CY_BB", "JARA 最新季訂單÷生產（金額）", f"='追蹤指標'!H{JROW['2026Q3']}", "倍", "0.00", "> 1 表示訂單領先出貨，營收認列在後"); r += 1
crow(r, "CY_OY", "JARA 最新季受注 YoY", f"='追蹤指標'!C{JROW['2026Q3']}", "%", "0.0%"); r += 1
crow(r, "CY_LTM", "JARA 近四季受注 ÷ 2025 全年", f"='追蹤指標'!B{JROW['LTM']}/'追蹤指標'!B{JROW['2025']}", "倍", "0.00"); r += 1
crow(r, "CY_G26_GAP", "情境基準 2026 裝機年增 − JARA 前三季生產 YoY 平均", f"={I('G26')}-AVERAGE('追蹤指標'!E{JROW['2026Q1']}:E{JROW['2026Q3']})", "百分點", "0.0%", "負值＝基準（IFR）比日系生產端保守；日系生產額含漲價與匯率"); r += 1
for col, w in zip("ABCDEFG", [14, 52, 16, 14, 14, 10, 50]):
    ws.column_dimensions[col].width = w
ws.column_dimensions["E"].width = 16

# ================= 敏感度 =================
# 輸出 1：2030 全球減速機市場（基準情境）；輸出 2：國產 6 軸回本期


def mkt2030(v):
    t = 5
    I2030 = f"({v['G_INST_2025']}*(1+{v['G26']})*(1+{v['G2729']})^3*(1+{v['G3032']}))"
    c25 = f"({v['COBOT_INST_2025']}/{v['G_INST_2025']})"
    csh = f"({c25}+({v['COBOT_SH_2032']}-{c25})*{t}/7)"
    cob = f"({I2030}*{csh})"
    trad = f"({I2030}-{cob})"
    rvu = f"({trad}*{v['MIX_ART']}*{v['RVTYPE_SH']}*{v['N_RV_ART']}*{v['AFTER']})"
    hdu = (f"(({trad}*({v['MIX_ART']}*({v['RVTYPE_SH']}*{v['N_HD_ART_RV']}+(1-{v['RVTYPE_SH']})*{v['N_HD_SMALL']})"
           f"+{v['MIX_SCARA']}*{v['N_HD_SCARA']})+{cob}*{v['N_HD_COBOT']})*{v['AFTER']})")
    return f"({rvu}*{v['ASP_RV']}+{hdu}*{v['ASP_HD']})*(1+{v['ASP_DECL']})^{t}/1000000"


def payback(v):
    sys_ = f"({v['BODY_6AX']}*(1+{v['INTEG']}))"
    lab = f"({v['WAGE']}*{v['LOAD']}*{v['FTE']}*{v['SHIFTS']})"
    pwr = f"({v['KW']}*{v['HRS']}*{v['SHIFTS']}*{v['ELEC']})"
    return f"{sys_}/({lab}-{pwr}-{sys_}*{v['MAINT']})"


OUTS = [
 ("2030 全球機器人減速機市場（RV＋諧波，基準情境）", "USD m", mkt2030,
  ["G_INST_2025", "G26", "G2729", "G3032", "COBOT_INST_2025", "COBOT_SH_2032", "MIX_ART", "MIX_SCARA", "RVTYPE_SH",
   "N_RV_ART", "N_HD_ART_RV", "N_HD_SMALL", "N_HD_SCARA", "N_HD_COBOT", "AFTER", "ASP_RV", "ASP_HD", "ASP_DECL"],
  ("情境", B("mred", 2030)), "#,##0"),
 ("國產 6 軸回本期（中國一般工業）", "年", payback,
  ["BODY_6AX", "INTEG", "WAGE", "LOAD", "FTE", "SHIFTS", "KW", "HRS", "ELEC", "MAINT"],
  ("計算", f"'計算'!C{PB_ROW['PB_YR']}"), "0.00"),
]


def pyeval(expr):
    return eval(expr.replace("^", "**"))


ws = wsT
ws["A1"] = "敏感度：承重變數排序（一次只動一個變數到其低／高值，其餘取基準；依擺幅由大到小）"; ws["A1"].font = TITLE
ws["A2"] = "先看這頁再看基準結果：排名前面的變數決定結論是否成立。公式全部引用「輸入」頁，可追溯。"
r = 4
SENS_TOP = {}
for oi, (oname, ounit, fn, vars_, (chk_sheet, chk_ref), fmt) in enumerate(OUTS):
    ws.cell(r, 1, f"輸出 {oi+1}：{oname}（{ounit}）").font = BOLD; ws.cell(r, 1).fill = SEC
    r += 1
    base_ref = {k: I(k) for k in vars_}
    ws.cell(r, 1, "基準值"); c = ws.cell(r, 4, "=" + fn(base_ref)); c.number_format = fmt; base_r = r
    ws.cell(r, 5, "對帳："); c = ws.cell(r, 6, f"={chk_ref}"); c.number_format = fmt
    c = ws.cell(r, 7, f'=IF(ABS(D{r}-F{r})<1E-6*MAX(1,ABS(F{r})),"一致","不一致")')
    ws.cell(r, 8, f"對帳來源：{chk_sheet}頁")
    r += 1
    hdr(ws, r, ["排名", "變數 ID", "變數", "低值", "高值", "輸出@低", "輸出@高", "擺幅（絕對值）", "擺幅÷基準", "標記", "排名（公式核對）"])
    r += 1
    # Python 端估值以決定排列順序
    nb = {k: VAL[k]["E"] for k in vars_}
    swings = []
    for k in vars_:
        if VAL[k]["D"] == VAL[k]["F"]: continue
        lo = dict(nb); lo[k] = VAL[k]["D"]
        hi = dict(nb); hi[k] = VAL[k]["F"]
        swings.append((abs(pyeval(fn(hi)) - pyeval(fn(lo))), k))
    swings.sort(reverse=True)
    first = r
    last = r + len(swings) - 1
    SENS_TOP[oi] = []
    for rank, (_, k) in enumerate(swings, 1):
        lo = dict(base_ref); lo[k] = I(k, "D")
        hi = dict(base_ref); hi[k] = I(k, "F")
        ws.cell(r, 1, rank); ws.cell(r, 2, k); ws.cell(r, 3, f"='輸入'!B{IROW[k]}")
        c = ws.cell(r, 4, f"={I(k, 'D')}"); c.number_format = wsI.cell(IROW[k], 5).number_format
        c = ws.cell(r, 5, f"={I(k, 'F')}"); c.number_format = wsI.cell(IROW[k], 5).number_format
        c = ws.cell(r, 6, "=" + fn(lo)); c.number_format = fmt
        c = ws.cell(r, 7, "=" + fn(hi)); c.number_format = fmt
        c = ws.cell(r, 8, f"=ABS(G{r}-F{r})"); c.number_format = fmt
        c = ws.cell(r, 9, f"=H{r}/$D${base_r}"); c.number_format = "0%"
        ws.cell(r, 10, f"='輸入'!G{IROW[k]}")
        ws.cell(r, 11, f"=RANK(H{r},$H${first}:$H${last})")
        if rank <= 3: SENS_TOP[oi].append(r)
        r += 1
    SENS_TOP[(oi, "base")] = base_r
    r += 2
for col, w in zip("ABCDEFGHIJK", [8, 16, 40, 12, 12, 14, 14, 14, 10, 18, 10]):
    ws.column_dimensions[col].width = w

# ================= 下游取數 =================
ws = wsO
ws["A1"] = "下游取數（給個股模型讀：哈默納科、Nabtesco、Fanuc、安川、上銀等）。值全部是連結公式；以 ID 取數，不要以列號取數。"; ws["A1"].font = BOLD
ws["A2"] = "口徑：裝機＝IFR 新裝機台數（含協作）；減速機需求＝裝機×每台顆數×替換係數（僅機器人用途，不含半導體設備、人形）；市場＝需求×全球混合均價（USD m）"
hdr(ws, 4, ["ID", "名稱", "年", "情境", "值", "單位", "口徑", "來源儲存格", "標記"])
r = 5
SER = [
 ("IR.inst.global", "inst", "全球工業機器人裝機（含協作）", "台", "IFR 新裝機口徑；2025 實績，之後情境", "#,##0"),
 ("IR.inst.china", "cn", "中國工業機器人裝機（含協作）", "台", "IFR 中國裝機口徑（非國家統計局產量）", "#,##0"),
 ("IR.inst.cobot", "cob", "全球協作機器人裝機", "台", "IFR 協作口徑", "#,##0"),
 ("IR.share.cobot", "cobsh", "協作占全球裝機比", "%", "台數比", "0.0%"),
 ("IR.inst.trad", "trad", "全球非協作工業機器人裝機", "台", "裝機−協作", "#,##0"),
 ("IR.red.rv.units", "rvu", "全球機器人 RV 減速機需求", "顆", "僅工業機器人；含替換係數", "#,##0"),
 ("IR.red.hd.units", "hdu", "全球機器人諧波減速機需求", "顆", "僅工業機器人（含協作）；不含人形、半導體設備", "#,##0"),
 ("IR.red.hd.cobot_units", "hdcob", "其中協作拉動的諧波需求", "顆", "協作裝機×每台諧波×替換係數", "#,##0"),
 ("IR.red.rv.mkt", "mrv", "全球機器人 RV 減速機市場", "USD m", "需求×全球混合均價", "#,##0"),
 ("IR.red.hd.mkt", "mhd", "全球機器人諧波減速機市場", "USD m", "同上", "#,##0"),
 ("IR.red.total.mkt", "mred", "全球機器人減速機市場（RV＋諧波）", "USD m", "不含行星、擺線以外產品", "#,##0"),
 ("IR.red.rv.units.china", "cnrv", "中國 RV 減速機需求", "顆", "假設中國機型結構同全球", "#,##0"),
 ("IR.red.hd.units.china", "cnhd", "中國諧波減速機需求", "顆", "同上", "#,##0"),
 ("IR.mv.global", "mv", "全球工業機器人本體市場值", "USD bn", "IFR 本體價值口徑×單價變動", "#,##0.0"),
]
for sid, key, name, unit, kou, fmt in SER:
    for sname, scol in SCN:
        for y in YEARS:
            if y == 2025 and sname != "基準": continue
            ref = f"'情境'!{YCOL[y]}{SROW[(sname, key)]}"
            vals = [f"{sid}.{y}.{ {'低':'lo','基準':'base','高':'hi'}[sname] }".replace(" ", ""), name, y,
                    "實績" if y == 2025 else sname, f"={ref}", unit, kou, ref.replace("'", ""),
                    "[Verified]" if (y == 2025 and key in ("inst", "cn", "cob", "cobsh", "trad", "mv")) else "[Derived]"]
            for j, v in enumerate(vals): ws.cell(r, 1 + j, v)
            ws.cell(r, 5).number_format = fmt
            r += 1
# 單值
SINGLE = []
for k, name in [("N_RV_ART", "中大負載 6 軸每台 RV 顆數"), ("N_HD_ART_RV", "中大負載 6 軸每台諧波顆數"),
                ("N_HD_SMALL", "小負載 6 軸每台諧波顆數"), ("N_HD_SCARA", "SCARA 每台諧波顆數"),
                ("N_HD_COBOT", "協作每台諧波顆數"), ("RVTYPE_SH", "多關節中 RV 型占比")]:
    for sname, scol in SCN:
        SINGLE.append((f"IR.unit.{k}.{ {'低':'lo','基準':'base','高':'hi'}[sname] }".replace(" ", ""), name, "—", sname,
                       f"={I(k, scol)}", wsI.cell(IROW[k], 3).value, "每台物理用量", f"輸入!{scol}{IROW[k]}", wsI.cell(IROW[k], 7).value))
for key, name, unit in [("U_RV_TRAD", "每台非協作機器人平均 RV 顆數（混合）", "顆/台"), ("U_HD_TRAD", "每台非協作機器人平均諧波顆數（混合）", "顆/台"),
                        ("U_RV_ALL25", "每台工業機器人平均 RV 顆數（2025 全部裝機）", "顆/台"), ("U_HD_ALL25", "每台工業機器人平均諧波顆數（2025 全部裝機）", "顆/台")]:
    SINGLE.append((f"IR.unit.{key}", name, 2025 if "25" in key else "—", "基準", f"='計算'!C{CROW[key]}", unit, "機型結構加權", f"計算!C{CROW[key]}", "[Derived]"))
for colL, lab in [("C", "國產 6 軸"), ("D", "外資 6 軸"), ("E", "國產協作")]:
    SINGLE.append((f"IR.payback.{ {'C':'cn6ax','D':'for6ax','E':'cncobot'}[colL] }".replace(" ", ""), f"{lab}回本期（中國一般工業）", "—", "基準",
                   f"='計算'!{colL}{PB_ROW['PB_YR']}", "年", "系統投資÷年淨節省", f"計算!{colL}{PB_ROW['PB_YR']}", "[Derived]"))
SINGLE += [
 ("IR.cycle.jara.orders.latest", "JARA 最新季受注額（2026 年 7–9 月）", "2026Q3", "實績", f"='追蹤指標'!B{JROW['2026Q3']}", "JPY m", "會員口徑", f"追蹤指標!B{JROW['2026Q3']}", "[Verified]"),
 ("IR.cycle.jara.orders_yoy.latest", "JARA 最新季受注 YoY", "2026Q3", "實績", f"='追蹤指標'!C{JROW['2026Q3']}", "%", "會員口徑", f"追蹤指標!C{JROW['2026Q3']}", "[Verified]"),
 ("IR.cycle.jara.prod.latest", "JARA 最新季生產額", "2026Q3", "實績", f"='追蹤指標'!D{JROW['2026Q3']}", "JPY m", "會員口徑", f"追蹤指標!D{JROW['2026Q3']}", "[Verified]"),
 ("IR.cycle.jara.prod_yoy.latest", "JARA 最新季生產 YoY", "2026Q3", "實績", f"='追蹤指標'!E{JROW['2026Q3']}", "%", "會員口徑", f"追蹤指標!E{JROW['2026Q3']}", "[Verified]"),
 ("IR.cycle.jara.b2b.latest", "JARA 最新季訂單÷生產（金額）", "2026Q3", "實績", f"='追蹤指標'!H{JROW['2026Q3']}", "倍", "會員口徑", f"追蹤指標!H{JROW['2026Q3']}", "[Derived]"),
 ("IR.cycle.jara.orders.ltm", "JARA 近四季受注合計（2025Q4–2026Q3）", "LTM", "實績", f"='追蹤指標'!B{JROW['LTM']}", "JPY m", "會員口徑", f"追蹤指標!B{JROW['LTM']}", "[Derived]"),
 ("IR.cycle.jara.orders.2025", "JARA 2025 全年受注", 2025, "實績", f"='追蹤指標'!B{JROW['2025']}", "JPY m", "會員口徑", f"追蹤指標!B{JROW['2025']}", "[Verified]"),
 ("IR.cycle.jara.prod.2025", "JARA 2025 全年生產", 2025, "實績", f"='追蹤指標'!D{JROW['2025']}", "JPY m", "會員口徑", f"追蹤指標!D{JROW['2025']}", "[Verified]"),
 ("IR.cycle.jara.orders_fc.2026", "JARA 2026 全年受注預測", 2026, "協會預測", f"='追蹤指標'!B{JROW['FC']}", "JPY m", "含非會員（與季度會員口徑不同）", f"追蹤指標!B{JROW['FC']}", "[Interested-party]"),
 ("IR.cycle.cn_prod.ytd", "中國工業機器人產量 2026 年 1–9 月", "2026", "實績", f"='追蹤指標'!C{KROW['中國工業機器人產量（累計）']}", "台", "國家統計局產量口徑（含出口）", f"追蹤指標!C{KROW['中國工業機器人產量（累計）']}", "[Verified]"),
 ("IR.cycle.cn_prod.ytd_yoy", "中國工業機器人產量 2026 年 1–9 月 YoY", "2026", "實績", f"='追蹤指標'!E{KROW['中國工業機器人產量（累計）']}", "%", "同上", f"追蹤指標!E{KROW['中國工業機器人產量（累計）']}", "[Verified]"),
 ("IR.share.china_domestic.2025", "中國裝機國產品牌占比", 2025, "實績", f"={I('CN_DOM_2025')}", "%", "IFR", f"輸入!E{IROW['CN_DOM_2025']}", "[Verified]"),
]
for row in SINGLE:
    for j, v in enumerate(row): ws.cell(r, 1 + j, v)
    u = row[5]
    ws.cell(r, 5).number_format = "0.0%" if u == "%" else ("0.00" if u in ("顆/台", "年", "倍") else "#,##0")
    r += 1
for col, w in zip("ABCDEFGHI", [36, 40, 8, 10, 14, 10, 40, 22, 18]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "C5"
ws.auto_filter.ref = f"A4:I{r-1}"

# ================= 說明 =================
ws = wsD
ws["A1"] = f"線 2a 工業機器人（含協作）單位經濟模型 {VERSION}"; ws["A1"].font = TITLE
info = [
 ("產生", "2026-11-16；由 build_industrial_robot.py 產生；數字待未參與製作的代理獨立核對"),
 ("依據", S["frame"] + "（第 3 節等式、第 4 節門檻、第 5 節追蹤指標）"),
 ("核心問題", "(1) 單台：中國一般工業的系統投資與回本期，是否跨過約 2–3 年門檻；(2) 產業：裝機情境與每台減速機用量 → RV 與諧波需求與市場，供減速機與整機個股模型讀取；(3) 循環：JARA 訂單→生產的傳導位置"),
 ("頁面", "說明｜輸入（藍字輸入、黃底判斷）｜計算（框架等式）｜情境（2025A、2026–2032 低／基準／高）｜敏感度（承重變數排序）｜下游取數（以 ID 取數）｜追蹤指標（JARA 季度、最新讀數、頻率表）"),
 ("標記", "[Verified]／[Interested-party]／[Analogy]／[Assumed]／[Derived]；Analogy、Assumed 一律給低／基準／高"),
 ("不做", "人形（線 2b）、倉儲 AMR（線 3）、機器人智能與控制器軟體、晶片（AI半導體只讀）；減速機的半導體設備與人形需求不在本線市場內"),
]
r = 3
for k, v in info:
    ws.cell(r, 1, k).font = BOLD; c = ws.cell(r, 2, v); c.alignment = WRAP; r += 1
r += 1
ws.cell(r, 1, "一、先看敏感度（承重變數前三名）").font = BOLD; ws.cell(r, 1).fill = SEC; r += 1
hdr(ws, r, ["輸出", "排名", "變數", "輸出@低", "輸出@高", "擺幅÷基準"]); r += 1
for oi, (oname, ounit, *_ ) in enumerate(OUTS):
    for rank, sr in enumerate(SENS_TOP[oi], 1):
        ws.cell(r, 1, f"{oname}（{ounit}）")
        ws.cell(r, 2, rank)
        ws.cell(r, 3, f"='敏感度'!C{sr}")
        fmt = OUTS[oi][5]
        c = ws.cell(r, 4, f"='敏感度'!F{sr}"); c.number_format = fmt
        c = ws.cell(r, 5, f"='敏感度'!G{sr}"); c.number_format = fmt
        c = ws.cell(r, 6, f"='敏感度'!I{sr}"); c.number_format = "0%"
        r += 1
r += 1
ws.cell(r, 1, "二、基準結果").font = BOLD; ws.cell(r, 1).fill = SEC; r += 1
hdr(ws, r, ["項目", "低", "基準", "高", "單位", "來源"]); r += 1


def triple(label, key, y, unit, fmt):
    global r
    ws.cell(r, 1, label)
    for j, (sname, _) in enumerate(SCN):
        c = ws.cell(r, 2 + j, f"='情境'!{YCOL[y]}{SROW[(sname, key)]}"); c.number_format = fmt
    ws.cell(r, 5, unit); ws.cell(r, 6, "情境頁")
    r += 1


triple("全球裝機 2030", "inst", 2030, "台", "#,##0")
triple("全球裝機 2032", "inst", 2032, "台", "#,##0")
triple("中國裝機 2030", "cn", 2030, "台", "#,##0")
triple("協作裝機 2030", "cob", 2030, "台", "#,##0")
triple("全球 RV 需求 2030", "rvu", 2030, "顆", "#,##0")
triple("全球諧波需求 2030", "hdu", 2030, "顆", "#,##0")
triple("機器人減速機市場 2025", "mred", 2025, "USD m", "#,##0")
triple("機器人減速機市場 2030", "mred", 2030, "USD m", "#,##0")
for colL, lab in [("C", "國產 6 軸"), ("D", "外資 6 軸"), ("E", "國產協作")]:
    ws.cell(r, 1, f"{lab}回本期（基準）"); c = ws.cell(r, 3, f"='計算'!{colL}{PB_ROW['PB_YR']}"); c.number_format = "0.00"
    ws.cell(r, 5, "年"); ws.cell(r, 6, "計算頁 A"); r += 1
ws.cell(r, 1, "JARA 最新季訂單÷生產"); c = ws.cell(r, 3, f"='計算'!C{CROW['CY_BB']}"); c.number_format = "0.00"; ws.cell(r, 5, "倍"); ws.cell(r, 6, "追蹤指標頁"); r += 1
r += 1
ws.cell(r, 1, "三、已知口徑衝突與限制").font = BOLD; ws.cell(r, 1).fill = SEC; r += 1
notes = [
 "1. 減速機市場規模：公開產業報告數字互相矛盾（全球 RV 2024 年 20.8 億或 29.1 億美元；全球諧波 2023 年 4.4 億或 2024 年 12 億美元；全球諧波需求 2024 年 282 萬顆或 2025 年 107 萬顆）。本模型改由下而上（裝機×每台顆數×均價），並以 Nabtesco CMP 與 HDS 營收推得上限對帳（計算頁 D）。",
 "2. 協作裝機：IFR 2025 年 6.6 萬台；觀研天下另列 2025E 全球 16.4 萬台，口徑不同（可能含非工業用途），本模型用 IFR。",
 "3. 中國國家統計局產量（2025 年約 71 萬台，推得）約為 IFR 中國裝機的 2 倍：含出口、可能以套計。下游只用其增速。",
 "4. JARA 季度數字為會員口徑；JARA 全年預測（1.22 兆日圓）含非會員，兩者不可直接相比。",
 "5. 回本期的本體價區間極寬（6kg 小負載 2 萬元到 20kg 15 萬元），而整合倍數與替代人數沒有公開統計，回本期只能當區間讀，不是觀測值。",
 "6. 機型結構與每台顆數三情境共用基準；不確定性放在敏感度。中國減速機需求假設與全球同結構（中國小負載與 SCARA 比重可能更高）。",
 "7. 減速機均價為判斷值：以 2022 年中國均價（RV 約 4,530 元、諧波約 2,048 元，智研諮詢經東莞證券轉引）為錨，並非 2025 年觀測值。",
]
for n in notes:
    c = ws.cell(r, 1, n); c.alignment = WRAP; ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 45; r += 1
r += 1
ws.cell(r, 1, "四、版本").font = BOLD; ws.cell(r, 1).fill = SEC; r += 1
ws.cell(r, 1, "v0.1（2026-11-16）"); ws.cell(r, 2, "初版：由框架 v0.1 建成 Excel；補查 IFR 預測、JARA 四季、國家統計局工資、HDS 上半年決算、減速機用量與均價")
ws.column_dimensions["A"].width = 44
for col in "BCDE": ws.column_dimensions[col].width = 16
ws.column_dimensions["B"].width = 60
ws.column_dimensions["F"].width = 14

os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)
print("wrote", OUT)

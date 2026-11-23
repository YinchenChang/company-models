# 線 1b 智慧車（四輪機器人）單位經濟模型 v0.1
# 用法：python3 physical_ai/lines/smart_car/build_smart_car.py ；再以 LibreOffice headless 重算
# 依據：線1b_智慧車_框架_v0.1.md（2026-11-16）第 3 節式 A／式 B；第 5 節追蹤指標
# 頁面：說明｜輸入｜計算｜情境｜敏感度｜下游取數｜追蹤指標｜來源
# 金額：美元（USD）為主，線控底盤以人民幣（RMB）輸入再換算；車數：中國以萬台、Tesla 以百萬台
import os, copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "dist", "20261116_智慧車單位經濟_v0.1.xlsx")
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7"); JUDG = PatternFill("solid", fgColor="FFF2CC")
YEARS = list(range(2026, 2033))
SC = [("低", "D"), ("基準", "E"), ("高", "F")]

# ---------------- 來源 ----------------
S = {
 "tsla_q3": "Tesla 2026Q3 財報與簡報（經 notateslaapp 2026-10 轉述：FSD 活躍 174 萬、北美搭配率 >75%）https://www.notateslaapp.com/news/4862/tesla-fsd-hits-record-subs-purchased-with-75-of-new-orders",
 "tsla_del": "Tesla 2026Q3 交付公告（Electrek 2026-10-02：Q3 交付 486,532、前三季 1,324,681、Q3 產量 464,391）https://electrek.co/2026/10/02/tesla-q3-2026-deliveries-486532/",
 "tsla_8m": "Tesla 第 800 萬輛下線（2025-06，Giga Berlin）https://thedriven.io/2025/06/10/tesla-produces-its-8-millionth-ev-and-its-the-new-tesla-model-y/",
 "fsd_price": "Electrek 2026-01-14：美國 2026-02 起停售買斷、只剩訂閱 https://electrek.co/2026/01/14/tesla-tsla-stop-selling-full-self-driving-package-subscription-only/ ；月費 99 美元（news18a）https://english.news18a.com/m/news/english_232483.html",
 "eu": "Electrive 2026-09-28：EU 表決最快 12 月 https://www.electrive.com/2026/09/28/eu-vote-on-tesla-fsd-not-expected-before-december/ ；Electrive 2026-10-12 更名 Tesla Assisted Driving https://electrive.com/2026/10/12/from-full-self-driving-supervised-to-tesla-assisted-driving-with-a-new-name-aims-to-secure-eu-approval",
 "cn_ff": "CleanTechnica 2026-08-12／SCMP：中國資料不出境，FSD 核准仍遠 https://cleantechnica.com/2026/08/12/tesla-fsd-hurdles-in-china-secrecy-in-europe/",
 "sina_noa": "新浪汽車 2026-08-26：2026H1 城市 NOA 162 萬／861 萬＝18.8%；灼識諮詢 2030 年 62%；乾崑市佔 19.6% https://auto.sina.cn/2026-08-26/detail-iniprytn5653084.d.html",
 "ggai": "高工智能汽車研究院 2026H1（太平洋汽車轉述，網易）：城區 NOA（含選裝、預埋）204.69 萬輛；前裝標配智駕域控 317.32 萬套、滲透率 37.21%；城區 NOA 晶片英偉達 39.43%、地平線 22.82% https://c.m.163.com/news/a/L33K11L005279745.html",
 "caict": "CnEVPost 2026-10-24 引中國信通院：截至 2026-08 具 NOA 乘用車保有約 1,500 萬、L2 約 6,100 萬 https://cnevpost.com/2026/10/24/china-15-million-cars-with-noa",
 "cpca": "乘聯會 2026-10（MarkLines 轉述）：10 月零售 164.3 萬（−27.1%）；1–10 月零售 1,497.8 萬（−22.3%）；新能源零售滲透率 67.5% https://www.marklines.com/cn/report/cpca085_202611",
 "sina_sbw": "新浪汽車 2026-10-05：SBW 2026 滲透率約 5%、2030 有望超過 30%；國產 SBW 每套成本約 8,000 元（過去 2 萬元以上）；線控底盤整體 2030 達 40% https://auto.sina.cn/2026-10-05/detail-iniuawyt3411838.d.html",
 "kr36": "36氪／車云：EMB 每車約 2,500 元（2025）→ 約 1,100 元（2030）；EMB 滲透率 2026 約 5%、2030 約 40%；耐世特中國 SBW 份額約 35% https://www.36kr.com/p/3888250770750216",
 "robosense": "Gasgoo：速騰聚創 2026Q3 ADAS LiDAR 45.04 萬台（+200%），前三季 ADAS 88.7 萬 https://autonews.gasgoo.com/articles/market-industry/robosense-quarterly-lidar-sales-hit-new-high-in-q3-2026-2108495209121140737",
 "mbly": "Mobileye 2026Q3 8-K：EyeQ 出貨 950 萬套、ASP 49.1 美元（前年 51.7）https://www.sec.gov/Archives/edgar/data/0001910139/000110465926119830/tm2628197d1_ex99-1.htm",
 "hor": "Gasgoo：地平線 2026 中報，授權收入 11.29 億元、毛利率 66% https://autonews.gasgoo.com/articles/ICV/horizon-robotics-2026-interim-report-gross-margin-maintains-high-of-66-2095039844043800577",
 "momenta": "Gasgoo：Momenta 累計裝機破 100 萬（2026-07）https://autonews.gasgoo.com/articles/market-industry/momenta-officially-announces-installations-surpass-one-million-2074842447635861505",
 "byd": "Gasgoo：比亞迪輔助駕駛累計搭載 >285 萬（至 2026-03-31）https://autonews.gasgoo.com/articles/news/li-yunfei-reveals-byds-driver-assistance-installed-base-exceeds-285-million-units-2044742748492447745",
 "l3": "CnEVPost 2026-10-22：中國第二批 L3 試點推進 https://cnevpost.com/2026/10/22/china-advances-2nd-batch-level-3-approvals/ ；2026-08-04 L3/L4 安全基線（2027-07-01 生效）https://cnevpost.com/2026/08/04/china-sets-safety-baseline-l3-l4-autonomous-driving",
 "frame": "線1b_智慧車_框架_v0.1.md（2026-11-16）",
}

# ---------------- 輸入定義 ----------------
# (鍵, 項目, 單位, 低情境值, 基準, 高情境值, 標記, 類別, 出處, 說明)；低／高為 None → 等於基準（非判斷值）
# 欄位方向：「低情境」＝讓本線規模較小的值；「高情境」＝讓規模較大的值（不是數值大小）
INP = [
 ("sec", "A. 中國：分母與滲透率（式 A：滲透率）"),
 ("cn_sales26", "中國乘用車國內零售 2026", "萬台", 1800, 1830, 1860, "Derived", "A", S["cpca"], "1–10 月 1,497.8 萬＋11–12 月假設每月 150–180 萬；城市 NOA 滲透率分母"),
 ("cn_g", "中國乘用車國內銷量年增率 2027–2032", "%", -0.03, 0.0, 0.02, "Assumed", "A", "—", "2026 年 1–10 月 −22.3%（購置稅與補貼退坡），之後假設持平附近"),
 ("city26", "中國城市 NOA 滲透率 2026", "%", 0.188, 0.214, 0.240, "Verified", "I", S["sina_noa"] + "；" + S["ggai"], "兩口徑：18.8%（新浪，H1 162/861 萬）；24.0%（高工，含選裝與硬體預埋 204.69/852.8 萬）；基準取中點"),
 ("city30", "中國城市 NOA 滲透率 2030", "%", 0.40, 0.55, 0.62, "Assumed", "I", S["sina_noa"], "高＝灼識諮詢 62%（Interested-party：常為 IPO 顧問）；低／基準為判斷"),
 ("city32", "中國城市 NOA 滲透率 2032", "%", 0.50, 0.65, 0.75, "Assumed", "I", "—", "2030 後放緩；入門車款（<10 萬元）為天花板"),
 ("dc26", "中國前裝標配智駕域控制器滲透率 2026H1", "%", None, 0.3721, None, "Verified", "A", S["ggai"], "用來推高速 NOA（不含城市）＝域控滲透率−城市 NOA"),
 ("hwy30", "中國高速 NOA（不含城市）滲透率 2030", "%", 0.30, 0.25, 0.20, "Assumed", "I", "—", "城市 NOA 上升時高速-only 被吸收；低情境城市少故高速-only 多"),
 ("hwy32", "中國高速 NOA（不含城市）滲透率 2032", "%", 0.30, 0.20, 0.15, "Assumed", "I", "—", "同上"),
 ("l3_30", "中國 L3（車廠擔責）滲透率 2030", "%", 0.01, 0.04, 0.10, "Assumed", "I", S["l3"], "2026 僅試點路段（約 0%）；國標 2027-07-01 生效；L3 視為城市 NOA 的子集"),
 ("l3_32", "中國 L3 滲透率 2032", "%", 0.03, 0.08, 0.15, "Assumed", "I", "—", "同上"),
 ("sbw26", "中國線控轉向（SBW）滲透率 2026", "%", None, 0.05, None, "Interested-party", "I", S["sina_sbw"], "產業媒體預估值；GB 17675-2025 2026-07 取消機械連接強制"),
 ("sbw30", "中國 SBW 滲透率 2030", "%", 0.15, 0.30, 0.40, "Assumed", "I", S["sina_sbw"], "基準＝媒體「有望超過 30%」；高＝線控底盤整體 40%"),
 ("sbw32", "中國 SBW 滲透率 2032", "%", 0.25, 0.40, 0.55, "Assumed", "I", "—", "L3 冗餘需求帶動"),
 ("emb26", "中國 EMB 滲透率 2026", "%", None, 0.05, None, "Interested-party", "I", S["kr36"], "產業媒體彙整業界估計"),
 ("emb30", "中國 EMB 滲透率 2030", "%", 0.15, 0.30, 0.40, "Assumed", "I", S["kr36"], "高＝媒體 40%；平價點約 2028 年"),
 ("emb32", "中國 EMB 滲透率 2032", "%", 0.25, 0.40, 0.55, "Assumed", "I", "—", ""),
 ("sec", "B. 每車智能內容（式 A：內容單價，2026 年價）"),
 ("sens_city", "城市 NOA：相機＋毫米波雷達（不含 LiDAR）", "USD/車", 150, 220, 300, "Assumed", "P", "—", "約 11 相機＋1–5 雷達；未取得公開拆解"),
 ("lidar_att", "城市 NOA 車搭載 LiDAR 比例", "%", 0.50, 0.65, 0.80, "Assumed", "I", S["robosense"], "速騰 2026 前三季 ADAS LiDAR 88.7 萬台，對照城市 NOA 車數推估；純視覺陣營（Tesla、部分比亞迪）不裝"),
 ("lidar_n", "每車 LiDAR 數", "顆", 1.0, 1.0, 1.2, "Assumed", "P", "—", "主流 1 顆；旗艦 2–3 顆"),
 ("lidar_asp", "ADAS LiDAR 單價", "USD/顆", 150, 200, 300, "Analogy", "P", "—", "類比：速騰 MX、禾賽 ATX 主力款報價約 200 美元級（未取得 2026 公司揭露 ASP）"),
 ("comp_city", "城市 NOA：運算（SoC＋域控）", "USD/車", 200, 350, 500, "Assumed", "P", "只讀 AI供應鏈（待接）", "本線不建模晶片；暫放區間，對帳時以 AI供應鏈 為準"),
 ("lic_city", "城市 NOA：方案／授權費", "USD/車", 40, 80, 150, "Assumed", "I", S["hor"] + "；" + S["mbly"], "地平線授權毛利 66%；Mobileye 全產品 ASP 49.1 美元（含 L2 基本款）"),
 ("sens_hwy", "高速 NOA：感測", "USD/車", 60, 100, 150, "Assumed", "P", "—", ""),
 ("comp_hwy", "高速 NOA：運算（只讀，待接）", "USD/車", 60, 100, 150, "Assumed", "P", "只讀 AI供應鏈（待接）", ""),
 ("lic_hwy", "高速 NOA：方案／授權費", "USD/車", 10, 25, 50, "Assumed", "I", S["mbly"], "Mobileye ASP 49.1 美元為上緣參考"),
 ("l3_add", "L3 冗餘增量（第二套運算、感測、電源）", "USD/車", 400, 800, 1500, "Assumed", "P", "—", "不含線控底盤（另列）"),
 ("e_dec", "電子內容年降價率", "%", 0.12, 0.08, 0.05, "Assumed", "P", "—", "低情境降價快；套用感測、運算、授權、L3 增量"),
 ("sbw_v26", "SBW 每套價值 2026", "RMB/車", None, 8000, None, "Interested-party", "P", S["sina_sbw"], "國產供應商每套成本約 8,000 元（媒體；框架 v0.1 原為 Assumed 2,000–6,000，本版改用此值，對帳時注意）；為整套價值，非相對 EPS 增量"),
 ("sbw_dec", "SBW 年降價率", "%", 0.15, 0.10, 0.06, "Assumed", "P", "—", "過去 2 萬元以上降到 8,000 元"),
 ("emb_v25", "EMB 每車價值 2025", "RMB/車", None, 2500, None, "Interested-party", "P", S["kr36"], "業界估計"),
 ("emb_v30", "EMB 每車價值 2030", "RMB/車", None, 1100, None, "Interested-party", "P", S["kr36"], "隱含年降約 15%；EHB 平價點約 1,350 元（約 2028）"),
 ("fx", "人民幣兌美元", "RMB/USD", 7.3, 7.1, 6.9, "Assumed", "A", "—", "未取當日匯率；給區間"),
 ("sec", "C. Tesla FSD（式 B）"),
 ("act_q3", "FSD 活躍數 2026Q3 末（含早期買斷）", "百萬", None, 1.74, None, "Verified", "A", S["tsla_q3"], "年增 67%；Q2 末 1.48 百萬"),
 ("q4_add", "FSD 2026Q4 淨增", "百萬", 0.18, 0.24, 0.30, "Assumed", "A", S["tsla_q3"], "Q3 淨增約 0.26 百萬"),
 ("legacy", "活躍數中早期買斷戶（不付月費）2026 末", "百萬", 1.0, 0.8, 0.6, "Assumed", "A", "未取得（Tesla 未揭露拆分）", "低情境＝買斷戶多、付費戶少"),
 ("d_ytd", "Tesla 交付 2026 前三季", "百萬", None, 1.324681, None, "Verified", "A", S["tsla_del"], ""),
 ("d_q4", "Tesla 交付 2026Q4", "百萬", 0.42, 0.46, 0.50, "Assumed", "A", S["tsla_del"], "Q3 486,532；應以 Tesla 個股模型為準（對帳）"),
 ("d_g", "Tesla 交付年增率 2027–2032", "%", -0.03, 0.05, 0.12, "Assumed", "A", "—", "Robotaxi 專用車（Cybercab）另計於線 1a；應以 Tesla 個股模型為準（對帳）"),
 ("fleet26", "Tesla 在用車隊 2026 末", "百萬", 10.2, 10.5, 10.8, "Derived", "A", S["tsla_8m"] + "；" + S["tsla_del"], "800 萬（2025-06）＋2025H2 約 0.85–0.92 百萬＋2026 約 1.75 百萬，扣報廢約 2–3%"),
 ("scrap", "車隊年報廢率", "%", 0.03, 0.025, 0.02, "Assumed", "A", "—", "車齡年輕"),
 ("cov26", "FSD 可售市場占 Tesla 交付比重（2026，北美＋已核准國）", "%", 0.45, 0.50, 0.55, "Assumed", "I", S["eu"], "Tesla 未揭露分區交付"),
 ("eu_sh", "歐洲占 Tesla 交付比重", "%", 0.15, 0.18, 0.22, "Assumed", "A", "—", "可由 ACEA 登記數核對（未取）"),
 ("eu_yr", "EU 全體核准生效年", "年", 2029, 2027, 2027, "Assumed", "I", S["eu"], "表決最快 2026-12；需 15 國以上合格多數"),
 ("cn_sh", "中國占 Tesla 交付比重", "%", 0.30, 0.33, 0.36, "Assumed", "A", "—", "可由乘聯會 Tesla 中國零售核對（未取）"),
 ("cn_yr", "中國 FSD 核准生效年", "年", 2040, 2029, 2028, "Assumed", "I", S["cn_ff"], "低情境 2040＝期間內不核准"),
 ("att26", "可售市場新車 FSD 搭配率 2026", "%", None, 0.75, None, "Interested-party", "I", S["tsla_q3"], "Tesla 自揭北美 Q3 >75%（Q2 55%）；含試用轉換，有推升估值誘因"),
 ("att30", "可售市場新車 FSD 搭配率 2030（2031–32 持平）", "%", 0.45, 0.65, 0.80, "Assumed", "I", "—", "新市場搭配率低於北美"),
 ("conv", "存量非 FSD 車每年轉訂閱比例（可售市場）", "%", 0.02, 0.04, 0.07, "Assumed", "I", "—", "HW3 車能否跑新版為限制"),
 ("churn", "FSD 活躍數年流失率", "%", 0.10, 0.06, 0.03, "Assumed", "I", "—", "訂閱可隨時取消"),
 ("arpu26", "FSD 月費 2026（美國）", "USD/月", None, 99, None, "Verified", "A", S["fsd_price"], "歐洲傳 99 歐元未經官方核實"),
 ("arpu30", "FSD 混合月費 2030（含區域折價；2031–32 持平）", "USD/月", 60, 80, 99, "Assumed", "I", "—", ""),
 ("mc", "FSD 每車年邊際成本（推論、OTA、客服、責任準備）", "USD/車年", 200, 120, 60, "Assumed", "I", "未取得（Tesla 未揭露）", "框架：L3 責任準備 50–300 美元／車年（Analogy 車險）；FSD 仍屬監督式 L2，故取較低"),
]

wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR

# ---------------- 說明 ----------------
DOC = wb.active; DOC.title = "說明"
doc = [
 "線 1b 智慧車（四輪機器人）單位經濟模型 v0.1｜2026-11-16",
 "依據：" + S["frame"] + "。範圍：自駕相關每車內容（感測、運算、線控底盤、方案授權）與 Tesla FSD 軟體收入；乘用車銷量只作分母。運算晶片只讀 AI供應鏈（本模型放暫定區間待接）；訓練算力不做；Robotaxi 在線 1a。",
 "式 A（每車智能內容，中國新車平均）＝ 城市NOA滲透率×城市NOA內容 ＋ 高速NOA滲透率×高速NOA內容 ＋ L3滲透率×L3冗餘增量 ＋ (SBW滲透率×SBW價值 ＋ EMB滲透率×EMB價值)÷匯率",
 "式 B（FSD）：活躍數(t) ＝ 活躍數(t−1)×(1−流失) ＋ 交付(t)×可售比重(t)×搭配率(t) ＋ (車隊(t−1)−活躍數(t−1))×可售比重(t)×轉換率；年末年化營收＝(活躍數−買斷戶)×月費×12；每車年毛利＝月費×12−邊際成本",
 "格式：藍字＝輸入；黃底＝判斷（Assumed／Analogy／Interested-party 的情境值，可改）；黑字＝公式。輸入頁 D／E／F 欄＝低情境／基準／高情境所用的值；「低」指讓本線規模較小的值，不一定是數值較小（例如降價率、流失率、邊際成本，低情境用較大值）。",
 "標記：[Verified] 可公開驗證附出處；[Interested-party] 寫明利害方；[Analogy] 寫明類比；[Assumed] 判斷（一律給三值）；[Derived] 由其他數推算。類別：[P] 物理／[A] 會計恆等式／[I] 誘因資訊。",
 "情境：計算頁三個區塊（低／基準／高）用同一組公式，只換讀輸入頁的欄；情境頁彙整 2026–2032。內插：2026→2030 線性、2030→2032 線性。",
 "敏感度：以 build 腳本的 Python 鏡像（與計算頁同一組等式）逐一把單一輸入換成低情境／高情境值、其餘維持基準，量測 2030 年輸出變動；敏感度頁上方為承重變數排序，下方才是基準結果與一致性核對（Python 基準＝Excel 計算頁，差額應為 0）。",
 "下游取數：給個股模型（Tesla、耐世特、地平線、Mobileye 等）讀的輸出，全部以公式連到計算頁。",
 "已知限制：(1) 情境內各變數同向（低＝全部偏保守），未建變數間相關（例如 SBW 降價會提高滲透率）；(2) Tesla 交付與車隊應以 Tesla 個股模型為準，本模型暫放區間並列入對帳；(3) 運算內容為暫定區間，待接 AI供應鏈；(4) FSD 邊際成本與買斷戶拆分 Tesla 未揭露；(5) 中國城市 NOA 兩個口徑差約 5 個百分點；(6) 不含海外（歐美日）非 Tesla ADAS 內容，Mobileye 只列追蹤。",
 "建置：python3 physical_ai/lines/smart_car/build_smart_car.py；LibreOffice headless 重算後檢查 0 錯誤值。",
]
for i, t in enumerate(doc, 1):
    DOC.cell(i, 1, t).alignment = Alignment(wrap_text=True, vertical="top")
DOC["A1"].font = Font(bold=True, size=13); DOC.column_dimensions["A"].width = 150

# ---------------- 輸入 ----------------
IN = wb.create_sheet("輸入")
IN["A1"] = "輸入（藍字；黃底＝判斷可改）。D 低情境｜E 基準｜F 高情境；非判斷值的 D、F 以公式等於 E"
hdr(IN, 2, ["鍵", "項目", "單位", "低情境", "基準", "高情境", "標記", "類別", "出處", "說明"])
ROW = {}; r = 3; BASE = {}; LOW = {}; HIGH = {}
for it in INP:
    if it[0] == "sec":
        IN.cell(r, 1, it[1]).font = BOLD; r += 1; continue
    k, lab, unit, lo, b, hi, tag, cls, src, note = it
    ROW[k] = r; BASE[k] = b; LOW[k] = b if lo is None else lo; HIGH[k] = b if hi is None else hi
    IN.cell(r, 1, k); IN.cell(r, 2, lab); IN.cell(r, 3, unit)
    judg = tag in ("Assumed", "Analogy") or (lo is not None)
    for col, v in (("D", lo), ("E", b), ("F", hi)):
        c = IN[f"{col}{r}"]
        if col != "E" and v is None:
            c.value = f"=E{r}"
        else:
            c.value = v; c.font = BLUE
            if judg: c.fill = JUDG
        if unit == "%": c.number_format = "0.0%"
    IN.cell(r, 7, f"[{tag}]"); IN.cell(r, 8, f"[{cls}]"); IN.cell(r, 9, src); IN.cell(r, 10, note)
    r += 1
for col, w in zip("ABCDEFGHIJ", (12, 46, 10, 10, 10, 10, 18, 7, 70, 60)):
    IN.column_dimensions[col].width = w

def I(k, col):
    return f"輸入!${col}${ROW[k]}"

# ---------------- 計算 ----------------
CA = wb.create_sheet("計算")
CA["A1"] = "計算（框架第 3 節式 A、式 B）。三區塊同式，只換讀輸入頁欄位：低＝D、基準＝E、高＝F。欄 C–I＝2026–2032"
YC = {y: CL(3 + i) for i, y in enumerate(YEARS)}
LINES = [  # (代號, 名稱, 單位, 格式)
 ("sales", "中國乘用車國內零售", "萬台", "#,##0"),
 ("city", "城市 NOA 滲透率", "%", "0.0%"),
 ("hwy", "高速 NOA（不含城市）滲透率", "%", "0.0%"),
 ("l3", "L3 滲透率（城市 NOA 子集）", "%", "0.0%"),
 ("sbw", "SBW 滲透率", "%", "0.0%"),
 ("emb", "EMB 滲透率", "%", "0.0%"),
 ("idx", "電子內容價格指數（2026＝1）", "倍", "0.000"),
 ("c_city", "城市 NOA 每車內容", "USD/車", "#,##0"),
 ("c_hwy", "高速 NOA 每車內容", "USD/車", "#,##0"),
 ("c_l3", "L3 冗餘增量", "USD/車", "#,##0"),
 ("v_sbw", "SBW 每套價值", "RMB/車", "#,##0"),
 ("v_emb", "EMB 每車價值", "RMB/車", "#,##0"),
 ("avg_e", "每車自駕電子內容（新車平均）", "USD/車", "#,##0"),
 ("avg_w", "每車線控底盤內容（新車平均）", "USD/車", "#,##0"),
 ("avg", "每車智能內容（新車平均，式 A）", "USD/車", "#,##0"),
 ("tam", "中國智能內容市場", "十億 USD", "0.00"),
 ("sbw_mkt", "中國 SBW 市場", "億 RMB", "#,##0"),
 ("city_n", "城市 NOA 新車數", "萬台", "#,##0"),
 ("del", "Tesla 交付", "百萬", "0.000"),
 ("fleet", "Tesla 在用車隊（年末）", "百萬", "0.00"),
 ("cov", "FSD 可售市場比重", "%", "0.0%"),
 ("att", "可售市場新車搭配率", "%", "0.0%"),
 ("act", "FSD 活躍數（年末）", "百萬", "0.00"),
 ("adopt", "FSD 採用率（活躍÷車隊）", "%", "0.0%"),
 ("leg", "早期買斷戶", "百萬", "0.00"),
 ("pay", "付費訂閱戶", "百萬", "0.00"),
 ("arpu", "混合月費", "USD/月", "0.0"),
 ("rev", "FSD 年末年化營收", "十億 USD", "0.00"),
 ("gp_car", "FSD 每車年毛利", "USD/車年", "#,##0"),
 ("gp", "FSD 年末年化毛利", "十億 USD", "0.00"),
]
CELL = {}  # CELL[(sc, code, year)] = "計算!$C$5"
def interp(y, a, b, c):  # a=2026, b=2030, c=2032 formula strings
    return f"IF({y}<=2030,{a}+({b}-{a})*({y}-2026)/4,{b}+({c}-{b})*({y}-2030)/2)"
row = 3
BLOCK_START = {}
for sc, col in SC:
    CA.cell(row, 1, f"情境：{sc}（讀輸入頁 {col} 欄）").font = BOLD; row += 1
    hdr(CA, row, ["項目", "單位"] + YEARS); yr_row = row; row += 1
    R = {}
    for code, name, unit, fmt in LINES:
        R[code] = row; row += 1
    BLOCK_START[sc] = R
    g = lambda k: I(k, col)
    for code, name, unit, fmt in LINES:
        rr = R[code]; CA.cell(rr, 1, name); CA.cell(rr, 2, unit)
        for i, y in enumerate(YEARS):
            c = YC[y]; p = CL(2 + i)  # previous year column
            ref = lambda cd, cc=c: f"{cc}{R[cd]}"
            pref = lambda cd: f"{p}{R[cd]}"
            first = (y == 2026)
            Y = f"{c}${yr_row}"
            f = {
             "sales": g("cn_sales26") if first else f"{pref('sales')}*(1+{g('cn_g')})",
             "city": interp(Y, g("city26"), g("city30"), g("city32")),
             "hwy": interp(Y, f"MAX(0,{g('dc26')}-{g('city26')})", g("hwy30"), g("hwy32")),
             "l3": f"MIN({ref('city')},{interp(Y, '0', g('l3_30'), g('l3_32'))})",
             "sbw": interp(Y, g("sbw26"), g("sbw30"), g("sbw32")),
             "emb": interp(Y, g("emb26"), g("emb30"), g("emb32")),
             "idx": f"(1-{g('e_dec')})^({Y}-2026)",
             "c_city": f"({g('sens_city')}+{g('lidar_att')}*{g('lidar_n')}*{g('lidar_asp')}+{g('comp_city')}+{g('lic_city')})*{ref('idx')}",
             "c_hwy": f"({g('sens_hwy')}+{g('comp_hwy')}+{g('lic_hwy')})*{ref('idx')}",
             "c_l3": f"{g('l3_add')}*{ref('idx')}",
             "v_sbw": f"{g('sbw_v26')}*(1-{g('sbw_dec')})^({Y}-2026)",
             "v_emb": f"{g('emb_v25')}*({g('emb_v30')}/{g('emb_v25')})^(({Y}-2025)/5)",
             "avg_e": f"{ref('city')}*{ref('c_city')}+{ref('hwy')}*{ref('c_hwy')}+{ref('l3')}*{ref('c_l3')}",
             "avg_w": f"({ref('sbw')}*{ref('v_sbw')}+{ref('emb')}*{ref('v_emb')})/{g('fx')}",
             "avg": f"{ref('avg_e')}+{ref('avg_w')}",
             "tam": f"{ref('sales')}*10000*{ref('avg')}/1E9",
             "sbw_mkt": f"{ref('sales')}*10000*{ref('sbw')}*{ref('v_sbw')}/1E8",
             "city_n": f"{ref('sales')}*{ref('city')}",
             "del": f"{g('d_ytd')}+{g('d_q4')}" if first else f"{pref('del')}*(1+{g('d_g')})",
             "fleet": g("fleet26") if first else f"{pref('fleet')}*(1-{g('scrap')})+{ref('del')}",
             "cov": f"MIN(0.95,{g('cov26')}+IF({Y}>={g('eu_yr')},{g('eu_sh')},0)+IF({Y}>={g('cn_yr')},{g('cn_sh')},0))",
             "att": interp(Y, g("att26"), g("att30"), g("att30")),
             "act": f"{g('act_q3')}+{g('q4_add')}" if first else f"{pref('act')}*(1-{g('churn')})+{ref('del')}*{ref('cov')}*{ref('att')}+({pref('fleet')}-{pref('act')})*{ref('cov')}*{g('conv')}",
             "adopt": f"{ref('act')}/{ref('fleet')}",
             "leg": f"{g('legacy')}*(1-{g('scrap')})^({Y}-2026)",
             "pay": f"MAX(0,{ref('act')}-{ref('leg')})",
             "arpu": interp(Y, g("arpu26"), g("arpu30"), g("arpu30")),
             "rev": f"{ref('pay')}*{ref('arpu')}*12/1000",
             "gp_car": f"{ref('arpu')}*12-{g('mc')}",
             "gp": f"{ref('pay')}*{ref('gp_car')}/1000",
            }[code]
            cell = CA[f"{c}{rr}"]; cell.value = "=" + f; cell.number_format = fmt
            CELL[(sc, code, y)] = f"計算!${c}${rr}"
    row += 1
CA.column_dimensions["A"].width = 34; CA.column_dimensions["B"].width = 11
for y in YEARS: CA.column_dimensions[YC[y]].width = 11

# ---------------- Python 鏡像（敏感度用；與計算頁同式） ----------------
def ip(y, a, b, c):
    return a + (b - a) * (y - 2026) / 4 if y <= 2030 else b + (c - b) * (y - 2030) / 2
def model(P):
    o = {}; sales = None; prev = {}
    for y in YEARS:
        sales = P["cn_sales26"] if y == 2026 else sales * (1 + P["cn_g"])
        city = ip(y, P["city26"], P["city30"], P["city32"])
        hwy = ip(y, max(0, P["dc26"] - P["city26"]), P["hwy30"], P["hwy32"])
        l3 = min(city, ip(y, 0, P["l3_30"], P["l3_32"]))
        sbw = ip(y, P["sbw26"], P["sbw30"], P["sbw32"]); emb = ip(y, P["emb26"], P["emb30"], P["emb32"])
        idx = (1 - P["e_dec"]) ** (y - 2026)
        c_city = (P["sens_city"] + P["lidar_att"] * P["lidar_n"] * P["lidar_asp"] + P["comp_city"] + P["lic_city"]) * idx
        c_hwy = (P["sens_hwy"] + P["comp_hwy"] + P["lic_hwy"]) * idx
        c_l3 = P["l3_add"] * idx
        v_sbw = P["sbw_v26"] * (1 - P["sbw_dec"]) ** (y - 2026)
        v_emb = P["emb_v25"] * (P["emb_v30"] / P["emb_v25"]) ** ((y - 2025) / 5)
        avg = city * c_city + hwy * c_hwy + l3 * c_l3 + (sbw * v_sbw + emb * v_emb) / P["fx"]
        tam = sales * 1e4 * avg / 1e9
        d = P["d_ytd"] + P["d_q4"] if y == 2026 else prev["del"] * (1 + P["d_g"])
        fleet = P["fleet26"] if y == 2026 else prev["fleet"] * (1 - P["scrap"]) + d
        cov = min(0.95, P["cov26"] + (P["eu_sh"] if y >= P["eu_yr"] else 0) + (P["cn_sh"] if y >= P["cn_yr"] else 0))
        att = ip(y, P["att26"], P["att30"], P["att30"])
        act = P["act_q3"] + P["q4_add"] if y == 2026 else prev["act"] * (1 - P["churn"]) + d * cov * att + (prev["fleet"] - prev["act"]) * cov * P["conv"]
        leg = P["legacy"] * (1 - P["scrap"]) ** (y - 2026)
        pay = max(0, act - leg)
        arpu = ip(y, P["arpu26"], P["arpu30"], P["arpu30"])
        rev = pay * arpu * 12 / 1000
        prev = {"del": d, "fleet": fleet, "act": act}
        o[y] = dict(avg=avg, tam=tam, city=city, sbw=sbw, act=act, adopt=act / fleet, rev=rev)
    return o

# ---------------- 敏感度 ----------------
SE = wb.create_sheet("敏感度")
OUTS = [("tam", "中國智能內容市場 2030（十億 USD）"), ("act", "FSD 活躍數 2030（百萬）"), ("rev", "FSD 年末年化營收 2030（十億 USD）")]
base_o = model(BASE)
SE["A1"] = "敏感度：單一輸入換成低情境／高情境值、其餘維持基準，2030 年輸出的變動。先看承重變數排序，基準結果在下方"
SE["A2"] = "計算方式：build 腳本 Python 鏡像，與計算頁同一組等式（值，非公式）；下方一致性核對以公式比對計算頁基準，差額須為 0"
row = 4; SENS_TOP = {}
for code, title in OUTS:
    SE.cell(row, 1, f"承重變數排序｜{title}").font = BOLD; row += 1
    hdr(SE, row, ["排名", "輸入鍵", "項目", "低情境值", "高情境值", "輸出（低）", "輸出（高）", "擺幅（高−低）", "擺幅占基準"]); row += 1
    res = []
    for k in ROW:
        if LOW[k] == HIGH[k]: continue
        Pl = dict(BASE); Pl[k] = LOW[k]; Ph = dict(BASE); Ph[k] = HIGH[k]
        ol = model(Pl)[2030][code]; oh = model(Ph)[2030][code]
        if abs(oh - ol) < 1e-12: continue
        res.append((abs(oh - ol), k, ol, oh))
    res.sort(reverse=True); SENS_TOP[code] = res[:3]
    b = base_o[2030][code]
    labels = {it[0]: it[1] for it in INP if it[0] != "sec"}
    for n, (sw, k, ol, oh) in enumerate(res[:10], 1):
        vals = [n, k, labels[k], LOW[k], HIGH[k], ol, oh, oh - ol, (oh - ol) / b]
        for j, v in enumerate(vals, 1):
            c = SE.cell(row, j, v)
            if j in (6, 7, 8): c.number_format = "0.00"
            if j == 9: c.number_format = "0%"
        row += 1
    row += 1
SE.cell(row, 1, "基準結果與一致性核對（2030）").font = BOLD; row += 1
hdr(SE, row, ["輸出", "Python 基準", "計算頁基準（公式）", "差額"]); row += 1
for code, title in OUTS:
    SE.cell(row, 1, title); SE.cell(row, 2, base_o[2030][code])
    SE.cell(row, 3, "=" + CELL[("基準", code, 2030)]); SE.cell(row, 4, f"=ROUND(C{row}-B{row},6)")
    row += 1
SE.column_dimensions["A"].width = 40; SE.column_dimensions["C"].width = 46
for col in "BDEFGHI": SE.column_dimensions[col].width = 13

# ---------------- 情境 ----------------
SN = wb.create_sheet("情境")
SN["A1"] = "情境（2026–2032；低／基準／高）— 全部以公式連到計算頁"
SHOW = [("city", "中國城市 NOA 滲透率", "0.0%"), ("l3", "中國 L3 滲透率", "0.0%"), ("sbw", "中國 SBW 滲透率", "0.0%"), ("emb", "中國 EMB 滲透率", "0.0%"),
        ("avg", "每車智能內容（中國新車平均）USD", "#,##0"), ("c_city", "城市 NOA 每車內容 USD", "#,##0"), ("tam", "中國智能內容市場 十億 USD", "0.00"),
        ("sbw_mkt", "中國 SBW 市場 億 RMB", "#,##0"), ("act", "FSD 活躍數 百萬", "0.00"), ("adopt", "FSD 採用率（活躍÷車隊）", "0.0%"),
        ("rev", "FSD 年末年化營收 十億 USD", "0.00"), ("gp", "FSD 年末年化毛利 十億 USD", "0.00")]
hdr(SN, 3, ["指標", "情境"] + YEARS); row = 4
for code, name, fmt in SHOW:
    for sc, _ in SC:
        SN.cell(row, 1, name); SN.cell(row, 2, sc)
        for i, y in enumerate(YEARS):
            c = SN.cell(row, 3 + i, "=" + CELL[(sc, code, y)]); c.number_format = fmt
        row += 1
SN.column_dimensions["A"].width = 36

# ---------------- 下游取數 ----------------
DS = wb.create_sheet("下游取數")
DS["A1"] = "下游取數：給個股模型（Tesla、耐世特、伯特利、Mobis、地平線、Mobileye、Momenta、禾賽／速騰等）讀的輸出；值為公式，重算後即最新"
hdr(DS, 2, ["名稱", "年", "低", "基準", "高", "單位", "口徑", "建議讀者"])
DSL = [
 ("FSD 活躍數（年末，含早期買斷）", "act", [2026, 2027, 2028, 2030, 2032], "百萬", "Tesla 揭露口徑（訂閱＋買斷）；2026 為 Q3 實績＋Q4 假設", "Tesla"),
 ("FSD 採用率（活躍÷在用車隊）", "adopt", [2026, 2028, 2030, 2032], "%", "存量口徑；車隊為 Tesla 全球在用車（估）", "Tesla"),
 ("FSD 可售市場新車搭配率", "att", [2026, 2028, 2030], "%", "流量口徑；2026＝北美 Q3 自揭 >75%", "Tesla"),
 ("FSD 年末年化營收", "rev", [2026, 2027, 2028, 2030], "十億 USD", "（活躍−買斷戶）×混合月費×12；不含買斷一次性收入", "Tesla"),
 ("FSD 年末年化毛利", "gp", [2027, 2030], "十億 USD", "邊際成本 Tesla 未揭露（Assumed）", "Tesla"),
 ("每車智能內容（中國新車平均）", "avg", [2026, 2028, 2030, 2032], "USD/車", "式 A：自駕電子（感測、運算暫定、授權、L3 冗餘）＋線控底盤（SBW、EMB 整套價值）", "地平線、Momenta、禾賽／速騰、耐世特、伯特利"),
 ("城市 NOA 每車內容", "c_city", [2026, 2030], "USD/車", "城市 NOA 車；運算為暫定待接 AI供應鏈", "地平線、Momenta、禾賽／速騰"),
 ("中國城市 NOA 滲透率", "city", [2026, 2027, 2028, 2030, 2032], "%", "國內乘用車零售；2026 為兩口徑中點（18.8% 標配口徑／24.0% 含選裝預埋）", "地平線、Momenta、比亞迪、小鵬、理想、小米"),
 ("中國城市 NOA 新車數", "city_n", [2026, 2030], "萬台", "滲透率×國內零售", "地平線、Momenta"),
 ("中國 L3 滲透率", "l3", [2028, 2030, 2032], "%", "車廠擔責 L3；城市 NOA 子集", "耐世特、伯特利、Mobis"),
 ("中國線控轉向（SBW）滲透率", "sbw", [2026, 2027, 2028, 2030, 2032], "%", "國內乘用車零售；2026 為產業媒體預估 5%", "耐世特、浙江世宝、Mobis"),
 ("中國 SBW 每套價值", "v_sbw", [2026, 2030], "RMB/車", "整套價值（非相對 EPS 增量）", "耐世特"),
 ("中國 SBW 市場", "sbw_mkt", [2026, 2030, 2032], "億 RMB", "滲透率×價值×銷量", "耐世特、浙江世宝"),
 ("中國 EMB 滲透率", "emb", [2026, 2030], "%", "國內乘用車零售", "伯特利"),
 ("中國智能內容市場", "tam", [2026, 2030, 2032], "十億 USD", "每車智能內容×國內零售", "全線"),
]
row = 3; DSROW = {}
for name, code, yrs, unit, scope, who in DSL:
    fmt = dict((l[0], l[3]) for l in LINES)[code]
    for y in yrs:
        DS.cell(row, 1, name); DS.cell(row, 2, y)
        for j, (sc, _) in enumerate(SC):
            c = DS.cell(row, 3 + j, "=" + CELL[(sc, code, y)]); c.number_format = fmt
        DS.cell(row, 6, unit); DS.cell(row, 7, scope); DS.cell(row, 8, who)
        DSROW[(code, y)] = row; row += 1
for col, w in zip("ABCDEFGH", (34, 7, 11, 11, 11, 11, 70, 40)): DS.column_dimensions[col].width = w

# ---------------- 追蹤指標 ----------------
TR = wb.create_sheet("追蹤指標")
TR["A1"] = "追蹤指標（框架第 5 節）＋最新值與日期"
hdr(TR, 2, ["指標", "頻率", "資料來源", "最新值", "日期", "出處", "為何重要"])
TRK = [
 ("FSD 活躍數", "季", "Tesla 財報與簡報", "174 萬（年增 67%；Q2 末 148 萬）", "2026Q3（2026-10-21）", S["tsla_q3"], "式 B 核心"),
 ("FSD 北美新車搭配率", "季", "Tesla 財報與簡報", ">75%（Q2 55%）", "2026Q3", S["tsla_q3"], "式 B 流量"),
 ("Tesla 交付", "季", "Tesla 交付公告", "Q3 486,532；前三季 1,324,681", "2026-10-02", S["tsla_del"], "分母"),
 ("EU FSD（Tesla Assisted Driving）表決", "事件", "歐盟 TCMV、荷蘭 RDW", "表決延後，最快 2026-12；已 9 國單獨核准", "2026-10", S["eu"], "歐洲車隊能否計入"),
 ("中國 L3 第二批試點", "事件", "工信部裝備工業一司", "第二批推進中；首批長安、北汽（2025-12）", "2026-10-22", S["l3"], "L3 擴散"),
 ("中國城市 NOA 滲透率", "月／半年", "高工智能汽車、新浪／灼識、懂車帝", "18.8%（標配口徑）／24.0%（含選裝預埋）", "2026H1", S["sina_noa"] + "；" + S["ggai"], "式 A 滲透率"),
 ("中國具 NOA 乘用車保有", "半年", "中國信通院", "約 1,500 萬（L2 約 6,100 萬）", "2026-08", S["caict"], "存量"),
 ("中國智駕域控標配滲透率", "半年", "高工智能汽車", "37.21%（317.32 萬套）", "2026H1", S["ggai"], "高速 NOA 推算"),
 ("城區 NOA 晶片份額", "半年", "高工智能汽車", "英偉達 39.43%、地平線 22.82%、華為第三", "2026H1", S["ggai"], "價值歸屬"),
 ("中國乘用車零售", "月", "乘聯會", "10 月 164.3 萬（−27.1%）；1–10 月 1,497.8 萬（−22.3%）", "2026-10", S["cpca"], "分母"),
 ("SBW 滲透率", "年", "產業媒體、耐世特／浙江世宝揭露", "2026 約 5%（預估）；2030 >30%（預估）", "2026-10-05", S["sina_sbw"], "線控底盤放量"),
 ("EMB 價值與滲透率", "年", "產業媒體、伯特利揭露", "2,500 元（2025）→1,100 元（2030）；滲透率 5%→40%", "2026", S["kr36"], "線控底盤放量"),
 ("Mobileye ASP 與出貨", "季", "Mobileye 8-K", "950 萬套；ASP 49.1 美元", "2026Q3", S["mbly"], "海外 ADAS 價格"),
 ("地平線授權收入", "半年", "地平線中報", "11.29 億元；毛利率 66%", "2026H1", S["hor"], "授權模式"),
 ("Momenta 裝機", "半年", "Momenta 港股揭露", ">100 萬", "2026-07", S["momenta"], "第三方方案商"),
 ("比亞迪輔助駕駛累計搭載", "半年", "比亞迪", ">285 萬", "2026-03-31", S["byd"], "標配免費路線規模"),
 ("速騰 ADAS LiDAR 出貨", "季", "速騰聚創", "Q3 45.04 萬台（+200%）", "2026Q3", S["robosense"], "LiDAR 搭載率"),
 ("SBW／EMB 定點與量產車型", "季", "耐世特、伯特利、Mobis 財報", "未取得（本輪未查）", "—", "—", "放量節奏"),
 ("NHTSA AV Framework 規則進度", "事件", "NHTSA、Federal Register", "2026-06 提 FMVSS 135、撤回 AV STEP", "2026-06", S["frame"], "美國無方向盤車輛"),
]
for i, t in enumerate(TRK, 3):
    for j, v in enumerate(t, 1): TR.cell(i, j, v)
for col, w in zip("ABCDEFG", (30, 9, 30, 44, 18, 70, 20)): TR.column_dimensions[col].width = w

# ---------------- 來源 ----------------
SR = wb.create_sheet("來源")
hdr(SR, 1, ["代碼", "來源"])
for i, (k, v) in enumerate(S.items(), 2):
    SR.cell(i, 1, k); SR.cell(i, 2, v)
SR.column_dimensions["B"].width = 160

wb.move_sheet("情境", offset=-1)
wb.save(OUT)
print("saved", OUT)
for code, _ in OUTS:
    print(code, "base2030", round(base_o[2030][code], 3), "top3", [(k, round(sw, 3)) for sw, k, _, _ in SENS_TOP[code]])

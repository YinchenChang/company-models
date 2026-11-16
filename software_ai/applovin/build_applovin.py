#!/usr/bin/env python3
"""AppLovin（APP）個股模型 v0.1：原型 A（按廣告計價）的單位經濟驅動單一官方分部 Advertising。
用法：python3 software_ai/applovin/build_applovin.py <輸出 raw.xlsx> [--sens sens.json] [--kreq kreq.json] [--set key=值 ...] [--rowmap rowmap.json]
全部活公式（openpyxl）；輸出須經 LibreOffice 重算後才有快取值。--set：覆寫 In 頁基準欄（敏感度與反解整本重算用）。
"""
import csv, datetime as dt, hashlib, json, os, re, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

HERE = os.path.dirname(os.path.abspath(__file__))
VER = "0.1"
BUILD_DATE = dt.date(2026, 11, 16)
DATA_DATE = dt.date(2026, 11, 16)
PX_DATE = dt.date(2026, 11, 13)
VAL_DATE = dt.date(2026, 9, 30)

# ───────────────────────── 事實總帳（data/facts_applovin.csv）─────────────────────────
FACTS = {}
for r in csv.DictReader(open(os.path.join(HERE, "data", "facts_applovin.csv"), encoding="utf-8-sig")):
    FACTS[(r["指標"], r["期間"])] = r


def F(metric, period):
    v = FACTS[(metric, period)]["數值"]
    return float(v)


# ───────────────────────── 樣式 ─────────────────────────
BLUE = Font(name="Arial", size=10, color="0000FF")
BLK = Font(name="Arial", size=10, color="000000")
BOLD = Font(name="Arial", size=10, bold=True)
HDR = Font(name="Arial", size=10, bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="0B7A80")
SFILL = PatternFill("solid", fgColor="E6F4F3")
YFILL = PatternFill("solid", fgColor="FFF2CC")
WRAP = Alignment(wrap_text=True, vertical="top")

YEARS = list(range(2023, 2033))          # C..L
YC = {y: L(3 + i) for i, y in enumerate(YEARS)}   # 2023→C … 2032→L
FC = [YC[y] for y in range(2027, 2033)]  # G..L 預測年
COL9, COLQ4 = "M", "N"                   # 9M26A、Q4 26E
ISO = lambda d: dt.datetime(d.year, d.month, d.day)

# ───────────────────────── In 頁定義 ─────────────────────────
# (key, 項目, 單位, 低, 基準, 高, 標記, 來源, 說明, 來源誘因)
IP_CO = "公司管理層自揭／口頭陳述：有維持成長敘事與股價的誘因；口徑由公司自訂、未經查核"
IP_SELL = "賣方預估或資料商彙整的賣方共識：賣方與承銷、評等業務有利害關係；資料商彙整時點與樣本數不一，未經查核"
IP_WEB = "第三方資料網站：計算期間與口徑未完全揭露，無直接利益但未經查核"
SRC_Q3 = "AppLovin 2026Q3 財報新聞稿（8-K 附件 99.1，2026-11-04）"
SRC_10Q = "AppLovin 10-Q（2026-09-30 止季度）"
SRC_10K = "AppLovin 10-K FY2025（2026-02-19）"
TKSNAP = "software_ai/applovin/data/tokenomics_snapshot.json（Tokenomics v22.0，commit 79f6b30）"

IN_SECTIONS = []   # [(標題, [條目…])]


def S(title, *items):
    IN_SECTIONS.append((title, list(items)))


def E(key, label, unit, lo, base, hi, tag, src, note="", inc=""):
    return dict(key=key, label=label, unit=unit, lo=lo, base=base, hi=hi, tag=tag, src=src, note=note, inc=inc)


S("A. 日期、現價、股數與資產負債（評價日 2026-09-30）",
  E("data_date", "資料截止日（本版最晚一筆輸入：共識頁 TickFlow 2026-11-16；Tokenomics 快照擷取日 2026-11-16）", "日期", None, ISO(DATA_DATE), None, "Derived", "data/facts_applovin.csv 最晚擷取日；data/tokenomics_snapshot.json source.extractedAt", "原則第 10 節第 41 條：README 首列與 HTML 表頭分列資料截止日、成品日、股價基準日"),
  E("px_date", "股價基準日（現價日；最近一個交易日）", "日期", None, ISO(PX_DATE), None, "Verified", "同 px 列來源", "2026-11-16（週一）為資料截止日，當日收盤尚未取得，故基準日為 2026-11-13（週五）"),
  E("val_date", "評價日（最近一季資產負債表日）", "日期", None, ISO(VAL_DATE), None, "Verified", "AppLovin 10-Q 資產負債表日（2026-09-30）"),
  E("tp_date", "目標時點＝EDATE(股價基準日, 12)", "日期", None, "=EDATE({px_date},12)", None, "Derived", "原則第 10 節第 26 條", "持有期間＝現價日起 12 個月；折現期間＝評價日起（Val 頁並列）"),
  E("px", "現價（2026-11-13 收盤）", "US$", None, 409.10, None, "Verified", "StockAnalysis 歷史價格頁 https://stockanalysis.com/stocks/app/history/（2026-11-16 讀取）", "單一來源（統計頁、預估頁亦為同站）；同頁 11-06 收盤 391.22 與 TIKR 文載 391.22 一致。只影響上檔空間、權益權重與庫藏股法股數"),
  E("sh_out", "流通股數（Class A＋B；2026-09-30）", "百萬股", None, 331.322, None, "Verified", SRC_Q3 + " 資產負債表附註（Class A 301.258＋Class B 30.064）"),
  E("rsu_unv", "未歸屬 RSU（2025-12-31；Q3 未揭露新值）", "百萬股", None, 0.358162, None, "Verified", SRC_10K + " Note 11 Equity awards", "逾期一季：2026-09-30 未歸屬數未載；DCF 標準規格第 3 節：RSU 全額計入（無沒收率揭露）"),
  E("opt_n", "選擇權流通數（全數已可行使；2025-12-31）", "百萬股", None, 1.256114, None, "Verified", SRC_10K + " Note 11"),
  E("opt_k", "選擇權加權平均履約價", "US$", None, 6.10, None, "Verified", SRC_10K + " Note 11"),
  E("dil_inc", "評價股數計入未歸屬 RSU（1＝是；0＝敏感度，原口徑）", "x", 0, 1, None, "Assumed", "DCF 標準規格第 3 節（Andy 2026-11-09 J1 選 b）", "2025 年 10 月授予的績效股 PSU 0.920526 百萬股（績效／市場條件，達標機率無證據）只作說明，不計入"),
  E("cash_0", "現金及約當現金（2026-09-30）", "US$M", None, 3173.731, None, "Verified", SRC_Q3 + " 資產負債表"),
  E("debt_0", "優先無擔保票據本金（2029／2031／2034／2054；循環信貸未動用）", "US$M", None, 3550.0, None, "Verified", SRC_10K + " Note 9 Debt（本金 3,550；帳面 2026-09-30 為 3,516.136）", "10-Q 載票據總額 3.6bn；未見 2026 年償還紀錄"),
  E("fin_lease", "融資租賃負債", "US$M", None, 0.0, None, "Assumed", "未找到：新聞稿未單列融資租賃負債（年度租賃本金支付約 18.7）", "金額小，列 0 並在未找到清單註明"),
  E("eq_inv", "權益法投資（Tripledot 持股；帳面）", "US$M", 0.0, 298.333, None, "Verified", SRC_Q3 + " 資產負債表", "非營業資產按帳面計入；低案＝全額減損"),
  E("legal_oth", "類債：訴訟與和解（未入帳；無損失估計）", "US$M", None, 0.0, 1000.0, "Assumed", SRC_10Q + " Note 5：無須提列或揭露之重大損失；2026-09-16 起證券集體訴訟、衍生訴訟（loss 未估）", "事件型尾端風險（原則第 10 節第 19 條）：基準 0；高案 1,000 為示範尾端值，非估計"),
  E("cons_tp", "市場共識目標價（32 位；2026-11-10）", "US$", None, 453.28, None, "Interested-party", "StockAnalysis https://stockanalysis.com/stocks/app/forecast/（S&P Global 彙整；高 600、低 347）", "只作對照，不進模型", IP_SELL),
)
S("B. 折現率與永續價值",
  E("rf", "無風險利率（10 年美債；2026-10-27，跨模型沿用）", "%", None, 0.0499, None, "Verified", "美國財政部每日殖利率曲線（10 Yr）2026-10-27 4.99%（Meta v0.2.1 In 頁 rf，沿用同一讀值）", "逾期：與現價日相差超過 2 個交易日；原則第 10 節第 15 條：跨模型共用取值維持一致，下次全面改版一併更新"),
  E("beta_a", "原始 β：StockAnalysis（5Y；2026-11-07）", "x", None, 2.50, None, "Interested-party", "https://stockanalysis.com/stocks/app/statistics/", "", IP_WEB),
  E("beta_b", "原始 β：Yahoo Finance「Beta (5Y Monthly)」（日期未載）", "x", None, 2.56, None, "Interested-party", "https://finance.yahoo.com/quote/APP/key-statistics/", "", IP_WEB),
  E("blume", "Blume 調整係數", "x", None, 0.67, None, "Derived", "Blume（1971）；同 Meta／Alphabet"),
  E("beta", "β（Blume 調整後兩個來源的中位數）", "x", "=MIN({blume}*{beta_a}+(1-{blume}),{blume}*{beta_b}+(1-{blume}))", "=MEDIAN({blume}*{beta_a}+(1-{blume}),{blume}*{beta_b}+(1-{blume}))", "=MAX({blume}*{beta_a}+(1-{blume}),{blume}*{beta_b}+(1-{blume}))", "Derived", "同 Meta v0.2.1 F4 做法", "低／高＝兩個來源各自的調整值；原始 β 約 2.5 使股權成本約 15%，是 DCF 的主導輸入"),
  E("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節預設 5%（敏感度 4–6%）"),
  E("int_q3", "2026Q3 利息費用（單季）", "US$M", None, 51.066, None, "Verified", SRC_Q3 + " Statements of Operations"),
  E("kd", "稅前債務成本＝2026Q3 利息費用年化 ÷ 票據本金（有效利率）", "%", None, "=4*{int_q3}/{debt_0}", None, "Derived", "DCF 標準規格第 2 節：公司債有效利率", "10-K 票面利率區間未逐檔核對，不用票面"),
  E("wacc_adj", "WACC 加碼（敏感度用；基準 0）", "%", -0.01, 0.0, 0.01, "Assumed", "DCF 標準規格第 4 節：WACC ±1pt"),
  E("g_tv", "永續成長率（FY2032 後）", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格第 3 節預設 3%（敏感度 2–4%）"),
  E("ronic_sp", "永續期 RONIC 相對 WACC 的差", "%", None, 0.0, 0.05, "Assumed", "DCF 標準規格第 3 節：RONIC＝WACC（價值中性）", "Sens：RONIC＝WACC＋5pt"),
  E("tv_alt", "終值方法（0＝NOPAT×(1−g÷RONIC)；1＝FCFF×(1＋g) 替代法）", "x", None, 0, 1, "Assumed", "DCF 標準規格第 4 節：終值改用 FCFF×(1＋g) 的替代法", "只作敏感度"),
  E("sbc_add", "DCF 是否加回股票薪酬（0＝不加回，預設；1＝敏感度）", "x", None, 0, 1, "Assumed", "DCF 標準規格第 1 節：股票薪酬不加回（預設；加回只作敏感度）"),
  E("mid", "期中慣例：年度現金流距年初", "年", None, 0.5, None, "Derived", "DCF 標準規格第 1 節"),
)
S("C. 稅、現金、回購與股數",
  E("tax", "有效稅率（預測期）", "%", 0.131, 0.17, 0.21, "Assumed", "2026 年前三季有效稅率 16.4%（769.815 ÷ 4,680.868，Derived）、Q3 17.5%；FY2025 13.1%（低案）；美國法定稅率 21%（高案）", "公司未給 2026 全年稅率指引（未找到）"),
  E("r_cash", "現金收益率（利息收入）", "%", 0.03, 0.04, 0.05, "Assumed", "判斷：略低於 10 年美債 4.99%", "2026 年前三季其他收益 191.071 含權益法投資損益與匯兌，預測期只計利息"),
  E("bb_pct", "庫藏股買回占自由現金流比例", "%", 0.60, 0.85, 1.00, "Assumed", "2026 年前三季買回 2,811 ÷ 自由現金流 3,569＝79%；Q3 買回 1.28bn ÷ 1.42bn＝90%（Derived）", "董事會 2026-10 增加授權 3.2bn，餘額 3.6bn（SEC 10-Q 載 394.3 百萬加 3.2bn）"),
  E("px_g", "買回價格年漂移（自現價日起）", "%", 0.0, 0.10, 0.20, "Assumed", "判斷：折現率 15% 下以較低漂移保守估計"),
  E("sh_iss", "年度股票薪酬發行淨股數（扣淨股份交割代扣）", "百萬股", 0.0, 0.5, 1.5, "Assumed", "2026 年前三季流通股 338.313→331.322（−6.991）、買回退休 7.329，推得淨發行約 +0.34（Derived）"),
  E("dil_base", "加權稀釋增加股數（2026Q3；RSU、選擇權）", "百萬股", None, 1.141, None, "Derived", SRC_Q3 + " 加權稀釋股數 334.769 減基本 333.628"),
  E("min_cash", "最低現金（低於時補新債）", "US$M", None, 1000.0, None, "Assumed", "判斷；基準情境不觸發"),
)
S("D. 命題模組（原型 A：按廣告計價；安裝量 × 單次安裝淨收入）",
  E("q4_qoq", "2026Q4 營收季增率（相對 Q3）", "%", 0.08, 0.125, 0.19, "Assumed", "歷史 Q4 對 Q3：2024 年 +19.7%（Advertising 999.487 ÷ 835）、2025 年 +18.0%（1,657.944 ÷ 1,405.045）；Q4 指引 2,365–2,395 隱含 +8% 至 +10%（Interested-party）", "基準取介於指引與歷史季節性之間；低案＝指引中點附近、高案＝歷史均值。不以指引校準（原則第 10 節第 59 條）"),
  E("ginst0", "安裝量年增（組織性，扣分流前）", "%", -0.03, 0.02, 0.06, "Assumed", "公司 KPI：FY2024 +50%、FY2025 +3%、2026 年前三季 −7%、Q3 +2%（Interested-party）", "公司只揭露年增率，不揭露絕對值；本模型以 FY2026 ＝1.00 的指數表示"),
  E("nrpi_base", "單次安裝淨收入年增（非 AI 的價格與組合）", "%", 0.0, 0.03, 0.06, "Assumed", "判斷：一般廣告通膨；無 AI 對照路徑用"),
  E("u_decay", "AI 提升（單次安裝淨收入）逐年衰減係數", "x", 0.5, 0.78, 0.85, "Assumed", "判斷：新模型世代的邊際收益遞減；公司說明改善「不會按季度節奏」（Interested-party）"),
  E("e_decay", "新垂直（電商、連網電視）增量逐年衰減係數", "x", 0.6, 0.8, 0.9, "Assumed", "判斷：基期擴大後相對成長遞減"),
  E("div_yrs", "分流比例爬升年數（FY2027–FY2032）", "年", None, 6, None, "Assumed", "判斷：線性爬升到期末值"),
  E("u0_A", "情境 A：FY2027 AI 提升單次安裝淨收入", "%", None, 0.15, None, "Assumed", "情境依據：新模型全面上線（公司說明目前部分上線）；公司長期複合成長約 30% 的口頭說法為上緣參照（Interested-party）", "情境 A（AI 效益充分釋放）"),
  E("u0_B", "情境 B：FY2027 AI 提升單次安裝淨收入", "%", None, 0.10, None, "Assumed", "基準：FY2027 營收成長約 26%，與共識 FY2027 成長（約 22–26%）同量級，由本模型參數推得，非校準值", "情境 B（基準）"),
  E("u0_C", "情境 C：FY2027 AI 提升單次安裝淨收入", "%", None, 0.05, None, "Assumed", "情境依據：新模型進度如 Q2「慢於正常」的延續", "情境 C（AI 效益遞延）"),
  E("e0_A", "情境 A：FY2027 新垂直增量（占前一年營收）", "%", None, 0.16, None, "Assumed", "情境依據：消費廣告主支出年增三位數（Interested-party），自助平台全面開放", "情境 A"),
  E("e0_B", "情境 B：FY2027 新垂直增量（占前一年營收）", "%", None, 0.12, None, "Assumed", "基準；無電商營收絕對值（未找到），為承重 Assumed", "情境 B"),
  E("e0_C", "情境 C：FY2027 新垂直增量（占前一年營收）", "%", None, 0.05, None, "Assumed", "情境依據：消費廣告主放緩、增量維持小", "情境 C"),
  E("cdiv_A", "情境 A：FY2032 累計分流（AI 聊天介面分走安裝量）", "%", None, 0.01, None, "Assumed", "情境依據：分流有限；ChatGPT 廣告 2026-08 年化營收約 10 億美元、仍以零售與金融為主（Sensor Tower 經 PPC Land 轉述，Interested-party）", "情境 A"),
  E("cdiv_B", "情境 B：FY2032 累計分流", "%", None, 0.04, None, "Assumed", "基準：分流自 FY2027 起線性爬升；OpenAI 已接 AppsFlyer 做安裝歸因（二手報導）", "情境 B"),
  E("cdiv_C", "情境 C：FY2032 累計分流", "%", None, 0.10, None, "Assumed", "情境依據：聊天介面成為應用程式探索入口，吃掉行動遊戲安裝預算", "情境 C"),
  E("dc_26", "資料中心成本率（2026 年前三季；占營收）", "%", None, "={dc9m}/{rev9m}", None, "Derived", "10-Q Note 11：9M26 資料中心成本 537.843 ÷ 營收 5,950.950", "Axon 推薦系統訓練與服務的運算成本（雲端租用）；公司指引含較高訓練運算成本"),
  E("dcr_A", "情境 A：FY2029 起資料中心成本率", "%", None, 0.085, None, "Assumed", "情境依據：運算效率提升；Q3 為 8.9%（10-Q 衍生）", "情境 A"),
  E("dcr_B", "情境 B：FY2029 起資料中心成本率", "%", None, 0.09, None, "Assumed", "基準＝近四期水準（9M26 9.0%）持平", "情境 B"),
  E("dcr_C", "情境 C：FY2029 起資料中心成本率", "%", None, 0.11, None, "Assumed", "情境依據：模型世代更替使訓練運算占比上升（FY2023 13.6%、FY2024 12.2% 為歷史高點）", "情境 C"),
  E("llm_tasks27", "LLM 生成互動式廣告素材任務數（FY2027；百萬件）", "百萬件", 20, 200, 2000, "Assumed", "公司 10-Q 2026Q1 起載「AI 生成互動式廣告與影片素材」工具；任務數未揭露（未找到），判斷值", "推論成本只讀 Tokenomics：任務 token ＝一般聊天 token × IF_TaskTokMult"),
  E("llm_g", "LLM 素材任務數年成長倍數", "x", 1.5, 2.0, 3.0, "Assumed", "判斷"),
  E("tk_chat", "一般聊天每次嘗試總 token（新鮮輸入＋快取輸入＋decode）", "token", None, "={tk_fresh}+{tk_cached}+{tk_dec}", None, "Derived", TKSNAP + " IF_TaskTokFresh＋IF_TaskTokCached＋IF_TaskTokDec（第一任務類）", "契約：每任務 token 數＝IF_TaskTokMult × 一般聊天 token 數"),
  E("tk_fresh", "Tokenomics IF_TaskTokFresh（一般聊天）", "token", None, 1500, None, "Derived", TKSNAP),
  E("tk_cached", "Tokenomics IF_TaskTokCached（一般聊天）", "token", None, 1000, None, "Derived", TKSNAP),
  E("tk_dec", "Tokenomics IF_TaskTokDec（一般聊天）", "token", None, 500, None, "Derived", TKSNAP),
  E("tk_mult", "素材生成任務相對一般聊天的 token 倍數（取 IF_TaskTokMult 欄 0.5 小時任務）", "x", 2.0666666666666667, 9.2, 36.8, "Assumed", TKSNAP + " IF_TaskTokMult［1、2.067、9.2、36.8、498］", "任務類別選取為判斷（Assumed）；低＝0.1 小時任務、高＝2 小時任務"),
  E("tk_px", "軟體公司向模型商購買的市場價（IF_PriceRef_Sol；有效牌價 $/M token）", "US$/百萬 token", 0.784, 1.568, 8.0355555555555, "Derived", TKSNAP + " IF_PriceRef_Sol（Interface!C108）；低＝IF_PriceRef_Terra、高＝IF_PriceRef_Astra", "市場價欄；全成本 IF_FullCost_* 只作模型商自做成本下限的對照（Recon），不進本欄"),
  E("oc_r", "其他營業成本率（營業成本扣資料中心；含無形資產攤銷、付款處理）", "%", 0.018, "=({cogs9m}-{dc9m})/{rev9m}", 0.028, "Derived", "2026 年前三季（670.198 − 537.843）÷ 5,950.950", ""),
  E("sm_r", "銷售與行銷費用率", "%", 0.028, "={sm9m}/{rev9m}", 0.045, "Derived", "2026 年前三季 192.948 ÷ 5,950.950", "Q3 為 3.1%；新垂直與自助平台推廣使費用率有上行風險"),
  E("rd_r", "研發費用率", "%", 0.04, "={rd9m}/{rev9m}", 0.07, "Derived", "2026 年前三季 295.822 ÷ 5,950.950", "Q3 研發 +132% 主要為股票薪酬相關人事費（10-Q MD&A）；上行風險"),
  E("ga_r", "一般管理費用率", "%", 0.02, "={ga9m}/{rev9m}", 0.03, "Derived", "2026 年前三季 148.804 ÷ 5,950.950", ""),
  E("sbc_r", "股票薪酬率（含於各費用內；只影響非 GAAP 與現金流加回）", "%", 0.035, "={sbc9m}/{rev9m}", 0.055, "Derived", "2026 年前三季 253.417 ÷ 5,950.950（Adjusted EBITDA 調節表）", "DCF 不加回（sbc_add＝0）"),
  E("da_fix", "折舊攤銷與註銷（年）", "US$M", 110.0, 135.0, 160.0, "Assumed", "2026 年前三季 100.418（年化約 134）；其中無形資產攤銷約 97（FY2025），為收購遺留，預測期持平", ""),
  E("capex_r", "資本支出率（不動產設備、無形資產購置、融資租賃本金；占營收）", "%", 0.004, 0.008, 0.015, "Assumed", "FY2025：(0.473＋28.318＋18.669) ÷ 5,480.717＝0.87%；2026 年前三季約 0.8%（Derived）", "運算以租用為主，資本密集度低；Q3 資本支出 7.5 為近年高點"),
  E("wc_k", "營運資金增加 ÷ 營收增量（年度）", "x", 0.05, 0.12, 0.25, "Assumed", "2026Q3 淨營運資金（應收 2,388.8 − 應付 838.2 − 應計 661.2）＝890.4，占年化 Q4 營收約 10–12%（Derived）；FY2025 流入、2026 年前三季流出 584（現金流量表）", "DCF 標準規格第 1 節與原則第 10 節第 23 條：以營收增量比率建模"),
  E("wc_q", "2026Q4 營運資金增加 ÷ 季營收季增（Q3 淨營運資金 ÷ Q3 營收）", "x", None, "=({ar}-{ap}-{accr})/{rev_q3}", None, "Derived", SRC_Q3 + " 資產負債表與損益表", ""),
)
S("E. 估值（同業前瞻本益比、權重）",
  E("pe_pins", "同業預估本益比：PINS（2026-11-06）", "x", None, 8.89, None, "Interested-party", "https://stockanalysis.com/stocks/pins/statistics/", "口徑：分母為賣方共識 EPS（多為 non-GAAP，排除股票薪酬）、期間未載；與 APP 的 GAAP EPS 口徑不同（原則第 10 節第 30 條）", IP_SELL),
  E("pe_ttd", "同業預估本益比：TTD（2026-11-06）", "x", None, 13.87, None, "Interested-party", "https://stockanalysis.com/stocks/ttd/statistics/", "同上", IP_SELL),
  E("pe_rddt", "同業預估本益比：RDDT（2026-11-13）", "x", None, 15.68, None, "Interested-party", "https://stockanalysis.com/stocks/rddt/statistics/", "同上", IP_SELL),
  E("pe_u", "同業預估本益比：Unity（2026-10-26；逾期，早於 Q3 財報後急漲）", "x", None, 27.06, None, "Interested-party", "https://stockanalysis.com/stocks/u/statistics/", "逾期：與現價日相差超過 2 個交易日", IP_SELL),
  E("pe_meta", "同業預估本益比：META（2026-11-13）", "x", None, 19.48, None, "Interested-party", "https://stockanalysis.com/stocks/meta/statistics/", "同上", IP_SELL),
  E("pe_googl", "同業預估本益比：GOOGL（2026-11-13）", "x", None, 21.45, None, "Interested-party", "https://stockanalysis.com/stocks/googl/statistics/", "同上", IP_SELL),
  E("pe_app", "APP 自身預估本益比（2026-11-06，391.22 價位）", "x", None, 19.87, None, "Interested-party", "https://stockanalysis.com/stocks/app/statistics/", "只作對照，不進模型；不計入同業中位數", IP_SELL),
  E("pe_rel", "相對倍數調整係數（1.00＝未調整）", "x", 0.7, 1.0, 1.3, "Assumed", "未調整：缺同業 NTM EPS 同一日期、同一 GAAP／non-GAAP 口徑資料（原則第 10 節第 18 條；承重輸入）", "寫「1.00（未調整：缺…）」；Sens 低／高＝倍數 ±30%"),
  E("w_dcf", "目標價中 DCF 的權重（其餘＝同業前瞻本益比）", "%", 0.3, 0.5, 0.7, "Assumed", "原則第 4 節預設 DCF 50%＋Comparable 50%"),
  E("tp_wA", "情境權重 A", "%", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25"),
  E("tp_wB", "情境權重 B", "%", None, 0.5, None, "Assumed", "同上"),
  E("tp_wC", "情境權重 C", "%", None, 0.25, None, "Assumed", "同上"),
)
# 由 Hist 與其他頁讀取的衍生來源（以 token 表示，建置時換成儲存格參照）：dc9m、rev9m、cogs9m、sm9m、rd9m、ga9m、sbc9m、ar、ap、accr、rev_q3
IN_ORDER = [e for _, items in IN_SECTIONS for e in items]
INKEYS = [e["key"] for e in IN_ORDER]

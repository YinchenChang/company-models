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
  E("beta", "β（Blume 調整後兩個來源的中位數）", "x", 2.005, "=MEDIAN({blume}*{beta_a}+(1-{blume}),{blume}*{beta_b}+(1-{blume}))", 2.0452, "Derived", "同 Meta v0.2.1 F4 做法", "低／高＝兩個來源各自的調整值；原始 β 約 2.5 使股權成本約 15%，是 DCF 的主導輸入"),
  E("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節預設 5%（敏感度 4–6%）"),
  E("int_q3", "2026Q3 利息費用（單季）", "US$M", None, 51.066, None, "Verified", SRC_Q3 + " Statements of Operations"),
  E("kd", "稅前債務成本＝2026Q3 利息費用年化 ÷ 票據本金（有效利率）", "%", None, "=4*{int_q3}/{debt_0}", None, "Derived", "DCF 標準規格第 2 節：公司債有效利率", "10-K 票面利率區間未逐檔核對，不用票面"),
  E("wacc_adj", "WACC 加碼（敏感度用；基準 0）", "%", -0.01, 0.0, 0.01, "Assumed", "DCF 標準規格第 4 節：WACC ±1pt"),
  E("g_tv", "永續成長率（FY2032 後）", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格第 3 節預設 3%（敏感度 2–4%）"),
  E("ronic_sp", "永續期 RONIC 相對 WACC 的差", "%", None, 0.0, 0.05, "Assumed", "DCF 標準規格第 3 節：RONIC＝WACC（價值中性）", "Sens：RONIC＝WACC＋5pt"),
  E("tv_alt", "終值方法（0＝NOPAT×(1−g÷RONIC)；1＝FCFF×(1＋g) 替代法）", "x", None, 0, 1, "Assumed", "DCF 標準規格第 4 節：終值改用 FCFF×(1＋g) 的替代法", "只作敏感度"),
  E("sbc_add", "DCF 是否加回股票薪酬（0＝不加回，預設；1＝敏感度）", "x", None, 0, 1, "Assumed", "DCF 標準規格第 1 節：股票薪酬不加回（預設；加回只作敏感度）"),
  E("mid_c", "期中慣例：年度現金流距年初", "年", None, 0.5, None, "Derived", "DCF 標準規格第 1 節"),
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
  E("u0_A", "情境 A：FY2027 AI 提升單次安裝淨收入", "%", None, 0.12, None, "Assumed", "情境依據：新模型全面上線（公司說明目前部分上線）；公司長期複合成長約 30% 的口頭說法為上緣參照（Interested-party）", "情境 A（AI 效益充分釋放）"),
  E("u0_B", "情境 B：FY2027 AI 提升單次安裝淨收入", "%", None, 0.08, None, "Assumed", "基準：由本模型參數推得的 FY2027 營收年增與賣方 FY2027 共識（財報前）同量級對照見 Recon，不以共識校準", "情境 B（基準）"),
  E("u0_C", "情境 C：FY2027 AI 提升單次安裝淨收入", "%", None, 0.03, None, "Assumed", "情境依據：新模型進度如 Q2「慢於正常」的延續", "情境 C（AI 效益遞延）"),
  E("e0_A", "情境 A：FY2027 新垂直增量（占前一年營收）", "%", None, 0.14, None, "Assumed", "情境依據：消費廣告主支出年增三位數（Interested-party），自助平台全面開放", "情境 A"),
  E("e0_B", "情境 B：FY2027 新垂直增量（占前一年營收）", "%", None, 0.10, None, "Assumed", "基準；無電商營收絕對值（未找到），為承重 Assumed", "情境 B"),
  E("e0_C", "情境 C：FY2027 新垂直增量（占前一年營收）", "%", None, 0.04, None, "Assumed", "情境依據：消費廣告主放緩、增量維持小", "情境 C"),
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
  E("tk_mult", "素材生成任務相對一般聊天的 token 倍數（取 IF_TaskTokMult 欄 0.5 小時任務）", "x", 2.0666666666666667, 9.2, 36.8, "Assumed", TKSNAP + " IF_TaskTokMult［1、2.067、9.2、36.8、498］", "公司未揭露 token 量或任務量（未找到；G2-1 無揭露可校準），故用快照倍數，任務類別選取為判斷；低＝0.1 小時任務、高＝2 小時任務"),
  E("tk_px", "軟體公司向模型商購買的市場價（IF_PriceRef_Sol；有效牌價 $/M token）", "US$/百萬 token", 0.784, 1.568, 8.0355555555555, "Analogy", TKSNAP + " IF_PriceRef_Sol；低＝IF_PriceRef_Terra、高＝IF_PriceRef_Astra", "選哪一層 token 價格屬類比（G2-1），不標 Verified；高案只動成本一邊，標「上界」，聯動案另列 Sens"),
  E("oc_r", "其他營業成本率（營業成本扣資料中心；含無形資產攤銷、付款處理）", "%", 0.018, "=({cogs9m}-{dc9m})/{rev9m}", 0.028, "Derived", "2026 年前三季（670.198 − 537.843）÷ 5,950.950", ""),
  E("sm_r", "銷售與行銷費用率", "%", 0.028, "={sm_ttm}/{rev_ttm}", 0.045, "Derived", "最近四季（2025Q4＋2026 年前三季）：（48.776＋192.948）÷（1,657.944＋5,950.950）（Derived）", "G2-3：基準全期持平於最近四季實績；收斂只放 Sens。Q3 為 3.1%；新垂直與自助平台推廣使費用率有上行風險"),
  E("rd_r", "研發費用率", "%", 0.04, "={rd_ttm}/{rev_ttm}", 0.07, "Derived", "最近四季：（82.220＋295.822）÷（1,657.944＋5,950.950）（Derived）", "G2-3 同上。Q3 研發 +132% 主要為股票薪酬相關人事費（10-Q MD&A）；上行風險"),
  E("ga_r", "一般管理費用率", "%", 0.02, "={ga_ttm}/{rev_ttm}", 0.03, "Derived", "最近四季：（68.176＋148.804）÷（1,657.944＋5,950.950）（Derived）", "G2-3 同上"),
  E("sbc_r", "股票薪酬率（含於各費用內；只影響非 GAAP 與現金流加回）", "%", 0.035, "={sbc9m}/{rev9m}", 0.055, "Derived", "2026 年前三季 253.417 ÷ 5,950.950（Adjusted EBITDA 調節表）", "DCF 不加回（sbc_add＝0）"),
  E("da_fix", "折舊攤銷與註銷（年）", "US$M", 110.0, 135.0, 160.0, "Assumed", "2026 年前三季 100.418（年化約 134）；其中無形資產攤銷約 97（FY2025），為收購遺留，預測期持平", ""),
  E("capex_r", "資本支出率（不動產設備、無形資產購置、融資租賃本金；占營收）", "%", 0.004, 0.008, 0.015, "Assumed", "FY2025：(0.473＋28.318＋18.669) ÷ 5,480.717＝0.87%；2026 年前三季約 0.8%（Derived）", "運算以租用為主，資本密集度低；Q3 資本支出 7.5 為近年高點"),
  E("wc_k", "營運資金增加 ÷ 營收增量（年度）", "x", 0.05, 0.12, 0.25, "Assumed", "2026Q3 淨營運資金（應收 2,388.8 − 應付 838.2 − 應計 661.2）＝890.4，占年化 Q4 營收約 10–12%（Derived）；FY2025 流入、2026 年前三季流出 584（現金流量表）", "DCF 標準規格第 1 節與原則第 10 節第 23 條：以營收增量比率建模"),
  E("wc_q", "2026Q4 營運資金增加 ÷ 季營收季增（Q3 淨營運資金 ÷ Q3 營收）", "x", None, "=(bs_ar-bs_ap-bs_accr)/rev_q3", None, "Derived", SRC_Q3 + " 資產負債表與損益表", ""),
)
S("E. 估值（同業前瞻本益比、權重）",
  E("pe_pins", "同業預估本益比：PINS（2026-11-06）", "x", None, 8.89, None, "Interested-party", "https://stockanalysis.com/stocks/pins/statistics/", "逾期：2026-11-06 與現價日相差超過 2 個交易日。口徑：分母為賣方共識 EPS（多為 non-GAAP，排除股票薪酬）、期間未載；與 APP 的 GAAP EPS 口徑不同（原則第 10 節第 30 條）", IP_SELL),
  E("pe_ttd", "同業預估本益比：TTD（2026-11-06）", "x", None, 13.87, None, "Interested-party", "https://stockanalysis.com/stocks/ttd/statistics/", "逾期：2026-11-06 與現價日相差超過 2 個交易日；口徑同上", IP_SELL),
  E("pe_rddt", "同業預估本益比：RDDT（2026-11-13）", "x", None, 15.68, None, "Interested-party", "https://stockanalysis.com/stocks/rddt/statistics/", "同上", IP_SELL),
  E("pe_u", "同業預估本益比：Unity（2026-10-26；逾期，早於 Q3 財報後急漲）", "x", None, 27.06, None, "Interested-party", "https://stockanalysis.com/stocks/u/statistics/", "逾期：與現價日相差超過 2 個交易日", IP_SELL),
  E("pe_meta", "同業預估本益比：META（2026-11-13）", "x", None, 19.48, None, "Interested-party", "https://stockanalysis.com/stocks/meta/statistics/", "同上", IP_SELL),
  E("pe_googl", "同業預估本益比：GOOGL（2026-11-13）", "x", None, 21.45, None, "Interested-party", "https://stockanalysis.com/stocks/googl/statistics/", "同上", IP_SELL),
  E("pe_app", "APP 自身預估本益比（2026-11-06，391.22 價位）", "x", None, 19.87, None, "Interested-party", "https://stockanalysis.com/stocks/app/statistics/", "逾期：2026-11-06 與現價日相差超過 2 個交易日。只作對照，不進模型；不計入同業中位數", IP_SELL),
  E("pe_rel", "相對倍數調整係數（1.00＝未調整）", "x", 0.7, 1.0, 1.3, "Assumed", "未調整：缺同業 NTM EPS 同一日期、同一 GAAP／non-GAAP 口徑資料（原則第 10 節第 18 條；承重輸入）", "寫「1.00（未調整：缺…）」；Sens 低／高＝倍數 ±30%"),
  E("w_dcf", "目標價中 DCF 的權重（其餘＝同業前瞻本益比）", "%", 0.3, 0.5, 0.7, "Assumed", "原則第 4 節預設 DCF 50%＋Comparable 50%"),
  E("tp_wA", "情境權重 A", "%", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25"),
  E("tp_wB", "情境權重 B", "%", None, 0.5, None, "Assumed", "同上"),
  E("tp_wC", "情境權重 C", "%", None, 0.25, None, "Assumed", "同上"),
)
# 由 Hist 與其他頁讀取的衍生來源（以 token 表示，建置時換成儲存格參照）：dc9m、rev9m、cogs9m、sm9m、rd9m、ga9m、sbc9m、ar、ap、accr、rev_q3
IN_ORDER = [e for _, items in IN_SECTIONS for e in items]
INKEYS = [e["key"] for e in IN_ORDER]


# ═════════════════════════════ 參數與基本函式 ═════════════════════════════
import argparse, subprocess
from openpyxl.workbook.defined_name import DefinedName
ap = argparse.ArgumentParser()
ap.add_argument("out", nargs="?")
ap.add_argument("--sens"); ap.add_argument("--kreq")
ap.add_argument("--set", action="append", default=[])
ap.add_argument("--cases", action="store_true", help="輸出敏感度案例 JSON 後結束")
A = ap.parse_args()
OVR = {}
for _s in A.set:
    _k, _v = _s.split("=", 1); OVR[_k] = float(_v)

IND = {e["key"]: e for e in IN_ORDER}


def _tok(x):
    return re.sub(r"\{(\w+)\}", r"\1", x) if isinstance(x, str) else x


# ── 敏感度案例（In 頁每個有低／高區間的鍵一案；G2-2 聯動案另列）──
def make_cases():
    cases = []
    for e in IN_ORDER:
        for side, key in (("低", "lo"), ("高", "hi")):
            v = e[key]
            if v is None or isinstance(v, str): continue
            cases.append(dict(id=f"{e['key']}|{side}", keys=[e["key"]], side=side, label=e["label"].split("（")[0][:30], sets=[[e["key"], v]]))
    tk = IND
    def tv(k, s): return tk[k]["lo" if s == "低" else "hi"]
    cases.append(dict(id="tkjoint|高", keys=["tk_px", "tk_mult", "llm_tasks27"], side="高", label="成本聯動上界：價格、倍數與任務數同取高值",
                      sets=[["tk_px", tv("tk_px", "高")], ["tk_mult", tv("tk_mult", "高")], ["llm_tasks27", tv("llm_tasks27", "高")]]))
    cases.append(dict(id="tkroute|高", keys=["tk_px", "tk_mult"], side="高", label="路由到便宜模型：價格取低層、任務 token 取高倍數",
                      sets=[["tk_px", tv("tk_px", "低")], ["tk_mult", tv("tk_mult", "高")]]))
    return cases


if A.cases:
    print(json.dumps(make_cases(), ensure_ascii=False)); sys.exit(0)

# ── build_id ──
def _build_id():
    h = hashlib.sha256()
    fl = ["build_applovin.py"] + sorted(os.path.join("data", f) for f in os.listdir(os.path.join(HERE, "data")) if not f.startswith(("sens_", "kreq_")))
    for fn in fl:
        h.update(fn.encode()); h.update(open(os.path.join(HERE, fn), "rb").read())
    return f"applovin-v{VER}-{h.hexdigest()[:10]}"


BUILD_ID = _build_id()
TK = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot.json"), encoding="utf-8"))
SENS = json.load(open(A.sens, encoding="utf-8")) if A.sens and os.path.exists(A.sens) else {}
KREQ = json.load(open(A.kreq, encoding="utf-8")) if A.kreq and os.path.exists(A.kreq) else {}
wb = openpyxl.Workbook()
RD = wb.active; RD.title = "README"
IN = wb.create_sheet("In"); HI = wb.create_sheet("Hist"); TH = wb.create_sheet("Thesis"); PL = wb.create_sheet("PL"); CF = wb.create_sheet("CF")
V = wb.create_sheet("Val"); RV = wb.create_sheet("RevDCF"); RC = wb.create_sheet("Recon"); SE = wb.create_sheet("Sens"); SO = wb.create_sheet("Sources")
ROWMAP = {}


def dn(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def hdr(ws, row, labels, c0=1):
    for j, h in enumerate(labels, c0):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = SFILL


def fx(*parts):
    """公式組字：str＝文字常數（每段 ≤ 200 字），tuple(("f", 運算式))＝公式片段。"""
    out = []
    for p in parts:
        if isinstance(p, tuple):
            out.append(p[1])
        else:
            for i in range(0, len(p), 200):
                out.append('"' + p[i:i + 200].replace('"', '""') + '"')
    return "=" + "&".join(out)


def T(expr, f="0.00"):
    return ("f", f'TEXT({expr},"{f}")')


def put(ws, r, c, v, fmt=None, blue=False, bold=False):
    cell = ws.cell(r, c, v)
    if fmt: cell.number_format = fmt
    if blue or (isinstance(v, (int, float)) and not isinstance(v, bool)): cell.font = BLUE
    if bold: cell.font = BOLD
    return cell


# ═════════════════════════════ Hist ═════════════════════════════
PER = {"C": "FY2023", "D": "FY2024", "E": "FY2025", "F": "2025 年前三季累計（9M25）", "G": "2026 年前三季累計（9M26）", "H": "2025Q4", "I": "2026Q3"}


def fget(prefix, per):
    for (m, p), r in FACTS.items():
        if m.startswith(prefix) and p == per: return r
    return None


HI["A1"] = "Hist — 歷史實績（持續營運單位；US$M，EPS 為美元，股數為百萬股）。數字取自 data/facts_applovin.csv（事實總帳；逐期出處、擷取日、標記見該檔），本頁只讀、不重估。FY2023 無稀釋股數與 EPS（未找到）。"
HI["A2"] = "L 欄＝最近四季（2025Q4＋2026 年前三季）。Adjusted EBITDA、自由現金流為公司定義的非 GAAP 指標（Interested-party）。"
hdr(HI, 3, ["項目", "單位", "FY2023", "FY2024", "FY2025", "9M25", "9M26", "2025Q4", "2026Q3", "標記", "來源", "最近四季", "來源誘因（Interested-party 列）"])
HR = {}
HROWS = [("rev", "營收（持續營運單位）", "US$M", "合併營收（持續營運單位）", "營收（持續營運單位）"),
         ("cogs", "營業成本", "US$M", "營業成本（Cost"), ("dc", "其中：資料中心成本", "US$M", "資料中心成本"),
         ("sm", "銷售與行銷費用", "US$M", "銷售與行銷費用"), ("rd", "研發費用", "US$M", "研發費用"), ("ga", "一般管理費用", "US$M", "一般管理費用"),
         ("oi", "營業利益", "US$M", "營業利益（Income"), ("int", "利息費用（含債務清償損失）", "US$M", "利息費用"), ("oth", "其他收益（費損）淨額", "US$M", "其他收益（費損）"),
         ("pti", "稅前利益", "US$M", "稅前利益"), ("tax", "所得稅費用", "US$M", "所得稅費用"), ("ni", "持續營運稅後淨利", "US$M", "持續營運稅後淨利"),
         ("eps", "稀釋 EPS（持續營運）", "US$", "稀釋 EPS"), ("shd", "稀釋加權平均股數", "百萬股", "稀釋加權平均股數"),
         ("sbc", "股票薪酬（SBC）", "US$M", "股票薪酬"), ("da", "折舊攤銷與註銷", "US$M", "折舊攤銷與註銷"), ("aebitda", "Adjusted EBITDA（公司定義）", "US$M", "Adjusted EBITDA")]
r = 4
for row in HROWS:
    key, lab, unit, pre = row[0], row[1], row[2], row[3]
    HR[key] = r; HI.cell(r, 1, lab); HI.cell(r, 2, unit)
    tags, srcs = set(), []
    for col, per in PER.items():
        f = None
        if key == "rev":
            f = fget("合併營收（持續營運單位）", per) if col in "CDE" else fget("營收（持續營運單位）", per)
        else:
            f = fget(pre, per)
        if f:
            put(HI, r, "ABCDEFGHI".index(col) + 1, float(f["數值"]), "#,##0.000" if key not in ("eps",) else "0.00")
            tags.add(f["標記"]); srcs.append(f["出處"])
    HI.cell(r, 10, "Interested-party" if "Interested-party" in tags else "Verified")
    HI.cell(r, 11, (srcs[0][:140] if srcs else "未找到") + "（逐期出處見 data/facts_applovin.csv）")
    if "Interested-party" in tags: HI.cell(r, 13, IP_CO)
    r += 1
# 推得值（FY2023 稅前、其他收益）
HI.cell(HR["pti"], 3, f"=C{HR['ni']}+C{HR['tax']}").number_format = "#,##0.000"
HI.cell(HR["oth"], 3, f"=C{HR['pti']}-C{HR['oi']}+C{HR['int']}").number_format = "#,##0.000"
HI.cell(HR["pti"], 10, "Verified"); HI.cell(HR["oth"], 10, "Derived")
HI.cell(HR["oth"], 11, "FY2023 稅前利益＝淨利＋所得稅；其他收益＝稅前利益−營業利益＋利息費用（推得；10-K FY2023 損益表未逐項取得）")
for k in ("rev", "cogs", "dc", "sm", "rd", "ga", "oi", "int", "oth", "pti", "tax", "ni", "sbc", "da", "aebitda"):
    HI.cell(HR[k], 12, f"=IF(AND(ISNUMBER(G{HR[k]}),ISNUMBER(H{HR[k]})),G{HR[k]}+H{HR[k]},\"未找到\")").number_format = "#,##0.000"
HI.cell(HR["dc"], 12, "未找到（2025Q4 單季資料中心成本未取得）")
HI.cell(HR["sbc"], 12, "未找到（2025Q4 單季股票薪酬未取得）")
r += 1
HI.cell(r, 1, "資產負債表（2026-09-30）"); HI.cell(r, 1).font = BOLD; r += 1
for key, lab, pre in (("bs_ar", "應收帳款淨額", "應收帳款淨額"), ("bs_ap", "應付帳款", "應付帳款"), ("bs_accr", "應計及其他流動負債", "應計及其他流動負債")):
    f = fget(pre, "2026-09-30")
    HI.cell(r, 1, lab); HI.cell(r, 2, "US$M"); put(HI, r, 7, float(f["數值"]), "#,##0.000"); HI.cell(r, 10, "Verified"); HI.cell(r, 11, f["出處"][:140]); HR[key] = r; r += 1
HI.column_dimensions["A"].width = 36; HI.column_dimensions["K"].width = 60
for nm_, col, k in (("rev9m", "G", "rev"), ("cogs9m", "G", "cogs"), ("dc9m", "G", "dc"), ("sbc9m", "G", "sbc"), ("rev_q3", "I", "rev"), ("rev_ttm", "L", "rev"),
                    ("sm_ttm", "L", "sm"), ("rd_ttm", "L", "rd"), ("ga_ttm", "L", "ga"), ("bs_ar", "G", "bs_ar"), ("bs_ap", "G", "bs_ap"), ("bs_accr", "G", "bs_accr")):
    dn(nm_, f"Hist!${col}${HR[k]}")

# ═════════════════════════════ In ═════════════════════════════
IN["A1"] = "In — 輸入（藍字；基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁、不進目標價；情境 A／B／C 的驅動見本頁最後一節「情境驅動對照」）。金額 US$M；每股 US$；股數百萬股。標記：Verified／Interested-party／Analogy／Assumed／Derived（Analogy、Assumed 一律給低／高區間）"
IN["A2"] = fx("事實總帳：data/facts_applovin.csv；Tokenomics 快照 data/tokenomics_snapshot.json（v22.0，只讀）。評價日 ", T("val_date", "yyyy-mm-dd"), "，股價基準日 ", T("px_date", "yyyy-mm-dd"),
              "，目標時點 ", T("tp_date", "yyyy-mm-dd"), "（＝股價基準日＋12 個月）。逾期＝與股價基準日相差超過 2 個交易日。未找到項目見 README 與各列說明。")
hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "說明", "來源誘因（Interested-party 列）"])
r = 4
INR = {}
for title, items in IN_SECTIONS:
    IN.cell(r, 1, title).font = BOLD
    for c in range(1, 11): IN.cell(r, c).fill = SFILL
    r += 1
    for e in items:
        INR[e["key"]] = r; r += 1
    # 列號先配好（公式用名稱，不需列號）
r = 4
for title, items in IN_SECTIONS:
    r += 1
    for e in items:
        k = e["key"]; rr = INR[k]
        IN.cell(rr, 1, k); IN.cell(rr, 2, e["label"]); IN.cell(rr, 3, e["unit"])
        for col, v in ((4, e["lo"]), (5, e["base"]), (6, e["hi"])):
            if col == 5 and k in OVR: v = OVR[k]
            v = _tok(v)
            if v is None: continue
            cell = IN.cell(rr, col, v)
            if isinstance(v, dt.datetime): cell.number_format = "yyyy-mm-dd"
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.font = BLUE
                cell.number_format = "0.0%" if e["unit"] == "%" else ("#,##0.000" if abs(v) < 100 else "#,##0.0")
            elif isinstance(v, str) and v.startswith("="):
                cell.number_format = "0.0%" if e["unit"] == "%" else "#,##0.000"
        IN.cell(rr, 7, e["tag"]); IN.cell(rr, 8, _tok(e["src"])); IN.cell(rr, 9, e["note"]); IN.cell(rr, 10, e["inc"] or None)
        if e["tag"] in ("Assumed", "Analogy"): IN.cell(rr, 5).fill = YFILL
        assert not re.fullmatch(r"[A-Za-z]{1,3}\d+", k), k
        dn(k, f"In!$E${rr}")
    r += len(items)
# 讓有 hist 名稱的日期以 yyyy-mm-dd 顯示
for k in ("data_date", "px_date", "val_date", "tp_date"): IN.cell(INR[k], 5).number_format = "yyyy-mm-dd"
IN.cell(INR["dil_inc"], 5).number_format = "0"
r = max(INR.values()) + 2
SCN_SEC = r
IN.cell(r, 1, "J. 情境驅動對照（純展示；三情境並排，進加權目標價的驅動見上方各情境專屬列）").font = BOLD; r += 1
hdr(IN, r, ["驅動", "項目", "單位", "情境 A", "情境 B", "情境 C", "", "說明"]); r += 1
SCN_ROWS = {}
for k, lab, unit, a, b, c in (("scn_u", "FY2027 AI 提升單次安裝淨收入", "%", "u0_A", "u0_B", "u0_C"), ("scn_e", "FY2027 新垂直增量（占前一年營收）", "%", "e0_A", "e0_B", "e0_C"),
                              ("scn_cdiv", "FY2032 累計分流（AI 聊天介面分走安裝量）", "%", "cdiv_A", "cdiv_B", "cdiv_C"), ("scn_dcr", "FY2029 起資料中心成本率", "%", "dcr_A", "dcr_B", "dcr_C"),
                              ("scn_w", "情境權重", "%", "tp_wA", "tp_wB", "tp_wC")):
    IN.cell(r, 1, k); IN.cell(r, 2, lab); IN.cell(r, 3, unit)
    for j, kk in enumerate((a, b, c)):
        cell = IN.cell(r, 4 + j, f"=E{INR[kk]}"); cell.number_format = "0.0%"
    IN.cell(r, 8, f"引用 In 頁 {a}／{b}／{c}")
    SCN_ROWS[k] = r; r += 1
IN.cell(r + 1, 1, "情境分化依據：A＝新模型全面上線、新垂直放大、分流有限、運算效率提升；B＝基準；C＝AI 效益遞延、分流較大、模型世代更替使訓練運算占比上升。AI 淨效果（增量營收−分流侵蝕−推論成本）三情境差距見 Thesis 頁。")
IN_LAST = r
for w, col in ((14, "A"), (46, "B"), (10, "C"), (12, "D"), (12, "E"), (12, "F"), (16, "G"), (50, "H"), (60, "I")): IN.column_dimensions[col].width = w


# ═════════════════════════════ PL／CF／Thesis ═════════════════════════════
SCS = ["B", "A", "C"]                  # 先排基準（標準損益表在最上）
SCN_NAME = {"A": "情境 A（AI 效益充分釋放）", "B": "情境 B（基準）", "C": "情境 C（AI 效益遞延）"}
YCOL = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]
HYR = {"C": "FY2023A", "D": "FY2024A", "E": "FY2025A", "F": "FY2026E", "G": "FY2027E", "H": "FY2028E", "I": "FY2029E", "J": "FY2030E", "K": "FY2031E", "L": "FY2032E"}
FCOL = ["G", "H", "I", "J", "K", "L"]
PREV = {"D": "C", "E": "D", "F": "E", "G": "F", "H": "G", "I": "H", "J": "I", "K": "J", "L": "K"}
BPREV = {"G": "N", "H": "G", "I": "H", "J": "I", "K": "J", "L": "K"}
KIDX = {c: i for i, c in enumerate(FCOL)}
PLK = ["hdr", "s_ads", "rev", "yoy", "chk_seg", "u", "e", "cd", "rev_noai", "rev_nodiv", "dcr", "tasks", "cogs_dc", "cogs_oth", "cogs_llm", "cogs", "gp", "gm", "sm", "rd", "ga", "opx", "oi", "opm",
       "int", "oth", "pti", "tax", "ni", "nm", "niy", "shd", "eps", "da", "sbc", "aeb", "aebm", "chk_oi", "chk_pti", "chk_ni"]
CFK = ["hdr", "ni", "da", "sbc", "dnwc", "cfo", "capex", "fcf", "div", "fcfd", "bb", "pxbb", "shbb", "shend", "shavg", "cashpre", "newdebt", "cash", "debt", "chk_cash",
       "ebit", "nopat", "da2", "sbcadj", "capex2", "dnwc2", "fcff", "tm", "df", "pv"]
PLR, CFR = {}, {}
for i, X in enumerate(SCS):
    PLR[X] = {k: 6 + i * (len(PLK) + 3) + j for j, k in enumerate(PLK)}
    CFR[X] = {k: 5 + i * (len(CFK) + 3) + j for j, k in enumerate(CFK)}
GATE_ROW = 5 + 3 * (len(CFK) + 3) + 1
M9, Q4 = "M", "N"

PL["A1"] = "PL — 三情境損益（US$M；EPS 為美元）。單一官方分部 Advertising（10-K FY2025 起為單一應報導分部；Apps 為停業單位）＝合併營收；FY2023–FY2025 為持續營運單位口徑。損益表為 GAAP；調整後 EBITDA 為 Non-GAAP（本模型口徑）。"
PL["A2"] = "欄位：C–E 實績、F＝FY2026E（＝M 欄 9M26 實績＋N 欄 Q4 26E）、G–L 預測。命題模組（原型 A）：營收＝安裝量指數 × 單次安裝淨收入指數；公司只揭露年增率、不揭露絕對值，故以 FY2026＝基期的年增因子表示（Thesis 頁）。三情境只調 AI 提升、新垂直增量、分流與資料中心成本率（In 頁）。"
PL["A3"] = "標準損益表見情境 B 區塊（第 6 列起）。檢查列應全為 0。"

LBL = dict(hdr="項目", s_ads="Advertising", rev="營收", yoy="營收年增率", chk_seg="檢查：分部合計 − 合併營收", u="AI 提升：單次安裝淨收入年增（模型輸入路徑）", e="新垂直（電商、連網電視）增量（占前一年營收）",
           cd="累計分流（AI 聊天介面分走安裝量）", rev_noai="反事實營收（無 AI 提升、無新垂直、無分流）", rev_nodiv="營收（有 AI 提升與新垂直、無分流）", dcr="資料中心成本率（占營收）",
           tasks="LLM 生成素材任務數", cogs_dc="營業成本：資料中心（Axon 訓練與服務）", cogs_oth="營業成本：其他（無形資產攤銷、付款處理等）", cogs_llm="營業成本：LLM 推論（Tokenomics 快照價）",
           cogs="營業成本", gp="毛利", gm="毛利率", sm="銷售與行銷費用", rd="研發費用", ga="一般管理費用", opx="營業費用", oi="營業利益", opm="營業利益率", int="利息費用",
           oth="其他收入（淨）", pti="稅前利益", tax="所得稅", ni="稅後淨利", nm="淨利率", niy="淨利年增率", shd="稀釋加權平均股數", eps="EPS（GAAP 稀釋）", da="折舊攤銷與註銷",
           sbc="股票薪酬（含於各費用內）", aeb="調整後 EBITDA（Non-GAAP；營業利益＋折舊攤銷＋股票薪酬）", aebm="調整後 EBITDA 利益率", chk_oi="檢查：營業利益 − 財報值", chk_pti="檢查：稅前利益 − 財報值", chk_ni="檢查：稅後淨利 − 財報值")
UNIT = dict(s_ads="US$M", rev="US$M", yoy="%", chk_seg="US$M", u="%", e="%", cd="%", rev_noai="US$M", rev_nodiv="US$M", dcr="%", tasks="百萬件", cogs_dc="US$M", cogs_oth="US$M", cogs_llm="US$M",
            cogs="US$M", gp="US$M", gm="%", sm="US$M", rd="US$M", ga="US$M", opx="US$M", oi="US$M", opm="%", int="US$M", oth="US$M", pti="US$M", tax="US$M", ni="US$M", nm="%", niy="%",
            shd="百萬股", eps="US$", da="US$M", sbc="US$M", aeb="US$M", aebm="%", chk_oi="US$M", chk_pti="US$M", chk_ni="US$M")
PCT = {"yoy", "u", "e", "cd", "dcr", "gm", "opm", "nm", "niy", "aebm"}


def H_(k, col):   # Hist 儲存格；缺值回 None
    return f"Hist!{col}{HR[k]}"


def hist_has(k, col):
    return bool(HI.cell(HR[k], "ABCDEFGHI".index(col) + 1).value) if k in HR else False


for X in SCS:
    R = PLR[X]; CR = CFR[X]
    sfx = "" if X == "B" else f" ｜情境 {X}"
    PL.cell(R["hdr"] - 1, 1, SCN_NAME[X] + ("：標準損益表" if X == "B" else "")).font = BOLD
    hdr(PL, R["hdr"], ["項目", "單位"] + [HYR[c] for c in YCOL] + ["9M26A", "Q4 26E", "說明"])
    for k in PLK[1:]:
        PL.cell(R[k], 1, LBL[k] + (sfx if k not in ("chk_seg",) or X != "B" else "")); PL.cell(R[k], 2, UNIT[k])
        if k == "chk_seg" and X != "B": PL.cell(R[k], 1, f"檢查：分部合計 − 合併營收{sfx}")
    def S_(k, col, f):
        c = PL[f"{col}{R[k]}"]; c.value = f
        c.number_format = "0.0%" if k in PCT else ("0.00" if k == "eps" else "#,##0.0")
        if k in ("tasks",): c.number_format = "#,##0"
    rr = lambda k, c: f"{c}{R[k]}"
    # 歷史與 9M26
    for col, hc in (("C", "C"), ("D", "D"), ("E", "E"), (M9, "G")):
        S_("s_ads", col, f"={H_('rev', hc)}"); S_("rev", col, f"={H_('rev', hc)}")
        S_("cogs_dc", col, f"={H_('dc', hc)}"); S_("cogs", col, f"={H_('cogs', hc)}"); S_("cogs_oth", col, f"={rr('cogs', col)}-{rr('cogs_dc', col)}")
        S_("sm", col, f"={H_('sm', hc)}"); S_("rd", col, f"={H_('rd', hc)}"); S_("ga", col, f"={H_('ga', hc)}")
        S_("int", col, f"={H_('int', hc)}"); S_("oth", col, f"={H_('oth', hc)}"); S_("tax", col, f"={H_('tax', hc)}")
        S_("da", col, f"={H_('da', hc)}")
        if hist_has("sbc", hc): S_("sbc", col, f"={H_('sbc', hc)}")
        if hist_has("eps", hc): S_("eps", col, f"={H_('eps', hc)}")
        if hist_has("shd", hc): S_("shd", col, f"={H_('shd', hc)}")
        S_("dcr", col, f"={rr('cogs_dc', col)}/{rr('rev', col)}")
        S_("chk_oi", col, f"={rr('oi', col)}-{H_('oi', hc)}"); S_("chk_pti", col, f"={rr('pti', col)}-{H_('pti', hc)}"); S_("chk_ni", col, f"={rr('ni', col)}-{H_('ni', hc)}")
    S_("yoy", "D", "=D{0}/C{0}-1".format(R["rev"])); S_("yoy", "E", "=E{0}/D{0}-1".format(R["rev"])); S_("yoy", M9, f"={rr('rev', M9)}/{H_('rev', 'F')}-1")
    S_("niy", "D", "=D{0}/C{0}-1".format(R["ni"])); S_("niy", "E", "=E{0}/D{0}-1".format(R["ni"])); S_("niy", M9, f"={rr('ni', M9)}/{H_('ni', 'F')}-1")
    # Q4 26E
    q = Q4
    S_("s_ads", q, "=rev_q3*(1+q4_qoq)"); S_("rev", q, f"={rr('s_ads', q)}")
    S_("yoy", q, f"={rr('rev', q)}/{H_('rev', 'H')}-1"); S_("niy", q, f"={rr('ni', q)}/{H_('ni', 'H')}-1")
    S_("dcr", q, "=dc_26"); S_("cogs_dc", q, f"={rr('dcr', q)}*{rr('rev', q)}"); S_("cogs_oth", q, f"=oc_r*{rr('rev', q)}"); S_("tasks", q, 0); S_("cogs_llm", q, 0)
    S_("sm", q, f"=sm_r*{rr('rev', q)}"); S_("rd", q, f"=rd_r*{rr('rev', q)}"); S_("ga", q, f"=ga_r*{rr('rev', q)}")
    S_("int", q, "=debt_0*kd/4"); S_("oth", q, "=cash_0*r_cash/4"); S_("tax", q, f"=tax*{rr('pti', q)}")
    S_("da", q, "=da_fix/4"); S_("sbc", q, f"=sbc_r*{rr('rev', q)}")
    S_("shd", q, f"=CF!{q}{CR['shavg']}+dil_base"); S_("eps", q, f"={rr('ni', q)}/{rr('shd', q)}")
    # FY2026E ＝ 9M26A ＋ Q4 26E
    for k in ("s_ads", "rev", "cogs_dc", "cogs_oth", "cogs_llm", "sm", "rd", "ga", "int", "oth", "tax", "da", "sbc"):
        S_(k, "F", f"={rr(k, M9)}+{rr(k, q)}")
    S_("cogs_llm", M9, 0); S_("tasks", M9, 0); S_("tasks", "F", 0); S_("cogs_llm", "F", f"={rr('cogs_llm', q)}")
    S_("dcr", "F", f"={rr('cogs_dc', 'F')}/{rr('rev', 'F')}")
    S_("yoy", "F", "=F{0}/E{0}-1".format(R["rev"])); S_("niy", "F", "=F{0}/E{0}-1".format(R["ni"]))
    S_("shd", "F", f"=0.75*{rr('shd', M9)}+0.25*{rr('shd', q)}"); S_("eps", "F", f"={rr('eps', M9)}+{rr('eps', q)}")
    S_("rev_noai", "F", f"={rr('s_ads', 'F')}"); S_("rev_nodiv", "F", f"={rr('s_ads', 'F')}"); S_("cd", "F", 0)
    # 預測年
    for c in FCOL:
        p = PREV[c]; k = KIDX[c]; bp = BPREV[c]
        S_("u", c, f"=u0_{X}*u_decay^{k}"); S_("e", c, f"=e0_{X}*e_decay^{k}")
        S_("cd", c, f"=cdiv_{X}*MIN({k + 1},div_yrs)/div_yrs")
        S_("s_ads", c, f"={p}{R['s_ads']}*(1+ginst0)*(1+nrpi_base)*(1+{c}{R['u']})*(1+{c}{R['e']})*(1-{c}{R['cd']})/(1-{p}{R['cd']})")
        S_("rev", c, f"={rr('s_ads', c)}")
        S_("yoy", c, f"={c}{R['rev']}/{p}{R['rev']}-1")
        S_("rev_noai", c, f"={p}{R['rev_noai']}*(1+ginst0)*(1+nrpi_base)")
        S_("rev_nodiv", c, f"={p}{R['rev_nodiv']}*(1+ginst0)*(1+nrpi_base)*(1+{c}{R['u']})*(1+{c}{R['e']})")
        S_("dcr", c, {"G": f"=dc_26+(dcr_{X}-dc_26)/3", "H": f"=dc_26+(dcr_{X}-dc_26)*2/3"}.get(c, f"=dcr_{X}"))
        S_("tasks", c, "=llm_tasks27" if c == "G" else f"={p}{R['tasks']}*llm_g")
        S_("cogs_dc", c, f"={rr('dcr', c)}*{rr('rev', c)}"); S_("cogs_oth", c, f"=oc_r*{rr('rev', c)}")
        S_("cogs_llm", c, f"={rr('tasks', c)}*tk_chat*tk_mult*tk_px/1000000")
        S_("sm", c, f"=sm_r*{rr('rev', c)}"); S_("rd", c, f"=rd_r*{rr('rev', c)}"); S_("ga", c, f"=ga_r*{rr('rev', c)}")
        S_("int", c, f"=CF!{bp}{CR['debt']}*kd"); S_("oth", c, f"=CF!{bp}{CR['cash']}*r_cash"); S_("tax", c, f"=tax*{rr('pti', c)}")
        S_("da", c, "=da_fix"); S_("sbc", c, f"=sbc_r*{rr('rev', c)}")
        S_("shd", c, f"=CF!{c}{CR['shavg']}+dil_base"); S_("eps", c, f"={rr('ni', c)}/{rr('shd', c)}")
        S_("niy", c, f"={c}{R['ni']}/{p}{R['ni']}-1")
    # 全欄位共通
    for c in YCOL + [M9, Q4]:
        if c in ("C", "D", "E", "F") or c in FCOL or c in (M9, Q4):
            S_("cogs", c, f"={rr('cogs_dc', c)}+{rr('cogs_oth', c)}+{rr('cogs_llm', c)}") if c not in ("C", "D", "E", M9) else None
            S_("gp", c, f"={rr('rev', c)}-{rr('cogs', c)}"); S_("gm", c, f"={rr('gp', c)}/{rr('rev', c)}")
            S_("opx", c, f"={rr('sm', c)}+{rr('rd', c)}+{rr('ga', c)}"); S_("oi", c, f"={rr('gp', c)}-{rr('opx', c)}"); S_("opm", c, f"={rr('oi', c)}/{rr('rev', c)}")
            S_("pti", c, f"={rr('oi', c)}-{rr('int', c)}+{rr('oth', c)}"); S_("ni", c, f"={rr('pti', c)}-{rr('tax', c)}"); S_("nm", c, f"={rr('ni', c)}/{rr('rev', c)}")
            if c in ("C", "D", "E", M9) and not hist_has("sbc", "CDEG"["CDEM".index(c) if c != M9 else 3]): pass
            if c not in ("C", "D", "E", M9) or hist_has("sbc", {"C": "C", "D": "D", "E": "E", M9: "G"}[c]):
                S_("aeb", c, f"={rr('oi', c)}+{rr('da', c)}+{rr('sbc', c)}"); S_("aebm", c, f"={rr('aeb', c)}/{rr('rev', c)}")
    for c in ("C", "D", "E", M9):   # 實績欄的分部合計核對
        S_("chk_seg", c, f"={rr('s_ads', c)}-{rr('rev', c)}")
    for c in ["F", Q4] + FCOL:
        S_("chk_seg", c, f"={rr('s_ads', c)}-{rr('rev', c)}")
    for c in ("C", "D", "E", M9) :
        pass
    PL.cell(R["s_ads"], 15, "官方單一應報導分部（10-K FY2025）；2025-06-30 出售 Apps（停業單位）")
    PL.cell(R["u"], 15, "FY2027 起：u0×衰減係數^年（In 頁）")
    PL.cell(R["e"], 15, "新垂直增量：e0×衰減係數^年（公司未揭露電商營收，未找到）")
    PL.cell(R["cd"], 15, "線性爬升到 FY2032 期末值；相對前一年的增量折減營收")
    PL.cell(R["cogs_llm"], 15, "任務數 × 每任務 token × 價格（Tokenomics 快照；Analogy）；公司未揭露任務量")
    PL.cell(R["eps"], 15, "歷史欄取財報稀釋 EPS；FY2026E＝9M26 實績＋Q4 26E")
    PL.cell(R["oth"], 15, "預測＝期初現金 × 現金收益率；9M26 含權益法投資損益與匯兌")
    PL.cell(R["aeb"], 15, "Non-GAAP；與公司定義（含其他調整）有差，對照見 Recon")
for w_, col in ((50, "A"), (8, "B")): PL.column_dimensions[col].width = w_
for c in YCOL + ["M", "N"]: PL.column_dimensions[c].width = 11
PL.column_dimensions["O"].width = 70
SUMROW = None

# ── CF ──
CF["A1"] = "CF — 現金流（US$M）。N 欄＝Q4 26E（評價日 2026-09-30 之後未滿一年）、G–L＝FY2027–FY2032。上半：槓桿後自由現金流與股東回饋（股數、現金、債務）；下半：FCFF（NOPAT＋折舊攤銷−資本支出−營運資金增加；股票薪酬預設不加回）與折現。"
CF["A2"] = "自由現金流＝稅後淨利＋折舊攤銷＋股票薪酬−營運資金增加−資本支出（含融資租賃本金）。買回＝自由現金流 × 買回比例；無股利。現金低於最低現金時補新債；利息按期初餘額，無循環。"
CFL = dict(hdr="項目", ni="稅後淨利", da="＋折舊攤銷與註銷", sbc="＋股票薪酬", dnwc="−營運資金增加", cfo="營運現金流", capex="−資本支出", fcf="自由現金流", div="股利", fcfd="自由現金流扣股利",
           bb="庫藏股買回", pxbb="買回價格", shbb="買回股數", shend="期末流通股數", shavg="平均流通股數", cashpre="補債前現金", newdebt="新增舉債", cash="期末現金", debt="期末債務本金",
           chk_cash="檢查：現金平衡", ebit="營業利益（EBIT）", nopat="NOPAT＝EBIT×(1−稅率)", da2="＋折舊攤銷與註銷", sbcadj="＋股票薪酬（稅後；預設不加回）", capex2="−資本支出", dnwc2="−營運資金增加",
           fcff="無槓桿自由現金流（FCFF）", tm="折現期間（距評價日，期中；年）", df="折現因子", pv="FCFF 現值")
for X in SCS:
    R = CFR[X]; P = PLR[X]
    CF.cell(R["hdr"] - 1, 1, SCN_NAME[X]).font = BOLD
    hdr(CF, R["hdr"], ["項目", "單位"] + [HYR[c] for c in YCOL] + ["9M26A", "Q4 26E"])
    for k in CFK[1:]:
        CF.cell(R[k], 1, CFL[k] + ("" if X == "B" else f" ｜情境 {X}")); CF.cell(R[k], 2, "US$M" if k not in ("pxbb", "df", "tm", "shbb", "shend", "shavg") else {"pxbb": "US$", "df": "x", "tm": "年"}.get(k, "百萬股"))
    for c in [Q4] + FCOL:
        bp = BPREV[c] if c != Q4 else None
        ln = "0.25" if c == Q4 else "1"
        def Cc(k, f, fmt="#,##0.0"):
            cell = CF[f"{c}{R[k]}"]; cell.value = f; cell.number_format = fmt
        Cc("ni", f"=PL!{c}{P['ni']}"); Cc("da", f"=PL!{c}{P['da']}"); Cc("sbc", f"=PL!{c}{P['sbc']}")
        if c == Q4: Cc("dnwc", f"=wc_q*(PL!{c}{P['rev']}-rev_q3)")
        else: Cc("dnwc", f"=wc_k*(PL!{c}{P['rev']}-PL!{PREV[c]}{P['rev']})")
        Cc("cfo", f"={c}{R['ni']}+{c}{R['da']}+{c}{R['sbc']}-{c}{R['dnwc']}")
        Cc("capex", f"=capex_r*PL!{c}{P['rev']}"); Cc("fcf", f"={c}{R['cfo']}-{c}{R['capex']}"); Cc("div", 0); Cc("fcfd", f"={c}{R['fcf']}-{c}{R['div']}")
        Cc("bb", f"=bb_pct*MAX({c}{R['fcfd']},0)")
        Cc("pxbb", "=px" if c == Q4 else f"=px*(1+px_g)^{KIDX[c] + 1}", "#,##0.00"); Cc("shbb", f"={c}{R['bb']}/{c}{R['pxbb']}", "#,##0.000")
        Cc("shend", (f"=sh_out+sh_iss*{ln}-{c}{R['shbb']}" if c == Q4 else f"={bp}{R['shend']}+sh_iss-{c}{R['shbb']}"), "#,##0.000")
        Cc("shavg", (f"=(sh_out+{c}{R['shend']})/2" if c == Q4 else f"=({bp}{R['shend']}+{c}{R['shend']})/2"), "#,##0.000")
        Cc("cashpre", (f"=cash_0+{c}{R['fcfd']}-{c}{R['bb']}" if c == Q4 else f"={bp}{R['cash']}+{c}{R['fcfd']}-{c}{R['bb']}"))
        Cc("newdebt", f"=MAX(0,min_cash-{c}{R['cashpre']})"); Cc("cash", f"={c}{R['cashpre']}+{c}{R['newdebt']}")
        Cc("debt", (f"=debt_0+{c}{R['newdebt']}" if c == Q4 else f"={bp}{R['debt']}+{c}{R['newdebt']}"))
        Cc("chk_cash", f"={c}{R['cash']}-({'cash_0' if c == Q4 else bp + str(R['cash'])}+{c}{R['fcfd']}-{c}{R['bb']}+{c}{R['newdebt']})")
        Cc("ebit", f"=PL!{c}{P['oi']}"); Cc("nopat", f"={c}{R['ebit']}*(1-tax)"); Cc("da2", f"={c}{R['da']}"); Cc("sbcadj", f"={c}{R['sbc']}*(1-tax)*sbc_add")
        Cc("capex2", f"={c}{R['capex']}"); Cc("dnwc2", f"={c}{R['dnwc']}")
        Cc("fcff", f"={c}{R['nopat']}+{c}{R['da2']}+{c}{R['sbcadj']}-{c}{R['capex2']}-{c}{R['dnwc2']}")
        Cc("tm", "=0.25/2" if c == Q4 else ("=0.25+mid_c" if c == "G" else f"={bp}{R['tm']}+1"), "0.000")
        Cc("df", f"=1/(1+wacc)^{c}{R['tm']}", "0.0000"); Cc("pv", f"={c}{R['fcff']}*{c}{R['df']}")
CF.cell(GATE_ROW - 1, 1, "收支關卡").font = BOLD
CF.cell(GATE_ROW, 1, "收支關卡：自由現金流扣股利（五年合計）"); CF.cell(GATE_ROW, 2, "情境 A→D 欄依序為 A、B、C 的 FY2027–FY2031 合計")
for j, X in enumerate(("A", "B", "C")):
    cell = CF.cell(GATE_ROW, 2 + j, f"=SUM(G{CFR[X]['fcfd']}:K{CFR[X]['fcfd']})"); cell.number_format = "#,##0.0"
CF.cell(GATE_ROW, 2).number_format = "#,##0.0"
CF.cell(GATE_ROW + 1, 1, "（B、C、D 欄＝情境 A、B、C；收支關卡通過＝合計不小於 0。APP 無股利，扣股利即自由現金流）")
CF.column_dimensions["A"].width = 46
for c in YCOL + ["M", "N"]: CF.column_dimensions[c].width = 11


# ═════════════════════════════ Val ═════════════════════════════
VR = {}
_vr = [3]


def vrow(key, label, vals=None, fmt="#,##0.00", note=None, bold=False):
    r = _vr[0]; _vr[0] += 1; VR[key] = r
    V.cell(r, 1, label)
    if bold: V.cell(r, 1).font = BOLD
    for j, v in enumerate(vals or []):
        if v is None: continue
        c = V.cell(r, 2 + j, v); c.number_format = fmt
        if isinstance(v, (int, float)): c.font = BLUE
    if note: V.cell(r, 5, note)
    return r


def vsec(t):
    r = _vr[0]; _vr[0] += 1
    V.cell(r, 1, t).font = BOLD
    for cc in range(1, 7): V.cell(r, cc).fill = SFILL
    return r


def tierf(u): return f'IF({u}>0.5,"大於 50%",IF({u}>0.25,"25%–50%",IF({u}>=-0.25,"−25%～25%",IF({u}>=-0.5,"−50%～−25%","小於 −50%"))))'
def recf(u): return f'IF({u}>0.5,"強力買進",IF({u}>0.25,"買進",IF({u}>=-0.25,"中立",IF({u}>=-0.5,"放空","強力放空"))))'


V["A1"] = "Val — 估值（US$；股數百萬股）。目標價＝DCF 與 Comparable 加權（預設各半），兩腿各依情境權重加權（25／50／25）。DCF：FCFF 到 FY2032，永續價值 NOPAT×(1+g)×(1−g÷RONIC)÷(WACC−g)，RONIC＝WACC；每股價值以 ke 推到目標時點。Comparable：同業前瞻本益比中位數 × 目標時點下一年（NTM）GAAP EPS。"
V["A2"] = "評等分界：加權目標價相對現價的上檔空間 > 50%＝強力買進、> 25%＝買進、−25%～25%＝中立、−50%～−25%＝放空、< −50%＝強力放空；收支關卡不通過（五年自由現金流扣股利合計 < 0）時，買進以上的評等改為「收支不足」。"
# 目標價主列（先佔位，公式於後填入）
for k, l in (("tp", "目標價（DCF 與 Comparable 加權）"), ("up", "上檔空間：加權目標價"), ("tier", "上檔空間級距"), ("vrate", "估值評等（五級）"),
             ("gate", "入不敷出判定用：收支關卡合計（情境 B）"), ("rec", "投資建議"), ("tpdcf", "DCF 目標價（情境加權；12 個月後）"),
             ("tpcomp", "Comparable 目標價（情境加權；同業前瞻本益比）"), ("cons", "市場共識目標價（只作對照，不進模型）"), ("px", "現價（股價基準日收盤）"), ("upc", "上檔空間：共識目標價")):
    vrow(k, l, bold=k in ("tp", "rec"))
_vr[0] += 1
vsec("折現率與股數")
for k, l in (("rf", "無風險利率 rf"), ("beta", "β（Blume 調整後中位數）"), ("erp", "股權風險溢酬"), ("ke", "股權成本 ke＝rf＋β×ERP"), ("kd", "稅前債務成本 kd"), ("kdat", "稅後債務成本"),
             ("shval", "評價日股數（流通股＋RSU＋選擇權庫藏股法淨股數；百萬股）"), ("emkt", "股權市值（現價 × 評價日股數；US$M）"), ("dmkt", "債務（本金＋融資租賃；US$M）"),
             ("wacc", "WACC"), ("ronic", "RONIC（永續期）"), ("gtv", "永續成長率 g")):
    vrow(k, l)
_vr[0] += 1
VS = vsec("DCF 彙總（三情境；US$M，每股為 US$）")
hdr(V, VS, ["DCF 彙總（三情境）", "情境 A", "情境 B", "情境 C", "說明"])
for k, l in (("pvx", "明確期 FCFF 現值（Q4 26E–FY2032E）"), ("pvtv", "永續價值現值"), ("ev", "企業價值 EV"), ("cash", "＋現金及約當現金"), ("debt", "−優先無擔保票據本金"), ("fl", "−融資租賃負債"),
             ("legal", "−類債：訴訟與和解"), ("eqinv", "＋權益法投資（帳面）"), ("eqv", "股權價值"), ("shs", "評價日股數（百萬股）"), ("ps0", "每股價值（評價日）"),
             ("ps12", "DCF 每股價值（推到目標時點）"), ("tvsh", "永續價值占 EV"), ("tv", "永續價值（FY2032 年底）")):
    vrow(k, l)
_vr[0] += 1
CS = vsec("Comparable（三情境；同業前瞻本益比 × 目標時點 NTM EPS）")
hdr(V, CS, ["Comparable（三情境）", "情境 A", "情境 B", "情境 C", "說明"])
for k, l in (("eps27", "FY2027E EPS（GAAP 稀釋）"), ("eps28", "FY2028E EPS（GAAP 稀釋）"), ("fw", "FY2027 在 NTM 的權重（目標時點起算）"), ("ntm", "目標時點 NTM EPS"),
             ("mult", "採用倍數＝同業中位數 × 相對調整係數"), ("tpc_s", "Comparable 每股價值")):
    vrow(k, l)
_vr[0] += 1
WS = vsec("情境加權")
hdr(V, WS, ["情境加權", "情境 A", "情境 B", "情境 C", "說明"])
for k, l in (("w", "情境權重"), ("tp_s", "情境目標價＝DCF 權重×DCF＋(1−DCF 權重)×Comparable"), ("up_s", "情境上檔空間"), ("wchk", "檢查：情境權重合計 − 1"), ("sc_chk", "檢查：加權 DCF／Comparable 與加權目標價")):
    vrow(k, l)
_vr[0] += 1
MS = vsec("同業前瞻本益比（只作 Comparable 倍數；口徑與日期見說明）")
hdr(V, MS, ["同業", "預估本益比", "資料日", "逾期？（與現價日相差 > 2 個交易日）", "會計準則／租賃處理與分母口徑"])
PEERS = [("PINS", "pe_pins", dt.datetime(2026, 11, 6)), ("TTD", "pe_ttd", dt.datetime(2026, 11, 6)), ("RDDT", "pe_rddt", dt.datetime(2026, 11, 13)),
         ("Unity", "pe_u", dt.datetime(2026, 10, 26)), ("META", "pe_meta", dt.datetime(2026, 11, 13)), ("GOOGL", "pe_googl", dt.datetime(2026, 11, 13))]
PEER_ROWS = []
for nm, key, d in PEERS:
    r = vrow("peer_" + key, nm, [f"={key}", d, f'=IF(NETWORKDAYS(C{_vr[0]},px_date)-1>2,"逾期","")'], "0.00",
             "美國發行人，US GAAP；P/E 不受租賃分類影響；分母為賣方共識 EPS（多為 non-GAAP，排除股票薪酬）、期間未載（Interested-party）")
    V.cell(r, 3).number_format = "yyyy-mm-dd"; PEER_ROWS.append(r)
vrow("mult_med", "同業預估本益比中位數", ["=MEDIAN(" + ",".join(f"B{x}" for x in PEER_ROWS) + ")"], "0.00", "六家中位數；不含 APP 自身")
vrow("app_pe", "APP 自身預估本益比（2026-11-06；只作對照）", ["=pe_app"], "0.00", "逾期：2026-11-06")
vrow("app_ratio", "本公司現行預估本益比 ÷ 同業中位數", ["=B{0}/B{1}".format(VR["app_pe"], VR["mult_med"])], "0.00")
vrow("rerate", "隱含重估：若 APP 倍數回到同業中位數的價格變動", ["=B{0}/B{1}-1".format(VR["mult_med"], VR["app_pe"])], "+0.0%;-0.0%;0.0%", "只作展示；不進目標價")
_vr[0] += 1
vsec("備註")
V.cell(_vr[0], 1, "分母口徑：本公司 EPS 為 GAAP 稀釋（扣股票薪酬）；同業分母多為 non-GAAP。兩者不同，故 Comparable 腿偏保守；相對調整係數基準 1（未調整，缺同業 NTM EPS 同日期同口徑資料）。")

def sc_col(X): return {"A": "B", "B": "C", "C": "D"}[X]
vv = lambda k, c="B": f"{c}{VR[k]}"
Vset = lambda k, c, f, fmt="#,##0.00": (V.cell(VR[k], "ABCDEF".index(c) + 1, f), setattr(V.cell(VR[k], "ABCDEF".index(c) + 1), "number_format", fmt))
Vset("rf", "B", "=rf", "0.00%"); Vset("beta", "B", "=beta", "0.000"); Vset("erp", "B", "=erp", "0.00%"); Vset("ke", "B", "=rf+beta*erp", "0.00%")
Vset("kd", "B", "=kd", "0.00%"); Vset("kdat", "B", "=kd*(1-tax)", "0.00%")
Vset("shval", "B", "=sh_out+rsu_unv*dil_inc+opt_n*MAX(0,1-opt_k/px)", "#,##0.000")
Vset("emkt", "B", f"=px*B{VR['shval']}", "#,##0"); Vset("dmkt", "B", "=debt_0+fin_lease", "#,##0")
Vset("wacc", "B", f"=B{VR['emkt']}/(B{VR['emkt']}+B{VR['dmkt']})*B{VR['ke']}+B{VR['dmkt']}/(B{VR['emkt']}+B{VR['dmkt']})*B{VR['kdat']}+wacc_adj", "0.00%")
Vset("ronic", "B", f"=B{VR['wacc']}+ronic_sp", "0.00%"); Vset("gtv", "B", "=g_tv", "0.00%")
for nm_, k in (("wacc", "wacc"), ("ke", "ke"), ("shval", "shval")): dn(nm_, f"Val!$B${VR[k]}")
V.cell(VR["shval"], 5, "RSU 為 2025-12-31 未歸屬數（Q3 未揭露新值）；沒收率未揭露、全額計入；選擇權履約價 6.10 以庫藏股法")
V.cell(VR["beta"], 5, "兩來源 Blume 調整後的中位數；β 約 2 使 ke 約 15%，是 DCF 的主導輸入（見 Recon 反解）")
for X in ("A", "B", "C"):
    c = sc_col(X); Pr = PLR[X]; Cr = CFR[X]
    cols = ["N"] + FCOL
    Vset("pvx", c, f"=SUM(CF!N{Cr['pv']},CF!G{Cr['pv']}:L{Cr['pv']})", "#,##0")
    Vset("tv", c, f"=IF(tv_alt=1,CF!L{Cr['fcff']}*(1+g_tv),CF!L{Cr['nopat']}*(1+g_tv)*(1-g_tv/B{VR['ronic']}))/(B{VR['wacc']}-g_tv)", "#,##0")
    Vset("pvtv", c, f"={c}{VR['tv']}/(1+B{VR['wacc']})^(CF!L{Cr['tm']}+0.5)", "#,##0")
    Vset("ev", c, f"={c}{VR['pvx']}+{c}{VR['pvtv']}", "#,##0")
    Vset("cash", c, "=cash_0", "#,##0"); Vset("debt", c, "=-debt_0", "#,##0"); Vset("fl", c, "=-fin_lease", "#,##0"); Vset("legal", c, "=-legal_oth", "#,##0"); Vset("eqinv", c, "=eq_inv", "#,##0")
    Vset("eqv", c, f"=SUM({c}{VR['ev']},{c}{VR['cash']}:{c}{VR['legal']},{c}{VR['eqinv']})", "#,##0")
    Vset("shs", c, f"=B{VR['shval']}", "#,##0.000"); Vset("ps0", c, f"={c}{VR['eqv']}/{c}{VR['shs']}")
    Vset("ps12", c, f"={c}{VR['ps0']}*(1+B{VR['ke']})^((tp_date-val_date)/365)"); Vset("tvsh", c, f"={c}{VR['pvtv']}/{c}{VR['ev']}", "0.0%")
    Vset("eps27", c, f"=PL!G{Pr['eps']}"); Vset("eps28", c, f"=PL!H{Pr['eps']}")
    Vset("fw", c, "=MAX(0,MIN(1,(DATE(2027,12,31)-tp_date)/365))", "0.000"); Vset("ntm", c, f"={c}{VR['fw']}*{c}{VR['eps27']}+(1-{c}{VR['fw']})*{c}{VR['eps28']}")
    Vset("mult", c, f"=B{VR['mult_med']}*pe_rel", "0.00"); Vset("tpc_s", c, f"={c}{VR['mult']}*{c}{VR['ntm']}")
    Vset("w", c, f"=tp_w{X}", "0%"); Vset("tp_s", c, f"=w_dcf*{c}{VR['ps12']}+(1-w_dcf)*{c}{VR['tpc_s']}"); Vset("up_s", c, f"={c}{VR['tp_s']}/px-1", "+0.0%;-0.0%;0.0%")
V.cell(VR["tv"], 5, "RONIC＝WACC＋In 頁差；終值＝NOPAT(FY2033)×(1−g÷RONIC)÷(WACC−g)；tv_alt＝1 時改 FCFF×(1+g)")
V.cell(VR["pvtv"], 5, "折現期間＝FY2032 年底（距評價日 6.25 年）")
V.cell(VR["ps12"], 5, "每股價值（評價日）以 ke 推到目標時點；無股利")
Vset("wchk", "B", f"=SUM(B{VR['w']}:D{VR['w']})-1", "0.000000")
Vset("tpdcf", "B", f"=SUMPRODUCT(B{VR['w']}:D{VR['w']},B{VR['ps12']}:D{VR['ps12']})")
Vset("tpcomp", "B", f"=SUMPRODUCT(B{VR['w']}:D{VR['w']},B{VR['tpc_s']}:D{VR['tpc_s']})")
Vset("tp", "B", f"=SUMPRODUCT(B{VR['w']}:D{VR['w']},B{VR['tp_s']}:D{VR['tp_s']})")
Vset("sc_chk", "B", f"=B{VR['tp']}-(w_dcf*B{VR['tpdcf']}+(1-w_dcf)*B{VR['tpcomp']})", "0.000000")
Vset("px", "B", "=px", "#,##0.00"); V.cell(VR["px"], 3, "=px_date").number_format = "yyyy-mm-dd"
Vset("cons", "B", "=cons_tp", "#,##0.00"); Vset("upc", "B", f"=B{VR['cons']}/px-1", "+0.0%;-0.0%;0.0%")
Vset("up", "B", f"=B{VR['tp']}/px-1", "+0.0%;-0.0%;0.0%"); Vset("tier", "B", "=" + tierf(f"B{VR['up']}"), "@"); Vset("vrate", "B", "=" + recf(f"B{VR['up']}"), "@")
Vset("gate", "B", f"=CF!C{GATE_ROW}", "#,##0.0")
Vset("rec", "B", f'=IF(AND(B{VR["gate"]}<0,OR(B{VR["vrate"]}="買進",B{VR["vrate"]}="強力買進")),"收支不足",B{VR["vrate"]})', "@")
V.cell(VR["tp"], 5, "兩腿權重 w_dcf；各腿先依情境權重加權")
V.cell(VR["vrate"], 5, "依上檔空間分界；入不敷出判定用列等於 CF 頁收支關卡（情境 B）")
V.column_dimensions["A"].width = 58
for c in "BCD": V.column_dimensions[c].width = 14
V.column_dimensions["E"].width = 70

# ═════════════════════════════ Thesis ═════════════════════════════
TH["A1"] = fx("命題（原型 A：按廣告計價）：AppLovin 的營收＝安裝量 × 單次安裝淨收入（公司只揭露年增率）。AI 在此的作用是提高單次安裝淨收入與打開電商、連網電視等新垂直（增量），同時有 AI 聊天介面分走安裝量的風險（侵蝕），並耗用資料中心運算與 LLM 推論（成本）。三情境的 AI 淨效果於 FY2032 為 ",
              T("D{0}".format(0), "#,##0"), "")
TH["A2"] = "單位經濟（每 US$1,000 營收）與 AI 淨效果：AI 淨效果（營業利益口徑）＝(AI 增量營收＋分流侵蝕)×(1−資料中心成本率−其他營業成本率)−LLM 推論成本。推論成本只讀 Tokenomics 快照；公司未揭露 token 量與任務量（未找到），任務數與倍數為 Assumed。變數名稱與 software_ai/unit_economics 一致處見 README；v0.2 改讀 unit_economics。"
THR = {}
r = 4
TH.cell(r, 1, "單位經濟：每 US$1,000 營收（情境 B；實績與預測）").font = BOLD; r += 1
hdr(TH, r, ["項目", "單位"] + [HYR[c] for c in YCOL]); THR["hdr_ue"] = r; r += 1
UE = [("ue_rev", "營收", "rev", 1), ("ue_dc", "資料中心成本（AI 運算）", "cogs_dc", 1), ("ue_llm", "LLM 推論成本", "cogs_llm", 1), ("ue_oc", "其他營業成本", "cogs_oth", 1),
      ("ue_gp", "毛利", "gp", 1), ("ue_opx", "營業費用", "opx", 1), ("ue_oi", "營業利益", "oi", 1), ("ue_sbc", "其中股票薪酬", "sbc", 1)]
for key, lab, pk, _ in UE:
    THR[key] = r; TH.cell(r, 1, lab); TH.cell(r, 2, "US$／千元營收")
    for c in YCOL:
        if pk in ("cogs_llm",) and c in ("C", "D", "E"): continue
        if pk == "sbc" and c in ("C", "D", "E") and not hist_has("sbc", c): continue
        TH[f"{c}{r}"] = f"=1000*PL!{c}{PLR['B'][pk]}/PL!{c}{PLR['B']['rev']}"; TH[f"{c}{r}"].number_format = "#,##0.0"
    r += 1
r += 1
THS = {}
for X in ("A", "B", "C"):
    P = PLR[X]
    TH.cell(r, 1, SCN_NAME[X] + "：AI 淨效果").font = BOLD; r += 1
    hdr(TH, r, ["項目", "單位"] + [HYR[c] for c in YCOL]); THS[X, "hdr"] = r; r += 1
    keys = [("noai", "反事實營收（無 AI 提升、新垂直、分流）", "US$M"), ("inc", "AI 增量營收（AI 提升＋新垂直）", "US$M"), ("ero", "分流侵蝕（營收減少；負值）", "US$M"), ("rev", "營收（＝PL）", "US$M"),
            ("ctb", "貢獻利潤率＝1−資料中心成本率−其他營業成本率", "%"), ("llm", "LLM 推論成本", "US$M"), ("net", "AI 淨效果（營業利益口徑）", "US$M"), ("chk", "檢查：反事實＋增量＋侵蝕 − 營收", "US$M")]
    for k, lab, u in keys:
        THS[X, k] = r; TH.cell(r, 1, lab + ("" if X == "B" else f" ｜情境 {X}")); TH.cell(r, 2, u); r += 1
    for c in ["F"] + FCOL:
        f = lambda k, v, fmt="#,##0.0": (TH.__setitem__(f"{c}{THS[X, k]}", v), setattr(TH[f"{c}{THS[X, k]}"], "number_format", fmt))
        f("noai", f"=PL!{c}{P['rev_noai']}"); f("inc", f"=PL!{c}{P['rev_nodiv']}-PL!{c}{P['rev_noai']}"); f("ero", f"=PL!{c}{P['rev']}-PL!{c}{P['rev_nodiv']}"); f("rev", f"=PL!{c}{P['rev']}")
        f("ctb", f"=1-PL!{c}{P['dcr']}-oc_r", "0.0%"); f("llm", f"=PL!{c}{P['cogs_llm']}")
        f("net", f"=({c}{THS[X, 'inc']}+{c}{THS[X, 'ero']})*{c}{THS[X, 'ctb']}-{c}{THS[X, 'llm']}")
        f("chk", f"={c}{THS[X, 'noai']}+{c}{THS[X, 'inc']}+{c}{THS[X, 'ero']}-{c}{THS[X, 'rev']}", "0.000")
    r += 1
TH.cell(r, 1, "AI 淨效果三情境對照（FY2032 與 FY2027–FY2032 累計）").font = BOLD; r += 1
hdr(TH, r, ["項目", "單位", "情境 A", "情境 B", "情境 C"]); THR["cmp_h"] = r; r += 1
for key, lab, fn in (("c_inc", "FY2032 AI 增量營收", lambda X: f"=L{THS[X, 'inc']}"), ("c_ero", "FY2032 分流侵蝕", lambda X: f"=L{THS[X, 'ero']}"), ("c_llm", "FY2032 LLM 推論成本", lambda X: f"=L{THS[X, 'llm']}"),
                     ("c_net", "FY2032 AI 淨效果（營業利益口徑）", lambda X: f"=L{THS[X, 'net']}"), ("c_cum", "FY2027–FY2032 累計 AI 淨效果", lambda X: f"=SUM(G{THS[X, 'net']}:L{THS[X, 'net']})")):
    THR[key] = r; TH.cell(r, 1, lab); TH.cell(r, 2, "US$M")
    for j, X in enumerate("ABC"): TH.cell(r, 3 + j, fn(X)).number_format = "#,##0"
    r += 1
THR["c_spread"] = r; TH.cell(r, 1, "檢查：三情境 AI 淨效果分化（情境 A − 情境 C，應大於 0）"); TH.cell(r, 3, f"=C{THR['c_net']}-E{THR['c_net']}").number_format = "#,##0"
TH.column_dimensions["A"].width = 58
for c in YCOL: TH.column_dimensions[c].width = 11
TH["A1"] = fx("命題（原型 A：按廣告計價）：營收＝安裝量 × 單次安裝淨收入（公司只揭露年增率，絕對值未找到）。AI 的作用有三：提高單次安裝淨收入並打開電商、連網電視等新垂直（增量）；AI 聊天介面分走安裝量（侵蝕）；耗用資料中心運算與 LLM 推論（成本）。FY2032 AI 淨效果（營業利益口徑）情境 A／B／C＝",
              T(f"C{THR['c_net']}", "#,##0"), "／", T(f"D{THR['c_net']}", "#,##0"), "／", T(f"E{THR['c_net']}", "#,##0"), " 百萬美元。")


# ═════════════════════════════ RevDCF ═════════════════════════════
RVR = {}
RV["A1"] = "RevDCF — 反向 DCF：現價隱含什麼（情境 B 路徑；封閉解，不另跑反解）。以現價回推企業價值，扣掉明確期 FCFF 現值後，得到終值需要的 FY2032 NOPAT，與模型比較。"
RV["A2"] = "只作展示、不進目標價。現價日與評價日不同（現價日在評價日之後），隱含值以評價日折現口徑直接比較，未做時間調整。"
CB = CFR["B"]; PB = PLR["B"]
rows = [("sum_dcf", None, None, None), ("sum_cmp", None, None, None),
        ("px", "現價（US$）", "=px", "#,##0.00"), ("mcap", "股權市值（現價 × 評價日股數；US$M）", f"=Val!B{VR['emkt']}", "#,##0"),
        ("nc", "淨現金與非營業項目（現金−債務−融資租賃−類債＋權益法投資；US$M）", "=cash_0-debt_0-fin_lease-legal_oth+eq_inv", "#,##0"),
        ("evm", "現價隱含企業價值 EV（US$M）", None, "#,##0"), ("pvx", "明確期 FCFF 現值（模型，情境 B；US$M）", f"=Val!C{VR['pvx']}", "#,##0"),
        ("tvn", "現價隱含的終值（FY2032 年底；US$M）", None, "#,##0"), ("tvm", "模型終值（情境 B；US$M）", f"=Val!C{VR['tv']}", "#,##0"),
        ("nopn", "現價隱含的 FY2032 NOPAT（US$M）", None, "#,##0"), ("nopm", "模型 FY2032 NOPAT（US$M）", f"=CF!L{CB['nopat']}", "#,##0"),
        ("scl", "隱含 NOPAT ÷ 模型 NOPAT（1＝現價與模型一致）", None, "0.00"),
        ("rv0", "模型 FY2032 營收（US$M）", f"=PL!L{PB['rev']}", "#,##0"), ("rv1", "隱含 FY2032 營收（同利潤率；US$M）", None, "#,##0"),
        ("g0", "模型營收年複合成長（FY2026–FY2032）", f"=(PL!L{PB['rev']}/PL!F{PB['rev']})^(1/6)-1", "0.0%"), ("g1", "隱含營收年複合成長（FY2026–FY2032）", None, "0.0%"),
        ("w0", "折現率 WACC（模型）", f"=Val!B{VR['wacc']}", "0.00%"),
        ("pe_px", "現價 ÷ 目標時點 NTM EPS（情境 B；倍）", f"=px/Val!C{VR['ntm']}", "0.0"), ("pe_m", "Comparable 採用倍數（倍）", f"=Val!C{VR['mult']}", "0.0"), ("pe_gap", "現價隱含倍數相對採用倍數", None, "+0.0%;-0.0%;0.0%")]
r = 4
for k, lab, f, fmt in rows:
    RVR[k] = r; RV.cell(r, 1, lab)
    if f: c = RV.cell(r, 2, f); c.number_format = fmt
    r += 1
def RS(k, f, fmt): c = RV.cell(RVR[k], 2, f); c.number_format = fmt
RS("evm", f"=B{RVR['mcap']}-B{RVR['nc']}", "#,##0")
RS("tvn", f"=(B{RVR['evm']}-B{RVR['pvx']})*(1+B{RVR['w0']})^(CF!L{CB['tm']}+0.5)", "#,##0")
RS("nopn", f"=B{RVR['tvn']}*(B{RVR['w0']}-g_tv)/((1+g_tv)*(1-g_tv/Val!B{VR['ronic']}))", "#,##0")
RS("scl", f"=B{RVR['nopn']}/B{RVR['nopm']}", "0.00"); RS("rv1", f"=B{RVR['rv0']}*B{RVR['scl']}", "#,##0")
RS("g1", f"=(B{RVR['rv1']}/PL!F{PB['rev']})^(1/6)-1", "0.0%"); RS("pe_gap", f"=B{RVR['pe_px']}/B{RVR['pe_m']}-1", "+0.0%;-0.0%;0.0%")
RV.cell(RVR["sum_dcf"], 1, fx("現金流折現法：現價相當於 FY2032 稅後營業利益為模型的 ", T(f"B{RVR['scl']}", "0.00"), " 倍（同利潤率下 FY2032 營收約 ", T(f"B{RVR['rv1']}", "#,##0"), " 百萬美元，年複合成長 ",
                                T(f"B{RVR['g1']}", "0.0%"), "；模型為 ", T(f"B{RVR['g0']}", "0.0%"), "）。折現率 ", T(f"B{RVR['w0']}", "0.0%"), "。"))
RV.cell(RVR["sum_cmp"], 1, fx("同業本益比法：現價相當於目標時點下一年 EPS 的 ", T(f"B{RVR['pe_px']}", "0.0"), " 倍，採用倍數為 ", T(f"B{RVR['pe_m']}", "0.0"), " 倍（現價隱含相對採用 ", T(f"B{RVR['pe_gap']}", "+0.0%;-0.0%"), "）。"))
RV.column_dimensions["A"].width = 70; RV.column_dimensions["B"].width = 16

# ═════════════════════════════ Recon ═════════════════════════════
RCR = {}
RC["A1"] = "Recon — 對帳：賣方共識、公司指引、DCF 與 Comparable 落差、反解與 Tokenomics 讀值。共識與指引皆為 Interested-party，只作對照，不進模型。"
r = 3
def rcw(r, vals, fmts=None):
    for j, v in enumerate(vals):
        if v is None: continue
        c = RC.cell(r, 1 + j, v)
        if fmts and j < len(fmts) and fmts[j] and not isinstance(v, str) or (fmts and j < len(fmts) and fmts[j] and isinstance(v, str) and v.startswith("=")): c.number_format = fmts[j]
RC.cell(r, 1, "1. 賣方共識對照（營收；US$M）").font = BOLD; RCR["s1"] = r; r += 1
hdr(RC, r, ["項目", "年度", "情境 A", "情境 B", "情境 C", "共識", "B 對共識", "說明"]); r += 1
cons_rows = [("營收", "FY2026E", "F", 8340.0, "Q3 財報後共識；9M26 實績＋Q4 指引已隱含較高值（另一來源 8,050 疑含過期預估）"), ("營收", "FY2027E", "G", 10189.0, "財報前共識（四季相加 10,490 為推導值）；財報後未取得官方彙總（未找到）"),
             ("營收", "FY2028E", "H", 12421.0, "財報前共識")]
for lab, yr, col, cv, note in cons_rows:
    RCR[("cons", col)] = r
    rcw(r, [lab, yr] + [f"=PL!{col}{PLR[X]['rev']}" for X in "ABC"] + [cv, f"=D{r}/F{r}-1", note], [None, None, "#,##0", "#,##0", "#,##0", "#,##0", "+0.0%;-0.0%", None])
    RC.cell(r, 6).font = BLUE; r += 1
RCR["eps_note"] = r
RC.cell(r, 1, "EPS：賣方共識為 non-GAAP（排除股票薪酬），本模型為 GAAP 稀釋，口徑不同，不直接比較。FY2026 共識 17.35、FY2027 財報前 19.89（Interested-party）；本模型 FY2026E／FY2027E GAAP EPS 見 PL 頁。"); r += 2
RC.cell(r, 1, "2. 公司指引對照（2026Q4；US$M；指引為 Interested-party，不用於校準）").font = BOLD; RCR["s2"] = r; r += 1
hdr(RC, r, ["項目", "本模型（情境 B）", "指引低", "指引高", "情境 A", "說明"]); r += 1
RCR["g_rev"] = r
rcw(r, ["Q4 營收", f"=PL!N{PLR['B']['rev']}", 2365.0, 2395.0, f"=PL!N{PLR['A']['rev']}", "基準取介於指引與歷史 Q4 季增之間；公司近年指引多被超越（Q3 指引 2,055–2,085，實際 2,184.8）"], [None, "#,##0", "#,##0", "#,##0", "#,##0"])
RC.cell(r, 3).font = BLUE; RC.cell(r, 4).font = BLUE; r += 1
RCR["g_gap"] = r
rcw(r, ["Q4 營收高於指引中點", f"=B{r - 1}/AVERAGE(C{r - 1}:D{r - 1})-1", None, None, f"=E{r - 1}/AVERAGE(C{r - 1}:D{r - 1})-1", "Q3 實際高於指引中點約 +5.9%（2,184.815 ÷ 2,070）"], [None, "+0.0%;-0.0%", None, None, "+0.0%;-0.0%"]); r += 1
RCR["g_eb"] = r
rcw(r, ["Q4 調整後 EBITDA（本模型口徑）", f"=PL!N{PLR['B']['aeb']}", 1995.0, 2025.0, None, "公司定義含其他調整；本模型＝營業利益＋折舊攤銷＋股票薪酬"], [None, "#,##0", "#,##0", "#,##0"]); RC.cell(r, 3).font = BLUE; RC.cell(r, 4).font = BLUE; r += 2
RC.cell(r, 1, "3. DCF 與 Comparable 落差分解（情境 B；US$／股；公式組字）").font = BOLD; RCR["s3"] = r; r += 1
hdr(RC, r, ["項目", "數值", "說明"]); r += 1
VB = lambda k: f"Val!C{VR[k]}"
gaprows = [("gap_comp", "Comparable 每股價值（情境 B）", f"={VB('tpc_s')}", "#,##0.00", "同業中位數 × NTM EPS"),
           ("gap_dcf", "DCF 每股價值（情境 B；目標時點）", f"={VB('ps12')}", "#,##0.00", ""),
           ("gap", "落差（Comparable − DCF）", None, "+#,##0.00;-#,##0.00", ""),
           ("gap_pe", "DCF 隱含目標時點前瞻本益比（倍）", f"={VB('ps12')}/{VB('ntm')}", "0.0", "DCF 每股價值 ÷ NTM EPS"),
           ("gap_mult", "倍數項：(同業採用倍數 − DCF 隱含倍數) × NTM EPS", f"=({VB('mult')}-B{{gap_pe}})*{VB('ntm')}", "+#,##0.00;-#,##0.00", "與落差相等；落差的全部以倍數差表示"),
           ("gap_time", "時點項：DCF 由評價日推到目標時點的增加（ke 推進）", f"={VB('ps12')}-{VB('ps0')}", "+#,##0.00;-#,##0.00", "DCF 已含，不是落差來源，列示以供對照"),
           ("gap_sbc", "股票薪酬（參考）：NTM 股票薪酬稅後每股 × 同業倍數", f"=PL!H{PLR['B']['sbc']}*(1-tax)/shval*{VB('mult')}", "#,##0.00", "同業分母若排除股票薪酬，套到 GAAP EPS 的方向性差距（保守側）"),
           ("gap_lev", "槓桿（參考）：每股淨現金（DCF 的 EV 橋接項）", f"=(Val!C{VR['eqv']}-Val!C{VR['ev']})/shval", "#,##0.00", "債務小，槓桿口徑差異有限"),
           ("gap_ke", "折現率（參考）：ke 與 WACC", f"=Val!B{VR['ke']}", "0.0%", "β 約 2 使 ke 偏高，主導 DCF 低於同業倍數；反解見下節")]
for k, lab, f, fmt, note in gaprows:
    RCR[k] = r; r += 1
for k, lab, f, fmt, note in gaprows:
    rr_ = RCR[k]; RC.cell(rr_, 1, lab)
    if f: RC.cell(rr_, 2, f.replace("{gap_pe}", str(RCR["gap_pe"]))).number_format = fmt
    RC.cell(rr_, 3, note)
RC.cell(RCR["gap"], 2, f"=B{RCR['gap_comp']}-B{RCR['gap_dcf']}").number_format = "+#,##0.00;-#,##0.00"
RCR["gap_txt"] = r
RC.cell(r, 1, fx("落差歸因：Comparable ", T(f"B{RCR['gap_comp']}", "#,##0.00"), " 高於 DCF ", T(f"B{RCR['gap_dcf']}", "#,##0.00"), " 約 ", T(f"B{RCR['gap']}", "#,##0.00"),
                 " 美元；DCF 隱含前瞻本益比 ", T(f"B{RCR['gap_pe']}", "0.0"), " 倍，同業採用 ", T(f"{VB('mult')}", "0.0"), " 倍。"))
r += 2
RC.cell(r, 1, "4. 反解：支持現價所需的主導變數值（整本重算；只作展示，不進目標價）").font = BOLD; RCR["s4"] = r; r += 1
hdr(RC, r, ["主導變數（In 頁鍵）", "基準值", "使加權目標價＝現價的值", "使上檔空間＝+25%（買進邊界）的值", "狀態"]); r += 1
RCR["kreq"] = {}
for key, lab, fmt in (("beta", "β（Blume 調整後）", "0.000"), ("u0_B", "情境 B：FY2027 AI 提升單次安裝淨收入", "0.0%"), ("e0_B", "情境 B：FY2027 新垂直增量", "0.0%")):
    RCR["kreq"][key] = r
    RC.cell(r, 1, f"{lab}（{key}）"); RC.cell(r, 2, f"={key}").number_format = fmt
    kk = KREQ.get(key)
    if kk:
        RC.cell(r, 3, kk["req"]["value"]).number_format = fmt; RC.cell(r, 3).font = BLUE
        RC.cell(r, 4, kk["flip"]["value"]).number_format = fmt; RC.cell(r, 4).font = BLUE
        RC.cell(r, 5, "已解" if kk["req"]["solved"] and kk["flip"]["solved"] else "未收斂")
    else:
        RC.cell(r, 5, "尚未重跑（KREQ＝1）")
    r += 1
RCR["kreq_chk"] = r
RC.cell(r, 1, "反解快照檢查：反解時的基準加權目標價 − 目前加權目標價（應約為 0；不為 0＝反解已過期）")
if KREQ.get("base_tp") is not None:
    RC.cell(r, 2, f"={KREQ['base_tp']}-Val!B{VR['tp']}").number_format = "0.000"
r += 2
RC.cell(r, 1, "5. Tokenomics 讀值（只讀快照 v22.0；推論成本不自行估價）").font = BOLD; RCR["s5"] = r; r += 1
hdr(RC, r, ["項目", "數值", "說明"]); r += 1
TKR = {}
for k, lab, f, fmt, note in (("t_terra", "IF_PriceRef_Terra（US$／百萬 token）", f"={TK['items']['IF_PriceRef_Terra']['values'] if False else 0.784}", "0.000", "快照原值"),):
    pass
tkv = lambda n: TK["items"][n]["values"]
def _first(v):
    return v if isinstance(v, (int, float)) else None
for k, lab, nm in (("t_terra", "IF_PriceRef_Terra（低層）", "IF_PriceRef_Terra"), ("t_sol", "IF_PriceRef_Sol（本模型採用）", "IF_PriceRef_Sol"), ("t_astra", "IF_PriceRef_Astra（高層）", "IF_PriceRef_Astra")):
    TKR[k] = r; RC.cell(r, 1, lab + "；US$／百萬 token"); RC.cell(r, 2, _first(tkv(nm))).number_format = "0.000"; RC.cell(r, 2).font = BLUE
    RC.cell(r, 3, f"Tokenomics 快照 {nm}；Analogy（選哪一層價格為類比）"); r += 1
TKR["cost27"] = r
RC.cell(r, 1, "FY2027E LLM 推論成本（情境 B；US$M；本模型採用價）"); RC.cell(r, 2, f"=PL!G{PLR['B']['cogs_llm']}").number_format = "#,##0.0"
RC.cell(r, 3, "任務數與倍數為 Assumed（公司未揭露，未找到），占營收比例極低；高案聯動見 Sens 的上界與路由案例"); r += 1
TKR["share27"] = r
RC.cell(r, 1, "FY2027E LLM 推論成本占營收比"); RC.cell(r, 2, f"=B{r - 1}/PL!G{PLR['B']['rev']}").number_format = "0.000%"; r += 2
RC.cell(r, 1, "6. 逾期標示（與股價基準日相差超過 2 個交易日）").font = BOLD; RCR["s6"] = r; r += 1
for lab, dtv in (("無風險利率 rf（2026-10-27；跨模型沿用 Meta v0.2.1 讀值）", dt.datetime(2026, 10, 27)), ("Unity 預估本益比（2026-10-26）", dt.datetime(2026, 10, 26)),
                 ("PINS、TTD、APP 預估本益比（2026-11-06）", dt.datetime(2026, 11, 6))):
    RC.cell(r, 1, lab); c = RC.cell(r, 2, dtv); c.number_format = "yyyy-mm-dd"
    RC.cell(r, 3, f'=IF(NETWORKDAYS(B{r},px_date)-1>2,"逾期","")'); r += 1
r += 1
RC.cell(r, 1, "7. 未找到清單").font = BOLD; RCR["s7"] = r; r += 1
NF = ["電商（e-commerce）廣告營收：未找到（公司未分項揭露）", "廣告主數量：未找到", "AI 歸因營收：未找到（公司不揭露 AI 帶來的營收）", "產品線營收（AXON 廣告、MAX 變現、Adjust 等）：未找到（單一應報導分部）",
      "安裝量與單次安裝淨收入的絕對值：未找到（只揭露年增率）", "每任務 token 與代理任務量：未找到（G2-1 無揭露可校準，採快照倍數，標 Assumed）", "FY2027 財報後官方共識彙總：未找到（僅取得財報前與資料商相加值）",
      "SEC 調查現況的主要文件確認：未找到（僅管理層口頭說已結束；10-Q 載證券集體訴訟，損失未估）", "融資租賃負債：未找到（新聞稿未單列，列 0）", "RSU 沒收率與 2026-09-30 未歸屬 RSU：未找到（取 2025-12-31）",
      "2025Q4 單季資料中心成本與股票薪酬：未找到", "法定申報封面最新流通股數（G2-4 主來源）：未找到（改用 2026-09-30 資產負債表 331.322 百萬股）"]
for t in NF: RC.cell(r, 1, t); r += 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFG": RC.column_dimensions[c].width = 16
RC.column_dimensions["H"].width = 60

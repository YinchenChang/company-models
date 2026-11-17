# Tesla Robotaxi 分部模型 v0.2（Tesla 分部，視同未上市）— 建置腳本
# v0.2：價格、營運成本（第 3–8 項）、稼動率改讀 Robotaxi 線 v0.2 成品（Tesla 車隊加權；建置時由線 xlsx 讀值，不手打）；
#       自有車閘門便利溢價基準 cf＝1.0（對帳 1 判斷題預設乙）；新增輸出 RT_VAL_A/B/C、RT_REV_<情境>_<年>。
# 規格：/home/claude/model_brief.md「Robotaxi 分部模型特別規格」＋ 2026-11-16 22:38 補充（滲透頁最上方「價格」區塊）；
#       docs/plan/個股DCF標準規格.md（FCFF、WACC、RONIC＝WACC）；docs/plan/個股模型原則_命題加營收分部.md 第 8–10 節（In 頁版面、Sens 覆蓋、治理）。
# 不讀 Siliconomics／Tokenomics／Wattonomics 協調檔（實體 AI 線不適用）。
# 用法：python3 physical_ai/tesla_robotaxi/build_tesla_robotaxi.py   → dist/20261116_TeslaRobotaxi個股模型_v0.1.xlsx
#       之後以 LibreOffice 重算並跑 check_tesla_robotaxi.py（公式值＝Python 引擎值、0 錯誤值、三支 repo 核對程式）。
import glob, hashlib, json, math, os, sys
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
VER = "v0.2"
DATE = "20261116"
OUT = os.path.join(HERE, "dist", f"{DATE}_TeslaRobotaxi個股模型_{VER}.xlsx")
YEARS = list(range(2026, 2036))           # 2026＝觀察／校準欄；2027–2035 明確期
DCF_YEARS = list(range(2027, 2036))
SCN = ["A", "B", "C"]
SCN_NAME = {"A": "情境 A（樂觀）", "B": "情境 B（基準）", "C": "情境 C（保守）"}

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SEC = PatternFill("solid", fgColor="FCE4D6")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------------------------------------------------------- 來源（README 與各頁共用）
S = {
    "q3call": "Not a Tesla App，Tesla Q3 2026 法說摘要（2026-10-22）https://notateslaapp.com/news/4867/tesla-q3-2026-earnings-call-recap-everything-you-need-to-know",
    "q3res": "Drive Tesla Canada，Tesla Q3 2026 財報（2026-10-21）https://driveteslacanada.ca/news/tesla-q3-2026-earnings-results/",
    "safety": "Not a Tesla App，Tesla Robotaxi 里程與安全報告（資料 2026-01-22～10-18）https://notateslaapp.com/news/4860/tesla-names-next-robotaxi-cities-shares-miles-driven-safety-data",
    "q2deck": "EVwire 轉述 Tesla Q2 2026 Shareholder Update https://evwire.com/p/tesla-robotaxi-2-4-million-paid-miles-six-metros",
    "ramp": "Basenor，Tesla 無人化進度（2026-10）https://www.basenor.com/blogs/news/teslas-driverless-ramp-explained-in-5-points",
    "mpd": "Basenor 轉述 Sawyer Merritt（2026-10-21）https://www.basenor.com/blogs/news/tesla-robotaxi-hits-80-600-miles-per-day-launches-first-responder-training-site",
    "txreg": "EVwire，德州 TxDMV AV 名錄（2026-11-11）https://evwire.com/p/tesla-overtakes-waymo-texas-robotaxi-registrations",
    "tracker_bn": "Basenor Tesla Robotaxi Tracker（2026-11-16）https://www.basenor.com/pages/tesla-robotaxi-tracker",
    "geo": "EVwire，Austin 服務區 440 平方英里（2026-11-09）https://evwire.com/p/tesla-robotaxi-austin-geofence-expansion-440-square-miles-update",
    "may": "CarScoops 引 Robotaxi Tracker（2026-05-27）https://www.carscoops.com/2026/05/tesla-robotaxi-fleet-shrinking/",
    "hey": "heygotrade（2026-11-10 收盤）https://www.heygotrade.com/en/news/tesla-tsla-stock-news-november-11-2026/",
    "obi": "Automotive World，Obi 9.4 萬趟（2025-11～2026-01）https://www.automotiveworld.com/news/obi-study-compares-tesla-waymo-uber-pricing/",
    "tc_obi": "TechCrunch，Obi Bay Area 平均車資與 ETA（2026-01-27）https://techcrunch.com/2026/01/27/the-price-gap-between-waymo-and-uber-is-narrowing",
    "price_dtc": "Drive Tesla Canada，Austin 動態計價（2025-07-30）https://driveteslacanada.ca/news/tesla-introduces-dynamic-fare-pricing-for-robotaxi-rides-in-austin/",
    "price_aol": "AOL／TheStreet，Austin 基本費調至 3.25（2026-03-07）https://www.aol.com/finance/tesla-robotaxi-prices-just-jumped-180300054.html",
    "price_bn": "Basenor，Austin 計價歷程（2026-03，07-16 更新）https://www.basenor.com/blogs/news/tesla-robotaxi-pricing-just-changed-base-fee-triples-to-3",
    "price_tp": "tparts，Austin 每英里費率 +40%（2026-03-12）https://www.tparts.com/blogs/tesla-latest-news/tesla-raises-robotaxi-fares-in-austin-second-price-hike-in-weeks-sees-mileage-fees-jump-40",
    "price_driven": "The Driven，Tesla 調漲 Robotaxi 價格最多 40%（2026-03-17）https://thedriven.io/2026/03/17/tesla-lifts-robotaxi-prices-by-up-40-pct-less-than-a-year-after-launch/",
    "cc_ride": "Basenor 轉述 Joe Tegtmeyer Cybercab 19.4 英里行程（2026-09-09）https://www.basenor.com/blogs/news/tesla-robotaxi-vs-uber-the-numbers-are-in",
    "jef": "Stocktwits 轉述 Jefferies Austin 實測（2026-02-26）https://stocktwits.com/news-articles/markets/equity/tesla-robotaxis-underperform-rideshare-in-austin-on-key-metrics-despite-discounting-jefferies/cZTGn7rRIRa/amp",
    "gridwise": "Gridwise Q3 2026 Rideshare Report https://gridwise.io/analytics/q3-2026-rideshare-delivery-report",
    "lyft": "Lyft Q3 2026 8-K 新聞稿 https://www.sec.gov/Archives/edgar/data/0001759509/000162828026071043/lyft-20260930xpressrelease.htm",
    "uber": "Uber Q3 2026 新聞稿 https://seekingalpha.com/pr/20712391",
    "aaa": "AAA Your Driving Costs 2025 Fact Sheet https://newsroom.aaa.com/wp-content/uploads/2025/09/UPDATE-AAA-Fact-Sheet-Your-Driving-Cost-9.2025-1.pdf",
    "fred": "FRED M12MTVUSM227NFWA（FHWA 12 個月移動 VMT，2026-08）https://fred.stlouisfed.org/series/M12MTVUSM227NFWA",
    "te10y": "Trading Economics 美國 10 年期公債（2026-11-13 約 5.15%）https://tradingeconomics.com/united-states/government-bond-yield/news/598729",
    "fortune": "Fortune，Waymo 1,260 億美元估值、Other Bets Q3 2026（2026-10-30）https://fortune.com/2026/10/30/waymo-126-billion-alphabet-372-million-2-05-billion-google",
    "line1a": "線 1a Robotaxi 框架 v0.1（2026-11-16，/mnt/user-data/outputs/lines/線1a_Robotaxi_框架_v0.1.md）",
    "yipit": "YipitData，Waymo 採用者叫車行為（2026-08-21）https://ask.yipitdata.com/insights/uber-riders-waymo-total-rideshare-growth",
    "cpuc": "CPUC AV Program Quarterly Reporting https://www.cpuc.ca.gov/regulatory-services/licensing/transportation-licensing-and-analysis-branch/autonomous-vehicle-programs/quarterly-reporting",
    "cpuc_tsla": "Electrek，CPUC 確認 Tesla 在加州非 AV 服務（2026-03-25）https://electrek.co/2026/03/25/california-regulator-confirms-tesla-not-operating-autonomous-vehicle-service/",
    "rtt": "robotaxitracker.com（社群追蹤；Methodology 頁 https://robotaxitracker.com/methodology）",
}

# ---------------------------------------------------------------- 讀線 1a v0.2（權威：physical_ai/lines/robotaxi/dist）
LINE_GLOB = os.path.join(HERE, "..", "lines", "robotaxi", "dist", "*Robotaxi單位經濟_v0.2.xlsx")
LINE_ANCH = [2026, 2028, 2030, 2032]            # 線錨點年
CASE_IDX = {"A": 0, "B": 1, "C": 2}              # A＝低成本＋低價；B＝基準；C＝高成本＋高價
ST_MY, ST_CC = 28, 51                            # 堆疊頁 Tesla Model Y／Cybercab 區塊第 1 項所在列
ST_FLEET = 124                                   # 堆疊頁「Tesla 合計（車隊英里加權）｜單位營運成本」低成本列（+1 基準、+2 高成本；欄 C＝2026…I＝2032）
PR_MY, PR_CC = 56, 81                            # 成本輸入頁 TY_PRICE／TC_PRICE


def _geo_path(vals):
    """錨點 2026/28/30/32 → 2026–2035：奇數年幾何平均，2033 起持平於 2032（對帳 1 判斷題 2 預設）"""
    out = {}
    for y in range(2026, 2036):
        if y in vals:
            out[y] = vals[y]
        elif y < 2032:
            out[y] = math.sqrt(vals[y - 1] * vals[y + 1])
        else:
            out[y] = vals[2032]
    return out


def load_line():
    path = sorted(glob.glob(LINE_GLOB))[-1]
    wb = openpyxl.load_workbook(path, data_only=True)
    st, ci = wb["每英里成本堆疊"], wb["成本輸入"]
    # 版面斷言：列號改動時立即失敗，不靜默讀錯格
    assert str(st.cell(ST_MY - 2, 1).value).startswith("Tesla Model Y") and str(st.cell(ST_MY, 1).value).startswith("1 車輛")
    assert str(st.cell(ST_CC - 2, 1).value).startswith("Tesla Cybercab") and str(st.cell(ST_CC, 1).value).startswith("1 車輛")
    assert str(st.cell(ST_MY + 15, 1).value).startswith("備忘：每車每日付費英里") and str(st.cell(ST_CC + 15, 1).value).startswith("備忘：每車每日付費英里")
    for k in range(3):
        assert "Tesla 合計（車隊英里加權）｜單位營運成本" in str(st.cell(ST_FLEET + k, 1).value)
    assert ci.cell(PR_MY, 1).value == "TY_PRICE" and ci.cell(PR_CC, 1).value == "TC_PRICE"
    g = lambda ws, r, c: ws.cell(r, c).value
    L = {"file": os.path.basename(path), "path": path, "w": {}, "opx": {}, "pmd": {}, "price": {}, "cells": {}}
    for s, c in CASE_IDX.items():
        a_opx, a_pmd, a_pr = {}, {}, {}
        for i, yr in enumerate(LINE_ANCH):
            sc, ic = 3 + i * 3 + c, 5 + i * 3 + c                      # 堆疊頁欄、成本輸入頁欄
            my = {k: g(st, ST_MY + k - 1, sc) for k in range(1, 9)}
            cc = {k: g(st, ST_CC + k - 1, sc) for k in range(1, 9)}
            fleet = g(st, ST_FLEET + c, 3 + (yr - 2026))                # 車隊加權單位營運成本（1–8 項）
            w = (fleet - sum(cc.values())) / (sum(my.values()) - sum(cc.values()))
            a_opx[yr] = w * sum(my[k] for k in range(3, 9)) + (1 - w) * sum(cc[k] for k in range(3, 9))
            a_pmd[yr] = w * g(st, ST_MY + 15, sc) + (1 - w) * g(st, ST_CC + 15, sc)
            a_pr[yr] = w * g(ci, PR_MY, ic) + (1 - w) * g(ci, PR_CC, ic)
            L["w"][(s, yr)] = w
            L["cells"][(s, yr)] = dict(sc=CL(sc), ic=CL(ic), fc=CL(3 + (yr - 2026)), fr=ST_FLEET + c)
        L["opx"][s], L["pmd"][s], L["price"][s] = _geo_path(a_opx), _geo_path(a_pmd), _geo_path(a_pr)
    return L


LINE = load_line()

# ---------------------------------------------------------------- In 頁輸入
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 出處, 來源誘因/說明)；基準可為公式字串（Derived）
IN = [
    ("sec", "1. 市場基準：美國叫車（人類司機）"),
    ("lyft_rq", "Lyft 單季乘車趟數（Q3 2026）", "百萬趟", None, 273.5, None, "Verified", S["lyft"], "法定申報（8-K 附件）"),
    ("lyft_na", "Lyft 趟數中北美占比（扣 FreeNow 歐洲）", "比例", 0.85, 0.90, 0.95, "Assumed", "Lyft 未分地區揭露趟數", "區間；FreeNow 2025 年併入"),
    ("lyft_sh", "Lyft 美國叫車趟數份額", "比例", 0.23, 0.26, 0.30, "Analogy", "類比：Bloomberg Second Measure／YipitData 等刷卡資料歷年約 1/4（本版未取得 2026 原文數字）", "區間；份額越低 → 美國總趟數越大"),
    ("rh_fare", "UberX 平均每趟車資（乘客付）", "US$/趟", None, 24.30, None, "Verified", S["gridwise"], "Gridwise 司機樣本，非 Uber 揭露"),
    ("rh_take", "Uber 平台抽成中位數", "比例", None, 0.2244, None, "Verified", S["gridwise"], "同上"),
    ("drv_pm", "Uber 司機每英里毛收入", "US$/英里", None, 2.65, None, "Verified", S["gridwise"], "同上"),
    ("trip_mi", "人類叫車平均行程長度", "英里/趟", 6.0, "=rh_fare*(1-rh_take)/drv_pm", 8.0, "Derived", "車資 ×（1 − 抽成）÷ 司機每英里收入（Gridwise 三數相除）", "推算；低／高為 Assumed 區間"),
    ("tip", "小費加成（乘客實付 ÷ 車資）", "比例", 0.05, 0.08, 0.15, "Assumed", "Gridwise 未揭露小費率", "區間"),
    ("rh_g", "美國叫車里程年成長（2026→2035，不含 AV 誘發）", "年增率", 0.03, 0.06, 0.09, "Assumed", "對照 Lyft Q3 2026 趟數年增 10%、Uber 全平台趟數年增 16%", "區間"),
    ("vmt", "美國 12 個月移動車輛行駛里程（2026-08）", "十億英里", None, 3335.695, None, "Verified", S["fred"], "政府統計"),
    ("ldv_sh", "VMT 中可被 Robotaxi 取代的小客車占比", "比例", 0.85, 0.90, 0.92, "Assumed", "扣除重卡、商用車；FHWA 分車種表未取用", "區間"),
    ("p_own", "自有車每英里全成本（AAA 2025，15,000 英里/年平均）", "US$/英里", 0.56, 0.77, 0.77, "Verified", S["aaa"], "低＝AAA 小轎車 0.5587"),
    ("cf", "便利溢價：Robotaxi 每英里價格可高於自有車成本的倍數（轉換閘門門檻）", "倍", 1.0, 1.0, 1.5, "Assumed", "不用開車、停車、保險的時間價值；無直接證據。v0.2 基準由 1.5 改 1.0（對帳 1 判斷題預設乙：價格 ≤ 0.77 全開、≥ 1.155 關閉、中間線性；上緣 1.5＝v0.1 設定）", "區間 1.0–1.5；閘門＝MIN(1,MAX(0,(1.5×門檻−價)/(0.5×門檻)))：價≤門檻全開、價≥1.5×門檻關閉"),
    ("eps", "叫車需求價格彈性（誘發，傑文斯效應）", "彈性", 0.3, 0.7, 1.2, "Assumed", "觀察值未取得（見滲透頁「觀察到的價格彈性」）", "區間；誘發倍數＝(人類叫車價 ÷ Robotaxi 價)^彈性"),
    ("lk", "Robotaxi 滲透 S 曲線斜率", "1/年", 0.6, 0.8, 1.2, "Assumed", "—", "區間"),
    ("lmid", "Robotaxi 滲透 S 曲線中點年", "年", 2030, 2031, 2032, "Assumed", "—", "區間"),
    ("sec", "2. Tesla 現況（2026 校準欄）"),
    ("mi_day", "Tesla 無人 Robotaxi 每日里程（2026-10）", "英里/日", None, 80600, None, "Interested-party", S["mpd"], "社群轉述（Sawyer Merritt 為 Tesla 長期支持者）；Tesla 未在財報揭露"),
    ("mi_ytd", "Tesla 無人里程累計（2026-01-22～10-18，不含 Bay Area）", "英里", None, 3100000, None, "Interested-party", S["safety"], "Tesla 自揭安全報告；有推升估值誘因"),
    ("fleet26", "Tesla Robotaxi 車隊（2026 年底，含 Bay Area 有安全員）", "台", 1100, 1500, 2000, "Derived", "德州 TxDMV 名錄 1,302 台（2026-11-11）＋ Bay Area 約 100–170 台（Obi 抽樣期 156 台）", "名錄＝已登記，不等於在營運；低／高為 Assumed"),
    ("p26", "Tesla 2026 實收每英里價格（混合）", "US$/英里", 1.40, 1.80, 2.10, "Derived", "Austin 3＋1.40×英里（5 英里＝2.00）、Cybercab 單趟 1.41、Bay Area 1.99/公里＝3.20/英里（有安全員）", "推算；混合權重 Assumed"),
    ("sig26", "Tesla 占美國 Robotaxi 里程份額（2026）", "比例", 0.08, 0.15, 0.25, "Assumed", "Tesla 約 8 萬英里/日 vs Waymo 每週 40–50 萬趟 × 約 5 英里 ÷ 7（約 30 萬英里/日）", "區間；Waymo 每趟英里未揭露"),
    ("sec", "3. 情境驅動（2035 年終點，2026/2027 起線性或 S 曲線）"),
]
SCN_IN = [
    ("s35", "Robotaxi 占美國叫車基準里程（取代率，所有營運商）2035", "比例", (0.50, 0.25, 0.08), "Assumed", "對照 Waymo 舊金山 2026 已占叫車市場可觀份額（YipitData 只揭露行為變化）；無 2035 證據"),
    ("c35", "自有車 VMT 轉換為 Robotaxi 的比例（閘門開啟時）2035", "比例", (0.10, 0.03, 0.005), "Assumed", "傑文斯路徑：價格低於自有車成本×便利溢價時，私人里程轉入；無證據"),
    ("sig35", "Tesla 占美國 Robotaxi 里程份額 2035", "比例", (0.45, 0.30, 0.12), "Assumed", "A：車隊成本與規模勝出；C：Waymo、Zoox、中國廠與 Uber 平台分散"),
    ("o35", "車主車輛（unsupervised FSD）占 Tesla 網路里程 2035", "比例", (0.50, 0.30, 0.10), "Assumed", "Tesla 稱約 2027 Q2 開放車主車輛加入（法說）"),
    ("intl35", "海外里程 ÷ 美國里程 2035（2028 起）", "倍", (1.0, 0.5, 0.1), "Assumed", "FSD 歐洲 9 國核准（監督版）；中國資料不出境"),
    ("w", "情境權重", "權重", (0.25, 0.50, 0.25), "Assumed", "DCF 規格預設 25／50／25"),
]
IN2 = [
    ("sec", "4. 每英里經濟（Tesla 自有車隊）"),
    ("take", "車主車輛里程的 Tesla 抽成", "比例", 0.25, 0.30, 0.50, "Assumed", "既有 Tesla 模型 2035 情境 75%（Tesla）／25%（車主）作上緣對照", "區間"),
    ("tsla_opx_tgt", "Tesla 法說每英里營運成本目標（對照）", "US$/英里", None, 0.30, None, "Interested-party", S["q3call"], "只作對照，不進模型；v0.2 營運成本讀線（見第 8 節），是否含折舊未明，有推升估值誘因"),
    ("veh27", "單車取得成本 2027（Cybercab 與 Model Y 混合）", "US$/台", 28000, 32000, 38000, "Assumed", "Cybercab 個人版售價 29,990（法說後媒體）、Musk「低於 3 萬」為 Interested-party", "區間"),
    ("veh35", "單車取得成本 2035", "US$/台", 20000, 25000, 30000, "Assumed", "學習曲線；既有 Tesla 模型 Cybercab 23,000＋套件 2,000", "區間"),
    ("util_k", "稼動率乘數（對線每車每日付費英里 × 365 的調整；基準 1.0）", "倍", 0.8, 1.0, 1.2, "Assumed", "v0.2 每車年付費英里＝線 Tesla 車隊加權每車每日付費英里 × 365 × 本乘數；線的稼動率是承重 Assumed／Derived（對帳 1 工程結論）", "區間；只在 Sens 動"),
    ("life", "車輛經濟壽命", "年", 5, 6, 8, "Assumed", "既有 Tesla 模型 6 年", "區間"),
    ("sga", "平台、研發分攤、管銷 ÷ 營收（不含訓練算力）", "比例", 0.10, 0.15, 0.25, "Assumed", "訓練算力支出屬算力收支（只讀，本版未取得），不在此列", "區間"),
    ("tax", "稅率", "比例", 0.15, 0.21, 0.25, "Assumed", "美國聯邦 21%；分部虧損可抵 Tesla 合併所得", "區間"),
    ("wc", "營運資金增加 ÷ 營收增量", "比例", -0.05, 0.0, 0.05, "Assumed", "乘客預付、無應收；近零", "區間"),
    ("sec", "5. 折現率（DCF 規格；Tesla 幾無有息負債，視為全權益）"),
    ("rf", "美國 10 年期公債殖利率（2026-11-13）", "比例", None, 0.0515, None, "Verified", S["te10y"], "—"),
    ("beta", "β（Tesla，Blume 調整）", "倍", 1.6, 2.0, 2.4, "Assumed", "兩個來源的 Blume 調整值本版未取得；以 Tesla 長期高 β 為區間", "區間；待補兩來源"),
    ("erp", "股權風險溢酬", "比例", 0.04, 0.05, 0.06, "Assumed", "DCF 規格預設 5%（敏感度 4–6%）", "—"),
    ("g_cap", "永續成長率上限", "比例", 0.02, 0.03, 0.04, "Assumed", "DCF 規格：g＝MIN(3%, rf)", "—"),
    ("ronic_sp", "RONIC − WACC", "比例", None, 0.0, 0.05, "Assumed", "DCF 規格預設 RONIC＝WACC", "—"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,11,16)", None, "Assumed", "成品日", "—"),
    ("sec", "6. 對照（不進模型）"),
    ("px", "TSLA 收盤價（2026-11-10，逾期 3 個交易日）", "US$", None, 379.78, None, "Verified", S["hey"], "只作對照，不進模型"),
    ("mcap", "TSLA 市值（2026-11-10）", "十億美元", None, 1500, None, "Verified", S["hey"], "只作對照，不進模型"),
    ("waymo_val", "Waymo 最近一輪融資估值（2026-02，160 億美元輪）", "十億美元", None, 126, None, "Verified", S["fortune"], "只作對照，不進模型"),
    ("waymo_wk", "Waymo 每週付費趟數（2026-02 宣布時）", "萬趟/週", None, 40, None, "Interested-party", S["fortune"], "Waymo 自揭里程碑；只作對照，不進模型"),
]

def line_in_rows():
    """In 頁「線 1a v0.2 讀入」：(鍵, 標籤, 單位, 值, 出處, 說明)；值由 load_line() 從線成品讀入（不手打）。
    價格、營運成本、稼動率：逐年 2026–2035 × 三情境；車隊權重 w_MY：錨點年 × 三情境（只作對照，不進模型）。"""
    fn = LINE["file"]
    rows = []
    anc = lambda y: y in LINE_ANCH

    def how(y):
        if anc(y):
            return "錨點"
        if y < 2032:
            return f"{y - 1}／{y + 1} 錨點的幾何平均（奇數年）"
        return "2032 錨點持平（線的權威只到 2032；趨勢版列敏感度）"

    for kind, key, lab, unit in (("price", "lp", "線 Tesla 車隊加權價格", "US$/英里"),
                                 ("opx", "lx", "線 Tesla 車隊加權營運成本（第 3–8 項，不含車輛攤提與融資）", "US$/英里"),
                                 ("pmd", "lm", "線 Tesla 車隊加權每車每日付費英里", "英里/日")):
        for y in range(2026, 2036):
            for s in SCN:
                ay = y if anc(y) else None
                if ay:
                    cc = LINE["cells"][(s, ay)]
                    if kind == "price":
                        src = f"{fn}｜成本輸入!{cc['ic']}{PR_MY}（TY_PRICE）、{cc['ic']}{PR_CC}（TC_PRICE）× 車隊權重 w_MY 加權（{LINE_CASE_LAB[s]}欄）"
                    elif kind == "opx":
                        src = (f"{fn}｜每英里成本堆疊!{cc['sc']}{ST_MY + 2}:{cc['sc']}{ST_MY + 7}（Model Y 第 3–8 項）、"
                               f"{cc['sc']}{ST_CC + 2}:{cc['sc']}{ST_CC + 7}（Cybercab 第 3–8 項）× 車隊權重 w_MY 加權")
                    else:
                        src = f"{fn}｜每英里成本堆疊!{cc['sc']}{ST_MY + 15}（Model Y）、{cc['sc']}{ST_CC + 15}（Cybercab）× 車隊權重 w_MY 加權"
                else:
                    src = f"{fn}｜{how(y)}"
                rows.append((f"{key}_{y}_{s}", f"情境 {s}：{lab} {y}", unit, LINE[kind][s][y], src, how(y)))
    for y in LINE_ANCH:
        for s in SCN:
            cc = LINE["cells"][(s, y)]
            src = (f"{fn}｜每英里成本堆疊!{cc['fc']}{cc['fr']}（Tesla 合計車隊加權單位營運成本）；"
                   f"w_MY＝(該值 − Cybercab 第 1–8 項合計 {cc['sc']}{ST_CC}:{cc['sc']}{ST_CC + 7}) ÷ (Model Y 第 1–8 項合計 {cc['sc']}{ST_MY}:{cc['sc']}{ST_MY + 7} − Cybercab 合計)")
            rows.append((f"wmy_{y}_{s}", f"情境 {s}：線 Tesla 車隊權重 w_MY（Model Y 占比）{y}", "比例", LINE["w"][(s, y)], src, "只作對照，不進模型"))
    return rows


LINE_CASE_LAB = {"A": "低價（情境 A）", "B": "基準（情境 B）", "C": "高價（情境 C）"}
LINE_DRIVERS = ([(f"lp_{y}", f"線 Tesla 車隊加權價格 {y}", "US$/英里") for y in range(2026, 2036)]
                + [(f"lx_{y}", f"線 Tesla 車隊加權營運成本（第 3–8 項）{y}", "US$/英里") for y in range(2026, 2036)]
                + [(f"lm_{y}", f"線 Tesla 車隊加權每車每日付費英里 {y}", "英里/日") for y in range(2026, 2036)]
                + [(f"wmy_{y}", f"線 Tesla 車隊權重 w_MY {y}", "比例") for y in LINE_ANCH])


# ---------------------------------------------------------------- Python 引擎（與 xlsx 公式同式；Sens 與核對用）
def base_params():
    P = {}
    for row in IN + IN2:
        if row[0] == "sec":
            continue
        P[row[0]] = row[4]
    for k, lab, u, vals, tag, src in SCN_IN:
        for s, v in zip(SCN, vals):
            P[f"{k}_{s}"] = v
    for k, lab, unit, v, src, note in line_in_rows():
        P[k] = v
    P["trip_mi"] = P["rh_fare"] * (1 - P["rh_take"]) / P["drv_pm"]
    P["val_date"] = 46342  # Excel serial of 2026-11-16
    return P


def derived(P):
    D = {}
    D["rh26"] = P["lyft_rq"] * 4 * P["lyft_na"] / P["lyft_sh"] / 1000 * P["trip_mi"]  # 十億英里
    D["p_h"] = P["rh_fare"] / P["trip_mi"] * (1 + P["tip"])
    D["thr"] = P["p_own"] * P["cf"]
    D["wacc"] = P["rf"] + P["beta"] * P["erp"]
    D["g"] = min(P["g_cap"], P["rf"])
    D["ronic"] = D["wacc"] + P["ronic_sp"]
    D["frac0"] = (46387 - P["val_date"]) / 365  # 2026-12-31 serial 46387
    D["t26_mi"] = (P["mi_ytd"] + P["mi_day"] * 74) / 1e9
    D["rev26"] = D["t26_mi"] * P["p26"]
    return D


def run(P, s):
    D = derived(P)
    L = lambda t: 1 / (1 + math.exp(-P["lk"] * (t - P["lmid"])))
    out = {k: {} for k in ["rh", "ln", "s", "p", "ind", "sub", "gate", "conv", "rt", "sig", "tus", "intl", "tall", "o",
                           "own", "peer", "rev", "opx", "mpv", "veh", "fleet", "capex", "da", "opex", "sga", "ebit", "nopat",
                           "dwc", "fcff", "df", "pv"]}
    fleet_prev, rev_prev = P["fleet26"], D["rev26"]
    for t in DCF_YEARS:
        o = {}
        o["rh"] = D["rh26"] * (1 + P["rh_g"]) ** (t - 2026)
        o["ln"] = L(t) / L(2035)
        o["s"] = P[f"s35_{s}"] * o["ln"]
        o["p"] = P[f"lp_{t}_{s}"]
        o["ind"] = (D["p_h"] / o["p"]) ** P["eps"]
        o["sub"] = o["rh"] * o["s"] * o["ind"]
        o["gate"] = min(1, max(0, (1.5 * D["thr"] - o["p"]) / (0.5 * D["thr"])))
        o["conv"] = P["vmt"] * P["ldv_sh"] * P[f"c35_{s}"] * o["ln"] * o["gate"]
        o["rt"] = o["sub"] + o["conv"]
        o["sig"] = P["sig26"] + (P[f"sig35_{s}"] - P["sig26"]) * (t - 2026) / 9
        o["tus"] = o["rt"] * o["sig"]
        o["intl"] = P[f"intl35_{s}"] * max(0, t - 2027) / 8
        o["tall"] = o["tus"] * (1 + o["intl"])
        o["o"] = P[f"o35_{s}"] * (t - 2027) / 8
        o["own"] = o["tall"] * (1 - o["o"])
        o["peer"] = o["tall"] * o["o"]
        o["rev"] = o["own"] * o["p"] + o["peer"] * o["p"] * P["take"]
        o["opx"] = P[f"lx_{t}_{s}"]
        o["mpv"] = P[f"lm_{t}_{s}"] * 365 * P["util_k"]
        o["veh"] = P["veh27"] + (P["veh35"] - P["veh27"]) * (t - 2027) / 8
        o["fleet"] = o["own"] * 1e9 / o["mpv"]
        o["capex"] = (max(0, o["fleet"] - fleet_prev) + fleet_prev / P["life"]) * o["veh"] / 1e9
        o["da"] = o["fleet"] * o["veh"] / P["life"] / 1e9
        o["opex"] = o["own"] * o["opx"]
        o["sga"] = P["sga"] * o["rev"]
        o["ebit"] = o["rev"] - o["opex"] - o["da"] - o["sga"]
        o["nopat"] = o["ebit"] * (1 - P["tax"])
        o["dwc"] = P["wc"] * (o["rev"] - rev_prev)
        o["fcff"] = o["nopat"] + o["da"] - o["capex"] - o["dwc"]
        o["df"] = (1 + D["wacc"]) ** -(D["frac0"] + (t - 2027) + 0.5)
        o["pv"] = o["fcff"] * o["df"]
        for k, v in o.items():
            out[k][t] = v
        fleet_prev, rev_prev = o["fleet"], o["rev"]
    tv = out["nopat"][2035] * (1 + D["g"]) * (1 - D["g"] / D["ronic"]) / (D["wacc"] - D["g"])
    pv_tv = tv * (1 + D["wacc"]) ** -(D["frac0"] + 9)
    ev = sum(out["pv"].values()) + pv_tv
    return dict(out=out, tv=tv, pv_tv=pv_tv, ev=ev, pv_exp=sum(out["pv"].values()))


def value(P):
    r = {s: run(P, s)["ev"] for s in SCN}
    w = sum(P[f"w_{s}"] * r[s] for s in SCN)
    return w, r


# ---------------------------------------------------------------- 工作簿
def build():
    P = base_params()
    wb = Workbook()
    wsR = wb.active; wsR.title = "README"
    wsP = wb.create_sheet("滲透")
    wsM = wb.create_sheet("市場天花板與情境")
    wsU = wb.create_sheet("每英里經濟")
    wsV = wb.create_sheet("估值")
    wsT = wb.create_sheet("給Tesla模型")
    wsSrc = wb.create_sheet("資料來源比較")
    wsRec = wb.create_sheet("對帳")
    wsI = wb.create_sheet("In")
    wsS = wb.create_sheet("Sens")
    wsLog = wb.create_sheet("改版紀錄_v0.2")

    def dn(name, ref):
        wb.defined_names[name] = DefinedName(name, attr_text=ref)

    # ---------------- In
    wsI["A1"] = "In 頁：藍字＝輸入。基準欄 E 進主模型；D／F（敏感度低／高）只用於 Sens 頁（一次動一個參數），不進分部價值。情境差異只放在情境專屬列（鍵尾 _A／_B／_C）。"
    wsI["A2"] = f"Tesla Robotaxi 分部模型 {VER}｜評價日 2026-11-16｜單位：里程十億英里、金額十億美元（另註者除外）"
    hdr = ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "出處", "來源誘因／說明"]
    for j, h in enumerate(hdr, 1):
        c = wsI.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4
    rowof = {}

    def put_in(k, lab, unit, lo, base, hi, tag, src, note):
        nonlocal r
        wsI.cell(r, 1, k); wsI.cell(r, 2, lab); wsI.cell(r, 3, unit)
        for j, v in ((4, lo), (5, base), (6, hi)):
            if v is not None:
                c = wsI.cell(r, j, v)
                if not (isinstance(v, str) and v.startswith("=") and tag == "Derived"):
                    c.font = BLUE
        wsI.cell(r, 7, tag); wsI.cell(r, 8, src); wsI.cell(r, 9, note)
        dn(k, f"In!$E${r}")
        rowof[k] = r
        r += 1

    def put_sec(t):
        nonlocal r
        c = wsI.cell(r, 1, t); c.font = BOLD; c.fill = SEC
        r += 1

    for row in IN:
        if row[0] == "sec":
            put_sec(row[1]); continue
        put_in(*row)
    for k, lab, unit, vals, tag, src in SCN_IN:
        for s, v in zip(SCN, vals):
            t2 = tag if not (k == "opx35" and s != "A") else "Assumed"
            put_in(f"{k}_{s}", f"情境 {s}：{lab}", unit, None, v, None, t2, src,
                   "Tesla 管理層目標，有推升估值誘因" if t2 == "Interested-party" else "情境專屬（2035 終點）")
    for row in IN2:
        if row[0] == "sec":
            put_sec(row[1]); continue
        put_in(*row)
    # 衍生（In 頁內計算，非輸入）
    put_sec("7. 衍生值（公式，非輸入）")
    for k, lab, unit, f, src in [
        ("rh26", "美國叫車基準里程 2026", "十億英里", "=lyft_rq*4*lyft_na/lyft_sh/1000*trip_mi", "Lyft 年化趟數 × 北美占比 ÷ 份額 × 每趟英里"),
        ("p_h", "人類叫車每英里乘客實付（含小費）", "US$/英里", "=rh_fare/trip_mi*(1+tip)", "UberX 車資 ÷ 行程長度 ×（1＋小費）"),
        ("thr", "轉換閘門門檻價（自有車成本 × 便利溢價）", "US$/英里", "=p_own*cf", "—"),
        ("wacc", "WACC（＝ke，全權益）", "比例", "=rf+beta*erp", "CAPM"),
        ("g", "永續成長率", "比例", "=MIN(g_cap,rf)", "DCF 規格"),
        ("ronic", "RONIC", "比例", "=wacc+ronic_sp", "DCF 規格"),
        ("frac0", "評價日至 2026 年底（年）", "年", "=(DATE(2026,12,31)-val_date)/365", "期中慣例"),
        ("t26_mi", "Tesla 2026 付費里程（估）", "十億英里", "=(mi_ytd+mi_day*74)/1e9", "1/22–10/18 累計 ＋ 每日里程 × 74 日（10/19–12/31）"),
        ("rev26", "Tesla Robotaxi 2026 營收（估）", "十億美元", "=t26_mi*p26", "—"),
        ("sh_d", "TSLA 隱含股數（市值 ÷ 股價）", "十億股", "=mcap/px", "只作對照，不進模型"),
    ]:
        put_in(k, lab, unit, None, f, None, "Derived", src, "公式" if "只作對照" not in src else "只作對照，不進模型")
    # 線 1a v0.2 讀入
    put_sec("8. 線 1a v0.2 讀入（Derived；建置程式從線成品 xlsx 讀值，非手打。價格、營運成本、稼動率逐年 2026–2035；2026／2028／2030／2032 為線錨點）")
    for k, lab, unit, v, src, note in line_in_rows():
        put_in(k, lab, unit, None, v, None, "Derived", src, note)
    # 情境驅動對照
    r += 1
    c = wsI.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）"); c.font = BOLD; c.fill = SEC
    r += 1
    for j, h in ((1, "鍵"), (2, "項目"), (4, "情境 A（樂觀）"), (5, "情境 B（基準）"), (6, "情境 C（保守）"), (8, "引用")):
        wsI.cell(r, j, h).font = BOLD
    r += 1
    for k, lab, unit, vals, tag, src in SCN_IN:
        kk = "scn_w" if k == "w" else k
        wsI.cell(r, 1, kk); wsI.cell(r, 2, lab); wsI.cell(r, 3, unit)
        for j, s in zip((4, 5, 6), SCN):
            wsI.cell(r, j, f"=E{rowof[f'{k}_{s}']}")
        wsI.cell(r, 8, f"引用 In 頁 {k}_A／{k}_B／{k}_C；{tag}")
        r += 1
    for k, lab, unit in LINE_DRIVERS:
        wsI.cell(r, 1, k); wsI.cell(r, 2, lab); wsI.cell(r, 3, unit)
        for j, s in zip((4, 5, 6), SCN):
            wsI.cell(r, j, f"=E{rowof[f'{k}_{s}']}")
        wsI.cell(r, 8, f"引用 In 頁 {k}_A／{k}_B／{k}_C；Derived（線 1a v0.2 讀入）")
        r += 1
    for col, wd in zip("ABCDEFGHI", (12, 52, 12, 11, 13, 11, 15, 70, 40)):
        wsI.column_dimensions[col].width = wd

    # ---------------- 市場天花板與情境（引擎第一段）
    ycol = {y: CL(3 + i) for i, y in enumerate(YEARS)}
    R = {}  # (scn, key) -> row on its sheet

    def yhdr(ws, row, title):
        ws.cell(row, 1, title).font = BOLD
        ws.cell(row, 2, "單位／說明").font = BOLD
        for y in YEARS:
            c = ws.cell(row, 3 + y - 2026, y); c.font = BOLD; c.fill = HDR

    wsM["A1"] = "市場天花板與情境：美國叫車基準里程 × 取代率 × 價格誘發（傑文斯）＋ 自有車 VMT 轉換（價格閘門）→ Robotaxi 里程 → Tesla 份額 → Tesla Robotaxi 營收。2026 欄為觀察／校準值。"
    wsM["A2"] = "讀法：收支只作估值的一環；本頁回答「天花板多高、Tesla 能拿多少」。里程單位十億英里；營收十億美元。"
    rr = 4
    for s in SCN:
        c = wsM.cell(rr, 1, SCN_NAME[s]); c.font = BOLD; c.fill = SEC
        rr += 1
        yhdr(wsM, rr, "項目"); rr += 1
        spec = [
            ("rh", "美國叫車基準里程", "十億英里", lambda y, C: f"=rh26*(1+rh_g)^({y}-2026)"),
            ("ln", "S 曲線（2035＝1）", "—", lambda y, C: f"=(1/(1+EXP(-lk*({y}-lmid))))/(1/(1+EXP(-lk*(2035-lmid))))"),
            ("s", "Robotaxi 取代率（占基準里程）", "比例", lambda y, C: f"=s35_{s}*{C}{R[(s,'ln')]}"),
            ("p", "Tesla 實收每英里價格（讀線 1a v0.2 車隊加權）", "US$/英里", lambda y, C: f"=lp_{y}_{s}"),
            ("ind", "價格誘發倍數（人類叫車價 ÷ Robotaxi 價）^彈性", "倍", lambda y, C: f"=(p_h/{C}{R[(s,'p')]})^eps"),
            ("sub", "叫車取代＋誘發里程", "十億英里", lambda y, C: f"={C}{R[(s,'rh')]}*{C}{R[(s,'s')]}*{C}{R[(s,'ind')]}"),
            ("gate", "自有車轉換閘門（價格 vs 自有車成本 × 便利溢價）", "0–1", lambda y, C: f"=MIN(1,MAX(0,(1.5*thr-{C}{R[(s,'p')]})/(0.5*thr)))"),
            ("conv", "自有車 VMT 轉換里程", "十億英里", lambda y, C: f"=vmt*ldv_sh*c35_{s}*{C}{R[(s,'ln')]}*{C}{R[(s,'gate')]}"),
            ("rt", "美國 Robotaxi 總里程（全部營運商）", "十億英里", lambda y, C: f"={C}{R[(s,'sub')]}+{C}{R[(s,'conv')]}"),
            ("sig", "Tesla 份額", "比例", lambda y, C: f"=sig26+(sig35_{s}-sig26)*({y}-2026)/9"),
            ("tus", "Tesla 美國里程", "十億英里", lambda y, C: f"={C}{R[(s,'rt')]}*{C}{R[(s,'sig')]}"),
            ("intl", "海外 ÷ 美國", "倍", lambda y, C: f"=intl35_{s}*MAX(0,{y}-2027)/8"),
            ("tall", "Tesla 全球網路付費里程", "十億英里", lambda y, C: f"={C}{R[(s,'tus')]}*(1+{C}{R[(s,'intl')]})"),
            ("o", "車主車輛占比", "比例", lambda y, C: f"=o35_{s}*({y}-2027)/8"),
            ("own", "Tesla 自有車隊付費里程", "十億英里", lambda y, C: f"={C}{R[(s,'tall')]}*(1-{C}{R[(s,'o')]})"),
            ("peer", "車主車輛付費里程", "十億英里", lambda y, C: f"={C}{R[(s,'tall')]}*{C}{R[(s,'o')]}"),
            ("rev", "Tesla Robotaxi 營收（自有全額＋車主抽成）", "十億美元", lambda y, C: f"={C}{R[(s,'own')]}*{C}{R[(s,'p')]}+{C}{R[(s,'peer')]}*{C}{R[(s,'p')]}*take"),
            ("pen", "Robotaxi 里程 ÷ 叫車基準里程（滲透倍數）", "倍", lambda y, C: f"={C}{R[(s,'rt')]}/{C}{R[(s,'rh')]}"),
            ("pen_vmt", "Robotaxi 里程 ÷ 小客車 VMT", "比例", lambda y, C: f"={C}{R[(s,'rt')]}/(vmt*ldv_sh)"),
        ]
        for key, lab, unit, fn in spec:
            R[(s, key)] = rr
            wsM.cell(rr, 1, lab); wsM.cell(rr, 2, unit)
            for y in DCF_YEARS:
                wsM.cell(rr, 3 + y - 2026, fn(y, ycol[y]))
            rr += 1
        # 2026 校準欄
        wsM.cell(R[(s, "rh")], 3, "=rh26")
        wsM.cell(R[(s, "p")], 3, "=p26")
        wsM.cell(R[(s, "sig")], 3, "=sig26")
        wsM.cell(R[(s, "tall")], 3, "=t26_mi")
        wsM.cell(R[(s, "own")], 3, "=t26_mi")
        wsM.cell(R[(s, "rev")], 3, "=rev26")
        rr += 1
    msum = rr
    wsM.cell(rr, 1, "情境比較：2030 與 2035").font = BOLD; rr += 1
    for j, h in enumerate(["項目", "年", "情境 A", "情境 B", "情境 C", "加權"], 1):
        wsM.cell(rr, j, h).font = BOLD
    rr += 1
    for key, lab in (("rt", "美國 Robotaxi 總里程（十億英里）"), ("pen", "Robotaxi ÷ 叫車基準里程（倍）"), ("pen_vmt", "Robotaxi ÷ 小客車 VMT"),
                     ("tall", "Tesla 全球付費里程（十億英里）"), ("rev", "Tesla Robotaxi 營收（十億美元）")):
        for y in (2030, 2035):
            wsM.cell(rr, 1, lab); wsM.cell(rr, 2, y)
            for j, s in enumerate(SCN):
                wsM.cell(rr, 3 + j, f"={ycol[y]}{R[(s,key)]}")
            wsM.cell(rr, 6, f"=w_A*C{rr}+w_B*D{rr}+w_C*E{rr}")
            rr += 1
    wsM.column_dimensions["A"].width = 46; wsM.column_dimensions["B"].width = 14

    # ---------------- 每英里經濟（引擎第二段）
    wsU["A1"] = "每英里經濟：價格 vs 每英里全成本（營運成本＋車輛折舊＋平台管銷分攤）；價格低於全成本＝補貼訊號。之後為車隊、資本支出與 EBIT。"
    wsU["A2"] = "口徑與 waymo 模型相同：每付費英里；全成本不含訓練算力（算力收支權威，本版未取得）。"
    ru = 4
    # 2026 觀察：Austin 計價 vs 全成本
    wsU.cell(ru, 1, "2026 觀察：價格 vs 每英里全成本（線 v0.2 2026 基準成本）").font = BOLD; ru += 1
    for j, h in enumerate(["項目", "單位", "短程 2 英里", "中程 5 英里", "長程 10 英里", "說明"], 1):
        wsU.cell(ru, j, h).font = BOLD
    ru += 1
    obs = ru
    wsU.cell(ru, 1, "Austin 現行牌價每英里（3＋1.40×英里）"); wsU.cell(ru, 2, "US$/英里")
    for j, d in enumerate((2, 5, 10)):
        wsU.cell(ru, 3 + j, f"=(3+1.4*{d})/{d}")
    wsU.cell(ru, 6, "牌價，未含促銷；見滲透頁價格區塊"); ru += 1
    wsU.cell(ru, 1, "每英里全成本（2026 線基準情境 B：營運＋以 2027 單車成本折舊＋管銷分攤）"); wsU.cell(ru, 2, "US$/英里")
    for j in range(3):
        cc = CL(3 + j)
        wsU.cell(ru, 3 + j, f"=lx_2026_B+veh27/life/(lm_2026_B*365*util_k)+sga*{cc}{ru-1}")
    wsU.cell(ru, 6, "營運成本與稼動率讀線 v0.2（2026 基準：每車每日付費英里約 45，安全員與遠端監看占大宗）；單車成本與壽命仍為 Assumed"); ru += 1
    wsU.cell(ru, 1, "價格 − 全成本"); wsU.cell(ru, 2, "US$/英里")
    for j in range(3):
        cc = CL(3 + j)
        wsU.cell(ru, 3 + j, f"={cc}{ru-2}-{cc}{ru-1}")
    ru += 1
    wsU.cell(ru, 1, "是否低於成本（補貼訊號）"); wsU.cell(ru, 2, "判定")
    for j in range(3):
        cc = CL(3 + j)
        wsU.cell(ru, 3 + j, f'=IF({cc}{ru-1}<0,"低於成本（補貼）","高於成本")')
    ru += 1
    wsU.cell(ru, 1, "對照：Cybercab 單趟實付 1.41/英里（19.4 英里）− 全成本"); wsU.cell(ru, 2, "US$/英里")
    wsU.cell(ru, 5, f"=1.41-E{obs+1}"); wsU.cell(ru, 6, S["cc_ride"]); ru += 2

    for s in SCN:
        c = wsU.cell(ru, 1, SCN_NAME[s]); c.font = BOLD; c.fill = SEC; ru += 1
        yhdr(wsU, ru, "項目"); ru += 1
        M = lambda key: R[(s, key)]
        spec = [
            ("pp", "價格（實收每英里）", "US$/英里", lambda y, C: f"=市場天花板與情境!{C}{M('p')}"),
            ("opx", "營運成本（不含折舊；讀線 v0.2 第 3–8 項車隊加權）", "US$/英里", lambda y, C: f"=lx_{y}_{s}"),
            ("mpv", "每車年付費英里（線每車每日付費英里 × 365 × 稼動率乘數）", "英里/年", lambda y, C: f"=lm_{y}_{s}*365*util_k"),
            ("veh", "單車取得成本", "US$/台", lambda y, C: f"=veh27+(veh35-veh27)*({y}-2027)/8"),
            ("dpm", "車輛折舊每英里", "US$/英里", lambda y, C: f"={C}{R[(s,'veh')]}/life/{C}{R[(s,'mpv')]}"),
            ("spm", "平台管銷分攤每英里", "US$/英里", lambda y, C: f"=sga*{C}{R[(s,'pp')]}"),
            ("fc", "每英里全成本", "US$/英里", lambda y, C: f"={C}{R[(s,'opx')]}+{C}{R[(s,'dpm')]}+{C}{R[(s,'spm')]}"),
            ("mg", "價格 − 全成本", "US$/英里", lambda y, C: f"={C}{R[(s,'pp')]}-{C}{R[(s,'fc')]}"),
            ("flag", "是否低於成本（補貼訊號）", "判定", lambda y, C: f'=IF({C}{R[(s,"mg")]}<0,"低於成本","高於成本")'),
            ("vs_h", "價格 ÷ 人類叫車每英里（含小費）", "比例", lambda y, C: f"={C}{R[(s,'pp')]}/p_h"),
            ("vs_own", "價格 ÷ 自有車每英里全成本（AAA）", "倍", lambda y, C: f"={C}{R[(s,'pp')]}/p_own"),
            ("fleet", "Tesla 自有車隊（年底）", "台", lambda y, C: f"=市場天花板與情境!{C}{M('own')}*1e9/{C}{R[(s,'mpv')]}"),
            ("capex", "車隊資本支出（新增＋汰換）", "十億美元", lambda y, C: f"=(MAX(0,{C}{R[(s,'fleet')]}-{CL(ord(C)-65)}{R[(s,'fleet')]})+{CL(ord(C)-65)}{R[(s,'fleet')]}/life)*{C}{R[(s,'veh')]}/1e9"),
            ("da", "折舊", "十億美元", lambda y, C: f"={C}{R[(s,'fleet')]}*{C}{R[(s,'veh')]}/life/1e9"),
            ("opex", "營運成本", "十億美元", lambda y, C: f"=市場天花板與情境!{C}{M('own')}*{C}{R[(s,'opx')]}"),
            ("sgaa", "平台、研發、管銷", "十億美元", lambda y, C: f"=sga*市場天花板與情境!{C}{M('rev')}"),
            ("ebit", "分部 EBIT", "十億美元", lambda y, C: f"=市場天花板與情境!{C}{M('rev')}-{C}{R[(s,'opex')]}-{C}{R[(s,'da')]}-{C}{R[(s,'sgaa')]}"),
            ("em", "EBIT 率", "比例", lambda y, C: f"=IF(市場天花板與情境!{C}{M('rev')}=0,0,{C}{R[(s,'ebit')]}/市場天花板與情境!{C}{M('rev')})"),
        ]
        for key, lab, unit, fn in spec:
            R[(s, key)] = ru
            wsU.cell(ru, 1, lab); wsU.cell(ru, 2, unit)
            ru += 1
        for key, lab, unit, fn in spec:
            for y in DCF_YEARS:
                wsU.cell(R[(s, key)], 3 + y - 2026, fn(y, ycol[y]))
        wsU.cell(R[(s, "fleet")], 3, "=fleet26")
        wsU.cell(R[(s, "pp")], 3, "=p26")
        ru += 1
    wsU.column_dimensions["A"].width = 46; wsU.column_dimensions["B"].width = 12

    # ---------------- 估值（FCFF DCF）
    wsV["A1"] = "估值：分部 DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加；明確期 2027–2035；終值＝NOPAT(2036) ×（1 − g ÷ RONIC）÷（WACC − g）；期中慣例）。評等：未上市（Tesla 分部）。"
    wsV["A2"] = "分部價值＝企業價值（分部不配置現金與負債；Tesla 合併淨現金在 Tesla 模型 SOTP 處理）。2026 第 4 季已發生的車隊支出不另計（量小）。"
    rv = 4
    wsV.cell(rv, 1, "折現率").font = BOLD; rv += 1
    for lab, f in (("rf", "=rf"), ("β", "=beta"), ("ERP", "=erp"), ("WACC（＝ke）", "=wacc"), ("永續成長率 g", "=g"), ("RONIC", "=ronic"), ("評價日至 2026 年底（年）", "=frac0")):
        wsV.cell(rv, 1, lab); wsV.cell(rv, 3, f); rv += 1
    rv += 1
    VR = {}
    for s in SCN:
        c = wsV.cell(rv, 1, SCN_NAME[s]); c.font = BOLD; c.fill = SEC; rv += 1
        yhdr(wsV, rv, "項目"); rv += 1
        U = lambda key: R[(s, key)]
        spec = [
            ("ebit", "EBIT", lambda y, C: f"=每英里經濟!{C}{U('ebit')}"),
            ("nopat", "NOPAT", lambda y, C: f"={C}{VR[(s,'ebit')]}*(1-tax)"),
            ("da", "＋折舊", lambda y, C: f"=每英里經濟!{C}{U('da')}"),
            ("capex", "− 資本支出", lambda y, C: f"=每英里經濟!{C}{U('capex')}"),
            ("dwc", "− 營運資金增加", lambda y, C: f"=wc*(市場天花板與情境!{C}{R[(s,'rev')]}-市場天花板與情境!{CL(ord(C)-65)}{R[(s,'rev')]})"),
            ("fcff", "FCFF", lambda y, C: f"={C}{VR[(s,'nopat')]}+{C}{VR[(s,'da')]}-{C}{VR[(s,'capex')]}-{C}{VR[(s,'dwc')]}"),
            ("chk", "檢查：各列加總 − FCFF（應為 0）", lambda y, C: f"=({C}{VR[(s,'nopat')]}+{C}{VR[(s,'da')]}-{C}{VR[(s,'capex')]}-{C}{VR[(s,'dwc')]})-{C}{VR[(s,'fcff')]}"),
            ("df", "折現因子", lambda y, C: f"=(1+wacc)^-(frac0+({y}-2027)+0.5)"),
            ("pv", "FCFF 現值", lambda y, C: f"={C}{VR[(s,'fcff')]}*{C}{VR[(s,'df')]}"),
        ]
        for key, lab, fn in spec:
            VR[(s, key)] = rv; wsV.cell(rv, 1, lab); wsV.cell(rv, 2, "十億美元" if key not in ("df",) else "—"); rv += 1
        for key, lab, fn in spec:
            for y in DCF_YEARS:
                wsV.cell(VR[(s, key)], 3 + y - 2026, fn(y, ycol[y]))
        L_ = ycol[2035]
        for key, lab, f in (
            ("pvx", "明確期現值合計", f"=SUM(D{VR[(s,'pv')]}:{L_}{VR[(s,'pv')]})"),
            ("tv", "終值（2035 年底）", f"={L_}{VR[(s,'nopat')]}*(1+g)*(1-g/ronic)/(wacc-g)"),
            ("pvtv", "終值現值", f"=C{rv+1}*(1+wacc)^-(frac0+9)"),
            ("ev", "分部價值（EV，評價日）", f"=C{rv}+C{rv+2}"),
            ("tvsh", "終值占 EV（>80% 標示）", f"=IF(C{rv+3}=0,0,C{rv+2}/C{rv+3})"),
            ("tvflag", "終值占比旗標", f'=IF(C{rv+4}>0.8,"終值占比超過 80%","—")'),
            ("v12", "推到 12 個月後（× (1＋WACC)）", f"=C{rv+3}*(1+wacc)"),
            ("ps", "每股（÷ 隱含股數，只作對照）", f"=C{rv+3}/sh_d"),
        ):
            VR[(s, key)] = rv; wsV.cell(rv, 1, lab); wsV.cell(rv, 3, f); rv += 1
        rv += 1
    rv += 1
    wsV.cell(rv, 1, "分部價值彙總（十億美元）").font = BOLD; rv += 1
    for j, h in enumerate(["項目", "", "情境 A", "情境 B", "情境 C", "加權（25／50／25）"], 1):
        wsV.cell(rv, j, h).font = BOLD
    rv += 1
    VS = {}
    for key, lab in (("pvx", "明確期現值"), ("pvtv", "終值現值"), ("ev", "分部價值（EV）"), ("tvsh", "終值占 EV"), ("v12", "推到 12 個月後"), ("ps", "每股（US$，只作對照）")):
        wsV.cell(rv, 1, lab)
        for j, s in enumerate(SCN):
            wsV.cell(rv, 3 + j, f"=C{VR[(s,key)]}")
        wsV.cell(rv, 6, f"=w_A*C{rv}+w_B*D{rv}+w_C*E{rv}")
        VS[key] = rv; rv += 1
    wsV.cell(rv, 1, "分部價值 ÷ TSLA 市值（只作對照）")
    for j in range(4):
        cc = CL(3 + j)
        wsV.cell(rv, 3 + j, f"={cc}{VS['ev']}/mcap")
    VS["pct"] = rv; rv += 2
    wsV.cell(rv, 1, "評等").font = BOLD; wsV.cell(rv, 3, "未上市（Tesla 分部；不給目標價與評等，只給分部價值區間）"); rv += 2
    # Waymo 融資估值對照
    wsV.cell(rv, 1, "對照：Waymo 最近一輪融資估值（不進結論）").font = BOLD; rv += 1
    wsV.cell(rv, 1, "Waymo 估值（2026-02）"); wsV.cell(rv, 3, "=waymo_val"); wsV.cell(rv, 4, S["fortune"]); rv += 1
    wsV.cell(rv, 1, "Waymo 估值 ÷ 年化付費趟數（US$/年化趟）"); wsV.cell(rv, 3, "=waymo_val*1e9/(waymo_wk*1e4*52)"); rv += 1
    wsV.cell(rv, 1, "Tesla 年化趟數（估：每日里程 × 365 ÷ 5 英里/趟）"); wsV.cell(rv, 3, "=mi_day*365/5"); wsV.cell(rv, 4, "5 英里/趟為 Assumed（Tesla 未揭露趟數）"); rv += 1
    wsV.cell(rv, 1, "以 Waymo 每年化趟估值倍數套用 Tesla（十億美元）"); wsV.cell(rv, 3, f"=C{rv-2}*C{rv-1}/1e9"); rv += 1
    wsV.cell(rv, 1, "讀法：Waymo 輪次估值含 Alphabet 策略溢價與期權價值；趟數倍數只顯示「現況規模」口徑下 Tesla 的對照值，不作依據。"); rv += 1
    wsV.column_dimensions["A"].width = 40

    # ---------------- 給 Tesla 模型
    wsT["A1"] = "給 Tesla 模型（physical_ai/tesla/ 讀取）：Robotaxi 分部價值，十億美元，評價日 2026-11-16 現值；營收百萬美元。具名儲存格：RT_VAL_LO／BASE／HI（區間）、RT_VAL_A／B／C（情境價值）、RT_REV_<A/B/C>_<2027–2035>（Tesla Robotaxi 營收，百萬美元）。"
    rows = [
        ("RT_VAL_LO", "Robotaxi 分部價值 低（＝情境 C）", f"=估值!E{VS['ev']}"),
        ("RT_VAL_BASE", "Robotaxi 分部價值 基準（＝情境加權 25／50／25）", f"=估值!F{VS['ev']}"),
        ("RT_VAL_HI", "Robotaxi 分部價值 高（＝情境 A）", f"=估值!C{VS['ev']}"),
        ("RT_VAL_B", "對照：情境 B 單一情境（＝RT_VAL_B）", f"=估值!D{VS['ev']}"),
        ("RT_PS_BASE", "基準每股（÷ 隱含股數，只作對照）", f"=估值!F{VS['ps']}"),
        ("RT_REV2030_B", "情境 B 2030 營收（十億美元）", f"=市場天花板與情境!{ycol[2030]}{R[('B','rev')]}"),
        ("RT_REV2035_B", "情境 B 2035 營收（十億美元）", f"=市場天花板與情境!{ycol[2035]}{R[('B','rev')]}"),
        ("RT_VAL_A", "Robotaxi 分部價值 情境 A（十億美元）", f"=估值!C{VS['ev']}"),
        ("RT_VAL_C", "Robotaxi 分部價值 情境 C（十億美元）", f"=估值!E{VS['ev']}"),
    ]
    for s_ in SCN:
        for y in DCF_YEARS:
            rows.append((f"RT_REV_{s_}_{y}", f"情境 {s_} {y} Tesla Robotaxi 營收（百萬美元）", f"=市場天花板與情境!{ycol[y]}{R[(s_,'rev')]}*1000"))
    wsT.cell(3, 1, "具名儲存格").font = BOLD; wsT.cell(3, 2, "項目").font = BOLD; wsT.cell(3, 3, "值（十億美元／百萬美元／US$）").font = BOLD
    for i, (nm, lab, f) in enumerate(rows):
        rr_ = 4 + i
        wsT.cell(rr_, 1, nm); wsT.cell(rr_, 2, lab); wsT.cell(rr_, 3, f)
        dn(nm, f"給Tesla模型!$C${rr_}")
    nr = 4 + len(rows) + 1
    wsT.cell(nr, 1, "口徑：分部 EV（不含 Tesla 合併淨現金）；不含 FSD 個人車主訂閱（屬 FSD 軟體分部）；車主車輛加入 Robotaxi 網路的抽成在本分部；訓練算力支出未扣（算力收支只讀、未取得）→ Tesla SOTP 若另扣 AI 訓練成本，避免重複。")
    wsT.cell(nr + 1, 1, "標記：Derived（價格、營運成本、稼動率讀線 1a v0.2；其餘情境驅動為 Assumed）。Tesla 模型 SOTP 讀 RT_VAL_A／B／C；DCF 腿的 Robotaxi 營收讀 RT_REV_<情境>_<年>（FY2027–FY2035），不再讀原檔引擎 × rt_k。")
    wsT.cell(nr + 2, 1, "RT_REV 口徑：Tesla 全球網路營收（自有車隊全額＋車主車輛抽成），名目 US$ 百萬；成本與毛利不在此提供（Tesla 模型沿用其 Robotaxi 毛利處理，見該模型 README 限制）。")
    wsT.column_dimensions["A"].width = 16; wsT.column_dimensions["B"].width = 48; wsT.column_dimensions["C"].width = 16

    return wb, P, R, VR, VS, rowof, (wsR, wsP, wsSrc, wsRec, wsS, wsI, wsM, wsU, wsV, wsLog), ycol


def write_pen(wsP, R, ycol):
    """滲透頁：最上方價格區塊、城市表、時間序列、與 Waymo 並排口徑。"""
    r = 1
    wsP.cell(r, 1, "滲透（主頁）：Tesla Robotaxi 的滲透狀況。最上方為核心指標「每英里實收價格」（決定傑文斯效應能否實現），其後為城市、車隊、里程、份額與時間序列。口徑與 physical_ai/waymo/ 相同。").font = BOLD
    r += 2
    # ---- A. 價格
    wsP.cell(r, 1, "A. 價格：每英里實收價格（按行程長度：短 2／中 5／長 10 英里）").font = BOLD; wsP.cell(r, 1).fill = SEC; r += 1
    hd = ["日期", "城市", "計價", "基本費 US$", "每英里 US$", "短程 2 英里 $/mi", "中程 5 英里 $/mi", "長程 10 英里 $/mi", "標記", "出處", "來源誘因／說明"]
    for j, h in enumerate(hd, 1):
        c = wsP.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    hist = [
        ("2025-06-22", "Austin", "固定每趟", None, None, 4.20, "Verified", S["price_dtc"], "上線價（有安全員）；固定價"),
        ("2025-07 中", "Austin", "固定每趟", None, None, 6.90, "Verified", S["price_dtc"], "服務區擴大時調高"),
        ("2025-07-30", "Austin", "基本費＋每英里", 1.00, 1.00, None, "Interested-party", S["price_aol"], "app 25.7.10 起按距離；公式由社群推得，Tesla 未公告"),
        ("2026-03-07", "Austin", "基本費＋每英里", 3.25, 1.00, None, "Interested-party", S["price_aol"], "社群回報（Joe Tegtmeyer、Tesla Newswire）；先 3.00 再 3.25"),
        ("2026-03-12", "Austin", "基本費＋每英里", 3.00, 1.40, None, "Interested-party", S["price_tp"], "每英里 +40%（另見 The Driven 2026-03-17）；本模型現行牌價"),
    ]
    PR = {}
    for d, city, kind, base, pm, flat, tag, src, note in hist:
        wsP.cell(r, 1, d); wsP.cell(r, 2, city); wsP.cell(r, 3, kind)
        if flat is not None:
            wsP.cell(r, 4, flat).font = BLUE
            for j, mi in enumerate((2, 5, 10)):
                wsP.cell(r, 6 + j, f"=D{r}/{mi}")
        else:
            wsP.cell(r, 4, base).font = BLUE; wsP.cell(r, 5, pm).font = BLUE
            for j, mi in enumerate((2, 5, 10)):
                wsP.cell(r, 6 + j, f"=(D{r}+E{r}*{mi})/{mi}")
        wsP.cell(r, 9, tag); wsP.cell(r, 10, src); wsP.cell(r, 11, note)
        PR[d] = r; r += 1
    # 實付觀察
    obs = [
        ("2025-11～2026-01", "SF Bay Area", "Obi 實測平均（有安全員）", 1.99 * 1.609344, "Verified", S["obi"], "US$1.99/公里 × 1.609＝每英里；平均每趟 8.17；ETA 15.32 分"),
        ("2026-09-04", "Austin", "Model Y 單趟實付（vs Uber 21.00）", None, "Interested-party", S["cc_ride"], "16.99，距離未揭露 → 每英里未取得；Tesla 支持者貼文"),
        ("2026-09-09", "Austin", "Cybercab 單趟實付（19.4 英里）", 1.41, "Interested-party", S["cc_ride"], "單一樣本；長程"),
        ("2026-02", "Austin", "Jefferies 實測：Tesla 對 UberX 折扣約 60%", None, "Interested-party", S["jef"], "15 趟樣本；賣方（有覆蓋利益）"),
    ]
    wsP.cell(r, 1, "實付觀察（非牌價）").font = BOLD; r += 1
    for d, city, kind, v, tag, src, note in obs:
        wsP.cell(r, 1, d); wsP.cell(r, 2, city); wsP.cell(r, 3, kind)
        if v is not None:
            c = wsP.cell(r, 7, v); c.font = BLUE
        else:
            wsP.cell(r, 7, "未取得")
        wsP.cell(r, 9, tag); wsP.cell(r, 10, src); wsP.cell(r, 11, note)
        PR[kind] = r; r += 1
    # 錨
    wsP.cell(r, 1, "比較錨").font = BOLD; r += 1
    cur = PR["2026-03-12"]
    anchors = [
        ("錨 1：人類叫車 UberX 每英里（含小費）", "=p_h", "Derived", "Gridwise 車資 24.30 ÷ 行程 7.1 英里 ×（1＋小費 8%）", "所有長度同值（無距離拆分資料）"),
        ("錨 1b：Lyft 每英里（含小費，假設同行程長度）", "=19.76/trip_mi*(1+tip)", "Derived", S["gridwise"] + "（Lyft 每趟 19.76）", "Lyft 行程長度未揭露"),
        ("錨 1c：Bay Area UberX 每趟 17.47（Obi）", 17.47, "Verified", S["tc_obi"], "每英里未揭露；只作每趟對照"),
        ("錨 2：自有車每英里全成本（AAA 2025，15,000 英里/年平均）", "=p_own", "Verified", S["aaa"], "0.77"),
        ("錨 2b：自有小轎車每英里全成本（AAA 2025）", 0.5587, "Verified", S["aaa"], "0.5587"),
    ]
    AR = {}
    for lab, v, tag, src, note in anchors:
        wsP.cell(r, 1, lab)
        for j in range(3):
            if lab.startswith("錨 1c"):
                break
            c = wsP.cell(r, 6 + j, v)
            if not isinstance(v, str):
                c.font = BLUE
        wsP.cell(r, 9, tag); wsP.cell(r, 10, src); wsP.cell(r, 11, note)
        AR[lab[:4]] = r; r += 1
    a1 = AR["錨 1："]; a2 = AR["錨 2："]
    for lab, f in (("現行牌價 ÷ 錨 1（UberX 每英里）", lambda C: f"={C}{cur}/{C}{a1}"),
                   ("現行牌價 ÷ 錨 2（自有車每英里）", lambda C: f"={C}{cur}/{C}{a2}"),
                   ("是否低於自有車成本（傑文斯門檻）", lambda C: f'=IF({C}{cur}<{C}{a2},"低於自有車成本","高於自有車成本")')):
        wsP.cell(r, 1, lab)
        for j in range(3):
            C = CL(6 + j)
            wsP.cell(r, 6 + j, f(C))
        wsP.cell(r, 9, "Derived"); r += 1
    wsP.cell(r, 1, "觀察到的價格彈性 Δln(付費里程)／Δln(價格)")
    wsP.cell(r, 7, "未取得")
    wsP.cell(r, 11, "2026-03 調價前後無同口徑付費里程（Tesla 只揭露累計里程；Q2 付費里程較 Q1 下降的媒體說法同時混有車隊縮減）→ 無法分離價格效果。模型用 In 頁 eps 0.3–1.2（Assumed）。")
    r += 2

    # ---- B. 城市表
    wsP.cell(r, 1, "B. 每城市現況（2026-11-16）").font = BOLD; wsP.cell(r, 1).fill = SEC; r += 1
    hd = ["城市", "州", "上線日", "狀態", "車隊（登記／觀察）", "服務區（平方英里）", "累計自駕里程（1/22–10/18）", "叫車市場份額", "等候時間（分）", "標記", "出處"]
    for j, h in enumerate(hd, 1):
        c = wsP.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    cities = [
        ("Austin", "TX", "2025-06-22", "無人（2026-01 移除安全員）", "德州名錄全州 1,302（MY 729＋Cybercab 573）", 440, 2118015, "未取得", "12–15（Austin 社群回報）", "Interested-party", S["geo"] + "；" + S["safety"]),
        ("Dallas", "TX", "2026-04-18", "無人；Cybercab 2026-12 起", "未取得（含在德州 1,302）", "未取得", 608130, "未取得", "未取得", "Interested-party", S["tracker_bn"] + "；" + S["safety"]),
        ("Houston", "TX", "2026-04-18", "無人；10 月服務區約擴 10 倍", "未取得", "未取得", 225964, "未取得", "未取得", "Interested-party", S["tracker_bn"]),
        ("Fort Worth", "TX", "2026-11-02", "無人", "約 24（GuruFocus）", "未取得", "未上線", "未取得", "未取得", "Interested-party", S["hey"]),
        ("San Antonio", "TX", "2026-11-09", "未揭露", "未取得", 275, "未上線", "未取得", "未取得", "Interested-party", S["geo"]),
        ("Miami", "FL", "2026-07-03", "無人", "約 5（GuruFocus）", "未取得", 46262, "未取得", "未取得", "Interested-party", S["hey"] + "；" + S["safety"]),
        ("Orlando", "FL", "2026-07-21", "無人（Q2 deck）", "未取得", "未取得", 61053, "未取得", "未取得", "Interested-party", S["q2deck"]),
        ("Tampa", "FL", "2026-07-21", "無人（Q2 deck）", "未取得", "未取得", 53117, "未取得", "未取得", "Interested-party", S["q2deck"]),
        ("SF Bay Area", "CA", "2025-07", "有安全員（CPUC TCP 包車許可，非 AV 服務）", "約 156–168（Obi 抽樣期）；2026-05 觀察 9", "未取得", "不列（非自駕）", "未取得", 15.32, "Verified", S["obi"] + "；" + S["cpuc_tsla"]),
    ]
    c0 = r
    for row in cities:
        for j, v in enumerate(row, 1):
            wsP.cell(r, j, v)
        r += 1
    wsP.cell(r, 1, "合計（自駕里程）"); wsP.cell(r, 7, f"=SUM(G{c0}:G{r-1})"); wsP.cell(r, 11, "Tesla 報告合計約 310 萬英里"); r += 1
    wsP.cell(r, 1, "已宣布未上線：Las Vegas、Phoenix、Jacksonville（法說 2026-10）"); r += 1
    wsP.cell(r, 1, "城市叫車份額：德州與佛州無 TNC 逐城趟數公開資料 → 份額未取得；可行替代：YipitData／Second Measure 刷卡面板（付費）。"); r += 2

    # ---- C. 時間序列
    wsP.cell(r, 1, "C. 全國時間序列（季；空白＝未取得）").font = BOLD; wsP.cell(r, 1).fill = SEC; r += 1
    hd = ["期末", "無人營運城市數", "總城市（含 Bay Area）", "車隊（無人，觀察）", "車隊（登記／總）", "累計付費里程（百萬）", "累計無人里程（百萬）", "每日里程（無人）", "Austin 牌價 5 英里 $/mi", "出處", "說明"]
    for j, h in enumerate(hd, 1):
        c = wsP.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    ts = [
        ("2025Q2", 0, 1, None, "約 10–20", None, None, None, f"=G{PR['2025-06-22']}", S["price_dtc"], "Austin 有安全員上線"),
        ("2025Q3", 0, 2, None, None, None, None, None, f"=G{PR['2025-07-30']}", S["price_aol"], "Bay Area（有駕駛）"),
        ("2025Q4", 0, 2, None, "約 500（Austin＋Bay Area，Musk Q4 法說）", None, None, None, f"=G{PR['2025-07-30']}", S["jef"], ""),
        ("2026Q1", 1, 2, None, None, None, None, None, f"=G{PR['2026-03-12']}", S["may"], "Austin 2026-01 移除安全員"),
        ("2026Q2", 6, 7, "20（7 日）／33（30 日），2026-05 下旬", "34（2026-05 下旬）；德州名錄 42（05-28）", 2.4, 0.38, None, f"=G{PR['2026-03-12']}", S["q2deck"] + "；" + S["may"] + "；" + S["ramp"], "Q2 deck：累計付費約 240 萬英里（含有監督）"),
        ("2026Q3", 6, 7, None, "德州名錄 >1,000（10 月下旬）", None, 2.0, None, f"=G{PR['2026-03-12']}", S["q3res"] + "；" + S["txreg"], "Q3 deck：累計無人里程近 200 萬；Cybercab 2026-09-03 開始載客"),
        ("2026-10-18", 6, 7, None, None, 2.5, 3.1, 80600, f"=G{PR['2026-03-12']}", S["safety"] + "；" + S["mpd"] + "；" + S["ramp"], "安全報告 310 萬自駕英里；付費「接近 250 萬」口徑不同（含監督、定義未明）"),
        ("2026-11-11", 8, 9, None, "德州名錄 1,302（MY 729＋Cybercab 573）", None, None, None, f"=G{PR['2026-03-12']}", S["txreg"], "Fort Worth、San Antonio 新增；Tesla 德州登記數首度超過 Waymo（1,297）"),
    ]
    for row in ts:
        for j, v in enumerate(row, 1):
            if v is not None:
                wsP.cell(r, j, v)
        r += 1
    wsP.cell(r, 1, "注意：累計付費里程（240 萬@6 月底、近 250 萬@10 月中）與累計自駕里程（310 萬@10/18）口徑不同且彼此難以相容；只作方向參考。Tesla 未揭露每週付費趟數。"); r += 1
    wsP.cell(r, 1, "每週付費趟數（估）＝每日里程 × 7 ÷ 每趟英里（5，Assumed）"); wsP.cell(r, 8, "=mi_day*7/5"); wsP.cell(r, 9, "Derived"); r += 2

    # ---- D. 與 Waymo 並排口徑
    wsP.cell(r, 1, "D. 與 Waymo 並排（同口徑；Waymo 權威數字見 physical_ai/waymo/）").font = BOLD; wsP.cell(r, 1).fill = SEC; r += 1
    hd = ["指標", "Tesla", "Waymo", "日期", "出處"]
    for j, h in enumerate(hd, 1):
        c = wsP.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    for row in (
        ("營運城市（無人）", 8, 15, "2026-11／2026-10", S["tracker_bn"] + "；" + S["fortune"]),
        ("德州 AV 登記車數", 1302, 1297, "2026-11-11", S["txreg"]),
        ("每週付費趟數", "未揭露（估約 11 萬，Derived）", "40 萬+（2026-02）", "—", S["fortune"]),
        ("Bay Area 每英里價格", "3.20（有安全員）", "9.21（5.72/公里）", "2025-11～2026-01", S["obi"]),
        ("Bay Area 等候時間（分）", 15.32, 5.74, "2025-11～2026-01", S["obi"]),
        ("Austin 服務區（平方英里）", 440, "未取得", "2026-11-09", S["geo"]),
    ):
        for j, v in enumerate(row, 1):
            wsP.cell(r, j, v)
        r += 1
    for col, wd in zip("ABCDEFGHIJK", (26, 14, 20, 22, 30, 16, 16, 16, 15, 60, 50)):
        wsP.column_dimensions[col].width = wd


def write_src(ws):
    ws["A1"] = "資料來源比較（Robotaxi 滲透；與 waymo 模型同表格式）"
    hd = ["來源", "涵蓋", "頻率", "可靠度", "建議用途", "標記", "網址／說明"]
    for j, h in enumerate(hd, 1):
        c = ws.cell(3, j, h); c.font = BOLD; c.fill = HDR
    rows = [
        ("Tesla 財報 Shareholder Update／法說", "城市、累計里程（付費、無人）、Cybercab 進度、成本目標", "季", "中：數字少、口徑會換（付費 vs 無人 vs 自駕）", "里程錨點、城市清單", "Interested-party", S["q3res"]),
        ("Tesla Robotaxi 安全報告", "逐城累計自駕里程、事故對比", "不定期（2026-10 首次）", "中：公司自編，未經第三方稽核", "逐城里程分配", "Interested-party", S["safety"]),
        ("robotaxitracker.com（Ethan McKanna）", "Tesla、Waymo、Zoox、Avride 車輛名錄、服務區地圖、7／30 日活躍車數、價格、等候時間、政府資料彙整", "近即時", "中低：社群提交＋等候時間遙測＋政府紀錄＋審核流程（Methodology 頁摘要）；頁面為 JS 渲染，本環境無法擷取明細與方法全文（403）；與 Tesla 揭露常有落差", "活躍車數、等候時間趨勢（交叉驗證後）", "Interested-party", S["rtt"] + "；個人經營，社群來源，非公司揭露"),
        ("EVwire／Texas AV Tracker（TxDMV 名錄解析）", "德州 AV 登記車數（VIN 解車型）", "日", "中高：政府名錄；但登記≠營運", "車隊上限與趨勢", "Derived", S["txreg"]),
        ("加州 CPUC AV 季報", "Waymo 等 AV 趟數、里程、空駛", "季", "高：監管申報；Tesla 在加州為 TCP 包車，不在 AV 報告內", "Waymo 份額；Tesla 不適用", "Verified", S["cpuc"]),
        ("加州 DMV（測試／部署許可、脫離報告）", "許可與測試里程", "年／即時", "高", "Tesla 是否取得無人部署許可", "Verified", "https://www.dmv.ca.gov/portal/vehicle-industry-services/autonomous-vehicles/"),
        ("NHTSA SGO 事故資料", "ADS 事故逐筆", "月", "高（但有遮蔽欄位）", "安全、停駛風險", "Verified", "https://www.nhtsa.gov/laws-regulations/standing-general-order-crash-reporting"),
        ("Obi 價格研究", "Bay Area 9.4 萬次模擬叫車：價格、ETA", "半年", "中高：第三方、大樣本；限 Bay Area", "每英里價格、等候時間對照", "Verified", S["obi"]),
        ("Gridwise 季報", "司機每英里收入、抽成、每趟車資", "季", "中：司機 app 樣本", "人類叫車價格錨", "Verified", S["gridwise"]),
        ("YipitData／Second Measure（付費）", "刷卡面板：Uber、Lyft、Waymo 份額與使用者行為", "月", "中高（付費；本版只取公開摘要）", "城市叫車份額（Tesla 份額待購買）", "Analogy", S["yipit"]),
        ("Uber／Lyft 法說", "趟數、GB、AV 合作車數", "季", "高（法定申報）", "叫車市場基準", "Verified", S["lyft"] + "；" + S["uber"]),
        ("社群 X 帳號（Sawyer Merritt、Whole Mars 等）", "每日里程、單趟價格", "即時", "低：Tesla 支持者、樣本小", "只作方向，需交叉驗證", "Interested-party", S["mpd"]),
        ("賣方實測（Jefferies 等）", "價格、等候、路線", "不定期", "中：樣本小、有覆蓋利益", "價格折扣方向", "Interested-party", S["jef"]),
    ]
    for i, row in enumerate(rows):
        for j, v in enumerate(row, 1):
            ws.cell(4 + i, j, v).alignment = WRAP
    for col, wd in zip("ABCDEFG", (30, 40, 12, 40, 28, 14, 60)):
        ws.column_dimensions[col].width = wd


def write_rec(ws, rowof):
    ws["A1"] = "對帳：本模型使用的「線層級」變數（v0.2：價格、營運成本、稼動率讀 physical_ai/lines/robotaxi 線 1a v0.2 成品；其餘沿用線框架 v0.1 取值）。出處欄列出讀線的儲存格（錨點年；奇數年幾何平均、2033 起持平）。"
    hd = ["變數名", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處", "標記", "影響本模型哪些儲存格", "與線框架 v0.1 差異"]
    for j, h in enumerate(hd, 1):
        c = ws.cell(3, j, h); c.font = BOLD; c.fill = HDR
    E = lambda k: f"=In!E{rowof[k]}"

    def line_txt(key, fmt):
        yrs = LINE_ANCH if key == "wmy" else (2026, 2028, 2030, 2032, 2035)
        parts = []
        for s_ in SCN:
            parts.append('"' + s_ + '：' + '／'.join(str(y) for y in yrs) + ' "&' + '&"／"&'.join(f'TEXT(In!E{rowof[f"{key}_{y}_{s_}"]},"{fmt}")' for y in yrs))
        return "=" + '&"；"&'.join(parts)

    def line_src(kind):
        fn = LINE["file"]
        if kind == "price":
            return f"{fn}｜成本輸入!E:P 第 {PR_MY} 列（TY_PRICE）與第 {PR_CC} 列（TC_PRICE）；欄＝錨點年×（低價／基準／高價）：2026 E–G、2028 H–J、2030 K–M、2032 N–P"
        if kind == "opx":
            return f"{fn}｜每英里成本堆疊!C:N 第 {ST_MY + 2}–{ST_MY + 7} 列（Model Y 第 3–8 項）與第 {ST_CC + 2}–{ST_CC + 7} 列（Cybercab）；欄＝2026 C–E、2028 F–H、2030 I–K、2032 L–N（低成本／基準／高成本）"
        if kind == "pmd":
            return f"{fn}｜每英里成本堆疊!C:N 第 {ST_MY + 15} 列（Model Y）、第 {ST_CC + 15} 列（Cybercab）「備忘：每車每日付費英里」"
        return f"{fn}｜每英里成本堆疊!C{ST_FLEET}:I{ST_FLEET + 2}（Tesla 合計車隊加權單位營運成本，2026–2032）；Model Y 第 {ST_MY}–{ST_MY + 7} 列、Cybercab 第 {ST_CC}–{ST_CC + 7} 列"
    rows = [
        ("美國叫車基準里程 2026", "1a", E("rh26"), "十億英里", "Lyft 年化趟數 ÷ 份額 × 每趟英里", "Lyft 8-K；Gridwise", "Derived", "市場天花板與情境 rh 列 → 全部", "線框架未給總趟數（新增）"),
        ("人類叫車每英里乘客價（含小費）", "1a", E("p_h"), "US$/英里", "UberX 乘客實付", S["gridwise"], "Derived", "誘發倍數、滲透頁錨 1", "框架給司機每英里 2.65（司機端），本模型換成乘客端"),
        ("Uber 平台抽成", "1a", E("rh_take"), "比例", "中位數", S["gridwise"], "Verified", "trip_mi、p_h", "同框架"),
        ("司機每英里毛收入", "1a", E("drv_pm"), "US$/英里", "UberX", S["gridwise"], "Verified", "trip_mi", "同框架"),
        ("Tesla 營運成本目標", "1a", E("tsla_opx_tgt"), "US$/英里", "Tesla 法說，是否含折舊未明", S["q3call"], "Interested-party", "不進模型（v0.2 營運成本讀線）", "v0.2：不再作情境 A 2035 營運成本；線 Tesla 車隊加權營運成本 2032 基準約 0.71，低成本約 0.30"),
        ("Tesla 車隊加權每英里價格（A 低價／B 基準／C 高價）", "1a", line_txt("lp", "0.00"), "US$/英里", "車隊加權（Model Y／Cybercab）；2026／2028／2030／2032 錨點，奇數年幾何平均，2033–2035 持平於 2032", line_src("price"), "Derived（線內為 Assumed）", "市場天花板與情境 p 列 → 誘發倍數、閘門、營收", "取代 v0.1 的 p26→p35 直線（B：1.80→1.25）"),
        ("Tesla 車隊加權營運成本（第 3–8 項，不含車輛攤提與融資）", "1a", line_txt("lx", "0.00"), "US$/英里", "能源、維修清潔、保險理賠、遠端、安全員、場站；分部另算折舊與折現", line_src("opx"), "Derived（線內多為 Assumed）", "每英里經濟 opx 列 → EBIT", "取代 v0.1 的 opx27→opx35 直線（B：0.70→0.40）；線 B 2027 約 5.7，2032 約 0.71"),
        ("Tesla 車隊加權每車每日付費英里（稼動率）", "1a", line_txt("lm", "0"), "英里/日", "每車年付費英里＝本值 × 365 × 稼動率乘數 util_k（基準 1.0）", line_src("pmd"), "Derived（線內為 Assumed）", "每英里經濟 mpv 列 → 車隊、折舊、資本支出", "取代 v0.1 固定 5 萬英里/年（每日 137）；線 B 2026 約 43 → 2032 約 165"),
        ("車隊權重 w_MY（Model Y 占 Tesla 車隊英里）", "1a", line_txt("wmy", "0.000"), "比例", "w＝(堆疊頁 Tesla 合計車隊加權單位營運成本 − Cybercab 第 1–8 項合計) ÷ (Model Y 第 1–8 項合計 − Cybercab 合計)；每錨點年、每成本情境", line_src("w"), "Derived", "只用於推得上列三項的車隊加權（In 頁 wmy 列只作對照）", "線堆疊頁無直接輸出加權價格與稼動率，由本權重反解"),
        ("自有車轉換閘門的便利溢價 cf", "1a／1b", E("cf"), "倍", "價格 ≤ p_own×cf 全開、≥ 1.5×p_own×cf 關閉、中間線性", "對帳 1 判斷題預設乙（判斷題未獲回覆，採建議預設，可撤回）", "Assumed", "市場天花板與情境 gate 列 → 自有車轉換里程", "線 v0.2 只標年份、不計入營收；v0.1 為 1.5（價格 ≤ 1.155 全開）；乙＝0.77 全開、1.155 關閉"),
        ("單車取得成本（Cybercab 混合）", "1a", f"=TEXT(In!D{rowof['veh27']},\"#,##0\")&\"–\"&TEXT(In!F{rowof['veh27']},\"#,##0\")", "US$/台", "含自駕套件", "Musk、媒體", "Assumed", "折舊、資本支出", "框架：Cybercab <3 萬（Interested-party）；Waymo Ojai 約 12.5 萬"),
        ("車輛壽命", "1a", f"=TEXT(In!D{rowof['life']},\"0\")&\"–\"&TEXT(In!F{rowof['life']},\"0\")", "年", "經濟壽命", "—", "Assumed", "折舊、汰換", "框架用生涯英里；本模型 6 年 × 5 萬＝30 萬英里，落在框架下緣"),
        ("自有車每英里全成本（AAA）", "1a／1b", E("p_own"), "US$/英里", "15,000 英里/年平均", S["aaa"], "Verified", "轉換閘門、滲透頁錨 2", "框架未列（新增，專案規格指定）"),
        ("美國 VMT", "1b", E("vmt"), "十億英里", "FHWA 12 個月移動", S["fred"], "Verified", "自有車轉換里程", "框架未列（新增）"),
        ("叫車需求價格彈性", "1a", f"=TEXT(In!D{rowof['eps']},\"0.0\")&\"–\"&TEXT(In!F{rowof['eps']},\"0.0\")", "彈性", "誘發", "—", "Assumed", "誘發倍數", "框架未列（新增）"),
        ("FSD 付費用戶", "1b", 1.74, "百萬", "含買斷", S["q3res"], "Verified", "未進本模型（車主車輛占比的上限參考）", "同框架 174 萬"),
        ("Waymo 每週付費趟數", "1a", E("waymo_wk"), "萬趟", "2026-02", S["fortune"], "Interested-party", "Tesla 份額 sig26 區間、估值對照", "框架 2026-03 50 萬（insideevs）；本模型用 2026-02 40 萬"),
    ]
    for i, row in enumerate(rows):
        for j, v in enumerate(row, 1):
            ws.cell(4 + i, j, v).alignment = WRAP
    for col, wd in zip("ABCDEFGHI", (30, 8, 16, 12, 30, 40, 14, 30, 44)):
        ws.column_dimensions[col].width = wd


SENS_CASES = [
    # (驅動, 設定, 鍵, {鍵: (低值, 高值)} 或 特殊)
    ("價格彈性", "eps 0.3／1.2", "eps"),
    ("美國叫車成長", "rh_g 3%／9%", "rh_g"),
    ("Lyft 份額", "lyft_sh 23%／30%", "lyft_sh"),
    ("Lyft 北美占比", "lyft_na 85%／95%", "lyft_na"),
    ("行程長度", "trip_mi 6／8 英里", "trip_mi"),
    ("小費", "tip 5%／15%", "tip"),
    ("小客車 VMT 占比", "ldv_sh", "ldv_sh"),
    ("自有車成本", "p_own 0.56（小轎車）", "p_own"),
    ("便利溢價", "cf 1.0（基準）／1.5（v0.1 設定）", "cf"),
    ("S 曲線斜率", "lk 0.6／1.2", "lk"),
    ("S 曲線中點", "lmid 2030／2032（高＝提早）", "lmid"),
    ("2026 車隊", "fleet26", "fleet26"),
    ("2026 實收價格", "p26 1.40／2.10", "p26"),
    ("2026 Tesla 份額", "sig26", "sig26"),
    ("車主抽成", "take 25%／50%", "take"),
    ("2027 單車成本", "veh27", "veh27"),
    ("2035 單車成本", "veh35", "veh35"),
    ("稼動率", "util_k ×0.8／×1.2（線每車每日付費英里 × 365）", "util_k"),
    ("車輛壽命", "life 5／8", "life"),
    ("管銷研發比例", "sga 10%／25%", "sga"),
    ("稅率", "tax 15%／25%", "tax"),
    ("營運資金", "wc −5%／5%", "wc"),
    ("β", "beta 1.6／2.4", "beta"),
    ("ERP", "erp 4%／6%", "erp"),
    ("永續成長", "g_cap 2%／4%", "g_cap"),
    ("RONIC", "ronic_sp＝WACC＋5pt（只有高）", "ronic_sp"),
]


def in_lohi(k):
    for row in IN + IN2:
        if row[0] == k:
            return row[3], row[5]
    raise KeyError(k)


def sens_rows(P):
    base_w, base_r = value(P)
    out = []
    for lab, setting, k in SENS_CASES:
        lo, hi = in_lohi(k)
        res = []
        for v in (lo, hi):
            if v is None:
                res.append(None); continue
            Q = dict(P); Q[k] = v
            if k in ("rh_fare", "rh_take", "drv_pm"):
                Q["trip_mi"] = Q["rh_fare"] * (1 - Q["rh_take"]) / Q["drv_pm"]
            res.append(value(Q)[0])
        out.append((lab, setting, k, res[0], res[1]))
    # WACC ±1pt（以 erp 等效？改以直接加減 WACC：用 rf 不行（g 會變）→ 用 beta 換算 ±0.01/erp）
    for d, nm in ((-0.01, "WACC −1pt"), (0.01, "WACC +1pt")):
        Q = dict(P); Q["beta"] = P["beta"] + d / P["erp"]
        out.append((nm, "以 β 等效調整（g、RONIC 隨 WACC）", "wacc", value(Q)[0] if d < 0 else None, value(Q)[0] if d > 0 else None))
    # 情境權重
    for nm, wts in (("情境權重 33／34／33", (1/3, 0.34, 1 - 1/3 - 0.34)), ("只用情境 B", (0, 1, 0))):
        Q = dict(P); Q["w_A"], Q["w_B"], Q["w_C"] = wts
        out.append((nm, "w_A／w_B／w_C", "w_A／w_B／w_C", value(Q)[0], None))
    # 情境驅動整組（A／C 各自單獨套用到加權：顯示敏感度）
    for k in ("s35", "c35", "sig35", "o35", "intl35"):
        lo = dict(P); hi = dict(P)
        lo[f"{k}_B"] = P[f"{k}_C"]; hi[f"{k}_B"] = P[f"{k}_A"]
        out.append((f"情境 B 的 {k} 改用 C／A 值", f"{k}_B → {k}_C／{k}_A", f"{k}_B", value(lo)[0], value(hi)[0]))
    # 線路徑（價格 lp、營運成本 lx、稼動率 lm）：情境 B 整條路徑改用 C／A 路徑
    for k, nm in (("lp", "價格"), ("lx", "營運成本"), ("lm", "稼動率")):
        lo = dict(P); hi = dict(P)
        for y in range(2026, 2036):
            lo[f"{k}_{y}_B"] = P[f"{k}_{y}_C"]; hi[f"{k}_{y}_B"] = P[f"{k}_{y}_A"]
        out.append((f"情境 B 的線{nm}路徑改用 C／A 路徑", f"{k}_*_B → {k}_*_C／{k}_*_A（整條 2026–2035）", f"{k}_*_B", value(lo)[0], value(hi)[0]))
    # 2032 後外推：價格、營運成本、稼動率延續 2030–2032 趨勢（對帳 1 判斷題 2 的敏感度版）
    Q = dict(P)
    for k in ("lp", "lx", "lm"):
        for s_ in SCN:
            r_ = (P[f"{k}_2032_{s_}"] / P[f"{k}_2030_{s_}"]) ** 0.5
            for y in (2033, 2034, 2035):
                Q[f"{k}_{y}_{s_}"] = P[f"{k}_2032_{s_}"] * r_ ** (y - 2032)
    out.append(("2032 後外推：價格、成本、稼動率延續 2030–32 趨勢", "lp／lx／lm 2033–2035 ＝ 2032 × (2032÷2030)^((年−2032)/2)", "lp_*／lx_*／lm_* 2033–35", value(Q)[0], None))
    # 終值替代法
    return base_w, base_r, out


def write_sens(ws, P, base_w, base_r, rows, build_id, fp):
    ws["A1"] = "Sens：一次動一個參數，整本重算（Python 引擎與 xlsx 公式同式；核對程式比對兩者基準值）。本頁「目標價」＝Tesla Robotaxi 加權分部價值（十億美元，評價日現值），非每股目標價。"
    ws["A2"] = ("說明：ronic_sp 只有高案（RONIC＝WACC＋5pt）。val_date、rf、lyft_rq 等 Verified 單點值無區間不列：屬公開數據，無低／高。"
                "情境專屬列（s35_A 等）以「情境 B 改用 C／A 值」案涵蓋；trip_mi 的低／高直接取代推算值。")
    ws.cell(3, 1, "build_id"); ws.cell(3, 4, build_id)
    ws.cell(4, 1, "模型狀態"); ws.cell(4, 2, "基準加權分部價值（建置時）"); ws.cell(4, 3, base_w)
    ws.cell(4, 4, f'=IF(ABS(估值!F{{EVROW}}-C4)<0.0005,"CURRENT","STALE")')
    ws.cell(5, 1, "輸入指紋"); ws.cell(5, 4, fp)
    hdr = ["驅動", "設定", "鍵", "低案目標價", "高案目標價", "低案變化", "高案變化", "DCF單法變化（低）", "DCF單法變化（高）"]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(7, j, h); c.font = BOLD; c.fill = HDR
    ws.cell(8, 1, "基準（情境加權）"); ws.cell(8, 3, "—"); ws.cell(8, 4, base_w); ws.cell(8, 5, base_w)
    r = 9
    for lab, setting, k, lo, hi in rows:
        ws.cell(r, 1, lab); ws.cell(r, 2, setting); ws.cell(r, 3, k)
        if lo is not None:
            ws.cell(r, 4, lo); ws.cell(r, 6, f"=D{r}/$D$8-1"); ws.cell(r, 8, f"=D{r}-$D$8")
        if hi is not None:
            ws.cell(r, 5, hi); ws.cell(r, 7, f"=E{r}/$D$8-1"); ws.cell(r, 9, f"=E{r}-$D$8")
        r += 1
    ws.cell(r + 1, 1, "讀法：本分部 100% 以 DCF 估值（無同口徑獲利同業），DCF 單法變化＝加權分部價值變化（十億美元）。")
    for col, wd in zip("ABCDEFGHI", (30, 40, 18, 14, 14, 11, 11, 16, 16)):
        ws.column_dimensions[col].width = wd


def write_readme(ws, build_id, VS):
    lines = [
        f"Tesla Robotaxi 分部模型 {VER}｜build_id {build_id}｜資料截止 2026-11-16｜成品日 2026-11-16｜股價基準 2026-11-10（逾期，只作對照）",
        "命題：Tesla Robotaxi 的價值取決於「每英里實收價格能否壓到自有車成本附近」——價格一旦低於自有車每英里全成本 × 便利溢價，私人里程開始轉入（傑文斯效應），市場天花板從叫車（百億英里量級，見 In 頁 rh26）跳到小客車 VMT（兆英里量級，見 In 頁 vmt）；否則只是在叫車市場與 Waymo 分份額。",
        "主頁＝滲透（最上方價格區塊）；市場天花板與情境 → 每英里經濟 → 估值（分部 DCF）→ 給Tesla模型（RT_VAL_LO／BASE／HI、RT_VAL_A／B／C、RT_REV_<情境>_<年>）。v0.2：價格、營運成本（第 3–8 項）、稼動率讀 Robotaxi 線 v0.2（Tesla 車隊加權；In 頁第 8 節），自有車閘門 cf＝1.0；判斷題預設見 README.md。",
        "分部價值（十億美元，評價日現值）：低（情境 C）／基準（加權）／高（情境 A）見下列公式；評等：未上市。",
    ]
    for i, t in enumerate(lines, 1):
        ws.cell(i, 1, t)
    ws.cell(6, 1, "低（情境 C）"); ws.cell(6, 2, f"=估值!E{VS['ev']}")
    ws.cell(7, 1, "基準（加權）"); ws.cell(7, 2, f"=估值!F{VS['ev']}")
    ws.cell(8, 1, "高（情境 A）"); ws.cell(8, 2, f"=估值!C{VS['ev']}")
    ws.cell(9, 1, "基準 ÷ TSLA 市值"); ws.cell(9, 2, f"=估值!F{VS['pct']}")
    ws.cell(10, 1, "摘要句"); ws.cell(10, 2, f'="Robotaxi 分部價值 "&TEXT(B6,"0")&"／"&TEXT(B7,"0")&"／"&TEXT(B8,"0")&" 十億美元（低／基準／高），基準約占 TSLA 市值 "&TEXT(B9,"0.0%")&"。"')
    ws.cell(12, 1, "失效條件（可觀察）：(1) Austin 每英里牌價在 2027 年仍高於 2 美元且無下降 → 情境 A 失效；(2) 2027 Q2 車主車輛未獲准加入 → o35 下修；(3) 2027 年底 Tesla 無人城市 < 10 或每日里程 < 50 萬英里 → 情境 B 份額路徑偏高；(4) 重大事故導致停駛 → 尾端風險（未建模，見 README.md 風險登錄）。")
    ws.cell(13, 1, "不讀 Siliconomics／Tokenomics／Wattonomics（算力線協調檔不適用實體 AI）；訓練算力支出未扣（算力收支只讀、未取得）。負值不歸零（分部虧損是補貼訂價的一部分）。")
    ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 100


def write_log(ws):
    ws["A1"] = "改版紀錄（v0.1 → v0.2；2026-11-16；依對帳 1《Tesla Robotaxi 分部 vs Robotaxi 線》，判斷題未獲回覆，採建議預設，可撤回）"
    for j, h in enumerate(["#", "項目", "v0.1", "v0.2", "影響／出處"], 1):
        c = ws.cell(3, j, h); c.font = BOLD; c.fill = HDR
    rows = [
        (1, "價格路徑", "p26→p35 直線（B：1.80→1.25）", "讀線 Tesla 車隊加權價格（A 低價／B 基準／C 高價）；2026／28／30／32 錨點、奇數年幾何平均、2033–2035 持平於 2032", "In 第 8 節 lp_<年>_<情境>；B 2030 為 1.11、2032 為 0.90"),
        (2, "營運成本（不含車輛折舊）", "opx27→opx35 直線（B：0.70→0.40）", "讀線堆疊第 3–8 項車隊加權（A 低成本／B 基準／C 高成本），同樣插補與持平", "In lx_<年>_<情境>；B 2027 為 5.69、2030 為 1.07、2032 為 0.71"),
        (3, "稼動率", "固定 mpv 5 萬英里/年", "線車隊加權每車每日付費英里 × 365 × util_k（基準 1.0；Sens ×0.8／×1.2）", "In lm_<年>_<情境>、util_k；B 2026 為 43、2032 為 165 英里/日"),
        (4, "車隊權重", "—", "w_MY＝(堆疊頁 Tesla 合計車隊加權單位營運成本 − Cybercab 第 1–8 項合計) ÷ (Model Y 合計 − Cybercab 合計)，每錨點年、每成本情境", "In wmy_<年>_<情境>（只作對照）；對帳頁列出讀線儲存格"),
        (5, "自有車閘門便利溢價 cf", "1.5（價格 ≤ 1.155 全開、≥ 1.73 關閉）", "1.0（價格 ≤ 0.77 全開、≥ 1.155 關閉、中間線性；區間 1.0–1.5，Assumed）", "對帳 1 判斷題 1 預設乙"),
        (6, "負值", "—", "不歸零（對帳 1 判斷題 3 預設）", "分部虧損是補貼訂價的一部分"),
        (7, "給Tesla模型輸出", "RT_VAL_LO／BASE／HI", "另加 RT_VAL_A／B／C 與 RT_REV_<A/B/C>_<2027–2035>（百萬美元）", "Tesla 模型 SOTP 與 DCF 腿讀取，避免兩條 Robotaxi 路徑並存"),
        (8, "結果（十億美元）", "A 260.9／B 48.3／C 2.2／加權 89.9", "見 README 頁與 README.md 結果表", "與對帳 1 變體乙（2032 後持平）一致：加權 29.8、A 133.2、B −1.0、C −12.0"),
    ]
    for i, row in enumerate(rows):
        for j, v in enumerate(row, 1):
            ws.cell(4 + i, j, v).alignment = WRAP
    for col, wd in zip("ABCDE", (5, 22, 36, 70, 56)):
        ws.column_dimensions[col].width = wd


def main():
    wb, P, R, VR, VS, rowof, sheets, ycol = build()
    wsR, wsP, wsSrc, wsRec, wsS, wsI, wsM, wsU, wsV, wsLog = sheets
    write_pen(wsP, R, ycol)
    write_src(wsSrc)
    write_rec(wsRec, rowof)
    base_w, base_r, rows = sens_rows(P)
    fp_src = json.dumps({k: P[k] for k in sorted(P)}, sort_keys=True, default=str)
    fp = hashlib.sha256(fp_src.encode()).hexdigest()[:12]
    src_hash = hashlib.sha256(open(__file__, "rb").read()).hexdigest()[:8]
    build_id = f"{VER}-{fp}-{src_hash}"
    write_sens(wsS, P, base_w, base_r, rows, build_id, fp)
    wsS["D4"] = wsS["D4"].value.replace("{EVROW}", str(VS["ev"]))
    write_readme(wsR, build_id, VS)
    write_log(wsLog)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    eng = {s: run(P, s) for s in SCN}
    json.dump({"build_id": build_id, "weighted": base_w, "scn": {s: eng[s]["ev"] for s in SCN},
               "rev": {s: {y: eng[s]["out"]["rev"][y] for y in DCF_YEARS} for s in SCN},
               "rt": {s: {y: eng[s]["out"]["rt"][y] for y in DCF_YEARS} for s in SCN},
               "fleet": {s: {y: eng[s]["out"]["fleet"][y] for y in DCF_YEARS} for s in SCN},
               "tvsh": {s: eng[s]["pv_tv"] / eng[s]["ev"] for s in SCN},
               "ev_row": VS["ev"], "sens": rows},
              open(os.path.join(HERE, "engine_snapshot.json"), "w"), ensure_ascii=False, indent=1, default=str)
    print(OUT, build_id, "weighted", round(base_w, 2), {s: round(eng[s]["ev"], 2) for s in SCN})


if __name__ == "__main__":
    main()

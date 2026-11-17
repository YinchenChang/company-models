# HubSpot 個股模型：輸入表（In 頁）、歷史資料（Hist 頁）與 Tokenomics 快照讀取。由 build_hubspot.py 匯入。
# 金額 US$M；客戶數為戶；ASRPC 為 US$。所有數字的逐列出處見 data/facts_hubspot.csv（facts_hubspot.py 產生）。
import datetime as dt
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
D = dt.datetime
SEC = "#"
TK = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v22.0.json"), encoding="utf-8"))
TKS = "data/tokenomics_snapshot_v22.0.json（Tokenomics v22.0 快照，commit 79f6b30；工作單 G2.1：無 token 揭露時以快照為 Analogy）"


def _tk(name):
    """取 Tokenomics 快照值（陣列取任務 1＝0.1 小時任務；價格為純量或三級）"""
    d = TK.get("items", TK)
    v = d[name]
    return v["values"] if isinstance(v, dict) else v


def tkv(name, idx=None):
    v = _tk(name)
    if isinstance(v, list):
        return float(v[idx if idx is not None else 0])
    return float(v)


# 公司來源簡稱
K25 = "https://www.sec.gov/Archives/edgar/data/1404655/000119312526046646/hubs-20251231.htm（10-K FY2025）"
Q3Q = "https://www.sec.gov/Archives/edgar/data/1404655/000119312526447363/hubs-20260930.htm（10-Q 2026Q3）"
Q3PR = "https://www.sec.gov/Archives/edgar/data/1404655/000119312526447249/hubs-ex99_1.htm（Q3 2026 新聞稿，2026-11-04）"
PPC = "https://ppc.land/hubspot-forecasts-q4-revenue-below-q3-as-growth-slows-to-12-from-17/（法說摘要，第二來源）"
HSP = "https://www.hubspot.com/company-news/hubspots-customer-agent-and-prospecting-agent-now-you-pay-when-the-task-is-complete（HubSpot 官網）"
CR = "https://chartrow.com/quote/{}（2026-11-13；只一個資料商）"
VEND = "資料商整理值，非法定文件；口徑未核"
CO = "公司自訂指標或管理層說法，傾向呈現成長與獲利改善"
SELL = "賣方預估；賣方有評等與投行業務關係"
NOTM = "只作對照，不進模型"

# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 說明, 來源誘因)
INP = [
    (SEC, "A. 日期、現價、股數與資產負債（評價日 2026-09-30；US$M；股數百萬股）"),
    ("data_date", "資料截止日（本版最晚一筆輸入）", "日期", None, D(2026, 11, 16), None, "Derived", "TickFlow 賣方目標價頁標示日 2026-11-16", "README 首列與 HTML 表頭分列資料截止日、成品日、股價基準日", None),
    ("px_date", "股價基準日（現價日）", "日期", None, D(2026, 11, 13), None, "Verified", "QuiverQuant 歷史收盤價：2026-11-13（ChartRow 同值）", "2026-11-13 為最近完整交易日（11-16 為週一、收盤未取）", None),
    ("val_date", "評價日（最近一季資產負債表日）", "日期", None, D(2026, 9, 30), None, "Verified", Q3Q + " 資產負債表日", "Q4 2026 財報預計 2027-02", None),
    ("tp_date", "目標時點＝EDATE(股價基準日, 12)", "日期", None, "=EDATE(px_date,12)", None, "Derived", "原則第 10 節第 26 條", "持有期間＝現價日起 12 個月；折現期間＝評價日起", None),
    ("px", "現價（2026-11-13 收盤）", "US$", None, 217.25, None, "Verified", "QuiverQuant 與 ChartRow 兩處同值 217.25", "2026-11-04 法說日 216.64、11-05 204.96、11-12 215.84；52 週區間 170.28–405.83", None),
    ("sh_basic", "流通股數（10-Q 封面，2026-10-30）", "百萬股", None, 48.836461, None, "Verified", Q3Q + " 封面", "法定申報封面為主來源（工作單 G2.4）", None),
    ("rsu_dil", "RSU 與 PSU 潛在股數（Q3 反稀釋排除之 3 個月加權；2026-09-30 期末餘額未找到）", "百萬股", 1.471, 2.525, None, "Derived", Q3Q + " 每股盈餘附註；低＝2025 年底未歸屬 1.471（10-K）", "期末未歸屬餘額為未找到，以此代理；全數視為稀釋（庫藏股法所需之未攤銷費用未取）", None),
    ("forfeit", "RSU 預期沒收率", "%", 0.0, 0.0, 0.15, "Assumed", "10-K 附註未取得沒收率；工作單 G2.4 要求扣除預期沒收", "基準 0 為保守（未找到）；高案 15% 為敏感度", None),
    ("opt_dil", "選擇權庫藏股法淨增股數", "百萬股", None, 0.0, 0.093, "Assumed", Q3Q + "：潛在稀釋選擇權 93 千股；履約價未載（未找到），基準 0，高案＝全數視為價內（上界）", "履約價未找到，無法做庫藏股法", None),
    ("cash_0", "現金、短期與長期投資（2026-09-30）", "US$M", None, "=h_cashinv_q326", None, "Derived", Q3PR + " 資產負債表：現金 1,121.342＋短期投資 175.414＋長期投資 13.747", "長期投資 13.7 計入（與現金同流動性口徑差異小）", None),
    ("debt_0", "債務本金與融資租賃", "US$M", None, 0.0, None, "Verified", Q3Q + " Revolving Credit Facility 附註：5 億美元額度未動用", "營業租賃不計入淨債務（工作單 G2.5 採用口徑）；融資租賃未找到明細，視為 0", None),
    ("sh_eps0", "Q4 2026 稀釋股數（EPS 用）", "百萬股", None, 48.9, None, "Interested-party", Q3PR + " Business Outlook：Q4 非 GAAP EPS 以 48.9 百萬稀釋股計", "EPS 用庫藏股法口徑；DCF 每股價值另以含全數 RSU 的評價股數（見 Val）", CO),
    ("bb_q4", "Q4 2026 回購金額", "US$M", 0, 200, 400, "Assumed", "Q3 回購 230.8；2026 年授權 10 億美元已用約 936（剩餘 964.1 為 Q3 後；法說）", "Q1–Q3 回購 206.7／536.2／230.8", None),
    (SEC, "B. 折現率、估值與目標價設定"),
    ("rf", "無風險利率（美國 10 年期公債，2026-11-12）", "%", None, 0.0510, None, "Verified", "FRED DGS10（2026-11-12 收盤 5.10%；現價日前 1 個交易日；工作單 G5.2 統一值）", "v0.1 曾用 Trading Economics 2026-11-13 的 5.15%；v0.2 依 G5.2 統一改 5.10%", None),
    ("rf_date", "rf 資料日", "日期", None, D(2026, 11, 12), None, "Verified", "FRED DGS10", "供逾期核對（Recon）", None),
    ("beta_sa", "原始 β：StockAnalysis（5 年）", "x", None, 1.24, None, "Analogy", "https://stockanalysis.com/stocks/hubs/statistics/（2026-10-26）", "未載報酬頻率", None),
    ("beta_mb", "原始 β：MarketBeat", "x", None, 1.19, None, "Analogy", "https://www.marketbeat.com/instant-alerts/hubspot-nysehubs-releases-quarterly-earnings-results-beats-estimates-by-024-eps-2026-08-05/（Q2 後，無日期）", "無資料日", None),
    ("blume_a", "Blume 調整權重", "x", None, 0.67, None, "Assumed", "Blume (1971) 慣例 β_adj＝0.67×β＋0.33", None, None),
    ("beta", "β（Blume 調整後兩來源平均）", "x", None, "=AVERAGE(blume_a*beta_sa+1-blume_a,blume_a*beta_mb+1-blume_a)", None, "Derived", "兩來源 Blume 調整後平均", None, None),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "個股 DCF 標準規格 v1.0 第 2 節預設 5%（4–6%）", None, None),
    ("tax", "稅率（NOPAT 與損益；公司 non-GAAP 稅率 15%）", "%", 0.12, 0.15, 0.21, "Assumed", Q3PR + "：non-GAAP 稅率 15%；GAAP 有效稅率 FY2025 33.9%、9M26 11.3%（扭曲，不採）", "未建現金稅橋（已付稅 FY2025 19.3 對 non-GAAP 稅前）；列未做", None),
    ("g_tv", "永續成長率", "%", 0.02, 0.03, 0.04, "Assumed", "個股 DCF 標準規格 v1.0 第 3 節：預設 3%（2–4%）", None, None),
    ("ronic_sp", "永續期 RONIC 相對 WACC 的差", "%", 0, 0, 0.05, "Assumed", "個股 DCF 標準規格：預設 RONIC＝WACC", None, None),
    ("mid", "期中慣例：年度現金流距期初", "年", None, 0.5, None, "Derived", "個股 DCF 標準規格", None, None),
    ("tp_wA", "情境權重 A", "%", None, 0.25, None, "Assumed", "個股 DCF 標準規格：預設 25／50／25", None, None),
    ("tp_wB", "情境權重 B", "%", None, 0.5, None, "Assumed", "同上", None, None),
    ("tp_wC", "情境權重 C", "%", None, 0.25, None, "Assumed", "同上", None, None),
    ("w_dcf", "目標價中 DCF 的權重（其餘＝Comparable）", "%", 0.3, 0.5, 0.7, "Assumed", "原則第 4 節預設 DCF 50%＋Comparable 50%", None, None),
    ("t_tgt", "折現期間：評價日 → 目標時點", "年", None, "=(tp_date-val_date)/365", None, "Derived", "(tp_date − val_date) ÷ 365", None, None),
    ("t_hold", "持有期間：股價基準日 → 目標時點", "年", None, "=(tp_date-px_date)/365", None, "Derived", "(tp_date − px_date) ÷ 365", None, None),
    ("wacc_adj", "WACC 調整（只用於敏感度；基準 0）", "%", -0.01, 0.0, 0.01, "Assumed", "個股 DCF 標準規格第 4 節：WACC ±1pt", None, None),
    ("tv_alt", "終值方法開關（0＝NOPAT 價值驅動式；1＝FCFF×(1+g) 替代法）", "0／1", 0, 0, 1, "Assumed", "個股 DCF 標準規格第 4 節", "只用於敏感度", None),
    ("sbc_ab", "DCF 是否加回股票薪酬（0＝不加回；1＝加回）", "0／1", 0, 0, 1, "Assumed", "個股 DCF 標準規格第 1 節：預設不加回", "只用於敏感度", None),
    ("pe_crm", "同業前瞻 P/E：Salesforce（美國 GAAP；分母＝前瞻一年每股盈餘）", "x", None, 13.6, None, "Analogy", CR.format("crm"), "分母口徑（GAAP 或非 GAAP）與前瞻期間頁面未載；租賃處理未核", None),
    ("pe_now", "同業前瞻 P/E：ServiceNow（美國 GAAP）", "x", None, 21.3, None, "Analogy", CR.format("now"), "同上", None),
    ("pe_wday", "同業前瞻 P/E：Workday（美國 GAAP）", "x", None, 14.4, None, "Analogy", CR.format("wday"), "同上", None),
    ("pe_adbe", "同業前瞻 P/E：Adobe（美國 GAAP）", "x", None, 8.9, None, "Analogy", CR.format("adbe"), "同上", None),
    ("pe_intu", "同業前瞻 P/E：Intuit（美國 GAAP）", "x", None, 12.6, None, "Analogy", CR.format("intu"), "同上", None),
    ("pe_team", "同業前瞻 P/E：Atlassian（美國 GAAP）", "x", None, 22.4, None, "Analogy", CR.format("team"), "同上", None),
    ("pe_mndy", "同業前瞻 P/E：monday.com（美國 GAAP；以色列公司）", "x", None, 10.2, None, "Analogy", CR.format("mndy"), "同上", None),
    ("pe_docu", "同業前瞻 P/E：DocuSign（美國 GAAP）", "x", None, 15.4, None, "Analogy", CR.format("docu"), "同上", None),
    ("pe_twlo", "同業前瞻 P/E：Twilio（美國 GAAP）", "x", None, 30.8, None, "Analogy", CR.format("twlo"), "同上", None),
    ("pe_zm", "同業前瞻 P/E：Zoom（美國 GAAP）", "x", None, 15.4, None, "Analogy", CR.format("zm"), "同上", None),
    ("pe_ddog", "同業前瞻 P/E：Datadog（美國 GAAP）", "x", None, 57.9, None, "Analogy", CR.format("ddog"), "同上；成長 35% 遠高於 HubSpot，倍數不可比（中位數不受影響）", None),
    ("pe_date", "同業倍數資料日", "日期", None, D(2026, 11, 13), None, "Verified", "同上", "供逾期核對", None),
    ("pe_med", "同業前瞻 P/E 中位數", "x", None, "=MEDIAN(pe_crm,pe_now,pe_wday,pe_adbe,pe_intu,pe_team,pe_mndy,pe_docu,pe_twlo,pe_zm,pe_ddog)", None, "Derived", "11 家中位數；少於約 8 家才不做迴歸，此處 11 家但分母口徑未核，不做迴歸", "股權倍數（有槓桿）配本公司非 GAAP EPS（有槓桿）；本益比法只作對照列（工作單 G6）", None),
    ("ps_peer", "同業中位 P/S（最近一季末；relmult_hist，11 家）", "x", None, 4.03, None, "Derived", "software_ai/data/relmult_hist/HUBS.csv 與 out/hubspot.json 的 ps_series 最後一點：peer_median（季末 2026-06-30；市值÷TTM 營收，StockAnalysis）", "Comparable 主腿（工作單 G6）：P/S 口徑與 PS／PSG 相對值一致；季末距現價日逾 2 個交易日，已於 Recon 標逾期", None),
    ("ps_peer_date", "同業 P/S 資料日（季末）", "日期", None, D(2026, 6, 30), None, "Verified", "同上", "供逾期核對（Recon）", None),
    ("rel_ps_now", "HUBS 現行 P/S 相對值（最近一季末 ÷ 同業中位）", "x", None, 0.672456575682382, None, "Derived", "software_ai/data/relmult_hist/out/hubspot.json：ps_series 最後一點 rel（HUBS 2.71 ÷ 同業 4.03）", "只供對照（Val 與 Recon 的現行相對值）；不進主模型", None),
    ("rel_mult", "相對倍數係數（HUBS 對同業中位數；取相對倍數歷史讀數的幾何平均基準）", "x", 0.6465043833518652, 1.2714865566748963, 1.579667603932714, "Derived", "software_ai/data/relmult_hist/out/hubspot.json：geo_base（base_source：PS 與 PSG 兩讀數幾何平均，Derived）；低／高＝sens.low／sens.high；窗口 2021-09-30 至 2026-06-30，同業 11 家", "PS 相對值與 PSG 相對值（PSG＝PS÷TTM 營收年增）的歷史中位數；本模型 Comparable 主腿為 P/S 口徑（工作單 G6），與 PS／PSG 相對值同口徑；P/E 只作對照列；PS 與 PSG 兩讀數相差 2.0%（未超過 30%）；HUBS 成長／同業成長比 1.52（TTM）；現行前瞻 P/E 相對值見 Recon", None),
    ("rm_ps", "relmult_hist：ps_only 讀數（PS 相對值歷史中位數）", "x", None, 1.2591749963481562, None, "Derived", "software_ai/data/relmult_hist/out/hubspot.json：ps_only", "只供 Sens 案例；不進主模型", None),
    ("rm_psg", "relmult_hist：psg_only 讀數（PSG 相對值歷史中位數）", "x", None, 1.283918493055893, None, "Derived", "software_ai/data/relmult_hist/out/hubspot.json：psg_only", "只供 Sens 案例；不進主模型", None),
    (SEC, "C. 命題驅動：客戶數 × 每客戶訂閱營收；AI 代理按任務計價（增量）與席位被代理取代（侵蝕）；v0.2 改讀 unit_economics"),
    ("na_q4", "Q4 2026 客戶淨增", "戶", 2000, 4000, 5000, "Assumed", PPC + "：Q4 淨增「低於 5,000」（管理層說法）", "Q3 實績 5,258；Q2 6,988", None),
    ("q4_asrpc", "Q4 2026 ASRPC 季增（相對 Q3 12,038 美元）", "%", -0.02, -0.01, 0.0, "Assumed", "指引中點隱含約 −1.8%（只在 Recon 對照）；基準取持平與指引之間", "Q4 指引營收 945–947 低於 Q3 947.6", None),
    ("na_27", "客戶淨增率：FY2027（期末客戶年增）", "%", 0.04, 0.06, 0.08, "Assumed", "FY2026 約 +9.4%（Q3 後含 Q4 假設）；Q3 淨增 5.3k、Q4 低於 5k，換算年化約 6–7%", "線性收斂到 FY2032 值", None),
    ("na_32", "客戶淨增率：FY2032", "%", 0.02, 0.03, 0.04, "Assumed", "收斂假設，無直接證據", "承重 Assumed 之一", None),
    ("g_up", "每客戶訂閱營收的年增（不含席位被取代；價格、升級、交叉銷售）", "%", 0.01, 0.03, 0.05, "Assumed", "ASRPC 年變動：FY2022 +6.5%、FY2023 +2.0%、FY2024 −0.4%、FY2025 +0.6%（Derived，10-K）；Q4 2025 至 Q3 2026 季增 +3.0%（含匯率）", "承重 Assumed 之一；NRR 101%（Q3 法說）", None),
    ("seat_sh", "訂閱營收中按席位計價的占比", "%", 0.4, 0.6, 0.8, "Assumed", "HubSpot 為席位加用量混合；占比未揭露（未找到）", "承重 Assumed 之一；新試點方案改為 seats and credits（PPC）", None),
    ("repl_A", "情境 A：每年被 AI 代理取代的席位占比", "%", None, 0.008, None, "Assumed", "情境設定：取代慢", "席位被取代率 × 席位占比 ＝ 每客戶訂閱營收年侵蝕", None),
    ("repl_B", "情境 B：每年被 AI 代理取代的席位占比", "%", None, 0.025, None, "Assumed", "情境設定：NRR 由 103.5% 降至 101%（約 2pt）全數歸因 AI 之上界的一半", "公司稱 FY2026 NRR 較 FY2025 降約 2 個百分點（PPC）", None),
    ("repl_C", "情境 C：每年被 AI 代理取代的席位占比", "%", None, 0.05, None, "Assumed", "情境設定：取代快", None, None),
    ("tk_ref", "Customer Agent 啟用客戶數對照（8,000 以上；2026-04）", "戶", None, 8000, None, "Interested-party", HSP, "只作對照，不進模型", CO),
    ("ai_rev_26", "FY2026 代理按任務計價營收（Customer Agent＋Prospecting Agent）", "US$M", 1.5, 3.0, 6.0, "Assumed", "Q3 解決件數逾 200 萬件 × 0.50＝約 1.0 的季營收（Derived，PPC＋HSP）；前幾季較少", "公司未揭露 AI 營收（未找到）；Prospecting Agent 與其他 credits 未計", None),
    ("tasks_0", "每採用客戶每年任務數（FY2026）", "次", 600, 1000, 1500, "Assumed", "Q3 200 萬件 ÷ 8,000 客戶 ÷ 1 季 ＝ 250 件/季（Derived；客戶數為 2026-04 說法，日期不一致）", "承重 Assumed 之一", None),
    ("g_tasks", "每採用客戶任務數年增", "%", 0.10, 0.25, 0.40, "Assumed", "agent 動作 5 倍、credits 消耗約 2.5 倍（Q3 法說，基期未載）", None, None),
    ("price_0", "Customer Agent 單價（每次解決對話；2026-04-14 起）", "US$", None, 0.50, None, "Verified", HSP, "原為每次對話 1.00 美元；50 Credits／次，每 Credit 約 0.01 美元", None),
    ("adopt_A", "情境 A：FY2032 付費採用 agent 的客戶占比", "%", None, 0.25, None, "Assumed", "情境設定", "現況：Pro 以上客戶中 20% 使用 agents（Interested-party），付費任務占比小", None),
    ("adopt_B", "情境 B：FY2032 付費採用占比", "%", None, 0.15, None, "Assumed", "情境設定", None, None),
    ("adopt_C", "情境 C：FY2032 付費採用占比", "%", None, 0.08, None, "Assumed", "情境設定", None, None),
    ("dp_A", "情境 A：每任務單價年變動", "%", None, 0.0, None, "Assumed", "情境設定：價格持平", None, None),
    ("dp_B", "情境 B：每任務單價年變動", "%", None, -0.05, None, "Assumed", "情境設定：年降 5%（單價 2026-04 已由 1.00 降到 0.50）", None, None),
    ("dp_C", "情境 C：每任務單價年變動", "%", None, -0.10, None, "Assumed", "情境設定：年降 10%", None, None),
    ("tok_fresh", "Tokenomics IF_TaskTokFresh（任務 2＝0.1 小時級，向量第 2 欄）", "tok", None, tkv("IF_TaskTokFresh", 1), None, "Analogy", TKS, "公司未揭露每任務 token（未找到）；以 Tokenomics 工作負載為類比，不標 Verified", None),
    ("tok_cache", "Tokenomics IF_TaskTokCached（向量第 2 欄）", "tok", None, tkv("IF_TaskTokCached", 1), None, "Analogy", TKS, None, None),
    ("tok_dec", "Tokenomics IF_TaskTokDec（向量第 2 欄）", "tok", None, tkv("IF_TaskTokDec", 1), None, "Analogy", TKS, None, None),
    ("ctx_m", "客服對話相對 0.1 小時任務的 token 倍數", "x", 2, 4, 10, "Assumed", "多輪對話與知識庫檢索使每解決對話的 token 高於基礎任務；無揭露", "承重 Assumed 之一", None),
    ("succ", "任務成功率（IF_TaskSucc_Terra 任務 1）", "x", None, tkv("IF_TaskSucc_Terra", 1), None, "Analogy", TKS, "每解決任務成本＝每次嘗試成本 ÷ 成功率", None),
    ("tkp_fL", "Luna 新鮮輸入價（US$／百萬 token）", "US$", None, tkv("IF_PriceFresh_Luna"), None, "Derived", TKS, "市場價，不自行重估", None),
    ("tkp_cL", "Luna 快取輸入價", "US$", None, tkv("IF_PriceCached_Luna"), None, "Derived", TKS, None, None),
    ("tkp_oL", "Luna 輸出價", "US$", None, tkv("IF_PriceOut_Luna"), None, "Derived", TKS, None, None),
    ("tkp_fT", "Terra 新鮮輸入價", "US$", None, tkv("IF_PriceFresh_Terra"), None, "Derived", TKS, None, None),
    ("tkp_cT", "Terra 快取輸入價", "US$", None, tkv("IF_PriceCached_Terra"), None, "Derived", TKS, None, None),
    ("tkp_oT", "Terra 輸出價", "US$", None, tkv("IF_PriceOut_Terra"), None, "Derived", TKS, None, None),
    ("tkp_fS", "Sol 新鮮輸入價", "US$", None, tkv("IF_PriceFresh_Sol"), None, "Derived", TKS, None, None),
    ("tkp_cS", "Sol 快取輸入價", "US$", None, tkv("IF_PriceCached_Sol"), None, "Derived", TKS, None, None),
    ("tkp_oS", "Sol 輸出價", "US$", None, tkv("IF_PriceOut_Sol"), None, "Derived", TKS, None, None),
    ("mix_L", "模型組合：Luna 占比", "%", None, 0.2, None, "Analogy", "選哪一層價格屬類比（工作單 G2.1）；HubSpot 不揭露所用模型", None, None),
    ("mix_T", "模型組合：Terra 占比", "%", None, 0.6, None, "Analogy", "同上", None, None),
    ("mix_S", "模型組合：Sol 占比", "%", None, 0.2, None, "Analogy", "同上", None, None),
    ("route", "路由到便宜模型（0＝基準組合；1＝全部走 Luna）", "0／1", 0, 0, 1, "Assumed", "工作單 G2.2：模型組合案例", "只用於敏感度", None),
    ("tk_pscale", "token 價格倍數（基準 1；只動成本一邊）", "x", 0.5, 1.0, 2.0, "Assumed", "Tokenomics 價格不確定度；高案為上界（只動成本、售價不連動）", "聯動案例見 Sens「token 價格與任務售價聯動」", None),
    ("link_pass", "token 價格是否同步傳導到任務售價（0＝否；1＝是）", "0／1", None, 0, 1, "Assumed", "工作單 G2.2：價格與成本聯動案例", "只用於敏感度；1＝任務售價乘 tk_pscale", None),
    ("pd_tok", "token 價格年變動（待 AI供應鏈 登錄）", "%", -0.40, -0.25, 0.0, "Assumed", "Tokenomics 無年降幅登錄；待 AI供應鏈 登錄", "只作敏感度標記；基準與 Terra 近年降幅一致性未核", None),
    ("emb_pct", "內嵌 AI 推論成本占營收（FY2026；已含於近四季成本率）", "%", 0.005, 0.01, 0.02, "Assumed", "非 GAAP 訂閱毛利率 85.4% 對去年同期 86.7%（−1.3pt，含 AI 功能使用）；訂閱與託管成本年增 23.2（Q3 10-Q，部分因 AI 使用）", "公司未拆分（未找到）", None),
    ("emb_g", "內嵌 AI 使用量年增", "%", 0.3, 0.6, 1.5, "Assumed", "credits 消耗約 2.5 倍、agent 動作 5 倍（Q3 法說，基期未載）", "成本＝使用量 ×（1＋token 價格年變動）", None),
    (SEC, "D. 損益與現金流驅動（GAAP；公司 non-GAAP 只作起點與對照；比率基準全期持平於近四季，工作單 G2.3）"),
    ("ng_cogs", "非 GAAP 營業成本率（含內嵌 AI 推論）", "%", None, "=h_ngcogs_ltm/h_rev_ltm", 0.1639, "Derived", "近四季＝Q4 2025＋9M 2026（Hist 頁）", "基準全期持平；高案＝毛利下降 1pt（公司預期毛利率隨時間小幅下降，10-Q）", None),
    ("ng_opex", "非 GAAP 營業費用率（只用於 Q4 2026 單季；FY2027 起改依增量利益率 inc_m）", "%", None, "=h_ngopx_ltm/h_rev_ltm", None, "Derived", "近四季＝Q4 2025＋9M 2026（Hist 頁）", "Q4 為單季預測，持平於近四季", None),
    ("inc_m", "增量非 GAAP 營業利益率（營收每增 1 美元的營業利益增量）：FY2027 起", "%", 0.2083, "=(h_ngoi_25-h_ngoi_23)/(h_rev_25-h_rev_23)", 0.367, "Derived", "FY2023 至 FY2025 斜率：非 GAAP 營業利益 330.3→581.9、營收 2,170.2→3,131.3（Hist 頁）；非 GAAP 營業利益率 FY2022 9.8%、FY2023 15.2%、FY2024 17.5%、FY2025 18.6%、近四季 20.8%（歷史有規模經濟，工作單 G3.1）", "固定加變動成本拆分：營業利益＝近四季營業利益＋斜率×營收增量。低案＝0.2083＝全期持平於近四季費用率（下緣，工作單 G2.3）；高案＝FY2022→FY2023 斜率 36.7%。管理層 2027 目標 23–24%、2030 目標 30%（Interested-party）只在 Recon 對照；承重 Assumed 之一", None),
    ("sbc_pct", "股票薪酬占營收（只用於 Q4 2026 單季；FY2027 起改依增量 sbc_inc）", "%", None, "=h_sbc_ltm/h_rev_ltm", None, "Derived", "近四季現金流量表 SBC ÷ 營收（Hist 頁）", "Q4 為單季預測", None),
    ("sbc_inc", "增量股票薪酬率（營收每增 1 美元的 SBC 增量）：FY2027 起", "%", 0.046, "=(h_sbc_25-h_sbc_23)/(h_rev_25-h_rev_23)", 0.1404, "Derived", "FY2023 至 FY2025 斜率：SBC 432.3→528.2（Hist 頁）；SBC 占營收 FY2021 12.8%、FY2022 15.9%、FY2023 19.9%、FY2024 19.2%、FY2025 16.9%、近四季 14.0%", "SBC＝近四季 SBC＋斜率×營收增量。低案＝FY2024→FY2025 斜率 4.6%；高案＝0.1404＝全期持平於近四季占比（下緣）。DCF 不加回 SBC", None),
    ("adj_pct", "非 GAAP 調節其他項（收購攤銷、收購費用、重組、SBC 雇主稅等）占營收", "%", 0.01, "=(h_ngoi_ltm-h_oi_ltm-h_sbc_ltm)/h_rev_ltm", 0.02, "Derived", "近四季：非 GAAP 營業利益 − GAAP 營業利益 − SBC（Hist 頁）", "全期持平；Q4 2026 重組費用另列", None),
    ("restr_q4", "Q4 2026 重組費用（2026 Restructuring Plan，約 7% 員工）", "US$M", 65, 70, 75, "Interested-party", Q3Q + " 附註：預估 65–75 百萬美元，多數於 Q4 認列", "只影響 Q4 一次性費用與現金；後續節省不入基準（工作單 G2.3）", CO),
    ("ps_pct", "專業服務與其他營收 ÷ 訂閱營收", "%", 0.015, "=(h_ps_ltm)/(h_sub_ltm)", 0.025, "Derived", "近四季（Hist 頁）", "全期持平", None),
    ("da_pct", "折舊與攤銷占營收", "%", 0.04, "=h_da_ltm/h_rev_ltm", 0.055, "Derived", "近四季（Hist 頁）", None, None),
    ("cx_pct", "資本支出（含資本化軟體）占營收", "%", 0.05, "=(h_ppe_ltm+h_capsw_ltm)/h_rev_ltm", 0.07, "Derived", "近四季（Hist 頁）", None, None),
    ("nwc_k", "營運資金流入 ÷ 營收增量", "x", 0.0, "=h_nwc_ltm/(h_rev_ltm-h_rev_ltm1_ltm)", 0.19, "Derived", "近四季營運資金變動 48.6 ÷ 營收增量 599.4（Hist 頁）；高案＝FY2025 19.0%", "遞延營收為主", None),
    ("int_y", "現金與投資收益率", "%", 0.015, 0.026, 0.04, "Assumed", "9M26 利息收入 30.2；平均餘額未取得（Hist 頁僅期末）", "收益率由利息收入年化÷期末現金投資粗估，標 Assumed", None),
    ("bb_pct", "回購占自由現金流比例（FY2027 起）", "%", 0.5, 0.8, 1.0, "Assumed", "9M26 回購 973.7 對 FCF 約 495（Hist 頁）；以後降為 80%", "不影響 DCF（FCFF），只影響每股盈餘與期末現金", None),
    (SEC, "E. 公司指引、賣方共識與其他對照（只作對照，不進模型）"),
    ("g_rev_q4", "指引 Q4 2026 營收（中點）", "US$M", None, 946.0, None, "Interested-party", Q3PR + "：945–947", NOTM, CO),
    ("g_ngoi_q4", "指引 Q4 2026 非 GAAP 營業利益（中點）", "US$M", None, 216.5, None, "Interested-party", Q3PR + "：216–217", NOTM, CO),
    ("g_rev26", "指引 FY2026 營收（中點）", "US$M", None, 3686.0, None, "Interested-party", Q3PR + "：3,685–3,687", NOTM, CO),
    ("g_ngoi26", "指引 FY2026 非 GAAP 營業利益（中點）", "US$M", None, 772.5, None, "Interested-party", Q3PR + "：772–773", NOTM, CO),
    ("g_eps26", "指引 FY2026 非 GAAP EPS（中點）", "US$", None, 13.60, None, "Interested-party", Q3PR + "：13.59–13.61", NOTM, CO),
    ("g_m27", "管理層目標：2027 非 GAAP 營業利益率（中點）", "%", None, 0.235, None, "Interested-party", PPC + "：23–24%", NOTM, CO),
    ("c_rev26", "賣方共識 FY2026 營收", "US$M", None, 3690.0, None, "Interested-party", "https://twelvedata.com/markets/402519/stock/mta/1hubs/analysis（無日期，Q3 後）", NOTM, SELL),
    ("c_rev27", "賣方共識 FY2027 營收", "US$M", None, 4100.0, None, "Interested-party", "同上（32 位；區間 3,980–4,220）", NOTM, SELL),
    ("c_tp", "賣方平均目標價（49 位；低 190、高 300）", "US$", None, 243.18, None, "Interested-party", "https://tickflow.io/stock/HUBS/forecast（2026-11-16）", NOTM, SELL),
    ("c_pe_hubs", "HUBS 現行前瞻 P/E（ChartRow）", "x", None, 13.2, None, "Analogy", CR.format("hubs"), NOTM + "；口徑同同業（未核）", None),
    ("c_date", "共識資料日", "日期", None, D(2026, 11, 16), None, "Verified", "TickFlow 頁面標示", NOTM, None),
]
# 情境驅動對照：(鍵, 項目, 單位, 引用鍵(A,B,C))
SCN_ROWS = [
    ("scn_repl", "每年被 AI 代理取代的席位占比", "%", ("repl_A", "repl_B", "repl_C")),
    ("scn_adopt", "FY2032 付費採用 agent 的客戶占比", "%", ("adopt_A", "adopt_B", "adopt_C")),
    ("scn_dp", "每任務單價年變動", "%", ("dp_A", "dp_B", "dp_C")),
    ("scn_w", "情境權重", "%", ("tp_wA", "tp_wB", "tp_wC")),
]

# ───────── Hist ─────────
# 欄：C FY2023, D FY2024, E FY2025, F 9M25, G 9M26, H Q4'25, I LTM(Q4'25+9M26), J Q1'26, K Q2'26, L Q3'26
HCOLS = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]
HSUF = ["23", "24", "25", "9m25", "9m26", "q425", "ltm", "q126", "q226", "q326"]
HLAB = ["FY2023", "FY2024", "FY2025", "9M25", "9M26", "Q4'25", "近四季（Q4'25＋9M26）", "Q1'26", "Q2'26", "Q3'26"]
N = None
# (鍵, 項目, [FY23, FY24, FY25, 9M25, 9M26, Q4'25, (LTM 公式或 None), Q1'26, Q2'26, Q3'26], 標記, ltm 規則 'sum'|None|'v')
HIST = [
    ("sub", "訂閱營收", [2123.479, 2569.546, 3063.917, 2234.937, 2686.531, 828.980, "sum", 862.264, 894.025, 930.242], "Verified"),
    ("ps", "專業服務與其他營收", [46.751, 57.997, 67.349, 49.583, 53.770, 17.766, "sum", 18.731, 17.715, 17.323], "Verified"),
    ("rev", "營收合計", [2170.230, 2627.543, 3131.266, 2284.520, 2740.301, 846.746, "sum", 880.995, 911.740, 947.565], "Verified"),
    ("cogs", "營業成本（GAAP）", [338.362, 393.265, 508.487, N, "=SUM(J{r}:L{r})", 137.701, "sum", 145.694, 160.879, 164.680], "Verified"),
    ("rd", "研發", [617.745, 778.714, 905.943, N, N, N, N, N, N, N], "Verified"),
    ("sm", "銷售與行銷", [1068.560, 1218.844, 1379.376, N, N, N, N, N, N, N], "Verified"),
    ("ga", "一般與行政", [249.649, 300.332, 326.045, N, N, N, N, N, N, N], "Verified"),
    ("restr", "重組費用", [96.843, 3.990, 4.036, N, N, N, N, N, N, N], "Verified"),
    ("oi", "營業利益（GAAP）", [-200.929, -67.602, 7.379, N, 143.155, 48.238, "sum", 27.943, 43.319, 71.892], "Verified"),
    ("pt", "稅前利益", [-150.575, 28.677, 69.463, N, 179.068, N, N, N, N, 90.207], "Verified"),
    ("tx", "所得稅費用", [13.935, 24.049, 23.552, N, 20.205, N, N, N, N, 7.236], "Verified"),
    ("ni", "稅後淨利（GAAP）", [-164.510, 4.628, 45.911, N, "=SUM(J{r}:L{r})", 54.426, "sum", 32.554, 43.338, 82.971], "Verified"),
    ("eps", "EPS 稀釋（GAAP；US$）", [-3.30, 0.09, 0.86, N, "=SUM(J{r}:L{r})", 1.04, N, 0.62, 0.86, 1.67], "Verified"),
    ("dsh", "稀釋加權股數（百萬股）", [49.877, 51.819, 53.194, N, N, 52.547, N, 52.578, 50.615, 49.535], "Verified"),
    ("sbc", "股票薪酬（現金流量表）", [432.271, 504.770, 528.153, N, 371.347, 132.399, "sum", 115.744, 128.482, 127.121], "Verified"),
    ("da", "折舊與攤銷", [72.673, 96.828, 136.295, N, 130.813, 38.569, "sum", 40.239, 44.396, 46.178], "Verified"),
    ("ppe", "購置不動產與設備", [33.718, 37.939, 53.165, N, 51.095, 13.450, "sum", 15.422, 19.652, 16.021], "Verified"),
    ("capsw", "資本化軟體開發成本", [66.372, 89.636, 130.624, N, 112.994, 30.269, "sum", 34.339, 39.953, 38.702], "Verified"),
    ("cfo", "營業現金流", [350.971, 598.599, 760.717, N, 659.607, 247.430, "sum", 198.825, 222.755, 238.027], "Verified"),
    ("ngoi", "營業利益（非 GAAP）", [330.275, 460.211, 581.901, N, 556.257, 191.006, "sum", 156.764, 185.342, 214.154], "Interested-party"),
    ("ngcogs", "非 GAAP 營業成本（訂閱＋專業服務）", [N, N, N, N, 427.218, 124.874, "sum", N, N, N], "Interested-party"),
    ("ngopx", "非 GAAP 營業費用（研發＋銷售行銷＋一般行政）", [N, N, N, N, 1756.826, 530.866, "sum", N, N, N], "Interested-party"),
    ("ngeps", "非 GAAP 稀釋 EPS（US$）", [N, N, 9.70, N, "=SUM(J{r}:L{r})", 3.09, N, 2.72, 3.26, 3.81], "Interested-party"),
    ("nwc", "營運資金變動合計（現金流量表六項）", [-17.7, 56.8, 95.6, 61.5, 14.5, N, "=E{r}+G{r}-F{r}", N, N, N], "Derived"),
    ("cust", "Customers（期末；戶）", [205091, 247939, 288706, N, N, 288706, N, 299458, 306446, 311704], "Interested-party"),
    ("asrpc", "Average Subscription Revenue per Customer（US$）", [11384, 11343, 11414, N, N, 11683, N, 11722, 11800, 12038], "Interested-party"),
    ("intinc", "利息收入", [N, N, N, N, 30.209, N, N, N, N, N], "Verified"),
    ("cashinv", "現金＋短期投資＋長期投資（期末）", [N, N, N, N, N, N, N, N, N, 1310.503], "Verified"),
    ("rev_ltm1", "近四季前一年同期營收（Q4'24＋9M25；增量用）", [N, N, N, N, N, N, "=F{rev}+D{rev_q424}", N, N, N], "Derived"),
    ("rev_q424", "Q4 2024 營收（近四季增量用）", [N, 703.172, N, N, N, N, N, N, N, N], "Verified"),
    ("emp", "全職員工（期末）", [7663, 8246, 8882, N, N, N, N, N, N, N], "Verified"),
    ("nrr", "Net Revenue Retention（%；FY2023–25 新口徑）", [1.030, 1.018, 1.035, N, N, N, N, N, N, 1.01], "Interested-party"),
]

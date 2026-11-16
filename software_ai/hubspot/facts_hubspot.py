# HubSpot（HUBS）事實總帳：每個數字一列（指標、數值、單位、期間、口徑、出處、擷取日、標記、來源誘因）。
# 本檔是 data/facts_hubspot.csv、data/research_hubspot.md 與 xlsx Hist 頁的單一來源（改數字只改這裡）。
# 金額單位：百萬美元（US$M）；10-K／10-Q／新聞稿原表為千美元，除以 1,000。擷取日 2026-11-16。
# 用法：python3 software_ai/hubspot/facts_hubspot.py   → 寫 data/facts_hubspot.csv
import csv
import os

HERE = os.path.dirname(os.path.abspath(__file__))
ASOF = "2026-11-16"
E = "https://www.sec.gov/Archives/edgar/data/1404655/"
SRC = {
    "K25": (E + "000119312526046646/hubs-20251231.htm", "HubSpot 10-K FY2025（2026-02-11）"),
    "K23": (E + "000095017024015277/hubs-20231231.htm", "HubSpot 10-K FY2023（2024-02-14）"),
    "PR23": (E + "000095017024015231/hubs-ex99_1.htm", "8-K 附件 99.1 Q4／FY2023 新聞稿（2024-02-14）"),
    "PR_Q4": (E + "000119312526046563/hubs-ex99_1.htm", "8-K 附件 99.1 Q4／FY2025 新聞稿（2026-02-11）"),
    "PR_Q1": (E + "000119312526211923/hubs-ex99_1.htm", "8-K 附件 99.1 Q1 2026 新聞稿（2026-05-07）"),
    "PR_Q2": (E + "000119312526335148/hubs-ex99_1.htm", "8-K 附件 99.1 Q2 2026 新聞稿（2026-08-05）"),
    "PR_Q3": (E + "000119312526447249/hubs-ex99_1.htm", "8-K 附件 99.1 Q3 2026 新聞稿（2026-11-04）"),
    "Q3Q": (E + "000119312526447363/hubs-20260930.htm", "HubSpot 10-Q 2026Q3（2026-11-04）"),
    "PPC": ("https://ppc.land/hubspot-forecasts-q4-revenue-below-q3-as-growth-slows-to-12-from-17/", "PPC Land：HubSpot Q3 2026 法說摘要（第二來源，轉述法說與 10-Q）"),
    "WB": ("https://www.webull.com/news/15870491618810880", "Webull：HUBS Q3 2026 deep dive（第二來源，轉述法說）"),
    "HSP": ("https://www.hubspot.com/company-news/hubspots-customer-agent-and-prospecting-agent-now-you-pay-when-the-task-is-complete", "HubSpot 官網公告：Customer Agent 與 Prospecting Agent 按結果計價（頁面註記 2026-04-13 更新）"),
    "SAAS": ("https://www.saastr.com/hubspot-switching-ai-pricing-from-per-use-to-per-resolution/", "SaaStr：HubSpot 改以每次解決計價（第二來源）"),
    "CR": ("https://chartrow.com/quote/hubs", "ChartRow HUBS 報價頁"),
    "QQ": ("https://www.quiverquant.com/stock/HUBS/historical-prices/", "QuiverQuant HUBS 歷史收盤價"),
    "FRED": ("https://fred.stlouisfed.org/series/DGS10", "FRED DGS10（美國 10 年期公債固定到期殖利率；工作單 G5.2 統一值，2026-11-12 收盤 5.10%）"),
    "TE": ("https://tradingeconomics.com/united-states/government-bond-yield/news/598729", "Trading Economics 美國 10 年期公債殖利率新聞頁（含 2026-11-13 更新）"),
    "SA": ("https://stockanalysis.com/stocks/hubs/statistics/", "StockAnalysis HUBS statistics（2026-10-26）"),
    "MB": ("https://www.marketbeat.com/instant-alerts/hubspot-nysehubs-releases-quarterly-earnings-results-beats-estimates-by-024-eps-2026-08-05/", "MarketBeat HUBS 頁（Q2 後，無明確日期）"),
    "TF": ("https://tickflow.io/stock/HUBS/forecast", "TickFlow HUBS forecast（頁面標示 2026-11-16）"),
    "TG": ("https://www.tickergate.com/stocks/hubs/stock-forecast", "TickerGate HUBS forecast（無明確日期，最新分析師紀錄 2026-11-13）"),
    "TD": ("https://twelvedata.com/markets/402519/stock/mta/1hubs/analysis", "Twelve Data HUBS analysis（無明確日期；含 2026-09-18 至 11-06 分析師動作）"),
    "SAF": ("https://stockanalysis.com/stocks/hubs/forecast/", "StockAnalysis HUBS forecast（2026-10-23，財報前）"),
    "MS": ("https://uk.marketscreener.com/quote/stock/HUBSPOT-INC-45166580/finances/", "MarketScreener HUBS 財務頁（頁面時間 2026-10-12，財報前）"),
}
V, I, A, AS, D = "Verified", "Interested-party", "Analogy", "Assumed", "Derived"
INC_CO = "公司自訂指標／非 GAAP 口徑或管理層說法，傾向呈現成長與獲利改善；定義可變更"
INC_SELL = "賣方預估與目標價；賣方有評等與投行業務關係"
INC_MKT = "資料商頁面，單一來源或未標示日期／口徑"
F = []   # 事實列


def f(metric, val, unit, period, basis, key, loc, tag, inc="", asof=ASOF):
    url, doc = SRC[key]
    if tag == I and not inc:
        inc = INC_CO
    F.append([metric, val, unit, period, basis, f"{url}｜{doc}｜{loc}", asof, tag, inc if tag == I else ""])


# ---------------- 年度損益（FY：1 月 1 日至 12 月 31 日；US$M；GAAP） ----------------
YRS = ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025"]
HIST = {   # 指標: ([FY2021..FY2025], 單位, 來源 key 規則, 備註)
    "訂閱營收": [1258.319, 1690.538, 2123.479, 2569.546, 3063.917],
    "專業服務與其他營收": [42.339, 40.431, 46.751, 57.997, 67.349],
    "營收合計": [1300.658, 1730.969, 2170.230, 2627.543, 3131.266],
    "營業成本：訂閱": [211.132, 257.513, 283.675, 336.878, 445.336],
    "營業成本：專業服務與其他": [47.725, 56.746, 54.687, 56.387, 63.151],
    "營業成本合計": [258.857, 314.259, 338.362, 393.265, 508.487],
    "毛利": [1041.801, 1416.710, 1831.868, 2234.278, 2622.779],
    "研發": [301.970, 442.022, 617.745, 778.714, 905.943],
    "銷售與行銷": [649.681, 886.069, 1068.560, 1218.844, 1379.376],
    "一般與行政": [144.949, 197.720, 249.649, 300.332, 326.045],
    "重組費用": [0.0, 0.0, 96.843, 3.990, 4.036],
    "營業費用合計": [1096.600, 1525.811, 2032.797, 2301.880, 2615.400],
    "營業利益（GAAP）": [-54.799, -109.101, -200.929, -67.602, 7.379],
    "其他收入合計": [-19.019, 4.409, 50.354, 96.279, 62.084],
    "稅前利益": [-73.818, -104.692, -150.575, 28.677, 69.463],
    "所得稅費用": [4.019, 8.057, 13.935, 24.049, 23.552],
    "稅後淨利（GAAP）": [-77.837, -112.749, -164.510, 4.628, 45.911],
    "EPS 稀釋（GAAP）": [-1.66, -2.35, -3.30, 0.09, 0.86],
    "稀釋加權股數": [46.891, 48.065, 49.877, 51.819, 53.194],
    "股票薪酬": [166.761, 275.849, 432.271, 504.770, 528.153],
    "折舊與攤銷": [45.159, 58.150, 72.673, 96.828, 136.295],
    "購置不動產與設備": [28.726, 37.426, 33.718, 37.939, 53.165],
    "資本化軟體開發成本": [33.139, 44.345, 66.372, 89.636, 130.624],
    "營業現金流": [238.728, 273.174, 350.971, 598.599, 760.717],
}
for name, vals in HIST.items():
    unit = "US$" if name.startswith("EPS") else ("百萬股" if "股數" in name else "US$M")
    for y, v in zip(YRS, vals):
        old = y in ("FY2021", "FY2022")
        key = "K23" if old else "K25"
        loc = "合併損益表" if name not in ("股票薪酬", "折舊與攤銷", "購置不動產與設備", "資本化軟體開發成本", "營業現金流") else "合併現金流量表"
        if y == "FY2023" and name in ("營業成本：訂閱", "營業成本合計", "毛利", "營業利益（GAAP）", "稅前利益", "所得稅費用", "稅後淨利（GAAP）", "EPS 稀釋（GAAP）"):
            basis = "GAAP；FY2025 10-K 重編後數（FY2023 10-K 原數不同，見研究包第 6 節）"
        elif old:
            basis = "GAAP；FY2023 10-K 所載（FY2021–FY2022 未重編）"
        else:
            basis = "GAAP"
        if y == "FY2023" and name == "營業現金流":
            basis = "GAAP"
        f(name, v, unit, f"{y}（1/1–12/31）", basis, key, loc, V)

# 公司自揭 KPI（Interested-party）
KPI = {
    "Customers（期末）": ([135442, 167386, 205091, 247939, 288706], "戶", "K23,K23,K25,K25,K25"),
    "Average Subscription Revenue per Customer": ([10486, 11163, 11384, 11343, 11414], "US$", "K23,K23,K25,K25,K25"),
}
for name, (vals, unit, keys) in KPI.items():
    for y, v, k in zip(YRS, vals, keys.split(",")):
        f(name, v, unit, f"{y}（1/1–12/31）", "公司自訂 KPI（Item 7 Key Business Metrics）", k, "Item 7 Key Business Metrics", I)
for y, v, k, b in [("FY2021", 115.2, "K23", "舊口徑（含 Solutions Partner 佣金）"), ("FY2022", 110.3, "K23", "舊口徑（含 Solutions Partner 佣金）"),
                   ("FY2023", 103.0, "K25", "新口徑（2025 年起剔除不再符合佣金資格的 Solutions Partner 佣金；舊口徑 FY2023 103.9%、FY2024 102.2%）"),
                   ("FY2024", 101.8, "K25", "新口徑（舊口徑 102.2%）"), ("FY2025", 103.5, "K25", "新口徑")]:
    f("Net Revenue Retention", v, "%", f"{y}（1/1–12/31）", b, k, "Item 7 Key Business Metrics", I)
for y, v, k in [("FY2022", 7433, "K23"), ("FY2023", 7663, "K23"), ("FY2024", 8246, "K25"), ("FY2025", 8882, "K25")]:
    f("全職員工數（期末）", v, "人", f"{y}年底", "10-K 人力資源段／風險因素", k, "Item 1 Human Capital／Item 1A", V)
# 非 GAAP 營業利益（公司自訂口徑）
for y, v, k, loc in [("FY2022", 169.072, "PR23", "Reconciliation of non-GAAP operating income（FY2022 欄）"), ("FY2023", 330.275, "PR23", "同上（FY2023 原數；FY2025 10-K 重編 GAAP 後未見重編非 GAAP）"),
                     ("FY2024", 460.211, "PR_Q4", "Reconciliation of non-GAAP operating income（FY2024 欄）"), ("FY2025", 581.901, "PR_Q4", "同上（FY2025 欄）")]:
    f("營業利益（非 GAAP）", v, "US$M", f"{y}（1/1–12/31）", "非 GAAP：排除股票薪酬、收購無形資產攤銷、收購相關費用、重組", k, loc, I)
# 地理（營收）
GEO = {"Americas": [797.986, 1073.932, 1321.568, 1588.886, 1872.171], "Europe": [390.379, 507.507, 668.263, 825.598, 1010.968], "Asia Pacific": [112.293, 149.530, 180.399, 213.059, 248.127]}
for g, vals in GEO.items():
    for y, v in zip(YRS, vals):
        f(f"營收：{g}", v, "US$M", f"{y}（1/1–12/31）", "GAAP；依客戶所在地", "K23" if y in ("FY2021", "FY2022", "FY2023") and y != "FY2023" else "K25", "合併財報附註 11 地理資訊", V)

# ---------------- 季度（US$M；GAAP 除非註明）----------------
Q = {  # 期間: dict
    "Q4'25": dict(sub=828.980, ps=17.766, rev=846.746, cogs=137.701, oi=48.238, oing=191.006, ni=54.426, sbc=132.399, da=38.569, cfo=247.430, ppe=13.450, capsw=30.269, eps=1.04, dsh=52.547,
                  cust=288706, asrpc=11683, key="PR_Q4"),
    "Q1'26": dict(sub=862.264, ps=18.731, rev=880.995, cogs=145.694, oi=27.943, oing=156.764, ni=32.554, sbc=115.744, da=40.239, cfo=198.825, ppe=15.422, capsw=34.339, eps=0.62, dsh=52.578,
                  cust=299458, asrpc=11722, key="PR_Q1"),
    "Q2'26": dict(sub=894.025, ps=17.715, rev=911.740, cogs=160.879, oi=43.319, oing=185.342, ni=43.338, sbc=128.482, da=44.396, cfo=222.755, ppe=19.652, capsw=39.953, eps=0.86, dsh=50.615,
                  cust=306446, asrpc=11800, key="PR_Q2"),
    "Q3'26": dict(sub=930.242, ps=17.323, rev=947.565, cogs=164.680, oi=71.892, oing=214.154, ni=82.971, sbc=127.121, da=46.178, cfo=238.027, ppe=16.021, capsw=38.702, eps=1.67, dsh=49.535,
                  cust=311704, asrpc=12038, key="PR_Q3"),
}
QN = {"sub": ("訂閱營收", V), "ps": ("專業服務與其他營收", V), "rev": ("營收合計", V), "cogs": ("營業成本合計", V), "oi": ("營業利益（GAAP）", V), "oing": ("營業利益（非 GAAP）", I),
      "ni": ("稅後淨利（GAAP）", V), "sbc": ("股票薪酬", V), "da": ("折舊與攤銷", V), "cfo": ("營業現金流", V), "ppe": ("購置不動產與設備", V), "capsw": ("資本化軟體開發成本", V),
      "eps": ("EPS 稀釋（GAAP）", V), "dsh": ("稀釋加權股數", V), "cust": ("Customers（期末）", I), "asrpc": ("Average Subscription Revenue per Customer", I)}
for q, d in Q.items():
    for k, (name, tag) in QN.items():
        unit = "US$" if k in ("eps", "asrpc") else ("百萬股" if k == "dsh" else "戶" if k == "cust" else "US$M")
        basis = "非 GAAP" if k == "oing" else "公司自訂 KPI" if tag == I else "GAAP"
        f(name, d[k], unit, f"{q}（單季）", basis, d["key"], "新聞稿財報表與 Business Highlights", tag)
f("9M26 訂閱營收", 2686.531, "US$M", "9M26（2026-01-01–09-30）", "GAAP", "PR_Q3", "Consolidated Statements of Operations", V)
f("9M26 專業服務與其他營收", 53.770, "US$M", "9M26", "GAAP", "PR_Q3", "同上", V)
f("9M26 營收合計", 2740.301, "US$M", "9M26", "GAAP", "PR_Q3", "同上", V)
f("9M25 營收合計", 2284.520, "US$M", "9M25", "GAAP", "PR_Q3", "同上", V)
f("9M26 營業利益（GAAP）", 143.155, "US$M", "9M26", "GAAP", "PR_Q3", "同上", V)
f("9M26 營業利益（非 GAAP）", 556.257, "US$M", "9M26", "非 GAAP", "PR_Q3", "Reconciliation of non-GAAP operating income", I)
f("9M26 股票薪酬", 371.347, "US$M", "9M26", "GAAP（現金流量表）", "PR_Q3", "Consolidated Statements of Cash Flows", V)
f("9M26 非 GAAP 營業成本（訂閱＋專業服務）", 427.218, "US$M", "9M26", "非 GAAP＝383.255＋43.963", "PR_Q3", "Reconciliation of non-GAAP expense", I)
f("9M26 非 GAAP 營業費用（研發＋銷售行銷＋一般行政）", 1756.826, "US$M", "9M26", "非 GAAP＝493.139＋1,075.457＋188.230", "PR_Q3", "Reconciliation of non-GAAP expense", I)
f("Q4'25 非 GAAP 營業成本", 124.874, "US$M", "Q4'25", "非 GAAP＝109.643＋15.231", "PR_Q4", "Reconciliation of non-GAAP expense", I)
f("Q4'25 非 GAAP 營業費用", 530.866, "US$M", "Q4'25", "非 GAAP＝149.118＋322.389＋59.359", "PR_Q4", "Reconciliation of non-GAAP expense", I)
f("9M26 營運資金變動合計（現金流量表六項）", 14.5, "US$M", "9M26", "應收＋預付＋遞延佣金＋使用權資產＋應付＋應計＋租賃負債＋遞延營收", "PR_Q3", "Consolidated Statements of Cash Flows", D)
f("9M25 營運資金變動合計", 61.5, "US$M", "9M25", "同上", "PR_Q3", "同上", D)
f("FY2025 營運資金變動合計", 95.6, "US$M", "FY2025", "同上", "K25", "合併現金流量表", D)
f("FY2024 營運資金變動合計", 56.8, "US$M", "FY2024", "同上", "K25", "合併現金流量表", D)
f("FY2023 營運資金變動合計", -17.7, "US$M", "FY2023", "同上", "K25", "合併現金流量表", D)
f("9M26 利息收入", 30.209, "US$M", "9M26", "GAAP", "PR_Q3", "Consolidated Statements of Operations", V)
f("9M26 所得稅費用", 20.205, "US$M", "9M26", "GAAP", "PR_Q3", "同上", V)
f("9M26 稅前利益", 179.068, "US$M", "9M26", "GAAP", "PR_Q3", "同上", V)
f("Q3'26 稅前利益", 90.207, "US$M", "Q3'26", "GAAP", "PR_Q3", "同上", V)
f("Q3'26 所得稅費用", 7.236, "US$M", "Q3'26", "GAAP", "PR_Q3", "同上", V)
f("9M26 回購普通股（現金流量表）", 973.690, "US$M", "9M26", "GAAP", "PR_Q3", "Consolidated Statements of Cash Flows", V)
f("FY2025 回購普通股", 500.019, "US$M", "FY2025", "GAAP", "K25", "合併現金流量表", V)
f("現金與投資（2026-09-30）", 1310.503, "US$M", "2026-09-30", "現金 1,121.342＋短期投資 175.414＋長期投資 13.747", "PR_Q3", "Consolidated Balance Sheets", D)
f("借款（2026-09-30）", 0.0, "US$M", "2026-09-30", "循環信貸額度 5 億美元，無已動用餘額", "Q3Q", "Revolving Credit Facility 附註", V)
f("流通股數（2026-10-30）", 48.836461, "百萬股", "2026-10-30", "10-Q 封面", "Q3Q", "封面", V)
f("RSU 與 PSU（反稀釋排除，3 個月加權）", 2.525, "百萬股", "Q3'26", "附註：潛在稀釋股數（反稀釋）；非期末未歸屬餘額", "Q3Q", "Note 每股盈餘", D)
f("未歸屬 RSU＋PSU（2025-12-31）", 1.471, "百萬股", "2025-12-31", "RSU 1,396 千＋PSU 75 千", "K25", "Note 股票薪酬", V)
f("遞延營收（2026-09-30）", 1034.580, "US$M", "2026-09-30", "GAAP", "PR_Q3", "Consolidated Balance Sheets", V)
f("非 GAAP 自由現金流", 188.0, "US$M", "Q3'26", "公司定義：營業現金流 − 購置不動產設備 − 資本化軟體 ＋ 重組現金", "PR_Q3", "Reconciliation of free cash flow", I)
f("非 GAAP 訂閱毛利率", 0.854, "比率", "Q3'26", "非 GAAP（去年同期 86.7%）", "PR_Q3", "Reconciliation of non-GAAP subscription margin", I)
f("訂閱與託管成本年增（含 AI 功能使用量）", 23.246, "US$M", "Q3'26 對 Q3'25", "10-Q MD&A 成本變動表", "Q3Q", "MD&A Cost of revenue", V)
f("營業成本內容含 AI 推論", "訂閱營業成本含「AI inferencing costs」", "文字", "2026Q3", "10-Q 成本科目定義", "Q3Q", "Key Components of Consolidated Statements of Operations", V)
f("重組計畫（2026-10-01 董事會核准）", "裁撤約 7% 員工；預估費用 65–75 百萬美元，多數於 Q4 認列", "文字", "2026Q4", "10-Q 附註 Restructuring Plan", "Q3Q", "Restructuring Plan 附註", V)
f("營收：Americas（Q3'26）", 556.758, "US$M", "Q3'26", "GAAP", "Q3Q", "地理資訊附註", V)
f("營收：Europe（Q3'26）", 309.812, "US$M", "Q3'26", "GAAP", "Q3Q", "地理資訊附註", V)
f("營收：Asia Pacific（Q3'26）", 80.995, "US$M", "Q3'26", "GAAP", "Q3Q", "地理資訊附註", V)
f("營收：Americas／Europe／Asia Pacific（Q2'26）", "533.662／301.051／77.027", "US$M", "Q2'26", "GAAP", "PR_Q2", "地理資訊（取自 Q2 10-Q）", V)
f("營收：Americas／Europe／Asia Pacific（Q1'26）", "516.523／292.631／71.841", "US$M", "Q1'26", "GAAP", "PR_Q1", "地理資訊（取自 Q1 10-Q）", V)

# ---------------- 指引（Interested-party） ----------------
f("指引：Q4 2026 營收", "945.0–947.0", "US$M", "Q4 2026", "公司指引（2026-11-04）", "PR_Q3", "Business Outlook", I)
f("指引：Q4 2026 非 GAAP 營業利益", "216.0–217.0", "US$M", "Q4 2026", "公司指引", "PR_Q3", "Business Outlook", I)
f("指引：Q4 2026 非 GAAP EPS", "3.85–3.87", "US$", "Q4 2026", "公司指引；稀釋股數約 48.9 百萬", "PR_Q3", "Business Outlook", I)
f("指引：FY2026 營收", "3,685–3,687", "US$M", "FY2026", "公司指引", "PR_Q3", "Business Outlook", I)
f("指引：FY2026 非 GAAP 營業利益", "772.0–773.0", "US$M", "FY2026", "公司指引", "PR_Q3", "Business Outlook", I)
f("指引：FY2026 非 GAAP EPS", "13.59–13.61", "US$", "FY2026", "公司指引；稀釋股數約 50.4 百萬", "PR_Q3", "Business Outlook", I)
f("長期目標：2027 非 GAAP 營業利益率", "23%–24%", "比率", "FY2027", "管理層說法（法說）", "PPC", "Margin targets 段", I)
f("長期目標：2030 非 GAAP 營業利益率", "30%", "比率", "FY2030", "管理層說法（法說）", "PPC", "Margin targets 段", I)

# ---------------- AI 與定價（第 5 節） ----------------
f("Customer Agent 價格（2026-04-14 起）", 0.50, "US$／次解決對話", "2026-04-14 起", "原為每次對話 1.00 美元；50 HubSpot Credits／次解決", "HSP", "Customer Agent 段", V)
f("Prospecting Agent 價格（2026-04-14 起）", 1.00, "US$／筆推薦線索", "2026-04-14 起", "原為每月每筆登錄聯絡人收費；100 Credits／筆", "HSP", "Prospecting Agent 段", V)
f("Customer Agent 解決率與啟用客戶", "65% 對話由代理解決；啟用客戶 8,000 以上", "文字", "2026-04", "公司行銷宣稱", "HSP", "Performance claims 段", I)
f("Customer Agent 單季解決件數", "Q3 2026 超過 200 萬件", "文字", "Q3'26", "法說口頭說法（第二來源轉述）", "PPC", "AI 段", I)
f("使用 HubSpot agents 的 Pro 以上客戶占比", "20%（較前期 +11 個百分點）；使用 AI 功能 64%", "文字", "Q3'26", "法說口頭說法（第二來源轉述）", "PPC", "AI 段", I)
f("月度 agent 動作數與 credits 消耗", "agent 動作 5 倍；credits 消耗約 2.5 倍", "文字", "Q3'26", "法說口頭說法；基期未載（年增或季增未明）", "PPC", "AI 段", I)
f("客戶定價方案：seats and credits 試點", "新客戶限定；北歐 2026-04、荷比盧 2026-06、EMEA 全部新客戶 Q4 2026；現有客戶無時程；2027-02 公布 EMEA 試點結果", "文字", "2026", "法說口頭說法（第二來源轉述）；試點區 credits 用量 2–3 倍、初始 ASP 較低（未揭露數字）", "PPC", "Pricing Changes 段", I)
f("淨營收留存率（Net Revenue Retention）", 1.01, "比率", "Q3'26", "法說（第二來源轉述）；FY2026 預期較 FY2025 約降 2 個百分點；客戶金額留存「高 80% 區間」", "PPC", "Customers and Retention 段", I)
f("Customers 季淨增", "Q2'26 7,000；Q3'26 5,300；Q4'26 預期低於 5,000", "戶", "2026", "法說（第二來源轉述）；2026Q3 起客戶定義納入 YouSpot Pro 訂閱者（人數未揭露）", "PPC", "Customers and Retention 段", I)
f("Q4 2026 指引的管理層說明", "Q2 的逆風在 Q3 更明顯，預期延續到 2027；AI 改變客戶發現、評估與購買軟體的方式", "文字", "Q4 2026", "法說口頭說法（第二來源轉述）", "PPC", "Q4 Guidance 段", I)
# ---------------- 市場資料 ----------------
f("HUBS 收盤價", 217.25, "US$", "2026-11-13", "收盤", "QQ", "歷史收盤價表（ChartRow 同值 217.25）", V, asof="2026-11-16")
f("HUBS 收盤價（對照）", 217.25, "US$", "2026-11-13", "收盤；第二來源", "CR", "Latest close", V)
f("HUBS 收盤價", 215.84, "US$", "2026-11-12", "收盤", "QQ", "歷史收盤價表", V)
f("HUBS 收盤價", 216.64, "US$", "2026-11-04", "法說當日收盤", "QQ", "歷史收盤價表", V)
f("HUBS 收盤價", 204.96, "US$", "2026-11-05", "財報後首日收盤", "QQ", "歷史收盤價表", V)
f("美國 10 年期公債殖利率", 0.0510, "比率", "2026-11-12", "工作單 G5.2 統一值（FRED DGS10；現價日前 1 個交易日）；v0.1 曾用 Trading Economics 2026-11-13 約 5.15%", "FRED", "DGS10 2026-11-12", V)
f("HUBS β（5 年）", 1.24, "無單位", "2026-10-26", "StockAnalysis；未載報酬頻率", "SA", "Beta (5Y)", A, INC_MKT)
f("HUBS β", 1.19, "無單位", "Q2 後（無日期）", "MarketBeat", "MB", "Beta", A, INC_MKT)
f("同業前瞻 P/E：Salesforce", 13.6, "倍", "2026-11-13", "ChartRow Forward P/E（口徑與日期以頁面為準，疑為非 GAAP 一年期前瞻；未核）", "CR", "https://chartrow.com/quote/crm", A, INC_MKT)
PEERS = [("Salesforce（CRM）", 13.6, 4.9, 10.8, "crm"), ("ServiceNow（NOW）", 21.3, 7.3, 20.5, "now"), ("Workday（WDAY）", 14.4, 4.7, 12.8, "wday"), ("Adobe（ADBE）", 8.9, 3.7, 12.9, "adbe"),
         ("Intuit（INTU）", 12.6, 4.4, 13.7, "intu"), ("Atlassian（TEAM）", 22.4, 5.5, 23.7, "team"), ("monday.com（MNDY）", 10.2, 2.0, 26.7, "mndy"), ("DocuSign（DOCU）", 15.4, 4.4, 9.4, "docu"),
         ("Twilio（TWLO）", 30.8, 5.6, 19.9, "twlo"), ("Zoom（ZM）", 15.4, 5.7, 4.9, "zm"), ("Datadog（DDOG）", 57.9, 15.6, 35.2, "ddog")]
F.pop()   # 上方單列示例改由 PEERS 迴圈寫入
for n, pe, evs, gr, tk in PEERS:
    f(f"同業前瞻 P/E：{n}", pe, "倍", "2026-11-13", "ChartRow Forward P/E（口徑未核）", "CR", f"https://chartrow.com/quote/{tk} Forward P/E", A, INC_MKT)
    f(f"同業 EV/Sales（TTM）：{n}", evs, "倍", "2026-11-13", "ChartRow", "CR", f"https://chartrow.com/quote/{tk} EV/Sales", A, INC_MKT)
    f(f"同業營收年增：{n}", gr, "%", "最近期", "ChartRow（期間依頁面，未統一）", "CR", f"https://chartrow.com/quote/{tk} Revenue growth", A, INC_MKT)
f("HUBS 前瞻 P/E／EV/Sales（TTM）", "13.2 倍／2.7 倍（市值 106 億美元、EV 98 億美元）", "文字", "2026-11-13", "ChartRow", "CR", "Valuation multiples", A, INC_MKT)
# ---------------- 共識（Interested-party；賣方）----------------
f("賣方平均目標價", 243.18, "US$", "2026-11-16", "49 位分析師；最低 190、最高 300；評等 Hold（2.71）", "TF", "Analyst price targets", I, INC_SELL)
f("賣方平均目標價（對照）", 245.25, "US$", "約 2026-11-13", "47 位分析師；最低 190、最高 300", "TG", "Price target", I, INC_SELL)
f("賣方平均目標價（財報前）", 248.77, "US$", "2026-10-23", "35 位（S&P Global）；最低 190、最高 320", "SAF", "Price targets", I, INC_SELL)
f("共識營收 FY2026", 3690.0, "US$M", "FY2026", "32 位；財報後（本季當期共識 946.97 百萬）", "TD", "Revenue estimates", I, INC_SELL)
f("共識營收 FY2027", 4100.0, "US$M", "FY2027", "32 位（區間 3,980–4,220）；頁面無日期，因當季營收共識 946.97 百萬判斷為 Q3 財報後", "TD", "Revenue estimates", I, INC_SELL)
f("共識營收 FY2027（財報前）", 4198.0, "US$M", "FY2027", "財報前（頁面時間 2026-10-12）", "MS", "Net sales 2027", I, INC_SELL)
f("共識營收 FY2028（財報前）", 4769.0, "US$M", "FY2028", "財報前（頁面時間 2026-10-12）", "MS", "Net sales 2028", I, INC_SELL)
f("共識非 GAAP EPS FY2027（財報前）", 16.85, "US$", "FY2027", "StockAnalysis 摘要框「EPS Next Year」；財報前 2026-10-23；FY2027 財報後值未找到", "SAF", "EPS Next Year", I, INC_SELL)
f("共識 EBIT FY2027／FY2028（財報前）", "966.3／1,178", "US$M", "FY2027／FY2028", "MarketScreener；財報前；口徑未載", "MS", "EBIT", I, INC_SELL)
f("共識非 GAAP EPS FY2026", 13.30, "US$", "FY2026", "StockAnalysis；財報前 2026-10-23", "SAF", "FY2026 forecast", I, INC_SELL)


NOT_FOUND = [
    "Pro 以上客戶占全體 Customers 的比例（法說只給「Pro 以上客戶中 20% 使用 agents」，分母未揭露）",
    "付費席位（seats）數與每席位年費（HubSpot 只揭露 Customers 與 Average Subscription Revenue per Customer；本模型以 Customer 為計算單位）",
    "HubSpot Credits 總營收、Customer Agent 與 Prospecting Agent 營收、AI 相關 ARR（公司未揭露；10-Q 只說 credits 屬訂閱營收的一部分）",
    "Customer Agent 每季解決件數的完整歷史（只取得 2026Q3 超過 200 萬件的口頭說法）",
    "seats and credits 新方案對 ASP 的影響幅度（法說只說試點區 credits 用量 2–3 倍、初始 ASP 較低，無數字）；2027-02 公布試點結果",
    "FY2021 非 GAAP 營業利益與 FY2021 年底員工數（本代理未取得 FY2021 年報與 FY2022 年報）",
    "FY2023 在 FY2025 10-K 與 FY2023 10-K 的營業損失不同（−200.929 對 −208.056；訂閱營業成本 283.675 對 290.802）的重編原因（10-K 擷取段落未找到說明）",
    "Q3 財報後的 FY2027、FY2028 EPS 共識與 FY2028 營收共識（取得的頁面不是未標日期，就是財報前或欄位互相矛盾，不採用）",
    "選擇權履約價（10-Q 只載潛在稀釋選擇權 93 千股，履約價未載；基準以 0 計）",
    "未歸屬 RSU 與 PSU 的 2026-09-30 期末餘額（10-Q 只載反稀釋排除之加權股數 2.525 百萬；2025 年底為 1.471 百萬）",
    "同業前瞻 P/E 的 EPS 口徑（GAAP 或非 GAAP）、NTM 起訖與逐一核對（ChartRow 頁面只給 Forward P/E 數字）",
    "10 年期公債殖利率的第二個來源（財政部與 FRED 頁面在本環境無法讀取）",
]

if __name__ == "__main__":
    out = os.path.join(HERE, "data", "facts_hubspot.csv")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["指標", "數值", "單位", "期間", "口徑", "出處", "擷取日", "標記", "來源誘因"])
        w.writerows(F)
    print(f"facts_hubspot.csv：{len(F)} 列")

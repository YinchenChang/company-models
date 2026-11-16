#!/usr/bin/env python3
"""產生 facts_paloalto.csv（逐列紀錄）。研究包第 4 條格式；建置腳本 build_paloalto.py 由此 CSV 讀歷史值。
用法：python3 software_ai/paloalto/data/make_research.py   （輸出同資料夾 facts_paloalto.csv）"""
import csv, os
D = os.path.dirname(os.path.abspath(__file__))
TODAY = "2026-11-16"
U = {  # 出處縮寫 → 完整出處
 "K26": "PANW 10-K FY2026（2026-09-10 提交）https://www.sec.gov/Archives/edgar/data/1327567/000132756726000023/panw-20260731.htm",
 "R26Q4": "PANW 8-K 附件 99.1 FY2026 Q4 財報新聞稿（2026-09-01）https://www.sec.gov/Archives/edgar/data/0001327567/000132756726000019/ex991q426earningsrelease.htm",
 "R26Q3": "PANW FY2026 Q3 財報新聞稿 https://www.paloaltonetworks.com/company/press/2026/palo-alto-networks-reports-fiscal-third-quarter-2026-financial-results",
 "R26Q2": "PANW FY2026 Q2 財報新聞稿 https://investors.paloaltonetworks.com/node/20246/pdf",
 "R26Q1": "PANW FY2026 Q1 財報新聞稿（2025-11-19）https://www.paloaltonetworks.com/company/press/2025/palo-alto-networks-reports-fiscal-first-quarter-2026-financial-results",
 "R25Q4": "PANW FY2025 Q4 財報新聞稿 https://www.paloaltonetworks.com/company/press/2025/palo-alto-networks-reports-fiscal-fourth-quarter-and-fiscal-year-2025-financial-results",
 "R23Q4": "PANW 8-K 附件 99.1 FY2023 Q4 財報新聞稿 https://www.sec.gov/Archives/edgar/data/1327567/000132756723000020/ex991q423earningsrelease.htm",
 "R22Q4": "PANW 8-K 附件 99.1 FY2022 Q4 財報新聞稿 https://www.sec.gov/Archives/edgar/data/1327567/000132756722000020/ex991q422earningsrelease.htm",
 "PXY": "PANW 2026 年股東會委託書 DEF 14A https://investors.paloaltonetworks.com/node/20946/html",
 "SA": "StockAnalysis.com PANW 預測頁（S&P Global 資料）https://stockanalysis.com/stocks/panw/forecast/",
 "SAS": "StockAnalysis.com PANW 統計頁 https://stockanalysis.com/stocks/panw/statistics/",
 "MS": "MarketScreener PANW 共識頁 https://www.marketscreener.com/quote/stock/PALO-ALTO-NETWORKS-INC-11067980/consensus/",
 "YF": "Yahoo Finance PANW 統計頁（頁面未標日期）https://finance.yahoo.com/quote/PANW/key-statistics/",
 "WS": "247WallSt 2026-09-24 https://247wallst.com/investing/2026/09/24/palo-alto-is-knocking-on-400-as-wall-street-rushes-into-cyber-buy-hold-or-sell/",
 "PH": "Phaetrix Substack 2026-02-16 https://phaetrix.substack.com/p/palo-alto-networks-panw-solid-execution",
}
ROWS = []
def add(metric, val, unit, period, basis, src, tag="Verified", inc="", where=""):
    s = U.get(src, src) + (f"；{where}" if where else "")
    ROWS.append([metric, val, unit, period, basis, s, TODAY, tag, inc])
FY = "FY（8 月 1 日至翌年 7 月 31 日）"
def fyrow(metric, vals, unit, basis, src, where, tag="Verified", inc=""):
    for p, v in vals.items():
        if v is not None: add(metric, v, unit, p, basis, src if isinstance(src, str) else src[p], tag, inc, where)
# ---- 營收（法定申報／新聞稿）
SRC_FY = {"FY2022": "R22Q4", "FY2023": "R23Q4", "FY2024": "R25Q4", "FY2025": "R25Q4", "FY2026": "R26Q4"}
fyrow("Product 營收", {"FY2022": 1363.1, "FY2023": 1578.4, "FY2024": 1603.3, "FY2025": 1801.9, "FY2026": 2280}, "US$M", "GAAP", SRC_FY, "損益表")
fyrow("Subscription and support 營收", {"FY2022": 4138.4, "FY2023": 5314.3, "FY2024": 6424.2, "FY2025": 7419.6, "FY2026": 9200}, "US$M", "GAAP", SRC_FY, "損益表")
fyrow("合併營收", {"FY2022": 5501.5, "FY2023": 6892.7, "FY2024": 8027.5, "FY2025": 9221.5, "FY2026": 11480}, "US$M", "GAAP", SRC_FY, "損益表")
for p, v in {"FY2024": 4188, "FY2025": 4974, "FY2026": 6239}.items(): add("其中 Subscription 營收", v, "US$M", p, "GAAP", "K26", where="Item 7 營收表")
for p, v in {"FY2024": 2236, "FY2025": 2445, "FY2026": 2961}.items(): add("其中 Support 營收", v, "US$M", p, "GAAP", "K26", where="Item 7 營收表")
for p, v in {"FY2026": 7679, "FY2025": 6205, "FY2024": 5483}.items(): add("Americas 營收", v, "US$M", p, "GAAP", "K26", where="Item 7 地區表")
for p, v in {"FY2026": 2428, "FY2025": 1917, "FY2024": 1602}.items(): add("EMEA 營收", v, "US$M", p, "GAAP", "K26", where="Item 7 地區表")
for p, v in {"FY2026": 1373, "FY2025": 1099, "FY2024": 942}.items(): add("APAC 營收", v, "US$M", p, "GAAP", "K26", where="Item 7 地區表")
for p, (a, b, c) in {"FY2024": (1603, 6424, 8027), "FY2025": (1802, 7419, 9221), "FY2026": (2280, 9200, 11480)}.items():
    add("Product 營收（10-K 口徑，百萬整數）", a, "US$M", p, "GAAP", "K26", where="Item 7 營收表")
    add("Subscription and support 營收（10-K 口徑，百萬整數）", b, "US$M", p, "GAAP", "K26", where="Item 7 營收表")
    add("合併營收（10-K 口徑，百萬整數）", c, "US$M", p, "GAAP", "K26", where="Item 7 營收表")
# ---- 季度（FY2026 Q1–Q4）
Q = {"FY2026Q1": ("R26Q1", 434, 2040, 2474, 309, 746, 334, 0.47, 0.93, 387, 5.9, 15.5),
     "FY2026Q2": ("R26Q2", 514, 2080, 2594, 397, 785, 432, 0.61, 1.03, 321, 6.3, 16.0),
     "FY2026Q3": ("R26Q3", 594, 2408, 3002, -183, 814, -177, -0.22, 0.85, 517, 8.1, 18.4),
     "FY2026Q4": ("R26Q4", 738, 2672, 3410, 172, 1011, -282, -0.35, 1.02, 487, 9.1, 21.2)}
for p, (s, pr, ss, tot, oi, ngoi, ni, eps, ngeps, sbc, arr, rpo) in Q.items():
    add("Product 營收", pr, "US$M", p, "GAAP", s, where="損益表")
    add("Subscription and support 營收", ss, "US$M", p, "GAAP", s, where="損益表")
    add("合併營收", tot, "US$M", p, "GAAP", s, where="損益表")
    add("GAAP 營業利益", oi, "US$M", p, "GAAP", s)
    add("Non-GAAP 營業利益", ngoi, "US$M", p, "Non-GAAP（公司定義）", s, "Interested-party", "公司自訂 non-GAAP 調節，排除股票薪酬、併購攤銷等", "GAAP 至 non-GAAP 調節表")
    add("GAAP 淨利", ni, "US$M", p, "GAAP", s)
    add("GAAP 稀釋 EPS", eps, "US$", p, "GAAP（FY2025 起為 2 拆 1 後口徑）", s)
    add("Non-GAAP 稀釋 EPS", ngeps, "US$", p, "Non-GAAP（公司定義）", s, "Interested-party", "公司自訂 non-GAAP", "GAAP 至 non-GAAP 調節表")
    add("SBC-related charges", sbc, "US$M", p, "Non-GAAP 調節項", s)
    add("NGS ARR（期末）", arr, "US$B", p, "公司 KPI（定義見研究包第 5 節）", s, "Interested-party", "公司自揭 KPI，無第三方查核；定義可變（見 NGS ARR 定義列）", "Financial Highlights")
    add("RPO（期末）", rpo, "US$B", p, "公司 KPI（ASC 606 剩餘履約義務）", s, "Interested-party", "公司自揭 KPI；法定附註（10-K Note 2）僅 FY2026 年末可核對", "Financial Highlights")
add("NGS ARR（期末）", 5.6, "US$B", "FY2025Q4", "公司 KPI", "R25Q4", "Interested-party", "公司自揭 KPI", "Key metrics")
add("RPO（期末）", 15.8, "US$B", "FY2025Q4", "公司 KPI", "R25Q4", "Interested-party", "公司自揭 KPI", "Key metrics")
add("NGS ARR 年增率（期末）", 0.63, "比率", "FY2026Q4", "公司 KPI（含 CyberArk、Chronosphere 併入）", "R26Q4", "Interested-party", "公司自揭 KPI；含併購", "Q4 Financial Highlights")
add("RPO 年增率", 0.34, "比率", "FY2026Q4", "公司 KPI", "R26Q4", "Interested-party", "公司自揭 KPI", "Q4 Financial Highlights")
add("本季淨新增 NGS ARR", "接近 1（CEO：nearly $1 billion）", "US$B", "FY2026Q4", "公司口述", "R26Q4", "Interested-party", "CEO 引言", "CEO 引言")
add("下 12 個月預期認列的 RPO", 9.3, "US$B", "FY2026 年末", "ASC 606", "K26", where="Note 2 Revenue")
# ---- 損益科目（年）
fyrow("GAAP 毛利", {"FY2022": 3782.8, "FY2023": 4983.0, "FY2024": 5968.3, "FY2025": 6769.9, "FY2026": 8077}, "US$M", "GAAP", SRC_FY, "損益表")
fyrow("總營業成本（Cost of revenue）", {"FY2024": 2059.2, "FY2025": 2451.6, "FY2026": 3403}, "US$M", "GAAP", SRC_FY, "損益表")
fyrow("營業費用合計", {"FY2024": 5284.4, "FY2025": 5527.0, "FY2026": 7382}, "US$M", "GAAP", SRC_FY, "損益表")
fyrow("GAAP 營業利益", {"FY2022": -188.8, "FY2023": 387.3, "FY2024": 683.9, "FY2025": 1242.9, "FY2026": 695}, "US$M", "GAAP", SRC_FY, "損益表")
fyrow("其他收入（費用）淨額", {"FY2024": 312.7, "FY2025": 355.8, "FY2026": -159}, "US$M", "GAAP", SRC_FY, "損益表；FY2026 含可轉債與 capped call 公允價值損失 562")
fyrow("稅前利益", {"FY2024": 988.3, "FY2025": 1595.7, "FY2026": 536}, "US$M", "GAAP", SRC_FY, "損益表")
fyrow("所得稅（費用）", {"FY2024": -1589.3, "FY2025": 461.8, "FY2026": 229}, "US$M", "GAAP", SRC_FY, "損益表；FY2024 含估值備抵釋放")
fyrow("GAAP 淨利", {"FY2022": -267.0, "FY2023": 439.7, "FY2024": 2577.6, "FY2025": 1133.9, "FY2026": 307}, "US$M", "GAAP", SRC_FY, "損益表")
fyrow("GAAP 稀釋 EPS", {"FY2024": 3.64, "FY2025": 1.60, "FY2026": 0.40}, "US$", "GAAP（2 拆 1 後口徑）", SRC_FY, "損益表")
fyrow("GAAP 稀釋加權股數", {"FY2024": 707.9, "FY2025": 709.3, "FY2026": 764}, "百萬股", "GAAP", SRC_FY, "損益表／10-K Note 17")
fyrow("Non-GAAP 營業利益", {"FY2025": 2652, "FY2026": 3356}, "US$M", "Non-GAAP（公司定義）", "R26Q4", "GAAP 至 non-GAAP 調節表", "Interested-party", "公司自訂 non-GAAP")
fyrow("Non-GAAP 淨利", {"FY2024": 1948.1, "FY2025": 2344.5, "FY2026": 2931}, "US$M", "Non-GAAP（公司定義）", SRC_FY, "調節表", "Interested-party", "公司自訂 non-GAAP")
fyrow("Non-GAAP 稀釋 EPS", {"FY2024": 2.84, "FY2025": 3.34, "FY2026": 3.84}, "US$", "Non-GAAP（公司定義；2 拆 1 後）", SRC_FY, "調節表", "Interested-party", "公司自訂 non-GAAP")
fyrow("Non-GAAP 稀釋股數", {"FY2024": 687.2, "FY2025": 702.0, "FY2026": 764}, "百萬股", "Non-GAAP", SRC_FY, "調節表", "Interested-party", "公司自訂 non-GAAP")
fyrow("SBC-related charges", {"FY2021": 936.5, "FY2022": 1072.0, "FY2023": 1145.1, "FY2024": 1161.7, "FY2025": 1386.4, "FY2026": 1712}, "US$M", "Non-GAAP 調節項", {"FY2021": "R22Q4", "FY2022": "R22Q4", "FY2023": "R23Q4", "FY2024": "R25Q4", "FY2025": "R25Q4", "FY2026": "R26Q4"}, "調節表")
for p, v in {"FY2024": 1079, "FY2025": 1300, "FY2026": 1815}.items(): add("股票薪酬費用（10-K MD&A）", v, "US$M", p, "GAAP", "K26", where="Item 7 Share-Based Compensation 註腳")
add("股票薪酬費用（現金流量表）", 1774, "US$M", "FY2026", "GAAP", "K26", where="現金流量表（與 MD&A 1,815 差 41，未釐清）")
add("股票薪酬 ÷ 營收", 0.158, "比率", "FY2026", "公司披露", "PXY", "Interested-party", "公司在委託書自揭，用於說明薪酬治理", "董事長信")
add("股票薪酬 ÷ 營收", 0.218, "比率", "FY2021", "公司披露", "PXY", "Interested-party", "同上", "董事長信")
add("尚未認列的 SBC 費用", 3300, "US$M", "FY2026 年末", "GAAP", "K26", where="Item 7 SBC 註腳；加權平均約 2.5 年認列")
for p, v in {"FY2025": 164, "FY2026": 638}.items(): add("收購無形資產攤銷", v, "US$M", p, "Non-GAAP 調節項", "R26Q4", where="調節表")
add("收購無形資產攤銷", 281, "US$M", "FY2026Q4", "Non-GAAP 調節項", "R26Q4", where="調節表")
for p, v in {"FY2025": -110, "FY2026": 295}.items(): add("併購相關費用", v, "US$M", p, "Non-GAAP 調節項", "R26Q4", where="調節表；FY2025 為負值（或有對價調整）")
add("折舊與攤銷（現金流量表）", 855, "US$M", "FY2026", "GAAP", "K26", where="現金流量表")
add("折舊與攤銷（現金流量表）", 343, "US$M", "FY2025", "GAAP", "K26", where="現金流量表")
for p, (a, b, c) in {"FY2024": (0.743, 0.783, 0.734), "FY2025": (0.734, 0.771, 0.725), "FY2026": (0.704, 0.751, 0.692)}.items():
    add("GAAP 毛利率（合計）", a, "比率", p, "GAAP", "K26", where="Item 7 Gross Margin")
    add("GAAP 毛利率（Product）", b, "比率", p, "GAAP", "K26", where="Item 7 Gross Margin")
    add("GAAP 毛利率（Subscription and support）", c, "比率", p, "GAAP", "K26", where="Item 7 Gross Margin")
add("有效稅率（GAAP）", 0.427, "比率", "FY2026", "GAAP", "K26", where="Item 7 所得稅")
add("已付所得稅（淨額）", 199, "US$M", "FY2026", "現金", "K26", where="所得稅附註")
# ---- 現金流
fyrow("營業活動現金流", {"FY2024": 3258, "FY2025": 3716, "FY2026": 4553}, "US$M", "GAAP", {"FY2024": "K26", "FY2025": "R26Q4", "FY2026": "R26Q4"}, "現金流量表")
fyrow("資本支出（PP&E 與其他資產）", {"FY2025": 247, "FY2026": 440}, "US$M", "GAAP", "R26Q4", "FCF 調節表")
fyrow("自由現金流（公司定義）", {"FY2024": 3101, "FY2025": 3469, "FY2026": 4113}, "US$M", "Non-GAAP（OCF − capex）", {"FY2024": "K26", "FY2025": "R26Q4", "FY2026": "R26Q4"}, "FCF 調節表", "Interested-party", "公司定義的 FCF，SBC 不扣")
fyrow("調整後自由現金流（公司定義）", {"FY2025": 3507, "FY2026": 4414}, "US$M", "Non-GAAP", "R26Q4", "FCF 調節表", "Interested-party", "公司定義，加回總部購地、特定資產支出、併購付款")
add("調整後自由現金流率", 0.384, "比率", "FY2026", "Non-GAAP", "R26Q4", "Interested-party", "公司定義", "FCF 調節表")
add("總部購地資本支出", 91, "US$M", "FY2026", "GAAP", "R26Q4", where="FCF 調節表")
# ---- 資產負債表（2026-07-31）
for m, v in {"現金及約當現金": 2514, "短期投資": 557, "長期投資": 4835, "長期可轉換票據（公允價值）": 1774, "遞延收入（流動）": 7747, "遞延收入（長期）": 7009, "商譽": 22010, "無形資產淨額": 7017, "長期營業租賃負債": 726}.items():
    add(m, v, "US$M", "2026-07-31", "GAAP", "R26Q4", where="資產負債表")
for m, v in {"現金及約當現金": 2268.6, "短期投資": 634.6, "長期投資": 5555.6, "遞延收入（流動）": 6302.2, "遞延收入（長期）": 6449.7}.items():
    add(m, v, "US$M", "2025-07-31", "GAAP", "R25Q4", where="資產負債表")
add("Capped call 公允價值", 153, "US$M", "2026-07-31", "公允價值（Level 2）", "K26", where="Note 3")
add("2030 可轉債剩餘本金", 1100, "US$M", "2026-07-31", "面額", "K26", where="Note 11；CyberArk 原發行 1,250，票面 0%，2030-06-15 到期，換股價約 211.24")
add("2030 可轉債轉換股數", 5.4, "百萬股", "2026-07-31", "初始", "K26", where="Note 11")
add("未認列採購承諾（含雲端代管 7,686）", 8242, "US$M", "2026-07-31", "合約", "K26", where="Note 13")
add("流通股數", 818, "百萬股", "2026-08-31", "封面", "K26", where="封面")
add("未歸屬 RSU", 17, "百萬股", "2026-07-31", "全額", "K26", where="Note 15（加權授予日公允價值 177.41）")
add("未歸屬 PSU", 8, "百萬股", "2026-07-31", "績效型", "K26", where="Note 15（加權授予日公允價值 155.85）")
add("選擇權", 0, "百萬股", "2026-07-31", "—", "K26", where="Note 15：無流通在外")
add("FY2026 買回金額", 1000, "US$M", "FY2026", "現金流量表", "K26", where="約 7 百萬股")
add("買回授權剩餘", 1000, "US$M", "2026-07-31", "—", "K26", where="Item 5／7；授權期限延至 2026-12-31")
fyrow("員工數", {"FY2025": 16068, "FY2026": 21921}, "人", "期末", "K26", "Item 1 Human Capital（含 CyberArk 併入約 4,223）")
# ---- 併購
add("CyberArk 併購總對價", 21061, "US$M", "FY2026（2026-02-11 完成）", "購買價格", "K26", where="Note 8：現金 2,308＋股票 18,488（112 百萬股）；每股 CyberArk 得 45.00 美元現金＋2.2005 股 PANW")
add("CyberArk 併購：商譽", 14800, "US$M", "FY2026", "購買價格分攤", "K26", where="Note 8")
add("CyberArk 併購：可辨認無形資產", 6279, "US$M", "FY2026", "購買價格分攤", "K26", where="Note 8")
add("Chronosphere 併購總對價", 2951, "US$M", "FY2026（2026-01-29 完成）", "購買價格", "K26", where="Note 8，現金 2,842")
add("Chronosphere＋CyberArk 併入後貢獻營收", 930, "US$M", "FY2026（自各自完成日）", "GAAP", "K26", where="Note 8")
add("Chronosphere＋CyberArk 併入後營業損失", -797, "US$M", "FY2026（自各自完成日）", "GAAP", "K26", where="Note 8")
add("Koi Security 併購對價", 231, "US$M", "FY2026（2026-04-14）", "購買價格", "K26", where="Note 8")
add("Portkey 併購對價", 117, "US$M", "FY2026（2026-05-29）", "購買價格", "K26", where="Note 8")
add("Embrace Mobile 併購對價", 325, "US$M", "FY2027Q1（2026-08-27 完成）", "現金，待調整", "K26", where="Note 8／20")
add("Console Systems 併購對價", 500, "US$M", "FY2027Q1（2026-09-01 完成）", "現金，待調整", "K26", where="Note 8／20")
add("報告分部數", 1, "個", "FY2026", "ASC 280", "K26", where="Note 19：一個應報導分部；CODM 以合併淨利衡量")
# ---- 推導
add("有機營收（扣除 Chronosphere 與 CyberArk 併入後貢獻）", 10550, "US$M", "FY2026", "合併營收 11,480 − 930", "由 K26 Note 8 與 R26Q4 相減", "Derived", "", "")
add("有機營收年增率", 0.1441, "比率", "FY2026", "10,550 ÷ 9,221.5 − 1", "同上", "Derived", "", "")
add("Non-GAAP 營業利益率", 0.2923, "比率", "FY2026", "3,356 ÷ 11,480", "R26Q4", "Derived", "", "")
add("Non-GAAP 所得稅率（推導）", 0.22, "比率", "FY2026", "（3,356＋403−2,931）÷（3,356＋403），其中 403＝其他收入淨額 −159＋公允價值損失 562＝403", "R26Q4", "Derived", "", "")
# ---- 指引與目標（公司自揭）
G = "R26Q4"
for m, v, u in [("FY2027 營收指引（低）", 14100, "US$M"), ("FY2027 營收指引（高）", 14200, "US$M"), ("FY2027 NGS ARR 指引（低）", 11.075, "US$B"), ("FY2027 NGS ARR 指引（高）", 11.175, "US$B"),
                ("FY2027 Non-GAAP 營業利益率指引", 0.295, "比率"), ("FY2027 Non-GAAP 稀釋 EPS 指引（低）", 4.16, "US$"), ("FY2027 Non-GAAP 稀釋 EPS 指引（高）", 4.19, "US$"),
                ("FY2027 調整後自由現金流率指引", 0.38, "比率"), ("FY2027 稀釋股數指引（低）", 844, "百萬股"), ("FY2027 稀釋股數指引（高）", 847, "百萬股"),
                ("FY2027 RPO 指引（低）", 25.2, "US$B"), ("FY2027 RPO 指引（高）", 25.4, "US$B")]:
    add(m, v, u, "FY2027", "公司指引（2026-09-01）", G, "Interested-party", "公司指引，管理層有保守設定動機", "Financial Outlook")
add("FY2030 NGS ARR 目標", 20, "US$B", "FY2030", "管理層目標", G, "Interested-party", "管理層長期目標，用於資本市場溝通", "CEO／CFO 引言")
add("FY2028 調整後自由現金流率目標", 0.40, "比率", "FY2028", "管理層目標", G, "Interested-party", "同上", "CFO 引言")
add("FY2027Q1 營收指引（低）", 3300, "US$M", "FY2027Q1", "公司指引", G, "Interested-party", "公司指引", "Financial Outlook")
add("FY2027Q1 營收指引（高）", 3310, "US$M", "FY2027Q1", "公司指引", G, "Interested-party", "公司指引", "Financial Outlook")
add("NGS ARR 定義", "公司：年化分配營收（active contracts；含 product、subscription、support；排除硬體、legacy attached subscriptions、professional services）", "文字", "FY2026", "公司定義", "PXY", "Interested-party", "公司自訂 KPI 定義", "KPI 說明；10-K Item 7 Key Financial Metrics 同義")
# ---- AI 相關
add("Prisma AIRS（AI 安全）發布", "2026 年 4 月發布 Prisma AIRS 與 Unit 42 Frontier AI Defense（法定申報僅質性描述，未揭露營收）", "文字", "FY2026", "質性", "K26", "Interested-party", "公司自述", "Item 7 Overview")
add("Prisma AIRS ARR 逾 1 億美元（媒體轉述）", 100, "US$M", "GA 後四季內", "媒體轉述；8-K／10-K 未找到", "WS", "Interested-party", "媒體轉述公司說法，未在法定申報找到；誘因：公司推廣 AI 安全產品", "Bull case 段")
add("平台化客戶數（媒體轉述）", 1450, "家", "FY2026Q1", "媒體轉述 Q1 法說", "PH", "Interested-party", "作者轉述公司法說，未取原文", "Bull Arguments")
add("CyberArk 調查：機器身分對人類身分比", 82, "比（x:1）", "2025", "第三方調查（CyberArk，資安廠商）", "https://www.govinfosecurity.com/cyberark-rise-in-machine-identities-poses-new-risks-a-28967（經 unit_economics README S02 引用）", "Interested-party", "調查方為銷售身分安全產品的廠商（CyberArk 已併入 PANW）", "")
# ---- 市場資料與共識
add("現價（收盤）", 322.24, "US$", "2026-11-13", "收盤", "SAS", "Verified", "", "頁首：Nov 13, 2026 close；單一來源，待第二來源（Yahoo 頁面未標日期，值為 342.00，不採用）")
add("市值", 263.18, "US$B", "2026-11-13", "—", "SAS", "Verified", "", "")
add("Beta（5Y 月）：StockAnalysis", 0.97, "x", "2026-11-14", "原始值", "SAS", "Verified", "", "")
add("Beta（5Y 月）：Yahoo Finance", 0.93, "x", "頁面未標日期", "原始值", "YF", "Verified", "", "")
add("共識 FY2027 營收", 14170, "US$M", "2026-10-15", "共識平均（52 家）", "SA", "Interested-party", "賣方預估，受公司指引錨定", "Revenue and EPS 表")
add("共識 FY2027 Non-GAAP EPS", 4.18, "US$", "2026-10-15", "共識平均（52 家）", "SA", "Interested-party", "賣方預估", "Revenue and EPS 表")
add("共識 FY2028 營收／EPS", "未找到（免費頁鎖定）", "—", "FY2028", "—", "SA", "Interested-party", "—", "")
add("賣方平均目標價", 398.63, "US$", "2026-11-14", "55 家（S&P Global）", "SAS", "Interested-party", "賣方目標價；投行與公司有業務往來", "統計頁 analyst note")
add("賣方目標價（最低）", 190, "US$", "2026-10-23", "—", "MS", "Interested-party", "同上", "")
add("賣方目標價（最高）", 525, "US$", "2026-10-23", "—", "MS", "Interested-party", "同上", "")
add("賣方目標價（中位）", 415, "US$", "2026-10-15", "—", "SA", "Interested-party", "同上", "")
# ---- 同業（2026-11-13 收盤；StockAnalysis 統計頁）
for tk, px, mc, ev, evs, fpe, beta in [("FTNT", 127.22, 93.47, 89.23, 11.32, 33.85, 1.04), ("ZS", 192.47, 31.38, 29.78, 8.88, 39.22, 0.84), ("CHKP", 147.18, 15.03, 12.88, 4.67, 13.70, 0.42), ("NET", 305.81, 109.60, 109.18, 39.95, 174.39, 1.37)]:
    base = f"StockAnalysis.com {tk} 統計頁 https://stockanalysis.com/stocks/{tk.lower()}/statistics/"
    add(f"同業 {tk} 前瞻本益比", fpe, "x", "2026-11-13", "Forward P/E（非 GAAP EPS 口徑，分母為共識 NTM；資料商未標示精確期間）", base, "Interested-party", "共識 EPS 為賣方預估", "")
    add(f"同業 {tk} EV／營收（TTM）", evs, "x", "2026-11-13", "TTM", base, "Verified", "", "")
    add(f"同業 {tk} 收盤價", px, "US$", "2026-11-13", "收盤", base, "Verified", "", "")
    add(f"同業 {tk} Beta", beta, "x", "2026-11-13", "5Y", base, "Verified", "", "")
add("同業 PANW 前瞻本益比（資料商）", 77.10, "x", "2026-11-13", "Forward P/E", "SAS", "Interested-party", "共識 EPS 為賣方預估", "")
add("同業 PANW EV／營收（TTM，資料商）", 22.88, "x", "2026-11-13", "TTM", "SAS", "Verified", "", "")
# ---- 歷史相對倍數（G3 第 4 條；2026-11-17 擷取）
_MT = {"PANW": "https://www.macrotrends.net/stocks/charts/PANW/palo-alto-networks/price-sales", "ZS": "https://www.macrotrends.net/stocks/charts/ZS/zscaler/price-sales", "FTNT": "https://www.macrotrends.net/stocks/charts/FTNT/fortinet/price-sales", "CHKP": "https://www.macrotrends.net/stocks/charts/CHKP/check-point-software-technologies/price-sales"}
for _k, _u in _MT.items():
    ROWS.append([f"{_k} 市銷率（TTM）逐季序列 2020-04 至 2025-04", "21 點，見模型 RelHist 頁", "x", "2020-04-30 至 2025-04-30", "TTM 市銷率（價格 ÷ 每股銷售額）；PANW、ZS 取財季末，FTNT、CHKP 取前一個日曆季末", f"macrotrends.net {_k} price-sales 頁 {_u}", "2026-11-17", "Verified", ""])
ROWS.append(["PANW ÷ 同業中位市銷率：中位數", 0.924, "x", "2020-04 至 2025-04（21 點）", "同業＝ZS、FTNT、CHKP；Derived", "RelHist 頁（macrotrends 序列自算）", "2026-11-17", "Derived", ""])
ROWS.append(["PANW ÷ 同業中位市銷率：最低", 0.577, "x", "2020-04 至 2025-04", "同上", "RelHist 頁", "2026-11-17", "Derived", ""])
ROWS.append(["PANW ÷ 同業中位市銷率：最高", 1.794, "x", "2020-04 至 2025-04", "同上", "RelHist 頁", "2026-11-17", "Derived", ""])
ROWS.append(["PANW ÷ 同業中位 EV／營收（現行）", 2.577, "x", "2026-11-13", "同業＝ZS、FTNT、CHKP（3 家）；含 NET 為 2.265", "RelHist 頁（StockAnalysis 現行值自算）", "2026-11-17", "Derived", ""])
ROWS.append(["PANW 歷史前瞻本益比序列／NET 歷史市銷率", "未找到（共識 EPS 歷史不公開；macrotrends 禁止擷取 Cloudflare 頁）", "—", "—", "—", "—", "2026-11-17", "Verified", ""])
with open(os.path.join(D, "facts_paloalto.csv"), "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f); w.writerow(["指標", "數值", "單位", "期間", "口徑", "出處", "擷取日", "標記", "來源誘因"]); w.writerows(ROWS)
print(len(ROWS), "rows")

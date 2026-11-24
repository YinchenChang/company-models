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
# ---- 歷史相對倍數（G5 第 1 條；共用庫 software_ai/data/relmult_hist/out/paloalto.json，2026-11-17）
_J = "共用庫 software_ai/data/relmult_hist/out/paloalto.json（資料商 StockAnalysis 季度比率頁；同業 FTNT、ZS、CHKP）"
for _m, _v, _t in [("PS 讀數（歷史中位）", 1.1594302795370166, "Derived"), ("PSG 讀數", 1.4416760569670555, "Derived"), ("相對倍數基準 geo_base（兩讀數幾何平均）", 1.2928738816029726, "Derived"), ("Sens 低", 0.4982987722595606, "Derived"), ("Sens 高", 2.4330435896191496, "Derived"), ("成長比（本公司 ÷ 同業中位，TTM）", 1.3047030171601444, "Derived")]:
    ROWS.append([_m, _v, "x", "2021-09 至 2026-06（20 個財季）", "PANW ÷ 同業中位；PS 與 PSG", _J, "2026-11-17", _t, ""])
for _m, _v, _d in [("同業 P/S：Fortinet", 11.86, "2026-11-13"), ("同業 P/S：Zscaler", 9.36, "2026-11-13"), ("同業 P/S：Check Point", 5.44, "2026-11-13")]:
    ROWS.append([_m, _v, "x", _d, "PS＝市值÷TTM 營收；市值÷最新 TTM 營收，價格日 2026-11-13（無逾期家數）", "共用庫 software_ai/data/relmult_hist/out/peer_ps_20261113.json", "2026-11-17", "Verified", ""])
ROWS.append(["無風險利率（美國 10 年期）", 0.051, "比率", "2026-11-12", "FRED DGS10", "FRED https://fred.stlouisfed.org/series/DGS10（G5 第 2 條統一值）", "2026-11-17", "Verified", ""])
ROWS.append(["PANW 歷史前瞻本益比序列", "未找到（共識 EPS 歷史不公開；相對倍數改用共用庫 PS／PSG）", "—", "—", "—", "—", "2026-11-17", "Verified", ""])
# ======================================================================================
# v0.2.2 新增（2026-11-24）：FY2027 Q1（截至 2026-10-31）實績、新指引、10-Q 股數、現價與共識更新、擬收購 Applied Compute。
# 舊列保留原擷取日（2026-11-16／17）；新列擷取日 2026-11-24。找不到者寫「未找到」，不估。
ND = "2026-11-24"
U.update({
 "R27Q1": "PANW 8-K 附件 99.1 FY2027 Q1 財報新聞稿（2026-11-18）https://www.sec.gov/Archives/edgar/data/0001327567/000132756726000030/ex991q127earningsrelease.htm",
 "R27AC": "PANW 8-K 附件 99.2 Applied Compute 新聞稿（2026-11-18）https://www.sec.gov/Archives/edgar/data/0001327567/000132756726000030/ex992pressrelease_november.htm",
 "Q27Q1": "PANW 10-Q（季度截至 2026-10-31；簽署日 2026-11-18；取自公司 IR 檔案）https://investors.paloaltonetworks.com/static-files/6d8f675d-45ea-4c8e-98d5-89138518ae1d",
 "FOX": "Fox Business PANW 報價頁 https://www.foxbusiness.com/quote?stockTicker=PANW（標 2026-11-23 4:00 PM EDT 收盤；頁面另有一個 Close 欄為 343.32，與 Closed Market 的 353.43 不一致，採與前收盤＋漲跌額核對一致者）",
 "STW": "Stocktwits PANW 頁 https://stocktwits.com/symbol/panw（標示 Closed 353.43、漲 10.11、2.94%；頁面未標日期，以前收盤 343.32 核對為 2026-11-23）",
 "SAH": "StockAnalysis.com PANW 報價頁 https://stockanalysis.com/stocks/panw/ 與 YCharts https://ycharts.com/companies/PANW（兩者皆為 2026-11-20 收盤 343.32、前收盤 361.73）",
 "SA2": "StockAnalysis.com PANW 預測頁（S&P Global 資料；頁面標 2026-11-20 更新）https://stockanalysis.com/stocks/panw/forecast/",
 "CSO": "CSO Online 2026-11-19 https://csoonline.com/article/4250669/palo-alto-networks-bets-4-5b-on-enterprise-ai-with-applied-compute-acquisition.html",
 "BZ": "Benzinga KR PANW 圖表頁 https://kr.benzinga.com/quote/PANW/charts（標 2026-11-17 收盤 309.29；單一來源）",
})
def addn(metric, val, unit, period, basis, src, tag="Verified", inc="", where=""):
    s = U.get(src, src) + (f"；{where}" if where else "")
    ROWS.append([metric, val, unit, period, basis, s, ND, tag, inc])
Q1, Q1P = "FY2027Q1（2026-08-01 至 2026-10-31）", "FY2026Q1（2025-08-01 至 2025-10-31）"
# ---- 損益表（GAAP；8-K 附件 99.1 與 10-Q 相互核對：毛利率、分項毛利、營收三項一致）
for m, a, b in [("Product 營收", 595, 434), ("Subscription and support 營收", 2850, 2040), ("合併營收", 3445, 2474),
                ("營業成本：Product", 206, 89), ("營業成本：Subscription and support", 897, 549), ("總營業成本（Cost of revenue）", 1103, 638),
                ("GAAP 毛利", 2342, 1836), ("研發費用", 798, 528), ("銷售與行銷費用", 1167, 820), ("一般與行政費用", 281, 179), ("營業費用合計", 2246, 1527),
                ("GAAP 營業利益", 96, 309), ("其他收入（費用）淨額", -71, 103), ("稅前利益", 25, 412), ("所得稅（費用；負值＝利益）", -457, 78), ("GAAP 淨利", 482, 334)]:
    addn(m, a, "US$M", Q1, "GAAP", "R27Q1", where="損益表（與 10-Q 一致）")
    addn(m, b, "US$M", Q1P, "GAAP", "R27Q1", where="損益表上年同期欄")
for m, a, b in [("GAAP 稀釋 EPS", 0.57, 0.47), ("GAAP 基本 EPS", 0.59, 0.49)]:
    addn(m, a, "US$", Q1, "GAAP", "R27Q1", where="損益表"); addn(m, b, "US$", Q1P, "GAAP", "R27Q1", where="損益表上年同期欄")
addn("GAAP 稀釋加權股數", 838, "百萬股", Q1, "GAAP", "R27Q1", where="損益表；10-Q Note 16：基本 819＋可轉債 2＋員工股權計畫 17")
addn("GAAP 稀釋加權股數", 709, "百萬股", Q1P, "GAAP", "R27Q1", where="損益表上年同期欄")
addn("所得稅利益主因", 369, "US$M", Q1, "GAAP", "Q27Q1", where="MD&A：FY2018 過渡年稅務立場達 more-likely-than-not 門檻的單次美國聯邦利益；有效稅率 −1,828%（上年同期 18.9%）；一次性")
addn("GAAP 毛利率（合計）", 0.680, "比率", Q1, "GAAP", "Q27Q1", where="MD&A（上年同期 74.2%）")
addn("GAAP 毛利率（Product）", 0.654, "比率", Q1, "GAAP", "Q27Q1", where="MD&A（上年同期 79.5%）；硬體供應鏈成本、產品組合、無形資產攤銷")
addn("GAAP 毛利率（Subscription and support）", 0.685, "比率", Q1, "GAAP", "Q27Q1", where="MD&A（上年同期 73.1%）；無形資產攤銷、雲端服務成本")
for m, a, b in [("Americas 營收", 2325, 1641), ("EMEA 營收", 728, 521), ("APAC 營收", 392, 312)]:
    addn(m, a, "US$M", Q1, "GAAP", "Q27Q1", where="地區表"); addn(m, b, "US$M", Q1P, "GAAP", "Q27Q1", where="地區表上年同期欄")
addn("Subscription 與 Support 分項營收", "未找到（10-Q 與新聞稿 Q1 皆只列 Subscription and support 合計）", "—", Q1, "—", "Q27Q1", "Verified", "", "營收表")
# ---- Non-GAAP 調節（公司自訂）
INCO = "公司自訂 non-GAAP 調節，排除股票薪酬、併購攤銷等"
for m, a, b in [("Non-GAAP 營業利益", 1059, 746), ("SBC-related charges", 622, 387), ("併購相關費用", 55, 5), ("收購無形資產攤銷", 282, 39), ("訴訟相關費用", 4, 6), ("Non-GAAP 淨利", 881, 662)]:
    addn(m, a, "US$M", Q1, "Non-GAAP（公司定義）", "R27Q1", "Interested-party", INCO, "GAAP 至 non-GAAP 調節表")
    addn(m, b, "US$M", Q1P, "Non-GAAP（公司定義）", "R27Q1", "Interested-party", INCO, "調節表上年同期欄")
addn("Non-GAAP 稀釋 EPS", 1.05, "US$", Q1, "Non-GAAP（公司定義）", "R27Q1", "Interested-party", INCO, "調節表"); addn("Non-GAAP 稀釋 EPS", 0.93, "US$", Q1P, "Non-GAAP（公司定義）", "R27Q1", "Interested-party", INCO, "調節表上年同期欄")
addn("Non-GAAP 稀釋股數", 837, "百萬股", Q1, "Non-GAAP（GAAP 838 減票據避險與 capped call 反稀釋 1）", "R27Q1", "Interested-party", INCO, "調節表")
addn("股票薪酬費用（10-Q 費用表）", 606, "US$M", Q1, "GAAP", "Q27Q1", where="成本與費用表（上年同期 370）；現金流量表 596、權益變動表 593，三表差異未釐清")
addn("尚未認列的 SBC 費用", 4500, "US$M", "2026-10-31", "GAAP", "Q27Q1", where="加權平均約 2.4 年認列（約數）")
# ---- 現金流與資產負債（2026-10-31）
for m, a, b in [("營業活動現金流", 1929, 1771), ("資本支出（PP&E 與其他資產）", 87, 84), ("自由現金流（公司定義）", 1842, 1687), ("調整後自由現金流（公司定義）", 1937, 1713)]:
    tg, inc = (("Interested-party", "公司定義，加回併購付款等") if "調整後" in m else (("Interested-party", "公司定義的 FCF，SBC 不扣") if "自由" in m else ("Verified", "")))
    addn(m, a, "US$M", Q1, "Non-GAAP" if tg == "Interested-party" else "GAAP", "R27Q1", tg, inc, "現金流量表與 FCF 調節表")
    addn(m, b, "US$M", Q1P, "Non-GAAP" if tg == "Interested-party" else "GAAP", "R27Q1", tg, inc, "上年同期欄")
addn("調整後自由現金流率", 0.562, "比率", Q1, "Non-GAAP", "R27Q1", "Interested-party", "公司定義", "FCF 調節表（上年同期 69.2%；Q1 為季節性高點）")
addn("投資活動現金流", -1900, "US$M", Q1, "GAAP", "R27Q1", where="現金流量表（上年同期 −983）")
for m, v in {"現金及約當現金": 2594, "短期投資": 461, "長期投資": 6080, "短期可轉換票據（公允價值；自長期重分類）": 1930, "遞延收入（流動）": 7730, "遞延收入（長期）": 6859, "商譽": 22559, "無形資產淨額": 6869, "資產總計": 49110}.items():
    addn(m, v, "US$M", "2026-10-31", "GAAP", "R27Q1", where="資產負債表（現金與投資與票據亦見 10-Q）")
addn("現金、短期與長期投資合計", 9135, "US$M", "2026-10-31", "GAAP", "Q27Q1", where="MD&A Liquidity（2026-07-31 為 7,906）；營運資金赤字 −3,906（7 月底 −1,280）")
addn("2030 可轉債剩餘本金", 1100, "US$M", "2026-10-31", "面額", "Q27Q1", where="CyberArk 原發行 1,250；0%，2030-06-15 到期；有效換股價約 211.24；初始交換條件：每 1,000 本金約 4.3161 股加 88.2630 現金；2026 年第四曆季可轉換（價格條件已達）")
addn("Capped call 公允價值", 169, "US$M", "2026-10-31", "公允價值", "Q27Q1", where="履約價約 211.24；上限價約 287.21 至 291.44")
addn("流通股數（10-Q 封面）", 821, "百萬股", "2026-11-11", "封面", "Q27Q1", where="封面：截至 2026-11-11")
addn("未歸屬 RSU", 16, "百萬股", "2026-10-31", "全額", "Q27Q1", where="Note 12（加權授予日公允價值 195.13）")
addn("未歸屬 PSU", 10, "百萬股", "2026-10-31", "績效型（最大可得股數）", "Q27Q1", where="Note 12（加權授予日公允價值 235.87）")
addn("買回授權剩餘", 1000, "US$M", "2026-10-31", "—", "Q27Q1", where="總授權 5,100；期限 2027-12-31")
addn("員工數", 22339, "人", "2026-10-31", "期末", "Q27Q1", where="MD&A（FY2026 年底 21,921）")
addn("兩家經銷商占營收", 0.25, "比率", Q1, "各占 10% 以上合計", "Q27Q1", where="客戶集中度（單一經銷商占比未揭露）")
addn("Embrace Mobile 購買價格分攤", "對價 292（現金 287）；商譽 190；無形資產 104", "US$M", "2026-08-27 完成", "購買價格", "Q27Q1", where="Note 7（FY2026 10-K 估 325，完成後 10-Q 為 292）")
addn("Console Systems 購買價格分攤", "對價 393（現金 347）；商譽 354；無形資產 31", "US$M", "2026-09-01 完成", "購買價格", "Q27Q1", where="Note 7（FY2026 10-K 估 500，完成後 10-Q 為 393）")
addn("Chronosphere、CyberArk 本季貢獻營收與營業損失", "未找到（10-Q 稱併購影響不重大；未列金額）", "—", Q1, "—", "Q27Q1")
# ---- KPI（公司自訂）
addn("NGS ARR（期末）", 9.73, "US$B", "FY2027Q1", "公司 KPI（含 CyberArk、Chronosphere；10-Q 取整為 9.7）", "R27Q1", "Interested-party", "公司自揭 KPI，無第三方查核；定義可變", "Financial Highlights")
addn("NGS ARR 年增率（期末）", 0.66, "比率", "FY2027Q1", "公司 KPI（含併購）", "R27Q1", "Interested-party", "公司自揭 KPI；含併購", "Financial Highlights")
addn("RPO（期末）", 21.0, "US$B", "FY2027Q1", "ASC 606 剩餘履約義務", "Q27Q1", "Verified", "", "Note 2；下 12 個月預期認列約 9.4（FY2026 年底 9.3）；8-K 年增 35%")
addn("本季淨新增 NGS ARR", 0.63, "US$B", "FY2027Q1", "9.73 − 9.1（相鄰兩季相減）", "R27Q1 與 R26Q4", "Derived", "", "含併購（Embrace 2026-08-27、Console 2026-09-01 完成）併入，有機部分公司未拆分")
addn("NGS ARR 有機成長", "未找到（公司未拆分；新聞稿未說明是否含併購 ARR）", "—", "FY2027Q1", "—", "R27Q1")
# ---- 推導：TTM（至 2026-10-31）與 Q1 對 Q1 指引
TTMB = "TTM＝FY2026 − FY2026Q1 ＋ FY2027Q1（Derived）"
for m, v in [("Product 營收", 2441), ("Subscription and support 營收", 10010), ("合併營收", 12451), ("GAAP 營業利益", 482), ("Non-GAAP 營業利益", 3669), ("SBC-related charges", 1947), ("收購無形資產攤銷", 881), ("營業活動現金流", 4711), ("資本支出（PP&E 與其他資產）", 443), ("自由現金流（公司定義）", 4268)]:
    addn(m, v, "US$M", "TTM 至 2026-10-31", TTMB, "由 R26Q4 與 R27Q1 相加減", "Derived")
for m, v, b in [("Non-GAAP 營業利益率", 3669 / 12451, "3,669 ÷ 12,451"), ("SBC-related charges ÷ 營收", 1947 / 12451, "1,947 ÷ 12,451")]:
    addn(m, round(v, 4), "比率", "TTM 至 2026-10-31", b, "由上列相除", "Derived")
for m, v, b in [("Non-GAAP 營業利益率", 1059 / 3445, "1,059 ÷ 3,445"), ("SBC-related charges ÷ 營收", 622 / 3445, "622 ÷ 3,445"), ("營收年增率", 3445 / 2474 - 1, "3,445 ÷ 2,474 − 1")]:
    addn(m, round(v, 4), "比率", "FY2027Q1", b, "由上列相除", "Derived")
addn("SBC-related charges ÷ 營收", round(487 / 3410, 4), "比率", "FY2026Q4", "487 ÷ 3,410（上一季）", "由本檔 FY2026Q4 列相除", "Derived")
addn("Q1 實績對 Q1 營收指引上緣（3,310）", round(3445 / 3310 - 1, 4), "比率", "FY2027Q1", "3,445 ÷ 3,310 − 1", "R27Q1 與 R26Q4", "Derived")
addn("Q1 實績對 Q1 NGS ARR 指引上緣（9.56）", round(9.73 / 9.56 - 1, 4), "比率", "FY2027Q1", "9.73 ÷ 9.56 − 1", "R27Q1 與 R26Q4", "Derived")
# ---- 舊 Q1 指引（2026-09-01；補列，供與實績對照）
for m, v, u in [("FY2027Q1 NGS ARR 指引（低）", 9.54, "US$B"), ("FY2027Q1 NGS ARR 指引（高）", 9.56, "US$B"), ("FY2027Q1 RPO 指引（低）", 20.8, "US$B"), ("FY2027Q1 RPO 指引（高）", 20.9, "US$B"),
                ("FY2027Q1 Non-GAAP 稀釋 EPS 指引（低）", 0.96, "US$"), ("FY2027Q1 Non-GAAP 稀釋 EPS 指引（高）", 0.98, "US$")]:
    addn(m, v, u, "FY2027Q1", "公司指引（2026-09-01）", "R26Q4", "Interested-party", "公司指引，管理層有保守設定動機", "Financial Outlook")
# ---- 新指引（2026-11-18）
for m, v, u in [("FY2027 營收指引（低）", 14520, "US$M"), ("FY2027 營收指引（高）", 14570, "US$M"), ("FY2027 NGS ARR 指引（低）", 11.475, "US$B"), ("FY2027 NGS ARR 指引（高）", 11.525, "US$B"),
                ("FY2027 RPO 指引（低）", 25.3, "US$B"), ("FY2027 RPO 指引（高）", 25.5, "US$B"), ("FY2027 Non-GAAP 營業利益率指引", 0.295, "比率"),
                ("FY2027 Non-GAAP 稀釋 EPS 指引（低）", 4.26, "US$"), ("FY2027 Non-GAAP 稀釋 EPS 指引（高）", 4.30, "US$"), ("FY2027 稀釋股數指引（低）", 844, "百萬股"), ("FY2027 稀釋股數指引（高）", 847, "百萬股"),
                ("FY2027 調整後自由現金流率指引", 0.38, "比率"),
                ("FY2027Q2 營收指引（低）", 3530, "US$M"), ("FY2027Q2 營收指引（高）", 3550, "US$M"), ("FY2027Q2 NGS ARR 指引（低）", 10.20, "US$B"), ("FY2027Q2 NGS ARR 指引（高）", 10.22, "US$B"),
                ("FY2027Q2 RPO 指引（低）", 21.5, "US$B"), ("FY2027Q2 RPO 指引（高）", 21.6, "US$B"), ("FY2027Q2 Non-GAAP 稀釋 EPS 指引（低）", 1.01, "US$"), ("FY2027Q2 Non-GAAP 稀釋 EPS 指引（高）", 1.03, "US$"),
                ("FY2027Q2 稀釋股數指引（低）", 837, "百萬股"), ("FY2027Q2 稀釋股數指引（高）", 842, "百萬股")]:
    per = "FY2027Q2" if "Q2" in m else "FY2027"
    addn(m, v, u, per, "公司指引（2026-11-18；取代 2026-09-01 指引）", "R27Q1", "Interested-party", "公司指引，管理層有保守設定動機", "Financial Outlook")
addn("FY2027 營收指引中點變動（新舊）", 14545 - 14150, "US$M", "FY2027", "(14,520＋14,570)÷2 − (14,100＋14,200)÷2", "R27Q1 與 R26Q4", "Derived", "", "")
addn("FY2027 營收指引上調幅度扣除 Q1 營收超出 Q1 指引中點後之餘額", (14545 - 14150) - (3445 - 3305), "US$M", "FY2027", "395 − (3,445 − 3,305)", "R27Q1 與 R26Q4", "Derived", "", "餘額＝管理層對 Q2–Q4 營收的上調（Interested-party 指引的推導，不入模型）")
addn("FY2030 NGS ARR 目標、FY2028 調整後自由現金流率目標是否更新", "未找到（Q1 新聞稿摘要與媒體報導皆未提；沿用 2026-09-01 的 20 與 40%）", "—", "FY2027Q1", "—", "R27Q1")
# ---- 擬收購 Applied Compute（未交割；不入模型）
addn("擬收購 Applied Compute：對價", 4500, "US$M", "2026-11-18 簽約", "全股票（以 PANW 普通股支付），subject to adjustments；無現金成分、無留任或或有對價載明", "R27AC", "Verified", "", "新聞稿原文：The all-stock transaction is valued at $4.5 billion, subject to adjustments；10-Q Note 19 同；CSO Online 2026-11-19 為第三來源")
addn("擬收購 Applied Compute：預期交割", "FY2027 第三季（2027-02 至 2027-04），含主管機關核准等條件；可能不成", "文字", "2026-11-18", "—", "Q27Q1", "Verified", "", "Note 19 Subsequent Events")
addn("擬收購 Applied Compute：標的營收、ARR、虧損", "未找到（未揭露）", "—", "—", "—", "R27AC")
addn("擬收購 Applied Compute：是否已含於 FY2027 指引", "未找到（新聞稿與 10-Q 未說明；Q1 新聞稿指引段以 intent to acquire 描述）", "—", "FY2027", "—", "R27Q1")
addn("擬收購 Applied Compute：約當發行股數", round(4500 / 353.43, 1), "百萬股", "估算（以 2026-11-23 收盤 353.43 計）", "4,500 ÷ 353.43；實際股數依交割時價格與調整條款而定", "由對價與現價相除", "Derived", "", "占 10-Q 封面流通股 821 的約 1.6%；交割時價格未知，只作量級")
# ---- 現價與共識（財報後）
addn("現價（收盤）", 353.43, "US$", "2026-11-23", "收盤（財報後最近一個交易日）", "FOX", "Verified", "", "第二來源 Stocktwits 同值；前收盤 343.32（2026-11-20）加漲跌額 10.11（2.94%）＝353.43")
addn("現價（收盤）第二來源", 353.43, "US$", "2026-11-23", "收盤", "STW", "Verified", "", "頁面未標日期；與第一來源與前收盤算術一致")
addn("前一交易日收盤", 343.32, "US$", "2026-11-20", "收盤", "SAH", "Verified", "", "兩來源一致；前收盤 361.73（2026-11-19）")
addn("財報前收盤", 309.29, "US$", "2026-11-17", "收盤（財報公布前最後一個交易日）", "BZ", "Verified", "", "單一來源")
addn("市值（資料商）", 281.87, "US$B", "2026-11-20", "以 2026-11-20 收盤 343.32 與 821 百萬股", "SAH", "Verified", "", "")
addn("賣方平均目標價", 409.72, "US$", "2026-11-20", "54 家（S&P Global）", "SA2", "Interested-party", "賣方目標價；投行與公司有業務往來", "預測頁；中位 425、最高 510、最低 190")
addn("賣方目標價（中位）", 425, "US$", "2026-11-20", "—", "SA2", "Interested-party", "同上", "")
addn("賣方目標價（最高）", 510, "US$", "2026-11-20", "—", "SA2", "Interested-party", "同上", "")
addn("賣方目標價（最低）", 190, "US$", "2026-11-20", "—", "SA2", "Interested-party", "同上", "")
addn("共識 FY2027 營收", 14600, "US$M", "2026-11-20", "共識平均（44 家；頁面只給到 14.6 十億，四捨五入）", "SA2", "Interested-party", "賣方預估，受公司指引錨定", "預測頁；高 15.0、低 14.5 十億")
addn("共識 FY2027 Non-GAAP EPS", 4.28, "US$", "2026-11-20", "共識平均（44 家）", "SA2", "Interested-party", "賣方預估", "預測頁；高 4.41、低 4.22")
addn("共識 FY2028 營收／EPS", "未找到（免費頁鎖定）", "—", "FY2028", "—", "SA2", "Interested-party", "—", "")
# ---- 逾期標示（G10-3、第 15 條）
addn("同業 P/S 價格日與現價日相差", 6, "交易日", "2026-11-13 對 2026-11-23", "週一至週五計，無休市日", "共用庫 peer_ps_20261113.json", "Derived", "", "超過 2 個交易日：標逾期；本版仍讀 peer_ps_20261113.json（工作單 G10-3）")
addn("無風險利率日與現價日相差", 7, "交易日", "2026-11-12 對 2026-11-23", "週一至週五計，無休市日", "FRED DGS10（G5 統一值）", "Derived", "", "超過 2 個交易日：標逾期；rf 不動（工作單 G10-3）")

with open(os.path.join(D, "facts_paloalto.csv"), "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f); w.writerow(["指標", "數值", "單位", "期間", "口徑", "出處", "擷取日", "標記", "來源誘因"]); w.writerows(ROWS)
print(len(ROWS), "rows")

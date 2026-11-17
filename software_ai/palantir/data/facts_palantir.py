# Palantir（PLTR）事實總帳：每個數字一列（工作單「共同規則」第 4 條）。本檔是唯一來源：
#   - 產生 data/facts_palantir.csv（研究包）
#   - build_palantir.py 的 Q／Hist／In 頁輸入由本檔取值（import FACTS）
# 單位：金額除另註外為 US$ 百萬；股數為百萬股。擷取日 2026-11-16。
# 標記：Verified／Interested-party／Analogy／Assumed／Derived（五選一）。
import csv, os

D = "2026-11-16"
U_Q3PR = "https://www.sec.gov/Archives/edgar/data/0001321655/000132165526000045/a2026q3ex991pressrelease.htm｜8-K 附件 99.1（2026Q3 財報新聞稿）"
U_Q3Q = "https://www.sec.gov/Archives/edgar/data/0001321655/000132165526000048/pltr-20260930.htm｜10-Q（2026-09-30）"
U_Q2Q = "https://www.sec.gov/Archives/edgar/data/0001321655/000132165526000041/pltr-20260630.htm｜10-Q（2026-06-30）"
U_K25 = "https://investors.palantir.com/files/2025%20FY%20PLTR%2010-K.pdf｜10-K FY2025"
U_K22 = "https://www.sec.gov/Archives/edgar/data/1321655/000132165523000011/pltr-20221231.htm｜10-K FY2022"
U_K23 = "https://www.sec.gov/Archives/edgar/data/1321655/000132165524000022/pltr-20231231.htm｜10-K FY2023"
U_Q1PR = "https://www.businesswire.com/news/home/20260503338048/en/Palantir-Reports-Q1-2026-U.S.-Revenue-Growth-of-104-YY-and-Revenue-Grow｜2026Q1 財報新聞稿"
U_Q2MB = "https://www.marketbeat.com/instant-alerts/palantir-technologies-q2-earnings-call-highlights-2026-08-03/｜MarketBeat 2026Q2 法說摘要（第二來源）"
U_Q4RA = "https://rallies.ai/news/palantir-posts-137-q4-commercial-revenue-surge-guides-718b720b-for-2026｜rallies.ai 2025Q4 摘要（第二來源）"
U_FUT = "https://futurumgroup.com/?p=97554｜Futurum 2026Q3 分析（第二來源）"
U_SA = "https://stockanalysis.com/stocks/pltr/history/｜StockAnalysis PLTR 歷史股價"

FY = "會計年度（日曆年，1/1–12/31）"
GAAP = "GAAP"
NG = "non-GAAP（公司定義）"

FACTS = []   # (id, 指標, 數值, 單位, 期間, 口徑, 出處, 標記, 來源誘因)
def f(i, name, v, unit, per, basis, src, tag, inc=""):
    FACTS.append(dict(id=i, 指標=name, 數值=v, 單位=unit, 期間=per, 口徑=basis, 出處=src, 擷取日=D, 標記=tag, 來源誘因=inc))

# ---- 分部營收：年度 FY2021–FY2025 ----
GOV = {2021: 897.356, 2022: 1071.776, 2023: 1222.215, 2024: 1569.605, 2025: 2402.287}
COM = {2021: 644.533, 2022: 834.095, 2023: 1002.797, 2024: 1295.902, 2025: 2073.159}
REV = {2021: 1541.889, 2022: 1905.871, 2023: 2225.012, 2024: 2865.507, 2025: 4475.446}
SRC_SEG = {2021: U_K22 + "；MD&A Revenue 表（2022 與 2021 欄）", 2022: U_K22 + "；MD&A Revenue 表", 2023: U_K25 + "；Note 分部（FY2023 欄）", 2024: U_K25 + "；Note 分部（FY2024 欄）；MD&A p.75", 2025: U_K25 + "；Note 分部（FY2025 欄）；MD&A p.75"}
for y in range(2021, 2026):
    f(f"seg.gov.{y}", "Government 分部營收", GOV[y], "US$M", f"FY{y}", GAAP + "；分部營收", SRC_SEG[y], "Verified")
    f(f"seg.com.{y}", "Commercial 分部營收", COM[y], "US$M", f"FY{y}", GAAP + "；分部營收", SRC_SEG[y], "Verified")
    f(f"rev.{y}", "合併營收", REV[y], "US$M", f"FY{y}", GAAP, (U_K22 if y <= 2022 else U_K25) + "；損益表（FY2023–2025：p.86；FY2021–2022：MD&A）", "Verified")
# ---- 分部營收：最近 4 季 ----
QGOV = {"2025Q4": 729.665, "2026Q1": 858.410, "2026Q2": 990.032, "2026Q3": 1145.540}
QCOM = {"2025Q4": 677.137, "2026Q1": 774.173, "2026Q2": 945.432, "2026Q3": 1167.871}
QREV = {"2025Q4": 1406.802, "2026Q1": 1632.583, "2026Q2": 1935.464, "2026Q3": 2313.411}
f("seg.gov.2026Q3", "Government 分部營收", 1145.540, "US$M", "2026Q3（7–9 月）", GAAP, U_Q3Q + "；Note 12 分部表", "Verified")
f("seg.com.2026Q3", "Commercial 分部營收", 1167.871, "US$M", "2026Q3", GAAP, U_Q3Q + "；Note 12", "Verified")
f("rev.2026Q3", "合併營收", 2313.411, "US$M", "2026Q3", GAAP, U_Q3Q + "；損益表 p.4", "Verified")
f("seg.gov.2026Q2", "Government 分部營收", 990.032, "US$M", "2026Q2", GAAP, U_Q2Q + "；Note 12", "Verified")
f("seg.com.2026Q2", "Commercial 分部營收", 945.432, "US$M", "2026Q2", GAAP, U_Q2Q + "；Note 12", "Verified")
f("rev.2026Q2", "合併營收", 1935.464, "US$M", "2026Q2", GAAP, U_Q2Q + "；Note 12 合計", "Verified")
f("seg.gov.2026Q1", "Government 分部營收", 858.410, "US$M", "2026Q1", GAAP + "；＝6M 1,848.442 − Q2 990.032", U_Q2Q + "；Note 12（6M 欄減 Q2）", "Derived")
f("seg.com.2026Q1", "Commercial 分部營收", 774.173, "US$M", "2026Q1", GAAP + "；＝6M 1,719.605 − Q2 945.432", U_Q2Q + "；Note 12", "Derived")
f("rev.2026Q1", "合併營收", 1632.583, "US$M", "2026Q1", GAAP, U_Q1PR + "；Revenue 1,632,583 千美元", "Verified")
f("seg.gov.2025Q4", "Government 分部營收", 729.665, "US$M", "2025Q4", GAAP + "；＝FY2025 2,402.287 − 9M25 1,672.622", U_K25 + "與" + U_Q3Q + "（9M 2025 欄）", "Derived")
f("seg.com.2025Q4", "Commercial 分部營收", 677.137, "US$M", "2025Q4", GAAP + "；＝FY2025 2,073.159 − 9M25 1,396.022", U_K25 + "與" + U_Q3Q, "Derived")
f("rev.2025Q4", "合併營收", 1406.802, "US$M", "2025Q4", GAAP + "；＝FY2025 4,475.446 − 9M25 3,068.644", U_K25 + "與" + U_Q3Q, "Derived")
f("seg.gov.9M26", "Government 分部營收", 2993.982, "US$M", "2026 前三季", GAAP, U_Q3Q + "；Note 12", "Verified")
f("seg.com.9M26", "Commercial 分部營收", 2887.476, "US$M", "2026 前三季", GAAP, U_Q3Q + "；Note 12", "Verified")
f("rev.9M26", "合併營收", 5881.458, "US$M", "2026 前三季", GAAP, U_Q3Q + "；損益表 p.4", "Verified")
f("seg.gov.9M25", "Government 分部營收", 1672.622, "US$M", "2025 前三季", GAAP, U_Q3Q + "；Note 12", "Verified")
f("seg.com.9M25", "Commercial 分部營收", 1396.022, "US$M", "2025 前三季", GAAP, U_Q3Q + "；Note 12", "Verified")
f("rev.9M25", "合併營收", 3068.644, "US$M", "2025 前三季", GAAP, U_Q3Q + "；損益表 p.4", "Verified")
f("seg.gov.Q3_25", "Government 分部營收", 632.676, "US$M", "2025Q3", GAAP, U_Q3Q + "；Note 12", "Verified")
f("seg.com.Q3_25", "Commercial 分部營收", 548.416, "US$M", "2025Q3", GAAP, U_Q3Q + "；Note 12", "Verified")
# ---- 分部 contribution（公司定義的分部利益）----
f("ctb.gov.2025", "Government contribution", 1576.070, "US$M", "FY2025", "公司定義分部利益（非 GAAP 營業利益）", U_K25 + "；分部 Note", "Verified")
f("ctb.com.2025", "Commercial contribution", 1366.627, "US$M", "FY2025", "同上", U_K25 + "；分部 Note", "Verified")
f("ctb.gov.2024", "Government contribution", 948.440, "US$M", "FY2024", "同上", U_K25 + "；分部 Note", "Verified")
f("ctb.com.2024", "Commercial contribution", 771.508, "US$M", "FY2024", "同上", U_K25 + "；分部 Note", "Verified")
f("ctb.gov.2023", "Government contribution", 724.970, "US$M", "FY2023", "同上", U_K25 + "；分部 Note", "Verified")
f("ctb.com.2023", "Commercial contribution", 520.585, "US$M", "FY2023", "同上", U_K25 + "；分部 Note", "Verified")
f("ctb.gov.9M26", "Government contribution", 2127.775, "US$M", "2026 前三季", "同上", U_Q3Q + "；Note 12", "Verified")
f("ctb.com.9M26", "Commercial contribution", 2195.425, "US$M", "2026 前三季", "同上", U_Q3Q + "；Note 12", "Verified")
f("ctb.gov.9M25", "Government contribution", 1069.081, "US$M", "2025 前三季", "同上", U_Q3Q + "；Note 12", "Verified")
f("ctb.com.9M25", "Commercial contribution", 889.098, "US$M", "2025 前三季", "同上", U_Q3Q + "；Note 12", "Verified")
f("ctb.gov.2026Q3", "Government contribution", 797.258, "US$M", "2026Q3", "同上", U_Q3Q + "；Note 12", "Verified")
f("ctb.com.2026Q3", "Commercial contribution", 897.525, "US$M", "2026Q3", "同上", U_Q3Q + "；Note 12", "Verified")
f("ctb.gov.2026Q2", "Government contribution", 701.106, "US$M", "2026Q2", "同上", U_Q2Q + "；Note 12", "Verified")
f("ctb.com.2026Q2", "Commercial contribution", 734.944, "US$M", "2026Q2", "同上", U_Q2Q + "；Note 12", "Verified")
f("ctb.gov.2026Q1", "Government contribution", 629.411, "US$M", "2026Q1", "同上；＝6M 1,330.517 − Q2 701.106", U_Q2Q + "；Note 12", "Derived")
f("ctb.com.2026Q1", "Commercial contribution", 562.956, "US$M", "2026Q1", "同上；＝6M 1,297.900 − Q2 734.944", U_Q2Q + "；Note 12", "Derived")
f("ctb.gov.2025Q4", "Government contribution", 506.989, "US$M", "2025Q4", "同上；＝FY2025 − 9M25", U_K25 + "與" + U_Q3Q, "Derived")
f("ctb.com.2025Q4", "Commercial contribution", 477.529, "US$M", "2025Q4", "同上；＝FY2025 − 9M25", U_K25 + "與" + U_Q3Q, "Derived")
# ---- 地理與美國商業／政府 ----
f("us.com.2026Q3", "美國商業營收", 984, "US$M", "2026Q3", "公司揭露（四捨五入到百萬）", U_Q3PR + "；Highlights", "Interested-party", "公司自揭 KPI；管理層薪酬與股價連動")
f("us.gov.2026Q3", "美國政府營收", 953, "US$M", "2026Q3", "同上", U_Q3PR + "；Highlights", "Interested-party", "公司自揭 KPI")
f("us.com.2026Q2", "美國商業營收", 764, "US$M", "2026Q2", "同上", U_Q2MB, "Interested-party", "公司自揭 KPI（經媒體轉述，第二來源）")
f("us.gov.2026Q2", "美國政府營收", 809, "US$M", "2026Q2", "同上", U_Q2MB, "Interested-party", "同上")
f("us.com.2026Q1", "美國商業營收", 595, "US$M", "2026Q1", "同上", U_Q1PR, "Interested-party", "公司自揭 KPI")
f("us.gov.2026Q1", "美國政府營收", 687, "US$M", "2026Q1", "同上", U_Q1PR, "Interested-party", "公司自揭 KPI")
f("us.com.2025Q4", "美國商業營收", 507, "US$M", "2025Q4", "同上", U_Q4RA, "Interested-party", "公司自揭 KPI（經媒體轉述，第二來源）")
f("us.gov.2025Q4", "美國政府營收", 570, "US$M", "2025Q4", "同上", U_Q4RA, "Interested-party", "同上")
f("ic.2026Q3", "國際商業營收（＝Commercial 1,167.871 − 美國商業 984）", 183.871, "US$M", "2026Q3", "Derived", U_Q3Q + "與" + U_Q3PR, "Derived")
f("us.rev.9M26", "美國營收", 4792.302, "US$M", "2026 前三季", GAAP + "；Note 12 地理", U_Q3Q + "；Note 12", "Verified")
f("row.rev.9M26", "其他地區營收", 1089.156, "US$M", "2026 前三季", GAAP, U_Q3Q + "；Note 12", "Verified")
f("us.rev.2025", "美國營收", 3320.043, "US$M", "FY2025", GAAP, U_K25 + "；Note 地理", "Verified")
f("row.rev.2025", "其他地區營收（合計 3,320.043 + 1,155.403 ＝ 4,475.446；Note 列出 UK 427.398 另計）", 1155.403, "US$M", "FY2025", "Derived（合併 − 美國）", U_K25, "Derived")
# ---- 損益（GAAP）FY2023–FY2025 ----
IS = {
 "cogs": ("營業成本（Cost of revenue）", {2023: 431.105, 2024: 565.990, 2025: 789.177}),
 "gp": ("毛利", {2023: 1793.907, 2024: 2299.517, 2025: 3686.269}),
 "sm": ("銷售與行銷", {2023: 744.992, 2024: 887.755, 2025: 1056.859}),
 "rd": ("研發", {2023: 404.624, 2024: 507.878, 2025: 557.677}),
 "ga": ("一般與管理", {2023: 524.325, 2024: 593.481, 2025: 657.718}),
 "oi": ("營業利益（GAAP）", {2023: 119.966, 2024: 310.403, 2025: 1414.015}),
 "int": ("利息收入", {2023: 132.572, 2024: 196.792, 2025: 229.181}),
 "oth": ("其他收入（費用）淨額", {2023: -15.447, 2024: -18.022, 2025: 14.172}),
 "pbt": ("稅前利益", {2023: 237.091, 2024: 489.173, 2025: 1657.368}),
 "tax": ("所得稅費用", {2023: 19.716, 2024: 21.255, 2025: 22.724}),
 "ni": ("淨利", {2023: 217.375, 2024: 467.918, 2025: 1634.644}),
 "dsh": ("稀釋加權平均股數", {2023: 2297.927, 2024: 2450.818, 2025: 2565.197}),
 "sbc": ("股票薪酬（現金流量表）", {2023: 475.903, 2024: 691.638, 2025: 684.033}),
 "cfo": ("營業現金流", {2023: 712.183, 2024: 1153.865, 2025: 2134.473}),
 "capex": ("資本支出", {2023: 15.114, 2024: 12.634, 2025: 33.882}),
 "da": ("折舊與攤銷", {2023: 33.354, 2024: 31.587, 2025: 26.145}),
}
for k, (nm, d) in IS.items():
    for y, v in d.items():
        unit = "百萬股" if k == "dsh" else "US$M"
        f(f"is.{k}.{y}", nm, v, unit, f"FY{y}", GAAP, U_K25 + "；損益表 p.86／現金流量表 p.89", "Verified")
for y, v in {2023: 0.09, 2024: 0.19, 2025: 0.63}.items():
    f(f"is.eps.{y}", "稀釋 EPS（GAAP）", v, "US$", f"FY{y}", GAAP, U_K25 + "；損益表 p.86", "Verified")
f("is.oi.2022", "營業利益（GAAP）", -161.201, "US$M", "FY2022", GAAP, U_K22 + "；MD&A", "Verified")
f("is.oi.2021", "營業利益（GAAP）", -411.046, "US$M", "FY2021", GAAP, U_K22 + "；MD&A", "Verified")
# ---- 2026 前三季 ----
N9 = {"cogs": ("營業成本", 895.954, "Derived（營收 5,881.458 − 毛利 4,985.504）"), "gp": ("毛利", 4985.504, GAAP), "oi": ("營業利益（GAAP）", 2750.847, GAAP),
      "pbt": ("稅前利益", 3117.947, "Derived（淨利 3,076.805 + 稅 41.142）"), "tax": ("所得稅費用", 41.142, GAAP), "ni": ("淨利", 3076.805, GAAP),
      "cfo": ("營業現金流", 3319.902, GAAP), "capex": ("資本支出", 29.969, GAAP), "da": ("折舊與攤銷", 22.031, GAAP), "sbc": ("股票薪酬（現金流量表）", 785.690, GAAP)}
for k, (nm, v, b) in N9.items():
    f(f"is.{k}.9M26", nm, v, "US$M", "2026 前三季", b, U_Q3Q + "；損益表 p.4／現金流量表 p.8", "Derived" if "Derived" in b else "Verified")
for k, nm, v in (("rev", "營收", 3068.644), ("oi", "營業利益（GAAP）", 838.621), ("ni", "淨利", 1023.037), ("cfo", "營業現金流", 1357.178), ("capex", "資本支出", 20.610), ("da", "折舊與攤銷", 19.127), ("sbc", "股票薪酬", 487.628), ("gp", "毛利", 2495.433)):
    f(f"is.{k}.9M25", nm, v, "US$M", "2025 前三季", GAAP, U_Q3Q + "；損益表／現金流量表（2025 欄）", "Verified")
f("is.oi.2026Q3", "營業利益（GAAP）", 1084.845, "US$M", "2026Q3", GAAP, U_Q3PR + "；Statements of Operations", "Verified")
f("is.oi.2026Q1", "營業利益（GAAP）", 753.998, "US$M", "2026Q1", GAAP, U_Q1PR, "Verified")
f("is.ni.2026Q3", "淨利", 1134.441, "US$M", "2026Q3", GAAP, U_Q3PR, "Verified")
f("is.sbc.2026Q3", "股票薪酬（合計）", 318.889, "US$M", "2026Q3", GAAP, U_Q3PR + "；Statements of Operations 註 1", "Verified")
f("is.sbc.2025Q3", "股票薪酬（合計）", 172.318, "US$M", "2025Q3", GAAP, U_Q3PR, "Verified")
f("is.cogs.2026Q3", "營業成本", 383.286, "US$M", "2026Q3", GAAP, U_Q3PR, "Verified")
f("is.opx.2026Q3", "營業費用（S&M 396.711＋R&D 219.720＋G&A 228.849）", 845.280, "US$M", "2026Q3", GAAP + "；Derived 相加", U_Q3PR, "Derived")
f("is.int.2026Q3", "利息收入", 88.936, "US$M", "2026Q3", GAAP, U_Q3PR, "Verified")
f("is.eps.2026Q3", "稀釋 EPS（GAAP）", 0.44, "US$", "2026Q3", GAAP, U_Q3PR, "Verified")
f("is.eps.9M26", "稀釋 EPS（GAAP）", 1.19, "US$", "2026 前三季", GAAP, U_Q3Q + "；損益表 p.4", "Verified")
f("adj.oi.2026Q3", "調整後營業利益", 1426.840, "US$M", "2026Q3", NG, U_Q3PR + "；Non-GAAP Reconciliations", "Interested-party", "公司自訂 non-GAAP 口徑（排除股票薪酬與其雇主稅）")
f("adj.oi.2026Q1", "調整後營業利益", 983.545, "US$M", "2026Q1", NG, U_Q1PR, "Interested-party", "公司自訂 non-GAAP 口徑")
f("adj.oi.9M26", "調整後營業利益", 3604.857, "US$M", "2026 前三季", NG, U_Q3PR + "；Nine-month 表", "Interested-party", "公司自訂 non-GAAP 口徑")
f("adj.oi.2025", "調整後營業利益", 2254.100, "US$M", "FY2025", NG, U_K25 + "；p.72", "Interested-party", "公司自訂 non-GAAP 口徑")
f("adj.fcf.2026Q3", "調整後自由現金流", 1216.077, "US$M", "2026Q3", NG, U_Q3PR, "Interested-party", "公司自訂 non-GAAP 口徑")
f("gm.gaap.2025", "GAAP 毛利率", 0.8237, "比率", "FY2025", "Derived（毛利 3,686.269 ÷ 營收 4,475.446）", U_K25, "Derived")
f("gm.gaap.9M26", "GAAP 毛利率", 0.8477, "比率", "2026 前三季", "Derived（4,985.504 ÷ 5,881.458）", U_Q3Q, "Derived")
f("om.gaap.9M26", "GAAP 營業利益率", 0.4677, "比率", "2026 前三季", "Derived（2,750.847 ÷ 5,881.458）", U_Q3Q, "Derived")
f("om.gaap.2025", "GAAP 營業利益率", 0.3160, "比率", "FY2025", "Derived（1,414.015 ÷ 4,475.446）", U_K25, "Derived")
f("sbc.rev.9M26", "股票薪酬占營收", 0.1336, "比率", "2026 前三季", "Derived（785.690 ÷ 5,881.458）", U_Q3Q, "Derived")
f("sbc.rev.2025", "股票薪酬占營收", 0.1528, "比率", "FY2025", "Derived（684.033 ÷ 4,475.446）", U_K25, "Derived")
# ---- 資產負債與股數（2026-09-30；封面 2026-10-26）----
f("bs.cash", "現金及約當現金", 3619.414, "US$M", "2026-09-30", GAAP, U_Q3Q + "；資產負債表 p.3", "Verified")
f("bs.sec", "有價證券（短期美國公債）", 6942.117, "US$M", "2026-09-30", GAAP, U_Q3Q + "；資產負債表 p.3", "Verified")
f("bs.debt", "借款餘額（2014 Credit Facility 循環額度 500 百萬，2027-03-31 到期）", 0, "US$M", "2026-09-30", GAAP, U_Q3Q + "；Note 6", "Verified")
f("bs.lease", "營業租賃負債（非流動）", 225.267, "US$M", "2026-09-30", GAAP, U_Q3Q + "；資產負債表 p.3", "Verified")
f("bs.cash25", "現金及約當現金", 1423.796, "US$M", "2025-12-31", GAAP, U_K25 + "；資產負債表 p.85", "Verified")
f("bs.sec25", "有價證券", 5753.247, "US$M", "2025-12-31", GAAP, U_K25 + "；資產負債表 p.85", "Verified")
f("bs.ar", "應收帳款淨額", 1883.498, "US$M", "2026-09-30", GAAP, U_Q3Q + "；資產負債表", "Verified")
f("sh.a", "Class A 流通股", 2306.705341, "百萬股", "2026-10-26", "封面", U_Q3Q + "；封面", "Verified")
f("sh.b", "Class B 流通股", 101.787133, "百萬股", "2026-10-26", "封面", U_Q3Q + "；封面", "Verified")
f("sh.f", "Class F 流通股", 1.005, "百萬股", "2026-10-26", "封面", U_Q3Q + "；封面", "Verified")
f("sh.rsu", "未歸屬 RSU", 35.503, "百萬股", "2026-09-30", "加權平均授予日公允價值 48.26 美元", U_Q3Q + "；Note 9", "Verified")
f("sh.opt", "選擇權流通數", 146.841, "百萬股", "2026-09-30", "加權平均履約價 10.16 美元", U_Q3Q + "；Note 9", "Verified")
f("sh.optk", "選擇權加權履約價", 10.16, "US$", "2026-09-30", "加權平均", U_Q3Q + "；Note 9", "Verified")
f("sh.sar", "SAR 流通數", 18.511, "百萬股", "2026-09-30", "加權平均履約價 134.77 美元", U_Q3Q + "；Note 9", "Verified")
f("sh.sark", "SAR 加權履約價", 134.77, "US$", "2026-09-30", "加權平均", U_Q3Q + "；Note 9", "Verified")
f("sh.prsu", "未歸屬 P-RSU", 0.140, "百萬股", "2026-09-30", "績效型；模型不計入（數量小，列為已知差異）", U_Q3Q + "；Note 9", "Verified")
f("sh.dil.q3", "Q3 稀釋加權平均股數", 2573.926, "百萬股", "2026Q3", GAAP, U_Q3PR, "Verified")
f("tax.rate.9M26", "有效稅率（前三季）", 0.0132, "比率", "2026 前三季", "Derived（41.142 ÷ 3,117.947）；美、英遞延所得稅資產仍全額備抵", U_Q3Q + "；Note 10", "Derived")
f("emp.2025", "全職員工數", 4429, "人", "2025-12-31", "其中 28% 在美國以外", U_K25 + "；p.11、15", "Verified")
# ---- 營運指標 ----
f("kpi.cust.q3", "客戶數（近 12 個月）", 1090, "家", "2026-09-30", "公司定義（TTM）", U_Q3Q + "；MD&A", "Interested-party", "公司自揭 KPI 與定義")
f("kpi.cust.q2", "客戶數（近 12 個月）", 1049, "家", "2026-06-30", "公司定義", U_Q2Q + "；MD&A", "Interested-party", "公司自揭 KPI")
f("kpi.cust.25", "客戶數（近 12 個月）", 954, "家", "2025-12-31", "公司定義", U_K25 + "；p.6", "Interested-party", "公司自揭 KPI")
f("kpi.usc.q3", "美國商業客戶數", 686, "家", "2026-09-30", "公司定義", U_FUT, "Interested-party", "公司自揭 KPI（經分析機構轉述，10-Q 未載；第二來源）")
f("kpi.usc.q2", "美國商業客戶數", 653, "家", "2026-06-30", "公司定義", U_Q2MB, "Interested-party", "公司自揭 KPI（經媒體轉述）")
f("kpi.top20.q3", "前 20 大客戶平均營收（近 12 個月）", 146, "US$M", "2026-09-30", "公司定義", U_Q3Q + "；MD&A", "Interested-party", "公司自揭 KPI")
f("kpi.top20.25", "前 20 大客戶平均營收（近 12 個月）", 93.9, "US$M", "FY2025", "公司定義", U_K25 + "；p.6", "Interested-party", "公司自揭 KPI")
f("kpi.top3.25", "前三大客戶占營收比", 0.16, "比率", "FY2025", "公司揭露", U_K25 + "；p.14", "Verified")
f("kpi.ndr.q3", "淨美元留存率（NDR）", 1.65, "比率", "2026Q3", "公司定義", U_FUT, "Interested-party", "公司自揭 KPI（第二來源；一手新聞稿本次未載）")
f("kpi.ndr.q2", "淨美元留存率（NDR）", 1.57, "比率", "2026Q2", "公司定義", U_Q2MB, "Interested-party", "公司自揭 KPI")
f("kpi.rpo.q3", "剩餘履約義務（RPO）", 4900, "US$M", "2026-09-30", "約 44% 於 12 個月內認列、37% 於第 13–36 個月", U_Q3Q + "；Note 3", "Verified")
f("kpi.rpo.25", "剩餘履約義務（RPO）", 4100, "US$M", "2025-12-31", "約 38% 於 12 個月內認列", U_K25 + "；Note 3 p.97", "Verified")
f("kpi.rdv.25", "總剩餘交易價值（RDV）", 11200, "US$M", "2025-12-31", "公司定義；商業 6.8B、政府 4.4B", U_K25 + "；p.15", "Interested-party", "公司自訂 KPI")
f("kpi.tcv.q3", "總合約價值（TCV）", 2918, "US$M", "2026Q3", "公司定義；年增 6%", U_Q3PR + "；Highlights", "Interested-party", "公司自訂 KPI")
f("kpi.usctcv.q3", "美國商業 TCV", 1463, "US$M", "2026Q3", "公司定義；年增 11%", U_Q3PR, "Interested-party", "公司自訂 KPI")
f("kpi.usctcv.q2", "美國商業 TCV", 2132, "US$M", "2026Q2", "公司定義；年增 153%", U_Q2MB, "Interested-party", "公司自訂 KPI（第二來源）")
f("kpi.uscrdv.q3", "美國商業剩餘交易價值", 6517, "US$M", "2026-09-30", "公司定義；年增 79%", U_Q3PR, "Interested-party", "公司自訂 KPI")
f("kpi.deals.q3", "≥1M／≥5M／≥10M 交易筆數", 212, "筆", "2026Q3", "≥100 萬美元交易 212；≥500 萬 107；≥1,000 萬 64", U_Q3PR, "Interested-party", "公司自揭 KPI")
f("kpi.r40.q3", "Rule of 40 分數", 1.58, "比率", "2026Q3", "公司定義", U_Q3PR, "Interested-party", "公司自訂 KPI")
# ---- 指引（公司）----
f("gd.q4.rev.lo", "2026Q4 營收指引下限", 2649, "US$M", "2026Q4", "公司指引", U_Q3PR + "；Outlook", "Interested-party", "公司指引；管理層保守傾向（歷來上修）")
f("gd.q4.rev.hi", "2026Q4 營收指引上限", 2653, "US$M", "2026Q4", "公司指引", U_Q3PR + "；Outlook", "Interested-party", "同上")
f("gd.fy.rev.lo", "FY2026 營收指引下限", 8530, "US$M", "FY2026", "公司指引（2026-11-03 上修）", U_Q3PR, "Interested-party", "同上")
f("gd.fy.rev.hi", "FY2026 營收指引上限", 8534, "US$M", "FY2026", "公司指引", U_Q3PR, "Interested-party", "同上")
f("gd.fy.usc", "FY2026 美國商業營收指引（下限，超過）", 3559, "US$M", "FY2026", "公司指引", U_Q3PR, "Interested-party", "同上")
f("gd.fy.adjoi.lo", "FY2026 調整後營業利益指引下限", 5245, "US$M", "FY2026", NG + "；公司指引", U_Q3PR, "Interested-party", "同上")
f("gd.fy.adjoi.hi", "FY2026 調整後營業利益指引上限", 5249, "US$M", "FY2026", NG + "；公司指引", U_Q3PR, "Interested-party", "同上")
f("gd.fy.fcf.lo", "FY2026 調整後自由現金流指引下限", 4800, "US$M", "FY2026", NG + "；公司指引", U_Q3PR, "Interested-party", "同上")
f("gd.fy.fcf.hi", "FY2026 調整後自由現金流指引上限", 5000, "US$M", "FY2026", NG + "；公司指引", U_Q3PR, "Interested-party", "同上")
# ---- 市場與共識 ----
f("mkt.px.1113", "PLTR 收盤價", 167.99, "US$", "2026-11-13", "普通股收盤", U_SA + "；表列 Nov 13, 2026", "Verified")
f("mkt.px.1109", "PLTR 收盤價", 169.20, "US$", "2026-11-09", "普通股收盤（與 heygotrade 11-09 報導相符）", U_SA, "Verified")
f("mkt.beta.y", "原始 β：Yahoo Finance「Beta (5Y Monthly)」", 1.60, "x", "頁面當日（日期未載）", "資料商計算", "https://finance.yahoo.com/quote/PLTR/key-statistics｜Yahoo Finance 關鍵統計頁", "Interested-party", "資料商計算值，無交易誘因；基準指數與日期未載")
f("mkt.beta.s", "原始 β：StockAnalysis「Beta (5Y)」", 1.61, "x", "2026-11-13 收盤", "資料商計算", "https://stockanalysis.com/stocks/pltr/statistics/｜StockAnalysis 統計頁（更新 2026-11-14）", "Interested-party", "資料商計算值，無交易誘因")
f("mkt.rf", "無風險利率（10 年美債）", 0.0510, "比率", "2026-11-12", "FRED DGS10 收盤（現價日 2026-11-13 前 1 個交易日）", "https://fred.stlouisfed.org/series/DGS10（工作單 G5 第 2 條指定美國 10 年期 5.10%）", "Verified")
f("mkt.mcap", "市值", 404770, "US$M", "2026-11-13", "資料商", "https://stockanalysis.com/stocks/pltr/statistics/", "Verified")
f("mkt.evs", "EV／營收（TTM）", 54.12, "x", "2026-11-13", "資料商", "https://stockanalysis.com/stocks/pltr/statistics/", "Verified")
f("cons.rev26", "賣方共識營收 FY2026（平均；29 家）", 8220, "US$M", "FY2026", "共識；2026-10-23 收盤，財報前", "https://finance.yahoo.com/quote/PLTR/analysis/｜Yahoo Finance 分析頁", "Interested-party", "賣方預估；投行業務關係；財報前，只作歷史對照")
f("cons.rev27", "賣方共識營收 FY2027（平均；28 家；高 14,620／低 11,000）", 12360, "US$M", "FY2027", "共識；2026-10-23，財報前", "同上", "Interested-party", "賣方預估；財報前，只作歷史對照")
f("cons.eps26", "賣方共識 EPS FY2026（平均；29 家）", 1.62, "US$", "FY2026", "共識（多為 non-GAAP）；2026-10-23，財報前", "同上", "Interested-party", "賣方預估；口徑與 GAAP 不同")
f("cons.eps27", "賣方共識 EPS FY2027（平均；29 家）", 2.36, "US$", "FY2027", "共識；2026-10-23，財報前", "同上", "Interested-party", "賣方預估；口徑與 GAAP 不同")
f("cons.tp", "賣方平均目標價（32 家；高 255／低 80）", 191.68, "US$", "2026-08-31", "共識；Q2 財報後、Q3 財報前", "https://ca.marketscreener.com/quote/stock/PALANTIR-TECHNOLOGIES-INC-113108869/consensus/｜MarketScreener", "Interested-party", "賣方目標價；投行業務關係；財報前，只作歷史對照")
f("cons.tp.rbc", "RBC 目標價（Underperform；2026-11-03 由 90 上調至 130）", 130, "US$", "2026-11-03", "單家賣方", "https://247wallst.com/investing/2026/11/03/palantirs-96-revenue-growth-drives-analyst-optimism-as-price-targets-rise｜247WallSt", "Interested-party", "單家賣方；Q3 後唯一低於現價者")
f("cons.tp.citi", "Citi 目標價（Buy；250→265）", 265, "US$", "2026-11-03", "單家賣方", "同上", "Interested-party", "單家賣方")
f("cons.tp.ms", "Morgan Stanley 目標價（Equal Weight；205→233）", 233, "US$", "2026-11-03", "單家賣方", "同上", "Interested-party", "單家賣方")
f("cons.tp.piper", "Piper Sandler 目標價（Overweight；231→235）", 235, "US$", "2026-11-03", "單家賣方", "同上", "Interested-party", "單家賣方")
# ---- 同業（StockAnalysis 統計頁；EV/Sales 為 TTM；3Y 營收成長為賣方預測）----
PEERS = [
 ("SNOW", "Snowflake", "商業", 18.58, 31.03, "2026-11-14", "https://stockanalysis.com/stocks/snow/statistics/"),
 ("DDOG", "Datadog", "商業", 14.62, 25.59, "2026-11-15", "https://stockanalysis.com/stocks/ddog/statistics/"),
 ("NET", "Cloudflare", "商業", 39.95, 32.98, "2026-11-14", "https://stockanalysis.com/stocks/net/statistics/"),
 ("MDB", "MongoDB", "商業", 8.21, 20.82, "2026-11-08", "https://stockanalysis.com/stocks/mdb/statistics/"),
 ("NOW", "ServiceNow", "商業", 6.95, 19.72, "2026-11-08", "https://stockanalysis.com/stocks/now/statistics/"),
 ("BAH", "Booz Allen Hamilton", "政府", 1.30, 3.31, "2026-10-26", "https://stockanalysis.com/stocks/bah/statistics/"),
 ("LDOS", "Leidos", "政府", 1.23, 4.16, "2026-11-15", "https://stockanalysis.com/stocks/ldos/statistics/"),
 ("CACI", "CACI International", "政府", 2.03, 8.35, "2026-10-26", "https://stockanalysis.com/stocks/caci/statistics/"),
]
for tk, nm, grp, evs, g3, dt, url in PEERS:
    f(f"peer.{tk}.evs", f"{nm}（{tk}）EV／營收", evs, "x", f"頁面更新 {dt}", "TTM 營收；US GAAP；EV 不含營業租賃", url, "Verified")
    f(f"peer.{tk}.g3", f"{nm}（{tk}）賣方 3 年營收成長預測", g3 / 100, "比率", f"頁面更新 {dt}", "賣方預測（年複合）", url, "Interested-party", "賣方預測")

import csv as _csv, os as _os
PSL = {}
for tk, nm, grp, evs, g3, dt, url in PEERS:
    _fp = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "..", "..", "data", "relmult_hist", f"{tk}.csv")
    _rows = [r for r in _csv.reader(open(_fp, encoding="utf-8-sig"))][1:]
    _rows = [r for r in _rows if r[5].strip()]
    PSL[tk] = (_rows[-1][0], float(_rows[-1][5]), _rows[-1][1])
    f(f"peer.{tk}.ps", f"{nm}（{tk}）P/S（市值÷TTM 營收；共用庫最近一季）", PSL[tk][1], "x", f"季末 {PSL[tk][0]}（財季實際期末 {PSL[tk][2]}）", "software_ai/data/relmult_hist/%s.csv 最近一季；Derived：市值÷TTM 營收（StockAnalysis 季度比率頁）" % tk, f"software_ai/data/relmult_hist/{tk}.csv", "Derived")
FACT_BY_ID = {x["id"]: x for x in FACTS}
def V(i): return FACT_BY_ID[i]["數值"]
FIELDS = ["指標", "數值", "單位", "期間", "口徑", "出處", "擷取日", "標記", "來源誘因"]

if __name__ == "__main__":
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "facts_palantir.csv")
    with open(out, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh); w.writerow(FIELDS)
        for x in FACTS: w.writerow([x[k] for k in FIELDS])
    # 恆等式核對（分部合計＝合併）
    for y in range(2021, 2026): assert abs(GOV[y] + COM[y] - REV[y]) < 0.0015, y
    for q in QREV: assert abs(QGOV[q] + QCOM[q] - QREV[q]) < 0.0015, q
    assert abs(sum(QGOV[q] for q in ("2026Q1", "2026Q2", "2026Q3")) - 2993.982) < 0.002
    assert abs(sum(QCOM[q] for q in ("2026Q1", "2026Q2", "2026Q3")) - 2887.476) < 0.002
    print(f"{len(FACTS)} 筆事實 → {out}")

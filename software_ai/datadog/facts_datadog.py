# Datadog（DDOG）事實總帳：逐列紀錄 → data/facts_datadog.csv（工作單共同規則第 4 條欄位）與 data/datadog_facts.json（build_datadog.py 讀取）。
# 單位：US$M（百萬美元）；股數：百萬股；每股：美元。會計年度＝日曆年（12 月 31 日止）。擷取日 2026-11-16。
# 用法：python3 software_ai/datadog/facts_datadog.py
import csv, json, os

HERE = os.path.dirname(os.path.abspath(__file__))
D = "2026-11-16"
U = {
    "q326pr": "https://www.sec.gov/Archives/edgar/data/0001561550/000162828026069792/ex-991x20260930x8k.htm（8-K 附件 99.1，2026Q3 財報新聞稿，2026-11-04；財務報表表格）",
    "q326q": "https://www.sec.gov/Archives/edgar/data/0001561550/000162828026070292/ddog-20260930.htm（10-Q 2026Q3）",
    "q226pr": "https://investors.datadoghq.com/news-releases/news-release-details/datadog-announces-second-quarter-2026-financial-results（2026Q2 財報新聞稿，日期未載於擷取內容）",
    "q126pr": "https://investors.datadoghq.com/news-releases/news-release-details/datadog-announces-first-quarter-2026-financial-results（2026Q1 財報新聞稿，指引日 2026-05-07）",
    "q425pr": "https://investors.datadoghq.com/news-releases/news-release-details/datadog-announces-fourth-quarter-and-fiscal-year-2025-financial（2025Q4／FY2025 財報新聞稿，2026-02-10）",
    "q424pr": "https://investors.datadoghq.com/news-releases/news-release-details/datadog-announces-fourth-quarter-and-fiscal-year-2024-financial（2024Q4／FY2024 財報新聞稿）",
    "k25": "https://investors.datadoghq.com/static-files/0a0cecbd-9dfe-43c8-9174-1bcffad62c20（10-K FY2025，股數日 2026-02-05）",
    "k22": "https://www.sec.gov/Archives/edgar/data/1561550/000156155023000006/ddog-20221231.htm（10-K FY2022）",
    "q421pr": "https://www.sec.gov/Archives/edgar/data/1561550/000156155022000004/ex-991x20211231x8k.htm（8-K 附件 99.1，FY2021 財報新聞稿）",
    "call326": "https://singjupost.com/transcript-of-datadog-ddog-q3-2026-earnings-call/（2026Q3 法說逐字稿，2026-11-04；第二來源）",
    "stack": "https://www.thestack.technology/customers-snapped-up-datadogs-ai-products-in-q3-but-budgets-are-getting-tighter/（The Stack，2026-11-04）",
    "secapi": "https://sec-api.io/insights/datadog-ddog-q3-2026-earnings-largest-customer-cuts-usage-as-rpo-reaches-3-85bn-sec-filing-analysis（sec-api.io 整理 2026-11-04；非原文，引述歷次法說口頭與 10-Q）",
    "call425": "https://www.fool.com/earnings/call-transcripts/2026/02/10/datadog-ddog-q4-2025-earnings-transcript/（2025Q4 法說逐字稿，2026-02-10）",
    "pricing": "https://www.datadoghq.com/pricing/?product=bits-ai-sre（Datadog 官網定價頁 AI Credits，2026-11-16 讀取）",
}
INC_CO = "公司管理層（KPI、法說口頭說法）：呈現成長與維持市場信心的誘因，口徑由公司定義、未經查核"
INC_3P = "第三方整理或媒體：無直接利益，但二手轉述、未經原文查核"
F = []   # (id, 指標, 數值, 單位, 期間, 口徑, 出處, 標記, 來源誘因)


def f(i, m, v, u, p, b, s, t="Verified", inc=""):
    F.append((i, m, v, u, p, b, s, t, inc))


# ---------- 損益：年度（US$M）----------
FY = {  # 年: rev, cogs, rd, sm, ga, oi, ni, sbc, ngoi
    2021: dict(rev=1028.784, cogs=234.245, oi=-19.156, ni=-20.745, sbc=163.737, ngoi=165.058),
    2022: dict(rev=1675.100, cogs=346.743, oi=-58.695, sbc=363.154),
    2023: dict(rev=2128.359, cogs=409.908, rd=962.447, sm=609.276, ga=180.192, oi=-33.464, ni=48.568, sbc=482.300, ngoi=490.243),
    2024: dict(rev=2684.275, cogs=515.531, rd=1152.703, sm=756.605, ga=205.152, oi=54.284, ni=183.746, sbc=570.336, ngoi=674.213),
    2025: dict(rev=3427.158, cogs=686.957, rd=1548.451, sm=956.423, ga=279.700, oi=-44.373, ni=107.741, sbc=750.671, ngoi=768.044),
}
SRC_FY = {2021: "q421pr", 2022: "k22", 2023: "q424pr", 2024: "q424pr", 2025: "q425pr"}
NAMES = dict(rev="營收", cogs="營業成本", rd="研發費用", sm="銷售與行銷費用", ga="一般及行政費用", oi="營業利益（GAAP）", ni="淨利（GAAP）", sbc="股票薪酬（總額）", ngoi="營業利益（non-GAAP）")
for y, d in FY.items():
    for k, v in d.items():
        f(f"is.{k}.fy{y}", NAMES[k], v, "US$M", f"FY{y}（12 月 31 日止）", "non-GAAP（公司定義：排除 SBC、員工股票交易薪資稅、併購攤銷與費用）" if k == "ngoi" else "GAAP", U[SRC_FY[y]])
f("is.ni.fy2022", "淨利（GAAP）", -50.2, "US$M", "FY2022（12 月 31 日止）", "GAAP", U["k25"] + "；Item 1A 風險因素稱 2022 年淨損 50.2 百萬美元（取整至 0.1）")
# 2024／2025 補充（新聞稿）
for y, d in {2024: dict(ngp=2201.053, ngrd=758.268, ngsm=629.007, ngga=139.565, ngni=653.828, ptx=43.126, amort=6.467, mna=0.0, da=54.933, ocf=870.603, capex=34.719, capsw=60.781, fcf=775.103, onet=149.656, pti=203.940, tax=20.194, dsh=358.636, bsh=336.172, eps=0.52, ngeps=1.82),
             2025: dict(ngp=2776.053, ngrd=1038.742, ngsm=793.082, ngga=176.184, ngni=746.582, ptx=53.799, amort=6.373, mna=1.574, da=55.756, ocf=1050.135, capex=49.578, capsw=85.840, fcf=914.717, onet=171.394, pti=127.021, tax=19.280, dsh=363.471, bsh=347.309, eps=0.31, ngeps=2.05, acq=117.982)}.items():
    nm = dict(ngp="毛利（non-GAAP）", ngrd="研發費用（non-GAAP）", ngsm="銷售與行銷費用（non-GAAP）", ngga="一般及行政費用（non-GAAP）", ngni="淨利（non-GAAP，21% 長期稅率）", ptx="員工股票交易薪資稅",
              amort="取得無形資產攤銷", mna="併購交易費用", da="折舊與攤銷", ocf="營運現金流", capex="購置不動產與設備", capsw="資本化軟體開發成本", fcf="自由現金流（公司定義：營運現金流 − 資本支出 − 資本化軟體）",
              onet="其他收入淨額（利息收入及其他 − 利息費用）", pti="稅前淨利", tax="所得稅費用", dsh="稀釋加權股數", bsh="基本加權股數", eps="稀釋 EPS（GAAP）", ngeps="稀釋 EPS（non-GAAP）", acq="收購現金支出（扣除取得現金）")
    un = {"dsh": "百萬股", "bsh": "百萬股", "eps": "US$", "ngeps": "US$"}
    for k, v in d.items():
        ng = k.startswith("ng") or k in ("ptx", "amort", "mna")
        f(f"is.{k}.fy{y}", nm[k], v, un.get(k, "US$M"), f"FY{y}（12 月 31 日止）", "non-GAAP" if k.startswith("ng") else "GAAP", U["q425pr"] if y == 2025 else U["q424pr"])
f("is.da.fy2023", "折舊與攤銷", 44.465, "US$M", "FY2023", "GAAP", U["q424pr"]); f("is.ocf.fy2023", "營運現金流", 659.954, "US$M", "FY2023", "GAAP", U["q424pr"])
f("is.capex.fy2023", "購置不動產與設備", 27.586, "US$M", "FY2023", "GAAP", U["q424pr"]); f("is.capsw.fy2023", "資本化軟體開發成本", 34.820, "US$M", "FY2023", "GAAP", U["q424pr"])
f("is.fcf.fy2023", "自由現金流（公司定義）", 597.548, "US$M", "FY2023", "公司定義", U["q424pr"]); f("is.ngp.fy2023", "毛利（non-GAAP）", 1744.434, "US$M", "FY2023", "non-GAAP", U["q424pr"])
f("is.dsh.fy2023", "稀釋加權股數", 350.292, "百萬股", "FY2023", "GAAP", U["q424pr"]); f("is.eps.fy2023", "稀釋 EPS（GAAP）", 0.14, "US$", "FY2023", "GAAP", U["q424pr"])
f("is.dsh.3q26", "稀釋加權股數（單季）", 375.014, "百萬股", "2026Q3", "GAAP", U["q326pr"])
f("is.pti.fy2023", "稅前淨利", 60.235, "US$M", "FY2023", "GAAP", U["k25"] + "（Note 損益表；有效稅率推算段）"); f("is.tax.fy2023", "所得稅費用", 11.667, "US$M", "FY2023", "GAAP", U["k25"])
f("is.ocf.fy2021", "營運現金流", 286.545, "US$M", "FY2021", "GAAP", U["q421pr"]); f("is.fcf.fy2021", "自由現金流（公司定義）", 250.520, "US$M", "FY2021", "公司定義", U["q421pr"])
# ---------- 年度 營收地理 ----------
for y, na, intl in ((2023, 1487.319, 641.040), (2024, 1874.321, 809.954), (2025, 2433.083, 994.075)):
    f(f"geo.na.fy{y}", "營收：北美（依客戶帳單地址）", na, "US$M", f"FY{y}", "GAAP", U["k25"] + "（Note 11，第 82 頁）")
    f(f"geo.intl.fy{y}", "營收：國際", intl, "US$M", f"FY{y}", "GAAP", U["k25"] + "（Note 11，第 82 頁）")
for y, us in ((2023, 1411.0), (2024, 1785.5), (2025, 2320.3)):
    f(f"geo.us.fy{y}", "營收：美國", us, "US$M", f"FY{y}", "GAAP", U["k25"] + "（Note 11 註腳，第 82 頁）")
for k, v, nm_ in (("na", 2429.953, "北美"), ("intl", 895.017, "國際")):
    f(f"geo.{k}.9m26", f"營收：{nm_}", v, "US$M", "9M26（2026-01-01 至 09-30）", "GAAP", U["q326q"] + "（Note 11）")
for k, v, nm_ in (("na", 1754.428, "北美"), ("intl", 719.536, "國際")):
    f(f"geo.{k}.9m25", f"營收：{nm_}", v, "US$M", "9M25", "GAAP", U["q326q"] + "（Note 11）")
# ---------- 季（US$M）----------
Q = {  # 季: rev, cogs, oi, ngoi, sbc, ni, eps
    "1q25": (761.553, 157.628, -12.420, 166.505, 164.265, 24.642, 0.07), "2q25": (826.760, 165.978, -35.500, 164.108, 180.462, 2.647, 0.01),
    "3q25": (885.651, 176.457, -5.809, 207.364, 200.558, 33.885, 0.10), "4q25": (953.194, 186.894, 9.356, 230.068, 205.386, 46.567, 0.13),
    "1q26": (1006.426, 209.228, 7.327, 223.497, 196.841, 52.574, 0.15), "2q26": (1121.454, 240.113, 5.455, 257.031, 220.251, 44.558, 0.12),
    "3q26": (1197.090, 237.167, 74.706, 334.289, 237.181, 118.316, 0.32), "4q24": (737.727, 144.178, 9.392, 179.168, 158.461, 45.594, 0.13)}
QS = {"1q25": "q126pr", "2q25": "q226pr", "3q25": "q326pr", "4q25": "q425pr", "1q26": "q126pr", "2q26": "q226pr", "3q26": "q326pr", "4q24": "q425pr"}
for q, vals in Q.items():
    for k, v, un_, nm_ in zip(("rev", "cogs", "oi", "ngoi", "sbc", "ni", "eps"), vals, ("US$M",) * 6 + ("US$",), ("營收", "營業成本", "營業利益（GAAP）", "營業利益（non-GAAP）", "股票薪酬", "淨利（GAAP）", "稀釋 EPS（GAAP）")):
        f(f"q.{k}.{q}", nm_, v, un_, f"20{q[2:]}Q{q[0]}", "non-GAAP" if k == "ngoi" else "GAAP", U[QS[q]])
f("q.cust100k.3q26", "ARR ≥ $100K 客戶數", 4860, "戶", "2026-09-30", "公司定義（約數）", U["q326pr"], "Interested-party", INC_CO)
f("q.cust100k.2q26", "ARR ≥ $100K 客戶數", 4720, "戶", "2026-06-30", "公司定義（約數）", U["q226pr"], "Interested-party", INC_CO)
f("q.cust100k.1q26", "ARR ≥ $100K 客戶數", 4550, "戶", "2026-03-31", "公司定義（約數）", U["q126pr"], "Interested-party", INC_CO)
f("q.cust100k.4q25", "ARR ≥ $100K 客戶數", 4310, "戶", "2025-12-31", "公司定義（約數）", U["q425pr"] + "；" + U["k25"], "Interested-party", INC_CO)
f("q.cust100k.3q25", "ARR ≥ $100K 客戶數", 4060, "戶", "2025-09-30", "公司定義（約數）", U["q326pr"], "Interested-party", INC_CO)
f("q.cust100k.4q24", "ARR ≥ $100K 客戶數", 3610, "戶", "2024-12-31", "公司定義（約數）", U["q424pr"], "Interested-party", INC_CO)
f("q.cust100k.4q22", "ARR ≥ $100K 客戶數", 2780, "戶", "2022-12-31", "公司定義（約數）", U["k22"], "Interested-party", INC_CO)
f("q.cust100k.4q21", "ARR ≥ $100K 客戶數", 2010, "戶", "2021-12-31", "公司定義（約數）", U["q421pr"], "Interested-party", INC_CO)
# ---------- 九個月累計（US$M）----------
M9 = dict(rev=3324.970, cogs=686.508, gp=2638.462, rd=1404.836, sm=894.640, ga=251.498, oi=87.488, intx=-9.609, intinc=157.363, onet=147.754, pti=235.242, tax=19.794, ni=215.448,
          dsh=371.363, bsh=356.067, ngp=2671.289, ngrd=957.098, ngsm=746.780, ngga=152.593, ngoi=814.818, sbc=654.274, amort=5.733, ptx=63.500, mna=3.823, iss=3.146, da=58.202,
          ocf=940.446, capex=32.119, capsw=89.897, fcf=818.430, acq=115.609, ngni=762.918, eps=0.59, ngeps=2.05)
M9N = dict(rev="營收", cogs="營業成本", gp="毛利（GAAP）", rd="研發費用", sm="銷售與行銷費用", ga="一般及行政費用", oi="營業利益（GAAP）", intx="利息費用", intinc="利息收入及其他收入淨額", onet="其他收入淨額",
           pti="稅前淨利", tax="所得稅費用", ni="淨利（GAAP）", dsh="稀釋加權股數", bsh="基本加權股數", ngp="毛利（non-GAAP）", ngrd="研發費用（non-GAAP）", ngsm="銷售與行銷費用（non-GAAP）",
           ngga="一般及行政費用（non-GAAP）", ngoi="營業利益（non-GAAP）", sbc="股票薪酬（總額）", amort="取得無形資產攤銷", ptx="員工股票交易薪資稅", mna="併購交易費用", iss="發行成本攤銷（利息費用）",
           da="折舊與攤銷", ocf="營運現金流", capex="購置不動產與設備", capsw="資本化軟體開發成本", fcf="自由現金流（公司定義）", acq="收購現金支出", ngni="淨利（non-GAAP，21% 稅率調整後）", eps="稀釋 EPS（GAAP）", ngeps="稀釋 EPS（non-GAAP）")
for k, v in M9.items():
    f(f"is.{k}.9m26", M9N[k], v, "百萬股" if k in ("dsh", "bsh") else ("US$" if "eps" in k else "US$M"), "9M26（2026-01-01 至 09-30）", "non-GAAP" if k.startswith("ng") else "GAAP", U["q326pr"])
M9_25 = dict(rev=2473.964, cogs=500.063, oi=-53.729, ni=61.174, ngoi=537.977, sbc=545.285, ocf=723.066, fcf=623.696, capex=40.686, capsw=58.684, da=39.125, ngp=1999.756, dsh=361.905, onet=127.280, pti=73.551, tax=12.377)
for k, v in M9_25.items():
    f(f"is.{k}.9m25", M9N.get(k, k), v, "百萬股" if k == "dsh" else "US$M", "9M25", "non-GAAP" if k.startswith("ng") else "GAAP", U["q326pr"])
# ---------- 資產負債表 ----------
for k, v, nm_ in (("cash", 881.230, "現金及約當現金"), ("mkt", 4352.357, "有價證券"), ("conv", 986.595, "可轉換票據帳面（非流動）"), ("dr_c", 1235.133, "遞延收入（流動）"), ("dr_nc", 65.081, "遞延收入（非流動）"),
                  ("ar", 793.977, "應收帳款淨額"), ("ppe", 438.173, "不動產及設備淨額"), ("gw", 706.644, "商譽")):
    f(f"bs.{k}.20260930", nm_, v, "US$M", "2026-09-30", "GAAP", U["q326pr"])
for k, v, nm_ in (("cash", 401.305, "現金及約當現金"), ("mkt", 4073.531, "有價證券"), ("conv", 983.449, "可轉換票據帳面（非流動）")):
    f(f"bs.{k}.20251231", nm_, v, "US$M", "2025-12-31", "GAAP", U["q425pr"])
f("debt.face.20260930", "2029 年到期 0.00% 可轉換票據本金", 1000.0, "US$M", "2026-09-30", "面額；到期 2029-12-01；初始轉換價約 217.60 美元（4.5955 股／千美元）", U["q326q"] + "（債務附註）")
f("conv.px", "2029 票據初始轉換價", 217.60, "US$", "2026-09-30", "約數", U["q326q"] + "；" + U["k25"] + "（Note 8，第 78–80 頁）")
f("rpo.20260930", "剩餘履約義務（RPO）", 3847.7, "US$M", "2026-09-30", "GAAP 附註；公司預期絕大部分於 24 個月內認列", U["q326q"] + "（Note 11）")
f("rpo.20251231", "剩餘履約義務（RPO）", 3461.2, "US$M", "2025-12-31", "GAAP 附註", U["k25"] + "（Note 11，第 82 頁）")
f("rpo.20241231", "剩餘履約義務（RPO）", 2273.1, "US$M", "2024-12-31", "GAAP 附註", U["k25"] + "（Note 11，第 82 頁）")
f("sh.out.20261028", "流通股數（Class A 337,307,033＋Class B 23,735,512）", 361.043, "百萬股", "2026-10-28", "10-Q 封面", U["q326q"] + "（封面）")
f("rsu.20260930", "未歸屬 RSU／PSU", 16.346867, "百萬股", "2026-09-30", "加權平均授予日公允價值 140.63 美元", U["q326q"] + "（Note 12）")
f("opt.n.20260930", "流通選擇權數量", 1.306895, "百萬股", "2026-09-30", "", U["q326q"] + "（Note 12）")
f("opt.k.20260930", "流通選擇權加權平均履約價", 8.12, "US$", "2026-09-30", "", U["q326q"] + "（Note 12）")
f("emp.20251231", "員工數（約）", 8100, "人", "2025-12-31", "約數；35 國，約 44% 在美國以外", U["k25"] + "（Item 1 Human Capital，第 12 頁）", "Interested-party", INC_CO)
f("cust.20251231", "客戶總數（約）", 32700, "戶", "2025-12-31", "公司定義（約數）", U["k25"] + "（Item 1，第 7 頁）", "Interested-party", INC_CO)
f("cust.20240930", "客戶總數（約，上年同期）", 32000, "戶", "2025-09-30", "法說口頭「約 32,000，上年同期」", U["call326"], "Interested-party", INC_CO)
f("cust.20260930", "客戶總數（約）", 33500, "戶", "2026-09-30", "法說口頭", U["call326"], "Interested-party", INC_CO)
f("cust1m.20251231", "ARR ≥ $1M 客戶數", 603, "戶", "2025-12-31", "公司定義（約數）", U["q425pr"], "Interested-party", INC_CO)
f("cust1m.20241231", "ARR ≥ $1M 客戶數", 462, "戶", "2024-12-31", "公司定義（約數）", U["q425pr"], "Interested-party", INC_CO)
f("nrr.20251231", "美元基礎淨留存率（過去 12 個月）", 1.20, "比率", "2025-12-31", "公司定義：同客戶 ARR 比較、排除新客戶；過去 12 個月點對點加權平均", U["k25"] + "（MD&A，第 48 頁）", "Interested-party", INC_CO)
f("nrr.20260930", "美元基礎淨留存率（過去 12 個月）", 1.21, "比率", "2026-09-30", "法說口頭「low 120s」；取 121％ 為代表值（區間下緣 120％）", U["stack"] + "；" + U["secapi"], "Interested-party", INC_CO)
f("grr.20260930", "毛留存率（過去 12 個月）", 0.965, "比率", "2026-09-30", "法說口頭「mid to high 90s」；取 96.5％ 為區間中點代表值（Assumed 取點）", U["call326"], "Interested-party", INC_CO)
for k, v in (("2p", 0.84), ("4p", 0.55), ("6p", 0.33), ("8p", 0.18), ("10p", 0.09)):
    f(f"prod.{k}.20251231", f"使用 {k[:-1]} 項以上產品的客戶占比", v, "比率", "2025-12-31", "公司定義（約數）", U["k25"] + "（MD&A，第 48 頁）", "Interested-party", INC_CO)
f("prod.4p.20260930", "使用 4 項以上產品的客戶占比", 0.61, "比率", "2026-09-30", "法說口頭（上年同期 54％）", U["stack"] + "；" + U["call326"], "Interested-party", INC_CO)
f("ai.cust.20260930", "AI 原生客戶數（約，超過）", 850, "戶", "2026-09-30", "法說口頭「over 850」；其中 39 戶年支出逾 100 萬美元", U["call326"], "Interested-party", INC_CO)
f("ai.cust.20260630", "AI 原生客戶數（約）", 750, "戶", "2026-06-30", "法說口頭（上季）", U["call326"], "Interested-party", INC_CO)
f("ai.sh.3q24", "AI 原生客戶占營收比", 0.06, "比率", "2024Q3", "法說口頭", U["secapi"], "Interested-party", INC_CO)
f("ai.sh.2q25", "AI 原生客戶占營收比", 0.11, "比率", "2025Q2", "法說口頭", U["secapi"], "Interested-party", INC_CO)
f("ai.sh.3q25", "AI 原生客戶占營收比", 0.12, "比率", "2025Q3", "法說口頭（sec-api 整理；2025Q4 起未再揭露）", U["secapi"], "Interested-party", INC_CO)
f("ai.pp.4q25", "AI 原生客戶對營收年增率的貢獻", 7.0, "百分點", "2025Q4", "10-K 風險因素：約 7 個百分點", U["k25"] + "（Item 1A，第 16 頁）", "Interested-party", INC_CO)
f("ai.nonai.3q26", "非 AI 客戶營收年增（法說：high 20s）", 0.28, "比率", "2026Q3", "法說口頭「高 20%」區間；取 28％ 為代表值（區間 27％–29％）；上季約低 1 點", U["call326"], "Interested-party", INC_CO)
f("ai.exlargest.4q26g", "Q4 2026 指引：排除最大客戶的成長（low-to-mid 30s）", 0.33, "比率", "2026Q4 指引", "法說口頭；取 33％ 為代表值", U["secapi"] + "；" + U["call326"], "Interested-party", INC_CO)
f("ai.largest.minq4", "最大客戶 Q4 指引假設：以合約最低承諾計", 1, "旗標", "2026Q4 指引", "法說口頭：「our Q4 guidance contemplates their revenue at their minimum commitment」；金額未揭露", U["call326"], "Interested-party", INC_CO)
f("ai.exlargest.fy26", "FY2026 指引假設：排除最大客戶，全年至少成長 20%", 0.20, "比率", "FY2026（2026-02 指引時）", "CFO 口頭", U["call425"], "Interested-party", INC_CO)
# 價格與用量
f("px.credit.annual", "AI Credits 年約牌價", 1.00, "US$/credit", "2026-11-16", "$500／500 credits（月）；年繳", U["pricing"], "Verified", "")
f("px.credit.od", "AI Credits 隨需牌價", 1.30, "US$/credit", "2026-11-16", "on-demand", U["pricing"], "Verified", "")
f("px.credit.inv", "Bits Investigation 預估耗用 credits／次", 4, "credits", "2026-11-16", "官網稱「僅供參考」，實際依用量", U["pricing"], "Interested-party", "公司定價頁估算：呈現產品親和價格的誘因；實際耗用依任務而異")
f("ai.pen.3q26", "使用一項以上 AI 產品的客戶占比", 1 / 3, "比率", "2026Q3", "法說口頭「about a third」；$100K 以上客戶約四分之三", U["stack"], "Interested-party", INC_CO)
# 指引與市場
for k, v, nm_, per in (("g_q4lo", 1190.0, "Q4 2026 營收指引下緣", "2026Q4"), ("g_q4hi", 1200.0, "Q4 2026 營收指引上緣", "2026Q4"), ("g_fylo", 4515.0, "FY2026 營收指引下緣", "FY2026"), ("g_fyhi", 4525.0, "FY2026 營收指引上緣", "FY2026"),
                       ("g_ng_fylo", 1095.0, "FY2026 non-GAAP 營業利益指引下緣", "FY2026"), ("g_ng_fyhi", 1105.0, "FY2026 non-GAAP 營業利益指引上緣", "FY2026"),
                       ("g_ngeps_fylo", 2.73, "FY2026 non-GAAP EPS 指引下緣", "FY2026"), ("g_ngeps_fyhi", 2.75, "FY2026 non-GAAP EPS 指引上緣", "FY2026"),
                       ("g_ng_q4lo", 280.0, "Q4 2026 non-GAAP 營業利益指引下緣", "2026Q4"), ("g_ng_q4hi", 290.0, "Q4 2026 non-GAAP 營業利益指引上緣", "2026Q4"),
                       ("g_ngeps_q4lo", 0.69, "Q4 2026 non-GAAP EPS 指引下緣", "2026Q4"), ("g_ngeps_q4hi", 0.71, "Q4 2026 non-GAAP EPS 指引上緣", "2026Q4")):
    f(f"guid.{k}", nm_, v, "US$" if "eps" in k else "US$M", per, "non-GAAP（公司指引，2026-11-04）" if "ng" in k else "公司指引（2026-11-04）", U["q326pr"], "Interested-party", "公司管理層指引：偏保守以便超越的誘因（法說自述「imply conservatism」）")
f("guid.q4dil", "Q4 2026 指引假設稀釋股數", 379.0, "百萬股", "2026Q4", "公司指引假設", U["q326pr"], "Interested-party", INC_CO)
f("guid.fydil", "FY2026 指引假設稀釋股數", 375.0, "百萬股", "FY2026", "公司指引假設", U["q326pr"], "Interested-party", INC_CO)
f("guid.cx", "FY2026 資本支出與資本化軟體占營收比（指引區間 3％–4％）", 0.035, "比率", "FY2026", "法說口頭（CFO）；取區間中點 3.5％", U["stack"] + "；" + U["call326"], "Interested-party", INC_CO)
f("guid.beat1", "歷史：2026Q1 營收實績超出指引上緣（％）", 1006.426 / 961.0 - 1, "比率", "2026Q1", "實績 ÷ Q1 指引上緣 961（US$M）− 1", U["q126pr"] + "；" + U["q425pr"], "Derived")
f("guid.beat2", "歷史：2026Q2 營收實績超出指引上緣（％）", 1121.454 / 1080.0 - 1, "比率", "2026Q2", "實績 ÷ Q2 指引上緣 1,080 − 1", U["q226pr"] + "；" + U["q126pr"], "Derived")
f("guid.beat3", "歷史：2026Q3 營收實績超出指引上緣（％）", 1197.090 / 1145.0 - 1, "比率", "2026Q3", "實績 ÷ Q3 指引上緣 1,145 − 1", U["q326pr"] + "；" + U["q226pr"], "Derived")
f("mkt.px.20261113", "收盤價", 184.16, "US$", "2026-11-13", "Nasdaq 收盤", "https://stockanalysis.com/stocks/ddog/statistics/（2026-11-16 讀取；fn2.ai 分析師頁前收價 184.16 相符）")
f("mkt.px.20261112", "收盤價", 184.46, "US$", "2026-11-12", "Nasdaq 收盤", "https://www.tipranks.com/stocks/ddog/stock-forecast（2026-11-16 讀取）；stockanalysis 前收價 184.46 相符")
f("mkt.beta.si", "β（5Y 月）", 1.51, "x", "2026-11-13", "stockanalysis", "https://stockanalysis.com/stocks/ddog/statistics/")
f("mkt.beta.yh", "β（5Y 月）", 1.51, "x", "2026-10-23 頁面", "Yahoo Finance", "https://finance.yahoo.com/quote/DDOG/key-statistics/（2026-11-16 讀取，頁面價格日 2026-10-23）")
f("mkt.rf.20261113", "10 年期美國公債殖利率", 5.15, "%", "2026-11-13", "財政部日殖利率曲線", "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202611")
f("cons.tp.tipranks", "賣方平均目標價（高 335、低 171；家數頁面未載）", 276.92, "US$", "2026-11-12", "12 個月目標價", "https://www.tipranks.com/stocks/ddog/stock-forecast（2026-11-16 讀取）", "Interested-party", "賣方分析師：評等與投行／交易業務誘因，目標價傾向跟隨股價")
f("cons.tp.tikr", "賣方平均目標價（家數與日期頁面未載）", 272.0, "US$", "約 2026-11-09", "12 個月目標價", "https://www.tikr.com/blog/datadog-sinks-17-in-a-week-despite-strong-adoption-across-our-platform", "Interested-party", "賣方分析師；二手整理")
f("cons.tp.fn2", "賣方平均目標價（52 週區間；高 335、低 75.76）", 237.0, "US$", "2026-11-15", "頁面標示；口徑未載明（含舊目標）", "https://fn2.ai/earnings/ddog/analysts", "Interested-party", "賣方分析師；二手整理")
f("cons.rev27g", "市場共識 FY2027 營收成長（財報後）", 0.23, "比率", "2026-11-09", "TIKR 文字：分析師預期 2027 成長 23％；營收與 EPS 絕對值未取得", "https://tikr.com/blog/datadog-sinks-13-as-flat-q4-revenue-outlook-overshadows-earnings-beat", "Interested-party", "賣方分析師；二手整理")
f("cons.rev27.pre", "市場共識 FY2027 營收（財報前）", 5490.0, "US$M", "2026-10-26（財報前，只作歷史對照）", "stockanalysis 摘要「next-year」", "https://stockanalysis.com/stocks/ddog/forecast/（2026-11-16 讀取，頁面日 2026-10-26）", "Interested-party", "賣方分析師；財報前")
f("cons.eps27.pre", "市場共識 FY2027 EPS（non-GAAP，財報前）", 2.98, "US$", "2026-10-26（財報前，只作歷史對照）", "stockanalysis 摘要「next-year」", "https://stockanalysis.com/stocks/ddog/forecast/", "Interested-party", "賣方分析師；財報前")
f("cons.rev26.pre", "市場共識 FY2026 營收（財報前）", 4500.0, "US$M", "2026-10-26（財報前，只作歷史對照）", "stockanalysis（45 家；四捨五入至十億）", "https://stockanalysis.com/stocks/ddog/forecast/", "Interested-party", "賣方分析師；財報前")
f("cons.eps26.pre", "市場共識 FY2026 EPS（non-GAAP，財報前）", 2.53, "US$", "2026-10-26（財報前，只作歷史對照）", "stockanalysis（45 家）", "https://stockanalysis.com/stocks/ddog/forecast/", "Interested-party", "賣方分析師；財報前")
# 同業（stockanalysis 統計頁；Forward PS／PE 的分母口徑未載明；日期不一）
PEERS = [("Snowflake（SNOW）", "2026-11-13", 14.31, 115.37, 18.58, 0.6703, 0.2199, 0.0, 0.3103), ("Cloudflare（NET）", "2026-11-13", 29.98, 174.39, 39.95, 0.7190, 0.1292, 0.0, 0.3298),
         ("MongoDB（MDB）", "2026-11-06", 7.67, 45.04, 8.21, 0.7269, 0.2368, 0.0, 0.2082), ("Dynatrace（DT）", "2026-11-06", 4.97, 22.65, 5.43, 0.8134, 0.2653, 0.0, 0.1569),
         ("Elastic（ESTC）", "2026-11-06", 4.36, 24.48, 4.53, 0.7559, 0.1937, 0.0, 0.1488)]
for nm_, dt, ps, pe, evs, gm, fcfm, _, g3 in PEERS:
    for k, v, un_, mt in (("fps", ps, "x", "Forward P/S"), ("fpe", pe, "x", "Forward P/E"), ("evs", evs, "x", "EV/Sales（TTM）"), ("gm", gm, "比率", "毛利率（TTM）"), ("fcfm", fcfm, "比率", "自由現金流率（TTM）"), ("g3", g3, "比率", "三年營收成長預估（資料商）")):
        f(f"peer.{k}.{nm_[nm_.index('（') + 1:-1].lower()}", f"{nm_} {mt}", v, un_, dt, "資料商計算；Forward 分母（FY 或 NTM）未載明；GAAP／IFRS 與租賃處理未載明", f"https://stockanalysis.com/stocks/{nm_[nm_.index('（') + 1:-1].lower()}/statistics/（2026-11-16 讀取）",
          "Interested-party", "資料商計算值（共識預估為分母）：無直接利益，但分母為賣方預估，口徑與日期未經查核")
f("peer.fps.ddog", "Datadog Forward P/S", 12.73, "x", "2026-11-13", "同上口徑", "https://stockanalysis.com/stocks/ddog/statistics/", "Interested-party", "資料商計算值")
f("peer.fpe.ddog", "Datadog Forward P/E", 61.83, "x", "2026-11-13", "同上口徑", "https://stockanalysis.com/stocks/ddog/statistics/", "Interested-party", "資料商計算值")
f("fx.taxrate.ng", "non-GAAP 長期稅率", 0.21, "比率", "FY2026 起", "公司採用", U["q326pr"], "Interested-party", INC_CO)
f("tax.paid.fy2025", "已付所得稅", 17.649, "US$M", "FY2025", "現金流量表補充", U["k25"] + "（第 66 頁）")
f("tax.cash.fy26g", "FY2026 現金稅預期", 35.0, "US$M", "FY2026", "法說口頭：3,000 萬至 4,000 萬美元（取中點）", U["secapi"], "Interested-party", INC_CO)
f("pc.cloud.commit", "不可取消採購承諾（主要為雲端代管）", 1400.0, "US$M", "2025-12-31", "10-K 附註；供應商別與到期時點未揭露", U["k25"] + "（Note 9，第 80 頁）")
f("cogs.host.9m26", "第三方雲端基礎設施代管與軟體成本增加（9M26 較 9M25）", 158.1, "US$M", "9M26", "10-Q MD&A：COGS 增 186.4（+37％）中的雲端代管增額", U["q326q"] + "（MD&A）")


for k, v, nm_ in (("ngp", 776.297, "毛利（non-GAAP）"), ("ngrd", 280.263, "研發費用（non-GAAP）"), ("ngsm", 219.144, "銷售與行銷費用（non-GAAP）"), ("ngga", 46.822, "一般及行政費用（non-GAAP）"),
                  ("sbc", 205.386, "股票薪酬（總額）"), ("ptx", 13.283, "員工股票交易薪資稅"), ("amort", 1.842, "取得無形資產攤銷"), ("mna", 0.201, "併購交易費用"), ("ngoi", 230.068, "營業利益（non-GAAP）")):
    f(f"is.{k}.4q25", nm_, v, "US$M", "2025Q4", "non-GAAP" if k.startswith("ng") else "GAAP", U["q425pr"])

# ---------- 輸出 ----------
os.makedirs(os.path.join(HERE, "data"), exist_ok=True)
with open(os.path.join(HERE, "data", "facts_datadog.csv"), "w", newline="", encoding="utf-8-sig") as fh:
    w = csv.writer(fh)
    w.writerow(["指標", "數值", "單位", "期間", "口徑", "出處", "擷取日", "標記", "來源誘因", "id"])
    for i, m, v, u, p, b, s, t, inc in F:
        w.writerow([m, v, u, p, b, s, D, t, inc, i])
ids = [x[0] for x in F]
assert len(ids) == len(set(ids)), [i for i in ids if ids.count(i) > 1]
json.dump({"asOf": D, "items": [dict(id=i, metric=m, value=v, unit=u, period=p, basis=b, source=s, tag=t, incentive=inc) for i, m, v, u, p, b, s, t, inc in F]},
          open(os.path.join(HERE, "data", "datadog_facts.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(f"事實 {len(F)} 列")

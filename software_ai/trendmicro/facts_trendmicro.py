#!/usr/bin/env python3
"""趨勢科技（4704）事實總帳：產生 data/facts_trendmicro.csv 與 data/facts_trendmicro.json（建置腳本讀 json）。
每列：id、指標、數值、單位、期間、口徑、出處、擷取日、標記、來源誘因。標記只用 Verified／Interested-party／Analogy／Assumed／Derived。
金額單位：百萬日圓（除非另註）。擷取日一律 2026-11-16（本代理讀取日）。
用法：python3 software_ai/trendmicro/facts_trendmicro.py
"""
import csv, json, os

HERE = os.path.dirname(os.path.abspath(__file__))
D = "2026-11-16"
U = "https://www.trendmicro.com/content/dam/trendmicro/global/en/about/investor-relations/shareholder-meetings/"
AR25 = U + "annual-report-fiscal-year-2025.pdf"
AR24 = U + "annual-report-fiscal-year-2024.pdf"
AR23 = U + "annual-report-fiscal-year-2023.pdf"
R_Q4 = "https://newsroom.trendmicro.com/2026-02-18-Trend-Micro-Reports-Earnings-Results-for-Q4-2025-and-FY-2025"
R_Q1 = "https://newsroom.trendmicro.com/2026-05-14-Trend-Micro-Reports-Earnings-Results-for-Q1-2026"
R_Q2 = "https://newsroom.trendmicro.com/2026-08-13-Trend-Micro-Reports-Earnings-Results-for-Q2-2026"
R_Q3 = "https://www.prnewswire.com/news-releases/trend-micro-reports-q3-2026-results-with-9-net-sales-growth-and-an-18-operating-margin-302938200.html"
CALL3 = "https://www.investing.com/news/transcripts/earnings-call-transcript-trend-micro-posts-record-q3-sales-stock-slips-085-on-outlook-93CH-5022947"
SA_IS = "https://stockanalysis.com/quote/tyo/4704/financials/income-statement/"
SA_CF = "https://stockanalysis.com/quote/tyo/4704/financials/cash-flow-statement/"
SA_BS = "https://stockanalysis.com/quote/tyo/4704/financials/balance-sheet/"
SA_Q = "https://stockanalysis.com/quote/tyo/4704/"
SA_F = "https://stockanalysis.com/quote/tyo/4704/forecast/"
ENV_S = "https://envalith.com/en/companies/4704/segments"
ENV_O = "https://envalith.com/en/companies/4704"
GYO = "https://www.gyokaidigest.com/en/companies/trend-micro/report/2026-Q2"
EDB = "https://edinetdb.jp/company/E04999/text?fy=2025"
TE = "https://tradingeconomics.com/japan/government-bond-yield/news/600233"
REUT = "https://www.thestar.com.my/tech/tech-news/2025/02/13/exclusive-buyout-firms-vie-for-cybersecurity-firm-trend-micro-sources-say"
IV = "公司自揭 KPI／管理層說法：有為自身辯護、呈現成長的誘因"
IV_SV = "賣方預估：賣方有維持客戶關係與評等立場的誘因；資料商轉載"

F = []
def f(i, metric, val, unit, period, basis, src, tag, inc=""):
    F.append(dict(id=i, metric=metric, value=val, unit=unit, period=period, basis=basis, source=src, date=D, tag=tag, incentive=inc))

# ---- 合併損益（日本會計準則；12 月制）----
for fy, rev, ar in ((2021, 190359, SA_IS + "（與 " + ENV_O + " 二源一致；法定申報原文未讀）"), (2022, 223795, AR25 + " p.4 事業成績表"), (2023, 248691, AR25 + " p.4"),
                    (2024, 272638, AR25 + " p.4；AR24 p.26"), (2025, 275984, AR25 + " p.4、p.20")):
    f(f"rev.fy{fy}", "合併營收（net sales）", rev, "百萬日圓", f"FY{fy}（{fy}-01-01～{fy}-12-31）", "日本會計準則（J-GAAP）合併", ar, "Verified")
for fy, oi, ar in ((2021, 43642, SA_IS + "（envalith 43,641；差 1 為四捨五入）"), (2022, 31341, SA_IS + "（envalith 31,340）"), (2023, 32603, AR23 + " p.2 為 32,602；" + SA_IS + " 為 32,603"),
                   (2024, 48105, AR24 + " p.26"), (2025, 57777, AR25 + " p.1、p.20")):
    f(f"oi.fy{fy}", "營業利益（財報口徑）", oi, "百萬日圓", f"FY{fy}", "J-GAAP 合併；財報口徑", ar, "Verified")
for fy, v, ar in ((2021, 42017, SA_IS), (2022, 54300, SA_IS), (2023, 63662, SA_IS), (2024, 64980, AR24 + " p.26"), (2025, 63717, AR25 + " p.20")):
    f(f"cogs.fy{fy}", "營業成本（cost of sales）", v, "百萬日圓", f"FY{fy}", "J-GAAP 合併", ar, "Verified")
for fy, v, ar in ((2024, 159551, AR24 + " p.26"), (2025, 154490, "AR25 p.20 之 毛利 212,267 − 營業利益 57,777（" + SA_IS + " 列 SG&A 154,490）")):
    f(f"sga.fy{fy}", "銷管費（含研發；SG&A）", v, "百萬日圓", f"FY{fy}", "J-GAAP 合併", ar, "Verified" if fy == 2024 else "Derived")
for fy, v in ((2021, 148342), (2022, 169495), (2023, 185029), (2024, 207658), (2025, 212267)):
    f(f"gp.fy{fy}", "毛利", v, "百萬日圓", f"FY{fy}", "營收 − 營業成本", "由 rev 與 cogs 相減；與 " + SA_IS + " 一致", "Derived")
for fy, v, ar in ((2021, 52371, SA_IS), (2022, 41897, SA_IS), (2023, 31976, SA_IS), (2024, 55579, SA_IS), (2025, 52329, SA_IS)):
    f(f"pti.fy{fy}", "稅前利益", v, "百萬日圓", f"FY{fy}", "J-GAAP 合併", ar + "（資料商；法定申報原文未讀）", "Verified")
for fy, v, ar in ((2021, 38367, SA_IS), (2022, 29843, AR25 + " p.4"), (2023, 10731, AR25 + " p.4"), (2024, 34358, AR25 + " p.4"), (2025, 34523, AR25 + " p.4")):
    f(f"ni.fy{fy}", "歸屬母公司淨利", v, "百萬日圓", f"FY{fy}", "J-GAAP 合併", ar, "Verified")
for fy, v, ar in ((2022, 213.59, AR25 + " p.4"), (2023, 78.45, AR25 + " p.4"), (2024, 259.05, AR25 + " p.4"), (2025, 262.42, AR25 + " p.4、p.39")):
    f(f"eps.fy{fy}", "每股盈餘（基本）", v, "日圓", f"FY{fy}", "J-GAAP", ar, "Verified")
for fy, v, ar in ((2023, 53073, AR23 + " p.2"), (2024, 77636, AR24 + " p.1"), (2025, 80799, AR25 + " p.2")):
    f(f"pregaap.fy{fy}", "Pre-GAAP 營業利益（公司自訂；加回遞延收入調整）", v, "百萬日圓", f"FY{fy}", "公司 non-GAAP 類指標", ar, "Interested-party", IV)

# ---- 官方報告分部（所在地別；AR25 p.38 附註）----
REG = {"Japan": ("日本", [None, None, 83002, 85756, 87840]), "Americas": ("美洲", [None, 52930, 57643, 58827, 55187]),
       "Europe": ("歐洲", [None, 41479, 50643, 58546, 61439]), "Asia Pacific": ("亞太", [None, 47284, 57401, 69507, 71516])}
for rn, (zh, vals) in REG.items():
    for fy, v in zip((2021, 2022, 2023, 2024, 2025), vals):
        if v is None:
            continue
        src = {2023: AR23 + " p.1（營運概況）", 2024: AR24 + " p.40 附註", 2025: AR25 + " p.38 附註", 2022: "由 AR23 p.1 之 FY2023 值與年增率回推（≈）"}[fy]
        f(f"seg.{rn}.fy{fy}", f"分部營收：{rn}（{zh}）", v, "百萬日圓", f"FY{fy}", "所在地別報告分部；外部客戶營收", src, "Derived" if fy == 2022 else "Verified")
for rn, v, ar in (("Japan", 20651, ENV_S), ("Americas", 11027, ENV_S), ("Europe", 12334, ENV_S), ("Asia Pacific", 12982, ENV_S)):
    f(f"segoi.{rn}.fy2025", f"分部利益：{rn}", v, "百萬日圓", "FY2025", "報告分部利益（調整前；合計 56,994 與合併營業利益 57,777 差 783 為調整）", ar + "（資料商；法定申報原文未讀）", "Verified")
for rn, v25, v26 in (("Japan", 22072, 22136), ("Americas", 14174, 14498), ("Europe", 14261, 16990), ("Asia Pacific", 16993, 20231)):
    f(f"seg.{rn}.q1fy2025", f"分部營收：{rn}", v25, "百萬日圓", "2025Q1", "所在地別報告分部", ENV_S + "（資料商；法定申報原文未讀）", "Verified")
    f(f"seg.{rn}.q1fy2026", f"分部營收：{rn}", v26, "百萬日圓", "2026Q1", "所在地別報告分部", ENV_S + "（資料商；法定申報原文未讀）", "Verified")
for rn, v, yy in (("Japan", 44184, 0.007), ("Americas", 29059, 0.075), ("Europe", 33865, 0.166), ("Asia Pacific", 41486, 0.222)):
    f(f"seg.{rn}.h1fy2026", f"分部營收：{rn}", v, "百萬日圓", "2026H1（1–6 月累計）", f"所在地別報告分部；年增 {yy:.1%}", GYO + "（資料商轉載短信；法定申報原文未讀）", "Verified")

# ---- 季度合併 ----
for q, rev, oi, ni, src in (("2025Q4", 73230, 13303, 7580, R_Q4), ("2026Q1", 73856, 15558, 11775, R_Q1), ("2026Q2", 74739, 6233, 3455, R_Q2), ("2026Q3", 75300, 13200, 8600, R_Q3 + "（新聞稿以十億日圓一位小數表達：75.3／13.2／8.6）")):
    f(f"rev.{q}", "合併營收", rev, "百萬日圓", q, "J-GAAP 合併；季度", src, "Verified")
    f(f"oi.{q}", "營業利益", oi, "百萬日圓", q, "J-GAAP 合併；季度", src, "Verified")
    f(f"ni.{q}", "歸屬母公司淨利", ni, "百萬日圓", q, "J-GAAP 合併；季度", src, "Verified")
f("rev.h1fy2026", "合併營收（上半年累計）", 148596, "百萬日圓", "2026H1", "J-GAAP；年增 11.0%", GYO, "Verified")
f("oi.h1fy2026", "營業利益（上半年累計）", 21792, "百萬日圓", "2026H1", "J-GAAP；年增 −23.5%", GYO, "Verified")
f("ordinary.h1fy2026", "經常利益（上半年累計）", 24824, "百萬日圓", "2026H1", "J-GAAP", GYO, "Verified")
f("ent.h1fy2026", "企業業務（TrendAI）營收（上半年累計）", 119444, "百萬日圓", "2026H1", "公司分業務營收；年增 11.9%", GYO, "Verified")
f("con.h1fy2026", "個人業務（TrendLife）營收（上半年累計）", 29151, "百萬日圓", "2026H1", "公司分業務營收；年增 7.2%", GYO, "Verified")
f("cfo.h1fy2026", "營業現金流（上半年累計）", 35541, "百萬日圓", "2026H1", "J-GAAP 合併現金流", GYO, "Verified")
f("bb.h1fy2026", "庫藏股買回（上半年累計）", 4999, "百萬日圓", "2026H1", "現金支出", GYO, "Verified")

# ---- 指引 ----
f("guide.rev26", "FY2026 營收指引", 301500, "百萬日圓", "FY2026", "J-GAAP；匯率假設 156 日圓／美元；2026-02-18 發布、2026-05-14、08-13、11-12 維持", R_Q4 + "；" + R_Q3, "Interested-party", "公司指引：管理層設定可達成的目標")
f("guide.oi26.init", "FY2026 營業利益指引（初始）", 56400, "百萬日圓", "FY2026", "2026-02-18 發布；2026-05-14 未修正", R_Q4 + "；" + R_Q1, "Interested-party", "公司指引")
f("guide.oi26", "FY2026 營業利益指引（現行）", 44400, "百萬日圓", "FY2026", "2026-08-13 下修（−12,000）；2026-11-12 維持", R_Q2 + "；" + R_Q3, "Interested-party", "公司指引；下修理由寫「雲端相關費用（含 AI token 成本）顯著超出原預估」")
f("guide.ni26", "FY2026 歸屬母公司淨利指引（現行）", 30700, "百萬日圓", "FY2026", "2026-08-13 由 36,600 下修", R_Q2 + "；" + R_Q3, "Interested-party", "公司指引")
f("guide.ord26", "FY2026 經常利益指引（現行）", 46000, "百萬日圓", "FY2026", "2026-08-13 由 55,100 下修（資料商轉載）", GYO, "Interested-party", "公司指引")
f("target.opm28", "FY2028 營業利益率目標", "28–31", "%", "FY2028", "管理層法說口頭；2026-11-12 維持；FY2029–2030 路線圖不再延伸、年度 IR Day 取消", CALL3, "Interested-party", IV)

# ---- ARR 與 KPI（公司自揭；美元按固定匯率 155.95）----
for i, m, v, u, p, b, s in (
    ("arr.total.q3", "全公司 ARR", 1760, "US$M", "2026Q3", "固定匯率 155.95；年增 6%（Q1 3%、Q2 5%）", R_Q3),
    ("arr.ent.q3", "企業（TrendAI）ARR", 1292, "US$M", "2026Q3", "固定匯率；年增 8%", R_Q3),
    ("arr.con.q3", "個人（TrendLife）ARR", 463, "US$M", "2026Q3", "固定匯率；年增 1%；Worry-Free 轉入 TrendLife 後重編", R_Q3),
    ("arr.v1.q3", "Vision One SaaS ARR", 564, "US$M", "2026Q3", "固定匯率；年增 46%；與法說譯稿「近 800M、占企業 ARR 61%」口徑不一致，未解", R_Q3),
    ("arr.v1.q3.call", "Vision One ARR（法說譯稿口徑）", 800, "US$M", "2026Q3", "約數（nearly）；年增 37%；占 TrendAI ARR 約 61%；口徑與新聞稿不同", CALL3),
    ("arr.svc.q3", "服務 ARR", 88, "US$M", "2026Q3", "年增 25%", R_Q3),
    ("arr.dlp.q3", "Digital Life Protection ARR", 141, "US$M", "2026Q3", "年增 37%；占 TrendLife ARR 30%", R_Q3),
    ("nrr.q3", "淨營收留存率（NRR）", 1.19, "比率", "2026Q3", "企業；Q2 122%", R_Q3),
    ("grr.q3", "毛留存率（GRR）", 0.87, "比率", "2026Q3", "企業；法說", CALL3),
    ("arr.v1.g.q3", "Vision One ARR 年增率", 0.46, "比率", "2026Q3", "新聞稿口徑（Q1 50%、Q2 49%）", R_Q3),
    ("arr.plat.fy25", "大型企業平台 ARR（Vision One Platform＋Platform Cyber Services）", 467, "US$M", "2025Q4", "固定匯率；年增 58%", R_Q4),
    ("arr.con.q2", "個人（TrendLife）ARR", 374, "US$M", "2026Q2", "年增 4%（與 Q3 463 口徑不同：重編前後）", R_Q2),
    ("netnew.ent.q3", "企業淨新增 ARR", 31, "US$M", "2026Q3", "Q2 為 21", R_Q3),
    ("pipe.q4", "Q4 以後管線", 1200, "US$M", "2026Q3", "約數；法說", CALL3),
    ("arpu.con.q1", "個人客戶 ARPU", 33.3, "US$", "2026Q1", "年增 2%（Q3 年增 9%）", R_Q1),
):
    f(i, m, v, u, p, b, s, "Interested-party", IV)
f("fy25.ent.arr", "企業 ARR（FY2025 年底）", ">1300", "US$M", "2025Q4", "固定匯率；2025 全年成長 2%", R_Q4, "Interested-party", IV)
f("fy25.con.sales", "個人業務營收年增", -0.09, "比率", "FY2025", "公司說法", R_Q4, "Interested-party", IV)

# ---- 人力、現金、股數、股利 ----
f("emp.fy25", "員工數（期末）", 6717, "人", "FY2025 年底", "合併；R&D 2,228、銷售 1,811、產品支援 1,406、行銷 489、管理 783", AR25 + " p.5", "Verified")
f("emp.fy24", "員工數（期末）", 6869, "人", "FY2024 年底", "合併", AR24 + " p.5", "Verified")
f("emp.q3", "員工數（約）", 6300, "人", "2026Q3 末", "法說；Q4 預計低於 6,200；全年減 300–400 人（譯稿 300 與 400 並見）", CALL3, "Interested-party", IV)
f("sev.q3", "離職相關費用（約）", 1000, "百萬日圓", "2026Q3", "法說約數", CALL3, "Interested-party", IV)
f("sev.q4", "離職相關費用（預計）", 4000, "百萬日圓", "2026Q4", "法說約 40 億日圓、約 300 人", CALL3, "Interested-party", IV)
f("aws.credit.q3", "AWS 抵減（Q3 雲端費用）", 11, "US$M", "2026Q3", "法說（譯稿換算日圓 1.1／1.7／1.8 十億並見，口徑未解）", CALL3, "Interested-party", IV)
f("cash.fy25", "現金及存款", 220092, "百萬日圓", "2025-12-31", "合併資產負債表", AR25 + " p.19", "Verified")
f("cashsec.fy25", "現金及存款＋有價證券", 231030, "百萬日圓", "2025-12-31", "資料商", SA_BS, "Verified")
f("cash.q2", "現金及存款", 219902, "百萬日圓", "2026-06-30", "短信；資料商轉載", GYO + "；" + SA_BS, "Verified")
f("cashsec.q2", "現金及存款＋短期投資", 231363, "百萬日圓", "2026-06-30", "資料商（現金 219,902＋短期投資 11,461）", SA_BS, "Verified")
f("defrev.fy25", "合約負債（遞延收入，流動）", 236085, "百萬日圓", "2025-12-31", "資料商", SA_BS, "Verified")
f("defrev.q2", "合約負債（遞延收入，流動）", 232485, "百萬日圓", "2026-06-30", "資料商", SA_BS, "Verified")
f("recv.fy25", "應收帳款", 76126, "百萬日圓", "2025-12-31", "資料商", SA_BS, "Verified")
f("debt", "有利息負債", "未找到", "百萬日圓", "2026-06-30", "資產負債表資料商頁無借款列", SA_BS, "Assumed")
f("lease", "租賃負債", "未找到", "百萬日圓", "2026-06-30", "資料商頁未列", SA_BS, "Assumed")
f("sh.issued", "已發行股數", 140901604, "股", "2025-12-31", "", AR25 + " p.35", "Verified")
f("sh.treasury", "庫藏股", 10221368, "股", "2025-12-31", "", AR25 + " p.35", "Verified")
f("sh.out.fy25", "流通股數", 130680236, "股", "2025-12-31", "", AR25 + " p.6", "Verified")
f("sh.out.now", "流通股數", 129.93, "百萬股", "2026-11-13", "資料商", SA_Q, "Verified")
f("div.fy25", "每股股利（FY2025 年度）", 185, "日圓", "FY2025", "股利總額 24,175 百萬；除息日 2025-12-29", AR25 + " p.35；" + SA_Q, "Verified")
f("div.fy24", "每股股利（FY2024 年度）", 184, "日圓", "FY2024", "股利總額 24,158 百萬", AR25 + " p.35", "Verified")
for fy, cfo, cx, da, bb in ((2021, 55267, 2156, 20894, 0), (2022, 56903, 1402, 25185, 7035), (2023, 57227, 2178, 28315, 24999), (2024, 46781, 1220, 28927, 40000)):
    f(f"cfo.fy{fy}", "營業現金流", cfo, "百萬日圓", f"FY{fy}", "合併現金流", SA_CF, "Verified")
    f(f"capex.fy{fy}", "資本支出", cx, "百萬日圓", f"FY{fy}", "（FY2023 2,178、FY2024 1,220 與 AR 營運概況一致）", SA_CF, "Verified")
    f(f"da.fy{fy}", "折舊攤銷（含合約成本等攤銷，口徑待釐清）", da, "百萬日圓", f"FY{fy}", "資料商", SA_CF, "Verified")
    f(f"bb.fy{fy}", "庫藏股買回", bb, "百萬日圓", f"FY{fy}", "現金支出", SA_CF, "Verified")
f("rd.fy25", "研究開發費用", 6651, "百萬日圓", "FY2025", "有價證券報告書「研究開發活動」（二手轉載；與員工 R&D 2,228 人不相稱，疑僅計特定資本化／外部項目，口徑待釐清）", EDB, "Verified")
f("sbc", "股票薪酬費用", "未找到", "—", "FY2025", "法定申報僅見董事股票選擇權費用約 7 百萬日圓（FY2024）；合併股票薪酬未單獨揭露", AR24 + " p.13", "Assumed")

# ---- 市場資料 ----
f("px", "收盤價", 7214, "日圓", "2026-11-13", "東證；Q3 財報（11-12 盤後）後 +8.32%（前日收盤 6,660）", SA_Q, "Verified")
f("px.prev", "收盤價", 6660, "日圓", "2026-11-12", "財報當日", CALL3 + "（譯稿標題誤標為美元）；與 11-13 +554 日圓相符", "Verified")
f("mcap", "市值", 937340, "百萬日圓", "2026-11-13", "資料商", SA_Q, "Verified")
f("beta.sa", "Beta（5 年）", 0.43, "—", "2026-11-13", "資料商 stockanalysis", SA_Q, "Verified")
f("rf", "日本 10 年期公債殖利率", 0.0277, "比率", "2026-11-11", "Trading Economics", TE, "Verified")
f("cons.tp", "賣方平均目標價", 6433, "日圓", "2026-09-30 價格日（資料更新 2026-09-11）", "S&P Global 彙整 9 家；中位數 6,400、低 5,000、高 9,000；評等 2 強買、0 買、5 持有、1 賣、1 強賣；財報前，只作歷史對照", SA_F, "Interested-party", IV_SV)
f("cons.rev26", "賣方共識營收", 298960, "百萬日圓", "FY2026", "S&P Global 彙整（家數未載；範圍 282.0–303.6 十億日圓）；財報前", SA_F, "Interested-party", IV_SV)
f("cons.rev27", "賣方共識營收", 316690, "百萬日圓", "FY2027", "S&P Global 彙整；財報前", SA_F, "Interested-party", IV_SV)
f("cons.eps", "賣方共識 EPS（FY2026、FY2027）", "未找到", "日圓", "FY2026／FY2027", "資料商頁 EPS 為 n/a 或需付費", SA_F, "Interested-party", IV_SV)

# ---- 同業（stockanalysis 統計頁；EV 倍數為 TTM 口徑，NTM 未取得）----
PEERS = [("CHKP", "Check Point", 4.12, 28.31, 13.70, 16.51, 15.78, 4.67, 0.42, "2026-11-13"), ("GEN", "Gen Digital", 6.74, 42.80, 8.25, 10.20, 9.20, 4.36, 1.26, "2026-10-26"),
         ("QLYS", "Qualys", 8.61, 34.45, 28.03, 29.48, 28.04, 10.16, 0.60, "2026-10-26"), ("FTNT", "Fortinet", 14.25, 32.91, 33.85, 34.40, 32.40, 11.32, 1.04, "2026-11-13"),
         ("ZS", "Zscaler", 16.68, -3.52, 39.22, None, None, 8.88, 0.84, "2026-11-13"), ("OKTA", "Okta", 10.31, 7.68, 44.83, 125.95, 92.60, 9.67, 0.72, "2026-11-13"),
         ("PANW", "Palo Alto Networks", 17.28, 7.05, 77.10, 324.69, 168.38, 22.88, 0.97, "2026-11-13")]
for tk, nm, g3, om, fpe, evebit, evebitda, evs, beta, dt in PEERS:
    u = f"https://stockanalysis.com/stocks/{tk.lower()}/statistics/"
    f(f"peer.{tk}", f"{nm}（{tk}）", {"g3": g3, "opm": om, "fpe": fpe, "evebit": evebit, "evebitda": evebitda, "evs": evs, "beta": beta}, "混合",
      dt, "3 年營收成長預估（%）、營業利益率（%）、前瞻 P/E、EV/EBIT、EV/EBITDA、EV/Sales（TTM）、Beta；P/E 口徑（GAAP 或 non-GAAP）因資料商而異，未逐家核對", u + "（美國 GAAP；租賃處理未逐家核對）", "Verified")

# ---- AI 相關揭露 ----
for i, m, b, s in (
    ("ai.token", "2026-08-13 下修指引的理由：雲端相關費用（含 AI token 成本）顯著超出原預估", "營業利益指引 56,400 → 44,400 百萬日圓（−12,000，其中含產品功能投資與競爭環境，未拆 token 部分）", R_Q2),
    ("ai.anthropic", "與 Anthropic 策略協議（2026 年初簽署）；內部大量使用 Claude；加入 Project Glasswing，以 Claude Opus 5 做漏洞優先排序與虛擬修補；Claude Compliance API 整合至 Vision One（Q3 延伸至 Claude Code、Cowork）", "金額未揭露", R_Q2 + "；" + CALL3),
    ("ai.nvidia", "NVIDIA：2025-10-28 起代理式 AI 防護；Q3 支援 NVIDIA Agent Safety Platform；以 Nemotron 開源模型作為 Vision One 平台以降低 token 成本", "金額未揭露", R_Q4 + "；" + R_Q3 + "；" + CALL3),
    ("ai.magna", "Magna AI：與 Wistron、NVIDIA 的主權 AI 合資（2025-10 宣布；2026-01 前後設立）；VicOne 與 Magna 合計營收年增 8 倍、ARR 年增 36 倍（Q2）", "金額未揭露；FY2026 營業利益指引含 Magna 啟動前期成本（資料商說法）", R_Q2 + "；https://the-shashi.com/tse/4704/decisions/2025-magna-ai-trendai-pivot/"),
    ("ai.aesir", "ÆSIR AI 安全研究平台（Q1）；Autonomous Vulnerability Remediation 私測（Q3）", "金額未揭露", R_Q1 + "；" + R_Q3),
    ("ai.pkg", "TrendAI Vision One AI Security Package（企業 AI 風險管理）", "未單獨揭露營收或客戶數", R_Q4),
    ("price.raise", "定價：過去曾漲價、可能延續到 2027 年；彈性定價；免費授權方案協助舊產品客戶遷移（無預算者）", "管理層預期對 Q4 與全年指引無重大影響；漲價日期與幅度未找到", CALL3),
    ("ai.units", "AI 產品付費客戶數、AI 營收、AI 訂單", "未找到（公司未揭露）", "—")):
    f(i, m, b, "文字", "2025Q4–2026Q3", "公司揭露／管理層說法", s, "Interested-party" if s != "—" else "Assumed", IV if s != "—" else "")
f("buyout", "私募收購傳聞：2025-02-13 Reuters 報導 Bain Capital、Advent、EQT 競逐，KKR 亦表興趣；公司未置評；報導稱不保證成交", "股價當日漲停 +1,500 日圓至 10,860", "文字", "2025-02-13", "匿名消息人士；2024 年已報導公司探索出售", REUT, "Interested-party", "匿名消息來源與媒體：無法驗證，來源可能有試探市場的誘因")
f("listed", "上市狀態：2026-11-12 仍發布 Q3 財報；2026-11-13 東證收盤 7,214 日圓；EDINET 大量持有申報（2026-10-28 野村 10.40%、2026-07-21 三菱 UFJ 5.01%、ValueAct 3.68% 於 2025-09-19）未見公開收購或私有化申報", "仍上市；不下市；未找到任何已宣告的收購或私有化", "文字", "2026-11-16", "查證", "https://edinetdb.jp/company/E04999/ma；" + R_Q3 + "；" + SA_Q, "Verified")


def main():
    with open(os.path.join(HERE, "data", "facts_trendmicro.csv"), "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["指標", "數值", "單位", "期間", "口徑", "出處", "擷取日", "標記", "來源誘因"])
        for x in F:
            v = json.dumps(x["value"], ensure_ascii=False) if isinstance(x["value"], dict) else x["value"]
            w.writerow([x["metric"], v, x["unit"], x["period"], x["basis"], x["source"], x["date"], x["tag"], x["incentive"]])
    json.dump(F, open(os.path.join(HERE, "data", "facts_trendmicro.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    bad = [x["id"] for x in F if x["tag"] not in ("Verified", "Interested-party", "Analogy", "Assumed", "Derived") or (x["tag"] == "Interested-party" and not x["incentive"])]
    ids = [x["id"] for x in F]
    dup = {i for i in ids if ids.count(i) > 1}
    assert not bad and not dup, (bad, dup)
    # 分部合計核對
    for fy in (2023, 2024, 2025):
        s = sum(next(x["value"] for x in F if x["id"] == f"seg.{r}.fy{fy}") for r in REG)
        t = next(x["value"] for x in F if x["id"] == f"rev.fy{fy}")
        print(f"FY{fy} 分部合計 {s} − 合併 {t} = {s - t}")
    print(f"事實 {len(F)} 筆")


if __name__ == "__main__":
    main()

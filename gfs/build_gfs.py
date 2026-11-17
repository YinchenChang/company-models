# 格芯（GlobalFoundries，GFS；Nasdaq）個股模型 v0.1.0（AI 半導體 Project：代工層，成熟與特殊製程晶圓代工；樣板 umc/、vis/）
# 用法：python3 gfs/build_gfs.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)。金額：十億美元（US$B）；晶圓：千片 300 mm 約當（kWSE）；售價：美元／片；EPS：美元（基本、IFRS）。
# 營收＝製造服務（成熟與特殊製程晶圓：產能 × 稼動率 × 平均售價；不含矽光子）＋矽光子與資料中心（美元）＋技術服務（MIPS、ARC IP 與設計服務 NRE；美元）。
# 本檔上半部為公司設定（數字、文字、來源），下半部為與 umc/build_umc.py 同一套共用建置程式（幣別改美元、匯率恆為 1）。
PR3 = "https://www.sec.gov/Archives/edgar/data/0001709048/000170904826000259/globalfoundries3q2026earni.htm"
PR1 = "https://www.sec.gov/Archives/edgar/data/0001709048/000170904826000111/globalfoundries1q2026earni.htm"
PR2 = "https://www.globenewswire.com/news-release/2026/08/05/3339162/0/en/GlobalFoundries-Reports-Second-Quarter-2026-Financial-Results.html"
MB2 = "https://www.marketbeat.com/instant-alerts/globalfoundries-q2-earnings-call-highlights-2026-08-05/"
CL3 = "https://www.webull.com/news/15866071678360576"
FV3 = "https://finviz.com/news/415654/globalfoundries-misses-the-whisper-by-a-penny-but-accelerating-ai-growth-arrives-with-a-4-billion-capex-bill-and-more-dilution"
FS3 = "https://finsee.ai/earnings/gfs/2026/q3/en/"
W247 = "https://247wallst.com/cards/globalfoundries-q3-2026-earnings-gfs-01m6k7qqa5why13p60hfm4bqc7"
SAI = "https://stockanalysis.com/stocks/gfs/financials/income-statement/"
SACF = "https://stockanalysis.com/stocks/gfs/financials/cash-flow-statement/"
SABS = "https://stockanalysis.com/stocks/gfs/financials/balance-sheet/"
SAR = "https://stockanalysis.com/stocks/gfs/financials/ratios/"
YSN = "docs/workorders/20261117_新增五家_Yahoo基本面快照.md"
CSV1 = "markets/invest_map/data/closes_20261020_20261113.csv"
CSV2 = "markets/invest_map/data/closes_extra_20261113.csv"
C = dict(kind="umc", folder="gfs", name="格芯", wafer_unit="300 mm 約當", wafer_name="製造服務晶圓", wafer_rev_label="製造服務晶圓營收（不含矽光子與資料中心、技術服務）")
C["usd"] = True
C["urlname"] = {PR3: "公司 6-K 3Q26 財報新聞稿（2026-11-04）", PR1: "公司 6-K 1Q26 財報新聞稿", PR2: "公司 2Q26 財報新聞稿（GlobeNewswire 2026-08-05）", MB2: "MarketBeat 2Q26 法說摘要",
                CL3: "3Q26 法說逐字稿（Webull 轉載）", FV3: "Finviz 3Q26 財報評論", FS3: "Finsee 3Q26 財報摘要", W247: "24/7 Wall St 3Q26 財報摘要",
                SAI: "stockanalysis.com 年度損益表", SACF: "stockanalysis.com 年度現金流量表", SABS: "stockanalysis.com 年度資產負債表", SAR: "stockanalysis.com 比率頁"}
C["cap26_src"] = "Q 頁：各季出貨 ÷ 稼動率（2Q26、3Q26 公司稱高 80%；1Q26 假設；4Q26 假設）"
C["cap26_note"] = "公司未揭露產能與稼動率數值；以出貨與稼動率推得（Derived／Assumed）"
C["capg"] = {27: (0.04, "公司：2027 年淨資本支出約 40 億美元（3Q26 法說）；新產能（Malta 矽光子與中介層、Vermont 擴產、Dresden、Singapore）多在 2027 下半年至 2029 年投產，2027 年產能增幅取 4%（Assumed）"),
             28: (0.06, "矽光子產能 2027 年約三倍（法說）屬特殊製程，晶圓當量有限；Malta 中介層 2028 上半年放量（TSMC 協議）；取 6%（Assumed）"), 29: (0.06, "同上；日美合作新廠（Malta，預計 2029 年投產）由 GF 營運收管理與技術費，不併入本公司產能（Assumed）")}
C["ut"] = {27: {"A": 0.92, "B": 0.89, "C": 0.84}, 28: {"A": 0.93, "B": 0.89, "C": 0.82}, 29: {"A": 0.93, "B": 0.90, "C": 0.82}}
C["ut_src"] = "公司：2Q26、3Q26 稼動率「高 80%」（3Q26 法說逐字稿）；2027 年預期隨產能增加仍有改善空間"
C["asp"] = {27: {"A": 0.08, "B": 0.05, "C": 0.02}, 28: {"A": 0.03, "B": 0.02, "C": 0.0}, 29: {"A": 0.02, "B": 0.01, "C": -0.01}}
C["asp_src"] = "公司：3Q26 混合 ASP 季增約 2.5%、年增近 5%；2027 年「混合 ASP 與產品組合」約提升 10%（3Q26 法說）；矽光子已另列，故製造服務售價年增 A 取 8%、B 5%、C 2%（Assumed）"
C["asp26_f"] = "=(a26_rev-ai_26-int_26)/a26_ship*1000000"
C["asp26_src"] = "（FY2026 營收 − 矽光子 − 技術服務）÷ 出貨；美元／片"
C["ai26"] = 0.5
C["ai26_src"] = CL3 + "：矽光子 2026 年營收約 5 億美元；2027 年底年化營收約 10 億美元，產能約增三倍"
C["aig27"] = {"A": 0.8, "B": 0.6, "C": 0.4}
C["ai_src2"] = CL3 + "：矽光子與 SiGe 產能已訂滿至 2027、2028 年部分；公司 2027 年底年化 10 億美元（FY2027 全年約 8 億美元＝+60%）"
C["aig_L"] = 0.35
C["int"] = {27: {"A": 1.10, "B": 1.02, "C": 0.92}, 28: {"A": 1.20, "B": 1.08, "C": 0.92}, 29: {"A": 1.30, "B": 1.14, "C": 0.92}}
C["int26"] = 0.88
C["int_src"] = MB2 + "：技術服務約占營收 11%（2Q26）、約 12%（3Q26 法說稱 2026 年約 12%）；MIPS、ARC 2026 年貢獻 1.0–1.2 億美元；為絕對金額假設（美元）"
C["da26"] = "=0.311+0.307+0.335+0.337"; C["da26_tag"] = "Derived"
C["da26_src"] = "1Q26 3.11、2Q26 3.07、3Q26 3.35（公司現金流量表，十億美元）；4Q26 取 3.37（3Q26 ×1.006，Assumed）"
C["da26_note"] = "第四季折舊攤銷：公司未指引"
C["dag"] = {27: (0.06, "新產能陸續投產、收購無形資產攤銷（2Q26 起 ARC、MIPS）；2026 年折舊約 12.9 億美元，2027 年取 +6%（Assumed）"), 28: (0.10, "2027–28 年大額資本支出（淨資本支出 2027 年約 40 億美元）之後的折舊（Assumed）"), 29: (0.10, "同上")}
C["opx_src"] = "FY2025 營業費用 8.93（stockanalysis）、2026 年前三季 0.271＋0.331＋0.341（公司）、4Q26 指引 0.287"
C["nonop"] = (-0.05, 0.0, 0.05)
C["nonop_src"] = "4Q26 公司指引：淨利息與其他 −5 百萬至 +1 百萬美元／季；FY2027 起取 0（Assumed）"
C["nonop_note"] = "3Q26 其他費用 4.39 億美元含美國政府股份發行 4.41 億（非現金、一次性）：不外推；後續政府股份發行的非現金費用公司指引排除，本模型未計（見 README 待判斷）"
C["tax_src"] = "FY2025 約 2.5%（含一次性利益）；1Q26 43.8%、2Q26 14.8%、3Q26 稅前虧損下的稅費 0.070；4Q26 指引稅費 44–68 百萬美元（有效稅率約 17%）"
C["sh_src"] = "4Q26 指引完全稀釋股數約 5.75 億股；公司：2027 年起約 6.00 億股（美國政府持股約 7%、尚待交割；3Q26 法說）"
C["sh_note"] = "2027 年起股數步階式增加見 In 頁 sh27（Interested-party：公司指引）；之後不變"
C["fx26"] = 1.0
C["fx26_src"] = "美元報表，匯率恆為 1（Verified）"
C["sh27"] = 0.600
C["q_inputs"] = [
    ("q4_rev", "4Q26 營收（公司指引中點）", "US$B", 2.000, "Interested-party", PR3 + "：4Q26 營收 20 億美元 ± 0.25 億", "取中點"),
    ("q4_ship", "4Q26 晶圓出貨季增", "%", 0.03, "Assumed", "公司未指引出貨；2026 年各季年增 6.6–7.6%（Finsee）；4Q25 619 千片、3Q26 647 千片", "取 3%"),
    ("q4_ut", "4Q26 稼動率", "%", 0.89, "Assumed", "公司：2Q26、3Q26 高 80%；2027 年仍有改善空間", "取 89%"),
    ("q4_gm", "4Q26 毛利率（IFRS）", "%", 0.307, "Interested-party", PR3 + "：4Q26 IFRS 毛利率 30.7% ± 100 bps", "取中點"),
    ("q4_opx", "4Q26 營業費用（IFRS）", "US$B", 0.287, "Interested-party", PR3 + "：4Q26 IFRS 營業費用 287 百萬美元 ± 10", "取中點"),
    ("q4_nonop", "4Q26 營業外淨額", "US$B", -0.002, "Interested-party", PR3 + "：淨利息與其他收入（費用）−5 至 +1 百萬美元", "取中點 −2 百萬"),
    ("q4_tax", "4Q26 有效稅率（稅費 ÷ 稅前）", "%", 0.172, "Interested-party", PR3 + "：4Q26 稅費 44–68 百萬美元；中點 56 百萬 ÷ 稅前約 326 百萬", "—"),
    ("q_sh", "加權股數（4Q26 完全稀釋指引）", "十億股", 0.575, "Interested-party", PR3 + "：4Q26 完全稀釋股數約 5.75 億股", "含預期對美國政府的股份發行")]
C["cons"] = [
    ("cons_rev26", "共識營收 FY2026（只作對照，不進模型）", "US$B", 7.323, "Yahoo Finance 2026-11-17 快照（" + YSN + "；第三方，Interested-party）", "2026-12 期末"),
    ("cons_eps26", "共識 EPS FY2026（非 IFRS；只作對照）", "US$", 1.988, "同上", "共識為非 IFRS EPS（排除股權激勵、攤銷與一次性）；本模型為 IFRS EPS，口徑不同"),
    ("cons_rev27", "共識營收 FY2027", "US$B", 8.294, "同上", "2027-12 期末"),
    ("cons_eps27", "共識 EPS FY2027（非 IFRS）", "US$", 2.600, "同上", "口徑同上"),
    ("co_asp27", "公司：2027 年混合平均售價與產品組合約提升 10%", "%", 0.10, CL3, "含產品組合；本模型的製造服務售價不含矽光子，A 取 8%"),
    ("co_sip27", "公司：2027 年底矽光子年化營收約 10 億美元", "US$B", 1.0, CL3, "年化運行率；FY2027 全年應低於此值"),
    ("co_capex27", "公司：2027 年淨資本支出約 40 億美元", "US$B", 4.0, CL3 + "；" + FV3, "約三分之一由客戶預付款支應")]
C["px"] = 41.47; C["px_date"] = "2026-11-13"
C["px_src"] = "收盤價 2026-11-13（" + CSV2 + "：GFS 41.47）"; C["px_src2"] = "工作單第 3 節：價格只取 11-13 收盤兩個 CSV"
C["eps_ttm"] = 0.35; C["eps_ttm_src"] = "近 12 月 IFRS EPS：4Q25 0.36＋1Q26 0.18＋2Q26 0.30＋3Q26 −0.49＝0.35（公司；3Q26 含美國政府股份發行費用 0.79）；P/E 僅供參考"
C["fpe_lab"] = "前瞻 P/E（Yahoo 2026-11-17 快照，非 IFRS EPS 基礎）"; C["fpe"] = 15.16; C["fpe_src"] = YSN
C["cfeps_lab"] = "共識前瞻 12 月 EPS（非 IFRS：1.6／12 × FY2026＋10.4／12 × FY2027）"; C["cfeps"] = "=1.6/12*1.988+10.4/12*2.6"
C["cfeps_src"] = "Yahoo 2026-11-17 快照共識 EPS：FY2026 1.988、FY2027 2.600（非 IFRS）"
C["cons_tp"] = 64.77; C["cons_tp_src"] = "Yahoo 2026-11-17 快照：22 位分析師目標價均值 64.77（第三方，Interested-party）"
C["cons_tp_src2"] = "口徑：共識 EPS 為非 IFRS"
C["nim5"] = "=AVERAGE(1.448/8.108,1.020/7.392,-0.265/6.750,0.885/6.791)"
C["nim5_src"] = "stockanalysis 損益表：FY2022–25 淨利與營收（FY2021 淨損 2.50 億，上市前）；只有四年"
C["roe_f"] = "=AVERAGE(0.1607,0.0964,-0.0238,0.0779)"; C["roe_src"] = "stockanalysis 比率頁：FY2022 16.07%、FY2023 9.64%、FY2024 −2.38%、FY2025 7.79%"
C["q_a1"] = "Q — 實績與 FY2026 錨點（十億美元；EPS 為美元、基本、IFRS；股數為十億股；晶圓為千片 300 mm 約當）。藍字＝輸入（公司 6-K、stockanalysis）；黑字＝公式"
C["q_a2"] = "FY2026 錨點＝1Q、2Q、3Q26 實績（3Q26 財報 2026-11-04 公布）＋4Q26 公司指引（營收 20 億美元、IFRS 毛利率 30.7%、營業費用 2.87 億、稅費 44–68 百萬）；稼動率公司只說「高 80%」，1Q26 與 4Q26 為假設"
C["q3_state"] = "實績"
C["q_s1"] = "公司 6-K：1Q26（2026-05-05）、2Q26（2026-08-05）、3Q26（2026-11-04）新聞稿；FY2025：stockanalysis 年度損益表（轉錄 20-F）"
C["qd"] = dict(
    rev={"B": 6.791, "C": 1.634, "D": 1.786, "E": 1.899, "F": "=q4_rev"}, rev_tag="Verified／Interested-party（4Q 指引）",
    rev_s2="公司 6-K 各季；4Q26＝公司指引中點 20 億美元",
    ship={"B": 2345, "C": 579, "D": 625, "E": 647, "F": "=E5*(1+q4_ship)"}, ship_tag="Verified／Assumed",
    ship_s1="公司：晶圓出貨（300 mm 約當，千片）1Q26 579、2Q26 625、3Q26 647（約 647.3）；FY2025＝543＋581＋602＋619", ship_s2="4Q26 為假設（季增 3%）；FY2025 各季取公司歷季新聞稿（1Q25 543、2Q25 581、3Q25 602、4Q25 619）",
    ut={"C": 0.82, "D": 0.88, "E": 0.88, "F": "=q4_ut"}, ut_tag="Interested-party／Assumed", ut_s1="公司：2Q26、3Q26「高 80%」（取 88%）；1Q26 未揭露（假設 82%）", ut_s2="FY2026 欄＝出貨合計 ÷ 產能合計",
    cap={c: f"={c}5/{c}6" for c in "CDEF"}, cap_s1="出貨 ÷ 稼動率", cap_s2="公司未揭露產能",
    fxq={c: "fx_26" for c in "CDEF"}, asp_s2="含矽光子與技術服務；模型的製造服務售價另扣（In 頁 asp_26）",
    gp={"B": 1.690, "C": 0.451, "D": 0.505, "E": 0.571, "F": "=F4*q4_gm"}, gp_tag="Verified／Interested-party（4Q 指引）",
    gp_s2="IFRS 毛利（公司）；4Q26＝營收 × 30.7%", gm_s2="公司：IFRS 毛利率 1Q26 27.6%、2Q26 28.3%、3Q26 30.1%（非 IFRS 29.0%、29.9%、31.4%）",
    opx={"B": "=B9-B12", "C": "=C9-C12", "D": "=D9-D12", "E": "=E9-E12", "F": "=q4_opx"},
    oi={"B": 0.797, "C": 0.180, "D": 0.174, "E": 0.230, "F": "=F9-F11"}, oi_tag="Verified／Derived", oi_s2="IFRS 營業利益（公司）",
    nonop={"B": "=B14-B12", "C": "=C14-C12", "D": "=D14-D12", "E": "=E14-E12", "F": "=q4_nonop"}, nonop_s2="稅前 − 營業利益＝財務收益淨額＋其他收入（費用）；3Q26 含美國政府股份發行費用 −0.441（非現金）",
    pti={"B": 0.911, "C": "=0.180+0.015-0.010", "D": "=0.174+0.009+0.013", "E": "=0.230+0.009-0.439", "F": "=F12+F13"}, pti_tag="Verified／Derived", pti_s2="營業利益＋財務收益淨額＋其他收入（費用）（公司損益表）；FY2025 取 stockanalysis 稅前 0.911",
    tax={"B": "=B14-B17-B18", "C": "=C14-C17-C18", "D": "=D14-D17-D18", "E": "=E14-E17-E18", "F": "=F14*q4_tax"},
    ni={"B": 0.885, "C": 0.104, "D": 0.167, "E": -0.270, "F": "=F14-F15-F18"}, ni_tag="Verified／Derived", ni_s2="公司淨利（IFRS）；2026 年各季非控制權益僅約 0.5–1 百萬美元，併入淨利不分列",
    mi={"B": 0.003, "C": 0, "D": 0, "E": 0, "F": 0}, mi_s2="FY2025：stockanalysis 少數股東損益 3 百萬；2026 各季未分列（取 0，Assumed）",
    sh={"B": 0.555, "C": 0.555, "D": 0.549, "E": 0.551, "F": "=q_sh"}, sh_s2="基本加權股數（公司損益表）：FY2025 555、1Q26 555、2Q26 549、3Q26 551（百萬股）；4Q26＝完全稀釋指引 575",
    eps_rep={"B": 1.59, "C": 0.19, "D": 0.30, "E": -0.49}, eps_s1="公司：FY2025 基本 1.59（稀釋 1.59）；1Q26 基本 0.19（稀釋 0.18）；2Q26 0.30；3Q26 −0.49", eps_s2="9M26 基本合計 0.00")
C["q_checks"] = [("FY2026 營收年增", "=G4/B4-1", "0.0%"), ("3Q26 營收季增（公司 +6%）", "=E4/D4-1", "0.0%"), ("FY2026 IFRS 毛利率", "=G10", "0.0%"),
                 ("FY2026 營業利益率", "=G12/G4", "0.0%"), ("檢查：4Q26 EPS 自算 − 公司指引中點 0.47（應約為 0）", "=F20-0.47", "0.000"),
                 ("檢查：9M26 營收 自算 − 公司前三季合計 5.319（應約為 0）", "=SUM(C4:E4)-5.319", "0.000"),
                 ("FY2026 EPS 不含 3Q26 美國政府股份發行費用 0.441（非現金；參考；稅額影響為 0）", "=(G17+0.441)/G19", "0.00")]
C["seg"] = dict(title="官方報告分部與終端市場（最近完整年度＋最近四季 4Q25–3Q26；分部原則第 3 條）",
    cols=["FY2025A", "4Q25", "1Q26", "2Q26", "3Q26", "最近四季合計"],
    rows=[("晶圓代工與技術服務（單一報導部門）", {"B": 6.791, "C": 1.830, "D": "=C4", "E": "=D4", "F": "=E4", "G": "=SUM(C36:F36)"}, "#,##0.000", "Verified",
           "公司 6-K 各季合併營收（公司為單一營運部門）", "分部金額以外的部門資料：公司只揭露終端市場比重（見下列）"),
          ("  智慧行動裝置占營收（公司揭露比重）× 營收", {"E": "=E36*0.36", "F": "=F36*0.32", "G": "=SUM(C37:F37)"}, "#,##0.000", "Interested-party",
           MB2 + "：2Q26 約 36%；" + FV3 + "：3Q26 約 32%（Finsee 稱 36% 屬上季）", "FY2025、4Q25、1Q26 比重：未找到；最近四季合計僅含 2Q26、3Q26"),
          ("  通訊基礎設施與資料中心 × 營收", {"E": "=E36*0.16", "F": "=F36*0.17", "G": "=SUM(C38:F38)"}, "#,##0.000", "Interested-party",
           MB2 + "：2Q26 約 16%；" + CL3 + "：3Q26 17%", "其餘期間比重：未找到"),
          ("  其他終端市場（車用、家用與工業物聯網、PC 與其他；差額，僅 2Q26、3Q26）", {"E": "=E36-E37-E38", "F": "=F36-F37-F38", "G": "=SUM(C39:F39)"}, "#,##0.000", "Derived", "差額", "車用與物聯網各約 19%（2Q26，MarketBeat）；3Q26 未揭露"),
          ("檢查：分部合計 − 合併營收", {"B": "=B36-B4", "C": "=C36-1.83", "D": "=D36-C4", "E": "=E36-D4", "F": "=F36-E4"}, "#,##0.000", "Derived", "—", "應為 0"),
          ("檢查：子列合計 − 分部（僅 2Q26、3Q26）", {"E": "=E37+E38+E39-E36", "F": "=F37+F38+F39-F36"}, "#,##0.000", "Derived", "—", "應為 0")])
C["scn_t"] = {"A": "情境 A（樂觀）：稼動率 92%／93%／93%；製造服務售價 +8%／+3%／+2%；矽光子 FY2027 +80%、之後 +35% × 顆數 A；技術服務 1.10／1.20／1.30 十億美元",
              "B": "情境 B（基準）：稼動率 89%／89%／90%；售價 +5%／+2%／+1%；矽光子 +60%、之後 × 顆數 B（0.8／0.71）；技術服務 1.02／1.08／1.14",
              "C": "情境 C（保守）：稼動率 84%／82%／82%；售價 +2%／0／−1%（中國成熟製程擴產與智慧行動裝置需求下滑）；矽光子 +40%、之後 × 顆數 C；技術服務 0.92（持平）"}
C["units_a1"] = "Units — 格芯製造服務晶圓（千片 300 mm 約當）、稼動率、售價與矽光子、技術服務（三情境）。產能 FY2026 由出貨 ÷ 稼動率推得；矽光子 FY2028–29 乘 Siliconomics 同字母顆數係數"
C["units_note"] = "讀法：格芯 EPS 由稼動率、混合售價與資本支出後的折舊決定；矽光子（資料中心光互連）FY2027 年底年化約 10 億美元，在 FY2029 仍低於營收兩成；2027 年淨資本支出約 40 億美元使折舊與自由現金流承壓。"
_SRC = "stockanalysis.com 格芯季度報表與公司 3Q26 6-K"
C["peers"] = [("台積電（TSMC，ADR）", "TSM", "=398.73/(1.6/12*17.5+10.4/12*23.39)", "2026-11-13", "收盤 398.73（" + CSV1 + "）÷ 前瞻 12 月 EPS（Yahoo 2026-11-17 快照：FY2026 17.50、FY2027 23.39；1.6／12、10.4／12 加權）"),
              ("聯電（UMC，ADR）", "UMC", "=15.4/(1.6/12*1.135+10.4/12*0.9)", "2026-11-13", "收盤 15.40 ÷ 前瞻 12 月 EPS（Yahoo：FY2026 1.135、FY2027 0.900；FY2026 含 2Q26 一次性業外）"),
              ("高塔半導體（Tower）", "TSEM", "=146.89/(1.6/12*4.067+10.4/12*7.004)", "2026-11-13", "收盤 146.89 ÷ 前瞻 12 月 EPS（Yahoo：FY2026 4.067、FY2027 7.004）"),
              ("德州儀器（TXN；類比與成熟製程類比對象）", "TXN", "=303.6/(1.6/12*9.056+10.4/12*11.58)", "2026-11-13", "收盤 303.60（" + CSV2 + "）÷ 前瞻 12 月 EPS（Yahoo：FY2026 9.056、FY2027 11.58）"),
              ("中芯國際（SMIC，港股）", "0981.HK", "=48.52/7.8443/(1.6/12*0.188+10.4/12*0.241)", "2026-11-13", "收盤 48.52 港元 ÷ HKD=X 7.8443 ÷ 前瞻 12 月 EPS（Yahoo：FY2026 0.188、FY2027 0.241 美元；快照註記幣別混用）")]
C["peer_note"] = "晶圓代工與成熟製程同業（台積電、聯電、高塔、德州儀器、中芯港股）：收盤價取 2026-11-13，EPS 為 Yahoo 共識（多為非 GAAP；台積電、聯電為 IFRS），格芯 EPS 為 IFRS，口徑不同；同業組合為本模型自訂（工作單未指定）"
C["tp_method"] = ("DCF（標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，以 WACC 折現到 2026 年底（期中慣例），終值＝NOPAT(FY2033) ×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, 美國 10 年期公債殖利率)、RONIC＝WACC；"
                  "ke＝rf＋β ×（ERP 5%＋國家風險溢酬 0），β 為格芯 Yahoo β（單一來源）去槓桿後依市值 D/E 重新加槓桿；EV 加淨現金（含有價證券）÷ 完全稀釋股數 6.00 億股，以 ke 推到 12 個月後並扣期間股利（DCF 頁、tools/fcff_dcf.py）。"
                  "折舊取 FY2023–25 占營收平均；資本支出 FY2027–28 取營收 35%（公司指引 2027 年淨資本支出約 40 億美元，Assumed），FY2029 起收斂到折舊略高水準；營運資金取營收 25%（Assumed）；客戶預付款未另計。情境價值負值取 0。"
                  "Comparable：台積電、聯電、高塔、德州儀器、中芯前瞻 P/E 中位數 × 12 個月後前瞻 EPS（IFRS）；相對倍數調整係數：格芯 Bloomberg 前瞻 P/E 只有 2022–2026 五年，不足 7 年門檻，係數取 1（不調整）。"
                  "主要偏誤：稼動率與售價屬公司層級假設（待併入 Siliconomics）；Comparable 對 IFRS EPS 套非 GAAP 倍數偏保守；FY2030 起只用收斂法。情境對應：A＝樂觀、B＝基準、C＝保守（中國成熟製程擴產、智慧行動裝置下滑）")
C["sens_cases"] = [("cap_adj", "產能係數", "x", "0.00"), ("ut_adj", "稼動率調整", "pt", "0.0%"), ("asp_adj", "售價水準調整", "%", "0.0%"),
                   ("dag_adj", "折舊年增調整", "pt", "0.0%"), ("cc_g", "每片現金成本年增", "%", "0.0%"), ("phi", "營業費用彈性", "x", "0.00"),
                   ("nonop_f", "營業外淨額", "US$B", "0.000"), ("tax", "稅率", "%", "0%"), ("ai_adj", "矽光子營收係數", "x", "0.00"), ("m_new", "新業務邊際貢獻率", "%", "0%")]
C["sens_read"] = "讀法：售價水準 ±3% 影響最大（加權目標價約 ±14%），其次是產能、折舊與每片現金成本（約 ±10%）、稼動率（約 ±6%，固定成本高）；DCF 另受資本支出強度（In 區間 25%／45% 約 +13%／−12%）與完全稀釋股數（0.57／0.62 十億股約 +5%／−3%）影響；矽光子營收 ±30% 約 ±7%。各案例都是一次動一個輸入，不是門檻。"
C["pe_cmc"] = [None, None, None, None, None, None, 20.2, 32.9, None, 21.8]
C["pe_sa"] = [None, None, None, None, None, None, 20.29, 32.87, None, 21.93]
C["pe_note"] = "companiesmarketcap（2022 20.2、2023 32.9、2025 21.8；網址見 Sources 頁）與 stockanalysis 年度 P/E 只有 FY2022（20.29）、FY2023（32.87）、FY2025（21.93）；FY2021、FY2024 為虧損無 P/E。上市（2021-10）以來只有三個有效年度，十年規則僅作參考。"
C["recon"] = [("sec", "A. FY2026–27 錨點 vs 共識（Yahoo 2026-11-17 快照；第三方）"),
    ("FY2026 營收 vs 共識", "=cons_rev26", ("rev", 2026), "Yahoo 快照共識", "4Q26 為公司指引中點", "#,##0.000", False),
    ("FY2026 EPS（IFRS）vs 共識（非 IFRS）", "=cons_eps26", ("eps", 2026), "Yahoo 快照共識", "口徑不同：模型含 3Q26 政府股份發行費用 0.79 與股權激勵", "#,##0.00", False),
    ("FY2027 營收 vs 共識", "=cons_rev27", ("rev", 2027), "Yahoo 快照共識", "共識 2027 年成長約 13%", "#,##0.000", False),
    ("FY2027 EPS（IFRS）vs 共識（非 IFRS）", "=cons_eps27", ("eps", 2027), "Yahoo 快照共識", "口徑不同：IFRS 較非 IFRS 低約股權激勵與攤銷 0.4–0.5 美元", "#,##0.00", False),
    ("4Q26 毛利率 vs 公司指引（百分點）", "=q4_gm", ("=Q!F10", 0), "公司 6-K（2026-11-04）", "三欄同值（FY2026 錨點）", "0.0%", True),
    ("sec", "B. FY2027–29 vs 公司說法（校準未使用）"),
    ("FY2027 製造服務售價年增 vs 公司混合售價與組合約 10%（百分點）", "=co_asp27", ("aspg", 2027), "3Q26 法說逐字稿", "A 取 8%（矽光子已另列）；B、C 低於公司", "0.0%", True),
    ("FY2027 矽光子營收 vs 公司 2027 年底年化 10 億美元", "=co_sip27", ("ai_usd", 2027), "3Q26 法說逐字稿", "全年營收應低於年化運行率；A、B、C 皆低於 10 億美元", "0.000", False),
    ("FY2027 技術服務＋其他（隱含）— 不比較", None, ("int_usd", 2027), "—", "公司只說 2026 年約占營收 12%；無 2027 目標", "0.000", False),
    ("共識前瞻 12 月 EPS（非 IFRS）vs 模型（情境加權 IFRS）", "=cons_feps", ("feps", 0), "Yahoo 快照共識", "口徑不同；三欄同值（情境加權）", "#,##0.00", False)]
C["sources"] = [
    ("公司：6-K 3Q26 財報新聞稿（2026-11-04；損益、資產負債表、現金流量、非 IFRS 調節、4Q26 指引、股利、庫藏股、美國政府股份發行 4.41 億美元非現金費用）", PR3),
    ("公司：6-K 2Q26 財報新聞稿（2026-08-05；損益、資產負債表、現金流量、3Q26 指引）", PR2),
    ("公司：6-K 1Q26 財報新聞稿（2026-05-05；損益、資產負債表、現金流量、庫藏股 4 億美元、2Q26 指引）", PR1),
    ("媒體：MarketBeat 2Q26 法說摘要（終端市場比重、Synopsys ARC 收購、技術服務、2026 年各終端市場成長指引）", MB2),
    ("媒體：Webull 轉載 3Q26 法說逐字稿（2027 年淨資本支出約 40 億美元、稼動率高 80%、混合 ASP 與組合 2027 年約 +10%、矽光子 2026 年約 5 億美元與 2027 年底年化 10 億美元、資料中心占營收 17%、技術服務約 12%、完全稀釋股數 2027 年約 6 億股）", CL3),
    ("媒體：Finviz 3Q26 評論（智慧行動裝置約占 32%、長期毛利率目標 40%、2027 年淨資本支出 40 億美元）", FV3),
    ("媒體：Finsee 3Q26 財報摘要（出貨季增、年增；與逐字稿口徑差異見 README）", FS3), ("媒體：24/7 Wall St 3Q26 財報摘要", W247),
    ("資料商：stockanalysis（FY2021–25 損益、現金流量、資產負債表、年度 P/E、ROE；原文為公司 20-F，轉錄等級）", SAI + "；" + SACF + "；" + SABS + "；" + SAR),
    ("第三方彙整：Yahoo Finance 2026-11-17 基本面快照（共識 EPS、目標價、β、同業 EPS）；只作交叉檢查與 Comparable 同業 EPS", YSN),
    ("資料商：companiesmarketcap 年底 P/E（2022、2023、2025；第三方，Interested-party）", "https://companiesmarketcap.com/globalfoundries/pe-ratio/"),
    ("價格：2026-11-13 收盤（GFS、同業；只取自工作單指定兩個 CSV）", CSV1 + "；" + CSV2),
    ("行業層：Siliconomics 快照（AI 加速器顆數係數 um28／um29；本 repo，唯讀）", "siliconomics/data/CURRENT_SNAPSHOT；siliconomics/reader.py"),
    ("待補（人工）：產能與稼動率數值（公司未揭露）、終端市場完整比重（20-F 與投影片）、技術服務逐季金額、各平台 ASP 與晶圓數、Mubadala 現持股比例、庫藏股後續授權、美國政府股份發行的後續交割時程與費用、日美合作新廠的 GF 收費條件", "未找到"),
    ("SemiAnalysis", "未使用")]

# 標準版 DCF 參數（tools/fcff_dcf.in_rows；docs/plan/個股DCF標準規格.md）。資產負債表 2026-09-30（公司 3Q26 6-K）
BS3 = PR3 + "（資產負債表 2026-09-30）"
C["dcf_P"] = dict(
    date="2026-12-31", ccy="US$", amt="US$B", shu="十億股",
    rf=(0.05, "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；與 mrvl 等美股模型同一來源）",
        "2026-11-13 殖利率未取得（財政部頁面本次無法讀取），沿用 10-23 值", "無風險利率（美國 10 年期公債殖利率，2026-10-23）"),
    beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Derived",
                "格芯 β：Yahoo Finance 2026-11-17 快照 1.85（YSN；單一來源、第三方）；依市值 D/E 去槓桿後 Blume 調整 0.67β＋0.33，再以本公司 D/E 與稅率重新加槓桿（Hamada）；同業 β 未取得",
                "Sens 低／高（Analogy）：β_u ∓ beta_du（0.3），各自同式重新加槓桿"),
    rf_us=(0.05, "同 rf（美元計價，無需換算）", "Sens「美元口徑折現」對本公司等同基準"),
    crp=(0.0, "Assumed", "美國：國家風險溢酬 0", "ERP＝5%（工作單預設）"),
    kd=("=rf+0.01", "Assumed", "美國 10 年期公債＋信用利差 1%（Assumed）", "債務占資本約 6%（債務與租賃 1.66 十億美元 ÷ 市值約 23 十億美元）"),
    tax_w=("=tax", "Derived", "＝In 頁 tax（FY2027 起有效稅率 17%）", "NOPAT 用同一稅率"),
    debt=(1.136, "Verified", BS3 + "：長期債務流動 99、非流動 1,037（百萬美元）", "不含租賃"),
    lease=(0.524, "Verified", BS3 + "：非流動租賃負債 524（百萬美元）", "流動租賃負債併入其他流動負債，3Q26 新聞稿未單列（未找到）"),
    cash=(3.212, "Verified", BS3 + "：現金 997＋流動有價證券 1,258＋非流動有價證券 957（百萬美元）", "非流動有價證券 957 為債務證券（公司揭露「現金、約當現金與有價證券 32 億美元」）"),
    inv=(0.0, "Assumed", "非營業權益投資：未單列（併入其他非流動資產 804 百萬美元）", "不加入評價（保守）"),
    mi=(0.057, "Verified", BS3 + "：非控制權益 57 百萬美元", "帳面值"),
    stub=dict(bs_date="2026-09-30", val_date="2026-12-31", bs_src=BS3,
              div=(0.067, "Verified", "2026-10-09 已付季股利每股 0.12 美元（公司 2Q26 新聞稿；記錄日 2026-09-23），約 5.57 億股 × 0.12＝0.067 十億美元，在資產負債表日之後、評價日之前", "4Q26 庫藏股買回：公司未指引（未計）")),
    dps=(0.48, "Assumed", "季股利每股 0.12 美元（2026-10-27 董事會核准 2027-01-14 發放）；評價日後 12 個月以 4 次計＝0.48 美元", "Assumed 股利政策不變"),
    sh=("=sh27", "Interested-party", "＝In 頁 sh27（公司指引：2027 年起完全稀釋股數約 6.00 億股）", "含預期對美國政府的股份發行（約 7% 持股，尚待交割）；以完全稀釋股數評價，保守"),
    fc_da=(None, "=AVERAGE(1.451/7.392,1.565/6.750,1.301/6.791)", None, "Derived",
           SACF + "：折舊攤銷 FY2023 1,451、FY2024 1,565、FY2025 1,301（百萬美元）；" + "營收 FY2023 7.392、FY2024 6.750、FY2025 6.791（十億美元；stockanalysis 損益表）；比例 19.6%、23.2%、19.2%",
           "PL 推導區塊的銷貨成本內折舊（da_tot26 × da_sh，年增 dag）只用於 EBIT；DCF 加回的折舊用本比例"),
    fc_cx=(0.25, 0.35, 0.45, "Assumed",
           "公司：2027 年淨資本支出約 40 億美元（3Q26 法說；FY2027 營收約 8.3 十億美元，占 48%）；FY2026 約 14 億美元（占 19%）；FY2023–25 平均 14.8%（stockanalysis 現金流量表：1,804／625／722）。FY2027–28 兩年取平均 35%（Assumed）",
           "Sens 用 fc_cx_d ±1pt；低／高 25%／45%；客戶預付款（約三分之一）未另計入營運資金（保守）"),
    fc_cx_ss=("=fc_da_ss*(1+g_tv)*1.1", "Assumed", "FY2029 起線性收斂，FY2032 達到：折舊 ×（1＋g）× 1.1（建置期後的維持加成長性資本支出）", "Andy 可改；原模板為近期比例與維持水準的中點，格芯 2027 年資本支出屬建置期，不適用"),
    fc_wc=(0.15, 0.25, 0.35, "Assumed",
           "營運資金＝應收＋存貨 − 應付：FY2023–25 占營收 31.7%、38.9%、39.2%（stockanalysis；「應收」含預付款與其他，高估）；基準取 25%（Assumed）",
           "Sens 低／高 15%／35%（Assumed）；客戶預付款可抵銷，未計"))
C["dcf_ncbs"] = ("3.212+0-1.136-0.524-0.057", "2026-09-30 資產負債表（十億美元；公司 3Q26 6-K）：現金與有價證券 3.212 − 長期債務 1.136 − 非流動租賃 0.524 − 非控制權益 0.057（流動租賃未單列）")
C["readme"] = [
 ("格芯（GlobalFoundries，GFS）個股模型 v0.1.0（AI 半導體 Project：代工層，成熟與特殊製程晶圓代工；樣板 umc/、vis/）　2026-11-17", True),
 ("版本：v0.1.0（初版）。營收＝製造服務晶圓（產能 × 稼動率 × 平均售價）＋矽光子與資料中心（美元）＋技術服務（MIPS、ARC IP 與設計服務；美元）。滾入 3Q26 實績（2026-11-04 財報）與 4Q26 指引。報表與股價皆為美元，匯率恆為 1。", False),
 ("模型命題：格芯的獲利由成熟與特殊製程的稼動率與混合售價決定；2026 年的轉折是資料中心與矽光子（2026 年約 5 億美元、2027 年底年化 10 億美元）、2027 年混合 ASP 與組合約 +10%，但 2027 年淨資本支出約 40 億美元與完全稀釋股數增至約 6 億股（美國政府持股）會壓低每股自由現金流與 IFRS EPS。", False),
 ("核心問題：稼動率（高 80%）與售價能回升多少；矽光子在 FY2029 補多少營收；2027–28 年資本支出與稀釋後，現價 41.47 美元（2026-11-13）相對 DCF 與同業倍數是高是低。3Q26 IFRS 淨損 2.70 億美元主要來自 4.41 億美元的非現金股份發行費用（發給美國政府 990 萬股），不是營運惡化（營業利益 2.30 億美元）。", False),
 ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。FY2025A 實績、FY2026 錨點（1Q–3Q26 實績＋4Q26 公司指引）、FY2027–FY2029 驅動因子推導、FY2030–FY2032 收斂法。金額：十億美元；EPS：美元（基本、IFRS）；口徑：歸屬普通股淨利。晶圓：千片 300 mm 約當。", False),
 ("", False),
 ("驅動因子 → 推導量（行業變數讀 Siliconomics；成熟與特殊製程稼動率與售價為公司層級，待併入 Siliconomics）", True),
 ("製造服務晶圓：產能（FY2026 由各季出貨 ÷ 稼動率推得約 2,900 千片；FY2027–29 年增 4%／6%／6%）× 稼動率（A／B／C：FY2027 92%／89%／84%、FY2028 93%／89%／82%、FY2029 93%／90%／82%）× 美元售價（FY2026 約 2,300 美元；FY2027 +8%／+5%／+2%、FY2028 +3%／+2%／0、FY2029 +2%／+1%／−1%）；匯率恆為 1。", False),
 ("新業務：矽光子 FY2026 5 億美元（公司）、FY2027 +80%／+60%／+40%、之後年增 35% 再乘 Siliconomics 同字母顆數係數（um28、um29）；技術服務 FY2026 約 8.8 億美元（公司稱約占營收 12%）、FY2027 1.10／1.02／0.92、FY2028 1.20／1.08／0.92、FY2029 1.30／1.14／0.92 十億美元；新業務邊際貢獻率 45%（Assumed）。", False),
 ("成本：銷貨成本內折舊（FY2026 約 11.6 億美元＝全年折舊攤銷 12.9 億 × 0.9；年增 6%／10%／10%）＋每片現金成本（FY2026 殘差，年增 2%）；營業費用彈性 0.5、年增下限 3%；營業外 0（FY2026 含 3Q26 政府股份費用 −4.41 億，不外推）、稅率 17%、股數 FY2027 起 6.00 億股（公司指引）。", False),
 ("校準：FY2026 ＝ 1Q–3Q26 實績（公司 6-K）＋4Q26 公司指引（營收 20 億美元、IFRS 毛利率 30.7%、營業費用 2.87 億、稅費 56 百萬、稀釋股數 5.75 億）。製造服務售價與每片現金成本以 FY2026 校準；不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF（標準版：無槓桿自由現金流 FCFF＋WACC，FY2027–32 明確期＋永續，三情境權重 0.3／0.5／0.2，情境價值負值取 0）＋0.5 × Comparable（台積電、聯電、高塔、德州儀器、中芯前瞻 P/E 中位數 × 12 個月後前瞻 EPS；相對倍數調整係數：格芯 Bloomberg 前瞻 P/E 只有五年、不足門檻，取 1）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m，FY2032 達到；PL 頁灰底）。十年平均淨利率不適用（上市僅四個完整會計年度），B、C 依工作單以 FY2029 × 0.9／0.8 代替。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。公司數字取自 6-K 新聞稿；法說內容取自逐字稿與媒體轉述；券商、資料商、Yahoo 快照為 Interested-party／第三方。SemiAnalysis 未使用。", False),
 ("", False),
 ("已知限制：(1) 公司不揭露產能、稼動率數值與製造服務單價，產能與售價為推得；(2) 稼動率與售價為公司層級假設，Siliconomics 尚無成熟與特殊製程層（建議補進）；(3) 矽光子與技術服務金額為公司口述或推估，毛利率未揭露，邊際貢獻率為 Assumed；(4) 對美國政府後續股份發行的非現金費用（公司指引排除）未建模，股數以公司 6 億股指引取代；(5) 客戶預付款（2027 年淨資本支出約三分之一）未另計入 DCF，偏保守；(6) 共識 EPS 為非 IFRS，模型 EPS 為 IFRS，Comparable 套用同業倍數偏保守；(7) 流動租賃負債與 Mubadala 持股比例未找到；(8) 現價為 2026-11-13 收盤，3Q26 財報後 11-04 起股價反映已含。", False),
 ("", False),
 ("情境對照表（A／B／C 讀同字母的行業情境；公司特有變數 A 最有利、C 最不利）", True),
 ("行業：矽光子營收 FY2028–29 乘 Siliconomics um28／um29（A 1.0／1.0、B 0.8／0.71、C 0.6／0.6）；匯率不適用（美元）。成熟與特殊製程稼動率與價格：Siliconomics 未建，公司層級（待併入 Siliconomics）。不讀 markets/equipment、markets/network。", False),
 ("公司特有：稼動率 FY2027 92%／89%／84%、FY2028 93%／89%／82%、FY2029 93%／90%／82%；售價 FY2027 +8%／+5%／+2%、FY2028 +3%／+2%／0、FY2029 +2%／+1%／−1%（C＝中國成熟製程擴產與智慧行動裝置需求下滑）；矽光子 FY2027 +80%／+60%／+40%；技術服務 FY2027–29 A 1.10／1.20／1.30、B 1.02／1.08／1.14、C 0.92 十億美元。", False)]
C["obs"] = ["", "", ""]

# ======================= 共用建置程式（格芯、世界先進同一套結構；數字與文字在上方 C 設定） =======================
import sys, os, json, re
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.dont_write_bytecode = True      # 不在上游資料夾寫入 __pycache__（上游唯讀）
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
sys.path.insert(0, os.path.join(ROOT, "markets", "scen")); import scen as SCN   # noqa: E402  v0.16：產業情境 A／C 參考值（市場層；B 讀 Siliconomics）
import cycle_position as CP
import fcff_dcf as FD   # v0.2：標準版 DCF（無槓桿自由現金流＋WACC）共用模組
import valuation_rule as VR

KIND = C["kind"]          # "umc"：成熟製程晶圓＋矽光子（先進封裝、矽光子）＋技術服務；"vis"：8 吋晶圓＋VSMC 12 吋
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
SCN_NAME = {"A": "樂觀", "B": "基準", "C": "保守"}
SILV = f"Siliconomics {SIL.VERSION}（siliconomics/data/{SIL.FILE}）"
TBD = "待併入 Siliconomics"

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = (f"In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億美元（US$B）；晶圓：千片（{C['wafer_unit']}）；售價：美元／片；匯率：美元／美元。"
            "公司財報與法說為 Interested-party（公司說法）或 Verified（財報數字）；券商、資料商、媒體為 Interested-party")
IN["A2"] = ("標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 以情境或 Sens 區間呈現，不當點估計。"
            f"行業變數讀 {SILV}；成熟製程稼動率與售價屬公司層級，說明欄標「{TBD}」。SemiAnalysis 未使用")
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
URLNAME = C["urlname"]

sec(f"A. 行業變數（{SILV}；A／B／C 讀同字母上游，不自訂行業路徑）")
fl, fb, fh = 1.0, 1.0, 1.0
inp("fx", "匯率（美元報表，恆為 1）", "US$/US$", fl, fb, fh, "Verified", "美元報表與美元股價", "無匯率風險")
for s in (SC if C.get("use_um", True) else ()):
    for y in (28, 29):
        k = f"um{y}_{s}"
        inp(k, f"情境 {s}：CY20{y} AI 加速器顆數 ÷ 基準", "x", None, SCN.base(k), None, SCN.V[k]["tag"], (f"{SILV}｜{k}" if s == "B" else f"{SCN.SRC}｜{k}"),
            "矽光子營收 CY2028–29 依同字母顆數係數縮放（行業情境；A＝原始顆數）" if (s == "A" and y == 28) else "同上")

sec(f"B. 產能與稼動率（{C['wafer_unit']}；公司層級，{TBD}：Siliconomics 尚無成熟製程稼動率）")
inp("cap_26", f"FY2026 產能（{C['wafer_unit']}，千片／年）", "千片", None, "=Q!G7", None, "Derived", C["cap26_src"], C["cap26_note"])
for y, (v, s1) in C["capg"].items():
    inp(f"capg{y}", f"FY20{y} 產能年增", "%", None, v, None, "Assumed", s1, f"三情境相同（產能為公司已核准的擴產；{TBD}）")
inp("cap_adj", "產能係數（Sens 用；FY2027 起各年同乘）", "x", 0.95, 1.0, 1.05, "Assumed", "無", "±5%：擴產時程提前或延後")
for y in (27, 28, 29):
    for s in SC:
        inp(f"ut{y}_{s}", f"FY20{y} 平均稼動率（情境 {s}）", "%", None, C["ut"][y][s], None, "Assumed",
            C["ut_src"] if (s == "A" and y == 27) else "同上", (f"A 最有利、C 最不利（C＝中國成熟製程擴產的價格與份額壓力）；{TBD}" if s == "A" else TBD))
inp("ut_adj", "稼動率調整（Sens 用；FY2027 起各年加減，上限 98%）", "pt", -0.03, 0.0, 0.03, "Assumed", "無", "±3 個百分點")

sec(f"C. 平均售價（美元，每片 {C['wafer_unit']}；公司層級，{TBD}：Siliconomics 尚無成熟製程價格）")
inp("asp_26", f"FY2026 平均售價（{C['wafer_name']}；美元／片）", "US$/片", None, C["asp26_f"], None, "Derived", C["asp26_src"], "校準年；FY2027 起依情境年增")
for y in (27, 28, 29):
    for s in SC:
        inp(f"asp{y}_{s}", f"FY20{y} 平均售價年增（美元；情境 {s}）", "%", None, C["asp"][y][s], None, "Assumed",
            C["asp_src"] if (s == "A" and y == 27) else "同上", (f"A＝漲價延續；C＝中國成熟製程擴產造成價格壓力；{TBD}" if s == "A" else TBD))
inp("asp_adj", "售價水準調整（Sens 用；FY2027 起各年同乘 1＋x）", "%", -0.03, 0.0, 0.03, "Assumed", "無", "±3%")

if KIND == "umc":
    sec("D. 新業務：矽光子（先進封裝矽中介層、矽光子）與 技術服務合作（美元營收；毛利以邊際貢獻率計）")
    inp("ai_26", "FY2026 矽光子營收（先進封裝、矽光子）", "US$B", None, C["ai26"], None, "Interested-party", C["ai26_src"], "公司未揭露金額與毛利率；本模型把它從晶圓營收拆出")
    for s in SC:
        inp(f"aig27_{s}", f"FY2027 矽光子營收年增（情境 {s}）", "%", None, C["aig27"][s], None, "Assumed",
            C["ai_src2"] if s == "A" else "同上", "A 最有利、C 最不利" if s == "A" else "—")
    inp("aig_L", "FY2028–29 矽光子營收年增（情境 A 路徑；B、C 再乘同字母顆數係數）", "%", None, C["aig_L"], None, "Assumed", C["ai_src2"], "情境 A 到 FY2029 約 US$1.0B（公司目標）")
    inp("ai_adj", "矽光子營收係數（Sens 用；FY2027 起）", "x", 0.7, 1.0, 1.3, "Assumed", "無", "±30%")
    inp("m_new", "新業務（矽光子、技術服務）邊際貢獻率（毛利 ÷ 營收）", "%", 0.35, 0.45, 0.55, "Assumed", "公司未揭露；先進封裝中介層、矽光子屬特殊製程", "成熟製程晶圓的現金成本以 FY2026 校準，新業務另計")
    inp("int_26", "FY2026 技術服務營收（MIPS、ARC IP 與設計服務；公司稱約占營收 12%）", "US$B", None, C["int26"], None, "Interested-party", C["int_src"], "FY2026 校準值（Derived：12% × FY2026 營收約 7.3 十億美元）")
    for y in (27, 28, 29):
        for s in SC:
            inp(f"int{y}_{s}", f"FY20{y} 技術服務營收（情境 {s}）", "US$B", None, C["int"][y][s], None, "Assumed",
                C["int_src"] if (s == "A" and y == 27) else "同上", "金額為假設；A 最有利、C 最不利" if s == "A" else "—")
else:
    sec("D. VSMC（新加坡 12 吋廠，與 NXP 合資，世界先進持股 60%、合併報表）：產能 × 稼動率 × 售價；毛利率路徑；非控制權益 40%")
    for y, (v, s1) in C["vcap"].items():
        inp(f"vcap{y}", f"FY20{y} VSMC 產能（12 吋，千片／年）", "千片", None, v, None, "Assumed", s1, "三情境相同（產能為公司規劃）")
    inp("vcap_adj", "VSMC 產能係數（Sens 用）", "x", 0.8, 1.0, 1.2, "Assumed", "無", "±20%：量產爬坡快慢")
    for s in SC:
        inp(f"vut_{s}", f"VSMC 稼動率（FY2027 起；情境 {s}）", "%", None, C["vut"][s], None, "Assumed",
            C["vut_src"] if s == "A" else "同上", "A 最有利、C 最不利" if s == "A" else "—")
    inp("vasp", "VSMC 平均售價（美元／12 吋片）", "US$", 1500, C["vasp"], 2100, "Assumed", C["vasp_src"], "130–40 奈米特殊製程 12 吋；公司未揭露")
    for y in (27, 28, 29):
        for s in SC:
            inp(f"vgm{y}_{s}", f"FY20{y} VSMC 毛利率（含折舊；情境 {s}）", "%", None, C["vgm"][y][s], None, "Assumed",
                C["vgm_src"] if (s == "A" and y == 27) else "同上", "量產初期折舊重；A 最有利、C 最不利" if s == "A" else "—")
    inp("vmi", "VSMC 非控制權益比例（NXP 持股）", "%", None, 0.4, None, "Verified", C["vmi_src"], "非控制權益＝VSMC 毛利 ×（1 − 稅率）× 40%（簡化：VSMC 營業費用未另計）")
    inp("ai_sh26", "FY2026 矽光子 ÷ 營收（記憶列）", "%", None, C["ai_sh26"], None, "Interested-party", C["ai_sh_src"], "公司：全年雙位數；只作記憶列，不進營收加總")
    for s in SC:
        inp(f"ai_sh27_{s}", f"FY2027 起 矽光子 ÷ 營收（記憶列；情境 {s}）", "%", None, C["ai_sh27"][s], None, "Assumed",
            C["ai_sh_src2"] if s == "A" else "同上", "AI 需求已反映在稼動率與售價情境；本列不進營收加總" if s == "A" else "—")

sec("E. 成本：折舊（銷貨成本內）、每片現金成本、營業費用、營業外、稅、股數")
inp("da_tot26", "FY2026 折舊與攤銷（全公司）", "US$B", None, C["da26"], None, C["da26_tag"], C["da26_src"], C["da26_note"])
inp("da_sh", "折舊中計入銷貨成本的比例", "%", None, 0.9, None, "Assumed", "公司未揭露拆分；晶圓廠折舊多數計入銷貨成本", "—")
for y, (v, s1) in C["dag"].items():
    inp(f"dag_{y}", f"FY20{y} 銷貨成本內折舊年增", "%", None, v, None, "Assumed", s1, "三情境相同（擴產已核准）")
inp("dag_adj", "折舊年增調整（Sens 用；FY2027 起各年加減）", "pt", -0.05, 0.0, 0.05, "Assumed", "無", "±5 個百分點")
inp("cc_g", "每片現金成本年增（FY2027 起）", "%", 0.0, 0.02, 0.04, "Assumed", "原物料、電價、人力（Assumed）", "每片現金成本為 FY2026 殘差（Derived），之後依本列成長")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.3, 0.5, 0.8, "Assumed", C["opx_src"], "—")
inp("g_opx_min", "營業費用年增下限", "%", None, 0.03, None, "Assumed", "研發、新業務投入", "—")
inp("nonop_f", "營業外淨額（FY2027 起每年）", "US$B", C["nonop"][0], C["nonop"][1], C["nonop"][2], "Assumed", C["nonop_src"], C["nonop_note"])
inp("tax", "有效稅率（FY2027 起）", "%", 0.13, 0.17, 0.21, "Assumed", C["tax_src"], "—")
inp("sh27", "FY2027 起完全稀釋股數（公司指引約 6.00 億股）", "十億股", 0.57, C["sh27"], 0.62, "Interested-party", "公司 3Q26 法說：完全稀釋股數 2027 年起約 6 億股（美國政府持股約 7%、尚待交割）", "Sens 低／高：5.70／6.20 億股")
inp("sh_g", "加權股數年變動", "%", None, 0.0, None, "Assumed", C["sh_src"], C["sh_note"])

sec("F. FY2026 錨點（1Q、2Q26 實績＋3Q26＋4Q26 指引或假設）")
inp("fx_26", "FY2026 平均匯率", "US$/US$", None, C["fx26"], None, "Assumed", C["fx26_src"], "實際平均匯率：未找到")
for k, lab, unit, v, tag, s1, s2 in C["q_inputs"]:
    inp(k, lab, unit, None, v, None, tag, s1, s2)

sec("G. 公司說法、共識與券商（只作 Recon 對帳，不作校準）")
for k, lab, unit, v, s1, s2 in C["cons"]:
    inp(k, lab, unit, None, v, None, "Interested-party", s1, s2)

sec("H. 價格與市場倍數")
inp("px", f"股價（{C['px_date']} 收盤）", "US$", None, C["px"], None, "Interested-party", C["px_src"], C["px_src2"])
inp("pe_ttm", f"近 12 月 P/E（現價 ÷ 近 12 月 EPS {C['eps_ttm']}）", "x", None, f"=px/{C['eps_ttm']}", None, "Derived", C["eps_ttm_src"], "—")
inp("fpe", C["fpe_lab"], "x", None, C["fpe"], None, "Interested-party", C["fpe_src"], "—")
inp("cons_feps", C["cfeps_lab"], "US$", None, C["cfeps"], None, "Derived", C["cfeps_src"], "共識前瞻 EPS 的直接數字：未找到")
inp("cons_tp", "市場共識目標價（賣方平均）", "US$", None, C["cons_tp"], None, "Interested-party", C["cons_tp_src"], C["cons_tp_src2"])
inp("nim_5", "近五年平均淨利率（FY2021–25；對照用）", "%", None, C["nim5"], None, "Derived", C["nim5_src"], "十年平均（2016–2025）：未找到（2016–2020 未取得）；收斂法依工作單改用最後一年 × 0.8")

sec("I. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 樂觀／B 基準／C 保守）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_g", "PL 延伸期營收年增的收斂目標（情境 B；FY2032；v0.2 起不再用於 DCF，DCF 永續成長率見 M2 節 g_tv）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "v0.2 刪除 tp_r、tp_roe（舊股東現金流法 DCF 參數）；本值只決定 FY2030–FY2032 營收收斂路徑")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（2026-10-25 11:53 定 DCF／Comparable 各半）", "—")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（FY2029 與「FY2029 × 0.8」的中點）", "%", None, 0, None, "Assumed", "十年平均未找到，依工作單以最後一年 × 0.8 代替", "—")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（FY2029 × 0.8）", "%", None, 0, None, "Assumed", "同上", "—")
for _row in FD.in_rows(C["dcf_P"]):
    rows.append(_row)
inp("beta_u", "產業無槓桿 β（自算 格芯 Yahoo β，Blume 調整後；採用值）", "x", None, "=0.67*beta_u_med+0.33", None, "Derived",
    "自算產業 β：格芯 Yahoo β 同業 13 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 臺灣加權指數 ^TWII 迴歸，Hamada 去槓桿、現金修正）；＝0.67 × 中位數 beta_u_med ＋ 0.33", "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；用現金修正值：淨現金另計，β 只對營業資產")
inp("beta_u_med", "格芯無槓桿 β（Yahoo β 1.85 去槓桿；單一來源，Blume 調整前）", "x", None, "=1.85/(1+(1-tax_w)*beta_de)", None, "Derived",
    "自算產業 β：格芯 Yahoo β 同業 13 家中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 臺灣加權指數 ^TWII 迴歸，Hamada 去槓桿、現金修正）", "只用於計算 beta_u；第二輪修正起 β 敏感度改 β_u ± beta_du（原為中位數 × 0.8 再 Blume／中位數未調整）")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務本金 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 股數 sh_v)", "與 WACC 權重同一口徑")
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3", "取代原低＝中位數 × 0.8 再 Blume、高＝中位數未調整")
TP_KEYS += ["tp_g", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]

r = 4; INROW = {}
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, s1, s2 = srow
    assert key not in INROW, key
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): IN.cell(r, j).font = BLUE
    for u_, nm_ in URLNAME.items():          # 網址只列在 Sources 頁；In 頁寫來源名稱
        s1 = s1.replace(u_, nm_); s2 = s2.replace(u_, nm_)
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    INROW[key] = r; r += 1
for k in TP_KEYS:
    for c in range(1, 10): IN.cell(INROW[k], c).fill = KEY
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
FD.yellow(IN, FD.JUDG + ("beta_u",))
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績與 FY2026 錨點（US$B） ----------------
Q = wb.create_sheet("Q")
Q["A1"] = C["q_a1"]
Q["A2"] = C["q_a2"]
hdr(Q, 3, ["項目", "FY2025A", "1Q26（實績）", "2Q26（實績）", "3Q26（" + C["q3_state"] + "）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1", "來源2／說明"])
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=C["q_s1"], s2="—", ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"G{r}"] = ann or f"=SUM(C{r}:F{r})"; Q[f"G{r}"].number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, s1); Q.cell(r, 10, s2)
QD = C["qd"]
qrow(4, "營收", QD["rev"], tag=QD["rev_tag"], s2=QD["rev_s2"])
qrow(5, f"晶圓出貨（千片，{C['wafer_unit']}）", QD["ship"], fmt="#,##0.0", tag=QD["ship_tag"], s1=QD["ship_s1"], s2=QD["ship_s2"])
qrow(6, "稼動率", QD["ut"], fmt="0.0%", ann="=G5/G7", tag=QD["ut_tag"], s1=QD["ut_s1"], s2=QD["ut_s2"])
qrow(7, f"產能（千片，{C['wafer_unit']}）", QD["cap"], fmt="#,##0.0", tag="Derived", s1=QD["cap_s1"], s2=QD["cap_s2"])
qrow(8, "平均售價（全部營收 ÷ 出貨；美元／片）", {c: f"={c}4/{c}5/{QD['fxq'][c]}*1000000" for c in "CDEF"}, fmt="#,##0", ann="=G4/G5/fx_26*1000000", tag="Derived", s1="營收 ÷ 出貨 ÷ 匯率", s2=QD["asp_s2"])
qrow(9, "毛利", QD["gp"], tag=QD["gp_tag"], s2=QD["gp_s2"])
qrow(10, "毛利率", {c: f"={c}9/{c}4" for c in "BCDEFG"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2=QD["gm_s2"])
qrow(11, "營業費用", QD["opx"], tag="Derived／Assumed", s2="毛利 − 營業利益")
qrow(12, "營業利益", QD["oi"], tag=QD["oi_tag"], s2=QD["oi_s2"])
qrow(13, "營業外淨額", QD["nonop"], tag="Derived／Assumed", s2=QD["nonop_s2"])
qrow(14, "稅前利益", QD["pti"], tag=QD["pti_tag"], s2=QD["pti_s2"])
qrow(15, "所得稅", QD["tax"], tag="Derived／Assumed", s2="稅前 −（歸屬母公司淨利＋非控制權益）")
qrow(16, "  有效稅率", {c: f"={c}15/{c}14" for c in "BCDEFG"}, fmt="0.0%", ann="", tag="Derived", s1="—")
qrow(17, "歸屬母公司淨利", QD["ni"], tag=QD["ni_tag"], s2=QD["ni_s2"])
qrow(18, "  非控制權益", QD["mi"], tag="Derived／Assumed", s2=QD["mi_s2"])
qrow(19, "加權股數（十億股）", QD["sh"], fmt="0.00000", ann="=AVERAGE(C19:F19)", tag="Derived", s1="淨利 ÷ EPS（公司）", s2=QD["sh_s2"])
qrow(20, "EPS（基本）", {c: f"={c}17/{c}19" for c in "BCDEF"}, fmt="#,##0.00", ann="=G17/G19", tag="Derived", s1="淨利 ÷ 股數")
qrow(21, "  公司揭露 EPS", QD["eps_rep"], fmt="#,##0.00", ann="", tag="Verified", s1=QD["eps_s1"], s2=QD["eps_s2"])
Q.cell(23, 1, "FY2026 錨點的檢查（參考；校準未使用券商數字）").font = BOLD
for i, (lab, f, fm) in enumerate(C["q_checks"], 24):
    Q.cell(i, 1, lab); Q[f"B{i}"] = f; Q[f"B{i}"].number_format = fm
for nm, rr in (("rev", 4), ("gp", 9), ("opx", 11), ("oi", 12), ("nonop", 13), ("tax", 15), ("ni", 17), ("mi", 18), ("sh", 19), ("eps", 20)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$G${rr}")
defname("a26_ship", "Q!$G$5")
# 官方分部與子列（分部原則第 3 條；check_segments 讀本區）
SEG = C["seg"]
Q.cell(34, 1, SEG["title"]).font = BOLD
hdr(Q, 35, ["項目"] + SEG["cols"])
Q.cell(35, 8, "標記").font = BOLD; Q.cell(35, 9, "來源1").font = BOLD; Q.cell(35, 10, "來源2／說明").font = BOLD
for _r, (_lab, _vals, _fmt, _tag, _s1, _s2) in enumerate(SEG["rows"], 36):
    Q.cell(_r, 1, _lab)
    for col, v in _vals.items():
        Q[f"{col}{_r}"] = v; Q[f"{col}{_r}"].number_format = _fmt
        if not (isinstance(v, str) and v.startswith("=")): Q[f"{col}{_r}"].font = BLUE
    Q.cell(_r, 8, _tag); Q.cell(_r, 9, _s1); Q.cell(_r, 10, _s2)
Q.column_dimensions["A"].width = 56
for c in "BCDEFG": Q.column_dimensions[c].width = 14
Q.column_dimensions["H"].width = 30; Q.column_dimensions["I"].width = 50; Q.column_dimensions["J"].width = 70
Q.freeze_panes = "B4"

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G", 2030: "H", 2031: "I", 2032: "J"}
FADE = (2030, 2031, 2032); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
BU = (2026, 2027, 2028, 2029); F3 = (2027, 2028, 2029)
DEFP = dict(scn="A", fx="fx", cap_adj="cap_adj", ut_adj="ut_adj", asp_adj="asp_adj", dag_adj="dag_adj", cc_g="cc_g", phi="phi",
            nonop_f="nonop_f", tax="tax", ai_adj="ai_adj", m_new="m_new", vcap_adj="vcap_adj", vasp="vasp", fm_self=False, wm=None)
FM_FACTOR = {"A": "1", "B": "0.9", "C": "0.8"}
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.00", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        if note: ws.cell(r, 11, note)
        R[key] = r; r += 1
    def sec_(t):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
    yy = lambda y: y % 100
    sec_(f"A. {C['wafer_name']}：產能 × 稼動率 × 平均售價（D 欄＝FY2026 校準年）")
    row("cap", f"產能（{C['wafer_unit']}）", "千片", {"D": "=cap_26", "E": f"=D{r}*(1+capg27)*{P['cap_adj']}", "F": f"=E{r}*(1+capg28)", "G": f"=F{r}*(1+capg29)"}, "#,##0")
    row("ut", "平均稼動率", "%", {"D": "=a26_ship/cap_26", **{FC[y]: f"=MIN(0.98,ut{yy(y)}_{s}+{P['ut_adj']})" for y in F3}}, "0.0%")
    row("ship", f"晶圓出貨（{C['wafer_unit']}）", "千片", {FC[y]: f"={FC[y]}{R['cap']}*{FC[y]}{R['ut']}" for y in BU}, "#,##0")
    row("asp", f"平均售價（{C['wafer_name']}）", "US$/片", {"D": "=asp_26", "E": f"=D{r}*(1+asp27_{s})*(1+{P['asp_adj']})", "F": f"=E{r}*(1+asp28_{s})", "G": f"=F{r}*(1+asp29_{s})"}, "#,##0")
    row("aspg", "平均售價年增（美元）", "%", {FC[y]: f"={FC[y]}{R['asp']}/{FC[y - 1]}{R['asp']}-1" for y in F3}, "0.0%")
    row("fxr", "匯率", "US$/US$", {"D": "=fx_26", **{FC[y]: f"={P['fx']}" for y in F3}})
    row("wrev", C["wafer_rev_label"], "US$B", {FC[y]: f"={FC[y]}{R['ship']}*{FC[y]}{R['asp']}*{FC[y]}{R['fxr']}/1000000" for y in BU})
    if KIND == "umc":
        sec_("B. 新業務：矽光子（先進封裝矽中介層、矽光子）與 技術服務（美元營收）")
        if P["wm"]:   # v0.13：Wattonomics 錨定（矽光子營收 FY2027 起 × wta；FY2028 起先除去前一年係數再乘當年，避免連乘）
            w_ = P["wm"]
            row("ai_usd", "矽光子營收（美元；FY2028 起 × 同字母顆數係數；FY2027 起 × Wattonomics 錨定係數）", "US$B", {"D": "=ai_26", "E": f"=D{r}*(1+aig27_{s})*{P['ai_adj']}*{w_[2027]}",
                "F": f"=E{r}/{w_[2027]}*(1+aig_L)*um28_{s}*{w_[2028]}", "G": f"=F{r}/um28_{s}/{w_[2028]}*(1+aig_L)*um29_{s}*{w_[2029]}"}, "0.000")
        else:
            row("ai_usd", "矽光子營收（美元；FY2028 起 × 同字母顆數係數）", "US$B", {"D": "=ai_26", "E": f"=D{r}*(1+aig27_{s})*{P['ai_adj']}",
                "F": f"=E{r}*(1+aig_L)*um28_{s}", "G": f"=F{r}/um28_{s}*(1+aig_L)*um29_{s}"}, "0.000")
        row("ai_rev", "矽光子營收", "US$B", {FC[y]: f"={FC[y]}{R['ai_usd']}*{FC[y]}{R['fxr']}" for y in BU})
        row("int_usd", "技術服務營收（美元）", "US$B", {"D": "=int_26", "E": f"=int27_{s}", "F": f"=int28_{s}", "G": f"=int29_{s}"}, "0.000")
        row("int_rev", "技術服務營收", "US$B", {FC[y]: f"={FC[y]}{R['int_usd']}*{FC[y]}{R['fxr']}" for y in BU})
        row("rev", "營收", "US$B", {"C": "=a25_rev", **{FC[y]: f"={FC[y]}{R['wrev']}+{FC[y]}{R['ai_rev']}+{FC[y]}{R['int_rev']}" for y in BU}})
        row("aish", "矽光子 ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['ai_rev']}/{FC[y]}{R['rev']}" for y in BU}, "0.0%")
    else:
        sec_("B. VSMC（新加坡 12 吋廠；FY2027 第一季量產）")
        row("vcap", "VSMC 產能（12 吋）", "千片", {"D": "=0", **{FC[y]: f"=vcap{yy(y)}*{P['vcap_adj']}" for y in F3}}, "#,##0")
        row("vut", "VSMC 稼動率", "%", {FC[y]: f"=vut_{s}" for y in F3}, "0.0%")
        row("vship", "VSMC 出貨（12 吋）", "千片", {"D": "=0", **{FC[y]: f"={FC[y]}{R['vcap']}*{FC[y]}{R['vut']}" for y in F3}}, "#,##0")
        row("vrev", "VSMC 營收（12 吋）", "US$B", {"D": "=0", **{FC[y]: f"={FC[y]}{R['vship']}*{P['vasp']}*{FC[y]}{R['fxr']}/1000000" for y in F3}})
        row("rev", "營收", "US$B", {"C": "=a25_rev", **{FC[y]: f"={FC[y]}{R['wrev']}+{FC[y]}{R['vrev']}" for y in BU}})
        row("v12sh", "12 吋（VSMC）÷ 營收", "%", {FC[y]: f"={FC[y]}{R['vrev']}/{FC[y]}{R['rev']}" for y in BU}, "0.0%")
        row("aish", "矽光子 ÷ 營收（記憶列，不進營收加總）", "%", {"D": "=ai_sh26", **{FC[y]: f"=ai_sh27_{s}" for y in F3}}, "0.0%")
    row("g", "營收年增", "%", {FC[y]: f"={FC[y]}{R['rev']}/{FC[y - 1]}{R['rev']}-1" for y in FYS[1:]}, "0.0%")
    sec_("C. 成本與損益")
    row("da", "銷貨成本內折舊", "US$B", {"D": "=da_tot26*da_sh", **{FC[y]: f"={FC[y - 1]}{r}*(1+dag_{yy(y)}+{P['dag_adj']})" for y in F3}})
    if KIND == "umc":
        cal = f"=(D{R['rev']}-a26_gp-D{R['da']}-(1-{P['m_new']})*(D{R['ai_rev']}+D{R['int_rev']}))/D{R['ship']}*1000"
    else:
        cal = f"=(D{R['wrev']}-a26_gp-D{R['da']})/D{R['ship']}*1000"
    row("ccw", "每片現金成本（FY2026 殘差，之後年增 cc_g）", "US$K/片", {"D": cal, **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['cc_g']})" for y in F3}}, "#,##0.00",
        note="Derived：讓 FY2026 毛利等於錨點；含原物料、人力、電力等變動與固定現金成本")
    row("cogs", "現金成本合計（每片 × 出貨）", "US$B", {FC[y]: f"={FC[y]}{R['ccw']}*{FC[y]}{R['ship']}/1000" for y in BU})
    if KIND == "umc":
        row("gpn", "新業務毛利（邊際貢獻率 × 矽光子與 Intel 營收）", "US$B", {FC[y]: f"={P['m_new']}*({FC[y]}{R['ai_rev']}+{FC[y]}{R['int_rev']})" for y in BU})
    else:
        row("vgm", "VSMC 毛利率", "%", {FC[y]: f"=vgm{yy(y)}_{s}" for y in F3}, "0.0%")
        row("gpn", "VSMC 毛利", "US$B", {"D": "=0", **{FC[y]: f"={FC[y]}{R['vrev']}*{FC[y]}{R['vgm']}" for y in F3}})
    row("gp", "毛利", "US$B", {"C": "=a25_gp", **{FC[y]: f"={FC[y]}{R['wrev']}-{FC[y]}{R['da']}-{FC[y]}{R['cogs']}+{FC[y]}{R['gpn']}" for y in BU}})
    row("gm", "毛利率", "%", {FC[y]: f"={FC[y]}{R['gp']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("opx", "營業費用", "US$B", {"C": "=a25_opx", "D": "=a26_opx", **{FC[y]: f"={FC[y - 1]}{r}*(1+MAX(g_opx_min,{P['phi']}*{FC[y]}{R['g']}))" for y in F3}})
    row("oi", "營業利益", "US$B", {FC[y]: f"={FC[y]}{R['gp']}-{FC[y]}{R['opx']}" for y in FYS})
    row("oim", "營業利益率", "%", {FC[y]: f"={FC[y]}{R['oi']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("nonop", "營業外淨額", "US$B", {"C": "=a25_nonop", "D": "=a26_nonop", **{FC[y]: f"={P['nonop_f']}" for y in F3}})
    row("pti", "稅前利益", "US$B", {FC[y]: f"={FC[y]}{R['oi']}+{FC[y]}{R['nonop']}" for y in FYS})
    row("tax", "所得稅", "US$B", {"C": "=a25_tax", "D": "=a26_tax", **{FC[y]: f"={FC[y]}{R['pti']}*{P['tax']}" for y in F3}})
    if KIND == "umc":
        row("mi", "非控制權益", "US$B", {"C": "=a25_mi", "D": "=a26_mi", **{FC[y]: "=0" for y in F3}}, note="FY2027 起假設為 0（歷史上接近 0）")
    else:
        row("mi", "非控制權益（NXP 占 VSMC 40%）", "US$B", {"C": "=a25_mi", "D": "=a26_mi", **{FC[y]: f"=vmi*{FC[y]}{R['gpn']}*(1-{P['tax']})" for y in F3}},
            note="VSMC 虧損時為負值（母公司淨利少扣）")
    row("ni", "歸屬母公司淨利", "US$B", {FC[y]: f"={FC[y]}{R['pti']}-{FC[y]}{R['tax']}-{FC[y]}{R['mi']}" for y in FYS})
    row("shr", "加權股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", **{FC[y]: (f"=sh27" if y == 2027 else f"={FC[y - 1]}{r}*(1+sh_g)") for y in F3}}, "0.00000")
    row("eps", "EPS（基本）", "US$", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['shr']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {FC[y]: f"={FC[y]}{R['eps']}/{FC[y - 1]}{R['eps']}-1" for y in FYS[1:]}, "0.0%")
    row("nim", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    fg = f"fade_g{s}"
    fm = f"$G${R['nim']}*{FM_FACTOR[s]}" if P["fm_self"] else f"fade_m{s}"
    for y in FADE:      # FY2030–32：收斂法（上游只到 2029；非驅動因子推導）
        c, pc, k = FC[y], FC[y - 1], y - 2029
        ws[f"{c}{R['g']}"] = f"=$G{R['g']}+({fg}-$G{R['g']})*{k}/3"
        ws[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
        ws[f"{c}{R['nim']}"] = f"=$G{R['nim']}+({fm}-$G{R['nim']})*{k}/3"
        ws[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
        ws[f"{c}{R['shr']}"] = f"={pc}{R['shr']}*(1+sh_g)"
        ws[f"{c}{R['eps']}"] = f"={c}{R['ni']}/{c}{R['shr']}"
        ws[f"{c}{R['epsg']}"] = f"={c}{R['eps']}/{pc}{R['eps']}-1"
        for key, fmt_ in (("g", "0.0%"), ("rev", "#,##0.00"), ("nim", "0.00%"), ("ni", "#,##0.00"), ("shr", "0.00000"), ("eps", "#,##0.00"), ("epsg", "0.0%")):
            ws[f"{c}{R[key]}"].number_format = fmt_; ws[f"{c}{R[key]}"].fill = FADEFILL
    ws.cell(R["rev"], 11, "FY2030 起為收斂法：營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m（FY2032 達到），非驅動因子推導")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = (f"PL — {C['name']}損益（十億美元；EPS 美元、基本；歸屬母公司）。FY2025A 實績、FY2026 錨點（三情境相同），FY2027–FY2029 由驅動因子推導，"
            "FY2030–FY2032 為收斂法。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要")
PB = {}; r0 = 60
for s in SC:
    # v0.2.4 三層協調 C4 另一側讀數建置（tools/coord/alt_run.py；SIL.COORD_ALT）：情境 B、C 於 CY2027 起 矽光子營收 × wta（＝另一側讀數比值，只在對帳中年份；其他年份＝1）；情境 A 不縮放。正式建置不乘（wm=None，公式與前版相同）
    PB[s] = engine(PL, r0, C["scn_t"][s], scn=s, wm=({2027: "wta_27", 2028: "wta_28", 2029: "wta_29"} if (SIL.COORD_ALT and KIND == "umc" and s in ("B", "C")) else None))
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
sr = 4; SUMR = {}
SUMK = [("rev", "營收", "US$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("wrev", "  " + C["wafer_rev_label"], "US$B", "#,##0.0")]
SUMK += ([("ai_rev", "  矽光子營收", "US$B", "#,##0.0"), ("int_rev", "  技術服務營收", "US$B", "#,##0.0")] if KIND == "umc"
         else [("vrev", "  VSMC 營收（12 吋）", "US$B", "#,##0.0"), ("aish", "  矽光子 ÷ 營收（記憶列）", "%", "0.0%")])
SUMK += [("ut", "稼動率", "%", "0.0%"), ("asp", "平均售價（美元）", "US$/片", "#,##0"),
         ("gm", "毛利率", "%", "0.00%"), ("oim", "營業利益率", "%", "0.00%"), ("ni", "歸屬母公司淨利", "US$B", "#,##0.0"), ("nim", "  淨利率", "%", "0.00%"),
         ("eps", "EPS", "US$", "#,##0.00"), ("epsg", "  EPS 年增", "%", "0.0%")]
LONGK = ("rev", "g", "ni", "nim", "eps", "epsg")
NOC = {"ut", "asp", "wrev", "ai_rev", "int_rev", "vrev", "aish"}   # FY2025A 無驅動因子拆分
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in SUMK:
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS + (FADE if key in LONGK else ()):
            if key in ("g", "epsg") and y == 2025: continue
            if key in NOC and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
            if y in FADE: PL[f"{FC[y]}{sr}"].fill = FADEFILL
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
IN.cell(INROW["fade_mA"], 5, f"=PL!G{SUMR[('A', 'nim')]}")
IN.cell(INROW["fade_mB"], 5, f"=PL!G{SUMR[('B', 'nim')]}*0.9")
IN.cell(INROW["fade_mC"], 5, f"=PL!G{SUMR[('C', 'nim')]}*0.8")
for k in ("fade_mA", "fade_mB", "fade_mC"): IN.cell(INROW[k], 5).font = Font(color="000000")
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型 − 錨點").font = BOLD; sr += 1
CHK = []
for lab, f in (("  三情境：營收差（US$B）", "+".join(f"ABS(D{PB[s]['rev']}-a26_rev)" for s in SC)),
               ("  三情境：毛利差（US$B）", "+".join(f"ABS(D{PB[s]['gp']}-a26_gp)" for s in SC)),
               ("  三情境：EPS 差（US$）", "+".join(f"ABS(D{PB[s]['eps']}-a26_eps)" for s in SC))):
    PL.cell(sr, 1, lab); PL[f"D{sr}"] = "=" + f; PL[f"D{sr}"].number_format = "0.0000"; CHK.append(sr); sr += 1
assert sr < 60, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 10; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 13
PL.cell(2, 1, "灰底＝FY2030–FY2032 收斂法（上游顆數只到 CY2029）：營收年增線性收斂到 fade_g、淨利率線性收斂到 fade_m（In 頁 I 節），非驅動因子推導")

# ---------------- Units ----------------
UN = wb.create_sheet("Units")
UN["A1"] = C["units_a1"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "FY2026", "FY2027", "FY2028", "FY2029", "說明"])
ur = 4; UR = {}
UN.cell(ur, 1, f"1. {C['wafer_name']}：三情境稼動率、出貨與售價").font = BOLD; ur += 1
for key, lab, unit, fmt in (("ut", "稼動率", "%", "0.0%"), ("ship", "晶圓出貨", "千片", "#,##0"), ("asp", "平均售價（美元）", "US$/片", "#,##0"), ("wrev", "晶圓營收", "US$B", "#,##0.0")):
    for s in SC:
        UN.cell(ur, 1, f"{lab}（情境 {s}）"); UN.cell(ur, 2, unit); UR[f"{key}_{s}"] = ur
        for y in BU:
            UN[f"{cyc(y)}{ur}"] = f"=PL!{FC[y]}{PB[s][key]}"; UN[f"{cyc(y)}{ur}"].number_format = fmt
        UN.cell(ur, 7, "PL 頁各情境推導區塊" if s == "A" else "同上"); ur += 1
ur += 1
UN.cell(ur, 1, "2. " + ("新業務：矽光子與 技術服務（三情境）" if KIND == "umc" else "VSMC 12 吋廠與 矽光子（三情境）")).font = BOLD; ur += 1
NEWK = (("ai_usd", "矽光子營收（美元）", "US$B", "0.000"), ("aish", "矽光子 ÷ 營收", "%", "0.0%"), ("int_usd", "技術服務營收（美元）", "US$B", "0.000")) if KIND == "umc" else \
       (("vship", "VSMC 出貨（12 吋）", "千片", "#,##0"), ("vrev", "VSMC 營收", "US$B", "#,##0.0"), ("v12sh", "12 吋（VSMC）÷ 營收", "%", "0.0%"), ("aish", "矽光子 ÷ 營收（記憶列）", "%", "0.0%"))
for key, lab, unit, fmt in NEWK:
    for s in SC:
        UN.cell(ur, 1, f"{lab}（情境 {s}）"); UN.cell(ur, 2, unit); UR[f"{key}_{s}"] = ur
        for y in BU:
            UN[f"{cyc(y)}{ur}"] = f"=PL!{FC[y]}{PB[s][key]}"; UN[f"{cyc(y)}{ur}"].number_format = fmt
        UN.cell(ur, 7, "PL 頁各情境推導區塊" if s == "A" else "同上"); ur += 1
ur += 1
UN.cell(ur, 1, "3. 校準值（只用 FY2026 已發生的實績與錨點；不用共識、不用 FY2027 起的目標）").font = BOLD; ur += 1
for key, lab, f, fmt, note in (("asp26", "FY2026 平均售價（美元／片）", f"=PL!D{PB['A']['asp']}", "#,##0", "晶圓營收 ÷ 出貨 ÷ 匯率"),
                               ("ccw", "FY2026 每片現金成本（US$K）", f"=PL!D{PB['A']['ccw']}", "#,##0.00", "讓 FY2026 毛利等於錨點"),
                               ("da", "FY2026 銷貨成本內折舊（US$B）", f"=PL!D{PB['A']['da']}", "#,##0.00", "全公司折舊 × da_sh")):
    UN.cell(ur, 1, lab); UN[f"C{ur}"] = f; UN[f"C{ur}"].number_format = fmt; UN.cell(ur, 7, note); UR[key] = ur; ur += 1
ur += 1
UN.cell(ur, 1, C["units_note"]).font = BOLD
UN_NOTE = ur
UN.column_dimensions["A"].width = 56; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 13
UN.column_dimensions["G"].width = 56

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（美元）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，年底落後 P/E）只作參考"
V["A2"] = f"循環位置：{C['name']}不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照創意、聯發科；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
O = 51
V.cell(O, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, O + 1, ["項目", "A", "B", "C", "說明"])
V.cell(O + 2, 1, f"現價（{C['px_date']} 收盤）"); V[f"B{O + 2}"] = "=px"; V[f"B{O + 2}"].number_format = "#,##0.00"
for k, y, lab in ((3, 2026, "現價 ÷ FY2026 EPS（錨點）"), (4, 2027, "現價 ÷ FY2027 EPS（模型）"), (5, 2028, "現價 ÷ FY2028 EPS（模型）"), (6, 2029, "現價 ÷ FY2029 EPS（模型）")):
    V.cell(O + k, 1, lab)
    for j, s in enumerate(SC):
        c = "BCD"[j]; V[f"{c}{O + k}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{O + k}"].number_format = "0.0"
V.cell(O + 7, 1, "未來 12 月 EPS（模型 B）≈ 0.2 × FY2026＋0.8 × FY2027"); V[f"B{O + 7}"] = f"=0.2*PL!D{SUMR[('B', 'eps')]}+0.8*PL!E{SUMR[('B', 'eps')]}"; V[f"B{O + 7}"].number_format = "0.00"
V.cell(O + 8, 1, "  現價 ÷ 未來 12 月 EPS（模型 B）"); V[f"B{O + 8}"] = f"=px/B{O + 7}"; V[f"B{O + 8}"].number_format = "0.0"
V.cell(O + 9, 1, "近 12 月 P/E（現價 ÷ 近 12 月 EPS）"); V[f"B{O + 9}"] = "=pe_ttm"; V[f"B{O + 9}"].number_format = "0.0"
V.cell(O + 10, 1, C["fpe_lab"]); V[f"B{O + 10}"] = "=fpe"; V[f"B{O + 10}"].number_format = "0.0"
V[f"E{O + 3}"] = "FY2026 三情境相同（錨點）"; V[f"E{O + 4}"] = "FY2027 三情境差在稼動率、售價與新業務；顆數情境係數從 CY2028 開始"
CY0 = O + 12
V.cell(CY0, 1, "2. 循環位置（本模型營收最高年；FY(t)＝CY(t)）").font = BOLD
hdr(V, CY0 + 1, ["情境", "營收最高年（FY2026–FY2029）", "FY2027 判定", "規則倍數（P/E）", "見頂年是否未知", "說明"])
CYR = {}
for i, s in enumerate(SC, CY0 + 2):
    rr = f"PL!D{PB[s]['rev']}:G{PB[s]['rev']}"
    V.cell(i, 1, s); CYR[s] = i
    V.cell(i, 2, f"=INDEX({{2026,2027,2028,2029}},MATCH(MAX({rr}),{rr},0))")
    V.cell(i, 3, '=IF(B{0}=2029,"中段",'.format(i) + CP.excel_rule("2027", f"B{i}") + ")")
    V.cell(i, 4, f'=IF(C{i}="見頂",pe_peak,IF(C{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 5, f'=IF(B{i}=2029,"是：營收到觀測期末仍在增加","否")')
    V.cell(i, 6, "tools/cycle_position.py excel_rule；營收最高年取代設備市場見頂年" if s == "A" else "同上")
R30 = CY0 + 6
V.cell(R30, 1, "3. 規則值隱含股價（美元；約 FY2027 年底，未折現）＝ FY2027 EPS × 規則倍數").font = BOLD
hdr(V, R30 + 1, ["情境", "FY2027 EPS", "（同左）", "規則倍數", "隱含股價", "相對現價", "現價 ÷ FY2027 EPS"])
VAL = {}
for j, s in enumerate(SC):
    i = R30 + 2 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 3, f"=B{i}").number_format = "0.00"
    V.cell(i, 4, f"=D{CYR[s]}").number_format = "0.0"
    V.cell(i, 5, f"=C{i}*D{i}").number_format = "#,##0.00"
    V.cell(i, 6, f"=E{i}/px-1").number_format = "0%"
    V.cell(i, 7, f"=px/C{i}").number_format = "0.0"
    for c in range(1, 8): V.cell(i, c).fill = KEY
    VAL[s] = i
W0 = R30 + 6
V.cell(W0, 1, "4. What must be true：FY2027 年底價 ＝ 現價 ×（1＋報酬）所需的倍數（FY2028 EPS；約 2027 年底的價格看下一年 EPS）").font = BOLD
hdr(V, W0 + 1, ["FY2027 年底價 ÷ 現價", "所需倍數（÷ FY2028 EPS，A）", "所需倍數（÷ FY2028 EPS，B）", "所需倍數（÷ FY2028 EPS，C）", "對照：十年中位數"])
vr = W0 + 2
for x in (1.0, 1.1, 1.21, 1.5, 2.0):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    for j, s in enumerate(SC):
        V.cell(vr, 2 + j, f"=px*A{vr}/PL!F{SUMR[(s, 'eps')]}").number_format = "0.0"
    V.cell(vr, 5, "=pe_mid").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數低於歷史中位數時，現價相對便宜；倍數是否合理，另看前瞻 P/E 歷史（Bloomberg）。本版未做先驗，無 Andy 合理 P/E 對照。")
WMT_NOTE = vr
T10 = vr + 2
V.cell(T10, 1, "5. 十年 P/E（2016–2025 年底；兩來源平均，分位排名用；無負值或 >200 需剔除）").font = BOLD
hdr(V, T10 + 1, ["年底", "兩來源平均 P/E", "companiesmarketcap", "stockanalysis"])
RB0 = T10 + 16
for i, y in enumerate(range(2016, 2026)):
    rr = T10 + 2 + i
    V.cell(rr, 1, str(y))
    V.cell(rr, 2, f'=IF(COUNT(C{rr}:D{rr})=0,"",AVERAGE(C{rr}:D{rr}))').number_format = "0.00"
    V.cell(rr, 3, f'=IF(B{RB0 + 5 + i}="","",B{RB0 + 5 + i})').number_format = "0.00"
    V.cell(rr, 4, f'=IF(C{RB0 + 5 + i}="","",C{RB0 + 5 + i})').number_format = "0.00"
PE10 = f"$B${T10 + 2}:$B${T10 + 11}"
V.cell(T10 + 12, 1, "年數"); V[f"B{T10 + 12}"] = f"=COUNT({PE10})"
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "companiesmarketcap（年底 P/E）", "vals": C["pe_cmc"], "tag": "Interested-party（資料商；2016–2025 共 10 年）"},
    {"label": "stockanalysis（年度 P/E）", "vals": C["pe_sa"], "tag": "Interested-party（資料商；只有 FY2021–25，2016–2020 未找到）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位，兩個來源分別取分位再平均。" + C["pe_note"])
assert V.cell(RB0 + 5, 1).value == "2016" and V.cell(RB0 + 14, 1).value == "2025", (V.cell(RB0 + 5, 1).value, V.cell(RB0 + 14, 1).value)
rule_rows = [rw for rw in range(RB0, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")

# 0. 目標價（Andy 2026-10-25 11:28／11:53／12:22：DCF 與 Comparable 各半；DCF 負值取 0；列標籤依工作單逐字）
YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
V.cell(4, 1, "0. 目標價").font = BOLD
T = {}
def tset(r, key, label, fs, fmt, note=""):
    V.cell(r, 1, label); T[key] = r
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{r}"] = f; V[f"{c}{r}"].number_format = fmt
    if note: V[f"E{r}"] = note
# v0.2：標準版 DCF（FCFF＋WACC；tools/fcff_dcf.py）。DCF 明細在 DCF 頁；EBIT：FY2027–29＝推導區塊營業利益；FY2030–32（收斂法只有淨利）由淨利還原
DCFW = wb.create_sheet("DCF")
DCFW["A1"] = ("DCF — 標準版（模組 v2）：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例）；"
              "終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py；金額：十億美元")
DCFW["A2"] = ("EBIT：FY2027–FY2029 取各推導區塊「營業利益」；FY2030–FY2032（收斂法只有淨利）＝（歸屬母公司淨利 ×（1＋FY2029 非控制權益 ÷ 淨利））÷（1 − 稅率）− 營業外淨額。"
              "稅率＝In 頁 tax。營業外收支不進 NOPAT，現金與投資另在評價時加入；非控制權益以帳面扣除。折舊、資本支出、營運資金：近三年（FY2023–FY2025）占營收平均，資本支出 FY2029 起線性收斂到 fc_cx_ss。")
DCF_NEXT = [4]
def dpaths(blocks, pre, tax="tax", nonop="nonop_f", lab=None):
    out = []
    for s_ in SC:
        bl = blocks[s_]
        ebit = {y: (f"{pre}{FC[y]}{bl['oi']}" if y <= 2029 else
                    f"=({pre}{FC[y]}{bl['ni']}*(1+{pre}$G${bl['mi']}/{pre}$G${bl['ni']}))/(1-{tax})-{nonop}") for y in YRS6}
        out.append(dict(name=s_, label=(lab or "情境") + f" {s_}（{SCN_NAME[s_]}）", rev={y: f"{pre}{FC[y]}{bl['rev']}" for y in YRS6}, rev0=f"{pre}D{bl['rev']}",
                        ebit=ebit, tax={y: tax for y in YRS6}, ebit_note="FY2030 起由淨利還原", stub_fcff="stub_fcff4q"))   # v0.2.2：stub＝Q 頁 4Q26 FCFF（三情境相同）
    return out
def ddetail(blocks, pre, title, **kw):
    R_ = FD.write_detail(DCFW, DCF_NEXT[0], dpaths(blocks, pre, **kw), list(YRS6), 1.0, conv_from=2029, title=title, fmt="#,##0.00", stub=True)
    DCF_NEXT[0] = R_["end"] + 1
    return R_
def v12ref(R_):
    return {s_: f"DCF!$C${R_['paths'][s_]['v12']}" for s_ in SC}
RD_ = ddetail(PB, "PL!", "三情境 FCFF 明細（十億美元；會計年度＝日曆年；評價日 2026-12-31）")
DCF_NEXT[0] = FD.write_checks(DCFW, DCF_NEXT[0], nc_bs=C["dcf_ncbs"][0], nc_note=C["dcf_ncbs"][1]) + 1
# v0.2.2（第三輪修正第 2 點）：stub 期（2026-09-30 → 12-31）FCFF＝Q 頁 4Q26 季度數（三情境相同；v0.2.1 誤用 FY2027 FCFF × stub_frac）
_q0 = DCF_NEXT[0]
DCFW.cell(_q0, 1, "Stub 期 FCFF：Q 頁 4Q26（十億美元；2026-09-30 → 12-31 恰為一季；三情境相同）").font = BOLD
for _i, (_a, _f, _n) in enumerate((("營收 4Q26", "=Q!F4", "Q 頁（3Q26 × 出貨與售價指引）"), ("營業利益 4Q26", "=Q!F12", "Q 頁（毛利率指引 − 營業費用）"),
                                    ("NOPAT＝營業利益 ×（1 − q4_tax）", f"=C{_q0 + 2}*(1-q4_tax)", "In 頁 q4_tax"),
                                    ("＋折舊攤銷（營收 × fc_da）", f"=C{_q0 + 1}*fc_da", "同明細比例"), ("− 資本支出（營收 × fc_cx）", f"=C{_q0 + 1}*fc_cx", "同明細比例（FY2023–25 平均）"),
                                    ("− 營運資金增加（fc_wc × 年化營收增量 4 ×（4Q26 − 3Q26））", f"=fc_wc*4*(C{_q0 + 1}-Q!E4)", "年化季營收變動"),
                                    ("4Q26 FCFF（＝stub 期 FCFF）", f"=C{_q0 + 3}+C{_q0 + 4}-C{_q0 + 5}-C{_q0 + 6}", "進淨現金滾到評價日")), 1):
    DCFW.cell(_q0 + _i, 1, _a); DCFW.cell(_q0 + _i, 2, _n); DCFW.cell(_q0 + _i, 3, _f).number_format = "#,##0.000"
wb.defined_names["stub_fcff4q"] = DefinedName("stub_fcff4q", attr_text=f"DCF!$C${_q0 + 7}")
DCF_NEXT[0] = _q0 + 9
DCFW.column_dimensions["A"].width = 58; DCFW.column_dimensions["B"].width = 34
for _c in "CDEFGH": DCFW.column_dimensions[_c].width = 12
SDCF = FD.write_summary(wb, V, 5, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt="#,##0.00",
                        title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；明細見 DCF 頁）", beta_note="自算產業 β 重新加槓桿")
for _j, _h in enumerate(["項目", "A（樂觀）", "B（基準）", "C（保守）", "說明"], 1): V.cell(SDCF["h"], _j, _h)
T.update(dcfv=SDCF["dcfv"], dcf12=SDCF["v12"], dcftp=SDCF["dcf"])
V.cell(SDCF["stub"], 5, '="資產負債表日 "&bs_date&"；stub 期 FCFF＝Q 頁 4Q26 季度數（DCF 頁下方；三情境相同）"')
assert SDCF["dcf"] <= 24
_PE = lambda c, y: f"PL!{FC[y]}{SUMR[(dict(zip('BCD', SC))[c], 'eps')]}"
EPSR = {}   # v0.2：舊 0.1 節 EPS 列已移除（ROWMAP 相容）
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = C["peers"]
for i, (nm, tk, pe, dt, src) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
assert PEER_LAST <= 32
tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", C["peer_note"])
tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=" + "+".join(f"tp_w{dict(zip('BCD', SC))[c]}*(2/12*{_PE(c, 2027)}+10/12*{_PE(c, 2028)})" for c in "BCD"), None, None], "#,##0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], "#,##0.00", "同業中位數 × 前瞻 EPS")
V.cell(37, 1, "0.3 結論").font = BOLD
hdr(V, 38, ["項目", "數值", "", "", "說明"])
tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B35", None, None], "#,##0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable（相對倍數調整後，見第 0.5 節）")
tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], "#,##0.00")
tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], "#,##0.00", C["cons_tp_src"])
tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", f"現價＝In 頁 px（{C['px_date']} 收盤）")
tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS 見 In 頁 cons_feps")
tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
tset(46, "band", "價格帶（DCF 保守～樂觀）", [f'=TEXT(D{T["dcf12"]},"#,##0")&"～"&TEXT(B{T["dcf12"]},"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=IF(px/{_PE(c, 2028)}>=MAX({PE10}),1,IF(px/{_PE(c, 2028)}<=MIN({PE10}),0,PERCENTRANK({PE10},px/{_PE(c, 2028)})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 5 節兩來源平均）的分位；高於十年最大值記為 100%、低於最小值記為 0%；只作檢查，不進目標價")
V.cell(48, 1, "目標價方法")
V["B48"] = C["tp_method"]
for r_ in (T["dcftp"], 35, 39):
    V[f"B{r_}"].fill = KEY
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = ("Sens — 一次動一個輸入（In 頁敏感度低／高值），三情境都在本頁下方逐案重算（FY2026 恆等於錨點，每片現金成本在每個案例內重算），"
            "表內為情境加權後的 DCF、Comparable（相對倍數調整後）與加權目標價；另列 FY2027、FY2028 情境 A 的 EPS")
SE["A2"] = "相對基準＝案例加權目標價 ÷ 基準加權目標價 − 1。DCF 欄為標準版 DCF（FCFF＋WACC；每個案例三情境各在 DCF 頁另寫一個 FCFF 明細區塊）；DCF 參數（WACC、β、ERP、g、RONIC、資本支出、營運資金等）的敏感度見本頁最下方「目標價參數敏感度」表"
hdr(SE, 3, ["案例", "設定", "鍵", "FY2027 EPS（A）", "FY2028 EPS（A）", "FY2028 EPS 相對基準（A）", "DCF 目標價", "Comparable（調整後）目標價", "加權目標價", "相對基準"])
cases = C["sens_cases"]
SENS_ROWS = {}
def put_case(i, name, setting, keys, blocks, pre, dref, w="tp_w_dcf", scale=None):
    """blocks：{情境: 推導區塊列號 dict}；pre：區塊所在工作表前綴（PL! 或 空字串＝本頁）"""
    SE.cell(i, 1, name); SE.cell(i, 2, setting); SE.cell(i, 3, keys)
    _sc = f"*{scale}" if scale else ""
    SE[f"D{i}"] = f"={pre}E{blocks['A']['eps']}{_sc}"; SE[f"E{i}"] = f"={pre}F{blocks['A']['eps']}{_sc}"
    SE[f"F{i}"] = f"=E{i}/E$4-1"
    SE[f"G{i}"] = "=" + "+".join(f"tp_w{s}*{dref[s]}" for s in SC)
    fe = "+".join(f"tp_w{s}*(2/12*{pre}E{blocks[s]['eps']}+10/12*{pre}F{blocks[s]['eps']})" for s in SC)
    SE[f"H{i}"] = f"=Val!$B$33*({fe})*rm_factor{_sc}"
    SE[f"I{i}"] = f"={w}*G{i}+(1-{w})*H{i}"
    SE[f"J{i}"] = f"=I{i}/I$4-1"
    for c, fm in (("D", "#,##0.00"), ("E", "#,##0.00"), ("F", "0.0%"), ("G", "#,##0.00"), ("H", "#,##0.00"), ("I", "#,##0.00"), ("J", "0.0%")):
        SE[f"{c}{i}"].number_format = fm
put_case(4, "基準（三情境加權）", "PL 頁三個推導區塊", "—", PB, "PL!", v12ref(RD_))
# v0.2：舊 tp_r／tp_g／tp_w_dcf 六案改為標準版 DCF 參數案例（WACC、β、ERP、g、RONIC、資本支出、營運資金、權益投資、終值替代法、DCF 權重；閉式公式與 DCF 頁明細同式，tools/fcff_dcf.value_expr）
_KE = lambda beta="beta", erp="erp": f"(rf+{beta}*({erp}+crp))"
_WK = lambda k: f"(w_e*{k}+(1-w_e)*kd_at+wacc_adj)"
FDC = [("WACC −1pt", "WACC −1pt（RONIC 隨 WACC）", "wacc_adj", dict(w="(wacc+wacc_adj_lo)"), None),
       ("WACC ＋1pt", "WACC ＋1pt（RONIC 隨 WACC）", "wacc_adj", dict(w="(wacc+wacc_adj_hi)"), None),
       ("β：低", '="β "&TEXT(beta_lo,"0.00")&"（基準 "&TEXT(beta,"0.00")&"；β_u "&TEXT(beta_u,"0.000")&" − "&TEXT(beta_du,"0.0")&" 重新加槓桿）"', "beta", dict(ke=_KE(beta="beta_lo"), w=_WK(_KE(beta="beta_lo"))), None),
       ("β：高", '="β "&TEXT(beta_hi,"0.00")&"（基準 "&TEXT(beta,"0.00")&"；β_u ＋ "&TEXT(beta_du,"0.0")&" 重新加槓桿）"', "beta", dict(ke=_KE(beta="beta_hi"), w=_WK(_KE(beta="beta_hi"))), None),
       ("ERP：低", '="ERP "&TEXT(erp_lo,"0%")&"（基準 "&TEXT(erp,"0%")&"）"', "erp", dict(ke=_KE(erp="erp_lo"), w=_WK(_KE(erp="erp_lo"))), None),
       ("ERP：高", '="ERP "&TEXT(erp_hi,"0%")&"（基準 "&TEXT(erp,"0%")&"）"', "erp", dict(ke=_KE(erp="erp_hi"), w=_WK(_KE(erp="erp_hi"))), None),
       ("永續成長率 g：低", '="g "&TEXT(g_tv_lo,"0.0%")&"（基準 "&TEXT(g_tv,"0.0%")&"）"', "g_tv", dict(g="g_tv_lo"), None),
       ("永續成長率 g：高", '="g "&TEXT(g_tv_hi,"0.0%")&"（上限 rf）"', "g_tv", dict(g="g_tv_hi"), None),
       ("RONIC＝WACC＋5pt", "永續期新投資報酬率 RONIC＝WACC → WACC＋5pt", "ronic_sp", dict(ronic="(wacc+ronic_sp_hi)"), None),
       ("資本支出 ÷ 營收 −1pt", "每年資本支出比例 −1 個百分點", "fc_cx_d", dict(cxd="fc_cx_d_lo"), None),
       ("資本支出 ÷ 營收 ＋1pt", "每年資本支出比例 ＋1 個百分點", "fc_cx_d", dict(cxd="fc_cx_d_hi"), None),
       ("營運資金 ÷ 營收：低", '="營運資金比例 "&TEXT(fc_wc_lo,"0%")&"（基準 "&TEXT(fc_wc,"0.0%")&"）"', "fc_wc", dict(wc="fc_wc_lo"), None),
       ("營運資金 ÷ 營收：高", '="營運資金比例 "&TEXT(fc_wc_hi,"0%")&"（基準 "&TEXT(fc_wc,"0.0%")&"）"', "fc_wc", dict(wc="fc_wc_hi"), None),
       ("淨現金不含長期投資", "資金與投資（3Q26 簡報口徑：流動金融資產＋長期投資）不計入評價", "inv_k", dict(nc="(cash_v+inv_k_lo*inv_v-debt_v-lease_v-mi_v)"), None),
       ("終值改用 FCFF ×（1＋g）", "終值＝FY2032 FCFF ×（1＋g）÷（WACC − g）（替代法）", "—", dict(tv="alt"), None),
       ("美元口徑折現", '="rf＝美國 10 年期 "&TEXT(rf_us,"0.00%")&"（基準當地 "&TEXT(rf,"0.00%")&"）＋ β ×（ERP＋CRP）、g＝MIN("&TEXT(g_cap,"0%")&", "&TEXT(rf_us,"0.00%")&")；其他不變"', "rf_us",
        dict(ke="(rf_us+beta*(erp+crp))", w=_WK("(rf_us+beta*(erp+crp))"), g="MIN(g_cap,rf_us)"), None),
       ("DCF 權重 0.3", "Comparable 0.7", "tp_w_dcf", {}, "0.3"), ("DCF 權重 0.7", "Comparable 0.3", "tp_w_dcf", {}, "0.7")]
# 公司特有案例（v0.1.0）：In 頁區間較大的資本支出強度（fc_cx 25%／45%）與完全稀釋股數（sh27 0.57／0.62 十億股）。
# 資本支出：cxd 覆寫為（區間端點 − 基準 ＋ fc_cx_d），所有明確期年度與延伸期同移；股數：每股 DCF ＝ 加總價值 ÷ 股數，故以 sh_v ÷ 端點股數縮放；EPS（FY2027 起）同比例縮放。
_SHX = lambda e, s_x: f"(({e})*sh_v/{s_x}*(1+ke)^t_tgt-dps_tgt)"
FDC += [("資本支出 ÷ 營收 25%（In 區間低端）", '="資本支出占營收 "&TEXT(fc_cx_lo,"0%")&"（基準 "&TEXT(fc_cx,"0%")&"）；所有年度同移"', "fc_cx", dict(cxd="(fc_cx_lo-fc_cx+fc_cx_d)"), None),
        ("資本支出 ÷ 營收 45%（In 區間高端）", '="資本支出占營收 "&TEXT(fc_cx_hi,"0%")&"（基準 "&TEXT(fc_cx,"0%")&"）；所有年度同移"', "fc_cx", dict(cxd="(fc_cx_hi-fc_cx+fc_cx_d)"), None),
        ("完全稀釋股數 0.57（低端）", '="FY2027 起股數 "&TEXT(sh27_lo*1000,"0")&" 百萬股（基準 "&TEXT(sh27*1000,"0")&"）；DCF 每股與 EPS 同比例縮放"', "sh27", dict(sh="sh27_lo"), None),
        ("完全稀釋股數 0.62（高端）", '="FY2027 起股數 "&TEXT(sh27_hi*1000,"0")&" 百萬股（基準 "&TEXT(sh27*1000,"0")&"）；DCF 每股與 EPS 同比例縮放"', "sh27", dict(sh="sh27_hi"), None)]
n_tp = len(FDC) + 1
STOP = 4 + 2 * len(cases) + n_tp + 4
base_row = STOP + 2; SB = {}
SE.cell(STOP, 1, "敏感度「逐案推導區塊」：以下每個案例三個情境各一個完整推導區塊（不是案例列）").font = BOLD
i = 5
for key, name, unit, fmtv in cases:
    for side in ("lo", "hi"):
        blocks = {}
        for s in SC:
            ov = {key: f"{key}_{side}", "scn": s, "fm_self": True}
            blocks[s] = engine(SE, base_row, f"案例區塊：{name}：{'低' if side == 'lo' else '高'}（情境 {s}）", **ov)
            base_row = blocks[s]["end"] + 2
        SB[(key, side)] = blocks
        setting = f'=TEXT({key}_{side},"{fmtv}")&"（基準 "&TEXT({key},"{fmtv}")&"）"'
        _kw = dict(tax=f"tax_{side}") if key == "tax" else (dict(nonop=f"nonop_f_{side}") if key == "nonop_f" else {})
        _Rc = ddetail(blocks, "Sens!", f"Sens 案例：{name}：{'低' if side == 'lo' else '高'}（三情境 FCFF 明細）", lab="案例", **_kw)
        put_case(i, f"{name}：{'低' if side == 'lo' else '高'}", setting, key, blocks, "", v12ref(_Rc))
        SENS_ROWS[(key, side)] = i; i += 1
TP_ROWS = {}
for nm, st, k, over, wd in FDC:
    if "sh" in over:
        _x = over["sh"]
        put_case(i, nm, st, k, PB, "PL!", {s_: _SHX(FD.value_expr(RD_, s_, push=False), _x) for s_ in SC}, scale=f"(sh27/{_x})")
    else:
        put_case(i, nm, st, k, PB, "PL!", {s_: FD.value_expr(RD_, s_, **over) for s_ in SC}, **({"w": wd} if wd else {}))
    TP_ROWS[nm] = i; i += 1
SE.cell(i, 1, "檢查：閉式公式的基準 DCF 目標價 − Val 第 0.1 節（應為 0；閉式與 DCF 頁明細同式）")
SE.cell(i, 7, "=" + "+".join(f"tp_w{s_}*{FD.value_expr(RD_, s_)}" for s_ in SC) + f"-Val!B{T['dcftp']}").number_format = "0.000000"; i += 1
SENS_LAST = i - 1
SE.cell(i + 1, 1, "讀法：" + C["sens_read"])
SENS_NOTE = i + 1
assert SENS_NOTE < STOP, (SENS_NOTE, STOP)
SE.column_dimensions["A"].width = 34; SE.column_dimensions["B"].width = 26; SE.column_dimensions["C"].width = 12
for c in "DEFGHIJ": SE.column_dimensions[c].width = 15

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = f"Recon — 本模型 vs 公司說法、共識、券商。只作事後對照，不為對上數字而調參數。券商與資料商數字為 Interested-party，不用於校準"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.00", pt=False, gap=True):
    global rr_
    RC.cell(rr_, 1, label)
    if ref is not None: RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        if ref is not None and gap:
            L = chr(64 + cv)
            RC.cell(rr_, cd, f"={L}{rr_}-B{rr_}" if pt else f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y: (lambda s: f"=PL!{FC[y]}{PB[s][key]}")
for item in C["recon"]:
    if item[0] == "sec":
        RC.cell(rr_, 1, item[1]).font = BOLD; rr_ += 1; continue
    lab, ref, (key, y), src, note, fmt, pt = item
    if key == "feps": getter = lambda s: f"=Val!B{T['feps']}"
    elif key.startswith("="): getter = (lambda f_: (lambda s: f_))(key)
    else: getter = pl(key, y)
    rec(lab, ref, getter, src, note, fmt=fmt, pt=pt)
RC.cell(rr_, 1, "預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "成熟製程稼動率與價格的行業層（Siliconomics）尚未建立：上線後以本模型稼動率、售價情境對照行業路徑"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.column_dimensions["A"].width = 80
for c in "BCDEFGH": RC.column_dimensions[c].width = 11
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60
RECON_LAST = rr_ - 1

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司財報與法說內容部分為媒體轉述；券商、資料商、媒體皆為 Interested-party）"; SO["A1"].font = BOLD
for i, (a, b) in enumerate(C["sources"], 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 100; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
for i, (t, b) in enumerate(C["readme"], 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

_up = T["upw"]
GRADE_ROW = V.max_row + 2; REC_ROW = GRADE_ROW + 1
V.cell(GRADE_ROW, 1, "上檔空間級距（六級，供分組）")
V.cell(GRADE_ROW, 2, f'=IF(B{_up}>0.5,"大於 50%",IF(B{_up}>0.25,"25%–50%",IF(B{_up}>=0,"0%–25%",IF(B{_up}>=-0.25,"0%～−25%",IF(B{_up}>=-0.5,"−25%～−50%","小於 −50%")))))')
V.cell(GRADE_ROW, 5, f"以「上檔空間：加權目標價」（B{_up}）分六級（Andy 2026-10-25 14:00）")
V.cell(REC_ROW, 1, "投資建議")
V.cell(REC_ROW, 2, f'=IF(B{_up}>0.5,"強力買進",IF(B{_up}>0.25,"買進",IF(B{_up}>=-0.25,"中立",IF(B{_up}>=-0.5,"放空","強力放空"))))')
V.cell(REC_ROW, 5, "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
_rr = RD.max_row + 1
RD.cell(_rr, 1, "")
RD.cell(_rr + 1, 1, f'="投資建議："&Val!B{REC_ROW}&"（依上檔空間固定規則）"').font = Font(bold=True)
OBS_ROW0 = _rr + 2
for _j, _t in enumerate(C["obs"], 1): RD.cell(_rr + 1 + _j, 1, f"重點觀察 {_j}：" + _t)

order = ["README", "In", "Q", "Units", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
import relative_multiple as RMX  # noqa: E402
RMX._universe = lambda: {"gfs": {"own_ticker": "GFS", "公司": "格芯"}}   # 不改 tools/：以程式內設定取代 universe.csv／peers.csv（gfs 尚未列入）
RMX._peers = lambda folder: [("台積電（TSMC）", "TSM"), ("聯電（UMC）", "UMC"), ("高塔半導體（Tower）", "TSEM"), ("德州儀器（TXN）", "TXN"), ("中芯國際（SMIC）", "0981.HK")]
RMX.apply(wb, C["folder"])
RD.cell(RD.max_row, 1, "相對倍數調整（tools/relative_multiple.py；Val 第 0.5 節）已含在加權目標價：Comparable（相對倍數調整後）＝同業前瞻倍數中位數 × 本公司 10 年相對同業的倍數比（Bloomberg 前瞻倍數）。「Comparable 目標價（市場主流同業法）」維持不調整。重點觀察的數字以本版重算值撰寫（第 1 點第一句為公式）。")
sys.path.insert(0, os.path.join(ROOT, "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.1.0", "2026-11-17")
for _r in range(4, IN.max_row + 1):   # v0.2.2（第三輪修正第 15、23 點）：表頭統一順利／平平／不如預期；收斂法列名中性
    if str(IN.cell(_r, 1).value or "").startswith("情境驅動對照"):
        for _j, _h in zip((4, 5, 6), ("情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）")): IN.cell(_r + 1, _j).value = _h
    if str(IN.cell(_r, 1).value or "").startswith("scn_fade"): IN.cell(_r, 2).value = re.sub(r"（[^）]*）$", "", str(IN.cell(_r, 2).value))
# ---------------- v0.13（2026-11-04）：Sens「Wattonomics 錨定」案例（三層校準原則；決策 S37）。In 頁最後三列、Sens 頁最下方小表與案例區塊；既有格不動 ----------------
WTA_NOTE = "Wattonomics 錨定：以情境 B 為底，產業量 CY2027 起 × wta（Siliconomics SIF_WTAnchor_Base＝R4 比值）；FY2026 維持錨點。三層校準原則（Andy 2026-11-04）；個股基準暫不改，2027-01～02 雲端資本支出指引後再決定。"
WTA_MF = {2027: "wta_27", 2028: "wta_28", 2029: "wta_29"}
def wta_in_rows():
    r = IN.max_row + 2
    IN.cell(r, 1, f"Z. Wattonomics 錨定係數（敏感度用；Siliconomics {SIL.VERSION}）").font = BOLD; r += 1
    def _inp(key, item, unit, lo, base, hi, tag, s1, s2):
        nonlocal r
        for j, v in enumerate((key, item, unit, lo, base, hi, tag, s1, s2), 1):
            if v is not None: IN.cell(r, j, v)
        IN.cell(r, 5).number_format = "0.000"
        wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
        r += 1
    SIL.wta_rows(_inp)
def wta_sens():
    t0 = SE.max_row + 3; brow_ = t0 + 9; WB = {}
    for s in SC:
        WB[s] = engine(SE, brow_, f"案例區塊：Wattonomics 錨定（情境 {s}；FY2027 起 矽光子營收 × wta_27／wta_28／wta_29）", scn=s, fm_self=True, wm=WTA_MF)
        brow_ = WB[s]["end"] + 2
    SE.cell(t0, 1, "Wattonomics 錨定（三層校準原則）").font = BOLD
    hdr(SE, t0 + 1, [SE.cell(3, c).value for c in range(1, 11)] + ["FY2027 EPS（B）", "FY2028 EPS（B）"])
    put_case(t0 + 2, "基準（三情境加權；同第 4 列）", "PL 頁三個推導區塊", "—", PB, "PL!", v12ref(RD_))
    _Rw = ddetail(WB, "Sens!", "Sens 案例：Wattonomics 錨定（三情境 FCFF 明細）", lab="錨定")
    put_case(t0 + 3, "Wattonomics 錨定", "FY2027 起 矽光子營收（先進封裝中介層、矽光子）× wta（三情境同乘；權重不變）", "wta_27／wta_28／wta_29", WB, "", v12ref(_Rw))
    for rr, bl, pre in ((t0 + 2, PB, "PL!"), (t0 + 3, WB, "")):
        SE[f"K{rr}"] = f"={pre}E{bl['B']['eps']}"; SE[f"L{rr}"] = f"={pre}F{bl['B']['eps']}"
        for c in "KL": SE[f"{c}{rr}"].number_format = "#,##0.00"
    SE.cell(t0 + 4, 1, "差異%：Wattonomics 錨定 ÷ 基準 − 1")
    for c in "DEGHIKL":
        SE[f"{c}{t0 + 4}"] = f'=IFERROR({c}{t0 + 3}/{c}{t0 + 2}-1,"")'; SE[f"{c}{t0 + 4}"].number_format = "0.0%"
    SE.cell(t0 + 5, 1, WTA_NOTE)
    SE.cell(t0 + 6, 1, "本模型 AI 量驅動只有 矽光子營收（FY2026 約占營收 3.4%）；成熟製程晶圓與 技術服務不乘。FY(t)＝CY(t)；FY2030–FY2032 沿用收斂法。目標價為三情境皆 × wta 後依情境權重加權。")
    return WB
wta_in_rows(); WTA_R = wta_sens()   # v0.13：放在 In 版面套用之後，確保在 In 頁最後
_vr = lambda p: next(r for r in range(1, V.max_row + 1) if str(V.cell(r, 1).value or "").strip().startswith(p))
_B = lambda k, c="B": f"Val!{c}{T[k]}"
_E = lambda s, y: f"PL!{FC[y]}{SUMR[(s, 'eps')]}"
_U = lambda k, s, y: f"Units!{dict(zip((2026, 2027, 2028, 2029), 'CDEF'))[y]}{UR[f'{k}_{s}']}"
_S = lambda k, side: f"Sens!J{SENS_ROWS[(k, side)]}"
_NB = lambda s_, y=2029: f"(PL!{FC[y]}{SUMR[(s_, 'ai_rev')]}+PL!{FC[y]}{SUMR[(s_, 'int_rev')]})/PL!{FC[y]}{SUMR[(s_, 'rev')]}"   # 矽光子＋技術服務 ÷ 營收
_own, _npeer, _pmed10, _raw = (f"Val!B{_vr(p)}" for p in ("本公司 10 年中位數", "納入的同業家數", "同業 10 年中位數的中位數", "原始比值"))
_rec, _f0 = f"Val!B{REC_ROW}", f"Val!B{O + 8}"
_T = lambda x, f="0.0": f'TEXT({x},"{f}")'
_pc = lambda x: f'TEXT({x},"+0.0%;-0.0%")'
def XF(*parts):
    out = []
    for p_ in parts:
        if isinstance(p_, tuple): out.append(p_[0]); continue
        t_ = p_.replace('"', '""')
        out += [f'"{t_[i:i + 200]}"' for i in range(0, len(t_), 200)]
    return "=" + "&".join(out)
_SP = '"+0.0%;-0.0%"'
_FDN = {0: ("WACC −1pt", "WACC ＋1pt"), 1: ("β：低", "β：高"), 2: ("ERP：低", "ERP：高"), 5: ("資本支出 ÷ 營收 −1pt", "資本支出 ÷ 營收 ＋1pt")}
_K = lambda k: f"Sens!J{TP_ROWS[_FDN[k][0]]}"; _KH = lambda k: f"Sens!J{TP_ROWS[_FDN[k][1]]}"
_tv = SDCF["tvsh"]
RD.cell(OBS_ROW0, 1, XF("重點觀察 1：目標價 ", (f'TEXT({_B("tpw")},"0.0")',), " 美元（DCF ", (f'TEXT({_B("dcftp")},"0.0")',), "、Comparable 相對倍數調整後 ", ('TEXT(tp_comp_adj,"0.0")',), "；DCF 權重 ", ('TEXT(tp_w_dcf,"0%")',),
    "），比現價 ", ('TEXT(px,"0.00")',), " 美元（2026-11-13 收盤）", (f'IF({_B("upw")}>=0," 高 "," 低 ")',), (f'TEXT(ABS({_B("upw")}),"0.0%")',), "，投資建議 ", (_rec,),
    "。相對倍數調整係數 ", ('TEXT(rm_factor,"0.00")',), "（格芯 Bloomberg 前瞻 P/E 只有 2022–2026 五年，不足七年門檻，不調整）；Comparable 單獨（市場主流同業法）上檔空間 ", (f'TEXT({_B("upc")},{_SP})',),
    "；分析師共識目標價 ", (f'TEXT({_B("cons")},"0.0")',), " 美元（Yahoo 快照，非 IFRS EPS 基礎）。"))
RD.cell(OBS_ROW0 + 1, 1, XF("重點觀察 2：", (f'IF(ABS({_B("upw")})>0.25,"與現價的差距超過 25%，主因："&IF({_B("upw")}<0,"DCF 偏低與 Comparable 對 IFRS EPS 套非 GAAP 同業倍數。","DCF 與 Comparable 都高於現價。"),"與現價的差距在 25% 以內，評等中立。")',),
    " DCF 受 2027–28 年大額資本支出（公司指引 2027 年淨資本支出約 40 億美元，模型取營收 ", ('TEXT(fc_cx,"0%")',), "）與完全稀釋股數 ", ('TEXT(sh27*1000,"0")',), " 百萬股影響；WACC ", ('TEXT(wacc,"0.0%")',), "（β ", ('TEXT(beta,"0.00")',),
    "，單一來源）。Sens：資本支出占營收 −1／＋1 個百分點使加權目標價 ", (f'TEXT({_K(5)},{_SP})',), "／", (f'TEXT({_KH(5)},{_SP})',), "；WACC −1／＋1 個百分點 ", (f'TEXT({_K(0)},{_SP})',), "／", (f'TEXT({_KH(0)},{_SP})',), "。"))
RD.cell(OBS_ROW0 + 2, 1, XF("重點觀察 3：稼動率與混合售價是承重變數：公司稱 2Q26、3Q26 稼動率高 80%、2027 年混合 ASP 與產品組合約 +10%（模型情境 A 取製造服務售價 +8%、B +5%、C +2%）；售價低／高 3% 使加權目標價 ",
    (f'TEXT({_S("asp_adj", "lo")},{_SP})',), "／", (f'TEXT({_S("asp_adj", "hi")},{_SP})',), "，稼動率低／高 3 個百分點 ", (f'TEXT({_S("ut_adj", "lo")},{_SP})',), "／", (f'TEXT({_S("ut_adj", "hi")},{_SP})',),
    "，矽光子營收 ±30% 只有 ", (f'TEXT({_S("ai_adj", "lo")},{_SP})',), "／", (f'TEXT({_S("ai_adj", "hi")},{_SP})',), "。3Q26 IFRS 淨損來自 4.41 億美元非現金股份發行（美國政府），營業利益 2.30 億美元為正。下一個檢驗點是 2027-02 的 4Q26 法說（2027 年資本支出、售價與股數；日期未公告）。"))
r = RD.max_row + 2
RD.cell(r, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；上方各列原文保留給下游）").font = Font(bold=True); r += 1
RD.cell(r, 1, "命題：格芯的獲利由成熟與特殊製程的稼動率與混合售價決定；資料中心與矽光子是增量，2027 年大額資本支出與股數增加是壓力。以下三個問題依序回答。"); r += 1
RD.cell(r, 1, "命題 1：稼動率與售價能回升多少。稼動率已在高 80%、公司估 2027 年混合售價與組合約 +10%，這個回升能撐起多少 EPS。"); r += 1
RD.cell(r, 1, "命題 2：矽光子與資料中心能補多少。矽光子 2026 年約 5 億美元、2027 年底年化約 10 億美元，到 FY2029 占營收多少。"); r += 1
RD.cell(r, 1, "命題 3：每股值多少。用 DCF（含 2027 年約 40 億美元資本支出與約 6 億股）與同業倍數算出的目標價，和現價相比是高是低。"); r += 1
RD.cell(r, 1, f'="重點 1：加權目標價約 "&{_T(_B("tpw"), "0.0")}&" 美元，比現價 "&{_T("px", "0.00")}&" 美元"&IF({_B("upw")}>=0,"高 ","低 ")&TEXT(ABS({_B("upw")}),"0.0%")&"，投資建議"&{_rec}&"。DCF 約 "&{_T(_B("dcftp"))}&" 美元（權重 "&TEXT(tp_w_dcf,"0%")&"）、Comparable（相對倍數調整後）約 "&{_T("tp_comp_adj")}&" 美元；Comparable 用同業前瞻 P/E 中位數 "&{_T(_B("pmed"))}&" 倍 × 12 個月後 EPS（IFRS，情境加權）"&{_T(_B("feps"), "0.00")}&" 美元，現價相當於 "&{_T("px/" + _B("feps"))}&" 倍。相對倍數調整係數 "&{_T("rm_factor", "0.00")}&"（資料不足，不調整）。同業倍數為非 GAAP EPS 基礎，對格芯 IFRS EPS 偏保守。"'); r += 1
RD.cell(r, 1, f'="重點 2：三情境的差別主要在稼動率與售價，兩者都是假設。三情境 FY2027 稼動率 "&TEXT(ut27_A,"0%")&"／"&TEXT(ut27_B,"0%")&"／"&TEXT(ut27_C,"0%")&"、製造服務售價年增 "&TEXT(asp27_A,"+0%;-0%;0%")&"／"&TEXT(asp27_B,"+0%;-0%;0%")&"／"&TEXT(asp27_C,"+0%;-0%;0%")&"，FY2027 EPS 分別約 "&{_T(_E("A", 2027), "0.00")}&"、"&{_T(_E("B", 2027), "0.00")}&"、"&{_T(_E("C", 2027), "0.00")}&" 美元（FY2026 錨點 "&{_T(_E("B", 2026), "0.00")}&" 美元含 3Q26 政府股份發行費用 0.79）。敏感度（Sens 頁以三情境加權目標價為基準）：售價水準低／高 3% 使 "&{_pc(_S("asp_adj", "lo"))}&"／"&{_pc(_S("asp_adj", "hi"))}&"，稼動率低／高 3 個百分點使 "&{_pc(_S("ut_adj", "lo"))}&"／"&{_pc(_S("ut_adj", "hi"))}&"；矽光子營收 ±30% 只有 "&{_pc(_S("ai_adj", "lo"))}&"／"&{_pc(_S("ai_adj", "hi"))}&"。"'); r += 1
RD.cell(r, 1, f'="重點 3：矽光子與資料中心補得有限但成長快。FY2029 矽光子營收三情境約 "&{_T(_U("ai_usd", "A", 2029), "0.00")}&"／"&{_T(_U("ai_usd", "B", 2029), "0.00")}&"／"&{_T(_U("ai_usd", "C", 2029), "0.00")}&" 十億美元，占營收 "&TEXT({_U("aish", "A", 2029)},"0.0%")&"／"&TEXT({_U("aish", "B", 2029)},"0.0%")&"／"&TEXT({_U("aish", "C", 2029)},"0.0%")&"；技術服務 FY2029 約 "&{_T(_U("int_usd", "A", 2029), "0.00")}&"／"&{_T(_U("int_usd", "B", 2029), "0.00")}&"／"&{_T(_U("int_usd", "C", 2029), "0.00")}&" 十億美元；兩者合計約占營收 "&TEXT({_NB("A")},"0.0%")&"／"&TEXT({_NB("B")},"0.0%")&"／"&TEXT({_NB("C")},"0.0%")&"。矽光子金額為公司口述（2026 年約 5 億、2027 年底年化 10 億），毛利率未揭露，皆為假設。"'); r += 1
RD.cell(r, 1, XF("重點 4：DCF 目標價 ", (f'TEXT({_B("dcftp")},"0.0")',), " 美元。折現率 WACC ", ('TEXT(wacc,"0.0%")',), "（股權成本 ", ('TEXT(ke,"0.0%")',), "＝rf ", ('TEXT(rf,"0.00%")',), " ＋ β ", ('TEXT(beta,"0.00")',),
    " ×（ERP 5% ＋ 國家風險溢酬 0）；β 為 Yahoo 單一來源去槓桿後重新加槓桿，ERP 5% 與 RONIC＝WACC 為假設），永續成長率 ", ('TEXT(g_tv,"0.0%")',), "。情境 C 終值占企業價值 ", (f'TEXT(Val!D{_tv},"0%")',),
    "；情境 B 占 ", (f'TEXT(Val!C{_tv},"0%")',), "、隱含終值 EV／EBITDA ", (f'TEXT(Val!C{SDCF["mult"]},"0.0")',), " 倍。DCF 參數敏感度（以情境加權目標價為基準）：β 低／高 ", (f'TEXT({_K(1)},{_SP})',), "／", (f'TEXT({_KH(1)},{_SP})',),
    "、ERP 4%／6% ", (f'TEXT({_K(2)},{_SP})',), "／", (f'TEXT({_KH(2)},{_SP})',), "、資本支出 ±1 個百分點 ", (f'TEXT({_K(5)},{_SP})',), "／", (f'TEXT({_KH(5)},{_SP})',), "。")); r += 1
RD.cell(r, 1, "重點 5：美元計價，Sens「美元口徑折現」案例與基準相同（無匯率與利率口徑差）；不另列。"); r += 1
r += 1
RD.cell(r, 1, "命題回覆（HTML 用；依本模型計算逐一回答上面三個命題）").font = Font(bold=True); r += 1
RD.cell(r, 1, f'="稼動率與售價能回升多少？答：有限。FY2026 平均稼動率約 "&TEXT({_U("ut", "B", 2026)},"0.0%")&"（高 80% 為公司說法；數值為假設推得），FY2027 三情境設為 "&TEXT(ut27_A,"0%")&"／"&TEXT(ut27_B,"0%")&"／"&TEXT(ut27_C,"0%")&"；製造服務平均售價由 FY2026 約 "&{_T(_U("asp", "B", 2026), "#,##0")}&" 美元／片到 FY2027 "&{_T(_U("asp", "A", 2027), "#,##0")}&"／"&{_T(_U("asp", "B", 2027), "#,##0")}&"／"&{_T(_U("asp", "C", 2027), "#,##0")}&"。FY2027 EPS（IFRS）三情境約 "&{_T(_E("A", 2027), "0.00")}&"／"&{_T(_E("B", 2027), "0.00")}&"／"&{_T(_E("C", 2027), "0.00")}&" 美元；完全稀釋股數增至 "&TEXT(sh27*1000,"0")&" 百萬股、折舊隨資本支出上升，抵銷部分量價成長。"'); r += 1
RD.cell(r, 1, f'="矽光子與資料中心能補多少？答：FY2029 兩項新業務（矽光子與技術服務）合計占營收約 "&TEXT({_NB("A")},"0.0%")&"／"&TEXT({_NB("B")},"0.0%")&"／"&TEXT({_NB("C")},"0.0%")&"（情境 A／B／C），其中矽光子約 "&TEXT({_U("aish", "A", 2029)},"0.0%")&"／"&TEXT({_U("aish", "B", 2029)},"0.0%")&"／"&TEXT({_U("aish", "C", 2029)},"0.0%")&"。矽光子情境 A 約 "&{_T(_U("ai_usd", "A", 2029), "0.00")}&" 十億美元（FY2027 年底年化目標 "&{_T("co_sip27", "0.0")}&" 十億美元）。矽光子營收 ±30% 對加權目標價只有 "&{_pc(_S("ai_adj", "lo"))}&"／"&{_pc(_S("ai_adj", "hi"))}&"。資料中心的更大影響在毛利率組合與稼動率，已反映在稼動率與售價情境。"'); r += 1
RD.cell(r, 1, f'="每股值多少？答：加權目標價約 "&{_T(_B("tpw"), "0.0")}&" 美元（DCF "&{_T(_B("dcftp"))}&"、Comparable 相對倍數調整後 "&{_T("tp_comp_adj")}&"，權重各半），比現價 "&{_T("px", "0.00")}&" 美元"&IF({_B("upw")}>=0,"高 ","低 ")&TEXT(ABS({_B("upw")}),"0.0%")&"，投資建議"&{_rec}&"。市場共識目標價 "&{_T(_B("cons"))}&" 美元（非 IFRS EPS 基礎）。DCF 三情境價值（推到 12 個月後）樂觀、基準、保守約 "&{_T(_B("dcf12", "B"))}&"、"&{_T(_B("dcf12", "C"))}&"、"&{_T(_B("dcf12", "D"))}&" 美元。現價約為未來 12 個月 EPS（情境 B，IFRS）的 "&{_T(_f0)}&" 倍。"'); r += 1
r += 1
RD.cell(r, 1, "展示層輔助：Comparable 目標價（相對倍數調整後）的上檔空間（HTML 估值表用；加權目標價使用的 Comparable 即此值；投資建議橫幅中間格放只看 Comparable〔市場主流同業法〕，見 Val 頁）")
RD.cell(r, 2, "=tp_comp_adj/px-1").number_format = "0.0%"
# 三層協調 C4：對帳中年份並列兩讀數目標價（Coord 頁；快照 gfs/coord_alt.json；tools/coord/）
sys.path.insert(0, os.path.join(ROOT, "tools", "coord"))
import coord as COORD  # noqa: E402
_HERE = os.path.dirname(os.path.abspath(__file__))
_CR = COORD.add_sheet(wb, _HERE, SIL)
if _CR:
    COORD.stamp(wb, "v0.1.0", "v0.1.0", "2026-11-17")
    sys.path.insert(0, _HERE)
    import gfs_post  # noqa: E402
    print(gfs_post.apply(wb, _HERE, SIL, _CR, dict(PB=PB, C=C, SENS_ROWS=SENS_ROWS, TP_ROWS=TP_ROWS)))
wb.save(sys.argv[1])
if os.environ.get("ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "STOP": STOP, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE,
               "UN_NOTE": UN_NOTE, "T": T, "T10": T10, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0,
               "GRADE_ROW": GRADE_ROW, "REC_ROW": REC_ROW, "OBS_ROW0": OBS_ROW0, "RD_LAST": RD.max_row,
               "SENS_ROWS": {f"{k[0]}|{k[1]}": v for k, v in SENS_ROWS.items()}, "TP_ROWS": TP_ROWS, "SEG_N": len(SEG["rows"])},
              open(os.environ["ROWMAP"], "w"), ensure_ascii=False, indent=1)

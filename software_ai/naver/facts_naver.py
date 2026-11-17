#!/usr/bin/env python3
"""Naver（035420）事實總帳：build_naver.py 匯入本檔的 Q 序列與 FACTS；單獨執行時寫出 data/facts_naver.csv。
金額十億韓元（K-IFRS，會計年度＝曆年），EPS 與股價韓元，股數百萬股。擷取日 2026-11-16。標記只用 Verified／Interested-party／Analogy／Assumed／Derived。
"""
import csv, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
RET = "2026-11-16"
DL = "https://www.navercorp.com/api/article/download/"
DECK = {"3Q26": DL + "0cdefb0e-a7c5-4a8b-b785-9e026da9c305", "2Q26": DL + "07b3b329-bbbe-481c-ac47-b3d0d7fea855", "1Q26": DL + "896dc2b0-33ea-4dbf-a354-c9608c6d989b",
        "4Q25": DL + "197f039a-83a5-4cdd-be24-02f61a5c003c", "4Q24": DL + "8c766af7-6185-4eb1-aab6-373f659d0948", "4Q23": DL + "966e8489-dc3b-4724-9fd5-39524bdcfe7d",
        "4Q22": "https://www.navercorp.com/navercorp_/ir/earningsRelease/2023/4Q22%20Naver_Earnings_Eng.pdf"}
PR3 = "https://www.navercorp.com/en/media/pressReleasesDetail?seq=10034882"
BRK = "https://www.navercorp.com/en/media/pressReleasesDetail?seq=10034796"
IRPT = "https://www.navercorp.com/api/article/read/0479c513-b992-458e-a96b-d9a5151d88fc/NAVER_2025_Integrated%20Report.pdf"
QS = ["4Q24", "1Q25", "2Q25", "3Q25", "4Q25", "1Q26", "2Q26", "3Q26"]
# 1Q25–4Q25 為 2026 年起新分部口徑重編的比較欄（取自 1Q26／2Q26／3Q26 簡報的 YoY 比較欄）；4Q24 取 4Q25 簡報比較欄。
SRC_Q = {"4Q24": DECK["4Q25"], "1Q25": DECK["1Q26"], "2Q25": DECK["2Q26"], "3Q25": DECK["3Q26"], "4Q25": DECK["4Q25"], "1Q26": DECK["1Q26"], "2Q26": DECK["2Q26"], "3Q26": DECK["3Q26"]}
N = None
QDATA = {  # 十億韓元；4Q24 欄部分科目未取（N）
    "rev":   [2885.6, 2786.8, 2915.1, 3138.1, 3195.1, 3241.1, 3388.8, 3472.4],
    "np":    [1680.7, 1604.7, 1693.1, 1816.6, 1850.3, 1839.8, 1902.2, 1948.4],
    "ads":   [1322.0, 1276.2, 1346.6, 1398.1, 1415.3, 1394.5, 1447.2, 1474.8],
    "svc":   [358.7, 328.5, 346.6, 418.6, 434.9, 445.3, 455.0, 473.6],
    "fin":   [395.1, 386.7, 405.6, 427.3, 448.6, 459.7, 470.7, 487.7],
    "glob":  [809.8, 795.4, 816.4, 894.1, 896.2, 941.6, 1015.9, 1036.4],
    "c2c":   [214.0, 222.7, 227.5, 250.6, 285.7, 351.1, 397.9, 370.4],
    "cont":  [453.7, 446.2, 462.1, 498.2, 444.9, 440.1, 464.3, 472.6],
    "ent":   [142.1, 126.6, 126.7, 145.3, 165.6, 150.5, 153.7, 193.3],
    "devops": [N, 671.1, 683.1, 732.7, 743.7, 748.9, 786.9, 750.9],
    "partner": [N, 990.9, 1030.2, 1114.1, 1129.8, 1188.0, 1256.8, 1324.0],
    "infra": [N, 189.3, 197.7, 218.6, 205.2, 250.8, 219.8, 282.3],
    "mkt":   [N, 430.2, 482.4, 502.0, 505.7, 511.7, 604.9, 663.9],
    "opex":  [N, 2281.5, 2393.5, 2567.4, 2584.4, 2699.3, 2868.4, 3021.1],
    "oi":    [N, 505.3, 521.6, 570.6, 610.6, 541.8, 520.3, 451.3],
    "nonop": [N, 74.8, 162.1, 286.2, -311.1, -148.3, 381.7, 79.6],
    "eqm":   [N, -2.7, 161.2, 276.2, 60.5, -131.5, 238.7, 39.8],
    "pbt":   [N, 580.1, 683.6, 856.9, 299.5, 393.5, 902.0, 530.9],
    "tax":   [N, 156.4, 186.3, 122.2, 136.5, 102.5, 197.8, 128.5],
    "ni":    [N, 423.7, 497.4, 734.7, 163.0, 291.0, 704.3, 402.5],
    "np_m":  [N, 0.342, 0.327, 0.336, 0.333, 0.288, 0.285, 0.240],   # Naver Platform 營業利益率（公司揭露，一位小數百分比）
    "glob_oi": [N, -74.2, -65.0, -68.1, -31.2, -14.8, -35.6, -26.1],
    "tpv":   [N, 19.6, 20.8, 22.7, 23.0, 24.2, 25.2, 26.2],          # 支付總額（兆韓元，Financial Platform 的 Naver Pay）
}
QLAB = {"rev": "合併營收", "np": "Naver Platform 營收", "ads": "Naver Platform：廣告（Ads）營收", "svc": "Naver Platform：服務（Services）營收", "fin": "Financial Platform 營收", "glob": "Global Opportunities 營收",
        "c2c": "Global Opportunities：C2C 營收", "cont": "Global Opportunities：Contents 營收", "ent": "Global Opportunities：Enterprise 營收",
        "devops": "營業費用：開發與營運（Dev&Ops）", "partner": "營業費用：合作夥伴（Partner）", "infra": "營業費用：基礎設施（Infrastructure）", "mkt": "營業費用：行銷（Marketing）", "opex": "營業費用合計",
        "oi": "營業利益", "nonop": "營業外損益", "eqm": "營業外：權益法損益", "pbt": "稅前利益", "tax": "所得稅費用", "ni": "稅後淨利（合計）", "np_m": "Naver Platform 營業利益率",
        "glob_oi": "Global Opportunities 分部營業損益", "tpv": "Naver Pay 支付總額"}
QUNIT = {"np_m": "%", "tpv": "兆韓元"}
QBASIS = {"np": "2026 年起新分部（Naver Platform＝廣告＋服務）；1Q25–4Q25 為重編比較欄", "glob": "新分部 Global Opportunities（簡報英文稿另稱 Global Initiatives）",
          "devops": "2026 年起費用分類重編（Dev&Ops 與 Partner 間有重分類，與 FY2023–FY2024 舊口徑不連續）", "partner": "同上", "ni": "合計淨利（含非控制權益）；歸屬母公司拆分未找到"}

# 年度（舊五分部口徑；FY2021–FY2025）
YDATA = {
    "seg_search": ("Search Platform", [3307.8, 3568.0, 3589.1, 3946.2, 4168.9]),
    "seg_comm": ("Commerce", [1488.5, 1801.1, 2546.6, 2923.0, 3688.4]),
    "seg_fint": ("Fintech", [979.0, 1186.6, 1354.8, 1508.4, 1690.7]),
    "seg_cont": ("Content", [659.6, 1261.5, 1733.0, 1796.4, 1899.2]),
    "seg_cloud": ("Cloud／Enterprise", [382.6, 402.9, 447.2, 563.7, 587.8]),
    "y_rev": ("合併營收", [6817.6, 8220.1, 9670.6, 10737.7, 12035.0]),
    "y_oi": ("營業利益", [1325.5, 1304.7, 1488.8, 1979.3, 2208.1]),
    "y_ni": ("稅後淨利", [16477.6, 664.0, 988.4, 1932.0, 1818.7]),
}
YYR = ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025"]
YSRC = ["4Q22（FY2021 比較欄）", "4Q22", "4Q23", "4Q24", "4Q25"]
YOPX = {  # FY2023、FY2024 營業費用與損益（4Q23、4Q25 簡報）
    "y_devops": ("營業費用：Dev&Ops", {"FY2023": 2572.7, "FY2024": 2664.4}), "y_partner": ("營業費用：Partner", {"FY2023": 3533.6, "FY2024": 3789.9}),
    "y_infra": ("營業費用：Infrastructure", {"FY2023": 595.7, "FY2024": 704.8}), "y_mkt": ("營業費用：Marketing", {"FY2023": 1479.8, "FY2024": 1599.4}),
    "y_nonop": ("營業外損益", {"FY2023": -7.4, "FY2024": 342.9}), "y_pbt": ("稅前利益", {"FY2023": 1481.4, "FY2024": 2322.2}), "y_tax": ("所得稅費用", {"FY2023": 496.4, "FY2024": 390.2}),
}
ADJ_EBITDA = {"FY2021": 1914.6, "FY2022": 1802.8, "FY2023": 2133.8, "FY2024": 2664.4}
COMM_ADS_Q = [308.3, 335.6, 366.3, 390.3]   # 1Q25–4Q25 商務廣告（舊分類）

# 單點事實：key -> (指標, 數值, 單位, 期間, 口徑, 出處, 標記, 誘因)
import json as _j, os as _o
_PS = _j.load(open(_o.path.join(_o.path.dirname(_o.path.abspath(__file__)), "data", "peer_ps_naver.json"), encoding="utf-8"))
SINGLE = {
  "px": ("股價（收盤）", 172500, "韓元", "2026-11-13", "KRX 收盤（11-13 為最近交易日）", "韓國投資證券個股頁（KIS，當日 −5.12%）；與教保證券報告 73.9% 上檔空間隱含價一致", "Verified", ""),
  "sh": ("流通股數（扣庫藏）", 149.86, "百萬股", "2026-10", "第二來源（StockAnalysis）；法定申報封面股數未取得（未找到）", "https://stockanalysis.com/quote/krx/035420/statistics/", "Derived", ""),
  "nv_sh": ("NVIDIA 新發行股數（暫定）", 7.241564, "百萬股", "2026-11", "每股 204,500 韓元，共約 1 兆韓元；繳款日 2026-11-06、上市日 2026-11-27", BRK + "（Naver 新聞稿）；3Q26 簡報", "Verified", ""),
  "nv_px": ("NVIDIA 認購價", 204500, "韓元", "2026-11", "每股發行價（事件價）", BRK, "Verified", ""),
  "debt": ("借款（有息負債，公司定義）", 2996.3, "十億韓元", "2026-09-30", "借款合計；租賃負債未找到、不計（營業租賃一律不計）", DECK["3Q26"], "Verified", ""),
  "cash": ("合併現金與約當現金", 7710.3, "十億韓元", "2026-09-30", "含 NAVER Financial 之現金（金融事業現金不計入淨現金）", DECK["3Q26"], "Verified", ""),
  "netcash": ("淨現金（公司定義，不含 NAVER Financial）", 1974.9, "十億韓元", "2026-09-30", "公司揭露", DECK["3Q26"], "Verified", ""),
  "dps": ("每股股利", 2630, "韓元", "FY2025", "股利總額 393.6 十億韓元；占兩年平均自由現金流 30%", PR3 + "；FY2025 新聞稿 seq=10034311", "Verified", ""),
  "emp": ("員工數", 5047, "人", "2025-12", "整合報告 p.5；母公司或合併口徑未載明", IRPT, "Verified", ""),
  "ads_comm_fy25": ("商務廣告（舊分類，FY2025）", sum(COMM_ADS_Q), "十億韓元", "FY2025", "1Q25–4Q25 季度加總；舊五分部口徑（Commerce 內之廣告）", DECK["4Q25"], "Derived", ""),
  "ebitda25": ("EBITDA（第二來源）", 2953, "十億韓元", "FY2025", "MarketScreener；與公司口徑不同", "https://www.marketscreener.com/quote/stock/NAVER-CORPORATION-6491497/finances/", "Derived", ""),
  "capex25": ("資本支出（第二來源）", 1235, "十億韓元", "FY2025", "MarketScreener；公司未於簡報揭露年度資本支出（未找到）", "https://www.marketscreener.com/quote/stock/NAVER-CORPORATION-6491497/finances/", "Derived", ""),
  "beta_ss": ("β（StockAnalysis）", 1.19, "倍", "2026-11", "5 年月", "https://stockanalysis.com/quote/krx/035420/statistics/", "Verified", ""),
  "beta_yh": ("β（Yahoo Finance）", 1.20, "倍", "2026-11", "5 年月", "https://finance.yahoo.com/quote/035420.KS/", "Verified", ""),
  "rf": ("無風險利率（韓國 10 年期國債）", 0.0399, "比例", "2026-11-13", "殖利率收盤；Trading Economics 頁載 2026-11-16 為 3.986%、單日變動 −0.001，11-13 約 3.987%（頁面無 11-12、11-13 逐日值）", "https://tradingeconomics.com/south-korea/government-bond-yield", "Verified", ""),
  "rm_geo": ("歷史相對倍數基準（PS 與 PSG 兩讀數幾何平均）", 0.5617479268406912, "倍", "2021-09 至 2026-06", "本公司 ÷ 同業中位（GOOGL、META、PINS、Kakao、LY）；規格 docs/plan/歷史相對倍數_PS與PSG.md", "software_ai/data/relmult_hist/out/naver.json（geo_base；base_source＝PS 與 PSG 兩讀數幾何平均）", "Derived", ""),
  "rm_ps": ("歷史相對倍數：PS 讀數（歷史相對中位）", 0.6807807444453408, "倍", "2021-09 至 2026-06", "20 季；最低 0.36、最高 1.17", "同上（ps_reading）", "Derived", ""),
  **{f"ps_{k.lower()}": (f"同業 P/S（TTM 營收口徑）：{k}", v["ps"], "倍", v["d"], "2026-11-13 現價口徑；資料日不是 11-13 者於 Val 標逾期（G7）；Comparable 主腿", _PS["source"], "Derived", "") for k, v in _PS["peers"].items()},
  "rm_psg": ("歷史相對倍數：PSG 讀數（歷史相對中位 x 成長比）", 0.4635277009294886, "倍", "2021-09 至 2026-06", "16 季；成長比 0.722（TTM，本公司 14.6%、同業中位 20.2%）", "同上（psg_reading）", "Derived", ""),
  "ai_ads_1q": ("AI 對 1Q26 廣告成長的貢獻（公司稱）", 0.5, "占比（超過）", "1Q26", "公司稱 AI 貢獻廣告成長逾 50%", DECK["1Q26"], "Interested-party", "管理層向市場說明 AI 成效"),
  "ai_ads_2q": ("AI 對 2Q26 廣告成長的貢獻（公司稱）", 0.6, "占比（超過）", "2Q26", "公司稱 AI 貢獻廣告成長逾 60%", DECK["2Q26"], "Interested-party", "管理層向市場說明 AI 成效"),
  "ai_ctr": ("AI Briefing 廣告點擊率提升（公司稱）", 0.30, "比例", "2026-07-21 起", "AI Briefing 廣告上線後公司稱 CTR +30%、轉換 3 倍（非隨機對照）", DECK["2Q26"], "Interested-party", "管理層向市場說明 AI 成效"),
  "ai_tab_u": ("AI Tab 使用者", 10, "百萬人", "2026-07-15", "公司揭露", DECK["2Q26"], "Interested-party", "管理層向市場說明 AI 採用率"),
  "ai_brief_u": ("AI Briefing 月使用者（逾）", 30, "百萬人", "2026Q2", "公司揭露", DECK["2Q26"], "Interested-party", "管理層向市場說明 AI 採用率"),
  "kr_share": ("Naver 韓國搜尋占有率", 0.8134, "比例", "2026-05", "Internet Trend（一年前 59.26%）；第三方統計口徑", "https://www.internettrend.co.kr/", "Verified", ""),
  "af_contract": ("AI Factory 合約（20MW，4–5 年）", 2000, "十億韓元", "2026", "約 2 兆韓元；年化 400–500 十億韓元；Naver 保證使用費最長 8 年", BRK, "Interested-party", "管理層宣布合約，利害關係人說法"),
  "af_mw27": ("AI Factory 上線容量（2027 上半年）", 55, "MW", "1H27", "公司規劃", BRK, "Interested-party", "管理層規劃"),
  "af_mw28": ("AI Factory 上線容量（2028）", 200, "MW", "2028", "公司規劃；整體目標 1GW 規模", BRK, "Interested-party", "管理層規劃"),
  "gpu_n": ("GPU 取得數量（NVIDIA 合作備忘錄）", 60000, "片", "2025-10", "公司宣布", "https://www.navercorp.com/en/media/pressReleases", "Interested-party", "管理層宣布"),
  "gpu_life": ("GPU 折舊年限", 6, "年", "2026", "由 5 年延長為 6 年，年度折舊遞延約 100 十億韓元", DECK["1Q26"], "Verified", ""),
  "tb_cancel": ("庫藏股註銷", 4.901094, "百萬股", "2026-08-03", "約 1 兆韓元、3.1%", DECK["2Q26"], "Verified", ""),
  "ret_prog": ("股東回饋方案 FY2025–FY2027", "25%–35%", "占平均自由現金流", "FY2025–FY2027", "公司政策", PR3, "Verified", ""),
  "plat_goal": ("Naver Platform 長期營業利益率目標", 0.25, "比例", "長期", "公司稱「穩定在 25%」；無正式年度指引", DECK["3Q26"], "Interested-party", "管理層目標"),
  "cons_tp": ("共識平均目標價", 286538, "韓元", "2026-11-11", "MarketScreener 26 位分析師；財報前（財報前股價 194,300），只作歷史對照", "https://www.marketscreener.com/quote/stock/NAVER-CORPORATION-6491497/consensus/", "Interested-party", "賣方預估"),
  "cons_rev26": ("共識營收 FY2026E", 13741, "十億韓元", "2026-11-11", "財報前", "https://www.marketscreener.com/quote/stock/NAVER-CORPORATION-6491497/finances/", "Interested-party", "賣方預估"),
  "cons_oi26": ("共識營業利益 FY2026E", 2107, "十億韓元", "2026-11-11", "財報前", "同上", "Interested-party", "賣方預估"),
  "cons_eps26": ("共識 EPS FY2026E", 12318, "韓元", "2026-11-11", "財報前", "同上", "Interested-party", "賣方預估"),
  "cons_rev27": ("共識營收 FY2027E", 15791, "十億韓元", "2026-11-11", "財報前", "同上", "Interested-party", "賣方預估"),
  "cons_oi27": ("共識營業利益 FY2027E", 2465, "十億韓元", "2026-11-11", "財報前", "同上", "Interested-party", "賣方預估"),
  "cons_eps27": ("共識 EPS FY2027E", 14039, "韓元", "2026-11-11", "財報前", "同上", "Interested-party", "賣方預估"),
  "cons_capex26": ("共識資本支出 FY2026E", 1455, "十億韓元", "2026-11-11", "財報前", "同上", "Interested-party", "賣方預估"),
  "cons_capex27": ("共識資本支出 FY2027E", 1485, "十億韓元", "2026-11-11", "財報前", "同上", "Interested-party", "賣方預估"),
  "pe_goog": ("同業前瞻本益比：Alphabet（GOOGL）", 21.45, "倍", "2026-11-13", "US GAAP；NTM EPS；StockAnalysis", "https://stockanalysis.com/stocks/googl/statistics/", "Verified", ""),
  "pe_meta": ("同業前瞻本益比：Meta（META）", 19.48, "倍", "2026-11-13", "US GAAP；NTM EPS", "https://stockanalysis.com/stocks/meta/statistics/", "Verified", ""),
  "pe_pins": ("同業前瞻本益比：Pinterest（PINS）", 8.89, "倍", "2026-11-06", "US GAAP；NTM EPS；與現價日相差逾 2 個交易日", "https://stockanalysis.com/stocks/pins/statistics/", "Verified", ""),
  "pe_kakao": ("同業前瞻本益比：Kakao（035720）", 21.21, "倍", "2026-09-09", "K-IFRS；NTM EPS；逾期", "https://stockanalysis.com/quote/krx/035720/statistics/", "Verified", ""),
  "pe_ly": ("同業前瞻本益比：LY Corporation（4689）", 16.86, "倍", "2026-10-13", "IFRS；NTM EPS；逾期", "https://stockanalysis.com/quote/tyo/4689/statistics/", "Verified", ""),
  "tp_hana": ("券商目標價：Hana（財報後）", 340000, "韓元", "2026-11-13", "由 400,000 下調", "新聞報導（財報後各券商報告）", "Interested-party", "賣方預估"),
  "tp_kis": ("券商目標價：Korea Investment（財報後）", 300000, "韓元", "2026-11-13", "由 360,000 下調", "同上", "Interested-party", "賣方預估"),
  "tp_meritz": ("券商目標價：Meritz（財報後）", 270000, "韓元", "2026-11-13", "由 330,000 下調", "同上", "Interested-party", "賣方預估"),
  "tp_sk": ("券商目標價：SK（財報後）", 260000, "韓元", "2026-11-13", "", "同上", "Interested-party", "賣方預估"),
  "tp_kyobo": ("券商目標價：Kyobo（財報後）", 300000, "韓元", "2026-11-13", "由 390,000 下調；其 2027E 營收 15.9 兆、營業利益 2.3 兆", "同上", "Interested-party", "賣方預估"),
}
NOTFOUND = ["搜尋查詢量（年）", "AI Briefing 占查詢比例與每次回答的 token 量（公司未揭露，故每任務 token 無法以揭露校準）", "HyperCLOVA X 使用 GPU 數量與自用推論成本金額",
            "年度資本支出（FY2023–FY2024 與逐季；公司簡報未揭露）", "折舊攤銷與股票薪酬金額（K-IFRS 附註未取得）", "租賃負債", "權益法投資帳面價值", "RSU、選擇權明細與稀釋股數", "非控制權益拆分（歸屬母公司淨利）",
            "FY2021–FY2025 每股盈餘（K-IFRS 基本）", "4Q24 分部營業利益（新口徑）", "第二個共識來源（FnGuide 僅見單季營業利益共識）", "NAVER Financial 與 Dunamu 合併換股比例與會計影響",
            "法定申報封面股數（DART 季報）", "Baidu 等未取得同業本益比"]


def facts_rows():
    rows = []
    def add(metric, val, unit, period, basis, src, tag, inc=""):
        rows.append([metric, val, unit, period, basis, src, RET, tag, inc])
    for k, vals in QDATA.items():
        for q, v in zip(QS, vals):
            if v is None: continue
            tag = "Derived" if k in ("np_m",) and False else "Verified"
            add(QLAB[k], v * 100 if k == "np_m" else v, QUNIT.get(k, "十億韓元"), q, QBASIS.get(k, "K-IFRS 合併；2025 各季為新口徑重編比較欄"), "Naver 財報簡報 " + q + "（" + SRC_Q[q] + "）；頁：業績表（頁碼未逐頁記錄）", tag)
    for k, (lab, vals) in YDATA.items():
        for y, v, s in zip(YYR, vals, YSRC):
            add(lab + "（舊五分部口徑）" if k.startswith("seg_") else lab, v, "十億韓元", y, "舊口徑；FY2025 為 2026 年重編後之季度加總" if y == "FY2025" else ("含 LINE-Z 合併利益" if k == "y_ni" and y == "FY2021" else "K-IFRS 合併"), "Naver 財報簡報 " + s + "（" + DECK[s[:4]] + "）；頁：年度業績表（頁碼未逐頁記錄）", "Derived" if y == "FY2025" and k != "y_oi" else "Verified")
    for k, (lab, d) in YOPX.items():
        for y, v in d.items():
            add(lab, v, "十億韓元", y, "舊口徑", "Naver 財報簡報 " + ("4Q23" if y == "FY2023" else "4Q25") + "（" + DECK["4Q23" if y == "FY2023" else "4Q25"] + "）；頁：年度費用表（頁碼未逐頁記錄）", "Verified")
    for y, v in ADJ_EBITDA.items():
        add("調整後 EBITDA（公司定義，不含 SBC）", v, "十億韓元", y, "非 K-IFRS；公司自訂", "Naver 財報簡報 4Q22／4Q23／4Q24", "Interested-party", "管理層自訂口徑")
    for q, v in zip(["1Q25", "2Q25", "3Q25", "4Q25"], COMM_ADS_Q):
        add("商務廣告（舊分類）", v, "十億韓元", q, "舊五分部口徑，Commerce 內之廣告", DECK["4Q25"], "Verified")
    for k, (m, v, u, p, b, s, t, i) in SINGLE.items():
        add(m, v, u, p, b, s, t, i)
    for nf in NOTFOUND:
        add(nf, "未找到", "—", "—", "不估計；模型以 Assumed 輸入並標明", "—", "Assumed", "")
    return rows


if __name__ == "__main__":
    out = os.path.join(HERE, "data", "facts_naver.csv")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    rows = facts_rows()
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["指標", "數值", "單位", "期間", "口徑", "出處", "擷取日", "標記", "來源誘因"])
        w.writerows(rows)
    print("facts_naver.csv", len(rows), "列")

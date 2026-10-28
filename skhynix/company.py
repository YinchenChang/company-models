# SK 海力士設定檔（樣板 v0.2）。只放參數；結構在 tools/memory_model/build.py 與 markets/memory/market.py
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))), "siliconomics"))
import reader as _SIL   # HBM 營收份額 sh25–sh29 讀 Siliconomics M24（v0.8 起；數值權威在本庫 In 頁）
S = "sec"
CONFIG = {
 "name": "SK 海力士", "version": "v0.5", "date": "2026-10-25",
 "unit": "兆韓元", "ccy_per_usdB": "fx/1000", "share_scale": 1000000,
 "base_fy": 2026, "fy_offset": 0, "fy_label": "{}", "fy_desc": "日曆年（K-IFRS 合併）",
 "thesis": "SK 海力士的盈餘由兩條價格線決定——每季議價的一般 DRAM 與年約鎖價的 HBM；股價取決於這個循環在哪一年見頂，以及見頂時累積了多少淨值。",
 "q_source": "SEC 424B4（2026-07）、6-K 半年報（2026-08-18）、公司公告",
 "q_note": "2026Q1 DRAM 77.3%、NAND 22%、毛利率約 79%；Q2 DRAM 73%、NAND 27%、毛利率 83%（公司簡報）。2025 季度拆分未取得。2025Q3、Q4 折舊為合計 7.113 平均分配 [Derived]",
 "hbm_share": {2025: "sh25", 2026: "sh26", 2027: "sh27", 2028: "sh28", 2029: "sh29"},
 "pb_names": ["P/B 1.0x", "P/B 歷史中位 1.6x", "P/B 券商遠期 3.3x"], "pe_vals": [4, 8, 12],
 "quarters": [
  {"label": "2025Q1", "type": "A", "rev": 17.639, "op": 7.441, "ni": 8.108, "da": 3.334},
  {"label": "2025Q2", "type": "A", "rev": 22.232, "op": 9.213, "ni": 6.996, "da": 3.443},
  {"label": "2025Q3", "type": "A", "rev": 24.449, "op": 11.383, "ni": 12.598, "da": 3.556},
  {"label": "2025Q4", "type": "A", "rev": 32.827, "op": 19.170, "ni": 15.246, "da": 3.557},
  {"label": "2026Q1", "type": "A", "rev": 52.576, "dram": "=F4*0.773", "nand": "=F4*0.22", "cogs": "=F4*(1-0.79)", "op": 37.610, "ni": 40.346, "da": 3.726},
  {"label": "2026Q2", "type": "A", "rev": 79.319, "dram": "=G4*0.73", "nand": "=G4*0.27", "cogs": "=G4*(1-0.83)", "op": 60.543, "ni": 93.923, "da": 4.028},
  {"label": "2026Q3E", "type": "F", "keys": {"db": "q3_db", "dp": "q3_dp", "nb": "q3_nb", "np": "q3_np", "da": "da_q3"}},
  {"label": "2026Q4E", "type": "F", "keys": {"db": "q4_db", "dp": "q4_dp", "nb": "q4_nb", "np": "q4_np", "da": "da_q4"}},
 ],
 "inputs": [
  (S, "A. 匯率、股價"),
  ("fx", "美元兌韓元（2026–2028 平均）", "KRW/USD", *_SIL.vals("fx_krw"), "Derived", "Q2 2026 營收 79.32 兆 ≈ US$55.2B → 約 1,437"),
  ("px", "股價（2026-10-23 收盤）", "KRW", None, 1277000, None, "Verified", "markets/invest_map/data/prices.csv（tradersunion、digrin；Yonhap 當日 +4.4%）；MarketScreener 2026-10-23 收盤同（v0.4.2 為 10-07 收盤 1,723,000）"),
  (S, "B. 基準年剩餘季度（公司指引＋TrendForce）"),
  ("q3_db", "Q3 DRAM 位元季增", "%", 0.08, 0.10, 0.12, "Verified", "公司 Q2 法說：約 +10%"),
  ("q3_dp", "Q3 DRAM 混合 ASP 季增", "%", 0.08, 0.12, 0.16, "Derived", "TrendForce 一般 DRAM +13–18%；HBM 年約鎖價拉低混合值"),
  ("q3_nb", "Q3 NAND 位元季增", "%", 0.0, 0.02, 0.04, "Verified", "公司：低個位數"),
  ("q3_np", "Q3 NAND ASP 季增", "%", 0.10, 0.125, 0.20, "Interested-party", "TrendForce +10–15%；MS 實際約 +20%"),
  ("q4_db", "Q4 DRAM 位元季增", "%", 0.0, 0.05, 0.08, "Assumed", "公司：下半年位元成長優於上半年"),
  ("q4_dp", "Q4 DRAM 混合 ASP 季增", "%", 0.05, 0.10, 0.15, "Derived", "TrendForce +10–15%；BofA 約 +8%"),
  ("q4_nb", "Q4 NAND 位元季增", "%", 0.0, 0.05, 0.08, "Assumed", ""),
  ("q4_np", "Q4 NAND ASP 季增", "%", 0.10, 0.175, 0.20, "Interested-party", "TrendForce +15–20%"),
  ("da_q3", "Q3 折舊攤銷", "兆韓元", None, 4.3, None, "Assumed", "Q1 3.726、Q2 4.028 延伸"),
  ("da_q4", "Q4 折舊攤銷", "兆韓元", None, 4.6, None, "Assumed", ""),
  ("costdn_q", "單位成本季降", "%", 0.01, 0.03, 0.05, "Assumed", ""),
  ("opex_fix_q", "營業費用固定部分（每季）", "兆韓元", None, 1.5, None, "Derived", "Q1、Q2 2026 營業費用 3.97、5.24 對營收迴歸"),
  ("nonop_q", "下半年每季業外淨收益", "兆韓元", 0.5, 1.0, 2.0, "Assumed", "淨現金 × 約 2%"),
  (S, "C. 年度驅動（公司別）"),
  ("cb_1", "一般 DRAM 位元成長 第 1 年", "%", 0.05, 0.10, 0.15, "Derived", "BofA DRAM 供給 +19%；HBM 晶圓占比 22%→30% 吃掉增量"),
  ("cb_2", "一般 DRAM 位元成長 第 2 年", "%", 0.10, 0.15, 0.19, "Interested-party", "BofA 2028 供給 +19%"),
  ("nb_1", "NAND 位元成長 第 1 年", "%", 0.15, 0.20, 0.25, "Verified", "Micron：產業 2027 中 20% 段"),
  ("nb_2", "NAND 位元成長 第 2 年", "%", 0.15, 0.20, 0.25, "Verified", "同上"),
  ("sh25", "HBM 營收市占 2025", "%", *_SIL.vals("hbmsh_skh_25"), "Interested-party", "Counterpoint 3Q25 57%；出貨 62%"),
  ("sh26", "HBM 營收市占 2026", "%", *_SIL.vals("hbmsh_skh_26"), "Interested-party", "Counterpoint Q1 58%、Q2 50%；IDC Q1 56.4%"),
  ("sh27", "HBM 營收市占 2027", "%", *_SIL.vals("hbmsh_skh_27"), "Interested-party", "UBS 2027 位元：SK 39%、三星 41%、美光 20%"),
  ("sh28", "HBM 營收市占 2028", "%", *_SIL.vals("hbmsh_skh_28"), "Assumed", ""),
  ("other_y", "其他營收（每年）", "兆韓元", None, 1.6, None, "Verified", ""),
  ("w_c", "成本權重：一般 DRAM", "%", None, 0.45, None, "Assumed", ""),
  ("w_h", "成本權重：HBM", "%", None, 0.30, None, "Assumed", "HBM 每位元耗晶圓約 DDR5 的 3 倍（Micron）"),
  ("w_n", "成本權重：NAND", "%", None, 0.25, None, "Assumed", ""),
  ("costdn_y", "單位成本年降", "%", 0.05, 0.10, 0.15, "Assumed", ""),
  ("da_life", "新增資本支出轉為折舊的年數", "年", 4, 5, 6, "Assumed", ""),
  ("da_cogs", "折舊計入營業成本比例", "%", None, 0.9, None, "Assumed", ""),
  ("opex_fix_y", "營業費用固定部分（年，基準年）", "兆韓元", None, 6.0, None, "Derived", "＝每季 1.5 × 4"),
  ("opex_var", "營業費用變動率", "% 營收", 0.04, 0.047, 0.06, "Derived", "同上迴歸"),
  ("opex_g", "營業費用固定部分年增", "%", None, 0.10, None, "Assumed", ""),
  ("tax", "有效稅率", "%", 0.20, 0.23, 0.25, "Verified", "Q1 21.8%、Q2 23.5%"),
  ("rf", "淨現金收益率", "%", None, 0.02, None, "Assumed", ""),
  ("cap_0", "資本支出 基準年 2026", "兆韓元", 45, 48.7, 50, "Verified／Interested-party", "指引「40 兆後段」；共識 48.7"),
  ("cap_1", "資本支出 2027", "兆韓元", 55, 62.3, 75, "Interested-party", "MarketScreener 共識 62.3"),
  ("cap_2", "資本支出 2028", "兆韓元", 60, 71.2, 85, "Interested-party", "MarketScreener 共識 71.2"),
  ("cb_3", "一般 DRAM 位元成長 第 3 年（2029）", "%", 0.10, 0.15, 0.20, "Assumed", "龍仁一廠 2027 起投產"),
  ("nb_3", "NAND 位元成長 第 3 年", "%", 0.15, 0.20, 0.25, "Assumed", ""),
  ("sh29", "HBM 營收市占 2029", "%", *_SIL.vals("hbmsh_skh_29"), "Assumed", ""),
  ("cap_3", "資本支出 2029", "兆韓元", 60, 75, 90, "Assumed", "龍仁 2–6 期"),
  ("dps_3", "每股股利 2029", "KRW", None, 5800, None, "Assumed", "延續成長"),
  ("nm_op_q", "非記憶體事業每季營業利益（基準年推估季）", "—", None, 0, None, "Verified", "純記憶體公司，無"),
  ("nm_rev_0", "非記憶體事業營收 第 0 年", "—", None, 0, None, "Verified", "無"),
  ("nm_op_0", "非記憶體事業營業利益 第 0 年", "—", None, 0, None, "Verified", "無"),
  ("nm_rev_1", "非記憶體事業營收 第 1 年", "—", None, 0, None, "Verified", "無"),
  ("nm_op_1", "非記憶體事業營業利益 第 1 年", "—", None, 0, None, "Verified", "無"),
  ("nm_rev_2", "非記憶體事業營收 第 2 年", "—", None, 0, None, "Verified", "無"),
  ("nm_op_2", "非記憶體事業營業利益 第 2 年", "—", None, 0, None, "Verified", "無"),
  ("nm_rev_3", "非記憶體事業營收 第 3 年", "—", None, 0, None, "Verified", "無"),
  ("nm_op_3", "非記憶體事業營業利益 第 3 年", "—", None, 0, None, "Verified", "無"),
  (S, "D. 資本與股東回饋"),
  ("eq_last", "最後實績季末權益（2026-06-30）", "兆韓元", None, 262.693, None, "Verified", "6-K"),
  ("nc_last", "最後實績季末淨現金", "兆韓元", None, 69.4, None, "Verified", "Q2 簡報"),
  ("sh_last", "最後實績季末流通股數", "百萬股", None, 711.0755, None, "Verified", "6-K"),
  ("extra_eq", "基準年其餘增資（ADS 2026-07）", "兆韓元", None, 40.4, None, "Derived", "US$26.25B × 1,538.05"),
  ("extra_sh", "基準年其餘新股（ADS）", "百萬股", None, 17.79, None, "Verified", "424B4"),
  ("bb_0", "基準年其餘買回", "兆韓元", None, 40.0, None, "Verified", "2026-08-19 公告，註銷"),
  ("bb_sh_0", "基準年其餘買回股數", "百萬股", None, 24.07, None, "Verified", ""),
  ("div_0", "基準年其餘股利", "兆韓元", None, 0.54, None, "Derived", "每季 375 韓元 × 2 季"),
  ("capex_rem", "基準年其餘資本支出（現金）", "兆韓元", None, 29.2, None, "Derived", "全年 48.7 × 約 60%（上半年實付 18.3）"),
  ("dps_1", "每股股利 2027", "KRW", None, 4541, None, "Interested-party", "MarketScreener 共識"),
  ("dps_2", "每股股利 2028", "KRW", None, 5249, None, "Interested-party", "同上"),
  ("payout", "股東回饋 ÷ FCF", "%", 0.3, 0.5, 0.7, "Verified", "公司：2025–27 累計 FCF 的 50% 以上"),
  ("bb_px", "未來買回均價", "KRW", None, 1723000, None, "Assumed", "＝目前股價"),
  (S, "E. 估值倍數"),
  ("pb_1", "P/B 1.0x", "x", None, 1.0, None, "Assumed", "循環谷底倍數"),
  ("pb_2", "P/B 歷史中位", "x", None, 1.6, None, "Interested-party", "2010–2025 年底 0.81–2.00，中位約 1.6"),
  ("pb_3", "P/B 券商遠期", "x", None, 3.3, None, "Interested-party", "Shinhan 2026-07"),
  ("pe_1", "P/E 4x", "x", None, 4, None, "Assumed", "共識 2027E P/E 約 3.7x"),
  ("pe_2", "P/E 8x", "x", None, 8, None, "Assumed", ""),
  ("pe_3", "P/E 12x", "x", None, 12, None, "Assumed", ""),
 ],
 "recon": [
  ("2026E 營收 vs 共識", "=Y!C{rev}", 344.2, "MarketScreener 2026-10-02"),
  ("2026E 營業利益 vs 共識", "=Y!C{op}", 266.1, "MarketScreener；FnGuide 265.2"),
  ("2026E 淨利 vs 共識", "=Y!C{ni}", 256.3, "MarketScreener"),
  ("2026E HBM 營收 vs BofA", "=Y!C{hbm}", 45.7, "BofA 2026-09-13；CLSA 49.4"),
  ("2027E 營收（A）vs 共識", "=Y!D{rev}", 531.3, "MarketScreener"),
  ("2027E 營業利益（A）vs 共識", "=Y!D{op}", 417.2, "MarketScreener；FnGuide 391.8（08-14）"),
  ("2027E 淨利（A）vs 共識", "=Y!D{ni}", 336.2, "MarketScreener"),
  ("2028E 營業利益（A）vs 共識", "=Y!E{op}", 454.0, "MarketScreener；FnGuide 399.1"),
  ("2026Q3E 營收 vs 共識", "=Q!H4", 94.1, "FnGuide 2026-09-25"),
  ("2026Q3E 營業利益 vs 共識", "=Q!H10", 76.1, "FnGuide 74.1–78.1"),
 ],
 "sources": [
  ("SK hynix 424B4（2026-07）", "https://www.sec.gov/Archives/edgar/data/2120882/000119312526299963/d32785d424b4.htm"),
  ("SK hynix 6-K 半年報（2026-08-18）", "https://www.sec.gov/Archives/edgar/data/2120882/000119312526354777/d147827d6k.htm"),
  ("Q2 2026 法說逐字稿", "https://ng.investing.com/news/transcripts/earnings-call-transcript-sk-hynix-posts-record-q2-2026-results-as-shares-fall-93CH-2622277"),
  ("MarketScreener 共識", "https://www.marketscreener.com/quote/stock/SK-HYNIX-INC-6494929/finances/"),
  ("Herald Business：FnGuide、目標價", "https://biz.heraldcorp.com/article/10841314"),
  ("BofA HBM（Economy Tribune）", "https://www.economytribune.co.kr/news/articleView.html?idxno=3903874"),
  ("40 兆買回（Korea Herald）", "https://www.koreaherald.com/article/10845792"),
  ("UBS 2027 HBM 市占（Korea Herald）", "https://www.koreaherald.com/article/10828143"),
  ("companiesmarketcap P/B", "https://companiesmarketcap.com/sk-hynix/pb-ratio/"),
  ("TrendForce 4Q26 價格", "https://www.trendforce.com/presscenter/news/20260930-13258.html"),
  ("Citi 2027 見頂", "https://www.tipranks.com/news/memory-prices-will-peak-next-year-warns-citi-as-it-cuts-micron-mu-price-target-by-18"),
  ("Bernstein 2028", "https://www.notebookcheck.net/DRAM-crisis-Analysts-expect-drastic-price-drop-in-2028.1337992.0.html"),
 ],
 "readme_extra": ["vv0.4.1：Val 頁新增「十年倍數與規則值」區塊（倍數歷史、分位、規則值與隱含股價；原放在 markets/overview/extract.py），總覽與 HTML 展示層改讀此處；損益、EPS 與原有儲存格不變。", "v0.4：HBM 出貨位元隨情境變動（markets/memory）：2028–29 需求位元 × 情境顆數係數（B 0.8／0.85、C 0.6／0.6；與台積電、Advantest 同一套），供給與 2028 前的單價假設不變。2026–27 與 v0.3.1 相同；情境 A 全部相同；B、C 的 2028–29 HBM 營收下降，營業成本量指數隨 HBM 位元同步下降。", "v0.3.1：Tokenomics v19.0 同步。只有對帳頁的「每顆 GPU 的 NVIDIA 內容」由 51.6 改 42.8 千美元（橋接表 v0.2：GB300 機架基準價 $4.3M）；損益、淨值、估值與 v0.3 相同。", "v0.3：預測延長到 2029（三情境各加 2029 年走法）；前兩年數字與 v0.2 相同。", "v0.2 相對 v0.1：(1) 一般 DRAM、NAND 年均價變動改由共用季價路徑（Path）依會計年度平均算出，不再逐年手填；(2) 新增情境 C；(3) HBM 市場移到共用市場層 markets/memory。",
                  "已知簡化：營運資金未計；股數用期末；HBM 毛利率無法由公開資料分離。"],
}

# 估值規則的十年倍數歷史（2016–2025 年底；規則見 tools/valuation_rule.py；值原放在 markets/overview/extract.py，2026-10-23 移入）
CONFIG["rule"] = {
 "pb": {"sources": [
   {"label": "companiesmarketcap", "vals": [1.31, 1.59, .8832, 1.34, 1.56, 1.45, .81, 1.87, 1.62, 3.86], "tag": "Verified"},
   {"label": "marketcapof（2016–24）", "vals": [1.35, 1.60, .90, 1.34, 1.56, 1.45, .82, 1.83, 1.62, None], "tag": "Verified", "note": "2025 年底未提供"}]},
 "pe": {"sources": [
   {"label": "companiesmarketcap（年底 P/E，含虧損年）", "vals": [9.57, 4.61, 2.54, 30.2, 16.2, 9.01, 22.6, -10.6, 5.99, 10.5], "tag": "Verified", "use": False, "note": "2023 年虧損（P/E 為負）"},
   {"label": "採用（虧損年不計）", "vals": ['=IF(B{r}>0,B{r},"")'] * 10, "tag": "Derived", "note": "虧損年不計入分位"},
   {"label": "stockanalysis（核對，不採用）", "vals": [None, None, None, None, None, 9.38, 23.13, None, 6.06, 10.47], "tag": "Verified", "use": False, "note": "2021–2025 與 companiesmarketcap 差距 ≤ 8%"}]},
}

# v0.4.2（2026-10-25）：目標價與上檔空間（Andy 2026-10-25 要求；Val 頁最後一節、In 頁最後三列；既有儲存格不變）
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.4.2（2026-10-25）：Val 頁新增「0. 目標價與上檔空間」（放在 Val 既有內容之後，不位移既有列）：目標價＝In 頁權重 tp_wA／tp_wB／tp_wC（預設 A 0.3／B 0.5／C 0.2，Assumed，Andy 可改）× 三個 P/B 規則值（A＝中段、B＝見頂、C＝谷底）；既有儲存格的值不變。"]
CONFIG["tp_bias"] = "十年 P/B 含 2016–2023 年 AI 前（HBM 未放量）的低倍數期，偏誤方向為低估。"

# v0.5（2026-10-25）：目標價改為 DCF 與 Comparable 加權（Andy 11:28、11:53）；三情境延伸到 2032（收斂法）；現價更新到 10-23
MS_EST = "MarketScreener 預估頁（2026-10-25 擷取；頁面為 10-23 收盤）"
CONFIG["tp"] = {
 "roe_label": "ROE（近三年平均，2023–2025）", "roe": "=AVERAGE(-0.1565,0.3107,0.4414)",
 "roe_src": "stockanalysis 比率頁：2023 −15.65%、2024 31.07%、2025 44.14%（單一來源）",
 "nim10": "=AVERAGE(2.960/17.198,10.642/30.109,15.540/40.445,2.016/26.991,4.759/31.900,0.2233,0.0500,-0.2781,0.2989,0.4418)",
 "nim10_src": "2021–2025：stockanalysis 損益表淨利率（22.33%、5.00%、−27.81%、29.89%、44.18%）；2016–2020：公司年報營收與淨利（兆韓元，本次未重新開啟原文，單一來源）",
 "cons_tp": (3096288, 1535000, 4700000, "MarketScreener 共識頁：39 位分析師平均 3,096,288（最低 1,535,000、最高 4,700,000），2026-10-25 擷取"),
 "cons": [("cons_eps27", "共識 EPS 2027", 475721, MS_EST), ("cons_eps28", "共識 EPS 2028", 537981, MS_EST), ("cons_bvps28", "共識每股淨值 2028 年底", 1398106, MS_EST)],
 "cons_src": MS_EST,
 "peers": [("三星電子", "005930.KS", 1.26, "2026-10-23", "MarketScreener：現價 ÷ 2027E 每股淨值"),
           ("美光", "MU", 3.35, "2026-10-23", "MarketScreener：現價 ÷ FY2027E（2027-09）每股淨值"),
           ("鎧俠（Kioxia）", "285A.T", 2.22, "2026-10-09", "MarketScreener：現價 ÷ FY2028E（2028-03）每股淨值；頁面價格為 10-09"),
           ("南亞科", "2408.TW", 1.66, "2026-10-22", "MarketScreener：現價 ÷ 2027E 每股淨值")],
 "peer_note": "記憶體同業（三星、美光、鎧俠、南亞科；無 A 股）。前瞻 P/B＝現價 ÷ 下一個會計年度末（約 2027 年底）共識每股淨值；各家會計年度不同（美光 9 月、鎧俠 3 月）",
 "fwd_w": ("2/12", "10/12"), "fwd_note": "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × 2027＋10／12 × 2028（會計年度＝日曆年）",
 "method": ("DCF：股東現金流 2027–2029 用每股（淨利＋折舊－資本支出），2030–2032 收斂期用 EPS ×（1 − g／ROE），r 10%、g 4%、ROE 為 2023–2025 平均（含 2023 虧損年）；明確期＋永續，基準日 2026 年底。"
            "Comparable：三星、美光、鎧俠、南亞科前瞻 P/B 中位數 × 2028 年底每股淨值（情境加權）。"
            "收斂期（2030–2032）淨利率收斂到循環中段：A 十年平均 × 1.2、B 十年平均、C 十年平均 × 0.8（In 頁 cyc_mA／cyc_mC），永續期以 2032 中段水準為基礎。偏誤：2027–2029 仍為高峰期盈餘且權重大，DCF 仍可能偏高估；同業 P/B 為循環高峰前的市場倍數、且 SK 海力士淨值增速高於同業，Comparable 方向不明。"
            "情境對應：A＝順利、B＝平平、C＝不如預期（原為 TrendForce／BofA 型價格情境）。"),
}
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.5（2026-10-25）：Val 頁新增「0. 目標價」（DCF 與 Comparable 加權，tp_w_dcf 0.5）；原「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」，情境對應改為 A 75 分位、B 中位數、C 25 分位 × 各情境 2027 年底每股淨值。Y 頁三情境延伸到 2032（2030 起收斂法）。現價更新為 2026-10-23 收盤 1,277,000。2026–2028 EPS 與 v0.4.2 相同；2029 因修正共用工具「前一年」基期（原誤用 2027）而改變。收斂期淨利率收斂到循環中段（十年平均 × 1.2／1.0／0.8）。"]

# v0.5.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00、14:17；樣板 v0.4.1，數值不變）
CONFIG["version"] = "v0.5.1"
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.5.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁最後新增「上檔空間級距」「投資建議」兩列（依「上檔空間：加權目標價」固定規則）；README 頁最後新增投資建議與重點觀察三點。"]
CONFIG["obs"] = [
 "目標價 ₩3,087,845（DCF 3,709,130、Comparable 2,466,561 各半），比現價 1,277,000（2026-10-23）高 141.8%，級距 大於 50%；Comparable 單獨為 +93.2%。主因是倍數不是 EPS：本模型 2028 年底每股淨值 1,271,423 比共識 1,398,106 低 9.1%（前瞻 EPS 439,185 比共識低 16.8%），但現價只隱含約 1.0 倍前瞻 P/B，同業（三星、美光、鎧俠、南亞科）中位數 1.94 倍、共識目標價 3,096,288 隱含 2.21 倍；DCF 另被 2027–2029 高峰期自由現金流拉高（A 情境 DCF 為現價 5.10 倍）。加權目標價與共識目標價只差 0.3%，上檔空間幾乎全部來自現價偏低。",
 "情境分歧集中在 2028 年後：EPS 2028 A 714,923、B 331,149、C 195,351；DCF 情境價值 A 6,513,978、B 2,986,311、C 1,308,904，C 情境約等於現價（1.02 倍），現價隱含的循環路徑比 C 還悲觀。",
 "2027 A 營業利益對一般 DRAM 收入最敏感（±20% → ±13.0%），NAND ±30% → ±9.0%、HBM ±30% → ±6.0%；折現率 ±1pt 使加權目標價 −6.2%／+8.6%。下一個檢驗事件是 3Q26 財報（2026-10 下旬，日期未確認）的 DRAM 價格與 2027 年 HBM 供貨協議。"]   # 重點觀察（與 README.md 相同文字）

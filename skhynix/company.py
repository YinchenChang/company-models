# SK 海力士設定檔（樣板 v0.2）。只放參數；結構在 tools/memory_model/build.py 與 markets/memory/market.py
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))), "siliconomics"))
import reader as _SIL   # HBM 營收份額 sh25–sh29 讀 Siliconomics M24（v0.8 起；數值權威在本庫 In 頁）
S = "sec"
CONFIG = {
 "name": "SK 海力士", "version": "v0.5", "date": "2026-10-29",
 "unit": "兆韓元", "ccy_per_usdB": "fx/1000", "share_scale": 1000000,
 "base_fy": 2026, "fy_offset": 0, "fy_label": "{}", "fy_desc": "日曆年（K-IFRS 合併）",
 "thesis": "SK 海力士的盈餘由兩條價格線決定——每季議價的一般 DRAM 與年約鎖價的 HBM；股價取決於這個循環在哪一年見頂，以及見頂時累積了多少淨值。",
 "q_source": "SEC 424B4（2026-07）、6-K 半年報（2026-08-18）、3Q26 財報新聞稿與法說（2026-10-28）、公司公告",
 "q_note": "2026Q3 實績：營收 104.616、營業利益 83.334、淨利 39.776 兆（業外淨損 32.6 兆：匯兌損失 11.6、投資資產評價損失 22.9、其他淨收益 1.9）；DRAM 75.7%／NAND 24.3% 與毛利率 85.8% 為推估 [Derived：Q2 結構 × 法說位元與 ASP 季增（DRAM +10%／高 20% 段、NAND 小幅季減／中 20% 段）後依實際營收縮放；毛利＝營業利益＋營業費用迴歸值]；折舊 4.4＝EBITDA 87.7 − 營業利益。2026Q1 DRAM 77.3%、NAND 22%、毛利率約 79%；Q2 DRAM 73%、NAND 27%、毛利率 83%（公司簡報）。2025 季度拆分未取得。2025Q3、Q4 折舊為合計 7.113 平均分配 [Derived]",
 "hbm_share": {2025: "sh25", 2026: "sh26", 2027: "sh27", 2028: "sh28", 2029: "sh29"},
 "pb_names": ["P/B 1.0x", "P/B 歷史中位 1.6x", "P/B 券商遠期 3.3x"], "pe_vals": [4, 8, 12],
 "quarters": [
  {"label": "2025Q1", "type": "A", "rev": 17.639, "op": 7.441, "ni": 8.108, "da": 3.334},
  {"label": "2025Q2", "type": "A", "rev": 22.232, "op": 9.213, "ni": 6.996, "da": 3.443},
  {"label": "2025Q3", "type": "A", "rev": 24.449, "op": 11.383, "ni": 12.598, "da": 3.556},
  {"label": "2025Q4", "type": "A", "rev": 32.827, "op": 19.170, "ni": 15.246, "da": 3.557},
  {"label": "2026Q1", "type": "A", "rev": 52.576, "dram": "=F4*0.773", "nand": "=F4*0.22", "cogs": "=F4*(1-0.79)", "op": 37.610, "ni": 40.346, "da": 3.726},
  {"label": "2026Q2", "type": "A", "rev": 79.319, "dram": "=G4*0.73", "nand": "=G4*0.27", "cogs": "=G4*(1-0.83)", "op": 60.543, "ni": 93.923, "da": 4.028},
  {"label": "2026Q3", "type": "A", "rev": 104.616, "dram": "=H4*0.757", "nand": "=H4*0.243", "cogs": "=H4*(1-0.858)", "op": 83.334, "ni": 39.776, "da": 4.4},
  {"label": "2026Q4E", "type": "F", "keys": {"db": "q4_db", "dp": "q4_dp", "nb": "q4_nb", "np": "q4_np", "da": "da_q4"}},
 ],
 "inputs": [
  (S, "A. 匯率、股價"),
  ("fx", "美元兌韓元（2026–2028 平均）", "KRW/USD", *_SIL.vals("fx_krw"), "Derived", "Q2 2026 營收 79.32 兆 ≈ US$55.2B → 約 1,437"),
  ("px", "股價（2026-10-23 收盤）", "KRW", None, 1277000, None, "Verified", "markets/invest_map/data/prices.csv（tradersunion、digrin；Yonhap 當日 +4.4%）；MarketScreener 2026-10-23 收盤同（v0.4.2 為 10-07 收盤 1,723,000）"),
  (S, "B. 基準年剩餘季度（公司指引＋TrendForce）"),
  ("q4_db", "Q4 DRAM 位元季增", "%", 0.20, 0.25, 0.28, "Interested-party", "公司 3Q26 法說（2026-10-28）：4Q DRAM 位元季增中 20% 段（1z 製程、HBM4 放量）"),
  ("q4_dp", "Q4 DRAM 混合 ASP 季增", "%", 0.05, 0.10, 0.15, "Derived", "TrendForce +10–15%；BofA 約 +8%"),
  ("q4_nb", "Q4 NAND 位元季增", "%", 0.03, 0.05, 0.07, "Interested-party", "公司 3Q26 法說：4Q NAND 位元季增中個位數"),
  ("q4_np", "Q4 NAND ASP 季增", "%", 0.10, 0.175, 0.20, "Interested-party", "TrendForce +15–20%"),
  ("da_q4", "Q4 折舊攤銷", "兆韓元", None, 4.6, None, "Assumed", ""),
  ("costdn_q", "單位成本季降", "%", 0.01, 0.03, 0.05, "Assumed", ""),
  ("opex_fix_q", "營業費用固定部分（每季）", "兆韓元", None, 1.5, None, "Derived", "Q1、Q2 2026 營業費用 3.97、5.24 對營收迴歸"),
  ("nonop_q", "下半年每季業外淨收益", "兆韓元", 0.5, 1.0, 2.0, "Assumed", "淨現金 × 約 2%"),
  (S, "C. 年度驅動（公司別）"),
  ("cb_1", "一般 DRAM 位元成長 第 1 年", "%", 0.05, 0.10, 0.15, "Derived", "BofA DRAM 供給 +19%；HBM 晶圓占比 22%→30% 吃掉增量"),
  ("cb_2", "一般 DRAM 位元成長 第 2 年", "%", 0.10, 0.15, 0.19, "Interested-party", "BofA 2028 供給 +19%"),
  ("nb_1", "NAND 位元成長 第 1 年", "%", 0.15, 0.20, 0.25, "Derived", "美光 FQ4-26 法說：產業 2027、2028 皆中 20% 段；本模型取低於產業（SK 海力士優先高價值產品，3Q26 NAND 位元季減）。2026-10-28 來源複核：原標 Verified 不當"),
  ("nb_2", "NAND 位元成長 第 2 年", "%", 0.15, 0.20, 0.25, "Derived", "同上"),
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
  ("eq_last", "最後實績季末權益（2026-09-30，推估）", "兆韓元", None, "=262.693+39.776+40.4-40-0.27", None, "Derived／Assumed", "6-30 權益 262.693（6-K）＋ 3Q 淨利 39.776 ＋ ADS 增資 40.4（7 月）− 40 兆買回（假設 3Q 執行完畢；3Q 淨現金 +29.3 兆大致相容，現金橋未拆非現金業外損失，待季報核對）− 2Q 股利 0.27；3Q 資產負債表待季報"),
  ("nc_last", "最後實績季末淨現金（2026-09-30）", "兆韓元", None, 98.7, None, "Verified", "3Q26 法說：現金 114.5 兆、有息負債 15.8 兆"),
  ("sh_last", "最後實績季末流通股數（推估）", "百萬股", None, "=711.0755+17.79-24.07", None, "Derived", "6-30 股數（6-K）＋ ADS 新股 17.79 − 40 兆買回註銷 24.07"),
  ("extra_eq", "基準年其餘增資", "兆韓元", None, 0, None, "Verified", "ADS（2026-07）已含在 9-30 權益"),
  ("extra_sh", "基準年其餘新股", "百萬股", None, 0, None, "Verified", "ADS 已含在 9-30 股數"),
  ("bb_0", "基準年其餘買回（4Q）", "兆韓元", None, 20.0, None, "Verified", "公司 6-K（2026-10-28）：董事會 10-28 決議買回 20 兆（約 10 兆股東回饋、預計註銷；約 10 兆員工酬勞）；決議在 9-30 之後，不在 9-30 淨現金內，故不重複扣除。執行期間公司未載明，假設 4Q 完成（Assumed）"),
  ("bb_sh_0", "基準年其餘註銷股數（4Q）", "百萬股", None, "=10/1.277", None, "Derived", "註銷 10 兆 ÷ 現價 127.7 萬韓元；員工酬勞部分不註銷"),
  ("div_0", "基準年其餘股利（4Q 支付）", "兆韓元", None, "=20000*(711.0755+17.79-24.07)/1e6", None, "Derived", "3Q 股利每股 20,000 韓元（一般 10,000＋鎧俠投資收益特別 10,000；3Q26 法說）× 9-30 股數"),
  ("capex_rem", "基準年其餘資本支出（4Q 現金）", "兆韓元", 10, 14.6, 20, "Assumed", "原下半年 29.2 兆（全年 48.7 − 上半年實付 18.3 後約 60%）的一半；3Q 實付待季報"),
  ("dps_1", "每股股利 2027（現金支付年）", "KRW", None, 22050, None, "Derived", "公司 6-K（2026-10-28）：FY2026 現金股利 30 兆以上；已宣告 1Q 375＋2Q 375＋3Q 20,000 韓元（約 14.6 兆），餘額 ≥15.4 兆為 FY2026 年度股利（含第二筆鎧俠特別股利），2027 支付 ÷ 年底股數約 697 百萬股＝下限；2027 年內季度股利未計（政策未公告）。v0.5.1 為 MarketScreener 共識 4,541"),
  ("dps_2", "每股股利 2028", "KRW", None, 5249, None, "Interested-party", "MarketScreener 共識（10-25 擷取，未隨 3Q 政策更新）；股東回饋總額由 payout × FCF 決定，本值只影響股利與買回的分配"),
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
  ("2026Q3 營收（實績）vs 事前共識", "=Q!H4", 94.1, "FnGuide 2026-09-25"),
  ("2026Q3 營業利益（實績）vs 事前共識", "=Q!H10", 76.1, "FnGuide 74.1–78.1"),
 ],
 "sources": [
  ("SK hynix 424B4（2026-07）", "https://www.sec.gov/Archives/edgar/data/2120882/000119312526299963/d32785d424b4.htm"),
  ("SK hynix 6-K 半年報（2026-08-18）", "https://www.sec.gov/Archives/edgar/data/2120882/000119312526354777/d147827d6k.htm"),
  ("Q2 2026 法說逐字稿", "https://ng.investing.com/news/transcripts/earnings-call-transcript-sk-hynix-posts-record-q2-2026-results-as-shares-fall-93CH-2622277"),
  ("MarketScreener 共識", "https://www.marketscreener.com/quote/stock/SK-HYNIX-INC-6494929/finances/"),
  ("Herald Business：FnGuide、目標價", "https://biz.heraldcorp.com/article/10841314"),
  ("BofA HBM（Economy Tribune）", "https://www.economytribune.co.kr/news/articleView.html?idxno=3903874"),
  ("40 兆買回（Korea Herald）", "https://www.koreaherald.com/article/10845792"),
  ("SK hynix 6-K 3Q26 股東回饋（2026-10-28）", "https://www.sec.gov/Archives/edgar/data/0002120882/000119312526435008/d236990d6k.htm"),
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
            "情境對應：A＝樂觀、B＝基準、C＝保守（原為 TrendForce／BofA 型價格情境）。"),
}
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.5（2026-10-25）：Val 頁新增「0. 目標價」（DCF 與 Comparable 加權，tp_w_dcf 0.5）；原「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」，情境對應改為 A 75 分位、B 中位數、C 25 分位 × 各情境 2027 年底每股淨值。Y 頁三情境延伸到 2032（2030 起收斂法）。現價更新為 2026-10-23 收盤 1,277,000。2026–2028 EPS 與 v0.4.2 相同；2029 因修正共用工具「前一年」基期（原誤用 2027）而改變。收斂期淨利率收斂到循環中段（十年平均 × 1.2／1.0／0.8）。"]

# v0.5.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00、14:17；樣板 v0.4.1，數值不變）
CONFIG["version"] = "v0.5.1"
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.5.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁最後新增「上檔空間級距」「投資建議」兩列（依「上檔空間：加權目標價」固定規則）；README 頁最後新增投資建議與重點觀察三點。"]
CONFIG["obs"] = [
 "目標價 ₩3,087,845（DCF 3,709,130、Comparable 2,466,561 各半），比現價 1,277,000（2026-10-23）高 141.8%，級距 大於 50%；Comparable 單獨為 +93.2%。主因是倍數不是 EPS：本模型 2028 年底每股淨值 1,271,423 比共識 1,398,106 低 9.1%（前瞻 EPS 439,185 比共識低 16.8%），但現價只隱含約 1.0 倍前瞻 P/B，同業（三星、美光、鎧俠、南亞科）中位數 1.94 倍、共識目標價 3,096,288 隱含 2.21 倍；DCF 另被 2027–2029 高峰期自由現金流拉高（A 情境 DCF 為現價 5.10 倍）。加權目標價與共識目標價只差 0.3%，上檔空間幾乎全部來自現價偏低。",
 "情境分歧集中在 2028 年後：EPS 2028 A 714,923、B 331,149、C 195,351；DCF 情境價值 A 6,513,978、B 2,986,311、C 1,308,904，C 情境約等於現價（1.02 倍），現價隱含的循環路徑比 C 還悲觀。",
 "2027 A 營業利益對一般 DRAM 收入最敏感（±20% → ±13.0%），NAND ±30% → ±9.0%、HBM ±30% → ±6.0%；折現率 ±1pt 使加權目標價 −6.2%／+8.6%。下一個檢驗事件是 3Q26 財報（2026-10 下旬，日期未確認）的 DRAM 價格與 2027 年 HBM 供貨協議。"]   # 重點觀察（與 README.md 相同文字）

# v0.6（2026-10-29）：滾入 2026Q3 實績（2026-10-28 公布）與 4Q 指引
CONFIG["version"] = "v0.6"
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.6（2026-10-29）：滾入 2026Q3 實績（營收 104.616、營業利益 83.334、淨利 39.776 兆；業外淨損 32.6 兆＝匯兌損失 11.6＋投資資產評價損失 22.9−其他淨收益 1.9）；4Q DRAM 位元季增改為公司指引中 20% 段（原 5%）、NAND 中個位數；資產負債錨點移到 9-30（淨現金 98.7 兆為公司數字；權益、股數為推估，假設 40 兆買回於 3Q 執行完畢）；4Q 股東回饋改為 10-28 決議的買回 20 兆與 3Q 每股 20,000 韓元股利；2027 每股股利改為 FY2026 年度股利下限 22,050 韓元（6-K：全年現金股利 30 兆以上）。結果：2027 EPS（A）550,036 → 617,759（+12.3%）、加權目標價 3,087,845 → 3,408,692（+10.4%）。3Q 資產負債表待季報後校正。"]
CONFIG["obs"] = [
 "目標價 ₩3,408,692（DCF 4,288,814、Comparable 2,528,569 各半），比現價 1,277,000（2026-10-23）高 166.9%，級距 大於 50%；Comparable 單獨為 +98.0%。v0.6 滾入 3Q26 實績與 4Q 指引（DRAM 位元季增中 20% 段）後，2026 年第四季營收由 113.1 升到 140.7 兆，2027 起各年營收隨第四季基期上移約 9–10%；FY2026 年度股利（2027 支付）依公司「全年現金股利 30 兆以上」改為每股 22,050 韓元後，同樣的股東回饋總額中買回減少、股數多約 1.2%，加權目標價合計 +10.4%。上檔空間仍主要來自倍數：現價隱含約 1.0 倍前瞻 P/B，同業中位數 1.94 倍。",
 "情境分歧集中在 2028 年後：EPS 2028 A 818,444、B 375,667、C 222,213；DCF 情境價值（推到 12 個月後）A 7,641,482、B 3,401,631、C 1,477,771，C 情境為現價 1.16 倍。2026 年淨利 264 兆幾乎不變：3Q 業外淨損 32.6 兆（匯兌損失 11.6、投資資產評價損失 22.9、其他淨收益 1.9）抵銷了 4Q 上調。",
 "2027 A 營業利益對一般 DRAM 收入最敏感（±20% → ±13.7%），NAND ±30% → ±8.3%、HBM ±30% → ±5.4%。待校正：3Q 資產負債表（權益、股數為推估，假設 40 兆買回於 3Q 執行完畢）於季報公布後更新；下一個檢驗事件是 4Q26 財報（2027-01 下旬）的 DRAM 位元與 HBM4 占比。"]

# v0.6.1（2026-10-30）：In 頁版面（個股模型原則第 8 節；數值不變）
CONFIG["version"] = "v0.6.1"

# v0.6.2（2026-11-05）：只改展示層（個股 HTML 版面規格 v1.0；README 新增「展示層」「命題回覆」兩段，公式帶數字；其他頁數值不變）
CONFIG["in_layout_version"] = "v0.6.1"   # In 頁版面說明的版本標籤（v0.6.1 套用；之後升版不改）
CONFIG["version"] = "v0.6.2"
CONFIG["date"] = "2026-11-05"
CONFIG["a1_note"] = "；v0.6.2 只改展示層（README 新增展示層與命題回覆兩段、HTML 新版面），其他頁與 v0.6.1 相同"
CONFIG["pres"] = {
 "lead": "命題：SK 海力士的盈餘由每季議價的一般 DRAM 與年約鎖價的 HBM 兩條價格線決定；股價取決於這個循環在哪一年見頂，以及見頂時累積了多少淨值。三個問題如下。",
 "props": [
  "命題 1：哪條價格線決定盈餘：每季議價的一般 DRAM 與年約鎖價的 HBM，哪一條對營業利益的影響較大。",
  "命題 2：循環在哪一年見頂：三種價格情境下，盈餘在哪一年到高點、之後回落多少。",
  "命題 3：見頂時累積的淨值撐不撐得住股價：用同業倍數與現金流折現估算，每股價值是否高於現價。"],
 "obs": [
  "重點 1：合理價值{updesc}現價，{upsrc}。我們估每股約 {tpw|P} 韓元（DCF {dcftp|P} 占 {wdcf|0%}、Comparable〔相對倍數調整後〕{compadj|P} 占其餘），相對現價 {px|P} 為 {upw|+0%;−0%}，投資建議{rec}。把加權目標價的上檔拆成兩個來源（各自的權重 ×（目標價 − 現價））：DCF 貢獻 {dcfsh|0%}、Comparable 貢獻 {compsh|0%}；DCF 反映 {y27}–{y29} 高峰期的大量自由現金流（情境 A 的 DCF 價值為現價 {dcfpxA|0.0} 倍），對收斂期假設敏感。Comparable 比現價 {upcadj|+0%;−0%}，可拆成兩個因素相乘：每股淨值由 {y27} 年底成長到 {y28} 年底（情境加權）{bvg|+0%;−0%}，以及同業倍數（相對倍數調整後 {pmadj|0.00} 倍）比現價隱含的 {y27} 年底 P/B（{pxbv27|0.00} 倍）高 {pmg|0%;−0%}，{compdrv}。本模型的每股淨值只比市場共識{gbvdir} {gbvabs|0.0%}。",
  "重點 2：三種情境的淨利高點落在 {peakA}（情境 A）、{peakB}（情境 B）、{peakC}（情境 C）；分歧集中在 {y28} 以後：{y28} 每股盈餘情境 A 約 {epsA28|E}、B 約 {epsB28|E}、C 約 {epsC28|E} 韓元。最保守的情境 C，DCF 價值仍為現價的 {dcfpxC|0.00} 倍。",
  "重點 3：營業利益主要看{s_top} 價格。以 {s_year} 情境 {s_scn} 為例，{s_dram_l} 使營業利益變動 {s_dram|+0.0%;−0.0%}；{s_hbm_l} 為 {s_hbm|+0.0%;−0.0%}、{s_nand_l} 為 {s_nand|+0.0%;−0.0%}（各收入線的測試幅度不同，是敏感度測試，不是門檻）。"],
 "ans": [
  "一般 DRAM 和 HBM 哪條價格線對盈餘影響較大？答：{s_top}。以 {s_year} 情境 {s_scn} 為例，一般 DRAM 營收約 {convA27|#,##0.0} 兆韓元、HBM 約 {hbmA27|#,##0.0} 兆韓元（HBM 占記憶體營收 {hbmshA27|0%}）；{s_dram_l} 使營業利益變動 {s_dram|+0.0%;−0.0%}，{s_hbm_l} 只變動 {s_hbm|+0.0%;−0.0%}。HBM 年約鎖價讓這部分收入較穩，但金額小於一般 DRAM，盈餘的波動主要來自每季議價的一般 DRAM。",
  "循環在哪一年見頂？答：依本模型三種情境的價格路徑（情境設定，不是單一預測），淨利高點在情境 A {peakA}、情境 B {peakB}、情境 C {peakC}。情境 B 高點淨利約 {nipkB|#,##0} 兆韓元，{y28} 為 {niB28|#,##0} 兆韓元（比高點 {nichgB28|+0%;−0%}）；情境 C {y28} 為 {niC28|#,##0} 兆韓元（比高點 {nichgC28|+0%;−0%}）。",
  "見頂時累積的淨值撐不撐得住股價？答：依本模型{verdict}。現價 {px|P} 韓元只相當於 {y27} 年底每股淨值（情境加權）的 {pxbv27|0.00} 倍，記憶體同業同口徑的前瞻 P/B 中位數 {pmed|0.00} 倍；目標價把同業倍數（經相對倍數調整後 {pmadj|0.00} 倍）套在 12 個月後的前瞻每股淨值（{y28} 年底，情境加權）{fbv|P} 上，得 Comparable 目標價 {compadj|P}（比現價 {upcadj|+0.0%;−0.0%}）。加上 DCF 目標價 {dcftp|P}，加權目標價 {tpw|P}，比現價 {upw|+0.0%;−0.0%}，投資建議{rec}。{dcfall}（最低的情境 C 為現價 {dcfpxC|0.00} 倍）。"],
}

# v0.7（2026-11-05）：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；樣板 v0.5、tools/fcff_dcf.py；規格 docs/plan/個股DCF標準規格.md v1.0）。
# Comparable、情境權重、DCF 權重 50% 與投資建議規則不變；舊 DCF 參數 tp_r、tp_roe 刪除；淨現金收益率鍵 rf 改名 ncy（值不變）。
_DAM = "Damodaran Betas by Sector (Global) 2026 年 1 月（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/BetasGlobal.html；2026-11-05 讀取）"
_DAMUS = "Damodaran Betas by Sector (US) 2026 年 1 月（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/Betas.html；2026-11-05 讀取）"
_CRP = "Damodaran Country Default Spreads and Risk Premiums（2026-01-05 更新；https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html；2026-11-05 讀取）"
_KR10 = ("Trading Economics「South Korea 10-Year Yield Lowest Since June」2026-10-22：約 4.14%（https://tradingeconomics.com/south-korea/government-bond-yield/news/593996；2026-11-05 讀取）",
         "同站 2026-10-27「Lowest Since May」約 4.00%、2026-11-05 為 3.90%；取與股價（2026-10-23）最近的 10-22 值")
_SA = "stockanalysis.com SK hynix（KRX:000660）財報頁，2026-11-05 讀取；百萬韓元"
CONFIG["dcf"] = {
 "frac0": 1.0,
 "capex_note": "",
 "fd": dict(
  date="2026-12-31（基準年底；同 v0.6 DCF 基準日）",
  rf=(0.0414, _KR10[0], _KR10[1], "無風險利率（韓國 10 年期公債殖利率，2026-10-22）"),
  crp=(0.0064, "Verified", _CRP + "：South Korea（Aa2）國家風險溢酬 0.64%", "Andy 2026-11-05 20:45：非美國市場 ERP＝5%＋CRP，不設股權成本下限"),
  kd=("=rf+0.01", "Assumed", "rf ＋ 1.0pt 信用利差（公司債利率未逐筆查）", "負債占資本約 2%，對 WACC 影響可忽略"),
  debt=(15.8, "Verified", "3Q26 法說（2026-10-28）：有息負債 15.8 兆韓元（與 In 頁 nc_last 同一來源）", "只用於 WACC 權重與 β 重新加槓桿；評價日淨現金取 Y 頁基準年末"),
  lease=(1.963, "Verified", _SA + "：FY2025 Long-Term Leases 1,962,647", "2026-09-30 租賃負債未找到；公司「有息負債」是否含租賃未載明，視為不含"),
  inv=(0, "Assumed", "鎧俠等權益投資未計入（9-30 帳面值未取得；投資資產評價損益不進營業利益）", "偏保守；Sens 案例「淨現金不含權益投資」因此與基準相同"),
  mi=(0.151, "Verified", _SA + "：FY2025 Minority Interest 150,573", "—"),
  dps=("=dps_1", "Derived", "＝In 頁 dps_1（2027 年支付的每股股利，含 FY2026 年度股利）", "評價日 2026-12-31 到目標時點 2027-12-31"),
  t_tgt=(1.0, "目標價為評價日（2026-12-31）後 12 個月，與 v0.6 DCF 目標價同一時點"),
  fc_da=(None, "=(13.421531/32.765719+12.581537/66.19296+13.93013/97.146675)/3", None, "Derived",
         _SA + "：近三年（2023–2025）折舊攤銷 ÷ 營收平均＝41.0%／19.0%／14.3%（D&A 13,421,531／12,581,537／13,930,130；營收 32,765,719／66,192,960／97,146,675）",
         "記憶體樣板：明確期 2027–2029 用 Y 頁折舊；本值只作 2030 起收斂的長期穩態（含谷底年，近似循環平均）"),
  fc_cx=(None, "=(8.325138/32.765719+15.945534/66.19296+27.518924/97.146675)/3", None, "Derived",
         _SA + "：近三年（2023–2025）資本支出 ÷ 營收平均＝25.4%／24.1%／28.3%（8,325,138／15,945,534／27,518,924）",
         "明確期 2027–2029 用 Y 頁資本支出（In 頁 cap_1–cap_3）；本值只作 2030 起長期穩態；穩態資本支出 ≥ 折舊 ×（1＋g）（DCF 頁一致性檢查）"),
  fc_wc=(0.20, "=((6.737251+13.480659+0.392361-1.845537-1.680914)/32.765719+(13.019006+13.313937+0.302814-2.277347-3.977166)/66.19296+(18.199078+14.28939+0.300766-2.848455-6.277237)/97.146675)/3", 0.50, "Derived",
         _SA + "：近三年營運資金 ÷ 營收平均＝52.1%／30.8%／24.4%；營運資金＝應收＋存貨＋其他流動資產 − 應付 − 應計費用（不含其他流動負債）",
         "Sens 低／高 20%／50%（Assumed）；2023 谷底年營收低使比例偏高"),
 ),
 "beta_u": (1.87, "產業無槓桿 β：Damodaran Global Semiconductor（現金修正後；2026 年 1 月）",
            _DAM + "：Semiconductor 675 家，β 1.88、D/E 4.13%、稅率 8.46%、無槓桿 β 1.82、現金修正後 1.87", "非美國公司用 Global 表（Japan／Emerging 表未取得）；用現金修正值：淨現金另計，β 只對營業資產"),
 "beta_extra": [("beta_u_us", "對照：Damodaran US Semiconductor 無槓桿 β（現金修正後；β 低案例）", 1.50, _DAMUS + "：Semiconductor 66 家，現金修正後無槓桿 β 1.50")],
 "beta_lo": "=beta_u_us*(1+(1-tax_w)*beta_de)", "beta_hi": None,
 "beta_lohi_note": "低＝Damodaran US Semiconductor 1.50 重新加槓桿（Analogy）；高＝與低對稱（2 × 基準 − 低，Assumed）。個股迴歸 β 未取得",
}
CONFIG["tp"]["method"] = ("DCF（v0.7 標準版）：無槓桿自由現金流 FCFF＝NOPAT（合併營業利益 ×（1 − 稅率））＋折舊 − 資本支出 − 營運資金增加，2027–2029 用 Y 頁折舊與資本支出，2030–2032 收斂期由淨利還原營業利益、折舊與資本支出比例收斂到近三年平均；"
                          "以 WACC（股權成本＝韓國 10 年期公債 rf ＋ β ×（ERP 5%＋韓國 CRP 0.64%），β 為 Damodaran Global 半導體無槓桿 β 重新加槓桿）折到 2026 年底，終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g），g＝3%、RONIC＝WACC；"
                          "加淨現金 ÷ 股數，以股權成本推到 12 個月後並扣 2027 年股利（DCF 頁、tools/fcff_dcf.py）。"
                          "Comparable：三星、美光、鎧俠、南亞科前瞻 P/B 中位數 × 2028 年底每股淨值（情境加權）。收斂期（2030–2032）淨利率收斂到循環中段：A 十年平均 × 1.2、B 十年平均、C 十年平均 × 0.8（In 頁 cyc_mA／cyc_mC）。"
                          "偏誤：2027–2029 高峰期現金流權重大；永續期以 2032 循環中段為基礎；同業 P/B 為循環高峰前的市場倍數，Comparable 方向不明。情境對應：A＝樂觀、B＝基準、C＝保守。")
CONFIG["version"] = "v0.7"
CONFIG["a1_note"] = "；v0.6.2 只改展示層（README 展示層與命題回覆、HTML 新版面）；v0.7（2026-11-05）模型改版：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數），Comparable、情境權重、DCF 權重 50% 與投資建議規則不變"
CONFIG["obs_tpl"] = {
 2: "情境分歧集中在 2028 年後：EPS 2028 A {epsA28|E}、B {epsB28|E}、C {epsC28|E}；DCF 情境價值（v0.7 標準版 FCFF＋WACC，推到 12 個月後）A {dcfA|P}、B {dcfB|P}、C {dcfC|P}，C 情境為現價 {dcfpxC|0.00} 倍。2026 年淨利與 v0.6 前相近：3Q 業外淨損 32.6 兆（匯兌損失 11.6、投資資產評價損失 22.9、其他淨收益 1.9）抵銷了 4Q 上調。",
 3: "2027 A 營業利益對一般 DRAM 收入最敏感（+20% → {s_dram|+0.0%;−0.0%}），NAND +30% → {s_nand|+0.0%;−0.0%}、HBM +30% → {s_hbm|+0.0%;−0.0%}；WACC −1pt／＋1pt 使加權目標價 {s_wlo|+0.0%;−0.0%}／{s_whi|+0.0%;−0.0%}。待校正：3Q 資產負債表（權益、股數為推估，假設 40 兆買回於 3Q 執行完畢）於季報公布後更新；下一個檢驗事件是 4Q26 財報（2027-01 下旬）的 DRAM 位元與 HBM4 占比。",
}
CONFIG["readme_tail"] = ["v0.7（2026-11-05）：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；記憶體樣板 v0.5、tools/fcff_dcf.py；docs/plan/個股DCF標準規格.md v1.0）。新增 DCF 頁（三情境與 Wattonomics 錨定路徑的 FCFF 明細、檢查列）；Val 0.1 節改為折現率與 DCF 彙總（列數不變，0.2 節以後列號不變）；Sens 目標價參數改為標準案例（WACC、β、ERP、g、RONIC、資本支出、營運資金、DCF 權重、終值替代法）；In 頁新增 M2 節（rf、β、ERP、CRP、資本結構、折舊／資本支出／營運資金穩態比例），刪除 tp_r、tp_roe，淨現金收益率鍵 rf 改名 ncy（值不變）。Y、Q、Mkt、Recon 頁與 v0.6.2 相同。"]   # 接在 README 最後（不位移既有列）
CONFIG["readme_extra"] = [("已知簡化：Y 頁淨現金與 ROE 未計營運資金（只在 DCF 以 fc_wc × 營收增量計入）；鎧俠等權益投資未計入 DCF 淨現金；股數用期末；HBM 毛利率無法由公開資料分離。" if t.startswith("已知簡化") else t) for t in CONFIG["readme_extra"]]

# v0.7.1（2026-11-05）：β 改自算產業 β（Andy 2026-11-05 22:55；docs/plan/自算產業β_20261105.md 對照表）：KR 半導體 同業中位數 1.72，Blume 調整 0.67β＋0.33，依本公司 D/E 重新加槓桿；取代 Damodaran 表。其他模型內容不變。
_BSRC = "自算產業 β：KR 半導體 同業 4 家（SK 海力士、三星電子、韓美半導體、Samsung Electro-Mechanics）中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 韓國綜合指數 ^KS11 迴歸，Hamada 去槓桿、現金修正）"
CONFIG["dcf"].update({
 "beta_u": ("=0.67*beta_u_med+0.33", "產業無槓桿 β（自算 KR 半導體，Blume 調整後；採用值）", _BSRC + "；＝0.67 × 中位數 beta_u_med ＋ 0.33",
            "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；用現金修正值：淨現金另計，β 只對營業資產", "Derived"),
 "beta_extra": [("beta_u_med", "產業無槓桿 β 中位數（自算 KR 半導體，Blume 調整前；β 低／高案例用）", 1.72, _BSRC + "；β 低案例＝中位數 × 0.8 再 Blume 調整；高案例＝中位數未調整（Analogy）", "Derived")],
 "beta_lo": "=(0.67*0.8*beta_u_med+0.33)*(1+(1-tax_w)*beta_de)", "beta_hi": "=beta_u_med*(1+(1-tax_w)*beta_de)",
 "beta_lohi_note": "低＝本組 Blume 前中位數 × 0.8 再 Blume、高＝本組 Blume 前中位數（未調整），重新加槓桿（Analogy）；個股迴歸 β 未取得",
 "beta_src": _BSRC + "；Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；",
 "beta_obs": "自算 KR 半導體產業無槓桿 β {betau|0.000}（同業中位數 Blume 調整）",
 "beta_sens": "β 低／高案例（本組中位數 × 0.8 再 Blume／中位數未調整）為 {s_blo|+0.0%;−0.0%}／{s_bhi|+0.0%;−0.0%}，",
})
CONFIG["version"] = "v0.7.1"
CONFIG["a1_note"] = CONFIG["a1_note"] + "；v0.7.1（2026-11-05）：β 改自算產業 β（KR 半導體，Blume 調整）"
CONFIG["readme_tail"] = CONFIG["readme_tail"] + ["v0.7.1（2026-11-05）：β 改自算產業 β（Andy 2026-11-05 22:55）：KR 半導體 同業無槓桿 β 中位數 {betaum|0.00}（Yahoo Finance 月報酬 2021-11～2026-10 對 韓國綜合指數 ^KS11 迴歸，Hamada 去槓桿、現金修正），Blume 調整後 {betau|0.000}，依市值 D/E 重新加槓桿為 {beta|0.00}，取代 Damodaran 表；β 低／高案例改為中位數 × 0.8 再 Blume／中位數未調整。"
    "評價日 2026-12-31 不變（已是統一評價日）。"
    "新舊對照（v0.7 → v0.7.1）：β 1.90 → {beta|0.00}，WACC 14.6% → {wacc|0.0%}，DCF 目標價 2,178,721 → {dcftp|P}，加權目標價 2,320,245 → {tpw|P}，上檔空間 +81.7% → {upw|+0.0%;−0.0%}，投資建議 強力買進 → {rec}。其他模型內容不變。"]

_m = CONFIG["tp"]["method"]; assert "β 為 Damodaran Global 半導體無槓桿 β 重新加槓桿" in _m
CONFIG["tp"]["method"] = _m.replace("β 為 Damodaran Global 半導體無槓桿 β 重新加槓桿", "β 為自算 KR 半導體產業無槓桿 β（Blume 調整）重新加槓桿")

# v0.7.2（2026-11-05）：工作單第二輪修正（Andy 2026-11-05 23:16 定案；共用模組 tools/fcff_dcf.py v2、記憶體樣板 v0.5.2）
# (1) β 改「記憶體」組（docs/plan/自算產業β_20261105.md 第二輪修正新增組：MU、WDC、SK 海力士、三星電子、南亞科 5 家，無槓桿中位數 2.294 → Blume 1.867）；β 低／高＝β_u ∓ 0.3 重新加槓桿
# (2) 資產負債表日 2026-09-30（3Q26 財報新聞稿：現金 114.5、借款 15.8 兆）；淨現金滾到評價日 2026-12-31＝9-30 淨現金 ＋ 4Q 模型 FCFF − 已宣告股利與買回
# (3) 明確期折舊／資本支出照舊讀 Y 頁逐年列（模組 v2 da／cx），收斂年度資本支出 ≥ 折舊 ×（1＋g）；(4) Sens 新增「美元口徑折現」
_BSRC2 = ("自算產業 β：記憶體 同業 5 家（美光 MU、Western Digital WDC、SK 海力士 000660、三星電子 005930、南亞科 2408；SanDisk、Kioxia 上市不足 5 年不入中位數）中位數"
          "（Yahoo Finance 月報酬 2021-11～2026-10，各自對本地大盤迴歸，Hamada 去槓桿、現金修正）")
_US10 = ("美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；2026-11-05 讀取）",
         "FRED DGS10 同序列；與股價同日（2026-10-23）；只用於 Sens「美元口徑折現」（工作單第二輪修正，Andy 2026-11-05 23:16）")
_REL3Q = "SK hynix 3Q26 財報新聞稿（2026-10-28；https://www.placera.se/pressmeddelanden/sk-hynix-announces-3q26-financial-results-20261028，2026-11-05 讀取）"
_SAQ = "stockanalysis.com SK hynix（KRX:000660）季度資產負債表（https://stockanalysis.com/quote/krx/000660/financials/balance-sheet/?p=quarterly，2026-11-05 讀取；最新一季 2026-06-30，百萬韓元）"
CONFIG["dcf"].update({
 "beta_u": ("=0.67*beta_u_med+0.33", "產業無槓桿 β（自算記憶體組，Blume 調整後；採用值）", _BSRC2 + "；＝0.67 × 中位數 beta_u_med ＋ 0.33",
            "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；記憶體組為第二輪修正新增（2026-11-05 23:20，取代 KR 半導體 1.482）；用現金修正值：淨現金另計，β 只對營業資產", "Derived"),
 "beta_extra": [("beta_u_med", "產業無槓桿 β 中位數（自算記憶體組，Blume 調整前；只作來源對照）", 2.294, _BSRC2, "Derived"),
                ("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", 0.3, "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3；取代 v0.7.1 的中位數 × 0.8 再 Blume／中位數未調整", "Assumed")],
 "beta_du": True,
 "beta_lo": "=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "beta_hi": "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)",
 "beta_lohi_note": "Sens 低／高（Analogy；v0.7.2 起）：β_u ∓ beta_du（0.3），各自同式重新加槓桿（tools/fcff_dcf.py v2 write_sens）；個股迴歸 β 未取得",
 "beta_src": _BSRC2 + "；Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；",
 "beta_obs": "自算記憶體產業無槓桿 β {betau|0.000}（同業中位數 Blume 調整）",
 "beta_sens": "β 低／高案例（β_u ∓ 0.3 各自重新加槓桿）為 {s_blo|+0.0%;−0.0%}／{s_bhi|+0.0%;−0.0%}，",
 "beta_note": "自算產業 β 重新加槓桿",
 "stub": dict(bs_date="2026-09-30", val_date="2026-12-31", bs_src=_REL3Q + "：現金 114.5 兆、借款 15.8 兆（初步、未經核閱）；租賃、少數股權取最新一季 2026-06-30（" + _SAQ + "）",
              div=("=div_0+bb_0", "Derived", "＝In 頁 div_0（3Q 股利每股 20,000 韓元 × 9-30 股數，2026-10-28 宣告、4Q 支付）＋ bb_0（10-28 董事會決議買回 20 兆）",
                   "只扣已宣告者；8 月決議的 40 兆買回依新聞稿於 10-28 前一週完成，部分可能在 9-30 之後執行，本模型沿用「3Q 執行完畢」假設（季報公布後核對；未扣部分會高估淨現金）")),
 "nc_bs": "114.5-15.8-lease_v-mi_v", "nc_bs_note": "2026-09-30（3Q26 財報新聞稿）：現金 114.5 − 借款 15.8 − 租賃 − 少數股權（後兩者為 6-30 值）",
})
CONFIG["dcf"]["stub"]["fcff"] = "(Q!I10+nm_op_q)*(1-tax)+Q!I13-capex_rem-fc_wc*(Q!I4-Q!E4)"   # 4Q26 模型 FCFF：營業利益 ×（1 − 稅率）＋折舊 − 4Q 資本支出 − fc_wc ×（2026Q4 − 2025Q4 營收；近四季營收增量）
CONFIG["dcf"]["fd"] = dict(CONFIG["dcf"]["fd"],
    rf_us=(0.05, _US10[0], _US10[1]),
    cash=(114.5, "Verified", _REL3Q + "：2026-09-30 現金 114.5 兆韓元（初步）", "淨現金 98.7＝現金 114.5 − 借款 15.8（與 In 頁 nc_last 同一來源）；DCF 淨現金以本列計，再以 stub 期滾到評價日"),
    debt=(15.8, "Verified", _REL3Q + "：2026-09-30 借款 15.8 兆韓元", "WACC 權重、β 重新加槓桿與評價淨現金皆用本值"),
    lease=("=(2000.044+526.461)/1000", "Verified", _SAQ + "：Long-Term Leases 2,000,044＋Current Portion of Leases 526,461", "9-30 租賃負債未揭露（季報約 11 月中），用最新一季 6-30 值；視為不含在借款內"),
    mi=(0.313, "Verified", _SAQ + "：Minority Interest 312,618", "9-30 未揭露，用 6-30 值"),
)
CONFIG["obs_tpl"] = dict(CONFIG["obs_tpl"])
CONFIG["obs_tpl"][1] = ("目標價 {tpw|P}（DCF {dcftp|P}、Comparable（相對倍數調整後）{compadj|P}；DCF 權重 {wdcf|0%}），相對現價 {px|P}（2026-10-23）為 {upw|+0.0%;−0.0%}，投資建議{rec}。"
                        "{upsrc}：加權目標價的上檔中 DCF 貢獻 {dcfsh|0%}、Comparable 貢獻 {compsh|0%}（各自權重 ×（目標價 − 現價））；Comparable 單獨（市場主流同業法、未調整）{upc|+0.0%;−0.0%}。")
CONFIG["obs_tpl"][3] = CONFIG["obs_tpl"][3].replace("待校正：3Q 資產負債表（權益、股數為推估，假設 40 兆買回於 3Q 執行完畢）於季報公布後更新",
                                                    "待校正：3Q 季報（權益、股數為推估；8 月 40 兆買回依新聞稿於 10-28 前一週完成，本模型假設 3Q 執行完畢）公布後更新")
CONFIG["version"] = "v0.7.2"
CONFIG["a1_note"] = CONFIG["a1_note"] + "；v0.7.2（2026-11-05）：第二輪修正（記憶體組 β、資產負債表日 9-30 滾到評價日、β 敏感度 ±0.3、Sens 美元口徑折現）"
CONFIG["readme_tail"] = CONFIG["readme_tail"] + ["v0.7.2（2026-11-05）：工作單第二輪修正（Andy 2026-11-05 23:16 定案；共用模組 tools/fcff_dcf.py v2、記憶體樣板 v0.5.2）。"
    "(1) β 改「記憶體」組（MU、WDC、SK 海力士、三星電子、南亞科 5 家無槓桿中位數 {betaum|0.000}，Blume 調整 {betau|0.000}），依市值 D/E 重新加槓桿為 {beta|0.00}；β 敏感度改 β_u ∓ 0.3 重新加槓桿。"
    "(2) 資產負債表日 2026-09-30（3Q26 財報新聞稿：現金 114.5、借款 15.8 兆）：DCF 淨現金與投資 {dcfnc|#,##0.0} 兆，滾到評價日 2026-12-31＝＋4Q 模型 FCFF {stubB|#,##0.0} 兆 − 已宣告股利與買回 {stubdiv|#,##0.0} 兆＝{ncrB|#,##0.0} 兆。"
    "(3) 明確期 2027–2029 折舊／資本支出照舊讀 Y 頁逐年列（模組 v2），2030 起收斂且資本支出 ≥ 折舊 ×（1＋g）。(4) Sens 新增「美元口徑折現」：加權目標價 {usd|P}（{usdp|+0.0%;−0.0%}）。"
    "新舊對照（v0.7.1 → v0.7.2）：β 1.51 → {beta|0.00}，WACC 12.5% → {wacc|0.0%}，DCF 目標價 2,367,571 → {dcftp|P}，加權目標價 2,414,670 → {tpw|P}，上檔空間 +89.1% → {upw|+0.0%;−0.0%}，投資建議 強力買進 → {rec}。"]

CONFIG["tp"]["method"] = CONFIG["tp"]["method"].replace("β 為自算 KR 半導體產業無槓桿 β（Blume 調整）重新加槓桿", "β 為自算記憶體產業無槓桿 β（Blume 調整）重新加槓桿")
assert "β 為自算記憶體產業無槓桿 β（Blume 調整）重新加槓桿" in CONFIG["tp"]["method"]

# v0.7.3（2026-11-06）：第三輪（合併前）修正（記憶體樣板 v0.5.3；共用模組 tools/fcff_dcf.py v2 不變）
# (1) 現價統一 2026-11-04 收盤（Yahoo Finance 日線收盤）；同業前瞻 P/B 以同日收盤 ÷ 共識每股淨值重算（Kioxia 與 SanDisk／Kioxia 模型同一 11-04 口徑）
# (2) bb_px、bb_sh_0 改公式連 px；(3) README 版本沿革改固定文字；(4) 自算產業 β 來源改引 repo 文件 docs/plan/自算產業β_20261105.md
_PX = "Yahoo Finance 日線收盤 2026-11-04"
def _set_in(key, label=None, base=None, tag=None, src=None):
    i = next(j for j, x in enumerate(CONFIG["inputs"]) if x[0] == key)
    x = list(CONFIG["inputs"][i])
    if label is not None: x[1] = label
    if base is not None: x[4] = base
    if tag is not None: x[6] = tag
    if src is not None: x[7] = src
    CONFIG["inputs"][i] = tuple(x)
_set_in("px", label="股價（2026-11-04 收盤）", base=1396000, tag="Verified", src=_PX + "（第三輪修正統一價格日）；v0.7.2 為 2026-10-23 收盤 1,277,000")
_set_in("bb_px", base="=px", tag="Derived", src="＝目前股價（公式連 In 頁 px）；v0.7.2 為寫死的 10-07 舊價 1,723,000")
_set_in("bb_sh_0", base="=10/px*1e6", src="註銷 10 兆 ÷ 現價 px（百萬股；公式連 px）；v0.7.2 為 10 ÷ 127.7 萬韓元")
_MSP = "MarketScreener 共識每股淨值（2026-10-25 擷取，未更新）"
CONFIG["tp"]["peers"] = [
    ("三星電子", "005930.KS", "=218000/(215500/1.26)", "2026-11-04", f"11-04 收盤 218,000（{_PX}）÷ 2027E 每股淨值（＝10-23 收盤 215,500 ÷ 頁面 P/B 1.26；{_MSP}）"),
    ("美光", "MU", "=795.23/(877.55/3.35)", "2026-11-04", f"11-04 收盤 795.23（{_PX}）÷ FY2027E（2027-09）每股淨值（＝10-23 收盤 877.55 ÷ 頁面 P/B 3.35；{_MSP}）"),
    ("鎧俠（Kioxia）", "285A.T", "=10735/7808", "2026-11-04", f"11-04 收盤 ¥10,735（{_PX}）÷ FY2028E（2028-03）每股淨值 ¥7,808（{_MSP}；sndk/、kioxia/ 模型同值）；v0.7.2 為 10-09 頁面 P/B 2.22"),
    ("南亞科", "2408.TW", "=399.5/248.9", "2026-11-04", f"11-04 收盤 399.5（{_PX}）÷ 2027E 每股淨值 248.9（MarketScreener 2026-11-06 讀取；頁面 10-22 收盤 414 對應 P/B 1.66）")]
CONFIG["tp"]["peer_note"] = CONFIG["tp"]["peer_note"] + "；第三輪修正（v0.7.3）起報價統一 2026-11-04 收盤（Yahoo Finance），共識每股淨值不變"
_fd = CONFIG["dcf"]["fd"]
_fd["rf"] = (_fd["rf"][0], _fd["rf"][1], _fd["rf"][2].replace("取與股價（2026-10-23）最近的 10-22 值", "取 v0.7.2 現價日（2026-10-23）最近的 10-22 值；第三輪修正現價改 2026-11-04，rf 未改（不在該輪範圍）"), _fd["rf"][3])
_fd["rf_us"] = (_fd["rf_us"][0], _fd["rf_us"][1], _fd["rf_us"][2].replace("與股價同日（2026-10-23）", "v0.7.2 現價日（2026-10-23）值，第三輪修正未改"))
assert "rf 未改" in _fd["rf"][2] and "第三輪修正未改" in _fd["rf_us"][2]
CONFIG["dcf"]["stub"]["fcff_label"] = "2026Q4E 模型季度 FCFF（Q 頁：營業利益 ×（1 − 稅率）＋折舊 − 4Q 資本支出 capex_rem − fc_wc × 近四季營收增量）"
CONFIG["obs_tpl"][1] = CONFIG["obs_tpl"][1].replace("相對現價 {px|P}（2026-10-23）", "相對現價 {px|P}（2026-11-04）")
assert "（2026-11-04）" in CONFIG["obs_tpl"][1]
_FROZEN = {"v0.7.1（": 'v0.7.1（2026-11-05）：β 改自算產業 β（Andy 2026-11-05 22:55）：KR 半導體 同業無槓桿 β 中位數 1.72（Yahoo Finance 月報酬 2021-11～2026-10 對 韓國綜合指數 ^KS11 迴歸，Hamada 去槓桿、現金修正），Blume 調整後 1.482，依市值 D/E 重新加槓桿為 1.51，取代 Damodaran 表；β 低／高案例改為中位數 × 0.8 再 Blume／中位數未調整。評價日 2026-12-31 不變（已是統一評價日）。新舊對照（v0.7 → v0.7.1）：β 1.90 → 1.51，WACC 14.6% → 12.5%，DCF 目標價 2,178,721 → 2,367,571，加權目標價 2,320,245 → 2,414,670，上檔空間 +81.7% → +89.1%，投資建議 強力買進 → 強力買進。其他模型內容不變。',
           "v0.7.2（": 'v0.7.2（2026-11-05）：工作單第二輪修正（Andy 2026-11-05 23:16 定案；共用模組 tools/fcff_dcf.py v2、記憶體樣板 v0.5.2）。(1) β 改「記憶體」組（MU、WDC、SK 海力士、三星電子、南亞科 5 家無槓桿中位數 2.294，Blume 調整 1.867），依市值 D/E 重新加槓桿為 1.90；β 敏感度改 β_u ∓ 0.3 重新加槓桿。(2) 資產負債表日 2026-09-30（3Q26 財報新聞稿：現金 114.5、借款 15.8 兆）：DCF 淨現金與投資 95.9 兆，滾到評價日 2026-12-31＝＋4Q 模型 FCFF 40.5 兆 − 已宣告股利與買回 34.1 兆＝102.3 兆。(3) 明確期 2027–2029 折舊／資本支出照舊讀 Y 頁逐年列（模組 v2），2030 起收斂且資本支出 ≥ 折舊 ×（1＋g）。(4) Sens 新增「美元口徑折現」：加權目標價 2,258,319（−1.3%）。新舊對照（v0.7.1 → v0.7.2）：β 1.51 → 1.90，WACC 12.5% → 14.6%，DCF 目標價 2,367,571 → 2,114,216，加權目標價 2,414,670 → 2,287,993，上檔空間 +89.1% → +79.2%，投資建議 強力買進 → 強力買進。'}
CONFIG["readme_tail"] = [next((v for k, v in _FROZEN.items() if t.startswith(k)), t) for t in CONFIG["readme_tail"]]
CONFIG["version"] = "v0.7.3"
CONFIG["date"] = "2026-11-06"
CONFIG["a1_note"] = CONFIG["a1_note"] + "；v0.7.3（2026-11-06）：第三輪修正（現價與同業報價統一 2026-11-04、買回價連現價、版本沿革固定文字）"
CONFIG["readme_tail"] = CONFIG["readme_tail"] + ["v0.7.3（2026-11-06）：第三輪（合併前）修正（記憶體樣板 v0.5.3）。"
    "(1) 現價改 2026-11-04 收盤 1,396,000（Yahoo Finance 日線收盤；v0.7.2 為 10-23 的 1,277,000）；同業前瞻 P/B 以同日收盤 ÷ 共識每股淨值重算（三星 1.26 → 1.27、美光 3.35 → 3.04、Kioxia 2.22（10-09 頁面）→ 1.37、南亞科 1.66 → 1.61；Kioxia 與 SanDisk、Kioxia 模型同一 11-04 口徑），同業中位數 1.94 → 1.49。"
    "(2) 未來買回均價 bb_px（原寫死 10-07 舊價 1,723,000）與 4Q 註銷股數 bb_sh_0 改公式連 px。"
    "(3) README 版本沿革列改為當版固定文字（原為活公式，會顯示現值）；Stub 期說明改依實際口徑（2026Q4E 模型季度 FCFF）；情境驅動對照收斂法列名中性化、循環中段係數情境 B 佔位 0 → 1；HTML DCF 明細只有比例列用百分比。"
    "(4) 自算產業 β 來源改引 repo 文件 docs/plan/自算產業β_20261105.md（值不變）。"
    "新舊對照（v0.7.2 → v0.7.3）：β 1.90 → 1.89，DCF 目標價 2,114,216 → 2,111,490，Comparable（相對倍數調整後）2,461,770 → 2,027,217，加權目標價 2,287,993 → 2,069,354，上檔空間 +79.2% → +48.2%，投資建議 強力買進 → 買進。"]


# v0.7.4（2026-11-06，工作單 px1105；Andy 2026-11-06 07:28、03:14）：現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守（記憶體樣板同步改名）
# 同業前瞻 P/B＝現行倍數（11-04 基準）× 11-05 收盤 ÷ 11-04 收盤（共識每股淨值不變）；rf、共識、預測匯率、模型假設不變
_PX5 = "Yahoo Finance 日線收盤 2026-11-05"
_set_in("px", label="股價（2026-11-05 收盤）", base=1264000, tag="Verified", src=_PX5 + "；v0.7.3 為 2026-11-04 收盤 1,396,000")
CONFIG["tp"]["peers"] = [
    ("三星電子", "005930.KS", "=218000/(215500/1.26)*201500/218000", "2026-11-05", f"現行 P/B（11-04 收盤 218,000 ÷ 2027E 每股淨值，＝10-23 收盤 215,500 ÷ 頁面 P/B 1.26；{_MSP}）× 11-05 收盤 201,500 ÷ 11-04 收盤 218,000（{_PX5}）"),
    ("美光", "MU", "=795.23/(877.55/3.35)*656.39/795.23", "2026-11-05", f"現行 P/B（11-04 收盤 795.23 ÷ FY2027E（2027-09）每股淨值，＝10-23 收盤 877.55 ÷ 頁面 P/B 3.35；{_MSP}）× 11-05 收盤 656.39 ÷ 11-04 收盤 795.23（{_PX5}）"),
    ("鎧俠（Kioxia）", "285A.T", "=10735/7808*9324/10735", "2026-11-05", f"現行 P/B（11-04 收盤 ¥10,735 ÷ FY2028E（2028-03）每股淨值 ¥7,808；{_MSP}；sndk/、kioxia/ 模型同值）× 11-05 收盤 ¥9,324 ÷ 11-04 收盤 ¥10,735（{_PX5}）"),
    ("南亞科", "2408.TW", "=399.5/248.9*360/399.5", "2026-11-05", f"現行 P/B（11-04 收盤 399.5 ÷ 2027E 每股淨值 248.9，MarketScreener 2026-11-06 讀取）× 11-05 收盤 360 ÷ 11-04 收盤 399.5（{_PX5}）")]
CONFIG["tp"]["peer_note"] = CONFIG["tp"]["peer_note"] + "；v0.7.4 起報價改 2026-11-05 收盤（Yahoo Finance；現行倍數 × 11-05 ÷ 11-04 收盤），共識每股淨值不變"
CONFIG["obs_tpl"][1] = CONFIG["obs_tpl"][1].replace("相對現價 {px|P}（2026-11-04）", "相對現價 {px|P}（2026-11-05）")
assert "（2026-11-05）" in CONFIG["obs_tpl"][1]
CONFIG["version"] = "v0.7.4"
CONFIG["date"] = "2026-11-06"
CONFIG["a1_note"] = CONFIG["a1_note"] + "；v0.7.4（2026-11-06）：現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守"
CONFIG["readme_tail"] = CONFIG["readme_tail"] + ["v0.7.4：現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守"]

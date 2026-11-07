# 三星電子設定檔（記憶體樣板 v0.3；集團公司，只建模記憶體事業，其餘事業為輸入）
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))), "siliconomics"))
import reader as _SIL   # HBM 營收份額 sh25–sh29 讀 Siliconomics M24（v0.8 起；數值權威在本庫 In 頁）
S = "sec"
Z = lambda k, nm: (k, nm, "兆韓元", None, 0, None, "Verified", "")
CONFIG = {
 "name": "三星電子", "version": "v0.3", "date": "2026-10-25",
 "unit": "兆韓元", "ccy_per_usdB": "fx/1000", "share_scale": 1000000,
 "base_fy": 2026, "fy_offset": 0, "fy_label": "{}", "fy_desc": "日曆年（K-IFRS 合併）；記憶體事業以外為輸入",
 "thesis": "三星電子 2026 年的獲利幾乎全部來自記憶體，盈餘由一般 DRAM 價格循環與 HBM 份額追趕兩條線決定；股價取決於循環見頂的年度、見頂時累積的淨值，以及非記憶體事業（手機、晶圓代工、面板、家電）在谷底時能撐住多少。",
 "q_source": "三星 IR 季度簡報（2025Q1–2026Q2）、2026Q3 初步財報（2026-10-08）",
 "q_note": "三星只揭露記憶體營收，不揭露記憶體營業利益與 DRAM／NAND 拆分。記憶體營業利益＝DS 營業利益＋晶圓代工與 LSI 虧損估計（2025 每季約 2 兆、2026 約 2 兆）[Derived／Assumed]；DRAM 68%／NAND 32% 依 TrendForce 2Q26 三星 DRAM 營收約 82 兆推得 [Derived]；營業費用約 7% 營收、折舊為 Assumed。淨利為合併歸屬母公司。",
 "hbm_share": {2025: "sh25", 2026: "sh26", 2027: "sh27", 2028: "sh28", 2029: "sh29"},
 "pb_names": ["P/B 1.0x（谷底）", "P/B 歷史中位 1.4x", "P/B 前高 2.0x"], "pe_vals": [4, 8, 12],
 "quarters": [
  {"label": "2025Q1", "type": "A", "rev": 19.1, "op": 3.1, "ni": 8.0, "da": 5.0},
  {"label": "2025Q2", "type": "A", "rev": 21.2, "op": 2.4, "ni": 4.9, "da": 5.0},
  {"label": "2025Q3", "type": "A", "rev": 26.7, "op": 9.0, "ni": 12.0, "da": 5.0},
  {"label": "2025Q4", "type": "A", "rev": 37.1, "op": 18.4, "ni": 19.3, "da": 5.0},
  {"label": "2026Q1", "type": "A", "rev": 74.8, "dram": "=F4*0.68", "nand": "=F4*0.32", "cogs": "=F4*(1-0.745-0.07)", "op": 55.7, "ni": 47.1, "da": 6.0},
  {"label": "2026Q2", "type": "A", "rev": 120.8, "dram": "=G4*0.68", "nand": "=G4*0.32", "cogs": "=G4*(1-0.755-0.07)", "op": 91.2, "ni": 71.3, "da": 6.3},
  {"label": "2026Q3E", "type": "F", "keys": {"db": "q3_db", "dp": "q3_dp", "nb": "q3_nb", "np": "q3_np", "da": "da_q3"}},
  {"label": "2026Q4E", "type": "F", "keys": {"db": "q4_db", "dp": "q4_dp", "nb": "q4_nb", "np": "q4_np", "da": "da_q4"}},
 ],
 "inputs": [
  (S, "A. 匯率、股價"),
  ("fx", "美元兌韓元", "KRW/USD", *_SIL.vals("fx_krw"), "Derived", "同 SK 海力士"),
  ("px", "普通股股價（2026-10-23 收盤）", "KRW", None, 215500, None, "Verified", "markets/invest_map/data/prices.csv（tradersunion，單一來源）；MarketScreener 2026-10-23 收盤 215,500（第二來源，本次核對）；v0.2.2 為 10-08 的 262,000"),
  (S, "B. 基準年剩餘季度（公司指引＋TrendForce）"),
  ("q3_db", "Q3 DRAM 位元季增", "%", 0.03, 0.05, 0.07, "Verified", "公司 Q2 法說：中個位數"),
  ("q3_dp", "Q3 DRAM ASP 季增", "%", 0.08, 0.10, 0.15, "Derived", "以 2026Q3 初步財報校準：合併營收 195 兆（+13.7%）− 非記憶體約 55 兆 → 記憶體約 140 兆（+16%）；TrendForce 產業 +13–18%，三星年約價格帶使實現價較低"),
  ("q3_nb", "Q3 NAND 位元季增", "%", 0.06, 0.08, 0.09, "Verified", "公司：高個位數"),
  ("q3_np", "Q3 NAND ASP 季增", "%", 0.08, 0.10, 0.15, "Derived", "同上校準；TrendForce +10–15%"),
  ("q4_db", "Q4 DRAM 位元季增", "%", 0.0, 0.05, 0.08, "Assumed", "P4 全面量產"),
  ("q4_dp", "Q4 DRAM ASP 季增", "%", 0.05, 0.10, 0.15, "Interested-party", "TrendForce +10–15%"),
  ("q4_nb", "Q4 NAND 位元季增", "%", 0.0, 0.05, 0.08, "Assumed", ""),
  ("q4_np", "Q4 NAND ASP 季增", "%", 0.10, 0.175, 0.20, "Interested-party", "TrendForce +15–20%"),
  ("da_q3", "Q3 記憶體折舊", "兆韓元", None, 6.6, None, "Assumed", "未揭露"),
  ("da_q4", "Q4 記憶體折舊", "兆韓元", None, 7.0, None, "Assumed", ""),
  ("costdn_q", "單位成本季降", "%", 0.01, 0.03, 0.05, "Assumed", ""),
  ("opex_fix_q", "記憶體營業費用固定部分（每季）", "兆韓元", None, 1.0, None, "Assumed", "約 7% 營收"),
  ("nonop_q", "每季業外淨收益", "兆韓元", 0.5, 1.5, 2.5, "Assumed", "淨現金約 168 兆 × 約 3%"),
  ("nm_op_q", "非記憶體事業每季營業利益（Q3、Q4）", "兆韓元", -1.0, 2.0, 4.0, "Assumed", "MX 新機季節性、SDC、晶圓代工虧損；Q1 1.5、Q2 −1.7 [Derived]"),
  (S, "C. 年度驅動"),
  ("cb_1", "一般 DRAM 位元成長 2027", "%", 0.10, 0.15, 0.20, "Derived", "P4 提前全面量產、P5 提前 6 個月（TrendForce）；HBM 晶圓占比上升"),
  ("cb_2", "一般 DRAM 位元成長 2028", "%", 0.12, 0.18, 0.22, "Derived", "P5 投產"),
  ("cb_3", "一般 DRAM 位元成長 2029", "%", 0.10, 0.15, 0.20, "Assumed", ""),
  ("nb_1", "NAND 位元成長 2027", "%", 0.10, 0.15, 0.20, "Assumed", ""),
  ("nb_2", "NAND 位元成長 2028", "%", 0.12, 0.18, 0.22, "Assumed", ""),
  ("nb_3", "NAND 位元成長 2029", "%", 0.12, 0.18, 0.22, "Assumed", ""),
  ("sh25", "HBM 營收市占 2025", "%", *_SIL.vals("hbmsh_ss_25"), "Interested-party", "Counterpoint 2Q25 15%、3Q25 22%"),
  ("sh26", "HBM 營收市占 2026", "%", *_SIL.vals("hbmsh_ss_26"), "Interested-party", "Counterpoint 1Q26 21%、2Q26 33%；下半年 HBM4 放量"),
  ("sh27", "HBM 營收市占 2027", "%", *_SIL.vals("hbmsh_ss_27"), "Interested-party", "UBS 2027 位元 41%（SK 39%、美光 20%）；與 SK 42%、美光 20% 加總約 100%"),
  ("sh28", "HBM 營收市占 2028", "%", *_SIL.vals("hbmsh_ss_28"), "Assumed", "公司目標：HBM 份額向 DRAM 份額（約 39%）靠攏"),
  ("sh29", "HBM 營收市占 2029", "%", *_SIL.vals("hbmsh_ss_29"), "Assumed", ""),
  ("other_y", "其他記憶體營收", "兆韓元", None, 0, None, "Derived", "DRAM＋NAND 已涵蓋"),
  ("w_c", "成本權重：一般 DRAM", "%", None, 0.55, None, "Assumed", ""),
  ("w_h", "成本權重：HBM", "%", None, 0.15, None, "Assumed", ""),
  ("w_n", "成本權重：NAND", "%", None, 0.30, None, "Assumed", ""),
  ("costdn_y", "單位成本年降", "%", 0.05, 0.10, 0.15, "Assumed", ""),
  ("da_life", "新增資本支出轉為折舊的年數", "年", 4, 5, 6, "Assumed", ""),
  ("da_cogs", "折舊計入營業成本比例", "%", None, 0.9, None, "Assumed", ""),
  ("opex_fix_y", "記憶體營業費用固定部分（年）", "兆韓元", None, 4.0, None, "Assumed", ""),
  ("opex_var", "記憶體營業費用變動率", "% 營收", 0.04, 0.06, 0.08, "Assumed", "未揭露"),
  ("opex_g", "營業費用固定部分年增", "%", None, 0.10, None, "Assumed", ""),
  ("tax", "有效稅率", "%", 0.15, 0.20, 0.24, "Derived", "2Q26 約 20%（淨利 71.6 ÷ 稅前推估）；2025 為 8.7%"),
  ("rf", "淨現金收益率", "%", None, 0.03, None, "Assumed", ""),
  ("cap_0", "記憶體資本支出 2026", "兆韓元", 45, 52.5, 60, "Derived／Assumed", "全公司約 70（首爾經濟日報）× 記憶體約 75%（Assumed）"),
  ("cap_1", "記憶體資本支出 2027", "兆韓元", 50, 60, 75, "Assumed", "P5 主體工程（>80 兆，多年）"),
  ("cap_2", "記憶體資本支出 2028", "兆韓元", 55, 66, 80, "Assumed", ""),
  ("cap_3", "記憶體資本支出 2029", "兆韓元", 55, 66, 80, "Assumed", ""),
  ("nm_rev_0", "非記憶體事業營收 2026", "兆韓元", None, 224.8, None, "Derived／Assumed", "合併 − 記憶體：Q1 59.1、Q2 50.7（含內部抵銷）；Q3 55、Q4 60 Assumed"),
  ("nm_op_0", "非記憶體事業營業利益 2026", "兆韓元", None, 3.8, None, "Derived／Assumed", "Q1 1.5、Q2 −1.7（Derived）；Q3、Q4 各 2.0（Assumed）"),
  ("nm_rev_1", "非記憶體事業營收 2027", "兆韓元", None, 230, None, "Assumed", ""),
  ("nm_op_1", "非記憶體事業營業利益 2027", "兆韓元", 0, 6, 12, "Assumed", "晶圓代工虧損縮小（HBM4 基底晶粒）；MX 受記憶體成本擠壓"),
  ("nm_rev_2", "非記憶體事業營收 2028", "兆韓元", None, 240, None, "Assumed", ""),
  ("nm_op_2", "非記憶體事業營業利益 2028", "兆韓元", 0, 8, 15, "Assumed", "記憶體價格回落時 MX 成本減輕"),
  ("nm_rev_3", "非記憶體事業營收 2029", "兆韓元", None, 250, None, "Assumed", ""),
  ("nm_op_3", "非記憶體事業營業利益 2029", "兆韓元", 0, 8, 15, "Assumed", ""),
  (S, "D. 資本與股東回饋"),
  ("eq_last", "權益（歸屬母公司）2026-06-30", "兆韓元", None, 565.06, None, "Verified", "2Q26 合併資產負債表"),
  ("nc_last", "淨現金 2026-06-30", "兆韓元", None, 167.59, None, "Derived", "現金 190.00 − 借款 22.41"),
  ("sh_last", "流通股數（普通股扣庫藏＋特別股）", "百萬股", None, 6566.56, None, "Verified", "普通 5,846.28 − 庫藏 82.09 ＋ 特別 802.37"),
  ("extra_eq", "基準年其餘增資", "兆韓元", None, 0, None, "Verified", ""),
  ("extra_sh", "基準年其餘新股", "百萬股", None, 0, None, "Verified", ""),
  ("bb_0", "2026 下半年買回", "兆韓元", None, 15, None, "Verified", "2026-08-21 董事會：約 15 兆（員工酬勞用）"),
  ("bb_sh_0", "2026 下半年買回股數", "百萬股", None, 57.3, None, "Derived", "15 兆 ÷ 262,000"),
  ("div_0", "2026 下半年股利", "兆韓元", None, 32.5, None, "Verified／Derived", "3Q26 約 30 兆現金股利（含定期）＋4Q 定期 2.45"),
  ("capex_rem", "2026 下半年記憶體資本支出", "兆韓元", None, 31.4, None, "Derived", "全年約 70 − 上半年 28.1 ＝ 41.9 × 75%"),
  ("dps_1", "每股股利 2027", "KRW", None, 1500, None, "Verified", "定期股利 9.8 兆／年 ÷ 6,567M 股"),
  ("dps_2", "每股股利 2028", "KRW", None, 1500, None, "Verified", "同上"),
  ("dps_3", "每股股利 2029", "KRW", None, 1500, None, "Assumed", ""),
  ("payout", "股東回饋 ÷ FCF", "%", 0.3, 0.5, 0.7, "Derived", "2024–26 合計 120–140 兆；2026 年 90–110 兆；剩餘部分 2027-01 決定"),
  ("bb_px", "未來買回均價", "KRW", None, 262000, None, "Assumed", "＝目前股價"),
  (S, "E. 估值倍數"),
  ("pb_1", "P/B 1.0x（谷底）", "x", None, 1.0, None, "Interested-party", "2015、2018、2022 年約 1.0；2024 年 0.89"),
  ("pb_2", "P/B 歷史中位", "x", None, 1.4, None, "Interested-party", "2010–2025 年底中位約 1.4"),
  ("pb_3", "P/B 前高", "x", None, 2.0, None, "Interested-party", "2020 年 1.97、2025 年 1.96"),
  ("pe_1", "P/E 4x", "x", None, 4, None, "Assumed", "共識 2027E P/E 約 3.7x"),
  ("pe_2", "P/E 8x", "x", None, 8, None, "Assumed", ""),
  ("pe_3", "P/E 12x", "x", None, 12, None, "Assumed", ""),
 ],
 "recon": [
  ("2026Q3E 合併營業利益 vs 公司初步財報", "=Q!H10+nm_op_q", 107.4, "三星 2026-10-08 初步財報（約 107.4 兆）[Verified]"),
  ("2026E 合併營收 vs 共識", "=Y!C{crev}", 711.7, "MarketScreener 2026-10-06"),
  ("2026E 合併營業利益 vs 共識", "=Y!C{cop}", 373.7, "MarketScreener；KB 381"),
  ("2026E 淨利 vs 共識", "=Y!C{ni}", 303.3, "MarketScreener"),
  ("2027E 合併營業利益（A）vs 共識", "=Y!D{cop}", 568.7, "MarketScreener；FnGuide 544.7；KB 574"),
  ("2027E 淨利（A）vs 共識", "=Y!D{ni}", 454.7, "MarketScreener"),
  ("2028E 合併營業利益（A）vs 共識", "=Y!E{cop}", 578.4, "MarketScreener；FnGuide 543.9"),
 ],
 "sources": [
  ("三星 2026Q3 初步財報（2026-10-08）", "https://news.samsung.com/global/samsung-electronics-announces-earnings-guidance-for-third-quarter-2026"),
  ("三星 IR 2Q26 簡報", "https://images.samsung.com/is/content/samsung/assets/global/ir/docs/2026_2Q_conference_eng.pdf"),
  ("三星 IR 1Q26 簡報", "https://images.samsung.com/is/content/samsung/assets/global/ir/docs/2026_1Q_conference_eng.pdf"),
  ("三星 IR 4Q25 簡報", "https://images.samsung.com/is/content/samsung/assets/global/ir/docs/2025_4Q_conference_eng.pdf"),
  ("三星 2Q26 資產負債表", "https://images.samsung.com/kdp/ir/financial-info/2026/2026_con_quarter02_bs.pdf"),
  ("三星 2026 股東回饋（2026-08-21）", "https://news.samsung.com/global/samsung-electronics-to-implement-largest-ever-shareholder-return-in-2026-estimated-at-krw-90-to-110-trillion"),
  ("TrendForce 2Q26 法說摘要", "https://www.trendforce.com/news/?p=61735"),
  ("FN News（TrendForce 2Q26 DRAM 營收）", "https://www.fnnews.com/news/202609071940555902"),
  ("MarketScreener 共識", "https://www.marketscreener.com/quote/stock/SAMSUNG-ELECTRONICS-CO-6494906/finances/"),
  ("companiesmarketcap P/B", "https://companiesmarketcap.com/samsung/pb-ratio/"),
  ("Seoul Economic Daily（P4、資本支出）", "https://en.sedaily.com/finance/2026/04/22/samsung-accelerates-p4-fab-advances-full-production-by-six"),
 ],
 "readme_extra": ["vv0.2.1：Val 頁新增「十年倍數與規則值」區塊（倍數歷史、分位、規則值與隱含股價；原放在 markets/overview/extract.py），總覽與 HTML 展示層改讀此處；損益、EPS 與原有儲存格不變。", "v0.2：HBM 出貨位元隨情境變動（markets/memory）：2028–29 需求位元 × 情境顆數係數（B 0.8／0.85、C 0.6／0.6；與台積電、Advantest 同一套），供給與 2028 前的單價假設不變。2026–27 與 v0.1.1 相同；情境 A 全部相同；B、C 的 2028–29 HBM 營收下降，營業成本量指數隨 HBM 位元同步下降。", "v0.1.1：Tokenomics v19.0 同步。只有對帳頁的「每顆 GPU 的 NVIDIA 內容」由 51.6 改 42.8 千美元（橋接表 v0.2：GB300 機架基準價 $4.3M）；損益、淨值、估值與 v0.1 相同。", "集團公司處理：只建模記憶體事業（DRAM、HBM、NAND），手機、晶圓代工、面板、家電、Harman 以「非記憶體事業營收、營業利益」輸入；權益、淨現金、股數、估值為全公司。",
                  "最大限制：三星不揭露記憶體營業利益；基準年記憶體利潤率由 DS 營業利益加回晶圓代工虧損估計推得，並以 2026Q3 初步財報（合併營業利益 107.4 兆）核對。"],
}

# 估值規則的十年倍數歷史（2016–2025 年底；規則見 tools/valuation_rule.py；值原放在 markets/overview/extract.py，2026-10-23 移入）
CONFIG["rule"] = {
 "pb": {"sources": [
   {"label": "companiesmarketcap", "vals": [1.27, 1.59, 1.04, 1.41, 1.97, 1.73, 1.04, 1.43, .8865, 1.96], "tag": "Verified", "note": "核對：GuruFocus 十年中位 1.44"}]},
 "pe": {"sources": [
   {"label": "companiesmarketcap（年底 P/E）", "vals": [9.15, 6.95, 5.06, 15.3, 19.1, 12.5, 6.50, 36.1, 10.6, 18.2], "tag": "Verified"},
   {"label": "stockanalysis（核對，不採用）", "vals": [None, None, None, None, None, 13.40, 6.79, 35.92, 10.53, 17.39], "tag": "Verified", "use": False, "note": "2021–2025 與 companiesmarketcap 差距 ≤ 8%"}]},
}

# v0.2.2（2026-10-25）：目標價與上檔空間（Andy 2026-10-25 要求；Val 頁最後一節、In 頁最後三列；既有儲存格不變）
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.2.2（2026-10-25）：Val 頁新增「0. 目標價與上檔空間」（放在 Val 既有內容之後，不位移既有列）：目標價＝In 頁權重 tp_wA／tp_wB／tp_wC（預設 A 0.3／B 0.5／C 0.2，Assumed，Andy 可改）× 三個 P/B 規則值（A＝中段、B＝見頂、C＝谷底）；既有儲存格的值不變。"]
CONFIG["tp_bias"] = "十年 P/B 含 AI 前的低倍數期，且為全公司口徑（含手機、晶圓代工等非記憶體事業的集團折價），偏誤方向為低估。"

# v0.3（2026-10-25）：目標價改為 DCF 與 Comparable 加權（Andy 11:28、11:53）；三情境延伸到 2032（收斂法）；現價更新到 10-23
MS_EST = "MarketScreener 預估頁（2026-10-25 擷取；頁面為 10-23 收盤）"
CONFIG["tp"] = {
 "roe_label": "ROE（近三年平均，2023–2025）", "roe": "=AVERAGE(0.0431,0.0900,0.1078)",
 "roe_src": "stockanalysis 比率頁：2023 4.31%、2024 9.00%、2025 10.78%（單一來源）",
 "nim10": "=AVERAGE(22.726/201.867,42.187/239.575,44.345/243.771,21.739/230.401,26.408/236.807,0.1404,0.1811,0.0559,0.1118,0.1327)",
 "nim10_src": "2021–2025：stockanalysis 損益表淨利率（14.04%、18.11%、5.59%、11.18%、13.27%）；2016–2020：公司年報合併營收與淨利（兆韓元，本次未重新開啟原文，單一來源）",
 "cons_tp": (478906, 290000, 725000, "MarketScreener 共識頁：36 位分析師平均 478,906（最低 290,000、最高 725,000），2026-10-25 擷取"),
 "cons": [("cons_eps27", "共識 EPS 2027", 72279, MS_EST), ("cons_eps28", "共識 EPS 2028", 75531, MS_EST), ("cons_bvps28", "共識每股淨值 2028 年底", 232055, MS_EST)],
 "cons_src": MS_EST,
 "peers": [("SK 海力士", "000660.KS", 1.35, "2026-10-23", "MarketScreener：現價 ÷ 2027E 每股淨值"),
           ("美光", "MU", 3.35, "2026-10-23", "MarketScreener：現價 ÷ FY2027E（2027-09）每股淨值"),
           ("鎧俠（Kioxia）", "285A.T", 2.22, "2026-10-09", "MarketScreener：現價 ÷ FY2028E（2028-03）每股淨值；頁面價格為 10-09"),
           ("南亞科", "2408.TW", 1.66, "2026-10-22", "MarketScreener：現價 ÷ 2027E 每股淨值")],
 "peer_note": "記憶體同業（SK 海力士、美光、鎧俠、南亞科；無 A 股）。前瞻 P/B＝現價 ÷ 下一個會計年度末（約 2027 年底）共識每股淨值；各家會計年度不同（美光 9 月、鎧俠 3 月）。三星為集團公司，純記憶體同業倍數未含集團折價",
 "fwd_w": ("2/12", "10/12"), "fwd_note": "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × 2027＋10／12 × 2028（會計年度＝日曆年）",
 "method": ("DCF：股東現金流 2027–2029 用每股（淨利＋折舊－資本支出；折舊與資本支出為記憶體事業口徑，淨利為合併），2030–2032 收斂期用 EPS ×（1 − g／ROE），r 10%、g 4%、ROE 為 2023–2025 平均約 8%（再投資率約五成）；明確期＋永續，基準日 2026 年底。"
            "Comparable：SK 海力士、美光、鎧俠、南亞科前瞻 P/B 中位數 × 2028 年底每股淨值（情境加權）。"
            "收斂期（2030–2032）淨利率收斂到循環中段：A 十年平均 × 1.2、B 十年平均、C 十年平均 × 0.8（In 頁 cyc_mA／cyc_mC），永續期以 2032 中段水準為基礎。偏誤：2027–2029 仍為高峰期盈餘且權重大，DCF 仍可能偏高估；純記憶體同業倍數未含集團折價，Comparable 偏高估。"
            "情境對應：A＝樂觀、B＝基準、C＝保守（原為 TrendForce／BofA 型價格情境）。"),
}
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.3（2026-10-25）：Val 頁新增「0. 目標價」（DCF 與 Comparable 加權，tp_w_dcf 0.5）；原「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」，情境對應改為 A 75 分位、B 中位數、C 25 分位 × 各情境 2027 年底每股淨值。Y 頁三情境延伸到 2032（2030 起收斂法）。現價更新為 2026-10-23 收盤 215,500。2026–2028 EPS 與 v0.2.2 相同；2029 因修正共用工具「前一年」基期（原誤用 2027）而改變。收斂期淨利率收斂到循環中段（十年平均 × 1.2／1.0／0.8）。"]

# v0.3.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00、14:17；樣板 v0.4.1，數值不變）
CONFIG["version"] = "v0.3.1"
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.3.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁最後新增「上檔空間級距」「投資建議」兩列（依「上檔空間：加權目標價」固定規則）；README 頁最後新增投資建議與重點觀察三點。"]
CONFIG["obs"] = [
 "目標價 ₩477,143（DCF 486,542、Comparable 467,744 各半），比現價 215,500（2026-10-23）高 121.4%，級距 大於 50%；Comparable 單獨為 +117.1%。主因是倍數不是 EPS：本模型 2028 年底每股淨值 241,105 只比共識 232,055 高 3.9%（前瞻 EPS 78,917 比共識高 5.2%），但現價只隱含 0.89 倍前瞻 P/B，記憶體同業中位數 1.94 倍、共識目標價 478,906 隱含 2.06 倍；同業倍數未扣三星的集團折價，Comparable 偏高估。",
 "DCF 情境價值 A 818,535、B 394,410、C 218,882（現價的 3.80／1.83／1.02 倍），對應 2028 EPS A 129,982、B 59,278、C 36,498；收斂期用的 ROE 為 2023–2025 平均 8.0%，低於折現率 10%，再投資率約五成，壓低 2030 年後的股東現金流。",
 "2028 B 營業利益對一般 DRAM 收入最敏感（±20% → ±15.2%），NAND ±30% → ±9.2%、HBM ±30% → ±7.3%；折現率 ±1pt 使加權目標價 −4.1%／+5.6%。下一個檢驗事件是 3Q26 確定財報與法說（約 2026-10-30，日期未確認；2026-10-08 初步財報已納入）。"]   # 重點觀察（與 README.md 相同文字）

# v0.3.2（2026-10-28）：分部原則重建（Q 頁新增官方報告分部與核對列；記憶體樣板選用區塊，數值不變）
CONFIG["version"] = "v0.3.2"
CONFIG["segments"] = {
 "title": "官方報告分部（兆韓元；分部營收含內部銷售；三星 IR 季度簡報）",
 "rows": [
  ("DX（裝置體驗：MX／網路、VD／DA）", [51.7, 43.6, 48.4, 44.3, 52.7, 48.0, None, None]),
  ("DS（裝置解決方案）", [25.1, 27.9, 33.1, 44.0, 81.7, 127.5, None, None]),
  ("SDC（三星顯示器）", [5.9, 6.4, 8.1, 9.5, 6.7, 7.5, None, None]),
  ("Harman", [3.4, 3.8, None, 4.6, 3.8, 4.6, None, None]),
  ("內部抵銷 [Derived＝合併營收 − 四分部合計]", [-7.0, -7.1, None, -8.6, -11.0, -16.1, None, None]),
 ],
 "total": ("合併營收", [79.1, 74.6, 86.1, 93.8, 133.9, 171.5, None, None]),
 "check_label": "檢查：分部合計 − 合併營收",
 "notes": ["來源：2026Q1 簡報（韓文版，含 2025Q1、2025Q4）、2026Q2 簡報（含 2025Q2；第三方轉存 gelonghui，三星官網原檔未取得）、2025Q3 財報新聞稿（news.samsungsemiconductor.com）。",
           "2025Q3：DX 未揭露 [Derived]＝2025Q2 DX × 1.11（新聞稿：季增 11%）；Harman 未揭露，該季不計核對。三星不揭露內部抵銷，故抵銷列為差額，核對列只驗證分部表本身。",
           "記憶體（第 4 列「記憶體營收」）屬 DS；DS 其餘為晶圓代工與 System LSI。推估季不拆分部：非記憶體事業以 In 頁 nm_rev／nm_op 整體輸入。"],
}
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.3.2（2026-10-28）：分部原則重建。Q 頁第 20 列起新增官方報告分部 DX／DS／SDC／Harman、內部抵銷（差額）與合併營收（2025Q1～2026Q2 實績），核對列「分部合計 − 合併營收」；記憶體營收屬 DS。損益、EPS、目標價與 v0.3.1 相同。"]

# v0.3.3（2026-10-30）：In 頁版面（個股模型原則第 8 節；數值不變）
CONFIG["version"] = "v0.3.3"

# v0.4（2026-10-30）：滾入 2026Q3 確定財報（2026-10-29）與 3Q26 法說的 4Q 指引；資產負債錨點移到 2026-09-30（推估，9-30 資產負債表待季報）
CONFIG["version"] = "v0.4"
CONFIG["date"] = "2026-10-30"
CONFIG["q_source"] = "三星 IR 季度簡報（2025Q1–2026Q2）、2026Q3 確定財報與法說（2026-10-29）"
CONFIG["q_note"] = CONFIG["q_note"] + "　2026Q3 實績：記憶體營收 140.7 兆（公司）；記憶體營業利益 108.0＝DS 營業利益 106.5＋晶圓代工與 LSI 虧損估計 1.5 [Derived／Assumed]；DRAM 68%／NAND 32% 沿用 2Q 結構 [Assumed]（法說 DRAM ASP 約 +20%、位元低至中個位數，NAND ASP 十幾 %、位元中個位數，兩者營收增幅相近；但價量相乘約 +20%，高於記憶體營收實際季增 16%，差距可能來自韓元升值的匯率影響〔法說：季減約 10 兆，以零組件為主〕與 ASP 口徑）；淨利 74.8＝合併淨利 75.1361 × 2Q 歸屬母公司比例 0.9955 [Derived]；折舊 Assumed。"
CONFIG["quarters"][6] = {"label": "2026Q3", "type": "A", "rev": 140.7, "dram": "=H4*0.68", "nand": "=H4*0.32", "cogs": "=H4*(1-0.768-0.07)", "op": 108.0, "ni": 74.8, "da": 6.6}
_UPD = {
 "q3_db": ("Q3 DRAM 位元季增（v0.4 起 Q3 為實績，本列不再使用）", "%", 0.03, 0.05, 0.07, "Interested-party", "公司 Q2 法說：中個位數；3Q26 法說實際為低至中個位數"),
 "q3_dp": ("Q3 DRAM ASP 季增（v0.4 起不再使用）", "%", 0.08, 0.10, 0.15, "Derived", "3Q26 法說實際約 +20%"),
 "q3_nb": ("Q3 NAND 位元季增（v0.4 起不再使用）", "%", 0.06, 0.08, 0.09, "Interested-party", "3Q26 法說實際為中個位數"),
 "q3_np": ("Q3 NAND ASP 季增（v0.4 起不再使用）", "%", 0.08, 0.10, 0.15, "Derived", "3Q26 法說實際為十幾 %"),
 "q4_db": ("Q4 DRAM 位元季增", "%", 0.12, 0.16, 0.19, "Interested-party", "公司 3Q26 法說（2026-10-29）：4Q DRAM 位元季增中至高十幾 %（v0.3.3 為 5%，Assumed）"),
 "q4_nb": ("Q4 NAND 位元季增", "%", 0.12, 0.15, 0.18, "Interested-party", "公司 3Q26 法說：4Q NAND 位元季增中十幾 %（v0.3.3 為 5%，Assumed）"),
 "da_q3": ("Q3 記憶體折舊（v0.4 起為實績列折舊）", "兆韓元", None, 6.6, None, "Assumed", "未揭露"),
 "nm_op_q": ("非記憶體事業每季營業利益（Q4；Q3 起為實績）", "兆韓元", -1.0, 2.0, 4.0, "Assumed", "Q1 1.5、Q2 −1.7、Q3 −0.56（合併 107.4405 − 記憶體 108.0）[Derived]；Q4 MX 新機旺季、晶圓代工營收雙位數成長（公司）"),
 "nm_rev_0": ("非記憶體事業營收 2026", "兆韓元", None, 224.4, None, "Derived／Assumed", "合併 − 記憶體：Q1 59.1、Q2 50.7、Q3 54.55（195.2532 − 140.7）（含內部抵銷）；Q4 60 Assumed"),
 "nm_op_0": ("非記憶體事業營業利益 2026", "兆韓元", None, 1.2, None, "Derived／Assumed", "Q1 1.5、Q2 −1.7、Q3 −0.56（Derived）；Q4 2.0（Assumed）"),
 "cap_0": ("記憶體資本支出 2026", "兆韓元", 55, 63.3, 70, "Derived／Assumed", "公司 3Q26：全年約 84.4 兆（DS 76.8、SDC 3.9）× 記憶體約 75%（Assumed；v0.3.3 用約 70 兆 × 75%）"),
 "eq_last": ("權益（歸屬母公司）2026-09-30（推估）", "兆韓元", None, "=565.06+74.8-2.45", None, "Derived／Assumed", "6-30 權益 565.06（2Q26 資產負債表）＋ 3Q 淨利 74.8 − 2Q 股利 2.45（8 月支付）；15 兆買回假設在 4Q 執行（bb_0）；9-30 資產負債表待季報（約 11 月中）校正"),
 "nc_last": ("淨現金 2026-09-30（推估）", "兆韓元", None, "=167.59+74.8+6.6-29.0*0.75-2.45", None, "Derived／Assumed", "6-30 淨現金 167.59 ＋ 3Q 淨利 74.8 ＋ 記憶體折舊 6.6 − 3Q 記憶體資本支出（3Q 全公司 29.0 × 75%）− 2Q 股利 2.45；與樣板現金橋同口徑（只計記憶體折舊與資本支出），待季報校正"),
 "capex_rem": ("2026 第四季記憶體資本支出", "兆韓元", None, "=(84.4-56.9)*0.75", None, "Derived", "全年 84.4 − 9M 56.9（公司 3Q26）＝ 27.5 × 75%"),
 "div_0": ("2026 第四季起支付的 2026 股利（3Q 特別＋4Q 定期）", "兆韓元", None, 32.5, None, "Verified／Derived", "3Q26 現金股利每股 4,601 韓元、合計約 30 兆（公司 10-29，11 月支付）＋4Q 定期 2.45；2026 股東回饋總額 90–110 兆、剩餘 55–75 兆（含大規模買回註銷）由 2027-01 董事會決定，未計入本列"),
 "bb_0": ("2026 第四季買回", "兆韓元", None, 15, None, "Verified／Assumed", "2026-08-21 董事會：約 15 兆（員工酬勞用）；執行期間未載明，假設 4Q（9-30 錨點未扣）"),
}
CONFIG["inputs"] = [((k,) + _UPD[k] if isinstance(r, tuple) and r[0] in _UPD and (k := r[0]) else r) for r in CONFIG["inputs"]]
CONFIG["recon"] = [r for r in CONFIG["recon"] if not r[0].startswith("2026Q3E")] + [
 ("2026Q3 記憶體營業利益（推估）vs Meritz 估計", "=Q!H10", 111.9, "Meritz Securities 2026-10-08（Herald Business 轉述）[Interested-party]；本模型＝DS 106.5＋晶圓代工與 LSI 虧損估計 1.5")]
CONFIG["sources"] = [("三星 2026Q3 確定財報（2026-10-29）", "https://news.samsung.com/global/samsung-electronics-announces-third-quarter-2026-results"),
                     ("3Q26 法說摘要（VLSI Korea，位元與 ASP、4Q 指引、淨利）", "https://www.vlsi.kr/en/samsung-q3-2026-earnings-11-facts-en/")] + CONFIG["sources"]
_seg = CONFIG["segments"]
for (lab, vals), v in zip(_seg["rows"], [52.1, 147.1, 8.0, 4.3, None]):
    vals[6] = v
_seg["rows"][4][1][6] = round(195.2532 - (52.1 + 147.1 + 8.0 + 4.3), 4)
_seg["total"][1][6] = 195.2532
_seg["notes"] = _seg["notes"] + ["2026Q3：確定財報（2026-10-29）DX 52.1、DS 147.1、SDC 8.0、Harman 4.3（VLSI Korea 轉述法說）、合併 195.2532。"]
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.4（2026-10-30）：滾入 2026Q3 確定財報（2026-10-29：合併營收 195.2532、營業利益 107.4405、淨利 75.1361 兆；DS 147.1／106.5、記憶體營收 140.7）。記憶體營業利益 108.0 為 DS 加回晶圓代工與 LSI 虧損估計 1.5 [Derived／Assumed]；4Q DRAM 位元季增改為公司指引中至高十幾 %（16%，原 5%）、NAND 中十幾 %（15%，原 5%）；2026 記憶體資本支出依公司全年 84.4 兆改為 63.3 兆（原 52.5）；資產負債錨點移到 9-30（權益、淨現金為推估，待季報校正）；Q 頁分部表補 2026Q3。結果：2027 EPS（A）95,824 → 98,695（+3.0%）、2028 EPS（B）59,278 → 60,865（+2.7%）、加權目標價 485,941 → 498,385（+2.6%），強力買進不變。"]
CONFIG["obs"] = [
 "目標價 ₩498,385（DCF 506,370、Comparable 相對倍數調整後 490,401 各半），比現價 215,500（2026-10-23）高 131.3%，級距 大於 50%；Comparable（市場主流同業法）單獨為 +119.3%。主因仍是倍數不是 EPS：v0.4 滾入 3Q26 實績與 4Q 指引後，2028 年底每股淨值 243,619 只比共識 232,055 高 5.0%，現價隱含的前瞻 P/B 不到 0.9 倍，記憶體同業中位數 1.94 倍。",
 "DCF 情境價值 A 858,019、B 407,898、C 225,075（現價的 3.98／1.89／1.04 倍），對應 2028 EPS A 134,854、B 60,865、C 37,230。4Q 位元季增由 5% 上調到公司指引的 15–16%，使 2027 起營收基期上移約 3%；2026 記憶體資本支出依公司全年 84.4 兆上調 10.8 兆，部分抵銷自由現金流。",
 "2028 B 營業利益對一般 DRAM 收入最敏感（±20% → ±15.5%），NAND ±30% → ±9.2%、HBM ±30% → ±7.2%；折現率 ±1pt 使加權目標價 −4.1%／+5.6%。待校正：9-30 資產負債表（權益、淨現金為推估）於季報公布後更新；2026 股東回饋剩餘 55–75 兆（含大規模買回註銷）待 2027-01 董事會，未計入。"]

CONFIG["in_layout_version"] = "v0.3.3"   # In 頁版面說明的版本標籤（v0.3.3 套用；之後升版不改）

# v0.4.1（2026-11-05）：只改展示層（個股 HTML 版面規格 v1.0；README 新增「展示層」「命題回覆」兩段，公式帶數字；其他頁數值不變）
CONFIG["version"] = "v0.4.1"
CONFIG["date"] = "2026-11-05"
CONFIG["a1_note"] = "；v0.4.1 只改展示層（README 新增展示層與命題回覆兩段、HTML 新版面），其他頁與 v0.4 相同"
CONFIG["pres"] = {
 "lead": "命題：三星電子的獲利幾乎全部來自記憶體，盈餘由一般 DRAM 價格循環與 HBM 份額追趕兩條線決定；股價取決於循環見頂的年度、見頂時累積的淨值，以及非記憶體事業在谷底時能撐住多少。四個問題如下。",
 "props": [
  "命題 1：哪條線決定盈餘：一般 DRAM 價格循環與 HBM 份額追趕，哪一條對營業利益的影響較大。",
  "命題 2：循環在哪一年見頂：三種價格情境下，盈餘在哪一年到高點、之後回落多少。",
  "命題 3：見頂時累積的淨值撐不撐得住股價：用同業倍數與現金流折現估算，每股價值是否高於現價。",
  "命題 4：非記憶體事業（手機、晶圓代工、面板、家電）在記憶體谷底時能撐住多少盈餘。"],
 "obs": [
  "重點 1：合理價值{updesc}現價，{upsrc}。我們估每股約 {tpw|P} 韓元（DCF {dcftp|P} 占 {wdcf|0%}、Comparable〔相對倍數調整後〕{compadj|P} 占其餘），相對現價 {px|P} 為 {upw|+0%;−0%}，投資建議{rec}。把加權目標價的上檔拆成兩個來源（各自的權重 ×（目標價 − 現價））：DCF 貢獻 {dcfsh|0%}、Comparable 貢獻 {compsh|0%}；DCF 反映 {y27}–{y29} 高峰期的大量自由現金流（情境 A 的 DCF 價值為現價 {dcfpxA|0.0} 倍），對收斂期假設敏感。Comparable 比現價 {upcadj|+0%;−0%}，可拆成兩個因素相乘：每股淨值由 {y27} 年底成長到 {y28} 年底（情境加權）{bvg|+0%;−0%}，以及同業倍數（相對倍數調整後 {pmadj|0.00} 倍）比現價隱含的 {y27} 年底 P/B（{pxbv27|0.00} 倍）高 {pmg|0%;−0%}，{compdrv}。本模型的每股淨值只比市場共識{gbvdir} {gbvabs|0.0%}。",
  "重點 2：三種情境的淨利高點落在 {peakA}（情境 A）、{peakB}（情境 B）、{peakC}（情境 C）；分歧集中在 {y28} 以後：{y28} 每股盈餘情境 A 約 {epsA28|E}、B 約 {epsB28|E}、C 約 {epsC28|E} 韓元。最保守的情境 C，DCF 價值為現價的 {dcfpxC|0.00} 倍。",
  "重點 3：營業利益主要看{s_top} 價格。以 {s_year} 情境 {s_scn} 為例，{s_dram_l} 使營業利益變動 {s_dram|+0.0%;−0.0%}；{s_nand_l} 為 {s_nand|+0.0%;−0.0%}、{s_hbm_l} 為 {s_hbm|+0.0%;−0.0%}（各收入線的測試幅度不同，是敏感度測試，不是門檻）。非記憶體事業是輸入假設、不建模。"],
 "ans": [
  "一般 DRAM 價格循環與 HBM 份額追趕，哪一條對盈餘影響較大？答：{s_top}。以 {s_year} 情境 {s_scn} 為例，一般 DRAM 營收約 {convA27|#,##0.0} 兆韓元、HBM 約 {hbmA27|#,##0.0} 兆韓元（HBM 占記憶體營收 {hbmshA27|0%}）；{s_dram_l} 使營業利益變動 {s_dram|+0.0%;−0.0%}，{s_hbm_l} 只變動 {s_hbm|+0.0%;−0.0%}。HBM 份額追趕的影響小於一般 DRAM 的價格循環。",
  "循環在哪一年見頂？答：依本模型三種情境的價格路徑（情境設定，不是單一預測），淨利高點在情境 A {peakA}、情境 B {peakB}、情境 C {peakC}。情境 B 高點淨利約 {nipkB|#,##0} 兆韓元，{y28} 為 {niB28|#,##0} 兆韓元（比高點 {nichgB28|+0%;−0%}）；情境 C {y28} 為 {niC28|#,##0} 兆韓元（比高點 {nichgC28|+0%;−0%}）。",
  "見頂時累積的淨值撐不撐得住股價？答：依本模型{verdict}。現價 {px|P} 韓元只相當於 {y27} 年底每股淨值（情境加權）的 {pxbv27|0.00} 倍，記憶體同業同口徑的前瞻 P/B 中位數 {pmed|0.00} 倍；目標價把同業倍數（經相對倍數調整後 {pmadj|0.00} 倍）套在 12 個月後的前瞻每股淨值（{y28} 年底，情境加權）{fbv|P} 上，得 Comparable 目標價 {compadj|P}（比現價 {upcadj|+0.0%;−0.0%}）。加上 DCF 目標價 {dcftp|P}，加權目標價 {tpw|P}，比現價 {upw|+0.0%;−0.0%}，投資建議{rec}。{dcfall}（最低的情境 C 為現價 {dcfpxC|0.00} 倍）。",
  "非記憶體事業在谷底能撐住多少？答：本模型沒有建模非記憶體事業，營收與營業利益是輸入假設。情境 C 的 {y28}、{y29}，非記憶體事業營業利益約 {nmopC28|#,##0.0}、{nmopC29|#,##0.0} 兆韓元，占合併營業利益 {nmshC28|0%}、{nmshC29|0%}；依這組假設，{nmsupport}。"],
}

# v0.5（2026-11-05）：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；樣板 v0.5、tools/fcff_dcf.py；規格 docs/plan/個股DCF標準規格.md v1.0）。
# Comparable、情境權重、DCF 權重 50% 與投資建議規則不變；舊 DCF 參數 tp_r、tp_roe 刪除；淨現金收益率鍵 rf 改名 ncy（值不變）。
_DAM = "Damodaran Betas by Sector (Global) 2026 年 1 月（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/BetasGlobal.html；2026-11-05 讀取）"
_DAMUS = "Damodaran Betas by Sector (US) 2026 年 1 月（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/Betas.html；2026-11-05 讀取）"
_CRP = "Damodaran Country Default Spreads and Risk Premiums（2026-01-05 更新；https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html；2026-11-05 讀取）"
_KR10 = ("Trading Economics「South Korea 10-Year Yield Lowest Since June」2026-10-22：約 4.14%（https://tradingeconomics.com/south-korea/government-bond-yield/news/593996；2026-11-05 讀取）",
         "同站 2026-10-27「Lowest Since May」約 4.00%、2026-11-05 為 3.90%；取與股價（2026-10-23）最近的 10-22 值；同 SK 海力士 v0.7")
_SA = "stockanalysis.com Samsung Electronics（KRX:005930）財報頁，2026-11-05 讀取；百萬韓元"
CONFIG["dcf"] = {
 "frac0": 1.0,
 "capex_note": "（記憶體事業口徑；非記憶體事業只有營業利益輸入，其折舊與資本支出在明確期視為相抵）",
 "fd": dict(
  date="2026-12-31（基準年底；同 v0.4 DCF 基準日）",
  rf=(0.0414, _KR10[0], _KR10[1], "無風險利率（韓國 10 年期公債殖利率，2026-10-22）"),
  crp=(0.0064, "Verified", _CRP + "：South Korea（Aa2）國家風險溢酬 0.64%", "Andy 2026-11-05 20:45：非美國市場 ERP＝5%＋CRP，不設股權成本下限"),
  kd=("=rf+0.01", "Assumed", "rf ＋ 1.0pt 信用利差（借款利率未逐筆查）", "負債占資本約 2%，對 WACC 影響可忽略"),
  debt=(22.41, "Verified", "2Q26 合併資產負債表：借款 22.41 兆韓元（In 頁 nc_last 同一來源；9-30 資產負債表待季報）", "只用於 WACC 權重與 β 重新加槓桿；評價日淨現金取 Y 頁基準年末"),
  lease=(3.674, "Verified", _SA + "：FY2025 Long-Term Leases 3,673,503", "2026 年租賃負債未找到；視為不含在借款內"),
  inv=(0, "Assumed", "關聯企業與權益投資（三星 SDI、三星生物等持股）未計入（帳面值未逐項取得）", "偏保守；Sens 案例「淨現金不含權益投資」因此與基準相同"),
  mi=(12.007, "Verified", _SA + "：FY2025 Minority Interest 12,007,082", "合併營業利益含子公司少數股權份額，評價時扣除"),
  dps=("=dps_1", "Derived", "＝In 頁 dps_1（2027 每股定期股利 1,500 韓元）", "2026 股東回饋剩餘 55–75 兆（含大規模買回）待 2027-01 董事會，未計入"),
  t_tgt=(1.0, "目標價為評價日（2026-12-31）後 12 個月，與 v0.4 DCF 目標價同一時點"),
  fc_da=(None, "=(38.581541/258.935494+42.630822/300.870903+46.926592/333.605938)/3", None, "Derived",
         _SA + "：近三年（2023–2025）全公司折舊攤銷 ÷ 合併營收平均＝14.9%／14.2%／14.1%（D&A 38,581,541／42,630,822／46,926,592；營收 258,935,494／300,870,903／333,605,938）",
         "明確期 2027–2029 用 Y 頁記憶體折舊；本值為 2030 起收斂的長期穩態（全公司口徑，含非記憶體事業）"),
  fc_cx=(None, "=(57.611292/258.935494+51.406355/300.870903+47.522179/333.605938)/3", None, "Derived",
         _SA + "：近三年（2023–2025）全公司資本支出 ÷ 合併營收平均＝22.2%／17.1%／14.2%（57,611,292／51,406,355／47,522,179）",
         "明確期用 Y 頁記憶體資本支出（In 頁 cap_1–cap_3）；本值為 2030 起長期穩態（全公司口徑）；穩態資本支出 ≥ 折舊 ×（1＋g）（DCF 頁一致性檢查）"),
  fc_wc=(0.15, "=((43.280641+51.625874+5.256702-11.319824-26.013273)/258.935494+(53.246047+51.754865+6.04674-12.370177-29.613258)/300.870903+(58.608969+52.636828+6.964529-13.03938-32.707431)/333.605938)/3", 0.35, "Derived",
         _SA + "：近三年全公司營運資金 ÷ 合併營收平均＝24.3%／23.0%／21.7%；營運資金＝應收＋存貨＋其他流動資產 − 應付 − 應計費用（不含其他流動負債）",
         "Sens 低／高 15%／35%（Assumed）"),
 ),
 "beta_u": (1.87, "產業無槓桿 β：Damodaran Global Semiconductor（現金修正後；2026 年 1 月）",
            _DAM + "：Semiconductor 675 家，β 1.88、D/E 4.13%、稅率 8.46%、無槓桿 β 1.82、現金修正後 1.87",
            "三星為集團公司，本模型盈餘以記憶體為主，故用半導體；同表 Electronics (General) 現金修正後 1.57（對照）。非美國公司用 Global 表；用現金修正值：淨現金另計"),
 "beta_extra": [("beta_u_us", "對照：Damodaran US Semiconductor 無槓桿 β（現金修正後；β 低案例）", 1.50, _DAMUS + "：Semiconductor 66 家，現金修正後無槓桿 β 1.50")],
 "beta_lo": "=beta_u_us*(1+(1-tax_w)*beta_de)", "beta_hi": None,
 "beta_lohi_note": "低＝Damodaran US Semiconductor 1.50 重新加槓桿（Analogy；與 Global Electronics (General) 1.57 相近）；高＝與低對稱（2 × 基準 − 低，Assumed）。個股迴歸 β 未取得",
}
CONFIG["tp"]["method"] = ("DCF（v0.5 標準版）：無槓桿自由現金流 FCFF＝NOPAT（合併營業利益 ×（1 − 稅率））＋折舊 − 資本支出 − 營運資金增加；2027–2029 折舊與資本支出為 Y 頁記憶體事業口徑（非記憶體事業視為相抵），2030–2032 收斂期由淨利還原營業利益、比例收斂到全公司近三年平均；"
                          "以 WACC（股權成本＝韓國 10 年期公債 rf ＋ β ×（ERP 5%＋韓國 CRP 0.64%），β 為 Damodaran Global 半導體無槓桿 β 重新加槓桿）折到 2026 年底，終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g），g＝3%、RONIC＝WACC；"
                          "加淨現金、扣少數股權 ÷ 股數，以股權成本推到 12 個月後並扣 2027 年股利（DCF 頁、tools/fcff_dcf.py）。"
                          + CONFIG["tp"]["method"][CONFIG["tp"]["method"].index("Comparable"):])
CONFIG["version"] = "v0.5"
CONFIG["a1_note"] = "；v0.4.1 只改展示層（README 展示層與命題回覆、HTML 新版面）；v0.5（2026-11-05）模型改版：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數），Comparable、情境權重、DCF 權重 50% 與投資建議規則不變"
CONFIG["obs_tpl"] = {
 2: "DCF 情境價值（v0.5 標準版 FCFF＋WACC，推到 12 個月後）A {dcfA|P}、B {dcfB|P}、C {dcfC|P}（現價的 {dcfpxA|0.00}／{dcfpxB|0.00}／{dcfpxC|0.00} 倍），對應 2028 EPS A {epsA28|E}、B {epsB28|E}、C {epsC28|E}。4Q 位元季增由 5% 上調到公司指引的 15–16%，使 2027 起營收基期上移約 3%；2026 記憶體資本支出依公司全年 84.4 兆上調 10.8 兆，部分抵銷自由現金流。",
 3: "2028 B 營業利益對一般 DRAM 收入最敏感（±20% → ±15.5%），NAND ±30% → ±9.2%、HBM ±30% → ±7.2%；WACC −1pt／＋1pt 使加權目標價 {s_wlo|+0.0%;−0.0%}／{s_whi|+0.0%;−0.0%}。待校正：9-30 資產負債表（權益、淨現金為推估）於季報公布後更新；2026 股東回饋剩餘 55–75 兆（含大規模買回註銷）待 2027-01 董事會，未計入。",
}
CONFIG["readme_tail"] = ["v0.5（2026-11-05）：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；記憶體樣板 v0.5、tools/fcff_dcf.py；docs/plan/個股DCF標準規格.md v1.0）。新增 DCF 頁（三情境與 Wattonomics 錨定路徑的 FCFF 明細、檢查列）；Val 0.1 節改為折現率與 DCF 彙總（列數不變，0.2 節以後列號不變）；Sens 目標價參數改為標準案例（WACC、β、ERP、g、RONIC、資本支出、營運資金、DCF 權重、終值替代法）；In 頁新增 M2 節（rf、β、ERP、CRP、資本結構、折舊／資本支出／營運資金穩態比例），刪除 tp_r、tp_roe，淨現金收益率鍵 rf 改名 ncy（值不變）。Y、Q、Mkt、Recon 頁與 v0.4.1 相同。已知簡化：非記憶體事業只有營業利益輸入，明確期折舊與資本支出視為相抵；營運資金只在 DCF 計入；權益投資未計入 DCF 淨現金。"]

# v0.5.1（2026-11-05）：β 改自算產業 β（Andy 2026-11-05 22:55；docs/plan/自算產業β_20261105.md 對照表）：KR 半導體 同業中位數 1.72，Blume 調整 0.67β＋0.33，依本公司 D/E 重新加槓桿；取代 Damodaran 表。其他模型內容不變。
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
CONFIG["version"] = "v0.5.1"
CONFIG["a1_note"] = CONFIG["a1_note"] + "；v0.5.1（2026-11-05）：β 改自算產業 β（KR 半導體，Blume 調整）"
CONFIG["readme_tail"] = CONFIG["readme_tail"] + ["v0.5.1（2026-11-05）：β 改自算產業 β（Andy 2026-11-05 22:55）：KR 半導體 同業無槓桿 β 中位數 {betaum|0.00}（Yahoo Finance 月報酬 2021-11～2026-10 對 韓國綜合指數 ^KS11 迴歸，Hamada 去槓桿、現金修正），Blume 調整後 {betau|0.000}，依市值 D/E 重新加槓桿為 {beta|0.00}，取代 Damodaran 表；β 低／高案例改為中位數 × 0.8 再 Blume／中位數未調整。"
    "評價日 2026-12-31 不變（已是統一評價日）。"
    "新舊對照（v0.5 → v0.5.1）：β 1.90 → {beta|0.00}，WACC 14.6% → {wacc|0.0%}，DCF 目標價 337,484 → {dcftp|P}，加權目標價 414,653 → {tpw|P}，上檔空間 +92.4% → {upw|+0.0%;−0.0%}，投資建議 強力買進 → {rec}。其他模型內容不變。"]

_m = CONFIG["tp"]["method"]; assert "β 為 Damodaran Global 半導體無槓桿 β 重新加槓桿" in _m
CONFIG["tp"]["method"] = _m.replace("β 為 Damodaran Global 半導體無槓桿 β 重新加槓桿", "β 為自算 KR 半導體產業無槓桿 β（Blume 調整）重新加槓桿")
CONFIG["dcf"]["beta_u"] = CONFIG["dcf"]["beta_u"][:3] + (CONFIG["dcf"]["beta_u"][3] + "；三星為集團公司，本模型盈餘以記憶體為主，且本身在 KR 半導體同業內",) + CONFIG["dcf"]["beta_u"][4:]

# v0.5.2（2026-11-05）：工作單第二輪修正（Andy 2026-11-05 23:16 定案；共用模組 tools/fcff_dcf.py v2、記憶體樣板 v0.5.2）
# (1) β 改「記憶體」組（MU、WDC、SK 海力士、三星電子、南亞科 5 家，無槓桿中位數 2.294 → Blume 1.867；取代 KR 半導體 1.482）；β 低／高＝β_u ∓ 0.3 重新加槓桿
# (2) 資產負債表日 2026-06-30（2Q26 半期報告：最新已公布的資產負債表；3Q26 確定財報 10-29 未附資產負債表，季報約 11 月中）；
#     淨現金滾到評價日 2026-12-31＝6-30 淨現金 ＋ 3Q（實績）＋ 4Q（模型）FCFF − 已宣告股利與買回（2Q 股利 2.45、3Q 股利約 30 兆、8-21 決議買回 15 兆；4Q 定期股利尚未宣告，不扣）
# (3) 明確期折舊／資本支出照舊讀 Y 頁（模組 v2）；(4) Sens 新增「美元口徑折現」
_BSRC2 = ("自算產業 β：記憶體 同業 5 家（美光 MU、Western Digital WDC、SK 海力士 000660、三星電子 005930、南亞科 2408；SanDisk、Kioxia 上市不足 5 年不入中位數）中位數"
          "（Yahoo Finance 月報酬 2021-11～2026-10，各自對本地大盤迴歸，Hamada 去槓桿、現金修正）")
_US10 = ("美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；2026-11-05 讀取）",
         "FRED DGS10 同序列；與股價同日（2026-10-23）；只用於 Sens「美元口徑折現」（工作單第二輪修正，Andy 2026-11-05 23:16）")
_SAQ = "stockanalysis.com Samsung Electronics（KRX:005930）季度資產負債表（https://stockanalysis.com/quote/krx/005930/financials/balance-sheet/?p=quarterly，2026-11-05 讀取；最新一季 2026-06-30，百萬韓元）"
_R3Q = "三星 3Q26 確定財報（2026-10-29；https://news.samsung.com/global/samsung-electronics-announces-third-quarter-2026-results，2026-11-05 讀取）"
CONFIG["dcf"].update({
 "beta_u": ("=0.67*beta_u_med+0.33", "產業無槓桿 β（自算記憶體組，Blume 調整後；採用值）", _BSRC2 + "；＝0.67 × 中位數 beta_u_med ＋ 0.33",
            "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；記憶體組為第二輪修正新增（2026-11-05 23:20，取代 KR 半導體 1.482）；三星為集團公司，本模型盈餘以記憶體為主；用現金修正值", "Derived"),
 "beta_extra": [("beta_u_med", "產業無槓桿 β 中位數（自算記憶體組，Blume 調整前；只作來源對照）", 2.294, _BSRC2, "Derived"),
                ("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", 0.3, "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3；取代 v0.5.1 的中位數 × 0.8 再 Blume／中位數未調整", "Assumed")],
 "beta_du": True,
 "beta_lo": "=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "beta_hi": "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)",
 "beta_lohi_note": "Sens 低／高（Analogy；v0.5.2 起）：β_u ∓ beta_du（0.3），各自同式重新加槓桿（tools/fcff_dcf.py v2 write_sens）；個股迴歸 β 未取得",
 "beta_src": _BSRC2 + "；Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；",
 "beta_obs": "自算記憶體產業無槓桿 β {betau|0.000}（同業中位數 Blume 調整）",
 "beta_sens": "β 低／高案例（β_u ∓ 0.3 各自重新加槓桿）為 {s_blo|+0.0%;−0.0%}／{s_bhi|+0.0%;−0.0%}，",
 "beta_note": "自算產業 β 重新加槓桿",
 "stub": dict(bs_date="2026-06-30", val_date="2026-12-31", bs_src=_SAQ + "（2Q26 半期報告；3Q26 確定財報未附資產負債表，季報約 11 月中公布）",
              div=("=2.45+30+bb_0", "Derived", "已宣告：2Q 定期股利 2.45 兆（7 月宣告、8 月支付）＋ 3Q 現金股利每股 4,601 韓元合計約 30 兆（" + _R3Q + "，11 月支付）＋ In 頁 bb_0（2026-08-21 董事會決議買回約 15 兆）",
                   "只扣已宣告者：4Q 定期股利（約 2.45 兆，2027-01 宣告）與 2026 股東回饋剩餘 55–75 兆（2027-01 董事會）不扣")),
 "nc_bs": "189.999028-22.408721-lease_v-mi_v",
 "nc_bs_note": _SAQ + "：Total Cash 189,999,028 − Total Debt 22,408,721 − 租賃 − 少數股權（2026-06-30）",
})
CONFIG["dcf"]["stub"]["fcff"] = ("(107.4405+Q!I10+nm_op_q)*(1-tax)+Q!H13+Q!I13-(29.0*0.75+capex_rem)-fc_wc*((Q!H4+Q!I4)-(Q!D4+Q!E4))")
# ↑ 3Q 實績合併營業利益 107.4405（_R3Q）＋ 4Q 模型（記憶體營業利益＋非記憶體 nm_op_q），×（1 − 稅率）；＋3Q、4Q 記憶體折舊；− 3Q 記憶體資本支出（29.0 × 75%）與 4Q capex_rem；
#   − fc_wc × 近兩季記憶體營收年增量（2026Q3＋Q4 − 2025Q3＋Q4；非記憶體營收季度未建模，略）
CONFIG["dcf"]["fd"] = dict(CONFIG["dcf"]["fd"],
    rf_us=(0.05, _US10[0], _US10[1]),
    cash=(189.999028, "Verified", _SAQ + "：Total Cash 189,999,028（現金 92,916,382＋短期投資 97,036,646＋其他 46,000；2026-06-30）", "DCF 淨現金以資產負債表日科目計，再以 stub 期滾到評價日；Y 頁 nc_last（9-30 推估）只用於 EPS 推估"),
    debt=(22.408721, "Verified", "2Q26 合併資產負債表：借款 22.41 兆韓元（" + _SAQ + " Total Debt 22,408,721）", "WACC 權重、β 重新加槓桿與評價淨現金皆用本值"),
    lease=("=(3879.269+1195.504)/1000", "Verified", _SAQ + "：Long-Term Leases 3,879,269＋Current Portion of Leases 1,195,504", "視為不含在借款內"),
    mi=(14.245, "Verified", _SAQ + "：Minority Interest 14,244,936（2026-06-30）", "合併營業利益含子公司少數股權份額，評價時扣除；v0.5.1 為 FY2025 12.007"),
)
CONFIG["obs_tpl"] = dict(CONFIG["obs_tpl"])
CONFIG["obs_tpl"][1] = ("目標價 {tpw|P}（DCF {dcftp|P}、Comparable（相對倍數調整後）{compadj|P}；DCF 權重 {wdcf|0%}），相對現價 {px|P}（2026-10-23）為 {upw|+0.0%;−0.0%}，投資建議{rec}。"
                        "{upsrc}：加權目標價的上檔中 DCF 貢獻 {dcfsh|0%}、Comparable 貢獻 {compsh|0%}（各自權重 ×（目標價 − 現價））；Comparable 單獨（市場主流同業法、未調整）{upc|+0.0%;−0.0%}。")
CONFIG["obs_tpl"][3] = CONFIG["obs_tpl"][3].replace("待校正：9-30 資產負債表（權益、淨現金為推估）於季報公布後更新",
                                                    "待校正：9-30 資產負債表（季報約 11 月中；DCF 淨現金暫以 6-30 資產負債表加 stub 期計，Y 頁權益與淨現金為推估）公布後更新")
assert "待校正：9-30 資產負債表（季報約 11 月中" in CONFIG["obs_tpl"][3]
CONFIG["tp"]["method"] = CONFIG["tp"]["method"].replace("β 為自算 KR 半導體產業無槓桿 β（Blume 調整）重新加槓桿", "β 為自算記憶體產業無槓桿 β（Blume 調整）重新加槓桿")
CONFIG["version"] = "v0.5.2"
CONFIG["a1_note"] = CONFIG["a1_note"] + "；v0.5.2（2026-11-05）：第二輪修正（記憶體組 β、資產負債表日 6-30 滾到評價日、β 敏感度 ±0.3、Sens 美元口徑折現）"
CONFIG["readme_tail"] = CONFIG["readme_tail"] + ["v0.5.2（2026-11-05）：工作單第二輪修正（Andy 2026-11-05 23:16 定案；共用模組 tools/fcff_dcf.py v2、記憶體樣板 v0.5.2）。"
    "(1) β 改「記憶體」組（MU、WDC、SK 海力士、三星電子、南亞科 5 家無槓桿中位數 {betaum|0.000}，Blume 調整 {betau|0.000}），依市值 D/E 重新加槓桿為 {beta|0.00}；β 敏感度改 β_u ∓ 0.3 重新加槓桿。"
    "(2) 資產負債表日 2026-06-30（2Q26 半期報告，最新已公布；3Q26 確定財報未附資產負債表）：DCF 淨現金與投資 {dcfnc|#,##0.0} 兆，滾到評價日 2026-12-31＝＋3Q（實績營業利益）與 4Q（模型）FCFF {stubB|#,##0.0} 兆 − 已宣告股利與買回 {stubdiv|#,##0.0} 兆＝{ncrB|#,##0.0} 兆；4Q 定期股利與剩餘股東回饋尚未宣告，不扣。"
    "(3) 明確期 2027–2029 折舊／資本支出照舊讀 Y 頁記憶體事業逐年列（模組 v2），2030 起收斂且資本支出 ≥ 折舊 ×（1＋g）。(4) Sens 新增「美元口徑折現」：加權目標價 {usd|P}（{usdp|+0.0%;−0.0%}）。"
    "新舊對照（v0.5.1 → v0.5.2）：β 1.50 → {beta|0.00}，WACC 12.5% → {wacc|0.0%}，DCF 目標價 360,480 → {dcftp|P}，加權目標價 426,150 → {tpw|P}，上檔空間 +97.7% → {upw|+0.0%;−0.0%}，投資建議 強力買進 → {rec}。"]

# v0.5.3（2026-11-06）：第三輪（合併前）修正（記憶體樣板 v0.5.3；共用模組 tools/fcff_dcf.py v2 不變）
# (1) 現價統一 2026-11-04 收盤（Yahoo Finance 日線收盤）；同業前瞻 P/B 以同日收盤 ÷ 共識每股淨值重算（Kioxia 與 SanDisk／Kioxia 模型同一 11-04 口徑）
# (2) bb_px 改公式連 px；(3) 非控制權益：上市子公司 Rainbow Robotics（277810.KQ，持股 35%、合併子公司）改市值；
# (4) README 版本沿革改固定文字、重點觀察 3 改公式；(5) 自算產業 β 來源改引 repo 文件 docs/plan/自算產業β_20261105.md
_PX = "Yahoo Finance 日線收盤 2026-11-04"
def _set_in(key, label=None, base=None, tag=None, src=None):
    i = next(j for j, x in enumerate(CONFIG["inputs"]) if x[0] == key)
    x = list(CONFIG["inputs"][i])
    if label is not None: x[1] = label
    if base is not None: x[4] = base
    if tag is not None: x[6] = tag
    if src is not None: x[7] = src
    CONFIG["inputs"][i] = tuple(x)
_set_in("px", label="普通股股價（2026-11-04 收盤）", base=218000, tag="Verified", src=_PX + "（第三輪修正統一價格日）；v0.5.2 為 2026-10-23 收盤 215,500")
_set_in("bb_px", base="=px", tag="Derived", src="＝目前股價（公式連 In 頁 px）；v0.5.2 為寫死的 10-08 舊價 262,000")
_MSP = "MarketScreener 共識每股淨值（2026-10-25 擷取，未更新）"
CONFIG["tp"]["peers"] = [
    ("SK 海力士", "000660.KS", "=1396000/(1277000/1.35)", "2026-11-04", f"11-04 收盤 1,396,000（{_PX}）÷ 2027E 每股淨值（＝10-23 收盤 1,277,000 ÷ 頁面 P/B 1.35；{_MSP}）"),
    ("美光", "MU", "=795.23/(877.55/3.35)", "2026-11-04", f"11-04 收盤 795.23（{_PX}）÷ FY2027E（2027-09）每股淨值（＝10-23 收盤 877.55 ÷ 頁面 P/B 3.35；{_MSP}）"),
    ("鎧俠（Kioxia）", "285A.T", "=10735/7808", "2026-11-04", f"11-04 收盤 ¥10,735（{_PX}）÷ FY2028E（2028-03）每股淨值 ¥7,808（{_MSP}；sndk/、kioxia/ 模型同值）；v0.5.2 為 10-09 頁面 P/B 2.22"),
    ("南亞科", "2408.TW", "=399.5/248.9", "2026-11-04", f"11-04 收盤 399.5（{_PX}）÷ 2027E 每股淨值 248.9（MarketScreener 2026-11-06 讀取；頁面 10-22 收盤 414 對應 P/B 1.66）")]
CONFIG["tp"]["peer_note"] = CONFIG["tp"]["peer_note"] + "；第三輪修正（v0.5.3）起報價統一 2026-11-04 收盤（Yahoo Finance），共識每股淨值不變"
_fd = CONFIG["dcf"]["fd"]
_fd["rf"] = (_fd["rf"][0], _fd["rf"][1], _fd["rf"][2].replace("取與股價（2026-10-23）最近的 10-22 值", "取 v0.5.2 現價日（2026-10-23）最近的 10-22 值；第三輪修正現價改 2026-11-04，rf 未改（不在該輪範圍）"), _fd["rf"][3])
_fd["rf_us"] = (_fd["rf_us"][0], _fd["rf_us"][1], _fd["rf_us"][2].replace("與股價同日（2026-10-23）", "v0.5.2 現價日（2026-10-23）值，第三輪修正未改"))
assert "rf 未改" in _fd["rf"][2] and "第三輪修正未改" in _fd["rf_us"][2]
# 非控制權益（第三輪修正第 1、13 項「上市子公司一律市值」）：Rainbow Robotics 為合併子公司（三星 2024-12-31 公告持股由 14.7% 增至 35%、納入合併報表；2026 半年報持股 35.00%）
_RB = ("Rainbow Robotics（277810.KQ）：三星持股 35%，2025 起為合併子公司（The Star 2025-01-01：Samsung to become largest shareholder in Rainbow Robotics…「integrated into Samsung's consolidated financial statements as a subsidiary」；"
       "https://www.thestar.com.my/business/business-news/2025/01/01/samsung-to-become-largest-shareholder-in-rainbow-robotics）；"
       "股數 19.4 百萬股（kr.investing.com）；股價 2026-11-04 收盤未取得，沿用 2026-10-26 收盤 348,500 韓元（stockhub.kr，2026-11-06 讀取）；"
       "6-30 權益 132,869 百萬韓元（stockanalysis.com 季度資產負債表）")
_fd["mi"] = ("=14.245-0.65*0.132869+0.65*19.4*0.3485", "Verified／Derived",
             "2026-06-30 帳面少數股權 14.245 兆（" + _SAQ + "）中 Rainbow Robotics 部分改市值：− 帳面 65% × 0.133 兆 ＋ 市值 65% × 19.4 百萬股 × 348,500 韓元（" + _RB + "）",
             "上市子公司一律市值、其餘（Samsung Display 等未上市）帳面；E（WACC 權重）未含非控制權益市值（模組 v2 口徑，影響 < 0.5%）")
CONFIG["dcf"]["stub"]["fcff_label"] = "2026Q3 實績＋2026Q4E 模型季度 FCFF（合併營業利益 ×（1 − 稅率）＋3Q、4Q 記憶體折舊 − 3Q、4Q 記憶體資本支出 − fc_wc × 近兩季記憶體營收年增量）"
CONFIG["obs_tpl"][1] = CONFIG["obs_tpl"][1].replace("相對現價 {px|P}（2026-10-23）", "相對現價 {px|P}（2026-11-04）")
assert "（2026-11-04）" in CONFIG["obs_tpl"][1]
_o3 = "2028 B 營業利益對一般 DRAM 收入最敏感（±20% → ±15.5%），NAND ±30% → ±9.2%、HBM ±30% → ±7.2%；"
assert CONFIG["obs_tpl"][3].startswith(_o3)
CONFIG["obs_tpl"][3] = ("{sg2} 營業利益對各收入線的敏感度：一般 DRAM 收入 +20%／−20% → {sg2_dram_hi|+0.0%;−0.0%}／{sg2_dram_lo|+0.0%;−0.0%}，"
                        "NAND ±30% → {sg2_nand_hi|+0.0%;−0.0%}／{sg2_nand_lo|+0.0%;−0.0%}、HBM ±30% → {sg2_hbm_hi|+0.0%;−0.0%}／{sg2_hbm_lo|+0.0%;−0.0%}；") + CONFIG["obs_tpl"][3][len(_o3):]
_FROZEN = {"v0.5.1（": 'v0.5.1（2026-11-05）：β 改自算產業 β（Andy 2026-11-05 22:55）：KR 半導體 同業無槓桿 β 中位數 1.72（Yahoo Finance 月報酬 2021-11～2026-10 對 韓國綜合指數 ^KS11 迴歸，Hamada 去槓桿、現金修正），Blume 調整後 1.482，依市值 D/E 重新加槓桿為 1.50，取代 Damodaran 表；β 低／高案例改為中位數 × 0.8 再 Blume／中位數未調整。評價日 2026-12-31 不變（已是統一評價日）。新舊對照（v0.5 → v0.5.1）：β 1.90 → 1.50，WACC 14.6% → 12.5%，DCF 目標價 337,484 → 360,480，加權目標價 414,653 → 426,150，上檔空間 +92.4% → +97.7%，投資建議 強力買進 → 強力買進。其他模型內容不變。',
           "v0.5.2（": 'v0.5.2（2026-11-05）：工作單第二輪修正（Andy 2026-11-05 23:16 定案；共用模組 tools/fcff_dcf.py v2、記憶體樣板 v0.5.2）。(1) β 改「記憶體」組（MU、WDC、SK 海力士、三星電子、南亞科 5 家無槓桿中位數 2.294，Blume 調整 1.867），依市值 D/E 重新加槓桿為 1.90；β 敏感度改 β_u ∓ 0.3 重新加槓桿。(2) 資產負債表日 2026-06-30（2Q26 半期報告，最新已公布；3Q26 確定財報未附資產負債表）：DCF 淨現金與投資 148.3 兆，滾到評價日 2026-12-31＝＋3Q（實績營業利益）與 4Q（模型）FCFF 115.7 兆 − 已宣告股利與買回 47.5 兆＝216.5 兆；4Q 定期股利與剩餘股東回饋尚未宣告，不扣。(3) 明確期 2027–2029 折舊／資本支出照舊讀 Y 頁記憶體事業逐年列（模組 v2），2030 起收斂且資本支出 ≥ 折舊 ×（1＋g）。(4) Sens 新增「美元口徑折現」：加權目標價 406,648（−0.9%）。新舊對照（v0.5.1 → v0.5.2）：β 1.50 → 1.90，WACC 12.5% → 14.6%，DCF 目標價 360,480 → 328,708，加權目標價 426,150 → 410,264，上檔空間 +97.7% → +90.4%，投資建議 強力買進 → 強力買進。'}
CONFIG["readme_tail"] = [next((v for k, v in _FROZEN.items() if t.startswith(k)), t) for t in CONFIG["readme_tail"]]
CONFIG["version"] = "v0.5.3"
CONFIG["date"] = "2026-11-06"
CONFIG["a1_note"] = CONFIG["a1_note"] + "；v0.5.3（2026-11-06）：第三輪修正（現價與同業報價統一 2026-11-04、買回價連現價、上市子公司非控制權益改市值、版本沿革固定文字）"
CONFIG["readme_tail"] = CONFIG["readme_tail"] + ["v0.5.3（2026-11-06）：第三輪（合併前）修正（記憶體樣板 v0.5.3）。"
    "(1) 現價改 2026-11-04 收盤 218,000（Yahoo Finance 日線收盤；v0.5.2 為 10-23 的 215,500）；同業前瞻 P/B 以同日收盤 ÷ 共識每股淨值重算（SK 海力士 1.35 → 1.48、美光 3.35 → 3.04、Kioxia 2.22（10-09 頁面）→ 1.37、南亞科 1.66 → 1.61；Kioxia 與 SanDisk、Kioxia 模型同一 11-04 口徑），同業中位數 1.94 → 1.54。"
    "(2) 未來買回均價 bb_px（原寫死 10-08 舊價 262,000）改公式連 px。"
    "(3) 非控制權益：上市合併子公司 Rainbow Robotics（277810.KQ，持股 35%）改市值（65% × 19.4 百萬股 × 348,500 韓元；11-04 收盤未取得，沿用 10-26），少數股權 14.245 → 18.553 兆；其餘子公司（未上市）帳面。"
    "(4) README 版本沿革列改為當版固定文字（原為活公式，會顯示現值）；重點觀察 3 的 Sens 數字改公式；Stub 期說明改依實際口徑（3Q 實績＋4Q 模型季度 FCFF）；情境驅動對照收斂法列名中性化、循環中段係數情境 B 佔位 0 → 1；HTML DCF 明細只有比例列用百分比。"
    "(5) 自算產業 β 來源改引 repo 文件 docs/plan/自算產業β_20261105.md（值不變）。"
    "新舊對照（v0.5.2 → v0.5.3）：β 1.90 → 1.90，DCF 目標價 328,708 → 327,937，Comparable（相對倍數調整後）491,821 → 419,399，加權目標價 410,264 → 373,668，上檔空間 +90.4% → +71.4%，投資建議 強力買進 → 強力買進。"]


# v0.5.4（2026-11-06，工作單 px1105；Andy 2026-11-06 07:28、03:14）：現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守（記憶體樣板同步改名）
# 同業前瞻 P/B＝現行倍數（11-04 基準）× 11-05 收盤 ÷ 11-04 收盤（共識每股淨值不變）；rf、共識、預測匯率、模型假設不變
_PX5 = "Yahoo Finance 日線收盤 2026-11-05"
_set_in("px", label="普通股股價（2026-11-05 收盤）", base=201500, tag="Verified", src=_PX5 + "；v0.5.3 為 2026-11-04 收盤 218,000")
CONFIG["tp"]["peers"] = [
    ("SK 海力士", "000660.KS", "=1396000/(1277000/1.35)*1264000/1396000", "2026-11-05", f"現行 P/B（11-04 收盤 1,396,000 ÷ 2027E 每股淨值，＝10-23 收盤 1,277,000 ÷ 頁面 P/B 1.35；{_MSP}）× 11-05 收盤 1,264,000 ÷ 11-04 收盤 1,396,000（{_PX5}）"),
    ("美光", "MU", "=795.23/(877.55/3.35)*656.39/795.23", "2026-11-05", f"現行 P/B（11-04 收盤 795.23 ÷ FY2027E（2027-09）每股淨值，＝10-23 收盤 877.55 ÷ 頁面 P/B 3.35；{_MSP}）× 11-05 收盤 656.39 ÷ 11-04 收盤 795.23（{_PX5}）"),
    ("鎧俠（Kioxia）", "285A.T", "=10735/7808*9324/10735", "2026-11-05", f"現行 P/B（11-04 收盤 ¥10,735 ÷ FY2028E（2028-03）每股淨值 ¥7,808；{_MSP}；sndk/、kioxia/ 模型同值）× 11-05 收盤 ¥9,324 ÷ 11-04 收盤 ¥10,735（{_PX5}）"),
    ("南亞科", "2408.TW", "=399.5/248.9*360/399.5", "2026-11-05", f"現行 P/B（11-04 收盤 399.5 ÷ 2027E 每股淨值 248.9，MarketScreener 2026-11-06 讀取）× 11-05 收盤 360 ÷ 11-04 收盤 399.5（{_PX5}）")]
CONFIG["tp"]["peer_note"] = CONFIG["tp"]["peer_note"] + "；v0.5.4 起報價改 2026-11-05 收盤（Yahoo Finance；現行倍數 × 11-05 ÷ 11-04 收盤），共識每股淨值不變"
CONFIG["obs_tpl"][1] = CONFIG["obs_tpl"][1].replace("相對現價 {px|P}（2026-11-04）", "相對現價 {px|P}（2026-11-05）")
assert "（2026-11-05）" in CONFIG["obs_tpl"][1]
_RB5 = _RB.replace("股價 2026-11-04 收盤未取得，沿用 2026-10-26 收盤 348,500 韓元（stockhub.kr，2026-11-06 讀取）", "股價 2026-11-05 收盤 334,500 韓元（" + _PX5 + "；v0.5.3 沿用 10-26 收盤 348,500）")
assert _RB5 != _RB
_fd["mi"] = ("=14.245-0.65*0.132869+0.65*19.4*0.3345", "Verified／Derived",
             "2026-06-30 帳面少數股權 14.245 兆（" + _SAQ + "）中 Rainbow Robotics 部分改市值：− 帳面 65% × 0.133 兆 ＋ 市值 65% × 19.4 百萬股 × 334,500 韓元（" + _RB5 + "）",
             _fd["mi"][3])
CONFIG["version"] = "v0.5.4"
CONFIG["date"] = "2026-11-06"
CONFIG["a1_note"] = CONFIG["a1_note"] + "；v0.5.4（2026-11-06）：現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守"
CONFIG["readme_tail"] = CONFIG["readme_tail"] + ["v0.5.4：現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守"]

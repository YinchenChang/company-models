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
            "情境對應：A＝順利、B＝平平、C＝不如預期（原為 TrendForce／BofA 型價格情境）。"),
}
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.3（2026-10-25）：Val 頁新增「0. 目標價」（DCF 與 Comparable 加權，tp_w_dcf 0.5）；原「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」，情境對應改為 A 75 分位、B 中位數、C 25 分位 × 各情境 2027 年底每股淨值。Y 頁三情境延伸到 2032（2030 起收斂法）。現價更新為 2026-10-23 收盤 215,500。2026–2028 EPS 與 v0.2.2 相同；2029 因修正共用工具「前一年」基期（原誤用 2027）而改變。收斂期淨利率收斂到循環中段（十年平均 × 1.2／1.0／0.8）。"]

# v0.3.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00、14:17；樣板 v0.4.1，數值不變）
CONFIG["version"] = "v0.3.1"
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.3.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁最後新增「上檔空間級距」「投資建議」兩列（依「上檔空間：加權目標價」固定規則）；README 頁最後新增投資建議與重點觀察三點。"]
CONFIG["obs"] = [
 "目標價 ₩477,143（DCF 486,542、Comparable 467,744 各半），比現價 215,500（2026-10-23）高 121.4%，級距 大於 50%；Comparable 單獨為 +117.1%。主因是倍數不是 EPS：本模型 2028 年底每股淨值 241,105 只比共識 232,055 高 3.9%（前瞻 EPS 78,917 比共識高 5.2%），但現價只隱含 0.89 倍前瞻 P/B，記憶體同業中位數 1.94 倍、共識目標價 478,906 隱含 2.06 倍；同業倍數未扣三星的集團折價，Comparable 偏高估。",
 "DCF 情境價值 A 818,535、B 394,410、C 218,882（現價的 3.80／1.83／1.02 倍），對應 2028 EPS A 129,982、B 59,278、C 36,498；收斂期用的 ROE 為 2023–2025 平均 8.0%，低於折現率 10%，再投資率約五成，壓低 2030 年後的股東現金流。",
 "2028 B 營業利益對一般 DRAM 收入最敏感（±20% → ±15.2%），NAND ±30% → ±9.2%、HBM ±30% → ±7.3%；折現率 ±1pt 使加權目標價 −4.1%／+5.6%。下一個檢驗事件是 3Q26 確定財報與法說（約 2026-10-30，日期未確認；2026-10-08 初步財報已納入）。"]   # 重點觀察（與 README.md 相同文字）

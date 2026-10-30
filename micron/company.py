# 美光設定檔（樣板 v0.2 複製測試）。只放參數
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))), "siliconomics"))
import reader as _SIL   # HBM 營收份額 sh25–sh29 讀 Siliconomics M24（v0.8 起；數值權威在本庫 In 頁）
S = "sec"
GM = [38.4, 36.8, 37.7, 44.7, 56.0, 74.4, 84.6, 86.8]
REV = [8.709, 8.053, 9.301, 11.315, 13.643, 23.860, 41.456, 54.229]
DRAM = [6.400, 6.123, 7.071, 8.984, 10.812, 18.768, 31.328, 39.8]
NAND = [2.241, 1.855, 2.155, 2.252, 2.743, 4.997, 9.943, 14.1]
OP = [2.174, 1.773, 2.169, 3.654, 6.136, 16.135, 33.318, 43.751]
NI = [1.870, 1.583, 1.885, 3.201, 5.240, 13.785, 28.243, 37.701]
DA = [2.030, 2.079, 2.094, 2.149, 2.212, 2.286, 2.364, 2.641]
LAB = ["FQ1-25", "FQ2-25", "FQ3-25", "FQ4-25", "FQ1-26", "FQ2-26", "FQ3-26", "FQ4-26"]
COL = "BCDEFGHI"
CONFIG = {
 "name": "美光", "version": "v0.4", "date": "2026-10-25",
 "unit": "十億美元", "ccy_per_usdB": "1", "share_scale": 1000,
 "base_fy": 2026, "fy_offset": 1, "fy_label": "FY{}", "fy_desc": "會計年度約 9 月初結束；FY27 ≈ 2026Q4–2027Q3（日曆季對齊為近似，實際差約 1 個月）",
 "thesis": "美光的盈餘由一般 DRAM 與 NAND 的合約價循環決定，HBM 份額約與其 DRAM 份額相當（約兩成）；股價取決於循環見頂的年度與見頂時累積的淨值，以及「超額現金 100% 回饋」把淨值轉成買回的速度。",
 "q_source": "SEC 8-K 財報新聞稿、10-Q、10-K、CFO 書面說明",
 "segments": {
  "title": "官方報告分部（事業單位，FQ4-25 起新架構；8-K 財報新聞稿）",
  "rows": [
   ("CMBU 雲端記憶體（含 HBM）", [2.648, 2.947, 3.382, 4.543, 5.284, 7.749, 13.769, 16.283]),
   ("CDBU 核心資料中心", [2.292, 1.830, 1.531, 1.577, 2.379, 5.687, 11.524, 18.002]),
   ("MCBU 行動與用戶端", [2.608, 2.236, 3.256, 3.760, 4.255, 7.711, 11.521, 13.114]),
   ("AEBU 車用與嵌入式", [1.158, 1.034, 1.124, 1.434, 1.720, 2.708, 4.634, 6.824]),
  ],
  "check_label": "檢查：分部合計 − 合併營收",
  "notes": ["分部來源：FQ1-26、FQ2-26、FQ4-26 新聞稿（各含前期比較數，FQ1-25、FQ2-25、FQ4-25 為公司重編）；FQ3-25 當季新聞稿未揭露新架構 [Derived]＝FY2025 全年（FY2026 10-K 第 1 項，取整至千萬美元）− 其他三季。",
            "差額為四捨五入與未揭露的「其他」（每季 ≤ 0.008）。推估年度不拆分部：模型以技術別（一般 DRAM／HBM／NAND）推估；HBM（Y 頁）屬 CMBU。"],
 },
 "q_note": "營業成本＝營收 × (1 − GAAP 毛利率)；FQ4-26 為 14 週，DRAM／NAND 依書面說明 73%／26%（10-K 未出）",
 "hbm_share": {2025: "sh25", 2026: "sh26", 2027: "sh27", 2028: "sh28", 2029: "sh29"},
 "pb_names": ["P/B 1.0x（谷底）", "P/B 歷史中位 1.4x", "P/B 前高 3.0x"], "pe_vals": [4, 8, 12],
 "quarters": [dict(label=LAB[i], type="A", rev=REV[i], dram=DRAM[i], nand=NAND[i], cogs=f"={COL[i]}4*(1-{GM[i]/100})", op=OP[i], ni=NI[i], da=DA[i]) for i in range(8)],
 "inputs": [
  (S, "A. 股價"),
  ("px", "股價（2026-10-23 收盤）", "USD", None, 877.55, None, "Verified", "markets/invest_map/data/prices.csv（stockscan、bullfincher）；MarketScreener 2026-10-23 收盤同；v0.3.2 為 10-07 的 1,088"),
  ("fx", "（美元公司不使用匯率）", "—", None, 1, None, "Verified", ""),
  (S, "B. 季度推估用（美光基準年 FY26 全為實績，以下不使用，保留樣板欄位）"),
  ("costdn_q", "單位成本季降", "%", None, 0.03, None, "Assumed", "不使用"),
  ("opex_fix_q", "營業費用固定部分（每季）", "十億美元", None, 1.25, None, "Assumed", "不使用"),
  ("nonop_q", "每季業外淨收益", "十億美元", None, 0.6, None, "Assumed", "不使用"),
  (S, "C. 年度驅動（公司別）"),
  ("cb_1", "一般 DRAM 位元成長 FY27", "%", 0.06, 0.12, 0.18, "Verified／Derived", "美光：產業 DRAM 位元 CY27 低 20% 段（含 HBM）；HBM 吃掉部分晶圓"),
  ("cb_2", "一般 DRAM 位元成長 FY28", "%", 0.10, 0.15, 0.20, "Verified／Derived", "同上"),
  ("nb_1", "NAND 位元成長 FY27", "%", 0.12, 0.18, 0.25, "Derived", "美光 FQ4-26 法說：產業 CY27、CY28 皆中 20% 段；美光供給低於產業（本模型取值）。2026-10-28 來源複核：原標 Verified 不當"),
  ("nb_2", "NAND 位元成長 FY28", "%", 0.15, 0.20, 0.25, "Derived", "同上"),
  ("sh25", "HBM 營收市占 2025", "%", *_SIL.vals("hbmsh_mu_25"), "Interested-party", "Counterpoint 3Q25 21%；美光：與 DRAM 市占相當"),
  ("sh26", "HBM 營收市占 2026", "%", *_SIL.vals("hbmsh_mu_26"), "Interested-party", "Counterpoint Q1 21%、Q2 18%"),
  ("sh27", "HBM 營收市占 2027", "%", *_SIL.vals("hbmsh_mu_27"), "Interested-party", "UBS 2027 位元 20%"),
  ("sh28", "HBM 營收市占 2028", "%", *_SIL.vals("hbmsh_mu_28"), "Assumed", ""),
  ("other_y", "其他營收（每年）", "十億美元", None, 0.7, None, "Derived", "FY26 營收 − DRAM − NAND"),
  ("w_c", "成本權重：一般 DRAM", "%", None, 0.50, None, "Assumed", ""),
  ("w_h", "成本權重：HBM", "%", None, 0.22, None, "Assumed", "HBM 3:1 晶圓（美光）"),
  ("w_n", "成本權重：NAND", "%", None, 0.28, None, "Assumed", "NAND 約 24% 營收、毛利較低"),
  ("costdn_y", "單位成本年降", "%", 0.05, 0.10, 0.15, "Assumed", ""),
  ("da_life", "新增資本支出轉為折舊的年數", "年", 4, 5, 6, "Assumed", ""),
  ("da_cogs", "折舊計入營業成本比例", "%", None, 0.9, None, "Assumed", ""),
  ("opex_fix_y", "營業費用固定部分（年，基準年）", "十億美元", None, 5.0, None, "Derived", "FY26 營業費用 7.6（GAAP）；FY27 指引約 +2.5（研發、獎金）"),
  ("opex_var", "營業費用變動率", "% 營收", 0.01, 0.015, 0.03, "Derived", "使 FY27 營業費用約 9.3–10（指引口徑）"),
  ("opex_g", "營業費用固定部分年增", "%", None, 0.15, None, "Derived", "研發擴張"),
  ("tax", "有效稅率", "%", 0.14, 0.155, 0.18, "Verified", "FY26 14.8%；FY27 指引約 15.5%"),
  ("rf", "淨現金收益率", "%", None, 0.04, None, "Assumed", "美元利率"),
  ("cap_0", "資本支出 FY26（淨）", "十億美元", None, 27.367, None, "Verified", "FY26 8-K（毛 30.7）"),
  ("cap_1", "資本支出 FY27", "十億美元", 50, 55, 65, "Derived", "FQ1 約 11.5、上半年約 25、下半年更高（書面說明）"),
  ("cap_2", "資本支出 FY28", "十億美元", 50, 60, 75, "Assumed", ""),
  ("cb_3", "一般 DRAM 位元成長 FY29", "%", 0.10, 0.15, 0.20, "Assumed", ""),
  ("nb_3", "NAND 位元成長 FY29", "%", 0.15, 0.20, 0.25, "Assumed", ""),
  ("sh29", "HBM 營收市占 2029", "%", *_SIL.vals("hbmsh_mu_29"), "Assumed", ""),
  ("cap_3", "資本支出 FY29", "十億美元", 50, 60, 75, "Assumed", ""),
  ("dps_3", "每股股利 FY29", "USD", None, 0.85, None, "Assumed", ""),
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
  ("eq_last", "FY26 末權益", "十億美元", None, 138.378, None, "Verified", "FQ4-26 8-K"),
  ("nc_last", "FY26 末淨現金", "十億美元", None, 68.3, None, "Derived", "現金與投資 73.48 − 負債 5.18"),
  ("sh_last", "流通股數", "百萬股", None, 1129.4, None, "Verified", "10-Q 封面 2026-06-17（FY26 末數未出）"),
  ("extra_eq", "基準年其餘增資", "十億美元", None, 0, None, "Verified", "無"),
  ("extra_sh", "基準年其餘新股", "百萬股", None, 0, None, "Verified", "無"),
  ("bb_0", "基準年其餘買回", "十億美元", None, 0, None, "Verified", "基準年全為實績"),
  ("bb_sh_0", "基準年其餘買回股數", "百萬股", None, 0, None, "Verified", ""),
  ("div_0", "基準年其餘股利", "十億美元", None, 0, None, "Verified", ""),
  ("capex_rem", "基準年其餘資本支出", "十億美元", None, 0, None, "Verified", ""),
  ("dps_1", "每股股利 FY27", "USD", None, 0.60, None, "Verified", "每季 $0.15（2026-03 起）"),
  ("dps_2", "每股股利 FY28", "USD", None, 0.72, None, "Assumed", "+20%"),
  ("payout", "股東回饋 ÷ FCF", "%", 0.5, 0.8, 1.0, "Verified／Assumed", "公司：2026-12-09 起逐步把「超額現金 100%」回饋；取 80%"),
  ("bb_px", "未來買回均價", "USD", None, 1088, None, "Assumed", "＝目前股價"),
  (S, "E. 估值倍數"),
  ("pb_1", "P/B 1.0x（谷底）", "x", None, 1.0, None, "Interested-party", "歷史谷底 0.52–1.32（Macrotrends）"),
  ("pb_2", "P/B 歷史中位", "x", None, 1.4, None, "Interested-party", "2010–2025 年 8 月底 P/B 中位約 1.35"),
  ("pb_3", "P/B 前高", "x", None, 3.0, None, "Interested-party", "前高 3.02（2014）、3.11（2024）"),
  ("pe_1", "P/E 4x", "x", None, 4, None, "Assumed", "目前 FY27 共識 P/E 約 6.2x"),
  ("pe_2", "P/E 8x", "x", None, 8, None, "Assumed", ""),
  ("pe_3", "P/E 12x", "x", None, 12, None, "Assumed", ""),
 ],
 "recon": [
  ("FY26 營收（實績核對）", "=Y!C{rev}", 133.188, "FQ4-26 8-K"),
  ("FY27 營收（A）vs 共識", "=Y!D{rev}", 275.0, "Yahoo／LSEG（39 位分析師）2026-10-07"),
  ("FY27 淨利（A）vs 共識（EPS 176.15 × 約 1.15B 股）", "=Y!D{ni}", 202.6, "Yahoo／LSEG [Derived]"),
  ("FY28 營收（A）vs 共識", "=Y!E{rev}", 319.1, "Yahoo／LSEG"),
  ("FY28 淨利（A）vs 共識（EPS 206.33 × 1.15B）", "=Y!E{ni}", 237.3, "Yahoo／LSEG [Derived]"),
  ("FY27 平均每季營收（A）vs FQ1-27 指引", "=Y!D{rev}/4", 61.5, "公司指引 ±1.5；管理層：FQ1 毛利率為 FY27 低點"),
 ],
 "sources": [
  ("FQ4 FY26 財報新聞稿（2026-09-30）", "https://www.sec.gov/Archives/edgar/data/723125/000072312526000018/a2026q4ex991-pressrelease.htm"),
  ("FQ4 FY26 書面說明", "https://s25.q4cdn.com/621799436/files/doc_financials/2026/q4/Q4-FY26-Prepared-Remarks.pdf"),
  ("FQ4 FY26 法說逐字稿", "https://www.investing.com/news/transcripts/earnings-call-transcript-micron-tops-q4-2026-estimates-as-demand-stays-hot-93CH-4925992"),
  ("FQ3 FY26 新聞稿", "https://www.sec.gov/Archives/edgar/data/723125/000072312526000013/a2026q3ex991-pressrelease.htm"),
  ("FQ2 FY26 新聞稿", "https://www.sec.gov/Archives/edgar/data/723125/000072312526000004/a2026q2ex991-pressrelease.htm"),
  ("FQ1 FY26 新聞稿", "https://www.sec.gov/Archives/edgar/data/723125/000072312525000044/a2026q1ex991-pressrelease.htm"),
  ("10-Q FQ3 FY26", "https://www.sec.gov/Archives/edgar/data/723125/000072312526000015/mu-20260528.htm"),
  ("10-K FY25", "https://www.sec.gov/Archives/edgar/data/723125/000072312525000028/mu-20250828.htm"),
  ("Yahoo Finance 共識", "https://finance.yahoo.com/quote/MU/analysis/"),
  ("Macrotrends P/B", "https://www.macrotrends.net/stocks/charts/MU/micron-technology/price-book"),
 ],
 "readme_extra": ["vv0.3.1：Val 頁新增「十年倍數與規則值」區塊（倍數歷史、分位、規則值與隱含股價；原放在 markets/overview/extract.py），總覽與 HTML 展示層改讀此處；損益、EPS 與原有儲存格不變。", "v0.3：HBM 出貨位元隨情境變動（markets/memory）：2028–29 需求位元 × 情境顆數係數（B 0.8／0.85、C 0.6／0.6；與台積電、Advantest 同一套），供給與 2028 前的單價假設不變。2026–27 與 v0.2.1 相同；情境 A 全部相同；B、C 的 2028–29 HBM 營收下降，營業成本量指數隨 HBM 位元同步下降。", "v0.2.1：Tokenomics v19.0 同步。只有對帳頁的「每顆 GPU 的 NVIDIA 內容」由 51.6 改 42.8 千美元（橋接表 v0.2：GB300 機架基準價 $4.3M）；損益、淨值、估值與 v0.2 相同。", "v0.2：預測延長到 FY2029，情境 C 的谷底（FY28–29）已可見；前兩年與 v0.1 相同。", "複製測試：本檔只換參數（設定檔 micron/company.py）；樣板結構不變。為了美光而改的樣板結構（會計年度、幣別、季價路徑）已先回頭改進樣板 v0.2，SK 海力士 v0.2 同步重建。",
                  "已知簡化：會計年度與日曆季對齊差約 1 個月；HBM 營收美光未揭露（由市場 × 市占推得）；FQ4-26 為 14 週。"],
}

# 估值規則的十年倍數歷史（會計年度底，2016–2025；規則見 tools/valuation_rule.py；值原放在 markets/overview/extract.py，2026-10-23 移入）
CONFIG["rule"] = {
 "pb": {"sources": [
   {"label": "macrotrends（會計年度 8 月底）", "vals": [1.29, 1.78, 1.79, 1.33, 1.27, 1.83, 1.22, 1.72, 2.35, 2.46], "tag": "Verified"},
   {"label": "companiesmarketcap（日曆年底，參考不採用）", "vals": [1.83, 2.03, 1.02, 1.64, 2.10, 2.27, 1.11, 2.20, 2.11, 6.04], "tag": "Verified", "use": False, "ref_rows": True}]},
 # SEC 10-K XBRL：P/B（macrotrends）、權益 $B、庫藏股（成本）$B、歸屬淨利 $B；P/E＝P/B × 權益 ÷ 淨利、買回加回 P/B＝P/B × 權益 ÷（權益＋庫藏股）
 "sec": [(2016, 1.29, 12.080, 1.029, -0.276), (2017, 1.78, 18.621, 0.067, 5.089), (2018, 1.79, 32.294, 0.429, 14.135),
         (2019, 1.33, 35.881, 3.221, 6.313), (2020, 1.27, 38.996, 3.495, 2.687), (2021, 1.83, 43.933, 4.695, 5.861),
         (2022, 1.22, 49.907, 7.127, 8.687), (2023, 1.72, 44.120, 7.552, -5.833), (2024, 2.35, 45.131, 7.852, 0.778),
         (2025, 2.46, 54.165, 7.852, 8.539)],
}

# v0.3.2（2026-10-25）：目標價與上檔空間（Andy 2026-10-25 要求；Val 頁最後一節、In 頁最後三列；既有儲存格不變）
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.3.2（2026-10-25）：Val 頁新增「0. 目標價與上檔空間」（放在 Val 既有內容之後，不位移既有列）：目標價＝In 頁權重 tp_wA／tp_wB／tp_wC（預設 A 0.3／B 0.5／C 0.2，Assumed，Andy 可改）× 三個 P/B 規則值（A＝中段、B＝見頂、C＝谷底）；既有儲存格的值不變。"]
CONFIG["tp_bias"] = "十年 P/B 含 AI 前的低倍數期，且大量買回壓低帳面每股淨值（第 6 節另列買回加回 P/B 對照），偏誤方向為低估。"

# v0.4（2026-10-25）：目標價改為 DCF 與 Comparable 加權（Andy 11:28、11:53）；三情境延伸到 FY2032（收斂法）；現價更新到 10-23
MS_EST = "MarketScreener 預估頁（2026-10-25 擷取；頁面為 10-23 收盤；會計年度 9 月）"
CONFIG["tp"] = {
 "roe_label": "ROE（近三年平均，FY2024–FY2026）", "roe": "=AVERAGE(0.0174,0.1720,0.8826)",
 "roe_src": "stockanalysis 比率頁：FY2024 1.74%、FY2025 17.20%、FY2026 88.26%（單一來源）",
 "nim10": "=AVERAGE(-0.276/12.399,5.089/20.322,14.135/30.391,6.313/23.406,2.687/21.435,5.861/27.705,0.2824,-0.3754,0.0310,0.2284)",
 "nim10_src": "FY2016–FY2025：淨利為 SEC 10-K（Val 第 5 節）；營收 FY2016–FY2021 為公司 10-K（本次未重新開啟原文，單一來源），FY2022–FY2025 淨利率為 stockanalysis（28.24%、−37.54%、3.10%、22.84%）",
 "cons_tp": (1560.57, 361, 3000, "MarketScreener 共識頁：49 位分析師平均 1,560.57（最低 361、最高 3,000），2026-10-25 擷取（頁面收盤日 10-16）"),
 "cons": [("cons_eps27", "共識 EPS FY2027", 173, MS_EST), ("cons_eps28", "共識 EPS FY2028", 201.1, MS_EST), ("cons_bvps28", "共識每股淨值 FY2028 年底", 418.4, MS_EST)],
 "cons_src": MS_EST,
 "peers": [("SK 海力士", "000660.KS", 1.35, "2026-10-23", "MarketScreener：現價 ÷ 2027E 每股淨值"),
           ("三星電子", "005930.KS", 1.26, "2026-10-23", "MarketScreener：現價 ÷ 2027E 每股淨值"),
           ("鎧俠（Kioxia）", "285A.T", 2.22, "2026-10-09", "MarketScreener：現價 ÷ FY2028E（2028-03）每股淨值；頁面價格為 10-09"),
           ("南亞科", "2408.TW", 1.66, "2026-10-22", "MarketScreener：現價 ÷ 2027E 每股淨值")],
 "peer_note": "記憶體同業（SK 海力士、三星、鎧俠、南亞科；無 A 股）。前瞻 P/B＝現價 ÷ 下一個會計年度末（約 2027 年底）共識每股淨值；鎧俠會計年度 3 月。美光自身前瞻 P/B 3.35x 遠高於同業（大量買回壓低帳面淨值）",
 "fwd_w": ("0", "1"), "fwd_note": "會計年度 9 月初結束：約 2027-10 時的未來 12 月 ≈ FY2028（2027-09～2028-08），故取 FY2028 EPS",
 "dcf_note": "；美光 FY2027 約於 2027-09 結束，t 以會計年度計（近似）",
 "method": ("DCF：股東現金流 FY2027–FY2029 用每股（淨利＋折舊－資本支出），FY2030–FY2032 收斂期用 EPS ×（1 − g／ROE），r 10%、g 4%、ROE 為 FY2024–FY2026 平均約 36%；明確期＋永續，基準日 2026 年底。"
            "Comparable：SK 海力士、三星、鎧俠、南亞科前瞻 P/B 中位數 × FY2028 年底每股淨值（情境加權）。"
            "收斂期（FY2030–FY2032）淨利率收斂到循環中段：A 十年平均 × 1.2、B 十年平均、C 十年平均 × 0.8（In 頁 cyc_mA／cyc_mC），永續期以 FY2032 中段水準為基礎。偏誤：FY2027–FY2029 仍為高峰期盈餘，DCF 仍可能偏高估；美光大量買回使帳面淨值偏低，同業 P/B 套用在美光淨值上，Comparable 偏低估。"
            "情境對應：A＝順利、B＝平平、C＝不如預期（原為 TrendForce／BofA 型價格情境）。"),
}
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.4（2026-10-25）：Val 頁新增「0. 目標價」（DCF 與 Comparable 加權，tp_w_dcf 0.5）；原「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」，情境對應改為 A 75 分位、B 中位數、C 25 分位 × 各情境 FY2027 年底每股淨值。Y 頁三情境延伸到 FY2032（FY2030 起收斂法）。現價更新為 2026-10-23 收盤 877.55。FY2026–FY2028 EPS 與 v0.3.2 相同；FY2029 因修正共用工具「前一年」基期（原誤用 FY2027）而改變。收斂期淨利率收斂到循環中段（十年平均 × 1.2／1.0／0.8）。"]

# v0.4.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00、14:17；樣板 v0.4.1，數值不變）
CONFIG["version"] = "v0.4.1"
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.4.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁最後新增「上檔空間級距」「投資建議」兩列（依「上檔空間：加權目標價」固定規則）；README 頁最後新增投資建議與重點觀察三點。"]
CONFIG["obs"] = [
 "目標價 US$684（DCF 933、Comparable 436 各半），比現價 877.55（2026-10-23）低 22.0%，級距 0%～−25%；Comparable 單獨為 −50.4%。兩法差距來自倍數：記憶體同業前瞻 P/B 中位數 1.51 倍 × FY2028 年底每股淨值 289.40，遠低於共識目標價 1,560.57 隱含的 3.73 倍（比同業高 148%）；美光自身前瞻 P/B 3.35 倍，因大量買回壓低帳面淨值。",
 "本模型比共識保守：FY2028 前瞻 EPS（情境加權）157.33 比共識 201.1 低 21.8%、FY2028 年底每股淨值比共識 418.4 低 30.8%；情境 DCF 價值 A 1,714、B 708、C 322（現價的 1.95／0.81／0.37 倍），分歧集中在 FY2028 以後的 DRAM 價格路徑。",
 "FY2028 B 營業利益對一般 DRAM 收入最敏感（±20% → ±16.5%），NAND ±30% → ±8.3%、HBM ±30% → ±6.5%；DCF 權重改 0.3／0.7 使加權目標價 −14.5%／+14.5%。下一個檢驗事件是 FQ1 FY2027 財報（約 2026-12 下旬）的 DRAM 合約價與 HBM 出貨。"]   # 重點觀察（與 README.md 相同文字）

# v0.4.2（2026-10-28）：分部原則重建（Q 頁新增官方報告分部與核對列；樣板 v0.4.2 選用區塊，數值不變）
CONFIG["version"] = "v0.4.2"
CONFIG["readme_extra"] = CONFIG["readme_extra"] + ["v0.4.2（2026-10-28）：分部原則重建。Q 頁第 20 列起新增官方報告分部（CMBU／CDBU／MCBU／AEBU，FQ1-25～FQ4-26 實績）與核對列「分部合計 − 合併營收」（每季 ≤ 0.008，四捨五入與未揭露的其他）；推估年度仍以技術別推估、不拆分部，HBM 屬 CMBU。損益、EPS、目標價與 v0.4.1 相同。"]

# v0.4.3（2026-10-30）：In 頁版面（個股模型原則第 8 節；數值不變）
CONFIG["version"] = "v0.4.3"

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
            "情境對應：A＝樂觀、B＝基準、C＝保守（原為 TrendForce／BofA 型價格情境）。"),
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

# v0.4.4（2026-11-05）：只改展示層（個股 HTML 版面規格 v1.0；README 新增「展示層」「命題回覆」兩段，公式帶數字；其他頁數值不變）
CONFIG["in_layout_version"] = "v0.4.3"   # In 頁版面說明的版本標籤（v0.4.3 套用；之後升版不改）
CONFIG["version"] = "v0.4.4"
CONFIG["date"] = "2026-11-05"
CONFIG["a1_note"] = "；v0.4.4 只改展示層（README 新增展示層與命題回覆兩段、HTML 新版面），其他頁與 v0.4.3 相同"
CONFIG["pres"] = {
 "lead": "命題：美光的盈餘由一般 DRAM 與 NAND 的合約價循環決定，HBM 份額約與其 DRAM 份額相當；股價取決於循環見頂的年度、見頂時累積的淨值，以及「超額現金全數回饋」把淨值轉成買回的速度。三個問題如下。",
 "props": [
  "命題 1：哪條收入線決定盈餘：一般 DRAM、NAND 與 HBM 的收入變動，各自使營業利益變動多少。",
  "命題 2：循環在哪一年見頂：三種價格情境下，盈餘在哪一年到高點、之後回落多少。",
  "命題 3：見頂時累積的淨值與買回撐不撐得住股價：大量買回把淨值轉成更少的股數，用同業倍數與現金流折現估算，每股價值是否高於現價。"],
 "obs": [
  "重點 1：合理價值{updesc}現價，兩種方法差距很大。我們估每股約 {tpw|P} 美元（DCF {dcftp|P} 占 {wdcf|0%}、Comparable〔相對倍數調整後〕{compadj|P} 占其餘），相對現價 {px|P} 為 {upw|+0.0%;−0.0%}，投資建議{rec}。DCF 比現價 {dcfup|+0%;−0%}，Comparable 比現價 {upcadj|+0%;−0%}：同業前瞻 P/B 經相對倍數調整後 {pmadj|0.00} 倍，套在 {y28} 年底每股淨值（情境加權）{fbv|P} 上；目標價方法已註明，美光大量買回壓低帳面淨值，用同業 P/B 估值可能偏低。",
  "重點 2：三種情境的淨利高點落在 {peakA}（情境 A）、{peakB}（情境 B）、{peakC}（情境 C）；分歧集中在 {y28} 以後：{y28} 每股盈餘情境 A 約 {epsA28|E}、B 約 {epsB28|E}、C 約 {epsC28|E} 美元。情境 DCF 價值為現價的 {dcfpxA|0.00}／{dcfpxB|0.00}／{dcfpxC|0.00} 倍（A／B／C）。",
  "重點 3：營業利益主要看{s_top} 價格。以 {s_year} 情境 {s_scn} 為例，{s_dram_l} 使營業利益變動 {s_dram|+0.0%;−0.0%}；{s_nand_l} 為 {s_nand|+0.0%;−0.0%}、{s_hbm_l} 為 {s_hbm|+0.0%;−0.0%}（各收入線的測試幅度不同，是敏感度測試，不是門檻）。"],
 "ans": [
  "一般 DRAM、NAND 與 HBM 哪條收入線對盈餘影響最大？答：{s_top}。以 {s_year} 情境 {s_scn} 為例，一般 DRAM 營收約 {convA27|#,##0.0} 十億美元、NAND 約 {nandA27|#,##0.0}、HBM 約 {hbmA27|#,##0.0} 十億美元（HBM 占記憶體營收 {hbmshA27|0%}）；{s_dram_l} 使營業利益變動 {s_dram|+0.0%;−0.0%}，{s_nand_l} 為 {s_nand|+0.0%;−0.0%}，{s_hbm_l} 為 {s_hbm|+0.0%;−0.0%}。",
  "循環在哪一年見頂？答：依本模型三種情境的價格路徑（情境設定，不是單一預測），淨利高點在情境 A {peakA}、情境 B {peakB}、情境 C {peakC}。情境 B 高點淨利約 {nipkB|#,##0.0} 十億美元，{y28} 為 {niB28|#,##0.0} 十億美元（比高點 {nichgB28|+0%;−0%}）；情境 C {y28} 為 {niC28|#,##0.0} 十億美元（比高點 {nichgC28|+0%;−0%}）。",
  "見頂時累積的淨值與買回撐不撐得住股價？答：依本模型{verdict}。買回使情境 B 的期末股數由基準年到 {y29} 變動 {shchgB29|+0.0%;−0.0%}，帳面淨值因此偏低：現價 {px|P} 美元相當於 {y27} 年底每股淨值（情境加權）的 {pxbv27|0.00} 倍，記憶體同業同口徑的前瞻 P/B 中位數 {pmed|0.00} 倍。目標價把同業倍數（經相對倍數調整後 {pmadj|0.00} 倍）套在 12 個月後的前瞻每股淨值（{y28} 年底，情境加權）{fbv|P} 上，得 Comparable 目標價 {compadj|P}（比現價 {upcadj|+0.0%;−0.0%}）；DCF 目標價 {dcftp|P}（比現價 {dcfup|+0.0%;−0.0%}）。加權目標價 {tpw|P}，比現價 {upw|+0.0%;−0.0%}，投資建議{rec}；{dcfall}（情境 C 為現價 {dcfpxC|0.00} 倍）。"],
}

# v0.5（2026-11-05）：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；樣板 v0.5、tools/fcff_dcf.py；規格 docs/plan/個股DCF標準規格.md v1.0）。
# Comparable、情境權重、DCF 權重 50% 與投資建議規則不變；舊 DCF 參數 tp_r、tp_roe 刪除；淨現金收益率鍵 rf 改名 ncy（值不變）。
_DAMUS = "Damodaran Betas by Sector (US) 2026 年 1 月（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/Betas.html；2026-11-05 讀取）"
_DAM = "Damodaran Betas by Sector (Global) 2026 年 1 月（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/BetasGlobal.html；2026-11-05 讀取）"
_SA = "stockanalysis.com Micron（MU）財報頁，2026-11-05 讀取；百萬美元"
CONFIG["dcf"] = {
 "frac0": 1.0, "nc_incl_lease": True,
 "capex_note": "",
 "fd": dict(
  date="2026-09-03（FY2026 年底；淨現金錨點；與 v0.4 DCF 同樣以第 1 年＝FY2027 折現）",
  rf=(0.0500, "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；與 NVIDIA v0.3 同一來源）",
      "FRED DGS10（https://fred.stlouisfed.org/series/DGS10）同序列；與股價同日（2026-10-23）", "無風險利率（美國 10 年期公債殖利率，2026-10-23）"),
  crp=(0.0, "Assumed", "美國：0（同 NVIDIA v0.3；Damodaran 2026-01 表的美國 CRP 為 0.23%，未採用）", "—"),
  kd=("=rf+0.01", "Assumed", "rf ＋ 1.0pt 信用利差（BBB 級；公司債利率未逐筆查）", "負債占資本 < 1%，對 WACC 影響可忽略"),
  debt=(2.792, "Verified", _SA + "：FY2026 長期債務 2,792（總債務 5,179＝長期債務 2,792＋流動租賃 491＋長期租賃 1,896）", "In 頁 nc_last 68.3＝現金與投資 73.48 − 負債 5.18（含租賃）"),
  lease=(2.387, "Verified", _SA + "：FY2026 租賃負債 491＋1,896", "公司淨現金（nc_last）已扣租賃，故評價日現金加回租賃，避免重複扣除"),
  inv=(0, "Assumed", "無重大權益投資", "—"),
  mi=(0, "Verified", _SA + "：未報少數股權", "—"),
  dps=("=dps_1", "Derived", "＝In 頁 dps_1（FY2027 每股股利）", "評價日到目標時點約一年"),
  t_tgt=(1.0, "目標價為評價日後 12 個月（約 2027-09；與 v0.4 DCF 目標價同一時點口徑）"),
  fc_da=(None, "=(7.78/25.111+8.352/37.378+9.503/133.188)/3", None, "Derived",
         _SA + "：近三年（FY2024–FY2026）折舊攤銷 ÷ 營收平均＝31.0%／22.3%／7.1%（D&A 7,780／8,352／9,503；營收 25,111／37,378／133,188）",
         "記憶體樣板：明確期 FY2027–FY2029 用 Y 頁折舊；本值只作 FY2030 起收斂的長期穩態"),
  fc_cx=(None, "=(8.386/25.111+15.857/37.378+30.712/133.188)/3", None, "Derived",
         _SA + "：近三年（FY2024–FY2026）資本支出（毛額）÷ 營收平均＝33.4%／42.4%／23.1%（8,386／15,857／30,712）",
         "明確期用 Y 頁資本支出（In 頁 cap_1–cap_3）；本值只作 FY2030 起長期穩態；穩態資本支出 ≥ 折舊 ×（1＋g）（DCF 頁一致性檢查）"),
  fc_wc=(0.20, "=((6.615+8.875+0.765-2.726-1.117)/25.111+(9.265+8.355+0.91-3.132-1.116)/37.378+(36.197+10.372+1.04-3.518-4.447)/133.188)/3", 0.50, "Derived",
         _SA + "：近三年營運資金 ÷ 營收平均＝49.4%／38.2%／29.8%；營運資金＝應收＋存貨＋其他流動資產 − 應付 − 應計費用（不含其他流動負債）",
         "Sens 低／高 20%／50%（Assumed）"),
 ),
 "beta_u": (1.50, "產業無槓桿 β：Damodaran US Semiconductor（現金修正後；2026 年 1 月）",
            _DAMUS + "：Semiconductor 66 家，β 1.52、D/E 2.59%、稅率 5.11%、無槓桿 β 1.49、現金修正後 1.50", "同 NVIDIA v0.3；用現金修正值：淨現金另計，β 只對營業資產"),
 "beta_extra": [("beta_u_gl", "對照：Damodaran Global Semiconductor 無槓桿 β（現金修正後；β 高案例）", 1.87, _DAM + "：Semiconductor 675 家，現金修正後無槓桿 β 1.87")],
 "beta_lo": None, "beta_hi": "=beta_u_gl*(1+(1-tax_w)*beta_de)",
 "beta_lohi_note": "高＝Damodaran Global Semiconductor 1.87 重新加槓桿（Analogy）；低＝與高對稱（2 × 基準 − 高，Assumed）。個股迴歸 β 未取得",
}
CONFIG["tp"]["method"] = ("DCF（v0.5 標準版）：無槓桿自由現金流 FCFF＝NOPAT（營業利益 ×（1 − 稅率））＋折舊 − 資本支出 − 營運資金增加，FY2027–FY2029 用 Y 頁折舊與資本支出，FY2030–FY2032 收斂期由淨利還原營業利益、折舊與資本支出比例收斂到近三年平均；"
                          "以 WACC（股權成本＝美國 10 年期公債 rf ＋ β × ERP 5%，β 為 Damodaran 美國半導體無槓桿 β 重新加槓桿）折到 FY2026 年底，終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g），g＝3%、RONIC＝WACC；"
                          "加淨現金 ÷ 股數，以股權成本推到 12 個月後並扣 FY2027 股利（DCF 頁、tools/fcff_dcf.py）。"
                          + CONFIG["tp"]["method"][CONFIG["tp"]["method"].index("Comparable"):])
CONFIG["version"] = "v0.5"
CONFIG["a1_note"] = "；v0.4.4 只改展示層（README 展示層與命題回覆、HTML 新版面）；v0.5（2026-11-05）模型改版：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數），Comparable、情境權重、DCF 權重 50% 與投資建議規則不變"
CONFIG["obs_tpl"] = {
 2: "本模型比共識保守：12 個月後的前瞻 EPS（情境加權）{feps|P} 比共識同口徑 {geps|+0.0%;−0.0%}、FY2028 年底每股淨值（情境加權）比共識 {gbv|+0.0%;−0.0%}；情境 DCF 價值（v0.5 標準版 FCFF＋WACC，推到 12 個月後）A {dcfA|P}、B {dcfB|P}、C {dcfC|P}（現價的 {dcfpxA|0.00}／{dcfpxB|0.00}／{dcfpxC|0.00} 倍），分歧集中在 FY2028 以後的 DRAM 價格路徑。",
 3: "FY2028 B 營業利益對一般 DRAM 收入最敏感（±20% → ±16.5%），NAND ±30% → ±8.3%、HBM ±30% → ±6.5%；DCF 權重改 0.3／0.7 使加權目標價 {s_d03|+0.0%;−0.0%}／{s_d07|+0.0%;−0.0%}，WACC −1pt／＋1pt 為 {s_wlo|+0.0%;−0.0%}／{s_whi|+0.0%;−0.0%}。下一個檢驗事件是 FQ1 FY2027 財報（約 2026-12 下旬）的 DRAM 合約價與 HBM 出貨。",
}
CONFIG["readme_tail"] = ["v0.5（2026-11-05）：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；記憶體樣板 v0.5、tools/fcff_dcf.py；docs/plan/個股DCF標準規格.md v1.0）。新增 DCF 頁（三情境與 Wattonomics 錨定路徑的 FCFF 明細、檢查列）；Val 0.1 節改為折現率與 DCF 彙總（列數不變，0.2 節以後列號不變）；Sens 目標價參數改為標準案例（WACC、β、ERP、g、RONIC、資本支出、營運資金、DCF 權重、終值替代法）；In 頁新增 M2 節（rf、β、ERP、CRP、資本結構、折舊／資本支出／營運資金穩態比例），刪除 tp_r、tp_roe，淨現金收益率鍵 rf 改名 ncy（值不變）。Y、Q、Mkt、Recon 頁與 v0.4.4 相同。"]
CONFIG["readme_extra"] = [(t + "營運資金只在 DCF 以 fc_wc × 營收增量計入（Y 頁淨現金未計）。" if t.startswith("已知簡化") else t) for t in CONFIG["readme_extra"]]

# v0.5.1（2026-11-05）：β 改自算產業 β（Andy 2026-11-05 22:55；docs/plan/自算產業β_20261105.md 對照表）：US 半導體 同業中位數 2.32，Blume 調整 0.67β＋0.33，依本公司 D/E 重新加槓桿；取代 Damodaran 表。其他模型內容不變。
_BSRC = "自算產業 β：US 半導體 同業 13 家（NVDA AMD AVGO MRVL QCOM INTC MU ARM CRDO ALAB RMBS MTSI GFS）中位數（Yahoo Finance 月報酬 2021-11～2026-10，對 S&P 500 ^GSPC 迴歸，Hamada 去槓桿、現金修正）"
CONFIG["dcf"].update({
 "beta_u": ("=0.67*beta_u_med+0.33", "產業無槓桿 β（自算 US 半導體，Blume 調整後；採用值）", _BSRC + "；＝0.67 × 中位數 beta_u_med ＋ 0.33",
            "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；用現金修正值：淨現金另計，β 只對營業資產", "Derived"),
 "beta_extra": [("beta_u_med", "產業無槓桿 β 中位數（自算 US 半導體，Blume 調整前；β 低／高案例用）", 2.32, _BSRC + "；β 低案例＝中位數 × 0.8 再 Blume 調整；高案例＝中位數未調整（Analogy）", "Derived")],
 "beta_lo": "=(0.67*0.8*beta_u_med+0.33)*(1+(1-tax_w)*beta_de)", "beta_hi": "=beta_u_med*(1+(1-tax_w)*beta_de)",
 "beta_lohi_note": "低＝本組 Blume 前中位數 × 0.8 再 Blume、高＝本組 Blume 前中位數（未調整），重新加槓桿（Analogy）；個股迴歸 β 未取得",
 "beta_src": _BSRC + "；Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；",
 "beta_obs": "自算 US 半導體產業無槓桿 β {betau|0.000}（同業中位數 Blume 調整）",
 "beta_sens": "β 低／高案例（本組中位數 × 0.8 再 Blume／中位數未調整）為 {s_blo|+0.0%;−0.0%}／{s_bhi|+0.0%;−0.0%}，",
})
CONFIG["version"] = "v0.5.1"
CONFIG["a1_note"] = CONFIG["a1_note"] + "；v0.5.1（2026-11-05）：β 改自算產業 β（US 半導體，Blume 調整）"
CONFIG["readme_tail"] = CONFIG["readme_tail"] + ["v0.5.1（2026-11-05）：β 改自算產業 β（Andy 2026-11-05 22:55）：US 半導體 同業無槓桿 β 中位數 {betaum|0.00}（Yahoo Finance 月報酬 2021-11～2026-10 對 S&P 500 ^GSPC 迴歸，Hamada 去槓桿、現金修正），Blume 調整後 {betau|0.000}，依市值 D/E 重新加槓桿為 {beta|0.00}，取代 Damodaran 表；β 低／高案例改為中位數 × 0.8 再 Blume／中位數未調整。"
    "評價日由 2026-09-03（FY2026 年底）統一為 2026-12-31（FY2027 只計評價日後 8 個月；目標時點 2027-12-31；淨現金與股數沿用 FY2026 年底 2026-09-03 資產負債表）。"
    "新舊對照（v0.5 → v0.5.1）：β 1.51 → {beta|0.00}，WACC 12.5% → {wacc|0.0%}，DCF 目標價 525.00 → {dcftp|P}，加權目標價 546.79 → {tpw|P}，上檔空間 −37.7% → {upw|+0.0%;−0.0%}，投資建議 放空 → {rec}。其他模型內容不變。"]

_m = CONFIG["tp"]["method"]; assert "β 為 Damodaran 美國半導體無槓桿 β 重新加槓桿）折到 FY2026 年底" in _m
CONFIG["tp"]["method"] = _m.replace("β 為 Damodaran 美國半導體無槓桿 β 重新加槓桿）折到 FY2026 年底", "β 為自算 US 半導體產業無槓桿 β（Blume 調整）重新加槓桿）折到 2026-12-31（FY2027 只計評價日後 8 個月；淨現金沿用 FY2026 年底資產負債表）")
CONFIG["dcf"]["frac0"] = 8 / 12   # 評價日 2026-12-31：FY2027（2026-09～2027-08）剩 8 個月
CONFIG["dcf"]["fd"] = dict(CONFIG["dcf"]["fd"],
    date="2026-12-31（統一評價日；FY2027 只計評價日後 8 個月；淨現金與股數沿用 FY2026 年底 2026-09-03）",
    dps=("=dps_1", "Derived", "＝In 頁 dps_1（FY2027 每股股利）", "評價日 2026-12-31 到目標時點 2027-12-31 約一年（以 FY2027 每股股利近似）"),
    t_tgt=(1.0, "目標價為評價日後 12 個月（2027-12-31）"))

# v0.5.2（2026-11-05）：工作單第二輪修正（Andy 2026-11-05 23:16 定案；共用模組 tools/fcff_dcf.py v2、記憶體樣板 v0.5.2）
# (1) β 改「記憶體」組（MU、WDC、SK 海力士、三星電子、南亞科 5 家，無槓桿中位數 2.294 → Blume 1.867；取代 US 半導體 1.884）；β 低／高＝β_u ∓ 0.3 重新加槓桿
# (2) 資產負債表日 2026-09-03（FY2026 10-K／FQ4 新聞稿，最新已公布一季）；淨現金滾到評價日 2026-12-31＝＋FY2027 FCFF × stub_frac − 已宣告股利（每股 0.15 美元，9-30 宣告、10-29 發放）
# (3) 明確期折舊／資本支出照舊讀 Y 頁（模組 v2）。美國公司不加 Sens「美元口徑折現」（＝基準）
_BSRC2 = ("自算產業 β：記憶體 同業 5 家（美光 MU、Western Digital WDC、SK 海力士 000660、三星電子 005930、南亞科 2408；SanDisk、Kioxia 上市不足 5 年不入中位數）中位數"
          "（Yahoo Finance 月報酬 2021-11～2026-10，各自對本地大盤迴歸，Hamada 去槓桿、現金修正）")
_PR4 = "Micron FQ4-26 新聞稿（8-K Ex.99.1，2026-09-30；https://www.sec.gov/Archives/edgar/data/0000723125/000072312526000018/a2026q4ex991-pressrelease.htm，2026-11-05 讀取）"
_10K = "Micron FY2026 10-K（https://www.sec.gov/Archives/edgar/data/0000723125/000072312526000023/mu-20260903.htm，2026-11-05 讀取）"
# 股數維持 v0.5.1 的 1,129.4 百萬股（10-Q 封面 2026-06-17）：10-K 封面 1,131.4 百萬股（2026-10-02）差 0.2%，改動會牽動 Y 頁 EPS／每股淨值與 markets/overview 總覽資料，留待下次滾入 FQ1-27 時一併更新
CONFIG["dcf"].update({
 "beta_u": ("=0.67*beta_u_med+0.33", "產業無槓桿 β（自算記憶體組，Blume 調整後；採用值）", _BSRC2 + "；＝0.67 × 中位數 beta_u_med ＋ 0.33",
            "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；記憶體組為第二輪修正新增（2026-11-05 23:20，取代 US 半導體 1.884）；用現金修正值：淨現金另計，β 只對營業資產", "Derived"),
 "beta_extra": [("beta_u_med", "產業無槓桿 β 中位數（自算記憶體組，Blume 調整前；只作來源對照）", 2.294, _BSRC2, "Derived"),
                ("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", 0.3, "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3；取代 v0.5.1 的中位數 × 0.8 再 Blume／中位數未調整", "Assumed")],
 "beta_du": True,
 "beta_lo": "=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "beta_hi": "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)",
 "beta_lohi_note": "Sens 低／高（Analogy；v0.5.2 起）：β_u ∓ beta_du（0.3），各自同式重新加槓桿（tools/fcff_dcf.py v2 write_sens）；個股迴歸 β 未取得",
 "beta_src": _BSRC2 + "；Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；",
 "beta_obs": "自算記憶體產業無槓桿 β {betau|0.000}（同業中位數 Blume 調整）",
 "beta_sens": "β 低／高案例（β_u ∓ 0.3 各自重新加槓桿）為 {s_blo|+0.0%;−0.0%}／{s_bhi|+0.0%;−0.0%}，",
 "beta_note": "自算產業 β 重新加槓桿",
 "stub": dict(bs_date="2026-09-03", val_date="2026-12-31", bs_src=_PR4 + "：FY2026 年底（2026-09-03）資產負債表；" + _10K,
              div=("=0.15*sh_v", "Derived", _PR4 + "：2026-09-30 宣告季度股利每股 0.15 美元（登記日 2026-10-14、2026-10-29 發放）× 評價用股數 sh_v",
                   "只扣已宣告者；下一次股利（FQ1-27 財報，約 2026-12 下旬）尚未宣告；買回無已宣告金額")),
 "nc_bs": "73.48-5.18-mi_v",
 "nc_bs_note": _PR4 + "：現金與投資 73.48（現金 38.364＋短期投資 5.070＋長期有價投資 30.019＋受限現金 0.027）− 債務 5.18（流動 0.491＋長期 4.688，含租賃）− 少數股權（US$B，2026-09-03；與 In 頁 nc_last 同口徑、四捨五入到 0.01）",
})
CONFIG["dcf"]["fd"] = dict(CONFIG["dcf"]["fd"],
    date="2026-12-31（統一評價日；FY2027 只計評價日後 8 個月；資產負債表日 2026-09-03，淨現金以 stub 期滾到評價日）")
CONFIG["obs_tpl"] = dict(CONFIG["obs_tpl"])
CONFIG["obs_tpl"][1] = ("目標價 {tpw|P}（DCF {dcftp|P}、Comparable（相對倍數調整後）{compadj|P}；DCF 權重 {wdcf|0%}），相對現價 {px|P}（2026-10-23）為 {upw|+0.0%;−0.0%}，投資建議{rec}。"
                        "加權目標價相對現價的差距中 DCF 占 {dcfsh|0%}、Comparable 占 {compsh|0%}（各自權重 ×（目標價 − 現價））；Comparable 單獨（市場主流同業法、未調整）{upc|+0.0%;−0.0%}。")
CONFIG["version"] = "v0.5.2"
CONFIG["a1_note"] = CONFIG["a1_note"] + "；v0.5.2（2026-11-05）：第二輪修正（記憶體組 β、資產負債表日 2026-09-03 滾到評價日、β 敏感度 ±0.3）"
CONFIG["readme_tail"] = CONFIG["readme_tail"] + ["v0.5.2（2026-11-05）：工作單第二輪修正（Andy 2026-11-05 23:16 定案；共用模組 tools/fcff_dcf.py v2、記憶體樣板 v0.5.2）。"
    "(1) β 改「記憶體」組（MU、WDC、SK 海力士、三星電子、南亞科 5 家無槓桿中位數 {betaum|0.000}，Blume 調整 {betau|0.000}），依市值 D/E 重新加槓桿為 {beta|0.00}；β 敏感度改 β_u ∓ 0.3 重新加槓桿。"
    "(2) 資產負債表日 2026-09-03（FY2026 10-K）：DCF 淨現金與投資 {dcfnc|#,##0.0}，滾到評價日 2026-12-31＝＋stub 期 FCFF（FY2027 FCFF × stub_frac {stubfr|0.000} 年；情境 B {stubB|#,##0.0}）− 已宣告股利 {stubdiv|#,##0.00}＝{ncrB|#,##0.0}（US$B）；股數維持 1,129.4 百萬股（10-K 封面 1,131.4 百萬股差 0.2%，下次滾入時更新）。"
    "(3) 明確期 FY2027–FY2029 折舊／資本支出照舊讀 Y 頁逐年列（模組 v2），FY2030 起收斂且資本支出 ≥ 折舊 ×（1＋g）。美國公司不加「美元口徑折現」案例（＝基準）。"
    "新舊對照（v0.5.1 → v0.5.2）：β 1.89 → {beta|0.00}，WACC 14.4% → {wacc|0.0%}，DCF 目標價 481.47 → {dcftp|P}，加權目標價 525.02 → {tpw|P}，上檔空間 −40.2% → {upw|+0.0%;−0.0%}，投資建議 放空 → {rec}。"]

CONFIG["tp"]["method"] = CONFIG["tp"]["method"].replace("β 為自算 US 半導體產業無槓桿 β（Blume 調整）重新加槓桿", "β 為自算記憶體產業無槓桿 β（Blume 調整）重新加槓桿")
assert "β 為自算記憶體產業無槓桿 β（Blume 調整）重新加槓桿" in CONFIG["tp"]["method"]

# v0.5.3（2026-11-06）：第三輪（合併前）修正（記憶體樣板 v0.5.3；共用模組 tools/fcff_dcf.py v2 不變）
# (1) 現價統一 2026-11-04 收盤（Yahoo Finance 日線收盤）；同業前瞻 P/B 以同日收盤 ÷ 共識每股淨值重算（Kioxia 與 SanDisk／Kioxia 模型同一 11-04 口徑）
# (2) bb_px 改公式連 px；(3) README 版本沿革改固定文字、重點觀察 3 與展示層重點 1 改公式；(4) 自算產業 β 來源改引 repo 文件 docs/plan/自算產業β_20261105.md
_PX = "Yahoo Finance 日線收盤 2026-11-04"
def _set_in(key, label=None, base=None, tag=None, src=None):
    i = next(j for j, x in enumerate(CONFIG["inputs"]) if x[0] == key)
    x = list(CONFIG["inputs"][i])
    if label is not None: x[1] = label
    if base is not None: x[4] = base
    if tag is not None: x[6] = tag
    if src is not None: x[7] = src
    CONFIG["inputs"][i] = tuple(x)
_set_in("px", label="股價（2026-11-04 收盤）", base=795.23, tag="Verified", src=_PX + "（第三輪修正統一價格日）；v0.5.2 為 2026-10-23 收盤 877.55")
_set_in("bb_px", base="=px", tag="Derived", src="＝目前股價（公式連 In 頁 px）；v0.5.2 為寫死的 10-07 舊價 1,088")
_MSP = "MarketScreener 共識每股淨值（2026-10-25 擷取，未更新）"
CONFIG["tp"]["peers"] = [
    ("SK 海力士", "000660.KS", "=1396000/(1277000/1.35)", "2026-11-04", f"11-04 收盤 1,396,000（{_PX}）÷ 2027E 每股淨值（＝10-23 收盤 1,277,000 ÷ 頁面 P/B 1.35；{_MSP}）"),
    ("三星電子", "005930.KS", "=218000/(215500/1.26)", "2026-11-04", f"11-04 收盤 218,000（{_PX}）÷ 2027E 每股淨值（＝10-23 收盤 215,500 ÷ 頁面 P/B 1.26；{_MSP}）"),
    ("鎧俠（Kioxia）", "285A.T", "=10735/7808", "2026-11-04", f"11-04 收盤 ¥10,735（{_PX}）÷ FY2028E（2028-03）每股淨值 ¥7,808（{_MSP}；sndk/、kioxia/ 模型同值）；v0.5.2 為 10-09 頁面 P/B 2.22"),
    ("南亞科", "2408.TW", "=399.5/248.9", "2026-11-04", f"11-04 收盤 399.5（{_PX}）÷ 2027E 每股淨值 248.9（MarketScreener 2026-11-06 讀取；頁面 10-22 收盤 414 對應 P/B 1.66）")]
CONFIG["tp"]["peer_note"] = CONFIG["tp"]["peer_note"] + "；第三輪修正（v0.5.3）起報價統一 2026-11-04 收盤（Yahoo Finance），共識每股淨值不變"
_fd = CONFIG["dcf"]["fd"]
_fd["rf"] = (_fd["rf"][0], _fd["rf"][1], _fd["rf"][2].replace("與股價同日（2026-10-23）", "v0.5.2 現價日（2026-10-23）值；第三輪修正現價改 2026-11-04，rf 未改（不在該輪範圍）"), _fd["rf"][3])
assert "rf 未改" in _fd["rf"][2]
CONFIG["obs_tpl"][1] = CONFIG["obs_tpl"][1].replace("相對現價 {px|P}（2026-10-23）", "相對現價 {px|P}（2026-11-04）")
assert "（2026-11-04）" in CONFIG["obs_tpl"][1]
_o3 = "FY2028 B 營業利益對一般 DRAM 收入最敏感（±20% → ±16.5%），NAND ±30% → ±8.3%、HBM ±30% → ±6.5%；"
assert CONFIG["obs_tpl"][3].startswith(_o3)
CONFIG["obs_tpl"][3] = ("{sg2} 營業利益對各收入線的敏感度：一般 DRAM 收入 +20%／−20% → {sg2_dram_hi|+0.0%;−0.0%}／{sg2_dram_lo|+0.0%;−0.0%}，"
                        "NAND ±30% → {sg2_nand_hi|+0.0%;−0.0%}／{sg2_nand_lo|+0.0%;−0.0%}、HBM ±30% → {sg2_hbm_hi|+0.0%;−0.0%}／{sg2_hbm_lo|+0.0%;−0.0%}；") + CONFIG["obs_tpl"][3][len(_o3):]
_p1 = CONFIG["pres"]["obs"][0]; assert _p1.startswith("重點 1：合理價值{updesc}現價，兩種方法差距很大。")
CONFIG["pres"]["obs"][0] = ("重點 1：合理價值{updesc}現價。我們估每股約 {tpw|P} 美元（DCF {dcftp|P} 占 {wdcf|0%}、Comparable〔相對倍數調整後〕{compadj|P} 占其餘），相對現價 {px|P} 為 {upw|+0.0%;−0.0%}，投資建議{rec}。"
    "{dirsame}：DCF 比現價 {dcfup|+0%;−0%}、Comparable 比現價 {upcadj|+0%;−0%}；加權目標價相對現價的差距中 DCF 占 {dcfsh|0%}、Comparable 占 {compsh|0%}（各自權重 ×（目標價 − 現價））。"
    "Comparable 為同業前瞻 P/B 經相對倍數調整後 {pmadj|0.00} 倍，套在 {y28} 年底每股淨值（情境加權）{fbv|P} 上；美光大量買回壓低帳面淨值，用同業 P/B 估值可能偏低（目標價方法已註明）。")
_FROZEN = {"v0.5.1（": 'v0.5.1（2026-11-05）：β 改自算產業 β（Andy 2026-11-05 22:55）：US 半導體 同業無槓桿 β 中位數 2.32（Yahoo Finance 月報酬 2021-11～2026-10 對 S&P 500 ^GSPC 迴歸，Hamada 去槓桿、現金修正），Blume 調整後 1.884，依市值 D/E 重新加槓桿為 1.89，取代 Damodaran 表；β 低／高案例改為中位數 × 0.8 再 Blume／中位數未調整。評價日由 2026-09-03（FY2026 年底）統一為 2026-12-31（FY2027 只計評價日後 8 個月；目標時點 2027-12-31；淨現金與股數沿用 FY2026 年底 2026-09-03 資產負債表）。新舊對照（v0.5 → v0.5.1）：β 1.51 → 1.89，WACC 12.5% → 14.4%，DCF 目標價 525.00 → 481.47，加權目標價 546.79 → 525.02，上檔空間 −37.7% → −40.2%，投資建議 放空 → 放空。其他模型內容不變。',
           "v0.5.2（": 'v0.5.2（2026-11-05）：工作單第二輪修正（Andy 2026-11-05 23:16 定案；共用模組 tools/fcff_dcf.py v2、記憶體樣板 v0.5.2）。(1) β 改「記憶體」組（MU、WDC、SK 海力士、三星電子、南亞科 5 家無槓桿中位數 2.294，Blume 調整 1.867），依市值 D/E 重新加槓桿為 1.88；β 敏感度改 β_u ∓ 0.3 重新加槓桿。(2) 資產負債表日 2026-09-03（FY2026 10-K）：DCF 淨現金與投資 68.3，滾到評價日 2026-12-31＝＋stub 期 FCFF（FY2027 FCFF × stub_frac 0.326 年；情境 B 32.0）− 已宣告股利 0.17＝100.2（US$B）；股數維持 1,129.4 百萬股（10-K 封面 1,131.4 百萬股差 0.2%，下次滾入時更新）。(3) 明確期 FY2027–FY2029 折舊／資本支出照舊讀 Y 頁逐年列（模組 v2），FY2030 起收斂且資本支出 ≥ 折舊 ×（1＋g）。美國公司不加「美元口徑折現」案例（＝基準）。新舊對照（v0.5.1 → v0.5.2）：β 1.89 → 1.88，WACC 14.4% → 14.3%，DCF 目標價 481.47 → 514.75，加權目標價 525.02 → 541.67，上檔空間 −40.2% → −38.3%，投資建議 放空 → 放空。'}
CONFIG["readme_tail"] = [next((v for k, v in _FROZEN.items() if t.startswith(k)), t) for t in CONFIG["readme_tail"]]
CONFIG["version"] = "v0.5.3"
CONFIG["date"] = "2026-11-06"
CONFIG["a1_note"] = CONFIG["a1_note"] + "；v0.5.3（2026-11-06）：第三輪修正（現價與同業報價統一 2026-11-04、買回價連現價、版本沿革固定文字）"
CONFIG["readme_tail"] = CONFIG["readme_tail"] + ["v0.5.3（2026-11-06）：第三輪（合併前）修正（記憶體樣板 v0.5.3）。"
    "(1) 現價改 2026-11-04 收盤 795.23（Yahoo Finance 日線收盤；v0.5.2 為 10-23 的 877.55）；同業前瞻 P/B 以同日收盤 ÷ 共識每股淨值重算（SK 海力士 1.35 → 1.48、三星 1.26 → 1.27、Kioxia 2.22（10-09 頁面）→ 1.37、南亞科 1.66 → 1.61；Kioxia 與 SanDisk、Kioxia 模型同一 11-04 口徑），同業中位數 1.51 → 1.43。"
    "(2) 未來買回均價 bb_px（原寫死 10-07 舊價 1,088）改公式連 px。"
    "(3) README 版本沿革列改為當版固定文字（原為活公式，會顯示現值）；重點觀察 3 的 Sens 數字與展示層重點 1（DCF／Comparable 占比）改公式；命題回覆 3 補「三種情境都低於現價」分支；情境驅動對照收斂法列名中性化、循環中段係數情境 B 佔位 0 → 1；HTML DCF 明細只有比例列用百分比。"
    "(4) 自算產業 β 來源改引 repo 文件 docs/plan/自算產業β_20261105.md（值不變）。"
    "新舊對照（v0.5.2 → v0.5.3）：β 1.88 → 1.88，DCF 目標價 514.75 → 514.79，Comparable（相對倍數調整後）568.58 → 583.43，加權目標價 541.67 → 549.11，上檔空間 −38.3% → −30.9%，投資建議 放空 → 放空。"]


# v0.5.4（2026-11-06，工作單 px1105；Andy 2026-11-06 07:28、03:14）：現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守（記憶體樣板同步改名）
# 同業前瞻 P/B＝現行倍數（11-04 基準）× 11-05 收盤 ÷ 11-04 收盤（共識每股淨值不變）；rf、共識、預測匯率、模型假設不變
_PX5 = "Yahoo Finance 日線收盤 2026-11-05"
_set_in("px", label="股價（2026-11-05 收盤）", base=656.39, tag="Verified", src=_PX5 + "；v0.5.3 為 2026-11-04 收盤 795.23")
CONFIG["tp"]["peers"] = [
    ("SK 海力士", "000660.KS", "=1396000/(1277000/1.35)*1264000/1396000", "2026-11-05", f"現行 P/B（11-04 收盤 1,396,000 ÷ 2027E 每股淨值，＝10-23 收盤 1,277,000 ÷ 頁面 P/B 1.35；{_MSP}）× 11-05 收盤 1,264,000 ÷ 11-04 收盤 1,396,000（{_PX5}）"),
    ("三星電子", "005930.KS", "=218000/(215500/1.26)*201500/218000", "2026-11-05", f"現行 P/B（11-04 收盤 218,000 ÷ 2027E 每股淨值，＝10-23 收盤 215,500 ÷ 頁面 P/B 1.26；{_MSP}）× 11-05 收盤 201,500 ÷ 11-04 收盤 218,000（{_PX5}）"),
    ("鎧俠（Kioxia）", "285A.T", "=10735/7808*9324/10735", "2026-11-05", f"現行 P/B（11-04 收盤 ¥10,735 ÷ FY2028E（2028-03）每股淨值 ¥7,808；{_MSP}；sndk/、kioxia/ 模型同值）× 11-05 收盤 ¥9,324 ÷ 11-04 收盤 ¥10,735（{_PX5}）"),
    ("南亞科", "2408.TW", "=399.5/248.9*360/399.5", "2026-11-05", f"現行 P/B（11-04 收盤 399.5 ÷ 2027E 每股淨值 248.9，MarketScreener 2026-11-06 讀取）× 11-05 收盤 360 ÷ 11-04 收盤 399.5（{_PX5}）")]
CONFIG["tp"]["peer_note"] = CONFIG["tp"]["peer_note"] + "；v0.5.4 起報價改 2026-11-05 收盤（Yahoo Finance；現行倍數 × 11-05 ÷ 11-04 收盤），共識每股淨值不變"
CONFIG["obs_tpl"][1] = CONFIG["obs_tpl"][1].replace("相對現價 {px|P}（2026-11-04）", "相對現價 {px|P}（2026-11-05）")
assert "（2026-11-05）" in CONFIG["obs_tpl"][1]
CONFIG["version"] = "v0.5.4"
CONFIG["date"] = "2026-11-06"
CONFIG["a1_note"] = CONFIG["a1_note"] + "；v0.5.4（2026-11-06）：現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守"
CONFIG["readme_tail"] = CONFIG["readme_tail"] + ["v0.5.4：現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守"]

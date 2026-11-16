# 野村總合研究所 NRI（4307）事實總帳：研究包 facts_nri.csv 與建置腳本共用（只放公司與市場事實；單位百萬日圓，除非另註）。
# 擷取方式：WebFetch 讀 PDF 後由摘要模型轉錄（非逐字人工核對）；同一數字在兩份文件互見者已互證（見 CHK）。擷取日 2026-11-16。
# 用法：python3 software_ai/nri/facts_nri.py  → 寫 software_ai/nri/data/facts_nri.csv
import csv, os
HERE = os.path.dirname(os.path.abspath(__file__))
EXT = "2026-11-16"
SRC = {
 "FS26": "https://ir.nri.com/en/ir/library/financial/main/011111114/teaserItems2/05/linkList/02/link/20260424_financialstatements_e.pdf｜NRI 2026年3月期 決算短信〔IFRS〕(連結)（2026-04-24）",
 "FS25": "https://www2.jpx.co.jp/disc/43070/140120250422519942.pdf｜NRI 2025年3月期 決算短信〔IFRS〕(連結)（2025-04-24）",
 "FS24": "https://www2.jpx.co.jp/disc/43070/140120240422573620.pdf｜NRI 2024年3月期 決算短信〔IFRS〕(連結)（2024-04-25）",
 "FS23": "https://equity.jiji.com/storage/tdnet/140120230424551670.pdf｜NRI 2023年3月期 決算短信〔IFRS〕(連結)（2023-04-27）",
 "H1": "https://ir.nri.com/en/ir/news/auto_20261027556550/pdfFile.pdf｜NRI Consolidated Financial Results for the Six Months Ended September 30, 2026（2026-10-29）",
 "H1P": "https://ir.nri.com/en/ir/library/financial/main/011111115/teaserItems2/03/linkList/0/link/20261029_financialresultspresentation_e.pdf｜NRI 2Q FY3/27 決算說明資料（2026-10-29）",
 "Q1": "https://ir.nri.com/en/ir/library/financial/main/011111115/teaserItems2/02/linkList/02/link/20260730_financialstatements_e.pdf｜NRI 2026年4～6月 決算短信（2026-07-30）",
 "Q3": "https://ir.nri.com/en/ir/library/financial/main/011111114/teaserItems2/04/linkList/02/link/20260129_financialstatements_e.pdf｜NRI 2026年4～12月 決算短信（2026-01-29）",
 "H1_26": "https://ir.nri.com/en/ir/library/financial/main/011111114/teaserItems2/03/linkList/02/link/20251030_financialstatements_e.pdf｜NRI 2025年4～9月 決算短信（2025-10-30）",
 "P26": "https://ir.nri.com/en/ir/library/financial/main/011111114/teaserItems2/05/linkList/0/link/20260424_financialresultspresentation_e.pdf｜NRI FY3/26 決算說明資料（2026-04-24）",
 "SA": "https://stockanalysis.com/quote/tyo/4307/statistics/｜StockAnalysis NRI 統計頁（資料日 2026-11-09）",
 "SAF": "https://stockanalysis.com/quote/tyo/4307/forecast/｜StockAnalysis NRI 預測頁（S&P Global；更新 2026-09-11）",
 "KB": "https://kabuyoho.jp/sp/reportAnalyst?bcode=4307｜株予報 NRI 業績予想（2026-07-24 共識）",
 "KT": "https://kabuyoho.jp/sp/reportTarget?bcode=4307｜株予報 NRI 目標株価（2026-09-10）",
 "YH": "https://finance.yahoo.co.jp/quote/4307.T/history｜Yahoo!ファイナンス 4307 時系列（最新日 2026-10-20）",
 "MOF": "https://www.mof.go.jp/english/policy/jgbs/auction/calendar/eresul/eresul20261105a.htm｜財務省 10年債入札結果（2026-11-05）",
}
CHK = "同一數字互證：FY2025 營收 764,813／營業利益 134,907 見 FS26 與 FS25；FY2024 營收 736,556 見 FS24 與 FS25；FY2023 營收 692,165 見 FS23 與 FS24；FY2026 營收 814,708 見 FS26 與 H1；H1 FY27 營收 429,906 見 H1 與 H1P"
YRS = ["FY2022", "FY2023", "FY2024", "FY2025", "FY2026"]
# 年度合併（FY2022＝2022-03 期；4 月制）：科目 → [FY2022..FY2026]
ANN = {
 "rev":   ([611634, 692165, 736556, 764813, 814708], "FS23,FS23,FS24,FS25,FS26", "營收（売上収益）"),
 "cogs":  ([395562, 452336, 475549, 489517, 514556], "FS23,FS23,FS24,FS25,FS26", "營業成本（売上原価）"),
 "gp":    ([216071, 239829, 261006, 275295, 300151], "FS23,FS23,FS24,FS25,FS26", "毛利"),
 "sga":   ([113536, 131580, 142353, 144071, 145642], "FS23,FS23,FS24,FS25,FS26", "販管費"),
 "oi":    ([106218, 111832, 120411, 134907, 58273], "FS23,FS23,FS24,FS25,FS26", "營業利益（FY2026 含減損 97,586）"),
 "pbt":   ([104671, 108499, 117224, 134161, 58851], "FS23,FS23,FS24,FS25,FS26", "稅前利益"),
 "tax":   ([32878, 32002, 37062, 39778, 43067], "FS23,FS23,FS24,FS25,FS26", "所得稅"),
 "ni":    ([71445, 76307, 79643, 93762, 15257], "FS23,FS23,FS24,FS25,FS26", "歸屬母公司淨利"),
 "eps":   ([120.57, 128.92, 136.90, 163.57, 26.62], "FS23,FS23,FS24,FS25,FS26", "基本 EPS（日圓）"),
 "ocf":   ([98137, 118899, 142277, 130196, 147641], "FS23,FS23,FS24,FS25,FS26", "營業活動現金流"),
 "cx_p":  ([6824, 12061, 7484, 6660, 5844], "FS23,FS23,FS24,FS25,FS26", "資本支出：有形資產取得（現金）"),
 "cx_i":  ([36153, 42640, 42313, 42942, 40467], "FS23,FS23,FS24,FS25,FS26", "資本支出：無形資產取得（現金）"),
 "da":    ([41941, 45092, 48157, 48910, 51275], "FS23,FS23,FS24,FS25,FS26", "折舊攤銷（現金流量表；含使用權資產）"),
 "div":   ([22647, 25372, 27583, 33313, 39516], "FS23,FS23,FS24,FS25,FS26", "股利支付"),
 "bb":    ([60136, 20156, 68427, 30306, 0], "FS23,FS23,FS24,FS25,FS26", "庫藏股取得"),
 "cash":  ([115610, 129257, 173935, 168597, 132620], "FS23,FS23,FS24,FS25,FS26", "現金及約當現金"),
 "debt":  ([209627, 205823, 268104, 246277, 204962], "FS23,FS23,FS24,FS25,FS26", "有利子負債"),
 "eqo":   ([339360, 399006, 399532, 434040, 433905], "FS23,FS23,FS24,FS25,FS26", "歸屬母公司權益"),
 "dps":   ([40, 45, 53, 63, 77], "FS23,FS23,FS24,FS25,FS26", "每股股利（日圓）"),
 "sh_iss": ([610.696500, 593.652242, 580.796911, 581.241511, 581.241511], "FS23,FS23,FS24,FS25,FS26", "期末發行股數（百萬股，含庫藏）"),
 "sh_tre": ([21.109262, 1.959169, 4.545600, 9.133722, 6.837002], "FS23,FS23,FS24,FS25,FS26", "期末庫藏股（百萬股；含員工持股信託）"),
}
# 分部（外部客戶營收；FY2022–FY2025 為當時揭露的舊口徑，FY2026 同時給舊口徑與 2026-04 起新口徑〔澳洲由 Industrial 移到 Financial〕）
SEGX = {  # 外部營收
 "old": {"Consulting": [42807, 46100, 51959, 60333, 63556], "Financial IT Solutions": [303635, 328576, 349872, 366605, 399792],
         "Industrial IT Solutions": [222583, 266723, 275923, 267487, 271241], "IT Platform Services": [42607, 48153, 55628, 67041, 77039],
         "Others／Adjustments": [0, 2612, 3172, 3346, 3077]},
 "new26": {"Consulting": 63556, "Financial IT Solutions": 455522, "Industrial IT Solutions": 215511, "IT Platform Services": 77039, "Others／Adjustments": 3077},
}
SEGO = {  # 分部營業利益（含內部交易口徑；FY2026 舊口徑含 Industrial 減損 97,060）
 "Consulting": [12820, 12329, 13929, 18398, 19225], "Financial IT Solutions": [43877, 49710, 54651, 61581, 74255],
 "Industrial IT Solutions": [25449, 24393, 23405, 24287, -74616], "IT Platform Services": [20955, 23046, 28167, 30341, 38564],
 "Others／Adjustments": [3116, 2352, 257, 298, 844]}
SUBL26 = {"Securities": 155369, "Insurance": 89366, "Banking": 84413, "Other financial": 126373, "Distribution": 62186, "Manufacturing, service & other": 153325}
# 半期（新口徑）
H1 = {  # [1H FY2026（新口徑重編）, 1H FY2027]
 "rev": [397065, 429906], "cogs": [249047, 268510], "gp": [148017, 161396], "sga": [69845, 74383], "oi": [78799, 88272], "bp": [79213, 88281],
 "fin_i": [3085, 3415], "fin_c": [2428, 1983], "pbt": [79456, 89703], "tax": [25640, 29446], "pl": [53816, 60257], "ni": [53565, 60014], "eps": [93.53, 106.61],
 "da": [24798, 26109], "ocf": [73668, 94534], "cx_p": [2271, 4635], "cx_i": [22789, 16408], "sh_avg": [572.737291, 562.930511],
}
H1SEG_EXT = {"Consulting": [29321, 30516], "Financial IT Solutions": [222386, 241660], "Industrial IT Solutions": [108197, 114312], "IT Platform Services": [35685, 41879], "Others／Adjustments": [1474, 1536]}
H1SEG_TOT = {"Consulting": [31706, 33192], "Financial IT Solutions": [224857, 244539], "Industrial IT Solutions": [112518, 118310], "IT Platform Services": [106619, 119198]}
H1SEG_OI = {"Consulting": [8348, 7784], "Financial IT Solutions": [35939, 42317], "Industrial IT Solutions": [14853, 14924], "IT Platform Services": [19279, 22712], "Others／Adjustments": [377, 533]}
H1SUB27 = {"Securities": 79477, "Insurance": 45091, "Banking": 46470, "Other financial": 70621, "Distribution": 36142, "Manufacturing, service & other": 78170}
Q1SEG_EXT = {"Consulting": 13838, "Financial IT Solutions": 119092, "Industrial IT Solutions": 56873, "IT Platform Services": 19874, "Others／Adjustments": 776}
Q1SEG_OI = {"Consulting": 2901, "Financial IT Solutions": 20742, "Industrial IT Solutions": 6711, "IT Platform Services": 11096, "Others／Adjustments": 259}
Q1 = {"rev": 210455, "oi": 41711, "ni": 29214, "eps": 51.42}
M9_26 = {"rev": 602333, "oi": 118780, "ni": 83201, "eps": 145.21}      # 2025 年 4～12 月
H1_26OLD_EXT = {"Consulting": 29321, "Financial IT Solutions": 193547, "Industrial IT Solutions": 137036, "IT Platform Services": 35685, "Others／Adjustments": 1474}
M9_26OLD_EXT = {"Consulting": 46475, "Financial IT Solutions": 295450, "Industrial IT Solutions": 203246, "IT Platform Services": 54901, "Others／Adjustments": 2259}
BS27 = {"cash": 142837, "ofa": 14040, "debt": 218810, "mi": 3916, "eqo": 383746, "sh_iss": 561.197409, "sh_tre": 5.811620, "lease_mar26": 31736, "ofa_mar26": 61975}
OTHER = {"emp": 16894, "sub_jp": 9500, "sub_off": 7000, "turnover": 0.036, "ai_inv_fy26": 18100, "backlog_mar26": 470600, "backlog_sep26": 341087, "orders_h1": 294549,
         "fx_usd_h1": 146.03, "fx_aud_h1": 94.50, "capex_guide27": 49500, "da_guide27": 40600, "guid_rev27": 870000, "guid_oi27": 180000, "guid_ni27": 123000, "guid_eps27": 218.50, "guid_dps27": 88,
         "px": 6056.0, "px_val": 4952.0, "rf": 0.02768, "beta_raw": 0.34, "tp_cons_avg": 5978, "tp_cons_n": 14, "tp_cons_hi": 6700, "tp_cons_lo": 5190, "rev27_cons": 867500, "rev27_cons_hi": 884600, "rev27_cons_lo": 855800,
         "rev27_kb": 858121, "oi27_kb": 178992, "ni27_kb": 122488}
PEERS = [  # 代號, 名稱, 前瞻 P/E, EV/EBITDA, 資料日, 會計準則
 ("3626.T", "TIS", 15.42, 9.20, "2026-11-16", "IFRS"), ("8056.T", "BIPROGY", 13.86, 7.90, "2026-10-05", "IFRS"), ("6702.T", "Fujitsu", 23.96, 13.03, "2026-10-23", "IFRS"),
 ("2327.T", "NS Solutions", 20.88, 10.80, "2026-09-15", "J-GAAP"), ("4684.T", "OBIC", 24.17, 18.71, "2026-09-11", "J-GAAP"),
 ("ACN", "Accenture", 17.61, 11.41, "2026-10-30", "US GAAP"), ("CTSH", "Cognizant", 11.11, 8.00, "2026-10-27", "US GAAP"), ("INFY", "Infosys", 14.43, 10.30, "2026-10-20", "IFRS")]

def rows():
    out = []
    def a(metric, v, unit, per, basis, src, tag, inc="", page=""):
        s = SRC[src] + (f"（{page}）" if page else "")
        out.append([metric, v, unit, per, basis, s, EXT, tag, inc])
    for k, (vals, srcs, nm) in ANN.items():
        unit = "日圓" if k in ("eps", "dps") else ("百萬股" if k.startswith("sh_") else "百萬日圓")
        for y, v, s in zip(YRS, vals, srcs.split(",")):
            a(nm, v, unit, f"{y}（{int(y[2:])-1}-04～{y[2:]}-03）", "IFRS 合併", s, "Verified")
    for nm, vals in SEGX["old"].items():
        for y, v, s in zip(YRS, vals, ("FS23", "FS24", "FS24", "FS26", "FS26")):
            a(f"分部外部營收：{nm}", v, "百萬日圓", f"{y}（舊口徑）", "分部揭露；FY2025 為 FY3/26 決算短信重編", s, "Verified")
    for nm, v in SEGX["new26"].items():
        a(f"分部外部營收：{nm}", v, "百萬日圓", "FY2026（2025-04～2026-03；2026-04 新口徑重編）", "澳洲 NRI Australia 由 Industrial IT 移至 Financial IT", "H1", "Verified", page="p.9 預測表之前期實績欄")
    for nm, vals in SEGO.items():
        for y, v, s in zip(YRS, vals, ("FS23", "FS24", "FS25", "FS26", "FS26")):
            a(f"分部營業利益：{nm}", v, "百萬日圓", f"{y}（舊口徑）", "含內部交易口徑；FY2026 Industrial 含減損 97,060", s, "Verified")
    for nm, v in SUBL26.items():
        a(f"產品線外部營收：{nm}", v, "百萬日圓", "FY2026（新口徑）", "公司固定揭露的分部下產業別", "H1", "Verified", page="p.9")
    for k, (v26, v27) in H1.items():
        unit = "日圓" if k == "eps" else ("百萬股" if k == "sh_avg" else "百萬日圓")
        a(f"半期合併 {k}", v26, unit, "1H FY2026（2025-04～09）", "IFRS", "H1", "Verified", page="p.1,3,13")
        a(f"半期合併 {k}", v27, unit, "1H FY2027（2026-04～09）", "IFRS", "H1", "Verified", page="p.1,3,13")
    for nm, (v26, v27) in H1SEG_EXT.items():
        a(f"半期分部外部營收：{nm}", v26, "百萬日圓", "1H FY2026（新口徑）", "", "H1", "Verified", page="p.5-6,19"); a(f"半期分部外部營收：{nm}", v27, "百萬日圓", "1H FY2027", "", "H1", "Verified", page="p.5-6,19")
    for nm, (v26, v27) in H1SEG_TOT.items():
        a(f"半期分部總營收（含內部）：{nm}", v26, "百萬日圓", "1H FY2026（新口徑）", "", "H1", "Verified", page="p.5-6,19"); a(f"半期分部總營收（含內部）：{nm}", v27, "百萬日圓", "1H FY2027", "", "H1", "Verified", page="p.5-6,19")
    for nm, (v26, v27) in H1SEG_OI.items():
        a(f"半期分部營業利益：{nm}", v26, "百萬日圓", "1H FY2026（新口徑）", "", "H1", "Verified", page="p.5-6,19"); a(f"半期分部營業利益：{nm}", v27, "百萬日圓", "1H FY2027", "", "H1", "Verified", page="p.5-6,19")
    for nm, v in H1SUB27.items():
        a(f"半期產品線外部營收：{nm}", v, "百萬日圓", "1H FY2027", "", "H1P", "Verified", page="p.6-7")
    for nm, v in Q1SEG_EXT.items():
        a(f"季分部外部營收：{nm}", v, "百萬日圓", "1Q FY2027（2026-04～06）", "新口徑", "Q1", "Verified")
    for nm, v in Q1SEG_OI.items():
        a(f"季分部營業利益：{nm}", v, "百萬日圓", "1Q FY2027", "新口徑", "Q1", "Verified")
    for k, v in Q1.items():
        a(f"季合併 {k}", v, "日圓" if k == "eps" else "百萬日圓", "1Q FY2027", "IFRS", "Q1", "Verified")
    for k, v in M9_26.items():
        a(f"9M 合併 {k}", v, "日圓" if k == "eps" else "百萬日圓", "9M FY2026（2025-04～12）", "IFRS", "Q3", "Verified", page="p.10")
    for nm, v in M9_26OLD_EXT.items():
        a(f"9M 分部外部營收：{nm}", v, "百萬日圓", "9M FY2026（舊口徑）", "", "Q3", "Verified", page="p.16")
    for nm, v in H1_26OLD_EXT.items():
        a(f"半期分部外部營收：{nm}", v, "百萬日圓", "1H FY2026（舊口徑）", "", "H1_26", "Verified")
    for k, v in BS27.items():
        a(f"資產負債表／股數 {k}", v, "百萬股" if k.startswith("sh_") else "百萬日圓", "2026-09-30" if "mar26" not in k else "2026-03-31", "IFRS", "H1" if "mar26" not in k else "FS26", "Verified", page="p.7,11-12")
    a("員工數（NRI 集團）", 16894, "人", "2026-03-31", "公司 IR 資料", "H1P", "Interested-party", "公司自揭 KPI：呈現規模與人力投入的誘因，口徑由公司定義", "p.44")
    a("外包人力（日本約 9,500、海外約 7,000；欄位標示含糊）", 16500, "人", "2026-03-31", "公司 IR 資料；兩數相加為 Derived", "H1P", "Interested-party", "公司自揭 KPI：口徑由公司定義，含糊處未核對", "p.44")
    a("AI 相關投資（FY3/26）", 18100, "百萬日圓", "FY2026", "公司 IR 資料", "P26", "Interested-party", "公司自揭：呈現 AI 投入的誘因；口徑未定義", "p.24")
    a("訂單餘額（2026-03）", 470600, "百萬日圓", "2026-03-31", "公司 IR 資料", "P26", "Interested-party", "公司自揭 KPI", "p.17-18")
    a("訂單餘額（2026-09）", 341087, "百萬日圓", "2026-09-30", "公司 IR 資料", "H1P", "Interested-party", "公司自揭 KPI", "p.10,23-24")
    a("上半期訂單（2026-04～09）", 294549, "百萬日圓", "1H FY2027", "公司 IR 資料", "H1P", "Interested-party", "公司自揭 KPI", "p.25-26")
    a("美元兌日圓平均匯率（1H FY2027）", 146.03, "日圓／美元", "1H FY2027", "公司 IR 資料", "H1P", "Verified", "", "p.22")
    a("FY2027 營收指引", 870000, "百萬日圓", "FY2027（2026-04～2027-03）", "公司指引（2026-10-29 上修；4 月為 850,000）", "H1", "Interested-party", "管理層指引：維持市場信心的誘因，未經查核", "p.9")
    a("FY2027 營業利益指引", 180000, "百萬日圓", "FY2027", "公司指引（4 月為 175,000）", "H1", "Interested-party", "管理層指引", "p.9")
    a("FY2027 歸屬母公司淨利指引", 123000, "百萬日圓", "FY2027", "公司指引", "H1", "Interested-party", "管理層指引", "p.9")
    a("FY2027 EPS 指引", 218.50, "日圓", "FY2027", "公司指引", "H1", "Interested-party", "管理層指引", "p.9")
    a("FY2027 DPS 指引（中間 42＋期末 46）", 88, "日圓", "FY2027", "公司指引", "H1", "Interested-party", "管理層指引", "p.1,9-10")
    a("FY2027 資本支出指引（不含使用權資產）", 49500, "百萬日圓", "FY2027", "公司指引", "H1P", "Interested-party", "管理層指引", "p.41")
    a("FY2027 折舊攤銷指引（不含使用權資產）", 40600, "百萬日圓", "FY2027", "公司指引", "H1P", "Interested-party", "管理層指引", "p.41")
    a("現價（2026-11-09 收盤）", 6056, "日圓", "2026-11-09", "TSE 收盤（延遲報價）；單一來源，第二來源未找到", "SA", "Verified")
    a("評價日收盤（2026-09-30）", 4952, "日圓", "2026-09-30", "TSE 收盤", "YH", "Verified")
    a("10 年期日本公債殖利率（10 年債入札平均殖利率）", 0.02768, "比率", "2026-11-05", "財務省入札結果", "MOF", "Verified")
    a("Beta（5Y，資料商）", 0.34, "倍", "2026-11-09", "StockAnalysis；對照頁 2026-09-09 為 0.32", "SA", "Verified")
    a("賣方目標價平均（14 家）", 5978, "日圓", "更新 2026-09-11（財報前）", "S&P Global 經 StockAnalysis；高 6,700、低 5,190", "SAF", "Interested-party", "賣方分析師：評等與交易業務誘因，目標價傾向跟隨股價；財報前只作歷史對照")
    a("賣方 FY2027 營收共識平均", 867500, "百萬日圓", "更新 2026-09-11（財報前）", "高 884,600、低 855,800", "SAF", "Interested-party", "賣方分析師：財報前只作歷史對照")
    a("賣方 FY2027 共識（營收 858,121／營業利益 178,992／淨利 122,488）", 178992, "百萬日圓", "2026-07-24（財報前）", "營業利益；營收 858,121、淨利 122,488", "KB", "Interested-party", "賣方分析師：財報前只作歷史對照")
    a("賣方目標價平均（13 家；評等 強買 11／中立 1／弱 1）", 5954, "日圓", "2026-09-10（財報前）", "", "KT", "Interested-party", "賣方分析師：評等與交易業務誘因")
    inc_co = "公司管理層：展示 AI 推進與成長的誘因，口徑由公司定義，未經查核"
    a("AI 開發工具適用比例（2026 年）", 0.25, "比例", "2026（公司敘述）", "公司稱 AI 驅動開發適用約 25%，目標 FY2028 全面導入（75% 為公司圖表讀值，摘要擷取非逐字）", "H1P", "Interested-party", inc_co, "p.11")
    a("AX 顧問營收年增（1H27）", 2.0, "倍", "1H FY2027", "公司稱約倍增（roughly doubled；數量未揭露）", "H1P", "Interested-party", inc_co, "p.13")
    a("離職率", 0.036, "比例", "FY2026", "公司人才資料", "H1P", "Interested-party", inc_co, "p.44")
    a("token 量、代理任務量、每任務 token", "未找到", "—", "—", "公司未揭露", "H1P", "Derived", "", "")
    for t, n, pe, ev, d, g in PEERS:
        out.append(["同業前瞻 P/E：" + n + "（" + t + "）", pe, "倍", d, f"資料商前瞻 P/E（{g}）；與 NRI 估值日不同（逾期標示）", "https://stockanalysis.com/（各公司 statistics 頁）｜StockAnalysis 統計頁", EXT, "Interested-party", "資料商前瞻倍數含賣方預估：估值日與口徑未與 NRI 對齊"])
        out.append(["同業 EV/EBITDA：" + n + "（" + t + "）", ev, "倍", d, f"資料商（{g}）；TTM 口徑，不用於估值", "https://stockanalysis.com/（各公司 statistics 頁）｜StockAnalysis 統計頁", EXT, "Interested-party", "資料商計算，口徑未核對"])
    return out

if __name__ == "__main__":
    os.makedirs(os.path.join(HERE, "data"), exist_ok=True)
    R = rows()
    with open(os.path.join(HERE, "data", "facts_nri.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f); w.writerow(["指標", "數值", "單位", "期間", "口徑", "出處", "擷取日", "標記", "來源誘因"]); w.writerows(R)
    print(f"facts_nri.csv：{len(R)} 列")

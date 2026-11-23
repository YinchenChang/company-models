# Symbotic（SYM，Nasdaq）個股模型 v0.1 — AI實體應用 Project（線 5 倉儲物流自動化）
# 用法：python3 physical_ai/symbotic/build_symbotic.py            → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、segments.json、sens.json
#       python3 physical_ai/symbotic/build_symbotic.py --raw OUT  → 只寫一本（讀環境變數 SYM_OVERRIDE 的 JSON 覆寫 In 頁基準值）
# 會計年度：52／53 週制，9 月最後一個週六前後結束（FY2025 止於 2025-09-27）。金額：十億美元（US$B）；股數：百萬股；EPS、股價：美元。
# 結構：README、In、Hist、Q（季度營收線與分部對帳）、PL（三情境：在手訂單轉換＋新訂單 → 營收線 → 標準損益 → 現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 公司只有一個報告分部（單一營運分部）；營收依 Systems／Software maintenance and support／Operation services 三條營收線揭露，本模型以三條營收線當分部列。
# DCF 為本檔自寫（未用 tools/fcff_dcf.py），公式同 docs/plan/個股DCF標準規格.md。實體 AI 線不讀 Siliconomics／Tokenomics／Wattonomics。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
VER = "v0.1"
DATE = "20261123"
NAME = f"{DATE}_Symbotic個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032

# ---------------- 出處 ----------------
Q3K = "SEC 8-K 附件 99.1：Symbotic Reports Third Quarter Fiscal Year 2026 Results（2026-08-05）https://www.sec.gov/Archives/edgar/data/0001837240/000183724026000042/q3268-k_ex991.htm"
Q2K = "SEC 8-K 附件 99.1：第 2 季 FY2026 財報（季度截至 2026-03-28）https://www.sec.gov/Archives/edgar/data/0001837240/000183724026000023/q2268-k_ex991.htm"
Q10 = "SEC 10-Q（季度截至 2026-06-27；封面股數 2026-08-03）https://www.sec.gov/Archives/edgar/data/1837240/000183724026000043/sym-20260627.htm"
FY25 = "公司新聞稿 Symbotic Reports Fourth Quarter and Fiscal Year 2025 Results（2025-11-24；8-K 附件同文，經 finviz 轉載）https://finviz.com/news/237861/symbotic-reports-fourth-quarter-and-fiscal-year-2025-results"
FY24 = "SEC 8-K 附件 99.1：第 4 季與 FY2024 財報（FY2023 比較欄）https://www.sec.gov/Archives/edgar/data/1837240/000183724024000213/q4248-k_ex9911.htm"
CALL = "MarketBeat 經 Nasdaq 轉載：Symbotic Q3 earnings call highlights（2026-08-06）https://www.nasdaq.com/articles/symbotic-q3-earnings-call-highlights"
PAN = "Panabee（AI 整理 FY2025 10-K，2025-11-24；資料商，非原文）https://www.panabee.com/news/symbotic-earnings-2025-annual"
SAF = "stockanalysis.com SYM forecast 頁（2026-10-22；S&P Global 彙整 17 位分析師）https://stockanalysis.com/stocks/sym/forecast/"
SAS = "stockanalysis.com 各公司 statistics 頁（讀取 2026-11-23；頁面報價日見各列）"
YCL = "Yahoo Finance 日線（physical_ai/data/closes2_20261109_20261120.csv）"
UST = "美國 10 年期公債殖利率 5.10%（FRED DGS10 2026-11-12；與 software_ai/hubspot 同一取值，跨模型一致）"
WHL = "physical_ai/lines/warehouse/dist/20261116_倉儲物流單位經濟_v0.1.xlsx「下游取數」頁"
PROF = "physical_ai/map/profiles/SYM.md（轉述 DEF 14A 2026-01-16 受益所有權表；未回原文逐字核對）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"

# ---------------- In 列 ----------------
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. FY2026 錨點：前三季實績（Hist 頁）＋第 4 季公司指引（Interested-party）")
inp("g4_rev", "公司指引：第 4 季 FY2026 營收（7.6–7.8 億美元中點）", "US$B", None, 0.770, None, "Interested-party", Q3K + "：Q4 營收指引 760–780 百萬美元", "FY2026 隱含營收＝前三季 2.027 ＋ 0.770 ≈ 2.80（共識 2.80，" + SAF + "）", CO)
inp("g4_ebitda", "公司指引：第 4 季調整後 EBITDA（1.00–1.05 億美元中點；只作對照，不進模型）", "US$B", None, 0.1025, None, "Interested-party", Q3K + "：Q4 調整後 EBITDA 指引 100–105 百萬美元", "Recon 對照模型推得值，只作對照，不進模型", CO)
inp("g_anc", "第 4 季營收實績相對公司指引", "%", -0.02, 0.0, 0.02, "Assumed", "基準＝指引中點；第 4 季財報約 2026-11 下旬公布（本版截止前未公布）", "低／高 ±2%（指引區間 ±1.3%）")
inp("g4_gm", "第 4 季毛利率（管理層：接近第 2 季水準）", "%", 0.205, "=h_gp_q2/h_rev_q2", 0.225, "Interested-party", CALL + "：Q4 毛利率接近 Q2 水準", "Q2 毛利 149.952 ÷ 營收 676.480 ＝ 22.2%（" + Q2K + "）；低／高 Assumed", CO)
inp("g4_opx", "第 4 季營業費用相對第 3 季增幅（管理層：小幅增加）", "%", 0.0, 0.03, 0.08, "Assumed", CALL + "：Q4 營業費用、研發小幅增加（定性）", "Q3 營業費用 128.015 百萬美元")
inp("g4_oi", "第 4 季業外收入淨額", "US$B", 0.01, 0.015, 0.03, "Assumed", "現金約 17.5 億 × 約 3.5% ÷ 4", "Q3 30.587 百萬美元含非經常項目（未拆）")
inp("g4_eqm", "第 4 季權益法投資損益（Exol，原 GreenBox 合資）", "US$B", None, -0.010, None, "Assumed", "Q3 −9.631、Q2 −6.945 百萬美元（" + Q3K + "）", "")
inp("q4_fcf", "第 4 季自由現金流（評價日前 stub）", "US$B", 0.0, 0.10, 0.30, "Interested-party", CALL + "：管理層預期 Q4 與全年自由現金流為正", "前三季 242.8 百萬美元；Q3 −164.6（專案時點）；金額 Assumed", CO)
inp("ops_26", "運作中系統數（FY2026 年末）", "套", 58, 60, 62, "Assumed", Q10 + "：Q3 運作中 56 套（一年前 42）、部署中 77 套", "Q4 新增約 4 套（Q3 新增 4 套），Assumed")

sec("B. 命題模組：在手訂單（RPO）轉營收 ＋ 新訂單 ＋ Walmart 門市後場（Systems 營收線）")
inp("rpo0", "剩餘履約義務（RPO，即 backlog；2026-06-27）", "US$B", None, 22.5, None, "Verified", Q10 + "：RPO 約 225 億美元，Walmart 與 Exol 占絕大部分", "FY2025 年末亦約 225 億（" + PAN + "）→ 九個月新簽約 ≈ 認列營收（book-to-bill 約 1，Derived）")
inp("r12", "RPO 中未來 12 個月認列比例（2026-07～2027-06）", "%", 0.12, 0.15, 0.18, "Verified", Q10 + "：約 15% 於未來 12 個月認列", "FY2025 10-K 為約 12%（" + PAN + "）；低＝12%（前期值）、高＝18%（Assumed）")
inp("sys_sh", "RPO 中 Systems 占比（其餘為 15 年期軟體維護與營運服務）", "%", 0.80, 0.88, 0.95, "Assumed", "未揭露：公司不拆分 RPO", "軟體維護合約一般 15 年、預收（" + Q10 + "）；133 套 × 每套年約 0.9 百萬 × 15 年 ≈ 18 億 → Systems 約 85–92%（Derived 粗估）")
inp("n_rem", "其餘 RPO（未來 12 個月以後）認列年數（自 2027-07 起平均）", "年", 5, 7, 10, "Assumed", "未揭露：公司未給 12 個月以後的認列時程", "Walmart 42 座區域配送中心合約（2022）與 Exol 系統分批部署；部署中 77 套、每季開工約 11 套（" + CALL + "）")
inp("bk_cut", "尾端：既有 Systems backlog 取消或無限期遞延比例", "%", None, 0.0, 0.2, "Assumed", "基準 0；高＝20%（單一客戶占營收一成以上，第 10 節第 19 條尾端 Sens）", "RPO 價格調整已使 Q3 backlog 略降（" + CALL + "）")
for s, v, lo, hi, tag, note in (
        ("A", 3.5, None, None, "Assumed", "樂觀：新客戶（Southern Glazer's 後續場址、C&S、Exol 多客戶場址）與新一代儲存結構帶動簽約加速"),
        ("B", 2.5, 1.5, 3.5, "Analogy", "類比 FY2026 前三季：RPO 由約 225 億維持 225 億，新簽約 ≈ 認列營收 2.03 → 年化約 2.7、Systems 約 94%"),
        ("C", 1.5, None, None, "Assumed", "保守：Walmart 以外客戶簽約緩慢，Walmart 新單放緩")):
    inp(f"o27_{s}", f"情境 {s}：新簽 Systems 訂單 FY2027（不含 Walmart 門市後場）", "US$B", lo, v, hi, tag, note, "區間見情境 A／B／C 三列" if s != "B" else "類比對象：前三季簽約速度（口徑含價格調整與 Exol）")
for s, ref, fml, lo, hi in (("A", "高／樂觀欄 E8、E11", "=(19.60365/10.92945)^(1/6)-1", None, None),
                            ("B", "基準欄 D8、D11", "=(16.94452/10.69434)^(1/6)-1", 0.053, 0.102),
                            ("C", "低／悲觀欄 C8、C11", "=(14.27712/10.45241)^(1/6)-1", None, None)):
    inp(f"go_{s}", f"情境 {s}：新簽訂單年增 FY2028 起（讀線模型 AS/RS 市場 2026→2032 年複合）", "%/年", lo, fml, hi, "Derived",
        f"{WHL} WH_ASRS_MKT_2026／WH_ASRS_MKT_2032（{ref}）", "MarketsandMarkets 口徑（含製造、汽車），與 Symbotic 整箱棧板系統口徑不同；只取成長率")
for s, v, lo, hi in (("A", 5.0, None, None), ("B", 2.5, 0.0, 5.0), ("C", 0.0, None, None)):
    inp(f"apd_{s}", f"情境 {s}：Walmart 門市後場系統（400 套）簽約金額，FY2028 簽約", "US$B", lo, v, hi, "Assumed",
        Q10 + "：2025 年 Master Automation Agreement 的或有承諾，可使 RPO 增加逾 50 億美元（不在 225 億內）", CALL + "：Walmart 依第二個門市場址結果決定擴大；最早 2028 年初才影響業績（Interested-party）")
inp("apd_L", "門市後場系統認列年數（FY2029 起平均）", "年", 4, 5, 8, "Assumed", "400 套門市系統分批建置", "")
inp("L_new", "新簽 Systems 訂單認列年數（簽約次年起平均）", "年", 3, 4, 6, "Assumed", "配送中心系統自開工到上線約 2–3 年；新訂單另有排程等待", "")
inp("wmt_no", "新簽訂單中 Walmart 占比（不含門市後場）", "%", 0.4, 0.6, 0.8, "Assumed", Q10 + "：Customer A 占營收 9M FY2026 87.0%、Q3 90.5%；占應收 82.8%", "新訂單較歷史分散（Exol、Southern Glazer's、C&S）")
inp("wmt_tail", "尾端：Walmart 停止新簽約（含門市後場）", "x", None, 0.0, 1.0, "Assumed", "基準 0；高＝1（Walmart 新訂單與門市後場全數不發生；既有 backlog 照常認列）", "第 10 節第 19 條：單一客戶占營收一成以上建尾端 Sens")
for s, v, lo, hi in (("A", 25, None, None), ("B", 18, 12, 25), ("C", 12, None, None)):
    inp(f"comp_{s}", f"情境 {s}：每年新增運作中系統數 FY2027 起", "套/年", lo, v, hi, "Analogy" if s == "B" else "Assumed",
        Q10 + "：運作中 42 → 56 套（一年 +14）；部署中 46 → 77 套", "B 以部署中 77 套 ÷ 約 4 年平均建置期推估" if s == "B" else "區間見情境 A–C")
inp("sw_rate", "軟體維護收入：每套運作中系統每年", "US$B/套", 0.0007294, "=h_sw_q3*4/56", 0.0010941, "Derived", "＝Q3 軟體維護營收 12.765 百萬 × 4 ÷ 56 套（" + Q3K + "、" + Q10 + "）", "低／高 ±20%（Assumed）；合約一般 15 年、預收")
inp("svc_r", "營運服務營收 ÷ Systems 營收", "%", 0.04, "=h_svc_9m/h_sys_9m", 0.06, "Derived", "＝9M FY2026 94.989 ÷ 1,895.740 ＝ 5.0%（" + Q3K + "）", "FY2025 4.6%；低／高 Assumed")
for s, v, lo, hi in (("A", 0.27, None, None), ("B", 0.24, 0.21, 0.27), ("C", 0.20, None, None)):
    inp(f"gms_{s}", f"情境 {s}：Systems 毛利率 FY2029 目標（FY2026 起線性收斂）", "%", lo, v, hi, "Assumed",
        "Systems 毛利率實績：FY2024 14.0%、FY2025 19.3%、9M FY2026 21.5%、Q3 22.0%（Derived，" + Q3K + "）", CALL + "：新一代儲存結構在下一會計年度下半起成為較顯著的毛利驅動（Interested-party）；關稅與保固為抵銷項")
inp("gm_sw", "軟體維護毛利率", "%", 0.65, "=1-h_csw_9m/h_sw_9m", 0.80, "Derived", "＝1 − 9.808 ÷ 36.574 ＝ 73.2%（9M FY2026）", "FY2025 71.1%")
inp("gm_svc", "營運服務毛利率", "%", 0.05, "=1-h_csvc_9m/h_svc_9m", 0.15, "Derived", "＝1 − 84.178 ÷ 94.989 ＝ 11.4%（9M FY2026）", "FY2025 −7.0%")
inp("g_opx", "營業費用（研發＋銷管＋重組）年增 FY2027 起", "%/年", 0.02, 0.06, 0.12, "Assumed", "FY2025 537.6 → FY2026 約 531（前三季 398.8 ＋ Q4 估）", CALL + "：研發與營業費用小幅增加；含股權報酬（9M 142.9 百萬）")

sec("C. 損益表下半段（經濟口徑：V 類股視同轉換）")
inp("oi", "業外收入淨額 FY2027 起（利息收入為主）", "US$B", 0.03, 0.06, 0.10, "Assumed", "9M FY2026 54.688 百萬（Q3 30.587 含非經常項目）；FY2025 36.063", "無借款")
inp("eqm", "權益法投資損益（Exol）FY2027 起", "US$B", -0.06, -0.03, 0.0, "Assumed", "9M FY2026 −22.375、FY2025 −13.716 百萬（" + Q3K + "）", "Exol 亞特蘭大場址已上線、加州 Lathrop 場址 60–90 天內上線（" + CALL + "）")
inp("tax_w", "有效稅率（經濟口徑：全部 V 類視同轉換為 A 類，公司層級課稅）", "%", 0.15, 0.23, 0.26, "Assumed", "美國聯邦 21%＋州稅；Up-C 結構下 GAAP 稅負只計 A 類持有的 21.4%，加上虧損扣抵，目前 GAAP 有效稅率約 0", "低＝15%（虧損扣抵與股權報酬扣除延長）")
inp("nci_sh", "非控制權益占比（V 類股 ÷ 全部股數）", "%", None, "=(sh_v1+sh_v3)/(sh_a+sh_v1+sh_v3)", None, "Derived", "＝(71.369 ＋ 403.559) ÷ 604.802（" + Q10 + " 封面股數）", Q10 + "：2026-06-27 Symbotic Holdings 由 A 類持有 21.4%、V 類 78.6%")
inp("sbc_r", "股權報酬 ÷ 營收（只用於 Recon 調整後 EBITDA 對照；DCF 不加回）", "%", None, "=h_sbc_9m/h_rev_9m", None, "Derived", "＝9M FY2026 142.919 ÷ 2,027.303（" + Q10 + " 現金流量表）", "")
inp("pay_r", "配息率", "%", None, 0.0, None, "Verified", Q10 + "：未發放股利、無買回", "")

sec("D. 估值輸入（評價日＝FY2026 年末 2026-09-26；資產負債表 2026-06-27）")
inp("px", "股價（2026-11-20 收盤）", "US$", None, 31.1, None, "Verified", YCL + "：SYM 2026-11-20 收盤 31.10", "量級對照：2026-10-23 收盤 39.28（" + SAF + "）；2026-08 財報後單日約 −15%，11-13 前累計約 −25%（公司檔案）")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("px_oct", "對照：2026-10-23 收盤（只作對照，不進模型）", "US$", None, 39.28, None, "Interested-party", SAF, "量級核對：與 csv 收盤同一量級", DV)
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("val_date", "評價日（FY2026 年末；52 週制推算）", "日期", None, "=DATE(2026,9,26)", None, "Derived", "FY2025 止於 2025-09-27 ＋ 52 週", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, "=(tp_date-val_date)/365", None, "Derived", "", "")
inp("sh_a", "A 類普通股（公開交易；2026-08-03）", "百萬股", None, 129.873381, None, "Verified", Q10 + " 封面：129,873,381 股", "")
inp("sh_v1", "V-1 類普通股（每股 1 票；對應 Holdings 單位，可換 A 類；2026-08-03）", "百萬股", None, 71.369131, None, "Verified", Q10 + " 封面：71,369,131 股", "主要持有人 Walmart（持 V-1 約 84%，" + PROF + "）")
inp("sh_v3", "V-3 類普通股（每股 3 票；創辦人家族信託等；2026-08-03）", "百萬股", None, 403.559196, None, "Verified", Q10 + " 封面：403,559,196 股", "對應 Holdings 單位，可一比一換 A 類")
inp("sh_tot", "全部普通股（A＋V-1＋V-3）", "百萬股", None, "=sh_a+sh_v1+sh_v3", None, "Derived", "經濟權益：V 類持有 Holdings 單位，與 A 類同享每單位盈餘", "")
inp("rsu", "RSU 等稀釋股（以 Q3 稀釋與基本加權股數差代）", "百萬股", None, "=133.252947-128.076383", 15.0, "Derived", Q3K + "：Q3 A 類加權股數 稀釋 133,252,947、基本 128,076,383", "RSU 流通在外總數未找到；高＝15 百萬股（Assumed）")
inp("war_n", "Exol 認股權證股數（未歸屬，8 批）", "百萬股", None, 11.43436, None, "Verified", Q10 + "：至多 11,434,360 股 A 類、履約價 41.9719、2026-06-27 尚未歸屬任何一批、2027-07-23 到期（展延則 2029）", "Walmart 2022 年認股權證（Holdings 單位）現況未找到")
inp("war_k", "Exol 認股權證履約價", "US$", None, 41.9719, None, "Verified", Q10, "")
inp("war_tsm", "Exol 認股權證庫藏股法淨股數（現價低於履約價時為 0）", "百萬股", None, "=MAX(0,war_n*(1-war_k/px))", None, "Derived", "DCF 標準規格第 3 節；未歸屬仍以庫藏股法計入（現價下為 0）", "")
inp("sh_v", "評價日股數（全部普通股＋RSU＋認股權證淨股數）", "百萬股", None, "=sh_tot+rsu+war_tsm", None, "Derived", "DCF 標準規格第 3 節", "")
inp("cash_v", "現金及約當現金（2026-06-27）", "US$B", None, 1.746446, None, "Verified", Q3K + "：1,746,446 千美元", "含客戶預收款（遞延收入 17.37 億），預收款的回吐在 FCFF 營運資金列扣除")
inp("debt_v", "借款（2026-06-27）", "US$B", None, 0.0, None, "Verified", Q3K + "：資產負債表無借款科目", "租賃負債金額未找到（未計入）")
inp("inv_v", "權益法投資：Exol（2026-06-27 帳面）", "US$B", None, 0.140468, None, "Verified", Q3K + "：140,468 千美元", Q10 + "：最大損失曝險 1,491.3 百萬（含資金承諾）；9M 注資 73.2 百萬")
inp("inv_k", "Exol 投資計入比例", "x", 0.0, 1.0, None, "Assumed", "基準按帳面；低＝0（持續虧損、資金承諾）", "")
inp("tra_v", "應收稅負協議（TRA）預期未來支付（未入帳）", "US$B", None, 0.487, None, "Verified", Q10 + "：TRA 預期未來支付 487.0 百萬，2026-06-27 未認列負債", "Panabee 轉述 10-K 稱 TRA 65 億（全部換股情境；資料商，未查核，不作依據）")
inp("tra_k", "TRA 扣除比例", "x", None, 0.0, 1.0, "Assumed", "基準 0：本模型以全部股數（A＋V 視同轉換）估每單位價值，TRA 是單位持有人之間的移轉，且對應的稅基墊高利益未建模", "高＝1：全額視為類債（A 類股東視角）")
inp("rf", "無風險利率（美國 10 年期公債）", "%", None, 0.051, None, "Verified", UST, "與股價基準日相差逾 2 個交易日（逾期），維持跨模型一致")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
B1, B2 = 0.67 * 2.11 + 0.33, 0.67 * 1.98 + 0.33
BETA = statistics.median([B1, B2])
inp("beta", "β（兩來源 Blume 調整後中位數）", "x", round(BETA - 0.3, 6), "=MEDIAN(0.67*2.11+0.33,0.67*1.98+0.33)", round(BETA + 0.3, 6), "Derived",
    SAS + "：SYM 5 年 β 2.11（頁面資料日 2025-11）；同業 5 年 β 中位數 1.98（AutoStore 1.98、Daifuku 1.03、KION 2.05；Geek+ 未取得）", "Blume 0.67β＋0.33；敏感度 ±0.3（Assumed）")
inp("kd", "稅前負債成本（無借款，權重 0）", "%", None, "=rf+0.02", None, "Assumed", "rf＋2pt（無借款，不影響 WACC）", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("fc_da", "折舊攤銷 ÷ 營收", "%", 0.012, "=h_da_9m/h_rev_9m", 0.025, "Derived", "＝9M FY2026 30.277 ÷ 2,027.303（" + Q10 + "）", "低／高 Assumed")
inp("fc_cx", "資本支出 ÷ 營收（含資本化軟體）", "%", 0.02, "=h_cx_9m/h_rev_9m", 0.05, "Derived", "＝9M FY2026 62.753 ÷ 2,027.303；FY2025 79.030 ÷ 2,246.922 ＝ 3.5%", "低／高 Assumed")
inp("nwc_26", "營運資金 ÷ 營收（FY2026 年末；應收＋未開票應收＋存貨 − 遞延收入）", "%", None, "=(h_ar_q3+h_uar_q3+h_inv_q3-h_drc_q3-h_drn_q3)/(h_rev_9m+h_rev_q4fy25)", None, "Derived",
    "＝2026-06-27 (288.5＋459.8＋220.8 − 1,553.7 − 182.8) ÷ 近四季營收 2,645.8 ＝ −29.0%（" + Q3K + "）", "FY2025 年末 −37.1%（Hist 頁）：預收款緩衝已在縮小")
inp("nwc_lt", "營運資金 ÷ 營收（FY2030 起；FY2026 起線性收斂）", "%", -0.29, -0.15, -0.05, "Assumed",
    "預收款（遞延收入）隨 backlog 轉營收而回吐：已收的現金不再是未來現金流", "低＝維持 −29%（預收比例不變）；高＝−5%（預收大幅回吐）")
inp("w_dcf", "DCF 權重（其餘為 Comparable：同業前瞻本益比）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺 Symbotic 對同業的歷史相對本益比；公司 2022 年上市，FY2025 前 GAAP 虧損）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("E. 同業前瞻 P/E（資料商頁面報價日各異，以 2026-11-20 收盤等比例換算至同一估值日；分母 EPS 不變；Derived）")
PEERS = [("pe_auto", "AutoStore AUTO.OL（立方格 AS/RS）", "=26.20*14.79/14.68", "2026-10-22 頁面 26.20 倍、收盤 14.68 挪威克朗 → 2026-11-20 收盤 14.79"),
         ("pe_daif", "Daifuku 大福 6383.T（物料搬運系統）", "=21.33*4129/5429", "2026-09-11 頁面 21.33 倍、收盤 5,429 日圓 → 2026-11-20 收盤 4,129"),
         ("pe_kion", "KION KGX.DE（Dematic 倉儲系統整合＋堆高機）", "=9.85*29.2/40.58", "2026-09-18 頁面 9.85 倍、收盤 40.58 歐元 → 2026-11-20 收盤 29.20"),
         ("pe_geek", "極智嘉 Geek+ 2590.HK（AMR 貨到人）", "=56.32*5.405/9.22", "2026-09-04 頁面 56.32 倍、收盤 9.22 港元 → 2026-11-20 收盤 5.405")]
for k, lab, v, note in PEERS:
    inp(k, f"前瞻 P/E：{lab}（換算至 2026-11-20）", "x", None, v, None, "Derived", f"{SAS}：{note}（{YCL}）",
        "原值 Interested-party（賣方共識 EPS）；會計口徑：AutoStore、KION IFRS（租賃資本化）、Daifuku 日本基準、Geek+ IFRS；P/E 為股權口徑不受租賃處理影響", DV)
inp("pe_med", "同業前瞻 P/E 中位數（4 家）", "x", None, "=MEDIAN(pe_auto,pe_daif,pe_kion,pe_geek)", None, "Derived", "少於 8 家，不做迴歸（第 10 節第 7 條）；Ocado 虧損不列", "")
inp("cons_tp", "賣方平均目標價（只作對照，不進模型）", "US$", None, 62.50, None, "Interested-party", SAF + "：17 位，平均 62.50、中位數 67.50、區間 38–85；評等 Buy", "Q4 FY2026 財報前", DV)
inp("cons_rev26", "共識 FY2026 營收（只作對照，不進模型）", "US$B", None, 2.80, None, "Interested-party", SAF + "：16 位平均 2.80（+24.7%）", "", DV)
inp("cons_eps26", "共識 FY2026 EPS（non-GAAP，A 類；只作對照，不進模型）", "US$", None, 0.57, None, "Interested-party", SAF + "：16 位平均 0.57（0.50–0.68）", "", DV)
inp("sb_a", "軟銀持有 A 類比例（只作對照，不進模型）", "%", None, 0.368, None, "Derived", PROF, "比例以委託書當時 A 類股數為分母", DV)
inp("wmt_a", "Walmart 持有 A 類比例（只作對照，不進模型）", "%", None, 0.122, None, "Derived", PROF, "同上", DV)

# ---------------- Hist（年度＋9M＋季度） ----------------
# (鍵, 項目, FY2023, FY2024, FY2025, 9M FY2026, 標記, 來源)
HROWS = [("sys", "Systems 營收", 1.138059, 1.705440, 2.118868, 1.895740),
         ("sw", "Software maintenance and support 營收", 0.006601, 0.014173, 0.029602, 0.036574),
         ("svc", "Operation services 營收", 0.032231, 0.068566, 0.098452, 0.094989),
         ("rev", "合併營收", 1.176891, 1.788179, 2.246922, 2.027303),
         ("csys", "Systems 營業成本", 0.940076, 1.466841, 1.710446, 1.489031),
         ("csw", "Software 營業成本", 0.009222, 0.008949, 0.008562, 0.009808),
         ("csvc", "Operation services 營業成本", 0.037854, 0.066723, 0.105304, 0.084178),
         ("cogs", "營業成本合計", 0.987152, 1.542513, 1.824312, 1.583017),
         ("gp", "毛利", 0.189739, 0.245666, 0.422610, 0.444286),
         ("rd", "研發費用", 0.195042, 0.173457, 0.216013, 0.138069),
         ("sga", "銷售及管理費用", 0.217927, 0.188934, 0.298730, 0.258020),
         ("rst", "重組費用", 0.0, 0.0, 0.022873, 0.002685),
         ("opx", "營業費用合計", 0.412969, 0.362391, 0.537616, 0.398774),
         ("op", "營業利益（損失）", -0.223230, -0.116725, -0.115006, 0.045512),
         ("oi", "業外收入淨額", 0.010716, 0.037042, 0.036063, 0.054688),
         ("pbt", "稅前利益（權益法前）", -0.212514, -0.079683, -0.078943, 0.100200),
         ("tax", "所得稅利益（費用為負）", 0.004620, -0.004212, 0.001627, -0.000038),
         ("eqm", "權益法投資損益", 0.0, -0.000777, -0.013716, -0.022375),
         ("ni", "稅後淨利（含非控制權益）", -0.207894, -0.084672, -0.091032, 0.077787),
         ("nci", "非控制權益損益", -0.184028, -0.071182, -0.074095, 0.061543),
         ("nia", "歸屬 A 類股東淨利", -0.023866, -0.013490, -0.016937, 0.016244),
         ("eps", "EPS（GAAP，A 類稀釋，美元）", -0.37, -0.14, -0.16, 0.12),
         ("ocf", "營業現金流", 0.230794, -0.058077, 0.866939, 0.305584),
         ("cx", "資本支出（含資本化軟體）", 0.021326, 0.044374, 0.079030, 0.062753),
         ("fcf", "自由現金流（營業現金流 − 資本支出）", 0.209468, -0.102451, 0.787909, 0.242831),
         ("da", "折舊攤銷", None, None, None, 0.030277),
         ("sbc", "股權報酬", None, None, None, 0.142919),
         ("cash", "現金及約當現金（期末）", 0.258770, 0.727310, 1.244993, 1.746446)]
HSRC = {"FY2023": FY24 + "（FY2023 比較欄，原公告口徑）", "FY2024": FY25 + "（FY2024 比較欄；較原公告營收 1.822 少 0.034，公司重編）",
        "FY2025": FY25, "9M FY2026": Q3K + "；折舊攤銷、股權報酬取自 " + Q10 + " 現金流量表"}
QI = [("sw_q3", "Q3 FY2026 Software 營收", 0.012765, Q3K), ("svc_q3", "Q3 FY2026 Operation services 營收", 0.037121, Q3K),
      ("opx_q3", "Q3 FY2026 營業費用合計", 0.128015, Q3K), ("gp_q2", "Q2 FY2026 毛利", 0.149952, Q2K), ("rev_q2", "Q2 FY2026 營收", 0.676480, Q2K),
      ("rev_q4fy25", "Q4 FY2025 營收（FY2025 − 9M FY2025）", 0.618457, FY25),
      ("ar_q3", "應收帳款（2026-06-27）", 0.288533, Q3K), ("uar_q3", "未開票應收（2026-06-27）", 0.459843, Q3K), ("inv_q3", "存貨（2026-06-27）", 0.220841, Q3K),
      ("drc_q3", "遞延收入－流動（2026-06-27）", 1.553749, Q3K), ("drn_q3", "遞延收入－非流動（2026-06-27）", 0.182810, Q3K),
      ("ar_25", "應收帳款（2025-09-27）", 0.186705, Q3K), ("uar_25", "未開票應收（2025-09-27）", 0.181658, Q3K), ("inv_25", "存貨（2025-09-27）", 0.164390, Q3K),
      ("drc_25", "遞延收入－流動（2025-09-27）", 1.242312, Q3K), ("drn_25", "遞延收入－非流動（2025-09-27）", 0.124932, Q3K)]

# 季度營收線（Q 頁）：(期間, Systems, Software, Services, 合併, 標記, 來源)
QROWS = [("FY2023", "=h_sys_23", "=h_sw_23", "=h_svc_23", "=h_rev_23", "Verified", FY24),
         ("FY2024", "=h_sys_24", "=h_sw_24", "=h_svc_24", "=h_rev_24", "Verified", FY25),
         ("FY2025", "=h_sys_25", "=h_sw_25", "=h_svc_25", "=h_rev_25", "Verified", FY25),
         ("9M FY2025", 1.536539, 0.020331, 0.071595, 1.628465, "Verified", Q3K + "（前年同期欄）"),
         ("Q3 FY2025", 0.559108, 0.008121, 0.024892, 0.592121, "Verified", Q3K + "（前年同期欄）"),
         ("Q4 FY2025（FY2025 − 9M）", "=B5-B6", "=C5-C6", "=D5-D6", "=E5-E6", "Derived", "年度減前三季"),
         ("Q1 FY2026（9M − Q2 − Q3）", "=B12-B10-B11", "=C12-C10-C11", "=D12-D10-D11", "=E12-E10-E11", "Derived", "九個月減第 2、3 季"),
         ("Q2 FY2026", 0.634496, 0.012924, 0.029060, 0.676480, "Verified", Q2K),
         ("Q3 FY2026", 0.670952, 0.012765, 0.037121, 0.720838, "Verified", Q3K),
         ("9M FY2026", "=h_sys_9m", "=h_sw_9m", "=h_svc_9m", "=h_rev_9m", "Verified", Q3K)]


def build(path, ov=None, sens=None, rev=None):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 US$B（十億美元）、股數百萬股、EPS 與股價美元"
    IN["A2"] = "標記：Verified（附出處）／Interested-party（寫明是誰、誘因欄）／Analogy（寫明類比對象）／Assumed（給區間）／Derived（寫明換算）；查不到寫「未找到」"
    hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源1", "來源2／說明", "來源誘因"])
    r = 4; KR = {}
    for row in ROWS:
        if row[0] == "sec":
            IN.cell(r, 1, row[1]).font = BOLD; r += 1; continue
        k, lab, u, lo, b, hi, tag, s1, s2, inc = row
        if k in ov: b = ov[k]
        IN.cell(r, 1, k); IN.cell(r, 2, lab); IN.cell(r, 3, u)
        for col, v in ((4, lo), (5, b), (6, hi)):
            if v is None: continue
            c = IN.cell(r, col, v)
            if col == 5: c.font = BLUE
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2); IN.cell(r, 10, inc)
        dn(k, f"In!$E${r}"); KR[k] = r; r += 1
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("o27", "新簽 Systems 訂單 FY2027"), ("go", "新簽訂單年增 FY2028 起（讀線模型）"), ("apd", "Walmart 門市後場簽約金額"),
           ("comp", "每年新增運作中系統數"), ("gms", "Systems 毛利率 FY2029 目標")]
    for f, lab in FAM:
        IN.cell(r, 1, f + "_scn"); IN.cell(r, 2, lab)
        for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR[f + '_' + s]}")
        IN.cell(r, 8, f"引用 In 頁 {f}_A／{f}_B／{f}_C；{lab}"); r += 1
    IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重")
    for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR['w_' + s]}")
    IN.cell(r, 8, "引用 In 頁 w_A／w_B／w_C；合計＝1"); r += 1
    IN.column_dimensions["B"].width = 56; IN.column_dimensions["H"].width = 70

    # ---------------- Hist ----------------
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（US$B；GAAP）。公式只引用本頁儲存格（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "9M FY2026", "標記", "來源"])
    for i, (k, lab, *vals) in enumerate(HROWS, 3):
        H.cell(i, 1, lab)
        for j, (sfx, v) in enumerate(zip(("23", "24", "25", "9m"), vals)):
            if v is None:
                H.cell(i, 2 + j, "未找到"); continue
            H.cell(i, 2 + j, v).font = BLUE
            dn(f"h_{k}_{sfx}", f"Hist!${'BCDE'[j]}${i}")
        H.cell(i, 6, "Verified"); H.cell(i, 7, "；".join(f"{p}：{s}" for p, s in HSRC.items()))
    i = 3 + len(HROWS)
    H.cell(i, 1, "檢查：三條營收線合計 − 合併營收（應為 0）")
    for j, sfx in enumerate(("23", "24", "25", "9m")):
        H.cell(i, 2 + j, f"=ROUND(h_sys_{sfx}+h_sw_{sfx}+h_svc_{sfx}-h_rev_{sfx},6)")
    H.cell(i + 1, 1, "檢查：營收 − 營業成本 − 毛利（應為 0）")
    for j, sfx in enumerate(("23", "24", "25", "9m")):
        H.cell(i + 1, 2 + j, f"=ROUND(h_rev_{sfx}-h_cogs_{sfx}-h_gp_{sfx},6)")
    H.cell(i + 2, 1, "檢查：毛利 − 營業費用 − 營業利益（應為 0）")
    for j, sfx in enumerate(("23", "24", "25", "9m")):
        H.cell(i + 2, 2 + j, f"=ROUND(h_gp_{sfx}-h_opx_{sfx}-h_op_{sfx},6)")
    H.cell(i + 3, 1, "Systems 毛利率（Derived）")
    for j, sfx in enumerate(("23", "24", "25", "9m")):
        H.cell(i + 3, 2 + j, f"=1-h_csys_{sfx}/h_sys_{sfx}")
    q0 = i + 5
    hdr(H, q0, ["季度與資產負債表項目", "值", "", "", "", "標記", "來源"])
    for j, (k, lab, v, src) in enumerate(QI, q0 + 1):
        H.cell(j, 1, lab); H.cell(j, 2, v).font = BLUE; H.cell(j, 6, "Verified" if "−" not in lab else "Derived"); H.cell(j, 7, src)
        dn(f"h_{k}", f"Hist!$B${j}")
    j = q0 + 1 + len(QI)
    H.cell(j, 1, "營運資金 ÷ 營收：FY2025 年末（Derived）"); H.cell(j, 2, "=(h_ar_25+h_uar_25+h_inv_25-h_drc_25-h_drn_25)/h_rev_25")
    H.cell(j + 1, 1, "營運資金 ÷ 近四季營收：2026-06-27（Derived）"); H.cell(j + 1, 2, "=(h_ar_q3+h_uar_q3+h_inv_q3-h_drc_q3-h_drn_q3)/(h_rev_9m+h_rev_q4fy25)")
    H.cell(j + 2, 1, "FY2023–FY2025 折舊攤銷、股權報酬：未找到（新聞稿未列；10-K 原文未逐項讀取）")
    H.column_dimensions["A"].width = 46

    # ---------------- Q ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 營收線（US$B）。公司為單一報告分部；營收依三條營收線揭露。分部對帳：三線合計 − 合併營收"
    hdr(Q, 2, ["期間（季度截至）", "Systems", "Software maintenance and support", "Operation services", "合併營收", "檢查：分部合計 − 合併營收（應為 0）", "Walmart（Customer A）占營收", "標記", "來源", "來源誘因"])
    # 列：3＝FY2023 … 12＝9M FY2026
    cust = {"FY2025": 0.85, "9M FY2025": 0.851, "Q3 FY2025": 0.838, "Q3 FY2026": 0.905, "9M FY2026": 0.870}
    for i, (p, a, b_, c_, d_, tag, src) in enumerate(QROWS, 3):
        Q.cell(i, 1, p)
        for jj, v in enumerate((a, b_, c_, d_)):
            c = Q.cell(i, 2 + jj, v)
            if isinstance(v, float): c.font = BLUE
        Q.cell(i, 6, f"=ROUND(B{i}+C{i}+D{i}-E{i},6)")
        pk = p.split("（")[0]
        if pk in cust: Q.cell(i, 7, cust[pk]).font = BLUE
        Q.cell(i, 8, tag); Q.cell(i, 9, src + ("；客戶集中度：" + (PAN if pk == "FY2025" else Q10) if pk in cust else ""))
    Q.cell(14, 1, "註：Q1 FY2025、Q2 FY2025 單季未取得（只取得九個月累計）；Walmart 占營收 FY2025 85% 取自資料商轉述 10-K（Panabee），其餘取自 10-Q「Customer A」")
    Q.column_dimensions["A"].width = 30; Q.column_dimensions["I"].width = 60

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：在手訂單（RPO）轉換＋新訂單（命題模組）→ 三條營收線 → 標準損益表（GAAP 架構；預測期稅與每股為 A＋V 視同轉換的經濟口徑）→ 現金流（US$B；FY2023–FY2025 實績、FY2026＝前三季實績＋第 4 季指引、FY2027 起模型）"
    PR = {}
    rr = 3
    labs = [("yr", "年度（數值，公式用）"), ("ex", "既有 backlog：Systems 認列（子列）"), ("bkr", "既有 Systems backlog 剩餘（年末）"),
            ("ord", "新簽 Systems 訂單（當年簽約；不含門市後場）"), ("nw", "新訂單：Systems 認列（子列）"), ("apd", "Walmart 門市後場：Systems 認列（子列）"),
            ("sys", "Systems 營收"), ("chk1", "檢查：Systems 子列合計 − Systems（應為 0）"), ("ops", "運作中系統數（年末，套）"),
            ("sw", "Software maintenance and support 營收"), ("svc", "Operation services 營收"), ("rev", "營收（合併）"),
            ("chk2", "檢查：分部合計 − 合併營收（應為 0）"), ("gr", "營收年增率"), ("gms", "Systems 毛利率"),
            ("cogs", "營業成本"), ("gp", "毛利"), ("gm", "毛利率"), ("opx", "營業費用（研發＋銷管＋重組）"), ("op", "營業利益"), ("opm", "營業利益率"),
            ("oi", "業外收入淨額（其他收入）"), ("pbt", "稅前利益（權益法前）"), ("tax", "所得稅（費用）"), ("eqm", "權益法投資損益（Exol）"),
            ("ni", "稅後淨利（含非控制權益）"), ("nci", "非控制權益損益（V 類股）"), ("nia", "歸屬 A 類股東淨利"), ("nm", "淨利率"), ("niy", "淨利年增率"),
            ("sh", "完全稀釋股數（A＋V＋RSU，百萬股）"), ("eps", "EPS（完全稀釋；歷史＝GAAP A 類，預測＝A＋V 視同轉換）"),
            ("da", "折舊攤銷"), ("cx", "資本支出"), ("nwc", "營運資金（年末）"), ("dwc", "營運資金增加"),
            ("fcff", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加；歷史＝營業現金流 − 資本支出）"), ("div", "股利"), ("fcfad", "自由現金流扣股利（收支關卡用）")]
    HK = {"rev": "rev", "cogs": "cogs", "gp": "gp", "opx": "opx", "op": "op", "oi": "oi", "eqm": "eqm", "ni": "ni", "nci": "nci", "nia": "nia",
          "sys": "sys", "sw": "sw", "svc": "svc", "eps": "eps", "cx": "cx"}
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, (k, _) in enumerate(labs)}
        for k, l in labs: P.cell(R[k], 1, l)
        F, Lc = COL[2027], COL[2032]
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1)
            cell = lambda k, c=c: f"{c}{R[k]}"; prev = lambda k, p=p: f"{p}{R[k]}"
            def w(k, f, c=c): P[f"{c}{R[k]}"] = f
            w("yr", y)
            if y <= 2025:
                sfx = str(y)[2:]
                for k, hk in HK.items(): w(k, f"=h_{hk}_{sfx}")
                w("chk2", f"=ROUND({cell('sys')}+{cell('sw')}+{cell('svc')}-{cell('rev')},6)")
                w("gms", f"=1-h_csys_{sfx}/h_sys_{sfx}")
                w("pbt", f"=h_pbt_{sfx}"); w("tax", f"=-h_tax_{sfx}")
                w("gm", f"={cell('gp')}/{cell('rev')}"); w("opm", f"={cell('op')}/{cell('rev')}"); w("nm", f"={cell('ni')}/{cell('rev')}")
                if y > 2023:
                    w("gr", f"={cell('rev')}/{prev('rev')}-1")
                    w("niy", f'=IF({prev("ni")}>0,{cell("ni")}/{prev("ni")}-1,"N/M（前期虧損）")')
                w("fcff", f"=h_ocf_{sfx}-h_cx_{sfx}"); w("div", 0); w("fcfad", f"={cell('fcff')}-{cell('div')}")
                if y == 2025:
                    w("nwc", "=(h_ar_25+h_uar_25+h_inv_25-h_drc_25-h_drn_25)")
                continue
            if y == 2026:
                q4 = "g4_rev*(1+g_anc)"
                w("ex", f"=h_sys_9m+({q4}-h_sw_q3-h_svc_q3)"); w("nw", 0); w("apd", 0)
                w("bkr", f"=rpo0*sys_sh*(1-bk_cut)-({q4}-h_sw_q3-h_svc_q3)")
                w("sys", f"=h_sys_9m+{q4}-h_sw_q3-h_svc_q3"); w("ops", "=ops_26")
                w("sw", "=h_sw_9m+h_sw_q3"); w("svc", "=h_svc_9m+h_svc_q3")
                w("gp", f"=h_gp_9m+{q4}*g4_gm")
                w("gms", f"=(h_sys_9m-h_csys_9m+{q4}*g4_gm-h_sw_q3*gm_sw-h_svc_q3*gm_svc)/{cell('sys')}")
                w("opx", "=h_opx_9m+h_opx_q3*(1+g4_opx)"); w("oi", "=h_oi_9m+g4_oi")
                w("tax", "=-h_tax_9m"); w("eqm", "=h_eqm_9m+g4_eqm")
                w("nci", f"=h_nci_9m+({cell('ni')}-h_ni_9m)*nci_sh")
                w("sh", "=sh_v")
                w("da", f"=h_da_9m+fc_da*{q4}"); w("cx", f"=h_cx_9m+fc_cx*{q4}")
                w("nwc", f"=nwc_26*{cell('rev')}")
                w("fcff", "=h_fcf_9m+q4_fcf")
            else:
                ex = f"=IF({y}=2027,rpo0*sys_sh*(1-bk_cut)*r12-({COL[2026]}{R['sys']}-h_sys_9m),0)+rpo0*sys_sh*(1-bk_cut)*(1-r12)/n_rem*MAX(0,MIN({y},2026.75+n_rem)-MAX({y}-1,2026.75))"
                w("ex", ex); w("bkr", f"={prev('bkr')}-{cell('ex')}")
                w("ord", f"=o27_{s}*(1+go_{s})^({y}-2027)*(1-wmt_tail*wmt_no)")
                w("nw", f"=SUMPRODUCT(({F}{R['yr']}:{Lc}{R['yr']}<{y})*({F}{R['yr']}:{Lc}{R['yr']}>={y}-L_new)*({F}{R['ord']}:{Lc}{R['ord']}))/L_new")
                w("apd", f"=apd_{s}*(1-wmt_tail)/apd_L*IF(AND({y}>=2029,{y}<2029+apd_L),1,0)")
                w("sys", f"={cell('ex')}+{cell('nw')}+{cell('apd')}")
                w("ops", f"={prev('ops')}+comp_{s}")
                w("sw", f"=sw_rate*({prev('ops')}+{cell('ops')})/2"); w("svc", f"=svc_r*{cell('sys')}")
                w("gms", f"=${COL[2026]}{R['gms']}+(gms_{s}-${COL[2026]}{R['gms']})*MIN(1,({y}-2026)/3)")
                w("gp", f"={cell('sys')}*{cell('gms')}+{cell('sw')}*gm_sw+{cell('svc')}*gm_svc")
                w("opx", f"={prev('opx')}*(1+g_opx)"); w("oi", "=oi")
                w("tax", f"=MAX(0,{cell('pbt')})*tax_w"); w("eqm", "=eqm")
                w("nci", f"={cell('ni')}*nci_sh"); w("sh", "=sh_v")
                w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}")
                w("nwc", f"=(nwc_26+(nwc_lt-nwc_26)*MIN(1,({y}-2026)/4))*{cell('rev')}")
                w("dwc", f"={cell('nwc')}-{prev('nwc')}")
                w("fcff", f"={cell('op')}-MAX(0,{cell('op')})*tax_w+{cell('da')}-{cell('cx')}-{cell('dwc')}")
                w("chk1", f"=ROUND({cell('ex')}+{cell('nw')}+{cell('apd')}-{cell('sys')},6)")
            if y == 2026:
                w("chk1", f"=ROUND({cell('ex')}+{cell('nw')}+{cell('apd')}-{cell('sys')},6)")
            w("rev", f"={cell('sys')}+{cell('sw')}+{cell('svc')}")
            w("chk2", f"=ROUND({cell('sys')}+{cell('sw')}+{cell('svc')}-{cell('rev')},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            w("cogs", f"={cell('rev')}-{cell('gp')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
            w("op", f"={cell('gp')}-{cell('opx')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            w("pbt", f"={cell('op')}+{cell('oi')}")
            w("ni", f"={cell('pbt')}-{cell('tax')}+{cell('eqm')}")
            w("nia", f"={cell('ni')}-{cell('nci')}"); w("nm", f"={cell('ni')}/{cell('rev')}")
            w("niy", f'=IF({prev("ni")}>0,{cell("ni")}/{prev("ni")}-1,"N/M（前期虧損）")')
            w("eps", f"={cell('ni')}/{cell('sh')}*1000")
            w("div", "=pay_r*MAX(0," + cell("nia") + ")"); w("fcfad", f"={cell('fcff')}-{cell('div')}")
        P.cell(R["ex"], 2, "FY2027＝RPO×Systems 占比×12 個月比例 − FY2026 Q4 Systems；其後＝其餘 ÷ 認列年數（自 2027-07 起）")
        P.cell(R["nw"], 2, "各年新簽訂單於簽約次年起 L_new 年平均認列")
        P.cell(R["apd"], 2, "FY2028 簽約、FY2029 起 apd_L 年平均認列")
        P.cell(R["gms"], 2, "FY2026 實績＋指引 → FY2029 收斂到 gms 目標")
        P.cell(R["tax"], 2, "歷史＝GAAP；預測＝稅前 × tax_w（經濟口徑）")
        P.cell(R["eps"], 2, "預測＝稅後淨利（含非控制權益）÷ 完全稀釋股數")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 ×(1 − tax_w)；股權報酬不加回；FY2026＝前三季 FCF＋q4_fcf")
        for k in ("gr", "gms", "gm", "opm", "nm", "niy"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(labs) + 2
    P.column_dimensions["A"].width = 50; P.column_dimensions["B"].width = 34
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "sys", "fcfad"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-09-26（FY2026 年末），明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*erp"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（全部股數 × 現價；US$B）"; D["B4"] = "=px*sh_v/1000"
    D["A5"] = "負債 D（借款）"; D["B5"] = "=debt_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金與投資（2026-06-27）＝現金 − 借款 ＋ Exol × inv_k − TRA × tra_k"; D["B10"] = "=cash_v-debt_v+inv_k*inv_v-tra_k*tra_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND(((px*sh_v/1000)*(rf+beta*erp)+debt_v*kd*(1-tax_w))/((px*sh_v/1000)+debt_v)+wacc_adj-wacc,9)"
    for rr_ in range(3, 11): D[f"B{rr_}"].number_format = "0.00%" if rr_ in (3, 6, 7, 8, 9) else "#,##0.000"
    DR = {}; rr = 13
    for s in SC:
        R = PR[s]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}" for y in range(2027, 2033)]); b = rr + 2
        rows = ["EBIT（營業利益）", "NOPAT", "＋折舊攤銷", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF（應為 0）", "折現年數（期中）", "折現因子", "FCFF 現值"]
        for i, l in enumerate(rows): D.cell(b + i, 1, l)
        for j, y in enumerate(range(2027, 2033)):
            c = "BCDEFG"[j]; pc = COL[y]
            D[f"{c}{b}"] = f"=PL!{pc}{R['op']}"; D[f"{c}{b+1}"] = f"={c}{b}-MAX(0,{c}{b})*tax_w"
            D[f"{c}{b+2}"] = f"=PL!{pc}{R['da']}"; D[f"{c}{b+3}"] = f"=-PL!{pc}{R['cx']}"; D[f"{c}{b+4}"] = f"=-PL!{pc}{R['dwc']}"
            D[f"{c}{b+5}"] = f"=PL!{pc}{R['fcff']}"; D[f"{c}{b+6}"] = f"=ROUND(SUM({c}{b+1}:{c}{b+4})-{c}{b+5},9)"
            D[f"{c}{b+7}"] = j + 0.5; D[f"{c}{b+8}"] = f"=1/(1+wacc)^{c}{b+7}"; D[f"{c}{b+9}"] = f"={c}{b+5}*{c}{b+8}"
        t = b + 11
        items = [("明確期現值合計", f"=SUM(B{b+9}:G{b+9})"),
                 ("NOPAT FY2033＝NOPAT FY2032 ×(1＋g)", f"=G{b+1}*(1+g_tv)"),
                 ("終值（FY2032 年底）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^6"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金與投資（2026-06-27）", "=dcf_nc"),
                 ("＋Stub：FY2026 第 4 季自由現金流", "=q4_fcf"),
                 ("（保留列：無其他調整）", 0),
                 ("股權價值", f"=B{t+4}+B{t+5}+B{t+6}+B{t+7}"),
                 ("每股價值（評價日）", f"=MAX(0,B{t+8})/sh_v*1000"),
                 ("DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）", f"=B{t+9}*(1+ke)^t_tgt"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+11}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=IF(G{b+1}>0,B{t+2}/G{b+1},\"N/M\")"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）", f"=MAX(0,(B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6)+B{t+5}+B{t+6}+B{t+7})/sh_v*1000*(1+ke)^t_tgt")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 10, ps=t + 9, ev=t + 4, tvs=t + 11, alt=t + 14, chk=b + 6)
        dn(f"dcf_{s}", f"DCF!$B${t+10}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 64

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價：DCF（權重 w_dcf）＋ Comparable（同業前瞻本益比 × FY2028 經濟口徑 EPS，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}")
        V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "DCF 頁")
    r_cu = vrow("Comparable（未調整）：同業前瞻 P/E 中位數 × FY2028 EPS（A＋V 視同轉換）", [f"=MAX(0,PL!H{PR[s]['eps']})*pe_med" for s in SC],
                "目標時點 2027-11 起算的 NTM 約 10/12 FY2028＋2/12 FY2029，以 FY2028 代；EPS 為稅後淨利（含非控制權益）÷ 完全稀釋股數")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（美元）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    r_px = one("現價（In 頁 px）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_upc = one("Comparable 目標價上檔空間", f"=B{r_cmpt}/px-1")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（US$B；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=TEXT(px*(1+IF(B{r_up}>0.5,9,IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5))))),"#,##0.00")&"／"&IF(B{r_up}<-0.5,"（已是最低級）",TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,-0.5)))),"#,##0.00"))', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    r_liq = one("期間流動性：情境 B 最低年末現金代理（評價日現金＋累計 FCFF，FY2027–FY2030）", f"=cash_v+q4_fcf+MIN(PL!G{PR['B']['fcff']},SUM(PL!G{PR['B']['fcff']}:PL!H{PR['B']['fcff']}),SUM(PL!G{PR['B']['fcff']}:PL!I{PR['B']['fcff']}),SUM(PL!G{PR['B']['fcff']}:PL!J{PR['B']['fcff']}))", "第 10 節第 55 條；無借款，最大單年新增舉債 0")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2026-09-26 → 目標時點，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 EPS", f'=IF(PL!H{PR["B"]["eps"]}>0,B{r_tp}/PL!H{PR["B"]["eps"]},"N/M")', "分母：模型情境 B FY2028 EPS（經濟口徑）")
    r_pe2 = one("現行前瞻 P/E：現價 ÷ 共識 FY2026 EPS（non-GAAP）", "=px/cons_eps26", "只作對照；共識 2026-10-22、財報前；分母口徑 non-GAAP A 類")
    r_pe3 = one("現價 ÷ 情境 B FY2028 EPS（模型）", f'=IF(PL!H{PR["B"]["eps"]}>0,px/PL!H{PR["B"]["eps"]},"N/M")', "分母：模型情境 B FY2028")
    vr += 1
    V.cell(vr, 1, "股權結構（公開流通股與總股數差異；2026-08-03 股數）").font = BOLD; vr += 1
    r_s1 = one("A 類普通股（公開交易）", "=sh_a")
    r_s2 = one("V-1＋V-3 類（對應 Holdings 單位，可一比一換 A 類；不在市場交易）", "=sh_v1+sh_v3")
    r_s3 = one("全部普通股", "=sh_tot")
    r_s4 = one("A 類占全部普通股", "=sh_a/sh_tot")
    r_s5 = one("市值（全部普通股 × 現價，US$B；本模型估值口徑）", "=sh_tot*px/1000")
    r_s6 = one("只以 A 類計的市值（US$B；若誤用會低估約 4/5）", "=sh_a*px/1000")
    r_s7 = one("估計公開流通 A 類（扣除軟銀、Walmart 持有；百萬股）", "=sh_a*(1-sb_a-wmt_a)", "比例取自委託書（2026-01），分母為當時 A 類股數；Derived，只作對照")
    r_s8 = one("估計公開流通 A 類占全部普通股", f"=B{r_s7}/sh_tot")
    r_s9 = one("企業價值（全部普通股市值 − 淨現金與投資，US$B）", f"=B{r_s5}-dcf_nc")
    r_s10 = one("EV ÷ FY2026 營收（倍）", f"=B{r_s9}/PL!F{PR['B']['rev']}")
    for rr_ in (r_up, r_upc, r_a1, r_a2, r_a3, r_s4, r_s8): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0.00"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "#,##0.00"
    V.column_dimensions["A"].width = 72; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, upc=r_upc, rt=r_rt, rec=r_rec, gate=r_gate, gj=r_gj, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, pd2=r_pd2, pe=r_pe, liq=r_liq, s5=r_s5, s6=r_s6, s4=r_s4, s8=r_s8, s10=r_s10)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司指引、市場共識、在手訂單對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或共識", "差", "說明"])
    B_ = PR["B"]
    rc = [("FY2026 營收 vs 前三季實績＋第 4 季指引中點", f"=PL!F{B_['rev']}", "=h_rev_9m+g4_rev", "公司指引（Interested-party）；差額來自 g_anc"),
          ("FY2026 營收 vs 共識", f"=PL!F{B_['rev']}", "=cons_rev26", "資料商共識（2026-10-22，Q4 財報前，只作對照）"),
          ("第 4 季調整後 EBITDA：模型（營業利益＋折舊＋股權報酬）vs 指引中點", f"=(PL!F{B_['op']}-h_op_9m)+fc_da*g4_rev*(1+g_anc)+sbc_r*g4_rev*(1+g_anc)", "=g4_ebitda", "模型以股權報酬 ÷ 營收近似；口徑不同，只看方向"),
          ("FY2027 營收 vs RPO × 未來 12 個月比例（2026-07～2027-06 視窗）", f"=PL!G{B_['rev']}", "=rpo0*r12", "視窗差一季：RPO 15% 含 FY2026 Q4；模型 FY2027 另含新訂單與軟體服務"),
          ("FY2032 年末既有 Systems backlog 剩餘", f"=PL!L{B_['bkr']}", "=rpo0*sys_sh", "認列年數 n_rem 越長，剩餘越多（backlog 不是現金：剩餘部分要到 FY2033 後才認列）"),
          ("FY2027–FY2032 累計營收：既有 backlog 貢獻占比", f"=SUM(PL!G{B_['ex']}:PL!L{B_['ex']})/SUM(PL!G{B_['rev']}:PL!L{B_['rev']})", "", "其餘來自新訂單、門市後場與軟體服務"),
          ("FY2026 EPS（模型經濟口徑）vs 共識 non-GAAP EPS", f"=PL!F{B_['eps']}", "=cons_eps26", "口徑不同（共識加回股權報酬等），只看方向"),
          ("加權目標價 vs 賣方平均目標價", "=tp", "=cons_tp", "賣方 17 位平均（Interested-party；Q4 財報前）"),
          ("市值隱含：EV ÷ FY2026 營收（倍）", f"=Val!B{VR['s10']}", "", "全部普通股口徑")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f'=IF(ISNUMBER(C{i}),B{i}-C{i},"")'); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_symbotic.py 網格＋線性內插；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 64; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_symbotic.py 以 LibreOffice 重算每一案例）；結果為常數快照"
    base_tp = sens["base"]["tp"] if sens else None
    SE["A2"] = "快照基準值（加權，美元）"; SE["B2"] = base_tp
    SE["C2"] = "本檔即時值（加權，美元）"; SE["D2"] = "=tp"
    SE["E2"] = '=IF(ABS(D2-B2)>0.005,"STALE：In 頁已改，Sens 須重跑","OK：快照與本檔一致")'
    SE["F2"] = f"快照輸入指紋 {sens['fp'] if sens else ''}"
    hdr(SE, 4, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    sr = 5
    if sens:
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]); SE.cell(sr, 5, cs["tp"] / base_tp - 1)
            SE.cell(sr, 6, cs["dcf"]); SE.cell(sr, 7, (cs["dcf"] / sens["base"]["dcf"] - 1) if sens["base"]["dcf"] else 0); SE.cell(sr, 8, cs["rt"])
            SE[f"E{sr}"].number_format = "0.0%"; SE[f"G{sr}"].number_format = "0.0%"; SE[f"D{sr}"].number_format = "#,##0.00"; SE[f"F{sr}"].number_format = "#,##0.00"
            sr += 1
    SE.cell(sr + 1, 1, "說明：相對倍數係數 rel_k 基準即為 1.00（未調整），故「相對倍數係數＝1」案例＝基準。無區間的鍵不列案例。尾端案例（bk_cut、wmt_tail、tra_k）基準為 0。")
    sb = sr + 3
    SE.cell(sb, 1, "build_id"); SE.cell(sb, 4, "{BID}")
    SE.cell(sb + 1, 1, "模型狀態"); SE.cell(sb + 1, 4, '=IF(ABS(D2-B2)>0.005,"STALE","CURRENT")')
    SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 34

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = "對帳 — 本模型使用的線層級變數（線 5 倉儲物流自動化；讀 " + WHL + "；與線模型不一致處寫明，不靜默改口徑）"
    hdr(RE, 2, ["變數", "所屬線", "線模型值（低／基準／高）", "本模型使用值", "單位", "口徑", "標記", "影響本模型哪些儲存格／差異說明"])
    rec = [("AS/RS 市場規模 2026（WH_ASRS_MKT_2026）", "5 倉儲", "10.45／10.69／10.93", "同左（只取 2026→2032 成長率）", "十億美元", "MarketsandMarkets 口徑（含製造、汽車）", "Derived", "In go_A／go_B／go_C 分母；下游取數 C8:E8"),
           ("AS/RS 市場規模 2032（WH_ASRS_MKT_2032）", "5 倉儲", "14.28／16.94／19.60", "同左", "十億美元", "同上", "Derived", "In go_* 分子；下游取數 C11:E11"),
           ("新簽訂單年增 FY2028 起（由上兩列推得）", "5 倉儲", "5.3%／8.0%／10.2%", "=TEXT(go_C,\"0.0%\")&\"／\"&TEXT(go_B,\"0.0%\")&\"／\"&TEXT(go_A,\"0.0%\")", "%/年", "市場年複合當作 Symbotic 新簽訂單成長", "Derived", "PL 各情境「新簽 Systems 訂單」列；一致"),
           ("Symbotic 營收 ÷ AS/RS 市場 2026（隱含份額）", "5 倉儲", "—", "=TEXT(PL!F" + str(PR['B']['rev']) + "/10.69434,\"0%\")", "比例", "口徑不同：線模型 AS/RS 市場以立方格與堆垛機為主，Symbotic 整箱棧板系統是否全數計入未確認", "Derived", "只作量級對照，未進公式；份額看似偏高反映口徑差，不據以下結論"),
           ("美國倉儲全負擔時薪 2026（WH_W_LOADED_US）", "5 倉儲", "33.83／35.13／36.43", "未進公式", "美元/時", "BLS NAICS 493 × 全負擔係數", "Derived", "客戶端回本期的背景；Symbotic 系統未在線模型單獨建模（線模型說明）"),
           ("AS/RS（立方格）回本期（WH_ASRS_PAYBACK）", "5 倉儲", "20.05／2.16／0.71 年", "未進公式", "年", "每埠＝每座倉，美國 2026", "Derived", "Symbotic 整箱棧板系統屬不同方案，線模型未單獨建模；本模型不以此推需求"),
           ("Symbotic backlog、12 個月交付比例、運作中系統（追蹤指標 D8）", "5 倉儲", "225 億；約 12%；運作 56、部署 77", "225 億；15%（10-Q 新值）；56、77", "—", "線模型追蹤頁（框架 v0.1，2026-08）", "Verified", "不一致：12 個月比例線模型 12%（FY2025 10-K）、本模型 15%（2026-06-27 10-Q 新值）；r12 低案即 12%"),
           ("人形每人工當量小時成本（WH_HUMANOID_COST_HR）", "5 倉儲", "155.6／40.0／12.5", "未進公式", "美元/時", "Agility 示意 RaaS", "Derived", "人形進倉經濟性未過關（40 > 35.1），對 Symbotic 整箱作業的替代威脅列風險登錄")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 46; RE.column_dimensions["C"].width = 30; RE.column_dimensions["H"].width = 60

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模或只在 Sens 呈現的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("Walmart 集中（9M FY2026 占營收 87%、Q3 90.5%；同時是股東）", "是（尾端 Sens：wmt_tail、bk_cut）", "單一客戶占營收一成以上（第 10 節第 19 條）", "10-Q Customer A 比例、Walmart 新配送中心與門市後場決定"),
             ("財務報導內部控制重大缺失（FY2025 會計師出具否定意見：成本與營收認列時點）", "否", "百分比完工法下認列時點可能再調整；FY2024 曾重編季報", "10-K 內控意見、重編公告"),
             ("Exol（原 GreenBox）資金承諾與虧損", "部分（eqm、inv_k）", "最大損失曝險 14.9 億美元（含資金承諾），未建資本注入現金流", "10-Q VIE 附註、Exol 場址上線與客戶數"),
             ("預收款回吐（遞延收入 17.4 億）", "是（nwc_lt）", "已收現金不再是未來現金流；backlog 轉營收時現金流低於營收增長", "遞延收入、未開票應收季變動"),
             ("應收稅負協議（TRA）", "部分（tra_k Sens）", "V 類換股時支付稅負節省的 85%；預期未來支付 4.87 億（10-Q）", "換股量、TRA 負債認列"),
             ("多類股制與治理（創辦人家族掌多數投票權）", "否", "V-3 每股 3 票；公開流通 A 類約占全部股數一成", "委託書受益所有權表、換股公告"),
             ("關稅與保固成本", "否", "已在毛利率歷史實績內；未逐項建模", "Systems 毛利率季變動"),
             ("人形與移動操作機器人替代整箱作業", "否", "線模型：人形每人工當量小時成本 40 美元，高於倉儲時薪 35.1", "GXO、Agility 試點結果")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 56; RK.column_dimensions["C"].width = 60

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([Q3K, Q2K, Q10, FY25, FY24, CALL, PAN, SAF, SAS + "：AUTO（osl）、6383（tyo）、KGX（etr）、2590（hkg）、SYM", YCL, UST, WHL, PROF], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"Symbotic（SYM）個股模型 {VER}｜build_id {{BID}}"
    RD["A2"] = '="資料截止日 2026-11-23（第 4 季 FY2026 財報尚未公布；最新為 2026-08-05 第 3 季）；成品日 2026-11-23；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.00")&" 美元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：Symbotic 的股價取決於 225 億美元在手訂單（RPO）能多快、以多少毛利轉成營收與現金，以及 Walmart（占營收約九成）以外能否簽到足夠的新訂單；在手訂單不是現金——已預收的部分不會再流入，未來 12 個月只認列約 15%"
    RD["A4"] = (f'="在 RPO 其餘部分 "&TEXT(n_rem,"0")&" 年認列、情境 B 新簽 Systems 訂單 FY2027 "&TEXT(o27_B,"0.0")&" 十億美元、Systems 毛利率 FY2029 達 "&TEXT(gms_B,"0%")&"、β "&TEXT(beta,"0.00")&" 的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 美元（DCF "&TEXT(Val!B{r_dcft},"0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"0.00")&" 十億美元）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 美元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：RPO 未來 12 個月比例升到 20% 以上且 RPO 不降（轉換加速）；季度新簽約明顯高於認列（book-to-bill 持續 > 1.3）；Walmart 簽下門市後場 400 套；Systems 毛利率連兩季 > 25%"
    RD["A7"] = "範圍：單一報告分部，營收以 Systems、Software maintenance and support、Operation services 三條營收線建模；不讀 Siliconomics／Tokenomics／Wattonomics；線 5 倉儲線模型只讀 AS/RS 市場成長率。HTML 展示層見 dist/ 同名 .html"
    RD["A8"] = f'="股權結構：A 類 "&TEXT(sh_a,"0.0")&" 百萬股只占全部普通股 "&TEXT(sh_tot,"0.0")&" 百萬股的 "&TEXT(Val!B{r_s4},"0.0%")&"（扣除軟銀、Walmart 持有後的估計公開流通約 "&TEXT(Val!B{r_s8},"0%")&"）；V 類對應 Holdings 單位、可一比一換 A 類，估值以全部股數計：市值 "&TEXT(Val!B{r_s5},"0.0")&" 十億美元，只算 A 類會誤得 "&TEXT(Val!B{r_s6},"0.0")'
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值（批次） ----------------
LO = "file://" + os.path.join(os.environ.get("HOME") or tempfile.gettempdir(), "lo_profile_sym")
def recalc_many(srcs, outdir):
    os.makedirs(outdir, exist_ok=True)
    for i in range(0, len(srcs), 12):
        chunk = srcs[i:i + 12]
        for attempt in range(2):
            try:
                subprocess.run(["soffice", f"-env:UserInstallation={LO}", "--headless", "--convert-to", "xlsx", "--outdir", outdir] + chunk,
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=175)
                break
            except subprocess.TimeoutExpired:
                subprocess.run(["pkill", "-f", "lo_profile_sym"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                if attempt == 1: raise
    return [os.path.join(outdir, os.path.basename(s)) for s in srcs]

def read_vals(path, VR, DR=None):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]
    d = dict(tp=V[f"B{VR['tp']}"].value, dcf=V[f"B{VR['dcf']}"].value, cmp=V[f"B{VR['cmp']}"].value, up=V[f"B{VR['up']}"].value, rt=V[f"B{VR['rt']}"].value)
    if DR:
        I = wb["In"]; w = {}; wd = None
        for row in I.iter_rows(values_only=True):
            if row[0] in ("w_A", "w_B", "w_C"): w[row[0][-1]] = row[4]
            if row[0] == "w_dcf": wd = row[4]
        alt = sum(w[s] * wb["DCF"][f"B{DR[s]['alt']}"].value for s in SC)
        d["alt_dcf"] = alt; d["alt_tp"] = wd * alt + (1 - wd) * d["cmp"]
    return d

def run_cases(ovs):
    d = tempfile.mkdtemp(prefix="sym_"); paths = []; VRs = []
    for i, ov in enumerate(ovs):
        p = os.path.join(d, f"c{i:03d}.xlsx"); VR, _, DR = build(p, ov); paths.append(p); VRs.append((VR, DR))
    outs = recalc_many(paths, os.path.join(d, "o"))
    res = [read_vals(o, VR, DR) for o, (VR, DR) in zip(outs, VRs)]
    shutil.rmtree(d, ignore_errors=True)
    return res

def in_val(k, col):
    for row in ROWS:
        if row[0] == k: return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)

def fingerprint():
    return hashlib.sha1(json.dumps([r for r in ROWS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]

LABEL = {r[0]: r[1] for r in ROWS if r[0] != "sec"}
def rate(u): return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"


def write_readme(path, VR, PR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]
    g = lambda c: V[c].value
    f2 = lambda x: f"{x:,.2f}"; pc = lambda x: f"{x*100:+.1f}%"
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:8]
    def fv(x): return x if isinstance(x, str) else f"{x:.4g}"
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {fv(r['val'])} | {r['base']} | {r['note']} |" for r in revs)
    pev = g(f"B{VR['pe']}"); pe_txt = pev if isinstance(pev, str) else f"{pev:.1f}"
    def yr(s, k, fmt="{:,.2f}"):
        out = []
        for y in YRS:
            v = P[f"{COL[y]}{PR[s][k]}"].value
            out.append(fmt.format(v) if isinstance(v, (int, float)) else "—")
        return " | ".join(out)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f2(g(chr(66+i)+str(VR['dcfr'])))} | {f2(g(chr(66+i)+str(VR['cu'])))} | {f2(g(chr(66+i)+str(VR['tps'])))} |" for i, s in enumerate(SC))
    Iv = {wb['In'][f'A{r}'].value: wb['In'][f'E{r}'].value for r in range(1, wb['In'].max_row + 1)}
    rts = sorted(set(c['rt'] for c in sens['cases']))
    txt = f"""# physical_ai/symbotic — Symbotic（SYM，Nasdaq）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

線 5 倉儲物流自動化。本 README 的數字由 `build_symbotic.py` 從重算後的成品讀出產生，不手寫。

**本版沒有做的部分：** 第 4 季 FY2026（截至約 2026-09-26）財報在本版截止前尚未公布，FY2026 以前三季實績＋公司第 4 季指引中點代；FY2023–FY2025 折舊攤銷與股權報酬、RSU 流通總數、租賃負債、Walmart 2022 年認股權證現況未找到；RPO 的 Systems 占比與 12 個月以後的認列時程公司未揭露（Assumed，承重）。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A8'].value}

## 估值（評價日 2026-09-26；目標時點約 12 個月）
| 情境 | DCF（推到目標時點） | Comparable（同業前瞻 P/E） | 情境目標價 |
|---|---|---|---|
{scn}

- 加權目標價 **{f2(tp)} 美元**（DCF {f2(dcf)}、Comparable {f2(cmp_)}），現價 {f2(Iv['px'])}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")}（情境 B FY2026–FY2030 自由現金流扣股利 {g(f"C{VR['gate']}"):.2f} 十億美元）→ 投資建議 **{g(f"B{VR['rec']}")}**。
- 方法：DCF 標準版 FCFF（{V['B' + str(VR['pe'] - 1)].value}，為 WACC／ke／g／RONIC；無借款，WACC＝ke）50%＋Comparable 50%（AutoStore、Daifuku、KION、極智嘉 4 家前瞻 P/E 中位數 {Iv['pe_med']:.1f} 倍 × 情境 FY2028 EPS；相對倍數係數 1.00，未調整）。
- 隱含前瞻 P/E（加權目標價 ÷ 情境 B FY2028 EPS）{pe_txt} 倍；賣方平均目標價 {f2(Iv['cons_tp'])}（17 位，2026-10-22，Q4 財報前）。
- 現價量級核對：csv 2026-11-20 收盤 {f2(Iv['px'])}，資料商 2026-10-23 收盤 {f2(Iv['px_oct'])}，同一量級。

## 在手訂單不是現金：三個處理
1. **轉換時程：** RPO {Iv['rpo0']:.1f} 十億美元（2026-06-27），公司揭露約 {Iv['r12']*100:.0f}% 於未來 12 個月認列；其餘依 Assumed 的 {Iv['n_rem']:.0f} 年平均認列。情境 B 下 FY2027–FY2032 累計營收中既有 backlog 貢獻 {wb['Recon']['B8'].value*100:.0f}%，FY2032 年末既有 Systems backlog 仍剩 {wb['Recon']['B7'].value:.1f} 十億美元。
2. **預收款回吐：** 營運資金 ÷ 營收由 FY2025 年末 −37% 到 2026-06-27 −29%（遞延收入 17.4 億美元，已收現金），基準假設 FY2030 收斂到 {Iv['nwc_lt']*100:.0f}%，回吐在 FCFF 扣除；現金 {Iv['cash_v']:.2f} 十億美元只加一次。
3. **客戶集中：** Walmart（Customer A）占營收 9M FY2026 87.0%、Q3 90.5%、占應收 82.8%；尾端 Sens：Walmart 停止新簽約（含門市後場 400 套）、既有 Systems backlog 取消 20%。
4. **股權結構：** 見上方股權結構句；估值、市值、EPS 一律用全部股數（A＋V＋RSU），歷史 EPS 為 GAAP A 類口徑。

## 營收線與損益（情境 B，US$B）
| 列 | {" | ".join(f"FY{y}" for y in YRS)} |
|---|{"---|" * len(YRS)}
| 既有 backlog：Systems 認列 | {yr('B','ex')} |
| 新訂單：Systems 認列 | {yr('B','nw')} |
| Walmart 門市後場：Systems 認列 | {yr('B','apd')} |
| Systems 營收 | {yr('B','sys')} |
| Software maintenance and support | {yr('B','sw','{:,.3f}')} |
| Operation services | {yr('B','svc','{:,.3f}')} |
| 營收（合併） | {yr('B','rev')} |
| 毛利率 | {yr('B','gm','{:.1%}')} |
| 營業利益 | {yr('B','op','{:,.3f}')} |
| 稅後淨利（含非控制權益） | {yr('B','ni','{:,.3f}')} |
| EPS（美元） | {yr('B','eps')} |
| FCFF | {yr('B','fcff','{:,.3f}')} |

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| RPO 認列年數（n_rem）、Systems 占比（sys_sh） | 10-Q RPO 與「未來 12 個月」比例；部署中與運作中系統數 |
| 新簽 Systems 訂單（o27） | 季度 RPO 變動＋認列營收（隱含 book-to-bill） |
| Systems 毛利率（gms） | Systems 季度毛利率；新一代儲存結構上線時點 |
| 門市後場（apd） | Walmart 第二個門市場址結果與擴大決定 |
| 營運資金收斂（nwc_lt） | 遞延收入、未開票應收季變動 |
| β | 本身 5 年 β 2.11（資料商，2025-11 頁面） |

{RD['A6'].value}。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f2(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁；各案例估值評等出現「{'、'.join(rts)}」。

## 反解（整本重算；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
讀 `physical_ai/lines/warehouse/dist/20261116_倉儲物流單位經濟_v0.1.xlsx`「下游取數」：WH_ASRS_MKT_2026／2032 的低／基準／高 → 新簽訂單年增 go_C／go_B／go_A（{Iv['go_C']*100:.1f}%／{Iv['go_B']*100:.1f}%／{Iv['go_A']*100:.1f}%）。與線模型不一致：線模型追蹤頁的 12 個月交付比例 12%（FY2025 10-K）較舊，本模型用 2026-06-27 10-Q 的 15%（低案 12%）。線模型未單獨建模 Symbotic 整箱棧板系統的回本期，本模型不以 AMR／AS/RS 回本期推需求。

## 限制與未找到資料
- 第 4 季 FY2026 財報未公布；RPO 的 Systems 占比、12 個月以後的認列時程未揭露（Assumed）。
- FY2023–FY2025 折舊攤銷、股權報酬未找到（新聞稿未列；10-K 原文未逐項讀取）；FY2023 為原公告口徑（FY2024 曾重編季報，FY2025 10-K 的 FY2024 比較欄較原公告營收少 0.034）。
- RSU 流通總數、租賃負債、Walmart 2022 年認股權證現況未找到；稀釋股以 Q3 稀釋與基本加權股數差代。
- 同業前瞻 P/E 取自資料商頁面（報價日 2026-09-04～10-22），以 2026-11-20 收盤等比例換算（Derived，分母 EPS 不變）；只有 4 家，不做迴歸；KION、Daifuku、極智嘉的 csv 收盤較頁面報價日低 24–41%。
- SYM β 取自資料商 2025-11 頁面（未更新）；rf 取 2026-11-12 值（逾期，跨模型一致）。
- 股權結構中軟銀、Walmart 持有 A 類比例轉述自委託書（2026-01），未回原文逐字核對，只作對照。
- 未做：未參與製作的代理獨立核對（Project 規則第 10 條，待主代理安排）；tools/model_governance/check_display.py（無 display_check.json）。

## 檔案
- `build_symbotic.py`：建置＋整本重算 Sens／反解＋產生本 README 與 `segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-23）：初版；現價 {f2(Iv['px'])}（2026-11-20 收盤）；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


def main():
    if "--raw" in sys.argv:
        ov = json.loads(os.environ.get("SYM_OVERRIDE", "{}"))
        build(sys.argv[sys.argv.index("--raw") + 1], ov); return
    px0 = in_val("px", "base")
    # 案例清單：In 頁每個有低／高且與基準不同的鍵
    specs = []
    for row in ROWS:
        if row[0] == "sec": continue
        k, lab = row[0], row[1]
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if v is None or isinstance(v, str): continue
            b_ = in_val(k, "base")
            if isinstance(b_, (int, float)) and abs(v - b_) < 1e-12: continue
            specs.append((k, side, v, lab))
    ovs = [None] + [{k: v} for k, side, v, lab in specs] + [{f"w_{x}": (1.0 if x == s else 0.0) for x in SC} for s in SC]
    res = [None] + run_cases(ovs[1:])
    # 基準需要 alt
    d = tempfile.mkdtemp(prefix="symb_"); p = os.path.join(d, "base.xlsx"); VR, _, DR = build(p, {})
    base = read_vals(recalc_many([p], os.path.join(d, "o"))[0], VR, DR); shutil.rmtree(d, ignore_errors=True)
    print("基準", base)
    cases = []
    for (k, side, v, lab), r in zip(specs, res[1:1 + len(specs)]):
        cases.append(dict(drv=lab + ("（低）" if side == "lo" else "（高）"), set=f"{k}＝{v:.4g}", key=k, **r))
    for s, r in zip(SC, res[1 + len(specs):]):
        cases.append(dict(drv=f"只看情境 {s}", set=f"w_{s}＝1，其餘 0", key=f"w_{s}", **r))
    up_ = base["alt_tp"] / px0 - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rate(up_)))
    # 反解：網格整本重算 → 線性內插
    up0 = base["up"]; tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    targets = [("目標價＝現價（上檔 0%）", 0.0)] + ([(f"升一級邊界（上檔 {upper:+.0%}）", upper)] if upper is not None else []) + ([(f"降一級邊界（上檔 {lower:+.0%}）", lower)] if lower is not None else [])
    revs = []
    for key, grid, b0, note in (("o27_B", [0.5, 1.5, 2.5, 4, 6, 8, 11, 15, 20, 30], 2.5, "情境 B 新簽 Systems 訂單 FY2027（其後依 go_B 成長；整本重算；情境 A／C 不動）"),
                                ("gms_B", [0.10, 0.17, 0.24, 0.30, 0.36, 0.42, 0.50, 0.60], 0.24, "情境 B Systems 毛利率 FY2029 目標（整本重算；情境 A／C 不動）"),
                                ("beta", [0.6, 0.9, 1.2, 1.5, BETA, 2.0], round(BETA, 4), "β（三情境同動；整本重算）")):
        rs = run_cases([{key: x} for x in grid]); ups = [r["up"] for r in rs]
        for lab, tgt in targets:
            val = None
            for i in range(len(grid) - 1):
                a, b2 = ups[i] - tgt, ups[i + 1] - tgt
                if a == 0: val = grid[i]; break
                if a * b2 < 0: val = grid[i] + (grid[i + 1] - grid[i]) * (-a) / (b2 - a); break
            revs.append(dict(label=lab, key=key, val=val if val is not None else f"區間 {grid[0]}～{grid[-1]} 內無解", base=b0, note=note + "；網格線性內插"))
            print("反解", key, lab, val)
    sens = dict(base=base, cases=cases, fp=fingerprint())
    json.dump(sens, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
    # 成品
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="symf_"); raw = os.path.join(tmp, NAME)
    VR, PR, DR = build(raw, {}, sens, revs)
    bid = f"{VER}-{fingerprint()}"
    wb = openpyxl.load_workbook(raw)
    wb["README"]["A1"] = wb["README"]["A1"].value.replace("{BID}", bid)
    SE = wb["Sens"]
    for r in range(1, SE.max_row + 1):
        if SE.cell(r, 4).value == "{BID}": SE.cell(r, 4).value = bid
    SE["H2"] = f"build_id {bid}"; wb.save(raw)
    out = recalc_many([raw], os.path.join(tmp, "o"))[0]
    shutil.copy(out, os.path.join(dist, NAME))
    wbv = openpyxl.load_workbook(os.path.join(dist, NAME), data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10]); assert not errs, errs
    write_readme(os.path.join(dist, NAME), VR, PR, bid, sens, revs)
    json.dump(dict(company="symbotic", ticker="SYM", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=v, unit="US$B", source=FY25 + "（單一報告分部；三條營收線）")
                                      for n, l, v in (("Systems", "Systems 營收", 2.118868), ("Software maintenance and support", "Software maintenance and support 營收", 0.029602),
                                                      ("Operation services", "Operation services 營收", 0.098452))],
                   sublines=[{"parent": "Systems", "labels": ["既有 backlog：Systems 認列（子列）", "新訂單：Systems 認列（子列）", "Walmart 門市後場：Systems 認列（子列）"]}],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）", "檢查：Systems 子列合計 − Systems（應為 0）"],
                   note="公司為單一報告分部（10-K），營收依三條營收線揭露，本模型以三條營收線為分部列；Q 頁另有 FY2023–9M FY2026 營收線對帳（檢查列為 0）。Systems 三條子列（既有 backlog、新訂單、門市後場）合計＝Systems"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)
    shutil.rmtree(tmp, ignore_errors=True)

if __name__ == "__main__":
    main()

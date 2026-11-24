# KION（KGX.DE，法蘭克福）個股模型 v0.1 — AI實體應用 Project（線 5 倉儲物流自動化）
# 用法：python3 physical_ai/kion/build_kion.py            → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、segments.json、sens.json
#       python3 physical_ai/kion/build_kion.py --raw OUT  → 只寫一本（讀環境變數 KION_OVERRIDE 的 JSON 覆寫 In 頁基準值）
# 金額：十億歐元（EUR B）；股數：百萬股；EPS、股價：歐元。會計年度＝曆年（12 月底）。評價日 2026-12-31。
# 結構：README、In、Hist、PL（三情境：IAS 與 ITS 分部營收 → 調整後 EBIT → 標準損益 → 現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md。實體 AI 線不讀 Siliconomics／Tokenomics／Wattonomics。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
VER = "v0.1"; DATE = "20261124"
NAME = f"{DATE}_KION個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032

Q3R = "KION 新聞稿：9M 2026 結果與全年指引更新（2026-10-29）https://www.kiongroup.com/en/Press/Press-Releases/Press-Releases-Detail.html?id=486077511"
KF = "KION 2025 年報「Key figures for 2025」https://reports.kiongroup.com/2025/ar/kion-group/key-figures-for-2025.html"
KP = "KION 2025 年報「Business situation and financial performance」https://reports.kiongroup.com/2025/ar/management-report/report-on-the-economic-position/financial-position-and-financial-performance/business-situation-and-financial-performance-of-the-kion-group.html"
KM = "KION 2025 年報「Multi-year overview」https://reports.kiongroup.com/2025/ar/kion-group/multi-year-overview.html"
KS = "KION 新聞稿：FY2025 solid financial year（分部數字）https://www.kiongroup.com/en/Newsroom/Story-Categories/Growth/Article/KION-with-solid-financial-year-2025-–-strong-order-intake.html"
KN = "KION 2025 年報「Net assets」https://reports.kiongroup.com/2025/ar/management-report/report-on-the-economic-position/financial-position-and-financial-performance/net-assets.html"
KD = "KION 2025 年報與 2026 半年報（搜尋摘要）：工業淨負債 2,685.8、淨確定給付義務 527.5（2025-12-31）／481.7（2026-06-30）、權益 6,123.0；https://reports.kiongroup.com/2025/ar/ 、https://reports.kiongroup.com/2026/ir/2/"
SAS = "stockanalysis.com 各公司 statistics 頁（讀取 2026-11-23；頁面報價日見各列）"
YCL = "Yahoo Finance 日線（physical_ai/data/closes2_20261109_20261120.csv）"
BUND = "德國 10 年期公債殖利率：2026-11 初至中約 3.25%–3.4%（Trading Economics 新聞標題彙整，日期標示不一致，取 3.30%）https://tradingeconomics.com/germany/government-bond-yield"
WHL = "physical_ai/lines/warehouse/dist/20261116_倉儲物流單位經濟_v0.1.xlsx「下游取數」頁"
PROF = "physical_ai/map/profiles/KGX.md"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"

ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. FY2026 錨點：公司指引（2026-10-14 下修；2026-10-29 重申）")
inp("ias26", "公司指引：IAS 營收 FY2026（33.5–34.5 億歐元中點）", "EUR B", 3.35, 3.400, 3.45, "Interested-party", Q3R + "：IAS 營收指引 3,350–3,450 百萬歐元", "前三季 2.519（+16.9%）；隱含第 4 季 0.88 ≈ 第 3 季 0.890", CO)
inp("its26", "公司指引：ITS 營收 FY2026（81–83 億歐元中點）", "EUR B", 8.10, 8.200, 8.30, "Interested-party", Q3R + "：ITS 營收指引 8,100–8,300 百萬歐元（7 月為 8,200–8,500）", "前三季 6.034（−0.7%）", CO)
inp("cor_rev", "公司本部與內部抵銷營收（FY2026 起維持）", "EUR B", -0.06, -0.040, -0.02, "Derived", "＝FY2025 11.297 − 8.272 − 3.071 ＝ −0.046；9M FY2026 8.522 − 6.034 − 2.519 ＝ −0.031（" + KS + "；" + Q3R + "）", "金額小，區間 Assumed")
inp("ias_e26", "公司指引：IAS 調整後 EBIT FY2026（2.30–2.50 億歐元中點）", "EUR B", 0.230, 0.240, 0.250, "Interested-party", Q3R + "：IAS 調整後 EBIT 指引 230–250 百萬歐元", "前三季 0.1722（率 6.8%）", CO)
inp("its_e26", "公司指引：ITS 調整後 EBIT FY2026（6.90–7.40 億歐元中點）", "EUR B", 0.690, 0.715, 0.740, "Interested-party", Q3R + "：ITS 調整後 EBIT 指引 690–740 百萬歐元（7 月為 765–835）", "前三季 0.5287（率 8.8%）", CO)
inp("cor_e", "公司本部調整後 EBIT（＝集團指引中點 8.45 − IAS − ITS；FY2026 起維持）", "EUR B", -0.130, "=0.845-ias_e26-its_e26", -0.090, "Derived", Q3R + "：集團調整後 EBIT 指引 815–875 百萬歐元", "前三季 0.6325 − 0.1722 − 0.5287 ＝ −0.0684；區間 Assumed")
inp("adj26", "調整後 EBIT → EBIT 的調整項 FY2026（購買價格分攤攤銷＋非經常項目）", "EUR B", 0.080, 0.120, 0.160, "Assumed", "FY2025 調整後 788.6 − EBIT 500.9 ＝ 287.7（含非經常項目 201.1）；FY2024 139.4（" + KF + "；" + KS + "）", "FY2026 前三季淨利 309.1 反推調整項約 0.08–0.12（Derived 粗估）；未取得 9M 報告 EBIT")
inp("adjx", "調整項 FY2027 起（購買價格分攤攤銷為主）", "EUR B", 0.060, 0.090, 0.120, "Assumed", "FY2024 調整項 0.139 含非經常 0.028（Derived）→ 攤銷約 0.11；假設逐年下降", "")
inp("fcf26", "公司指引：自由現金流 FY2026（4.0–5.0 億歐元中點；股權口徑，利息後）", "EUR B", 0.400, 0.450, 0.500, "Interested-party", Q3R + "：自由現金流指引 400–500 百萬歐元（7 月為 420–540）", "前三季 0.1211（去年同期 0.3928）", CO)

sec("B. 命題模組：IAS（Dematic 倉儲自動化）營收＝市場成長 × 相對表現；ITS（堆高機與服務）循環")
for s, v, lo, hi, tag, note in (("A", 0.12, None, None, "Assumed", "樂觀：訂單回補，新簽專案加速（FY2025 SCS 訂單 +39.5%、營收 +4.4%）"),
                                ("B", 0.06, 0.00, 0.12, "Analogy", "類比：FY2026 前三季 IAS 訂單 25.75 億（−12.4%）仍略高於營收 25.19（訂單營收比約 1.02）；FY2026 營收 +10.7%（指引）後減速"),
                                ("C", -0.04, None, None, "Assumed", "保守：訂單下滑傳導到營收")):
    inp(f"g27i_{s}", f"情境 {s}：IAS 營收年增 FY2027", "%", lo, v, hi, tag, note, "區間見情境 A／B／C 三列" if s != "B" else "類比對象：前三季訂單營收比（Interested-party 數字，推論 Assumed）")
for s, ref, fml, lo, hi in (("A", "高／樂觀欄 E8、E11", "=(19.60365/10.92945)^(1/6)-1", None, None),
                            ("B", "基準欄 D8、D11", "=(16.94452/10.69434)^(1/6)-1", 0.053, 0.102),
                            ("C", "低／悲觀欄 C8、C11", "=(14.27712/10.45241)^(1/6)-1", None, None)):
    inp(f"go_{s}", f"情境 {s}：市場年增，用於 IAS FY2028 起（讀線模型 AS/RS 市場 2026→2032 年複合）", "%/年", lo, fml, hi, "Derived",
        f"{WHL} WH_ASRS_MKT_2026／WH_ASRS_MKT_2032（{ref}）", "MarketsandMarkets 口徑（含製造、汽車），與 Dematic 系統整合口徑不同；只取成長率")
inp("ias_sp", "IAS 相對市場成長的差（FY2028 起；正＝增加市占）", "pt/年", -0.03, 0.0, 0.03, "Assumed", "無證據顯示 Dematic 份額持續增減；FY2026 IAS +10.7% 高於市場 6–8%（Q3R）", "")
inp("ias_tm_A", "情境 A：IAS 調整後 EBIT 率 FY2029 目標（FY2026 起線性收斂）", "%", None, 0.09, None, "Assumed", "FY2024 3.8%、FY2025 6.0%、9M FY2026 6.8%（" + KS + "；" + Q3R + "）", "")
inp("ias_tm_B", "情境 B：IAS 調整後 EBIT 率 FY2029 目標", "%", 0.065, 0.075, 0.09, "Assumed", "同上；FY2026 指引中點 7.1%", "上限＝樂觀情境值")
inp("ias_tm_C", "情境 C：IAS 調整後 EBIT 率 FY2029 目標", "%", None, 0.060, None, "Assumed", "回到 FY2025 水準 6.0%", "")
for s, a, lo, hi in (("A", 0.04, None, None), ("B", 0.02, -0.02, 0.04), ("C", -0.03, None, None)):
    inp(f"g27t_{s}", f"情境 {s}：ITS 營收年增 FY2027", "%", lo, a, hi, "Assumed", "FY2024 −（8.609）、FY2025 8.272（−3.9%）、FY2026 指引 8.2（−0.9%）；公司稱經濟型產品與亞太疲弱（" + Q3R + "）", "ITS 為成熟堆高機循環，無線模型變數可讀")
for s, a, lo, hi in (("A", 0.04, None, None), ("B", 0.03, 0.01, 0.05), ("C", 0.01, None, None)):
    inp(f"g28t_{s}", f"情境 {s}：ITS 營收年增 FY2028 起", "%", lo, a, hi, "Assumed", "名目成長 1–5%（通膨加輕微成長）", "")
inp("its_tm_A", "情境 A：ITS 調整後 EBIT 率 FY2029 目標（FY2026 起線性收斂）", "%", None, 0.105, None, "Assumed", "FY2024 10.7%、FY2025 8.7%、9M FY2026 8.8%（" + KS + "；" + Q3R + "）", "回到 FY2024 水準")
inp("its_tm_B", "情境 B：ITS 調整後 EBIT 率 FY2029 目標", "%", 0.085, 0.090, 0.105, "Assumed", "FY2026 指引中點 8.7%；公司兩位數 EBIT 率目標已延到 2027 年之後（" + PROF + "）", "")
inp("its_tm_C", "情境 C：ITS 調整後 EBIT 率 FY2029 目標", "%", None, 0.075, None, "Assumed", "Q3 FY2026 8.3% 續降", "")

sec("C. 損益表下半段與現金流（歐元；FY2026 起）")
for s, v, lo, hi in (("A", 0.270, None, None), ("B", 0.265, 0.258, 0.272), ("C", 0.255, None, None)):
    inp(f"gm_{s}", f"情境 {s}：毛利率 FY2026 起（營業費用由毛利與營業利益倒推）", "%", lo, v, hi, "Assumed", "FY2023–FY2025 毛利率 FY2024 26.9%、FY2025 26.1%（" + KP + "）", "FY2023 毛利未找到；毛利率只影響營業費用列的拆分，不影響營業利益與估值")
inp("fin", "淨財務費用 FY2026 起（負值）", "EUR B", -0.150, -0.130, -0.110, "Assumed", "FY2025 −0.1474、FY2024 −0.1880（" + KP + "）", "債務下降趨勢；FY2026 前三季未取得")
inp("tax_w", "有效稅率", "%", 0.25, 0.28, 0.32, "Assumed", "FY2025 32.0%（含非經常項目）、FY2024 37.4%（" + KP + "）", "德國企業稅率約 30%；高稅率年與非經常項目有關")
inp("nci", "非控制權益損益 FY2026 起", "EUR B", 0.008, 0.012, 0.016, "Assumed", "FY2025 0.0104、FY2024 0.0089（" + KP + "）", "")
inp("sh_v", "股數（流通普通股；基本＝稀釋）", "百萬股", None, 131.1, None, "Verified", KP + "：131.1 百萬股（FY2024、FY2025 相同；基本與稀釋 EPS 同值）", "")
inp("pay_r", "配息率（占前一年歸屬母公司淨利）FY2027 起", "%", 0.30, 0.35, 0.45, "Assumed", "FY2025 建議股利 0.62 歐元＝81.3 百萬，占淨利約 35%；FY2024 0.82（" + KP + "）", "FY2026 已付股利為 0.0813（Hist）")
inp("fc_da", "折舊攤銷 ÷ 營收（含租賃事業出租資產折舊）", "%", None, "=h_da_25/h_rev_25", None, "Derived", "＝EBITDA − EBIT ＝ 1,199.7 ÷ 11,297.2（" + KP + "）", "")
inp("fc_cx", "資本支出（含租賃事業資產投資）÷ 營收", "%", 0.100, 0.104, 0.110, "Assumed", "未取得年報資本支出；以 FY2026 自由現金流指引中點反推校準（模型 FY2026 股權口徑自由現金流 ≈ 指引區間內；見 Recon）", "折舊攤銷 10.6%；資本支出略低（出租資產投資扣除處分）")
inp("nwc_r", "營運資金 ÷ 營收", "%", 0.12, "=h_nwc_25/h_rev_25", 0.15, "Derived", "＝1,508.5 ÷ 11,297.2 ＝ 13.4%（" + KN + "）", "FY2024 15.5%；增量營運資金 ＝ 比率 × 營收增量")
inp("nci_v", "非控制權益價值（評價日；未取得帳面）", "EUR B", 0.08, 0.12, 0.20, "Assumed", "FY2025 非控制權益損益 10.4 百萬 × 約 10–12 倍", "資產負債表非控制權益金額未找到")

sec("D. 估值輸入（評價日＝FY2026 年末 2026-12-31）")
inp("px", "股價（2026-11-20 收盤，歐元）", "EUR", None, 29.200001, None, "Verified", YCL + "：KGX.DE 2026-11-20 收盤 29.20", "2026-11-18 31.53；2026-09-18 資料商頁面收盤 40.58（" + SAS + "）；全球 AI 與半導體股當週大跌")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("val_date", "評價日（FY2026 年末）", "日期", None, "=DATE(2026,12,31)", None, "Derived", "會計年度為曆年", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, "=(tp_date-val_date)/365", None, "Derived", "", "")
inp("ind_25", "工業淨負債 2025-12-31（含淨確定給付義務與租賃事業淨融資；保守視為債務）", "EUR B", None, 2.6858, None, "Verified", KD, "淨財務負債 0.584（" + KF + "）僅是其中一部分；租賃事業融資與出租資產折舊、資本支出一併入 FCFF 口徑，取工業淨負債較保守")
inp("div26", "FY2026 支付股利（FY2025 股利 0.62 歐元 × 131.1 百萬股）", "EUR B", None, 0.0813, None, "Verified", KP + "：建議股利總額 81.3 百萬歐元", "")
inp("nd_v", "評價日淨負債＝2025 年底工業淨負債 − FY2026 自由現金流（指引）＋ FY2026 股利", "EUR B", None, "=ind_25-fcf26+div26", None, "Derived", "DCF 標準規格第 3 節", "忽略 FY2026 其他現金項目（匯率、購併：Smart Innovation 70% 股權款項未找到）")
inp("rf", "無風險利率（德國 10 年期公債）", "%", None, 0.033, None, "Derived", BUND, "與股價基準日相差 0–2 週，日期不明確（逾期風險）")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("beta", "β（Blume 調整後）", "x", round(0.67 * 2.05 + 0.33 - 0.3, 6), "=0.67*2.05+0.33", round(0.67 * 2.05 + 0.33 + 0.3, 6), "Derived", SAS + "：KION 5 年 β 2.05（頁面資料日 2025-11；未更新）", "Blume 0.67β＋0.33；敏感度 ±0.3（Assumed）")
inp("kd", "稅前負債成本", "%", None, "=rf+0.02", None, "Assumed", "rf＋2pt（投資等級邊緣）", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable：同業前瞻本益比）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺 KION 對同業的歷史相對本益比）；改以下列 ITS 低倍數與 IAS 同業倍數按盈餘占比混合", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("E. 同業前瞻 P/E（資料商頁面報價日各異，以 2026-11-20 收盤等比例換算；KION 自身不納入中位數；Derived）")
PEERS = [("pe_auto", "AutoStore AUTO.OL（立方格 AS/RS）", "=26.20*14.79/14.68", "2026-10-22 頁面 26.20 倍、收盤 14.68 挪威克朗 → 2026-11-20 收盤 14.79"),
         ("pe_daif", "Daifuku 大福 6383.T（物料搬運系統）", "=21.33*4129/5429", "2026-09-11 頁面 21.33 倍、收盤 5,429 日圓 → 2026-11-20 收盤 4,129"),
         ("pe_geek", "極智嘉 Geek+ 2590.HK（AMR 貨到人）", "=56.32*5.405/9.22", "2026-09-04 頁面 56.32 倍、收盤 9.22 港元 → 2026-11-20 收盤 5.405")]
for k, lab, v, note in PEERS:
    inp(k, f"前瞻 P/E：{lab}（換算至 2026-11-20）", "x", None, v, None, "Derived", f"{SAS}：{note}（{YCL}）", "原值 Interested-party（賣方共識 EPS）；P/E 為股權口徑", DV)
inp("pe_med", "同業前瞻 P/E 中位數（3 家，供 IAS 部分）", "x", None, "=MEDIAN(pe_auto,pe_daif,pe_geek)", None, "Derived", "少於 8 家，不做迴歸（第 10 節第 7 條）。報價日未取得、不納入：Symbotic（GAAP 虧損，無前瞻 P/E）、Ocado（虧損）、Jungheinrich／豐田自動織機／Hyster-Yale（收盤檔無資料）", "同業皆為自動化設備或系統商，偏高成長；與 KION 的 ITS 堆高機部分不可比")
inp("pe_its", "ITS（堆高機與服務）適用前瞻 P/E", "x", 6.5, 8.0, 12.0, "Analogy", "類比：KION 自身頁面前瞻 P/E 9.85（2026-09-18，價 40.58）換算 2026-11-20 收盤 29.20 ＝ 7.1 倍（" + SAS + "）；歐洲工業機械常見 10–14 倍", "自我參照；缺獨立的堆高機同業報價（見 pe_med 說明），為承重 Analogy", DV)
inp("cons_pe", "對照：KION 自身前瞻 P/E（換算至 2026-11-20；不進模型）", "x", None, "=9.85*29.2/40.58", None, "Derived", SAS + "：2026-09-18 頁面 9.85 倍、收盤 40.58 歐元 → 29.20", "", DV)

# ---------------- Hist ----------------
# (鍵, 項目, FY2023, FY2024, FY2025)  單位 EUR B
HROWS = [("rev", "合併營收", 11.4337, 11.5032, 11.2972), ("ias", "SCS／IAS 分部營收", None, 2.943, 3.071), ("its", "ITS 分部營收", None, 8.609, 8.272),
         ("cogs", "營業成本", None, 8.4097, 8.3470), ("gp", "毛利", None, 3.0935, 2.9502), ("ebit", "EBIT（報告）", 0.6606, 0.7778, 0.5009),
         ("adj", "調整後 EBIT", 0.7905, 0.9172, 0.7886), ("iase", "SCS／IAS 調整後 EBIT", None, 0.1129, 0.1832), ("itse", "ITS 調整後 EBIT", None, 0.9175, 0.7218),
         ("fin", "淨財務費用", None, -0.1880, -0.1474), ("pbt", "稅前利益", None, 0.5898, 0.3535), ("tax", "所得稅費用（正數）", None, 0.2205, 0.1130),
         ("ni", "淨利（含非控制權益）", 0.3144, 0.3692, 0.2405), ("nci", "非控制權益損益", None, 0.0089, 0.0104), ("nia", "歸屬母公司淨利", None, 0.3603, 0.2301),
         ("eps", "EPS（基本＝稀釋，歐元）", 2.33, 2.75, 1.75), ("da", "折舊攤銷（EBITDA − EBIT）", 1.0530, 1.1392, 1.1997), ("fcf", "自由現金流（公司定義）", 0.7152, 0.7020, 0.7095),
         ("dps", "每股股利（歐元）", 0.70, 0.82, 0.62)]
HSRC = {"FY2023": KM, "FY2024": KP + "；" + KS + "；" + KM, "FY2025": KP + "；" + KS + "；" + KF}
HS = [("nwc_25", "營運資金 2025-12-31（存貨＋應收＋合約資產 − 應付 − 合約負債）", 1.5085, KN), ("nwc_24", "營運資金 2024-12-31", 1.7832, KN),
      ("rev9m", "9M FY2026 營收", 8.522, Q3R), ("adj9m", "9M FY2026 調整後 EBIT", 0.6325, Q3R), ("ias9m", "9M FY2026 IAS 營收", 2.519, Q3R), ("its9m", "9M FY2026 ITS 營收", 6.034, Q3R),
      ("iase9m", "9M FY2026 IAS 調整後 EBIT", 0.1722, Q3R), ("itse9m", "9M FY2026 ITS 調整後 EBIT", 0.5287, Q3R), ("ni9m", "9M FY2026 淨利", 0.3091, Q3R), ("nia9m", "9M FY2026 歸屬母公司淨利", 0.3012, Q3R),
      ("eps9m", "9M FY2026 EPS（歐元）", 2.30, Q3R), ("fcf9m", "9M FY2026 自由現金流", 0.1211, Q3R),
      ("iaso9m", "9M FY2026 IAS 訂單", 2.575, Q3R), ("its_o9m", "9M FY2026 ITS 訂單", 5.897, Q3R), ("pen_26", "淨確定給付義務 2026-06-30", 0.4817, KD)]


def build(path, ov=None, sens=None, rev=None):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 EUR B（十億歐元）、股數百萬股、EPS 與股價歐元"
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
    FAM = [("g27i", "IAS 營收年增 FY2027"), ("go", "市場年增（IAS FY2028 起，讀線模型）"), ("ias_tm", "IAS 調整後 EBIT 率 FY2029 目標"), ("g27t", "ITS 營收年增 FY2027"),
           ("g28t", "ITS 營收年增 FY2028 起"), ("its_tm", "ITS 調整後 EBIT 率 FY2029 目標"), ("gm", "毛利率")]
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
    H["A1"] = "Hist — 歷史財務（EUR B；IFRS）。公式只引用本頁儲存格（定義名稱 h_*）；FY2023 分部、毛利、稅等「未找到」"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "標記", "來源"])
    for i, (k, lab, *vals) in enumerate(HROWS, 3):
        H.cell(i, 1, lab)
        for j, (sfx, v) in enumerate(zip(("23", "24", "25"), vals)):
            if v is None:
                H.cell(i, 2 + j, "未找到"); continue
            H.cell(i, 2 + j, v).font = BLUE; dn(f"h_{k}_{sfx}", f"Hist!${'BCD'[j]}${i}")
        H.cell(i, 5, "Verified"); H.cell(i, 6, "；".join(f"{p}：{s}" for p, s in HSRC.items()))
    i = 3 + len(HROWS)
    H.cell(i, 1, "檢查：毛利 − 營業成本 − 營收（應為 0；FY2024、FY2025）")
    for j, sfx in enumerate(("24", "25")): H.cell(i, 3 + j, f"=ROUND(h_rev_{sfx}-h_cogs_{sfx}-h_gp_{sfx},3)")
    H.cell(i + 1, 1, "檢查：稅前 − 稅 − 淨利（應為 0；FY2024、FY2025）")
    for j, sfx in enumerate(("24", "25")): H.cell(i + 1, 3 + j, f"=ROUND(h_pbt_{sfx}-h_tax_{sfx}-h_ni_{sfx},3)")
    q0 = i + 3
    hdr(H, q0, ["其他歷史與 2026 前三季項目", "值", "", "", "標記", "來源"])
    for j, (k, lab, v, src) in enumerate(HS, q0 + 1):
        H.cell(j, 1, lab); H.cell(j, 2, v).font = BLUE; H.cell(j, 5, "Verified"); H.cell(j, 6, src); dn(f"h_{k}", f"Hist!$B${j}")
    H.column_dimensions["A"].width = 56

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：IAS 與 ITS 分部營收 → 調整後 EBIT → 標準損益表（IFRS）→ 現金流（EUR B；FY2023–FY2025 實績、FY2026＝公司指引中點、FY2027 起模型）"
    PR = {}; rr = 3
    labs = [("yr", "年度（數值，公式用）"), ("ias", "IAS（Dematic 等）營收"), ("its", "ITS（Linde、STILL 等）營收"), ("cor", "公司本部與抵銷營收"), ("rev", "營收（合併）"),
            ("chk", "檢查：分部合計 − 合併營收（應為 0）"), ("gr", "營收年增率"), ("iasm", "IAS 調整後 EBIT 率"), ("itsm", "ITS 調整後 EBIT 率"), ("adj", "調整後 EBIT（分部合計＋本部）"),
            ("adjm", "調整後 EBIT 率"), ("cogs", "營業成本"), ("gp", "毛利"), ("gm", "毛利率"), ("opx", "營業費用（銷管＋研發＋其他；毛利 − 營業利益倒推）"),
            ("adjx", "調整項（購買價格分攤與非經常項目）"), ("op", "營業利益（EBIT，報告口徑）"), ("opm", "營業利益率"), ("fin", "業外：淨財務費用"), ("pbt", "稅前利益"), ("tax", "所得稅"),
            ("ni", "淨利（含非控制權益）"), ("nci", "非控制權益損益"), ("nia", "歸屬母公司淨利"), ("nm", "淨利率"), ("niy", "淨利年增率"), ("sh", "股數（百萬股）"), ("eps", "EPS（歐元）"),
            ("da", "折舊攤銷"), ("cx", "資本支出（含租賃事業資產投資）"), ("nwc", "營運資金（年末）"), ("dwc", "營運資金增加"),
            ("fcff", "FCFF（EBIT 稅後＋折舊 − 資本支出 − 營運資金增加；歷史＝公司定義自由現金流）"), ("intr", "稅後淨利息"), ("div", "股利"), ("fcfad", "自由現金流扣股利（收支關卡用；FCFF＋稅後淨利息 − 股利）")]
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, (k, _) in enumerate(labs)}
        for k, l in labs: P.cell(R[k], 1, l)
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1)
            cell = lambda k, c=c: f"{c}{R[k]}"; prev = lambda k, p=p: f"{p}{R[k]}"
            def w(k, f, c=c): P[f"{c}{R[k]}"] = f
            w("yr", y)
            if y <= 2025:
                sfx = str(y)[2:]
                w("rev", f"=h_rev_{sfx}"); w("ebit_dummy", 0) if False else None
                if y >= 2024:
                    w("ias", f"=h_ias_{sfx}"); w("its", f"=h_its_{sfx}"); w("cor", f"={cell('rev')}-{cell('ias')}-{cell('its')}"); w("chk", f"=ROUND({cell('ias')}+{cell('its')}+{cell('cor')}-{cell('rev')},6)")
                    w("iasm", f"=h_iase_{sfx}/h_ias_{sfx}"); w("itsm", f"=h_itse_{sfx}/h_its_{sfx}")
                    w("cogs", f"=h_cogs_{sfx}"); w("gp", f"=h_gp_{sfx}"); w("gm", f"={cell('gp')}/{cell('rev')}"); w("opx", f"={cell('gp')}-{cell('op')}")
                    w("fin", f"=h_fin_{sfx}"); w("pbt", f"=h_pbt_{sfx}"); w("tax", f"=h_tax_{sfx}"); w("nci", f"=h_nci_{sfx}"); w("nia", f"=h_nia_{sfx}")
                    w("gr", f"={cell('rev')}/{prev('rev')}-1"); w("niy", f"={cell('ni')}/{prev('ni')}-1")
                else:
                    for k in ("ias", "its", "cor", "iasm", "itsm", "cogs", "gp", "gm", "opx", "fin", "pbt", "tax", "nci", "nia"): w(k, "未找到")
                    w("chk", "未找到")
                w("adj", f"=h_adj_{sfx}"); w("adjm", f"={cell('adj')}/{cell('rev')}"); w("adjx", f"={cell('adj')}-{cell('op')}")
                w("op", f"=h_ebit_{sfx}"); w("opm", f"={cell('op')}/{cell('rev')}"); w("ni", f"=h_ni_{sfx}"); w("nm", f"={cell('ni')}/{cell('rev')}")
                w("sh", "=sh_v"); w("eps", f"=h_eps_{sfx}"); w("da", f"=h_da_{sfx}"); w("fcff", f"=h_fcf_{sfx}")
                if y == 2025: w("nwc", "=h_nwc_25")
                if y == 2024: w("nwc", "=h_nwc_24")
                continue
            if y == 2026: w("ias", "=ias26"); w("its", "=its26")
            elif y == 2027: w("ias", f"={prev('ias')}*(1+g27i_{s})"); w("its", f"={prev('its')}*(1+g27t_{s})")
            else: w("ias", f"={prev('ias')}*(1+go_{s}+ias_sp)"); w("its", f"={prev('its')}*(1+g28t_{s})")
            w("cor", "=cor_rev"); w("rev", f"={cell('ias')}+{cell('its')}+{cell('cor')}"); w("chk", f"=ROUND({cell('ias')}+{cell('its')}+{cell('cor')}-{cell('rev')},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            if y == 2026: w("iasm", "=ias_e26/ias26"); w("itsm", "=its_e26/its26")
            else:
                w("iasm", f"=$F{R['iasm']}+(ias_tm_{s}-$F{R['iasm']})*MIN(1,({y}-2026)/3)"); w("itsm", f"=$F{R['itsm']}+(its_tm_{s}-$F{R['itsm']})*MIN(1,({y}-2026)/3)")
            w("adj", f"={cell('ias')}*{cell('iasm')}+{cell('its')}*{cell('itsm')}+cor_e"); w("adjm", f"={cell('adj')}/{cell('rev')}")
            w("adjx", "=adj26" if y == 2026 else "=adjx"); w("op", f"={cell('adj')}-{cell('adjx')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            w("gm", f"=gm_{s}"); w("gp", f"={cell('rev')}*{cell('gm')}"); w("cogs", f"={cell('rev')}-{cell('gp')}"); w("opx", f"={cell('gp')}-{cell('op')}")
            w("fin", "=fin"); w("pbt", f"={cell('op')}+{cell('fin')}"); w("tax", f"=MAX(0,{cell('pbt')})*tax_w"); w("ni", f"={cell('pbt')}-{cell('tax')}")
            w("nci", "=nci"); w("nia", f"={cell('ni')}-{cell('nci')}"); w("nm", f"={cell('ni')}/{cell('rev')}"); w("niy", f"={cell('ni')}/{prev('ni')}-1")
            w("sh", "=sh_v"); w("eps", f"={cell('nia')}/{cell('sh')}*1000")
            w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}"); w("nwc", f"=nwc_r*{cell('rev')}"); w("dwc", f"={cell('nwc')}-{prev('nwc')}")
            w("fcff", f"={cell('op')}-MAX(0,{cell('op')})*tax_w+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("intr", f"={cell('fin')}*(1-tax_w)")
            w("div", "=div26" if y == 2026 else f"=pay_r*{prev('nia')}")
            w("fcfad", f"={cell('fcff')}+{cell('intr')}-{cell('div')}")
        P.cell(R["ias"], 2, "FY2026＝指引中點；FY2027＝年增 g27i；FY2028 起＝市場年增（線模型）＋相對差 ias_sp")
        P.cell(R["its"], 2, "FY2026＝指引中點；其後依 g27t、g28t")
        P.cell(R["iasm"], 2, "FY2026 指引 → FY2029 線性收斂到情境目標")
        P.cell(R["opx"], 2, "毛利率為輸入，營業費用倒推；不影響營業利益")
        P.cell(R["fcff"], 2, "歷史＝公司定義自由現金流（利息後）；預測＝稅後 EBIT＋折舊 − 資本支出 − 營運資金增加")
        for k in ("gr", "iasm", "itsm", "adjm", "gm", "opm", "nm", "niy"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(labs) + 2
    P.column_dimensions["A"].width = 56; P.column_dimensions["B"].width = 34
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "ias", "its", "fcfad"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（EBIT 稅後＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*erp"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（股數 × 現價；EUR B）"; D["B4"] = "=px*sh_v/1000"
    D["A5"] = "負債 D（工業淨負債 2025-12-31）"; D["B5"] = "=ind_25"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "評價日淨負債與非控制權益（EUR B）＝nd_v ＋ nci_v"; D["B10"] = "=nd_v+nci_v"; dn("dcf_nd", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND(((px*sh_v/1000)*(rf+beta*erp)+ind_25*kd*(1-tax_w))/((px*sh_v/1000)+ind_25)+wacc_adj-wacc,9)"
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
                 ("− 評價日淨負債（工業淨負債口徑）", "=-nd_v"), ("− 非控制權益", "=-nci_v"), ("（保留列：無其他調整）", 0),
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
    V["A1"] = "Val — 目標價：DCF（權重 w_dcf）＋ Comparable（混合前瞻本益比 × FY2028 EPS，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}"); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "DCF 頁")
    r_wi = vrow("IAS 占 FY2028 調整後 EBIT 比例（混合倍數權重）", [f"=PL!H{PR[s]['ias']}*PL!H{PR[s]['iasm']}/PL!H{PR[s]['adj']}" for s in SC], "IAS 適用同業 P/E 中位數，其餘適用 ITS 倍數 pe_its")
    r_mx = vrow("混合前瞻 P/E＝權重 × 同業中位數 ＋ (1 − 權重) × pe_its", [f"={L}{r_wi}*pe_med+(1-{L}{r_wi})*pe_its" for L in "BCD"], "KION 一半以上盈餘來自堆高機，不能直接套自動化同業倍數")
    r_cu = vrow("Comparable（未調整）：混合前瞻 P/E × FY2028 EPS", [f"=MAX(0,PL!H{PR[s]['eps']})*{L}{r_mx}" for s, L in zip(SC, "BCD")], "目標時點 2027-11 的 NTM 約 10/12 FY2028＋2/12 FY2029，以 FY2028 代")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（歐元）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    one("現價（In 頁 px）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_upc = one("Comparable 目標價上檔空間", f"=B{r_cmpt}/px-1")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（EUR B；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=TEXT(px*(1+IF(B{r_up}>0.5,9,IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5))))),"#,##0.00")&"／"&IF(B{r_up}<-0.5,"（已是最低級）",TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,-0.5)))),"#,##0.00"))', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    r_liq = one("期間流動性：情境 B 評價日淨負債（EUR B；FY2026 年底）", "=nd_v", "第 10 節第 55 條；KION 有備用信貸與租賃事業再融資，未建模最大單年新增舉債")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    one("期間：模型折現（評價日 2026-12-31 → 目標時點，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 EPS", f'=IF(PL!H{PR["B"]["eps"]}>0,B{r_tp}/PL!H{PR["B"]["eps"]},"N/M")')
    r_pe3 = one("現價 ÷ 情境 B FY2028 EPS（模型）", f'=IF(PL!H{PR["B"]["eps"]}>0,px/PL!H{PR["B"]["eps"]},"N/M")')
    r_pe4 = one("現價 ÷ 情境 B FY2026 EPS（模型）", f'=px/PL!F{PR["B"]["eps"]}')
    r_mc = one("市值（EUR B）", "=px*sh_v/1000")
    r_ev = one("企業價值（市值＋工業淨負債口徑 nd_v＋非控制權益；EUR B）", "=B{0}+nd_v+nci_v".format(r_mc))
    r_evs = one("EV ÷ FY2026 調整後 EBIT（倍；情境 B）", f"=B{r_ev}/PL!F{PR['B']['adj']}")
    for rr_ in (r_up, r_upc, r_a1, r_a2, r_a3): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_wi,):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_mx, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0.00"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "#,##0.00"
    V.column_dimensions["A"].width = 72; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, upc=r_upc, rt=r_rt, rec=r_rec, gate=r_gate, gj=r_gj, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, pd2=r_pd2, pe=r_pe, liq=r_liq, mx=r_mx, wi=r_wi, mc=r_mc, evs=r_evs, pe4=r_pe4)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司指引、前三季實績對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或實績", "差", "說明"])
    B_ = PR["B"]
    rc = [("FY2026 營收 vs 集團指引中點（11.45–11.75）", f"=PL!F{B_['rev']}", "=(11.45+11.75)/2", "模型為 IAS＋ITS 指引中點＋本部抵銷；公司指引（Interested-party）"),
          ("FY2026 調整後 EBIT vs 集團指引中點（0.815–0.875）", f"=PL!F{B_['adj']}", "=(0.815+0.875)/2", "本部＝指引中點倒推，故應相等"),
          ("FY2026 IAS 營收：模型 − 前三季 → 隱含第 4 季", f"=PL!F{B_['ias']}-h_ias9m", "=h_ias9m/3", "對照：前三季平均每季；第 4 季隱含值應在其附近"),
          ("FY2026 EPS（模型）vs 前三季 EPS 2.30", f"=PL!F{B_['eps']}", "=h_eps9m", "全年應高於前三季；第 4 季隱含 EPS 為差額"),
          ("FY2026 股權口徑自由現金流（模型 FCFF＋稅後淨利息）vs 指引中點 0.45", f"=PL!F{B_['fcff']}+PL!F{B_['intr']}", "=fcf26", "資本支出比率校準的結果；落在指引 0.40–0.50 內才合理"),
          ("IAS 訂單營收比（前三季）", "=h_iaso9m/h_ias9m", "", "約 1.02＝g27i 基準的 Analogy 依據"),
          ("FY2025 實績 EPS vs 模型 FY2026 EPS 年增", f"=PL!F{B_['eps']}/h_eps_25-1", "", "FY2026 EPS 年增主要來自非經常項目減少"),
          ("加權目標價 vs 現價", "=tp", "=px", "現價 2026-11-20；收盤檔 2026-11-18 為 31.53"),
          ("市值隱含：EV ÷ FY2026 調整後 EBIT（倍）", f"=Val!B{VR['evs']}", "", "工業淨負債口徑")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f'=IF(ISNUMBER(C{i}),B{i}-C{i},"")'); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_kion.py 網格＋線性內插；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 64; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_kion.py 以 LibreOffice 重算每一案例）；結果為常數快照"
    base_tp = sens["base"]["tp"] if sens else None
    SE["A2"] = "快照基準值（加權，歐元）"; SE["B2"] = base_tp
    SE["C2"] = "本檔即時值（加權，歐元）"; SE["D2"] = "=tp"
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
    SE.cell(sr + 1, 1, "說明：相對倍數係數 rel_k 基準即為 1.00（未調整），故不列案例。無區間的鍵不列案例。")
    sb = sr + 3
    SE.cell(sb, 1, "build_id"); SE.cell(sb, 4, "{BID}")
    SE.cell(sb + 1, 1, "模型狀態"); SE.cell(sb + 1, 4, '=IF(ABS(D2-B2)>0.005,"STALE","CURRENT")')
    SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 34

    # ---------------- 對帳 ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = "對帳 — 本模型使用的線層級變數（線 5 倉儲物流自動化；讀 " + WHL + "；與線模型不一致處寫明，不靜默改口徑）"
    hdr(RE, 2, ["變數", "所屬線", "線模型值（低／基準／高）", "本模型使用值", "單位", "口徑", "標記", "影響本模型哪些儲存格／差異說明"])
    rec = [("AS/RS 市場規模 2026（WH_ASRS_MKT_2026）", "5 倉儲", "10.45／10.69／10.93", "同左（只取 2026→2032 成長率）", "十億美元", "MarketsandMarkets 口徑（含製造、汽車）", "Derived", "In go_A／go_B／go_C 分母；下游取數 C8:E8"),
           ("AS/RS 市場規模 2032（WH_ASRS_MKT_2032）", "5 倉儲", "14.28／16.94／19.60", "同左", "十億美元", "同上", "Derived", "In go_* 分子；下游取數 C11:E11"),
           ("市場年增（由上兩列推得）", "5 倉儲", "5.3%／8.0%／10.2%", "=TEXT(go_C,\"0.0%\")&\"／\"&TEXT(go_B,\"0.0%\")&\"／\"&TEXT(go_A,\"0.0%\")", "%/年", "市場年複合當作 IAS FY2028 起基本成長，另加相對差 ias_sp", "Derived", "PL IAS 營收列；一致。FY2027 用 g27i 自設（訂單減速的 Analogy），與線模型不同，理由見 In g27i_B"),
           ("KION IAS 調整後 EBIT 率（追蹤指標 D10）", "5 倉儲", "6.8%（9M 2026）", "=TEXT(PL!F" + str(PR['B']['iasm']) + ",\"0.0%\")&\"（FY2026 指引）\"", "%", "線模型追蹤頁取 2026 前三季；本模型取全年指引中點", "Verified", "一致（前三季 6.8% vs 全年指引中點 7.1%）"),
           ("IAS 訂單年減（追蹤指標 D10）", "5 倉儲", "−12.4%（9M 2026）", "−12.4%（Hist h_iaso9m）", "%", "同上", "Verified", "一致；用於 g27i_B 的 Analogy"),
           ("AMR 與 AS/RS 每行成本（WH_COST_AMR／ASRS／MANUAL）", "5 倉儲", "0.205／0.197／0.351", "未進公式", "美元/訂單行", "線模型單位經濟", "Derived", "自動化相對人工的成本優勢只作需求背景；未建 Dematic 單位經濟"),
           ("人形每人工當量小時成本（WH_HUMANOID_COST_HR）", "5 倉儲", "155.6／40.0／12.5", "未進公式", "美元/時", "Agility 示意 RaaS", "Derived", "人形進倉經濟性未過關（40 > 35.1）；對 KION 非承重"),
           ("堆高機（ITS）市場", "—", "線模型無此變數", "ITS 全由自設成長率驅動", "—", "線模型只涵蓋自主堆高機於 Interact 行動機器人口徑，未拆分", "Assumed", "ITS 占營收約七成，為承重 Assumed；線模型無可讀變數")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 46; RE.column_dimensions["C"].width = 30; RE.column_dimensions["H"].width = 60

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模或只在 Sens 呈現的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("堆高機（ITS）循環與經濟型產品、亞太疲弱", "是（g27t、its_tm Sens）", "ITS 占營收約七成、調整後 EBIT 約八成；10 月已下修指引", "ITS 訂單、季度調整後 EBIT 率"),
             ("IAS 訂單減速（前三季 −12.4%）", "是（g27i Sens）", "訂單領先營收", "IAS 季度訂單與訂單營收比"),
             ("租賃事業融資與工業淨負債口徑", "部分（nd_v 取工業淨負債）", "租賃事業負債與資產配對；若視為經營性，估值偏保守", "淨財務負債 vs 工業淨負債"),
             ("大股東濰柴動力持股 46.5%（投票權）", "否", "控股股東可影響股利與資本配置；流通股比例較低", "投票權通知、年度股東會"),
             ("購併（Smart Innovation 70% 股權）與自主堆高機", "否", "金額與產品化進度未找到", "公司公告"),
             ("價格基準日與同業倍數換算", "部分", "同業 P/E 為 9–10 月頁面值按收盤比例換算", "新頁面 P/E")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 56; RK.column_dimensions["C"].width = 60

    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([Q3R, KF, KP, KM, KS, KN, KD, SAS + "：AUTO（osl）、6383（tyo）、2590（hkg）、KGX（etr）", YCL, BUND, WHL, PROF], 1):
        SO.cell(i, 1, s_)

    # ---------------- README ----------------
    RD["A1"] = f"KION（KGX.DE）個股模型 {VER}｜build_id {{BID}}"
    RD["A2"] = '="資料截止日 2026-11-24（最新為 2026-10-29 前三季財報）；成品日 2026-11-24；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.00")&" 歐元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：KION 的盈餘七成靠堆高機（ITS）循環，三成來自 Dematic 倉儲自動化（IAS）；股價取決於 IAS 訂單減速是否傳導到營收、IAS 利潤率能否站穩 7% 以上，以及 ITS 利潤率能否回到 9% 以上；市場把它當堆高機公司定價，倉儲自動化的成長只獲得部分估值"
    RD["A4"] = (f'="在 IAS 營收 FY2027 年增 "&TEXT(g27i_B,"0%")&"、之後隨線模型市場年增 "&TEXT(go_B,"0.0%")&"、IAS 與 ITS 調整後 EBIT 率 FY2029 達 "&TEXT(ias_tm_B,"0.0%")&"／"&TEXT(its_tm_B,"0.0%")&"、ITS 適用 P/E "&TEXT(pe_its,"0.0")&" 倍的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 歐元（DCF "&TEXT(Val!B{r_dcft},"0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"0.00")&" 十億歐元）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 歐元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：IAS 季度訂單連兩季年減超過 10% 且營收年增掉到個位數以下；ITS 調整後 EBIT 率跌破 8%；FY2026 全年指引再下修；自由現金流低於 3 億歐元"
    RD["A7"] = "範圍：兩個報告分部（ITS、IAS）加本部抵銷；IAS 營收由線 5 倉儲線模型的 AS/RS 市場年增推導，ITS 為自設成長率；不讀 Siliconomics／Tokenomics／Wattonomics。HTML 展示層見 dist/ 同名 .html"
    RD["A8"] = f'="估值口徑：EV 以工業淨負債（含租賃事業淨融資與退休金義務）"&TEXT(ind_25,"0.00")&" 十億歐元（2025 年底）起算，較保守；市值 "&TEXT(Val!B{r_mc},"0.00")&" 十億歐元，現價為 FY2026 模型 EPS 的 "&TEXT(Val!B{r_pe4},"0.0")&" 倍"'
    wb.save(path)
    return VR, PR, DR


LO = "file://" + os.path.join(os.environ.get("HOME") or tempfile.gettempdir(), "lo_kion")
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
                subprocess.run(["pkill", "-f", "lo_kion"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
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
    d = tempfile.mkdtemp(prefix="kion_"); paths = []; VRs = []
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

def fingerprint(): return hashlib.sha1(json.dumps([r for r in ROWS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]
def rate(u): return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"


def write_readme(path, VR, PR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]
    g = lambda c: V[c].value
    f2 = lambda x: f"{x:,.2f}"; pc = lambda x: f"{x*100:+.1f}%"
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:8]
    def fv(x): return x if isinstance(x, str) else f"{x:.4g}"
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {fv(r['val'])} | {r['base']} | {r['note']} |" for r in revs)
    def yr(s, k, fmt="{:,.2f}"):
        return " | ".join(fmt.format(v) if isinstance((v := P[f"{COL[y]}{PR[s][k]}"].value), (int, float)) else "—" for y in YRS)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f2(g(chr(66+i)+str(VR['dcfr'])))} | {f2(g(chr(66+i)+str(VR['cu'])))} | {f2(g(chr(66+i)+str(VR['tps'])))} |" for i, s in enumerate(SC))
    Iv = {wb['In'][f'A{r}'].value: wb['In'][f'E{r}'].value for r in range(1, wb['In'].max_row + 1)}
    rts = sorted(set(c['rt'] for c in sens['cases']))
    txt = f"""# physical_ai/kion — KION（KGX.DE，法蘭克福）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

線 5 倉儲物流自動化。本 README 的數字由 `build_kion.py` 從重算後的成品讀出產生，不手寫。

**本版沒有做的部分：** FY2023 分部、毛利、稅與非控制權益未找到（只取到合併營收、EBIT、淨利與 EPS）；FY2026 前三季報告 EBIT、資本支出、非控制權益帳面金額、Smart Innovation 收購金額未找到；FY2026 的第 4 季以公司指引中點代；Jungheinrich、豐田自動織機、Hyster-Yale 等堆高機同業的報價日 2026-11-20 收盤不在收盤檔，未納入。ITS 倍數（pe_its）為自我參照的 Analogy，是承重輸入。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A8'].value}

## 估值（評價日 2026-12-31；目標時點約 12 個月）
| 情境 | DCF（推到目標時點） | Comparable（混合前瞻 P/E） | 情境目標價 |
|---|---|---|---|
{scn}

- 加權目標價 **{f2(tp)} 歐元**（DCF {f2(dcf)}、Comparable {f2(cmp_)}），現價 {f2(Iv['px'])}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")}（情境 B FY2026–FY2030 自由現金流扣股利 {g(f"C{VR['gate']}"):.2f} 十億歐元）→ 投資建議 **{g(f"B{VR['rec']}")}**。
- 方法：DCF 標準版 FCFF 50%＋Comparable 50%；Comparable 倍數＝IAS 盈餘占比 × 同業（AutoStore、Daifuku、極智嘉）前瞻 P/E 中位數 {Iv['pe_med']:.1f} 倍＋ITS 占比 × ITS 倍數 {Iv['pe_its']:.1f} 倍（Analogy），再乘情境 FY2028 EPS；相對倍數係數 1.00。
- 現價相當於情境 B FY2026 EPS 的 {g(f"B{VR['pe4']}"):.1f} 倍；全球 AI 與半導體股 2026-11-16 至 11-20 當週大跌，收盤價下跌不是資料錯誤。

## 營收與損益（情境 B，EUR B）
| 列 | {" | ".join(f"FY{y}" for y in YRS)} |
|---|{"---|" * len(YRS)}
| IAS 營收 | {yr('B','ias')} |
| ITS 營收 | {yr('B','its')} |
| 營收（合併） | {yr('B','rev')} |
| 調整後 EBIT | {yr('B','adj','{:,.3f}')} |
| 營業利益（EBIT） | {yr('B','op','{:,.3f}')} |
| 歸屬母公司淨利 | {yr('B','nia','{:,.3f}')} |
| EPS（歐元） | {yr('B','eps')} |
| FCFF | {yr('B','fcff','{:,.3f}')} |

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| ITS 與 IAS 調整後 EBIT 率 FY2029 目標（its_tm、ias_tm） | 季度分部調整後 EBIT 率；公司全年指引 |
| IAS 營收年增 FY2027（g27i） | IAS 季度訂單與訂單營收比 |
| ITS 適用 P/E（pe_its） | 歐洲堆高機同業前瞻 P/E（本版缺報價） |
| 資本支出比率（fc_cx）、工業淨負債口徑 | 年報資本支出、淨財務負債 vs 工業淨負債 |
| β | 資料商 5 年 β 2.05（2025-11 頁面） |

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
讀 `physical_ai/lines/warehouse/dist/20261116_倉儲物流單位經濟_v0.1.xlsx`「下游取數」：WH_ASRS_MKT_2026／2032 的低／基準／高 → IAS FY2028 起市場年增 go_C／go_B／go_A（{Iv['go_C']*100:.1f}%／{Iv['go_B']*100:.1f}%／{Iv['go_A']*100:.1f}%）。與線模型不一致：FY2027 IAS 年增用自設的訂單減速 Analogy（基準 {Iv['g27i_B']*100:.0f}%），不用市場年增；ITS（堆高機）線模型沒有可讀變數，全為自設。

## 限制與未找到資料
- FY2023 分部、毛利、稅、非控制權益；資本支出、FY2026 前三季報告 EBIT、非控制權益帳面金額、Smart Innovation 收購金額：未找到。
- 同業前瞻 P/E 取自資料商頁面（報價日 2026-09-04～10-22），以 2026-11-20 收盤等比例換算（Derived）；只有 3 家，不做迴歸；同業皆為自動化高成長公司，對 ITS 不可比，故用混合倍數。報價日未取得、不納入：Symbotic、Ocado（虧損）、堆高機同業（收盤檔無資料）。
- β 取自資料商 2025-11 頁面（未更新）；rf 為德國 10 年期公債 2026-11 初至中的區間彙整（日期標示不一致）。
- 工業淨負債（2,685.8）與淨確定給付義務取自搜尋摘要轉述年報，未逐頁回原表核對。
- 未做：未參與製作的代理獨立核對（Project 規則第 10 條）；tools/model_governance/check_display.py（無 display_check.json）。

## 檔案
- `build_kion.py`：建置＋整本重算 Sens／反解＋產生本 README 與 `segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-24）：初版；現價 {f2(Iv['px'])}（2026-11-20 收盤）；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


def main():
    if "--raw" in sys.argv:
        ov = json.loads(os.environ.get("KION_OVERRIDE", "{}"))
        build(sys.argv[sys.argv.index("--raw") + 1], ov); return
    px0 = in_val("px", "base")
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
    d = tempfile.mkdtemp(prefix="kionb_"); p = os.path.join(d, "base.xlsx"); VR, _, DR = build(p, {})
    base = read_vals(recalc_many([p], os.path.join(d, "o"))[0], VR, DR); shutil.rmtree(d, ignore_errors=True)
    print("基準", base)
    cases = []
    for (k, side, v, lab), r in zip(specs, res[1:1 + len(specs)]):
        cases.append(dict(drv=lab + ("（低）" if side == "lo" else "（高）"), set=f"{k}＝{v:.4g}", key=k, **r))
    for s, r in zip(SC, res[1 + len(specs):]):
        cases.append(dict(drv=f"只看情境 {s}", set=f"w_{s}＝1，其餘 0", key=f"w_{s}", **r))
    up_ = base["alt_tp"] / px0 - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rate(up_)))
    up0 = base["up"]; tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    targets = [("目標價＝現價（上檔 0%）", 0.0)] + ([(f"升一級邊界（上檔 {upper:+.0%}）", upper)] if upper is not None else []) + ([(f"降一級邊界（上檔 {lower:+.0%}）", lower)] if lower is not None else [])
    revs = []
    for key, grid, b0, note in (("its_tm_B", [0.03, 0.05, 0.07, 0.09, 0.105, 0.12], 0.09, "情境 B ITS 調整後 EBIT 率 FY2029 目標（整本重算；情境 A／C 不動）"),
                                ("ias_tm_B", [0.0, 0.03, 0.05, 0.075, 0.10, 0.13], 0.075, "情境 B IAS 調整後 EBIT 率 FY2029 目標（整本重算；情境 A／C 不動）"),
                                ("pe_its", [3, 5, 6.5, 8, 10, 12, 16], 8.0, "ITS 適用前瞻 P/E（整本重算）")):
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
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="kionf_"); raw = os.path.join(tmp, NAME)
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
    SRC = KS + "；" + Q3R
    json.dump(dict(company="kion", ticker="KGX.DE", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=v, unit="EUR B", source=KS) for n, l, v in
                                      (("ITS", "ITS（Linde、STILL 等）營收", 8.272), ("SCS／IAS", "IAS（Dematic 等）營收", 3.071))],
                   sublines=[], check_labels=["檢查：分部合計 − 合併營收（應為 0）"],
                   note="兩個報告分部（ITS、SCS 自 2026 年改稱 IAS）；本部與抵銷列為差額（FY2025 −0.046）。"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)
    shutil.rmtree(tmp, ignore_errors=True)

if __name__ == "__main__":
    main()

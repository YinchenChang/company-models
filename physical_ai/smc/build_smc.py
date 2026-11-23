# SMC（SMC株式会社，6273.T）個股模型 v0.1 — AI實體應用 Project（線 6 共用零組件：氣動元件、電動致動器、夾爪）
# 用法：python3 physical_ai/smc/build_smc.py            → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、segments.json、sens.json，並跑核對
#       python3 physical_ai/smc/build_smc.py --raw OUT  → 只寫一本（讀環境變數 SMC_OVERRIDE 的 JSON 覆寫 In 頁基準值）
# 會計年度：3 月底決算（日本會計準則）。年度標籤採「起始年」：FY2026＝2026-04～2027-03（2027年3月期，本期公司指引年），
#   FY2025＝2026年3月期（最近實績），與本專案哈默納科模型同一口徑。金額：十億日圓（¥B）；股數：百萬股；EPS、股價：日圓。
# 結構：README、In、Hist、PL（三情境：命題驅動→地區子列→標準損益表→現金流）、DCF、Val、Recon（含實體 AI 曝險節）、Sens、對帳、Risk、Sources。
# 公司只有單一報告分部（自動制御機器事業），地區別營收（地域ごとの情報）為分部之下的子列。
# 不讀 Siliconomics／Tokenomics／Wattonomics（實體 AI 線）。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261123"
NAME = f"{DATE}_SMC個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032
HISTY = (2023, 2024, 2025)

# ---------------- 出處 ----------------
TS2Q = "公司 2027年3月期 第2四半期（中間期）決算短信〔日本基準〕(連結)（2026-11-13；JPX 適時開示）https://www2.jpx.co.jp/disc/62730/140120261112575619.pdf"
BR2Q = "公司 2027年3月期 第2四半期 決算説明資料（2026-11-13）https://www2.jpx.co.jp/disc/62730/140120261112575686.pdf"
KBT = "株探 2026-11-13 決算速報（中間期實績、通期修正、Q2 單季）https://s.kabutan.jp/stocks/6273/news/k202611130114/"
FWM = "FourWeekMBA 2026-11-13 報導（中間期與修正指引、庫藏股）https://fourweekmba.com/ai-smc-raises-profit-outlook-34-says-capex-is-centred-on-ai/"
EDB = "edinetdb.jp（EDINET 有價證券報告書轉載；FY2025 MD&A）https://edinetdb.jp/company/E01673/earnings/FY2026"
SAI = "stockanalysis.com 年度／季度損益（資料商；讀取 2026-11-23）https://stockanalysis.com/quote/tyo/6273/financials/income-statement/"
SAB = "stockanalysis.com 資產負債表（資料商）https://stockanalysis.com/quote/tyo/6273/financials/balance-sheet/"
SAC = "stockanalysis.com 現金流量表（資料商）https://stockanalysis.com/quote/tyo/6273/financials/cash-flow-statement/"
SAF = "stockanalysis.com forecast 頁（2026-09-02 頁面；目標價 2026-08-13 更新；S&P Global）https://stockanalysis.com/quote/tyo/6273/forecast/"
SHA = "the-shashi.com 地域別売上（有價證券報告書「地域ごとの情報」轉載；百萬日圓）https://the-shashi.com/tse/6273/financials/region/"
YCL = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
JGB10 = "日本財務省 國債金利情報 10 年：2026-11-02 2.817%（與本專案 Nabtesco、哈默納科、Fanuc 模型同一取值，跨模型一致）"
FAN = "physical_ai/fanuc/dist/20261123_發那科個股模型_v0.1.1.xlsx"
HDS = "physical_ai/harmonic_drive/dist/20261123_哈默納科個股模型_v0.1.1.xlsx"
NAB = "physical_ai/nabtesco/dist/20261123_Nabtesco個股模型_v0.1.1.xlsx"
LC = "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx"
LIR = "physical_ai/lines/industrial_robot/dist/20261116_工業機器人單位經濟_v0.1.xlsx"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整（立場：中立彙整，來源為公司申報或賣方；未逐項核對法定原文）"
SELL = "賣方分析師（立場：券商研究與交易關係）"

# ---------------- In 列 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源1, 來源2／說明, 來源誘因)
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. FY2026（2026-04～2027-03）錨點：公司指引（2026-11-13 上修；Interested-party）")
inp("g26_rev", "公司指引：FY2026 營收", "¥B", None, 1073.0, None, "Interested-party", TS2Q + "：売上高 1,073,000 百萬日圓（前次 2026-05-14 1,000,000）", "中間期實績 553.636 → 下半年隱含 519.364（較上半 −6.2%，Derived）；" + KBT, CO)
inp("g26_op", "公司指引：FY2026 營業利益", "¥B", None, 294.0, None, "Interested-party", TS2Q + "：営業利益 294,000（前次 219,000）", "下半年隱含 136.054（上半 157.946）", CO)
inp("g26_ord", "公司指引：FY2026 經常利益", "¥B", None, 325.0, None, "Interested-party", TS2Q + "：経常利益 325,000（前次 239,000）", "", CO)
inp("g26_ni", "公司指引：FY2026 歸屬母公司淨利", "¥B", None, 238.0, None, "Interested-party", TS2Q + "：親会社株主に帰属する当期純利益 238,000（前次 170,000）", "", CO)
inp("g26_eps", "公司指引：FY2026 EPS", "¥", None, 3789.21, None, "Interested-party", TS2Q + "：1株当たり当期純利益 3,789.21", "株探以期末股數換算為 3,818.1（" + KBT + "）；本模型用公司原文", CO)
inp("g26_gm", "公司指引：FY2026 毛利率", "%", None, 0.475, None, "Interested-party", BR2Q + "：通期修正 売上総利益率 47.5%（前次 44.9%；FY2025 實績 45.3%）", "上半實績 47.9%（265.292 ÷ 553.636）", CO)
inp("g26_sga", "公司指引：FY2026 營業費用率（販管費 ÷ 營收）", "%", None, "=g26_gm-g26_op/g26_rev", None, "Derived", "＝毛利率 47.5% − 營業利益率（294 ÷ 1,073 ＝ 27.4%）＝ 20.1%", BR2Q + "：販管費率 20.1%（公司原文，Interested-party，一致）")
inp("g26_da", "公司計畫：FY2026 折舊", "¥B", None, 62.3, None, "Interested-party", BR2Q + "：減価償却費 623 億（5 月計畫；11 月未改列）", "上半實績 26.361（" + TS2Q + "）", CO)
inp("g26_cx", "公司計畫：FY2026 設備投資", "¥B", None, 100.0, None, "Interested-party", BR2Q + "：設備投資計畫 1,000 億（上半完成 328 億）", "FY2025 實績 1,502 億", CO)
inp("cx_27", "公司說法：FY2027 設備投資", "¥B", 80.0, 95.0, 100.0, "Interested-party", BR2Q + "：「来期の設備投資も1,000億円弱となる見通し」", "基準 95（「弱」取 −5%）；低 80／高 100 為 Assumed", CO)
inp("g26_dps", "公司計畫：FY2026 每股股利（中間 500＋期末 500）", "¥", None, 1000, None, "Interested-party", TS2Q + "：年間配當予想 1,000（不變）", "", CO)
inp("ord_idx", "受注指數 2026-10（FY2025＝100；只作對照，不進模型）", "指數", None, 151, None, "Interested-party", BR2Q + "：連結受注 26/1Q 135、26/2Q 140、10 月 151", "Recon 對照用；只作對照，不進模型", CO)
inp("g_anc", "FY2026 實績相對公司指引（營收同比例；毛利率、費用率不變）", "%", -0.03, 0.0, 0.05, "Assumed", "基準＝公司指引（上半已實績，偏離只來自下半）", "高 +5%：10 月受注指數 151（FY2025＝100）高於上半營收年化指數 131，下半指引較上半低 6%；低 −3%：中國電動車、二次電池投資放緩（" + BR2Q + "）")

sec("B. 命題驅動：工廠自動化設備投資循環 → SMC 銷售量 × 價格（單一分部；地區子列以 FY2025 結構分配）")
for s, v, lo, hi, tag, note in (
        ("A", 0.15, None, None, "Assumed", "樂觀：受注指數 2026Q1 135 → Q2 140 → 10 月 151（FY2025＝100），半導體／電機 205（" + BR2Q + "）；10 月指數 ÷ FY2026 指引營收指數 127（Derived）≈ 1.19 倍延續到 FY2027"),
        ("B", 0.06, 0.0, 0.12, "Assumed", "基準：受注領先營收，但 FY2026 營收已含匯率貢獻（上半 +11.5%）；以量 +6% 承接訂單、不假設匯率續升"),
        ("C", -0.10, None, None, "Assumed", "保守：設備投資循環回落；類比 FY2023（2024年3月期）營收 −5.8%、FY2019（2020年3月期）−8.8%（" + SHA + " 合計，Derived）")):
    inp(f"vg27_{s}", f"情境 {s}：FY2027 銷售量年增", "%", lo, v, hi, tag, note, "區間見情境 A／B／C 三列（A 15%、C −10%）" if s == "B" else "")
for s, v, lo, hi in (("A", 0.08, None, None), ("B", 0.05, 0.02, 0.08), ("C", 0.02, None, None)):
    inp(f"vgm_{s}", f"情境 {s}：FY2028–FY2029 銷售量年增", "%/年", lo, v, hi, "Analogy" if s == "B" else "Assumed",
        "類比：SMC 營收 FY2013（2014年3月期）395.3 → FY2025 842.5，12 年年複合 6.5%（含匯率與價格；" + SHA + "，Derived）" if s == "B" else "A／C 以 B ± 3pt 設定",
        "量的成長，價格另列；工業機器人裝機線模型 2026→2032 年複合 6.6%（對帳頁），機器人只占 SMC 需求一小部分，不直接採用")
for s, v, lo, hi in (("A", 0.06, None, None), ("B", 0.04, 0.02, 0.06), ("C", 0.02, None, None)):
    inp(f"vglt_{s}", f"情境 {s}：FY2030–FY2032 銷售量年增", "%/年", lo, v, hi, "Assumed", "長期：氣動元件市場成長（未取得第三方市場規模；以公司長期營收成長減價格與匯率約 4% 為中心）", "")
for s, v, lo, hi in (("A", 0.015, None, None), ("B", 0.005, 0.0, 0.015), ("C", 0.0, None, None)):
    inp(f"prc_{s}", f"情境 {s}：每年價格變化（FY2027 起）", "%/年", lo, v, hi, "Assumed",
        BR2Q + "：上半營收年增 38.3% 中價格效果 +0.6%（日本 +19 億、海外 +4 億）", "全球份額 36%（公司自稱，Interested-party）支撐價格，但無明確新漲價計畫")
for s, v, lo, hi in (("A", 0.0, None, None), ("B", -0.01, -0.025, 0.0), ("C", -0.025, None, None)):
    inp(f"gadj_{s}", f"情境 {s}：FY2027 起毛利率相對 FY2026 指引調整", "pt", lo, v, hi, "Assumed",
        "毛利率史：FY2023 46.7%、FY2024 45.8%、FY2025 45.3%、FY2026 指引 47.5%（Hist 頁；Derived）", "B −1pt：指引含人民幣升值 16% 與高稼動，回到近三年均值 45.9% 與指引之間；C −2.5pt：回到 FY2025 水準")
inp("e_sga", "營業費用對營收的彈性（營業費用年增 ÷ 營收年增）", "x", 0.5, 0.8, 1.2, "Assumed",
    "上半：營業費用 +18.9% ÷ 營收 +38.3% ＝ 0.49；FY2025：+10.4% ÷ +6.4% ＝ 1.6；FY2024：+3.6% ÷ +2.0% ＝ 1.8（Hist、" + BR2Q + "；Derived）", "短期固定成本攤薄、長期接近 1；基準 0.8")

sec("C. 損益表下半段")
inp("nonop", "業外損益（營業外＋特別損益）FY2027 起", "¥B", 12.0, 22.0, 32.0, "Assumed", "上半：受取利息 11.022、匯兌損失 −3.082、特別損益 +1.566（" + TS2Q + "）；匯兌損益不預測（設 0）", "以利息收入年約 22 為基準；FY2025 業外 46.4（含匯兌利益 19.7）")
inp("xo_26", "FY2026 特別損益（上半實績；下半假設 0）", "¥B", None, "=2.339-0.773", None, "Derived", TS2Q + "：特別利益 2,339、特別損失 773", "")
inp("tax_w", "有效稅率 FY2027 起", "%", 0.26, 0.28, 0.30, "Derived", "FY2025：69.677 ÷ 236.988 ＝ 29.4%；上半：45.780 ÷ 174.772 ＝ 26.2%（" + SAI + "、" + TS2Q + "）", "取兩者之間 28%；低／高 Assumed")
inp("nci_r", "非支配股東損益 ÷ 稅後淨利", "%", None, 0.0, None, "Derived", "FY2025 非支配股東損益 0.009（" + SAI + "）÷ 稅後淨利 167.3 ≈ 0；非支配股東持分 0.242（" + TS2Q + "）", "")
inp("pay_r", "FY2027 起配息率（股利 ÷ 歸屬母公司淨利）", "%", 0.26, 0.30, 0.40, "Assumed", "FY2025 實績 37.9%（DPS 1,000 ÷ EPS 2,640）；FY2026 指引 26.4%（1,000 ÷ 3,789）", "公司方針「安定的な配当の継続」＋機動的自己株式取得（" + BR2Q + "）；只影響收支關卡，不進 DCF 與目標價")

sec("D. 現金流")
inp("fc_da", "折舊 ÷ 營收（FY2027 起）", "%", 0.05, "=g26_da/g26_rev", 0.065, "Derived", "＝FY2026 計畫 62.3 ÷ 指引營收 1,073 ＝ 5.8%", "FY2025 44.846 ÷ 842.541 ＝ 5.3%（" + SAC + "）；低／高 Assumed")
inp("fc_cx", "資本支出 ÷ 營收（FY2028 起）", "%", 0.05, 0.07, 0.10, "Assumed", "FY2023 13.2%、FY2024 13.5%、FY2025 18.1%（投資高峰，" + SAC + "）；FY2026 計畫 9.3%、FY2027 約 95（公司）", "長期假設回到 7%（高於折舊 5.8%，支撐量的成長）")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.50, 0.70, 0.85, "Derived", "（應收＋存貨 − 應付）÷ 營收：FY2024 (206.733＋489.660 − 46.898) ÷ 792.108 ＝ 82.0%；FY2025 84.1%（" + SAB + "）；2026-09-30 767.8 ÷ 上半年化 1,107 ＝ 69.3%（" + TS2Q + "）", "SMC 刻意持有高存貨（短交期政策）；取最近一期 70%")
inp("stub_wc", "下半年（2026-10～2027-03）營運資金增加", "¥B", -30.0, 0.0, 30.0, "Assumed", "上半營運資金增加 57.6（Hist，Derived）；下半指引營收較上半低 6%，應收可能回收", "基準 0；低 −30（回收）／高 +30")

sec("E. 估值輸入（評價日 2027-03-31；資產負債表 2026-09-30）")
inp("px", "股價（2026-11-20 收盤）", "¥", None, 55300, None, "Verified", YCL + "：6273.T 2026-11-20 收盤 55,300", "未還原股利；中間決算（11-13 14:00）當日收盤 62,320，之後回落")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("sh_out", "流通股數（2026-09-30；發行 63,869,359 − 自己株式 1,534,226）", "百萬股", None, "=63.869359-1.534226", None, "Derived", TS2Q + "：期末發行済株式数、自己株式数（原文 Verified）", "無選擇權、RSU 揭露 → 不另加")
inp("bb_rem", "評價日前尚待執行的庫藏股買回金額", "¥B", 0.0, 50.0, None, "Assumed", "公司 2026-11-13 決議：上限 100 萬股（1.6%）或 500 億円，期間 2026-11-25～2027-03-24（" + FWM + "；" + BR2Q + "）", "基準＝以現價全數執行（Assumed）；低案＝不再買回")
inp("sh_v", "評價日股數", "百萬股", None, "=sh_out-bb_rem/px*1000", None, "Derived", "＝流通股數 − 尚待買回金額 ÷ 股價", "DCF 標準規格第 3 節")
inp("cash_v", "現金及び預金（2026-09-30）", "¥B", None, 689.169, None, "Verified", TS2Q + "：689,169", "")
inp("sec_v", "有価証券（2026-09-30）", "¥B", None, 46.678, None, "Verified", TS2Q + "：46,678", "")
inp("inv_v", "投資有価証券（2026-09-30）", "¥B", None, 76.636, None, "Verified", TS2Q + "：76,636（2026-03-31 121,714）", "")
inp("inv_k", "投資有価証券計入比例", "x", 0.7, 1.0, None, "Assumed", "基準按帳面全額；低案 70%（處分時的稅負與折價）", "")
inp("debt_v", "借款（2026-09-30）", "¥B", None, 0.0, None, "Verified", TS2Q + "：短期借入金「－」（2026-03-31 5,092）", "租賃負債未單獨揭露（未取得）；金額小，不計")
inp("mi_v", "非支配股東持分（2026-09-30）", "¥B", None, 0.242, None, "Verified", TS2Q + "：242", "")
inp("dps_int", "評價日前支付的中間股利（2026-12 支付；基準日 2026-09-30 已過）", "¥", None, 500, None, "Interested-party", TS2Q + "：中間配當 500（決議）", "股價日已除權息，該現金不屬於現價持有人 → 自股權價值扣除", CO)
inp("dps_ye", "FY2026 期末股利（基準日 2027-03-31；2027-06 支付）", "¥", None, 500, None, "Interested-party", TS2Q + "：期末配當予想 500", "評價日股權價值為除息後 → 扣除", CO)
inp("dps_tgt", "評價日到目標時點的每股股利（FY2027 中間，基準日 2027-09-30）", "¥", None, 500, None, "Assumed", "假設與 FY2026 中間股利相同", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, "=(tp_date-DATE(2027,3,31))/365", None, "Derived", "2027-03-31 → 目標時點", "")
inp("rf", "無風險利率（日本 10 年期國債）", "%", None, 0.02817, None, "Verified", JGB10, "")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "日本國家風險溢酬", "%", None, 0.0091, None, "Verified", "Damodaran 國家風險溢酬（2026-01-05）：日本 0.91%（與本專案日本個股同值）", "")
inp("beta", "β（兩來源 Blume 調整後中位數）", "x", "=E{r}-0.3", "=MEDIAN(0.67*0.83+0.33,0.67*0.97+0.33)", "=E{r}+0.3", "Derived",
    "SMC 5 年 β 0.83；同業 8 家 5 年 β 中位數 0.97（Fanuc 0.98、安川 1.24、THK 0.76、哈默納科 0.96、Nabtesco 0.93、Nidec 1.07、Knorr-Bremse 1.11、ASSA ABLOY 0.82；stockanalysis.com，讀取 2026-11-16，取自 " + NAB + " In 頁 beta 列）", "Blume 0.67β＋0.33；敏感度 ±0.3（Assumed）")
inp("kd", "稅前負債成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt；評價日無有息負債，WACC 權重為 0", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable 本益比）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺 SMC 對同業 10 年前瞻 P/E 比值歷史）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("F. 同業前瞻 P/E（2026-11-20 收盤口徑；取自本專案其他個股模型的同業表，跨模型一致）")
PEERS = [("pe_keyence", "Keyence 6861（感測、FA 零組件）", 29.8353471749489, FAN + " Val!D41", "traders.co.jp PER（收盤 ÷ 公司本期 EPS 予想）"),
         ("pe_fanuc", "Fanuc 6954（機器人、CNC）", 20.4505989144675, FAN + " Val!D46", "traders.co.jp PER（公司本期 EPS 予想）"),
         ("pe_nab", "Nabtesco 6268（減速機）", 14.6736010072748, FAN + " Val!D42", "traders.co.jp PER（公司本期 EPS 予想）"),
         ("pe_okuma", "Okuma 6103（工具機）", 10.5625158831004, FAN + " Val!D43", "traders.co.jp PER（公司本期 EPS 予想）"),
         ("pe_dmg", "DMG Mori 6141（工具機）", 23.7848260447036, FAN + " Val!D44", "traders.co.jp PER（公司本期 EPS 予想）"),
         ("pe_yask", "安川電機 6506（伺服、機器人）", 15.6738990825688, HDS + " Comp!F7", "stockanalysis.com forward P/E（賣方共識 NTM EPS）"),
         ("pe_thk", "THK 6481（線性傳動）", 10.6780788096795, HDS + " Comp!F8", "stockanalysis.com forward P/E（賣方共識 NTM EPS）"),
         ("pe_nidec", "Nidec 6594（馬達）", 11.5719333333333, HDS + " Comp!F10", "stockanalysis.com forward P/E（賣方共識 NTM EPS）")]
for k, lab, v, cell, basis in PEERS:
    inp(k, f"前瞻 P/E：{lab}", "x", None, v, None, "Derived", f"取自 {cell}（2026-11-20 收盤；原報價日 P/E × 新收盤 ÷ 舊收盤，分母 EPS 不變）",
        f"分母口徑：{basis}；會計口徑：日本基準或 IFRS（P/E 為股權口徑，不受租賃處理影響）")
inp("pe_med", "同業前瞻 P/E 中位數（8 家）", "x", None, "=MEDIAN(pe_keyence,pe_fanuc,pe_nab,pe_okuma,pe_dmg,pe_yask,pe_thk,pe_nidec)", None, "Derived",
    "8 家，等於約 8 家門檻，不做迴歸（第 10 節第 7 條）；兩種分母口徑（公司予想 5 家、賣方共識 3 家）混用，列為限制", "未取得：CKD 6407、Parker-Hannifin、Festo（未上市）等氣動同業報價（csv 無此檔）")
inp("pe_self", "SMC 現行前瞻 P/E（現價 ÷ FY2026 公司指引 EPS；只作對照，不進模型）", "x", None, "=px/g26_eps", None, "Derived", "＝55,300 ÷ 3,789.21；" + FAN + " Val!D40 traders.co.jp 口徑 14.51（以 3,812 為分母），只作對照，不進模型", "")
inp("cons_tp", "賣方平均目標價（13 位；只作對照，不進模型）", "¥", None, 88154, None, "Interested-party", SAF + "：平均 88,154、中位數 88,000、低 78,500、高 100,000；評等 Buy", "中間決算前（2026-08-13 更新），只作歷史對照，不進模型", SELL)
inp("cons_rev26", "共識 FY2026 營收（只作對照，不進模型）", "¥B", None, 1070.0, None, "Interested-party", SAF + "：平均 1.07 兆（13 位，兆日圓取一位小數）", "中間決算前，只作歷史對照，不進模型", SELL)

sec("G. 實體 AI 曝險（只作對照，不進模型；Recon 頁「實體 AI 曝險」節使用，不影響目標價）")
for cy, col in ((2026, "D"), (2030, "H"), (2031, "I")):
    inp(f"pne_{cy}", f"機器人（工業＋協作）拉動的氣動元件市場 CY{cy}", "百萬美元", f"={{lo{cy}}}", f"={{b{cy}}}", f"={{hi{cy}}}", "Derived",
        f"線 6 {LC} 情境!{col}44／{col}45／{col}46（低／基準／高）；每台內容 傳統 500、協作 300 美元（輸入!D57、E57，Assumed）", "五線歸因口徑：人形、AMR、車用的氣動內容基準為 0；只作對照，不進模型")
inp("pai_sh", "SMC 在機器人用氣動元件的份額", "%", 0.25, 0.36, 0.45, "Assumed", BR2Q + "：SMC 全球氣動份額 36%（FY2025，公司自稱，Interested-party）", "機器人末端與周邊氣動的份額未取得，以全球份額代、給區間；只作對照，不進模型")
inp("fx_usd", "美元兌日圓（換算用）", "¥", None, 158.0, None, "Interested-party", TS2Q + "：通期想定レート USD 158.00", "只作對照，不進模型", CO)
inp("ir_25", "全球工業機器人裝機 CY2025（含協作）", "台", None, 603307, None, "Verified", f"線 3 {LIR} 下游取數 IR.inst.global.2025.base（E12；IFR）", "只作對照，不進模型")
inp("ir_26", "全球工業機器人裝機 CY2026（基準）", "台", None, 655010.4099, None, "Derived", f"線 3 {LIR} 下游取數 IR.inst.global.2026.base（E13）", "只作對照，不進模型")
inp("ir_32", "全球工業機器人裝機 CY2032（基準）", "台", None, 959983.818818984, None, "Derived", f"線 3 {LIR} 下游取數 IR.inst.global.2032.base（E19）", "只作對照，不進模型")

LINE_CACHE = {}
def line_vals():
    """讀線 6 成品的氣動元件市場（情境!D44:I46）；值寫成常數，來源儲存格寫在 In 頁說明。"""
    if LINE_CACHE: return LINE_CACHE
    wb = openpyxl.load_workbook(os.path.join(ROOT, LC), data_only=True); ws = wb["情境"]
    assert ws["A45"].value == "PNE" and ws["C45"].value == "基準", "線 6 情境頁 PNE 列位置改變"
    for cy, col in ((2026, "D"), (2030, "H"), (2031, "I")):
        LINE_CACHE[f"lo{cy}"] = ws[f"{col}44"].value; LINE_CACHE[f"b{cy}"] = ws[f"{col}45"].value; LINE_CACHE[f"hi{cy}"] = ws[f"{col}46"].value
    w2 = openpyxl.load_workbook(os.path.join(ROOT, LIR), data_only=True)["下游取數"]
    ids = {w2.cell(r, 1).value: w2.cell(r, 5).value for r in range(5, w2.max_row + 1)}
    LINE_CACHE["ir"] = {k: ids[k] for k in ("IR.inst.global.2025.base", "IR.inst.global.2026.base", "IR.inst.global.2032.base")}
    return LINE_CACHE

# ---------------- Hist ----------------
# FY2023＝2024年3月期、FY2024＝2025年3月期、FY2025＝2026年3月期
HR = [("h_rev", "營收（合併）", (776.873, 792.108, 842.541), "Verified", EDB + "；" + SAI + "（兩路徑一致：FY2025 842,541）"),
      ("h_cogs", "營業成本（売上原価）", (413.796, 429.069, 461.089), "Interested-party", SAI + "；FY2025 毛利率 45.3% 與 " + BR2Q + " 一致"),
      ("h_op", "營業利益", (196.226, 190.246, 190.558), "Verified", EDB + "；" + KBT + "（FY2025 190,558）"),
      ("h_pbt", "稅前利益（税金等調整前当期純利益）", (252.653, 211.067, 236.989), "Verified", EDB + "（FY2025 236,989）；" + SAI),
      ("h_tax", "所得稅（法人税等）", (74.300, 54.608, 69.677), "Interested-party", SAI),
      ("h_nci", "非支配股東損益", (0.032, 0.115, 0.009), "Interested-party", SAI),
      ("h_ni", "歸屬母公司淨利", (178.321, 156.344, 167.302), "Verified", EDB + "；" + KBT),
      ("h_eps", "EPS（日圓）", (2766.94, 2444.63, 2640.04), "Verified", EDB + "（FY2025 2,640.04）"),
      ("h_da", "折舊攤銷", (33.355, 34.308, 44.846), "Interested-party", SAC),
      ("h_cx", "資本支出", (102.347, 106.739, 152.701), "Interested-party", SAC),
      ("h_dwc", "營運資金增加（現金流量表；正＝占用）", (96.674, -18.926, 42.215), "Interested-party", SAC + "（Change in working capital 變號）"),
      ("h_ocf", "營業現金流", (98.207, 196.656, 188.917), "Interested-party", SAC),
      ("h_div", "股利支付", (58.063, 64.048, 63.535), "Interested-party", SAC),
      ("h_bb", "庫藏股買回", (30.628, 25.018, 30.018), "Interested-party", SAC)]
REG = [("jp", "日本"), ("cn", "中國"), ("as", "亞洲（中國以外）"), ("eu", "歐洲"), ("us", "美國"), ("ot", "其他地區")]
REGV = {"jp": (159.172, 158.116, 158.823), "cn": (195.218, 208.690, 234.939), "as": (136.726, 151.612, 158.955),
        "eu": (149.487, 144.414, 162.250), "us": (93.200, 88.937, 83.708), "ot": (43.068, 40.336, 43.864)}
H1 = [("h1_rev", "營收", 553.636), ("h1_cogs", "營業成本", 288.344), ("h1_sga", "營業費用（販管費）", 107.346), ("h1_op", "營業利益", 157.946),
      ("h1_ord", "經常利益", 173.206), ("h1_pbt", "稅前利益", 174.772), ("h1_tax", "所得稅", 45.780), ("h1_ni", "歸屬母公司淨利", 128.979),
      ("h1_da", "折舊", 26.361), ("h1_cx", "有形固定資產取得", 29.721), ("h1_ocf", "營業現金流", 110.449), ("h1_div", "股利支付", 31.576), ("h1_bb", "自己株式取得", 56.656)]
WC = [("wc_rec", "受取手形及び売掛金", 235.325, 289.015), ("wc_inv", "棚卸資產（商品製品＋仕掛品＋原材料）", "=198.419+33.289+274.989", "=188.958+36.754+292.408"),
      ("wc_ap", "支払手形及び買掛金", 31.822, 39.336)]


def build(path, ov=None, sens=None, rev=None):
    ov = ov or {}
    LV = line_vals()
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 ¥B、股數百萬股、EPS 與股價日圓。年度＝起始年（FY2026＝2026-04～2027-03）"
    IN["A2"] = "標記：Verified（附出處）／Interested-party（寫明是誰、誘因欄）／Analogy（寫明類比對象）／Assumed（給區間）／Derived（寫明換算）；查不到寫「未取得」"
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
            if isinstance(v, str):
                v = v.replace("{r}", str(r))
                for kk, vv in LV.items():
                    if not isinstance(vv, dict): v = v.replace("{" + kk + "}", repr(vv))
            c = IN.cell(r, col, v)
            if col == 5: c.font = BLUE
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2); IN.cell(r, 10, inc)
        if k in ("px_date", "tp_date"): IN.cell(r, 5).number_format = "yyyy-mm-dd"
        dn(k, f"In!$E${r}"); KR[k] = r; r += 1
    in_last = r - 1
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("vg27", "FY2027 銷售量年增"), ("vgm", "FY2028–FY2029 銷售量年增"), ("vglt", "FY2030–FY2032 銷售量年增"),
           ("prc", "每年價格變化"), ("gadj", "FY2027 起毛利率調整")]
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
    H["A1"] = "Hist — 歷史財務（¥B；日本會計準則；年度＝起始年：FY2023＝2024年3月期 … FY2025＝2026年3月期）。公式只引用本頁儲存格（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "標記", "來源"])
    i = 3
    for k, lab, vals, tag, src in HR:
        H.cell(i, 1, lab)
        for j, v in enumerate(vals): H.cell(i, 2 + j, v).font = BLUE
        H.cell(i, 5, tag); H.cell(i, 6, src)
        for j, y in enumerate(HISTY): dn(f"{k}_{y % 100}", f"Hist!${'BCD'[j]}${i}")
        i += 1
    H.cell(i, 1, "營業費用（販管費＋其他營業費用；＝毛利 − 營業利益，Derived）")
    for j, y in enumerate(HISTY):
        yy = y % 100; H.cell(i, 2 + j, f"=h_rev_{yy}-h_cogs_{yy}-h_op_{yy}"); dn(f"h_sga_{yy}", f"Hist!${'BCD'[j]}${i}")
    H.cell(i, 5, "Derived"); H.cell(i, 6, "資料商 SG&A 列較本列小約 0.5（其他營業費用另列）；公司 FY2025 販管費率 22.7%（" + BR2Q + "）與本列一致"); i += 1
    H.cell(i, 1, "營運資金 ÷ 營收（年底；Derived；FY2023 年底明細未取得）"); H.cell(i, 3, "=(206.733+489.660-46.898)/h_rev_24"); H.cell(i, 4, "=(234.366+506.697-31.822)/h_rev_25")
    H.cell(i, 6, SAB + "：應收、存貨、應付"); i += 2
    # 地區別
    hdr(H, i, ["地區別營收（地域ごとの情報；顧客所在地）", "FY2023", "FY2024", "FY2025", "標記", "來源"]); i += 1
    r0 = i
    for k, lab in REG:
        H.cell(i, 1, lab)
        for j, v in enumerate(REGV[k]): H.cell(i, 2 + j, v).font = BLUE
        H.cell(i, 5, "Interested-party"); H.cell(i, 6, SHA + "（資料商轉載）")
        for j, y in enumerate(HISTY): dn(f"h_{k}_{y % 100}", f"Hist!${'BCD'[j]}${i}")
        i += 1
    H.cell(i, 1, "檢查：地區合計 − 合併營收（四捨五入差 < 0.01）")
    for j, L in enumerate("BCD"): H.cell(i, 2 + j, f"=ROUND(SUM({L}{r0}:{L}{i-1})-{L}3,2)")
    i += 1
    H.cell(i, 1, "FY2025 地區占比（FY2026 起地區子列分配用；Derived）")
    hdr(H, i + 1, ["地區", "占比"])
    for n, (k, lab) in enumerate(REG):
        H.cell(i + 2 + n, 1, lab); H.cell(i + 2 + n, 2, f"=h_{k}_25/SUM(D{r0}:D{r0+5})"); dn(f"mix_{k}", f"Hist!$B${i+2+n}")
    i += 2 + len(REG) + 1
    # 上半實績
    hdr(H, i, ["FY2026 上半（2026-04～09）實績", "值", "", "", "標記", "來源"]); i += 1
    for k, lab, v in H1:
        H.cell(i, 1, lab); H.cell(i, 2, v).font = BLUE; H.cell(i, 5, "Verified"); H.cell(i, 6, TS2Q); dn(k, f"Hist!$B${i}"); i += 1
    H.cell(i, 1, "毛利率（上半；Derived）"); H.cell(i, 2, "=(h1_rev-h1_cogs)/h1_rev"); i += 1
    hdr(H, i + 1, ["營運資金（決算短信資產負債表）", "2026-03-31", "2026-09-30", "", "標記", "來源"]); i += 2
    w0 = i
    for k, lab, a, b in WC:
        H.cell(i, 1, lab); H.cell(i, 2, a).font = BLUE; H.cell(i, 3, b).font = BLUE; H.cell(i, 5, "Verified"); H.cell(i, 6, TS2Q); i += 1
    H.cell(i, 1, "營運資金（應收＋存貨 − 應付；Derived）"); H.cell(i, 2, f"=B{w0}+B{w0+1}-B{w0+2}"); H.cell(i, 3, f"=C{w0}+C{w0+1}-C{w0+2}"); i += 1
    H.cell(i, 1, "上半營運資金增加（Derived）"); H.cell(i, 2, f"=C{i-1}-B{i-1}"); dn("h1_dwc", f"Hist!$B${i}"); i += 1
    H.cell(i + 1, 1, "受注指數（FY2025＝100；公司簡報，Interested-party）：2026Q1 135、Q2 140、10 月 151；10 月產業別：半導體／電機 205、汽車 111、工具機 134、食品 109、醫療 125（" + BR2Q + "）")
    H.cell(i + 2, 1, "上半營收年增 38.3＝量 +26.0%＋價格 +0.6%＋匯率 +11.5%；量的地區別：日本 +21%、北美 +12%、歐洲 +10%、中國台灣香港 +42%、其他亞洲 +37%（" + BR2Q + "）")
    H.column_dimensions["A"].width = 52; H.column_dimensions["F"].width = 60

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題驅動（銷售量 × 價格）→ 營收（單一報告分部，地區子列）→ 標準損益表（日本會計準則）→ 現金流（¥B；FY2023–FY2025 實績、FY2026 公司指引 ×(1＋g_anc)、FY2027 起模型）"
    PR = {}
    rr = 3
    KEYS = ["vol", "prc", "jp", "cn", "as", "eu", "us", "ot", "rev", "chk", "gr", "cogs", "gp", "gm", "sga", "sgar", "op", "opm", "nonop", "pbt", "tax", "nit", "nci", "ni",
            "nim", "nig", "sh", "eps", "da", "cx", "dwc", "fcff", "div", "fcfad"]
    LAB = ["銷售量年增（命題驅動）", "價格年變化", "地區子列：日本", "地區子列：中國", "地區子列：亞洲（中國以外）", "地區子列：歐洲", "地區子列：美國", "地區子列：其他地區",
           "營收（合併；自動制御機器事業，單一報告分部）", "檢查：地區子列合計 − 合併營收", "營收年增率",
           "營業成本", "毛利", "毛利率", "營業費用（販管費）", "營業費用率", "營業利益", "營業利益率", "業外損益（營業外＋特別損益）", "稅前利益", "所得稅", "稅後淨利（含非支配）",
           "非支配股東損益", "淨利（歸屬母公司）", "淨利率", "淨利年增率", "平均股數（百萬股）", "EPS（日圓）",
           "折舊攤銷", "資本支出", "營運資金增加", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "股利", "自由現金流扣股利（收支關卡用）"]
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, k in enumerate(KEYS)}
        for i, l in enumerate(LAB): P.cell(b + i, 1, l)
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1); cell = lambda k: f"{c}{R[k]}"; prev = lambda k: f"{p}{R[k]}"
            def w(k, f): P[cell(k)] = f
            if y in HISTY:
                yy = y % 100
                for k, _ in REG: w(k, f"=h_{k}_{yy}")
                w("rev", f"=h_rev_{yy}"); w("chk", f"=ROUND(SUM({c}{R['jp']}:{c}{R['ot']})-{c}{R['rev']},2)")
                if y > 2023: w("gr", f"={cell('rev')}/{prev('rev')}-1"); w("nig", f"={cell('ni')}/{prev('ni')}-1")
                w("cogs", f"=h_cogs_{yy}"); w("gp", f"={cell('rev')}-{cell('cogs')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
                w("sga", f"=h_sga_{yy}"); w("sgar", f"={cell('sga')}/{cell('rev')}"); w("op", f"={cell('gp')}-{cell('sga')}"); w("opm", f"={cell('op')}/{cell('rev')}")
                w("pbt", f"=h_pbt_{yy}"); w("nonop", f"={cell('pbt')}-{cell('op')}"); w("tax", f"=h_tax_{yy}"); w("nit", f"={cell('pbt')}-{cell('tax')}")
                w("nci", f"=h_nci_{yy}"); w("ni", f"=h_ni_{yy}"); w("nim", f"={cell('ni')}/{cell('rev')}"); w("eps", f"=h_eps_{yy}"); w("sh", f"={cell('ni')}/{cell('eps')}*1000")
                w("da", f"=h_da_{yy}"); w("cx", f"=h_cx_{yy}"); w("dwc", f"=h_dwc_{yy}")
                w("fcff", f"={cell('op')}*(1-{cell('tax')}/{cell('pbt')})+{cell('da')}-{cell('cx')}-{cell('dwc')}")
                w("div", f"=h_div_{yy}"); w("fcfad", f"={cell('fcff')}-{cell('div')}")
                continue
            if y == 2026:
                w("rev", "=g26_rev*(1+g_anc)"); w("gp", f"={cell('rev')}*g26_gm"); w("sga", f"={cell('rev')}*g26_sga")
                w("nonop", "=g26_ord-g26_op+xo_26"); w("tax", f"={cell('pbt')}*(1-g26_ni/(g26_ord+xo_26))")
                w("sh", "=g26_ni/g26_eps*1000")
                w("da", "=g26_da"); w("cx", "=g26_cx"); w("dwc", "=h1_dwc+stub_wc"); w("div", f"=g26_dps*{cell('sh')}/1000")
            else:
                vg = f"vg27_{s}" if y == 2027 else (f"vgm_{s}" if y <= 2029 else f"vglt_{s}")
                w("vol", f"={vg}"); w("prc", f"=prc_{s}")
                w("rev", f"={prev('rev')}*(1+{cell('vol')})*(1+{cell('prc')})")
                w("gp", f"={cell('rev')}*(g26_gm+gadj_{s})")
                w("sga", f"={prev('sga')}*(1+e_sga*({cell('rev')}/{prev('rev')}-1))")
                w("nonop", "=nonop"); w("tax", f"={cell('pbt')}*tax_w"); w("sh", "=sh_v")
                w("da", f"=fc_da*{cell('rev')}"); w("cx", "=cx_27" if y == 2027 else f"=fc_cx*{cell('rev')}")
                w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})"); w("div", f"=pay_r*{cell('ni')}")
            for k, _ in REG: w(k, f"={cell('rev')}*mix_{k}")
            w("chk", f"=ROUND(SUM({c}{R['jp']}:{c}{R['ot']})-{c}{R['rev']},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            w("cogs", f"={cell('rev')}-{cell('gp')}"); w("gm", f"={cell('gp')}/{cell('rev')}"); w("sgar", f"={cell('sga')}/{cell('rev')}")
            w("op", f"={cell('gp')}-{cell('sga')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            w("pbt", f"={cell('op')}+{cell('nonop')}"); w("nit", f"={cell('pbt')}-{cell('tax')}")
            w("nci", f"={cell('nit')}*nci_r"); w("ni", f"={cell('nit')}-{cell('nci')}"); w("nim", f"={cell('ni')}/{cell('rev')}"); w("nig", f"={cell('ni')}/{prev('ni')}-1")
            w("eps", f"={cell('ni')}/{cell('sh')}*1000")
            taxr = "(1-g26_ni/(g26_ord+xo_26))" if y == 2026 else "tax_w"
            w("fcff", f"={cell('op')}*(1-{taxr})+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("fcfad", f"={cell('fcff')}-{cell('div')}")
        P.cell(R["vol"], 2, "命題驅動：FY2027 vg27、FY2028–29 vgm、FY2030–32 vglt（情境專屬）")
        P.cell(R["rev"], 2, "FY2026＝公司指引 ×(1＋g_anc)；之後＝前一年 ×(1＋量)×(1＋價格)")
        P.cell(R["jp"], 2, "FY2026 起＝營收 × FY2025 地區占比（Assumed：結構固定）")
        P.cell(R["gp"], 2, "FY2026＝指引毛利率；FY2027 起＝指引毛利率＋gadj")
        P.cell(R["sga"], 2, "FY2027 起＝前一年 ×(1＋e_sga × 營收年增)")
        P.cell(R["nonop"], 2, "FY2023–25＝稅前 − 營業利益；FY2026＝指引經常 − 營業＋上半特別損益")
        P.cell(R["tax"], 2, "FY2026＝指引隱含稅率；FY2027 起 tax_w")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 ×(1 − 稅率)；FY2026 營運資金＝上半實績＋下半假設")
        P.cell(R["div"], 2, "FY2023–25 實績支付；FY2026＝DPS 1,000；之後 pay_r × 淨利")
        for k in ("gr", "gm", "sgar", "opm", "nim", "nig", "vol", "prc"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        for k in KEYS:
            if k not in ("gr", "gm", "sgar", "opm", "nim", "nig", "vol", "prc", "eps"):
                for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "#,##0.0"
        for y in YRS: P[f"{COL[y]}{R['eps']}"].number_format = "#,##0"
        PR[s] = R; rr = b + len(LAB) + 2
    P.column_dimensions["A"].width = 46; P.column_dimensions["B"].width = 36

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2027-03-31（FY2026 年底），明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（¥B）"; D["B4"] = "=px*sh_v/1000"
    D["A5"] = "負債 D（借款）"; D["B5"] = "=debt_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金與投資（資產負債表 2026-09-30）＝現金＋有價證券 − 借款 − 非支配 ＋ 投資有價證券 × inv_k"; D["B10"] = "=cash_v+sec_v-debt_v-mi_v+inv_k*inv_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC"; D["B11"] = "=ROUND(((px*sh_v/1000)*(rf+beta*(erp+crp))+debt_v*kd*(1-tax_w))/((px*sh_v/1000)+debt_v)+wacc_adj-wacc,9)"
    for rr_ in range(3, 11): D[f"B{rr_}"].number_format = "0.00%" if rr_ in (3, 6, 7, 8, 9) else "#,##0.0"
    DR = {}; rr = 13
    for s in SC:
        R = PR[s]; F = COL[2026]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}" for y in range(2027, 2033)]); b = rr + 2
        rows = ["EBIT（營業利益）", "NOPAT", "＋折舊攤銷", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF", "折現年數（期中）", "折現因子", "FCFF 現值"]
        for i, l in enumerate(rows): D.cell(b + i, 1, l)
        for j, y in enumerate(range(2027, 2033)):
            c = "BCDEFG"[j]; pc = COL[y]
            D[f"{c}{b}"] = f"=PL!{pc}{R['op']}"; D[f"{c}{b+1}"] = f"={c}{b}*(1-tax_w)"
            D[f"{c}{b+2}"] = f"=PL!{pc}{R['da']}"; D[f"{c}{b+3}"] = f"=-PL!{pc}{R['cx']}"; D[f"{c}{b+4}"] = f"=-PL!{pc}{R['dwc']}"
            D[f"{c}{b+5}"] = f"=PL!{pc}{R['fcff']}"; D[f"{c}{b+6}"] = f"=ROUND(SUM({c}{b+1}:{c}{b+4})-{c}{b+5},9)"
            D[f"{c}{b+7}"] = j + 0.5; D[f"{c}{b+8}"] = f"=1/(1+wacc)^{c}{b+7}"; D[f"{c}{b+9}"] = f"={c}{b+5}*{c}{b+8}"
        t = b + 11
        items = [("明確期現值合計", f"=SUM(B{b+9}:G{b+9})"),
                 ("NOPAT FY2033＝NOPAT FY2032 ×(1＋g)", f"=G{b+1}*(1+g_tv)"),
                 ("終值（FY2032 年底）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^6"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金與投資（2026-09-30）", "=dcf_nc"),
                 ("＋下半年 FCFF（2026-10～2027-03：指引全年 − 上半實績）",
                  f"=(PL!{F}{R['op']}-h1_op)*(1-(1-g26_ni/(g26_ord+xo_26)))+(PL!{F}{R['da']}-h1_da)-(PL!{F}{R['cx']}-h1_cx)-stub_wc"),
                 ("− 尚待買回金額", "=-bb_rem"),
                 ("− 評價日前支付的中間股利與 FY2026 期末股利", "=-(dps_int*sh_out+dps_ye*sh_v)/1000"),
                 ("股權價值（評價日）", f"=B{t+4}+B{t+5}+B{t+6}+B{t+7}+B{t+8}"),
                 ("每股價值（評價日）", f"=MAX(0,B{t+9})/sh_v*1000"),
                 ("DCF 情境價值推到目標時點（×(1＋ke)^t − 期間股利）", f"=B{t+10}*(1+ke)^t_tgt-dps_tgt"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+12}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=B{t+2}/G{b+1}"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到目標時點）",
                  f"=((B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6)+B{t+5}+B{t+6}+B{t+7}+B{t+8})/sh_v*1000*(1+ke)^t_tgt-dps_tgt")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 11, ps=t + 10, ev=t + 4, tvs=t + 12, alt=t + 15, chk=b + 6)
        dn(f"dcf_{s}", f"DCF!$B${t+11}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 64

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價：DCF（權重 w_dcf）＋ Comparable 本益比（1 − w_dcf；同業 8 家前瞻 P/E 中位數 × 目標時點起算 NTM EPS）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}")
        V.cell(vr, 6, note); vr += 1; return vr - 1
    G, Hc = COL[2027], COL[2028]
    r_ntm = vrow("NTM EPS（目標時點 2027-11 起算：FY2027 × 1/3＋FY2028 × 2/3；日圓）", [f"=PL!{G}{PR[s]['eps']}/3+PL!{Hc}{PR[s]['eps']}*2/3" for s in SC], "分母：本模型情境 EPS（非共識）")
    r_dcf = vrow("DCF 情境價值推到目標時點", [f"=dcf_{s}" for s in SC], "DCF 頁")
    r_cu = vrow("Comparable（未調整）：同業前瞻 P/E 中位數 × NTM EPS", [f"=pe_med*{L}{r_ntm}" for L in "BCD"], "同業 8 家中位數（In 頁 F 節）")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（日圓）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    r_px = one("現價（In 頁 px）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（¥B；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!{COL[2026]}{PR[s]['fcfad']}:PL!{COL[2030]}{PR[s]['fcfad']})")
    V.cell(vr, 6, "口徑：FCFF − 股利（未扣庫藏股買回；無有息負債）"); vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=TEXT(px*(1+IF(B{r_up}>0.5,9,IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5))))),"#,##0")&"／"&TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,IF(B{r_up}>=-0.5,-0.5,-9))))),"#,##0")', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_cmpup = one("Comparable 目標價上檔空間（只作對照，評等不依此）", f"=B{r_cmpt}/px-1")
    r_pd1 = one("期間：模型折現（評價日 2027-03-31 → 目標時點，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B NTM EPS", f"=B{r_tp}/C{r_ntm}", "分母：模型情境 B NTM EPS")
    r_pe2 = one("現行前瞻 P/E（現價 ÷ FY2026 指引 EPS）÷ 同業中位數", "=pe_self/pe_med", "只作對照；顯示用")
    r_tvsh = one("DCF 終值占 EV（情境 B）", f"=DCF!B{DR['B']['tvs']}")
    for rr_ in (r_up, r_a1, r_a2, r_a3, r_cmpup, r_tvsh): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_ntm, r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "#,##0"
    V.column_dimensions["A"].width = 70; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, ca=r_ca, ntm=r_ntm, pd2=r_pd2, pe=r_pe, pe2=r_pe2, gj=r_gj, wk=r_wk, cmpup=r_cmpup, tvsh=r_tvsh, vg=r_vg)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司指引、市場共識對照；實體 AI 曝險（只作對照，不進目標價）；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或共識", "差", "說明"])
    F = COL[2026]; B_ = PR["B"]
    rc = [("FY2026 營收", f"=PL!{F}{B_['rev']}", "=g26_rev", "公司指引（Interested-party）；基準 g_anc＝0，與指引相同"),
          ("FY2026 營業利益", f"=PL!{F}{B_['op']}", "=g26_op", "公司指引"),
          ("FY2026 歸屬母公司淨利", f"=PL!{F}{B_['ni']}", "=g26_ni", "公司指引（模型含上半特別損益，稅率由指引反推）"),
          ("FY2026 EPS", f"=PL!{F}{B_['eps']}", "=g26_eps", "公司指引"),
          ("FY2026 營收 vs 中間決算前共識", f"=PL!{F}{B_['rev']}", "=cons_rev26", "資料商共識（2026-09-02 頁面，財報前，只作歷史對照）"),
          ("10 月受注指數 ÷ FY2026 營收指數（FY2025＝100）", "=ord_idx/100", f"=PL!{F}{B_['rev']}/PL!{COL[2025]}{B_['rev']}", "受注 151（公司簡報，Interested-party）÷ 指引營收指數；口徑：受注為單月、營收為全年，只看方向"),
          ("加權目標價 vs 賣方平均目標價", "=tp", "=cons_tp", "賣方 13 位平均（2026-08-13，財報前；Interested-party）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    # 實體 AI 曝險
    RC.cell(i, 1, "實體 AI 曝險（只作對照，不進目標價；線 6 氣動元件市場 × SMC 份額 × 匯率）").font = BOLD; i += 1
    hdr(RC, i, ["項目", "低", "基準", "高", "說明"]); i += 1
    pai0 = i
    RC.cell(i, 1, "機器人拉動的氣動市場 FY2026（百萬美元；9/12 × CY2026＋3/12 × CY2027 以 CY2026 代）")
    for j, col in enumerate("DEF"): RC.cell(i, 2 + j, f"=In!{col}{KR['pne_2026']}")
    RC.cell(i, 5, "線 6 情境!D44:D46（CY2026）；會計年度錯開 3 個月，CY2027 未讀，以 CY2026 代（差約 1–2%）"); i += 1
    RC.cell(i, 1, "機器人拉動的氣動市場 FY2030（百萬美元；9/12 × CY2030＋3/12 × CY2031）")
    for j, col in enumerate("DEF"): RC.cell(i, 2 + j, f"=0.75*In!{col}{KR['pne_2030']}+0.25*In!{col}{KR['pne_2031']}")
    RC.cell(i, 5, "線 6 情境!H44:I46；原則第 10 節第 49 條月數加權"); i += 1
    RC.cell(i, 1, "SMC 份額（低／基準／高）")
    for j, col in enumerate("DEF"): RC.cell(i, 2 + j, f"=In!{col}{KR['pai_sh']}")
    RC.cell(i, 5, "Assumed；全球氣動份額 36% 為公司自稱"); i += 1
    RC.cell(i, 1, "SMC 機器人用氣動營收 FY2026（¥B）")
    for j, L in enumerate("BCD"): RC.cell(i, 2 + j, f"={L}{pai0}*{L}{pai0+2}*fx_usd/1000")
    RC.cell(i, 5, "市場 × 份額 × 158 日圓"); i += 1
    RC.cell(i, 1, "SMC 機器人用氣動營收 FY2030（¥B）")
    for j, L in enumerate("BCD"): RC.cell(i, 2 + j, f"={L}{pai0+1}*{L}{pai0+2}*fx_usd/1000")
    RC.cell(i, 5, "匯率固定 158"); i += 1
    RC.cell(i, 1, "占 FY2026 合併營收（情境 B）")
    for j, L in enumerate("BCD"): RC.cell(i, 2 + j, f"={L}{pai0+3}/PL!{F}{B_['rev']}")
    RC.cell(i, 5, "實體 AI（機器人）曝險占比"); i += 1
    RC.cell(i, 1, "占 FY2030 合併營收（情境 B）")
    for j, L in enumerate("BCD"): RC.cell(i, 2 + j, f"={L}{pai0+4}/PL!{COL[2030]}{B_['rev']}")
    RC.cell(i, 5, "人形、AMR、車用的氣動內容在線模型基準為 0，未計"); i += 1
    RC.cell(i, 1, "對照：全球工業機器人裝機 CY2025 → CY2026 → CY2032（台；線 3）")
    RC.cell(i, 2, "=ir_25"); RC.cell(i, 3, "=ir_26"); RC.cell(i, 4, "=ir_32"); RC.cell(i, 5, "線 3 下游取數 IR.inst.global.*.base"); i += 1
    RC.cell(i, 1, "對照：機器人裝機 CY2026→CY2032 年複合（線 3）vs 本模型 FY2027–FY2032 營收年複合（情境 B）")
    RC.cell(i, 2, "=(ir_32/ir_26)^(1/6)-1"); RC.cell(i, 3, f"=(PL!{COL[2032]}{B_['rev']}/PL!{F}{B_['rev']})^(1/6)-1"); RC.cell(i, 4, f"=C{i}-B{i}")
    RC.cell(i, 5, "SMC 量的路徑不讀機器人裝機（機器人占需求小）；差異列出、不校準"); i += 1
    for rr_ in range(pai0 + 2, pai0 + 2 + 1):
        for L in "BCD": RC[f"{L}{rr_}"].number_format = "0%"
    for rr_ in (pai0 + 5, pai0 + 6, i - 1):
        for L in "BCD": RC[f"{L}{rr_}"].number_format = "0.0%"
    PAI = dict(start=pai0, end=i - 1)
    i += 1
    RC.cell(i, 1, "反解（整本重算，build_smc.py 線性內插後整本重算驗證；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 64; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_smc.py 以 LibreOffice 重算每一案例）；加權目標價、DCF 目標價為常數快照，相對基準與評等為公式"
    SE["A2"] = ("pne_2026、pne_2030、pne_2031、pai_sh 不列：只進 Recon 頁實體 AI 曝險節，對加權目標價結構上恆為 0。"
                "pay_r 不列：只影響收支關卡的股利，對加權目標價結構上恆為 0（收支關卡另見 Val 頁）。")
    SE["A3"] = "build_id"; SE["D3"] = sens["build_id"] if sens else "（基準檔）"
    SE["A4"] = "In 頁輸入指紋（敏感度重算時）"; SE["D4"] = sens["fp_sum"] if sens else 0
    SE["A5"] = "In 頁輸入指紋（本檔即時）"; SE["D5"] = f"=SUM(In!E4:E{in_last})"
    SE["A6"] = "基準加權目標價（敏感度重算時）"; SE["D6"] = sens["base"]["tp"] if sens else 0
    SE["A7"] = "模型狀態"; SE["D7"] = f'=IF(AND(ABS(D5-D4)<=1E-6*MAX(1,ABS(D4)),ABS(tp-D6)<1E-6*MAX(1,ABS(D6))),"CURRENT","STALE")'
    hdr(SE, 9, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    sr = 10
    if sens:
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]); SE.cell(sr, 5, f"=D{sr}/tp-1")
            SE.cell(sr, 6, cs["dcf"]); SE.cell(sr, 7, f"=F{sr}/Val!$B${r_dcft}-1")
            SE.cell(sr, 8, f'=IF(D{sr}/px-1>0.5,"強力買進",IF(D{sr}/px-1>0.25,"買進",IF(D{sr}/px-1>=-0.25,"中立",IF(D{sr}/px-1>=-0.5,"放空","強力放空"))))')
            SE[f"E{sr}"].number_format = "+0.0%;-0.0%"; SE[f"G{sr}"].number_format = "+0.0%;-0.0%"; SE[f"D{sr}"].number_format = "#,##0"; SE[f"F{sr}"].number_format = "#,##0"
            sr += 1
    SE.cell(sr + 1, 1, "註：相對倍數係數 rel_k 基準即為 1.00（未調整），故「相對倍數係數＝1」案例＝基準。相對基準以本檔即時加權目標價為分母；模型狀態 STALE 表示 In 頁改過、須重跑 build_smc.py")
    SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 34

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = "對帳 — 本模型讀到的線層級變數（線 6 共用零組件、線 3 工業機器人；取自 physical_ai/lines/ 成品，寫明檔名與儲存格）。與線模型不一致的自設值列在下方，不靜默改口徑"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔名＋儲存格）", "標記", "影響本模型哪些儲存格"])
    f0 = lambda x: f"{x:,.0f}"
    rec = [("每台工業機器人氣動內容（傳統／協作）", "6", "低 200／100、基準 500／300、高 1,500／800", "美元／台", "末端夾爪與周邊氣動", LC + " 輸入!D35:E35、D57:E57、D79:E79", "Assumed", "未進公式；說明 pne_*"),
           ("機器人拉動的氣動市場 CY2026", "6", f"{f0(LV['lo2026'])}／{f0(LV['b2026'])}／{f0(LV['hi2026'])}", "百萬美元", "五線歸因", LC + " 情境!D44:D46", "Derived", "In pne_2026 → Recon 實體 AI 曝險"),
           ("機器人拉動的氣動市場 CY2030", "6", f"{f0(LV['lo2030'])}／{f0(LV['b2030'])}／{f0(LV['hi2030'])}", "百萬美元", "五線歸因", LC + " 情境!H44:H46", "Derived", "In pne_2030 → Recon"),
           ("機器人拉動的氣動市場 CY2031", "6", f"{f0(LV['lo2031'])}／{f0(LV['b2031'])}／{f0(LV['hi2031'])}", "百萬美元", "五線歸因", LC + " 情境!I44:I46", "Derived", "In pne_2031 → Recon"),
           ("人形、AMR、車用每台氣動內容", "6", "0（基準）", "美元／台", "線模型假設", LC + " 輸入!F57:L57", "Assumed", "未計入曝險（若人形末端改用氣動夾爪才會增加）"),
           ("SMC 份額（氣動元件）", "6", "線模型：未取得（份額!F17）；本模型 36%（25–45%）", "%", "公司自稱全球份額", LC + " 份額!F17；" + BR2Q, "Interested-party", "In pai_sh（本模型自設，線模型無值）"),
           ("全球工業機器人裝機 CY2025、CY2026、CY2032（基準）", "3", f"{f0(LV['ir']['IR.inst.global.2025.base'])}、{f0(LV['ir']['IR.inst.global.2026.base'])}、{f0(LV['ir']['IR.inst.global.2032.base'])}", "台", "IFR 新裝機（含協作）", LIR + " 下游取數 E12、E13、E19（IR.inst.global.*.base）", "Derived", "In ir_* → Recon 對照列"),
           ("不一致 1：SMC 銷售量年增 vs 機器人裝機年增", "3", "本模型 FY2028–29 5%、FY2030–32 4%（情境 B）；線 3 裝機 2026→2032 年複合 6.6%", "%/年", "SMC 量的路徑不讀機器人裝機", "本模型 In vgm_B、vglt_B；" + LIR + " 下游取數", "Assumed", "理由：機器人拉動的氣動市場（基準 CY2026 約 3 億美元）只占 SMC 營收約 4–5%，需求主體是半導體、汽車、工具機、食品等一般工廠自動化"),
           ("不一致 2：氣動市場口徑", "6", "線模型為五線歸因（只計機器人）；SMC 營收為全部工廠自動化", "—", "口徑不同", LC + " 說明!A14", "—", "曝險節只作對照，不進目標價")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 44; RE.column_dimensions["F"].width = 56

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("設備投資循環轉折（半導體、資料中心投資放緩）", "是（vg27_C、gadj_C）", "情境 C 與 Sens", "受注指數（月）、半導體／電機產業別受注、SEMI 設備出貨"),
             ("匯率（人民幣、美元、歐元升值反轉）", "否", "公司揭露 1% 變動對營業利益：歐元 48.3 億、人民幣 19.3 億、美元 15 億（" + BR2Q + "）；未逐年建模，業外匯兌損益設 0", "USDJPY、EURJPY、CNYJPY"),
             ("中國電動車、二次電池投資放緩", "部分（g_anc 低案）", "公司下半年說法", "中國產業別受注"),
             ("中國在地化（半導體在地生產、陸系氣動廠）", "否", "份額數據未取得", "中國地區受注與份額（公司自稱 36%）"),
             ("高存貨政策的營運資金占用", "是（fc_wc）", "Sens", "存貨 ÷ 營收、營業現金流"),
             ("人形與 AMR 末端採用氣動夾爪的選擇權", "否（只作對照）", "線 6 基準每台內容為 0", "人形整機 BOM 拆解、公司法說")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 60

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([TS2Q, BR2Q, KBT, FWM, EDB, SAI, SAB, SAC, SAF, SHA, YCL, JGB10, FAN, HDS, NAB, LC, LIR], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"SMC（6273.T）個股模型 {VER}｜build_id {{BID}}"
    RD["A2"] = '="資料截止日 2026-11-23（中間決算 2026-11-13）；成品日 2026-11-23；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"#,##0")&" 日圓）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：SMC 的股價由工廠自動化設備投資循環（現由半導體與資料中心帶動）決定——營收＝銷售量 × 價格（單一報告分部，地區子列），毛利率與營業費用彈性決定利潤槓桿；機器人（實體 AI）只占需求一小部分，另列曝險、不主導目標價"
    RD["A4"] = (f'="在 FY2027 銷售量情境 B 年增 "&TEXT(vg27_B,"0.0%")&"、FY2027 起毛利率較指引 "&TEXT(gadj_B*100,"+0.0;-0.0")&" 個百分點、同業前瞻 P/E 中位數 "&TEXT(pe_med,"0.0")&" 倍的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"#,##0")&" 日圓（DCF "&TEXT(Val!B{r_dcft},"#,##0")&"、Comparable "&TEXT(Val!B{r_cmpt},"#,##0")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0.0")&"）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 日圓。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：受注指數（FY2025＝100）連兩個月低於 FY2026 營收指數 127；半導體／電機受注指數跌破 150；單季毛利率跌破 45%；季營收年增轉負"
    RD["A7"] = f'="實體 AI 曝險（只作對照，不進目標價）：機器人拉動的氣動營收 FY2026 約 "&TEXT(Recon!C{pai0+3},"0.0")&" 十億日圓（區間 "&TEXT(Recon!B{pai0+3},"0.0")&"–"&TEXT(Recon!D{pai0+3},"0.0")&"），占營收 "&TEXT(Recon!C{pai0+5},"0.0%")&"（區間 "&TEXT(Recon!B{pai0+5},"0.0%")&"–"&TEXT(Recon!D{pai0+5},"0.0%")&"）；FY2030 "&TEXT(Recon!C{pai0+6},"0.0%")&"。Assumed：線 6 每台氣動內容與 SMC 份額皆為假設值"'
    RD["A8"] = "範圍：不讀 Siliconomics／Tokenomics／Wattonomics（算力線）；自用 AI 讀法、上線時程等規格條文不適用。匯率不預測（業外匯兌損益設 0，營收以 FY2026 想定匯率延續）。HTML 展示層見 dist/ 同名 .html"
    wb.save(path)
    return dict(VR=VR, PR=PR, DR=DR, KR=KR, PAI=PAI, in_last=in_last)


# ---------------- 重算與讀值 ----------------
LO = "file:///tmp/h_smc/lo_profile"
def recalc(paths, outdir):
    os.makedirs(outdir, exist_ok=True)
    for i in range(0, len(paths), 20):
        chunk = paths[i:i + 20]
        for attempt in (1, 2):
            try:
                subprocess.run(["soffice", f"-env:UserInstallation={LO}", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", outdir] + chunk,
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
                break
            except subprocess.TimeoutExpired:
                subprocess.run(["pkill", "-f", "lo_profile"], check=False)
                if attempt == 2: raise
    return [os.path.join(outdir, os.path.basename(p)) for p in paths]


def read_vals(path, M):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; VR = M["VR"]; DR = M["DR"]
    d = dict(tp=V[f"B{VR['tp']}"].value, dcf=V[f"B{VR['dcf']}"].value, cmp=V[f"B{VR['cmp']}"].value, up=V[f"B{VR['up']}"].value, rt=V[f"B{VR['rt']}"].value,
             gate=V[f"C{VR['gate']}"].value)
    I = wb["In"]; w = {}; wd = None; s = 0.0
    for r in range(4, M["in_last"] + 1):
        k = I.cell(r, 1).value; v = I.cell(r, 5).value
        if isinstance(v, (int, float)): s += v
        if k in ("w_A", "w_B", "w_C"): w[k[-1]] = v
        if k == "w_dcf": wd = v
    alt = sum(w[x] * wb["DCF"][f"B{DR[x]['alt']}"].value for x in SC)
    d["alt_dcf"] = alt; d["alt_tp"] = wd * alt + (1 - wd) * d["cmp"]; d["fp_sum"] = wb["Sens"]["D5"].value
    return d


def run_batch(jobs):
    """jobs：{tag: override dict} → {tag: 讀值}。一次建置、批次重算。"""
    d = tempfile.mkdtemp(prefix="smc_"); paths = {}; M = None
    for tag, ov in jobs.items():
        p = os.path.join(d, f"c_{tag}.xlsx"); M = build(p, ov); paths[tag] = p
    outs = recalc(list(paths.values()), os.path.join(d, "o"))
    res = {tag: read_vals(o, M) for tag, o in zip(paths.keys(), outs)}
    shutil.rmtree(d, ignore_errors=True)
    return res


def in_val(k, col):
    for row in ROWS:
        if row[0] == k:
            return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)


def fingerprint():
    return hashlib.sha1(json.dumps([r for r in ROWS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]


def rating(u):
    return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"


CASES = [  # (鍵, 驅動)
    ("g_anc", "FY2026 實績相對指引"), ("vg27_B", "FY2027 銷售量年增（情境 B）"), ("vgm_B", "FY2028–29 銷售量年增（情境 B）"), ("vglt_B", "FY2030–32 銷售量年增（情境 B）"),
    ("prc_B", "每年價格變化（情境 B）"), ("gadj_B", "毛利率調整（情境 B）"), ("e_sga", "營業費用彈性"), ("nonop", "業外損益"), ("tax_w", "有效稅率"),
    ("fc_da", "折舊 ÷ 營收"), ("cx_27", "FY2027 設備投資"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("stub_wc", "下半年營運資金增加"),
    ("bb_rem", "尚待買回金額"), ("inv_k", "投資有價證券計入比例"), ("beta", "β"), ("erp", "ERP"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"),
    ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重")]
BETA_B = statistics.median([0.67 * 0.83 + 0.33, 0.67 * 0.97 + 0.33])
TIERS = [0.5, 0.25, -0.25, -0.5]
REV_KEYS = [("pe_med", 15.17, 5.0, 35.0, "同業前瞻 P/E 中位數（倍；三情境同動，只影響 Comparable 腿）"),
            ("gadj_B", -0.01, -0.15, 0.06, "情境 B 的 FY2027 起毛利率調整（pt；其他不變，整本重算；情境 A／C 不動）"),
            ("wacc_adj", 0.0, -0.04, 0.08, "WACC 加減（三情境同動；只影響 DCF 腿）")]


def main():
    if "--raw" in sys.argv:
        ov = json.loads(os.environ.get("SMC_OVERRIDE", "{}"))
        build(sys.argv[sys.argv.index("--raw") + 1], ov); return
    if "--post" in sys.argv:   # 沿用 sens.json 的整本重算快照，只重做成品、HTML、核對、README（In 頁改過時模型狀態會顯示 STALE）
        sj = json.load(open(os.path.join(HERE, "sens.json")))
        sens = dict(base=sj["base"], cases=sj["cases"], fp=sj["fp"], fp_sum=sj["base"]["fp_sum"], build_id=sj["build_id"])
        finalize(sens, sj["revs"]); return
    px0 = in_val("px", "base")
    jobs = {"base": {}}
    meta = {}
    for k, drv in CASES:
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if k == "beta": v = BETA_B + (-0.3 if side == "lo" else 0.3)
            if v is None: continue
            b_ = in_val(k, "base")
            if isinstance(b_, (int, float)) and abs(v - b_) < 1e-12: continue
            tag = f"{k}_{side}"; jobs[tag] = {k: v}
            meta[tag] = dict(drv=drv + ("（低）" if side == "lo" else "（高）"), set=f"{k}＝{v:.4g}", key=k)
    for s in SC:
        tag = f"only_{s}"; jobs[tag] = {f"w_{x}": (1.0 if x == s else 0.0) for x in SC}
        meta[tag] = dict(drv=f"只看情境 {s}", set=f"w_{s}＝1，其餘 0", key=f"w_{s}")
    # 反解格點
    GRID = 9
    for key, b0, lo_, hi_, _ in REV_KEYS:
        for n in range(GRID):
            jobs[f"rv_{key}_{n}"] = {key: lo_ + (hi_ - lo_) * n / (GRID - 1)}
    res = run_batch(jobs)
    base = res["base"]; print("基準", {k: base[k] for k in ("tp", "dcf", "cmp", "up", "rt", "gate")})
    cases = []
    for tag, m in meta.items():
        r = res[tag]; cases.append(dict(**m, tp=r["tp"], dcf=r["dcf"], cmp=r["cmp"], up=r["up"], rt=r["rt"]))
    up_ = base["alt_tp"] / px0 - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rating(up_)))
    # 反解：格點線性內插 → 整本重算驗證
    up0 = base["up"]
    upper = min([t for t in TIERS if t > up0], default=None); lower = max([t for t in TIERS if t <= up0], default=None)
    targets = [("目標價＝現價（上檔 0%）", 0.0)] + ([(f"升一級邊界（上檔 {upper:+.0%}）", upper)] if upper is not None else []) + ([(f"降一級邊界（上檔 {lower:+.0%}）", lower)] if lower is not None else [])
    sol = {}; vjobs = {}
    for key, b0, lo_, hi_, note in REV_KEYS:
        xs = [lo_ + (hi_ - lo_) * n / (GRID - 1) for n in range(GRID)]; ys = [res[f"rv_{key}_{n}"]["up"] for n in range(GRID)]
        for lab, tgt in targets:
            x = None
            for n in range(GRID - 1):
                if (ys[n] - tgt) * (ys[n + 1] - tgt) <= 0 and ys[n] != ys[n + 1]:
                    x = xs[n] + (tgt - ys[n]) * (xs[n + 1] - xs[n]) / (ys[n + 1] - ys[n]); break
            sol[(key, lab)] = x
            if x is not None: vjobs[f"vf_{key}_{len(vjobs)}"] = (key, lab, x)
    vres = run_batch({t: {k: x} for t, (k, lab, x) in vjobs.items()}) if vjobs else {}
    vmap = {(k, lab): vres[t]["up"] for t, (k, lab, x) in vjobs.items()}
    revs = []
    for key, b0, lo_, hi_, note in REV_KEYS:
        for lab, tgt in targets:
            x = sol[(key, lab)]
            if x is None:
                revs.append(dict(label=lab, key=key, val=f"區間 {lo_:+.0%}～{hi_:+.0%} 內無解", base=b0, note=note))
            else:
                revs.append(dict(label=lab, key=key, val=x, base=b0, note=note + f"；驗證：重算上檔 {vmap[(key, lab)]:+.2%} vs 目標 {tgt:+.0%}"))
            print("反解", key, lab, x)
    fp = fingerprint(); bid = f"smc-{VER}-{fp}"
    sens = dict(base=base, cases=cases, fp=fp, fp_sum=base["fp_sum"], build_id=bid)
    json.dump(dict(base={k: base[k] for k in ("tp", "dcf", "cmp", "up", "rt", "alt_dcf", "alt_tp", "fp_sum")}, cases=cases, fp=fp, build_id=bid, revs=revs),
              open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
    finalize(sens, revs)


def finalize(sens, revs):
    bid = sens["build_id"]
    # 成品
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="smcf_"); raw = os.path.join(tmp, NAME)
    M = build(raw, {}, sens, revs)
    wb = openpyxl.load_workbook(raw); wb["README"]["A1"] = wb["README"]["A1"].value.replace("{BID}", bid); wb.save(raw)
    out = recalc([raw], os.path.join(tmp, "o"))[0]
    shutil.copy(out, os.path.join(dist, NAME))
    X = os.path.join(dist, NAME)
    wbv = openpyxl.load_workbook(X, data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value[:4] in ("#DIV", "#VAL", "#REF", "#NAM", "#N/A", "#NUM", "#NUL") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10])
    assert not errs, errs
    assert abs(wbv["Val"][f"B{M['VR']['tp']}"].value - base["tp"]) < 1e-6 * base["tp"], "成品與快照基準不一致"
    json.dump(dict(company="smc", ticker="6273.T", model_sheet="PL",
                   official_segments=[dict(name="自動制御機器事業（單一報告分部）", label="營收（合併；自動制御機器事業，單一報告分部）", period="FY2025（2026年3月期）",
                                           value=842.541, unit="¥B", source=TS2Q + "（セグメント情報：単一セグメントのため記載省略）；" + EDB)],
                   sublines=[dict(parent="自動制御機器事業（單一報告分部）", labels=[f"地區子列：{lab}" for _, lab in REG])],
                   check_labels=["檢查：地區子列合計 − 合併營收"],
                   note="公司只有單一報告分部；地區別營收（有價證券報告書「地域ごとの情報」，顧客所在地）為子列，FY2023–FY2025 為實績（四捨五入差 < 0.01），FY2026 起以 FY2025 結構分配。PL 頁三情境各有「檢查：地區子列合計 − 合併營收」"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)
    print("成品", X)
    # HTML 展示層
    subprocess.run([sys.executable, os.path.join(HERE, "gen_html_json.py"), X], check=True, cwd=ROOT)
    subprocess.run([sys.executable, os.path.join(ROOT, "tools/company_html/build_company_html.py"), HERE], check=True, cwd=ROOT)
    html = X[:-5] + ".html"
    ok, chk = run_checks(X, html)
    write_readme(X, M, bid, sens, revs, chk)
    print("\n".join(chk)); print("核對全過" if ok else "核對有失敗")


CHECKS = [["tools/in_layout/check_in_layout.py", "{x}"], ["tools/sens_coverage/check_sens_coverage.py", "{x}", "--require-keycol"],
          ["tools/model_governance/check_governance.py", "{x}", "--html", "{h}"], ["tools/model_governance/check_rating_gate.py", "{x}", "--strict"],
          ["tools/model_governance/check_zero_rows.py", "{x}"], ["tools/model_governance/check_cached_values.py", "{x}"],
          ["tools/segment_check/check_segments.py", "{d}"], ["tools/company_html/check_company_html.py", "{d}"],
          ["tools/company_html/check_layout_v13.py", "{h}", "--xlsx", "{x}"]]


def run_checks(x, h):
    out, ok = [], True
    for c in CHECKS:
        cmd = [sys.executable, os.path.join(ROOT, c[0])] + [a.format(x=x, h=h, d=HERE) for a in c[1:]]
        rp = subprocess.run(cmd, capture_output=True, text=True, cwd=ROOT, env=dict(os.environ))
        lines = [l for l in rp.stdout.strip().splitlines() if l.strip()]
        line = lines[0] if lines else rp.stderr.strip()[-200:]
        line = line.replace(ROOT + "/", "")
        out.append(f"`{os.path.basename(c[0])}`：{'通過' if rp.returncode == 0 else '失敗'}（{line[:160]}）")
        ok &= rp.returncode == 0
    return ok, out


def write_readme(X, M, bid, sens, revs, chk=None):
    wb = openpyxl.load_workbook(X, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; VR = M["VR"]; PR = M["PR"]
    g = lambda c: V[c].value
    f0 = lambda x: f"{x:,.0f}"; pc = lambda x: f"{x*100:+.1f}%"
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:8]
    def rvv(r):
        return r["val"] if isinstance(r["val"], str) else f"{r['val']:.4f}"
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {rvv(r)} | {r['base']} | {r['note']} |" for r in revs)
    yr = lambda s, k, fm="{:,.1f}": " | ".join(fm.format(P[f'{COL[y]}{PR[s][k]}'].value) if isinstance(P[f'{COL[y]}{PR[s][k]}'].value, (int, float)) else "—" for y in YRS)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f0(g(chr(66+i)+str(VR['dcfr'])))} | {f0(g(chr(66+i)+str(VR['ca'])))} | {f0(g(chr(66+i)+str(VR['tps'])))} |" for i, s in enumerate(SC))
    IN = wb["In"]; inr = {IN.cell(r, 1).value: r for r in range(4, IN.max_row + 1) if IN.cell(r, 1).value}
    iv = lambda k: IN.cell(inr[k], 5).value
    rts = sorted(set(c['rt'] for c in sens['cases']))
    chk_txt = "\n".join(f"- {c}" for c in (chk or ["（建置時未跑核對；見工作單回報）"]))
    txt = f"""# physical_ai/smc — SMC（SMC株式会社，6273.T）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

東證 Prime 上市（亞太基金可投資）。會計年度 3 月底，年度標籤採起始年（FY2026＝2026-04～2027-03，即 2027年3月期）。本 README 的數字由 `build_smc.py` 從重算後的成品讀出產生，不手寫。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

**實體 AI 曝險：** {RD['A7'].value}

## 估值（評價日 2027-03-31；目標時點 2027-11-20）
| 情境 | DCF（推到目標時點） | Comparable（同業 P/E × NTM EPS） | 情境目標價 |
|---|---|---|---|
{scn}

- 加權目標價 **{f0(tp)} 日圓**（DCF {f0(dcf)}、Comparable {f0(cmp_)}），現價 {f0(iv('px'))}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")}（情境 B FY2026–FY2030 自由現金流扣股利 {g(f"C{VR['gate']}"):,.1f} 十億日圓）→ 投資建議 **{g(f"B{VR['rec']}")}**。
- 方法：DCF 標準版 FCFF（WACC／ke／g／RONIC＝{g(f"B{VR['wk']}")}；無有息負債，WACC＝ke）50%＋Comparable 50%（同業 8 家前瞻 P/E 中位數 {iv('pe_med'):.2f} 倍 × 目標時點起算 NTM EPS）。相對倍數係數 1.00（未調整：缺 SMC 對同業的歷史相對倍數）。
- 隱含前瞻 P/E（加權目標價 ÷ 情境 B NTM EPS）{g(f"B{VR['pe']}"):.1f} 倍；現行前瞻 P/E {iv('pe_self'):.2f} 倍（現價 ÷ FY2026 公司指引 EPS 3,789.21）。DCF 終值占 EV（情境 B）{g(f"B{VR['tvsh']}")*100:.1f}%。
- 賣方平均目標價 {f0(iv('cons_tp'))}（13 位，2026-08-13，中間決算前；只作對照）。

## 營收與損益（PL 頁三情境；情境 B 摘要，¥B）
| 列 | {" | ".join(f"FY{y}" for y in YRS)} |
|---|{"---|" * len(YRS)}
| 營收（單一分部合計） | {yr('B','rev')} |
| 　其中 中國 | {yr('B','cn')} |
| 　其中 日本 | {yr('B','jp')} |
| 毛利率 | {yr('B','gm','{:.1%}')} |
| 營業利益 | {yr('B','op')} |
| 營業利益率 | {yr('B','opm','{:.1%}')} |
| 淨利（歸屬母公司） | {yr('B','ni')} |
| EPS（日圓） | {yr('B','eps','{:,.0f}')} |
| 自由現金流扣股利 | {yr('B','fcfad')} |

- FY2023–FY2025 實績（日本會計準則；FY2025＝2026年3月期營收 842.5、營業利益 190.6）；FY2026＝公司 2026-11-13 上修指引（營收 1,073、營業利益 294、淨利 238）；FY2027 起為模型。
- 命題驅動：營收＝前一年 ×(1＋銷售量年增)×(1＋價格變化)；毛利率＝FY2026 指引 47.5%＋情境調整；營業費用＝前一年 ×(1＋彈性 × 營收年增)。
- 分部：公司只有單一報告分部（自動制御機器事業）；地區子列（日本、中國、亞洲、歐洲、美國、其他）FY2023–FY2025 為有價證券報告書實績，FY2026 起以 FY2025 結構分配（Assumed，未取得 FY2026 地區別計畫）。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 銷售量年增（情境 A／B／C）：FY2027 15%／6%／−10%；FY2028–29 8%／5%／2%；FY2030–32 6%／4%／2%。價格每年 1.5%／0.5%／0%。
- 毛利率較 FY2026 指引 0／−1／−2.5 個百分點；營業費用彈性 0.8（0.5–1.2）；業外損益 22（12–32，匯兌不預測）；稅率 28%。
- 現金流：折舊 5.8% × 營收、FY2027 設備投資 95（公司「1,000 億円弱」）、之後 7% × 營收、營運資金 70% × 營收增量；下半年營運資金增加 0（−30～+30）。
- 折現：rf 2.817%、β {iv('beta'):.3f}（SMC 0.83 與同業 0.97 的 Blume 中位數）、ERP 5%、CRP 0.91%、g＝MIN(3%, rf)、RONIC＝WACC。
- 評價日股數：2026-09-30 流通股 62.34 百萬股，扣除 11-13 決議的 500 億円買回（以現價全數執行；Assumed）。股權價值另扣 2026-12 中間股利與 FY2026 期末股利（股價日已過中間配當基準日）。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| FY2027 銷售量（vg27） | 公司月度受注指數（10 月 151，FY2025＝100）、半導體／電機產業別受注（205） |
| 中期銷售量（vgm、vglt） | 季營收年增（扣匯率）、地區別量的變化（簡報） |
| 毛利率調整（gadj） | 季毛利率（上半 47.9%）、價格效果（上半 +0.6%） |
| 營業費用彈性（e_sga） | 季販管費年增 ÷ 營收年增 |
| 營運資金（fc_wc） | 存貨 ÷ 營收、營業現金流 |

{RD['A6'].value}。Sens 各案例估值評等出現「{'、'.join(rts)}」。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f0(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁。

## 反解（整本重算；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
- 線 6（`{LC}`）：每台工業機器人氣動內容 傳統 500、協作 300 美元（輸入!D57:E57，Assumed）；機器人拉動的氣動市場 CY2026 基準約 3.1 億美元、CY2030 約 3.9 億美元（情境!D45、H45；低／高列 44、46）→ In pne_* → Recon「實體 AI 曝險」節。人形、AMR、車用每台氣動內容基準為 0。
- 線 3（`{LIR}`）：全球工業機器人裝機 CY2025 603,307 台、CY2026 655,010、CY2032 959,984（下游取數 E12、E13、E19）→ Recon 對照列。
- 與線模型不一致（自設值）：SMC 銷售量路徑不讀機器人裝機（線 3 裝機年複合 6.6%，本模型中期 4–5%）；理由：機器人拉動的氣動需求只占 SMC 營收約 {wb['Recon'].cell(M['PAI']['start'] + 5, 3).value*100:.1f}%（FY2026，情境 B），需求主體是一般工廠自動化。線 6 未給 SMC 份額（份額!F17 未取得），本模型以公司自稱 36%（25–45%）。

## 限制與未取得資料
- **產業別營收未取得**（簡報只有受注指數與定性說明）；半導體／資料中心相關營收占比未知，命題以全公司銷售量代替。
- FY2026 地區別營收計畫未取得；FY2026 起地區子列以 FY2025 結構分配（中國、亞洲上半量的成長較快，結構固定會低估其占比）。
- 氣動元件市場的第三方規模與 SMC 以外同業（CKD、Festo、Parker）報價未取得；同業 P/E 為 FA／機器人 8 家，兩種分母口徑（公司予想 5 家、賣方共識 3 家）混用。
- 租賃負債未單獨揭露（未取得，金額小未計）；FY2023 年底營運資金明細未取得。
- 賣方共識目標價與營收為中間決算前（2026-08／09），只作歷史對照；共識 EPS 逐年值未取得（付費牆）。
- 匯率：模型不預測匯率，營收以 FY2026 想定匯率延續；公司揭露 1% 匯率變動對營業利益的影響列在 Risk 頁。

## 核對（建置時執行）
{chk_txt}

## 檔案
- `build_smc.py`：建置＋整本重算 Sens／反解＋產生本 README 與 `segments.json`、`sens.json`，並跑核對。
- `dist/{NAME}`：成品（工作表 README、In、Hist、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-23）：初版；現價 {f0(iv('px'))}（2026-11-20 收盤）；加權目標價 {f0(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


if __name__ == "__main__":
    main()

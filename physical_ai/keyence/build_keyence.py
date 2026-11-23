# Keyence（キーエンス，6861.T）個股模型 v0.1 — AI實體應用 Project（線 6 共用零組件：工業感測與機器視覺）
# 用法：python3 physical_ai/keyence/build_keyence.py   → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、segments.json、sens.json
# 會計年度：3 月 20 日決算（日本會計基準）。本模型 FY 以「起始年」命名，與 Fanuc 模型、公司「2026年度」用語一致：
#   FY2025＝2025-03-21～2026-03-20（2026年3月期）；FY2026＝2026-03-21～2027-03-20（2027年3月期）。
# 金額：十億日圓（¥B）；股數：百萬股；EPS、股價、股利：日圓。
# **股票分割：2026-11-21 生效 1 拆 5（基準日 2026-11-20；Yahoo 日線 2026-11-19 起調整）。本模型股數、EPS、每股股利一律為分割後口徑；
#   財報原文的分割前數字放在 Hist／In 頁，以 split_r 換算（Derived）。**
# 結構：README、In、Hist、Q、PL（三情境：國內＋海外 → 單一官方分部 → 標準損益表 → 現金流）、DCF、Val、Recon、Sens、對帳、實體AI、Risk、Sources。
# 不讀 Siliconomics／Tokenomics／Wattonomics（實體 AI 線不適用）。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261123"
NAME = f"{DATE}_Keyence個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032

# ---------------- 出處 ----------------
ER26 = "公司 2026年3月期 決算短信〔日本基準〕(連結)（2026-04）https://www.keyence.co.jp/pdf/EarningsRelease_202604_ja.pdf"
ER25 = "公司 2025年3月期 決算短信（2025-04；含 2024年3月期 比較欄）https://www.keyence.co.jp/pdf/EarningsRelease_202504_ja.pdf"
ERQ1 = "公司 2027年3月期 第1四半期 決算短信（2026-07）https://www.keyence.co.jp/pdf/EarningsRelease_202607_ja.pdf"
ERH1 = "公司 2027年3月期 第2四半期（中間期）決算短信（2026-10-29）https://www.keyence.co.jp/pdf/EarningsRelease_202610_ja.pdf"
FRH1 = "公司 2026年度 第2四半期 決算補足資料（2026-10-29）https://www.keyence.co.jp/pdf/FinancialResults_202610_ja.pdf"
MS = "MarketScreener KEYENCE 財務頁（共識 16 位；讀取 2026-11-20 收盤後）https://ca.marketscreener.com/quote/stock/KEYENCE-CORPORATION-105520876/finances/"
KY = "株予報 kabuyoho.jp 6861 アナリスト頁（資料更新 2026-07-17）https://kabuyoho.jp/reportAnalyst?bcode=6861"
SA = "stockanalysis.com statistics 頁"
TRD = "traders.co.jp 個股頁（2026-11-16 收盤；經 physical_ai/fanuc v0.1.1 Val 頁同業表）"
YCD = "Yahoo Finance 日線（physical_ai/data/closes2_20261109_20261120.csv；分割調整後）"
SPL = "physical_ai/data/closes2_20261109_20261120.csv.splits.txt（6861.T 2026-11-19 分割 5.0）"
JGB10 = "日本財務省 國債金利情報 10 年：2026-11-02 2.817%（與 physical_ai/fanuc、harmonic_drive 同一取值，跨模型一致）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"
CP = "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx"
IRX = "physical_ai/lines/industrial_robot/dist/20261116_工業機器人單位經濟_v0.1.xlsx"

# ---------------- In 列 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源1, 來源2／說明, 來源誘因)
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. 股票分割與股數（2026-11-21 生效 1 拆 5；本模型全部用分割後口徑）")
inp("split_r", "股票分割比率（1 股分為 5 股；基準日 2026-11-20，生效 2026-11-21）", "x", None, 5, None, "Verified",
    ERH1 + "：普通株式 1 株につき 5 株の割合で株式分割", SPL + "；csv 收盤為分割調整後價格。Hist 頁 EPS、股利、股數為分割前原文，PL 頁除以 split_r（Derived）")
inp("sh_iss", "發行股數（2026-09-20，分割前）", "百萬股", None, 243.207684, None, "Verified", ERH1 + "：期末発行済株式数 243,207,684 株", "")
inp("sh_trs", "自己株式數（2026-09-20，分割前）", "百萬股", None, 0.682755, None, "Verified", ERH1 + "：期末自己株式数 682,755 株", "")
inp("sh_v", "評價日股數（分割後）＝(發行 − 自己株式) × split_r", "百萬股", None, "=(sh_iss-sh_trs)*split_r", None, "Derived",
    "＝(243.207684 − 0.682755) × 5 ＝ 1,212.62 百萬股", "無 RSU、選擇權或潛在股份揭露（決算短信未列潛在株式調整後 EPS）→ DCF 標準規格第 3 節不另加")
inp("sh_26", "FY2026 平均股數（分割後；中間期平均 × split_r）", "百萬股", None, "=242.525062*split_r", None, "Derived", ERH1 + "：期中平均株式数 242,525,062 株（分割前）", "")

sec("B. FY2026 上半實績（2026-03-21～09-20；Verified）")
inp("h1_rev", "FY2026 上半 營收", "¥B", None, 735.399, None, "Verified", ERH1 + "：売上高 735,399 百萬日圓（+34.9%）", "前年同期 545,300")
inp("h1_cogs", "FY2026 上半 營業成本（売上原価）", "¥B", None, 105.433, None, "Verified", ERH1, "毛利率 85.7%（Derived）")
inp("h1_sga", "FY2026 上半 營業費用（販管費）", "¥B", None, 221.703, None, "Verified", ERH1, "")
inp("h1_op", "FY2026 上半 營業利益", "¥B", None, 408.262, None, "Verified", ERH1 + "：408,262（+50.0%）", "")
inp("h1_nonop", "FY2026 上半 業外收支淨額（營業外收益 18,799 − 費用 2,713）", "¥B", None, "=18.799-2.713", None, "Derived", ERH1 + "：受取利息 13,967、持分法 3,202、為替差損 1,992", "")
inp("h1_ni", "FY2026 上半 歸屬母公司淨利", "¥B", None, 298.675, None, "Verified", ERH1 + "：298,675（+49.4%）；EPS 1,231.53（分割前）", "")
inp("h1_dom_g", "FY2026 上半 國內營收年增（日圓）", "%", None, 0.214, None, "Verified", FRH1 + "：上期 國內 +21.4%", "")
inp("h1_ovs_lc", "FY2026 上半 海外營收年增（當地貨幣）", "%", None, 0.273, None, "Verified", FRH1 + "：上期 海外 円ベース +41.8%、現地通貨 +27.3%", "北中南米 +21.8%、アジア +39.5%、欧州・その他 +12.5%（現地通貨）")
inp("h1_fx", "FY2026 上半 匯率換算效果（(1＋日圓增速) ÷ (1＋當地貨幣增速) − 1）", "%", None, "=1.418/1.273-1", None, "Derived", FRH1 + "：海外 +41.8% ÷ +27.3%；為替影響額 売上 約 +519 億、営業利益 約 +384 億", "平均匯率 上期 USD 160（前年 146）、EUR 185（166）、CNY 23.5（20.3）")

sec("C. 命題模組：營收＝國內 ×(1＋國內增速) ＋ 海外 ×(1＋當地貨幣增速)×(1＋匯率換算效果)；公司不揭露業績預測")
inp("gd26", "FY2026 全年 國內營收年增", "%", 0.16, 0.19, 0.22, "Assumed", "上半 +21.4%（Verified）；下半基期較高（FY2025 下半營收 624.0 vs 上半 545.3，Derived）", "低／高 ±3pt")
inp("go26", "FY2026 全年 海外營收年增（當地貨幣）", "%", 0.21, 0.25, 0.28, "Assumed", "上半 +27.3%（Verified）；業種別上半 海外 半導體・液晶 +65%、電機・精密 +35%、自動車 +10%（" + FRH1 + "）", "")
inp("fx26", "FY2026 全年 海外匯率換算效果", "%", 0.05, 0.08, 0.10, "Assumed", "上半 +11.4%（Derived）；2026-11-20 USDJPY 159.71（" + YCD + "）≈ 上半平均 160", "下半前年同期平均匯率未取得 → 下半效果假設約 +4–5%，全年約 +8%")
inp("gm26", "FY2026 毛利率", "%", 0.850, 0.855, 0.860, "Assumed", "上半 85.7%（Derived）；FY2025 83.0%、FY2024 83.8%（Hist）", "")
inp("sga26_r", "FY2026 營業費用率（販管費 ÷ 營收）", "%", 0.29, 0.30, 0.31, "Assumed", "上半 30.1%（Derived）；FY2025 32.1%", "")
for s, v, lo, hi, n in (("A", 0.10, None, None, "樂觀：日本設備投資回升延續（上半國內 +21.4%）"), ("B", 0.05, 0.02, 0.08, "基準：FY2026 高基期後放緩"),
                        ("C", -0.05, None, None, "保守：半導體與電子設備投資週期回落")):
    inp(f"gd27_{s}", f"情境 {s}：國內營收年增 FY2027", "%", lo, v, hi, "Assumed", n, "區間見情境 A／B／C 三列")
for s, v, lo, hi in (("A", 0.07, None, None), ("B", 0.05, 0.03, 0.07), ("C", 0.02, None, None)):
    inp(f"gdm_{s}", f"情境 {s}：國內營收年增 FY2028–FY2029", "%/年", lo, v, hi, "Assumed", "FY2023→FY2025 國內年複合 6.4%（Derived：344.567 → 390.066）", "")
for s, v, lo, hi in (("A", 0.05, None, None), ("B", 0.04, 0.02, 0.06), ("C", 0.02, None, None)):
    inp(f"gdlt_{s}", f"情境 {s}：國內營收年增 FY2030–FY2032", "%/年", lo, v, hi, "Assumed", "長期：日本名目設備投資成長＋滲透", "")
for s, v, lo, hi, n in (("A", 0.16, None, None, "樂觀：AI 半導體設備投資延續（上半海外 半導體・液晶 +65%）"), ("B", 0.10, 0.06, 0.14, "基準：高基期後放緩；共識 FY2027 營收 +10.7%（日圓，Interested-party）"),
                        ("C", -0.08, None, None, "保守：半導體與電子投資回落、中國需求轉弱")):
    inp(f"go27_{s}", f"情境 {s}：海外營收年增 FY2027（當地貨幣）", "%", lo, v, hi, "Assumed", n, "")
for s, v, lo, hi in (("A", 0.12, None, None), ("B", 0.09, 0.06, 0.12), ("C", 0.05, None, None)):
    inp(f"gom_{s}", f"情境 {s}：海外營收年增 FY2028–FY2029（當地貨幣）", "%/年", lo, v, hi, "Assumed", "FY2023→FY2025 海外日圓年複合 11.9%（Derived；含匯率）", "")
for s, v, lo, hi in (("A", 0.09, None, None), ("B", 0.07, 0.05, 0.09), ("C", 0.04, None, None)):
    inp(f"golt_{s}", f"情境 {s}：海外營收年增 FY2030–FY2032（當地貨幣）", "%/年", lo, v, hi, "Assumed", "線 2a 全球工業機器人裝機 2026→2032 年複合約 6.6%（對帳頁；方向對照，不驅動）", "")
for s, v, lo, hi, n in (("A", 0.02, None, None, "樂觀：日圓再走弱約 2%"), ("B", 0.0, -0.04, 0.03, "基準：現價日匯率持平（USDJPY 159.71 ≈ FY2026 平均）"),
                        ("C", -0.07, None, None, "保守：日圓升值回到約 FY2025 上半水準（USD 146）")):
    inp(f"fx27_{s}", f"情境 {s}：FY2027 海外匯率換算效果（FY2028 起持平）", "%", lo, v, hi, "Assumed", n, "敏感度：營業利益對 USD 1 円約 13 億、EUR 1 円約 5 億、CNY 0.1 円約 10 億（" + FRH1 + "，Interested-party）")
for s, v, lo, hi in (("A", 0.860, None, None), ("B", 0.855, 0.845, 0.862), ("C", 0.840, None, None)):
    inp(f"gm_{s}", f"情境 {s}：毛利率 FY2027 起", "%", lo, v, hi, "Assumed", "2Q FY2026 單季 86.2%（" + FRH1 + "）；FY2023–FY2025 83.0–83.8%", "C：量縮與產品組合回到 FY2025 水準附近")
inp("e_sga", "營業費用彈性（販管費增速 ＝ 彈性 × 營收增速 ＋ (1 − 彈性) × 底線增速）", "x", 0.5, 0.7, 0.9, "Assumed",
    "FY2023→FY2025 販管費年複合 10.4%、營收 9.9%（Derived）：近 1 倍；FY2026 上半 販管費 +24.1% vs 營收 +34.9%（槓桿）", "直銷人員為主的費用結構，降速時費用不隨營收等比例下降")
inp("sga_fl", "營業費用底線增速（人員與據點擴張）", "%/年", 0.04, 0.06, 0.08, "Assumed", "未取得公司人員增速揭露", "")

sec("D. 損益表下半段")
inp("nonop", "業外收支淨額（受取利息、持分法、匯差）FY2027 起每年；FY2026＝上半實績＋本值 ÷ 2", "¥B", 20.0, 32.0, 45.0, "Assumed",
    "FY2025 40.0（含匯差）、FY2024 11.2、FY2023 24.3（Hist，Derived）；上半 FY2026 16.1", "現金與有價證券約 3.2 兆日圓 × 利率；不隨現金成長外推")
inp("tax_r", "有效稅率", "%", 0.29, 0.30, 0.31, "Derived", "FY2025 190.570 ÷ 635.756 ＝ 30.0%；上半 FY2026 29.6%（" + ER26 + "、" + ERH1 + "）", "")
inp("dps_26", "FY2026 每股股利（分割後；中間 325 円 ÷ 5 ＋ 期末 65 円予想）", "¥", None, "=325/split_r+65", None, "Interested-party",
    ERH1 + "：中間 325 円（分割前）、期末予想 65 円（分割後；不考慮分割為 325 円、年間 650 円）", FRH1 + "：配當總額予想 1,576 億円", CO)
inp("pay_r", "FY2027 起配息率（股利 ÷ 淨利）", "%", None, 0.30, None, "Derived", "FY2025 550 ÷ 1,835.63 ＝ 30.0%（" + ER26 + "）", "只用於收支關卡；不進 DCF")

sec("E. 估值輸入（評價日 2027-03-20＝FY2026 期末；資產負債表 2026-09-20）")
inp("px", "股價（2026-11-20 收盤，分割調整後）", "¥", None, 14025, None, "Verified", YCD + "：6861.T 2026-11-20 收盤 14,025", "分割前口徑約 70,125（＝14,025 × 5；Derived）。11-13 收盤 14,916、11-19 14,690（同檔）")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("val_date", "評價日（FY2026 期末）", "日期", None, "=DATE(2027,3,20)", None, "Derived", "會計年度末；DCF 明確期 FY2027–FY2032 由此折現", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, "=(tp_date-val_date)/365", None, "Derived", "2027-03-20 → 2027-11-20", "")
inp("cash_v", "現金及預金＋有價證券（2026-09-20）", "¥B", None, "=698.507+803.618", None, "Derived", ERH1 + "：現金及び預金 698,507、有価証券 803,618（Verified 原文相加）", "")
inp("inv_v", "投資有価証券（2026-09-20）", "¥B", None, 1691.841, None, "Verified", ERH1 + "：1,691,841", "內容（債券、存款、持分法投資）明細未取得")
inp("inv_k", "投資有価証券計入比例", "x", 0.8, 1.0, None, "Assumed", "基準按帳面全額；低案 80%（長期債券與持分法投資的市價、流動性折價未取得）", "")
inp("debt_v", "有息負債與租賃負債", "¥B", None, 0.0, None, "Verified", ERH1 + "：無借入金等有息負債科目；固定負債合計 11,853", "日本基準營業租賃不入表（未取得金額）")
inp("stub_frac", "Stub 期 FCFF 比例（2026-09-20 → 評價日 2027-03-20，FY2026 FCFF 的一半）", "x", None, 0.5, None, "Derived", "半年 ÷ 一年；上半現金流量表未取得，以全年 FCFF 的一半代", "")
inp("stub_div", "Stub 期支付的股利（FY2026 中間 325 円 × 分割前流通股；2026-12 支付）", "¥B", None, "=325*(sh_iss-sh_trs)/1000", None, "Derived", ERH1 + "：中間配當 325 円（分割前）", "")
inp("rf", "無風險利率（日本 10 年期國債）", "%", None, 0.02817, None, "Verified", JGB10, "與現價日相差逾 2 個交易日：逾期（跨模型一致，下次全面改版更新）")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "日本國家風險溢酬", "%", None, 0.0091, None, "Verified", "Damodaran 國家風險溢酬（2026-01-05）：日本 0.91%（與 Nabtesco v0.1.1 同值）", "")
inp("beta", "β（兩來源 Blume 調整後中位數）", "x", "=E{r}-0.3", "=MEDIAN(0.67*1.02+0.33,0.67*MEDIAN(1.48,1.14,1.70,0.85,0.83,0.98)+0.33)", "=E{r}+0.3", "Derived",
    SA + "：Keyence 5 年 β 1.02（2026-10-23）；同業 6 家 5 年 β 中位數 1.06（Cognex 1.48、Omron 1.14、Novanta 1.70、Hexagon 0.85、SMC 0.83、Fanuc 0.98）", "Blume 0.67β＋0.33；敏感度 ±0.3（Assumed）")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("fc_da", "折舊攤銷 ÷ 營收", "%", 0.012, 0.015, 0.018, "Derived", "FY2025 17.227 ÷ 1,169.289 ＝ 1.5%；FY2024 1.4%（" + ER26 + "）", "")
inp("fc_cx", "資本支出 ÷ 營收", "%", 0.012, 0.020, 0.030, "Assumed", "有形固定資產取得：FY2025 28.371（2.4%）、FY2024 14.342（1.4%）、FY2023 12.492（1.3%）", "補足資料「設備投資額」FY2025 四季合計含無形；FY2026 上半 100 億（" + FRH1 + "）")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.30, 0.38, 0.45, "Derived", "（應收＋存貨 − 應付）÷ 營收：FY2025 38.8%、FY2024 36.2%、FY2023 38.4%（Hist）", "以營運資金占營收比率代增量比率")
inp("dps_tgt", "評價日到目標時點之每股股利（分割後）", "¥", 120, 140, 160, "Assumed", "FY2026 期末 65 円（基準日 2027-03-20，6 月支付）＋FY2027 中間（假設約 75 円）", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("F. 同業前瞻 P/E（Comparable；同業 6 家，少於 8 家不做迴歸；csv 有舊報價日收盤者換算至 2026-11-20，其餘報價日未取得）")
PEERS = [("pe_cgnx", "Cognex CGNX（機器視覺）", 31.30, "Interested-party", f"{SA}：2026-10-30 收盤頁面值；報價日未取得 2026-11-20（csv 無此檔）", "VM"),
         ("pe_omron", "Omron 6645（感測、控制）", 26.57, "Interested-party", f"{SA}：2026-09-18 頁面值；報價日未取得 2026-11-20（csv 無此檔）", "VM"),
         ("pe_hexa", "Hexagon HEXA-B（量測、視覺）", 27.28, "Interested-party", f"{SA}：2026-11-13 頁面值；報價日未取得 2026-11-20（csv 無此檔）", "VM"),
         ("pe_novt", "Novanta NOVT（精密光學、視覺元件）", "=29.67*101.99/118.88", "Derived", f"{SA}：2026-11-13 頁面值 29.67 × 2026-11-20 收盤 101.99 ÷ 2026-11-13 收盤 118.88（{YCD}；分母 EPS 不變）", "VM"),
         ("pe_smc", "SMC 6273（氣動 FA 零組件）", "=15.81*55300/60270", "Derived", f"{TRD}：PER 15.81 × 2026-11-20 收盤 55,300 ÷ 2026-11-16 收盤 60,270（分母 EPS 不變）", "FA"),
         ("pe_fanuc", "Fanuc 6954（CNC、機器人）", "=21.65*5047/5343", "Derived", f"{TRD}：PER 21.65 × 2026-11-20 收盤 5,047 ÷ 舊收盤 5,343（分母＝公司 FY2026 EPS 予想）", "FA")]
for k, lab, v, tag, src, g in PEERS:
    inp(k, f"前瞻 P/E：{lab}", "x", None, v, None, tag, src, f"群組 {g}（VM＝視覺與量測、FA＝工廠自動化零組件）；會計口徑：美國 US GAAP、日本 日本基準或 IFRS、瑞典 IFRS；P/E 為股權口徑不受租賃處理影響", DV)
inp("pe_med", "同業前瞻 P/E 中位數（6 家；進 Comparable）", "x", None, "=MEDIAN(pe_cgnx,pe_omron,pe_hexa,pe_novt,pe_smc,pe_fanuc)", None, "Derived", "少於 8 家，不做迴歸（第 10 節第 7 條）", "")
inp("pe_vm", "子群中位數：視覺與量測 4 家（純展示）", "x", None, "=MEDIAN(pe_cgnx,pe_omron,pe_hexa,pe_novt)", None, "Derived", "第 10 節第 18 條：依商業模式並列子群中位數", "")
inp("pe_fa", "子群中位數：FA 零組件 2 家（純展示）", "x", None, "=MEDIAN(pe_smc,pe_fanuc)", None, "Derived", "", "")
inp("cons_eps26", "共識 FY2026（2027年3月期）EPS（分割後）", "¥", None, 491.2, None, "Interested-party", MS + "：EPS 491.2（FY2025 367.1＝1,835.63 ÷ 5，確認為分割後口徑）", "對照用（不進目標價）；財報後（10-29 中間決算後）", DV)
inp("pe_self", "Keyence 現行前瞻 P/E（現價 ÷ 共識 FY2026 EPS）", "x", None, "=px/cons_eps26", None, "Derived", "＝14,025 ÷ 491.2；資料商頁面 2027E P/E 28.6 倍一致", "對照：traders.co.jp 2026-11-16 PER 31.25（分割前口徑）", DV)
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, "=pe_self/pe_med", "Assumed", "1.00（未調整：缺 Keyence 與同業的 10 年前瞻 P/E 歷史；tools/data/fwd_mult_hist.csv 無 6861）",
    "高案＝現行 Keyence P/E ÷ 同業中位數（維持現行溢價）；第 10 節第 18、47 條：列為承重輸入")
inp("cons_tp", "賣方平均目標價（分割後；對照用，不進目標價）", "¥", None, 19906.25, None, "Interested-party", MS + "：16 位分析師平均 19,906.25、評等 BUY（2026-11-20）", "2026-07-17 株予報 平均 91,142（分割前）、2026-07-30 Simply Wall St 92,250（分割前，16 位）", DV)
for k, lab, v, n in (("cons_rev26", "共識 FY2026 營收", 1502.586, "EBIT 818.088、淨利 598.809"), ("cons_op26", "共識 FY2026 營業利益", 818.088, ""),
                     ("cons_rev27", "共識 FY2027 營收", 1663.559, "EBIT 898.104、淨利 653.537、EPS 536.1"), ("cons_op27", "共識 FY2027 營業利益", 898.104, ""),
                     ("cons_eps27", "共識 FY2027 EPS（分割後）", 536.1, "FY2028 EPS 587.4")):
    inp(k, lab + "（對照用，不進目標價）", "¥B" if "EPS" not in lab else "¥", None, v, None, "Interested-party", MS + "：" + (n or "同頁"), "財報後；" + KY + " 的 FY2026 共識（營收 1,329.7）為第 1 季決算前，只作歷史對照", DV)

sec("G. 實體 AI 曝險（不進目標價；只用於實體AI 頁）")
inp("kshr", "Keyence 占「五線歸因」工業感測與機器視覺市場份額（基準）", "%", None, 0.25, None, "Assumed", "未取得（線 6 份額頁 F18「未取得」）", "區間見 kshr_lo／kshr_hi；只用於實體AI 頁")
inp("kshr_lo", "同上（低端）", "%", None, 0.10, None, "Assumed", "Cognex、Omron、Sick、Basler 等分食", "只用於實體AI 頁")
inp("kshr_hi", "同上（高端）", "%", None, 0.40, None, "Assumed", "感測器（光電、位移）Keyence 一般被視為日本龍頭；全球份額未取得", "只用於實體AI 頁")
inp("usdjpy", "美元兌日圓（2026-11-20）", "¥/US$", None, 159.707001, None, "Verified", YCD + "：JPY=X 2026-11-20 收盤 159.707", "")


def build(path, ov=None, sens=None, rev=None):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 ¥B、股數百萬股（分割後）、EPS 與股價日圓（分割後）"
    IN["A2"] = "標記：Verified（附出處）／Interested-party（寫明是誰、誘因欄）／Analogy（寫明類比對象）／Assumed（給區間）／Derived（寫明換算）；查不到寫「未取得」。股票分割 1 拆 5（2026-11-21 生效）：股數、EPS、每股股利全部為分割後口徑"
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
            if isinstance(v, str): v = v.replace("{r}", str(r))
            c = IN.cell(r, col, v)
            if col == 5: c.font = BLUE
            if u == "日期" and col == 5: c.number_format = "yyyy-mm-dd"
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2); IN.cell(r, 10, inc)
        dn(k, f"In!$E${r}"); KR[k] = r; r += 1
    last_in = r
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("gd27", "國內營收年增 FY2027"), ("gdm", "國內營收年增 FY2028–FY2029"), ("gdlt", "國內營收年增 FY2030–FY2032"),
           ("go27", "海外營收年增 FY2027（當地貨幣）"), ("gom", "海外營收年增 FY2028–FY2029"), ("golt", "海外營收年增 FY2030–FY2032"),
           ("fx27", "FY2027 匯率換算效果"), ("gm", "毛利率 FY2027 起")]
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
    H["A1"] = "Hist — 歷史財務（年度；日本基準；¥B；EPS、股利、股數為分割前原文）。公式只引用本頁儲存格（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2023（2024年3月期）", "FY2024（2025年3月期）", "FY2025（2026年3月期）", "標記", "來源"])
    HR = [("h_rev", "營收（売上高）", 967.288, 1059.145, 1169.289),
          ("h_dom", "國內營收", 344.567, 372.753, 390.066),
          ("h_ovs", "海外營收", 622.720, 686.391, 779.222),
          ("h_cogs", "營業成本（売上原価）", 164.782, 171.444, 198.552),
          ("h_sga", "營業費用（販売費及び一般管理費）", 307.491, 337.925, 374.978),
          ("h_op", "營業利益", 495.014, 549.775, 595.759),
          ("h_noi", "營業外收益", 24.969, 16.320, 40.491),
          ("h_noe", "營業外費用", 0.688, 5.085, 0.494),
          ("h_pbt", "稅前利益（税金等調整前当期純利益；無特別損益）", 519.295, 561.010, 635.756),
          ("h_tax", "所得稅（法人税等合計）", 149.653, 162.353, 190.570),
          ("h_ni", "歸屬母公司淨利（無非支配股東損益）", 369.642, 398.656, 445.185),
          ("h_eps", "EPS（分割前，日圓）", 1524.14, 1643.77, 1835.63),
          ("h_dps", "每股股利（分割前，日圓）", 300, 350, 550),
          ("h_da", "折舊攤銷", 13.767, 15.193, 17.227),
          ("h_cx", "有形固定資產取得", 12.492, 14.342, 28.371),
          ("h_rec", "受取手形及び売掛金", 307.526, 320.312, 379.590),
          ("h_inv", "棚卸資産", 77.613, 77.892, 85.273),
          ("h_ap", "支払手形及び買掛金", 13.725, 14.890, 10.890),
          ("h_ocf", "營業現金流", 387.916, 409.522, 430.680),
          ("h_div", "股利支付總額（現金流量表）", 72.757, 78.820, 109.136)]
    for i, (k, lab, a, b, c) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b, c)): H.cell(i, 2 + j, v).font = BLUE
        H.cell(i, 5, "Verified"); H.cell(i, 6, (ER25 + "（FY2023 為比較欄）；" if True else "") + ER26)
        for j, y in enumerate((23, 24, 25)): dn(f"{k}_{y}", f"Hist!${'BCD'[j]}${i}")
    i = 3 + len(HR)
    H.cell(i, 1, "平均股數（分割前，百萬股）＝淨利 ÷ EPS（Derived）")
    for j, y in enumerate((23, 24, 25)): H.cell(i, 2 + j, f"=h_ni_{y}/h_eps_{y}*1000"); dn(f"h_sh_{y}", f"Hist!${'BCD'[j]}${i}")
    H.cell(i, 5, "Derived"); H.cell(i, 6, "FY2025 原文 242,525,222 股、FY2024 242,525,467 股（" + ER26 + "）")
    H.cell(i + 1, 1, "營運資金 ÷ 營收（Derived）")
    for j, y in enumerate((23, 24, 25)): H.cell(i + 1, 2 + j, f"=(h_rec_{y}+h_inv_{y}-h_ap_{y})/h_rev_{y}")
    H.cell(i + 1, 5, "Derived")
    H.cell(i + 2, 1, "FY2025 上半（2025-03-21～09-20）營收 545.300、營業利益 272.179、淨利 199.959（" + ERH1 + " 前年同期欄）；FY2025 下半＝全年 − 上半（Derived）")
    H.cell(i + 3, 1, "未取得：分地區（北中南米、アジア、欧州）年度營收金額（補足資料只揭露增速與單季構成比）；研究開發費年度合計；人員數")
    H.column_dimensions["A"].width = 46

    # ---------------- Q ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 季度損益（¥B；上半與第 1 季為 Verified 原文，第 2 季＝上半 − 第 1 季，Derived）"
    hdr(Q, 2, ["項目", "Q1'25", "Q2'25", "H1'25", "Q1'26", "Q2'26", "H1'26", "標記", "來源", "來源誘因"])
    qd = [("營收", 261.076, 545.300, 346.621, 735.399), ("營業成本", 45.412, 94.422, 51.866, 105.433), ("營業費用（販管費）", 86.363, 178.698, 107.679, 221.703),
          ("營業利益", 129.301, 272.179, 187.076, 408.262), ("歸屬母公司淨利", 92.116, 199.959, 139.133, 298.675)]
    for i, (lab, q1a, h1a, q1b, h1b) in enumerate(qd, 3):
        Q.cell(i, 1, lab)
        for col, v in ((2, q1a), (4, h1a), (5, q1b), (7, h1b)): Q.cell(i, col, v).font = BLUE
        Q.cell(i, 3, f"=D{i}-B{i}"); Q.cell(i, 6, f"=G{i}-E{i}")
        Q.cell(i, 8, "Verified"); Q.cell(i, 9, ERQ1 + "；" + ERH1); Q.cell(i, 10, "")
    i = 3 + len(qd)
    Q.cell(i, 1, "營業利益率"); [Q.cell(i, c, f"={L}6/{L}3") for c, L in zip(range(2, 8), "BCDEFG")]; Q.cell(i, 8, "Derived"); Q.cell(i, 9, "＝營業利益 ÷ 營收")
    Q.cell(i + 1, 1, "毛利率"); [Q.cell(i + 1, c, f"=1-{L}4/{L}3") for c, L in zip(range(2, 8), "BCDEFG")]; Q.cell(i + 1, 8, "Derived"); Q.cell(i + 1, 9, "＝1 − 營業成本 ÷ 營收")
    Q.cell(i + 2, 1, "檢查：營收 − 營業成本 − 營業費用 − 營業利益（應為 0）"); [Q.cell(i + 2, c, f"=ROUND({L}3-{L}4-{L}5-{L}6,3)") for c, L in zip(range(2, 8), "BCDEFG")]
    Q.cell(i + 2, 8, "Derived"); Q.cell(i + 2, 9, "")
    Q.cell(i + 3, 1, "公司補足資料單季：Q2'26 營收 3,888 億（+36.8%）、營業利益 2,212 億（+54.8%）、毛利率 86.2%、營業利益率 56.9%；Q2'26 構成比 國內 31.9%、海外 68.1%（北中南米 35.3%、アジア 46.1%、欧州・その他 18.5%）（" + FRH1 + "）")
    Q.cell(i + 4, 1, "地區別單季營收年增（円／現地通貨）Q2'26：國內 +22.8%；海外 +44.5%／+30.5%；北中南米 +38.9%／+24.9%；アジア +61.4%／+45.2%；欧州・その他 +22.1%／+12.1%")
    for L, w in zip("AI", (44, 60)): Q.column_dimensions[L].width = w

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境標準損益表（日本基準＝本表的 GAAP 口徑；¥B；EPS 日圓、分割後）：國內＋海外 → 單一官方分部（電子應用機器）→ 損益 → 現金流。FY2023–FY2025 實績，FY2026＝上半實績＋下半模型，FY2027 起模型"
    PR = {}; rr = 3
    LAB = [("dom", "國內營收（子列）"), ("ovs", "海外營收（子列）"), ("rev", "合併營收（單一分部：電子應用機器）"), ("chk", "檢查：國內＋海外 − 合併營收（應為 0）"),
           ("gr", "營收年增率"), ("ovsr", "海外營收占比"), ("cogs", "營業成本（売上原価）"), ("gp", "毛利"), ("gm", "毛利率"),
           ("sga", "營業費用（販売費及び一般管理費）"), ("sgar", "營業費用率"), ("op", "營業利益"), ("opm", "營業利益率"),
           ("chk2", "檢查：毛利 − 營業費用 − 營業利益（應為 0）"), ("nonop", "業外收支淨額（其他收入淨額）"), ("pbt", "稅前利益"), ("tax", "所得稅"),
           ("ni", "稅後淨利（歸屬母公司）"), ("nim", "淨利率"), ("nig", "淨利年增率"), ("sh", "平均股數（百萬股，分割後）"), ("eps", "EPS（日圓，分割後；無潛在股份）"),
           ("chk3", "檢查：營業利益＋業外 − 稅前利益（應為 0）"),
           ("da", "折舊攤銷"), ("cx", "資本支出"), ("dwc", "營運資金增加"), ("fcff", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）"), ("dps", "每股股利（日圓，分割後）"),
           ("div", "股利"), ("fcfad", "自由現金流扣股利（收支關卡用）")]
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, (k, _) in enumerate(LAB)}; R["hdr"] = rr + 1
        for k, l in LAB: P.cell(R[k], 1, l)
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1)
            cell = lambda k: f"{c}{R[k]}"; prev = lambda k: f"{p}{R[k]}"
            def w(k, f): P[cell(k)] = f
            if y <= 2025:
                t = str(y)[2:]
                for k, nm in (("dom", "h_dom"), ("ovs", "h_ovs"), ("cogs", "h_cogs"), ("sga", "h_sga"), ("op", "h_op"), ("tax", "h_tax"), ("ni", "h_ni"), ("da", "h_da"), ("cx", "h_cx"), ("pbt", "h_pbt")):
                    w(k, f"={nm}_{t}")
                w("rev", f"=h_rev_{t}"); w("nonop", f"=h_noi_{t}-h_noe_{t}")
                w("sh", f"=h_sh_{t}*split_r"); w("eps", f"=h_eps_{t}/split_r"); w("dps", f"=h_dps_{t}/split_r"); w("div", f"=h_div_{t}")
                if y >= 2024:
                    pt = str(y - 1)[2:]
                    w("dwc", f"=(h_rec_{t}+h_inv_{t}-h_ap_{t})-(h_rec_{pt}+h_inv_{pt}-h_ap_{pt})")
                    w("fcff", f"={cell('op')}*(1-tax_r)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
                    w("fcfad", f"={cell('fcff')}-{cell('div')}")
            elif y == 2026:
                w("dom", "=h_dom_25*(1+gd26)"); w("ovs", "=h_ovs_25*(1+go26)*(1+fx26)")
                w("cogs", f"={cell('rev')}*(1-gm26)"); w("sga", f"={cell('rev')}*sga26_r"); w("op", f"={cell('gp')}-{cell('sga')}")
                w("nonop", "=h1_nonop+nonop/2"); w("sh", "=sh_26"); w("dps", "=dps_26")
            else:
                g_d = f"gd27_{s}" if y == 2027 else (f"gdm_{s}" if y <= 2029 else f"gdlt_{s}")
                g_o = f"go27_{s}" if y == 2027 else (f"gom_{s}" if y <= 2029 else f"golt_{s}")
                fx = f"*(1+fx27_{s})" if y == 2027 else ""
                w("dom", f"={prev('dom')}*(1+{g_d})"); w("ovs", f"={prev('ovs')}*(1+{g_o}){fx}")
                w("cogs", f"={cell('rev')}*(1-gm_{s})")
                w("sga", f"={prev('sga')}*(1+e_sga*{cell('gr')}+(1-e_sga)*sga_fl)"); w("op", f"={cell('gp')}-{cell('sga')}")
                w("nonop", "=nonop"); w("sh", "=sh_v"); w("dps", f"=pay_r*{cell('eps')}")
            if y >= 2026:
                w("pbt", f"={cell('op')}+{cell('nonop')}"); w("tax", f"={cell('pbt')}*tax_r"); w("ni", f"={cell('pbt')}-{cell('tax')}")
                w("eps", f"={cell('ni')}/{cell('sh')}*1000")
                w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}"); w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
                w("fcff", f"={cell('op')}*(1-tax_r)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
                w("div", f"={cell('dps')}*{'sh_v' if y >= 2026 else cell('sh')}/1000"); w("fcfad", f"={cell('fcff')}-{cell('div')}")
                w("rev", f"={cell('dom')}+{cell('ovs')}")
            w("chk", f"=ROUND({cell('dom')}+{cell('ovs')}-{cell('rev')},6)")
            w("gp", f"={cell('rev')}-{cell('cogs')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
            w("sgar", f"={cell('sga')}/{cell('rev')}"); w("opm", f"={cell('op')}/{cell('rev')}"); w("ovsr", f"={cell('ovs')}/{cell('rev')}")
            w("chk2", f"=ROUND({cell('gp')}-{cell('sga')}-{cell('op')},6)"); w("chk3", f"=ROUND({cell('op')}+{cell('nonop')}-{cell('pbt')},6)")
            w("nim", f"={cell('ni')}/{cell('rev')}")
            if p:
                w("gr", f"={cell('rev')}/{prev('rev')}-1"); w("nig", f"={cell('ni')}/{prev('ni')}-1")
        P.cell(R["dom"], 2, "FY2026＝FY2025×(1＋gd26)；FY2027 起依情境國內增速")
        P.cell(R["ovs"], 2, "×(1＋當地貨幣增速)×(1＋匯率換算效果；只在 FY2026、FY2027)")
        P.cell(R["sga"], 2, "FY2026＝營收 × sga26_r；FY2027 起＝前一年 ×(1＋彈性×營收增速＋(1−彈性)×底線增速)")
        P.cell(R["nonop"], 2, "FY2026＝上半實績＋nonop ÷ 2")
        P.cell(R["eps"], 2, "FY2023–FY2025＝原文 EPS ÷ split_r（Derived）")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 ×(1 − tax_r)；FY2024–FY2025 為同式回算（Derived）")
        P.cell(R["div"], 2, "FY2023–FY2025＝現金流量表支付額；FY2026 起＝宣告每股股利 × 評價日股數")
        for k in ("gr", "ovsr", "gm", "sgar", "opm", "nim", "nig"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(LAB) + 2
    P.column_dimensions["A"].width = 46; P.column_dimensions["B"].width = 34
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "fcfad"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2027-03-20，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（¥B）"; D["B4"] = "=px*sh_v/1000"
    D["A5"] = "負債 D（有息負債＋租賃）"; D["B5"] = "=debt_v"
    D["A6"] = "稅後 kd（無負債，不適用）"; D["B6"] = "=rf*(1-tax_r)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金與投資（2026-09-20）＝現金及預金＋有價證券 ＋ 投資有価証券 × inv_k − 有息負債"; D["B10"] = "=cash_v+inv_k*inv_v-debt_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND(((px*sh_v/1000)*(rf+beta*(erp+crp))+debt_v*rf*(1-tax_r))/((px*sh_v/1000)+debt_v)+wacc_adj-wacc,9)"
    for r_ in range(3, 11): D[f"B{r_}"].number_format = "0.00%" if r_ in (3, 6, 7, 8, 9) else "#,##0.0"
    DR = {}; rr = 13
    for s in SC:
        R = PR[s]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}" for y in range(2027, 2033)]); b = rr + 2
        rows = ["EBIT（營業利益）", "NOPAT", "＋折舊攤銷", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF（應為 0）", "折現年數（期中）", "折現因子", "FCFF 現值"]
        for i, l in enumerate(rows): D.cell(b + i, 1, l)
        for j, y in enumerate(range(2027, 2033)):
            c = "BCDEFG"[j]; pc = COL[y]
            D[f"{c}{b}"] = f"=PL!{pc}{R['op']}"; D[f"{c}{b+1}"] = f"={c}{b}*(1-tax_r)"
            D[f"{c}{b+2}"] = f"=PL!{pc}{R['da']}"; D[f"{c}{b+3}"] = f"=-PL!{pc}{R['cx']}"; D[f"{c}{b+4}"] = f"=-PL!{pc}{R['dwc']}"
            D[f"{c}{b+5}"] = f"=PL!{pc}{R['fcff']}"; D[f"{c}{b+6}"] = f"=ROUND(SUM({c}{b+1}:{c}{b+4})-{c}{b+5},9)"
            D[f"{c}{b+7}"] = j + 0.5; D[f"{c}{b+8}"] = f"=1/(1+wacc)^{c}{b+7}"; D[f"{c}{b+9}"] = f"={c}{b+5}*{c}{b+8}"
        t = b + 11
        items = [("明確期現值合計", f"=SUM(B{b+9}:G{b+9})"),
                 ("NOPAT FY2033＝NOPAT FY2032 ×(1＋g)", f"=G{b+1}*(1+g_tv)"),
                 ("終值（FY2032 年底）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^6"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金與投資（2026-09-20）", "=dcf_nc"),
                 ("＋Stub 期 FCFF 與稅後業外（FY2026 FCFF × stub_frac ＋ 下半業外 ×(1 − 稅率)）", f"=PL!F{R['fcff']}*stub_frac+nonop/2*(1-tax_r)"),
                 ("− Stub 期支付的中間股利（2026-12）", "=-stub_div"),
                 ("股權價值", f"=B{t+4}+B{t+5}+B{t+6}+B{t+7}"),
                 ("每股價值（評價日，分割後）", f"=MAX(0,B{t+8})/sh_v*1000"),
                 ("DCF 情境價值推到目標時點（×(1＋ke)^t − 期間股利）", f"=B{t+9}*(1+ke)^t_tgt-dps_tgt"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+11}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=B{t+2}/G{b+1}"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到目標時點）", f"=((B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6)+B{t+5}+B{t+6}+B{t+7})/sh_v*1000*(1+ke)^t_tgt-dps_tgt")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 10, ps=t + 9, ev=t + 4, tvs=t + 11, alt=t + 14)
        dn(f"dcf_{s}", f"DCF!$B${t+10}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 64

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價：DCF（權重 w_dcf）＋ Comparable（同業前瞻 P/E 中位數 × 目標時點 NTM EPS，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}")
        V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到目標時點", [f"=dcf_{s}" for s in SC], "DCF 頁")
    r_ntm = vrow("目標時點 NTM EPS（4/12 × FY2027 ＋ 8/12 × FY2028；分割後）", [f"=4/12*PL!G{PR[s]['eps']}+8/12*PL!H{PR[s]['eps']}" for s in SC], "目標時點 2027-11-20 起算 12 個月：FY2027 剩 4 個月、FY2028 前 8 個月")
    r_cu = vrow("Comparable（未調整）：同業前瞻 P/E 中位數 × NTM EPS", [f"=pe_med*{L}{r_ntm}" for L in "BCD"], "同業 6 家中位數 pe_med")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（日圓，分割後）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    r_px = one("現價（In 頁 px，分割調整後）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（¥B；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=TEXT(px*(1+IF(B{r_up}>0.5,9,IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5))))),"#,##0")&"／"&TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,IF(B{r_up}>=-0.5,-0.5,-9))))),"#,##0")', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2027-03-20 → 目標時點，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B 目標時點 NTM EPS", f"=B{r_tp}/C{r_ntm}", "")
    r_pe2 = one("現行前瞻 P/E（現價 ÷ 共識 FY2026 EPS）÷ 同業中位數", "=pe_self/pe_med", "顯示用：Keyence 對同業的現行溢價")
    r_dsh = one("DCF 在加權目標價的占比", f"=w_dcf*B{r_dcft}/B{r_tp}")
    r_cash = one("每股淨現金與投資（分割後，日圓）", "=dcf_nc/sh_v*1000", "現金及有價證券＋投資有価証券；占現價比例見下一列")
    r_cashp = one("每股淨現金占現價", f"=B{r_cash}/px")
    for r_ in (r_up, r_a1, r_a2, r_a3, r_dsh, r_cashp): V[f"B{r_}"].number_format = "0.0%"
    for r_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{r_}"].number_format = "#,##0"
    for r_ in (r_tp, r_dcft, r_cmpt, r_cash): V[f"B{r_}"].number_format = "#,##0"
    V.column_dimensions["A"].width = 70; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, ca=r_ca, ntm=r_ntm, pd2=r_pd2, pe=r_pe, pe2=r_pe2, gj=r_gj, wk=r_wk, dsh=r_dsh, cash=r_cash, cashp=r_cashp)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與市場共識對照（公司不揭露業績預測）；FY2026 隱含下半年檢查；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "共識或實績", "差", "說明"])
    B_ = PR["B"]
    rc = [("FY2026 營收 vs 共識", f"=PL!F{B_['rev']}", "=cons_rev26", "共識 16 位（" + "MarketScreener，2026-11-20；中間決算後）"),
          ("FY2026 營業利益 vs 共識", f"=PL!F{B_['op']}", "=cons_op26", "同上"),
          ("FY2026 EPS（分割後） vs 共識", f"=PL!F{B_['eps']}", "=cons_eps26", "同上"),
          ("FY2027 營收 vs 共識", f"=PL!G{B_['rev']}", "=cons_rev27", "同上"),
          ("FY2027 營業利益 vs 共識", f"=PL!G{B_['op']}", "=cons_op27", "同上"),
          ("FY2027 EPS（分割後） vs 共識", f"=PL!G{B_['eps']}", "=cons_eps27", "同上"),
          ("FY2026 隱含下半營收（模型全年 − 上半實績） vs FY2025 下半實績", f"=PL!F{B_['rev']}-h1_rev", "=h_rev_25-545.3", "FY2025 下半＝全年 1,169.289 − 上半 545.300（Derived）"),
          ("FY2026 隱含下半營業利益 vs FY2025 下半實績", f"=PL!F{B_['op']}-h1_op", "=h_op_25-272.179", "上半 FY2025 營業利益 272.179"),
          ("加權目標價 vs 賣方平均目標價", "=tp", "=cons_tp", "賣方 16 位平均（Interested-party）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}/C{i}-1"); RC.cell(i, 5, n)
        RC[f"D{i}"].number_format = "+0.0%;-0.0%"
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_keyence.py 二分法；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 62; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_keyence.py 以 LibreOffice 重算每一案例）；案例值為常數快照"
    SE["A2"] = ("kshr、kshr_lo、kshr_hi、usdjpy 不列：只用於實體AI 頁的曝險估計，不進目標價。"
                "rel_k 高案＝現行 Keyence 前瞻 P/E ÷ 同業中位數（維持現行溢價）。")
    SE["A3"] = "build_id"; SE["D3"] = sens["build_id"] if sens else "（基準檔）"
    SE["A4"] = "In 頁輸入指紋（敏感度重算時）"; SE["D4"] = sens["fp"] if sens else 0
    SE["A5"] = "In 頁輸入指紋（本檔即時）"; SE["D5"] = f"=SUM(In!E4:E{last_in - 1})"
    SE["A6"] = "基準加權價（敏感度重算時）"; SE["D6"] = sens["tp0"] if sens else 0
    SE["A7"] = "模型狀態"; SE["D7"] = f'=IF(AND(ABS(D5-D4)<=1E-6*MAX(1,ABS(D4)),ABS(Val!B{r_tp}-D6)<1E-6),"CURRENT","STALE")'
    hdr(SE, 9, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    if sens:
        sr = 10
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]).number_format = "#,##0"; SE.cell(sr, 5, f"=D{sr}/Val!$B${r_tp}-1").number_format = "+0.0%;-0.0%"
            SE.cell(sr, 6, cs["dcf"]).number_format = "#,##0"; SE.cell(sr, 7, f"=F{sr}/Val!$B${r_dcft}-1").number_format = "+0.0%;-0.0%"
            SE.cell(sr, 8, f'=IF(D{sr}/px-1>0.5,"強力買進",IF(D{sr}/px-1>0.25,"買進",IF(D{sr}/px-1>=-0.25,"中立",IF(D{sr}/px-1>=-0.5,"放空","強力放空"))))')
            sr += 1
        SE.cell(sr + 1, 1, "註：相對基準欄以本檔 Val 頁加權目標價（即時）為分母；模型狀態 STALE 表示 In 頁改過、須重跑 build_keyence.py")
    SE.column_dimensions["A"].width = 50; SE.column_dimensions["B"].width = 34

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = "對帳 — 本模型讀取的線層級變數（線 6 共用零組件、線 2a 工業機器人成品；只讀，不在本模型改口徑）"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型是否一致"])
    rec = [("工業感測與機器視覺（FAS）市場 2026／2030／2032（基準）", "6", "502.4／626.7／674.1", "百萬美元", "五線歸因（工業＋協作機器人工作站）", CP + " 情境!D48、H48、J48（下游取數頁未列 FAS，改讀情境頁）", "Derived（上游 Assumed 每台內容）", "實體AI 頁 曝險估計", "一致（直接引用）"),
           ("FAS 市場 低／高（2030）", "6", "208.1／1,805.0", "百萬美元", "同上", CP + " 情境!H47、H49", "Derived", "實體AI 頁 曝險區間", "一致"),
           ("每台工作站感測視覺內容（傳統／協作；低／基準／高）", "6", "300／200、800／500、2,000／1,500", "美元/台", "Assumed", CP + " 輸入!D36:E36、D58:E58、D80:E80", "Assumed", "說明用", "一致"),
           ("人形、AMR、智慧車、Robotaxi 的 FAS 每台內容", "6", "0", "美元/台", "線模型假設（視覺以 VIS 相機另列）", CP + " 輸入!F58:L58", "Assumed", "實體AI 頁說明", "一致；本模型不另加人形曝險"),
           ("Keyence FAS 份額", "6", "未取得（本模型 Assumed 10%／25%／40%）", "%", "", CP + " 份額!F18", "Assumed", "In kshr、kshr_lo、kshr_hi", "線模型未取得；本模型自設區間"),
           ("全球工業機器人裝機 2025 實績、2026／2030／2032 基準", "2a", "60.3 萬；65.5 萬／85.4 萬／96.0 萬", "台", "IFR 新裝機（含協作）", IRX + " 下游取數 IR.inst.global.2025.base、.2026.base、.2030.base、.2032.base", "Verified／Derived", "實體AI 頁每台內容換算；海外長期增速方向對照", "一致（直接引用）"),
           ("全球工業機器人裝機 2026→2032 年複合（基準）", "2a", "約 6.6%", "%/年", "Derived", IRX + " 下游取數 IR.inst.global.2026.base → .2032.base", "Derived", "golt_B（7%）方向對照，不驅動", "不一致但有理由：Keyence 需求主體是一般產線自動化，非機器人台數；本模型不以裝機量驅動營收"),
           ("JARA 受注 YoY（2026 年 7–9 月）", "2a", "+46.6%", "%", "會員口徑", IRX + " 下游取數 IR.cycle.jara.orders_yoy.latest", "Verified", "情境 A（FA 投資上行）說明", "一致"),
           ("中國工業機器人產量 1–9 月 YoY", "2a", "+29.1%", "%", "國家統計局", IRX + " 下游取數 IR.cycle.cn_prod.ytd_yoy", "Verified", "go26 說明（アジア 上半 +39.5% 當地貨幣）", "一致")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 44; RE.column_dimensions["C"].width = 30; RE.column_dimensions["F"].width = 60

    # ---------------- 實體AI ----------------
    X = wb.create_sheet("實體AI")
    X["A1"] = "實體AI — 機器人工作站相關的感測視覺營收曝險（Assumed，不進目標價）：線 6 FAS 市場（五線歸因）× Keyence 份額 × 匯率 → 十億日圓，對照情境 B 合併營收"
    hdr(X, 2, ["項目", "單位", "來源"] + [f"FY{y}" for y in range(2026, 2033)])
    fas = {"lo": [179.18, 186.005049401046, 193.090068102945, 200.444958457081, 208.08, 212.003337237884, 216],
           "b": [502.385, 530.938057613278, 561.105276797501, 592.977343186269, 626.65, 649.947296324864, 674.1],
           "hi": [1319.2, 1426.82232133103, 1543.1832036411, 1668.98832769768, 1805, 1932.05706437465, 2068]}
    inst = [655010.4099, 701909.15524884, 752165.850764657, 806020.925679407, 854382.181220171, 905645.112093381, 959983.818818984]
    XR = {}
    rows_x = [("fas_lo", "FAS 市場（低）", "百萬美元", CP + " 情境!D47:J47（年 2026–2032；Derived）"),
              ("fas_b", "FAS 市場（基準）", "百萬美元", CP + " 情境!D48:J48"),
              ("fas_hi", "FAS 市場（高）", "百萬美元", CP + " 情境!D49:J49"),
              ("inst", "全球工業機器人裝機（基準）", "台", IRX + " 下游取數 IR.inst.global.<年>.base")]
    for i, (k, lab, u, src) in enumerate(rows_x, 3):
        X.cell(i, 1, lab); X.cell(i, 2, u); X.cell(i, 3, src); XR[k] = i
        vals = fas[k[4:]] if k.startswith("fas") else inst
        for j, v in enumerate(vals): X.cell(i, 4 + j, v).font = BLUE
    i = 7
    X.cell(i, 1, "隱含每台裝機 FAS 內容（基準）"); X.cell(i, 2, "美元/台"); X.cell(i, 3, "＝FAS 市場 ÷ 裝機（Derived；線模型以工作站計，口徑近似）")
    for j in range(7): L = "DEFGHIJ"[j]; X.cell(i, 4 + j, f"={L}4*1000000/{L}6")
    defs = [("exp_lo", "Keyence 機器人相關感測視覺營收（低：低市場 × 份額低端）", f"{{L}}3*kshr_lo*usdjpy/1000"),
            ("exp_b", "Keyence 機器人相關感測視覺營收（基準：基準市場 × 份額基準）", f"{{L}}4*kshr*usdjpy/1000"),
            ("exp_hi", "Keyence 機器人相關感測視覺營收（高：高市場 × 份額高端）", f"{{L}}5*kshr_hi*usdjpy/1000")]
    for n, (k, lab, f) in enumerate(defs):
        rr_ = 8 + n; X.cell(rr_, 1, lab); X.cell(rr_, 2, "¥B"); X.cell(rr_, 3, "Assumed（份額未取得）；匯率固定 2026-11-20"); XR[k] = rr_
        for j in range(7): L = "DEFGHIJ"[j]; X.cell(rr_, 4 + j, "=" + f.replace("{L}", L))
    X.cell(11, 1, "情境 B 合併營收"); X.cell(11, 2, "¥B"); X.cell(11, 3, "PL 頁情境 B")
    for j, y in enumerate(range(2026, 2033)): X.cell(11, 4 + j, f"=PL!{COL[y]}{PR['B']['rev']}")
    for n, (k, lab) in enumerate((("shr_lo", "占合併營收（低）"), ("shr_b", "占合併營收（基準）"), ("shr_hi", "占合併營收（高）"))):
        rr_ = 12 + n; X.cell(rr_, 1, lab); X.cell(rr_, 2, "%"); X.cell(rr_, 3, "Assumed"); XR[k] = rr_
        for j in range(7):
            L = "DEFGHIJ"[j]; c = X.cell(rr_, 4 + j, f"={L}{8 + n}/{L}11"); c.number_format = "0.0%"
    X.cell(16, 1, "讀法：線 6 把每台機器人工作站的感測與視覺內容設為傳統 800 美元、協作 500 美元（Assumed，區間 200–2,000），人形、AMR、車用為 0（視覺以相機模組另列）。"
                  "即使 Keyence 拿下全部 FAS 市場，2030 年基準也只約占營收數個百分點；Keyence 的需求主體是一般產線自動化與檢測，非機器人台數。")
    X.cell(17, 1, "不計入本頁：半導體設備投資（上半海外 半導體・液晶 +65%，當地貨幣）屬 AI 算力資本支出的間接需求，非實體 AI；人形與 AMR 工廠端的檢測需求未取得量化證據。")
    X.cell(18, 1, "重判條件：公司或第三方揭露機器人／協作工作站相關營收或份額；人形量產線採用 Keyence 視覺或位移感測的公開案例。")
    X.column_dimensions["A"].width = 58; X.column_dimensions["C"].width = 50
    for k, rr_ in XR.items(): dn(f"pa_{k}", f"'實體AI'!$D${rr_}:$J${rr_}")

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("半導體與電子設備投資循環反轉（上半成長主力）", "是（go27_C、gd27_C、gm_C）", "情境 C 與 Sens", "補足資料業種別增速（半導體・液晶）、SEMI 設備出貨、台積電資本支出"),
             ("日圓升值（海外 68%）", "是（fx27_*）", "情境 C −7%；公司揭露營業利益敏感度 USD 1 円約 13 億", "USDJPY、EURJPY、CNYJPY"),
             ("中國需求與出口管制", "否", "地區別金額未取得（アジア 構成比 46.1%）", "アジア 現地通貨增速、中國製造業 PMI"),
             ("直銷人員擴張帶來的費用剛性", "是（e_sga、sga_fl）", "Sens", "販管費增速 vs 營收增速"),
             ("投資有価証券的市價與流動性", "部分（inv_k）", "明細未取得", "有價證券報告書明細"),
             ("股票分割後口徑錯置（舊資料商數字為分割前）", "是（split_r）", "所有每股數字以 split_r 換算", "資料商每股數字與 csv 口徑")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 60

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([ER26, ER25, ERQ1, ERH1, FRH1, MS, KY, SA + "（Keyence、Cognex、Omron、Novanta、Hexagon、SMC）", TRD, YCD, SPL, JGB10, CP, IRX,
                            "Damodaran 國家風險溢酬（2026-01-05）"], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    bid = sens["build_id"] if sens else "（基準檔）"
    RD["A1"] = f"Keyence（6861.T）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-20（財報、共識、同業倍數）；成品日 2026-11-23；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"#,##0")&" 日圓，分割調整後）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")&"；股票分割 1 拆 5（2026-11-21 生效）：股數、EPS、每股股利皆為分割後口徑"'
    RD["A3"] = "命題：Keyence 是工廠自動化感測與機器視覺的高毛利直銷商，股價由全球製造業設備投資循環（2026 年上半由半導體與電子帶動）、匯率與營業利益率決定；實體 AI（機器人工作站）只是其營收的一小部分"
    RD["A4"] = (f'="在 FY2027 海外當地貨幣營收增速（情境 B）"&TEXT(go27_B,"0%")&"、毛利率 "&TEXT(gm_B,"0.0%")&"、匯率持平的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"#,##0")&" 日圓（DCF "&TEXT(Val!B{r_dcft},"#,##0")&"、Comparable "&TEXT(Val!B{r_cmpt},"#,##0")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0")&" 十億日圓）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 日圓。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")&"。現行前瞻 P/E 為同業中位數的 "&TEXT(Val!B{r_pe2},"0.00")&" 倍"'
    RD["A6"] = "失效條件（可觀察）：補足資料海外當地貨幣增速連兩季低於 5%；半導體・液晶 業種增速轉負；單季毛利率跌破 84%；USDJPY 跌破 145"
    RD["A7"] = (f'="實體 AI 曝險（Assumed，不進目標價）：機器人工作站相關感測視覺營收 FY2030 約 "&TEXT(\'實體AI\'!H8,"0")&"–"&TEXT(\'實體AI\'!H10,"0")&" 十億日圓（基準 "&TEXT(\'實體AI\'!H9,"0")&"），占情境 B 營收 "&TEXT(\'實體AI\'!H12,"0.0%")&"–"&TEXT(\'實體AI\'!H14,"0.0%")&"；人形、AMR 的感測內容在線 6 設為 0"')
    RD["A8"] = "範圍：不讀 Siliconomics／Tokenomics／Wattonomics（算力線）；自用 AI 讀法、Wattonomics 上線時程等規格條文不適用。HTML 展示層見 dist/ 同名 .html"
    RD.column_dimensions["A"].width = 160
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算（批次）與讀值 ----------------
LO = "file:///tmp/h_keyence/lo_profile"
def recalc_many(srcs, outdir):
    os.makedirs(outdir, exist_ok=True)
    for k in range(0, len(srcs), 12):
        chunk = srcs[k:k + 12]
        for attempt in range(2):
            try:
                subprocess.run(["soffice", f"-env:UserInstallation={LO}", "--headless", "--convert-to", "xlsx", "--outdir", outdir] + chunk,
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
                break
            except subprocess.TimeoutExpired:
                subprocess.run(["pkill", "-f", "lo_profile"], check=False)
                if attempt == 1: raise
    return [os.path.join(outdir, os.path.basename(s)) for s in srcs]

def read_vals(path, VR, DR, full=False):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]
    d = dict(tp=V[f"B{VR['tp']}"].value, dcf=V[f"B{VR['dcf']}"].value, cmp=V[f"B{VR['cmp']}"].value, up=V[f"B{VR['up']}"].value, rt=V[f"B{VR['rt']}"].value)
    if full:
        I = wb["In"]; w = {}; wd = None
        for row in I.iter_rows(values_only=True):
            if row[0] in ("w_A", "w_B", "w_C"): w[row[0][-1]] = row[4]
            if row[0] == "w_dcf": wd = row[4]
        alt = sum(w[s] * wb["DCF"][f"B{DR[s]['alt']}"].value for s in SC)
        d["alt_dcf"] = alt; d["alt_tp"] = wd * alt + (1 - wd) * d["cmp"]
        se = wb["Sens"]; d["fp"] = se["D5"].value
    return d

def run_cases(ovs):
    """ovs：[(tag, override dict)]；批次建置與重算，回傳 {tag: 值}。"""
    d = tempfile.mkdtemp(prefix="key_"); paths = []; meta = None
    for tag, ov in ovs:
        p = os.path.join(d, f"c_{tag}.xlsx"); meta = build(p, ov); paths.append(p)
    outs = recalc_many(paths, os.path.join(d, "o"))
    VR, _, DR = meta
    res = {tag: read_vals(o, VR, DR, full=(tag == "base")) for (tag, _), o in zip(ovs, outs)}
    shutil.rmtree(d, ignore_errors=True); return res

def in_val(k, col):
    for row in ROWS:
        if row[0] == k:
            return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)

def fingerprint():
    return hashlib.sha1(json.dumps([r for r in ROWS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]

def rating(u):
    return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"

CASES = [("gd26", "FY2026 國內營收年增"), ("go26", "FY2026 海外營收年增（當地貨幣）"), ("fx26", "FY2026 匯率換算效果"), ("gm26", "FY2026 毛利率"), ("sga26_r", "FY2026 營業費用率"),
         ("gd27_B", "國內營收年增 FY2027（情境 B）"), ("gdm_B", "國內營收年增 FY2028–29（情境 B）"), ("gdlt_B", "國內營收年增 FY2030–32（情境 B）"),
         ("go27_B", "海外營收年增 FY2027（情境 B）"), ("gom_B", "海外營收年增 FY2028–29（情境 B）"), ("golt_B", "海外營收年增 FY2030–32（情境 B）"),
         ("fx27_B", "FY2027 匯率換算效果（情境 B）"), ("gm_B", "毛利率 FY2027 起（情境 B）"), ("e_sga", "營業費用彈性"), ("sga_fl", "營業費用底線增速"),
         ("nonop", "業外收支淨額"), ("tax_r", "有效稅率"), ("beta", "β"), ("erp", "ERP"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"),
         ("fc_da", "折舊 ÷ 營收"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("inv_k", "投資有価証券計入比例"), ("dps_tgt", "期間股利"),
         ("w_dcf", "DCF 權重"), ("rel_k", "相對倍數係數")]


def write_readme(path, VR, PR, DR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f0 = lambda x: f"{x:,.0f}"; pc = lambda x: f"{x*100:+.1f}%"
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    inv = {I.cell(r, 1).value: I.cell(r, 5).value for r in range(1, I.max_row + 1)}
    base_tp = sens["tp0"]
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / base_tp - 1))[:8]
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {r['val'] if isinstance(r['val'], str) else f'{r[chr(118)+chr(97)+chr(108)]:.4f}'} | {r['base']} | {r['note']} |" for r in revs)
    def yr(k, fmt="{:,.1f}"):
        out = []
        for y in YRS:
            v = P[f"{COL[y]}{PR['B'][k]}"].value
            out.append(fmt.format(v) if isinstance(v, (int, float)) else "—")
        return " | ".join(out)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f0(g(chr(66+i)+str(VR['dcfr'])))} | {f0(g(chr(66+i)+str(VR['ca'])))} | {f0(g(chr(66+i)+str(VR['tps'])))} |" for i, s in enumerate(SC))
    rts = sorted(set(rating(c["tp"] / inv["px"] - 1) for c in sens["cases"]))
    X = wb["實體AI"]
    txt = f"""# physical_ai/keyence — Keyence（キーエンス，6861.T）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

東證 Prime 上市（亞太基金可投資）。所屬線：線 6 共用零組件（工業感測與機器視覺）；需求變數另讀線 2a 工業機器人。本 README 的數字由 `build_keyence.py` 從重算後的成品讀出產生，不手寫。

**股票分割：** 2026-11-21 生效 1 拆 5（基準日 2026-11-20；`physical_ai/data/closes2_20261109_20261120.csv.splits.txt` 記 2026-11-19）。csv 收盤為分割調整後價格；本模型股數（評價日 {inv['sh_v']:,.2f} 百萬股）、EPS、每股股利、目標價全部用分割後口徑，財報原文的分割前數字以 `split_r`＝5 換算，In 頁 A2 寫明。

**公司不揭露業績預測**（決算短信只有股利預測）：FY2026（2027年3月期）＝上半實績＋下半模型，Recon 以資料商共識（16 位，中間決算後）對照。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

## 估值（評價日 2027-03-20；目標時點 2027-11-20）
| 情境 | DCF（推到目標時點） | Comparable（同業 P/E × NTM EPS） | 情境目標價 |
|---|---|---|---|
{scn}

- 加權目標價 **{f0(tp)} 日圓**（DCF {f0(dcf)}、Comparable {f0(cmp_)}），現價 {f0(inv['px'])}（2026-11-20 收盤，分割調整後），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")}（情境 B FY2026–30 自由現金流扣股利 {g(f"C{VR['gate']}"):,.0f} 十億日圓）→ 投資建議 **{g(f"B{VR['rec']}")}**。
- 方法：DCF 標準版 FCFF（WACC／ke／g／RONIC＝{g(f"B{VR['wk']}")}；無有息負債，WACC＝ke）50%＋Comparable 50%（同業 6 家前瞻 P/E 中位數 {inv['pe_med']:.2f} 倍 × 目標時點 NTM EPS；視覺與量測子群 {inv['pe_vm']:.2f} 倍、FA 零組件子群 {inv['pe_fa']:.2f} 倍，純展示）。相對倍數係數 1.00（未調整：缺 Keyence 對同業的 10 年前瞻 P/E 歷史）。
- Keyence 現行前瞻 P/E {inv['pe_self']:.1f} 倍（現價 ÷ 共識 FY2026 EPS），為同業中位數的 {g(f"B{VR['pe2']}"):.2f} 倍；Comparable 以同業中位數計，等於假設這個溢價在 12 個月內消失——這是 Comparable 低於現價的主因（見 Sens 的 rel_k 高案）。
- 每股淨現金與投資 {g(f"B{VR['cash']}"):,.0f} 日圓（占現價 {g(f"B{VR['cashp']}")*100:.1f}%）。隱含前瞻 P/E（加權目標價 ÷ 情境 B NTM EPS）{g(f"B{VR['pe']}"):.1f} 倍。

## 營收分部與標準損益表（官方單一分部；PL 頁三情境，情境 B 摘要，¥B）
| 列 | {" | ".join(f"FY{y}" for y in YRS)} |
|---|{"---|" * len(YRS)}
| 國內營收（子列） | {yr('dom')} |
| 海外營收（子列） | {yr('ovs')} |
| 合併營收 | {yr('rev')} |
| 毛利率 | {yr('gm', '{:.1%}')} |
| 營業利益 | {yr('op')} |
| 營業利益率 | {yr('opm', '{:.1%}')} |
| 歸屬母公司淨利 | {yr('ni')} |
| EPS（日圓，分割後） | {yr('eps', '{:,.0f}')} |

- 官方分部：單一分部（決算短信「電子応用機器の製造・販売を中心に事業活動を展開する単一セグメント」）；決算短信固定揭露國內／海外營收，作為子列（合計＝合併營收，核對列為 0）。
- FY2023–FY2025 為實績（日本基準）；FY2026＝上半實績 + 下半模型（國內 +{inv['gd26']*100:.0f}%、海外當地貨幣 +{inv['go26']*100:.0f}%、匯率換算 +{inv['fx26']*100:.0f}%）；FY2027 起依情境。
- 命題模組：營收＝國內 ×(1＋國內增速) ＋ 海外 ×(1＋當地貨幣增速)×(1＋匯率換算效果)。營業成本＝營收 ×(1 − 毛利率)；營業費用＝前一年 ×(1＋彈性 × 營收增速＋(1 − 彈性) × 底線增速)，下行時有費用剛性。

## 主要假設（In 頁；Assumed 皆給區間）
- 國內營收年增（情境 A／B／C）：FY2027 10%／5%／−5%；FY2028–29 7%／5%／2%；FY2030–32 5%／4%／2%。
- 海外營收年增（當地貨幣）：FY2027 16%／10%／−8%；FY2028–29 12%／9%／5%；FY2030–32 9%／7%／4%。FY2027 匯率換算 +2%／0%／−7%。
- 毛利率 FY2027 起 86.0%／85.5%／84.0%；營業費用彈性 0.7（0.5–0.9）、底線增速 6%（4–8%）。
- 折現：rf 2.817%（日本 10 年期，與 Fanuc、哈默納科同值）、β {inv['beta']:.3f}（Keyence 1.02 與同業中位數 1.06 的 Blume 中位數）、ERP 5%、CRP 0.91%、g＝MIN(3%, rf)、RONIC＝WACC；資本支出 2.0%、折舊 1.5%、營運資金 38% × 營收增量。
- 評價日股數：2026-09-20 流通股 242.52 百萬股 × 5 ＝ {inv['sh_v']:,.2f} 百萬股（分割後）；無 RSU、選擇權。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| 海外當地貨幣增速（go27、gom） | 公司補足資料每季地區別現地通貨增速（Q2'26 海外 +30.5%）、業種別（半導體・液晶 +65%） |
| 匯率（fx27） | USDJPY、EURJPY、CNYJPY；公司揭露營業利益敏感度 |
| 毛利率（gm） | 單季毛利率（Q2'26 86.2%） |
| 營業費用彈性（e_sga、sga_fl） | 販管費年增 vs 營收年增 |
| 相對倍數係數（rel_k） | Keyence 對同業前瞻 P/E 比值的歷史（未取得） |

失效條件：{RD['A6'].value.replace('失效條件（可觀察）：', '')}。評等翻轉需要的條件見下方反解（Sens 各案例評等：{'、'.join(rts)}）。

## 實體 AI 曝險（另列；Assumed，不進目標價）
{RD['A7'].value}。計算：線 6 FAS 市場（五線歸因，{CP} 情境!D47:J49）× Keyence 份額（Assumed 10%／25%／40%，線 6 份額頁未取得）× 2026-11-20 匯率。半導體設備投資帶來的成長屬 AI 算力資本支出的間接需求，不計入實體 AI。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f0(c['tp'])} | {pc(c['tp']/base_tp-1)} | {rating(c['tp']/inv['px']-1)} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁。

## 反解（整本重算，二分法；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
- 線 6：FAS 市場 2026／2030／2032 基準 502／627／674 百萬美元（情境!D48、H48、J48；下游取數頁未列 FAS，改讀情境頁）、低／高；每台內容 800／500 美元（輸入!D58:E58）；人形、AMR 為 0；Keyence 份額未取得。
- 線 2a：全球裝機 IR.inst.global.<年>.base（2026→2032 年複合約 6.6%）只作方向對照；JARA 受注 +46.6%、中國產量 1–9 月 +29.1% 為情境 A 與 FY2026 假設的說明。
- 與線模型不一致處：本模型不以機器人裝機驅動營收（Keyence 需求主體是一般產線自動化）；Keyence 份額線模型未取得，本模型自設區間。

## 限制與未取得資料
- 公司不揭露業績預測與分地區金額（只揭露增速與單季構成比），FY2026 下半與 FY2027 起全為模型假設；下半前年同期平均匯率未取得。
- 投資有価証券的內容（債券、存款、持分法）未取得，以帳面計入（低案 80%）。
- 同業前瞻 P/E：Cognex、Omron、Hexagon 不在收盤 csv，報價日未取得 2026-11-20（頁面日期 2026-09-18～11-13）；Novanta、SMC、Fanuc 換算至 2026-11-20。少於 8 家，不做迴歸。
- 相對倍數係數未調整（缺 10 年前瞻 P/E 歷史）；這是評等的承重輸入。
- 未做：display_check.json；未參與製作的代理獨立核對（Project 規則第 10 條，待主代理安排）。
- 規格中 Siliconomics／Tokenomics／Wattonomics 協調不適用於實體 AI 線，略過。DCF 為本模型自寫（未用 `tools/fcff_dcf.py`），公式同 `docs/plan/個股DCF標準規格.md`。

## 檔案
- `build_keyence.py`：建置＋整本重算 Sens／反解＋產生本 README、`segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、實體AI、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-23）：初版；現價 {f0(inv['px'])}（2026-11-20 收盤，分割調整後）；加權目標價 {f0(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


def main():
    beta_b = statistics.median([0.67 * 1.02 + 0.33, 0.67 * statistics.median([1.48, 1.14, 1.70, 0.85, 0.83, 0.98]) + 0.33])
    jobs = [("base", {})]; meta = []
    for k, drv in CASES:
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if k == "beta": v = beta_b + (-0.3 if side == "lo" else 0.3)
            if v is None: continue
            if k == "rel_k" and side == "hi": continue   # 公式，基準跑完後再算
            b_ = in_val(k, "base")
            if isinstance(b_, (int, float)) and abs(v - b_) < 1e-12: continue
            tag = f"{k}_{side}"; jobs.append((tag, {k: v})); meta.append((tag, k, side, v, drv))
    for s in SC:
        jobs.append((f"only_{s}", {f"w_{x}": (1.0 if x == s else 0.0) for x in SC}))
    res = run_cases(jobs)
    base = res["base"]; px0 = in_val("px", "base"); print("基準", base)
    # rel_k 高案：現行 P/E ÷ 同業中位數（以基準檔值）
    wb0 = None
    d = tempfile.mkdtemp(prefix="keyb_"); p = os.path.join(d, "b.xlsx"); build(p, {}); o = recalc_many([p], os.path.join(d, "o"))[0]
    I0 = openpyxl.load_workbook(o, data_only=True)["In"]
    rk_hi = next(I0.cell(r, 6).value for r in range(1, I0.max_row + 1) if I0.cell(r, 1).value == "rel_k"); shutil.rmtree(d, ignore_errors=True)
    r2 = run_cases([("rel_k_hi", {"rel_k": rk_hi})])
    res.update(r2); meta.append(("rel_k_hi", "rel_k", "hi", rk_hi, "相對倍數係數"))
    cases = []
    for tag, k, side, v, drv in meta:
        r = res[tag]; cases.append(dict(drv=drv + ("（低）" if side == "lo" else "（高）"), set=f"{k}＝{v:.4g}", key=k, tp=r["tp"], dcf=r["dcf"], cmp=r["cmp"]))
    for s in SC:
        r = res[f"only_{s}"]; cases.append(dict(drv=f"只看情境 {s}", set=f"w_{s}＝1，其餘 0", key=f"w_{s}", tp=r["tp"], dcf=r["dcf"], cmp=r["cmp"]))
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"]))
    # 反解（批次二分法）
    up0 = base["up"]; tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    targets = [("目標價＝現價（上檔 0%）", 0.0)] + ([(f"升一級邊界（上檔 {upper:+.0%}）", upper)] if upper is not None else []) + ([(f"降一級邊界（上檔 {lower:+.0%}）", lower)] if lower is not None else [])
    SOLVE = [("gom_B", 0.09, -0.30, 0.80, "情境 B 的 FY2028–FY2029 海外營收年增（當地貨幣；其他不變，整本重算；情境 A／C 不動）"),
             ("golt_B", 0.07, -0.30, 0.60, "情境 B 的 FY2030–FY2032 海外營收年增（當地貨幣；情境 A／C 不動）"),
             ("gm_B", 0.855, 0.70, 0.95, "情境 B 的 FY2027 起毛利率（上限 95%；情境 A／C 不動）"),
             ("rel_k", 1.0, 0.3, 3.0, "相對倍數調整係數（Comparable 腿，三情境同動）")]
    probs = [(lab, tgt, key, b0, lo_, hi_, note) for key, b0, lo_, hi_, note in SOLVE for lab, tgt in targets]
    st = []
    ends = run_cases([(f"e{i}_{s}", {key: (lo_ if s == "lo" else hi_)}) for i, (lab, tgt, key, b0, lo_, hi_, note) in enumerate(probs) for s in ("lo", "hi")])
    for i, (lab, tgt, key, b0, lo_, hi_, note) in enumerate(probs):
        flo, fhi = ends[f"e{i}_lo"]["up"] - tgt, ends[f"e{i}_hi"]["up"] - tgt
        st.append(dict(lo=lo_, hi=hi_, flo=flo, ok=flo * fhi <= 0))
    for it in range(14):
        jobs = [(f"m{i}", {probs[i][2]: (s["lo"] + s["hi"]) / 2}) for i, s in enumerate(st) if s["ok"]]
        if not jobs: break
        rr = run_cases(jobs)
        for i, s in enumerate(st):
            if not s["ok"]: continue
            mid = (s["lo"] + s["hi"]) / 2; fm = rr[f"m{i}"]["up"] - probs[i][1]
            if s["flo"] * fm <= 0: s["hi"] = mid
            else: s["lo"], s["flo"] = mid, fm
    revs = []
    for (lab, tgt, key, b0, lo_, hi_, note), s in zip(probs, st):
        x = (s["lo"] + s["hi"]) / 2 if s["ok"] else f"區間 {lo_:+.2f}～{hi_:+.2f} 內無解"
        revs.append(dict(label=lab, key=key, val=x, base=b0, note=note)); print("反解", key, lab, x)
    bid = f"{VER}-{fingerprint()}"
    sens = dict(build_id=bid, fp=base["fp"], tp0=base["tp"], base=base, cases=cases)
    json.dump(sens, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1, default=str)
    # 成品
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="keyf_"); raw = os.path.join(tmp, NAME)
    VR, PR, DR = build(raw, {}, sens, revs)
    out = recalc_many([raw], os.path.join(tmp, "o"))[0]
    shutil.copy(out, os.path.join(dist, NAME))
    wbv = openpyxl.load_workbook(os.path.join(dist, NAME), data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10]); assert not errs, errs
    print("模型狀態", wbv["Sens"]["D7"].value)
    write_readme(os.path.join(dist, NAME), VR, PR, DR, bid, sens, revs)
    json.dump(dict(company="keyence", ticker="6861.T", model_sheet="PL",
                   official_segments=[dict(name="電子応用機器（単一セグメント）", label="合併營收（單一分部：電子應用機器）", period="FY2025（2026年3月期）", value=1169.289, unit="¥B",
                                           source=ER26 + "：セグメント情報「単一セグメントであるため、記載を省略」")],
                   sublines=[dict(parent="電子応用機器（単一セグメント）", labels=["國內營收（子列）", "海外營收（子列）"])],
                   check_labels=["檢查：國內＋海外 − 合併營收（應為 0）"],
                   note="單一官方分部；決算短信固定揭露國內／海外營收（FY2025 國內 390.066、海外 779.222，合計＝1,169.289）。季度地區別金額未揭露（只有增速與構成比），分部對帳以年度為準。"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)

if __name__ == "__main__":
    main()

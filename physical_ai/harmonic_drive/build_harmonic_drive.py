# 哈默納科 Harmonic Drive Systems（6324.T）個股模型 v0.1.1 建置腳本
# 產出：physical_ai/harmonic_drive/dist/20261123_哈默納科個股模型_v0.1.1.xlsx（LibreOffice 重算後的成品，含公式與快取值）
# 結構：命題模組（Thesis：諧波減速機需求＝工業機器人循環＋協作＋人形，另有半導體設備與其他用途）→ 官方分部（日本／中國／北美／歐洲）→ 損益 → FCFF DCF＋同業 P/E。
# 敏感度：每個案例覆寫 In 頁基準欄後整本重建、LibreOffice 重算（整本重算，不用閉式近似）；反解同法以二分搜尋。
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md。實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics 協調檔。
# 用法：python3 physical_ai/harmonic_drive/build_harmonic_drive.py   （需要 soffice；暫存於 /tmp/hds_build）
import datetime as dt
import glob
import hashlib
import json
import os
import shutil
import subprocess
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1.1"
OUT = os.path.join(HERE, "dist", f"20261123_哈默納科個股模型_{VER}.xlsx")
TMP = "/tmp/hds_build"
LO = ["soffice", "-env:UserInstallation=file:///tmp/lo_hds_profile", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML"]

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDRF = PatternFill("solid", fgColor="DDEBF7")
SECF = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------- 出處 ----------------
S_Q2 = "公司 2027年3月期 第2四半期（中間期）決算短信〔日本基準〕（連結）英文版（2026-11-11；https://finance-frontend-pc-dist.west.edge.storage-yahoo.jp/disclosure/20261111/20261111573783.pdf）"
S_REV = "公司「業績予想及び配当予想の修正に関するお知らせ」（2026-11-11；https://finance-frontend-pc-dist.west.edge.storage-yahoo.jp/disclosure/20261111/20261111572884.pdf）"
S_Q1 = "公司 2027年3月期 第1四半期 決算短信〔日本基準〕（連結）（2026-08-07；https://www2.jpx.co.jp/disc/63240/140120260804507706.pdf）"
S_INV = "Investing.com 日文版報導（2026-11-11 20:09；https://jp.investing.com/news/earnings/article-93CH-1752424）；原文寫「2,630億円」等，與 Q2 營收 194 億對照可知單位多一位數，本模型除以 10"
S_EDB = "edinetdb.jp（EDINET 有價證券報告書彙整；https://edinetdb.jp/company/E01712 ；2026-11-16 讀取；有報原文未直接讀取，資料商為第二來源）"
S_ENV = "envalith.com 分部資料（有報彙整；https://envalith.com/en/companies/6324/segments ；2026-11-16 讀取；有報原文未直接讀取）"
S_QR4 = "Quartr FY2026/3 摘要（https://quartr.com/events/harmonic-drive-systems-inc-6324-q4-2026_F9xDWNfe ；設備投資 5,838 百萬日圓）"
S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes_20261109_20261120.csv；未還原股利，GitHub Actions 以 yfinance 擷取）"
S_PX_OLD = "Eulerpool（https://eulerpool.com/stock/JP3765150002 ；頁面更新 2026-11-15 23:12，對應 2026-11-13（五）收盤；單一來源，未與交易所交叉核對）"
S_SA = "stockanalysis.com statistics 頁（2026-11-16 讀取；原報價日各列不同，見 Comp 頁 L、M 欄）；股價改 Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes_20261109_20261120.csv），P/E 與市值按 新收盤 ÷ 原報價 等比例換算（EPS 不變）"
S_TR = "TipRanks（https://www.tipranks.com/stocks/jp:6324/forecast ；2026-11-16 讀取）"
S_IFR = "IFR World Robotics 2026（https://ifr.org/ifr-press-releases/news/presidents-report-by-jane-heffner-3-2026 ；經線 2a 工業機器人框架 v0.1 引用）"
S_L2A = "線 2a 工業機器人框架 v0.1（2026-11-16）"
S_L2B = "線 2b 人形機器人框架 v0.1（2026-11-16）"
S_JGB = "日本財務省 國債金利情報 jgbcme.csv 10 年：2026-11-02 2.817%（https://www.mof.go.jp/english/policy/jgbs/reference/interest_rate/jgbcme.csv ；與 accretech 模型同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：日本 CRP 0.91%（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html ；與 accretech 模型同一取值）"

IP_CO = "發行公司：公司預測與政策，有達標與股價誘因"
IP_MEDIA = "媒體轉述（可能引用券商或公司說明會），未見原始文件；單位經本模型修正"
IP_IFR = "IFR 為機器人產業協會，對產業成長有推廣立場"
IP_L2B = "產業網站與賣方彙整（PICEA、MS 等），有推廣或銷售研究的誘因"

# ---------------- In 頁輸入 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 來源誘因, 說明)
IN = [
    ("sec", "1. 市場資料與日期"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, "", "資料截止日 2026-11-16（財報、同業 EPS、β）；股價 2026-11-20；成品日 2026-11-23"),
    ("px", "現價（2026-11-20 收盤）", "日圓", None, 3722, None, "Verified", S_PX, "", "v0.1 用 2026-11-13 收盤 4,775（" + S_PX_OLD + "；與 csv 同日值一致）；11-06 收盤 4,265（TipRanks）；11-11 決算後當日 −7.1%（Investing.com）"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,12,31)", None, "Assumed", "本專案慣例（與 accretech 同）：評價日 2026-12-31，FY2026 只計 2027-01～03", "", "折現期間由評價日起算"),
    ("tp_date", "目標時點（px_date 後 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", ""),
    ("bs_date", "資產負債表日", "日期", None, "=DATE(2026,9,30)", None, "Verified", S_Q2, "", ""),
    ("stub_frac", "Stub 期（資產負債表日→評價日）", "年", None, "=(val_date-bs_date)/365", None, "Derived", "＝(評價日 − 資產負債表日) ÷ 365", "", "Stub FCFF＝FY2026 FCFF × stub_frac"),
    ("frac_0", "FY2026 在評價日後的比例", "年", None, "=(DATE(2027,3,31)-val_date)/365", None, "Derived", "＝(2027-03-31 − 評價日) ÷ 365", "", "期中慣例：FY2026 折現年數＝frac_0 ÷ 2"),
    ("sh_v", "評價日股數（流通股）", "百萬股", None, "=(96315400-1638685)/1e6", None, "Verified", S_Q2 + "：発行済 96,315,400 − 自己株式 1,638,685（2026-09-30）", "", "RSU／選擇權：本輪未取得揭露（董事股票報酬規模小），以 0 計；見 README 限制"),
    ("cash_v", "現金及存款（2026-09-30）", "十億日圓", None, 21.855, None, "Verified", S_Q2, "", "3 月底 21.564"),
    ("invs_v", "投資有價證券（2026-09-30）", "十億日圓", None, 0.848, None, "Verified", S_Q2, "", "全額計入淨現金"),
    ("debt_v", "借款合計（短借＋一年內長借＋長借）", "十億日圓", None, "=0.704+1.890+8.179", None, "Verified", S_Q2 + "：短期借入金 704、1年内返済予定の長期借入金 1,890、長期借入金 8,179（百萬日圓）", "", "有淨值、獲利相關財務約款（Yahoo!ファイナンス 風險欄）"),
    ("lease_v", "租賃負債（流動＋非流動）", "十億日圓", None, "=0.851+3.684", None, "Verified", S_Q2 + "：リース債務 流動 851、固定 3,684（百萬日圓）", "", "日本基準，海外子公司租賃上表"),
    ("int_dps", "中間配當（2026-12-04 支付，在 stub 期內）", "日圓／股", None, 12, None, "Verified", S_REV + "：中間配當 12 円（前回予想 10 円），総額 1,136 百萬日圓，效力發生日 2026-12-04", "", "stub_div＝12 × 股數"),
    ("stub_div", "Stub 期已宣告股利", "十億日圓", None, "=int_dps*sh_v/1000", None, "Derived", "＝中間配當 × 評價日股數 ÷ 1000（≈ 公司揭露総額 1.136）", "", ""),
    ("dps_tgt", "評價日→目標時點間支付的股利（FY2026 期末配當）", "日圓／股", None, 16, None, "Interested-party", S_REV + "：期末配當予想 16 円（2027-06 支付）", IP_CO, "FY2027 中間配當於 2027-12 支付，在目標時點後，不扣"),
    ("sec", "2. 公司指引（FY2026＝2027年3月期，2026-11-11 上修版）"),
    ("g_rev26", "FY2026 營收指引", "十億日圓", None, 78.0, None, "Interested-party", S_REV + "：売上高 78,000 百萬日圓（前回 74,500；期初 68,000）", IP_CO, "模型 FY2026 營收＝指引（三情境同）"),
    ("g_op26", "FY2026 營業利益指引", "十億日圓", None, 10.2, None, "Interested-party", S_REV + "：営業利益 10,200（前回 8,500）", IP_CO, ""),
    ("g_ord26", "FY2026 經常利益指引", "十億日圓", None, 10.0, None, "Interested-party", S_REV + "：経常利益 10,000", IP_CO, ""),
    ("g_ni26", "FY2026 歸屬母公司淨利指引", "十億日圓", None, 7.2, None, "Interested-party", S_REV + "：親会社株主に帰属する当期純利益 7,200；EPS 76.05 円", IP_CO, "只作對照，不進模型（模型淨利＝經常利益 ×（1 − tax）；基準稅率由本列推得，見 tax）"),
    ("payout", "配當性向目標", "%", None, 0.35, None, "Interested-party", S_REV + "：連結配當性向 35% を目標", IP_CO, "用於收支關卡的股利扣除"),
    ("sec", "3. 命題模組：用途拆分（FY2026 基期）"),
    ("mix_rob", "FY2026 營收中產業機器人（含協作）占比", "%", 0.25, 0.28, 0.35, "Assumed", S_INV + "：第 2 四半期受注 263 億中產業機器人 72 億（27%）", "", "公司用途別圖表需 JavaScript，本輪未取得數字；受注比例≠營收比例"),
    ("mix_semi", "FY2026 營收中半導體設備占比", "%", 0.18, 0.22, 0.26, "Assumed", S_INV + "：第 2 四半期受注中半導體設備 65 億（25%，年增 178%）", "", "半導體受注急增，營收占比應低於受注占比"),
    ("hum_26", "FY2026 人形／AI 機器人營收", "十億日圓", 0.4, 0.8, 1.6, "Assumed", S_INV + "：第 2 四半期「AI ロボット」受注 2 億 × 4 季", "", "以受注代營收（Assumed）"),
    ("ifr_25", "全球工業機器人裝機 CY2025", "千台", None, 603, None, "Verified", S_IFR + "：60.3 萬台（+11%）", "", "線 2a 同值"),
    ("ifr_26", "全球工業機器人裝機 CY2026（IFR 預測）", "千台", None, 655, None, "Interested-party", S_IFR + "：2026 年 65.5 萬台", IP_IFR, "CY2027–2028 以 CY2026→CY2029 幾何內插（Derived）"),
    ("ifr_29", "全球工業機器人裝機 CY2029（IFR 預測）", "千台", None, 806, None, "Interested-party", S_IFR + "：2029 年 80.6 萬台", IP_IFR, ""),
    ("hu_25", "全球人形機器人出貨 CY2025", "千台", None, 15.5, None, "Interested-party", S_L2B + "：2025 年約 1.3–1.8 萬台（取中點）", IP_L2B, "各機構口徑差 30–50%"),
    ("n_harm", "每台人形使用的諧波減速機數（有效）", "個", 8, 12, 16, "Analogy", S_L2B + "：Optimus Gen 2 類旋轉關節諧波 14–16 個；類比對象＝Optimus 類全尺寸；部分設計改用行星減速機", "", "基準取 12（考慮行星替代）"),
    ("asp_26", "人形用諧波減速機 FY2026 單價（HDS）", "千日圓／個", 80, 120, 180, "Analogy", S_L2B + "：高階諧波 800–1,500 美元（類比對象＝產業網站彙整的高階報價）× 150 日圓／美元；中國廠便宜 30–50%", "", ""),
    ("asp_lr", "人形用諧波單價年降幅", "%／年", 0.05, 0.10, 0.15, "Assumed", S_L2B + "：學習率每倍增 10–25%（Analogy 工業機器人與電池）；本模型以年降幅近似", "", ""),
    ("cn_mix", "預測期 中國分部占合併營收", "%", None, "=Q!D12/Q!D9", None, "Derived", "＝2026 年 4–6 月中國分部 ÷ 合併營收（Q 頁）", "", "預測期各分部占比固定；用途→地區對照未揭露"),
    ("na_mix", "預測期 北美分部占合併營收", "%", None, "=Q!D13/Q!D9", None, "Derived", "＝2026 年 4–6 月北美分部 ÷ 合併營收（Q 頁）", "", ""),
    ("eu_mix", "預測期 歐洲分部占合併營收", "%", None, "=Q!D14/Q!D9", None, "Derived", "＝2026 年 4–6 月歐洲分部 ÷ 合併營收（Q 頁）", "", "日本＝1 − 其他三區"),
    ("sec", "4. 命題模組：情境專屬驅動（A 樂觀／B 基準／C 保守；只列基準欄，不設敏感度區間）"),
    ("g27_rob_A", "情境 A：FY2027 產業機器人營收年增", "%", None, 0.20, None, "Assumed", "依據：JARA 2026 年 7–9 月受注 +46.6%；HDS 受注÷營收 1.35–1.45（Recon）", "", "受注領先營收，FY2027 消化受注殘"),
    ("g27_rob_B", "情境 B：FY2027 產業機器人營收年增", "%", None, 0.12, None, "Assumed", "同上", "", ""),
    ("g27_rob_C", "情境 C：FY2027 產業機器人營收年增", "%", None, -0.10, None, "Assumed", "同上；循環在 FY2027 轉向", "", ""),
    ("ginst_A", "情境 A：CY2030 起全球裝機年增", "%", None, 0.08, None, "Assumed", "IFR 2026→2029 年複合約 7.2%（Derived）作參照", "", ""),
    ("ginst_B", "情境 B：CY2030 起全球裝機年增", "%", None, 0.05, None, "Assumed", "同上", "", ""),
    ("ginst_C", "情境 C：CY2030 起全球裝機年增", "%", None, 0.02, None, "Assumed", "同上", "", ""),
    ("shd_A", "情境 A：產業機器人用諧波份額年漂移", "%／年", None, 0.0, None, "Assumed", S_L2A + "：陸製諧波比 HDS 便宜 40–60%（Interested-party）；IFR 中國國產比 2025 年降到 55%", "", "相對全球裝機成長的超額（負＝失份額）"),
    ("shd_B", "情境 B：產業機器人用諧波份額年漂移", "%／年", None, -0.02, None, "Assumed", "同上；FY2025 中國分部營收 −28%（Hist）", "", ""),
    ("shd_C", "情境 C：產業機器人用諧波份額年漂移", "%／年", None, -0.05, None, "Assumed", "同上", "", ""),
    ("g27_semi_A", "情境 A：FY2027 半導體設備營收年增", "%", None, 0.20, None, "Assumed", S_INV + "：Q2 半導體受注 +178%", "", ""),
    ("g27_semi_B", "情境 B：FY2027 半導體設備營收年增", "%", None, 0.10, None, "Assumed", "同上", "", ""),
    ("g27_semi_C", "情境 C：FY2027 半導體設備營收年增", "%", None, -0.10, None, "Assumed", "同上", "", ""),
    ("gsemi_A", "情境 A：FY2028 起半導體設備營收年增", "%", None, 0.08, None, "Assumed", "設備循環長期成長（AI半導體線只讀，本輪未取得其 WFE 路徑）", "", ""),
    ("gsemi_B", "情境 B：FY2028 起半導體設備營收年增", "%", None, 0.05, None, "Assumed", "同上", "", ""),
    ("gsemi_C", "情境 C：FY2028 起半導體設備營收年增", "%", None, 0.0, None, "Assumed", "同上", "", ""),
    ("g27_oth_A", "情境 A：FY2027 其他用途營收年增", "%", None, 0.05, None, "Assumed", S_INV + "：Q1、Q2 各含醫療機器人整批受注約 20 億", "", "整批受注不一定重複"),
    ("g27_oth_B", "情境 B：FY2027 其他用途營收年增", "%", None, 0.0, None, "Assumed", "同上", "", ""),
    ("g27_oth_C", "情境 C：FY2027 其他用途營收年增", "%", None, -0.10, None, "Assumed", "同上", "", ""),
    ("goth_A", "情境 A：FY2028 起其他用途營收年增", "%", None, 0.05, None, "Assumed", "醫療、工具機、航太等", "", ""),
    ("goth_B", "情境 B：FY2028 起其他用途營收年增", "%", None, 0.03, None, "Assumed", "同上", "", ""),
    ("goth_C", "情境 C：FY2028 起其他用途營收年增", "%", None, 0.01, None, "Assumed", "同上", "", ""),
    ("hu32_A", "情境 A：CY2032 全球人形出貨", "千台", None, 1500, None, "Assumed", S_L2B + "：Tesla Fremont 設計產能 100 萬台／年、Hyundai 2028 年 3 萬台（Interested-party）", "", "CY2025→CY2032 幾何路徑"),
    ("hu32_B", "情境 B：CY2032 全球人形出貨", "千台", None, 400, None, "Assumed", "同上", "", ""),
    ("hu32_C", "情境 C：CY2032 全球人形出貨", "千台", None, 100, None, "Assumed", "同上", "", ""),
    ("shr_hum_A", "情境 A：HDS 占全球人形諧波用量份額", "%", None, 0.10, None, "Assumed", S_L2B + "：中國整機多用國產諧波；候選 A 論點：量增利潤落在中國量產廠", "", ""),
    ("shr_hum_B", "情境 B：HDS 占全球人形諧波用量份額", "%", None, 0.04, None, "Assumed", "同上", "", ""),
    ("shr_hum_C", "情境 C：HDS 占全球人形諧波用量份額", "%", None, 0.015, None, "Assumed", "同上", "", ""),
    ("tp_wA", "情境權重 A", "", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", ""),
    ("tp_wB", "情境權重 B", "", None, 0.50, None, "Assumed", "同上", "", ""),
    ("tp_wC", "情境權重 C", "", None, 0.25, None, "Assumed", "同上", "", ""),
    ("sec", "5. 獲利與現金流"),
    ("inc_m", "增量營業利益率（相對 FY2026 指引）", "%", 0.25, 0.35, 0.45, "Assumed", "FY2025→FY2026 指引：Δ營業利益 7.63 ÷ Δ營收 18.44＝41%（Derived）；景氣下行時同率反向", "", "固定成本高、營運槓桿大"),
    ("nonop", "業外損益（每年）", "十億日圓", None, "=g_ord26-g_op26", None, "Derived", "＝FY2026 經常利益指引 − 營業利益指引", "", ""),
    ("tax", "有效稅率", "%", 0.26, "=1-g_ni26/g_ord26", 0.32, "Derived", "＝1 − 淨利指引 ÷ 經常利益指引（＝28%）；用於淨利、NOPAT 與 WACC", "", "敏感度低／高 26%／32%（Assumed）"),
    ("da_26", "FY2026 折舊攤銷 ÷ 營收", "%", None, "=2*3.523/g_rev26", None, "Derived", S_Q2 + "：上半年減価償却費 3,523 × 2 ÷ 營收指引", "", "之後線性收斂到資本支出比例 × da_ss"),
    ("da_ss", "穩態折舊 ÷ 資本支出", "x", None, 0.9, None, "Assumed", "穩態資本支出高於折舊（成長 g）", "", ""),
    ("cx_r", "資本支出 ÷ 營收", "%", 0.06, 0.08, 0.11, "Assumed", S_QR4 + "：FY2025 設備投資 5.838 ÷ 59.557＝9.8%；" + S_Q2 + "：上半年取得有形固定資產 2.051 ÷ 36.124＝5.7%", "", "公司 FY2026 設備投資計畫本輪未取得"),
    ("nwc_r", "營運資金增加 ÷ 營收增量", "%", 0.20, 0.30, 0.40, "Assumed", S_Q2 + "：上半年存貨 12.404→14.176（+1.77）", "", "應收、應付明細本輪未取得"),
    ("sec", "6. 折現率（DCF 標準規格第 2 節）"),
    ("rf", "無風險利率（日本 10 年期國債）", "%", None, 0.02817, None, "Verified", S_JGB, "", "日期 2026-11-02，較股價基準日 2026-11-20 早 13 個交易日（逾期；跨模型一致，下次全面改版更新）"),
    ("crp", "日本國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", ""),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", ""),
    ("b_own", "HDS 5 年 β（原始）", "x", None, 0.96, None, "Verified", "stockanalysis.com TYO:6324 statistics（頁面報價日 2026-10-02）", "", ""),
    ("b_peer", "同業 5 年 β 中位數（原始）", "x", None, "=MEDIAN(Comp!H5:H8)", None, "Derived", "＝Comp 頁核心同業 4 家 β 中位數", "", ""),
    ("beta", "β（兩來源 Blume 調整後中位數）", "x", "=beta-0.3", "=MEDIAN(0.67*b_own+0.33,0.67*b_peer+0.33)", "=beta+0.3", "Derived", "Blume 0.67β＋0.33；兩來源＝HDS 自身、同業中位數", "", "敏感度 ∓0.3（跨模型慣例）"),
    ("kd", "稅前債務成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt（銀行借款為主）", "", "負債占資本不到 3%"),
    ("wacc_adj", "WACC 調整（敏感度用）", "pt", -0.01, 0, 0.01, "Assumed", "基準 0；敏感度 ±1pt", "", ""),
    ("g_cap", "永續成長率上限", "%", None, 0.03, None, "Assumed", "DCF 標準規格：3%，且不超過計價幣別 10 年期公債殖利率", "", ""),
    ("g_tv", "永續成長率", "%", 0.02, "=MIN(g_cap,rf)", 0.04, "Derived", "＝MIN(g_cap, rf)", "", "敏感度 2%／4%（4% 超過 rf，只作敏感度）"),
    ("ronic_sp", "RONIC − WACC", "pt", None, 0, 0.05, "Assumed", "RONIC 預設＝WACC（新投資價值中性）", "", ""),
    ("tv_alt", "終值改用 FCFF ×（1＋g）（0＝否，1＝是）", "開關", None, 0, 1, "Assumed", "DCF 標準規格第 4 節敏感度", "", ""),
    ("sec", "7. 同業倍數與權重"),
    ("pe_peer", "同業前瞻 P/E 中位數", "x", "=MIN(Comp!F5:F8)", "=MEDIAN(Comp!F5:F8)", "=MAX(Comp!F5:F8)", "Derived", "＝Comp 頁核心同業 4 家（Nabtesco、Fanuc、安川、THK）前瞻 P/E 中位數；股價皆為 2026-11-20 收盤、P/E 由原報價日等比例換算（EPS 分母不變，見 Comp）", "", "敏感度＝同業最低／最高"),
    ("rel_k", "相對倍數係數（HDS ÷ 同業）", "x", None, 1.0, 2.5, "Assumed", "1.00（未調整：缺 HDS 對同業十年相對倍數歷史）", "", "高案 2.5 倍＝v0.1（2026-11-13 現價口徑）現價對 FY2026 指引 EPS 約 63 倍 ÷ 同業約 20 倍的一半以上（未查核，不作依據；v0.1.1 維持不改）"),
    ("w_dcf", "DCF 權重（Comparable＝1 − 本值）", "%", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設 DCF 50%＋同業倍數 50%", "", ""),
]
SCN_DRIVERS = [("g27_rob", "FY2027 產業機器人營收年增"), ("ginst", "CY2030 起全球裝機年增"), ("shd", "產業機器人用諧波份額年漂移"),
               ("g27_semi", "FY2027 半導體設備營收年增"), ("gsemi", "FY2028 起半導體設備營收年增"), ("g27_oth", "FY2027 其他用途營收年增"),
               ("goth", "FY2028 起其他用途營收年增"), ("hu32", "CY2032 全球人形出貨"), ("shr_hum", "HDS 人形諧波份額")]

FY = list(range(2024, 2033))
COL = {y: openpyxl.utils.get_column_letter(3 + i) for i, y in enumerate(FY)}   # C=FY2024 … K=FY2032
CYC = {y: openpyxl.utils.get_column_letter(4 + (y - 2025)) for y in range(2025, 2034)}   # Thesis：D=CY2025 … L=CY2033
SC = "ABC"


def hdr(ws, r, labels, c0=1):
    for j, h in enumerate(labels):
        c = ws.cell(r, c0 + j, h); c.font = BOLD; c.fill = HDRF


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def T(x, f="#,##0"):
    return f'TEXT({x},"{f}")'


def make(over=None, sens=None):
    """over：{鍵: 數值} 覆寫 In 基準欄；sens：最終建置時寫入 Sens 頁的結果 dict。回傳 Workbook。"""
    over = over or {}
    wb = openpyxl.Workbook()
    RD = wb.active; RD.title = "README"
    IW = wb.create_sheet("In"); HI = wb.create_sheet("Hist"); Q = wb.create_sheet("Q"); TH = wb.create_sheet("Thesis")
    PL = wb.create_sheet("PL"); DC = wb.create_sheet("DCF"); V = wb.create_sheet("Val"); CP = wb.create_sheet("Comp")
    SE = wb.create_sheet("Sens"); RC = wb.create_sheet("Recon"); DZ = wb.create_sheet("對帳"); RK = wb.create_sheet("Risk")

    # ================= In =================
    IW["A1"] = ("In 頁：藍字為輸入。基準欄 E 進主模型；敏感度低（D）／敏感度高（F）只用於 Sens 頁（一次動一個參數、整本重算），不進目標價。"
                "情境 A／B／C 的差異只寫在情境專屬列（第 4 節），最後一節「情境驅動對照」純展示。")
    IW["A2"] = "單位：十億日圓（¥B）除另註；FY2026＝2027年3月期（4 月制，年度以起始年命名，與 accretech 同）。"
    hdr(IW, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "來源誘因", "說明"])
    r = 4; KR = {}
    for row in IN:
        if row[0] == "sec":
            IW.cell(r, 1, row[1]).font = BOLD; IW.cell(r, 1).fill = SECF; r += 1; continue
        k, lab, unit, lo, base, hi, tag, src, inc, note = row
        if k in over: base = over[k]
        IW.cell(r, 1, k); IW.cell(r, 2, lab); IW.cell(r, 3, unit)
        for c, v in ((4, lo), (5, base), (6, hi)):
            if v is not None:
                cc = IW.cell(r, c, v); cc.font = BLUE
                if unit == "%" or unit == "%／年": cc.number_format = "0.0%"
                if unit == "日期": cc.number_format = "yyyy-mm-dd"
        IW.cell(r, 7, tag); IW.cell(r, 8, src); IW.cell(r, 9, inc); IW.cell(r, 10, note)
        dn(wb, k, f"In!$E${r}"); KR[k] = r; r += 1
    LAST_IN = r - 1
    r += 1
    IW.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字，本節不被任何公式引用）").font = BOLD; r += 1
    hdr(IW, r, ["驅動", "項目", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "引用與說明"]); r += 1
    for d, lab in SCN_DRIVERS:
        IW.cell(r, 1, d); IW.cell(r, 2, lab)
        for j, s in enumerate(SC):
            IW.cell(r, 4 + j, f"=E{KR[d + '_' + s]}")
        IW.cell(r, 8, f"引用 In 頁 {d}_A／{d}_B／{d}_C；情境專屬輸入")
        r += 1
    IW.cell(r, 1, "scn_w"); IW.cell(r, 2, "情境權重")
    for j, s in enumerate(SC):
        IW.cell(r, 4 + j, f"=E{KR['tp_w' + s]}")
    IW.cell(r, 8, "引用 In 頁 tp_wA／tp_wB／tp_wC；合計＝1")
    for col, w in zip("ABCDEFGHIJ", (14, 44, 10, 11, 12, 11, 15, 70, 30, 50)):
        IW.column_dimensions[col].width = w

    # ================= Hist（年度，FY＝起始年；FY2025＝2026年3月期）=================
    HI["A1"] = "Hist：歷史財務（十億日圓；FY2025＝2026年3月期）。公式只引用本頁（定義名稱 h_*）。"
    hdr(HI, 3, ["項目", "FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"])
    HROWS = [
        ("h_rev", "合併營收", [57.087, 71.527, 55.796, 55.645, 59.557], "Verified", S_EDB),
        ("h_op", "營業利益", [8.739, 10.224, 0.124, 0.006, 2.567], "Verified", S_EDB),
        ("h_ord", "經常利益", [9.108, 10.757, 0.570, 0.151, 2.539], "Verified", S_EDB),
        ("h_ni", "歸屬母公司淨利", [6.643, 7.595, -24.806, 3.473, 1.608], "Verified", S_EDB + "；FY2023 含德國子公司商譽等減損"),
        ("h_eps", "EPS（日圓）", [69.0, 79.7, -261.0, 36.6, 16.99], "Verified", S_EDB),
        ("h_dps", "每股股利（日圓）", [21, 28, 20, 20, 20], "Verified", S_EDB),
        ("h_ocf", "營業現金流", [9.881, 10.850, 12.728, 7.516, 6.425], "Verified", S_EDB),
        ("h_fcf", "自由現金流（資料商定義）", [5.178, 2.187, 6.778, 8.996, 1.483], "Verified", S_EDB),
        ("h_ebitda", "EBITDA（資料商定義）", [16.017, 18.745, 9.314, 8.030, 9.928], "Verified", S_EDB),
        ("h_da", "折舊攤銷（＝EBITDA − 營業利益）", ["=B12-B5", "=C12-C5", "=D12-D5", "=E12-E5", "=F12-F5"], "Derived", "＝EBITDA − 營業利益"),
        ("h_cx", "設備投資", [None, None, None, None, 5.838], "Verified", S_QR4),
        ("h_jp", "日本 營收（外部）", [None, None, None, 21.727, 26.627], "Verified", S_ENV),
        ("h_cn", "中國 營收（外部）", [None, None, None, 5.623, 4.047], "Verified", S_ENV + "；中國原含於日本，近年獨立為分部"),
        ("h_na", "北美 營收（外部）", [None, None, None, 11.629, 12.108], "Verified", S_ENV),
        ("h_eu", "歐洲 營收（外部）", [None, None, None, 16.667, 16.777], "Verified", S_ENV),
        ("h_jp_p", "日本 分部利益（經常利益口徑）", [None, None, None, 2.224, 3.694], "Verified", S_ENV),
        ("h_cn_p", "中國 分部利益", [None, None, None, 0.302, 0.382], "Verified", S_ENV),
        ("h_na_p", "北美 分部利益", [None, None, None, 0.557, 0.529], "Verified", S_ENV),
        ("h_eu_p", "歐洲 分部利益", [None, None, None, -0.053, 0.645], "Verified", S_ENV),
        ("h_red", "減速裝置 營收", [None, None, None, None, 46.333], "Verified", S_EDB),
        ("h_mec", "メカトロニクス 營收", [None, None, None, None, 13.224], "Verified", S_EDB),
        ("h_ord_q", "受注高", [None, None, None, None, 61.613], "Verified", S_EDB + "：616 億 13 百萬日圓（+16.2%）"),
    ]
    for i, (k, lab, vals, tag, src) in enumerate(HROWS):
        rr = 4 + i
        HI.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = HI.cell(rr, 2 + j, v); c.font = BLUE if not isinstance(v, str) else Font()
        HI.cell(rr, 7, tag); HI.cell(rr, 8, src)
        dn(wb, k, f"Hist!$B${rr}:$F${rr}")
    HR = {k: 4 + i for i, (k, *_r) in enumerate(HROWS)}
    rr = 4 + len(HROWS) + 1
    HI.cell(rr, 1, "檢查：分部合計 − 合併營收（FY2024、FY2025；四捨五入誤差）")
    for c in "EF":
        HI[f"{c}{rr}"] = f"={c}{HR['h_jp']}+{c}{HR['h_cn']}+{c}{HR['h_na']}+{c}{HR['h_eu']}-{c}{HR['h_rev']}"
    HI.cell(rr + 1, 1, "檢查：減速裝置＋メカトロニクス − 合併營收（FY2025）"); HI[f"F{rr + 1}"] = f"=F{HR['h_red']}+F{HR['h_mec']}-F{HR['h_rev']}"
    HI.column_dimensions["A"].width = 36; HI.column_dimensions["H"].width = 80

    # ================= Q（季度；長格式，一列一個數字，單一標記）=================
    Q["A1"] = "Q：季度資料（十億日圓）。FY2026Q1＝2026 年 4–6 月。一列一個數字、單一主要標記。"
    hdr(Q, 3, ["項目", "FY2025Q1", "FY2025Q2", "FY2026Q1", "FY2026Q2", "標記", "來源", "來源誘因", "說明"])
    QROWS = [  # (列, 項目, [Q1-25, Q2-25, Q1-26, Q2-26], 標記, 來源, 誘因, 說明)
        (4, "受注高", [None, "=E5/1.87", 24.128, None], "Verified", S_Q1, "", "FY2025Q2＝FY2026Q2 ÷ 1.87（Derived，不另列）；FY2025Q1＝24.128 ÷ 1.557"),
        (5, "受注高（FY2026Q2，媒體轉述）", [None, None, None, 26.3], "Interested-party", S_INV + "：年增 87%、季增 9%；市場預期 240 億", IP_MEDIA, "公司用途別受注數字頁需 JavaScript，本輪未取得原文"),
        (6, "受注殘高（期末）", [None, None, 29.457, "=D6+E5-E9"], "Derived", S_Q1 + "：29.457（+45.6%）；FY2026Q2＝Q1 受注殘＋Q2 受注 − Q2 營收", "", "FY2026Q1 為 Verified 原值；Q2 為推算"),
        (7, "受注÷營收", [None, None, "=D4/D9", "=E5/E9"], "Derived", "＝受注 ÷ 營收", "", "線 2a 引用 Q1 約 1.45"),
        (8, "受注：產業機器人／AI 機器人／半導體設備（FY2026Q2）", [None, None, None, "7.2／0.2／6.5"], "Interested-party", S_INV, IP_MEDIA, "產業機器人年增 51%；半導體年增 178%；醫療整批約 2.0"),
        (9, "合併營收", [13.496, "=27.838-B9", 16.681, "=36.124-D9"], "Verified", S_Q1 + "；" + S_Q2 + "（上半年 36.124、前年同期 27.838，Q2＝上半年 − Q1）", "", ""),
        (10, "營業利益", [0.122, "=0.465-B10", 1.840, "=4.205-D10"], "Verified", S_Q1 + "；" + S_Q2, "", ""),
        (11, "日本 營收（外部）", ["=D11/1.362", None, 8.090, None], "Verified", S_Q1 + "：8,090（+36.2%）", "", "FY2025Q1 由年增率推得"),
        (12, "中國 營收（外部）", ["=D12/0.817", None, 1.007, None], "Verified", S_Q1 + "：1,007（−18.3%）", "", ""),
        (13, "北美 營收（外部）", ["=D13/1.124", None, 3.035, None], "Verified", S_Q1 + "：3,035（+12.4%）", "", ""),
        (14, "歐洲 營收（外部）", ["=D14/1.255", None, 4.547, None], "Verified", S_Q1 + "：4,547（+25.5%）", "", "上半年分部數字：英文版決算短信未含，本輪未取得"),
        (15, "檢查：分部合計 − 合併營收", [None, None, "=SUM(D11:D14)-D9", None], "Derived", "四捨五入誤差", "", ""),
        (16, "減速裝置 營收", [None, None, 12.791, None], "Verified", S_Q1 + "：+18.4%", "", ""),
        (17, "メカトロニクス 營收", [None, None, 3.889, None], "Verified", S_Q1 + "：+44.5%", "", ""),
        (18, "歸屬母公司淨利", [-0.038, "=0.317-B18", 1.274, "=2.986-D18"], "Verified", S_Q1 + "；" + S_Q2, "", ""),
    ]
    # 修正：Q 頁第 9 列為合併營收（cn_mix 等引用 Q!D12/Q!D9）
    for (rr, lab, vals, tag, src, inc, note) in QROWS:
        Q.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = Q.cell(rr, 2 + j, v)
                if not isinstance(v, str) or not v.startswith("="): c.font = BLUE
        Q.cell(rr, 6, tag); Q.cell(rr, 7, src); Q.cell(rr, 8, inc); Q.cell(rr, 9, note)
    Q["B4"] = "=D4/1.557"
    Q.column_dimensions["A"].width = 40; Q.column_dimensions["G"].width = 80

    # ================= Thesis（命題模組）=================
    TH["A1"] = ("命題模組：HDS 營收＝產業機器人（含協作；全球裝機 × 份額漂移）＋半導體設備＋人形（出貨 × 每台諧波數 × HDS 份額 × 單價）＋其他用途。"
                "FY2026＝公司指引；FY（4 月制）＝9/12 × CY(t) ＋ 3/12 × CY(t+1)（原則第 10 節第 49 條）。")
    TH["A2"] = "欄：D＝CY2025／FY2025 … K＝CY2032／FY2032，L＝CY2033（只用於 FY2032 加權）。十億日圓。"
    hdr(TH, 3, ["項目", "鍵／說明", ""] + [f"FY{y}" for y in range(2025, 2033)] + ["CY2033"])
    THR = {}
    r = 4
    for s in SC:
        TH.cell(r, 1, f"情境 {s}").font = BOLD; TH.cell(r, 1).fill = SECF; r += 1
        R = {}
        # 全球裝機（CY）
        R["inst"] = r; TH.cell(r, 1, "全球工業機器人裝機（CY，千台）")
        for y in range(2025, 2034):
            c = CYC[y]
            if y == 2025: f = "=ifr_25"
            elif y == 2026: f = "=ifr_26"
            elif y in (2027, 2028): f = f"=ifr_26*(ifr_29/ifr_26)^({y - 2026}/3)"
            elif y == 2029: f = "=ifr_29"
            else: f = f"={CYC[y - 1]}{r}*(1+ginst_{s})"
            TH[f"{c}{r}"] = f
        r += 1
        R["instfy"] = r; TH.cell(r, 1, "FY 加權裝機（千台）")
        for y in range(2025, 2033):
            TH[f"{CYC[y]}{r}"] = f"=0.75*{CYC[y]}{R['inst']}+0.25*{CYC[y + 1]}{R['inst']}"
        r += 1
        R["rob"] = r; TH.cell(r, 1, "產業機器人（含協作）營收")
        for y in range(2026, 2033):
            c = CYC[y]
            if y == 2026: f = "=g_rev26*mix_rob"
            elif y == 2027: f = f"={CYC[2026]}{r}*(1+g27_rob_{s})"
            else: f = f"={CYC[y - 1]}{r}*{c}{R['instfy']}/{CYC[y - 1]}{R['instfy']}*(1+shd_{s})"
            TH[f"{c}{r}"] = f
        r += 1
        R["semi"] = r; TH.cell(r, 1, "半導體設備 營收")
        for y in range(2026, 2033):
            c = CYC[y]
            TH[f"{c}{r}"] = "=g_rev26*mix_semi" if y == 2026 else (f"={CYC[2026]}{r}*(1+g27_semi_{s})" if y == 2027 else f"={CYC[y - 1]}{r}*(1+gsemi_{s})")
        r += 1
        R["hu"] = r; TH.cell(r, 1, "全球人形出貨（CY，千台）")
        for y in range(2025, 2034):
            TH[f"{CYC[y]}{r}"] = f"=hu_25*(hu32_{s}/hu_25)^(({y}-2025)/7)"
        r += 1
        R["hufy"] = r; TH.cell(r, 1, "FY 加權人形出貨（千台）")
        for y in range(2025, 2033):
            TH[f"{CYC[y]}{r}"] = f"=0.75*{CYC[y]}{R['hu']}+0.25*{CYC[y + 1]}{R['hu']}"
        r += 1
        R["asp"] = r; TH.cell(r, 1, "人形用諧波單價（千日圓／個）")
        for y in range(2026, 2033):
            TH[f"{CYC[y]}{r}"] = "=asp_26" if y == 2026 else f"={CYC[y - 1]}{r}*(1-asp_lr)"
        r += 1
        R["hum"] = r; TH.cell(r, 1, "人形／AI 機器人 營收")
        TH.cell(r, 2, "FY2026＝hum_26；之後＝FY 出貨 × n_harm × 份額 × 單價 ÷ 1000")
        for y in range(2026, 2033):
            c = CYC[y]
            TH[f"{c}{r}"] = "=hum_26" if y == 2026 else f"={c}{R['hufy']}*n_harm*shr_hum_{s}*{c}{R['asp']}/1000"
        r += 1
        R["oth"] = r; TH.cell(r, 1, "其他用途（醫療、工具機、航太等）營收")
        TH.cell(r, 2, "FY2026＝指引 − 其他三項（Derived 殘差）")
        for y in range(2026, 2033):
            c = CYC[y]
            TH[f"{c}{r}"] = (f"=g_rev26-{c}{R['rob']}-{c}{R['semi']}-{c}{R['hum']}" if y == 2026 else
                             (f"={CYC[2026]}{r}*(1+g27_oth_{s})" if y == 2027 else f"={CYC[y - 1]}{r}*(1+goth_{s})"))
        r += 1
        R["tot"] = r; TH.cell(r, 1, "用途合計（＝合併營收）")
        for y in range(2026, 2033):
            c = CYC[y]
            TH[f"{c}{r}"] = f"={c}{R['rob']}+{c}{R['semi']}+{c}{R['hum']}+{c}{R['oth']}"
        r += 1
        R["humsh"] = r; TH.cell(r, 1, "人形占合併營收")
        for y in range(2026, 2033):
            c = CYC[y]; TH[f"{c}{r}"] = f"={c}{R['hum']}/{c}{R['tot']}"; TH[f"{c}{r}"].number_format = "0.0%"
        r += 2
        THR[s] = R
    for row in TH.iter_rows(min_row=4, max_row=r, min_col=4, max_col=12):
        for c in row:
            if c.number_format == "General": c.number_format = "#,##0.00"
    TH.column_dimensions["A"].width = 36; TH.column_dimensions["B"].width = 30

    # ================= PL（三情境）=================
    PL["A1"] = "PL：官方分部（日本／中國／北美／歐洲，地區別報告分部）→ 合併營收 → 損益 → FCFF。C＝FY2024、D＝FY2025 為實績；E＝FY2026（指引錨）… K＝FY2032。十億日圓。"
    hdr(PL, 3, ["項目", "單位"] + [f"FY{y}" for y in FY])
    PR = {}
    r = 4
    for s in ["B", "A", "C"]:
        R = {}; th = THR[s]
        pre = "" if s == "B" else f"（情境 {s}）"
        PL.cell(r, 1, f"情境 {s}" + ("（基準）" if s == "B" else "")).font = BOLD; PL.cell(r, 1).fill = SECF; r += 1
        for key, lab in (("rob", "用途：產業機器人（含協作）"), ("semi", "用途：半導體設備"), ("hum", "用途：人形／AI 機器人"), ("oth", "用途：其他")):
            R[key] = r; PL.cell(r, 1, lab + pre); PL.cell(r, 2, "¥B")
            for y in range(2026, 2033): PL[f"{COL[y]}{r}"] = f"=Thesis!{CYC[y]}{th[key]}"
            r += 1
        R["rev"] = r; PL.cell(r, 1, "合併營收" + pre); PL.cell(r, 2, "¥B")
        PL[f"C{r}"] = f"=Hist!E{HR['h_rev']}"; PL[f"D{r}"] = f"=Hist!F{HR['h_rev']}"
        for y in range(2026, 2033): PL[f"{COL[y]}{r}"] = f"=SUM({COL[y]}{R['rob']}:{COL[y]}{R['oth']})"
        r += 1
        for key, lab, hk, mix in (("jp", "日本 營收", "h_jp", None), ("cn", "中國 營收", "h_cn", "cn_mix"), ("na", "北美 營收", "h_na", "na_mix"), ("eu", "歐洲 營收", "h_eu", "eu_mix")):
            R[key] = r; PL.cell(r, 1, lab + pre); PL.cell(r, 2, "¥B")
            PL[f"C{r}"] = f"=Hist!E{HR[hk]}"; PL[f"D{r}"] = f"=Hist!F{HR[hk]}"
            for y in range(2026, 2033):
                PL[f"{COL[y]}{r}"] = f"={COL[y]}{R['rev']}*(1-cn_mix-na_mix-eu_mix)" if mix is None else f"={COL[y]}{R['rev']}*{mix}"
            r += 1
        R["chk"] = r; PL.cell(r, 1, ("檢查：分部合計 − 合併營收" if s == "B" else f"檢查（情境 {s}）：分部合計 − 合併營收"))
        for y in FY: PL[f"{COL[y]}{r}"] = f"=SUM({COL[y]}{R['jp']}:{COL[y]}{R['eu']})-{COL[y]}{R['rev']}"
        r += 1
        R["chk2"] = r; PL.cell(r, 1, "檢查：用途合計 − 合併營收" + pre)
        for y in range(2026, 2033): PL[f"{COL[y]}{r}"] = f"=Thesis!{CYC[y]}{th['tot']}-{COL[y]}{R['rev']}"
        r += 1
        R["op"] = r; PL.cell(r, 1, "營業利益" + pre); PL.cell(r, 2, "¥B")
        PL[f"C{r}"] = f"=Hist!E{HR['h_op']}"; PL[f"D{r}"] = f"=Hist!F{HR['h_op']}"; PL[f"E{r}"] = "=g_op26"
        for y in range(2027, 2033): PL[f"{COL[y]}{r}"] = f"=$E{r}+inc_m*({COL[y]}{R['rev']}-$E{R['rev']})"
        r += 1
        R["opm"] = r; PL.cell(r, 1, "營業利益率" + pre); PL.cell(r, 2, "%")
        for y in FY: PL[f"{COL[y]}{r}"] = f"={COL[y]}{R['op']}/{COL[y]}{R['rev']}"; PL[f"{COL[y]}{r}"].number_format = "0.0%"
        r += 1
        R["nonop"] = r; PL.cell(r, 1, "業外損益" + pre); PL.cell(r, 2, "¥B")
        PL[f"C{r}"] = f"=Hist!E{HR['h_ord']}-Hist!E{HR['h_op']}"; PL[f"D{r}"] = f"=Hist!F{HR['h_ord']}-Hist!F{HR['h_op']}"
        for y in range(2026, 2033): PL[f"{COL[y]}{r}"] = "=nonop"
        r += 1
        R["ord"] = r; PL.cell(r, 1, "經常利益" + pre); PL.cell(r, 2, "¥B")
        for y in FY: PL[f"{COL[y]}{r}"] = f"={COL[y]}{R['op']}+{COL[y]}{R['nonop']}"
        r += 1
        R["ni"] = r; PL.cell(r, 1, "歸屬母公司淨利" + pre); PL.cell(r, 2, "¥B")
        PL[f"C{r}"] = f"=Hist!E{HR['h_ni']}"; PL[f"D{r}"] = f"=Hist!F{HR['h_ni']}"
        for y in range(2026, 2033): PL[f"{COL[y]}{r}"] = f"={COL[y]}{R['ord']}*(1-tax)"
        r += 1
        R["eps"] = r; PL.cell(r, 1, "EPS" + pre); PL.cell(r, 2, "日圓")
        for y in FY: PL[f"{COL[y]}{r}"] = f"={COL[y]}{R['ni']}/sh_v*1000"
        r += 1
        R["da"] = r; PL.cell(r, 1, "折舊攤銷" + pre); PL.cell(r, 2, "¥B")
        PL[f"C{r}"] = f"=Hist!E{HR['h_da']}"; PL[f"D{r}"] = f"=Hist!F{HR['h_da']}"
        for i, y in enumerate(range(2026, 2033)):
            PL[f"{COL[y]}{r}"] = f"={COL[y]}{R['rev']}*(da_26+(cx_r*da_ss-da_26)*{i}/6)"
        r += 1
        R["cx"] = r; PL.cell(r, 1, "資本支出" + pre); PL.cell(r, 2, "¥B")
        PL[f"D{r}"] = f"=Hist!F{HR['h_cx']}"
        for y in range(2026, 2033): PL[f"{COL[y]}{r}"] = f"={COL[y]}{R['rev']}*cx_r"
        r += 1
        R["dnwc"] = r; PL.cell(r, 1, "營運資金增加" + pre); PL.cell(r, 2, "¥B")
        for y in range(2026, 2033): PL[f"{COL[y]}{r}"] = f"=nwc_r*({COL[y]}{R['rev']}-{COL[y - 1]}{R['rev']})"
        r += 1
        R["nopat"] = r; PL.cell(r, 1, "NOPAT（＝營業利益 ×（1 − 稅率））" + pre); PL.cell(r, 2, "¥B")
        for y in range(2026, 2033): PL[f"{COL[y]}{r}"] = f"={COL[y]}{R['op']}*(1-tax)"
        r += 1
        R["fcff"] = r; PL.cell(r, 1, "FCFF" + pre); PL.cell(r, 2, "¥B")
        for y in range(2026, 2033):
            c = COL[y]; PL[f"{c}{r}"] = f"={c}{R['nopat']}+{c}{R['da']}-{c}{R['cx']}-{c}{R['dnwc']}"
        r += 1
        R["fchk"] = r; PL.cell(r, 1, "檢查：FCFF 各列加總 − FCFF" + pre)
        for y in range(2026, 2033):
            c = COL[y]; PL[f"{c}{r}"] = f"=({c}{R['op']}*(1-tax)+{c}{R['da']}-{c}{R['cx']}-{c}{R['dnwc']})-{c}{R['fcff']}"
        r += 1
        R["div"] = r; PL.cell(r, 1, "股利（淨利 × 配當性向）" + pre); PL.cell(r, 2, "¥B")
        for y in range(2026, 2033): PL[f"{COL[y]}{r}"] = f"=MAX(0,{COL[y]}{R['ni']})*payout"
        r += 1
        R["gate"] = r; PL.cell(r, 1, "FCFF 扣股利（收支）" + pre); PL.cell(r, 2, "¥B")
        for y in range(2026, 2033): PL[f"{COL[y]}{r}"] = f"={COL[y]}{R['fcff']}-{COL[y]}{R['div']}"
        r += 2
        PR[s] = R
    for row in PL.iter_rows(min_row=4, max_row=r, min_col=3, max_col=11):
        for c in row:
            if c.number_format == "General": c.number_format = "#,##0.00"
    PL.column_dimensions["A"].width = 40

    # ================= DCF =================
    DC["A1"] = "DCF：FCFF 折現（期中慣例；FY2026 只計評價日後 frac_0）。終值＝NOPAT(FY2032) ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）；tv_alt＝1 時改用 FCFF ×（1＋g）÷（WACC − g）。"
    hdr(DC, 3, ["項目"] + [f"FY{y}" for y in range(2026, 2033)])
    DCOL = {y: openpyxl.utils.get_column_letter(2 + y - 2026) for y in range(2026, 2033)}
    DC.cell(4, 1, "計入比例")
    DC.cell(5, 1, "折現年數（評價日起，期中）")
    DC.cell(6, 1, "折現因子")
    for y in range(2026, 2033):
        c = DCOL[y]
        DC[f"{c}4"] = "=frac_0" if y == 2026 else 1
        DC[f"{c}5"] = "=frac_0/2" if y == 2026 else f"=frac_0+{y - 2027}+0.5"
        DC[f"{c}6"] = f"=1/(1+wacc)^{c}5"
    DR = {}
    r = 8
    for s in SC:
        R = {}
        DC.cell(r, 1, f"情境 {s}").font = BOLD; r += 1
        R["f"] = r; DC.cell(r, 1, "FCFF")
        for y in range(2026, 2033): DC[f"{DCOL[y]}{r}"] = f"=PL!{COL[y]}{PR[s]['fcff']}"
        r += 1
        R["pv"] = r; DC.cell(r, 1, "現值")
        for y in range(2026, 2033): c = DCOL[y]; DC[f"{c}{r}"] = f"={c}{R['f']}*{c}4*{c}6"
        r += 1
        R["sum"] = r; DC.cell(r, 1, "明確期現值合計"); DC[f"B{r}"] = f"=SUM(B{R['pv']}:H{R['pv']})"; r += 1
        R["tv"] = r; DC.cell(r, 1, "終值（FY2032 年底）")
        DC[f"B{r}"] = (f"=IF(tv_alt=1,PL!K{PR[s]['fcff']}*(1+g_tv)/(wacc-g_tv),"
                       f"PL!K{PR[s]['nopat']}*(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv))"); r += 1
        R["tvpv"] = r; DC.cell(r, 1, "終值現值"); DC[f"B{r}"] = f"=B{R['tv']}/(1+wacc)^(frac_0+6)"; r += 1
        R["ev"] = r; DC.cell(r, 1, "企業價值 EV"); DC[f"B{r}"] = f"=B{R['sum']}+B{R['tvpv']}"; r += 1
        R["tvs"] = r; DC.cell(r, 1, "終值占 EV"); DC[f"B{r}"] = f"=IF(B{R['ev']}>0,B{R['tvpv']}/B{R['ev']},\"N/M\")"; DC[f"B{r}"].number_format = "0.0%"; r += 1
        R["tvw"] = r; DC.cell(r, 1, "檢查：終值占 EV 是否超過 80%"); DC[f"B{r}"] = f"=IF(AND(B{R['ev']}>0,B{R['tvs']}<=0.8),\"未超過\",\"超過或 N/M（報告標示）\")"; r += 1
        R["stub"] = r; DC.cell(r, 1, "＋Stub 期 FCFF（FY2026 FCFF × stub_frac）"); DC[f"B{r}"] = f"=B{R['f']}*stub_frac"; r += 1
        R["nc"] = r; DC.cell(r, 1, "淨現金滾到評價日（現金＋投資 − 借款 − 租賃 ＋ Stub FCFF − 中間配當）"); DC[f"B{r}"] = f"=cash_v+invs_v-debt_v-lease_v+B{R['stub']}-stub_div"; r += 1
        R["eq"] = r; DC.cell(r, 1, "股權價值"); DC[f"B{r}"] = f"=B{R['ev']}+B{R['nc']}"; r += 1
        R["ps"] = r; DC.cell(r, 1, "每股價值（評價日，日圓）"); DC[f"B{r}"] = f"=MAX(0,B{R['eq']}/sh_v*1000)"; r += 1
        R["v12"] = r; DC.cell(r, 1, "推到目標時點（×（1＋ke）^t − 期間股利）"); DC[f"B{r}"] = f"=MAX(0,B{R['ps']}*(1+ke)^((tp_date-val_date)/365)-dps_tgt)"; r += 2
        DR[s] = R
    for row in DC.iter_rows(min_row=4, max_row=r, min_col=2, max_col=8):
        for c in row:
            if c.number_format == "General": c.number_format = "#,##0.00"
    DC.column_dimensions["A"].width = 60

    # ================= Comp =================
    CP["A1"] = ("Comp：同業前瞻倍數（stockanalysis.com statistics）。股價＝Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes_20261109_20261120.csv，Verified）；"
                "前瞻 P/E 與市值＝原值 × 新收盤 ÷ 原報價（Derived，EPS 與股數不變）；EV/EBITDA、β 維持原報價日（未更新）。核心同業 4 家進中位數，少於 8 家不做迴歸。")
    hdr(CP, 4, ["公司", "代號", "股價", "報價日", "市值（十億）", "前瞻 P/E", "EV/EBITDA（近 12 月）", "5 年 β", "會計口徑", "進中位數", "說明", "原報價（stockanalysis）", "原報價日", "原市值（十億）", "原前瞻 P/E", "標記與換算"])
    # (公司, 代號, 2026-11-20 收盤, 原報價, 原報價日, 原市值, 原前瞻 P/E, EV/EBITDA, β, 口徑, 進中位數, 說明)
    PEERS = [
        ("Nabtesco", "6268.T", 3105, 4149, "2026-11-13", 486.31, 18.40, 9.98, 0.93, "IFRS（未查核）", "是", "RV 減速機；精密減速機分部為主要可比"),
        ("Fanuc", "6954.T", 5047, 5541, "2026-11-13", 5120, 23.22, 14.45, 0.98, "日本基準（未查核）", "是", "機器人整機與 FA"),
        ("安川電機", "6506.T", 3180, 4360, "2026-09-16", 1130, 21.49, 17.57, 1.24, "IFRS（未查核）", "是", "原報價日逾期 2 個月；原報價與 csv 無重疊日，未能交叉核對"),
        ("THK", "6481.T", 4249, 6116, "2026-10-19", 685.44, 15.37, 13.17, 0.76, "IFRS（未查核）", "是", "線性傳動；人形絲杠題材；原報價與 csv 無重疊日，未能交叉核對"),
        ("綠的諧波 Leaderdrive", "688017.SS", 230.80, 263.76, "2026-10-09", 48.36, 225.96, 278.34, 0.79, "中國會計準則", "否", "諧波直接競爭者；倍數由人形題材定價，只作對照（市值單位人民幣十億）"),
        ("Nidec", "6594.T", 1644, 1800, "2026-10-16", 2060, 12.67, None, 1.07, "IFRS（未查核）", "否", "近 12 月虧損、業務組合不同，只作對照"),
    ]
    for i, (nm, tk, px1, px0, d0, mc0, pe0, eve, bt, acc, med, note) in enumerate(PEERS):
        r = 5 + i
        CP.cell(r, 1, nm); CP.cell(r, 2, tk)
        CP.cell(r, 3, px1).font = BLUE; CP.cell(r, 4, "2026-11-20")
        CP.cell(r, 5, f"=N{r}*C{r}/L{r}"); CP.cell(r, 6, f"=O{r}*C{r}/L{r}")
        if eve is not None: CP.cell(r, 7, eve).font = BLUE
        CP.cell(r, 8, bt).font = BLUE; CP.cell(r, 9, acc); CP.cell(r, 10, med)
        CP.cell(r, 11, note + f"；EV/EBITDA、β 為原報價日 {d0} 值（報價日未更新）")
        CP.cell(r, 12, px0).font = BLUE; CP.cell(r, 13, d0); CP.cell(r, 14, mc0).font = BLUE; CP.cell(r, 15, pe0).font = BLUE
        CP.cell(r, 16, "股價 Verified（Yahoo Finance 日線 2026-11-20 收盤）；P/E、市值 Derived：＝原值 × 2026-11-20 收盤 ÷ 原報價（分母 EPS、股數不變）"
                + ("；原報價與 csv 同日收盤一致" if d0 == "2026-11-13" else ""))
    CP.cell(12, 1, "核心同業前瞻 P/E 中位數"); CP["F12"] = "=MEDIAN(F5:F8)"
    CP.cell(13, 1, "核心同業 EV/EBITDA 中位數（只作對照）"); CP["G13"] = "=MEDIAN(G5:G8)"
    CP.cell(14, 1, "HDS 現價 ÷ FY2026 指引 EPS（分母：公司 2026-11-11 指引 76.05 円，FY2026）"); CP["F14"] = "=px/(g_ni26/sh_v*1000)"
    CP.cell(15, 1, "HDS 現價 ÷ 模型 FY2027 EPS（分母：情境 B，FY2027）"); CP["F15"] = f"=px/PL!F{PR['B']['eps']}"
    CP.cell(16, 1, "HDS 現行倍數 ÷ 同業中位數（FY2026 指引口徑）"); CP["F16"] = "=F14/F12"
    CP.cell(17, 1, "來源：" + S_SA + "；HDS 自身報價見 In 頁 px")
    CP.column_dimensions["A"].width = 52; CP.column_dimensions["K"].width = 50; CP.column_dimensions["P"].width = 70

    # ================= Val =================
    V["A1"] = "Val：DCF（FCFF，WACC）50% ＋ 同業 P/E（Comparable）50%，三情境 25／50／25 加權；約 12 個月目標價。"
    VR = {}
    r = 3
    V.cell(r, 1, "0. 折現率").font = BOLD; r += 1
    for key, lab, f, fmt in (("ke", "股權成本 ke＝rf＋β ×（ERP＋CRP）", "=rf+beta*(erp+crp)", "0.00%"),
                             ("kd_at", "稅後債務成本", "=kd*(1-tax)", "0.00%"),
                             ("e_mv", "股權市值（¥B）", "=px*sh_v/1000", "#,##0.0"),
                             ("d_bv", "負債（借款＋租賃，¥B）", "=debt_v+lease_v", "#,##0.0"),
                             ("wacc", "WACC", "=(e_mv*ke+d_bv*kd_at)/(e_mv+d_bv)+wacc_adj", "0.00%"),
                             ("g_show", "永續成長率 g", "=g_tv", "0.00%"),
                             ("ronic", "RONIC", "=wacc+ronic_sp", "0.00%")):
        V.cell(r, 1, lab); V[f"B{r}"] = f; V[f"B{r}"].number_format = fmt
        if key != "g_show": dn(wb, key, f"Val!$B${r}")
        VR[key] = r; r += 1
    V.cell(r, 1, "檢查：WACC 由 In 頁輸入重算 − Val 的 WACC")
    V[f"B{r}"] = "=((px*sh_v/1000)*(rf+beta*(erp+crp))+(debt_v+lease_v)*kd*(1-tax))/((px*sh_v/1000)+(debt_v+lease_v))+wacc_adj-wacc"; r += 2
    V.cell(r, 1, "1. 各情境價值（日圓／股，目標時點）").font = BOLD; r += 1
    hdr(V, r, ["項目", "情境 A", "情境 B", "情境 C", "說明"]); r += 1
    def row3(key, lab, fs, fmt="#,##0", note=""):
        nonlocal r
        V.cell(r, 1, lab)
        for j, f in enumerate(fs):
            V.cell(r, 2 + j, f).number_format = fmt
        V.cell(r, 5, note); VR[key] = r; r += 1
    row3("ev", "企業價值 EV（¥B）", [f"=DCF!B{DR[s]['ev']}" for s in SC], "#,##0.0")
    row3("tvs", "終值占 EV", [f"=DCF!B{DR[s]['tvs']}" for s in SC], "0.0%")
    row3("dcf", "DCF 情境價值推到 12 個月後", [f"=DCF!B{DR[s]['v12']}" for s in SC])
    row3("eps27", "FY2027 EPS（模型）", [f"=PL!F{PR[s]['eps']}" for s in SC], "#,##0.0", "分母：模型各情境 FY2027（2028年3月期）EPS")
    row3("cmpu", "Comparable 未調整（同業 P/E × FY2027 EPS）", [f"=MAX(0,{c}{r - 1}*pe_peer)" for c in "BCD"], "#,##0", "同業 P/E＝In 頁 pe_peer（核心同業前瞻 P/E 中位數）")
    row3("cmp", "Comparable（× 相對倍數係數 rel_k）", [f"={c}{VR['cmpu']}*rel_k" for c in "BCD"], "#,##0", "rel_k＝1.00（未調整：缺十年相對倍數歷史）")
    row3("wt", "情境加權前：DCF × w ＋ Comparable ×（1 − w）", [f"=w_dcf*{c}{VR['dcf']}+(1-w_dcf)*{c}{VR['cmp']}" for c in "BCD"])
    row3("w", "情境權重", ["=tp_wA", "=tp_wB", "=tp_wC"], "0%")
    r += 1
    V.cell(r, 1, "2. 目標價與評等").font = BOLD; r += 1
    def row1(key, lab, f, fmt="#,##0", note=""):
        nonlocal r
        V.cell(r, 1, lab); V[f"B{r}"] = f; V[f"B{r}"].number_format = fmt; V.cell(r, 5, note); VR[key] = r; r += 1
    row1("dcf_tp", "DCF 目標價（情境加權）", f"=SUMPRODUCT(B{VR['dcf']}:D{VR['dcf']},B{VR['w']}:D{VR['w']})")
    row1("cmp_tp", "Comparable 目標價（情境加權；同業 P/E × 相對倍數係數）", f"=SUMPRODUCT(B{VR['cmp']}:D{VR['cmp']},B{VR['w']}:D{VR['w']})")
    row1("cmpu_tp", "Comparable 未調整（情境加權）", f"=SUMPRODUCT(B{VR['cmpu']}:D{VR['cmpu']},B{VR['w']}:D{VR['w']})")
    row1("tp", "目標價（DCF 與 Comparable 加權）", f"=w_dcf*B{VR['dcf_tp']}+(1-w_dcf)*B{VR['cmp_tp']}")
    row1("px_show", "現價（股價基準日）", "=px")
    row1("up", "上檔空間：加權目標價", f"=B{VR['tp']}/px-1", "0.0%")
    row1("rating", "估值評等（只看上檔空間）", f'=IF(B{VR["up"]}>0.5,"強力買進",IF(B{VR["up"]}>0.25,"買進",IF(B{VR["up"]}>=-0.25,"中立",IF(B{VR["up"]}>=-0.5,"放空","強力放空"))))', "@",
         "級距：> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空（原則第 10 節第 24 條）")
    V.cell(r, 1, "收支關卡：FY2026–FY2030 FCFF 扣股利合計（情境 A／B／C，¥B）")
    for j, s in enumerate(SC):
        V.cell(r, 2 + j, f"=SUM(PL!E{PR[s]['gate']}:I{PR[s]['gate']})").number_format = "#,##0.0"
    V.cell(r, 5, "口徑：FCFF − 股利（未扣利息，借款利息小）；基準情境 B < 0＝入不敷出"); VR["gate"] = r; r += 1
    row1("gate_b", "入不敷出判定用（情境 B 合計）", f"=C{VR['gate']}", "#,##0.0")
    row1("gate_t", "入不敷出判定", f'=IF(B{VR["gate_b"]}<0,"不通過（入不敷出）","通過")', "@")
    row1("rec", "投資建議", f'=IF(AND(B{VR["gate_b"]}<0,OR(B{VR["rating"]}="買進",B{VR["rating"]}="強力買進")),"收支不足",B{VR["rating"]})', "@",
         "原則第 10 節第 63 條：收支不通過且估值評等為買進以上時＝收支不足")
    row1("dist", "距最近級距邊界（百分點）", f"=MIN(ABS(B{VR['up']}-0.5),ABS(B{VR['up']}-0.25),ABS(B{VR['up']}+0.25),ABS(B{VR['up']}+0.5))", "0.0%")
    row1("up_b", "升一級所需目標價（上一級邊界）", f'=IF(B{VR["up"]}>0.5,"—（已是最高級）",px*(1+IF(B{VR["up"]}<-0.5,-0.5,IF(B{VR["up"]}<-0.25,-0.25,IF(B{VR["up"]}<=0.25,0.25000001,0.50000001)))))')
    row1("dn_b", "降一級門檻目標價（下一級邊界）", f'=IF(B{VR["up"]}<-0.5,"—（已是最低級）",px*(1+IF(B{VR["up"]}>0.5,0.5,IF(B{VR["up"]}>0.25,0.25,IF(B{VR["up"]}>=-0.25,-0.25000001,-0.50000001)))))')
    r += 1
    V.cell(r, 1, "3. 上檔空間三段歸因與相對倍數（原則第 10 節第 47 條）").font = BOLD; r += 1
    row1("att_d", "DCF 貢獻＝w ×（DCF − 現價）÷ 現價", f"=w_dcf*(B{VR['dcf_tp']}-px)/px", "0.0%")
    row1("att_c", "未調整 Comparable 貢獻＝(1 − w) ×（未調整 − 現價）÷ 現價", f"=(1-w_dcf)*(B{VR['cmpu_tp']}-px)/px", "0.0%")
    row1("att_r", "相對倍數調整貢獻", f"=(1-w_dcf)*(B{VR['cmp_tp']}-B{VR['cmpu_tp']})/px", "0.0%")
    row1("att_chk", "檢查：三段合計 − 上檔空間", f"=B{VR['att_d']}+B{VR['att_c']}+B{VR['att_r']}-B{VR['up']}", "0.0000")
    row1("relnow", "HDS 現行前瞻倍數 ÷ 同業中位數（FY2026 指引 EPS 口徑）", "=Comp!F16", "0.00")
    row1("relreq", "支持現價所需的相對倍數係數（DCF 不變）", f"=(px-w_dcf*B{VR['dcf_tp']})/((1-w_dcf)*B{VR['cmpu_tp']})", "0.00",
         "反解：目標價＝現價；DCF 以 RONIC＝WACC（無超額報酬）而高相對倍數隱含持續超額報酬，兩法對競爭優勢的假設相反")
    row1("relflip", "評等升一級所需的相對倍數係數", f'=IF(ISNUMBER(B{VR["up_b"]}),(B{VR["up_b"]}-w_dcf*B{VR["dcf_tp"]})/((1-w_dcf)*B{VR["cmpu_tp"]}),"—")', "0.00")
    r += 1
    V.cell(r, 1, "4. 期間").font = BOLD; r += 1
    row1("t_model", "模型折現期間（評價日→目標時點，年）", "=(tp_date-val_date)/365", "0.00")
    row1("t_hold", "持有期間（股價基準日→目標時點，年；上檔空間不是年化報酬）", "=(tp_date-px_date)/365", "0.00")
    row1("tp_date_show", "目標時點", "=tp_date", "yyyy-mm-dd")
    V.column_dimensions["A"].width = 62; V.column_dimensions["E"].width = 70
    for c in "BCD": V.column_dimensions[c].width = 14

    # ================= Sens =================
    SE["A1"] = "Sens：一次動一個參數（In 頁敏感度低／高或指定值），整本重建並以 LibreOffice 重算後讀加權目標價（不是閉式近似）；建置腳本產生。"
    SE["A2"] = "build_id"; SE["A3"] = "模型狀態"; SE["A4"] = "In 頁輸入指紋（Sens 計算時）"; SE["A5"] = "In 頁輸入指紋（本檔即時）"
    fp_live = f"=SUMPRODUCT(In!E4:E{LAST_IN},ROW(In!E4:E{LAST_IN}))"
    SE["D5"] = fp_live
    hdr(SE, 7, ["驅動", "設定", "鍵", "輸入值", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    if sens:
        SE["D2"] = sens["build_id"]; SE["D4"] = sens["fp"]
        SE["D3"] = f'=IF(AND(ABS(D5-D4)<1E-6,ABS(Val!B{VR["tp"]}-{sens["base_tp"]})<0.01),"CURRENT","STALE")'
        r = 8
        for c in sens["cases"]:
            SE.cell(r, 1, c["drv"]); SE.cell(r, 2, c["set"]); SE.cell(r, 3, c["key"])
            if c.get("val") is not None: SE.cell(r, 4, c["val"])
            SE.cell(r, 5, c["tp"]).number_format = "#,##0"; SE.cell(r, 6, c["tp"] / sens["base_tp"] - 1).number_format = "+0.0%;-0.0%"
            SE.cell(r, 7, c["dcf"]).number_format = "#,##0"; SE.cell(r, 8, c["dcf"] / sens["base_dcf"] - 1 if sens["base_dcf"] else None).number_format = "+0.0%;-0.0%"
            SE.cell(r, 9, c["rating"]); r += 1
        r += 1
        SE.cell(r, 1, "讀法：相對基準＝案例加權目標價 ÷ 基準加權目標價 − 1；同一案例 DCF 單法與加權目標價並列（原則第 10 節第 32 條）。情境專屬驅動以「情境 B 改為情境 A／C 值」呈現。")
        SE.cell(r + 1, 1, "說明：模型已讓折舊、資本支出、營運資金隨營收等比例變動（第 44 條）；增量利益率對上行與下行對稱。")
    SE.column_dimensions["A"].width = 48; SE.column_dimensions["B"].width = 44

    # ================= Recon =================
    RC["A1"] = "Recon：模型對公司指引、受注、市場共識的同口徑對照；反解支持現價所需的主導變數。"
    hdr(RC, 3, ["項目", "模型", "對照", "差距", "說明"])
    rr = 4
    def rrow(lab, a, b, note, fmt="#,##0.00", diff=True):
        nonlocal rr
        RC.cell(rr, 1, lab); RC.cell(rr, 2, a).number_format = fmt
        if b is not None: RC.cell(rr, 3, b).number_format = fmt
        if diff and b is not None: RC.cell(rr, 4, f"=B{rr}-C{rr}").number_format = fmt
        RC.cell(rr, 5, note); rr += 1
    rrow("FY2026 營收（情境 B）vs 指引", f"=PL!E{PR['B']['rev']}", "=g_rev26", "模型以指引為錨（差距應為 0）")
    rrow("FY2026 營業利益 vs 指引", f"=PL!E{PR['B']['op']}", "=g_op26", "同上")
    rrow("FY2026 淨利 vs 指引", f"=PL!E{PR['B']['ni']}", "=g_ni26", "稅率由指引推得")
    rrow("FY2026 下半年隱含營收（指引 − 上半年實績）", "=g_rev26-36.124", "=Q!E9*2", "對照＝Q2 營收 × 2（下半年若維持 Q2 水準）；模型隱含下半年需求略高於 Q2 年化")
    rrow("FY2026Q2 受注 ÷ 營收", "=Q!E7", None, "> 1 表示受注殘增加；線 2a 候選 1 證偽條件：< 1.0", "0.00", False)
    rrow("FY2026Q1 受注 ÷ 營收", "=Q!D7", None, "", "0.00", False)
    rrow("FY2027 營收（情境 B）vs 賣方共識（Stockopedia 2028E）", f"=PL!F{PR['B']['rev']}", 87.202, "共識為決算前（2026-10-27 頁面），只作歷史對照；Stockopedia 年度標籤 2028E＝2028年3月期（推定）")
    rrow("FY2026 營收 vs 賣方共識（Stockopedia 2027E）", f"=PL!E{PR['B']['rev']}", 75.436, "財報前，只作歷史對照")
    rrow("加權目標價 vs 賣方平均目標價（TipRanks，4 家，近 3 個月）", f"=Val!B{VR['tp']}", 8317.5, S_TR + "；區間 6,470–11,000；只作對照，不進模型", "#,##0")
    rr += 1
    RC.cell(rr, 1, "反解（整本重算二分搜尋；建置時計算，數值為快照）").font = BOLD; rr += 1
    RV = rr
    if sens:
        for lab, v, note in sens["solve"]:
            RC.cell(rr, 1, lab); RC.cell(rr, 2, v).number_format = "#,##0.00"; RC.cell(rr, 5, note); rr += 1
    rr += 1
    RC.cell(rr, 1, "競爭者對照：綠的諧波（Leaderdrive, 688017）").font = BOLD; rr += 1
    RC.cell(rr, 1, "綠的諧波 近 12 月營收（人民幣十億）"); RC.cell(rr, 2, 0.66785); RC.cell(rr, 5, "stockanalysis.com SHA:688017（報價日 2026-10-09）；毛利率 30.8%"); rr += 1
    RC.cell(rr, 1, "換算日圓（× 23 日圓／人民幣，公司下半年匯率假設）"); RC.cell(rr, 2, f"=B{rr - 1}*23"); RC.cell(rr, 5, "HDS 匯率假設來自 " + S_REV); rr += 1
    RC.cell(rr, 1, "相對 HDS FY2026 營收指引"); RC.cell(rr, 2, f"=B{rr - 1}/g_rev26").number_format = "0.0%"; rr += 1
    RC.cell(rr, 1, "綠的諧波 前瞻 P/E ÷ HDS 現價對 FY2026 指引 P/E"); RC.cell(rr, 2, "=Comp!F9/Comp!F14").number_format = "0.0"; rr += 1
    RC.column_dimensions["A"].width = 60; RC.column_dimensions["E"].width = 80

    # ================= 對帳 =================
    DZ["A1"] = "對帳：本模型使用的「線層級」變數（與 physical_ai/lines/ 各線模型最後對帳；本模型讀線框架 v0.1，不靜默改口徑）。"
    hdr(DZ, 3, ["變數名", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處", "標記", "影響本模型哪些儲存格"])
    ZR = [
        ("全球工業機器人裝機 CY2025", "2a", "=ifr_25", "千台", "IFR 年度裝機（不是中國統計局產量）", S_IFR, "Verified", f"Thesis 各情境「全球工業機器人裝機」D 欄"),
        ("全球工業機器人裝機 CY2026（預測）", "2a", "=ifr_26", "千台", "IFR 預測", S_IFR, "Interested-party", "Thesis 產業機器人 FY2027 起成長"),
        ("全球工業機器人裝機 CY2029（預測）", "2a", "=ifr_29", "千台", "IFR 預測；CY2027–28 幾何內插", S_IFR, "Interested-party", "同上"),
        ("CY2030 起裝機年增（A／B／C）", "2a", "8%／5%／2%", "%／年", "本模型假設", "In 頁 ginst_*", "Assumed", "Thesis 產業機器人 FY2030–FY2032"),
        ("陸製諧波相對 HDS 價差", "2a／2b", "低 40–60%（2a）；便宜 30–50%（2b）", "%", "產業網站", S_L2A + "；" + S_L2B, "Interested-party", "份額漂移 shd_*（定性依據，不直接進公式）"),
        ("HDS 受注÷營收（FY2026Q1）", "2a", "=Q!D7", "x", "受注 ÷ 營收（季）", S_Q1, "Derived", "g27_rob_* 的依據（定性）"),
        ("JARA 受注 2026 年 7–9 月 YoY", "2a", "+46.6%", "%", "日系全體受注額", S_L2A + "（JARA）", "Verified", "g27_rob_* 的依據（定性）"),
        ("全球人形出貨 CY2025", "2b", "=hu_25", "千台", "1.3–1.8 萬台取中點", S_L2B, "Interested-party", "Thesis 人形出貨路徑起點"),
        ("全球人形出貨 CY2032（A／B／C）", "2b", "1,500／400／100", "千台", "本模型假設（線 2b 未給出貨預測）", "In 頁 hu32_*", "Assumed", "Thesis 人形營收"),
        ("每台人形諧波減速機數", "2b", "=n_harm", "個", "有效用量（14–16 個打折，考慮行星替代）", S_L2B, "Analogy", "Thesis 人形營收"),
        ("人形用諧波單價", "2b", "=asp_26", "千日圓／個", "高階 800–1,500 美元 × 150", S_L2B, "Analogy", "Thesis 人形營收"),
        ("諧波單價年降幅", "2b", "=asp_lr", "%／年", "學習率近似", S_L2B, "Assumed", "Thesis 人形營收"),
        ("協作機器人占全球裝機", "2a", "11%（2025）", "%", "IFR", S_IFR, "Verified", "未直接使用（併入產業機器人用途）；協作全用諧波，下一版拆子列"),
        ("協作 BOM 中諧波占比", "2a", "15–20%", "%", "產業報告", S_L2A, "Interested-party", "未直接使用"),
    ]
    for i, z in enumerate(ZR):
        for j, v in enumerate(z):
            DZ.cell(4 + i, 1 + j, v)
    DZ.column_dimensions["A"].width = 36; DZ.column_dimensions["F"].width = 60; DZ.column_dimensions["H"].width = 50

    # ================= Risk =================
    RK["A1"] = "Risk：風險登錄（原則第 10 節第 10、19 條）。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    RISKS = [
        ("中國廠（綠的諧波等）份額侵蝕：產業機器人與人形", "是（shd_*、shr_hum_*）", "情境 C 每年 −5% 份額", "HDS 中國分部營收；綠的諧波季營收與毛利率；IFR 中國國產比"),
        ("人形設計改用行星減速機或其他傳動，諧波用量下降", "部分（n_harm 敏感度 8–16）", "設計選擇未定", "Tesla Optimus V3 拆解、Figure／宇樹 BOM 揭露"),
        ("半導體設備循環反轉", "是（g27_semi_C、gsemi_C）", "", "HDS 半導體用途受注；日本半導體設備協會訂單"),
        ("醫療機器人整批受注不重複", "是（g27_oth_*）", "Q1、Q2 各約 20 億", "季受注中醫療用途"),
        ("匯率（日圓升值）", "否", "公司假設 157 日圓／美元、180 日圓／歐元；海外分部約 55%", "日圓匯率；公司匯率敏感度揭露（未取得）"),
        ("財務約款（淨值、獲利）", "否", "淨現金接近零但約款存在", "決算短信注記"),
        ("大股東 KODEN Holdings 持股 34.77%（治理、流動性）", "否", "不影響現金流", "大股東持股變動"),
        ("關稅與美中供應鏈分流（FCC 禁令等）", "否", "方向不確定：可能有利非中國供應商", "美國對中國機器人零組件限制"),
    ]
    for i, k in enumerate(RISKS):
        for j, v in enumerate(k):
            RK.cell(4 + i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["D"].width = 60

    # ================= README（xlsx 內）=================
    def S(x, fmt="#,##0"):
        return T(x, fmt)
    v = lambda k: f"Val!B{VR[k]}"
    lines = [
        ('="哈默納科 Harmonic Drive Systems（6324.T）個股模型 ' + VER + '｜build_id "&Sens!D2&"｜資料截止日 2026-11-16｜成品日 2026-11-23｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'),
        "命題：諧波減速機需求＝工業機器人循環＋協作＋人形；訂單 ÷ 營收看循環位置；中國廠（綠的諧波）份額是主要威脅。",
        (f'="結論（條件式）：在 FY2026 以公司 11-11 上修指引為錨、增量營業利益率 "&{S("inc_m","0%")}&"、情境 B 的 CY2032 全球人形出貨 "&{S("hu32_B","#,##0")}&" 千台且 HDS 份額 "&{S("shr_hum_B","0.0%")}&"、同業 P/E 未加溢價（rel_k＝"&{S("rel_k","0.00")}&"）的假設下，加權目標價約 "&{S(v("tp"))}&" 日圓，相對現價 "&{S("px")}&" 日圓 "&{S(v("up"),"+0.0%;-0.0%")}&"，估值評等 "&{v("rating")}&"，投資建議 "&{v("rec")}&"。"'),
        (f'="DCF 目標價 "&{S(v("dcf_tp"))}&"、Comparable 目標價 "&{S(v("cmp_tp"))}&"（各半）；WACC "&{S("wacc","0.00%")}&"、ke "&{S("ke","0.00%")}&"、g "&{S("g_tv","0.00%")}&"；情境 B 終值占 EV "&{S(f"Val!C{VR['tvs']}","0%")}&"。"'),
        (f'="上檔空間歸因：DCF 貢獻 "&{S(v("att_d"),"+0.0%;-0.0%")}&"、未調整 Comparable 貢獻 "&{S(v("att_c"),"+0.0%;-0.0%")}&"、相對倍數調整貢獻 "&{S(v("att_r"),"+0.0%;-0.0%")}&"。HDS 現價對 FY2026 指引 EPS 的倍數是同業中位數的 "&{S(v("relnow"),"0.0")}&" 倍；支持現價需要相對倍數係數 "&{S(v("relreq"),"0.0")}&"，升一級需要 "&{S(v("relflip"),"0.0")}&"。"'),
        (f'="收支關卡：情境 B FY2026–FY2030 FCFF 扣股利合計 "&{S(v("gate_b"),"#,##0.0")}&" 十億日圓，"&{v("gate_t")}&"；距最近級距邊界 "&{S(v("dist"),"0.0%")}&"（不到 5 個百分點時須揭露）：目標價達 "&{S(v("up_b"))}&" 日圓即升一級。"'),
        "主要失效條件（可觀察）：(1) HDS 受注 ÷ 營收連兩季 < 1.0；(2) 中國分部營收續跌且綠的諧波營收加速；(3) Tesla／Figure 等非中國整機的旋轉關節改用行星減速機或中國諧波；(4) 人形出貨若遠超情境 B（Sens 與 Recon 反解）則本評等翻轉。",
        "方法：DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，WACC 折現，明確期到 FY2032，RONIC＝WACC）50%＋同業前瞻 P/E × 模型 FY2027 EPS 50%；三情境 25／50／25。",
        "限制：分部預測以 2026 年 4–6 月地區占比固定；用途拆分為 Assumed（公司用途別圖表本輪未取得）；同業 P/E 由原報價日按 2026-11-20 收盤等比例換算（EPS 未更新）；Siliconomics／Tokenomics／Wattonomics 協調不適用（實體 AI 模型）。詳見 repo README.md。",
    ]
    for i, t in enumerate(lines):
        RD.cell(1 + i, 1, t).alignment = WRAP
    RD.column_dimensions["A"].width = 160
    return wb, dict(VR=VR, KR=KR, PR=PR, LAST_IN=LAST_IN)


# ---------------- 建置、重算、讀值 ----------------
def recalc(paths, outdir):
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(LO + ["--outdir", outdir] + paths, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def read_val(path, meta):
    wb = openpyxl.load_workbook(path, data_only=True)
    V = wb["Val"]; VR = meta["VR"]
    g = lambda k, c="B": V[f"{c}{VR[k]}"].value
    return dict(tp=g("tp"), dcf=g("dcf_tp"), cmp=g("cmp_tp"), rating=g("rating"), up=g("up"))


def run_cases(cases, meta, sub):
    d = os.path.join(TMP, sub); shutil.rmtree(d, ignore_errors=True); raw = os.path.join(d, "raw"); rc = os.path.join(d, "rc"); os.makedirs(raw)
    paths = []
    for i, ov in enumerate(cases):
        wb, _ = make(ov); p = os.path.join(raw, f"c{i:03d}.xlsx"); wb.save(p); paths.append(p)
    recalc(paths, rc)
    return [read_val(os.path.join(rc, f"c{i:03d}.xlsx"), meta) for i in range(len(cases))]


def in_values(path, last):
    wb = openpyxl.load_workbook(path, data_only=True); ws = wb["In"]
    out = {}
    for r in range(4, last + 1):
        k = ws.cell(r, 1).value
        if isinstance(k, str) and ws.cell(r, 5).value is not None:
            out[k] = (ws.cell(r, 4).value, ws.cell(r, 5).value, ws.cell(r, 6).value, ws.cell(r, 2).value)
    return out


def write_readme(path, meta, sens):
    """repo README.md 由成品重算值產生（原則第 10 節第 51 條：不手寫數字）。"""
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; VR = meta["VR"]
    g = lambda k, c="B": V[f"{c}{VR[k]}"].value
    I = {}
    ws = wb["In"]
    for r in range(4, meta["LAST_IN"] + 1):
        if ws.cell(r, 1).value: I[ws.cell(r, 1).value] = ws.cell(r, 5).value
    PL = wb["PL"]; PR = meta["PR"]["B"]
    rev = [PL[f"{c}{PR['rev']}"].value for c in "EFGHIJK"]; opm = [PL[f"{c}{PR['opm']}"].value for c in "EFGHIJK"]
    eps = [PL[f"{c}{PR['eps']}"].value for c in "EFGHIJK"]; hum = [PL[f"{c}{PR['hum']}"].value for c in "EFGHIJK"]
    top = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base_tp"] - 1))[:8]
    sv = {a: b for a, b, _ in sens["solve"]}
    t = open(os.path.join(HERE, "README_template.md"), encoding="utf-8").read()
    rep = dict(
        build_id=sens["build_id"], tp=f"{g('tp'):,.0f}", dcf=f"{g('dcf_tp'):,.0f}", cmp=f"{g('cmp_tp'):,.0f}", px=f"{I['px']:,.0f}",
        up=f"{g('up'):+.1%}", rating=g("rating"), rec=g("rec"), gate=f"{g('gate_b'):,.1f}", gate_t=g("gate_t"), dist=f"{g('dist') * 100:.1f}",
        up_b=f"{g('up_b'):,.0f}" if isinstance(g("up_b"), (int, float)) else str(g("up_b")),
        wacc=f"{g('wacc'):.2%}", ke=f"{g('ke'):.2%}", gtv=f"{I['g_tv']:.2%}", beta=f"{I['beta']:.2f}",
        dA=f"{g('dcf', 'B'):,.0f}", dB=f"{g('dcf', 'C'):,.0f}", dC=f"{g('dcf', 'D'):,.0f}",
        cA=f"{g('cmp', 'B'):,.0f}", cB=f"{g('cmp', 'C'):,.0f}", cC=f"{g('cmp', 'D'):,.0f}",
        tvB=f"{g('tvs', 'C'):.0%}", tvA=f"{g('tvs', 'B'):.0%}", pe=f"{I['pe_peer']:.2f}", relnow=f"{g('relnow'):.1f}", relreq=f"{g('relreq'):.1f}",
        relflip=f"{g('relflip'):.2f}" if isinstance(g("relflip"), (int, float)) else str(g("relflip")),
        att_d=f"{g('att_d'):+.1%}", att_c=f"{g('att_c'):+.1%}", att_r=f"{g('att_r'):+.1%}",
        rev=" → ".join(f"{x:.1f}" for x in rev), opm=" → ".join(f"{x:.1%}" for x in opm), eps=" → ".join(f"{x:.0f}" for x in eps),
        hum=" → ".join(f"{x:.1f}" for x in hum),
        s1=f"{sens['solve'][0][1]:,.0f}", s2=f"{sens['solve'][1][1]:,.0f}", s3=f"{sens['solve'][2][1]:.0%}",
        bl1=f"{wb['Q']['D6'].value:.1f}", bl2=f"{wb['Q']['E6'].value:.1f}", oq1=f"{wb['Q']['D7'].value:.2f}", oq2=f"{wb['Q']['E7'].value:.2f}",
        up_lvl={"強力放空": "放空", "放空": "中立", "中立": "買進", "買進": "強力買進"}.get(g("rating"), "—"),
        near=("評等離邊界近（不到 5 個百分點），" if g("dist") < 0.05 else f"評等距最近級距邊界 {g('dist') * 100:.1f} 個百分點；"),
        sens_tbl="\n".join(f"| {c['drv']} | `{c['key']}` | {c['tp']:,.0f} | {c['tp'] / sens['base_tp'] - 1:+.1%} | {c['dcf']:,.0f} | {c['rating']} |" for c in top),
    )
    for k, v in rep.items():
        t = t.replace("{{" + k + "}}", str(v))
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write(t)


def main():
    os.makedirs(os.path.join(HERE, "dist"), exist_ok=True)
    shutil.rmtree(TMP, ignore_errors=True); os.makedirs(TMP)
    wb, meta = make(); base_raw = os.path.join(TMP, "base.xlsx"); wb.save(base_raw)
    recalc([base_raw], os.path.join(TMP, "base_rc"))
    base_rc = os.path.join(TMP, "base_rc", "base.xlsx")
    base = read_val(base_rc, meta)
    iv = in_values(base_rc, meta["LAST_IN"])
    wbv = openpyxl.load_workbook(base_rc, data_only=True)
    fp = wbv["Sens"]["D5"].value
    print("基準：", base)

    # 案例：In 頁有低／高區間的鍵
    specs = []
    for k, (lo, e, hi, lab) in iv.items():
        if k in ("scn_w",): continue
        for side, val in (("敏感度低", lo), ("敏感度高", hi)):
            if isinstance(val, (int, float)) and isinstance(e, (int, float)) and abs(val - e) > 1e-12:
                specs.append(dict(drv=f"{lab}（{side}）", set=f"{k}＝{val:.4g}（基準 {e:.4g}）", key=k, val=val, ov={k: val}))
    # 業務主因與情境（原則第 10 節第 32 條）
    ivE = {k: v[1] for k, v in iv.items()}
    extra = [
        ("情境 B 人形 CY2032 出貨改為情境 A 值", "hu32_B", ivE["hu32_A"]),
        ("情境 B 人形 CY2032 出貨改為情境 C 值", "hu32_B", ivE["hu32_C"]),
        ("情境 B 人形份額改為情境 A 值", "shr_hum_B", ivE["shr_hum_A"]),
        ("情境 B 人形份額改為情境 C 值", "shr_hum_B", ivE["shr_hum_C"]),
        ("情境 B 份額漂移改為 0（無中國侵蝕）", "shd_B", 0.0),
        ("情境 B 份額漂移改為情境 C 值", "shd_B", ivE["shd_C"]),
        ("情境 B FY2027 產業機器人年增改為情境 C 值（循環轉向）", "g27_rob_B", ivE["g27_rob_C"]),
        ("相對倍數係數＝1（基準即為 1）", "rel_k", 1.0),
    ]
    for drv, k, val in extra:
        specs.append(dict(drv=drv, set=f"{k}＝{val:.4g}", key=k, val=val, ov={k: val}))
    specs.append(dict(drv="情境權重全給 A", set="tp_wA＝1、tp_wB＝0、tp_wC＝0", key="tp_wA tp_wB tp_wC", val=None, ov={"tp_wA": 1, "tp_wB": 0, "tp_wC": 0}))
    specs.append(dict(drv="情境權重全給 C", set="tp_wA＝0、tp_wB＝0、tp_wC＝1", key="tp_wA tp_wB tp_wC", val=None, ov={"tp_wA": 0, "tp_wB": 0, "tp_wC": 1}))
    res = run_cases([s["ov"] for s in specs], meta, "sens")
    cases = []
    for s, r in zip(specs, res):
        cases.append(dict(drv=s["drv"], set=s["set"], key=s["key"], val=s["val"], tp=r["tp"], dcf=r["dcf"], rating=r["rating"]))
        print(f"  {s['drv'][:40]:40s} {r['tp']:9.0f} {r['tp'] / base['tp'] - 1:+.1%}")

    # 反解：情境 B 人形 CY2032 出貨（整本重算，對數二分）
    px = ivE["px"]
    def solve(key, target_tp, lo, hi, n=14):
        import math
        a, b = math.log(lo), math.log(hi)
        ra = run_cases([{key: hi}], meta, "solve")[0]["tp"]
        if ra < target_tp: return None
        for _ in range(n):
            m = (a + b) / 2
            t = run_cases([{key: math.exp(m)}], meta, "solve")[0]["tp"]
            if t < target_tp: a = m
            else: b = m
        return math.exp((a + b) / 2)
    up = base["up"]
    flip_tp = px * (1 + (-0.5 if up < -0.5 else -0.25 if up < -0.25 else 0.25 if up <= 0.25 else 0.5))
    s1 = solve("hu32_B", px, ivE["hu32_B"], 200000)
    s2 = solve("hu32_B", flip_tp, ivE["hu32_B"], 200000)
    s3 = solve("shr_hum_B", px, ivE["shr_hum_B"], 1.0)
    solve_rows = [
        ("支持現價（目標價＝現價）所需的情境 B CY2032 全球人形出貨（千台）", s1 if s1 else -1, "其他輸入不變；−1＝在 2 億台內無解" if not s1 else "其他輸入不變；基準 " + f"{ivE['hu32_B']:,.0f}"),
        ("評等升一級所需的情境 B CY2032 全球人形出貨（千台）", s2 if s2 else -1, f"升一級邊界目標價 {flip_tp:,.0f} 日圓"),
        ("支持現價所需的情境 B HDS 人形份額", s3 if s3 else -1, "其他輸入不變；−1＝份額 100% 仍無解" if not s3 else "其他輸入不變"),
    ]
    print("反解：", solve_rows)

    # build_id
    h = hashlib.sha1(json.dumps([list(map(str, x)) for x in IN], ensure_ascii=False).encode() + open(__file__, "rb").read()).hexdigest()[:8]
    build_id = f"{VER}+{h}"
    sens = dict(build_id=build_id, fp=fp, base_tp=base["tp"], base_dcf=base["dcf"], cases=cases, solve=solve_rows)
    wb, meta2 = make(sens=sens)
    raw = os.path.join(TMP, "final", os.path.basename(OUT)); os.makedirs(os.path.dirname(raw), exist_ok=True); wb.save(raw)
    recalc([raw], os.path.join(TMP, "final_rc"))
    for f in glob.glob(os.path.join(HERE, "dist", "*.xlsx")):
        if f != OUT: print("注意：dist 另有檔案", f)
    shutil.copy(os.path.join(TMP, "final_rc", os.path.basename(OUT)), OUT)
    fin = read_val(OUT, meta2)
    print("成品：", OUT, fin, "build_id", build_id)
    json.dump(dict(build_id=build_id, base=fin, solve=solve_rows), open(os.path.join(TMP, "summary.json"), "w"), ensure_ascii=False, indent=1)

    write_readme(OUT, meta2, sens)

    # 錯誤值掃描與核對程式
    wbv = openpyxl.load_workbook(OUT, data_only=True)
    ERR = ("#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!")
    errs = [f"{ws.title}!{c.coordinate}={c.value}" for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value in ERR or c.value.startswith("Err:"))]
    print("錯誤值：", len(errs), errs[:10])
    rc = 0
    for cmd in (["tools/in_layout/check_in_layout.py", OUT], ["tools/sens_coverage/check_sens_coverage.py", OUT, "--require-keycol"],
                ["tools/model_governance/check_governance.py", OUT], ["tools/model_governance/check_rating_gate.py", OUT, "--strict"],
                ["tools/segment_check/check_segments.py", HERE]):
        p = subprocess.run([sys.executable, os.path.join(ROOT, cmd[0])] + cmd[1:], cwd=ROOT)
        rc |= p.returncode
    if errs: rc |= 1
    sys.exit(rc)


if __name__ == "__main__":
    main()

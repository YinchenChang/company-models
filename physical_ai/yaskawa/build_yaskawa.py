# 安川電機（Yaskawa Electric，6506.T）個股模型 v0.1.1 建置腳本（v0.1.1：現價換 2026-11-20 收盤）
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md（FCFF＋WACC，明確期到 FY2032）。
# 實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics 協調檔（brief 第 2 條），相關規格條文不適用，見 README。
# 會計年度：安川 2 月底年結；本模型沿用公司命名 FY2026＝2026-03～2027-02（＝2027年2月期）。單位：十億日圓（¥B），每股為日圓。
# 用法：python3 build_yaskawa.py          → dist/20261123_安川電機個股模型_v0.1.1.xlsx（公式版）
#       之後以 LibreOffice 重算並核對（verify()：0 錯誤值、Python 鏡像與試算表目標價一致）。
import hashlib, json, os, subprocess, sys, shutil, tempfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
VER = "v0.1.1"
OUT = os.path.join(HERE, "dist", "20261123_安川電機個股模型_v0.1.1.xlsx")
BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
WARN = PatternFill("solid", fgColor="FCE4D6")

# ---------------------------------------------------------------- 出處
S_TAN26 = "2026年2月期 決算短信〔IFRS〕（連結）2026-04-10，https://www.yaskawa.co.jp/wp-content/uploads/2026/04/20260410.pdf"
S_TAN1Q = "2027年2月期 第1四半期 決算短信 2026-07-10，https://www.yaskawa.co.jp/wp-content/uploads/2026/07/20260710.pdf"
S_TAN2Q = "2027年2月期 第2四半期（中間期）決算短信 2026-10-09，https://www.yaskawa.co.jp/wp-content/uploads/2026/10/20261009.pdf"
S_HAIFU = "2027年2月期 第1四半期 決算説明資料 2026-07-10（受注推移頁，億日圓），https://www.yaskawa.co.jp/wp-content/uploads/2026/07/20260710_haifu.pdf"
S_QA1Q = "第1四半期 決算説明会 Q&A 2026-07，https://www.yaskawa.co.jp/wp-content/uploads/2026/07/261Q_QA_1.pdf"
S_TECH = "科技新報 2026-10-12（引述公司 2Q 決算説明資料：受注、分部通期預估），https://cdnfinance.technews.tw/2026/10/12/yaskawa-orders-hit-record-raises-profit-forecast"
S_DASH = "中期経営計画「Dash 35」2026-05-22，https://www.yaskawa.co.jp/wp-content/uploads/2026/05/20260522.pdf"
S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes_20261109_20261120.csv）"
S_PX_OLD = "gyokaidigest 安川電機 2027-Q2 報告頁（日次終値，聚合網站），https://www.gyokaidigest.com/en/companies/yaskawa/report/2027-Q2"
S_PX2 = "Yahoo Finance 日線 2026-11-16 收盤（physical_ai/data/closes_20261109_20261120.csv）"
S_RF = "日本 10 年期國債 2.82%（財務省 2026-11-02；與 repo accretech v0.3.6 同值，跨模型一致）"
S_CRP = "日本 CRP 0.91%（與 repo accretech v0.3.6 同值，Damodaran 國家風險溢酬）"
S_L2A = "線 2a 工業機器人框架 v0.1（/mnt/user-data/outputs/lines/線2a_工業機器人_框架_v0.1.md）"

# ---------------------------------------------------------------- 歷史（Hist 頁；¥B）
# 列：(鍵, 項目, FY2024, FY2025, H1 FY2025, Q1 FY2026, H1 FY2026, 標記, 出處)；H2 FY2025＝FY2025 − H1 FY2025、Q2 FY2026＝H1 − Q1 由公式。
HIST = [
    ("mc_rev", "運動控制 營收（外部）", 238.752, 236.053, 112.837, 67.635, 137.286),
    ("rob_rev", "機器人 營收（外部）", 237.413, 247.012, 119.204, 56.728, 118.404),
    ("se_rev", "系統工程 營收（外部）", 38.352, 38.744, 18.690, 9.816, 20.778),
    ("oth_rev", "其他 營收（外部）", 23.164, 20.311, 9.462, 4.802, 9.128),
    ("rev", "合併營收（揭露）", 537.682, 542.122, 260.195, 138.982, 285.598),
    ("mc_op", "運動控制 營業利益", 23.005, 24.384, 12.024, 7.562, 15.973),
    ("rob_op", "機器人 營業利益", 23.751, 20.418, 10.542, 0.888, 5.974),
    ("se_op", "系統工程 營業利益", 4.605, 4.989, 1.939, 1.921, 3.480),
    ("oth_op", "其他 營業利益", 1.591, 1.988, 0.935, 0.193, 0.575),
    ("adj_op", "調整額（全社費用等）", -2.797, -4.473, -2.107, -2.080, -3.944),
    ("op", "合併營業利益（揭露）", 50.156, 47.307, 23.334, 8.486, 22.058),
    ("pt", "稅前利益", 78.454, 49.563, 25.204, 8.506, 25.963),
    ("tax", "所得稅費用", 20.758, 12.981, None, None, None),
    ("ni", "歸屬母公司淨利", 56.987, 35.240, 18.247, 5.445, 18.988),
    ("da", "折舊攤銷及減損", 20.828, 21.094, None, 6.148, 13.148),
    ("cx", "資本支出（有形＋無形取得）", 37.304, 46.240, None, 7.607, 27.286),
    ("ocf", "營業活動現金流", 56.505, 52.170, None, 21.362, 45.628),
]
HSRC = {
    "FY": "FY2024、FY2025：" + S_TAN26,
    "H1": "H1 FY2025（前年同期欄）、H1 FY2026：" + S_TAN2Q,
    "Q1": "Q1 FY2026：" + S_TAN1Q,
}
# 受注（億日圓）：FY2024 Q1–Q4、FY2025 Q1–Q4、FY2026 Q1 取自 1Q 決算説明資料；FY2026 Q2 取自科技新報引述 2Q 資料
ORD_Q = ["FY24 Q1", "FY24 Q2", "FY24 Q3", "FY24 Q4", "FY25 Q1", "FY25 Q2", "FY25 Q3", "FY25 Q4", "FY26 Q1", "FY26 Q2"]
ORD_TOT = [1351, 1263, 1205, 1267, 1275, 1340, 1385, 1524, 1638, 1854]
ORD_MC = [568, 556, 506, 537, 554, 582, 679, 789, 821, 911]
ORD_ROB = [652, 546, 559, 549, 614, 583, 582, 508, 697, 776]

# ---------------------------------------------------------------- In 頁輸入
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源1, 來源2／來源誘因)
SEC = "sec"
IN = [
    (SEC, "1. 市場資料與日期"),
    ("px", "現價（收盤）", "円", None, 3180, None, "Verified", S_PX, "Yahoo 日線未還原股利；v0.1 用 2026-11-16 收盤 3,561（見 px_alt）"),
    ("px_alt", "對照：2026-11-16 收盤", "円", None, 3561, None, "Verified", S_PX2, "只作對照，不進模型；v0.1 現價，與 " + S_PX_OLD + " 同值"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, ""),
    ("tp_date", "目標時點", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 26 條）", ""),
    ("val_date", "評價日（FY2026 Q3 季末）", "日期", None, "=DATE(2026,11,30)", None, "Assumed", "取最接近現價日的季末；FY2026 剩餘 Dec–Feb 單獨一欄", "v0.1.1 股價基準日改 2026-11-20 後最接近的季末仍為 2026-11-30，不變"),
    ("bs_date", "資產負債表日", "日期", None, "=DATE(2026,8,31)", None, "Verified", S_TAN2Q, ""),
    (SEC, "2. 資產負債與股數（2026-08-31）"),
    ("cash_v", "現金及約當現金", "¥B", None, 66.083, None, "Verified", S_TAN2Q, ""),
    ("debt_v", "借款及公司債（流動 43.218＋非流動 66.084）", "¥B", None, 109.302, None, "Verified", S_TAN2Q, ""),
    ("lease_v", "租賃負債（流動 3.959＋非流動 12.083）", "¥B", None, 16.042, None, "Verified", S_TAN2Q, ""),
    ("nci_v", "非支配股東權益（帳面）", "¥B", None, 10.489, None, "Verified", S_TAN2Q, ""),
    ("inv_v", "權益法投資及有價證券（未取得，基準 0）", "¥B", 0, 0, 20, "Assumed", "未取得：決算短信摘要未列投資明細；高案 20 為假設上限", "前年部分出售煙台東星股份（利得 26.777），尚有持股但金額未取得"),
    ("sh_b", "流通股數（發行 266.690497 − 庫藏 7.251147）", "百萬股", None, "=266.690497-7.251147", None, "Derived", S_TAN2Q, ""),
    ("dil_f", "稀釋係數（基本 EPS ÷ 稀釋 EPS，H1）", "倍", None, "=73.21/73.11", None, "Derived", S_TAN2Q + "；基本 73.21、稀釋 73.11 円", "股份報酬（業績連動）以稀釋 EPS 隱含股數代理"),
    ("sh_d", "評價日股數（稀釋）", "百萬股", None, "=sh_b*dil_f", None, "Derived", "＝sh_b × dil_f", ""),
    ("stub_div", "Stub 期已宣告股利（中間配當 36 円 × 流通股）", "¥B", None, "=36*sh_b/1000", None, "Derived", S_TAN2Q + "；基準日 8-31、11 月支付（在評價日前）", ""),
    ("h1_fcff", "H1 FY2026 已實現自由現金流代理（營業 CF − 資本支出）", "¥B", None, "=45.628-27.286", None, "Derived", S_TAN2Q, "營業 CF 已扣利息與稅，作 FCFF 代理"),
    (SEC, "3. FY2026 公司指引（2026-10-09 修正）"),
    ("gd_rev", "FY2026 營收指引", "¥B", None, 600.0, None, "Interested-party", S_TAN2Q, "公司指引；下半年匯率假設 155 円／美元"),
    ("gd_op", "FY2026 營業利益指引", "¥B", None, 57.5, None, "Interested-party", S_TAN2Q, "公司指引（由 60.0 下修：核心系統切換影響出貨）"),
    ("gd_pt", "FY2026 稅前利益指引", "¥B", None, 65.5, None, "Interested-party", S_TAN2Q, "公司指引"),
    ("gd_ni", "FY2026 歸屬母公司淨利指引", "¥B", None, 47.5, None, "Interested-party", S_TAN2Q, "公司指引；只作 Recon 對照"),
    ("mc_fy26_rev", "運動控制 FY2026 營收（公司分部預估）", "¥B", None, 290.3, None, "Interested-party", S_TECH, "公司 2Q 説明資料分部預估，經媒體轉述；原文未取得"),
    ("mc_fy26_op", "運動控制 FY2026 營業利益（公司分部預估）", "¥B", None, 39.4, None, "Interested-party", S_TECH, "同上"),
    ("rob_fy26_rev", "機器人 FY2026 營收（公司分部預估）", "¥B", None, 251.0, None, "Interested-party", S_TECH, "同上"),
    ("rob_fy26_op", "機器人 FY2026 營業利益（公司分部預估）", "¥B", None, 17.8, None, "Interested-party", S_TECH, "同上；隱含利益率 7.1%（線 2a 同值）"),
    ("dps_26", "FY2026 每股股利指引", "円", None, 72, None, "Interested-party", S_TAN2Q, "公司指引（中間 36＋期末 36）"),
    (SEC, "4. 分部驅動：系統工程、其他、全社（三情境共用，無分化機制）"),
    ("se_g", "系統工程 營收年增（FY2027 起）", "%", 0.02, 0.04, 0.06, "Assumed", "FY2025 +1.0%、H1 FY2026 +11.2%（Verified）；取中間", ""),
    ("se_m", "系統工程 營業利益率（H2 FY2026 起）", "%", 0.10, 0.13, 0.16, "Assumed", "FY2025 12.9%、H1 FY2026 16.7%（Derived）", ""),
    ("oth_g", "其他 營收年增（FY2027 起）", "%", -0.03, 0.0, 0.03, "Assumed", "FY2025 −12.3%、H1 FY2026 −3.5%", ""),
    ("oth_m", "其他 營業利益率（H2 FY2026 起）", "%", 0.05, 0.07, 0.09, "Assumed", "FY2025 9.8%、H1 FY2026 6.3%", ""),
    ("corp_pct", "全社調整額 ÷ 合併營收（FY2027 起）", "%", -0.013, -0.010, -0.008, "Assumed", "FY2024 −0.52%、FY2025 −0.83%、H1 FY2026 −1.38%（含核心系統切換）", ""),
    ("g_lt", "分部營收年增收斂值（FY2032）", "%", 0.02, 0.03, 0.04, "Assumed", "FY2029–FY2032 由 FY2028 年增線性收斂", ""),
    (SEC, "5. 損益表以下與現金流"),
    ("nonop", "業外淨額（FY2027 起）", "¥B", 1.0, 3.0, 5.0, "Assumed", "FY2025 稅前 − 營業 2.256；H1 FY2026 3.905；FY2026 指引 8.0", ""),
    ("tax_r", "有效稅率", "%", 0.25, 0.27, 0.30, "Assumed", "FY2025 實績 26.2%（12.981 ÷ 49.563，Derived）", "日本法定約 30.6%"),
    ("nci_pct", "非支配股東損益 ÷ 稅後淨利", "%", 0.02, 0.035, 0.05, "Assumed", "FY2025 1.342 ÷ 36.582＝3.7%（Derived）", ""),
    ("payout", "配息率（FY2027 起）", "%", 0.40, 0.40, 0.50, "Interested-party", S_DASH, "Dash 35 目標 40% 以上；FY2025 實績 50%"),
    ("da_26", "FY2026 折舊攤銷", "¥B", None, "=13.148*2", None, "Derived", "H1 FY2026 13.148 × 2；" + S_TAN2Q, "核心系統（S/4HANA）年折舊增 0.7–0.8（" + S_QA1Q + "）"),
    ("da_pct", "折舊攤銷 ÷ 營收（FY2027 起）", "%", 0.040, 0.045, 0.050, "Assumed", "FY2025 3.9%、H1 FY2026 4.6%（Derived）", ""),
    ("cx_26", "FY2026 資本支出", "¥B", 45, 50, 55, "Assumed", "H1 FY2026 27.286（Verified）× 約 1.8", "公司全年資本支出計畫未取得"),
    ("cx_27", "資本支出 ÷ 營收（FY2027）", "%", 0.06, 0.07, 0.08, "Assumed", "FY2024 6.9%、FY2025 8.5%、H1 FY2026 9.6%（擴產）", ""),
    ("cx_ss", "資本支出 ÷ 營收（FY2032 穩態）", "%", 0.045, 0.050, 0.060, "Assumed", "須 ≥ 折舊率 ×（1＋g）", ""),
    ("wc_r", "營運資金增加 ÷ 營收增量", "%", 0.15, 0.25, 0.35, "Assumed", "未取得資產負債表營運資金明細；FY2025 現金流量表營運資金三項合計 +1.59（流入）與營收 +4.4 不足以估比率", "原則第 23 條：歷史比率未能建立，列為限制"),
    (SEC, "6. 折現率與永續"),
    ("rf", "無風險利率（日本 10 年期國債）", "%", None, 0.0282, None, "Verified", S_RF, "逾期（2026-11-02，相差超過 2 交易日）；跨模型一致"),
    ("beta_u", "無槓桿 β", "倍", 0.90, 1.10, 1.40, "Assumed", "自算產業 β 尚無「JP 工廠自動化」組（docs/plan/自算產業β_20261105.md）；未取得", "待自算：同業 6954、6506、6268、6324、6273、6861"),
    ("erp", "股權風險溢酬 ERP", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節預設 5%（4–6%）", ""),
    ("crp", "國家風險溢酬 CRP", "%", None, 0.0091, None, "Analogy", S_CRP, ""),
    ("kd_sp", "債務利差（稅前 kd＝rf＋利差）", "%", 0.003, 0.005, 0.010, "Assumed", "公司債有效利率未取得", ""),
    ("wacc_adj", "WACC 調整（Sens 用）", "%", -0.01, 0.0, 0.01, "Assumed", "DCF 規格第 4 節 WACC ±1pt", ""),
    ("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 規格：g＝MIN(3%, rf)", ""),
    ("ronic_sp", "RONIC − WACC", "%", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）", ""),
    (SEC, "7. 同業倍數（分部加總，EV ÷ NTM EBIT）與權重"),
    ("mult_mc", "運動控制 EV/EBIT", "倍", 12, 15, 18, "Assumed", "同日同業倍數未取得（Fanuc、Nabtesco、哈默納科、Keyence、SMC 同日報價與共識未能讀取）", "原則第 18 條：相對倍數調整 1.00（未調整：缺同日同業倍數）"),
    ("mult_rob", "機器人 EV/EBIT", "倍", 10, 13, 16, "Assumed", "同上", ""),
    ("mult_se", "系統工程 EV/EBIT", "倍", 8, 10, 12, "Assumed", "同上", ""),
    ("mult_oth", "其他 EV/EBIT", "倍", 6, 8, 10, "Assumed", "同上", ""),
    ("w_dcf", "DCF 權重", "%", 0.3, 0.5, 0.7, "Assumed", "原則第 4 節預設 DCF 50%＋分部加總 50%", ""),
    (SEC, "8. 命題驅動：運動控制、機器人（情境專屬列）"),
    ("mc_g27_A", "情境 A：運動控制 FY2027 營收年增", "%", None, 0.15, None, "Assumed", "Q2 FY2026 受注 911 億（+56%）年化 3,644 億 vs FY2026 營收 2,903 億", "受注÷營收 1.31（Derived）"),
    ("mc_g27_B", "情境 B：運動控制 FY2027 營收年增", "%", 0.06, 0.10, 0.14, "Assumed", "同上；半導體、資料中心驅動器帶動", ""),
    ("mc_g27_C", "情境 C：運動控制 FY2027 營收年增", "%", None, 0.03, None, "Assumed", "訂單於 FY2027 上半見頂", ""),
    ("mc_g28_A", "情境 A：運動控制 FY2028 營收年增", "%", None, 0.10, None, "Assumed", "", ""),
    ("mc_g28_B", "情境 B：運動控制 FY2028 營收年增", "%", 0.03, 0.06, 0.09, "Assumed", "", ""),
    ("mc_g28_C", "情境 C：運動控制 FY2028 營收年增", "%", None, -0.06, None, "Assumed", "循環下行（FY2023 型）", ""),
    ("mc_m_A", "情境 A：運動控制 營業利益率（FY2028 起）", "%", None, 0.17, None, "Assumed", "H2 FY2026 指引隱含 15.3%（Derived）", ""),
    ("mc_m_B", "情境 B：運動控制 營業利益率（FY2028 起）", "%", 0.13, 0.15, 0.17, "Assumed", "＝H2 FY2026 指引隱含水準", ""),
    ("mc_m_C", "情境 C：運動控制 營業利益率（FY2028 起）", "%", None, 0.115, None, "Assumed", "FY2024 9.6%、FY2025 10.3%", ""),
    ("rob_g27_A", "情境 A：機器人 FY2027 營收年增", "%", None, 0.12, None, "Assumed", "Q2 FY2026 受注 776 億（+33%）年化 3,104 億 vs FY2026 營收 2,510 億", "受注÷營收 1.26（Derived）"),
    ("rob_g27_B", "情境 B：機器人 FY2027 營收年增", "%", 0.04, 0.08, 0.12, "Assumed", "同上；半導體機器人約占受注 15%（" + S_QA1Q + "）", ""),
    ("rob_g27_C", "情境 C：機器人 FY2027 營收年增", "%", None, 0.02, None, "Assumed", "", ""),
    ("rob_g28_A", "情境 A：機器人 FY2028 營收年增", "%", None, 0.08, None, "Assumed", "", ""),
    ("rob_g28_B", "情境 B：機器人 FY2028 營收年增", "%", 0.02, 0.05, 0.08, "Assumed", "", ""),
    ("rob_g28_C", "情境 C：機器人 FY2028 營收年增", "%", None, -0.05, None, "Assumed", "", ""),
    ("rob_m_A", "情境 A：機器人 營業利益率（FY2028 起）", "%", None, 0.12, None, "Assumed", "回到 FY2022–23 高點附近（高點值未取得，假設）", ""),
    ("rob_m_B", "情境 B：機器人 營業利益率（FY2028 起）", "%", 0.08, 0.095, 0.11, "Assumed", "FY2024 10.0%、FY2025 8.3%、H2 FY2026 指引隱含 8.9%（Derived）", "命題承重變數；證偽測試見 Thesis 頁"),
    ("rob_m_C", "情境 C：機器人 營業利益率（FY2028 起）", "%", None, 0.065, None, "Assumed", "價格戰延續：H1 FY2026 5.0%、Q2 FY2026 8.2%", ""),
    ("scn_wA", "情境權重 A", "%", None, 0.25, None, "Assumed", "預設 25／50／25", ""),
    ("scn_wB", "情境權重 B", "%", None, 0.50, None, "Assumed", "", ""),
    ("scn_wC", "情境權重 C", "%", None, 0.25, None, "Assumed", "", ""),
]
SCN_KEYS = ["mc_g27", "mc_g28", "mc_m", "rob_g27", "rob_g28", "rob_m"]

YEARS = list(range(2024, 2033))  # FY2024..FY2032
YC = {y: CL(3 + i) for i, y in enumerate(YEARS)}  # C..K

# ---------------------------------------------------------------- Python 鏡像（與試算表同式；Sens／反解用）
H = {k: dict(zip(["fy24", "fy25", "h1_25", "q1_26", "h1_26"], v)) for k, _, *v in HIST}


def base_inputs():
    P = {}
    import datetime as _dt
    for row in IN:
        if row[0] == SEC:
            continue
        k, v = row[0], row[4]
        if isinstance(v, str):
            continue
        P[k] = v
    P["sh_b"] = 266.690497 - 7.251147
    P["dil_f"] = 73.21 / 73.11
    P["sh_d"] = P["sh_b"] * P["dil_f"]
    P["stub_div"] = 36 * P["sh_b"] / 1000
    P["h1_fcff"] = 45.628 - 27.286
    P["da_26"] = 13.148 * 2
    return P


def model(P, s):
    """回傳 dict：每列為 FY2024..FY2032 的 list（index 0..8）。"""
    n = len(YEARS)
    r = {k: [None] * n for k in ("mc", "rob", "se", "oth", "rev", "mc_op", "rob_op", "se_op", "oth_op", "corp", "op",
                                  "nonop", "pt", "tax", "np", "nci", "ni", "eps", "dps", "da", "cx", "dwc", "nopat", "fcff", "div", "fcf_div")}
    for i, y in enumerate((2024, 2025)):
        f = "fy24" if y == 2024 else "fy25"
        r["mc"][i], r["rob"][i], r["se"][i], r["oth"][i] = H["mc_rev"][f], H["rob_rev"][f], H["se_rev"][f], H["oth_rev"][f]
        r["mc_op"][i], r["rob_op"][i], r["se_op"][i], r["oth_op"][i], r["corp"][i] = H["mc_op"][f], H["rob_op"][f], H["se_op"][f], H["oth_op"][f], H["adj_op"][f]
        r["da"][i], r["cx"][i] = H["da"][f], H["cx"][f]
    # FY2026
    h1se, h1o = H["se_rev"]["h1_26"], H["oth_rev"]["h1_26"]
    rem = P["gd_rev"] - P["mc_fy26_rev"] - P["rob_fy26_rev"] - h1se - h1o
    se_h2, o_h2 = rem * h1se / (h1se + h1o), rem * h1o / (h1se + h1o)
    i = 2
    r["mc"][i], r["rob"][i], r["se"][i], r["oth"][i] = P["mc_fy26_rev"], P["rob_fy26_rev"], h1se + se_h2, h1o + o_h2
    r["mc_op"][i], r["rob_op"][i] = P["mc_fy26_op"], P["rob_fy26_op"]
    r["se_op"][i] = H["se_op"]["h1_26"] + se_h2 * P["se_m"]
    r["oth_op"][i] = H["oth_op"]["h1_26"] + o_h2 * P["oth_m"]
    r["corp"][i] = P["gd_op"] - r["mc_op"][i] - r["rob_op"][i] - r["se_op"][i] - r["oth_op"][i]
    r["da"][i], r["cx"][i] = P["da_26"], P["cx_26"]
    m26 = {"mc": P["mc_fy26_op"] / P["mc_fy26_rev"], "rob": P["rob_fy26_op"] / P["rob_fy26_rev"]}
    for i in range(3, n):
        for seg in ("mc", "rob"):
            g27, g28, mss = P[f"{seg}_g27_{s}"], P[f"{seg}_g28_{s}"], P[f"{seg}_m_{s}"]
            if i == 3:
                g = g27
            elif i == 4:
                g = g28
            else:
                g = g28 + (P["g_lt"] - g28) * (i - 4) / 4
            r[seg][i] = r[seg][i - 1] * (1 + g)
            m = (m26[seg] + mss) / 2 if i == 3 else mss
            r[seg + "_op"][i] = r[seg][i] * m
        r["se"][i] = r["se"][i - 1] * (1 + P["se_g"])
        r["oth"][i] = r["oth"][i - 1] * (1 + P["oth_g"])
        r["se_op"][i] = r["se"][i] * P["se_m"]
        r["oth_op"][i] = r["oth"][i] * P["oth_m"]
    for i in range(n):
        r["rev"][i] = r["mc"][i] + r["rob"][i] + r["se"][i] + r["oth"][i]
        if i >= 3:
            r["corp"][i] = r["rev"][i] * P["corp_pct"]
            r["da"][i] = r["rev"][i] * P["da_pct"]
            r["cx"][i] = r["rev"][i] * (P["cx_27"] + (P["cx_ss"] - P["cx_27"]) * (i - 3) / 5)
        r["op"][i] = r["mc_op"][i] + r["rob_op"][i] + r["se_op"][i] + r["oth_op"][i] + r["corp"][i]
    for i in range(2, n):
        r["nonop"][i] = (P["gd_pt"] - P["gd_op"]) if i == 2 else P["nonop"]
        r["pt"][i] = r["op"][i] + r["nonop"][i]
        r["tax"][i] = r["pt"][i] * P["tax_r"]
        r["np"][i] = r["pt"][i] - r["tax"][i]
        r["nci"][i] = r["np"][i] * P["nci_pct"]
        r["ni"][i] = r["np"][i] - r["nci"][i]
        r["eps"][i] = r["ni"][i] / P["sh_b"] * 1000
        r["dps"][i] = P["dps_26"] if i == 2 else r["eps"][i] * P["payout"]
        r["dwc"][i] = P["wc_r"] * (r["rev"][i] - r["rev"][i - 1])
        r["nopat"][i] = r["op"][i] * (1 - P["tax_r"])
        r["fcff"][i] = r["nopat"][i] + r["da"][i] - r["cx"][i] - r["dwc"][i]
        r["div"][i] = r["dps"][i] * P["sh_b"] / 1000
        r["fcf_div"][i] = r["fcff"][i] - r["div"][i]
    return r


def rates(P):
    E = P["px"] * P["sh_b"] / 1000
    D = P["debt_v"] + P["lease_v"]
    bl = P["beta_u"] * (1 + (1 - P["tax_r"]) * D / E)
    ke = P["rf"] + bl * (P["erp"] + P["crp"])
    kd = (P["rf"] + P["kd_sp"]) * (1 - P["tax_r"])
    w = (E * ke + D * kd) / (E + D) + P["wacc_adj"]
    g = min(P["g_cap"], P["rf"])
    ron = w + P["ronic_sp"]
    return dict(E=E, D=D, bl=bl, ke=ke, kd=kd, wacc=w, g=g, ronic=ron)


def valuation(P, tv_alt=False):
    R = rates(P)
    out = {}
    for s in "ABC":
        m = model(P, s)
        rem26 = (m["fcff"][2] - P["h1_fcff"]) / 2
        nc = P["cash_v"] - P["debt_v"] - P["lease_v"] - P["nci_v"] + P["inv_v"] + rem26 - P["stub_div"]
        pv = rem26 / (1 + R["wacc"]) ** 0.125
        for i in range(3, 9):
            pv += m["fcff"][i] / (1 + R["wacc"]) ** (0.25 + (i - 3) + 0.5)
        if tv_alt:
            tv = m["fcff"][8] * (1 + R["g"]) / (R["wacc"] - R["g"])
        else:
            tv = m["nopat"][8] * (1 + R["g"]) * (1 - R["g"] / R["ronic"]) / (R["wacc"] - R["g"])
        pvtv = tv / (1 + R["wacc"]) ** 6.25
        ev = pv + pvtv
        ps = (ev + nc) / P["sh_d"] * 1000
        dps_tgt = 36 + 0.5 * m["dps"][3]
        v12 = ps * (1 + R["ke"]) - dps_tgt
        # 分部加總
        ntm = {k: 0.25 * m[k][3] + 0.75 * m[k][4] for k in ("mc_op", "rob_op", "se_op", "oth_op", "corp")}
        segev = ntm["mc_op"] * P["mult_mc"] + ntm["rob_op"] * P["mult_rob"] + ntm["se_op"] * P["mult_se"] + ntm["oth_op"] * P["mult_oth"]
        wm = segev / (ntm["mc_op"] + ntm["rob_op"] + ntm["se_op"] + ntm["oth_op"])
        evc = segev + ntm["corp"] * wm
        comp = (evc + nc) / P["sh_d"] * 1000
        out[s] = dict(m=m, pv=pv, pvtv=pvtv, ev=ev, nc=nc, ps=ps, dcf=v12, comp=comp, tvsh=pvtv / ev,
                      gate=sum(m["fcf_div"][2:7]))
    w = {s: P["scn_w" + s] for s in "ABC"}
    dcf = sum(w[s] * out[s]["dcf"] for s in "ABC")
    comp = sum(w[s] * out[s]["comp"] for s in "ABC")
    tp = P["w_dcf"] * dcf + (1 - P["w_dcf"]) * comp
    return dict(R=R, S=out, dcf=dcf, comp=comp, tp=tp, up=tp / P["px"] - 1)


def tier(up):
    return "強力買進" if up > 0.5 else "買進" if up > 0.25 else "中立" if up >= -0.25 else "放空" if up >= -0.5 else "強力放空"


# ---------------------------------------------------------------- 試算表
def hdr(ws, r, labels, c0=1):
    for j, h in enumerate(labels):
        c = ws.cell(r, c0 + j, h)
        c.font = BOLD
        c.fill = HDR


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def build():
    P0 = base_inputs()
    V0 = valuation(P0)
    wb = openpyxl.Workbook()
    RD = wb.active
    RD.title = "README"
    for t in ("In", "Hist", "Thesis", "PL", "DCF", "Comp", "Val", "Sens", "Recon", "對帳", "Risk"):
        wb.create_sheet(t)
    # ---------------- In
    ws = wb["In"]
    ws["A1"] = ("In：藍字為輸入。基準欄 E 進主模型；D／F（敏感度低／高）只用於 Sens 頁（一次動一個參數），不進目標價。"
                "單位：十億日圓（¥B），每股為日圓。會計年度沿用公司命名：FY2026＝2026-03～2027-02。")
    hdr(ws, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源 1", "來源 2／來源誘因"])
    r = 4
    INROW = {}
    for row in IN:
        if row[0] == SEC:
            ws.cell(r, 1, row[1]).font = BOLD
            r += 1
            continue
        k, item, unit, lo, base, hi, tag, s1, s2 = row
        ws.cell(r, 1, k)
        ws.cell(r, 2, item)
        ws.cell(r, 3, unit)
        for c, v in ((4, lo), (5, base), (6, hi)):
            if v is not None:
                cc = ws.cell(r, c, v)
                if not (isinstance(v, str) and v.startswith("=") and c == 5 and tag in ("Derived",)):
                    cc.font = BLUE
                if unit == "%":
                    cc.number_format = "0.0%"
                elif unit == "日期":
                    cc.number_format = "yyyy-mm-dd"
                else:
                    cc.number_format = "#,##0.000" if unit in ("¥B", "百萬股") else "#,##0.00"
        ws.cell(r, 7, tag)
        ws.cell(r, 8, s1)
        ws.cell(r, 9, s2)
        dn(wb, k, f"In!$E${r}")
        INROW[k] = r
        r += 1
    # 情境驅動對照（純展示）
    r += 1
    ws.cell(r, 1, "Z. 情境驅動對照（純展示；修改請改上方藍字）").font = BOLD
    r += 1
    hdr(ws, r, ["鍵", "項目", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"])
    r += 1
    for k in SCN_KEYS:
        ws.cell(r, 1, k)
        ws.cell(r, 2, next(x[1] for x in IN if x[0] == k + "_B").replace("情境 B：", ""))
        ws.cell(r, 3, "%")
        for j, s in enumerate("ABC"):
            c = ws.cell(r, 4 + j, f"=E{INROW[k + '_' + s]}")
            c.number_format = "0.0%"
        ws.cell(r, 8, f"引用 In 頁 {k}_A／{k}_B／{k}_C")
        r += 1
    ws.cell(r, 1, "scn_w")
    ws.cell(r, 2, "情境權重")
    for j, s in enumerate("ABC"):
        ws.cell(r, 4 + j, f"=E{INROW['scn_w' + s]}").number_format = "0%"
    ws.cell(r, 8, "引用 In 頁 scn_wA／scn_wB／scn_wC；合計＝1")
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["H"].width = 60

    # ---------------- Hist
    hs = wb["Hist"]
    hs["A1"] = "Hist：歷史財務（十億日圓，IFRS；分部營收為外部營收）。H2 FY2025＝FY2025 − H1 FY2025、Q2 FY2026＝H1 FY2026 − Q1 FY2026（Derived）。"
    hdr(hs, 3, ["鍵", "項目", "FY2024（2025/2期）", "FY2025（2026/2期）", "H1 FY2025", "H2 FY2025", "Q1 FY2026", "Q2 FY2026", "H1 FY2026", "標記", "出處"])
    HR = {}
    for i, (k, lab, fy24, fy25, h125, q126, h126) in enumerate(HIST):
        rr = 4 + i
        HR[k] = rr
        hs.cell(rr, 1, k)
        hs.cell(rr, 2, lab)
        for c, v in ((3, fy24), (4, fy25), (5, h125), (7, q126), (9, h126)):
            if v is not None:
                hs.cell(rr, c, v).font = BLUE
        if fy25 is not None and h125 is not None:
            hs.cell(rr, 6, f"=D{rr}-E{rr}")
        if q126 is not None and h126 is not None:
            hs.cell(rr, 8, f"=I{rr}-G{rr}")
        hs.cell(rr, 10, "Verified")
        hs.cell(rr, 11, HSRC["FY"] + "；" + HSRC["H1"] + "；" + HSRC["Q1"])
        for c in range(3, 10):
            hs.cell(rr, c).number_format = "#,##0.000"
    rr = 4 + len(HIST) + 1
    hs.cell(rr, 1, "chk_rev")
    hs.cell(rr, 2, "檢查：分部合計 − 合併營收（歷史）")
    for c in range(3, 10):
        L = CL(c)
        hs.cell(rr, c, f"=ROUND({L}{HR['mc_rev']}+{L}{HR['rob_rev']}+{L}{HR['se_rev']}+{L}{HR['oth_rev']}-{L}{HR['rev']},2)")
    hs.cell(rr, 11, "四捨五入誤差（百萬日圓取整）容許 ±0.01")
    rr += 1
    hs.cell(rr, 1, "chk_op")
    hs.cell(rr, 2, "檢查：分部營業利益＋調整額 − 合併營業利益（歷史）")
    for c in range(3, 10):
        L = CL(c)
        hs.cell(rr, c, f"=ROUND({L}{HR['mc_op']}+{L}{HR['rob_op']}+{L}{HR['se_op']}+{L}{HR['oth_op']}+{L}{HR['adj_op']}-{L}{HR['op']},2)")
    hs.column_dimensions["B"].width = 36
    hs.column_dimensions["K"].width = 80

    # ---------------- Thesis
    th = wb["Thesis"]
    th["A1"] = "Thesis：命題模組——訂單強，但機器人利益率下滑是結構性（價格戰）還是一次性（歐洲重組、核心系統切換）？"
    th["A2"] = ("命題（條件式）：在 Q2 FY2026 受注÷營收維持 1.2 以上、且機器人利益率於 H2 FY2026 回到公司指引隱含的 8.9% 附近的前提下，"
                "訂單上行會在 FY2027 轉成營收與利益；若受注強而機器人利益率在 FY2027 仍低於 8%，代表利益率下滑是結構性，命題被證偽。")
    hdr(th, 4, ["季度受注（億日圓）"] + ORD_Q + ["標記／出處"])
    for j, (lab, arr) in enumerate((("合併受注", ORD_TOT), ("運動控制 受注", ORD_MC), ("機器人 受注", ORD_ROB))):
        th.cell(5 + j, 1, lab)
        for q, v in enumerate(arr):
            th.cell(5 + j, 2 + q, v).font = BLUE
        th.cell(5 + j, 12, "Verified：FY24 Q1–FY26 Q1 " + S_HAIFU + "；FY26 Q2 " + S_TECH + "（媒體轉述公司資料，YoY 與公司 +38%／+56%／+33% 一致）")
    th.cell(8, 1, "系統工程受注、地區別受注")
    th.cell(8, 2, "未取得（説明資料表格未能讀取）")
    th.cell(9, 1, "機器人 受注 YoY")
    th.cell(10, 1, "運動控制 受注 YoY")
    for q in range(4, 10):
        L, Lp = CL(2 + q), CL(2 + q - 4)
        th.cell(9, 2 + q, f"={L}7/{Lp}7-1").number_format = "0.0%"
        th.cell(10, 2 + q, f"={L}6/{Lp}6-1").number_format = "0.0%"
    # 受注 ÷ 營收
    hdr(th, 12, ["受注 ÷ 營收（Derived）", "H1 FY2025", "H2 FY2025", "Q1 FY2026", "Q2 FY2026", "說明"])
    th.cell(13, 1, "運動控制")
    th.cell(13, 2, f"=(F6+G6)/10/Hist!E{HR['mc_rev']}")
    th.cell(13, 3, f"=(H6+I6)/10/Hist!F{HR['mc_rev']}")
    th.cell(13, 4, f"=J6/10/Hist!G{HR['mc_rev']}")
    th.cell(13, 5, f"=K6/10/Hist!H{HR['mc_rev']}")
    th.cell(14, 1, "機器人")
    th.cell(14, 2, f"=(F7+G7)/10/Hist!E{HR['rob_rev']}")
    th.cell(14, 3, f"=(H7+I7)/10/Hist!F{HR['rob_rev']}")
    th.cell(14, 4, f"=J7/10/Hist!G{HR['rob_rev']}")
    th.cell(14, 5, f"=K7/10/Hist!H{HR['rob_rev']}")
    th.cell(13, 6, "營收為外部營收；受注是否含分部間交易未揭露，比率只看方向")
    for rr_ in (13, 14):
        for c in range(2, 6):
            th.cell(rr_, c).number_format = "0.00"
    # 機器人利益率
    hdr(th, 16, ["機器人 營業利益率", "FY2024", "FY2025", "H1 FY2025", "H2 FY2025", "Q1 FY2026", "Q2 FY2026", "FY2026 指引", "H2 FY2026 指引隱含", "說明"])
    th.cell(17, 1, "利益率")
    for c, L in ((2, "C"), (3, "D"), (4, "E"), (5, "F"), (6, "G"), (7, "H")):
        th.cell(17, c, f"=Hist!{L}{HR['rob_op']}/Hist!{L}{HR['rob_rev']}")
    th.cell(17, 8, "=rob_fy26_op/rob_fy26_rev")
    th.cell(17, 9, f"=(rob_fy26_op-Hist!I{HR['rob_op']})/(rob_fy26_rev-Hist!I{HR['rob_rev']})")
    th.cell(17, 10, "分部營業利益 ÷ 外部營收（Derived）")
    th.cell(18, 1, "Q1 FY2026 一次性項目加回後")
    th.cell(18, 6, f"=(Hist!G{HR['rob_op']}+1.0+1.0)/Hist!G{HR['rob_rev']}")
    th.cell(18, 10, "加回歐洲重組約 1.0 與核心系統切換約 1.0（¥B；" + S_QA1Q + "，Interested-party：公司口頭）")
    th.cell(19, 1, "運動控制 利益率")
    for c, L in ((2, "C"), (3, "D"), (4, "E"), (5, "F"), (6, "G"), (7, "H")):
        th.cell(19, c, f"=Hist!{L}{HR['mc_op']}/Hist!{L}{HR['mc_rev']}")
    th.cell(19, 8, "=mc_fy26_op/mc_fy26_rev")
    th.cell(19, 9, f"=(mc_fy26_op-Hist!I{HR['mc_op']})/(mc_fy26_rev-Hist!I{HR['mc_rev']})")
    for rr_ in (17, 18, 19):
        for c in range(2, 10):
            th.cell(rr_, c).number_format = "0.0%"
    # 證偽測試
    hdr(th, 21, ["證偽測試", "門檻", "最新讀數", "判定", "期限／下一個讀點", "說明"])
    tests = [
        ("T1 訂單強：機器人受注÷營收（Q2 FY2026）", 1.1, "=E14", '=IF(C22>B22,"成立（訂單強）","不成立")', "每季（1、4、7、10 月）", "命題前提"),
        ("T2 機器人利益率年減（Q2 FY2026 對 Q2 FY2025）", 0, f"=G17-(Hist!E{HR['rob_op']}-5.0)/(Hist!E{HR['rob_rev']}-55.6)", '=IF(C23<B23,"仍在下滑","止跌")', "每季",
         "Q2 FY2025 機器人由 H1 − Q1（Q1 營收 55.6、利益 5.0 取自說明資料，億日圓取整）推得"),
        ("T3 H2 FY2026 機器人利益率 ≥ 指引隱含", "=I17", "未出（FY2026 決算 2027-04）", "待判", "2027-04 決算", "若 H2 低於 8%，指引失效、偏情境 C"),
        ("T4 證偽：FY2027 機器人利益率 < 8% 且營收年增 > 5%", 0.08, "未出（FY2027 決算 2028-04）", "待判", "2028-04 決算", "成立即命題證偽：量增不增利為結構性"),
    ]
    for j, (a, b, c, d, e, f) in enumerate(tests):
        rr_ = 22 + j
        th.cell(rr_, 1, a)
        th.cell(rr_, 2, b)
        th.cell(rr_, 3, c)
        th.cell(rr_, 4, d)
        th.cell(rr_, 5, e)
        th.cell(rr_, 6, f)
    th["C23"].number_format = "0.0%"
    th["B24"].number_format = "0.0%"
    th["B25"].number_format = "0.0%"
    th.cell(27, 1, "人形機器人、AI 機器人（MOTOMAN NEXT）：Dash 35 列為成長方向，但無營收揭露，本版不另建子列（併入機器人分部成長率）。")
    th.column_dimensions["A"].width = 46

    # ---------------- PL（三情境區塊；B 在最上）
    pl = wb["PL"]
    pl["A1"] = "PL：分部營收 → 合併營收 → 分部營業利益 → 合併營業利益 → 淨利、EPS；現金流（FCFF）接在後。FY2024–FY2025 為實績（引用 Hist），FY2026 由 H1 實績＋公司指引，FY2027 起由 In 頁驅動。"
    PLR = {}

    def block(s, r0):
        R = {}
        sfx = "" if s == "B" else f"（情境 {s}）"
        pl.cell(r0, 1, f"情境 {s}" + ("（基準）" if s == "B" else "（樂觀）" if s == "A" else "（保守）")).font = BOLD
        hdr(pl, r0 + 1, ["項目", "說明"] + [f"FY{y}" + ("A" if y <= 2025 else "E") for y in YEARS])
        rows = ["mc", "rob", "se", "oth", "rev", "chk", "mc_g", "rob_g", "mc_op", "rob_op", "se_op", "oth_op", "corp", "op", "mc_m", "rob_m", "opm",
                "nonop", "pt", "tax", "np", "nci", "ni", "eps", "dps", "da", "cx", "dwc", "nopat", "fcff", "div", "fcf_div"]
        labels = {
            "mc": "運動控制 營收", "rob": "機器人 營收", "se": "系統工程 營收", "oth": "其他 營收", "rev": "合併營收",
            "chk": "檢查：分部合計 − 合併營收", "mc_g": "運動控制 營收年增", "rob_g": "機器人 營收年增",
            "mc_op": "運動控制 營業利益", "rob_op": "機器人 營業利益", "se_op": "系統工程 營業利益", "oth_op": "其他 營業利益",
            "corp": "調整額（全社）", "op": "合併營業利益（EBIT）", "mc_m": "運動控制 利益率", "rob_m": "機器人 利益率", "opm": "合併營業利益率",
            "nonop": "業外淨額", "pt": "稅前利益", "tax": "所得稅", "np": "稅後淨利", "nci": "非支配股東損益", "ni": "歸屬母公司淨利",
            "eps": "基本 EPS（円）", "dps": "每股股利（円）", "da": "折舊攤銷", "cx": "資本支出", "dwc": "營運資金增加", "nopat": "NOPAT",
            "fcff": "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "div": "現金股利", "fcf_div": "FCFF − 股利"}
        for j, k in enumerate(rows):
            R[k] = r0 + 2 + j
            pl.cell(R[k], 1, labels[k] + sfx)
        hc = {2024: "C", 2025: "D"}
        for y in YEARS:
            c = YC[y]
            i = y - 2024
            p = YC.get(y - 1)
            f = {}
            if y <= 2025:
                h = hc[y]
                for k, hk in (("mc", "mc_rev"), ("rob", "rob_rev"), ("se", "se_rev"), ("oth", "oth_rev"), ("mc_op", "mc_op"), ("rob_op", "rob_op"),
                              ("se_op", "se_op"), ("oth_op", "oth_op"), ("corp", "adj_op"), ("da", "da"), ("cx", "cx")):
                    f[k] = f"=Hist!{h}{HR[hk]}"
                f["chk"] = f"=ROUND({c}{R['rev']}-Hist!{h}{HR['rev']},2)"
                f["pt"] = f"=Hist!{h}{HR['pt']}"
                f["tax"] = f"=Hist!{h}{HR['tax']}"
                f["np"] = f"={c}{R['pt']}-{c}{R['tax']}"
                f["ni"] = f"=Hist!{h}{HR['ni']}"
                f["nci"] = f"={c}{R['np']}-{c}{R['ni']}"
                f["nonop"] = f"={c}{R['pt']}-{c}{R['op']}"
            elif y == 2026:
                se_h1, o_h1 = f"Hist!I{HR['se_rev']}", f"Hist!I{HR['oth_rev']}"
                rem = f"(gd_rev-mc_fy26_rev-rob_fy26_rev-{se_h1}-{o_h1})"
                f["mc"] = "=mc_fy26_rev"
                f["rob"] = "=rob_fy26_rev"
                f["se"] = f"={se_h1}+{rem}*{se_h1}/({se_h1}+{o_h1})"
                f["oth"] = f"={o_h1}+{rem}*{o_h1}/({se_h1}+{o_h1})"
                f["mc_op"] = "=mc_fy26_op"
                f["rob_op"] = "=rob_fy26_op"
                f["se_op"] = f"=Hist!I{HR['se_op']}+({c}{R['se']}-{se_h1})*se_m"
                f["oth_op"] = f"=Hist!I{HR['oth_op']}+({c}{R['oth']}-{o_h1})*oth_m"
                f["corp"] = f"=gd_op-{c}{R['mc_op']}-{c}{R['rob_op']}-{c}{R['se_op']}-{c}{R['oth_op']}"
                f["chk"] = f"=ROUND({c}{R['rev']}-gd_rev,2)"
                f["nonop"] = "=gd_pt-gd_op"
                f["da"] = "=da_26"
                f["cx"] = "=cx_26"
                f["dps"] = "=dps_26"
            else:
                for seg in ("mc", "rob"):
                    if y == 2027:
                        g = f"{seg}_g27_{s}"
                        m = f"({seg}_fy26_op/{seg}_fy26_rev+{seg}_m_{s})/2"
                    elif y == 2028:
                        g = f"{seg}_g28_{s}"
                        m = f"{seg}_m_{s}"
                    else:
                        g = f"({seg}_g28_{s}+(g_lt-{seg}_g28_{s})*{y - 2028}/4)"
                        m = f"{seg}_m_{s}"
                    f[seg] = f"={p}{R[seg]}*(1+{g})"
                    f[seg + "_op"] = f"={c}{R[seg]}*{m}"
                f["se"] = f"={p}{R['se']}*(1+se_g)"
                f["oth"] = f"={p}{R['oth']}*(1+oth_g)"
                f["se_op"] = f"={c}{R['se']}*se_m"
                f["oth_op"] = f"={c}{R['oth']}*oth_m"
                f["corp"] = f"={c}{R['rev']}*corp_pct"
                f["chk"] = f"=ROUND({c}{R['mc']}+{c}{R['rob']}+{c}{R['se']}+{c}{R['oth']}-{c}{R['rev']},6)"
                f["nonop"] = "=nonop"
                f["da"] = f"={c}{R['rev']}*da_pct"
                f["cx"] = f"={c}{R['rev']}*(cx_27+(cx_ss-cx_27)*{y - 2027}/5)"
                f["dps"] = f"={c}{R['eps']}*payout"
            f["rev"] = f"={c}{R['mc']}+{c}{R['rob']}+{c}{R['se']}+{c}{R['oth']}"
            f["op"] = f"={c}{R['mc_op']}+{c}{R['rob_op']}+{c}{R['se_op']}+{c}{R['oth_op']}+{c}{R['corp']}"
            f["mc_m"] = f"={c}{R['mc_op']}/{c}{R['mc']}"
            f["rob_m"] = f"={c}{R['rob_op']}/{c}{R['rob']}"
            f["opm"] = f"={c}{R['op']}/{c}{R['rev']}"
            if p:
                f["mc_g"] = f"={c}{R['mc']}/{p}{R['mc']}-1"
                f["rob_g"] = f"={c}{R['rob']}/{p}{R['rob']}-1"
            if y >= 2026:
                f["pt"] = f"={c}{R['op']}+{c}{R['nonop']}"
                f["tax"] = f"={c}{R['pt']}*tax_r"
                f["np"] = f"={c}{R['pt']}-{c}{R['tax']}"
                f["nci"] = f"={c}{R['np']}*nci_pct"
                f["ni"] = f"={c}{R['np']}-{c}{R['nci']}"
                f["eps"] = f"={c}{R['ni']}/sh_b*1000"
                f["dwc"] = f"=wc_r*({c}{R['rev']}-{p}{R['rev']})"
                f["nopat"] = f"={c}{R['op']}*(1-tax_r)"
                f["fcff"] = f"={c}{R['nopat']}+{c}{R['da']}-{c}{R['cx']}-{c}{R['dwc']}"
                f["div"] = f"={c}{R['dps']}*sh_b/1000"
                f["fcf_div"] = f"={c}{R['fcff']}-{c}{R['div']}"
            for k, v in f.items():
                cc = pl[f"{c}{R[k]}"]
                cc.value = v
                cc.number_format = "0.0%" if k in ("mc_g", "rob_g", "mc_m", "rob_m", "opm") else "#,##0.00" if k in ("eps", "dps") else "#,##0.0"
        pl.cell(R["chk"], 2, "歷史：對揭露合併營收；FY2026：對指引；預測：恆等")
        pl.cell(R["se"], 2, "FY2026＝H1 實績＋（指引營收 − 運動控制 − 機器人 − H1 系統工程與其他）按 H1 比例分配")
        pl.cell(R["corp"], 2, "FY2026＝指引營業利益 − 各分部（Derived 殘差）；FY2027 起＝營收 × corp_pct")
        pl.cell(R["nonop"], 2, "FY2026＝指引稅前 − 指引營業利益")
        pl.cell(R["fcff"], 2, "DCF 規格第 1 節；股權報酬不加回")
        return R

    PLR["B"] = block("B", 3)
    nrow = 3 + 2 + 32 + 2
    PLR["A"] = block("A", nrow)
    PLR["C"] = block("C", nrow + 36)
    pl.column_dimensions["A"].width = 40
    pl.column_dimensions["B"].width = 30

    # ---------------- DCF
    dc = wb["DCF"]
    dc["A1"] = ("DCF（FCFF＋WACC，DCF 標準規格 v1.0）：評價日 2026-11-30；FY2026 剩餘 Dec–Feb 單獨一欄（＝(FY2026 FCFF − H1 已實現) ÷ 2）；"
                "Sep–Nov 同額計入淨現金（stub）。期中慣例；終值＝NOPAT(FY2033) ×（1 − g ÷ RONIC）÷（WACC − g）。")
    dc["A3"] = "折現率"
    dc["A3"].font = BOLD
    rate_rows = [
        ("E_mv", "股權市值 E", "=px*sh_b/1000", "#,##0.0"),
        ("D_bv", "負債 D（借款＋租賃）", "=debt_v+lease_v", "#,##0.0"),
        ("beta_l", "β（重新加槓桿）", "=beta_u*(1+(1-tax_r)*D_bv/E_mv)", "0.000"),
        ("ke", "股權成本 ke", "=rf+beta_l*(erp+crp)", "0.00%"),
        ("kd_at", "稅後債務成本", "=(rf+kd_sp)*(1-tax_r)", "0.00%"),
        ("w_e", "權益占資本", "=E_mv/(E_mv+D_bv)", "0.0%"),
        ("wacc", "WACC", "=w_e*ke+(1-w_e)*kd_at+wacc_adj", "0.00%"),
        ("g_tv", "永續成長率 g＝MIN(g_cap, rf)", "=MIN(g_cap,rf)", "0.00%"),
        ("ronic", "RONIC", "=wacc+ronic_sp", "0.00%"),
    ]
    for j, (k, lab, fml, fmt) in enumerate(rate_rows):
        rr_ = 4 + j
        dc.cell(rr_, 1, lab)
        dc.cell(rr_, 2, fml).number_format = fmt
        dn(wb, k, f"DCF!$B${rr_}")
    # 各情境
    DR = {}
    r0 = 15
    for s in "ABC":
        P_ = PLR[s]
        dc.cell(r0, 1, f"情境 {s}").font = BOLD
        hdr(dc, r0 + 1, ["項目", "FY2026 剩餘"] + [f"FY{y}" for y in range(2027, 2033)] + ["終值"])
        rows = {"fcff": r0 + 2, "t": r0 + 3, "df": r0 + 4, "pv": r0 + 5}
        dc.cell(rows["fcff"], 1, "FCFF")
        dc.cell(rows["t"], 1, "折現年數（期中）")
        dc.cell(rows["df"], 1, "折現因子")
        dc.cell(rows["pv"], 1, "現值")
        dc.cell(rows["fcff"], 2, f"=(PL!E{P_['fcff']}-h1_fcff)/2")
        dc.cell(rows["t"], 2, 0.125)
        for j, y in enumerate(range(2027, 2033)):
            c = CL(3 + j)
            dc[f"{c}{rows['fcff']}"] = f"=PL!{YC[y]}{P_['fcff']}"
            dc[f"{c}{rows['t']}"] = 0.25 + j + 0.5
        dc.cell(rows["fcff"], 9, f"=PL!K{P_['nopat']}*(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)")
        dc.cell(rows["t"], 9, 6.25)
        for cidx in range(2, 10):
            c = CL(cidx)
            dc[f"{c}{rows['df']}"] = f"=1/(1+wacc)^{c}{rows['t']}"
            dc[f"{c}{rows['pv']}"] = f"={c}{rows['fcff']}*{c}{rows['df']}"
            for k in ("fcff", "pv"):
                dc[f"{c}{rows[k]}"].number_format = "#,##0.0"
            dc[f"{c}{rows['df']}"].number_format = "0.0000"
        q = r0 + 6
        items = [
            ("pv_exp", "明確期現值", f"=SUM(B{rows['pv']}:H{rows['pv']})", "#,##0.0"),
            ("pv_tv", "終值現值", f"=I{rows['pv']}", "#,##0.0"),
            ("ev", "企業價值 EV", f"=B{q}+B{q + 1}", "#,##0.0"),
            ("nc_bs", "＋淨現金與投資（資產負債表）", "=cash_v-debt_v-lease_v-nci_v+inv_v", "#,##0.0"),
            ("stub", "＋Stub 期 FCFF（Sep–Nov）", f"=B{rows['fcff']}", "#,##0.0"),
            ("sdiv", "− Stub 期已宣告股利", "=stub_div", "#,##0.0"),
            ("nc", "淨現金滾到評價日", f"=B{q + 3}+B{q + 4}-B{q + 5}", "#,##0.0"),
            ("eq", "股權價值", f"=B{q + 2}+B{q + 6}", "#,##0.0"),
            ("ps", "每股價值（評價日，円）", f"=B{q + 7}/sh_d*1000", "#,##0"),
            ("dtgt", "期間每股股利（FY2026 期末 36＋FY2027 中間）", f"=36+0.5*PL!F{P_['dps']}", "#,##0.0"),
            ("v12", "DCF 情境價值推到 12 個月後（円）", f"=B{q + 8}*(1+ke)-B{q + 9}", "#,##0"),
            ("tvsh", "終值占 EV", f"=B{q + 1}/B{q + 2}", "0.0%"),
            ("tvchk", "檢查：終值占 EV ≤ 80% 且 EV > 0", f'=IF(AND(B{q + 2}>0,B{q + 11}<=0.8),"通過","超過 80% 或 EV≤0")', None),
            ("fchk", "檢查：FCFF 各列加總 − FCFF（FY2027）", f"=ROUND(PL!F{P_['nopat']}+PL!F{P_['da']}-PL!F{P_['cx']}-PL!F{P_['dwc']}-PL!F{P_['fcff']},6)", "0.000000"),
        ]
        DR[s] = {}
        for j, (k, lab, fml, fmt) in enumerate(items):
            rr_ = q + j
            dc.cell(rr_, 1, lab)
            cc = dc.cell(rr_, 2, fml)
            if fmt:
                cc.number_format = fmt
            DR[s][k] = rr_
        r0 = q + len(items) + 2
    dc.column_dimensions["A"].width = 44

    # ---------------- Comp（分部加總）
    cp = wb["Comp"]
    cp["A1"] = ("Comp：分部加總（EV ÷ NTM EBIT）。NTM＝目標時點（2027-11）往後 12 個月 ≈ 0.25 × FY2027＋0.75 × FY2028。"
                "倍數為 [Assumed]：同日同業倍數未取得；相對倍數調整 1.00（未調整：缺同日同業倍數）。全社費用以分部加權倍數資本化。")
    CR = {}
    r0 = 3
    for s in "ABC":
        P_ = PLR[s]
        cp.cell(r0, 1, f"情境 {s}").font = BOLD
        hdr(cp, r0 + 1, ["分部", "NTM EBIT", "倍數", "EV"])
        segs = [("運動控制", "mc_op", "mult_mc"), ("機器人", "rob_op", "mult_rob"), ("系統工程", "se_op", "mult_se"), ("其他", "oth_op", "mult_oth")]
        for j, (lab, k, mk) in enumerate(segs):
            rr_ = r0 + 2 + j
            cp.cell(rr_, 1, lab)
            cp.cell(rr_, 2, f"=0.25*PL!F{P_[k]}+0.75*PL!G{P_[k]}").number_format = "#,##0.0"
            cp.cell(rr_, 3, f"={mk}").number_format = "0.0"
            cp.cell(rr_, 4, f"=B{rr_}*C{rr_}").number_format = "#,##0.0"
        a, b = r0 + 2, r0 + 5
        rr_ = r0 + 6
        cp.cell(rr_, 1, "全社費用（分部加權倍數）")
        cp.cell(rr_, 2, f"=0.25*PL!F{P_['corp']}+0.75*PL!G{P_['corp']}").number_format = "#,##0.0"
        cp.cell(rr_, 3, f"=SUM(D{a}:D{b})/SUM(B{a}:B{b})").number_format = "0.0"
        cp.cell(rr_, 4, f"=B{rr_}*C{rr_}").number_format = "#,##0.0"
        cp.cell(rr_ + 1, 1, "企業價值（分部加總）")
        cp.cell(rr_ + 1, 4, f"=SUM(D{a}:D{rr_})").number_format = "#,##0.0"
        cp.cell(rr_ + 2, 1, "＋淨現金滾到評價日（同 DCF）")
        cp.cell(rr_ + 2, 4, f"=DCF!B{DR[s]['nc']}").number_format = "#,##0.0"
        cp.cell(rr_ + 3, 1, "隱含 EV ÷ NTM EBIT")
        cp.cell(rr_ + 3, 4, f"=D{rr_ + 1}/SUM(B{a}:B{rr_})").number_format = "0.0"
        cp.cell(rr_ + 4, 1, "分部加總每股價值（円）")
        cp.cell(rr_ + 4, 4, f"=(D{rr_ + 1}+D{rr_ + 2})/sh_d*1000").number_format = "#,##0"
        CR[s] = rr_ + 4
        r0 = rr_ + 6
    cp.column_dimensions["A"].width = 34

    # ---------------- Val
    va = wb["Val"]
    va["A1"] = "Val：DCF 50%＋分部加總 50%；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）。"
    hdr(va, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    va.cell(4, 1, "情境權重")
    for j, s in enumerate("ABC"):
        va.cell(4, 2 + j, f"=scn_w{s}").number_format = "0%"
    va.cell(4, 5, "=SUM(B4:D4)").number_format = "0%"
    va.cell(5, 1, "DCF 情境價值推到 12 個月後（円）")
    va.cell(6, 1, "分部加總每股價值（円）")
    for j, s in enumerate("ABC"):
        va.cell(5, 2 + j, f"=DCF!B{DR[s]['v12']}").number_format = "#,##0"
        va.cell(6, 2 + j, f"=Comp!D{CR[s]}").number_format = "#,##0"
    va.cell(5, 5, "=SUMPRODUCT(B4:D4,B5:D5)").number_format = "#,##0"
    va.cell(6, 5, "=SUMPRODUCT(B4:D4,B6:D6)").number_format = "#,##0"
    va.cell(5, 6, "DCF 目標價（情境加權）")
    va.cell(6, 6, "分部加總目標價（情境加權；倍數 Assumed，未做相對倍數調整）")
    rows_v = [
        ("DCF 目標價（情境加權）", "=E5", "#,##0"),
        ("分部加總目標價（情境加權）", "=E6", "#,##0"),
        ("DCF 權重", "=w_dcf", "0%"),
        ("加權目標價（円）", "=w_dcf*E5+(1-w_dcf)*E6", "#,##0"),
        ("現價（円）", "=px", "#,##0"),
        ("上檔空間", "=B10/B11-1", "0.0%"),
        ("估值評等（只看上檔空間）", '=IF(B12>0.5,"強力買進",IF(B12>0.25,"買進",IF(B12>=-0.25,"中立",IF(B12>=-0.5,"放空","強力放空"))))', None),
        ("距最近級距邊界（百分點）", "=MIN(ABS(B12-0.5),ABS(B12-0.25),ABS(B12+0.25),ABS(B12+0.5))*100", "0.0"),
        ("入不敷出判定用（情境 B，FY2026–FY2030 FCFF − 股利合計，¥B）", f"=SUM(PL!E{PLR['B']['fcf_div']}:I{PLR['B']['fcf_div']})", "#,##0.0"),
        ("入不敷出判定", '=IF(B15<0,"不通過（入不敷出）","通過")', None),
        ("投資建議", '=IF(AND(B15<0,OR(B13="買進",B13="強力買進")),"收支不足",B13)', None),
        ("評價日", "=val_date", "yyyy-mm-dd"),
        ("股價基準日", "=px_date", "yyyy-mm-dd"),
        ("目標時點", "=tp_date", "yyyy-mm-dd"),
        ("模型折現期間（評價日 → 目標時點，年）", "=(tp_date-val_date)/365", "0.00"),
        ("持有期間（現價日 → 目標時點，年）", "=(tp_date-px_date)/365", "0.00"),
        ("上檔空間三段歸因：DCF 貢獻（百分點）", "=w_dcf*(E5/px-1)*100", "0.0"),
        ("上檔空間三段歸因：分部加總貢獻（未調整，百分點）", "=(1-w_dcf)*(E6/px-1)*100", "0.0"),
        ("上檔空間三段歸因：相對倍數調整貢獻（百分點）", 0, "0.0"),
    ]
    VR = {}
    for j, (lab, fml, fmt) in enumerate(rows_v):
        rr_ = 7 + j
        va.cell(rr_, 1, lab)
        cc = va.cell(rr_, 2, fml)
        if fmt:
            cc.number_format = fmt
        VR[lab] = rr_
    va.cell(VR["上檔空間三段歸因：相對倍數調整貢獻（百分點）"], 3, "1.00（未調整：缺同日同業倍數）")
    rr_ = 7 + len(rows_v) + 1
    va.cell(rr_, 1, "收支關卡（FY2026–FY2030 FCFF − 股利合計，¥B）")
    for j, s in enumerate("ABC"):
        va.cell(rr_, 2 + j, f"=SUM(PL!E{PLR[s]['fcf_div']}:I{PLR[s]['fcf_div']})").number_format = "#,##0.0"
    VR["gate"] = rr_
    rr_ += 1
    va.cell(rr_, 1, "期間最低年度 FCFF − 股利（情境 B，¥B）")
    va.cell(rr_, 3, f"=MIN(PL!E{PLR['B']['fcf_div']}:I{PLR['B']['fcf_div']})").number_format = "#,##0.0"
    rr_ += 2
    va.cell(rr_, 1, "情境區間（全部成員 MIN／MAX）")
    va.cell(rr_, 2, "=MIN(B5:D6)").number_format = "#,##0"
    va.cell(rr_, 3, "=MAX(B5:D6)").number_format = "#,##0"
    va.cell(rr_, 4, '=IF(AND(B10>=B' + str(rr_) + ',B10<=C' + str(rr_) + '),"基準在區間內","基準在區間外")')
    rr_ += 1
    va.cell(rr_, 1, "投資建議句（條件式）")
    va.cell(rr_, 2, '="在機器人利益率依情境 B 回到 "&TEXT(rob_m_B,"0.0%")&"、運動控制 "&TEXT(mc_m_B,"0.0%")&"、分部倍數為假設值的條件下，加權目標價 "&TEXT(B10,"#,##0")&" 円，上檔空間 "&TEXT(B12,"0.0%")&"，投資建議："&B17')
    va.column_dimensions["A"].width = 58

    # ---------------- Sens（Python 整本重算快照）
    se = wb["Sens"]
    bid = build_id()
    se["A1"] = f"Sens：一次動一個參數，整本重算（Python 鏡像 build_yaskawa.py:valuation，與試算表同式）。build_id {bid}"
    se["A2"] = ("讀法：數值為快照；核對列比對快照基準與 Val 頁即時目標價。px_alt 不列：只作對照，不進模型。"
                "情境專屬列 A／C 不列：已由三情境加權涵蓋。")
    hdr(se, 4, ["驅動", "設定", "鍵", "低案 加權目標價", "高案 加權目標價", "低案 相對基準", "高案 相對基準", "低案 DCF 單法", "高案 DCF 單法", "低案 估值評等", "高案 估值評等"])
    cases = []
    for row in IN:
        if row[0] == SEC:
            continue
        k, item, unit, lo, base, hi = row[:6]
        if (lo is not None or hi is not None) and not isinstance(base, str):
            cases.append((item, f"低 {lo}／高 {hi}", k, lo, hi))
    rr_ = 5
    se.cell(rr_, 1, "基準")
    se.cell(rr_, 3, "—")
    se.cell(rr_, 4, round(V0["tp"], 4)).number_format = "#,##0"
    se.cell(rr_, 8, round(V0["dcf"], 4)).number_format = "#,##0"
    se.cell(rr_, 10, tier(V0["up"]))
    rr_ += 1
    SENS = []
    for item, sett, k, lo, hi in cases:
        vals = []
        for v in (lo, hi):
            if v is None:
                vals.append(None)
                continue
            P = dict(P0)
            P[k] = v
            vals.append(valuation(P))
        se.cell(rr_, 1, item)
        se.cell(rr_, 2, sett)
        se.cell(rr_, 3, k)
        for j, Vx in enumerate(vals):
            if Vx is None:
                continue
            se.cell(rr_, 4 + j, round(Vx["tp"], 4)).number_format = "#,##0"
            se.cell(rr_, 6 + j, Vx["tp"] / V0["tp"] - 1).number_format = "0.0%"
            se.cell(rr_, 8 + j, round(Vx["dcf"], 4)).number_format = "#,##0"
            se.cell(rr_, 10 + j, tier(Vx["up"]))
        SENS.append((item, k, vals))
        rr_ += 1
    # 特殊案例
    Vt = valuation(P0, tv_alt=True)
    se.cell(rr_, 1, "終值改用 FCFF ×（1＋g）÷（WACC − g）")
    se.cell(rr_, 3, "tv_alt")
    se.cell(rr_, 4, round(Vt["tp"], 4)).number_format = "#,##0"
    se.cell(rr_, 6, Vt["tp"] / V0["tp"] - 1).number_format = "0.0%"
    se.cell(rr_, 8, round(Vt["dcf"], 4)).number_format = "#,##0"
    se.cell(rr_, 10, tier(Vt["up"]))
    rr_ += 1
    for lab, s in (("只用情境 A（樂觀）", "A"), ("只用情境 C（保守）", "C")):
        P = dict(P0)
        P.update({"scn_wA": 1.0 if s == "A" else 0.0, "scn_wB": 0.0, "scn_wC": 1.0 if s == "C" else 0.0})
        Vx = valuation(P)
        se.cell(rr_, 1, lab)
        se.cell(rr_, 3, "scn_wA" if s == "A" else "scn_wC")
        se.cell(rr_, 4, round(Vx["tp"], 4)).number_format = "#,##0"
        se.cell(rr_, 6, Vx["tp"] / V0["tp"] - 1).number_format = "0.0%"
        se.cell(rr_, 8, round(Vx["dcf"], 4)).number_format = "#,##0"
        se.cell(rr_, 10, tier(Vx["up"]))
        rr_ += 1
    P = dict(P0)
    P["px"] = P0["px_alt"]
    Vx = valuation(P)
    se.cell(rr_, 1, f"對照：2026-11-16 收盤 {P0['px_alt']:,.0f}（上檔空間；β 權重隨市值）")
    se.cell(rr_, 3, "px")
    se.cell(rr_, 4, round(Vx["tp"], 4)).number_format = "#,##0"
    se.cell(rr_, 6, Vx["tp"] / V0["tp"] - 1).number_format = "0.0%"
    se.cell(rr_, 10, tier(Vx["up"]))
    se.cell(rr_, 11, f"上檔空間 {Vx['up']:.1%}")
    rr_ += 2
    se.cell(rr_, 1, "檢查：快照基準 − Val 頁即時加權目標價")
    se.cell(rr_, 4, f"=ROUND(D5-Val!B{VR['加權目標價（円）']},2)")
    se.cell(rr_, 6, '=IF(ABS(D' + str(rr_) + ')<0.5,"一致","STALE")')
    SENS_CHK = rr_
    se.column_dimensions["A"].width = 44

    # ---------------- Recon
    rc = wb["Recon"]
    rc["A1"] = "Recon：模型 vs 公司指引、中期計畫與市場；反解支持現價所需的值。"
    hdr(rc, 3, ["項目", "模型（情境 B）", "公司或市場", "差距", "說明"])
    B = PLR["B"]
    rec = [
        ("FY2026 營收", f"=PL!E{B['rev']}", "=gd_rev", "依建構相同（FY2026 以指引為錨，Interested-party）"),
        ("FY2026 營業利益", f"=PL!E{B['op']}", "=gd_op", "同上"),
        ("FY2026 歸屬母公司淨利", f"=PL!E{B['ni']}", "=gd_ni", "模型以 tax_r、nci_pct 推得，與指引差異為稅率口徑"),
        ("FY2029 營收（Dash 35 目標 650，匯率 145 円／美元）", f"=PL!H{B['rev']}", 650, "公司目標 Interested-party（" + S_DASH + "）；模型匯率隱含現行水準，口徑不同"),
        ("FY2029 營業利益（Dash 35 目標 100）", f"=PL!H{B['op']}", 100, "同上；模型低於公司目標"),
        ("FY2029 營業利益率（Dash 35 目標 15.4%）", f"=PL!H{B['opm']}", 0.154, "同上"),
        ("FY2026 機器人利益率（公司分部預估 7.1%；線 2a 同值）", f"=PL!E{B['rob_m']}", "=rob_fy26_op/rob_fy26_rev", "依建構相同"),
    ]
    for j, (lab, m, c, note) in enumerate(rec):
        rr_ = 4 + j
        rc.cell(rr_, 1, lab)
        rc.cell(rr_, 2, m)
        rc.cell(rr_, 3, c)
        rc.cell(rr_, 4, f"=B{rr_}-C{rr_}")
        rc.cell(rr_, 5, note)
        fmt = "0.0%" if "率" in lab else "#,##0.0"
        for cidx in (2, 3, 4):
            rc.cell(rr_, cidx).number_format = fmt
    rr_ = 4 + len(rec) + 1
    rc.cell(rr_, 1, "市場對照：共識目標價 6,935 円、預估 PER 24.9 倍（kabuyoho，2026-09-02，財報前，只作歷史對照，未查核，不作依據）")
    rr_ += 2
    # 反解（Python 整本重算）
    rc.cell(rr_, 1, "反解（整本重算；一次動一個驅動）").font = BOLD
    hdr(rc, rr_ + 1, ["驅動", "基準", "目標價＝現價", "降一級（上檔 −25%）", "升一級（上檔 +25%）", "說明"])
    solves = []
    for k in ("rob_m_B", "mc_m_B", "beta_u", "mult_rob"):
        row = [k, P0[k]]
        for target in (0.0, -0.25, 0.25):
            row.append(solve(P0, k, target))
        solves.append(row)
    RSROW = {}
    for j, row in enumerate(solves):
        r2 = rr_ + 2 + j
        RSROW[row[0]] = r2
        for c, v in enumerate(row):
            cc = rc.cell(r2, 1 + c, v if v is not None else "區間內無解")
            if isinstance(v, float):
                cc.number_format = "0.0%" if row[0] in ("rob_m_B", "mc_m_B") else "0.00"
        rc.cell(r2, 6, "反解值是目標價＝現價（或級距邊界）的點，不是評等門檻；搜尋區間：利益率 0–30%、β 0.3–3、倍數 3–40")
    RSOLVE = solves
    rc.column_dimensions["A"].width = 56

    # ---------------- 對帳
    rb = wb["對帳"]
    rb["A1"] = "對帳：本模型用到的線層級變數（線框架 v0.1 取值，口徑相同；不同處照寫，對帳時處理）。"
    hdr(rb, 3, ["變數名", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處", "標記", "影響本模型哪些儲存格"])
    recs = [
        ("整機廠營業利益率（安川機器人）FY2026", "2a", "=rob_fy26_op/rob_fy26_rev", "%", "安川機器人分部營業利益 ÷ 外部營收；公司預估", S_L2A + "；" + S_TECH, "Interested-party（線 2a 標 Derived）",
         "PL!E 列機器人營業利益；In rob_fy26_op、rob_fy26_rev"),
        ("全球工業機器人裝機 2025", "2a", 603000, "台", "IFR 年度裝機", "IFR World Robotics 2026（經 " + S_L2A + "）", "Verified", "未直接進模型；作機器人營收年增合理性對照"),
        ("全球裝機預測 2026／2029", "2a", "65.5 萬／80.6 萬", "台", "IFR 預測", "同上", "Interested-party（IFR 為產業協會）", "對照 rob_g27、rob_g28 情境；2026→2029 年均約 +7%"),
        ("JARA 受注額 YoY（2026 年 7–9 月）", "2a", 0.466, "%", "日本機器人工業會會員受注金額", "JARA（經 " + S_L2A + "）", "Verified", "對照 Thesis T1；未進公式"),
        ("中國工業機器人產量 YoY（2026 年 1–9 月）", "2a", 0.291, "%", "國家統計局規模以上產量（含出口，與 IFR 裝機口徑不同）", "國家統計局（經 " + S_L2A + "）", "Verified", "對照機器人中國受注；未進公式"),
        ("中國裝機國產比 2025", "2a", 0.55, "%", "IFR 中國", "IFR 中國新聞稿（經 " + S_L2A + "）", "Verified", "情境 C（價格戰）證據；未進公式"),
        ("安川機器人受注 YoY（2026 年 6–8 月）", "2a", "=K7/G7-1", "%", "公司受注（億日圓）", S_TECH, "Verified（媒體轉述）", "Thesis 第 7 列；線 2a 引同值 +33%"),
        ("半導體用機器人占機器人受注", "2a", "約 15%", "%", "公司口頭（FY2026 全年預期）", S_QA1Q, "Interested-party", "rob_g27_B 說明；未進公式"),
        ("人形機器人每台 BOM 中伺服／致動器價值", "2b", "未使用", "—", "—", "—", "—", "本版未建人形子列（Dash 35 列為方向，無營收揭露）"),
    ]
    for j, row in enumerate(recs):
        for c, v in enumerate(row):
            cc = rb.cell(4 + j, 1 + c, v)
            if c == 2 and isinstance(v, (float, str)) and (isinstance(v, float) or str(v).startswith("=")) and row[3] == "%":
                cc.number_format = "0.0%"
    rb["C10"].value = "=Thesis!K7/Thesis!G7-1"
    rb.column_dimensions["A"].width = 40
    rb.column_dimensions["F"].width = 60

    # ---------------- Risk
    rk = wb["Risk"]
    rk["A1"] = "Risk：風險登錄（未建模的尾部風險；原則第 10 條第 10、19 條）"
    hdr(rk, 3, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [
        ("中國價格戰使機器人利益率結構性下滑", "情境 C（rob_m_C 6.5%）", "命題核心", "機器人分部季利益率；中國 6 軸均價（GGII）；IFR 國產比"),
        ("半導體／AI 資料中心需求循環反轉", "情境 C（FY2028 營收 −5～−6%）", "運動控制受注 +56% 主要來自半導體設備與資料中心驅動器", "運動控制受注 QoQ；日本工具機受注；SEMI 設備出貨"),
        ("匯率（日圓升值）", "否", "指引下半年 155 円／美元；Dash 35 以 145 計；敏感度原文欄位未能可靠讀取", "USD/JPY；公司匯率敏感度（待取得）"),
        ("核心系統（S/4HANA）切換再度影響出貨", "否（只在 FY2026 指引內）", "公司稱 8 月起穩定；FY2026 指引已下修 2.5", "Q3 FY2026 營收 vs 受注差距"),
        ("中東情勢成本增加", "否", "公司全年風險假設約 4.0（¥B）", "季度說明"),
        ("權益法投資與有價證券價值", "Sens（inv_v 0／20）", "金額未取得", "有價證券報告書"),
    ]
    for j, row in enumerate(risks):
        for c, v in enumerate(row):
            rk.cell(4 + j, 1 + c, v)
    rk.column_dimensions["A"].width = 40

    # ---------------- README 頁
    rd = RD
    rd["A1"] = f"安川電機（6506.T）個股模型 {VER}｜build_id {bid}"
    rd["A2"] = '="資料截止日 2026-10-09（2Q 決算短信）｜成品日 2026-11-23｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"｜評價日 "&TEXT(val_date,"yyyy-mm-dd")&"｜目標時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    rd["A3"] = ("命題：訂單創新高（Q2 FY2026 受注 1,854 億日圓 +38%，運動控制 +56%、機器人 +33%），但機器人利益率由 FY2024 10.0% 降到 H1 FY2026 5.0%。"
                "本模型把「機器人利益率能否回到 9–10%」當成承重變數與證偽測試（線 2a）。")
    rd["A4"] = f'="加權目標價 "&TEXT(Val!B{VR["加權目標價（円）"]},"#,##0")&" 円（DCF "&TEXT(Val!E5,"#,##0")&"、分部加總 "&TEXT(Val!E6,"#,##0")&"）｜現價 "&TEXT(px,"#,##0")&"｜上檔空間 "&TEXT(Val!B{VR["上檔空間"]},"0.0%")&"｜投資建議："&Val!B{VR["投資建議"]}'
    rd["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{VR["距最近級距邊界（百分點）"]},"0.0")&" 個百分點｜收支關卡（情境 B）"&TEXT(Val!B{VR["入不敷出判定用（情境 B，FY2026–FY2030 FCFF − 股利合計，¥B）"]},"#,##0.0")&" ¥B："&Val!B{VR["入不敷出判定"]}'
    rd["A6"] = f'="WACC "&TEXT(wacc,"0.00%")&"｜ke "&TEXT(ke,"0.00%")&"｜g "&TEXT(g_tv,"0.00%")&"｜終值占 EV（情境 B）"&TEXT(DCF!B{DR["B"]["tvsh"]},"0.0%")'
    rd["A7"] = "分部：運動控制、機器人、系統工程、其他（官方應報導部門，IFRS 外部營收）。檢查列：Hist!chk_rev、PL「檢查：分部合計 − 合併營收」。"
    rd["A8"] = "限制：同日同業倍數、β 自算、營運資金明細、權益法投資、匯率敏感度、FY2025 Q3／Q4 分部季資料未取得（詳 README.md）。"
    rd["A9"] = "不適用：Siliconomics／Tokenomics／Wattonomics 協調（實體 AI 模型不讀算力線協調檔）；自用 AI 讀法 2 不適用（無自用算力）。"
    rd["A10"] = f"Sens 快照核對：Sens!F{SENS_CHK}"
    # 展示層（v0.1.1 新增；HTML 用白話列，公式帶數字；只新增，不改上方各列與數值格）
    vb = lambda k: f"Val!B{VR[k]}"
    rd["A12"] = "展示層（HTML 用；白話，公式帶數字）"
    rd["A13"] = ('="訂單強：2026 年 6–8 月（Q2 FY2026）合併受注 "&TEXT(Thesis!K5,"#,##0")&" 億日圓，年增 "&TEXT(Thesis!K5/Thesis!G5-1,"0%")'
                 '&"；受注 ÷ 營收 機器人 "&TEXT(Thesis!E14,"0.00")&"、運動控制 "&TEXT(Thesis!E13,"0.00")&"，訂單領先營收。"')
    rd["A14"] = ('="機器人利益率：FY2024 "&TEXT(Thesis!B17,"0.0%")&"、FY2025 "&TEXT(Thesis!C17,"0.0%")&"、Q1 FY2026 "&TEXT(Thesis!F17,"0.0%")&"、Q2 FY2026 "&TEXT(Thesis!G17,"0.0%")'
                 '&"；公司指引隱含下半年 "&TEXT(Thesis!I17,"0.0%")&"。下滑是一次性還是價格戰，要看 2027-04 決算的下半年實績。"')
    rd["A15"] = (f'="估值：DCF 目標價 "&TEXT(Val!E5,"#,##0")&" 円、分部加總 "&TEXT(Val!E6,"#,##0")&" 円，各半加權 "&TEXT({vb("加權目標價（円）")},"#,##0")'
                 f'&" 円；對現價 "&TEXT(px,"#,##0")&" 円上檔空間 "&TEXT({vb("上檔空間")},"0.0%")&"，投資建議："&{vb("投資建議")}&"。"')
    rd["A16"] = (f'="現價隱含：要讓目標價等於現價，機器人穩態利益率需約 "&TEXT(Recon!C{RSROW["rob_m_B"]},"0.0%")&"（模型情境 B "&TEXT(rob_m_B,"0.0%")'
                 f'&"），或運動控制需約 "&TEXT(Recon!C{RSROW["mc_m_B"]},"0.0%")&"（模型 "&TEXT(mc_m_B,"0.0%")&"），或無槓桿 β 降到約 "&TEXT(Recon!C{RSROW["beta_u"]},"0.00")&"。"')
    rd["A17"] = (f'="收支：情境 B FY2026–FY2030 自由現金流扣股利合計 "&TEXT({vb("入不敷出判定用（情境 B，FY2026–FY2030 FCFF − 股利合計，¥B）")},"#,##0.0")'
                 f'&" 十億日圓（"&{vb("入不敷出判定")}&"）。"')
    rd.column_dimensions["A"].width = 160

    wb.save(OUT)
    return P0, V0, SENS, RSOLVE, VR, DR, PLR, bid


def sens_line(P0, V0, SENS):
    """README「Sens 影響最大」句由 Sens 快照產生（v0.1.1 起；原為手寫）。"""
    rows = []
    for item, k, vals in SENS:
        ch = [None if v is None else v["tp"] / V0["tp"] - 1 for v in vals]
        mx = max(abs(x) for x in ch if x is not None)
        flips = [lab + "翻" + tier(v["up"]) for lab, v in zip(("低案", "高案"), vals) if v is not None and tier(v["up"]) != tier(V0["up"])]
        rows.append((mx, item, ch, flips))
    rows.sort(key=lambda t: -t[0])
    parts = []
    for mx, item, ch, flips in rows[:7]:
        txt = "／".join("—" if x is None else f"{x:+.1%}" for x in ch)
        parts.append(f"{item}（低／高 {txt}" + ("，" + "、".join(flips) if flips else "") + "）")
    return "Sens 影響最大（加權目標價相對基準）：" + "、".join(parts) + "。"


def write_readme(P0, V0, RS, bid, SENS=()):
    S = V0["S"]
    R = V0["R"]
    mB = S["B"]["m"]
    f0 = lambda x: f"{x:,.0f}"
    pc = lambda x: f"{x:.1%}"
    rs = {r[0]: r for r in RS}
    alt = dict(P0); alt["px"] = P0["px_alt"]; Va = valuation(alt)
    L = []
    a = L.append
    a(f"# physical_ai/yaskawa/ — 安川電機（Yaskawa Electric，6506.T）個股模型 {VER}")
    a("")
    a(f"build_id `{bid}`｜資料截止日 2026-10-09（2Q 決算短信）｜成品日 2026-11-23｜股價基準日 2026-11-20｜評價日 2026-11-30｜目標時點 2027-11-20。本檔數字由 `build_yaskawa.py` 依 Python 鏡像產生，與 xlsx 同式（Sens 頁核對列＝0）。")
    a("")
    a(f"成品：`dist/{os.path.basename(OUT)}`（LibreOffice 重算後存檔，0 錯誤值）；HTML 展示層 `dist/{os.path.splitext(os.path.basename(OUT))[0]}.html`（`html.json` 由 `gen_html_json.py` 產生，版面規格 v1.3）。亞太掛牌（東證 Prime）。")
    a("")
    a("## 結論（條件式）")
    a("")
    a(f"在機器人利益率依情境 B 回到 {pc(P0['rob_m_B'])}、運動控制 {pc(P0['mc_m_B'])}、分部倍數為假設值（同日同業倍數未取得）的條件下：")
    a("")
    a("| 項目 | 數值 |")
    a("|---|---|")
    a(f"| DCF 目標價（情境加權） | {f0(V0['dcf'])} 円 |")
    a(f"| 分部加總目標價（情境加權，EV/NTM EBIT） | {f0(V0['comp'])} 円 |")
    a(f"| 加權目標價（DCF 50%＋分部加總 50%） | **{f0(V0['tp'])} 円** |")
    a(f"| 現價（2026-11-20 收盤，Yahoo Finance 日線） | {f0(P0['px'])} 円 |")
    a(f"| 上檔空間 | {pc(V0['up'])} |")
    a(f"| 估值評等／投資建議 | **{tier(V0['up'])}**（收支關卡通過：情境 B FY2026–FY2030 FCFF − 股利合計 {S['B']['gate']:.1f} ¥B） |")
    a(f"| 距放空邊界（−25%） | {abs(V0['up'] + 0.25) * 100:.1f} 個百分點；對照：2026-11-16 收盤 {f0(P0['px_alt'])} 時上檔空間 {pc(Va['up'])}（{tier(Va['up'])}） |")
    a(f"| 情境價值 A／B／C（DCF） | {f0(S['A']['dcf'])}／{f0(S['B']['dcf'])}／{f0(S['C']['dcf'])} 円 |")
    a(f"| 情境價值 A／B／C（分部加總） | {f0(S['A']['comp'])}／{f0(S['B']['comp'])}／{f0(S['C']['comp'])} 円 |")
    a(f"| WACC／ke／g／RONIC | {R['wacc']:.2%}／{R['ke']:.2%}／{R['g']:.2%}／{R['ronic']:.2%}（β 重新加槓桿 {R['bl']:.3f}） |")
    a(f"| 終值占 EV（情境 B） | {pc(S['B']['tvsh'])} |")
    a("")
    a("上檔空間歸因：DCF 貢獻 {:.1f} 個百分點、分部加總貢獻 {:.1f}、相對倍數調整 0（1.00，未調整：缺同日同業倍數）。DCF 以 RONIC＝WACC（成長價值中性）計，分部加總以假設倍數計，兩法對競爭優勢的隱含假設不同。".format(
        P0["w_dcf"] * (V0["dcf"] / P0["px"] - 1) * 100, (1 - P0["w_dcf"]) * (V0["comp"] / P0["px"] - 1) * 100))
    a("")
    a("反解（整本重算，一次動一個驅動）：")
    a("")
    a("| 驅動 | 基準 | 目標價＝現價 | 降到放空（上檔 −25%） |")
    a("|---|---|---|---|")
    for k, lab, fmt in (("rob_m_B", "機器人穩態利益率（情境 B）", pc), ("mc_m_B", "運動控制穩態利益率（情境 B）", pc), ("beta_u", "無槓桿 β", lambda x: f"{x:.2f}"), ("mult_rob", "機器人 EV/EBIT", lambda x: f"{x:.1f}")):
        r = rs[k]
        a(f"| {lab} | {fmt(r[1])} | {fmt(r[2]) if r[2] is not None else '區間內無解'} | {fmt(r[3]) if r[3] is not None else '區間內無解'} |")
    a("")
    a(f"讀法：要支持現價，機器人穩態利益率需約 {pc(rs['rob_m_B'][2])}（FY2024 實績 10.0%），或運動控制需約 {pc(rs['mc_m_B'][2])}——市場價格隱含的不是「利益率回到過去高點」，而是 Dash 35（FY2029 營業利益 1,000 億日圓、利益率 15.4%）以上的結構性提升，或比本模型更低的折現率。")
    a("")
    a("## 命題（線 2a 的證偽測試）")
    a("")
    a("- **命題**：安川訂單創新高（2026 年 6–8 月受注 1,854 億日圓 +38%；運動控制 911 億 +56%、機器人 776 億 +33%），驅動是半導體設備與資料中心（驅動器）；但機器人利益率由 FY2024 10.0%、FY2025 8.3% 降到 H1 FY2026 5.0%。問題是：利益率下滑是一次性（歐洲重組約 1.8、核心系統 S/4HANA 切換）還是結構性（中國價格戰、汽車需求平）？")
    a("- **命題模組**（Thesis 頁）：季度受注、受注 ÷ 營收、機器人與運動控制利益率序列、Q1 一次性加回、四個證偽測試。")
    a("- **目前讀數**：T1 訂單強成立（機器人受注 ÷ 營收 Q2 FY2026 1.26、運動控制 1.31）；T2 Q2 FY2026 機器人利益率 8.2%，對 Q2 FY2025 約 8.7% 仍略降（−0.5pt）；Q1 加回一次性後僅 5.1%，Q2 回到 8.2% 是改善訊號。")
    a("- **失效條件（可觀察）**：(1) H2 FY2026 機器人利益率低於 8%（公司指引隱含 8.9%；2027-04 決算）→ 偏情境 C；(2) FY2027 機器人利益率 < 8% 而營收年增 > 5% → 量增不增利為結構性，命題證偽；(3) 運動控制受注 QoQ 連兩季下降 → 半導體循環見頂，情境 C 的 FY2028 −6%。")
    a("- 人形機器人、AI 機器人（MOTOMAN NEXT）：Dash 35 列為方向，無營收揭露，本版不建子列（併入機器人成長率），對帳頁標 2b「未使用」。")
    a("")
    a("## 分部（官方應報導部門，IFRS 外部營收，¥B）")
    a("")
    a("| 分部 | FY2024 | FY2025 | H1 FY2026 | FY2026E（指引） | FY2029E 情境 B | 驅動 |")
    a("|---|---|---|---|---|---|---|")
    segs = [("運動控制", "mc_rev", "mc", "受注 → FY2027／FY2028 年增（情境專屬）＋穩態利益率；FY2026 用公司分部預估 290.3／39.4"),
            ("機器人", "rob_rev", "rob", "同上；FY2026 用公司分部預估 251.0／17.8（利益率 7.1%）"),
            ("系統工程", "se_rev", "se", "年增 se_g、利益率 se_m（三情境共用）"),
            ("其他", "oth_rev", "oth", "年增 oth_g、利益率 oth_m（三情境共用）")]
    for lab, hk, mk, d in segs:
        a(f"| {lab} | {H[hk]['fy24']:.1f} | {H[hk]['fy25']:.1f} | {H[hk]['h1_26']:.1f} | {mB[mk][2]:.1f} | {mB[mk][5]:.1f} | {d} |")
    a(f"| 合併營收 | {H['rev']['fy24']:.1f} | {H['rev']['fy25']:.1f} | {H['rev']['h1_26']:.1f} | {mB['rev'][2]:.1f} | {mB['rev'][5]:.1f} | 分部合計（核對列 0） |")
    a(f"| 合併營業利益 | {H['op']['fy24']:.1f} | {H['op']['fy25']:.1f} | {H['op']['h1_26']:.1f} | {mB['op'][2]:.1f} | {mB['op'][5]:.1f} | 分部利益＋全社調整（FY2026 為指引殘差） |")
    a("")
    a(f"FY2029 情境 B 營業利益 {mB['op'][5]:.1f}、利益率 {mB['op'][5] / mB['rev'][5]:.1%}，低於 Dash 35 目標（1,000 億、15.4%；匯率假設 145 円／美元），差距主要在利益率；營收高於目標（模型營收含 FY2026 的 155 円匯率水準與訂單上行）。")
    a("")
    a("## 主要假設（全為 Assumed，In 頁有低／高，Sens 頁皆有案例）")
    a("")
    a("| 鍵 | 情境 A／B／C | 依據 | 可證偽指標 |")
    a("|---|---|---|---|")
    for k, lab, ind in (("mc_g27", "運動控制 FY2027 年增", "運動控制季受注（每季決算説明資料）"), ("mc_g28", "運動控制 FY2028 年增", "日本工具機受注、SEMI 設備出貨"),
                        ("mc_m", "運動控制 穩態利益率", "運動控制分部季利益率"), ("rob_g27", "機器人 FY2027 年增", "機器人季受注、JARA 受注"),
                        ("rob_g28", "機器人 FY2028 年增", "IFR 裝機、中國工業機器人產量"), ("rob_m", "機器人 穩態利益率", "機器人分部季利益率；中國 6 軸均價")):
        a(f"| {k}_A／B／C | {lab}：{pc(P0[k + '_A'])}／{pc(P0[k + '_B'])}／{pc(P0[k + '_C'])} | 見 In 頁 | {ind} |")
    a(f"| mult_mc／rob／se／oth | EV/NTM EBIT {P0['mult_mc']}／{P0['mult_rob']}／{P0['mult_se']}／{P0['mult_oth']} 倍（區間 12–18／10–16／8–12／6–10） | 同日同業倍數未取得 | Fanuc、Nabtesco、哈默納科同日 EV/EBIT |")
    a(f"| beta_u | {P0['beta_u']}（0.9–1.4） | 自算產業 β 尚無 JP 工廠自動化組 | 依 docs/plan/自算產業β 方法自算 |")
    a(f"| wc_r | {pc(P0['wc_r'])}（15–35%） | 營運資金明細未取得 | 有報營運資金三項 ÷ 營收 |")
    a("")
    a("其他：rf 2.82%（日本 10 年期，2026-11-02，與 repo 其他日股模型同值，逾期）、ERP 5%、CRP 0.91%、稅率 27%、配息率 40%（Dash 35 ≥ 40%）、FY2026 資本支出 50（H1 27.3）。情境權重 25／50／25。")
    a("")
    a(sens_line(P0, V0, SENS))
    a("")
    a("## 限制與未取得資料")
    a("")
    a("- **同業倍數未取得**：Fanuc、Nabtesco、哈默納科、Keyence、SMC 同一日期的報價與共識 EPS 未能讀取（行情網站拒絕或快取過期），分部加總倍數全為 Assumed 區間；原則第 7、18 條未滿足，相對倍數調整 1.00（未調整）。待其他代理（Fanuc、Nabtesco、哈默納科）完成後以同日倍數替換。")
    a(f"- **現價**：2026-11-20 收盤 {f0(P0['px'])}（Yahoo Finance 日線，未還原股利）。v0.1 用的 gyokaidigest 2026-11-16 收盤 3,561 與 Yahoo 同日值一致；Sens 保留「對照：2026-11-16 收盤 {f0(P0['px_alt'])}」案例。")
    a("- **β 未自算**：沿用 Assumed 1.10（0.9–1.4）；rf 為 2026-11-02 值（逾期）。")
    a("- **營運資金歷史比率未建立**（原則第 23 條）：只用 Assumed 25% 的營收增量比率。權益法投資與有價證券金額未取得（基準 0、Sens 20）。")
    a("- **分部季資料**：FY2025 Q3、Q4 分部營收與利益未取得，歷史對帳以 FY2024、FY2025、H1／H2 FY2025、Q1／Q2 FY2026 代替最近四季。")
    a("- **受注**：系統工程受注、地區別受注絕對值未取得；Q2 FY2026 受注來自媒體轉述公司資料（YoY 與公司一致）；FY2026 運動控制、機器人分部預估亦來自媒體轉述（Interested-party）。")
    a("- **匯率敏感度**：說明資料表格欄位未能可靠讀取，未建模（列風險登錄）。")
    a("- **文字組字**：README 頁 A4–A6 與 Val 投資建議句由公式組字；A3 命題句的歷史數字為手寫（下一版改公式）。")
    a("- HTML 展示層（v1.3）：模型沒有營業成本、毛利、營業費用與淨利率、合併營收年增率列（分部利益直接到營業利益），標準損益表缺這些列；本版不為展示層補列。共識目標價只有文字對照（未查核），未進估值長條圖。")
    a("")
    a("## 不適用的規格條文")
    a("")
    a("- Siliconomics／Tokenomics／Wattonomics 協調、兩讀數（C4）、自用 AI 讀法 2（原則第 7 節、第 10 節第 52–54 條）：實體 AI 模型不讀算力線協調檔，安川也無自用算力。")
    a("- 融資可行性、股權瀑布（第 9、62 條）：無新債補缺口、無增資。")
    a("")
    a("## 對帳（線層級變數，xlsx「對帳」頁）")
    a("")
    a("使用線 2a 框架 v0.1 的口徑：安川機器人 FY2026 利益率 7.1%（本模型標 Interested-party，線 2a 標 Derived；同一數字、標記不同，對帳時處理）、IFR 裝機與預測、JARA 受注、中國產量、國產比、安川機器人受注 +33%、半導體機器人占受注約 15%。線 2b（人形 BOM）本版未使用。")
    a("")
    a("## 驗證")
    a("")
    a("- LibreOffice 重算 0 錯誤值；Python 鏡像與試算表加權目標價差 0。")
    a("- `tools/segment_check/check_segments.py physical_ai/yaskawa` 符合；`tools/in_layout/check_in_layout.py` 通過；`tools/sens_coverage/check_sens_coverage.py --require-keycol` 通過；`tools/model_governance/check_rating_gate.py --strict` 符合。")
    a("- 尚未由未參與製作的代理獨立核對數字（Project 規則第 10 條）。")
    a("- HTML：`tools/company_html/check_company_html.py physical_ai/yaskawa` 通過；`tools/company_html/check_layout_v13.py` 除 D 節缺列（營收年增率、營業成本、毛利、毛利率、營業費用、淨利率、淨利年增率：模型損益頁沒有這些列，未為展示層補列）外通過；`layout_check.py` 1280／390 無水平捲動。")
    a("")
    a("## 版本紀錄")
    a("")
    a("- v0.1（2026-11-16）：初版。現價 3,561（2026-11-16 收盤），加權目標價 2,852 円，上檔空間 -19.9%，中立。")
    a(f"- v0.1.1（2026-11-23）：現價由 3,561（2026-11-16 收盤，gyokaidigest）改為 {f0(P0['px'])}（2026-11-20 收盤，Yahoo Finance 日線）；股價基準日 2026-11-20、目標時點 2027-11-20（EDATE＋12 月），評價日仍 2026-11-30（最接近的季末）。"
      f"加權目標價 2,852 → {f0(V0['tp'])} 円（只經 WACC 的市值權重變動），上檔空間 -19.9% → {pc(V0['up'])}，評等 中立 → {tier(V0['up'])}；其他假設不變。同業：模型無同業報價格（分部倍數全為 Assumed），無換算。新增 HTML 展示層（html.json、gen_html_json.py）與 README 頁展示層列。")
    a("")
    a("## 出處")
    a("")
    for s_ in (S_TAN26, S_TAN1Q, S_TAN2Q, S_HAIFU, S_QA1Q, S_TECH, S_DASH, S_PX, S_PX2, S_PX_OLD, S_L2A):
        a(f"- {s_}")
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write("\n".join(L) + "\n")


def solve(P0, k, target, lo=None, hi=None):
    lo = lo if lo is not None else (0.0 if k.endswith("_m_B") else 0.3 if k == "beta_u" else 3.0)
    hi = hi if hi is not None else (0.30 if k.endswith("_m_B") else 3.0 if k == "beta_u" else 40.0)

    def f(x):
        P = dict(P0)
        P[k] = x
        return valuation(P)["up"] - target
    a, b = f(lo), f(hi)
    if a * b > 0:
        return None
    for _ in range(80):
        m = (lo + hi) / 2
        fm = f(m)
        if (fm > 0) == (a > 0):
            lo, a = m, fm
        else:
            hi = m
    return (lo + hi) / 2


def build_id():
    blob = json.dumps([x for x in IN], ensure_ascii=False, default=str) + json.dumps(HIST) + json.dumps([ORD_TOT, ORD_MC, ORD_ROB])
    return VER + "+" + hashlib.sha1(blob.encode()).hexdigest()[:8]


def recalc(path):
    tmp = tempfile.mkdtemp()
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, path], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    return os.path.join(tmp, os.path.basename(path))


def verify(path, V0, VR):
    rp = recalc(path)
    wb = openpyxl.load_workbook(rp, data_only=True)
    errs = []
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("#") and c.value[1:4].isupper():
                    errs.append(f"{ws.title}!{c.coordinate}={c.value}")
    va = wb["Val"]
    tp = va.cell(VR["加權目標價（円）"], 2).value
    print("錯誤格：", len(errs), errs[:10])
    print(f"試算表加權目標價 {tp:.2f}；Python 鏡像 {V0['tp']:.2f}；差 {tp - V0['tp']:.4f}")
    return rp, errs, tp


if __name__ == "__main__":
    P0, V0, SENS, RS, VR, DR, PLR, bid = build()
    print("built", OUT, bid)
    rp, errs, tp = verify(OUT, V0, VR)
    R = V0["R"]
    print({k: round(v, 4) for k, v in R.items()})
    for s in "ABC":
        x = V0["S"][s]
        print(s, "DCF", round(x["dcf"]), "Comp", round(x["comp"]), "TV%", round(x["tvsh"], 3), "gate", round(x["gate"], 1),
              "OP29", round(x["m"]["op"][5], 1), "rev29", round(x["m"]["rev"][5], 1))
    print("DCF", round(V0["dcf"]), "Comp", round(V0["comp"]), "TP", round(V0["tp"]), "up", round(V0["up"], 4), tier(V0["up"]))
    print("solves", RS)
    if errs or abs(tp - V0["tp"]) > 0.5:
        sys.exit(1)
    shutil.copyfile(rp, OUT)  # 成品存重算後版本（含快取值，核對程式讀得到）
    print("saved recalculated", OUT)
    write_readme(P0, V0, RS, bid, SENS)
    print("README.md written")

#!/usr/bin/env python3
"""Palo Alto Networks（PANW）個股模型 v0.1：命題（原型 P 按防護範圍計價）＋營收線 → 損益表 → FCFF DCF／同業本益比。
用法：python3 software_ai/paloalto/build_paloalto.py [--sens-only] [--no-sens]
  1. 建基準檔並以 LibreOffice 重算；2. 逐案整本重算（Sens）；3. 寫成品 dist/<日期>_Palo Alto Networks個股模型_v0.1.xlsx 並重算。
需要：python3＋openpyxl、LibreOffice（soffice）。暫存檔在 /home/claude/scratch/paloalto。"""
import copy, datetime, glob, hashlib, json, os, re, shutil, subprocess, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
SCR = "/home/claude/scratch/paloalto"
VER, VDATE = "0.1", "20261116"
OUTNAME = f"{VDATE}_Palo Alto Networks個股模型_v{VER}.xlsx"
TKJ = os.path.join(ROOT, "software_ai", "unit_economics", "data", "tk_snapshot.json")
TK = json.load(open(TKJ, encoding="utf-8"))
D = datetime.datetime
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); GREY = Font(color="666666")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
YEARS = list(range(2024, 2033)); COLS = {y: CL(3 + i) for i, y in enumerate(YEARS)}   # C..K
HC = {2022: "D", 2023: "E", 2024: "F", 2025: "G", 2026: "H"}                        # Hist 欄
SRC = {  # 出處縮寫
    "K26": "PANW 10-K FY2026（2026-09-10）https://www.sec.gov/Archives/edgar/data/1327567/000132756726000023/panw-20260731.htm",
    "R26Q4": "PANW 8-K 附件 99.1 FY2026 Q4 新聞稿（2026-09-01）https://www.sec.gov/Archives/edgar/data/0001327567/000132756726000019/ex991q426earningsrelease.htm",
    "R25Q4": "PANW FY2025 Q4 新聞稿 https://www.paloaltonetworks.com/company/press/2025/palo-alto-networks-reports-fiscal-fourth-quarter-and-fiscal-year-2025-financial-results",
    "R23Q4": "PANW 8-K 附件 99.1 FY2023 Q4 新聞稿 https://www.sec.gov/Archives/edgar/data/1327567/000132756723000020/ex991q423earningsrelease.htm",
    "R22Q4": "PANW 8-K 附件 99.1 FY2022 Q4 新聞稿 https://www.sec.gov/Archives/edgar/data/1327567/000132756722000020/ex991q422earningsrelease.htm",
    "SAS": "StockAnalysis.com 統計頁 https://stockanalysis.com/stocks/panw/statistics/（2026-11-13 收盤；頁面 2026-11-14 更新）",
    "SA": "StockAnalysis.com 預測頁（S&P Global）https://stockanalysis.com/stocks/panw/forecast/（2026-10-15）",
    "YF": "Yahoo Finance https://finance.yahoo.com/quote/PANW/key-statistics/（頁面未標日期）",
    "UE": "software_ai/unit_economics/build_unit_economics.py 的 P 原型輸入（Assumed 區間；變數名稱一致）",
}
INC_CO = "公司自訂指標或指引，有對資本市場溝通的動機"
INC_SELL = "賣方預估或目標價，與公司有承銷、研究業務往來"
INC_DATA = "資料商共識取自賣方預估"

# ====================================================================== 輸入清單
# (key, 項目, 單位, 低, 基準, 高, 標記, 來源, 說明, 來源誘因, 格式)
IN = []
def sec(t): IN.append(("sec", t))
def inp(key, item, unit, lo, base, hi, tag, src, note="", inc="", fmt="0.00"):
    IN.append((key, item, unit, lo, base, hi, tag, src, note, inc, fmt))
PCT, PCT1, N0, N2 = "0.0%", "0.00%", "#,##0", "#,##0.00"

sec("A. 日期、現價、股數與資產負債（評價日＝FY2026 年底 2026-07-31，最近一期法定資產負債表）")
inp("data_date", "資料截止日（本版最晚一筆輸入）", "日期", None, D(2026, 11, 16), None, "Derived", "研究包 data/facts_paloalto.csv 擷取日", "README 首列分列資料截止日、成品日、股價基準日（原則第 10 節第 41 條）", "", "yyyy-mm-dd")
inp("px_date", "股價基準日（現價日）", "日期", None, D(2026, 11, 13), None, "Verified", SRC["SAS"], "目標時點＝本值加 12 個月（第 26 條）", "", "yyyy-mm-dd")
inp("val_date", "評價日（FY2026 年底資產負債表日）", "日期", None, D(2026, 7, 31), None, "Verified", SRC["K26"] + "；資產負債表", "FY2027 Q1（2026-10-31）資產負債表尚未公布，沿用 FY2026 年底", "", "yyyy-mm-dd")
inp("tp_date", "目標時點＝EDATE(股價基準日, 12)", "日期", None, "=EDATE(px_date,12)", None, "Derived", "公式", "折現期間自評價日起、持有期間自現價日起，Val 頁並列", "", "yyyy-mm-dd")
inp("px", "現價（2026-11-13 收盤）", "US$", None, 322.24, None, "Verified", SRC["SAS"], "單一來源，待第二來源；Yahoo 頁面未標日期且值為 342.00，不採用。2026-09-24 媒體報導收盤 393.30（WS，見研究包）", "", N2)
inp("sh_out", "流通股數（2026-08-31，10-K 封面）", "百萬股", None, 818, None, "Verified", SRC["K26"] + "；封面", "", "", N0)
inp("rsu", "未歸屬 RSU（全額計入 DCF 股數）", "百萬股", None, 17, None, "Verified", SRC["K26"] + "；Note 15", "DCF 標準規格第 3 節：RSU 全額；沒收率未揭露", "", N0)
inp("psu", "未歸屬 PSU（績效型，達標機率無證據）", "百萬股", None, 8, None, "Verified", SRC["K26"] + "；Note 15", "只作敏感度（psu_in）", "", N0)
inp("psu_in", "PSU 計入比例", "比例", 0, 0, 1, "Assumed", "DCF 標準規格第 3 節（績效型獎酬只作敏感度）", "基準 0，高案全數計入", "", "0.00")
inp("cash_0", "現金、短期與長期投資（2026-07-31）", "US$M", None, "=h_cash_26+h_sti_26+h_lti_26", None, "Derived", SRC["R26Q4"] + "；資產負債表 2,514＋557＋4,835", "", "", N0)
inp("conv", "長期可轉換票據（公允價值，視為債務）", "US$M", None, "=h_conv_26", None, "Verified", SRC["R26Q4"] + "；資產負債表", "CyberArk 2030 可轉債剩餘本金約 1,100，換股價約 211.24，現價在價內；以公允價值 1,774 視為債務，換股 5.4 百萬股不再計入股數（避免重複）", "", N0)
inp("capcall", "Capped call 公允價值（非營業資產）", "US$M", None, 153, None, "Verified", SRC["K26"] + "；Note 3", "", "", N0)
inp("sh_q4", "FY2026 Q4 稀釋股數（期末水準）", "百萬股", None, 832, None, "Interested-party", SRC["R26Q4"] + "；調節表 non-GAAP 稀釋股數", "預測股數起點", INC_CO, N0)

sec("B. 折現率與終值")
inp("rf", "無風險利率（10 年美債）", "%", None, 0.0499, None, "Verified", "美國財政部每日殖利率 2026-10-27 4.99%（Meta v0.2.1 In 頁 rf，跨模型共用）", "與現價日相差逾 2 個交易日，標逾期（第 15 條）；跨模型共用值，下次全面改版一併更新", "", PCT1)
inp("beta", "β（Blume 調整後兩來源中位數）", "x", "=0.67*0.93+0.33", "=MEDIAN(0.67*0.97+0.33,0.67*0.93+0.33)", "=0.67*0.97+0.33", "Derived", SRC["SAS"] + "原始 0.97；" + SRC["YF"] + "原始 0.93", "低／高＝兩來源各自的 Blume 調整值", "", "0.000")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節（預設 5%，區間 4–6%）", "", "", PCT)
inp("kd", "稅前債務成本＝rf＋利差", "%", None, 0.0569, None, "Derived", "同 Meta v0.2.1 In 頁 kd（rf＋0.70%）", "債務只占資本約 0.7%，影響極小", "", PCT1)
inp("g", "永續成長率", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格第 3 節（預設 3%，區間 2–4%）", "", "", PCT)
inp("wacc_adj", "WACC 加減（Sens 用）", "%", -0.01, 0, 0.01, "Assumed", "DCF 標準規格第 4 節：WACC ±1pt", "基準 0", "", PCT)
inp("ronic_x", "RONIC 高於 WACC 的幅度", "%", None, 0, 0.05, "Assumed", "DCF 標準規格第 3 節：RONIC 預設＝WACC", "基準 0（新投資價值中性）", "", PCT)
inp("tv_alt", "終值改用 FCFF×(1＋g) 的替代法（0＝NOPAT 法）", "開關", None, 0, 1, "Assumed", "DCF 標準規格第 4 節", "", "", "0")

sec("C. 命題模組（原型 P 按防護範圍計價；變數名稱與 software_ai/unit_economics 一致，v0.2 改讀 unit_economics）")
inp("P_mod", "每受保護單位模組數", "個", 2, 4, 7, "Assumed", SRC["UE"], "PANW 不揭露受保護單位數與模組採用分布；只作比例尺（單位數由 NGS ARR 反推）", "", "0.0")
inp("P_px", "每模組每受保護單位年價", "US$/年", 8, 20, 40, "Assumed", SRC["UE"], "同上，只影響單位數與推論成本占比的尺度，不影響營收", "", N2)
inp("P_rho", "代理身分單價 ÷ AI 前單位單價", "x", 0.02, 0.10, 0.30, "Assumed", SRC["UE"], "代理身分以批量或用量計價，單價遠低於人類端點", "", "0.00")
inp("P_n26", "受保護代理身分數（每 1 個 AI 前單位）FY2026", "個", 0.02, 0.05, 0.15, "Assumed", SRC["UE"] + "；機器身分對人類身分約 82：1（CyberArk 調查，Interested-party）", "只計被納入計價者", "", "0.00")
inp("P_m26", "模組加購（占 AI 前單位營收）FY2026", "%", 0.0, 0.02, 0.05, "Assumed", SRC["UE"], "", "", PCT)
inp("P_e26", "AI 原生競爭造成的價格侵蝕 FY2026", "%", 0.0, 0.01, 0.03, "Assumed", SRC["UE"], "", "", PCT)
inp("P_k26", "AI 資安分析次數（每單位每年）FY2026", "次/年", 1, 4, 10, "Assumed", SRC["UE"], "任務固定為 Tokenomics 任務 3「單代理（工具迴圈）」", "", N0)
inp("P_tier", "分析使用的模型層級（1 Luna／2 Terra／3 Sol／4 Astra）", "序號", 1, 2, 3, "Assumed", SRC["UE"], "大量告警用中低層級", "", "0")
inp("g_px", "模型 API 市場單價年變化｜待 AI供應鏈 登錄", "%/年", -0.50, -0.25, 0.0, "Assumed", SRC["UE"], "Tokenomics 契約無未來單價路徑名稱，依工作單以區間作敏感度", "", PCT)
inp("g_tok", "每任務 token 數年變化", "%/年", 0.0, 0.20, 0.50, "Assumed", SRC["UE"], "代理迴圈加長、思考加深", "", PCT)
inp("ovh", "實際工作負載 token ÷ Tokenomics 參考工作負載", "x", 1.0, 3.0, 8.0, "Analogy", SRC["UE"], "重試、驗證、長上下文；類比證據見 unit_economics README", "", "0.0")
inp("n32_A", "情境 A：受保護代理身分數 FY2032", "個", None, 1.5, None, "Assumed", SRC["UE"] + "（P_n32 區間 0.20–3.00）", "樂觀：代理部署加速、身分安全被當作代理的必要閘門", "", "0.00")
inp("n32_B", "情境 B：受保護代理身分數 FY2032", "個", None, 0.8, None, "Assumed", SRC["UE"] + "（P_n32 基準）", "", "", "0.00")
inp("n32_C", "情境 C：受保護代理身分數 FY2032", "個", None, 0.2, None, "Assumed", SRC["UE"] + "（P_n32 下緣）", "保守：代理身分多由雲端與模型商內建、不另計價", "", "0.00")
inp("m32_A", "情境 A：模組加購 FY2032", "%", None, 0.18, None, "Assumed", SRC["UE"] + "（P_m32 區間 3–25%）", "樂觀：攻擊面擴大帶動平台化加購", "", PCT)
inp("m32_B", "情境 B：模組加購 FY2032", "%", None, 0.10, None, "Assumed", SRC["UE"] + "（P_m32 基準）", "", "", PCT)
inp("m32_C", "情境 C：模組加購 FY2032", "%", None, 0.03, None, "Assumed", SRC["UE"] + "（P_m32 下緣）", "", "", PCT)
inp("e32_A", "情境 A：價格侵蝕 FY2032", "%", None, 0.02, None, "Assumed", SRC["UE"] + "（P_e32 區間 0–15%）", "樂觀：AI 原生新進者未取得價格影響力", "", PCT)
inp("e32_B", "情境 B：價格侵蝕 FY2032", "%", None, 0.05, None, "Assumed", SRC["UE"] + "（P_e32 基準）", "", "", PCT)
inp("e32_C", "情境 C：價格侵蝕 FY2032", "%", None, 0.12, None, "Assumed", SRC["UE"] + "（P_e32 上緣附近）", "保守：模型商與雲端大廠把基礎安全功能內建、拉低單價", "", PCT)
inp("k32_A", "情境 A：AI 資安分析次數 FY2032", "次/年", None, 100, None, "Assumed", SRC["UE"] + "（P_k32 區間 10–150）", "樂觀情境代理活動多，分析次數與推論成本同步較高", "", N0)
inp("k32_B", "情境 B：AI 資安分析次數 FY2032", "次/年", None, 40, None, "Assumed", SRC["UE"] + "（P_k32 基準）", "", "", N0)
inp("k32_C", "情境 C：AI 資安分析次數 FY2032", "次/年", None, 15, None, "Assumed", SRC["UE"] + "（P_k32 下緣）", "", "", N0)

sec("D. 營收線驅動（Product、Subscription and support 之下的 NGS ARR 相關營收與其他）")
inp("gu27", "受保護單位數年增 FY2027", "%", 0.09, 0.13, 0.17, "Assumed", "類比：FY2026 有機營收年增 14.4%（研究包第 3 節，Derived）；NGS ARR 不含單位口徑揭露", "FY2027 起線性收斂到 FY2032；不以公司指引校準（FY2027 指引 NGS ARR 年增 22–23% 只在 Recon 對照）", "", PCT)
inp("gu32", "受保護單位數年增 FY2032", "%", 0.05, 0.08, 0.11, "Assumed", "同上", "", "", PCT)
inp("ng_ss", "NGS ARR 相關營收列入 Subscription and support 的比例", "%", 0.85, 0.92, 1.00, "Assumed", "NGS ARR 定義含 product 軟體授權（2026 委託書）；列入 Product 的比例未揭露", "餘額在 Product 路徑內", "", PCT)
inp("g_os", "其他 Subscription and support 年增（硬體支援、legacy、專業服務）", "%", -0.05, 0.03, 0.10, "Assumed", "無直接揭露；FY2026 以 S&S 減 NGS ARR 相關營收推得，約占 S&S 三成", "", "", PCT)
inp("gp27", "Product 營收年增 FY2027", "%", 0.08, 0.16, 0.26, "Assumed", "FY2026 Product +27%（含 CyberArk 部分年度授權）；FY2027 含全年度 CyberArk 授權", "FY2027 起線性收斂到 FY2032", "", PCT)
inp("gp32", "Product 營收年增 FY2032", "%", 0.02, 0.05, 0.08, "Assumed", "硬體更新週期與軟體授權混合；無直接證據", "", "", PCT)
inp("tail_dist", "經銷商流失（FY2028 一次性損失受保護單位比例）", "%", 0.0, 0.0, 0.10, "Assumed", SRC["K26"] + "；Item 1：兩家經銷商占 FY2026 營收 30%", "事件型尾端風險（第 19 條）：基準 0，只作用於 NGS ARR 相關營收", "", PCT)

sec("E. 利潤、現金流與股數（基準＝FY2026 實績，全期持平；第 35 條）")
inp("nm0", "Non-GAAP 營業利益率（扣 AI 增量後的核心率）", "%", 0.27, "=h_ngebit_26/h_rev_26", 0.33, "Interested-party", SRC["R26Q4"] + "；調節表（3,356÷11,480）", "公司自訂 non-GAAP；FY2027 指引 29.5% 只在 Recon 對照。低案含記憶體成本壓縮硬體利潤；高案接近管理層 FY2028 現金流目標所暗示的效率", INC_CO, PCT)
inp("nm_up", "Non-GAAP 營業利益率每年改善（Sens 用）", "pt/年", None, 0, 0.01, "Assumed", "管理層 FY2028 調整後自由現金流率 40% 目標（Interested-party）", "基準 0（無證據不入基準）", INC_CO, PCT)
inp("sbc_pct", "股票薪酬（SBC-related charges）占營收", "%", 0.12, "=h_sbc_26/h_rev_26", 0.17, "Derived", SRC["R26Q4"] + "；調節表 1,712÷11,480", "DCF 以 NOPAT 扣股票薪酬、不加回", "", PCT)
inp("sbc_dn", "SBC 占營收每年下降（Sens 用）", "pt/年", None, 0, 0.01, "Assumed", "公司自揭：SBC 占營收 FY2021 21.8% → FY2026 15.8%（2026 委託書，約 1.2pt／年）", "基準 0；README 指為 DCF 低估主因之一，故列 Sens", INC_CO, PCT)
inp("amort_run", "收購無形資產年攤銷（年化 Q4）", "US$M", 900, "=4*h_amq4_26", 1300, "Derived", SRC["R26Q4"] + "；Q4 FY2026 281×4", "受剩餘餘額限制（MIN）；終值年排除", "", N0)
inp("acq27", "併購整合與交易費用 FY2027", "US$M", 100, 250, 400, "Assumed", SRC["K26"] + "；FY2026 295，Embrace、Console 已完成", "", "", N0)
inp("acq28", "併購整合與交易費用 FY2028", "US$M", 0, 100, 200, "Assumed", "同上", "FY2029 起 0", "", N0)
inp("ysd", "現金與投資收益率（占期初餘額）", "%", 0.035, "=(h_oth_26+562)/AVERAGE(h_cashinv_25,cash_0)", 0.055, "Derived", SRC["R26Q4"] + "；其他收入淨額 −159＋公允價值損失 562＝403，除以平均現金與投資", "公允價值變動不預測", "", PCT)
inp("tax", "有效稅率（GAAP 與 non-GAAP 預測用）", "%", 0.18, 0.22, 0.26, "Derived", SRC["R26Q4"] + "；non-GAAP 稅率由調節表推得 22.0%", "FY2026 GAAP 有效稅率 42.7% 含可轉債公允價值變動，不用", "", PCT)
inp("tax_dcf", "DCF 現金稅率（NOPAT 用）", "%", 0.12, 0.22, 0.25, "Assumed", SRC["K26"] + "；FY2026 已付所得稅 199（約 5% 稅前）", "CyberArk 商譽大部分可抵稅，現金稅率可能長期低於模型稅率；第 40 條：差距大於 3 個百分點故列 Sens", "", PCT)
inp("dep_pct", "折舊（不含無形資產攤銷）占營收", "%", 0.015, "=(h_da_26-h_amort_26)/h_rev_26", 0.025, "Derived", SRC["K26"] + "；折舊與攤銷 855 減攤銷 638", "", "", PCT)
inp("cx_pct", "資本支出占營收（不含總部購地）", "%", 0.025, "=(h_cx_26-91)/h_rev_26", 0.04, "Derived", SRC["R26Q4"] + "；FCF 調節表 440 減總部購地 91", "", "", PCT)
inp("nwc_r", "營運資金淨流入 ÷ 營收增量（含遞延收入）", "比率", 0.15, 0.30, 0.45, "Assumed", "對照：OCF 減淨利、SBC、折舊攤銷所得殘差 ÷ 營收增量，FY2025 0.79、FY2026 0.72（Derived，含遞延稅項雜訊，未取得 10-K 現金流量表逐項）", "基準取低於對照；逐項待補（第 23 條）", "", "0.00")
inp("bb", "年度買回（占用現金，不計入 DCF）", "US$M", 0, 1000, 2000, "Assumed", SRC["K26"] + "；FY2026 買回 1,000，授權剩餘 1,000（至 2026-12-31）", "股數以 ke 推算買回價", "", N0)
inp("dil", "年度淨發行（股票薪酬）占股數", "%", 0.008, 0.018, 0.030, "Assumed", "無直接揭露；不以公司指引股數校準", "", "", PCT)
inp("cogs_sh", "營業成本占總成本（營業成本＋營業費用）", "比例", None, "=h_cogs_26/(h_cogs_26+h_opex_26)", None, "Derived", SRC["R26Q4"] + "；3,403÷(3,403＋7,382)", "只決定營業成本與營業費用的切分，不影響營業利益", "", "0.0000")

sec("F. 目標價規則與同業")
inp("w_dcf", "目標價中 DCF 的權重（其餘＝同業本益比）", "%", 0.3, 0.5, 0.7, "Assumed", "原則第 4 節（預設各半）", "", "", PCT)
inp("tp_wA", "情境權重 A", "%", None, 0.25, None, "Assumed", "原則第 8 節（25／50／25）", "", "", PCT)
inp("tp_wB", "情境權重 B", "%", None, 0.50, None, "Assumed", "同上", "", "", PCT)
inp("tp_wC", "情境權重 C", "%", None, 0.25, None, "Assumed", "同上", "", "", PCT)
inp("pe_f", "同業前瞻本益比 Fortinet（FTNT）", "x", None, 33.85, None, "Interested-party", "StockAnalysis.com FTNT 統計頁（2026-11-13 收盤）", "分母為共識 NTM EPS（資料商未標精確期間）；US GAAP", INC_DATA, "0.00")
inp("pe_z", "同業前瞻本益比 Zscaler（ZS）", "x", None, 39.22, None, "Interested-party", "StockAnalysis.com ZS 統計頁（同上）", "US GAAP", INC_DATA, "0.00")
inp("pe_c", "同業前瞻本益比 Check Point（CHKP）", "x", None, 13.70, None, "Interested-party", "StockAnalysis.com CHKP 統計頁（同上）", "US GAAP（以色列公司，申報 US GAAP）", INC_DATA, "0.00")
inp("pe_n", "同業前瞻本益比 Cloudflare（NET）", "x", None, 174.39, None, "Interested-party", "StockAnalysis.com NET 統計頁（同上）", "US GAAP", INC_DATA, "0.00")
inp("pe_p", "PANW 資料商前瞻本益比（展示；同源同日）", "x", None, 77.10, None, "Interested-party", SRC["SAS"], "只用於相對倍數高案與 Recon 展示", INC_DATA, "0.00")
inp("rel_adj", "相對倍數調整係數（未調整＝1.00）", "x", 0.8, 1.0, "=H_PEREL", "Assumed", "1.00（未調整：缺同業 NTM 成長率、利益率與同日 EPS；4 家同業不做迴歸）", "高案＝PANW 資料商前瞻本益比 ÷ 同業中位數（Val 頁）；列入承重輸入（第 18 條）", "", "0.00")

sec("G. 只作對照的公司指引、目標與共識（不進模型）")
for k, it, u, v, f, inc in [
    ("g_rev_lo", "FY2027 營收指引（低）", "US$M", 14100, N0, INC_CO), ("g_rev_hi", "FY2027 營收指引（高）", "US$M", 14200, N0, INC_CO),
    ("g_arr_lo", "FY2027 NGS ARR 指引（低）", "US$B", 11.075, "0.000", INC_CO), ("g_arr_hi", "FY2027 NGS ARR 指引（高）", "US$B", 11.175, "0.000", INC_CO),
    ("g_nm", "FY2027 Non-GAAP 營業利益率指引", "%", 0.295, PCT, INC_CO), ("g_eps_lo", "FY2027 Non-GAAP 稀釋 EPS 指引（低）", "US$", 4.16, N2, INC_CO),
    ("g_eps_hi", "FY2027 Non-GAAP 稀釋 EPS 指引（高）", "US$", 4.19, N2, INC_CO), ("g_fcfm", "FY2027 調整後自由現金流率指引", "%", 0.38, PCT, INC_CO),
    ("g_sh_lo", "FY2027 稀釋股數指引（低）", "百萬股", 844, N0, INC_CO), ("g_sh_hi", "FY2027 稀釋股數指引（高）", "百萬股", 847, N0, INC_CO),
    ("t_arr30", "FY2030 NGS ARR 管理層目標", "US$B", 20, N2, INC_CO), ("t_fcf28", "FY2028 調整後自由現金流率管理層目標", "%", 0.40, PCT, INC_CO)]:
    inp(k, it, u, None, v, None, "Interested-party", SRC["R26Q4"] + "；Financial Outlook／CEO、CFO 引言", "只作對照，不進模型", inc + "；只作對照，不進模型", f)
inp("acq_rev", "Chronosphere＋CyberArk 併入後貢獻營收（FY2026）", "US$M", None, 930, None, "Verified", SRC["K26"] + "；Note 8", "只作對照，不進模型", "", N0)
inp("cons_rev27", "共識 FY2027 營收", "US$M", None, 14170, None, "Interested-party", SRC["SA"], "52 家；只作對照，不進模型", INC_SELL + "；只作對照，不進模型", N0)
inp("cons_eps27", "共識 FY2027 Non-GAAP EPS", "US$", None, 4.18, None, "Interested-party", SRC["SA"], "52 家；只作對照，不進模型", INC_SELL + "；只作對照，不進模型", N2)
inp("cons_tp", "賣方平均目標價（55 家）", "US$", 190, 398.63, 525, "Interested-party", SRC["SAS"] + "（2026-11-14，S&P Global）；低高值 MarketScreener 2026-10-23", "共識目標價的最低、最高只作對照", INC_SELL + "；只作對照，不進模型", N2)

# ====================================================================== 歷史（Hist 頁）
# 鍵: (項目, 單位, {年: 值}, 標記, 來源縮寫, 誘因)
HIST = [
 ("prod", "Product 營收", "US$M", {2022: 1363.1, 2023: 1578.4, 2024: 1603.3, 2025: 1801.9, 2026: 2280}, "Verified", "R25Q4", ""),
 ("ss", "Subscription and support 營收", "US$M", {2022: 4138.4, 2023: 5314.3, 2024: 6424.2, 2025: 7419.6, 2026: 9200}, "Verified", "R25Q4", ""),
 ("rev", "合併營收", "US$M", {2022: 5501.5, 2023: 6892.7, 2024: 8027.5, 2025: 9221.5, 2026: 11480}, "Verified", "R25Q4", ""),
 ("sub", "其中 Subscription", "US$M", {2024: 4188, 2025: 4974, 2026: 6239}, "Verified", "K26", ""),
 ("sup", "其中 Support", "US$M", {2024: 2236, 2025: 2445, 2026: 2961}, "Verified", "K26", ""),
 ("cogs", "營業成本", "US$M", {2024: 2059.2, 2025: 2451.6, 2026: 3403}, "Verified", "R26Q4", ""),
 ("opex", "營業費用", "US$M", {2024: 5284.4, 2025: 5527.0, 2026: 7382}, "Verified", "R26Q4", ""),
 ("ebit", "GAAP 營業利益", "US$M", {2022: -188.8, 2023: 387.3, 2024: 683.9, 2025: 1242.9, 2026: 695}, "Verified", "R26Q4", ""),
 ("oth", "其他收入（費用）淨額", "US$M", {2024: 312.7, 2025: 355.8, 2026: -159}, "Verified", "R26Q4", ""),
 ("pti", "稅前利益", "US$M", {2024: 988.3, 2025: 1595.7, 2026: 536}, "Verified", "R26Q4", ""),
 ("tax", "所得稅費用（負值＝利益）", "US$M", {2024: -1589.3, 2025: 461.8, 2026: 229}, "Verified", "R26Q4", ""),
 ("ni", "GAAP 淨利", "US$M", {2022: -267.0, 2023: 439.7, 2024: 2577.6, 2025: 1133.9, 2026: 307}, "Verified", "R26Q4", ""),
 ("eps", "GAAP 稀釋 EPS（2 拆 1 後）", "US$", {2024: 3.64, 2025: 1.60, 2026: 0.40}, "Verified", "R26Q4", ""),
 ("shdil", "GAAP 稀釋加權股數", "百萬股", {2024: 707.9, 2025: 709.3, 2026: 764}, "Verified", "R26Q4", ""),
 ("ngebit", "Non-GAAP 營業利益", "US$M", {2025: 2652, 2026: 3356}, "Interested-party", "R26Q4", INC_CO),
 ("ngni", "Non-GAAP 淨利", "US$M", {2024: 1948.1, 2025: 2344.5, 2026: 2931}, "Interested-party", "R26Q4", INC_CO),
 ("ngeps", "Non-GAAP 稀釋 EPS（2 拆 1 後）", "US$", {2024: 2.84, 2025: 3.34, 2026: 3.84}, "Interested-party", "R26Q4", INC_CO),
 ("sbc", "SBC-related charges（調節項）", "US$M", {2024: 1161.7, 2025: 1386.4, 2026: 1712}, "Interested-party", "R26Q4", INC_CO),
 ("amort", "收購無形資產攤銷", "US$M", {2025: 164, 2026: 638}, "Interested-party", "R26Q4", INC_CO),
 ("amq4", "收購無形資產攤銷 Q4", "US$M", {2026: 281}, "Interested-party", "R26Q4", INC_CO),
 ("acq", "併購相關費用（調節項）", "US$M", {2025: -110, 2026: 295}, "Interested-party", "R26Q4", INC_CO),
 ("lit", "訴訟相關費用（調節項）", "US$M", {2025: -31, 2026: 16}, "Interested-party", "R26Q4", INC_CO),
 ("da", "折舊與攤銷（現金流量表）", "US$M", {2025: 343, 2026: 855}, "Verified", "K26", ""),
 ("cx", "資本支出", "US$M", {2025: 247, 2026: 440}, "Verified", "R26Q4", ""),
 ("ocf", "營業活動現金流", "US$M", {2024: 3258, 2025: 3716, 2026: 4553}, "Verified", "K26", ""),
 ("fcf", "自由現金流（OCF − 資本支出）", "US$M", {2024: 3101, 2025: 3469, 2026: 4113}, "Verified", "R26Q4", ""),
 ("cash", "現金及約當現金（年底）", "US$M", {2025: 2268.6, 2026: 2514}, "Verified", "R26Q4", ""),
 ("sti", "短期投資（年底）", "US$M", {2025: 634.6, 2026: 557}, "Verified", "R26Q4", ""),
 ("lti", "長期投資（年底）", "US$M", {2025: 5555.6, 2026: 4835}, "Verified", "R26Q4", ""),
 ("conv", "長期可轉換票據（公允價值，年底）", "US$M", {2026: 1774}, "Verified", "R26Q4", ""),
 ("intang", "無形資產淨額（年底）", "US$M", {2025: 763, 2026: 7017}, "Verified", "R26Q4", ""),
 ("arr", "NGS ARR（年底）", "US$B", {2025: 5.6, 2026: 9.1}, "Interested-party", "R26Q4", "公司自揭 KPI，無第三方查核；Q3 起含 CyberArk、Chronosphere"),
]
HIST_KEYS = {h[0] for h in HIST}
# Hist 衍生
EXTRA_HIST = [("cashinv", "現金、短期與長期投資（年底）", "US$M", {2025: "=h_cash_25+h_sti_25+h_lti_25", 2026: "=h_cash_26+h_sti_26+h_lti_26"}, "Derived", "公式", "")]
QDATA = {  # 季度（FY2026 Q1–Q4）
 "prod": (434, 514, 594, 738), "ss": (2040, 2080, 2408, 2672), "rev": (2474, 2594, 3002, 3410),
 "ebit": (309, 397, -183, 172), "ngebit": (746, 785, 814, 1011), "arr": (5.9, 6.3, 8.1, 9.1), "arr0": 5.6, "rpo": (15.5, 16.0, 18.4, 21.2),
}
QSRC = ("PANW 各季財報新聞稿：Q1 https://www.paloaltonetworks.com/company/press/2025/palo-alto-networks-reports-fiscal-first-quarter-2026-financial-results；"
        "Q2 https://investors.paloaltonetworks.com/node/20246/pdf；Q3 https://www.paloaltonetworks.com/company/press/2026/palo-alto-networks-reports-fiscal-third-quarter-2026-financial-results；"
        "Q4 " + SRC["R26Q4"])
TKT = ["Luna", "Terra", "Sol", "Astra"]
def tkv(name):
    return TK["items"][name]["values"]

# ====================================================================== 工作表與寫入輔助
def fmt_cell(c, f):
    if f: c.number_format = f
def bname(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

def build(ov=None, sens=None, fp=None, rec_tp=None, build_id="BUILD"):
    """建一本活頁簿。ov＝覆寫 In 基準值（Sens 整本重算用）；sens＝敏感度結果（成品用）；fp／rec_tp＝記錄的指紋與基準加權目標價。"""
    ov = ov or {}
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    names = ["README", "In", "Q", "Hist", "TK", "PL", "PL_A", "PL_C", "Val", "RevDCF", "Recon", "Sens", "Sources"]
    S = {n: wb.create_sheet(n) for n in names}
    # ---------------- In
    ws = S["In"]
    ws["A1"] = "In — 輸入（藍字；基準欄 E 進主模型；D 敏感度低、F 敏感度高，只用於 Sens 頁，不進目標價）"
    ws["A2"] = "研究包：software_ai/paloalto/data/（research_paloalto.md、facts_paloalto.csv）；每列單一主要標記（Verified／Interested-party／Analogy／Assumed／Derived）；本版未做先驗（工作單 G 節：先驗跳過）。"
    for j, h in enumerate(["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "說明", "來源誘因（Interested-party 列；混合來源列寫其他欄位）"]):
        c = ws.cell(3, j + 1, h); c.font = BOLD
    r = 4; INROW = {}
    for it in IN:
        if it[0] == "sec":
            ws.cell(r, 1, it[1]).font = BOLD; r += 1; continue
        k, item, unit, lo, base, hi, tag, src, note, inc, f = it
        INROW[k] = r
        ws.cell(r, 1, k); ws.cell(r, 2, item); ws.cell(r, 3, unit)
        b = ov.get(k, base)
        for col, v in ((4, lo), (5, b), (6, hi)):
            if v is None: continue
            if isinstance(v, str) and v == "=H_PEREL": v = "=pe_rel_live"
            c = ws.cell(r, col, v); c.font = BLUE; fmt_cell(c, f)
        ws.cell(r, 7, tag); ws.cell(r, 8, src); ws.cell(r, 9, note); ws.cell(r, 10, inc)
        if tag == "Assumed" and k in ("w_dcf", "tp_wA", "tp_wB", "tp_wC"): ws.cell(r, 5).fill = YELLOW
        r += 1
    for k, rr in INROW.items(): bname(wb, k, f"In!$E${rr}")
    # 情境驅動對照
    r += 1; ws.cell(r, 1, "J. 情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    for j, h in enumerate(["鍵", "項目", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）"]): ws.cell(r, j + 1, h).font = BOLD
    ws.cell(r, 8, "說明").font = BOLD
    r += 1
    for k, item, unit, trip, why in [
        ("n32", "受保護代理身分數 FY2032", "個", ("n32_A", "n32_B", "n32_C"), "引用 In 頁 n32_A／n32_B／n32_C；分化依據：代理身分是否被另行計價（unit_economics P_n32 區間 0.20–3.00），無直接量測，Assumed"),
        ("m32", "模組加購 FY2032", "%", ("m32_A", "m32_B", "m32_C"), "引用 In 頁 m32_A／m32_B／m32_C；分化依據：攻擊面擴大是否帶動平台化加購（P_m32 區間 3–25%）"),
        ("e32", "AI 原生競爭價格侵蝕 FY2032", "%", ("e32_A", "e32_B", "e32_C"), "引用 In 頁 e32_A／e32_B／e32_C；分化依據：基礎安全功能是否被模型商與雲端大廠內建（P_e32 區間 0–15%）"),
        ("k32", "AI 資安分析次數 FY2032", "次/年", ("k32_A", "k32_B", "k32_C"), "引用 In 頁 k32_A／k32_B／k32_C；分化依據：代理活動量（影響推論成本，方向與營收增量同向）"),
        ("scn_w", "情境權重", "%", ("tp_wA", "tp_wB", "tp_wC"), "引用 In 頁 tp_wA／tp_wB／tp_wC；情境名稱依 AI 對資安單位經濟的方向，不暗示機率；單位數成長（gu27、gu32）與 Product 路徑三情境相同：無機制與證據支持分化，寫明不分化")]:
        ws.cell(r, 1, k); ws.cell(r, 2, item); ws.cell(r, 3, unit)
        for j, kk in enumerate(trip): ws.cell(r, 4 + j, f"=E{INROW[kk]}")
        ws.cell(r, 8, why); r += 1
    for col, w in zip("ABCDEFGHIJ", (14, 52, 10, 11, 12, 11, 16, 60, 70, 40)): ws.column_dimensions[col].width = w
    global INLAST
    INLAST = r
    # ---------------- Hist
    wh = S["Hist"]
    wh["A1"] = "Hist — 歷史財務（US$M，會計年度 8 月 1 日至翌年 7 月 31 日）；公式只引用本頁儲存格（定義名稱 h_<鍵>_<年>）"
    for j, h in enumerate(["鍵", "項目", "單位", "FY2022", "FY2023", "FY2024", "FY2025", "FY2026", "標記", "來源", "來源誘因"]): wh.cell(3, j + 1, h).font = BOLD
    r = 4
    for k, item, unit, vals, tag, src, inc in HIST + EXTRA_HIST:
        wh.cell(r, 1, k); wh.cell(r, 2, item); wh.cell(r, 3, unit)
        for y, v in vals.items():
            c = wh[f"{HC[y]}{r}"]; c.value = v; c.font = BLUE if not (isinstance(v, str) and v.startswith("=")) else Font(color="000000")
            c.number_format = "#,##0.0##"
            bname(wb, f"h_{k}_{str(y)[2:]}", f"Hist!${HC[y]}${r}")
        wh.cell(r, 9, tag); wh.cell(r, 10, SRC.get(src, src)); wh.cell(r, 11, inc); r += 1
    wh.cell(r, 1, "arravg"); wh.cell(r, 2, "FY2026 NGS ARR 季平均（各季期初期末平均的四季平均）"); wh.cell(r, 3, "US$B")
    wh[f"H{r}"] = "=AVERAGE(Q!C12:F12)"; bname(wb, "h_arravg_26", f"Hist!$H${r}")
    wh.cell(r, 9, "Derived"); wh.cell(r, 10, "Q 頁 NGS ARR 各季末值"); r += 1
    for col, w in zip("ABCDEFGHIJK", (10, 40, 8, 11, 11, 11, 11, 11, 16, 90, 40)): wh.column_dimensions[col].width = w
    # ---------------- Q（FY2026 四季）
    wq = S["Q"]
    wq["A1"] = "Q — FY2026 四季（Q1＝2025-08 至 2025-10；Q4＝2026-05 至 2026-07）；分部對帳（最近四季）"
    hdrs = ["項目", "單位", "Q1", "Q2", "Q3", "Q4", "四季合計", "FY2026 全年", "合計−全年", "標記", "來源", "來源誘因"]
    for j, h in enumerate(hdrs): wq.cell(3, j + 1, h).font = BOLD
    def qrow(r, label, unit, vals, fy, tag, src, inc="", fmtq="#,##0"):
        wq.cell(r, 1, label); wq.cell(r, 2, unit)
        for j, v in enumerate(vals): c = wq.cell(r, 3 + j, v); c.font = BLUE; c.number_format = fmtq
        wq.cell(r, 7, f"=SUM(C{r}:F{r})").number_format = fmtq
        if fy is not None: wq.cell(r, 8, fy).number_format = fmtq; wq.cell(r, 9, f"=G{r}-H{r}").number_format = "0.0"
        wq.cell(r, 10, tag); wq.cell(r, 11, src); wq.cell(r, 12, inc)
    qrow(4, "Product", "US$M", QDATA["prod"], "=h_prod_26", "Verified", QSRC)
    qrow(5, "Subscription and support", "US$M", QDATA["ss"], "=h_ss_26", "Verified", QSRC)
    qrow(6, "合併營收（公司揭露）", "US$M", QDATA["rev"], "=h_rev_26", "Verified", QSRC)
    wq.cell(7, 1, "檢查：Product＋Subscription and support－合併營收")
    for j in range(4): wq.cell(7, 3 + j, f"={CL(3+j)}4+{CL(3+j)}5-{CL(3+j)}6")
    wq.cell(7, 7, "=G4+G5-G6"); wq.cell(7, 10, "Derived"); wq.cell(7, 11, "同列相減")
    qrow(8, "GAAP 營業利益", "US$M", QDATA["ebit"], "=h_ebit_26", "Verified", QSRC)
    qrow(9, "Non-GAAP 營業利益", "US$M", QDATA["ngebit"], "=h_ngebit_26", "Interested-party", QSRC, INC_CO)
    wq.cell(11, 1, "NGS ARR（期末；Q0＝FY2025 年底）").font = BOLD
    wq.cell(10, 1, "RPO（期末）"); wq.cell(10, 2, "US$B")
    for j, v in enumerate(QDATA["rpo"]): c = wq.cell(10, 3 + j, v); c.font = BLUE
    wq.cell(10, 10, "Interested-party"); wq.cell(10, 11, QSRC); wq.cell(10, 12, INC_CO)
    wq.cell(12, 1, "NGS ARR（期末）"); wq.cell(12, 2, "US$B")
    for j, v in enumerate(QDATA["arr"]): c = wq.cell(12, 3 + j, v); c.font = BLUE
    wq.cell(12, 10, "Interested-party"); wq.cell(12, 11, QSRC); wq.cell(12, 12, "公司自揭 KPI，無第三方查核；Q3 起含 CyberArk（2026-02-11）、Chronosphere（2026-01-29）")
    wq.cell(13, 1, "NGS ARR 期初（Q0＝FY2025 年底）"); wq.cell(13, 2, "US$B"); wq.cell(13, 3, 5.6).font = BLUE
    wq.cell(13, 10, "Interested-party"); wq.cell(13, 11, SRC["R25Q4"]); wq.cell(13, 12, "公司自揭 KPI")
    wq.cell(14, 1, "NGS ARR 各季平均（期初期末平均）"); wq.cell(14, 2, "US$B")
    wq["C14"] = "=AVERAGE(C13,C12)"; wq["D14"] = "=AVERAGE(C12,D12)"; wq["E14"] = "=AVERAGE(D12,E12)"; wq["F14"] = "=AVERAGE(E12,F12)"
    wq.cell(14, 7, "=AVERAGE(C14:F14)"); wq.cell(14, 10, "Derived"); wq.cell(14, 11, "上列相加除以二")
    wq.cell(15, 1, "檢查：四季合計與全年差（Product、Subscription and support、合併營收）")
    wq.cell(15, 3, "=I4"); wq.cell(15, 4, "=I5"); wq.cell(15, 5, "=I6")
    wq.cell(15, 10, "Derived"); wq.cell(15, 11, "上方各列「合計−全年」")
    # 修正：Hist 的 arravg 指向 Q!C14:F14
    wh[f"H{r-1}"] = "=AVERAGE(Q!C14:F14)"
    wq.column_dimensions["A"].width = 52
    # ---------------- TK（Tokenomics 快照）
    wt = S["TK"]
    src = TK["source"]
    wt["A1"] = f"TK — Tokenomics 快照 {src['version']}（commit {src['commit'][:10]}；擷取 {src['extractedAt']}；檔案 {src['file']}；讀自 software_ai/unit_economics/data/tk_snapshot.json，不自行估 token 價格）"
    wt["A2"] = "推論成本只讀契約名稱 IF_PriceRef_<層級>（市場價）、IF_TaskSucc_<層級>、IF_TaskTok*；全成本 IF_FullCost_<層級>（GB300 基準）只作「模型商自己做這件事的成本下限」對照。"
    for j, h in enumerate(["項目", "契約名稱", "Luna（1）", "Terra（2）", "Sol（3）", "Astra（4）", "單位"]): wt.cell(3, j + 1, h).font = BOLD
    wt.cell(4, 1, "參考請求混合 $/M（市場價）"); wt.cell(4, 2, "IF_PriceRef_<層級>"); wt.cell(4, 7, "US$/百萬 token")
    wt.cell(5, 1, "成功率 p（任務 3）"); wt.cell(5, 2, "IF_TaskSucc_<層級>[3]"); wt.cell(5, 7, "比例")
    wt.cell(6, 1, "全成本 $/M（GB300 基準；對照下限）"); wt.cell(6, 2, "IF_FullCost_<層級>"); wt.cell(6, 7, "US$/百萬 token")
    for j, t in enumerate(TKT):
        wt.cell(4, 3 + j, tkv(f"IF_PriceRef_{t}")); wt.cell(5, 3 + j, tkv(f"IF_TaskSucc_{t}")[2]); wt.cell(6, 3 + j, tkv(f"IF_FullCost_{t}")["GB300 NVL72"]["基準"])
    wt.cell(7, 1, "任務 3 每次嘗試 token 數（新鮮輸入＋快取輸入＋decode）"); wt.cell(7, 2, "IF_TaskTokFresh／Cached／Dec [3]")
    wt.cell(7, 3, tkv("IF_TaskTokFresh")[2] + tkv("IF_TaskTokCached")[2] + tkv("IF_TaskTokDec")[2]); wt.cell(7, 7, "token")
    wt.cell(8, 1, "每成功任務推論成本，市場價（FY2026 參考工作負載，未乘 ovh）"); wt.cell(8, 7, "US$/次")
    wt.cell(9, 1, "每成功任務推論成本，全成本對照（模型商自做的下限）"); wt.cell(9, 7, "US$/次")
    for j in range(4):
        L = CL(3 + j)
        wt[f"{L}8"] = f"=$C$7/1000000*{L}4/{L}5"; wt[f"{L}9"] = f"=$C$7/1000000*{L}6/{L}5"
    wt.cell(10, 1, "市場價 ÷ 全成本");
    for j in range(4): wt[f"{CL(3+j)}10"] = f"={CL(3+j)}8/{CL(3+j)}9"
    wt.cell(11, 1, "取用層級的每成功任務推論成本（市場價）"); wt["C11"] = "=INDEX(C8:F8,P_tier)"; wt.cell(11, 7, "US$/次")
    wt.cell(12, 1, "取用層級的全成本對照"); wt["C12"] = "=INDEX(C9:F9,P_tier)"; wt.cell(12, 7, "US$/次")
    wt.column_dimensions["A"].width = 60; wt.column_dimensions["B"].width = 30
    bname(wb, "tk_cps", "TK!$C$11")
    # ---------------- 情境頁
    RW = {}
    for sh, sc in (("PL", "B"), ("PL_A", "A"), ("PL_C", "C")):
        RW[sh] = scenario(wb, S[sh], sc)
    # ---------------- Val / RevDCF / Recon / Sens / README / Sources
    VR = valuation(wb, S, RW)
    global RW_REV, RW_RC
    RW_REV = revdcf(wb, S["RevDCF"], RW["PL"], VR)
    RW_RC = recon(wb, S["Recon"], RW["PL"], VR)
    sensitivity(wb, S["Sens"], VR, sens, fp, rec_tp, build_id)
    readme(wb, S["README"], RW, VR, build_id, sens)
    sources(S["Sources"])
    wb._rw = RW; wb._vr = VR; wb._inrow = INROW
    return wb


# ====================================================================== 情境頁
def scenario(wb, ws, sc):
    """寫一個情境（sc＝A／B／C）的完整損益、現金流、DCF。回傳 {key: 列}。"""
    ws["A1"] = f"{ws.title} — 情境 {sc}（{ {'A':'樂觀','B':'基準','C':'保守'}[sc] }）：命題模組（原型 P）→ 營收線 → 損益表 → 現金流 → DCF（US$M，會計年度 8 月 1 日至翌年 7 月 31 日）"
    ws["A2"] = "FY2024A–FY2026A＝實績（Hist 頁）；FY2027E–FY2032E＝預測。分化只在 AI 相關驅動（代理身分、模組加購、價格侵蝕、分析次數）；單位數成長與 Product 路徑三情境相同。"
    spec = []   # (key, label, unit, hist[C,D,E], fut template, fmt, note, e_override)
    def R(key, label, unit="US$M", h=None, f=None, fmt="#,##0.0", note="", e=None, bold=False):
        spec.append(dict(key=key, label=label, unit=unit, h=h, f=f, fmt=fmt, note=note, e=e, bold=bold))
    def HD(t): spec.append(dict(key=None, label=t, hdr=True))
    HD("年度")
    R("yr", "會計年度（FY，年底 7 月 31 日）", "年", [2024, 2025, 2026], None, "0", e=None)
    spec[-1]["fut"] = "yr"
    R("tt", "距 FY2026 年數", "年", ["=[yr]-2026"] * 3, "=[yr]-2026", "0")
    HD("一、營收分部（官方營收線；單一應報導分部）")
    R("prod", "Product", "US$M", ["=h_prod_24", "=h_prod_25", "=h_prod_26"], "=[prod-1]*(1+gp27+(gp32-gp27)*([yr]-2027)/5)", "#,##0.0", "Product 以成長路徑驅動（Assumed）")
    R("ss", "Subscription and support", "US$M", ["=h_ss_24", "=h_ss_25", "=h_ss_26"], "=[ngrev]+[osup]", "#,##0.0", "預測＝NGS ARR 相關營收＋其他")
    R("sub", "Subscription", "US$M", ["=h_sub_24", "=h_sub_25", "=h_sub_26"], "=[ss]*$E[sub]/$E[ss]", "#,##0.0", "預測期依 FY2026 占比拆分，只作 10-K 子列展示")
    R("sup", "Support", "US$M", ["=h_sup_24", "=h_sup_25", "=h_sup_26"], "=[ss]-[sub]", "#,##0.0")
    R("ngrev", "NGS ARR 相關營收", "US$M", [None, None, "=[arravg]*1000*ng_ss"], "=[arravg]*1000*ng_ss", "#,##0.0", "命題模組驅動：平均 NGS ARR × 列入 S&S 的比例")
    R("osup", "其他 Subscription and support", "US$M", [None, None, "=[ss]-[ngrev]"], "=[osup-1]*(1+g_os)", "#,##0.0", "FY2026 為 S&S 減 NGS ARR 相關營收（Derived）")
    R("rev", "合併營收", "US$M", ["=h_rev_24", "=h_rev_25", "=h_rev_26"], "=[prod]+[ss]", "#,##0.0", bold=True)
    R("chk1", "檢查：Product＋Subscription and support－合併營收", "US$M", ["=[prod]+[ss]-[rev]"] * 3, "=[prod]+[ss]-[rev]", "0.0")
    R("chk2", "檢查：NGS ARR 相關營收＋其他 Subscription and support－Subscription and support", "US$M", [None, None, "=[ngrev]+[osup]-[ss]"], "=[ngrev]+[osup]-[ss]", "0.0")
    R("chk3", "檢查：Subscription＋Support－Subscription and support", "US$M", [None, None, "=[sub]+[sup]-[ss]"], "=[sub]+[sup]-[ss]", "0.0", "FY2024、FY2025 因 10-K 整數與新聞稿小數差 0.2／0.6，不列入檢查")
    HD("二、命題模組 P：按防護範圍計價（NGS ARR ＝ 受保護單位 × 每單位營收）")
    R("nn", "受保護代理身分數（每 1 個 AI 前單位）", "個", [None, None, "=P_n26"], "=P_n26+(n32_%s-P_n26)*[tt]/6" % sc, "0.000")
    R("mm", "模組加購 ÷ R0", "%", [None, None, "=P_m26"], "=P_m26+(m32_%s-P_m26)*[tt]/6" % sc, "0.0%")
    R("ee", "AI 原生競爭價格侵蝕", "%", [None, None, "=P_e26"], "=P_e26+(e32_%s-P_e26)*[tt]/6" % sc, "0.0%")
    R("kk", "AI 資安分析次數（每單位每年）", "次/年", [None, None, "=P_k26"], "=P_k26+(k32_%s-P_k26)*[tt]/6" % sc, "#,##0.0")
    R("R0", "AI 前單位營收 R0（＝1 × 模組數 × 每模組單價）", "US$", [None, None, "=P_mod*P_px"], "=P_mod*P_px", "#,##0.00")
    R("Rper", "每單位營收（R0 × (1 − 侵蝕 ＋ 加購 ＋ 代理身分數 × 代理單價比)）", "US$", [None, None, "=[R0]*(1-[ee]+[mm]+[nn]*P_rho)"], "=[R0]*(1-[ee]+[mm]+[nn]*P_rho)", "#,##0.00")
    R("gu", "受保護單位數年增", "%", None, "=gu27+(gu32-gu27)*([yr]-2027)/5", "0.0%")
    R("U", "受保護單位數（百萬；由 FY2026 NGS ARR 與 Assumed 每單位營收反推）", "百萬", [None, None, "=h_arr_26*1000/[Rper]"], "=[U-1]*(1+[gu])*(1-IF([yr]=2028,tail_dist,0))", "#,##0.0", "絕對值依賴 Assumed 的 R0，只用比例")
    R("ARR", "NGS ARR（期末）", "US$B", [None, "=h_arr_25", "=[U]*[Rper]/1000"], "=[U]*[Rper]/1000", "0.000", "FY2026 與公司揭露 9.1 的恆等式核對見檢查列", bold=True)
    R("arravg", "NGS ARR 年平均（認列營收的基礎）", "US$B", [None, None, "=h_arravg_26"], "=([ARR-1]+[ARR])/2", "0.000", "FY2026 為四季平均（Q 頁）；預測期為期初期末平均")
    R("chk4", "檢查：FY2026 NGS ARR（U × 每單位營收）－公司揭露", "US$B", [None, None, "=[ARR]-h_arr_26"], None, "0.000")
    R("cps", "每次成功分析推論成本（市場價；含工作負載係數與年路徑）", "US$/次", [None, None, "=tk_cps*ovh"], "=tk_cps*ovh*((1+g_px)*(1+g_tok))^[tt]", "0.0000", "只讀 Tokenomics 快照（TK 頁）")
    R("incr_u", "每單位 AI 營收增量（相對 FY2026）", "US$", [None, None, "=0"], "=[R0]*(([nn]-$E[nn])*P_rho+([mm]-$E[mm]))", "#,##0.0000")
    R("eros_u", "每單位定價侵蝕（相對 FY2026）", "US$", [None, None, "=0"], "=[R0]*([ee]-$E[ee])", "#,##0.0000")
    R("dinf_u", "每單位推論成本增量（相對 FY2026）", "US$", [None, None, "=0"], "=[kk]*[cps]-$E[kk]*$E[cps]", "#,##0.0000")
    R("incr", "AI 營收增量（代理身分＋模組加購；相對 FY2026）", "US$M", [None, None, "=0"], "=[U]*[incr_u]", "#,##0.0")
    R("eros", "定價侵蝕（相對 FY2026）", "US$M", [None, None, "=0"], "=[U]*[eros_u]", "#,##0.0")
    R("dinf", "推論成本增量（市場價；相對 FY2026）", "US$M", [None, None, "=0"], "=[U]*[dinf_u]", "#,##0.0")
    R("ainet", "AI 淨效果（增量 − 侵蝕 − 推論成本）", "US$M", [None, None, "=0"], "=[incr]-[eros]-[dinf]", "#,##0.0", bold=True)
    R("infsh", "推論成本增量 ÷ 營業成本", "%", None, None, "0.000%")
    R("infsh2", "推論成本增量（全成本對照，模型商自做的下限）÷ 營業成本", "%", None, None, "0.000%", "以任務成本比例縮放，不另估 token 價格")
    R("be", "損益兩平所需代理身分單價 ÷ AI 前單價（反解）", "x", None, '=IF([nn]-$E[nn]<=0,"",(([eros_u]+[dinf_u])/[R0]-([mm]-$E[mm]))/([nn]-$E[nn]))', "0.000", "淨效果＝0 時的 P_rho；負值＝不需代理身分收入即為正")
    HD("三、損益鏈（Non-GAAP → GAAP）")
    R("nmeff", "核心 Non-GAAP 營業利益率（基準全期持平）", "%", None, "=nm0+nm_up*[tt]", "0.0%")
    R("ngcost", "Non-GAAP 營業成本與費用（(1 − 核心率) × (營收＋定價侵蝕) ＋ 推論成本增量）", "US$M", None, "=(1-[nmeff])*([rev]+[eros])+[dinf]", "#,##0.0", "價格侵蝕不減成本")
    R("ngebit", "Non-GAAP 營業利益", "US$M", [None, "=h_ngebit_25", "=h_ngebit_26"], "=[rev]-[ngcost]", "#,##0.0", bold=True)
    R("ngm", "Non-GAAP 營業利益率", "%", [None, "=[ngebit]/[rev]", "=[ngebit]/[rev]"], "=[ngebit]/[rev]", "0.0%")
    R("sbcx", "股票薪酬（SBC-related charges）", "US$M", ["=h_sbc_24", "=h_sbc_25", "=h_sbc_26"], "=[rev]*MAX(0,sbc_pct-sbc_dn*[tt])", "#,##0.0")
    R("bal", "收購無形資產餘額（年底）", "US$M", [None, "=h_intang_25", "=h_intang_26"], "=[bal-1]-[amortx]", "#,##0.0")
    R("amortx", "收購無形資產攤銷", "US$M", [None, "=h_amort_25", "=h_amort_26"], "=MIN(amort_run,[bal-1])", "#,##0.0")
    R("acqx", "併購相關費用", "US$M", [None, "=h_acq_25", "=h_acq_26"], "=IF([yr]=2027,acq27,IF([yr]=2028,acq28,0))", "#,##0.0")
    R("litx", "訴訟相關費用", "US$M", [None, "=h_lit_25", "=h_lit_26"], "=0", "#,##0.0")
    R("ebitg", "GAAP EBIT（Non-GAAP 營業利益 − 股票薪酬 − 攤銷 − 併購費用 − 訴訟）", "US$M", ["=h_ebit_24", "=[ngebit]-[sbcx]-[amortx]-[acqx]-[litx]", "=[ngebit]-[sbcx]-[amortx]-[acqx]-[litx]"], "=[ngebit]-[sbcx]-[amortx]-[acqx]-[litx]", "#,##0.0")
    R("chk5", "檢查：GAAP EBIT 調節－公司揭露 GAAP 營業利益", "US$M", [None, "=[ebitg]-h_ebit_25", "=[ebitg]-h_ebit_26"], None, "0.0", "FY2025 差 0.3：調節表整數四捨五入")
    R("othx", "其他收入（淨額）", "US$M", ["=h_oth_24", "=h_oth_25", "=h_oth_26"], "=ysd*[cash-1]", "#,##0.0", "預測＝收益率 × 期初現金與投資；公允價值變動不預測")
    R("ptix", "稅前利益", "US$M", ["=h_pti_24", "=h_pti_25", "=h_pti_26"], "=[ebitg]+[othx]", "#,##0.0")
    R("taxx", "所得稅", "US$M", ["=h_tax_24", "=h_tax_25", "=h_tax_26"], "=tax*[ptix]", "#,##0.0")
    R("nix", "稅後淨利（GAAP）", "US$M", ["=h_ni_24", "=h_ni_25", "=h_ni_26"], "=[ptix]-[taxx]", "#,##0.0", bold=True)
    R("pxb", "買回價（現價以 ke 推算）", "US$", None, "=px*(1+ke)^[tt]", "#,##0.00")
    R("shend", "稀釋股數（期末水準）", "百萬股", [None, None, "=sh_q4"], "=[shend-1]*(1+dil)-bb/[pxb]", "#,##0.0")
    R("shavg", "稀釋加權股數", "百萬股", ["=h_shdil_24", "=h_shdil_25", "=h_shdil_26"], "=([shend-1]+[shend])/2", "#,##0.0")
    R("epsx", "GAAP 稀釋 EPS", "US$", ["=h_eps_24", "=h_eps_25", "=h_eps_26"], "=[nix]/[shavg]", "0.00")
    R("ngni", "Non-GAAP 淨利（(Non-GAAP 營業利益＋其他收入) × (1 − 稅率)）", "US$M", ["=h_ngni_24", "=h_ngni_25", "=h_ngni_26"], "=([ngebit]+[othx])*(1-tax)", "#,##0.0")
    R("ngeps", "Non-GAAP 稀釋 EPS", "US$", ["=h_ngeps_24", "=h_ngeps_25", "=h_ngeps_26"], "=[ngni]/[shavg]", "0.00", bold=True)
    HD("四、現金流（公司 FCF 定義：營業現金流 − 資本支出；SBC 不扣）")
    R("dax", "折舊與攤銷", "US$M", [None, "=h_da_25", "=h_da_26"], "=dep_pct*[rev]+[amortx]", "#,##0.0")
    R("cxx", "資本支出", "US$M", [None, "=h_cx_25", "=h_cx_26"], "=cx_pct*[rev]", "#,##0.0")
    R("nwc", "營運資金淨流入（含遞延收入）", "US$M", None, "=nwc_r*([rev]-[rev-1])", "#,##0.0")
    R("cfo", "營業活動現金流（淨利＋SBC＋折舊攤銷＋營運資金淨流入）", "US$M", ["=h_ocf_24", "=h_ocf_25", "=h_ocf_26"], "=[nix]+[sbcx]+[dax]+[nwc]", "#,##0.0")
    R("fcf", "自由現金流（扣股利前；本公司無股利）", "US$M", ["=h_fcf_24", "=h_fcf_25", "=h_fcf_26"], "=[cfo]-[cxx]", "#,##0.0", bold=True)
    R("fcfm", "自由現金流率", "%", ["=[fcf]/[rev]"] * 3, "=[fcf]/[rev]", "0.0%")
    R("cash", "現金與投資（年底；扣買回）", "US$M", [None, "=h_cashinv_25", "=h_cashinv_26"], "=[cash-1]+[fcf]-bb", "#,##0.0")
    HD("五、FCFF DCF（標準規格：NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；期中折現）")
    R("nopat", "NOPAT（GAAP EBIT × (1 − DCF 現金稅率)；股票薪酬不加回）", "US$M", None, "=[ebitg]*(1-tax_dcf)", "#,##0.0")
    R("fcff", "FCFF（無槓桿自由現金流）", "US$M", None, "=[nopat]+[dax]-[cxx]+[nwc]", "#,##0.0", bold=True)
    R("dfm", "期中折現因子", "x", None, "=1/(1+wacc)^([tt]-0.5)", "0.0000")
    R("pv", "FCFF 現值", "US$M", None, "=[fcff]*[dfm]", "#,##0.0")
    # 展開列
    rows = {}; r = 4
    for s in spec:
        if s.get("hdr"): r += 1; s["row"] = r; r += 1; continue
        s["row"] = r; rows[s["key"]] = r; r += 1
    scal_start = r + 1
    # 標量（DCF 彙總）放在欄 C
    sc_rows = {}
    scal = [
        ("ebitn", "終端年 EBIT（FY2032 Non-GAAP 營業利益 − 股票薪酬；排除收購無形資產攤銷與整合費用）", "US$M", "=K{ngebit}-K{sbcx}"),
        ("nopat33", "終端年 NOPAT（FY2033＝終端年 EBIT × (1＋g) × (1 − 稅率)）", "US$M", "=C{ebitn}*(1+g)*(1-tax_dcf)"),
        ("ronic", "RONIC（預設＝WACC）", "%", "=wacc+ronic_x"),
        ("tv", "永續價值（FY2032 年底）", "US$M", "=IF(tv_alt=1,K{fcff}*(1+g)/(wacc-g),C{nopat33}*(1-g/C{ronic})/(wacc-g))"),
        ("pvtv", "永續價值現值（期中慣例：折 5.5 年）", "US$M", "=C{tv}/(1+wacc)^5.5"),
        ("pvx", "明確期 FCFF 現值合計（FY2027–FY2032）", "US$M", "=SUM(F{pv}:K{pv})"),
        ("ev", "企業價值 EV", "US$M", "=C{pvx}+C{pvtv}"),
        ("ncash", "淨現金（現金與投資 − 可轉債公允價值＋capped call）", "US$M", "=cash_0-conv+capcall"),
        ("eqv", "股權價值", "US$M", "=C{ev}+C{ncash}"),
        ("shv", "評價日股數（流通股＋RSU＋PSU×計入比例）", "百萬股", "=sh_out+rsu+psu*psu_in"),
        ("vps", "每股價值（評價日 2026-07-31）", "US$", "=C{eqv}/C{shv}"),
        ("roll", "推到目標時點的倍數（ke，自評價日起）", "x", "=(1+ke)^((tp_date-val_date)/365)"),
        ("tpdcf", "DCF 目標價（目標時點；負值取 0）", "US$", "=MAX(0,C{vps}*C{roll})"),
        ("tvsh", "永續價值占 EV", "%", '=IF(C{ev}>0,C{pvtv}/C{ev},"N/M")'),
        ("tpcmp", "同業本益比目標價（同業中位前瞻本益比 × 調整係數 × FY2028 Non-GAAP EPS）", "US$", "=pe_med*rel_adj*G{ngeps}"),
        ("tps", "情境目標價（DCF 與同業本益比依權重）", "US$", "=w_dcf*C{tpdcf}+(1-w_dcf)*C{tpcmp}"),
        ("gate", "收支關卡：FY2027–FY2031 自由現金流（扣股利）合計", "US$M", "=SUM(F{fcf}:J{fcf})"),
    ]
    r = scal_start; ws.cell(r - 1, 1, "六、DCF 彙總與目標價").font = BOLD
    for k, lab, u, f in scal:
        sc_rows[k] = r; r += 1
    allrows = dict(rows); allrows.update({k: v for k, v in sc_rows.items()})
    pat = re.compile(r"\[([a-z0-9_A-Z]+?)(-1)?\]")
    LK = rows
    def conv(t, col, prevcol):
        def sub(m):
            k, prev = m.group(1), m.group(2)
            return f"{prevcol if prev else col}{LK[k]}"
        t = re.sub(r"\$E\[([A-Za-z0-9_]+)\]", lambda m: f"$E${LK[m.group(1)]}", t)
        return pat.sub(sub, t)
    ws["A3"] = "項目"; ws["B3"] = "單位"
    for y in YEARS: ws[f"{COLS[y]}3"] = f"FY{y}{'A' if y <= 2026 else 'E'}"
    ws["L3"] = "說明"
    for c in "ABCDEFGHIJKL": ws[f"{c}3"].font = BOLD
    for s in spec:
        if s.get("hdr"):
            ws.cell(s["row"], 1, s["label"]).font = BOLD; continue
        r = s["row"]; ws.cell(r, 1, s["label"]); ws.cell(r, 2, s["unit"])
        if s["bold"]: ws.cell(r, 1).font = BOLD
        for i, y in enumerate(YEARS):
            col = COLS[y]; prevcol = CL(2 + i)
            if s["key"] == "yr": v = y
            elif y <= 2026:
                h = s["h"]; v = h[i] if h else None
                if isinstance(v, str): v = conv(v, col, prevcol)
            else:
                v = s["f"]; v = conv(v, col, prevcol) if v else None
            if v is None: continue
            c = ws[f"{col}{r}"]; c.value = v; c.number_format = s["fmt"]
        if s["note"]: ws.cell(r, 12, s["note"])
    # 標量
    for k, lab, u, f in scal:
        r = sc_rows[k]; ws.cell(r, 1, lab); ws.cell(r, 2, u)
        ff = f.format(**{kk: vv for kk, vv in allrows.items()})
        c = ws.cell(r, 3, ff); c.number_format = "#,##0.00" if u in ("US$", "x") else ("0.0%" if u == "%" else "#,##0.0")
    ws.column_dimensions["A"].width = 78; ws.column_dimensions["B"].width = 9; ws.column_dimensions["L"].width = 70
    for y in YEARS: ws.column_dimensions[COLS[y]].width = 11
    # 標準損益表（GAAP；HTML 與規格 D 節用）
    r = max(sc_rows.values()) + 2
    ws.cell(r, 1, "標準損益表（GAAP；US$M，EPS 為 US$）——由上方各列引用").font = BOLD; r += 1
    std = {}
    LK = dict(rows)
    ws.cell(r, 1, "標準損益表年度"); std["hdr"] = r
    for y in YEARS: ws[f"{COLS[y]}{r}"] = f"FY{y}{'A' if y <= 2026 else 'E'}"
    r += 1
    def srow(key, label, f_hist, f_fut=None, fmt="#,##0.0", unit="US$M"):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit); std[key] = r
        for i, y in enumerate(YEARS):
            if y == 2024 or True:
                col = COLS[y]; prevcol = CL(2 + i)
                t = f_hist if (y <= 2026 or f_fut is None) else f_fut
                if t is None: continue
                ws[f"{col}{r}"] = conv(t, col, prevcol) if "[" in t else t
                ws[f"{col}{r}"].number_format = fmt
        LK[key] = r
        r += 1
    srow("rev", "營收", "=[rev]")
    srow("yoy", "營收年增率", "=[rev]/[rev-1]-1", fmt="0.0%", unit="%")
    ws[f"C{std['yoy']}"] = "=C%d/h_rev_23-1" % std["rev"]
    ws[f"C{std['yoy']}"].number_format = "0.0%"
    srow("cogs", "營業成本", None)
    srow("gp", "毛利", None)
    srow("gm", "毛利率", "=[gp]/[rev]", fmt="0.0%", unit="%")
    srow("opex", "營業費用", None)
    srow("oi", "營業利益", "=[ebitg]")
    srow("om", "營業利益率", "=[oi]/[rev]", fmt="0.0%", unit="%")
    srow("oth", "其他收入（淨額）", "=[othx]")
    srow("pti", "稅前利益", "=[ptix]")
    srow("tax", "所得稅", "=[taxx]")
    srow("ni", "稅後淨利", "=[nix]")
    srow("nm", "淨利率", "=[ni]/[rev]", fmt="0.0%", unit="%")
    srow("eps", "EPS（完全稀釋，GAAP）", "=[epsx]", fmt="0.00", unit="US$")
    ws.cell(r, 1, "比率區").font = BOLD; std["ratio"] = r; r += 1
    ws.cell(r, 1, "營收年增率（比率區）"); ws.cell(r, 2, "%"); std["r_yoy"] = r
    ws.cell(r + 1, 1, "毛利率（比率區）"); ws.cell(r + 1, 2, "%"); std["r_gm"] = r + 1
    ws.cell(r + 2, 1, "營業利益率（比率區）"); ws.cell(r + 2, 2, "%"); std["r_om"] = r + 2
    ws.cell(r + 3, 1, "淨利率（比率區）"); ws.cell(r + 3, 2, "%"); std["r_nm"] = r + 3
    ws.cell(r + 4, 1, "淨利年增率"); ws.cell(r + 4, 2, "%"); std["r_niy"] = r + 4
    for i, y in enumerate(YEARS):
        col = COLS[y]; prevcol = CL(2 + i)
        ws[f"{col}{r}"] = f"={col}{std['yoy']}"; ws[f"{col}{r+1}"] = f"={col}{std['gm']}"; ws[f"{col}{r+2}"] = f"={col}{std['om']}"; ws[f"{col}{r+3}"] = f"={col}{std['nm']}"
        if i > 0: ws[f"{col}{r+4}"] = f'=IF({prevcol}{std["ni"]}>0,{col}{std["ni"]}/{prevcol}{std["ni"]}-1,"N/M")'
        for k in range(5): ws[f"{col}{r+k}"].number_format = "0.0%"
    r += 6
    ws.cell(r, 1, "標準損益表說明：GAAP 口徑（財報法定口徑，便於對 10-K）；公司與市場主要使用 non-GAAP（上方「三、損益鏈」與 Non-GAAP EPS 列）。預測期營業成本＝FY2026 營業成本占總成本比 ×（營收 − GAAP EBIT），營業費用為其餘（Assumed 切分，不影響營業利益）。").alignment = Alignment(wrap_text=False)
    std["note"] = r
    # 營業成本／費用預測切分（覆寫 cogs、opex）
    for i, y in enumerate(YEARS):
        col = COLS[y]
        if y <= 2026:
            ws[f"{col}{std['cogs']}"] = f"=h_cogs_{str(y)[2:]}"; ws[f"{col}{std['opex']}"] = f"=h_opex_{str(y)[2:]}"
        else:
            ws[f"{col}{std['cogs']}"] = f"=cogs_sh*({col}{rows['rev']}-{col}{rows['ebitg']})"; ws[f"{col}{std['opex']}"] = f"=({col}{rows['rev']}-{col}{rows['ebitg']})-{col}{std['cogs']}"
        ws[f"{col}{std['gp']}"] = f"={col}{std['rev']}-{col}{std['cogs']}"
        for k in ("cogs", "opex", "gp"): ws[f"{col}{std[k]}"].number_format = "#,##0.0"
    ws.cell(r + 1, 1, "檢查：標準損益表 營業利益 − (毛利 − 營業費用)")
    for i, y in enumerate(YEARS):
        col = COLS[y]; ws[f"{col}{r+1}"] = f"={col}{std['oi']}-({col}{std['gp']}-{col}{std['opex']})"; ws[f"{col}{r+1}"].number_format = "0.0"
    std["chk"] = r + 1
    # cogs 於損益鏈引用（infsh）：[cogs] 指標準損益表營業成本列
    for y in YEARS:
        if y >= 2027:
            col = COLS[y]
            ws[f"{col}{rows['infsh']}"] = f"={col}{rows['dinf']}/{col}{std['cogs']}"
            ws[f"{col}{rows['infsh2']}"] = f"={col}{rows['dinf']}*TK!$C$12/TK!$C$11/{col}{std['cogs']}"
    # 檢查：FY2026 公式一致、恆等式
    ws.cell(std["chk"] + 1, 1, "檢查：營收 ＝ 各項乘積（FY2026 受保護單位 × 每單位營收 ÷ 1000 − 公司揭露 NGS ARR）")
    ws[f"E{std['chk']+1}"] = f"=E{rows['chk4']}"
    allrows.update({"std_" + k: v for k, v in std.items()})
    rows.update({k: v for k, v in sc_rows.items()}); rows.update({"std_" + k: v for k, v in std.items()})
    return rows


# ====================================================================== 估值頁
def tier_f(u):
    return f'IF({u}>0.5,"強力買進",IF({u}>0.25,"買進",IF({u}>=-0.25,"中立",IF({u}>=-0.5,"放空","強力放空"))))'

def valuation(wb, S, RW):
    ws = S["Val"]; R = RW["PL"]
    ws["A1"] = "Val — 估值：DCF（FCFF）與同業本益比（Comparable）加權；三情境各算一次再依權重加權"
    rows = {}
    def put(key, label, b=None, c=None, d=None, fmt="#,##0.00", note="", bold=False):
        r = put.r; rows[key] = r; put.r += 1
        ws.cell(r, 1, label).font = BOLD if bold else Font()
        for col, v in zip((2, 3, 4), (b, c, d)):
            if v is not None:
                cc = ws.cell(r, col, v); cc.number_format = fmt
        if note: ws.cell(r, 5, note)
        return r
    put.r = 3
    ws.cell(2, 1, "項目").font = BOLD; ws.cell(2, 2, "值／情境 A").font = BOLD; ws.cell(2, 3, "情境 B").font = BOLD; ws.cell(2, 4, "情境 C").font = BOLD
    # 先放折現率
    put("sec1", "折現率（WACC，CAPM）", bold=True)
    put("rf", "無風險利率", "=rf", fmt="0.00%")
    put("beta", "β（Blume 調整後中位數）", "=beta", fmt="0.000")
    put("erp", "股權風險溢酬", "=erp", fmt="0.00%")
    put("ke", "股權成本 ke＝rf＋β×ERP", "=B{rf}+B{beta}*B{erp}", fmt="0.00%")
    put("kdat", "稅後債務成本", "=kd*(1-tax)", fmt="0.00%")
    put("ew", "股權市值（現價 × 評價日股數）", "=px*(sh_out+rsu)", fmt="#,##0")
    put("dw", "債務（可轉債公允價值）", "=conv", fmt="#,##0")
    put("wacc", "WACC（含 Sens 加減）", "=B{ew}/(B{ew}+B{dw})*B{ke}+B{dw}/(B{ew}+B{dw})*B{kdat}+wacc_adj", fmt="0.00%", bold=True)
    put("gg", "永續成長率 g", "=g", fmt="0.00%")
    put("ronic", "RONIC（預設＝WACC）", "=B{wacc}+ronic_x", fmt="0.00%")
    # 修正公式佔位
    for k in ("ke", "ew", "dw", "wacc"):
        pass
    def fx(s): return s
    # 目標價摘要
    put("sec2", "目標價與評等（目標時點＝現價日起 12 個月）", bold=True)
    put("px", "現價（2026-11-13 收盤）", "=px", fmt="#,##0.00")
    put("tp", "加權目標價（DCF 與同業本益比加權）", "=w_dcf*B{tpdcf}+(1-w_dcf)*B{tpcmp}", bold=True)
    put("tpdcf", "DCF 目標價（三情境加權）", "=tp_wA*PL_A!C%d+tp_wB*PL!C%d+tp_wC*PL_C!C%d" % ((R["tpdcf"],) * 3))
    put("tpcmp", "Comparable 目標價（同業前瞻本益比中位數法，三情境加權）", "=tp_wA*PL_A!C%d+tp_wB*PL!C%d+tp_wC*PL_C!C%d" % ((R["tpcmp"],) * 3))
    put("cons", "市場共識目標價（賣方 55 家平均，2026-11-14）", "=cons_tp", note="顯示用，只作對照，不進模型")
    put("up", "加權目標價上檔空間", "=B{tp}/B{px}-1", fmt="+0.0%;-0.0%")
    put("upd", "DCF 目標價上檔空間", "=B{tpdcf}/B{px}-1", fmt="+0.0%;-0.0%")
    put("upc", "Comparable 目標價上檔空間", "=B{tpcmp}/B{px}-1", fmt="+0.0%;-0.0%")
    put("tier", "級距（> +50%、> +25%、≥ −25%、≥ −50%）", '=' + tier_f("B{up}").replace("B{up}", "B{up}"))
    put("vrate", "估值評等（只看上檔空間）", '=' + tier_f("B{up}"), bold=True)
    put("gate", "入不敷出判定用（收支關卡：FY2027–FY2031 自由現金流扣股利合計，情境 B）", "=PL!C%d" % R["gate"], fmt="#,##0")
    put("gatev", "收支關卡 A／B／C（FY2027–FY2031 自由現金流扣股利合計，US$M）", "=PL_A!C%d" % R["gate"], "=PL!C%d" % R["gate"], "=PL_C!C%d" % R["gate"], fmt="#,##0")
    put("gtxt", "入不敷出判定", '=IF(B{gate}<0,"不通過","通過")')
    put("rec", "投資建議", '=IF(AND(B{gate}<0,OR(B{vrate}="買進",B{vrate}="強力買進")),"收支不足",B{vrate})', bold=True)
    put("scn", "情境目標價 A／B／C", "=PL_A!C%d" % R["tps"], "=PL!C%d" % R["tps"], "=PL_C!C%d" % R["tps"])
    put("scndcf", "情境 DCF 目標價 A／B／C", "=PL_A!C%d" % R["tpdcf"], "=PL!C%d" % R["tpdcf"], "=PL_C!C%d" % R["tpdcf"])
    put("scncmp", "情境 Comparable 目標價 A／B／C", "=PL_A!C%d" % R["tpcmp"], "=PL!C%d" % R["tpcmp"], "=PL_C!C%d" % R["tpcmp"])
    put("scnw", "情境權重 A／B／C", "=tp_wA", "=tp_wB", "=tp_wC", fmt="0%")
    put("sec3", "DCF 彙總（情境 A／B／C；US$M）", bold=True)
    for k, lab, u in [("pvx", "明確期 FCFF 現值", "#,##0"), ("pvtv", "永續價值現值", "#,##0"), ("ev", "企業價值 EV", "#,##0"), ("ncash", "淨現金", "#,##0"), ("eqv", "股權價值", "#,##0"),
                      ("shv", "評價日股數（百萬）", "#,##0.0"), ("vps", "每股價值（評價日）", "#,##0.00"), ("tpdcf", "DCF 目標價", "#,##0.00"), ("tvsh", "永續價值占 EV", "0.0%")]:
        put("d_" + k, lab, "=PL_A!C%d" % R[k], "=PL!C%d" % R[k], "=PL_C!C%d" % R[k], fmt=u)
    put("sec4", "同業本益比（Comparable）", bold=True)
    put("pef", "同業前瞻本益比：Fortinet", "=pe_f", fmt="0.00")
    put("pez", "同業前瞻本益比：Zscaler", "=pe_z", fmt="0.00")
    put("pec", "同業前瞻本益比：Check Point", "=pe_c", fmt="0.00")
    put("pen", "同業前瞻本益比：Cloudflare", "=pe_n", fmt="0.00")
    put("pemed", "同業中位數（4 家；不做迴歸）", "=MEDIAN(B{pef}:B{pen})", fmt="0.00", bold=True)
    put("adj", "相對倍數調整係數", "=rel_adj", fmt="0.00", note="1.00（未調整：缺同業 NTM 成長率、利益率與同日 EPS）")
    put("eps28", "FY2028 Non-GAAP EPS（A／B／C）", "=PL_A!G%d" % R["ngeps"], "=PL!G%d" % R["ngeps"], "=PL_C!G%d" % R["ngeps"])
    put("cmp", "Comparable 目標價（A／B／C）", "=PL_A!C%d" % R["tpcmp"], "=PL!C%d" % R["tpcmp"], "=PL_C!C%d" % R["tpcmp"])
    put("pxpe", "PANW 資料商前瞻本益比 ÷ 同業中位數（展示；同一資料商同日口徑）", "=pe_p/B{pemed}", fmt="0.00", note="分子為 StockAnalysis 2026-11-13 前瞻本益比（Interested-party，資料商口徑）；高案相對倍數係數取此值")
    put("pemod", "PANW 現價 ÷ 模型 FY2027 Non-GAAP EPS（情境 B；NTM 近似）", "=px/PL!F%d" % R["ngeps"], fmt="0.00")
    put("sec5", "期間與評等邊界", bold=True)
    put("per_d", "折現期間（評價日 → 目標時點，年）", "=(tp_date-val_date)/365", fmt="0.00")
    put("per_h", "持有期間（現價日 → 目標時點，年）", "=(tp_date-px_date)/365", fmt="0.00")
    put("bdist", "加權上檔空間距最近評等邊界（百分點）", "=MIN(ABS(B{up}-0.5),ABS(B{up}-0.25),ABS(B{up}+0.25),ABS(B{up}+0.5))*100", fmt="0.0")
    put("bflag", "距邊界揭露（不到 5 個百分點才顯示）", '=IF(B{bdist}<5,"加權上檔空間距最近評等邊界只差 "&TEXT(B{bdist},"0.0")&" 個百分點","距最近評等邊界 "&TEXT(B{bdist},"0.0")&" 個百分點（未貼近）")')
    put("attr_d", "上檔空間歸因：DCF 貢獻（w × (DCF − 現價) ÷ 現價）", "=w_dcf*(B{tpdcf}-B{px})/B{px}", fmt="+0.0%;-0.0%")
    put("attr_c", "上檔空間歸因：Comparable 貢獻（(1 − w) × (Comparable − 現價) ÷ 現價）", "=(1-w_dcf)*(B{tpcmp}-B{px})/B{px}", fmt="+0.0%;-0.0%")
    # 解析 {key}
    for r in range(3, put.r):
        for col in (2, 3, 4):
            c = ws.cell(r, col)
            if isinstance(c.value, str) and "{" in c.value:
                c.value = re.sub(r"\{(\w+)\}", lambda m: str(rows[m.group(1)]), c.value)
        a = ws.cell(r, 1)
    for k in ("ke", "wacc"): bname(wb, k, f"Val!$B${rows[k]}")
    bname(wb, "pe_med", f"Val!$B${rows['pemed']}")
    bname(wb, "pe_rel_live", f"Val!$B${rows['pxpe']}")
    ws.column_dimensions["A"].width = 78; ws.column_dimensions["E"].width = 70
    for c in "BCD": ws.column_dimensions[c].width = 14
    return rows


# ====================================================================== RevDCF
def revdcf(wb, ws, R, V):
    ws["A1"] = "RevDCF — 反向 DCF：現價隱含什麼（情境 B；DCF 單法；封閉解，因 EV 對核心 Non-GAAP 營業利益率為線性）"
    ws["A2"] = "只解一個變數：核心 Non-GAAP 營業利益率（全期持平，含終值）；其餘輸入不動。反解值是「DCF 目標價＝現價」的點，不是評等門檻（第 4 條）；加權目標價另含同業本益比，不在此解。"
    rows = {}
    def put(key, label, f, fmt="#,##0.00", note=""):
        r = put.r; rows[key] = r; put.r += 1
        ws.cell(r, 1, label); c = ws.cell(r, 2, f); c.number_format = fmt
        if note: ws.cell(r, 3, note)
    put.r = 4
    P = lambda k: f"PL!C{R[k]}"
    put("px", "現價", "=px")
    put("vv", "現價折回評價日的每股價值（以 ke 折現）", "=px/(1+ke)^((px_date-val_date)/365)")
    put("shv", "評價日股數（百萬）", "=" + P("shv"), "#,##0.0")
    put("evt", "現價隱含企業價值 EV（US$M）", "=B{vv}*B{shv}-" + P("ncash"), "#,##0")
    put("evb", "模型 EV（情境 B）", "=" + P("ev"), "#,##0")
    put("d1", "每 1 個百分點核心率對 EV 的影響（明確期）", "=0.01*(1-tax_dcf)*SUMPRODUCT((PL!F%d:K%d+PL!F%d:K%d),PL!F%d:K%d)" % (R["rev"], R["rev"], R["eros"], R["eros"], R["dfm"], R["dfm"]), "#,##0")
    put("d2", "每 1 個百分點核心率對 EV 的影響（終值）", "=IF(tv_alt=1,0,0.01*(PL!K%d+PL!K%d)*(1+g)*(1-tax_dcf)*(1-g/%s)/(wacc-g)/(1+wacc)^5.5)" % (R["rev"], R["eros"], P("ronic")), "#,##0")
    put("m0", "模型核心 Non-GAAP 營業利益率", "=nm0", "0.0%")
    put("m1", "現價隱含核心 Non-GAAP 營業利益率（全期持平）", "=B{m0}+(B{evt}-B{evb})/(B{d1}+B{d2})*0.01", "0.0%")
    put("m1x", "隱含 Non-GAAP 營業利益率 − 模型", "=B{m1}-B{m0}", "+0.0%;-0.0%")
    put("evs", "現價隱含 EV ÷ FY2027E 營收（情境 B）", "=B{evt}/PL!F%d" % R["rev"], "0.0")
    put("evs2", "模型 EV ÷ FY2027E 營收", "=B{evb}/PL!F%d" % R["rev"], "0.0")
    put("tvsh", "模型永續價值占 EV", "=" + P("tvsh"), "0.0%")
    put("sbc", "對照：FY2032 SBC 占營收（模型）", "=PL!K%d/PL!K%d" % (R["sbcx"], R["rev"]), "0.0%")
    put("note", "解讀", '="現價隱含核心 Non-GAAP 營業利益率 "&TEXT(B{m1},"0.0%")&"，"&IF(B{m1}>B{m0},"高於","低於")&"模型的 "&TEXT(B{m0},"0.0%")&"（差 "&TEXT(ABS(B{m1x})*100,"0.0")&" 個百分點）；全期持平且 SBC 占營收不變下，現價要求的利益率已接近或超過營收減去全部成本的上限。"', "@")
    for r in range(4, put.r):
        c = ws.cell(r, 2)
        if isinstance(c.value, str) and "{" in c.value: c.value = re.sub(r"\{(\w+)\}", lambda m: str(rows[m.group(1)]), c.value)
    ws.column_dimensions["A"].width = 70; ws.column_dimensions["B"].width = 16
    return rows


# ====================================================================== Recon
def recon(wb, ws, R, V):
    ws["A1"] = "Recon — 對帳：模型（情境 B）與公司指引、市場共識、實績的口徑對照；方向字樣由公式依數值組字（第 31、57 條）"
    r = 3; RC = {}
    def head(t, cols):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
        for j, h in enumerate(cols): ws.cell(r, 1 + j, h).font = BOLD
        r += 1
    def row(key, vals, fmts=None):
        nonlocal r
        RC[key] = r
        for j, v in enumerate(vals):
            if v is None: continue
            c = ws.cell(r, 1 + j, v)
            if fmts and j < len(fmts) and fmts[j]: c.number_format = fmts[j]
        r += 1
    def rng(m, lo, hi, unit_div=1, dec="0.0"):
        return (f'=IF({m}>{hi},"高於指引上緣 "&TEXT(({m}/{hi}-1)*100,"{dec}")&" 個百分點",IF({m}<{lo},"低於指引下緣 "&TEXT((1-{m}/{lo})*100,"{dec}")&" 個百分點","落在指引區間內"))')
    P = lambda k, col="F": f"PL!{col}{R[k]}"
    head("1. 合併層：FY2027 模型（情境 B）對公司指引與共識（指引只作對照，不進基準；X3）", ["項目", "模型（情境 B）", "指引低", "指引高", "共識", "判定（模型對指引）", "說明"])
    row("rev", ["FY2027 營收（US$M）", "=" + P("rev"), "=g_rev_lo", "=g_rev_hi", "=cons_rev27", rng("B{r}", "C{r}", "D{r}"), "指引 2026-09-01；共識 52 家 2026-10-15（財報後）"], [None, N0, N0, N0, N0])
    row("arr", ["FY2027 年底 NGS ARR（US$B）", "=" + P("ARR"), "=g_arr_lo", "=g_arr_hi", None, rng("B{r}", "C{r}", "D{r}"), "指引含 CyberArk 全年度；模型單位成長不以指引校準"], [None, "0.000", "0.000", "0.000"])
    row("nm", ["FY2027 Non-GAAP 營業利益率", "=" + P("ngm"), "=g_nm", "=g_nm", None, '=IF(B{r}>D{r},"高於指引 "&TEXT((B{r}-D{r})*100,"0.0")&" 個百分點",IF(B{r}<C{r},"低於指引 "&TEXT((C{r}-B{r})*100,"0.0")&" 個百分點","與指引相同"))', "指引為單點 29.5%；模型基準為 FY2026 實績持平，扣 AI 增量成本"], [None, "0.0%", "0.0%", "0.0%"])
    row("eps", ["FY2027 Non-GAAP EPS（US$）", "=" + P("ngeps"), "=g_eps_lo", "=g_eps_hi", "=cons_eps27", rng("B{r}", "C{r}", "D{r}", dec="0.0"), "共識貼近指引；模型股數另依淨發行與買回推算"], [None, "0.00", "0.00", "0.00", "0.00"])
    row("sh", ["FY2027 稀釋加權股數（百萬）", "=" + P("shavg"), "=g_sh_lo", "=g_sh_hi", None, rng("B{r}", "C{r}", "D{r}"), "指引 844–847"], [None, "#,##0.0", "#,##0", "#,##0"])
    row("fcf", ["FY2027 自由現金流率（模型：營業現金流 − 資本支出；指引為調整後口徑）", "=" + P("fcfm"), "=g_fcfm", "=g_fcfm", None, '=IF(B{r}>D{r},"高於",IF(B{r}<C{r},"低於","等於"))&"指引 "&TEXT(ABS(B{r}-D{r})*100,"0.0")&" 個百分點（口徑不同：指引加回總部購地與併購付款）"', "第 29、33 條：口徑不同，只作參照"], [None, "0.0%", "0.0%", "0.0%"])
    row("arr30", ["FY2030 年底 NGS ARR（US$B）對管理層目標", "=" + P("ARR", "I"), "=t_arr30", "=t_arr30", None, '=IF(B{r}<C{r},"低於目標 "&TEXT((1-B{r}/C{r})*100,"0.0")&"%","高於目標 "&TEXT((B{r}/C{r}-1)*100,"0.0")&"%")', "管理層目標 Interested-party；不校準"], [None, "0.00", "0.00", "0.00"])
    row("req", ["達到管理層目標所需的額外年增（FY2027–FY2030 四年；其餘不動，反解）", '=(C{ra}/B{ra})^(1/4)-1', None, None, None, '="需要 NGS ARR 每年比模型快 "&TEXT(B{r}*100,"0.0")&" 個百分點，才在 FY2030 達 "&TEXT(C{ra},"0")&" 十億美元"', "Derived"], [None, "0.0%"])
    ws.cell(RC["req"], 2).value = ws.cell(RC["req"], 2).value.replace("{ra}", str(RC["arr30"]))
    ws.cell(RC["req"], 6).value = ws.cell(RC["req"], 6).value.replace("{ra}", str(RC["arr30"]))
    row("fcf28", ["FY2028 自由現金流率對管理層目標（調整後口徑）", "=" + P("fcfm", "G"), "=t_fcf28", "=t_fcf28", None, '=IF(B{r}<C{r},"低於目標 "&TEXT((C{r}-B{r})*100,"0.0")&" 個百分點","高於目標 "&TEXT((B{r}-C{r})*100,"0.0")&" 個百分點")', "口徑不同，只作參照"], [None, "0.0%", "0.0%", "0.0%"])
    r += 1
    head("2. 逐分部（單一應報導分部；官方營收線 Product、Subscription and support）", ["項目", "FY2026 模型欄", "公司揭露", "差額", None, "指引／共識", "說明"])
    row("p", ["Product", "=PL!E%d" % R["prod"], "=h_prod_26", "=B{r}-C{r}", None, "未找到（公司不給分部指引）", "歷史＝公司揭露"], [None, N0, N0, "0.0"])
    row("s", ["Subscription and support", "=PL!E%d" % R["ss"], "=h_ss_26", "=B{r}-C{r}", None, "未找到", ""], [None, N0, N0, "0.0"])
    row("t", ["合併營收", "=PL!E%d" % R["rev"], "=h_rev_26", "=B{r}-C{r}", None, "指引見第 1 節", ""], [None, N0, N0, "0.0"])
    row("q", ["最近四季合計 − FY2026 全年（Product／Subscription and support／合併營收）", "=Q!I4", "=Q!I5", "=Q!I6", None, "Q 頁", "四個值皆應為 0"], [None, "0.0", "0.0", "0.0"])
    r += 1
    head("3. 一次性項目與有機成長（FY2026；Derived）", ["項目", "值", None, None, None, "說明", "出處"])
    row("org", ["有機營收（扣除 Chronosphere＋CyberArk 併入後貢獻）US$M", "=h_rev_26-acq_rev", None, None, None, "", "10-K Note 8、8-K"], [None, N0])
    row("orgg", ["有機營收年增率", "=B{o}/h_rev_25-1".replace("{o}", str(RC["org"])), None, None, None, '="合併年增 "&TEXT(h_rev_26/h_rev_25-1,"0.0%")&"，有機年增 "&TEXT(B{r},"0.0%")&"，差 "&TEXT((h_rev_26/h_rev_25-1-B{r})*100,"0.0")&" 個百分點來自併購"', "Derived"], [None, "0.0%"])
    row("bridge", ["GAAP 營業利益調節（Non-GAAP 營業利益 − 股票薪酬 − 攤銷 − 併購費用 − 訴訟）－公司揭露", "=PL!E%d" % R["chk5"], None, None, None, "應為 0", "Non-GAAP 調節表"], [None, "0.0"])
    row("sbcm", ["FY2026 GAAP 營業利益率 vs Non-GAAP 營業利益率", "=h_ebit_26/h_rev_26", "=h_ngebit_26/h_rev_26", None, None, '="差 "&TEXT((C{r}-B{r})*100,"0.0")&" 個百分點，主要為股票薪酬 "&TEXT(h_sbc_26/h_rev_26,"0.0%")&" 與攤銷 "&TEXT(h_amort_26/h_rev_26,"0.0%")&"（占營收）"', ""], [None, "0.0%", "0.0%"])
    r += 1
    head("4. 同業可比性（同一估值日 2026-11-13；皆 US GAAP；每組不足 8 家不做迴歸，X8）", ["同業", "前瞻本益比（x）", "EV／營收 TTM（x）", "成長率", "利益率", "會計口徑", "說明"])
    for tk, pe, evs in (("Fortinet（FTNT）", "pe_f", 11.32), ("Zscaler（ZS）", "pe_z", 8.88), ("Check Point（CHKP）", "pe_c", 4.67), ("Cloudflare（NET）", "pe_n", 39.95)):
        row("peer_" + tk[:3], [tk, "=" + pe, evs, "未取得", "未取得", "US GAAP；租賃為營業租賃", "StockAnalysis.com 統計頁 2026-11-13；成長率、利益率未取得同口徑，故不調整"], [None, "0.00", "0.00"])
    row("pemed", ["同業中位數", "=pe_med", "=MEDIAN(C%d:C%d)" % (r - 4, r - 1), None, None, "", "調整係數 1.00（未調整：缺成長與利益率）"], [None, "0.00", "0.00"])
    row("panw", ["PANW（資料商同日）", "=pe_p", 22.88, None, None, "US GAAP", '="PANW 前瞻本益比為同業中位數的 "&TEXT(B{r}/B{m},"0.00")&" 倍；同業中位落在口徑相近的 FTNT 與 ZS 之間"'.replace("{m}", str(RC["pemed"]))], [None, "0.00", "0.00"])
    row("pemod", ["PANW 現價 ÷ 模型 FY2027 Non-GAAP EPS（情境 B）", "=Val!B%d" % V["pemod"], None, None, None, "分母：模型情境 B、FY2027（8 月至 7 月）", "倍數分母標示（第 43 條）"], [None, "0.00"])
    r += 1
    head("5. Tokenomics 讀值（只讀快照，不自行估價）", ["項目", "值", None, None, None, "說明", "出處"])
    row("tkv", ["快照版本", TK["source"]["version"], None, None, None, "commit " + TK["source"]["commit"][:10], "software_ai/unit_economics/data/tk_snapshot.json"])
    row("tk1", ["取用層級每成功任務推論成本（市場價，US$/次）", "=TK!C11", None, None, None, "", "TK 頁"], [None, "0.0000"])
    row("tk2", ["對照：全成本（模型商自做成本下限，US$/次）", "=TK!C12", None, None, None, "", "TK 頁；只作對照"], [None, "0.0000"])
    row("tk3", ["FY2032 推論成本增量 ÷ 營業成本（情境 B）", "=" + P("infsh", "K"), None, None, None, "", "PL 頁"], [None, "0.000%"])
    r += 1
    head("6. 風險登錄（未建模的尾端風險；第 10、19 條）", ["風險", "是否建模", "理由", None, None, "觀察指標", None])
    for a, b, c_, d in [
        ("CyberArk 商譽稅務抵減（商譽 14.8 十億美元，大部分可抵稅）使現金稅率長期低於模型稅率", "Sens（tax_dcf 低案）", "金額與攤提年限未取得", "10-K 所得稅附註的遞延稅項與已付所得稅"),
        ("兩家經銷商合計占 FY2026 營收 30%", "Sens（tail_dist）", "單一經銷商占比未揭露", "10-K Item 1 經銷商集中度揭露"),
        ("記憶體與儲存成本上升壓縮硬體利潤（硬體約占營收一成）", "只在 nm0 低案", "管理層口述，無量化", "Product 毛利率（FY2026 75.1%）與 Q1 FY2027 毛利率"),
        ("併購整合（CyberArk、Chronosphere、Koi、Portkey、Embrace、Console）失敗或續做大型併購", "否", "無可觀察量化", "有機 NGS ARR 成長、併購費用與員工數"),
        ("可轉債（CyberArk 2030，換股價約 211.24）在價內，稀釋 5.4 百萬股", "以公允價值視為債務，不計股數", "避免重複", "轉換或贖回公告"),
        ("AI 原生競爭或雲端／模型商內建基礎安全功能，價格侵蝕高於基準", "情境 C", "無直接量測", "NGS ARR 淨新增、模組加購、同業價格公告"),
    ]:
        row("risk", [a, b, c_, None, None, d, None])
    ws.column_dimensions["A"].width = 78
    for c in "BCDE": ws.column_dimensions[c].width = 18
    ws.column_dimensions["F"].width = 60; ws.column_dimensions["G"].width = 60
    # {r} 取代
    for rr in range(3, r):
        for cc in range(1, 8):
            cell = ws.cell(rr, cc)
            if isinstance(cell.value, str) and "{r}" in cell.value: cell.value = cell.value.replace("{r}", str(rr))
    return RC


# ====================================================================== 案例
def val(x, key=None):
    if x is None: return None
    if isinstance(x, (int, float)): return float(x)
    if isinstance(x, str) and x == "=H_PEREL":
        import statistics
        return 77.10 / statistics.median([33.85, 39.22, 13.70, 174.39])
    if isinstance(x, str) and x.startswith("="):
        return float(eval(x[1:], {"MEDIAN": lambda *a: sorted(a)[len(a) // 2] if len(a) % 2 else sum(sorted(a)[len(a)//2-1:len(a)//2+1]) / 2}))
    return None

def make_cases():
    """回傳 [(label, key, setting, lo_ov or None, hi_ov or None)]。每個有低／高區間的鍵一案；加兩個群組案。"""
    out = []
    SKIP = {"cons_tp"}
    for it in IN:
        if it[0] == "sec": continue
        k, item, unit, lo, base, hi, *_ = it
        if k in SKIP: continue
        lov, hiv = val(lo, k), val(hi, k)
        bv = None
        try: bv = val(base, k)
        except Exception: bv = None
        if k in ("nm0", "sbc_pct", "amort_run", "ysd", "dep_pct", "cx_pct", "beta"):
            bv = {"nm0": 3356 / 11480, "sbc_pct": 1712 / 11480, "amort_run": 4 * 281, "ysd": None, "dep_pct": (855 - 638) / 11480, "cx_pct": (440 - 91) / 11480, "beta": (0.67 * 0.97 + 0.33 + 0.67 * 0.93 + 0.33) / 2}[k]
        has_lo = lo is not None and lov != bv
        has_hi = hi is not None and hiv != bv
        if not (has_lo or has_hi): continue
        setting = f"{k} 取 {'低' if has_lo else ''}{'、' if has_lo and has_hi else ''}{'高' if has_hi else ''}值"
        if k in ("g_px", "P_tier", "ovh", "g_tok"): setting += "（高案為上界：只動成本一邊，未連動價格與路由）"
        out.append((item, k, setting, {k: lov} if has_lo else None, {k: hiv} if has_hi else None))
    b = {k: [it for it in IN if it[0] == k][0][4] for k in ("n32_B", "m32_B", "e32_B", "k32_B")}
    cs = {}
    for k in ("n32", "m32", "e32", "k32"):
        for L in "AC": cs[f"{k}_{L}"] = b[f"{k}_B"]
    out.append(("三情境 AI 參數同時取情境 B（情境分化歸零）", "scn_collapse", "n32、m32、e32、k32 的情境 A、C 都設成情境 B 的值", cs, None))
    out.append(("價格與成本聯動（路由到便宜模型）", "cost_link", "P_tier 取 1（最低層級）且 g_px 取低值；低案＝聯動降本，高案＝層級 3 且單價不降", {"P_tier": 1.0, "g_px": -0.5}, {"P_tier": 3.0, "g_px": 0.0}))
    return out


# ====================================================================== Sens
def sensitivity(wb, ws, V, sens, fp, rec_tp, build_id):
    ws["A1"] = "Sens — 敏感度：一次動一個輸入，整本重算（LibreOffice）後記錄目標價；本頁為建置快照（X6），不是活公式"
    ws["A2"] = "讀法：低案、高案＝In 頁敏感度低、敏感度高欄；加權目標價＝DCF 與同業本益比依 w 加權。群組案「三情境 AI 參數同時取情境 B」說明情境分化的貢獻。IF_HoldEcon 群組（he_*）不適用：本模型不讀 Tokenomics 每 MW 持有成本。"
    ws["A3"] = "build_id"; ws["D3"] = build_id
    ws["A4"] = "模型狀態"
    ws["A5"] = "輸入指紋（敏感度重算時記錄；In 頁 D 至 F 欄數值合計）"; ws["D5"] = fp
    ws["A6"] = "輸入指紋（即時）"; ws["D6"] = f"=SUM(In!D4:F{INLAST})"
    ws["A7"] = "基準加權目標價（敏感度重算時記錄）"; ws["D7"] = rec_tp
    ws["A8"] = "基準加權目標價（即時）"; ws["D8"] = f"=Val!B{V['tp']}"
    ws["D4"] = '=IF(AND(ISNUMBER(D5),ISNUMBER(D7),ABS(D6-D5)<=0.000001*MAX(1,ABS(D5)),ABS(D8-D7)<0.005),"CURRENT","STALE")'
    ws["A9"] = "WACC 加減、g、β、ERP、資本支出、終值替代法（tv_alt）、股票薪酬下降路徑（sbc_dn）、營運資金比（nwc_r）都在下表各有一列；DCF 單法與 Comparable 單法並列。"
    hdr = 11
    for j, h in enumerate(["案例", "設定", "鍵", "低案加權目標價", "高案加權目標價", "低案 DCF 目標價", "高案 DCF 目標價", "低案 Comparable 目標價", "高案 Comparable 目標價"]): ws.cell(hdr, 1 + j, h).font = BOLD
    r = hdr + 1
    cases = make_cases()
    first = r
    for i, (label, key, setting, lo, hi) in enumerate(cases):
        ws.cell(r, 1, label); ws.cell(r, 2, setting); ws.cell(r, 3, key)
        if sens:
            res = sens["cases"][i]
            for j, side in enumerate(("low", "high")):
                v = res.get(side)
                if v:
                    for col, kk in ((4 + j, "tp"), (6 + j, "dcf"), (8 + j, "cmp")):
                        ws.cell(r, col, v[kk]).number_format = "#,##0.00"
        r += 1
    last = r - 1
    r += 1
    ws.cell(r, 1, "加權目標價相對基準（HTML 讀取）：案例目標價 ÷ 基準加權目標價 − 1").font = BOLD
    t0 = r; r += 1
    for j, h in enumerate(["案例", "低案相對基準", "高案相對基準"]): ws.cell(r, 1 + j, h).font = BOLD
    r += 1
    for i in range(len(cases)):
        src = first + i
        ws.cell(r, 1, f"=A{src}")
        ws.cell(r, 2, f'=IF(ISNUMBER(D{src}),D{src}/$D$7-1,"")').number_format = "+0.0%;-0.0%"
        ws.cell(r, 3, f'=IF(ISNUMBER(E{src}),E{src}/$D$7-1,"")').number_format = "+0.0%;-0.0%"
        r += 1
    ws.cell(r + 1, 1, "未涵蓋說明：he_all（IF_HoldEcon 四世代）不適用本模型；cons_tp 低高值只作對照，不進模型。")
    ws.column_dimensions["A"].width = 60; ws.column_dimensions["B"].width = 46
    for c in "DEFGHI": ws.column_dimensions[c].width = 16
    ws._t0 = t0


# ====================================================================== README
def readme(wb, ws, RW, V, build_id, sens):
    R = RW["PL"]; RA = RW["PL_A"]
    ws["A1"] = f"{build_id}｜Palo Alto Networks（PANW）個股模型 v{VER}｜命題（原型 P 按防護範圍計價）＋營收線 → 損益表｜FCFF DCF 與同業本益比各半"
    ws["A2"] = '="資料截止日 "&TEXT(data_date,"yyyy-mm-dd")&"；成品日 2026-11-16；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（與現價日相差超過 2 個交易日者標逾期：無风險利率 2026-10-27、同業與共識見各列日期）"'
    rows = {}
    def put(key, f):
        r = put.r; rows[key] = r; put.r += 1
        ws.cell(r, 1, f)
    put.r = 4
    v = lambda k, c="B": f"Val!{c}{V[k]}"
    tx = lambda ref, f: f'TEXT({ref},"{f}")'
    put("hero", '="估值評等"&' + v("vrate") + '&"；投資建議"&' + v("rec") + '&"。加權目標價 "&' + tx(v("tp"), "0") + '&" 美元對現價 "&' + tx(v("px"), "0.00") + '&" 美元（"&TEXT(px_date,"yyyy-mm-dd")&" 收盤）"&' + tx(v("up"), "+0%;-0%") + '&"；DCF "&' + tx(v("tpdcf"), "0") + '&"、同業本益比 "&' + tx(v("tpcmp"), "0") + '&"、賣方共識 "&' + tx(v("cons"), "0") + '&"。條件：核心營業利益率與股票薪酬占營收持平；現價隱含核心營業利益率 "&' + tx(f"RevDCF!B{{m1}}", "0%") + '&"，模型 "&' + tx("nm0", "0%") + '&"。"')
    put("rec", '="在核心 non-GAAP 營業利益率與股票薪酬占營收持平、永續成長 "&TEXT(g,"0%")&"、同業中位前瞻本益比 "&' + tx(v("pemed"), "0.0") + '&" 倍的基準假設下，加權目標價 "&' + tx(v("tp"), "0.00") + '&" 美元、上檔空間 "&' + tx(v("up"), "+0.0%;-0.0%") + '&"，估值評等"&' + v("vrate") + '&"（收支關卡"&' + v("gtxt") + '&"，FY2027–FY2031 自由現金流合計 "&' + tx(v("gate"), "#,##0") + '&" 百萬美元）。失效條件：核心營業利益率擴張或 SBC 占營收下降（Sens 有列），或同業估值上移。"')
    put("t1", '="命題：AI 讓資安的受保護範圍從人與端點擴大到代理（非人類身分），NGS ARR ＝ 受保護單位 × 每單位營收，使用者問的是 AI 帶來的營收增量，能否大於定價侵蝕與推論成本。"')
    put("t2", '="營收線：Product 與 Subscription and support（單一應報導分部下 10-K 固定揭露的兩條線）；命題模組只驅動 Subscription and support 之下的 NGS ARR 相關營收。"')
    put("t3", '="未做先驗（工作單 G 節：先驗跳過）；單位數、每單位年價、模組數公司不揭露，只作比例尺（Assumed）。"')
    put("a1", '="問題一：AI 增量是否大於侵蝕與推論成本？答：情境 B 下 FY2032 AI 淨效果（增量 − 侵蝕 − 推論成本）為 "&TEXT(PL!K' + str(R["ainet"]) + ',"#,##0")&" 百萬美元（情境 A "&TEXT(PL_A!K' + str(R["ainet"]) + ',"#,##0")&"、情境 C "&TEXT(PL_C!K' + str(R["ainet"]) + ',"#,##0")&"）；推論成本增量只占營業成本的 "&TEXT(PL!K' + str(R["infsh"]) + ',"0.00%")&"，量級遠小於增量與侵蝕，因此結果由代理身分數、加購與侵蝕三個 Assumed 決定，不由推論成本決定。"')
    put("a2", '="問題二：現價要求什麼？答：DCF 單法下，現價隱含核心 non-GAAP 營業利益率 "&TEXT(RevDCF!B{m1},"0.0%")&"，"&IF(RevDCF!B{m1}>nm0,"高於","低於")&"模型的 "&TEXT(nm0,"0.0%")&"；現價隱含 EV 為 FY2027E 營收的 "&TEXT(RevDCF!B{evs},"0.0")&" 倍。這是模型輸出，不是獨立證據。"')
    put("a3", '="問題三：與管理層目標的差距？答：模型情境 B 的 FY2030 NGS ARR 為 "&TEXT(PL!I' + str(R["ARR"]) + ',"0.0")&" 十億美元，管理層目標 "&TEXT(t_arr30,"0")&" 十億美元；差距需要每年額外 "&TEXT(Recon!B{req}*100,"0.0")&" 個百分點的成長才能彌補（管理層目標為 Interested-party，不校準）。"')
    put("o1", '="1. 評等主要由估值決定：DCF 目標價 "&' + tx(v("tpdcf"), "0") + '&" 美元、同業本益比目標價 "&' + tx(v("tpcmp"), "0") + '&" 美元，兩者"&IF(MAX(' + v("tpdcf") + ',' + v("tpcmp") + ')<' + v("px") + ',"都低於","並非都低於")&"現價 "&' + tx(v("px"), "0.00") + '&" 美元。"')
    put("o2", '="2. DCF 偏低的主因是股票薪酬：FY2032 股票薪酬占營收 "&TEXT(RevDCF!B{sbc},"0.0%")&"，標準 FCFF 以 NOPAT 扣股票薪酬、不加回；Sens 有股票薪酬占營收低案與每年下降案。"')
    put("o3", '="3. 營收成長含併購：FY2026 有機年增 "&TEXT(Recon!B{orgg},"0.0%")&"，合併年增 "&TEXT(h_rev_26/h_rev_25-1,"0.0%")&"；模型 FY2027 營收 "&TEXT(PL!F' + str(R["rev"]) + '/1000,"0.0")&" 十億美元，"&Recon!F{rev}&"。"')
    put("o4", '="4. 同業中位前瞻本益比 "&' + tx(v("pemed"), "0.0") + '&" 倍（4 家，不做迴歸），PANW 資料商前瞻本益比為其 "&' + tx(v("pxpe"), "0.0") + '&" 倍；調整係數 1.00（缺成長與利益率資料）；高案取 PANW 目前溢價。"')
    put("o5", '="5. AI 推論成本不是重點：FY2032 推論成本增量只占營業成本 "&TEXT(PL!K' + str(R["infsh"]) + ',"0.00%")&"；全成本對照（模型商自做）占 "&TEXT(PL!K' + str(R["infsh2"]) + ',"0.000%")&"。可證偽指標：Subscription and support 毛利率走勢（FY2026 "&TEXT(0.692,"0.0%")&"）。"')
    put("drv", '="驅動因子：Product＝上年 ×（1＋成長）；Subscription and support＝NGS ARR 相關營收（年平均 NGS ARR × 列入比例）＋其他；NGS ARR＝受保護單位 × 每單位營收（R0 ×（1 − 侵蝕 ＋ 加購 ＋ 代理身分數 × 代理單價比））；三項相加為合併營收。"')
    put("ver", '="v' + VER + '（2026-11-16）：首版。命題原型 P；DCF 標準 FCFF 至 FY2032；同業本益比 4 家；Sens 整本重算 "&COUNTA(Sens!C12:C60)&" 個鍵；未做先驗。"')
    r = put.r + 1
    ws.cell(r, 1, "不適用與未做項目（工作單 G 節）").font = BOLD; r += 1
    for t in ["不適用：Wattonomics 對照與自用算力（Meta 的自用 AI 讀法 2、在役 MW、Wattonomics 上線時程）——理由：PANW 不自建算力，推論成本以向模型商購買的市場價計（Tokenomics 快照，只讀）。",
              "不適用：Tokenomics 每 MW 持有成本（IF_HoldEcon）群組與 he_all 案——理由：本模型不讀每 MW 成本。",
              "不適用：三層協調（Coord 頁、C0、兩讀數）——理由：非 Siliconomics 覆蓋。",
              "不適用：分部加總（SOTP）——理由：單一應報導分部；Comparable 以合併 non-GAAP EPS × 同業中位本益比。",
              "不適用：收支模組、反向 DCF 之外的融資瀑布——理由：淨現金為正（FY2026 年底現金與投資 7.9 十億美元），不以新債補缺口。",
              "未做：先驗（工作單 G 節 2026-11-16 23:27：先驗跳過）。",
              "v0.2 計畫：該版改讀 software_ai/unit_economics 的 P 原型輸出（本版 In 頁自建，變數名稱一致）。",
              "與 unit_economics 的差異：本模型的成本不隨價格侵蝕下降（侵蝕不減成本），unit_economics 的 cogsx 隨淨營收；推論成本與 AI 增量取相對 FY2026 的差。",
              "承重 Assumed 與可證偽指標：(1) 受保護代理身分數與加購（n32、m32）——可證偽：公司揭露非人類身分產品營收或客戶採用；(2) 價格侵蝕（e32）——Subscription and support 毛利率與 NGS ARR 淨新增；(3) 核心 non-GAAP 營業利益率持平與 SBC 占營收持平——FY2027 Q1 起的營業利益率與 SBC 占營收實績。",
              "公司指標與模型參數的差異：NGS ARR 為公司自訂 KPI（含併購、排除硬體與 legacy），模型以其為 FY2026 錨點並標 Interested-party；調整後自由現金流含加回項目，模型 FCF 為營業現金流減資本支出。"]:
        ws.cell(r, 1, t); r += 1
    # 解析 {m1} 等
    res = {"m1": RW_REV["m1"], "evs": RW_REV["evs"], "sbc": RW_REV["sbc"], "req": RW_RC["req"], "orgg": RW_RC["orgg"], "rev": RW_RC["rev"]}
    for rr in range(1, r + 1):
        c = ws.cell(rr, 1)
        if isinstance(c.value, str) and "{" in c.value: c.value = re.sub(r"\{(\w+)\}", lambda m: str(res[m.group(1)]), c.value)
    ws.column_dimensions["A"].width = 160
    ws._rows = rows


def sources(ws):
    ws["A1"] = "Sources — 出處清單（逐列紀錄見 data/facts_paloalto.csv）"
    r = 3
    for k, v in SRC.items():
        ws.cell(r, 1, k); ws.cell(r, 2, v); r += 1
    for t in ["WebSearch／WebFetch 擷取日 2026-11-16", "software_ai/unit_economics/data/tk_snapshot.json（Tokenomics v22.0）"]:
        ws.cell(r, 2, t); r += 1
    ws.column_dimensions["B"].width = 160


# ====================================================================== 重算與主流程
def recalc(paths, outdir):
    os.makedirs(outdir, exist_ok=True)
    prof = os.path.join(SCR, "lo_profile")
    for i in range(0, len(paths), 25):
        batch = paths[i:i + 25]
        subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", outdir] + batch,
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"), timeout=900)

def read_out(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    v = wb["Val"]; lab = {str(v.cell(r, 1).value): r for r in range(1, v.max_row + 1) if v.cell(r, 1).value}
    g = lambda s: v.cell(lab[s], 2).value
    out = {"tp": g("加權目標價（DCF 與同業本益比加權）"), "dcf": g("DCF 目標價（三情境加權）"), "cmp": g("Comparable 目標價（同業前瞻本益比中位數法，三情境加權）")}
    return out, wb

def input_hash():
    h = hashlib.sha256()
    for it in IN:
        h.update(repr(it).encode("utf-8"))
    h.update(TK["source"]["commit"].encode())
    return h.hexdigest()[:10]

def main():
    os.makedirs(SCR, exist_ok=True)
    bid = f"v{VER}-{input_hash()}"
    wd = os.path.join(SCR, "run"); shutil.rmtree(wd, ignore_errors=True); os.makedirs(wd)
    base = build(build_id=bid); bp = os.path.join(wd, "base.xlsx"); base.save(bp)
    cases = make_cases()
    paths = [bp]; idx = []
    if "--no-sens" not in sys.argv:
        for i, (label, key, setting, lo, hi) in enumerate(cases):
            for side, ov in (("low", lo), ("high", hi)):
                if ov:
                    p = os.path.join(wd, f"c{i:03d}_{side}.xlsx"); build(ov=ov, build_id=bid).save(p); paths.append(p); idx.append((i, side, p))
    recalc(paths, os.path.join(wd, "out"))
    bout, bwb = read_out(os.path.join(wd, "out", "base.xlsx"))
    ws = bwb["In"]
    fp = bwb["Sens"]["D6"].value
    print("基準：", bout, "fp", fp)
    sens = {"base": bout, "cases": [dict() for _ in cases]}
    for i, side, p in idx:
        o, _ = read_out(os.path.join(wd, "out", os.path.basename(p))); sens["cases"][i][side] = o
    final = build(sens=sens, fp=fp, rec_tp=bout["tp"], build_id=bid)
    fr = os.path.join(wd, OUTNAME); final.save(fr)
    recalc([fr], os.path.join(wd, "fin"))
    os.makedirs(os.path.join(HERE, "dist"), exist_ok=True)
    for old in glob.glob(os.path.join(HERE, "dist", "*.xlsx")): os.remove(old)
    shutil.copy(os.path.join(wd, "fin", OUTNAME), os.path.join(HERE, "dist", OUTNAME))
    json.dump(sens, open(os.path.join(HERE, "data", f"sens_v{VER.replace('.', '')}_{VDATE}.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("完成", OUTNAME)

if __name__ == "__main__":
    main()

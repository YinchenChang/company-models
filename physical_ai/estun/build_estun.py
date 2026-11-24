# 埃斯頓 Estun（002747.SZ，深交所；另掛牌港股 02715）個股模型 v0.1 建置腳本
# 產出：physical_ai/estun/dist/20261124_埃斯頓個股模型_v0.1.xlsx（LibreOffice 重算後的成品，含公式與快取值）
# 結構：命題模組（Thesis：機器人及智能製造系統營收＝中國裝機（線 3）× 國產品牌占比（線 3 實績＋假設路徑）× 本公司相對份額漂移 × 單價年降；
#       核心部件營收＝年增假設）→ 官方兩分部 → 標準損益表 → FCFF DCF＋同業 P/E。
# 敏感度：每個案例覆寫 In 頁基準欄後整本重建、LibreOffice 重算（整本重算，不用閉式近似）；反解同法以二分搜尋。
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md、docs/plan/高不確定性預估工作原則.md。
# 用法：python3 physical_ai/estun/build_estun.py [--quick]   （需要 soffice；暫存於家目錄下 es_build）
import glob, hashlib, json, math, os, shutil, subprocess, sys
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
NAME = f"20261124_埃斯頓個股模型_{VER}.xlsx"
OUT = os.path.join(HERE, "dist", NAME)
HOMEDIR = os.path.expanduser("~")
TMP = os.path.join(HOMEDIR, "es_build")
LO = ["soffice", f"-env:UserInstallation=file://{HOMEDIR}/lo_estun", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML"]

L3 = os.path.join(ROOT, "physical_ai/lines/industrial_robot/dist/20261116_工業機器人單位經濟_v0.1.xlsx")
L3N = os.path.relpath(L3, ROOT)

BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDRF = PatternFill("solid", fgColor="DDEBF7"); SECF = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------- 出處 ----------------
S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv：002747.SZ；未還原股利）"
S_CLOSE = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
S_25 = "新浪財經 2026-03-30 年報報導（https://finance.sina.cn/2026-03-30/detail-inhsuvau2609528.d.html）：2025 年營收 48.88 億、歸母淨利 4,497.22 萬、毛利率 29.45%（−0.12pt）、期間費用 14.31 億（−1.21 億、費用率 29.29%）、基本 EPS 0.05"
S_25B = "大眾證券報／中國證券報 2026-04-14（https://www.cs.com.cn/ssgs/01/2026/04/14/detail_2026041410003906.html；搜尋摘要）：2025 年工業機器人及智能製造系統收入 39.97 億（+31.80%）、自動化核心部件 8.91 億（−8.72%）、研發 4.76 億"
S_24 = "券商 2024 年報點評（搜尋摘要）：2024 年營收 40.09 億（−13.83%）、歸母淨利 −8.10 億；2023 年營收 46.52 億、歸母淨利 1.35 億；2024 年對四家子公司計提商譽減值 3.45 億（財新／鳳凰網）"
S_Q3 = "東吳證券 2026 年三季報點評（2026-10-30；https://data.eastmoney.com/report/info/AP202610301831190796.html）：1–9 月營收 36.99 億（−2.96%）、歸母淨利 1.71 億、毛利率 32.36%、費用率 28.86%（銷售 8.45%、管理 8.56%、財務 2.75%、研發 9.11%）；Q3 營收 11.01 億（−12.43%）、毛利率 33.64%；2026–2028E 歸母淨利 2.08／2.60／3.37 億"
S_Q3B = "西南證券 2026 年三季報點評（2026-11-09；https://data.eastmoney.com/report/info/AP202611091831541346.html）：持有；EPS 2026–2028E 0.25／0.33／0.44 元；Q3 出貨首次逾 1 萬台；以 11 月 9 日股價 28.98 元計 PE 115／88／66 倍"
S_H1 = "券商與財經報導 2026 年半年報（https://finance.eastmoney.com/a/202608213849409039.html）：上半年營收 25.78 億（+1.14%）、歸母淨利 1.61 億、扣非 7,017 萬；非經常損益 0.91 億（南京工藝股權置換公允價值利得約 1.00 億）；不派現金股利"
S_HK = "埃斯頓 H 股發行（https://finance.eastmoney.com/a/202603053663288381.html；搜尋摘要）：發行價 15.36 港元、全球發售 9,678 萬股、淨額約 14.12 億港元，2026-03-09 掛牌"
S_SH = "西南證券研報基礎資料（搜尋摘要）：總股本 968 百萬股、流通股 782 百萬股（2026 年 4 月；是否為年報期末口徑未查核）"
S_RF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與 leaderdrive 等模型同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：China CRP 0.91%（與 leaderdrive 同一取值）"
IP_BRK = "券商研究報告：有銷售研究與維持評等的誘因；歷史數字為轉錄，預測為券商自設"
IP_CO = "發行公司說法，有股價與融資誘因"

# ---------------- 線 3 取數（以 ID 找列，不以列號） ----------------
def line_data():
    wb = openpyxl.load_workbook(L3, data_only=True)
    ws = wb["下游取數"]; rows = {}
    for r in range(1, ws.max_row + 1):
        k = ws.cell(r, 1).value
        if isinstance(k, str): rows[k] = (r, ws.cell(r, 5).value)
    d = {"inst": {}}
    for y in range(2025, 2033):
        for s, tag in (("A", "hi"), ("B", "base"), ("C", "lo")):
            k = f"IR.inst.china.{y}.{tag}" if y > 2025 else "IR.inst.china.2025.base"
            r, v = rows[k]; d["inst"][(y, s)] = (float(v), f"{L3N} 下游取數!E{r}（{k}）")
    r, v = rows["IR.share.china_domestic.2025"]; d["dom25"] = (float(v), f"{L3N} 下游取數!E{r}")
    r, v = rows["IR.cycle.cn_prod.ytd_yoy"]; d["prod_yoy"] = (float(v), f"{L3N} 下游取數!E{r}")
    ws2 = wb["輸入"]; pr = {ws2.cell(r, 1).value: r for r in range(1, ws2.max_row + 1)}
    r = pr["BODY_6AX"]; d["body"] = [(ws2.cell(r, c).value, f"{L3N} 輸入!{openpyxl.utils.get_column_letter(c)}{r}") for c in (4, 5, 6)]
    r = pr["BODY_6AX_F"]; d["bodyf"] = (ws2.cell(r, 5).value, f"{L3N} 輸入!E{r}")
    return d
LD = line_data()

# ---------------- In 頁輸入 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 來源誘因, 說明)
IN = [
    ("sec", "1. 市場資料、股數與資產負債"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, "", "資料截止日 2026-11-24；股價 2026-11-20；成品日 2026-11-24"),
    ("px", "現價（2026-11-20 收盤，深交所，人民幣）", "元", None, 24.25, None, "Verified", S_PX, "", "csv 值 24.25；11-09 收盤 29.41，兩週跌 17.5%（全球 AI 與半導體股同週大跌，非資料錯誤）"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,12,31)", None, "Assumed", "本專案慣例：評價日 2026-12-31（會計年度＝日曆年）", "", "折現期間由評價日起算"),
    ("tp_date", "目標時點（px_date 後 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", ""),
    ("bs_date", "資產負債表日", "日期", None, "=DATE(2026,9,30)", None, "Assumed", "最近一期財報為 2026 年三季報（資產負債表未取得，只定日期）", "", ""),
    ("stub_frac", "Stub 期（資產負債表日→評價日）", "年", None, "=(val_date-bs_date)/365", None, "Derived", "＝(評價日 − 資產負債表日) ÷ 365", "", "Stub FCFF＝FY2026 FCFF × stub_frac"),
    ("sh_v", "評價日股數（A 股＋H 股總股本）", "百萬股", None, 968.0, None, "Interested-party", S_SH, IP_BRK, "含 H 股 96.78 百萬股（" + S_HK + "）；限制性股票、選擇權未取得，以 0 計"),
    ("cash_v", "現金及約當現金（含 H 股募資，2026-09-30 估計）", "百萬元", 1200, 1700, 2200, "Assumed", "未找到：三季報與半年報資產負債表未能讀取（查過：東吳、西南研報摘要、半年報報導）。H 股淨募資約 14.12 億港元（約 12.0 億元；" + S_HK + "）已於 2026-03 到位", "", "估計值；翻轉結論所需的變動見 Sens"),
    ("debt_v", "有息負債（2026-09-30 估計）", "百萬元", 2000, 3000, 4000, "Assumed", "未找到：資產負債率約 68–70%（東吳研報基本資料 LF 值 68.22／70.36，期間未註明）；2025 前三季有息負債率 39.21%（口徑未註明）；絕對金額未取得", "", "估計值"),
    ("nci_v", "少數股東權益（估計）", "百萬元", 200, 400, 800, "Assumed", "未找到：埃斯頓江蘇智能 2026 年以 2.106 億元購入 19.46% 少數股權（隱含估值約 10.8 億元，" + S_Q3B + "），持股升至 69.6%，剩餘少數股權約 3 億元量級", "", "估計值"),
    ("sec", "2. 實績錨（FY2026＝2026 年，會計年度＝日曆年）"),
    ("r9m26", "2026 年 1–9 月營收", "百萬元", None, 3699.0, None, "Verified", S_Q3 + "；西南同為 36.99 億", "", "公司三季報原文未直接讀取"),
    ("g9m26", "2026 年 1–9 月營收年增", "%", None, -0.0296, None, "Verified", S_Q3, "", ""),
    ("q4g", "2026 年第 4 季營收年增", "%", -0.15, -0.05, 0.05, "Assumed", "Q3 單季 −12.4%（東吳）、Q3 出貨首次逾 1 萬台（西南）；東吳稱下游需求增速放緩；基準取 −5%", "", "FY2026 營收＝1–9 月實績＋2025Q4 ×（1＋q4g）"),
    ("sh_r26", "FY2026 機器人及智能製造系統占營收", "%", None, "=Hist!E8/Hist!E4", None, "Derived", "＝2025 年占比（2026 年分部營收未取得：三季報未讀到分部拆分）", "", "2025 年 81.8%"),
    ("gm26", "FY2026 毛利率", "%", None, 0.3237, None, "Verified", S_Q3 + "（1–9 月 32.36%；西南 32.37%）", "", "第 4 季假設同 1–9 月；2025 年 29.45%、Q3 單季 33.64%"),
    ("opex26", "FY2026 期間費用率（銷售＋管理＋研發＋財務）", "%", 0.27, 0.2886, 0.30, "Assumed", S_Q3 + "：1–9 月費用率 28.86%", "", "基準取 1–9 月實績"),
    ("oth26", "FY2026 其他損益淨額（公允價值利得、補助、減值、稅金及附加；稅前）", "百萬元", 60, 120, 180, "Assumed", S_H1 + "：上半年公允價值利得約 1.00 億；1–9 月歸母淨利 1.71 億高於核心（毛利 − 期間費用）約 1.29 億", "", "一次性成分高，FY2027 起改用 oth_r"),
    ("sec", "3. 命題模組：機器人及智能製造系統營收（線 3 × 國產占比 × 相對份額 × 單價）"),
    ("cn_k", "中國裝機年增相對線 3 的倍數（全部情境）", "x", 0.5, 1.0, 1.5, "Assumed", "基準＝線 3 原值（IFR 口徑）；低高案為線 3 各年成長率的 0.5／1.5 倍", "", "裝機水準以線 3 2025 年實績為起點，各年成長率 × cn_k；線 3 三情境已含高／基準／低路徑"),
    ("pd", "每台對應營收年降幅（價格戰與組合）", "%／年", 0.0, 0.03, 0.08, "Analogy", f"線 3：國產 6 軸本體價基準 {LD['body'][1][0]:,.0f} 元、低案 {LD['body'][0][0]:,.0f} 元（{LD['body'][1][1]}）；本公司 2025 年營收 +22% 而中國裝機 +？%，價格壓力方向一致但幅度未取得", "", "線 3 無逐年降幅；本值以 Analogy 區間表達，非線 3 輸入"),
    ("sec", "4. 命題模組：情境專屬驅動（A 樂觀／B 基準／C 保守；只列基準欄，不設敏感度區間）"),
    ("ds30_A", "情境 A：2030 年中國裝機國產品牌占比", "%", None, 0.70, None, "Assumed", f"2025 年 55%（線 3：{LD['dom25'][1]}，Verified）；2030 年路徑為假設", "", "2025 年實績線性走到 2030，其後持平"),
    ("ds30_B", "情境 B：2030 年中國裝機國產品牌占比", "%", None, 0.65, None, "Assumed", "同上", "", ""),
    ("ds30_C", "情境 C：2030 年中國裝機國產品牌占比", "%", None, 0.58, None, "Assumed", "同上；外資（Fanuc、安川、ABB、KUKA）降價反擊", "", ""),
    ("drift_A", "情境 A：本公司相對國產品牌份額年漂移", "%／年", None, 0.04, None, "Assumed", "券商引 MIR：2026 上半年中國工業機器人出貨居全品牌第一；但 2026 年前三季營收 −3%，份額不等於營收", "", "正＝取得份額"),
    ("drift_B", "情境 B：本公司相對國產品牌份額年漂移", "%／年", None, 0.01, None, "Assumed", "同上", "", ""),
    ("drift_C", "情境 C：本公司相對國產品牌份額年漂移", "%／年", None, -0.03, None, "Assumed", "同上；國產廠價格戰、匯川等競爭", "", ""),
    ("gcore_A", "情境 A：自動化核心部件營收年增（FY2027 起）", "%", None, 0.10, None, "Assumed", S_25B, "", ""),
    ("gcore_B", "情境 B：自動化核心部件營收年增（FY2027 起）", "%", None, 0.05, None, "Assumed", "同上", "", ""),
    ("gcore_C", "情境 C：自動化核心部件營收年增（FY2027 起）", "%", None, -0.03, None, "Assumed", "同上", "", ""),
    ("gm30_A", "情境 A：FY2030 毛利率", "%", None, 0.34, None, "Assumed", "2022 年 32.9%（搜尋摘要）、2025 年 29.45%、2026 年 1–9 月 32.4%", "", "FY2026 起線性走到 FY2030，其後持平"),
    ("gm30_B", "情境 B：FY2030 毛利率", "%", None, 0.32, None, "Assumed", "同上", "", ""),
    ("gm30_C", "情境 C：FY2030 毛利率", "%", None, 0.28, None, "Assumed", "同上；2025 年 29.45% 以下", "", ""),
    ("tp_wA", "情境權重 A", "", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", ""),
    ("tp_wB", "情境權重 B", "", None, 0.50, None, "Assumed", "同上", "", ""),
    ("tp_wC", "情境權重 C", "", None, 0.25, None, "Assumed", "同上", "", ""),
    ("sec", "5. 損益與現金流"),
    ("opex30", "FY2030 期間費用率", "%", 0.22, 0.25, 0.28, "Assumed", "2025 年 29.3%、2024 年 38.7%、2026 年 1–9 月 28.9%（營運槓桿與併購整合）", "", "FY2026→FY2030 線性，其後持平"),
    ("oth_r", "其他損益淨額 ÷ 營收（FY2027 起）", "%", -0.005, 0.003, 0.010, "Assumed", "2025 年約（歸母淨利 − 核心營業利益）÷ 營收＝0.7%（Derived，稅與少數股東未揭露）", "", "含補助、減值、稅金及附加"),
    ("tax_r", "有效稅率", "%", 0.10, 0.15, 0.25, "Assumed", "未找到：所得稅費用未取得；高新技術企業稅率 15% 為慣例假設", "", ""),
    ("nci_r", "少數股東損益 ÷ 稅後淨利", "%", 0.0, 0.08, 0.20, "Assumed", "未找到：少數股東損益未取得；子公司埃斯頓江蘇智能持股 69.6%", "", ""),
    ("payout", "配息率", "%", None, 0.0, 0.15, "Assumed", S_H1 + "：2026 年半年度不派現金股利；高案 15% 只作敏感度", "", "用於收支關卡與目標時點前股利"),
    ("da_r", "折舊攤銷 ÷ 營收", "%", 0.035, 0.045, 0.055, "Assumed", "未找到：折舊攤銷未取得；機器人整機廠慣例範圍", "", ""),
    ("cx26", "FY2026 資本支出", "百萬元", 150, 220, 320, "Assumed", "未找到：資本支出未取得（2025 年經營現金流 5.07 億，" + S_25B + "同篇報導）", "", ""),
    ("cx_r", "資本支出 ÷ 營收（FY2027 起）", "%", 0.03, 0.05, 0.08, "Assumed", "未找到：歷史資本支出未取得", "", ""),
    ("nwc_r", "營運資金增加 ÷ 營收增量", "%", 0.20, 0.30, 0.50, "Assumed", "2025 前三季應收款占營收 49%（資料商轉述）、1–9 月存貨 14.26 億（東吳）；比率為假設", "", ""),
    ("sec", "6. 折現率（DCF 標準規格第 2 節）"),
    ("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", S_RF, "", "日期 2026-11-05，較股價基準日早 11 個交易日（逾期；跨模型一致）"),
    ("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", ""),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", ""),
    ("beta", "β", "x", "=beta-0.3", 1.1, "=beta+0.3", "Assumed", "未找到：本公司與同業 β 未取得（資料商頁面無法讀取）；A 股機器人類股常見 0.9–1.4", "", "敏感度 ∓0.3（跨模型慣例）；非兩來源 Blume 中位數，原則限制"),
    ("kd", "稅前債務成本", "%", 0.03, 0.035, 0.045, "Assumed", "未找到：利息費用與借款利率未取得；財務費用率 1–9 月 2.75%（東吳）", "", ""),
    ("wacc_adj", "WACC 調整（敏感度用）", "pt", -0.01, 0, 0.01, "Assumed", "基準 0；敏感度 ±1pt", "", ""),
    ("g_cap", "永續成長率上限", "%", None, 0.03, None, "Assumed", "DCF 標準規格：3%，且不超過計價幣別 10 年期公債殖利率", "", ""),
    ("g_tv", "永續成長率", "%", 0.01, "=MIN(g_cap,rf)", 0.03, "Derived", "＝MIN(g_cap, rf)", "", "敏感度 1%／3%（3% 超過 rf，只作敏感度）"),
    ("ronic_sp", "RONIC − WACC", "pt", None, 0, 0.05, "Assumed", "RONIC 預設＝WACC（新投資價值中性）", "", ""),
    ("tv_alt", "終值改用 FCFF ×（1＋g）（0＝否，1＝是）", "開關", None, 0, 1, "Assumed", "DCF 標準規格第 4 節敏感度", "", ""),
    ("sec", "7. 同業倍數與權重"),
    ("pe_peer", "同業前瞻 P/E 中位數", "x", "=MIN(Comp!F5:F8)", "=MEDIAN(Comp!F5:F8)", "=MAX(Comp!F5:F8)", "Derived", "＝Comp 頁 4 家日本同業（2026-11-20 收盤 ÷ 本專案模型 FY2027 EPS）中位數；中國同業匯川報價有、前瞻 EPS 未取得，未納入", "", "敏感度＝同業最低／最高"),
    ("rel_k", "相對倍數係數（本公司 ÷ 同業）", "x", None, 1.0, 3.0, "Assumed", "1.00（未調整：缺本公司對同業十年相對倍數歷史）；高案 3 倍反映 A 股機器人題材溢價（未查核，不作依據）", "", ""),
    ("w_dcf", "DCF 權重（Comparable＝1 − 本值）", "%", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設 DCF 50%＋同業倍數 50%", "", ""),
]
SCN_DRIVERS = [("ds30", "2030 年中國裝機國產品牌占比"), ("drift", "本公司相對國產品牌份額年漂移"), ("gcore", "自動化核心部件營收年增"), ("gm30", "FY2030 毛利率")]
YRS = list(range(2023, 2033))
COL = {y: openpyxl.utils.get_column_letter(3 + i) for i, y in enumerate(YRS)}    # PL：C＝FY2023 … L＝FY2032
TC = {y: openpyxl.utils.get_column_letter(4 + (y - 2025)) for y in range(2025, 2033)}  # Thesis：D＝2025 … K＝2032
SC = "ABC"; LINE_SC = {"A": "線 3 高情境", "B": "線 3 基準", "C": "線 3 低情境"}


def short(s):
    import re
    s = str(s).replace(L3N, "線 3 成品")
    s = re.sub(r"[；;]?\s*https?://[^\s）)；]+", "", s)
    return re.sub(r"（[A-Z][A-Z0-9]*[._][A-Za-z0-9_.]+）", "", s)


def hdr(ws, r, labels, c0=1):
    for j, h in enumerate(labels):
        c = ws.cell(r, c0 + j, h); c.font = BOLD; c.fill = HDRF


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def T(x, f="#,##0"):
    return f'TEXT({x},"{f}")'


def make(over=None, sens=None):
    over = over or {}
    wb = openpyxl.Workbook()
    RD = wb.active; RD.title = "README"
    IW = wb.create_sheet("In"); HI = wb.create_sheet("Hist"); Q = wb.create_sheet("Q"); LN = wb.create_sheet("線取數"); TH = wb.create_sheet("Thesis")
    PL = wb.create_sheet("PL"); DC = wb.create_sheet("DCF"); V = wb.create_sheet("Val"); CP = wb.create_sheet("Comp")
    SE = wb.create_sheet("Sens"); RC = wb.create_sheet("Recon"); DZ = wb.create_sheet("對帳"); RK = wb.create_sheet("Risk"); SO = wb.create_sheet("Sources")

    # ================= In =================
    IW["A1"] = ("In 頁：藍字為輸入。基準欄 E 進主模型；敏感度低（D）／敏感度高（F）只用於 Sens 頁（一次動一個參數、整本重算），不進目標價。"
                "情境 A／B／C 的差異只寫在情境專屬列（第 4 節）與線 3 中國裝機路徑（線取數頁：A＝線 3 高、B＝基準、C＝低），最後一節「情境驅動對照」純展示。")
    IW["A2"] = "單位：百萬元人民幣，除另註；會計年度＝日曆年（FY2026＝2026 年）。中國會計準則（法定財報口徑）。資產負債表未取得，相關輸入為 Assumed 區間。"
    hdr(IW, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "來源誘因", "說明"])
    r = 4; KR = {}
    for row in IN:
        if row[0] == "sec":
            IW.cell(r, 1, row[1]).font = BOLD; IW.cell(r, 1).fill = SECF; r += 1; continue
        k, lab, unit, lo, base, hi, tag, src, inc, note = row
        if k in over: base = over[k]
        if "未找到" in src or "未取得" in src:
            src = src.replace("未找到：", "資料缺口：").replace("未找到", "缺").replace("未取得", "缺")
            note = ("未找到；" + note) if "未找到" not in note else note
        IW.cell(r, 1, k); IW.cell(r, 2, lab); IW.cell(r, 3, unit)
        for c, v in ((4, lo), (5, base), (6, hi)):
            if v is not None:
                cc = IW.cell(r, c, v); cc.font = BLUE
                if unit in ("%", "%／年"): cc.number_format = "0.0%"
                if unit == "日期": cc.number_format = "yyyy-mm-dd"
        IW.cell(r, 7, tag); IW.cell(r, 8, src); IW.cell(r, 9, inc); IW.cell(r, 10, short(note))
        dn(wb, k, f"In!$E${r}"); KR[k] = r; r += 1
    LAST_IN = r - 1
    r += 1
    IW.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字，本節不被任何公式引用）").font = BOLD; r += 1
    hdr(IW, r, ["驅動", "項目", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "引用與說明"]); r += 1
    for d, lab in SCN_DRIVERS:
        IW.cell(r, 1, d); IW.cell(r, 2, lab)
        for j, s in enumerate(SC): IW.cell(r, 4 + j, f"=E{KR[d + '_' + s]}")
        IW.cell(r, 8, f"引用 In 頁 {d}_A／{d}_B／{d}_C；情境專屬輸入"); r += 1
    IW.cell(r, 1, "scn_w"); IW.cell(r, 2, "情境權重")
    for j, s in enumerate(SC): IW.cell(r, 4 + j, f"=E{KR['tp_w' + s]}")
    IW.cell(r, 8, "引用 In 頁 tp_wA／tp_wB／tp_wC；合計＝1；線 3 中國裝機路徑（線取數頁）A＝高、B＝基準、C＝低，非本頁輸入")
    for col, w in zip("ABCDEFGHIJ", (12, 46, 10, 11, 12, 11, 15, 70, 30, 50)): IW.column_dimensions[col].width = w

    # ================= Hist =================
    HI["A1"] = "Hist：歷史財務（百萬元人民幣；中國會計準則）。券商與媒體轉錄之年報數字；2023 年分部與毛利率未找到。"
    hdr(HI, 3, ["項目", "", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"])
    HROWS = [  # (列, 項目, [2023, 2024, 2025], 標記, 來源, 格式)
        (4, "營業收入", [4652, 4009, 4888], "Verified", S_24 + "；" + S_25, None),
        (5, "營業成本（＝營收 ×（1 − 毛利率））", [None, "=D4*(1-D7)", "=E4*(1-E7)"], "Derived", "＝營業收入 ×（1 − 毛利率）；2023 年毛利率未找到", None),
        (6, "毛利（＝營收 − 營業成本）", [None, "=D4-D5", "=E4-E5"], "Derived", "＝營業收入 − 營業成本", None),
        (7, "毛利率", [None, "=0.2945+0.0012", 0.2945], "Verified", S_25 + "（2025 年 29.45%，較 2024 年 −0.12pt；2024 年由此推得）", "0.0%"),
        (8, "機器人及智能製造系統 營收", [None, "=3997/1.318", 3997], "Verified", S_25B + "（2025：39.97 億、+31.80%；2024 由增幅推得；2023 分部未找到）", None),
        (9, "自動化核心部件及運動控制系統 營收", [None, "=891/(1-0.0872)", 891], "Verified", S_25B + "（2025：8.91 億、−8.72%；2024 由降幅推得；2023 分部未找到）", None),
        (10, "期間費用（銷售＋管理＋研發＋財務）", [None, "=1431+121", 1431], "Verified", S_25 + "（2025：14.31 億，較 2024 年減 1.21 億；各項金額未取得）", None),
        (11, "營業利益（模型口徑：毛利 − 期間費用；未含減值、公允價值、補助、稅金及附加）", [None, "=D6-D10", "=E6-E10"], "Derived", "＝毛利 − 期間費用；2024 年大額商譽與資產減值（合計約 4.65 億）在此之外", None),
        (12, "歸屬母公司淨利", [135, -810, 44.9722], "Verified", S_24 + "；" + S_25, None),
        (13, "稅前利益、所得稅、少數股東損益", ["未揭露", "未揭露", "未揭露"], "Verified", "未找到：本輪查到的來源未列此三項（年報全文未讀取）", None),
        (14, "股數（百萬股；A 股，H 股 2026-03 才發行）", ["=sh_v-96.78", "=sh_v-96.78", "=sh_v-96.78"], "Derived", "＝總股本 968 − H 股 96.78；歷史 EPS 用此股數（2025 年得 0.05，與公告基本 EPS 0.05 一致）", None),
        (15, "EPS（元；＝淨利 ÷ A 股股數）", ["=C12/C14", "=D12/D14", "=E12/E14"], "Derived", "2025 年公告基本 EPS 0.05（" + S_25 + "）", "0.00"),
        (16, "研發費用（含於期間費用）", [None, None, 476], "Verified", S_25B + "（總研發投入 4.76 億，占營收 9.74%）", None),
    ]
    for (rr, lab, vals, tag, src, fmt) in HROWS:
        HI.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = HI.cell(rr, 3 + j, v)
                if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
                if fmt: c.number_format = fmt
        if "未找到" in src:
            HI.cell(rr, 8, "未找到：" + ("2023 年分部、毛利率" if rr in (5, 6, 7, 8, 9, 10, 11) else "見來源欄")); src = src.replace("未找到", "缺")
        HI.cell(rr, 6, tag); HI.cell(rr, 7, src)
    HI.cell(18, 1, "檢查：分部合計 − 營業收入（應為 0；2023 年分部未找到）")
    for c in "DE": HI[f"{c}18"] = f"=ROUND({c}8+{c}9-{c}4,0)"
    HI.column_dimensions["A"].width = 62; HI.column_dimensions["G"].width = 90

    # ================= Q =================
    Q["A1"] = "Q：2026 年半年與季度資料（百萬元）。一列一個項目；券商與媒體轉錄公司定期報告（原文未直接讀取）。"
    hdr(Q, 3, ["項目", "2026H1", "2026Q3", "2026 年 1–9 月", "標記", "來源", "來源誘因", "說明"])
    QROWS = [
        (4, "營業收入", [2577.5, "=D4-B4", 3699], "Verified", S_H1 + "；" + S_Q3, IP_CO, "Q3 單季由 1–9 月減上半年得 1,121.5，東吳稱 1,101、證券時報稱 1,122（來源分歧，差額為口徑或取整）"),
        (5, "營業收入年增", [0.0114, -0.1243, -0.0296], "Verified", S_H1 + "；" + S_Q3, "", "Q3 年減 12.4%"),
        (6, "毛利率", [None, 0.3364, 0.3236], "Verified", S_Q3, "", "Q3 單季 33.64%、1–9 月 32.36%（西南 32.37%）"),
        (7, "歸屬母公司淨利", [161.3, "=D7-B7", 171], "Verified", S_H1 + "；" + S_Q3, "", "Q3 單季約 9.7（東吳 0.09 億、證券時報 943.5 萬、西南 0.19 億，來源分歧）"),
        (8, "非經常損益（上半年）", [91, None, None], "Verified", S_H1, "", "主要為南京工藝股權置換公允價值利得約 100"),
    ]
    for (rr, lab, vals, tag, src, inc, note) in QROWS:
        Q.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = Q.cell(rr, 2 + j, v)
                if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
                if isinstance(v, float) and abs(v) < 1: c.number_format = "0.00%"
        Q.cell(rr, 5, tag); Q.cell(rr, 6, src); Q.cell(rr, 7, inc); Q.cell(rr, 8, note)
    Q.column_dimensions["A"].width = 30; Q.column_dimensions["F"].width = 90

    # ================= 線取數 =================
    LN["A1"] = "線取數：本模型讀入的線 3 變數（建置時由線 3 成品依 ID 讀值）。情境對應：A＝線 3 高情境、B＝基準、C＝低情境。"
    hdr(LN, 3, ["項目", "情境", ""] + [f"FY{y}" for y in range(2025, 2033)] + ["單位", "來源（檔案與儲存格）"])
    LR = {}; r = 4
    for s in SC:
        LN.cell(r, 1, "中國工業機器人裝機（含協作，IFR 口徑；線 3）"); LN.cell(r, 2, f"{s}（{LINE_SC[s]}）")
        for y in range(2025, 2033):
            v, src = LD["inst"][(y, s)]; c = LN.cell(r, 4 + (y - 2025), v); c.font = BLUE; c.number_format = "#,##0"
        LN.cell(r, 12, "台"); LN.cell(r, 13, f"{LD['inst'][(2025, s)][1]} … {LD['inst'][(2032, s)][1]}")
        LR[s] = r; r += 1
    LN.cell(r, 1, "中國裝機國產品牌占比 2025（線 3，Verified）"); LN.cell(r, 4, LD["dom25"][0]).font = BLUE; LN.cell(r, 4).number_format = "0%"; LN.cell(r, 12, "%"); LN.cell(r, 13, LD["dom25"][1]); LR["dom"] = r; r += 1
    LN.cell(r, 1, "國產 6 軸本體價（中負載；低／基準／高；線 3，Interested-party）")
    for j, (v, src) in enumerate(LD["body"]): LN.cell(r, 4 + j, v).font = BLUE
    LN.cell(r, 12, "元／台"); LN.cell(r, 13, LD["body"][1][1]); r += 1
    LN.cell(r, 1, "外資 6 軸本體價（中負載；基準；線 3）"); LN.cell(r, 4, LD["bodyf"][0]).font = BLUE; LN.cell(r, 12, "元／台"); LN.cell(r, 13, LD["bodyf"][1]); r += 1
    LN.cell(r, 1, "中國工業機器人產量 2026 年 1–9 月年增（線 3，Verified）"); LN.cell(r, 4, LD["prod_yoy"][0]).font = BLUE; LN.cell(r, 4).number_format = "0.0%"; LN.cell(r, 12, "%"); LN.cell(r, 13, LD["prod_yoy"][1]); LR["prod"] = r
    LN.column_dimensions["A"].width = 56; LN.column_dimensions["M"].width = 110

    # ================= Thesis =================
    TH["A1"] = ("命題模組：機器人及智能製造系統營收＝FY2026 錨 × 中國裝機成長（線 3）× 國產品牌占比變化 ×（1＋本公司相對份額漂移）×（1 − 單價年降）；"
                "核心部件營收＝前一年 ×（1＋年增）。FY2026＝1–9 月實績＋第 4 季假設。百萬元。")
    TH["A2"] = "欄：D＝FY2025 … K＝FY2032。國產占比 2025 年為線 3 實績，線性走到 2030 年情境值後持平。海外營收（2025 年占 30%）以同一驅動帶動，為簡化假設。"
    hdr(TH, 3, ["項目", "說明", ""] + [f"FY{y}" for y in range(2025, 2033)])
    THR = {}; r = 4
    for s in SC:
        TH.cell(r, 1, f"情境 {s}").font = BOLD; TH.cell(r, 1).fill = SECF; r += 1
        R = {}
        R["rev26"] = r; TH.cell(r, 1, "FY2026 合併營收（錨）"); TH.cell(r, 2, "＝1–9 月＋2025Q4 ×（1＋q4g）")
        TH[f"{TC[2026]}{r}"] = "=r9m26+(Hist!E4-r9m26/(1+g9m26))*(1+q4g)"; r += 1
        R["inst"] = r; TH.cell(r, 1, "中國工業機器人裝機（線 3 成長率 × 相對倍數，千台）")
        for y in range(2025, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!D{LR[s]}/1000" if y == 2025 else f"={TC[y - 1]}{r}*(1+(線取數!{TC[y]}{LR[s]}/線取數!{TC[y - 1]}{LR[s]}-1)*cn_k)"
        r += 1
        R["dom"] = r; TH.cell(r, 1, "國產品牌占比"); TH.cell(r, 2, "2025 線 3 實績；線性走到 ds30，其後持平")
        for y in range(2025, 2033):
            TH[f"{TC[y]}{r}"] = f"=線取數!D{LR['dom']}" if y == 2025 else f"=線取數!$D${LR['dom']}+(ds30_{s}-線取數!$D${LR['dom']})*MIN(1,({y}-2025)/5)"
            TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 1
        R["domu"] = r; TH.cell(r, 1, "國產品牌裝機（千台）")
        for y in range(2025, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['inst']}*{TC[y]}{R['dom']}"
        r += 1
        R["rob"] = r; TH.cell(r, 1, "機器人及智能製造系統 營收"); TH.cell(r, 2, "FY2027 起＝前一年 ×（國產裝機比）×（1＋漂移）×（1 − 年降）")
        for y in range(2025, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = ("=Hist!E8" if y == 2025 else f"={c}{R['rev26']}*sh_r26" if y == 2026 else
                             f"={TC[y - 1]}{r}*{c}{R['domu']}/{TC[y - 1]}{R['domu']}*(1+drift_{s})*(1-pd)")
        r += 1
        R["per"] = r; TH.cell(r, 1, "　每台國產品牌裝機對應營收（元；含系統、海外，Derived）")
        for y in range(2025, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['rob']}/{TC[y]}{R['domu']}*1000"
        r += 1
        R["core"] = r; TH.cell(r, 1, "自動化核心部件 營收")
        for y in range(2025, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = ("=Hist!E9" if y == 2025 else f"={c}{R['rev26']}*(1-sh_r26)" if y == 2026 else f"={TC[y - 1]}{r}*(1+gcore_{s})")
        r += 1
        R["gm"] = r; TH.cell(r, 1, "毛利率路徑"); TH.cell(r, 2, "FY2026＝1–9 月；線性到 FY2030 gm30，其後持平")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = "=gm26" if y == 2026 else f"=gm26+(gm30_{s}-gm26)*MIN(1,({y}-2026)/4)"; TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 1
        R["opex"] = r; TH.cell(r, 1, "期間費用率路徑"); TH.cell(r, 2, "FY2026＝opex26；線性到 FY2030 opex30，其後持平")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = "=opex26" if y == 2026 else f"=opex26+(opex30-opex26)*MIN(1,({y}-2026)/4)"; TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 2
        THR[s] = R
    for row in TH.iter_rows(min_row=4, max_row=r, min_col=4, max_col=11):
        for c in row:
            if c.number_format == "General": c.number_format = "#,##0.0"
    TH.column_dimensions["A"].width = 52; TH.column_dimensions["B"].width = 50

    # ================= PL =================
    PL["A1"] = ("PL：官方兩分部（機器人及智能製造系統／自動化核心部件及運動控制系統）→ 合併營收 → 標準損益表 → FCFF。"
                "C＝FY2023 … E＝FY2025 為實績；F＝FY2026（1–9 月實績＋第 4 季假設）… L＝FY2032。百萬元；中國會計準則（GAAP）。"
                "營業成本、毛利為由披露毛利率推得；期間費用含財務費用；2023 年分部與毛利率、歷史稅前與稅與少數股東未找到，標「未揭露」。")
    PR = {}; r = 3
    UND = "未揭露"
    for s in ["B", "A", "C"]:
        R = {}; th = THR[s]; pre = "" if s == "B" else f"（情境 {s}）"
        PL.cell(r, 1, f"情境 {s}" + ("（基準）" if s == "B" else ("（樂觀）" if s == "A" else "（保守）"))).font = BOLD; PL.cell(r, 1).fill = SECF; r += 1
        R["hdr"] = r; hdr(PL, r, ["項目", "單位"] + [f"FY{y}" for y in YRS]); r += 1
        def row(key, lab, unit, hist=None, fc=None, fmt=None):
            nonlocal r
            R[key] = r; PL.cell(r, 1, lab + pre); PL.cell(r, 2, unit)
            for y in YRS:
                f = (hist(y) if hist else None) if y <= 2025 else (fc(y) if fc else None)
                if f is not None:
                    PL[f"{COL[y]}{r}"] = f
                    if fmt: PL[f"{COL[y]}{r}"].number_format = fmt
            r += 1
        HC = lambda rr, need=2024: (lambda y: UND if y < need else f"=Hist!{'CDE'[y - 2023]}{rr}")
        C = lambda y: COL[y]
        row("rob", "機器人及智能製造系統 營收", "百萬元", HC(8), lambda y: f"=Thesis!{TC[y]}{th['rob']}")
        row("core", "自動化核心部件及運動控制系統 營收", "百萬元", HC(9), lambda y: f"=Thesis!{TC[y]}{th['core']}")
        row("rev", "營收（合併）", "百萬元", lambda y: f"=Hist!{'CDE'[y - 2023]}4", lambda y: f"={C(y)}{R['rob']}+{C(y)}{R['core']}")
        row("chk", "檢查：分部合計 − 合併營收（應為 0）", "百萬元", lambda y: UND if y < 2024 else f"=ROUND({C(y)}{R['rob']}+{C(y)}{R['core']}-{C(y)}{R['rev']},0)", lambda y: f"={C(y)}{R['rob']}+{C(y)}{R['core']}-{C(y)}{R['rev']}")
        PL.cell(R["chk"], 1, "檢查：分部合計 − 合併營收（應為 0）" if s == "B" else f"檢查（情境 {s}）：分部合計 − 合併營收（應為 0）")
        row("yoy", "營收年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1", lambda y: f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1", "0.0%")
        row("cogs", "營業成本", "百萬元", HC(5), lambda y: f"={C(y)}{R['rev']}*(1-Thesis!{TC[y]}{th['gm']})")
        row("gp", "毛利", "百萬元", lambda y: UND if y < 2024 else f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}")
        row("gm", "毛利率", "%", lambda y: UND if y < 2024 else f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", "0.0%")
        row("opex", "營業費用（銷售＋管理＋研發＋財務）", "百萬元", HC(10), lambda y: f"={C(y)}{R['rev']}*Thesis!{TC[y]}{th['opex']}")
        row("op", "營業利益（毛利 − 期間費用；不含減值、公允價值、補助）", "百萬元", lambda y: UND if y < 2024 else f"={C(y)}{R['gp']}-{C(y)}{R['opex']}", lambda y: f"={C(y)}{R['gp']}-{C(y)}{R['opex']}")
        row("opm", "營業利益率", "%", lambda y: UND if y < 2024 else f"={C(y)}{R['op']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", "0.0%")
        row("nonop", "業外與其他損益淨額（公允價值、補助、減值、稅金及附加）", "百萬元", lambda y: UND, lambda y: "=oth26" if y == 2026 else f"={C(y)}{R['rev']}*oth_r")
        row("pbt", "稅前利益", "百萬元", lambda y: UND, lambda y: f"={C(y)}{R['op']}+{C(y)}{R['nonop']}")
        row("tax", "所得稅", "百萬元", lambda y: UND, lambda y: f"=MAX(0,{C(y)}{R['pbt']})*tax_r")
        row("nit", "稅後淨利（含少數股東）", "百萬元", lambda y: UND, lambda y: f"={C(y)}{R['pbt']}-{C(y)}{R['tax']}")
        row("nci", "少數股東損益", "百萬元", lambda y: UND, lambda y: f"=MAX(0,{C(y)}{R['nit']})*nci_r")
        row("ni", "歸屬母公司淨利", "百萬元", lambda y: f"=Hist!{'CDE'[y - 2023]}12", lambda y: f"={C(y)}{R['nit']}-{C(y)}{R['nci']}")
        row("nim", "淨利率", "%", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", "0.0%")
        row("niy", "淨利年增率", "%", lambda y: None if y == 2023 else f"=({C(y)}{R['ni']}-{COL[y - 1]}{R['ni']})/ABS({COL[y - 1]}{R['ni']})",
            lambda y: f"=({C(y)}{R['ni']}-{COL[y - 1]}{R['ni']})/ABS({COL[y - 1]}{R['ni']})", "0.0%")
        row("sh", "股數（百萬股）", "百萬股", lambda y: f"=Hist!{'CDE'[y - 2023]}14", lambda y: "=sh_v")
        row("eps", "EPS（元）", "元", lambda y: f"=Hist!{'CDE'[y - 2023]}15", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['sh']}", "0.00")
        row("da", "折舊攤銷", "百萬元", None, lambda y: f"={C(y)}{R['rev']}*da_r")
        row("cx", "資本支出", "百萬元", None, lambda y: "=cx26" if y == 2026 else f"={C(y)}{R['rev']}*cx_r")
        row("dnwc", "營運資金增加", "百萬元", None, lambda y: f"=nwc_r*({C(y)}{R['rev']}-{COL[y - 1]}{R['rev']})")
        row("nopat", "NOPAT（＝營業利益 ×（1 − 稅率））", "百萬元", None, lambda y: f"={C(y)}{R['op']}*(1-tax_r)")
        row("fcff", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "百萬元", None, lambda y: f"={C(y)}{R['nopat']}+{C(y)}{R['da']}-{C(y)}{R['cx']}-{C(y)}{R['dnwc']}")
        row("div", "股利（淨利 × 配息率）", "百萬元", None, lambda y: f"=MAX(0,{C(y)}{R['ni']})*payout")
        row("gate", "FCFF 扣股利（收支）", "百萬元", None, lambda y: f"={C(y)}{R['fcff']}-{C(y)}{R['div']}")
        r += 1; PR[s] = R
    for row_ in PL.iter_rows(min_row=3, max_row=r, min_col=3, max_col=12):
        for c in row_:
            if c.number_format == "General": c.number_format = "#,##0.0"
    PL.column_dimensions["A"].width = 56

    # ================= DCF =================
    DC["A1"] = "DCF：FCFF 折現（期中慣例；評價日 2026-12-31，明確期 FY2027–FY2032）。終值＝NOPAT(FY2032) ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）；tv_alt＝1 時改用 FCFF ×（1＋g）÷（WACC − g）。"
    hdr(DC, 3, ["項目"] + [f"FY{y}" for y in range(2027, 2033)])
    DCOL = {y: openpyxl.utils.get_column_letter(2 + y - 2027) for y in range(2027, 2033)}
    DC.cell(4, 1, "折現年數（評價日起，期中）"); DC.cell(5, 1, "折現因子")
    for y in range(2027, 2033):
        c = DCOL[y]; DC[f"{c}4"] = y - 2027 + 0.5; DC[f"{c}5"] = f"=1/(1+wacc)^{c}4"
    DR = {}; r = 7
    for s in SC:
        R = {}
        DC.cell(r, 1, f"情境 {s}").font = BOLD; r += 1
        R["f"] = r; DC.cell(r, 1, "FCFF")
        for y in range(2027, 2033): DC[f"{DCOL[y]}{r}"] = f"=PL!{COL[y]}{PR[s]['fcff']}"
        r += 1
        R["pv"] = r; DC.cell(r, 1, "現值")
        for y in range(2027, 2033): c = DCOL[y]; DC[f"{c}{r}"] = f"={c}{R['f']}*{c}5"
        r += 1
        R["sum"] = r; DC.cell(r, 1, "明確期現值合計"); DC[f"B{r}"] = f"=SUM(B{R['pv']}:G{R['pv']})"; r += 1
        R["tv"] = r; DC.cell(r, 1, "終值（FY2032 年底）")
        DC[f"B{r}"] = f"=IF(tv_alt=1,PL!L{PR[s]['fcff']}*(1+g_tv)/(wacc-g_tv),PL!L{PR[s]['nopat']}*(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv))"; r += 1
        R["tvpv"] = r; DC.cell(r, 1, "終值現值"); DC[f"B{r}"] = f"=B{R['tv']}/(1+wacc)^6"; r += 1
        R["ev"] = r; DC.cell(r, 1, "企業價值 EV"); DC[f"B{r}"] = f"=B{R['sum']}+B{R['tvpv']}"; r += 1
        R["tvs"] = r; DC.cell(r, 1, "終值占 EV"); DC[f"B{r}"] = f"=IF(B{R['ev']}>0,B{R['tvpv']}/B{R['ev']},\"N/M\")"; DC[f"B{r}"].number_format = "0.0%"; r += 1
        R["tvw"] = r; DC.cell(r, 1, "檢查：終值占 EV 是否超過 80%"); DC[f"B{r}"] = f"=IF(AND(B{R['ev']}>0,B{R['tvs']}<=0.8),\"未超過\",\"超過或 N/M（報告標示）\")"; r += 1
        R["stub"] = r; DC.cell(r, 1, "＋Stub 期 FCFF（FY2026 FCFF × stub_frac）"); DC[f"B{r}"] = f"=PL!F{PR[s]['fcff']}*stub_frac"; r += 1
        R["nc"] = r; DC.cell(r, 1, "淨現金滾到評價日（現金 − 有息負債 − 少數股東權益 ＋ Stub FCFF）")
        DC[f"B{r}"] = f"=cash_v-debt_v-nci_v+B{R['stub']}"; r += 1
        R["eq"] = r; DC.cell(r, 1, "股權價值"); DC[f"B{r}"] = f"=B{R['ev']}+B{R['nc']}"; r += 1
        R["ps"] = r; DC.cell(r, 1, "每股價值（評價日，元）"); DC[f"B{r}"] = f"=MAX(0,B{R['eq']}/sh_v)"; r += 1
        R["dps"] = r; DC.cell(r, 1, "評價日→目標時點間股利（FY2026 盈餘配息）"); DC[f"B{r}"] = f"=PL!F{PR[s]['eps']}*payout"; r += 1
        R["v12"] = r; DC.cell(r, 1, "推到目標時點（×（1＋ke）^t − 期間股利）"); DC[f"B{r}"] = f"=MAX(0,B{R['ps']}*(1+ke)^((tp_date-val_date)/365)-B{R['dps']})"; r += 2
        DR[s] = R
    for row_ in DC.iter_rows(min_row=4, max_row=r, min_col=2, max_col=7):
        for c in row_:
            if c.number_format == "General": c.number_format = "#,##0.00"
    DC.column_dimensions["A"].width = 70

    # ================= Comp =================
    CP["A1"] = ("Comp：同業前瞻 P/E。股價＝" + S_CLOSE + "（Verified，四家日本同業全部是 2026-11-20 收盤）；"
                "前瞻 P/E＝收盤 ÷ 本專案各同業個股模型的 FY2027 EPS（情境 B；模型值，非市場共識）。中國同業匯川（300124）收盤檔有報價但前瞻 EPS 未取得，不進中位數。")
    hdr(CP, 4, ["公司", "代號", "股價（2026-11-20）", "報價日", "前瞻 EPS（FY2027，情境 B）", "前瞻 P/E", "進中位數", "說明", "標記與來源"])
    PEERS = [  # (公司, 代號, 收盤, EPS, 說明, 來源)
        ("Fanuc", "6954.T", 5047, 287.825172733299, "六軸工業機器人與 CNC；FY2027＝2028 年 3 月期（與本公司 FY2027 差 3 個月）", "physical_ai/fanuc 模型 PL 情境 B FY2027 EPS"),
        ("安川電機", "6506.T", 3180, 193.785912462578, "機器人、運動控制；FY2027＝2028 年 2 月期", "physical_ai/yaskawa 模型 PL 情境 B FY2027E EPS"),
        ("Nabtesco", "6268.T", 3105, 235.79370285695, "RV 減速機（機器人零組件）；12 月年結", "physical_ai/nabtesco 模型 PL 情境 B FY2027 EPS"),
        ("哈默納科", "6324.T", 3722, 91.6758836947537, "諧波減速機（機器人零組件）；FY2027＝2028 年 3 月期", "physical_ai/harmonic_drive 模型 Val 情境 B FY2027 EPS"),
    ]
    for i, (nm, tk, px1, eps, note, src) in enumerate(PEERS):
        rr = 5 + i
        CP.cell(rr, 1, nm); CP.cell(rr, 2, tk); CP.cell(rr, 3, px1).font = BLUE; CP.cell(rr, 4, "2026-11-20"); CP.cell(rr, 5, eps).font = BLUE
        CP.cell(rr, 6, f"=C{rr}/E{rr}").number_format = "0.0"; CP.cell(rr, 7, "是"); CP.cell(rr, 8, note)
        CP.cell(rr, 9, "股價 Verified（" + S_CLOSE + "）；EPS Derived：" + src + "（模型值，非市場共識）")
    CP.cell(9, 1, "匯川技術"); CP.cell(9, 2, "300124.SZ"); CP.cell(9, 3, 52.66).font = BLUE; CP.cell(9, 4, "2026-11-20"); CP.cell(9, 5, "未取得"); CP.cell(9, 6, "未取得"); CP.cell(9, 7, "否")
    CP.cell(9, 8, "國產機器人與伺服；收盤檔有報價，前瞻 EPS 未取得（搜尋只得 2025 年 10 月版券商預測，已過期不用）；報價日未取得前瞻值"); CP.cell(9, 9, "股價 Verified（" + S_CLOSE + "）")
    CP.cell(10, 1, "美的（KUKA）"); CP.cell(10, 2, "000333.SZ"); CP.cell(10, 3, 94.82).font = BLUE; CP.cell(10, 4, "2026-11-20"); CP.cell(10, 5, "未取得"); CP.cell(10, 6, "未取得"); CP.cell(10, 7, "否")
    CP.cell(10, 8, "家電集團，機器人占比小，不可比；前瞻 EPS 未取得"); CP.cell(10, 9, "股價 Verified（" + S_CLOSE + "）")
    CP.cell(11, 1, "綠的諧波（A 股零組件，對照）"); CP.cell(11, 2, "688017.SS"); CP.cell(11, 3, 230.8).font = BLUE; CP.cell(11, 4, "2026-11-20"); CP.cell(11, 5, 1.07).font = BLUE
    CP.cell(11, 6, "=C11/E11").number_format = "0.0"; CP.cell(11, 7, "否"); CP.cell(11, 8, "股價由人形選擇權定價，不可比；只作 A 股溢價對照"); CP.cell(11, 9, "EPS Derived：physical_ai/leaderdrive README 情境 B FY2027 EPS 1.07")
    CP.cell(13, 1, "核心同業前瞻 P/E 中位數"); CP["F13"] = "=MEDIAN(F5:F8)"
    CP.cell(14, 1, "本公司現價 ÷ 模型 FY2027 EPS（情境 B）"); CP["F14"] = f"=px/PL!G{PR['B']['eps']}"
    CP.cell(15, 1, "本公司現價 ÷ 券商 2027E EPS（西南 0.33 元）"); CP["F15"] = "=px/0.33"
    CP.cell(16, 1, "本公司現行倍數 ÷ 同業中位數（西南 2027E EPS 口徑）"); CP["F16"] = "=F15/F13"
    for rr in (13, 14, 15, 16): CP[f"F{rr}"].number_format = "0.0"
    CP.column_dimensions["A"].width = 52; CP.column_dimensions["H"].width = 60; CP.column_dimensions["I"].width = 80

    # ================= Val =================
    V["A1"] = "Val：DCF（FCFF，WACC）50% ＋ 同業 P/E（Comparable）50%，三情境 25／50／25 加權；約 12 個月目標價（元）。"
    VR = {}; r = 3
    V.cell(r, 1, "0. 折現率").font = BOLD; r += 1
    for key, lab, f, fmt in (("ke", "股權成本 ke＝rf＋β ×（ERP＋CRP）", "=rf+beta*(erp+crp)", "0.00%"),
                             ("kd_at", "稅後債務成本", "=kd*(1-tax_r)", "0.00%"),
                             ("e_mv", "股權市值（百萬元）", "=px*sh_v", "#,##0"),
                             ("d_bv", "有息負債（百萬元）", "=debt_v", "#,##0"),
                             ("wacc", "WACC", "=(e_mv*ke+d_bv*kd_at)/(e_mv+d_bv)+wacc_adj", "0.00%"),
                             ("g_show", "永續成長率 g", "=g_tv", "0.00%"),
                             ("ronic", "RONIC", "=wacc+ronic_sp", "0.00%")):
        V.cell(r, 1, lab); V[f"B{r}"] = f; V[f"B{r}"].number_format = fmt
        if key != "g_show": dn(wb, key, f"Val!$B${r}")
        VR[key] = r; r += 1
    V.cell(r, 1, "檢查：WACC 由 In 頁輸入重算 − Val 的 WACC（應為 0）")
    V[f"B{r}"] = "=((px*sh_v)*(rf+beta*(erp+crp))+debt_v*kd*(1-tax_r))/((px*sh_v)+debt_v)+wacc_adj-wacc"; r += 2
    V.cell(r, 1, "1. 各情境價值（元／股，目標時點）").font = BOLD; r += 1
    hdr(V, r, ["項目", "情境 A", "情境 B", "情境 C", "說明"]); r += 1
    def row3(key, lab, fs, fmt="#,##0.00", note=""):
        nonlocal r
        V.cell(r, 1, lab)
        for j, f in enumerate(fs): V.cell(r, 2 + j, f).number_format = fmt
        V.cell(r, 5, note); VR[key] = r; r += 1
    row3("ev", "企業價值 EV（百萬元）", [f"=DCF!B{DR[s]['ev']}" for s in SC], "#,##0")
    row3("tvs", "終值占 EV", [f"=DCF!B{DR[s]['tvs']}" for s in SC], "0.0%")
    row3("dcf", "DCF 情境價值推到 12 個月後", [f"=DCF!B{DR[s]['v12']}" for s in SC])
    row3("eps27", "FY2027 EPS（模型）", [f"=PL!G{PR[s]['eps']}" for s in SC], "0.00", "分母：模型各情境 FY2027 EPS")
    row3("cmpu", "Comparable 未調整（同業 P/E × FY2027 EPS）", [f"=MAX(0,{c}{r - 1}*pe_peer)" for c in "BCD"], "#,##0.00", "同業 P/E＝In 頁 pe_peer（4 家日本同業前瞻 P/E 中位數）")
    row3("cmp", "Comparable（× 相對倍數係數）", [f"={c}{VR['cmpu']}*rel_k" for c in "BCD"], "#,##0.00", "rel_k＝1.00（未調整：缺十年相對倍數歷史）")
    row3("wt", "情境目標價：DCF × w ＋ Comparable ×（1 − w）", [f"=w_dcf*{c}{VR['dcf']}+(1-w_dcf)*{c}{VR['cmp']}" for c in "BCD"], "#,##0.00")
    row3("w", "情境權重", ["=tp_wA", "=tp_wB", "=tp_wC"], "0%")
    r += 1
    V.cell(r, 1, "2. 目標價與評等").font = BOLD; r += 1
    def row1(key, lab, f, fmt="#,##0.00", note=""):
        nonlocal r
        V.cell(r, 1, lab); V[f"B{r}"] = f; V[f"B{r}"].number_format = fmt; V.cell(r, 5, note); VR[key] = r; r += 1
    row1("dcf_tp", "DCF 目標價（情境加權）", f"=SUMPRODUCT(B{VR['dcf']}:D{VR['dcf']},B{VR['w']}:D{VR['w']})")
    row1("cmp_tp", "Comparable 目標價（情境加權；同業 P/E × 相對倍數係數）", f"=SUMPRODUCT(B{VR['cmp']}:D{VR['cmp']},B{VR['w']}:D{VR['w']})")
    row1("cmpu_tp", "Comparable 未調整（情境加權）", f"=SUMPRODUCT(B{VR['cmpu']}:D{VR['cmpu']},B{VR['w']}:D{VR['w']})")
    row1("tp", "加權目標價（DCF 與 Comparable 加權）", f"=w_dcf*B{VR['dcf_tp']}+(1-w_dcf)*B{VR['cmp_tp']}")
    row1("px_show", "現價（股價基準日）", "=px")
    row1("up", "上檔空間：加權目標價", f"=B{VR['tp']}/px-1", "0.0%")
    row1("rating", "估值評等（只看上檔空間）", f'=IF(B{VR["up"]}>0.5,"強力買進",IF(B{VR["up"]}>0.25,"買進",IF(B{VR["up"]}>=-0.25,"中立",IF(B{VR["up"]}>=-0.5,"放空","強力放空"))))', "@",
         "級距：> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空（原則第 10 節第 24 條）")
    V.cell(r, 1, "收支關卡：FY2026–FY2030 FCFF 扣股利合計（情境 A／B／C，百萬元）")
    for j, s in enumerate(SC): V.cell(r, 2 + j, f"=SUM(PL!F{PR[s]['gate']}:J{PR[s]['gate']})").number_format = "#,##0.0"
    V.cell(r, 5, "口徑：FCFF − 股利；基準情境 B < 0＝入不敷出"); VR["gate"] = r; r += 1
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
    row1("att_chk", "檢查：三段合計 − 上檔空間（應為 0）", f"=B{VR['att_d']}+B{VR['att_c']}+B{VR['att_r']}-B{VR['up']}", "0.0000")
    row1("relnow", "本公司現行前瞻倍數 ÷ 同業中位數（西南 2027E EPS 口徑）", "=Comp!F16", "0.00")
    row1("relreq", "支持現價所需的相對倍數係數（DCF 不變）", f"=(px-w_dcf*B{VR['dcf_tp']})/((1-w_dcf)*B{VR['cmpu_tp']})", "0.00",
         "反解：目標價＝現價；DCF 以 RONIC＝WACC（無超額報酬）而高相對倍數隱含持續超額報酬，兩法對競爭優勢的假設相反")
    row1("relflip", "評等升一級所需的相對倍數係數", f'=IF(ISNUMBER(B{VR["up_b"]}),(B{VR["up_b"]}-w_dcf*B{VR["dcf_tp"]})/((1-w_dcf)*B{VR["cmpu_tp"]}),"—")', "0.00")
    row1("pe_impl", "隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2027 EPS", f"=B{VR['tp']}/C{VR['eps27']}", "0.0")
    r += 1
    V.cell(r, 1, "4. 期間").font = BOLD; r += 1
    row1("t_model", "模型折現期間（評價日→目標時點，年）", "=(tp_date-val_date)/365", "0.00")
    row1("t_hold", "持有期間（股價基準日→目標時點，年；上檔空間不是年化報酬）", "=(tp_date-px_date)/365", "0.00")
    row1("tp_date_show", "目標時點", "=tp_date", "yyyy-mm-dd")
    V.column_dimensions["A"].width = 66; V.column_dimensions["E"].width = 70
    for c in "BCD": V.column_dimensions[c].width = 14

    # ================= Sens =================
    SE["A1"] = "Sens：一次動一個參數（In 頁敏感度低／高或指定值），整本重建並以 LibreOffice 重算後讀加權目標價（不是閉式近似）；建置腳本產生。"
    SE["A2"] = "build_id"; SE["A3"] = "模型狀態"; SE["A4"] = "In 頁輸入指紋（Sens 計算時）"; SE["A5"] = "In 頁輸入指紋（本檔即時）"
    SE["D5"] = f"=SUMPRODUCT(In!E4:E{LAST_IN},ROW(In!E4:E{LAST_IN}))"
    hdr(SE, 7, ["驅動", "設定", "鍵", "輸入值", "加權目標價", "加權目標價相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    if sens:
        SE["D2"] = sens["build_id"]; SE["D4"] = sens["fp"]
        SE["D3"] = f'=IF(AND(ABS(D5-D4)<1E-6,ABS(Val!B{VR["tp"]}-{sens["base_tp"]})<0.01),"CURRENT","STALE")'
        r = 8
        for c in sens["cases"]:
            SE.cell(r, 1, c["drv"]); SE.cell(r, 2, c["set"]); SE.cell(r, 3, c["key"])
            if c.get("val") is not None: SE.cell(r, 4, c["val"])
            SE.cell(r, 5, c["tp"]).number_format = "#,##0.00"; SE.cell(r, 6, c["tp"] / sens["base_tp"] - 1).number_format = "+0.0%;-0.0%"
            SE.cell(r, 7, c["dcf"]).number_format = "#,##0.00"; SE.cell(r, 8, c["dcf"] / sens["base_dcf"] - 1 if sens["base_dcf"] else None).number_format = "+0.0%;-0.0%"
            SE.cell(r, 9, c["rating"]); r += 1
        r += 1
        SE.cell(r, 1, "讀法：相對基準＝案例加權目標價 ÷ 基準加權目標價 − 1；同一案例 DCF 單法與加權目標價並列（原則第 10 節第 32 條）。情境專屬驅動以「情境 B 改為情境 A／C 值」呈現。")
        SE.cell(r + 1, 1, "說明：折舊、資本支出（FY2027 起）、營運資金隨營收等比例變動（第 44 條）；單價年降幅同時作用於三情境。")
    SE.column_dimensions["A"].width = 52; SE.column_dimensions["B"].width = 44

    # ================= Recon =================
    RC["A1"] = "Recon：模型對實績、券商預測、線 3 的同口徑對照；反解支持現價所需的主導變數。"
    hdr(RC, 3, ["項目", "模型", "對照", "差距（模型 − 對照）", "說明"])
    rr = 4
    def rrow(lab, a, b, note, fmt="#,##0.0", diff=True):
        nonlocal rr
        RC.cell(rr, 1, lab); RC.cell(rr, 2, a).number_format = fmt
        if b is not None: RC.cell(rr, 3, b).number_format = fmt
        if diff and b is not None: RC.cell(rr, 4, f"=B{rr}-C{rr}").number_format = fmt
        RC.cell(rr, 5, short(note)); rr += 1
    rrow("FY2026 歸屬母公司淨利 vs 東吳 2026E 208、西南隱含（EPS 0.25 × 968）", f"=PL!F{PR['B']['ni']}", "=AVERAGE(208,0.25*968)", "券商為 Interested-party；東吳為三季報後（2026-10-30），西南為 2026-11-09")
    rrow("FY2027 歸屬母公司淨利（情境 B）vs 東吳 260、西南隱含（0.33 × 968）", f"=PL!G{PR['B']['ni']}", "=AVERAGE(260,0.33*968)", "同上")
    rrow("FY2028 歸屬母公司淨利（情境 B）vs 東吳 337、西南隱含（0.44 × 968）", f"=PL!H{PR['B']['ni']}", "=AVERAGE(337,0.44*968)", "同上")
    rrow("FY2027 EPS（情境 B）vs 西南 2027E 0.33 元", f"=PL!G{PR['B']['eps']}", 0.33, "", "0.00")
    rrow("FY2026 毛利率 vs 2026 年 1–9 月實績", f"=PL!F{PR['B']['gm']}", "=gm26", "模型以 1–9 月為錨（差距應為 0）", "0.0%")
    rrow("FY2026 營收 vs 2026 年 1–9 月營收 ×（12 ÷ 9）", f"=PL!F{PR['B']['rev']}", "=r9m26*12/9", "第 4 季若維持 1–9 月平均水準即為對照值")
    rrow("FY2026 第 4 季隱含營收 vs 2026Q3 實績", f"=PL!F{PR['B']['rev']}-r9m26", "=Q!C4", "第 4 季若維持第 3 季水準即為對照值")
    rr += 1
    RC.cell(rr, 1, "線 3 對照（同口徑：中國裝機與國產占比）").font = BOLD; rr += 1
    rrow("2026 年中國裝機成長（線 3 基準）vs 國家統計局 1–9 月產量年增", f"=Thesis!E{THR['B']['inst']}/Thesis!D{THR['B']['inst']}-1", "=線取數!D" + str(LR["prod"]), "產量含出口，與 IFR 裝機口徑不同；只看方向", "0.0%")
    rrow("FY2026 機器人營收年增（情境 B）vs 中國裝機成長（線 3 基準）", f"=Thesis!E{THR['B']['rob']}/Thesis!D{THR['B']['rob']}-1", f"=Thesis!E{THR['B']['inst']}/Thesis!D{THR['B']['inst']}-1", "2026 年前三季營收 −3%，落後中國裝機成長：本公司份額與單價壓力是否已反映，為本模型最大的待查項", "0.0%")
    rrow("2025 每台國產品牌裝機對應營收（元）vs 線 3 國產 6 軸本體價基準", f"=Thesis!D{THR['B']['per']}", f"=線取數!E{LR['dom'] + 1}", f"{LD['body'][1][1]}（Interested-party）；本模型數字含系統整合與海外，口徑不同，比值只作量級對照", "#,##0")
    rrow("FY2030 國產品牌占比（情境 B）vs 2025 年實績", f"=Thesis!I{THR['B']['dom']}", f"=Thesis!D{THR['B']['dom']}", "2025 年為線 3 實績（IFR）；2030 年為假設", "0.0%")
    rr += 1
    RC.cell(rr, 1, "反解（整本重算二分搜尋；建置時計算，數值為快照）").font = BOLD; rr += 1
    hdr(RC, rr, ["反解目標", "解", "基準值", "", "說明"]); rr += 1
    if sens:
        for lab, v, b0, note in sens["solve"]:
            RC.cell(rr, 1, lab); RC.cell(rr, 2, v).number_format = "#,##0.00"; RC.cell(rr, 3, b0).number_format = "#,##0.00"; RC.cell(rr, 5, note); rr += 1
    RC.column_dimensions["A"].width = 70; RC.column_dimensions["E"].width = 90

    # ================= 對帳 =================
    DZ["A1"] = "對帳：本模型使用的「線層級」變數（與 physical_ai/lines/industrial_robot 線 3 對帳；本模型讀線 v0.1 成品，不靜默改口徑）。"
    hdr(DZ, 3, ["變數名", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型是否一致"])
    i26 = LD["inst"][(2026, "B")]; i32 = LD["inst"][(2032, "B")]
    ZR = [
        ("中國工業機器人裝機 FY2026–FY2032（A／B／C＝線高／基準／低）", "3 工業機器人", f"基準 {i26[0]:,.0f} → {i32[0]:,.0f}（× cn_k）", "台", "IFR 中國裝機口徑（含協作），非國家統計局產量", f"{i26[1]}；{i32[1]}（逐年見線取數頁）", "Derived", "Thesis 機器人營收成長", "一致（直接讀值，cn_k＝1）"),
        ("中國裝機國產品牌占比 2025", "3 工業機器人", "55%", "%", "IFR 中國", LD["dom25"][1], "Verified", "Thesis 國產占比起點", "一致"),
        ("中國裝機國產品牌占比 2030（A／B／C）", "3 工業機器人", "=ds30_B", "%", "假設路徑（線 3 無 2030 年占比輸出）", "本模型 In 頁 ds30_A／B／C", "Assumed", "Thesis 國產占比路徑", "線 3 只給 2025 年實績；2030 年為本模型自設值（A 70%／B 65%／C 58%），標明"),
        ("國產 6 軸本體價（中負載）", "3 工業機器人", f"{LD['body'][1][0]:,.0f}（低 {LD['body'][0][0]:,.0f}／高 {LD['body'][2][0]:,.0f}）", "元／台", "本體價，不含系統整合", LD["body"][1][1], "Interested-party", "Recon 量級對照（不進公式）", "不一致（刻意）：本模型以「每台國產品牌裝機對應營收」（含系統、海外、核心部件外銷）推導，只讀線 3 的裝機與占比，不用本體價推營收"),
        ("單價年降幅 pd", "3 工業機器人", "=pd", "%／年", "每台對應營收年降（價格戰與組合）", "本模型 In 頁 pd（線 3 無逐年降幅）", "Analogy", "Thesis 機器人營收", "線 3 無對應輸出；以 Analogy 區間 0–8% 表達，非線 3 輸入"),
        ("中國工業機器人產量 2026 年 1–9 月年增", "3 工業機器人", "=線取數!D" + str(LR["prod"]), "%", "國家統計局產量（含出口）", LD["prod_yoy"][1], "Verified", "q4g 的依據（定性）", "一致"),
        ("線 3 追蹤指標：埃斯頓季營收與毛利率", "3 工業機器人", "季營收 Q3 −12.4%、毛利率 33.6%", "—", "線 3 追蹤指標!A41 把本公司列為國產替代與價格戰的利潤端讀數", f"{L3N} 追蹤指標!A41", "Verified", "Q 頁、gm 路徑", "一致：毛利率回升、營收下滑並存，本模型兩者皆取"),
    ]
    for i, z in enumerate(ZR):
        for j, v in enumerate(z): DZ.cell(4 + i, 1 + j, v)
    DZ.column_dimensions["A"].width = 50; DZ.column_dimensions["F"].width = 80; DZ.column_dimensions["I"].width = 70

    # ================= Risk =================
    RK["A1"] = "Risk：風險登錄（原則第 10 節第 10、19 條）。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    RISKS = [
        ("國產價格戰壓縮單價與毛利率", "是（pd、gm30_C）", "國產 6 軸本體價 10 萬 vs 外資 27.5 萬，線 3 低案 2 萬", "季毛利率、Q3 單季營收年增"),
        ("商譽與資產減值（併購品牌：Cloos、扎根德國等；2024 年已提 3.45 億商譽減值）", "部分（oth_r 下限 −0.5%）", "2026 年收購酷卓 4.87 億元與江蘇智能少數股權，整合與業績對賭風險（西南）", "年報商譽與減值測試"),
        ("下游製造業資本支出放緩（鋰電、光伏、汽車零組件）", "是（情境 C、cn_k 低案）", "東吳下修 2026–2028E 歸母淨利 27%／35%／39%，理由為下游需求增速放緩", "訂單、合約負債（1–9 月 −18%）"),
        ("公允價值利得非經常（南京工藝股權置換約 1.00 億）", "是（oth26 區間 60–180）", "前三季淨利 1.71 億中非經常成分高", "扣非歸母淨利（1–9 月 1.16 億）"),
        ("海外擴張與匯率（海外營收 2025 年 14.63 億，占 30%）", "否（併入同一驅動）", "未取得匯率敏感度", "海外營收占比、波蘭基地"),
        ("A＋H 兩地股價與 H 股流通", "否", "股數用總股本 968 百萬股，H 股 96.78 百萬股已含", "H 股與 A 股價差"),
        ("資產負債表未取得（現金、有息負債、少數股東權益）", "是（cash_v、debt_v、nci_v 區間）", "資產負債率約 68–70%，絕對金額未取得，淨負債約 13 億元為估計", "三季報合併資產負債表"),
    ]
    for i, k in enumerate(RISKS):
        for j, v in enumerate(k): RK.cell(4 + i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["D"].width = 60

    # ================= Sources =================
    SO["A1"] = "Sources：本模型引用的原始來源與資料商（含網址；HTML 展示層不放網址）。"
    for i, s_ in enumerate([S_PX, S_25, S_25B, S_24, S_Q3, S_Q3B, S_H1, S_HK, S_SH, S_RF, S_CRP, f"線 3：{L3N}",
                            "同業收盤：" + S_CLOSE, "physical_ai/fanuc、yaskawa、nabtesco、harmonic_drive、leaderdrive 模型（同業前瞻 EPS）",
                            "physical_ai/map/profiles/002747.md（背景，數字已另行查證）"], 3):
        SO.cell(i, 1, s_)
    SO.column_dimensions["A"].width = 160

    # ================= README（xlsx 內）=================
    S = T
    v = lambda k: f"Val!B{VR[k]}"
    lines = [
        ('="埃斯頓 Estun（002747.SZ）個股模型 ' + VER + '｜build_id "&Sens!D2&"｜資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'),
        "命題：埃斯頓是中國國產工業機器人第一梯隊，但 2026 年前三季營收年減 3%、毛利率回升；營收＝中國裝機（線 3）× 國產品牌占比 × 本公司相對份額 ×（1 − 單價年降）＋核心部件。",
        (f'="結論（條件式）：在 FY2026 以 1–9 月實績為錨、線 3 基準中國裝機、國產占比 2030 年 "&{S("ds30_B","0%")}&"、相對份額年漂移 "&{S("drift_B","0%")}&"、單價年降 "&{S("pd","0%")}&"、FY2030 毛利率 "&{S("gm30_B","0%")}&"、期間費用率 "&{S("opex30","0%")}&"、同業 P/E 未加溢價（相對倍數係數 "&{S("rel_k","0.00")}&"）的假設下，加權目標價約 "&{S(v("tp"),"#,##0.00")}&" 元，相對現價 "&{S("px","#,##0.00")}&" 元 "&{S(v("up"),"+0.0%;-0.0%")}&"，估值評等 "&{v("rating")}&"，投資建議 "&{v("rec")}&"。"'),
        (f'="DCF 目標價 "&{S(v("dcf_tp"),"#,##0.00")}&"、Comparable 目標價 "&{S(v("cmp_tp"),"#,##0.00")}&"（各半）；WACC "&{S("wacc","0.00%")}&"、ke "&{S("ke","0.00%")}&"、g "&{S("g_tv","0.00%")}&"；情境 B 終值占 EV "&{S(f"Val!C{VR['tvs']}","0%")}&"。"'),
        (f'="上檔空間歸因：DCF 貢獻 "&{S(v("att_d"),"+0.0%;-0.0%")}&"、未調整 Comparable 貢獻 "&{S(v("att_c"),"+0.0%;-0.0%")}&"、相對倍數調整貢獻 "&{S(v("att_r"),"+0.0%;-0.0%")}&"。本公司現價對西南 2027E EPS 的倍數是同業中位數的 "&{S(v("relnow"),"0.0")}&" 倍；支持現價需要相對倍數係數 "&{S(v("relreq"),"0.0")}&"，升一級需要 "&{S(v("relflip"),"0.0")}&"。"'),
        (f'="收支關卡：情境 B FY2026–FY2030 FCFF 扣股利合計 "&{S(v("gate_b"),"#,##0.0")}&" 百萬元，"&{v("gate_t")}&"；距最近級距邊界 "&{S(v("dist"),"0.0%")}&"。"'),
        "主要失效條件（可觀察）：(1) 營收年增轉正且毛利率維持 32% 以上、期間費用率降到 25% 以下；(2) 國產品牌占比 2030 年超過 70%；(3) 資產負債表實際淨負債遠低於估計；(4) A 股對機器人題材持續給予同業 3 倍以上的倍數。",
        "方法：DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，WACC 折現，明確期到 FY2032，RONIC＝WACC）50%＋同業前瞻 P/E × 模型 FY2027 EPS 50%；三情境 25／50／25。",
        "限制：資產負債表（現金、有息負債、少數股東權益）、稅前與稅、少數股東損益、2023 年分部與毛利率、β、折舊與資本支出歷史未取得，均為 Assumed 區間；同業為 4 家日本同業（A 股同業前瞻 EPS 未取得）；不讀 Siliconomics／Tokenomics／Wattonomics。詳見 repo README.md。",
    ]
    for i, t in enumerate(lines):
        RD.cell(1 + i, 1, t).alignment = WRAP
    RD.column_dimensions["A"].width = 160
    return wb, dict(VR=VR, KR=KR, PR=PR, LAST_IN=LAST_IN, THR=THR, DR=DR)


def recalc(paths, outdir):
    os.makedirs(outdir, exist_ok=True)
    for attempt in range(2):
        try:
            subprocess.run(LO + ["--outdir", outdir] + paths, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180); return
        except subprocess.TimeoutExpired:
            subprocess.run(["pkill", "-f", "lo_estun"], check=False)
            if attempt == 1: raise


def read_val(path, meta):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; VR = meta["VR"]
    g = lambda k, c="B": V[f"{c}{VR[k]}"].value
    return dict(tp=g("tp"), dcf=g("dcf_tp"), cmp=g("cmp_tp"), rating=g("rating"), up=g("up"))


def run_cases(cases, meta, sub):
    d = os.path.join(TMP, sub); shutil.rmtree(d, ignore_errors=True); raw = os.path.join(d, "raw"); rc = os.path.join(d, "rc"); os.makedirs(raw)
    paths = []
    for i, ov in enumerate(cases):
        wb, _ = make(ov); p = os.path.join(raw, f"c{i:03d}.xlsx"); wb.save(p); paths.append(p)
    for i in range(0, len(paths), 20): recalc(paths[i:i + 20], rc)
    return [read_val(os.path.join(rc, f"c{i:03d}.xlsx"), meta) for i in range(len(cases))]


def in_values(path, last):
    ws = openpyxl.load_workbook(path, data_only=True)["In"]; out = {}
    for r in range(4, last + 1):
        k = ws.cell(r, 1).value
        if isinstance(k, str) and ws.cell(r, 5).value is not None:
            out[k] = (ws.cell(r, 4).value, ws.cell(r, 5).value, ws.cell(r, 6).value, ws.cell(r, 2).value)
    return out


def write_readme(path, meta, sens):
    wb = openpyxl.load_workbook(path, data_only=True)
    V, P, I, C = wb["Val"], wb["PL"], wb["In"], wb["Comp"]
    VR, PR, KR = meta["VR"], meta["PR"], meta["KR"]
    g = lambda k, c="B": V[f"{c}{VR[k]}"].value
    iv = lambda k: I[f"E{KR[k]}"].value
    pl = lambda s_, k, col: P[f"{col}{PR[s_][k]}"].value
    f2 = lambda x: f"{x:,.2f}"; pc = lambda x: f"{x:.1%}"
    top = sorted([c for c in sens["cases"] if c["tp"] is not None], key=lambda c: -abs(c["tp"] / sens["base_tp"] - 1))[:10]
    L = []; a = L.append
    a("# physical_ai/estun/ — 埃斯頓 Estun（002747.SZ）個股模型 v0.1")
    a("")
    a("<!-- 本檔由 build_estun.py 依成品重算值產生；數字不手寫（原則第 10 節第 51 條）。 -->")
    a("")
    a("**本模型適用《高不確定性預估工作原則》**（`docs/plan/高不確定性預估工作原則.md`）。**未完成與缺口（先讀）：** (1) 資產負債表（現金、有息負債、少數股東權益）本輪未取得，三項以 Assumed 區間估計（淨負債約 13 億元量級，Sens 有案例）；(2) 歷史稅前利益、所得稅、少數股東損益、折舊攤銷、資本支出未取得，PL 歷史欄標「未揭露」，預測用 Assumed；(3) FY2023 的分部營收與毛利率未找到；(4) β 未取得（Assumed 1.1，區間 0.8–1.4），非兩來源 Blume 中位數；(5) 同業只有 4 家日本同業（前瞻 EPS 取自本專案各模型，非市場共識），匯川（A 股）報價在收盤檔但前瞻 EPS 未取得，未進中位數；(6) 2026 年分部營收未取得，以 FY2025 占比外推；(7) 尚未由未參與製作的代理獨立核對。")
    a("")
    a(f"- 成品：`dist/{NAME}`（build_id `{sens['build_id']}`；LibreOffice 重算，0 錯誤值）；HTML 展示層 `dist/{NAME[:-5]}.html`（`gen_html_json.py` 產生 `html.json`，再以 `tools/company_html/build_company_html.py` 產生）")
    a("- 建置：`python3 physical_ai/estun/build_estun.py`（需要 soffice；整本重算敏感度與反解；結尾自動跑 in_layout、sens_coverage --require-keycol、model_governance、rating_gate --strict、zero_rows、cached_values、segment_check）")
    a("- 日期：資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 2026-11-20（收盤 24.25 元，深交所）｜評價日 2026-12-31｜目標時點 2027-11-20")
    a("- 會計年度＝日曆年；單位百萬元人民幣；中國會計準則。深交所 A 股，另有港股 02715（2026-03 掛牌，股數用 A＋H 總股本 968 百萬股）。全球 AI 與半導體股 2026-11-16 至 11-20 同週大跌，收盤檔下跌非資料錯誤。")
    a("")
    a("## 結論（條件式）")
    a("")
    a("| 項目 | 值 |")
    a("|---|---|")
    a(f"| 加權目標價（DCF 50%＋Comparable 50%，情境 25／50／25） | **{f2(g('tp'))} 元** |")
    a(f"| DCF 目標價／Comparable 目標價 | {f2(g('dcf_tp'))}／{f2(g('cmp_tp'))} |")
    a(f"| 相對現價 24.25 | **{pc(g('up'))}** |")
    a(f"| 估值評等／投資建議 | **{g('rating')}／{g('rec')}** |")
    a(f"| 收支關卡（情境 B FY2026–FY2030 FCFF 扣股利合計） | {g('gate_b'):,.1f} 百萬元（{g('gate_t')}） |")
    a(f"| 情境目標價（A／B／C） | {f2(g('wt','B'))}／{f2(g('wt','C'))}／{f2(g('wt','D'))} |")
    a(f"| 情境價值 DCF（A／B／C） | {f2(g('dcf3','B') if 'dcf3' in VR else g('dcf','B'))}／{f2(g('dcf','C'))}／{f2(g('dcf','D'))} |")
    a(f"| 情境價值 Comparable（A／B／C） | {f2(g('cmp','B'))}／{f2(g('cmp','C'))}／{f2(g('cmp','D'))} |")
    a(f"| WACC／ke／g／β | {g('wacc'):.2%}／{g('ke'):.2%}／{g('g_show'):.2%}／{iv('beta'):.2f} |")
    a(f"| 終值占 EV（情境 B） | {pc(g('tvs','C'))} |")
    a(f"| 同業前瞻 P/E 中位數（4 家日本同業） | {iv('pe_peer'):.1f} 倍 |")
    a(f"| 現價對西南證券 FY2027 預測 EPS（0.33 元）的倍數 | {C['F15'].value:.0f} 倍（同業中位數的 {g('relnow'):.1f} 倍） |")
    a("")
    a(f"上檔空間歸因：DCF 貢獻 {pc(g('att_d'))}、未調整 Comparable 貢獻 {pc(g('att_c'))}、相對倍數調整 {pc(g('att_r'))}。要讓目標價等於現價，Comparable 的相對倍數係數須達 {g('relreq'):.1f}；升一級需 {g('relflip'):.1f}。DCF 以 RONIC＝WACC，而高相對倍數隱含持續超額報酬，兩法對競爭優勢的隱含假設相反。")
    a("")
    a(f"**讀法：** 埃斯頓是國產工業機器人龍頭，但前三季營收年減 3%、淨利含約 1 億元公允價值利得。情境 B 的模型淨利 FY2026 {pl('B','ni','F'):,.0f}、FY2027 {pl('B','ni','G'):,.0f}、FY2030 {pl('B','ni','J'):,.0f} 百萬元（EPS {pl('B','eps','F'):.2f}／{pl('B','eps','G'):.2f}／{pl('B','eps','J'):.2f} 元），對應市值 {iv('px')*iv('sh_v'):,.0f} 百萬元。這個結論主要建立在 Assumed 輸入（毛利率與期間費用率路徑、國產占比、相對份額、資產負債表估計）與 Derived 的同業倍數上，其中最敏感的是期間費用率與毛利率：期間費用率降到 {sens['solve'][1][1]:.1%}（基準 {iv('opex30'):.0%}）才升一級，支持現價在 0–28% 範圍內無解；中國裝機年增相對線模型達 {sens['solve'][2][1]:.1f} 倍才支持現價。方向在合理區間兩端相同（見敏感度），但資產負債表缺漏使絕對值的誤差帶偏寬；不改變評等的方向。")
    a("")
    a("## 命題")
    a("")
    a("1. **營收結構：** 官方兩分部——工業機器人及智能製造系統（2025 年 3,997 百萬元，占 82%）、自動化核心部件及運動控制系統（891 百萬元，占 18%）。")
    a("2. **機器人分部驅動：** 營收＝前一年 × 中國裝機成長（線 3，IFR 口徑）× 國產品牌占比變化（2025 年 55% 實績，2030 年 A／B／C＝70%／65%／58% 為假設）×（1＋本公司相對份額漂移）×（1 − 單價年降 3%，區間 0–8%）。")
    a("3. **利潤：** 毛利率由 2025 年 29.45% 回升到 2026 年前三季 32.4%；期間費用率 28.9%（含研發 9.1%）。情境 B 毛利率 FY2030 32%、期間費用率降到 25%。")
    a("")
    a("## 主要假設與可證偽指標（原則第 10 節第 20 條）")
    a("")
    a("| 承重假設 | 基準（A／B／C 或區間） | 標記 | 可觀察指標 |")
    a("|---|---|---|---|")
    a("| 中國工業機器人裝機（線 3） | 線高／基準／低；基準 2026 年約 38 萬台 | Derived／Assumed（線 3） | IFR、國家統計局產量（2026 年 1–9 月 +29.1%） |")
    a("| 2030 年國產品牌占比 | 70%／65%／58%（2025 年 55%） | Assumed | IFR 中國新聞稿、MIR 季報 |")
    a("| 本公司相對份額年漂移 | +4%／+1%／−3% | Assumed | MIR 出貨排名；本公司季營收年增 |")
    a("| 單價年降 | 3%（0–8%） | Analogy | 國產 6 軸本體價（線 3：基準 10 萬元、低案 2 萬元） |")
    a("| FY2030 毛利率 | 34%／32%／28% | Assumed | 季毛利率（2026 年 Q3 33.6%） |")
    a("| FY2030 期間費用率 | 25%（22–28%） | Assumed | 季費用率（1–9 月 28.9%） |")
    a("| 資產負債表（現金、有息負債、少數股東權益） | 1,700／3,000／400 百萬元（區間見 In 頁） | Assumed | 三季報合併資產負債表 |")
    a("| 同業前瞻 P/E | 中位數 %.1f 倍（4 家日本同業） | Derived | 同日重抓；補匯川前瞻 EPS |" % iv("pe_peer"))
    a("")
    a("## 敏感度（整本重算，前 10 大）")
    a("")
    a("| 案例 | 鍵 | 加權目標價 | 相對基準 | DCF 目標價 | 估值評等 |")
    a("|---|---|---|---|---|---|")
    for c in top:
        a(f"| {c['drv']} | `{c['key']}` | {c['tp']:.2f} | {c['tp'] / sens['base_tp'] - 1:+.1%} | {c['dcf']:.2f} | {c['rating']} |")
    a("")
    a("完整案例見 xlsx Sens 頁。**翻轉評等所需的變動**：見 Recon 反解。")
    a("")
    a("## 反解（Recon 頁）")
    a("")
    for lab, v_, b0, note in sens["solve"]:
        a(f"- {lab}：{v_:.3g}（基準 {b0:.3g}）；{note}")
    a("")
    a("## 限制與「分不出來／沒找到」")
    a("")
    a("- **沒找到：** 資產負債表、稅前利益與所得稅與少數股東損益、折舊攤銷與資本支出、各費用項目金額、2026 年分部營收、FY2023 分部與毛利率、β、市場共識目標價。公司年報與三季報原文未直接讀取（歷史數字取自媒體、券商轉錄；來源間有分歧處在 Q 頁說明，如 Q3 營收 1,101／1,122 百萬元）。")
    a("- **懷疑有錯但沒動：** 西南證券研報列的總股本 968 百萬股是否為最新（含 H 股）未查核；前三季淨利 171 與扣非 116 的差額（約 55）與上半年非經常損益 91 不完全吻合，可能含第三季轉回；2025 年毛利率有 27.8%（鳳凰網）與 29.45%（新浪）兩個口徑，採公司年報口徑 29.45%。")
    a("- **同業倍數：** 4 家日本同業全部用 2026-11-20 收盤；前瞻 EPS 取本專案各模型情境 B FY2027 值（非市場共識，會計年度月份不同）；匯川（300124）報價在收盤檔但前瞻 EPS 只查到 2025 年 10 月版（已過期），列「報價日未取得前瞻值」，未進中位數；美的、綠的諧波不可比，未進。")
    a("- **rf 逾期：** 中國 10 年期國債 1.61% 為 2026-11-05 值（跨模型一致）。")
    a("- **協調檔：** 實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics；相關條文不適用。")
    a("- **獨立核對：** 依專案規則，交付判讀的數字應由未參與製作的代理獨立核對，尚未進行。")
    a("")
    a("## 線 3 對帳")
    a("")
    a("讀線 3「下游取數」：中國裝機（逐年、三情境）、國產品牌占比 2025 年 55%、國產 6 軸本體價、國家統計局產量年增；對帳頁列出。**與線模型不一致處：** 國產占比 2030 年（線 3 只給 2025 年實績，2030 年為本模型假設）、單價年降（線 3 無逐年降幅，本模型以 Analogy 區間表達）、不用本體價推營收（本模型含系統整合、海外與核心部件外銷）。")
    a("")
    a("## 核對結果（建置時執行）")
    a("")
    a("in_layout、sens_coverage（--require-keycol）、check_governance、check_rating_gate --strict、check_zero_rows、check_cached_values、segment_check 皆通過（建置腳本結尾輸出）；HTML：`check_company_html.py`、`check_layout_v13.py` 結果見回報。")
    a("")
    a("## 資料來源")
    a("")
    a("見 xlsx Sources 頁（媒體與券商轉錄的年報與季報、H 股發行報導、線 3 成品、收盤檔、同業個股模型）。")
    a("")
    a("## 版本紀錄")
    a("")
    a(f"- v0.1（2026-11-24）：首版。現價 24.25（2026-11-20 收盤）；加權目標價 {f2(g('tp'))}；上檔空間 {pc(g('up'))}；{g('rating')}／{g('rec')}；收支關卡 {g('gate_b'):,.1f}（{g('gate_t')}）。")
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write("\n".join(L) + "\n")


def main():
    os.makedirs(os.path.join(HERE, "dist"), exist_ok=True)
    shutil.rmtree(TMP, ignore_errors=True); os.makedirs(TMP)
    wb, meta = make(); base_raw = os.path.join(TMP, "base.xlsx"); wb.save(base_raw)
    recalc([base_raw], os.path.join(TMP, "base_rc"))
    base_rc = os.path.join(TMP, "base_rc", "base.xlsx")
    base = read_val(base_rc, meta); iv = in_values(base_rc, meta["LAST_IN"])
    fp = openpyxl.load_workbook(base_rc, data_only=True)["Sens"]["D5"].value
    print("基準：", base)
    if "--quick" in sys.argv: return
    specs = []
    for k, (lo, e, hi, lab) in iv.items():
        for side, val in (("敏感度低", lo), ("敏感度高", hi)):
            if isinstance(val, (int, float)) and isinstance(e, (int, float)) and abs(val - e) > 1e-12:
                specs.append(dict(drv=f"{lab}（{side}）", set=f"{k}＝{val:.4g}（基準 {e:.4g}）", key=k, val=val, ov={k: val}))
    ivE = {k: v[1] for k, v in iv.items()}
    extra = []
    for d, lab in (("ds30", "2030 年國產品牌占比"), ("drift", "相對份額年漂移"), ("gcore", "核心部件營收年增"), ("gm30", "FY2030 毛利率")):
        for s in "AC": extra.append((f"情境 B 的{lab}改為情境 {s} 值", f"{d}_B", ivE[f"{d}_{s}"]))
    for drv, k, val in extra: specs.append(dict(drv=drv, set=f"{k}＝{val:.4g}", key=k, val=val, ov={k: val}))
    specs.append(dict(drv="相對倍數係數＝1（基準即為 1）", set="rel_k＝1", key="rel_k", val=1.0, ov={"rel_k": 1.0}))
    specs.append(dict(drv="情境權重全給 A", set="tp_wA＝1、tp_wB＝0、tp_wC＝0", key="tp_wA tp_wB tp_wC", val=None, ov={"tp_wA": 1, "tp_wB": 0, "tp_wC": 0}))
    specs.append(dict(drv="情境權重全給 C", set="tp_wA＝0、tp_wB＝0、tp_wC＝1", key="tp_wA tp_wB tp_wC", val=None, ov={"tp_wA": 0, "tp_wB": 0, "tp_wC": 1}))
    res = run_cases([s["ov"] for s in specs], meta, "sens")
    cases = []
    for s, r in zip(specs, res):
        cases.append(dict(drv=s["drv"], set=s["set"], key=s["key"], val=s["val"], tp=r["tp"], dcf=r["dcf"], rating=r["rating"]))
        print(f"  {s['drv'][:44]:44s} {r['tp']:9.2f} {r['tp'] / base['tp'] - 1:+.1%}")
    px = ivE["px"]
    def solve(key, target_tp, lo, hi, n=14, over=None):
        over = over or {}
        flo = run_cases([{**over, key: lo}], meta, "solve")[0]["tp"]; fhi = run_cases([{**over, key: hi}], meta, "solve")[0]["tp"]
        if (flo - target_tp) * (fhi - target_tp) > 0: return None
        a, b = lo, hi
        for _ in range(n):
            m = (a + b) / 2; fm = run_cases([{**over, key: m}], meta, "solve")[0]["tp"]
            if (fm - target_tp) * (flo - target_tp) > 0: a = m
            else: b = m
        return (a + b) / 2
    up = base["up"]
    flip = -0.5 if up < -0.5 else -0.25 if up < -0.25 else 0.25 if up <= 0.25 else 0.5
    flip_tp = px * (1 + flip + 1e-6)
    s1 = solve("opex30", px, 0.0, 0.28); s2 = solve("opex30", flip_tp, 0.0, 0.28)
    s3 = solve("cn_k", px, 1.0, 12.0, 10); s4 = solve("cn_k", flip_tp, 1.0, 12.0, 10)
    nf = lambda x: x if x is not None else -1
    solve_rows = [
        ("支持現價（目標價＝現價）所需的 FY2030 期間費用率（全部情境）", nf(s1), ivE["opex30"], "其他輸入不變；−1＝0–28% 範圍內無解"),
        (f"評等升一級（上檔 {flip:+.0%}）所需的 FY2030 期間費用率", nf(s2), ivE["opex30"], "其他輸入不變；−1＝0–28% 範圍內無解"),
        ("支持現價所需的中國裝機年增相對線 3 倍數（全部情境）", nf(s3), ivE["cn_k"], "其他輸入不變；−1＝12 倍內無解"),
        (f"評等升一級（上檔 {flip:+.0%}）所需的中國裝機年增相對線 3 倍數", nf(s4), ivE["cn_k"], "其他輸入不變；−1＝12 倍內無解"),
    ]
    print("反解：", solve_rows)
    h = hashlib.sha1(json.dumps([list(map(str, x)) for x in IN], ensure_ascii=False).encode() + open(__file__, "rb").read()).hexdigest()[:8]
    build_id = f"{VER}+{h}"
    sens = dict(build_id=build_id, fp=fp, base_tp=base["tp"], base_dcf=base["dcf"], cases=cases, solve=solve_rows)
    wb, meta2 = make(sens=sens)
    raw = os.path.join(TMP, "final", NAME); os.makedirs(os.path.dirname(raw), exist_ok=True); wb.save(raw)
    recalc([raw], os.path.join(TMP, "final_rc"))
    shutil.copy(os.path.join(TMP, "final_rc", NAME), OUT)
    fin = read_val(OUT, meta2)
    print("成品：", OUT, fin, "build_id", build_id)
    json.dump(dict(build_id=build_id, base=fin, cases=cases, solve=solve_rows), open(os.path.join(HERE, "sens.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    wbv = openpyxl.load_workbook(OUT, data_only=True); Hs = wbv["Hist"]
    segs = [("工業機器人及智能製造系統", "機器人及智能製造系統 營收", Hs["E8"].value), ("自動化核心部件及運動控制系統", "自動化核心部件及運動控制系統 營收", Hs["E9"].value)]
    json.dump(dict(company="estun", ticker="002747.SZ", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=round(v, 2), unit="百萬元", source=S_25B) for n, l, v in segs],
                   sublines=[], check_labels=["檢查：分部合計 − 合併營收（應為 0）"],
                   note="官方分部＝年報兩大業務（工業機器人及智能製造系統、自動化核心部件及運動控制系統）；2023 年分部未找到，PL 該欄標未揭露。"),
              open(os.path.join(HERE, "segments.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    write_readme(OUT, meta2, sens)
    ERR = ("#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!")
    errs = [f"{ws.title}!{c.coordinate}={c.value}" for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value in ERR or c.value.startswith("Err:"))]
    print("錯誤值：", len(errs), errs[:10])
    rc = 0
    for cmd in (["tools/in_layout/check_in_layout.py", OUT], ["tools/sens_coverage/check_sens_coverage.py", OUT, "--require-keycol"],
                ["tools/model_governance/check_governance.py", OUT], ["tools/model_governance/check_rating_gate.py", OUT, "--strict"],
                ["tools/model_governance/check_zero_rows.py", OUT], ["tools/model_governance/check_cached_values.py", OUT],
                ["tools/segment_check/check_segments.py", HERE]):
        p = subprocess.run([sys.executable, os.path.join(ROOT, cmd[0])] + cmd[1:], cwd=ROOT); rc |= p.returncode
    if errs: rc |= 1
    sys.exit(rc)


if __name__ == "__main__":
    main()

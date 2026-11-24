# 雙環傳動（002472.SZ，深圳證券交易所）個股模型 v0.1 建置腳本
# 產出：physical_ai/shuanghuan/dist/20261124_雙環傳動個股模型_v0.1.xlsx（LibreOffice 重算後的成品，含公式與快取值）
# 結構：命題模組（Thesis：新能源車齒輪＋燃油乘用車齒輪＋商用車＋工程機械＋電動工具與摩托車＋智能執行機構＋減速器及其他（含 RV 機器人減速機，讀線 6 RV 市場規模））
#       → 官方產品分部 → 標準損益表 → FCFF DCF＋同業 P/E。人形：線 4 與線 6 顯示 RV 減速機不因人形增量，本模型不另列人形營收（見對帳）。
# 敏感度：每個案例覆寫 In 頁基準欄後整本重建、LibreOffice 重算（整本重算，不用閉式近似）；反解同法以二分搜尋。
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md、docs/plan/高不確定性預估工作原則.md。
# 用法：python3 physical_ai/shuanghuan/build_shuanghuan.py   （需要 soffice；暫存於 $HOME/sh_build）
import datetime as dt
import glob
import hashlib
import json
import math
import os
import shutil
import subprocess
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
NAME = f"20261124_雙環傳動個股模型_{VER}.xlsx"
OUT = os.path.join(HERE, "dist", NAME)
HOMEDIR = os.path.expanduser("~")
TMP = os.path.join(HOMEDIR, "sh_build")
LO = ["soffice", f"-env:UserInstallation=file://{HOMEDIR}/lo_shuanghuan", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML"]

L4 = os.path.join(ROOT, "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx")
L6 = os.path.join(ROOT, "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx")
L3N = "physical_ai/lines/industrial_robot"
L4N, L6N = (os.path.relpath(p, ROOT) for p in (L4, L6))

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDRF = PatternFill("solid", fgColor="DDEBF7")
SECF = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------- 出處 ----------------
S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv：002472.SZ；未還原股利）"
S_AR = "雙環傳動 2025 年年度報告全文（2026-04；https://pdf.dfcfw.com/pdf/H2_AN202604231821516013_1.pdf；公司原文，直接讀取）"
S_Q3 = "雙環傳動 2026 年三季報，經東方財富研報頁轉錄（東吳證券 2026-10-29 https://data.eastmoney.com/report/info/AP202610291831181268.html；華龍證券 2026-10-28 https://data.eastmoney.com/report/info/AP202610281831101191.html；同花順新聞 2026-10-29 https://news.10jqka.com.cn/20261029/c681258420.shtml）：公司三季報原文未直接讀取"
S_H1 = "雙環傳動 2026 年半年報，經東方財富轉錄（https://finance.eastmoney.com/a/202608263854278566.html）：上半年淨利年增 1.65%、擬每 10 股派 1.2 元"
S_TP = "同花順 機構目標價彙整（2026-09-10；https://stock.10jqka.com.cn/20260910/c679791350.shtml）：近六個月 12 家機構，2026 年目標價 45.00–57.00、平均 49.74 元"
S_CICC = "中金公司 2025 年報點評（2026-04；引自搜尋結果摘要，原文未讀取）：目標價 45.0 元、2026 年淨利 14.3 億"
S_RF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與 leaderdrive 等模型同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：China CRP 0.91%（與 leaderdrive 等模型同一取值）"
S_CLOSE = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
S_SAP = "stockanalysis.com statistics 頁（資料商；原報價日見 Comp 頁 M 欄；數值取自 physical_ai/leaderdrive 模型 Comp 頁，2026-11-23 讀取）"
IP_BRK = "券商研究報告：有銷售研究與維持評等的誘因；歷史數字為轉錄，預測為券商自設"
IP_SA = "資料商彙整賣方共識，財報前的值未必更新"
IP_CO = "發行公司說法，有股價與融資誘因"


# ---------------- 線模型取數（以 ID 找列，不以列號） ----------------
def _line_rows(path, idcol, sheet="下游取數"):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    out = {}
    for r in range(1, ws.max_row + 1):
        k = ws.cell(r, idcol).value
        if isinstance(k, str): out[k] = (r, [ws.cell(r, c).value for c in range(1, 10)])
    return wb, out


def line_data():
    _, l4 = _line_rows(L4, 1)
    _, l6 = _line_rows(L6, 1)
    d = {"rv": {}, "hum": {}}
    for y in (2026, 2030, 2032):
        r, v = l6[f"RV 減速機市場規模 {y}"]
        for s, c, col in (("A", 3, "D"), ("B", 2, "C"), ("C", 1, "B")):   # v[1]=低 v[2]=基準 v[3]=高
            d["rv"][(y, s)] = (v[c], f"{L6N} 下游取數!{col}{r}（RV 減速機市場規模 {y}）")
    for y in (2026, 2030):
        r, v = l4[f"SHIP_{y}"]
        d["hum"][y] = (v[3], f"{L4N} 下游取數!D{r}（SHIP_{y}）")
    return d


LD = line_data()

# ---------------- In 頁輸入 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 來源誘因, 說明)
IN = [
    ("sec", "1. 市場資料與日期"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, "", "資料截止日 2026-11-24；股價 2026-11-20；成品日 2026-11-24"),
    ("px", "現價（2026-11-20 收盤）", "元", None, 35.45, None, "Verified", S_PX, "", "csv 值 35.450001；11-09 收盤 44.20，兩週跌 20%（全球 AI 與半導體股同週大跌）"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,12,31)", None, "Assumed", "本專案慣例：評價日 2026-12-31（會計年度＝日曆年，FY2026 在評價日已結束）", "", "折現期間由評價日起算"),
    ("tp_date", "目標時點（px_date 後 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", ""),
    ("bs_date", "資產負債表日", "日期", None, "=DATE(2025,12,31)", None, "Verified", S_AR, "", "2026-09-30 資產負債表未取得（資料商頁面 404，三季報原文未讀），故以年報日為準，Stub 期含全年 FY2026 FCFF"),
    ("stub_frac", "Stub 期（資產負債表日→評價日）", "年", None, "=(val_date-bs_date)/365", None, "Derived", "＝(評價日 − 資產負債表日) ÷ 365", "", "Stub FCFF＝FY2026 FCFF × stub_frac（約 1 年）"),
    ("sh_v", "評價日股數（總股本）", "百萬股", None, 849.77, None, "Verified", S_AR + "：股本 849,772,642 股（2025-12-31）", "", "員工持股計畫、限制性股票之稀釋未取得，以 0 計"),
    ("cash_v", "現金＋交易性金融資產（2025-12-31）", "百萬元", None, "=2490.47+221.19", None, "Verified", S_AR + "：貨幣資金 2,490.47（其中受限 123.3）、交易性金融資產 221.19", "", "受限資金未扣除"),
    ("ltinv_v", "長期投資（2025-12-31；長期股權投資＋其他權益工具投資）", "百萬元", None, "=38.86+46.16", None, "Verified", S_AR, "", "規模小"),
    ("inv_k", "長期投資計入淨現金的比例", "%", 0.5, 1.0, None, "Assumed", "基準全額計入；低案 50%", "", ""),
    ("debt_v", "有息負債（2025-12-31）", "百萬元", None, "=3122.49+592.09+273.99+12.71", None, "Verified", S_AR + "：短期借款 3,122.49、一年內到期非流動負債 592.09（含租賃 18.2）、長期借款 273.99、租賃負債 12.71", "", "2024 年底約 2,732；一年內到期含長借與租賃"),
    ("nci_v", "少數股東權益（2025-12-31）", "百萬元", None, 492.30, None, "Verified", S_AR, "", "自淨現金扣除；子公司環動科技等"),
    ("fx", "美元兌人民幣（2026-11-20）", "元／美元", None, 6.6665, None, "Verified", S_CLOSE + "：CNY=X", "", "用於把線 6 的 RV 市場規模（美元）換成人民幣"),
    ("cons_tp", "賣方目標價平均（12 家，2026-09，三季報前）", "元", None, 49.74, None, "Interested-party", S_TP, IP_SA, "只作對照，不進模型"),
    ("tp_cicc", "中金公司目標價（2026-04，年報後）", "元", None, 45.0, None, "Interested-party", S_CICC, IP_BRK, "只作對照，不進模型"),
    ("sec", "2. FY2026 各分部年增（錨：2026 年 1–9 月營收 +10.1%、第 3 季 +14.4%，第 3 季新能源車齒輪 +25%、燃油車齒輪 −14%）"),
    ("r9m26", "2026 年 1–9 月營收", "百萬元", None, 7118.0, None, "Verified", S_Q3, "", "Recon 用：檢查全年隱含第 4 季年增"),
    ("g9m26", "2026 年 1–9 月營收年增", "%", None, 0.1008, None, "Verified", S_Q3, "", ""),
    ("g_nev26", "FY2026 新能源車齒輪營收年增", "%", 0.16, 0.24, 0.30, "Assumed", "2025 年 +21.7%（" + S_AR + "）；2026 第 3 季約 +25%（" + S_Q3 + "，公司投資人關係紀錄）", "", "分部全年年增，公司只揭露第 3 季單季"),
    ("g_fuel26", "FY2026 燃油乘用車齒輪營收年增", "%", -0.20, -0.14, -0.08, "Assumed", "第 3 季傳統燃油車齒輪約 −14%（" + S_Q3 + "）", "", "燃油乘用車＝乘用車齒輪 − 新能源車齒輪（Derived）"),
    ("g_cv26", "FY2026 商用車齒輪營收年增", "%", -0.08, 0.0, 0.08, "Assumed", "2025 年 −6.95%；年報稱國內商用車全年產銷 426.1 萬輛、市場回暖（" + S_AR + "）", "", "2026 分部數字未取得"),
    ("g_con26", "FY2026 工程機械齒輪營收年增", "%", 0.0, 0.08, 0.15, "Assumed", "2025 年 +9.2%（" + S_AR + "）", "", ""),
    ("g_tm26", "FY2026 電動工具與摩托車齒輪營收年增", "%", -0.08, 0.0, 0.08, "Assumed", "2025 年合計 −6.4%（電動工具 +3.6%、摩托車 −20.0%）", "", "兩項合併列示"),
    ("g_act26", "FY2026 智能執行機構營收年增", "%", 0.08, 0.18, 0.28, "Assumed", "2025 年 +24.2%（" + S_AR + "）", "", ""),
    ("g_red26", "FY2026 減速器及其他營收年增", "%", 0.05, 0.12, 0.20, "Assumed", "2025 年 +21.2%；2026 上半年出口與匈牙利、越南廠爬坡（背景檔）", "", "含 RV 機器人減速機、其他減速器與雜項"),
    ("gm26", "FY2026 毛利率", "%", 0.255, 0.267, 0.275, "Assumed", "2026 第 3 季 26.78%（" + S_Q3 + "，華龍證券）；2025 年 27.34%（" + S_AR + "）", "", "原材料上漲壓縮毛利（東吳）"),
    ("opex26", "FY2026 銷售＋管理＋研發費用率", "%", 0.100, 0.107, 0.115, "Assumed", "2025 年 10.70%、2024 年 10.15%（Hist，Derived）", "", "管理費用含員工持股計畫攤銷"),
    ("sec", "3. 命題模組：減速器（RV）與各分部後續成長（量價合一的營收成長；RV 讀線 6 市場規模）"),
    ("rv_sh26", "FY2026「減速器及其他」中 RV 機器人減速機占比", "%", 0.40, 0.60, 0.80, "Assumed", "公司未揭露 RV 營收；環動科技 2024 年 RV 銷量 11.7 萬台、國內第一（背景檔引招股書）", "", "隱含占線 6 RV 市場份額見 Thesis 與 Recon"),
    ("sec", "4. 情境專屬驅動（A 樂觀／B 基準／C 保守；只列基準欄，不設敏感度區間）"),
    ("nevg_A", "情境 A：新能源車齒輪營收年增（FY2027）", "%", None, 0.24, None, "Assumed", "2025 年 +21.7%；外包趨勢與匈牙利廠（年報）；東吳 2027 淨利預測 +23%", "", "其後每年成長率乘上 nev_fade"),
    ("nevg_B", "情境 B：新能源車齒輪營收年增（FY2027）", "%", None, 0.16, None, "Assumed", "同上；乘用車終端需求與原材料壓力（東吳 2026-10 下調預測）", "", ""),
    ("nevg_C", "情境 C：新能源車齒輪營收年增（FY2027）", "%", None, 0.08, None, "Assumed", "同上；客戶電驅自製回流", "", ""),
    ("gfuel_A", "情境 A：燃油乘用車齒輪營收年增（FY2027 起）", "%", None, -0.06, None, "Assumed", "第 3 季 −14%；燃油車持續被替代", "", ""),
    ("gfuel_B", "情境 B：燃油乘用車齒輪營收年增（FY2027 起）", "%", None, -0.10, None, "Assumed", "同上", "", ""),
    ("gfuel_C", "情境 C：燃油乘用車齒輪營收年增（FY2027 起）", "%", None, -0.16, None, "Assumed", "同上", "", ""),
    ("gact_A", "情境 A：智能執行機構營收年增（FY2027 起）", "%", None, 0.25, None, "Assumed", "2025 年 +24.2%（年報）", "", ""),
    ("gact_B", "情境 B：智能執行機構營收年增（FY2027 起）", "%", None, 0.15, None, "Assumed", "同上", "", ""),
    ("gact_C", "情境 C：智能執行機構營收年增（FY2027 起）", "%", None, 0.05, None, "Assumed", "同上", "", ""),
    ("rvd_A", "情境 A：RV 份額年漂移（相對線 6 市場）", "%／年", None, 0.06, None, "Assumed", "國產替代；Nabtesco 全球 RV 份額 >53%、中國 60%（線 6 份額頁，Interested-party）", "", "相對線 6 RV 市場成長的超額（正＝取得份額）"),
    ("rvd_B", "情境 B：RV 份額年漂移（相對線 6 市場）", "%／年", None, 0.03, None, "Assumed", "同上；年報稱國產 RV 國產化率穩步提升", "", ""),
    ("rvd_C", "情境 C：RV 份額年漂移（相對線 6 市場）", "%／年", None, -0.02, None, "Assumed", "同上；中大力德等競爭", "", ""),
    ("gm30_A", "情境 A：FY2030 毛利率", "%", None, 0.285, None, "Assumed", "2024 年 25.0%（Derived）→2025 年 27.34%", "", "FY2026 起線性走到 FY2030，其後持平"),
    ("gm30_B", "情境 B：FY2030 毛利率", "%", None, 0.272, None, "Assumed", "同上；原材料上漲、海外廠爬坡", "", ""),
    ("gm30_C", "情境 C：FY2030 毛利率", "%", None, 0.255, None, "Assumed", "同上", "", ""),
    ("tp_wA", "情境權重 A", "", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", ""),
    ("tp_wB", "情境權重 B", "", None, 0.50, None, "Assumed", "同上", "", ""),
    ("tp_wC", "情境權重 C", "", None, 0.25, None, "Assumed", "同上", "", ""),
    ("sec", "5. 各情境共用的成長與損益參數"),
    ("nev_fade", "新能源車齒輪年增逐年衰減係數", "x", 0.6, 0.75, 0.9, "Assumed", "成長率每年乘以此係數（FY2028 起）；基數大、客戶集中度高", "", ""),
    ("gcv", "商用車齒輪營收年增（FY2027 起）", "%", -0.03, 0.02, 0.07, "Assumed", "2025 年 −6.95%；新能源商用車滲透率 26.9%（年報）", "", ""),
    ("gcon", "工程機械、電動工具與摩托車齒輪營收年增（FY2027 起）", "%", 0.0, 0.05, 0.10, "Assumed", "2025 年 +9.2%／−6.4%", "", ""),
    ("goth", "「減速器及其他」中非 RV 部分營收年增（FY2027 起）", "%", 0.03, 0.08, 0.15, "Assumed", "2025 年減速器及其他 +21.2%", "", ""),
    ("opex30", "FY2030 銷售＋管理＋研發費用率", "%", 0.095, 0.100, 0.107, "Assumed", "2025 年 10.70%、2024 年 10.15%；營運槓桿", "", "FY2026→FY2030 線性，其後持平"),
    ("ts_r", "稅金及附加 ÷ 營收", "%", None, "=Hist!E19/Hist!E4", None, "Derived", "＝2025 年稅金及附加 ÷ 營收", "", ""),
    ("oi_r", "其他收益（政府補助等）÷ 營收", "%", 0.008, "=Hist!E24/Hist!E4", 0.02, "Derived", "＝2025 年其他收益 145.0 ÷ 營收（" + S_AR + "）", "", "計入營業利益；2024 年 1.17%"),
    ("imp_r", "減值與資產處置損失 ÷ 營收", "%", 0.002, 0.005, 0.01, "Assumed", "2025 年 0.44%、2024 年 1.08%（信用減值＋資產減值＋處置）", "", "計入營業利益"),
    ("int_r", "有息負債利率（業外利息支出）", "%", 0.018, "=82.0/AVERAGE(2732,4001)", 0.03, "Derived", "＝2025 年利息費用 82.0 ÷ 平均有息負債（2024 年底約 2,732、2025 年底 4,001）", "", "業外＝現金收益 − 利息支出"),
    ("cy_r", "現金收益率", "%", 0.008, "=29.84/AVERAGE(1407,2490)", 0.025, "Derived", "＝2025 年利息收入 29.84 ÷ 平均貨幣資金", "", ""),
    ("tax_r", "有效稅率", "%", 0.10, "=Hist!E33/Hist!E32", 0.20, "Derived", "＝2025 年所得稅 ÷ 稅前利益（高新技術企業）", "", "敏感度低／高 10%／20%（Assumed）"),
    ("nci_r", "少數股東損益 ÷ 稅後淨利", "%", None, "=Hist!E35/Hist!E34", None, "Derived", "＝2025 年少數股東損益 ÷ 稅後淨利", "", "環動科技等子公司；2024 年 3.1%"),
    ("payout", "配息率", "%", None, "=0.28/1.50", None, "Derived", "＝2025 年每股股利 0.28 元 ÷ EPS 1.50 元（" + S_AR + "）", "", "用於收支關卡與目標時點前股利"),
    ("da26", "FY2026 折舊攤銷 ÷ 營收", "%", None, "=Hist!E40/Hist!E4", None, "Derived", "＝2025 年折舊攤銷 918 ÷ 營收", "", "之後線性走向資本支出比例 × da_ss（FY2032）"),
    ("da_ss", "穩態折舊 ÷ 資本支出", "x", None, 0.8, None, "Assumed", "穩態資本支出高於折舊（成長）", "", ""),
    ("cx26", "FY2026 資本支出", "百萬元", 1500, 1800, 2100, "Assumed", "2024 年 1,989、2025 年 2,069（現金流量表購建支出；" + S_AR + "）；在建工程 2,182", "", "2026 年指引未取得"),
    ("cx_r", "資本支出 ÷ 營收（FY2027 起）", "%", 0.10, 0.15, 0.20, "Assumed", "2024、2025 年約 22.6%／22.7%（Derived）；擴產高峰後回落", "", "決定自由現金流與 DCF 的主要輸入"),
    ("nwc_r", "營運資金增加 ÷ 營收增量", "%", 0.15, 0.25, 0.40, "Assumed", "2025 年底應收 2,758、存貨 2,226、應付帳款與票據 3,039（年報）", "", ""),
    ("sec", "6. 折現率（DCF 標準規格第 2 節）"),
    ("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", S_RF, "", "日期 2026-11-05，較股價基準日早 11 個交易日（逾期；跨模型一致，下次全面改版更新）"),
    ("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", ""),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", ""),
    ("b_own", "本公司 5 年 β（原始）", "x", None, 0.52, None, "Interested-party", S_SAP + "：β 0.52", IP_SA, "只有一個資料商來源"),
    ("b_peer", "同業 5 年 β 中位數（原始）", "x", None, "=MEDIAN(Comp!H5:H8)", None, "Derived", "＝Comp 頁核心同業 4 家 β 中位數", "", ""),
    ("beta", "β（兩來源 Blume 調整後中位數）", "x", "=beta-0.3", "=MEDIAN(0.67*b_own+0.33,0.67*b_peer+0.33)", "=beta+0.3", "Derived", "Blume 0.67β＋0.33；兩來源＝本公司、同業中位數", "", "敏感度 ∓0.3（跨模型慣例）"),
    ("kd", "稅前債務成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt（模型值；實際利息率見 int_r）", "", ""),
    ("wacc_adj", "WACC 調整（敏感度用）", "pt", -0.01, 0, 0.01, "Assumed", "基準 0；敏感度 ±1pt", "", ""),
    ("g_cap", "永續成長率上限", "%", None, 0.03, None, "Assumed", "DCF 標準規格：3%，且不超過計價幣別 10 年期公債殖利率", "", ""),
    ("g_tv", "永續成長率", "%", 0.01, "=MIN(g_cap,rf)", 0.03, "Derived", "＝MIN(g_cap, rf)", "", "敏感度 1%／3%（3% 超過 rf，只作敏感度）"),
    ("ronic_sp", "RONIC − WACC", "pt", None, 0, 0.05, "Assumed", "RONIC 預設＝WACC（新投資價值中性）", "", ""),
    ("tv_alt", "終值改用 FCFF ×（1＋g）（0＝否，1＝是）", "開關", None, 0, 1, "Assumed", "DCF 標準規格第 4 節敏感度", "", ""),
    ("sec", "7. 同業倍數與權重"),
    ("pe_peer", "同業前瞻 P/E 中位數", "x", "=MIN(Comp!F5:F8)", "=MEDIAN(Comp!F5:F8)", "=MAX(Comp!F5:F8)", "Derived", "＝Comp 頁核心同業 4 家前瞻 P/E 中位數；股價皆為 2026-11-20 收盤、P/E 由原報價日等比例換算（EPS 分母不變，見 Comp）", "", "敏感度＝同業最低／最高"),
    ("rel_k", "相對倍數係數（本公司 ÷ 同業）", "x", 0.5, 1.0, 1.5, "Assumed", "1.00（未調整：缺本公司對同業十年相對倍數歷史）；低案 0.5 約等於本公司現行倍數 ÷ 中位數", "", "同業含機器人選擇權定價的公司，見 Comp 說明"),
    ("w_dcf", "DCF 權重（Comparable＝1 − 本值）", "%", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設 DCF 50%＋同業倍數 50%", "", ""),
]
SCN_DRIVERS = [("nevg", "新能源車齒輪營收年增（FY2027）"), ("gfuel", "燃油乘用車齒輪營收年增"), ("gact", "智能執行機構營收年增"),
               ("rvd", "RV 份額年漂移"), ("gm30", "FY2030 毛利率")]

YRS = list(range(2023, 2033))
COL = {y: openpyxl.utils.get_column_letter(3 + i) for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032
TC = {y: openpyxl.utils.get_column_letter(4 + (y - 2025)) for y in range(2025, 2033)}  # Thesis：D＝2025 … K＝2032
SC = "ABC"
LINE_SC = {"A": "線模型高情境", "B": "線模型基準", "C": "線模型低情境"}


def short(s):
    """Recon 說明欄（HTML 會顯示）：去掉網址與 repo 路徑，改用簡稱。"""
    import re
    s = str(s)
    for p_, n_ in ((L3N, "線 3 成品"), (L4N, "線 4 成品"), (L6N, "線 6 成品")):
        s = s.replace(p_, n_)
    s = re.sub(r"[；;]?\s*https?://[^\s）)；]+", "", s)
    s = re.sub(r"（[A-Z][A-Z0-9]*_[A-Z0-9_]+）", "", s)
    return s


def hdr(ws, r, labels, c0=1):
    for j, h in enumerate(labels):
        c = ws.cell(r, c0 + j, h); c.font = BOLD; c.fill = HDRF


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def T(x, f="#,##0"):
    return f'TEXT({x},"{f}")'


def make(over=None, sens=None):
    """over：{鍵: 數值} 覆寫 In 基準欄；sens：最終建置時寫入 Sens 頁的結果 dict。"""
    over = over or {}
    wb = openpyxl.Workbook()
    RD = wb.active; RD.title = "README"
    IW = wb.create_sheet("In"); HI = wb.create_sheet("Hist"); Q = wb.create_sheet("Q"); LN = wb.create_sheet("線取數"); TH = wb.create_sheet("Thesis")
    PL = wb.create_sheet("PL"); DC = wb.create_sheet("DCF"); V = wb.create_sheet("Val"); CP = wb.create_sheet("Comp")
    SE = wb.create_sheet("Sens"); RC = wb.create_sheet("Recon"); DZ = wb.create_sheet("對帳"); RK = wb.create_sheet("Risk"); SO = wb.create_sheet("Sources")

    # ================= In =================
    IW["A1"] = ("In 頁：藍字為輸入。基準欄 E 進主模型；敏感度低（D）／敏感度高（F）只用於 Sens 頁（一次動一個參數、整本重算），不進目標價。"
                "情境 A／B／C 的差異只寫在情境專屬列（第 4 節）與線模型情境路徑（線取數頁：A＝線模型高、B＝基準、C＝低），最後一節「情境驅動對照」純展示。")
    IW["A2"] = "單位：百萬元人民幣，除另註；會計年度＝日曆年（FY2026＝2026 年）。中國會計準則（法定財報口徑）。"
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
                if unit in ("%", "%／年"): cc.number_format = "0.0%"
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
        IW.cell(r, 8, f"引用 In 頁 {d}_A／{d}_B／{d}_C；情境專屬輸入"); r += 1
    IW.cell(r, 1, "scn_w"); IW.cell(r, 2, "情境權重")
    for j, s in enumerate(SC):
        IW.cell(r, 4 + j, f"=E{KR['tp_w' + s]}")
    IW.cell(r, 8, "引用 In 頁 tp_wA／tp_wB／tp_wC；合計＝1；線模型需求路徑（線取數頁）A＝高、B＝基準、C＝低，非本頁輸入")
    for col, w in zip("ABCDEFGHIJ", (12, 46, 10, 11, 12, 11, 15, 70, 30, 50)):
        IW.column_dimensions[col].width = w
    # ================= Hist（年度；百萬元；中國會計準則）=================
    HI["A1"] = "Hist：歷史財務（百萬元人民幣；中國會計準則）。FY2024、FY2025 取自公司 2025 年報（元換算為百萬元）；FY2023 只取得營收、淨利、EPS、營業現金流（來自年報主要財務指標表），其餘未取得。"
    hdr(HI, 3, ["項目", "", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"])
    HROWS = [  # (列, 鍵, 項目, [2023, 2024, 2025], 標記, 來源)
        (4, "h_rev", "營業收入", [8074.19, 8781.40, 9112.17], "Verified", S_AR),
        (5, "h_cogs", "營業成本", [None, 6585.49, 6621.25], "Verified", S_AR),
        (6, "h_gp", "毛利（＝營收 − 營業成本）", [None, "=D4-D5", "=E4-E5"], "Derived", "＝營業收入 − 營業成本"),
        (7, "h_gm", "毛利率", [None, "=D6/D4", "=E6/E4"], "Derived", "＝毛利 ÷ 營收（2025 年 27.34%＝年報）"),
        (13, "h_sell", "銷售費用", [None, 85.63, 99.50], "Verified", S_AR),
        (14, "h_admin", "管理費用", [None, 349.97, 384.86], "Verified", S_AR + "；2025 年含員工持股計畫攤銷 59.1"),
        (15, "h_rnd", "研發費用", [None, 456.04, 490.51], "Verified", S_AR),
        (16, "h_opex", "營業費用（銷售＋管理＋研發）", [None, "=SUM(D13:D15)", "=SUM(E13:E15)"], "Derived", "＝三項合計"),
        (17, "h_fin", "財務費用", [None, 37.99, 46.99], "Verified", S_AR),
        (18, "h_inv", "投資收益＋公允價值變動", [None, -11.81, "=-9.21-3.90"], "Verified", S_AR),
        (19, "h_ts", "稅金及附加", [None, 55.74, 54.97], "Verified", S_AR),
        (20, "h_opp", "營業利潤（中國會計準則）", [None, 1207.15, 1505.73], "Verified", S_AR),
        (21, "h_ebit", "營業利益（核心：營業利潤 − 投資收益與公允價值 ＋ 財務費用）", [None, "=D20-D18+D17", "=E20-E18+E17"], "Derived", "本模型營業利益不含金融項目（DCF 另以淨現金計入）"),
        (22, "h_oth_op", "其他營業損益（＝核心營業利益 − 毛利 ＋ 稅金及附加 ＋ 營業費用）", [None, "=D21-D6+D19+D16", "=E21-E6+E19+E16"], "Derived", "其他收益、減值、資產處置合計（殘差）"),
        (23, "h_imp", "減值與資產處置損失", [None, "=-31.49-54.55-8.48", "=-13.34-19.79-7.13"], "Verified", S_AR + "：信用減值、資產減值、資產處置"),
        (24, "h_oi", "其他收益（政府補助等）", [None, 102.95, 145.00], "Verified", S_AR),
        (25, "h_ebitm", "營業利益率（核心）", [None, "=D21/D4", "=E21/E4"], "Derived", ""),
        (32, "h_pbt", "利潤總額（稅前）", [None, 1203.21, 1504.84], "Verified", S_AR),
        (26, "h_nonop", "業外（＝稅前 − 核心營業利益）", [None, "=D32-D21", "=E32-E21"], "Derived", "利息淨支出、投資收益、公允價值、營業外收支合計"),
        (33, "h_tax", "所得稅", [None, 146.61, 179.57], "Verified", S_AR),
        (34, "h_nit", "稅後淨利", [None, 1056.60, 1325.27], "Verified", S_AR),
        (35, "h_nci", "少數股東損益", [None, 32.69, 63.76], "Verified", S_AR),
        (36, "h_ni", "歸屬母公司淨利", [816.41, 1023.91, 1261.52], "Verified", S_AR),
        (37, "h_eps", "EPS（基本，元）", [0.97, 1.22, 1.50], "Verified", S_AR),
        (38, "h_dps", "每股股利（元）", [None, None, 0.28], "Verified", S_AR + "：每 10 股派 2.8 元；2024 年股利未取得"),
        (39, "h_sh", "股數（＝淨利 ÷ EPS，百萬股）", ["=C36/C37", "=D36/D37", "=E36/E37"], "Derived", ""),
        (40, "h_da", "折舊攤銷", [None, "=735.76+14.58+19.38", "=877.68+16.86+23.81"], "Verified", S_AR + "：現金流量表補充資料（固定資產折舊＋無形資產攤銷＋長期待攤）"),
        (41, "h_cx", "資本支出（現金購建長期資產）", [None, 1988.62, 2069.42], "Verified", S_AR),
        (42, "h_ocf", "營業現金流", [1571.68, 1681.24, 2391.31], "Verified", S_AR),
        (44, "h_pc", "乘用車齒輪 營收", [None, 5324.89, 6004.52], "Verified", S_AR),
        (45, "h_nev", "　其中新能源車齒輪", [None, "=E45/1.2172", 4102.38], "Verified", S_AR + "：2025 年 41.02 億、+21.72%；2024＝2025 ÷ 1.2172（Derived）"),
        (46, "h_fuel", "　其中燃油乘用車齒輪（＝乘用車 − 新能源）", [None, "=D44-D45", "=E44-E45"], "Derived", "殘差；含少量非新能源乘用車以外項目"),
        (47, "h_cv", "商用車齒輪 營收", [None, 721.67, 671.53], "Verified", S_AR),
        (48, "h_con", "工程機械齒輪 營收", [None, 579.94, 633.13], "Verified", S_AR),
        (49, "h_tm", "電動工具與摩托車齒輪 營收", [None, "=129.60+95.09", "=134.28+76.04"], "Verified", S_AR + "：電動工具、摩托車合計"),
        (50, "h_act", "智能執行機構 營收", [None, 642.35, 797.62], "Verified", S_AR),
        (51, "h_red", "減速器及其他 營收", [None, 656.10, 795.05], "Verified", S_AR + "（含 RV 機器人減速機；RV 營收未揭露）"),
        (52, "h_steel", "鋼材銷售（貿易）營收", [None, 631.76, 0], "Verified", S_AR + "：2025 年停止"),
    ]
    for (rr, k, lab, vals, tag, src) in HROWS:
        HI.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = HI.cell(rr, 3 + j, v)
                if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
                if k in ("h_gm", "h_ebitm"): c.number_format = "0.0%"
        HI.cell(rr, 6, tag); HI.cell(rr, 7, src)
        dn(wb, k, f"Hist!$C${rr}:$E${rr}")
    HI.cell(53, 1, "檢查：分部合計 − 營業收入（應為 0；FY2024、FY2025）")
    for c in "DE": HI[f"{c}53"] = f"={c}44+{c}47+{c}48+{c}49+{c}50+{c}51+{c}52-{c}4"
    HI.cell(29, 1, "差額：核心營業利益＋業外 − 稅前（應為 0）")
    for c in "DE": HI[f"{c}29"] = f"={c}21+{c}26-{c}32"
    HI.column_dimensions["A"].width = 58; HI.column_dimensions["G"].width = 90

    # ================= Q（季度；一列一個數字，單一標記）=================
    Q["A1"] = "Q：2026 年季度資料（百萬元）。一列一個項目；三季報經券商與新聞轉錄（原文未直接讀取）。"
    hdr(Q, 3, ["項目", "2026Q1", "2026Q2", "2026H1", "2026Q3", "2026 年 1–9 月", "標記", "來源", "來源誘因", "說明"])
    QROWS = [
        (4, "營業收入", [2095, "=D4-B4", 4559, 2559, "=D4+E4"], "Verified", S_Q3 + "；" + S_H1, "", "1–9 月合計與 In 頁 r9m26（71.18 億）差為取整"),
        (5, "營業收入年增", [0.015, None, 0.078, 0.144, 0.1008], "Verified", S_Q3, "", ""),
        (6, "歸屬母公司淨利", [284, "=D6-B6", "=F6-E6", 333, 920], "Verified", S_Q3, "", "Q1 +2.9%、上半年 +1.65%、Q3 +3.76%、1–9 月 +2.40%"),
        (7, "毛利率", [None, None, None, 0.2678, None], "Verified", S_Q3 + "（華龍證券：Q3 毛利率 26.78%，年減 0.07pt、季增 0.67pt）", "", ""),
        (8, "新能源車齒輪占營收", [None, None, None, 0.50, None], "Verified", S_Q3 + "：Q3 超過 50%、年增約 25%、季增約 18%", "", "公司投資人關係紀錄，下限值"),
        (9, "燃油車齒輪年增", [None, None, None, -0.14, None], "Verified", S_Q3, "", "季增約持平"),
    ]
    for (rr, lab, vals, tag, src, inc, note) in QROWS:
        Q.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = Q.cell(rr, 2 + j, v)
                if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
                if isinstance(v, float) and abs(v) < 1: c.number_format = "0.00%"
        Q.cell(rr, 7, tag); Q.cell(rr, 8, src); Q.cell(rr, 9, inc); Q.cell(rr, 10, note)
    Q.column_dimensions["A"].width = 30; Q.column_dimensions["H"].width = 90

    # ================= 線取數（線 6 RV 市場規模；藍字＝讀入值，M 欄寫來源檔與儲存格）=================
    LN["A1"] = ("線取數：RV 減速機市場規模（線 6 下游取數；五線歸因，只來自傳統工業機器人；百萬美元）。2026、2030、2032 為線模型讀入值（藍字），中間年以幾何內插（Derived）。"
                "情境對應：A＝線模型高、B＝基準、C＝低。")
    hdr(LN, 3, ["項目", "情境", ""] + [f"FY{y}" for y in range(2025, 2033)] + ["單位", "來源（檔案與儲存格）"])
    LR = {}
    r = 4
    for s in SC:
        LN.cell(r, 1, "RV 減速機市場規模（線 6）"); LN.cell(r, 2, f"{s}（{LINE_SC[s]}）")
        for y in (2026, 2030, 2032):
            c = LN.cell(r, 4 + (y - 2025), LD["rv"][(y, s)][0]); c.font = BLUE; c.number_format = "#,##0"
        for y, (a, b, n0, n1) in {2027: (2026, 2030, 2026, 4), 2028: (2026, 2030, 2026, 4), 2029: (2026, 2030, 2026, 4), 2031: (2030, 2032, 2030, 2)}.items():
            ca, cb = TC[a], TC[b if y != 2031 else 2032]
            LN[f"{TC[y]}{r}"] = f"={ca}{r}*({cb}{r}/{ca}{r})^(({y}-{a})/{n1})"; LN[f"{TC[y]}{r}"].number_format = "#,##0"
        LN.cell(r, 12, "百萬美元"); LN.cell(r, 13, f"{LD['rv'][(2026, s)][1]}；{LD['rv'][(2030, s)][1]}；{LD['rv'][(2032, s)][1]}")
        LR[s] = r; r += 1
    LN.column_dimensions["A"].width = 40; LN.column_dimensions["M"].width = 110

    # ================= Thesis（命題模組）=================
    TH["A1"] = ("命題模組：各分部營收由驅動因子推導。FY2026＝FY2025 × (1＋分部年增)（錨：1–9 月實績與第 3 季分部年增，見 Recon 全年隱含第 4 季）；新能源車齒輪＝前一年 ×（1＋首年增 × 衰減^(年−2027)）；"
                "燃油乘用車、商用車、工程機械等＝前一年 ×（1＋年增）；RV 機器人減速機＝前一年 × 線 6 RV 市場成長 ×（1＋份額漂移）；人形不另列（線 6：RV 不因人形增量）。百萬元。")
    TH["A2"] = "欄：D＝FY2025 … K＝FY2032。"
    hdr(TH, 3, ["項目", "說明", ""] + [f"FY{y}" for y in range(2025, 2033)])
    THR = {}
    r = 4
    for s in SC:
        TH.cell(r, 1, f"情境 {s}").font = BOLD; TH.cell(r, 1).fill = SECF; r += 1
        R = {}
        def trow(key, lab, note, f25, f26, fn):
            global_r = None
        def put(key, lab, note=""):
            nonlocal r
            R[key] = r; TH.cell(r, 1, lab); TH.cell(r, 2, note); r += 1
            return R[key]
        def seg(key, lab, hrow, g26, gk):
            rr = put(key, lab, f"FY2026＝FY2025 ×（1＋{g26}）；之後＝前一年 ×（1＋{gk}）")
            TH[f"{TC[2025]}{rr}"] = f"=Hist!E{hrow}"
            for y in range(2026, 2033):
                TH[f"{TC[y]}{rr}"] = (f"={TC[y - 1]}{rr}*(1+{g26})" if y == 2026 else f"={TC[y - 1]}{rr}*(1+{gk})")
        # 新能源
        rr = put("nev", "新能源車齒輪 營收", "FY2026＝FY2025 ×（1＋g_nev26）；之後＝前一年 ×（1＋nevg × nev_fade^(年−2027)）")
        TH[f"{TC[2025]}{rr}"] = "=Hist!E45"
        for y in range(2026, 2033):
            TH[f"{TC[y]}{rr}"] = f"={TC[y - 1]}{rr}*(1+g_nev26)" if y == 2026 else f"={TC[y - 1]}{rr}*(1+nevg_{s}*nev_fade^({y}-2027))"
        seg("fuel", "燃油乘用車齒輪 營收", 46, "g_fuel26", f"gfuel_{s}")
        seg("cv", "商用車齒輪 營收", 47, "g_cv26", "gcv")
        seg("con", "工程機械齒輪 營收", 48, "g_con26", "gcon")
        seg("tm", "電動工具與摩托車齒輪 營收", 49, "g_tm26", "gcon")
        seg("act", "智能執行機構 營收", 50, "g_act26", f"gact_{s}")
        rr = put("red26", "減速器及其他 FY2026 合計（錨）", "＝FY2025 ×（1＋g_red26）")
        TH[f"{TC[2025]}{rr}"] = "=Hist!E51"; TH[f"{TC[2026]}{rr}"] = f"={TC[2025]}{rr}*(1+g_red26)"
        rr = put("mkt", "線 6 RV 減速機市場規模（百萬美元）", "讀線取數頁")
        for y in range(2026, 2033): TH[f"{TC[y]}{rr}"] = f"=線取數!{TC[y]}{LR[s]}"
        rr = put("rv", "RV 機器人減速機 營收", "FY2026＝減速器及其他 × rv_sh26；之後＝前一年 × 市場成長 ×（1＋rvd）")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{rr}"] = f"={c}{R['red26']}*rv_sh26" if y == 2026 else f"={TC[y - 1]}{rr}*{c}{R['mkt']}/{TC[y - 1]}{R['mkt']}*(1+rvd_{s})"
        rr = put("rvsh", "　RV 營收隱含占線 6 市場份額（營收 ÷（市場 × 匯率））", "")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{rr}"] = f"={TC[y]}{R['rv']}/({TC[y]}{R['mkt']}*fx)"; TH[f"{TC[y]}{rr}"].number_format = "0.00%"
        rr = put("rvo", "其他減速器與雜項 營收", "FY2026＝減速器及其他 ×（1 − rv_sh26）；之後＝前一年 ×（1＋goth）")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{rr}"] = f"={c}{R['red26']}*(1-rv_sh26)" if y == 2026 else f"={TC[y - 1]}{rr}*(1+goth)"
        rr = put("red", "減速器及其他 合計")
        TH[f"{TC[2025]}{rr}"] = "=Hist!E51"
        for y in range(2026, 2033): TH[f"{TC[y]}{rr}"] = f"={TC[y]}{R['rv']}+{TC[y]}{R['rvo']}"
        rr = put("rev", "合併營收")
        TH[f"{TC[2025]}{rr}"] = "=Hist!E4"
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{rr}"] = f"={c}{R['nev']}+{c}{R['fuel']}+{c}{R['cv']}+{c}{R['con']}+{c}{R['tm']}+{c}{R['act']}+{c}{R['red']}"
        rr = put("nevsh", "新能源車齒輪占合併營收")
        for y in range(2025, 2033):
            TH[f"{TC[y]}{rr}"] = f"={TC[y]}{R['nev']}/{TC[y]}{R['rev']}"; TH[f"{TC[y]}{rr}"].number_format = "0.0%"
        rr = put("gm", "毛利率路徑", "FY2026＝gm26；線性到 FY2030 gm30，其後持平")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{rr}"] = "=gm26" if y == 2026 else f"=gm26+(gm30_{s}-gm26)*MIN(1,({y}-2026)/4)"; TH[f"{TC[y]}{rr}"].number_format = "0.0%"
        rr = put("opex", "營業費用率路徑", "FY2026＝opex26；線性到 FY2030 opex30，其後持平")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{rr}"] = "=opex26" if y == 2026 else f"=opex26+(opex30-opex26)*MIN(1,({y}-2026)/4)"; TH[f"{TC[y]}{rr}"].number_format = "0.0%"
        r += 1
        THR[s] = R
    for row in TH.iter_rows(min_row=4, max_row=r, min_col=4, max_col=11):
        for c in row:
            if c.number_format == "General": c.number_format = "#,##0.0"
    TH.column_dimensions["A"].width = 44; TH.column_dimensions["B"].width = 56

    # ================= PL（標準損益表，三情境）=================
    PL["A1"] = ("PL：官方產品分部（乘用車齒輪［新能源／燃油］、商用車、工程機械、電動工具與摩托車、智能執行機構、減速器及其他［RV／其他］、鋼材銷售）→ 合併營收 → 標準損益表 → FCFF。"
                "C＝FY2023 … E＝FY2025 為實績（FY2023 只取得營收與淨利）；F＝FY2026 … L＝FY2032。百萬元；中國會計準則（法定財報口徑，GAAP）。")
    PR = {}
    r = 3
    HAS23 = {4, 36, 37, 39}
    for s in ["B", "A", "C"]:
        R = {}; th = THR[s]
        pre = "" if s == "B" else f"（情境 {s}）"
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
        H = lambda rr: (lambda y: None if (y == 2023 and rr not in HAS23) else f"=Hist!{'CDE'[y - 2023]}{rr}")
        H24 = lambda rr: (lambda y: None if y == 2023 else f"=Hist!{'CDE'[y - 2023]}{rr}")
        C = lambda y: COL[y]
        T_ = lambda k: (lambda y: f"=Thesis!{TC[y]}{th[k]}")
        row("nev", "乘用車：新能源車齒輪（子列）", "百萬元", H24(45), T_("nev"))
        row("fuel", "乘用車：燃油乘用車齒輪（子列）", "百萬元", H24(46), T_("fuel"))
        row("pc", "乘用車齒輪 營收", "百萬元", H24(44), lambda y: f"={C(y)}{R['nev']}+{C(y)}{R['fuel']}")
        row("cv", "商用車齒輪 營收", "百萬元", H24(47), T_("cv"))
        row("con", "工程機械齒輪 營收", "百萬元", H24(48), T_("con"))
        row("tm", "電動工具與摩托車齒輪 營收", "百萬元", H24(49), T_("tm"))
        row("act", "智能執行機構 營收", "百萬元", H24(50), T_("act"))
        row("rv", "減速器：RV 機器人減速機（子列）", "百萬元", None, T_("rv"))
        row("rvo", "減速器：其他減速器與雜項（子列）", "百萬元", None, T_("rvo"))
        row("red", "減速器及其他 營收", "百萬元", H24(51), lambda y: f"={C(y)}{R['rv']}+{C(y)}{R['rvo']}")
        row("steel", "鋼材銷售 營收", "百萬元", H24(52), lambda y: "=0")
        row("rev", "營收（合併）", "百萬元", H(4), lambda y: f"={C(y)}{R['pc']}+{C(y)}{R['cv']}+{C(y)}{R['con']}+{C(y)}{R['tm']}+{C(y)}{R['act']}+{C(y)}{R['red']}+{C(y)}{R['steel']}")
        row("chk", "檢查：分部合計 − 合併營收（應為 0）", "百萬元",
            lambda y: None if y == 2023 else f"={C(y)}{R['pc']}+{C(y)}{R['cv']}+{C(y)}{R['con']}+{C(y)}{R['tm']}+{C(y)}{R['act']}+{C(y)}{R['red']}+{C(y)}{R['steel']}-{C(y)}{R['rev']}",
            lambda y: f"={C(y)}{R['pc']}+{C(y)}{R['cv']}+{C(y)}{R['con']}+{C(y)}{R['tm']}+{C(y)}{R['act']}+{C(y)}{R['red']}+{C(y)}{R['steel']}-{C(y)}{R['rev']}")
        PL.cell(R["chk"], 1, "檢查：分部合計 − 合併營收（應為 0）" if s == "B" else f"檢查（情境 {s}）：分部合計 − 合併營收（應為 0）")
        row("chk2", "檢查：子列合計 − 分部（應為 0）", "百萬元", None,
            lambda y: f"={C(y)}{R['rv']}+{C(y)}{R['rvo']}-Thesis!{TC[y]}{th['red']}+{C(y)}{R['nev']}+{C(y)}{R['fuel']}-{C(y)}{R['pc']}")
        row("yoy", "營收年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1",
            lambda y: f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1", "0.0%")
        row("cogs", "營業成本", "百萬元", H24(5), lambda y: f"={C(y)}{R['rev']}*(1-Thesis!{TC[y]}{th['gm']})")
        row("gp", "毛利", "百萬元", lambda y: None if y == 2023 else f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}")
        row("gm", "毛利率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", "0.0%")
        row("ts", "稅金及附加", "百萬元", H24(19), lambda y: f"={C(y)}{R['rev']}*ts_r")
        row("opex", "營業費用（銷售＋管理＋研發）", "百萬元", H24(16), lambda y: f"={C(y)}{R['rev']}*Thesis!{TC[y]}{th['opex']}")
        row("othop", "其他營業損益（其他收益 − 減值等）", "百萬元", H24(22), lambda y: f"={C(y)}{R['rev']}*(oi_r-imp_r)")
        row("op", "營業利益（核心，不含金融項目）", "百萬元", H24(21), lambda y: f"={C(y)}{R['gp']}-{C(y)}{R['ts']}-{C(y)}{R['opex']}+{C(y)}{R['othop']}")
        row("opm", "營業利益率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['op']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", "0.0%")
        row("nonop", "其他收入淨額（業外：利息淨額、投資收益、公允價值、營業外）", "百萬元", H24(26), lambda y: "=cash_v*cy_r-debt_v*int_r")
        row("pbt", "稅前利益", "百萬元", H24(32), lambda y: f"={C(y)}{R['op']}+{C(y)}{R['nonop']}")
        row("tax", "所得稅", "百萬元", H24(33), lambda y: f"={C(y)}{R['pbt']}*tax_r")
        row("nit", "稅後淨利（含少數股東）", "百萬元", H24(34), lambda y: f"={C(y)}{R['pbt']}-{C(y)}{R['tax']}")
        row("nci", "少數股東損益", "百萬元", H24(35), lambda y: f"={C(y)}{R['nit']}*nci_r")
        row("ni", "歸屬母公司淨利", "百萬元", H(36), lambda y: f"={C(y)}{R['nit']}-{C(y)}{R['nci']}")
        row("nim", "淨利率", "%", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", "0.0%")
        row("niy", "淨利年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1",
            lambda y: f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1", "0.0%")
        row("sh", "股數（百萬股）", "百萬股", H(39), lambda y: "=sh_v")
        row("eps", "EPS（元）", "元", H(37), lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['sh']}", "0.00")
        row("da", "折舊攤銷", "百萬元", H24(40), lambda y: f"={C(y)}{R['rev']}*(da26+(cx_r*da_ss-da26)*{y - 2026}/6)")
        row("cx", "資本支出", "百萬元", H24(41), lambda y: "=cx26" if y == 2026 else f"={C(y)}{R['rev']}*cx_r")
        row("dnwc", "營運資金增加", "百萬元", None, lambda y: f"=nwc_r*({C(y)}{R['rev']}-{COL[y - 1]}{R['rev']})")
        row("nopat", "NOPAT（＝營業利益 ×（1 − 稅率））", "百萬元", None, lambda y: f"={C(y)}{R['op']}*(1-tax_r)")
        row("fcff", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "百萬元", None, lambda y: f"={C(y)}{R['nopat']}+{C(y)}{R['da']}-{C(y)}{R['cx']}-{C(y)}{R['dnwc']}")
        row("div", "股利（淨利 × 配息率）", "百萬元", None, lambda y: f"=MAX(0,{C(y)}{R['ni']})*payout")
        row("gate", "FCFF 扣股利（收支）", "百萬元", None, lambda y: f"={C(y)}{R['fcff']}-{C(y)}{R['div']}")
        r += 1
        PR[s] = R
    for row_ in PL.iter_rows(min_row=3, max_row=r, min_col=3, max_col=12):
        for c in row_:
            if c.number_format == "General": c.number_format = "#,##0.0"
    PL.column_dimensions["A"].width = 52

    # ================= DCF =================
    DC["A1"] = "DCF：FCFF 折現（期中慣例；評價日 2026-12-31，明確期 FY2027–FY2032）。終值＝NOPAT(FY2032) ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）；tv_alt＝1 時改用 FCFF ×（1＋g）÷（WACC − g）。"
    hdr(DC, 3, ["項目"] + [f"FY{y}" for y in range(2027, 2033)])
    DCOL = {y: openpyxl.utils.get_column_letter(2 + y - 2027) for y in range(2027, 2033)}
    DC.cell(4, 1, "折現年數（評價日起，期中）"); DC.cell(5, 1, "折現因子")
    for y in range(2027, 2033):
        c = DCOL[y]; DC[f"{c}4"] = y - 2027 + 0.5; DC[f"{c}5"] = f"=1/(1+wacc)^{c}4"
    DR = {}
    r = 7
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
        R["nc"] = r; DC.cell(r, 1, "淨現金滾到評價日（現金與投資 − 有息負債 − 少數股東權益 ＋ Stub FCFF）")
        DC[f"B{r}"] = f"=cash_v+ltinv_v*inv_k-debt_v-nci_v+B{R['stub']}"; r += 1
        R["eq"] = r; DC.cell(r, 1, "股權價值"); DC[f"B{r}"] = f"=B{R['ev']}+B{R['nc']}"; r += 1
        R["ps"] = r; DC.cell(r, 1, "每股價值（評價日，元）"); DC[f"B{r}"] = f"=MAX(0,B{R['eq']}/sh_v)"; r += 1
        R["dps"] = r; DC.cell(r, 1, "評價日→目標時點間股利（FY2026 盈餘配息，約 2027-06 支付）"); DC[f"B{r}"] = f"=PL!F{PR[s]['eps']}*payout"; r += 1
        R["v12"] = r; DC.cell(r, 1, "推到目標時點（×（1＋ke）^t − 期間股利）"); DC[f"B{r}"] = f"=MAX(0,B{R['ps']}*(1+ke)^((tp_date-val_date)/365)-B{R['dps']})"; r += 2
        DR[s] = R
    for row_ in DC.iter_rows(min_row=4, max_row=r, min_col=2, max_col=7):
        for c in row_:
            if c.number_format == "General": c.number_format = "#,##0.00"
    DC.column_dimensions["A"].width = 70
    # ================= Comp =================
    CP["A1"] = ("Comp：同業前瞻 P/E（stockanalysis.com statistics，資料商；數值取自 leaderdrive 模型 Comp 頁）。股價＝" + S_CLOSE + "（Verified）；前瞻 P/E＝原值 × 新收盤 ÷ 原報價（Derived，EPS 不變）；"
                "哈默納科、Nabtesco 取本專案模型同日值。核心同業 4 家（齒輪／減速／汽車零件）進中位數；鳴志（馬達，P/E 逾 100 倍）、中大力德與五洲新春（資料商無前瞻 P/E）、綠的諧波（諧波減速，P/E 逾 200 倍）不進中位數。"
                "同業多以機器人選擇權定價，本公司主體為汽車齒輪，倍數口徑不同，見 Val 頁相對倍數反解。")
    hdr(CP, 4, ["公司", "代號", "股價（2026-11-20）", "報價日", "前瞻 P/E 分母", "前瞻 P/E", "近 12 月 P/E（原報價日）", "5 年 β", "會計口徑", "進中位數", "說明", "原報價", "原報價日", "原前瞻 P/E", "標記與換算"])
    PEERS = [  # (公司, 代號, 11-20 收盤, 原報價, 原報價日, 原前瞻 P/E, 近 12 月 P/E, β, 口徑, 進中位數, 說明, 分母)
        ("拓普集團", "601689.SS", 48.11, 46.52, "2026-09-30", 26.89, 32.10, 0.80, "中國會計準則", "是", "汽車零件、人形致動器", "資料商 NTM 共識 EPS"),
        ("Nabtesco", "6268.T", 3105, 3105, "2026-11-20", 13.77, None, 0.93, "IFRS（未查核）", "是", "RV 減速機；取 physical_ai/nabtesco 模型 In 頁 pe_self", "資料商 NTM 共識 EPS（11-13 值換算）"),
        ("哈默納科", "6324.T", 3722, 3722, "2026-11-20", 48.94, None, 0.96, "日本基準（未查核）", "是", "諧波減速機；取 physical_ai/harmonic_drive 模型 Comp!F14", "公司 FY2026 指引 EPS（2027年3月期）"),
        ("兆威機電", "003021.SZ", 76.83, 70.25, "2026-09-30", 59.70, 87.15, 0.44, "中國會計準則", "是", "微型傳動、靈巧手", "資料商 NTM 共識 EPS"),
        ("鳴志電器", "603728.SS", 44.52, 51.50, "2026-09-30", 124.10, 302.24, 0.62, "中國會計準則", "否", "步進與空心杯馬達；P/E 逾 100 倍、業務為馬達，不進中位數", "資料商 NTM 共識 EPS"),
        ("中大力德", "002896.SZ", 48.36, 56.87, "2026-10-09", None, 259.22, 0.25, "中國會計準則", "否", "RV、行星減速器；資料商無前瞻 P/E（報價日未取得前瞻值）", "—"),
        ("五洲新春", "603667.SS", 41.60, 44.45, "2026-10-13", None, 221.31, 0.86, "中國會計準則", "否", "軸承、行星滾柱絲杠；資料商無前瞻 P/E", "—"),
        ("綠的諧波", "688017.SS", 230.80, 263.76, "2026-10-09", 225.96, 342.80, 0.79, "中國會計準則", "否", "諧波減速；P/E 逾 200 倍，不進中位數", "資料商 NTM 共識 EPS（三季報前）"),
        ("雙環傳動（本公司，對照）", "002472.SZ", 35.45, 41.21, "2026-10-30", 22.32, 27.11, 0.52, "中國會計準則", "否", "本公司；只作對照", "資料商 NTM 共識 EPS"),
    ]
    for i, (nm, tk, px1, px0, d0, pe0, ttm, bt, acc, med, note, den) in enumerate(PEERS):
        rr = 5 + i
        CP.cell(rr, 1, nm); CP.cell(rr, 2, tk)
        CP.cell(rr, 3, px1).font = BLUE; CP.cell(rr, 4, "2026-11-20"); CP.cell(rr, 5, den)
        if pe0 is not None: CP.cell(rr, 6, f"=N{rr}*C{rr}/L{rr}")
        else: CP.cell(rr, 6, "未取得")
        if ttm is not None: CP.cell(rr, 7, ttm).font = BLUE
        CP.cell(rr, 8, bt).font = BLUE; CP.cell(rr, 9, acc); CP.cell(rr, 10, med); CP.cell(rr, 11, note)
        CP.cell(rr, 12, px0).font = BLUE; CP.cell(rr, 13, d0)
        if pe0 is not None: CP.cell(rr, 14, pe0).font = BLUE
        CP.cell(rr, 15, "股價 Verified（" + S_CLOSE + "）；前瞻 P/E Derived：＝原值 × 2026-11-20 收盤 ÷ 原報價（分母 EPS 不變）；原值 " + S_SAP)
    CP.cell(15, 1, "核心同業前瞻 P/E 中位數"); CP["F15"] = "=MEDIAN(F5:F8)"
    CP.cell(16, 1, "本公司現價 ÷ 資料商前瞻 EPS（換算至 2026-11-20；分母：資料商 NTM 共識，三季報前）"); CP["F16"] = "=F13"
    CP.cell(17, 1, "本公司現價 ÷ 券商 FY2027 EPS 平均（分母：東吳 1,606、華龍 1,641 百萬元淨利 ÷ 849.77 百萬股）"); CP["F17"] = "=px/AVERAGE(1606/849.77,1641/849.77)"
    CP.cell(18, 1, "本公司現價 ÷ 模型 FY2027 EPS（分母：情境 B，FY2027）"); CP["F18"] = f"=px/PL!G{PR['B']['eps']}"
    CP.cell(19, 1, "本公司現行倍數 ÷ 同業中位數（資料商 NTM 口徑）"); CP["F19"] = "=F16/F15"
    CP.column_dimensions["A"].width = 60; CP.column_dimensions["K"].width = 50; CP.column_dimensions["O"].width = 80

    # ================= Val =================
    V["A1"] = "Val：DCF（FCFF，WACC）50% ＋ 同業 P/E（Comparable）50%，三情境 25／50／25 加權；約 12 個月目標價（元）。"
    VR = {}
    r = 3
    V.cell(r, 1, "0. 折現率").font = BOLD; r += 1
    for key, lab, f, fmt in (("ke", "股權成本 ke＝rf＋β ×（ERP＋CRP）", "=rf+beta*(erp+crp)", "0.00%"),
                             ("kd_at", "稅後債務成本", "=kd*(1-tax_r)", "0.00%"),
                             ("e_mv", "股權市值（百萬元）", "=px*sh_v", "#,##0"),
                             ("d_bv", "有息負債（百萬元）", "=debt_v", "#,##0.00"),
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
    def row3(key, lab, fs, fmt="#,##0.0", note=""):
        nonlocal r
        V.cell(r, 1, lab)
        for j, f in enumerate(fs): V.cell(r, 2 + j, f).number_format = fmt
        V.cell(r, 5, note); VR[key] = r; r += 1
    row3("ev", "企業價值 EV（百萬元）", [f"=DCF!B{DR[s]['ev']}" for s in SC], "#,##0")
    row3("tvs", "終值占 EV", [f"=DCF!B{DR[s]['tvs']}" for s in SC], "0.0%")
    row3("dcf", "DCF 情境價值推到 12 個月後", [f"=DCF!B{DR[s]['v12']}" for s in SC])
    row3("eps27", "FY2027 EPS（模型）", [f"=PL!G{PR[s]['eps']}" for s in SC], "0.00", "分母：模型各情境 FY2027 EPS")
    row3("cmpu", "Comparable 未調整（同業 P/E × FY2027 EPS）", [f"=MAX(0,{c}{r - 1}*pe_peer)" for c in "BCD"], "#,##0.0", "同業 P/E＝In 頁 pe_peer（核心同業前瞻 P/E 中位數）")
    row3("cmp", "Comparable（× 相對倍數係數）", [f"={c}{VR['cmpu']}*rel_k" for c in "BCD"], "#,##0.0", "rel_k＝1.00（未調整：缺十年相對倍數歷史）")
    row3("wt", "情境目標價：DCF × w ＋ Comparable ×（1 − w）", [f"=w_dcf*{c}{VR['dcf']}+(1-w_dcf)*{c}{VR['cmp']}" for c in "BCD"])
    row3("w", "情境權重", ["=tp_wA", "=tp_wB", "=tp_wC"], "0%")
    r += 1
    V.cell(r, 1, "2. 目標價與評等").font = BOLD; r += 1
    def row1(key, lab, f, fmt="#,##0.0", note=""):
        nonlocal r
        V.cell(r, 1, lab); V[f"B{r}"] = f; V[f"B{r}"].number_format = fmt; V.cell(r, 5, note); VR[key] = r; r += 1
    row1("dcf_tp", "DCF 目標價（情境加權）", f"=SUMPRODUCT(B{VR['dcf']}:D{VR['dcf']},B{VR['w']}:D{VR['w']})")
    row1("cmp_tp", "Comparable 目標價（情境加權；同業 P/E × 相對倍數係數）", f"=SUMPRODUCT(B{VR['cmp']}:D{VR['cmp']},B{VR['w']}:D{VR['w']})")
    row1("cmpu_tp", "Comparable 未調整（情境加權）", f"=SUMPRODUCT(B{VR['cmpu']}:D{VR['cmpu']},B{VR['w']}:D{VR['w']})")
    row1("tp", "加權目標價（DCF 與 Comparable 加權）", f"=w_dcf*B{VR['dcf_tp']}+(1-w_dcf)*B{VR['cmp_tp']}")
    row1("px_show", "現價（股價基準日）", "=px")
    row1("cons_show", "賣方目標價平均（同花順 12 家，2026-09；只作對照）", "=cons_tp")
    row1("up", "上檔空間：加權目標價", f"=B{VR['tp']}/px-1", "0.0%")
    row1("rating", "估值評等（只看上檔空間）", f'=IF(B{VR["up"]}>0.5,"強力買進",IF(B{VR["up"]}>0.25,"買進",IF(B{VR["up"]}>=-0.25,"中立",IF(B{VR["up"]}>=-0.5,"放空","強力放空"))))', "@",
         "級距：> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空（原則第 10 節第 24 條）")
    V.cell(r, 1, "收支關卡：FY2026–FY2030 FCFF 扣股利合計（情境 A／B／C，百萬元）")
    for j, s in enumerate(SC):
        V.cell(r, 2 + j, f"=SUM(PL!F{PR[s]['gate']}:J{PR[s]['gate']})").number_format = "#,##0.0"
    V.cell(r, 5, "口徑：FCFF − 股利（不含金融資產收益）；基準情境 B < 0＝入不敷出"); VR["gate"] = r; r += 1
    row1("gate_b", "入不敷出判定用（情境 B 合計）", f"=C{VR['gate']}", "#,##0.0")
    row1("gate_t", "入不敷出判定", f'=IF(B{VR["gate_b"]}<0,"不通過（入不敷出）","通過")', "@")
    row1("rec", "投資建議", f'=IF(AND(B{VR["gate_b"]}<0,OR(B{VR["rating"]}="買進",B{VR["rating"]}="強力買進")),"收支不足",B{VR["rating"]})', "@",
         "原則第 10 節第 63 條：收支不通過且估值評等為買進以上時＝收支不足")
    row1("gate_cover", "收支缺口 ÷ 淨現金（情境 B；負值才有意義）", f"=B{VR['gate_b']}/(cash_v+ltinv_v*inv_k-debt_v)", "0.0%", "淨現金可否覆蓋 FY2026–FY2030 缺口")
    row1("dist", "距最近級距邊界（百分點）", f"=MIN(ABS(B{VR['up']}-0.5),ABS(B{VR['up']}-0.25),ABS(B{VR['up']}+0.25),ABS(B{VR['up']}+0.5))", "0.0%")
    row1("up_b", "升一級所需目標價（上一級邊界）", f'=IF(B{VR["up"]}>0.5,"—（已是最高級）",px*(1+IF(B{VR["up"]}<-0.5,-0.5,IF(B{VR["up"]}<-0.25,-0.25,IF(B{VR["up"]}<=0.25,0.25000001,0.50000001)))))')
    row1("dn_b", "降一級門檻目標價（下一級邊界）", f'=IF(B{VR["up"]}<-0.5,"—（已是最低級）",px*(1+IF(B{VR["up"]}>0.5,0.5,IF(B{VR["up"]}>0.25,0.25,IF(B{VR["up"]}>=-0.25,-0.25000001,-0.50000001)))))')
    r += 1
    V.cell(r, 1, "3. 上檔空間三段歸因與相對倍數（原則第 10 節第 47 條）").font = BOLD; r += 1
    row1("att_d", "DCF 貢獻＝w ×（DCF − 現價）÷ 現價", f"=w_dcf*(B{VR['dcf_tp']}-px)/px", "0.0%")
    row1("att_c", "未調整 Comparable 貢獻＝(1 − w) ×（未調整 − 現價）÷ 現價", f"=(1-w_dcf)*(B{VR['cmpu_tp']}-px)/px", "0.0%")
    row1("att_r", "相對倍數調整貢獻", f"=(1-w_dcf)*(B{VR['cmp_tp']}-B{VR['cmpu_tp']})/px", "0.0%")
    row1("att_chk", "檢查：三段合計 − 上檔空間（應為 0）", f"=B{VR['att_d']}+B{VR['att_c']}+B{VR['att_r']}-B{VR['up']}", "0.0000")
    row1("relnow", "本公司現行前瞻倍數 ÷ 同業中位數（資料商 NTM 口徑）", "=Comp!F19", "0.00")
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
            SE.cell(r, 5, c["tp"]).number_format = "#,##0.0"; SE.cell(r, 6, c["tp"] / sens["base_tp"] - 1).number_format = "+0.0%;-0.0%"
            SE.cell(r, 7, c["dcf"]).number_format = "#,##0.0"; SE.cell(r, 8, c["dcf"] / sens["base_dcf"] - 1 if sens["base_dcf"] else None).number_format = "+0.0%;-0.0%"
            SE.cell(r, 9, c["rating"]); r += 1
        r += 1
        SE.cell(r, 1, "讀法：相對基準＝案例加權目標價 ÷ 基準加權目標價 − 1；同一案例 DCF 單法與加權目標價並列（原則第 10 節第 32 條）。情境專屬驅動以「情境 B 改為情境 A／C 值」呈現。")
        SE.cell(r + 1, 1, "說明：折舊、資本支出（FY2027 起）、營運資金隨營收等比例變動（第 44 條）；各分部年增為情境與共用參數；RV 讀線 6。")
    SE.column_dimensions["A"].width = 52; SE.column_dimensions["B"].width = 44

    # ================= Recon =================
    RC["A1"] = "Recon：模型對實績、券商預測、線 6 市場規模的同口徑對照；反解現價與評等邊界所需的主導變數。"
    hdr(RC, 3, ["項目", "模型", "對照", "差距（模型 − 對照）", "說明"])
    rr = 4
    def rrow(lab, a, b, note, fmt="#,##0.0", diff=True):
        nonlocal rr
        RC.cell(rr, 1, lab); RC.cell(rr, 2, a).number_format = fmt
        if b is not None: RC.cell(rr, 3, b).number_format = fmt
        if diff and b is not None: RC.cell(rr, 4, f"=B{rr}-C{rr}").number_format = fmt
        RC.cell(rr, 5, short(note)); rr += 1
    rrow("FY2026 全年隱含第 4 季營收年增（情境 B，百萬元口徑）vs 第 3 季實績年增 14.4%", f"=(PL!F{PR['B']['rev']}-r9m26)/(Hist!E4-r9m26/(1+g9m26))-1", 0.144, "全年營收由分部年增推導，再與 1–9 月實績對照；差距大代表分部假設需修正", "0.0%")
    rrow("FY2026 營收年增（情境 B）vs 1–9 月實績年增", f"=PL!F{PR['B']['yoy']}", "=g9m26", "", "0.0%")
    rrow("FY2026 毛利率 vs 2026 第 3 季實績", f"=PL!F{PR['B']['gm']}", "=Q!E7", "第 3 季 26.78%（華龍證券）", "0.0%")
    rrow("FY2026 歸屬母公司淨利（情境 B）vs 券商 2026E 平均（東吳 1,301、華龍 1,381，三季報後）", f"=PL!F{PR['B']['ni']}", "=AVERAGE(1301,1381)", "券商預測為 Interested-party；1–9 月實績 920")
    rrow("FY2026 第 4 季隱含歸母淨利（全年 − 1–9 月 920）vs 2025 年第 4 季（1,261.5 − 920÷1.024）", f"=PL!F{PR['B']['ni']}-Q!F6", "=Hist!E36-Q!F6/1.024", "2025 年第 4 季約 363")
    rrow("FY2027 歸屬母公司淨利（情境 B）vs 券商 2027E 平均（東吳 1,606、華龍 1,641）", f"=PL!G{PR['B']['ni']}", "=AVERAGE(1606,1641)", "")
    rrow("FY2028 歸屬母公司淨利（情境 B）vs 券商 2028E 平均（東吳 1,943、華龍 1,982）", f"=PL!H{PR['B']['ni']}", "=AVERAGE(1943,1982)", "")
    rrow("FY2027 EPS（情境 B）vs 券商 2027E 隱含（淨利 ÷ 849.77 百萬股）", f"=PL!G{PR['B']['eps']}", "=AVERAGE(1606,1641)/sh_v", "", "0.00")
    rrow("FY2026–FY2030 年均資本支出（情境 B）vs 2024–2025 年均實績", f"=AVERAGE(PL!F{PR['B']['cx']}:J{PR['B']['cx']})", "=AVERAGE(Hist!D41:E41)", "模型年均低於實績年均的程度，是自由現金流為正的主要假設")
    rrow("加權目標價 vs 賣方目標價平均（同花順 12 家，2026-09）", f"=Val!B{VR['tp']}", "=cons_tp", S_TP + "；只作對照，不進模型")
    rrow("加權目標價 vs 中金目標價（2026-04）", f"=Val!B{VR['tp']}", "=tp_cicc", "只作對照，不進模型")
    rr += 1
    RC.cell(rr, 1, "線 6 市場對照（RV 減速機，五線歸因，只來自傳統工業機器人）").font = BOLD; rr += 1
    rrow("FY2026 RV 營收隱含占線 6 市場份額（情境 B）", f"=Thesis!E{THR['B']['rvsh']}", None, "RV 營收＝減速器及其他 × RV 占比假設（Assumed）；對照 Nabtesco 全球 >53%（線 6 份額頁，Interested-party）", "0.0%", False)
    rrow("FY2030 RV 營收隱含占線 6 市場份額（情境 B）", f"=Thesis!I{THR['B']['rvsh']}", None, "份額漂移累積後", "0.0%", False)
    rrow("FY2032 RV 營收隱含占線 6 市場份額（情境 B）", f"=Thesis!K{THR['B']['rvsh']}", None, "", "0.0%", False)
    rrow("線 6 RV 市場規模成長 FY2026→FY2030（情境 B）", f"=Thesis!I{THR['B']['mkt']}/Thesis!E{THR['B']['mkt']}-1", None, f"{LD['rv'][(2026, 'B')][1]}；{LD['rv'][(2030, 'B')][1]}", "0.0%", False)
    rr += 1
    RC.cell(rr, 1, "反解（整本重算二分搜尋；建置時計算，數值為快照）").font = BOLD; rr += 1
    hdr(RC, rr, ["反解目標", "解", "基準值", "", "說明"]); rr += 1
    if sens:
        for lab, v, b0, note in sens["solve"]:
            RC.cell(rr, 1, lab); RC.cell(rr, 2, v).number_format = "#,##0.00"; RC.cell(rr, 3, b0).number_format = "#,##0.00"; RC.cell(rr, 5, note); rr += 1
    RC.column_dimensions["A"].width = 70; RC.column_dimensions["E"].width = 90

    # ================= 對帳 =================
    DZ["A1"] = "對帳：本模型使用的「線層級」變數（與 physical_ai/lines/ 各線模型對帳；本模型讀線 v0.1 成品，不靜默改口徑）。"
    hdr(DZ, 3, ["變數名", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型是否一致"])
    r26 = LD["rv"][(2026, "B")]; r30 = LD["rv"][(2030, "B")]; r32 = LD["rv"][(2032, "B")]
    ZR = [
        ("RV 減速機市場規模 2026／2030／2032（A／B／C＝線高／基準／低）", "6 共用零組件", f"基準 {r26[0]:,.0f}／{r30[0]:,.0f}／{r32[0]:,.0f}", "百萬美元", "五線歸因，只來自傳統工業機器人", f"{r26[1]}；{r30[1]}；{r32[1]}", "Derived", "Thesis RV 營收成長（線取數頁；中間年幾何內插，Derived）", "一致（直接讀值）；2027–2029、2031 為本模型內插"),
        ("人形機器人對 RV 減速機的增量", "6 共用零組件／4 人形", "0", "—", "線 6：人形主流用諧波＋行星，RV 不因人形增量", f"{L6N} 下游取數（RV 市場說明欄）；背景檔 map/profiles/002472.md", "Derived", "本模型不列人形營收", "一致（採線 6 口徑）；人形用行星減速或執行機構的雙環營收未揭露，含在智能執行機構與其他分部的假設成長內"),
        ("全球人形機器人出貨 2026／2030（基準）", "4 人形", f"{LD['hum'][2026][0]:,.0f}／{LD['hum'][2030][0]:,.0f}", "台", "全球、含小型機", f"{LD['hum'][2026][1]}；{LD['hum'][2030][1]}", "Assumed", "未使用（只作對照）", "未使用：本模型沒有人形用營收列；若日後揭露人形客戶，再接 SHIP_* 與每台用量"),
        ("RV 減速機全球份額（Nabtesco）", "6 共用零組件", ">53%（全球）；60%（中國），2025", "%", "全球", f"{L6N} 份額頁第 7 列", "Interested-party", "Recon 對照（不進公式）", "雙環份額線 6 未建立（背景檔：份額!F8 空白）；本模型隱含份額見 Recon"),
        ("環動科技 RV 銷量 2024", "—", "11.7 萬台（國內第一）", "台", "公司招股書引述", "背景檔 map/profiles/002472.md（格隆匯引環動招股書）", "Interested-party", "rv_sh26 的依據（定性）", "線模型未讀；雙環 RV 營收未揭露"),
    ]
    for i, z in enumerate(ZR):
        for j, v in enumerate(z): DZ.cell(4 + i, 1 + j, v)
    DZ.column_dimensions["A"].width = 50; DZ.column_dimensions["F"].width = 80; DZ.column_dimensions["I"].width = 70

    # ================= Risk =================
    RK["A1"] = "Risk：風險登錄（原則第 10 節第 10、19 條）。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    RISKS = [
        ("新能源車終端需求與客戶電驅自製回流", "是（nevg_C、g_nev26 低案）", "東吳 2026-10 下調預測", "第 3 季起分部營收、前五大客戶占比（2025 年 32.68%）"),
        ("原材料上漲壓縮毛利", "是（gm26、gm30）", "第 3 季毛利率 26.78%，年減 0.07pt", "季毛利率"),
        ("擴產資本支出高、自由現金流薄", "是（cx26、cx_r、收支關卡）", "2024、2025 年購建支出約 2,000 百萬元／年；短期借款 3,122", "在建工程 2,182、有息負債"),
        ("燃油車與商用車齒輪持續下滑", "是（g_fuel26、gfuel、gcv）", "第 3 季燃油車齒輪 −14%", "分部營收"),
        ("RV 機器人減速機國產替代速度與價格競爭", "部分（rvd、rv_sh26）", "RV 營收未揭露；市場不因人形增量", "環動科技銷量、中大力德"),
        ("海外廠（匈牙利、越南）爬坡成本", "部分（gm 路徑）", "國外營收 2025 年 +58.9%、毛利率 26.0% 低於內銷 28.2%", "國外營收占比與毛利率"),
        ("員工持股計畫攤銷與稀釋", "否", "2025 年攤銷 59.1；稀釋股數未取得", "股本變動"),
        ("環動科技分拆上市終止（2026-09）", "否", "母公司保留機器人業務全部權益；少數股東權益 492 已自淨現金扣除", "股東會決議"),
    ]
    for i, k in enumerate(RISKS):
        for j, v in enumerate(k): RK.cell(4 + i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["D"].width = 60

    # ================= Sources =================
    SO["A1"] = "Sources：本模型引用的原始來源與資料商（含網址；HTML 展示層不放網址）。"
    for i, s_ in enumerate([S_PX, S_AR, S_Q3, S_H1, S_TP, S_CICC, S_RF, S_CRP, S_SAP,
                            f"線 4：{L4N}", f"線 6：{L6N}",
                            "physical_ai/harmonic_drive 模型（哈默納科前瞻 P/E）、physical_ai/nabtesco 模型（Nabtesco 前瞻 P/E）、physical_ai/leaderdrive 模型（A 股同業前瞻 P/E）",
                            "physical_ai/map/profiles/002472.md（背景，數字已另行查證）"], 3):
        SO.cell(i, 1, s_)
    SO.column_dimensions["A"].width = 160

    # ================= README（xlsx 內）=================
    S = T
    v = lambda k: f"Val!B{VR[k]}"
    lines = [
        ('="雙環傳動（002472.SZ）個股模型 ' + VER + '｜build_id "&Sens!D2&"｜資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'),
        "命題：雙環傳動的獲利主要來自新能源車齒輪放量（外包趨勢與海外廠），燃油車與商用車齒輪下滑；機器人減速機（RV）受線 6 市場規模（不因人形增量）限制，不是短期獲利來源。營收由各官方分部的年增推導，RV 讀線 6。",
        (f'="結論（條件式）：在新能源車齒輪 FY2027 年增 "&{S("nevg_B","0%")}&"（逐年衰減 "&{S("nev_fade","0.00")}&"）、燃油乘用車齒輪年降 "&{S("-gfuel_B","0%")}&"、FY2030 毛利率 "&{S("gm30_B","0.0%")}&"、資本支出占營收 "&{S("cx_r","0%")}&"、同業 P/E 未加溢價（相對倍數係數 "&{S("rel_k","0.00")}&"）的假設下，加權目標價約 "&{S(v("tp"),"#,##0.0")}&" 元，相對現價 "&{S("px","#,##0.00")}&" 元 "&{S(v("up"),"+0.0%;-0.0%")}&"，估值評等 "&{v("rating")}&"，投資建議 "&{v("rec")}&"。"'),
        (f'="DCF 目標價 "&{S(v("dcf_tp"),"#,##0.0")}&"、Comparable 目標價 "&{S(v("cmp_tp"),"#,##0.0")}&"（各半）；WACC "&{S("wacc","0.00%")}&"、ke "&{S("ke","0.00%")}&"、g "&{S("g_tv","0.00%")}&"；情境 B 終值占 EV "&{S(f"Val!C{VR['tvs']}","0%")}&"。"'),
        (f'="上檔空間歸因：DCF 貢獻 "&{S(v("att_d"),"+0.0%;-0.0%")}&"、未調整 Comparable 貢獻 "&{S(v("att_c"),"+0.0%;-0.0%")}&"、相對倍數調整貢獻 "&{S(v("att_r"),"+0.0%;-0.0%")}&"。本公司現價對資料商前瞻 EPS 的倍數是同業中位數的 "&{S(v("relnow"),"0.00")}&" 倍；支持現價需要相對倍數係數 "&{S(v("relreq"),"0.00")}&"，升一級需要 "&{S(v("relflip"),"0.00")}&"。"'),
        (f'="收支關卡：情境 B FY2026–FY2030 FCFF 扣股利合計 "&{S(v("gate_b"),"#,##0.0")}&" 百萬元，"&{v("gate_t")}&"；距最近級距邊界 "&{S(v("dist"),"0.0%")}&"。"'),
        "主要失效條件（可觀察）：(1) 新能源車齒輪年增低於 10%（客戶電驅自製回流或終端需求轉弱）；(2) 毛利率跌破 25%（原材料與海外廠爬坡）；(3) 資本支出不回落、自由現金流持續為負；(4) RV 或人形用傳動件取得具名大客戶，營收揭露明顯放大。",
        "方法：DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，WACC 折現，明確期到 FY2032，RONIC＝WACC；金融資產以淨現金計入）50%＋同業前瞻 P/E × 模型 FY2027 EPS 50%；三情境 25／50／25。",
        "限制：FY2023 只取得營收與淨利；RV 營收、人形營收公司未揭露（RV 占比為 Assumed）；2026-09-30 資產負債表未取得，淨現金用 2025 年底；三季報為轉錄；同業 P/E 由原報價日按 2026-11-20 收盤等比例換算（EPS 未更新）且同業含機器人選擇權定價；不讀 Siliconomics／Tokenomics／Wattonomics。詳見 repo README.md。",
    ]
    for i, t in enumerate(lines):
        RD.cell(1 + i, 1, t).alignment = WRAP
    RD.column_dimensions["A"].width = 160
    return wb, dict(VR=VR, KR=KR, PR=PR, LAST_IN=LAST_IN, THR=THR, DR=DR)


# ---------------- 建置、重算、讀值 ----------------
def recalc(paths, outdir):
    os.makedirs(outdir, exist_ok=True)
    for attempt in range(2):
        try:
            subprocess.run(LO + ["--outdir", outdir] + paths, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
            return
        except subprocess.TimeoutExpired:
            subprocess.run(["pkill", "-f", "lo_profile"], check=False)
            if attempt == 1: raise


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
    out = []
    for i in range(0, len(paths), 20):
        recalc(paths[i:i + 20], rc)
    return [read_val(os.path.join(rc, f"c{i:03d}.xlsx"), meta) for i in range(len(cases))]


def in_values(path, last):
    wb = openpyxl.load_workbook(path, data_only=True); ws = wb["In"]
    out = {}
    for r in range(4, last + 1):
        k = ws.cell(r, 1).value
        if isinstance(k, str) and ws.cell(r, 5).value is not None:
            out[k] = (ws.cell(r, 4).value, ws.cell(r, 5).value, ws.cell(r, 6).value, ws.cell(r, 2).value)
    return out



def main():
    os.makedirs(os.path.join(HERE, "dist"), exist_ok=True)
    shutil.rmtree(TMP, ignore_errors=True); os.makedirs(TMP)
    wb, meta = make(); base_raw = os.path.join(TMP, "base.xlsx"); wb.save(base_raw)
    recalc([base_raw], os.path.join(TMP, "base_rc"))
    base_rc = os.path.join(TMP, "base_rc", "base.xlsx")
    base = read_val(base_rc, meta)
    iv = in_values(base_rc, meta["LAST_IN"])
    fp = openpyxl.load_workbook(base_rc, data_only=True)["Sens"]["D5"].value
    print("基準：", base)
    if "--quick" in sys.argv:
        return

    specs = []
    for k, (lo, e, hi, lab) in iv.items():
        for side, val in (("敏感度低", lo), ("敏感度高", hi)):
            if isinstance(val, (int, float)) and isinstance(e, (int, float)) and abs(val - e) > 1e-12:
                specs.append(dict(drv=f"{lab}（{side}）", set=f"{k}＝{val:.4g}（基準 {e:.4g}）", key=k, val=val, ov={k: val}))
    ivE = {k: v[1] for k, v in iv.items()}
    extra = []
    for k, lab in (("nevg", "新能源車齒輪 FY2027 年增"), ("gfuel", "燃油乘用車齒輪年增"), ("gact", "智能執行機構年增"), ("rvd", "RV 份額漂移"), ("gm30", "FY2030 毛利率")):
        extra.append((f"情境 B 的{lab}改為情境 A 值", f"{k}_B", ivE[f"{k}_A"]))
        extra.append((f"情境 B 的{lab}改為情境 C 值", f"{k}_B", ivE[f"{k}_C"]))
    for drv, k, val in extra:
        specs.append(dict(drv=drv, set=f"{k}＝{val:.4g}", key=k, val=val, ov={k: val}))
    specs.append(dict(drv="情境權重全給 A", set="tp_wA＝1、tp_wB＝0、tp_wC＝0", key="tp_wA tp_wB tp_wC", val=None, ov={"tp_wA": 1, "tp_wB": 0, "tp_wC": 0}))
    specs.append(dict(drv="情境權重全給 C", set="tp_wA＝0、tp_wB＝0、tp_wC＝1", key="tp_wA tp_wB tp_wC", val=None, ov={"tp_wA": 0, "tp_wB": 0, "tp_wC": 1}))
    res = run_cases([s["ov"] for s in specs], meta, "sens")
    cases = []
    for s, r in zip(specs, res):
        cases.append(dict(drv=s["drv"], set=s["set"], key=s["key"], val=s["val"], tp=r["tp"], dcf=r["dcf"], rating=r["rating"]))
        print(f"  {s['drv'][:44]:44s} {r['tp']:9.1f} {r['tp'] / base['tp'] - 1:+.1%}")

    # 反解（整本重算，二分）：求加權目標價等於升一級／降一級邊界所需的值
    px = ivE["px"]
    def solve(key, target_tp, lo, hi, n=12, increasing=True):
        f = lambda x: run_cases([{key: x}], meta, "solve")[0]["tp"]
        flo, fhi = f(lo), f(hi)
        if not ((flo <= target_tp <= fhi) if increasing else (fhi <= target_tp <= flo)): return None
        a, b = lo, hi
        for _ in range(n):
            m = (a + b) / 2
            if (f(m) < target_tp) == increasing: a = m
            else: b = m
        return (a + b) / 2
    up = base["up"]
    up_edge = 0.5 if up > 0.25 else 0.25 if up > -0.25 else -0.25 if up > -0.5 else -0.5
    dn_edge = 0.25 if up > 0.5 else -0.25 if up > 0.25 else -0.5 if up > -0.25 else None
    if up > 0.5: up_edge = None
    solve_rows = []
    def add(lab, key, tp_t, lo, hi, note):
        s_ = solve(key, tp_t, lo, hi) if tp_t else None
        solve_rows.append((lab, s_ if s_ is not None else -1, ivE[key], note + ("；−1＝搜尋區間內無解" if s_ is None else "")))
    add("支持現價（目標價＝現價）所需的相對倍數係數（同業 P/E 的比例）", "rel_k", px, 0.1, 3.0, "其他輸入不變；基準 1.00＝同業中位數")
    if up_edge is not None:
        add(f"評等升一級（上檔 {up_edge:+.0%}）所需的情境 B FY2030 毛利率", "gm30_B", px * (1 + up_edge + 1e-6), 0.15, 0.40, "其他輸入不變")
        add(f"評等升一級（上檔 {up_edge:+.0%}）所需的情境 B 新能源車齒輪 FY2027 年增", "nevg_B", px * (1 + up_edge + 1e-6), 0.0, 0.8, "其他輸入不變")
    if dn_edge is not None:
        add(f"評等降一級（上檔 {dn_edge:+.0%}）所需的相對倍數係數", "rel_k", px * (1 + dn_edge - 1e-6), 0.1, 3.0, "其他輸入不變")
        add(f"評等降一級（上檔 {dn_edge:+.0%}）所需的情境 B FY2030 毛利率", "gm30_B", px * (1 + dn_edge - 1e-6), 0.15, 0.40, "其他輸入不變")
        add(f"評等降一級（上檔 {dn_edge:+.0%}）所需的情境 B 新能源車齒輪 FY2027 年增", "nevg_B", px * (1 + dn_edge - 1e-6), 0.0, 0.8, "其他輸入不變")
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
    wbv = openpyxl.load_workbook(OUT, data_only=True)
    Hs = wbv["Hist"]
    segs = [("乘用車齒輪", "乘用車齒輪 營收", Hs["E44"].value), ("商用車齒輪", "商用車齒輪 營收", Hs["E47"].value), ("工程機械齒輪", "工程機械齒輪 營收", Hs["E48"].value),
            ("電動工具與摩托車齒輪", "電動工具與摩托車齒輪 營收", Hs["E49"].value), ("智能執行機構", "智能執行機構 營收", Hs["E50"].value),
            ("減速器及其他", "減速器及其他 營收", Hs["E51"].value), ("鋼材銷售", "鋼材銷售 營收", Hs["E52"].value)]
    json.dump(dict(company="shuanghuan", ticker="002472.SZ", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=round(v, 2), unit="百萬元", source=S_AR) for n, l, v in segs],
                   sublines=[dict(parent="乘用車齒輪", labels=["乘用車：新能源車齒輪（子列）", "乘用車：燃油乘用車齒輪（子列）"]),
                             dict(parent="減速器及其他", labels=["減速器：RV 機器人減速機（子列）", "減速器：其他減速器與雜項（子列）"])],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）", "檢查：子列合計 − 分部（應為 0）"],
                   note="官方分部＝2025 年年報分產品營業收入（乘用車齒輪、電動工具齒輪、工程機械齒輪、摩托車齒輪、商用車齒輪、智能執行機構、減速器及其他、鋼材銷售）；電動工具與摩托車合併為一列。新能源車齒輪（年報另列 41.02 億）為乘用車齒輪的子列；RV 機器人減速機為減速器及其他的子列（公司未揭露，Assumed 拆分）。"),
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
        p = subprocess.run([sys.executable, os.path.join(ROOT, cmd[0])] + cmd[1:], cwd=ROOT)
        rc |= p.returncode
    if errs: rc |= 1
    sys.exit(rc)


def write_readme(path, meta, sens):
    """repo README.md 由成品重算值產生（原則第 10 節第 51 條：不手寫數字）。"""
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; VR = meta["VR"]
    g = lambda k, c="B": V[f"{c}{VR[k]}"].value
    ws = wb["In"]; I = {}
    for r in range(4, meta["LAST_IN"] + 1):
        if ws.cell(r, 1).value: I[ws.cell(r, 1).value] = ws.cell(r, 5).value
    PL = wb["PL"]; PR = meta["PR"]["B"]
    ser = lambda k, f: " → ".join(f.format(PL[f"{c}{PR[k]}"].value) for c in "FGHIJKL")
    top = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base_tp"] - 1))[:10]
    TH = wb["Thesis"]; THB = meta["THR"]["B"]
    fn = lambda x, f="{:,.1f}": f.format(x) if isinstance(x, (int, float)) else str(x)
    sol = sens["solve"]
    t = f"""# physical_ai/shuanghuan/ — 雙環傳動（002472.SZ）個股模型 {VER}

<!-- 本檔由 build_shuanghuan.py 依成品重算值產生；數字不手寫（原則第 10 節第 51 條）。 -->

**本模型適用《高不確定性預估工作原則》**（`docs/plan/高不確定性預估工作原則.md`）：RV 機器人減速機營收與人形用量公司未揭露，相關輸入一律給區間並標 Assumed／Analogy；結論只在證據邊界內成立。

## 未完成項目與限制（先看這裡）

- FY2023 只取得營收、淨利、EPS、營業現金流（年報主要財務指標表）；FY2023 的分部、成本、費用未取得，PL 該欄留空。
- 2026-09-30 資產負債表未取得（資料商頁面無法讀取、三季報原文未讀），淨現金以 2025-12-31 年報為準，Stub 期含全年 FY2026 FCFF。
- 三季報、半年報數字經券商與新聞轉錄；公司原文只讀了 2025 年年報。
- RV 機器人減速機營收、人形用營收公司未揭露：RV 占「減速器及其他」60%（區間 40–80%）為 Assumed；線 6 顯示 RV 不因人形增量，本模型不另列人形營收，只在對帳頁列線 4 出貨作對照。
- 券商營收預測未取得（只取得淨利預測）；員工持股計畫、限制性股票之稀釋未取得。
- 同業前瞻 P/E 取自 leaderdrive 模型 Comp 頁（原報價日 2026-09-30～10-30，以 2026-11-20 收盤等比例換算，EPS 未更新）；核心同業 4 家（拓普、Nabtesco、哈默納科、兆威），其餘因 P/E 逾 100 倍、無前瞻值或業務不同未進中位數。這些同業多含機器人選擇權定價，與本公司汽車齒輪主體口徑不同。
- 獨立核對：交付判讀的數字應由未參與製作的代理獨立核對；本次未做。

- 成品：`dist/{NAME}`（build_id `{sens['build_id']}`；LibreOffice 重算，0 錯誤值）；HTML 展示層 `dist/{NAME[:-5]}.html`（`gen_html_json.py` 產生 `html.json`，再以 `tools/company_html/build_company_html.py` 產生）
- 建置：`python3 physical_ai/shuanghuan/build_shuanghuan.py`（需要 soffice；整本重算敏感度與反解；結尾自動跑 in_layout、sens_coverage --require-keycol、model_governance、rating_gate --strict、zero_rows、cached_values、segment_check）
- 日期：資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 2026-11-20（收盤 {I['px']:,.2f} 元）｜評價日 2026-12-31｜目標時點 2027-11-20
- 會計年度＝日曆年；單位百萬元人民幣；中國會計準則（法定財報口徑，GAAP）。深圳證券交易所。

## 結論（條件式）

在「新能源車齒輪 FY2027 年增 {I['nevg_B']:.0%}（逐年衰減 {I['nev_fade']:.2f}）、燃油乘用車齒輪年降 {-I['gfuel_B']:.0%}、FY2030 毛利率 {I['gm30_B']:.1%}、資本支出占營收 {I['cx_r']:.0%}、同業 P/E 不加溢價（相對倍數係數 {I['rel_k']:.2f}）」的假設下：

| 項目 | 值 |
|---|---|
| 加權目標價（DCF 50%＋Comparable 50%，情境 25／50／25） | **{g('tp'):,.1f} 元** |
| DCF 目標價／Comparable 目標價 | {g('dcf_tp'):,.1f}／{g('cmp_tp'):,.1f} |
| 相對現價 {I['px']:,.2f} | **{g('up'):+.1%}** |
| 估值評等／投資建議 | **{g('rating')}／{g('rec')}** |
| 收支關卡（情境 B FY2026–FY2030 FCFF 扣股利合計） | {g('gate_b'):,.1f} 百萬元（{g('gate_t')}） |
| 情境價值（DCF A／B／C） | {g('dcf','B'):,.1f}／{g('dcf','C'):,.1f}／{g('dcf','D'):,.1f} |
| 情境價值（Comparable A／B／C） | {g('cmp','B'):,.1f}／{g('cmp','C'):,.1f}／{g('cmp','D'):,.1f} |
| 情境目標價（A／B／C） | {g('wt','B'):,.1f}／{g('wt','C'):,.1f}／{g('wt','D'):,.1f} |
| WACC／ke／g／β | {g('wacc'):.2%}／{g('ke'):.2%}／{I['g_tv']:.2%}／{I['beta']:.2f} |
| 終值占 EV（情境 B） | {g('tvs','C'):.0%} |
| 賣方目標價平均（同花順 12 家，2026-09；只作對照） | {I['cons_tp']:,.2f} |

上檔空間歸因：DCF 貢獻 {g('att_d'):+.1%}、未調整 Comparable 貢獻 {g('att_c'):+.1%}、相對倍數調整 {g('att_r'):+.1%}。本公司現價對資料商前瞻 EPS 的倍數是核心同業中位數（{I['pe_peer']:.1f} 倍）的 {g('relnow'):.2f} 倍；要讓目標價＝現價，Comparable 的相對倍數係數須達 {g('relreq'):.2f}。同業含機器人選擇權定價，本公司主體是汽車齒輪，Comparable 一半的權重因此偏向高估；DCF 才是對自身現金流的直接檢驗。

**讀法：** 情境 B 營收 {ser('rev', '{:,.0f}')}（FY2026→FY2032，百萬元）；新能源車齒輪占合併營收由 {TH[f'D{THB["nevsh"]}'].value:.0%}（FY2025）升到 {TH[f'K{THB["nevsh"]}'].value:.0%}（FY2032）。結論主要建立在 Assumed 輸入（各分部年增、毛利率、資本支出路徑）上，Verified 的部分是 FY2024–FY2025 年報與 2026 年第 3 季實績；最敏感的輸入見下表與 Recon 反解。

## 命題

1. **營收結構：** 官方分部為乘用車齒輪（2025 年 {wb['Hist']['E44'].value:,.0f} 百萬元，占 {wb['Hist']['E44'].value / wb['Hist']['E4'].value:.0%}；其中新能源車齒輪 {wb['Hist']['E45'].value:,.0f}）、商用車、工程機械、電動工具與摩托車、智能執行機構、減速器及其他；各分部由年增推導。
2. **新能源車齒輪：** 外包趨勢與海外廠帶動，2025 年 +21.7%；模型首年增（FY2027）情境 A／B／C＝{I['nevg_A']:.0%}／{I['nevg_B']:.0%}／{I['nevg_C']:.0%}，之後逐年衰減。標記 Assumed。
3. **機器人減速機（RV）：** 營收＝前一年 × 線 6 RV 市場成長 ×（1＋份額漂移）；線 6 RV 市場 2026→2030 基準幾乎持平（只來自傳統工業機器人），人形不增量，因此不是成長來源。
4. **獲利與現金：** 毛利率 FY2026 {I['gm26']:.1%}、FY2030 情境 A／B／C＝{I['gm30_A']:.1%}／{I['gm30_B']:.1%}／{I['gm30_C']:.1%}；資本支出 2024–2025 年占營收約 22%，FY2027 起假設降到 {I['cx_r']:.0%}，這是自由現金流與 DCF 的關鍵假設。

情境 B：營收 {ser('rev', '{:,.0f}')}；毛利率 {ser('gm', '{:.1%}')}；EPS {ser('eps', '{:.2f}')} 元（FY2026→FY2032）。

## 主要假設與可證偽指標（原則第 10 節第 20 條）

| 承重假設 | 基準（A／B／C 或區間） | 標記 | 可觀察指標 |
|---|---|---|---|
| 新能源車齒輪 FY2027 年增 | {I['nevg_A']:.0%}／{I['nevg_B']:.0%}／{I['nevg_C']:.0%}，衰減 {I['nev_fade']:.2f}（0.6–0.9） | Assumed | 季報新能源車齒輪占比（Q3 已逾 50%）、分部年增 |
| FY2030 毛利率 | {I['gm30_A']:.1%}／{I['gm30_B']:.1%}／{I['gm30_C']:.1%} | Assumed | 季毛利率（2026 第 3 季 26.78%） |
| 資本支出占營收（FY2027 起） | {I['cx_r']:.0%}（10–20%） | Assumed | 年報購建支出、在建工程（2025 年底 2,182） |
| FY2026 新能源車齒輪年增 | {I['g_nev26']:.0%}（16–30%） | Assumed | 第 4 季營收（模型隱含年增見 Recon） |
| RV 占減速器及其他 | {I['rv_sh26']:.0%}（40–80%） | Assumed | 環動科技揭露、年報分產品拆分 |
| 同業前瞻 P/E | 中位數 {I['pe_peer']:.1f} 倍（4 家） | Derived | 同一估值日重抓 |

## 敏感度（整本重算，前 10 大）

| 案例 | 鍵 | 加權目標價 | 相對基準 | DCF 目標價 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | `{c['key']}` | {c['tp']:,.1f} | {c['tp'] / sens['base_tp'] - 1:+.1%} | {c['dcf']:,.1f} | {c['rating']} |" for c in top) + f"""

完整案例見 xlsx Sens 頁（每個 In 頁有低／高區間的鍵都有案例；另有情境 B 改 A／C 值與情境權重案例）。

## 反解（Recon 頁）

""" + "\n".join(f"- {a}：{fn(b, '{:,.3f}')}（基準 {fn(c0, '{:,.3f}')}）；{d}" for a, b, c0, d in sol) + f"""

## 其他限制

- **線模型不一致：** 無。線 6 RV 市場規模直接讀值（中間年內插）；線 4 人形出貨只作對照。見 xlsx「對帳」頁。
- **rf 逾期：** 中國 10 年期國債 1.61% 為 2026-11-05 值（跨模型一致）。
- **協調檔：** 實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics；相關條文不適用。

## 資料來源

見 xlsx Sources 頁（公司 2025 年報、三季報轉錄、資料商、線 4／6 成品）。

## 版本紀錄

- {VER}（2026-11-24）：首版。現價 {I['px']:,.2f}（2026-11-20 收盤）；加權目標價 {g('tp'):,.1f}；上檔空間 {g('up'):+.1%}；{g('rating')}／{g('rec')}；收支關卡 {g('gate_b'):,.1f}（{g('gate_t')}）。
"""
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write(t)


if __name__ == "__main__":
    main()

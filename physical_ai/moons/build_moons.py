# 鳴志電器（603728.SS，上交所）個股模型 v0.1 建置腳本
# 產出：physical_ai/moons/dist/20261124_鳴志電器個股模型_v0.1.xlsx（LibreOffice 重算後的成品，含公式與快取值）
# 結構：命題模組（Thesis：控制電機及驅動系統＝非機器人用＋機器人應用（非人形）＋人形空心杯用（線 4 出貨 × 線 6 每台用量 × 份額 × 單價））
#       → 官方分部（控制電機及其驅動系統／其他板塊）→ 標準損益表 → FCFF DCF＋同業 P/E。
# 敏感度：每個案例覆寫 In 頁基準欄後整本重建、LibreOffice 重算（整本重算，不用閉式近似）；反解同法以二分搜尋。
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md、docs/plan/高不確定性預估工作原則.md。
#       實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics 協調檔。
# 用法：python3 physical_ai/moons/build_moons.py   （需要 soffice；暫存於 $HOME/mn_build）
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
NAME = f"20261124_鳴志電器個股模型_{VER}.xlsx"
OUT = os.path.join(HERE, "dist", NAME)
HOMEDIR = os.path.expanduser("~")
TMP = os.path.join(HOMEDIR, "mn_build")
LO = ["soffice", f"-env:UserInstallation=file://{HOMEDIR}/lo_moons", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML"]

L4 = os.path.join(ROOT, "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx")
L6 = os.path.join(ROOT, "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx")
L4N, L6N = (os.path.relpath(p, ROOT) for p in (L4, L6))

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDRF = PatternFill("solid", fgColor="DDEBF7")
SECF = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------- 出處 ----------------
S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv：603728.SS；未還原股利）"
S_Q3 = "鳴志電器 2026 年第三季度報告（上交所公告，2026-10-28，未經審計；https://stockmc.xueqiu.com/202610/603728_20261028_108A.pdf）：合併資產負債表、利潤表、現金流量表原文"
S_AR = "鳴志電器 2025 年年度報告摘要（公司公告，2026-04-24；https://pdf.dfcfw.com/pdf/H2_AN202604241821560097_1.pdf）：近三年主要會計數據、分季度數據"
S_HA = "華安證券 鳴志電器公司點評（2026-11-05；https://pdf.dfcfw.com/pdf/H3_AP202611081831535216_1.pdf）：上半年分產品、機器人應用占比、三季報後預測"
S_QY = "群益證券 鳴志電器研報（2026-10-28 前後；https://data.eastmoney.com/report/info/AP202610281831089523.html）：三季報點評與預測"
S_SA = "stockanalysis.com SHA:603728 財務與現金流量表頁（資料商，2026-11-24 讀取；核心營業利益、毛利、折舊攤銷、資本支出、股利為資料商整理值）"
S_NEWS = "網易財經 鳴志電器 2025 年營收 27.62 億（https://www.163.com/money/article/KSTF9B7P00259SAH.html）與搜狐、新浪財經年報解讀（2026-04-25）：轉錄 2025 年報控制電機及驅動系統收入 22.93 億、年增 17.04%、毛利率 36.8%，研發費用 2.69 億"
S_SA_ST = "stockanalysis.com SHA:603728 statistics（資料商；β 與前瞻 P/E；原報價日見 Comp 頁）"
S_RF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與 leaderdrive、innolight 等同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：China CRP 0.91%（與 innolight 同一取值）"
S_CLOSE = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
S_SAP = "stockanalysis.com statistics 頁（資料商；原報價日見 Comp 頁 M 欄）"
IP_BRK = "券商研究報告：有銷售研究與維持評等的誘因；歷史數字為轉錄，預測為券商自設"
IP_DATA = "資料商整理值，非公司原始文件"

# ---------------- 線模型取數（以 ID／代號找列，不以列號） ----------------
def _line_rows(path, idcol, sheet="下游取數"):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    out = {}
    for r in range(1, ws.max_row + 1):
        k = ws.cell(r, idcol).value
        if isinstance(k, str): out[k] = (r, [ws.cell(r, c).value for c in range(1, 10)])
    return wb, out




def line_data():
    wb4, l4 = _line_rows(L4, 1)
    wb6, _ = _line_rows(L6, 1)
    d = {"hum": {}}
    for y in range(2026, 2033):
        r, v = l4[f"SHIP_{y}"]
        for s, c, col in (("A", 4, "E"), ("B", 3, "D"), ("C", 2, "C")):
            d["hum"][(y, s)] = (v[c], f"{L4N} 下游取數!{col}{r}（SHIP_{y}）")
    ws6 = wb6["輸入"]
    cm = [r for r in range(1, ws6.max_row + 1) if ws6.cell(r, 1).value == "CM"]
    units = [r for r in cm if ws6.cell(r, 3).value == "顆"]          # 低、基準、高三個矩陣
    pr = [r for r in cm if ws6.cell(r, 3).value == "美元/顆"][0]
    assert len(units) == 3
    d["n_cup"] = [(ws6.cell(r, 6).value, f"{L6N} 輸入!F{r}") for r in units]
    d["pd"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{openpyxl.utils.get_column_letter(c)}{pr}") for c in (7, 8, 9)]
    d["p26"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{openpyxl.utils.get_column_letter(c)}{pr}") for c in (4, 5, 6)]
    d["mkt30"] = (wb6["計算"].cell(146, 7).value, f"{L6N} 計算!G146")
    return d


LD = line_data()

# ---------------- In 頁輸入 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 來源誘因, 說明)
IN = [
    ("sec", "1. 市場資料與日期"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, "", "資料截止日 2026-11-24（財報、券商預測、同業倍數）；股價 2026-11-20；成品日 2026-11-24"),
    ("px", "現價（2026-11-20 收盤）", "元", None, 44.52, None, "Verified", S_PX, "", "11-09 至 11-20 全球 AI 與半導體股大跌，收盤檔下跌非資料錯誤；華安報告 11-05 收盤 60.49"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,12,31)", None, "Assumed", "本專案慣例：評價日 2026-12-31（會計年度＝日曆年，FY2026 在評價日已結束）", "", "折現期間由評價日起算"),
    ("tp_date", "目標時點（px_date 後 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", ""),
    ("bs_date", "資產負債表日", "日期", None, "=DATE(2026,9,30)", None, "Verified", S_Q3, "", "2026 年三季報"),
    ("stub_frac", "Stub 期（資產負債表日→評價日）", "年", None, "=(val_date-bs_date)/365", None, "Derived", "＝(評價日 − 資產負債表日) ÷ 365", "", "Stub FCFF＝FY2026 FCFF × stub_frac"),
    ("sh_v", "評價日股數（總股本）", "百萬股", None, 418.8825, None, "Verified", S_Q3 + "：實收資本 418,882,500 元（每股面值 1 元）", "", "公司另有可轉債、期權未見於三季報摘錄，以 0 計"),
    ("cash_v", "貨幣資金＋交易性金融資產（2026-09-30）", "百萬元", None, 880.29, None, "Verified", S_Q3 + "：貨幣資金 880.13、交易性金融資產 0.16", "", "2025-12-31 貨幣資金 871.0；部分為受限資金，明細未找到"),
    ("debt_v", "有息負債（2026-09-30）", "百萬元", None, 803.51, None, "Verified", S_Q3 + "：短期借款 524.0＋一年內到期非流動負債 44.4＋長期借款 220.9＋租賃負債 14.2", "", "2025-12-31：短期借款 420、長期借款 112"),
    ("nci_v", "少數股東權益（2026-09-30）", "百萬元", None, 14.05, None, "Verified", S_Q3, "", "自淨現金扣除"),
    ("fx", "美元兌人民幣（2026-11-20）", "元／美元", None, 6.6665, None, "Verified", S_CLOSE + "：CNY=X", "", "用於線 6 單價（美元）換算人民幣"),
    ("sec", "2. 實績錨（FY2026＝2026 年，會計年度＝日曆年）"),
    ("g9m26", "2026 年 1–9 月營收年增", "%", None, "=Q!D4/Q!E4-1", None, "Derived", "＝Q 頁 2026 年 1–9 月營收 ÷ 2025 年 1–9 月營收 − 1（與公告 +21.40% 一致）", "", ""),
    ("q4g", "2026 年第 4 季營收年增", "%", 0.10, 0.20, 0.30, "Assumed", "第 3 季 +25.3%、1–9 月 +21.4%（" + S_Q3 + "）；2025 年第 4 季 718.9 為基期；華安 2026E 營收 3,375 隱含第 4 季年增約 +26%", "", "FY2026 營收＝1–9 月實績＋2025Q4 ×（1＋q4g）"),
    ("sh_cm26", "FY2026 控制電機及驅動系統 占營收", "%", None, 0.8815, None, "Interested-party", S_HA + "：2026 上半年控制電機及驅動系統 13.81 億元、占營收 88.15%", IP_BRK, "全年占比以上半年代入；其餘為電源與照明控制、設備狀態管理、貿易"),
    ("r_sh26", "FY2026 機器人應用 占營收", "%", 0.07, 0.09, 0.12, "Interested-party", S_HA + "：上半年機器人應用收入年增約 58%、占營收約 9%（絕對值公司未單獨披露）", IP_BRK, "機器人應用含人形與非人形（物流機器人、醫療等）；公司未定義口徑"),
    ("hum_sh26", "FY2026 機器人應用中人形空心杯用占比", "%", 0.03, 0.10, 0.25, "Assumed", "公司未揭露人形別營收；已與百餘家靈巧手與整機企業合作（" + S_QY + "，" + S_HA + "）；基準 10% 隱含 2026 全球人形空心杯市場份額約 14%（Thesis）", "", "反映合作多但量產有限的判斷，無直接證據"),
    ("gm26", "FY2026 毛利率", "%", None, "=Q!D7", None, "Derived", "＝2026 年 1–9 月毛利率 38.13%（Q 頁；" + S_Q3 + "）", "", "第 4 季假設同 1–9 月（第 3 季 38.87%）"),
    ("opex26", "FY2026 營業費用率（含稅金及附加、其他營業損益；＝毛利率 − 核心營業利益率）", "%", 0.32, "=(Q!D6-Q!D15)/Q!D4", 0.35, "Derived", "＝(1–9 月毛利 − 核心營業利益) ÷ 營收（Q 頁）；核心營業利益不含財務費用與投資收益", "", "敏感度 32%／35%（Assumed）；2025 年 34.4%、2024 年 34.3%"),
    ("sec", "3. 命題模組：人形空心杯（量 × 價；線 4、線 6）"),
    ("p_cm", "空心杯 2026 基準單價（線 6）", "美元／顆", LD["p26"][0][0], LD["p26"][1][0], LD["p26"][2][0], "Interested-party", f"線 6：{LD['p26'][1][1]}（低 {LD['p26'][0][1]}、高 {LD['p26'][2][1]}）；線 6 引述國產低階 100–4,000 人民幣（TechBuzz China）", "線 6 引述產業報告與新聞，非公司原始文件", "全球混合單價；本公司實際單價未揭露。單價水準只影響「隱含份額」的換算，不影響 2026 人形營收錨"),
    ("pd", "空心杯單價年降幅（線 6）", "%／年", LD["pd"][0][0], LD["pd"][1][0], LD["pd"][2][0], "Analogy", f"線 6：{LD['pd'][1][1]}（低 {LD['pd'][0][1]}、高 {LD['pd'][2][1]}）", "", "國產空心杯以價格切入"),
    ("n_cup", "人形每台空心杯用量（線 6）", "顆／台", LD["n_cup"][0][0], LD["n_cup"][1][0], LD["n_cup"][2][0], "Interested-party",
     f"線 6：{LD['n_cup'][1][1]}（低 {LD['n_cup'][0][1]}、高 {LD['n_cup'][2][1]}）", "線 6 輸入為外部報告與拆解推估，非公司原始文件", "取決於靈巧手自由度"),
    ("hu_k", "人形出貨相對線 4 的倍數（全部情境）", "x", None, 1.0, 2.0, "Assumed", "基準＝線 4 原值；高案 2 倍（中國 2026 年人形產量可能高於線 4 全球基準；未查核）", "", "反解變數"),
    ("g_o", "其他板塊（電源與照明、設備狀態管理、貿易）年增（FY2027 起）", "%", 0.0, 0.05, 0.10, "Assumed", "上半年控制電機及驅動系統 +25.7%，其餘約占 12%，成長未揭露", "", "三情境共用，Sens 動低高"),
    ("gm_adj", "FY2030 毛利率整體調整（三情境同加；基準 0）", "pt", -0.03, 0.0, 0.03, "Assumed", "敏感度與反解用：同時平移三情境 gm30", "", "反解變數"),
    ("sec", "4. 命題模組：情境專屬驅動（A 樂觀／B 基準／C 保守；只列基準欄，不設敏感度區間）"),
    ("gind_A", "情境 A：控制電機非機器人用營收年增（FY2027 起）", "%", None, 0.14, None, "Assumed", "2025 年控制電機及驅動系統 +17.0%、2026 年 1–9 月總營收 +21.4%；海外去庫完成與液冷、數據中心帶動", "", ""),
    ("gind_B", "情境 B：控制電機非機器人用營收年增（FY2027 起）", "%", None, 0.10, None, "Assumed", "同上；華安 2027E／2028E 營收 +18%／+14%（含機器人）", "", ""),
    ("gind_C", "情境 C：控制電機非機器人用營收年增（FY2027 起）", "%", None, 0.04, None, "Assumed", "同上；2024 年全公司營收 −5%（週期下行）", "", ""),
    ("grob_A", "情境 A：機器人應用（非人形）營收年增（FY2027 起）", "%", None, 0.35, None, "Assumed", "上半年機器人應用年增約 58%（Interested-party）", "", ""),
    ("grob_B", "情境 B：機器人應用（非人形）營收年增（FY2027 起）", "%", None, 0.25, None, "Assumed", "同上，基期小，成長遞減", "", ""),
    ("grob_C", "情境 C：機器人應用（非人形）營收年增（FY2027 起）", "%", None, 0.12, None, "Assumed", "同上", "", ""),
    ("hs30_A", "情境 A：2030 年占全球人形空心杯用量份額", "%", None, 0.20, None, "Assumed", "2026 隱含份額見 Thesis（基準約 14%）；國產空心杯價格競爭激烈，份額線性走到 2030 年再持平", "", ""),
    ("hs30_B", "情境 B：2030 年占全球人形空心杯用量份額", "%", None, 0.10, None, "Assumed", "同上", "", ""),
    ("hs30_C", "情境 C：2030 年占全球人形空心杯用量份額", "%", None, 0.05, None, "Assumed", "同上；靈巧手整機廠自製馬達或轉用其他驅動（" + "線 6 份額頁、profile 提示" + "）", "", ""),
    ("gm30_A", "情境 A：FY2030 毛利率", "%", None, 0.405, None, "Assumed", "2023 年 36.2%、2024 年 36.6%、2025 年 36.4%、2026 年 1–9 月 38.1%（Hist、Q）；華安預測 2028 年 38.9%", "", "FY2026 起線性走到 FY2030，其後持平"),
    ("gm30_B", "情境 B：FY2030 毛利率", "%", None, 0.385, None, "Assumed", "同上", "", ""),
    ("gm30_C", "情境 C：FY2030 毛利率", "%", None, 0.355, None, "Assumed", "同上；回到 2025 年以下", "", ""),
    ("tp_wA", "情境權重 A", "", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", ""),
    ("tp_wB", "情境權重 B", "", None, 0.50, None, "Assumed", "同上", "", ""),
    ("tp_wC", "情境權重 C", "", None, 0.25, None, "Assumed", "同上", "", ""),
    ("sec", "5. 損益與現金流"),
    ("opex30", "FY2030 營業費用率（含稅金及附加、其他營業損益）", "%", 0.285, 0.305, 0.33, "Assumed", "2023 年 30.8%、2025 年 34.4%、2026 年 1–9 月 33.3%；華安 2028E 營業費用（銷售＋管理＋財務）約 22.5% 加研發約 9% 約 31.5%", "", "FY2026→FY2030 線性，其後持平"),
    ("fin_r", "淨財務費用 ÷ 有息負債（含匯兌）", "%", 0.025, 0.035, 0.045, "Assumed", S_Q3 + "：1–9 月財務費用 21.1 百萬元（年化約 28）÷ 有息負債約 804＝3.5%", "", "業外＝−淨財務費用率 × 有息負債；2026 年匯兌損失放大"),
    ("tax_r", "有效稅率", "%", 0.03, 0.08, 0.15, "Assumed", "2023 年約 5%、2024 年約 −3%、2025 年 −34%（所得稅 −16，稅前 47）、2026 年 1–9 月 −1.5%（" + S_AR + "、" + S_Q3 + "）：高新技術企業 15% 與研發加計扣除、海外稅率差異", "", "歷史有效稅率極低且波動；基準 8%，高案 15% 為法定高新稅率"),
    ("nci_r", "少數股東損益 ÷ 稅後淨利", "%", None, "=Hist!E15/Hist!E17", None, "Derived", "＝2025 年少數股東損益 ÷ 稅後淨利（華安 1／63）", "", ""),
    ("payout", "配息率", "%", None, "=Hist!E20/Hist!E19", None, "Derived", "＝2025 年每股股利 0.025 元 ÷ EPS 0.1459 元", "", "用於收支關卡與目標時點前股利"),
    ("da26", "FY2026 折舊攤銷 ÷ 營收", "%", None, "=Hist!E22/Hist!E4", None, "Derived", "＝2025 年折舊攤銷 116.4（資料商；華安 132）÷ 營收", "", "之後線性收斂到資本支出比例 × da_ss（FY2032）"),
    ("da_ss", "穩態折舊 ÷ 資本支出", "x", None, 0.7, None, "Assumed", "穩態資本支出高於折舊（成長）", "", ""),
    ("cx26", "FY2026 資本支出", "百萬元", 250, 290, 330, "Assumed", S_Q3 + "：1–9 月購建固定資產等 220.5（上年同期 220.9）；華安 2026E 262", "", ""),
    ("cx_r", "資本支出 ÷ 營收（FY2027 起）", "%", 0.05, 0.065, 0.09, "Assumed", "2023 年 7.7%、2024 年 12.3%、2025 年 11.5%（Hist，Derived）；華安 2027E 6.6%、2028E 4.8%", "", ""),
    ("nwc_r", "營運資金增加 ÷ 營收增量", "%", 0.15, 0.25, 0.35, "Assumed", S_Q3 + "：2025-12→2026-09 應收 +104.6、存貨 +128.6、應付 +123.0（Derived）；華安 2027E 約 30%", "", "存貨與應收隨營收上升"),
    ("sec", "6. 折現率（DCF 標準規格第 2 節）"),
    ("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", S_RF, "", "日期 2026-11-05，較股價基準日早 11 個交易日（逾期；跨模型一致，下次全面改版更新）"),
    ("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", ""),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", ""),
    ("b_own", "本公司 5 年 β（原始）", "x", None, 0.62, None, "Verified", S_SA_ST + "（2026-10 讀取；與 leaderdrive 模型 Comp 頁同值）", IP_DATA, ""),
    ("b_peer", "同業 5 年 β 中位數（原始）", "x", None, "=MEDIAN(Comp!H5:H10)", None, "Derived", "＝Comp 頁核心同業 6 家 β 中位數", "", ""),
    ("beta", "β（兩來源 Blume 調整後中位數）", "x", "=beta-0.3", "=MEDIAN(0.67*b_own+0.33,0.67*b_peer+0.33)", "=beta+0.3", "Derived", "Blume 0.67β＋0.33；兩來源＝本公司、同業中位數", "", "敏感度 ∓0.3（跨模型慣例）"),
    ("kd", "稅前債務成本", "%", None, "=rf+0.015", None, "Assumed", "rf ＋ 1.5pt（有息負債 8 億元，財務費用率見 fin_r）", "", ""),
    ("wacc_adj", "WACC 調整（敏感度用）", "pt", -0.01, 0, 0.01, "Assumed", "基準 0；敏感度 ±1pt", "", ""),
    ("g_cap", "永續成長率上限", "%", None, 0.03, None, "Assumed", "DCF 標準規格：3%，且不超過計價幣別 10 年期公債殖利率", "", ""),
    ("g_tv", "永續成長率", "%", 0.01, "=MIN(g_cap,rf)", 0.03, "Derived", "＝MIN(g_cap, rf)", "", "敏感度 1%／3%（3% 超過 rf，只作敏感度）"),
    ("ronic_sp", "RONIC − WACC", "pt", None, 0, 0.05, "Assumed", "RONIC 預設＝WACC（新投資價值中性）", "", ""),
    ("tv_alt", "終值改用 FCFF ×（1＋g）（0＝否，1＝是）", "開關", None, 0, 1, "Assumed", "DCF 標準規格第 4 節敏感度", "", ""),
    ("sec", "7. 同業倍數與權重"),
    ("pe_peer", "同業前瞻 P/E 中位數", "x", "=MIN(Comp!F5:F10)", "=MEDIAN(Comp!F5:F10)", "=MAX(Comp!F5:F10)", "Derived", "＝Comp 頁核心同業 6 家前瞻 P/E 中位數；股價皆為 2026-11-20 收盤、P/E 由原報價日等比例換算（EPS 分母不變，見 Comp）", "", "敏感度＝同業最低／最高"),
    ("rel_k", "相對倍數係數（本公司 ÷ 同業）", "x", None, 1.0, 2.0, "Assumed", "1.00（未調整：缺本公司對同業十年相對倍數歷史）", "", "高案 2 倍（未查核，不作依據）"),
    ("w_dcf", "DCF 權重（Comparable＝1 − 本值）", "%", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設 DCF 50%＋同業倍數 50%", "", ""),
]
SCN_DRIVERS = [("gind", "控制電機非機器人用營收年增"), ("grob", "機器人應用（非人形）營收年增"), ("hs30", "2030 年人形空心杯份額"),
               ("gm30", "FY2030 毛利率")]

YRS = list(range(2023, 2033))
COL = {y: openpyxl.utils.get_column_letter(3 + i) for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032
TC = {y: openpyxl.utils.get_column_letter(4 + (y - 2025)) for y in range(2025, 2033)}  # Thesis：D＝2025 … K＝2032
SC = "ABC"
LINE_SC = {"A": "線模型高情境", "B": "線模型基準", "C": "線模型低情境"}

def short(s):
    """Recon 說明欄（HTML 會顯示）：去掉網址與 repo 路徑，改用簡稱。"""
    import re
    s = str(s)
    for p_, n_ in ((L4N, "線 4 成品"), (L6N, "線 6 成品")):
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
    HI["A1"] = "Hist：歷史財務（百萬元人民幣；中國會計準則）。公式只引用本頁（定義名稱 h_*）。核心營業利益＝資料商營業利益，不含財務費用與投資收益（2025：華安營業利潤 48＋財務費用 8＝56 與之相符）。"
    hdr(HI, 3, ["項目", "", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"])
    HROWS = [  # (列, 鍵, 項目, [2023, 2024, 2025], 標記, 來源)
        (4, "h_rev", "營業收入", [2542.79, 2415.93, 2761.98], "Verified", S_AR),
        (5, "h_cogs", "營業成本（＝營業收入 − 毛利）", ["=C4-920.16", "=D4-883.88", "=E4-1005"], "Interested-party", S_SA + "：毛利 920.16、883.88、1,005；2025 全年毛利率 36.4%、2024 年 37.7%（公司年報摘要）"),
        (6, "h_gp", "毛利（＝營收 − 營業成本）", ["=C4-C5", "=D4-D5", "=E4-E5"], "Derived", "＝營業收入 − 營業成本"),
        (7, "h_gm", "毛利率", ["=C6/C4", "=D6/D4", "=E6/E4"], "Derived", "＝毛利 ÷ 營收"),
        (8, "h_seg_cm", "控制電機及其驅動系統 營收", [None, "=E8/1.1704", 2293], "Interested-party", S_NEWS + "（2025：22.93 億，+17.04%；2024＝2025 ÷ 1.1704，Derived；2023 未取得）"),
        (9, "h_seg_o", "其他板塊 營收（電源與照明控制、設備狀態管理、貿易；＝營收 − 控制電機）", [None, "=D4-D8", "=E4-E8"], "Derived", "年報摘要只給合計，三個其他板塊的分項營收未取得，以殘差計"),
        (10, "h_opexall", "營業費用合計（含稅金及附加、其他營業損益；＝毛利 − 核心營業利益）", ["=C6-C11", "=D6-D11", "=E6-E11"], "Derived", "＝毛利 − 核心營業利益（銷售、管理、研發、稅金、減值、其他收益合計）"),
        (11, "h_ebit", "營業利益（核心，不含財務費用與投資收益）", [137.56, 54.44, 56.05], "Interested-party", S_SA),
        (12, "h_ebitm", "營業利益率（核心）", ["=C11/C4", "=D11/D4", "=E11/E4"], "Derived", ""),
        (13, "h_pbt", "利潤總額（稅前）", [148.53, 75.47, 46.62], "Verified", S_AR + "（148,530,145；75,471,521；46,618,441 元）"),
        (14, "h_nonop", "業外（＝稅前 − 核心營業利益；含財務費用、投資收益、營業外收支）", ["=C13-C11", "=D13-D11", "=E13-E11"], "Derived", ""),
        (15, "h_nci", "少數股東損益", [0, 0, 1.0], "Interested-party", S_HA + "（2025：1；2023、2024 未取得，以 0 計，影響小於 1 百萬元）"),
        (16, "h_ni", "歸屬母公司淨利", [140.44, 77.91, 61.13], "Verified", S_AR),
        (17, "h_nit", "稅後淨利（含少數股東）", ["=C16+C15", "=D16+D15", "=E16+E15"], "Derived", "＝歸母淨利 ＋ 少數股東損益"),
        (18, "h_tax", "所得稅（＝稅前 − 稅後淨利；負＝稅收利益）", ["=C13-C17", "=D13-D17", "=E13-E17"], "Derived", "2025 年華安轉錄所得稅 −16 與之相符"),
        (19, "h_eps", "EPS（基本，元）", [0.3371, 0.1856, 0.1459], "Verified", S_AR),
        (20, "h_dps", "每股股利（元）", [0.04, 0.024, 0.025], "Interested-party", S_SA + "；2025 年度預案每股 0.025 元（" + S_AR + "）"),
        (21, "h_sh", "股數（＝淨利 ÷ EPS，百萬股）", ["=C16/C19", "=D16/D19", "=E16/E19"], "Derived", "公司總股本 418.88 百萬股；2023 年加權股數較少"),
        (22, "h_da", "折舊攤銷", [101.81, 101.4, 116.41], "Interested-party", S_SA + "（華安 2025 年 132，口徑不同）"),
        (23, "h_cx", "資本支出", [194.98, 297.72, 317.2], "Interested-party", S_SA),
        (24, "h_ocf", "營業現金流", [203.14, 246.15, 219.71], "Verified", S_AR),
    ]
    for (rr, k, lab, vals, tag, src) in HROWS:
        HI.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = HI.cell(rr, 3 + j, v)
                if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
                if k in ("h_gm", "h_ebitm"): c.number_format = "0.0%"
        HI.cell(rr, 6, tag); HI.cell(rr, 7, src)
        if tag == "Interested-party": HI.cell(rr, 8, IP_DATA if "stockanalysis" in src else IP_BRK)
        dn(wb, k, f"Hist!$C${rr}:$E${rr}")
    HI.cell(28, 1, "檢查：分部合計 − 營業收入（FY2024、FY2025；其他板塊為殘差，恆為 0）")
    for c in "DE": HI[f"{c}28"] = f"={c}8+{c}9-{c}4"
    HI.cell(29, 1, "檢查：核心營業利益＋業外 − 稅前（應為 0）")
    for c in "CDE": HI[f"{c}29"] = f"={c}11+{c}14-{c}13"
    HI.column_dimensions["A"].width = 58; HI.column_dimensions["G"].width = 90

    # ================= Q（2026 年 1–9 月；一列一個數字，單一標記）=================
    Q["A1"] = "Q：2026 年三季報（公司公告原文，未經審計）與 2025 年分季營收（百萬元）。一列一個項目、單一主要標記。"
    hdr(Q, 3, ["項目", "2025Q4", "2026Q3", "2026 年 1–9 月", "2025 年 1–9 月", "標記", "來源", "來源誘因", "說明"])
    QROWS = [   # (列, 項目, [B,C,D,E], 標記, 來源, 說明)
        (4, "營業收入", [718.86, 913.12, 2480.25, "=595.31+719.26+728.54"], "Verified", S_AR + "（2025 分季）；" + S_Q3, "1–9 月年增 21.40%、Q3 年增 25.33%"),
        (5, "營業成本", [None, None, 1534.60, 1307.81], "Verified", S_Q3, ""),
        (6, "毛利", [None, None, "=D4-D5", "=E4-E5"], "Derived", "＝營收 − 營業成本", ""),
        (7, "毛利率", [None, 0.3887, "=D6/D4", "=E6/E4"], "Derived", "＝毛利 ÷ 營收；Q3 單季 38.87% 取自" + S_HA, "1–9 月 38.13%（華安同值）"),
        (8, "稅金及附加", [None, None, 13.43, None], "Verified", S_Q3, ""),
        (9, "銷售費用", [None, None, 225.25, None], "Verified", S_Q3, ""),
        (10, "管理費用", [None, None, 344.41, None], "Verified", S_Q3, ""),
        (11, "研發費用", [None, None, 232.05, None], "Verified", S_Q3, ""),
        (12, "其他收益", [None, None, 14.23, None], "Verified", S_Q3, ""),
        (13, "信用減值損失", [None, None, -5.20, None], "Verified", S_Q3, ""),
        (14, "資產減值損失", [None, None, -19.04, None], "Verified", S_Q3, "上年同期為 +5.33"),
        (15, "核心營業利益（毛利 − 稅金 − 三費 ＋ 其他收益 ＋ 減值）", [None, None, "=D6-D8-D9-D10-D11+D12+D13+D14", None], "Derived", "＝上列加總；不含財務費用與投資收益", ""),
        (16, "財務費用", [None, None, 21.10, None], "Verified", S_Q3, "上年同期 2.31，主因匯兌與借款增加"),
        (17, "利潤總額", [None, None, 99.76, None], "Verified", S_Q3, "與核心營業利益 − 財務費用 ＋ 投資收益等合計之差為營業外收支取整"),
        (18, "歸屬母公司淨利", [11.29, 63.00, 100.20, 49.84], "Verified", S_AR + "（2025Q4）；" + S_Q3, "1–9 月 +101.05%"),
    ]
    for (rr, lab, vals, tag, src, note) in QROWS:
        Q.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = Q.cell(rr, 2 + j, v)
                if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
                if rr == 7: c.number_format = "0.00%"
        Q.cell(rr, 6, tag); Q.cell(rr, 7, src); Q.cell(rr, 8, ""); Q.cell(rr, 9, note)
    Q.column_dimensions["A"].width = 44; Q.column_dimensions["G"].width = 90

    # ================= 線取數（線模型值；藍字＝讀入值，M 欄寫來源檔與儲存格）=================
    LN["A1"] = "線取數：本模型讀入的線層級變數年度路徑（建置時由線 4 成品依 ID 讀值）。情境對應：A＝線模型高情境、B＝基準、C＝低情境。每台空心杯用量與單價年降讀線 6（見 In 頁與對帳頁）。"
    hdr(LN, 3, ["項目", "情境", ""] + [f"FY{y}" for y in range(2025, 2033)] + ["單位", "來源（檔案與儲存格）"])
    LR = {}
    r = 4
    for s in SC:
        LN.cell(r, 1, "全球人形機器人出貨（線 4）"); LN.cell(r, 2, f"{s}（{LINE_SC[s]}）")
        srcs = []
        for y in range(2026, 2033):
            v, src = LD["hum"][(y, s)]
            c = LN.cell(r, 4 + (y - 2025), v); c.font = BLUE; c.number_format = "#,##0"
            srcs.append(src)
        LN.cell(r, 12, "台"); LN.cell(r, 13, f"{srcs[0]} … {srcs[-1]}")
        LR[("hum", s)] = r; r += 1
    LN.column_dimensions["A"].width = 52; LN.column_dimensions["M"].width = 110

    # ================= Thesis（命題模組）=================
    TH["A1"] = ("命題模組：控制電機及驅動系統營收＝非機器人用（FY2026 錨，逐年成長）＋機器人應用（非人形，逐年成長）＋人形空心杯用（線 4 出貨 × 線 6 每台用量 × 份額 × 單價）。"
                "FY2026＝1–9 月實績＋第 4 季假設；單價＝線 6 2026 單價 × 匯率 ×（1 − 年降）^(年 − 2026)。百萬元。")
    TH["A2"] = "欄：D＝FY2025 … K＝FY2032。人形 2026 隱含份額以線 4 基準路徑校準（三情境同一實績年）。人形空心杯用營收是整體營收的小部分（見 PL），本頁據此顯示它對估值的重要性有限。"
    hdr(TH, 3, ["項目", "說明", ""] + [f"FY{y}" for y in range(2025, 2033)])
    THR = {}
    r = 4
    for s in SC:
        TH.cell(r, 1, f"情境 {s}").font = BOLD; TH.cell(r, 1).fill = SECF; r += 1
        R = {}
        R["rev26"] = r; TH.cell(r, 1, "FY2026 合併營收（錨）"); TH.cell(r, 2, "＝1–9 月＋2025Q4 ×（1＋q4g）")
        TH[f"{TC[2026]}{r}"] = "=Q!D4+Q!B4*(1+q4g)"; r += 1
        R["px"] = r; TH.cell(r, 1, "空心杯單價（元／顆）"); TH.cell(r, 2, "＝線 6 單價 × 匯率，逐年 ×（1 − 年降）")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = "=p_cm*fx" if y == 2026 else f"={TC[y - 1]}{r}*(1-pd)"
        r += 1
        R["ship"] = r; TH.cell(r, 1, "全球人形出貨（線 4 × 出貨倍數，千台）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR[('hum', s)]}*hu_k/1000"
        r += 1
        R["hdem"] = r; TH.cell(r, 1, "全球人形空心杯需求（千顆）"); TH.cell(r, 2, "＝出貨 × 每台用量（線 6）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['ship']}*n_cup"
        r += 1
        R["lmk"] = r; TH.cell(r, 1, "線模型口徑的全球人形空心杯市場（百萬元）"); TH.cell(r, 2, "＝需求 × 單價 ÷ 1000")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['hdem']}*{TC[y]}{R['px']}/1000"
        r += 1
        R["cm26"] = r; TH.cell(r, 1, "FY2026 控制電機及驅動系統（錨）"); TH[f"{TC[2026]}{r}"] = f"={TC[2026]}{R['rev26']}*sh_cm26"; r += 1
        R["hsh"] = r; TH.cell(r, 1, "人形空心杯份額"); TH.cell(r, 2, "FY2026＝隱含（以線 4 基準校準）；線性走到 hs30，其後持平")
        R["hum"] = r + 1
        TH[f"{TC[2026]}{r}"] = f"={TC[2026]}{R['hum']}/(線取數!{TC[2026]}{LR[('hum', 'B')]}*hu_k/1000*n_cup*{TC[2026]}{R['px']}/1000)"
        for y in range(2027, 2033):
            TH[f"{TC[y]}{r}"] = f"=${TC[2026]}${r}+(hs30_{s}-${TC[2026]}${r})*MIN(1,({y}-2026)/4)"
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 1
        TH.cell(r, 1, "控制電機：人形空心杯用 營收"); TH.cell(r, 2, "FY2026＝營收 × 機器人占比 × 人形占比；FY2027 起＝線模型市場 × 份額")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = f"={c}{R['rev26']}*r_sh26*hum_sh26" if y == 2026 else f"={c}{R['lmk']}*{c}{R['hsh']}"
        r += 1
        R["robo"] = r; TH.cell(r, 1, "控制電機：機器人應用（非人形）營收")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = f"={c}{R['rev26']}*r_sh26*(1-hum_sh26)" if y == 2026 else f"={TC[y - 1]}{r}*(1+grob_{s})"
        r += 1
        R["ind"] = r; TH.cell(r, 1, "控制電機：非機器人用 營收"); TH.cell(r, 2, "FY2026＝控制電機 − 機器人應用；之後逐年成長")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = f"={c}{R['cm26']}-{c}{R['robo']}-{c}{R['hum']}" if y == 2026 else f"={TC[y - 1]}{r}*(1+gind_{s})"
        r += 1
        R["cm"] = r; TH.cell(r, 1, "控制電機及其驅動系統 合計")
        for y in range(2026, 2033): c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['ind']}+{c}{R['robo']}+{c}{R['hum']}"
        r += 1
        R["oth"] = r; TH.cell(r, 1, "其他板塊（電源與照明、設備狀態管理、貿易）")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['rev26']}*(1-sh_cm26)" if y == 2026 else f"={TC[y - 1]}{r}*(1+g_o)"
        r += 1
        R["humsh"] = r; TH.cell(r, 1, "人形空心杯用占合併營收")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['hum']}/({c}{R['cm']}+{c}{R['oth']})"; TH[f"{c}{r}"].number_format = "0.0%"
        r += 1
        R["gm"] = r; TH.cell(r, 1, "毛利率路徑"); TH.cell(r, 2, "FY2026＝1–9 月；線性到 FY2030（gm30 ＋ gm_adj），其後持平")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = "=gm26" if y == 2026 else f"=gm26+(gm30_{s}+gm_adj-gm26)*MIN(1,({y}-2026)/4)"; TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 1
        R["opex"] = r; TH.cell(r, 1, "營業費用率路徑"); TH.cell(r, 2, "FY2026＝opex26；線性到 FY2030 opex30，其後持平")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = "=opex26" if y == 2026 else f"=opex26+(opex30-opex26)*MIN(1,({y}-2026)/4)"; TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 2
        THR[s] = R
    for row in TH.iter_rows(min_row=4, max_row=r, min_col=4, max_col=11):
        for c in row:
            if c.number_format == "General": c.number_format = "#,##0.0"
    TH.column_dimensions["A"].width = 40; TH.column_dimensions["B"].width = 50

    # ================= PL（標準損益表，三情境）=================
    PL["A1"] = ("PL：官方分部（控制電機及其驅動系統／其他板塊）→ 合併營收 → 標準損益表 → FCFF。"
                "C＝FY2023 … E＝FY2025 為實績；F＝FY2026（1–9 月實績＋第 4 季假設）… L＝FY2032。百萬元；中國會計準則（法定財報口徑，GAAP）。")
    PR = {}
    r = 3
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
        H = lambda rr: (lambda y: f"=Hist!{'CDE'[y - 2023]}{rr}")
        H24 = lambda rr: (lambda y: None if y == 2023 else f"=Hist!{'CDE'[y - 2023]}{rr}")
        C = lambda y: COL[y]
        row("ind", "控制電機：非機器人用（子列）", "百萬元", None, lambda y: f"=Thesis!{TC[y]}{th['ind']}")
        row("robo", "控制電機：機器人應用（非人形，子列）", "百萬元", None, lambda y: f"=Thesis!{TC[y]}{th['robo']}")
        row("hum", "控制電機：人形空心杯用（子列）", "百萬元", None, lambda y: f"=Thesis!{TC[y]}{th['hum']}")
        row("cm", "控制電機及其驅動系統 營收", "百萬元", H24(8), lambda y: f"=SUM({C(y)}{R['ind']}:{C(y)}{R['hum']})")
        row("oth", "其他板塊 營收（電源與照明、設備狀態管理、貿易）", "百萬元", H24(9), lambda y: f"=Thesis!{TC[y]}{th['oth']}")
        row("rev", "營收（合併）", "百萬元", H(4), lambda y: f"=SUM({C(y)}{R['cm']}:{C(y)}{R['oth']})")
        row("chk", "檢查：分部合計 − 合併營收（應為 0）", "百萬元",
            lambda y: None if y == 2023 else f"={C(y)}{R['cm']}+{C(y)}{R['oth']}-{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['cm']}+{C(y)}{R['oth']}-{C(y)}{R['rev']}")
        PL.cell(R["chk"], 1, "檢查：分部合計 − 合併營收（應為 0）" if s == "B" else f"檢查（情境 {s}）：分部合計 − 合併營收（應為 0）")
        row("chk2", "檢查：控制電機子列合計 − 控制電機分部（應為 0）", "百萬元", None, lambda y: f"={C(y)}{R['ind']}+{C(y)}{R['robo']}+{C(y)}{R['hum']}-Thesis!{TC[y]}{th['cm']}")
        row("yoy", "營收年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1",
            lambda y: f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1", "0.0%")
        row("cogs", "營業成本", "百萬元", H(5), lambda y: f"={C(y)}{R['rev']}*(1-Thesis!{TC[y]}{th['gm']})")
        row("gp", "毛利", "百萬元", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}")
        row("gm", "毛利率", "%", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", "0.0%")
        row("opex", "營業費用（銷售＋管理＋研發，含稅金及附加與其他營業損益）", "百萬元", H(10), lambda y: f"={C(y)}{R['rev']}*Thesis!{TC[y]}{th['opex']}")
        row("op", "營業利益（核心，不含財務費用與投資收益）", "百萬元", H(11), lambda y: f"={C(y)}{R['gp']}-{C(y)}{R['opex']}")
        row("opm", "營業利益率", "%", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", "0.0%")
        row("nonop", "其他收入淨額（業外：財務費用、投資收益、營業外）", "百萬元", H(14), lambda y: "=-fin_r*debt_v")
        row("pbt", "稅前利益", "百萬元", H(13), lambda y: f"={C(y)}{R['op']}+{C(y)}{R['nonop']}")
        row("tax", "所得稅", "百萬元", H(18), lambda y: f"={C(y)}{R['pbt']}*tax_r")
        row("nit", "稅後淨利（含少數股東）", "百萬元", H(17), lambda y: f"={C(y)}{R['pbt']}-{C(y)}{R['tax']}")
        row("nci", "少數股東損益", "百萬元", H(15), lambda y: f"={C(y)}{R['nit']}*nci_r")
        row("ni", "歸屬母公司淨利", "百萬元", H(16), lambda y: f"={C(y)}{R['nit']}-{C(y)}{R['nci']}")
        row("nim", "淨利率", "%", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", "0.0%")
        row("niy", "淨利年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1",
            lambda y: f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1", "0.0%")
        row("sh", "股數（百萬股）", "百萬股", H(21), lambda y: "=sh_v")
        row("eps", "EPS（元）", "元", H(19), lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['sh']}", "0.00")
        row("da", "折舊攤銷", "百萬元", H(22), lambda y: f"={C(y)}{R['rev']}*(da26+(cx_r*da_ss-da26)*{y - 2026}/6)")
        row("cx", "資本支出", "百萬元", H(23), lambda y: "=cx26" if y == 2026 else f"={C(y)}{R['rev']}*cx_r")
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
        R["nc"] = r; DC.cell(r, 1, "淨現金滾到評價日（現金 − 有息負債 − 少數股東權益 ＋ Stub FCFF）")
        DC[f"B{r}"] = f"=cash_v-debt_v-nci_v+B{R['stub']}"; r += 1
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
    CP["A1"] = ("Comp：同業前瞻 P/E（stockanalysis.com statistics，資料商）。股價＝" + S_CLOSE + "（Verified）；前瞻 P/E＝原值 × 新收盤 ÷ 原報價（Derived，EPS 不變）；"
                "哈默納科、Nabtesco 取本專案模型同日換算值。核心同業 6 家進中位數；少於 8 家不做迴歸。報價日未取得的同業另列，不進中位數。")
    hdr(CP, 4, ["公司", "代號", "股價（2026-11-20）", "報價日", "前瞻 P/E 分母", "前瞻 P/E", "近 12 月 P/E（原報價日）", "5 年 β", "會計口徑", "進中位數", "說明", "原報價", "原報價日", "原前瞻 P/E", "標記與換算"])
    PEERS = [  # (公司, 代號, 11-20 收盤, 原報價, 原報價日, 原前瞻 P/E, 近 12 月 P/E, β, 口徑, 進中位數, 說明, 分母)
        ("雙環傳動", "002472.SZ", 35.45, 41.21, "2026-10-30", 22.32, 27.11, 0.52, "中國會計準則", "是", "齒輪、精密減速器；人形關節", "資料商 NTM 共識 EPS"),
        ("兆威機電", "003021.SZ", 76.83, 70.25, "2026-09-30", 59.70, 87.15, 0.44, "中國會計準則", "是", "微型傳動、靈巧手（可能的下游或競爭者）", "資料商 NTM 共識 EPS"),
        ("拓普集團", "601689.SS", 48.11, 46.52, "2026-09-30", 26.89, 32.10, 0.80, "中國會計準則", "是", "汽車零件、人形致動器", "資料商 NTM 共識 EPS"),
        ("綠的諧波", "688017.SS", 230.80, 263.76, "2026-10-09", 225.96, 342.80, 0.79, "中國會計準則", "是", "諧波減速器；人形零組件題材股", "資料商 NTM 共識 EPS（三季報前）"),
        ("哈默納科", "6324.T", 3722, 3722, "2026-11-20", 48.94, None, 0.96, "日本基準（未查核）", "是", "諧波減速機；取 physical_ai/harmonic_drive 模型 Comp!F14", "公司 FY2026 指引 EPS（2027年3月期）"),
        ("Nabtesco", "6268.T", 3105, 3105, "2026-11-20", 13.77, None, 0.93, "IFRS（未查核）", "是", "RV 減速機；取 physical_ai/nabtesco 模型 In 頁 pe_self", "資料商 NTM 共識 EPS（11-13 值換算）"),
        ("鳴志電器（本公司，對照）", "603728.SS", 44.52, 51.50, "2026-09-30", 124.10, 302.24, 0.62, "中國會計準則", "否", "本公司；只作對照", "資料商 NTM 共識 EPS（三季報前）"),
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
    CP.cell(12, 1, "報價日未取得（不進中位數）：Maxon（未上市）、Regal Rexnord（Kollmorgen）、Nidec 6594——收盤檔無 2026-11-20 報價或未讀取資料商前瞻 P/E")
    CP.cell(15, 1, "核心同業前瞻 P/E 中位數"); CP["F15"] = "=MEDIAN(F5:F10)"
    CP.cell(16, 1, "本公司現價 ÷ 資料商前瞻 EPS（換算至 2026-11-20；分母：資料商 NTM 共識，三季報前）"); CP["F16"] = "=F11"
    CP.cell(17, 1, "本公司現價 ÷ 券商 FY2027 EPS 平均（分母：華安 0.56、群益 2.2 億÷股本）"); CP["F17"] = "=px/AVERAGE(0.56,220/sh_v)"
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
    row1("gate_cover", "收支缺口 ÷ 淨現金（情境 B；負值才有意義）", f"=B{VR['gate_b']}/(cash_v-debt_v)", "0.0%", "淨現金可否覆蓋 FY2026–FY2030 缺口")
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
        SE.cell(r + 1, 1, "說明：折舊、資本支出（FY2027 起）、營運資金隨營收等比例變動（第 44 條）；單價年降幅同時作用於工業與人形用。")
    SE.column_dimensions["A"].width = 52; SE.column_dimensions["B"].width = 44


    # ================= Recon =================
    RC["A1"] = "Recon：模型對實績、券商預測、線 6 市場規模的同口徑對照；反解支持現價所需的主導變數。"
    hdr(RC, 3, ["項目", "模型", "對照", "差距（模型 − 對照）", "說明"])
    rr = 4
    def rrow(lab, a, b, note, fmt="#,##0.0", diff=True):
        nonlocal rr
        RC.cell(rr, 1, lab); RC.cell(rr, 2, a).number_format = fmt
        if b is not None: RC.cell(rr, 3, b).number_format = fmt
        if diff and b is not None: RC.cell(rr, 4, f"=B{rr}-C{rr}").number_format = fmt
        RC.cell(rr, 5, short(note)); rr += 1
    rrow("FY2026 營收（三情境同）vs 華安 2026E（3,375）", f"=PL!F{PR['B']['rev']}", 3375, S_HA + "；三季報後；券商預測為 Interested-party")
    rrow("FY2027 營收（情境 B）vs 華安 2027E（3,991）", f"=PL!G{PR['B']['rev']}", 3991, "")
    rrow("FY2028 營收（情境 B）vs 華安 2028E（4,550）", f"=PL!H{PR['B']['rev']}", 4550, "")
    rrow("FY2026 歸屬母公司淨利 vs 華安 2026E 162、群益 140 的平均", f"=PL!F{PR['B']['ni']}", "=AVERAGE(162,140)", S_HA + "；" + S_QY)
    rrow("FY2027 歸屬母公司淨利（情境 B）vs 華安 233、群益 220 的平均", f"=PL!G{PR['B']['ni']}", "=AVERAGE(233,220)", "")
    rrow("FY2028 歸屬母公司淨利（情境 B）vs 華安 306、群益 280 的平均", f"=PL!H{PR['B']['ni']}", "=AVERAGE(306,280)", "")
    rrow("FY2027 EPS（情境 B）vs 華安 2027E（0.56 元）", f"=PL!G{PR['B']['eps']}", 0.56, "", "0.00")
    rrow("FY2026 毛利率 vs 2026 年 1–9 月實績", f"=PL!F{PR['B']['gm']}", "=gm26", "模型以 1–9 月為錨（差距應為 0）", "0.0%")
    rrow("FY2026 第 4 季隱含營收 vs 2026Q3 實績", f"=PL!F{PR['B']['rev']}-Q!D4", "=Q!C4", "第 4 季若維持第 3 季水準即為對照值")
    rr += 1
    RC.cell(rr, 1, "線 6 市場對照（人形空心杯；同口徑：線 6 計算頁 2030 基準市場，美元）").font = BOLD; rr += 1
    rrow("線模型口徑 2030 人形空心杯市場（情境 B，百萬美元）vs 線 6 計算頁 2030 基準", f"=Thesis!J{THR['B']['lmk']}/fx", LD["mkt30"][0], f"線 6 市場規模 2030 基準 {LD['mkt30'][0]:,.1f} 百萬美元（{LD['mkt30'][1]}）；差距應約為 0", "#,##0.0")
    rrow("FY2030 人形空心杯用營收（情境 B，百萬元）", f"=PL!J{PR['B']['hum']}", None, "＝線模型市場 × 份額", "#,##0.0", False)
    rrow("FY2030 人形空心杯用營收占合併營收（情境 B）", f"=Thesis!J{THR['B']['humsh']}", None, "人形用占比小：量級大幅偏離線 4 基準也難改變評等（見敏感度、反解）", "0.0%", False)
    rrow("FY2026 人形隱含全球份額（以線 4 基準校準）", f"=Thesis!E{THR['B']['hsh']}", None, "＝人形空心杯用營收 ÷（線 4 出貨 × 線 6 每台用量 × 單價）", "0.0%", False)
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
    h26 = LD["hum"][(2026, "B")]; h32 = LD["hum"][(2032, "B")]
    ZR = [
        ("全球人形機器人出貨 FY2026–FY2032（A／B／C＝線高／基準／低）", "4 人形", f"基準 {h26[0]:,.0f} → {h32[0]:,.0f}（× hu_k）", "台", "全球、含小型機", f"{h26[1]}；{h32[1]}", "Assumed", "Thesis 人形空心杯需求與份額", "一致（hu_k＝1）；高案 hu_k＝2 為本模型自設，列 Sens 不改基準"),
        ("人形每台空心杯用量", "6 共用零組件", "=n_cup", "顆／台", "出貨組合平均；取決於靈巧手自由度", f"{LD['n_cup'][1][1]}（低 {LD['n_cup'][0][1]}、高 {LD['n_cup'][2][1]}）", "Interested-party", "Thesis 全球人形空心杯需求", "一致（直接讀值）；線 4 未另給空心杯用量，未見不一致"),
        ("空心杯 2026 單價（線 6 輸入列）", "6 共用零組件", "=p_cm", "美元／顆", "全球混合；國產低階與 Maxon、Faulhaber 價差大", f"{LD['p26'][1][1]}（低 {LD['p26'][0][1]}、高 {LD['p26'][2][1]}）", "Interested-party", "Thesis 單價與隱含份額", "一致（直接讀值）；本公司實際單價未揭露，僅影響隱含份額換算"),
        ("空心杯單價年降幅", "6 共用零組件", "=pd", "%／年", "線 6 輸入", f"{LD['pd'][1][1]}", "Analogy", "Thesis 單價路徑", "一致（直接讀值）"),
        ("空心杯市場規模 2030（基準）", "6 共用零組件", f"{LD['mkt30'][0]:,.1f}", "百萬美元", "全部來自人形", f"{LD['mkt30'][1]}", "Derived", "Recon 對照（不進公式）", "一致：本模型 Thesis 同口徑重算值見 Recon 第 1 列"),
        ("本公司空心杯份額", "6 共用零組件", "未取得", "—", "線 6 份額頁列 Maxon、Kollmorgen，本公司份額未取得", f"{L6N} 份額頁", "Assumed", "hum_sh26、hs30_*（本模型自設）", "線模型無值；本模型自設並標 Assumed，區間見 In 頁"),
        ("線 3 工業機器人諧波需求", "3 工業機器人", "未使用", "—", "本公司不產諧波減速器", "—", "—", "—", "不適用"),
    ]
    for i, z in enumerate(ZR):
        for j, v in enumerate(z): DZ.cell(4 + i, 1 + j, v)
    DZ.column_dimensions["A"].width = 50; DZ.column_dimensions["F"].width = 80; DZ.column_dimensions["I"].width = 70

    # ================= Risk =================
    RK["A1"] = "Risk：風險登錄（原則第 10 節第 10、19 條）。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    RISKS = [
        ("毛利率與費用率回升不如預期", "是（gm30、opex30、gm_adj）", "2025 年淨利 −22%；三季報毛利率 38.1%、費用率下降是目前估值的核心", "季毛利率、管理與研發費用率"),
        ("海外需求與匯率（境外收入占 45%）", "部分（gind、fin_r）", "2025 年財務費用與 2026 年 1–9 月財務費用 21.1 百萬元（上年 2.3）含匯兌", "匯率、越南廠產能利用率"),
        ("人形空心杯量產延後或份額低於預期", "是（情境 C、hs30、hu_k）", "人形用營收占比小，對估值影響有限", "靈巧手整機廠量產公告；因時、兆威供應鏈"),
        ("整機廠自製空心杯或價格競爭", "是（pd、hs30_C）", "國產低階價格低", "靈巧手報價"),
        ("資本支出與營運資金吃掉現金（收支關卡）", "是（cx_r、nwc_r、收支關卡）", "2025 年資本支出 317 高於營業現金流 220", "營業現金流、存貨與應收"),
        ("資訊揭露品質（2026-11-06 更正三季報前十大股東表）", "否", "更正為股東欄位列示，財務數字未見更正", "後續公告"),
        ("單一客戶集中度", "否", "年報摘要不含前五大客戶，全文未取得", "年報全文"),
        ("租稅（有效稅率歷史極低且波動）", "部分（tax_r 3–15%）", "稅率回到 15% 會直接壓低 EPS", "高新技術企業資格、遞延所得稅"),
    ]
    for i, k in enumerate(RISKS):
        for j, v in enumerate(k): RK.cell(4 + i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["D"].width = 60

    # ================= Sources =================
    SO["A1"] = "Sources：本模型引用的原始來源與資料商（含網址；HTML 展示層不放網址）。"
    for i, s_ in enumerate([S_PX, S_Q3, S_AR, S_HA, S_QY, S_SA, S_NEWS, S_SA_ST, S_RF, S_CRP, S_SAP,
                            f"線 4：{L4N}", f"線 6：{L6N}",
                            "physical_ai/leaderdrive 模型 Comp 頁（同業原報價與 β）、physical_ai/harmonic_drive 模型（哈默納科前瞻 P/E）、physical_ai/nabtesco 模型（Nabtesco 前瞻 P/E）",
                            "physical_ai/map/profiles/603728.md（背景，數字已另行查證）"], 3):
        SO.cell(i, 1, s_)
    SO.column_dimensions["A"].width = 160

    # ================= README（xlsx 內）=================
    S = T
    v = lambda k: f"Val!B{VR[k]}"
    lines = [
        ('="鳴志電器（603728.SS）個股模型 ' + VER + '｜build_id "&Sens!D2&"｜資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'),
        "命題：鳴志電器是控制電機（步進、無刷、空心杯）廠，股價有人形機器人題材，但人形空心杯用營收在線模型口徑下只占營收的小部分；估值主要取決於控制電機本業的毛利率與費用率回升，以及自由現金流（資本支出與營運資金）。",
        (f'="結論（條件式）：在 FY2026 以 1–9 月實績為錨（毛利率 "&{S("gm26","0.0%")}&"）、情境 B 的 2030 年毛利率 "&{S("gm30_B","0.0%")}&"、營業費用率降到 "&{S("opex30","0.0%")}&"、人形出貨 × "&{S("hu_k","0.0")}&" 倍線 4、同業 P/E 未加溢價（相對倍數係數 "&{S("rel_k","0.00")}&"）的假設下，加權目標價約 "&{S(v("tp"),"#,##0.0")}&" 元，相對現價 "&{S("px","#,##0.00")}&" 元 "&{S(v("up"),"+0.0%;-0.0%")}&"，估值評等 "&{v("rating")}&"，投資建議 "&{v("rec")}&"。"'),
        (f'="DCF 目標價 "&{S(v("dcf_tp"),"#,##0.0")}&"、Comparable 目標價 "&{S(v("cmp_tp"),"#,##0.0")}&"（各半）；WACC "&{S("wacc","0.00%")}&"、ke "&{S("ke","0.00%")}&"、g "&{S("g_tv","0.00%")}&"；情境 B 終值占 EV "&{S(f"Val!C{VR['tvs']}","0%")}&"。"'),
        (f'="上檔空間歸因：DCF 貢獻 "&{S(v("att_d"),"+0.0%;-0.0%")}&"、未調整 Comparable 貢獻 "&{S(v("att_c"),"+0.0%;-0.0%")}&"、相對倍數調整貢獻 "&{S(v("att_r"),"+0.0%;-0.0%")}&"。本公司現價對資料商前瞻 EPS 的倍數是同業中位數的 "&{S(v("relnow"),"0.0")}&" 倍；支持現價需要相對倍數係數 "&{S(v("relreq"),"0.0")}&"，升一級需要 "&{S(v("relflip"),"0.0")}&"。"'),
        (f'="收支關卡：情境 B FY2026–FY2030 FCFF 扣股利合計 "&{S(v("gate_b"),"#,##0.0")}&" 百萬元，"&{v("gate_t")}&"；距最近級距邊界 "&{S(v("dist"),"0.0%")}&"。"'),
        "主要失效條件（可觀察）：(1) 毛利率持續高於 40% 且費用率降到 28% 以下；(2) 人形與液冷等新業務營收占比快速上升且公司揭露；(3) 自由現金流轉正（資本支出回落）；(4) 有效稅率維持極低。",
        "方法：DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，WACC 折現，明確期到 FY2032，RONIC＝WACC）50%＋同業前瞻 P/E × 模型 FY2027 EPS 50%；三情境 25／50／25。",
        "限制：人形空心杯用營收占比、機器人應用口徑為 Assumed／Interested-party（公司未揭露絕對值）；分部歷史只有控制電機一項（年報全文未取得）；同業 P/E 由原報價日按 2026-11-20 收盤等比例換算（EPS 未更新）；不讀 Siliconomics／Tokenomics／Wattonomics。詳見 repo README.md。",
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
    extra = [
        ("情境 B 的 2030 年人形空心杯份額改為情境 A 值", "hs30_B", ivE["hs30_A"]),
        ("情境 B 的 2030 年人形空心杯份額改為情境 C 值", "hs30_B", ivE["hs30_C"]),
        ("情境 B 非機器人用年增改為情境 A 值", "gind_B", ivE["gind_A"]),
        ("情境 B 非機器人用年增改為情境 C 值", "gind_B", ivE["gind_C"]),
        ("情境 B 機器人應用（非人形）年增改為情境 A 值", "grob_B", ivE["grob_A"]),
        ("情境 B 機器人應用（非人形）年增改為情境 C 值", "grob_B", ivE["grob_C"]),
        ("情境 B 的 FY2030 毛利率改為情境 A 值", "gm30_B", ivE["gm30_A"]),
        ("情境 B 的 FY2030 毛利率改為情境 C 值", "gm30_B", ivE["gm30_C"]),
        ("人形出貨為線 4 的 5 倍（全部情境）", "hu_k", 5.0),
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
        print(f"  {s['drv'][:44]:44s} {r['tp']:9.1f} {r['tp'] / base['tp'] - 1:+.1%}")

    # 反解（整本重算二分搜尋）
    px = ivE["px"]
    def solve(key, target_tp, lo, hi, n=11, log=True):
        f = (lambda x: math.log(x)) if log else (lambda x: x)
        g = (lambda x: math.exp(x)) if log else (lambda x: x)
        a, b = f(lo), f(hi)
        if run_cases([{key: hi}], meta, "solve")[0]["tp"] < target_tp: return None
        if run_cases([{key: lo}], meta, "solve")[0]["tp"] >= target_tp: return lo
        for _ in range(n):
            m = (a + b) / 2
            if run_cases([{key: g(m)}], meta, "solve")[0]["tp"] < target_tp: a = m
            else: b = m
        return g((a + b) / 2)
    up = base["up"]
    flip = -0.5 if up < -0.5 else -0.25 if up < -0.25 else 0.25 if up <= 0.25 else 0.5
    flip_tp = px * (1 + flip + 1e-6)
    s1 = solve("hu_k", px, 1.0, 300.0)
    s2 = solve("gm_adj", px, 0.0, 0.5, log=False)
    s3 = solve("gm_adj", flip_tp, 0.0, 0.5, log=False)
    nf = lambda x: x if x else -1
    solve_rows = [
        ("支持現價（目標價＝現價）所需的人形出貨倍數（相對線 4，全部情境）", nf(s1), 1.0, "其他輸入不變；−1＝300 倍內無解" if not s1 else f"其他輸入不變；情境 B 的 CY2030 全球人形出貨相當於 {LD['hum'][(2030, 'B')][0] * s1 / 1e6:,.1f} 百萬台"),
        ("支持現價所需的 FY2030 毛利率整體上調（三情境同加，百分點）", nf(s2), 0.0, "其他輸入不變；−1＝上調 50 個百分點仍無解"),
        (f"評等升一級（上檔 {flip:+.0%}）所需的 FY2030 毛利率整體上調（百分點）", nf(s3), 0.0, "其他輸入不變；−1＝上調 50 個百分點仍無解"),
    ]
    print("反解：", solve_rows)

    h = hashlib.sha1(json.dumps([list(map(str, x)) for x in IN], ensure_ascii=False).encode() + open(__file__, "rb").read()).hexdigest()[:8]
    build_id = f"{VER}+{h}"
    sens = dict(build_id=build_id, fp=fp, base_tp=base["tp"], base_dcf=base["dcf"], cases=cases, solve=solve_rows)
    wb, meta2 = make(sens=sens)
    raw = os.path.join(TMP, "final", NAME); os.makedirs(os.path.dirname(raw), exist_ok=True); wb.save(raw)
    recalc([raw], os.path.join(TMP, "final_rc"))
    for f in glob.glob(os.path.join(HERE, "dist", "*.xlsx")):
        if f != OUT: print("注意：dist 另有檔案", f)
    shutil.copy(os.path.join(TMP, "final_rc", NAME), OUT)
    fin = read_val(OUT, meta2)
    print("成品：", OUT, fin, "build_id", build_id)
    json.dump(dict(build_id=build_id, base=fin, cases=cases, solve=solve_rows), open(os.path.join(HERE, "sens.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    wbv = openpyxl.load_workbook(OUT, data_only=True)
    Hs = wbv["Hist"]
    src = S_NEWS + "；" + S_AR
    segs = [("控制電機及其驅動系統", "控制電機及其驅動系統 營收", Hs["E8"].value), ("其他板塊（電源與照明控制、設備狀態管理、貿易；殘差）", "其他板塊 營收（電源與照明、設備狀態管理、貿易）", Hs["E9"].value)]
    json.dump(dict(company="moons", ticker="603728.SS", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=round(v, 2), unit="百萬元", source=src) for n, l, v in segs],
                   sublines=[dict(parent="控制電機及其驅動系統", labels=["控制電機：非機器人用（子列）", "控制電機：機器人應用（非人形，子列）", "控制電機：人形空心杯用（子列）"])],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）", "檢查：控制電機子列合計 − 控制電機分部（應為 0）"],
                   note="官方分部＝年報四大板塊（控制電機及其驅動系統、電源與照明控制系統、設備狀態管理系統、貿易類）；年報摘要只提供合計與控制電機一項，後三者合併為「其他板塊」殘差。控制電機的用途子列（非機器人、機器人非人形、人形空心杯）為命題拆分（Assumed／Interested-party），合計＝控制電機分部。"),
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
    up_b = g("up_b")
    t = f"""# physical_ai/moons/ — 鳴志電器（603728.SS）個股模型 {VER}

<!-- 本檔由 build_moons.py 依成品重算值產生；數字不手寫（原則第 10 節第 51 條）。 -->

**本模型適用《高不確定性預估工作原則》**（`docs/plan/高不確定性預估工作原則.md`）：人形機器人需求尚未發生、公司未揭露機器人與人形別營收，相關輸入一律給區間並標 Assumed／Interested-party；結論只在證據邊界內成立。

- 成品：`dist/{NAME}`（build_id `{sens['build_id']}`；LibreOffice 重算，0 錯誤值）；HTML 展示層 `dist/{NAME[:-5]}.html`（`gen_html_json.py` 產生 `html.json`，再以 `tools/company_html/build_company_html.py` 產生）
- 建置：`python3 physical_ai/moons/build_moons.py`（需要 soffice；整本重算敏感度與反解；結尾自動跑 in_layout、sens_coverage --require-keycol、model_governance、rating_gate --strict、zero_rows、cached_values、segment_check）
- 日期：資料截止日 2026-11-24（財報、券商預測、同業倍數）｜成品日 2026-11-24｜股價基準日 2026-11-20（收盤 {I['px']:,.2f} 元）｜評價日 2026-12-31｜目標時點 2027-11-20
- 會計年度＝日曆年；單位百萬元人民幣；中國會計準則（法定財報口徑，GAAP）。上交所主板。

## 結論（條件式）

在「FY2026 以 1–9 月實績為錨（毛利率 {I['gm26']:.1%}）、情境 B 的 2030 年毛利率 {I['gm30_B']:.1%}、營業費用率由 {I['opex26']:.1%} 降到 {I['opex30']:.1%}、人形出貨為線 4 的 {I['hu_k']:.1f} 倍、同業 P/E 不加溢價（相對倍數係數 {I['rel_k']:.2f}）」的假設下：

| 項目 | 值 |
|---|---|
| 加權目標價（DCF {I['w_dcf']:.0%}＋Comparable，情境 25／50／25） | **{g('tp'):,.1f} 元** |
| DCF 目標價／Comparable 目標價 | {g('dcf_tp'):,.1f}／{g('cmp_tp'):,.1f} |
| 相對現價 {I['px']:,.2f} | **{g('up'):+.1%}** |
| 估值評等／投資建議 | **{g('rating')}／{g('rec')}** |
| 收支關卡（情境 B FY2026–FY2030 FCFF 扣股利合計） | {g('gate_b'):,.1f} 百萬元（{g('gate_t')}） |
| 情境價值（DCF A／B／C） | {g('dcf','B'):,.1f}／{g('dcf','C'):,.1f}／{g('dcf','D'):,.1f} |
| 情境價值（Comparable A／B／C） | {g('cmp','B'):,.1f}／{g('cmp','C'):,.1f}／{g('cmp','D'):,.1f} |
| 情境目標價（A／B／C） | {g('wt','B'):,.1f}／{g('wt','C'):,.1f}／{g('wt','D'):,.1f} |
| WACC／ke／g／β | {g('wacc'):.2%}／{g('ke'):.2%}／{I['g_tv']:.2%}／{I['beta']:.2f} |
| 終值占 EV（情境 B） | {g('tvs','C'):.0%} |

上檔空間歸因：DCF 貢獻 {g('att_d'):+.1%}、未調整 Comparable 貢獻 {g('att_c'):+.1%}、相對倍數調整 {g('att_r'):+.1%}。本公司現價對資料商前瞻 EPS 的倍數是核心同業中位數（{I['pe_peer']:.1f} 倍）的 {g('relnow'):.1f} 倍；要讓目標價＝現價，Comparable 的相對倍數係數須達 {g('relreq'):.1f}。DCF 以 RONIC＝WACC（新投資無超額報酬），而高相對倍數隱含持續超額報酬，兩法對競爭優勢的隱含假設相反。賣方（華安買入、群益買進）目標價未取得，不作對照。

**讀法：** 股價帶有人形機器人題材，但線模型口徑下人形空心杯用營收在情境 B 只由 {PL[f'F{PR["hum"]}'].value:,.0f} 增至 {PL[f'J{PR["hum"]}'].value:,.0f} 百萬元（FY2026→FY2030），占合併營收由 {TH[f'E{THB["humsh"]}'].value:.1%} 到 {TH[f'J{THB["humsh"]}'].value:.1%}。所以估值主要取決於本業：毛利率回升、費用率下降與自由現金流。整本重算反解：其他輸入不變時，{('支持現價需要人形出貨達線 4 的 ' + fn(sol[0][1], '{:,.0f}') + ' 倍') if sol[0][1] != -1 else '人形出貨放大到線 4 的 300 倍仍不足以支持現價'}；或三情境 FY2030 毛利率整體上調 {fn(sol[1][1] * 100, '{:,.1f}') if sol[1][1] != -1 else '逾 50'} 個百分點。結論主要建立在 Assumed（用途拆分、份額、毛利率與費用率路徑、稅率）與資料商／券商轉錄的 Interested-party 輸入上。

## 命題

1. **營收結構：** 官方四大板塊中，控制電機及其驅動系統 2025 年 {wb['Hist']['E8'].value:,.0f} 百萬元，占 {wb['Hist']['E8'].value / wb['Hist']['E4'].value:.0%}；其餘（電源與照明控制、設備狀態管理、貿易）年報摘要未拆分，合為殘差。機器人應用 2026 上半年約占營收 {I['r_sh26']:.0%}（華安轉錄，公司未單獨披露絕對值），人形空心杯用占其中 {I['hum_sh26']:.0%}（Assumed，區間 3–25%）。
2. **量：** 人形空心杯用＝線 4 全球出貨 × 線 6 每台 {I['n_cup']:.0f} 顆 × 份額；2026 年隱含全球份額 {TH[f'E{THB["hsh"]}'].value:.0%}（以線 4 基準校準）。非機器人與機器人（非人形）依成長率假設延伸。
3. **價：** 空心杯單價讀線 6（2026 年 {I['p_cm']:.0f} 美元，年降 {I['pd']:.0%}）。本公司整體毛利率 2025 年 36.4%，2026 年 1–9 月 38.1%。

情境 B：營收 {ser('rev', '{:,.0f}')}；毛利率 {ser('gm', '{:.1%}')}；EPS {ser('eps', '{:.2f}')} 元（FY2026→FY2032）。

## 主要假設與可證偽指標（原則第 10 節第 20 條）

| 承重假設 | 基準（A／B／C 或區間） | 標記 | 可觀察指標 |
|---|---|---|---|
| FY2030 毛利率 | {I['gm30_A']:.1%}／{I['gm30_B']:.1%}／{I['gm30_C']:.1%} | Assumed | 季毛利率（2026 第 3 季 38.9%） |
| FY2030 營業費用率 | 28.5%／{I['opex30']:.1%}／33.0%（區間） | Assumed | 管理、研發費用率（第 3 季期間費用率 30.6%） |
| 非機器人用年增 | {I['gind_A']:.0%}／{I['gind_B']:.0%}／{I['gind_C']:.0%} | Assumed | 控制電機分產品收入（上半年 +25.7%） |
| 機器人應用占營收（2026） | {I['r_sh26']:.0%}（7–12%） | Interested-party | 年報全文機器人收入揭露 |
| 人形空心杯份額（2030） | {I['hs30_A']:.0%}／{I['hs30_B']:.0%}／{I['hs30_C']:.0%} | Assumed | 因時、兆威等靈巧手供應鏈揭露 |
| 全球人形出貨（線 4） | 線 4 高／基準／低；× hu_k {I['hu_k']:.0f}（高案 2） | Assumed（線 4） | IDC、TrendForce 季出貨 |
| 有效稅率 | 3%／{I['tax_r']:.0%}／15% | Assumed | 高新技術企業資格 |
| 同業前瞻 P/E | 中位數 {I['pe_peer']:.1f} 倍（6 家） | Derived | 同一估值日重抓 |

## 敏感度（整本重算，前 10 大）

| 案例 | 鍵 | 加權目標價 | 相對基準 | DCF 目標價 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | `{c['key']}` | {c['tp']:,.1f} | {c['tp'] / sens['base_tp'] - 1:+.1%} | {c['dcf']:,.1f} | {c['rating']} |" for c in top) + f"""

完整案例見 xlsx Sens 頁（每個 In 頁有低／高區間的鍵都有案例；另有情境 B 改 A／C 值、人形出貨 5 倍、情境權重案例）。**翻轉評等所需的變動**：見 Recon 反解；升一級需加權目標價達 {fn(up_b)} 元。

## 反解（Recon 頁）

""" + "\n".join(f"- {a}：{fn(b, '{:,.3f}')}（基準 {fn(c0, '{:,.2f}')}）；{d}" for a, b, c0, d in sol) + f"""

## 限制與「分不出來／沒找到」

- **分不出來（資料性質）：** 機器人應用與人形空心杯用營收公司未揭露絕對值；人形量級屬未發生事件，線 4 三情境差距約 4–5 倍。
- **沒找到（工作未完）：** 2025 年年報全文（只取得摘要與媒體轉錄；電源與照明、設備狀態管理、貿易分項營收、前五大客戶、2023 年分部營收）；2023、2024 年少數股東損益與費用明細（以核心營業利益殘差處理）；賣方目標價；可轉債與期權；受限資金明細；2026 年第 4 季指引。
- **資料商口徑：** 毛利、核心營業利益、折舊攤銷、資本支出、股利取自 stockanalysis.com（資料商，非公司原文），標 Interested-party；2026 年三季報、年報摘要為公司原文。
- **線模型：** 空心杯每台用量、單價、年降皆直接讀線 6，無不一致；線 6 未給本公司份額，本模型自設（Assumed）。見 xlsx「對帳」頁。
- **同業倍數：** 4 家 A 股同業原報價日 2026-09-30～10-30，前瞻 P/E 依 2026-11-20 收盤等比例換算（EPS 未更新）；哈默納科、Nabtesco 取本專案模型值；Maxon（未上市）、Regal Rexnord、Nidec 報價日未取得，未進中位數；少於 8 家不做迴歸。綠的諧波本身題材倍數極高，中位數對它不敏感（6 家取中位數）。
- **rf 逾期：** 中國 10 年期國債 1.61% 為 2026-11-05 值（跨模型一致）。
- **協調檔：** 實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics；相關條文不適用。
- **獨立核對：** 依專案規則，交付判讀的數字應由未參與製作的代理獨立核對；待外層流程安排。

## 資料來源

見 xlsx Sources 頁（公司公告、券商報告、資料商、線 4／6 成品）。

## 版本紀錄

- {VER}（2026-11-24）：首版。現價 {I['px']:,.2f}（2026-11-20 收盤）；加權目標價 {g('tp'):,.1f}；上檔空間 {g('up'):+.1%}；{g('rating')}／{g('rec')}；收支關卡 {g('gate_b'):,.1f}（{g('gate_t')}）。
"""
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write(t)


if __name__ == "__main__":
    main()

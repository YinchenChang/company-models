# 恒立液壓 Hengli Hydraulic（601100.SS，上海主板）個股模型 v0.1 建置腳本
# 產出：physical_ai/hengli/dist/20261124_恒立液壓個股模型_v0.1.xlsx（LibreOffice 重算後的成品，含公式與快取值）
# 結構：命題模組（Thesis：液壓油缸、泵閥及馬達、液壓系統、配件鑄件及其他；人形用滾柱絲杠＝線 4 絲杠需求 × 線 4 隱含單價 × 份額）
#       → 官方產品分部 → 標準損益表 → FCFF DCF＋同業 P/E。
# 敏感度：每個案例覆寫 In 頁基準欄後整本重建、LibreOffice 重算（整本重算，不用閉式近似）；反解同法以二分搜尋。
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md、docs/plan/高不確定性預估工作原則.md。
#       實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics 協調檔。
# 用法：python3 physical_ai/hengli/build_hengli.py   （需要 soffice；暫存於家目錄下 hl_build）
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
NAME = f"20261124_恒立液壓個股模型_{VER}.xlsx"
OUT = os.path.join(HERE, "dist", NAME)
HOMEDIR = os.path.expanduser("~")
TMP = os.path.join(HOMEDIR, "hl_build")
LO = ["soffice", f"-env:UserInstallation=file://{HOMEDIR}/lo_hengli", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML"]

L4 = os.path.join(ROOT, "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx")
L6 = os.path.join(ROOT, "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx")
L3N = "physical_ai/lines/industrial_robot（本模型未使用）"
L4N, L6N = (os.path.relpath(p, ROOT) for p in (L4, L6))

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDRF = PatternFill("solid", fgColor="DDEBF7")
SECF = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------- 出處 ----------------
S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv：601100.SS；未還原股利）"
S_SA = "stockanalysis.com SHA:601100（資料商；年度與季度財務報表、資產負債表、現金流量表、statistics 頁，2026-11-24 讀取；法定財報原文未直接讀取）"
S_SA_ST = "stockanalysis.com SHA:601100 statistics（資料商；頁面報價 94.00 元為 2026-11-24 盤中，只作對照）"
S_Q3 = "長江證券 恒立液壓 2026 年三季報點評（https://m.gelonghui.com/news/5344583；2026-10-30 後）：轉錄公司 2026 年三季報（營收 102.33 億、歸母淨利 21.54 億、Q3 毛利率 40.19%）"
S_ZY = "中郵證券 恒立液壓 2026 年三季報點評（https://data.eastmoney.com/report/info/AP202611021831332773.html；2026-11-02）：期間費用率、2026–2028 年預測"
S_Y25 = "券商 2025 年報點評摘要（https://data.eastmoney.com/report/info/AP202604231821483078.html，2026-04-23；網頁搜尋摘要，PDF 原文未讀）：轉錄公司 2025 年報分產品營收（油缸 52.54 億 +10.36%、泵閥及馬達 43.26 億 +20.72%、液壓系統 3.85 億 +30.00%、配件及鑄件 8.91 億 +30.31%）"
S_H1 = "券商 2026 年半年報點評摘要（https://pdf.dfcfw.com/pdf/H3_AP202609091829178059_1.pdf 等；網頁搜尋摘要）：上半年營收 68.33 億（+32.13%）、毛利率 39.27%；線性驅動募投累計投入 12.79 億、進度 91.38%、預計 2026-12 達可使用狀態"
S_DW = "東吳證券 2026 年三季報點評（網頁搜尋摘要）：2026–2028 年歸母淨利預測 31.6／42.8／52.1 億元"
S_GX = "國信證券 2026 年三季報點評（網頁搜尋摘要）：2026–2028 年歸母淨利預測 31.08／38.55／46.89 億元（下調，理由為匯兌損益）"
S_RF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與 innolight、eoptolink、victory_giant、leaderdrive 同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：China CRP 0.91%（與 leaderdrive 同一取值）"
S_CLOSE = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
S_SAP = "各同業模型 Comp／In 頁（原值來源 stockanalysis.com，原報價日見各列）"
IP_BRK = "券商研究報告：有銷售研究與維持評等的誘因；歷史數字為轉錄，預測為券商自設"
IP_SA = "資料商彙整賣方共識，財報前的值未必更新"

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
    _, l4 = _line_rows(L4, 1)
    _, l6 = _line_rows(L6, 1)
    d = {"scr": {}, "sp": {}}
    for y in range(2026, 2033):
        r, v = l4[f"SCREW_UNITS_{y}"]
        for s, c, col in (("A", 4, "E"), ("B", 3, "D"), ("C", 2, "C")):
            d["scr"][(y, s)] = (v[c], f"{L4N} 下游取數!{col}{r}（SCREW_UNITS_{y}）")
    # 線 4 隱含滾柱絲杠單價（美元）＝ SCREW_MKT ÷ SCREW_UNITS，2026／2028／2030 三點；其間對數線性內插，2031–2032 沿 2028–2030 年降幅外推
    for s, c, col in (("A", 4, "E"), ("B", 3, "D"), ("C", 2, "C")):
        pts = {}
        for y in (2026, 2028, 2030):
            rm, m = l4[f"SCREW_MKT_{y}"]; ru, u = l4[f"SCREW_UNITS_{y}"]
            pts[y] = m[c] * 1e6 / u[c]
        src = f"{L4N} 下游取數!{col}（SCREW_MKT_2026／2028／2030 ÷ SCREW_UNITS_2026／2028／2030，Derived；其間內插、2031–32 外推）"
        for y in range(2026, 2033):
            if y <= 2028: p = pts[2026] * (pts[2028] / pts[2026]) ** ((y - 2026) / 2)
            elif y <= 2030: p = pts[2028] * (pts[2030] / pts[2028]) ** ((y - 2028) / 2)
            else: p = pts[2030] * (pts[2030] / pts[2028]) ** ((y - 2030) / 2)
            d["sp"][(y, s)] = (p, src)
    for k in ("行星滾柱絲杠市場規模 2026（基準單價口徑＝覓途）", "行星滾柱絲杠市場規模 2030（基準單價口徑＝覓途）", "行星滾柱絲杠市場規模 2032（基準單價口徑＝覓途）"):
        r, v = l6[k]; d[k[:15]] = (v[3], f"{L6N} 下游取數!D{r}")
    r, v = l6["行星滾柱絲杠需求量 2030"]; d["scr30_6"] = (v[3], f"{L6N} 下游取數!D{r}")
    r, v = l4["Q_SCREW_AVG"]; d["q_screw"] = (v[3], f"{L4N} 下游取數!D{r}（Q_SCREW_AVG）")
    r, v = l4["P_ROLLER_SCREW"]; d["p_screw4"] = (v[3], f"{L4N} 下游取數!D{r}（P_ROLLER_SCREW）")
    return d


LD = line_data()

# ---------------- In 頁輸入 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 來源誘因, 說明)
IN = [
    ("sec", "1. 市場資料與日期"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, "", "資料截止日 2026-11-24（財報、券商預測、同業倍數）；股價 2026-11-20；成品日 2026-11-24"),
    ("px", "現價（2026-11-20 收盤）", "元", None, 88.82, None, "Verified", S_PX, "", "csv 值 88.82；11-24 資料商頁面 94.00（盤中，不採用）"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,12,31)", None, "Assumed", "本專案慣例：評價日 2026-12-31（會計年度＝日曆年，FY2026 在評價日已結束）", "", "折現期間由評價日起算"),
    ("tp_date", "目標時點（px_date 後 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", ""),
    ("bs_date", "資產負債表日", "日期", None, "=DATE(2026,9,30)", None, "Verified", S_SA, "", "2026 年三季報"),
    ("stub_frac", "Stub 期（資產負債表日→評價日）", "年", None, "=(val_date-bs_date)/365", None, "Derived", "＝(評價日 − 資產負債表日) ÷ 365", "", "Stub FCFF＝FY2026 FCFF × stub_frac"),
    ("sh_v", "評價日股數（總股本）", "百萬股", None, 1341.0, None, "Verified", S_SA + "：股數 1,341 百萬股（statistics 頁 1.34 十億）", "", "限制性股票、選擇權未取得揭露，以 0 計"),
    ("cash_v", "現金＋短期投資＋交易性金融資產（2026-09-30）", "百萬元", None, 10418.0, None, "Verified", S_SA + "：現金 10,331、交易性金融資產 86.81", "", "2025-12-31 為 9,216"),
    ("ltinv_v", "長期投資（2026-09-30）", "百萬元", None, 16.43, None, "Verified", S_SA, "", "金額小；明細未取得"),
    ("inv_k", "長期投資計入淨現金的比例", "%", 0.5, 1.0, None, "Assumed", "基準全額計入；低案 50%", "", ""),
    ("debt_v", "有息負債（2026-09-30；含租賃負債）", "百萬元", None, 919.9, None, "Verified", S_SA + "：短期借款 848.57、長期借款 60.07、租賃 11.26", "", "短期借款自 2025-12 的 13.5 升至 848.6（用途未查）"),
    ("nci_v", "少數股東權益（2026-09-30）", "百萬元", None, 59.9, None, "Verified", S_SA, "", "自淨現金扣除"),
    ("fx", "美元兌人民幣（2026-11-20）", "元／美元", None, 6.6665, None, "Verified", S_CLOSE + "：CNY=X", "", "用於人形絲杠市場（美元）換算人民幣"),
    ("cons_tp", "賣方平均目標價（23 家，頁面 2026-11-24）", "元", None, 112.40, None, "Interested-party", S_SA_ST, IP_SA, "只作對照，不進模型；現價高於 11-20 收盤"),
    ("sec", "2. 實績錨（FY2026＝2026 年，會計年度＝日曆年）"),
    ("r9m26", "2026 年 1–9 月營收", "百萬元", None, "=Q!E4", None, "Derived", "＝Q 頁 Q1＋Q2＋Q3（10,233；" + S_Q3 + "：102.33 億，+31.36%，兩者一致）", "", ""),
    ("r_q425", "2025 年第 4 季營收", "百萬元", None, "=Q!F4", None, "Derived", "＝Q 頁 2025Q4（資料商季報 3,151；＝FY2025 10,941 − 1–9 月 7,790）", "", ""),
    ("q4g", "2026 年第 4 季營收年增", "%", 0.05, 0.15, 0.30, "Assumed", "Q1–Q3 年增 32.5%／31.8%／29.8%；中郵 2026E 營收 137.64 億隱含第 4 季約 +12%（" + S_ZY + "）；2025 年第 4 季基期已高（+28%）", "", "FY2026 營收＝1–9 月實績＋2025Q4 ×（1＋q4g）"),
    ("sh_cyl26", "FY2026 液壓油缸占營收", "%", 0.44, 0.48, 0.50, "Assumed", "2025 年占 48.0%（券商 2025 年報點評摘要，出處見 Sources 頁）；2026 年分產品數字未取得（半年報、三季報摘要無分產品）", "", "泵閥與非挖增速較高，占比可能下降"),
    ("sh_pump26", "FY2026 液壓泵閥及馬達占營收", "%", 0.38, 0.40, 0.45, "Assumed", "2025 年占 39.5%；2025 年成長 20.7% 高於油缸 10.4%", "", ""),
    ("sh_sys26", "FY2026 液壓系統占營收", "%", None, 0.035, None, "Assumed", "2025 年占 3.5%", "", "配件、鑄件及其他＝餘數（扣除人形絲杠）"),
    ("gm26", "FY2026 毛利率", "%", None, "=Q!E5/Q!E4", None, "Derived", "＝1–9 月毛利 ÷ 營收（Q 頁；39.58%，與 " + S_Q3 + " 一致）", "", "第 4 季假設同 1–9 月"),
    ("opex26", "FY2026 營業費用（含稅金及附加、研發，毛利 − 營業利益）÷ 營收", "%", 0.12, "=(Q!E5-Q!E6)/Q!E4", 0.15, "Derived", "＝(1–9 月毛利 − 營業利益) ÷ 營收（Q 頁；13.6%）；中郵：銷售 2.26%、管理 5.43%、研發 5.33%（" + S_ZY + "）", "", "資料商營業利益口徑；財務費用在業外"),
    ("nonop9m", "2026 年 1–9 月業外（稅前利益 − 營業利益）", "百萬元", None, "=Q!E7-Q!E6", None, "Derived", "＝Q 頁稅前（資料商以利潤率推算）− 營業利益；匯兌損失、外匯掉期公允價值下降（" + S_Q3 + "）", "", "財務費用率 1.78%（中郵）"),
    ("nonop_q4", "2026 年第 4 季業外", "百萬元", -150, 0.0, 150, "Assumed", "1–9 月 −280；匯兌與掉期損益無法預測，基準假設第 4 季打平", "", "券商 2026E 淨利 29.5–31.6 億隱含第 4 季利潤回升"),
    ("sec", "3. 命題模組：產品分部（驅動＝年增速逐年遞減）與人形滾柱絲杠（線 4、線 6）"),
    ("fade", "分部年增速每年遞減係數（FY2027 起首年 ×1，其後逐年 ×fade）", "x", 0.70, 0.85, 0.95, "Assumed", "券商 2027E／2028E 營收年增 20.9%／16.8%（" + S_ZY + "），隱含逐年遞減約 0.8", "", "同一係數用於油缸、泵閥、系統、配件"),
    ("g_sys", "液壓系統營收年增（FY2027 起首年）", "%", 0.08, 0.18, 0.30, "Assumed", "2025 年 +30.0%、基期小（3.85 億）", "", ""),
    ("g_parts", "配件、鑄件及其他（不含人形絲杠）營收年增（FY2027 起首年）", "%", 0.05, 0.12, 0.20, "Assumed", "2025 年配件及鑄件 +30.3%", "", ""),
    ("hu_k", "人形出貨相對線 4 的倍數（全部情境）", "x", None, 1.0, 2.0, "Assumed", "基準＝線 4 原值；高案 2 倍（中國人形 2026 產量可能高於線 4 基準，未查核）", "", "反解變數"),
    ("hum26", "FY2026 人形機器人用滾柱絲杠營收", "百萬元", 2, 10, 40, "Assumed", "公司未揭露；線性驅動項目 2025 年具備滾珠絲杠年產 7 萬套、客戶逾 300 家（券商 2025 年報點評摘要），行星滾柱絲杠切入人形僅見券商定性描述", "", "對照：線 4 基準 2026 人形絲杠市場約 1,085 百萬元；隱含份額見 Thesis"),
    ("sec", "4. 命題模組：情境專屬驅動（A 樂觀／B 基準／C 保守；只列基準欄，不設敏感度區間）"),
    ("gcyl_A", "情境 A：液壓油缸營收年增（FY2027 起首年）", "%", None, 0.14, None, "Assumed", "2025 年 +10.4%；2026 上半年挖掘機銷量 +26.4%（國信摘要）；出口 +33.5%", "", "挖掘機週期上行延續"),
    ("gcyl_B", "情境 B：液壓油缸營收年增（FY2027 起首年）", "%", None, 0.10, None, "Assumed", "同上；2026 年 Q3 挖掘機銷量 +15.8%（長江摘要）", "", ""),
    ("gcyl_C", "情境 C：液壓油缸營收年增（FY2027 起首年）", "%", None, 0.03, None, "Assumed", "同上；週期見頂下行", "", ""),
    ("gpump_A", "情境 A：液壓泵閥及馬達營收年增（FY2027 起首年）", "%", None, 0.25, None, "Assumed", "2025 年 +20.7%；東吳稱挖機泵閥份額提升、非挖多點開花", "", ""),
    ("gpump_B", "情境 B：液壓泵閥及馬達營收年增（FY2027 起首年）", "%", None, 0.20, None, "Assumed", "同上", "", ""),
    ("gpump_C", "情境 C：液壓泵閥及馬達營收年增（FY2027 起首年）", "%", None, 0.10, None, "Assumed", "同上", "", ""),
    ("hs30_A", "情境 A：2030 年占全球人形滾柱絲杠需求份額", "%", None, 0.06, None, "Assumed", "2026 隱含份額見 Thesis（基準約 0.9%）；歐系與陸系專業絲杠廠領先（線 6 份額頁）", "", "2026 隱含份額線性走到 2030，其後持平"),
    ("hs30_B", "情境 B：2030 年占全球人形滾柱絲杠需求份額", "%", None, 0.03, None, "Assumed", "同上", "", ""),
    ("hs30_C", "情境 C：2030 年占全球人形滾柱絲杠需求份額", "%", None, 0.01, None, "Assumed", "同上", "", ""),
    ("gm30_A", "情境 A：FY2030 毛利率", "%", None, 0.41, None, "Assumed", "2025 年 40.6%（Hist）、2026 年 1–9 月 39.6%（海外產能爬坡）", "", "FY2026 起線性走到 FY2030，其後持平"),
    ("gm30_B", "情境 B：FY2030 毛利率", "%", None, 0.385, None, "Assumed", "同上", "", ""),
    ("gm30_C", "情境 C：FY2030 毛利率", "%", None, 0.34, None, "Assumed", "同上；週期下行與價格競爭", "", ""),
    ("tp_wA", "情境權重 A", "", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", ""),
    ("tp_wB", "情境權重 B", "", None, 0.50, None, "Assumed", "同上", "", ""),
    ("tp_wC", "情境權重 C", "", None, 0.25, None, "Assumed", "同上", "", ""),
    ("sec", "5. 損益與現金流"),
    ("opex30", "FY2030 營業費用（含稅金及附加、研發）÷ 營收", "%", 0.12, 0.13, 0.15, "Assumed", "2023–2025 年 14.3%／16.6%／15.3%（Hist）；2026 年 1–9 月 13.6%", "", "FY2026→FY2030 線性，其後持平"),
    ("inv_y", "淨現金收益率（業外，FY2027 起）", "%", 0.01, 0.02, 0.03, "Assumed", "2025 年業外 264 ÷ 年底現金 9,216＝2.9%（Derived，含匯兌與公允價值）；2026 年為負", "", "業外；DCF 另以淨現金計入，不重複"),
    ("tax_r", "有效稅率", "%", 0.09, "=Hist!E17/Hist!E15", 0.15, "Derived", "＝2025 年所得稅 ÷ 稅前利益（資料商推算稅前；高新技術企業 15% 加研發加計扣除）", "", "敏感度 9%／15%（Assumed）"),
    ("payout", "配息率", "%", None, "=Hist!E20/Hist!E19", None, "Derived", "＝2025 年每股股利 0.86 元 ÷ EPS 2.04 元（" + S_SA + "）", "", "用於收支關卡與目標時點前股利"),
    ("da26", "FY2026 折舊攤銷 ÷ 營收", "%", None, "=Hist!E22/Hist!E4", None, "Derived", "＝2025 年折舊攤銷 527 ÷ 營收", "", "之後線性收斂到資本支出比例 × da_ss（FY2032）"),
    ("da_ss", "穩態折舊 ÷ 資本支出", "x", None, 0.8, None, "Assumed", "穩態資本支出高於折舊（成長）", "", ""),
    ("cx26", "FY2026 資本支出", "百萬元", 700, 900, 1100, "Assumed", S_SA + "：近 12 個月資本支出 843、FY2025 924；線性驅動募投 12.79 億已投入 91%（" + S_H1 + "）", "", ""),
    ("cx_r", "資本支出 ÷ 營收（FY2027 起）", "%", 0.06, 0.075, 0.10, "Assumed", "FY2023 15.2%、FY2024 11.4%、FY2025 8.4%（Hist，Derived）；印尼廠建設中", "", ""),
    ("nwc_r", "營運資金增加 ÷ 營收增量", "%", 0.15, 0.25, 0.40, "Assumed", S_SA + "：2025-12→2026-09 存貨 +543、應收 +437、應付 +427（淨 +553），營收年增量約 2,500", "", ""),
    ("sec", "6. 折現率（DCF 標準規格第 2 節）"),
    ("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", S_RF, "", "日期 2026-11-05，較股價基準日早 11 個交易日（逾期；跨模型一致，下次全面改版更新）"),
    ("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", ""),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", ""),
    ("b_own", "本公司 5 年 β（原始）", "x", None, 0.88, None, "Verified", S_SA_ST, "", "上銀模型 Comps 同為 0.87（較早日期）"),
    ("b_peer", "同業 5 年 β 中位數（原始）", "x", None, "=MEDIAN(Comp!H5:H10)", None, "Derived", "＝Comp 頁核心同業 6 家 β 中位數", "", ""),
    ("beta", "β（兩來源 Blume 調整後中位數）", "x", "=beta-0.3", "=MEDIAN(0.67*b_own+0.33,0.67*b_peer+0.33)", "=beta+0.3", "Derived", "Blume 0.67β＋0.33；兩來源＝本公司、同業中位數", "", "敏感度 ∓0.3（跨模型慣例）"),
    ("kd", "稅前債務成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt", "", "有息負債占市值不到 1%"),
    ("wacc_adj", "WACC 調整（敏感度用）", "pt", -0.01, 0, 0.01, "Assumed", "基準 0；敏感度 ±1pt", "", ""),
    ("g_cap", "永續成長率上限", "%", None, 0.03, None, "Assumed", "DCF 標準規格：3%，且不超過計價幣別 10 年期公債殖利率", "", ""),
    ("g_tv", "永續成長率", "%", 0.01, "=MIN(g_cap,rf)", 0.03, "Derived", "＝MIN(g_cap, rf)", "", "敏感度 1%／3%（3% 超過 rf，只作敏感度）"),
    ("ronic_sp", "RONIC − WACC", "pt", None, 0, 0.05, "Assumed", "RONIC 預設＝WACC（新投資價值中性）", "", ""),
    ("tv_alt", "終值改用 FCFF ×（1＋g）（0＝否，1＝是）", "開關", None, 0, 1, "Assumed", "DCF 標準規格第 4 節敏感度", "", ""),
    ("sec", "7. 同業倍數與權重"),
    ("pe_peer", "同業前瞻 P/E 中位數", "x", "=MIN(Comp!F5:F10)", "=MEDIAN(Comp!F5:F10)", "=MAX(Comp!F5:F10)", "Derived", "＝Comp 頁核心同業 6 家前瞻 P/E 中位數；股價皆為 2026-11-20 收盤（取自同日各同業模型值）", "", "敏感度＝同業最低／最高"),
    ("rel_k", "相對倍數係數（本公司 ÷ 同業）", "x", None, 1.0, 2.0, "Assumed", "1.00（未調整：缺本公司對同業十年相對倍數歷史）", "", "高案 2 倍（約等於現行相對倍數，未查核，不作依據）"),
    ("w_dcf", "DCF 權重（Comparable＝1 − 本值）", "%", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設 DCF 50%＋同業倍數 50%", "", ""),
]
SCN_DRIVERS = [("gcyl", "液壓油缸營收年增"), ("gpump", "液壓泵閥及馬達營收年增"), ("hs30", "2030 年人形絲杠份額"), ("gm30", "FY2030 毛利率")]

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


    # ================= Hist（年度；百萬元；資料商口徑，中國會計準則）=================
    HI["A1"] = "Hist：歷史財務（百萬元人民幣；資料商彙整中國會計準則報表）。營業利益為資料商口徑；稅前利益為資料商以稅前利潤率推算，精度約 ±0.5%。公式只引用本頁（定義名稱 h_*）。"
    hdr(HI, 3, ["項目", "", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"])
    HROWS = [  # (列, 鍵, 項目, [2023, 2024, 2025], 標記, 來源)
        (4, "h_rev", "營業收入", [8985, 9390, 10941], "Verified", S_SA + "；2025 另見券商 109.41 億（+16.52%）"),
        (5, "h_cogs", "營業成本（＝營收 − 毛利）", [5290, 5422, 6498], "Verified", S_SA + "（毛利 3,695／3,968／4,443；2025 毛利率 41.58% 與券商一致）"),
        (6, "h_gp", "毛利（＝營收 − 營業成本）", ["=C4-C5", "=D4-D5", "=E4-E5"], "Derived", "＝營業收入 − 營業成本"),
        (7, "h_gm", "毛利率", ["=C6/C4", "=D6/D4", "=E6/E4"], "Derived", "＝毛利 ÷ 營收"),
        (8, "h_seg_cyl", "液壓油缸 營收", [None, "=E8/1.1036", 5254], "Verified", S_Y25 + "；2024＝2025 ÷ 1.1036（Derived）；2023 未取得"),
        (9, "h_seg_pump", "液壓泵閥及馬達 營收", [None, "=E9/1.2072", 4326], "Verified", S_Y25 + "；2024＝2025 ÷ 1.2072（Derived）；2023 未取得"),
        (10, "h_seg_sys", "液壓系統 營收", [None, "=E10/1.30", 385], "Verified", S_Y25 + "；2024＝2025 ÷ 1.30（Derived）；2023 未取得"),
        (11, "h_seg_oth", "配件、鑄件及其他業務 營收（＝營收 − 上三項）", [None, "=D4-D8-D9-D10", "=E4-E8-E9-E10"], "Derived", "配件及鑄件 2025 年 891（+30.31%）；餘額為其他業務收入（約 85）；以殘差計"),
        (12, "h_op", "營業利益（資料商口徑）", [2408, 2413, 2768], "Verified", S_SA),
        (13, "h_opex", "營業費用（含稅金及附加、研發；＝毛利 − 營業利益）", ["=C6-C12", "=D6-D12", "=E6-E12"], "Derived", "銷售、管理、研發、稅金及附加的淨額（含其他收益）；細項未取得（券商摘要：2025 年銷售＋管理＋財務 91 億元＝8.31%）"),
        (14, "h_opm", "營業利益率", ["=C12/C4", "=D12/D4", "=E12/E4"], "Derived", ""),
        (15, "h_pbt", "稅前利益（資料商推算）", [2830, 2801, 3032], "Interested-party", S_SA + "：由稅前利潤率 × 營收推算，非報表直讀", "資料商推算"),
        (16, "h_nonop", "業外（＝稅前 − 營業利益）", ["=C15-C12", "=D15-D12", "=E15-E12"], "Derived", "利息收入、匯兌、公允價值變動、營業外收支合計"),
        (17, "h_tax", "所得稅（＝稅前 − 淨利）", ["=C15-C18", "=D15-D18", "=E15-E18"], "Derived", "稅前為資料商推算，稅額含推算誤差"),
        (18, "h_ni", "歸屬母公司淨利", [2499, 2509, 2734], "Verified", S_SA + "；2025 另見券商 27.34 億（+8.99%）"),
        (19, "h_eps", "EPS（基本，元）", [1.86, 1.87, 2.04], "Verified", S_SA),
        (20, "h_dps", "每股股利（元）", [0.70, 0.70, 0.86], "Verified", S_SA),
        (21, "h_sh", "股數（＝淨利 ÷ EPS，百萬股）", ["=C18/C19", "=D18/D19", "=E18/E19"], "Derived", "資料商股數 1,341"),
        (22, "h_da", "折舊攤銷", [410.2, 426.24, 526.72], "Verified", S_SA),
        (23, "h_cx", "資本支出", [1366, 1071, 924.2], "Verified", S_SA),
        (24, "h_ocf", "營業現金流", [2677, 2479, 1811], "Verified", S_SA),
    ]
    for hr in HROWS:
        (rr, k, lab, vals, tag, src) = hr[:6]
        inc = hr[6] if len(hr) > 6 else ""
        HI.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = HI.cell(rr, 3 + j, v)
                if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
                if k in ("h_gm", "h_opm"): c.number_format = "0.0%"
        HI.cell(rr, 6, tag); HI.cell(rr, 7, src); HI.cell(rr, 8, inc)
        dn(wb, k, f"Hist!$C${rr}:$E${rr}")
    HI.cell(26, 1, "檢查：分部合計 − 營業收入（應為 0；2023 分部未取得）")
    for c in "DE": HI[f"{c}26"] = f"={c}8+{c}9+{c}10+{c}11-{c}4"
    HI.column_dimensions["A"].width = 58; HI.column_dimensions["G"].width = 90

    # ================= Q（季度；一列一個數字，單一標記）=================
    Q["A1"] = "Q：2025Q4 與 2026 年季度資料（百萬元）。資料商季度報表；稅前利益為資料商以稅前利潤率推算。"
    hdr(Q, 3, ["項目", "2026Q1", "2026Q2", "2026Q3", "2026 年 1–9 月", "2025Q4", "標記", "來源", "來源誘因", "說明"])
    QROWS = [
        (4, "營業收入", [3210, 3623, 3400, "=SUM(B4:D4)", 3151], "Verified", S_SA + "；" + S_Q3, "", "1–9 月合計 10,233＝券商 102.33 億"),
        (5, "毛利", [1238, 1384, 1428, "=SUM(B5:D5)", 1175], "Verified", S_SA, "", "1–9 月毛利率 39.58%（券商同值）；Q3 40.19%"),
        (6, "營業利益（資料商口徑）", [795.36, 928.67, 932.38, "=SUM(B6:D6)", 694.24], "Verified", S_SA, "", ""),
        (7, "稅前利益（資料商推算）", [737, 858, 781, "=SUM(B7:D7)", 696], "Interested-party", S_SA, "資料商推算", "利潤率 × 營收，非報表直讀"),
        (8, "歸屬母公司淨利", [652.19, 784.15, 717.29, "=SUM(B8:D8)", 647.27], "Verified", S_SA + "；" + S_Q3, "", "1–9 月 2,154（+3.20%）"),
        (9, "營業收入年增", [0.3252, 0.3178, 0.2984, 0.3136, 0.2841], "Verified", S_SA + "；" + S_Q3, "", ""),
    ]
    for (rr, lab, vals, tag, src, inc, note) in QROWS:
        Q.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            c = Q.cell(rr, 2 + j, v)
            if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
            if rr == 9: c.number_format = "0.00%"
        Q.cell(rr, 7, tag); Q.cell(rr, 8, src); Q.cell(rr, 9, inc); Q.cell(rr, 10, note)
    Q.column_dimensions["A"].width = 30; Q.column_dimensions["H"].width = 90

    # ================= 線取數（線模型下游取數；藍字＝讀入值，M 欄寫來源檔與儲存格）=================
    LN["A1"] = "線取數：本模型讀入的線層級變數年度路徑（建置時由線模型成品依 ID 讀值）。情境對應：A＝線模型高情境、B＝基準、C＝低情境。"
    hdr(LN, 3, ["項目", "情境", ""] + [f"FY{y}" for y in range(2025, 2033)] + ["單位", "來源（檔案與儲存格）"])
    LR = {}
    r = 4
    for key, lab, unit, dd in (("scr", "全球人形機器人滾柱絲杠需求（線 4）", "支", LD["scr"]),
                               ("sp", "人形用滾柱絲杠隱含單價（線 4 市場 ÷ 需求；Derived）", "美元／支", LD["sp"])):
        for s in SC:
            LN.cell(r, 1, lab); LN.cell(r, 2, f"{s}（{LINE_SC[s]}）")
            srcs = []
            for y in range(2026, 2033):
                v, src = dd[(y, s)]
                c = LN.cell(r, 4 + (y - 2025), v); c.font = BLUE; c.number_format = "#,##0"
                srcs.append(src)
            LN.cell(r, 12, unit); LN.cell(r, 13, f"{srcs[0]} … {srcs[-1]}")
            LR[(key, s)] = r; r += 1
    r += 1
    LN.cell(r, 1, "線 6 行星滾柱絲杠市場規模（基準單價口徑；百萬美元；供 Recon 對照）").font = BOLD; r += 1
    MR = {}
    for k_, y_ in (("行星滾柱絲杠市場規模 2026", 2026), ("行星滾柱絲杠市場規模 2030", 2030), ("行星滾柱絲杠市場規模 2032", 2032)):
        v, src = LD[k_]
        LN.cell(r, 1, k_); LN.cell(r, 2, "B（基準）"); c = LN.cell(r, 4 + (y_ - 2025), v); c.font = BLUE; c.number_format = "#,##0"
        LN.cell(r, 12, "百萬美元"); LN.cell(r, 13, src); MR[y_] = r; r += 1
    LN.column_dimensions["A"].width = 56; LN.column_dimensions["M"].width = 110

    # ================= Thesis（命題模組）=================
    TH["A1"] = ("命題模組：油缸、泵閥及馬達、液壓系統、配件鑄件及其他各以 FY2026 錨（1–9 月實績＋第 4 季假設 × 分產品占比）起算、FY2027 起以逐年遞減的年增速外推；"
                "人形滾柱絲杠營收＝線 4 絲杠需求 × hu_k × 線 4 隱含單價 × 匯率 × 恒立份額。百萬元。")
    TH["A2"] = "欄：D＝FY2025 … K＝FY2032。人形 2026 隱含份額以線 4 基準路徑校準（三情境同一實績年）。"
    hdr(TH, 3, ["項目", "說明", ""] + [f"FY{y}" for y in range(2025, 2033)])
    THR = {}
    r = 4
    for s in SC:
        TH.cell(r, 1, f"情境 {s}").font = BOLD; TH.cell(r, 1).fill = SECF; r += 1
        R = {}
        R["rev26"] = r; TH.cell(r, 1, "FY2026 合併營收（錨）"); TH.cell(r, 2, "＝1–9 月＋2025Q4 ×（1＋q4g）")
        TH[f"{TC[2026]}{r}"] = "=r9m26+r_q425*(1+q4g)"; r += 1
        R["scr"] = r; TH.cell(r, 1, "全球人形滾柱絲杠需求（線 4 × 出貨倍數，千支）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR[('scr', s)]}*hu_k/1000"
        r += 1
        R["sp"] = r; TH.cell(r, 1, "人形用滾柱絲杠單價（美元／支，線 4 隱含）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR[('sp', s)]}"
        r += 1
        R["mkt"] = r; TH.cell(r, 1, "全球人形滾柱絲杠市場（百萬元人民幣）"); TH.cell(r, 2, "＝需求（千支）× 單價 × 匯率 ÷ 1000")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['scr']}*{TC[y]}{R['sp']}*fx/1000"
        r += 1
        R["hsh"] = r; TH.cell(r, 1, "人形滾柱絲杠份額"); TH.cell(r, 2, "FY2026＝隱含（hum26 ÷ 線 4 基準市場）；線性走到 hs30，其後持平")
        TH[f"{TC[2026]}{r}"] = f"=hum26/(線取數!{TC[2026]}{LR[('scr', 'B')]}*hu_k*線取數!{TC[2026]}{LR[('sp', 'B')]}*fx/1e6)"
        for y in range(2027, 2033):
            TH[f"{TC[y]}{r}"] = f"=${TC[2026]}${r}+(hs30_{s}-${TC[2026]}${r})*MIN(1,({y}-2026)/4)"
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"].number_format = "0.00%"
        r += 1
        R["hum"] = r; TH.cell(r, 1, "人形機器人用滾柱絲杠 營收"); TH.cell(r, 2, "FY2026＝hum26；FY2027 起＝市場 × 份額")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = "=hum26" if y == 2026 else f"={TC[y]}{R['mkt']}*{TC[y]}{R['hsh']}"
        r += 1
        def seg(key, lab, share, g, note):
            nonlocal r
            R[key] = r; TH.cell(r, 1, lab); TH.cell(r, 2, note)
            for y in range(2026, 2033):
                c = TC[y]
                TH[f"{c}{r}"] = (f"={c}{R['rev26']}*{share}" if y == 2026 else f"={TC[y - 1]}{r}*(1+{g}*fade^({y}-2027))")
            r += 1
        seg("cyl", "液壓油缸", "sh_cyl26", f"gcyl_{s}", "FY2027 起＝前一年 ×（1＋首年增速 × fade^(年−2027)）")
        seg("pump", "液壓泵閥及馬達", "sh_pump26", f"gpump_{s}", "同上")
        seg("sys", "液壓系統", "sh_sys26", "g_sys", "同上")
        R["parts"] = r; TH.cell(r, 1, "配件、鑄件及其他（不含人形絲杠）"); TH.cell(r, 2, "FY2026＝餘數；FY2027 起同上")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = (f"={c}{R['rev26']}-{c}{R['cyl']}-{c}{R['pump']}-{c}{R['sys']}-{c}{R['hum']}" if y == 2026
                             else f"={TC[y - 1]}{r}*(1+g_parts*fade^({y}-2027))")
        r += 1
        R["partsT"] = r; TH.cell(r, 1, "配件、鑄件及其他業務 合計")
        for y in range(2026, 2033): c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['parts']}+{c}{R['hum']}"
        r += 1
        R["humsh"] = r; TH.cell(r, 1, "人形絲杠占合併營收")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['hum']}/({c}{R['cyl']}+{c}{R['pump']}+{c}{R['sys']}+{c}{R['partsT']})"; TH[f"{c}{r}"].number_format = "0.00%"
        r += 1
        R["gm"] = r; TH.cell(r, 1, "毛利率路徑"); TH.cell(r, 2, "FY2026＝1–9 月；線性到 FY2030 gm30，其後持平")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = "=gm26" if y == 2026 else f"=gm26+(gm30_{s}-gm26)*MIN(1,({y}-2026)/4)"; TH[f"{TC[y]}{r}"].number_format = "0.0%"
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
    PL["A1"] = ("PL：官方產品分部（液壓油缸／液壓泵閥及馬達／液壓系統／配件、鑄件及其他業務）→ 合併營收 → 標準損益表 → FCFF。"
                "C＝FY2023 … E＝FY2025 為實績；F＝FY2026（1–9 月實績＋第 4 季假設）… L＝FY2032。百萬元；中國會計準則（資料商彙整口徑）。")
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
        row("cyl", "液壓油缸 營收", "百萬元", H24(8), lambda y: f"=Thesis!{TC[y]}{th['cyl']}")
        row("pump", "液壓泵閥及馬達 營收", "百萬元", H24(9), lambda y: f"=Thesis!{TC[y]}{th['pump']}")
        row("sys", "液壓系統 營收", "百萬元", H24(10), lambda y: f"=Thesis!{TC[y]}{th['sys']}")
        row("oth", "配件鑄件及其他（不含人形絲杠）（子列）", "百萬元", None, lambda y: f"=Thesis!{TC[y]}{th['parts']}")
        row("hum", "人形機器人用滾柱絲杠（子列）", "百萬元", None, lambda y: f"=Thesis!{TC[y]}{th['hum']}")
        row("parts", "配件、鑄件及其他業務 營收", "百萬元", H24(11), lambda y: f"={C(y)}{R['oth']}+{C(y)}{R['hum']}")
        row("rev", "營收（合併）", "百萬元", H(4), lambda y: f"={C(y)}{R['cyl']}+{C(y)}{R['pump']}+{C(y)}{R['sys']}+{C(y)}{R['parts']}")
        row("chk", "檢查：分部合計 − 合併營收（應為 0）", "百萬元",
            lambda y: None if y == 2023 else f"={C(y)}{R['cyl']}+{C(y)}{R['pump']}+{C(y)}{R['sys']}+{C(y)}{R['parts']}-{C(y)}{R['rev']}",
            lambda y: f"={C(y)}{R['cyl']}+{C(y)}{R['pump']}+{C(y)}{R['sys']}+{C(y)}{R['parts']}-{C(y)}{R['rev']}")
        PL.cell(R["chk"], 1, "檢查：分部合計 − 合併營收（應為 0）" if s == "B" else f"檢查（情境 {s}）：分部合計 − 合併營收（應為 0）")
        row("chk2", "檢查：子列合計 − 配件鑄件及其他分部（應為 0）", "百萬元", None, lambda y: f"={C(y)}{R['oth']}+{C(y)}{R['hum']}-Thesis!{TC[y]}{th['partsT']}")
        row("yoy", "營收年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1",
            lambda y: f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1", "0.0%")
        row("cogs", "營業成本", "百萬元", H(5), lambda y: f"={C(y)}{R['rev']}*(1-Thesis!{TC[y]}{th['gm']})")
        row("gp", "毛利", "百萬元", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}")
        row("gm", "毛利率", "%", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", "0.0%")
        row("opex", "營業費用（含稅金及附加、研發；扣除其他收益）", "百萬元", H(13), lambda y: f"={C(y)}{R['rev']}*Thesis!{TC[y]}{th['opex']}")
        row("op", "營業利益（資料商口徑）", "百萬元", H(12), lambda y: f"={C(y)}{R['gp']}-{C(y)}{R['opex']}")
        row("opm", "營業利益率", "%", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", "0.0%")
        row("nonop", "其他收入淨額（業外：利息、匯兌、公允價值、營業外）", "百萬元", H(16), lambda y: "=nonop9m+nonop_q4" if y == 2026 else "=(cash_v-debt_v)*inv_y")
        row("pbt", "稅前利益", "百萬元", H(15), lambda y: f"={C(y)}{R['op']}+{C(y)}{R['nonop']}")
        row("tax", "所得稅", "百萬元", H(17), lambda y: f"={C(y)}{R['pbt']}*tax_r")
        row("ni", "歸屬母公司淨利（資料商淨利＝歸母；少數股東損益未另揭露）", "百萬元", H(18), lambda y: f"={C(y)}{R['pbt']}-{C(y)}{R['tax']}")
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
        R["nc"] = r; DC.cell(r, 1, "淨現金滾到評價日（現金與投資 − 有息負債 − 少數股東權益 ＋ Stub FCFF）")
        DC[f"B{r}"] = f"=cash_v+ltinv_v*inv_k-debt_v-nci_v+B{R['stub']}"; r += 1
        R["eq"] = r; DC.cell(r, 1, "股權價值"); DC[f"B{r}"] = f"=B{R['ev']}+B{R['nc']}"; r += 1
        R["ps"] = r; DC.cell(r, 1, "每股價值（評價日，元）"); DC[f"B{r}"] = f"=MAX(0,B{R['eq']}/(sh_v))"; r += 1
        R["dps"] = r; DC.cell(r, 1, "評價日→目標時點間股利（FY2026 盈餘配息，約 2027-06 支付）"); DC[f"B{r}"] = f"=PL!F{PR[s]['eps']}*payout"; r += 1
        R["v12"] = r; DC.cell(r, 1, "推到目標時點（×（1＋ke）^t − 期間股利）"); DC[f"B{r}"] = f"=MAX(0,B{R['ps']}*(1+ke)^((tp_date-val_date)/365)-B{R['dps']})"; r += 2
        DR[s] = R
    for row_ in DC.iter_rows(min_row=4, max_row=r, min_col=2, max_col=7):
        for c in row_:
            if c.number_format == "General": c.number_format = "#,##0.00"
    DC.column_dimensions["A"].width = 70

    # ================= Comp =================
    CP["A1"] = ("Comp：同業前瞻 P/E。股價＝" + S_CLOSE + "（Verified）；前瞻 P/E＝原值 × 2026-11-20 收盤 ÷ 原報價（Derived，EPS 不變）。"
                "核心同業 6 家（皆有 2026-11-20 收盤與可換算的前瞻 P/E）進中位數；少於 8 家不做迴歸。"
                "國內工程機械與液壓同業（三一重工、徐工機械、中聯重科、艾迪精密）不在收盤檔，列「報價日未取得」，未進中位數。")
    hdr(CP, 4, ["公司", "代號", "股價（2026-11-20）", "報價日", "前瞻 P/E 分母", "前瞻 P/E", "近 12 月 P/E", "5 年 β", "會計口徑", "進中位數", "說明", "原報價", "原報價日", "原前瞻 P/E", "標記與換算"])
    PEERS = [  # (公司, 代號, 11-20 收盤, 原報價, 原報價日, 原前瞻 P/E, β, 口徑, 進中位數, 說明, 分母)
        ("Nabtesco", "6268.T", 3105, 3105, "2026-11-20", 13.77, 0.93, "IFRS（未查核）", "是", "建機行走馬達（液壓）與 RV 減速機；取 physical_ai/nabtesco 模型 In 頁 pe_self（11-13 值換算）", "資料商 NTM 共識 EPS"),
        ("SMC", "6273.T", 55300, 55300, "2026-11-20", 14.5941, 0.83, "日本基準（未查核）", "是", "氣動與流體控制元件；取 physical_ai/smc 模型 In 頁 pe_self", "公司 FY2026 指引 EPS"),
        ("THK", "6481.T", 4249, 6116, "2026-10-19", 15.37, 0.76, "日本基準（未查核）", "是", "線性導軌與滾珠絲杠；取 physical_ai/thk 模型值（上銀模型 Comps!F5 換算）", "資料商 NTM 共識 EPS"),
        ("Fanuc", "6954.T", 5047, 5541, "2026-11-13", 23.22, 0.98, "日本基準（未查核）", "是", "機器人與 CNC；取上銀模型 Comps!F12", "資料商 NTM 共識 EPS"),
        ("安川電機", "6506.T", 3180, 4360, "2026-09-16", 21.49, 1.24, "日本基準（未查核）", "是", "伺服與機器人；取上銀模型 Comps!F13", "資料商 NTM 共識 EPS"),
        ("雙環傳動", "002472.SZ", 35.45, 41.21, "2026-10-30", 22.32, 0.52, "中國會計準則", "是", "齒輪與精密減速器；取 physical_ai/leaderdrive 模型 Comp 頁", "資料商 NTM 共識 EPS"),
        ("貝斯特", "300580.SZ", 14.0, None, "2026-09-29", None, 0.60, "中國會計準則", "否", "線性傳動；資料商前瞻 P/E 26.78（2026-09-29），舊報價日收盤不在收盤檔，無法換算至 11-20 → 報價日未取得", "—"),
        ("上銀", "2049.TW", 242.5, None, "2026-10-22", None, 1.08, "台灣 IFRS", "否", "線性滑軌與滾珠螺桿；資料商前瞻 P/E 28.35（2026-10-22），舊報價日收盤不在收盤檔，無法換算 → 報價日未取得", "—"),
        ("恒立液壓（本公司，對照）", "601100.SS", 88.82, 94.00, "2026-11-24", 32.39, 0.88, "中國會計準則", "否", "本公司；資料商 11-24 頁面前瞻 P/E 32.39（價 94.00）換算至 11-20 收盤；只作對照", "資料商 NTM 共識 EPS（三季報後）"),
    ]
    for i, (nm, tk, px1, px0, d0, pe0, bt, acc, med, note, den) in enumerate(PEERS):
        rr = 5 + i
        CP.cell(rr, 1, nm); CP.cell(rr, 2, tk)
        CP.cell(rr, 3, px1).font = BLUE; CP.cell(rr, 4, "2026-11-20"); CP.cell(rr, 5, den)
        if pe0 is not None: CP.cell(rr, 6, f"=N{rr}*C{rr}/L{rr}")
        else: CP.cell(rr, 6, "未取得")
        CP.cell(rr, 8, bt).font = BLUE; CP.cell(rr, 9, acc); CP.cell(rr, 10, med); CP.cell(rr, 11, note)
        if px0 is not None: CP.cell(rr, 12, px0).font = BLUE
        CP.cell(rr, 13, d0)
        if pe0 is not None: CP.cell(rr, 14, pe0).font = BLUE
        CP.cell(rr, 15, "股價 Verified（" + S_CLOSE + "）；前瞻 P/E Derived：＝原值 × 2026-11-20 收盤 ÷ 原報價（分母 EPS 不變）；原值 " + S_SAP + "；β 為資料商 5 年值")
    CP.cell(14, 1, "報價日未取得（不進中位數）：三一重工 600031、徐工機械 000425、中聯重科 000157、艾迪精密 603638（國內工程機械與液壓同業；2026-11-20 收盤不在收盤檔）")
    CP.cell(15, 1, "核心同業前瞻 P/E 中位數"); CP["F15"] = "=MEDIAN(F5:F10)"
    CP.cell(16, 1, "本公司現價 ÷ 資料商前瞻 EPS（換算至 2026-11-20；分母：資料商 NTM 共識，三季報後）"); CP["F16"] = "=F13"
    CP.cell(17, 1, "本公司現價 ÷ 券商 FY2027 EPS 平均（分母：東吳 42.8、國信 38.55、中郵 42.02 億元 ÷ 股數，FY2027）"); CP["F17"] = "=px/(AVERAGE(4280,3855,4202)/sh_v)"
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
    row1("cons_show", "市場共識目標價（資料商 23 家平均，頁面 2026-11-24；只作對照）", "=cons_tp")
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
        SE.cell(r + 1, 1, "說明：折舊、資本支出（FY2027 起）、營運資金隨營收等比例變動（第 44 條）；分部增速逐年遞減係數 fade 同時作用於油缸、泵閥、系統與配件。")
    SE.column_dimensions["A"].width = 52; SE.column_dimensions["B"].width = 44


    # ================= Recon =================
    RC["A1"] = "Recon：模型對實績、券商預測、線 4／線 6 絲杠市場的同口徑對照；反解支持現價所需的主導變數。"
    hdr(RC, 3, ["項目", "模型", "對照", "差距（模型 − 對照）", "說明"])
    rr = 4
    def rrow(lab, a, b, note, fmt="#,##0.0", diff=True):
        nonlocal rr
        RC.cell(rr, 1, lab); RC.cell(rr, 2, a).number_format = fmt
        if b is not None: RC.cell(rr, 3, b).number_format = fmt
        if diff and b is not None: RC.cell(rr, 4, f"=B{rr}-C{rr}").number_format = fmt
        RC.cell(rr, 5, short(note)); rr += 1
    rrow("FY2026 營收（三情境同）vs 中郵 2026E（13,764）", f"=PL!F{PR['B']['rev']}", 13764, S_ZY + "；單一券商（Interested-party）")
    rrow("FY2027 營收（情境 B）vs 中郵 2027E（16,634）", f"=PL!G{PR['B']['rev']}", 16634, "單一券商（Interested-party）")
    rrow("FY2028 營收（情境 B）vs 中郵 2028E（19,433）", f"=PL!H{PR['B']['rev']}", 19433, "單一券商；本模型基準偏保守")
    rrow("FY2026 歸屬母公司淨利 vs 券商 2026E 平均（東吳 3,160、國信 3,108、中郵 2,947）", f"=PL!F{PR['B']['ni']}", "=AVERAGE(3160,3108,2947)", S_DW + "；" + S_GX + "；券商隱含第 4 季利潤回升，模型第 4 季業外假設打平")
    rrow("FY2027 歸屬母公司淨利（情境 B）vs 券商 2027E 平均（4,280、3,855、4,202）", f"=PL!G{PR['B']['ni']}", "=AVERAGE(4280,3855,4202)", "券商預測為 Interested-party")
    rrow("FY2028 歸屬母公司淨利（情境 B）vs 券商 2028E 平均（5,210、4,689、5,174）", f"=PL!H{PR['B']['ni']}", "=AVERAGE(5210,4689,5174)", "")
    rrow("FY2027 EPS（情境 B）vs 券商 2027E 平均（淨利平均 ÷ 股數）", f"=PL!G{PR['B']['eps']}", "=AVERAGE(4280,3855,4202)/sh_v", "", "0.00")
    rrow("FY2026 毛利率 vs 2026 年 1–9 月實績", f"=PL!F{PR['B']['gm']}", "=gm26", "模型以 1–9 月為錨（差距應為 0）", "0.0%")
    rrow("FY2026 第 4 季隱含營收 vs 2026Q3 實績", f"=PL!F{PR['B']['rev']}-r9m26", "=Q!D4", "第 4 季若維持第 3 季水準即為對照值")
    rrow("加權目標價 vs 市場共識目標價（資料商 23 家，頁面 2026-11-24）", f"=Val!B{VR['tp']}", "=cons_tp", S_SA_ST + "；只作對照，不進模型")
    rr += 1
    RC.cell(rr, 1, "線 6 市場對照（人形用行星滾柱絲杠；百萬美元換人民幣；本公司人形絲杠營收占全球市場）").font = BOLD; rr += 1
    for y_, col_ in ((2026, "F"), (2030, "J"), (2032, "L")):
        rrow(f"FY{y_} 人形絲杠營收（情境 B）÷ 線 6 絲杠市場 {y_}（基準單價口徑）", f"=PL!{col_}{PR['B']['hum']}/fx/線取數!{TC[y_]}{MR[y_]}", None,
             f"線 6 市場規模 {y_} 基準 {LD['行星滾柱絲杠市場規模 %d' % y_][0]:,.0f} 百萬美元（{LD['行星滾柱絲杠市場規模 %d' % y_][1]}）；線 4 對應市場較小（見對帳）", "0.00%", False)
    rrow("FY2026 人形絲杠隱含全球份額（以線 4 基準校準）", f"=Thesis!E{THR['B']['hsh']}", None, "＝hum26 ÷（線 4 絲杠需求 × 線 4 隱含單價 × 匯率）", "0.00%", False)
    rrow("FY2030 人形絲杠營收（情境 B）占合併營收", f"=Thesis!I{THR['B']['humsh']}", None, "即使份額三倍於 2026 年，人形絲杠對合併營收貢獻仍小", "0.00%", False)
    rr += 1
    RC.cell(rr, 1, "反解（整本重算二分搜尋；建置時計算，數值為快照）").font = BOLD; rr += 1
    hdr(RC, rr, ["反解目標", "解", "基準值", "", "說明"]); rr += 1
    if sens:
        for lab, v, b0, note in sens["solve"]:
            RC.cell(rr, 1, lab); RC.cell(rr, 2, v).number_format = "#,##0.00"; RC.cell(rr, 3, b0).number_format = "#,##0.00"; RC.cell(rr, 5, note); rr += 1
    RC.column_dimensions["A"].width = 70; RC.column_dimensions["E"].width = 90

    # ================= 對帳 =================
    DZ["A1"] = "對帳：本模型使用的「線層級」變數（與 physical_ai/lines/ 各線模型對帳；本模型讀線 v0.1 成品，不靜默改口徑）。恒立液壓主體為工程機械液壓件，線模型沒有挖掘機或工程機械變數，該部分為本模型自設（見 In 頁）。"
    hdr(DZ, 3, ["變數名", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型是否一致"])
    s26 = LD["scr"][(2026, "B")]; s32 = LD["scr"][(2032, "B")]; p26 = LD["sp"][(2026, "B")]; p30 = LD["sp"][(2030, "B")]
    m6 = {y: LD['行星滾柱絲杠市場規模 %d' % y] for y in (2026, 2030, 2032)}
    ZR = [
        ("全球人形滾柱絲杠需求 FY2026–FY2032（A／B／C＝線高／基準／低）", "4 人形", f"基準 {s26[0]:,.0f} → {s32[0]:,.0f}（× hu_k）", "支", "全球人形，平均每台 4.2 支（全尺寸 14 支 × 採用比例）", f"{s26[1]}；{s32[1]}（逐年見線取數頁）", "Assumed", "Thesis 人形絲杠市場與營收", "一致（直接讀值，hu_k＝1）；線 6 2030 年絲杠需求 400 萬根（基準）高於線 4 的 210 萬支，兩線不一致，本模型採線 4（有逐年路徑）"),
        ("人形用滾柱絲杠隱含單價 2026／2030（美元／支）", "4 人形", f"{p26[0]:,.0f}／{p30[0]:,.0f}", "美元／支", "線 4 市場 ÷ 需求（含學習曲線）", f"{p26[1]}", "Derived", "Thesis 單價路徑", f"一致（派生）；線 4 P_ROLLER_SCREW 基準 {LD['p_screw4'][0]:,.0f} 美元（{LD['p_screw4'][1]}）"),
        ("行星滾柱絲杠市場規模 2026／2030／2032", "6 共用零組件", f"{m6[2026][0]:,.0f}／{m6[2030][0]:,.0f}／{m6[2032][0]:,.0f}", "百萬美元", "覓途占比推算單價口徑", f"{m6[2026][1]}", "Derived", "Recon 隱含份額（對照）", "與線 4 不一致：線 4 的 2030 年人形絲杠市場約 747 百萬美元，線 6 基準 1,559 百萬美元（單價口徑不同），本模型營收讀線 4、對照列線 6"),
        ("本公司於行星滾柱絲杠的份額", "6 共用零組件", "未建立", "%", "—", f"{L6N} 份額頁（恒立液壓未列份額；五洲新春 2025 出貨約 8.7 萬根）", "Interested-party", "hum26、hs30_*（Assumed）", "線 6 未提供本公司份額，本模型自設並給區間"),
        ("挖掘機銷量、油缸與泵閥單價", "（無）", "In 頁 gcyl_*、gpump_*", "%／年", "本模型自設", "線模型無工程機械變數；參考券商摘要：2026 年 1–6 月挖掘機銷量 +26.4%（出口 +33.5%）、Q3 +15.8%", "Assumed", "Thesis 油缸、泵閥營收路徑", "不適用（線模型無對應變數）；為本模型主要驅動，列為證據邊界"),
    ]
    for i, z in enumerate(ZR):
        for j, v in enumerate(z): DZ.cell(4 + i, 1 + j, v)
    DZ.column_dimensions["A"].width = 50; DZ.column_dimensions["F"].width = 80; DZ.column_dimensions["I"].width = 70

    # ================= Risk =================
    RK["A1"] = "Risk：風險登錄（原則第 10 節第 10、19 條）。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    RISKS = [
        ("工程機械週期反轉（挖掘機銷量）", "是（情境 C：gcyl_C、gpump_C）", "油缸與泵閥合計約 88% 營收", "挖掘機月銷量與出口、開工小時"),
        ("匯率與外匯掉期損益", "部分（nonop_q4、inv_y）", "2026 年前三季財務費用率 1.78%，利潤受匯兌拖累", "季度財務費用、公允價值變動"),
        ("海外產能爬坡（墨西哥、印尼）", "部分（gm30、cx_r）", "Q3 毛利率年減 1.98 個百分點", "季毛利率、在建工程"),
        ("線性驅動與人形絲杠進度", "是（hum26、hs30_*）", "營收占比低，對估值影響小", "募投 2026-12 達可使用狀態；客戶量產訂單"),
        ("短期借款上升（2026 年 9 月 849 百萬元，2025 年底 14）", "否", "用途未查；占現金 8%", "年報借款附註"),
        ("國內同業倍數未納入", "否", "三一、徐工、中聯、艾迪精密不在收盤檔，本模型同業為日本與傳動公司", "同日報價取得後重算 pe_peer"),
        ("Q4 利潤回升預期（券商隱含）", "部分", "模型第 4 季業外打平，低於券商隱含", "2026 年報"),
    ]
    for i, k in enumerate(RISKS):
        for j, v in enumerate(k): RK.cell(4 + i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["D"].width = 60

    # ================= Sources =================
    SO["A1"] = "Sources：本模型引用的原始來源與資料商（含網址；HTML 展示層不放網址）。"
    for i, s_ in enumerate([S_PX, S_SA, S_SA_ST, S_Q3, S_ZY, S_Y25, S_H1, S_DW, S_GX, S_RF, S_CRP, S_SAP,
                            f"線 4：{L4N}", f"線 6：{L6N}",
                            "physical_ai/nabtesco、smc、thk、hiwin、fanuc、harmonic_drive、leaderdrive 模型（同業前瞻 P/E 與 β）",
                            "physical_ai/map/profiles/601100.md（背景，數字已另行查證）"], 3):
        SO.cell(i, 1, s_)
    SO.column_dimensions["A"].width = 160

    # ================= README（xlsx 內）=================
    S = T
    v = lambda k: f"Val!B{VR[k]}"
    lines = [
        ('="恒立液壓 Hengli Hydraulic（601100.SS）個股模型 ' + VER + '｜build_id "&Sens!D2&"｜資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'),
        "命題：恒立液壓的營收約九成來自挖掘機等工程機械的油缸與泵閥，2026 年前三季營收年增 31% 但淨利只增 3%（匯兌與新廠成本）；人形機器人滾柱絲杠（線 4 需求 × 線 4 單價 × 份額）目前占比不到 1%，不是估值主軸。",
        (f'="結論（條件式）：在 FY2026 以 1–9 月實績為錨、油缸首年增速 "&{S("gcyl_B","0%")}&"、泵閥 "&{S("gpump_B","0%")}&"（逐年遞減 ×"&{S("fade","0.00")}&"）、FY2030 毛利率 "&{S("gm30_B","0.0%")}&"、同業 P/E 未加溢價（相對倍數係數 "&{S("rel_k","0.00")}&"）的假設下，加權目標價約 "&{S(v("tp"),"#,##0.0")}&" 元，相對現價 "&{S("px","#,##0.00")}&" 元 "&{S(v("up"),"+0.0%;-0.0%")}&"，估值評等 "&{v("rating")}&"，投資建議 "&{v("rec")}&"。"'),
        (f'="DCF 目標價 "&{S(v("dcf_tp"),"#,##0.0")}&"、Comparable 目標價 "&{S(v("cmp_tp"),"#,##0.0")}&"（各半）；WACC "&{S("wacc","0.00%")}&"、ke "&{S("ke","0.00%")}&"、g "&{S("g_tv","0.00%")}&"；情境 B 終值占 EV "&{S(f"Val!C{VR['tvs']}","0%")}&"。"'),
        (f'="上檔空間歸因：DCF 貢獻 "&{S(v("att_d"),"+0.0%;-0.0%")}&"、未調整 Comparable 貢獻 "&{S(v("att_c"),"+0.0%;-0.0%")}&"、相對倍數調整貢獻 "&{S(v("att_r"),"+0.0%;-0.0%")}&"。本公司現價對資料商前瞻 EPS 的倍數是同業中位數的 "&{S(v("relnow"),"0.0")}&" 倍；支持現價需要相對倍數係數 "&{S(v("relreq"),"0.0")}&"，升一級需要 "&{S(v("relflip"),"0.0")}&"。"'),
        (f'="收支關卡：情境 B FY2026–FY2030 FCFF 扣股利合計 "&{S(v("gate_b"),"#,##0.0")}&" 百萬元，"&{v("gate_t")}&"；距最近級距邊界 "&{S(v("dist"),"0.0%")}&"。"'),
        "主要失效條件（可觀察）：(1) 挖掘機銷量與出口持續高於情境 A；(2) 國內工程機械同業（本模型未取得同日報價）倍數遠高於日本傳動同業，使同業中位數低估；(3) 毛利率回到 41% 以上；(4) 人形絲杠取得實質訂單。",
        "方法：DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，WACC 折現，明確期到 FY2032，RONIC＝WACC）50%＋同業前瞻 P/E × 模型 FY2027 EPS 50%；三情境 25／50／25。",
        "限制：分產品占比（FY2026）與分部增速為 Assumed；稅前利益與營業利益為資料商口徑；同業為日本與傳動公司（國內同業未取得同日報價）；不讀 Siliconomics／Tokenomics／Wattonomics。詳見 repo README.md。",
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
        ("情境 B 的 2030 年人形絲杠份額改為情境 A 值", "hs30_B", ivE["hs30_A"]),
        ("情境 B 的 2030 年人形絲杠份額改為情境 C 值", "hs30_B", ivE["hs30_C"]),
        ("情境 B 油缸首年增速改為情境 A 值", "gcyl_B", ivE["gcyl_A"]),
        ("情境 B 油缸首年增速改為情境 C 值", "gcyl_B", ivE["gcyl_C"]),
        ("情境 B 泵閥及馬達首年增速改為情境 A 值", "gpump_B", ivE["gpump_A"]),
        ("情境 B 泵閥及馬達首年增速改為情境 C 值", "gpump_B", ivE["gpump_C"]),
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

    # 反解（整本重算，對數二分）
    px = ivE["px"]
    def solve(key, target_tp, lo, hi, n=12, over=None):
        over = over or {}
        a, b = math.log(lo), math.log(hi)
        if run_cases([{**over, key: hi}], meta, "solve")[0]["tp"] < target_tp: return None
        if run_cases([{**over, key: lo}], meta, "solve")[0]["tp"] >= target_tp: return lo
        for _ in range(n):
            m = (a + b) / 2
            if run_cases([{**over, key: math.exp(m)}], meta, "solve")[0]["tp"] < target_tp: a = m
            else: b = m
        return math.exp((a + b) / 2)
    up = base["up"]
    flip = -0.5 if up < -0.5 else -0.25 if up < -0.25 else 0.25 if up <= 0.25 else 0.5
    flip_tp = px * (1 + flip + 1e-6)
    s1 = solve("pe_peer", px, 5.0, 120.0)
    s2 = solve("pe_peer", flip_tp, 5.0, 120.0)
    s3 = solve("hu_k", px, 1.0, 200.0)
    s4 = solve("gm30_B", px, 0.2, 0.9) if False else None
    nf = lambda x: x if x else -1
    pe0 = ivE["pe_peer"]
    solve_rows = [
        ("支持現價（目標價＝現價）所需的同業前瞻 P/E 中位數（倍）", nf(s1), pe0, "其他輸入不變；−1＝120 倍內無解"),
        (f"評等升一級（上檔 {flip:+.0%}）所需的同業前瞻 P/E 中位數（倍）", nf(s2), pe0, "其他輸入不變；−1＝120 倍內無解"),
        ("支持現價所需的人形出貨相對線 4 倍數", nf(s3), 1.0, "其他輸入不變；−1＝200 倍內無解（人形絲杠占營收不到 1%）"),
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
    # segments.json
    wbv = openpyxl.load_workbook(OUT, data_only=True)
    Hs = wbv["Hist"]
    segs = [("液壓油缸", "液壓油缸 營收", Hs["E8"].value), ("液壓泵閥及馬達", "液壓泵閥及馬達 營收", Hs["E9"].value),
            ("液壓系統", "液壓系統 營收", Hs["E10"].value), ("配件、鑄件及其他業務（殘差）", "配件、鑄件及其他業務 營收", Hs["E11"].value)]
    json.dump(dict(company="hengli", ticker="601100.SS", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=round(v, 2), unit="百萬元", source=S_Y25) for n, l, v in segs],
                   sublines=[dict(parent="配件、鑄件及其他業務", labels=["配件鑄件及其他（不含人形絲杠）（子列）", "人形機器人用滾柱絲杠（子列）"])],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）", "檢查：子列合計 − 配件鑄件及其他分部（應為 0）"],
                   note="官方分部＝年報主營業務分產品（液壓油缸、液壓泵閥及馬達、液壓系統、配件及鑄件）；配件及鑄件與其他業務收入合併為一列（殘差）。人形滾柱絲杠子列為命題拆分（Assumed），公司未揭露；合計＝配件、鑄件及其他業務分部。"),
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
    t = f"""# physical_ai/hengli/ — 恒立液壓 Hengli Hydraulic（601100.SS）個股模型 {VER}

<!-- 本檔由 build_hengli.py 依成品重算值產生；數字不手寫（原則第 10 節第 51 條）。 -->

**本模型適用《高不確定性預估工作原則》**（`docs/plan/高不確定性預估工作原則.md`）：公司未揭露 2026 年分產品營收與人形絲杠營收，相關輸入一律給區間並標 Assumed／Analogy；結論只在證據邊界內成立。

- 成品：`dist/{NAME}`（build_id `{sens['build_id']}`；LibreOffice 重算，0 錯誤值）；HTML 展示層 `dist/{NAME[:-5]}.html`（`gen_html_json.py` 產生 `html.json`，再以 `tools/company_html/build_company_html.py` 產生）
- 建置：`python3 physical_ai/hengli/build_hengli.py`（需要 soffice；整本重算敏感度與反解；結尾自動跑 in_layout、sens_coverage --require-keycol、model_governance、rating_gate --strict、zero_rows、cached_values、segment_check）
- 日期：資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 2026-11-20（收盤 {I['px']:,.2f} 元）｜評價日 2026-12-31｜目標時點 2027-11-20
- 會計年度＝日曆年；單位百萬元人民幣；中國會計準則（資料商彙整口徑）。上海證券交易所主板。

## 結論（條件式）

在「FY2026 以 1–9 月實績為錨、油缸首年增速 {I['gcyl_B']:.0%}／泵閥及馬達 {I['gpump_B']:.0%}（逐年遞減 ×{I['fade']:.2f}）、FY2030 毛利率 {I['gm30_B']:.1%}、同業 P/E 不加溢價（相對倍數係數 {I['rel_k']:.2f}）」的假設下：

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
| 市場共識目標價（資料商 23 家，頁面 2026-11-24；只作對照） | {I['cons_tp']:,.2f} |

上檔空間歸因：DCF 貢獻 {g('att_d'):+.1%}、未調整 Comparable 貢獻 {g('att_c'):+.1%}、相對倍數調整 {g('att_r'):+.1%}。本公司現價對資料商前瞻 EPS 的倍數是核心同業中位數（{I['pe_peer']:.1f} 倍）的 {g('relnow'):.1f} 倍；要讓目標價＝現價，Comparable 的相對倍數係數須達 {g('relreq'):.1f}。DCF 以 RONIC＝WACC（新投資無超額報酬），而高相對倍數隱含持續超額報酬，兩法對競爭優勢的隱含假設相反。

**讀法：** 恒立液壓是工程機械液壓件公司，人形機器人用滾柱絲杠營收情境 B 為 {ser('hum', '{:,.0f}')}（FY2026→FY2032，百萬元），占合併營收只從 {TH[f'E{THB["humsh"]}'].value:.2%} 到 {TH[f'K{THB["humsh"]}'].value:.2%}，因此不是估值主軸。同業中位數 {I['pe_peer']:.1f} 倍來自日本與傳動公司，**國內工程機械同業（三一、徐工、中聯、艾迪精密）不在收盤檔、未進中位數**，若國內同業倍數明顯較高，Comparable 腿會偏低。整本重算反解：其他輸入不變時，**支持現價需要同業前瞻 P/E 中位數 {fn(sol[0][1], '{:,.1f}')} 倍**（基準 {I['pe_peer']:.1f}）；評等升一級需要 {fn(sol[1][1], '{:,.1f}')} 倍。結論主要建立在 Assumed（分產品占比與增速、毛利率路徑）與 Derived（同業倍數）輸入上。

## 命題

1. **營收結構：** 官方產品分部為液壓油缸（2025 年 {wb['Hist']['E8'].value:,.0f} 百萬元，占 {wb['Hist']['E8'].value / wb['Hist']['E4'].value:.0%}）、液壓泵閥及馬達（{wb['Hist']['E9'].value:,.0f}）、液壓系統（{wb['Hist']['E10'].value:,.0f}）、配件鑄件及其他（{wb['Hist']['E11'].value:,.0f}）。2026 年分產品數字未取得，以占比 Assumed 起算。
2. **量與價：** 油缸與泵閥以逐年遞減的年增速外推（挖掘機銷量 2026 年上半年 +26.4%、Q3 +15.8%）；線模型沒有工程機械變數，此為本模型自設。人形絲杠＝線 4 絲杠需求 × 線 4 隱含單價 × 份額。
3. **利潤：** 毛利率 2025 年 {wb['Hist']['E7'].value:.1%}、2026 年 1–9 月 {I['gm26']:.1%}；營業費用率（含稅金附加與研發）{I['opex26']:.1%} 走向 {I['opex30']:.0%}；2026 年前三季業外為負（匯兌與掉期）。

情境 B：營收 {ser('rev', '{:,.0f}')}；毛利率 {ser('gm', '{:.1%}')}；EPS {ser('eps', '{:.2f}')} 元（FY2026→FY2032）。

## 主要假設與可證偽指標（原則第 10 節第 20 條）

| 承重假設 | 基準（A／B／C 或區間） | 標記 | 可觀察指標 |
|---|---|---|---|
| 油缸首年增速（FY2027） | {I['gcyl_A']:.0%}／{I['gcyl_B']:.0%}／{I['gcyl_C']:.0%} | Assumed | 挖掘機月銷量與出口 |
| 泵閥及馬達首年增速 | {I['gpump_A']:.0%}／{I['gpump_B']:.0%}／{I['gpump_C']:.0%} | Assumed | 半年報分產品表 |
| 增速遞減係數 | {I['fade']:.2f}（0.70–0.95） | Assumed | 券商 2027／2028 年營收預測 |
| FY2030 毛利率 | {I['gm30_A']:.1%}／{I['gm30_B']:.1%}／{I['gm30_C']:.1%} | Assumed | 季毛利率（2026 第 3 季 40.2%） |
| 2026 年油缸、泵閥占營收 | {I['sh_cyl26']:.0%}／{I['sh_pump26']:.0%} | Assumed | 年報分產品 |
| 2030 年人形絲杠份額 | {I['hs30_A']:.0%}／{I['hs30_B']:.0%}／{I['hs30_C']:.0%} | Assumed（線 4 需求為 Assumed） | 募投達可使用狀態、人形客戶訂單 |
| 同業前瞻 P/E | 中位數 {I['pe_peer']:.1f} 倍（6 家） | Derived | 同一估值日重抓，補入國內同業 |

## 敏感度（整本重算，前 10 大）

| 案例 | 鍵 | 加權目標價 | 相對基準 | DCF 目標價 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | `{c['key']}` | {c['tp']:,.1f} | {c['tp'] / sens['base_tp'] - 1:+.1%} | {c['dcf']:,.1f} | {c['rating']} |" for c in top) + f"""

完整案例見 xlsx Sens 頁（每個 In 頁有低／高區間的鍵都有案例；另有情境 B 改 A／C 值、人形出貨 5 倍、情境權重案例）。升一級需加權目標價達 {fn(up_b)} 元。

## 反解（Recon 頁）

""" + "\n".join(f"- {a}：{fn(b, '{:,.2f}')}（基準 {fn(c0, '{:,.2f}')}）；{d}" for a, b, c0, d in sol) + f"""

## 限制與「分不出來／沒找到」

- **分不出來（資料性質）：** 人形絲杠營收與線性驅動營收公司未揭露；2026 年分產品營收未取得。
- **沒找到（工作未完）：** 公司年報、三季報原文未直接讀取（歷史數字取自資料商與券商轉錄；年報 PDF 無法解析）；銷售、管理、研發費用細項（以「毛利 − 營業利益」合併為營業費用）；稅前利益為資料商推算；2023 年分產品營收；少數股東損益；短期借款用途；限制性股票與選擇權。
- **線模型不一致：** 線 6 2030 年絲杠需求 400 萬根高於線 4 的 210 萬支；線 6 市場規模與線 4 單價口徑不同（採線 4，對照列線 6）。見 xlsx「對帳」頁。
- **同業倍數：** 6 家核心同業（Nabtesco、SMC、THK、Fanuc、安川電機、雙環傳動）皆取 2026-11-20 收盤換算值；貝斯特、上銀舊報價日收盤不在收盤檔，無法換算；三一重工、徐工機械、中聯重科、艾迪精密不在收盤檔。以上列「報價日未取得」，未進中位數。少於 8 家不做迴歸。
- **rf 逾期：** 中國 10 年期國債 1.61% 為 2026-11-05 值（跨模型一致）。
- **協調檔：** 實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics；相關條文不適用。
- **獨立核對：** 依專案規則，交付判讀的數字應由未參與製作的代理獨立核對；待外層流程安排。

## 資料來源

見 xlsx Sources 頁（資料商財務報表、券商轉錄年報與季報、同業模型、線 4／6 成品）。

## 版本紀錄

- {VER}（2026-11-24）：首版。現價 {I['px']:,.2f}（2026-11-20 收盤）；加權目標價 {g('tp'):,.1f}；上檔空間 {g('up'):+.1%}；{g('rating')}／{g('rec')}；收支關卡 {g('gate_b'):,.1f}（{g('gate_t')}）。
"""
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write(t)


if __name__ == "__main__":
    main()

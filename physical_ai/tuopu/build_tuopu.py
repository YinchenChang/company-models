# 拓普集團 Tuopu（601689.SS，上海主板）個股模型 v0.1 建置腳本
# 產出：physical_ai/tuopu/dist/20261124_拓普集團個股模型_v0.1.xlsx（LibreOffice 重算後的成品，含公式與快取值）
# 結構：命題模組（Thesis：營收＝汽車零部件（本業，依年增情境）＋人形致動器（線 4 出貨 × 線 4 每台致動器內容 × 組合係數 × 份額））
#       → 標準損益表 → FCFF DCF＋同業 P/E。公司未揭露機器人營收，機器人用營收為 Assumed。
# 敏感度：每個案例覆寫 In 頁基準欄後整本重建、LibreOffice 重算；反解同法以二分搜尋。
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md、docs/plan/高不確定性預估工作原則.md。
# 用法：python3 physical_ai/tuopu/build_tuopu.py   （需要 soffice；暫存於 $HOME/tp_build）
import csv
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
NAME = f"20261124_拓普集團個股模型_{VER}.xlsx"
OUT = os.path.join(HERE, "dist", NAME)
HOMEDIR = os.path.expanduser("~")
TMP = os.path.join(HOMEDIR, "tp_build")
LO = ["soffice", f"-env:UserInstallation=file://{HOMEDIR}/lo_profile", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML"]

L4 = os.path.join(ROOT, "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx")
L6 = os.path.join(ROOT, "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx")
L4N, L6N = (os.path.relpath(p, ROOT) for p in (L4, L6))
CLOSES = os.path.join(ROOT, "physical_ai/data/closes2_20261109_20261120.csv")


def close(sym, day="2026-11-20"):
    with open(CLOSES, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if r["symbol"] == sym and r["date"] == day: return round(float(r["close"]), 2)
    return None


BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDRF = PatternFill("solid", fgColor="DDEBF7")
SECF = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")

S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv：601689.SS；未還原股利）"
S_SA = "stockanalysis.com SHA:601689（資料商；年度與季度損益表、資產負債表季頁、現金流量表、statistics 頁，2026-11-24 讀取；公司財報原文未直接讀取，利潤口徑為資料商標準化）"
S_SA_ST = "stockanalysis.com SHA:601689 statistics（資料商；2026-11-24 讀取）"
S_RF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與 leaderdrive 等模型同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：China CRP 0.91%（與 leaderdrive 同一取值）"
S_CLOSE = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
S_SAP = "stockanalysis.com statistics 頁（資料商；2026-11-24 讀取；原報價日見 Comp 頁 M 欄）"
S_PROF = "physical_ai/map/profiles/601689.md（背景；其中券商點評與新聞標記為外部、未經本專案檢核）"
IP_SA = "資料商彙整賣方共識，報價日與家數未完整標示"


def _line_rows(path, idcol, sheet="下游取數"):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    out = {}
    for r in range(1, ws.max_row + 1):
        k = ws.cell(r, idcol).value
        if isinstance(k, str): out[k] = (r, [ws.cell(r, c).value for c in range(1, 10)])
    return out


def line_data():
    l4 = _line_rows(L4, 1); l6 = _line_rows(L6, 1)
    d = {"ship": {}, "bom": {}}
    for y in range(2026, 2033):
        r, v = l4[f"SHIP_{y}"]
        for s_, c, col in (("A", 4, "E"), ("B", 3, "D"), ("C", 2, "C")):
            d["ship"][(y, s_)] = (v[c], f"{L4N} 下游取數!{col}{r}（SHIP_{y}）")
    # 每台全尺寸 BOM（基準欄 D）：線 4 給 2026、2028、2030、2035，其餘年份幾何內插（Derived）
    pts = {}
    for y in (2026, 2028, 2030, 2035):
        r, v = l4[f"HUM_BOM_{y}"]; pts[y] = (v[3], r)
    ys = sorted(pts)
    for y in range(2026, 2033):
        lo = max(a for a in ys if a <= y); hi = min(a for a in ys if a >= y)
        if lo == hi: val = pts[lo][0]; src = f"{L4N} 下游取數!D{pts[lo][1]}（HUM_BOM_{lo}）"
        else:
            val = pts[lo][0] * (pts[hi][0] / pts[lo][0]) ** ((y - lo) / (hi - lo))
            src = f"Derived：{L4N} 下游取數 HUM_BOM_{lo}（D{pts[lo][1]}）與 HUM_BOM_{hi}（D{pts[hi][1]}）幾何內插"
        d["bom"][y] = (val, src)
    r, v = l4["HUM_ACT_SHARE"]; d["act"] = [(v[2], f"{L4N} 下游取數!C{r}"), (v[3], f"{L4N} 下游取數!D{r}（HUM_ACT_SHARE）"), (v[4], f"{L4N} 下游取數!E{r}")]
    r, v = l4["Q_HARMONIC_AVG"]; r2, v2 = l4["Q_ROTARY"]
    d["mix"] = (v[3] / v2[3], f"{L4N} 下游取數!D{r}（Q_HARMONIC_AVG）÷ D{r2}（Q_ROTARY）＝出貨組合平均用量 ÷ 全尺寸用量（Derived）", v[2] / v2[3], v[4] / v2[3])
    r, v = l4["P_ROLLER_SCREW"]; d["screw"] = (v[3], f"{L4N} 下游取數!D{r}（P_ROLLER_SCREW）")
    r, v = l4["Q_ROLLER_SCREW"]; d["qscrew"] = (v[3], f"{L4N} 下游取數!D{r}（Q_ROLLER_SCREW）")
    k6 = [k for k in l6 if k.startswith("人形每台零組件內容 2026")][0]
    r, v = l6[k6]; d["c6"] = (v[2], f"{L6N} 下游取數!C{r}（{k6}）")
    k6b = [k for k in l6 if k.startswith("人形每台零組件內容 2030")][0]
    r, v = l6[k6b]; d["c6_30"] = (v[2], f"{L6N} 下游取數!C{r}（{k6b}）")
    return d


LD = line_data()

# ---------------- In 頁輸入 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 來源誘因, 說明)
IN = [
    ("sec", "1. 市場資料與日期"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, "", "資料截止日 2026-11-24（財報、同業倍數）；股價 2026-11-20；成品日 2026-11-24"),
    ("px", "現價（2026-11-20 收盤）", "元", None, close("601689.SS"), None, "Verified", S_PX, "", "資料商頁面 2026-11-24 價格較高（見 Comp 頁），本模型一律用 11-20 收盤"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,12,31)", None, "Assumed", "本專案慣例：評價日 2026-12-31（會計年度＝日曆年）", "", "折現期間由評價日起算"),
    ("tp_date", "目標時點（px_date 後 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", ""),
    ("bs_date", "資產負債表日", "日期", None, "=DATE(2026,9,30)", None, "Verified", S_SA, "", "2026 年三季報"),
    ("stub_frac", "Stub 期（資產負債表日→評價日）", "年", None, "=(val_date-bs_date)/365", None, "Derived", "＝(評價日 − 資產負債表日) ÷ 365", "", "Stub FCFF＝FY2026 FCFF × stub_frac"),
    ("sh_v", "評價日股數（總股本）", "百萬股", None, 1738.0, None, "Verified", S_SA + "：資產負債表頁股數 1,738 百萬股", "", "限制性股票與可轉債未取得揭露，以 0 計"),
    ("cash_v", "現金＋短期投資（2026-09-30）", "百萬元", None, 5057.0, None, "Verified", S_SA, "", "2025-12-31 為 5,620"),
    ("ltinv_v", "長期投資（本輪未取得）", "百萬元", None, 0.0, None, "Assumed", "資料商淨現金口徑未列長期投資；以 0 計（保守）", "", "拓普另有對外投資，規模未取得"),
    ("inv_k", "長期投資計入淨現金的比例", "%", 0.5, 1.0, None, "Assumed", "長期投資為 0，本鍵無作用（保留以符合範本結構）", "", ""),
    ("debt_v", "有息負債（2026-09-30）", "百萬元", None, 5782.0, None, "Verified", S_SA, "", "2025-12-31 為 5,201；淨負債 725"),
    ("nci_v", "少數股東權益（2026-09-30）", "百萬元", None, 39.76, None, "Verified", S_SA, "", "自淨現金扣除"),
    ("fx", "人民幣兌美元（2026-11-20，元／美元）", "元／美元", None, close("CNY=X"), None, "Verified", S_CLOSE + "：CNY=X", "", "用於人形致動器市場由美元換算人民幣，以及 Recon 對照"),
    ("cons_tp", "資料商平均目標價（18 家，報價日未標）", "元", None, 56.08, None, "Interested-party", S_SA_ST, IP_SA, "只作對照，不進模型"),
    ("sec", "2. 實績錨（FY2026＝2026 年，會計年度＝日曆年）"),
    ("r9m26", "2026 年 1–9 月營收", "百萬元", None, "=Q!F4", None, "Derived", "＝Q 頁 Q1＋Q2＋Q3（資料商季度營收）", "", ""),
    ("g9m26", "2026 年 1–9 月營收年增", "%", None, "=Q!F4/(Hist!E4-Q!B9)-1", None, "Derived", "＝1–9 月營收 ÷（FY2025 營收 − 2025Q4 營收）− 1", "", ""),
    ("q4g", "2026 年第 4 季營收年增", "%", 0.0, 0.06, 0.12, "Assumed", "基準＝介於 1–9 月年增（約 8%）與第 3 季年增（約 5%）之間；季度成長趨緩", "", "FY2026 營收＝1–9 月實績＋2025Q4 ×（1＋q4g）"),
    ("gm26", "FY2026 毛利率", "%", None, "=Q!F5/Q!F4", None, "Derived", "＝1–9 月毛利 ÷ 營收（Q 頁）", "", "第 4 季假設同 1–9 月（FY2025 為 19.2%）"),
    ("opex26", "FY2026 營業費用及其他營業損益淨額 ÷ 營收", "%", 0.075, "=(Q!F5-Q!F7)/Q!F4", 0.095, "Derived", "＝（1–9 月毛利 − 營業利益）÷ 營收（Q 頁；資料商標準化口徑，含銷管研、稅金及附加、減值、其他收益）", "", "FY2023–FY2025 為 7.9%、8.5%、8.9%（Hist）"),
    ("sec", "3. 命題模組：營收＝汽車零部件＋人形致動器（線 4、線 6）"),
    ("hum_rev26", "FY2026 機器人（人形）致動器營收", "百萬元", 30, 100, 300, "Assumed", "公司未揭露機器人營收；中報後券商稱線性、旋轉執行器與靈巧手電機進入小批量交付（" + S_PROF + "）；區間為量級假設（占營收 0.1%–1%）", "", "反映小批量交付；隱含全球份額見 Thesis"),
    ("act_sh", "致動器占人形 BOM 比重（線 4，全尺寸）", "%", LD["act"][0][0], LD["act"][1][0], LD["act"][2][0], "Analogy", f"線 4：{LD['act'][1][1]}（低 {LD['act'][0][1]}、高 {LD['act'][2][1]}）", "", "致動器含旋轉與線性關節，以線 4 口徑為準"),
    ("mixf", "出貨組合係數（平均每台致動器價值 ÷ 全尺寸機種）", "x", LD["mix"][2], LD["mix"][0], LD["mix"][3], "Analogy", LD["mix"][1], "", "線 4 BOM 以全尺寸機種計；出貨含小型機，以諧波用量比（8 ÷ 16）折算；低／高＝線 4 用量低／高案"),
    ("hu_k", "人形出貨相對線 4 的倍數（全部情境）", "x", None, 1.0, 2.0, "Assumed", "基準＝線 4 原值；高案 2 倍：媒體稱中國 2026 全年人形產量可能逾 10 萬台，線 4 基準全球 5.5 萬台（外部、未經本專案檢核）", "", "反解變數"),
    ("gauto_A", "情境 A：汽車零部件營收年增（FY2027 起）", "%", None, 0.14, None, "Assumed", "FY2023–FY2025 營收年增 23%、35%、11%（Hist）；A＝回到近年中高段", "", "汽車零部件含減震、內飾、底盤、熱管理與電子；公司分部未取得，以合併營收扣機器人估算"),
    ("gauto_B", "情境 B：汽車零部件營收年增（FY2027 起）", "%", None, 0.08, None, "Assumed", "2026 年 1–9 月年增約 8%；乘用車需求與匯率壓力（" + S_PROF + "）", "", ""),
    ("gauto_C", "情境 C：汽車零部件營收年增（FY2027 起）", "%", None, 0.03, None, "Assumed", "客戶需求轉弱、價格壓力", "", ""),
    ("hs30_A", "情境 A：2030 年占全球人形致動器價值份額", "%", None, 0.08, None, "Assumed", "2026 隱含份額見 Thesis；媒體稱為 Optimus 致動器供應商，供應商名單未經財報驗證（Interested-party）", "", "2026 隱含份額線性走到 2030，其後持平"),
    ("hs30_B", "情境 B：2030 年占全球人形致動器價值份額", "%", None, 0.04, None, "Assumed", "同上；模組整合次產業不可替代性最低一級（本專案 segments.csv ACT）", "", ""),
    ("hs30_C", "情境 C：2030 年占全球人形致動器價值份額", "%", None, 0.01, None, "Assumed", "同上；整機廠自製關節或改用其他 Tier 1", "", ""),
    ("gm30_A", "情境 A：FY2030 毛利率", "%", None, 0.215, None, "Assumed", "FY2023 22.6%、FY2024 20.6%、FY2025 19.2%、2026 年 1–9 月 18.5%（Hist、Q）；A＝回升至 FY2024 水準", "", "FY2026 起線性走到 FY2030，其後持平"),
    ("gm30_B", "情境 B：FY2030 毛利率", "%", None, 0.195, None, "Assumed", "同上；B＝持平 FY2025 水準", "", ""),
    ("gm30_C", "情境 C：FY2030 毛利率", "%", None, 0.17, None, "Assumed", "同上；C＝續降（價格壓力、折舊上升）", "", ""),
    ("tp_wA", "情境權重 A", "", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", ""),
    ("tp_wB", "情境權重 B", "", None, 0.50, None, "Assumed", "同上", "", ""),
    ("tp_wC", "情境權重 C", "", None, 0.25, None, "Assumed", "同上", "", ""),
    ("sec", "4. 損益與現金流"),
    ("opex30", "FY2030 營業費用及其他營業損益淨額 ÷ 營收", "%", 0.075, 0.085, 0.095, "Assumed", "FY2023–FY2025 7.9%、8.5%、8.9%；2026 年 1–9 月 8.6%", "", "FY2026→FY2030 線性，其後持平"),
    ("nonop_r", "業外淨額 ÷ 營收", "%", -0.002, "=Hist!E26/Hist!E4", 0.007, "Derived", "＝FY2025（稅前利益 − 營業利益）÷ 營收（資料商標準化口徑，含利息與投資收益）", "", "淨負債增加，業外趨降"),
    ("tax_r", "有效稅率", "%", 0.10, "=Hist!E33/Hist!E32", 0.18, "Derived", "＝FY2025 所得稅 ÷ 稅前利益（資料商由利潤率推得，近似值）", "", "敏感度低／高 10%／18%（Assumed）"),
    ("nci_r", "少數股東損益 ÷ 稅後淨利", "%", None, "=Hist!E35/Hist!E34", None, "Derived", "資料商未列；少數股東權益僅 40 百萬元，以 0 計", "", ""),
    ("payout", "現金股利支付 ÷ 淨利", "%", None, "=Hist!E38/Hist!E36", None, "Derived", "＝FY2025 現金流量表股利支付 ÷ FY2025 淨利（" + S_SA + "）", "", "用於收支關卡與目標時點前股利"),
    ("da26", "FY2026 折舊攤銷 ÷ 營收", "%", None, "=Hist!E40/Hist!E4", None, "Derived", "＝FY2025 折舊攤銷 ÷ 營收", "", "之後線性收斂到資本支出比例 × da_ss（FY2032）"),
    ("da_ss", "穩態折舊 ÷ 資本支出", "x", None, 0.8, None, "Assumed", "穩態資本支出高於折舊（成長）", "", ""),
    ("cx26", "FY2026 資本支出", "百萬元", 3000, 3300, 3800, "Assumed", S_SA + "：近四季資本支出 3,232、FY2025 3,497", "", ""),
    ("cx_r", "資本支出 ÷ 營收（FY2027 起）", "%", 0.08, 0.10, 0.13, "Assumed", "FY2023 16%、FY2024 12%、FY2025 12%（Hist，Derived）；擴產高峰過後回落", "", "資本支出強度是 DCF 最敏感的輸入之一"),
    ("nwc_r", "營運資金增加 ÷ 營收增量", "%", 0.10, 0.20, 0.35, "Assumed", "FY2025 營業現金流 4,482 與淨利加折舊 4,733 相近，營運資金占用小；區間為判斷", "", ""),
    ("sec", "5. 折現率（DCF 標準規格第 2 節）"),
    ("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", S_RF, "", "日期 2026-11-05，較股價基準日早 11 個交易日（逾期；跨模型一致，下次全面改版更新）"),
    ("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", ""),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", ""),
    ("b_own", "本公司 5 年 β（原始）", "x", None, 0.89, None, "Verified", S_SA_ST, "", ""),
    ("b_peer", "同業 5 年 β 中位數（原始）", "x", None, "=MEDIAN(Comp!H5:H10)", None, "Derived", "＝Comp 頁核心同業 6 家 β 中位數", "", ""),
    ("beta", "β（兩來源 Blume 調整後中位數）", "x", "=beta-0.3", "=MEDIAN(0.67*b_own+0.33,0.67*b_peer+0.33)", "=beta+0.3", "Derived", "Blume 0.67β＋0.33；兩來源＝本公司、同業中位數", "", "敏感度 ∓0.3（跨模型慣例）"),
    ("kd", "稅前債務成本", "%", None, "=rf+0.01", None, "Assumed", "rf ＋ 1.0pt", "", ""),
    ("wacc_adj", "WACC 調整（敏感度用）", "pt", -0.01, 0, 0.01, "Assumed", "基準 0；敏感度 ±1pt", "", ""),
    ("g_cap", "永續成長率上限", "%", None, 0.03, None, "Assumed", "DCF 標準規格：3%，且不超過計價幣別 10 年期公債殖利率", "", ""),
    ("g_tv", "永續成長率", "%", 0.01, "=MIN(g_cap,rf)", 0.03, "Derived", "＝MIN(g_cap, rf)", "", "敏感度 1%／3%（3% 超過 rf，只作敏感度）"),
    ("ronic_sp", "RONIC − WACC", "pt", None, 0, 0.05, "Assumed", "RONIC 預設＝WACC（新投資價值中性）", "", ""),
    ("tv_alt", "終值改用 FCFF ×（1＋g）（0＝否，1＝是）", "開關", None, 0, 1, "Assumed", "DCF 標準規格第 4 節敏感度", "", ""),
    ("sec", "6. 同業倍數與權重"),
    ("pe_peer", "同業前瞻 P/E 中位數", "x", "=MIN(Comp!F5:F10)", "=MEDIAN(Comp!F5:F10)", "=MAX(Comp!F5:F10)", "Derived", "＝Comp 頁核心同業 6 家前瞻 P/E 中位數；股價皆為 2026-11-20 收盤、P/E 由原報價日等比例換算（EPS 分母不變，見 Comp）", "", "敏感度＝同業最低／最高"),
    ("rel_k", "相對倍數係數（本公司 ÷ 同業）", "x", None, 1.0, 1.5, "Assumed", "1.00（未調整：缺本公司對同業十年相對倍數歷史）", "", "高案 1.5 倍（未查核，不作依據）"),
    ("w_dcf", "DCF 權重（Comparable＝1 − 本值）", "%", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設 DCF 50%＋同業倍數 50%", "", ""),
]
SCN_DRIVERS = [("gauto", "汽車零部件營收年增"), ("hs30", "2030 年人形致動器份額"), ("gm30", "FY2030 毛利率")]

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
    IW["A2"] = "單位：百萬元人民幣，除另註；會計年度＝日曆年（FY2026＝2026 年）。損益口徑為資料商標準化數字（營業利益與稅前利益含近似處理），非公司財報逐項轉錄。"
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
    HI["A1"] = "Hist：歷史財務（百萬元人民幣；資料商標準化數字）。公式只引用本頁（定義名稱 h_*）。"
    hdr(HI, 3, ["項目", "", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"])
    HROWS = [  # (列, 鍵, 項目, [2023, 2024, 2025], 標記, 來源)
        (4, "h_rev", "營業收入", [19701, 26600, 29581], "Verified", S_SA + "：年度損益表"),
        (5, "h_cogs", "營業成本", [15255, 21125, 23902], "Derived", "＝營收 − 毛利（資料商毛利）"),
        (6, "h_gp", "毛利（＝營收 − 營業成本）", ["=C4-C5", "=D4-D5", "=E4-E5"], "Derived", "＝營業收入 − 營業成本；與資料商毛利 4,446／5,475／5,679 一致"),
        (7, "h_gm", "毛利率", ["=C6/C4", "=D6/D4", "=E6/E4"], "Derived", "＝毛利 ÷ 營收"),
        (8, "h_seg_a", "汽車零部件及其他（公司未揭露機器人營收，歷史全數計入）", ["=C4", "=D4", "=E4"], "Derived", "公司分產品營收本輪未取得；機器人營收未揭露，歷史視為 0"),
        (16, "h_opex", "營業費用及其他營業損益淨額（＝毛利 − 營業利益）", ["=C6-C21", "=D6-D21", "=E6-E21"], "Derived", "資料商標準化口徑：含銷管研、稅金及附加、減值、其他收益"),
        (21, "h_ebit", "營業利益（資料商）", [2406, 3227, 3048], "Verified", S_SA),
        (25, "h_ebitm", "營業利益率", ["=C21/C4", "=D21/D4", "=E21/E4"], "Derived", ""),
        (26, "h_nonop", "業外（＝稅前 − 營業利益）", ["=C32-C21", "=D32-D21", "=E32-E21"], "Derived", "利息、投資收益等合計"),
        (32, "h_pbt", "稅前利益（資料商由稅前利潤率 × 營收推得，近似）", [2463, 3421, 3150], "Derived", S_SA + "：稅前利潤率 × 營收，四捨五入誤差約 ±5"),
        (33, "h_tax", "所得稅（＝稅前 − 淨利，近似）", ["=C32-C34", "=D32-D34", "=E32-E34"], "Derived", "＝稅前利益 − 淨利"),
        (34, "h_nit", "稅後淨利", [2151, 3001, 2779], "Verified", S_SA),
        (35, "h_nci", "少數股東損益（未列，以 0 計）", [0, 0, 0], "Assumed", "資料商未列；少數股東權益僅 40 百萬元（2026-09-30）"),
        (36, "h_ni", "歸屬母公司淨利（＝稅後淨利）", ["=C34-C35", "=D34-D35", "=E34-E35"], "Derived", "＝稅後淨利 − 少數股東損益"),
        (37, "h_eps", "EPS（基本，元）", [1.35, 1.78, 1.61], "Verified", S_SA),
        (38, "h_dps", "現金股利支付（現金流量表，百萬元）", [648.39, 788.54, 1016], "Verified", S_SA + "：現金流量表"),
        (39, "h_sh", "股數（＝淨利 ÷ EPS，百萬股）", ["=C36/C37", "=D36/D37", "=E36/E37"], "Derived", ""),
        (40, "h_da", "折舊攤銷", [1184, 1605, 1954], "Verified", S_SA + "：現金流量表"),
        (41, "h_cx", "資本支出", [3177, 3146, 3497], "Verified", S_SA + "：現金流量表"),
        (42, "h_ocf", "營業現金流", [3366, 3236, 4482], "Verified", S_SA + "：現金流量表"),
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
    HI.cell(28, 1, "檢查：毛利 − 營業費用及其他 − 營業利益（應為 0）")
    for c in "CDE": HI[f"{c}28"] = f"={c}6-{c}16-{c}21"
    HI.cell(29, 1, "差額：營業利益＋業外 − 稅前（應為 0）")
    for c in "CDE": HI[f"{c}29"] = f"={c}21+{c}26-{c}32"
    HI.column_dimensions["A"].width = 58; HI.column_dimensions["G"].width = 90

    # ================= Q（季度；一列一個數字，單一標記）=================
    Q["A1"] = "Q：2026 年季度資料（百萬元）。一列一個項目、單一主要標記；資料商季度損益表（公司季報原文未直接讀取）。"
    hdr(Q, 3, ["項目", "2026Q1", "2026Q2", "2026H1", "2026Q3", "2026 年 1–9 月", "標記", "來源", "來源誘因", "說明"])
    QROWS = [
        (4, "營業收入", [6628, 7571, "=B4+C4", 8416, "=D4+E4"], "Verified", S_SA + "：季度損益表", "", "2025Q3 營收 7,994；Q3 年增約 5%"),
        (5, "毛利", [1276, 1302, "=B5+C5", 1616, "=D5+E5"], "Verified", "同上", "", ""),
        (6, "毛利率", ["=B5/B4", "=C5/C4", "=D5/D4", "=E5/E4", "=F5/F4"], "Derived", "＝毛利 ÷ 營收", "", ""),
        (7, "營業利益", [791.97, 406.09, "=B7+C7", 1051, "=D7+E7"], "Verified", "同上", "", "Q2 偏低，原因本輪未查"),
        (8, "歸屬母公司淨利", [551.8, 470.75, "=B8+C8", 639.51, "=D8+E8"], "Verified", "同上", "", "2025Q3 淨利 671.62，Q3 年減約 5%"),
        (9, "2025Q4 營收（算 FY2026 第 4 季年增基期）", [8653, None, None, None, None], "Verified", "同上", "", ""),
    ]
    for (rr, lab, vals, tag, src, inc, note) in QROWS:
        Q.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = Q.cell(rr, 2 + j, v)
                if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
                if isinstance(v, str) and v.startswith("=") and rr == 6: c.number_format = "0.0%"
        Q.cell(rr, 7, tag); Q.cell(rr, 8, src); Q.cell(rr, 9, inc); Q.cell(rr, 10, note)
    Q.column_dimensions["A"].width = 30; Q.column_dimensions["H"].width = 90

    # ================= 線取數（線模型下游取數；藍字＝讀入值，H 欄寫來源檔與儲存格）=================
    LN["A1"] = "線取數：本模型讀入的線 4／線 6 變數年度路徑（建置時由線模型成品依 ID 讀值）。情境對應：A＝線模型高情境、B＝基準、C＝低情境。"
    hdr(LN, 3, ["項目", "情境", ""] + [f"FY{y}" for y in range(2025, 2033)] + ["單位", "來源（檔案與儲存格）"])
    LR = {}
    r = 4
    for s_ in SC:
        LN.cell(r, 1, "全球人形機器人出貨（線 4）"); LN.cell(r, 2, f"{s_}（{LINE_SC[s_]}）")
        srcs = []
        for y in range(2026, 2033):
            v, src = LD["ship"][(y, s_)]
            c = LN.cell(r, 4 + (y - 2025), v); c.font = BLUE; c.number_format = "#,##0"
            srcs.append(src)
        LN.cell(r, 12, "台"); LN.cell(r, 13, f"{srcs[0]} … {srcs[-1]}")
        LR[("hum", s_)] = r; r += 1
    LN.cell(r, 1, "全尺寸人形每台 BOM（線 4，基準；三情境共用）"); LN.cell(r, 2, "B（基準）")
    srcs = []
    for y in range(2026, 2033):
        v, src = LD["bom"][y]
        c = LN.cell(r, 4 + (y - 2025), v); c.font = BLUE; c.number_format = "#,##0"
        srcs.append(src)
    LN.cell(r, 12, "美元／台"); LN.cell(r, 13, f"{srcs[0]} … {srcs[-1]}（2027、2029、2031、2032 為幾何內插）")
    LR["bom"] = r
    LN.column_dimensions["A"].width = 52; LN.column_dimensions["M"].width = 110

    # ================= Thesis（命題模組）=================
    TH["A1"] = ("命題模組：營收＝汽車零部件（FY2026 錨，其後依情境年增）＋人形致動器（線 4 出貨 × 每台致動器內容 × 組合係數 × 匯率 × 份額）。"
                "FY2026＝1–9 月實績＋第 4 季假設，其中機器人營收為 Assumed。百萬元。")
    TH["A2"] = "欄：D＝FY2025 … K＝FY2032。每台致動器內容＝線 4 全尺寸 BOM × 致動器占比 × 出貨組合係數。人形 2026 隱含份額以線 4 基準路徑校準（三情境同一實績年）。"
    hdr(TH, 3, ["項目", "說明", ""] + [f"FY{y}" for y in range(2025, 2033)])
    THR = {}
    r = 4
    for s in SC:
        TH.cell(r, 1, f"情境 {s}").font = BOLD; TH.cell(r, 1).fill = SECF; r += 1
        R = {}
        R["rev26"] = r; TH.cell(r, 1, "FY2026 合併營收（錨）"); TH.cell(r, 2, "＝1–9 月＋2025Q4 ×（1＋q4g）")
        TH[f"{TC[2026]}{r}"] = "=r9m26+(Hist!E4-r9m26/(1+g9m26))*(1+q4g)"; r += 1
        R["ship"] = r; TH.cell(r, 1, "全球人形出貨（線 4 × 出貨倍數，千台）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR[('hum', s)]}*hu_k/1000"
        r += 1
        R["cont"] = r; TH.cell(r, 1, "每台致動器內容（美元）"); TH.cell(r, 2, "＝線 4 BOM × act_sh × mixf")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR['bom']}*act_sh*mixf"
        r += 1
        R["pool"] = r; TH.cell(r, 1, "全球人形致動器市場（百萬元人民幣）"); TH.cell(r, 2, "＝出貨（千台）× 每台內容 ÷ 1000 × 匯率")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['ship']}*{TC[y]}{R['cont']}/1000*fx"
        r += 1
        R["hsh"] = r; TH.cell(r, 1, "人形致動器份額"); TH.cell(r, 2, "FY2026＝隱含（以線 4 基準校準）；線性走到 hs30，其後持平")
        TH[f"{TC[2026]}{r}"] = f"=hum_rev26/(線取數!{TC[2026]}{LR[('hum', 'B')]}*hu_k/1000*線取數!{TC[2026]}{LR['bom']}*act_sh*mixf/1000*fx)"
        for y in range(2027, 2033):
            TH[f"{TC[y]}{r}"] = f"=${TC[2026]}${r}+(hs30_{s}-${TC[2026]}${r})*MIN(1,({y}-2026)/4)"
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"].number_format = "0.00%"
        r += 1
        R["hum"] = r; TH.cell(r, 1, "機器人（人形）致動器營收"); TH.cell(r, 2, "FY2026＝hum_rev26；FY2027 起＝市場 × 份額")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = "=hum_rev26" if y == 2026 else f"={TC[y]}{R['pool']}*{TC[y]}{R['hsh']}"
        r += 1
        R["auto"] = r; TH.cell(r, 1, "汽車零部件營收"); TH.cell(r, 2, "FY2026＝合併營收 − 機器人；FY2027 起＝前一年 ×（1＋gauto）")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = f"={c}{R['rev26']}-{c}{R['hum']}" if y == 2026 else f"={TC[y - 1]}{r}*(1+gauto_{s})"
        r += 1
        R["humsh"] = r; TH.cell(r, 1, "人形占合併營收")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['hum']}/({c}{R['hum']}+{c}{R['auto']})"; TH[f"{c}{r}"].number_format = "0.0%"
        r += 1
        R["gm"] = r; TH.cell(r, 1, "毛利率路徑"); TH.cell(r, 2, "FY2026＝1–9 月；線性到 FY2030 gm30，其後持平")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = "=gm26" if y == 2026 else f"=gm26+(gm30_{s}-gm26)*MIN(1,({y}-2026)/4)"; TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 1
        R["opex"] = r; TH.cell(r, 1, "營業費用及其他率路徑"); TH.cell(r, 2, "FY2026＝opex26；線性到 FY2030 opex30，其後持平")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = "=opex26" if y == 2026 else f"=opex26+(opex30-opex26)*MIN(1,({y}-2026)/4)"; TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 2
        THR[s] = R
    for row in TH.iter_rows(min_row=4, max_row=r, min_col=4, max_col=11):
        for c in row:
            if c.number_format == "General": c.number_format = "#,##0.0"
    TH.column_dimensions["A"].width = 40; TH.column_dimensions["B"].width = 50

    # ================= PL（標準損益表，三情境）=================
    PL["A1"] = ("PL：汽車零部件與人形致動器兩列營收 → 合併營收 → 標準損益表 → FCFF。"
                "C＝FY2023 … E＝FY2025 為實績；F＝FY2026（1–9 月實績＋第 4 季假設）… L＝FY2032。百萬元；資料商標準化口徑。")
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
        C = lambda y: COL[y]
        row("auto", "汽車零部件營收（歷史含未揭露的機器人）", "百萬元", H(8), lambda y: f"=Thesis!{TC[y]}{th['auto']}")
        row("hum", "機器人（人形）致動器營收", "百萬元", lambda y: "未揭露", lambda y: f"=Thesis!{TC[y]}{th['hum']}")
        row("rev", "營收（合併）", "百萬元", H(4), lambda y: f"={C(y)}{R['auto']}+{C(y)}{R['hum']}")
        row("chk", "檢查：分部合計 − 合併營收（應為 0）", "百萬元",
            lambda y: f"={C(y)}{R['auto']}-{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['auto']}+{C(y)}{R['hum']}-{C(y)}{R['rev']}")
        PL.cell(R["chk"], 1, "檢查：分部合計 − 合併營收（應為 0）" if s == "B" else f"檢查（情境 {s}）：分部合計 − 合併營收（應為 0）")
        row("yoy", "營收年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1",
            lambda y: f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1", "0.0%")
        row("cogs", "營業成本", "百萬元", H(5), lambda y: f"={C(y)}{R['rev']}*(1-Thesis!{TC[y]}{th['gm']})")
        row("gp", "毛利", "百萬元", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}")
        row("gm", "毛利率", "%", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", "0.0%")
        row("opex", "營業費用及其他營業損益淨額", "百萬元", H(16), lambda y: f"={C(y)}{R['rev']}*Thesis!{TC[y]}{th['opex']}")
        row("op", "營業利益（資料商口徑）", "百萬元", H(21), lambda y: f"={C(y)}{R['gp']}-{C(y)}{R['opex']}")
        row("opm", "營業利益率", "%", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", "0.0%")
        row("nonop", "其他收入淨額（業外：利息、投資收益等）", "百萬元", H(26), lambda y: f"={C(y)}{R['rev']}*nonop_r")
        row("pbt", "稅前利益", "百萬元", H(32), lambda y: f"={C(y)}{R['op']}+{C(y)}{R['nonop']}")
        row("tax", "所得稅", "百萬元", H(33), lambda y: f"={C(y)}{R['pbt']}*tax_r")
        row("nit", "稅後淨利（含少數股東）", "百萬元", H(34), lambda y: f"={C(y)}{R['pbt']}-{C(y)}{R['tax']}")
        row("nci", "少數股東損益", "百萬元", H(35), lambda y: f"={C(y)}{R['nit']}*nci_r")
        row("ni", "歸屬母公司淨利", "百萬元", H(36), lambda y: f"={C(y)}{R['nit']}-{C(y)}{R['nci']}")
        row("nim", "淨利率", "%", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", "0.0%")
        row("niy", "淨利年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1",
            lambda y: f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1", "0.0%")
        row("sh", "股數（百萬股）", "百萬股", H(39), lambda y: "=sh_v")
        row("eps", "EPS（元）", "元", H(37), lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['sh']}", "0.00")
        row("da", "折舊攤銷", "百萬元", H(40), lambda y: f"={C(y)}{R['rev']}*(da26+(cx_r*da_ss-da26)*{y - 2026}/6)")
        row("cx", "資本支出", "百萬元", H(41), lambda y: "=cx26" if y == 2026 else f"={C(y)}{R['rev']}*cx_r")
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
        R["hcash"] = r; DC.cell(r, 1, "＋股權募資（本模型不假設，0）"); DC[f"B{r}"] = "=0"; r += 1
        R["nc"] = r; DC.cell(r, 1, "淨現金滾到評價日（現金與投資 − 有息負債 − 少數股東權益 ＋ Stub FCFF ＋ H 股募資）")
        DC[f"B{r}"] = f"=cash_v+ltinv_v*inv_k-debt_v-nci_v+B{R['stub']}+B{R['hcash']}"; r += 1
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
    CP["A1"] = ("Comp：同業前瞻 P/E（stockanalysis.com statistics，資料商）。股價＝" + S_CLOSE + "（Verified）；前瞻 P/E＝原值 × 新收盤 ÷ 原報價（Derived，EPS 不變）。"
                "核心同業 6 家進中位數，少於 8 家不做迴歸。三花智控、北特科技等為 A 股，同業中位數受北特、兆威高倍數拉高（見 F 欄）；報價日未取得者不進中位數。")
    hdr(CP, 4, ["公司", "代號", "股價（2026-11-20）", "報價日", "前瞻 P/E 分母", "前瞻 P/E", "近 12 月 P/E（原報價日）", "5 年 β", "會計口徑", "進中位數", "說明", "原報價", "原報價日", "原前瞻 P/E", "標記與換算"])
    PEERS = [  # (公司, 代號, 原報價, 原報價日, 原前瞻 P/E, 近 12 月 P/E, β, 口徑, 進中位數, 說明, 分母)
        ("伯特利", "603596.SS", 18.52, "2026-11-24", 13.87, 18.04, 0.75, "中國會計準則", "是", "汽車電子煞車、底盤（與拓普電子煞車重疊）", "資料商 NTM 共識 EPS"),
        ("三花智控", "002050.SZ", 34.36, "2026-11-24", 27.01, 37.16, 0.78, "中國會計準則", "是", "熱管理、人形致動器（拓普最近的同業）", "資料商 NTM 共識 EPS"),
        ("貝斯特", "300580.SZ", 15.81, "2026-11-24", 24.45, 30.38, 0.63, "中國會計準則", "是", "汽車零件、絲杠與人形致動器零件", "資料商 NTM 共識 EPS"),
        ("北特科技", "603009.SS", 47.80, "2026-11-24", 92.27, 127.56, 0.71, "中國會計準則", "是", "汽車底盤件、行星滾柱絲杠", "資料商 NTM 共識 EPS"),
        ("雙環傳動", "002472.SZ", 41.21, "2026-10-30", 22.32, 27.11, 0.52, "中國會計準則", "是", "齒輪、精密減速器；人形關節（數值取自 leaderdrive 模型 Comp 頁，原報價日較舊）", "資料商 NTM 共識 EPS"),
        ("兆威機電", "003021.SZ", 70.25, "2026-09-30", 59.70, 87.15, 0.44, "中國會計準則", "是", "微型傳動、靈巧手（數值取自 leaderdrive 模型 Comp 頁，原報價日較舊）", "資料商 NTM 共識 EPS"),
        ("匯川技術", "300124.SZ", None, "—", None, None, None, "中國會計準則", "否", "報價日未取得：收盤檔有收盤價，但本輪未取得資料商前瞻 P/E", "—"),
        ("恆立液壓", "601100.SS", None, "—", None, None, None, "中國會計準則", "否", "報價日未取得：收盤檔有收盤價，但本輪未取得資料商前瞻 P/E", "—"),
        ("拓普集團（本公司，對照）", "601689.SS", 54.31, "2026-11-24", 29.56, None, 0.89, "中國會計準則", "否", "本公司；原報價由資料商目標價 56.08 與其標示的高於現價 3.26% 反推（Derived），只作對照", "資料商 NTM 共識 EPS"),
    ]
    for i, (nm, tk, px0, d0, pe0, ttm, bt, acc, med, note, den) in enumerate(PEERS):
        rr = 5 + i
        CP.cell(rr, 1, nm); CP.cell(rr, 2, tk)
        CP.cell(rr, 3, close(tk)).font = BLUE; CP.cell(rr, 4, "2026-11-20"); CP.cell(rr, 5, den)
        if pe0 is not None: CP.cell(rr, 6, f"=N{rr}*C{rr}/L{rr}")
        else: CP.cell(rr, 6, "未取得")
        if ttm is not None: CP.cell(rr, 7, ttm).font = BLUE
        if bt is not None: CP.cell(rr, 8, bt).font = BLUE
        CP.cell(rr, 9, acc); CP.cell(rr, 10, med); CP.cell(rr, 11, note)
        if px0 is not None: CP.cell(rr, 12, px0).font = BLUE
        CP.cell(rr, 13, d0)
        if pe0 is not None: CP.cell(rr, 14, pe0).font = BLUE
        CP.cell(rr, 15, "股價 Verified（" + S_CLOSE + "）；前瞻 P/E Derived：＝原值 × 2026-11-20 收盤 ÷ 原報價（分母 EPS 不變）；原值 " + S_SAP)
    CP.cell(15, 1, "核心同業前瞻 P/E 中位數"); CP["F15"] = "=MEDIAN(F5:F10)"
    CP.cell(16, 1, "本公司現價 ÷ 資料商前瞻 EPS（換算至 2026-11-20；分母：資料商 NTM 共識）"); CP["F16"] = "=F13"
    CP.cell(17, 1, "核心同業前瞻 P/E 最低／最高"); CP["F17"] = "=MIN(F5:F10)"; CP["G17"] = "=MAX(F5:F10)"
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
    row1("cons_show", "資料商平均目標價（18 家，報價日未標；只作對照）", "=cons_tp")
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
        SE.cell(r + 1, 1, "說明：折舊、資本支出（FY2027 起）、營運資金隨營收等比例變動（第 44 條）；單價年降幅同時作用於工業與人形用。")
    SE.column_dimensions["A"].width = 52; SE.column_dimensions["B"].width = 44

    # ================= Recon =================
    RC["A1"] = "Recon：模型對實績與資料商共識的同口徑對照；人形致動器市場對照線 4／線 6；反解。"
    hdr(RC, 3, ["項目", "模型", "對照", "差距（模型 − 對照）", "說明"])
    rr = 4
    def rrow(lab, a, b, note, fmt="#,##0.0", diff=True):
        nonlocal rr
        RC.cell(rr, 1, lab); RC.cell(rr, 2, a).number_format = fmt
        if b is not None: RC.cell(rr, 3, b).number_format = fmt
        if diff and b is not None: RC.cell(rr, 4, f"=B{rr}-C{rr}").number_format = fmt
        RC.cell(rr, 5, short(note)); rr += 1
    rrow("FY2026 毛利率 vs 2026 年 1–9 月實績", f"=PL!F{PR['B']['gm']}", "=gm26", "模型以 1–9 月為錨（差距應為 0）", "0.0%")
    rrow("FY2026 第 4 季隱含營收 vs 2026Q3 實績", f"=PL!F{PR['B']['rev']}-r9m26", "=Q!E4", "第 4 季若維持第 3 季水準即為對照值")
    rrow("FY2026 隱含第 4 季歸屬母公司淨利 vs 2026Q3 實績", f"=PL!F{PR['B']['ni']}-Q!F8", "=Q!E8", "模型 FY2026 淨利減 1–9 月實績")
    rrow("FY2027 EPS（情境 B）vs 資料商前瞻 EPS（現價 54.31 ÷ 29.56＝1.84，NTM）", f"=PL!G{PR['B']['eps']}", "=54.31/29.56", "資料商 NTM 為 FY2026 第 4 季至 FY2027 第 3 季混合，非同期；本專案無券商逐年預測可對照", "0.00")
    rrow("加權目標價 vs 資料商平均目標價（18 家）", f"=Val!B{VR['tp']}", "=cons_tp", S_SA_ST + "；只作對照，不進模型")
    rr += 1
    RC.cell(rr, 1, "人形致動器市場對照（美元；線 4 全尺寸 BOM × 致動器占比 × 組合係數，與線 6 每台零組件內容對照）").font = BOLD; rr += 1
    rrow("FY2026 每台致動器內容（美元，模型）vs 線 6 人形每台零組件內容 2026（基準）", f"=Thesis!E{THR['B']['cont']}", LD["c6"][0], f"{LD['c6'][1]}；線 6 值為全尺寸、不含運算電池結構，未乘出貨組合係數，故不應相等；只看量級", "#,##0")
    rrow("FY2030 每台致動器內容（美元，模型）vs 線 6 人形每台零組件內容 2030（基準）", f"=Thesis!I{THR['B']['cont']}", LD["c6_30"][0], f"{LD['c6_30'][1]}；同上", "#,##0")
    rrow("FY2030 全球人形致動器市場（百萬美元，情境 B）", f"=Thesis!I{THR['B']['pool']}/fx", None, "＝線 4 基準出貨 × 每台致動器內容", "#,##0", False)
    rrow("FY2026 人形隱含全球份額（以線 4 基準校準）", f"=Thesis!E{THR['B']['hsh']}", None, "＝機器人營收假設 ÷（線 4 出貨 × 每台致動器內容 × 匯率）", "0.00%", False)
    rrow("FY2030 人形營收占合併營收（情境 B）", f"=Thesis!I{THR['B']['humsh']}", None, "人形仍屬小比重，目標價主要由汽車本業決定", "0.0%", False)
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
    h26 = LD["ship"][(2026, "B")]; h32 = LD["ship"][(2032, "B")]
    ZR = [
        ("全球人形機器人出貨 FY2026–FY2032（A／B／C＝線高／基準／低）", "4 人形", f"基準 {h26[0]:,.0f} → {h32[0]:,.0f}（× hu_k）", "台", "全球、含小型機", f"{h26[1]}；{h32[1]}（逐年見線取數頁）", "Assumed", "Thesis 人形致動器市場", "一致（hu_k＝1）；高案 hu_k＝2 依媒體引述中國 2026 產量逾 10 萬台，與線 4 基準 5.5 萬台不一致，列 Sens 不改基準"),
        ("致動器占人形 BOM", "4 人形", "=act_sh", "%", "全尺寸", LD["act"][1][1], "Analogy", "Thesis 每台致動器內容", "一致（直接讀值；低、高為線 4 區間）"),
        ("出貨組合係數（出貨組合平均用量 ÷ 全尺寸用量）", "4 人形", "=mixf", "x", "以諧波用量比 8 ÷ 16 代表小型機占比", LD["mix"][1], "Analogy", "Thesis 每台致動器內容", "本模型自設換算：線 4 BOM 為全尺寸，出貨含小型機；以諧波用量比代理致動器價值比，未經驗證"),
        ("全尺寸每台 BOM 2026／2028／2030", "4 人形", f"{LD['bom'][2026][0]:,.0f}／{LD['bom'][2028][0]:,.0f}／{LD['bom'][2030][0]:,.0f}", "美元／台", "基準；2027、2029、2031、2032 為幾何內插", LD["bom"][2026][1], "Derived", "Thesis 每台致動器內容", "一致；三情境共用基準 BOM，未隨出貨情境調整（線 4 高出貨對應較低 BOM，此處保守忽略）"),
        ("線性關節零件（絲杠 14 支、單價 704 美元；線 4）", "4 人形", f"{LD['qscrew'][0]:g} 支 × {LD['screw'][0]:,.0f} 美元", "美元", "僅作對照，不進公式", f"{LD['screw'][1]}；{LD['qscrew'][1]}", "Derived", "未使用", "不使用：拓普賣整個致動器模組，不是單賣絲杠"),
        ("人形每台零組件內容 2026／2030（線 6）", "6 共用零組件", f"{LD['c6'][0]:,.0f}／{LD['c6_30'][0]:,.0f}", "美元／台", "全尺寸、不含運算電池結構", f"{LD['c6'][1]}；{LD['c6_30'][1]}", "Derived", "Recon 對照", "只作量級對照，見 Recon"),
    ]
    for i, z in enumerate(ZR):
        for j, v in enumerate(z): DZ.cell(4 + i, 1 + j, v)
    DZ.column_dimensions["A"].width = 50; DZ.column_dimensions["F"].width = 80; DZ.column_dimensions["I"].width = 70

    # ================= Risk =================
    RK["A1"] = "Risk：風險登錄（原則第 10 節第 10、19 條）。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    RISKS = [
        ("汽車需求與價格壓力（乘用車需求、客戶降價）", "是（gauto、gm30、情境 C）", "汽車為營收與獲利主體", "季度毛利率、客戶產量"),
        ("匯率（人民幣升值影響出口與獲利）", "否", "未建匯率情境", "季度匯兌損益"),
        ("資本支出與折舊上升", "是（cx_r、cx26、da_ss）", "近年資本支出約等於營業現金流，自由現金流薄", "在建工程、現金流量表"),
        ("人形致動器客戶集中與供應商名單未經驗證", "是（hs30 區間、情境 C 1%）", "媒體稱為 Optimus 供應商，財報未揭露", "年報機器人分部、客戶集中度"),
        ("整機廠自製關節，模組整合附加價值被壓縮", "部分（hs30_C）", "致動器模組整合次產業不可替代性最低一級", "整機廠 BOM 揭露"),
        ("人形量產延後或不及線 4 基準", "是（情境 C＝線 4 低）", "", "IDC、TrendForce 季出貨"),
        ("美國關稅與地緣政治（北美客戶與產線）", "否", "未取得北美營收占比", "年報地區別營收"),
        ("控股股東持股 57.88%（外部、未經本專案檢核）", "否", "公司治理與流通性，未建模", "股東變動公告"),
    ]
    for i, k in enumerate(RISKS):
        for j, v in enumerate(k): RK.cell(4 + i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["D"].width = 60

    # ================= Sources =================
    SO["A1"] = "Sources：本模型引用的原始來源與資料商（含網址；HTML 展示層不放網址）。"
    for i, s_ in enumerate([S_PX, S_SA, S_SA_ST, S_RF, S_CRP, S_SAP,
                            f"線 4：{L4N}", f"線 6：{L6N}",
                            S_PROF], 3):
        SO.cell(i, 1, s_)
    SO.column_dimensions["A"].width = 160

    # ================= README（xlsx 內）=================
    S = T
    v = lambda k: f"Val!B{VR[k]}"
    lines = [
        ('="拓普集團 Tuopu（601689.SS）個股模型 ' + VER + '｜build_id "&Sens!D2&"｜資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'),
        "命題：拓普的營收與獲利主體是汽車零部件，人形致動器是尚小的選擇權；營收＝汽車零部件（依年增情境）＋人形致動器（線 4 出貨 × 每台致動器內容 × 組合係數 × 份額）。",
        (f'="結論（條件式）：在 FY2026 以 1–9 月實績為錨、汽車零部件年增 "&{S("gauto_B","0%")}&"、毛利率 FY2030 "&{S("gm30_B","0.0%")}&"、線 4 基準人形出貨（× "&{S("hu_k","0.0")}&"）、情境 B 的 2030 年人形致動器份額 "&{S("hs30_B","0%")}&"、同業 P/E 未加溢價（相對倍數係數 "&{S("rel_k","0.00")}&"）的假設下，加權目標價約 "&{S(v("tp"),"#,##0.0")}&" 元，相對現價 "&{S("px","#,##0.00")}&" 元 "&{S(v("up"),"+0.0%;-0.0%")}&"，估值評等 "&{v("rating")}&"，投資建議 "&{v("rec")}&"。"'),
        (f'="DCF 目標價 "&{S(v("dcf_tp"),"#,##0.0")}&"、Comparable 目標價 "&{S(v("cmp_tp"),"#,##0.0")}&"（各半）；WACC "&{S("wacc","0.00%")}&"、ke "&{S("ke","0.00%")}&"、g "&{S("g_tv","0.00%")}&"；情境 B 終值占 EV "&{S(f"Val!C{VR['tvs']}","0%")}&"。"'),
        (f'="上檔空間歸因：DCF 貢獻 "&{S(v("att_d"),"+0.0%;-0.0%")}&"、未調整 Comparable 貢獻 "&{S(v("att_c"),"+0.0%;-0.0%")}&"、相對倍數調整貢獻 "&{S(v("att_r"),"+0.0%;-0.0%")}&"。本公司現價對資料商前瞻 EPS 的倍數是同業中位數的 "&{S(v("relnow"),"0.0")}&" 倍；支持現價需要相對倍數係數 "&{S(v("relreq"),"0.0")}&"，升一級需要 "&{S(v("relflip"),"0.0")}&"。"'),
        (f'="收支關卡：情境 B FY2026–FY2030 FCFF 扣股利合計 "&{S(v("gate_b"),"#,##0.0")}&" 百萬元，"&{v("gate_t")}&"；距最近級距邊界 "&{S(v("dist"),"0.0%")}&"。"'),
        "主要失效條件（可觀察）：(1) 汽車零部件年增與毛利率偏離情境 B；(2) 資本支出強度不降（自由現金流持續薄）；(3) 機器人營收揭露顯示量級遠高於 Assumed 區間；(4) 同業 P/E 中位數大幅變動。",
        "方法：DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，WACC 折現，明確期到 FY2032，RONIC＝WACC）50%＋同業前瞻 P/E × 模型 FY2027 EPS 50%；三情境 25／50／25。",
        "限制：損益為資料商標準化數字；公司分產品營收與機器人營收未取得（機器人為 Assumed）；同業 P/E 由原報價日按 2026-11-20 收盤等比例換算（EPS 未更新）；不讀 Siliconomics／Tokenomics／Wattonomics。詳見 repo README.md。",
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
        ("情境 B 的 2030 年人形致動器份額改為情境 A 值", "hs30_B", ivE["hs30_A"]),
        ("情境 B 的 2030 年人形致動器份額改為情境 C 值", "hs30_B", ivE["hs30_C"]),
        ("情境 B 汽車零部件年增改為情境 A 值", "gauto_B", ivE["gauto_A"]),
        ("情境 B 汽車零部件年增改為情境 C 值", "gauto_B", ivE["gauto_C"]),
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
    flip_tp = px * (1 + flip + (1e-6 if flip < 0 else 1e-6))
    s1 = solve("hu_k", px, 1.0, 200.0)
    s2 = solve("hu_k", flip_tp, 1.0, 200.0)
    s3 = solve("hs30_B", px, max(ivE["hs30_B"], 0.005), 1.0)
    s4 = solve("gm30_B", px, ivE["gm30_B"], 0.6)
    nf = lambda x, txt: x if x else -1
    solve_rows = [
        ("支持現價（目標價＝現價）所需的人形出貨倍數（相對線 4，全部情境）", nf(s1, ""), 1.0, "其他輸入不變；−1＝200 倍內無解" if not s1 else f"其他輸入不變；情境 B 的 CY2030 全球人形出貨相當於 {LD['ship'][(2030, 'B')][0] * s1 / 1e6:,.1f} 百萬台" if (2030, 'B') in LD['ship'] else ""),
        (f"評等變動（上檔 {flip:+.0%} 邊界）所需的人形出貨倍數", nf(s2, ""), 1.0, "其他輸入不變；−1＝200 倍內無解"),
        ("支持現價所需的情境 B 2030 年人形致動器份額", nf(s3, ""), ivE["hs30_B"], "其他輸入不變；−1＝份額 100% 仍無解"),
        ("支持現價所需的情境 B FY2030 毛利率", nf(s4, ""), ivE["gm30_B"], "其他輸入不變；−1＝毛利率 60% 仍無解"),
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
    json.dump(dict(company="tuopu", ticker="601689.SS", model_sheet="PL",
                   official_segments=[dict(name="汽車零部件及其他（歷史含未揭露的機器人）", label="汽車零部件營收（歷史含未揭露的機器人）", period="FY2025", value=round(Hs["E8"].value, 2), unit="百萬元", source=S_SA)],
                   sublines=[],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）"],
                   note="本輪未取得公司分產品營收（減震器、內飾、底盤、熱管理、電子）；以「汽車零部件」與「人形致動器」兩列呈現，機器人營收公司未揭露，歷史視為 0、預測為 Assumed。"),
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
    t = f"""# physical_ai/tuopu/ — 拓普集團 Tuopu（601689.SS）個股模型 {VER}

<!-- 本檔由 build_tuopu.py 依成品重算值產生；數字不手寫（原則第 10 節第 51 條）。 -->

**本模型適用《高不確定性預估工作原則》**（`docs/plan/高不確定性預估工作原則.md`）：公司未揭露機器人營收、人形需求尚未發生，相關輸入一律給區間並標 Assumed／Analogy；結論只在證據邊界內成立。

**未完成項目（先看這裡）：** (1) 公司分產品營收（減震器、內飾、底盤、熱管理、電子）與客戶集中度本輪未取得，營收只分汽車零部件與人形致動器兩列；(2) 損益為資料商標準化數字，公司財報原文未直接讀取；(3) 沒有券商逐年預測可對照，Recon 只對照資料商前瞻 EPS 與目標價；(4) 三花智控為最近同業，其前瞻 P/E 來自資料商 2026-11-24 頁面按 11-20 收盤換算；匯川技術、恆立液壓報價日未取得、未進中位數；(5) 獨立核對尚未做。

- 成品：`dist/{NAME}`（build_id `{sens['build_id']}`；LibreOffice 重算，0 錯誤值）；HTML 展示層 `dist/{NAME[:-5]}.html`
- 建置：`python3 physical_ai/tuopu/build_tuopu.py`（需要 soffice；整本重算敏感度與反解；結尾自動跑 in_layout、sens_coverage --require-keycol、model_governance、rating_gate --strict、zero_rows、cached_values、segment_check）
- 日期：資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 2026-11-20（收盤 {I['px']:,.2f} 元）｜評價日 2026-12-31｜目標時點 2027-11-20
- 會計年度＝日曆年；單位百萬元人民幣。上海主板。

## 結論（條件式）

在「FY2026 以 1–9 月實績為錨、汽車零部件年增 {I['gauto_B']:.0%}、FY2030 毛利率 {I['gm30_B']:.1%}、線 4 基準人形出貨、情境 B 的 2030 年人形致動器份額 {I['hs30_B']:.0%}、同業 P/E 不加溢價（相對倍數係數 {I['rel_k']:.2f}）」的假設下：

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
| 資料商平均目標價（18 家，報價日未標；只作對照） | {I['cons_tp']:,.2f} |

上檔空間歸因：DCF 貢獻 {g('att_d'):+.1%}、未調整 Comparable 貢獻 {g('att_c'):+.1%}、相對倍數調整 {g('att_r'):+.1%}。本公司現價對資料商前瞻 EPS 的倍數是核心同業中位數（{I['pe_peer']:.1f} 倍）的 {g('relnow'):.1f} 倍；要讓目標價＝現價，Comparable 的相對倍數係數須達 {g('relreq'):.2f}。

**讀法：** 拓普的價值主要由汽車本業決定，人形致動器在預測期仍是小比重。情境 B 人形致動器營收 {ser('hum', '{:,.0f}')}（FY2026→FY2032，百萬元），占合併營收由 {TH[f'E{THB["humsh"]}'].value:.1%} 升到 {TH[f'K{THB["humsh"]}'].value:.1%}。整本重算反解：{sol[0][0]}為 {fn(sol[0][1], '{:,.1f}')}（其他輸入不變）；{sol[2][0]}為 {fn(sol[2][1], '{:,.2f}')}（−1＝無解）；{sol[3][0]}為 {fn(sol[3][1], '{:,.2f}')}。結論主要建立在 Assumed（汽車年增、毛利率路徑、資本支出強度、機器人份額）輸入上；人形變數即使放大到線 4 高情境也只改變小部分價值。

## 命題

1. **營收結構：** 汽車零部件（本業）＋人形致動器。機器人營收公司未揭露，FY2026 以 Assumed {I['hum_rev26']:,.0f} 百萬元（區間 30–300）為錨。
2. **量與價：** 人形致動器營收＝線 4 全球出貨 × 每台致動器內容（線 4 全尺寸 BOM × 致動器占比 × 出貨組合係數）× 匯率 × 份額。2026 年隱含份額 {TH[f'E{THB["hsh"]}'].value:.2%}（以線 4 基準校準）。
3. **汽車本業：** 營收年增依情境 A／B／C {I['gauto_A']:.0%}／{I['gauto_B']:.0%}／{I['gauto_C']:.0%}（FY2027 起）；毛利率由 2026 年 1–9 月水準線性走到 FY2030 的 {I['gm30_A']:.1%}／{I['gm30_B']:.1%}／{I['gm30_C']:.1%}。

情境 B：營收 {ser('rev', '{:,.0f}')}；毛利率 {ser('gm', '{:.1%}')}；EPS {ser('eps', '{:.2f}')} 元（FY2026→FY2032）。

## 主要假設與可證偽指標（原則第 10 節第 20 條）

| 承重假設 | 基準（A／B／C 或區間） | 標記 | 可觀察指標 |
|---|---|---|---|
| 汽車零部件年增（FY2027 起） | {I['gauto_A']:.0%}／{I['gauto_B']:.0%}／{I['gauto_C']:.0%} | Assumed | 季度營收年增、客戶產量 |
| FY2030 毛利率 | {I['gm30_A']:.1%}／{I['gm30_B']:.1%}／{I['gm30_C']:.1%} | Assumed | 季度毛利率（2026 年 1–9 月 18.5%） |
| 資本支出占營收（FY2027 起） | {I['cx_r']:.0%}（8–13%） | Assumed | 現金流量表、在建工程 |
| 全球人形出貨（線 4） | 線 4 高／基準／低；× hu_k {I['hu_k']:.0f}（高案 2） | Assumed（線 4） | IDC、TrendForce 季出貨 |
| 2030 年人形致動器份額 | {I['hs30_A']:.0%}／{I['hs30_B']:.0%}／{I['hs30_C']:.0%} | Assumed | 年報機器人分部、Tesla 供應商揭露 |
| 致動器占 BOM、出貨組合係數 | {I['act_sh']:.0%}、{I['mixf']:.2f} | Analogy（線 4） | 整機拆解 |
| 同業前瞻 P/E | 中位數 {I['pe_peer']:.1f} 倍（6 家） | Derived | 同一估值日重抓 |

## 敏感度（整本重算，前 10 大）

| 案例 | 鍵 | 加權目標價 | 相對基準 | DCF 目標價 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | `{c['key']}` | {c['tp']:,.1f} | {c['tp'] / sens['base_tp'] - 1:+.1%} | {c['dcf']:,.1f} | {c['rating']} |" for c in top) + f"""

完整案例見 xlsx Sens 頁。**翻轉評等所需的變動**：見 Recon 反解；升一級需加權目標價達 {fn(up_b)} 元。

## 反解（Recon 頁）

""" + "\n".join(f"- {a}：{fn(b, '{:,.2f}')}（基準 {fn(c0, '{:,.2f}')}）；{d}" for a, b, c0, d in sol) + f"""

## 限制與「分不出來／沒找到」

- **分不出來（資料性質）：** 機器人營收公司未揭露；人形量級屬未發生事件，線 4 三情境差距約 4–5 倍。
- **沒找到（工作未完）：** 公司財報原文（歷史數字取自資料商）；分產品營收；客戶集中度；北美營收占比；長期投資明細；限制性股票與可轉債；券商逐年預測。
- **線模型：** 讀線 4 出貨、BOM、致動器占比；線 6 人形每台零組件內容只作對照。三情境共用線 4 基準 BOM；出貨組合係數為本模型自設換算。見 xlsx「對帳」頁。
- **同業倍數：** 同業報價一律用收盤檔 2026-11-20；前瞻 P/E 由資料商原報價日等比例換算（EPS 未更新）；雙環傳動、兆威機電原報價日為 2026-10-30、09-30，較舊；匯川技術、恆立液壓報價日未取得、未進中位數；少於 8 家不做迴歸。
- **rf 逾期：** 中國 10 年期國債 1.61% 為 2026-11-05 值（跨模型一致）。
- **協調檔：** 實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics；相關條文不適用。
- **獨立核對：** 尚未由未參與製作者核對。

## 資料來源

見 xlsx Sources 頁（資料商、線 4／6 成品、公司檔案）。

## 版本紀錄

- {VER}（2026-11-24）：首版。現價 {I['px']:,.2f}（2026-11-20 收盤）；加權目標價 {g('tp'):,.1f}；上檔空間 {g('up'):+.1%}；{g('rating')}／{g('rec')}；收支關卡 {g('gate_b'):,.1f}（{g('gate_t')}）。
"""
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write(t)


if __name__ == "__main__":
    main()

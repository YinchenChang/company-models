# 三花智控（002050.SZ，深交所）個股模型 v0.1 建置腳本
# 產出：physical_ai/sanhua/dist/20261124_三花智控個股模型_v0.1.xlsx（LibreOffice 重算後的成品，含公式與快取值）
# 結構：命題模組（Thesis：制冷空調（含液冷）與汽車零部件以年增推導；人形致動器＝線 4 出貨 × 每台致動器價值 × 份額）
#       → 官方分產品（制冷空調電器零部件／汽車零部件）＋機電執行器（另列）→ 標準損益表 → FCFF DCF＋同業 P/E。
# 敏感度：每個案例覆寫 In 頁基準欄後整本重建、LibreOffice 重算；反解同法以二分搜尋。
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md、docs/plan/高不確定性預估工作原則.md。
# 用法：python3 physical_ai/sanhua/build_sanhua.py   （需要 soffice）
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
NAME = f"20261124_三花智控個股模型_{VER}.xlsx"
OUT = os.path.join(HERE, "dist", NAME)
HOMEDIR = os.path.expanduser("~")
TMP = os.path.join(HOMEDIR, "sh_build")
LO = ["soffice", f"-env:UserInstallation=file://{HOMEDIR}/lo_sanhua", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML"]

L3 = os.path.join(ROOT, "physical_ai/lines/industrial_robot/dist/20261116_工業機器人單位經濟_v0.1.xlsx")
L4 = os.path.join(ROOT, "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx")
L6 = os.path.join(ROOT, "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx")
L3N, L4N, L6N = (os.path.relpath(p, ROOT) for p in (L3, L4, L6))

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDRF = PatternFill("solid", fgColor="DDEBF7")
SECF = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------- 出處 ----------------
S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv：002050.SZ；未還原股利）"
S_AR = "公司 2025 年年度報告全文（2026 年 3 月；巨潮資訊 https://static.cninfo.com.cn/finalpage/2026-03-24/1225026522.PDF）：合併財務報表、分產品營收、附註；原文讀取"
S_Q3 = "太平洋證券 三花智控 2026 年三季報點評（2026-10-28；https://data.eastmoney.com/report/info/AP202611031831351033.html）：轉錄公司 2026 年三季報；同花順 2026-10-29 報導同值（https://news.10jqka.com.cn/20261029/c681256306.shtml），含第 3 季分板塊營收"
S_23 = "資料商轉述公司 2024 年財務決算報告對照欄與 2023 年報點評（網頁摘要，原文未直接讀取；https://static.cninfo.com.cn/finalpage/2025-03-27/1222913177.PDF）"
S_BR = "太平洋證券 2026-10-28 三季報點評之盈利預測表（2026E／2027E／2028E 歸母淨利 42.37／50.11／60.46 億元、EPS 1.01／1.19／1.44 元）"
S_MD = "招商證券美的集團報告（2026-04；https://pdf.dfcfw.com/pdf/H3_AP202604301821815625_1.pdf 附近同批報告；經網頁搜尋摘要轉述，原表未直接讀取）：2027E EPS 6.57 元（非三季報後）"
S_IN = "華興證券 匯川技術快報點評（2026-08-25；經網頁搜尋摘要轉述，原表未直接讀取）：2027E EPS 2.61 元（非三季報後）"
S_LD = "physical_ai/leaderdrive 模型 Comp 頁（資料商 NTM 共識前瞻 P/E 與 β；原報價日見 Comp 頁）"
S_RF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與 innolight、eoptolink、victory_giant 同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：China CRP 0.91%（與 innolight 同一取值）"
S_CLOSE = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
IP_BRK = "券商研究報告：有銷售研究與維持評等的誘因；歷史數字為轉錄，預測為券商自設"
IP_CO = "發行公司說法，有股價與融資誘因"

# ---------------- 線模型取數（以 ID／代號找列，不以列號） ----------------
def _line_rows(path, idcol, sheet="下游取數"):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    out = {}
    for r in range(1, ws.max_row + 1):
        k = ws.cell(r, idcol).value
        if isinstance(k, str): out[k] = (r, [ws.cell(r, c).value for c in range(1, 10)])
    return wb, out


def _f(x): return float(x)


def line_data():
    _, l4 = _line_rows(L4, 1)
    _, l6 = _line_rows(L6, 1)
    d = {"ship": {}, "bom": {}}
    for y in range(2026, 2033):
        r, v = l4[f"SHIP_{y}"]
        for s, c, col in (("A", 4, "E"), ("B", 3, "D"), ("C", 2, "C")):
            d["ship"][(y, s)] = (_f(v[c]), f"{L4N} 下游取數!{col}{r}（SHIP_{y}）")
    r, v = l4["HUM_BOM_2026"]
    for s in "ABC": d["bom"][(2026, s)] = (_f(v[3]), f"{L4N} 下游取數!D{r}（HUM_BOM_2026 基準；2026 的低／高欄是輸入全取低／高，不是出貨情境，三情境同用基準）")
    for y in (2028, 2030, 2035):
        r, v = l4[f"HUM_BOM_{y}"]
        for s, c, col in (("A", 4, "E"), ("B", 3, "D"), ("C", 2, "C")):
            d["bom"][(y, s)] = (_f(v[c]), f"{L4N} 下游取數!{col}{r}（HUM_BOM_{y}；低出貨情境 BOM 較高）")
    r, v = l4["HUM_ACT_SHARE"]; d["act"] = [(_f(v[2]), _f(v[3]), _f(v[4])), f"{L4N} 下游取數!C{r}:E{r}（HUM_ACT_SHARE）"]
    r1, a = l4["Q_HARMONIC_AVG"]; r2, b = l4["Q_ROTARY"]
    d["eq"] = [(_f(a[2]) / _f(b[2]), _f(a[3]) / _f(b[3]), _f(a[4]) / _f(b[4])), f"{L4N} 下游取數!C{r1}:E{r1}（Q_HARMONIC_AVG）÷ C{r2}:E{r2}（Q_ROTARY）"]
    r, v = l6["人形每台零組件內容 2026（基準）"]; d["l6c"] = (_f(v[2]), f"{L6N} 下游取數!C{r}")
    return d


LD = line_data()

# ---------------- In 頁輸入 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 來源誘因, 說明)
IN = [
    ("sec", "1. 市場資料與日期"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, "", "資料截止日 2026-11-24（財報、券商預測、同業倍數）；股價 2026-11-20；成品日 2026-11-24"),
    ("px", "現價（2026-11-20 收盤，深圳 A 股）", "元", None, 31.40, None, "Verified", S_PX, "", "csv 值 31.400000；11-09 收盤 34.24，兩週跌 8%"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,12,31)", None, "Assumed", "本專案慣例：評價日 2026-12-31（會計年度＝日曆年，FY2026 在評價日已結束）", "", "折現期間由評價日起算"),
    ("tp_date", "目標時點（px_date 後 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", ""),
    ("bs_date", "資產負債表日（最近一期已讀原文：2025 年報）", "日期", None, "=DATE(2025,12,31)", None, "Verified", S_AR, "", "2026 年三季報資產負債表未找到；以年報日起算，Stub 期＝整個 FY2026"),
    ("stub_frac", "Stub 期（資產負債表日→評價日）", "年", None, "=(val_date-bs_date)/365", None, "Derived", "＝(評價日 − 資產負債表日) ÷ 365", "", "Stub FCFF＝FY2026 FCFF × stub_frac（約 1.0）"),
    ("sh_v", "評價日股數（總股本，A＋H）", "百萬股", None, 4208.01, None, "Verified", S_AR + "：股份總數 4,208,013,935 股（含 2025 年 H 股發行後；回購戶 8.35 百萬股未扣）", "", "限制性股票尚未解鎖部分未另計"),
    ("cash_v", "現金＋交易性金融資產（2025-12-31）扣已宣告股利", "百萬元", None, "=14911.99+508.33-1175.91", None, "Verified", S_AR + "：貨幣資金 14,911.99、交易性金融資產 508.33；扣 2025 年度擬派現金股利 1,175.91（每 10 股 2.8 元，2026 年支付）", "", "2025 年中 H 股募資淨額約 96.5 億元使現金大增"),
    ("ltinv_v", "長期股權投資（2025-12-31）", "百萬元", None, 44.59, None, "Verified", S_AR, "", ""),
    ("inv_k", "長期投資計入淨現金的比例", "%", 0.5, 1.0, None, "Assumed", "基準全額計入；低案 50%", "", "金額小（約 0.1%市值）"),
    ("debt_v", "有息負債（2025-12-31；含租賃負債）", "百萬元", None, 4051.97, None, "Verified", S_AR + "：短期借款 1,391.78＋一年內到期非流動負債 1,591.56（含部分租賃）＋長期借款 862.70＋租賃負債 205.93", "", "一年內到期項含租賃，與租賃負債可能小幅重疊（偏保守）"),
    ("nci_v", "少數股東權益（2025-12-31）", "百萬元", None, 291.30, None, "Verified", S_AR, "", "自淨現金扣除"),
    ("fx", "人民幣兌美元（2026-11-20，CNY 每美元）", "元／美元", None, 6.6665, None, "Verified", S_CLOSE + "：CNY=X", "", "用於人形致動器營收的美元換算"),
    ("sec", "2. 實績錨（FY2026＝2026 年，會計年度＝日曆年）"),
    ("r9m26", "2026 年 1–9 月營收", "百萬元", None, 25274.0, None, "Verified", S_Q3, "", "公司三季報原文未直接讀取"),
    ("g9m26", "2026 年 1–9 月營收年增", "%", None, 0.0518, None, "Verified", S_Q3, "", "第 3 季單季 +7.82%"),
    ("q4g", "2026 年第 4 季營收年增", "%", 0.03, 0.07, 0.12, "Assumed", "基準＝第 3 季年增 7.8% 與 1–9 月 5.2% 之間", "", "FY2026 營收＝1–9 月實績＋2025Q4 ×（1＋q4g）；2025Q4 營收 6,982（年報分季表）"),
    ("sh_c26", "FY2026 制冷空調電器零部件 占營收", "%", None, "=48.8/(48.8+34.9)", None, "Derived", S_Q3 + "：第 3 季制冷 48.8 億、汽零 34.9 億；1–9 月分板塊未找到，以第 3 季比重代表", "", "2025 年為 59.9%；汽車零部件＝餘數；FY2026 人形致動器營收以 0 計（上半年尚無收入，見 Thesis）"),
    ("gm26", "FY2026 毛利率", "%", 0.275, 0.2889, 0.295, "Assumed", S_Q3 + "：第 3 季毛利率 28.89%（+0.87pt）；2025 年全年 28.78%（年報）", "", "1–9 月毛利率未找到，以第 3 季代表"),
    ("opex26", "FY2026 銷售＋管理＋研發費用率", "%", 0.12, 0.131, 0.14, "Assumed", S_Q3 + "：第 3 季銷售 1.87%、管理 5.94%、研發 5.29%（合計 13.1%）；2025 年 13.0%（年報）", "", ""),
    ("nonop26", "FY2026 業外淨額（財務費用含匯兌、投資收益、公允價值）", "百萬元", -600, -350, -100, "Assumed", S_Q3 + "：第 3 季財務費用率 0.71%（匯兌損失）、淨利年減 9.6% 而扣非年增；以 1–9 月淨利 30.67 億元與第 4 季營收推得全年業外約 −3.5 億元（Derived 推算，區間 Assumed）；2025 年為 −1.1 億元（Hist）", "", "2027 年起改用 inv_y × 淨現金"),
    ("sec", "3. 命題模組：核心板塊成長與人形致動器（線 4 出貨 × 單台致動器價值 × 份額）"),
    ("hu_k", "人形出貨相對線 4 的倍數（全部情境）", "x", None, 1.0, 2.0, "Assumed", "基準＝線 4 原值；高案 2 倍為判斷（中國 2026 年人形產量報導遠高於線 4 全球 5.5 萬台，見對帳）", "", "反解變數"),
    ("act_sh", "致動器占全尺寸人形 BOM（線 4）", "%", LD["act"][0][0], LD["act"][0][1], LD["act"][0][2], "Analogy", LD["act"][1] + "（低／基準／高）", "", "致動器＝旋轉＋線性關節，含模組整合；本公司供應其中的比例用份額表示"),
    ("fs_eq", "全尺寸等效係數（出貨組合中相當於全尺寸機的比例）", "x", LD["eq"][0][0], LD["eq"][0][1], LD["eq"][0][2], "Analogy", LD["eq"][1] + "：全球平均每台諧波用量 ÷ 全尺寸每台用量，作為組合中小型機的折算", "", "判斷：以諧波用量比代表致動器價值比；實際小型機價值比未知"),
    ("hs30_A", "情境 A：2030 年占全球人形致動器價值份額", "%", None, 0.15, None, "Assumed", "媒體稱本公司為 Optimus 致動器供應商（Interested-party：媒體，未經本專案檢核）；2026 上半年尚無收入（同花順 2026-10-29）；份額無任何揭露", "", "FY2026＝0，線性走到 2030，其後持平"),
    ("hs30_B", "情境 B：2030 年占全球人形致動器價值份額", "%", None, 0.06, None, "Assumed", "同上；基準假設只拿到 Tesla 一家整機廠的一部分，且整機廠可能自製或跳過模組廠（地圖檔：Tier 1 不可替代性最低一級）", "", ""),
    ("hs30_C", "情境 C：2030 年占全球人形致動器價值份額", "%", None, 0.0, None, "Assumed", "同上；零＝報導不成真或被替代", "", ""),
    ("gcool_A", "情境 A：制冷空調（含液冷）營收年增（FY2027 起）", "%", None, 0.10, None, "Assumed", S_Q3 + "：第 3 季制冷 +6.5%（資料中心液冷貢獻增量）；2025 年 +12.2%（年報）", "", "液冷規模公司未揭露，併入此成長率"),
    ("gcool_B", "情境 B：制冷空調（含液冷）營收年增（FY2027 起）", "%", None, 0.07, None, "Assumed", "同上", "", ""),
    ("gcool_C", "情境 C：制冷空調（含液冷）營收年增（FY2027 起）", "%", None, 0.03, None, "Assumed", "同上；家電需求與關稅壓力", "", ""),
    ("gauto_A", "情境 A：汽車零部件營收年增（FY2027 起）", "%", None, 0.12, None, "Assumed", S_Q3 + "：第 3 季汽零 +9.6%；2025 年 +9.1%（年報）", "", ""),
    ("gauto_B", "情境 B：汽車零部件營收年增（FY2027 起）", "%", None, 0.08, None, "Assumed", "同上", "", ""),
    ("gauto_C", "情境 C：汽車零部件營收年增（FY2027 起）", "%", None, 0.03, None, "Assumed", "同上；電動車成長放緩", "", ""),
    ("gm30_A", "情境 A：FY2030 毛利率", "%", None, 0.305, None, "Assumed", "2023 年 27.6%、2024 年 27.5%、2025 年 28.8%（Hist，Derived）；趨勢緩升", "", "FY2026 起線性走到 FY2030，其後持平"),
    ("gm30_B", "情境 B：FY2030 毛利率", "%", None, 0.295, None, "Assumed", "同上", "", ""),
    ("gm30_C", "情境 C：FY2030 毛利率", "%", None, 0.275, None, "Assumed", "同上；回到 2023–2024 水準", "", ""),
    ("tp_wA", "情境權重 A", "", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", ""),
    ("tp_wB", "情境權重 B", "", None, 0.50, None, "Assumed", "同上", "", ""),
    ("tp_wC", "情境權重 C", "", None, 0.25, None, "Assumed", "同上", "", ""),
    ("sec", "5. 損益與現金流"),
    ("ts_r", "稅金及附加 ÷ 營收", "%", None, "=Hist!E19/Hist!E4", None, "Derived", "＝2025 年稅金及附加 ÷ 營收", "", ""),
    ("opex30", "FY2030 銷售＋管理＋研發費用率", "%", 0.115, 0.125, 0.135, "Assumed", "2023 年 13.1%、2024 年 13.9%、2025 年 13.0%（Hist）；規模效益", "", "FY2026→FY2030 線性，其後持平"),
    ("oi_r", "其他收益（政府補助等）÷ 營收", "%", 0.005, "=Hist!E24/Hist!E4", 0.015, "Derived", "＝2025 年其他收益 342 ÷ 營收（" + "年報" + "）", "", "計入營業利益"),
    ("imp_r", "減值損失（資產＋信用）÷ 營收", "%", 0.001, 0.003, 0.008, "Assumed", "2025 年資產減值 122.7、信用減值回沖 23.5，淨 99.2（0.32%）；2024 年 119.3（0.43%）", "", "計入營業利益"),
    ("inv_y", "現金與投資淨收益率（利息收入 − 利息費用 − 匯兌，占期初淨現金與投資）", "%", 0.0, 0.006, 0.015, "Assumed", S_AR + "：2025 年利息收入 262.6、利息費用 145.8、匯兌損失使財務費用 101.2；" + S_Q3 + "：第 3 季財務費用率 0.71%（匯兌損失）", "", "業外；DCF 另以淨現金計入，不重複"),
    ("tax_r", "有效稅率", "%", 0.12, "=Hist!E33/Hist!E32", 0.20, "Derived", "＝2025 年所得稅 ÷ 利潤總額（高新技術企業優惠）", "", "敏感度低／高 12%／20%（Assumed）"),
    ("nci_r", "少數股東損益 ÷ 稅後淨利", "%", None, "=Hist!E35/Hist!E34", None, "Derived", "＝2025 年少數股東損益 ÷ 淨利", "", ""),
    ("payout", "配息率", "%", None, "=0.28/1.03", None, "Derived", "＝2025 年每股股利 0.28 元（每 10 股 2.8 元；年報）÷ EPS 1.03 元", "", "用於收支關卡與目標時點前股利"),
    ("da26", "FY2026 折舊攤銷 ÷ 營收", "%", None, "=Hist!E40/Hist!E4", None, "Derived", "＝2025 年折舊攤銷 ÷ 營收", "", "之後線性收斂到資本支出比例 × da_ss（FY2032）"),
    ("da_ss", "穩態折舊 ÷ 資本支出", "x", None, 0.8, None, "Assumed", "穩態資本支出高於折舊（成長）", "", ""),
    ("cx26", "FY2026 資本支出", "百萬元", 2500, 3000, 3500, "Assumed", S_AR + "：2025 年購建長期資產支付 3,152、2024 年 3,290", "", "擴產（液冷、汽零、機電執行器海外產能）"),
    ("cx_r", "資本支出 ÷ 營收（FY2027 起）", "%", 0.06, 0.08, 0.10, "Assumed", "2024 年 11.8%、2025 年 10.2%（Hist，Derived）；擴產高峰後回落", "", ""),
    ("nwc_r", "營運資金增加 ÷ 營收增量", "%", 0.10, 0.20, 0.35, "Assumed", S_AR + "：2025 年應收 7,370、存貨 5,640、應付（票據＋帳款）10,189", "", "淨營運資金約占營收一成至二成"),
    ("h_iss", "再融資占擴大後股本（尾端事件，基準 0）", "%", None, 0.0, 0.05, "Assumed", "2025 年已完成 H 股發行（募資淨額 96.5 億元）；再融資無公告，列尾端風險", "", "原則第 10 節第 19 條：事件型尾端風險進 Sens、不進基準"),
    ("h_disc", "再融資價格相對現價折價", "%", None, 0.10, None, "Assumed", "慣例折價（未查核，不作依據）", "", "只在 h_iss＞0 時有作用"),
    ("sec", "6. 折現率（DCF 標準規格第 2 節）"),
    ("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", S_RF, "", "日期 2026-11-05，較股價基準日早 11 個交易日（逾期；跨模型一致，下次全面改版更新）"),
    ("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", ""),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", ""),
    ("b_own", "本公司 5 年 β（原始）", "x", None, 0.90, None, "Assumed", "資料商 β 頁讀取失敗（未找到）；以 0.9 暫代（A 股大型製造股常見 0.8–1.1，未查核）", "", "β 另有 ∓0.3 敏感度"),
    ("b_peer", "同業 5 年 β 中位數（原始）", "x", None, "=MEDIAN(Comp!H5:H8)", None, "Derived", "＝Comp 頁核心同業 β 中位數（美的、匯川 β 未找到，只含拓普、雙環）", "", ""),
    ("beta", "β（兩來源 Blume 調整後中位數）", "x", "=beta-0.3", "=MEDIAN(0.67*b_own+0.33,0.67*b_peer+0.33)", "=beta+0.3", "Derived", "Blume 0.67β＋0.33；兩來源＝本公司、同業中位數", "", "敏感度 ∓0.3（跨模型慣例）"),
    ("kd", "稅前債務成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt", "", ""),
    ("wacc_adj", "WACC 調整（敏感度用）", "pt", -0.01, 0, 0.01, "Assumed", "基準 0；敏感度 ±1pt", "", ""),
    ("g_cap", "永續成長率上限", "%", None, 0.03, None, "Assumed", "DCF 標準規格：3%，且不超過計價幣別 10 年期公債殖利率", "", ""),
    ("g_tv", "永續成長率", "%", 0.01, "=MIN(g_cap,rf)", 0.03, "Derived", "＝MIN(g_cap, rf)", "", "敏感度 1%／3%（3% 超過 rf，只作敏感度）"),
    ("ronic_sp", "RONIC − WACC", "pt", None, 0, 0.05, "Assumed", "RONIC 預設＝WACC（新投資價值中性）", "", ""),
    ("tv_alt", "終值改用 FCFF ×（1＋g）（0＝否，1＝是）", "開關", None, 0, 1, "Assumed", "DCF 標準規格第 4 節敏感度", "", ""),
    ("sec", "7. 同業倍數與權重"),
    ("pe_peer", "同業前瞻 P/E 中位數", "x", "=MIN(Comp!F5:F8)", "=MEDIAN(Comp!F5:F8)", "=MAX(Comp!F5:F8)", "Derived", "＝Comp 頁核心同業 4 家（美的、匯川、拓普、雙環）前瞻 P/E 中位數；股價皆為 2026-11-20 收盤，P/E 的 EPS 分母口徑不一（見 Comp）", "", "敏感度＝同業最低／最高"),
    ("rel_k", "相對倍數係數（本公司 ÷ 同業）", "x", 0.8, 1.0, 1.3, "Assumed", "1.00（未調整：缺本公司對同業十年相對倍數歷史）", "", "低 0.8／高 1.3（未查核，不作依據）"),
    ("w_dcf", "DCF 權重（Comparable＝1 − 本值）", "%", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設 DCF 50%＋同業倍數 50%", "", ""),
]
SCN_DRIVERS = [("gcool", "制冷空調（含液冷）營收年增"), ("gauto", "汽車零部件營收年增"), ("hs30", "2030 年人形致動器份額"), ("gm30", "FY2030 毛利率")]

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
    IW = wb.create_sheet("In"); HI = wb.create_sheet("Hist"); LN = wb.create_sheet("線取數"); TH = wb.create_sheet("Thesis")
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
    HI["A1"] = "Hist：歷史財務（百萬元人民幣；中國會計準則）。FY2024、FY2025 取自公司 2025 年年報原文；FY2023 為資料商轉述（部分項目未取得，標註處以 0 或比例估計）。公式只引用本頁（定義名稱 h_*）。"
    hdr(HI, 3, ["項目", "", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"])
    HROWS = [  # (列, 鍵, 項目, [2023, 2024, 2025], 標記, 來源)
        (4, "h_rev", "營業收入", [24557.80, 27947.164516, 31011.74451], "Verified", S_AR + "（2025、2024）；2023：" + S_23),
        (5, "h_cogs", "營業成本", [17773.86, 20269.83, 22087.50], "Verified", S_AR + "（2025、2024）；2023：" + S_23 + "（另一券商表列 17,708，差 0.4%）"),
        (6, "h_gp", "毛利（＝營收 − 營業成本）", ["=C4-C5", "=D4-D5", "=E4-E5"], "Derived", "＝營業收入 − 營業成本"),
        (7, "h_gm", "毛利率", ["=C6/C4", "=D6/D4", "=E6/E4"], "Derived", "＝毛利 ÷ 營收（年報 2025 年 28.78%）"),
        (8, "h_seg_c", "制冷空調電器零部件 營收", [None, 16560.605388, 18584.743718], "Verified", S_AR + "：分產品營收（2025 年 +12.22%）；2023 年分部未取得"),
        (9, "h_seg_a", "汽車零部件 營收", [None, 11386.559128, 12427.000792], "Verified", S_AR + "：分產品營收（2025 年 +9.14%）；2023 年分部未取得"),
        (10, "h_seg_h", "仿生機器人機電執行器 營收（未單獨揭露，以 0 計）", [None, 0, 0], "Assumed", "年報僅稱配合客戶研發、試製、送樣；分產品營收無此項；2026 上半年尚未產生收入（同花順 2026-10-29，外部報導）"),
        (13, "h_sell", "銷售費用", [597.57, 726.44, 751.53], "Verified", S_AR + "；2023：" + S_23),
        (14, "h_admin", "管理費用", [1476.33, 1767.45, 1909.33], "Verified", S_AR + "；2023：" + S_23),
        (15, "h_rnd", "研發費用", [1096.83, 1351.80, 1374.07], "Verified", S_AR + "；2023：" + S_23),
        (16, "h_opex", "營業費用（銷售＋管理＋研發）", ["=SUM(C13:C15)", "=SUM(D13:D15)", "=SUM(E13:E15)"], "Derived", "＝三項合計"),
        (17, "h_fin", "財務費用（負＝淨利息收入）", [-72.75, -43.78, 101.22], "Verified", S_AR + "；2023：" + S_23),
        (18, "h_inv", "投資收益＋公允價值變動", [0, "=-16.28-90.73", "=-7.62+11.68"], "Verified", S_AR + "（2025：投資收益 −7.62、公允價值 +11.68；2024：−16.28、−90.73）；2023 年未取得，以 0 計（由殘差項吸收）"),
        (19, "h_ts", "稅金及附加", ["=C4*D19/D4", 171.28, 175.29], "Verified", S_AR + "；2023 年未取得，以 2024 年占營收比例估（Derived）"),
        (20, "h_opp", "營業利潤（中國會計準則）", [3552.96, 3710.06, 4859.29], "Verified", S_AR + "；2023：" + S_23),
        (21, "h_ebit", "營業利益（核心：營業利潤 − 投資收益與公允價值 ＋ 財務費用）", ["=C20-C18+C17", "=D20-D18+D17", "=E20-E18+E17"], "Derived", "本模型營業利益不含金融資產收益與財務費用（DCF 另以淨現金計入）"),
        (22, "h_oth_op", "其他營業損益（＝核心營業利益 − 毛利 ＋ 稅金及附加 ＋ 營業費用）", ["=C21-C6+C19+C16", "=D21-D6+D19+D16", "=E21-E6+E19+E16"], "Derived", "其他收益、減值、資產處置等合計（殘差；2023 年含未取得項目的誤差）"),
        (23, "h_imp", "減值損失淨額（資產減值＋信用減值）", [None, "=-56.38-62.92", "=23.48-122.71"], "Verified", S_AR),
        (24, "h_oi", "其他收益", [None, 229.21, 342.15], "Verified", S_AR),
        (25, "h_ebitm", "營業利益率（核心）", ["=C21/C4", "=D21/D4", "=E21/E4"], "Derived", ""),
        (32, "h_pbt", "利潤總額（稅前）", [3553.26, 3691.69, 4843.76], "Verified", S_AR + "；2023：" + S_23),
        (26, "h_nonop", "業外（＝稅前 − 核心營業利益）", ["=C32-C21", "=D32-D21", "=E32-E21"], "Derived", "投資收益、公允價值、財務費用、營業外收支合計"),
        (33, "h_tax", "所得稅", ["=C32-C34", 579.96, 737.07], "Verified", S_AR + "；2023＝利潤總額 − 稅後淨利（Derived；少數股東損益未取得以 0 計）"),
        (34, "h_nit", "稅後淨利", ["=C36+C35", 3111.73, 4106.69], "Verified", S_AR + "；2023＝歸母淨利＋少數股東損益"),
        (35, "h_nci", "少數股東損益", [0, 12.56, 43.80], "Verified", S_AR + "；2023 年未取得，以 0 計"),
        (36, "h_ni", "歸屬母公司淨利", [2920.99, 3099.17, 4062.89], "Verified", S_AR),
        (37, "h_eps", "EPS（基本，元）", [0.81, 0.84, 1.03], "Verified", S_AR + "（加權股數；2025 年因 H 股發行股數增加）"),
        (38, "h_dps", "每股股利（元）", [None, "=932.42/3732.39", 0.28], "Verified", S_AR + "：2025 年每 10 股 2.8 元；2024 年現金股利 932.42 ÷ 股本 3,732.39"),
        (39, "h_sh", "股數（＝淨利 ÷ EPS，百萬股；加權平均）", ["=C36/C37", "=D36/D37", "=E36/E37"], "Derived", ""),
        (40, "h_da", "折舊攤銷", [None, "=832.16+117.16+46.18+33.33", "=1055.05+104.29+35.47+61.86"], "Verified", S_AR + "：現金流量表補充資料（固定資產折舊＋使用權資產折舊＋無形資產攤銷＋長期待攤費用攤銷）"),
        (41, "h_cx", "資本支出（購建固定資產、無形資產和其他長期資產支付的現金）", [None, 3290.15, 3152.39], "Verified", S_AR),
        (42, "h_ocf", "營業現金流", [3723.96, 4366.63, 5090.76], "Verified", S_AR),
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
    HI.cell(28, 1, "檢查：分部合計 − 營業收入（FY2024、FY2025 應為 0）")
    for c in "DE": HI[f"{c}28"] = f"={c}8+{c}9+{c}10-{c}4"
    HI.cell(29, 1, "差額：核心營業利益＋業外 − 稅前（應為 0）")
    for c in "CDE": HI[f"{c}29"] = f"={c}21+{c}26-{c}32"
    HI.column_dimensions["A"].width = 58; HI.column_dimensions["G"].width = 90

    # ================= 線取數（線模型下游取數；藍字＝讀入值，M 欄寫來源檔與儲存格）=================
    LN["A1"] = "線取數：本模型讀入的線 4 變數年度路徑（建置時由線模型成品依 ID 讀值）。情境對應：A＝線模型高情境、B＝基準、C＝低情境。BOM 的 2027、2029、2031、2032 為線性內插（Derived）。"
    hdr(LN, 3, ["項目", "情境", ""] + [f"FY{y}" for y in range(2025, 2033)] + ["單位", "來源（檔案與儲存格）", "FY2035"])
    LR = {}
    r = 4
    for s in SC:
        LN.cell(r, 1, "全球人形機器人出貨（線 4）"); LN.cell(r, 2, f"{s}（{LINE_SC[s]}）")
        for y in range(2026, 2033):
            v, src = LD["ship"][(y, s)]
            c = LN.cell(r, 4 + (y - 2025), v); c.font = BLUE; c.number_format = "#,##0"
        LN.cell(r, 12, "台"); LN.cell(r, 13, f"{LD['ship'][(2026, s)][1]} … {LD['ship'][(2032, s)][1]}")
        LR[("ship", s)] = r; r += 1
    for s in SC:
        LN.cell(r, 1, "全尺寸人形每台 BOM（線 4）"); LN.cell(r, 2, f"{s}（{LINE_SC[s]}）")
        for y, col in ((2026, "E"), (2028, "G"), (2030, "I")):
            c = LN[f"{col}{r}"]; c.value = LD["bom"][(y, s)][0]; c.font = BLUE; c.number_format = "#,##0"
        c = LN[f"N{r}"]; c.value = LD["bom"][(2035, s)][0]; c.font = BLUE; c.number_format = "#,##0"
        LN[f"F{r}"] = f"=(E{r}+G{r})/2"; LN[f"H{r}"] = f"=(G{r}+I{r})/2"
        LN[f"J{r}"] = f"=I{r}+(N{r}-I{r})/5"; LN[f"K{r}"] = f"=I{r}+(N{r}-I{r})*2/5"
        for col in "FHJK": LN[f"{col}{r}"].number_format = "#,##0"
        LN.cell(r, 12, "美元／台"); LN.cell(r, 13, f"{LD['bom'][(2026, s)][1]}；{LD['bom'][(2028, s)][1]}；{LD['bom'][(2030, s)][1]}；{LD['bom'][(2035, s)][1]}")
        LR[("bom", s)] = r; r += 1
    LN.column_dimensions["A"].width = 40; LN.column_dimensions["M"].width = 110

    # ================= Thesis（命題模組）=================
    TH["A1"] = ("命題模組：合併營收＝制冷空調電器零部件（含資料中心液冷；FY2026 錨 × 年增）＋汽車零部件（FY2026 錨 × 年增）＋仿生機器人機電執行器"
                "（線 4 出貨 × 全尺寸等效係數 × 全尺寸每台 BOM × 致動器占比 × 本公司份額 × 匯率）。FY2026＝1–9 月實績＋第 4 季假設；人形致動器 FY2026＝0。百萬元。")
    TH["A2"] = "欄：D＝FY2025 … K＝FY2032。份額 FY2026＝0，線性走到 2030 年值，其後持平。"
    hdr(TH, 3, ["項目", "說明", ""] + [f"FY{y}" for y in range(2025, 2033)])
    THR = {}
    r = 4
    for s in SC:
        TH.cell(r, 1, f"情境 {s}").font = BOLD; TH.cell(r, 1).fill = SECF; r += 1
        R = {}
        R["rev26"] = r; TH.cell(r, 1, "FY2026 合併營收（錨）"); TH.cell(r, 2, "＝1–9 月＋2025Q4 ×（1＋q4g）")
        TH[f"{TC[2026]}{r}"] = "=r9m26+(Hist!E4-r9m26/(1+g9m26))*(1+q4g)"; r += 1
        R["cool"] = r; TH.cell(r, 1, "制冷空調電器零部件（含液冷） 營收"); TH.cell(r, 2, "FY2027 起＝前一年 ×（1＋gcool）")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['rev26']}*sh_c26" if y == 2026 else f"={TC[y - 1]}{r}*(1+gcool_{s})"
        r += 1
        R["auto"] = r; TH.cell(r, 1, "汽車零部件 營收"); TH.cell(r, 2, "FY2027 起＝前一年 ×（1＋gauto）")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['rev26']}*(1-sh_c26)" if y == 2026 else f"={TC[y - 1]}{r}*(1+gauto_{s})"
        r += 1
        R["ship"] = r; TH.cell(r, 1, "全球人形出貨（線 4 × 出貨倍數，千台）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR[('ship', s)]}*hu_k/1000"
        r += 1
        R["bom"] = r; TH.cell(r, 1, "全尺寸每台致動器價值（美元＝BOM × 致動器占比）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR[('bom', s)]}*act_sh"
        r += 1
        R["mkt"] = r; TH.cell(r, 1, "全球人形致動器市場（百萬美元）"); TH.cell(r, 2, "＝出貨 × 全尺寸等效係數 × 每台致動器價值")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['ship']}*1000*fs_eq*{TC[y]}{R['bom']}/1000000"
        r += 1
        R["hsh"] = r; TH.cell(r, 1, "本公司人形致動器份額"); TH.cell(r, 2, "FY2026＝0；線性走到 hs30，其後持平")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = 0 if y == 2026 else f"=hs30_{s}*MIN(1,({y}-2026)/4)"; TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 1
        R["hum"] = r; TH.cell(r, 1, "仿生機器人機電執行器 營收"); TH.cell(r, 2, "＝市場（百萬美元）× 份額 × 匯率")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['mkt']}*{TC[y]}{R['hsh']}*fx"
        r += 1
        R["humsh"] = r; TH.cell(r, 1, "人形致動器占合併營收")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['hum']}/({c}{R['cool']}+{c}{R['auto']}+{c}{R['hum']})"; TH[f"{c}{r}"].number_format = "0.0%"
        r += 1
        R["gm"] = r; TH.cell(r, 1, "毛利率路徑"); TH.cell(r, 2, "FY2026＝gm26；線性到 FY2030 gm30，其後持平（人形致動器毛利率未另設，併入合併）")
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
    TH.column_dimensions["A"].width = 44; TH.column_dimensions["B"].width = 50

    # ================= PL（標準損益表，三情境）=================
    PL["A1"] = ("PL：官方分產品（制冷空調電器零部件／汽車零部件）＋機電執行器（另列）→ 合併營收 → 標準損益表 → FCFF。C＝FY2023 … E＝FY2025 為實績；F＝FY2026（1–9 月實績＋第 4 季假設）… L＝FY2032。百萬元；中國會計準則。")
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
        hs = lambda rr: (lambda y: None if y == 2023 else f"=Hist!{'CDE'[y - 2023]}{rr}")
        row("cool", "制冷空調電器零部件（含資料中心液冷） 營收", "百萬元", hs(8), lambda y: f"=Thesis!{TC[y]}{th['cool']}")
        row("auto", "汽車零部件 營收", "百萬元", hs(9), lambda y: f"=Thesis!{TC[y]}{th['auto']}")
        row("hum", "仿生機器人機電執行器 營收（未單獨揭露；模型另列）", "百萬元", hs(10), lambda y: f"=Thesis!{TC[y]}{th['hum']}")
        row("rev", "營收（合併）", "百萬元", H(4), lambda y: f"=SUM({C(y)}{R['cool']}:{C(y)}{R['hum']})")
        row("chk", "檢查：分部合計 − 合併營收（應為 0）", "百萬元",
            lambda y: None if y == 2023 else f"={C(y)}{R['cool']}+{C(y)}{R['auto']}+{C(y)}{R['hum']}-{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['cool']}+{C(y)}{R['auto']}+{C(y)}{R['hum']}-{C(y)}{R['rev']}")
        PL.cell(R["chk"], 1, "檢查：分部合計 − 合併營收（應為 0）" if s == "B" else f"檢查（情境 {s}）：分部合計 − 合併營收（應為 0）")
        row("yoy", "營收年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1",
            lambda y: f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1", "0.0%")
        row("cogs", "營業成本", "百萬元", H(5), lambda y: f"={C(y)}{R['rev']}*(1-Thesis!{TC[y]}{th['gm']})")
        row("gp", "毛利", "百萬元", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}")
        row("gm", "毛利率", "%", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", "0.0%")
        row("ts", "稅金及附加", "百萬元", H(19), lambda y: f"={C(y)}{R['rev']}*ts_r")
        row("opex", "營業費用（銷售＋管理＋研發）", "百萬元", H(16), lambda y: f"={C(y)}{R['rev']}*Thesis!{TC[y]}{th['opex']}")
        row("othop", "其他營業損益（其他收益 − 減值等）", "百萬元", H(22), lambda y: f"={C(y)}{R['rev']}*(oi_r-imp_r)")
        row("op", "營業利益（核心，不含金融資產收益）", "百萬元", H(21), lambda y: f"={C(y)}{R['gp']}-{C(y)}{R['ts']}-{C(y)}{R['opex']}+{C(y)}{R['othop']}")
        row("opm", "營業利益率", "%", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", "0.0%")
        row("nonop", "其他收入淨額（業外：投資收益、公允價值、財務費用、營業外）", "百萬元", H(26), lambda y: "=nonop26" if y == 2026 else "=(cash_v+ltinv_v)*inv_y")
        row("pbt", "稅前利益", "百萬元", H(32), lambda y: f"={C(y)}{R['op']}+{C(y)}{R['nonop']}")
        row("tax", "所得稅", "百萬元", H(33), lambda y: f"={C(y)}{R['pbt']}*tax_r")
        row("nit", "稅後淨利（含少數股東）", "百萬元", H(34), lambda y: f"={C(y)}{R['pbt']}-{C(y)}{R['tax']}")
        row("nci", "少數股東損益", "百萬元", H(35), lambda y: f"={C(y)}{R['nit']}*nci_r")
        row("ni", "歸屬母公司淨利", "百萬元", H(36), lambda y: f"={C(y)}{R['nit']}-{C(y)}{R['nci']}")
        row("nim", "淨利率", "%", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", "0.0%")
        row("niy", "淨利年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1",
            lambda y: f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1", "0.0%")
        row("sh", "股數（百萬股）", "百萬股", H(39), lambda y: "=sh_v/(1-h_iss)")
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
        R["hcash"] = r; DC.cell(r, 1, "＋H 股募資（尾端事件；基準 0）"); DC[f"B{r}"] = "=sh_v*h_iss/(1-h_iss)*px*(1-h_disc)"; r += 1
        R["nc"] = r; DC.cell(r, 1, "淨現金滾到評價日（現金與投資 − 有息負債 − 少數股東權益 ＋ Stub FCFF ＋ H 股募資）")
        DC[f"B{r}"] = f"=cash_v+ltinv_v*inv_k-debt_v-nci_v+B{R['stub']}+B{R['hcash']}"; r += 1
        R["eq"] = r; DC.cell(r, 1, "股權價值"); DC[f"B{r}"] = f"=B{R['ev']}+B{R['nc']}"; r += 1
        R["ps"] = r; DC.cell(r, 1, "每股價值（評價日，元）"); DC[f"B{r}"] = f"=MAX(0,B{R['eq']}/(sh_v/(1-h_iss)))"; r += 1
        R["dps"] = r; DC.cell(r, 1, "評價日→目標時點間股利（FY2026 盈餘配息，約 2027-06 支付）"); DC[f"B{r}"] = f"=PL!F{PR[s]['eps']}*payout"; r += 1
        R["v12"] = r; DC.cell(r, 1, "推到目標時點（×（1＋ke）^t − 期間股利）"); DC[f"B{r}"] = f"=MAX(0,B{R['ps']}*(1+ke)^((tp_date-val_date)/365)-B{R['dps']})"; r += 2
        DR[s] = R
    for row_ in DC.iter_rows(min_row=4, max_row=r, min_col=2, max_col=7):
        for c in row_:
            if c.number_format == "General": c.number_format = "#,##0.00"
    DC.column_dimensions["A"].width = 70

    # ================= Comp =================
    CP["A1"] = ("Comp：同業前瞻 P/E。股價＝" + S_CLOSE + "（Verified）；前瞻 P/E 由原報價日按 2026-11-20 收盤等比例換算（Derived，EPS 分母不變）。"
                "核心同業 4 家進中位數（少於 8 家不做迴歸）；EPS 分母口徑不一（資料商 NTM 共識，或券商 2027E），見 O 欄。報價日未取得前瞻 EPS 者另列。")
    hdr(CP, 4, ["公司", "代號", "股價（2026-11-20）", "報價日", "前瞻 P/E 分母", "前瞻 P/E", "近 12 月 P/E（原報價日）", "5 年 β", "會計口徑", "進中位數", "說明", "原報價", "原報價日", "原前瞻 P/E", "標記與換算"])
    PEERS = [
        ("美的集團", "000333.SZ", 94.82, 94.82, "2026-11-20（EPS 反推）", round(94.82 / 6.57, 4), None, None, "中國會計準則", "是", "家電與空調（本公司制冷板塊的下游客戶之一）；分母為券商 2027E EPS 6.57 元（2026-04，非三季報後）", "招商證券 2027E EPS 6.57 元"),
        ("匯川技術", "300124.SZ", 52.66, 52.66, "2026-11-20（EPS 反推）", round(52.66 / 2.61, 4), None, None, "中國會計準則", "是", "工控與電動車電驅（汽車零部件、機器人相鄰）；分母為券商 2027E EPS 2.61 元（2026-08-25）", "華興證券 2027E EPS 2.61 元"),
        ("拓普集團", "601689.SS", 48.11, 46.52, "2026-09-30", 26.89, 32.10, 0.80, "中國會計準則", "是", "汽車零件、人形致動器 Tier 1（直接競爭）；" + S_LD, "資料商 NTM 共識 EPS"),
        ("雙環傳動", "002472.SZ", 35.45, 41.21, "2026-10-30", 22.32, 27.11, 0.52, "中國會計準則", "是", "汽車齒輪與精密傳動；" + S_LD, "資料商 NTM 共識 EPS"),
        ("兆威機電", "003021.SZ", 76.83, 70.25, "2026-09-30", 59.70, 87.15, 0.44, "中國會計準則", "否", "微型傳動、靈巧手；倍數被機器人預期主導（判斷，Assumed），不進中位數；" + S_LD, "資料商 NTM 共識 EPS"),
        ("鳴志電器", "603728.SS", 44.52, 51.50, "2026-09-30", 124.10, 302.24, 0.62, "中國會計準則", "否", "步進與空心杯馬達；倍數過高，不進中位數；" + S_LD, "資料商 NTM 共識 EPS"),
        ("三花智控（本公司，對照）", "002050.SZ", 31.40, 33.30, "2026-10-28", 33.01, None, None, "中國會計準則", "否", "本公司；P/E 33.0 為太平洋證券 2026E EPS 1.01 元於報告日之倍數（對應股價約 33.3 元，反推）", "券商 2026E EPS 1.01 元"),
    ]
    for i, (nm, tk, px1, px0, d0, pe0, ttm, bt, acc, med, note, den) in enumerate(PEERS):
        rr = 5 + i
        CP.cell(rr, 1, nm); CP.cell(rr, 2, tk)
        CP.cell(rr, 3, px1).font = BLUE; CP.cell(rr, 4, "2026-11-20"); CP.cell(rr, 5, den)
        CP.cell(rr, 6, f"=N{rr}*C{rr}/L{rr}")
        if ttm is not None: CP.cell(rr, 7, ttm).font = BLUE
        if bt is not None: CP.cell(rr, 8, bt).font = BLUE
        CP.cell(rr, 9, acc); CP.cell(rr, 10, med); CP.cell(rr, 11, note)
        CP.cell(rr, 12, px0).font = BLUE; CP.cell(rr, 13, d0)
        CP.cell(rr, 14, pe0).font = BLUE
        CP.cell(rr, 15, "股價 Verified（" + S_CLOSE + "）；前瞻 P/E Derived：＝原值 × 2026-11-20 收盤 ÷ 原報價（分母 EPS 不變）；原值來源見說明欄")
    CP.cell(13, 1, "報價日未取得（不納入中位數）：海爾智家、格力電器、盾安環境、銀輪股份（收盤檔無或前瞻 EPS 未取得）；日本同業與本公司業務無重疊，不列。")
    CP.cell(15, 1, "核心同業前瞻 P/E 中位數"); CP["F15"] = "=MEDIAN(F5:F8)"
    CP.cell(16, 1, "本公司現價 ÷ 券商 FY2027E EPS（分母：太平洋證券 1.19 元）"); CP["F16"] = "=px/1.19"
    CP.cell(17, 1, "本公司現價 ÷ 券商 FY2026E EPS（分母：太平洋證券 1.01 元）"); CP["F17"] = "=px/1.01"
    CP.cell(18, 1, "本公司現價 ÷ 模型 FY2027 EPS（分母：情境 B，FY2027）"); CP["F18"] = f"=px/PL!G{PR['B']['eps']}"
    CP.cell(19, 1, "本公司現行倍數 ÷ 同業中位數（券商 2027E 口徑）"); CP["F19"] = "=F16/F15"
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
    RC["A1"] = "Recon：模型對實績、券商預測的同口徑對照；人形致動器市場的隱含值；反解支持現價所需的主導變數。"
    hdr(RC, 3, ["項目", "模型", "對照", "差距（模型 − 對照）", "說明"])
    rr = 4
    def rrow(lab, a, b, note, fmt="#,##0.0", diff=True):
        nonlocal rr
        RC.cell(rr, 1, lab); RC.cell(rr, 2, a).number_format = fmt
        if b is not None: RC.cell(rr, 3, b).number_format = fmt
        if diff and b is not None: RC.cell(rr, 4, f"=B{rr}-C{rr}").number_format = fmt
        RC.cell(rr, 5, short(note)); rr += 1
    rrow("FY2026 歸屬母公司淨利（情境 B）vs 太平洋證券 2026E（42.37 億元）", f"=PL!F{PR['B']['ni']}", 4237, "券商預測為 Interested-party；2026 年前三季淨利 30.67 億元（−5.4%）")
    rrow("FY2027 歸屬母公司淨利（情境 B）vs 太平洋證券 2027E（50.11 億元）", f"=PL!G{PR['B']['ni']}", 5011, "")
    rrow("FY2028 歸屬母公司淨利（情境 B）vs 太平洋證券 2028E（60.46 億元）", f"=PL!H{PR['B']['ni']}", 6046, "")
    rrow("FY2026 EPS（情境 B）vs 太平洋證券 2026E（1.01 元）", f"=PL!F{PR['B']['eps']}", 1.01, "", "0.00")
    rrow("FY2027 EPS（情境 B）vs 太平洋證券 2027E（1.19 元）", f"=PL!G{PR['B']['eps']}", 1.19, "", "0.00")
    rrow("FY2026 毛利率 vs 第 3 季實績代表值 gm26", f"=PL!F{PR['B']['gm']}", "=gm26", "模型以第 3 季為錨（差距應為 0）", "0.0%")
    rrow("FY2026 第 4 季隱含營收 vs 2026Q3 實績（8,374 百萬元）", f"=PL!F{PR['B']['rev']}-r9m26", 8374, "第 4 季通常低於第 3 季（2025：6,982 vs 7,767）")
    rrow("FY2026 營收（三情境同）vs 2025 年", f"=PL!F{PR['B']['rev']}", "=Hist!E4", "年增見 PL", "#,##0")
    rr += 1
    RC.cell(rr, 1, "人形致動器市場隱含值（情境 B；百萬美元）").font = BOLD; rr += 1
    rrow("2030 全球人形致動器市場（情境 B）", f"=Thesis!I{THR['B']['mkt']}", None, "＝線 4 出貨 × 全尺寸等效係數 × 每台 BOM × 致動器占比", "#,##0", False)
    rrow("2030 本公司人形致動器營收（情境 B，百萬美元）", f"=PL!J{PR['B']['hum']}/fx", None, "＝市場 × 份額", "#,##0", False)
    rrow("FY2030 人形致動器占合併營收（情境 B）", f"=Thesis!I{THR['B']['humsh']}", None, "人形不是 2030 年前的主要營收來源（仍以制冷與汽零為主）", "0.0%", False)
    rrow("2026 每台致動器價值（美元，全尺寸）vs 線 6「人形每台零組件內容 2026」", f"=Thesis!E{THR['B']['bom']}", LD["l6c"][0], f"{LD['l6c'][1]}（口徑較窄：不含運算、電池、結構、組裝）；本模型用線 4 致動器占 BOM（含模組整合），見對帳", "#,##0")
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
        ("全球人形機器人出貨 FY2026–FY2032（A／B／C＝線高／基準／低）", "4 人形", f"基準 {h26[0]:,.0f} → {h32[0]:,.0f}（× hu_k）", "台", "全球、含小型機", f"{h26[1]}；{h32[1]}", "Assumed", "Thesis 人形致動器市場", "一致（hu_k＝1）；高案 hu_k＝2 為判斷，列 Sens 不改基準"),
        ("全尺寸人形每台 BOM（2026／2028／2030／2035）", "4 人形", "基準 " + " ／ ".join(f"{LD['bom'][(y, 'B')][0]:,.0f}" for y in (2026, 2028, 2030, 2035)), "美元／台", "全尺寸、硬體", f"{LD['bom'][(2026, 'B')][1]}；{LD['bom'][(2028, 'B')][1]}", "Assumed", "Thesis 每台致動器價值", "一致（直接讀值；中間年線性內插）"),
        ("致動器占全尺寸 BOM", "4 人形", "=act_sh", "%", "旋轉＋線性關節", LD["act"][1], "Analogy", "Thesis 每台致動器價值", "一致（區間取線 4 低／基準／高）"),
        ("全尺寸等效係數（全球平均每台諧波用量 ÷ 全尺寸每台用量）", "4 人形", "=fs_eq", "x", "出貨組合折算", LD["eq"][1], "Analogy", "Thesis 人形致動器市場", "本模型自設折算（以諧波用量比代表致動器價值比）；線 4 未直接給致動器價值的組合折算"),
        ("人形每台零組件內容 2026（基準）", "6 共用零組件", f"{LD['l6c'][0]:,.0f}", "美元／台", "不含運算、電池、結構、組裝", LD["l6c"][1], "Derived", "Recon 對照（不進公式）", "不一致（刻意）：線 6 為零組件（諧波、絲杠、馬達、感測器等）內容；本模型的致動器含模組整合與組裝，口徑較寬（本模型 2026 每台約 3.9 萬美元 vs 線 6 約 2.3 萬美元）。致動器模組廠是否能獲得整個差額未知，故以份額（Assumed 0–15%）吸收"),
        ("致動器占 BOM 與零組件內容的關係", "4 人形／6 共用零組件", "—", "—", "—", "—", "判斷", "—", "兩線之間的口徑差見上一列；若改用線 6 口徑，人形致動器營收約降 40%，對加權目標價影響見 Sens（hs30 情境）"),
    ]
    for i, z in enumerate(ZR):
        for j, v in enumerate(z): DZ.cell(4 + i, 1 + j, v)
    DZ.column_dimensions["A"].width = 50; DZ.column_dimensions["F"].width = 80; DZ.column_dimensions["I"].width = 70

    # ================= Risk =================
    RK["A1"] = "Risk：風險登錄（原則第 10 節第 10、19 條）。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    RISKS = [
        ("匯兌損失與衍生工具浮虧（2026 第 3 季財務費用率 0.71%，+0.93pt）", "部分（inv_y 區間）", "第 3 季淨利年減 9.6%，營業面仍成長", "財務費用、公允價值變動季變化"),
        ("人形致動器未產生收入、客戶集中於單一整機廠", "是（情境 C 份額 0、情境 A 15%）", "媒體報導未經本專案檢核；整機廠可能自製或直接向減速機、絲杠廠採購", "年報是否單獨揭露機電執行器營收；Tesla 量產節奏"),
        ("制冷板塊依賴家電與資料中心液冷，景氣與關稅", "是（gcool 三情境）", "液冷規模未揭露", "第 3 季制冷增速、液冷揭露"),
        ("電動車熱管理需求放緩與價格競爭", "是（gauto 三情境）", "", "汽零毛利率（2025 年 28.79%）"),
        ("擴產（資本支出約 30 億元／年）與折舊上升", "部分（cx26、cx_r）", "2025 年汽零折舊 +48.6%", "在建工程、毛利率"),
        ("同業倍數口徑不一（NTM 與券商 2027E 混用）", "部分（pe_peer 低／高）", "只有 4 家核心同業", "同日重抓同口徑 EPS"),
        ("再融資（A／H 股）", "部分（h_iss 尾端 Sens，基準 0）", "2025 年已發行 H 股，無新公告", "公告"),
        ("客戶集中度", "否", "財報揭露未找到（年報前五大客戶占比未查）", "年報前五大客戶"),
    ]
    for i, k in enumerate(RISKS):
        for j, v in enumerate(k): RK.cell(4 + i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["D"].width = 60

    # ================= Sources =================
    SO["A1"] = "Sources：本模型引用的原始來源與資料商（含網址；HTML 展示層不放網址）。"
    for i, s_ in enumerate([S_PX, S_AR, S_Q3, S_23, S_BR, S_MD, S_IN, S_LD, S_RF, S_CRP,
                            f"線 4：{L4N}", f"線 6：{L6N}",
                            "physical_ai/map/profiles/002050.md（背景，數字已另行查證）"], 3):
        SO.cell(i, 1, s_)
    SO.column_dimensions["A"].width = 160

    # ================= README（xlsx 內）=================
    S = T
    v = lambda k: f"Val!B{VR[k]}"
    lines = [
        ('="三花智控（002050.SZ）個股模型 ' + VER + '｜build_id "&Sens!D2&"｜資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'),
        "命題：三花智控的營收與利潤主要來自制冷空調電器零部件（含資料中心液冷）與汽車熱管理零部件，兩者合計近 100%；仿生機器人機電執行器目前沒有收入，只是選擇權。營收＝制冷（年增）＋汽零（年增）＋人形致動器（線 4 出貨 × 每台致動器價值 × 份額）。",
        (f'="結論（條件式）：在 FY2026 以 1–9 月實績為錨、制冷年增 "&{S("gcool_B","0%")}&"、汽零年增 "&{S("gauto_B","0%")}&"、情境 B 的 2030 年人形致動器份額 "&{S("hs30_B","0%")}&"、同業 P/E 未加溢價（相對倍數係數 "&{S("rel_k","0.00")}&"）的假設下，加權目標價約 "&{S(v("tp"),"#,##0.0")}&" 元，相對現價 "&{S("px","#,##0.00")}&" 元 "&{S(v("up"),"+0.0%;-0.0%")}&"，估值評等 "&{v("rating")}&"，投資建議 "&{v("rec")}&"。"'),
        (f'="DCF 目標價 "&{S(v("dcf_tp"),"#,##0.0")}&"、Comparable 目標價 "&{S(v("cmp_tp"),"#,##0.0")}&"（各半）；WACC "&{S("wacc","0.00%")}&"、ke "&{S("ke","0.00%")}&"、g "&{S("g_tv","0.00%")}&"；情境 B 終值占 EV "&{S(f"Val!C{VR['tvs']}","0%")}&"。"'),
        (f'="上檔空間歸因：DCF 貢獻 "&{S(v("att_d"),"+0.0%;-0.0%")}&"、未調整 Comparable 貢獻 "&{S(v("att_c"),"+0.0%;-0.0%")}&"、相對倍數調整貢獻 "&{S(v("att_r"),"+0.0%;-0.0%")}&"。本公司現價對券商 2027E EPS 的倍數是同業中位數的 "&{S(v("relnow"),"0.00")}&" 倍。"'),
        (f'="收支關卡：情境 B FY2026–FY2030 FCFF 扣股利合計 "&{S(v("gate_b"),"#,##0.0")}&" 百萬元，"&{v("gate_t")}&"；距最近級距邊界 "&{S(v("dist"),"0.0%")}&"。"'),
        "主要失效條件（可觀察）：(1) 制冷與汽零成長明顯偏離本模型區間（季報增速）；(2) 毛利率脫離 27.5–30.5% 區間；(3) 機電執行器開始單獨揭露營收且量級遠高於情境 A；(4) 匯兌與衍生工具損失持續擴大。",
        "方法：DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，WACC 折現，明確期到 FY2032，RONIC＝WACC；金融資產以淨現金計入）50%＋同業前瞻 P/E × 模型 FY2027 EPS 50%；三情境 25／50／25。",
        "限制：人形致動器份額為 Assumed（公司未揭露任何金額）；β 未取得資料商值（Assumed 0.9）；2026 三季報資產負債表未取得（用 2025 年報）；同業 EPS 口徑不一；不讀 Siliconomics／Tokenomics／Wattonomics。詳見 repo README.md。",
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
        ("情境 B 的 2030 年人形致動器份額改為情境 C 值（0）", "hs30_B", ivE["hs30_C"]),
        ("情境 B 制冷年增改為情境 A 值", "gcool_B", ivE["gcool_A"]),
        ("情境 B 制冷年增改為情境 C 值", "gcool_B", ivE["gcool_C"]),
        ("情境 B 汽零年增改為情境 A 值", "gauto_B", ivE["gauto_A"]),
        ("情境 B 汽零年增改為情境 C 值", "gauto_B", ivE["gauto_C"]),
        ("情境 B 的 FY2030 毛利率改為情境 A 值", "gm30_B", ivE["gm30_A"]),
        ("情境 B 的 FY2030 毛利率改為情境 C 值", "gm30_B", ivE["gm30_C"]),
        ("人形出貨為線 4 的 5 倍（全部情境）", "hu_k", 5.0),
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
    s1 = solve("hu_k", px, 1.0, 400.0)
    s2 = solve("hs30_B", px, 0.001, 1.0)
    s3 = solve("gm30_B", px, 0.2, 0.6)
    s4 = solve("gm30_B", flip_tp, 0.2, 0.6)
    nf = lambda x, txt: x if x else -1
    solve_rows = [
        ("支持現價（目標價＝現價）所需的人形出貨倍數（相對線 4，全部情境）", nf(s1, ""), 1.0, "其他輸入不變；−1＝400 倍內無解（或現價已低於基準目標價方向不同，見 README）"),
        ("支持現價所需的情境 B 2030 年人形致動器份額", nf(s2, ""), ivE["hs30_B"], "其他輸入不變；−1＝份額 100% 仍無解"),
        ("支持現價所需的情境 B FY2030 毛利率", nf(s3, ""), ivE["gm30_B"], "其他輸入不變；−1＝60% 內無解；若解低於 0.2 下限則回報下限"),
        (f"評等翻轉（目標價達 {flip_tp:,.2f}）所需的情境 B FY2030 毛利率", nf(s4, ""), ivE["gm30_B"], "其他輸入不變；−1＝無解"),
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
    segs = [("制冷空調電器零部件", "制冷空調電器零部件（含資料中心液冷） 營收", Hs["E8"].value), ("汽車零部件", "汽車零部件 營收", Hs["E9"].value)]
    json.dump(dict(company="sanhua", ticker="002050.SZ", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=round(v, 2), unit="百萬元", source=S_AR) for n, l, v in segs],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）"],
                   note="官方分產品＝年報主營業務分產品（制冷空調電器零部件、汽車零部件）；仿生機器人機電執行器年報未單獨揭露（FY2025 以 0 計），模型另列一行，未計入官方分部。"),
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
    t = f"""# physical_ai/sanhua/ — 三花智控（002050.SZ）個股模型 {VER}

<!-- 本檔由 build_sanhua.py 依成品重算值產生；數字不手寫（原則第 10 節第 51 條）。 -->

**本模型適用《高不確定性預估工作原則》**（`docs/plan/高不確定性預估工作原則.md`）：人形致動器尚無收入、公司未揭露金額，相關輸入一律給區間並標 Assumed／Analogy；結論只在證據邊界內成立。

**未完成項目（誠實揭露）：** (1) 2026 三季報資產負債表與分板塊 1–9 月營收未取得（資產負債表用 2025 年報，分板塊用第 3 季比重）；(2) 本公司 β 未取得（Assumed 0.9）；(3) 2023 年分部、稅金及附加、投資收益等為估計或以 0 計；(4) 同業僅 4 家進中位數，EPS 口徑不一；(5) 未讀取三季報原文，三季報數字取自券商與媒體轉錄；(6) 未找到市場共識目標價；(7) 獨立核對尚待主流程安排。

- 成品：`dist/{NAME}`（build_id `{sens['build_id']}`；LibreOffice 重算，0 錯誤值）；HTML 展示層 `dist/{NAME[:-5]}.html`（`gen_html_json.py` 產生 `html.json`，再以 `tools/company_html/build_company_html.py` 產生）
- 建置：`python3 physical_ai/sanhua/build_sanhua.py`（需要 soffice；整本重算敏感度與反解；結尾自動跑各項核對）
- 日期：資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 2026-11-20（收盤 {I['px']:,.2f} 元）｜評價日 2026-12-31｜目標時點 2027-11-20
- 會計年度＝日曆年；單位百萬元人民幣；中國會計準則。深交所 A 股（另有港股 2050）。

## 結論（條件式）

在「FY2026 以 1–9 月實績為錨、制冷年增 {I['gcool_B']:.0%}、汽零年增 {I['gauto_B']:.0%}、情境 B 的 2030 年人形致動器份額 {I['hs30_B']:.0%}、同業 P/E 不加溢價（相對倍數係數 {I['rel_k']:.2f}）」的假設下：

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

上檔空間歸因：DCF 貢獻 {g('att_d'):+.1%}、未調整 Comparable 貢獻 {g('att_c'):+.1%}、相對倍數調整 {g('att_r'):+.1%}。核心同業前瞻 P/E 中位數 {I['pe_peer']:.1f} 倍；本公司現價對券商 2027E EPS 的倍數是同業中位數的 {g('relnow'):.2f} 倍。

**讀法：** 這家公司的價值幾乎全在制冷與汽車熱管理兩個既有板塊；情境 B 人形致動器營收 {ser('hum', '{:,.0f}')}（FY2026→FY2032，百萬元），占合併營收只從 {TH[f'E{THB["humsh"]}'].value:.0%} 升到 {TH[f'K{THB["humsh"]}'].value:.1%}，因此人形假設對目標價影響有限（見敏感度）。結論主要建立在 Assumed 的板塊成長率、毛利率與 Derived 的同業倍數上；與市場價格的差距主要來自折現率與同業倍數的選擇。

## 命題

1. **營收結構：** 官方分產品為制冷空調電器零部件（2025 年 {wb['Hist']['E8'].value:,.0f} 百萬元，占 {wb['Hist']['E8'].value / wb['Hist']['E4'].value:.0%}，含資料中心液冷）與汽車零部件（{wb['Hist']['E9'].value:,.0f}）；機電執行器未單獨揭露，本模型另列一行（FY2026＝0）。
2. **量與價：** 兩個既有板塊以年增率推導（Assumed，參考第 3 季制冷 +6.5%、汽零 +9.6%）；人形致動器＝線 4 出貨 × 全尺寸等效係數 {I['fs_eq']:.2f} × 每台 BOM × 致動器占比 {I['act_sh']:.1%} × 份額 × 匯率。
3. **利潤：** 毛利率由第 3 季 {I['gm26']:.2%} 線性走到 FY2030 {I['gm30_B']:.1%}（情境 B）；費用率 {I['opex26']:.1%} → {I['opex30']:.1%}。

情境 B：營收 {ser('rev', '{:,.0f}')}；毛利率 {ser('gm', '{:.1%}')}；EPS {ser('eps', '{:.2f}')} 元（FY2026→FY2032）。

## 主要假設與可證偽指標

| 承重假設 | 基準（A／B／C 或區間） | 標記 | 可觀察指標 |
|---|---|---|---|
| 制冷（含液冷）年增 | {I['gcool_A']:.0%}／{I['gcool_B']:.0%}／{I['gcool_C']:.0%} | Assumed | 季報分板塊營收 |
| 汽車零部件年增 | {I['gauto_A']:.0%}／{I['gauto_B']:.0%}／{I['gauto_C']:.0%} | Assumed | 季報分板塊營收 |
| FY2030 毛利率 | {I['gm30_A']:.1%}／{I['gm30_B']:.1%}／{I['gm30_C']:.1%} | Assumed | 季毛利率（第 3 季 28.89%） |
| 2030 年人形致動器份額 | {I['hs30_A']:.0%}／{I['hs30_B']:.0%}／{I['hs30_C']:.0%} | Assumed | 年報是否揭露機電執行器營收 |
| 全球人形出貨（線 4） | 線 4 高／基準／低 | Assumed（線 4） | IDC、TrendForce 季出貨 |
| 同業前瞻 P/E | 中位數 {I['pe_peer']:.1f} 倍（4 家） | Derived | 同日重抓同口徑 EPS |

## 敏感度（整本重算，前 10 大）

| 案例 | 鍵 | 加權目標價 | 相對基準 | DCF 目標價 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | `{c['key']}` | {c['tp']:,.1f} | {c['tp'] / sens['base_tp'] - 1:+.1%} | {c['dcf']:,.1f} | {c['rating']} |" for c in top) + f"""

完整案例見 xlsx Sens 頁。**翻轉評等所需的變動**：見 Recon 反解；升一級邊界目標價 {fn(g('up_b'))}、降一級邊界 {fn(g('dn_b'))} 元。

## 反解（Recon 頁）

""" + "\n".join(f"- {a}：{fn(b, '{:,.2f}')}（基準 {fn(c0, '{:,.2f}')}）；{d}" for a, b, c0, d in sol) + f"""

## 限制與「分不出來／沒找到」

- **分不出來：** 人形致動器金額公司未揭露；液冷規模未揭露，併入制冷成長率。
- **沒找到：** 2026 三季報資產負債表與現金流、1–9 月分板塊營收；本公司 β；市場共識目標價；客戶集中度（年報前五大客戶占比未查）；2023 年分部營收。
- **線模型不一致：** 線 6「人形每台零組件內容」（約 2.3 萬美元）口徑窄於本模型致動器價值（線 4 BOM × 致動器占比，約 3.9 萬美元）；本模型用線 4，差額以份額吸收。見 xlsx「對帳」頁。
- **同業倍數：** 核心 4 家（美的、匯川、拓普、雙環）；美的、匯川以券商 2027E EPS 反推，拓普、雙環用資料商 NTM 共識由原報價日換算至 2026-11-20；兆威、鳴志倍數被機器人預期主導，不進中位數。海爾智家、格力電器、盾安環境、銀輪股份報價日未取得前瞻 EPS。
- **rf 逾期：** 中國 10 年期國債 1.61% 為 2026-11-05 值（跨模型一致）。
- **協調檔：** 實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics。

## 資料來源

見 xlsx Sources 頁（公司 2025 年年報原文、券商三季報點評、資料商、線 4／6 成品）。

## 版本紀錄

- {VER}（2026-11-24）：首版。現價 {I['px']:,.2f}（2026-11-20 收盤）；加權目標價 {g('tp'):,.1f}；上檔空間 {g('up'):+.1%}；{g('rating')}／{g('rec')}；收支關卡 {g('gate_b'):,.1f}（{g('gate_t')}）。
"""
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write(t)


if __name__ == "__main__":
    main()

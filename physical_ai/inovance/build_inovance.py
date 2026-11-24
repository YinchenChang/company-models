# 匯川技術 Inovance（300124.SZ，深交所創業板）個股模型 v0.1 建置腳本
# 產出：physical_ai/inovance/dist/20261124_匯川技術個股模型_v0.1.xlsx（LibreOffice 重算後的成品，含公式與快取值）
# 結構：命題模組（Thesis：通用自動化＝基礎＋人形伺服增量；智能機器人＝中國裝機（線 3）× 份額漂移 × 單價；新能源汽車、電梯、數字能源以情境成長路徑）
#       → 官方分部 → 標準損益表 → FCFF DCF＋同業 P/E。
# 敏感度：每個案例覆寫 In 頁基準欄後整本重建、LibreOffice 重算（整本重算，不用閉式近似）；反解同法以二分搜尋。
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md、docs/plan/高不確定性預估工作原則.md。
# 用法：python3 physical_ai/inovance/build_inovance.py   （需要 soffice；暫存於 $HOME/ic_build）
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
NAME = f"20261124_匯川技術個股模型_{VER}.xlsx"
OUT = os.path.join(HERE, "dist", NAME)
HOMEDIR = os.path.expanduser("~")
TMP = os.path.join(HOMEDIR, "ic_build")
LO = ["soffice", f"-env:UserInstallation=file://{HOMEDIR}/lo_profile", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML"]

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
S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv：300124.SZ；未還原股利）"
S_SAF = "stockanalysis.com SHE:300124 財務頁（資料商；2026-11-24 讀取；成本、稅前、所得稅、股數為頁面比率推得，取整）"
S_SAS = "stockanalysis.com SHE:300124 statistics 頁（資料商；頁面報價 2026-11-24 收盤 53.61 元，2026-11-24 讀取）"
S_SAB = "stockanalysis.com SHE:300124 季度資產負債表（資料商；2026-09-30；2026-11-24 讀取）"
S_SAC = "stockanalysis.com SHE:300124 現金流量表（資料商；2026-11-24 讀取）"
S_IF = "鳳凰財經 2025 年報點評（https://finance.ifeng.com/c/8udndDyY7Iz）：轉錄公司 2025 年報分業務營收、毛利率、研發費用"
S_DF1 = "東方財富研報 2026 Q3 點評（https://data.eastmoney.com/report/info/AP202610271831007496.html）：轉錄公司 2026 三季報分業務營收與盈利預測"
S_DF2 = "東方財富研報（2026-04-28；https://data.eastmoney.com/report/info/AP202604281821656918.html）：2025 年分業務、2026 年新能源車營收 230–240 億元展望、盈利預測"
S_SW = "申萬宏源 2026 Q3 點評（2026-11-05；https://m.gelonghui.com/news/5345524）：三季報分業務營收、伺服／變頻器／PLC 拆分、盈利預測"
S_10 = "同花順（每日經濟新聞）匯川三季報（2026-10-26；https://news.10jqka.com.cn/20261026/c681099525.shtml）：營收、淨利、毛利率、第 3 季公允價值收益約 4.2 億元"
S_PF = "physical_ai/map/profiles/300124.md（背景；數字已另行查證）"
S_RF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與其他中國模型同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：China CRP 0.91%（與其他中國模型同一取值）"
S_CLOSE = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
IP_BRK = "券商研究報告：有銷售研究與維持評等的誘因；歷史數字為轉錄，預測為券商自設"
IP_SA = "資料商彙整賣方共識，財報前的值未必更新"
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


def line_data():
    _, l3 = _line_rows(L3, 1)
    wb4, l4 = _line_rows(L4, 1)
    wb6, _ = _line_rows(L6, 1)
    d = {"cn": {}, "hum": {}}
    for y in range(2025, 2033):
        for s, tag in (("A", "hi"), ("B", "base"), ("C", "lo")):
            k = f"IR.inst.china.{y}.{tag}" if y > 2025 else "IR.inst.china.2025.base"
            r, v = l3[k]
            d["cn"][(y, s)] = (v[4], f"{L3N} 下游取數!E{r}（{k}）")
    for y in range(2026, 2033):
        r, v = l4[f"SHIP_{y}"]
        for s, c, col in (("A", 4, "E"), ("B", 3, "D"), ("C", 2, "C")):
            d["hum"][(y, s)] = (v[c], f"{L4N} 下游取數!{col}{r}（SHIP_{y}）")
    ws6 = wb6["輸入"]
    sd = [r for r in range(1, ws6.max_row + 1) if ws6.cell(r, 1).value == "SD"]   # 低、基準、高矩陣各一列，第 4 列為單價
    d["ax"] = [(ws6.cell(r, 6).value, f"{L6N} 輸入!F{r}") for r in sd[:3]]         # 人形每台伺服驅動軸數（低、基準、高）
    pr = sd[3]
    d["psd"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{openpyxl.utils.get_column_letter(c)}{pr}") for c in (4, 5, 6)]
    d["pdd"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{openpyxl.utils.get_column_letter(c)}{pr}") for c in (7, 8, 9)]
    r, v = l3["IR.share.china_domestic.2025"]; d["dom"] = (v[4], f"{L3N} 下游取數!E{r}")
    r, v = l3["IR.cycle.cn_prod.ytd_yoy"]; d["prod"] = (v[4], f"{L3N} 下游取數!E{r}")
    r, v = l3["IR.payback.cn6ax"]; d["pay"] = (v[4], f"{L3N} 下游取數!E{r}")
    return d


LD = line_data()

# ---------------- In 頁輸入 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 來源誘因, 說明)
IN = [
    ("sec", "1. 市場資料與日期"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, "", "資料截止日 2026-11-24（財報、券商預測、同業倍數）；股價 2026-11-20；成品日 2026-11-24"),
    ("px", "現價（2026-11-20 收盤）", "元", None, 52.66, None, "Verified", S_PX, "", "csv 值 52.659999；11-09 收盤 57.55，兩週跌 8.5%"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,12,31)", None, "Assumed", "本專案慣例：評價日 2026-12-31（會計年度＝日曆年，FY2026 在評價日已結束）", "", "折現期間由評價日起算"),
    ("tp_date", "目標時點（px_date 後 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", ""),
    ("bs_date", "資產負債表日", "日期", None, "=DATE(2026,9,30)", None, "Verified", S_SAB, "", "2026 年三季報"),
    ("stub_frac", "Stub 期（資產負債表日→評價日）", "年", None, "=(val_date-bs_date)/365", None, "Derived", "＝(評價日 − 資產負債表日) ÷ 365", "", "Stub FCFF＝FY2026 FCFF × stub_frac"),
    ("sh_v", "評價日股數（總股本）", "百萬股", None, 2709.0, None, "Verified", S_SAS + "：市值 1,452.2 億元 ÷ 53.61 元＝27.09 億股（頁面另列 27.1 億股）", "", "庫藏股、限制性股票未取得，以 0 計；H 股發行另見 h_iss"),
    ("cash_v", "現金＋短期投資（2026-09-30）", "百萬元", None, 10089.0, None, "Verified", S_SAB + "：現金 6,118＋短期投資 3,971", "", "資料商另列現金合計 12.98 十億元（口徑未明，未採用）"),
    ("ltinv_v", "長期投資（2026-09-30；含股權投資與基金）", "百萬元", None, 5810.0, None, "Verified", S_SAB, "", "明細未取得；含海外基金與聯營投資（2026 第 3 季公允價值收益約 4.2 億元）"),
    ("inv_k", "長期投資計入淨現金的比例", "%", 0.5, 1.0, None, "Assumed", "基準全額計入；低案 50%（股權投資流動性與折價）", "", ""),
    ("debt_v", "有息負債（2026-09-30）", "百萬元", None, 2972.0, None, "Verified", S_SAB + "：短期借款 550＋一年內到期長期負債 1,343＋長期借款 1,078", "", ""),
    ("nci_v", "少數股東權益（2026-09-30）", "百萬元", None, 1906.0, None, "Verified", S_SAB, "", "自淨現金扣除"),
    ("fx", "人民幣兌美元（2026-11-20）", "元／美元", None, 6.6665, None, "Verified", S_CLOSE + "：CNY=X", "", "人形伺服單價（美元）換算用；線 6 內部換算用 7.1，兩者不同，見對帳"),
    ("cons_tp", "賣方平均目標價（23 家，2026-11-24 讀取）", "元", None, 70.34, None, "Interested-party", S_SAS, IP_SA, "只作對照，不進模型"),
    ("sec", "2. 實績錨（FY2026＝2026 年，會計年度＝日曆年）"),
    ("q4g_ga", "通用自動化 2026 年第 4 季營收年增", "%", 0.10, 0.25, 0.35, "Assumed", "1–9 月 +33%（" + S_SW + "）；公司展望工控全年 25% 以上（" + S_DF2 + "）", "", "FY2026＝1–9 月實績＋2025Q4 ×（1＋年增）"),
    ("q4g_el", "智慧電梯 2026 年第 4 季營收年增", "%", -0.05, 0.03, 0.08, "Assumed", "1–9 月 +6%；公司展望 2026 全年 0–5%（" + S_DF2 + "）", "", ""),
    ("q4g_nev", "新能源汽車動力系統 2026 年第 4 季營收年增", "%", 0.10, 0.25, 0.35, "Assumed", "1–9 月 +8%；公司展望 2026 全年 230–240 億元（" + S_DF2 + "），基準值對應全年約 230 億元", "", ""),
    ("q4g_em", "新興業務（機器人＋數字能源）2026 年第 4 季營收年增", "%", 0.50, 0.80, 1.00, "Assumed", "1–9 月 +88%、上半年 +96%（" + S_SW + "）", "", ""),
    ("q4g_ot", "其他營收 2026 年第 4 季年增", "%", -0.20, 0.0, 0.20, "Assumed", "未揭露；殘差項目", "", ""),
    ("gm26", "FY2026 毛利率", "%", 0.265, 0.275, 0.285, "Assumed", "1–9 月 28.2%、第 3 季 25.5%（" + S_10 + "）；原料（銅、電子料）漲價，第 4 季假設接近第 3 季", "", ""),
    ("opex26", "FY2026 營業費用及其他營業損益淨額 ÷ 營收", "%", 0.165, 0.175, 0.185, "Assumed", "歷史（毛利 − 營業利益）÷ 營收：2023 年 17.9%、2024 年 16.1%、2025 年 16.5%（Hist，Derived）；2026 前 9 月資產減值 5.42 億元（上年同期 2.32 億元）", "", "銷管研分項本輪未取得；含稅金、減值、其他收益"),
    ("nonop26", "FY2026 業外淨額", "百萬元", 200, 500, 800, "Assumed", "2025 年稅前 − 營業利益約 178；2026 第 3 季股權投資公允價值收益約 420（" + S_10 + "）", "", ""),
    ("sec", "3. 命題模組：通用自動化、智能機器人（線 3）、人形伺服（線 4、線 6）"),
    ("rob_sh26", "FY2026 新興業務中智能機器人占比", "%", 0.40, 0.60, 0.75, "Assumed", "公司只揭露新興業務合計（機器人＋數字能源）；占比未揭露", "", "其餘為數字能源"),
    ("pdr", "智能機器人每台營收年降幅", "%／年", 0.0, 0.03, 0.06, "Assumed", "國產價格競爭：國產 6 軸約 10 萬元對外資約 27.5 萬元（線 3）；匯川單價未揭露", "", "線 3 無單價年降變數，自設"),
    ("drift_A", "情境 A：智能機器人相對中國裝機的份額年漂移", "%／年", None, 0.25, None, "Assumed", f"2026 前 9 月新興業務 +88%，中國裝機國產占比 {LD['dom'][0]:.0%}（{LD['dom'][1]}）；公司稱中國工業機器人份額 9.2%、SCARA 約 25%（" + S_SW.split("（")[0] + "轉述）", "", "正＝取得份額"),
    ("drift_B", "情境 B：智能機器人相對中國裝機的份額年漂移", "%／年", None, 0.15, None, "Assumed", "同上", "", ""),
    ("drift_C", "情境 C：智能機器人相對中國裝機的份額年漂移", "%／年", None, 0.05, None, "Assumed", "同上", "", ""),
    ("ax", "人形每台伺服驅動軸數（線 6）", "軸／台", LD["ax"][0][0], LD["ax"][1][0], LD["ax"][2][0], "Analogy", f"線 6：{LD['ax'][1][1]}（低 {LD['ax'][0][1]}、高 {LD['ax'][2][1]}）", "", "線 6 以人形矩陣 SD 列"),
    ("psd", "伺服驅動單價 2026（線 6，每軸）", "美元／軸", LD["psd"][0][0], LD["psd"][1][0], LD["psd"][2][0], "Assumed", f"線 6：{LD['psd'][1][1]}（線 6 自標 Assumed，未取得實際單價）", "", ""),
    ("pds", "伺服驅動單價年降幅（線 6）", "%／年", LD["pdd"][0][0], LD["pdd"][1][0], LD["pdd"][2][0], "Assumed", f"線 6：{LD['pdd'][1][1]}", "", ""),
    ("hu_k", "人形出貨相對線 4 的倍數（全部情境）", "x", None, 1.0, 2.0, "Assumed", "基準＝線 4 原值；高案 2 倍（中國人形產量高於線 4 的參考值）", "", "反解變數"),
    ("hs30_A", "情境 A：2030 年匯川占全球人形伺服驅動份額", "%", None, 0.08, None, "Assumed", "匯川為中國伺服龍頭；人形用伺服多為客製、廠內自製比例高，未取得揭露", "", "2026 年以 0 計（已含於通用自動化），增量自 2027 起線性到 2030"),
    ("hs30_B", "情境 B：2030 年匯川占全球人形伺服驅動份額", "%", None, 0.04, None, "Assumed", "同上", "", ""),
    ("hs30_C", "情境 C：2030 年匯川占全球人形伺服驅動份額", "%", None, 0.01, None, "Assumed", "同上", "", ""),
    ("sec", "4. 情境專屬驅動（A 樂觀／B 基準／C 保守；只列基準欄，不設敏感度區間）"),
    ("ga27_A", "情境 A：通用自動化 FY2027 營收年增", "%", None, 0.28, None, "Assumed", "2026 前 9 月 +33%；展望 25% 以上；國產替代與工業復甦", "", "年增由 FY2027 值線性走到 FY2032 值"),
    ("ga27_B", "情境 B：通用自動化 FY2027 營收年增", "%", None, 0.20, None, "Assumed", "同上", "", ""),
    ("ga27_C", "情境 C：通用自動化 FY2027 營收年增", "%", None, 0.10, None, "Assumed", "同上", "", ""),
    ("ga32_A", "情境 A：通用自動化 FY2032 營收年增", "%", None, 0.12, None, "Assumed", "成長減速（基期變大）", "", ""),
    ("ga32_B", "情境 B：通用自動化 FY2032 營收年增", "%", None, 0.08, None, "Assumed", "同上", "", ""),
    ("ga32_C", "情境 C：通用自動化 FY2032 營收年增", "%", None, 0.04, None, "Assumed", "同上", "", ""),
    ("nev27_A", "情境 A：新能源汽車動力系統 FY2027 營收年增", "%", None, 0.15, None, "Assumed", "2025 年 +26%、2026 前 9 月 +8%；國內新能源車需求下滑與價格壓力（" + S_10 + "）", "", ""),
    ("nev27_B", "情境 B：新能源汽車動力系統 FY2027 營收年增", "%", None, 0.08, None, "Assumed", "同上", "", ""),
    ("nev27_C", "情境 C：新能源汽車動力系統 FY2027 營收年增", "%", None, 0.0, None, "Assumed", "同上", "", ""),
    ("nev32_A", "情境 A：新能源汽車動力系統 FY2032 營收年增", "%", None, 0.07, None, "Assumed", "同上", "", ""),
    ("nev32_B", "情境 B：新能源汽車動力系統 FY2032 營收年增", "%", None, 0.04, None, "Assumed", "同上", "", ""),
    ("nev32_C", "情境 C：新能源汽車動力系統 FY2032 營收年增", "%", None, 0.0, None, "Assumed", "同上", "", ""),
    ("gel_A", "情境 A：智慧電梯營收年增（FY2027 起）", "%", None, 0.05, None, "Assumed", "2025 年 +5%、2026 前 9 月 +6%", "", ""),
    ("gel_B", "情境 B：智慧電梯營收年增（FY2027 起）", "%", None, 0.03, None, "Assumed", "同上", "", ""),
    ("gel_C", "情境 C：智慧電梯營收年增（FY2027 起）", "%", None, 0.0, None, "Assumed", "同上", "", ""),
    ("gde_A", "情境 A：數字能源營收年增（FY2027 起）", "%", None, 0.35, None, "Assumed", "新興業務 2025 年 +16%、2026 前 9 月 +88%（含機器人）；數字能源單獨未揭露", "", ""),
    ("gde_B", "情境 B：數字能源營收年增（FY2027 起）", "%", None, 0.25, None, "Assumed", "同上", "", ""),
    ("gde_C", "情境 C：數字能源營收年增（FY2027 起）", "%", None, 0.10, None, "Assumed", "同上", "", ""),
    ("gm30_A", "情境 A：FY2030 毛利率", "%", None, 0.295, None, "Assumed", "2025 年 28.95%、2026 前 9 月 28.2%；通用自動化毛利率約 44%、新能源車約 14.5%（" + S_IF + "），組合與原料成本主導", "", "FY2026 起線性走到 FY2030，其後持平"),
    ("gm30_B", "情境 B：FY2030 毛利率", "%", None, 0.28, None, "Assumed", "同上", "", ""),
    ("gm30_C", "情境 C：FY2030 毛利率", "%", None, 0.255, None, "Assumed", "同上；2026 第 3 季 25.5%", "", ""),
    ("tp_wA", "情境權重 A", "", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", ""),
    ("tp_wB", "情境權重 B", "", None, 0.50, None, "Assumed", "同上", "", ""),
    ("tp_wC", "情境權重 C", "", None, 0.25, None, "Assumed", "同上", "", ""),
    ("sec", "5. 損益與現金流"),
    ("g_ot", "其他營收年增（FY2027 起）", "%", 0.0, 0.05, 0.10, "Assumed", "占營收 2%，未揭露", "", ""),
    ("opex30", "FY2030 營業費用及其他營業損益淨額 ÷ 營收", "%", 0.155, 0.165, 0.18, "Assumed", "2025 年 16.5%；研發費用率 9.44%（" + S_IF + "）；營運槓桿與研發投入互抵", "", "FY2026→FY2030 線性，其後持平"),
    ("nonop_ss", "FY2027 起業外淨額", "百萬元", 0, 150, 400, "Assumed", "2025 年約 178；公允價值波動不預測", "", "業外；DCF 另以淨現金計入"),
    ("tax_r", "所得稅（含少數股東損益）÷ 稅前", "%", 0.05, "=Hist!E18/Hist!E16", 0.12, "Derived", "＝2025 年（稅前 − 歸母淨利）÷ 稅前；歷史不拆少數股東損益（高新技術企業 15% 與研發加計扣除）", "", "敏感度低／高 5%／12%（Assumed）"),
    ("payout", "配息率", "%", None, "=Hist!E21/Hist!E20", None, "Derived", "＝每股股利 0.50 元 ÷ 2025 年 EPS（" + S_SAF + "）", "", "用於收支關卡與目標時點前股利"),
    ("da26", "FY2026 折舊攤銷 ÷ 營收", "%", None, "=Hist!E23/Hist!E4", None, "Derived", "＝2025 年折舊攤銷 926 ÷ 營收（" + S_SAC + "）", "", "之後線性收斂到資本支出比例 × da_ss（FY2032）"),
    ("da_ss", "穩態折舊 ÷ 資本支出", "x", 0.4, 0.5, 0.8, "Assumed", "2025 年折舊 926 ÷ 資本支出 3,024＝31%（Hist，Derived）；擴產完工後折舊追上，基準收斂到 50%", "", ""),
    ("cx26", "FY2026 資本支出", "百萬元", 2500, 3300, 4000, "Assumed", "2025 年 3,024、2024 年 2,112（" + S_SAC + "）", "", ""),
    ("cx_r", "資本支出 ÷ 營收（FY2027 起）", "%", 0.05, 0.06, 0.08, "Assumed", "2023 年 5.0%、2024 年 5.7%、2025 年 6.7%（Hist，Derived）", "", ""),
    ("nwc_r", "營運資金增加 ÷ 營收增量", "%", 0.15, 0.25, 0.40, "Assumed", "2026 前 9 月存貨 124.58 億元、較年初 +54%（" + S_10 + "）；經營現金流 1–9 月 8.95 億元（−77%）", "", "營運資金吃現金為近期主要風險"),
    ("h_iss", "H 股發行占擴大後股本（尾端事件，基準 0）", "%", None, 0.0, 0.10, "Assumed", "公司 2026 年 1 月宣布籌備 H 股上市（" + S_PF + "）；規模、價格未揭露", "", "原則第 10 節第 19 條：事件型尾端風險進 Sens、不進基準"),
    ("h_disc", "H 股發行價相對現價折價", "%", None, 0.10, None, "Assumed", "港股 A＋H 慣例折價（未查核，不作依據）", "", "只在 h_iss＞0 時有作用"),
    ("sec", "6. 折現率（DCF 標準規格第 2 節）"),
    ("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", S_RF, "", "日期 2026-11-05，較股價基準日早 11 個交易日（逾期；跨模型一致，下次全面改版更新）"),
    ("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", ""),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", ""),
    ("b_own", "本公司 5 年 β（原始）", "x", None, 0.52, None, "Verified", S_SAS, "", ""),
    ("b_peer", "同業 5 年 β 中位數（原始；只有取得 β 的同業）", "x", None, "=MEDIAN(Comp!H5:H10)", None, "Derived", "＝Comp 頁取得 β 的同業（埃斯頓、美的）中位數；日本同業 β 本輪未取得", "", ""),
    ("beta", "β（兩來源 Blume 調整後中位數）", "x", "=beta-0.3", "=MEDIAN(0.67*b_own+0.33,0.67*b_peer+0.33)", "=beta+0.3", "Derived", "Blume 0.67β＋0.33；兩來源＝本公司、同業中位數", "", "敏感度 ∓0.3（跨模型慣例）；β 偏低使 WACC 偏低，見 README 限制"),
    ("kd", "稅前債務成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt", "", "負債占比低"),
    ("wacc_adj", "WACC 調整（敏感度用）", "pt", -0.01, 0, 0.01, "Assumed", "基準 0；敏感度 ±1pt", "", ""),
    ("g_cap", "永續成長率上限", "%", None, 0.03, None, "Assumed", "DCF 標準規格：3%，且不超過計價幣別 10 年期公債殖利率", "", ""),
    ("g_tv", "永續成長率", "%", 0.01, "=MIN(g_cap,rf)", 0.03, "Derived", "＝MIN(g_cap, rf)", "", "敏感度 1%／3%（3% 超過 rf，只作敏感度）"),
    ("ronic_sp", "RONIC − WACC", "pt", None, 0, 0.05, "Assumed", "RONIC 預設＝WACC（新投資價值中性）", "", ""),
    ("tv_alt", "終值改用 FCFF ×（1＋g）（0＝否，1＝是）", "開關", None, 0, 1, "Assumed", "DCF 標準規格第 4 節敏感度", "", ""),
    ("sec", "7. 同業倍數與權重"),
    ("pe_peer", "同業前瞻 P/E 中位數", "x", "=MIN(Comp!F5:F10)", "=MEDIAN(Comp!F5:F10)", "=MAX(Comp!F5:F10)", "Derived", "＝Comp 頁核心同業 6 家前瞻 P/E 中位數；股價皆為 2026-11-20 收盤（EPS 分母不變，見 Comp）", "", "敏感度＝同業最低／最高"),
    ("rel_k", "相對倍數係數（本公司 ÷ 同業）", "x", None, 1.0, 1.5, "Assumed", "1.00（未調整：缺本公司對同業十年相對倍數歷史）", "", "高案 1.5 倍（未查核，不作依據）"),
    ("w_dcf", "DCF 權重（Comparable＝1 − 本值）", "%", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設 DCF 50%＋同業倍數 50%", "", ""),
]
SCN_DRIVERS = [("ga27", "通用自動化 FY2027 年增"), ("ga32", "通用自動化 FY2032 年增"), ("nev27", "新能源汽車 FY2027 年增"), ("nev32", "新能源汽車 FY2032 年增"),
               ("drift", "智能機器人份額漂移"), ("gel", "智慧電梯年增"), ("gde", "數字能源年增"), ("hs30", "2030 年人形伺服份額"), ("gm30", "FY2030 毛利率")]

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
    HI["A1"] = "Hist：歷史財務（百萬元人民幣；中國會計準則）。資料商頁面取整數；稅前、所得稅、股數為頁面比率推得（approx）。公式只引用本頁（定義名稱 h_*）。"
    hdr(HI, 3, ["項目", "", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"])
    HROWS = [  # (列, 鍵, 項目, [2023, 2024, 2025], 標記, 來源)
        (4, "h_rev", "營業收入", [30420, 37041, 45105], "Verified", S_SAF + "；2025 年 451.05 億元與鳳凰財經轉錄年報一致（" + S_IF.split("（")[0] + "）"),
        (5, "h_cogs", "營業成本（＝營收 − 毛利）", ["=C4-C6", "=D4-D6", "=E4-E6"], "Derived", "＝營業收入 − 毛利"),
        (6, "h_gp", "毛利", [9635, 10329, 12682], "Verified", S_SAF + "；2025 年毛利率 28.95%（" + S_IF.split("（")[0] + "）"),
        (7, "h_gm", "毛利率", ["=C6/C4", "=D6/D4", "=E6/E4"], "Derived", "＝毛利 ÷ 營收"),
        (8, "h_seg_ga", "通用自動化 營收", [None, "=E8/1.23", 16900], "Verified", S_IF + "：2025 年 169.00 億元（+23%）；2024＝2025 ÷ 1.23（Derived）；2023 年分部本輪未取得"),
        (9, "h_seg_el", "智慧電梯 營收", [None, "=E9/1.05", 5200], "Verified", S_IF + "：2025 年 52.00 億元（+5%）；2024＝2025 ÷ 1.05（Derived）"),
        (10, "h_seg_nev", "新能源汽車動力系統 營收", [None, "=E10/1.26", 20323], "Verified", S_IF + "：2025 年 203.23 億元（+26%，毛利率 14.5%）；2024＝2025 ÷ 1.26（Derived）"),
        (11, "h_seg_em", "新興業務（智能機器人＋數字能源）營收", [None, "=E11/1.16", 1795], "Verified", S_IF + "：2025 年 17.95 億元（+16%，毛利率 28.3%）；2024＝2025 ÷ 1.16（Derived）"),
        (12, "h_seg_ot", "其他營收（殘差，含取整差）", [None, "=D4-SUM(D8:D11)", "=E4-SUM(E8:E11)"], "Derived", "官方其他 7.42 億元；工業自動化合計 222.45 億元與通用自動化＋電梯相差 1.45 億元，併入本列，故 2025 年為 88.7 而非 74.2"),
        (13, "h_opex", "營業費用及其他營業損益淨額（＝毛利 − 營業利益）", ["=C6-C14", "=D6-D14", "=E6-E14"], "Derived", "含銷售、管理、研發、稅金、減值、其他收益淨額；銷管分項本輪未取得（2025 年研發 42.56 億元，期間費用 75.5 億元含財務費用）"),
        (14, "h_ebit", "營業利益", [4192, 4358, 5257], "Verified", S_SAF),
        (15, "h_ebitm", "營業利益率", ["=C14/C4", "=D14/D4", "=E14/E4"], "Derived", ""),
        (16, "h_pbt", "稅前利益（approx）", [5001, 4608, 5435], "Verified", S_SAF + "：頁面稅前利益率 × 營收，取整"),
        (17, "h_nonop", "業外（＝稅前 − 營業利益）", ["=C16-C14", "=D16-D14", "=E16-E14"], "Derived", "投資收益、公允價值、財務費用、營業外收支合計"),
        (18, "h_tax", "所得稅（含少數股東損益；＝稅前 − 歸母淨利）", ["=C16-C19", "=D16-D19", "=E16-E19"], "Derived", "歷史不拆少數股東損益"),
        (19, "h_ni", "歸屬母公司淨利", [4742, 4285, 5050], "Verified", S_SAF),
        (20, "h_eps", "EPS（基本，元）", [1.78, 1.60, 1.85], "Verified", S_SAF),
        (21, "h_dps", "每股股利（元）", [0.45, 0.41, 0.50], "Verified", S_SAF + "：0.50 為頁面「Current」值，視為 FY2025 股利"),
        (22, "h_sh", "股數（＝淨利 ÷ EPS，百萬股）", ["=C19/C20", "=D19/D20", "=E19/E20"], "Derived", ""),
        (23, "h_da", "折舊攤銷", [589.94, 759.5, 925.67], "Verified", S_SAC),
        (24, "h_cx", "資本支出", [1509, 2112, 3024], "Verified", S_SAC),
        (25, "h_ocf", "營業現金流", [3370, 7200, 6681], "Verified", S_SAC),
        (26, "h_rnd", "研發費用", [None, "=E26/1.3523", 4256], "Verified", S_IF + "：2025 年 42.56 億元（+35.23%，費用率 9.44%）；2024 為回推"),
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
    HI.cell(28, 1, "檢查：分部合計 − 營業收入（應為 0）")
    for c in "DE": HI[f"{c}28"] = f"={c}8+{c}9+{c}10+{c}11+{c}12-{c}4"
    HI.column_dimensions["A"].width = 58; HI.column_dimensions["G"].width = 90

    # ================= Q（2026 年 1–9 月分業務）=================
    Q["A1"] = "Q：2026 年 1–9 月分業務（百萬元）。一列一個項目；數字為券商轉錄公司三季報（億元取整，原文未直接讀取）。D＝1–9 月 ÷（1＋年增）＝2025 年同期（Derived）；F＝2025 全年 − 同期＝2025 年第 4 季。"
    hdr(Q, 3, ["項目", "2026 年 1–9 月", "年增", "2025 年 1–9 月（回推）", "FY2025", "2025 年第 4 季（回推）", "標記", "來源", "來源誘因", "說明"])
    QROWS = [
        (4, "通用自動化", 15900, 0.33, "=Hist!E8", "Verified", S_SW + "；" + S_DF1, "約 159 億元、+33%（伺服 69、變頻器 48、PLC＆HMI 19 億元）"),
        (5, "智慧電梯", 3800, 0.06, "=Hist!E9", "Verified", S_SW, "約 38 億元、+6%"),
        (6, "新能源汽車動力系統", 15500, 0.08, "=Hist!E10", "Verified", S_SW, "約 155 億元、+8%；第 3 季聯合動力營收約 62 億元（+15%）、毛利率 10.8%"),
        (7, "新興業務（智能機器人＋數字能源）", 2100, 0.88, "=Hist!E11", "Verified", S_SW + "；" + S_DF1, "約 21 億元、+88%"),
        (8, "營業收入", 37956, 0.1988, "=Hist!E4", "Verified", S_10, "379.56 億元、+19.88%；第 3 季 132.81 億元（+19.08%）"),
    ]
    for (rr, lab, b, g, e, tag, src, note) in QROWS:
        Q.cell(rr, 1, lab); Q.cell(rr, 2, b).font = BLUE; c = Q.cell(rr, 3, g); c.font = BLUE; c.number_format = "0.0%"
        Q.cell(rr, 4, f"=B{rr}/(1+C{rr})"); Q.cell(rr, 5, e); Q.cell(rr, 6, f"=E{rr}-D{rr}")
        Q.cell(rr, 7, tag); Q.cell(rr, 8, src); Q.cell(rr, 10, note)
    Q.cell(9, 1, "其他營收（殘差）")
    Q["B9"] = "=B8-SUM(B4:B7)"; Q["D9"] = "=D8-SUM(D4:D7)"; Q["E9"] = "=Hist!E12"; Q["F9"] = "=E9-D9"; Q.cell(9, 7, "Derived"); Q.cell(9, 10, "營收減四個分業務；含取整差")
    Q.cell(10, 1, "毛利率（2026 年 1–9 月；第 3 季）"); Q["B10"] = 0.282; Q["D10"] = 0.255; Q["B10"].font = BLUE; Q["D10"].font = BLUE
    Q.cell(10, 7, "Verified"); Q.cell(10, 8, S_10); Q.cell(10, 10, "B＝前三季 28.2%、D＝第 3 季單季 25.5%（欄名不適用）")
    Q.cell(11, 1, "歸屬母公司淨利（2026 年 1–9 月）"); Q["B11"] = 4398; Q["B11"].font = BLUE; Q["C11"] = 0.0338; Q["C11"].font = BLUE
    Q.cell(11, 7, "Verified"); Q.cell(11, 8, S_10); Q.cell(11, 10, "43.98 億元、+3.38%；扣非 39.07 億元（+0.49%）；第 3 季含公允價值收益約 4.2 億元")
    Q.column_dimensions["A"].width = 34; Q.column_dimensions["H"].width = 80

    # ================= 線取數 =================
    LN["A1"] = "線取數：本模型讀入的線層級變數年度路徑（建置時由線模型成品依 ID 讀值）。情境對應：A＝線模型高情境、B＝基準、C＝低情境。"
    hdr(LN, 3, ["項目", "情境", ""] + [f"FY{y}" for y in range(2025, 2033)] + ["單位", "來源（檔案與儲存格）"])
    LR = {}
    r = 4
    for key, lab, unit, dd, y0 in (("cn", "中國工業機器人裝機（含協作；線 3）", "台", LD["cn"], 2025),
                                   ("hum", "全球人形機器人出貨（線 4）", "台", LD["hum"], 2026)):
        for s in SC:
            LN.cell(r, 1, lab); LN.cell(r, 2, f"{s}（{LINE_SC[s]}）")
            srcs = []
            for y in range(y0, 2033):
                v, src = dd[(y, s)]
                c = LN.cell(r, 4 + (y - 2025), v); c.font = BLUE; c.number_format = "#,##0"
                srcs.append(src)
            LN.cell(r, 12, unit); LN.cell(r, 13, f"{srcs[0]} … {srcs[-1]}")
            LR[(key, s)] = r; r += 1
    LN.column_dimensions["A"].width = 52; LN.column_dimensions["M"].width = 110

    # ================= Thesis =================
    TH["A1"] = ("命題模組：營收＝通用自動化（基礎成長＋人形伺服增量）＋智慧電梯＋新能源汽車＋新興業務（智能機器人＝中國裝機（線 3）× 份額漂移 × 單價；數字能源＝成長）＋其他。"
                "FY2026＝1–9 月實績＋2025 年第 4 季 ×（1＋假設年增）。百萬元。")
    TH["A2"] = "欄：D＝FY2025 … K＝FY2032。人形伺服增量＝全球人形伺服驅動市場（線 4 出貨 × 線 6 軸數 × 單價）× 匯川份額；2026 年已含於通用自動化故增量為 0。"
    hdr(TH, 3, ["項目", "說明", ""] + [f"FY{y}" for y in range(2025, 2033)])
    THR = {}
    r = 4
    for s in SC:
        TH.cell(r, 1, f"情境 {s}").font = BOLD; TH.cell(r, 1).fill = SECF; r += 1
        R = {}
        def put(key, lab, note, f25, f26, fn, fmt=None):
            nonlocal r
            R[key] = r; TH.cell(r, 1, lab); TH.cell(r, 2, note)
            for y in range(2025, 2033):
                f = f25 if y == 2025 else (f26 if y == 2026 else fn(y))
                if callable(f25) and y == 2025: f = f25
                if f is not None:
                    TH[f"{TC[y]}{r}"] = f
                    if fmt: TH[f"{TC[y]}{r}"].number_format = fmt
            r += 1
        P_ = lambda y, k: f"{TC[y - 1]}{R[k]}"
        put("cn", "中國工業機器人裝機（線 3，千台）", "線取數頁", f"=線取數!D{LR[('cn', s)]}/1000", f"=線取數!E{LR[('cn', s)]}/1000", lambda y: f"=線取數!{TC[y]}{LR[('cn', s)]}/1000")
        put("ship", "全球人形出貨（線 4 × 出貨倍數，千台）", "", None, f"=線取數!E{LR[('hum', s)]}*hu_k/1000", lambda y: f"=線取數!{TC[y]}{LR[('hum', s)]}*hu_k/1000")
        put("hsv", "全球人形伺服驅動市場（百萬元）", "＝出貨 × 每台軸數 × 單價 ×（1 − 年降）^(年 − 2026) × 匯率", None,
            f"=E{r - 1}*ax*psd*fx/1000", lambda y: f"={TC[y]}{r - 1}*ax*psd*(1-pds)^({y}-2026)*fx/1000")
        put("hsh", "匯川人形伺服份額", "2026 為 0（已含於通用自動化）；線性走到 hs30，其後持平", None, 0, lambda y: f"=hs30_{s}*MIN(1,({y}-2026)/4)", "0.0%")
        put("hinc", "人形伺服增量營收（百萬元）", "＝份額 × 市場", None, f"=E{R['hsh']}*E{R['hsv']}", lambda y: f"={TC[y]}{R['hsh']}*{TC[y]}{R['hsv']}")
        put("gab", "通用自動化：基礎（子列）", "FY2027 起＝前一年 ×（1＋年增；FY2027 值線性走到 FY2032 值）", "=Hist!E8", "=Q!B4+Q!F4*(1+q4g_ga)",
            lambda y: f"={P_(y, 'gab')}*(1+ga27_{s}+(ga32_{s}-ga27_{s})*({y}-2027)/5)")
        put("ga", "通用自動化 合計", "基礎＋人形伺服增量", "=Hist!E8", f"=E{R['gab']}+E{R['hinc']}", lambda y: f"={TC[y]}{R['gab']}+{TC[y]}{R['hinc']}")
        put("el", "智慧電梯", "", "=Hist!E9", "=Q!B5+Q!F5*(1+q4g_el)", lambda y: f"={P_(y, 'el')}*(1+gel_{s})")
        put("nev", "新能源汽車動力系統", "FY2027 年增線性走到 FY2032 年增", "=Hist!E10", "=Q!B6+Q!F6*(1+q4g_nev)",
            lambda y: f"={P_(y, 'nev')}*(1+nev27_{s}+(nev32_{s}-nev27_{s})*({y}-2027)/5)")
        put("emer", "新興業務 FY2026（錨）", "＝1–9 月＋2025Q4 ×（1＋q4g_em）", "=Hist!E11", "=Q!B7+Q!F7*(1+q4g_em)", lambda y: None)
        put("rob", "智能機器人（子列）", "FY2026＝新興 × rob_sh26；之後＝前一年 × 中國裝機成長 ×（1＋漂移）×（1 − pdr）", None, f"=E{R['emer']}*rob_sh26",
            lambda y: f"={P_(y, 'rob')}*{TC[y]}{R['cn']}/{P_(y, 'cn')}*(1+drift_{s})*(1-pdr)")
        put("de", "數字能源（子列）", "", None, f"=E{R['emer']}*(1-rob_sh26)", lambda y: f"={P_(y, 'de')}*(1+gde_{s})")
        put("em", "新興業務 合計", "", "=Hist!E11", f"=E{R['rob']}+E{R['de']}", lambda y: f"={TC[y]}{R['rob']}+{TC[y]}{R['de']}")
        put("oth", "其他營收", "", "=Hist!E12", "=Q!B9+Q!F9*(1+q4g_ot)", lambda y: f"={P_(y, 'oth')}*(1+g_ot)")
        put("rev", "合併營收", "", "=Hist!E4", f"=E{R['ga']}+E{R['el']}+E{R['nev']}+E{R['em']}+E{R['oth']}",
            lambda y: f"={TC[y]}{R['ga']}+{TC[y]}{R['el']}+{TC[y]}{R['nev']}+{TC[y]}{R['em']}+{TC[y]}{R['oth']}")
        put("robsh", "智能機器人隱含占中國裝機的營收強度（元／台，Derived）", "＝智能機器人營收 ÷ 中國裝機；只看趨勢", None, f"=E{R['rob']}/E{R['cn']}*1000", lambda y: f"={TC[y]}{R['rob']}/{TC[y]}{R['cn']}*1000", "#,##0")
        put("humsh", "人形伺服增量占合併營收", "", None, f"=E{R['hinc']}/E{R['rev']}", lambda y: f"={TC[y]}{R['hinc']}/{TC[y]}{R['rev']}", "0.0%")
        put("gm", "毛利率路徑", "FY2026＝gm26；線性到 FY2030 gm30，其後持平", None, "=gm26", lambda y: f"=gm26+(gm30_{s}-gm26)*MIN(1,({y}-2026)/4)", "0.0%")
        put("opex", "營業費用率路徑", "FY2026＝opex26；線性到 FY2030 opex30，其後持平", None, "=opex26", lambda y: f"=opex26+(opex30-opex26)*MIN(1,({y}-2026)/4)", "0.0%")
        r += 1
        THR[s] = R
    for row in TH.iter_rows(min_row=4, max_row=r, min_col=4, max_col=11):
        for c in row:
            if c.number_format == "General": c.number_format = "#,##0.0"
    TH.column_dimensions["A"].width = 44; TH.column_dimensions["B"].width = 50

    # ================= PL（標準損益表，三情境）=================
    PL["A1"] = ("PL：官方分部（通用自動化／智慧電梯／新能源汽車動力系統／新興業務／其他）→ 合併營收 → 標準損益表 → FCFF。"
                "C＝FY2023 … E＝FY2025 為實績；F＝FY2026（1–9 月實績＋第 4 季假設）… L＝FY2032。百萬元；中國會計準則（法定財報口徑，GAAP）。FY2023 分部營收未取得。")
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
        HS = lambda rr: (lambda y: None if y == 2023 else f"=Hist!{'CDE'[y - 2023]}{rr}")
        C = lambda y: COL[y]
        T_ = lambda k: (lambda y: f"=Thesis!{TC[y]}{th[k]}")
        row("gab", "通用自動化：基礎（子列）", "百萬元", None, T_("gab"))
        row("ghu", "通用自動化：人形伺服增量（子列）", "百萬元", None, T_("hinc"))
        row("ga", "通用自動化 營收", "百萬元", HS(8), lambda y: f"={C(y)}{R['gab']}+{C(y)}{R['ghu']}")
        row("el", "智慧電梯 營收", "百萬元", HS(9), T_("el"))
        row("nev", "新能源汽車動力系統 營收", "百萬元", HS(10), T_("nev"))
        row("rob", "新興業務：智能機器人（子列）", "百萬元", None, T_("rob"))
        row("de", "新興業務：數字能源（子列）", "百萬元", None, T_("de"))
        row("em", "新興業務（智能機器人＋數字能源）營收", "百萬元", HS(11), lambda y: f"={C(y)}{R['rob']}+{C(y)}{R['de']}")
        row("oth", "其他營收", "百萬元", HS(12), T_("oth"))
        row("rev", "營收（合併）", "百萬元", H(4), lambda y: f"={C(y)}{R['ga']}+{C(y)}{R['el']}+{C(y)}{R['nev']}+{C(y)}{R['em']}+{C(y)}{R['oth']}")
        chkf = lambda y: f"={C(y)}{R['ga']}+{C(y)}{R['el']}+{C(y)}{R['nev']}+{C(y)}{R['em']}+{C(y)}{R['oth']}-{C(y)}{R['rev']}"
        row("chk", "檢查：分部合計 − 合併營收（應為 0）", "百萬元", lambda y: None if y == 2023 else chkf(y), chkf)
        PL.cell(R["chk"], 1, "檢查：分部合計 − 合併營收（應為 0）" if s == "B" else f"檢查（情境 {s}）：分部合計 − 合併營收（應為 0）")
        row("chk2", "檢查：子列合計 − 分部（應為 0）", "百萬元", None, lambda y: f"={C(y)}{R['gab']}+{C(y)}{R['ghu']}-Thesis!{TC[y]}{th['ga']}+{C(y)}{R['rob']}+{C(y)}{R['de']}-Thesis!{TC[y]}{th['em']}")
        row("yoy", "營收年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1",
            lambda y: f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1", "0.0%")
        row("cogs", "營業成本", "百萬元", H(5), lambda y: f"={C(y)}{R['rev']}*(1-Thesis!{TC[y]}{th['gm']})")
        row("gp", "毛利", "百萬元", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}")
        row("gm", "毛利率", "%", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", "0.0%")
        row("opex", "營業費用（銷管研、稅金、減值與其他營業損益淨額）", "百萬元", H(13), lambda y: f"={C(y)}{R['rev']}*Thesis!{TC[y]}{th['opex']}")
        row("op", "營業利益", "百萬元", H(14), lambda y: f"={C(y)}{R['gp']}-{C(y)}{R['opex']}")
        row("opm", "營業利益率", "%", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", "0.0%")
        row("nonop", "業外（投資收益、公允價值、財務費用、營業外）", "百萬元", H(17), lambda y: "=nonop26" if y == 2026 else "=nonop_ss")
        row("pbt", "稅前利益", "百萬元", H(16), lambda y: f"={C(y)}{R['op']}+{C(y)}{R['nonop']}")
        row("tax", "所得稅（含少數股東損益；歷史不拆）", "百萬元", H(18), lambda y: f"={C(y)}{R['pbt']}*tax_r")
        row("ni", "歸屬母公司淨利（預估期＝稅前 − 所得稅含少數股東）", "百萬元", H(19), lambda y: f"={C(y)}{R['pbt']}-{C(y)}{R['tax']}")
        row("nim", "淨利率", "%", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", "0.0%")
        row("niy", "淨利年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1",
            lambda y: f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1", "0.0%")
        row("sh", "股數（百萬股）", "百萬股", H(22), lambda y: "=sh_v/(1-h_iss)")
        row("eps", "EPS（元）", "元", H(20), lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['sh']}", "0.00")
        row("da", "折舊攤銷", "百萬元", H(23), lambda y: f"={C(y)}{R['rev']}*(da26+(cx_r*da_ss-da26)*{y - 2026}/6)")
        row("cx", "資本支出", "百萬元", H(24), lambda y: "=cx26" if y == 2026 else f"={C(y)}{R['rev']}*cx_r")
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
    CP["A1"] = ("Comp：同業前瞻 P/E。股價＝" + S_CLOSE + "（Verified）；前瞻 P/E＝原值 × 新收盤 ÷ 原報價（Derived，EPS 分母不變）；"
                "日本同業取本專案同日（2026-11-20）模型值；埃斯頓、美的取資料商 statistics 頁（原報價日 2026-11-24）。核心同業 6 家進中位數，少於 8 家不做迴歸。")
    hdr(CP, 4, ["公司", "代號", "股價（2026-11-20）", "報價日", "前瞻 P/E 分母", "前瞻 P/E", "近 12 月 P/E（原報價日）", "5 年 β", "會計口徑", "進中位數", "說明", "原報價", "原報價日", "原前瞻 P/E", "標記與換算"])
    PEERS = [  # (公司, 代號, 11-20 收盤, 原報價, 原報價日, 原前瞻 P/E, β, 口徑, 說明, 分母, 來源)
        ("安川電機", "6506.T", 3180, 3180, "2026-11-20", 15.6739, None, "日本基準（未查核）", "伺服、機器人；匯川直接競爭者", "資料商 NTM 共識 EPS", "physical_ai/thk 模型 In 頁 pe_yask（2026-11-20 換算值）"),
        ("Fanuc", "6954.T", 5047, 5047, "2026-11-20", 21.1499, None, "日本基準（未查核）", "工業機器人、CNC", "資料商 NTM 共識 EPS", "physical_ai/thk 模型 In 頁 pe_fanuc（2026-11-20 換算值）"),
        ("埃斯頓", "002747.SZ", 24.25, 27.24, "2026-11-24", 79.78, 0.82, "中國會計準則", "國產工業機器人與伺服；獲利基期低，P/E 偏高", "資料商 NTM 共識 EPS", "stockanalysis.com SHE:002747 statistics（2026-11-24 讀取）"),
        ("美的集團", "000333.SZ", 94.82, 94.23, "2026-11-24", 14.63, 0.43, "中國會計準則", "家電為主，含 KUKA 機器人與工業自動化；相似度低", "資料商 NTM 共識 EPS", "stockanalysis.com SHE:000333 statistics（2026-11-24 讀取）"),
        ("Keyence", "6861.T", 14025, 14025, "2026-11-20", 28.5525, None, "日本基準（未查核）", "FA 感測與視覺", "資料商 NTM 共識 EPS", "physical_ai/keyence 模型 In 頁 pe_self（2026-11-20）"),
        ("SMC", "6273.T", 55300, 55300, "2026-11-20", 14.5941, None, "日本基準（未查核）", "氣動 FA 零組件", "公司 FY2026 指引 EPS", "physical_ai/smc 模型 In 頁 pe_self（2026-11-20）"),
        ("匯川技術（本公司，對照）", "300124.SZ", 52.66, 53.61, "2026-11-24", 22.84, 0.52, "中國會計準則", "本公司；只作對照", "資料商 NTM 共識 EPS", S_SAS),
    ]
    for i, (nm, tk, px1, px0, d0, pe0, bt, acc, note, den, srcp) in enumerate(PEERS):
        rr = 5 + i
        CP.cell(rr, 1, nm); CP.cell(rr, 2, tk)
        CP.cell(rr, 3, px1).font = BLUE; CP.cell(rr, 4, "2026-11-20"); CP.cell(rr, 5, den)
        CP.cell(rr, 6, f"=N{rr}*C{rr}/L{rr}")
        if bt is not None: CP.cell(rr, 8, bt).font = BLUE
        CP.cell(rr, 9, acc); CP.cell(rr, 10, "是" if i < 6 else "否"); CP.cell(rr, 11, note)
        CP.cell(rr, 12, px0).font = BLUE; CP.cell(rr, 13, d0); CP.cell(rr, 14, pe0).font = BLUE
        CP.cell(rr, 15, "股價 Verified（" + S_CLOSE + "）；前瞻 P/E Derived：＝原值 × 2026-11-20 收盤 ÷ 原報價；原值來源：" + srcp)
    CP.cell(13, 1, "報價日未取得（不進中位數）：安川電機以外的日系同業已用同日值；台股上銀（2049.TW）收盤檔有但本輪未取得其前瞻 P/E；台達電、Siemens、Rockwell 不在收盤檔")
    CP.cell(15, 1, "核心同業前瞻 P/E 中位數"); CP["F15"] = "=MEDIAN(F5:F10)"
    CP.cell(16, 1, "本公司現價 ÷ 資料商前瞻 EPS（換算至 2026-11-20；分母：資料商 NTM 共識）"); CP["F16"] = "=F11"
    CP.cell(17, 1, "本公司現價 ÷ 券商 FY2027 歸母淨利預測平均除以股數（申萬 6,746、東方財富 7,240、另一家 7,070 百萬元；FY2027）"); CP["F17"] = "=px/(AVERAGE(6746,7070,7240)/sh_v)"
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
    row1("cons_show", "市場共識目標價（資料商 23 家平均，2026-11-24 讀取；只作對照）", "=cons_tp")
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
        SE.cell(r + 1, 1, "說明：折舊、資本支出（FY2027 起）、營運資金隨營收等比例變動（第 44 條）；各驅動見 In 頁區間。")
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
    rrow("FY2026 歸屬母公司淨利（情境 B）vs 券商 2026E 平均（申萬 5,645、東方財富 5,820、另一家 6,010 百萬元）", f"=PL!F{PR['B']['ni']}", "=AVERAGE(5645,5820,6010)", "券商預測為 Interested-party；含第 3 季公允價值收益，模型業外 FY2026 假設 5 億元")
    rrow("FY2027 歸屬母公司淨利（情境 B）vs 券商 2027E 平均（6,746、7,070、7,240）", f"=PL!G{PR['B']['ni']}", "=AVERAGE(6746,7070,7240)", "")
    rrow("FY2028 歸屬母公司淨利（情境 B）vs 券商 2028E 平均（7,775、8,460、8,710）", f"=PL!H{PR['B']['ni']}", "=AVERAGE(7775,8460,8710)", "")
    rrow("FY2026 營收（情境 B）vs 1–9 月實績年化（379.56 億 ÷ 0.75）", f"=PL!F{PR['B']['rev']}", "=Q!B8/0.75", "第 4 季季節性較高，模型高於年化屬預期")
    rrow("FY2026 新能源車營收（情境 B）vs 公司展望 2026 年 230–240 億元（取中值 23,500）", f"=Thesis!E{THR['B']['nev']}", 23500, S_DF2)
    rrow("FY2026 毛利率 vs 2026 年 1–9 月實績", f"=PL!F{PR['B']['gm']}", "=Q!B10", "模型假設第 4 季接近第 3 季（25.5%），故低於 1–9 月", "0.0%")
    rrow("FY2026 第 4 季隱含營收 vs 2026Q3 實績（13,281 百萬元）", f"=PL!F{PR['B']['rev']}-Q!B8", 13281, "第 4 季季節性高於第 3 季")
    rrow("加權目標價 vs 市場共識目標價（資料商 23 家，2026-11-24 讀取）", f"=Val!B{VR['tp']}", "=cons_tp", S_SAS + "；只作對照，不進模型")
    rr += 1
    RC.cell(rr, 1, "線 3 對照").font = BOLD; rr += 1
    rrow("FY2026 智能機器人營收（情境 B；Assumed 占新興 60%）÷ 線 3 中國裝機＝每台營收（元）", f"=Thesis!E{THR['B']['robsh']}", None, "中國裝機基準 2026 年 379,906 台（線 3）；匯川機器人含 SCARA、協作、六軸，單價遠低於外資；只看趨勢", "#,##0", False)
    rrow("FY2032 智能機器人營收（情境 B）÷ 中國裝機", f"=Thesis!K{THR['B']['robsh']}", None, "同上；與 FY2026 相比的倍數反映份額漂移與單價年降的合計", "#,##0", False)
    rrow("FY2030 人形伺服增量（情境 B，百萬元）占全球人形伺服驅動市場", f"=Thesis!I{THR['B']['hsh']}", None, "＝情境 B 2030 年人形伺服份額；全球市場 FY2030 見 Thesis 頁", "0.0%", False)
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
    c26 = LD["cn"][(2026, "B")]; c32 = LD["cn"][(2032, "B")]; h26 = LD["hum"][(2026, "B")]; h32 = LD["hum"][(2032, "B")]
    ZR = [
        ("中國工業機器人裝機 FY2026–FY2032（A／B／C＝線高／基準／低）", "3 工業機器人", f"基準 {c26[0]:,.0f} → {c32[0]:,.0f}", "台", "IFR 新裝機含協作", f"{c26[1]}；{c32[1]}（逐年見線取數頁）", "Derived", "Thesis 智能機器人營收成長", "一致（直接讀值）"),
        ("全球人形機器人出貨 FY2026–FY2032（A／B／C＝線高／基準／低）", "4 人形", f"基準 {h26[0]:,.0f} → {h32[0]:,.0f}（× hu_k）", "台", "全球、含小型機", f"{h26[1]}；{h32[1]}", "Assumed", "Thesis 人形伺服市場", "一致（hu_k＝1）；高案 hu_k＝2 為本模型自設，列 Sens 不改基準"),
        ("人形每台伺服驅動軸數", "6 共用零組件", "=ax", "軸／台", "線 6 輸入頁人形矩陣 SD 列", f"{LD['ax'][1][1]}（低 {LD['ax'][0][1]}、高 {LD['ax'][2][1]}）", "Analogy", "Thesis 人形伺服市場", "一致（直接讀值）；線 4 Q_MOTOR 為 30 台無框馬達，口徑不同未用"),
        ("伺服驅動單價 2026 與年降幅", "6 共用零組件", "=psd", "美元／軸", "線 6 自標未取得實際單價（Assumed）", f"{LD['psd'][1][1]}；年降 {LD['pdd'][1][1]}", "Assumed", "Thesis 人形伺服市場", "一致（直接讀值）；線 6 以 7.1 元／美元換算，本模型以 2026-11-20 的 6.6665 換算，差約 6%"),
        ("中國裝機國產品牌占比 2025", "3 工業機器人", f"{LD['dom'][0]:.0%}", "%", "IFR", f"{LD['dom'][1]}", "Verified", "drift_* 的依據（定性）", "一致"),
        ("中國工業機器人產量 2026 年 1–9 月年增", "3 工業機器人", f"{LD['prod'][0]:.1%}", "%", "國家統計局產量（含出口）", f"{LD['prod'][1]}", "Verified", "q4g_em 的依據（定性）", "一致；本模型新興業務 +88% 遠高於產業 +29%，差額由份額漂移與數字能源承擔"),
        ("國產 6 軸回本期（中國一般工業）", "3 工業機器人", f"{LD['pay'][0]:.2f} 年", "年", "線 3 單位經濟", f"{LD['pay'][1]}", "Derived", "drift_*（定性）", "一致"),
        ("伺服在通用自動化中的位置（伺服約 69 億元、變頻器約 48 億元、PLC 約 19 億元，2026 前 9 月）", "6 共用零組件", "未建模", "億元", "線 6 無工業伺服市場規模；只列人形增量", S_SW, "Verified", "Thesis 通用自動化（以公司成長路徑，不由線變數推導）", "線 6 無對應變數；通用自動化主體為 Assumed 成長路徑，已在 README 標明"),
    ]
    for i, z in enumerate(ZR):
        for j, v in enumerate(z): DZ.cell(4 + i, 1 + j, v)
    DZ.column_dimensions["A"].width = 50; DZ.column_dimensions["F"].width = 80; DZ.column_dimensions["I"].width = 70

    # ================= Risk =================
    RK["A1"] = "Risk：風險登錄（原則第 10 節第 10、19 條）。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    RISKS = [
        ("新能源車動力系統（聯合動力）毛利率與需求", "是（nev27、nev32、gm30）", "第 3 季毛利率 10.8%、淨利僅 0.4 億元；2026 前 9 月 +8%", "聯合動力季報、客戶車型銷量"),
        ("原料成本（銅、電子料）與毛利率", "是（gm26、gm30）", "第 3 季毛利率年減 2.1pt", "季毛利率"),
        ("營運資金與現金流", "是（nwc_r）", "存貨較年初 +54%，經營現金流 1–9 月年減 77%", "存貨、應收、經營現金流"),
        ("國產機器人價格戰", "部分（pdr、drift_C）", "匯川單價未揭露", "SCARA 份額、毛利率"),
        ("H 股發行（規模、價格未揭露）", "部分（h_iss 尾端 Sens，基準 0）", "外資持股結構與估值參照會變", "股東會決議、聯交所遞表"),
        ("公允價值收益波動（海外基金與股權投資）", "部分（nonop26、nonop_ss）", "第 3 季約 4.2 億元非經常性收益，扣非淨利 +0.5%", "季報非經常性損益"),
        ("資產減值", "是（含於 opex）", "前 9 月減值 5.42 億元（上年同期 2.32 億元）", "季報減值"),
        ("深股通持股下降（19.58%→14.76%）", "否", "不影響盈利；影響股價供需", "持股比例"),
        ("客戶集中度", "否", "年報前五大客戶本輪未取得", "年報"),
    ]
    for i, k in enumerate(RISKS):
        for j, v in enumerate(k): RK.cell(4 + i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["D"].width = 60

    # ================= Sources =================
    SO["A1"] = "Sources：本模型引用的原始來源與資料商（含網址；HTML 展示層不放網址）。"
    for i, s_ in enumerate([S_PX, S_SAF, S_SAS, S_SAB, S_SAC, S_IF, S_DF1, S_DF2, S_SW, S_10, S_RF, S_CRP,
                            f"線 3：{L3N}", f"線 4：{L4N}", f"線 6：{L6N}",
                            "同業前瞻 P/E：physical_ai/thk、keyence、smc 模型（2026-11-20）；埃斯頓、美的取 stockanalysis.com statistics 頁（2026-11-24）",
                            S_PF], 3):
        SO.cell(i, 1, s_)
    SO.column_dimensions["A"].width = 160

    # ================= README（xlsx 內）=================
    S = T
    v = lambda k: f"Val!B{VR[k]}"
    lines = [
        ('="匯川技術 Inovance（300124.SZ）個股模型 ' + VER + '｜build_id "&Sens!D2&"｜資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'),
        "命題：匯川是中國工業自動化龍頭，營收主體為通用自動化（伺服、變頻器、PLC）與新能源汽車動力系統；實體 AI 曝險只在智能機器人（新興業務一部分，隨中國裝機與份額）與人形伺服增量，兩者占營收比例小，結論主要由通用自動化成長與毛利率決定。",
        (f'="結論（條件式）：在 FY2026 以 1–9 月實績為錨、通用自動化 FY2027 年增 "&{S("ga27_B","0%")}&"（線性降到 FY2032 "&{S("ga32_B","0%")}&"）、新能源車 FY2027 年增 "&{S("nev27_B","0%")}&"、FY2030 毛利率 "&{S("gm30_B","0.0%")}&"、同業 P/E 未加溢價（相對倍數係數 "&{S("rel_k","0.00")}&"）的假設下，加權目標價約 "&{S(v("tp"),"#,##0.0")}&" 元，相對現價 "&{S("px","#,##0.00")}&" 元 "&{S(v("up"),"+0.0%;-0.0%")}&"，估值評等 "&{v("rating")}&"，投資建議 "&{v("rec")}&"。"'),
        (f'="DCF 目標價 "&{S(v("dcf_tp"),"#,##0.0")}&"、Comparable 目標價 "&{S(v("cmp_tp"),"#,##0.0")}&"（各半）；WACC "&{S("wacc","0.00%")}&"、ke "&{S("ke","0.00%")}&"、g "&{S("g_tv","0.00%")}&"；情境 B 終值占 EV "&{S(f"Val!C{VR['tvs']}","0%")}&"。"'),
        (f'="上檔空間歸因：DCF 貢獻 "&{S(v("att_d"),"+0.0%;-0.0%")}&"、未調整 Comparable 貢獻 "&{S(v("att_c"),"+0.0%;-0.0%")}&"、相對倍數調整貢獻 "&{S(v("att_r"),"+0.0%;-0.0%")}&"。本公司現價對資料商前瞻 EPS 的倍數是同業中位數的 "&{S(v("relnow"),"0.0")}&" 倍；支持現價需要相對倍數係數 "&{S(v("relreq"),"0.0")}&"，升一級需要 "&{S(v("relflip"),"0.0")}&"。"'),
        (f'="收支關卡：情境 B FY2026–FY2030 FCFF 扣股利合計 "&{S(v("gate_b"),"#,##0.0")}&" 百萬元，"&{v("gate_t")}&"；距最近級距邊界 "&{S(v("dist"),"0.0%")}&"。"'),
        "主要失效條件（可觀察）：(1) 第 4 季與 2027 年毛利率回升或續降；(2) 新能源車營收低於公司展望 230–240 億元；(3) 存貨與經營現金流持續惡化；(4) 通用自動化年增掉到 15% 以下；(5) H 股高價大量發行。",
        "方法：DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，WACC 折現，明確期到 FY2032，RONIC＝WACC；金融資產以淨現金計入）50%＋同業前瞻 P/E × 模型 FY2027 EPS 50%；三情境 25／50／25。",
        "限制：通用自動化與新能源車主體為 Assumed 成長路徑（線模型無對應變數）；銷管研分項、FY2023 分部營收、前五大客戶、庫藏股未取得；分部數字為券商轉錄；β 只取本公司與兩家 A 股同業；同業 P/E 由原報價日按 2026-11-20 收盤等比例換算（EPS 未更新）。詳見 repo README.md。",
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
    for d, lab in SCN_DRIVERS:
        for s, nm in (("A", "樂觀"), ("C", "保守")):
            extra.append((f"情境 B 的{lab}改為情境 {s} 值", f"{d}_B", ivE[f"{d}_{s}"]))
    extra += [("人形出貨為線 4 的 5 倍（全部情境）", "hu_k", 5.0), ("相對倍數係數＝1（基準即為 1）", "rel_k", 1.0)]
    for drv, k, val in extra:
        specs.append(dict(drv=drv, set=f"{k}＝{val:.4g}", key=k, val=val, ov={k: val}))
    specs.append(dict(drv="情境權重全給 A", set="tp_wA＝1、tp_wB＝0、tp_wC＝0", key="tp_wA tp_wB tp_wC", val=None, ov={"tp_wA": 1, "tp_wB": 0, "tp_wC": 0}))
    specs.append(dict(drv="情境權重全給 C", set="tp_wA＝0、tp_wB＝0、tp_wC＝1", key="tp_wA tp_wB tp_wC", val=None, ov={"tp_wA": 0, "tp_wB": 0, "tp_wC": 1}))
    res = run_cases([s["ov"] for s in specs], meta, "sens")
    cases = []
    for s, r in zip(specs, res):
        cases.append(dict(drv=s["drv"], set=s["set"], key=s["key"], val=s["val"], tp=r["tp"], dcf=r["dcf"], rating=r["rating"]))
        print(f"  {s['drv'][:44]:44s} {r['tp']:9.1f} {r['tp'] / base['tp'] - 1:+.1%}")

    # 反解（整本重算，線性二分；目標價對所解變數遞增）
    px = ivE["px"]
    def solve(key, target_tp, lo, hi, n=9):
        if run_cases([{key: hi}], meta, "solve")[0]["tp"] < target_tp: return None
        if run_cases([{key: lo}], meta, "solve")[0]["tp"] >= target_tp: return lo
        a, b = lo, hi
        for _ in range(n):
            m = (a + b) / 2
            if run_cases([{key: m}], meta, "solve")[0]["tp"] < target_tp: a = m
            else: b = m
        return (a + b) / 2
    up = base["up"]
    flip = -0.5 if up < -0.5 else -0.25 if up < -0.25 else 0.25 if up <= 0.25 else 0.5
    flip_tp = px * (1 + flip + 1e-6)
    s1 = solve("gm30_B", px, 0.15, 0.45)
    s2 = solve("gm30_B", flip_tp, 0.15, 0.45)
    s3 = solve("ga27_B", px, 0.0, 0.8)
    s4 = solve("ga27_B", flip_tp, 0.0, 0.8)
    nf = lambda x: x if x is not None else -1
    solve_rows = [
        ("支持現價（目標價＝現價）所需的情境 B FY2030 毛利率", nf(s1), ivE["gm30_B"], "其他輸入不變；−1＝45% 內無解"),
        (f"評等升一級（上檔 {flip:+.0%}）所需的情境 B FY2030 毛利率", nf(s2), ivE["gm30_B"], "其他輸入不變；−1＝45% 內無解"),
        ("支持現價所需的情境 B 通用自動化 FY2027 年增", nf(s3), ivE["ga27_B"], "其他輸入不變（FY2032 年增不動）；−1＝80% 內無解"),
        (f"評等升一級（上檔 {flip:+.0%}）所需的情境 B 通用自動化 FY2027 年增", nf(s4), ivE["ga27_B"], "其他輸入不變；−1＝80% 內無解"),
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
    segs = [("通用自動化", "通用自動化 營收", Hs["E8"].value), ("智慧電梯", "智慧電梯 營收", Hs["E9"].value),
            ("新能源汽車動力系統", "新能源汽車動力系統 營收", Hs["E10"].value), ("新興業務（智能機器人＋數字能源）", "新興業務（智能機器人＋數字能源）營收", Hs["E11"].value),
            ("其他（殘差，含取整差）", "其他營收", Hs["E12"].value)]
    json.dump(dict(company="inovance", ticker="300124.SZ", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=round(v, 2), unit="百萬元", source=S_IF) for n, l, v in segs],
                   sublines=[dict(parent="通用自動化", labels=["通用自動化：基礎（子列）", "通用自動化：人形伺服增量（子列）"]),
                             dict(parent="新興業務（智能機器人＋數字能源）", labels=["新興業務：智能機器人（子列）", "新興業務：數字能源（子列）"])],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）", "檢查：子列合計 − 分部（應為 0）"],
                   note="官方分部取自 2025 年報分業務營收（通用自動化、智慧電梯、新能源汽車動力系統、新興業務、其他）；工業自動化合計與通用自動化加電梯的取整差併入其他。新興業務內智能機器人與數字能源的拆分、通用自動化的人形增量為命題拆分（Assumed），合計等於分部。"),
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
    t = f"""# physical_ai/inovance/ — 匯川技術 Inovance（300124.SZ）個股模型 {VER}

<!-- 本檔由 build_inovance.py 依成品重算值產生；數字不手寫（原則第 10 節第 51 條）。 -->

**本模型適用《高不確定性預估工作原則》**（`docs/plan/高不確定性預估工作原則.md`）：通用自動化與新能源汽車的成長路徑、毛利率為 Assumed 區間；實體 AI 曝險（智能機器人、人形伺服）占營收比例小，結論只在證據邊界內成立。

- 成品：`dist/{NAME}`（build_id `{sens['build_id']}`；LibreOffice 重算，0 錯誤值）；HTML 展示層 `dist/{NAME[:-5]}.html`（`gen_html_json.py` 產生 `html.json`，再以 `tools/company_html/build_company_html.py` 產生）
- 建置：`python3 physical_ai/inovance/build_inovance.py`（需要 soffice；整本重算敏感度與反解；結尾自動跑 in_layout、sens_coverage --require-keycol、model_governance、rating_gate --strict、zero_rows、cached_values、segment_check）
- 日期：資料截止日 2026-11-24（財報、券商預測、同業倍數）｜成品日 2026-11-24｜股價基準日 2026-11-20（收盤 {I['px']:,.2f} 元）｜評價日 2026-12-31｜目標時點 2027-11-20
- 會計年度＝日曆年；單位百萬元人民幣；中國會計準則（法定財報口徑，GAAP）。深交所創業板。

## 結論（條件式）

在「FY2026 以 1–9 月實績為錨、通用自動化 FY2027 年增 {I['ga27_B']:.0%}（線性降到 FY2032 {I['ga32_B']:.0%}）、新能源車 FY2027 年增 {I['nev27_B']:.0%}、FY2030 毛利率 {I['gm30_B']:.1%}、同業 P/E 不加溢價（相對倍數係數 {I['rel_k']:.2f}）」的假設下：

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
| 市場共識目標價（資料商 23 家，2026-11-24 讀取；只作對照） | {I['cons_tp']:,.2f} |

上檔空間歸因：DCF 貢獻 {g('att_d'):+.1%}、未調整 Comparable 貢獻 {g('att_c'):+.1%}、相對倍數調整 {g('att_r'):+.1%}。本公司現價對資料商前瞻 EPS 的倍數是核心同業中位數（{I['pe_peer']:.1f} 倍）的 {g('relnow'):.1f} 倍；要讓目標價＝現價，Comparable 的相對倍數係數須達 {g('relreq'):.2f}。DCF 以 RONIC＝WACC（新投資無超額報酬）。

**讀法：** 匯川的價值由通用自動化（伺服、變頻器、PLC）成長與全公司毛利率決定；實體 AI 曝險小：情境 B 人形伺服增量 {ser('ghu', '{:,.0f}')}（FY2026→FY2032，百萬元），FY2032 占合併營收 {TH[f'K{THB["humsh"]}'].value:.1%}；智能機器人 {ser('rob', '{:,.0f}')} 百萬元。反解：其他輸入不變時，支持現價需要情境 B FY2030 毛利率 {fn(sol[0][1], '{:.1%}')}（基準 {I['gm30_B']:.1%}）、或通用自動化 FY2027 年增 {fn(sol[2][1], '{:.0%}')}（基準 {I['ga27_B']:.0%}）；−1＝區間內無解。結論主要建立在 Assumed 成長路徑與毛利率上（線模型沒有伺服或通用自動化變數可讀），線層級變數（中國機器人裝機、人形出貨、伺服軸數與單價）只影響約一成以下營收。

## 命題

1. **營收結構：** 官方分業務（2025 年報）：通用自動化 {wb['Hist']['E8'].value:,.0f}、智慧電梯 {wb['Hist']['E9'].value:,.0f}、新能源汽車動力系統 {wb['Hist']['E10'].value:,.0f}、新興業務（智能機器人＋數字能源）{wb['Hist']['E11'].value:,.0f}、其他 {wb['Hist']['E12'].value:,.0f} 百萬元。新興業務內機器人與數字能源的拆分未揭露，Assumed（機器人占 {I['rob_sh26']:.0%}，區間 40–75%）。
2. **量：** 智能機器人＝中國工業機器人裝機（線 3）× 份額漂移 × 單價年降；通用自動化＝公司成長路徑（FY2026 前 9 月 +33%）＋人形伺服增量（線 4 出貨 × 線 6 軸數 × 單價 × 匯川份額，基準 2030 年 {I['hs30_B']:.0%}）。
3. **價與毛利：** 毛利率 2025 年 28.95%，前 9 月 28.2%、第 3 季 25.5%（新能源車毛利率 14.5%、通用自動化約 44%）；情境 B FY2030 {I['gm30_B']:.1%}。

情境 B：營收 {ser('rev', '{:,.0f}')}；毛利率 {ser('gm', '{:.1%}')}；EPS {ser('eps', '{:.2f}')} 元（FY2026→FY2032）。

## 主要假設與可證偽指標（原則第 10 節第 20 條）

| 承重假設 | 基準（A／B／C 或區間） | 標記 | 可觀察指標 |
|---|---|---|---|
| 通用自動化 FY2027 年增 | {I['ga27_A']:.0%}／{I['ga27_B']:.0%}／{I['ga27_C']:.0%} | Assumed | 季度工控營收、PMI、國產替代 |
| 新能源車 FY2027 年增 | {I['nev27_A']:.0%}／{I['nev27_B']:.0%}／{I['nev27_C']:.0%} | Assumed | 聯合動力季報；公司展望 2026 年 230–240 億元 |
| FY2030 毛利率 | {I['gm30_A']:.1%}／{I['gm30_B']:.1%}／{I['gm30_C']:.1%} | Assumed | 季毛利率（第 3 季 25.5%） |
| 智能機器人份額漂移 | {I['drift_A']:+.0%}／{I['drift_B']:+.0%}／{I['drift_C']:+.0%} 每年 | Assumed | 中國機器人份額（公司稱 9.2%）、SCARA |
| 人形伺服份額（2030） | {I['hs30_A']:.0%}／{I['hs30_B']:.0%}／{I['hs30_C']:.0%} | Assumed（市場：線 4、線 6 Analogy） | 人形整機 BOM 揭露 |
| 同業前瞻 P/E | 中位數 {I['pe_peer']:.1f} 倍（6 家） | Derived | 同一估值日重抓 |

## 敏感度（整本重算，前 10 大）

| 案例 | 鍵 | 加權目標價 | 相對基準 | DCF 目標價 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | `{c['key']}` | {c['tp']:,.1f} | {c['tp'] / sens['base_tp'] - 1:+.1%} | {c['dcf']:,.1f} | {c['rating']} |" for c in top) + f"""

完整案例見 xlsx Sens 頁。**翻轉評等所需的變動**：見 Recon 反解；升一級需加權目標價達 {fn(up_b)} 元。

## 反解（Recon 頁）

""" + "\n".join(f"- {a}：{fn(b, '{:,.3f}')}（基準 {fn(c0, '{:,.3f}')}）；{d}" for a, b, c0, d in sol) + f"""

## 限制與「分不出來／沒找到」

- **分不出來（資料性質）：** 新興業務內機器人與數字能源拆分；人形伺服用量屬未發生事件。
- **沒找到（工作未完）：** 公司年報與三季報原文未直接讀取（歷史數字取自資料商與券商轉錄）；FY2023 分部營收；銷售、管理費用分項（營業費用以「毛利 − 營業利益」殘差建模）；前五大客戶；庫藏股；長期投資明細；日本同業 β；上銀等收盤檔有但未取前瞻 P/E 的同業。
- **線模型不一致：** 線模型無伺服、通用自動化、新能源車變數，故主體為 Assumed；線 6 以 7.1 元／美元換算，本模型用 6.6665。見 xlsx「對帳」頁。
- **同業倍數：** 日本同業與本公司取 2026-11-20 同日值；埃斯頓、美的取 2026-11-24 資料商頁面並換算至 2026-11-20 收盤；少於 8 家不做迴歸；埃斯頓前瞻 P/E 偏高（獲利基期低）。
- **β 偏低：** 本公司 β 0.52、同業取得者 0.43–0.82，Blume 調整後 {I['beta']:.2f}，WACC 偏低，DCF 終值占比高，DCF 結果對折現率敏感（見敏感度）。
- **rf 逾期：** 中國 10 年期國債 1.61% 為 2026-11-05 值（跨模型一致）。
- **2026 第 3 季公允價值收益：** 約 4.2 億元非經常性，FY2026 業外假設 5 億元，不外推。
- **獨立核對：** 依專案規則，交付判讀的數字應由未參與製作的代理獨立核對，待外層流程安排。

## 資料來源

見 xlsx Sources 頁（資料商財務頁、券商轉錄年報與三季報、線 3／4／6 成品、同業模型）。

## 版本紀錄

- {VER}（2026-11-24）：首版。現價 {I['px']:,.2f}（2026-11-20 收盤）；加權目標價 {g('tp'):,.1f}；上檔空間 {g('up'):+.1%}；{g('rating')}／{g('rec')}；收支關卡 {g('gate_b'):,.1f}（{g('gate_t')}）。
"""
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write(t)


if __name__ == "__main__":
    main()

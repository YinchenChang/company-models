# 兆威機電（003021.SZ，深圳證券交易所）個股模型 v0.1 建置腳本
# 產出：physical_ai/zhaowei/dist/20261124_兆威機電個股模型_v0.1.xlsx（LibreOffice 重算後的成品，含公式與快取值）
# 結構：命題模組（Thesis：靈巧手與機器人模組營收＝線 4 人形出貨 × 每台靈巧手數 × 線 6 單價 × 本公司份額；其餘為傳統微型傳動、精密零件）
#       → 官方產品分部（微型傳動系統／精密零件／其他業務）→ 標準損益表 → FCFF DCF＋同業 P/E。
# 敏感度：每個案例覆寫 In 頁基準欄後整本重建、LibreOffice 重算（整本重算，不用閉式近似）；反解同法以二分搜尋。
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md、docs/plan/高不確定性預估工作原則.md。
#       實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics 協調檔。
# 用法：python3 physical_ai/zhaowei/build_zhaowei.py   （需要 soffice；暫存於家目錄下 zw_build）
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
NAME = f"20261124_兆威機電個股模型_{VER}.xlsx"
OUT = os.path.join(HERE, "dist", NAME)
HOMEDIR = os.path.expanduser("~")
TMP = os.path.join(HOMEDIR, "zw_build")
LO = ["soffice", f"-env:UserInstallation=file://{HOMEDIR}/lo_zhaowei", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML"]

L4 = os.path.join(ROOT, "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx")
L6 = os.path.join(ROOT, "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx")
L4N, L6N = (os.path.relpath(p, ROOT) for p in (L4, L6))
L3N = L4N

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDRF = PatternFill("solid", fgColor="DDEBF7")
SECF = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------- 出處 ----------------
S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv：003021.SZ；未還原股利）"
S_TJ = "同花順 F10 兆威機電財務頁（https://basic.10jqka.com.cn/003021/finance.html，2026-11-24 讀取；資料商）：2023–2025 營收、淨利、毛利率、每股盈餘、2025 年底貨幣資金"
S_XY = "興業證券 兆威機電 2025 年報點評（2026-04-02；https://pdf.dfcfw.com/pdf/H3_AP202604031821007822_1.pdf）：轉錄公司 2025 年報；內文為壓縮檔，數字取自搜尋引擎摘要（費用率、折舊攤銷、經營現金流、毛利率）"
S_LX = "理杏仁 兆威機電營收構成頁（https://www.lixinger.com/equity/company/detail/sz/003021/3021/operation-revenue-constitution；資料商，取自搜尋摘要）：2025 年微型傳動系統 10.31 億元、精密零件 5.96 億元"
S_SZ = "證券時報報導與公司 2024 年報摘要（https://static.cninfo.com.cn/finalpage/2025-04-29/1223375606.PDF）：2024 營收 1,524,598,553 元、歸母淨利 225,094,345 元、總股本 240,203,500 股、研發 1.55 億元"
S_Q3 = "東方財富 2026-10-29 報導（https://finance.eastmoney.com/a/202610293913601568.html）與財問新聞（https://www.caiwennews.com/article/1674704.shtml）：轉錄 2026 年三季報（公司公告原文未直接讀取）"
S_GJ = "國金證券 兆威機電點評（2026-10-29；https://data.eastmoney.com/report/info/AP202610291831150145.html）：2026E–2028E 營收與歸母淨利"
S_HK = "公司 H 股上市公告（2026-03-09；https://pdf.dfcfw.com/pdf/H2_AN202603091820414641_1.pdf）與 Baker McKenzie 2026-03 新聞稿：發售 26,748,300 股、每股 71.28 港元、淨額約 18.28 億港元、上市後總股本 267,482,700 股"
S_RF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與 innolight、eoptolink、victory_giant 同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：China CRP 0.91%（與 innolight 同一取值）"
S_CLOSE = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
S_SAP = "資料商 stockanalysis.com statistics 頁（原報價日見 Comp 頁；同業原值轉引自本專案綠的諧波模型 Comp 頁）"
IP_BRK = "券商研究報告：有銷售研究與維持評等的誘因；歷史數字為轉錄，預測為券商自設"
IP_SA = "資料商彙整賣方共識，財報前的值未必更新"
IP_CO = "發行公司說法，有股價與融資誘因"
IP_MEDIA = "媒體引述私募人士說法，無公司揭露佐證"


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
    wb6, l6 = _line_rows(L6, 1)
    d = {"hum": {}}
    for y in range(2026, 2033):
        r, v = l4[f"SHIP_{y}"]
        for s, c, col in (("A", 4, "E"), ("B", 3, "D"), ("C", 2, "C")):
            d["hum"][(y, s)] = (v[c], f"{L4N} 下游取數!{col}{r}（SHIP_{y}）")
    ws6 = wb6["輸入"]
    hr = [r for r in range(1, ws6.max_row + 1) if ws6.cell(r, 1).value == "HAND"]
    # 線 6 輸入頁「HAND」列：前三列＝低／基準／高情境每台人形靈巧手數（F 欄），第四列＝2026 單價（D／E／F 欄）與年降幅（G／H／I 欄）
    d["n_hand"] = [(ws6.cell(r, 6).value, f"{L6N} 輸入!F{r}") for r in hr[:3]]
    pr = hr[3]
    d["p26"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{openpyxl.utils.get_column_letter(c)}{pr}") for c in (4, 5, 6)]
    d["pd"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{openpyxl.utils.get_column_letter(c)}{pr}") for c in (7, 8, 9)]
    r, v = l4["P_HAND"]; d["p_hand4"] = (v[3], f"{L4N} 下游取數!D{r}（P_HAND）", v[2], v[4])
    return d


LD = line_data()


# ---------------- In 頁輸入 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 來源誘因, 說明)
IN = [
    ("sec", "1. 市場資料與日期"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, "", "資料截止日 2026-11-24；股價 2026-11-20；成品日 2026-11-24"),
    ("px", "現價（2026-11-20 收盤，A 股）", "元", None, 76.83, None, "Verified", S_PX, "", "csv 值 76.830002；11-09 收盤 103.72，兩週跌 26%"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,12,31)", None, "Assumed", "本專案慣例：評價日 2026-12-31（會計年度＝日曆年，FY2026 在評價日已結束）", "", "折現期間由評價日起算"),
    ("tp_date", "目標時點（px_date 後 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", ""),
    ("bs_date", "資產負債表日", "日期", None, "=DATE(2026,9,30)", None, "Verified", S_Q3, "", "2026 年三季報；資產負債表明細未取得，淨現金為推算（見 cash_v）"),
    ("stub_frac", "Stub 期（資產負債表日→評價日）", "年", None, "=(val_date-bs_date)/365", None, "Derived", "＝(評價日 − 資產負債表日) ÷ 365", "", "Stub FCFF＝FY2026 FCFF × stub_frac"),
    ("sh_v", "評價日股數（總股本，A 股＋H 股）", "百萬股", None, 267.4827, None, "Verified", S_HK, "", "A 股 240,734,400＋H 股 26,748,300；限制性股票與選擇權未取得，以 0 計。兩地股份同權，A 股收盤價用於全部股本"),
    ("fx", "美元兌人民幣（2026-11-20）", "元／美元", None, 6.6665, None, "Verified", S_CLOSE + "：CNY=X", "", "靈巧手美元單價換算人民幣；H 股募資換算"),
    ("fxh", "美元兌港幣（2026-11-20）", "港幣／美元", None, 7.8426, None, "Verified", S_CLOSE + "：HKD=X", "", "H 股募資淨額換算人民幣（以報價日匯率近似，Derived）"),
    ("cash_v", "現金與短期投資（2026-09-30，推算）", "百萬元", 1400.0, "=217+1828*fx/fxh+158-92.7-150", 1800.0, "Derived",
     S_TJ + "（2025 年底貨幣資金 217）＋" + S_HK + "（淨額 18.28 億港元）＋" + S_Q3 + "（1–9 月營運現金流 158）；減 2025 年度股利約 92.7（每股 0.385 元 × 約 2.41 億股，見 payout）；減 1–9 月資本支出（無揭露值，假設 150）", "",
     "三季報資產負債表未讀取：以上為推算值，不是財報值；低／高 1,400／1,800（Assumed）。其他金融資產（理財、存單）未取得"),
    ("ltinv_v", "長期投資（2026-09-30）", "百萬元", None, 0.0, None, "Assumed", "未取得長期投資明細；基準以 0 計（保守）", "", "若有長期存單或股權投資，淨現金偏低"),
    ("inv_k", "長期投資計入淨現金的比例", "%", 0.5, 1.0, None, "Assumed", "長期投資基準為 0，本欄只為版面一致", "", ""),
    ("debt_v", "有息負債（2026-09-30）", "百萬元", 0.0, 50.0, 300.0, "Assumed", "無揭露值；2025 年財務費用率 −1.72%（淨利息收入，" + S_XY + "）顯示淨現金位置，基準取小額", "", ""),
    ("nci_v", "少數股東權益（2026-09-30）", "百萬元", None, 0.0, None, "Assumed", "未取得；靈巧手子公司是否有少數股東未查證", "", "自淨現金扣除"),
    ("cons_ni26", "資料商共識 2026 年歸母淨利（只作對照）", "百萬元", None, 308.0, None, "Interested-party", "東方財富 2026-10-29 報導：機構對 2026 全年歸母淨利一致預測 3.08 億元（日期不詳，疑未更新）；" + S_GJ + "同期給 1.6 億元，兩者差一倍", IP_SA, "只作對照，不進模型；賣方目標價未取得"),
    ("sec", "2. 實績錨（FY2026＝2026 年，會計年度＝日曆年）"),
    ("r9m26", "2026 年 1–9 月營收", "百萬元", None, 1262.0, None, "Verified", S_Q3, "", "公司公告原文未直接讀取；億元兩位小數轉錄"),
    ("g9m26", "2026 年 1–9 月營收年增", "%", None, 0.0053, None, "Verified", S_Q3, "", "Q3 單季 −4.94%（扣掃地機客戶需求放緩，國金）"),
    ("q4g", "2026 年第 4 季營收年增", "%", -0.05, 0.05, 0.12, "Assumed", "1–9 月 +0.5%、Q3 −4.9%；國金 2026E 營收 17.9 億（+4%）隱含 Q4 約 +10%（" + S_GJ + "）", "", "FY2026 營收＝1–9 月實績＋2025Q4 ×（1＋q4g）；2025Q4＝2025 全年 − 1–9 月 ÷（1＋g9m26）"),
    ("sh_t26", "FY2026 微型傳動系統 占營收", "%", None, "=Hist!E8/Hist!E4", None, "Derived", "＝2025 年占比（" + S_LX + "）", "", "2026 年只揭露上半年毛利率，未取得分部營收；沿用 2025 年結構"),
    ("sh_p26", "FY2026 精密零件 占營收", "%", None, "=Hist!E9/Hist!E4", None, "Derived", "＝2025 年占比", "", "其他業務＝餘數"),
    ("gm26", "FY2026 毛利率", "%", None, 0.3135, None, "Verified", S_Q3 + "（1–9 月 31.35%；Q3 單季 32.43%）；國金同為 31.35%", "", "第 4 季假設同 1–9 月"),
    ("opex26", "FY2026 銷售＋管理＋研發費用率", "%", 0.195, 0.215, 0.23, "Assumed", "2025 年 20.27%（銷售 3.87%＋管理 6.24%＋研發 10.16%，" + S_XY + "）；2026 年 1–9 月研發 10.6%，淨利率由 14.8% 降至 9.7%，費用率與減值上升", "", "三季報費用明細未取得；由淨利率回推"),
    ("sec", "3. 命題模組：靈巧手與機器人模組（量 × 價 × 份額；線 4、線 6）"),
    ("rob_sh26", "FY2026 機器人模組營收占合併營收", "%", 0.008, 0.014, 0.04, "Assumed", "公司未揭露機器人營收；私募人士稱占比約 1.4%（" + "東方財富 2026-10-29 報導" + "）；報導把「靈巧手收入由千萬級到億級」列為驗證訊號，顯示現為千萬級", IP_MEDIA, "占比 1.4% 對應約 24 百萬元；高案 4% 對應約 70 百萬元（億級以下）"),
    ("n_hand", "每台人形機器人靈巧手數（線 6）", "隻／台", LD["n_hand"][0][0], LD["n_hand"][1][0], LD["n_hand"][2][0], "Analogy", f"線 6：{LD['n_hand'][1][1]}（低 {LD['n_hand'][0][1]}、高 {LD['n_hand'][2][1]}）；線 4 計算頁同為兩隻", "", "兩線一致"),
    ("p_hand26", "2026 年靈巧手單價（線 6）", "美元／隻", LD["p26"][0][0], LD["p26"][1][0], LD["p26"][2][0], "Analogy", f"線 6：{LD['p26'][1][1]}（低 {LD['p26'][0][1]}、高 {LD['p26'][2][1]}）", "", f"線 4 P_HAND 基準 {LD['p_hand4'][0]:,.0f}（{LD['p_hand4'][2]:,.0f}–{LD['p_hand4'][3]:,.0f}），與線 6 不一致，見對帳；本模型採零組件線"),
    ("pd", "靈巧手單價年降幅（線 6）", "%／年", LD["pd"][0][0], LD["pd"][1][0], LD["pd"][2][0], "Analogy", f"線 6：{LD['pd'][1][1]}（低 {LD['pd'][0][1]}、高 {LD['pd'][2][1]}）", "", ""),
    ("hu_k", "人形出貨相對線 4 的倍數（全部情境）", "x", None, 1.0, 2.0, "Assumed", "基準＝線 4 原值；高案 2 倍：中國 2026 年人形產量可能逾 10 萬台（線 4 基準全球 5.5 萬台；群益證券引述，本輪未重查）", "", "反解變數"),
    ("g_ao", "其他業務 年增（FY2027 起）", "%", 0.0, 0.05, 0.15, "Assumed", "占營收約 5%，組成未取得（微型馬達、模具等，未查證）", "", ""),
    ("sec", "4. 命題模組：情境專屬驅動（A 樂觀／B 基準／C 保守；只列基準欄，不設敏感度區間）"),
    ("hs30_A", "情境 A：2030 年占全球靈巧手需求份額", "%", None, 0.04, None, "Assumed", "2026 隱含份額見 Thesis（基準約 1%）；因時機器人出貨領先（媒體）、整機廠自研風險高；份額未取得", "", "2026 隱含份額線性走到 2030，其後持平"),
    ("hs30_B", "情境 B：2030 年占全球靈巧手需求份額", "%", None, 0.02, None, "Assumed", "同上；基準取隱含份額約兩倍", "", ""),
    ("hs30_C", "情境 C：2030 年占全球靈巧手需求份額", "%", None, 0.007, None, "Assumed", "同上；基準取隱含份額約持平（整機廠自研、因時等競爭）", "", ""),
    ("gcore_A", "情境 A：傳統微型傳動與精密零件營收年增（FY2027 起）", "%", None, 0.14, None, "Assumed", "2023–2025 年營收年增 26%、12.5%；2026 年 1–9 月 +0.5%；國金 2027E 全公司 +16%（" + S_GJ + "，含機器人）", "", "汽車、消費電子（扶地機器人、手機鏡頭等）；兩分部同一增速"),
    ("gcore_B", "情境 B：傳統微型傳動與精密零件營收年增（FY2027 起）", "%", None, 0.08, None, "Assumed", "同上", "", ""),
    ("gcore_C", "情境 C：傳統微型傳動與精密零件營收年增（FY2027 起）", "%", None, 0.02, None, "Assumed", "同上；Q3 單季 −4.9%", "", ""),
    ("gm30_A", "情境 A：FY2030 毛利率", "%", None, 0.345, None, "Assumed", "2023 年 29.1%、2024 年 31.4%、2025 年 33.7%、2026 年 1–9 月 31.35%（Q3 32.4%）", "", "FY2026 起線性走到 FY2030，其後持平"),
    ("gm30_B", "情境 B：FY2030 毛利率", "%", None, 0.32, None, "Assumed", "同上", "", ""),
    ("gm30_C", "情境 C：FY2030 毛利率", "%", None, 0.29, None, "Assumed", "同上；上半年微型傳動毛利率 25.4%（資料商）", "", ""),
    ("tp_wA", "情境權重 A", "", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", ""),
    ("tp_wB", "情境權重 B", "", None, 0.50, None, "Assumed", "同上", "", ""),
    ("tp_wC", "情境權重 C", "", None, 0.25, None, "Assumed", "同上", "", ""),
    ("sec", "5. 損益與現金流"),
    ("ts_r", "稅金及附加 ÷ 營收", "%", 0.003, 0.006, 0.01, "Assumed", "年報明細未取得；A 股製造業常見 0.4–0.8%", "", ""),
    ("opex30", "FY2030 銷售＋管理＋研發費用率", "%", 0.17, 0.185, 0.20, "Assumed", "2025 年 20.3%；研發費用率約 10% 為策略性支出，營運槓桿有限", "", "FY2026→FY2030 線性，其後持平"),
    ("oi_r", "其他收益（政府補助等）÷ 營收", "%", 0.005, 0.01, 0.02, "Assumed", "年報明細未取得", "", "計入營業利益"),
    ("imp_r", "減值損失 ÷ 營收", "%", 0.002, 0.005, 0.01, "Assumed", "年報明細未取得；2026 年淨利率下滑幅度大於毛利率下滑，推測含減值", "", "計入營業利益"),
    ("inv_y", "現金與金融資產收益率", "%", 0.015, 0.02, 0.03, "Assumed", "2025 年財務費用率 −1.72%（" + S_XY + "）；利率下行", "", "業外；DCF 另以淨現金計入，不重複"),
    ("tax_r", "有效稅率", "%", 0.08, 0.12, 0.16, "Assumed", "年報稅費未取得；高新技術企業稅率 15% 並有研發加計扣除，取 12%", "", "敏感度低／高 8%／16%"),
    ("nci_r", "少數股東損益 ÷ 稅後淨利", "%", None, 0.0, None, "Assumed", "未取得", "", ""),
    ("payout", "配息率", "%", None, "=0.385/1.06", None, "Derived", "＝2025 年度每股股利 0.385 元（每 10 股 3.85 元，" + S_XY + "搜尋摘要）÷ EPS 1.06 元（" + S_TJ + "）", "", "用於收支關卡與目標時點前股利"),
    ("da26", "FY2026 折舊攤銷 ÷ 營收", "%", None, "=Hist!E22/Hist!E4", None, "Derived", "＝2025 年折舊攤銷 105.6 ÷ 營收（" + S_XY + "）", "", "之後線性收斂到資本支出比例 × da_ss（FY2032）"),
    ("da_ss", "穩態折舊 ÷ 資本支出", "x", None, 0.8, None, "Assumed", "穩態資本支出高於折舊（成長）", "", ""),
    ("cx26", "FY2026 資本支出", "百萬元", 150, 220, 320, "Assumed", "未取得；H 股募資用於產能與研發（公告未逐項讀取）；2025 年折舊 106", "", ""),
    ("cx_r", "資本支出 ÷ 營收（FY2027 起）", "%", 0.07, 0.10, 0.15, "Assumed", "歷史資本支出未取得；折舊占營收 6.2%", "", ""),
    ("nwc_r", "營運資金增加 ÷ 營收增量", "%", 0.15, 0.25, 0.40, "Assumed", "未取得營運資金明細；製造業常見", "", ""),
    ("h_iss", "再融資占擴大後股本（尾端事件，基準 0）", "%", None, 0.0, 0.05, "Assumed", "H 股已於 2026-03 發行（" + S_HK + "）；再融資無公告，只作尾端情境", "", "原則第 10 節第 19 條：事件型尾端風險進 Sens、不進基準"),
    ("h_disc", "再融資價格相對現價折價", "%", None, 0.10, None, "Assumed", "慣例折價（未查核，不作依據）", "", "只在 h_iss＞0 時有作用"),
    ("sec", "6. 折現率（DCF 標準規格第 2 節）"),
    ("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", S_RF, "", "日期 2026-11-05，較股價基準日早 11 個交易日（逾期；跨模型一致，下次全面改版更新）"),
    ("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", ""),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", ""),
    ("b_own", "本公司 5 年 β（原始）", "x", None, 0.44, None, "Interested-party", S_SAP + "：兆威 β 0.44（2026-09-30 報價頁）", IP_SA, "A 股小型股 β 偏低，樣本窗口內與大盤相關性低"),
    ("b_peer", "同業 5 年 β 中位數（原始）", "x", None, "=MEDIAN(Comp!H5:H10)", None, "Derived", "＝Comp 頁核心同業 6 家 β 中位數", "", ""),
    ("beta", "β（兩來源 Blume 調整後中位數）", "x", "=beta-0.3", "=MEDIAN(0.67*b_own+0.33,0.67*b_peer+0.33)", "=beta+0.3", "Derived", "Blume 0.67β＋0.33；兩來源＝本公司、同業中位數", "", "敏感度 ∓0.3（跨模型慣例）"),
    ("kd", "稅前債務成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt", "", "幾無有息負債"),
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
SCN_DRIVERS = [("hs30", "2030 年靈巧手需求份額"), ("gcore", "傳統微型傳動與精密零件營收年增"), ("gm30", "FY2030 毛利率")]


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
    HI["A1"] = "Hist：歷史財務（百萬元人民幣；中國會計準則）。券商與資料商轉錄之年報數字；空白或「未取得」＝本輪未取得（年報全文為影像檔，未能讀取）。公式只引用本頁（定義名稱 h_*）。"
    hdr(HI, 3, ["項目", "", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"])
    NA = "未取得"
    HROWS = [  # (列, 鍵, 項目, [2023, 2024, 2025], 標記, 來源)
        (4, "h_rev", "營業收入", ["=D4/1.2642", 1524.60, 1716.0], "Verified", S_SZ + "（2024，+26.42%；2023＝2024 ÷ 1.2642，Derived）；" + S_TJ + "（2025：17.16 億，+12.5%；2023：12.06 億）"),
        (5, "h_cogs", "營業成本（＝營收 ×（1 − 毛利率））", ["=C4*(1-0.2905)", "=D4*(1-0.3143)", "=E4*(1-0.3366)"], "Derived", "毛利率 29.05%、31.43%、33.66%（" + S_TJ + "；2025 年與" + S_XY + "一致）"),
        (6, "h_gp", "毛利（＝營收 − 營業成本）", ["=C4-C5", "=D4-D5", "=E4-E5"], "Derived", "＝營業收入 − 營業成本"),
        (7, "h_gm", "毛利率", ["=C6/C4", "=D6/D4", "=E6/E4"], "Derived", "＝毛利 ÷ 營收"),
        (8, "h_seg_t", "微型傳動系統 營收", [NA, NA, 1031.0], "Interested-party", S_LX + "（2025：10.31 億，占 60.10%，毛利率 28.67%）；2023、2024 分部未取得"),
        (9, "h_seg_p", "精密零件 營收", [NA, NA, 596.0], "Interested-party", S_LX + "（2025：5.96 億，占 34.72%，毛利率 42.49%）"),
        (10, "h_seg_a", "其他業務 營收（＝營收 − 上兩項）", [NA, NA, "=E4-E8-E9"], "Derived", "殘差，約 5%；含微型馬達等，組成未取得"),
        (11, "h_sell", "銷售費用（＝營收 × 3.87%）", [NA, NA, "=0.0387*E4"], "Derived", S_XY + "：銷售費用率 3.87%"),
        (12, "h_admin", "管理費用（＝營收 × 6.24%）", [NA, NA, "=0.0624*E4"], "Derived", S_XY + "：管理費用率 6.24%"),
        (13, "h_rnd", "研發費用", [NA, 155.0, "=0.1016*E4"], "Derived", S_SZ + "（2024：1.55 億，占 10.18%）；" + S_XY + "（2025 研發費用率 10.16%）"),
        (14, "h_opex", "營業費用（銷售＋管理＋研發）", [NA, NA, "=SUM(E11:E13)"], "Derived", "＝三項合計；2023、2024 銷售與管理費用未取得"),
        (15, "h_fin", "財務費用（負＝淨利息收入；＝營收 × −1.72%）", [NA, NA, "=-0.0172*E4"], "Derived", S_XY + "：財務費用率 −1.72%；不進 PL 營業利益"),
        (16, "h_op", "營業利益（核心：毛利 − 三項費用；2025 未含稅金及附加與其他損益）", [NA, NA, "=E6-E14"], "Derived", "年報利潤表全文未讀取，稅金及附加、其他收益、減值、投資收益未取得；此值為近似"),
        (17, "h_opm", "營業利益率（核心）", [NA, NA, "=E16/E4"], "Derived", ""),
        (18, "h_ni", "歸屬母公司淨利", ["=D18/1.2511", 225.09, 254.29], "Verified", S_SZ + "（2024：225.09 百萬，+25.11%；2023＝2024 ÷ 1.2511，Derived）；" + S_XY + "（2025：2.54 億，+13%；財務預測表 254.29）"),
        (19, "h_eps", "EPS（基本，元）", [0.75, 0.94, 1.06], "Verified", S_TJ),
        (20, "h_dps", "每股股利（元）", [NA, NA, 0.385], "Interested-party", "2025 年報利潤分配預案：每 10 股派現 3.85 元（含稅；搜尋摘要轉述）"),
        (21, "h_sh", "股數（＝淨利 ÷ EPS，百萬股；上市 H 股前）", ["=C18/C19", "=D18/D19", "=E18/E19"], "Derived", "約 240 百萬股，與 2024 年報摘要總股本 240.20 百萬股相符"),
        (22, "h_da", "折舊攤銷", [NA, NA, 105.63], "Interested-party", S_XY + "（財務預測表）"),
        (23, "h_ocf", "營業現金流", [NA, 123.68, 159.41], "Interested-party", S_XY + "（財務預測表）"),
    ]
    for (rr, k, lab, vals, tag, src) in HROWS:
        HI.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = HI.cell(rr, 3 + j, v)
                if not (isinstance(v, str) and (v.startswith("=") or v == NA)): c.font = BLUE
                if k in ("h_gm", "h_opm"): c.number_format = "0.0%"
        HI.cell(rr, 6, tag); HI.cell(rr, 7, src)
        dn(wb, k, f"Hist!$C${rr}:$E${rr}")
    HI.cell(26, 1, "檢查：分部合計 − 營業收入（FY2025，應為 0）")
    HI["E26"] = "=E8+E9+E10-E4"
    HI.column_dimensions["A"].width = 58; HI.column_dimensions["G"].width = 90
    H25 = lambda rr: (lambda y: f"=Hist!E{rr}" if y == 2025 else "未取得")
    HM = lambda y: "未取得"

    # ================= Q（季度；一列一個數字，單一標記）=================
    Q["A1"] = "Q：2026 年上半年與第 3 季資料（百萬元）。一列一個項目、單一主要標記；媒體轉錄公司三季報（原文未直接讀取）；空白＝未取得。"
    hdr(Q, 3, ["項目", "2026Q1", "2026Q2", "2026H1", "2026Q3", "2026 年 1–9 月", "標記", "來源", "來源誘因", "說明"])
    QROWS = [
        (4, "營業收入", [None, None, 816, 445, "=D4+E4"], "Verified", S_Q3 + "；上半年 8.16 億（東方財富 2026-08-15 報導）", "", "1–9 月合計與 In 頁 r9m26（12.62 億）差為取整；Q1、Q2 單季未取得"),
        (5, "營業收入年增", [None, None, 0.038, -0.0494, 0.0053], "Verified", "同上", "", "Q3 年減 4.9%"),
        (6, "毛利率", [None, None, 0.3076, 0.3243, 0.3135], "Verified", S_Q3 + "；上半年毛利率 30.76%（資料商）；國金：Q3 32.43%、1–9 月 31.35%", "", ""),
        (7, "歸屬母公司淨利", [None, None, "=F7-E7", 47.95, 123.0], "Verified", S_Q3 + "（1–9 月 1.23 億；Q3 4,794.87 萬；上半年為差額）", "", ""),
        (8, "歸屬母公司淨利年增", [None, None, None, -0.2944, -0.3235], "Verified", S_Q3, "", "扣非淨利 9,673 萬（−39.63%）"),
        (9, "研發費用", [None, None, None, None, 134.0], "Verified", S_Q3 + "（1.34 億，+4.8%，占營收 10.6%）", "", ""),
        (10, "營業現金流", [None, None, None, None, 158.0], "Verified", S_Q3 + "（1.58 億，+54.67%）", "", ""),
        (11, "營收年增（國金 2026E 全年）", [None, None, None, None, 0.04], "Interested-party", S_GJ + "（2026E 17.9 億，+4%）", IP_BRK, "全年預測，只作對照"),
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

    # ================= 線取數（線模型下游取數；藍字＝讀入值，H 欄寫來源檔與儲存格）=================
    LN["A1"] = "線取數：本模型讀入的線層級變數年度路徑（建置時由線模型成品依 ID 讀值）。情境對應：A＝線模型高情境、B＝基準、C＝低情境。"
    hdr(LN, 3, ["項目", "情境", ""] + [f"FY{y}" for y in range(2025, 2033)] + ["單位", "來源（檔案與儲存格）"])
    LR = {}
    r = 4
    for key, lab, unit, dd, y0 in (("hum", "全球人形機器人出貨（線 4）", "台", LD["hum"], 2026),):
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

    # ================= Thesis（命題模組）=================
    TH["A1"] = ("命題模組：微型傳動系統營收＝靈巧手與機器人模組（FY2026 錨；其後＝線 4 全球人形出貨 × 線 6 每台靈巧手數 × 線 6 單價 × 本公司份額）＋傳統微型傳動（汽車、消費電子等；年增 gcore）；"
                "精密零件與其他業務各自成長。FY2026＝1–9 月實績＋第 4 季假設；單價＝2026 單價 ×（1 − 線 6 年降幅）^(年 − 2026)。百萬元。")
    TH["A2"] = "欄：D＝FY2025 … K＝FY2032。靈巧手 2026 隱含份額以線 4 基準出貨校準（三情境同一實績年）；機器人模組營收可能含靈巧手以外的模組，故份額為上界。"
    hdr(TH, 3, ["項目", "說明", ""] + [f"FY{y}" for y in range(2025, 2033)])
    THR = {}
    r = 4
    for s in SC:
        TH.cell(r, 1, f"情境 {s}").font = BOLD; TH.cell(r, 1).fill = SECF; r += 1
        R = {}
        R["rev26"] = r; TH.cell(r, 1, "FY2026 合併營收（錨）"); TH.cell(r, 2, "＝1–9 月＋2025Q4 ×（1＋q4g）")
        TH[f"{TC[2026]}{r}"] = "=r9m26+(Hist!E4-r9m26/(1+g9m26))*(1+q4g)"; r += 1
        R["asp"] = r; TH.cell(r, 1, "靈巧手單價（元／隻）"); TH.cell(r, 2, "＝線 6 單價（美元）× 匯率 ×（1 − 年降）^(年 − 2026)")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = "=p_hand26*fx" if y == 2026 else f"={TC[y - 1]}{r}*(1-pd)"
        r += 1
        R["ship"] = r; TH.cell(r, 1, "全球人形出貨（線 4 × 出貨倍數，千台）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR[('hum', s)]}*hu_k/1000"
        r += 1
        R["hdem"] = r; TH.cell(r, 1, "全球靈巧手市場（百萬元）"); TH.cell(r, 2, "＝出貨 × 每台靈巧手數 × 單價")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['ship']}*n_hand*{TC[y]}{R['asp']}/1000"
        r += 1
        R["drv26"] = r; TH.cell(r, 1, "FY2026 微型傳動系統（錨）"); TH[f"{TC[2026]}{r}"] = f"={TC[2026]}{R['rev26']}*sh_t26"; r += 1
        R["hsh"] = r; TH.cell(r, 1, "靈巧手全球份額"); TH.cell(r, 2, "FY2026＝隱含（以線 4 基準校準）；線性走到 hs30，其後持平")
        TH[f"{TC[2026]}{r}"] = f"={TC[2026]}{R['rev26']}*rob_sh26/(線取數!{TC[2026]}{LR[('hum', 'B')]}*hu_k/1000*n_hand*{TC[2026]}{R['asp']}/1000)"
        for y in range(2027, 2033):
            TH[f"{TC[y]}{r}"] = f"=${TC[2026]}${r}+(hs30_{s}-${TC[2026]}${r})*MIN(1,({y}-2026)/4)"
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 1
        R["hum"] = r; TH.cell(r, 1, "微型傳動：靈巧手與機器人模組 營收"); TH.cell(r, 2, "FY2027 起＝市場 × 份額")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = f"={c}{R['rev26']}*rob_sh26" if y == 2026 else f"={c}{R['hdem']}*{c}{R['hsh']}"
        r += 1
        R["ind"] = r; TH.cell(r, 1, "微型傳動：非機器人用途 營收"); TH.cell(r, 2, "FY2026＝微型傳動系統 − 機器人模組；其後年增 gcore")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = f"={c}{R['drv26']}-{c}{R['hum']}" if y == 2026 else f"={TC[y - 1]}{r}*(1+gcore_{s})"
        r += 1
        R["harm"] = r; TH.cell(r, 1, "微型傳動系統 合計")
        for y in range(2026, 2033): c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['ind']}+{c}{R['hum']}"
        r += 1
        R["mech"] = r; TH.cell(r, 1, "精密零件")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['rev26']}*sh_p26" if y == 2026 else f"={TC[y - 1]}{r}*(1+gcore_{s})"
        r += 1
        R["ao"] = r; TH.cell(r, 1, "其他業務")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['rev26']}*(1-sh_t26-sh_p26)" if y == 2026 else f"={TC[y - 1]}{r}*(1+g_ao)"
        r += 1
        R["humsh"] = r; TH.cell(r, 1, "機器人模組占合併營收")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['hum']}/({c}{R['harm']}+{c}{R['mech']}+{c}{R['ao']})"; TH[f"{c}{r}"].number_format = "0.0%"
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
    PL["A1"] = ("PL：官方產品分部（微型傳動系統／精密零件／其他業務）→ 合併營收 → 標準損益表 → FCFF。"
                "C＝FY2023 … E＝FY2025 為實績（未取得之項目標「未取得」）；F＝FY2026（1–9 月實績＋第 4 季假設）… L＝FY2032。百萬元；中國會計準則（法定財報口徑，GAAP）。")
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
        row("ind", "微型傳動：非機器人用途（子列）", "百萬元", None, lambda y: f"=Thesis!{TC[y]}{th['ind']}")
        row("hum", "微型傳動：靈巧手與機器人模組（子列）", "百萬元", None, lambda y: f"=Thesis!{TC[y]}{th['hum']}")
        row("harm", "微型傳動系統 營收", "百萬元", H25(8), lambda y: f"=SUM({C(y)}{R['ind']}:{C(y)}{R['hum']})")
        row("mech", "精密零件 營收", "百萬元", H25(9), lambda y: f"=Thesis!{TC[y]}{th['mech']}")
        row("ao", "其他業務 營收", "百萬元", H25(10), lambda y: f"=Thesis!{TC[y]}{th['ao']}")
        row("rev", "營收（合併）", "百萬元", H(4), lambda y: f"=SUM({C(y)}{R['harm']}:{C(y)}{R['ao']})")
        row("chk", "檢查：分部合計 − 合併營收（應為 0）", "百萬元",
            lambda y: f"={C(y)}{R['harm']}+{C(y)}{R['mech']}+{C(y)}{R['ao']}-{C(y)}{R['rev']}" if y == 2025 else None, lambda y: f"={C(y)}{R['harm']}+{C(y)}{R['mech']}+{C(y)}{R['ao']}-{C(y)}{R['rev']}")
        PL.cell(R["chk"], 1, "檢查：分部合計 − 合併營收（應為 0）" if s == "B" else f"檢查（情境 {s}）：分部合計 − 合併營收（應為 0）")
        row("chk2", "檢查：微型傳動子列合計 − 微型傳動分部（應為 0）", "百萬元", None, lambda y: f"={C(y)}{R['ind']}+{C(y)}{R['hum']}-Thesis!{TC[y]}{th['harm']}")
        row("yoy", "營收年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1",
            lambda y: f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1", "0.0%")
        row("cogs", "營業成本", "百萬元", H(5), lambda y: f"={C(y)}{R['rev']}*(1-Thesis!{TC[y]}{th['gm']})")
        row("gp", "毛利", "百萬元", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}")
        row("gm", "毛利率", "%", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", "0.0%")
        row("ts", "稅金及附加", "百萬元", HM, lambda y: f"={C(y)}{R['rev']}*ts_r")
        row("opex", "營業費用（銷售＋管理＋研發）", "百萬元", H25(14), lambda y: f"={C(y)}{R['rev']}*Thesis!{TC[y]}{th['opex']}")
        row("othop", "其他營業損益（其他收益 − 減值等）", "百萬元", HM, lambda y: f"={C(y)}{R['rev']}*(oi_r-imp_r)")
        row("op", "營業利益（核心，不含金融資產收益）", "百萬元", H25(16), lambda y: f"={C(y)}{R['gp']}-{C(y)}{R['ts']}-{C(y)}{R['opex']}+{C(y)}{R['othop']}")
        row("opm", "營業利益率", "%", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}" if y == 2025 else "未取得", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", "0.0%")
        row("nonop", "其他收入淨額（業外：投資收益、公允價值、財務費用、營業外）", "百萬元", HM, lambda y: "=(cash_v+ltinv_v)*inv_y")
        row("pbt", "稅前利益", "百萬元", HM, lambda y: f"={C(y)}{R['op']}+{C(y)}{R['nonop']}")
        row("tax", "所得稅", "百萬元", HM, lambda y: f"={C(y)}{R['pbt']}*tax_r")
        row("nit", "稅後淨利（含少數股東）", "百萬元", HM, lambda y: f"={C(y)}{R['pbt']}-{C(y)}{R['tax']}")
        row("nci", "少數股東損益", "百萬元", HM, lambda y: f"={C(y)}{R['nit']}*nci_r")
        row("ni", "歸屬母公司淨利", "百萬元", H(18), lambda y: f"={C(y)}{R['nit']}-{C(y)}{R['nci']}")
        row("nim", "淨利率", "%", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", "0.0%")
        row("niy", "淨利年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1",
            lambda y: f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1", "0.0%")
        row("sh", "股數（百萬股）", "百萬股", H(21), lambda y: "=sh_v/(1-h_iss)")
        row("eps", "EPS（元）", "元", H(19), lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['sh']}", "0.00")
        row("da", "折舊攤銷", "百萬元", H25(22), lambda y: f"={C(y)}{R['rev']}*(da26+(cx_r*da_ss-da26)*{y - 2026}/6)")
        row("cx", "資本支出", "百萬元", HM, lambda y: "=cx26" if y == 2026 else f"={C(y)}{R['rev']}*cx_r")
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
    CP["A1"] = ("Comp：同業前瞻 P/E（資料商 stockanalysis.com statistics，轉引自本專案綠的諧波模型 Comp 頁）。股價＝" + S_CLOSE + "（Verified）；前瞻 P/E＝原值 × 2026-11-20 收盤 ÷ 原報價（Derived，EPS 不變）；"
                "哈默納科、Nabtesco 取本專案模型同日換算值。核心同業 6 家進中位數，少於 8 家不做迴歸。報價日未取得：中大力德、五洲新春無前瞻 P/E；其他收盤檔內無報價的同業未納入。")
    hdr(CP, 4, ["公司", "代號", "股價（2026-11-20）", "報價日", "前瞻 P/E 分母", "前瞻 P/E", "近 12 月 P/E（原報價日）", "5 年 β", "會計口徑", "進中位數", "說明", "原報價", "原報價日", "原前瞻 P/E", "標記與換算"])
    PEERS = [  # (公司, 代號, 11-20 收盤, 原報價, 原報價日, 原前瞻 P/E, 近 12 月 P/E, β, 口徑, 進中位數, 說明, 分母)
        ("雙環傳動", "002472.SZ", 35.45, 41.21, "2026-10-30", 22.32, 27.11, 0.52, "中國會計準則", "是", "齒輪、精密減速器（RV）；人形關節", "資料商 NTM 共識 EPS"),
        ("拓普集團", "601689.SS", 48.11, 46.52, "2026-09-30", 26.89, 32.10, 0.80, "中國會計準則", "是", "汽車零件、人形致動器", "資料商 NTM 共識 EPS"),
        ("鳴志電器", "603728.SS", 44.52, 51.50, "2026-09-30", 124.10, 302.24, 0.62, "中國會計準則", "是", "步進與空心杯馬達", "資料商 NTM 共識 EPS"),
        ("哈默納科", "6324.T", 3722, 3722, "2026-11-20", 48.94, None, 0.96, "日本基準（未查核）", "是", "諧波減速機；取 physical_ai/harmonic_drive 模型 Comp!F14", "公司 FY2026 指引 EPS（2027年3月期）"),
        ("Nabtesco", "6268.T", 3105, 3105, "2026-11-20", 13.77, None, 0.93, "IFRS（未查核）", "是", "RV 減速機；取 physical_ai/nabtesco 模型 In 頁 pe_self", "資料商 NTM 共識 EPS（11-13 值換算）"),
        ("綠的諧波", "688017.SS", 230.80, 263.76, "2026-10-09", 225.96, 342.80, 0.79, "中國會計準則", "是", "諧波減速器；人形選擇權定價，倍數偏高；中位數對極端值不敏感", "資料商 NTM 共識 EPS（三季報前）"),
        ("中大力德", "002896.SZ", 48.36, 56.87, "2026-10-09", None, 259.22, 0.25, "中國會計準則", "否", "諧波、行星減速器；資料商無前瞻 P/E（報價日未取得前瞻值）", "—"),
        ("五洲新春", "603667.SS", 41.60, 44.45, "2026-10-13", None, 221.31, 0.86, "中國會計準則", "否", "軸承、行星滾柱絲杠；資料商無前瞻 P/E", "—"),
        ("兆威機電（本公司，對照）", "003021.SZ", 76.83, 70.25, "2026-09-30", 59.70, 87.15, 0.44, "中國會計準則", "否", "本公司；只作對照", "資料商 NTM 共識 EPS（三季報前）"),
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
    CP.cell(15, 1, "核心同業前瞻 P/E 中位數"); CP["F15"] = "=MEDIAN(F5:F10)"
    CP.cell(16, 1, "本公司現價 ÷ 資料商前瞻 EPS（換算至 2026-11-20；分母：資料商 NTM 共識，三季報前）"); CP["F16"] = "=F13"
    CP.cell(17, 1, "本公司現價 ÷ 國金 FY2027 EPS（分母：國金 2027E 歸母淨利 210 百萬元 ÷ 總股本）"); CP["F17"] = "=px/(210/sh_v)"
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
        SE.cell(r + 1, 1, "說明：折舊、資本支出（FY2027 起）、營運資金隨營收等比例變動（第 44 條）；單價年降幅作用於靈巧手。")
    SE.column_dimensions["A"].width = 52; SE.column_dimensions["B"].width = 44

    # ================= Recon =================
    RC["A1"] = "Recon：模型對實績、券商預測的同口徑對照；反解支持現價所需的主導變數。"
    hdr(RC, 3, ["項目", "模型", "對照", "差距（模型 − 對照）", "說明"])
    rr = 4
    def rrow(lab, a, b, note, fmt="#,##0.0", diff=True):
        nonlocal rr
        RC.cell(rr, 1, lab); RC.cell(rr, 2, a).number_format = fmt
        if b is not None: RC.cell(rr, 3, b).number_format = fmt
        if diff and b is not None: RC.cell(rr, 4, f"=B{rr}-C{rr}").number_format = fmt
        RC.cell(rr, 5, short(note)); rr += 1
    rrow("FY2026 營收（三情境同）vs 國金 2026E 營收 1,790", f"=PL!F{PR['B']['rev']}", 1790, "國金 2026-10-29（三季報後）；券商預測為 Interested-party")
    rrow("FY2027 營收（情境 B）vs 國金 2027E 營收 2,080", f"=PL!G{PR['B']['rev']}", 2080, "國金 +16%；含機器人。本模型基準較低，因 2026 年 1–9 月營收持平")
    rrow("FY2028 營收（情境 B）vs 國金 2028E 營收 2,580", f"=PL!H{PR['B']['rev']}", 2580, "國金 +24%")
    rrow("FY2026 歸屬母公司淨利 vs 國金 2026E 160", f"=PL!F{PR['B']['ni']}", 160, "國金 1.6 億元（−37%）；1–9 月已實現 123，隱含第 4 季僅 37")
    rrow("FY2026 歸屬母公司淨利 vs 資料商共識 308", f"=PL!F{PR['B']['ni']}", "=cons_ni26", "共識 3.08 億元（報導引述，日期不詳）與國金差一倍；1–9 月僅 123，3.08 億需第 4 季 185，不合理，疑為舊值")
    rrow("FY2027 歸屬母公司淨利（情境 B）vs 國金 2027E 210", f"=PL!G{PR['B']['ni']}", 210, "國金 +29%")
    rrow("FY2028 歸屬母公司淨利（情境 B）vs 國金 2028E 310", f"=PL!H{PR['B']['ni']}", 310, "國金 +50%")
    rrow("FY2027 EPS（情境 B）vs 國金 2027E 淨利 ÷ 總股本", f"=PL!G{PR['B']['eps']}", "=210/sh_v", "", "0.00")
    rrow("FY2026 毛利率 vs 2026 年 1–9 月實績", f"=PL!F{PR['B']['gm']}", "=gm26", "模型以 1–9 月為錨（差距應為 0）", "0.0%")
    rrow("FY2026 第 4 季隱含營收 vs 2026Q3 實績", f"=PL!F{PR['B']['rev']}-r9m26", "=Q!E4", "第 4 季若維持第 3 季水準即為對照值")
    rrow("FY2026 隱含第 4 季歸屬母公司淨利 vs 2026Q3 實績", f"=PL!F{PR['B']['ni']}-Q!F7", "=Q!E7", "模型 FY2026 淨利 − 1–9 月 123")
    rrow("FY2026 靈巧手與機器人模組營收（百萬元）vs 媒體引述 1.4% × 營收", f"=Thesis!E{THR['B']['hum']}", "=0.014*PL!F" + str(PR['B']['rev']), "媒體引述私募人士說法（Interested-party），本模型基準即取該值", "#,##0.0")
    rrow("FY2026 靈巧手隱含全球份額（以線 4 基準校準）", f"=Thesis!E{THR['B']['hsh']}", None, "＝機器人模組營收 ÷（線 4 出貨 × 每台靈巧手數 × 線 6 單價）；機器人模組可能含靈巧手以外，故為上界", "0.0%", False)
    rrow("FY2030 靈巧手與機器人模組營收（情境 B，百萬元）", f"=Thesis!I{THR['B']['hum']}", None, "＝全球靈巧手市場 × 份額（單價年降，市場以人民幣計增速低於出貨）", "#,##0.0", False)
    rrow("靈巧手 2026 單價（美元）：線 6 基準 vs 線 4 P_HAND", "=p_hand26", LD["p_hand4"][0], f"線 6：{LD['p26'][1][1]}；線 4：{LD['p_hand4'][1]}；兩線不一致，本模型採線 6（零組件線）。改用線 4 的值，隱含份額與營收預測會改變，見 Sens 的線 4 單價案例", "#,##0")
    rr += 1
    RC.cell(rr, 1, "反解（整本重算二分搜尋；建置時計算，數值為快照）").font = BOLD; rr += 1
    hdr(RC, rr, ["反解目標", "解", "基準值", "", "說明"]); rr += 1
    if sens:
        for lab, v, b0, note in sens["solve"]:
            RC.cell(rr, 1, lab); RC.cell(rr, 2, v).number_format = "#,##0.00"; RC.cell(rr, 3, b0).number_format = "#,##0.00"; RC.cell(rr, 5, note); rr += 1
    RC.column_dimensions["A"].width = 70; RC.column_dimensions["E"].width = 90

    # ================= 對帳 =================
    DZ["A1"] = "對帳：本模型使用的「線層級」變數（與 physical_ai/lines/ 各線模型對帳；本模型讀線 v0.1 成品，不靜默改口徑）。傳統微型傳動（汽車、消費電子）不讀線變數，為本模型自設，見 In 頁 gcore。"
    hdr(DZ, 3, ["變數名", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型是否一致"])
    h26 = LD["hum"][(2026, "B")]; h32 = LD["hum"][(2032, "B")]
    ZR = [
        ("全球人形機器人出貨 FY2026–FY2032（A／B／C＝線高／基準／低）", "4 人形", f"基準 {h26[0]:,.0f} → {h32[0]:,.0f}（× hu_k）", "台", "全球、含小型機", f"{h26[1]}；{h32[1]}", "Assumed", "Thesis 靈巧手市場", "一致（hu_k＝1）；高案 hu_k＝2 與線 4 基準不一致，列 Sens 不改基準"),
        ("每台人形機器人靈巧手數", "6 共用零組件", "=n_hand", "隻／台", "每台兩隻（低案 1.5）", f"{LD['n_hand'][1][1]}（低 {LD['n_hand'][0][1]}、高 {LD['n_hand'][2][1]}）；線 4 計算頁同為兩隻", "Analogy", "Thesis 全球靈巧手市場", "一致"),
        ("靈巧手 2026 單價", "6 共用零組件／4 人形", "=p_hand26", "美元／隻", "線 6 基準 3,500（2,000–4,750）", f"{LD['p26'][1][1]}；線 4：{LD['p_hand4'][1]}（{LD['p_hand4'][0]:,.0f}，區間 {LD['p_hand4'][2]:,.0f}–{LD['p_hand4'][3]:,.0f}）", "Analogy", "Thesis 靈巧手單價", "不一致：線 6 基準 3,500、線 4 基準 4,750（線 4 區間低端 2,100）。本模型採線 6；兩線互不一致屬線模型待處理事項，不在本輪修改"),
        ("靈巧手單價年降幅", "6 共用零組件", "=pd", "%／年", "線 6", f"{LD['pd'][1][1]}（低 {LD['pd'][0][1]}、高 {LD['pd'][2][1]}）", "Analogy", "Thesis 單價路徑", "一致"),
        ("兆威靈巧手份額", "6 共用零組件", "未取得", "%", "份額頁第 19 列", f"{L6N} 份額頁（因時機器人出貨領先，兆威份額未取得；Interested-party 媒體）", "Interested-party", "hs30_*（Assumed，非讀值）", "本模型份額為自設，Recon 列出 2026 隱含值"),
        ("傳統微型傳動、精密零件年增", "（無）", "A／B／C＝14%／8%／2%", "%／年", "本模型自設", "In 頁 gcore_A／B／C（Assumed）", "Assumed", "Thesis 非機器人用途與精密零件", "線模型沒有對應變數；汽車與消費電子需求未讀線模型"),
    ]
    for i, z in enumerate(ZR):
        for j, v in enumerate(z): DZ.cell(4 + i, 1 + j, v)
    DZ.column_dimensions["A"].width = 50; DZ.column_dimensions["F"].width = 80; DZ.column_dimensions["I"].width = 70

    # ================= Risk =================
    RK["A1"] = "Risk：風險登錄（原則第 10 節第 10、19 條）。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    RISKS = [
        ("整機廠自研靈巧手、因時機器人等競爭", "是（hs30_A／B／C、rob_sh26）", "供應格局未定型；兆威份額未取得", "兆威靈巧手收入（2026 年報）、客戶揭露"),
        ("傳統業務下滑（掃地機、汽車、消費電子客戶需求）", "是（gcore、gm30_C、q4g）", "Q3 單季營收 −4.9%、淨利 −29%", "季營收、季毛利率"),
        ("費用與減值上升（淨利率 14.8%→9.7%）", "部分（opex26、imp_r）", "三季報費用與減值明細未取得", "2026 年報利潤表"),
        ("客戶集中度（前五大客戶占比）", "否", "年報客戶明細未取得（檔案記載亦查無）", "年報前五大客戶"),
        ("人形量產延後或不及線 4 基準", "是（情境 C＝線 4 低）", "", "機構季度出貨；中國人形產量"),
        ("H 股募資運用（擴產、研發）與資本支出", "部分（cx26、cx_r）", "資本支出歷史與募資運用明細未取得", "半年報在建工程、現金流量表"),
        ("A／H 價差與再融資", "部分（h_iss 尾端 Sens，基準 0）", "H 股 2026-03 以 71.28 港元發行", "H 股報價（本輪未取得）"),
        ("全球 AI 與半導體股 2026-11-16 至 20 當週大跌（NVIDIA 約 −23%）", "否", "現價為下跌後價格；本模型用 11-20 收盤，不預測價格走勢", "股價基準日更新"),
    ]
    for i, k in enumerate(RISKS):
        for j, v in enumerate(k): RK.cell(4 + i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["D"].width = 60

    # ================= Sources =================
    SO["A1"] = "Sources：本模型引用的原始來源與資料商（含網址；HTML 展示層不放網址）。"
    for i, s_ in enumerate([S_PX, S_TJ, S_XY, S_LX, S_SZ, S_Q3, S_GJ, S_HK, S_RF, S_CRP, S_SAP,
                            f"線 4：{L4N}", f"線 6：{L6N}",
                            "physical_ai/harmonic_drive 模型（哈默納科前瞻 P/E）、physical_ai/nabtesco 模型（Nabtesco 前瞻 P/E）、physical_ai/leaderdrive 模型（A 股同業原值）",
                            "physical_ai/map/profiles/003021.md（背景，數字已另行查證）"], 3):
        SO.cell(i, 1, s_)
    SO.column_dimensions["A"].width = 160

    # ================= README（xlsx 內）=================
    S = T
    v = lambda k: f"Val!B{VR[k]}"
    lines = [
        ('="兆威機電（003021.SZ）個股模型 ' + VER + '｜build_id "&Sens!D2&"｜資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'),
        "命題：兆威機電目前仍是汽車與消費電子的微型傳動廠，股價（2026 年前三季淨利年減三成、滾動本益比約 130 倍）在為靈巧手選擇權定價；營收＝傳統微型傳動與精密零件（年增自設）＋靈巧手與機器人模組（線 4 出貨 × 每台靈巧手數 × 線 6 單價 × 份額），單價每年下降（線 6）。",
        (f'="結論（條件式）：在 FY2026 以 1–9 月實績為錨、傳統業務年增 "&{S("gcore_B","0%")}&"、線 4 基準人形出貨（× "&{S("hu_k","0.0")}&"）、靈巧手單價 "&{S("p_hand26","#,##0")}&" 美元且年降 "&{S("pd","0%")}&"、情境 B 的 2030 年靈巧手份額 "&{S("hs30_B","0.0%")}&"、同業 P/E 未加溢價（相對倍數係數 "&{S("rel_k","0.00")}&"）的假設下，加權目標價約 "&{S(v("tp"),"#,##0.0")}&" 元，相對現價 "&{S("px","#,##0.00")}&" 元 "&{S(v("up"),"+0.0%;-0.0%")}&"，估值評等 "&{v("rating")}&"，投資建議 "&{v("rec")}&"。"'),
        (f'="DCF 目標價 "&{S(v("dcf_tp"),"#,##0.0")}&"、Comparable 目標價 "&{S(v("cmp_tp"),"#,##0.0")}&"（各半）；WACC "&{S("wacc","0.00%")}&"、ke "&{S("ke","0.00%")}&"、g "&{S("g_tv","0.00%")}&"；情境 B 終值占 EV "&{S(f"Val!C{VR['tvs']}","0%")}&"。"'),
        (f'="上檔空間歸因：DCF 貢獻 "&{S(v("att_d"),"+0.0%;-0.0%")}&"、未調整 Comparable 貢獻 "&{S(v("att_c"),"+0.0%;-0.0%")}&"、相對倍數調整貢獻 "&{S(v("att_r"),"+0.0%;-0.0%")}&"。本公司現價對資料商前瞻 EPS 的倍數是同業中位數的 "&{S(v("relnow"),"0.0")}&" 倍；支持現價需要相對倍數係數 "&{S(v("relreq"),"0.0")}&"，升一級需要 "&{S(v("relflip"),"0.0")}&"。"'),
        (f'="收支關卡：情境 B FY2026–FY2030 FCFF 扣股利合計 "&{S(v("gate_b"),"#,##0.0")}&" 百萬元，"&{v("gate_t")}&"；距最近級距邊界 "&{S(v("dist"),"0.0%")}&"。"'),
        "主要失效條件（可觀察）：(1) 靈巧手營收由千萬級進入億級以上且客戶具名；(2) 傳統業務營收恢復雙位數成長、毛利率回到 33% 以上；(3) 全球人形出貨持續高於線 4 高情境；(4) 同業前瞻 P/E 重新定價（見 Sens 的同業前瞻 P/E 案例）。",
        "方法：DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，WACC 折現，明確期到 FY2032，RONIC＝WACC；現金以淨現金計入）50%＋同業前瞻 P/E × 模型 FY2027 EPS 50%；三情境 25／50／25。",
        "限制：機器人營收占比、份額、費用與稅率、淨現金皆未取得公司揭露（年報全文為影像檔未能讀取），以 Assumed 區間處理；淨現金為推算；同業 P/E 由原報價日按 2026-11-20 收盤等比例換算（EPS 未更新）；不讀 Siliconomics／Tokenomics／Wattonomics。詳見 repo README.md。",
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
            subprocess.run(["pkill", "-f", "lo_zhaowei"], check=False)
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
        ("情境 B 的 2030 年靈巧手份額改為情境 A 值", "hs30_B", ivE["hs30_A"]),
        ("情境 B 的 2030 年靈巧手份額改為情境 C 值", "hs30_B", ivE["hs30_C"]),
        ("情境 B 傳統業務年增改為情境 A 值", "gcore_B", ivE["gcore_A"]),
        ("情境 B 傳統業務年增改為情境 C 值", "gcore_B", ivE["gcore_C"]),
        ("情境 B 的 FY2030 毛利率改為情境 A 值", "gm30_B", ivE["gm30_A"]),
        ("情境 B 的 FY2030 毛利率改為情境 C 值", "gm30_B", ivE["gm30_C"]),
        ("人形出貨為線 4 的 5 倍（全部情境）", "hu_k", 5.0),
        ("靈巧手單價改用線 4 基準 4,750 美元", "p_hand26", 4750.0),
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
    s3 = solve("hs30_B", px, max(ivE["hs30_B"], 0.01), 1.0)
    s4 = solve("hu_k", px, 1.0, 200.0, over={"hs30_B": 0.5, "hs30_A": 0.5, "hs30_C": 0.5})
    nf = lambda x, txt: x if x else -1
    solve_rows = [
        ("支持現價（目標價＝現價）所需的人形出貨倍數（相對線 4，全部情境）", nf(s1, ""), 1.0, "其他輸入不變；−1＝200 倍內無解" if not s1 else f"其他輸入不變；情境 B 的 CY2030 全球人形出貨相當於 {LD['hum'][(2030, 'B')][0] * s1 / 1e6:,.1f} 百萬台"),
        (f"評等升一級（上檔 {flip:+.0%}）所需的人形出貨倍數", nf(s2, ""), 1.0, "其他輸入不變；−1＝200 倍內無解" if not s2 else f"情境 B 的 CY2030 全球人形出貨相當於 {LD['hum'][(2030, 'B')][0] * s2 / 1e6:,.2f} 百萬台"),
        ("支持現價所需的情境 B 2030 年靈巧手份額", nf(s3, ""), ivE["hs30_B"], "其他輸入不變；−1＝份額 100% 仍無解"),
        ("三情境靈巧手份額皆 50% 時，支持現價所需的人形出貨倍數", nf(s4, ""), 1.0, "份額與量同時放大的組合；−1＝200 倍內無解"),
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
    segs = [("微型傳動系統", "微型傳動系統 營收", Hs["E8"].value), ("精密零件", "精密零件 營收", Hs["E9"].value),
            ("其他業務（殘差）", "其他業務 營收", Hs["E10"].value)]
    json.dump(dict(company="zhaowei", ticker="003021.SZ", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=round(v, 2), unit="百萬元", source=S_LX + "；" + S_TJ) for n, l, v in segs],
                   sublines=[dict(parent="微型傳動系統", labels=["微型傳動：非機器人用途（子列）", "微型傳動：靈巧手與機器人模組（子列）"])],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）", "檢查：微型傳動子列合計 − 微型傳動分部（應為 0）"],
                   note="官方分部＝2025 年報分產品（資料商轉錄）：微型傳動系統、精密零件，其餘為其他業務（殘差）。靈巧手與機器人模組為微型傳動系統內的命題子列（Assumed），合計＝微型傳動分部。"),
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
    flag = open(os.path.join(HERE, "README_status.txt"), encoding="utf-8").read().strip() if os.path.exists(os.path.join(HERE, "README_status.txt")) else ""
    t = f"""# physical_ai/zhaowei/ — 兆威機電（003021.SZ）個股模型 {VER}

<!-- 本檔由 build_zhaowei.py 依成品重算值產生；數字不手寫（原則第 10 節第 51 條）。 -->

{flag}

**本模型適用《高不確定性預估工作原則》**（`docs/plan/高不確定性預估工作原則.md`）：靈巧手營收公司未揭露、人形需求尚未發生，年報全文為影像檔未能讀取，相關輸入一律給區間並標 Assumed／Analogy／未取得；結論只在證據邊界內成立。

- 成品：`dist/{NAME}`（build_id `{sens['build_id']}`；LibreOffice 重算，0 錯誤值）；HTML 展示層 `dist/{NAME[:-5]}.html`（`gen_html_json.py` 產生 `html.json`，再以 `tools/company_html/build_company_html.py` 產生）
- 建置：`python3 physical_ai/zhaowei/build_zhaowei.py`（需要 soffice；整本重算敏感度與反解；結尾自動跑 in_layout、sens_coverage --require-keycol、model_governance、rating_gate --strict、zero_rows、cached_values、segment_check）
- 日期：資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 2026-11-20（收盤 {I['px']:,.2f} 元，A 股）｜評價日 2026-12-31｜目標時點 2027-11-20
- 會計年度＝日曆年；單位百萬元人民幣；中國會計準則（法定財報口徑，GAAP）。深圳證券交易所（另有 H 股 02692，2026-03 掛牌；本模型以 A 股價格計全部股本）。

## 結論（條件式）

在「FY2026 以 1–9 月實績為錨、傳統業務年增 {I['gcore_B']:.0%}、線 4 基準人形出貨、靈巧手單價 {I['p_hand26']:,.0f} 美元（線 6）每年降 {I['pd']:.0%}、情境 B 的 2030 年靈巧手份額 {I['hs30_B']:.1%}、同業 P/E 不加溢價（相對倍數係數 {I['rel_k']:.2f}）」的假設下：

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
| 資料商共識 2026 年歸母淨利（只作對照；賣方目標價未取得） | {I['cons_ni26']:,.0f} 百萬元（國金 160） |

上檔空間歸因：DCF 貢獻 {g('att_d'):+.1%}、未調整 Comparable 貢獻 {g('att_c'):+.1%}、相對倍數調整 {g('att_r'):+.1%}。本公司現價對資料商前瞻 EPS 的倍數是核心同業中位數（{I['pe_peer']:.1f} 倍）的 {g('relnow'):.1f} 倍；要讓目標價＝現價，Comparable 的相對倍數係數須達 {g('relreq'):.1f}。DCF 以 RONIC＝WACC（新投資無超額報酬），而高相對倍數隱含持續超額報酬，兩法對競爭優勢的隱含假設相反。

**讀法：** 現價的溢價來自靈巧手選擇權，但按線模型推導，靈巧手營收在合理區間內仍不足以支撐。情境 B 靈巧手與機器人模組營收 {ser('hum', '{:,.0f}')}（FY2026→FY2032，百萬元），占合併營收由 {TH[f'E{THB["humsh"]}'].value:.1%} 升到 {TH[f'K{THB["humsh"]}'].value:.0%}。整本重算反解：其他輸入不變時，**支持現價需要人形出貨達線 4 的 {fn(sol[0][1], '{:,.1f}')} 倍**（{sol[0][3]}）；評等升一級需要 {fn(sol[1][1], '{:,.2f}')} 倍；只放大份額時，{('情境 B 2030 年份額達 100% 仍不足以支持現價' if sol[2][1] == -1 else f'需要情境 B 2030 年靈巧手份額 {sol[2][1]:.0%}')}。結論主要建立在 Assumed（機器人占比、份額、費用率、稅率、淨現金）與線模型的 Analogy 輸入上，不是公司揭露值；方向在合理區間兩端是否相同見敏感度表。

## 命題

1. **營收結構：** 官方產品分部為微型傳動系統（2025 年 {wb['Hist']['E8'].value:,.0f} 百萬元，占 {wb['Hist']['E8'].value / wb['Hist']['E4'].value:.0%}）、精密零件、其他業務；靈巧手與機器人模組屬微型傳動系統內，公司未揭露金額，本模型以 Assumed 設 2026 年占合併營收 {I['rob_sh26']:.1%}（區間 0.8–4%；媒體引述私募人士說法，無公司佐證）。
2. **量與份額：** 靈巧手市場＝線 4 全球人形出貨 × 線 6 每台靈巧手數 {I['n_hand']:.0f} × 線 6 單價；本公司份額 2026 年隱含 {TH[f'E{THB["hsh"]}'].value:.1%}（以線 4 基準校準），情境 B 到 2030 年 {I['hs30_B']:.1%}。
3. **價與傳統業務：** 靈巧手單價每年降 {I['pd']:.0%}（線 6；區間 8–20%）；傳統微型傳動與精密零件年增 A／B／C＝14%／8%／2%（自設，線模型無對應變數）。2026 年 1–9 月營收 +0.5%、淨利 −32%，毛利率 31.35%。

情境 B：營收 {ser('rev', '{:,.0f}')}；毛利率 {ser('gm', '{:.1%}')}；EPS {ser('eps', '{:.2f}')} 元（FY2026→FY2032）。

## 主要假設與可證偽指標（原則第 10 節第 20 條）

| 承重假設 | 基準（A／B／C 或區間） | 標記 | 可觀察指標 |
|---|---|---|---|
| 同業前瞻 P/E | 中位數 {I['pe_peer']:.1f} 倍（6 家，含綠的諧波） | Derived | 同一估值日重抓 |
| 傳統業務年增 | {I['gcore_A']:.0%}／{I['gcore_B']:.0%}／{I['gcore_C']:.0%} | Assumed | 季營收（2026 Q3 −4.9%） |
| 2030 年靈巧手份額 | {I['hs30_A']:.1%}／{I['hs30_B']:.1%}／{I['hs30_C']:.1%} | Assumed | 2026 年報靈巧手收入、客戶具名 |
| 2026 年機器人模組占營收 | {I['rob_sh26']:.1%}（0.8–4%） | Assumed | 年報或 H 股業績公告的機器人分部 |
| 全球人形出貨（線 4） | 線 4 高／基準／低；× hu_k {I['hu_k']:.0f}（高案 2） | Assumed（線 4） | 機構季度出貨 |
| 靈巧手單價與年降（線 6） | {I['p_hand26']:,.0f} 美元（2,000–4,750）；{I['pd']:.0%}（8–20%） | Analogy | 量產報價；線 4 基準為 4,750，兩線不一致 |
| FY2030 毛利率 | {I['gm30_A']:.1%}／{I['gm30_B']:.0%}／{I['gm30_C']:.0%} | Assumed | 季毛利率（2026 Q3 32.4%） |

## 敏感度（整本重算，前 10 大）

| 案例 | 鍵 | 加權目標價 | 相對基準 | DCF 目標價 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | `{c['key']}` | {c['tp']:,.1f} | {c['tp'] / sens['base_tp'] - 1:+.1%} | {c['dcf']:,.1f} | {c['rating']} |" for c in top) + f"""

完整案例見 xlsx Sens 頁（每個 In 頁有低／高區間的鍵都有案例；另有情境 B 改 A／C 值、人形出貨 5 倍、線 4 單價、情境權重案例）。**翻轉評等所需的變動**：升一級需加權目標價達 {fn(up_b)} 元（現價 {I['px']:,.2f}）；見 Recon 反解與相對倍數係數。

## 反解（Recon 頁）

""" + "\n".join(f"- {a}：{fn(b, '{:,.2f}')}（基準 {fn(c0, '{:,.2f}')}）；{d}" for a, b, c0, d in sol) + f"""

## 限制與「分不出來／沒找到」

- **分不出來（資料性質）：** 靈巧手營收與份額公司未揭露；人形量級屬未發生事件，線 4 三情境差距約 4–5 倍。
- **沒找到（工作未完）：** 公司 2025 年報全文為影像檔未能讀取（費用率、折舊攤銷、經營現金流取自券商摘要）；2023、2024 年分部營收與銷售管理費用；稅金及附加、其他收益、減值、所得稅、少數股東損益、資本支出、有息負債；2026-09-30 資產負債表（淨現金為推算，區間 1,400–1,800）；前五大客戶；賣方目標價；H 股現價。
- **線模型不一致：** 靈巧手單價線 6＝3,500 美元、線 4＝4,750 美元（採線 6）。見 xlsx「對帳」頁。
- **同業倍數：** 同業原報價日 2026-09-30～10-30，前瞻 P/E 依 2026-11-20 收盤等比例換算（EPS 未更新）；哈默納科、Nabtesco 取本專案模型值；中大力德、五洲新春資料商無前瞻 P/E，報價日未取得，未進中位數；少於 8 家不做迴歸。
- **市場背景：** 2026-11-16 至 11-20 全球 AI 與半導體股大跌（NVIDIA 當週約 −23%、日經 225 約 −7.7%），收盤檔下跌屬實；本公司兩週跌 26%。
- **rf 逾期：** 中國 10 年期國債 1.61% 為 2026-11-05 值（跨模型一致）。
- **協調檔：** 實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics；相關條文不適用。
- **獨立核對：** 依專案規則，交付判讀的數字應由未參與製作的代理獨立核對；本輪未安排。

## 資料來源

見 xlsx Sources 頁（資料商財務頁、券商點評、媒體轉錄三季報、H 股公告、線 4／6 成品）。

## 版本紀錄

- {VER}（2026-11-24）：首版。現價 {I['px']:,.2f}（2026-11-20 收盤）；加權目標價 {g('tp'):,.1f}；上檔空間 {g('up'):+.1%}；{g('rating')}／{g('rec')}；收支關卡 {g('gate_b'):,.1f}（{g('gate_t')}）。
"""
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write(t)


if __name__ == "__main__":
    main()

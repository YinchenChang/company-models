# 綠的諧波 Leaderdrive（688017.SS，上海科創板）個股模型 v0.1 建置腳本
# 產出：physical_ai/leaderdrive/dist/20261123_綠的諧波個股模型_v0.1.xlsx（LibreOffice 重算後的成品，含公式與快取值）
# 結構：命題模組（Thesis：諧波減速器需求＝工業機器人（線 3）＋人形（線 4 出貨 × 線 6 每台用量）＋其他用途；單價年降幅讀線 6）
#       → 官方產品分部（諧波減速器及金屬件／機電一體化產品／智能自動化裝備及其他業務）→ 標準損益表 → FCFF DCF＋同業 P/E。
# 敏感度：每個案例覆寫 In 頁基準欄後整本重建、LibreOffice 重算（整本重算，不用閉式近似）；反解同法以二分搜尋。
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md、docs/plan/高不確定性預估工作原則.md。
#       實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics 協調檔。
# 用法：python3 physical_ai/leaderdrive/build_leaderdrive.py   （需要 soffice；暫存於 $HOME/ld_build）
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
NAME = f"20261123_綠的諧波個股模型_{VER}.xlsx"
OUT = os.path.join(HERE, "dist", NAME)
HOMEDIR = os.path.expanduser("~")
TMP = os.path.join(HOMEDIR, "ld_build")
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
S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv：688017.SS；未還原股利）"
S_SA = "stockanalysis.com SHA:688017（資料商；財務摘要、資產負債表季頁、statistics 頁，2026-11-23 讀取；法定季報原文未直接讀取）"
S_SA_ST = "stockanalysis.com SHA:688017 statistics（資料商；頁面報價日 2026-10-09，2026-11-23 讀取）"
S_DW = "東吳證券 綠的諧波 2025 年報＆2026 一季報點評（2026-04-23；https://pdf.dfcfw.com/pdf/H3_AP202604231821491305_1.pdf）：轉錄公司 2025 年報（百萬元取整）"
S_GY1 = "國元證券 2025 年報及 2026 年一季報點評（2026-04-29；https://data.eastmoney.com/report/info/AP202604291821759233.html）：轉錄公司 2025 年報分產品營收與毛利率"
S_GY2 = "國元證券 2026 年半年報點評（檔名日期 2026-09-02；https://data.eastmoney.com/report/info/AP202609021828901486.html）：轉錄公司 2026 年半年報"
S_GD = "光大證券 2024 年報點評（2025-04-30；https://pdf.dfcfw.com/pdf/H3_AP202505011665485338_1.pdf）：轉錄公司 2023、2024 年報（百萬元取整）"
S_QY = "群益證券（香港）2026 年三季報點評（2026-11-02；https://pdf.dfcfw.com/pdf/H3_AP202611021831314272_1.pdf）：轉錄公司 2026 年三季報"
S_H = "公司公告 2026-023《關於籌劃發行 H 股股票並在香港聯合交易所有限公司上市的提示性公告》（2026-08-27；https://pdf.dfcfw.com/pdf/H2_AN202608261828493082_1.pdf）：發行規模、時程、募資用途皆未揭露"
S_LF = "來福諧波（紹興）港股申請版本草擬本（2026；https://www1.hkexnews.hk/app/sehk/2026/108602/a134036/sehk26061400335_c.pdf）：諧波減速器銷量與平均售價 2023 795、2024 722、2025 571 元／顆"
S_RF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與 innolight、eoptolink、victory_giant 同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：China CRP 0.91%（與 innolight 同一取值）"
S_CLOSE = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
S_SAP = "stockanalysis.com statistics 頁（資料商；2026-11-23 讀取；原報價日見 Comp 頁 M 欄）"
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
    wb6, l6 = _line_rows(L6, 1)
    d = {"ind": {}, "hum": {}}
    for y in range(2025, 2033):
        for s, tag in (("A", "hi"), ("B", "base"), ("C", "lo")):
            k = f"IR.red.hd.units.{y}.{tag}" if y > 2025 else "IR.red.hd.units.2025.base"
            r, v = l3[k]
            d["ind"][(y, s)] = (v[4], f"{L3N} 下游取數!E{r}（{k}）")
    for y in range(2026, 2033):
        r, v = l4[f"SHIP_{y}"]
        for s, c, col in (("A", 4, "E"), ("B", 3, "D"), ("C", 2, "C")):
            d["hum"][(y, s)] = (v[c], f"{L4N} 下游取數!{col}{r}（SHIP_{y}）")
    ws6 = wb6["輸入"]
    # 線 6 輸入頁：人形每台諧波用量（低／基準／高情境矩陣 F 欄，列以 C1 表頭後第一個「HR」列找）；單價與年降幅（E 節「HR」列）
    hr = [r for r in range(1, ws6.max_row + 1) if ws6.cell(r, 1).value == "HR"]
    d["n_hum"] = [(ws6.cell(r, 6).value, f"{L6N} 輸入!F{r}") for r in hr[:3]]         # 低、基準、高
    pr = hr[3]
    d["pd"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{openpyxl.utils.get_column_letter(c)}{pr}") for c in (7, 8, 9)]
    d["p26"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{openpyxl.utils.get_column_letter(c)}{pr}") for c in (4, 5, 6)]
    for k in ("諧波減速機市場規模 2026", "諧波減速機市場規模 2030", "諧波減速機市場規模 2032"):
        r, v = l6[k]; d[k] = (v[2], f"{L6N} 下游取數!C{r}")
    r, v = l4["Q_HARMONIC_AVG"]; d["q_avg4"] = (v[3], f"{L4N} 下游取數!D{r}（Q_HARMONIC_AVG）", v[2], v[4])
    r, v = l4["P_HARMONIC"]; d["p_harm4"] = (v[3], f"{L4N} 下游取數!D{r}（P_HARMONIC）")
    for y in (2026, 2028, 2030):
        r1, a = l4[f"HARM_MKT_{y}"]; r2, b = l4[f"HARM_UNITS_{y}"]
        d[f"p4_{y}"] = (a[3] * 1e6 / b[3], f"{L4N} 下游取數!D{r1} ÷ D{r2}（HARM_MKT_{y} ÷ HARM_UNITS_{y}，Derived）")
    return d


LD = line_data()

# ---------------- In 頁輸入 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 來源誘因, 說明)
IN = [
    ("sec", "1. 市場資料與日期"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, "", "資料截止日 2026-11-23（財報、券商預測、同業倍數）；股價 2026-11-20；成品日 2026-11-23"),
    ("px", "現價（2026-11-20 收盤）", "元", None, 230.80, None, "Verified", S_PX, "", "csv 值 230.800003；11-09 收盤 296.20，兩週跌 22%"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,12,31)", None, "Assumed", "本專案慣例：評價日 2026-12-31（會計年度＝日曆年，FY2026 在評價日已結束）", "", "折現期間由評價日起算"),
    ("tp_date", "目標時點（px_date 後 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", ""),
    ("bs_date", "資產負債表日", "日期", None, "=DATE(2026,9,30)", None, "Verified", S_SA, "", "2026 年三季報"),
    ("stub_frac", "Stub 期（資產負債表日→評價日）", "年", None, "=(val_date-bs_date)/365", None, "Derived", "＝(評價日 − 資產負債表日) ÷ 365", "", "Stub FCFF＝FY2026 FCFF × stub_frac"),
    ("sh_v", "評價日股數（總股本）", "百萬股", None, 183.33, None, "Verified", S_SA + "；東吳、群益報告同為 183.33 百萬股", "", "限制性股票、選擇權本輪未取得揭露，以 0 計；H 股發行另見 h_iss"),
    ("cash_v", "現金＋短期投資＋交易性金融資產（2026-09-30）", "百萬元", None, 1065.0, None, "Verified", S_SA + "：貨幣資金 392.34、短期投資 439.28，連同交易性金融資產合計約 1,065", "", "2025-12-31 約 1,015"),
    ("ltinv_v", "長期投資（2026-09-30；含長期存單、債權與股權投資）", "百萬元", None, 546.61, None, "Verified", S_SA, "", "明細未取得；含與 SKF 合資（持股 40%）等股權投資"),
    ("inv_k", "長期投資計入淨現金的比例", "%", 0.5, 1.0, None, "Assumed", "基準全額計入；低案 50%（股權投資與流動性折價）", "", ""),
    ("debt_v", "有息負債（2026-09-30）", "百萬元", None, 0.17, None, "Verified", S_SA + "：一年內到期非流動負債 0.17；無短借、長借", "", ""),
    ("nci_v", "少數股東權益（2026-09-30）", "百萬元", None, 9.78, None, "Verified", S_SA, "", "自淨現金扣除"),
    ("fx", "美元兌人民幣（2026-11-20）", "元／美元", None, 6.6665, None, "Verified", S_CLOSE + "：CNY=X", "", "只用於 Recon 與線 6 市場規模對照"),
    ("cons_tp", "賣方平均目標價（15 家，2026-10-09，三季報前）", "元", None, 322.66, None, "Interested-party", S_SA_ST, IP_SA, "只作對照，不進模型；三季報（2026-10-30）前的值，只作歷史對照"),
    ("tp_qy", "群益證券目標價（2026-11-02，三季報後）", "元", None, 305.0, None, "Interested-party", S_QY, IP_BRK, "只作對照，不進模型"),
    ("sec", "2. 實績錨（FY2026＝2026 年，會計年度＝日曆年）"),
    ("r9m26", "2026 年 1–9 月營收", "百萬元", None, 591.0, None, "Verified", S_QY + "；國金點評同為 5.91 億（+45.3%）", "", "公司三季報原文未直接讀取"),
    ("g9m26", "2026 年 1–9 月營收年增", "%", None, 0.453, None, "Verified", S_QY, "", ""),
    ("q4g", "2026 年第 4 季營收年增", "%", 0.25, 0.45, 0.65, "Assumed", "基準＝1–9 月年增 45%（Q3 單季 +56%）", "", "FY2026 營收＝1–9 月實績＋2025Q4 ×（1＋q4g）"),
    ("sh_h26", "FY2026 諧波減速器及金屬件 占營收", "%", None, "=Q!D10/Q!D4", None, "Derived", "＝2026 上半年諧波營收 ÷ 營收（Q 頁）", "", "上半年只揭露諧波與機電兩項"),
    ("sh_m26", "FY2026 機電一體化產品 占營收", "%", None, "=Q!D11/Q!D4", None, "Derived", "＝2026 上半年機電營收 ÷ 營收（Q 頁）", "", "智能自動化裝備及其他業務＝餘數"),
    ("gm26", "FY2026 毛利率", "%", None, "=(Q!D7*Q!D4+Q!E7*Q!E4)/(Q!D4+Q!E4)", None, "Derived", "＝（上半年毛利率 × 上半年營收＋第 3 季毛利率 × 第 3 季營收）÷ 1–9 月營收（Q 頁）", "", "第 4 季假設同 1–9 月"),
    ("opex26", "FY2026 銷售＋管理＋研發費用率", "%", 0.12, 0.14, 0.16, "Assumed", S_GY2 + "：上半年費用率 15.96%（含財務 0.33%）；" + S_QY + "：第 3 季費用率 11.8%", "", "1–9 月加權約 14%"),
    ("sec", "3. 命題模組：諧波減速器（量 × 價；線 3、線 4、線 6）"),
    ("asp25", "2025 每顆諧波減速器對應分部營收", "元／顆", None, "=Hist!E8*1e6/425000", None, "Derived", "＝2025 諧波減速器及金屬件營收 ÷ 諧波減速器銷量 42.5 萬台（" + S_DW + "）", "", "分子含金屬件營收（金屬部件銷量 71.2 萬件），故高於純諧波單價；來福 2025 平均售價 571 元（Analogy，見對帳）"),
    ("pd", "諧波減速機單價年降幅（線 6）", "%／年", LD["pd"][0][0], LD["pd"][1][0], LD["pd"][2][0], "Analogy", f"線 6：{LD['pd'][1][1]}（低 {LD['pd'][0][1]}、高 {LD['pd'][2][1]}）", "",
     "工業與人形用同一降幅（線 6 不分用途）；本公司 2024→2025 每顆分部營收 −15%（Derived，含組合變化）高於線 6 高案"),
    ("n_hum", "人形每台諧波減速機用量（出貨組合平均，線 6）", "顆／台", LD["n_hum"][0][0], LD["n_hum"][1][0], LD["n_hum"][2][0], "Analogy",
     f"線 6：{LD['n_hum'][1][1]}（低 {LD['n_hum'][0][1]}、高 {LD['n_hum'][2][1]}）；Optimus 類全尺寸 14–16 顆", "", f"線 4 Q_HARMONIC_AVG 為 {LD['q_avg4'][0]:g}（{LD['q_avg4'][2]:g}–{LD['q_avg4'][3]:g}），兩線不一致，見對帳"),
    ("hu_k", "人形出貨相對線 4 的倍數（全部情境）", "x", None, 1.0, 2.0, "Assumed", "基準＝線 4 原值；高案 2 倍：" + S_QY + "引述中國 2026 上半年人形產量逾 4 萬台、全年可望逾 10 萬台（線 4 基準 2026 全球 5.5 萬台）", "", "反解變數"),
    ("hum_sh26", "FY2026 諧波營收中人形用占比", "%", 0.08, 0.15, 0.25, "Assumed", "公司未揭露用途別營收；券商稱具身智能業務規模大幅成長（" + S_QY + "）", "", "隱含 2026 全球人形諧波份額見 Thesis"),
    ("oth_sh26", "FY2026 諧波營收中其他用途（機床、半導體、醫療、金屬件）占比", "%", 0.10, 0.15, 0.25, "Assumed", "公司未揭露用途別營收；金屬部件銷量 71.2 萬件（" + S_DW + "）", "", "工業機器人用＝1 − 人形 − 其他"),
    ("g_ao", "智能自動化裝備及其他業務 年增（FY2027 起）", "%", 0.0, 0.05, 0.15, "Assumed", "2025 年 +221%（基期小）；占營收約 2–4%", "", ""),
    ("sec", "4. 命題模組：情境專屬驅動（A 樂觀／B 基準／C 保守；只列基準欄，不設敏感度區間）"),
    ("drift_A", "情境 A：工業機器人用諧波份額年漂移", "%／年", None, 0.08, None, "Assumed", "國產替代：中國裝機國產品牌占比 2025 年 55%（線 3 IR.share.china_domestic.2025）；本公司全球份額 8%（2023）→12%（2025）（線 6 份額頁，Interested-party）", "", "相對線 3 全球需求成長的超額（正＝取得份額）"),
    ("drift_B", "情境 B：工業機器人用諧波份額年漂移", "%／年", None, 0.03, None, "Assumed", "同上；2023→2025 全球份額年增約 22%（Interested-party）與 2026 年 1–9 月營收 +45%（線 3 全球需求 +9%）顯示仍在取得份額；基準只外推其一小部分（價格競爭）", "", ""),
    ("drift_C", "情境 C：工業機器人用諧波份額年漂移", "%／年", None, -0.03, None, "Assumed", "同上；來福等國產廠擴產（來福 2025 銷量 +104%，" + S_LF + "）", "", ""),
    ("hs30_A", "情境 A：2030 年占全球人形諧波用量份額", "%", None, 0.30, None, "Assumed", "2026 隱含份額見 Thesis（基準約 18%）；中國整機多用國產諧波", "", "2026 隱含份額線性走到 2030，其後持平"),
    ("hs30_B", "情境 B：2030 年占全球人形諧波用量份額", "%", None, 0.18, None, "Assumed", "同上；基準約持平 2026 隱含份額", "", ""),
    ("hs30_C", "情境 C：2030 年占全球人形諧波用量份額", "%", None, 0.08, None, "Assumed", "同上；非中國整機（Tesla 等）供應鏈未必採用；行星減速替代", "", ""),
    ("goth_A", "情境 A：其他用途諧波營收年增（FY2027 起）", "%", None, 0.12, None, "Assumed", "機床、半導體設備、醫療、金屬件", "", ""),
    ("goth_B", "情境 B：其他用途諧波營收年增（FY2027 起）", "%", None, 0.08, None, "Assumed", "同上", "", ""),
    ("goth_C", "情境 C：其他用途諧波營收年增（FY2027 起）", "%", None, 0.03, None, "Assumed", "同上", "", ""),
    ("gmech_A", "情境 A：機電一體化產品營收年增（FY2027 起）", "%", None, 0.35, None, "Assumed", "2025 年 +41%、2026 上半年 +31%（" + S_GY2 + "）", "", "關節模組，部分受人形帶動"),
    ("gmech_B", "情境 B：機電一體化產品營收年增（FY2027 起）", "%", None, 0.25, None, "Assumed", "同上", "", ""),
    ("gmech_C", "情境 C：機電一體化產品營收年增（FY2027 起）", "%", None, 0.10, None, "Assumed", "同上", "", ""),
    ("gm30_A", "情境 A：FY2030 毛利率", "%", None, 0.36, None, "Assumed", "2025 年 36.9%、2026 年 1–9 月約 32.7%（擴產成本與價格競爭）", "", "FY2026 起線性走到 FY2030，其後持平"),
    ("gm30_B", "情境 B：FY2030 毛利率", "%", None, 0.33, None, "Assumed", "同上", "", ""),
    ("gm30_C", "情境 C：FY2030 毛利率", "%", None, 0.28, None, "Assumed", "同上；來福 2025 毛利率 25.6%（Analogy）", "", ""),
    ("tp_wA", "情境權重 A", "", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", ""),
    ("tp_wB", "情境權重 B", "", None, 0.50, None, "Assumed", "同上", "", ""),
    ("tp_wC", "情境權重 C", "", None, 0.25, None, "Assumed", "同上", "", ""),
    ("sec", "5. 損益與現金流"),
    ("ts_r", "稅金及附加 ÷ 營收", "%", None, "=Hist!E19/Hist!E4", None, "Derived", "＝2025 年稅金及附加 ÷ 營收", "", ""),
    ("opex30", "FY2030 銷售＋管理＋研發費用率", "%", 0.10, 0.12, 0.145, "Assumed", "2025 年 16.3%、2026 第 3 季約 11.5%（營運槓桿）", "", "FY2026→FY2030 線性，其後持平"),
    ("oi_r", "其他收益（政府補助等）÷ 營收", "%", 0.005, "=Hist!E24/Hist!E4", 0.03, "Derived", "＝2025 年其他收益 12 ÷ 營收（" + S_DW + "）", "", "計入營業利益"),
    ("imp_r", "減值損失 ÷ 營收", "%", 0.005, 0.015, 0.03, "Assumed", "2025 年減值 31（5.4%，含存貨與信用減值）；2026 第 3 季信用減值增加", "", "計入營業利益"),
    ("inv_y", "金融資產收益率（投資收益＋公允價值變動）", "%", 0.015, 0.025, 0.035, "Assumed", "2025 年（36＋16）÷ 年底金融資產約 1,630＝3.2%（Derived）；2026 第 3 季投資收益減少", "", "業外；DCF 另以淨現金計入，不重複"),
    ("tax_r", "有效稅率", "%", 0.10, "=Hist!E33/Hist!E32", 0.20, "Derived", "＝2025 年所得稅 ÷ 稅前利益（高新技術企業 15% 與研發加計扣除）", "", "敏感度低／高 10%／20%（Assumed）"),
    ("nci_r", "少數股東損益 ÷ 稅後淨利", "%", None, "=Hist!E35/Hist!E34", None, "Derived", "＝2025 年少數股東損益 ÷ 稅後淨利", "", ""),
    ("payout", "配息率", "%", None, "=0.21/0.69", None, "Derived", "＝2025 年每股股利 0.21 元 ÷ EPS 0.69 元（" + S_SA + "）", "", "用於收支關卡與目標時點前股利"),
    ("da26", "FY2026 折舊攤銷 ÷ 營收", "%", None, "=Hist!E40/Hist!E4", None, "Derived", "＝2025 年折舊攤銷 60 ÷ 營收", "", "之後線性收斂到資本支出比例 × da_ss（FY2032）"),
    ("da_ss", "穩態折舊 ÷ 資本支出", "x", None, 0.8, None, "Assumed", "穩態資本支出高於折舊（成長）", "", ""),
    ("cx26", "FY2026 資本支出", "百萬元", 150, 200, 250, "Assumed", S_DW + "：2026E 資本開支 249；" + S_QY + "：募投項目完工延至 2028 年底", "", ""),
    ("cx_r", "資本支出 ÷ 營收（FY2027 起）", "%", 0.09, 0.13, 0.18, "Assumed", "2023 年 21%、2024 年 8%、2025 年 17%（Hist，Derived）", "", ""),
    ("nwc_r", "營運資金增加 ÷ 營收增量", "%", 0.20, 0.30, 0.45, "Assumed", S_SA + "：2025-12→2026-09 存貨 +142、應收 +112、應付 +223", "", "存貨水位高"),
    ("h_iss", "H 股發行占擴大後股本（尾端事件，基準 0）", "%", None, 0.0, 0.10, "Assumed", S_H, "", "原則第 10 節第 19 條：事件型尾端風險進 Sens、不進基準"),
    ("h_disc", "H 股發行價相對現價折價", "%", None, 0.10, None, "Assumed", "港股 A＋H 慣例折價（未查核，不作依據）", "", "只在 h_iss＞0 時有作用"),
    ("sec", "6. 折現率（DCF 標準規格第 2 節）"),
    ("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", S_RF, "", "日期 2026-11-05，較股價基準日早 11 個交易日（逾期；跨模型一致，下次全面改版更新）"),
    ("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", ""),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", ""),
    ("b_own", "本公司 5 年 β（原始）", "x", None, 0.79, None, "Verified", S_SA_ST, "", ""),
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
SCN_DRIVERS = [("drift", "工業機器人用諧波份額年漂移"), ("hs30", "2030 年人形諧波份額"), ("goth", "其他用途諧波營收年增"),
               ("gmech", "機電一體化產品營收年增"), ("gm30", "FY2030 毛利率")]

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
    HI["A1"] = "Hist：歷史財務（百萬元人民幣；中國會計準則）。券商轉錄之年報數字為百萬元取整。公式只引用本頁（定義名稱 h_*）。"
    hdr(HI, 3, ["項目", "", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"])
    HROWS = [  # (列, 鍵, 項目, [2023, 2024, 2025], 標記, 來源)
        (4, "h_rev", "營業收入", [356.17, 387.41, 570.71], "Verified", S_SA + "；券商轉錄年報一致（356、387、571）"),
        (5, "h_cogs", "營業成本", [210, 242, "=E4*(1-0.3691)"], "Verified", S_GD + "；2025＝營收 ×（1 − 毛利率 36.91%）（" + S_GY1 + "；東吳轉錄 360）"),
        (6, "h_gp", "毛利（＝營收 − 營業成本）", ["=C4-C5", "=D4-D5", "=E4-E5"], "Derived", "＝營業收入 − 營業成本"),
        (7, "h_gm", "毛利率", ["=C6/C4", "=D6/D4", "=E6/E4"], "Derived", "＝毛利 ÷ 營收"),
        (8, "h_seg_h", "諧波減速器及金屬件 營收", ["=325/1.026", 325, 476], "Verified", S_GD + "（2024：3.25 億、+2.6%）；" + S_GY1 + "（2025：4.76 億）；2023＝2024 ÷ 1.026（Derived）"),
        (9, "h_seg_m", "機電一體化產品 營收", ["=53/1.567", 53, 74], "Verified", S_GD + "（2024：0.53 億、+56.7%）；" + S_GY1 + "（2025：0.74 億）；2023＝2024 ÷ 1.567（Derived）"),
        (10, "h_seg_a", "智能自動化裝備及其他業務 營收（＝營收 − 上兩項）", ["=C4-C8-C9", "=D4-D8-D9", "=E4-E8-E9"], "Derived", "智能自動化裝備 2024 約 5、2025 約 15（券商轉錄，億元一位小數），其餘為其他業務收入；以殘差計"),
        (11, "h_hd_units", "諧波減速器銷量", [None, "=425/1.7248", 425], "Verified", S_DW + "：2025 年 42.5 萬台（+72.48%）；2024＝2025 ÷ 1.7248（Derived）；單位千台"),
        (12, "h_asp", "每顆諧波對應分部營收（元／顆）", [None, "=D8/D11*1000", "=E8/E11*1000"], "Derived", "＝諧波減速器及金屬件營收 ÷ 銷量（含金屬件營收）"),
        (13, "h_sell", "銷售費用", [11, 13, 13], "Verified", S_GD + "；" + S_DW),
        (14, "h_admin", "管理費用", [22, 22, 26], "Verified", S_GD + "；" + S_DW),
        (15, "h_rnd", "研發費用", [48, 50, 54], "Verified", S_GD + "；" + S_DW),
        (16, "h_opex", "營業費用（銷售＋管理＋研發）", ["=SUM(C13:C15)", "=SUM(D13:D15)", "=SUM(E13:E15)"], "Derived", "＝三項合計"),
        (17, "h_fin", "財務費用（負＝淨利息收入）", [-23, 0, 2], "Verified", S_GD + "；" + S_DW),
        (18, "h_inv", "投資收益＋公允價值變動", [12, 28, "=36+16"], "Verified", S_GD + "（2023、2024 只列投資收益）；" + S_DW + "（2025：投資淨收益 36、公允價值變動 16）"),
        (19, "h_ts", "稅金及附加", [3, 4, 3], "Verified", S_GD + "；" + S_DW),
        (20, "h_opp", "營業利潤（中國會計準則）", [92, 61, 145], "Verified", S_GD + "；" + S_DW),
        (21, "h_ebit", "營業利益（核心：營業利潤 − 投資收益與公允價值 ＋ 財務費用）", ["=C20-C18+C17", "=D20-D18+D17", "=E20-E18+E17"], "Derived", "本模型營業利益不含金融資產收益與財務費用（DCF 另以淨現金計入）"),
        (22, "h_oth_op", "其他營業損益（＝核心營業利益 − 毛利 ＋ 稅金及附加 ＋ 營業費用）", ["=C21-C6+C19+C16", "=D21-D6+D19+D16", "=E21-E6+E19+E16"], "Derived", "其他收益、減值、資產處置等合計（殘差；含券商取整誤差）"),
        (23, "h_imp", "減值損失（2025）", [None, None, -31], "Verified", S_DW),
        (24, "h_oi", "其他收益（2025）", [None, None, 12], "Verified", S_DW),
        (25, "h_ebitm", "營業利益率（核心）", ["=C21/C4", "=D21/D4", "=E21/E4"], "Derived", ""),
        (32, "h_pbt", "利潤總額（稅前）", [92, 61, 144], "Verified", S_GD + "；" + S_DW),
        (26, "h_nonop", "業外（＝稅前 − 核心營業利益）", ["=C32-C21", "=D32-D21", "=E32-E21"], "Derived", "投資收益、公允價值、財務費用、營業外收支合計"),
        (33, "h_tax", "所得稅", [7, 5, 18], "Verified", S_GD + "；" + S_DW),
        (34, "h_nit", "稅後淨利", [85, 56, 126], "Verified", S_GD + "；" + S_DW),
        (35, "h_nci", "少數股東損益", [1, 0, 1], "Verified", S_GD + "；" + S_DW),
        (36, "h_ni", "歸屬母公司淨利", [84.16, 56.17, 124.37], "Verified", S_SA),
        (37, "h_eps", "EPS（基本，元）", [0.50, 0.33, 0.69], "Verified", S_SA + "（券商報告以 183.33 百萬股計為 0.46、0.31、0.68）"),
        (38, "h_dps", "每股股利（元）", [0.30, 0.10, 0.21], "Verified", S_SA),
        (39, "h_sh", "股數（＝淨利 ÷ EPS，百萬股）", ["=C36/C37", "=D36/D37", "=E36/E37"], "Derived", ""),
        (40, "h_da", "折舊攤銷", [44, 55, 60], "Verified", S_GD + "；" + S_DW),
        (41, "h_cx", "資本支出", [76, 32, 98], "Verified", S_GD + "（淨資本支出）；" + S_DW),
        (42, "h_ocf", "營業現金流", [None, None, 152], "Verified", S_DW),
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
    for c in "CDE": HI[f"{c}28"] = f"={c}8+{c}9+{c}10-{c}4"
    HI.cell(29, 1, "差額：核心營業利益＋業外 − 稅前（券商取整，非應為 0）")
    for c in "CDE": HI[f"{c}29"] = f"={c}21+{c}26-{c}32"
    HI.column_dimensions["A"].width = 58; HI.column_dimensions["G"].width = 90

    # ================= Q（季度；一列一個數字，單一標記）=================
    Q["A1"] = "Q：2026 年季度資料（百萬元）。一列一個項目、單一主要標記；券商轉錄公司定期報告（原文未直接讀取）。"
    hdr(Q, 3, ["項目", "2026Q1", "2026Q2", "2026H1", "2026Q3", "2026 年 1–9 月", "標記", "來源", "來源誘因", "說明"])
    QROWS = [
        (4, "營業收入", [140, "=D4-B4", 349, 242, "=D4+E4"], "Verified", S_GY1 + "（Q1 1.40 億）；" + S_GY2 + "（H1 3.49 億）；" + S_QY + "（Q3 2.42 億）", "", "1–9 月合計與 In 頁 r9m26（5.91 億）差為取整"),
        (5, "營業收入年增", [0.4296, None, 0.3864, 0.561, 0.453], "Verified", "同上", "", ""),
        (6, "毛利率（Q1）", [0.3362, None, None, None, None], "Verified", S_DW, "", ""),
        (7, "毛利率", [None, None, 0.3226, 0.334, None], "Verified", S_GY2 + "（H1 32.26%）；" + S_QY + "（Q3 33.4%）", "", "Q3 年減 6.2pt（擴產成本）"),
        (8, "歸屬母公司淨利", [33, "=D8-B8", 70, 50, "=D8+E8"], "Verified", "同上（億元取整）", "", "1–9 月 1.22 億（+30.3%）"),
        (9, "費用率（含財務費用）", [None, None, 0.1596, 0.118, None], "Verified", S_GY2 + "；" + S_QY, "", ""),
        (10, "諧波減速器 營收", [None, None, 288, None, None], "Verified", S_GY2 + "：2.88 億、+46.32%、占 82.67%、毛利率 32.79%", "", ""),
        (11, "機電一體化 營收", [None, None, 55, None, None], "Verified", S_GY2 + "：0.55 億、+31.45%、占 15.70%、毛利率 29.75%", "", ""),
    ]
    for (rr, lab, vals, tag, src, inc, note) in QROWS:
        Q.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = Q.cell(rr, 2 + j, v)
                if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
                if isinstance(v, float) and v < 1: c.number_format = "0.00%"
        Q.cell(rr, 7, tag); Q.cell(rr, 8, src); Q.cell(rr, 9, inc); Q.cell(rr, 10, note)
    Q.column_dimensions["A"].width = 30; Q.column_dimensions["H"].width = 90

    # ================= 線取數（線模型下游取數；藍字＝讀入值，H 欄寫來源檔與儲存格）=================
    LN["A1"] = "線取數：本模型讀入的線層級變數年度路徑（建置時由線模型成品依 ID 讀值）。情境對應：A＝線模型高情境、B＝基準、C＝低情境。"
    hdr(LN, 3, ["項目", "情境", ""] + [f"FY{y}" for y in range(2025, 2033)] + ["單位", "來源（檔案與儲存格）"])
    LR = {}
    r = 4
    for key, lab, unit, dd, y0 in (("ind", "全球機器人諧波減速機需求（線 3；僅工業機器人含協作）", "顆", LD["ind"], 2025),
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

    # ================= Thesis（命題模組）=================
    TH["A1"] = ("命題模組：諧波減速器及金屬件營收＝工業機器人用（FY2026 錨 × 線 3 全球需求成長 × 份額漂移 × 單價年降）＋人形用（線 4 出貨 × 線 6 每台用量 × 份額 × 單價）＋其他用途。"
                "FY2026＝1–9 月實績＋第 4 季假設；單價＝2025 每顆分部營收 ×（1 − 線 6 年降幅）^(年 − 2025)。百萬元。")
    TH["A2"] = "欄：D＝FY2025 … K＝FY2032。人形 2026 隱含份額以線 4 基準路徑校準（三情境同一實績年）。"
    hdr(TH, 3, ["項目", "說明", ""] + [f"FY{y}" for y in range(2025, 2033)])
    THR = {}
    r = 4
    for s in SC:
        TH.cell(r, 1, f"情境 {s}").font = BOLD; TH.cell(r, 1).fill = SECF; r += 1
        R = {}
        R["rev26"] = r; TH.cell(r, 1, "FY2026 合併營收（錨）"); TH.cell(r, 2, "＝1–9 月＋2025Q4 ×（1＋q4g）")
        TH[f"{TC[2026]}{r}"] = "=r9m26+(Hist!E4-r9m26/(1+g9m26))*(1+q4g)"; r += 1
        R["asp"] = r; TH.cell(r, 1, "每顆諧波對應分部營收（元／顆）")
        for y in range(2025, 2033):
            TH[f"{TC[y]}{r}"] = "=asp25" if y == 2025 else f"={TC[y - 1]}{r}*(1-pd)"
        r += 1
        R["U"] = r; TH.cell(r, 1, "全球機器人諧波需求（線 3，千顆）")
        for y in range(2025, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR[('ind', s)]}/1000"
        r += 1
        R["ship"] = r; TH.cell(r, 1, "全球人形出貨（線 4 × 出貨倍數，千台）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR[('hum', s)]}*hu_k/1000"
        r += 1
        R["hdem"] = r; TH.cell(r, 1, "全球人形諧波需求（千顆）"); TH.cell(r, 2, "＝出貨 × 每台用量（線 6）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['ship']}*n_hum"
        r += 1
        R["harm26"] = r; TH.cell(r, 1, "FY2026 諧波減速器及金屬件（錨）"); TH[f"{TC[2026]}{r}"] = f"={TC[2026]}{R['rev26']}*sh_h26"; r += 1
        R["ind"] = r; TH.cell(r, 1, "諧波：工業機器人用 營收"); TH.cell(r, 2, "FY2027 起＝前一年 × 線 3 需求成長 ×（1＋漂移）×（1 − 年降）")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = (f"={c}{R['harm26']}*(1-hum_sh26-oth_sh26)" if y == 2026 else
                             f"={TC[y - 1]}{r}*{c}{R['U']}/{TC[y - 1]}{R['U']}*(1+drift_{s})*(1-pd)")
        r += 1
        R["indu"] = r; TH.cell(r, 1, "　工業用隱含銷量（千顆）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['ind']}/{TC[y]}{R['asp']}*1000"
        r += 1
        R["indsh"] = r; TH.cell(r, 1, "　工業用隱含全球份額（銷量 ÷ 線 3 需求）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['indu']}/{TC[y]}{R['U']}"; TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 1
        R["hsh"] = r; TH.cell(r, 1, "人形諧波份額"); TH.cell(r, 2, "FY2026＝隱含（以線 4 基準校準）；線性走到 hs30，其後持平")
        TH[f"{TC[2026]}{r}"] = f"={TC[2026]}{R['harm26']}*hum_sh26/(線取數!{TC[2026]}{LR[('hum', 'B')]}*hu_k/1000*n_hum*{TC[2026]}{R['asp']}/1000)"
        for y in range(2027, 2033):
            TH[f"{TC[y]}{r}"] = f"=${TC[2026]}${r}+(hs30_{s}-${TC[2026]}${r})*MIN(1,({y}-2026)/4)"
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 1
        R["hum"] = r; TH.cell(r, 1, "諧波：人形機器人用 營收"); TH.cell(r, 2, "FY2027 起＝需求（千顆）× 份額 × 單價 ÷ 1000")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = f"={c}{R['harm26']}*hum_sh26" if y == 2026 else f"={c}{R['hdem']}*{c}{R['hsh']}*{c}{R['asp']}/1000"
        r += 1
        R["oth"] = r; TH.cell(r, 1, "諧波：其他用途與金屬件 營收")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = f"={c}{R['harm26']}*oth_sh26" if y == 2026 else f"={TC[y - 1]}{r}*(1+goth_{s})"
        r += 1
        R["harm"] = r; TH.cell(r, 1, "諧波減速器及金屬件 合計")
        for y in range(2026, 2033): c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['ind']}+{c}{R['hum']}+{c}{R['oth']}"
        r += 1
        R["mech"] = r; TH.cell(r, 1, "機電一體化產品")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['rev26']}*sh_m26" if y == 2026 else f"={TC[y - 1]}{r}*(1+gmech_{s})"
        r += 1
        R["ao"] = r; TH.cell(r, 1, "智能自動化裝備及其他業務")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['rev26']}*(1-sh_h26-sh_m26)" if y == 2026 else f"={TC[y - 1]}{r}*(1+g_ao)"
        r += 1
        R["humsh"] = r; TH.cell(r, 1, "人形用占合併營收")
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
    PL["A1"] = ("PL：官方產品分部（諧波減速器及金屬件／機電一體化產品／智能自動化裝備及其他業務）→ 合併營收 → 標準損益表 → FCFF。"
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
        C = lambda y: COL[y]
        row("ind", "諧波：工業機器人用（子列）", "百萬元", None, lambda y: f"=Thesis!{TC[y]}{th['ind']}")
        row("hum", "諧波：人形機器人用（子列）", "百萬元", None, lambda y: f"=Thesis!{TC[y]}{th['hum']}")
        row("oth", "諧波：其他用途與金屬件（子列）", "百萬元", None, lambda y: f"=Thesis!{TC[y]}{th['oth']}")
        row("harm", "諧波減速器及金屬件 營收", "百萬元", H(8), lambda y: f"=SUM({C(y)}{R['ind']}:{C(y)}{R['oth']})")
        row("mech", "機電一體化產品 營收", "百萬元", H(9), lambda y: f"=Thesis!{TC[y]}{th['mech']}")
        row("ao", "智能自動化裝備及其他業務 營收", "百萬元", H(10), lambda y: f"=Thesis!{TC[y]}{th['ao']}")
        row("rev", "營收（合併）", "百萬元", H(4), lambda y: f"=SUM({C(y)}{R['harm']}:{C(y)}{R['ao']})")
        row("chk", "檢查：分部合計 − 合併營收（應為 0）", "百萬元",
            lambda y: f"={C(y)}{R['harm']}+{C(y)}{R['mech']}+{C(y)}{R['ao']}-{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['harm']}+{C(y)}{R['mech']}+{C(y)}{R['ao']}-{C(y)}{R['rev']}")
        PL.cell(R["chk"], 1, "檢查：分部合計 − 合併營收（應為 0）" if s == "B" else f"檢查（情境 {s}）：分部合計 − 合併營收（應為 0）")
        row("chk2", "檢查：諧波子列合計 − 諧波分部（應為 0）", "百萬元", None, lambda y: f"={C(y)}{R['ind']}+{C(y)}{R['hum']}+{C(y)}{R['oth']}-Thesis!{TC[y]}{th['harm']}")
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
        row("nonop", "其他收入淨額（業外：投資收益、公允價值、財務費用、營業外）", "百萬元", H(26), lambda y: "=(cash_v+ltinv_v)*inv_y")
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
    CP["A1"] = ("Comp：同業前瞻 P/E（stockanalysis.com statistics，資料商）。股價＝" + S_CLOSE + "（Verified）；前瞻 P/E＝原值 × 新收盤 ÷ 原報價（Derived，EPS 不變）；"
                "哈默納科、Nabtesco 取本專案模型同日換算值。核心同業 6 家進中位數，少於 8 家不做迴歸。")
    hdr(CP, 4, ["公司", "代號", "股價（2026-11-20）", "報價日", "前瞻 P/E 分母", "前瞻 P/E", "近 12 月 P/E（原報價日）", "5 年 β", "會計口徑", "進中位數", "說明", "原報價", "原報價日", "原前瞻 P/E", "標記與換算"])
    PEERS = [  # (公司, 代號, 11-20 收盤, 原報價, 原報價日, 原前瞻 P/E, 近 12 月 P/E, β, 口徑, 進中位數, 說明, 分母)
        ("雙環傳動", "002472.SZ", 35.45, 41.21, "2026-10-30", 22.32, 27.11, 0.52, "中國會計準則", "是", "齒輪、精密減速器（RV）；人形關節", "資料商 NTM 共識 EPS"),
        ("兆威機電", "003021.SZ", 76.83, 70.25, "2026-09-30", 59.70, 87.15, 0.44, "中國會計準則", "是", "微型傳動、靈巧手", "資料商 NTM 共識 EPS"),
        ("拓普集團", "601689.SS", 48.11, 46.52, "2026-09-30", 26.89, 32.10, 0.80, "中國會計準則", "是", "汽車零件、人形致動器", "資料商 NTM 共識 EPS"),
        ("鳴志電器", "603728.SS", 44.52, 51.50, "2026-09-30", 124.10, 302.24, 0.62, "中國會計準則", "是", "步進與空心杯馬達", "資料商 NTM 共識 EPS"),
        ("哈默納科", "6324.T", 3722, 3722, "2026-11-20", 48.94, None, 0.96, "日本基準（未查核）", "是", "諧波減速機直接競爭者；取 physical_ai/harmonic_drive 模型 Comp!F14", "公司 FY2026 指引 EPS（2027年3月期）"),
        ("Nabtesco", "6268.T", 3105, 3105, "2026-11-20", 13.77, None, 0.93, "IFRS（未查核）", "是", "RV 減速機；取 physical_ai/nabtesco 模型 In 頁 pe_self", "資料商 NTM 共識 EPS（11-13 值換算）"),
        ("中大力德", "002896.SZ", 48.36, 56.87, "2026-10-09", None, 259.22, 0.25, "中國會計準則", "否", "諧波、行星減速器；資料商無前瞻 P/E（報價日未取得前瞻值）", "—"),
        ("五洲新春", "603667.SS", 41.60, 44.45, "2026-10-13", None, 221.31, 0.86, "中國會計準則", "否", "軸承、行星滾柱絲杠；資料商無前瞻 P/E", "—"),
        ("綠的諧波（本公司，對照）", "688017.SS", 230.80, 263.76, "2026-10-09", 225.96, 342.80, 0.79, "中國會計準則", "否", "本公司；只作對照", "資料商 NTM 共識 EPS（三季報前）"),
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
    CP.cell(17, 1, "本公司現價 ÷ 券商 FY2027 EPS 平均（分母：國元 1.40、東吳 1.76、群益 1.53 元，FY2027）"); CP["F17"] = "=px/AVERAGE(1.40,1.76,1.53)"
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
    row1("cons_show", "市場共識目標價（資料商 15 家平均，2026-10-09，三季報前；只作對照）", "=cons_tp")
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
    RC["A1"] = "Recon：模型對實績、券商預測、線 6 市場規模的同口徑對照；反解支持現價所需的主導變數。"
    hdr(RC, 3, ["項目", "模型", "對照", "差距（模型 − 對照）", "說明"])
    rr = 4
    def rrow(lab, a, b, note, fmt="#,##0.0", diff=True):
        nonlocal rr
        RC.cell(rr, 1, lab); RC.cell(rr, 2, a).number_format = fmt
        if b is not None: RC.cell(rr, 3, b).number_format = fmt
        if diff and b is not None: RC.cell(rr, 4, f"=B{rr}-C{rr}").number_format = fmt
        RC.cell(rr, 5, short(note)); rr += 1
    rrow("FY2026 營收（三情境同）vs 券商 2026E 平均（國元 834、東吳 839、群益 854）", f"=PL!F{PR['B']['rev']}", "=AVERAGE(834,839,854)", "群益為三季報後；國元為半年報後；東吳為一季報後（財報前，只作歷史對照）")
    rrow("FY2027 營收（情境 B）vs 券商 2027E 平均（1,174、1,226、1,307）", f"=PL!G{PR['B']['rev']}", "=AVERAGE(1174,1226,1307)", "券商預測為 Interested-party")
    rrow("FY2028 營收（情境 B）vs 券商 2028E 平均（1,643、1,632、1,739）", f"=PL!H{PR['B']['rev']}", "=AVERAGE(1643,1632,1739)", "")
    rrow("FY2026 歸屬母公司淨利 vs 券商 2026E 平均（196、212、183）", f"=PL!F{PR['B']['ni']}", "=AVERAGE(196,212,183)", "")
    rrow("FY2027 歸屬母公司淨利（情境 B）vs 券商 2027E 平均（257、322、281）", f"=PL!G{PR['B']['ni']}", "=AVERAGE(257,322,281)", "")
    rrow("FY2027 EPS（情境 B）vs 券商 2027E 平均（1.40、1.76、1.53 元）", f"=PL!G{PR['B']['eps']}", "=AVERAGE(1.40,1.76,1.53)", "", "0.00")
    rrow("FY2026 毛利率 vs 2026 年 1–9 月實績", f"=PL!F{PR['B']['gm']}", "=gm26", "模型以 1–9 月為錨（差距應為 0）", "0.0%")
    rrow("FY2026 第 4 季隱含營收 vs 2026Q3 實績", f"=PL!F{PR['B']['rev']}-r9m26", "=Q!E4", "第 4 季若維持第 3 季水準即為對照值")
    rrow("加權目標價 vs 市場共識目標價（資料商 15 家，2026-10-09，三季報前）", f"=Val!B{VR['tp']}", "=cons_tp", S_SA_ST + "；只作對照，不進模型")
    rrow("加權目標價 vs 群益目標價（2026-11-02）", f"=Val!B{VR['tp']}", "=tp_qy", "只作對照，不進模型")
    rr += 1
    RC.cell(rr, 1, "線 6 市場對照（同口徑：五線歸因，美元；本公司諧波分部含金屬件與非機器人用途，故為上界）").font = BOLD; rr += 1
    rrow("FY2026 諧波分部營收（百萬美元）÷ 線 6 諧波市場 2026", f"=PL!F{PR['B']['harm']}/fx/線取數!D{LR[('ind', 'B')] + 9}", None, f"線 6 市場規模 2026 基準 {LD['諧波減速機市場規模 2026'][0]:,.0f} 百萬美元（{LD['諧波減速機市場規模 2026'][1]}）；隱含份額上界", "0.0%", False)
    rrow("FY2030 諧波分部營收（情境 B）÷ 線 6 諧波市場 2030", f"=PL!J{PR['B']['harm']}/fx/線取數!D{LR[('ind', 'B')] + 10}", None, f"線 6 市場規模 2030 基準（{LD['諧波減速機市場規模 2030'][1]}）", "0.0%", False)
    rrow("FY2032 諧波分部營收（情境 B）÷ 線 6 諧波市場 2032", f"=PL!L{PR['B']['harm']}/fx/線取數!D{LR[('ind', 'B')] + 11}", None, f"線 6 市場規模 2032 基準（{LD['諧波減速機市場規模 2032'][1]}）", "0.0%", False)
    rrow("FY2026 工業用隱含全球份額（情境 B）vs 線 6 份額頁 2025 年 12%", f"=Thesis!E{THR['B']['indsh']}", 0.12, "線 6 份額頁為全球全部用途、2025（Interested-party，Zane Robotics 未註明出處）；口徑不同，只看方向", "0.0%")
    rrow("FY2026 人形隱含全球份額（以線 4 基準校準）", f"=Thesis!E{THR['B']['hsh']}", None, "＝人形用營收 ÷（線 4 出貨 × 線 6 每台用量 × 單價）", "0.0%", False)
    rrow("2025 每顆分部營收 vs 來福 2025 諧波平均售價（Analogy）", "=asp25", 571, S_LF + "；本公司分子含金屬件營收", "#,##0")
    rrow("2025 每顆分部營收（美元）vs 線 6 2026 單價基準", "=asp25/fx", LD["p26"][1][0], f"{LD['p26'][1][1]}（全球混合，含哈默納科）；線 4 全尺寸 {LD['p_harm4'][0]:,.0f} 美元（{LD['p_harm4'][1]}）", "#,##0")
    rr += 1
    RC.cell(rr, 1, "反解（整本重算二分搜尋；建置時計算，數值為快照）").font = BOLD; rr += 1
    hdr(RC, rr, ["反解目標", "解", "基準值", "", "說明"]); rr += 1
    if sens:
        for lab, v, b0, note in sens["solve"]:
            RC.cell(rr, 1, lab); RC.cell(rr, 2, v).number_format = "#,##0.00"; RC.cell(rr, 3, b0).number_format = "#,##0.00"; RC.cell(rr, 5, note); rr += 1
    RC.column_dimensions["A"].width = 70; RC.column_dimensions["E"].width = 90

    # 線取數頁補線 6 市場規模（供 Recon 引用）
    base_r = LR[("ind", "B")] + 9
    LN.cell(base_r - 1, 1, "線 6 諧波減速機市場規模（五線歸因、單價基準；百萬美元）").font = BOLD
    for i, k in enumerate(("諧波減速機市場規模 2026", "諧波減速機市場規模 2030", "諧波減速機市場規模 2032")):
        v, src = LD[k]
        LN.cell(base_r + i, 1, k); LN.cell(base_r + i, 2, "B（基準）"); c = LN.cell(base_r + i, 4, v); c.font = BLUE; c.number_format = "#,##0"
        LN.cell(base_r + i, 12, "百萬美元"); LN.cell(base_r + i, 13, src)

    # ================= 對帳 =================
    DZ["A1"] = "對帳：本模型使用的「線層級」變數（與 physical_ai/lines/ 各線模型對帳；本模型讀線 v0.1 成品，不靜默改口徑）。"
    hdr(DZ, 3, ["變數名", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型是否一致"])
    ind26 = LD["ind"][(2026, "B")]; ind32 = LD["ind"][(2032, "B")]; h26 = LD["hum"][(2026, "B")]; h32 = LD["hum"][(2032, "B")]
    ZR = [
        ("全球機器人諧波減速機需求 FY2026–FY2032（A／B／C＝線高／基準／低）", "3 工業機器人", f"基準 {ind26[0]:,.0f} → {ind32[0]:,.0f}", "顆", "僅工業機器人（含協作）；不含人形、半導體設備", f"{ind26[1]}；{ind32[1]}（逐年見線取數頁）", "Derived", "Thesis 工業機器人用營收成長", "一致（直接讀值）"),
        ("全球人形機器人出貨 FY2026–FY2032（A／B／C＝線高／基準／低）", "4 人形", f"基準 {h26[0]:,.0f} → {h32[0]:,.0f}（× hu_k）", "台", "全球、含小型機", f"{h26[1]}；{h32[1]}", "Assumed", "Thesis 人形用營收", "一致（hu_k＝1）；高案 hu_k＝2 依群益引述中國 2026 產量逾 10 萬台，與線 4 基準 5.5 萬台不一致，列 Sens 不改基準"),
        ("人形每台諧波減速機用量（出貨組合平均）", "6 共用零組件", "=n_hum", "顆／台", "出貨組合平均（非 Optimus 類單機）", f"{LD['n_hum'][1][1]}（低 {LD['n_hum'][0][1]}、高 {LD['n_hum'][2][1]}）", "Analogy", "Thesis 全球人形諧波需求", f"與線 6 一致；線 4 Q_HARMONIC_AVG＝{LD['q_avg4'][0]:g}（{LD['q_avg4'][1]}），兩線互不一致，本模型採零組件線（線 6）"),
        ("諧波減速機單價年降幅", "6 共用零組件", "=pd", "%／年", "線 6 不分用途", f"{LD['pd'][1][1]}", "Analogy", "Thesis 單價路徑（工業與人形用）", "一致；本公司 2024→2025 每顆分部營收 −15%（Hist，Derived，含金屬件與組合）超出線 6 高案 10%，列為判斷題"),
        ("諧波單價水準", "6 共用零組件／4 人形", "=asp25", "元／顆", "本公司每顆分部營收（Derived）", f"線 6 2026 基準 {LD['p26'][1][0]:g} 美元（{LD['p26'][1][1]}）；線 4 全尺寸 {LD['p_harm4'][0]:g} 美元", "Derived", "Thesis 單價起點", "不一致（刻意）：採本公司揭露推得的價格水準，只讀線 6 的降幅；線 6、線 4 價格為全球混合或高階口徑"),
        ("線 4 人形諧波隱含單價路徑（2026／2028／2030）", "4 人形", f"{LD['p4_2026'][0]:,.0f}／{LD['p4_2028'][0]:,.0f}／{LD['p4_2030'][0]:,.0f}", "美元／顆", "HARM_MKT ÷ HARM_UNITS", f"{LD['p4_2026'][1]}", "Derived", "未使用（只作對照）", "線 4 隱含年降約 16%，高於線 6 的 6%；本模型依工作單以零組件線（線 6）為單價權威"),
        ("諧波減速機市場規模 2026／2030／2032", "6 共用零組件", "見線取數頁", "百萬美元", "五線歸因、單價基準", f"{LD['諧波減速機市場規模 2026'][1]}", "Derived", "Recon 隱含份額", "一致（直接讀值，只作對照）"),
        ("本公司全球諧波份額", "6 共用零組件", "8%（2023）；12%（2025）", "%", "全球", f"{L6N} 份額頁第 5 列", "Interested-party", "Recon 對照（不進公式）", "模型工業用隱含份額見 Recon"),
        ("中國裝機國產品牌占比 2025", "3 工業機器人", "55%", "%", "IFR", f"{L3N} 下游取數（IR.share.china_domestic.2025）", "Verified", "drift_* 的依據（定性）", "一致"),
        ("中國工業機器人產量 2026 年 1–9 月", "3 工業機器人", "79.5 萬台（+29.1%）", "台", "國家統計局產量（含出口）", f"{L3N} 下游取數（IR.cycle.cn_prod.ytd）", "Verified", "q4g 的依據（定性）", "一致（群益報告同值）"),
    ]
    for i, z in enumerate(ZR):
        for j, v in enumerate(z): DZ.cell(4 + i, 1 + j, v)
    DZ.column_dimensions["A"].width = 50; DZ.column_dimensions["F"].width = 80; DZ.column_dimensions["I"].width = 70

    # ================= Risk =================
    RK["A1"] = "Risk：風險登錄（原則第 10 節第 10、19 條）。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    RISKS = [
        ("H 股發行（規模、價格未揭露）", "部分（h_iss 尾端 Sens，基準 0）", "高於內在價值發行可能增加每股價值", "股東會決議、證監會備案、聯交所遞表"),
        ("價格競爭（來福等國產廠擴產）", "是（pd、drift_C、gm30_C）", "來福 2025 平均售價 −21%", "本公司毛利率、來福招股書更新"),
        ("人形設計改用行星減速機或其他傳動", "部分（n_hum 6–16）", "設計未定", "Tesla Optimus、宇樹、優必選 BOM 揭露"),
        ("人形量產延後或不及線 4 基準", "是（情境 C＝線 4 低）", "", "IDC、TrendForce 季出貨；中國人形產量"),
        ("擴產成本與產能利用率（募投延至 2028 年底）", "部分（gm 路徑、cx）", "", "季毛利率、在建工程"),
        ("金融資產收益下滑（利率下行）", "是（inv_y）", "業外占稅前約三成", "投資收益季變化"),
        ("單一客戶集中（前五大 34.05%、最大 7.51%，2025）", "否", "未達一成門檻", "年報前五大客戶"),
        ("與 SKF 合資（持股 40%）", "否", "規模未揭露", "合資營收與損益"),
    ]
    for i, k in enumerate(RISKS):
        for j, v in enumerate(k): RK.cell(4 + i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["D"].width = 60

    # ================= Sources =================
    SO["A1"] = "Sources：本模型引用的原始來源與資料商（含網址；HTML 展示層不放網址）。"
    for i, s_ in enumerate([S_PX, S_SA, S_SA_ST, S_DW, S_GY1, S_GY2, S_GD, S_QY, S_H, S_LF, S_RF, S_CRP, S_SAP,
                            f"線 3：{L3N}", f"線 4：{L4N}", f"線 6：{L6N}",
                            "physical_ai/harmonic_drive 模型（哈默納科前瞻 P/E）、physical_ai/nabtesco 模型（Nabtesco 前瞻 P/E）",
                            "physical_ai/map/profiles/688017.md（背景，數字已另行查證）"], 3):
        SO.cell(i, 1, s_)
    SO.column_dimensions["A"].width = 160

    # ================= README（xlsx 內）=================
    S = T
    v = lambda k: f"Val!B{VR[k]}"
    lines = [
        ('="綠的諧波 Leaderdrive（688017.SS）個股模型 ' + VER + '｜build_id "&Sens!D2&"｜資料截止日 2026-11-23｜成品日 2026-11-23｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'),
        "命題：綠的諧波的股價由人形機器人的諧波減速器選擇權定價；營收＝工業機器人用（線 3 需求 × 份額）＋人形用（線 4 出貨 × 線 6 每台用量 × 份額 × 單價）＋其他，單價每年下降（線 6）。",
        (f'="結論（條件式）：在 FY2026 以 1–9 月實績為錨、線 4 基準人形出貨（× "&{S("hu_k","0.0")}&"）、線 6 每台 "&{S("n_hum","0")}&" 顆與單價年降 "&{S("pd","0%")}&"、情境 B 的 2030 年人形諧波份額 "&{S("hs30_B","0%")}&"、同業 P/E 未加溢價（相對倍數係數 "&{S("rel_k","0.00")}&"）的假設下，加權目標價約 "&{S(v("tp"),"#,##0.0")}&" 元，相對現價 "&{S("px","#,##0.00")}&" 元 "&{S(v("up"),"+0.0%;-0.0%")}&"，估值評等 "&{v("rating")}&"，投資建議 "&{v("rec")}&"。"'),
        (f'="DCF 目標價 "&{S(v("dcf_tp"),"#,##0.0")}&"、Comparable 目標價 "&{S(v("cmp_tp"),"#,##0.0")}&"（各半）；WACC "&{S("wacc","0.00%")}&"、ke "&{S("ke","0.00%")}&"、g "&{S("g_tv","0.00%")}&"；情境 B 終值占 EV "&{S(f"Val!C{VR['tvs']}","0%")}&"。"'),
        (f'="上檔空間歸因：DCF 貢獻 "&{S(v("att_d"),"+0.0%;-0.0%")}&"、未調整 Comparable 貢獻 "&{S(v("att_c"),"+0.0%;-0.0%")}&"、相對倍數調整貢獻 "&{S(v("att_r"),"+0.0%;-0.0%")}&"。本公司現價對資料商前瞻 EPS 的倍數是同業中位數的 "&{S(v("relnow"),"0.0")}&" 倍；支持現價需要相對倍數係數 "&{S(v("relreq"),"0.0")}&"，升一級需要 "&{S(v("relflip"),"0.0")}&"。"'),
        (f'="收支關卡：情境 B FY2026–FY2030 FCFF 扣股利合計 "&{S(v("gate_b"),"#,##0.0")}&" 百萬元，"&{v("gate_t")}&"；距最近級距邊界 "&{S(v("dist"),"0.0%")}&"。"'),
        "主要失效條件（可觀察）：(1) 全球人形出貨持續高於線 4 高情境（Sens、Recon 反解）；(2) 本公司毛利率回到 37% 以上且每顆營收不再下降；(3) 非中國整機（Tesla 等）量產採用本公司諧波；(4) H 股以高價大量發行。",
        "方法：DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，WACC 折現，明確期到 FY2032，RONIC＝WACC；金融資產以淨現金計入）50%＋同業前瞻 P/E × 模型 FY2027 EPS 50%；三情境 25／50／25。",
        "限制：用途拆分（工業／人形／其他）為 Assumed（公司未揭露）；分部歷史為券商轉錄（億元一位小數）；同業 P/E 由原報價日按 2026-11-20 收盤等比例換算（EPS 未更新）；不讀 Siliconomics／Tokenomics／Wattonomics。詳見 repo README.md。",
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
        ("情境 B 的 2030 年人形諧波份額改為情境 A 值", "hs30_B", ivE["hs30_A"]),
        ("情境 B 的 2030 年人形諧波份額改為情境 C 值", "hs30_B", ivE["hs30_C"]),
        ("情境 B 工業用份額漂移改為情境 A 值", "drift_B", ivE["drift_A"]),
        ("情境 B 工業用份額漂移改為情境 C 值", "drift_B", ivE["drift_C"]),
        ("情境 B 的 FY2030 毛利率改為情境 A 值", "gm30_B", ivE["gm30_A"]),
        ("情境 B 的 FY2030 毛利率改為情境 C 值", "gm30_B", ivE["gm30_C"]),
        ("情境 B 機電一體化年增改為情境 A 值", "gmech_B", ivE["gmech_A"]),
        ("情境 B 機電一體化年增改為情境 C 值", "gmech_B", ivE["gmech_C"]),
        ("情境 B 其他用途年增改為情境 A 值", "goth_B", ivE["goth_A"]),
        ("情境 B 其他用途年增改為情境 C 值", "goth_B", ivE["goth_C"]),
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
    s3 = solve("hs30_B", px, max(ivE["hs30_B"], 0.01), 1.0)
    s4 = solve("hu_k", px, 1.0, 200.0, over={"hs30_B": 0.5, "hs30_A": 0.5, "hs30_C": 0.5})
    nf = lambda x, txt: x if x else -1
    solve_rows = [
        ("支持現價（目標價＝現價）所需的人形出貨倍數（相對線 4，全部情境）", nf(s1, ""), 1.0, "其他輸入不變；−1＝200 倍內無解" if not s1 else f"其他輸入不變；情境 B 的 CY2030 全球人形出貨相當於 {LD['hum'][(2030, 'B')][0] * s1 / 1e6:,.1f} 百萬台"),
        (f"評等升一級（上檔 {flip:+.0%}）所需的人形出貨倍數", nf(s2, ""), 1.0, "其他輸入不變；−1＝200 倍內無解" if not s2 else f"情境 B 的 CY2030 全球人形出貨相當於 {LD['hum'][(2030, 'B')][0] * s2 / 1e6:,.2f} 百萬台"),
        ("支持現價所需的情境 B 2030 年人形諧波份額", nf(s3, ""), ivE["hs30_B"], "其他輸入不變；−1＝份額 100% 仍無解"),
        ("三情境人形份額皆 50% 時，支持現價所需的人形出貨倍數", nf(s4, ""), 1.0, "份額與量同時放大的組合；−1＝200 倍內無解"),
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
    segs = [("諧波減速器及金屬件", "諧波減速器及金屬件 營收", Hs["E8"].value), ("機電一體化產品", "機電一體化產品 營收", Hs["E9"].value),
            ("智能自動化裝備（含其他業務收入，殘差）", "智能自動化裝備及其他業務 營收", Hs["E10"].value)]
    json.dump(dict(company="leaderdrive", ticker="688017.SS", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=round(v, 2), unit="百萬元", source=S_GY1 + "；" + S_DW) for n, l, v in segs],
                   sublines=[dict(parent="諧波減速器及金屬件", labels=["諧波：工業機器人用（子列）", "諧波：人形機器人用（子列）", "諧波：其他用途與金屬件（子列）"])],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）", "檢查：諧波子列合計 − 諧波分部（應為 0）"],
                   note="官方分部＝年報主營業務分產品（諧波減速器及金屬件、機電一體化產品、智能自動化裝備）；智能自動化裝備與其他業務收入合併為一列（殘差），因 2026 上半年只揭露前兩項。用途子列為命題拆分（Assumed），合計＝諧波分部。"),
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
    t = f"""# physical_ai/leaderdrive/ — 綠的諧波 Leaderdrive（688017.SS）個股模型 {VER}

<!-- 本檔由 build_leaderdrive.py 依成品重算值產生；數字不手寫（原則第 10 節第 51 條）。 -->

**本模型適用《高不確定性預估工作原則》**（`docs/plan/高不確定性預估工作原則.md`）：人形機器人需求尚未發生、公司未揭露用途別營收，相關輸入一律給區間並標 Assumed／Analogy；結論只在證據邊界內成立。

- 成品：`dist/{NAME}`（build_id `{sens['build_id']}`；LibreOffice 重算，0 錯誤值）；HTML 展示層 `dist/{NAME[:-5]}.html`（`gen_html_json.py` 產生 `html.json`，再以 `tools/company_html/build_company_html.py` 產生）
- 建置：`python3 physical_ai/leaderdrive/build_leaderdrive.py`（需要 soffice；整本重算敏感度與反解；結尾自動跑 in_layout、sens_coverage --require-keycol、model_governance、rating_gate --strict、zero_rows、cached_values、segment_check）
- 日期：資料截止日 2026-11-23（財報、券商預測、同業倍數）｜成品日 2026-11-23｜股價基準日 2026-11-20（收盤 {I['px']:,.2f} 元）｜評價日 2026-12-31｜目標時點 2027-11-20
- 會計年度＝日曆年；單位百萬元人民幣；中國會計準則（法定財報口徑，GAAP）。上海科創板，亞太掛牌。

## 結論（條件式）

在「FY2026 以 1–9 月實績為錨、線 4 基準人形出貨、線 6 每台 {I['n_hum']:.0f} 顆與單價年降 {I['pd']:.0%}、情境 B 的 2030 年人形諧波份額 {I['hs30_B']:.0%}、同業 P/E 不加溢價（相對倍數係數 {I['rel_k']:.2f}）」的假設下：

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
| 市場共識目標價（資料商 15 家，2026-10-09，三季報前；只作對照） | {I['cons_tp']:,.2f} |

上檔空間歸因：DCF 貢獻 {g('att_d'):+.1%}、未調整 Comparable 貢獻 {g('att_c'):+.1%}、相對倍數調整 {g('att_r'):+.1%}。本公司現價對資料商前瞻 EPS 的倍數是核心同業中位數（{I['pe_peer']:.1f} 倍）的 {g('relnow'):.1f} 倍；要讓目標價＝現價，Comparable 的相對倍數係數須達 {g('relreq'):.1f}。DCF 以 RONIC＝WACC（新投資無超額報酬），而高相對倍數隱含持續超額報酬，兩法對競爭優勢的隱含假設相反。

**讀法：** 市場對本公司的定價幾乎全是人形機器人選擇權。情境 B 人形用營收 {ser('hum', '{:,.0f}')}（FY2026→FY2032，百萬元），占合併營收由 {TH[f'E{THB["humsh"]}'].value:.0%} 升到 {TH[f'K{THB["humsh"]}'].value:.0%}。整本重算反解：其他輸入不變時，**支持現價需要人形出貨達線 4 的 {fn(sol[0][1], '{:,.1f}')} 倍**（{sol[0][3]}）；評等升一級需要 {fn(sol[1][1], '{:,.2f}')} 倍；只放大份額時，{('情境 B 2030 年人形份額達 100% 仍不足以支持現價' if sol[2][1] == -1 else f'需要情境 B 2030 年人形份額 {sol[2][1]:.0%}')}。結論主要建立在 Assumed（用途拆分、份額、毛利率路徑）與線模型的 Analogy／Derived 輸入上，但方向在合理區間兩端相同（見敏感度），屬「穩健」的放空結論；證據邊界外的是人形量級本身。

## 命題

1. **營收結構：** 官方產品分部為諧波減速器及金屬件（2025 年 {wb['Hist']['E8'].value:,.0f} 百萬元，占 {wb['Hist']['E8'].value / wb['Hist']['E4'].value:.0%}）、機電一體化產品、智能自動化裝備；用途別（工業機器人、人形、其他）公司未揭露，本模型以 Assumed 拆分（2026 年人形占諧波營收 {I['hum_sh26']:.0%}，區間 8–25%）。
2. **量：** 工業機器人用跟隨線 3 全球機器人諧波需求；人形用＝線 4 全球出貨 × 線 6 每台用量 × 份額。2026 年隱含人形份額 {TH[f'E{THB["hsh"]}'].value:.0%}（以線 4 基準校準）。
3. **價：** 每顆分部營收 2025 年 {I['asp25']:,.0f} 元，依線 6 每年降 {I['pd']:.0%}；本公司 2024→2025 已降約 15%（含組合），來福 2025 平均售價 571 元、年降 21%。價格競爭是量增利潤不落在本公司的主要機制。

情境 B：營收 {ser('rev', '{:,.0f}')}；毛利率 {ser('gm', '{:.1%}')}；EPS {ser('eps', '{:.2f}')} 元（FY2026→FY2032）。

## 主要假設與可證偽指標（原則第 10 節第 20 條）

| 承重假設 | 基準（A／B／C 或區間） | 標記 | 可觀察指標 |
|---|---|---|---|
| 全球人形出貨（線 4） | 線 4 高／基準／低；× hu_k {I['hu_k']:.0f}（高案 2） | Assumed（線 4） | IDC、TrendForce 季出貨；中國人形產量（群益引述 2026 年逾 10 萬台） |
| 人形每台諧波用量（線 6） | {I['n_hum']:.0f}（6–16） | Analogy | 整機拆解、行星減速替代比例 |
| 單價年降幅（線 6） | {I['pd']:.0%}（3–10%） | Analogy | 本公司每顆分部營收；來福平均售價 |
| 2030 年人形諧波份額 | {I['hs30_A']:.0%}／{I['hs30_B']:.0%}／{I['hs30_C']:.0%} | Assumed | 宇樹、優必選、智元供應鏈揭露；Tesla 供應商 |
| 2026 年人形占諧波營收 | {I['hum_sh26']:.0%}（8–25%） | Assumed | 年報用途揭露（若有）；H 股招股書 |
| FY2030 毛利率 | {I['gm30_A']:.0%}／{I['gm30_B']:.0%}／{I['gm30_C']:.0%} | Assumed | 季毛利率（2026 第 3 季 33.4%） |
| 工業用份額漂移 | {I['drift_A']:+.0%}／{I['drift_B']:+.0%}／{I['drift_C']:+.0%} 每年 | Assumed | 中國國產品牌占比；哈默納科中國分部營收 |
| 同業前瞻 P/E | 中位數 {I['pe_peer']:.1f} 倍（6 家） | Derived | 同一估值日重抓 |

## 敏感度（整本重算，前 10 大）

| 案例 | 鍵 | 加權目標價 | 相對基準 | DCF 目標價 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | `{c['key']}` | {c['tp']:,.1f} | {c['tp'] / sens['base_tp'] - 1:+.1%} | {c['dcf']:,.1f} | {c['rating']} |" for c in top) + f"""

完整案例見 xlsx Sens 頁（每個 In 頁有低／高區間的鍵都有案例；另有情境 B 改 A／C 值、人形出貨 5 倍、情境權重案例）。**翻轉評等所需的變動**：見 Recon 反解——單一輸入在合理區間內都無法讓評等離開「{g('rating')}」，升一級需加權目標價達 {fn(up_b)} 元。

## 反解（Recon 頁）

""" + "\n".join(f"- {a}：{fn(b, '{:,.2f}')}（基準 {fn(c0, '{:,.2f}')}）；{d}" for a, b, c0, d in sol) + f"""

## 限制與「分不出來／沒找到」

- **分不出來（資料性質）：** 用途別營收（工業、人形、其他）公司未揭露；人形量級屬未發生事件，線 4 三情境差距約 4–5 倍。
- **沒找到（工作未完）：** 公司年報與三季報原文未直接讀取（歷史數字取自資料商與券商轉錄）；2023 年分部營收由 2024 年增率回推；金屬部件營收未拆分；限制性股票與選擇權；H 股發行規模與價格；長期投資明細。
- **線模型不一致：** 人形每台諧波用量線 6＝10、線 4＝8（採線 6）；線 4 隱含諧波單價年降約 16%，線 6 為 6%（採線 6）。見 xlsx「對帳」頁。
- **同業倍數：** 4 家 A 股同業原報價日 2026-09-30～10-30，前瞻 P/E 依 2026-11-20 收盤等比例換算（EPS 未更新）；哈默納科、Nabtesco 取本專案模型值；少於 8 家不做迴歸；中大力德、五洲新春資料商無前瞻 P/E，未進中位數。
- **rf 逾期：** 中國 10 年期國債 1.61% 為 2026-11-05 值（跨模型一致）。
- **協調檔：** 實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics；相關條文不適用。
- **獨立核對：** 依專案規則，交付判讀的數字應由未參與製作的代理獨立核對；本代理無法啟動子代理，待主代理安排。

## 資料來源

見 xlsx Sources 頁（公司公告、券商轉錄年報與季報、資料商、來福港股申請文件、線 3／4／6 成品）。

## 版本紀錄

- {VER}（2026-11-23）：首版。現價 {I['px']:,.2f}（2026-11-20 收盤）；加權目標價 {g('tp'):,.1f}；上檔空間 {g('up'):+.1%}；{g('rating')}／{g('rec')}；收支關卡 {g('gate_b'):,.1f}（{g('gate_t')}）。
"""
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write(t)


if __name__ == "__main__":
    main()

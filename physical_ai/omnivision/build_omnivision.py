# 豪威集團／韋爾股份 OmniVision（603501.SS，上海證券交易所）個股模型 v0.1 建置腳本
# 產出：physical_ai/omnivision/dist/20261124_豪威集團個股模型_v0.1.xlsx（LibreOffice 重算後的成品，含公式與快取值）
# 結構：命題模組（Thesis：設計業務營收＝汽車市場＋新興市場（含人形用相機，線 4 出貨 × 線 6 每台顆數 × 單價 × 本公司份額）＋智慧手機及其他；
#       半導體分銷另列）→ 官方分部（半導體設計／半導體分銷／其他）→ 標準損益表 → FCFF DCF＋同業 P/E。
# 敏感度：每個案例覆寫 In 頁基準欄後整本重建、LibreOffice 重算（整本重算，不用閉式近似）；反解同法以二分搜尋。
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md、docs/plan/高不確定性預估工作原則.md。
# 用法：python3 physical_ai/omnivision/build_omnivision.py   （需要 soffice；暫存於 $HOME/ov_build）
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
NAME = f"20261124_豪威集團個股模型_{VER}.xlsx"
OUT = os.path.join(HERE, "dist", NAME)
HOMEDIR = os.path.expanduser("~")
TMP = os.path.join(HOMEDIR, "ov_build")
LO = ["soffice", f"-env:UserInstallation=file://{HOMEDIR}/lo_omnivision", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML"]

L4 = os.path.join(ROOT, "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx")
L6 = os.path.join(ROOT, "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx")
L4N, L6N = (os.path.relpath(p, ROOT) for p in (L4, L6))

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDRF = PatternFill("solid", fgColor="DDEBF7")
SECF = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------- 出處 ----------------
S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv：603501.SS；未還原股利）"
S_SA = "stockanalysis.com SHA:603501（資料商，來源 S&P Global Market Intelligence；年度與季度損益表、資產負債表、現金流量表，2026-11-24 讀取；法定財報原文未直接讀取）"
S_SA_ST = "stockanalysis.com SHA:603501 statistics 頁（資料商；頁面報價 2026-11-24 收盤時段 86.68 元，2026-11-24 讀取）"
S_SEG = "2025 年年報分產品與分市場數字（轉載自證券之星、鳳凰科技、長江證券研報摘要的搜尋結果；2026-11-24 讀取；年報原文未直接讀取）：半導體設計 238.00 億（82.60%）、半導體代理 49.05 億（17.02%）、圖像傳感器 212.46 億（+10.71%）、模擬 16.13 億（+13.43%）、顯示 9.41 億（−8.47%）、汽車 74.71 億（+26.52%）、新興市場 23.69 億（+211.85%）、智能手機 −15.61%"
S_Q3 = "2026 年三季報摘要（轉載自 IT 之家、證券時報、中泰證券研報摘要的搜尋結果；2026-11-24 讀取）：前三季營收 221.49 億（+1.68%）、歸母淨利 22.99 億（−28.39%）、半導體設計 163.46 億（−9.40%）、代理 56.86 億（+55.67%）、毛利率約 29.2%；第 4 季預告營收 77.59–83.24 億、毛利率 29.44–30.74%"
S_IR = "豪威集團投資者關係活動紀錄（同花順轉載，2026-10-30）：汽車業務預期持續回升；全域快門感測器部分料號 Q4 量產（人形、工廠自動化）"
S_RF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與 leaderdrive、innolight 同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：China CRP 0.91%（與 leaderdrive 同一取值）"
S_CLOSE = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
S_PEER = "stockanalysis.com statistics 頁（資料商；2026-11-24 讀取；各檔原報價日見 Comp 頁 M 欄）"
IP_SA = "資料商彙整賣方共識，未必含最新財報"
IP_CO = "公司說法與預告，有股價與融資誘因"
IP_BRK = "券商或轉載媒體：有銷售研究與維持評等的誘因"

# ---------------- 線模型取數（以 ID／代號找列，不以列號） ----------------
def line_data():
    wb4 = openpyxl.load_workbook(L4, data_only=True); ws4 = wb4["下游取數"]
    l4 = {}
    for r in range(1, ws4.max_row + 1):
        k = ws4.cell(r, 1).value
        if isinstance(k, str): l4[k] = (r, [ws4.cell(r, c).value for c in range(1, 10)])
    wb6 = openpyxl.load_workbook(L6, data_only=True); ws6 = wb6["輸入"]; wc = wb6["計算"]
    d = {"hum": {}}
    for y in range(2026, 2033):
        r, v = l4[f"SHIP_{y}"]
        for s, c, col in (("A", 4, "E"), ("B", 3, "D"), ("C", 2, "C")):
            d["hum"][(y, s)] = (v[c], f"{L4N} 下游取數!{col}{r}（SHIP_{y}）")
    cam = [r for r in range(1, ws6.max_row + 1) if ws6.cell(r, 1).value == "CAM" and ws6.cell(r, 3).value == "顆"]
    d["cam_n"] = [(ws6.cell(r, 6).value, f"{L6N} 輸入!F{r}") for r in cam[:3]]      # 低、基準、高：人形每台相機顆數
    pr = [r for r in range(1, ws6.max_row + 1) if ws6.cell(r, 1).value == "CAM" and ws6.cell(r, 3).value == "美元/顆"][0]
    L = openpyxl.utils.get_column_letter
    d["cam_p"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{L(c)}{pr}") for c in (4, 5, 6)]
    d["cam_pd"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{L(c)}{pr}") for c in (7, 8, 9)]
    cv = [r for r in range(1, ws6.max_row + 1) if ws6.cell(r, 1).value == "CAMV" and ws6.cell(r, 3).value == "顆"]
    d["veh_n"] = [(ws6.cell(r, 8).value, f"{L6N} 輸入!H{r}") for r in cv[:3]]
    for rr, nm in ((89, "low"), (156, "base"), (223, "high")):
        assert wc.cell(rr, 1).value == "CAMV", rr
        d["veh_mkt_" + nm] = (wc.cell(rr, 3).value, f"{L6N} 計算!C{rr}")
    assert wc.cell(155, 1).value == "CAM"
    d["rob_mkt30"] = (wc.cell(155, 7).value, f"{L6N} 計算!G155")
    return d


LD = line_data()

# ---------------- In 頁輸入 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 來源誘因, 說明)
IN = [
    ("sec", "1. 市場資料與日期"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, "", "資料截止日 2026-11-24；股價 2026-11-20；成品日 2026-11-24"),
    ("px", "現價（2026-11-20 收盤）", "元", None, 83.91, None, "Verified", S_PX, "", "csv 值 83.910004；11-24 資料商頁面 86.68"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,12,31)", None, "Assumed", "本專案慣例：評價日 2026-12-31（會計年度＝日曆年，FY2026 在評價日已結束）", "", "折現期間由評價日起算"),
    ("tp_date", "目標時點（px_date 後 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", ""),
    ("bs_date", "資產負債表日", "日期", None, "=DATE(2026,9,30)", None, "Verified", S_SA, "", "2026 年三季報"),
    ("stub_frac", "Stub 期（資產負債表日→評價日）", "年", None, "=(val_date-bs_date)/365", None, "Derived", "＝(評價日 − 資產負債表日) ÷ 365", "", "Stub FCFF＝FY2026 FCFF × stub_frac"),
    ("sh_v", "評價日股數（A＋H 總股本）", "百萬股", None, 1250.0, None, "Interested-party", S_SA_ST + "：股數 12.5 億；同頁市值 1,068 億元 ÷ 股價 86.68 ≈ 12.32 億，兩者差 1.5%", IP_SA, "H 股 2026-01 上市，已含；限制性股票與選擇權本輪未取得揭露，以 0 計；差 1.5% 不改評等"),
    ("cash_v", "現金＋短期投資（2026-09-30）", "百萬元", None, 15494.0, None, "Verified", S_SA + "：現金 14,880、交易性金融資產 615", "", "2025-12-31 為 12,946"),
    ("ltinv_v", "長期投資（2026-09-30）", "百萬元", None, 6535.0, None, "Verified", S_SA, "", "明細未取得（含策略性股權投資，流動性與變現價值不確定）；2025-12-31 為 4,392"),
    ("inv_k", "長期投資計入淨現金的比例", "%", 0.0, 0.5, 1.0, "Assumed", "明細未取得；基準計入一半、低案 0、高案全額", "", ""),
    ("debt_v", "有息負債（2026-09-30，含租賃負債 159）", "百萬元", None, 13444.0, None, "Verified", S_SA, "", "2025-12-31 為 9,688；負債一年內增加 3,756，與現金同步增加"),
    ("nci_v", "少數股東權益（2026-09-30）", "百萬元", None, -19.38, None, "Verified", S_SA, "", "負值（少數股東虧損），自淨現金扣除即小幅加回"),
    ("fx", "人民幣兌美元（2026-11-20）", "元／美元", None, 6.6665, None, "Verified", S_CLOSE + "：CNY=X", "", "用於人形相機營收換算與 Recon 對照線 6 市場規模"),
    ("cons_tp", "賣方平均目標價（18 家，頁面報價 2026-11-24）", "元", None, 112.33, None, "Interested-party", S_SA_ST, IP_SA, "只作對照，不進模型"),
    ("sec", "2. 實績錨（FY2026＝2026 年，會計年度＝日曆年）"),
    ("r9m26", "2026 年 1–9 月營收", "百萬元", None, "=Q!E4", None, "Derived", "＝Q 頁 Q1＋Q2＋Q3（22,149；三季報摘要 221.49 億，一致）", "", ""),
    ("q4rev", "2026 年第 4 季營收（公司預告區間 7,759–8,324）", "百萬元", 7759.0, 8041.5, 8324.0, "Interested-party", S_Q3, IP_CO, "預告區間，基準取中點；FY2026 營收＝1–9 月實績＋本值"),
    ("q4gm", "2026 年第 4 季毛利率（公司預告區間 29.44%–30.74%）", "%", 0.2944, 0.3009, 0.3074, "Interested-party", S_Q3, IP_CO, "預告區間，基準取中點"),
    ("dist_sh26", "FY2026 半導體分銷 占營收", "%", 0.22, 0.256, 0.28, "Assumed", S_Q3 + "：前三季 25.67%（2025 全年 17.0%）", "", "第 4 季假設同前三季；公司稱受 AI 產業鏈需求拉動，持續性未知"),
    ("oth_sh26", "FY2026 其他業務 占營收", "%", None, 0.005, None, "Assumed", "2025 年其他業務約 150 百萬元（0.5%，Derived）", "", "半導體設計＝營收 − 分銷 − 其他"),
    ("gm26", "FY2026 毛利率", "%", None, "=(Q!E5+q4rev*q4gm)/(r9m26+q4rev)", None, "Derived", "＝（前三季毛利＋第 4 季預告營收 × 預告毛利率）÷ FY2026 營收", "", ""),
    ("opex26", "FY2026 營業費用及其他營業損益合計 ÷ 營收", "%", 0.14, "=Q!E7/Q!E4", 0.165, "Derived", "＝前三季（毛利 − 營業利益）÷ 營收（Q 頁）；資料商未提供費用分項", "", "含銷管研費用、稅金及附加、減值與其他收益，不可拆；低／高為 Assumed"),
    ("ga26", "FY2026 汽車市場營收年增", "%", 0.0, 0.08, 0.15, "Assumed", S_SEG + "：2025 年 +26.5%；" + S_IR + "：預期持續回升；前三季未揭露汽車金額", "", "2026 年分市場營收未揭露"),
    ("ge26", "FY2026 新興市場營收年增", "%", 0.0, 0.30, 0.80, "Assumed", S_SEG + "：2025 年 +211.9%（基期小）；2026 年未揭露", "", ""),
    ("sec", "3. 命題模組：人形機器人相機（線 4 出貨 × 線 6 每台顆數 × 單價 × 相機價值中本公司感測器占比 × 本公司份額）"),
    ("hu_k", "人形出貨相對線 4 的倍數（全部情境）", "x", None, 1.0, 2.0, "Assumed", "基準＝線 4 原值；高案 2 倍（中國 2026 人形產量高於線 4 基準的報導，見線 4 README）", "", "反解變數"),
    ("cam_n", "人形每台相機顆數（線 6）", "顆／台", LD["cam_n"][0][0], LD["cam_n"][1][0], LD["cam_n"][2][0], "Analogy", f"線 6：{LD['cam_n'][1][1]}（低 {LD['cam_n'][0][1]}、高 {LD['cam_n'][2][1]}）；線 6 註記人形頭部 2–8 顆", "", ""),
    ("cam_p", "相機模組單價 2026（線 6）", "美元／顆", LD["cam_p"][0][0], LD["cam_p"][1][0], LD["cam_p"][2][0], "Assumed", f"線 6：{LD['cam_p'][1][1]}（低 {LD['cam_p'][0][1]}、高 {LD['cam_p'][2][1]}）；線 6 標記未取得來源", "", "模組單價，非感測器單價"),
    ("pdc", "相機單價年降幅（線 6）", "%／年", LD["cam_pd"][0][0], LD["cam_pd"][1][0], LD["cam_pd"][2][0], "Assumed", f"線 6：{LD['cam_pd'][1][1]}（低 {LD['cam_pd'][0][1]}、高 {LD['cam_pd'][2][1]}）", "", ""),
    ("cis_val", "感測器占相機模組價值比例", "%", 0.20, 0.30, 0.40, "Assumed", "本專案假設，未找到可引用的拆解來源", "", "線 6 未提供；模組含鏡頭、封裝、組裝"),
    ("ov_sh", "本公司在人形相機感測器的份額", "%", 0.10, 0.25, 0.40, "Assumed", "未取得；公司稱全域快門感測器鎖定人形與工廠自動化（" + S_IR + "）", "", "份額未取得，給寬區間"),
    ("sec", "4. 情境專屬驅動（A 樂觀／B 基準／C 保守；只列基準欄，不設敏感度區間）"),
    ("gauto_A", "情境 A：汽車市場營收年增（FY2027 起）", "%", None, 0.18, None, "Assumed", S_SEG + "：2025 年 +26.5%；線 6 輕型車每車平均 " + f"{LD['veh_n'][1][0]:g} 顆（{LD['veh_n'][0][0]:g}–{LD['veh_n'][2][0]:g}；{LD['veh_n'][1][1]}）", "", "含相機顆數增加與高階化"),
    ("gauto_B", "情境 B：汽車市場營收年增（FY2027 起）", "%", None, 0.10, None, "Assumed", "同上；2026 年管理層預期汽車業務持續回升", "", ""),
    ("gauto_C", "情境 C：汽車市場營收年增（FY2027 起）", "%", None, 0.03, None, "Assumed", "同上；整車需求與價格壓力", "", ""),
    ("gem_A", "情境 A：新興市場（不含人形子列）營收年增（FY2027 起）", "%", None, 0.40, None, "Assumed", S_SEG + "：2025 年 +211.9%；機器視覺、AR／AI 眼鏡、醫療", "", ""),
    ("gem_B", "情境 B：新興市場（不含人形子列）營收年增（FY2027 起）", "%", None, 0.20, None, "Assumed", "同上", "", ""),
    ("gem_C", "情境 C：新興市場（不含人形子列）營收年增（FY2027 起）", "%", None, 0.05, None, "Assumed", "同上", "", ""),
    ("gph_A", "情境 A：智慧手機及其他營收年增（FY2027 起）", "%", None, 0.03, None, "Assumed", S_SEG + "：2025 年智能手機 −15.6%；" + S_Q3 + "：設計業務 2026 前三季 −9.4%", "", "含類比與顯示；2026 隱含年增見 Thesis"),
    ("gph_B", "情境 B：智慧手機及其他營收年增（FY2027 起）", "%", None, -0.03, None, "Assumed", "同上", "", ""),
    ("gph_C", "情境 C：智慧手機及其他營收年增（FY2027 起）", "%", None, -0.10, None, "Assumed", "同上", "", ""),
    ("gdist_A", "情境 A：半導體分銷營收年增（FY2027 起）", "%", None, 0.10, None, "Assumed", S_Q3 + "：2026 前三季 +55.7%、2025 全年 17.0% 占比；公司稱受 AI 產業鏈需求拉動", "", "持續性未知，區間寬"),
    ("gdist_B", "情境 B：半導體分銷營收年增（FY2027 起）", "%", None, -0.08, None, "Assumed", "同上；基準假設逐年回落", "", ""),
    ("gdist_C", "情境 C：半導體分銷營收年增（FY2027 起）", "%", None, -0.25, None, "Assumed", "同上；回到 2025 年水準方向", "", ""),
    ("gmd30_A", "情境 A：FY2030 設計業務毛利率", "%", None, 0.40, None, "Assumed", "2025 約 34–35%、2026 約 37%（Derived，依分銷毛利假設；公司稱設計業務結構優化）", "", "FY2026 起線性走到 FY2030，其後持平"),
    ("gmd30_B", "情境 B：FY2030 設計業務毛利率", "%", None, 0.37, None, "Assumed", "同上", "", ""),
    ("gmd30_C", "情境 C：FY2030 設計業務毛利率", "%", None, 0.33, None, "Assumed", "同上；手機高階化不順與價格競爭", "", ""),
    ("tp_wA", "情境權重 A", "", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", ""),
    ("tp_wB", "情境權重 B", "", None, 0.50, None, "Assumed", "同上", "", ""),
    ("tp_wC", "情境權重 C", "", None, 0.25, None, "Assumed", "同上", "", ""),
    ("sec", "5. 損益與現金流"),
    ("gmdist", "半導體分銷毛利率", "%", 0.05, 0.08, 0.12, "Assumed", "分銷毛利率未找到；公司稱分銷占比升高拉低整體毛利率（" + S_Q3 + "）", "", "用於拆分設計與分銷毛利；整體毛利率不受影響（FY2026 錨），影響混合變化後的毛利率"),
    ("opex30", "FY2030 營業費用及其他合計 ÷ 營收", "%", 0.125, 0.14, 0.155, "Assumed", "2023 16.2%、2024 15.4%、2025 14.4%、2026 前三季 15.2%（Hist／Q，Derived）", "", "FY2026→FY2030 線性，其後持平"),
    ("nonop_r", "業外淨額 ÷ 營收", "%", -0.025, -0.015, 0.004, "Assumed", "2025 年約 +0.4%（Hist，近似）；2026 前三季淨利 ÷ (1 − 稅率) 與營業利益相比，業外約 −2%（Derived，近似；負債 134 億的利息與匯兌）", "", "基準取 −1.5%（介於 2026 前三季與 2025）；高案＝2025 年水準"),
    ("tax_r", "有效稅率", "%", 0.08, "=Hist!E19/Hist!E17", 0.18, "Derived", "＝2025 年（稅前 − 淨利）÷ 稅前；稅前為近似值", "", "敏感度低／高為 Assumed"),
    ("payout", "配息率（現金股利 ÷ 前一年淨利）", "%", None, "=Hist!E23/Hist!D20", None, "Derived", "＝2025 年現金股利支付 978.54 ÷ 2024 年淨利 3,323", "", "用於收支關卡與目標時點前股利"),
    ("da26", "FY2026 折舊攤銷 ÷ 營收", "%", None, "=Hist!E24/Hist!E4", None, "Derived", "＝2025 年折舊攤銷 835 ÷ 營收", "", "之後線性收斂到資本支出比例 × da_ss（FY2032）"),
    ("da_ss", "穩態折舊 ÷ 資本支出", "x", None, 0.8, None, "Assumed", "穩態資本支出高於折舊（成長）", "", ""),
    ("cx26", "FY2026 資本支出", "百萬元", 1500, 2000, 2600, "Assumed", "2025 年 2,336、2024 年 1,249（Hist，資料商）；2026 未取得指引", "", ""),
    ("cx_r", "資本支出 ÷ 營收（FY2027 起）", "%", 0.05, 0.07, 0.09, "Assumed", "2023 年 5.0%、2024 年 4.9%、2025 年 8.1%（Hist，Derived）", "", ""),
    ("nwc_r", "營運資金增加 ÷ 營收增量", "%", 0.15, 0.30, 0.45, "Assumed", S_SA + "：2025-12→2026-09 存貨 +777、應收 +2,149、應付 +580", "", "存貨與應收水位高"),
    ("h_iss", "H 股發行占擴大後股本（尾端事件；H 股已於 2026-01 上市，基準 0）", "%", None, 0.0, None, "Assumed", "H 股已上市，股數已含；新增發行未見公告", "", "保留供公式引用，不設敏感度"),
    ("h_disc", "新發行價相對現價折價", "%", None, 0.10, None, "Assumed", "只在 h_iss＞0 時有作用", "", ""),
    ("sec", "6. 折現率（DCF 標準規格第 2 節）"),
    ("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", S_RF, "", "日期 2026-11-05，較股價基準日早 11 個交易日（逾期；跨模型一致，下次全面改版更新）"),
    ("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", ""),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", ""),
    ("b_own", "本公司 5 年 β（原始）", "x", None, 0.86, None, "Verified", S_SA_ST, "", ""),
    ("b_peer", "同業 5 年 β 中位數（原始）", "x", None, "=MEDIAN(Comp!H5:H8)", None, "Derived", "＝Comp 頁核心同業 4 家 β 中位數", "", ""),
    ("beta", "β（兩來源 Blume 調整後中位數）", "x", "=beta-0.3", "=MEDIAN(0.67*b_own+0.33,0.67*b_peer+0.33)", "=beta+0.3", "Derived", "Blume 0.67β＋0.33；兩來源＝本公司、同業中位數", "", "敏感度 ∓0.3（跨模型慣例）"),
    ("kd", "稅前債務成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt", "", "有息負債 134 億，占市值約 13%"),
    ("wacc_adj", "WACC 調整（敏感度用）", "pt", -0.01, 0, 0.01, "Assumed", "基準 0；敏感度 ±1pt", "", ""),
    ("g_cap", "永續成長率上限", "%", None, 0.03, None, "Assumed", "DCF 標準規格：3%，且不超過計價幣別 10 年期公債殖利率", "", ""),
    ("g_tv", "永續成長率", "%", 0.01, "=MIN(g_cap,rf)", 0.03, "Derived", "＝MIN(g_cap, rf)", "", "敏感度 1%／3%（3% 超過 rf，只作敏感度）"),
    ("ronic_sp", "RONIC − WACC", "pt", None, 0, 0.05, "Assumed", "RONIC 預設＝WACC（新投資價值中性）", "", ""),
    ("tv_alt", "終值改用 FCFF ×（1＋g）（0＝否，1＝是）", "開關", None, 0, 1, "Assumed", "DCF 標準規格第 4 節敏感度", "", ""),
    ("sec", "7. 同業倍數與權重"),
    ("pe_peer", "同業前瞻 P/E 中位數", "x", "=MIN(Comp!F5:F8)", "=MEDIAN(Comp!F5:F8)", "=MAX(Comp!F5:F8)", "Derived", "＝Comp 頁核心同業 4 家前瞻 P/E 中位數；股價皆為 2026-11-20 收盤、P/E 由原報價日等比例換算（EPS 分母不變，見 Comp）", "", "敏感度＝同業最低／最高；同業僅 4 家，區間寬"),
    ("rel_k", "相對倍數係數（本公司 ÷ 同業）", "x", None, 1.0, 2.0, "Assumed", "1.00（未調整：缺本公司對同業十年相對倍數歷史）", "", "高案 2 倍（未查核，不作依據）"),
    ("w_dcf", "DCF 權重（Comparable＝1 − 本值）", "%", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設 DCF 50%＋同業倍數 50%", "", ""),
]
SCN_DRIVERS = [("gauto", "汽車市場營收年增"), ("gem", "新興市場營收年增（不含人形）"), ("gph", "智慧手機及其他營收年增"),
               ("gdist", "半導體分銷營收年增"), ("gmd30", "FY2030 設計業務毛利率")]

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
    HI["A1"] = "Hist：歷史財務（百萬元人民幣）。資料商轉錄（S&P Global）；稅前利益由資料商稅前利潤率 × 營收推算（近似）；分部由年報增率回推 2024（Derived）。2023 年分部未找到。公式只引用本頁（定義名稱 h_*）。"
    hdr(HI, 3, ["項目", "", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"])
    HROWS = [  # (列, 鍵, 項目, [2023, 2024, 2025], 標記, 來源)
        (4, "h_rev", "營業收入", [21021, 25731, 28855], "Verified", S_SA),
        (5, "h_cogs", "營業成本（＝營收 − 毛利）", ["=C4-C6", "=D4-D6", "=E4-E6"], "Derived", "＝營業收入 − 毛利"),
        (6, "h_gp", "毛利", [4176, 7247, 8639], "Verified", S_SA),
        (7, "h_gm", "毛利率", ["=C6/C4", "=D6/D4", "=E6/E4"], "Derived", "＝毛利 ÷ 營收"),
        (8, "h_dsg", "半導體設計 營收（2024＝由 2025 年各產品增率回推）", [None, "=21246/1.1071+1613/1.1343+941/(1-0.0847)", 23800], "Verified", S_SEG + "；2024＝2025 年圖像傳感器 21,246、模擬 1,613、顯示 941 各除以（1＋增率）"),
        (9, "h_agt", "半導體分銷 營收（2024＝營收 − 設計，含其他業務）", [None, "=D4-D8", 4905], "Verified", S_SEG),
        (10, "h_oth", "其他業務 營收（殘差）", [None, None, "=E4-E8-E9"], "Derived", "＝營收 − 設計 − 分銷"),
        (11, "h_chk", "檢查：分部合計 − 營業收入（應為 0）", [None, "=D8+D9+D10-D4", "=E8+E9+E10-E4"], "Derived", ""),
        (12, "h_auto", "設計：汽車市場 營收（2024＝2025 ÷ 1.2652）", [None, "=E12/1.2652", 7471], "Verified", S_SEG + "；汽車市場口徑可能含非感測器產品（如類比），視為上界"),
        (13, "h_emer", "設計：新興市場 營收（2024＝2025 ÷ 3.1185）", [None, "=E13/3.1185", 2369], "Verified", S_SEG),
        (14, "h_phone", "設計：智慧手機及其他（殘差；含類比與顯示）", [None, "=D8-D12-D13", "=E8-E12-E13"], "Derived", "＝設計 − 汽車 − 新興"),
        (15, "h_opex", "營業費用及其他營業損益合計（＝毛利 − 營業利益）", ["=C6-C16", "=D6-D16", "=E6-E16"], "Derived", "資料商未提供銷管研分項；含稅金及附加、減值、其他收益"),
        (16, "h_op", "營業利益", [776.6, 3280, 4479], "Verified", S_SA),
        (17, "h_pbt", "稅前利益（＝營收 × 資料商稅前利潤率 3.29%／12.74%／15.94%）", ["=C4*0.0329", "=D4*0.1274", "=E4*0.1594"], "Derived", S_SA + "：稅前利潤率；金額為近似"),
        (18, "h_nonop", "業外（＝稅前 − 營業利益）", ["=C17-C16", "=D17-D16", "=E17-E16"], "Derived", "近似；2024 年稅前低於淨利，資料商口徑差異（見 README 懷疑項）"),
        (19, "h_tax", "所得稅（＝稅前 − 淨利，近似）", ["=C17-C20", "=D17-D20", "=E17-E20"], "Derived", "2024 為負（稅前低於淨利）"),
        (20, "h_ni", "歸屬母公司淨利", [555.62, 3323, 4045], "Verified", S_SA),
        (21, "h_eps", "EPS（元）", [0.47, 2.77, 3.36], "Verified", S_SA),
        (22, "h_sh", "股數（＝淨利 ÷ EPS，百萬股）", ["=C20/C21", "=D20/D21", "=E20/E21"], "Derived", ""),
        (23, "h_div", "現金股利支付", [568.31, 656.89, 978.54], "Verified", S_SA),
        (24, "h_da", "折舊攤銷", [1091, 1232, 835.23], "Verified", S_SA),
        (25, "h_cx", "資本支出", [1048, 1249, 2336], "Verified", S_SA),
        (26, "h_ocf", "營業現金流", [7537, 4772, 4120], "Verified", S_SA),
    ]
    for (rr, k, lab, vals, tag, src) in HROWS:
        HI.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = HI.cell(rr, 3 + j, v)
                if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
                if k == "h_gm": c.number_format = "0.0%"
        HI.cell(rr, 6, tag); HI.cell(rr, 7, src)
        dn(wb, k, f"Hist!$C${rr}:$E${rr}")
    HI.column_dimensions["A"].width = 62; HI.column_dimensions["G"].width = 90

    # ================= Q（季度；一列一個數字，單一標記）=================
    Q["A1"] = "Q：2026 年季度資料（百萬元）。資料商轉錄（S&P Global）；第 3 季與前三季合計與三季報摘要（營收 221.49 億、淨利 22.99 億）一致。"
    hdr(Q, 3, ["項目", "2026Q1", "2026Q2", "2026Q3", "2026 年 1–9 月", "標記", "來源", "來源誘因", "說明"])
    QROWS = [
        (4, "營業收入", [6414, 7611, 8124, "=SUM(B4:D4)"], "Verified", S_SA + "；" + S_Q3, "", "前三季 +1.68%"),
        (5, "毛利", [1884, 2023, 2566, "=SUM(B5:D5)"], "Verified", S_SA, "", ""),
        (6, "營業利益", [834.72, 865.16, 1401, "=SUM(B6:D6)"], "Verified", S_SA, "", ""),
        (7, "營業費用及其他合計（＝毛利 − 營業利益）", ["=B5-B6", "=C5-C6", "=D5-D6", "=E5-E6"], "Derived", "＝毛利 − 營業利益", "", ""),
        (8, "歸屬母公司淨利", [502.95, 716.81, 1079, "=SUM(B8:D8)"], "Verified", S_SA + "；" + S_Q3, "", "前三季 −28.4%（2025 同期含較多投資收益）"),
        (9, "毛利率", ["=B5/B4", "=C5/C4", "=D5/D4", "=E5/E4"], "Derived", "＝毛利 ÷ 營收", "", "前三季約 29.2%，與摘要一致"),
    ]
    for (rr, lab, vals, tag, src, inc, note) in QROWS:
        Q.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            c = Q.cell(rr, 2 + j, v)
            if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
            if rr == 9: c.number_format = "0.0%"
        Q.cell(rr, 6, tag); Q.cell(rr, 7, src); Q.cell(rr, 8, inc); Q.cell(rr, 9, note)
    Q.column_dimensions["A"].width = 40; Q.column_dimensions["G"].width = 90

    # ================= 線取數 =================
    LN["A1"] = "線取數：本模型讀入的線層級變數（建置時由線模型成品依鍵或儲存格讀值）。線 6 的相機項目不在其「下游取數」頁，故直接讀「輸入」「計算」頁儲存格（來源欄寫明）。情境對應：A＝線模型高、B＝基準、C＝低。"
    hdr(LN, 3, ["項目", "情境", ""] + [f"FY{y}" for y in range(2025, 2033)] + ["單位", "來源（檔案與儲存格）"])
    LR = {}
    r = 4
    for s in SC:
        LN.cell(r, 1, "全球人形機器人出貨（線 4）"); LN.cell(r, 2, f"{s}（{LINE_SC[s]}）")
        srcs = []
        for y in range(2026, 2033):
            v, src = LD["hum"][(y, s)]
            c = LN.cell(r, 4 + (y - 2025), v); c.font = BLUE; c.number_format = "#,##0"; srcs.append(src)
        LN.cell(r, 12, "台"); LN.cell(r, 13, f"{srcs[0]} … {srcs[-1]}")
        LR[("hum", s)] = r; r += 1
    r += 1
    LN.cell(r, 1, "線 6 輕型車相機市場 2026（全部輕型車；既有需求，不計入實體 AI 合計；百萬美元）").font = BOLD; r += 1
    LR["veh"] = r
    for nm, lab in (("low", "低"), ("base", "基準"), ("high", "高")):
        v, src = LD["veh_mkt_" + nm]
        LN.cell(r, 1, "輕型車相機市場 2026"); LN.cell(r, 2, lab); c = LN.cell(r, 4, v); c.font = BLUE; c.number_format = "#,##0"
        LN.cell(r, 12, "百萬美元"); LN.cell(r, 13, src); r += 1
    LN.cell(r, 1, "機器人與 Robotaxi 相機市場 2030（基準；百萬美元）"); v, src = LD["rob_mkt30"]
    c = LN.cell(r, 4, v); c.font = BLUE; c.number_format = "#,##0.0"; LN.cell(r, 12, "百萬美元"); LN.cell(r, 13, src); LR["rob"] = r
    LN.column_dimensions["A"].width = 60; LN.column_dimensions["M"].width = 110

    # ================= Thesis（命題模組）=================
    TH["A1"] = ("命題模組：設計業務營收＝汽車市場＋新興市場（其他新興＋人形機器人相機子列）＋智慧手機及其他（殘差）；半導體分銷與其他業務另列。"
                "FY2026＝1–9 月實績＋第 4 季預告；人形相機營收＝線 4 出貨 × 倍數 × 線 6 每台顆數 × 模組單價（逐年降）× 感測器占比 × 本公司份額 × 匯率。百萬元。")
    TH["A2"] = "欄：D＝FY2025 … K＝FY2032。2026 年分市場營收未揭露，為 Assumed 年增；智慧手機及其他為殘差，其 2026 隱含年增列於下方供檢查。"
    hdr(TH, 3, ["項目", "說明", ""] + [f"FY{y}" for y in range(2025, 2033)])
    THR = {}
    r = 4
    for s in SC:
        TH.cell(r, 1, f"情境 {s}").font = BOLD; TH.cell(r, 1).fill = SECF; r += 1
        R = {}
        R["rev26"] = r; TH.cell(r, 1, "FY2026 合併營收（錨）"); TH.cell(r, 2, "＝1–9 月實績＋第 4 季預告")
        TH[f"{TC[2026]}{r}"] = "=r9m26+q4rev"; r += 1
        R["ship"] = r; TH.cell(r, 1, "全球人形出貨（線 4 × 倍數，千台）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR[('hum', s)]}*hu_k/1000"
        r += 1
        R["hum"] = r; TH.cell(r, 1, "人形機器人相機 營收（新興市場子列）"); TH.cell(r, 2, "＝出貨（千台）× 顆數 × 模組單價 ×（1 − 年降）^(年 − 2026) × 感測器占比 × 份額 × 匯率 ÷ 1000")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['ship']}*cam_n*cam_p*(1-pdc)^({y}-2026)*cis_val*ov_sh*fx/1000"
        r += 1
        R["auto"] = r; TH.cell(r, 1, "設計：汽車市場 營收"); TH.cell(r, 2, "FY2026＝FY2025 ×（1＋ga26）；FY2027 起 ×（1＋gauto）")
        for y in range(2025, 2033):
            TH[f"{TC[y]}{r}"] = "=Hist!E12" if y == 2025 else (f"={TC[2025]}{r}*(1+ga26)" if y == 2026 else f"={TC[y - 1]}{r}*(1+gauto_{s})")
        r += 1
        R["emo"] = r; TH.cell(r, 1, "設計：新興市場（不含人形相機）營收"); TH.cell(r, 2, "FY2026＝FY2025 ×（1＋ge26）− 人形；FY2027 起 ×（1＋gem）")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = (f"=Hist!E13*(1+ge26)-{TC[2026]}{R['hum']}" if y == 2026 else f"={TC[y - 1]}{r}*(1+gem_{s})")
        r += 1
        R["emer"] = r; TH.cell(r, 1, "設計：新興市場 營收（含人形相機）")
        TH[f"{TC[2025]}{r}"] = "=Hist!E13"
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['emo']}+{TC[y]}{R['hum']}"
        r += 1
        R["dist"] = r; TH.cell(r, 1, "半導體分銷 營收"); TH.cell(r, 2, "FY2026＝營收 × 分銷占比；FY2027 起 ×（1＋gdist）")
        for y in range(2025, 2033):
            TH[f"{TC[y]}{r}"] = "=Hist!E9" if y == 2025 else (f"={TC[2026]}{R['rev26']}*dist_sh26" if y == 2026 else f"={TC[y - 1]}{r}*(1+gdist_{s})")
        r += 1
        R["oth"] = r; TH.cell(r, 1, "其他業務 營收")
        for y in range(2025, 2033):
            TH[f"{TC[y]}{r}"] = "=Hist!E10" if y == 2025 else (f"={TC[2026]}{R['rev26']}*oth_sh26" if y == 2026 else f"={TC[y - 1]}{r}")
        r += 1
        R["dsg"] = r; TH.cell(r, 1, "半導體設計 營收"); TH.cell(r, 2, "FY2026＝營收 − 分銷 − 其他；FY2027 起＝三個市場列合計")
        R["phone"] = r + 1
        TH[f"{TC[2025]}{r}"] = "=Hist!E8"
        TH[f"{TC[2026]}{r}"] = f"={TC[2026]}{R['rev26']}-{TC[2026]}{R['dist']}-{TC[2026]}{R['oth']}"
        for y in range(2027, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['auto']}+{TC[y]}{R['emer']}+{TC[y]}{R['phone']}"
        r += 1
        TH.cell(r, 1, "設計：智慧手機及其他（殘差；含類比與顯示）"); TH.cell(r, 2, "FY2026＝設計 − 汽車 − 新興（殘差）；FY2027 起 ×（1＋gph）")
        TH[f"{TC[2025]}{r}"] = "=Hist!E14"
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = (f"={TC[y]}{R['dsg']}-{TC[y]}{R['auto']}-{TC[y]}{R['emer']}" if y == 2026 else f"={TC[y - 1]}{r}*(1+gph_{s})")
        r += 1
        R["phg"] = r; TH.cell(r, 1, "　智慧手機及其他 年增（檢查：2026 為隱含值）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['phone']}/{TC[y - 1]}{R['phone']}-1"; TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 1
        R["humsh"] = r; TH.cell(r, 1, "人形相機營收占合併營收")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['hum']}/({TC[y]}{R['dsg']}+{TC[y]}{R['dist']}+{TC[y]}{R['oth']})"; TH[f"{TC[y]}{r}"].number_format = "0.000%"
        r += 1
        R["gmd"] = r; TH.cell(r, 1, "設計業務毛利率路徑"); TH.cell(r, 2, "FY2026＝(整體 − 分銷占比 × 分銷毛利)÷(1 − 分銷占比)；線性到 FY2030 gmd30，其後持平")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = ("=(gm26-dist_sh26*gmdist)/(1-dist_sh26)" if y == 2026 else
                                 f"=${TC[2026]}${r}+(gmd30_{s}-${TC[2026]}${r})*MIN(1,({y}-2026)/4)")
            TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 1
        R["gm"] = r; TH.cell(r, 1, "整體毛利率路徑"); TH.cell(r, 2, "＝設計毛利率 × 設計與其他占比 ＋ 分銷毛利率 × 分銷占比（混合）")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = f"=(1-{c}{R['dist']}/({c}{R['dsg']}+{c}{R['dist']}+{c}{R['oth']}))*{c}{R['gmd']}+{c}{R['dist']}/({c}{R['dsg']}+{c}{R['dist']}+{c}{R['oth']})*gmdist"
            TH[f"{c}{r}"].number_format = "0.0%"
        r += 1
        R["opex"] = r; TH.cell(r, 1, "營業費用及其他合計率路徑"); TH.cell(r, 2, "FY2026＝opex26；線性到 FY2030 opex30，其後持平")
        for y in range(2026, 2033):
            TH[f"{TC[y]}{r}"] = "=opex26" if y == 2026 else f"=opex26+(opex30-opex26)*MIN(1,({y}-2026)/4)"; TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 2
        THR[s] = R
    for row in TH.iter_rows(min_row=4, max_row=r, min_col=4, max_col=11):
        for c in row:
            if c.number_format == "General": c.number_format = "#,##0.0"
    TH.column_dimensions["A"].width = 44; TH.column_dimensions["B"].width = 50

    # ================= PL（標準損益表，三情境）=================
    PL["A1"] = ("PL：官方分部（半導體設計／半導體分銷／其他業務）→ 合併營收 → 標準損益表 → FCFF。"
                "C＝FY2023 … E＝FY2025 為實績；F＝FY2026（1–9 月實績＋第 4 季預告）… L＝FY2032。百萬元；中國會計準則。FY2023 分部未找到。")
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
        T_ = lambda k: (lambda y: f"=Thesis!{TC[y]}{th[k]}")
        row("auto", "設計：汽車市場（子列）", "百萬元", H24(12), T_("auto"))
        row("emer", "設計：新興市場（子列，含人形相機）", "百萬元", H24(13), T_("emer"))
        row("hum", "　其中：人形機器人相機（備忘，不加總）", "百萬元", None, T_("hum"))
        row("phone", "設計：智慧手機及其他（子列，殘差）", "百萬元", H24(14), lambda y: f"=Thesis!{TC[y]}{th['phone']}")
        row("dsg", "半導體設計 營收", "百萬元", H24(8), lambda y: f"={C(y)}{R['auto']}+{C(y)}{R['emer']}+{C(y)}{R['phone']}")
        row("agt", "半導體分銷 營收", "百萬元", H24(9), T_("dist"))
        row("oth", "其他業務 營收", "百萬元", lambda y: f"=Hist!E10" if y == 2025 else None, T_("oth"))
        row("rev", "營收（合併）", "百萬元", H(4), lambda y: f"={C(y)}{R['dsg']}+{C(y)}{R['agt']}+{C(y)}{R['oth']}")
        row("chk", "檢查：分部合計 − 合併營收（應為 0）", "百萬元",
            lambda y: None if y == 2023 else f"={C(y)}{R['dsg']}+{C(y)}{R['agt']}+{C(y)}{R['oth']}-{C(y)}{R['rev']}",
            lambda y: f"={C(y)}{R['dsg']}+{C(y)}{R['agt']}+{C(y)}{R['oth']}-{C(y)}{R['rev']}")
        PL.cell(R["chk"], 1, "檢查：分部合計 − 合併營收（應為 0）" if s == "B" else f"檢查（情境 {s}）：分部合計 − 合併營收（應為 0）")
        row("chk2", "檢查：市場子列合計 − 設計分部（應為 0）", "百萬元", None, lambda y: f"={C(y)}{R['auto']}+{C(y)}{R['emer']}+{C(y)}{R['phone']}-{C(y)}{R['dsg']}")
        row("yoy", "營收年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1",
            lambda y: f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1", "0.0%")
        row("cogs", "營業成本", "百萬元", H(5), lambda y: f"={C(y)}{R['rev']}*(1-Thesis!{TC[y]}{th['gm']})")
        row("gp", "毛利", "百萬元", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}")
        row("gm", "毛利率", "%", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", "0.0%")
        row("opex", "營業費用及其他營業損益合計", "百萬元", H(15), lambda y: f"={C(y)}{R['rev']}*Thesis!{TC[y]}{th['opex']}")
        row("op", "營業利益", "百萬元", H(16), lambda y: f"={C(y)}{R['gp']}-{C(y)}{R['opex']}")
        row("opm", "營業利益率", "%", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", "0.0%")
        row("nonop", "業外淨額（稅前 − 營業利益；歷史為近似）", "百萬元", H(18), lambda y: f"={C(y)}{R['rev']}*nonop_r")
        row("pbt", "稅前利益", "百萬元", H(17), lambda y: f"={C(y)}{R['op']}+{C(y)}{R['nonop']}")
        row("tax", "所得稅", "百萬元", H(19), lambda y: f"={C(y)}{R['pbt']}*tax_r")
        row("ni", "歸屬母公司淨利（少數股東損益未揭露，併入）", "百萬元", H(20), lambda y: f"={C(y)}{R['pbt']}-{C(y)}{R['tax']}")
        row("nim", "淨利率", "%", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['rev']}", "0.0%")
        row("niy", "淨利年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1",
            lambda y: f"={C(y)}{R['ni']}/{COL[y - 1]}{R['ni']}-1", "0.0%")
        row("sh", "股數（百萬股）", "百萬股", H(22), lambda y: "=sh_v/(1-h_iss)")
        row("eps", "EPS（元）", "元", H(21), lambda y: f"={C(y)}{R['ni']}/{C(y)}{R['sh']}", "0.00")
        row("da", "折舊攤銷", "百萬元", H(24), lambda y: f"={C(y)}{R['rev']}*(da26+(cx_r*da_ss-da26)*{y - 2026}/6)")
        row("cx", "資本支出", "百萬元", H(25), lambda y: "=cx26" if y == 2026 else f"={C(y)}{R['rev']}*cx_r")
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
    CP["A1"] = ("Comp：同業前瞻 P/E（stockanalysis.com statistics，資料商）。股價＝" + S_CLOSE + "（Verified）；前瞻 P/E＝原值 × 2026-11-20 收盤 ÷ 原報價（Derived，EPS 不變）。"
                "核心同業僅 4 家（收盤檔內有報價且資料商有前瞻 P/E 者），不做迴歸；收盤檔沒有的同業不進中位數，列「報價日未取得」。")
    hdr(CP, 4, ["公司", "代號", "股價（2026-11-20）", "報價日", "前瞻 P/E 分母", "前瞻 P/E", "近 12 月 P/E（原報價日）", "5 年 β", "會計口徑", "進中位數", "說明", "原報價", "原報價日", "原前瞻 P/E", "標記與換算"])
    PEERS = [  # (公司, 代號, 11-20 收盤, 原報價, 原報價日, 原前瞻 P/E, 近 12 月 P/E, β, 口徑, 進中位數, 說明, 分母)
        ("Sony", "6758.T", 4312.0, 4322.0, "2026-11-24", 18.07, None, 0.84, "日本基準／IFRS（未查核）", "是", "CIS 全球龍頭（影像感測器為其一部門）；近 12 月每股虧損", "資料商 NTM 共識 EPS"),
        ("舜宇光學", "2382.HK", 58.5, 63.0, "2026-11-24", 13.16, 12.35, 1.22, "IFRS（未查核）", "是", "相機模組（下游客戶端）", "資料商 NTM 共識 EPS"),
        ("Ambarella", "AMBA", 53.07, 55.56, "2026-11-23", 62.57, None, 2.36, "美國 GAAP", "是", "車用與邊緣視覺晶片；近 12 月虧損，前瞻 P/E 偏高", "資料商 NTM 共識 EPS"),
        ("Qualcomm", "QCOM", 149.94, 165.38, "2026-11-23", 17.18, 13.02, 1.76, "美國 GAAP", "是", "手機與車用晶片（相鄰同業，非 CIS）", "資料商 NTM 共識 EPS"),
        ("奧比中光", "688322.SS", 0, 0, "未取得", None, None, None, "中國會計準則", "否", "深度相機；收盤檔有報價但本輪未查前瞻 P/E（成本控制）", "—"),
        ("Samsung Electronics", "005930.KS", 0, 0, "報價日未取得", None, None, None, "IFRS", "否", "手機 CIS；收盤檔無此代號，報價日未取得，不進中位數", "—"),
        ("豪威（本公司，對照）", "603501.SS", 83.91, 86.68, "2026-11-24", 22.93, None, 0.86, "中國會計準則", "否", "本公司；只作對照", "資料商 NTM 共識 EPS"),
    ]
    for i, (nm, tk, px1, px0, d0, pe0, ttm, bt, acc, med, note, den) in enumerate(PEERS):
        rr = 5 + i
        CP.cell(rr, 1, nm); CP.cell(rr, 2, tk)
        if px1: CP.cell(rr, 3, px1).font = BLUE
        CP.cell(rr, 4, "2026-11-20" if px1 else "報價日未取得"); CP.cell(rr, 5, den)
        if pe0 is not None: CP.cell(rr, 6, f"=N{rr}*C{rr}/L{rr}")
        else: CP.cell(rr, 6, "未取得")
        if ttm is not None: CP.cell(rr, 7, ttm).font = BLUE
        if bt is not None: CP.cell(rr, 8, bt).font = BLUE
        CP.cell(rr, 9, acc); CP.cell(rr, 10, med); CP.cell(rr, 11, note)
        if px0: CP.cell(rr, 12, px0).font = BLUE
        CP.cell(rr, 13, d0)
        if pe0 is not None: CP.cell(rr, 14, pe0).font = BLUE
        CP.cell(rr, 15, ("股價 Verified（" + S_CLOSE + "）；前瞻 P/E Derived：＝原值 × 2026-11-20 收盤 ÷ 原報價（分母 EPS 不變）；原值 " + S_PEER) if px1 else "收盤檔無報價或未查，不進中位數")
    CP.cell(15, 1, "核心同業前瞻 P/E 中位數"); CP["F15"] = "=MEDIAN(F5:F8)"
    CP.cell(16, 1, "本公司現價 ÷ 資料商前瞻 EPS（換算至 2026-11-20；分母：資料商 NTM 共識）"); CP["F16"] = "=F11"
    CP.cell(17, 1, "本公司現價 ÷ 模型 FY2026 EPS（分母：情境 B，FY2026）"); CP["F17"] = f"=px/PL!F{PR['B']['eps']}"
    CP.cell(18, 1, "本公司現價 ÷ 模型 FY2027 EPS（分母：情境 B，FY2027）"); CP["F18"] = f"=px/PL!G{PR['B']['eps']}"
    CP.cell(19, 1, "本公司現行倍數 ÷ 同業中位數（資料商 NTM 口徑）"); CP["F19"] = "=F16/F15"
    for rr in range(5, 20):
        if CP.cell(rr, 6).value not in (None, "未取得"): CP.cell(rr, 6).number_format = "0.0"
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
    row1("cons_show", "賣方平均目標價（資料商 18 家，頁面報價 2026-11-24；只作對照）", "=cons_tp")
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
        SE.cell(r + 1, 1, "說明：折舊、資本支出（FY2027 起）、營運資金隨營收等比例變動（第 44 條）；人形相機單價年降幅只作用於人形子列。")
    SE.column_dimensions["A"].width = 52; SE.column_dimensions["B"].width = 44

    # ================= Recon =================
    RC["A1"] = "Recon：模型對實績、公司預告、線 6 市場規模的同口徑對照；反解支持現價所需的主導變數。"
    hdr(RC, 3, ["項目", "模型", "對照", "差距（模型 − 對照）", "說明"])
    rr = 4
    def rrow(lab, a, b, note, fmt="#,##0.0", diff=True):
        nonlocal rr
        RC.cell(rr, 1, lab); RC.cell(rr, 2, a).number_format = fmt
        if b is not None: RC.cell(rr, 3, b).number_format = fmt
        if diff and b is not None: RC.cell(rr, 4, f"=B{rr}-C{rr}").number_format = fmt
        RC.cell(rr, 5, short(note)); rr += 1
    B_ = PR["B"]
    rrow("FY2026 營收 vs 公司第 4 季預告區間下限推得（1–9 月＋7,759）", f"=PL!F{B_['rev']}", "=r9m26+7759", "預告區間 7,759–8,324；基準取中點")
    rrow("FY2026 營收 vs 預告區間上限推得（1–9 月＋8,324）", f"=PL!F{B_['rev']}", "=r9m26+8324", "")
    rrow("FY2026 毛利率 vs 前三季實績", f"=PL!F{B_['gm']}", "=Q!E9", "第 4 季預告毛利率 29.44–30.74%", "0.0%")
    rrow("FY2026 歸屬母公司淨利 − 前三季淨利（＝隱含第 4 季）vs 第 3 季淨利", f"=PL!F{B_['ni']}-Q!E8", "=Q!D8", "第 4 季淨利公司未預告；模型的業外與稅率為 Assumed")
    rrow("FY2026 營業利益 − 前三季營業利益（＝隱含第 4 季）vs 第 3 季營業利益", f"=PL!F{B_['op']}-Q!E6", "=Q!D6", "")
    rrow("FY2026 智慧手機及其他隱含年增（殘差）", f"=Thesis!E{THR['B']['phg']}", None, "由設計業務前三季 −9.4%、汽車與新興 Assumed 年增反推；若偏離，表示分市場年增假設需修正", "0.0%", False)
    rrow("加權目標價 vs 賣方平均目標價（資料商 18 家）", f"=Val!B{VR['tp']}", "=cons_tp", S_SA_ST + "；只作對照，不進模型")
    rr += 1
    RC.cell(rr, 1, "線 4／線 6 對照（人形相機曝險規模）").font = BOLD; rr += 1
    rrow("2025 汽車市場營收（百萬美元）÷ 線 6 輕型車相機市場 2026 基準", f"=Hist!E12/fx/線取數!D{LR['veh'] + 1}", None, f"線 6 基準 {LD['veh_mkt_base'][0]:,.0f} 百萬美元（{LD['veh_mkt_base'][1]}，低 {LD['veh_mkt_low'][0]:,.0f}／高 {LD['veh_mkt_high'][0]:,.0f}）；本公司汽車營收為感測器、線 6 為整個相機價值，故比例＝感測器占比與份額的乘積", "0.0%", False)
    rrow("FY2026 汽車市場營收（情境 B）÷ 線 6 輕型車相機市場 2026 基準", f"=Thesis!E{THR['B']['auto']}/fx/線取數!D{LR['veh'] + 1}", None, "同上", "0.0%", False)
    rrow("FY2030 人形相機營收（情境 B，百萬美元）vs 線 6 機器人與 Robotaxi 相機市場 2030 基準", f"=Thesis!I{THR['B']['hum']}/fx", f"=線取數!D{LR['rob']}", f"線 6：{LD['rob_mkt30'][1]}；機器人與 Robotaxi 全部相機（含 Robotaxi），本公司人形子列為其一小部分")
    rrow("FY2030 人形相機營收占合併營收（情境 B）", f"=Thesis!I{THR['B']['humsh']}", None, "人形對本公司營收的貢獻在證據範圍內可忽略", "0.000%", False)
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
    h26 = LD["hum"][(2026, "B")]; h30 = LD["hum"][(2030, "B")]
    ZR = [
        ("全球人形機器人出貨 FY2026–FY2032（A／B／C＝線高／基準／低）", "4 人形", f"基準 {h26[0]:,.0f} → {h30[0]:,.0f}（2030）（× hu_k）", "台", "全球、含小型機", f"{h26[1]}；{h30[1]}（逐年見線取數頁）", "Assumed", "Thesis 人形相機營收", "一致（直接讀值，hu_k＝1）"),
        ("人形每台相機顆數", "6 共用零組件", "=cam_n", "顆／台", "人形頭部 2–8 顆", f"{LD['cam_n'][1][1]}（低 {LD['cam_n'][0][1]}、高 {LD['cam_n'][2][1]}）", "Analogy", "Thesis 人形相機營收", "一致（線 6 輸入頁儲存格，非「下游取數」頁；該頁無相機項目）"),
        ("相機模組單價 2026／年降幅", "6 共用零組件", "=cam_p", "美元／顆", "模組單價", f"{LD['cam_p'][1][1]}；年降 {LD['cam_pd'][1][1]}", "Assumed", "Thesis 人形相機營收", "一致（線 6 標記該單價來源未取得）"),
        ("輕型車平均每車相機顆數", "6 共用零組件", f"{LD['veh_n'][1][0]:g}（{LD['veh_n'][0][0]:g}–{LD['veh_n'][2][0]:g}）", "顆／車", "全部輕型車，既有需求", f"{LD['veh_n'][1][1]}", "Assumed", "gauto 的依據（定性）", "一致；本模型汽車營收以本公司實績年增驅動，不用顆數乘單價"),
        ("輕型車相機市場 2026（低／基準／高）", "6 共用零組件", f"{LD['veh_mkt_low'][0]:,.0f}／{LD['veh_mkt_base'][0]:,.0f}／{LD['veh_mkt_high'][0]:,.0f}", "百萬美元", "全部輕型車；線 6 視為既有需求，不計入實體 AI 合計", f"{LD['veh_mkt_base'][1]}", "Derived", "Recon 對照（不進公式）", "一致（只作對照）"),
        ("機器人與 Robotaxi 相機市場 2030（基準）", "6 共用零組件", f"{LD['rob_mkt30'][0]:,.1f}", "百萬美元", "機器人與 Robotaxi", f"{LD['rob_mkt30'][1]}", "Derived", "Recon 對照（不進公式）", "一致（只作對照）"),
        ("感測器占相機模組價值、本公司人形份額", "—", "cis_val 30%、ov_sh 25%", "%", "本模型自設", "In 頁", "Assumed", "Thesis 人形相機營收", "不一致（線模型未提供）：自設值，區間寬；整體結論對其不敏感（見 Sens）"),
    ]
    for i, z in enumerate(ZR):
        for j, v in enumerate(z): DZ.cell(4 + i, 1 + j, v)
    DZ.column_dimensions["A"].width = 50; DZ.column_dimensions["F"].width = 80; DZ.column_dimensions["I"].width = 70
    # ================= Risk =================
    RK["A1"] = "Risk：風險登錄（原則第 10 節第 10、19 條）。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    RISKS = [
        ("半導體分銷（占前三季 25.7%）持續性與毛利率", "是（dist_sh26、gdist、gmdist）", "公司稱受 AI 產業鏈需求拉動；分銷毛利率未找到", "季度分銷占比、整體毛利率"),
        ("智慧手機 CIS 需求與份額", "是（殘差與 gph）", "設計業務前三季 −9.4%", "季度設計業務營收、國內手機品牌拉貨"),
        ("車用 CIS 回升力道與份額", "是（ga26、gauto）", "管理層稱持續回升，未揭露金額", "季度汽車市場營收、車廠規格"),
        ("人形與機器視覺放量", "是（子列，貢獻可忽略）", "全域快門 Q4 部分料號量產；人形占比極低", "年報新興市場營收、線 4 出貨"),
        ("長期投資 6,535 百萬元的價值與變現", "部分（inv_k 0–100%）", "明細未取得", "年報附註"),
        ("有息負債 13,444 百萬元增加（一年內增加 3,756）", "是（淨現金）", "與現金同步增加，用途未查", "季度負債與資本支出"),
        ("地緣政治與出口管制", "否", "無法量化", "實體清單、客戶出貨"),
        ("單一客戶集中", "否", "客戶集中度未找到", "年報前五大客戶"),
    ]
    for i, k in enumerate(RISKS):
        for j, v in enumerate(k): RK.cell(4 + i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["D"].width = 60
    # ================= Sources =================
    SO["A1"] = "Sources：本模型引用的原始來源與資料商（含網址；HTML 展示層不放網址）。"
    for i, s_ in enumerate([S_PX, S_SA, S_SA_ST, S_SEG, S_Q3, S_IR, S_RF, S_CRP, S_PEER,
                            f"線 4：{L4N}", f"線 6：{L6N}", "physical_ai/map/profiles/603501.md（背景，數字已另行查證）"], 3):
        SO.cell(i, 1, s_)
    SO.column_dimensions["A"].width = 160
    # ================= README（xlsx 內）=================
    S = T
    v = lambda k: f"Val!B{VR[k]}"
    lines = [
        ('="豪威集團（603501.SS）個股模型 ' + VER + '｜build_id "&Sens!D2&"｜資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'),
        "命題：豪威的價值來自設計業務（車用、手機、新興市場影像感測器）的毛利與規模；半導體分銷營收大、毛利低；人形機器人相機在證據範圍內只是微小子列。",
        (f'="結論（條件式）：在 FY2026＝前三季實績＋第 4 季預告中點、設計業務毛利率 FY2030 "&{S("gmd30_B","0%")}&"（情境 B）、同業前瞻 P/E 中位數 "&{S("pe_peer","0.0")}&" 倍且未加溢價的假設下，加權目標價約 "&{S(v("tp"),"#,##0.0")}&" 元，相對現價 "&{S("px","#,##0.00")}&" 元 "&{S(v("up"),"+0.0%;-0.0%")}&"，估值評等 "&{v("rating")}&"。"'),
        (f'="DCF 目標價 "&{S(v("dcf_tp"),"#,##0.0")}&"、Comparable 目標價 "&{S(v("cmp_tp"),"#,##0.0")}&"（各半）；WACC "&{S("wacc","0.00%")}&"、ke "&{S("ke","0.00%")}&"、g "&{S("g_tv","0.00%")}&"；情境 B 終值占 EV "&{S(f"Val!C{VR['tvs']}","0%")}&"。"'),
        (f'="上檔空間歸因：DCF 貢獻 "&{S(v("att_d"),"+0.0%;-0.0%")}&"、未調整 Comparable 貢獻 "&{S(v("att_c"),"+0.0%;-0.0%")}&"、相對倍數調整貢獻 "&{S(v("att_r"),"+0.0%;-0.0%")}&"。本公司現行前瞻倍數是同業中位數的 "&{S(v("relnow"),"0.0")}&" 倍；支持現價需要相對倍數係數 "&{S(v("relreq"),"0.0")}&"。"'),
        (f'="收支關卡：情境 B FY2026–FY2030 FCFF 扣股利合計 "&{S(v("gate_b"),"#,##0.0")}&" 百萬元，"&{v("gate_t")}&"；距最近級距邊界 "&{S(v("dist"),"0.0%")}&"。"'),
        "主要失效條件（可觀察）：(1) 分銷毛利率與占比；(2) 設計業務毛利率能否回到 40%；(3) 同業中位數僅 4 家，換同業組合結論可能改變；(4) 長期投資與負債的實際內容。",
        "方法：DCF（FCFF，WACC 折現，明確期到 FY2032，RONIC＝WACC；金融資產以淨現金計入）50%＋同業前瞻 P/E × 模型 FY2027 EPS 50%；三情境 25／50／25。",
        "限制：分市場營收 2026 年未揭露（Assumed）；歷史數字取自資料商與新聞轉載（年報原文未直接讀取）；稅前與稅額為近似；同業 P/E 由原報價日按 2026-11-20 收盤等比例換算；不讀 Siliconomics／Tokenomics／Wattonomics。詳見 repo README.md。",
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
            subprocess.run(["pkill", "-f", "lo_omnivision"], check=False)
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
    for d_, lab_ in SCN_DRIVERS:
        extra.append((f"情境 B 的{lab_}改為情境 A 值", f"{d_}_B", ivE[f"{d_}_A"]))
        extra.append((f"情境 B 的{lab_}改為情境 C 值", f"{d_}_B", ivE[f"{d_}_C"]))
    extra += [("人形出貨為線 4 的 20 倍（全部情境）", "hu_k", 20.0), ("相對倍數係數＝1（基準即為 1）", "rel_k", 1.0)]
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
    def solve_dir(key, target_tp, lo, hi, n=11, over=None):
        over = over or {}
        up_ = run_cases([{**over, key: hi}], meta, "solve")[0]["tp"] >= run_cases([{**over, key: lo}], meta, "solve")[0]["tp"]
        a, b = math.log(lo), math.log(hi)
        ok = lambda t: (t >= target_tp) if up_ else (t <= target_tp)
        if not ok(run_cases([{**over, key: hi}], meta, "solve")[0]["tp"]): return None
        for _ in range(n):
            m = (a + b) / 2
            if ok(run_cases([{**over, key: math.exp(m)}], meta, "solve")[0]["tp"]): b = m
            else: a = m
        return math.exp((a + b) / 2)
    s1 = solve_dir("pe_peer", px, 5.0, 200.0)
    s2 = solve_dir("gmd30_B", px, 0.2, 0.9) if False else None
    s3 = solve_dir("hu_k", px, 1.0, 5000.0)
    nf = lambda x: x if x else -1
    solve_rows = [
        ("支持現價（目標價＝現價）所需的同業前瞻 P/E 中位數（倍）", nf(s1), ivE["pe_peer"], "其他輸入不變；−1＝200 倍內無解；基準為 Comp 頁 4 家中位數"),
        ("支持現價所需的人形出貨倍數（相對線 4）", nf(s3), 1.0, "其他輸入不變；−1＝5,000 倍內無解（人形相機營收占比極低）"),
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
    segs = [("半導體設計", "半導體設計 營收", Hs["E8"].value), ("半導體分銷（代理銷售）", "半導體分銷 營收", Hs["E9"].value),
            ("其他業務（殘差）", "其他業務 營收", Hs["E10"].value)]
    json.dump(dict(company="omnivision", ticker="603501.SS", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=round(v, 2), unit="百萬元", source=S_SEG) for n, l, v in segs],
                   sublines=[dict(parent="半導體設計", labels=["設計：汽車市場（子列）", "設計：新興市場（子列，含人形相機）", "設計：智慧手機及其他（子列，殘差）"])],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）", "檢查：市場子列合計 − 設計分部（應為 0）"],
                   note="官方分部＝年報主營業務分行業（半導體設計、半導體代理銷售）；其他業務為殘差。設計業務的市場拆分（汽車、新興、智慧手機及其他）取自年報分市場數字，2026 年起為命題拆分（Assumed）；人形相機為新興市場內的備忘子列，不加總。"),
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
    t = f"""# physical_ai/omnivision/ — 豪威集團／韋爾股份 OmniVision（603501.SS）個股模型 {VER}

<!-- 本檔由 build_omnivision.py 依成品重算值產生；數字不手寫（原則第 10 節第 51 條）。 -->

**本模型適用《高不確定性預估工作原則》**（`docs/plan/高不確定性預估工作原則.md`）：2026 年分市場營收、分銷毛利率、人形相機份額未揭露，一律給區間並標 Assumed／Analogy；結論只在證據邊界內成立。

**未完成項目：** 年報與三季報原文未直接讀取（歷史與季度數字取自資料商與新聞轉載）；分銷毛利率、客戶集中度、銷管研費用分項、長期投資明細、限制性股票未找到；同業僅 4 家有收盤檔報價與前瞻 P/E；獨立核對待外層流程安排。

- 成品：`dist/{NAME}`（build_id `{sens['build_id']}`；LibreOffice 重算，0 錯誤值）；HTML 展示層 `dist/{NAME[:-5]}.html`（`gen_html_json.py` 產生 `html.json`，再以 `tools/company_html/build_company_html.py` 產生）
- 建置：`python3 physical_ai/omnivision/build_omnivision.py`（需要 soffice；整本重算敏感度與反解；結尾自動跑 in_layout、sens_coverage --require-keycol、model_governance、rating_gate --strict、zero_rows、cached_values、segment_check）
- 日期：資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 2026-11-20（收盤 {I['px']:,.2f} 元）｜評價日 2026-12-31｜目標時點 2027-11-20
- 會計年度＝日曆年；單位百萬元人民幣；中國會計準則。上海證券交易所，H 股 2026-01 於港交所上市（股數為 A＋H 合計）。

## 結論（條件式）

在「FY2026＝前三季實績＋第 4 季公司預告中點、情境 B 設計業務毛利率 FY2030 {I['gmd30_B']:.0%}、同業前瞻 P/E 中位數 {I['pe_peer']:.1f} 倍未加溢價」的假設下：

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
| 賣方平均目標價（資料商 18 家，只作對照） | {I['cons_tp']:,.2f} |

上檔空間歸因：DCF 貢獻 {g('att_d'):+.1%}、未調整 Comparable 貢獻 {g('att_c'):+.1%}、相對倍數調整 {g('att_r'):+.1%}。本公司現行前瞻倍數是核心同業中位數的 {g('relnow'):.1f} 倍；要讓目標價＝現價，Comparable 的相對倍數係數須達 {g('relreq'):.2f}。同業只有 4 家（Sony、舜宇、Ambarella、高通），中位數取其中間值，最低與最高值見敏感度；換同業組合結論可能改變。

**讀法：** 人形相機在本模型只是備忘子列：情境 B 人形相機營收 {ser('hum', '{:,.1f}')}（FY2026→FY2032，百萬元），占合併營收不到 {TH[f'K{THB["humsh"]}'].value:.2%}。評價主要取決於設計業務毛利率路徑、分銷業務的持續性與同業倍數，與實體 AI 無關。反解：{'；'.join(f"{a}：{fn(b, '{:,.1f}')}（基準 {fn(c0, '{:,.1f}')}）" for a, b, c0, d in sol)}。

## 命題

1. **營收結構：** 官方分行業為半導體設計（2025 年 {wb['Hist']['E8'].value:,.0f} 百萬元，{wb['Hist']['E8'].value / wb['Hist']['E4'].value:.0%}）與半導體分銷（{wb['Hist']['E9'].value:,.0f}）；2026 前三季設計 −9.4%、分銷 +55.7%，分銷占比 25.7%（Verified，轉載）。
2. **設計業務拆分：** 汽車市場、新興市場取自年報分市場數字，智慧手機及其他為殘差；2026 年起分市場年增為 Assumed。2026 年隱含智慧手機及其他年增 {TH[f'E{THB["phg"]}'].value:.0%}。
3. **毛利：** 設計與分銷毛利率分開；分銷毛利率未找到（Assumed 8%，5–12%），整體毛利率以前三季實績與第 4 季預告為錨。

情境 B：營收 {ser('rev', '{:,.0f}')}；毛利率 {ser('gm', '{:.1%}')}；EPS {ser('eps', '{:.2f}')} 元（FY2026→FY2032）。

## 主要假設與可證偽指標（原則第 10 節第 20 條）

| 承重假設 | 基準（A／B／C 或區間） | 標記 | 可觀察指標 |
|---|---|---|---|
| FY2026 第 4 季營收與毛利率 | 公司預告區間中點 | Interested-party | 2026 年報 |
| 設計業務毛利率 FY2030 | {I['gmd30_A']:.0%}／{I['gmd30_B']:.0%}／{I['gmd30_C']:.0%} | Assumed | 季度毛利率、分銷占比 |
| 分銷營收年增 FY2027 起 | {I['gdist_A']:+.0%}／{I['gdist_B']:+.0%}／{I['gdist_C']:+.0%} | Assumed | 季度分銷營收 |
| 汽車市場年增 FY2027 起 | {I['gauto_A']:.0%}／{I['gauto_B']:.0%}／{I['gauto_C']:.0%} | Assumed | 年報分市場營收 |
| 人形相機（線 4 出貨 × 線 6 顆數 × 份額） | 顆數 {I['cam_n']:.0f}（2–8）、份額 {I['ov_sh']:.0%}（10–40%） | Analogy／Assumed | 全域快門量產與客戶揭露 |
| 同業前瞻 P/E | 中位數 {I['pe_peer']:.1f} 倍（4 家） | Derived | 同一估值日重抓 |

## 敏感度（整本重算，前 10 大）

| 案例 | 鍵 | 加權目標價 | 相對基準 | DCF 目標價 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | `{c['key']}` | {c['tp']:,.1f} | {c['tp'] / sens['base_tp'] - 1:+.1%} | {c['dcf']:,.1f} | {c['rating']} |" for c in top) + f"""

完整案例見 xlsx Sens 頁。升一級所需加權目標價 {fn(g('up_b'))} 元；降一級門檻 {fn(g('dn_b'))} 元。

## 限制與「分不出來／沒找到」

- **分不出來（資料性質）：** 2026 年分市場營收、人形用相機營收、分銷毛利率公司未揭露。
- **沒找到：** 年報與三季報原文；分銷毛利率；銷管研費用分項；客戶集中度；長期投資與負債明細；限制性股票。
- **可疑但未動：** 資料商 2024 年稅前利潤率（12.74%）推得稅前 3,278 低於淨利 3,323，稅額為負，稅前與稅額為近似；股數資料商兩處（12.5 億、市值 ÷ 價格 12.32 億）差 1.5%，採 12.5 億；汽車市場營收口徑可能含非感測器產品。
- **線模型：** 線 6 下游取數頁無相機項目，直接讀線 6 輸入與計算頁儲存格（見對帳頁）；線 4 人形出貨直接讀下游取數。
- **同業倍數：** 4 家同業原報價日 2026-11-23／24，前瞻 P/E 依 2026-11-20 收盤等比例換算（EPS 未更新）；Samsung 收盤檔無報價（報價日未取得）、奧比中光未查前瞻 P/E，皆未進中位數。
- **rf 逾期：** 中國 10 年期國債 1.61% 為 2026-11-05 值（跨模型一致）。
- **獨立核對：** 待外層流程安排。

## 資料來源

見 xlsx Sources 頁。

## 版本紀錄

- {VER}（2026-11-24）：首版。現價 {I['px']:,.2f}（2026-11-20 收盤）；加權目標價 {g('tp'):,.1f}；上檔空間 {g('up'):+.1%}；{g('rating')}／{g('rec')}；收支關卡 {g('gate_b'):,.1f}（{g('gate_t')}）。
"""
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write(t)


if __name__ == "__main__":
    main()

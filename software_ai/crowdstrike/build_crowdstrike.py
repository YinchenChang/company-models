#!/usr/bin/env python3
"""CrowdStrike（CRWD）個股模型 v0.1 建置腳本（AI 軟體應用 Project；原型 P：按防護範圍計價）。

用法：
  python3 build_crowdstrike.py            # 建置 base、跑敏感度（整本重算）、輸出 dist/<日期>_CrowdStrike個股模型_v0.1.xlsx
  python3 build_crowdstrike.py --quick    # 只建置與重算（不跑敏感度；Sens 頁用快取 data/sens_v01.json）
公開 repo：只含公司與市場事實，不含任何個人情境。
"""
import copy
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
DIST = os.path.join(HERE, "dist")
SCR = "/home/claude/scratch/crowdstrike"
VERSION = "v0.1.1"
BUILD_DATE = "2026-11-16"
FILE_DATE = "20261117"
YEARS = list(range(2024, 2033))                 # FY2024A..FY2032E
YC = {y: openpyxl.utils.get_column_letter(3 + i) for i, y in enumerate(YEARS)}   # C..K
FC = [2027, 2028, 2029, 2030, 2031, 2032]
SC = ["A", "B", "C"]
SCN_NAME = {"A": "情境 A（AI 代理防護放量、價格侵蝕小）", "B": "情境 B（基準）", "C": "情境 C（代理防護慢、價格侵蝕大）"}
D = dt.datetime

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
YEL = PatternFill("solid", fgColor="FFF2CC")
HDR = PatternFill("solid", fgColor="DDEBF7")

# ───────────────────────── 輸入（In 頁）─────────────────────────
Q2PR = "https://ir.crowdstrike.com/news-releases/news-release-details/crowdstrike-reports-second-quarter-fiscal-year-2027-financial（Q2 FY2027 新聞稿，2026-08-26）"
TENQ = "https://www.sec.gov/Archives/edgar/data/1535527/000153552726000031/crwd-20260731.htm（Form 10-Q 截至 2026-07-31）"
K26 = "https://www.sec.gov/Archives/edgar/data/1535527/000153552726000010/crwd-20260131.htm（Form 10-K FY2026）"
Q4PR = "https://www.sec.gov/Archives/edgar/data/1535527/000153552726000007/crwd-20260303xex991.htm（8-K Ex.99.1，2026-03-03）"
TKS = "data/tk_snapshot_v22.0.json（Tokenomics v22.0 快照；取自 software_ai/unit_economics/data/tk_snapshot.json，2026-11-16）"
SA = "https://stockanalysis.com/stocks/{}/statistics/"
VEND = "資料商（StockAnalysis）整理值，非法定文件"
SEC = "#"
# (key, 項目, 單位, 低, 基準, 高, 標記, 來源, 說明, 來源誘因)
RMH = [1.412, 1.149, 2.163, 2.356, 3.411]
RMH_MAX = max(RMH)
RMSRC = "https://stockanalysis.com/stocks/crwd/financials/ratios/ 與 panw、ftnt、zs、okta、s、chkp 同路徑（2026-11-17 擷取；StockAnalysis 年度 EV／Sales）；data/relmult_history.csv"
INP = [
    (SEC, "A. 日期、現價、股數與資產負債（評價日 2026-07-31；US$M；股數百萬股，4 比 1 拆股後）"),
    ("data_date", "資料截止日（本版最晚一筆輸入）", "日期", None, D(2026, 11, 16), None, "Derived", "本腳本建置日；最晚輸入為 Tokenomics 快照擷取日 2026-11-16", "README 首列與 HTML 表頭分列資料截止日、成品日、股價基準日", None),
    ("px_date", "股價基準日（現價日）", "日期", None, D(2026, 11, 13), None, "Verified", "StockAnalysis CRWD 歷史收盤價頁：2026-11-13 收盤", "2026-11-13 為最近一個完整交易日；第二來源未取得（只一個資料商）", None),
    ("val_date", "評價日（最近一季資產負債表日）", "日期", None, D(2026, 7, 31), None, "Verified", TENQ + " 資產負債表日", "Q3 FY2027 尚未公布（預計 2026 年 12 月初）", None),
    ("tp_date", "目標時點＝EDATE(股價基準日, 12)", "日期", None, "=EDATE(px_date,12)", None, "Derived", "原則第 10 節第 26 條", "持有期間＝現價日起 12 個月；折現期間＝評價日起（Val 頁並列）", None),
    ("px", "現價（2026-11-13 收盤；拆股後）", "US$", None, 226.61, None, "Verified", "https://stockanalysis.com/stocks/crwd/history/（2026-11-13 收盤 226.61）", "拆股後口徑；11 月 4 日 241.00、11 月 5 日 217.20，近期波動大", None),
    ("sh_dil", "稀釋股數（Q2 FY2027 稀釋加權平均；拆股後）", "百萬股", None, 1044.469, None, "Verified", Q2PR + " 損益表", "G2-4：評價日股數＝公司稀釋加權平均（含 RSU 與選擇權的庫藏股法影響）；RSU 預期沒收與選擇權庫藏股法明細未取得（未找到）；法定封面流通股 sh_out 為主來源對照；ASR 與可轉債不適用（未見）", None),
    ("sh_out", "流通股數（2026-08-20，10-Q 封面）", "百萬股", None, 1023.934842, None, "Verified", TENQ + " 封面", "對照用，不進估值", None),
    ("cash_0", "現金及約當現金（2026-07-31）", "US$M", None, 5013.847, None, "Verified", Q2PR + " 資產負債表", "公司無另列短期投資", None),
    ("cash_26", "現金及約當現金（2026-01-31）", "US$M", None, 5230.125, None, "Verified", Q4PR + " 資產負債表", "預測期現金餘額起點", None),
    ("debt_0", "長期負債（帳面，2026-07-31）", "US$M", None, 746.216, None, "Verified", Q2PR + " 資產負債表", "債務口徑（G2-5）：只計優先票據帳面；營業租賃負債不計（Q 頁營運資本列另有，不入淨債務）；融資租賃未另列（未找到）；本金與票面利率未在取得的段落內，標未找到", None),
    ("h1_fcf_co", "FY2027 上半年自由現金流（公司口徑，含 SBC 加回與利息收入）", "US$M", None, 845.87, None, "Interested-party", Q2PR + " 現金流量表", "公司自訂 FCF 口徑（營運現金流 − 資本支出 − 資本化軟體）", "公司自訂 FCF 口徑，含利息收入"),
    ("h1_sbc", "FY2027 上半年股票薪酬費用（現金流量表）", "US$M", None, 674.6, None, "Verified", TENQ + " 現金流量表", "DCF 首期以『SBC 費用化』口徑扣除上半年實績：FCF − SBC", None),
    ("sbc_nc", "SBC 中屬非現金部分的比率（SBC ÷ SBC 加雇主薪資稅）", "%", None, "=h_sbc26/h_sbct26", None, "Derived", "FY2026：現金流量表 SBC 1,096.7 ÷ 調節表 SBC 與雇主薪資稅 1,130.6", "用於 FCF（現金口徑）與 SBC 加回案", None),
    ("split", "拆股比例（2026 年 7 月 4 比 1）", "倍", None, 4, None, "Verified", Q2PR + " 附註", "歷史每股數字除以此值", None),
    (SEC, "B. 折現率、估值與目標價設定"),
    ("rf", "無風險利率（10 年美債，2026-11-13）", "%", None, 0.0515, None, "Verified", "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202611（2026-11-13：5.15%）", None, None),
    ("beta_ya", "原始 β：Yahoo Finance「Beta (5Y Monthly)」", "x", None, 1.20, None, "Interested-party", "https://finance.yahoo.com/quote/CRWD/key-statistics/（讀取日 2026-11-16，頁面未標日期）", "未標資料日；市值欄與現價不一致，只取 β", "資料商整理值"),
    ("beta_sa", "原始 β：StockAnalysis 統計頁", "x", None, 1.19, None, "Interested-party", SA.format("crwd") + "（讀取日 2026-11-16）", "頁面疑為舊快取（市值與財報日對不上 2026），β 僅作第二來源，兩來源差 0.01", "資料商整理值"),
    ("blume_a", "Blume 調整權重", "x", None, 0.67, None, "Assumed", "Blume (1971) 慣例 β_adj＝0.67×β＋0.33", None, None),
    ("beta", "β（Blume 調整後兩來源的中位數）", "x", None, "=MEDIAN(blume_a*beta_ya+1-blume_a,blume_a*beta_sa+1-blume_a)", None, "Derived", "兩來源 Blume 調整後取中位數（兩個值時為平均）", None, None),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "個股 DCF 標準規格 v1.0 第 2 節：預設 5%（4–6%）", None, None),
    ("kd_sp", "稅前債務利差（相對 rf）", "%", 0.004, 0.008, 0.012, "Assumed", "判斷：投資等級邊緣資安公司約 40–120bp；公司評等未取得", "債務占資本比重不到 1%，對 WACC 影響很小", None),
    ("kd", "稅前債務成本＝rf＋利差", "%", None, "=rf+kd_sp", None, "Derived", "個股 DCF 標準規格", None, None),
    ("tax", "有效稅率（預測期，GAAP 稅前為正時適用）", "%", 0.18, 0.21, 0.24, "Assumed", "公司 non-GAAP 稅率 21%（Q2 FY2027 指引）僅作參照；GAAP 現金稅因 NOL 與遞延稅另有差異，未建模", "未對現金稅率建橋（未找到已付所得稅 ÷ non-GAAP 稅前）", None),
    ("g_tv", "永續成長率", "%", 0.02, 0.03, 0.04, "Assumed", "個股 DCF 標準規格 v1.0 第 3 節：預設 3%（2–4%）", None, None),
    ("ronic_sp", "永續期 RONIC 相對 WACC 的差", "%", 0, 0, 0.05, "Assumed", "個股 DCF 標準規格：預設 RONIC＝WACC（新投資價值中性）", None, None),
    ("mid", "期中慣例：年度現金流距期初", "年", None, 0.5, None, "Derived", "個股 DCF 標準規格", None, None),
    ("tp_wA", "情境權重 A", "%", None, 0.25, None, "Assumed", "個股 DCF 標準規格：預設 25／50／25", None, None),
    ("tp_wB", "情境權重 B", "%", None, 0.5, None, "Assumed", "同上", None, None),
    ("tp_wC", "情境權重 C", "%", None, 0.25, None, "Assumed", "同上", None, None),
    ("w_dcf", "目標價中 DCF 的權重（其餘＝Comparable）", "%", 0.3, 0.5, 0.7, "Assumed", "原則第 4 節預設 DCF 50%＋Comparable 50%", None, None),
    ("t_tgt", "折現期間：評價日 → 目標時點", "年", None, "=(tp_date-val_date)/365", None, "Derived", "(tp_date − val_date) ÷ 365", None, None),
    ("t_hold", "持有期間：股價基準日 → 目標時點", "年", None, "=(tp_date-px_date)/365", None, "Derived", "(tp_date − px_date) ÷ 365", None, None),
    ("wacc_adj", "WACC 調整（只用於敏感度；基準 0）", "%", -0.01, 0.0, 0.01, "Assumed", "個股 DCF 標準規格第 4 節：WACC ±1pt", None, None),
    ("tv_alt", "終值方法開關（0＝NOPAT 價值驅動式；1＝FCFF×(1+g) 替代法）", "0／1", 0, 0, 1, "Assumed", "個股 DCF 標準規格第 4 節：終值改用 FCFF ×(1＋g) 的替代法", "只用於敏感度", None),
    ("sbc_ab", "DCF 是否加回股票薪酬（0＝不加回；1＝加回）", "0／1", 0, 0, 1, "Assumed", "個股 DCF 標準規格第 1 節：股票薪酬不加回（預設）；加回只作敏感度", "只用於敏感度", None),
    ("pm_panw", "同業 EV／營收（TTM）：Palo Alto Networks", "x", None, 22.88, None, "Interested-party", SA.format("panw") + "（2026-11-14 更新，收盤 2026-11-13）", "US GAAP；租賃處理未核對；與 CRWD 同為平台型資安", VEND),
    ("pm_ftnt", "同業 EV／營收（TTM）：Fortinet", "x", None, 11.32, None, "Interested-party", SA.format("ftnt") + "（2026-11-15 更新）", "US GAAP；硬體占比高，成長較低；租賃處理未核對", VEND),
    ("pm_zs", "同業 EV／營收（TTM）：Zscaler", "x", None, 8.88, None, "Interested-party", SA.format("zs") + "（2026-11-14 更新）", "US GAAP；2026-11 價格；租賃處理未核對", VEND),
    ("pm_okta", "同業 EV／營收（TTM）：Okta", "x", None, 9.67, None, "Interested-party", SA.format("okta") + "（2026-11-14 更新）", "US GAAP；身分安全；租賃處理未核對", VEND),
    ("pm_s", "同業 EV／營收（TTM）：SentinelOne", "x", None, 7.48, None, "Interested-party", SA.format("s") + "（2026-10-27 更新，收盤 2026-10-26）", "逾期（與基準日相差超過 2 個交易日）；直接競爭者；租賃處理未核對", VEND),
    ("pm_chkp", "同業 EV／營收（TTM）：Check Point", "x", None, 4.67, None, "Interested-party", SA.format("chkp") + "（2026-11-15 更新）", "會計準則（IFRS 或 US GAAP）與租賃處理未核對；低成長", VEND),
    ("pm_med", "同業 EV／營收（TTM）中位數", "x", None, "=MEDIAN(pm_panw,pm_ftnt,pm_zs,pm_okta,pm_s,pm_chkp)", None, "Derived", "六家同業中位數；少於約 8 家，不做迴歸（原則第 10 節第 7 條）", "G2-6：分子 EV（無槓桿）、分母 TTM 營收，與本公司同口徑；Comparable 乘目標日 TTM 營收；分母為各家 TTM 營收；成長率、利益率、資本強度並列未做（未取得）", None),
    ("rm_22", "CRWD EV／營收 ÷ 同業中位數：FY2022 年末（2022-01-31）", "x", None, RMH[0], None, "Derived", RMSRC, "CRWD 25.80；同業六家 18.27（中位數）；各家取最接近的會計年度末（PANW、ZS 為 7 月底、其餘 12 月或 1 月底）", VEND),
    ("rm_23", "同上：FY2023 年末（2023-01-31）", "x", None, RMH[1], None, "Derived", RMSRC, "CRWD 9.76；同業中位數 8.50", VEND),
    ("rm_24", "同上：FY2024 年末（2024-01-31）", "x", None, RMH[2], None, "Derived", RMSRC, "CRWD 21.32；同業中位數 9.86", VEND),
    ("rm_25", "同上：FY2025 年末（2025-01-31）", "x", None, RMH[3], None, "Derived", RMSRC, "CRWD 23.93；同業中位數 10.16", VEND),
    ("rm_26", "同上：FY2026 年末（2026-01-31）", "x", None, RMH[4], None, "Derived", RMSRC, "CRWD 22.31；同業中位數 6.54", VEND),
    ("rel_mult", "相對倍數係數（CRWD 倍數 ÷ 同業中位數）：最近 5 個會計年度末平均", "x", 1.0, "=AVERAGE(rm_22,rm_23,rm_24,rm_25,rm_26)", RMH_MAX, "Derived", "G3-4：用公開股價與財報自算的 5 年相對倍數平均（data/relmult_history.csv）；低值 1.0＝與同業中位數平價；高值＝5 年最高", "承重輸入；五個年末觀察值跨度大（最低與最高見 In 頁各列），且 CRWD 現價隱含倍數遠高於 5 年平均，見 Recon", None),
    ("ev_hc", "事件尾端：營收認列與 ARR 報告方式資訊請求（DOJ／SEC）→ 重編下修訂閱營收比例", "%", None, 0.0, 0.10, "Assumed", "G3 與原則第 19 條：10-Q Note 11 揭露 DOJ 與 SEC 資訊請求，無金額或結論；高值 10% 為尾端壓力（無證據支持幅度，非預測）", "只列 Sens，基準 0；尾端案不進加權目標價", None),
    ("ev_delta", "事件尾端：Delta 案一次性現金負擔（US$M）", "US$M", None, 0.0, 500, "Verified", "Delta 公開求償額逾 US$500M（itpro.com 與 TechTarget 2024-10 報導；證據開示中）；高值以求償額上限計，非預期損失，僅作尾端壓力", "只列 Sens，基準 0；求償金額為原告主張，法院結果未定（Interested-party：原告陳述）", None),
    (SEC, "C. 命題驅動：按防護範圍計價（受保護單位 × 模組數 × 每模組年費）與 AI 代理（非人類身分）；v0.2 改讀 unit_economics"),
    ("u_0", "受保護單位數（端點＋身分；FY2026 年底）", "百萬", 25, 40, 60, "Assumed", "公司不揭露；10-K 稱客戶逾 88,000（含 MSSP 終端）。校準：ARR ÷ 單位數＝每單位年費", "承重 Assumed 之一；只影響每單位數字，不影響 ARR 路徑（p_0 由 ARR 反推）", None),
    ("m_0", "每單位平均模組數（FY2026 年底）", "個", 5, 6, 7, "Assumed", "依模組採用分布推估：6 個以上 50%、7 個以上 34%、8 個以上 24%（Q4 FY2026 新聞稿）", "ARR 加權平均高於客戶加權；公司只揭露門檻占比", None),
    ("p_0", "每單位每模組年費（FY2026 年底，由 ARR 反推）", "US$", None, "=h_arr26/(u_0*m_0)", None, "Derived", "FY2026 年底 ARR ÷（單位數 × 模組數）", None, None),
    ("gu_27", "人類端受保護單位年增：FY2027", "%", 0.11, 0.15, 0.19, "Assumed", "起點取最近實績：ARR 年增 25%、淨留存 115%（10-K FY2026），拆為單位、模組、價格三項", "線性收斂到 FY2032 值", None),
    ("gu_32", "人類端受保護單位年增：FY2032", "%", 0.03, 0.06, 0.09, "Assumed", "收斂值，無直接證據", None, None),
    ("gm_27", "每單位模組數年增率：FY2027", "%", 0.05, 0.09, 0.13, "Assumed", "FY2025 至 FY2026 6 個以上模組占比 +2pt、7 個以上 +2pt（公司揭露）；起點取近年模組滲透速度（約每單位模組數 +9%）", "線性收斂到 FY2032 值；以比率表示，使每單位基期模組數與年費（由 ARR 反推）不影響 ARR 路徑", None),
    ("gm_32", "每單位模組數年增率：FY2032", "%", 0.01, 0.025, 0.04, "Assumed", "收斂值，無直接證據", None, None),
    ("ero_A", "情境 A：每模組年費年變動（AI 帶來的打包折扣、價格侵蝕）", "%", None, 0.0, None, "Assumed", "情境設定：價格持平", "A／B／C 差異反映 AI 平台業者（如 Claude Code Security）與同業打包競爭下的價格保留程度", None),
    ("ero_B", "情境 B：每模組年費年變動", "%", None, -0.01, None, "Assumed", "情境設定：年降 1%（Falcon Flex 轉換 ARR 增幅逾 40% 為對照，公司口述）", None, None),
    ("ero_C", "情境 C：每模組年費年變動", "%", None, -0.03, None, "Assumed", "情境設定：年降 3%", None, None),
    ("r32_A", "情境 A：FY2032 AI 代理（非人類身分）受保護單位 ÷ 人類單位", "x", None, 2.0, None, "Assumed", "情境設定。公司稱 AIDR 可能大於 EDR（Kurtz，2026-08-26，Interested-party）；絕對 ARR 未揭露", "FY2026 起點＝0（公司未揭露 AIDR ARR 與客戶數）", None),
    ("r32_B", "情境 B：FY2032 AI 代理單位 ÷ 人類單位", "x", None, 1.0, None, "Assumed", "情境設定", None, None),
    ("r32_C", "情境 C：FY2032 AI 代理單位 ÷ 人類單位", "x", None, 0.3, None, "Assumed", "情境設定", None, None),
    ("ma_n", "每個 AI 代理單位的模組數", "個", 1, 2, 3, "Assumed", "AIDR、Falcon Shield、身分防護等；公司稱 AIDR 為獨立模組另行計價", None, None),
    ("pf_A", "情境 A：AI 代理單位每模組年費 ÷ 人類單位每模組年費（價格保留率）", "x", None, 0.8, None, "Assumed", "情境設定", "AIDR 含 token 計價（基準額度加超量 token 包，經 Flex）；單價結構未揭露", None),
    ("pf_B", "情境 B：價格保留率", "x", None, 0.5, None, "Assumed", "情境設定", None, None),
    ("pf_C", "情境 C：價格保留率", "x", None, 0.25, None, "Assumed", "情境設定", None, None),
    ("conv", "訂閱營收 ÷ 平均 ARR（FY2026 實績）", "x", None, "=h_sub26/AVERAGE(h_arr25,h_arr26)", None, "Derived", "10-K FY2026：訂閱營收 ÷（期初 ARR＋期末 ARR）/2", "ARR 為年化合約值，營收依合約期間按比例認列", None),
    ("ps_pct", "專業服務營收 ÷ 訂閱營收", "%", 0.04, "=Q!M6/Q!M5", 0.065, "Derived", "最近四季（Q3 FY2026–Q2 FY2027）專業服務 ÷ 訂閱", "原則第 10 節第 35 條：比率以最近四季實績為起點", None),
    ("tasks_0", "AI 輔助分析任務數／受保護單位／年（FY2026 基期，已含在基期毛利中）", "次", 20, 40, 60, "Assumed", "公司未揭露 Charlotte AI 用量；基期推論成本已含在實績毛利內，模型只計增量", None, None),
    ("tasks_A", "情境 A：FY2032 任務數／受保護單位／年", "次", None, 120, None, "Assumed", "情境設定", "Charlotte AI、Agentic SOAR 等；用量與付費客戶數公司未揭露", None),
    ("tasks_B", "情境 B：FY2032 任務數／受保護單位／年", "次", None, 80, None, "Assumed", "情境設定", None, None),
    ("tasks_C", "情境 C：FY2032 任務數／受保護單位／年", "次", None, 50, None, "Assumed", "情境設定", None, None),
    ("tok_task", "每任務總 token 數（新鮮輸入＋快取輸入＋輸出；Workload 任務 3，0.5 小時級）", "tok", 6200, "=tk_fresh+tk_cache+tk_dec", 110400, "Assumed", TKS + " IF_TaskTokFresh／Cached／Dec 第 3 欄；低＝任務 2、高＝任務 4", "任務類別選擇為 Assumed（SOC 調查約 0.5 小時級）；token 數取自 Tokenomics", None),
    ("tk_fresh", "Tokenomics IF_TaskTokFresh（任務 3）", "tok", None, 6840, None, "Derived", TKS, None, None),
    ("tk_cache", "Tokenomics IF_TaskTokCached（任務 3）", "tok", None, 18360, None, "Derived", TKS, None, None),
    ("tk_dec", "Tokenomics IF_TaskTokDec（任務 3）", "tok", None, 2400, None, "Derived", TKS, None, None),
    ("px_tok", "每百萬 token 市場價：Terra 參考請求混合（IF_PriceRef_Terra）", "US$/百萬 tok", 0.0803555555555556, 0.784, 1.568, "Derived", TKS + " IF_PriceRef_Terra；低＝Luna（IF_PriceRef_Luna）、高＝Sol（IF_PriceRef_Sol；只動成本一邊，故為上界，另列聯動與模型組合案，G2-2）；選哪一層價格屬 Analogy", "軟體公司向模型商購買的市場價；不自行重估；token 價格逐年下降未建模（保守）", None),
    ("fc_tok", "對照：模型商全成本（IF_FullCostDefault_Terra，GB300 基準）", "US$/百萬 tok", None, 0.212357464257912, None, "Derived", TKS, "只作『模型商自己做這件事的成本下限』對照，不進模型", None),
    (SEC, "D. 損益與現金流驅動（GAAP；公司 non-GAAP 只作起點與對照）"),
    ("gmsub_ng", "Non-GAAP 訂閱毛利率（不含增量 AI 推論成本）", "%", 0.79, 0.81, 0.83, "Interested-party", Q2PR + "：近四季各季均 81%", "公司自訂 non-GAAP 口徑（排除 SBC、攤銷等）", "公司自訂 non-GAAP 口徑；管理層稱 FY2029 目標模型（數值未取得）"),
    ("gmps_ng", "Non-GAAP 專業服務毛利率", "%", 0.25, 0.32, 0.35, "Interested-party", Q2PR + "：近四季 30%–35%", "公司自訂 non-GAAP 口徑", "公司自訂 non-GAAP 口徑"),
    ("opx_27", "Non-GAAP 營業費用率（S&M＋R&D＋G&A）：FY2027", "%", 0.53, "=Q!M11-Q!M17", 0.57, "Derived", "最近四季：non-GAAP 毛利率 − non-GAAP 營業利益率（Q 頁）", "原則第 10 節第 35 條：以最近四季實績為起點；FY2027 指引 non-GAAP 營業利益約 25% 只進 Recon", None),
    ("fix_ng", "Non-GAAP 固定成本（US$M；規模經濟：營運成本＝固定＋變動 × 營收）", "US$M", 0, "=INTERCEPT(Hist!C45:F45,Hist!C5:F5)", 400, "Derived", "G3-1：FY2023–FY2026 四年 non-GAAP 總成本（營收 − non-GAAP 營業利益）對營收的最小平方回歸截距（Hist 頁）；非 GAAP 營業利益率由 FY2023 15.9% 升至 FY2026 21.7%，FY2024 起大致持平，故固定部分不大；FY2027 以近四季費用率為起點，之後只吸收固定成本的稀釋（營業費用率 ＝ 起點費用率 − 固定成本 ×（1／FY2027 營收 − 1／當年營收））；低值 0＝全期持平（下緣，G2-3）", "高值 400 為規模經濟加大案（Assumed）", None),
    ("sbc_27", "SBC（含雇主薪資稅）占營收：FY2027", "%", 0.22, "=Q!M12/Q!M7", 0.27, "Derived", "最近四季 SBC 與相關薪資稅 ÷ 營收（Q 頁）", "現金流量表 SBC 占營收 FY2026 22.8%（Hist 頁）", None),
    ("fix_sbc", "SBC（含雇主薪資稅）固定部分（US$M）", "US$M", 0, "=MAX(0,INTERCEPT(Hist!C22:F22,Hist!C5:F5))", 400, "Derived", "G3-1：FY2023–FY2026 SBC（現金流量表）對營收的回歸截距為負（約 −3）：歷史 SBC 占營收 23.5%、21.2%、21.8%、22.8% 無下降趨勢，固定部分取 0（即基準＝近四季占營收比持平）；低值 0；高值 400 為 SBC 規模經濟案（Assumed，無歷史證據）", "DCF 以 NOPAT 扣 SBC，不加回（個股 DCF 標準規格預設）", None),
    ("oth_pct", "Non-GAAP 與 GAAP 營業利益的其他調整（攤銷、併購費用、7 月 19 日事件等）占營收", "%", 0.003, 0.007, 0.012, "Assumed", "Q2 FY2027 調節表：攤銷 13.3、併購費用 6.7、事件回收 −14.5（US$M）；近年約 0.7% 營收", None, None),
    ("gm_gap", "GAAP 與 non-GAAP 毛利率差（成本中的 SBC 與攤銷）", "%", 0.03, 0.035, 0.045, "Derived", "FY2026：non-GAAP 毛利率 78.2% 減 GAAP 74.7%（Hist 頁）", "僅分配 GAAP 毛利與營業費用，不影響營業利益", None),
    ("ext_d", "延伸期終點（FY2037）GAAP 營業利益率相對 FY2032 的變動", "%", -0.03, 0.0, 0.08, "Assumed", "G2-3：基準持平（0）；FY2033–FY2037 營收年增線性收斂到永續成長率 g，利益率不變。高值 +8pt＝規模槓桿收斂案，低值 −3pt＝競爭壓縮，只放 Sens", "延伸期必要性：明確期到 FY2032 時營收年增仍高於 10%，直接接永續價值會低估，故加 5 年延伸期（個股 DCF 標準規格第 1 節）", None),
    ("da_k", "折舊與攤銷 ÷ 前一年資本支出（資本支出以固定比例轉為折舊攤銷）", "x", 0.75, "=AVERAGE((h_da25+h_am25)/h_cxt24,(h_da26+h_am26)/h_cxt25)", 1.0, "Derived", "FY2025、FY2026 折舊攤銷（含無形資產攤銷）÷ 前一年資本支出合計（Hist 頁）的平均；使折舊與攤銷隨資本支出累積，不再固定占營收比（資本支出 9% 對折舊 6% 長期並存為內部不一致）", "營業利益由費用率決定，折舊攤銷只影響 FCFF 的非現金加回", None),
    ("cx_27", "資本支出占營收（含資本化軟體）：FY2027", "%", 0.075, 0.09, 0.105, "Assumed", "FY2026 7.7%、FY2027 上半年 9.5%（現金流量表）；以後者附近為起點", "資料中心設備與自建資料中心上升", None),
    ("fix_cx", "資本支出固定部分（US$M；含資本化軟體）", "US$M", 0, "=MAX(0,INTERCEPT(Hist!C46:F46,Hist!C5:F5))", 300, "Derived", "G3-1：FY2023–FY2026 資本支出合計對營收的回歸截距（Hist 頁）；FY2027 以 In 頁 cx_27 占營收比為起點，之後只吸收固定部分的稀釋；低值 0＝全期持平（下緣）；高值 300 為加大案（Assumed）", "FY2027 上半年占營收 9.5% 高於 FY2024–FY2026（7.4%–7.9%），回歸截距主要由 FY2023 的高占比（11.8%）帶出", None),
    ("wc_in", "營運資金流入 ÷ 營收增量（遞延收入為主，含遞延取得成本攤銷加回）", "x", 0.25, 0.40, 0.50, "Assumed", "FY2025 0.348、FY2026 0.454（Hist 頁，10-K 現金流量表營運資產負債變動，與營運現金流勾稽）", "個股模型原則第 10 節第 23 條；遞延收入占 FY2026 營運現金流來源大宗", None),
    ("int_y", "現金利率（利息收入 ÷ 期初現金）", "%", 0.03, 0.04, 0.05, "Assumed", "FY2026 利息收入 195.0 ÷ 平均現金約 4.78B＝4.1%", None, None),
    ("int_x", "利息費用（US$M／年，優先票據）", "US$M", None, 28.0, None, "Derived", "FY2026 利息費用 28.0（Q4 FY2026 新聞稿）", None, None),
    ("sh_g", "稀釋股數年增（SBC 稀釋，不含回購）", "%", 0.0, 0.015, 0.03, "Assumed", "FY2026 稀釋股數約 +2%（Hist 頁，拆股後）", "回購未建模（FY2027 上半年回購 175.6，授權剩餘約 1.3B）", None),
    (SEC, "E. 公司指引、賣方共識與其他對照（只作對照，不進模型）"),
    ("g_rev27", "公司指引 FY2027 營收（中點）", "US$M", None, 6001.1, None, "Interested-party", Q2PR + "：5,991.1–6,011.1", "只作對照，不進模型", "公司指引"),
    ("g_arr27", "公司指引 FY2027 期末 ARR（中點）", "US$M", None, 6607.45, None, "Interested-party", Q2PR + "：6,603.0–6,611.9", "只作對照，不進模型", "公司指引"),
    ("g_ngop27", "公司指引 FY2027 non-GAAP 營業利益（中點）", "US$M", None, 1502.8, None, "Interested-party", Q2PR + "：1,497.2–1,508.4", "只作對照，不進模型", "公司指引"),
    ("g_eps27", "公司指引 FY2027 non-GAAP 稀釋 EPS（中點；拆股後）", "US$", None, 1.255, None, "Interested-party", Q2PR + "：1.25–1.26", "只作對照，不進模型", "公司指引"),
    ("c_rev27", "賣方共識 FY2027 營收", "US$M", None, 6000, None, "Interested-party", "https://stockanalysis.com/stocks/crwd/forecast（S&P Global；49 位；2026-11-02 更新）", "只作對照，不進模型", "賣方預估"),
    ("c_eps27", "賣方共識 FY2027 non-GAAP EPS（拆股後）", "US$", None, 1.26, None, "Interested-party", "同上", "只作對照，不進模型", "賣方預估"),
    ("c_rev28", "賣方共識『次一年』營收（頁面摘要框；年度待確認為 FY2028）", "US$M", None, 7380, None, "Interested-party", "同上（明細表付費鎖定）", "只作對照，不進模型；年度與 EPS 1.61 待確認", "賣方預估"),
    ("cons_tp", "賣方平均目標價（53 位；中位 250；高 325；低 132）", "US$", None, 249.68, None, "Interested-party", "https://stockanalysis.com/stocks/crwd/forecast（2026-11-03 查看）", "只作對照，不進模型", "賣方預估"),
]
SCN_ROWS = [  # (key, 項目, 單位, 引用鍵, 標記說明)
    ("scn_ero", "每模組年費年變動（價格侵蝕）", "%", ("ero_A", "ero_B", "ero_C")),
    ("scn_r32", "FY2032 AI 代理單位 ÷ 人類單位", "x", ("r32_A", "r32_B", "r32_C")),
    ("scn_pf", "AI 代理單位價格保留率", "x", ("pf_A", "pf_B", "pf_C")),
    ("scn_tasks", "FY2032 任務數／受保護單位／年", "次", ("tasks_A", "tasks_B", "tasks_C")),
    ("scn_w", "情境權重", "%", ("tp_wA", "tp_wB", "tp_wC")),
]

# ───────────────────────── 歷史（Hist 頁，US$M）─────────────────────────
# (key, 項目, FY2023, FY2024, FY2025, FY2026, 出處, 標記)
HK = "10-K FY2026 與 Q4 新聞稿（見來源欄）"
HIST = [
    ("sub", "訂閱營收", 2111.660, 2870.557, 3761.480, 4564.683),
    ("ps", "專業服務營收", 129.576, 184.998, 192.144, 247.322),
    ("rev", "合併營收", 2241.236, 3055.555, 3953.624, 4812.005),
    ("cogs", "GAAP 營業成本合計", 601.231, 755.723, 990.172, 1218.929),
    ("gp", "GAAP 毛利", 1640.005, 2299.832, 2963.452, 3593.076),
    ("sm", "GAAP 銷售及行銷費用", 904.409, 1140.566, 1523.001, 1831.254),
    ("rd", "GAAP 研發費用", 608.364, 768.497, 1075.587, 1384.770),
    ("ga", "GAAP 一般及行政費用", 317.344, 392.764, 481.264, 670.344),
    ("ebit", "GAAP 營業利益（損）", -190.112, -1.995, -116.400, -293.292),
    ("ie", "利息費用", -25.319, -25.756, -26.311, -28.021),
    ("ii", "利息收入", 52.495, 148.930, 196.174, 194.969),
    ("oi", "其他收入（費用）淨額", 3.053, 1.638, 5.101, -0.645),
    ("pt", "GAAP 稅前利益（損）", -159.883, 122.817, 58.564, -126.989),
    ("tx", "所得稅費用", 22.402, 32.232, 71.130, 34.176),
    ("ni", "GAAP 歸屬 CrowdStrike 淨利（損）", -183.245, 89.327, -15.241, -162.502),
    ("shd", "GAAP 稀釋加權股數（拆股前）", 233.139, 243.635, 246.933 if False else 244.750, 250.576),
    ("ngebit", "Non-GAAP 營業利益", 355.622, 660.295, 879.888, 1046.467),
    ("ngeps", "Non-GAAP 稀釋 EPS（拆股前，US$）", 1.54, 3.09, 3.24, 3.73),
    ("nggp", "Non-GAAP 毛利（訂閱＋專業服務）", None, None, 3095.344, 3762.300),
    ("sbc", "股票薪酬費用（現金流量表）", 526.504, 648.665, 861.391, 1096.679),
    ("sbct", "股票薪酬與相關雇主薪資稅（non-GAAP 調節項）", None, None, 903.584, 1130.571),
    ("da", "折舊與攤銷", 77.245, 126.838, 187.952, 250.218),
    ("am", "無形資產攤銷", 16.565, 18.416, 26.004, 31.233),
    ("cx", "購置不動產設備", 235.019, 176.529, 254.852, 302.108),
    ("cs", "資本化內部使用軟體", 29.095, 49.457, 58.969, 68.751),
    ("ocf", "營運現金流", 941.007, 1166.207, 1381.727, 1612.349),
    ("fcf", "自由現金流（公司口徑）", 676.829, 938.190, 1065.079, 1235.274),
    ("arr", "期末 ARR（FY2024 為四捨五入值）", 2560.0, 3440.0, 4241.838, 5252.751),
    ("nnarr", "淨新增 ARR", None, None, 806.7, 1010.0),
    ("wc_ar", "營運資產負債：應收帳款", -258.109, -217.699, -274.219, -232.528),
    ("wc_dac", "營運資產負債：遞延合約取得成本（新增）", -298.716, -371.649, -584.484, -703.707),
    ("wc_pp", "營運資產負債：預付與其他資產", -46.807, -102.520, -190.232, -206.157),
    ("wc_ap", "營運資產負債：應付帳款", -15.463, -18.898, 84.939, -11.267),
    ("wc_ac", "營運資產負債：應計費用與其他負債", 58.923, 14.586, 218.518, 22.587),
    ("wc_pay", "營運資產負債：應計薪資", 65.226, 65.102, 85.873, 61.610),
    ("wc_ls", "營運資產負債：營業租賃負債", -10.364, -14.035, -15.657, -13.692),
    ("wc_dr", "營運資產負債：遞延收入", 825.751, 696.639, 669.264, 1023.863),
    ("dac_am", "遞延合約取得成本攤銷（非現金，加回）", None, None, 318.837, 449.413),
    ("emp", "期末全職員工（FY2023–FY2024 為第二來源）", 7273, 7925, 10118, 10698),
]
HIST_NOTE = {"sbc": "FY2024 取 10-K FY2026 現金流量表（修正後）648.665；FY2024 新聞稿原值 631.519",
             "ebit": "FY2024 為 Q4 FY2024 新聞稿原值（FY2026 10-K 對 SBC 作無重大影響的修正）",
             "shd": "拆股前；FY2025、FY2026 為虧損，稀釋＝基本"}
# 季（Q 頁）：Q3FY26, Q4FY26, Q1FY27, Q2FY27 + 去年同季
QTR = ["Q3 FY2026", "Q4 FY2026", "Q1 FY2027", "Q2 FY2027"]
QDATA = [  # (項目, 4 個季值, 標記, 來源, 來源誘因)
    ("訂閱營收", [1168.705, 1242.265, 1320.853, 1400.291], "Verified", "各季新聞稿損益表（Q3 FY2026、Q4 FY2026 8-K、Q1 FY2027、Q2 FY2027）", None),
    ("專業服務營收", [65.539, 63.110, 64.776, 70.606], "Verified", "同上", None),
    ("合併營收", [1234.244, 1305.375, 1385.629, 1470.897], "Verified", "同上", None),
    ("GAAP 營業利益（損）", [-69.443, -6.900, -30.600, -33.232], "Verified", "同上", None),
    ("Non-GAAP 營業利益", [264.586, 325.806, 325.670, 371.649], "Interested-party", "同上（公司調節表）", "公司自訂 non-GAAP 口徑"),
    ("Non-GAAP 毛利率（總）", [0.78, 0.79, 0.79, 0.79], "Interested-party", "同上（新聞稿四捨五入整數百分比）", "公司自訂 non-GAAP 口徑；整數百分比"),
    ("SBC 與相關雇主薪資稅", [293.303, 302.196, 317.627, 399.001], "Verified", "同上調節表", None),
    ("營運現金流", [397.5, 497.869, 590.937, 530.268], "Verified", "同上現金流量表（Q3 FY2026 為四捨五入 US$M）", None),
    ("自由現金流（公司口徑）", [295.9, 376.363, 468.463, 377.407], "Interested-party", "同上", "公司自訂 FCF 口徑"),
    ("淨新增 ARR", [265.3, 330.7, 255.8, 332.8], "Interested-party", "同上", "公司自訂 KPI；10-Q 披露 DOJ／SEC 就營收認列與 ARR 報告方式的資訊請求"),
    ("期末 ARR", [4920.0, 5252.751, 5510.0, 5841.421], "Interested-party", "同上（Q3 FY2026、Q1 FY2027 為四捨五入值）", "公司自訂 KPI；同上資訊請求"),
]
Q_NG_GM_NOTE = "Non-GAAP 毛利率為新聞稿整數百分比（四捨五入），TTM 為四季簡單平均"

# ───────────────────────── 工具 ─────────────────────────
class Reg:
    """記錄各頁關鍵列號，供其他頁與 html.json 引用"""
    rows = {}


def f_(ws, cell, v, fmt=None, bold=False, blue=False, fill=None, wrap=False):
    c = ws[cell]
    c.value = v
    if fmt:
        c.number_format = fmt
    if bold:
        c.font = BOLD
    if blue:
        c.font = BLUE
    if fill:
        c.fill = fill
    if wrap:
        c.alignment = Alignment(wrap_text=True, vertical="top")
    return c


def addname(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def build(ov=None, sens=None, out=None):
    """ov：In 頁基準欄覆寫 {key: value}（敏感度用）；sens：Sens 頁內容（dict）；回傳 (路徑, build_id)"""
    ov = ov or {}
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    names = ["README", "In", "Hist", "Q", "Thesis", "PL", "CF", "Val", "RevDCF", "Recon", "Sens", "Sources"]
    ws_ = {n: wb.create_sheet(n) for n in names}
    # ───────── In ─────────
    In = ws_["In"]
    f_(In, "A1", "In — 輸入（藍字；基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁、不進目標價；情境 A／B／C 的驅動見本頁最後一節「情境驅動對照」）。金額 US$M；每股 US$（4 比 1 拆股後）；股數百萬股。標記：Verified／Interested-party／Analogy／Assumed／Derived；Interested-party 列須有來源誘因。", wrap=False)
    f_(In, "A2", "事實總帳：data/facts_crowdstrike.csv（逐列出處）與 data/research_crowdstrike.md；Tokenomics v22.0 快照 data/tk_snapshot_v22.0.json（推論成本只讀快照）。本版未做先驗（Andy 2026-11-16 指示先跳過）。")
    for j, h in enumerate(["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "說明", "來源誘因（Interested-party 列；混合來源列寫其他欄位）"], 1):
        f_(In, f"{openpyxl.utils.get_column_letter(j)}3", h, bold=True, fill=HDR)
    r = 4
    keyrow = {}
    for it in INP:
        if it[0] == SEC:
            f_(In, f"A{r}", it[1], bold=True)
            r += 1
            continue
        key, lab, unit, lo, base, hi, tag, src, note, inc = it
        if key in ov:
            base = ov[key]
        keyrow[key] = r
        f_(In, f"A{r}", key)
        f_(In, f"B{r}", lab)
        f_(In, f"C{r}", unit)
        if lo is not None:
            f_(In, f"D{r}", lo, blue=True)
        f_(In, f"E{r}", base, blue=not (isinstance(base, str) and base.startswith("=")))
        if hi is not None:
            f_(In, f"F{r}", hi, blue=True)
        f_(In, f"G{r}", tag)
        f_(In, f"H{r}", src)
        if note:
            f_(In, f"I{r}", note)
        if inc:
            f_(In, f"J{r}", inc)
        if isinstance(base, D):
            In[f"E{r}"].number_format = "yyyy-mm-dd"
        elif unit == "%":
            for c in "DEF":
                In[f"{c}{r}"].number_format = "0.0%"
        addname(wb, key, f"In!$E${r}")
        r += 1
    r += 1
    f_(In, f"A{r}", "J. 情境驅動對照（三情境並排，進加權目標價；本節只引用上方各列，修改請改上方藍字）", bold=True)
    r += 1
    for j, h in enumerate(["鍵", "項目", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "標記", "說明"], 1):
        f_(In, f"{openpyxl.utils.get_column_letter(j)}{r}", h, bold=True, fill=HDR)
    r += 1
    for k, lab, unit, ks in SCN_ROWS:
        f_(In, f"A{r}", k)
        f_(In, f"B{r}", lab)
        f_(In, f"C{r}", unit)
        for c, kk in zip("DEF", ks):
            f_(In, f"{c}{r}", f"=E{keyrow[kk]}")
            if unit == "%":
                In[f"{c}{r}"].number_format = "0.0%"
        f_(In, f"G{r}", "Derived")
        f_(In, f"H{r}", "引用 In 頁 " + "／".join(ks) + "；說明見上方各列")
        r += 1
    In.column_dimensions["B"].width = 60
    In.column_dimensions["H"].width = 60
    In.column_dimensions["I"].width = 50
    # ───────── Hist ─────────
    H = ws_["Hist"]
    f_(H, "A1", "Hist — 歷史財務（US$M；拆股前之股數與 EPS；來源：10-K FY2026、Q4 FY2026／FY2025／FY2024／FY2023 新聞稿，逐列出處見 data/facts_crowdstrike.csv）。公式只引用本頁（原則第 10 節第 14 條）。", bold=True)
    for j, h in enumerate(["項目", "鍵", "FY2023", "FY2024", "FY2025", "FY2026", "標記", "來源", "說明"], 1):
        f_(H, f"{openpyxl.utils.get_column_letter(j)}2", h, bold=True, fill=HDR)
    hr = {}
    for i, row in enumerate(HIST):
        rr = 3 + i
        key, lab, *vals = row
        hr[key] = rr
        f_(H, f"A{rr}", lab)
        f_(H, f"B{rr}", key)
        for c, v in zip("CDEF", vals):
            if v is not None:
                f_(H, f"{c}{rr}", v, "#,##0.000", blue=True)
            else:
                f_(H, f"{c}{rr}", "未找到")
        tag = "Interested-party" if key in ("ngebit", "ngeps", "nggp", "fcf", "arr", "nnarr", "sbct") else "Verified"
        if key == "emp":
            tag = "Verified"
        f_(H, f"G{rr}", tag)
        f_(H, f"H{rr}", "10-K FY2026／Q4 新聞稿／Q4 FY2024 新聞稿（facts_crowdstrike.csv）")
        if tag == "Interested-party":
            f_(H, f"I{rr}", "來源誘因：公司自訂 non-GAAP／FCF／KPI 口徑；" + HIST_NOTE.get(key, ""))
        elif key in HIST_NOTE:
            f_(H, f"I{rr}", HIST_NOTE[key])
        for y, c in zip((2023, 2024, 2025, 2026), "CDEF"):
            addname(wb, f"h_{key}{str(y)[2:]}", f"Hist!${c}${rr}")
    rr = 3 + len(HIST) + 1
    f_(H, f"A{rr}", "核對：訂閱＋專業服務 − 合併營收", bold=True)
    for c in "CDEF":
        f_(H, f"{c}{rr}", f"={c}{hr['sub']}+{c}{hr['ps']}-{c}{hr['rev']}", "0.000")
    Reg.rows["Hist_check"] = rr
    r_ngc, r_cxt = rr + 2, rr + 3
    assert (r_ngc, r_cxt, hr["rev"], hr["sbc"]) == (45, 46, 5, 22), (r_ngc, r_cxt, hr["rev"], hr["sbc"])
    f_(H, f"A{r_ngc}", "Non-GAAP 總成本（營收 − non-GAAP 營業利益；規模經濟回歸用，Derived）")
    f_(H, f"A{r_cxt}", "資本支出合計（購置不動產設備＋資本化內部使用軟體；Derived）")
    for y, c in zip((2023, 2024, 2025, 2026), "CDEF"):
        f_(H, f"{c}{r_ngc}", f"={c}{hr['rev']}-{c}{hr['ngebit']}", "#,##0.000")
        f_(H, f"{c}{r_cxt}", f"={c}{hr['cx']}+{c}{hr['cs']}", "#,##0.000")
        addname(wb, f"h_ngc{str(y)[2:]}", f"Hist!${c}${r_ngc}")
        addname(wb, f"h_cxt{str(y)[2:]}", f"Hist!${c}${r_cxt}")
    f_(H, f"G{r_ngc}", "Derived"); f_(H, f"G{r_cxt}", "Derived")
    H.column_dimensions["A"].width = 46
    # ───────── Q ─────────
    Q = ws_["Q"]
    f_(Q, "A1", "Q — 最近四季實績（US$M；來源：Q3 FY2026、Q4 FY2026、Q1 FY2027、Q2 FY2027 新聞稿與 10-Q；會計年度 2 月制，Q2 FY2027＝2026-05-01～2026-07-31）。Q3 FY2027 尚未公布。")
    f_(Q, "A2", "核對列（實績必須為 0）：訂閱＋專業服務 − 合併營收（第 8 列）。" + Q_NG_GM_NOTE + "。")
    for j, h in enumerate(["項目", "Q3 FY2026", "Q4 FY2026", "Q1 FY2027", "Q2 FY2027", "", "", "", "", "", "", "", "近四季 TTM", "", "標記", "來源", "來源誘因（Interested-party 列）"], 1):
        f_(Q, f"{openpyxl.utils.get_column_letter(j)}3", h, bold=True, fill=HDR)
    f_(Q, "A4", "營收（官方營收線）與關鍵指標", bold=True)
    qrow = {}
    rr = 5
    # 順序：5 訂閱、6 專業服務、7 合併營收、8 核對、9 GAAP EBIT、10 ngEBIT、11 ng GM、12.. 其餘
    # 為了讓 In 頁公式固定引用，列號固定：5 sub 6 ps 7 rev 8 check 9 ebit 10 ngebit 11 nggm 12 sbc 13 ocf 14 fcf 15 nnarr 16 arr
    order = [0, 1, 2, "chk", 3, 4, 5, 6, 7, 8, 9, 10]
    for o in order:
        if o == "chk":
            f_(Q, f"A{rr}", "檢查：訂閱＋專業服務 − 合併營收")
            for c in "BCDE":
                f_(Q, f"{c}{rr}", f"={c}5+{c}6-{c}7", "0.000")
            f_(Q, f"M{rr}", "=SUM(B8:E8)", "0.000")
            f_(Q, f"O{rr}", "Derived")
            f_(Q, f"P{rr}", "由上列相加")
        else:
            lab, vals, tag, src, inc = QDATA[o]
            f_(Q, f"A{rr}", lab)
            for c, v in zip("BCDE", vals):
                f_(Q, f"{c}{rr}", v, "0%" if "毛利率" in lab else "#,##0.0", blue=True)
            if "毛利率" in lab:
                f_(Q, f"M{rr}", f"=AVERAGE(B{rr}:E{rr})", "0.0%")
            elif "期末 ARR" in lab:
                f_(Q, f"M{rr}", f"=E{rr}", "#,##0.0")
            else:
                f_(Q, f"M{rr}", f"=SUM(B{rr}:E{rr})", "#,##0.0")
            f_(Q, f"O{rr}", tag)
            f_(Q, f"P{rr}", src)
            if inc:
                f_(Q, f"Q{rr}", inc)
        qrow[rr] = True
        rr += 1
    # 列 12、13 為 In 公式用：opx（毛利率 − 營業利益率）所需分子分母
    f_(Q, "A17", "衍生：TTM non-GAAP 營業利益率")
    f_(Q, "M17", "=M10/M7", "0.0%")
    f_(Q, "O17", "Derived")
    f_(Q, "P17", "TTM non-GAAP 營業利益 ÷ TTM 營收")
    f_(Q, "A18", "衍生：TTM SBC 占營收")
    f_(Q, "M18", "=M12/M7", "0.0%")
    f_(Q, "O18", "Derived")
    f_(Q, "P18", "TTM SBC 與相關雇主薪資稅 ÷ TTM 營收")
    f_(Q, "A19", "衍生：TTM GAAP 營業利益率")
    f_(Q, "M19", "=M9/M7", "0.0%")
    f_(Q, "O19", "Derived")
    f_(Q, "P19", "TTM GAAP 營業利益 ÷ TTM 營收")
    f_(Q, "A20", "衍生：TTM 營收")
    f_(Q, "M20", "=M7", "#,##0.0")
    f_(Q, "O20", "Derived")
    f_(Q, "P20", "四季合計")
    Q.column_dimensions["A"].width = 44
    # ───────── 年度欄與共用 ─────────
    ycols = {y: YC[y] for y in YEARS}          # C..K
    def tf(*parts):
        """組字公式：字串＝文字（拆成 ≤200 字的片段）；(expr, fmt)＝TEXT(expr,fmt)；('raw', expr)＝原樣"""
        out = []
        for p in parts:
            if isinstance(p, str):
                for i in range(0, len(p), 200):
                    out.append('"' + p[i:i + 200].replace('"', '""') + '"')
            elif p[0] == "raw":
                out.append(p[1])
            else:
                out.append(f'TEXT({p[0]},"{p[1]}")')
        return "=" + "&".join(out)
    nA = {y: y - 2026 for y in YEARS}
    # ───────── Thesis ─────────
    T = ws_["Thesis"]
    f_(T, "A1", "Thesis — 命題模組（原型 P：按防護範圍計價）：ARR＝受保護單位 × 每單位模組數 × 每模組年費；AI 淨效果＝AI 代理防護增量 − 定價侵蝕 − 增量推論成本。US$M；單位為百萬；v0.2 改讀 unit_economics。", bold=True)
    f_(T, "A2", "推論成本只讀 Tokenomics v22.0 快照（In 頁 tok_task、px_tok）；基期（FY2026）推論成本已含在實績毛利，模型只把增量扣在 non-GAAP 毛利。")
    PB = {}
    P_OFF = {"title": 0, "yr": 1, "sub": 2, "ps": 3, "rev": 4, "g": 5, "chk": 6, "u": 7, "m": 8, "p": 9, "chk2": 10, "cogs": 11, "gp": 12,
             "gm": 13, "opx": 14, "ebit": 15, "ebitm": 16, "oth": 17, "pt": 18, "tax": 19, "ni": 20, "nim": 21, "eps": 22, "nig": 23, "sh": 24,
             "nggp": 25, "nggm": 26, "ngop": 27, "ngopm": 28, "sbc": 29, "ngeps": 30, "inf": 31}
    CB = {}
    C_OFF = {"title": 0, "yr": 1, "rev": 2, "gr": 3, "ebit": 4, "tax": 5, "nopat": 6, "da": 7, "cx": 8, "wc": 9, "sbcab": 10, "fcff": 11, "int": 12, "fcf": 13,
             "cash": 14, "t": 15, "cf": 16, "df": 17, "pv": 18, "spv": 19, "n33": 20, "tv": 21, "tvdf": 22, "tvpv": 23, "ev": 24, "nc": 25, "eq": 26, "ps0": 27,
             "ps1": 28, "tvsh": 29, "gate": 30, "mincash": 31}
    for bi, s in enumerate(SC):
        PB[s] = {k: 4 + bi * 34 + v for k, v in P_OFF.items()}
        CB[s] = {k: 4 + bi * 34 + v for k, v in C_OFF.items()}
    TB = {}
    T_OFF = {"title": 0, "yr": 1, "n": 2, "gu": 3, "U": 4, "dm": 5, "M": 6, "P": 7, "core": 8, "r": 9, "aU": 10, "aARR": 11, "ARR": 12,
             "nn": 13, "arrg": 14, "sub": 15, "ps": 16, "rev": 17, "aiRev": 18, "dev": 19, "ero": 20, "tpu": 21, "tasks": 22, "cost": 23,
             "inc": 24, "floor": 25, "net": 26, "share": 27, "be": 28, "chk": 29}
    for bi, s in enumerate(SC):
        r0 = 4 + bi * 32
        TB[s] = {k: r0 + v for k, v in T_OFF.items()}
        R = TB[s]
        f_(T, f"A{R['title']}", f"{SCN_NAME[s]}：命題模組", bold=True)
        labels = {"yr": "年度", "n": "年數（自 FY2026 起）", "gu": "人類端受保護單位年增", "U": "人類端受保護單位（百萬）",
                  "dm": "每單位模組數年增率", "M": "每單位平均模組數", "P": "每模組年費（US$）", "core": "核心 ARR（US$M）＝單位 × 模組 × 年費",
                  "r": "AI 代理單位 ÷ 人類單位", "aU": "AI 代理（非人類身分）受保護單位（百萬）", "aARR": "AI 代理防護 ARR（US$M）",
                  "ARR": "合計期末 ARR（US$M）", "nn": "淨新增 ARR（US$M）", "arrg": "ARR 年增率", "sub": "訂閱營收（US$M）＝訂閱轉換率 × 平均 ARR",
                  "ps": "專業服務營收（US$M）", "rev": "合併營收（US$M）", "aiRev": "AI 增量營收：AI 代理防護（US$M）",
                  "dev": "核心 ARR 相對基期年費的價差（US$M）", "ero": "定價侵蝕（US$M，負值）", "tpu": "AI 輔助任務數／受保護單位／年",
                  "cost": "AI 推論成本（US$M，市場價；含基期）", "inc": "增量推論成本（US$M，相對 FY2026 基期）", "floor": "對照：以模型商全成本計（US$M，下限，不進模型）",
                  "net": "AI 淨效果（US$M）＝AI 增量營收 + 定價侵蝕 − 增量推論成本", "share": "推論成本占營業成本比（non-GAAP）",
                  "be": "損益兩平價格保留率（AI 代理單位；使 AI 淨效果＝0）", "chk": "檢查：合計 ARR − (核心＋AI 代理)、營收 − (訂閱＋專業服務)",
                  "tasks": "總任務數（百萬次）"}
        for k, lab in labels.items():
            f_(T, f"A{R[k]}", lab)
        T_ = lambda k, c: f"{c}{R[k]}"
        for y in range(2026, 2033):
            c = ycols[y]
            pc = ycols[y - 1]
            f_(T, T_("yr", c), f"FY{y}{'A' if y == 2026 else 'E'}", bold=True)
            f_(T, T_("n", c), nA[y])
            if y == 2026:
                f_(T, T_("U", c), "=u_0", "0.0")
                f_(T, T_("M", c), "=m_0", "0.00")
                f_(T, T_("P", c), "=p_0", "0.0")
                f_(T, T_("r", c), 0, "0.00")
                f_(T, T_("tpu", c), "=tasks_0", "0")
                f_(T, T_("dev", c), 0, "#,##0.0")
                f_(T, T_("sub", c), "=h_sub26", "#,##0.0")
                f_(T, T_("ps", c), "=h_ps26", "#,##0.0")
                f_(T, T_("aiRev", c), 0, "#,##0.0")
                f_(T, T_("ero", c), 0, "#,##0.0")
                f_(T, T_("inc", c), 0, "#,##0.0")
            else:
                f_(T, T_("gu", c), f"=gu_27+(gu_32-gu_27)*({c}{R['n']}-1)/5", "0.0%")
                f_(T, T_("U", c), f"={pc}{R['U']}*(1+{c}{R['gu']})", "0.0")
                f_(T, T_("dm", c), f"=gm_27+(gm_32-gm_27)*({c}{R['n']}-1)/5", "0.0%")
                f_(T, T_("M", c), f"={pc}{R['M']}*(1+{c}{R['dm']})", "0.00")
                f_(T, T_("P", c), f"={pc}{R['P']}*(1+ero_{s})", "0.0")
                f_(T, T_("r", c), f"=r32_{s}*{c}{R['n']}/6", "0.00")
                f_(T, T_("tpu", c), f"=tasks_0+(tasks_{s}-tasks_0)*{c}{R['n']}/6", "0")
                f_(T, T_("dev", c), f"=({c}{R['P']}-$E{R['P']})*{c}{R['U']}*{c}{R['M']}", "#,##0.0")
                f_(T, T_("sub", c), f"=conv*(1-ev_hc)*AVERAGE({pc}{R['ARR']},{c}{R['ARR']})", "#,##0.0")
                f_(T, T_("ps", c), f"=ps_pct*{c}{R['sub']}", "#,##0.0")
                f_(T, T_("aiRev", c), f"=conv*(1-ev_hc)*AVERAGE({pc}{R['aARR']},{c}{R['aARR']})", "#,##0.0")
                f_(T, T_("ero", c), f"=conv*(1-ev_hc)*AVERAGE({pc}{R['dev']},{c}{R['dev']})", "#,##0.0")
                f_(T, T_("inc", c), f"={c}{R['cost']}-$E{R['cost']}", "#,##0.0")
                f_(T, T_("nn", c), f"={c}{R['ARR']}-{pc}{R['ARR']}", "#,##0.0")
                f_(T, T_("arrg", c), f"={c}{R['ARR']}/{pc}{R['ARR']}-1", "0.0%")
            f_(T, T_("core", c), f"={c}{R['U']}*{c}{R['M']}*{c}{R['P']}", "#,##0.0")
            f_(T, T_("aU", c), f"={c}{R['U']}*{c}{R['r']}", "0.0")
            f_(T, T_("aARR", c), f"={c}{R['U']}*{c}{R['r']}*ma_n*{c}{R['P']}*pf_{s}", "#,##0.0")
            f_(T, T_("ARR", c), f"={c}{R['core']}+{c}{R['aARR']}", "#,##0.0")
            f_(T, T_("rev", c), f"={c}{R['sub']}+{c}{R['ps']}", "#,##0.0")
            f_(T, T_("tasks", c), f"=({c}{R['U']}+{c}{R['aU']})*{c}{R['tpu']}", "#,##0")
            f_(T, T_("cost", c), f"={c}{R['tasks']}*tok_task*px_tok/1000000", "#,##0.0")
            f_(T, T_("floor", c), f"={c}{R['tasks']}*tok_task*fc_tok/1000000", "#,##0.0")
            f_(T, T_("net", c), f"={c}{R['aiRev']}+{c}{R['ero']}-{c}{R['inc']}", "#,##0.0")
            f_(T, T_("be", c), f'=IFERROR(({c}{R["inc"]}-{c}{R["ero"]})/({c}{R["aiRev"]}/pf_{s}),"n/m")', "0.00")
            f_(T, T_("chk", c), f"=({c}{R['ARR']}-({c}{R['core']}+{c}{R['aARR']}))+({c}{R['rev']}-({c}{R['sub']}+{c}{R['ps']}))", "0.000")
        for y in range(2027, 2033):
            c = ycols[y]
            f_(T, f"{c}{R['share']}", f"={c}{R['inc']}/(PL!{c}{PB[s]['rev']}-PL!{c}{PB[s]['nggp']})", "0.0%")
        f_(T, f"E{R['ARR']}", f"=E{R['core']}+E{R['aARR']}", "#,##0.0")
        f_(T, f"E{R['arrg']}", "=h_arr26/h_arr25-1", "0.0%")
        f_(T, f"E{R['nn']}", "=h_arr26-h_arr25", "#,##0.0")
        f_(T, f"E{R['rev']}", f"=E{R['sub']}+E{R['ps']}", "#,##0.0")
    T.column_dimensions["A"].width = 66
    # ───────── PL ─────────
    P = ws_["PL"]
    f_(P, "A1", "PL — 損益（GAAP 為主；Non-GAAP 為公司自訂口徑，另列；US$M；每股 US$，4 比 1 拆股後）。官方報告分部：單一應報導分部；營收線＝訂閱、專業服務（segments.json）。三情境並列：情境 A／B／C。", bold=True)
    f_(P, "A2", "FY2024A–FY2026A 為實績（Hist 頁）；FY2027E 起為模型；會計年度 2 月制（FY2027＝2026-02-01～2027-01-31）。FY2027 含上半年實績但整年以模型推估。")
    cf = ws_["CF"]
    f_(cf, "A1", "CF — 現金流與 DCF（FCFF＝NOPAT＋折舊攤銷 − 資本支出＋營運資金流入；股票薪酬不加回，加回只作敏感度）。US$M；DCF 折現到評價日 2026-07-31（期中慣例），FY2027 只計評價日後的下半年（FY2027E FCFF − 上半年實績 FCF）。", bold=True)
    f_(cf, "A2", "營運資金：以『營運資金流入 ÷ 營收增量』建簡化模組（In 頁 wc_in；FY2025 0.348、FY2026 0.454 見 Hist 頁），遞延收入為主。收支關卡＝FY2026–FY2030 自由現金流（公司口徑近似）合計，CRWD 無股利。")
    for s in SC:
        R, B_, C_ = PB[s], TB[s], CB[s]
        f_(P, f"A{R['title']}", f"{SCN_NAME[s]}：標準損益表（GAAP；US$M；EPS 為 US$）", bold=True)
        pl_labels = {"yr": "年度", "sub": "訂閱", "ps": "專業服務", "rev": "營收", "g": "營收年增率", "chk": "檢查：訂閱＋專業服務 − 合併營收",
                     "u": "端點與身分基數（命題子列）", "m": "模組密度（命題子列）", "p": "每模組年費（命題子列）",
                     "chk2": "檢查：命題子列（基數×模組×單價）− 訂閱營收", "cogs": "營業成本（GAAP）", "gp": "毛利（GAAP）", "gm": "毛利率（GAAP）",
                     "opx": "營業費用（GAAP）", "ebit": "營業利益（GAAP）", "ebitm": "營業利益率（GAAP）", "oth": "其他收入（淨利息與其他）",
                     "pt": "稅前利益", "tax": "所得稅", "ni": "稅後淨利（GAAP，歸屬 CrowdStrike）", "nim": "淨利率（GAAP）", "eps": "EPS（GAAP 稀釋，拆股後，US$）",
                     "nig": "淨利年增率", "sh": "稀釋股數（百萬股，拆股後）", "nggp": "Non-GAAP 毛利（公司口徑，不含 SBC 與攤銷）", "nggm": "Non-GAAP 毛利率",
                     "ngop": "Non-GAAP 營業利益（公司口徑）", "ngopm": "Non-GAAP 營業利益率", "sbc": "SBC 與相關雇主薪資稅（調節項）",
                     "ngeps": "Non-GAAP 稀釋 EPS（公司口徑，拆股後，US$）", "inf": "增量 AI 推論成本（已扣於 Non-GAAP 毛利）"}
        for k, lab in pl_labels.items():
            f_(P, f"A{R[k]}", lab)
        for k, un in (("u", "百萬單位"), ("m", "每單位模組數"), ("p", "US$ 每模組年費")):
            f_(P, f"B{R[k]}", un)
        hist_col = {2024: "D", 2025: "E", 2026: "F"}
        for y in YEARS:
            c = ycols[y]
            pc = ycols[y - 1] if y > 2024 else None
            f_(P, f"{c}{R['yr']}", f"FY{y}{'A' if y <= 2026 else 'E'}" + ("" if y <= 2026 else ""), bold=True)
            if y <= 2026:
                hc = hist_col[y]
                H_ = lambda k: f"Hist!{hc}{hr[k]}"
                f_(P, f"{c}{R['sub']}", f"={H_('sub')}", "#,##0.0")
                f_(P, f"{c}{R['ps']}", f"={H_('ps')}", "#,##0.0")
                f_(P, f"{c}{R['chk']}", f"={c}{R['sub']}+{c}{R['ps']}-{H_('rev')}", "0.000")
                f_(P, f"{c}{R['cogs']}", f"={H_('cogs')}", "#,##0.0")
                f_(P, f"{c}{R['opx']}", f"={H_('sm')}+{H_('rd')}+{H_('ga')}", "#,##0.0")
                f_(P, f"{c}{R['ebit']}", f"={H_('ebit')}", "#,##0.0")
                f_(P, f"{c}{R['oth']}", f"={H_('ie')}+{H_('ii')}+{H_('oi')}", "#,##0.0")
                f_(P, f"{c}{R['pt']}", f"={H_('pt')}", "#,##0.0")
                f_(P, f"{c}{R['tax']}", f"={H_('tx')}", "#,##0.0")
                f_(P, f"{c}{R['ni']}", f"={H_('ni')}", "#,##0.0")
                f_(P, f"{c}{R['sh']}", f"={H_('shd')}*split", "#,##0.0")
                f_(P, f"{c}{R['ngop']}", f"={H_('ngebit')}", "#,##0.0")
                f_(P, f"{c}{R['ngeps']}", f"={H_('ngeps')}/split", "0.00")
                if y >= 2025:
                    f_(P, f"{c}{R['nggp']}", f"={H_('nggp')}", "#,##0.0")
                    f_(P, f"{c}{R['sbc']}", f"={H_('sbct')}", "#,##0.0")
                    f_(P, f"{c}{R['nggm']}", f"={c}{R['nggp']}/{c}{R['rev']}", "0.0%")
                else:
                    f_(P, f"{c}{R['nggp']}", "未找到")
                    f_(P, f"{c}{R['sbc']}", "未找到")
                    f_(P, f"{c}{R['nggm']}", "未找到")
                if y == 2026:
                    f_(P, f"{c}{R['u']}", f"=Thesis!E{B_['U']}", "0.0")
                    f_(P, f"{c}{R['m']}", f"=Thesis!E{B_['M']}", "0.00")
                    f_(P, f"{c}{R['p']}", f"=Thesis!E{B_['P']}", "0.0")
                    f_(P, f"{c}{R['chk2']}", f"=conv*AVERAGE(h_arr25,{c}{R['u']}*{c}{R['m']}*{c}{R['p']})-{c}{R['sub']}", "0.000")
                else:
                    for k in ("u", "m", "p"):
                        f_(P, f"{c}{R[k]}", "未揭露")
                    f_(P, f"{c}{R['chk2']}", "n/a")
                f_(P, f"{c}{R['inf']}", 0, "#,##0.0")
                f_(P, f"{c}{R['gp']}", f"={H_('gp')}", "#,##0.0")
            else:
                tc = c
                f_(P, f"{c}{R['sub']}", f"=Thesis!{tc}{B_['sub']}", "#,##0.0")
                f_(P, f"{c}{R['ps']}", f"=Thesis!{tc}{B_['ps']}", "#,##0.0")
                f_(P, f"{c}{R['chk']}", f"={c}{R['sub']}+{c}{R['ps']}-Thesis!{tc}{B_['rev']}", "0.000")
                f_(P, f"{c}{R['u']}", f"=Thesis!{tc}{B_['U']}", "0.0")
                f_(P, f"{c}{R['m']}", f"=Thesis!{tc}{B_['M']}", "0.00")
                f_(P, f"{c}{R['p']}", f"=Thesis!{tc}{B_['P']}", "0.0")
                f_(P, f"{c}{R['chk2']}", f"=conv*(1-ev_hc)*AVERAGE({pc}{R['u']}*{pc}{R['m']}*{pc}{R['p']}+Thesis!{pc}{B_['aARR']},{c}{R['u']}*{c}{R['m']}*{c}{R['p']}+Thesis!{c}{B_['aARR']})-{c}{R['sub']}", "0.000")
                f_(P, f"{c}{R['inf']}", f"=Thesis!{tc}{B_['inc']}", "#,##0.0")
                n_ = f"Thesis!{tc}{B_['n']}"
                f_(P, f"{c}{R['nggp']}", f"={c}{R['sub']}*gmsub_ng+{c}{R['ps']}*gmps_ng-{c}{R['inf']}", "#,##0.0")
                f_(P, f"{c}{R['nggm']}", f"={c}{R['nggp']}/{c}{R['rev']}", "0.0%")
                f_(P, f"{c}{R['ngop']}", f"={c}{R['nggp']}-opx_27*{c}{R['rev']}+fix_ng*({c}{R['rev']}/$F{R['rev']}-1)", "#,##0.0")
                f_(P, f"{c}{R['sbc']}", f"=sbc_27*{c}{R['rev']}-fix_sbc*({c}{R['rev']}/$F{R['rev']}-1)", "#,##0.0")
                f_(P, f"{c}{R['ebit']}", f"={c}{R['ngop']}-{c}{R['sbc']}-oth_pct*{c}{R['rev']}", "#,##0.0")
                f_(P, f"{c}{R['gp']}", f"={c}{R['nggp']}-gm_gap*{c}{R['rev']}", "#,##0.0")
                f_(P, f"{c}{R['cogs']}", f"={c}{R['rev']}-{c}{R['gp']}", "#,##0.0")
                f_(P, f"{c}{R['opx']}", f"={c}{R['gp']}-{c}{R['ebit']}", "#,##0.0")
                f_(P, f"{c}{R['oth']}", f"=int_y*CF!{pc}{C_['cash']}-int_x", "#,##0.0")
                f_(P, f"{c}{R['pt']}", f"={c}{R['ebit']}+{c}{R['oth']}", "#,##0.0")
                f_(P, f"{c}{R['tax']}", f"=MAX(0,{c}{R['pt']})*tax", "#,##0.0")
                f_(P, f"{c}{R['ni']}", f"={c}{R['pt']}-{c}{R['tax']}", "#,##0.0")
                f_(P, f"{c}{R['sh']}", "=sh_dil" if y == 2027 else f"={pc}{R['sh']}*(1+sh_g)", "#,##0.0")
                f_(P, f"{c}{R['ngeps']}", f"=({c}{R['ngop']}+{c}{R['oth']})*(1-tax)/{c}{R['sh']}", "0.00")
            f_(P, f"{c}{R['rev']}", f"={c}{R['sub']}+{c}{R['ps']}", "#,##0.0")
            if y >= 2024:
                prev_rev = f"{pc}{R['rev']}" if pc else "Hist!C5"
                f_(P, f"{c}{R['g']}", f"={c}{R['rev']}/{prev_rev}-1", "0.0%")
            f_(P, f"{c}{R['gm']}", f"={c}{R['gp']}/{c}{R['rev']}", "0.0%")
            f_(P, f"{c}{R['ebitm']}", f"={c}{R['ebit']}/{c}{R['rev']}", "0.0%")
            f_(P, f"{c}{R['nim']}", f"={c}{R['ni']}/{c}{R['rev']}", "0.0%")
            f_(P, f"{c}{R['eps']}", f"={c}{R['ni']}/{c}{R['sh']}", "0.00")
            f_(P, f"{c}{R['ngopm']}", f"={c}{R['ngop']}/{c}{R['rev']}", "0.0%")
            if pc:
                f_(P, f"{c}{R['nig']}", f'=IF({pc}{R["ni"]}<=0,"n/m",{c}{R["ni"]}/{pc}{R["ni"]}-1)', "0.0%")
            else:
                f_(P, f"{c}{R['nig']}", "n/m")
        # ── CF 區塊 ──
        f_(cf, f"A{C_['title']}", f"{SCN_NAME[s]}：現金流與 DCF（US$M；FY2033–FY2037 為延伸期）", bold=True)
        cl = {"yr": "年度", "rev": "營收", "gr": "營收年增率", "ebit": "營業利益（GAAP）", "tax": "稅率", "nopat": "NOPAT＝EBIT ×（1 − 稅率）（EBIT 為負時不課稅）", "da": "＋折舊與攤銷",
              "cx": "− 資本支出（含資本化軟體）", "wc": "＋營運資金流入（遞延收入等）", "sbcab": "＋股票薪酬加回（預設 0）", "fcff": "＝無槓桿自由現金流（FCFF；SBC 費用化）",
              "int": "稅後淨利息與其他", "fcf": "自由現金流（現金口徑近似）＝FCFF＋非現金 SBC＋稅後淨利息", "cash": "年末現金（無回購、無併購）",
              "t": "折現時點（自評價日，年；期中）", "cf": "DCF 現金流（FY2027 扣上半年實績）", "df": "折現因子", "pv": "現值",
              "spv": "明確期與延伸期現值合計", "n33": "終年 NOPAT × (1＋g)（FY2038）", "tv": "終值", "tvdf": "終值折現因子", "tvpv": "終值現值",
              "ev": "企業價值 EV", "nc": "＋淨現金（現金 − 長期負債，2026-07-31）", "eq": "股權價值", "ps0": "每股價值（評價日，US$）",
              "ps1": "每股價值（目標時點，US$）", "tvsh": "終值占 EV", "gate": "收支關卡：FY2026–FY2030 自由現金流合計（US$M）", "mincash": "預測期最低年末現金（US$M）"}
        for k, lab in cl.items():
            f_(cf, f"A{C_[k]}", lab)
        ext_cols = {y: openpyxl.utils.get_column_letter(5 + (y - 2026)) for y in range(2026, 2038)}   # E..P
        for y in range(2026, 2038):
            c = ext_cols[y]
            pc = ext_cols[y - 1] if y > 2026 else None
            f_(cf, f"{c}{C_['yr']}", f"FY{y}{'A' if y == 2026 else 'E'}", bold=True)
            if y == 2026:
                f_(cf, f"{c}{C_['rev']}", f"=PL!E{R['rev']}", "#,##0.0")
                f_(cf, f"{c}{C_['fcf']}", "=h_fcf26", "#,##0.0")
                f_(cf, f"{c}{C_['cash']}", "=cash_26", "#,##0.0")
                continue
            k_ = y - 2032
            if y <= 2032:
                n_ = f"Thesis!{c}{B_['n']}"
                f_(cf, f"{c}{C_['rev']}", f"=PL!{c}{R['rev']}", "#,##0.0")
                f_(cf, f"{c}{C_['ebit']}", f"=PL!{c}{R['ebit']}", "#,##0.0")
                f_(cf, f"{c}{C_['cx']}", f"=cx_27*{c}{C_['rev']}-fix_cx*({c}{C_['rev']}/$F{C_['rev']}-1)", "#,##0.0")
                f_(cf, f"{c}{C_['sbcab']}", f"=sbc_ab*sbc_nc*PL!{c}{R['sbc']}", "#,##0.0")
            else:
                f_(cf, f"{c}{C_['gr']}", f"=$K{C_['gr']}+(g_tv-$K{C_['gr']})*{k_}/5", "0.0%")
                f_(cf, f"{c}{C_['rev']}", f"={pc}{C_['rev']}*(1+{c}{C_['gr']})", "#,##0.0")
                f_(cf, f"{c}{C_['ebit']}", f"=(PL!$K{R['ebitm']}+ext_d*{k_}/5)*{c}{C_['rev']}", "#,##0.0")
                f_(cf, f"{c}{C_['cx']}", f"=$K{C_['cx']}/$K{C_['rev']}*{c}{C_['rev']}", "#,##0.0")
                f_(cf, f"{c}{C_['sbcab']}", f"=sbc_ab*sbc_nc*PL!$K{R['sbc']}/PL!$K{R['rev']}*{c}{C_['rev']}", "#,##0.0")
            if y <= 2032:
                f_(cf, f"{c}{C_['gr']}", f"={c}{C_['rev']}/{pc}{C_['rev']}-1", "0.0%")
            f_(cf, f"{c}{C_['tax']}", "=tax", "0.0%")
            f_(cf, f"{c}{C_['nopat']}", f"=IF({c}{C_['ebit']}>0,{c}{C_['ebit']}*(1-{c}{C_['tax']}),{c}{C_['ebit']})", "#,##0.0")
            f_(cf, f"{c}{C_['da']}", ("=da_k*(h_cxt26)" if y == 2027 else f"=da_k*{pc}{C_['cx']}"), "#,##0.0")
            f_(cf, f"{c}{C_['wc']}", f"=wc_in*({c}{C_['rev']}-{pc}{C_['rev']})", "#,##0.0")
            f_(cf, f"{c}{C_['fcff']}", f"={c}{C_['nopat']}+{c}{C_['da']}-{c}{C_['cx']}+{c}{C_['wc']}+{c}{C_['sbcab']}", "#,##0.0")
            if y <= 2032:
                f_(cf, f"{c}{C_['int']}", f"=PL!{c}{R['oth']}*(1-IF(PL!{c}{R['pt']}>0,tax,0))", "#,##0.0")
                f_(cf, f"{c}{C_['fcf']}", f"={c}{C_['fcff']}+(1-sbc_ab)*sbc_nc*PL!{c}{R['sbc']}+{c}{C_['int']}", "#,##0.0")
                f_(cf, f"{c}{C_['cash']}", f"={pc}{C_['cash']}+{c}{C_['fcf']}", "#,##0.0")
            if y == 2027:
                f_(cf, f"{c}{C_['t']}", "=((DATE(2027,1,31)-val_date)/365)/2", "0.000")
                f_(cf, f"{c}{C_['cf']}", f"={c}{C_['fcff']}-(h1_fcf_co-(1-sbc_ab)*h1_sbc)", "#,##0.0")
            else:
                f_(cf, f"{c}{C_['t']}", f"=(DATE(2027,1,31)-val_date)/365+mid+{y - 2028}", "0.000")
                f_(cf, f"{c}{C_['cf']}", f"={c}{C_['fcff']}", "#,##0.0")
            f_(cf, f"{c}{C_['df']}", f"=1/(1+wacc)^{c}{C_['t']}", "0.0000")
            f_(cf, f"{c}{C_['pv']}", f"={c}{C_['cf']}*{c}{C_['df']}", "#,##0.0")
        E_ = lambda k: f"E{C_[k]}"
        f_(cf, E_("spv"), f"=SUM(F{C_['pv']}:P{C_['pv']})", "#,##0.0")
        f_(cf, E_("n33"), f"=P{C_['nopat']}*(1+g_tv)", "#,##0.0")
        f_(cf, E_("tv"), f"=IF(tv_alt=1,P{C_['fcff']}*(1+g_tv)/(wacc-g_tv),{E_('n33')}*(1-g_tv/ronic)/(wacc-g_tv))", "#,##0.0")
        f_(cf, E_("tvdf"), f"=1/(1+wacc)^((DATE(2027,1,31)-val_date)/365+10)", "0.0000")
        f_(cf, E_("tvpv"), f"={E_('tv')}*{E_('tvdf')}", "#,##0.0")
        f_(cf, E_("ev"), f"={E_('spv')}+{E_('tvpv')}", "#,##0.0")
        f_(cf, E_("nc"), "=cash_0-debt_0-ev_delta", "#,##0.0")
        f_(cf, E_("eq"), f"={E_('ev')}+{E_('nc')}", "#,##0.0")
        f_(cf, E_("ps0"), f"={E_('eq')}/sh_dil", "0.00")
        f_(cf, E_("ps1"), f"={E_('ps0')}*(1+ke)^t_tgt", "0.00")
        f_(cf, E_("tvsh"), f"={E_('tvpv')}/{E_('ev')}", "0.0%")
        f_(cf, E_("gate"), f"=SUM(E{C_['fcf']}:I{C_['fcf']})", "#,##0.0")
        f_(cf, E_("mincash"), f"=MIN(E{C_['cash']}:K{C_['cash']})", "#,##0.0")
    cf.column_dimensions["A"].width = 60
    P.column_dimensions["A"].width = 52
    # ───────── Val ─────────
    V = ws_["Val"]
    # WACC 區（具名）
    vw = {"rf": 30, "beta": 31, "erp": 32, "ke": 33, "kd": 34, "tax": 35, "mc": 36, "debt": 37, "wd": 38, "wacc": 39, "g": 40, "ronic": 41}
    f_(V, "A29", "折現率與資本結構（個股 DCF 標準規格第 2 節）", bold=True)
    rows_w = [("rf", "無風險利率 rf", "=rf", "0.00%"), ("beta", "β（Blume 調整後兩來源中位數）", "=beta", "0.000"), ("erp", "股權風險溢酬 ERP", "=erp", "0.00%"),
              ("ke", "股權成本 ke＝rf＋β × ERP", "=rf+beta*erp", "0.00%"), ("kd", "稅後債務成本＝(rf＋利差) ×(1 − 稅率)", "=kd*(1-tax)", "0.00%"),
              ("tax", "稅率", "=tax", "0.0%"), ("mc", "股權市值（現價 × 稀釋股數，US$M）", "=px*sh_dil", "#,##0"), ("debt", "債務（長期負債帳面，US$M）", "=debt_0", "#,##0"),
              ("wd", "債務權重 D÷(D＋E)", f"=B{vw['debt']}/(B{vw['debt']}+B{vw['mc']})", "0.00%"),
              ("wacc", "WACC（含敏感度調整）", f"=(1-B{vw['wd']})*B{vw['ke']}+B{vw['wd']}*B{vw['kd']}+wacc_adj", "0.00%"),
              ("g", "永續成長率 g", "=g_tv", "0.0%"), ("ronic", "永續期 RONIC＝WACC＋差", f"=B{vw['wacc']}+ronic_sp", "0.00%")]
    for k, lab, fm, nf in rows_w:
        f_(V, f"A{vw[k]}", lab)
        f_(V, f"B{vw[k]}", fm, nf)
    addname(wb, "ke", f"Val!$B${vw['ke']}")
    addname(wb, "wacc", f"Val!$B${vw['wacc']}")
    addname(wb, "ronic", f"Val!$B${vw['ronic']}")
    # 評價區
    f_(V, "A1", tf("Val — 目標價（股價基準日＋12 個月＝", ("tp_date", "yyyy-mm-dd"), "）＝DCF ", ("w_dcf", "0%"), "＋Comparable ", ("1-w_dcf", "0%"), "（情境 A／B／C 依權重加權）"))
    f_(V, "A2", "DCF：FCFF（NOPAT＋折舊攤銷 − 資本支出＋營運資金流入；股票薪酬不加回）以 WACC 折現到評價日（期中慣例），終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC)÷(WACC − g)，再以 ke 推到目標時點。")
    f_(V, "A3", "Comparable：同業 EV／營收（TTM）中位數 × 相對倍數係數 × CRWD 目標時點 TTM 營收（FY2027 與 FY2028 營收 1：3 加權）＋淨現金；少於約 8 家同業，不做迴歸；單一分部，SOTP 退化為單一營收倍數。")
    f_(V, "A4", "0.3 結論", bold=True)
    for j, h in enumerate(["項目", "數值", "對照", "對照", "說明"], 1):
        f_(V, f"{openpyxl.utils.get_column_letter(j)}5", h, bold=True, fill=HDR)
    cmp_rows = {"A": 60, "B": 61, "C": 62}
    f_(V, "A6", "DCF 目標價（情境加權）")
    f_(V, "B6", "=" + "+".join(f"tp_w{s}*CF!E{CB[s]['ps1']}" for s in SC), "0.00")
    f_(V, "A7", "Comparable 目標價（同業 EV／營收法；單一分部）")
    f_(V, "B7", f"=tp_wA*B{cmp_rows['A']}+tp_wB*B{cmp_rows['B']}+tp_wC*B{cmp_rows['C']}".replace("B60", "B64").replace("B61", "B64").replace("B62", "B64"), "0.00")
    # Comparable 表（列 64–70；欄 B、C、D＝情境 A、B、C）
    f_(V, "A8", "目標價（DCF 與 Comparable 加權）")
    f_(V, "B8", "=w_dcf*B6+(1-w_dcf)*B7", "0.00")
    f_(V, "E8", tf("DCF 權重 ", ("w_dcf", "0%"), "、Comparable ", ("1-w_dcf", "0%")))
    f_(V, "A9", "現價（拆股後）")
    f_(V, "B9", "=px", "0.00")
    f_(V, "E9", tf("股價基準日 ", ("px_date", "yyyy-mm-dd"), "（StockAnalysis 收盤；只有一個資料商來源）"))
    f_(V, "A10", "賣方平均目標價（只作對照）")
    f_(V, "B10", "=cons_tp", "0.00")
    f_(V, "E10", "StockAnalysis（S&P Global）53 位，2026-11-03 查看；只作對照")
    f_(V, "A11", "上檔空間：加權目標價")
    f_(V, "B11", "=B8/px-1", "+0.0%;-0.0%")
    f_(V, "A12", "上檔空間：Comparable 目標價")
    f_(V, "B12", "=B7/px-1", "+0.0%;-0.0%")
    f_(V, "A13", "上檔空間級距")
    f_(V, "B13", '=IF(B11>0.5,"大於 50%",IF(B11>0.25,"25%–50%",IF(B11>=0,"0%–25%",IF(B11>=-0.25,"0%～−25%",IF(B11>=-0.5,"−25%～−50%","小於 −50%")))))')
    f_(V, "A14", "投資建議")
    f_(V, "B14", '=IF(B16<0,IF(B11>0.25,"中立",IF(B11>=-0.25,"中立",IF(B11>=-0.5,"放空","強力放空"))),IF(B11>0.5,"強力買進",IF(B11>0.25,"買進",IF(B11>=-0.25,"中立",IF(B11>=-0.5,"放空","強力放空")))))')
    f_(V, "E14", "五級固定規則（原則第 24 條）：> +25% 買進、> +50% 強力買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空；收支關卡未過時不得買進（原則第 63 條）。")
    f_(V, "A15", "期間：折現期間（評價日起）／持有期間（股價基準日起）（年）")
    f_(V, "B15", "=t_tgt", "0.00")
    f_(V, "C15", "=t_hold", "0.00")
    f_(V, "E15", tf("評價日 ", ("val_date", "yyyy-mm-dd"), " → 目標時點 ", ("tp_date", "yyyy-mm-dd"), "；股價基準日 ", ("px_date", "yyyy-mm-dd"), "；上檔空間不是年化報酬"))
    f_(V, "A16", "收支關卡：情境 C 的 FY2026–FY2030 自由現金流合計（US$M；正值＝通過）")
    f_(V, "B16", f"=MIN(CF!E{CB['A']['gate']},CF!E{CB['B']['gate']},CF!E{CB['C']['gate']})", "#,##0")
    f_(V, "E16", tf("三情境最小值；", ("IF(B16<0,\"未通過\",\"通過\")", "@"), "；預測期最低年末現金（情境 C）US$M ", (f"CF!E{CB['C']['mincash']}", "#,##0")))
    f_(V, "A17", "DCF 單法上檔空間")
    f_(V, "B17", "=B6/px-1", "+0.0%;-0.0%")
    f_(V, "A18", "DCF 終值占 EV（情境 B；超過約 80% 時標示）")
    f_(V, "B18", f"=CF!E{CB['B']['tvsh']}", "0.0%")
    f_(V, "E18", '=IF(B18>0.8,"超過 80%：估值對永續成長與 WACC 高度敏感","未超過 80%")')
    f_(V, "A19", "現價隱含 EV／TTM 營收 ÷ 同業中位數（現行相對倍數）")
    f_(V, "B19", "=((px*sh_dil-(cash_0-debt_0-ev_delta))/Q!M7)/pm_med", "0.00")
    f_(V, "E19", tf("現價隱含 EV／TTM 營收 ", ("(px*sh_dil-(cash_0-debt_0-ev_delta))/Q!M7", "0.0"), "；同業中位數 ", ("pm_med", "0.0"), "；相對倍數係數基準 ", ("rel_mult", "0.00"), "（最近 5 個會計年度末 CRWD 倍數 ÷ 同業中位數的平均）"))
    f_(V, "A20", "上檔空間三段歸因：DCF 貢獻（權重 × DCF 上檔）")
    f_(V, "B20", "=w_dcf*(B6/px-1)", "+0.0%;-0.0%")
    f_(V, "A21", "上檔空間三段歸因：未調整 Comparable 貢獻（係數＝1）")
    f_(V, "B21", f"=(1-w_dcf)*((B7/rel_mult)/px-1)", "+0.0%;-0.0%")
    f_(V, "A22", "上檔空間三段歸因：相對倍數調整貢獻")
    f_(V, "B22", "=B11-B20-B21", "+0.0%;-0.0%")
    f_(V, "E22", "未調整 Comparable 的相對倍數係數＝1；基準係數為最近 5 年相對倍數平均（In 頁 rel_mult），此項即其調整貢獻（Sens 有係數低／高案）")
    f_(V, "A23", "重點：Comparable 價格＝EV 倍數 × 營收，隨 TTM 營收與同業倍數連動；DCF 價格依 WACC、終值與 SBC 費用化")
    f_(V, "A24", "加權中 DCF 與 Comparable 各占的目標價美元（DCF／Comparable）")
    f_(V, "B24", "=w_dcf*B6", "0.00")
    f_(V, "C24", "=(1-w_dcf)*B7", "0.00")
    f_(V, "A25", "估值評等（只看上檔空間）")
    f_(V, "B25", '=IF(B11>0.5,"強力買進",IF(B11>0.25,"買進",IF(B11>=-0.25,"中立",IF(B11>=-0.5,"放空","強力放空"))))', bold=True)
    f_(V, "A26", "入不敷出判定用（收支關卡：FY2026–FY2030 自由現金流合計，與 CF 頁第一個收支關卡列（情境 A）相同，US$M）")
    f_(V, "B26", f"=CF!E{CB['A']['gate']}", "#,##0")
    f_(V, "E26", "原則第 63 條：收支不足時投資建議不得為買進以上；B14 以三情境最小值判定（較保守，見 B16）；本列取情境 A 值，與 check_rating_gate 讀取 CF 頁第一個收支關卡列的口徑一致")
    f_(V, "A27", "入不敷出判定")
    f_(V, "B27", '=IF(B16<0,"不通過","通過")')
    # DCF 彙總表（列 44–54）
    f_(V, "A44", "DCF 彙總（三情境）", bold=True)
    for j, h in enumerate(["項目", "情境 A", "情境 B", "情境 C"], 1):
        f_(V, f"{openpyxl.utils.get_column_letter(j)}45", h, bold=True, fill=HDR)
    dsum = [("明確期現值合計（US$M）", "spv", "#,##0"), ("終值現值（US$M）", "tvpv", "#,##0"), ("企業價值 EV（US$M）", "ev", "#,##0"), ("淨現金（US$M）", "nc", "#,##0"),
            ("股權價值（US$M）", "eq", "#,##0"), ("每股價值：評價日（US$）", "ps0", "0.00"), ("每股價值：目標時點（US$）", "ps1", "0.00"), ("終值占 EV", "tvsh", "0.0%")]
    for i, (lab, k, nf) in enumerate(dsum):
        rr = 46 + i
        f_(V, f"A{rr}", lab)
        for c, s in zip("BCD", SC):
            f_(V, f"{c}{rr}", f"=CF!E{CB[s][k]}", nf)
    # Comparable 表（列 64–70）
    f_(V, "A58", "Comparable（情境 A／B／C；同業 EV／營收中位數 × 相對倍數係數 × 目標時點 TTM 營收）", bold=True)
    for j, h in enumerate(["項目", "情境 A", "情境 B", "情境 C"], 1):
        f_(V, f"{openpyxl.utils.get_column_letter(j)}59", h, bold=True, fill=HDR)
    cr = [("FY2027E 營收（US$M）", lambda s: f"=PL!F{PB[s]['rev']}", "#,##0"), ("FY2028E 營收（US$M）", lambda s: f"=PL!G{PB[s]['rev']}", "#,##0"),
          ("目標時點 TTM 營收（0.25 × FY2027 ＋ 0.75 × FY2028；Q3 FY2028 TTM 近似，US$M）", lambda s: None, "#,##0"),
          ("企業價值＝倍數 × 係數 × TTM 營收（US$M）", lambda s: None, "#,##0"), ("每股價值（目標時點，US$）", lambda s: None, "0.00")]
    for i, (lab, fn, nf) in enumerate(cr):
        f_(V, f"A{60 + i}", lab)
    for c, s in zip("BCD", SC):
        f_(V, f"{c}60", f"=PL!F{PB[s]['rev']}", "#,##0")
        f_(V, f"{c}61", f"=PL!G{PB[s]['rev']}", "#,##0")
        f_(V, f"{c}62", f"=0.25*{c}60+0.75*{c}61", "#,##0")
        f_(V, f"{c}63", f"=pm_med*rel_mult*{c}62", "#,##0")
        f_(V, f"{c}64", f"=({c}63+cash_0-debt_0-ev_delta)/sh_dil", "0.00")
    f_(V, "B7", "=tp_wA*B64+tp_wB*C64+tp_wC*D64", "0.00")
    V.column_dimensions["A"].width = 70
    V.column_dimensions["E"].width = 80
    # ───────── RevDCF ─────────
    RV = ws_["RevDCF"]
    f_(RV, "A1", "RevDCF — 反向 DCF：現價隱含什麼（封閉解，不依賴迭代；情境 B 路徑）", bold=True)
    f_(RV, "A2", "以公式反解：把現價折回評價日，扣除淨現金與首期現值，除以 FY2028 起各年 FCFF（含延伸期）與終值的現值，得到使 DCF＝現價所需的 FCFF 倍數 k（基準＝1）。")
    rb = CB["B"]
    rows_rv = [
        (6, "現價（拆股後，US$）", "=px", "0.00"),
        (7, "現價折回評價日（以 ke 折現，US$）", "=px/(1+ke)^((px_date-val_date)/365)", "0.00"),
        (8, "隱含股權價值（US$M）", "=B7*sh_dil", "#,##0"),
        (9, "淨現金（US$M）", "=cash_0-debt_0-ev_delta", "#,##0"),
        (10, "隱含企業價值 EV（US$M）", "=B8-B9", "#,##0"),
        (11, "情境 B 首期（FY2027 下半年）現值（US$M）", f"=CF!F{rb['pv']}", "#,##0"),
        (12, "情境 B FY2028–FY2037 現金流現值＋終值現值（US$M）", f"=SUM(CF!G{rb['pv']}:P{rb['pv']})+CF!E{rb['tvpv']}", "#,##0"),
        (13, "反解係數 k：FY2028 起 FCFF 與終值須為模型的幾倍才支持現價", "=(B10-B11)/B12", "0.00"),
        (14, "模型基準（k＝1）", 1, "0.00"),
        (16, "現價隱含 EV／TTM 營收（Q 頁 TTM）", "=(px*sh_dil-(cash_0-debt_0-ev_delta))/Q!M7", "0.0"),
        (17, "同業 EV／營收中位數", "=pm_med", "0.0"),
        (18, "現價隱含相對倍數（÷ 同業中位數）", "=B16/B17", "0.00"),
        (19, "模型相對倍數係數（基準）", "=rel_mult", "0.00"),
        (20, "情境 B 目標時點 TTM 營收（US$M）", "=Val!C62", "#,##0"),
        (21, "支持現價所需的 EV／TTM 營收（以目標時點 TTM 營收、不含折現）", "=(px*sh_dil-(cash_0-debt_0-ev_delta))/B20", "0.0"),
    ]
    for r_, lab, fm, nf in rows_rv:
        f_(RV, f"A{r_}", lab)
        f_(RV, f"B{r_}", fm, nf)
    f_(RV, "A24", tf("現金流折現法：現價要成立，FY2028 起 FCFF 與終值須為情境 B 基準的 ", ("raw", 'IF(B13<0,"（情境 B 基準的 FY2028 起 FCFF 與終值現值合計為負，倍數無經濟意義；改看下方整本重算反解）",TEXT(B13,"0.0")&" 倍")'), "（反解值是目標價＝現價的點，不是評等門檻）。"))
    f_(RV, "A25", tf("同業倍數法：現價隱含 EV／TTM 營收 ", ("B16", "0.0"), "，為同業中位數 ", ("B17", "0.0"), " 的 ", ("B18", "0.0"), " 倍；模型相對倍數係數基準為 ", ("B19", "0.00"), "（最近 5 個會計年度末 CRWD 倍數 ÷ 同業中位數的平均）。"))
    rev = {}
    rf_ = os.path.join(DATA, "reverse_v01.json")
    if os.path.exists(rf_):
        rev = json.load(open(rf_, encoding="utf-8"))
    f_(RV, "A27", "整本重算反解（G2-9；LibreOffice 全活頁簿重算，在其餘輸入不變下，使加權目標價＝現價所需的輸入值；靜態值，不隨 In 頁連動）", bold=True)
    for j, h in enumerate(["輸入", "基準值", "反解值", "是否在 In 頁敏感度範圍內", "說明"], 1):
        f_(RV, f"{openpyxl.utils.get_column_letter(j)}28", h, bold=True, fill=HDR)
    rr_ = 29
    for it in rev.get("items", []):
        f_(RV, f"A{rr_}", it["label"])
        f_(RV, f"B{rr_}", it["base"], "0.000")
        f_(RV, f"C{rr_}", it["solve"] if it["solve"] is not None else "未找到可行解", "0.000")
        f_(RV, f"D{rr_}", it["within"])
        f_(RV, f"E{rr_}", it["note"])
        rr_ += 1
    if not rev:
        f_(RV, "A29", "尚未執行 --reverse")
    rr_ += 1
    sensrow = {(c["key"], c["side"]): 4 + i for i, c in enumerate((sens or {}).get("cases", []))}
    f_(RV, f"A{rr_}", "DCF 與 Comparable 落差公式分解（G2-8；每股，US$）", bold=True)
    g0 = rr_ + 1
    f_(RV, f"A{g0}", "落差＝Comparable 目標價 − DCF 目標價")
    f_(RV, f"B{g0}", "=Val!B7-Val!B6", "0.00")
    f_(RV, f"A{g0+1}", "(1) 股票薪酬：DCF 加回 SBC 時的 DCF 增加量（Sens 重算值，連結 Sens 頁）")
    if ("sbc_ab", "高") in sensrow:
        f_(RV, f"B{g0+1}", f"=Sens!H{sensrow[('sbc_ab', '高')]}-Val!B6", "0.00")
    f_(RV, f"A{g0+2}", "(2) 槓桿：兩法同用淨現金／稀釋股數（cash_0−debt_0；sh_dil），差額")
    f_(RV, f"B{g0+2}", "=(cash_0-debt_0-ev_delta)/sh_dil-(cash_0-debt_0-ev_delta)/sh_dil", "0.00")
    f_(RV, f"A{g0+3}", "(3) 時點：兩法皆為目標時點每股價值（DCF 由評價日值複利，Comparable 用目標日 TTM），差額")
    f_(RV, f"B{g0+3}", "=0", "0.00")
    f_(RV, f"A{g0+4}", "(4) 倍數與成長（殘差）：落差 − (1) − (2) − (3)")
    f_(RV, f"B{g0+4}", f"=B{g0}-B{g0+1}-B{g0+2}-B{g0+3}", "0.00")
    f_(RV, f"A{g0+5}", tf("DCF 隱含 EV／目標日 TTM 營收 ", ("(Val!B6*sh_dil-(cash_0-debt_0-ev_delta))/B20", "0.0"), "；同業中位數 × 係數 ", ("pm_med*rel_mult", "0.0"), "；殘差主要反映同業六家的成長與利益率低於本模型情境之處，或本模型 FCFF 低於同業倍數隱含之現金流。"))
    RV.column_dimensions["A"].width = 80
    # ───────── Recon ─────────
    RC = ws_["Recon"]
    f_(RC, "A1", "Recon — 對帳：公司指引、賣方共識與模型（情境 B；FY2027＝2026-02-01～2027-01-31）", bold=True)
    f_(RC, "A2", "公司指引與共識為 Interested-party，只在本頁對照，不進基準（原則第 10 節第 59 條）。口徑：營收＝GAAP 合併營收；non-GAAP 為公司自訂口徑。")
    for j, h in enumerate(["項目", "模型（情境 B）", "指引或共識", "差距（模型相對對照）", "說明", "來源"], 1):
        f_(RC, f"{openpyxl.utils.get_column_letter(j)}4", h, bold=True, fill=HDR)
    pb = PB["B"]
    tb = TB["B"]
    rec = [
        (5, "FY2027 營收（US$M）：模型 vs 公司指引中點", f"=PL!F{pb['rev']}", "=g_rev27", "#,##0.0", "公司指引 5,991.1–6,011.1（Q2 FY2027 新聞稿）"),
        (6, "FY2027 期末 ARR（US$M）：模型 vs 公司指引中點", f"=Thesis!F{tb['ARR']}", "=g_arr27", "#,##0.0", "公司指引 6,603.0–6,611.9"),
        (7, "FY2027 non-GAAP 營業利益（US$M）：模型 vs 公司指引中點", f"=PL!F{pb['ngop']}", "=g_ngop27", "#,##0.0", "公司指引 1,497.2–1,508.4"),
        (8, "FY2027 non-GAAP 稀釋 EPS（US$，拆股後）：模型 vs 公司指引中點", f"=PL!F{pb['ngeps']}", "=g_eps27", "0.00", "公司指引 1.25–1.26"),
        (9, "FY2027 營收（US$M）：模型 vs 賣方共識", f"=PL!F{pb['rev']}", "=c_rev27", "#,##0.0", "StockAnalysis（S&P Global，49 位，2026-11-02）"),
        (10, "FY2027 non-GAAP EPS（US$）：模型 vs 賣方共識", f"=PL!F{pb['ngeps']}", "=c_eps27", "0.00", "同上"),
        (11, "FY2028 營收（US$M）：模型 vs 賣方『次一年』（年度待確認）", f"=PL!G{pb['rev']}", "=c_rev28", "#,##0.0", "同上；明細表付費鎖定，標待確認"),
        (12, "目標價（US$）：模型加權 vs 賣方平均", "=Val!B8", "=cons_tp", "0.00", "53 位；只作對照"),
    ]
    for r_, lab, fm, ref, nf, src in rec:
        f_(RC, f"A{r_}", lab)
        f_(RC, f"B{r_}", fm, nf)
        f_(RC, f"C{r_}", ref, nf)
        f_(RC, f"D{r_}", f"=B{r_}/C{r_}-1", "+0.0%;-0.0%")
        f_(RC, f"E{r_}", tf(('IF(D' + str(r_) + '>0,"高於","低於")', "@"), "對照 ", (f"ABS(D{r_})*100", "0.0"), " 個百分點"))
        f_(RC, f"F{r_}", src)
    f_(RC, "A14", "指引差距分解與反解（FY2027 期末 ARR；Recon 反解值不借用指引校準）", bold=True)
    f_(RC, "A15", "模型 FY2027 人類端受保護單位年增（基準）")
    f_(RC, "B15", f"=Thesis!F{tb['gu']}", "0.0%")
    f_(RC, "A16", "達到公司指引 ARR 所需的 FY2027 人類端單位年增（其餘驅動不變，公式反解）")
    f_(RC, "B16", f"=(g_arr27-Thesis!F{tb['aARR']})/(Thesis!E{tb['U']}*Thesis!F{tb['M']}*Thesis!F{tb['P']})-1", "0.0%")
    f_(RC, "E16", tf("基準 ARR ", (f"IF(Thesis!F{tb['ARR']}<g_arr27,\"低於\",\"高於\")", "@"), "指引 ", (f"ABS(Thesis!F{tb['ARR']}/g_arr27-1)*100", "0.0"), " 個百分點；差距落在單位年增（Assumed）一項"))
    f_(RC, "A17", "FY2027 上半年實績淨新增 ARR（US$M，公司口徑）／全年指引隱含下半年（US$M）")
    f_(RC, "B17", "=Q!D15+Q!E15", "#,##0.0")
    f_(RC, "C17", "=g_arr27-h_arr26-B17", "#,##0.0")
    f_(RC, "E17", "下半年隱含淨新增 ARR＝指引期末 ARR − FY2026 期末 ARR − 上半年實績；只作對照")
    f_(RC, "A19", "AI 淨效果（情境 A／B／C，FY2032E，US$M）", bold=True)
    for j, h in enumerate(["項目", "情境 A", "情境 B", "情境 C"], 1):
        f_(RC, f"{openpyxl.utils.get_column_letter(j)}20", h, bold=True, fill=HDR)
    for i, (lab, k, nf) in enumerate([("AI 增量營收（AI 代理防護）", "aiRev", "#,##0.0"), ("定價侵蝕（負值）", "ero", "#,##0.0"), ("增量推論成本", "inc", "#,##0.0"),
                                      ("AI 淨效果", "net", "#,##0.0"), ("損益兩平價格保留率", "be", "0.00")]):
        f_(RC, f"A{21 + i}", lab)
        for c, s in zip("BCD", SC):
            f_(RC, f"{c}{21 + i}", f"=Thesis!K{TB[s][k]}", nf)
    f_(RC, "A26", "情境分化依據：三情境相同的人類端單位與模組路徑；分化來自 (a) AI 代理單位數與價格保留率（AI 增量營收）、(b) 每模組年費年變動（定價侵蝕）、(c) 任務數（推論成本）。")
    f_(RC, "A28", "未找到清單（Recon 同步列出）", bold=True)
    nf_items = [("AI 專屬營收與 AI 產品付費客戶數", "未找到：公司未揭露", "研究包 (5)"),
                ("Charlotte AI 用量與 ARR；AIDR 絕對 ARR", "未找到：公司只給 AIDR ARR 季增倍數", "Q2 FY2027 法說"),
                ("受保護端點／身分總數", "未找到：公司不揭露；In 頁 u_0 為 Assumed", "10-K FY2026"),
                ("FY2028 與 FY2029 賣方共識明細", "未找到：頁面付費鎖定", "StockAnalysis"),
                ("RSU 與選擇權庫藏股法淨股數明細", "未找到：以稀釋加權平均近似", "10-Q 權益獎酬附註"),
                ("已付所得稅 ÷ non-GAAP 稅前（現金稅率）", "未找到：未建現金稅橋", "10-K 現金流量表補充資訊")]
    for i, (a, b, c) in enumerate(nf_items):
        f_(RC, f"A{29 + i}", a)
        f_(RC, f"C{29 + i}", b)
        f_(RC, f"F{29 + i}", c)
    RC.column_dimensions["A"].width = 70
    RC.column_dimensions["E"].width = 50
    # ───────── README ─────────
    RD = ws_["README"]
    inhash = hashlib.sha256(json.dumps([[str(x) for x in it] for it in INP] + [sorted((k, str(v)) for k, v in ov.items())] + [VERSION], ensure_ascii=False).encode()).hexdigest()[:10]
    build_id = f"crowdstrike-{VERSION}-{inhash}"
    f_(RD, "A1", tf(f"CrowdStrike（CRWD）個股模型 {VERSION}｜build_id {build_id}｜資料截止日 ", ("data_date", "yyyy-mm-dd"), f"｜成品日 {BUILD_DATE}｜股價基準日 ", ("px_date", "yyyy-mm-dd"), "｜評價日 ", ("val_date", "yyyy-mm-dd"), "｜本版未做先驗"))
    f_(RD, "A2", f"{VERSION} 要點：依工作單 G3 修正 v0.1：基準不再全期持平（改為依 FY2023–FY2026 回歸的固定＋變動成本規模經濟，全期持平移到 Sens 下緣）、折舊攤銷由前一年資本支出推得、相對倍數係數改用最近 5 個會計年度末自算平均、新增合理性關卡診斷與事件尾端 Sens（DOJ／SEC、Delta 案）。其餘依 software_ai 工作單 G 節建立：命題模組＝原型 P（按防護範圍計價）單位經濟，三情境（AI 代理防護放量／基準／價格侵蝕）；官方分部為單一分部，營收線＝訂閱、專業服務；FCFF DCF 到 FY2032＋同業 EV／營收法各半；推論成本只讀 Tokenomics v22.0 快照。本版未做先驗（Andy 2026-11-16 指示先跳過）。")
    f_(RD, "A3", tf("命題：AI 進入資安後，CrowdStrike 的 ARR 由『受保護單位 × 每單位模組數 × 每模組年費』驅動；AI 代理（非人類身分）創造新的受保護單位與模組（增量），AI 平台業者與打包競爭可能侵蝕每模組年費（定價侵蝕），推論成本進入毛利。現價要成立，FY2028 起現金流須為本模型基準的 ", ("RevDCF!B13", "0.0"), " 倍。"))
    f_(RD, "A4", tf("核心問題：FY2032 的 AI 淨效果（AI 代理防護增量營收 + 定價侵蝕 − 增量推論成本）在情境 A／B／C 各為 US$M ", (f"Thesis!K{TB['A']['net']}", "#,##0"), "／", (f"Thesis!K{TB['B']['net']}", "#,##0"), "／", (f"Thesis!K{TB['C']['net']}", "#,##0"), "；FY2032 營收情境 B 為 US$M ", (f"PL!K{PB['B']['rev']}", "#,##0"), "。"))
    f_(RD, "A5", tf("投資建議：", ("Val!B14", "@"), "——在『AI 代理防護以情境權重放量、每模組價格依情境小幅下降、股票薪酬依規格費用化（不加回）』的基準假設下；加權目標價 US$", ("Val!B8", "0.00"), "，上檔空間 ", ("Val!B11", "+0.0%;-0.0%"), "（現價 US$", ("px", "0.00"), "，", ("px_date", "yyyy-mm-dd"), "）。主要失效條件見下方。"))
    f_(RD, "A6", tf("重點觀察 1：目標價 ", ("Val!B8", "0.00"), "＝DCF ", ("Val!B6", "0.00"), " × ", ("w_dcf", "0%"), "＋Comparable ", ("Val!B7", "0.00"), " × ", ("1-w_dcf", "0%"), "；現價隱含 EV／TTM 營收為同業中位數的 ", ("Val!B19", "0.0"), " 倍（相對倍數係數基準 ", ("rel_mult", "0.00"), "，最近 5 年平均）。"))
    f_(RD, "A7", tf("重點觀察 2：FY2032 情境 B 的 ARR 為 US$M ", (f"Thesis!K{TB['B']['ARR']}", "#,##0"), "，其中 AI 代理防護 ", (f"Thesis!K{TB['B']['aARR']}", "#,##0"), "；增量推論成本 US$M ", (f"Thesis!K{TB['B']['inc']}", "#,##0"), "，占 non-GAAP 營業成本 ", (f"Thesis!K{TB['B']['share']}", "0.0%"), "，不是主要風險；價格保留率與代理單位數才是。"))
    f_(RD, "A8", tf("重點觀察 3：收支關卡（FY2026–FY2030 自由現金流合計，情境 C）US$M ", ("Val!B16", "#,##0"), "，", ("IF(Val!B16<0,\"未通過\",\"通過\")", "@"), "；DCF 終值占 EV ", ("Val!B18", "0%"), "；情境 B 的 FY2032 GAAP 營業利益率 ", (f"PL!K{PB['B']['ebitm']}", "0.0%"), "，SBC 費用化是 DCF 與 non-GAAP 口徑差距的主因。"))
    f_(RD, "A9", "投資建議依加權目標價判定；現價隱含的現金流遠高於本模型基準，主要差異在 AI 代理防護放量、股票薪酬費用化與終值。本版未做先驗。")
    f_(RD, "A10", "驅動 → 推導對照表（每頁輸入、推導量、因果方向）", bold=True)
    for j, h in enumerate(["頁", "輸入（In 頁鍵）", "推導量", "因果方向"], 1):
        f_(RD, f"{openpyxl.utils.get_column_letter(j)}11", h, bold=True)
    mp = [("Thesis", "u_0、m_0、gu_27／gu_32、gm_27／gm_32、ero_A／B／C", "核心 ARR、價格侵蝕", "單位與模組數上升 → ARR 上升；每模組年費下降 → ARR 下降"),
          ("Thesis", "r32_A／B／C、ma_n、pf_A／B／C", "AI 代理防護 ARR 與增量營收", "代理單位與價格保留率上升 → 增量營收上升"),
          ("Thesis", "tasks_0、tasks_A／B／C、tok_task、px_tok（Tokenomics 快照）", "AI 推論成本與兩平價格保留率", "任務數、token、單價上升 → 推論成本上升 → AI 淨效果下降"),
          ("PL", "conv、ps_pct、gmsub_ng、opx_27、fix_ng、sbc_27、fix_sbc、oth_pct、gm_gap", "營收、毛利、營業利益、EPS（GAAP 與 non-GAAP）", "費用率與 SBC 率下降 → 營業利益率上升"),
          ("CF", "da_k、cx_27、fix_cx、wc_in、tax", "FCFF、收支關卡、現金", "資本支出上升 → FCFF 下降；營運資金流入上升 → FCFF 上升"),
          ("Val", "rf、beta、erp、kd_sp、g_tv、ronic_sp、pm_*、rel_mult、w_dcf", "DCF、Comparable、加權目標價、投資建議", "WACC 上升 → DCF 下降；同業倍數與係數上升 → Comparable 上升")]
    for i, row in enumerate(mp):
        for j, v in enumerate(row):
            f_(RD, f"{openpyxl.utils.get_column_letter(j + 1)}{12 + i}", v)
    f_(RD, "A19", "版本要點", bold=True)
    f_(RD, "A20", "相對 v0.1：DCF 與 Comparable 目標價都上移（營業利益率不再凍結在虧損水準、相對倍數係數由 1.00 改為 5 年平均）；評等仍為強力放空，診斷見 README 頁合理性關卡一節；數值皆為本版輸入的函數。")
    f_(RD, "A21", "官方分部：單一應報導分部；營收線訂閱、專業服務；對帳核對列為 0（PL 頁各情境區塊、Hist 頁、Q 頁）。")
    f_(RD, "A22", "推論成本只讀 Tokenomics v22.0 快照（data/tk_snapshot_v22.0.json）；不自行估 token 價格。v0.2 改讀 software_ai/unit_economics。")
    f_(RD, "A23", "研究包：data/research_crowdstrike.md、facts_crowdstrike.csv、segments.json。")
    f_(RD, "A25", "市場資料日期、失效條件與可證偽指標", bold=True)
    f_(RD, "A26", tf("市場資料日期：現價 ", ("px_date", "yyyy-mm-dd"), "（StockAnalysis 收盤，單一資料商）；rf 2026-11-13（美國財政部）；賣方共識 2026-11-02 更新、2026-11-03 查看；同業倍數 2026-11-14／15 更新，SentinelOne 為 2026-10-26 收盤（逾期）；公司指引 2026-08-26（Q2 FY2027）。"))
    f_(RD, "A27", "下一個事件：Q3 FY2027 財報預計 2026 年 12 月初公布；共識與指引屆時更新，現有共識為 Q2 財報後，不屬『財報前』共識。")
    f_(RD, "A28", tf("承重 Assumed 1（人類端單位與模組路徑 gu、gm）。可證偽指標：Q3 FY2027 期末 ARR 對公司指引 ", ("g_arr27", "#,##0"), "（FY2027 期末）的軌跡；模型 FY2027 ARR 為 ", (f"Thesis!F{TB['B']['ARR']}", "#,##0"), "。"))
    f_(RD, "A29", "承重 Assumed 2（價格保留：ero、pf）。可證偽指標：公司揭露的美元淨留存率（FY2026 為 115%）、Falcon Flex 轉換 ARR 增幅（Q2 FY2027 逾 40%）、AIDR 絕對 ARR 與客戶數（目前未揭露）。")
    f_(RD, "A30", tf("承重 Assumed 3（相對倍數係數 rel_mult，基準為最近 5 年平均相對倍數，Derived）。費用率、SBC 率、資本支出率以近四季為起點，只依歷史回歸的固定成本吸收規模經濟（non-GAAP 營業利益率 ", ("Q!M17", "0.0%"), "、SBC 占營收 ", ("Q!M18", "0.0%"), "），無證據的改善只放 Sens；可證偽指標為這兩個比率的逐季路徑。"))
    f_(RD, "A31", "結論條件式：評等隨加權目標價相對現價的級距而定；DCF 與 Comparable 各半時，評等翻轉需要 (a) 現金流反解係數大幅下降，或 (b) 同業倍數與相對倍數係數上升（見 Recon 與 Sens）。")
    f_(RD, "A32", "失效條件（可觀察）：FY2027 以後淨新增 ARR 年增持續高於本模型基準；AIDR／代理防護揭露絕對 ARR 且高於情境 A；SBC 占營收快速下降；DOJ／SEC 資訊請求有結論。")
    f_(RD, "A33", "本版未做先驗；公開 repo，內容只含公司與市場事實。")
    f_(RD, "A35", "風險登錄", bold=True)
    for j, h in enumerate(["風險", "是否建模", "理由", "觀察指標"], 1):
        f_(RD, f"{openpyxl.utils.get_column_letter(j)}36", h, bold=True)
    risks = [("DOJ／SEC 就營收認列與 ARR 報告方式的資訊請求（10-Q Note 11）", "否", "無日期、無範圍量化；不能估機率與金額", "10-Q／8-K 披露；ARR 定義變更"),
             ("7 月 19 日事件相關訴訟（Delta 案在 Georgia 州法院證據開示中）", "否", "金額無揭露；僅事件成本已含在 GAAP 營業利益實績", "法院裁定、和解公告、10-Q 或有負債"),
             ("AWS 協議重談與雲端託管成本上升", "否", "成本變動只含在毛利率基期；10-K 未量化重談影響", "10-K 風險因子、雲端託管成本增額"),
             ("XM Cyber 技術資產併購（約 US$145M 現金與股份，未交割）", "否", "金額小，未含在公司 ARR 與營收指引；本模型亦不含", "交割公告、FY2027 下半年 ARR 貢獻"),
             ("AI 平台業者進入資安（例：Claude Code Security，2026-02-20）", "是（情境 C 的價格侵蝕與代理單位）", "以價格與單位路徑呈現，不單獨建競爭事件", "美元淨留存率、Flex 轉換 ARR 增幅、新客 ARR 占比")]
    for i, row in enumerate(risks):
        for j, v in enumerate(row):
            f_(RD, f"{openpyxl.utils.get_column_letter(j + 1)}{37 + i}", v)
    f_(RD, "A43", "命題回覆", bold=True)
    f_(RD, "A44", tf("1. AI 為 CrowdStrike 帶來增量還是侵蝕：FY2032 情境 A／B／C 的 AI 淨效果 US$M ", (f"Thesis!K{TB['A']['net']}", "#,##0"), "／", (f"Thesis!K{TB['B']['net']}", "#,##0"), "／", (f"Thesis!K{TB['C']['net']}", "#,##0"), "（AI 增量營收減去定價侵蝕與增量推論成本）。"))
    f_(RD, "A45", tf("2. 推論成本是否吃掉毛利：情境 B 的 FY2032 增量推論成本占 non-GAAP 營業成本 ", (f"Thesis!K{TB['B']['share']}", "0.0%"), "；以模型商全成本對照的下限為 US$M ", (f"Thesis!K{TB['B']['floor']}", "#,##0"), "（對照，不進模型）。"))
    f_(RD, "A46", tf("3. 現價隱含什麼：FY2028 起現金流須為情境 B 基準的 ", ("raw", 'IF(RevDCF!B13<0,"（基準現值合計為負，倍數無意義）",TEXT(RevDCF!B13,"0.0")&" 倍")'), "，或 EV／TTM 營收為同業中位數的 ", ("RevDCF!B18", "0.0"), " 倍。"))
    f_(RD, "A48", "不適用：算力 MW、Wattonomics 對照、Tokenomics 自用 AI 讀法 2（CrowdStrike 不自建算力出租；推論成本向模型商購買）。")
    f_(RD, "A49", "不適用：三層協調（Coord 頁、C0、兩讀數）：非 Siliconomics 覆蓋。")
    f_(RD, "A50", "未做：先驗（Andy 2026-11-16 指示先跳過）。")
    f_(RD, "A51", "未做：展示層數字一致性檢查設定（display_check.json，原則第 10 節第 51 條）；本版 README、HTML 的數字由儲存格組字，未另建對照檔。")
    f_(RD, "A52", "未做：RSU 與選擇權庫藏股法明細（以稀釋加權平均近似）；現金稅率橋（原則第 10 節第 40 條）；事件型尾端 Sens（DOJ／SEC、Delta 案；原則第 19 條）；同業成長與利益率並列（未取得）。")
    f_(RD, "A53", "v0.2 待辦：改讀 software_ai/unit_economics 的 P 原型；補 AIDR／Charlotte AI 揭露；以 10-K 核對 FY2022–FY2023 員工數。")
    sens_r = {(c["key"], c["side"]): 4 + i for i, c in enumerate((sens or {}).get("cases", []))}
    f_(RD, "A55", "合理性關卡診斷（工作單 G3-3；發布前必查）", bold=True)
    f_(RD, "A56", tf("(a) DCF 目標價 ", ("Val!B6", "0.00"), "＝現價的 ", ("Val!B6/px", "0.0%"), "：", ("raw", 'IF(OR(Val!B6<0.25*px,Val!B6>4*px),"觸發（低於 25% 或高於 4 倍）","未觸發")'), "；(b) 情境 B 的 DCF 企業價值 US$M ", (f"CF!E{CB['B']['ev']}", "#,##0"), "：", ("raw", f'IF(CF!E{CB["B"]["ev"]}<0,"為負，觸發","為正，未觸發")'), "；(c) 評等 ", ("Val!B14", "@"), "：", ("raw", 'IF(OR(Val!B14="強力放空",Val!B14="強力買進"),"觸發","未觸發")'), "；(d) FY2027 營收差距 ", ("Recon!D5", "+0.0%;-0.0%"), "、non-GAAP 營業利益差距 ", ("Recon!D7", "+0.0%;-0.0%"), "（門檻 10%）：", ("raw", 'IF(OR(ABS(Recon!D5)>0.1,ABS(Recon!D7)>0.1),"觸發","未觸發")'), "。"))
    f_(RD, "A57", tf("診斷量：FY2032 情境 B 的 GAAP 營業利益率 ", (f"PL!K{PB['B']['ebitm']}", "0.0%"), "、non-GAAP 營業利益率 ", (f"PL!K{PB['B']['ngopm']}", "0.0%"), "、SBC（含薪資稅）占營收 ", (f"PL!K{PB['B']['sbc']}/PL!K{PB['B']['rev']}", "0.0%"), "；DCF 加回 SBC 時的 DCF 目標價 ", ((f"Sens!H{sens_r[('sbc_ab', '高')]}" if ('sbc_ab', '高') in sens_r else "0"), "0.00"), "；SBC 規模經濟加大案（固定部分高值）的相對基準 ", ((f"Sens!E{sens_r[('fix_sbc', '高')]}" if ('fix_sbc', '高') in sens_r else "0"), "+0.0%;-0.0%"), "。"))
    f_(RD, "A58", tf("核對：(1) 營運資金為現金流入、符號正確（情境 B 的 FY2027 營運資金流入 US$M ", (f"CF!F{CB['B']['wc']}", "#,##0"), "）；(2) 模型 FY2027 公司口徑近似自由現金流 US$M ", (f"CF!F{CB['B']['fcf']}", "#,##0"), "，對照 FY2026 實績 ", ("h_fcf26", "#,##0"), "、FY2027 上半年實績 ", ("h1_fcf_co", "#,##0"), "；(3) 淨現金與股數取自 10-Q 與新聞稿（In 頁）；(4) 折舊攤銷由資本支出推得（In 頁 da_k），EV 為正。"))
    f_(RD, "A59", "診斷結論：DCF 偏低與強力放空不是公式或符號錯誤，而是輸入口徑的結果：GAAP 營業利益（SBC 費用化）在 FY2023–FY2026 實績中接近零或為負，歷史 SBC 占營收沒有下降趨勢，規模經濟只由固定成本回歸取得（幅度小），所以預測期 GAAP 營業利益率仍接近零（略為負），終值為負或近零。保留此結果的限制：它由『SBC 視為費用、無證據不假設其占比下降』這個口徑決定，不是對公司價值的經濟定論；DCF 加回 SBC、SBC 與費用的規模經濟加大案、延伸期利益率上移案在 Sens 各有列。Comparable 單法（同業倍數 × 最近 5 年相對倍數平均）才是本版加權目標價的主要來源。")
    RD.column_dimensions["A"].width = 120
    # ───────── Sens ─────────
    S = ws_["Sens"]
    f_(S, "A1", "Sens — 一次改一個輸入（In 頁敏感度低／高值），整本活頁簿以 LibreOffice 重算後的加權目標價（Val 頁）與 DCF、Comparable 單法目標價；由 build_crowdstrike.py 的敏感度程序產生。")
    f_(S, "A2", "相對基準＝案例目標價 ÷ 基準加權目標價 − 1。情境權重、WACC ±1pt、g、β 以外的來源 ERP、資本支出、終值替代法、股票薪酬加回皆列案。tok_task、px_tok 的低／高取自 Tokenomics 快照（任務 2／4；Luna／Sol 價）。情境專屬列（ero_*、r32_*、pf_*、tasks_*）以 A／B／C 三情境並列為其區間，不另列案。")
    for j, h in enumerate(["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "說明", "DCF 單法目標價", "Comparable 單法目標價"], 1):
        f_(S, f"{openpyxl.utils.get_column_letter(j)}3", h, bold=True, fill=HDR)
    cases = (sens or {}).get("cases", [])
    base_tp = (sens or {}).get("base", {})
    rr = 4
    case_row = {}
    for cs in cases:
        f_(S, f"A{rr}", cs["label"])
        f_(S, f"B{rr}", cs["setting"])
        f_(S, f"C{rr}", cs["side"])
        f_(S, f"D{rr}", cs["tp"], "0.00")
        f_(S, f"E{rr}", f"=D{rr}/Val!$B$8-1", "+0.0%;-0.0%")
        f_(S, f"F{rr}", cs["key"])
        f_(S, f"G{rr}", "重算值（整本重算）")
        f_(S, f"H{rr}", cs["dcf"], "0.00")
        f_(S, f"I{rr}", cs["cmp"], "0.00")
        case_row[(cs["key"], cs["side"])] = rr
        rr += 1
    rr += 1
    f_(S, f"A{rr}", "build_id")
    f_(S, f"D{rr}", build_id)
    bid_row = rr
    rr += 1
    f_(S, f"A{rr}", "模型狀態")
    f_(S, f"D{rr}", "CURRENT" if sens else "未跑敏感度")
    rr += 1
    f_(S, f"A{rr}", "註：In 頁輸入指紋（敏感度重算時的基準輸入雜湊）")
    f_(S, f"D{rr}", inhash)
    rr += 1
    f_(S, f"A{rr}", "註：基準加權目標價（敏感度重算時）／本檔 Val 頁")
    f_(S, f"D{rr}", base_tp.get("tp"), "0.00")
    f_(S, f"E{rr}", "=Val!B8", "0.00")
    f_(S, f"F{rr}", '=IF(ABS(D' + str(rr) + '-E' + str(rr) + ')<0.005,"一致","STALE")')
    rr += 2
    sum_hdr = rr
    for j, h in enumerate(["驅動（輸入）", "輸入取低值：加權目標價相對基準", "輸入取高值：加權目標價相對基準"], 1):
        f_(S, f"{openpyxl.utils.get_column_letter(j)}{sum_hdr}", h, bold=True, fill=HDR)
    rr += 1
    keys_seen = []
    for cs in cases:
        if cs["key"] not in keys_seen:
            keys_seen.append(cs["key"])
    sum_rows = {}
    for k in keys_seen:
        lab = next(c["label"] for c in cases if c["key"] == k).rsplit("：", 1)[0]
        f_(S, f"A{rr}", lab)
        if (k, "低") in case_row:
            f_(S, f"B{rr}", f"=E{case_row[(k, '低')]}", "+0.0%;-0.0%")
        if (k, "高") in case_row:
            f_(S, f"C{rr}", f"=E{case_row[(k, '高')]}", "+0.0%;-0.0%")
        sum_rows[k] = rr
        rr += 1
    Reg.rows["Sens_sum"] = (sum_hdr, sum_rows)
    Reg.rows["Sens_bid"] = bid_row
    S.column_dimensions["A"].width = 60
    # ───────── Sources ─────────
    SR = ws_["Sources"]
    f_(SR, "A1", "Sources — 來源（編號、文件、日期、網址、標記、用途）；逐列出處見 data/facts_crowdstrike.csv", bold=True)
    for j, h in enumerate(["編號", "文件", "日期", "網址", "標記", "用途"], 1):
        f_(SR, f"{openpyxl.utils.get_column_letter(j)}3", h, bold=True, fill=HDR)
    srcs = [("S01", "CrowdStrike Q2 FY2027 新聞稿", "2026-08-26", "https://ir.crowdstrike.com/news-releases/news-release-details/crowdstrike-reports-second-quarter-fiscal-year-2027-financial", "Verified", "Q2 實績、指引、模組採用、Flex"),
            ("S02", "CrowdStrike Form 10-Q（2026-07-31）", "2026-09", "https://www.sec.gov/Archives/edgar/data/1535527/000153552726000031/crwd-20260731.htm", "Verified", "資產負債表、RPO、員工、併購、DOJ／SEC 資訊請求"),
            ("S03", "CrowdStrike Form 10-K FY2026", "2026-03", "https://www.sec.gov/Archives/edgar/data/1535527/000153552726000010/crwd-20260131.htm", "Verified", "損益、ARR、淨留存、分部、RPO、員工"),
            ("S04", "CrowdStrike Q4／FY2026 新聞稿（8-K Ex.99.1）", "2026-03-03", "https://www.sec.gov/Archives/edgar/data/1535527/000153552726000007/crwd-20260303xex991.htm", "Verified", "FY2025／FY2026 損益與現金流"),
            ("S05", "CrowdStrike Q4 FY2024 新聞稿（8-K Ex.99.1）", "2024-03-05", "https://sec.gov/Archives/edgar/data/1535527/000153552724000004/crwd-20240305xex991.htm", "Verified", "FY2023／FY2024 損益與現金流"),
            ("S06", "美國財政部每日殖利率曲線（10 Yr）", "2026-11-13", "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202611", "Verified", "rf＝5.15%"),
            ("S07", "StockAnalysis：CRWD 歷史收盤價", "2026-11-13", "https://stockanalysis.com/stocks/crwd/history/", "Verified", "現價（單一資料商）"),
            ("S08", "StockAnalysis：CRWD 預測", "2026-11-03", "https://stockanalysis.com/stocks/crwd/forecast", "Interested-party", "賣方共識與目標價（只作對照）"),
            ("S09", "StockAnalysis：同業統計頁（PANW、FTNT、ZS、OKTA、S、CHKP）", "2026-11-14／15", "https://stockanalysis.com/stocks/panw/statistics/", "Interested-party", "同業 EV／營收（TTM）"),
            ("S10", "Q2 FY2027 法說會逐字稿（Webull 轉載）", "2026-08-26", "https://www.webull.com/news/15497781704532992", "Interested-party", "AIDR、Flex、定價、token 計價的口述"),
            ("S11", "Tokenomics v22.0 快照（data/tk_snapshot_v22.0.json）", "2026-11-16", "software_ai/unit_economics/data/tk_snapshot.json（子集）", "Derived", "推論成本只讀快照：token 價、任務 token 數、全成本對照"),
            ("S12", "研究包 data/research_crowdstrike.md", "2026-11-16", "software_ai/crowdstrike/data/", "Derived", "市場爭議、未找到清單")]
    for i, row in enumerate(srcs):
        for j, v in enumerate(row):
            f_(SR, f"{openpyxl.utils.get_column_letter(j + 1)}{4 + i}", v)
    Reg.rows["Sources_tk"] = 4 + 10
    SR.column_dimensions["B"].width = 60
    SR.column_dimensions["D"].width = 90
    # ───────── 儲存 ─────────
    out = out or os.path.join(SCR, "base.xlsx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    return out, build_id, TB, PB, CB


def recalc(paths, outdir):
    """LibreOffice headless 重算並轉存 xlsx（快取值）；回傳輸出路徑清單"""
    os.makedirs(outdir, exist_ok=True)
    cmd = ["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", outdir] + list(paths)
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=3000)
    return [os.path.join(outdir, os.path.basename(p)) for p in paths]


def read_val(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    v = wb["Val"]
    return {"tp": v["B8"].value, "dcf": v["B6"].value, "cmp": v["B7"].value}


def reverse():
    """G2-9：整本重算反解（網格＋線性內插，兩輪）。"""
    os.makedirs(SCR, exist_ok=True)
    inp = {it[0]: it for it in INP if it[0] != SEC}
    px = inp["px"][4]
    targets = [("rel_mult", "相對倍數係數 rel_mult（主導變數：Sens 最大）", 0.8, 12.0, "x"),
               ("gu_27", "人類端受保護單位年增 gu_27（FY2027；FY2032 值不動）", 0.0, 0.9, "%")]
    items = []
    for key, lab, lo, hi, unit in targets:
        base_v = inp[key][4]
        a_, b_ = lo, hi
        sol = None
        for rnd in range(2):
            grid = [a_ + (b_ - a_) * i / 8 for i in range(9)]
            fs = []
            for v in grid:
                pth, *_ = build(ov={key: v}, out=os.path.join(SCR, "rev", f"r_{key}_{rnd}_{v:.4f}.xlsx"))
                fs.append(pth)
            vals = [read_val(o)["tp"] for o in recalc(fs, os.path.join(SCR, "rev_rc"))]
            hit = None
            for i in range(8):
                if (vals[i] - px) * (vals[i + 1] - px) <= 0 and vals[i] != vals[i + 1]:
                    hit = i
                    break
            if hit is None:
                break
            x0, x1, y0, y1 = grid[hit], grid[hit + 1], vals[hit], vals[hit + 1]
            sol = x0 + (px - y0) * (x1 - x0) / (y1 - y0)
            a_, b_ = x0, x1
        within = "是" if sol is not None and inp[key][3] is not None and inp[key][3] <= sol <= (inp[key][5] if inp[key][5] is not None else 1e9) else "否"
        items.append({"key": key, "label": lab, "base": base_v, "solve": sol, "within": within if sol is not None else "不適用",
                      "note": f"其餘輸入不變；搜尋範圍 {lo:g}–{hi:g}" + ("；範圍內無交點" if sol is None else "")})
    json.dump({"px": px, "items": items}, open(os.path.join(DATA, "reverse_v01.json"), "w"), ensure_ascii=False, indent=1)
    print(items)


def main():
    if "--reverse" in sys.argv:
        reverse()
        return
    quick = "--quick" in sys.argv
    os.makedirs(SCR, exist_ok=True)
    base, bid, *_ = build(out=os.path.join(SCR, "base.xlsx"))
    rb = recalc([base], os.path.join(SCR, "rc"))[0]
    b = read_val(rb)
    print("base", b)
    sens_f = os.path.join(DATA, "sens_v01.json")
    if not quick:
        cases = []
        jobs = []
        for it in INP:
            if it[0] == SEC:
                continue
            key, lab, unit, lo, base_v, hi = it[:6]
            if lo is None and hi is None:
                continue
            jobs.append(it)
        files = []
        meta = []
        for it in jobs:
            key, lab, unit, lo, base_v, hi = it[:6]
            for side, val in (("低", lo), ("高", hi)):
                if val is None:
                    continue
                if not isinstance(base_v, str) and not isinstance(base_v, D) and abs(val - base_v) < 1e-12:
                    continue
                p, _bid, *_ = build(ov={key: val}, out=os.path.join(SCR, "sens", f"s_{key}_{side}.xlsx"))
                files.append(p)
                meta.append((key, lab, unit, base_v, lo, hi, side, val))
        outs = recalc(files, os.path.join(SCR, "sens_rc"))
        for (key, lab, unit, base_v, lo, hi, side, val), o in zip(meta, outs):
            r = read_val(o)
            fm = (lambda x: f"{x:.1%}" if unit == "%" else f"{x:g}")
            bs = "公式" if isinstance(base_v, str) else fm(base_v)
            cases.append({"key": key, "side": side, "label": f"{lab}：{side}", "setting": f"低／基準／高＝{fm(lo) if lo is not None else '—'}／{bs}／{fm(hi) if hi is not None else '—'}",
                          "tp": r["tp"], "dcf": r["dcf"], "cmp": r["cmp"]})
        inp = {it[0]: it for it in INP if it[0] != SEC}
        pxl, pxt, pxh = inp["px_tok"][3], 0.784, inp["px_tok"][5]
        joint = [
            ("jt_a", "聯動：token 價格取高案（Sol）且成本轉嫁客戶（情境 B 價格保留率取情境 A 值）", {"px_tok": pxh, "pf_B": inp["pf_A"][4]}, f"px_tok={pxh:g}、pf_B={inp['pf_A'][4]:g}"),
            ("jt_b", "模型組合：任務路由到便宜模型（Luna 與 Terra 價各半，Analogy）", {"px_tok": (pxl + pxt) / 2}, f"px_tok={(pxl + pxt) / 2:g}"),
            ("jt_c", "聯動：token 價格取低案（Luna）且客戶端同步壓價（情境 B 價格保留率取情境 C 值）", {"px_tok": pxl, "pf_B": inp["pf_C"][4]}, f"px_tok={pxl:g}、pf_B={inp['pf_C'][4]:g}"),
        ]
        jf = []
        for key, lab, ovd, st in joint:
            pth, _b, *_ = build(ov=ovd, out=os.path.join(SCR, "sens", f"s_{key}.xlsx"))
            jf.append(pth)
        for (key, lab, ovd, st), o in zip(joint, recalc(jf, os.path.join(SCR, "sens_rc"))):
            r = read_val(o)
            cases.append({"key": key, "side": "高", "label": f"{lab}：聯動", "setting": st, "tp": r["tp"], "dcf": r["dcf"], "cmp": r["cmp"]})
        json.dump({"base": b, "cases": cases}, open(sens_f, "w"), ensure_ascii=False, indent=1)
    sens = json.load(open(sens_f, encoding="utf-8"))
    sens["base"] = b
    final, bid, TB, PB, CB = build(sens=sens, out=os.path.join(SCR, "final", f"{FILE_DATE}_CrowdStrike個股模型_{VERSION}.xlsx"))
    rf = recalc([final], os.path.join(SCR, "final_rc"))[0]
    os.makedirs(DIST, exist_ok=True)
    for old in os.listdir(DIST):
        if old.endswith((".xlsx", ".html")):
            os.makedirs(os.path.join(DIST, "_待刪除"), exist_ok=True)
            shutil.move(os.path.join(DIST, old), os.path.join(DIST, "_待刪除", old))
    dst = os.path.join(DIST, os.path.basename(rf))
    shutil.copy(rf, dst)
    json.dump({"TB": TB, "PB": PB, "CB": CB, "reg": {k: (v if not isinstance(v, tuple) else [v[0], v[1]]) for k, v in Reg.rows.items()}}, open(os.path.join(SCR, "layout.json"), "w"))
    print("輸出", dst, bid)


if __name__ == "__main__":
    main()

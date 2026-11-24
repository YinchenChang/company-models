# 奧比中光 Orbbec（688322.SS，上海科創板）個股模型 v0.1 建置腳本
# 產出：physical_ai/orbbec/dist/20261124_奧比中光個股模型_v0.1.xlsx（LibreOffice 重算後的成品，含公式與快取值）
# 結構：命題模組（Thesis：機器人用 3D 視覺＝人形（線 4 出貨 × 線 6 每台相機數 × 份額 × 單價）＋其他機器人（線 6 AMR 銷量為成長指標）；
#       非機器人感測器、消費級應用設備（三維掃描與支付設備）、工業級另列）→ 官方產品分部 → 標準損益表 → FCFF DCF＋同業 P/E。
# 敏感度：每個案例覆寫 In 頁基準欄後整本重建、LibreOffice 重算（整本重算，不用閉式近似）；反解同法以二分搜尋。
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md、docs/plan/高不確定性預估工作原則.md。
#       實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics 協調檔。
# 用法：python3 physical_ai/orbbec/build_orbbec.py   （需要 soffice；暫存於家目錄下 ob_build）
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
NAME = f"20261124_奧比中光個股模型_{VER}.xlsx"
OUT = os.path.join(HERE, "dist", NAME)
HOMEDIR = os.path.expanduser("~")
TMP = os.path.join(HOMEDIR, "ob_build")
LO = ["soffice", f"-env:UserInstallation=file://{HOMEDIR}/lo_profile", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML"]

L4 = os.path.join(ROOT, "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx")
L6 = os.path.join(ROOT, "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx")
L4N, L6N = (os.path.relpath(p, ROOT) for p in (L4, L6))
L3N = "physical_ai/lines/industrial_robot"

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDRF = PatternFill("solid", fgColor="DDEBF7")
SECF = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------- 出處 ----------------
S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv：688322.SS；未還原股利）"
S_AR = "奧比中光 2025 年年度報告（2026-04-19 公告，天健會計師事務所標準無保留意見；公告 PDF 全文已讀取：主要會計數據、分產品與分行業營收、前五大客戶、三大報表、現金流量表補充資料）"
S_DW = "東吳證券 奧比中光-W 2026 年三季報點評（2026-10-30；東方財富研報 PDF）：轉錄三季報，附 Wind 三大表預測"
S_CM = "交銀國際證券 奧比中光 研報（東方財富研報頁，檔名日期 2026-10-29；買入、目標價 136.58 元）"
S_ZY = "中郵證券 奧比中光 研報（2026-04-26；年報後、三季報前，已過時，只作歷史對照）"
S_TJ = "同花順 2026-10-26 三季報報導（https://stock.10jqka.com.cn/20261026/c681101690.shtml）與 2026-10-27 投資者關係活動紀錄"
S_SA = "stockanalysis.com SHA:688322 statistics（資料商；2026-11-24 讀取；頁面報價為 11-23 收盤 101.95 元）"
S_RF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與 leaderdrive 等 A 股模型同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：China CRP 0.91%（與 A 股其他模型同一取值）"
S_CLOSE = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
S_SAP = "stockanalysis.com statistics 頁（資料商；2026-11-24 讀取，頁面報價日 2026-11-23 或 11-24；原報價與日期見 Comp 頁 L、M 欄）"
IP_BRK = "券商研究報告：有銷售研究與維持評等的誘因；預測為券商自設"
IP_SA = "資料商彙整，更新時點未必同步"
IP_CO = "發行公司說法，有股價與融資誘因"

# ---------------- 線模型取數（以 ID／代號找列，不以列號） ----------------
def line_data():
    wb4 = openpyxl.load_workbook(L4, data_only=True); ws4 = wb4["下游取數"]
    l4 = {ws4.cell(r, 1).value: (r, [ws4.cell(r, c).value for c in range(1, 10)]) for r in range(1, ws4.max_row + 1) if isinstance(ws4.cell(r, 1).value, str)}
    wb6 = openpyxl.load_workbook(L6, data_only=True); ws6 = wb6["輸入"]; wc = wb6["計算"]
    d = {"hum": {}, "amr": {}}
    for y in range(2026, 2033):
        r, v = l4[f"SHIP_{y}"]
        for s, c, col in (("A", 4, "E"), ("B", 3, "D"), ("C", 2, "C")):
            d["hum"][(y, s)] = (v[c], f"{L4N} 下游取數!{col}{r}（SHIP_{y}）")
    # 線 6 輸入頁：每台相機數（低／基準／高矩陣各一列 CAM，單位「顆」）、單價（單位「美元/顆」）
    cam = [r for r in range(1, ws6.max_row + 1) if ws6.cell(r, 1).value == "CAM" and ws6.cell(r, 3).value == "顆"]
    d["n_cam"] = [(ws6.cell(r, 6).value, f"{L6N} 輸入!F{r}") for r in cam[:3]]      # 人形：低、基準、高
    d["n_amr"] = [(ws6.cell(r, 7).value, f"{L6N} 輸入!G{r}") for r in cam[:3]]      # AMR
    pr = [r for r in range(1, ws6.max_row + 1) if ws6.cell(r, 1).value == "CAM" and ws6.cell(r, 3).value == "美元/顆"][0]
    L = openpyxl.utils.get_column_letter
    d["pd"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{L(c)}{pr}") for c in (7, 8, 9)]
    d["p26"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{L(c)}{pr}") for c in (4, 5, 6)]
    # AMR 銷量（萬台）：2026／2030／2032 低、基準、高；幾何插值（線 6 輸入頁「插值＝幾何」）
    ar = [r for r in range(1, ws6.max_row + 1) if ws6.cell(r, 1).value == "AMR"][0]
    pts = {"C": (5, 8, 11), "B": (6, 9, 12), "A": (7, 10, 13)}
    for s, (c26, c30, c32) in pts.items():
        a, b, c = (float(ws6.cell(ar, k).value) for k in (c26, c30, c32))
        for y in range(2026, 2033):
            if y <= 2030: v = a * (b / a) ** ((y - 2026) / 4)
            else: v = b * (c / b) ** ((y - 2030) / 2)
            d["amr"][(y, s)] = (v, f"{L6N} 輸入!{L(c26)}{ar}、{L(c30)}{ar}、{L(c32)}{ar}（2026／2030／2032 幾何插值，Derived）")
    # 線 6 計算頁：相機模組市場規模（百萬美元，2026–2032，基準）
    cr = None
    for r in range(100, wc.max_row + 1):
        if wc.cell(r, 1).value == "CAM" and isinstance(wc.cell(r, 3).value, (int, float)) and wc.cell(r, 10).value is None and wc.cell(r, 9).value is not None:
            cr = r; break
    d["cam_mkt"] = {y: (wc.cell(cr, 3 + y - 2026).value, f"{L6N} 計算!{L(3 + y - 2026)}{cr}") for y in range(2026, 2033)}
    r, v = l4["P_HARMONIC"] if "P_HARMONIC" in l4 else (None, [None] * 9)
    return d


LD = line_data()

# ---------------- In 頁輸入 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 來源誘因, 說明)
IN = [
    ("sec", "1. 市場資料與日期"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, "", "資料截止日 2026-11-24；股價 2026-11-20；成品日 2026-11-24"),
    ("px", "現價（2026-11-20 收盤）", "元", None, 85.09, None, "Verified", S_PX, "", "csv 值 85.089996；11-09 收盤 112.0，兩週跌 24%；資料商頁面 11-23 報價 101.95（隔個交易日反彈 20%，本模型不用）"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,12,31)", None, "Assumed", "本專案慣例：評價日 2026-12-31（會計年度＝日曆年，FY2026 在評價日已結束）", "", "折現期間由評價日起算"),
    ("tp_date", "目標時點（px_date 後 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", ""),
    ("bs_date", "資產負債表日", "日期", None, "=DATE(2026,9,30)", None, "Verified", S_DW, "", "2026 年三季報"),
    ("stub_frac", "Stub 期（資產負債表日→評價日）", "年", None, "=(val_date-bs_date)/365", None, "Derived", "＝(評價日 − 資產負債表日) ÷ 365", "", "Stub FCFF＝FY2026 FCFF × stub_frac"),
    ("sh_v", "評價日股數（總股本）", "百萬股", None, 412.83, None, "Verified", S_DW + "：總股本 412.83 百萬股（2025 年底約 400 百萬股，2026 年 6 月定增 9.75 億元淨額到帳後增加）；資料商頁 410.74", "", "限制性股票尚未歸屬部分未取得，以 0 計；H 股發行另見 h_iss"),
    ("cash_v", "貨幣資金＋短期投資（2026-09-30）", "百萬元", None, 1300.0, None, "Interested-party", S_SA + "：現金與短期投資 1.30 十億元（三季報原文未直接讀取）", IP_SA, "2025-12-31 貨幣資金 629.7＋交易性金融資產 310.1＝939.8（年報）；2026 年 6 月定增淨額 975"),
    ("ltinv_v", "長期金融資產（大額存單等，其他非流動資產；估計）", "百萬元", None, 1200.0, None, "Interested-party", S_DW + "：Wind 預測表 2026E 其他非流動資產 1,202（2025A 為 1,345，年報資產負債表）", IP_BRK, "明細（存單、理財）未取得；三季報原文未讀取，故敏感度低案只計 50%"),
    ("inv_k", "長期金融資產計入淨現金的比例", "%", 0.5, 1.0, None, "Assumed", "基準全額計入；低案 50%（流動性與期限不明）", "", ""),
    ("debt_v", "有息負債（2026-09-30）", "百萬元", None, 157.87, None, "Interested-party", S_SA + "：總負債（有息）157.87；2025 年底短期借款 127.3（年報）", IP_SA, ""),
    ("nci_v", "少數股東權益", "百萬元", None, 17.0, None, "Interested-party", S_DW + "：Wind 表 2025A 17", IP_BRK, "自淨現金扣除"),
    ("fx", "美元兌人民幣（2026-11-20）", "元／美元", None, 7.1, None, "Assumed", "與線 6 輸入頁匯率假設 7.1 一致（區間 6.9–7.3，僅用於 Recon 與線 6 市場規模對照）", "", "收盤檔 CNY=X 當日值未用於本模型，以線 6 假設對齊口徑"),
    ("cons_tp", "券商目標價（交銀國際，2026-10，三季報後）", "元", None, 136.58, None, "Interested-party", S_CM, IP_BRK, "只作對照，不進模型；東吳（首次買入）未列目標價"),
    ("sec", "2. 實績錨（FY2026＝2026 年，會計年度＝日曆年）"),
    ("r9m26", "2026 年 1–9 月營收", "百萬元", None, 710.9, None, "Verified", S_DW + "：約 7.11 億元、年增 −0.43%；" + S_TJ, "", "公司三季報原文未直接讀取；＝2025 年 1–9 月 714.0（年報分季）×（1 − 0.43%）"),
    ("g9m26", "2026 年 1–9 月營收年增", "%", None, -0.0043, None, "Verified", S_DW, "", "同花順報導為「基本持平」"),
    ("q4g", "2026 年第 4 季營收年增", "%", 0.0, 0.25, 0.75, "Assumed", "第 3 季單季 −1.9%；但存貨 9 月底 333（上年末 186）、合同負債翻倍，且採集設備 8 月上市後兩個月內逾百家客戶下單（" + S_TJ + "）；券商 FY2026 預測（" + S_DW + "：1,102；" + S_CM + "：1,131）隱含第 4 季約 +72%～+78%", "", "FY2026 營收＝1–9 月實績＋2025Q4 ×（1＋q4g）；基準取券商隱含值的約三分之一（證據不足以支持全額）"),
    ("sh_s26", "FY2026 3D 視覺感測器 占營收", "%", 0.28, 0.36, 0.44, "Assumed", "2025 年占 31%（年報分產品）；調整低毛利支付終端訂單、機器人視覺翻倍（" + S_DW + "）；2026 年分產品未揭露", "", "感測器為機器人用與非機器人用的合計"),
    ("sh_c26", "FY2026 消費級應用設備 占營收", "%", 0.45, 0.55, 0.65, "Assumed", "2025 年占 62%（三維掃描與支付『碰一下』設備）；公司減少傳統生物識別支付終端訂單（" + S_DW + "）；1–9 月營收持平而機器人翻倍，推知此項下降", "", "工業級及其他＝餘數"),
    ("gm26", "FY2026 毛利率", "%", None, 0.4909, None, "Verified", S_DW + "：2026 年 1–9 月毛利率 49.09%（年增 6.29pt）", "", "第 4 季假設同 1–9 月；2025 全年 43.8%"),
    ("opex26", "FY2026 銷售＋管理＋研發費用率", "%", 0.40, 0.43, 0.47, "Assumed", S_DW + "：1–9 月銷售 7.49%、管理 12.94%、研發 24.57%（合計 45.0%）；第 4 季營收較高，全年比率較低", "", "FY2025 為 38.7%"),
    ("sec", "3. 命題模組：機器人用 3D 視覺（量 × 價；線 4、線 6）"),
    ("asp25", "2025 每台 3D 視覺感測器平均營收", "元／台", None, "=Hist!E8/Hist!E11*100", None, "Derived", "＝2025 年 3D 視覺感測器營收 ÷ 銷量 88.35 萬台（" + S_AR + "）", "", "含全部用途（機器人、生物識別等）；線 6 相機模組單價 25 美元（約 178 元）為假設，口徑不同，見對帳"),
    ("pd", "相機單價年降幅（線 6）", "%／年", LD["pd"][0][0], LD["pd"][1][0], LD["pd"][2][0], "Assumed", f"線 6：{LD['pd'][1][1]}（低 {LD['pd'][0][1]}、高 {LD['pd'][2][1]}）", "", "線 6 標記為 Assumed（未取得相機單價資料）；本公司 2024→2025 感測器每台營收變動見 Hist"),
    ("n_cam", "人形每台相機數（線 6）", "顆／台", LD["n_cam"][0][0], LD["n_cam"][1][0], LD["n_cam"][2][0], "Analogy",
     f"線 6：{LD['n_cam'][1][1]}（低 {LD['n_cam'][0][1]}、高 {LD['n_cam'][2][1]}）；人形頭部 2–8 顆", "", "本公司深度相機未必每台都用；份額已吸收"),
    ("hu_k", "人形出貨相對線 4 的倍數（全部情境）", "x", None, 1.0, 2.0, "Assumed", "基準＝線 4 原值；高案 2 倍（中國人形 2026 年產量說法高於線 4 基準，未查核）", "", "反解變數"),
    ("rob_sh26", "FY2026 感測器營收中機器人用占比", "%", 0.40, 0.55, 0.70, "Assumed", "公司未揭露；公司稱國內機器人業務 2025 年翻倍、2026 年 1–9 月機器人視覺方案再翻倍（" + S_AR + "；" + S_TJ + "）；其餘為生物識別、三維掃描模組等", "", "推算 2025 年機器人用約 85–100 百萬元（占營收約 10%）"),
    ("hum_sh26", "FY2026 機器人用感測器中人形用占比", "%", 0.03, 0.10, 0.20, "Assumed", "公司未揭露；年報稱服務機器人客戶為主（云跡、擎朗、普渡、高仙等），人形客戶未具名", "", "隱含 2026 全球人形相機份額見 Thesis"),
    ("sec", "4. 命題模組：情境專屬驅動（A 樂觀／B 基準／C 保守；只列基準欄，不設敏感度區間）"),
    ("drift_A", "情境 A：其他機器人用感測器相對 AMR 銷量成長的年超額（量）", "%／年", None, 0.45, None, "Assumed", "線 6 AMR 銷量只是服務、物流機器人的代表指標；公司機器人視覺 1–9 月翻倍、境外直營翻倍；採集設備（EGO、UMI）為新品類", "", "超額成長含品類擴張與份額；樂觀"),
    ("drift_B", "情境 B：其他機器人用感測器相對 AMR 銷量成長的年超額（量）", "%／年", None, 0.25, None, "Assumed", "同上；基準取 2026 年增速（約 100%）的四分之一（證據只有一年）", "", ""),
    ("drift_C", "情境 C：其他機器人用感測器相對 AMR 銷量成長的年超額（量）", "%／年", None, 0.05, None, "Assumed", "同上；競爭者（Intel RealSense、舜宇、禾賽等）與客戶自製", "", ""),
    ("hs30_A", "情境 A：2030 年占全球人形相機用量份額", "%", None, 0.45, None, "Assumed", "2026 隱含份額見 Thesis；公司稱服務全球逾千家機器人客戶，人形客戶比重未揭露", "", "2026 隱含份額線性走到 2030，其後持平"),
    ("hs30_B", "情境 B：2030 年占全球人形相機用量份額", "%", None, 0.30, None, "Assumed", "同上；基準約持平 2026 隱含份額", "", ""),
    ("hs30_C", "情境 C：2030 年占全球人形相機用量份額", "%", None, 0.12, None, "Assumed", "同上；Tesla 等整機自製視覺、純視覺路線（線 6 對人形 4 顆為視覺相機假設）", "", ""),
    ("gnr_A", "情境 A：非機器人用感測器年增（FY2027 起）", "%", None, 0.05, None, "Assumed", "生物識別與三維掃描模組", "", ""),
    ("gnr_B", "情境 B：非機器人用感測器年增（FY2027 起）", "%", None, -0.05, None, "Assumed", "公司主動縮減低毛利支付終端訂單（" + S_DW + "）", "", ""),
    ("gnr_C", "情境 C：非機器人用感測器年增（FY2027 起）", "%", None, -0.15, None, "Assumed", "同上", "", ""),
    ("gcons_A", "情境 A：消費級應用設備年增（FY2027 起）", "%", None, 0.12, None, "Assumed", "三維掃描與支付『碰一下』設備；採集設備（EGO 等）若歸此類則上行；2025 年 +97%", "", "第一大客戶 30%、第二大（關聯方，螞蟻體系）27%，集中度高"),
    ("gcons_B", "情境 B：消費級應用設備年增（FY2027 起）", "%", None, 0.03, None, "Assumed", "同上", "", ""),
    ("gcons_C", "情境 C：消費級應用設備年增（FY2027 起）", "%", None, -0.15, None, "Assumed", "同上；大客戶訂單縮減", "", ""),
    ("gind_A", "情境 A：工業級應用設備及其他年增（FY2027 起）", "%", None, 0.15, None, "Assumed", "2025 年工業級 −2.1%、其他 +33.0%；基期小（合計約 7%）", "", ""),
    ("gind_B", "情境 B：工業級應用設備及其他年增（FY2027 起）", "%", None, 0.08, None, "Assumed", "同上", "", ""),
    ("gind_C", "情境 C：工業級應用設備及其他年增（FY2027 起）", "%", None, 0.0, None, "Assumed", "同上", "", ""),
    ("gm30_A", "情境 A：FY2030 毛利率", "%", None, 0.50, None, "Assumed", "2025 年 43.8%、2026 年 1–9 月 49.1%（結構轉向高毛利）；券商 2027E 46–50%（" + S_CM + "；" + S_DW + "）", "", "FY2026 起線性走到 FY2030，其後持平"),
    ("gm30_B", "情境 B：FY2030 毛利率", "%", None, 0.46, None, "Assumed", "同上；機器人產品價格競爭、規模擴大", "", ""),
    ("gm30_C", "情境 C：FY2030 毛利率", "%", None, 0.40, None, "Assumed", "同上；回到支付設備占比高的 2025 年水準", "", ""),
    ("tp_wA", "情境權重 A", "", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", ""),
    ("tp_wB", "情境權重 B", "", None, 0.50, None, "Assumed", "同上", "", ""),
    ("tp_wC", "情境權重 C", "", None, 0.25, None, "Assumed", "同上", "", ""),
    ("sec", "5. 損益與現金流"),
    ("ts_r", "稅金及附加 ÷ 營收", "%", None, "=Hist!E19/Hist!E4", None, "Derived", "＝2025 年稅金及附加 ÷ 營收", "", ""),
    ("opex30", "FY2030 銷售＋管理＋研發費用率", "%", 0.22, 0.28, 0.34, "Assumed", "2025 年 38.7%、2026 年 1–9 月 45.0%；券商 2027E 約 31–34%（" + S_DW + "：銷售＋管理＋研發 ÷ 營收）", "", "FY2026→FY2030 線性，其後持平；研發投入強度高，營運槓桿是否出現未驗證"),
    ("oi_r", "其他收益（政府補助等）÷ 營收", "%", 0.02, 0.04, 0.059, "Assumed", "2025 年其他收益 55.7＝5.9%（年報）；券商 2026E 36（約 3%）", "", "計入營業利益"),
    ("imp_r", "減值損失（含信用減值）÷ 營收", "%", 0.01, "=-Hist!E23/Hist!E4", 0.03, "Derived", "＝2025 年（資產減值 18.4＋信用減值 1.3）÷ 營收（年報）", "", "計入營業利益"),
    ("inv_y", "金融資產收益率（投資收益＋公允價值變動＋利息）", "%", 0.01, 0.02, 0.03, "Assumed", "2025 年投資收益 36.4＋公允價值 2.5＋淨利息收入 4.6＝43.5，÷ 平均金融資產約 1,900（Derived）", "", "業外；DCF 另以淨現金計入，不重複"),
    ("tax_r", "有效稅率", "%", 0.05, 0.10, 0.15, "Assumed", "2025 年所得稅 −8.8（遞延所得稅利益；稅前 118.8）；高新技術企業 15%＋研發加計扣除＋累計虧損可抵；2026E 券商約 1%", "", "稅率偏低且逐步正常化未建模（敏感度高 15%）"),
    ("nci_r", "少數股東損益 ÷ 稅後淨利", "%", None, "=Hist!E35/Hist!E34", None, "Derived", "＝2025 年少數股東損益 ÷ 稅後淨利", "", ""),
    ("payout", "配息率", "%", None, 0.0, None, "Verified", S_AR + "：母公司未分配利潤為負，2025 年度不分配；改以回購（2025 年 4,804 萬元）", "", "收支關卡以配息率 0 計；回購未計入股利"),
    ("da26", "FY2026 折舊攤銷 ÷ 營收", "%", None, "=Hist!E40/Hist!E4", None, "Derived", "＝2025 年折舊攤銷 49.8 ÷ 營收（年報現金流量表補充資料）", "", "之後線性收斂到資本支出比例 × da_ss（FY2032）"),
    ("da_ss", "穩態折舊 ÷ 資本支出", "x", None, 0.8, None, "Assumed", "穩態資本支出高於折舊（成長）", "", ""),
    ("cx26", "FY2026 資本支出", "百萬元", 100, 180, 250, "Assumed", S_DW + "：2026E 資本開支 180（越南工廠 2026-05 動工）；2025 年 63、2024 年 203（年報）", "", ""),
    ("cx_r", "資本支出 ÷ 營收（FY2027 起）", "%", 0.05, 0.08, 0.12, "Assumed", "2025 年 6.7%、2024 年 36%（廠房建設）；券商 2027E 約 10%", "", ""),
    ("nwc_r", "營運資金增加 ÷ 營收增量", "%", 0.20, 0.30, 0.45, "Assumed", "2025 年營運資本變動 −63（券商表）；2026 年 9 月存貨 333（上年末 186）", "", "存貨水位上升"),
    ("h_iss", "H 股發行占擴大後股本（尾端事件，基準 0）", "%", None, 0.0, 0.10, "Assumed", "2026-07-23 公告籌備 H 股上市；規模、時程未揭露（" + S_TJ + "）", IP_CO, "原則第 10 節第 19 條：事件型尾端風險進 Sens、不進基準"),
    ("h_disc", "H 股發行價相對現價折價", "%", None, 0.10, None, "Assumed", "港股 A＋H 慣例折價（未查核，不作依據）", "", "只在 h_iss＞0 時有作用"),
    ("sec", "6. 折現率（DCF 標準規格第 2 節）"),
    ("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", S_RF, "", "日期 2026-11-05，較股價基準日早 11 個交易日（逾期；跨模型一致，下次全面改版更新）"),
    ("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", ""),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", ""),
    ("b_own", "本公司 5 年 β（原始）", "x", None, 0.80, None, "Verified", S_SA, "", ""),
    ("b_peer", "同業 5 年 β 中位數（原始）", "x", None, "=MEDIAN(Comp!H5:H10)", None, "Derived", "＝Comp 頁核心同業 3 家 β 中位數（LG Innotek β 2.63 偏高，取中位數）", "", ""),
    ("beta", "β（兩來源 Blume 調整後中位數）", "x", "=beta-0.3", "=MEDIAN(0.67*b_own+0.33,0.67*b_peer+0.33)", "=beta+0.3", "Derived", "Blume 0.67β＋0.33；兩來源＝本公司、同業中位數", "", "敏感度 ∓0.3（跨模型慣例）"),
    ("kd", "稅前債務成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt", "", "有息負債占比小"),
    ("wacc_adj", "WACC 調整（敏感度用）", "pt", -0.01, 0, 0.01, "Assumed", "基準 0；敏感度 ±1pt", "", ""),
    ("g_cap", "永續成長率上限", "%", None, 0.03, None, "Assumed", "DCF 標準規格：3%，且不超過計價幣別 10 年期公債殖利率", "", ""),
    ("g_tv", "永續成長率", "%", 0.01, "=MIN(g_cap,rf)", 0.03, "Derived", "＝MIN(g_cap, rf)", "", "敏感度 1%／3%（3% 超過 rf，只作敏感度）"),
    ("ronic_sp", "RONIC − WACC", "pt", None, 0, 0.05, "Assumed", "RONIC 預設＝WACC（新投資價值中性）", "", ""),
    ("tv_alt", "終值改用 FCFF ×（1＋g）（0＝否，1＝是）", "開關", None, 0, 1, "Assumed", "DCF 標準規格第 4 節敏感度", "", ""),
    ("sec", "7. 同業倍數與權重"),
    ("pe_peer", "同業前瞻 P/E 中位數", "x", "=MIN(Comp!F5:F10)", "=MEDIAN(Comp!F5:F10)", "=MAX(Comp!F5:F10)", "Derived", "＝Comp 頁核心同業 3 家前瞻 P/E 中位數；股價皆為 2026-11-20 收盤、P/E 由原報價日等比例換算（EPS 分母不變，見 Comp）", "", "敏感度＝同業最低／最高；僅 3 家，樣本小"),
    ("rel_k", "相對倍數係數（本公司 ÷ 同業）", "x", None, 1.0, 2.0, "Assumed", "1.00（未調整：缺本公司對同業十年相對倍數歷史）", "", "高案 2 倍（未查核，不作依據）"),
    ("w_dcf", "DCF 權重（Comparable＝1 − 本值）", "%", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設 DCF 50%＋同業倍數 50%", "", ""),
]
SCN_DRIVERS = [("drift", "其他機器人用感測器年超額成長"), ("hs30", "2030 年人形相機份額"), ("gnr", "非機器人用感測器年增"),
               ("gcons", "消費級應用設備年增"), ("gind", "工業級及其他年增"), ("gm30", "FY2030 毛利率")]
SCN_INFO = {"gnr": ("A", "C"), "gcons": ("A", "B", "C"), "gind": ("A", "B", "C")}

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

    # ================= Hist =================
    HI["A1"] = "Hist：歷史財務（百萬元人民幣；中國會計準則；2024、2025 取自 2025 年年度報告，2023 只取年報主要會計數據表列示項目）。公式只引用本頁（定義名稱 h_*）。"
    hdr(HI, 3, ["項目", "", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"])
    HROWS = [  # (列, 鍵, 項目, [2023, 2024, 2025], 標記, 來源)
        (4, "h_rev", "營業收入", [360.01, 564.46, 940.74], "Verified", S_AR),
        (5, "h_cogs", "營業成本", [None, 328.62, 529.12], "Verified", S_AR),
        (6, "h_gp", "毛利（＝營收 − 營業成本）", [None, "=D4-D5", "=E4-E5"], "Derived", "＝營業收入 − 營業成本"),
        (7, "h_gm", "毛利率", [None, "=D6/D4", "=E6/E4"], "Derived", "＝毛利 ÷ 營收"),
        (8, "h_seg_h", "3D 視覺感測器 營收", [None, "=290.93/1.3679", 290.93], "Verified", S_AR + "：分產品表（2025 年 290.93、年增 36.79%）；2024＝2025 ÷ 1.3679（Derived）"),
        (9, "h_seg_m", "消費級應用設備 營收", [None, "=583.73/1.9729", 583.73], "Verified", S_AR + "：分產品表（2025 年 583.73、年增 97.29%；三維掃描、支付『碰一下』設備）；2024＝2025 ÷ 1.9729（Derived）"),
        (10, "h_seg_a", "工業級應用設備及其他 營收（＝營收 − 上兩項；含其他業務收入）", [None, "=D4-D8-D9", "=E4-E8-E9"], "Derived", "工業級 25.54、其他 34.84（年報分產品）＋其他業務收入約 5.7；以殘差計"),
        (11, "h_hd_units", "3D 視覺感測器銷量（萬台）", [None, "=88.35/1.0726", 88.35], "Verified", S_AR + "：產銷量表 2025 年 88.35 萬台（年增 7.26%）；2024＝2025 ÷ 1.0726（Derived）"),
        (12, "h_asp", "每台感測器營收（元／台）", [None, "=D8/D11*100", "=E8/E11*100"], "Derived", "＝感測器營收 ÷ 銷量"),
        (13, "h_sell", "銷售費用", [None, 73.44, 69.11], "Verified", S_AR),
        (14, "h_admin", "管理費用", [None, 104.87, 92.17], "Verified", S_AR),
        (15, "h_rnd", "研發費用", [None, 204.33, 202.53], "Verified", S_AR),
        (16, "h_opex", "營業費用（銷售＋管理＋研發）", [None, "=SUM(D13:D15)", "=SUM(E13:E15)"], "Derived", "＝三項合計"),
        (17, "h_fin", "財務費用（負＝淨利息收入）", [None, -35.78, -4.58], "Verified", S_AR),
        (18, "h_inv", "投資收益＋公允價值變動", [None, "=37.04+0.09", "=36.35+2.52"], "Verified", S_AR),
        (19, "h_ts", "稅金及附加", [None, 3.34, 7.39], "Verified", S_AR),
        (20, "h_opp", "營業利潤（中國會計準則）", [None, -62.79, 119.95], "Verified", S_AR),
        (21, "h_ebit", "營業利益（核心：營業利潤 − 投資收益與公允價值 ＋ 財務費用）", [None, "=D20-D18+D17", "=E20-E18+E17"], "Derived", "本模型營業利益不含金融資產收益與財務費用（DCF 另以淨現金計入）；2025 年 76.5 與東吳 EBIT 76 一致"),
        (22, "h_oth_op", "其他營業損益（＝核心營業利益 − 毛利 ＋ 稅金及附加 ＋ 營業費用）", [None, "=D21-D6+D19+D16", "=E21-E6+E19+E16"], "Derived", "其他收益、減值、資產處置合計（殘差）"),
        (23, "h_imp", "減值損失（資產減值＋信用減值）", [None, "=-(1.26+21.73)", "=-(1.28+18.42)"], "Verified", S_AR),
        (24, "h_oi", "其他收益（政府補助等）", [None, 37.24, 55.67], "Verified", S_AR),
        (25, "h_ebitm", "營業利益率（核心）", [None, "=D21/D4", "=E21/E4"], "Derived", ""),
        (32, "h_pbt", "利潤總額（稅前）", [-268.87, -65.42, 118.83], "Verified", S_AR + "（2023 取主要會計數據表）"),
        (26, "h_nonop", "業外（＝稅前 − 核心營業利益）", [None, "=D32-D21", "=E32-E21"], "Derived", "投資收益、公允價值、財務費用、營業外收支合計"),
        (33, "h_tax", "所得稅（負＝利益）", [None, -2.52, -8.79], "Verified", S_AR),
        (34, "h_nit", "稅後淨利", [None, -62.90, 127.62], "Verified", S_AR),
        (35, "h_nci", "少數股東損益", [None, "=D34-D36", "=E34-E36"], "Derived", "＝稅後淨利 − 歸屬母公司淨利"),
        (36, "h_ni", "歸屬母公司淨利", [-275.88, -62.91, 127.91], "Verified", S_AR),
        (37, "h_eps", "EPS（基本，元）", [-0.69, -0.16, 0.32], "Verified", S_AR),
        (38, "h_dps", "每股股利（元）", [None, 0, 0], "Verified", S_AR + "：不分配（母公司未分配利潤為負）"),
        (39, "h_sh", "股數（＝淨利 ÷ EPS，百萬股）", ["=C36/C37", "=D36/D37", "=E36/E37"], "Derived", ""),
        (40, "h_da", "折舊攤銷", [None, "=27.46+10.38+15.61+17.11", "=28.54+2.00+12.17+7.13"], "Verified", S_AR + "：現金流量表補充資料（固定資產折舊、使用權資產攤銷、無形資產攤銷、長期待攤費用攤銷）"),
        (41, "h_cx", "資本支出", [None, 202.82, 63.12], "Verified", S_AR + "：購建固定資產、無形資產和其他長期資產支付的現金"),
        (42, "h_ocf", "營業現金流", [-159.76, -86.34, 82.72], "Verified", S_AR),
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
    for c in "DE": HI[f"{c}28"] = f"={c}8+{c}9+{c}10-{c}4"
    HI.cell(29, 1, "差額：核心營業利益＋業外 − 稅前（應為 0）")
    for c in "DE": HI[f"{c}29"] = f"={c}21+{c}26-{c}32"
    HI.cell(30, 1, "說明：2023 年年報表格只取主要會計數據（營收、稅前、歸母淨利、EPS、營業現金流）；其餘科目未取得（2023 年年報未讀取），不補數。")
    HI.column_dimensions["A"].width = 58; HI.column_dimensions["G"].width = 90

    # ================= Q（季度；一列一個數字，單一標記）=================
    Q["A1"] = "Q：季度資料（百萬元）。2025 年分季取自年報「分季度主要財務數據」；2026 年取自券商轉錄的三季報（原文未直接讀取）。"
    hdr(Q, 3, ["項目", "2025Q1", "2025Q2", "2025Q3", "2025Q4", "2026Q3", "2026 年 1–9 月", "標記", "來源", "來源誘因", "說明"])
    QROWS = [
        (4, "營業收入", [191.06, 244.41, 278.53, 226.75, 273.3, 710.9], "Verified", S_AR + "（2025 分季）；" + S_DW + "（2026Q3 2.73 億、1–9 月 7.11 億）", "", "2026Q3 年增 −1.86%；1–9 月年增 −0.43%"),
        (5, "歸屬母公司淨利", [24.32, 35.88, 47.83, 19.89, 32.22, 73.86], "Verified", S_AR + "；" + S_DW, "", "2026 前三季年增 −31.6%（研發、管理費用增加）"),
        (6, "毛利率（2026 年 1–9 月）", [None, None, None, None, None, 0.4909], "Verified", S_DW, "", "年增 +6.29pt（調整低毛利支付終端訂單）"),
    ]
    for (rr, lab, vals, tag, src, inc, note) in QROWS:
        Q.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = Q.cell(rr, 2 + j, v); c.font = BLUE
                if isinstance(v, float) and v < 1: c.number_format = "0.00%"
        Q.cell(rr, 8, tag); Q.cell(rr, 9, src); Q.cell(rr, 10, inc); Q.cell(rr, 11, note)
    Q.column_dimensions["A"].width = 30; Q.column_dimensions["I"].width = 90

    # ================= 線取數 =================
    LN["A1"] = "線取數：本模型讀入的線層級變數年度路徑（建置時由線模型成品依 ID 讀值）。情境對應：A＝線模型高情境、B＝基準、C＝低情境。"
    hdr(LN, 3, ["項目", "情境", ""] + [f"FY{y}" for y in range(2025, 2033)] + ["單位", "來源（檔案與儲存格）"])
    LR = {}
    r = 4
    for key, lab, unit, dd, scale in (("hum", "全球人形機器人出貨（線 4）", "台", LD["hum"], 1),
                                      ("amr", "全球運輸物流機器人（AMR）年銷量（線 6 輸入頁，幾何插值）", "萬台", LD["amr"], 1)):
        for s in SC:
            LN.cell(r, 1, lab); LN.cell(r, 2, f"{s}（{LINE_SC[s]}）")
            srcs = []
            for y in range(2026, 2033):
                v, src = dd[(y, s)]
                c = LN.cell(r, 4 + (y - 2025), v); c.font = BLUE; c.number_format = "#,##0.0" if key == "amr" else "#,##0"
                srcs.append(src)
            LN.cell(r, 12, unit); LN.cell(r, 13, f"{srcs[0]} … {srcs[-1]}")
            LR[(key, s)] = r; r += 1
    r += 1
    LN.cell(r, 1, "線 6 相機模組市場規模（五線歸因；百萬美元；基準；只作 Recon 對照）").font = BOLD; r += 1
    LR["cam_mkt"] = r
    for y in range(2026, 2033):
        v, src = LD["cam_mkt"][y]
        c = LN.cell(r, 4 + (y - 2025), v); c.font = BLUE; c.number_format = "#,##0.0"
    LN.cell(r, 1, "線 6 相機模組市場規模"); LN.cell(r, 2, "B（基準）"); LN.cell(r, 12, "百萬美元"); LN.cell(r, 13, LD["cam_mkt"][2026][1] + " … " + LD["cam_mkt"][2032][1])
    LN.column_dimensions["A"].width = 52; LN.column_dimensions["M"].width = 110

    # ================= Thesis（命題模組）=================
    TH["A1"] = ("命題模組：機器人用 3D 視覺營收＝人形用（線 4 出貨 × 線 6 每台相機數 × 份額 × 單價）＋其他機器人用（前一年 × 線 6 AMR 銷量成長 ×（1＋超額）×（1 − 單價年降））。"
                "非機器人用感測器、消費級應用設備、工業級及其他各自以年增率推算。FY2026＝1–9 月實績＋第 4 季假設。百萬元。")
    TH["A2"] = "欄：D＝FY2025 … K＝FY2032。人形 2026 隱含份額以線 4 基準路徑校準（三情境同一實績年）。"
    hdr(TH, 3, ["項目", "說明", ""] + [f"FY{y}" for y in range(2025, 2033)])
    THR = {}
    r = 4
    for s in SC:
        TH.cell(r, 1, f"情境 {s}").font = BOLD; TH.cell(r, 1).fill = SECF; r += 1
        R = {}
        R["rev26"] = r; TH.cell(r, 1, "FY2026 合併營收（錨）"); TH.cell(r, 2, "＝1–9 月＋2025Q4 ×（1＋q4g）")
        TH[f"{TC[2026]}{r}"] = "=r9m26+(Hist!E4-r9m26/(1+g9m26))*(1+q4g)"; r += 1
        R["asp"] = r; TH.cell(r, 1, "每台感測器平均營收（元／台）")
        for y in range(2025, 2033):
            TH[f"{TC[y]}{r}"] = "=asp25" if y == 2025 else f"={TC[y - 1]}{r}*(1-pd)"
        r += 1
        R["amr"] = r; TH.cell(r, 1, "全球 AMR 銷量（線 6，萬台）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR[('amr', s)]}"
        r += 1
        R["ship"] = r; TH.cell(r, 1, "全球人形出貨（線 4 × 出貨倍數，千台）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR[('hum', s)]}*hu_k/1000"
        r += 1
        R["hdem"] = r; TH.cell(r, 1, "全球人形相機需求（千顆）"); TH.cell(r, 2, "＝出貨 × 每台相機數（線 6）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['ship']}*n_cam"
        r += 1
        R["sens26"] = r; TH.cell(r, 1, "FY2026 3D 視覺感測器（錨）"); TH[f"{TC[2026]}{r}"] = f"={TC[2026]}{R['rev26']}*sh_s26"; r += 1
        R["hsh"] = r; TH.cell(r, 1, "人形相機份額"); TH.cell(r, 2, "FY2026＝隱含（以線 4 基準校準）；線性走到 hs30，其後持平")
        TH[f"{TC[2026]}{r}"] = f"={TC[2026]}{R['sens26']}*rob_sh26*hum_sh26/(線取數!{TC[2026]}{LR[('hum', 'B')]}*hu_k/1000*n_cam*{TC[2026]}{R['asp']}/1000)"
        for y in range(2027, 2033):
            TH[f"{TC[y]}{r}"] = f"=${TC[2026]}${r}+(hs30_{s}-${TC[2026]}${r})*MIN(1,({y}-2026)/4)"
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 1
        R["hum"] = r; TH.cell(r, 1, "感測器：人形機器人用 營收"); TH.cell(r, 2, "FY2027 起＝需求（千顆）× 份額 × 單價 ÷ 1000")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = f"={c}{R['sens26']}*rob_sh26*hum_sh26" if y == 2026 else f"={c}{R['hdem']}*{c}{R['hsh']}*{c}{R['asp']}/1000"
        r += 1
        R["orob"] = r; TH.cell(r, 1, "感測器：其他機器人用 營收"); TH.cell(r, 2, "FY2027 起＝前一年 × AMR 銷量成長 ×（1＋超額）×（1 − 年降）")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = (f"={c}{R['sens26']}*rob_sh26*(1-hum_sh26)" if y == 2026 else
                             f"={TC[y - 1]}{r}*{c}{R['amr']}/{TC[y - 1]}{R['amr']}*(1+drift_{s})*(1-pd)")
        r += 1
        R["nrob"] = r; TH.cell(r, 1, "感測器：非機器人用 營收")
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = f"={c}{R['sens26']}*(1-rob_sh26)" if y == 2026 else f"={TC[y - 1]}{r}*(1+gnr_{s})"
        r += 1
        R["sens"] = r; TH.cell(r, 1, "3D 視覺感測器 合計")
        for y in range(2026, 2033): c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['hum']}+{c}{R['orob']}+{c}{R['nrob']}"
        r += 1
        R["cons"] = r; TH.cell(r, 1, "消費級應用設備")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['rev26']}*sh_c26" if y == 2026 else f"={TC[y - 1]}{r}*(1+gcons_{s})"
        r += 1
        R["indo"] = r; TH.cell(r, 1, "工業級應用設備及其他")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['rev26']}*(1-sh_s26-sh_c26)" if y == 2026 else f"={TC[y - 1]}{r}*(1+gind_{s})"
        r += 1
        R["humsh"] = r; TH.cell(r, 1, "人形用占合併營收")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['hum']}/({c}{R['sens']}+{c}{R['cons']}+{c}{R['indo']})"; TH[f"{c}{r}"].number_format = "0.0%"
        r += 1
        R["robsh"] = r; TH.cell(r, 1, "機器人用（人形＋其他機器人）占合併營收")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"=({c}{R['hum']}+{c}{R['orob']})/({c}{R['sens']}+{c}{R['cons']}+{c}{R['indo']})"; TH[f"{c}{r}"].number_format = "0.0%"
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
                if y == 2023 and key not in ("rev", "pbt", "ni", "nim", "eps", "sh"): f = None
                if f is not None:
                    PL[f"{COL[y]}{r}"] = f
                    if fmt: PL[f"{COL[y]}{r}"].number_format = fmt
            r += 1
        H = lambda rr: (lambda y: f"=Hist!{'CDE'[y - 2023]}{rr}")
        C = lambda y: COL[y]
        row("hum", "感測器：人形機器人用（子列）", "百萬元", None, lambda y: f"=Thesis!{TC[y]}{th['hum']}")
        row("orob", "感測器：其他機器人用（子列）", "百萬元", None, lambda y: f"=Thesis!{TC[y]}{th['orob']}")
        row("nrob", "感測器：非機器人用（子列）", "百萬元", None, lambda y: f"=Thesis!{TC[y]}{th['nrob']}")
        row("harm", "3D 視覺感測器 營收", "百萬元", H(8), lambda y: f"=SUM({C(y)}{R['hum']}:{C(y)}{R['nrob']})")
        row("mech", "消費級應用設備 營收", "百萬元", H(9), lambda y: f"=Thesis!{TC[y]}{th['cons']}")
        row("ao", "工業級應用設備及其他 營收", "百萬元", H(10), lambda y: f"=Thesis!{TC[y]}{th['indo']}")
        row("rev", "營收（合併）", "百萬元", H(4), lambda y: f"=SUM({C(y)}{R['harm']}:{C(y)}{R['ao']})")
        row("chk", "檢查：分部合計 − 合併營收（應為 0）", "百萬元",
            lambda y: f"={C(y)}{R['harm']}+{C(y)}{R['mech']}+{C(y)}{R['ao']}-{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['harm']}+{C(y)}{R['mech']}+{C(y)}{R['ao']}-{C(y)}{R['rev']}")
        PL.cell(R["chk"], 1, "檢查：分部合計 − 合併營收（應為 0）" if s == "B" else f"檢查（情境 {s}）：分部合計 − 合併營收（應為 0）")
        row("chk2", "檢查：感測器子列合計 − 感測器分部（應為 0）", "百萬元", None, lambda y: f"={C(y)}{R['hum']}+{C(y)}{R['orob']}+{C(y)}{R['nrob']}-Thesis!{TC[y]}{th['sens']}")
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
    CP["A1"] = ("Comp：同業前瞻 P/E（stockanalysis.com statistics，資料商）。股價＝" + S_CLOSE + "（Verified）；前瞻 P/E＝原值 × 新收盤 ÷ 原報價（Derived，EPS 不變）。"
                "核心同業僅 3 家（少於 8 家不做迴歸）；收盤檔無報價或資料商無前瞻值者另列「報價日未取得」，不進中位數。")
    hdr(CP, 4, ["公司", "代號", "股價（2026-11-20）", "報價日", "前瞻 P/E 分母", "前瞻 P/E", "近 12 月 P/E（原報價日）", "5 年 β", "會計口徑", "進中位數", "說明", "原報價", "原報價日", "原前瞻 P/E", "標記與換算"])
    PEERS = [  # (公司, 代號, 11-20 收盤, 原報價, 原報價日, 原前瞻 P/E, 近 12 月 P/E, β, 口徑, 進中位數, 說明, 分母)
        ("舜宇光學", "2382.HK", 58.5, 63.05, "2026-11-24", 13.16, 12.35, 1.22, "港股（IFRS，未查核）", "是", "手機與車載鏡頭、相機模組、LiDAR 光學", "資料商 NTM 共識 EPS"),
        ("韋爾股份", "603501.SS", 83.91, 86.68, "2026-11-24", 22.93, 34.18, 0.86, "中國會計準則", "是", "CMOS 影像感測器；視覺感測上游", "資料商 NTM 共識 EPS"),
        ("LG Innotek", "011070.KS", 267500, 284500, "2026-11-24", 7.75, 10.69, 2.63, "韓國 IFRS（未查核）", "是", "相機模組（含機器人相機）；蘋果相機模組供應商；β 偏高", "資料商 NTM 共識 EPS"),
        ("立訊精密", "002475.SZ", 45.70, None, "報價日未取得", None, None, None, "中國會計準則", "否", "連接器與消費電子代工；收盤檔有收盤價，但資料商頁面取得失敗（404），前瞻 P/E 報價日未取得", "—"),
        ("禾賽科技", "HSAI", 11.83, None, "報價日未取得", None, None, None, "美國公認會計準則（未查核）", "否", "機器人與車用 LiDAR；虧損，前瞻 P/E 無意義，不納入", "—"),
        ("速騰聚創", "2498.HK", 13.32, None, "報價日未取得", None, None, None, "港股（IFRS，未查核）", "否", "LiDAR；虧損，前瞻 P/E 無意義，不納入", "—"),
        ("Intel RealSense", "（已分拆，未上市）", None, None, "報價日未取得", None, None, None, "—", "否", "深度相機直接競爭者；非上市公司，無報價", "—"),
        ("思特威", "688213.SS", None, None, "報價日未取得", None, None, None, "中國會計準則", "否", "CMOS 影像感測器；收盤檔無此代號，不納入", "—"),
        ("奧比中光（本公司，對照）", "688322.SS", 85.09, 101.95, "2026-11-23", 125.09, 443.05, 0.80, "中國會計準則", "否", "本公司；只作對照", "資料商 NTM 共識 EPS"),
    ]
    for i, (nm, tk, px1, px0, d0, pe0, ttm, bt, acc, med, note, den) in enumerate(PEERS):
        rr = 5 + i
        CP.cell(rr, 1, nm); CP.cell(rr, 2, tk)
        if px1 is not None: CP.cell(rr, 3, px1).font = BLUE
        CP.cell(rr, 4, "2026-11-20" if px1 is not None else "—"); CP.cell(rr, 5, den)
        if pe0 is not None: CP.cell(rr, 6, f"=N{rr}*C{rr}/L{rr}")
        else: CP.cell(rr, 6, "未取得")
        if ttm is not None: CP.cell(rr, 7, ttm).font = BLUE
        if bt is not None: CP.cell(rr, 8, bt).font = BLUE
        CP.cell(rr, 9, acc); CP.cell(rr, 10, med); CP.cell(rr, 11, note)
        if px0 is not None: CP.cell(rr, 12, px0).font = BLUE
        CP.cell(rr, 13, d0)
        if pe0 is not None: CP.cell(rr, 14, pe0).font = BLUE
        if pe0 is not None:
            CP.cell(rr, 15, "股價 Verified（" + S_CLOSE + "）；前瞻 P/E Derived：＝原值 × 2026-11-20 收盤 ÷ 原報價（分母 EPS 不變）；原值 " + S_SAP)
        else:
            CP.cell(rr, 15, "報價日未取得：不進中位數")
    CP.cell(15, 1, "核心同業前瞻 P/E 中位數"); CP["F15"] = "=MEDIAN(F5:F10)"
    CP.cell(16, 1, "本公司現價 ÷ 資料商前瞻 EPS（換算至 2026-11-20）"); CP["F16"] = "=F13"
    CP.cell(17, 1, "本公司現價 ÷ 券商 FY2027 EPS 平均（分母：東吳 0.73、交銀國際 366÷412.83、中郵 430÷412.83 元，FY2027）"); CP["F17"] = "=px/AVERAGE(0.73,366/412.83,430/412.83)"
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
    row1("cons_show", "券商目標價（交銀國際，2026-10；只作對照）", "=cons_tp")
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
    rrow("FY2026 營收（情境 B）vs 券商 2026E 平均（東吳 1,102、交銀國際 1,131）", f"=PL!F{PR['B']['rev']}", "=AVERAGE(1102,1131)", "兩家皆為三季報後；券商隱含第 4 季營收 391～420（年增 +72%～+85%），本模型第 4 季年增 q4g，差距主要在此；中郵 1,460 為年報後、已過時")
    rrow("FY2027 營收（情境 B）vs 券商 2027E 平均（東吳 2,045、交銀國際 1,880、中郵 2,040）", f"=PL!G{PR['B']['rev']}", "=AVERAGE(2045,1880,2040)", "券商預測為 Interested-party；券商隱含採集設備與機器人視覺大幅放量，本模型證據不足以支持，故低於券商")
    rrow("FY2028 營收（情境 B）vs 券商 2028E 平均（東吳 3,505、交銀國際 2,789、中郵 2,570）", f"=PL!H{PR['B']['rev']}", "=AVERAGE(3505,2789,2570)", "")
    rrow("FY2026 歸屬母公司淨利 vs 券商 2026E 平均（東吳 180、交銀國際 152）", f"=PL!F{PR['B']['ni']}", "=AVERAGE(180.42,152)", "")
    rrow("FY2027 歸屬母公司淨利（情境 B）vs 券商 2027E 平均（東吳 301、交銀國際 366、中郵 430）", f"=PL!G{PR['B']['ni']}", "=AVERAGE(300.95,366,430)", "")
    rrow("FY2027 EPS（情境 B）vs 券商 2027E 平均（0.73、0.89、1.04 元）", f"=PL!G{PR['B']['eps']}", "=AVERAGE(0.73,366/412.83,430/412.83)", "", "0.00")
    rrow("FY2026 毛利率 vs 2026 年 1–9 月實績", f"=PL!F{PR['B']['gm']}", "=gm26", "模型以 1–9 月為錨（差距應為 0）", "0.0%")
    rrow("FY2026 第 4 季隱含營收 vs 2026Q3 實績", f"=PL!F{PR['B']['rev']}-r9m26", "=Q!F4", "第 4 季若維持第 3 季水準即為對照值")
    rrow("加權目標價 vs 交銀國際目標價（2026-10，三季報後）", f"=Val!B{VR['tp']}", "=cons_tp", S_CM + "；只作對照，不進模型")
    rr += 1
    RC.cell(rr, 1, "線 6 相機市場對照（五線歸因：工業機器人、人形、AMR、智慧車、Robotaxi；單價 25 美元為線 6 假設）").font = BOLD; rr += 1
    rrow("FY2026 機器人用感測器營收（人形＋其他機器人，百萬美元）vs 線 6 相機市場 2026", f"=(PL!F{PR['B']['hum']}+PL!F{PR['B']['orob']})/fx", f"=線取數!E{LR['cam_mkt']}",
         "本公司單一公司機器人用營收即高於線 6 全球相機市場，兩者口徑不同：線 6 為人形＋AMR＋Robotaxi 的 25 美元相機模組，未含服務機器人（掃地、清潔、割草等）、深度相機高單價；見對帳", "#,##0.0")
    rrow("FY2030 機器人用感測器營收（情境 B，百萬美元）vs 線 6 相機市場 2030", f"=(PL!J{PR['B']['hum']}+PL!J{PR['B']['orob']})/fx", f"=線取數!I{LR['cam_mkt']}", "同上", "#,##0.0")
    rrow("FY2026 人形隱含全球相機份額（以線 4 基準校準）", f"=Thesis!E{THR['B']['hsh']}", None, "＝人形用營收 ÷（線 4 出貨 × 線 6 每台相機數 × 本公司每台單價）", "0.0%", False)
    rrow("2025 每台感測器營收（元）vs 線 6 相機模組單價 25 美元（元）", "=asp25", f"=25*fx", f"線 6：{LD['p26'][1][1]}（Assumed，線 6 未取得相機價格資料）；本公司為深度相機與多用途平均", "#,##0")
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
    h26 = LD["hum"][(2026, "B")]; h30 = LD["hum"][(2030, "B")]; h32 = LD["hum"][(2032, "B")]
    a26 = LD["amr"][(2026, "B")]; a32 = LD["amr"][(2032, "B")]
    ZR = [
        ("全球人形機器人出貨 FY2026–FY2032（A／B／C＝線高／基準／低）", "4 人形", f"基準 {h26[0]:,.0f} → {h32[0]:,.0f}（× hu_k）", "台", "全球、含小型機", f"{h26[1]}；{h32[1]}", "Assumed", "Thesis 人形用營收", "一致（hu_k＝1）；高案 hu_k＝2 列 Sens 不改基準"),
        ("人形每台相機數", "6 共用零組件", "=n_cam", "顆／台", "人形頭部 2–8 顆", f"{LD['n_cam'][1][1]}（低 {LD['n_cam'][0][1]}、高 {LD['n_cam'][2][1]}）", "Analogy", "Thesis 全球人形相機需求", "一致（直接讀值）"),
        ("AMR 年銷量 FY2026–FY2032（A／B／C＝線高／基準／低）", "6 共用零組件", f"基準 {a26[0]:,.1f} → {a32[0]:,.1f}", "萬台", "運輸物流機器人（IFR 樣本）；本模型只當「服務、物流機器人需求」的成長指標，不直接當出貨量", f"{a26[1]}", "Derived", "Thesis 其他機器人用營收成長", "一致（直接讀值、幾何插值）；用途為代理指標，線 6 未涵蓋掃地、清潔、割草等服務機器人，列為判斷題（drift 吸收差距）"),
        ("相機單價年降幅", "6 共用零組件", "=pd", "%／年", "線 6 相機模組", f"{LD['pd'][1][1]}", "Assumed", "Thesis 每台單價路徑", "一致（直接讀值）；線 6 此列為 Assumed，未取得相機單價資料"),
        ("相機單價水準", "6 共用零組件", "=asp25", "元／台", "本公司每台感測器營收（Derived，含全部用途）", f"線 6：25 美元（約 178 元，{LD['p26'][1][1]}）", "Derived", "Thesis 單價起點", "不一致（刻意）：本公司每台 329 元（約 46 美元）為深度相機與多用途平均，採本公司揭露推得的水準，只讀線 6 的降幅"),
        ("線 6 相機模組市場規模 2026／2030（基準）", "6 共用零組件", f"{LD['cam_mkt'][2026][0]:,.1f}／{LD['cam_mkt'][2030][0]:,.1f}", "百萬美元", "五線歸因；25 美元單價", f"{LD['cam_mkt'][2026][1]}", "Derived", "Recon 對照（不進公式）", "不一致（量級）：本公司 2026 年機器人用感測器營收約 25–30 百萬美元，高於線 6 全球相機市場，原因見 Recon；建議線 6 下次改版補深度相機與服務機器人"),
        ("人形占相機市場比重 2030", "6 共用零組件", "約 51%（人形，基準）", "%", "線 6 計算頁 F245÷L245", f"{L6N} 計算!F245、L245", "Derived", "只作定性對照", "一致"),
        ("工業機器人（線 3）", "3 工業機器人", "未使用", "—", "線 6 每台工業機器人「機器視覺」美元內容 500 美元（Assumed）", f"{L6N} 輸入!E58", "Assumed", "未使用", "本公司工業級應用設備僅 2.7% 營收，不建模此驅動"),
    ]
    for i, z in enumerate(ZR):
        for j, v in enumerate(z): DZ.cell(4 + i, 1 + j, v)
    DZ.column_dimensions["A"].width = 50; DZ.column_dimensions["F"].width = 80; DZ.column_dimensions["I"].width = 70

    # ================= Risk =================
    RK["A1"] = "Risk：風險登錄（原則第 10 節第 10、19 條）。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    RISKS = [
        ("客戶集中：2025 年第一大客戶 30.0%、第二大（關聯方，螞蟻體系）26.8%，前五大合計 63.9%", "部分（gcons_A／C）", "消費級應用設備（三維掃描、支付『碰一下』設備）占營收 62%，大客戶訂單縮減即營收下降；2026 年 1–9 月營收持平即與此相關", "年報前五大客戶；消費級設備營收"),
        ("機器人用營收未揭露（只知『翻倍』）", "是（rob_sh26 區間）", "機器人占比、人形占比皆 Assumed，結論對此敏感", "投資者關係紀錄；H 股招股書（若遞表）"),
        ("研發與管理費用上升侵蝕利潤", "是（opex26、opex30）", "2026 年 1–9 月歸母淨利 −32%", "季費用率"),
        ("競爭與替代（Intel RealSense、舜宇、禾賽、純視覺路線）", "部分（hs30_C、gm30_C）", "", "客戶導入名單；季毛利率"),
        ("H 股發行（規模、價格未揭露）", "部分（h_iss 尾端 Sens，基準 0）", "高於內在價值發行可能增加每股價值", "聯交所遞表"),
        ("金融資產明細未取得（約 1.2 十億元長期金融資產為估計）", "部分（inv_k 低案 50%）", "淨現金占市值約 10%", "三季報原文資產負債表"),
        ("採集設備新品線需求未驗證", "部分（q4g 區間 0–75%）", "『上市兩個月逾百家客戶下單』為公司說法", "第 4 季與 2027 年一季報"),
        ("所得稅率偏低（累計虧損可抵、高新企業）", "部分（tax_r 5–15%）", "", "年報所得稅附註"),
    ]
    for i, k in enumerate(RISKS):
        for j, v in enumerate(k): RK.cell(4 + i, 1 + j, v)
    RK.column_dimensions["A"].width = 60; RK.column_dimensions["D"].width = 60

    # ================= Sources =================
    SO["A1"] = "Sources：本模型引用的原始來源與資料商（含網址；HTML 展示層不放網址）。"
    for i, s_ in enumerate([S_PX, S_AR, S_DW, S_CM, S_ZY, S_TJ, S_SA, S_SAP, S_RF, S_CRP,
                            f"線 4：{L4N}", f"線 6：{L6N}",
                            "physical_ai/map/profiles/688322.md（背景，數字已另行查證）"], 3):
        SO.cell(i, 1, s_)
    SO.column_dimensions["A"].width = 160

    # ================= README（xlsx 內）=================
    S = T
    v = lambda k: f"Val!B{VR[k]}"
    lines = [
        ('="奧比中光（688322.SS）個股模型 ' + VER + '｜build_id "&Sens!D2&"｜資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'),
        "命題：奧比中光的股價以機器人視覺（含人形）選擇權定價，但 2025 年營收增長主要來自三維掃描與支付設備，機器人用營收未揭露；本模型把營收拆為機器人用（人形：線 4 出貨 × 線 6 每台相機數 × 份額 × 單價；其他機器人：線 6 AMR 銷量成長 ×（1＋超額））與非機器人用，並檢驗現價隱含多少機器人成長。",
        (f'="結論（條件式）：在 FY2026 以 1–9 月實績為錨（第 4 季年增 "&{S("q4g","0%")}&"）、線 4 基準人形出貨（× "&{S("hu_k","0.0")}&"）、線 6 每台 "&{S("n_cam","0")}&" 顆相機與單價年降 "&{S("pd","0%")}&"、同業 P/E 未加溢價（相對倍數係數 "&{S("rel_k","0.00")}&"）的假設下，加權目標價約 "&{S(v("tp"),"#,##0.0")}&" 元，相對現價 "&{S(v("up"),"+0.0%;-0.0%")}&"，估值評等「"&{v("rating")}&"」。"'),
        (f'="DCF 目標價 "&{S(v("dcf_tp"),"#,##0.0")}&"、Comparable 目標價 "&{S(v("cmp_tp"),"#,##0.0")}&"（各半）；WACC "&{S("wacc","0.00%")}&"、ke "&{S("ke","0.00%")}&"、g "&{S("g_tv","0.00%")}&"；情境 B 終值占 EV "&{S(f"Val!C{VR['tvs']}","0%")}&"。"'),
        (f'="上檔空間歸因：DCF 貢獻 "&{S(v("att_d"),"+0.0%;-0.0%")}&"、未調整 Comparable 貢獻 "&{S(v("att_c"),"+0.0%;-0.0%")}&"、相對倍數調整貢獻 "&{S(v("att_r"),"+0.0%;-0.0%")}&"。本公司現價對資料商前瞻 EPS 的倍數是同業中位數（僅 3 家）的 "&{S(v("relnow"),"0.0")}&" 倍；支持現價需要相對倍數係數 "&{S(v("relreq"),"0.0")}&"。"'),
        (f'="收支關卡：情境 B FY2026–FY2030 FCFF 扣股利合計 "&{S(v("gate_b"),"#,##0.0")}&" 百萬元，"&{v("gate_t")}&"；距最近級距邊界 "&{S(v("dist"),"0.0%")}&"。"'),
        "主要失效條件（可觀察）：(1) 第 4 季與 2027 年營收接近券商預測（年增 70% 以上）；(2) 機器人用營收占比公開且高於 40%；(3) 毛利率維持 49% 以上且費用率下降；(4) H 股高價發行。",
        "方法：DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，WACC 折現，明確期到 FY2032，RONIC＝WACC；金融資產以淨現金計入）50%＋同業前瞻 P/E × 模型 FY2027 EPS 50%；三情境 25／50／25。",
        "限制：機器人用／非機器人用拆分為 Assumed（公司未揭露）；2026 年分產品未揭露；同業僅 3 家，P/E 由原報價日按 2026-11-20 收盤等比例換算（EPS 未更新）；2023 年損益細項未取得；不讀 Siliconomics／Tokenomics／Wattonomics。詳見 repo README.md。",
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
    for k, lab in (("hs30", "2030 年人形相機份額"), ("drift", "其他機器人用年超額成長"), ("gm30", "FY2030 毛利率"),
                   ("gcons", "消費級設備年增"), ("gnr", "非機器人用感測器年增"), ("gind", "工業級及其他年增")):
        for t_ in "AC":
            extra.append((f"情境 B 的{lab}改為情境 {t_} 值", f"{k}_B", ivE[f"{k}_{t_}"]))
    extra.append(("人形出貨為線 4 的 5 倍（全部情境）", "hu_k", 5.0))
    for drv, k, val in extra:
        specs.append(dict(drv=drv, set=f"{k}＝{val:.4g}", key=k, val=val, ov={k: val}))
    specs.append(dict(drv="情境權重全給 A", set="tp_wA＝1、tp_wB＝0、tp_wC＝0", key="tp_wA tp_wB tp_wC", val=None, ov={"tp_wA": 1, "tp_wB": 0, "tp_wC": 0}))
    specs.append(dict(drv="情境權重全給 C", set="tp_wA＝0、tp_wB＝0、tp_wC＝1", key="tp_wA tp_wB tp_wC", val=None, ov={"tp_wA": 0, "tp_wB": 0, "tp_wC": 1}))
    res = run_cases([s["ov"] for s in specs], meta, "sens")
    cases = []
    for s, r in zip(specs, res):
        cases.append(dict(drv=s["drv"], set=s["set"], key=s["key"], val=s["val"], tp=r["tp"], dcf=r["dcf"], rating=r["rating"]))
        print(f"  {s['drv'][:44]:44s} {r['tp']:9.1f} {r['tp'] / base['tp'] - 1:+.1%}")

    px = ivE["px"]
    def solve(key, target_tp, lo, hi, n=10, over=None):
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
    s1 = solve("hu_k", px, 1.0, 200.0)
    s2 = solve("drift_B", px, ivE["drift_B"], 20.0)
    s3 = solve("drift_B", flip_tp, ivE["drift_B"], 20.0)
    nf = lambda x: x if x else -1
    solve_rows = [
        ("支持現價（目標價＝現價）所需的人形出貨倍數（相對線 4，全部情境）", nf(s1), 1.0, "其他輸入不變；−1＝200 倍內無解" if not s1 else f"其他輸入不變；情境 B 的 CY2030 全球人形出貨相當於 {LD['hum'][(2030, 'B')][0] * s1 / 1e6:,.1f} 百萬台"),
        ("支持現價所需的情境 B 其他機器人用感測器年超額成長（量，相對 AMR 銷量成長）", nf(s2), ivE["drift_B"], "其他輸入不變；−1＝20（即 2,000%／年）內無解"),
        (f"評等升一級（上檔 {flip:+.0%}）所需的情境 B 其他機器人用年超額成長", nf(s3), ivE["drift_B"], "其他輸入不變；−1＝20（即 2,000%／年）內無解"),
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
    segs = [("3D 視覺感測器", "3D 視覺感測器 營收", Hs["E8"].value), ("消費級應用設備", "消費級應用設備 營收", Hs["E9"].value),
            ("工業級應用設備及其他（含其他業務收入，殘差）", "工業級應用設備及其他 營收", Hs["E10"].value)]
    json.dump(dict(company="orbbec", ticker="688322.SS", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=round(v, 2), unit="百萬元", source=S_AR) for n, l, v in segs],
                   sublines=[dict(parent="3D 視覺感測器", labels=["感測器：人形機器人用（子列）", "感測器：其他機器人用（子列）", "感測器：非機器人用（子列）"])],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）", "檢查：感測器子列合計 − 感測器分部（應為 0）"],
                   note="官方分部＝年報主營業務分產品（3D 視覺傳感器 290.93、消費級應用設備 583.73、工業級應用設備 25.54、其他 34.84）；工業級與其他及其他業務收入合併為一列（殘差）。用途子列（人形、其他機器人、非機器人）為命題拆分（Assumed），合計＝感測器分部。"),
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
    t = f"""# physical_ai/orbbec/ — 奧比中光 Orbbec（688322.SS）個股模型 {VER}

<!-- 本檔由 build_orbbec.py 依成品重算值產生；數字不手寫（原則第 10 節第 51 條）。 -->

**本模型適用《高不確定性預估工作原則》**（`docs/plan/高不確定性預估工作原則.md`）：機器人用營收與人形占比公司未揭露、人形需求尚未發生，相關輸入一律給區間並標 Assumed／Analogy；結論只在證據邊界內成立。

- 成品：`dist/{NAME}`（build_id `{sens['build_id']}`；LibreOffice 重算，0 錯誤值）；HTML 展示層 `dist/{NAME[:-5]}.html`（`gen_html_json.py` 產生 `html.json`，再以 `tools/company_html/build_company_html.py` 產生）
- 建置：`python3 physical_ai/orbbec/build_orbbec.py`（需要 soffice；整本重算敏感度與反解；結尾自動跑 in_layout、sens_coverage --require-keycol、model_governance、rating_gate --strict、zero_rows、cached_values、segment_check）
- 日期：資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 2026-11-20（收盤 {I['px']:,.2f} 元）｜評價日 2026-12-31｜目標時點 2027-11-20
- 會計年度＝日曆年；單位百萬元人民幣；中國會計準則（法定財報口徑，GAAP）。上海科創板。

## 結論（條件式）

在「FY2026 以 1–9 月實績為錨（第 4 季年增 {I['q4g']:.0%}）、線 4 基準人形出貨、線 6 每台 {I['n_cam']:.0f} 顆相機與單價年降 {I['pd']:.0%}、同業 P/E 不加溢價（相對倍數係數 {I['rel_k']:.2f}）」的假設下：

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
| 券商目標價（交銀國際，2026-10；只作對照） | {I['cons_tp']:,.2f} |

上檔空間歸因：DCF 貢獻 {g('att_d'):+.1%}、未調整 Comparable 貢獻 {g('att_c'):+.1%}、相對倍數調整 {g('att_r'):+.1%}。本公司現價對資料商前瞻 EPS 的倍數是核心同業中位數（{I['pe_peer']:.1f} 倍，僅 3 家）的 {g('relnow'):.1f} 倍；要讓目標價＝現價，Comparable 的相對倍數係數須達 {g('relreq'):.1f}。DCF 以 RONIC＝WACC（新投資無超額報酬），高相對倍數隱含持續超額報酬，兩法對競爭優勢的假設相反。

**讀法：** 市場以機器人視覺選擇權定價，但 2025 年營收增長（+67%）主要來自三維掃描與支付『碰一下』設備（消費級應用設備占 62%），2026 年 1–9 月營收持平（−0.4%）、歸母淨利 −32%；機器人用營收未揭露（只知 2025 年國內業務翻倍、2026 年 1–9 月視覺方案再翻倍）。情境 B 機器人用占合併營收由 {TH[f'E{THB["robsh"]}'].value:.0%}（FY2026）升到 {TH[f'K{THB["robsh"]}'].value:.0%}（FY2032），其中人形用只占 {TH[f'E{THB["humsh"]}'].value:.0%} → {TH[f'K{THB["humsh"]}'].value:.0%}。整本重算反解：其他輸入不變時，**支持現價需要人形出貨達線 4 的 {fn(sol[0][1])} 倍**（{sol[0][3]}）；或其他機器人用感測器每年超額成長達 {fn(sol[1][1], '{:,.0%}')}（基準 {fn(sol[1][2], '{:,.0%}')}）。結論主要建立在 Assumed（用途拆分、份額、毛利率與費用率路徑）與線模型 Analogy 輸入上；方向是否「穩健」看敏感度表。

## 命題

1. **營收結構：** 官方產品分部為 3D 視覺感測器（2025 年 {wb['Hist']['E8'].value:,.0f} 百萬元，占 {wb['Hist']['E8'].value / wb['Hist']['E4'].value:.0%}）、消費級應用設備（{wb['Hist']['E9'].value:,.0f}，占 {wb['Hist']['E9'].value / wb['Hist']['E4'].value:.0%}）、工業級應用設備及其他；用途別（機器人／非機器人、人形）公司未揭露，本模型以 Assumed 拆分（2026 年機器人用占感測器 {I['rob_sh26']:.0%}，區間 40–70%；人形占機器人用 {I['hum_sh26']:.0%}，區間 3–20%）。
2. **量：** 人形用＝線 4 全球出貨 × 線 6 每台相機數 × 份額（2026 年隱含份額 {TH[f'E{THB["hsh"]}'].value:.0%}，以線 4 基準校準）；其他機器人用＝前一年 × 線 6 AMR 銷量成長 ×（1＋超額）。
3. **價：** 每台感測器營收 2025 年 {I['asp25']:,.0f} 元（約 46 美元；線 6 相機單價假設 25 美元，口徑不同），依線 6 每年降 {I['pd']:.0%}。

情境 B：營收 {ser('rev', '{:,.0f}')}；毛利率 {ser('gm', '{:.1%}')}；EPS {ser('eps', '{:.2f}')} 元（FY2026→FY2032）。

## 主要假設與可證偽指標（原則第 10 節第 20 條）

| 承重假設 | 基準（A／B／C 或區間） | 標記 | 可觀察指標 |
|---|---|---|---|
| 2026 年第 4 季營收年增 | {I['q4g']:.0%}（0–75%；券商隱含約 +72% 以上） | Assumed | 2026 年年報、2027 年一季報 |
| 機器人用占感測器營收（2026） | {I['rob_sh26']:.0%}（40–70%） | Assumed | 投資者關係紀錄、H 股招股書 |
| 其他機器人用年超額成長 | {I['drift_A']:.0%}／{I['drift_B']:.0%}／{I['drift_C']:.0%} | Assumed | 機器人視覺營收揭露 |
| 人形每台相機數（線 6） | {I['n_cam']:.0f}（2–8） | Analogy | 整機拆解 |
| 2030 年人形相機份額 | {I['hs30_A']:.0%}／{I['hs30_B']:.0%}／{I['hs30_C']:.0%} | Assumed | 人形客戶導入名單 |
| FY2030 毛利率 | {I['gm30_A']:.0%}／{I['gm30_B']:.0%}／{I['gm30_C']:.0%} | Assumed | 季毛利率（2026 年 1–9 月 49.1%） |
| FY2030 費用率 | {I['opex30']:.0%}（22–34%） | Assumed | 季費用率（2026 年 1–9 月 45.0%） |
| 同業前瞻 P/E | 中位數 {I['pe_peer']:.1f} 倍（3 家） | Derived | 同一估值日重抓 |

## 敏感度（整本重算，前 10 大）

| 案例 | 鍵 | 加權目標價 | 相對基準 | DCF 目標價 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | `{c['key']}` | {c['tp']:,.1f} | {c['tp'] / sens['base_tp'] - 1:+.1%} | {c['dcf']:,.1f} | {c['rating']} |" for c in top) + f"""

完整案例見 xlsx Sens 頁（每個 In 頁有低／高區間的鍵都有案例；另有情境 B 改 A／C 值、人形出貨 5 倍、情境權重案例）。**翻轉評等所需的變動**：見 Recon 反解；升一級需加權目標價達 {fn(up_b)} 元。

## 反解（Recon 頁）

""" + "\n".join(f"- {a}：{fn(b, '{:,.2f}')}（基準 {fn(c0, '{:,.2f}')}）；{d}" for a, b, c0, d in sol) + f"""

## 限制與「分不出來／沒找到」

- **分不出來（資料性質）：** 機器人用營收與人形用營收公司未揭露；2026 年分產品未揭露；人形量級屬未發生事件。
- **沒找到（工作未完）：** 2026 年三季報原文與資產負債表明細未直接讀取（三季數字取自券商與資料商轉錄；現金 1,300 為資料商值，長期金融資產 1,200 為券商預測表估計）；2023 年損益細項；2026 年限制性股票；H 股發行規模與價格；2024、2025 年工業級與其他的細項分開；人形客戶名單。
- **線模型不一致：** 線 6 全球相機市場（2026 年約 16 百萬美元）小於本公司單家機器人用感測器營收（約 25–30 百萬美元），因線 6 為人形＋AMR＋Robotaxi 的 25 美元相機模組、未含服務機器人與深度相機高單價；本模型只讀線 6 的降幅、每台相機數、AMR 銷量路徑，單價水準採本公司揭露。見 xlsx「對帳」頁。
- **同業倍數：** 核心同業僅 3 家（舜宇、韋爾、LG Innotek），資料商頁報價日 2026-11-24，前瞻 P/E 依 2026-11-20 收盤等比例換算（EPS 未更新）；立訊精密、禾賽、速騰聚創、Intel RealSense、思特威「報價日未取得」或虧損、非上市，未進中位數；少於 8 家不做迴歸。
- **現價與資料商報價差：** 資料商頁面顯示 2026-11-23 報價 101.95 元，與收盤檔 2026-11-20 的 85.09 元差 20%；本模型一律用收盤檔。
- **rf 逾期：** 中國 10 年期國債 1.61% 為 2026-11-05 值（跨模型一致）。
- **協調檔：** 實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics；相關條文不適用。
- **獨立核對：** 依專案規則，交付判讀的數字應由未參與製作的代理獨立核對；本代理未安排，待外層流程。

## 資料來源

見 xlsx Sources 頁（公司 2025 年年度報告、券商研報、資料商、線 4／6 成品）。

## 版本紀錄

- {VER}（2026-11-24）：首版。現價 {I['px']:,.2f}（2026-11-20 收盤）；加權目標價 {g('tp'):,.1f}；上檔空間 {g('up'):+.1%}；{g('rating')}／{g('rec')}；收支關卡 {g('gate_b'):,.1f}（{g('gate_t')}）。
"""
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write(t)


if __name__ == "__main__":
    main()

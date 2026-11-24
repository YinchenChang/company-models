# 五洲新春 Wuzhou Xinchun（603667.SS，上海主板）個股模型 v0.1 建置腳本
# 產出：physical_ai/wuzhou/dist/20261124_五洲新春個股模型_v0.1.xlsx（LibreOffice 重算後的成品，含公式與快取值）
# 結構：命題模組（傳統本業＝軸承＋汽車零部件＋熱管理，以 FY2026 前三季實績為錨、之後給成長區間；
#       機器人零部件＝線 4 全球人形滾柱絲杠需求 × 本公司份額 × 單價（線 6 年降幅））→ 標準損益表 → FCFF DCF＋同業 P/E（另加非營運淨資產，含洛軸股份持股）。
# 敏感度：每個案例覆寫 In 頁基準欄後整本重建、LibreOffice 重算（整本重算，不用閉式近似）；反解同法以二分搜尋。
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md、docs/plan/高不確定性預估工作原則.md。
#       實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics 協調檔。
# 用法：python3 physical_ai/wuzhou/build_wuzhou.py   （需要 soffice；暫存於家目錄下 wz_build）
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
NAME = f"20261124_五洲新春個股模型_{VER}.xlsx"
OUT = os.path.join(HERE, "dist", NAME)
HOMEDIR = os.path.expanduser("~")
TMP = os.path.join(HOMEDIR, "wz_build")
LO = ["soffice", f"-env:UserInstallation=file://{HOMEDIR}/lo_profile", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML"]

L4 = os.path.join(ROOT, "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx")
L6 = os.path.join(ROOT, "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx")
L4N, L6N = (os.path.relpath(p, ROOT) for p in (L4, L6))

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDRF = PatternFill("solid", fgColor="DDEBF7")
SECF = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------- 出處 ----------------
S_PX = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv：603667.SS；未還原股利）"
S_CLOSE = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
S_SINA = "新浪財經 603667 合併利潤表（資料商；https://vip.stock.finance.sina.com.cn/corp/go.php/vFD_ProfitStatement/stockid/603667/ctrl/part/displaytype/4.phtml 及 ctrl/2024、ctrl/2023 頁，2026-11-24 讀取；單位萬元；2026 年各欄為年初至今累計；公司年報與三季報原文未直接讀取）"
S_SA = "stockanalysis.com SHA:603667（資料商；財務摘要、資產負債表季頁、現金流量表、statistics 頁，2026-11-24 讀取）"
S_J = "同花順 F10 財務摘要（資料商；https://basic.10jqka.com.cn/603667/finance.html，2026-11-24 讀取）：各期營收、毛利率、每股收益"
S_N = "同花順財經 2026-10-27〈投資洛軸股份「吃飽」五洲新春第三季度歸屬淨利潤增長超 3875%〉（https://stock.10jqka.com.cn/20261027/c681155613.shtml）：前三季營收 22.50 億元（−15.45%）、第 3 季營收 7.16 億元（−6.59%）、持有洛軸股份 50,983,200 股（7.04%，第四大股東）"
S_FP = "財報透視（東方財富，2026-10-28；https://wap.eastmoney.com/a/202610283911420312.html）：以搜尋摘要轉述，全文未讀；前三季歸母淨利 9.83 億元（+898%）、扣非 7,044 萬元（−22.9%）、非經常性損益 9.12 億元、公允價值與處置損益 10.69 億元；第 3 季扣非年減 34.88%（美元貶值匯兌損失、子公司虹新製冷業績下滑）"
S_SX = "山西證券研究報告（格隆匯轉載 https://m.gelonghui.com/news/5239848；轉載日期未能確認）：2025 年營收 33.43 億元（+2.41%）、歸母 0.91 億元；軸承 +1.95%、汽車零部件 +8.68%（毛利率 23.98%）；機器人零部件 2025 年收入 1,599.5 萬元（首次出現）；其預測（2026E 營收 38.65 億、淨利 1.43 億）早於三季報，不採用"
S_LZ = "洛軸股份（301699）2026-11-06 收盤 31.78 元（新浪財經／證券之星報導，搜尋摘要；不在本專案收盤檔，只作持股市值對照，不進模型公式）"
S_RF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與 innolight、eoptolink、victory_giant、leaderdrive 同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：China CRP 0.91%（與 innolight 同一取值）"
S_SAP = "stockanalysis.com statistics 頁（資料商；2026-11-24 讀取，頁面報價日見 Comp 頁 M 欄）"
S_HX = "physical_ai/hiwin 個股模型 Comps 頁（THK 前瞻 P/E，資料日 2026-11-20）"
S_LD = "physical_ai/leaderdrive 個股模型 Comp 頁（雙環傳動、拓普集團前瞻 P/E 與原報價）"
S_HW = "humanoid.guide（產業網站，經線 6 追蹤指標頁轉述）：五洲新春 2025 年行星滾柱絲杠出貨約 8.7 萬根、均價約 1,360 元人民幣、年減 39%"
IP_BRK = "券商研究報告：有銷售研究與維持評等的誘因；歷史數字為轉錄，預測為券商自設"
IP_SA = "資料商彙整，財報前的值未必更新"
IP_CO = "產業網站轉述，未經公司證實"


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
    d = {"scr": {}}
    for y in range(2026, 2033):
        r, v = l4[f"SCREW_UNITS_{y}"]
        for s, c, col in (("A", 4, "E"), ("B", 3, "D"), ("C", 2, "C")):
            d["scr"][(y, s)] = (v[c], f"{L4N} 下游取數!{col}{r}（SCREW_UNITS_{y}）")
    ws6 = wb6["輸入"]
    pr = [r for r in range(1, ws6.max_row + 1) if ws6.cell(r, 1).value == "PRS" and ws6.cell(r, 3).value == "美元/根"][0]
    L = openpyxl.utils.get_column_letter
    d["pd"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{L(c)}{pr}") for c in (7, 8, 9)]
    d["p26"] = [(ws6.cell(pr, c).value, f"{L6N} 輸入!{L(c)}{pr}") for c in (4, 5, 6)]
    r, v = l6["行星滾柱絲杠需求量 2030"]; d["dem30_6"] = (v[2] * 1e4, f"{L6N} 下游取數!C{r}")
    r, v = l6["行星滾柱絲杠市場規模 2030（基準單價口徑＝覓途）"]; d["mkt30_6"] = (v[2], f"{L6N} 下游取數!C{r}")
    r, v = l6["行星滾柱絲杠 2030：基準需求 × 單價口徑（中國量產／覓途／MS）"]; d["mkt30_cn"] = (v[1], f"{L6N} 下游取數!B{r}")
    r, v = l4["Q_SCREW_AVG"]; d["q_scr"] = (v[3], f"{L4N} 下游取數!D{r}（Q_SCREW_AVG）")
    r, v = l4["SCREW_MKT_2030"]; d["mkt30_4"] = (v[3], f"{L4N} 下游取數!D{r}（SCREW_MKT_2030）")
    return d


LD = line_data()

# ---------------- In 頁輸入 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 來源誘因, 說明)
IN = [
    ("sec", "1. 市場資料與日期"),
    ("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_PX, "", "資料截止日 2026-11-24（財報、同業倍數）；股價 2026-11-20；成品日 2026-11-24"),
    ("px", "現價（2026-11-20 收盤）", "元", None, 41.60, None, "Verified", S_PX, "", "csv 值 41.599998；11-09 至 11-20 全球 AI 與半導體股大跌，股價下跌非資料錯誤"),
    ("val_date", "評價日", "日期", None, "=DATE(2026,12,31)", None, "Assumed", "本專案慣例：評價日 2026-12-31（會計年度＝日曆年，FY2026 在評價日已結束）", "", "折現期間由評價日起算"),
    ("tp_date", "目標時點（px_date 後 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", ""),
    ("bs_date", "資產負債表日", "日期", None, "=DATE(2026,9,30)", None, "Verified", S_SA, "", "2026 年三季報"),
    ("stub_frac", "Stub 期（資產負債表日→評價日）", "年", None, "=(val_date-bs_date)/365", None, "Derived", "＝(評價日 − 資產負債表日) ÷ 365", "", "Stub FCFF＝FY2026 FCFF × stub_frac"),
    ("sh_v", "評價日股數（總股本）", "百萬股", None, 383.15, None, "Verified", S_SA + "：總股本 383.15 百萬股（2026-06-30、09-30 相同；2026-03-31 為 366.2）", "", "2026 上半年有增資；前三季加權股數約 375 百萬股（歸母 983 ÷ EPS 2.62）；限制性股票與選擇權未取得，以 0 計"),
    ("cash_v", "現金＋交易性金融資產（2026-09-30）", "百萬元", None, 956.54, None, "Verified", S_SA + "：貨幣資金 859.54＋交易性金融資產 97", "", "2026-06-30 為 1,166"),
    ("ltinv_v", "長期投資（2026-09-30；含洛軸股份持股的公允價值）", "百萬元", None, 1477.0, None, "Verified", S_SA + "：長期投資 1,477（2026-06-30 為 344.9，增加主要為洛軸股份公允價值）", "", "洛軸股份 50.98 百萬股（7.04%）；明細未取得，見 Recon 市值對照"),
    ("inv_k", "長期投資計入淨現金的比例", "%", 0.6, 0.85, 1.0, "Assumed", "持股為創業板新上市公司股票，可能有鎖定期與出售稅負（公允價值利得已計遞延所得稅，三季報所得稅 1.60 億元）；基準折 15%，低案 60%", "", "洛軸股份價格波動大，見 Recon"),
    ("debt_v", "有息負債（2026-09-30）", "百萬元", None, 678.19, None, "Verified", S_SA + "：短期借款等 476.77、長期 5，合計 678.19", "", "2026-06-30 為 804.26"),
    ("nci_v", "少數股東權益（2026-09-30）", "百萬元", None, 54.36, None, "Verified", S_SA, "", "自淨現金扣除"),
    ("fx", "美元兌人民幣（2026-11-20）", "元／美元", None, 6.6665, None, "Verified", S_CLOSE + "：CNY=X", "", "只用於 Recon 與線 6 市場規模對照"),
    ("lz_sh", "洛軸股份持股數", "百萬股", None, 50.9832, None, "Verified", S_N, "", "只用於 Recon 市值對照"),
    ("lz_px", "洛軸股份 2026-11-06 收盤", "元", None, 31.78, None, "Interested-party", S_LZ, IP_SA, "只用於 Recon 市值對照；收盤檔沒有洛軸股份，不進模型"),
    ("sec", "2. 實績錨（FY2026＝2026 年，會計年度＝日曆年）"),
    ("r9m26", "2026 年 1–9 月營收", "百萬元", None, "=Q!F4", None, "Derived", "＝Q 頁 1–9 月營收（" + S_SINA + "）", "", "同花順報導 22.50 億元（−15.45%）一致"),
    ("g9m26", "2026 年 1–9 月營收年增", "%", None, -0.1545, None, "Verified", S_N, "", "2025 年 1–9 月營收 2,660.90 百萬元（新浪）"),
    ("q4g", "2026 年第 4 季營收年增", "%", -0.15, -0.066, 0.03, "Assumed", "基準＝第 3 季單季年增 −6.59%（" + S_N + "）；低案＝上半年年減 19%（Q 頁）與第 3 季之間；高案＝回到小幅成長", "", "FY2026 營收＝1–9 月實績＋2025Q4 ×（1＋q4g）"),
    ("rob26", "FY2026 機器人零部件營收", "百萬元", 25.0, 45.0, 90.0, "Assumed", "公司未單獨揭露 2026 年機器人零部件營收；2025 年 16.0 百萬元（" + S_SX + "），2026 年券商稱絲杠交付節奏為觀察重點；2026 年事件：入股具身智能基金、成立智驅科技公司（profile，外部新聞）", "", "三情境同一實績年；隱含份額見 Thesis"),
    ("gm26", "FY2026 毛利率", "%", None, "=1-Q!F5/Q!F4", None, "Derived", "＝1 − 前三季營業成本 ÷ 營收（Q 頁；19.10%，與同花順摘要一致）", "", "第 4 季假設同 1–9 月"),
    ("opex26", "FY2026 銷售＋管理＋研發費用率", "%", 0.135, "=Q!F6/Q!F4", 0.155, "Derived", "＝前三季（銷售＋管理＋研發）÷ 營收（Q 頁）", "", "2025 年 12.8%；2026 年前三季管理費用明顯增加"),
    ("sec", "3. 命題模組：傳統本業與機器人零部件"),
    ("pdr", "滾柱絲杠單價年降幅（線 6）", "%／年", LD["pd"][0][0], LD["pd"][1][0], LD["pd"][2][0], "Analogy", f"線 6：{LD['pd'][1][1]}（低 {LD['pd'][0][1]}、高 {LD['pd'][2][1]}）", "",
     "本公司 2025 年均價年減 39%（" + S_HW + "；Interested-party）高於線 6 高案 20%，見對帳"),
    ("asp25", "2025 本公司滾柱絲杠均價", "元人民幣／根", None, 1360.0, None, "Interested-party", S_HW, IP_CO, "線 6 以此為「中國量產價」口徑；與雙林報價 499 元並列"),
    ("hu_k", "人形絲杠需求相對線 4 的倍數（全部情境）", "x", None, 1.0, 2.0, "Assumed", "基準＝線 4 原值；高案 2 倍為任意倍數，無證據支持（線 4 高情境已涵蓋樂觀路徑）", "", "反解變數"),
    ("sec", "4. 命題模組：情境專屬驅動（A 樂觀／B 基準／C 保守；只列基準欄，不設敏感度區間）"),
    ("gcore_A", "情境 A：傳統本業營收年增（FY2027 起）", "%", None, 0.05, None, "Assumed", "2023→2025 營收年複合約 3.7%（Hist）；熱管理回穩、線控轉向與制動絲杠螺母（辰致、上海大陸、吉利定點，" + S_SX + "）", "", "不含機器人零部件"),
    ("gcore_B", "情境 B：傳統本業營收年增（FY2027 起）", "%", None, 0.03, None, "Assumed", "同上；2025 年 +2.4%、2026 年前三季 −15%（上半年 −19%，熱管理受大客戶影響，見 profile 券商點評）", "", ""),
    ("gcore_C", "情境 C：傳統本業營收年增（FY2027 起）", "%", None, 0.0, None, "Assumed", "同上；熱管理續弱與軸承價格競爭", "", ""),
    ("sh30_A", "情境 A：2030 年占全球人形滾柱絲杠需求份額", "%", None, 0.20, None, "Assumed", "2026 隱含份額見 Thesis；公司稱已有頭部新勢力客戶配套（" + S_SX + "）", "", "2026 隱含份額線性走到 2030，其後持平"),
    ("sh30_B", "情境 B：2030 年占全球人形滾柱絲杠需求份額", "%", None, 0.12, None, "Assumed", "同上；新劍傳動據報建 100 萬根／年產線、雙林 2026-01 宣布量產單價 499 元（線 6 追蹤指標頁），價格競爭稀釋份額", "", ""),
    ("sh30_C", "情境 C：2030 年占全球人形滾柱絲杠需求份額", "%", None, 0.05, None, "Assumed", "同上；國際大廠（Schaeffler 等）與新進者搶單", "", ""),
    ("gm30_A", "情境 A：FY2030 毛利率", "%", None, 0.21, None, "Assumed", "2023–2025 年 17.5%、16.5%、18.05%；2026 年前三季 19.1%、第 1 季 19.8%（" + S_J + "；" + S_SX + "）", "", "FY2026 起線性走到 FY2030，其後持平"),
    ("gm30_B", "情境 B：FY2030 毛利率", "%", None, 0.19, None, "Assumed", "同上；基準＝持平於 2026 年前三季", "", ""),
    ("gm30_C", "情境 C：FY2030 毛利率", "%", None, 0.17, None, "Assumed", "同上；回到 2023–2024 年水準", "", ""),
    ("tp_wA", "情境權重 A", "", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", ""),
    ("tp_wB", "情境權重 B", "", None, 0.50, None, "Assumed", "同上", "", ""),
    ("tp_wC", "情境權重 C", "", None, 0.25, None, "Assumed", "同上", "", ""),
    ("sec", "5. 損益與現金流"),
    ("ts_r", "稅金及附加 ÷ 營收", "%", None, "=Hist!E19/Hist!E4", None, "Derived", "＝2025 年稅金及附加 ÷ 營收（" + S_SINA + "）", "", ""),
    ("opex30", "FY2030 銷售＋管理＋研發費用率", "%", 0.125, 0.135, 0.145, "Assumed", "2023–2025 年 6.9%→10.5%→12.8%（Hist）、2026 年前三季 14.7%；基準假設費用率自 FY2026 回落至 13.5%", "", "FY2026→FY2030 線性，其後持平"),
    ("oth_r", "其他營業損益（其他收益 − 減值等）÷ 營收", "%", -0.008, "=Hist!E22/Hist!E4", 0.004, "Derived", "＝2025 年殘差 ÷ 營收（Hist 第 22 列；殘差含新浪頁未列的其他收益與減值）", "", "計入營業利益"),
    ("fin_r", "財務費用 ÷ 營收（業外，扣除）", "%", 0.004, "=Hist!E17/Hist!E4", 0.010, "Derived", "＝2025 年財務費用 ÷ 營收（" + S_SINA + "）", "", "業外；有息負債 678 百萬元；DCF 另以淨現金計入"),
    ("tax_r", "有效稅率", "%", 0.08, "=SUM(Hist!C33:E33)/SUM(Hist!C32:E32)", 0.15, "Derived", "＝2023–2025 年所得稅合計 ÷ 稅前利益合計（高新技術企業 15%、研發加計扣除）", "", "敏感度低／高 8%／15%（Assumed）"),
    ("nci_r", "少數股東損益 ÷ 稅後淨利", "%", None, "=SUM(Hist!C35:E35)/SUM(Hist!C34:E34)", None, "Derived", "＝2023–2025 年少數股東損益合計 ÷ 稅後淨利合計", "", ""),
    ("payout", "現金股利 ÷ 歸母淨利", "%", 0.5, "=(96.28+101.04+81.93)/SUM(Hist!C36:E36)", 1.0, "Derived", S_SA + "：2023–2025 年支付現金股利 96.28、101.04、81.93 合計 ÷ 歸母淨利合計", "", "每股股利資料商欄位 0.12 元與現金支出不符（見 README 疑點）；用於收支關卡與目標時點前股利"),
    ("da26", "FY2026 折舊攤銷 ÷ 營收", "%", None, "=Hist!E40/Hist!E4", None, "Derived", "＝2025 年折舊攤銷 ÷ 營收（" + S_SA + "）", "", "之後線性收斂到資本支出比例 × da_ss（FY2032）"),
    ("da_ss", "穩態折舊 ÷ 資本支出", "x", None, 0.9, None, "Assumed", "傳統製造業，成長有限，穩態資本支出略高於折舊", "", ""),
    ("cx26", "FY2026 資本支出", "百萬元", 160, 200, 250, "Assumed", S_SA + "：2025 年資本支出 202.77；近 12 個月 194.61", "", ""),
    ("cx_r", "資本支出 ÷ 營收（FY2027 起）", "%", 0.04, 0.055, 0.07, "Assumed", "2023–2025 年 5.6%、3.9%、6.1%（Hist，Derived）", "", ""),
    ("nwc_r", "營運資金增加 ÷ 營收增量", "%", 0.15, 0.25, 0.40, "Assumed", S_SA + "：2026-09-30 存貨 929.7＋應收 1,065 約占營收 60%，扣應付後以 25% 估", "", ""),
    ("sec", "6. 折現率（DCF 標準規格第 2 節）"),
    ("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", S_RF, "", "日期 2026-11-05，較股價基準日早 11 個交易日（逾期；跨模型一致，下次全面改版更新）"),
    ("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", ""),
    ("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", ""),
    ("b_own", "本公司 5 年 β（原始）", "x", None, 0.93, None, "Verified", S_SA + "：statistics 頁 5 年 β 0.93", "", "hiwin 模型 Comps 頁（資料日 2026-10-13）為 0.86"),
    ("b_peer", "同業 5 年 β 中位數（原始）", "x", None, "=MEDIAN(Comp!H5:H10)", None, "Derived", "＝Comp 頁核心同業 6 家 β 中位數", "", ""),
    ("beta", "β（兩來源 Blume 調整後中位數）", "x", "=beta-0.3", "=MEDIAN(0.67*b_own+0.33,0.67*b_peer+0.33)", "=beta+0.3", "Derived", "Blume 0.67β＋0.33；兩來源＝本公司、同業中位數", "", "敏感度 ∓0.3（跨模型慣例）"),
    ("kd", "稅前債務成本", "%", None, "=rf+0.02", None, "Assumed", "rf ＋ 2pt（財務費用 25 百萬元 ÷ 有息負債約 700 百萬元≈3.6%，與此相近）", "", ""),
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
SCN_DRIVERS = [("gcore", "傳統本業營收年增"), ("sh30", "2030 年人形滾柱絲杠份額"), ("gm30", "FY2030 毛利率")]


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
    HI["A1"] = "Hist：歷史財務（百萬元人民幣；中國會計準則）。新浪財經合併利潤表（萬元）換算為百萬元，取兩位小數。公式只引用本頁（定義名稱 h_*）。"
    hdr(HI, 3, ["項目", "", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"])
    HROWS = [  # (列, 鍵, 項目, [2023, 2024, 2025], 標記, 來源)
        (4, "h_rev", "營業收入", [3106.08, 3264.55, 3343.29], "Verified", S_SINA + "；" + S_SA + "（3,106、3,265、3,343）一致"),
        (5, "h_cogs", "營業成本", [2561.27, 2725.18, 2739.81], "Verified", S_SINA + "；2025 毛利率 18.05%、2024 16.52%、2023 17.54%（" + S_J + "）與此一致"),
        (6, "h_gp", "毛利（＝營收 − 營業成本）", ["=C4-C5", "=D4-D5", "=E4-E5"], "Derived", "＝營業收入 − 營業成本"),
        (7, "h_gm", "毛利率", ["=C6/C4", "=D6/D4", "=E6/E4"], "Derived", "＝毛利 ÷ 營收"),
        (8, "h_seg_c", "軸承、汽車零部件及其他（傳統本業，殘差）營收", ["=C4-C9", "=D4-D9", "=E4-E9"], "Derived", "＝營收 − 機器人零部件；公司分產品金額（軸承、汽車零部件、熱管理）未取得：山西證券只轉述 2025 年軸承 +1.95%、汽車零部件 +8.68%（" + S_SX + "）"),
        (9, "h_seg_r", "機器人零部件 營收", [0, 0, 16.0], "Verified", S_SX + "：2025 年機器人零部件收入 1,599.5 萬元，首次出現；2023、2024 以 0 計（Derived：「0 到 1」）"),
        (13, "h_sell", "銷售費用", [83.88, 108.75, 117.55], "Verified", S_SINA),
        (14, "h_admin", "管理費用", [148.15, 163.70, 192.54], "Verified", S_SINA),
        (15, "h_rnd", "研發費用", [101.33, 98.70, 118.94], "Verified", S_SINA),
        (16, "h_opex", "營業費用（銷售＋管理＋研發）", ["=SUM(C13:C15)", "=SUM(D13:D15)", "=SUM(E13:E15)"], "Derived", "＝三項合計"),
        (17, "h_fin", "財務費用（正＝淨費用）", [14.87, 33.04, 25.22], "Verified", S_SINA),
        (18, "h_inv", "投資收益＋公允價值變動", [0.75, -0.02, 0.68], "Verified", S_SINA + "（2023：0.85−0.10；2025：1.13−0.45）"),
        (19, "h_ts", "稅金及附加", [24.35, 26.57, 28.54], "Verified", S_SINA),
        (20, "h_opp", "營業利潤（中國會計準則）", [162.75, 110.73, 113.50], "Verified", S_SINA),
        (21, "h_ebit", "營業利益（核心：營業利潤 − 投資收益與公允價值 ＋ 財務費用）", ["=C20-C18+C17", "=D20-D18+D17", "=E20-E18+E17"], "Derived", "本模型營業利益不含金融資產收益與財務費用（財務費用列業外）"),
        (22, "h_oth_op", "其他營業損益（＝核心營業利益 − 毛利 ＋ 稅金及附加 ＋ 營業費用）", ["=C21-C6+C19+C16", "=D21-D6+D19+D16", "=E21-E6+E19+E16"], "Derived", "其他收益、減值、資產處置等合計（殘差）"),
        (25, "h_ebitm", "營業利益率（核心）", ["=C21/C4", "=D21/D4", "=E21/E4"], "Derived", ""),
        (32, "h_pbt", "利潤總額（稅前）", [163.02, 106.10, 109.17], "Verified", S_SINA),
        (26, "h_nonop", "業外（＝稅前 − 核心營業利益）", ["=C32-C21", "=D32-D21", "=E32-E21"], "Derived", "投資收益、公允價值、財務費用、營業外收支合計"),
        (33, "h_tax", "所得稅", [19.18, 15.12, 8.46], "Verified", S_SINA),
        (34, "h_nit", "稅後淨利", [143.84, 90.98, 100.71], "Verified", S_SINA),
        (35, "h_nci", "少數股東損益", [5.66, -0.39, 9.32], "Verified", S_SINA),
        (36, "h_ni", "歸屬母公司淨利", [138.18, 91.37, 91.39], "Verified", S_SINA + "；" + S_SA + "一致"),
        (37, "h_eps", "EPS（基本，元）", [0.40, 0.25, 0.25], "Verified", S_J),
        (38, "h_div", "現金股利支出（百萬元）", [96.28, 101.04, 81.93], "Verified", S_SA + "：現金流量表「支付普通股股利」"),
        (39, "h_sh", "股數（＝淨利 ÷ EPS，百萬股；EPS 取兩位小數，誤差大）", ["=C36/C37", "=D36/D37", "=E36/E37"], "Derived", ""),
        (40, "h_da", "折舊攤銷", [156.78, 167.95, 180.99], "Verified", S_SA),
        (41, "h_cx", "資本支出", [172.95, 128.14, 202.77], "Verified", S_SA),
        (42, "h_ocf", "營業現金流", [283.47, 195.83, 236.70], "Verified", S_SA),
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
    for c in "CDE": HI[f"{c}28"] = f"={c}8+{c}9-{c}4"
    HI.cell(29, 1, "差額：核心營業利益＋業外 − 稅前（應為 0）")
    for c in "CDE": HI[f"{c}29"] = f"={c}21+{c}26-{c}32"
    HI.column_dimensions["A"].width = 58; HI.column_dimensions["G"].width = 90

    # ================= Q（季度；一列一個數字，單一標記）=================
    Q["A1"] = "Q：2026 年季度資料（百萬元）。新浪財經合併利潤表年初至今累計值（萬元換算）；Q2、Q3 為相減得出。"
    hdr(Q, 3, ["項目", "2026Q1", "2026Q2", "2026H1", "2026Q3", "2026 年 1–9 月", "標記", "來源", "來源誘因", "說明"])
    QROWS = [
        (4, "營業收入", [721.79, "=D4-B4", 1533.84, "=F4-D4", 2249.91], "Verified", S_SINA, "", "上半年年減 19.03%（同花順 F10 營運分析頁）；第 3 季 716.07 與同花順報導 7.16 億元一致"),
        (5, "營業成本", [579.18, "=D5-B5", 1234.15, "=F5-D5", 1820.22], "Verified", S_SINA, "", ""),
        (6, "銷售＋管理＋研發費用", [102.36, "=D6-B6", 210.97, "=F6-D6", 329.80], "Verified", S_SINA, "", "前三季管理費用 158.39（2025 全年 192.54）"),
        (7, "毛利率", ["=1-B5/B4", "=1-C5/C4", "=1-D5/D4", "=1-E5/E4", "=1-F5/F4"], "Derived", "＝1 − 營業成本 ÷ 營收", "", "第 1 季 19.76%（山西證券轉載）"),
        (8, "歸屬母公司淨利", [30.80, "=D8-B8", 60.91, "=F8-D8", 982.87], "Verified", S_SINA + "；" + S_FP, "", "含洛軸股份公允價值利得，非本業"),
        (9, "公允價值變動收益", [0.45, "=D9-B9", 0.45, "=F9-D9", 1068.73], "Verified", S_SINA, "", "第 3 季 1,068.28 為洛軸股份公允價值變動（財報透視摘要）"),
        (10, "扣非歸母淨利", [None, None, None, None, 70.44], "Verified", S_FP, "", "前三季年減 22.90%；第 3 季年減 34.88%"),
    ]
    for (rr, lab, vals, tag, src, inc, note) in QROWS:
        Q.cell(rr, 1, lab)
        for j, v in enumerate(vals):
            if v is not None:
                c = Q.cell(rr, 2 + j, v)
                if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
                if rr == 7: c.number_format = "0.00%"
                else: c.number_format = "#,##0.00"
        Q.cell(rr, 7, tag); Q.cell(rr, 8, src); Q.cell(rr, 9, inc); Q.cell(rr, 10, note)
    Q.column_dimensions["A"].width = 30; Q.column_dimensions["H"].width = 90

    # ================= 線取數（線模型下游取數；藍字＝讀入值，M 欄寫來源檔與儲存格）=================
    LN["A1"] = "線取數：本模型讀入的線層級變數年度路徑（建置時由線模型成品依 ID 讀值）。情境對應：A＝線模型高情境、B＝基準、C＝低情境。"
    hdr(LN, 3, ["項目", "情境", ""] + [f"FY{y}" for y in range(2025, 2033)] + ["單位", "來源（檔案與儲存格）"])
    LR = {}
    r = 4
    for s in SC:
        LN.cell(r, 1, "全球人形機器人用行星滾柱絲杠需求（線 4）"); LN.cell(r, 2, f"{s}（{LINE_SC[s]}）")
        srcs = []
        for y in range(2026, 2033):
            v, src = LD["scr"][(y, s)]
            c = LN.cell(r, 4 + (y - 2025), v); c.font = BLUE; c.number_format = "#,##0"
            srcs.append(src)
        LN.cell(r, 12, "根"); LN.cell(r, 13, f"{srcs[0]} … {srcs[-1]}")
        LR[s] = r; r += 1
    LN.column_dimensions["A"].width = 52; LN.column_dimensions["M"].width = 110

    # ================= Thesis（命題模組）=================
    TH["A1"] = ("命題模組：機器人零部件營收＝全球人形滾柱絲杠需求（線 4 × 倍數）× 本公司份額 × 單價（2025 年均價 ×（1 − 線 6 年降幅））；"
                "傳統本業（軸承、汽車零部件、熱管理）＝FY2026 實績錨 × 成長。FY2026＝1–9 月實績＋第 4 季假設。百萬元。")
    TH["A2"] = "欄：D＝FY2025 … K＝FY2032。2026 年機器人零部件營收為 Assumed（rob26），隱含份額以線 4 基準路徑校準（三情境同一實績年）。"
    hdr(TH, 3, ["項目", "說明", ""] + [f"FY{y}" for y in range(2025, 2033)])
    THR = {}
    r = 4
    for s in SC:
        TH.cell(r, 1, f"情境 {s}").font = BOLD; TH.cell(r, 1).fill = SECF; r += 1
        R = {}
        R["rev26"] = r; TH.cell(r, 1, "FY2026 合併營收（錨）"); TH.cell(r, 2, "＝1–9 月＋2025Q4 ×（1＋q4g）")
        TH[f"{TC[2026]}{r}"] = "=r9m26+(Hist!E4-r9m26/(1+g9m26))*(1+q4g)"; r += 1
        R["price"] = r; TH.cell(r, 1, "滾柱絲杠均價（元／根）"); TH.cell(r, 2, "＝2025 均價 ×（1 − 線 6 年降幅）^(年 − 2025)")
        for y in range(2025, 2033):
            TH[f"{TC[y]}{r}"] = "=asp25" if y == 2025 else f"={TC[y - 1]}{r}*(1-pdr)"
        r += 1
        R["dem"] = r; TH.cell(r, 1, "全球人形滾柱絲杠需求（線 4 × 倍數，千根）")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"=線取數!{TC[y]}{LR[s]}*hu_k/1000"
        r += 1
        R["mkt"] = r; TH.cell(r, 1, "全球人形滾柱絲杠市場（百萬元，本公司單價口徑）"); TH.cell(r, 2, "＝需求（千根）× 均價 ÷ 1000")
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"] = f"={TC[y]}{R['dem']}*{TC[y]}{R['price']}/1000"
        r += 1
        R["sh"] = r; TH.cell(r, 1, "本公司份額"); TH.cell(r, 2, "FY2026＝隱含（rob26 ÷ 線 4 基準市場）；線性走到 sh30，其後持平")
        TH[f"{TC[2026]}{r}"] = f"=rob26/(線取數!{TC[2026]}{LR['B']}*hu_k/1000*{TC[2026]}{R['price']}/1000)"
        for y in range(2027, 2033):
            TH[f"{TC[y]}{r}"] = f"=${TC[2026]}${r}+(sh30_{s}-${TC[2026]}${r})*MIN(1,({y}-2026)/4)"
        for y in range(2026, 2033): TH[f"{TC[y]}{r}"].number_format = "0.0%"
        r += 1
        R["rob"] = r; TH.cell(r, 1, "機器人零部件 營收"); TH.cell(r, 2, "FY2026＝rob26；FY2027 起＝市場 × 份額")
        TH[f"{TC[2025]}{r}"] = "=Hist!E9"
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = "=rob26" if y == 2026 else f"={c}{R['mkt']}*{c}{R['sh']}"
        r += 1
        R["core"] = r; TH.cell(r, 1, "傳統本業 營收（軸承、汽車零部件、熱管理、其他）"); TH.cell(r, 2, "FY2026＝合併營收 − rob26；FY2027 起＝前一年 ×（1＋gcore）")
        TH[f"{TC[2025]}{r}"] = "=Hist!E8"
        for y in range(2026, 2033):
            c = TC[y]
            TH[f"{c}{r}"] = f"={c}{R['rev26']}-rob26" if y == 2026 else f"={TC[y - 1]}{r}*(1+gcore_{s})"
        r += 1
        R["robsh"] = r; TH.cell(r, 1, "機器人零部件占合併營收")
        for y in range(2026, 2033):
            c = TC[y]; TH[f"{c}{r}"] = f"={c}{R['rob']}/({c}{R['rob']}+{c}{R['core']})"; TH[f"{c}{r}"].number_format = "0.0%"
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
    TH.column_dimensions["A"].width = 44; TH.column_dimensions["B"].width = 50

    # ================= PL（標準損益表，三情境）=================
    PL["A1"] = ("PL：分部（傳統本業、機器人零部件）→ 合併營收 → 標準損益表 → FCFF。"
                "C＝FY2023 … E＝FY2025 為實績；F＝FY2026（1–9 月實績＋第 4 季假設）… L＝FY2032。百萬元；中國會計準則（法定財報口徑，GAAP）。預估不含洛軸股份公允價值變動（不可預測，其價值計入淨現金）。")
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
        row("core", "傳統本業 營收（軸承、汽車零部件、熱管理）", "百萬元", H(8), lambda y: f"=Thesis!{TC[y]}{th['core']}")
        row("rob", "機器人零部件 營收", "百萬元", H(9), lambda y: f"=Thesis!{TC[y]}{th['rob']}")
        row("rev", "營收（合併）", "百萬元", H(4), lambda y: f"={C(y)}{R['core']}+{C(y)}{R['rob']}")
        row("chk", "檢查：分部合計 − 合併營收（應為 0）", "百萬元",
            lambda y: f"={C(y)}{R['core']}+{C(y)}{R['rob']}-{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['core']}+{C(y)}{R['rob']}-{C(y)}{R['rev']}")
        PL.cell(R["chk"], 1, "檢查：分部合計 − 合併營收（應為 0）" if s == "B" else f"檢查（情境 {s}）：分部合計 − 合併營收（應為 0）")
        row("yoy", "營收年增率", "%", lambda y: None if y == 2023 else f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1",
            lambda y: f"={C(y)}{R['rev']}/{COL[y - 1]}{R['rev']}-1", "0.0%")
        row("cogs", "營業成本", "百萬元", H(5), lambda y: f"={C(y)}{R['rev']}*(1-Thesis!{TC[y]}{th['gm']})")
        row("gp", "毛利", "百萬元", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}", lambda y: f"={C(y)}{R['rev']}-{C(y)}{R['cogs']}")
        row("gm", "毛利率", "%", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['gp']}/{C(y)}{R['rev']}", "0.0%")
        row("ts", "稅金及附加", "百萬元", H(19), lambda y: f"={C(y)}{R['rev']}*ts_r")
        row("opex", "營業費用（銷售＋管理＋研發）", "百萬元", H(16), lambda y: f"={C(y)}{R['rev']}*Thesis!{TC[y]}{th['opex']}")
        row("othop", "其他營業損益（其他收益 − 減值等）", "百萬元", H(22), lambda y: f"={C(y)}{R['rev']}*oth_r")
        row("op", "營業利益（核心，不含金融資產收益與財務費用）", "百萬元", H(21), lambda y: f"={C(y)}{R['gp']}-{C(y)}{R['ts']}-{C(y)}{R['opex']}+{C(y)}{R['othop']}")
        row("opm", "營業利益率", "%", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", lambda y: f"={C(y)}{R['op']}/{C(y)}{R['rev']}", "0.0%")
        row("nonop", "其他收入淨額（業外：投資收益、公允價值、財務費用、營業外）", "百萬元", H(26), lambda y: f"=-{C(y)}{R['rev']}*fin_r")
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
        R["hcash"] = r; DC.cell(r, 1, "＋增資或發行（本公司無此事件，0）"); DC[f"B{r}"] = "=0"; r += 1
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
    CP["A1"] = ("Comp：同業前瞻 P/E（stockanalysis.com statistics，資料商）。股價＝" + S_CLOSE + "（Verified）；前瞻 P/E＝原值 × 新收盤 ÷ 原報價（Derived，EPS 不變）；"
                "只有原報價與報價日可查的同業進中位數（6 家，少於 8 家不做迴歸）；原報價日未取得者另列、不進中位數。")
    hdr(CP, 4, ["公司", "代號", "股價（2026-11-20）", "報價日", "前瞻 P/E 分母", "前瞻 P/E", "近 12 月 P/E（原報價日）", "5 年 β", "會計口徑", "進中位數", "說明", "原報價", "原報價日", "原前瞻 P/E", "標記與換算"])
    PEERS = [  # (公司, 代號, 11-20 收盤, 原報價, 原報價日, 原前瞻 P/E, 近 12 月 P/E, β, 口徑, 進中位數, 說明, 分母, 來源)
        ("THK", "6481.T", 4249, 4249, "2026-11-20", 10.68, None, 0.76, "日本基準（未查核）", "是", "線性傳動、滾珠絲杠、軸承相關", "資料商 NTM 共識 EPS", S_HX),
        ("NSK", "6471.T", 892.5, 926.2, "2026-11-24", 11.70, None, 0.55, "日本基準（未查核）", "是", "軸承、滾珠絲杠、汽車零件", "資料商 NTM 共識 EPS", S_SAP),
        ("Schaeffler", "SHA0.DE", 4.83, 5.07, "2026-11-23", 9.66, None, 1.66, "歐盟 IFRS（未查核）", "是", "軸承、汽車零件、行星滾柱絲杠（歐元計價，P/E 無幣別問題）", "資料商 NTM 共識 EPS", S_SAP),
        ("雙環傳動", "002472.SZ", 35.45, 41.21, "2026-10-30", 22.32, None, 0.52, "中國會計準則", "是", "齒輪、精密減速器；汽車零件", "資料商 NTM 共識 EPS", S_LD),
        ("拓普集團", "601689.SS", 48.11, 46.52, "2026-09-30", 26.89, None, 0.80, "中國會計準則", "是", "汽車零件（含熱管理）、人形致動器", "資料商 NTM 共識 EPS", S_LD),
        ("北特科技", "603009.SS", 41.10, 47.80, "2026-11-24", 92.27, None, 0.71, "中國會計準則", "是", "汽車零件、行星滾柱絲杠（本專案列為同業）；前瞻 P/E 偏高但中位數不受單點影響", "資料商 NTM 共識 EPS", S_SAP),
        ("貝斯特", "300580.SZ", 14.00, None, "報價日未取得", None, None, 0.60, "中國會計準則", "否", "滾珠、滾柱絲杠；前瞻 P/E 26.78 來自 hiwin 模型 Comps 頁（資料日 2026-09-29），該頁未記原報價，無法換算至 2026-11-20，故不納入", "—", S_HX),
        ("中大力德", "002896.SZ", 48.36, None, "報價日未取得", None, None, 0.25, "中國會計準則", "否", "減速器；資料商無前瞻 P/E", "—", S_SAP),
    ]
    for i, (nm, tk, px1, px0, d0, pe0, ttm, bt, acc, med, note, den, srcx) in enumerate(PEERS):
        rr = 5 + i
        CP.cell(rr, 1, nm); CP.cell(rr, 2, tk)
        CP.cell(rr, 3, px1).font = BLUE; CP.cell(rr, 4, "2026-11-20"); CP.cell(rr, 5, den)
        if pe0 is not None: CP.cell(rr, 6, f"=N{rr}*C{rr}/L{rr}")
        else: CP.cell(rr, 6, "報價日未取得" if tk == "300580.SZ" else "未取得")
        CP.cell(rr, 8, bt).font = BLUE; CP.cell(rr, 9, acc); CP.cell(rr, 10, med); CP.cell(rr, 11, note)
        if px0 is not None: CP.cell(rr, 12, px0).font = BLUE
        CP.cell(rr, 13, d0)
        if pe0 is not None: CP.cell(rr, 14, pe0).font = BLUE
        CP.cell(rr, 15, "股價 Verified（" + S_CLOSE + "）；前瞻 P/E Derived：＝原值 × 2026-11-20 收盤 ÷ 原報價（分母 EPS 不變）；原值：" + srcx)
    CP.cell(13, 1, "五洲新春（本公司，對照）"); CP.cell(13, 2, "603667.SS"); CP.cell(13, 3, "=px"); CP.cell(13, 4, "2026-11-20")
    CP.cell(13, 6, "資料商無前瞻 P/E"); CP.cell(13, 7, "不適用（近 12 月淨利含公允價值利得）"); CP.cell(13, 8, "=b_own"); CP.cell(13, 10, "否"); CP.cell(13, 11, "本公司；只作對照")
    CP.cell(15, 1, "核心同業前瞻 P/E 中位數"); CP["F15"] = "=MEDIAN(F5:F10)"
    CP.cell(16, 1, "本公司現價 ÷ 模型 FY2027 EPS（情境 B；分母：本模型）"); CP["F16"] = f"=px/PL!G{PR['B']['eps']}"
    CP.cell(17, 1, "本公司現價 ÷ FY2025 EPS 0.25 元（實績）"); CP["F17"] = "=px/Hist!E37"
    CP.cell(18, 1, "本公司現價 ÷ 模型 FY2030 EPS（情境 B）"); CP["F18"] = f"=px/PL!J{PR['B']['eps']}"
    CP.cell(19, 1, "本公司現行倍數 ÷ 同業中位數（模型 FY2027 EPS 口徑）"); CP["F19"] = "=F16/F15"
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
                             ("ronic", "RONIC", "=wacc+ronic_sp", "0.00%"),
                             ("na_ps", "非營運淨資產每股（現金＋長期投資 × inv_k − 有息負債 − 少數股東權益，元）", "=(cash_v+ltinv_v*inv_k-debt_v-nci_v)/sh_v", "#,##0.00")):
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
    row3("cmpu", "Comparable 未調整（同業 P/E × FY2027 EPS ＋ 非營運淨資產每股）", [f"=MAX(0,{c}{r - 1}*pe_peer)+na_ps" for c in "BCD"], "#,##0.0", "同業 P/E＝In 頁 pe_peer（核心同業前瞻 P/E 中位數）")
    row3("cmp", "Comparable（本業部分 × 相對倍數係數）", [f"=({c}{VR['cmpu']}-na_ps)*rel_k+na_ps" for c in "BCD"], "#,##0.0", "rel_k＝1.00（未調整：缺十年相對倍數歷史）")
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
    row1("relreq", "支持現價所需的相對倍數係數（DCF 不變）", f"=((px-w_dcf*B{VR['dcf_tp']})/(1-w_dcf)-na_ps)/(B{VR['cmpu_tp']}-na_ps)", "0.00",
         "反解：目標價＝現價；DCF 以 RONIC＝WACC（無超額報酬）而高相對倍數隱含持續超額報酬，兩法對競爭優勢的假設相反")
    row1("relflip", "評等升一級所需的相對倍數係數", f'=IF(ISNUMBER(B{VR["up_b"]}),((B{VR["up_b"]}-w_dcf*B{VR["dcf_tp"]})/(1-w_dcf)-na_ps)/(B{VR["cmpu_tp"]}-na_ps),"—")', "0.00")
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
    RC["A1"] = "Recon：模型對實績、線 6 市場規模的同口徑對照；反解支持現價所需的主導變數。"
    hdr(RC, 3, ["項目", "模型", "對照", "差距（模型 − 對照）", "說明"])
    rr = 4
    def rrow(lab, a, b, note, fmt="#,##0.0", diff=True):
        nonlocal rr
        RC.cell(rr, 1, lab); RC.cell(rr, 2, a).number_format = fmt
        if b is not None: RC.cell(rr, 3, b).number_format = fmt
        if diff and b is not None: RC.cell(rr, 4, f"=B{rr}-C{rr}").number_format = fmt
        RC.cell(rr, 5, short(note)); rr += 1
    rrow("FY2026 營收（三情境同）vs 山西證券 2026E 營收 3,865（舊稿，不採用）", f"=PL!F{PR['B']['rev']}", 3865, "山西證券預測早於三季報，前三季營收年減 15% 使其不可達；只作歷史對照")
    rrow("FY2026 歸母淨利（情境 B，不含公允價值利得）vs 前三季扣非歸母年化（70.44 × 4／3）", f"=PL!F{PR['B']['ni']}", "=70.44*4/3", "前三季扣非年減 22.9%（" + S_FP + "）；模型以全年費用率與第 4 季假設推得")
    rrow("FY2027 EPS（情境 B）vs 山西證券 2027E EPS 0.55 元（舊稿，不採用）", f"=PL!G{PR['B']['eps']}", 0.55, "同上，只作歷史對照", "0.00")
    rrow("FY2026 毛利率 vs 2026 年 1–9 月實績", f"=PL!F{PR['B']['gm']}", "=gm26", "模型以 1–9 月為錨（差距應為 0）", "0.0%")
    rrow("FY2026 第 4 季隱含營收 vs 2026Q3 實績", f"=PL!F{PR['B']['rev']}-r9m26", "=Q!E4", "第 4 季若維持第 3 季水準即為對照值")
    rrow("洛軸股份持股市值（50.98 百萬股 × 2026-11-06 收盤 31.78）vs 2026-09-30 長期投資帳面", "=lz_sh*lz_px", "=ltinv_v", S_LZ + "；長期投資含其他項目，三季末公允價值日價格未取得；差距約 10%，已由 inv_k 折價涵蓋")
    rrow("洛軸股份持股市值 ÷ 本公司市值（現價 × 股數）", "=lz_sh*lz_px/(px*sh_v)", None, "持股約占本公司市值一成；洛軸股份不在收盤檔，價格為舊日期的資料商值", "0.0%", False)
    rrow("非營運淨資產每股（現金＋長期投資 × inv_k − 有息負債 − 少數股東權益）÷ 現價", f"=Val!B{VR['na_ps']}/px", None, "市值中由淨資產解釋的比例；其餘需由本業與機器人選擇權解釋", "0.0%", False)
    rr += 1
    RC.cell(rr, 1, "線 4／線 6 市場對照（滾柱絲杠；百萬美元）").font = BOLD; rr += 1
    rrow("FY2030 機器人零部件營收（情境 B，百萬美元）÷ 線 6 滾柱絲杠市場 2030（中國量產價口徑）", f"=PL!J{PR['B']['rob']}/fx/{LD['mkt30_cn'][0]}", None, f"線 6 2030 基準需求 × 中國量產價 {LD['mkt30_cn'][0]:,.0f} 百萬美元（{LD['mkt30_cn'][1]}）；本公司營收含非人形用途，故為上界", "0.0%", False)
    rrow("FY2030 機器人零部件營收（情境 B，百萬美元）÷ 線 6 滾柱絲杠市場 2030（覓途口徑）", f"=PL!J{PR['B']['rob']}/fx/{LD['mkt30_6'][0]}", None, f"線 6 基準 {LD['mkt30_6'][0]:,.0f} 百萬美元（{LD['mkt30_6'][1]}）", "0.0%", False)
    rrow("FY2030 全球人形滾柱絲杠需求：線 4 基準（千根）vs 線 6（千根）", f"=Thesis!J{THR['B']['dem']}", LD["dem30_6"][0] / 1000, f"線 6 {LD['dem30_6'][1]}；線 4 {LD['scr'][(2030, 'B')][1]}；兩線不一致，本模型採線 4（人形用量讀人形線）", "#,##0")
    rrow("FY2026 隱含全球人形絲杠份額（rob26 ÷ 線 4 基準市場）", f"=Thesis!E{THR['B']['sh']}", None, "＝機器人零部件營收 ÷（線 4 需求 × 單價）；機器人零部件含非人形用途，故為上界", "0.0%", False)
    rrow("2025 本公司均價（美元）vs 線 6 滾柱絲杠 2026 單價低案（中國量產價）", "=asp25/fx", LD["p26"][0][0], f"{LD['p26'][0][1]}；基準 {LD['p26'][1][0]:,.0f} 美元、高 {LD['p26'][2][0]:,.0f} 美元（價格口徑互相矛盾，線 6 已列敏感度）", "#,##0")
    rrow("2025 絲杠出貨 8.7 萬根 × 均價 1,360 元（百萬元）vs 2025 機器人零部件營收", "=87000*asp25/1000000", "=Hist!E9", "出貨量與營收差約 7 倍：出貨量為產業網站轉述，未經公司證實，或營收口徑不含全部絲杠（見 README 疑點）；本模型營收以公司口徑營收為準，不採出貨量")
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
    s26 = LD["scr"][(2026, "B")]; s30 = LD["scr"][(2030, "B")]; s32 = LD["scr"][(2032, "B")]
    ZR = [
        ("全球人形滾柱絲杠需求 FY2026–FY2032（A／B／C＝線高／基準／低）", "4 人形", f"基準 {s26[0]:,.0f} → {s32[0]:,.0f}（× hu_k）", "根", "全球、人形用（全尺寸用量 × 採用比例）", f"{s26[1]}；{s32[1]}（逐年見線取數頁）", "Assumed", "Thesis 全球市場與機器人零部件營收", "一致（直接讀值，hu_k＝1）"),
        ("滾柱絲杠單價年降幅", "6 共用零組件", "=pdr", "%／年", "線 6 行星滾柱絲杠，不分廠", f"{LD['pd'][1][1]}（低 {LD['pd'][0][1]}、高 {LD['pd'][2][1]}）", "Analogy", "Thesis 單價路徑", "與線 6 一致；本公司 2025 年均價年減 39%（產業網站，Interested-party）高於線 6 高案 20%，列為判斷題"),
        ("滾柱絲杠單價水準", "6 共用零組件", "=asp25", "元／根", "本公司 2025 年均價（中國量產價口徑）", f"線 6 輸入!F90、F91（五洲新春 2025 均價 1,360 元）；線 6 2026 單價低案 {LD['p26'][0][0]:,.0f} 美元、基準 {LD['p26'][1][0]:,.0f} 美元", "Interested-party", "Thesis 單價起點", "與線 6「中國量產價」口徑一致；不採線 6 基準單價（覓途口徑，約為 3.7 倍）：本公司實際成交價屬中國量產價"),
        ("FY2030 全球人形滾柱絲杠需求量", "4 人形／6 共用零組件", f"線 4 基準 {s30[0]:,.0f} 根", "根", "全球人形", f"{s30[1]}；線 6 {LD['dem30_6'][1]}（{LD['dem30_6'][0]:,.0f} 根）", "Assumed", "Thesis 需求", "不一致：線 6 為線 4 的約 1.9 倍；依工作單人形用量讀人形線（線 4），Recon 並列"),
        ("滾柱絲杠市場規模 2030（基準需求）", "4 人形／6 共用零組件", "見 Recon", "百萬美元", "線 4 單價隨學習曲線；線 6 覓途／中國量產價", f"線 4 {LD['mkt30_4'][1]}；線 6 {LD['mkt30_6'][1]}、{LD['mkt30_cn'][1]}", "Derived", "Recon 隱含份額", "只作對照，不進公式"),
        ("本公司滾柱絲杠 2025 年出貨量", "6 共用零組件", "87,000 根", "根", "產業網站轉述", f"線 6 追蹤指標!D11；份額!F9（份額未取得）", "Interested-party", "Recon（不進公式）", "與公司口徑營收 16.0 百萬元差約 7 倍，見 Recon 與 README 疑點；本模型不採"),
        ("全球人形出貨（只作對照）", "4 人形", f"基準 {LD['scr'][(2026, 'B')][0] / 4.2:,.0f} 台（2026，絲杠需求 ÷ 4.2）", "台", "全球、含小型機", f"{L4N} 下游取數（SHIP_2026、Q_SCREW_AVG）", "Assumed", "未使用", "不使用；只用絲杠需求"),
    ]
    for i, z in enumerate(ZR):
        for j, v in enumerate(z): DZ.cell(4 + i, 1 + j, v)
    DZ.column_dimensions["A"].width = 50; DZ.column_dimensions["F"].width = 80; DZ.column_dimensions["I"].width = 70

    # ================= Risk =================
    RK["A1"] = "Risk：風險登錄（原則第 10 節第 10、19 條）。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    RISKS = [
        ("洛軸股份股價波動與鎖定期（持股約 1,477 百萬元，占市值約一成）", "部分（inv_k 0.6–1.0）", "公允價值利得會隨股價反覆，且不屬本業", "洛軸股份收盤價、鎖定期公告、減持公告"),
        ("熱管理業務（虹新製冷）大客戶需求下滑", "是（gcore_C、q4g 低案）", "前三季營收年減 15%，第 3 季扣非年減 35%", "季報營收與毛利率、大客戶訂單"),
        ("匯兌損失（美元貶值）", "部分（fin_r）", "第 3 季扣非受匯兌拖累", "財務費用季變化"),
        ("滾柱絲杠價格戰（新劍傳動、雙林）", "是（pdr、sh30_C）", "雙林報價 499 元、新劍據報建 100 萬根年產線", "本公司機器人零部件營收、毛利率"),
        ("人形量產延後或不及線 4 基準", "是（情境 C＝線 4 低）", "", "IDC、TrendForce 季出貨；中國人形產量"),
        ("機器人零部件營收未單獨揭露", "是（rob26 區間 25–90）", "2026 年數字為 Assumed", "半年報、年報分產品"),
        ("客戶集中（最新可得為 2016 上半年：前五大 64.57%、Schaeffler 35.03%）", "否", "資料過舊，現況未取得", "年報前五大客戶"),
        ("傳統本業毛利率與費用（管理費用前三季 158 百萬元，已接近 2025 全年 193）", "是（gm30、opex30）", "", "季報費用率"),
    ]
    for i, k in enumerate(RISKS):
        for j, v in enumerate(k): RK.cell(4 + i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["D"].width = 60

    # ================= Sources =================
    SO["A1"] = "Sources：本模型引用的原始來源與資料商（含網址；HTML 展示層不放網址）。"
    for i, s_ in enumerate([S_PX, S_SINA, S_SA, S_J, S_N, S_FP, S_SX, S_LZ, S_HW, S_RF, S_CRP, S_SAP, S_HX, S_LD,
                            f"線 4：{L4N}", f"線 6：{L6N}",
                            "physical_ai/map/profiles/603667.md（背景，數字已另行查證）"], 3):
        SO.cell(i, 1, s_)
    SO.column_dimensions["A"].width = 160

    # ================= README（xlsx 內）=================
    S = T
    v = lambda k: f"Val!B{VR[k]}"
    lines = [
        ('="五洲新春 Wuzhou Xinchun（603667.SS）個股模型 ' + VER + '｜build_id "&Sens!D2&"｜資料截止日 2026-11-24｜成品日 2026-11-24｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'),
        "命題：五洲新春的營收九成九以上是軸承、汽車零部件與熱管理，2026 年前三季營收年減 15%；前三季歸母淨利 9.8 億元幾乎全是洛軸股份持股的公允價值利得，本業扣非只有 0.7 億元。股價中的機器人滾柱絲杠成分，由線 4 全球人形絲杠需求 × 本公司份額 × 單價（線 6 年降）推得。",
        (f'="結論（條件式）：在 FY2026 以 1–9 月實績為錨、傳統本業 FY2027 起年增 "&{S("gcore_B","0%")}&"、線 4 基準人形絲杠需求（× "&{S("hu_k","0.0")}&"）、情境 B 2030 年份額 "&{S("sh30_B","0%")}&"、單價年降 "&{S("pdr","0%")}&"、同業 P/E 未加溢價（相對倍數係數 "&{S("rel_k","0.00")}&"）的假設下，加權目標價約 "&{S(v("tp"),"#,##0.0")}&" 元，相對現價 "&{S("px","#,##0.00")}&" 元 "&{S(v("up"),"+0.0%;-0.0%")}&"，估值評等 "&{v("rating")}&"，投資建議 "&{v("rec")}&"。"'),
        (f'="DCF 目標價 "&{S(v("dcf_tp"),"#,##0.0")}&"、Comparable 目標價 "&{S(v("cmp_tp"),"#,##0.0")}&"（各半；Comparable 已加回非營運淨資產每股 "&{S(v("na_ps"),"#,##0.0")}&" 元）；WACC "&{S("wacc","0.00%")}&"、ke "&{S("ke","0.00%")}&"、g "&{S("g_tv","0.00%")}&"；情境 B 終值占 EV "&{S(f"Val!C{VR['tvs']}","0%")}&"。"'),
        (f'="上檔空間歸因：DCF 貢獻 "&{S(v("att_d"),"+0.0%;-0.0%")}&"、未調整 Comparable 貢獻 "&{S(v("att_c"),"+0.0%;-0.0%")}&"、相對倍數調整貢獻 "&{S(v("att_r"),"+0.0%;-0.0%")}&"。本公司現價對模型 FY2027 EPS 的倍數是同業中位數的 "&{S(v("relnow"),"0.0")}&" 倍；支持現價需要相對倍數係數 "&{S(v("relreq"),"0.0")}&"，升一級需要 "&{S(v("relflip"),"0.0")}&"。"'),
        (f'="收支關卡：情境 B FY2026–FY2030 FCFF 扣股利合計 "&{S(v("gate_b"),"#,##0.0")}&" 百萬元，"&{v("gate_t")}&"；距最近級距邊界 "&{S(v("dist"),"0.0%")}&"。"'),
        "主要失效條件（可觀察）：(1) 機器人零部件營收單季超過 1 億元且毛利率高於本業；(2) 熱管理與軸承營收回到年增 5% 以上；(3) 洛軸股份股價續漲而公司兌現持股；(4) 本公司份額高於情境 A 的 20%。",
        "方法：DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，WACC 折現，明確期到 FY2032，RONIC＝WACC；現金與長期投資（含洛軸股份，折價 inv_k）計入淨現金）50%＋同業前瞻 P/E × 模型 FY2027 EPS 加非營運淨資產每股 50%；三情境 25／50／25。",
        "限制：分產品金額（軸承、汽車零部件、熱管理）未取得；2026 年機器人零部件營收為 Assumed；同業 P/E 由原報價日按 2026-11-20 收盤等比例換算（EPS 未更新）；不讀 Siliconomics／Tokenomics／Wattonomics。詳見 repo README.md。",
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
        ("情境 B 的 2030 年人形絲杠份額改為情境 A 值", "sh30_B", ivE["sh30_A"]),
        ("情境 B 的 2030 年人形絲杠份額改為情境 C 值", "sh30_B", ivE["sh30_C"]),
        ("情境 B 傳統本業年增改為情境 A 值", "gcore_B", ivE["gcore_A"]),
        ("情境 B 傳統本業年增改為情境 C 值", "gcore_B", ivE["gcore_C"]),
        ("情境 B 的 FY2030 毛利率改為情境 A 值", "gm30_B", ivE["gm30_A"]),
        ("情境 B 的 FY2030 毛利率改為情境 C 值", "gm30_B", ivE["gm30_C"]),
        ("人形絲杠需求為線 4 的 5 倍（全部情境）", "hu_k", 5.0),
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
    s1 = solve("hu_k", px, 1.0, 400.0)
    s2 = solve("hu_k", flip_tp, 1.0, 400.0)
    s3 = solve("sh30_B", px, max(ivE["sh30_B"], 0.01), 1.0)
    s4 = solve("hu_k", px, 1.0, 400.0, over={"sh30_B": 0.5, "sh30_A": 0.5, "sh30_C": 0.5})
    nf = lambda x: x if x else -1
    d30 = LD["scr"][(2030, "B")][0]
    solve_rows = [
        ("支持現價（目標價＝現價）所需的人形絲杠需求倍數（相對線 4，全部情境）", nf(s1), 1.0, "其他輸入不變；−1＝400 倍內無解" if not s1 else f"其他輸入不變；情境 B 的 CY2030 全球人形絲杠需求相當於 {d30 * s1 / 1e6:,.1f} 百萬根（線 4 基準 {d30 / 1e6:.1f}）"),
        (f"評等升一級（上檔 {flip:+.0%}）所需的人形絲杠需求倍數", nf(s2), 1.0, "其他輸入不變；−1＝400 倍內無解" if not s2 else f"情境 B 的 CY2030 全球人形絲杠需求相當於 {d30 * s2 / 1e6:,.2f} 百萬根"),
        ("支持現價所需的情境 B 2030 年人形絲杠份額", nf(s3), ivE["sh30_B"], "其他輸入不變；−1＝份額 100% 仍無解"),
        ("三情境份額皆 50% 時，支持現價所需的人形絲杠需求倍數", nf(s4), 1.0, "份額與量同時放大的組合；−1＝400 倍內無解"),
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
    segs = [("傳統本業（軸承、汽車零部件、熱管理及其他，殘差）", "傳統本業 營收（軸承、汽車零部件、熱管理）", Hs["E8"].value), ("機器人零部件", "機器人零部件 營收", Hs["E9"].value)]
    json.dump(dict(company="wuzhou", ticker="603667.SS", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=round(v, 2), unit="百萬元", source=S_SX + "；" + S_SINA) for n, l, v in segs],
                   sublines=[],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）"],
                   note="公司分產品金額（軸承、汽車零部件、熱管理）在可取得的資料中未找到，只取得 2025 年增速（軸承 +1.95%、汽車零部件 +8.68%）與機器人零部件 2025 年收入 16.0 百萬元；故傳統本業以殘差列示，機器人零部件由命題（線 4 需求 × 份額 × 單價）推導。"),
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
    t = f"""# physical_ai/wuzhou/ — 五洲新春 Wuzhou Xinchun（603667.SS）個股模型 {VER}

<!-- 本檔由 build_wuzhou.py 依成品重算值產生；數字不手寫（原則第 10 節第 51 條）。 -->

**本模型適用《高不確定性預估工作原則》**（`docs/plan/高不確定性預估工作原則.md`）：機器人零部件營收公司未單獨揭露、人形需求尚未發生，相關輸入一律給區間並標 Assumed／Analogy；結論只在證據邊界內成立。

- 成品：`dist/{NAME}`（build_id `{sens['build_id']}`；LibreOffice 重算，0 錯誤值）；HTML 展示層 `dist/{NAME[:-5]}.html`（`gen_html_json.py` 產生 `html.json`，再以 `tools/company_html/build_company_html.py` 產生）
- 建置：`python3 physical_ai/wuzhou/build_wuzhou.py`（需要 soffice；整本重算敏感度與反解；結尾自動跑 in_layout、sens_coverage --require-keycol、model_governance、rating_gate --strict、zero_rows、cached_values、segment_check）
- 日期：資料截止日 2026-11-24（財報、同業倍數）｜成品日 2026-11-24｜股價基準日 2026-11-20（收盤 {I['px']:,.2f} 元）｜評價日 2026-12-31｜目標時點 2027-11-20
- 會計年度＝日曆年；單位百萬元人民幣；中國會計準則（法定財報口徑，GAAP）。上海證券交易所主板。

## 結論（條件式）

在「FY2026 以 1–9 月實績為錨、傳統本業 FY2027 起年增 {I['gcore_B']:.0%}、線 4 基準人形滾柱絲杠需求（× {I['hu_k']:.0f}）、情境 B 的 2030 年份額 {I['sh30_B']:.0%}、單價年降 {I['pdr']:.0%}、同業 P/E 不加溢價（相對倍數係數 {I['rel_k']:.2f}）」的假設下：

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
| 非營運淨資產每股（現金＋長期投資 × {I['inv_k']:.0%} − 負債；含洛軸股份持股） | {g('na_ps'):,.1f} 元 |

上檔空間歸因：DCF 貢獻 {g('att_d'):+.1%}、未調整 Comparable 貢獻 {g('att_c'):+.1%}、相對倍數調整 {g('att_r'):+.1%}。現價對模型 FY2027 EPS 的倍數是核心同業中位數（{I['pe_peer']:.1f} 倍）的 {g('relnow'):.1f} 倍；要讓目標價＝現價，Comparable 的相對倍數係數須達 {g('relreq'):.1f}。

**讀法：** 2026 年前三季歸母淨利 9.83 億元，其中約 9.1 億元是持有洛軸股份（7.04%）的公允價值與處置利得；扣非只有 0.70 億元、年減 22.9%，營收年減 15%。本模型預估不含這類利得，持股價值以折價後計入淨現金與 Comparable 的非營運淨資產。情境 B 機器人零部件營收 {ser('rob', '{:,.0f}')}（FY2026→FY2032，百萬元），占合併營收由 {TH[f'E{THB["robsh"]}'].value:.1%} 升到 {TH[f'K{THB["robsh"]}'].value:.1%}。整本重算反解：其他輸入不變時，**支持現價需要人形絲杠需求達線 4 的 {fn(sol[0][1], '{:,.1f}')} 倍**（{sol[0][3]}）；評等升一級需要 {fn(sol[1][1], '{:,.2f}')} 倍；只放大份額時，{('情境 B 2030 年份額達 100% 仍不足以支持現價' if sol[2][1] == -1 else f'需要情境 B 2030 年份額 {sol[2][1]:.0%}')}。結論主要建立在 Assumed（2026 年機器人零部件營收、份額、毛利率路徑）與線模型的 Assumed／Analogy 輸入上；傳統本業與淨資產部分為 Verified／Derived。

## 命題

1. **營收結構：** 公司分產品金額（軸承、汽車零部件含熱管理）在可取得的資料中未找到，只有 2025 年增速；傳統本業以殘差列示（2025 年 {wb['Hist']['E8'].value:,.0f} 百萬元，占 {wb['Hist']['E8'].value / wb['Hist']['E4'].value:.1%}），機器人零部件 2025 年 {wb['Hist']['E9'].value:,.1f} 百萬元（首次出現）。
2. **機器人零部件：** ＝線 4 全球人形滾柱絲杠需求 × 本公司份額 × 單價；單價 2025 年均價 {I['asp25']:,.0f} 元，每年降 {I['pdr']:.0%}（線 6）。2026 年營收為 Assumed {I['rob26']:,.0f} 百萬元（區間 25–90），隱含份額 {TH[f'E{THB["sh"]}'].value:.0%}。
3. **傳統本業：** FY2026 以前三季實績加第 4 季假設為錨，之後年增 {I['gcore_B']:.0%}（區間 0–5%）；毛利率走向 {I['gm30_B']:.0%}。

情境 B：營收 {ser('rev', '{:,.0f}')}；毛利率 {ser('gm', '{:.1%}')}；EPS {ser('eps', '{:.2f}')} 元（FY2026→FY2032）。

## 主要假設與可證偽指標（原則第 10 節第 20 條）

| 承重假設 | 基準（A／B／C 或區間） | 標記 | 可觀察指標 |
|---|---|---|---|
| 全球人形滾柱絲杠需求（線 4） | 線 4 高／基準／低；× hu_k {I['hu_k']:.0f}（高案 2） | Assumed（線 4） | IDC、TrendForce 季出貨 |
| 單價年降幅（線 6） | {I['pdr']:.0%}（8–20%） | Analogy | 本公司機器人零部件營收；雙林、新劍報價 |
| 2030 年絲杠份額 | {I['sh30_A']:.0%}／{I['sh30_B']:.0%}／{I['sh30_C']:.0%} | Assumed | 客戶揭露、出貨量 |
| 2026 年機器人零部件營收 | {I['rob26']:,.0f}（25–90） | Assumed | 半年報、年報分產品 |
| 傳統本業年增 | {I['gcore_A']:.0%}／{I['gcore_B']:.0%}／{I['gcore_C']:.0%} | Assumed | 季報營收、熱管理訂單 |
| FY2030 毛利率 | {I['gm30_A']:.0%}／{I['gm30_B']:.0%}／{I['gm30_C']:.0%} | Assumed | 季毛利率（2026 前三季 19.1%） |
| 長期投資計入比例 | {I['inv_k']:.0%}（60–100%） | Assumed | 洛軸股份收盤價、鎖定期 |
| 同業前瞻 P/E | 中位數 {I['pe_peer']:.1f} 倍（6 家） | Derived | 同一估值日重抓 |

## 敏感度（整本重算，前 10 大）

| 案例 | 鍵 | 加權目標價 | 相對基準 | DCF 目標價 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | `{c['key']}` | {c['tp']:,.1f} | {c['tp'] / sens['base_tp'] - 1:+.1%} | {c['dcf']:,.1f} | {c['rating']} |" for c in top) + f"""

完整案例見 xlsx Sens 頁。**翻轉評等所需的變動**：見 Recon 反解；升一級需加權目標價達 {fn(up_b)} 元。

## 反解（Recon 頁）

""" + "\n".join(f"- {a}：{fn(b, '{:,.2f}')}（基準 {fn(c0, '{:,.2f}')}）；{d}" for a, b, c0, d in sol) + f"""

## 限制與「分不出來／沒找到」

- **沒找到：** 分產品金額（軸承、汽車零部件、熱管理）；2026 年機器人零部件營收；客戶集中度最新值（最新可得為 2016 年）；公司年報與三季報原文未直接讀取（歷史數字取自新浪財經、stockanalysis.com、同花順；三季報淨利拆分取自搜尋摘要轉述）；洛軸股份持股成本與鎖定期；限制性股票與選擇權。
- **疑點（沒動）：** (1) 產業網站稱 2025 年絲杠出貨 8.7 萬根、均價 1,360 元，相乘約 118 百萬元，但公司口徑機器人零部件營收只有 16.0 百萬元，差約 7 倍，本模型以公司營收為準。(2) 資料商每股股利 0.12 元與現金流量表股利支出（約 0.22–0.28 元／股）不符，配息率採現金支出比。(3) 山西證券預測早於三季報，不採用。
- **線模型不一致：** 2030 年人形滾柱絲杠需求線 6＝4.0 百萬根、線 4＝2.1 百萬根（採線 4）；單價採中國量產價口徑（線 6 低案）而非線 6 基準；本公司 2025 年均價年減 39% 高於線 6 高案 20%。見 xlsx「對帳」頁。
- **同業倍數：** 6 家進中位數，原報價日 2026-09-30～11-24，前瞻 P/E 依 2026-11-20 收盤等比例換算（EPS 未更新）；貝斯特（原報價日未取得）、中大力德（無前瞻 P/E）未進中位數；少於 8 家不做迴歸。北特科技前瞻 P/E 92 倍偏高，中位數不受影響。
- **洛軸股份：** 301699 不在收盤檔，股價取資料商舊日期值，僅作持股市值對照，不進公式。
- **rf 逾期：** 中國 10 年期國債 1.61% 為 2026-11-05 值（跨模型一致）。
- **協調檔：** 實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics；相關條文不適用。
- **獨立核對：** 依專案規則，交付判讀的數字應由未參與製作的代理獨立核對；待外層流程安排。

## 資料來源

見 xlsx Sources 頁（新浪財經利潤表、stockanalysis.com、同花順、搜尋摘要轉述的財報透視、山西證券轉載、線 4／6 成品）。

## 版本紀錄

- {VER}（2026-11-24）：首版。現價 {I['px']:,.2f}（2026-11-20 收盤）；加權目標價 {g('tp'):,.1f}；上檔空間 {g('up'):+.1%}；{g('rating')}／{g('rec')}；收支關卡 {g('gate_b'):,.1f}（{g('gate_t')}）。
"""
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write(t)


if __name__ == "__main__":
    main()

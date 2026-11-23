# THK（6481.T）個股模型 v0.1 — AI實體應用 Project（線 6 共用零組件：線性滑軌、滾珠／行星滾柱絲杠；並讀線 3 工業機器人、線 4 人形）
# 用法：python3 physical_ai/thk/build_thk.py   → 建置、整本重算 Sens 與反解（LibreOffice 批次）、輸出 dist/ 成品、README.md、sens.json、segments.json
# 會計年度＝曆年（12 月決算，IFRS）。金額：十億日圓（¥B）；股數：百萬股；EPS、股價：日圓。
# 汽車與運輸機器事業 2026-06-01 轉讓（AP87），FY2025 起列停業單位；本模型只建持續營業（產業機器）事業。
# 結構：README、In、Hist、Q（官方所在地別分部＋日本單體受注）、PL（三情境：命題模組→地區營收→標準損益表→現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 不讀 Siliconomics／Tokenomics／Wattonomics 協調檔（實體 AI 線）。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261123"
NAME = f"{DATE}_THK個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032

# ---------------- 出處 ----------------
TS3Q = "公司 FY2026 第 3 四半期決算短信〔IFRS〕（英文版，2026-11-13）https://www.thk.com/jp/wordpress/wp-content/uploads/2026/11/261113_financial_result_en.pdf"
FB3Q = "公司 Investor Information FY2026-3Q（2026-11-13）https://www.thk.com/jp/wordpress/wp-content/uploads/2026/11/Investor_Information_FY2026-3Q_en.pdf"
TS25 = "公司 FY2025 決算短信〔IFRS〕（英文版，2026-02-12；持續營業口徑重編 FY2024）https://www.thk.com/jp/wordpress/wp-content/uploads/2026/02/260212_financial_result_en.pdf"
PR2Q = "公司 FY2026 第 2 四半期決算補足資料（2026-08-05）https://www.thk.com/jp/wordpress/wp-content/uploads/2026/08/260805_presentation_en.pdf"
QTR = "Quartr THK Q3 2026 earnings summary（資料商摘要，2026-11-13）https://quartr.com/events/thk-co-ltd-6481-q3-2026_F9NplsXB"
SA = "stockanalysis.com THK statistics（資料商；頁面 2026-10-19）https://stockanalysis.com/quote/tyo/6481/statistics"
YCL = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
JGB10 = "日本財務省 國債金利情報 10 年：2026-11-02 2.817%（與 physical_ai/nabtesco In!E90、AI半導體 東京精密 v0.3.6 同一取值，跨模型一致；逾期）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"
L6 = "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx"
L3 = "physical_ai/lines/industrial_robot/dist/20261116_工業機器人單位經濟_v0.1.xlsx"
L4 = "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx"
HWX = "physical_ai/hiwin/dist/20261123_上銀個股模型_v0.1.1.xlsx"
HDX = "physical_ai/harmonic_drive/dist/20261123_哈默納科個股模型_v0.1.1.xlsx"
NBX = "physical_ai/nabtesco/dist/20261123_Nabtesco個股模型_v0.1.1.xlsx"

# ---------------- In 列 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源1, 來源2／說明, 來源誘因)
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

REG = [("jp", "日本"), ("am", "美洲"), ("eu", "歐洲"), ("cn", "中國"), ("as", "亞洲及其他")]
sec("A. FY2025 實績：持續營業（產業機器）營收依客戶地區（汽車與運輸機器事業已列停業單位）")
for (k, n), v in zip(REG, (96.415, 27.168, 21.611, 71.188, 24.059)):
    inp(f"a25_{k}", f"營收 FY2025：{n}（客戶地區，持續營業）", "¥B", None, v, None, "Verified", FB3Q + " p.14 地區別營收", "FY2025 美洲、歐洲、亞洲不含已轉讓事業；五地區合計 240.441 ＝ 合併營收 240.444（取整差 0.003）")

sec("B. FY2026 錨點：公司指引（2026-11-13 第三次上修；Interested-party）")
for (k, n), v in zip(REG, (118.0, 35.5, 27.0, 104.5, 35.0)):
    inp(f"g26_{k}", f"公司指引：FY2026 營收 {n}", "¥B", None, v, None, "Interested-party", FB3Q + " p.14 FY2026E 地區別", "五地區合計 320.0 ＝ 合併營收指引", CO)
inp("g26_rev", "公司指引：FY2026 合併營收", "¥B", None, 320.0, None, "Interested-party", TS3Q + "：320,000（+33.1%）", "指引沿革：2026-02 260.0 → 05 276.0 → 08 310.0 → 11 320.0（" + PR2Q + "）；9M 實績 235.939 → Q4 隱含 84.1（Derived）", CO)
inp("g26_gp", "公司指引：FY2026 毛利", "¥B", None, 111.2, None, "Interested-party", FB3Q + " p.14：毛利 111,200", "9M 實績 81.651（毛利率 34.6%）", CO)
inp("g26_sga", "公司指引：FY2026 銷管費用", "¥B", None, 59.5, None, "Interested-party", FB3Q + " p.14：SG&A 59,500", "9M 實績 43.934", CO)
inp("g26_op", "公司指引：FY2026 營業利益", "¥B", None, 52.0, None, "Interested-party", TS3Q + "：52,000（+260.2%）", "前次 48.0（2026-08）；9M 實績 37.741 → Q4 隱含 14.3（Q3 單季 15.3，Derived）", CO)
inp("g26_pbt", "公司指引：FY2026 稅前利益", "¥B", None, 51.8, None, "Interested-party", TS3Q + "：51,800", "", CO)
inp("g26_ni", "公司指引：FY2026 歸屬母公司淨利（含停業單位）", "¥B", None, 58.8, None, "Interested-party", TS3Q + "：58,800；EPS 524.66", "含事業轉讓利益 16.436（停業單位）", CO)
inp("g26_eps", "公司指引：FY2026 基本 EPS（含停業單位）", "¥", None, 524.66, None, "Interested-party", TS3Q + "：524.66", "", CO)
inp("g26_disc", "FY2026 停業單位損益（9M 實績；轉讓已於 2026-06-01 完成，Q4 視為 0）", "¥B", None, 15.750, None, "Derived", TS3Q + "：9M 停業單位 15,750（含轉讓利益 16,436）", "Q4＝0 為假設（轉讓完成）")
inp("g26_nci", "FY2026 非支配股東損益（9M 實績）", "¥B", None, 0.092, None, "Verified", TS3Q + "：9M 92", "非支配股東權益隨汽車事業轉出，2026-09-30 為 0")
inp("da_26", "公司計畫：FY2026 折舊", "¥B", None, 18.2, None, "Interested-party", FB3Q + " p.19：減価償却費 18,200（合併）", "9M 實績 16.845（含轉讓前汽車事業）", CO)
inp("cx_26", "公司計畫：FY2026 資本支出", "¥B", None, 21.0, None, "Interested-party", FB3Q + " p.19：設備投資 21,000（合併）", "9M 實績 11.403；大連滾珠絲杠廠二期 2027-12 完工（" + PR2Q + "）", CO)
inp("dps_26", "FY2026 每股股利（中間 96 實績＋期末 99 予想）", "¥", None, 195, None, "Interested-party", TS3Q + "：年間 195；DOE 8% 方針", "", CO)
inp("dps_q4", "FY2026 期末股利（2027 年支付）", "¥", None, 99, None, "Interested-party", TS3Q, "", CO)
inp("g_anc", "FY2026 實績相對公司指引（營收、毛利、費用同比例）", "%", -0.02, 0.0, 0.02, "Assumed", "基準＝公司指引（9M 已實績，偏離只來自 Q4）", "Q4 隱含營收 84.1 ≈ Q3 單季 84.9；低／高 ±2%")

sec("C. 命題模組一：產業機器循環（工具機、半導體設備、電子、一般機械、機器人）→ 核心營收（五地區同一成長率）")
GTXT = {
    "g27": ("FY2027", {"A": (0.22, "樂觀：日本單體受注 2026Q3 495 億 vs 營收 281 億（訂單÷營收 1.76）、受注殘 570 億（2025 年底 183 億）；電子受注 2026Q3 138 億 vs 2025Q3 45 億（AI 半導體設備）"),
                       "B": (0.12, "基準：Q4 指引隱含季營收 84.1 年化 336（+5% 於 FY2026）＋受注殘消化；受注殘增加 387 億約當日本單體 1.4 季營收（Derived，" + FB3Q + " p.12）"),
                       "C": (0.0, "保守：受注含提前下單與重複下單，2027 年營收持平（2024–2025 年受注 YoY 為負的循環為類比）")}),
    "g28": ("FY2028", {"A": (0.08, "樂觀：半導體設備投資延續"), "B": (0.04, "基準：循環趨緩"), "C": (-0.10, "保守：半導體與工具機循環下行（2023–2024 年持續營業營收約 −5% 為類比；FY2024 222.7 → FY2025 240.4）")}),
    "gm": ("FY2029–FY2030", {"A": (0.06, ""), "B": (0.04, "基準：工業機器人裝機 2026→2030 年複合 6.9%（線 3 下游取數 E13→E17）為上緣類比；工具機、半導體設備未在實體 AI 線"), "C": (0.03, "")}),
    "glt": ("FY2031–FY2032", {"A": (0.04, ""), "B": (0.03, "長期：名目產業機器需求"), "C": (0.02, "")}),
}
GRNG = {"g27": (0.05, 0.20), "g28": (-0.04, 0.08), "gm": (0.02, 0.06), "glt": (0.02, 0.04)}
for f, (yl, d) in GTXT.items():
    for s in SC:
        v, note = d[s]
        lo, hi = GRNG[f] if s == "B" else (None, None)
        inp(f"{f}_{s}", f"情境 {s}：核心產業機器營收年增 {yl}（五地區同一成長率）", "%/年", lo, v, hi, "Assumed", note or "A／C 以 B ± 約 2pt 設定",
            "地區別成長差異無證據（不分化）；區間見情境 A／B／C 三列" if s == "B" else "")

sec("C2. 命題模組二：人形機器人線性關節（地區未分）＝行星滾柱絲杠市場（線 6 下游取數）× THK 份額 × 爬坡 × 匯率")
for yy, rr in ((2026, 12), (2030, 13), (2032, 14)):
    for s, col in (("A", "D"), ("B", "C"), ("C", "B")):
        v = {("A", 2026): 731.5, ("B", 2026): 298.571428571429, ("C", 2026): 89.5714285714286,
             ("A", 2030): 6545.958375, ("B", 2030): 1558.56151785714, ("C", 2030): 233.784227678571,
             ("A", 2032): 13794.2435339844, ("B", 2032): 3378.18208995536, ("C", 2032): 337.818208995536}[(s, yy)]
        inp(f"m{str(yy)[2:]}_{s}", f"情境 {s}：行星滾柱絲杠市場 {yy}（線 6；{'高' if s == 'A' else '基準' if s == 'B' else '低'}需求情境）", "百萬美元", None, v, None, "Derived",
            f"{L6} 下游取數!{col}{rr}（基準單價口徑＝覓途；五線歸因，全來自人形）", "線 6 需求情境：A＝高（D 欄）、B＝基準（C 欄）、C＝低（B 欄）；年間以幾何插值（Derived）")
for s, v, lo, hi in (("A", 0.04, None, None), ("B", 0.02, 0.005, 0.05), ("C", 0.005, None, None)):
    inp(f"shr_{s}", f"情境 {s}：THK 在人形線性關節（以行星滾柱絲杠市場為可及市場代理）的份額", "%", lo, v, hi, "Assumed",
        f"未取得：THK 未揭露人形機器人營收或客戶；行星滾柱絲杠現有領先者為歐系（Schaeffler、Rollvis、GSA、SKF），THK 未列在清單（{L6} 份額!B11）", "代理口徑：THK 以滾珠絲杠、微型 LM 導軌、致動器參與（公司 2026-06 義肢膝關節用滾珠絲杠＋LM 導軌案例，非人形）；份額數據未找到")
inp("rp27", "人形營收爬坡：FY2027 占穩態比例（線性升到 FY2030＝100%）", "%", 0.0, 0.25, 0.5, "Assumed", "比照上銀模型 Thesis 頁爬坡 2027 25% → 2030 100%（physical_ai/hiwin In 頁）", "")
inp("fx", "美元兌日圓（人形市場換算）", "¥/US$", 140.0, 157.58, 165.0, "Interested-party", FB3Q + " p.14：FY2026 平均匯率假設 157.58", "低／高為 Assumed", CO)

sec("D. 損益表結構（標準損益表；IFRS）")
inp("mm_gp", "邊際毛利率（Δ毛利 ÷ Δ營收，相對 FY2026）", "x", 0.30, 0.40, 0.50, "Assumed",
    "9M'26 vs 9M'25：(81.651 − 51.875) ÷ (235.939 − 175.480) ＝ 49.2%（Derived，" + TS3Q + "；含結構改革與價格）", "FY2024→FY2025：(70.495 − 68.249) ÷ (240.444 − 222.737) ＝ 12.7%（" + TS25 + "）；基準取 40%，減收時同率回落")
for s, v, lo, hi in (("A", 0.01, None, None), ("B", 0.0, -0.02, 0.01), ("C", -0.02, None, None)):
    inp(f"gmadj_{s}", f"情境 {s}：FY2027 起毛利率額外調整（價格、中國同業競爭）", "pt", lo, v, hi, "Assumed",
        "毛利率：FY2024 30.6%、FY2025 29.3%、9M'26 34.6%、FY2026 指引 34.75%（Derived）", "C：中國絲杠與導軌廠擴產（physical_ai/hiwin Comps 頁貝斯特、恒立、五洲新春）")
inp("sga_el", "銷管費用對營收的彈性（費用年增 ÷ 營收年增）", "x", 0.2, 0.3, 0.6, "Derived",
    "9M'26：SG&A +7.9% ÷ 營收 +34.5% ＝ 0.23；FY2026 指引：+9.5% ÷ +33.1% ＝ 0.29（" + TS3Q + "、" + FB3Q + "）", "產業機器結構改革 FY2027–29 目標改善 200 億以上（" + PR2Q + "，Interested-party）未另計，低案 0.2 部分反映")
inp("oth_op", "其他營業損益淨額（其他收益 − 其他費用 ＋ 持分法）FY2027 起", "¥B", -1.5, 0.3, 1.0, "Assumed", "FY2026 指引隱含 0.3（＝52.0 − 111.2 ＋ 59.5，Derived）", "9M'26 0.025、FY2025 −1.718、FY2024 1.313")
inp("nonop", "業外：金融收益 − 金融費用（FY2027 起）", "¥B", -1.5, -0.2, 1.5, "Assumed", "FY2026 指引隱含 −0.2（＝51.8 − 52.0，Derived）", "9M'26 +0.035、FY2025 +1.310、FY2024 +1.947；借款 1,498 億、現金 1,183 億")
inp("tax_w", "有效稅率（FY2027 起）", "%", 0.20, 0.28, 0.32, "Assumed", "日本法定實效約 30.6%、中國 25%（類比）", "實績：FY2024 37.3%、FY2025 32.9%、9M'26 17.2%（低，原因未查核）；FY2026 指引隱含約 16.7%（Derived）")

sec("E. 現金流與股東回饋")
inp("fc_da", "折舊攤銷 ÷ 營收（FY2027 起）", "%", 0.048, "=da_26/g26_rev", 0.065, "Interested-party", FB3Q + " p.19：FY2026 計畫折舊 182 億 ÷ 營收 3,200 億", "FY2025 持續營業 16.479 ÷ 240.444 ＝ 6.9%（減価償却費，合併）", CO)
inp("fc_cx", "資本支出 ÷ 營收（FY2027 起）", "%", 0.05, "=cx_26/g26_rev", 0.08, "Interested-party", FB3Q + " p.19：FY2026 計畫 210 億 ÷ 3,200 億", "FY2023 28.757、FY2024 30.712（含汽車）、FY2025 15.410；大連二期 2027-12 完工", CO)
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.20, 0.30, 0.45, "Assumed", "（應收＋存貨 − 應付）÷ 營收：2025-12-31 (63.528＋65.177 − 18.585) ÷ 240.444 ＝ 45.8%（Derived，" + TS25 + "）", "9M'26 營運資金增 15.668 ÷ 年化營收增量約 99 ＝ 16%（Derived）；基準 30%")
inp("doe", "股東權益配息率（DOE）", "%", None, 0.08, None, "Interested-party", TS3Q + "：DOE 8% 方針", "只用於收支關卡（股利＝DOE × 前一年底權益）；不進 DCF", CO)
inp("eq_26", "FY2026 年底歸屬母公司權益（2026-09-30 權益 ＋ Q4 淨利）", "¥B", None, "=270.631+(g26_ni-46.935)", None, "Derived", TS3Q + "：2026-09-30 權益合計 270.631（非支配 0）、9M 歸屬母公司淨利 46.935", "期末股利 99（2027-03 支付）不在 2026 扣除")

sec("F. 估值輸入（評價日 2026-12-31；資產負債表 2026-09-30）")
inp("px", "股價（2026-11-20 收盤）", "¥", None, 4249, None, "Verified", YCL + "：6481.T 2026-11-20 收盤 4,249", "未還原股利。第 3 季決算（2026-11-13 盤後）後：11-13 收盤 5,472、11-16 4,684（−14.4%）；下跌原因未查核")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("sh_out", "流通股數（2026-09-30；發行 119,099,803 − 自己株式 7,027,344）", "百萬股", None, "=119.099803-7.027344", None, "Derived", TS3Q + "：期末發行済株式数、自己株式数（Verified）", "9M 庫藏股買回僅 2 百萬日圓；股票報酬與選擇權：未取得，視為 0")
inp("sh_v", "評價日股數", "百萬股", None, "=sh_out", None, "Derived", "＝流通股數（無進行中的買回額度）", "DCF 標準規格第 3 節；公司「彈性取得自己株式」未建模")
inp("cash_v", "現金及約當現金＋其他流動金融資產（2026-09-30）", "¥B", None, "=118.325+2.520", None, "Derived", TS3Q + "：現金 118,325、其他流動金融資產 2,520（Verified）", "")
inp("debt_v", "社債及借款（2026-09-30；流動 80,002＋非流動 69,826）", "¥B", None, "=80.002+69.826", None, "Derived", TS3Q + "（Verified 原值相加）", "")
inp("lease_v", "租賃負債代理（2025-12-31 其他金融負債：流動 2,884＋非流動 6,621）", "¥B", 0.0, 9.505, None, "Derived", TS25 + "：租賃負債未單獨揭露，含在其他金融負債內（上限代理）", "低案 0；轉讓後應較低（未取得 2026-09-30 值）")
inp("inv_v", "其他非流動金融資產（2026-09-30）", "¥B", None, 24.896, None, "Verified", TS3Q + "：24,896（2025 年底 10,880）", "內容未拆分（增加可能與事業轉讓相關，未查核）")
inp("inv_k", "其他非流動金融資產計入比例", "x", 0.5, 1.0, None, "Assumed", "基準帳面全額；低案 50%", "")
inp("mi_v", "非支配股東權益（2026-09-30）", "¥B", None, 0.0, None, "Verified", TS3Q + "：—（汽車事業轉出後為 0）", "")
inp("stub_frac", "Stub 期（2026-09-30 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,9,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（日本 10 年期國債）", "%", None, 0.02817, None, "Verified", JGB10, "")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "日本國家風險溢酬", "%", None, 0.0091, None, "Verified", "Damodaran 國家風險溢酬（2026-01-05）：日本 0.91%（與 physical_ai/nabtesco In!E92 同值）", "")
inp("beta", "β（兩來源 Blume 調整後中位數）", "x", "=E{r}-0.3", "=MEDIAN(0.67*0.76+0.33,0.67*MEDIAN(0.53,1.04,0.99,1.08)+0.33)", "=E{r}+0.3", "Derived",
    SA + "：THK 5 年 β 0.76；線性傳動同業 5 年 β：NSK 0.53、IKO 1.04、直得 0.99（" + HWX + " Comps!J6:J8）、上銀 1.08（同檔 Comps!J16）", "Blume 0.67β＋0.33；敏感度 ±0.3（Assumed）")
inp("kd", "稅前負債成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt（銀行借款與社債；9M'26 金融費用 1.007 含匯損，未拆利息）", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("dps_tgt", "評價日到目標時點之每股股利", "¥", None, "=2*dps_q4", None, "Assumed", "FY2026 期末 99（2027-03 支付）＋FY2027 中間（假設同 99）", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable：同業前瞻 P/E）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺 THK 對同業 10 年前瞻 P/E 比值歷史）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業前瞻 P/E（取自同專案成品；報價日各異：只有 Fanuc、安川、Nabtesco 已換算至 2026-11-20，其餘報價日未更新）")
PEERS = [("pe_nsk", "NSK 6471（軸承、滾珠絲杠、導軌）", 15.35, "Interested-party", HWX + " Comps!F6（資料日 2026-09-18；報價日未更新）"),
         ("pe_iko", "Nippon Thompson（IKO）6480（線性導軌）", 11.81, "Interested-party", HWX + " Comps!F7（2026-09-04；報價日未更新）"),
         ("pe_cdg", "直得 1597（微型導軌）", 58.97, "Interested-party", HWX + " Comps!F8（2026-08-24；報價日未更新）"),
         ("pe_hiwin", "上銀 2049（線性滑軌、滾珠螺桿）", 28.35, "Interested-party", HWX + " Comps!F16（2026-10-22；報價日未更新）"),
         ("pe_fanuc", "Fanuc 6954（機器人、FA）", 21.1499, "Derived", HWX + " Comps!F12（換算至 2026-11-20）"),
         ("pe_yask", "安川電機 6506（伺服、機器人）", 15.6739, "Derived", HWX + " Comps!F13（換算至 2026-11-20）"),
         ("pe_nab", "Nabtesco 6268（精密減速機）", 13.7701, "Derived", HDX + " Comp!F5（換算至 2026-11-20）"),
         ("pe_smc", "SMC 6273（氣動元件）", 21.35, "Interested-party", NBX + " In!E113（資料日 2026-10-09；報價日未更新）")]
for k, lab, v, tag, src in PEERS:
    inp(k, f"前瞻 P/E：{lab}", "x", None, v, None, tag, src, "原值：stockanalysis.com 賣方共識 EPS；會計口徑：日本 IFRS 或日本基準、台灣 IFRS，P/E 為股權口徑不受租賃處理影響", DV)
inp("pe_med", "同業前瞻 P/E 中位數（8 家；進 Comparable）", "x", 13.77, "=MEDIAN(pe_nsk,pe_iko,pe_cdg,pe_hiwin,pe_fanuc,pe_yask,pe_nab,pe_smc)", 21.85, "Derived",
    "8 家中位數；不含哈默納科 84.44 倍（人形溢價離群，" + HWX + " Comps!F14）", "低＝Nabtesco 換算值 13.77（Analogy：已反映 11 月回落的同業）；高＝線性傳動 4 家中位數 21.85（Derived）；少於約 8 家不做迴歸")
inp("pe_self", "THK 現行前瞻 P/E（換算至 2026-11-20）", "x", None, 10.6781, None, "Derived", HWX + " Comps!F5：資料商 2026-10-19 頁面 15.37 × 4,249 ÷ 6,116（分母 EPS 不變）", "只作對照，不進模型", DV)
inp("cons_tp", "賣方平均目標價（12 位；資料商 2026-10-19 頁面，第 3 季決算前）", "¥", None, 8251.67, None, "Interested-party", SA + "：平均 8,251.67、評等 Buy", "只作對照，不進模型", DV)


def build(path, ov=None, sens=None, rev=None, bid="{BID}"):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 ¥B、股數百萬股、EPS 與股價日圓"
    IN["A2"] = "標記：Verified（附出處）／Interested-party（寫明是誰、誘因欄）／Analogy（寫明類比對象）／Assumed（給區間）／Derived（寫明換算）；查不到寫「未取得」"
    hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源1", "來源2／說明", "來源誘因"])
    r = 4; KR = {}
    for row in ROWS:
        if row[0] == "sec":
            IN.cell(r, 1, row[1]).font = BOLD; r += 1; continue
        k, lab, u, lo, b, hi, tag, s1, s2, inc = row
        if k in ov: b = ov[k]
        IN.cell(r, 1, k); IN.cell(r, 2, lab); IN.cell(r, 3, u)
        for col, v in ((4, lo), (5, b), (6, hi)):
            if v is None: continue
            if isinstance(v, str): v = v.replace("{r}", str(r))
            c = IN.cell(r, col, v)
            if col == 5: c.font = BLUE
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2); IN.cell(r, 10, inc)
        dn(k, f"In!$E${r}"); KR[k] = r; r += 1
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("g27", "核心營收年增 FY2027"), ("g28", "核心營收年增 FY2028"), ("gm", "核心營收年增 FY2029–FY2030"), ("glt", "核心營收年增 FY2031–FY2032"),
           ("m26", "行星滾柱絲杠市場 2026（線 6）"), ("m30", "行星滾柱絲杠市場 2030（線 6）"), ("m32", "行星滾柱絲杠市場 2032（線 6）"),
           ("shr", "THK 人形線性關節份額"), ("gmadj", "FY2027 起毛利率調整")]
    for f, lab in FAM:
        IN.cell(r, 1, f + "_scn"); IN.cell(r, 2, lab)
        for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR[f + '_' + s]}")
        IN.cell(r, 8, f"引用 In 頁 {f}_A／{f}_B／{f}_C；{lab}"); r += 1
    IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重")
    for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR['w_' + s]}")
    IN.cell(r, 8, "引用 In 頁 w_A／w_B／w_C；合計＝1"); r += 1
    IN.column_dimensions["B"].width = 56; IN.column_dimensions["H"].width = 60

    # ---------------- Hist ----------------
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（年度；¥B）。FY2024、FY2025＝持續營業口徑（汽車與運輸機器事業列停業單位，FY2025 決算短信重編）；FY2023＝原公告（含汽車事業，舊口徑，與後兩年不可比）。公式只引用本頁（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2023（舊口徑）", "FY2024", "FY2025", "標記", "來源"])
    NA = "未取得（舊口徑）"
    HR = [("h_rev", "營收", 351.939, 222.737, 240.444, "Verified", FB3Q + " p.14（FY2023）；" + TS25),
          ("h_cogs", "營業成本", "=B3-B5", 154.487, 169.949, "Verified", "FY2023＝營收 − 毛利（Derived）；" + TS25),
          ("h_gp", "毛利", 81.630, 68.249, 70.495, "Verified", FB3Q + "；" + TS25),
          ("h_sga", "銷管費用", 58.958, 53.639, 54.341, "Verified", FB3Q + "；" + TS25),
          ("h_oth", "其他營業損益淨額（其他收益 − 其他費用 ＋ 持分法）", "=B8-B5+B6", "=2.091-0.833+0.055", "=1.450-1.581-1.587", "Derived", "FY2023＝營業利益 − 毛利 ＋ 銷管費用；" + TS25),
          ("h_op", "營業利益", 23.707, 15.923, 14.436, "Verified", FB3Q + "；" + TS25),
          ("h_fin", "業外：金融收益 − 金融費用", NA, "=3.036-1.089", "=3.048-1.738", "Derived", TS25),
          ("h_pbt", "稅前利益", NA, 17.870, 15.746, "Verified", TS25),
          ("h_tax", "所得稅", NA, 6.659, 5.185, "Verified", TS25),
          ("h_cont", "持續營業稅後淨利", NA, 11.210, 10.561, "Verified", TS25),
          ("h_disc", "停業單位損益", NA, -0.552, -80.040, "Verified", TS25 + "（FY2025 含處分群組公允價值減損 81.639）"),
          ("h_ni", "歸屬母公司淨利", 18.398, 10.439, -69.891, "Verified", FB3Q + "；" + TS25),
          ("h_nci", "非支配股東損益", NA, 0.217, 0.412, "Verified", TS25),
          ("h_eps", "基本 EPS（含停業單位，日圓）", 150.08, 85.17, -618.66, "Verified", FB3Q + "；" + TS25),
          ("h_epsc", "基本 EPS（持續營業，日圓）", NA, 88.30, 89.85, "Verified", TS25),
          ("h_da", "折舊（合併，含汽車事業至轉讓）", 18.898, 20.808, 16.479, "Verified", FB3Q + " p.19"),
          ("h_cx", "設備投資（合併）", 28.757, 30.712, 15.410, "Verified", FB3Q + " p.19"),
          ("h_ocf", "營業現金流（含停業單位）", NA, 28.412, 42.748, "Verified", TS25),
          ("h_divp", "股利支付", NA, 4.167, 29.359, "Verified", TS25),
          ("h_bb", "自己株式取得", NA, 4.020, 36.519, "Verified", TS25)]
    for i, (k, lab, a, b, c, tag, src) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b, c)):
            cc = H.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): cc.font = BLUE
        H.cell(i, 5, tag); H.cell(i, 6, src)
        dn(k + "_23", f"Hist!$B${i}"); dn(k + "_24", f"Hist!$C${i}"); dn(k + "_25", f"Hist!$D${i}")
    H.column_dimensions["A"].width = 44
    HRW = {k: 3 + i for i, (k, *_x) in enumerate(HR)}

    # ---------------- Q（官方分部＋日本單體受注） ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 官方報告分部（所在地別，含已轉讓的汽車與運輸機器事業至 2026-05）9M 累計外部營收與日本單體受注（¥B）。分部對帳：分部合計 − 停業單位重分類 − 合併營收（持續營業）"
    hdr(Q, 2, ["項目", "9M'25", "9M'26", "標記", "來源", "來源誘因"])
    segs = [("日本（所在地別）營收", 80.416, 95.602), ("美洲（所在地別）營收", 66.477, 55.572), ("歐洲（所在地別）營收", 50.140, 40.820),
            ("中國（所在地別）營收", 55.763, 78.712), ("其他（所在地別）營收", 15.429, 20.616)]
    for i, (lab, a, b) in enumerate(segs, 3):
        Q.cell(i, 1, lab); Q.cell(i, 2, a).font = BLUE; Q.cell(i, 3, b).font = BLUE; Q.cell(i, 4, "Verified"); Q.cell(i, 5, TS3Q + " 分部情報（外部顧客營收）")
    Q.cell(8, 1, "報告分部合計"); Q.cell(8, 2, "=SUM(B3:B7)"); Q.cell(8, 3, "=SUM(C3:C7)"); Q.cell(8, 4, "Derived"); Q.cell(8, 5, "加總")
    Q.cell(9, 1, "停業單位重分類（汽車與運輸機器事業）"); Q.cell(9, 2, -92.746).font = BLUE; Q.cell(9, 3, -55.384).font = BLUE; Q.cell(9, 4, "Verified"); Q.cell(9, 5, TS3Q)
    Q.cell(10, 1, "合併營收（持續營業）"); Q.cell(10, 2, 175.480).font = BLUE; Q.cell(10, 3, 235.939).font = BLUE; Q.cell(10, 4, "Verified"); Q.cell(10, 5, TS3Q)
    Q.cell(11, 1, "檢查：分部合計 − 停業重分類 − 合併營收（四捨五入差 < 0.005）"); Q.cell(11, 2, "=ROUND(B8+B9-B10,3)"); Q.cell(11, 3, "=ROUND(C8+C9-C10,3)"); Q.cell(11, 4, "Derived"); Q.cell(11, 5, "百萬日圓取整差 1")
    Q.cell(12, 1, "分部利益合計（含停業單位；9M'26 含事業轉讓利益）"); Q.cell(12, 2, 9.840).font = BLUE; Q.cell(12, 3, 53.275).font = BLUE; Q.cell(12, 4, "Verified"); Q.cell(12, 5, TS3Q + "：調整 −0.111／0.270、停業重分類 0.033／−15.804 → 營業利益 9.761／37.741")
    Q.cell(13, 1, "註：官方分部依公司所在地、含已轉讓事業至 2026-05，無法向前推估；PL 頁營收改用公司揭露的持續營業「客戶地區別」營收（公司 Investor Information FY2026-3Q p.14），兩者口徑不同")
    hdr(Q, 15, ["日本單體（THK 本體）", "受注", "營收", "標記", "來源", "來源誘因", "受注殘", "受注 ÷ 營收"])
    ords = [("2024Q1", 19.768, 20.028, 19.984), ("2024Q2", 22.109, 22.387, 19.706), ("2024Q3", 22.374, 22.048, 20.032), ("2024Q4", 21.289, 23.941, 17.380),
            ("2025Q1", 19.936, 20.703, 16.613), ("2025Q2", 21.974, 20.689, 17.898), ("2025Q3", 21.294, 21.710, 17.482), ("2025Q4", 23.765, 22.928, 18.319),
            ("2026Q1", 29.780, 22.773, 25.325), ("2026Q2", 37.357, 27.180, 35.502), ("2026Q3", 49.548, 28.082, 56.967)]
    for i, (q, o, s_, bl) in enumerate(ords, 16):
        Q.cell(i, 1, q); Q.cell(i, 2, o).font = BLUE; Q.cell(i, 3, s_).font = BLUE; Q.cell(i, 7, bl).font = BLUE
        Q.cell(i, 8, f"=B{i}/C{i}"); Q.cell(i, 4, "Verified"); Q.cell(i, 5, FB3Q + " p.12（単体受注・売上・受注残）")
    q0 = 16 + len(ords) + 1
    hdr(Q, q0, ["日本單體受注：產業別", "工具機", "一般機械", "標記", "來源", "來源誘因", "電子（含半導體）"])
    ind = [("2025Q3", 3.060, 4.009, 4.536), ("2025Q4", 3.646, 4.011, 5.441), ("2026Q1", 3.918, 4.968, 6.740), ("2026Q2", 4.700, 7.051, 9.255), ("2026Q3", 6.538, 8.345, 13.790)]
    for i, (q, a, b, c) in enumerate(ind, q0 + 1):
        Q.cell(i, 1, q); Q.cell(i, 2, a).font = BLUE; Q.cell(i, 3, b).font = BLUE; Q.cell(i, 7, c).font = BLUE
        Q.cell(i, 4, "Verified"); Q.cell(i, 5, FB3Q + " p.13")
    Q.column_dimensions["A"].width = 46

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（核心產業機器循環＋人形線性關節）→ 地區營收 → 標準損益表（IFRS，持續營業）→ 現金流（¥B；FY2023 舊口徑、FY2024–25 實績、FY2026 公司指引 ×(1＋g_anc)、FY2027 起模型）"
    PR = {}
    rr = 3
    KEYS = ([f"c_{k}" for k, _ in REG] + ["core", "hum", "rev", "chk1", "chk2", "gr", "cogs", "gp", "gm", "sga", "oth", "opex", "op", "opm", "fin", "pbt", "tax", "cont",
             "disc", "nci", "ni", "nim", "sh", "eps", "epsc", "nig", "da", "cx", "dwc", "fcff", "eq", "div", "fcfad", "mkt", "rp"])
    LAB = ([f"營收：{n}（客戶地區，核心產業機器）" for _, n in REG] +
           ["核心產業機器營收（五地區合計）", "人形機器人線性關節（命題子列，地區未分）", "營收（合併）", "檢查：地區合計 − 核心營收", "檢查：核心＋人形 − 合併營收",
            "營收年增率", "營業成本", "毛利", "毛利率", "銷售及管理費用", "其他營業損益淨額（其他收益 − 其他費用 ＋ 持分法）", "營業費用（銷管費用 − 其他營業損益淨額）",
            "營業利益", "營業利益率", "業外：金融收益 − 金融費用（其他收入淨額）", "稅前利益", "所得稅", "持續營業稅後淨利", "停業單位損益（汽車與運輸機器事業）",
            "非支配股東損益", "稅後淨利（歸屬母公司）", "淨利率", "平均股數（百萬股）", "EPS（基本，日圓）", "EPS（持續營業，日圓）", "淨利年增率",
            "折舊攤銷", "資本支出", "營運資金增加", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "歸屬母公司權益（年底）", "股利", "自由現金流扣股利（收支關卡用）",
            "參考：行星滾柱絲杠市場（百萬美元，線 6）", "參考：人形營收爬坡（占穩態）"])
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, k in enumerate(KEYS)}
        for i, l in enumerate(LAB): P.cell(b + i, 1, l)
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1); cell = lambda k: f"{c}{R[k]}"; prev = lambda k: f"{p}{R[k]}"
            def w(k, f): P[cell(k)] = f
            if y <= 2025:
                yy = str(y)[2:]
                hm = lambda k: f"=h_{k}_{yy}"
                for k, hk in (("rev", "rev"), ("cogs", "cogs"), ("gp", "gp"), ("sga", "sga"), ("oth", "oth"), ("op", "op"), ("fin", "fin"), ("pbt", "pbt"), ("tax", "tax"),
                              ("cont", "cont"), ("disc", "disc"), ("nci", "nci"), ("ni", "ni"), ("eps", "eps"), ("epsc", "epsc"), ("da", "da"), ("cx", "cx")):
                    w(k, hm(hk))
                w("gm", f"={cell('gp')}/{cell('rev')}"); w("opex", f"={cell('sga')}-{cell('oth')}"); w("opm", f"={cell('op')}/{cell('rev')}")
                w("nim", f"={cell('ni')}/{cell('rev')}")
                if y == 2025:
                    for k, _ in REG: w(f"c_{k}", f"=a25_{k}")
                    w("core", f"=SUM({c}{R['c_jp']}:{c}{R['c_as']})"); w("hum", 0); w("rp", 0)
                    w("chk1", f"=ROUND(SUM({c}{R['c_jp']}:{c}{R['c_as']})-{cell('core')},6)")
                    w("chk2", f"=ROUND({cell('core')}+{cell('hum')}-{cell('rev')},3)")
                    w("gr", f"={cell('rev')}/{prev('rev')}-1"); w("nig", '="N/M（FY2025 為虧損）"')
                    w("sh", f"={cell('ni')}/{cell('eps')}*1000")
                elif y == 2024:
                    w("gr", '="不可比（FY2023 含汽車事業）"'); w("nig", '="不可比（FY2023 含汽車事業）"'); w("sh", f"={cell('ni')}/{cell('eps')}*1000")
                else:
                    for k in ("gr", "nig"): w(k, '="未取得"')
                    w("sh", f"={cell('ni')}/{cell('eps')}*1000")
                    for k in ("fin", "pbt", "tax", "cont", "disc", "nci", "epsc"): w(k, '="未取得（舊口徑）"')
                    w("nim", f"={cell('ni')}/{cell('rev')}")
                continue
            if y == 2026:
                for k, _ in REG: w(f"c_{k}", f"=g26_{k}*(1+g_anc)")
                w("hum", 0); w("rp", 0)
                w("gp", "=g26_gp*(1+g_anc)"); w("sga", "=g26_sga*(1+g_anc)"); w("oth", "=(g26_op-g26_gp+g26_sga)*(1+g_anc)")
                w("fin", "=g26_pbt-g26_op"); w("disc", "=g26_disc"); w("nci", "=g26_nci")
                w("tax", f"={cell('pbt')}-(g26_ni-g26_disc+g26_nci)*(1+g_anc)")
                w("sh", "=g26_ni/g26_eps*1000")
                w("da", "=da_26"); w("cx", "=cx_26"); w("eq", "=eq_26"); w("div", f"=dps_26*{cell('sh')}/1000")
            else:
                gk = "g27" if y == 2027 else "g28" if y == 2028 else "gm" if y <= 2030 else "glt"
                for k, _ in REG: w(f"c_{k}", f"={prev('c_' + k)}*(1+{gk}_{s})")
                if y <= 2029: mk = f"=m26_{s}*(m30_{s}/m26_{s})^(({y}-2026)/4)"
                elif y == 2030: mk = f"=m30_{s}"
                elif y == 2031: mk = f"=m30_{s}*(m32_{s}/m30_{s})^0.5"
                else: mk = f"=m32_{s}"
                w("mkt", mk); w("rp", f"=MIN(1,rp27+(1-rp27)*({y}-2027)/3)")
                w("hum", f"={cell('mkt')}*shr_{s}*{cell('rp')}*fx/1000")
                w("gp", f"=$F{R['gp']}+mm_gp*({cell('rev')}-$F{R['rev']})+gmadj_{s}*{cell('rev')}")
                w("sga", f"=$F{R['sga']}*(1+sga_el*({cell('rev')}/$F{R['rev']}-1))"); w("oth", "=oth_op")
                w("fin", "=nonop"); w("tax", f"={cell('pbt')}*tax_w"); w("disc", 0); w("nci", 0); w("sh", "=sh_v")
                w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}")
                w("div", f"=doe*{prev('eq')}"); w("eq", f"={prev('eq')}+{cell('ni')}-{cell('div')}")
            w("core", f"=SUM({c}{R['c_jp']}:{c}{R['c_as']})")
            w("rev", f"={cell('core')}+{cell('hum')}")
            w("chk1", f"=ROUND(SUM({c}{R['c_jp']}:{c}{R['c_as']})-{cell('core')},6)")
            w("chk2", f"=ROUND({cell('core')}+{cell('hum')}-{cell('rev')},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            w("cogs", f"={cell('rev')}-{cell('gp')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
            w("opex", f"={cell('sga')}-{cell('oth')}"); w("op", f"={cell('gp')}-{cell('opex')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            w("pbt", f"={cell('op')}+{cell('fin')}")
            w("cont", f"={cell('pbt')}-{cell('tax')}"); w("ni", f"={cell('cont')}+{cell('disc')}-{cell('nci')}")
            w("nim", f"={cell('ni')}/{cell('rev')}"); w("eps", f"={cell('ni')}/{cell('sh')}*1000"); w("epsc", f"=({cell('cont')}-{cell('nci')})/{cell('sh')}*1000")
            w("nig", f"=IF({prev('ni')}>0,{cell('ni')}/{prev('ni')}-1,\"N/M（前一年為虧損）\")")
            w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
            w("fcff", f"={cell('op')}*(1-tax_w)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("fcfad", f"={cell('fcff')}-{cell('div')}")
        P.cell(b, 2, "FY2027 起＝前一年 ×(1＋核心年增)")
        P.cell(R["hum"], 2, "市場 × 份額 × 爬坡 × 匯率；FY2026 以前未揭露（視為 0）")
        P.cell(R["chk2"], 2, "FY2025：五地區取整值合計與合併營收差 0.003")
        P.cell(R["gp"], 2, "FY2027 起＝FY2026 毛利 ＋ 邊際毛利率 × Δ營收 ＋ 毛利率調整 × 營收")
        P.cell(R["sga"], 2, "FY2027 起＝FY2026 ×(1＋彈性 × 營收相對 FY2026 增幅)")
        P.cell(R["tax"], 2, "FY2026＝指引稅前 − 指引持續營業淨利（Derived）")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 ×(1 − tax_w)；股票報酬未取得（不加回）")
        P.cell(R["div"], 2, "FY2027 起＝DOE × 前一年底權益")
        for k in ("gr", "gm", "opm", "nim", "nig", "rp"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(LAB) + 2
    P.column_dimensions["A"].width = 52; P.column_dimensions["B"].width = 30
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "fcfad", "hum"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（¥B）"; D["B4"] = "=px*sh_v/1000"
    D["A5"] = "負債 D（社債及借款＋租賃代理）"; D["B5"] = "=debt_v+lease_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金與投資（2026-09-30）＝現金 − 借款 − 租賃代理 − 非支配 ＋ 其他非流動金融資產 × inv_k"; D["B10"] = "=cash_v-debt_v-lease_v-mi_v+inv_k*inv_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND(((px*sh_v/1000)*(rf+beta*(erp+crp))+(debt_v+lease_v)*kd*(1-tax_w))/((px*sh_v/1000)+debt_v+lease_v)+wacc_adj-wacc,9)"
    for rr_ in range(3, 11): D[f"B{rr_}"].number_format = "0.00%" if rr_ in (3, 6, 7, 8, 9) else "#,##0.0"
    DR = {}; rr = 13
    for s in SC:
        R = PR[s]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}" for y in range(2027, 2033)]); b = rr + 2
        rows = ["EBIT（營業利益）", "NOPAT", "＋折舊攤銷", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF（應為 0）", "折現年數（期中）", "折現因子", "FCFF 現值"]
        for i, l in enumerate(rows): D.cell(b + i, 1, l)
        for j, y in enumerate(range(2027, 2033)):
            c = "BCDEFG"[j]; pc = COL[y]
            D[f"{c}{b}"] = f"=PL!{pc}{R['op']}"; D[f"{c}{b+1}"] = f"={c}{b}*(1-tax_w)"
            D[f"{c}{b+2}"] = f"=PL!{pc}{R['da']}"; D[f"{c}{b+3}"] = f"=-PL!{pc}{R['cx']}"; D[f"{c}{b+4}"] = f"=-PL!{pc}{R['dwc']}"
            D[f"{c}{b+5}"] = f"=PL!{pc}{R['fcff']}"; D[f"{c}{b+6}"] = f"=ROUND(SUM({c}{b+1}:{c}{b+4})-{c}{b+5},9)"
            D[f"{c}{b+7}"] = j + 0.5; D[f"{c}{b+8}"] = f"=1/(1+wacc)^{c}{b+7}"; D[f"{c}{b+9}"] = f"={c}{b+5}*{c}{b+8}"
        t = b + 11
        items = [("明確期現值合計", f"=SUM(B{b+9}:G{b+9})"),
                 ("NOPAT FY2033＝NOPAT FY2032 ×(1＋g)", f"=G{b+1}*(1+g_tv)"),
                 ("終值（FY2032 年底）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^6"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金與投資（2026-09-30）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!F{R['fcff']}*stub_frac"),
                 ("− 尚待買回金額（無）", 0),
                 ("股權價值", f"=B{t+4}+B{t+5}+B{t+6}+B{t+7}"),
                 ("每股價值（評價日）", f"=MAX(0,B{t+8})/sh_v*1000"),
                 ("DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）", f"=B{t+9}*(1+ke)^t_tgt-dps_tgt"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+11}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=B{t+2}/G{b+1}"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）", f"=((B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6)+B{t+5}+B{t+6}+B{t+7})/sh_v*1000*(1+ke)^t_tgt-dps_tgt")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 10, ps=t + 9, ev=t + 4, tvs=t + 11, alt=t + 14, chk=b + 6)
        dn(f"dcf_{s}", f"DCF!$B${t+10}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 60

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價：DCF（權重 w_dcf）＋ Comparable（同業前瞻 P/E 中位數 × 情境 FY2028 EPS，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}")
        V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "DCF 頁")
    r_cu = vrow("Comparable（未調整）：同業前瞻 P/E 中位數 × 情境 FY2028 EPS（持續營業）", [f"=pe_med*PL!H{PR[s]['epsc']}" for s in SC],
                "目標時點的 NTM 以 FY2028 代（2027-11 起算 NTM 約 1/12 FY2027＋11/12 FY2028）；分母：模型各情境 FY2028 持續營業 EPS")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（日圓）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    r_cons = one("市場共識目標價（賣方 12 位平均；資料商 2026-10-19 頁面，第 3 季決算前）", "=cons_tp", "只作對照（Interested-party）")
    r_px = one("現價（In 頁 px）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（¥B；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=IF(B{r_up}>0.5,"已在最高級",TEXT(px*(1+IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5)))),"#,##0"))&"／"&TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,IF(B{r_up}>=-0.5,-0.5,-9))))),"#,##0")', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 持續營業 EPS", f"=B{r_tp}/PL!H{PR['B']['epsc']}", "分母：模型情境 B FY2028 持續營業 EPS")
    r_pe1 = one("現價 ÷ 情境 B FY2027 持續營業 EPS", f"=px/PL!G{PR['B']['epsc']}", "分母：模型情境 B FY2027")
    r_pe2 = one("現行前瞻 P/E（資料商 NTM，換算至 2026-11-20）÷ 同業中位數", "=pe_self/pe_med", "只作對照；相對重估需要的倍數＝1 ÷ 本值")
    r_hum = one("人形線性關節營收占 FY2030 合併營收（情境 B）", f"=PL!J{PR['B']['hum']}/PL!J{PR['B']['rev']}")
    for rr_ in (r_up, r_a1, r_a2, r_a3, r_hum): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "#,##0"
    V.column_dimensions["A"].width = 70; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, ca=r_ca, pd2=r_pd2, pe=r_pe, pe1=r_pe1, pe2=r_pe2, gj=r_gj, cons=r_cons, wk=r_wk, hum=r_hum)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司指引、市場共識對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或共識", "差", "說明"])
    rc = [("FY2026 營收", f"=PL!F{PR['B']['rev']}", "=g26_rev", "公司指引（Interested-party）"),
          ("FY2026 營業利益", f"=PL!F{PR['B']['op']}", "=g26_op", "公司指引"),
          ("FY2026 歸屬母公司淨利", f"=PL!F{PR['B']['ni']}", "=g26_ni", "公司指引；含停業單位 15.75"),
          ("FY2026 毛利", f"=PL!F{PR['B']['gp']}", "=g26_gp", "公司指引"),
          ("FY2027 持續營業 EPS（模型 B） vs 現價 ÷ 現行前瞻 P/E 隱含 NTM EPS", f"=PL!G{PR['B']['epsc']}", "=px/pe_self", "口徑：資料商 NTM（2026-10-19 頁面，第 3 季決算前；起算約 2.5/12 FY2026＋9.5/12 FY2027）vs 模型 FY2027；口徑不同，只看方向"),
          ("加權目標價 vs 賣方平均目標價", "=tp", "=cons_tp", "賣方 12 位平均（2026-10-19，第 3 季決算前，只作歷史對照）"),
          ("FY2026 營業利益 vs 公司 ROE 10% 所需營業利益", f"=PL!F{PR['B']['op']}", 40.0, "公司 ROE 10% 所需營業利益 400 億（公司 FY2026 第 2 四半期決算補足資料，2026-08-05，Interested-party）；FY2026 指引已超過")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_thk.py 批次格點＋線性內插；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 60; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_thk.py 以 LibreOffice 批次重算每一案例）；結果為常數快照"
    base_tp = sens["base"]["tp"] if sens else None
    SE["A2"] = "快照基準值（加權，日圓）"; SE["B2"] = base_tp
    SE["C2"] = "本檔即時值（加權，日圓）"; SE["D2"] = "=tp"
    SE["E2"] = f"快照輸入指紋 {sens['fp'] if sens else ''}"
    SE["A3"] = "build_id"; SE["D3"] = bid
    SE["A4"] = "模型狀態"; SE["D4"] = '=IF(ABS(D2-B2)>0.5,"STALE","CURRENT")'
    hdr(SE, 6, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    sr = 7
    if sens:
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]); SE.cell(sr, 5, cs["tp"] / base_tp - 1)
            SE.cell(sr, 6, cs["dcf"]); SE.cell(sr, 7, cs["dcf"] / sens["base"]["dcf"] - 1); SE.cell(sr, 8, cs["rt"])
            SE[f"E{sr}"].number_format = "0.0%"; SE[f"G{sr}"].number_format = "0.0%"; SE[f"D{sr}"].number_format = "#,##0"; SE[f"F{sr}"].number_format = "#,##0"
            sr += 1
    SE.cell(sr + 1, 1, "說明：相對倍數係數 rel_k 基準即為 1.00（未調整），故「相對倍數係數＝1」案例＝基準。doe、eq_26 等無區間的鍵不列案例；人形市場 m26／m30／m32 的 B 列不另設區間（線 6 低／高已作情境 C／A）。")
    SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 34

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = "對帳 — 本模型使用的線層級變數（讀線模型成品「下游取數」頁；與線模型不一致的自設值標明理由，不靜默改口徑）"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型一致？"])
    rec = [("行星滾柱絲杠市場 2026（低／基準／高）", "6", "89.6／298.6／731.5", "百萬美元", "五線歸因；基準單價口徑＝覓途", L6 + " 下游取數!B12:D12", "Derived", "In m26_C／m26_B／m26_A → PL 人形子列", "一致（照取）"),
           ("行星滾柱絲杠市場 2030（低／基準／高）", "6", "233.8／1,558.6／6,546.0", "百萬美元", "同上", L6 + " 下游取數!B13:D13", "Derived", "In m30_*", "一致"),
           ("行星滾柱絲杠市場 2032（低／基準／高）", "6", "337.8／3,378.2／13,794.2", "百萬美元", "同上", L6 + " 下游取數!B14:D14", "Derived", "In m32_*", "一致"),
           ("行星滾柱絲杠 2030 單價口徑（中國量產／覓途／MS）", "6", "400.0／1,558.6／4,176.1", "百萬美元", "基準需求 × 三種單價口徑", L6 + " 下游取數!B15:D15", "Derived", "未進公式（以 shr 區間涵蓋）", "未使用"),
           ("滾珠絲杠需求量 2030（主要來自 EMB）", "6", "1,948.8／4,964.2／9,526.0", "萬根", "只計增量（EMB、SBW、機器人）", L6 + " 下游取數!B23:D23", "Derived", "未進公式", "不一致（自設 0）：THK 已於 2026-06-01 轉讓汽車與運輸機器事業，EMB／SBW 增量不計入；機器人部分以核心成長率涵蓋"),
           ("滾珠絲杠與線性導軌份額", "6", "未取得", "%", "THK、上銀、NSK", L6 + " 份額!F12", "—", "未進公式", "一致（線模型亦未取得）"),
           ("行星滾柱絲杠領先者", "6", "Schaeffler、Rollvis、GSA、SKF；THK 未列", "—", "定性", L6 + " 份額!B11", "Interested-party", "shr_* 說明（基準 2% 偏低）", "一致"),
           ("人形用滾柱絲杠市場 2030（低／基準／高）", "4", "205.2／747.0／1,563.7", "百萬美元", "單價隨學習曲線；絲杠情形 3", L4 + " 下游取數!C72:E72", "Derived", "未進公式（對照）", "線 4 與線 6 不一致：線 6 基準 1,558.6 為線 4 的 2.1 倍（單價口徑不同）；本模型照所屬線 6，差異以 shr 區間吸收"),
           ("人形全球出貨 2030（低／基準／高）", "4", "20 萬／50 萬／89 萬", "台", "全球、含小型機", L4 + " 下游取數!C38:E38", "Derived", "未進公式（線 6 市場已含）", "一致（經線 6）"),
           ("全尺寸人形每台滾柱絲杠", "4", "12／14／14", "支／台", "全尺寸", L4 + " 下游取數!C26:E26", "Derived", "說明用", "一致"),
           ("全球工業機器人裝機 2026→2030（基準）", "3", "65.5 萬 → 85.4 萬（年複合 6.9%）", "台", "IFR 新裝機口徑", L3 + " 下游取數!E13、E17", "Derived", "gm_B（4%）的上緣類比；未進公式", "自設低於線 3：THK 營收主體為工具機、半導體與電子設備，機器人只占一部分（用途別營收金額未揭露）"),
           ("全球工業機器人裝機 2030（低／高）", "3", "69.3 萬／103.5 萬", "台", "IFR 新裝機口徑", L3 + " 下游取數!E9、E24", "Derived", "gm_A／gm_C 說明", "同上"),
           ("JARA 最新季受注 YoY（2026 年 7–9 月）", "3", "+46.6%", "%", "日系整機訂單", L3 + " 下游取數!E339", "Verified", "g27_A 依據之一（與 THK 日本單體受注 +133% 同向）", "一致"),
           ("中國工業機器人產量 2026 年 1–9 月 YoY", "3", "+29.1%", "%", "國家統計局", L3 + " 下游取數!E348", "Verified", "說明用（中國營收指引 +46.8%）", "一致")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 30; RE.column_dimensions["F"].width = 60; RE.column_dimensions["I"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("半導體與 AI 設備投資循環反轉（受注含提前、重複下單）", "是（g27_C、g28_C）", "情境 C", "日本單體受注與受注殘、電子產業受注、SEMI 設備出貨"),
             ("中國線性導軌與絲杠廠價格競爭", "是（gmadj_C）", "情境 C 毛利率 −2pt", "中國營收、毛利率、陸廠擴產"),
             ("人形線性關節份額為零或採用行星滾柱絲杠以外的方案", "是（shr_C、Sens）", "人形子列占營收比例小（見 Val 頁）", "Tesla、Figure 等供應鏈揭露；公司人形產品發表"),
             ("事業轉讓的剩餘或有負債（AP87）", "否", "金額未揭露；轉讓利益已入帳", "後續決算短信注記"),
             ("匯率（美元、歐元、人民幣）", "否", "公司 FY2026 假設 157.58 日圓；未逐年建模", "USDJPY、CNYJPY"),
             ("資本政策：淨現金降到中性、容許淨負債 1–2 倍 EBITDA（自己株式取得）", "否", "以公允價格買回對每股價值中性；未建模", "自己株式取得公告"),
             ("稅率回到 30% 以上", "部分（tax_w 高案 32%）", "9M'26 17.2% 原因未查核", "季度有效稅率")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 50

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([TS3Q, FB3Q, TS25, PR2Q, QTR, SA, YCL, JGB10, L6 + "（下游取數、份額）", L3 + "（下游取數）", L4 + "（下游取數）",
                            HWX + "（Comps 頁同業倍數）", HDX + "（Comp 頁 Nabtesco 換算 P/E）", NBX + "（In 頁 SMC P/E、rf、CRP）"], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"THK（6481.T）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-20（第 3 季決算 2026-11-13）；成品日 2026-11-23；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"#,##0")&" 日圓）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：THK 是線性導軌（LM Guide）與滾珠絲杠的全球龍頭，轉讓汽車事業後成為純產業機器零件公司；股價由產業機器循環（工具機、半導體與電子設備、機器人）帶動的營收與營業槓桿決定，人形機器人線性關節是小型選擇權"
    RD["A4"] = (f'="在 FY2027 核心營收情境 B 年增 "&TEXT(g27_B,"0%")&"、邊際毛利率 "&TEXT(mm_gp,"0%")&"、同業前瞻 P/E 中位數 "&TEXT(pe_med,"0.0")&" 倍的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"#,##0")&" 日圓（DCF "&TEXT(Val!B{r_dcft},"#,##0")&"、Comparable "&TEXT(Val!B{r_cmpt},"#,##0")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0.0")&" 十億日圓）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 日圓。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：日本單體受注連兩季低於營收（受注 ÷ 營收 < 1）；電子產業受注 YoY 轉負；季度毛利率跌破 30%；FY2027 公司指引營收低於 FY2026 實績"
    RD["A7"] = "範圍：只建持續營業（產業機器）事業；不讀 Siliconomics／Tokenomics／Wattonomics（算力線）；自用 AI 讀法等規格條文不適用。HTML 展示層見 dist/ 同名 .html"
    RD["A8"] = f'="命題回覆：問：產業機器循環能否支持高於現價的目標價？答：情境 B FY2028 持續營業 EPS "&TEXT(PL!H{PR["B"]["epsc"]},"#,##0")&" 日圓，現價只有其 "&TEXT(px/PL!H{PR["B"]["epsc"]},"0.0")&" 倍，低於同業中位數 "&TEXT(pe_med,"0.0")&" 倍；加權目標價 "&TEXT(Val!B{r_tp},"#,##0")&" 日圓。問：人形機器人有多重要？答：情境 B FY2030 人形子列只占合併營收 "&TEXT(Val!B{r_hum},"0.0%")&"，結論不靠人形（模型輸出，非獨立證據）。"'
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
def lo_profile():
    d = os.path.join(os.environ.get("HOME", tempfile.gettempdir()), "lo_thk")
    os.makedirs(d, exist_ok=True)
    return "file://" + d

def recalc_many(srcs, outdir):
    os.makedirs(outdir, exist_ok=True)
    for i in range(0, len(srcs), 40):
        chunk = srcs[i:i + 40]
        for attempt in range(2):
            try:
                subprocess.run(["soffice", f"-env:UserInstallation={lo_profile()}", "--headless", "--convert-to", "xlsx", "--outdir", outdir] + chunk,
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
                break
            except subprocess.TimeoutExpired:
                subprocess.run(["pkill", "-f", "soffice"]); continue
    return [os.path.join(outdir, os.path.basename(s)) for s in srcs]

def read_vals(path, VR, DR=None):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]
    d = dict(tp=V[f"B{VR['tp']}"].value, dcf=V[f"B{VR['dcf']}"].value, cmp=V[f"B{VR['cmp']}"].value, up=V[f"B{VR['up']}"].value, rt=V[f"B{VR['rt']}"].value)
    if DR:
        I = wb["In"]; w = {}; wd = None
        for row in I.iter_rows(values_only=True):
            if row[0] in ("w_A", "w_B", "w_C"): w[row[0][-1]] = row[4]
            if row[0] == "w_dcf": wd = row[4]
        alt = sum(w[s] * wb["DCF"][f"B{DR[s]['alt']}"].value for s in SC)
        d["alt_dcf"] = alt; d["alt_tp"] = wd * alt + (1 - wd) * d["cmp"]
    return d

def run_cases(ovs):
    """ovs：[(tag, ov)] → {tag: vals}；一次 LibreOffice 批次重算。"""
    d = tempfile.mkdtemp(prefix="thk_"); paths = []; VR = DR = None
    for tag, ov in ovs:
        p = os.path.join(d, f"c_{tag}.xlsx"); VR, _, DR = build(p, ov); paths.append(p)
    outs = recalc_many(paths, os.path.join(d, "o"))
    res = {tag: read_vals(o, VR, DR if tag == "base" else None) for (tag, _), o in zip(ovs, outs)}
    shutil.rmtree(d, ignore_errors=True); return res

def in_val(k, col):
    for row in ROWS:
        if row[0] == k:
            return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)

def fingerprint():
    return hashlib.sha1(json.dumps([r for r in ROWS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]

def rating(u):
    return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"

CASES = [  # (鍵, 驅動白話)
    ("g_anc", "FY2026 實績相對公司指引"), ("g27_B", "核心營收年增 FY2027（情境 B）"), ("g28_B", "核心營收年增 FY2028（情境 B）"), ("gm_B", "核心營收年增 FY2029–30（情境 B）"),
    ("glt_B", "核心營收年增 FY2031–32（情境 B）"), ("shr_B", "人形線性關節份額（情境 B）"), ("rp27", "人形營收爬坡 FY2027"), ("fx", "美元兌日圓"),
    ("mm_gp", "邊際毛利率"), ("gmadj_B", "毛利率調整（情境 B）"), ("sga_el", "銷管費用彈性"), ("oth_op", "其他營業損益淨額"), ("nonop", "業外金融損益"), ("tax_w", "有效稅率"),
    ("fc_da", "折舊 ÷ 營收"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("lease_v", "租賃負債代理"), ("inv_k", "其他金融資產計入比例"),
    ("beta", "β"), ("erp", "股權風險溢酬"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重"), ("pe_med", "同業前瞻 P/E 中位數")]
BETA_B = statistics.median([0.67 * 0.76 + 0.33, 0.67 * statistics.median([0.53, 1.04, 0.99, 1.08]) + 0.33])
FXD = {"fc_da": 18.2 / 320, "fc_cx": 21.0 / 320, "pe_med": statistics.median([15.35, 11.81, 58.97, 28.35, 21.1499, 15.6739, 13.7701, 21.35])}


def solve_batch(key, lo, hi, targets, px0):
    """格點批次重算（13 點）→ 對每個目標上檔找相鄰兩點線性內插；再以 7 點細格點精修一次。"""
    import numpy as np
    xs = list(np.linspace(lo, hi, 13))
    res = run_cases([(f"g{i}", {key: x}) for i, x in enumerate(xs)])
    ys = [res[f"g{i}"]["up"] for i in range(len(xs))]
    out = {}
    for tgt in targets:
        if tgt is None: continue
        br = next((i for i in range(len(xs) - 1) if (ys[i] - tgt) * (ys[i + 1] - tgt) <= 0), None)
        if br is None: out[tgt] = None; continue
        a, b_ = xs[br], xs[br + 1]
        xs2 = list(np.linspace(a, b_, 7)); r2 = run_cases([(f"f{i}", {key: x}) for i, x in enumerate(xs2)]); y2 = [r2[f"f{i}"]["up"] for i in range(7)]
        j = next(i for i in range(6) if (y2[i] - tgt) * (y2[i + 1] - tgt) <= 0)
        x0, x1, y0, y1 = xs2[j], xs2[j + 1], y2[j], y2[j + 1]
        out[tgt] = x0 if y1 == y0 else x0 + (tgt - y0) * (x1 - x0) / (y1 - y0)
    return out


def main():
    px0 = in_val("px", "base")
    jobs = [("base", {})]; meta = []
    for k, drv in CASES:
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if k == "beta": v = BETA_B + (-0.3 if side == "lo" else 0.3)
            if v is None: continue
            b_ = in_val(k, "base")
            if not isinstance(b_, (int, float)): b_ = FXD.get(k, BETA_B if k == "beta" else None)
            if b_ is not None and abs(v - b_) < 1e-12: continue
            tag = f"{k}_{side}"; jobs.append((tag, {k: v})); meta.append((tag, drv + ("（低）" if side == "lo" else "（高）"), f"{k}＝{v:.4g}", k))
    for s in SC:
        tag = f"only_{s}"; jobs.append((tag, {f"w_{x}": (1.0 if x == s else 0.0) for x in SC})); meta.append((tag, f"只看情境 {s}", f"w_{s}＝1，其餘 0", f"w_{s}"))
    res = run_cases(jobs)
    base = res["base"]; print("基準", base)
    cases = [dict(drv=d, set=st, key=k, **res[t]) for t, d, st, k in meta]
    up_ = base["alt_tp"] / px0 - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rating(up_)))
    for c in cases: print(c["key"], c["drv"], round(c["tp"]), c["rt"])
    up0 = base["up"]; tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    revs = []
    for key, b0, lo_, hi_, note in (("g27_B", 0.12, -0.40, 0.60, "情境 B 的 FY2027 核心營收年增（其他不變，整本重算；情境 A／C 不動；之後年度在此基礎上成長）"),
                                    ("mm_gp", 0.40, -0.50, 1.00, "邊際毛利率（三情境同動）"),
                                    ("pe_med", FXD["pe_med"], 0.0, 40.0, "同業前瞻 P/E 中位數（Comparable；DCF 不受影響）")):
        sol = solve_batch(key, lo_, hi_, [0.0, upper, lower], px0)
        for lab, tgt in (("目標價＝現價（上檔 0%）", 0.0), (f"升一級邊界（上檔 {upper:+.0%}）" if upper is not None else None, upper), (f"降一級邊界（上檔 {lower:+.0%}）" if lower is not None else None, lower)):
            if lab is None: continue
            x = sol.get(tgt)
            revs.append(dict(label=lab, key=key, val=x if x is not None else f"區間 {lo_:+.2f}～{hi_:+.2f} 內無解", base=b0, note=note))
            print("反解", key, lab, x)
    sens = dict(base=base, cases=cases, fp=fingerprint())
    json.dump(sens, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="thkf_"); raw = os.path.join(tmp, NAME)
    bid = f"{VER}-{fingerprint()}"
    VR, PR, DR = build(raw, {}, sens, revs, bid)
    out = recalc_many([raw], os.path.join(tmp, "o"))[0]
    shutil.copy(out, os.path.join(dist, NAME))
    wbv = openpyxl.load_workbook(os.path.join(dist, NAME), data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10]); assert not errs, errs
    print("Sens 狀態", wbv["Sens"]["D4"].value, "tp", wbv["Val"][f"B{VR['tp']}"].value)
    write_readme(os.path.join(dist, NAME), VR, PR, DR, bid, sens, revs)
    json.dump(dict(company="thk", ticker="6481.T", model_sheet="Q",
                   official_segments=[dict(name=n, label=l, period="9M'26（累計，外部顧客）", value=v, unit="¥B", source=TS3Q + " 分部情報")
                                      for n, l, v in (("日本", "日本（所在地別）營收", 95.602), ("美洲", "美洲（所在地別）營收", 55.572), ("歐洲", "歐洲（所在地別）營收", 40.820),
                                                      ("中國", "中國（所在地別）營收", 78.712), ("其他", "其他（所在地別）營收", 20.616))],
                   sublines=[], check_labels=["檢查：分部合計 − 停業重分類 − 合併營收（四捨五入差 < 0.005）"],
                   note="官方報告分部為所在地別（含已轉讓的汽車與運輸機器事業至 2026-05），Q 頁以「分部合計 − 停業單位重分類 − 合併營收」對帳。PL 頁營收列改用公司揭露的持續營業客戶地區別營收（Investor Information p.14），另有「檢查：地區合計 − 核心營收」「檢查：核心＋人形 − 合併營收」。"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)


def write_readme(path, VR, PR, DR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f0 = lambda x: f"{x:,.0f}"; pc = lambda x: f"{x*100:+.1f}%"
    inv = lambda k: I['E' + str([r for r in range(1, I.max_row + 1) if I['A' + str(r)].value == k][0])].value
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:8]
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {r['val'] if isinstance(r['val'], str) else format(r['val'], '.4f')} | {r['base']:.4g} | {r['note']} |" for r in revs)
    def yr(s, k, fmt="{:,.1f}", ys=range(2024, 2033)):
        out = []
        for y in ys:
            v = P[f"{COL[y]}{PR[s][k]}"].value
            out.append(fmt.format(v) if isinstance(v, (int, float)) else "—")
        return " | ".join(out)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f0(g(chr(66+i)+str(VR['dcfr'])))} | {f0(g(chr(66+i)+str(VR['cu'])))} | {f0(g(chr(66+i)+str(VR['tps'])))} | {g(chr(66+i)+str(VR['gate'])):,.1f} |" for i, s in enumerate(SC))
    rts = sorted(set(c['rt'] for c in sens['cases']))
    YS = range(2024, 2033)
    txt = f"""# physical_ai/thk — THK（6481.T）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

東證 Prime 上市。本 README 的數字由 `build_thk.py` 從重算後的成品讀出產生，不手寫。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A8'].value}

**讀法：** 上檔主要來自兩件事：(1) 現價只有模型 FY2027 持續營業 EPS 的 {g(f"B{VR['pe1']}"):.1f} 倍，FY2026 指引已三度上修（營收 2,600 → 3,200 億日圓、營業利益 260 → 520 億日圓）且日本單體受注 ÷ 營收 1.76；(2) Comparable 用同業 8 家前瞻 P/E 中位數 {inv('pe_med'):.1f} 倍（其中 5 家報價日在 11 月同業回落之前、未換算，可能偏高）。結論主要建立在 Assumed 的 FY2027–28 成長率與邊際毛利率，以及部分逾期的同業倍數上；人形子列（情境 B FY2030 占營收 {g(f"B{VR['hum']}")*100:.1f}%）不是承重項。

## 估值（評價日 2026-12-31；目標時點約 12 個月）
| 情境 | DCF（推到 12 個月後） | Comparable（P/E × FY2028 EPS） | 情境目標價 | 收支關卡（FY2026–30 FCF 扣股利，¥B） |
|---|---|---|---|---|
{scn}

- 加權目標價 **{f0(tp)} 日圓**（DCF {f0(dcf)}、Comparable {f0(cmp_)}），現價 {f0(inv('px'))}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")} → 投資建議 **{g(f"B{VR['rec']}")}**。市場共識目標價 {f0(inv('cons_tp'))}（12 位，2026-10-19，第 3 季決算前，只作對照）。
- 方法：DCF 標準版 FCFF（{g(f"B{VR['wk']}")}，為 WACC／ke／g／RONIC）50%＋Comparable 50%（同業 8 家前瞻 P/E 中位數 {inv('pe_med'):.2f} 倍 × 各情境 FY2028 持續營業 EPS）。相對倍數係數 1.00（未調整：缺 THK 對同業的歷史相對倍數）。
- 隱含前瞻 P/E（加權目標價 ÷ 情境 B FY2028 持續營業 EPS）{g(f"B{VR['pe']}"):.1f} 倍；THK 現行前瞻 P/E {inv('pe_self'):.2f} 倍 ＝ 同業中位數的 {g(f"B{VR['pe2']}"):.2f} 倍。

## 營收與標準損益（情境 B，¥B；FY2024–25 持續營業實績、FY2026 公司指引、FY2027 起模型）
| 列 | {" | ".join(f"FY{y}" for y in YS)} |
|---|{"---|" * len(YS)}
| 日本 | {yr('B','c_jp')} |
| 美洲 | {yr('B','c_am')} |
| 歐洲 | {yr('B','c_eu')} |
| 中國 | {yr('B','c_cn')} |
| 亞洲及其他 | {yr('B','c_as')} |
| 人形線性關節（命題子列） | {yr('B','hum')} |
| 營收（合併） | {yr('B','rev')} |
| 毛利 | {yr('B','gp')} |
| 營業利益 | {yr('B','op')} |
| 歸屬母公司淨利 | {yr('B','ni')} |
| EPS（持續營業，日圓） | {yr('B','epsc','{:,.0f}')} |

- FY2024、FY2025 地區別：FY2024 持續營業地區別未取得（只有 FY2025 與 FY2026 指引），故 FY2024 地區列空白。FY2023 為舊口徑（含汽車事業），見 xlsx PL 頁。
- FY2026 歸屬母公司淨利含停業單位（事業轉讓利益 16.4）；持續營業 EPS 才是評價分母。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 核心產業機器營收年增（A／B／C）：FY2027 22%／12%／0%；FY2028 8%／4%／−10%；FY2029–30 6%／4%／3%；FY2031–32 4%／3%／2%。五地區同一成長率（地區差異無證據）。
- 人形：線 6 行星滾柱絲杠市場（A＝高、B＝基準、C＝低需求情境）× THK 份額 4%／2%／0.5% × 爬坡（FY2027 25% → FY2030 100%）× 157.58 日圓。
- 毛利：FY2026 指引毛利 ＋ 邊際毛利率 40% × Δ營收 ＋ 情境調整（A +1pt、C −2pt）；銷管費用彈性 0.3；稅率 28%（9M'26 實績 17.2%，原因未查核）。
- 現金流：折舊 5.7%、資本支出 6.6%（FY2026 計畫比率）、營運資金 30% × 營收增量；股利＝DOE 8% × 前一年底權益。
- 折現：rf 2.817%（與 Nabtesco 同值，逾期）、β {inv('beta'):.3f}、ERP 5%、CRP 0.91%、g＝MIN(3%, rf)、RONIC＝WACC。評價日股數 112.07 百萬股；淨負債約 {-wb['DCF']['B10'].value:,.1f}（含租賃代理、扣其他非流動金融資產）。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| FY2027–28 核心成長（g27、g28） | 日本單體受注 ÷ 營收（2026Q3 1.76）、受注殘（570 億）、電子產業受注 YoY；FY2027 公司指引（2027-02） |
| 邊際毛利率（mm_gp） | 季度毛利率（9M'26 34.6%）、Δ毛利 ÷ Δ營收 |
| 同業 P/E（pe_med） | 同業前瞻 P/E 換算到同一估值日後的中位數 |
| 人形份額（shr） | 公司人形相關產品與客戶揭露、Tesla／Figure 供應鏈 |

失效條件：{RD['A6'].value.split('：', 1)[1]}。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f0(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁；各案例估值評等出現「{'、'.join(rts)}」。

## 反解（整本重算；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
- 線 6：行星滾柱絲杠市場 2026／2030／2032 低／基準／高（下游取數!B12:D14）→ m26／m30／m32 的情境 C／B／A，照取；滾珠絲杠需求 2030（下游取數!C23，主要來自線控煞車）**不計入**：THK 已轉讓汽車事業（與線模型不一致，自設 0）；份額（份額!F12）線模型亦未取得。
- 線 4：人形用滾柱絲杠市場 2030 基準 747（下游取數 C72）與線 6 的 1,558.6 不一致（單價口徑），本模型照所屬線 6，差異由份額區間吸收。
- 線 3：工業機器人裝機 2026→2030 年複合 6.9%（下游取數 E13、E17）只作 FY2029–30 成長的上緣類比；THK 營收主體為工具機、半導體與電子設備，自設 4% 低於線 3。

## 限制與未取得資料
- **用途別（工具機、半導體、電子、機器人）營收金額未取得**（公司只揭露比例圓餅與日本單體受注產業別）；人形機器人營收、客戶未找到；FY2024 持續營業地區別營收未取得；FY2023 持續營業重編值未取得（舊口徑）。
- 官方報告分部為所在地別且含已轉讓事業至 2026-05，無法向前推估；PL 營收列改用客戶地區別（口徑不同，見 Q 頁）。
- 9M'26 有效稅率 17.2% 的原因、租賃負債 2026-09-30 值、其他非流動金融資產內容未查核；2026-11-13 決算後股價 −14.4% 的原因未查核。
- 同業前瞻 P/E：NSK、IKO、直得、上銀、SMC 報價日早於 2026-11-09（未換算，同業 11 月回落 6–22%，中位數可能偏高）；少於約 8 家不做迴歸。共識 EPS 逐年值未取得；共識目標價為第 3 季決算前。
- rf、CRP 沿用 2026-11-02／01-05 取值（逾期，跨模型一致）。DCF 為本模型自寫，公式與 `docs/plan/個股DCF標準規格.md` 相同。

## 檔案
- `build_thk.py`：建置＋整本重算 Sens／反解（LibreOffice 批次）＋產生本 README、`segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-23）：初版；現價 {f0(inv('px'))}（2026-11-20 收盤）；加權目標價 {f0(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


if __name__ == "__main__":
    main()

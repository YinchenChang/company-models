# NSK 日本精工（6471.T）個股模型 v0.1 — AI實體應用 Project（線 6 共用零組件：軸承、滾珠絲杠、線性導軌；人形用量另讀線 4）
# 用法：python3 physical_ai/nsk/build_nsk.py   → 建置、整本重算 Sens 與反解（LibreOffice 批次）、輸出 dist/ 成品、README.md、sens.json、segments.json
# 會計年度：3 月決算，IFRS；本檔沿用公司稱呼：FY2026＝2027 年 3 月期（公司指引年度），FY2025＝2026 年 3 月期（實績），FY2024＝2025 年 3 月期，FY2023＝2024 年 3 月期。
# 金額：十億日圓（¥B）；股數：百萬股；EPS、股價：日圓。
# 轉向（Steering）事業自 2025-09-01 起納入合併（FY2025 起報告分部）；FY2023–24 歷史數字為公司重編後口徑（資料商轉述），分部歷史未取得。
# 結構：README、In、Hist、Q（FY2026 公司指引分部與中間期實績）、PL（三情境：命題模組→分部營收→標準損益表→現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 不讀 Siliconomics／Tokenomics／Wattonomics（算力線）。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md（評價日 2026-09-30＝最近資產負債表日）。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import re
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261124"
NAME = f"{DATE}_NSK個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032

# ---------------- 出處 ----------------
TS2Q = "公司 2027年3月期 第2四半期（中間期）決算短信〔IFRS〕(連結)（2026-11；PDF 文字層無法讀取，數字取自下列資料商摘要）https://www.nsk.com/jp-ja/company/investors/ir-news/2026/financial-results-for-q2-fy2026--consolidated-/"
Q2 = "Quartr NSK Q2 2027 earnings summary（資料商摘要，2026-11）https://quartr.com/events/nsk-ltd-6471-q2-2027_FYqHy2GP"
Q1 = "Quartr NSK Q1 2027 earnings summary（資料商摘要，2026-07-31）https://quartr.com/events/nsk-ltd-6471-q1-2027_FYqHyiYN"
Q4 = "Quartr NSK Q4 2025 earnings summary（資料商摘要；FY2024＝2025 年 3 月期）https://quartr.com/events/nsk-ltd-6471-q4-2025_FYqHUzll"
SAI = "stockanalysis.com NSK income statement（資料商標準化口徑；頁面 2026-11-24）https://stockanalysis.com/quote/tyo/6471/financials/income-statement/"
SAB = "stockanalysis.com NSK balance sheet（季，2026-09-30；頁面 2026-11-24）https://stockanalysis.com/quote/tyo/6471/financials/balance-sheet/?p=quarterly"
SAS = "stockanalysis.com NSK statistics（資料商；頁面 2026-11-24）https://stockanalysis.com/quote/tyo/6471/statistics/"
YCL = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
JGB10 = "日本財務省 國債金利情報 10 年：2026-11-02 2.817%（與本專案 Nabtesco、哈默納科、Fanuc、SMC、THK 模型同一取值，跨模型一致；逾期）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整（立場：中立彙整，但來源為賣方或公司）"
SELL = "賣方分析師（立場：券商研究與交易關係）"
L6 = "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx"
L3 = "physical_ai/lines/industrial_robot/dist/20261116_工業機器人單位經濟_v0.1.xlsx"
L4 = "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx"
SMX = "physical_ai/smc/dist/20261123_SMC個股模型_v0.1.xlsx"
THX = "physical_ai/thk/dist/20261123_THK個股模型_v0.1.xlsx"

# ---------------- In 列 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源1, 來源2／說明, 來源誘因)
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

REG = [("im", "產業機械"), ("au", "汽車"), ("st", "轉向與其他（差額）")]
sec("A. FY2026 錨點：公司指引（2026-11 第 2 季決算，第二次上修；Interested-party）")
inp("g26_im", "公司指引：FY2026 產業機械營收", "¥B", None, 436.0, None, "Interested-party", Q2 + "：產業機械 FY 營收 436", "H2 221.5（隱含 H1 214.5，Derived）；營業利益 37.5（調整後 44）", CO)
inp("g26_au", "公司指引：FY2026 汽車營收", "¥B", None, 403.0, None, "Interested-party", Q2 + "：汽車 FY 營收 403", "營業利益 17.5；調整後營業利益率目標 5%；全球車產假設 9,100 萬台", CO)
inp("g26_rev", "公司指引：FY2026 合併營收", "¥B", None, 1050.0, None, "Interested-party", Q2 + "：1,050（較 2026-07 上修 +10；2026-05 初始 1,000）", "H1 實績 520.351；H2 隱含 529.6（Derived）", CO)
inp("g26_st", "FY2026 轉向事業與其他、內部調整（＝合併 − 產業機械 − 汽車；差額）", "¥B", None, "=g26_rev-g26_im-g26_au", None, "Derived", "公司未在本次取得的資料中給轉向分部營收；差額含轉向與內部交易調整", "轉向自 2025-09-01 納入合併（" + Q2 + "）；分部拆分未取得，差額不可當作轉向營收")
inp("g26_op", "公司指引：FY2026 營業利益", "¥B", None, 56.0, None, "Interested-party", Q2 + "：56（較 2026-07 上修 +6；2026-05 初始 42）", "調整後（扣除特定項目）71；H1 實績 30.066 → H2 隱含 25.9（Derived）", CO)
inp("g26_pbt", "公司指引：FY2026 稅前利益", "¥B", None, 53.5, None, "Interested-party", Q2 + "：53.5", "H1 實績 28.909", CO)
inp("g26_ni", "公司指引：FY2026 歸屬母公司淨利", "¥B", None, 34.0, None, "Interested-party", Q2 + "：34", "H1 實績 18.838（+78.8%）", CO)
inp("g26_nci", "FY2026 非支配股東損益", "¥B", 0.5, 1.5, 2.5, "Assumed", "未取得公司指引；FY2025 實績 1.440（" + SAI + "）", "")
inp("gm26", "FY2026 毛利率（毛利 ÷ 營收）", "%", 0.215, 0.225, 0.235, "Assumed", "未取得公司指引與 H1 毛利；FY2025 21.0%（191.291 ÷ 911.644）、TTM 22.05%（" + SAI + "；" + SAS + "）", "H1 營業利益率 5.8% 高於 FY2025 4.3%，且結構改革，故取 22.5%")
inp("oth26", "FY2026 其他營業損益淨額（結構改革等特定項目）", "¥B", -18.0, -15.0, -10.0, "Derived", "＝調整後營業利益 71 − 營業利益 56（" + Q2 + "）", "未取得公司拆分；FY2025 特定項目 6.4（45.2 − 38.8，" + "資料商 Dow Jones/Quartr 摘要" + "）", CO)
inp("dps_26", "FY2026 每股股利（年間）", "¥", None, 34.0, None, "Interested-party", Q1 + "：34 維持；DOE 下限 2.5%", "FY2025 實績年間 34（中間 17＋期末 17）", CO)
inp("g_anc", "FY2026 實績相對公司指引（營收同比例；營業利益同比例）", "%", -0.02, 0.0, 0.02, "Assumed", "基準＝公司指引（H1 已實績，偏離只來自 H2）", "H1 實績占指引 49.6%；兩次上修；低／高 ±2%")

sec("C. 命題模組一：產業機械（半導體設備、工具機、機器人與一般機械軸承、絲杠、導軌）營收成長；汽車與轉向（車產）營收成長")
GTXT = {
    "i27": ("FY2027", "產業機械", {"A": (0.12, "樂觀：H2 指引 221.5 年化 443（+1.6% 於 FY2026），半導體設備與工具機持續強勁（Q1 描述「特別是半導體與工具機」，" + Q1 + "）"),
                               "B": (0.05, "基準：H2 年化 443 ＋ 溫和成長；公司產業機械營業利益率目標 H2 ≥ 11%（" + Q2 + "）"),
                               "C": (-0.05, "保守：半導體與工具機循環回落（Analogy：上一輪 FY2024 營收持平、營業利益率 2–3%）")}),
    "i28": ("FY2028", "產業機械", {"A": (0.08, ""), "B": (0.04, "基準：循環趨緩"), "C": (-0.08, "保守：循環下行")}),
    "im": ("FY2029–FY2030", "產業機械", {"A": (0.06, ""), "B": (0.04, "基準：工業機器人裝機 2026→2030 年複合 6.9%（線 3 下游取數 E13→E17）為上緣類比"), "C": (0.02, "")}),
    "ilt": ("FY2031–FY2032", "產業機械", {"A": (0.04, ""), "B": (0.03, "長期：名目產業機械需求"), "C": (0.02, "")}),
    "x27": ("FY2027", "汽車與轉向", {"A": (0.04, "樂觀：車產小幅成長、價格轉嫁關稅"), "B": (0.01, "基準：公司假設全球車產 9,100 萬台（持平，" + Q1 + "、" + Q2 + "）；歐洲與中國需求疲弱（" + Q2 + "）"),
                                "C": (-0.05, "保守：歐中車產下滑、日圓轉強")}),
    "x28": ("FY2028", "汽車與轉向", {"A": (0.03, ""), "B": (0.015, "基準：車產持平略增"), "C": (-0.04, "保守")}),
    "xm": ("FY2029–FY2030", "汽車與轉向", {"A": (0.03, ""), "B": (0.02, "基準：名目車產與電動化零件含量"), "C": (0.01, "")}),
    "xlt": ("FY2031–FY2032", "汽車與轉向", {"A": (0.03, ""), "B": (0.02, ""), "C": (0.01, "")}),
}
GRNG = {"i27": (0.0, 0.10), "i28": (-0.02, 0.07), "im": (0.02, 0.06), "ilt": (0.02, 0.04), "x27": (-0.03, 0.04), "x28": (-0.02, 0.04), "xm": (0.0, 0.03), "xlt": (0.0, 0.03)}
for f, (yl, who, d) in GTXT.items():
    for s in SC:
        v, note = d[s]
        lo, hi = GRNG[f] if s == "B" else (None, None)
        inp(f"{f}_{s}", f"情境 {s}：{who}營收年增 {yl}", "%/年", lo, v, hi, "Assumed", note or "A／C 以 B ± 約 2–4pt 設定",
            "FY2026 後分部營收以公司指引為起點，分部歷史未取得（轉向 2025-09 才納入合併）" if s == "B" else "")

sec("C2. 命題模組二：人形機器人線性關節（併入產業機械以外的子列）＝行星滾柱絲杠市場（線 6 下游取數）× NSK 份額 × 爬坡 × 匯率")
for yy, rr in ((2026, 12), (2030, 13), (2032, 14)):
    for s, col in (("A", "D"), ("B", "C"), ("C", "B")):
        v = {("A", 2026): 731.5, ("B", 2026): 298.571428571429, ("C", 2026): 89.5714285714286,
             ("A", 2030): 6545.958375, ("B", 2030): 1558.56151785714, ("C", 2030): 233.784227678571,
             ("A", 2032): 13794.2435339844, ("B", 2032): 3378.18208995536, ("C", 2032): 337.818208995536}[(s, yy)]
        inp(f"m{str(yy)[2:]}_{s}", f"情境 {s}：行星滾柱絲杠市場 {yy}（線 6；{'高' if s == 'A' else '基準' if s == 'B' else '低'}需求情境）", "百萬美元", None, v, None, "Derived",
            f"{L6} 下游取數!{col}{rr}（基準單價口徑＝覓途；五線歸因，全來自人形）", "線 6 需求情境：A＝高（D 欄）、B＝基準（C 欄）、C＝低（B 欄）；年間以幾何插值（Derived）")
for s, v, lo, hi in (("A", 0.03, None, None), ("B", 0.015, 0.0, 0.04), ("C", 0.0, None, None)):
    inp(f"shr_{s}", f"情境 {s}：NSK 在人形線性關節（以行星滾柱絲杠市場為可及市場代理）的份額", "%", lo, v, hi, "Assumed",
        f"未取得：NSK 未揭露人形機器人營收或客戶；行星滾柱絲杠現有領先者為歐系（Schaeffler、Rollvis、GSA、SKF），NSK 未列在清單（{L6} 份額!B11）", "代理口徑：NSK 以滾珠絲杠、線性導軌、軸承參與；份額數據未找到；C 情境為 0")
inp("rp27", "人形營收爬坡：FY2027 占穩態比例（線性升到 FY2030＝100%）", "%", 0.0, 0.25, 0.5, "Assumed", "比照 THK 模型與上銀模型的爬坡 2027 25% → 2030 100%", "")
inp("fx", "美元兌日圓（人形市場換算）", "¥/US$", 140.0, 155.0, 165.0, "Interested-party", Q2 + "：H2 匯率假設 USD/JPY 155", "低／高為 Assumed", CO)

sec("D. 損益表結構（標準損益表；IFRS）")
inp("mm_gp", "邊際毛利率（Δ毛利 ÷ Δ營收，相對 FY2026）", "x", 0.18, 0.25, 0.32, "Assumed",
    "FY2025→FY2026 指引：(1,050×22.5% − 191.291) ÷ (1,050 − 911.644) ＝ 32%（Derived；含轉向合併與匯率）；FY2023→FY2024：(172.784 − 165.249) ÷ (796.667 − 788.867) ≈ 97%（取整雜訊）", "基準取 25%：低於合併貢獻期的 32%；減收時同率回落")
for s, v, lo, hi in (("A", 0.005, None, None), ("B", 0.0, -0.01, 0.005), ("C", -0.015, None, None)):
    inp(f"gmadj_{s}", f"情境 {s}：FY2027 起毛利率額外調整（價格、原料、中國同業競爭）", "pt", lo, v, hi, "Assumed",
        "毛利率：FY2023 20.9%、FY2024 21.7%、FY2025 21.0%、FY2026 假設 22.5%（Derived／Assumed）", "C：車用軸承價格壓力與中國同業")
inp("sga_el", "銷管費用（含研發）對營收的彈性（費用年增 ÷ 營收年增）", "x", 0.25, 0.4, 0.7, "Assumed",
    "FY2025：銷管＋研發 +11.1% ÷ 營收 +14.4% ＝ 0.77（含轉向合併）；FY2024：+2.3% ÷ +1.0%（" + SAI + "）", "結構改革 FY2026 起降固定費用（公司：特定項目 15，" + Q2 + "），基準 0.4")
inp("oth_op", "其他營業損益淨額（結構改革費用等）FY2027 起", "¥B", -10.0, -3.0, 0.0, "Assumed", "FY2026 指引含 15 的特定項目（Derived）；產業機械 H2 含約 5 結構改革費用（" + Q2 + "）", "FY2025 特定項目約 6.4（資料商）；基準假設 FY2027 起降到 3")
inp("nonop", "業外：金融收益 − 金融費用（FY2027 起）", "¥B", -5.0, -2.5, 0.0, "Assumed", "FY2026 指引隱含 −2.5（＝53.5 − 56.0，Derived）", "H1 實績 −1.2（28.909 − 30.066）；借款 3,351 億")
inp("tax_w", "有效稅率（FY2027 起）", "%", 0.28, 0.34, 0.40, "Assumed", "FY2026 指引隱含 (53.5 − 34 − 非支配)/53.5 ≈ 34%（Derived）；實績 FY2023 45.6%、FY2024 47.9%、FY2025 36.1%（" + SAI + "）", "日本法定實效約 30.6%；歷史偏高（虧損地區未認列稅損，原因未查核）")

sec("E. 現金流與股東回饋")
inp("fc_da", "折舊攤銷 ÷ 營收（FY2026 起）", "%", 0.052, 0.059, 0.065, "Derived", "TTM D&A 60.421 ÷ 營收 1,019.745 ＝ 5.9%（" + SAS + "；" + SAI + "）", "FY2025 55.115 ÷ 911.644 ＝ 6.0%（資料商標準化，含使用權資產折舊）")
inp("fc_cx", "資本支出 ÷ 營收（FY2026 起）", "%", 0.045, 0.060, 0.075, "Assumed", "FY2024 設備投資 56.5 ÷ 營收 796.7 ＝ 7.1%（" + Q4 + "）；資料商 TTM 資本支出 40.0（" + SAS + "）二者不一致", "基準取約等於折舊（6.0%）使終值不虛增；公司 FY2026 資本支出計畫未取得")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.20, 0.30, 0.45, "Assumed", "未取得應收、存貨、應付明細；軸承業存貨重（類比 THK 同口徑 45.8%，THK 模型 In!fc_wc）", "基準 30%")
inp("doe", "股東權益配息率（DOE）", "%", None, 0.025, None, "Interested-party", Q1 + "：DOE 下限 2.5%；年間股利 34 維持", "用於收支關卡（股利＝DOE × 前一年底權益；約 17 十億日圓，與 34 日圓 × 4.9 億股相符）；不進 DCF", CO)
inp("eq_26", "FY2026 年底歸屬母公司權益（2026-09-30 權益 ＋ H2 淨利）", "¥B", None, "=690.404+(g26_ni-18.838)", None, "Derived", SAB + "：普通股權益 690.404（含非支配前 710.202）；H1 淨利 18.838（" + Q2 + "）", "中間股利約 8.3 於 2026-12 支付，不在此扣除（與 THK 範本同口徑）")

sec("F. 估值輸入（評價日 2026-09-30＝最近資產負債表日；目標時點 2027-11-20）")
inp("px", "股價（2026-11-20 收盤）", "¥", None, 892.5, None, "Verified", YCL + "：6471.T 2026-11-20 收盤 892.5", "未還原股利。2026-11-13 收盤 1,002 → 11-20 892.5（−10.9%），與全球 AI 與半導體股當週大跌同期（下跌非資料錯誤）")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("sh_out", "流通股數（資料商：490.04 百萬股）", "百萬股", None, 490.04, None, "Derived", SAS + "：490.04M；EPS 基本 46.75 ÷ 淨利 22.867 ＝ 489 百萬股（" + SAI + "）", "未取得公司自己株式明細；股票報酬與選擇權視為 0")
inp("sh_v", "評價日股數", "百萬股", None, "=sh_out", None, "Derived", "＝流通股數", "DCF 標準規格第 3 節")
inp("cash_v", "現金及約當現金（2026-09-30）", "¥B", None, 150.054, None, "Verified", SAB + "：150.054", "短期投資：資料商未列")
inp("debt_v", "有息負債（2026-09-30；短期 135.314＋長期 199.785）", "¥B", None, "=135.314+199.785", None, "Verified", SAB + "：總負債 335.099", "資料商未單獨列租賃負債，可能已含在內（故 lease_v 基準 0）")
inp("lease_v", "租賃負債代理（若未含在有息負債內）", "¥B", 0.0, 0.0, 20.0, "Assumed", "未取得租賃負債明細", "高案 20 為 Assumed")
inp("inv_v", "其他非流動金融資產（政策保有股、關聯企業投資）", "¥B", None, 0.0, None, "Assumed", "未取得；NSK 有政策保有股與持分法投資，未查核金額，保守設 0", "設 0 使 DCF 偏保守")
inp("inv_k", "其他非流動金融資產計入比例", "x", None, 1.0, None, "Assumed", "基準帳面全額（inv_v＝0 故無影響）", "")
inp("mi_v", "非支配股東權益（2026-09-30）", "¥B", None, 19.798, None, "Verified", SAB + "：710.202 − 690.404 ＝ 19.798", "帳面值，未調整為市值")
inp("stub_frac", "Stub 期（2026-09-30 → FY2026 年底 2027-03-31）占 FY2026 FCFF 比例", "年", None, "=(DATE(2027,3,31)-DATE(2026,9,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（日本 10 年期國債）", "%", None, 0.02817, None, "Verified", JGB10, "")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "日本國家風險溢酬", "%", None, 0.0091, None, "Verified", "Damodaran 國家風險溢酬（2026-01-05）：日本 0.91%（與本專案日本個股同值）", "")
inp("beta", "β（NSK 與 THK 5 年 β，Blume 調整後中位數）", "x", "=E{r}-0.3", "=MEDIAN(0.67*0.55+0.33,0.67*0.76+0.33)", "=E{r}+0.3", "Derived",
    SAS + "：NSK 5 年 β 0.55；THK 5 年 β 0.76（" + THX + " In!beta 說明）", "Blume 0.67β＋0.33；敏感度 ±0.3（Assumed）")
inp("kd", "稅前負債成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt；FY2025 利息費用 5.293 ÷ 負債約 3,350 ＝ 1.6%（" + SAI + "），取較保守", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("dps_tgt", "評價日到目標時點之每股股利（2026-12 中間 17＋2027-06 期末 17）", "¥", None, 34.0, None, "Assumed", Q1 + "：年間 34 維持（17＋17）", "2027-11-20 前的支付：中間（2026-12）與期末（2027-06）")
inp("t_tgt", "評價日到目標時點（年）", "年", None, "=(tp_date-DATE(2026,9,30))/365", None, "Derived", "2026-09-30 → 2027-11-20", "")
inp("ntm_w", "目標時點 NTM 對 FY2027 的權重（2027-11-20 起 12 個月：FY2027 剩 4.4 個月、FY2028 7.6 個月）", "x", None, 0.36, None, "Derived", "（2028-03-31 − 2027-11-20）÷ 365 ＝ 0.36", "3 月決算，NTM EPS ＝ 0.36 × FY2027 ＋ 0.64 × FY2028")
inp("w_dcf", "DCF 權重（其餘為 Comparable：同業前瞻 P/E）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺 NSK 對同業 10 年前瞻 P/E 比值歷史）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業前瞻 P/E（一律 2026-11-20 收盤口徑；取自本專案其他個股模型的同業表；報價日無法換算到 11-20 的同業不納入）")
PEERS = [("pe_thk", "THK 6481（線性導軌、滾珠絲杠）", 10.6780788096795, THX + " Comp 區（SMC 模型 In!pe_thk 同值）", "stockanalysis.com forward P/E（賣方共識 NTM EPS）"),
         ("pe_nab", "Nabtesco 6268（精密減速機）", 14.6736010072748, SMX + " In!pe_nab", "traders.co.jp PER（公司本期 EPS 予想）"),
         ("pe_fanuc", "Fanuc 6954（機器人、CNC）", 20.4505989144675, SMX + " In!pe_fanuc", "traders.co.jp PER（公司本期 EPS 予想）"),
         ("pe_yask", "安川電機 6506（伺服、機器人）", 15.6738990825688, SMX + " In!pe_yask", "stockanalysis.com forward P/E（賣方共識 NTM EPS）"),
         ("pe_nidec", "Nidec 6594（馬達、車用）", 11.5719333333333, SMX + " In!pe_nidec", "stockanalysis.com forward P/E（賣方共識 NTM EPS）"),
         ("pe_okuma", "Okuma 6103（工具機）", 10.5625158831004, SMX + " In!pe_okuma", "traders.co.jp PER（公司本期 EPS 予想）"),
         ("pe_dmg", "DMG Mori 6141（工具機）", 23.7848260447036, SMX + " In!pe_dmg", "traders.co.jp PER（公司本期 EPS 予想）")]
for k, lab, v, src, basis in PEERS:
    inp(k, f"前瞻 P/E：{lab}", "x", None, v, None, "Derived", f"取自 {src}（2026-11-20 收盤；原報價日 P/E × 新收盤 ÷ 舊收盤，分母 EPS 不變）",
        f"分母口徑：{basis}；會計口徑：日本基準或 IFRS（P/E 為股權口徑，不受租賃處理影響）", DV)
inp("pe_med", "同業前瞻 P/E 中位數（7 家；進 Comparable）", "x", 10.68, "=MEDIAN(pe_thk,pe_nab,pe_fanuc,pe_yask,pe_nidec,pe_okuma,pe_dmg)", 20.45, "Derived",
    "7 家中位數；不含 Keyence 29.84（感測器，非機械零組件）。報價日未取得（收盤檔無該 P/E 資料日收盤或無該股，無法換算至 11-20，不納入）：Nippon Thompson（IKO）6480、直得 1597、上銀 2049、Schaeffler、五洲新春 603667",
    "低＝同業最低（THK 10.68）；高＝Fanuc 20.45（Assumed 區間）；7 家雙口徑混用（公司予想／賣方共識）列為限制；少於約 8 家不做迴歸")
inp("pe_self", "NSK 現行前瞻 P/E（現價 ÷ FY2026 指引 EPS＝34.0 ÷ 490.04）", "x", None, "=px/(g26_ni/sh_v*1000)", None, "Derived", "指引淨利 34 ÷ 股數；資料商 11-24 頁面 11.70（價 926.2，賣方 NTM），只作對照", "只作對照，不進模型", DV)
inp("cons_tp", "賣方平均目標價（6 位；資料商 2026-11-24 頁面）", "¥", None, 1296.0, None, "Interested-party", SAS + "：1,296、共識 Hold", "只作對照，不進模型；11-24 頁面價 926.2（11-20 收盤 892.5 之後）", SELL)

def nu(t):
    """含「未取得」的說明文字不放網址（HTML 自含，網址見 Sources 頁）。"""
    return re.sub(r"https?://[^\s；）)]+", "（網址見 Sources 頁）", t) if isinstance(t, str) and "未取得" in t else t


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
        IN.cell(r, 7, tag); IN.cell(r, 8, nu(s1)); IN.cell(r, 9, nu(s2)); IN.cell(r, 10, nu(inc))
        dn(k, f"In!$E${r}"); KR[k] = r; r += 1
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("i27", "產業機械營收年增 FY2027"), ("i28", "產業機械營收年增 FY2028"), ("im", "產業機械營收年增 FY2029–FY2030"), ("ilt", "產業機械營收年增 FY2031–FY2032"),
           ("x27", "汽車與轉向營收年增 FY2027"), ("x28", "汽車與轉向營收年增 FY2028"), ("xm", "汽車與轉向營收年增 FY2029–FY2030"), ("xlt", "汽車與轉向營收年增 FY2031–FY2032"),
           ("m26", "行星滾柱絲杠市場 2026（線 6）"), ("m30", "行星滾柱絲杠市場 2030（線 6）"), ("m32", "行星滾柱絲杠市場 2032（線 6）"),
           ("shr", "NSK 人形線性關節份額"), ("gmadj", "FY2027 起毛利率調整")]
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
    H["A1"] = "Hist — 歷史財務（年度；¥B）。FY2023＝2024 年 3 月期、FY2024＝2025 年 3 月期、FY2025＝2026 年 3 月期（公司稱呼）。營收、毛利、銷管（含研發）、稅前、稅、淨利、折舊取自資料商標準化口徑；FY2024–25 營業利益取公司值（資料商摘要），與標準化口徑不同（見來源欄）。公式只引用本頁（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2023（2024年3月期）", "FY2024（2025年3月期）", "FY2025（2026年3月期）", "標記", "來源"])
    NA = "未取得"
    HR = [("h_rev", "營收", 788.867, 796.667, 911.644, "Verified", SAI + "；FY2025 與 " + Q4 + " 一致（796.7）、FY2025 911.6（資料商摘要）"),
          ("h_cogs", "營業成本", "=B3-B5", "=C3-C5", "=D3-D5", "Derived", "營收 − 毛利"),
          ("h_gp", "毛利", 165.249, 172.784, 191.291, "Verified", SAI),
          ("h_sga", "銷管費用（含研發）", "=119.103+11.794", "=122.524+12.214", "=134.719+14.921", "Verified", SAI + "：SG&A＋R&D"),
          ("h_oth", "其他營業損益淨額（營業利益 − 毛利 ＋ 銷管；含結構改革費用、持分法）", "=B8-B5+B6", "=C8-C5+C6", "=D8-D5+D6", "Derived", "差額；FY2025 資料商標準化營業利益 31.075 與公司值 38.8 不同，差額反映口徑"),
          ("h_op", "營業利益", "=28.5/1.039", 28.5, 38.8, "Verified", "FY2024 28.5、FY2025 38.8（+36.4%）：" + Q4 + "；資料商摘要（公司決算短信 FY2025 原文未直接讀取）；FY2023＝28.5 ÷ 1.039（Quartr：FY2024 +3.9%，Derived）"),
          ("h_fin", "業外：稅前利益 − 營業利益（金融損益及其他）", "=B10-B8", "=C10-C8", "=D10-D8", "Derived", "稅前（資料商）− 營業利益（公司）；兩來源口徑混用，只作大致"),
          ("h_pbt", "稅前利益", 26.208, 25.099, 38.037, "Verified", SAI),
          ("h_tax", "所得稅", 11.955, 12.031, 13.730, "Verified", SAI),
          ("h_cont", "稅後淨利（持續營業，未扣非支配）", "=B10-B11", "=C10-C11", "=D10-D11", "Derived", "稅前 − 稅"),
          ("h_disc", "停業單位及其他（歸屬母公司淨利 − 持續營業歸屬母公司）", "=B14-(B12-B15)", "=C14-(C12-C15)", "=D14-(D12-D15)", "Derived", "FY2024 資料商：淨利 10.647 含停業單位（Quartr：持續營業 12.4、含停業 10.6）"),
          ("h_ni", "歸屬母公司淨利", 8.502, 10.647, 22.867, "Verified", SAI + "；FY2025 22.9（公司，" + Q4 + "）"),
          ("h_nci", "非支配股東損益", 1.062, 0.546, 1.440, "Verified", SAI),
          ("h_eps", "基本 EPS（日圓）", 17.26, 21.78, 46.75, "Verified", SAI),
          ("h_epsc", "基本 EPS（持續營業，日圓）", "=(B12-B15)/(B14/B16*1000)*1000", "=(C12-C15)/(C14/C16*1000)*1000", "=(D12-D15)/(D14/D16*1000)*1000", "Derived", "（稅後 − 非支配）÷ 股數"),
          ("h_da", "折舊攤銷（資料商標準化）", 54.121, 52.412, 55.115, "Verified", SAI),
          ("h_cx", "設備投資", NA, 56.5, NA, "Verified", Q4 + "：FY2024 56.5；FY2023、FY2025 未取得")]
    for i, (k, lab, a, b, c, tag, src) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b, c)):
            cc = H.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): cc.font = BLUE
        H.cell(i, 5, tag); H.cell(i, 6, nu(src))
        dn(k + "_23", f"Hist!$B${i}"); dn(k + "_24", f"Hist!$C${i}"); dn(k + "_25", f"Hist!$D${i}")
    H.column_dimensions["A"].width = 44
    HRW = {k: 3 + i for i, (k, *_x) in enumerate(HR)}

    # ---------------- Q（官方分部＋日本單體受注） ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — FY2026（2027年3月期）公司指引分部與中間期（H1）實績（¥B）。報告分部：產業機械、汽車、轉向（2025-09 起合併）；本次取得的來源未給轉向分部數字，第三列為差額"
    hdr(Q, 2, ["項目", "FY2026 指引", "H1 FY2026 實績", "標記", "來源", "來源誘因"])
    Q.cell(3, 1, "產業機械 營收"); Q.cell(3, 2, 436.0).font = BLUE; Q.cell(3, 3, "未取得"); Q.cell(3, 4, "Interested-party"); Q.cell(3, 5, Q2); Q.cell(3, 6, CO)
    Q.cell(4, 1, "汽車 營收"); Q.cell(4, 2, 403.0).font = BLUE; Q.cell(4, 3, "未取得"); Q.cell(4, 4, "Interested-party"); Q.cell(4, 5, Q2); Q.cell(4, 6, CO)
    Q.cell(5, 1, "轉向與其他、內部調整（差額）營收"); Q.cell(5, 2, "=B6-B3-B4"); Q.cell(5, 3, "未取得"); Q.cell(5, 4, "Derived"); Q.cell(5, 5, "合併 − 產業機械 − 汽車；未取得公司給的轉向分部營收，差額不是轉向分部營收")
    Q.cell(6, 1, "合併營收"); Q.cell(6, 2, 1050.0).font = BLUE; Q.cell(6, 3, 520.351).font = BLUE; Q.cell(6, 4, "Interested-party"); Q.cell(6, 5, Q2 + "；H1 520.351（對照 " + TS2Q + "）"); Q.cell(6, 6, CO)
    Q.cell(7, 1, "檢查：分部合計 − 合併營收（應為 0）"); Q.cell(7, 2, "=ROUND(B3+B4+B5-B6,3)"); Q.cell(7, 4, "Derived"); Q.cell(7, 5, "第三列為差額，故檢查必為 0；真正的分部加總檢查需轉向分部營收（未取得）")
    Q.cell(8, 1, "營業利益（合併）"); Q.cell(8, 2, 56.0).font = BLUE; Q.cell(8, 3, 30.066).font = BLUE; Q.cell(8, 4, "Interested-party"); Q.cell(8, 5, Q2); Q.cell(8, 6, CO)
    Q.cell(9, 1, "產業機械 營業利益（調整後 44.0）"); Q.cell(9, 2, 37.5).font = BLUE; Q.cell(9, 3, "未取得"); Q.cell(9, 4, "Interested-party"); Q.cell(9, 5, Q2); Q.cell(9, 6, CO)
    Q.cell(10, 1, "汽車 營業利益"); Q.cell(10, 2, 17.5).font = BLUE; Q.cell(10, 3, "未取得"); Q.cell(10, 4, "Interested-party"); Q.cell(10, 5, Q2); Q.cell(10, 6, CO)
    Q.cell(11, 1, "稅前利益"); Q.cell(11, 2, 53.5).font = BLUE; Q.cell(11, 3, 28.909).font = BLUE; Q.cell(11, 4, "Interested-party"); Q.cell(11, 5, Q2); Q.cell(11, 6, CO)
    Q.cell(12, 1, "歸屬母公司淨利"); Q.cell(12, 2, 34.0).font = BLUE; Q.cell(12, 3, 18.838).font = BLUE; Q.cell(12, 4, "Interested-party"); Q.cell(12, 5, Q2); Q.cell(12, 6, CO)
    Q.cell(13, 1, "H1 FY2025（前年同期）營收／營業利益"); Q.cell(13, 2, 412.250).font = BLUE; Q.cell(13, 3, 17.708).font = BLUE; Q.cell(13, 4, "Interested-party"); Q.cell(13, 5, Q2 + "（營業利益前年同期公司短信表為 16.467，資料商 17.708，未核對重編）"); Q.cell(13, 6, CO)
    Q.cell(14, 1, "H2 隱含（指引 − H1）：營收／營業利益"); Q.cell(14, 2, "=B6-C6"); Q.cell(14, 3, "=B8-C8"); Q.cell(14, 4, "Derived"); Q.cell(14, 5, "公式")
    Q.cell(15, 1, "註：轉向事業自 2025-09-01 納入合併，H1 年增 +26.2% 含合併效果（H1 FY2025 轉向貢獻營收增 15.2、營業利益增 3.8，資料商摘要）；FY2025 分部營收歷史未取得")
    Q.column_dimensions["A"].width = 46

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（產業機械、汽車與轉向循環＋人形線性關節）→ 分部營收 → 標準損益表（IFRS）→ 現金流（¥B；FY2023–25 實績（3 月期）、FY2026 公司指引 ×(1＋g_anc)、FY2027 起模型）"
    PR = {}
    rr = 3
    KEYS = ([f"c_{k}" for k, _ in REG] + ["core", "hum", "rev", "chk1", "chk2", "gr", "cogs", "gp", "gm", "sga", "oth", "opex", "op", "opm", "fin", "pbt", "tax", "cont",
             "disc", "nci", "ni", "nim", "sh", "eps", "epsc", "nig", "da", "cx", "dwc", "fcff", "eq", "div", "fcfad", "mkt", "rp"])
    LAB = ([f"營收：{n}（報告分部）" for _, n in REG] +
           ["核心營收（三分部合計）", "人形機器人線性關節（命題子列）", "營收（合併）", "檢查：分部合計 − 核心營收", "檢查：核心＋人形 − 合併營收",
            "營收年增率", "營業成本", "毛利", "毛利率", "銷售及管理費用", "其他營業損益淨額（其他收益 − 其他費用 ＋ 持分法）", "營業費用（銷管費用 − 其他營業損益淨額）",
            "營業利益", "營業利益率", "業外：金融收益 − 金融費用（其他收入淨額）", "稅前利益", "所得稅", "持續營業稅後淨利", "停業單位及其他損益",
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
                for k, _ in REG: w(f"c_{k}", '="未取得"')
                w("core", f"={cell('rev')}"); w("hum", 0); w("rp", 0)
                w("chk1", 0); w("chk2", f"=ROUND({cell('core')}+{cell('hum')}-{cell('rev')},3)")
                if y == 2023: w("gr", '="未取得（FY2022 未建）"'); w("nig", '="未取得（FY2022 未建）"')
                else:
                    w("gr", f"={cell('rev')}/{prev('rev')}-1"); w("nig", f"=IF({prev('ni')}>0,{cell('ni')}/{prev('ni')}-1,\"N/M\")")
                w("sh", f"={cell('ni')}/{cell('eps')}*1000")
                continue
            if y == 2026:
                for k, _ in REG: w(f"c_{k}", f"=g26_{k}*(1+g_anc)")
                w("hum", 0); w("rp", 0)
                w("gp", f"=gm26*{cell('rev')}"); w("oth", "=oth26"); w("sga", f"={cell('gp')}+oth26-g26_op*(1+g_anc)")
                w("fin", "=g26_pbt-g26_op"); w("disc", 0); w("nci", "=g26_nci")
                w("tax", f"={cell('pbt')}-(g26_ni+g26_nci)*(1+g_anc)")
                w("sh", "=sh_v")
                w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}"); w("eq", "=eq_26"); w("div", f"=dps_26*{cell('sh')}/1000")
            else:
                gk = "27" if y == 2027 else "28" if y == 2028 else "m" if y <= 2030 else "lt"
                for k, _ in REG: w(f"c_{k}", f"={prev('c_' + k)}*(1+{'i' if k == 'im' else 'x'}{gk}_{s})")
                if y <= 2029: mk = f"=m26_{s}*(m30_{s}/m26_{s})^(({y}-2026)/4)"
                elif y == 2030: mk = f"=m30_{s}"
                elif y == 2031: mk = f"=m30_{s}*(m32_{s}/m30_{s})^0.5"
                else: mk = f"=m32_{s}"
                w("mkt", mk); w("rp", f"=MIN(1,rp27+(1-rp27)*({y}-2027)/3)")
                w("hum", f"={cell('mkt')}*shr_{s}*{cell('rp')}*fx/1000")
                w("gp", f"=$F{R['gp']}+mm_gp*({cell('rev')}-$F{R['rev']})+gmadj_{s}*{cell('rev')}")
                w("sga", f"=$F{R['sga']}*(1+sga_el*({cell('rev')}/$F{R['rev']}-1))"); w("oth", "=oth_op")
                w("fin", "=nonop"); w("tax", f"={cell('pbt')}*tax_w"); w("disc", 0); w("nci", "=g26_nci"); w("sh", "=sh_v")
                w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}")
                w("div", f"=doe*{prev('eq')}"); w("eq", f"={prev('eq')}+{cell('ni')}-{cell('div')}")
            w("core", f"=SUM({c}{R['c_im']}:{c}{R['c_st']})")
            w("rev", f"={cell('core')}+{cell('hum')}")
            w("chk1", f"=ROUND(SUM({c}{R['c_im']}:{c}{R['c_st']})-{cell('core')},6)")
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
        P.cell(b, 2, "FY2027 起＝前一年 ×(1＋該分部成長率)；FY2026＝公司指引")
        P.cell(R["hum"], 2, "市場 × 份額 × 爬坡 × 匯率；FY2026 以前未揭露（視為 0）；疊加於三分部之上")
        P.cell(R["chk2"], 2, "核心＋人形 − 合併營收（應為 0）")
        P.cell(R["gp"], 2, "FY2026＝毛利率假設 × 營收；FY2027 起＝FY2026 毛利 ＋ 邊際毛利率 × Δ營收 ＋ 毛利率調整 × 營收")
        P.cell(R["sga"], 2, "FY2026＝使營業利益等於指引；FY2027 起＝FY2026 ×(1＋彈性 × 營收相對 FY2026 增幅)")
        P.cell(R["tax"], 2, "FY2026＝稅前 − （指引淨利＋非支配）（Derived）")
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
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-09-30，明確期 FY2027–FY2032（3 月期，期中慣例：FY2027 折現 1.0 年），Stub＝FY2026 FCFF × 半年；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
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
            D[f"{c}{b+7}"] = j + 1.0; D[f"{c}{b+8}"] = f"=1/(1+wacc)^{c}{b+7}"; D[f"{c}{b+9}"] = f"={c}{b+5}*{c}{b+8}"
        t = b + 11
        items = [("明確期現值合計", f"=SUM(B{b+9}:G{b+9})"),
                 ("NOPAT FY2033＝NOPAT FY2032 ×(1＋g)", f"=G{b+1}*(1+g_tv)"),
                 ("終值（FY2032 年底，2033-03-31）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^6.5"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金與投資（2026-09-30）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac；2026-10 至 2027-03）", f"=PL!F{R['fcff']}*stub_frac"),
                 ("− 尚待買回金額（無）", 0),
                 ("股權價值", f"=B{t+4}+B{t+5}+B{t+6}+B{t+7}"),
                 ("每股價值（評價日）", f"=MAX(0,B{t+8})/sh_v*1000"),
                 ("DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）", f"=B{t+9}*(1+ke)^t_tgt-dps_tgt"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+11}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=B{t+2}/G{b+1}"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）", f"=((B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6.5)+B{t+5}+B{t+6}+B{t+7})/sh_v*1000*(1+ke)^t_tgt-dps_tgt")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 10, ps=t + 9, ev=t + 4, tvs=t + 11, alt=t + 14, chk=b + 6)
        dn(f"dcf_{s}", f"DCF!$B${t+10}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 60

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價：DCF（權重 w_dcf）＋ Comparable（同業前瞻 P/E 中位數 × 情境 NTM EPS，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}")
        V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "DCF 頁")
    r_cu = vrow("Comparable（未調整）：同業前瞻 P/E 中位數 × 情境 NTM EPS（持續營業）", [f"=pe_med*(ntm_w*PL!G{PR[s]['epsc']}+(1-ntm_w)*PL!H{PR[s]['epsc']})" for s in SC],
                "目標時點（2027-11-20）NTM EPS＝ntm_w × FY2027 ＋ (1 − ntm_w) × FY2028（3 月期）；分母：模型各情境持續營業 EPS")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（日圓）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    r_cons = one("市場共識目標價（賣方 6 位平均；資料商 2026-11-24 頁面）", "=cons_tp", "只作對照（Interested-party）")
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
    r_pd1 = one("期間：模型折現（評價日 2026-09-30 → 目標時點 2027-11-20，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B NTM EPS", f"=B{r_tp}/(ntm_w*PL!G{PR['B']['epsc']}+(1-ntm_w)*PL!H{PR['B']['epsc']})", "分母：模型情境 B NTM EPS（目標時點）")
    r_pe1 = one("現價 ÷ 情境 B FY2027 持續營業 EPS", f"=px/PL!G{PR['B']['epsc']}", "分母：模型情境 B FY2027")
    r_pe2 = one("現行前瞻 P/E（現價 ÷ FY2026 指引 EPS）÷ 同業中位數", "=pe_self/pe_med", "只作對照；相對重估需要的倍數＝1 ÷ 本值")
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
          ("FY2026 營業利益", f"=PL!F{PR['B']['op']}", "=g26_op", "公司指引（含特定項目 15）"),
          ("FY2026 歸屬母公司淨利", f"=PL!F{PR['B']['ni']}", "=g26_ni", "公司指引"),
          ("FY2026 稅前利益", f"=PL!F{PR['B']['pbt']}", "=g26_pbt", "公司指引"),
          ("FY2026 調整後營業利益（營業利益 − 其他營業損益）vs 公司調整後 71", f"=PL!F{PR['B']['op']}-PL!F{PR['B']['oth']}", 71.0, "公司指引調整後營業利益 71（" + "資料商 Quartr 摘要" + "）"),
          ("FY2026 EPS（模型 B）vs 指引淨利 ÷ 股數", f"=PL!F{PR['B']['eps']}", "=g26_ni/sh_v*1000", "資料商 11-24：前瞻 P/E 11.70 對應 NTM EPS 約 79（價 926.2，賣方 NTM；口徑不同，只看方向）"),
          ("加權目標價 vs 賣方平均目標價", "=tp", "=cons_tp", "賣方 6 位平均（資料商 2026-11-24，只作對照）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_nsk.py 批次格點＋線性內插；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 60; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_nsk.py 以 LibreOffice 批次重算每一案例）；結果為常數快照"
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
           ("滾珠絲杠需求量 2030（主要來自線控煞車 EMB）", "6", "1,948.8／4,964.2／9,526.0", "萬根", "只計增量（EMB、SBW、機器人）", L6 + " 下游取數!B23:D23", "Derived", "未進公式", "不一致（自設 0）：NSK 有線控煞車用滾珠絲杠與轉向事業，但未取得 NSK 在該增量的份額或訂單，汽車與轉向營收改以車產成長率（x27 等）推導；此為最大的潛在上檔來源，未建模"),
           ("滾珠絲杠與線性導軌份額", "6", "未取得", "%", "THK、上銀、NSK", L6 + " 份額!F12", "—", "未進公式", "一致（線模型亦未取得）"),
           ("行星滾柱絲杠領先者", "6", "Schaeffler、Rollvis、GSA、SKF；NSK 未列", "—", "定性", L6 + " 份額!B11", "Interested-party", "shr_* 說明（基準 1.5% 為 Assumed）", "一致"),
           ("人形用滾柱絲杠市場 2030（低／基準／高）", "4", "205.2／747.0／1,563.7", "百萬美元", "單價隨學習曲線；絲杠情形 3", L4 + " 下游取數!C72:E72", "Derived", "未進公式（對照）", "線 4 與線 6 不一致：線 6 基準 1,558.6 為線 4 的 2.1 倍（單價口徑不同）；本模型照所屬線 6，差異以 shr 區間吸收"),
           ("人形全球出貨 2030（低／基準／高）", "4", "20 萬／50 萬／89 萬", "台", "全球、含小型機", L4 + " 下游取數!C38:E38", "Derived", "未進公式（線 6 市場已含）", "一致（經線 6）"),
           ("全球工業機器人裝機 2026→2030（基準）", "3", "65.5 萬 → 85.4 萬（年複合 6.9%）", "台", "IFR 新裝機口徑", L3 + " 下游取數!E13、E17", "Derived", "im_B（4%）的上緣類比；未進公式", "自設低於線 3：NSK 產業機械營收主體為工具機、半導體設備與一般機械，機器人只占一部分（未揭露）")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 30; RE.column_dimensions["F"].width = 60; RE.column_dimensions["I"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("半導體與工具機循環反轉", "是（i27_C、i28_C）", "情境 C", "產業機械季營收與營業利益率（H2 目標 ≥ 11%）、SEMI 設備出貨"),
             ("中國與歐洲車用軸承需求疲弱、價格競爭", "是（x27_C、gmadj_C）", "情境 C", "汽車分部營業利益率（目標 5%）、中國與歐洲車產"),
             ("人形線性關節份額為零或採用行星滾柱絲杠以外的方案", "是（shr_C、Sens）", "人形子列占營收比例小（見 Val 頁）", "Tesla、Figure 等供應鏈揭露；公司人形產品發表"),
             ("結構改革費用超出指引（FY2026 特定項目約 15）", "部分（oth_op 區間 −10 至 0）", "基準假設 FY2027 起降至 3", "調整後與報告營業利益差距"),
             ("匯率（美元、歐元、人民幣）", "否", "公司 H2 假設 USD/JPY 155、EUR/JPY 175、CNY/JPY 22.5；H1 匯率貢獻營業利益 7.1；未逐年建模", "USDJPY、EURJPY"),
             ("轉向事業合併後的整合、少數股權與美國關稅", "否", "關稅成本以價格轉嫁抵銷（H1 淨 0，" + "資料商摘要" + "）；轉向分部數字未取得", "分部披露、非支配股東損益"),
             ("稅率偏高（歷史 36–48%）", "部分（tax_w 區間 28–40%）", "原因未查核", "季度有效稅率")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 50

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([TS2Q, Q2, Q1, Q4, SAI, SAB, SAS, YCL, JGB10, L6 + "（下游取數、份額）", L3 + "（下游取數）", L4 + "（下游取數）",
                            SMX + "（In 頁同業 P/E，2026-11-20 口徑）", THX + "（THK P/E、β）"], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"NSK 日本精工（6471.T）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-20（2027 年 3 月期第 2 季決算 2026-11）；成品日 2026-11-24；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"#,##0")&" 日圓）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：NSK 是軸承（汽車、產業機械）與滾珠絲杠、線性導軌的全球大廠，營收約 1 兆日圓、營業利益率仍低（FY2026 指引 5.3%）；股價由產業機械（半導體設備、工具機）回升與結構改革能否把利益率拉高決定，人形機器人線性關節是小型選擇權，目前沒有 NSK 的實體 AI 營收數字"
    RD["A4"] = (f'="在 FY2027 產業機械情境 B 年增 "&TEXT(i27_B,"0%")&"、邊際毛利率 "&TEXT(mm_gp,"0%")&"、同業前瞻 P/E 中位數 "&TEXT(pe_med,"0.0")&" 倍的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"#,##0")&" 日圓（DCF "&TEXT(Val!B{r_dcft},"#,##0")&"、Comparable "&TEXT(Val!B{r_cmpt},"#,##0")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0.0")&" 十億日圓）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 日圓。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：產業機械分部營業利益率 H2 低於 11% 目標；公司 FY2026 指引再下修；汽車分部營業利益率低於 3%；FY2027 公司指引（2027-05）營業利益低於 FY2026 指引 56"
    RD["A7"] = "範圍：全公司（產業機械、汽車、轉向）；分部歷史與毛利指引未取得，見限制；不讀 Siliconomics／Tokenomics／Wattonomics（算力線）；自用 AI 讀法等規格條文不適用。HTML 展示層見 dist/ 同名 .html"
    RD["A8"] = f'="命題回覆：問：產業機械回升與結構改革能否支持高於現價的目標價？答：情境 B FY2027（2028 年 3 月期）持續營業 EPS "&TEXT(PL!G{PR["B"]["epsc"]},"#,##0")&" 日圓，現價只有其 "&TEXT(px/PL!G{PR["B"]["epsc"]},"0.0")&" 倍，低於同業中位數 "&TEXT(pe_med,"0.0")&" 倍；加權目標價 "&TEXT(Val!B{r_tp},"#,##0")&" 日圓。問：人形機器人有多重要？答：情境 B FY2030 人形子列只占合併營收 "&TEXT(Val!B{r_hum},"0.0%")&"，結論不靠人形（模型輸出，非獨立證據）。"'
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
def lo_profile():
    d = os.path.join(os.environ.get("HOME", tempfile.gettempdir()), "lo_nsk")
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
    d = tempfile.mkdtemp(prefix="nsk_"); paths = []; VR = DR = None
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
    ("g_anc", "FY2026 實績相對公司指引"), ("i27_B", "產業機械營收年增 FY2027（情境 B）"), ("i28_B", "產業機械營收年增 FY2028（情境 B）"), ("im_B", "產業機械營收年增 FY2029–30（情境 B）"),
    ("ilt_B", "產業機械營收年增 FY2031–32（情境 B）"), ("x27_B", "汽車與轉向營收年增 FY2027（情境 B）"), ("x28_B", "汽車與轉向營收年增 FY2028（情境 B）"), ("xm_B", "汽車與轉向營收年增 FY2029–30（情境 B）"),
    ("xlt_B", "汽車與轉向營收年增 FY2031–32（情境 B）"), ("shr_B", "人形線性關節份額（情境 B）"), ("rp27", "人形營收爬坡 FY2027"), ("fx", "美元兌日圓"),
    ("gm26", "FY2026 毛利率"), ("oth26", "FY2026 特定項目（其他營業損益）"), ("g26_nci", "FY2026 非支配股東損益"),
    ("mm_gp", "邊際毛利率"), ("gmadj_B", "毛利率調整（情境 B）"), ("sga_el", "銷管費用彈性"), ("oth_op", "其他營業損益淨額"), ("nonop", "業外金融損益"), ("tax_w", "有效稅率"),
    ("fc_da", "折舊 ÷ 營收"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("lease_v", "租賃負債代理"),
    ("beta", "β"), ("erp", "股權風險溢酬"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重"), ("pe_med", "同業前瞻 P/E 中位數")]
BETA_B = statistics.median([0.67 * 0.55 + 0.33, 0.67 * 0.76 + 0.33])
FXD = {"pe_med": statistics.median([10.6780788096795, 14.6736010072748, 20.4505989144675, 15.6738990825688, 11.5719333333333, 10.5625158831004, 23.7848260447036])}


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
    for key, b0, lo_, hi_, note in (("i27_B", 0.05, -0.40, 0.60, "情境 B 的 FY2027 產業機械營收年增（其他不變，整本重算；情境 A／C 不動；之後年度在此基礎上成長）"),
                                    ("mm_gp", 0.25, -0.50, 1.00, "邊際毛利率（三情境同動）"),
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
    tmp = tempfile.mkdtemp(prefix="nskf_"); raw = os.path.join(tmp, NAME)
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
    json.dump(dict(company="nsk", ticker="6471.T", model_sheet="Q",
                   official_segments=[dict(name=n, label=l, period="FY2026 公司指引（2027年3月期）", value=v, unit="¥B", source=Q2 + "（產業機械、汽車）；轉向與其他為差額（Derived）")
                                      for n, l, v in (("產業機械", "產業機械 營收", 436.0), ("汽車", "汽車 營收", 403.0), ("轉向與其他（差額）", "轉向與其他、內部調整（差額）營收", 211.0))],
                   sublines=[], check_labels=["檢查：分部合計 − 合併營收（應為 0）"],
                   note="報告分部：產業機械、汽車、轉向（2025-09 起合併）。本次取得的來源只給產業機械與汽車的 FY2026 指引營收，轉向與其他以合併減兩者的差額列示（Derived），故 Q 頁分部檢查必為 0，不具獨立驗證力；PL 頁分部營收以此為起點，FY2025 以前分部營收未取得。"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)


def write_readme(path, VR, PR, DR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f0 = lambda x: f"{x:,.0f}"; pc = lambda x: f"{x*100:+.1f}%"
    inv = lambda k: I['E' + str([r for r in range(1, I.max_row + 1) if I['A' + str(r)].value == k][0])].value
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:8]
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {r['val'] if isinstance(r['val'], str) else format(r['val'], '.4f')} | {r['base']:.4g} | {r['note']} |" for r in revs)
    def yr(s, k, fmt="{:,.1f}", ys=range(2023, 2033)):
        out = []
        for y in ys:
            v = P[f"{COL[y]}{PR[s][k]}"].value
            out.append(fmt.format(v) if isinstance(v, (int, float)) else "—")
        return " | ".join(out)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f0(g(chr(66+i)+str(VR['dcfr'])))} | {f0(g(chr(66+i)+str(VR['cu'])))} | {f0(g(chr(66+i)+str(VR['tps'])))} | {g(chr(66+i)+str(VR['gate'])):,.1f} |" for i, s in enumerate(SC))
    rts = sorted(set(c['rt'] for c in sens['cases']))
    YS = range(2023, 2033)
    txt = f"""# physical_ai/nsk — NSK 日本精工（6471.T）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

東證 Prime 上市；3 月決算（IFRS）。本檔沿用公司稱呼：FY2026＝2027 年 3 月期（公司指引年度）、FY2025＝2026 年 3 月期（實績）。本 README 的數字由 `build_nsk.py` 從重算後的成品讀出產生，不手寫。

## 未取得與限制（先讀）
- 公司 2027 年 3 月期第 2 季決算短信 PDF 的文字層無法讀取；FY2026 指引與 H1 實績取自資料商（Quartr）摘要，FY2023–25 財務取自資料商標準化口徑（營業利益 FY2024–25 取公司值，與標準化口徑不同）。未直接對照公司原文。
- **未取得**：FY2026 毛利與銷管費用指引（毛利率 22.5% 為 Assumed，區間 21.5–23.5%）；轉向分部營收與營業利益（Q 頁以差額列示，不是轉向營收）；FY2025 以前分部營收；資本支出計畫（資料商兩個數字不一致）；租賃負債、政策保有股與持分法投資金額（設 0，保守）；NSK 人形機器人營收與份額、客戶集中度。
- 同業前瞻 P/E 一律用 2026-11-20 收盤口徑，僅納入 7 家有換算值者；**報價日未取得、未納入**：Nippon Thompson（IKO）6480（收盤檔無此股）、直得 1597、上銀 2049、Schaeffler、五洲新春 603667（其 P/E 資料日為 8–10 月，收盤檔無該日收盤，無法換算）。7 家混用公司予想與賣方共識兩種分母口徑；少於約 8 家不做迴歸。
- 2026-11-16 至 11-20 全球 AI 與半導體股大跌屬實，收盤 892.5（11-13 為 1,002，−10.9%）不是資料錯誤；rf、CRP 沿用較早取值（跨模型一致，逾期）。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A8'].value}

**讀法：** 現價只有模型 FY2027 持續營業 EPS 的 {g(f"B{VR['pe1']}"):.1f} 倍，Comparable 用同業 7 家前瞻 P/E 中位數 {inv('pe_med'):.1f} 倍（區間 10.7–20.4）；FY2026 指引兩度上修（營收 1,000 → 1,050 十億日圓、營業利益 42 → 56）。**結論主要建立在 Assumed 輸入上**：FY2026 毛利率（22.5%）、FY2027 起產業機械與汽車成長率、邊際毛利率、結構改革費用降到 3（FY2026 為 15）與同業倍數；人形子列（情境 B FY2030 占營收 {g(f"B{VR['hum']}")*100:.1f}%）不是承重項。

## 估值（評價日 2026-09-30；目標時點 2027-11-20）
| 情境 | DCF（推到目標時點） | Comparable（P/E × NTM EPS） | 情境目標價 | 收支關卡（FY2026–30 FCF 扣股利，¥B） |
|---|---|---|---|---|
{scn}

- 加權目標價 **{f0(tp)} 日圓**（DCF {f0(dcf)}、Comparable {f0(cmp_)}），現價 {inv('px'):,.1f}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")} → 投資建議 **{g(f"B{VR['rec']}")}**。賣方平均目標價 {f0(inv('cons_tp'))}（6 位，2026-11-24 頁面，只作對照）。
- 方法：DCF 標準版 FCFF（{g(f"B{VR['wk']}")}，為 WACC／ke／g／RONIC）50%＋Comparable 50%（同業 7 家前瞻 P/E 中位數 {inv('pe_med'):.2f} 倍 × 各情境 NTM EPS）。相對倍數係數 1.00（未調整）。
- 隱含前瞻 P/E（加權目標價 ÷ 情境 B NTM EPS）{g(f"B{VR['pe']}"):.1f} 倍；NSK 現行前瞻 P/E（現價 ÷ 指引 EPS）{inv('pe_self'):.2f} 倍 ＝ 同業中位數的 {g(f"B{VR['pe2']}"):.2f} 倍。

## 營收與標準損益（情境 B，¥B；FY2023–25 實績（3 月期）、FY2026 公司指引、FY2027 起模型）
| 列 | {" | ".join(f"FY{y}" for y in YS)} |
|---|{"---|" * len(YS)}
| 產業機械 | {yr('B','c_im')} |
| 汽車 | {yr('B','c_au')} |
| 轉向與其他（差額） | {yr('B','c_st')} |
| 人形線性關節（命題子列） | {yr('B','hum')} |
| 營收（合併） | {yr('B','rev')} |
| 毛利 | {yr('B','gp')} |
| 營業利益 | {yr('B','op')} |
| 歸屬母公司淨利 | {yr('B','ni')} |
| EPS（持續營業，日圓） | {yr('B','epsc','{:,.1f}')} |

- 分部歷史未取得（FY2023–25 分部列為「—」）；FY2026 以後分部營收以公司指引為起點。FY2026 營業利益含特定項目（結構改革）15，調整後 71。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 產業機械營收年增（A／B／C）：FY2027 12%／5%／−5%；FY2028 8%／4%／−8%；FY2029–30 6%／4%／2%；FY2031–32 4%／3%／2%。汽車與轉向：FY2027 4%／1%／−5%；FY2028 3%／1.5%／−4%；之後 3%／2%／1%。
- 人形：線 6 行星滾柱絲杠市場（A＝高、B＝基準、C＝低需求情境）× NSK 份額 3%／1.5%／0% × 爬坡（FY2027 25% → FY2030 100%）× 155 日圓。
- 毛利：FY2026 毛利率 22.5% ＋ 邊際毛利率 25% × Δ營收 ＋ 情境調整（A +0.5pt、C −1.5pt）；銷管彈性 0.4；稅率 34%（歷史 36–48%，FY2026 指引隱含約 34%）。
- 現金流：折舊 5.9%、資本支出 6.0%、營運資金 30% × 營收增量；股利＝DOE 2.5% × 前一年底權益。
- 折現：rf 2.817%（逾期）、β {inv('beta'):.3f}、ERP 5%、CRP 0.91%、g＝MIN(3%, rf)、RONIC＝WACC。股數 490.04 百萬股；淨負債（含非支配、不含政策保有股）約 {-wb['DCF']['B10'].value:,.1f}。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| FY2027–28 產業機械成長（i27、i28） | 產業機械季營收與 H2 營業利益率 ≥ 11% 目標；半導體設備與工具機訂單；FY2027 公司指引（2027-05） |
| 邊際毛利率與結構改革費用（mm_gp、oth_op） | 季度毛利率、調整後與報告營業利益差距（FY2026 約 15） |
| 同業 P/E（pe_med） | 同業前瞻 P/E 換算到同一報價日後的中位數 |
| 汽車與轉向成長（x27） | 全球車產（公司假設 9,100 萬台）、中國與歐洲需求 |

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
- 線 6：行星滾柱絲杠市場 2026／2030／2032 低／基準／高（下游取數!B12:D14）→ m26／m30／m32 的情境 C／B／A，照取；滾珠絲杠需求 2030（下游取數!C23，主要來自線控煞車）**不計入**（自設 0，與線模型不一致）：NSK 有線控煞車與轉向事業，但未取得其份額，汽車與轉向營收改以車產成長率推導；這是未建模的潛在上檔。份額（份額!F12）線模型亦未取得。
- 線 4：人形用滾柱絲杠市場 2030 基準 747（下游取數 C72）與線 6 的 1,558.6 不一致（單價口徑），本模型照所屬線 6，差異由份額區間吸收。
- 線 3：工業機器人裝機 2026→2030 年複合 6.9%（下游取數 E13、E17）只作 FY2029–30 成長的上緣類比；自設 4% 低於線 3。

## 檔案
- `build_nsk.py`：建置＋整本重算 Sens／反解（LibreOffice 批次）＋產生本 README、`segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-24）：初版；現價 {inv('px'):,.1f}（2026-11-20 收盤）；加權目標價 {f0(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


if __name__ == "__main__":
    main()

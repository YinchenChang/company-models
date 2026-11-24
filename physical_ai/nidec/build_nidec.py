# Nidec 日本電產（6594.T）個股模型 v0.1 — AI實體應用 Project（線 6 共用零組件：馬達；並讀線 4 人形）
# 用法：python3 physical_ai/nidec/build_nidec.py   → 建置、整本重算 Sens 與反解（LibreOffice 批次）、輸出 dist/ 成品、README.md、sens.json、segments.json
# 會計年度＝3 月底決算（IFRS）。本模型以「FY2026＝2026 年 3 月期」稱呼（公司決算新聞稿標題為 Fiscal Year Ended March 31, 2026）；FY2027＝2027 年 3 月期（公司指引）。
# 金額：十億日圓（¥B）；股數：百萬股；EPS、股價：日圓。
# 結構：README、In、Hist、Q（官方分部）、PL（三情境）、DCF、Val、Recon、Sens、對帳、Risk、Sources。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261124"
NAME = f"{DATE}_Nidec個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032

# ---------------- 出處 ----------------
FY26 = "公司 FY2026（2026 年 3 月期）決算新聞稿〔IFRS〕（2026-09-30 公布，Business Wire／Barchart 轉載）https://www.webull.com/news/15672822030222336"
MNX = "Monex 適時開示彙整 6594（2026-09-30：FY2027 公司預想與市場共識）https://scouter.monex.co.jp/news/detail/6594_20260930_act_20260930_163009_1"
FY25 = "公司 FY2025（2025 年 3 月期）決算短信〔IFRS〕（英文版，2025-04-24；分部外部營收，原公告口徑）https://www.nidec.com/v=1777635081/files/topics/33642_ext_3_en_0.pdf"
SAI = "stockanalysis.com Nidec 財報與資產負債表（資料商；頁面 2026-11-24）https://stockanalysis.com/quote/tyo/6594/financials/income-statement/ ；…/financials/balance-sheet/"
SAS = "stockanalysis.com Nidec statistics（資料商；頁面 2026-11-24）https://stockanalysis.com/quote/tyo/6594/statistics/"
Q3S = "公司 FY2026 第 3 四半期決算短信 分部（搜尋結果摘錄：九個月累計各分部營業收入）https://www.nidec.com/files/user/www-nidec-com/corporate/news/2026/0930-09/260930-09en.pdf"
LATE = "公司適時開示 2026-11-16：2027 年 3 月期第 2 四半期決算短信將超過四半期末後 45 日才公布（Minkabu 轉載）https://s.minkabu.jp/news/4666436"
QTR = "Quartr Nidec 摘要（資料商；AI 彙整，日期與期別標示不一致，只作旁證）https://quartr.com/events/nidec-corporation-6594-q2-2026_ojaREW8g"
YCL = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
JGB10 = "日本財務省 國債金利情報 10 年：2026-11-02 2.817%（與 physical_ai/nabtesco In!E90、physical_ai/thk 同一取值，跨模型一致；逾期）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"
L6 = "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx"
L4 = "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx"
FNX = "physical_ai/fanuc/dist/20261123_發那科個股模型_v0.1.1.xlsx"
YKX = "physical_ai/yaskawa/dist/20261123_安川電機個股模型_v0.1.1.xlsx"
NBX = "physical_ai/nabtesco/dist/20261123_Nabtesco個股模型_v0.1.1.xlsx"
THX = "physical_ai/thk/dist/20261123_THK個股模型_v0.1.xlsx"
SMX = "physical_ai/smc/dist/20261123_SMC個股模型_v0.1.xlsx"

# ---------------- In 列 ----------------
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

SEG = [("sp", "精密小型馬達（SPMS）"), ("au", "車載（AMEC）"), ("mo", "家電・商業・產業用 MOEN"), ("ac", "家電・商業・產業用 ACIM"), ("ma", "機器與自動化（Nidec Machinery and Automation）"), ("gc", "集團公司事業（Group Company Business）")]
A25 = (391.265, 348.685, 573.202, 462.269, 213.386, 618.287)
A9M = (295.264, 270.057, 500.686, 325.329, 148.987, 446.034)
sec("A. 分部營收：FY2025 全年（原公告口徑，外部營收）與 FY2026 前九個月（2025-04～12）")
for (k, n), v in zip(SEG, A25):
    inp(f"a25_{k}", f"營收 FY2025：{n}", "¥B", None, v, None, "Verified", FY25 + " 分部外部營收", "原公告口徑；FY2025 其後重編（合併營收 2,607.094 → 2,606.347），PL 頁按重編合併營收等比例縮放（Derived）")
for (k, n), v in zip(SEG, A9M):
    inp(f"a9_{k}", f"營收 FY2026 九個月：{n}", "¥B", None, v, None, "Verified", Q3S, "九個月合計 1,986.357；Q4 分部未取得，PL 頁以九個月占比乘全年合併營收分配（Derived）")
inp("a25_tot", "FY2025 分部合計", "¥B", None, "=" + "+".join(f"a25_{k}" for k, _ in SEG), None, "Derived", "加總", "")
inp("a9_tot", "FY2026 九個月分部合計", "¥B", None, "=" + "+".join(f"a9_{k}" for k, _ in SEG), None, "Derived", "加總", "")

sec("B. FY2027（2027 年 3 月期）錨點：公司預想（2026-09-30 公布；Interested-party）。第 2 季決算短信延後公布，預想未更新")
inp("g27_rev", "公司預想：FY2027 合併營收", "¥B", None, 2800.0, None, "Interested-party", MNX + "：2,800,000 百萬日圓", "前期 2,708.7（+3.4%）；中間期預想 1,400", CO)
inp("g27_op", "公司預想：FY2027 營業利益", "¥B", None, 200.0, None, "Interested-party", MNX + "：200,000", "中間期預想 120；市場共識（2026-09-29）241.4", CO)
inp("g27_pbt", "公司預想：FY2027 稅前利益", "¥B", None, 180.0, None, "Interested-party", MNX + "：180,000", "中間期預想 110", CO)
inp("g27_ni", "公司預想：FY2027 歸屬母公司淨利", "¥B", None, 100.0, None, "Interested-party", MNX + "：100,000", "隱含稅與非支配合計 80（稅前 180 的 44%）；原因未取得（遞延所得稅資產評價等為推測，未查核）", CO)
inp("cs27_rev", "市場共識：FY2027 營收（2026-09-29）", "¥B", None, 2685.52, None, "Interested-party", MNX + "：共識 2,685,520", "只作對照", DV)
inp("cs27_op", "市場共識：FY2027 營業利益（2026-09-29）", "¥B", None, 241.38, None, "Interested-party", MNX + "：共識 241,380", "只作對照", DV)
inp("cs27_ni", "市場共識：FY2027 淨利（2026-09-29）", "¥B", None, 177.0, None, "Interested-party", MNX + "：共識 177,000", "只作對照", DV)
inp("g_anc", "FY2027 實績相對公司預想（營收）", "%", -0.03, 0.0, 0.03, "Assumed", "基準＝公司預想；2025 年起發生會計不正與調查，預想可信度未驗證（" + QTR + "）", "第 2 季決算短信延後（" + LATE + "）；低／高 ±3%")

sec("C. 命題模組一：全公司核心營收成長（六分部同一成長率＋分部傾斜）；FY2027 以公司預想營收為錨")
GTXT = {
    "g27": ("FY2027", {"A": (0.06, "樂觀：共識營收 2,685.5 低於公司預想，樂觀取高於預想；北美資料中心需求強（MOEN，" + FY26 + "）"),
                       "B": (0.03, "基準：公司預想營收 2,800 較 FY2026 2,708.7 成長 3.4%，扣人形子列後核心取 3.0%（Derived）"),
                       "C": (-0.03, "保守：車載與家電減量；FY2024→FY2025 合併營收 +11.1% 為上緣類比，下緣取歷史循環類比（Assumed）")}),
    "g28": ("FY2028", {"A": (0.05, ""), "B": (0.03, "基準：名目成長；FY2023→FY2026 營收年複合 6.7%（Derived，2,230.0 → 2,708.7）高於此"), "C": (-0.02, "保守：循環下行")}),
    "gm": ("FY2029–FY2030", {"A": (0.05, ""), "B": (0.035, "基準：資料中心冷卻與車載電動化長期成長"), "C": (0.01, "")}),
    "glt": ("FY2031–FY2032", {"A": (0.04, ""), "B": (0.03, "長期：名目需求"), "C": (0.02, "")}),
}
GRNG = {"g27": (-0.03, 0.06), "g28": (-0.02, 0.05), "gm": (0.01, 0.05), "glt": (0.02, 0.04)}
for f, (yl, d) in GTXT.items():
    for s in SC:
        v, note = d[s]
        lo, hi = GRNG[f] if s == "B" else (None, None)
        inp(f"{f}_{s}", f"情境 {s}：核心營收年增 {yl}（六分部同一成長率）", "%/年", lo, v, hi, "Assumed", note or "A／C 以 B ± 約 2pt 設定",
            "區間見情境 A／B／C 三列；分部差異另見傾斜列" if s == "B" else "")
TILT = {"sp": -0.02, "au": -0.04, "mo": 0.05, "ac": 0.0, "ma": -0.01, "gc": 0.0}
for (k, n) in SEG:
    inp(f"tl_{k}", f"分部成長傾斜（FY2027–FY2030，加在核心成長率上）：{n}", "pt", TILT[k] - 0.03, TILT[k], TILT[k] + 0.03, "Assumed",
        "MOEN：北美資料中心需求強、發電機與家電商用成長（" + FY26 + "）" if k == "mo" else "車載（AMEC）：歐洲車市疲弱、NPe 減損 1,163 億（" + FY26 + "）" if k == "au" else "未取得分部成長資料；以 0 附近設定",
        "六分部傾斜依 FY2026 前九個月占比加權合計約 +0.35pt（Derived），使基準 FY2027 營收≈公司預想 2,800；低／高為 ±3pt")

sec("C2. 命題模組二：人形機器人用馬達（無框力矩馬達；子列，地區未分）＝線 6 無框力矩馬達市場 × Nidec 份額 × 爬坡 × 匯率")
FMB = {26: 329.8, 30: 1901.6, 32: 4684.2}
inp("fm26", "線 6 無框力矩馬達市場 2026（基準）", "百萬美元", None, 329.8, None, "Derived", L6 + " 計算!C145（五線歸因基準；下游取數頁無馬達列）", "線模型只有基準列；低／高需求以線 6 行星滾柱絲杠低／高需求相對基準比例換算（下列，Analogy）")
inp("fm30", "線 6 無框力矩馬達市場 2030（基準）", "百萬美元", None, 1901.6, None, "Derived", L6 + " 計算!G145", "")
inp("fm32", "線 6 無框力矩馬達市場 2032（基準）", "百萬美元", None, 4684.2, None, "Derived", L6 + " 計算!I145", "")
SCL = {("A", 26): 731.5 / 298.571428571429, ("C", 26): 89.5714285714286 / 298.571428571429, ("A", 30): 6545.958375 / 1558.56151785714, ("C", 30): 233.784227678571 / 1558.56151785714,
       ("A", 32): 13794.2435339844 / 3378.18208995536, ("C", 32): 337.818208995536 / 3378.18208995536}
for (s, yy), v in SCL.items():
    inp(f"scl{yy}{'hi' if s == 'A' else 'lo'}", f"{yy} 年市場相對基準倍數：{'高' if s == 'A' else '低'}需求（線 6 行星滾柱絲杠{'高' if s == 'A' else '低'}需求 ÷ 基準；用於 m{yy}_{s}）", "x", None, v, None, "Analogy",
        L6 + f" 下游取數!{'D' if s == 'A' else 'B'}{ {26: 12, 30: 13, 32: 14}[yy] } ÷ C{ {26: 12, 30: 13, 32: 14}[yy] }", "借用絲杠需求情境比例作無框馬達情境比例（線模型無馬達低／高列），為類比")
for yy in (26, 30, 32):
    for s in SC:
        inp(f"m{yy}_{s}", f"情境 {s}：無框力矩馬達市場 20{yy}", "百萬美元", None, f"=fm{yy}" + ("" if s == "B" else f"*scl{yy}{'hi' if s == 'A' else 'lo'}"), None, "Derived", f"fm{yy}" + ("" if s == "B" else f" × scl{yy}{'hi' if s == 'A' else 'lo'}"), "")
for s, v, lo, hi in (("A", 0.03, None, None), ("B", 0.01, 0.0, 0.03), ("C", 0.002, None, None)):
    inp(f"shr_{s}", f"情境 {s}：Nidec 在人形無框力矩馬達的份額", "%", lo, v, hi, "Assumed",
        "未取得：Nidec 未揭露人形機器人營收或客戶；線 6 份額頁列「無框與空心杯馬達：Maxon、Kollmorgen（RRX）、鳴志」，Nidec 未列（" + L6 + " 份額!B14）", "實體 AI 營收曝險判斷為低（physical_ai/map/profiles/6594.md）；區間 0–3%")
inp("rp27", "人形營收爬坡：FY2027 占穩態比例（線性升到 FY2030＝100%）", "%", 0.0, 0.25, 0.5, "Assumed", "比照 physical_ai/thk、physical_ai/hiwin 爬坡設定", "")
inp("fx", "美元兌日圓（人形市場換算）", "¥/US$", 140.0, 159.707, 165.0, "Verified", YCL + "：JPY=X 2026-11-20 收盤 159.707", "低／高為 Assumed")

sec("D. 損益表結構（標準損益表；IFRS）：毛利率、營業費用率（含減損以外的其他營業損益）")
for s, v, lo, hi in (("A", 0.225, None, None), ("B", 0.215, 0.20, 0.23), ("C", 0.200, None, None)):
    inp(f"gm27_{s}", f"情境 {s}：FY2027 毛利率", "%", lo, v, hi, "Assumed",
        "毛利率：FY2023 18.1%、FY2024 21.2%、FY2025 21.1%（重編）、FY2026 19.4%（Derived，" + SAI + "）", "FY2027 公司預想營業利益率 7.1% 需要毛利率與費用率同時改善；共識營業利益 241.4（8.9%）")
inp("opx27", "FY2027 營業費用率（錨點：基準毛利率 21.5% − 公司預想營業利益率 200 ÷ 2,800）", "%", 0.1336, "=0.215-g27_op/g27_rev", 0.1636, "Derived", "使基準 FY2027 營業利益＝公司預想 200（基準毛利率 21.5% 寫死於本式）", "錨點費用率 14.4% 低於 FY2026 剔除減損後的 16.0%（Derived：毛利 526.6 − 剔除減損營業利益 94.1 ＝ 432.6，÷ 2,708.7）；改善來源未取得，為承重 Assumed；高案＝接近 FY2026 剔除減損水準")
for s, v, lo, hi in (("A", -0.003, None, None), ("B", 0.0, -0.005, 0.015), ("C", 0.010, None, None)):
    inp(f"opxadj_{s}", f"情境 {s}：營業費用率額外調整（FY2027 起）", "pt", lo, v, hi, "Assumed", "C：費用降不下來（調查、訴訟、品質成本，未量化）；A：略優於預想", "")
for s, v, lo, hi in (("A", 0.235, None, None), ("B", 0.225, 0.205, 0.24), ("C", 0.205, None, None)):
    inp(f"gmlt_{s}", f"情境 {s}：FY2030 起毛利率（FY2027 到 FY2030 線性）", "%", lo, v, hi, "Assumed", "FY2024–FY2025 毛利率 21.1–21.2%；車載虧損改善與產品組合（資料中心）", "")
inp("opx_step", "營業費用率每年改善幅度（FY2028–FY2030，三年）", "pt/年", 0.0, 0.002, 0.004, "Assumed", "結構改革（公司中期計畫數字檢討中，" + QTR + "）；每年 0.2pt", "")
inp("fin27", "業外：金融收益 − 金融費用（FY2027；公司預想隱含）", "¥B", None, "=g27_pbt-g27_op", None, "Derived", "＝公司預想稅前 180 − 營業利益 200", "")
inp("nonop", "業外：金融收益 − 金融費用（FY2028 起）", "¥B", -40.0, -20.0, 0.0, "Assumed", "FY2027 預想隱含 −20；借款 1,113 億級、現金 852 億（" + SAI + "）", "")
inp("tax27", "FY2027 有效稅率（公司預想隱含：稅與非支配合計 ÷ 稅前）", "%", None, "=(g27_pbt-g27_ni)/g27_pbt", None, "Derived", "＝(180 − 100) ÷ 180；非支配以 0 計", "FY2026 稅 126.7 在稅前虧損下仍為費用（" + SAI + "），稅負異常原因未查核")
inp("tax_w", "有效稅率（FY2028 起；DCF NOPAT 亦用）", "%", 0.25, 0.30, 0.40, "Assumed", "日本法定實效約 30.6%（類比）", "FY2024 37.8%、FY2025 36.3%（Derived，稅 ÷ 稅前）")

sec("E. 現金流與股東回饋")
inp("fc_da", "折舊攤銷 ÷ 營收（FY2027 起）", "%", 0.025, 0.032, 0.045, "Assumed", "FY2026 62.6 ÷ 2,708.7 ＝ 2.3%（減損後）；FY2025 79.4 ÷ 2,606.3 ＝ 3.0%；FY2024 114.5 ÷ 2,347.2 ＝ 4.9%（" + SAI + "）", "減損後折舊下降，之後隨投資回升")
inp("fc_cx", "資本支出 ÷ 營收（FY2027 起）", "%", 0.035, 0.045, 0.055, "Assumed", "FY2026 130.5 ÷ 2,708.7 ＝ 4.8%（" + SAS + "）", "印度等計畫視現金延後（" + QTR + "）")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.15, 0.25, 0.40, "Assumed", "未取得營運資金明細；以製造業類比", "")
inp("payout", "配息率（FY2028 起；FY2027 為 0）", "%", 0.0, 0.25, 0.40, "Assumed", "公司未決定股利（調查期間，" + QTR + "）；FY2026 年度無股利（" + SAS + "）", "只用於收支關卡（股利＝配息率 × 前一年淨利）；不進 DCF")
inp("eq_26", "FY2026 年底歸屬母公司權益", "¥B", None, 764.726, None, "Verified", SAI + "：Shareholders' equity 764,726", "非支配股東權益 −29.438（負值，不計入）")

sec("F. 估值輸入（評價日＝資產負債表日 2026-03-31；FY2027 第 2 季資產負債表未公布）")
inp("px", "股價（2026-11-20 收盤）", "¥", None, 1644, None, "Verified", YCL + "：6594.T 2026-11-20 收盤 1,644", "未還原股利；11-24 資料商頁面價 1,685（只作對照）")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("sh_v", "評價日股數（流通股）", "百萬股", None, 1146.0, None, "Derived", SAI + "：Shares Outstanding 1,146（FY2026 加權）；" + SAS + "：市值 1.93 兆 ÷ 1,685 ≒ 1,145", "庫藏股與選擇權未取得，視為已含")
inp("cash_v", "現金及約當現金（2026-03-31）", "¥B", None, 852.370, None, "Verified", SAI + "：852,370", "")
inp("debt_v", "有息負債（含租賃負債 69.9）（2026-03-31）", "¥B", None, 1113.404, None, "Verified", SAI + "：Total debt 1,113,404，內含租賃負債 69,885", "租賃已含，下列租賃代理設 0 避免重複")
inp("lease_v", "租賃負債代理（已含於有息負債）", "¥B", 0.0, 0.0, 0.0, "Derived", "已含於 debt_v", "")
inp("inv_v", "其他非流動金融資產", "¥B", None, 0.0, None, "Assumed", "未取得（FY2026 決算新聞稿摘要無此項）", "不計入")
inp("inv_k", "其他非流動金融資產計入比例", "x", 0.5, 1.0, None, "Assumed", "inv_v＝0，無影響", "")
inp("mi_v", "非支配股東權益（2026-03-31）", "¥B", None, 0.0, None, "Assumed", SAI + "：−29,438（負值）；負值不作為加項，取 0", "保守")
inp("stub_frac", "Stub 期（評價日＝資產負債表日，無 Stub）", "年", None, 0.0, None, "Derived", "評價日 2026-03-31；FY2027 起 6 年全部明確期", "")
inp("rf", "無風險利率（日本 10 年期國債）", "%", None, 0.02817, None, "Verified", JGB10, "")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "日本國家風險溢酬", "%", None, 0.0091, None, "Verified", "Damodaran 國家風險溢酬（2026-01-05）：日本 0.91%（與 physical_ai/nabtesco In!E92 同值）", "")
inp("beta", "β（5 年，Blume 調整）", "x", "=E{r}-0.3", "=0.67*1.14+0.33", "=E{r}+0.3", "Derived", SAS + "：β 1.14（5 年）；Blume 0.67β＋0.33", "敏感度 ±0.3（Assumed）；Nidec 近一年因會計調查股價波動特殊，β 可能失真")
inp("kd", "稅前負債成本", "%", None, "=rf+0.01", None, "Assumed", "rf ＋ 1.0pt（特別注意銘柄、會計師無法表示意見，信用利差高於一般日本公司；未取得實際利率）", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("dps_tgt", "評價日到目標時點之每股股利", "¥", None, 0.0, None, "Assumed", "公司未決定股利（調查期間）；視為 0", "")
inp("t_tgt", "評價日 2026-03-31 到目標時點 2027-11-20（年）", "年", None, "=(tp_date-DATE(2026,3,31))/365", None, "Derived", "評價日到目標價時點", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable：同業前瞻 P/E）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺 Nidec 對同業 10 年前瞻 P/E 比值歷史；Nidec 處於調查與特別注意銘柄，折價合理性另見 Risk 頁）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業前瞻 P/E（只納入已換算到 2026-11-20 收盤者；其餘報價日未取得，列於下方「報價日未取得」不進中位數）")
PEERS = [("pe_fanuc", "Fanuc 6954（機器人、FA）", 21.1499, "Derived", THX + " In!pe_fanuc；" + FNX + "（換算至 2026-11-20，用收盤檔 5,047）"),
         ("pe_yask", "安川電機 6506（伺服、機器人）", 15.6739, "Derived", THX + " In!pe_yask；" + YKX + "（換算至 2026-11-20，收盤檔 3,180）"),
         ("pe_nab", "Nabtesco 6268（精密減速機）", 13.7701, "Derived", THX + " In!pe_nab；" + NBX + "（換算至 2026-11-20，收盤檔 3,105）"),
         ("pe_thk", "THK 6481（線性導軌、滾珠絲杠）", 10.6781, "Derived", THX + " In!pe_self（換算至 2026-11-20，收盤檔 4,249）"),
         ("pe_smc", "SMC 6273（氣動元件）", 14.59, "Derived", SMX + " README：現價 55,300（2026-11-20）÷ FY2026 公司指引 EPS 3,789.21 ＝ 14.59")]
for k, lab, v, tag, src in PEERS:
    inp(k, f"前瞻 P/E：{lab}", "x", None, v, None, tag, src, "股權口徑 P/E；分母為資料商賣方共識或公司指引 EPS 固定不變，只換股價", DV)
inp("pe_med", "同業前瞻 P/E 中位數（5 家，皆已換算至 2026-11-20；進 Comparable）", "x", 10.68, "=MEDIAN(pe_fanuc,pe_yask,pe_nab,pe_thk,pe_smc)", 21.15, "Derived",
    "5 家中位數；Nidec 業務（車載、家電、小馬達）與這些 FA 與機械零件同業不同，屬 Analogy", "低＝5 家最低（THK）；高＝5 家最高（Fanuc）。少於約 8 家不做迴歸")
inp("pe_self", "Nidec 現行前瞻 P/E（換算至 2026-11-20）", "x", None, "=11.38*1644/1685", None, "Derived", SAS + "：前瞻 P/E 11.38（1,685 日圓，2026-11-24）× 1,644 ÷ 1,685（分母 EPS 不變）", "只作對照，不進模型；隱含共識 EPS 約 148 日圓（共識淨利 177 ÷ 1,146 股＝154，相近）", DV)
inp("cons_tp", "賣方平均目標價（6 位；資料商 2026-11-24 頁面）", "¥", None, 2353.33, None, "Interested-party", SAS + "：平均 2,353.33、評等 Hold、6 位", "只作對照，不進模型", DV)


def build(path, ov=None, sens=None, rev=None, bid="{BID}"):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 ¥B、股數百萬股、EPS 與股價日圓；FY2026＝2026 年 3 月期"
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
           ("m26", "無框力矩馬達市場 2026（線 6）"), ("m30", "無框力矩馬達市場 2030（線 6）"), ("m32", "無框力矩馬達市場 2032（線 6）"),
           ("shr", "Nidec 人形無框馬達份額"), ("gm27", "FY2027 毛利率"), ("gmlt", "FY2030 起毛利率"), ("opxadj", "營業費用率調整")]
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
    H["A1"] = "Hist — 歷史財務（年度；¥B；3 月期）。FY2025 為重編值（2026-09-30 新聞稿），FY2024、FY2023 為原公告值，三年口徑不完全可比；FY2026 含巨額減損。公式只引用本頁（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025（重編）", "FY2026", "標記", "來源"])
    NA = "未取得"
    HR = [("h_rev", "營收", 2230.027, 2347.159, 2606.347, 2708.705, "Verified", SAI + "；" + FY26),
          ("h_gp", "毛利", 403.051, 497.162, 548.627, 526.635, "Verified", SAI),
          ("h_cogs", "營業成本", "=B3-B4", "=C3-C4", "=D3-D4", "=E3-E4", "Derived", "營收 − 毛利"),
          ("h_op", "營業利益（IFRS 報告值；FY2026 含減損）", 89.923, 161.856, 128.222, -518.978, "Verified", FY25 + "（FY2024）；" + FY26 + "（FY2025 重編、FY2026）；FY2023：" + SAI),
          ("h_opex", "營業費用（毛利 − 營業利益；含減損與其他營業損益）", "=B4-B6", "=C4-C6", "=D4-D6", "=E4-E6", "Derived", "FY2026 含資產減損 613.0（" + SAI + " Asset Writedown）"),
          ("h_pbt", "稅前利益", 110.435, 201.669, 124.576, -496.214, "Verified", SAI + "；" + FY26),
          ("h_fin", "業外：稅前 − 營業利益", "=B8-B6", "=C8-C6", "=D8-D6", "=E8-E6", "Derived", "稅前利益 − 營業利益"),
          ("h_tax", "所得稅", 72.752, 76.274, 45.167, 126.696, "Verified", SAI + "（FY2026 稅前虧損仍有稅費，原因未查核）"),
          ("h_ni", "歸屬母公司淨利", 36.982, 124.455, 84.672, -564.624, "Verified", SAI + "；" + FY26),
          ("h_nci", "非支配股東損益（＝稅後利益 − 歸屬母公司淨利）", "=B8-B10-B11", "=C8-C10-C11", "=D8-D10-D11", "=E8-E10-E11", "Derived", "稅前 − 稅 − 歸屬母公司淨利（FY2026 為負：非支配股東分擔減損）"),
          ("h_eps", "基本 EPS（日圓）", 32.13, 108.30, 73.69, -492.55, "Verified", SAI + "；" + FY26),
          ("h_da", "折舊攤銷", 119.522, 114.481, 79.350, 62.591, "Verified", SAI),
          ("h_cx", "資本支出", NA, NA, NA, 130.48, "Verified", "stockanalysis.com Nidec statistics（資料商，頁面 2026-11-24；網址見 Sources 頁）Capital Expenditures，FY2026／TTM；其他年度未取得")]
    for i, (k, lab, a, b, c, d, tag, src) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b, c, d)):
            cc = H.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): cc.font = BLUE
        H.cell(i, 6, tag); H.cell(i, 7, src)
        for yy, cl in (("23", "B"), ("24", "C"), ("25", "D"), ("26", "E")): dn(f"{k}_{yy}", f"Hist!${cl}${i}")
    H.column_dimensions["A"].width = 48

    # ---------------- Q（官方分部） ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 官方報告分部（外部顧客營收，¥B）。FY2024、FY2025 為全年（原公告口徑）；9M'26＝2025-04 至 2025-12 累計。分部對帳：分部合計 − 合併營收"
    hdr(Q, 2, ["項目", "FY2024", "FY2025（原公告）", "9M'26", "標記", "來源", "來源誘因"])
    FY24S = (330.596, 338.540, 459.828, 428.009, 204.388, 585.798)
    for i, ((k, n), a, b, c) in enumerate(zip(SEG, FY24S, A25, A9M), 3):
        Q.cell(i, 1, n + " 營收")
        for j, v in enumerate((a, b, c)): Q.cell(i, 2 + j, v).font = BLUE
        Q.cell(i, 5, "Verified"); Q.cell(i, 6, FY25 + "（FY2024、FY2025）；" + Q3S + "（9M'26）"); Q.cell(i, 7, CO)
    Q.cell(9, 1, "報告分部合計"); Q.cell(9, 5, "Derived")
    for j, L in enumerate("BCD"): Q.cell(9, 2 + j, f"=SUM({L}3:{L}8)")
    Q.cell(10, 1, "合併營收（FY2025 原公告 2,607.094；9M'26 九個月合計）"); Q.cell(10, 5, "Verified")
    for j, v in enumerate((2347.159, 2607.094, 1986.357)): Q.cell(10, 2 + j, v).font = BLUE
    Q.cell(10, 6, FY25 + "；" + Q3S)
    Q.cell(11, 1, "檢查：分部合計 − 合併營收（四捨五入差 < 0.005）"); Q.cell(11, 5, "Derived")
    for j, L in enumerate("BCD"): Q.cell(11, 2 + j, f"=ROUND({L}9-{L}10,3)")
    Q.cell(12, 1, "註：官方分部名稱沿用公司縮寫（SPMS＝精密小型馬達、AMEC＝車載、MOEN 與 ACIM＝家電商用產業用）；FY2025 其後重編（合併營收 2,606.347），PL 頁以重編合併營收等比例縮放分部；FY2026 全年分部未取得，以前九個月占比分配（Derived）")
    Q.column_dimensions["A"].width = 52

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（核心營收＋人形無框馬達子列）→ 分部營收 → 標準損益表（IFRS）→ 現金流（¥B；FY2023–FY2026 實績，FY2027 以公司預想為錨，FY2028 起模型）。FY2026＝2026 年 3 月期"
    PR = {}
    rr = 3
    KEYS = ([f"c_{k}" for k, _ in SEG] + ["core", "hum", "rev", "chk1", "chk2", "gr", "cogs", "gp", "gm", "opex", "op", "opm", "fin", "pbt", "tax", "cont", "disc",
             "nci", "ni", "nim", "sh", "eps", "epsc", "nig", "da", "cx", "dwc", "fcff", "eq", "div", "fcfad", "mkt", "rp"])
    LAB = ([f"營收：{n}" for _, n in SEG] +
           ["核心營收（六分部合計）", "人形機器人用無框力矩馬達（命題子列，分部未分）", "營收（合併）", "檢查：分部合計 − 核心營收", "檢查：核心＋人形 − 合併營收",
            "營收年增率", "營業成本", "毛利", "毛利率", "營業費用（含減損與其他營業損益）", "營業利益", "營業利益率", "業外：稅前 − 營業利益（金融收支與其他）", "稅前利益", "所得稅",
            "稅後利益（含非支配）", "停業單位損益（無）", "非支配股東損益", "稅後淨利（歸屬母公司）", "淨利率", "平均股數（百萬股）", "EPS（基本，日圓）", "EPS（持續營業，日圓）", "淨利年增率",
            "折舊攤銷", "資本支出", "營運資金增加", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "歸屬母公司權益（年底）", "股利", "自由現金流扣股利（收支關卡用）",
            "參考：無框力矩馬達市場（百萬美元，線 6）", "參考：人形營收爬坡（占穩態）"])
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, k in enumerate(KEYS)}
        for i, l in enumerate(LAB): P.cell(b + i, 1, l)
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1); cell = lambda k: f"{c}{R[k]}"; prev = lambda k: f"{p}{R[k]}"
            def w(k, f): P[cell(k)] = f
            if y <= 2026:
                yy = str(y)[2:]
                for k, hk in (("rev", "rev"), ("cogs", "cogs"), ("gp", "gp"), ("opex", "opex"), ("op", "op"), ("fin", "fin"), ("tax", "tax"), ("nci", "nci"), ("ni", "ni"), ("eps", "eps"), ("da", "da")):
                    w(k, f"=h_{hk}_{yy}")
                w("pbt", f"=h_pbt_{yy}"); w("cont", f"={cell('pbt')}-{cell('tax')}"); w("disc", 0); w("epsc", f"={cell('eps')}")
                w("gm", f"={cell('gp')}/{cell('rev')}"); w("opm", f"={cell('op')}/{cell('rev')}"); w("nim", f"={cell('ni')}/{cell('rev')}")
                w("sh", f"={cell('ni')}/{cell('eps')}*1000"); w("hum", 0); w("rp", 0)
                if y == 2026: w("cx", f"=h_cx_{yy}")
                else: w("cx", '="未取得"')
                if y >= 2025:
                    for k, _ in SEG:
                        w(f"c_{k}", f"=a25_{k}*{cell('rev')}/a25_tot" if y == 2025 else f"=a9_{k}*{cell('rev')}/a9_tot")
                    w("core", f"=SUM({c}{R['c_sp']}:{c}{R['c_gc']})")
                    w("chk1", f"=ROUND(SUM({c}{R['c_sp']}:{c}{R['c_gc']})-{cell('core')},6)")
                    w("chk2", f"=ROUND({cell('core')}+{cell('hum')}-{cell('rev')},6)")
                else:
                    w("core", '="未取得"'); w("chk1", '="—"'); w("chk2", '="—"')
                w("gr", '="未取得"' if y == 2023 else f"={cell('rev')}/{prev('rev')}-1")
                w("nig", '="未取得"' if y == 2023 else f'=IF({prev("ni")}>0,{cell("ni")}/{prev("ni")}-1,"N/M（前一年為虧損）")' if y > 2023 else "")
                if y == 2026: w("nig", '="N/M（FY2026 為虧損）"'); w("eq", "=eq_26")
                continue
            gk = "g27" if y == 2027 else "g28" if y == 2028 else "gm" if y <= 2030 else "glt"
            tl = y <= 2030
            for k, _ in SEG:
                if y == 2027:
                    w(f"c_{k}", f"={prev('c_' + k)}*(1+{gk}_{s}+tl_{k})*(1+g_anc)")
                else:
                    w(f"c_{k}", f"={prev('c_' + k)}*(1+{gk}_{s}" + (f"+tl_{k}" if tl else "") + ")")
            w("core", f"=SUM({c}{R['c_sp']}:{c}{R['c_gc']})")
            xi = f"MIN(1,({y}-2027)/3)"
            w("mkt", f"=m26_{s}*(m30_{s}/m26_{s})^(({y}-2027)/4)" if y <= 2030 else f"=m30_{s}" if y == 2031 else f"=m30_{s}*(m32_{s}/m30_{s})^0.5")
            w("rp", f"=MIN(1,rp27+(1-rp27)*({y}-2027)/3)")
            w("hum", f"={cell('mkt')}*shr_{s}*{cell('rp')}*fx/1000")
            w("rev", f"={cell('core')}+{cell('hum')}")
            w("chk1", f"=ROUND(SUM({c}{R['c_sp']}:{c}{R['c_gc']})-{cell('core')},6)")
            w("chk2", f"=ROUND({cell('core')}+{cell('hum')}-{cell('rev')},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            w("gm", f"=gm27_{s}+(gmlt_{s}-gm27_{s})*{xi}")
            w("gp", f"={cell('rev')}*{cell('gm')}")
            w("cogs", f"={cell('rev')}-{cell('gp')}")
            opx = f"((opx27+opxadj_{s})-opx_step*MIN(3,{y}-2027))"
            w("opex", f"={cell('rev')}*{opx}")
            w("op", f"={cell('gp')}-{cell('opex')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            w("fin", "=fin27" if y == 2027 else "=nonop")
            w("pbt", f"={cell('op')}+{cell('fin')}")
            w("tax", f"={cell('pbt')}*" + ("tax27" if y == 2027 else "tax_w"))
            w("cont", f"={cell('pbt')}-{cell('tax')}"); w("disc", 0); w("nci", 0)
            w("ni", f"={cell('cont')}+{cell('disc')}-{cell('nci')}")
            w("nim", f"={cell('ni')}/{cell('rev')}"); w("sh", "=sh_v")
            w("eps", f"={cell('ni')}/{cell('sh')}*1000"); w("epsc", f"=({cell('cont')}-{cell('nci')})/{cell('sh')}*1000")
            w("nig", f"=IF({prev('ni')}>0,{cell('ni')}/{prev('ni')}-1,\"N/M（前一年為虧損）\")")
            w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}")
            w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
            w("fcff", f"={cell('op')}*(1-tax_w)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("div", "=0" if y == 2027 else f"=payout*MAX(0,{prev('ni')})")
            w("eq", f"={prev('eq')}+{cell('ni')}-{cell('div')}")
            w("fcfad", f"={cell('fcff')}-{cell('div')}")
        P.cell(R["c_sp"], 2, "FY2027 起＝前一年 ×(1＋核心成長＋分部傾斜)（FY2027 另乘 (1＋g_anc)）")
        P.cell(R["hum"], 2, "市場 × 份額 × 爬坡 × 匯率；FY2026 以前未揭露（視為 0）；日曆年市場對應 3 月期（FY2027≈CY2026）")
        P.cell(R["gm"], 2, "FY2027 In gm27 → FY2030 In gmlt 線性")
        P.cell(R["opex"], 2, "營業費用率＝基準毛利率 − 公司預想營業利益率 ＋ 調整 − 每年改善；FY2026 以前為報告值")
        P.cell(R["tax"], 2, "FY2027＝稅前 × 公司預想隱含稅率；FY2028 起 tax_w")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 ×(1 − tax_w)；FY2026 以前不算")
        P.cell(R["div"], 2, "FY2027＝0；之後＝配息率 × 前一年淨利（正值部分）")
        P.cell(R["c_sp"] + 6, 2, "核心＝六分部合計")
        for k in ("gr", "gm", "opm", "nim", "nig", "rp"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(LAB) + 2
    P.column_dimensions["A"].width = 52; P.column_dimensions["B"].width = 30
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "fcfad", "hum"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-03-31（資產負債表日），明確期 FY2027–FY2032（3 月期），期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（¥B）"; D["B4"] = "=px*sh_v/1000"
    D["A5"] = "負債 D（社債及借款＋租賃代理）"; D["B5"] = "=debt_v+lease_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金與投資（2026-03-31）＝現金 − 借款 − 租賃代理 − 非支配 ＋ 其他非流動金融資產 × inv_k"; D["B10"] = "=cash_v-debt_v-lease_v-mi_v+inv_k*inv_v"; dn("dcf_nc", "DCF!$B$10")
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
                 ("＋淨現金與投資（2026-03-31）", "=dcf_nc"),
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
    r_cu = vrow("Comparable（未調整）：同業前瞻 P/E 中位數 × 情境目標時點起算 NTM EPS（FY2028 ×1/3＋FY2029 ×2/3）", [f"=pe_med*(PL!H{PR[s]['epsc']}/3+PL!I{PR[s]['epsc']}*2/3)" for s in SC],
                "目標時點 2027-11-20 起算 NTM（3 月期）＝約 4 個月 FY2028＋8 個月 FY2029；分母：模型各情境對應 EPS 加權")
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
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2027–FY2031 自由現金流扣股利合計（¥B；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!G{PR[s]['fcfad']}:PL!K{PR[s]['fcfad']})")
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
    r_pd1 = one("期間：模型折現（評價日 2026-03-31 → 目標時點 2027-11-20，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B 目標時點起算 NTM EPS", f"=B{r_tp}/(PL!H{PR['B']['epsc']}/3+PL!I{PR['B']['epsc']}*2/3)", "分母：模型情境 B NTM EPS（FY2028 ×1/3＋FY2029 ×2/3）")
    r_pe1 = one("現價 ÷ 情境 B FY2027 持續營業 EPS", f"=px/PL!G{PR['B']['epsc']}", "分母：模型情境 B FY2027")
    r_pe2 = one("現行前瞻 P/E（資料商，換算至 2026-11-20）÷ 同業中位數", "=pe_self/pe_med", "只作對照；相對重估需要的倍數＝1 ÷ 本值")
    r_hum = one("人形無框馬達營收占 FY2030 合併營收（情境 B）", f"=PL!J{PR['B']['hum']}/PL!J{PR['B']['rev']}")
    for rr_ in (r_up, r_a1, r_a2, r_a3, r_hum): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "#,##0"
    V.column_dimensions["A"].width = 70; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, ca=r_ca, pd2=r_pd2, pe=r_pe, pe1=r_pe1, pe2=r_pe2, gj=r_gj, cons=r_cons, wk=r_wk, hum=r_hum)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司預想、市場共識對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或共識", "差", "說明"])
    rc = [("FY2027 營收", f"=PL!G{PR['B']['rev']}", "=g27_rev", "公司預想（Interested-party）；模型含人形子列"),
          ("FY2027 營業利益", f"=PL!G{PR['B']['op']}", "=g27_op", "公司預想；基準由營業費用率錨定使之相等"),
          ("FY2027 歸屬母公司淨利", f"=PL!G{PR['B']['ni']}", "=g27_ni", "公司預想（稅率取預想隱含值）"),
          ("FY2027 營業利益 vs 市場共識（2026-09-29）", f"=PL!G{PR['B']['op']}", "=cs27_op", "共識 241.4 高於公司預想 200"),
          ("FY2027 淨利 vs 市場共識（2026-09-29）", f"=PL!G{PR['B']['ni']}", "=cs27_ni", "共識淨利 177 高於公司預想 100（稅率與非支配假設差異）"),
          ("FY2027 EPS（模型 B） vs 現價 ÷ 現行前瞻 P/E 隱含 EPS", f"=PL!G{PR['B']['epsc']}", "=px/pe_self", "口徑：資料商前瞻 P/E 的分母（賣方共識）vs 模型 FY2027；口徑不同，只看方向"),
          ("加權目標價 vs 賣方平均目標價", "=tp", "=cons_tp", "賣方 6 位平均（2026-11-24 頁面，只作對照）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_nidec.py 批次格點＋線性內插；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 60; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_nidec.py 以 LibreOffice 批次重算每一案例）；結果為常數快照"
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
    RE["A1"] = "對帳 — 本模型使用的線層級變數（讀線模型成品；與線模型不一致的自設值標明理由，不靜默改口徑）"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型不一致？"])
    rec = [("無框力矩馬達市場 2026（基準）", "6", "329.8", "百萬美元", "五線歸因基準；下游取數頁無馬達列，取計算頁", L6 + " 計算!C145", "Derived", "In fm26 → m26_* → PL 人形子列", "一致（照取）"),
           ("無框力矩馬達市場 2030（基準）", "6", "1,901.6", "百萬美元", "同上", L6 + " 計算!G145", "Derived", "In fm30", "一致"),
           ("無框力矩馬達市場 2032（基準）", "6", "4,684.2", "百萬美元", "同上", L6 + " 計算!I145", "Derived", "In fm32", "一致"),
           ("無框馬達低／高需求倍數 2026／2030／2032", "6", "A：2.45／4.20／4.08；C：0.30／0.15／0.10", "x", "借用行星滾柱絲杠低／高需求對基準的比例", L6 + " 下游取數!B12:D14", "Analogy", "In scl26hi／lo、scl30hi／lo、scl32hi／lo", "自設（線模型無馬達低／高列）：以絲杠需求情境比例類比"),
           ("有框伺服馬達市場 2030", "6", "509.2", "百萬美元", "需求來自傳統工業機器人與 AMR，與人形關係小", L6 + " 計算!G144", "Derived", "未進公式", "不使用：Nidec 營收主體為車載、家電、小馬達，機器人用有框伺服占比未揭露"),
           ("伺服驅動、編碼器市場 2030", "6", "1,432.4／558.5", "百萬美元", "五線歸因", L6 + " 計算!G147、G148", "Derived", "未進公式", "不使用：Nidec 非主要驅動器或編碼器供應商（線 6 份額頁未列）"),
           ("無框與空心杯馬達領先廠", "6", "Maxon、Kollmorgen（RRX）、鳴志；Nidec 未列", "—", "定性", L6 + " 份額!B14", "Interested-party", "shr_* 說明（基準 1%）", "一致"),
           ("人形全球出貨 2030（低／基準／高）", "4", "20 萬／50 萬／89 萬", "台", "全球、含小型機", L4 + " 下游取數!C38:E38", "Derived", "未進公式（線 6 市場已含）", "一致（經線 6）"),
           ("人形 BOM 馬達類占比（中值）", "4", "約 15%（範圍 10–20%）", "%", "人形 BOM", "physical_ai/map/profiles/6594.md（segments.csv MOT，Analogy）", "Analogy", "說明用", "未使用")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 30; RE.column_dimensions["F"].width = 60; RE.column_dimensions["I"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("會計不正與品質問題調查、財報再重編（FY2025 已重編；會計師對有價證券報告書出具無法表示意見）", "部分（g_anc、opxadj_C）", "預想可信度未驗證；再重編幅度無法估計", "第 2 季決算短信（已延後）、調查結論、會計師意見"),
             ("特別注意銘柄（2025-10 起）與可能下市風險", "否", "無法量化機率；Comparable 未折價（rel_k＝1）", "東證公告、解除條件"),
             ("車載事業再減損與合約損失（FY2026 已認列 ACIM 2,988 億、NPe 1,163 億、NMOJ 731 億）", "部分（情境 C 毛利率與費用率）", "未取得剩餘資產帳面值", "車載分部營業損益、合約損失準備"),
             ("訴訟、客戶與供應商求償", "否", "未取得金額", "決算短信注記"),
             ("股利與買回：公司未決定股利", "是（payout）", "FY2027 設 0，FY2028 起 25%（區間 0–40%）", "股利公告"),
             ("資料中心需求與北美市場（MOEN 成長依賴）", "是（tl_mo）", "分部傾斜 +5pt（區間 +2～+8pt）", "分部營收、受注"),
             ("匯率", "否", "未逐年建模", "USDJPY"),
             ("人形份額為零或採用其他廠方案", "是（shr_C、Sens）", "人形子列占營收比例小（見 Val 頁）", "公司人形產品與客戶揭露")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 50

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([FY26, MNX, FY25, SAI, SAS, Q3S, LATE, QTR, YCL, JGB10, L6 + "（計算、份額）", L4 + "（下游取數）",
                            THX + "（In 頁同業 P/E）", FNX, YKX, NBX, SMX], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"Nidec 日本電產（6594.T）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-24（FY2027 第 2 季決算短信尚未公布）；成品日 2026-11-24；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"#,##0")&" 日圓）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：Nidec 是全品類馬達龍頭，股價取決於車載事業減損與會計調查之後，FY2027 公司預想的獲利回升（營業利益 2,000 億日圓）能否兌現；人形機器人用馬達只是小型選擇權，不是承重項"
    RD["A4"] = (f'="在 FY2027 核心營收情境 B 年增 "&TEXT(g27_B,"0%")&"、FY2030 起毛利率 "&TEXT(gmlt_B,"0.0%")&"、同業前瞻 P/E 中位數 "&TEXT(pe_med,"0.0")&" 倍的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"#,##0")&" 日圓（DCF "&TEXT(Val!B{r_dcft},"#,##0")&"、Comparable "&TEXT(Val!B{r_cmpt},"#,##0")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2027–31 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0.0")&" 十億日圓）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 日圓。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：第 2 季決算短信顯示上半年營業利益明顯低於中間期預想 1,200 億日圓；會計調查再度重編或會計師持續無法表示意見；車載分部再度減損；FY2027 公司預想下修"
    RD["A7"] = "範圍：只建合併模型（六分部營收由成長率與傾斜推導，利潤率為合併層級）；不讀 Siliconomics／Tokenomics／Wattonomics（算力線）；自用 AI 讀法等規格條文不適用。HTML 展示層見 dist/ 同名 .html"
    RD["A8"] = f'="命題回覆：問：FY2027 預想的獲利回升能否支持高於現價的目標價？答：情境 B FY2028 EPS "&TEXT(PL!H{PR["B"]["epsc"]},"#,##0")&" 日圓，現價為其 "&TEXT(px/PL!H{PR["B"]["epsc"]},"0.0")&" 倍，同業中位數 "&TEXT(pe_med,"0.0")&" 倍；加權目標價 "&TEXT(Val!B{r_tp},"#,##0")&" 日圓。問：人形機器人有多重要？答：情境 B FY2030 人形子列只占合併營收 "&TEXT(Val!B{r_hum},"0.0%")&"，結論不靠人形（模型輸出，非獨立證據）。"'
    wb.save(path)
    return VR, PR, DR



# ---------------- 重算與讀值 ----------------
def lo_profile():
    d = os.path.join(os.environ.get("HOME", tempfile.gettempdir()), "lo_nidec")
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
    d = tempfile.mkdtemp(prefix="nidec_"); paths = []; VR = DR = None
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


CASES = [  # (鍵, 驅動白話)
    ("g_anc", "FY2027 營收相對公司預想"), ("g27_B", "核心營收年增 FY2027（情境 B）"), ("g28_B", "核心營收年增 FY2028（情境 B）"), ("gm_B", "核心營收年增 FY2029–30（情境 B）"),
    ("glt_B", "核心營收年增 FY2031–32（情境 B）"), ("tl_mo", "MOEN 成長傾斜"), ("tl_au", "車載成長傾斜"), ("tl_sp", "精密小馬達成長傾斜"), ("tl_ac", "ACIM 成長傾斜"), ("tl_ma", "機器與自動化成長傾斜"), ("tl_gc", "集團公司事業成長傾斜"), ("inv_k", "其他金融資產計入比例"), ("shr_B", "人形無框馬達份額（情境 B）"), ("rp27", "人形營收爬坡 FY2027"), ("fx", "美元兌日圓"),
    ("gm27_B", "FY2027 毛利率（情境 B）"), ("gmlt_B", "FY2030 起毛利率（情境 B）"), ("opxadj_B", "營業費用率調整（情境 B）"),("opx27", "FY2027 營業費用率錨點"), ("opx_step", "營業費用率每年改善"), ("nonop", "業外金融損益"), ("tax_w", "有效稅率"),
    ("fc_da", "折舊 ÷ 營收"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("payout", "配息率"),
    ("beta", "β"), ("erp", "股權風險溢酬"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重"), ("pe_med", "同業前瞻 P/E 中位數")]
BETA_B = 0.67 * 1.14 + 0.33
FXD = {"pe_med": statistics.median([21.1499, 15.6739, 13.7701, 10.6781, 14.59])}



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
    for key, b0, lo_, hi_, note in (("g27_B", 0.03, -0.40, 0.60, "情境 B 的 FY2027 核心營收年增（其他不變，整本重算；情境 A／C 不動；之後年度在此基礎上成長）"),
                                    ("opx27", 0.143571, 0.0, 0.40, "FY2027 營業費用率錨點（三情境同動）"),
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
    tmp = tempfile.mkdtemp(prefix="nidf_"); raw = os.path.join(tmp, NAME)
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
    json.dump(dict(company="nidec", ticker="6594.T", model_sheet="Q",
                   official_segments=[dict(name=n, label=n + " 營收", period="9M'26（2025-04～2025-12 累計，外部顧客）", value=v, unit="¥B", source=Q3S)
                                      for (k, n), v in zip(SEG, A9M)],
                   sublines=[], check_labels=["檢查：分部合計 − 合併營收（四捨五入差 < 0.005）"],
                   note="官方報告分部為公司事業單位（SPMS、AMEC、MOEN、ACIM、機器與自動化、集團公司事業）；Q 頁以分部合計對合併營收，PL 頁以重編合併營收等比例縮放分部，另有「檢查：分部合計 − 核心營收」「檢查：核心＋人形 − 合併營收」。"),
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
    seg_rows = "\n".join(f"| {n} | {yr('B', 'c_' + k, ys=range(2025, 2033))} |" for k, n in SEG)
    txt = f"""# physical_ai/nidec — Nidec 日本電產（6594.T）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

東證 Prime 上市（2025-10 起列特別注意銘柄）。3 月底決算；本模型稱「FY2026＝2026 年 3 月期」。本 README 的數字由 `build_nidec.py` 從重算後的成品讀出產生，不手寫。

**未完成與限制（先讀）：** (1) FY2027 第 2 季決算短信因調查延後公布（2026-11-16 適時開示），最新資產負債表為 2026-03-31，公司預想為 2026-09-30 版；(2) 公司 FY2026 全年分部營收未取得，以前九個月占比分配；(3) 合併為模型最小完整版：分部只推營收，利潤率在合併層級；(4) 其他非流動金融資產、庫藏股、租賃負債拆分未取得；(5) 同業 5 家已換算至 2026-11-20，另有 NSK、IKO、直得、上銀等報價日未取得，未納入中位數；(6) 會計不正調查結果與再重編風險無法量化。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A8'].value}

**讀法：** 公司 FY2027 預想營業利益 2,000 億日圓（FY2026 剔除減損約 941 億），現價 {f0(inv('px'))} 日圓只有模型 FY2027 EPS 的 {g(f"B{VR['pe1']}"):.1f} 倍。結論主要建立在 Assumed 輸入（FY2027 起毛利率與營業費用率改善、核心成長、同業倍數）與公司預想（Interested-party）上，沒有獨立證據證明費用率能從 FY2026 的 16.0%（剔除減損）降到 14.4%；人形子列（情境 B FY2030 占營收 {g(f"B{VR['hum']}")*100:.2f}%）不是承重項。

## 估值（評價日 2026-03-31；目標時點 2027-11-20）
| 情境 | DCF（推到目標時點） | Comparable（P/E × NTM EPS） | 情境目標價 | 收支關卡（FY2027–31 FCF 扣股利，¥B） |
|---|---|---|---|---|
{scn}

- 加權目標價 **{f0(tp)} 日圓**（DCF {f0(dcf)}、Comparable {f0(cmp_)}），現價 {f0(inv('px'))}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")} → 投資建議 **{g(f"B{VR['rec']}")}**。賣方平均目標價 {f0(inv('cons_tp'))}（6 位，2026-11-24，只作對照）。
- 方法：DCF 標準版 FCFF（{g(f"B{VR['wk']}")}，為 WACC／ke／g／RONIC）50%＋Comparable 50%（同業 5 家前瞻 P/E 中位數 {inv('pe_med'):.2f} 倍，皆已換算到 2026-11-20，× 各情境目標時點起算 NTM EPS）。相對倍數係數 1.00（未調整）。
- 隱含前瞻 P/E {g(f"B{VR['pe']}"):.1f} 倍；Nidec 現行前瞻 P/E {inv('pe_self'):.2f} 倍 ＝ 同業中位數的 {g(f"B{VR['pe2']}"):.2f} 倍。

## 營收與標準損益（情境 B，¥B；FY2025–26 實績、FY2027 起模型）
| 列 | {" | ".join(f"FY{y}" for y in range(2025, 2033))} |
|---|{"---|" * 8}
{seg_rows}
| 人形無框馬達（命題子列） | {yr('B','hum',ys=range(2025,2033))} |
| 營收（合併） | {yr('B','rev',ys=range(2025,2033))} |
| 毛利 | {yr('B','gp',ys=range(2025,2033))} |
| 營業利益 | {yr('B','op',ys=range(2025,2033))} |
| 歸屬母公司淨利 | {yr('B','ni',ys=range(2025,2033))} |
| EPS（日圓） | {yr('B','epsc','{:,.0f}',ys=range(2025,2033))} |

- FY2026 營業損失 5,190 億含資產減損 6,130 億（ACIM 2,988 億、NPe 1,163 億、NMOJ 731 億），歸屬母公司淨損 5,646 億。FY2025 為 2026-09-30 重編值（原公告營業利益 2,402 億、淨利 1,677 億），FY2023–24 為原公告值，三年不可比。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 核心營收年增（A／B／C）：FY2027 6%／3%／−3%；FY2028 5%／3%／−2%；FY2029–30 5%／3.5%／1%；FY2031–32 4%／3%／2%；另有六分部傾斜（MOEN +5pt、車載 −4pt 等）。
- 毛利率 FY2027 22.5%／21.5%／20.0%，FY2030 起 23.5%／22.5%／20.5%；營業費用率錨點＝基準毛利率 − 公司預想營業利益率（14.4%），每年再改善 0.2pt；稅率 FY2027 取預想隱含值、FY2028 起 30%。
- 人形：線 6 無框力矩馬達市場（基準 2030 年 1,901.6 百萬美元）× Nidec 份額 3%／1%／0.2% × 爬坡 × 159.7 日圓；低／高需求倍數借用絲杠情境比例（Analogy）。
- 現金流：折舊 3.2%、資本支出 4.5%、營運資金 25% × 營收增量；配息率 FY2028 起 25%（FY2027 為 0）。
- 折現：rf 2.817%、β {inv('beta'):.3f}、ERP 5%、CRP 0.91%、kd＝rf＋1pt、g＝MIN(3%, rf)、RONIC＝WACC。股數 1,146 百萬；淨負債約 {-wb['DCF']['B10'].value:,.1f}（現金 852.4、有息負債含租賃 1,113.4）。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| FY2027–30 毛利率與營業費用率（gm27、gmlt、opx） | 第 2 季決算短信營業利益對中間期預想 1,200 億；季度毛利率；車載分部損益 |
| FY2027 核心成長（g27、tl_mo） | MOEN 分部營收、北美資料中心需求；公司預想修正 |
| 同業 P/E（pe_med） | 同業前瞻 P/E 換算到同一估值日後的中位數 |

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
- 線 6：無框力矩馬達市場 2026／2030／2032 基準（計算!C145、G145、I145；下游取數頁無馬達列）照取；低／高需求倍數借用絲杠情境比例（Analogy，自設）；有框伺服、伺服驅動、編碼器不使用（Nidec 營收主體非此）。
- 線 4：人形出貨與 BOM 馬達占比只作說明，未進公式。

## 限制與未取得資料
- 公司人形機器人營收、客戶、份額未取得；分部全年（FY2026）、分部營業損益 FY2026、FY2027 第 2 季數字未取得。
- FY2027 稅與非支配合計 80（稅前 180 的 44%）的原因未查核；FY2026 稅前虧損下仍有稅費 126.7 的原因未查核。
- 同業前瞻 P/E 5 家（Fanuc、安川、Nabtesco、THK、SMC）已換算到 2026-11-20；NSK、Nippon Thompson、直得、上銀、哈默納科、Regal Rexnord：報價日未取得（或與業務不可比），未納入。
- rf、CRP 沿用 2026-11-02／01-05 取值（逾期，跨模型一致）。

## 檔案
- `build_nidec.py`：建置＋整本重算 Sens／反解（LibreOffice 批次）＋產生本 README、`segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-24）：初版；現價 {f0(inv('px'))}（2026-11-20 收盤）；加權目標價 {f0(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


if __name__ == "__main__":
    main()

# Novanta（NOVT）個股模型 v0.1 — AI實體應用 Project（線 6 共用零組件：六維力與扭矩感測；並讀線 4 人形）
# 用法：python3 physical_ai/novanta/build_novanta.py → 建置、整本重算 Sens 與反解（LibreOffice 批次）、輸出 dist/ 成品、README.md、sens.json、segments.json
# 會計年度＝曆年（52／53 週制，近似曆年）。金額：百萬美元（US$M）；股數：百萬股；EPS、股價：美元。美國 GAAP。
# 兩個報告分部：Automation Enabling Technologies（AET，含 ATI 力感測）與 Medical Solutions（2026-07 併入 Riverpoint Medical）。
# 結構：README、In、Hist、Q、PL（三情境）、DCF、Val、Recon、Sens、對帳、Risk、Sources。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261124"
NAME = f"{DATE}_Novanta個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032

# ---------------- 出處 ----------------
Q3R = "公司 2026 年第 3 季財報新聞稿（Business Wire 經 Finviz 轉載，2026-11-10）https://finviz.com/news/419603/novanta-announces-financial-results-for-the-third-quarter-2026"
Q3C = "Earnings Whispers 第 3 季評論與逐字稿摘要（2026-11-10）https://www.earningswhispers.com/news/420263/novanta-beats-on-q3-and-bookings-but-a-stockport-halt-cuts-margins-and-q4-earnings-growth-to-0-5"
F25 = "公司 2025 全年財報新聞稿（2026-02-23；Stock Titan 轉載）https://www.stocktitan.net/news/NOVT/novanta-announces-financial-results-for-the-fourth-quarter-and-full-xwoj6lugo4hv.html"
F24 = "公司 2024 全年財報新聞稿（2025-02-25；搜尋摘要含 2023 比較數）https://www.sec.gov/Archives/edgar/data/1076930/000095017025026093/novt-ex99_1.htm"
KYB = "physical_ai/keyence/build_keyence.py（Novanta 前瞻 P/E 29.67 為資料商 2026-11-13 頁面值；5 年 β 1.70；同業 6 家 β 中位數 1.06）"
YCL = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
UST = "美國 10 年期公債殖利率 5.10%（FRED DGS10 2026-11-12；與 physical_ai/symbotic 同一取值，跨模型一致；逾期）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"
L6 = "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx"
L4 = "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx"
HWX = "physical_ai/hiwin/dist/20261123_上銀個股模型_v0.1.1.xlsx"
HDX = "physical_ai/harmonic_drive/dist/20261123_哈默納科個股模型_v0.1.1.xlsx"
NBX = "physical_ai/nabtesco/dist/20261123_Nabtesco個股模型_v0.1.1.xlsx"
SMCR = "physical_ai/smc/README.md"
KYR = "physical_ai/keyence/README.md"
RRXS = "StockStory 引 S&P 資料（Regal Rexnord 第 3 季 2026 報導；未來 12 個月 EPS 12.39）https://stockstory.org/us/stocks/nyse/rrx/news/earnings/regal-rexnord-nyserrx-exceeds-q3-2026-expectations"

ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. FY2025 實績：分部營收（US$M）")
inp("a25_aet", "營收 FY2025：Automation Enabling Technologies（AET）", "US$M", None, 500.8, None, "Verified", F25, "FY2024 490.6")
inp("a25_med", "營收 FY2025：Medical Solutions", "US$M", None, 479.8, None, "Verified", F25, "FY2024 458.6；合計 980.6 ＝ 合併營收")

sec("B. FY2026 錨點：公司指引（2026-11-10；Interested-party）")
inp("g26_rev", "公司指引：FY2026 合併營收（區間 1,135–1,140 取中點）", "US$M", None, 1137.5, None, "Interested-party", Q3R + "：1,135–1,140（約 +16%，有機約 +7%）", "9M 分部營收未取得；Q3 305.0、Q4 指引 306–312", CO)
inp("g26_ebitda", "公司指引：FY2026 調整後 EBITDA（274–277 取中點）", "US$M", None, 275.5, None, "Interested-party", Q3R + "：274–277（成長逾 24%）", "調整後＝非 GAAP，排除攤銷、股票報酬、併購與重整成本", CO)
inp("g26_epsa", "公司指引：FY2026 調整後稀釋 EPS（3.66–3.71 取中點）", "US$", None, 3.685, None, "Interested-party", Q3R, "前次 3.68–3.74（下修）", CO)
inp("g26_gma", "公司指引：FY2026 調整後毛利率", "%", None, 0.465, None, "Interested-party", Q3C + "：約 46.5%（前次更高；Stockport 英國廠微影產品出貨暫停至 2027Q1）", "Q3 調整後 47.0%；GAAP 44.5%（135.6÷305.0，Derived）", CO)
inp("gap26", "GAAP 毛利率低於調整後毛利率的幅度（FY2026）", "pt", 0.017, 0.022, 0.028, "Derived", "FY2025：GAAP 44.4%（435.3÷980.6）vs 調整後約 46.1%（約 452÷980.6）＝1.7pt；Q3'26：2.5pt（" + Q3R + "）", "取 2.2pt；區間 1.7–2.8")
inp("aet26", "AET FY2026 年增（相對 FY2025）", "%", 0.08, 0.11, 0.14, "Assumed", "Q3'26 AET 145 vs 125（+16%）；Q3 評論：ATI 所在部門有機成長（" + Q3R + "；" + Q3C + "）", "前兩季與全年 AET 營收未取得；Medical＝合併指引 − AET（Derived）")
inp("g_anc", "FY2026 實績相對公司指引（營收、EBITDA 同比例）", "%", -0.02, 0.0, 0.02, "Assumed", "基準＝公司指引中點（Q3 已實績，偏離只來自 Q4；Q3 與 Q4 指引都超過公司自己的前次區間）", "低／高 ±2%")
inp("da_26", "FY2026 折舊與攤銷（合計）", "US$M", 82.0, 88.0, 92.0, "Assumed", "Q3'26 27.4（含 Riverpoint 約兩個月；" + Q3R + "）；FY2025 61.9；Q4 以約 29 估", "Q1–Q2 合計未取得")
inp("sbc_26", "FY2026 股票報酬費用", "US$M", 30.0, 36.0, 40.0, "Assumed", "Q3'26 10.1（" + Q3R + "）；FY2025 未取得", "")
inp("oth_26", "FY2026 其他調整項（併購、整合、重整、庫存增值等；GAAP 營業利益與調整後 EBITDA 的差，扣除折舊攤銷與股票報酬後）", "US$M", 35.0, 50.0, 65.0, "Assumed",
    "Q3'26 反推：80.3 − 27.4 − 10.1 − 13.9（GAAP 營業利益）＝28.9（Derived，" + Q3R + "）；前兩季未取得", "Riverpoint 併購成本集中在 Q3；FY2026 取 50")
inp("fin_26", "FY2026 業外淨額（利息淨額與其他）", "US$M", -38.0, -34.0, -30.0, "Assumed", "Q3'26 利息淨額 −8.3（" + Q3R + "）；債務 7.8 → 8.3 億美元在 Q3 提高；FY2025 −24.4（稅前 69.6 − 營業利益 94.0）", "")
inp("sh_26", "FY2026 平均稀釋股數", "百萬股", 39.0, 40.0, 41.0, "Derived", "Q3'26 42.709、Q3'25 36.040、FY2025 36.7（" + Q3R + "；" + F25 + "）；Q3 發股籌資併購，前兩季約 36.8（假設）", "")

sec("C. 命題模組一：AET 核心與 Medical Solutions 營收（不含人形子列）")
GTXT = {
    "ga27": ("AET FY2027", {"A": (0.12, "樂觀：Q3 AET +16%、全公司訂單（bookings）+29%（不含 Riverpoint +18%）延續（" + Q3C + "）"),
                            "B": (0.08, "基準：FY2026 指引有機約 +7%，Q3 加速；訂單 +18%（不含併購）略高於營收"),
                            "C": (0.02, "保守：工業與半導體相關需求降溫（2025 AET +2.1%：490.6 → 500.8 為類比）")}),
    "ga28": ("AET FY2028", {"A": (0.10, ""), "B": (0.07, "基準：中個位數至高個位數"), "C": (0.0, "保守：循環下行")}),
    "gam": ("AET FY2029–FY2030", {"A": (0.08, ""), "B": (0.06, ""), "C": (0.03, "")}),
    "gal": ("AET FY2031–FY2032", {"A": (0.06, ""), "B": (0.05, ""), "C": (0.03, "")}),
    "gmed27": ("Medical FY2027", {"A": (0.25, "樂觀：Riverpoint 全年併入"), "B": (0.22, "基準：公司 2027 框架報告營收成長逾 15%（" + Q3C + "）；Riverpoint 於 2026-07 完成併購、2027 全年併入，扣除 AET 約 8% 後反推 Medical 約 22%（Derived）"), "C": (0.16, "保守：Riverpoint 貢獻低於預期")}),
    "gmed28": ("Medical FY2028", {"A": (0.08, ""), "B": (0.06, "基準：併購後有機成長中個位數"), "C": (0.03, "")}),
    "gmedm": ("Medical FY2029–FY2030", {"A": (0.07, ""), "B": (0.05, ""), "C": (0.03, "")}),
    "gmedl": ("Medical FY2031–FY2032", {"A": (0.06, ""), "B": (0.04, ""), "C": (0.03, "")}),
}
GRNG = {"ga27": (0.02, 0.12), "ga28": (0.0, 0.10), "gam": (0.03, 0.08), "gal": (0.03, 0.06), "gmed27": (0.16, 0.25), "gmed28": (0.03, 0.08), "gmedm": (0.03, 0.07), "gmedl": (0.03, 0.06)}
for f, (yl, d) in GTXT.items():
    for s in SC:
        v, note = d[s]
        lo, hi = GRNG[f] if s == "B" else (None, None)
        inp(f"{f}_{s}", f"情境 {s}：{yl} 營收年增", "%/年", lo, v, hi, "Assumed", note or "A／C 以 B 上下設定", "區間見情境 A／B／C 三列" if s == "B" else "")

sec("C2. 命題模組二：人形機器人六維力感測（命題子列；ATI 力感測；增量，不含在 AET 核心成長內）")
for s, col, v in (("A", "D", 3371.24590344), ("B", "C", 727.43565423), ("C", "B", 217.51758722304)):
    inp(f"m30_{s}", f"情境 {s}：六維力感測市場 2030（線 6；{'高' if s == 'A' else '基準' if s == 'B' else '低'}需求情境）", "百萬美元", None, v, None, "Derived",
        f"{L6} 下游取數!{col}22（五線歸因；人形占 93%；基準單價 1,500 美元、年降 18%）", "線 6 需求情境：A＝高、B＝基準、C＝低；2026 與 2032 市場值線模型未列入下游取數（只有 2030）")
for s, v, lo, hi in (("A", 0.12, None, None), ("B", 0.05, 0.01, 0.12), ("C", 0.01, None, None)):
    inp(f"shr_{s}", f"情境 {s}：ATI 在人形六維力市場（以線 6 市場為可及市場）的份額", "%", lo, v, hi, "Assumed",
        f"未取得：Novanta 未揭露人形營收或份額；線 6 份額表未列 Novanta（{L6} 份額!F15）；中國人形份額由藍點觸控超過 80%（Interested-party 賽迪）", "代理口徑；全球工業與協作機器人 ATI 為主要供應商但不在此子列")
inp("rp27", "人形營收爬坡：FY2027 占 2030 水準比例（線性升到 FY2030＝100%，FY2031 起持平）", "%", 0.0, 0.10, 0.25, "Assumed", "比照 physical_ai/thk 與上銀的爬坡設定；市場自身成長含在爬坡內（線模型只有 2030 值）", "2031–2032 不再成長（保守；線 6 的 2032 值未列入下游取數）")

sec("D. 損益結構與調整後指標")
inp("gm_aet", "AET 毛利率（GAAP，FY2027 起；人形子列同）", "%", 0.47, 0.495, 0.52, "Derived", "Q3'26 50.2%、Q3'25 47.3%（" + Q3R + "）", "分部毛利為公司揭露的近似 GAAP 值")
inp("gm_med", "Medical 毛利率（GAAP，FY2027 起）", "%", 0.38, 0.405, 0.43, "Derived", "Q3'26 39.9%、Q3'25 43.0%（含 Riverpoint 與 Stockport 影響；" + Q3R + "）", "")
for s, v, lo, hi in (("A", 0.005, None, None), ("B", 0.0, -0.01, 0.005), ("C", -0.015, None, None)):
    inp(f"gmadj_{s}", f"情境 {s}：FY2027 起毛利率額外調整", "pt", lo, v, hi, "Assumed", "Stockport 微影產品暫停至 2027Q1，Q4 毛利率約降 2pt（" + Q3C + "）；C：中國力感測同業價格競爭（藍點觸控等）", "")
for s, v, lo, hi in (("A", 0.263, None, None), ("B", 0.253, 0.24, 0.265), ("C", 0.24, None, None)):
    inp(f"eb27_{s}", f"情境 {s}：FY2027 調整後 EBITDA 率", "%", lo, v, hi, "Assumed", "FY2026 指引 24.2%（275.5÷1,137.5，Derived）；公司 2027 框架：報告營收 +15% 以上、調整後 EBITDA +20% 以上 → 約 25.3%（Derived，" + Q3C + "）", "")
for s, v, lo, hi in (("A", 0.005, None, None), ("B", 0.003, 0.0, 0.006), ("C", 0.0, None, None)):
    inp(f"ebx_{s}", f"情境 {s}：FY2028 起調整後 EBITDA 率每年改善", "pt/年", lo, v, hi, "Assumed", "規模與整合效益；無公司中期目標可引（未找到）", "")
inp("da_pct", "折舊攤銷 ÷ 營收（FY2027 起）", "%", 0.075, 0.09, 0.095, "Derived", "Q3'26 27.4 ÷ 305.0 ＝ 9.0%（含 Riverpoint 約兩個月，" + Q3R + "）", "FY2025 6.3%；攤銷隨併購無形資產上升")
inp("dep_sh", "折舊占折舊攤銷比例（調整後營業利益只加回攤銷）", "x", 0.5, 0.6, 0.7, "Derived", "Q3'26：調整後 EBITDA 80.3 − 調整後營業利益 65.1 ＝ 15.2 ÷ 27.4 ＝ 55%；Q3'25：68%（" + Q3R + "）", "取 60%，使 FY2026 調整後 EPS 對齊指引")
inp("sbc_pct", "股票報酬 ÷ 營收（FY2027 起）", "%", 0.028, 0.032, 0.036, "Derived", "Q3'26 10.1 ÷ 305.0 ＝ 3.3%（" + Q3R + "）", "DCF 與 GAAP 營業利益把股票報酬當成本")
inp("oth_27", "FY2027 其他調整項（整合與重整）", "US$M", 10.0, 25.0, 40.0, "Assumed", "Riverpoint 整合；FY2026 取 50", "")
inp("oth_lt", "FY2028 起其他調整項", "US$M", 5.0, 10.0, 20.0, "Assumed", "常態重整成本", "")
inp("fin_x", "業外淨額（FY2027 起；利息為主）", "US$M", -48.0, -42.0, -32.0, "Assumed", "債務 829.8、現金 149.2（2026-10-02；" + Q3R + "），淨債務 680.6 × 約 6.2%＝約 42", "公司以營運現金流降債；基準不另設還債（保守）")
inp("tax_w", "GAAP 有效稅率（FY2026 起）", "%", 0.18, 0.22, 0.26, "Assumed", "FY2025 22.7%（15.8÷69.6）、FY2024 19.0%（15.0÷79.1）（" + F25 + "）；Q3'26 68%（稅前利益過小，不採）", "")
inp("tax_adj", "調整後稅率（只用於調整後 EPS）", "%", 0.20, 0.23, 0.25, "Derived", "FY2025：調整後淨利約 120.7（3.29 × 36.7）÷（調整後營業利益約 179 − 利息 21.5）→ 約 23%；Q3'26 約 21%", "調整後營業利益由調整後 EBITDA 減折舊（Derived）")
inp("sh_v", "評價日與 FY2027 起稀釋股數", "百萬股", 42.0, 42.709, 43.5, "Derived", "Q3'26 稀釋加權平均 42.709（" + Q3R + "）；Q3 中旬增發，期末股數可能略高（未取得期末股數）", "選擇權稀釋已含在稀釋股數")

sec("E. 現金流")
inp("fc_cx", "資本支出 ÷ 營收（FY2027 起）", "%", 0.02, 0.03, 0.04, "Derived", "Q3'26 9.0 ÷ 305.0 ＝ 3.0%；FY2025 15.6 ÷ 980.6 ＝ 1.6%；FY2024 1.8%（" + Q3R + "；" + F25 + "）", "")
inp("cx_26", "FY2026 資本支出", "US$M", None, 33.0, None, "Assumed", "9M 未取得；Q3 9.0，全年約 3%（Q1–Q2 合計低於此）", "")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.10, 0.15, 0.25, "Assumed", "資產負債表營運資金細項未取得；製造業常態 10–25%", "")

sec("F. 估值輸入（評價日 2026-12-31；資產負債表 2026-10-02）")
inp("px", "股價（2026-11-20 收盤）", "US$", None, 101.99, None, "Verified", YCL + "：NOVT 2026-11-20 收盤 101.99", "同業大跌週：11-18 117.89、11-19 114.88；第 3 季財報後 2026-11-13 收盤 118.88")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("cash_v", "現金及約當現金（2026-10-02）", "US$M", None, 149.2, None, "Verified", Q3R, "")
inp("debt_v", "債務（流動 48.9＋長期 780.9；2026-10-02）", "US$M", None, "=48.9+780.9", None, "Verified", Q3R + "（原值相加 829.8）", "營業租賃負債、Riverpoint 或有對價（milestone 付款，金額未揭露）未計入")
inp("stub_frac", "Stub 期（2026-10-02 → 評價日 2026-12-31）占 FY2026 比例", "年", None, "=(DATE(2026,12,31)-DATE(2026,10,2))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac（粗略；FY2026 含併購前後）")
inp("rf", "無風險利率（美國 10 年期公債）", "%", None, 0.051, None, "Verified", UST, "")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "國家風險溢酬（美國）", "%", None, 0.0, None, "Assumed", "美國＝0", "")
inp("beta", "β（兩來源 Blume 調整後中位數）", "x", "=E{r}-0.3", "=MEDIAN(0.67*1.70+0.33,0.67*1.06+0.33)", "=E{r}+0.3", "Derived", KYB, "Novanta 自身 5 年 β 1.70；同業 6 家（Cognex、Omron、Novanta、Hexagon、SMC、Fanuc）中位數 1.06；Blume 0.67β＋0.33；敏感度 ±0.3")
inp("kd", "稅前負債成本", "%", None, 0.065, None, "Assumed", "債務主要為浮動利率信用額度：Q3 利息淨額 8.3 ÷ 債務約 0.8 億級（粗估 6–7%）", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("dps_tgt", "評價日到目標時點之每股股利", "US$", None, 0.0, None, "Verified", "公司未配發普通股股利（未找到股利紀錄；以 0 處理）", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable：同業前瞻 P/E × 調整後 EPS）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺 Novanta 對同業 10 年前瞻 P/E 比值歷史）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業前瞻 P/E（全部以 2026-11-20 收盤換算；收盤檔沒有的同業不納入中位數，列在下方「報價日未取得」）")
PEV = dict(pe_fanuc=21.1499, pe_yask=15.6739, pe_nab=13.7701, pe_smc=55300 / 3789.21, pe_kyn=28.6, pe_rrx=152.35 / 12.39)
PEERS = [("pe_fanuc", "Fanuc 6954（機器人、FA）", "21.1499", "Derived", HWX + " Comps!F12（換算至 2026-11-20；收盤 5,047）"),
         ("pe_yask", "安川電機 6506（伺服、機器人）", "15.6739", "Derived", HWX + " Comps!F13（換算至 2026-11-20；收盤 3,180）"),
         ("pe_nab", "Nabtesco 6268（精密減速機）", "13.7701", "Derived", HDX + " Comp!F5（換算至 2026-11-20；收盤 3,105）"),
         ("pe_smc", "SMC 6273（氣動元件）", "55300/3789.21", "Derived", f"{SMCR}：2026-11-20 收盤 55,300 ÷ 公司 FY2026 指引 EPS 3,789.21（14.59 倍）"),
         ("pe_kyn", "Keyence 6861（感測、視覺）", "28.6", "Derived", f"{KYR}：現價 ÷ 共識 FY2026 EPS 28.6 倍（收盤 14,025，2026-11-20）"),
         ("pe_rrx", "Regal Rexnord RRX（工業動力傳動、馬達）", "152.35/12.39", "Interested-party", f"2026-11-20 收盤 152.35（{YCL}）÷ 未來 12 個月 EPS 預期 12.39（{RRXS}；非曆年 2027 共識，資料商代理）")]
for k, lab, f, tag, src in PEERS:
    inp(k, f"前瞻 P/E：{lab}", "x", None, "=" + f, None, tag, src, "股權口徑 P/E；Novanta 的分母用調整後 EPS（排除攤銷與股票報酬），日本同業用 IFRS／日本基準 EPS，口徑不同", DV)
for k, lab, v, src in (("pe_cgnx", "Cognex CGNX（機器視覺）", 31.30, KYB.split("；")[0] + "；Cognex 2026-10-30 頁面值 31.30，見 physical_ai/keyence In 頁"),
                       ("pe_omron", "Omron 6645（感測、控制）", 26.57, "physical_ai/keyence In 頁：2026-09-18 頁面值 26.57"),
                       ("pe_hexa", "Hexagon HEXA-B（量測）", 27.28, "physical_ai/keyence In 頁：2026-11-13 頁面值 27.28")):
    inp(k, f"報價日未取得（不進中位數）：{lab}", "x", None, v, None, "Interested-party", src + "；收盤檔無此檔，報價日未取得 2026-11-20", "只作對照；若納入，中位數見 pe_med 高案", DV)
inp("pe_med", "同業前瞻 P/E 中位數（6 家，有 2026-11-20 收盤者；進 Comparable）", "x", 12.30, "=MEDIAN(pe_fanuc,pe_yask,pe_nab,pe_smc,pe_kyn,pe_rrx)", 21.15, "Derived",
    "6 家中位數；日本與美國工業自動化同業，非 Novanta 的醫療光學直接對手", "低＝6 家最低值（RRX 12.30）；高＝加入未取得報價日的 Cognex、Omron、Hexagon 後 9 家中位數 21.15（Derived）；少於約 8 家不做迴歸")
inp("pe_self", "Novanta 現行前瞻 P/E（換算至 2026-11-20）", "x", None, "=29.67*101.99/118.88", None, "Derived", KYB + "：2026-11-13 頁面值 29.67 × 11-20 收盤 101.99 ÷ 11-13 收盤 118.88（分母 EPS 不變）", "只作對照，不進模型", DV)
inp("cons_eps", "隱含共識前瞻 EPS（現價 ÷ 現行前瞻 P/E）", "US$", None, "=px/pe_self", None, "Derived", "同上", "只作對照，不進模型（口徑：資料商前瞻，非曆年 FY2027）", DV)


def build(path, ov=None, sens=None, rev=None, bid="{BID}"):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 US$M、股數百萬股、EPS 與股價美元"
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
    FAM = [("ga27", "AET 年增 FY2027"), ("ga28", "AET 年增 FY2028"), ("gam", "AET 年增 FY2029–FY2030"), ("gal", "AET 年增 FY2031–FY2032"),
           ("gmed27", "Medical 年增 FY2027"), ("gmed28", "Medical 年增 FY2028"), ("gmedm", "Medical 年增 FY2029–FY2030"), ("gmedl", "Medical 年增 FY2031–FY2032"),
           ("m30", "六維力感測市場 2030（線 6，百萬美元）"), ("shr", "ATI 人形六維力份額"), ("gmadj", "FY2027 起毛利率調整"), ("eb27", "FY2027 調整後 EBITDA 率"), ("ebx", "FY2028 起 EBITDA 率年改善")]
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
    H["A1"] = "Hist — 歷史財務（年度；US$M，美國 GAAP；調整後指標為公司非 GAAP 定義）。公式只引用本頁（定義名稱 h_*）；未取得者標「未取得」"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "標記", "來源"])
    NA = "未取得"
    HR = [("h_rev", "營收", 881.7, 949.2, 980.6, "Verified", F24 + "；" + F25),
          ("h_gp", "毛利（GAAP）", NA, 421.5, 435.3, "Verified", F25),
          ("h_op", "營業利益（GAAP）", 110.5, 110.6, 94.0, "Verified", F24 + "；" + F25),
          ("h_fin", "業外：稅前利益 − 營業利益（利息淨額與其他）", NA, "=79.1-110.6", "=69.6-94.0", "Derived", F25 + "（FY2024 利息淨額 −31.5；FY2025 −21.5，其餘為匯兌與其他）"),
          ("h_pbt", "稅前利益", NA, 79.1, 69.6, "Verified", F25),
          ("h_tax", "所得稅", NA, 15.0, 15.8, "Verified", F25),
          ("h_ni", "淨利（GAAP）", 72.9, 64.1, 53.8, "Verified", F24 + "；" + F25),
          ("h_eps", "稀釋 EPS（GAAP，美元）", 2.02, 1.77, 1.47, "Verified", F24 + "；" + F25),
          ("h_sh", "稀釋加權平均股數（百萬股）", NA, 36.1, 36.7, "Verified", F25),
          ("h_ebitda", "調整後 EBITDA（非 GAAP）", NA, 209.8, 221.0, "Interested-party", F25 + "（公司定義）"),
          ("h_epsa", "調整後稀釋 EPS（非 GAAP，美元）", NA, "=3.29/1.07", 3.29, "Interested-party", F25 + "：FY2025 3.29，年增約 7%；FY2024 為倒推（Derived）"),
          ("h_da", "折舊與攤銷", NA, 55.6, 61.9, "Verified", F25),
          ("h_cx", "資本支出", NA, 17.2, 15.6, "Verified", F25),
          ("h_ocf", "營業現金流", NA, 158.5, 64.1, "Verified", F25),
          ("h_aet", "AET 營收", NA, 490.6, 500.8, "Verified", F25),
          ("h_med", "Medical Solutions 營收", NA, 458.6, 479.8, "Verified", F25),
          ("h_cash", "現金（年底）", NA, 114.0, 380.9, "Verified", F25),
          ("h_debt", "債務（年底）", NA, 416.6, 250.8, "Verified", F25),
          ("h_eq", "股東權益（年底）", NA, 745.7, 1314.3, "Verified", F25)]
    for i, (k, lab, a, b, c, tag, src) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b, c)):
            cc = H.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): cc.font = BLUE
        H.cell(i, 5, tag); H.cell(i, 6, src)
        dn(k + "_23", f"Hist!$B${i}"); dn(k + "_24", f"Hist!$C${i}"); dn(k + "_25", f"Hist!$D${i}")
    H.column_dimensions["A"].width = 46

    # ---------------- Q（官方分部季度） ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 官方報告分部與 2026 年第 3 季（季末 2026-10-02）對去年同期（US$M，美國 GAAP；分部數字公司以百萬美元整數揭露）"
    hdr(Q, 2, ["項目", "Q3'25", "Q3'26", "標記", "來源", "來源誘因"])
    qrows = [("AET 營收", 125, 145), ("Medical Solutions 營收", 123, 160), ("分部營收合計", "=B3+B4", "=C3+C4"), ("合併營收", 247.836, 305.008),
             ("檢查：分部合計 − 合併營收（四捨五入至百萬）", "=ROUND(B5-B6,0)", "=ROUND(C5-C6,0)"), ("AET 毛利", 59, 73), ("Medical Solutions 毛利", 53, 64),
             ("營業利益（GAAP）", 21.848, 13.911), ("調整後 EBITDA", 58.1, 80.3), ("調整後稀釋 EPS（美元）", 0.87, 1.05), ("稀釋 EPS（GAAP，美元）", 0.30, 0.05),
             ("營業現金流", 8.5, 39.3), ("資本支出", 4.3, 9.0), ("折舊與攤銷", 15.711, 27.449), ("股票報酬", 6.698, 10.065), ("稀釋加權平均股數（百萬股）", 36.040, 42.709)]
    for i, (lab, a, b) in enumerate(qrows, 3):
        Q.cell(i, 1, lab)
        for j, v in ((2, a), (3, b)):
            c = Q.cell(i, j, v)
            if isinstance(v, (int, float)): c.font = BLUE
        Q.cell(i, 4, "Derived" if isinstance(a, str) else "Verified"); Q.cell(i, 5, Q3R); Q.cell(i, 6, CO)
    g0 = 3 + len(qrows) + 1
    hdr(Q, g0, ["公司指引（2026-11-10）", "低", "高", "標記", "來源", "來源誘因"])
    grows = [("Q4 2026 營收", 306, 312), ("Q4 2026 調整後 EBITDA", 76, 79), ("Q4 2026 調整後稀釋 EPS（美元）", 0.91, 0.96), ("FY2026 營收", 1135, 1140),
             ("FY2026 調整後 EBITDA", 274, 277), ("FY2026 調整後稀釋 EPS（美元）", 3.66, 3.71), ("FY2026 營業現金流（約）", 200, 200)]
    for i, (lab, a, b) in enumerate(grows, g0 + 1):
        Q.cell(i, 1, lab); Q.cell(i, 2, a).font = BLUE; Q.cell(i, 3, b).font = BLUE; Q.cell(i, 4, "Interested-party"); Q.cell(i, 5, Q3R); Q.cell(i, 6, CO)
    Q.cell(g0 + len(grows) + 2, 1, "註：2027 初步框架（報告營收成長 15% 以上、調整後 EBITDA 成長 20% 以上）、訂單 +29%（不含 Riverpoint +18%）見 " + Q3C)
    Q.column_dimensions["A"].width = 48

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（AET、Medical、人形六維力子列）→ 標準損益表（美國 GAAP）→ 現金流（US$M；FY2023–25 實績、FY2026 公司指引 ×(1＋g_anc)、FY2027 起模型）"
    PR = {}
    rr = 3
    KEYS = ["c_aet", "c_med", "hum", "rev", "chk1", "gr", "cogs", "gp", "gm", "opex", "op", "opm", "ebitda", "ebm", "fin", "pbt", "tax", "ni", "nim", "sh", "eps", "epsa", "nig",
            "da", "dep", "sbc", "oth", "cx", "dwc", "fcff", "fcfad", "mkt", "rp"]
    LAB = ["營收：Automation Enabling Technologies（AET）", "營收：Medical Solutions", "人形機器人六維力感測（命題子列，增量）", "營收（合併）", "檢查：分部合計 − 合併營收",
           "營收年增率", "營業成本", "毛利", "毛利率", "營業費用（毛利 − 營業利益；含銷管、研發、攤銷與併購整合）", "營業利益", "營業利益率",
           "調整後 EBITDA（非 GAAP）", "調整後 EBITDA 率", "業外：利息淨額與其他", "稅前利益", "所得稅", "稅後淨利", "淨利率", "稀釋平均股數（百萬股）", "EPS（GAAP 稀釋，美元）",
           "調整後 EPS（非 GAAP，美元；Comparable 分母）", "淨利年增率", "折舊與攤銷", "其中折舊（調整後營業利益只加回攤銷）", "股票報酬", "其他調整項（併購、整合、重整）", "資本支出", "營運資金增加",
           "FCFF（NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；股票報酬當成本）", "自由現金流扣稅後利息（收支關卡用；無股利）", "參考：六維力感測市場 2030（百萬美元，線 6）", "參考：人形營收爬坡（占 2030 水準）"]
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
                w("rev", f"=h_rev_{yy}"); w("op", f"=h_op_{yy}"); w("ni", f"=h_ni_{yy}"); w("eps", f"=h_eps_{yy}")
                w("opm", f"={cell('op')}/{cell('rev')}"); w("nim", f"={cell('ni')}/{cell('rev')}"); w("hum", 0); w("rp", 0)
                if y == 2023:
                    for k in ("c_aet", "c_med", "chk1", "gr", "cogs", "gp", "gm", "opex", "ebitda", "ebm", "fin", "pbt", "tax", "sh", "epsa", "nig", "da", "dep", "sbc", "oth", "cx", "dwc", "fcff", "fcfad", "mkt"):
                        w(k, "未取得")
                    w("hum", "未取得"); w("rp", "未取得"); continue
                w("c_aet", f"=h_aet_{yy}"); w("c_med", f"=h_med_{yy}")
                w("chk1", f"=ROUND({cell('c_aet')}+{cell('c_med')}+{cell('hum')}-{cell('rev')},3)")
                w("gp", f"=h_gp_{yy}"); w("cogs", f"={cell('rev')}-{cell('gp')}"); w("gm", f"={cell('gp')}/{cell('rev')}"); w("opex", f"={cell('gp')}-{cell('op')}")
                w("ebitda", f"=h_ebitda_{yy}"); w("ebm", f"={cell('ebitda')}/{cell('rev')}"); w("fin", f"=h_fin_{yy}"); w("pbt", f"=h_pbt_{yy}"); w("tax", f"=h_tax_{yy}")
                w("sh", f"=h_sh_{yy}"); w("epsa", f"=h_epsa_{yy}"); w("da", f"=h_da_{yy}"); w("cx", f"=h_cx_{yy}")
                for k in ("dep", "sbc", "oth", "dwc", "fcff", "fcfad", "mkt"): w(k, "未取得")
                w("gr", f"={cell('rev')}/{prev('rev')}-1"); w("nig", f"={cell('ni')}/{prev('ni')}-1")
                continue
            if y == 2026:
                w("rev", "=g26_rev*(1+g_anc)"); w("c_aet", "=a25_aet*(1+aet26)"); w("c_med", f"={cell('rev')}-{cell('c_aet')}"); w("hum", 0); w("rp", 0); w("mkt", 0)
                w("gp", f"={cell('rev')}*(g26_gma-gap26)"); w("ebitda", "=g26_ebitda*(1+g_anc)")
                w("da", "=da_26"); w("sbc", "=sbc_26"); w("oth", "=oth_26"); w("fin", "=fin_26"); w("sh", "=sh_26"); w("cx", "=cx_26")
            else:
                gA = "ga27" if y == 2027 else "ga28" if y == 2028 else "gam" if y <= 2030 else "gal"
                gM = "gmed27" if y == 2027 else "gmed28" if y == 2028 else "gmedm" if y <= 2030 else "gmedl"
                w("c_aet", f"={prev('c_aet')}*(1+{gA}_{s})"); w("c_med", f"={prev('c_med')}*(1+{gM}_{s})")
                w("mkt", f"=m30_{s}"); w("rp", f"=MIN(1,rp27+(1-rp27)*({y}-2027)/3)" if y <= 2030 else "=1")
                w("hum", f"={cell('mkt')}*shr_{s}*{cell('rp')}")
                w("rev", f"={cell('c_aet')}+{cell('c_med')}+{cell('hum')}")
                w("gp", f"={cell('c_aet')}*gm_aet+{cell('c_med')}*gm_med+{cell('hum')}*gm_aet+gmadj_{s}*{cell('rev')}")
                w("ebitda", f"={cell('rev')}*(eb27_{s}+ebx_{s}*({y}-2027))")
                w("da", f"=da_pct*{cell('rev')}"); w("sbc", f"=sbc_pct*{cell('rev')}"); w("oth", "=oth_27" if y == 2027 else "=oth_lt")
                w("fin", "=fin_x"); w("sh", "=sh_v"); w("cx", f"=fc_cx*{cell('rev')}")
            w("chk1", f"=ROUND({cell('c_aet')}+{cell('c_med')}+{cell('hum')}-{cell('rev')},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            w("cogs", f"={cell('rev')}-{cell('gp')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
            w("op", f"={cell('ebitda')}-{cell('da')}-{cell('sbc')}-{cell('oth')}"); w("opex", f"={cell('gp')}-{cell('op')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            w("ebm", f"={cell('ebitda')}/{cell('rev')}")
            w("pbt", f"={cell('op')}+{cell('fin')}"); w("tax", f"={cell('pbt')}*tax_w"); w("ni", f"={cell('pbt')}-{cell('tax')}"); w("nim", f"={cell('ni')}/{cell('rev')}")
            w("eps", f"={cell('ni')}/{cell('sh')}"); w("dep", f"=dep_sh*{cell('da')}")
            w("epsa", f"=({cell('ebitda')}-{cell('dep')}+{cell('fin')})*(1-tax_adj)/{cell('sh')}")
            w("nig", f"=IF({prev('ni')}>0,{cell('ni')}/{prev('ni')}-1,\"N/M（前一年為虧損）\")")
            w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
            w("fcff", f"={cell('op')}*(1-tax_w)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("fcfad", f"={cell('fcff')}+{cell('fin')}*(1-tax_w)")
        P.cell(R["c_aet"], 2, "FY2026＝FY2025 ×(1＋aet26)；FY2027 起＝前一年 ×(1＋年增)")
        P.cell(R["c_med"], 2, "FY2026＝合併指引 − AET（Derived）；FY2027 起同上")
        P.cell(R["hum"], 2, "市場 2030 × 份額 × 爬坡；FY2026 以前未揭露（視為 0）")
        P.cell(R["gp"], 2, "FY2026＝營收 ×(調整後毛利率指引 − GAAP 差)；FY2027 起＝分部營收 × 分部毛利率 ＋ 調整")
        P.cell(R["ebitda"], 2, "FY2026＝公司指引中點；FY2027 起＝營收 × EBITDA 率")
        P.cell(R["op"], 2, "＝調整後 EBITDA − 折舊攤銷 − 股票報酬 − 其他調整項")
        P.cell(R["tax"], 2, "FY2026 起＝稅前 × tax_w（全年套同一稅率）")
        P.cell(R["epsa"], 2, "＝(調整後 EBITDA − 折舊 ＋ 業外) ×(1 − 調整後稅率) ÷ 股數；FY2026 與指引對照見 Recon")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 ×(1 − tax_w)；不含併購現金支出（Riverpoint 約 12 億美元，已在 2026-10-02 資產負債表內）")
        for k in ("gr", "gm", "opm", "ebm", "nim", "nig", "rp"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(LAB) + 2
    P.column_dimensions["A"].width = 56; P.column_dimensions["B"].width = 30
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "epsa", "fcff", "fcfad", "hum"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；股票報酬當成本），WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（US$M）"; D["B4"] = "=px*sh_v"
    D["A5"] = "負債 D（債務；營業租賃未取得）"; D["B5"] = "=debt_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨債務（現金 − 債務；2026-10-02）"; D["B10"] = "=cash_v-debt_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND(((px*sh_v)*(rf+beta*(erp+crp))+debt_v*kd*(1-tax_w))/((px*sh_v)+debt_v)+wacc_adj-wacc,9)"
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
                 ("＋淨現金（負值為淨債務；2026-10-02）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!F{R['fcff']}*stub_frac"),
                 ("− 尚待買回金額（無）", 0),
                 ("股權價值", f"=B{t+4}+B{t+5}+B{t+6}+B{t+7}"),
                 ("每股價值（評價日）", f"=MAX(0,B{t+8})/sh_v"),
                 ("DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）", f"=B{t+9}*(1+ke)^t_tgt-dps_tgt"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+11}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=B{t+2}/G{b+1}"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）", f"=((B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6)+B{t+5}+B{t+6}+B{t+7})/sh_v*(1+ke)^t_tgt-dps_tgt")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 10, ps=t + 9, ev=t + 4, tvs=t + 11, alt=t + 14, chk=b + 6)
        dn(f"dcf_{s}", f"DCF!$B${t+10}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 60

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價：DCF（權重 w_dcf）＋ Comparable（同業前瞻 P/E 中位數 × 情境 FY2028 調整後 EPS，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}")
        V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "DCF 頁")
    r_cu = vrow("Comparable（未調整）：同業前瞻 P/E 中位數 × 情境 FY2028 調整後 EPS", [f"=pe_med*PL!H{PR[s]['epsa']}" for s in SC],
                "目標時點的 NTM 以 FY2028 代；分母為調整後 EPS（排除攤銷與股票報酬），與資料商前瞻 P/E 口徑相近；GAAP EPS 低很多（見 PL 頁）")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（美元）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    r_px = one("現價（In 頁 px）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣稅後利息合計（US$M；情境 A／B／C；無股利；不含併購現金支出）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=IF(B{r_up}>0.5,"已在最高級",TEXT(px*(1+IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5)))),"#,##0.00"))&"／"&IF(B{r_up}<-0.5,"已在最低級",TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,IF(B{r_up}>=-0.5,-0.5,-0.5))))),"#,##0.00"))', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 調整後 EPS", f"=B{r_tp}/PL!H{PR['B']['epsa']}", "分母：模型情境 B FY2028 調整後 EPS")
    r_pe1 = one("現價 ÷ 情境 B FY2027 調整後 EPS", f"=px/PL!G{PR['B']['epsa']}", "分母：模型情境 B FY2027")
    r_pe2 = one("Novanta 現行前瞻 P/E（資料商，換算至 2026-11-20）÷ 同業中位數", "=pe_self/pe_med", "只作對照")
    r_gaap = one("現價 ÷ 情境 B FY2027 GAAP EPS", f"=px/PL!G{PR['B']['eps']}", "GAAP 本益比；差距主要來自攤銷、股票報酬與併購整合成本")
    r_hum = one("人形六維力營收占 FY2030 合併營收（情境 B）", f"=PL!J{PR['B']['hum']}/PL!J{PR['B']['rev']}")
    for rr_ in (r_up, r_a1, r_a2, r_a3, r_hum): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0.00"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "#,##0.00"
    V.column_dimensions["A"].width = 70; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, ca=r_ca, pd2=r_pd2, pe=r_pe, pe1=r_pe1, pe2=r_pe2, gj=r_gj, wk=r_wk, hum=r_hum, gaap=r_gaap, px=r_px)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司指引、共識對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或共識", "差", "說明"])
    rc = [("FY2026 營收", f"=PL!F{PR['B']['rev']}", "=g26_rev", "公司指引中點（Interested-party）"),
          ("FY2026 調整後 EBITDA", f"=PL!F{PR['B']['ebitda']}", "=g26_ebitda", "公司指引中點"),
          ("FY2026 調整後 EPS（模型由 EBITDA 反推，驗證 dep_sh、利息、稅率與股數設定）", f"=PL!F{PR['B']['epsa']}", "=g26_epsa", "公司指引中點 3.66–3.71；差距在 0.05 內視為對齊"),
          ("FY2026 GAAP 營業利益（模型）", f"=PL!F{PR['B']['op']}", "未取得", "公司未給 GAAP 指引；Q3 實績 13.9"),
          ("FY2027 營收年增（模型 B）vs 公司 2027 框架下限", f"=PL!G{PR['B']['gr']}", 0.15, "公司初步框架：報告營收成長逾 15%（Interested-party）"),
          ("FY2027 調整後 EBITDA 年增（模型 B）vs 公司框架下限", f"=PL!G{PR['B']['ebitda']}/PL!F{PR['B']['ebitda']}-1", 0.20, "公司初步框架：調整後 EBITDA 成長逾 20%"),
          ("FY2027 調整後 EPS（模型 B）vs 資料商隱含前瞻 EPS", f"=PL!G{PR['B']['epsa']}", "=cons_eps", "資料商口徑為前瞻 12 個月，非曆年 FY2027；只看方向")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_)
        RC.cell(i, 4, f"=B{i}-C{i}" if not (isinstance(b_, str) and not b_.startswith("=")) else "未取得"); RC.cell(i, 5, n)
    for i in (7, 8): 
        for L in "BCD": RC[f"{L}{i}"].number_format = "0.0%"
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_novanta.py 批次格點＋線性內插；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 60; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_novanta.py 以 LibreOffice 批次重算每一案例）；結果為常數快照"
    base_tp = sens["base"]["tp"] if sens else None
    SE["A2"] = "快照基準值（加權，美元）"; SE["B2"] = base_tp
    SE["C2"] = "本檔即時值（加權，美元）"; SE["D2"] = "=tp"
    SE["E2"] = f"快照輸入指紋 {sens['fp'] if sens else ''}"
    SE["A3"] = "build_id"; SE["D3"] = bid
    SE["A4"] = "模型狀態"; SE["D4"] = '=IF(ABS(D2-B2)>0.05,"STALE","CURRENT")'
    hdr(SE, 6, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    sr = 7
    if sens:
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]); SE.cell(sr, 5, cs["tp"] / base_tp - 1)
            SE.cell(sr, 6, cs["dcf"]); SE.cell(sr, 7, cs["dcf"] / sens["base"]["dcf"] - 1); SE.cell(sr, 8, cs["rt"])
            SE[f"E{sr}"].number_format = "0.0%"; SE[f"G{sr}"].number_format = "0.0%"; SE[f"D{sr}"].number_format = "#,##0.00"; SE[f"F{sr}"].number_format = "#,##0.00"
            sr += 1
    SE.cell(sr + 1, 1, "說明：相對倍數係數 rel_k 基準即為 1.00（未調整），故「相對倍數係數＝1」案例＝基準。無區間的鍵（cx_26、cons_eps 等）不列案例；線 6 市場 m30 的 B 列不另設區間（線 6 低／高已作情境 C／A）。")
    SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 34

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = "對帳 — 本模型使用的線層級變數（讀線模型成品「下游取數」頁；與線模型不一致的自設值標明理由，不靜默改口徑）"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型一致？"])
    rec = [("六維力感測市場規模 2030（低／基準／高）", "6", "217.5／727.4／3,371.2", "百萬美元", "五線歸因；人形占 93%；基準單價 1,500 美元年降 18%", L6 + " 下游取數!B22:D22", "Derived", "In m30_C／m30_B／m30_A → PL 人形子列", "一致（照取）"),
           ("六維力單價（2026 基準、年降）", "6", "1,500 美元（700–3,000）、年降 18%（10–25%）", "美元", "零組件線輸入", L6 + " 輸入!D107:I107（見 map/profiles/NOVT.md）", "Interested-party", "未直接進公式（已含在 m30）", "線 6 與線 4 不一致：人形線六維力單價基準 1,000 美元（L4 輸入!F23）；本模型照所屬線 6，差異以 shr 區間涵蓋"),
           ("人形每台六維力顆數", "6／4", "2 顆（2–4）", "顆／台", "人形出貨組合平均", L6 + " 輸入!F33、F55、F77", "Derived", "含在 m30，未直接進公式", "一致（經線 6）"),
           ("六維力 2030 市場中人形占比", "6", "93%", "%", "五線歸因", L6 + " 計算!M239（見 map/profiles/NOVT.md）", "Derived", "命題子列依據", "一致"),
           ("Novanta 份額", "6", "未取得（線模型份額表未列 Novanta；中國由藍點觸控超過 80%）", "%", "中國人形", L6 + " 份額!F15", "—", "shr_* 為 Assumed 5%（1–12%）", "一致（線模型亦未取得）"),
           ("AET 核心（工業與協作機器人、其他工業感測）", "3", "不取線 3 裝機量", "—", "ATI 營收主體為工業感測與工具，工業機器人用量只占一部分（每台 0.05 顆傳統、0.3 顆協作；profile 引自線 6）", L6 + " 輸入!D55、E55", "Derived", "AET 成長以公司訂單與指引為主，未接線 3", "自設：AET 營收分部未拆到機器人用途，無法連到線 3 裝機量；不一致處以此說明"),
           ("Medical Solutions", "—", "不屬實體 AI 線", "—", "醫療視覺、流體、Riverpoint", "—", "—", "gmed* 為公司 2027 框架反推", "不適用")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 30; RE.column_dimensions["F"].width = 60; RE.column_dimensions["I"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("Riverpoint 整合與槓桿（淨債務 680.6、利息約 42／年）", "部分（fin_x、oth_27）", "不設還債；或有對價（milestone）金額未揭露、未計入", "季度利息、淨債務、整合成本"),
             ("Stockport 英國廠微影產品出貨暫停至 2027Q1、毛利率下修", "部分（gmadj、gm_med）", "公司 Q4 毛利率預期降約 2pt；復工時點為公司說法", "2027Q1 出貨恢復、毛利率"),
             ("人形與工業機器人需求低於線模型", "是（shr_C、m30_C、ga*_C）", "子列占營收比例小（見 Val）", "Optimus 等量產節奏、訂單"),
             ("中國國產六維力價格壓力（價格約 5,000 人民幣／顆，線 6）", "是（gmadj_C、shr_C）", "情境 C 毛利率與份額下調", "中國人形六維力出貨與價格"),
             ("調整後與 GAAP 差距：攤銷、股票報酬、併購成本", "是（Comparable 用調整後 EPS、DCF 用 GAAP 現金流）", "兩種口徑結論差距大，見 Val 的 GAAP 本益比列", "GAAP 與調整後 EPS 差"),
             ("匯率（歐元、英鎊）", "否", "未逐年建模", "歐元兌美元、英鎊兌美元"),
             ("營業租賃負債、股數期末值未取得", "否", "股數用 Q3 稀釋加權平均 42.709（Q3 中旬增發，期末可能略高）", "10-Q 資產負債表")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 50

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([Q3R, Q3C, F25, F24, KYB, YCL, UST, RRXS, L6 + "（下游取數、份額）", L4 + "（下游取數）",
                            HWX + "（Comps 頁 Fanuc、安川換算 P/E）", HDX + "（Comp 頁 Nabtesco 換算 P/E）", SMCR, KYR], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"Novanta（NOVT）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-20（第 3 季財報 2026-11-10）；成品日 2026-11-24；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"#,##0.00")&" 美元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：Novanta 的股價由兩件事決定：ATI 力感測與工業精密元件（AET）能否維持高於營收的獲利成長，以及併入 Riverpoint 後的醫療部門與槓桿能否消化；人形機器人六維力感測是小型選擇權，不是承重項"
    RD["A4"] = (f'="在 FY2027 AET 年增（情境 B）"&TEXT(ga27_B,"0%")&"、Medical 年增 "&TEXT(gmed27_B,"0%")&"、FY2027 調整後 EBITDA 率 "&TEXT(eb27_B,"0.0%")&"、同業前瞻 P/E 中位數 "&TEXT(pe_med,"0.0")&" 倍的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"#,##0.00")&" 美元（DCF "&TEXT(Val!B{r_dcft},"#,##0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"#,##0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣稅後利息 "&TEXT(Val!C{r_gate},"#,##0")&" 百萬美元）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 美元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：2027 年第 1 季 Stockport 仍未復工；FY2027 公司框架報告營收成長低於 15% 或調整後 EBITDA 成長低於 20%；訂單（不含併購）成長轉負；淨債務 ÷ 調整後 EBITDA 未在 2027 年下降"
    RD["A7"] = "範圍：兩個報告分部（AET、Medical）加人形六維力命題子列；不讀 Siliconomics／Tokenomics／Wattonomics（算力線）。實體 AI 營收曝險低（profile 判斷），Medical 約占一半。HTML 展示層見 dist/ 同名 .html"
    RD["A8"] = f'="命題回覆：問：AET 與醫療能否支持高於現價的目標價？答：情境 B FY2028 調整後 EPS "&TEXT(PL!H{PR["B"]["epsa"]},"0.00")&" 美元，現價是其 "&TEXT(px/PL!H{PR["B"]["epsa"]},"0.0")&" 倍，高於同業中位數 "&TEXT(pe_med,"0.0")&" 倍；加權目標價 "&TEXT(Val!B{r_tp},"#,##0.00")&" 美元。問：人形六維力有多重要？答：情境 B FY2030 子列只占合併營收 "&TEXT(Val!B{r_hum},"0.0%")&"，結論不靠人形（模型輸出，非獨立證據）。"'
    wb.save(path)
    return VR, PR, DR
# ---------------- 重算與讀值 ----------------
def lo_profile():
    d = os.path.join(os.environ.get("HOME", tempfile.gettempdir()), "lo_novanta")
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
    d = tempfile.mkdtemp(prefix="novanta_"); paths = []; VR = DR = None
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

CASES = [
    ("g_anc", "FY2026 實績相對公司指引"), ("aet26", "AET FY2026 年增（分部拆分）"), ("ga27_B", "AET 年增 FY2027（情境 B）"), ("ga28_B", "AET 年增 FY2028（情境 B）"),
    ("gam_B", "AET 年增 FY2029–30（情境 B）"), ("gal_B", "AET 年增 FY2031–32（情境 B）"), ("gmed27_B", "Medical 年增 FY2027（情境 B）"), ("gmed28_B", "Medical 年增 FY2028（情境 B）"),
    ("gmedm_B", "Medical 年增 FY2029–30（情境 B）"), ("gmedl_B", "Medical 年增 FY2031–32（情境 B）"), ("shr_B", "ATI 人形六維力份額（情境 B）"), ("rp27", "人形營收爬坡 FY2027"),
    ("gm_aet", "AET 毛利率"), ("gm_med", "Medical 毛利率"), ("gmadj_B", "毛利率調整（情境 B）"), ("eb27_B", "FY2027 調整後 EBITDA 率（情境 B）"), ("ebx_B", "EBITDA 率年改善（情境 B）"),
    ("gap26", "GAAP 與調整後毛利率差"), ("da_26", "FY2026 折舊攤銷"), ("sbc_26", "FY2026 股票報酬"), ("oth_26", "FY2026 其他調整項"), ("fin_26", "FY2026 業外淨額"), ("sh_26", "FY2026 平均股數"),
    ("da_pct", "折舊攤銷 ÷ 營收"), ("dep_sh", "折舊占折舊攤銷比例"), ("sbc_pct", "股票報酬 ÷ 營收"), ("oth_27", "FY2027 其他調整項"), ("oth_lt", "FY2028 起其他調整項"), ("fin_x", "業外淨額（FY2027 起）"),
    ("tax_w", "有效稅率"), ("tax_adj", "調整後稅率"), ("sh_v", "評價日股數"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"),
    ("beta", "β"), ("erp", "股權風險溢酬"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重"), ("pe_med", "同業前瞻 P/E 中位數")]
BETA_B = statistics.median([0.67 * 1.70 + 0.33, 0.67 * 1.06 + 0.33])
FXD = {"pe_med": statistics.median(PEV.values())}
PXN = 101.99


def solve_batch(key, lo, hi, targets):
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
    up_ = base["alt_tp"] / PXN - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rating(up_)))
    for c in cases: print(c["key"], c["drv"], round(c["tp"], 2), c["rt"])
    up0 = base["up"]; tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    revs = []
    for key, b0, lo_, hi_, note in (("eb27_B", 0.253, 0.10, 0.60, "情境 B 的 FY2027 調整後 EBITDA 率（其他不變，整本重算；情境 A／C 不動）"),
                                    ("pe_med", FXD["pe_med"], 0.0, 80.0, "同業前瞻 P/E 中位數（Comparable；DCF 不受影響）"),
                                    ("ga27_B", 0.08, -0.40, 1.50, "情境 B 的 FY2027 AET 年增（其他不變；之後年度在此基礎上成長）")):
        sol = solve_batch(key, lo_, hi_, [0.0, upper, lower])
        for lab, tgt in (("目標價＝現價（上檔 0%）", 0.0), (f"升一級邊界（上檔 {upper:+.0%}）" if upper is not None else None, upper), (f"降一級邊界（上檔 {lower:+.0%}）" if lower is not None else None, lower)):
            if lab is None: continue
            x = sol.get(tgt)
            revs.append(dict(label=lab, key=key, val=x if x is not None else f"區間 {lo_:+.2f}～{hi_:+.2f} 內無解", base=b0, note=note))
            print("反解", key, lab, x)
    sens = dict(base=base, cases=cases, fp=fingerprint())
    json.dump(sens, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="novf_"); raw = os.path.join(tmp, NAME)
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
    src = "公司 2025 全年財報新聞稿（2026-02-23）" + F25.split("）", 1)[1] if False else F25
    segs = [dict(name="Automation Enabling Technologies", label="營收：Automation Enabling Technologies（AET）", value=500.8),
            dict(name="Medical Solutions", label="營收：Medical Solutions", value=479.8)]
    json.dump(dict(company="novanta", ticker="NOVT", model_sheet="PL",
                   official_segments=[dict(period="FY2025", unit="US$M", source=F25 + "（兩個報告分部）", **x) for x in segs],
                   sublines=[{"parent": "Automation Enabling Technologies", "labels": ["人形機器人六維力感測（命題子列，增量）"]}],
                   check_labels=["檢查：分部合計 − 合併營收"],
                   note="兩個報告分部（AET、Medical Solutions）FY2025 營收加總等於合併營收 980.6；PL 頁「檢查：分部合計 − 合併營收」各年為零。人形六維力為命題子列（增量，屬 AET 的 ATI 力感測），計入合併營收。"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)


def write_readme(path, VR, PR, DR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f2 = lambda x: f"{x:,.2f}"; f0 = lambda x: f"{x:,.0f}"; pc = lambda x: f"{x*100:+.1f}%"
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
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f2(g(chr(66+i)+str(VR['dcfr'])))} | {f2(g(chr(66+i)+str(VR['cu'])))} | {f2(g(chr(66+i)+str(VR['tps'])))} | {g(chr(66+i)+str(VR['gate'])):,.0f} |" for i, s in enumerate(SC))
    rts = sorted(set(c['rt'] for c in sens['cases']))
    YS = range(2023, 2033)
    txt = f"""# physical_ai/novanta — Novanta（NOVT）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

NASDAQ 上市。本 README 的數字由 `build_novanta.py` 從重算後的成品讀出產生，不手寫。

## 未完成與限制（先看）
- 未取得：FY2026 前三季分部營收與 GAAP 損益（只有 Q3 與全年指引）、期末股數、營業租賃負債、Riverpoint 或有對價金額、Riverpoint 單獨營收與利潤、FY2023 毛利以下細項、賣方共識目標價與逐年 EPS。FY2026 分部拆分與 GAAP 營業利益是由指引反推（Assumed／Derived，見 In 頁）。
- 同業前瞻 P/E 只納入收盤檔有 2026-11-20 收盤的 6 家（Fanuc、安川、Nabtesco、SMC、Keyence、Regal Rexnord）；Cognex、Omron、Hexagon 收盤檔沒有，**報價日未取得**，不進中位數，只在 pe_med 高案（9 家中位數 21.15）出現。6 家多為日本 FA 同業，不是 Novanta 的醫療光學直接對手，Comparable 偏低的可能性要看高案。
- Comparable 分母用調整後 EPS（排除攤銷、股票報酬、併購成本）；DCF 用 GAAP 現金流並把股票報酬當成本。兩者口徑不同，現價對 GAAP EPS 的本益比見 Val 頁。
- DCF 偏低的結構原因：FY2032 營業利益仍扣除約 9% 營收的折舊攤銷（多為併購無形資產攤銷，非現金），終值用 NOPAT 法只約 9.9 倍；終值替代法（FCFF 法）見 Sens 頁 `tv_alt`，加權目標價較高。毛利率輸入（gm_aet、gm_med、gmadj、gap26）只影響營業費用拆分，不影響目標價（營業利益由 EBITDA 率推得），所以敏感度為零。
- 人形六維力份額、營收爬坡、FY2028 起 EBITDA 率改善為 Assumed；無公司中期目標可引。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A8'].value}

## 估值（評價日 2026-12-31；目標時點約 12 個月）
| 情境 | DCF（推到 12 個月後） | Comparable（P/E × FY2028 調整後 EPS） | 情境目標價 | 收支關卡（FY2026–30，US$M） |
|---|---|---|---|---|
{scn}

- 加權目標價 **{f2(tp)} 美元**（DCF {f2(dcf)}、Comparable {f2(cmp_)}），現價 {f2(inv('px'))}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")} → 投資建議 **{g(f"B{VR['rec']}")}**。
- 方法：DCF 標準版 FCFF（{g(f"B{VR['wk']}")}，為 WACC／ke／g／RONIC）50%＋Comparable 50%（同業 6 家前瞻 P/E 中位數 {inv('pe_med'):.2f} 倍 × 各情境 FY2028 調整後 EPS）。相對倍數係數 1.00（未調整）。
- 隱含前瞻 P/E（加權目標價 ÷ 情境 B FY2028 調整後 EPS）{g(f"B{VR['pe']}"):.1f} 倍；資料商現行前瞻 P/E {inv('pe_self'):.1f} 倍 ＝ 同業中位數的 {g(f"B{VR['pe2']}"):.2f} 倍；現價 ÷ 情境 B FY2027 GAAP EPS {g(f"B{VR['gaap']}"):.1f} 倍。

## 營收與標準損益（情境 B，US$M；FY2023–25 實績、FY2026 公司指引、FY2027 起模型）
| 列 | {" | ".join(f"FY{y}" for y in YS)} |
|---|{"---|" * len(YS)}
| AET | {yr('B','c_aet')} |
| Medical Solutions | {yr('B','c_med')} |
| 人形六維力（命題子列） | {yr('B','hum')} |
| 營收（合併） | {yr('B','rev')} |
| 毛利 | {yr('B','gp')} |
| 營業利益（GAAP） | {yr('B','op')} |
| 調整後 EBITDA | {yr('B','ebitda')} |
| 淨利（GAAP） | {yr('B','ni')} |
| EPS（GAAP，美元） | {yr('B','eps','{:,.2f}')} |
| 調整後 EPS（美元） | {yr('B','epsa','{:,.2f}')} |

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- AET 年增（A／B／C）：FY2027 12%／8%／2%；Medical FY2027 25%／22%／16%（由公司 2027 框架「報告營收 +15% 以上、調整後 EBITDA +20% 以上」反推，Derived）。
- 人形六維力：線 6 市場 2030（高／基準／低 3,371／727／218 百萬美元）× ATI 份額 12%／5%／1% × 爬坡（FY2027 10% → FY2030 100%）。
- 利潤：FY2027 調整後 EBITDA 率 26.3%／25.3%／24.0%；FY2028 起每年 +0.5／+0.3／0 pt；GAAP 營業利益＝EBITDA − 折舊攤銷（9% 營收）− 股票報酬（3.2%）− 其他調整項。
- 折現：rf 5.10%（逾期，跨模型一致）、β {inv('beta'):.3f}、ERP 5%、g＝MIN(3%, rf)、RONIC＝WACC；淨債務 {-wb['DCF']['B10'].value:,.1f}（2026-10-02）；股數 {inv('sh_v'):.3f} 百萬。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| FY2027 EBITDA 率（eb27） | 2027 年框架更新、季度調整後 EBITDA 率、Stockport 復工 |
| 同業 P/E（pe_med） | 同業倍數換算到同一估值日後的中位數；Novanta 自身前瞻 P/E |
| AET 與 Medical 成長（ga27、gmed27） | 訂單成長（不含併購）、Riverpoint 季度貢獻 |

失效條件：{RD['A6'].value.split('：', 1)[1]}。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f2(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁；各案例估值評等出現「{'、'.join(rts)}」。

## 反解（整本重算；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
- 線 6：六維力感測市場 2030 低／基準／高（下游取數!B22:D22）→ m30 的情境 C／B／A，照取；2026、2032 值線模型未列入下游取數，故爬坡含市場成長。
- 與線模型不一致：線 6 六維力單價基準 1,500 美元、線 4 為 1,000 美元（線模型自身待對帳），本模型照線 6；AET 營收未拆到機器人用途，不接線 3 裝機量（自設）。Medical 不屬實體 AI 線。

## 檔案
- `build_novanta.py`：建置＋整本重算 Sens／反解（LibreOffice 批次）＋產生本 README、`segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-24）：初版；現價 {f2(inv('px'))}（2026-11-20 收盤）；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


if __name__ == "__main__":
    main()

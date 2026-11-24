# Schaeffler 舍弗勒（SHA0.DE）個股模型 v0.1 — AI實體應用 Project（線 6 共用零組件：行星滾柱絲杠、致動器；並讀線 4 人形）
# 用法：python3 physical_ai/schaeffler/build_schaeffler.py   → 建置、整本重算 Sens 與反解（LibreOffice 批次）、輸出 dist/ 成品、README.md、sens.json、segments.json
# 會計年度＝曆年（12 月決算，IFRS）。金額：百萬歐元（€M）；股數：百萬股；EPS、股價：歐元。
# 2024 年 10 月合併 Vitesco：FY2024 報告值只含 Vitesco 一季，與 FY2025 不可比（FY2024 pro forma 營收 24,313）。
# 結構：README、In、Hist、Q（事業部門）、PL（三情境：命題模組→事業部門營收→標準損益表→現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 不讀 Siliconomics／Tokenomics／Wattonomics（算力線）。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261124"
NAME = f"{DATE}_舍弗勒個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032

# ---------------- 出處 ----------------
PR25 = "公司新聞稿「Schaeffler reports solid results for 2025」（2026-03-03）https://www.mynewsdesk.com/fi/schaeffler/pressreleases/schaeffler-reports-solid-results-for-2025-3435255"
SAI = "stockanalysis.com Schaeffler 損益表（資料商；頁面 2026-11 下旬）https://stockanalysis.com/quote/etr/SHA0/financials/income-statement/"
SAS = "stockanalysis.com Schaeffler statistics（資料商；價格 2026-11-23 17:37 CET，頁面更新 2026-11-24）https://stockanalysis.com/quote/etr/SHA0/statistics/"
Q3N = "公司 2026 年前九個月業績新聞稿（2026-11 初；經 industrial-production.de／computer-automation.de／menafn 轉載，未取得公司原文）https://www.computer-automation.de/robotics-en/financial-results-for-q3-2026-schaeffler-cuts-3-460-jobs-in-europe.htm"
YCL = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
BUND = "德國 10 年期公債殖利率：2026-11-09 約 3.26%（Trading Economics 新聞頁 https://tradingeconomics.com/germany/government-bond-yield/news/600706；11-20 值未取得；逾期）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"
L6 = "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx"
L4 = "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx"
NXX = "physical_ai/nexteer/dist/（耐世特個股模型 In 頁同業 P/E 與頁面報價）"

# ---------------- In 列 ----------------
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

REG = [("pc", "動力傳動與底盤（Powertrain & Chassis）"), ("em", "電動化（E-Mobility）"), ("vl", "車輛生命週期（Vehicle Lifetime Solutions，售後）"),
       ("bi", "軸承與工業（Bearings & Industrial Solutions）"), ("ot", "其他與調節（含人形、國防等成長領域；Derived）")]
sec("A. FY2025 實績：營收依事業部門（公司揭露；『其他與調節』＝合併營收 − 四部門合計）")
for (k, n), v in zip(REG[:4], (8900, 5015, 3038, 6368)):
    inp(f"a25_{k}", f"營收 FY2025：{n}", "€M", None, v, None, "Verified", PR25 + " 事業部門表", "")
inp("a25_ot", f"營收 FY2025：{REG[4][1]}", "€M", None, "=23492-(8900+5015+3038+6368)", None, "Derived", PR25 + "：合併營收 23,492 減四部門合計", "成長領域（人形、國防）另列於『Others』部門，公司未拆出金額")

sec("B. FY2026 錨點：公司指引（Interested-party；2026 年前九個月公告確認全年指引）")
for (k, n), v in zip(REG[:4], (8300, 5500, 3200, 6450)):
    inp(f"g26_{k}", f"公司指引中點：FY2026 營收 {n}", "€M", None, v, None, "Interested-party", PR25 + "：指引區間 P&C 8.0–8.6bn、E-Mobility 5.2–5.8bn、VLS 3.1–3.3bn、B&I 6.2–6.7bn，取中點（Derived）", "區間寬度 ±3–4%", CO)
inp("g26_rev", "公司指引中點：FY2026 合併營收", "€M", None, 23500, None, "Interested-party", PR25 + "：22.5–24.5bn（2026 前九個月公告確認）", "前九個月營收 17,481（去年同期 17,672）→ Q4 隱含約 6,019（Derived；Q3 5,814）", CO)
inp("g26_ot", f"FY2026 營收：{REG[4][1]}", "€M", None, "=g26_rev-(g26_pc+g26_em+g26_vl+g26_bi)", None, "Derived", "合併指引中點減四部門中點", "")
inp("ebitm_26", "FY2026 EBIT 利益率（特殊項目前）", "%", 0.035, 0.045, 0.055, "Interested-party", PR25 + "：指引 3.5–5.5%；前九個月 4.5%、Q3 4.2%（" + Q3N + "）", "基準取前九個月實績 4.5%（亦為指引中點）；低／高＝指引區間", CO)
inp("gmr_26", "FY2026 毛利率", "%", 0.18, 0.188, 0.195, "Assumed", "FY2025 毛利率 4,322 ÷ 23,492 ＝ 18.4%（Derived，" + SAI + "）；公司未給指引", "FY2023 21.5%、FY2024 21.1%（資料商）；基準略高於 FY2025（重組節省與價格），低／高為 Assumed")
inp("spi_26", "FY2026 特殊項目（重組、減損；EBIT 層級，負值）", "€M", -700, -480, -300, "Interested-party", Q3N + "：重組一次性成本約 480；前九個月特殊項目稅後 201、Q3 電動化減損與準備約 220–240", "基準取公司揭露的 480；低／高為 Assumed（範圍 300–700）", CO)
inp("fin_26", "FY2026 業外：金融收益 − 費用（含利息）", "€M", -380, -290, -220, "Assumed", "FY2025：稅前 −20 − 營業利益 190 ＝ −210；FY2024 −767（含合併相關，Derived，" + SAI + "）", "基準 −290（利息費用 FY2025 274 與 FY2024 275，" + SAI + "）；Assumed")
inp("tax_26", "FY2026 有效稅率", "%", 0.30, 0.35, 0.45, "Assumed", "FY2023 41.7%（240÷575）；FY2025 稅額 372 對稅前 −20（不可比，虧損國家無稅盾）（Derived）", "基準 35%；Assumed")
inp("nci_c", "非支配股東損益（每年）", "€M", 10, 35, 50, "Derived", "FY2023 26、FY2025 32（= 稅後 − 歸屬母公司淨利，Derived，" + SAI + "）", "")
inp("g_anc", "FY2026 實績相對公司指引中點（營收，四部門同比例）", "%", -0.03, 0.0, 0.03, "Assumed", "基準＝指引中點（前九個月已實績，偏離只來自 Q4）", "低／高 ±3%")
inp("da_26", "FY2026 折舊攤銷 ÷ 營收（基準欄）", "%", None, 0.05, None, "Derived", "FY2025 EBITDA（特殊項目前）≈ 淨財務負債 4,915 ÷ 2.1 ＝ 2,340，減 EBIT 936 ＝ D&A 約 1,404（6.0%）（Derived，" + PR25 + "）", "基準取 5.0% 偏保守（長期應向資本支出收斂）；高案＝6.0% 實績推算")

sec("C. 命題模組一：汽車與工業循環 → 核心營收（四部門＋其他同一成長率）")
GTXT = {
    "g27": ("FY2027", {"A": (0.04, "樂觀：歐洲車廠量回升、售後與工業成長（B&I 2025 +0.7%、VLS +5.0%）"),
                       "B": (0.02, "基準：FY2026 指引中點持平（23,500 對 FY2025 23,492）；前九個月匯率中性 −0.4%、Q3 −2.0%"),
                       "C": (-0.03, "保守：歐洲車廠中國專案量減（公司於 Q3 認列減損的原因）延續；P&C 2025 −5.2% 為類比")}),
    "g28": ("FY2028", {"A": (0.05, ""), "B": (0.025, "基準：名目成長"), "C": (-0.03, "保守：汽車循環下行")}),
    "gm": ("FY2029–FY2030", {"A": (0.04, ""), "B": (0.03, "基準：售後與工業占比提高；工業機器人裝機 2026→2030 年複合 6.9%（線 3）僅類比"), "C": (0.0, "")}),
    "glt": ("FY2031–FY2032", {"A": (0.03, ""), "B": (0.025, "長期：名目需求"), "C": (0.01, "")}),
}
GRNG = {"g27": (-0.02, 0.05), "g28": (-0.02, 0.05), "gm": (0.0, 0.05), "glt": (0.01, 0.04)}
for f, (yl, d) in GTXT.items():
    for s in SC:
        v, note = d[s]
        lo, hi = GRNG[f] if s == "B" else (None, None)
        inp(f"{f}_{s}", f"情境 {s}：核心營收年增 {yl}（五列同一成長率）", "%/年", lo, v, hi, "Assumed", note or "A／C 以 B ± 約 1.5–5pt 設定",
            "事業部門別成長差異無證據（不分化）；區間見情境 A／B／C 三列" if s == "B" else "")

sec("C2. 命題模組二：人形機器人致動器（地區與部門未分）＝行星滾柱絲杠市場（線 6 下游取數）× Schaeffler 份額 × 爬坡 ÷ EURUSD")
for yy, rr in ((2026, 12), (2030, 13), (2032, 14)):
    for s, col in (("A", "D"), ("B", "C"), ("C", "B")):
        v = {("A", 2026): 731.5, ("B", 2026): 298.571428571429, ("C", 2026): 89.5714285714286,
             ("A", 2030): 6545.958375, ("B", 2030): 1558.56151785714, ("C", 2030): 233.784227678571,
             ("A", 2032): 13794.2435339844, ("B", 2032): 3378.18208995536, ("C", 2032): 337.818208995536}[(s, yy)]
        inp(f"m{str(yy)[2:]}_{s}", f"情境 {s}：行星滾柱絲杠市場 {yy}（線 6；{'高' if s == 'A' else '基準' if s == 'B' else '低'}需求情境）", "百萬美元", None, v, None, "Derived",
            f"{L6} 下游取數!{col}{rr}（基準單價口徑＝覓途；五線歸因，全來自人形）", "線 6 需求情境：A＝高（D 欄）、B＝基準（C 欄）、C＝低（B 欄）；年間以幾何插值（Derived）")
for s, v, lo, hi in (("A", 0.25, None, None), ("B", 0.10, 0.03, 0.20), ("C", 0.02, None, None)):
    inp(f"shr_{s}", f"情境 {s}：Schaeffler 在人形致動器（以行星滾柱絲杠市場為可及市場代理）的份額", "%", lo, v, hi, "Assumed",
        f"公司說法：2030 年前人形相關訂單達數億歐元，目標拿下可服務 BOM（約整機 50%）的約 10%（{L6} 追蹤指標!D13；Interested-party，法說 2026-05-05）；線 6 把 Schaeffler（含 Ewellix）列為歐系現有技術領先者（份額!B11）",
        "份額數據未找到；基準 10% 對應 FY2030 約 1.4 億歐元營收（與『數億歐元訂單』同量級，訂單不等於營收）", CO)
inp("rp27", "人形營收爬坡：FY2027 占穩態比例（線性升到 FY2030＝100%）", "%", 0.0, 0.25, 0.5, "Assumed", "比照上銀模型爬坡 2027 25% → 2030 100%（physical_ai/hiwin In 頁）", "")
inp("fx", "EURUSD（人形市場換算；每歐元美元數）", "US$/€", 1.0, 1.096972, 1.2, "Verified", YCL + "：EURUSD=X 2026-11-20 收盤 1.096972", "低／高為 Assumed")

sec("D. 損益表結構（標準損益表；IFRS）")
inp("mm_gp", "邊際毛利率（Δ毛利 ÷ Δ營收，相對 FY2026）", "x", 0.12, 0.20, 0.30, "Assumed",
    "FY2025 毛利率 18.4%；FY2023 21.5%（資料商）；基準取高於平均（量增時固定成本稀釋＋重組節省）", "減收時同率回落；E-Mobility 虧損縮小的效應已含在此")
for s, v, lo, hi in (("A", 0.005, None, None), ("B", 0.0, -0.01, 0.005), ("C", -0.01, None, None)):
    inp(f"gmadj_{s}", f"情境 {s}：FY2027 起毛利率額外調整（價格、關稅、中國競爭）", "pt", lo, v, hi, "Assumed", "FY2025 事業部門 EBIT 利益率：P&C 10.5%、VLS 14.8%、B&I 7.5%、E-Mobility −16.0%（" + PR25 + "）", "")
inp("sga_el", "銷管研發費用對營收的彈性（費用年增 ÷ 營收年增）", "x", 0.3, 0.5, 0.8, "Assumed", "未找到逐年拆分；工業公司研發與管理費用多為固定", "FY2025 費用含特殊項目 746（Derived），已在 FY2025 列移出")
inp("sav", "重組節省：2030 年起年化（EBIT 層級）", "€M", 0, 300, 300, "Interested-party", Q3N + "：預期自 2030 年起每年改善約 300（對應裁撤 3,460 人、16 個據點，歐洲 3,460 人其中德國 2,660）", "公司宣稱；實現度見 sav_k", CO)
inp("sav_k", "重組節省實現度（扣除再投資、價格讓利）", "x", 0.0, 0.6, 1.0, "Assumed", "Assumed；ramp：FY2027 25% → FY2030 100%", "")
inp("oth_op", "其他營業損益淨額（特殊項目、其他）FY2027 起", "€M", -300, -50, 0, "Assumed", "FY2025 特殊項目 −746、FY2026 −480；續做重組的尾部費用（Assumed）", "")
inp("nonop", "業外：金融收益 − 金融費用（FY2027 起）", "€M", -380, -280, -200, "Assumed", "FY2025 −210、FY2026 假設 −290；淨財務負債 5,296（2026-09-30）× 約 5% 利率（Derived）", "")
inp("tax_w", "有效稅率（FY2027 起）", "%", 0.27, 0.32, 0.38, "Assumed", "德國法定約 30%；FY2023 41.7%（Derived）", "")

sec("E. 現金流與股東回饋")
inp("fc_da", "折舊攤銷 ÷ 營收（FY2027 起）", "%", 0.045, "=da_26", 0.06, "Derived", "見 da_26：FY2025 推算 6.0%；基準取 5.0%（Assumed）", "")
inp("fc_cx", "資本支出 ÷ 營收（FY2027 起）", "%", 0.04, 0.045, 0.06, "Assumed", "FY2025 資本支出 974 ÷ 23,492 ＝ 4.1%（" + PR25 + "）；FY2024 pro forma 1,460（6.0%）", "")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.10, 0.20, 0.30, "Assumed", "未取得資產負債表分項；汽車零件業約 20% 為類比", "")
inp("rcx_27", "重組現金支出（FY2027；計入營運資金列）", "€M", 100, 250, 400, "Assumed", "FY2025 自由現金流含 242 重組與整合現金流出；FY2026 指引含『顯著重組現金流出』（" + PR25 + "）", "FY2026 的特殊項目已在營業利益扣除，故只加 FY2027–28 尾款")
inp("rcx_28", "重組現金支出（FY2028）", "€M", 0, 100, 200, "Assumed", "同上", "")
inp("doe", "股利 ÷ 前一年底權益（股利政策代理）", "%", None, 0.09, None, "Derived", "FY2025 每股 0.30（提案）× 945 ÷ 權益約 3,123 ＝ 9.1%（Derived）", "只用於收支關卡（股利＝比率 × 前一年底權益）；不進 DCF", CO)
inp("dps_26", "FY2026 年內支付之每股股利（FY2025 年度，提案值）", "€", None, 0.30, None, "Interested-party", PR25 + "：提案 0.30（上年 0.25）", "", CO)
inp("eq_26", "FY2026 年底歸屬母公司權益（2026-09-30 權益 ＋ Q4 淨利）", "€M", None, "=5296/1.696+130", None, "Derived", Q3N + "：淨財務負債 5,296、槓桿 gearing 169.6% → 權益約 3,123；Q4 淨利以 130 假設（Assumed）", "")

sec("F. 估值輸入（評價日 2026-12-31；資產負債表 2026-09-30）")
inp("px", "股價（2026-11-20 收盤）", "€", None, 4.83, None, "Verified", YCL + "：SHA0.DE 2026-11-20 收盤 4.83", "未還原股利；11-13 收盤 5.54（同檔）。2026-11-23 價 5.07（" + SAS + "）")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("sh_out", "流通股數（資料商 944.88 百萬）", "百萬股", None, 944.88, None, "Derived", SAS, "公司原文未取得；FY2025 加權股數 945（" + SAI + "）；庫藏股、股票報酬未取得，視為 0")
inp("sh_v", "評價日股數", "百萬股", None, "=sh_out", None, "Derived", "＝流通股數", "")
inp("cash_v", "現金及約當現金（資料商）", "€M", None, 2280, None, "Derived", SAS + "：現金 2.28B", "")
inp("debt_v", "借款（淨財務負債 5,296 ＋ 現金 2,280）", "€M", None, "=5296+2280", None, "Derived", Q3N + "：2026-09-30 淨財務負債 5,296（2025 年底 4,915；槓桿 2.3x）", "資料商總負債 8.06B 減現金 2.28B ＝ 5.78B，與公司淨財務負債差 約 484（疑為租賃，未查核）")
inp("lease_v", "額外租賃負債代理（資料商與公司淨財務負債差）", "€M", 0, 0, 484, "Assumed", "見 debt_v 說明", "基準 0（假設已含在公司淨財務負債內）；高案 484")
inp("pens_v", "退休金負債（淨額；未找到本期值）", "€M", 0, 800, 1500, "Assumed", "未找到 2026-09-30 退休金準備；德國公司大型退休金義務為常識性存在（Assumed）", "區間 0–1,500；Vitesco 併入後可能更高——此為承重未取得資料")
inp("inv_v", "其他非流動金融資產", "€M", None, 0, None, "Assumed", "未取得", "")
inp("inv_k", "其他非流動金融資產計入比例", "x", None, 1.0, None, "Assumed", "基準無金額（0）", "")
inp("mi_v", "非支配股東權益", "€M", 0, 150, 300, "Assumed", "未取得（Vitesco 併入後有少數股東）；由非支配損益 32 以 5–10% 報酬倒推約 300–600（Derived 粗估），基準取較低值", "")
inp("stub_frac", "Stub 期（2026-09-30 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,9,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（德國 10 年期公債）", "%", None, 0.0326, None, "Verified", BUND, "")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "德國國家風險溢酬", "%", None, 0.0, None, "Assumed", "成熟市場 0（Assumed）", "")
inp("beta", "β（5 年，Blume 調整）", "x", "=E{r}-0.3", "=0.67*1.66+0.33", "=E{r}+0.3", "Derived", SAS + "：5 年 β 1.66（單一來源）", "Blume 0.67β＋0.33 ＝ 1.44；敏感度 ±0.3（Assumed）。β 偏高為單一資料商值")
inp("kd", "稅前負債成本", "%", 0.0476, "=rf+0.025", 0.0676, "Assumed", "Schaeffler 為投資級邊緣的歐系零組件廠；公債加 2.5pt（Assumed，未查公司債殖利率）", "")
inp("g_cap", "永續成長率上限", "%", 0.01, 0.02, 0.03, "Assumed", "歐元區名目成長較低：2%（敏感度 1–3%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("dps_tgt", "評價日到目標時點之每股股利", "€", None, "=dps_26", None, "Assumed", "FY2026 年度股利（2027 年股東會）假設同 FY2025 的 0.30", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable：同業前瞻 P/E）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", 0.7, 1.0, None, "Assumed", "1.00（未調整：缺 Schaeffler 對同業的歷史相對倍數）", "低案 0.7：歐洲汽車零件、高槓桿（淨負債 ÷ EBITDA 2.3x）、E-Mobility 虧損事業的折價；第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業前瞻 P/E（頁面值 × 2026-11-20 收盤 ÷ 頁面報價；取自耐世特模型 In 頁同業列；分母 EPS 不變）；歐洲同業收盤檔無報價，未納入")
PEERS = [("pe_mobis", "現代摩比斯 012330.KS（汽車零件）", 7.08, 352500, 323000, "2026-10-16"),
         ("pe_denso", "電裝 6902.T（汽車電子與熱管理）", 11.58, 1872.0, 1900.0, "2026-10-20"),
         ("pe_nsk", "NSK 6471.T（軸承、汽車零件）", 15.35, 1082.0, 892.5, "2026-09-18"),
         ("pe_tuopu", "拓普集團 601689.SS（底盤）", 26.89, 46.52, 48.11, "2026-09-30")]
for k, lab, pe, p0, p1, d in PEERS:
    inp(k, f"前瞻 P/E：{lab}（換算至 2026-11-20）", "x", None, f"={pe}*{p1}/{p0}", None, "Derived",
        f"{NXX}：{d} 頁面 forward P/E {pe}、頁面報價 {p0:,g}；2026-11-20 收盤 {p1:,g}（{YCL}）", "P/E 為股權口徑，不受租賃處理影響；會計口徑：K-IFRS、日本 IFRS、中國企業會計準則", DV)
inp("pe_nxt", "前瞻 P/E：耐世特 1316.HK（轉向系統；換算至 2026-11-20）", "x", None, "=9.22*4.18/4.515", None, "Derived", NXX + "：2026-11-13 forward P/E 9.22；收盤 4.515 → 4.18（" + YCL + "）", "", DV)
inp("pe_med", "同業前瞻 P/E 中位數（5 家；進 Comparable）", "x", 9.2, "=MEDIAN(pe_mobis,pe_denso,pe_nsk,pe_tuopu,pe_nxt)", 14.0, "Derived",
    "5 家中位數；未納入（報價日未取得 2026-11-20）：SKF、Continental、Valeo、Forvia、Aumovio、JTEKT", "低＝Schaeffler 自身現行 P/E 9.2（市場已定價的折價，Analogy）；高 14.0（Assumed）；少於約 8 家不做迴歸")
inp("pe_self", "Schaeffler 現行前瞻 P/E（換算至 2026-11-20）", "x", None, "=9.66*4.83/5.07", None, "Derived", SAS + "：價格 5.07 時 forward P/E 9.66；分母 EPS 不變 → 約 0.525 歐元（Derived）", "只作對照，不進模型", DV)
inp("cons_tp", "賣方平均目標價（11 位；資料商頁面 2026-11-23）", "€", None, 8.33, None, "Interested-party", SAS + "：平均 8.33、共識 Buy（11 位）", "只作對照，不進模型", DV)


def build(path, ov=None, sens=None, rev=None, bid="{BID}"):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 €M、股數百萬股、EPS 與股價歐元"
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
           ("shr", "Schaeffler 人形致動器份額"), ("gmadj", "FY2027 起毛利率調整")]
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
    H["A1"] = "Hist — 歷史財務（年度；€M）。FY2023＝Schaeffler 單獨；FY2024 只含 Vitesco 自 2024-10 起的一季（pro forma 營收 24,313，與 FY2025 不可比）；FY2025 全年合併。資料商（stockanalysis.com）損益表分類；公式只引用本頁（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "標記", "來源"])
    NA = "未取得"
    HR = [("h_rev", "營收", 16313, 18188, 23492, "Verified", SAI + "；FY2025 亦見 " + PR25),
          ("h_cogs", "營業成本", 12805, 14356, 19170, "Verified", SAI),
          ("h_gp", "毛利", 3508, 3832, 4322, "Verified", SAI),
          ("h_sga", "銷管與研發費用（FY2025 已移出特殊項目）", "=1797+768", "=2075+987", "=2550+1582-(936-190)", "Derived", SAI + "（SG&A＋R&D）；FY2025 移出特殊項目 746（＝EBIT 特殊項目前 936 − 營業利益 190，" + PR25 + "）"),
          ("h_oth", "其他營業損益淨額（FY2025＝特殊項目）", 0, 0, "=190-936", "Derived", "FY2023–24 特殊項目未取得（含在費用內）；FY2025＝營業利益 190 − EBIT 特殊項目前 936"),
          ("h_op", "營業利益（特殊項目後）", 943, 770, 190, "Verified", SAI),
          ("h_fin", "業外：金融收益 − 金融費用（稅前 − 營業利益）", "=575-943", "=3-770", "=-20-190", "Derived", SAI),
          ("h_pbt", "稅前利益", 575, 3, -20, "Verified", SAI),
          ("h_tax", "所得稅", 240, 608, 372, "Verified", SAI),
          ("h_cont", "稅後淨利（含非支配）", "=575-240", "=3-608", "=-20-372", "Derived", "稅前 − 所得稅"),
          ("h_disc", "停業單位損益", 0, 0, 0, "Assumed", "未見停業單位（資料商）"),
          ("h_ni", "歸屬母公司淨利", 309, -632, -424, "Verified", SAI + "；" + PR25),
          ("h_nci", "非支配股東損益（稅後 − 歸屬母公司淨利）", "=575-240-309", "=3-608+632", "=-20-372+424", "Derived", "Derived"),
          ("h_eps", "基本 EPS（歐元）", 0.46, -0.86, -0.45, "Verified", SAI + "；FY2025 亦見 " + PR25),
          ("h_epsc", "基本 EPS（持續營業，歐元；無停業單位＝同上）", 0.46, -0.86, -0.45, "Verified", SAI),
          ("h_da", "折舊攤銷（FY2025 由 EBITDA 倍數推算）", NA, NA, "=4915/2.1-936", "Derived", PR25 + "：淨財務負債 4,915 ÷ 2.1x ＝ EBITDA 2,340，減 EBIT 936"),
          ("h_cx", "資本支出", NA, NA, 974, "Verified", PR25 + "（FY2024 pro forma 1,460，不可比）")]
    for i, (k, lab, a, b, c, tag, src) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b, c)):
            cc = H.cell(i, 2 + j, v)
            if isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("=")): cc.font = BLUE
        H.cell(i, 5, tag); H.cell(i, 6, src)
        dn(k + "_23", f"Hist!$B${i}"); dn(k + "_24", f"Hist!$C${i}"); dn(k + "_25", f"Hist!$D${i}")
    H.column_dimensions["A"].width = 44
    HRW = {k: 3 + i for i, (k, *_x) in enumerate(HR)}

    # ---------------- Q（事業部門） ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 官方事業部門（公司 FY2025 新聞稿）：營收、特殊項目前 EBIT 利益率與 FY2026 指引（€M）。部門合計＋其他與調節＝合併營收"
    hdr(Q, 2, ["事業部門", "FY2025 營收", "FY2025 EBIT 利益率", "FY2026 營收指引低", "FY2026 營收指引高", "FY2026 EBIT 利益率指引（低）", "（高）", "標記", "來源", "來源誘因"])
    divs = [("動力傳動與底盤", 8900, 0.105, 8000, 8600, 0.10, 0.12), ("電動化", 5015, -0.16, 5200, 5800, -0.15, -0.13),
            ("車輛生命週期（售後）", 3038, 0.148, 3100, 3300, 0.135, 0.155), ("軸承與工業", 6368, 0.075, 6200, 6700, 0.07, 0.09)]
    for i, (n, r_, m_, lo_, hi_, ml, mh) in enumerate(divs, 3):
        Q.cell(i, 1, n)
        for j, v in enumerate((r_, m_, lo_, hi_, ml, mh)): Q.cell(i, 2 + j, v).font = BLUE
        Q.cell(i, 8, "Verified"); Q.cell(i, 9, PR25)
    Q.cell(7, 1, "四部門合計"); Q.cell(7, 2, "=SUM(B3:B6)"); Q.cell(7, 4, "=SUM(D3:D6)"); Q.cell(7, 5, "=SUM(E3:E6)"); Q.cell(7, 8, "Derived"); Q.cell(7, 9, "加總")
    Q.cell(8, 1, "其他與調節（含人形、國防成長領域）"); Q.cell(8, 2, "=B9-B7"); Q.cell(8, 8, "Derived"); Q.cell(8, 9, "合併營收減四部門合計")
    Q.cell(9, 1, "合併營收"); Q.cell(9, 2, 23492).font = BLUE; Q.cell(9, 4, 22500).font = BLUE; Q.cell(9, 5, 24500).font = BLUE; Q.cell(9, 6, 0.035).font = BLUE; Q.cell(9, 7, 0.055).font = BLUE
    Q.cell(9, 8, "Verified"); Q.cell(9, 9, PR25)
    Q.cell(10, 1, "檢查：四部門合計 ＋ 其他與調節 − 合併營收（應為 0）"); Q.cell(10, 2, "=ROUND(B7+B8-B9,3)"); Q.cell(10, 8, "Derived"); Q.cell(10, 9, "")
    Q.cell(12, 1, "2026 年前九個月（公司新聞稿，經轉載）").font = BOLD
    for i, (n, v) in enumerate((("營收（去年同期 17,672）", 17481), ("Q3 營收", 5814), ("EBIT 利益率（特殊項目前，9M）", 0.045), ("Q3 EBIT 利益率（特殊項目前）", 0.042),
                                ("歸屬母公司淨利 9M（資料來源品質低，轉載稿）", 11), ("自由現金流 9M（M&A 前；去年同期 47）", -40), ("淨財務負債 2026-09-30（2025 年底 4,915）", 5296)), 13):
        Q.cell(i, 1, n); Q.cell(i, 2, v).font = BLUE; Q.cell(i, 8, "Interested-party"); Q.cell(i, 9, Q3N); Q.cell(i, 10, CO)
    Q.column_dimensions["A"].width = 46

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（核心營收＋人形致動器）→ 事業部門營收 → 標準損益表（IFRS）→ 現金流（€M；FY2023–25 實績、FY2026 公司指引中點 ×(1＋g_anc)、FY2027 起模型）"
    PR = {}
    rr = 3
    KEYS = ([f"c_{k}" for k, _ in REG] + ["core", "hum", "rev", "chk1", "chk2", "gr", "cogs", "gp", "gm", "sga", "oth", "opex", "op", "opm", "fin", "pbt", "tax", "cont",
             "disc", "nci", "ni", "nim", "sh", "eps", "epsc", "nig", "da", "cx", "dwc", "fcff", "eq", "div", "fcfad", "mkt", "rp"])
    LAB = ([f"營收：{n}" for _, n in REG] +
           ["核心營收（五列合計）", "人形機器人致動器（命題子列）", "營收（合併）", "檢查：五列合計 − 核心營收", "檢查：核心＋人形 − 合併營收",
            "營收年增率", "營業成本", "毛利", "毛利率", "銷售及管理費用", "其他營業損益淨額（其他收益 − 其他費用 ＋ 持分法）", "營業費用（銷管費用 − 其他營業損益淨額）",
            "營業利益", "營業利益率", "業外：金融收益 − 金融費用（其他收入淨額）", "稅前利益", "所得稅", "稅後淨利（含非支配）", "停業單位損益",
            "非支配股東損益", "稅後淨利（歸屬母公司）", "淨利率", "平均股數（百萬股）", "EPS（基本，歐元）", "EPS（持續營業，歐元）", "淨利年增率",
            "折舊攤銷", "資本支出", "營運資金增加＋重組現金支出", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "歸屬母公司權益（年底）", "股利", "自由現金流扣股利（收支關卡用）",
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
                w("nim", f"={cell('ni')}/{cell('rev')}"); w("sh", f"={cell('ni')}/{cell('eps')}")
                w("hum", 0); w("rp", 0)
                if y == 2025:
                    for k, _ in REG: w(f"c_{k}", f"=a25_{k}")
                    w("core", f"=SUM({c}{R['c_pc']}:{c}{R['c_ot']})")
                    w("chk1", f"=ROUND(SUM({c}{R['c_pc']}:{c}{R['c_ot']})-{cell('core')},6)")
                    w("chk2", f"=ROUND({cell('core')}+{cell('hum')}-{cell('rev')},3)")
                    w("gr", f"={cell('rev')}/{prev('rev')}-1"); w("nig", '="N/M（FY2025 為虧損）"')
                elif y == 2024:
                    w("gr", f"={cell('rev')}/{prev('rev')}-1"); w("nig", '="N/M（FY2024 為虧損）"')
                else:
                    w("gr", '="未取得"'); w("nig", '="未取得"')
                continue
            if y == 2026:
                for k, _ in REG: w(f"c_{k}", f"=g26_{k}*(1+g_anc)")
                w("hum", 0); w("rp", 0)
                w("gp", f"=gmr_26*{cell('rev')}"); w("sga", f"={cell('gp')}-ebitm_26*{cell('rev')}"); w("oth", "=spi_26")
                w("fin", "=fin_26"); w("tax", f"={cell('pbt')}*tax_26"); w("disc", 0); w("nci", "=nci_c"); w("sh", "=sh_v")
                w("da", f"=da_26*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}"); w("eq", "=eq_26"); w("div", f"=dps_26*{cell('sh')}")
            else:
                gk = "g27" if y == 2027 else "g28" if y == 2028 else "gm" if y <= 2030 else "glt"
                for k, _ in REG: w(f"c_{k}", f"={prev('c_' + k)}*(1+{gk}_{s})")
                if y <= 2029: mk = f"=m26_{s}*(m30_{s}/m26_{s})^(({y}-2026)/4)"
                elif y == 2030: mk = f"=m30_{s}"
                elif y == 2031: mk = f"=m30_{s}*(m32_{s}/m30_{s})^0.5"
                else: mk = f"=m32_{s}"
                w("mkt", mk); w("rp", f"=MIN(1,rp27+(1-rp27)*({y}-2027)/3)")
                w("hum", f"={cell('mkt')}*shr_{s}*{cell('rp')}/fx")
                w("gp", f"=$F{R['gp']}+mm_gp*({cell('rev')}-$F{R['rev']})+gmadj_{s}*{cell('rev')}")
                w("sga", f"=$F{R['sga']}*(1+sga_el*({cell('rev')}/$F{R['rev']}-1))-sav*sav_k*MIN(1,({y}-2026)/4)"); w("oth", "=oth_op")
                w("fin", "=nonop"); w("tax", f"={cell('pbt')}*tax_w"); w("disc", 0); w("nci", "=nci_c"); w("sh", "=sh_v")
                w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}")
                w("div", f"=doe*{prev('eq')}"); w("eq", f"={prev('eq')}+{cell('ni')}-{cell('div')}")
            w("core", f"=SUM({c}{R['c_pc']}:{c}{R['c_ot']})")
            w("rev", f"={cell('core')}+{cell('hum')}")
            w("chk1", f"=ROUND(SUM({c}{R['c_pc']}:{c}{R['c_ot']})-{cell('core')},6)")
            w("chk2", f"=ROUND({cell('core')}+{cell('hum')}-{cell('rev')},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            w("cogs", f"={cell('rev')}-{cell('gp')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
            w("opex", f"={cell('sga')}-{cell('oth')}"); w("op", f"={cell('gp')}-{cell('opex')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            w("pbt", f"={cell('op')}+{cell('fin')}")
            w("cont", f"={cell('pbt')}-{cell('tax')}"); w("ni", f"={cell('cont')}+{cell('disc')}-{cell('nci')}")
            w("nim", f"={cell('ni')}/{cell('rev')}"); w("eps", f"={cell('ni')}/{cell('sh')}"); w("epsc", f"=({cell('cont')}-{cell('nci')})/{cell('sh')}")
            w("nig", f"=IF({prev('ni')}>0,{cell('ni')}/{prev('ni')}-1,\"N/M（前一年為虧損）\")")
            w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})+" + ("rcx_27" if y == 2027 else "rcx_28" if y == 2028 else "0"))
            w("fcff", f"={cell('op')}*(1-tax_w)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("fcfad", f"={cell('fcff')}-{cell('div')}")
        P.cell(b, 2, "FY2027 起＝前一年 ×(1＋核心年增)；FY2026＝指引中點 ×(1＋g_anc)")
        P.cell(R["hum"], 2, "市場 × 份額 × 爬坡 ÷ EURUSD；FY2026 以前未揭露（視為 0）")
        P.cell(R["chk2"], 2, "FY2025：其他與調節為倒推值")
        P.cell(R["gp"], 2, "FY2027 起＝FY2026 毛利 ＋ 邊際毛利率 × Δ營收 ＋ 毛利率調整 × 營收")
        P.cell(R["sga"], 2, "FY2027 起＝FY2026 ×(1＋彈性 × 營收相對 FY2026 增幅)")
        P.cell(R["tax"], 2, "FY2026＝稅前 × tax_26；FY2026 毛利＝gmr_26 × 營收、費用＝毛利 − EBIT 特殊項目前")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 ×(1 − tax_w)；FY2027–28 另扣重組現金支出（含在營運資金列）")
        P.cell(R["div"], 2, "FY2027 起＝股利比率 × 前一年底權益")
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
    D["A4"] = "股權市值 E（€M）"; D["B4"] = "=px*sh_v"
    D["A5"] = "負債 D（社債及借款＋租賃代理）"; D["B5"] = "=debt_v+lease_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金與投資（2026-09-30）＝現金 − 借款 − 租賃代理 − 非支配 − 退休金 ＋ 其他非流動金融資產 × inv_k"; D["B10"] = "=cash_v-debt_v-lease_v-mi_v-pens_v+inv_k*inv_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND(((px*sh_v)*(rf+beta*(erp+crp))+(debt_v+lease_v)*kd*(1-tax_w))/((px*sh_v)+debt_v+lease_v)+wacc_adj-wacc,9)"
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
    r_tp = one("加權目標價（歐元）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    r_cons = one("市場共識目標價（賣方 11 位平均；資料商 2026-11-23 頁面）", "=cons_tp", "只作對照（Interested-party）")
    r_px = one("現價（In 頁 px）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（€M；情境 A／B／C）")
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
    rc = [("FY2026 營收", f"=PL!F{PR['B']['rev']}", "=g26_rev", "公司指引中點（Interested-party）"),
          ("FY2026 EBIT（特殊項目前）", f"=PL!F{PR['B']['op']}-spi_26", "=ebitm_26*g26_rev", "指引利益率 4.5%（前九個月實績）× 指引營收中點"),
          ("FY2026 歸屬母公司淨利（模型 B）vs 前九個月實績 11（Q4 需約 120–140）", f"=PL!F{PR['B']['ni']}", 11.0, "前九個月淨利 11 百萬歐元（轉載稿，品質低）；差額＝Q4 隱含淨利"),
          ("FY2027 EPS（模型 B） vs 現價 ÷ 現行前瞻 P/E 隱含 NTM EPS", f"=PL!G{PR['B']['epsc']}", "=px/pe_self", "口徑：資料商 NTM（約 2026-11-23 頁面）vs 模型 FY2027；口徑不同，只看方向"),
          ("加權目標價 vs 賣方平均目標價", "=tp", "=cons_tp", "賣方 11 位平均（2026-11-23，只作對照）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_schaeffler.py 批次格點＋線性內插；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 60; RC.column_dimensions["E"].width = 70
    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_schaeffler.py 以 LibreOffice 批次重算每一案例）；結果為常數快照"
    base_tp = sens["base"]["tp"] if sens else None
    SE["A2"] = "快照基準值（加權，歐元）"; SE["B2"] = base_tp
    SE["C2"] = "本檔即時值（加權，歐元）"; SE["D2"] = "=tp"
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
           ("行星滾柱絲杠領先者", "6", "Schaeffler（含 Ewellix）、Rollvis、GSA、SKF", "—", "定性", L6 + " 份額!B11", "Interested-party", "shr_* 說明", "一致"),
           ("Schaeffler 人形目標：2030 年前訂單數億歐元、可服務 BOM（約整機 50%）約 10%", "6", "10%（shr_B 基準）", "%", "公司法說 2026-05-05", L6 + " 追蹤指標!D13", "Interested-party", "shr_B", "一致（本模型取公司目標為基準，但標 Assumed，營收≠訂單）"),
           ("人形用滾柱絲杠市場 2030（低／基準／高）", "4", "205.2／747.0／1,563.7", "百萬美元", "單價隨學習曲線；絲杠情形 3", L4 + " 下游取數!C72:E72", "Derived", "未進公式（對照）", "線 4 與線 6 不一致：線 6 基準 1,558.6 為線 4 的 2.1 倍（單價口徑不同）；本模型照所屬線 6，差異以 shr 區間吸收"),
           ("人形全球出貨 2030（低／基準／高）", "4", "20 萬／50 萬／89 萬", "台", "全球、含小型機", L4 + " 下游取數!C38:E38", "Derived", "未進公式（線 6 市場已含）", "一致（經線 6）"),
           ("全尺寸人形每台滾柱絲杠", "4", "12／14／14", "支／台", "全尺寸", L4 + " 下游取數!C26:E26", "Derived", "說明用", "一致"),
           ("線 6 滾珠絲杠／線控煞車需求（EMB）", "6", "不計入", "—", "汽車增量", L6 + " 下游取數!B23:D23", "Derived", "未進公式", "不一致（自設 0）：Schaeffler 汽車營收由 FY2026 指引與 g27–g28 成長率承擔，不另疊線控煞車增量，避免重複計算")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 30; RE.column_dimensions["F"].width = 60; RE.column_dimensions["I"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("歐洲車廠量與中國專案量減（Q3 認列電動化減損的原因）", "是（g27_C、g28_C）", "情境 C", "季度營收年增（匯率中性）、各部門訂單"),
             ("電動化事業持續虧損（FY2025 EBIT 利益率 −16%；指引 −15% 到 −13%）", "部分（mm_gp、gmr_26）", "併入毛利與特殊項目假設，未單獨建模", "電動化部門利益率、特殊項目"),
             ("重組成本與節省落空（裁 3,460 人；2030 年起年省約 300）", "是（sav_k、spi_26、rcx_27、rcx_28）", "實現度區間 0–1", "特殊項目、重組現金流出、員工數"),
             ("高槓桿（淨財務負債 5,296；2.3x）與退休金", "部分（pens_v、kd、WACC 權重）", "退休金金額未找到，以區間代入", "槓桿、自由現金流（9M −40）"),
             ("人形致動器份額為零或被旋轉關節、滾珠絲杠取代", "是（shr_C、Sens）", "人形子列占營收比小（見 Val）", "Tesla、Figure 供應鏈揭露；公司人形訂單"),
             ("關稅與匯率（EURUSD、人民幣）", "否", "僅人形子列用 EURUSD；其餘未逐年建模", "關稅公告、匯率"),
             ("同業倍數：歐洲同業未納入", "否", "歐洲同業 2026-11-20 報價未取得", "SKF、Continental、Valeo 前瞻 P/E")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 50

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([PR25, SAI, SAS, Q3N, YCL, BUND, L6 + "（下游取數、份額、追蹤指標）", L4 + "（下游取數）", NXX], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"Schaeffler 舍弗勒（SHA0.DE）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-24（2026 年前九個月公告）；成品日 2026-11-24；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.00")&" 歐元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：Schaeffler 是汽車傳動與工業軸承大廠，正由電動化虧損與歐洲重組中恢復獲利；股價由重組節省、電動化虧損收斂與汽車循環決定，人形機器人致動器（行星滾柱絲杠）是小型選擇權"
    RD["A4"] = (f'="在 FY2027 核心營收情境 B 年增 "&TEXT(g27_B,"0.0%")&"、邊際毛利率 "&TEXT(mm_gp,"0%")&"、重組節省 2030 年起 "&TEXT(sav*sav_k,"0")&" 百萬歐元（實現度 "&TEXT(sav_k,"0%")&"）、同業前瞻 P/E 中位數 "&TEXT(pe_med,"0.0")&" 倍的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 歐元（DCF "&TEXT(Val!B{r_dcft},"0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0")&" 百萬歐元）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 歐元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：季度 EBIT 利益率（特殊項目前）連兩季低於 3.5%（指引下緣）；電動化部門利益率未收斂到 −13% 以內；自由現金流（M&A 前）FY2026 低於指引下緣 100；FY2027 指引營收低於 FY2026 指引中點"
    RD["A7"] = "範圍：合併營收依四個事業部門加『其他與調節』，成長率同一；人形致動器為命題子列。不讀 Siliconomics／Tokenomics／Wattonomics（算力線）；自用 AI 讀法等規格條文不適用。HTML 展示層見 dist/ 同名 .html"
    RD["A8"] = f'="命題回覆：問：重組與電動化收斂能否支持高於現價的目標價？答：情境 B FY2028 EPS "&TEXT(PL!H{PR["B"]["epsc"]},"0.00")&" 歐元，現價只有其 "&TEXT(px/PL!H{PR["B"]["epsc"]},"0.0")&" 倍，低於同業中位數 "&TEXT(pe_med,"0.0")&" 倍；加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 歐元。問：人形機器人有多重要？答：情境 B FY2030 人形子列只占合併營收 "&TEXT(Val!B{r_hum},"0.0%")&"，結論不靠人形（模型輸出，非獨立證據）。"'
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
def lo_profile():
    d = os.path.join(os.environ.get("HOME", tempfile.gettempdir()), "lo_sch")
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
    d = tempfile.mkdtemp(prefix="sch_"); paths = []; VR = DR = None
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
    ("g_anc", "FY2026 實績相對公司指引"), ("ebitm_26", "FY2026 EBIT 利益率"), ("gmr_26", "FY2026 毛利率"), ("spi_26", "FY2026 特殊項目"),
    ("g27_B", "核心營收年增 FY2027（情境 B）"), ("g28_B", "核心營收年增 FY2028（情境 B）"), ("gm_B", "核心營收年增 FY2029–30（情境 B）"),
    ("glt_B", "核心營收年增 FY2031–32（情境 B）"), ("shr_B", "人形致動器份額（情境 B）"), ("rp27", "人形營收爬坡 FY2027"), ("fx", "EURUSD"),
    ("mm_gp", "邊際毛利率"), ("gmadj_B", "毛利率調整（情境 B）"), ("sga_el", "費用彈性"), ("sav", "重組節省（年化）"), ("sav_k", "重組節省實現度"),
    ("oth_op", "其他營業損益淨額"), ("nonop", "業外金融損益"), ("tax_w", "有效稅率"), ("fin_26", "FY2026 業外"), ("tax_26", "FY2026 有效稅率"), ("nci_c", "非支配損益"),
    ("fc_da", "折舊 ÷ 營收"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("rcx_27", "重組現金支出 FY2027"), ("rcx_28", "重組現金支出 FY2028"),
    ("lease_v", "租賃負債代理"), ("pens_v", "退休金負債"), ("mi_v", "非支配股東權益"), ("beta", "β"), ("kd", "負債成本"), ("erp", "股權風險溢酬"),
    ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重"), ("pe_med", "同業前瞻 P/E 中位數"), ("rel_k", "相對倍數係數")]
BETA_B = 0.67 * 1.66 + 0.33
_pe = [7.08 * 323000 / 352500, 11.58 * 1900.0 / 1872.0, 15.35 * 892.5 / 1082.0, 26.89 * 48.11 / 46.52, 9.22 * 4.18 / 4.515]
FXD = {"fc_da": 0.05, "kd": 0.0326 + 0.025, "pe_med": statistics.median(_pe)}


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
    for key, b0, lo_, hi_, note in (("g27_B", 0.02, -0.40, 0.60, "情境 B 的 FY2027 核心營收年增（其他不變，整本重算；情境 A／C 不動；之後年度在此基礎上成長）"),
                                    ("mm_gp", 0.20, -0.30, 1.00, "邊際毛利率（三情境同動）"),
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
    tmp = tempfile.mkdtemp(prefix="schf_"); raw = os.path.join(tmp, NAME)
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
    json.dump(dict(company="schaeffler", ticker="SHA0.DE", model_sheet="Q",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=v, unit="€M", source=PR25 + " 事業部門表")
                                      for n, l, v in (("動力傳動與底盤", "動力傳動與底盤", 8900), ("電動化", "電動化", 5015), ("車輛生命週期（售後）", "車輛生命週期（售後）", 3038), ("軸承與工業", "軸承與工業", 6368))],
                   sublines=[], check_labels=["檢查：四部門合計 ＋ 其他與調節 − 合併營收（應為 0）"],
                   note="官方事業部門四個加『其他與調節』（含人形、國防成長領域，公司未拆金額），Q 頁以合併營收倒推對帳；PL 營收列沿用同一分法，另有「檢查：五列合計 − 核心營收」「檢查：核心＋人形 − 合併營收」。"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)


def write_readme(path, VR, PR, DR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f2 = lambda x: f"{x:,.2f}"; pc = lambda x: f"{x*100:+.1f}%"
    inv = lambda k: I['E' + str([r for r in range(1, I.max_row + 1) if I['A' + str(r)].value == k][0])].value
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:8]
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {r['val'] if isinstance(r['val'], str) else format(r['val'], '.4f')} | {r['base']:.4g} | {r['note']} |" for r in revs)
    def yr(s, k, fmt="{:,.0f}", ys=range(2023, 2033)):
        out = []
        for y in ys:
            v = P[f"{COL[y]}{PR[s][k]}"].value
            out.append(fmt.format(v) if isinstance(v, (int, float)) else "—")
        return " | ".join(out)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f2(g(chr(66+i)+str(VR['dcfr'])))} | {f2(g(chr(66+i)+str(VR['cu'])))} | {f2(g(chr(66+i)+str(VR['tps'])))} | {g(chr(66+i)+str(VR['gate'])):,.0f} |" for i, s in enumerate(SC))
    rts = sorted(set(c['rt'] for c in sens['cases']))
    YS = range(2023, 2033)
    txt = f"""# physical_ai/schaeffler — Schaeffler 舍弗勒（SHA0.DE）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

法蘭克福證券交易所上市，12 月決算，IFRS；金額單位百萬歐元。本 README 的數字由 `build_schaeffler.py` 從重算後的成品讀出產生，不手寫。

## 未完成項目與限制（先看這裡）
- **退休金負債未找到**（德國公司的大額項目）：以 0／800／1,500 百萬歐元區間代入（Assumed），這是會動目標價的未取得資料；非支配股東權益、2026-09-30 資產負債表分項、折舊攤銷、資本支出指引也未取得（折舊由 EBITDA 倍數推算）。
- **2026 年前九個月數字來自轉載新聞稿與資料商，未取得公司原文**；前九個月歸屬母公司淨利（11 百萬歐元）來源品質低；公司 FY2026 的毛利率、淨利與特殊項目全年值未取得，皆為 Assumed。
- 同業只有 5 家（現代摩比斯、電裝、NSK、拓普、耐世特），皆為亞洲同業；SKF、Continental、Valeo、Forvia、Aumovio 與 JTEKT 在收盤檔無 2026-11-20 報價，**報價日未取得，未納入中位數**。
- β 只有單一資料商值 1.66；德國 10 年期公債殖利率用 2026-11-09 的約 3.26%（11-20 值未取得）。
- 人形致動器份額與營收無公開揭露，只有公司「2030 年前訂單數億歐元」說法；已標 Assumed，且人形子列占營收比重低，不是承重項。
- 事業部門成長率四列同一（無證據分化）；沒有建事業部門利益率與電動化虧損的逐項模型。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A8'].value}

**讀法：** 上檔主要來自三件事：(1) 現價只有模型 FY2027 EPS 的 {g(f"B{VR['pe1']}"):.1f} 倍，市場已把 Schaeffler 當作高槓桿的重組股（現行前瞻 P/E 約 {inv('pe_self'):.1f} 倍）；(2) DCF 假設 EBIT 利益率由 FY2026 的 4.5% 爬升，來自重組節省（公司宣稱 2030 年起每年約 300 百萬歐元）與毛利率回升；(3) Comparable 用同業 5 家中位數 {inv('pe_med'):.1f} 倍。結論**主要建立在 Assumed 輸入**（成長率、邊際毛利率、節省實現度、退休金、同業倍數），不是已揭露事實；人形子列（情境 B FY2030 占營收 {g(f"B{VR['hum']}")*100:.1f}%）不是承重項。

## 估值（評價日 2026-12-31；目標時點約 12 個月；歐元）
| 情境 | DCF（推到 12 個月後） | Comparable（P/E × FY2028 EPS） | 情境目標價 | 收支關卡（FY2026–30 FCF 扣股利，百萬歐元） |
|---|---|---|---|---|
{scn}

- 加權目標價 **{f2(tp)} 歐元**（DCF {f2(dcf)}、Comparable {f2(cmp_)}），現價 {f2(inv('px'))}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")} → 投資建議 **{g(f"B{VR['rec']}")}**。賣方共識目標價 {f2(inv('cons_tp'))}（11 位，2026-11-23，只作對照）。
- 方法：DCF 標準版 FCFF（{g(f"B{VR['wk']}")}，為 WACC／ke／g／RONIC）50%＋Comparable 50%（同業 5 家前瞻 P/E 中位數 {inv('pe_med'):.2f} 倍 × 各情境 FY2028 EPS）。相對倍數係數 1.00（未調整）；Sens 另列 0.7。
- 隱含前瞻 P/E（加權目標價 ÷ 情境 B FY2028 EPS）{g(f"B{VR['pe']}"):.1f} 倍；Schaeffler 現行前瞻 P/E {inv('pe_self'):.2f} 倍 ＝ 同業中位數的 {g(f"B{VR['pe2']}"):.2f} 倍。

## 營收與標準損益（情境 B，百萬歐元；FY2023–25 實績、FY2026 公司指引中點、FY2027 起模型）
| 列 | {" | ".join(f"FY{y}" for y in YS)} |
|---|{"---|" * len(YS)}
| 動力傳動與底盤 | {yr('B','c_pc')} |
| 電動化 | {yr('B','c_em')} |
| 車輛生命週期（售後） | {yr('B','c_vl')} |
| 軸承與工業 | {yr('B','c_bi')} |
| 其他與調節 | {yr('B','c_ot')} |
| 人形致動器（命題子列） | {yr('B','hum')} |
| 營收（合併） | {yr('B','rev')} |
| 毛利 | {yr('B','gp')} |
| 營業利益（特殊項目後） | {yr('B','op')} |
| 歸屬母公司淨利 | {yr('B','ni')} |
| EPS（歐元） | {yr('B','epsc','{:,.2f}')} |

- FY2024 只含 Vitesco 自 2024-10 起的一季（pro forma 營收 24,313），FY2025 年增不是同口徑；事業部門營收只有 FY2025 起；FY2023–24 無事業部門拆分。
- FY2026 的營業利益已扣特殊項目（基準 −480）；FY2027 起特殊項目假設 −50／年。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 核心營收年增（A／B／C）：FY2027 4%／2%／−3%；FY2028 5%／2.5%／−3%；FY2029–30 4%／3%／0%；FY2031–32 3%／2.5%／1%。
- 人形：線 6 行星滾柱絲杠市場（A＝高、B＝基準、C＝低）× Schaeffler 份額 25%／10%／2% × 爬坡（FY2027 25% → FY2030 100%）÷ EURUSD 1.097。
- 毛利：FY2026 毛利率 18.8%（Assumed），FY2027 起＝FY2026 毛利＋邊際毛利率 20% × Δ營收；重組節省 300 × 實現度 60%，FY2030 達滿；稅率 32%。
- 現金流：折舊 5.0%、資本支出 4.5%、營運資金 20% × 營收增量、FY2027–28 重組現金支出 250／100；股利＝9% × 前一年底權益。
- 折現：rf 3.26%、β {inv('beta'):.2f}、ERP 5%、kd＝rf＋2.5pt、g＝MIN(2%, rf)、RONIC＝WACC。股數 944.88 百萬；淨負債與退休金等合計 {-wb['DCF']['B10'].value:,.0f}（百萬歐元，含退休金 800 假設）。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| FY2027–28 核心成長（g27、g28） | 季度營收年增（匯率中性）、FY2027 指引（2027-03） |
| 邊際毛利率與重組節省（mm_gp、sav、sav_k） | 季度毛利率、EBIT 利益率（特殊項目前）、員工數與重組費用 |
| 退休金（pens_v） | 年報／季報退休金準備 |
| 同業 P/E（pe_med） | 歐洲同業前瞻 P/E（報價日 2026-11-20） |

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
- 線 6：行星滾柱絲杠市場 2026／2030／2032 低／基準／高（下游取數!B12:D14）→ 情境 C／B／A，照取；滾珠絲杠與線控煞車增量不計入（自設 0，避免與 FY2026 指引與成長率重複）。
- 線 4：人形用滾柱絲杠市場 2030 基準 747（下游取數 C72）與線 6 的 1,558.6 不一致（單價口徑），本模型照所屬線 6，差異由份額區間吸收。
- 份額：線 6 只記「歐系為現有技術領先者」；Schaeffler 10% 為公司目標（追蹤指標!D13），標 Assumed。

## 檔案
- `build_schaeffler.py`：建置＋整本重算 Sens／反解（LibreOffice 批次）＋產生本 README、`segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-24）：初版；現價 {f2(inv('px'))}（2026-11-20 收盤）；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


if __name__ == "__main__":
    main()

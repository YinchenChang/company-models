# Regal Rexnord（RRX，NYSE）個股模型 v0.1 — AI實體應用 Project（線 6 共用零組件：無框力矩馬達；人形用量另讀線 4）
# 用法：python3 physical_ai/regal_rexnord/build_rrx.py → 建置、整本重算 Sens 與反解（LibreOffice 批次）、輸出 dist/ 成品、README.md、sens.json、segments.json
# 會計年度＝曆年（12 月決算，美國 GAAP）。金額：百萬美元（US$M）；股數：百萬股；EPS、股價：美元。
# 結構：README、In、Hist、Q、PL（三情境：命題模組→分部營收→標準損益表→現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261124"
NAME = f"{DATE}_RegalRexnord個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2024, 2033)); COL = {y: "CDEFGHIJK"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2024 … K＝FY2032

Q326 = "公司新聞稿：Regal Rexnord Reports Third Quarter 2026 Financial Results（2026-11-03；Finviz 轉載）https://finviz.com/news/414042/regal-rexnord-reports-third-quarter-2026-financial-results"
Q425 = "公司 8-K 附件：2025 年第四季及全年財報新聞稿（2026-02-04）https://www.sec.gov/Archives/edgar/data/82811/000008281126000041/a4q2025earningsannouncement.htm"
FOX = "Fox Business RRX 報價頁（2026-10-23 更新；資料商）https://www.foxbusiness.com/quote?stockTicker=RRX"
YHO = "Yahoo Finance RRX 統計頁（資料商；5 年月 β 1.04）https://finance.yahoo.com/quote/RRX/key-statistics/"
MST = "Morningstar／Dow Jones：10-Year Treasury Yield Falls to 5.051%（2026-11-19 盤中）https://www.morningstar.com/news/dow-jones/202611198503/10-year-treasury-yield-falls-to-5051-data-talk"
YCL = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商（立場：中立彙整，但來源為賣方或第三方）"
L6 = "physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx"
L4 = "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx"
HWX = "physical_ai/hiwin/dist/20261123_上銀個股模型_v0.1.1.xlsx"
HDX = "physical_ai/harmonic_drive/dist/20261123_哈默納科個股模型_v0.1.1.xlsx"

ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

SEG = [("amc", "AMC 自動化與運動控制（Kollmorgen 等；不含無框馬達增量）"), ("ips", "IPS 工業動力傳動"), ("pes", "PES 電力效率方案（含資料中心）")]
SEGS = {"amc": "AMC", "ips": "IPS", "pes": "PES"}
sec("A. FY2025 實績與 9M'26 實績（Verified）")
for (k, n), v in zip(SEG, (1689.8, 2594.1, 1650.6)):
    inp(f"a25_{k}", f"營收 FY2025：{SEGS[k]}", "US$M", None, v, None, "Verified", Q425 + "：分部淨銷售", "FY2025 三分部合計 5,934.5 ＝ 合併營收")
for (k, n), v, y in zip(SEG, (1435.4, 1967.2, 1222.9), (0.187, 0.022, -0.042)):
    inp(f"m9_{k}", f"營收 9M'26：{SEGS[k]}", "US$M", None, v, None, "Verified", Q326 + "：9M 分部淨銷售", "")
    inp(f"m9g_{k}", f"營收 9M'26 年增（公司揭露）：{SEGS[k]}", "%", None, y, None, "Verified", Q326, "")
inp("m9_gp", "9M'26 毛利", "US$M", None, 1751.7, None, "Verified", Q326, "")
inp("m9_opx", "9M'26 營業費用（含處分損益；不含商譽減損）＝毛利 − 營業利益 − 減損", "US$M", None, "=1751.7-449.3-119.3", None, "Derived", Q326 + "：毛利 1,751.7、營業利益 449.3、Q3 商譽減損 119.3", "")
inp("imp26", "FY2026 商譽減損（AMC 自主移動方案；Q3 提列）", "US$M", None, 119.3, None, "Verified", Q326, "公司評估出售該業務；本模型不建出售（未取得價格）")
inp("m9_ni", "9M'26 歸屬母公司淨利", "US$M", None, 168.4, None, "Verified", Q326, "")
inp("g26_eps_lo", "公司指引：FY2026 GAAP 稀釋 EPS 下緣", "US$", None, 3.76, None, "Interested-party", Q326, "含 IEEPA 關稅退款 +0.58、商譽減損 −1.78（公司揭露）", CO)
inp("g26_eps_hi", "公司指引：FY2026 GAAP 稀釋 EPS 上緣", "US$", None, 4.06, None, "Interested-party", Q326, "", CO)
inp("g26_adj", "公司指引：FY2026 調整後稀釋 EPS 中點（10.60–10.90）", "US$", None, 10.75, None, "Interested-party", Q326, "調整後口徑排除收購無形資產攤銷等；本模型 Comparable 用自定義調整（見 Val 頁）", CO)
inp("q4g_amc", "FY2026 第四季年增：AMC", "%", 0.12, 0.20, 0.25, "Assumed", "Q3'26 年增 24.5%（有機 +25.2%）、9M +18.7%（" + Q326 + "）", "Q4'25＝FY2025 − 9M'25（9M'25 由 9M'26 與公司年增反推，Derived）")
inp("q4g_ips", "FY2026 第四季年增：IPS", "%", -0.05, -0.015, 0.02, "Assumed", "Q3'26 −1.9%、9M +2.2%（" + Q326 + "）", "")
inp("q4g_pes", "FY2026 第四季年增：PES", "%", -0.02, 0.03, 0.08, "Assumed", "Q3'26 +1.2%、9M −4.2%（" + Q326 + "）；資料中心訂單成長", "")
inp("q4_gm", "FY2026 第四季毛利率", "%", 0.36, 0.37, 0.38, "Assumed", "Q3'26 37.2%、9M 37.9%（" + Q326 + "）", "")
inp("q4_opx", "FY2026 第四季營業費用（不含減損）", "US$M", 385.0, 395.0, 410.0, "Assumed", "Q3'26 營業費用約 389.5（＝590.2 − 81.4 − 119.3，Derived）", "校準：使 FY2026 GAAP EPS 落在公司指引 3.76–4.06 內（見 Recon）")
inp("fin_26", "FY2026 淨利息與其他（＝− 淨利息率 × FY2025 年底借款）", "US$M", None, "=-nir*debt_25", None, "Derived", "FY2025 淨利息與其他 −328.3（" + Q425 + "）", "9M'26 利息未取得")
inp("tx26", "FY2026 有效稅率", "%", 0.22, 0.25, 0.28, "Derived", "反推：公司 GAAP EPS 指引中點 3.91 × 66.8 百萬股 ＝ 淨利 261；商譽減損部分不可抵稅", "FY2025 有效稅率 20.3%（" + Q425 + "）")
inp("dep_26", "FY2026 折舊（9M 111.2 × 4/3）", "US$M", None, "=111.2*4/3", None, "Derived", Q326, "")
inp("am_26", "FY2026 無形資產攤銷（9M 259.9 × 4/3）", "US$M", None, "=259.9*4/3", None, "Derived", Q326, "")
inp("cx_26", "FY2026 資本支出（9M 66.0 ＋ 第四季 30）", "US$M", 85.0, "=66+30", 120.0, "Assumed", Q326 + "：9M 66.0", "FY2025 97.7")
inp("dwc_26", "FY2026 營運資金與其他現金流出（校準 9M 營業現金流 411 低於獲利）", "US$M", 100.0, 250.0, 350.0, "Assumed", "9M'26 營業現金流 411.4 vs 9M'25 823.0（" + Q326 + "）", "調整後 EBITDA 9M 1,022 − 估計利息與稅後的差額（Derived，粗估）")
inp("sh_26", "FY2026 平均稀釋股數（Q3 調整後 66.8）", "百萬股", None, 66.8, None, "Verified", Q326, "")
inp("debt_q4", "FY2026 第四季淨償債（假設）", "US$M", 0.0, 100.0, 200.0, "Assumed", "9M 自由現金流 345.4；Q3 單季 193.7（" + Q326 + "）", "")

sec("B. 命題模組一：三分部營收成長（FY2027 起；各分部、各情境）")
GV = {"g27": {"amc": (0.14, 0.08, 0.0), "ips": (0.04, 0.02, -0.04), "pes": (0.10, 0.05, -0.03)},
      "g28": {"amc": (0.10, 0.06, -0.03), "ips": (0.04, 0.025, -0.02), "pes": (0.08, 0.04, 0.0)},
      "gm": {"amc": (0.08, 0.05, 0.02), "ips": (0.04, 0.03, 0.02), "pes": (0.06, 0.04, 0.02)},
      "glt": {"amc": (0.06, 0.04, 0.02), "ips": (0.03, 0.025, 0.02), "pes": (0.04, 0.03, 0.02)}}
GY = {"g27": "FY2027", "g28": "FY2028", "gm": "FY2029–FY2030", "glt": "FY2031–FY2032"}
GNOTE = {"amc": "AMC 9M'26 +18.7%（Q3 +24.5%）；基準假設回落至 FY2027 +8%（Assumed）；A／C 約 ±6pt",
         "ips": "IPS 9M'26 +2.2%、Q3 −1.9%；FY2024→FY2025 −0.2%（" + Q425 + "）；成熟傳動，成長貼近工業產值",
         "pes": "PES 9M'26 −4.2%、Q3 +1.2%；資料中心訂單（公司 2026-02 稱訂單 7.35 億美元，Interested-party）支撐 FY2027 回升；住宅與商用馬達疲弱為下緣"}
for f in GV:
    for sg, _n in SEG:
        for s, v in zip(SC, GV[f][sg]):
            lo, hi = (v - 0.04, v + 0.04) if s == "B" else (None, None)
            inp(f"{f}_{sg}_{s}", f"情境 {s}：{SEGS[sg]} 營收年增 {GY[f]}", "%/年", lo, v, hi, "Assumed", GNOTE[sg] if f == "g27" else "同上，逐年遞減趨向名目產值成長", "公司未給 FY2027 以後指引（未找到）" if s == "B" else "")

sec("C. 命題模組二：無框力矩馬達（人形、協作；線 6）→ AMC 增量營收＝Kollmorgen 份額 ×（市場 − 2026 市場）")
MK = {"A": (480.816, 827.598, 1461.82, 2634.127, 4813.731, 7480.625, 11661.807), "B": (329.814, 495.136, 762.111, 1195.584, 1901.636, 2975.172, 4684.238),
      "C": (224.856, 278.354, 353.345, 457.841, 602.87, 764.807, 975.746)}
MKR = {"A": 212, "B": 145, "C": 78}
for s in SC:
    for i, y in enumerate(range(2026, 2033)):
        inp(f"mk{str(y)[2:]}_{s}", f"情境 {s}：無框力矩馬達市場 {y}", "百萬美元", None, MK[s][i], None, "Derived",
            f"{L6} 計算!{'CDEFGHI'[i]}{MKR[s]}（五線歸因；{'高' if s == 'A' else '基準' if s == 'B' else '低'}需求情境）", "線 6 下游取數頁未列無框馬達；取自「計算」頁 FM 列（公司檔案 profiles/RRX.md 亦引同列）")
for s, v, lo, hi in (("A", 0.08, None, None), ("B", 0.04, 0.01, 0.10), ("C", 0.01, None, None)):
    inp(f"shr_{s}", f"情境 {s}：Kollmorgen 在無框力矩馬達的份額（增量部分）", "%", lo, v, hi, "Assumed",
        f"未取得：線模型份額頁列 Kollmorgen（RRX）為主要廠之一，份額「未取得」（{L6} 份額!F13）", "Maxon、鳴志等同列；公司未揭露無框馬達營收")

sec("D. 損益結構")
for s, v, lo, hi in (("A", 0.004, None, None), ("B", 0.002, -0.002, 0.006), ("C", -0.005, None, None)):
    inp(f"dgm_{s}", f"情境 {s}：FY2027–FY2030 每年毛利率變動（之後持平）", "pt/年", lo, v, hi, "Assumed",
        "毛利率：FY2024 36.3%、FY2025 37.4%、9M'26 37.9%（Derived，" + Q425 + "、" + Q326 + "）", "C：價格與關稅壓力")
inp("opx_el", "營業費用彈性（費用年增 ÷ 營收年增）", "x", 0.4, 0.6, 0.9, "Assumed", "FY2025 營收 −1.6%、營業費用 −1.3%（Derived，" + Q425 + "）", "半固定成本")
inp("tax_w", "有效稅率（FY2027 起）", "%", 0.18, 0.22, 0.26, "Assumed", "FY2024 20.0%、FY2025 20.3%（Derived，" + Q425 + "）", "")
inp("nir", "淨利息率（淨利息與其他 ÷ 年初借款）", "%", 0.055, 0.065, 0.075, "Derived", "FY2025 328.3 ÷ 4,788.7（年底借款）＝ 6.9%（" + Q425 + "）", "低案反映降息與再融資")

sec("E. 現金流與股東回饋")
inp("fc_dep", "折舊 ÷ 營收（FY2027 起）", "%", 0.022, 0.025, 0.030, "Derived", "FY2025 154.5 ÷ 5,934.5 ＝ 2.6%（" + Q425 + "）", "")
inp("am_f", "無形資產攤銷（FY2027 起，每年）", "US$M", 300.0, 335.0, 346.0, "Assumed", "FY2025 346.1、FY2026 約 346（" + Q425 + "）", "攤銷明細表未取得；假設逐步略降")
inp("fc_cx", "資本支出 ÷ 營收（FY2027 起）", "%", 0.014, 0.016, 0.025, "Derived", "FY2025 97.7 ÷ 5,934.5 ＝ 1.6%", "")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.08, 0.15, 0.25, "Assumed", "未取得營運資金明細", "")
inp("dps", "每股股利（年）", "US$", None, 1.40, None, "Interested-party", Q425 + "：FY2025 宣告 1.40（FY2024 同）", "FY2026 以後假設不變（Assumed）", CO)

sec("F. 估值輸入（評價日 2026-12-31；資產負債表 2026-09-30）")
inp("px", "股價（2026-11-20 收盤）", "US$", None, 152.350006, None, "Verified", YCL + "：RRX 2026-11-20 收盤 152.35", "未還原股利")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("sh_v", "評價日股數（Q3'26 調整後稀釋）", "百萬股", None, 66.8, None, "Verified", Q326, "流通股 66.6；選擇權與 RSU 已含於稀釋股數")
inp("cash_v", "現金（2026-09-30）", "US$M", None, 449.6, None, "Verified", Q326, "")
inp("debt_v", "借款總額（2026-09-30；長期 4,427.5＋一年內到期 24.7）", "US$M", None, 4452.2, None, "Verified", Q326, "公司揭露淨負債 4,001.0")
inp("debt_25", "借款總額（2025-12-31；4,764.6＋24.1）", "US$M", None, 4788.7, None, "Verified", Q425 + "；" + Q326, "")
inp("lease_v", "租賃負債代理", "US$M", 0.0, 0.0, 300.0, "Assumed", "未取得：10-Q 資產負債表未讀", "基準 0（偏樂觀）；高案 300")
inp("mi_v", "非支配股東權益", "US$M", None, 0.0, None, "Assumed", "未取得；9M 淨利與歸屬母公司淨利差 0.4（公司第 3 季新聞稿），金額小", "")
inp("stub_frac", "Stub 期（2026-09-30 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,9,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（美國 10 年期）", "%", 0.045, 0.0505, 0.055, "Derived", MST, "2026-11-19 盤中值；11-20 收盤未取得（11-17 收盤 5.10%）；殖利率處於 2007 年以來高位")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "國家風險溢酬（美國）", "%", None, 0.0, None, "Assumed", "美國 0", "")
inp("beta", "β（Blume 調整：兩來源 5 年 β 1.07、1.04）", "x", "=E{r}-0.3", "=0.67*MEDIAN(1.07,1.04)+0.33", "=E{r}+0.3", "Derived", FOX + "；" + YHO, "Blume 0.67β＋0.33；敏感度 ±0.3（Assumed）")
inp("kd", "稅前負債成本", "%", None, "=nir", None, "Derived", "＝淨利息率（含其他，粗估）", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("dps_tgt", "評價日到目標時點之每股股利", "US$", None, "=dps", None, "Assumed", "", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "同業僅 3 家，Comparable 可信度低，權重區間含 0.3")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺 RRX 對同業歷史相對倍數）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業前瞻 P/E（只納入報價日＝2026-11-20 者；取自同專案成品的換算值）")
PEERS = [("pe_fanuc", "Fanuc 6954（機器人、FA）", 21.1499, HWX + " Comps!F12（換算至 2026-11-20）"),
         ("pe_yask", "安川電機 6506（伺服、機器人）", 15.6739, HWX + " Comps!F13（換算至 2026-11-20）"),
         ("pe_nab", "Nabtesco 6268（精密減速機）", 13.7701, HDX + " Comp!F5（換算至 2026-11-20）")]
for k, lab, v, src in PEERS:
    inp(k, f"前瞻 P/E：{lab}", "x", None, v, None, "Derived", src, "原值：資料商賣方共識 EPS；P/E 為股權口徑。同為日本企業，與美國工業股估值水準可能不同（Analogy）", DV)
inp("pe_med", "同業前瞻 P/E 中位數（3 家；進 Comparable）", "x", 13.77, "=MEDIAN(pe_fanuc,pe_yask,pe_nab)", 21.15, "Derived",
    "3 家中位數；報價日未取得、不納入：SMC、THK、NSK、IKO、直得、上銀（報價日早於 11-20）；美國同業（Rockwell、Emerson、Parker 等）收盤檔無報價", "低＝Nabtesco；高＝Fanuc（Analogy）；少於約 8 家不做迴歸")
inp("pe_self", "RRX 現行前瞻 P/E（現價 ÷ 公司調整後 EPS 指引中點）", "x", None, "=px/g26_adj", None, "Derived", "現價 ÷ 10.75；公司調整後口徑，與 Comparable 的自定義調整 EPS 不同", "只作對照，不進模型", CO)
inp("cons_tp", "賣方平均目標價（13 位；2026-10-23 頁面，第 3 季財報前）", "US$", None, 231.10, None, "Interested-party", FOX, "只作對照，不進模型；來源間差異大（另見 Google Finance 約 262，2026-07）", DV)


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
    FAM = [(f"{f}_{sg}", f"{SEGS[sg]} 營收年增 {GY[f]}") for sg in SEGS for f in GV] + [("shr", "Kollmorgen 無框馬達份額")] + [(f"mk{yy}", f"無框力矩馬達市場 20{yy}（線 6）") for yy in range(26, 33)] + [("dgm", "每年毛利率變動 FY2027–30")]
    for f, lab in FAM:
        IN.cell(r, 1, f + "_scn"); IN.cell(r, 2, lab)
        for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR[f + '_' + s]}")
        IN.cell(r, 8, f"引用 In 頁 {f}_A／{f}_B／{f}_C"); r += 1
    IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重")
    for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR['w_' + s]}")
    IN.cell(r, 8, "引用 In 頁 w_A／w_B／w_C；合計＝1"); r += 1
    IN.column_dimensions["B"].width = 56; IN.column_dimensions["H"].width = 60

    # ---------------- Hist ----------------
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（年度；US$M；美國 GAAP）。FY2023 未取得（本輪未查 10-K）。FY2024 含已處分的 Industrial Systems 分部營收 157.8"
    hdr(H, 2, ["項目", "FY2024", "FY2025", "標記", "來源"])
    HR = [("h_rev", "營收", 6033.8, 5934.5), ("h_cogs", "營業成本", 3842.8, 3716.7), ("h_gp", "毛利", 2191.0, 2217.8),
          ("h_opx", "營業費用", 1552.5, 1532.5), ("h_los", "處分事業損失", 8.5, 4.5), ("h_op", "營業利益", 630.0, 680.8),
          ("h_fin", "淨利息與其他（利息費用 − 利息收入 ＋ 其他費用，取負）", "=-(399.7-18.8+1.1)", "=-(349.2-23.7+2.8)"),
          ("h_pbt", "稅前利益", 248.0, 352.5), ("h_tax", "所得稅", 49.6, 71.7), ("h_ni", "淨利", 198.4, 280.8),
          ("h_sh", "稀釋股數（百萬）", 66.7, 66.6), ("h_eps", "稀釋 EPS（US$）", 2.94, 4.20),
          ("h_amc", "AMC 營收", 1633.8, 1689.8), ("h_ips", "IPS 營收", 2598.1, 2594.1), ("h_pes", "PES 營收", 1644.1, 1650.6), ("h_isd", "Industrial Systems（已處分）營收", 157.8, 0),
          ("h_dep", "折舊（現金流量表）", 165.3, 154.5), ("h_am", "無形資產攤銷", 346.5, 346.1), ("h_cx", "資本支出", 109.5, 97.7),
          ("h_fcf", "自由現金流（營業現金流 − 資本支出）", 499.9, 893.1), ("h_debt", "借款總額（年底）", "未取得", 4788.7), ("h_div", "每股股利（宣告）", 1.40, 1.40)]
    for i, (k, lab, a, b) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b)):
            cc = H.cell(i, 2 + j, v)
            if isinstance(v, (int, float, str)) and v != "未取得": cc.font = BLUE
        H.cell(i, 4, "Verified"); H.cell(i, 5, Q425)
        dn(k + "_24", f"Hist!$B${i}"); dn(k + "_25", f"Hist!$C${i}")
    H.column_dimensions["A"].width = 52

    # ---------------- Q ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 官方報告分部（AMC、IPS、PES）9M'26 與 Q3'26 實績，及 9M'25（由公司揭露年增反推）。分部合計對合併營收核對（US$M）"
    hdr(Q, 2, ["項目", "9M'25（Derived）", "9M'26", "標記", "來源", "Q3'26 營收", "Q3'26 年增（公司）", "Q3'26 調整後 EBITDA 率", "來源誘因"])
    qd = [("AMC 營收", "amc", 500.6, 0.245, 0.192), ("IPS 營收", "ips", 649.5, -0.019, 0.260), ("PES 營收", "pes", 437.9, 0.012, 0.197)]
    for i, (lab, k, q3, g3, mg) in enumerate(qd, 3):
        Q.cell(i, 1, lab); Q.cell(i, 2, f"=m9_{k}/(1+m9g_{k})"); Q.cell(i, 3, f"=m9_{k}"); Q.cell(i, 4, "Verified"); Q.cell(i, 5, Q326 + "：分部淨銷售（9M'25 由年增反推）")
        Q.cell(i, 9, CO); Q.cell(i, 6, q3).font = BLUE; Q.cell(i, 7, g3).font = BLUE; Q.cell(i, 8, mg).font = BLUE
    Q.cell(6, 1, "分部合計"); Q.cell(6, 2, "=SUM(B3:B5)"); Q.cell(6, 3, "=SUM(C3:C5)"); Q.cell(6, 6, "=SUM(F3:F5)"); Q.cell(6, 4, "Derived")
    Q.cell(7, 1, "合併營收（公司揭露）"); Q.cell(7, 2, 4411.2).font = BLUE; Q.cell(7, 3, 4625.5).font = BLUE; Q.cell(7, 6, 1588.0).font = BLUE; Q.cell(7, 4, "Verified"); Q.cell(7, 5, Q326)
    Q.cell(8, 1, "檢查：分部合計 − 合併營收（四捨五入差 < 0.5）"); Q.cell(8, 2, "=ROUND(B6-B7,1)"); Q.cell(8, 3, "=ROUND(C6-C7,1)"); Q.cell(8, 6, "=ROUND(F6-F7,1)"); Q.cell(8, 4, "Derived")
    Q.cell(9, 1, "註：官方分部即 PL 頁分部（AMC 在 PL 頁另加無框馬達增量命題子列）；Q3 分部調整後 EBITDA 率為公司揭露，其他季度未取得")
    Q.column_dimensions["A"].width = 50

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（三分部成長＋無框馬達增量）→ 分部營收 → 標準損益表（美國 GAAP）→ 現金流（US$M；FY2024–25 實績、FY2026＝9M 實績＋第四季假設、FY2027 起模型）"
    PR = {}; rr = 3
    KEYS = [f"c_{k}" for k, _ in SEG] + ["core", "hum", "rev", "chk1", "chk2", "gr", "cogs", "gp", "gm", "sga", "oth", "opex", "op", "opm", "fin", "pbt", "tax", "ni", "nim", "sh", "eps",
            "nig", "am", "epsc", "da", "cx", "dwc", "fcff", "debt", "div", "fcfad", "mkt", "rp"]
    LAB = [f"營收：{n}" for _, n in SEG] + ["分部營收合計（不含無框馬達增量）", "無框力矩馬達增量（命題子列；列於 AMC）", "營收（合併）", "檢查：分部合計 − 核心營收", "檢查：核心＋無框增量＋已處分分部 − 合併營收",
            "營收年增率", "營業成本", "毛利", "毛利率", "營業費用（不含減損與處分損益）", "商譽減損與處分事業損失", "營業費用合計", "營業利益", "營業利益率", "業外：淨利息與其他",
            "稅前利益", "所得稅", "淨利", "淨利率", "稀釋股數（百萬股）", "EPS（GAAP 稀釋，US$）", "淨利年增率", "無形資產攤銷", "調整後 EPS（GAAP EPS ＋ 攤銷 ×(1 − 稅率) ÷ 股數；Comparable 用）",
            "折舊與攤銷", "資本支出", "營運資金增加", "FCFF（NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加）", "借款總額（年底）", "股利", "自由現金流扣股利（FCFF 扣稅後利息、股利；收支關卡用）",
            "參考：無框力矩馬達市場（百萬美元，線 6）", "參考：2026 年市場基期"]
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
                for k, _ in SEG: w(f"c_{k}", f"=h_{k}_{yy}")
                w("hum", 0); w("rp", 0)
                w("rev", f"=h_rev_{yy}"); w("cogs", f"=h_cogs_{yy}"); w("gp", f"=h_gp_{yy}"); w("sga", f"=h_opx_{yy}"); w("oth", f"=h_los_{yy}")
                w("op", f"=h_op_{yy}"); w("fin", f"=h_fin_{yy}"); w("pbt", f"=h_pbt_{yy}"); w("tax", f"=h_tax_{yy}"); w("ni", f"=h_ni_{yy}"); w("sh", f"=h_sh_{yy}"); w("eps", f"=h_eps_{yy}")
                w("am", f"=h_am_{yy}"); w("da", f"=h_dep_{yy}+h_am_{yy}"); w("cx", f"=h_cx_{yy}")
                w("chk2", f"=ROUND({cell('core')}+{cell('hum')}+h_isd_{yy}-{cell('rev')},1)")
                w("gr", '="未取得（FY2023）"' if y == 2024 else f"={cell('rev')}/{prev('rev')}-1"); w("nig", '="未取得（FY2023）"' if y == 2024 else f"={cell('ni')}/{prev('ni')}-1")
                if y == 2025: w("debt", "=h_debt_25")
                else: w("debt", '="未取得"')
                w("dwc", '="未取得"'); w("fcff", '="未取得"'); w("div", f"=h_div_{yy}*{cell('sh')}"); w("fcfad", '="未取得"')
            elif y == 2026:
                for k, _ in SEG:
                    w(f"c_{k}", f"=m9_{k}+(h_{k}_25-m9_{k}/(1+m9g_{k}))*(1+q4g_{k})")
                w("hum", 0); w("rp", 0)
                w("gp", f"=m9_gp+(({cell('rev')}-4625.5)*q4_gm)"); w("sga", "=m9_opx+q4_opx"); w("oth", "=imp26")
                w("fin", "=fin_26"); w("tax", f"={cell('pbt')}*tx26"); w("sh", "=sh_26"); w("am", "=am_26"); w("da", "=dep_26+am_26"); w("cx", "=cx_26"); w("dwc", "=dwc_26")
                w("debt", "=debt_v-debt_q4"); w("div", f"=dps*{cell('sh')}")
            else:
                gk = "g27" if y == 2027 else "g28" if y == 2028 else "gm" if y <= 2030 else "glt"
                for k, _ in SEG: w(f"c_{k}", f"={prev('c_' + k)}*(1+{gk}_{k}_{s})")
                w("mkt", f"=mk{str(y)[2:]}_{s}"); w("rp", f"=mk26_{s}")
                w("hum", f"=MAX(0,{cell('mkt')}-{cell('rp')})*shr_{s}")
                w("gm", f"={prev('gm')}+dgm_{s}" if y <= 2030 else f"={prev('gm')}"); w("gp", f"={cell('rev')}*{cell('gm')}")
                w("sga", f"=$E{R['sga']}*(1+opx_el*({cell('rev')}/$E{R['rev']}-1))"); w("oth", 0)
                w("fin", f"=-nir*{prev('debt')}"); w("tax", f"={cell('pbt')}*tax_w"); w("sh", "=sh_v"); w("am", "=am_f"); w("da", f"=fc_dep*{cell('rev')}+am_f"); w("cx", f"=fc_cx*{cell('rev')}")
                w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})"); w("div", f"=dps*{cell('sh')}")
                w("debt", f"=MAX(0,{prev('debt')}-({cell('fcff')}+{cell('fin')}*(1-tax_w)-{cell('div')}))")
            if y >= 2026:
                w("rev", f"={cell('core')}+{cell('hum')}")
                if y == 2026: w("cogs", f"={cell('rev')}-{cell('gp')}")
                else: w("cogs", f"={cell('rev')}-{cell('gp')}")
                w("op", f"={cell('gp')}-{cell('opex')}"); w("pbt", f"={cell('op')}+{cell('fin')}")
                w("ni", f"={cell('pbt')}-{cell('tax')}"); w("eps", f"={cell('ni')}/{cell('sh')}")
                w("gr", f"={cell('rev')}/{prev('rev')}-1"); w("nig", f"={cell('ni')}/{prev('ni')}-1")
                w("chk2", f"=ROUND({cell('core')}+{cell('hum')}-{cell('rev')},6)")
                w("fcff", f"={cell('op')}*(1-tax_w)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
                w("fcfad", f"={cell('fcff')}+{cell('fin')}*(1-tax_w)-{cell('div')}")
            w("core", f"=SUM({c}{R['c_amc']}:{c}{R['c_pes']})")
            w("chk1", f"=ROUND({cell('core')}-SUM({c}{R['c_amc']}:{c}{R['c_pes']}),6)")
            if y != 2024 or True:
                w("gm", f"={cell('gp')}/{cell('rev')}") if y <= 2026 else None
            w("opex", f"={cell('sga')}+{cell('oth')}"); w("opm", f"={cell('op')}/{cell('rev')}"); w("nim", f"={cell('ni')}/{cell('rev')}")
            w("epsc", f"={cell('eps')}+{cell('am')}*(1-tax_w)/{cell('sh')}")
        P.cell(R["hum"], 2, "份額 ×（市場 − 2026 市場）；FY2026 以前未揭露（視為 0，已含在 AMC 實績）")
        P.cell(R["gp"], 2, "FY2026＝9M 實績＋第四季營收 × 假設毛利率；FY2027 起＝營收 × 毛利率")
        P.cell(R["sga"], 2, "FY2027 起＝FY2026 ×(1＋彈性 × 營收相對 FY2026 增幅)")
        P.cell(R["fin"], 2, "FY2027 起＝− 淨利息率 × 前一年底借款")
        P.cell(R["debt"], 2, "年底借款＝前一年 −（FCFF ＋ 稅後利息 − 股利），下限 0；還清後累積現金不計利息收入（保守）")
        P.cell(R["chk2"], 2, "FY2024 加已處分 Industrial Systems 157.8")
        for k in ("gr", "gm", "opm", "nim", "nig"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(LAB) + 2
    P.column_dimensions["A"].width = 52; P.column_dimensions["B"].width = 30
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "epsc", "fcff", "fcfad", "hum"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$K${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（US$M）"; D["B4"] = "=px*sh_v"
    D["A5"] = "負債 D（借款＋租賃代理）"; D["B5"] = "=debt_v+lease_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金（2026-09-30）＝現金 − 借款 − 租賃代理 − 非支配"; D["B10"] = "=cash_v-debt_v-lease_v-mi_v"; dn("dcf_nc", "DCF!$B$10")
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
                 ("＋淨現金（2026-09-30；負值為淨負債）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!E{R['fcff']}*stub_frac"),
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
    r_cu = vrow("Comparable（未調整）：同業前瞻 P/E 中位數 × 情境 FY2028 調整後 EPS", [f"=pe_med*PL!G{PR[s]['epsc']}" for s in SC],
                "分母：自定義調整 EPS＝GAAP EPS ＋ 攤銷稅後（比公司調整後 EPS 窄，不加回重組等）；同業為日本企業（Analogy）")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（美元）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    r_cons = one("市場共識目標價（賣方 13 位平均；2026-10-23 頁面，第 3 季財報前）", "=cons_tp", "只作對照（Interested-party）")
    r_px = one("現價（In 頁 px）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（US$M；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!E{PR[s]['fcfad']}:PL!I{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=IF(B{r_up}>0.5,"已在最高級",TEXT(px*(1+IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5)))),"#,##0.00"))&"／"&TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,IF(B{r_up}>=-0.5,-0.5,-9))))),"#,##0.00")', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 調整後 EPS", f"=B{r_tp}/PL!G{PR['B']['epsc']}", "分母：模型情境 B FY2028 自定義調整 EPS")
    r_pe1 = one("現價 ÷ 情境 B FY2027 調整後 EPS", f"=px/PL!F{PR['B']['epsc']}", "分母：模型情境 B FY2027")
    r_pe2 = one("現行前瞻 P/E（現價 ÷ 公司調整後 EPS 指引中點）÷ 同業中位數", "=pe_self/pe_med", "只作對照")
    r_hum = one("無框力矩馬達增量營收占 FY2030 合併營收（情境 B）", f"=PL!I{PR['B']['hum']}/PL!I{PR['B']['rev']}")
    for rr_ in (r_up, r_a1, r_a2, r_a3, r_hum): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0.00"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "#,##0.00"
    V.column_dimensions["A"].width = 70; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, ca=r_ca, pd2=r_pd2, pe=r_pe, pe1=r_pe1, pe2=r_pe2, gj=r_gj, cons=r_cons, wk=r_wk, hum=r_hum)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司指引、市場共識對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或共識", "差", "說明"])
    rc = [("FY2026 GAAP EPS vs 公司指引中點（3.76–4.06）", f"=PL!E{PR['B']['eps']}", "=(g26_eps_lo+g26_eps_hi)/2", "公司指引（Interested-party）；差在 ±0.15 內視為校準通過"),
          ("FY2026 營收 9M 實績＋Q4 假設（公司未給營收指引）", f"=PL!E{PR['B']['rev']}", "=m9_amc+m9_ips+m9_pes", "後者只有 9M 實績，差＝第四季假設營收"),
          ("FY2026 調整後 EPS（自定義）vs 公司調整後 EPS 指引中點", f"=PL!E{PR['B']['epsc']}", "=g26_adj", "口徑不同：公司另排除重組、減損、關稅退款等；自定義只加回攤銷；差為負屬預期"),
          ("加權目標價 vs 賣方平均目標價", "=tp", "=cons_tp", "賣方 13 位平均（2026-10-23，第 3 季財報前，只作對照）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_rrx.py 批次格點＋線性內插；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 60; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_rrx.py 以 LibreOffice 批次重算每一案例）；結果為常數快照"
    base_tp = sens["base"]["tp"] if sens else None
    SE["A2"] = "快照基準值（加權，美元）"; SE["B2"] = base_tp
    SE["C2"] = "本檔即時值（加權，美元）"; SE["D2"] = "=tp"
    SE["E2"] = f"快照輸入指紋 {sens['fp'] if sens else ''}"
    SE["A3"] = "build_id"; SE["D3"] = bid
    SE["A4"] = "模型狀態"; SE["D4"] = '=IF(ABS(D2-B2)>0.005,"STALE","CURRENT")'
    hdr(SE, 6, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    sr = 7
    if sens:
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]); SE.cell(sr, 5, cs["tp"] / base_tp - 1)
            SE.cell(sr, 6, cs["dcf"]); SE.cell(sr, 7, cs["dcf"] / sens["base"]["dcf"] - 1 if sens["base"]["dcf"] else None); SE.cell(sr, 8, cs["rt"])
            SE[f"E{sr}"].number_format = "0.0%"; SE[f"G{sr}"].number_format = "0.0%"; SE[f"D{sr}"].number_format = "#,##0.00"; SE[f"F{sr}"].number_format = "#,##0.00"
            sr += 1
    SE.cell(sr + 1, 1, "說明：無區間的鍵（市場 mk*、a25_* 等實績）不列案例；線 6 低／高市場已作情境 C／A。")
    SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 34

    # ---------------- 對帳 ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = "對帳 — 本模型使用的線層級變數（讀線模型成品；與線模型不一致的自設值標明理由，不靜默改口徑）"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型一致？"])
    rec = [("無框力矩馬達市場 2026（低／基準／高）", "6", "224.9／329.8／480.8", "百萬美元", "五線歸因；基準單價 180 美元/顆", L6 + " 計算!C78／C145／C212", "Derived", "In mk26_C／B／A → PL 無框增量基期", "一致（照取；下游取數頁未列此變數，取自計算頁）"),
           ("無框力矩馬達市場 2030（低／基準／高）", "6", "602.9／1,901.6／4,813.7", "百萬美元", "同上；人形占約 95%", L6 + " 計算!G78／G145／G212", "Derived", "In mk30_*", "一致"),
           ("無框力矩馬達市場 2027–2029、2031–2032", "6", "見 In 頁 mk27–mk32", "百萬美元", "同上", L6 + " 計算!D–F、H–I 欄同列", "Derived", "In mk27_*…mk32_*", "一致"),
           ("無框馬達 2026 單價（低／基準／高）", "6", "100／180／300（年降 4／8／12%）", "美元/顆", "線 6 輸入", L6 + " 輸入!D103:F103、G103:I103", "Interested-party", "未進公式（已含在市場值）", "一致"),
           ("Kollmorgen（RRX）份額", "6", "未取得；本模型自設 1%／4%／8%", "%", "無框馬達", L6 + " 份額!F13", "Assumed", "In shr_A／B／C", "一致（線模型亦未取得；自設值為 Assumed）"),
           ("人形每台無框馬達 28 顆、協作每台 6 顆", "6／4", "未直接使用", "顆/台", "線 6 輸入", L6 + " 輸入!F51、E51", "Derived", "已含在市場值", "一致"),
           ("人形全球出貨（線 4）", "4", "未直接使用（經線 6 市場值）", "台", "—", L4 + " 下游取數", "Derived", "—", "未使用；本輪未重讀線 4 取數頁")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 30; RE.column_dimensions["F"].width = 60; RE.column_dimensions["I"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("資料中心訂單與電力設備需求退潮（PES、AMC 部分）", "是（g27_pes_C、g27_amc_C）", "情境 C", "公司訂單與積壓訂單、AI 資本支出"),
             ("新執行長投資組合檢討：AMR 業務出售或更大規模分拆", "否", "價格與範圍未取得；商譽減損 119.3 已入帳", "2027 年初 Q4 財報與公告"),
             ("高槓桿（借款 4,452、淨負債約 4.0 倍 EBITDA 以下 3 倍）與利率", "部分（nir 區間、債務隨現金流增減）", "再融資成本未逐筆建模", "利息費用、評等"),
             ("關稅與 IEEPA 退款（FY2026 指引含 +0.58 美元退款）", "否", "一次性；FY2027 起不含", "法院與海關公告"),
             ("人形無框馬達量產延後或 Kollmorgen 份額低於假設", "是（shr_C、Sens）", "無框增量占營收比例小（見 Val 頁）", "人形廠量產公告；線 6 份額更新"),
             ("同業比較面窄：僅 3 家日本同業以 11-20 報價換算；美國同業無同日報價", "部分（w_dcf 區間 0.3–0.7）", "Comparable 可信度低", "取得美國同業 11-20 收盤與共識 EPS 後重算"),
             ("報價日未取得、未納入中位數：SMC、THK、NSK、IKO、直得、上銀；美國：Rockwell、Emerson、Parker 等", "否", "收盤檔無該價或換算日不同", "—")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 50

    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([Q326, Q425, FOX, YHO, MST, YCL, L6 + "（計算、輸入、份額）", L4 + "（下游取數；本輪未重讀）", HWX + "（Comps 頁同業倍數）", HDX + "（Comp 頁 Nabtesco 換算 P/E）",
                            "公司檔案 physical_ai/map/profiles/RRX.md（背景；數字已自行查證或標明）"], 1):
        SO.cell(i, 1, s_)

    RD["A1"] = f"Regal Rexnord（RRX，NYSE）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-24（第 3 季財報 2026-11-03）；成品日 2026-11-24；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"#,##0.00")&" 美元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：Regal Rexnord 是以工業傳動、電機與資料中心電力設備為主的高槓桿工業公司；股價由資料中心帶動的電力與運動控制營收、毛利率與償債速度決定，無框力矩馬達（Kollmorgen，人形機器人用）只是小型選擇權"
    RD["A4"] = (f'="在 FY2027 AMC／IPS／PES 情境 B 營收年增 "&TEXT(g27_amc_B,"0%")&"／"&TEXT(g27_ips_B,"0%")&"／"&TEXT(g27_pes_B,"0%")&"、同業前瞻 P/E 中位數 "&TEXT(pe_med,"0.0")&" 倍的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"#,##0.00")&" 美元（DCF "&TEXT(Val!B{r_dcft},"#,##0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"#,##0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0")&" 百萬美元）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 美元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：PES 或 AMC 季營收年增轉負；毛利率跌破 35%；淨負債 ÷ 調整後 EBITDA 回升到 3.5 倍以上；FY2027 公司指引調整後 EPS 低於 FY2026 指引中點 10.75 美元"
    RD["A7"] = "範圍：只建三個官方分部（AMC、IPS、PES）與無框馬達增量；FY2023 未取得；AMR 業務出售、租賃負債、非支配權益、攤銷明細未取得。HTML 展示層見 dist/ 同名 .html"
    RD["A8"] = f'="命題回覆：問：資料中心帶動的電力與運動控制成長能否支撐高於現價的目標價？答：情境 B FY2028 調整後 EPS "&TEXT(PL!G{PR["B"]["epsc"]},"0.00")&" 美元，現價 "&TEXT(px,"0.00")&" 美元是其 "&TEXT(px/PL!G{PR["B"]["epsc"]},"0.0")&" 倍，同業中位數 "&TEXT(pe_med,"0.0")&" 倍；加權目標價 "&TEXT(Val!B{r_tp},"#,##0.00")&" 美元。問：人形機器人有多重要？答：情境 B FY2030 無框力矩馬達增量只占合併營收 "&TEXT(Val!B{r_hum},"0.00%")&"，結論不靠人形（模型輸出，非獨立證據）。"'
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
def lo_profile():
    d = os.path.join(os.environ.get("HOME", tempfile.gettempdir()), "lo_regal_rexnord")
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
    d = tempfile.mkdtemp(prefix="rrx_"); paths = []; VR = DR = None
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

CASES = [("q4g_amc", "FY2026 第四季年增 AMC"), ("q4g_ips", "FY2026 第四季年增 IPS"), ("q4g_pes", "FY2026 第四季年增 PES"), ("q4_gm", "FY2026 第四季毛利率"), ("q4_opx", "FY2026 第四季營業費用"),
         ("tx26", "FY2026 有效稅率"), ("cx_26", "FY2026 資本支出"), ("dwc_26", "FY2026 營運資金流出"), ("debt_q4", "FY2026 第四季償債"),
         ("g27_amc_B", "AMC 年增 FY2027（情境 B）"), ("g27_ips_B", "IPS 年增 FY2027（情境 B）"), ("g27_pes_B", "PES 年增 FY2027（情境 B）"),
         ("g28_amc_B", "AMC 年增 FY2028（情境 B）"), ("g28_ips_B", "IPS 年增 FY2028（情境 B）"), ("g28_pes_B", "PES 年增 FY2028（情境 B）"),
         ("gm_amc_B", "AMC 年增 FY2029–30（情境 B）"), ("gm_ips_B", "IPS 年增 FY2029–30（情境 B）"), ("gm_pes_B", "PES 年增 FY2029–30（情境 B）"),
         ("glt_amc_B", "AMC 年增 FY2031–32（情境 B）"), ("glt_ips_B", "IPS 年增 FY2031–32（情境 B）"), ("glt_pes_B", "PES 年增 FY2031–32（情境 B）"),
         ("shr_B", "Kollmorgen 無框馬達份額（情境 B）"), ("dgm_B", "每年毛利率變動（情境 B）"), ("opx_el", "營業費用彈性"), ("tax_w", "有效稅率"), ("nir", "淨利息率"),
         ("fc_dep", "折舊 ÷ 營收"), ("am_f", "無形資產攤銷"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("lease_v", "租賃負債代理"),
         ("rf", "無風險利率"), ("beta", "β"), ("erp", "股權風險溢酬"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重"), ("pe_med", "同業前瞻 P/E 中位數")]
BETA_B = 0.67 * statistics.median([1.07, 1.04]) + 0.33
FXD = {"cx_26": 96.0, "pe_med": statistics.median([21.1499, 15.6739, 13.7701])}
PES_ = None
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
            if isinstance(v, str): v = FXD[k] if k in FXD else None
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
    for c in cases: print(c["key"], c["drv"], round(c["tp"], 2), c["rt"])
    up0 = base["up"]; tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    revs = []
    for key, b0, lo_, hi_, note in (("g27_pes_B", 0.05, -0.30, 0.60, "情境 B 的 PES FY2027 營收年增（其他不變，整本重算；情境 A／C 不動）"),
                                    ("dgm_B", 0.002, -0.03, 0.03, "情境 B 每年毛利率變動 FY2027–30（其他不變）"),
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
    tmp = tempfile.mkdtemp(prefix="rrxf_"); raw = os.path.join(tmp, NAME)
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
    json.dump(dict(company="Regal Rexnord", ticker="RRX", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=v, unit="US$M", source=Q425 + "：分部淨銷售")
                                      for n, l, v in (("Automation & Motion Control", "營收：AMC 自動化與運動控制（Kollmorgen 等；不含無框馬達增量）", 1689.8), ("Industrial Powertrain Solutions", "營收：IPS 工業動力傳動", 2594.1), ("Power Efficiency Solutions", "營收：PES 電力效率方案（含資料中心）", 1650.6))],
                   sublines=[], check_labels=["檢查：分部合計 − 核心營收", "檢查：核心＋無框增量＋已處分分部 − 合併營收"],
                   note="PL 頁三分部即官方分部；FY2024 另有已處分的 Industrial Systems 157.8（核對列處理）；無框馬達增量為命題子列（FY2027 起，列於 AMC）。Q 頁有 9M'26 分部與合併營收核對。"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)


def write_readme(path, VR, PR, DR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f2 = lambda x: f"{x:,.2f}"; f0 = lambda x: f"{x:,.0f}"; pc = lambda x: f"{x*100:+.1f}%"
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
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f2(g(chr(66+i)+str(VR['dcfr'])))} | {f2(g(chr(66+i)+str(VR['cu'])))} | {f2(g(chr(66+i)+str(VR['tps'])))} | {g(chr(66+i)+str(VR['gate'])):,.0f} |" for i, s in enumerate(SC))
    rts = sorted(set(c['rt'] for c in sens['cases']))
    YS = range(2024, 2033)
    txt = f"""# physical_ai/regal_rexnord — Regal Rexnord（RRX，NYSE）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

線 6 共用零組件（無框力矩馬達；線模型成品 `physical_ai/lines/components/dist/20261116_共用零組件單位經濟_v0.1.xlsx`）。本 README 的數字由 `build_rrx.py` 從重算後的成品讀出產生，不手寫。

**範圍與未完成項目：** 只建三個官方分部（AMC、IPS、PES）與無框馬達增量命題子列；FY2023 未取得；AMR 業務出售（價格未取得，未建模）；租賃負債、非支配權益、攤銷明細未取得（以區間或 0 代替並標 Assumed）；同業只有 3 家日本公司使用 2026-11-20 報價（見下）；線 4 取數頁本輪未重讀。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A8'].value}

**讀法：** 現價 {f2(inv('px'))} 美元約為公司調整後 EPS 指引中點 10.75 美元的 {inv('pe_self'):.1f} 倍；Comparable 只用 3 家日本同業中位數 {inv('pe_med'):.1f} 倍，且套在自定義調整 EPS 上，可信度低。結論主要建立在 Assumed 的分部成長率與毛利率，以及面窄的同業倍數上；無框馬達增量（情境 B FY2030 占營收 {g(f"B{VR['hum']}")*100:.2f}%）不是承重項。

## 估值（評價日 2026-12-31；目標時點約 12 個月；美元）
| 情境 | DCF（推到 12 個月後） | Comparable（P/E × FY2028 調整後 EPS） | 情境目標價 | 收支關卡（FY2026–30，百萬美元） |
|---|---|---|---|---|
{scn}

- 加權目標價 **{f2(tp)} 美元**（DCF {f2(dcf)}、Comparable {f2(cmp_)}），現價 {f2(inv('px'))}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")} → 投資建議 **{g(f"B{VR['rec']}")}**。賣方平均目標價 {f2(inv('cons_tp'))}（13 位，2026-10-23，第 3 季財報前，只作對照）。
- 方法：DCF 標準版 FCFF（{g(f"B{VR['wk']}")}，為 WACC／ke／g／RONIC）50%＋Comparable 50%。相對倍數係數 1.00（未調整）。
- 同業：納入 Fanuc 21.15、安川 15.67、Nabtesco 13.77（皆換算至 2026-11-20 收盤，取自同專案成品）；**報價日未取得、不納入**：SMC、THK、NSK、IKO、直得、上銀（報價日早於 11-20）；美國同業收盤檔無報價。
- 隱含前瞻 P/E（加權目標價 ÷ 情境 B FY2028 調整後 EPS）{g(f"B{VR['pe']}"):.1f} 倍。

## 分部營收與標準損益（情境 B；US$M）
| 列 | {" | ".join(f"FY{y}" for y in YS)} |
|---|{"---|" * len(YS)}
| AMC（不含無框增量） | {yr('B','c_amc')} |
| IPS | {yr('B','c_ips')} |
| PES | {yr('B','c_pes')} |
| 無框力矩馬達增量 | {yr('B','hum')} |
| 營收（合併） | {yr('B','rev')} |
| 毛利 | {yr('B','gp')} |
| 營業利益 | {yr('B','op')} |
| 淨利 | {yr('B','ni')} |
| EPS（GAAP，美元） | {yr('B','eps','{:,.2f}')} |
| 調整後 EPS（自定義，美元） | {yr('B','epsc','{:,.2f}')} |
| 借款總額（年底） | {yr('B','debt','{:,.0f}')} |

- FY2024 營收含已處分的 Industrial Systems 157.8；FY2026＝9M 實績＋第四季假設（GAAP EPS 校準在公司指引 3.76–4.06 內，見 Recon）；FY2027 起為模型。
- FY2026 GAAP 含商譽減損 119.3 與關稅退款（公司揭露）。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 分部成長（A／B／C）：FY2027 AMC 14%／8%／0%、IPS 4%／2%／−4%、PES 10%／5%／−3%；之後逐年遞減（全為 Assumed，公司未給 FY2027 指引）。
- 無框馬達：線 6 市場（計算!C145:I145 等）× Kollmorgen 份額 8%／4%／1%（未取得，Assumed）× 2026 年以後增量。
- 毛利率每年 +0.4／+0.2／−0.5 個百分點至 FY2030；營業費用彈性 0.6；稅率 22%；淨利息率 6.5%；債務隨現金流增減。
- 折現：rf 5.05%（2026-11-19 盤中，非 11-20 收盤）、β {inv('beta'):.3f}、ERP 5%、g＝3%、RONIC＝WACC；評價日股數 66.8 百萬。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| 分部成長（g27、g28） | 季度分部營收年增、訂單與積壓訂單、FY2027 公司指引（2027-02） |
| 毛利率路徑（dgm） | 季度毛利率（9M'26 37.9%） |
| 同業 P/E（pe_med） | 美國工業同業同日倍數（目前未取得） |
| 無框馬達份額（shr） | 人形廠量產與供應鏈揭露 |

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
- 線 6：無框力矩馬達市場 2026–2032 低／基準／高（計算!C–I 欄第 78／145／212 列）→ 情境 C／B／A，照取；線 6 下游取數頁未列此變數，故取自計算頁。Kollmorgen 份額線模型未取得（份額!F13），自設 1%／4%／8%。
- 線 4：未重讀（無框馬達市場值已含人形需求）。

## 限制與未取得資料
- FY2023 財報、客戶集中度、10-Q 資產負債表（租賃、非支配權益）、攤銷明細、FY2026 營收指引（公司未給）、無框馬達營收與份額：未取得。
- rf 為 2026-11-19 盤中值；β 兩來源（1.07、1.04）為資料商。賣方目標價來源分歧（231 vs 262）。
- 自定義調整 EPS 只加回攤銷，與公司調整後 EPS（10.75）口徑不同，Comparable 因此偏保守。

## 檔案
- `build_rrx.py`：建置＋整本重算 Sens／反解（LibreOffice 批次）＋產生本 README、`segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-24）：初版；現價 {f2(inv('px'))}（2026-11-20 收盤）；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


if __name__ == "__main__":
    main()

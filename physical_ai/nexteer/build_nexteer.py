# 耐世特 Nexteer（1316.HK）個股模型 v0.1 — AI實體應用 Project（線 2 智慧車：線控轉向 SBW）
# 用法：python3 physical_ai/nexteer/build_nexteer.py            → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、segments.json、sens.json
#       python3 physical_ai/nexteer/build_nexteer.py --raw OUT  → 只寫一本（讀環境變數 NEX_OVERRIDE 的 JSON 覆寫 In 頁基準值），供重算
# 會計年度＝曆年（12 月決算，IFRS）。報表美元：金額 US$M（百萬美元）；股數百萬股；EPS 美元；股價、目標價港元（HK$，以 HKD=X 換算）。
# 結構：README、In、Hist、PL（三情境：命題模組 SBW → 官方地區分部 → 調整後 EBITDA → 標準損益表 → 現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 線層級變數（中國 SBW 市場）於建置時讀線模型 physical_ai/lines/smart_car/dist/*.xlsx「下游取數」頁，寫成 In 頁常數並記錄儲存格。
# 不讀 Siliconomics／Tokenomics／Wattonomics 協調檔（實體 AI 線）。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md。
import sys, os, json, subprocess, shutil, hashlib, tempfile, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261123"
NAME = f"{DATE}_耐世特個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032
LINE_REL = "physical_ai/lines/smart_car/dist/20261116_智慧車單位經濟_v0.1.xlsx"
LINE = os.path.join(ROOT, LINE_REL)
CSV = "physical_ai/data/closes2_20261109_20261120.csv"

# ---------------- 出處 ----------------
AR25 = "公司 2025 年報（港交所 2026-04-24）https://www1.hkexnews.hk/listedco/listconews/sehk/2026/0424/2026042402166.pdf"
AN24 = "公司 2024 年度業績公告（港交所 2025-03-19；含 2023 比較期）https://www1.hkexnews.hk/listedco/listconews/sehk/2025/0319/2025031900876.pdf"
IR26 = "公司 2026 中期業績公告（港交所 2026-08-12）https://www1.hkexnews.hk/listedco/listconews/sehk/2026/0812/2026081200874.pdf"
PR26 = "公司 2026 中期業績簡報（地區分部 EBITDA、SbW 專案）https://www.nexteer.com/wp-content/uploads/2026-Interim-Results-Announcement.pdf"
PR25 = "公司新聞稿 Nexteer Reports 2025 Year End Results（2026-03-24）https://www.newswire.com/news/nexteer-reports-2025-year-end-results"
SA = "stockanalysis.com（資料商）"
YCL = f"Yahoo Finance 日線 2026-11-20 收盤（{CSV}）"
US10 = "Trading Economics 美國 10 年期公債（2026-11-13 約 5.15%；與 tesla_robotaxi 同一取值，跨模型一致）https://tradingeconomics.com/united-states/government-bond-yield/news/598729"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"
LM = "線模型（本專案；立場：無外部誘因，但輸入多為產業媒體 Interested-party）"


# ---------------- 線模型「下游取數」讀值 ----------------
def read_line():
    wb = openpyxl.load_workbook(LINE, data_only=True); ws = wb["下游取數"]
    out = {}
    for r in range(3, ws.max_row + 1):
        a, y = ws.cell(r, 1).value, ws.cell(r, 2).value
        if a is None: continue
        out[(a, y)] = dict(row=r, lo=ws.cell(r, 3).value, b=ws.cell(r, 4).value, hi=ws.cell(r, 5).value, unit=ws.cell(r, 6).value)
    return out
LN = read_line()
def ln(name, y): return LN[(name, y)]
def lref(name, y, col): return f"{LINE_REL} 下游取數!{col}{ln(name, y)['row']}"
MKT = "中國 SBW 市場"

# ---------------- In 列 ----------------
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. FY2026 錨點：1H26 實績（Hist 頁）＋ 2H26 依 2H25 × (1＋各分部 1H26 年增＋g_anc)；公司不給營收與獲利指引")
inp("g_anc", "2H26 相對 2H25 的額外年增（各分部同加，於 1H26 分部年增之上）", "%", -0.03, 0.0, 0.03, "Assumed",
    IR26 + "：1H26 營收 +3.9%（北美 +2.7%、亞太 +3.2%、EMEASA +11.6%）", "公司無量化指引（中期公告「未來前景」為定性）；低／高 ±3pt")

sec("B. 命題模組：線控轉向（SBW）營收＝中國 SBW 市場（線模型）× 耐世特有效份額 ＋ 海外 SBW；扣除替代自家 EPS 的抵減")
for s, col in (("A", "E"), ("B", "D"), ("C", "C")):
    v = {"A": "hi", "B": "b", "C": "lo"}[s]
    for y, k in ((2026, "lmk26"), (2030, "lmk30"), (2032, "lmk32")):
        lo_ = ln(MKT, y)["lo"] if (s == "B" and y != 2026) else None
        hi_ = ln(MKT, y)["hi"] if (s == "B" and y != 2026) else None
        inp(f"{k}_{s}", f"情境 {s}：中國 SBW 市場 {y}（線模型{'高' if s == 'A' else '基準' if s == 'B' else '低'}值）", "億 RMB", lo_, ln(MKT, y)[v], hi_, "Derived",
            f"{lref(MKT, y, col)}（線 2 智慧車；滲透率 × 每套價值 × 國內零售）",
            ("敏感度低／高＝線模型低／高值（" + lref(MKT, y, 'C') + "、" + lref(MKT, y, 'E') + "）") if lo_ is not None else "情境 A＝線模型高值、B＝基準、C＝低值；中間年以幾何內插（PL 頁）", LM)
inp("fx_cny", "匯率 USD/CNY（預測期固定）", "RMB/US$", None, 6.6665, None, "Verified", YCL + "：CNY=X 2026-11-20 收盤 6.6665", "")
inp("sbw26", "耐世特 2026 中國 SBW 有效份額（÷ 線模型 2026 市場）", "%", 0.01, 0.02, 0.04, "Assumed",
    PR26 + "：2 個 SbW 專案量產（1H26 新增 1 個，首個 ASIL D 全線控）；累計 7 個專案、3 地區、8 家客戶（含 3 個開發中）", "SBW 營收公司未揭露；以 2% × 線模型 2026 市場約 1,098 百萬美元 ≈ 22 百萬美元。線模型 2026 市場 × 媒體所稱份額 35% ≈ 384 百萬美元，與公司 2 個量產專案的規模不符（見對帳頁）")
for s, v in (("A", 0.30), ("B", 0.15), ("C", 0.06)):
    inp(f"shr30_{s}", f"情境 {s}：耐世特中國 SBW 份額 FY2030（FY2027 起線性爬升，FY2031–32 持平）", "%", 0.08 if s == "B" else None, v, 0.25 if s == "B" else None, "Assumed",
        "產業媒體稱耐世特中國 SBW 份額約 35%（線模型 輸入!I17 引 36氪／車云；Interested-party，口徑為早期定點）", "競爭者：博世華域、採埃孚、捷太格特、浙江世宝、拓普、伯特利（陸廠國產替代）；份額數據未取得")
inp("cann", "SBW 替代自家 EPS 的營收抵減率（÷ SBW 營收）", "%", 0.15, 0.30, 0.45, "Assumed",
    "SBW 整套價值（線模型 2030 每套 5,249 元）高於 EPS；車型若原用耐世特 EPS，改用 SBW 後原 EPS 營收消失", "EPS 單價未取得；以 EPS 約為 SBW 價值的三到四成、且約一半 SBW 車型原為耐世特 EPS 估")
inp("os26", "海外（北美＋EMEASA）SBW 營收 FY2026", "US$M", 5.0, 10.0, 20.0, "Assumed", PR26 + "：已量產 2 個專案（含北美 Cybertruck 類全線控）", "公司未揭露 SBW 營收")
for s, v in (("A", 300.0), ("B", 120.0), ("C", 30.0)):
    inp(f"os30_{s}", f"情境 {s}：海外 SBW 營收 FY2030（FY2027 起線性）", "US$M", 50.0 if s == "B" else None, v, 250.0 if s == "B" else None, "Assumed",
        PR25 + "：2025 年續獲 SbW 訂單（中國新能源車廠）、首個歐系後輪轉向訂單", "7 個已獲專案 × 單一專案年營收未揭露；純假設")
inp("os_na", "海外 SBW 中北美占比（其餘歸 EMEASA）", "%", 0.4, 0.6, 0.8, "Assumed", "北美為最大客戶（GM）所在；只改分部歸屬", "SBW 利益率與抵減率不分地區")
inp("os_g", "海外 SBW 營收年增 FY2031–FY2032", "%/年", 0.05, 0.10, 0.20, "Assumed", "海外 L3 與線控法規晚於中國（線 2 框架）", "")
inp("m_sbw", "SBW 調整後 EBITDA 率", "%", 0.08, 0.14, 0.20, "Assumed",
    "類比：亞太分部 EBITDA 率 16.6%（2025；Hist 頁）；國產 SBW 每套由 2 萬元以上降到約 8,000 元（線模型 輸入!I32）", "新品初期折舊與工程費高；降價壓縮利潤")

sec("C. 官方地區分部（傳統轉向與傳動軸；成長路徑）與分部調整後 EBITDA 率")
for fam, lab, vals, rng, note in (
        ("gna", "北美傳統業務年增 FY2027 起", (0.04, 0.025, 0.0), (0.0, 0.045), "北美輕型車產量持平附近；最大客戶 GM 占營收 34%（2025）"),
        ("gap", "亞太傳統業務年增 FY2027 起", (0.07, 0.04, 0.0), (0.0, 0.07), "2025 訂單 49 億美元中 34% 來自中國車廠；中國價格戰"),
        ("gem", "EMEASA 傳統業務年增 FY2027 起", (0.08, 0.05, 0.01), (0.01, 0.08), "1H26 +11.6%（波蘭 +17.6%）")):
    for s, v in zip(SC, vals):
        inp(f"{fam}_{s}", f"情境 {s}：{lab}", "%/年", rng[0] if s == "B" else None, v, rng[1] if s == "B" else None, "Assumed", note,
            "FY2023→FY2025 年複合：北美 0.7%、亞太 10.0%、EMEASA 4.9%（Hist 頁，Derived）")
inp("m_na", "北美分部調整後 EBITDA 率（傳統業務）", "%", 0.065, "=(h_ebna_1h26+h_ebna_25-h_ebna_1h25)/(h_na_1h26+h_na_25-h_na_1h25)", 0.09, "Derived",
    "近四季（2H25＋1H26）＝(86＋174.3 − 86) ÷ (1,168.7＋2,290.4 − 1,138.3)", "2024 8.1%、2025 7.6%；低／高 Assumed")
inp("m_ap", "亞太分部調整後 EBITDA 率（傳統業務）", "%", 0.14, "=(h_ebap_1h26+h_ebap_25-h_ebap_1h25)/(h_ap_1h26+h_ap_25-h_ap_1h25)", 0.18, "Derived",
    "近四季＝(120＋243.3 − 116) ÷ (708.3＋1,468.9 − 686.5)", "2024 17.2%、2025 16.6%")
inp("m_em", "EMEASA 分部調整後 EBITDA 率（傳統業務）", "%", 0.085, "=(h_ebem_1h26+h_ebem_25-h_ebem_1h25)/(h_em_1h26+h_em_25-h_em_1h25)", 0.125, "Derived",
    "近四季＝(55＋68.9 − 35) ÷ (447.3＋798.9 − 400.9)", "2024 5.0%、2025 8.6%、1H26 12.2%")
for s, v in (("A", 0.005), ("B", 0.0), ("C", -0.01)):
    inp(f"madj_{s}", f"情境 {s}：FY2028 起傳統業務 EBITDA 率調整（年降價、客戶議價）", "pt", -0.01 if s == "B" else None, v, 0.005 if s == "B" else None, "Assumed",
        "車用零組件年降價約 1–3%；北美 EBITDA 率 2024 8.1% → 1H26 7.3%", "")
inp("eb_oth", "其他與未分配調整後 EBITDA FY2027 起", "US$M", -20.0, -12.0, -5.0, "Assumed", AR25 + "：2024 −20.1、2025 −14.7；1H26 約 +2（簡報合計 263 − 分部 261）", "")
inp("oth_k", "其他（Others）分部營收 FY2026 ＝ 1H26 × 本係數，其後持平", "x", None, 2.0, None, "Derived", IR26 + "：1H26 4.6（1H25 16.6）", "占營收約 0.2%")

sec("D. 損益表下半段與現金流")
inp("da_r", "折舊攤銷 ÷ 營收", "%", 0.058, "=(h_da_25+h_da_1h26)/(h_rev_25+h_rev_1h26)", 0.066, "Derived", "FY2025 285.0 ÷ 4,584.2 ＝ 6.2%；1H26 145.6 ÷ 2,328.8 ＝ 6.3%（年報、中期公告附註）", "含資本化開發費攤銷（4–8 年）")
inp("adj_x", "減損與其他調整（營業利益 −（調整後 EBITDA − 折舊攤銷））每年", "US$M", -25.0, -12.0, 0.0, "Assumed",
    "FY2024 −23.0、FY2025 −24.3（開發費減損 2024 46.9、2025 23.4）；1H26 約 −1.5（Hist 頁，Derived）", "")
inp("opex_r", "營業費用淨額 ÷ 營收（工程研發＋銷售管理 − 其他收益；只用於拆毛利，不影響營業利益）", "%", None, "=(h_gp_25-h_op_25+h_gp_1h26-h_op_1h26)/(h_rev_25+h_rev_1h26)", None, "Derived",
    "FY2025 (523.0 − 162.5) ÷ 4,584.2 ＝ 7.9%；1H26 (278.3 − 115.9) ÷ 2,328.8 ＝ 7.0%", "毛利＝營業利益＋營業費用淨額（Derived）；營業成本＝營收 − 毛利")
inp("nonop", "業外（財務收入 − 財務成本 ＋ 合資損益）FY2027 起", "US$M", 0.0, 4.0, 8.0, "Assumed", "FY2025 0.5＋4.9 ＝ 5.3；1H26 0.6 − 2.4＋1.0 ＝ −0.8", "淨現金 5 億美元，利息收入隨利率變動")
inp("tax_w", "有效稅率", "%", 0.22, 0.28, 0.33, "Assumed", "FY2023 29.7%、FY2024 35.9%、FY2025 33.0%、1H26 21.6%（Hist 頁，Derived）", "多國稅率（美、墨、中、波）與未認列虧損")
inp("nci_r", "非控股權益損益 ÷ 稅後淨利", "%", 0.05, 0.08, 0.12, "Derived", "FY2025 10.5 ÷ 112.5 ＝ 9.3%；FY2024 17.8%；1H26 5.0%", "中國合資子公司")
inp("pay_r", "配息率（股利 ÷ 歸屬母公司淨利；次年 7 月支付）", "%", 0.35, 0.45, 0.55, "Derived", AR25 + "：2025 末期股息每股 0.0183 美元（約 45.9 百萬美元）÷ 淨利 102.0 ＝ 45%", "2024 每股 0.0087 美元 ÷ 61.7 ＝ 35%")
inp("fc_cx", "資本支出（購置固定資產＋資本化開發）÷ 營收", "%", 0.056, "=(h_cx_24+h_cx_25+h_cx_1h26)/(h_rev_24+h_rev_25+h_rev_1h26)", 0.068, "Derived", "FY2024＋FY2025＋1H26 合計 691.9 ÷ 11,189.1 ＝ 6.2%（FY2024 6.5%、FY2025 5.6%、1H26 6.8%）", "與折舊攤銷率相近（維持性為主）；公司稱 2026 年下半年無重大投資計畫（中期公告）")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.02, 0.06, 0.10, "Derived", "（應收＋存貨 − 應付）÷ 營收：FY2024 5.4%、FY2025 5.6%（Hist 頁）", "")

sec("E. 估值輸入（評價日 2026-12-31；資產負債表 2026-06-30）")
inp("px", "股價（2026-11-20 收盤）", "HK$", None, 4.18, None, "Verified", YCL + "：1316.HK 2026-11-20 收盤 4.18", "未還原股利；11-13 收盤 4.515（同檔）")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("px_prev", "對照：2026-11-13 收盤（資料商前瞻 P/E 頁面日；只作對照，不進模型）", "HK$", None, 4.515, None, "Verified", YCL + "：1316.HK 2026-11-13 收盤 4.515", "pe_self 換算用")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("fx_hkd", "匯率 USD/HKD", "HK$/US$", None, 7.8426, None, "Verified", YCL + "：HKD=X 2026-11-20 收盤 7.8426", "每股美元價值 × 本匯率 ＝ 港元")
inp("sh_v", "評價日股數（2025-12-31 已發行股數；無庫藏股）", "百萬股", None, 2509.824293, None, "Verified", AR25 + "：已發行 2,509,824,293 股；1H26 加權平均 2,509,824 千股（中期公告）",
    "購股權 1,085,500 股（計畫 2024-06-05 屆滿；履約價未取得，占 0.04%，不計）")
inp("cash_v", "現金及約當現金＋受限存款（2026-06-30）", "US$M", None, "=596.506+0.103", None, "Derived", IR26 + "：596,506＋103 千美元（Verified 原文相加）", "")
inp("debt_v", "借款（2026-06-30；流動 147＋非流動 51,213 千美元）", "US$M", None, "=0.147+51.213", None, "Derived", IR26, "")
inp("lease_v", "租賃負債（2026-06-30；流動 8,509＋非流動 21,095 千美元）", "US$M", None, "=8.509+21.095", None, "Derived", IR26, "")
inp("mi_v", "非控股權益（帳面，2026-06-30）", "US$M", None, 53.491, None, "Verified", IR26, "以帳面扣除（DCF 標準規格第 3 節）")
inp("div_jul", "2026-07-09 支付的 2025 末期股息（2026-06-30 現金之後流出）", "US$M", None, 45.93, None, "Verified", IR26 + "：董事會 2026-06-17 宣派約 45,930 千美元，7 月 9 日支付", "")
inp("stub_frac", "Stub 期（2026-06-30 → 評價日 2026-12-31）", "年", None, 0.5, None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（美國 10 年期公債；現金流為美元）", "%", None, 0.0515, None, "Verified", US10, "2026-11-13 值，較股價基準日早 5 個交易日：逾期（跨模型一致，下次全面改版更新）")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "國家風險溢酬（營收加權：中國 27%、墨西哥 24%、波蘭 12%）", "%", 0.004, 0.008, 0.012, "Assumed", "未查 Damodaran 當期值；以營收加權約 0.8% 假設", "")
inp("beta", "β（兩來源 Blume 調整後中位數）", "x", "=E{r}-0.3", "=MEDIAN(0.67*1.85+0.33,0.67*0.67+0.33)", "=E{r}+0.3", "Derived",
    SA + "：耐世特 5 年 β 1.85（2026-11-13）；同業 5 家 5 年 β 中位數 0.67（現代摩比斯 1.36、電裝 0.61、NSK 0.53、捷太格特 0.67、拓普 0.80）", "Blume 0.67β＋0.33；兩來源差距大（2.01 vs 0.78）；敏感度 ±0.3（Assumed）")
inp("kd", "稅前負債成本", "%", None, "=rf+0.01", None, "Assumed", "rf ＋ 1pt（負債極少）", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", 0.75, 1.0, None, "Assumed", "1.00（未調整：缺耐世特對同業的歷史相對 P/E）",
    "低案 0.75 ≈ 現行前瞻 P/E 對同業中位數比（控股股東、港股流動性、客戶集中折價持續）；第 10 節第 18 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("F. 同業前瞻 P/E（資料商頁面值 × 2026-11-20 收盤 ÷ 頁面報價；分母 EPS 不變；Interested-party：賣方共識）")
PEERS = [("pe_mobis", "現代摩比斯 012330.KS（轉向、線控底盤）", 7.08, 352500, 323000, "2026-10-16", "KRW"),
         ("pe_denso", "電裝 6902.T（電子、電動轉向馬達）", 11.58, 1872.0, 1900.0, "2026-10-20", "JPY"),
         ("pe_nsk", "NSK 6471.T（軸承；轉向事業已分出）", 15.35, 1082.0, 892.5, "2026-09-18", "JPY"),
         ("pe_tuopu", "拓普集團 601689.SS（底盤、線控轉向）", 26.89, 46.52, 48.11, "2026-09-30", "CNY")]
for k, lab, pe, p0, p1, d, cur in PEERS:
    inp(k, f"前瞻 P/E：{lab}（換算至 2026-11-20）", "x", None, f"={pe}*{p1}/{p0}", None, "Derived",
        f"{SA}：{d} 頁面 forward P/E {pe}、報價 {p0:,g} {cur}；2026-11-20 收盤 {p1:,g}（{CSV}）", "會計口徑：韓國 K-IFRS、日本 IFRS、中國企業會計準則；P/E 為股權口徑，不受租賃處理影響", DV)
inp("pe_jtekt", "前瞻 P/E：捷太格特 6473.T（轉向；報價日未取得 2026-11-20）", "x", None, 11.28, None, "Interested-party",
    SA + "：2026-11-13 頁面值 11.28（報價 2,088.5 日圓）", "報價日未更新：6473.T 不在收盤 csv，未換算到 2026-11-20", DV)
inp("pe_med", "同業前瞻 P/E 中位數（5 家）", "x", None, "=MEDIAN(pe_mobis,pe_denso,pe_nsk,pe_tuopu,pe_jtekt)", None, "Derived",
    "少於 8 家，不做迴歸（第 10 節第 7 條）；伯特利 603596.SS 資料商頁面日期不明，不列", "")
inp("pe_self", "耐世特現行前瞻 P/E（換算至 2026-11-20；只作對照，不進模型）", "x", None, "=9.22*px/px_prev", None, "Derived",
    SA + "：2026-11-13 forward P/E 9.22 × 2026-11-20 收盤 ÷ 11-13 收盤", "", DV)
inp("cons_tp", "賣方平均目標價（只作對照，不進模型）", "HK$", None, 7.19, None, "Interested-party", SA + " forecast 頁（2026-11-18）：15 位分析師平均 7.19（5.00–11.40）", "", DV)
inp("cons_rev26", "共識 FY2026 營收（只作對照，不進模型）", "US$M", None, 4800.0, None, "Interested-party", SA + " forecast 頁（2026-11-18）：4.8B（4.7–5.2B）", "取整到億美元", DV)
inp("cons_eps26", "共識 FY2026 EPS（美元；只作對照，不進模型）", "US$", None, 0.06, None, "Interested-party", SA + " forecast 頁（2026-11-18）：0.06（0.06–0.07）", "取整到美分", DV)

FAM = [("lmk26", "中國 SBW 市場 2026（億 RMB；線模型）"), ("lmk30", "中國 SBW 市場 2030"), ("lmk32", "中國 SBW 市場 2032"), ("shr30", "耐世特中國 SBW 份額 FY2030"),
       ("os30", "海外 SBW 營收 FY2030（US$M）"), ("gna", "北美傳統業務年增"), ("gap", "亞太傳統業務年增"), ("gem", "EMEASA 傳統業務年增"), ("madj", "FY2028 起傳統業務 EBITDA 率調整")]

# ---------------- Hist 資料 ----------------
# (鍵, 項目, FY2023, FY2024, FY2025, 1H25, 1H26, 標記, 來源)；None＝未取得
HIST = [
    ("h_rev", "營收", 4206.793, 4276.086, 4584.217, 2242.248, 2328.838, "Verified", AN24 + "；" + AR25 + "；" + IR26),
    ("h_cogs", "營業成本", 3838.200, 3827.569, "=4584.217-523.005", 1983.341, 2050.535, "Verified", "同上；FY2025＝營收 − 毛利"),
    ("h_gp", "毛利", 368.593, 448.517, 523.005, 258.907, 278.303, "Verified", "同上"),
    ("h_epd", "工程及產品開發費", 150.694, 179.804, 174.8, 75.393, 72.102, "Verified", "同上；FY2025 為年報管理層討論百萬美元取整"),
    ("h_sda", "銷售及分銷＋行政費用", "=19.721+135.094", "=21.261+146.824", 192.9, "=10.704+81.474", "=12.635+87.390", "Verified", "同上；FY2025 為取整"),
    ("h_og", "其他收益淨額", -1.684, 14.137, "=162.502-523.005+174.8+192.9", 1.792, 9.768, "Verified", "同上；FY2025＝營業利益 − 毛利 ＋ 費用（Derived，約 7.2）"),
    ("h_op", "營業利益", 61.400, 114.765, 162.502, 93.128, 115.944, "Verified", "同上"),
    ("h_fin", "財務收支淨額", "=5.021-5.064", "=6.128-7.936", "=8.516-8.052", "=4.182-3.390", "=0.605-2.378", "Verified", "財務（投資）收入 − 財務成本"),
    ("h_jv", "應佔合資損益", 2.870, 4.275, "=167.822-162.502-8.516+8.052", 1.761, 0.996, "Verified", "FY2025＝稅前 − 營業利益 − 財務淨額（約 4.9）"),
    ("h_pbt", "稅前利益", 64.227, 117.232, 167.822, 95.681, 115.167, "Verified", "同上"),
    ("h_tax", "所得稅", 19.052, 42.121, 55.303, 26.667, 24.845, "Verified", "同上（正數＝費用）"),
    ("h_pat", "稅後淨利", 45.175, 75.111, 112.519, 69.014, 90.322, "Verified", "同上"),
    ("h_nci", "非控股權益損益", 8.438, 13.392, 10.515, 5.534, 4.516, "Verified", "同上"),
    ("h_ni", "歸屬母公司淨利", 36.737, 61.719, 102.004, 63.480, 85.806, "Verified", "同上"),
    ("h_na", "北美營收", 2259.055, 2193.003, 2290.418, 1138.304, 1168.703, "Verified", "分部資訊（地區）"),
    ("h_ap", "亞太營收", 1214.732, 1337.848, 1468.928, 686.533, 708.276, "Verified", "同上"),
    ("h_em", "EMEASA 營收", 725.921, 717.161, 798.900, 400.851, 447.299, "Verified", "同上"),
    ("h_ot", "其他營收", 7.085, 28.074, 25.971, 16.560, 4.560, "Verified", "同上"),
    ("h_ebna", "北美調整後 EBITDA", None, 178.493, 174.341, 86.0, 86.0, "Verified", AR25 + "；1H 為 " + PR26 + "（百萬美元取整）"),
    ("h_ebap", "亞太調整後 EBITDA", None, 230.114, 243.305, 116.0, 120.0, "Verified", "同上"),
    ("h_ebem", "EMEASA 調整後 EBITDA", None, 35.882, 68.919, 35.0, 55.0, "Verified", "同上"),
    ("h_ebot", "其他與未分配調整後 EBITDA", None, -20.071, -14.702, "=230-86-116-35", "=263-86-120-55", "Verified", "1H＝簡報 EBITDA 合計 − 三分部（Derived）"),
    ("h_ebitda", "調整後 EBITDA", 346.556, 424.418, 471.863, 230.0, 263.0, "Verified", "同上"),
    ("h_da", "折舊攤銷（固定資產＋使用權＋無形）", None, "=146.124+17.013+123.588", "=146.431+14.741+123.846", 138.289, 145.624, "Verified", AR25 + "；" + IR26),
    ("h_cx", "資本支出（購置固定資產＋新增無形資產）", None, "=165.990+110.836", "=133.858+123.707", "=54.994+55.491", "=88.113+69.433", "Verified", "同上"),
    ("h_rec", "應收帳款（期末）", None, 820.395, 924.089, None, 920.372, "Verified", "同上"),
    ("h_inv", "存貨（期末）", None, 283.826, 279.929, None, 322.635, "Verified", "同上"),
    ("h_ap2", "應付帳款（期末）", None, 873.172, 945.605, None, 1010.928, "Verified", "同上"),
    ("h_ocf", "營業現金流", None, 446.234, 404.933, 142.0, 262.0, "Verified", "同上；1H 為簡報取整"),
    ("h_div", "支付母公司股東股利", None, 7.529, 21.835, None, None, "Verified", AR25),
    ("h_pl_eps", "產品線：電動轉向（EPS）", 2860.746, 2887.391, 3128.936, 1525.379, 1624.862, "Verified", "產品線營收（含 SBW；公司未單獨揭露 SBW）"),
    ("h_pl_cis", "產品線：轉向柱與中間軸（CIS）", 379.419, 440.201, 463.605, 228.499, 220.890, "Verified", "同上"),
    ("h_pl_hps", "產品線：液壓轉向（HPS）", 168.082, 177.178, 181.923, 91.624, 84.022, "Verified", "同上"),
    ("h_pl_dl", "產品線：傳動軸（DL）", 798.546, 771.316, 809.753, 396.746, 399.064, "Verified", "同上"),
]
HSUF = ["23", "24", "25", "1h25", "1h26"]


def build(path, ov=None, sens=None, rev=None):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 US$M、股數百萬股、EPS 美元、股價港元"
    IN["A2"] = "標記：Verified（附出處）／Interested-party（寫明是誰、誘因欄）／Analogy／Assumed（給區間）／Derived（寫明換算）；查不到寫「未取得」"
    hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "說明", "來源誘因"])
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
    last_in = r
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
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
    H["A1"] = "Hist — 歷史財務（IFRS；US$M）。公式只引用本頁儲存格（定義名稱 h_*_23／_24／_25／_1h25／_1h26）；空白＝未取得"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "1H25", "1H26", "標記", "來源"])
    HR = {}
    for i, (k, lab, *vals, tag, src) in enumerate(HIST, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate(vals):
            if v is None:
                H.cell(i, 2 + j, "未取得"); continue
            H.cell(i, 2 + j, v).font = BLUE
            dn(f"{k}_{HSUF[j]}", f"Hist!${'BCDEF'[j]}${i}")
        H.cell(i, 7, tag); H.cell(i, 8, src); HR[k] = i
    i = 3 + len(HIST)
    H.cell(i, 1, "檢查：產品線合計 − 營收（應為 0）")
    for j, L in enumerate("BCDEF"):
        H.cell(i, 2 + j, f"=ROUND({L}{HR['h_pl_eps']}+{L}{HR['h_pl_cis']}+{L}{HR['h_pl_hps']}+{L}{HR['h_pl_dl']}-{L}{HR['h_rev']},3)")
    H.cell(i + 1, 1, "檢查：地區分部合計 − 營收（應為 0）")
    for j, L in enumerate("BCDEF"):
        H.cell(i + 1, 2 + j, f"=ROUND({L}{HR['h_na']}+{L}{HR['h_ap']}+{L}{HR['h_em']}+{L}{HR['h_ot']}-{L}{HR['h_rev']},3)")
    H.cell(i + 2, 1, "營運資金 ÷ 營收（Derived）")
    for L in "CD": H.cell(i + 2, "ABCDEF".index(L) + 1, f"=({L}{HR['h_rec']}+{L}{HR['h_inv']}-{L}{HR['h_ap2']})/{L}{HR['h_rev']}")
    H.cell(i + 3, 1, "減損與其他調整＝營業利益 −（調整後 EBITDA − 折舊攤銷）（Derived）")
    for L in "CDF": H.cell(i + 3, "ABCDEF".index(L) + 1, f"={L}{HR['h_op']}-({L}{HR['h_ebitda']}-{L}{HR['h_da']})")
    H.cell(i + 4, 1, "SBW 營收：公司未單獨揭露（含在 EPS 產品線）；2023 分部 EBITDA 與折舊攤銷、資本支出未取得（本版只取 2024 年度公告的比較期損益）")
    H.column_dimensions["A"].width = 44; H.column_dimensions["H"].width = 60

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（線控轉向 SBW）→ 官方地區分部 → 調整後 EBITDA → 標準損益表 → 現金流（US$M；FY2023–FY2025 實績、FY2026＝1H26 實績＋2H 外推、FY2027 起模型）"
    PR = {}
    rr = 3
    keys = ["cnm", "cns", "sbwap", "canap", "convap", "ap", "os", "sbwna", "canna", "convna", "na", "sbwem", "canem", "convem", "em", "ot", "rev",
            "chk1", "chk2", "gr", "sbw", "sbwsh",
            "ebna", "ebap", "ebem", "ebot", "ebitda", "ebm", "da", "adj",
            "cogs", "gp", "gpm", "opex", "op", "opm", "nonop", "pbt", "tax", "pat", "nci", "ni", "nim", "nig", "sh", "eps", "epshk",
            "cx", "dwc", "fcff", "div", "fcfad"]
    labs = ["中國 SBW 市場（線模型；億 RMB）", "耐世特中國 SBW 有效份額", "亞太：SBW 營收（中國；子列）", "亞太：SBW 替代自家 EPS 抵減（子列）", "亞太：傳統轉向與傳動軸（子列）",
            "亞太（Asia Pacific）營收", "海外 SBW 營收合計（北美＋EMEASA；說明列）", "北美：SBW 營收（子列）", "北美：SBW 替代自家 EPS 抵減（子列）", "北美：傳統轉向與傳動軸（子列）",
            "北美（North America）營收", "EMEASA：SBW 營收（子列）", "EMEASA：SBW 替代自家 EPS 抵減（子列）", "EMEASA：傳統轉向與傳動軸（子列）", "EMEASA（歐洲、中東、非洲及南美）營收",
            "其他（Others）營收", "營收（合併）",
            "檢查：子列合計 − 分部（三地區合計，應為 0）", "檢查：分部合計 − 合併營收（應為 0）", "營收年增率", "SBW 營收合計（三地區）", "SBW 占營收",
            "北美調整後 EBITDA", "亞太調整後 EBITDA", "EMEASA 調整後 EBITDA", "其他與未分配調整後 EBITDA", "調整後 EBITDA", "調整後 EBITDA 率", "折舊攤銷", "減損與其他調整",
            "營業成本", "毛利", "毛利率", "營業費用（淨額：工程研發＋銷管 − 其他收益）", "營業利益", "營業利益率", "業外（財務收支淨額＋合資損益）", "稅前利益", "所得稅", "稅後淨利",
            "非控股權益損益", "歸屬母公司淨利", "淨利率（歸屬母公司）", "淨利年增率", "平均股數（百萬股）", "EPS（美元，基本＝稀釋）", "EPS（港元，換算）",
            "資本支出", "營運資金增加", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "股利（母公司，現金支付年）", "自由現金流扣股利（收支關卡用）"]
    assert len(keys) == len(labs)
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, k in enumerate(keys)}; R["hdr"] = rr + 1
        for i, l in enumerate(labs): P.cell(b + i, 1, l)
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1)
            cell = lambda k, c=c: f"{c}{R[k]}"; prev = lambda k, p=p: f"{p}{R[k]}"
            def w(k, f, c=c): P[f"{c}{R[k]}"] = f
            sfx = {2023: "23", 2024: "24", 2025: "25"}.get(y)
            if sfx:
                for k, h in (("na", "h_na"), ("ap", "h_ap"), ("em", "h_em"), ("ot", "h_ot"), ("rev", "h_rev"), ("gp", "h_gp"), ("op", "h_op"), ("pbt", "h_pbt"),
                             ("tax", "h_tax"), ("pat", "h_pat"), ("nci", "h_nci"), ("ni", "h_ni"), ("ebitda", "h_ebitda")):
                    w(k, f"={h}_{sfx}")
                for k in ("sbwap", "canap", "sbwna", "canna", "sbwem", "canem"): w(k, 0)
                w("convap", f"={cell('ap')}"); w("convna", f"={cell('na')}"); w("convem", f"={cell('em')}")
                w("cogs", f"=h_cogs_{sfx}"); w("opex", f"={cell('gp')}-{cell('op')}"); w("nonop", f"={cell('pbt')}-{cell('op')}")
                w("sh", "=sh_v")
                if y >= 2024:
                    for k, h in (("ebna", "h_ebna"), ("ebap", "h_ebap"), ("ebem", "h_ebem"), ("ebot", "h_ebot"), ("da", "h_da"), ("cx", "h_cx")): w(k, f"={h}_{sfx}")
                    w("adj", f"={cell('op')}-({cell('ebitda')}-{cell('da')})")
                    w("div", f"=h_div_{sfx}")
                if y == 2025:
                    w("dwc", f"=(h_rec_25+h_inv_25-h_ap2_25)-(h_rec_24+h_inv_24-h_ap2_24)")
                    w("fcff", f"={cell('op')}*(1-tax_w)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            elif y == 2026:
                fy = lambda h: f"({h}_1h26+({h}_25-{h}_1h25)*({h}_1h26/{h}_1h25+g_anc))"
                w("cnm", "=lmk26_" + s); w("cns", "=sbw26")
                w("sbwap", f"={cell('cnm')}*{cell('cns')}*100/fx_cny"); w("canap", f"=-cann*{cell('sbwap')}")
                w("ap", "=" + fy("h_ap")); w("convap", f"={cell('ap')}-{cell('sbwap')}-{cell('canap')}")
                w("os", "=os26")
                w("sbwna", f"={cell('os')}*os_na"); w("canna", f"=-cann*{cell('sbwna')}"); w("na", "=" + fy("h_na")); w("convna", f"={cell('na')}-{cell('sbwna')}-{cell('canna')}")
                w("sbwem", f"={cell('os')}*(1-os_na)"); w("canem", f"=-cann*{cell('sbwem')}"); w("em", "=" + fy("h_em")); w("convem", f"={cell('em')}-{cell('sbwem')}-{cell('canem')}")
                w("ot", "=h_ot_1h26*oth_k")
                w("ebna", f"=h_ebna_1h26+(h_ebna_25-h_ebna_1h25)*(h_na_1h26/h_na_1h25+g_anc)")
                w("ebap", f"=h_ebap_1h26+(h_ebap_25-h_ebap_1h25)*(h_ap_1h26/h_ap_1h25+g_anc)")
                w("ebem", f"=h_ebem_1h26+(h_ebem_25-h_ebem_1h25)*(h_em_1h26/h_em_1h25+g_anc)")
                w("ebot", "=h_ebot_1h26+(h_ebot_25-h_ebot_1h25)")
                w("nonop", "=h_fin_1h26+h_jv_1h26+nonop/2"); w("cx", f"=fc_cx*{cell('rev')}"); w("div", "=div_jul")
            else:
                n = y - 2026
                if y <= 2029: w("cnm", f"=lmk26_{s}*(lmk30_{s}/lmk26_{s})^({n}/4)")
                elif y == 2030: w("cnm", f"=lmk30_{s}")
                elif y == 2031: w("cnm", f"=SQRT(lmk30_{s}*lmk32_{s})")
                else: w("cnm", f"=lmk32_{s}")
                w("cns", f"=sbw26+(shr30_{s}-sbw26)*{min(n, 4)}/4")
                w("sbwap", f"={cell('cnm')}*{cell('cns')}*100/fx_cny"); w("canap", f"=-cann*{cell('sbwap')}")
                w("os", f"=os26+(os30_{s}-os26)*{n}/4" if y <= 2030 else f"={prev('os')}*(1+os_g)")
                w("sbwna", f"={cell('os')}*os_na"); w("canna", f"=-cann*{cell('sbwna')}")
                w("sbwem", f"={cell('os')}*(1-os_na)"); w("canem", f"=-cann*{cell('sbwem')}")
                for seg, g in (("na", "gna"), ("ap", "gap"), ("em", "gem")):
                    w("conv" + seg, f"={prev('conv' + seg)}*(1+{g}_{s})")
                    w(seg, f"={cell('conv' + seg)}+{cell('sbw' + seg)}+{cell('can' + seg)}")
                w("ot", f"={prev('ot')}")
                adj = f"+madj_{s}" if y >= 2028 else ""
                for seg, m in (("na", "m_na"), ("ap", "m_ap"), ("em", "m_em")):
                    w("eb" + seg, f"=({m}{adj})*({cell('conv' + seg)}+{cell('can' + seg)})+m_sbw*{cell('sbw' + seg)}")
                w("ebot", "=eb_oth"); w("nonop", "=nonop"); w("cx", f"=fc_cx*{cell('rev')}"); w("div", f"=pay_r*{prev('ni')}")
            if y >= 2026:
                w("rev", f"={cell('na')}+{cell('ap')}+{cell('em')}+{cell('ot')}")
                w("ebitda", f"={cell('ebna')}+{cell('ebap')}+{cell('ebem')}+{cell('ebot')}")
                w("da", f"=da_r*{cell('rev')}"); w("adj", "=adj_x")
                w("op", f"={cell('ebitda')}-{cell('da')}+{cell('adj')}")
                w("opex", f"=opex_r*{cell('rev')}"); w("gp", f"={cell('op')}+{cell('opex')}"); w("cogs", f"={cell('rev')}-{cell('gp')}")
                w("pbt", f"={cell('op')}+{cell('nonop')}"); w("tax", f"={cell('pbt')}*tax_w"); w("pat", f"={cell('pbt')}-{cell('tax')}")
                w("nci", f"={cell('pat')}*nci_r"); w("ni", f"={cell('pat')}-{cell('nci')}"); w("sh", "=sh_v")
                w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
                w("fcff", f"={cell('op')}*(1-tax_w)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
                w("fcfad", f"={cell('fcff')}-{cell('div')}")
            # 共通
            w("chk1", f"=ROUND(({cell('convna')}+{cell('sbwna')}+{cell('canna')}-{cell('na')})+({cell('convap')}+{cell('sbwap')}+{cell('canap')}-{cell('ap')})+({cell('convem')}+{cell('sbwem')}+{cell('canem')}-{cell('em')}),6)")
            w("chk2", f"=ROUND({cell('na')}+{cell('ap')}+{cell('em')}+{cell('ot')}-{cell('rev')},6)")
            if y > 2023:
                w("gr", f"={cell('rev')}/{prev('rev')}-1"); w("nig", f"={cell('ni')}/{prev('ni')}-1")
            else:
                w("gr", ""); w("nig", "")
            w("sbw", f"={cell('sbwap')}+{cell('sbwna')}+{cell('sbwem')}"); w("sbwsh", f"={cell('sbw')}/{cell('rev')}")
            w("gpm", f"={cell('gp')}/{cell('rev')}"); w("opm", f"={cell('op')}/{cell('rev')}"); w("nim", f"={cell('ni')}/{cell('rev')}")
            w("ebm", f"={cell('ebitda')}/{cell('rev')}")
            w("eps", f"={cell('ni')}/{cell('sh')}"); w("epshk", f"={cell('eps')}*fx_hkd")
        for y in (2023, 2024):
            for k in ("dwc", "fcff", "fcfad"): P[f"{COL[y]}{R[k]}"] = ""
        P.cell(R["cnm"], 2, "FY2026、2030、2032＝線模型下游取數；中間年幾何內插")
        P.cell(R["sbwap"], 2, "＝市場 × 份額 × 100 ÷ USD/CNY（百萬美元）")
        P.cell(R["convap"], 2, "FY2026＝分部 − SBW − 抵減；FY2027 起 ×(1＋年增)")
        P.cell(R["ap"], 2, "FY2026＝1H26 ＋ 2H25 ×(1H26 年增＋g_anc)")
        P.cell(R["ebap"], 2, "FY2027 起＝傳統 EBITDA 率 ×(傳統＋抵減)＋SBW 率 × SBW")
        P.cell(R["gp"], 2, "預測期＝營業利益＋營業費用淨額（Derived）")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 ×(1 − tax_w)；股票報酬不加回")
        P.cell(R["div"], 2, "FY2026＝7 月已付 45.9；其後＝配息率 × 前一年淨利")
        for k in ("cns", "gr", "sbwsh", "ebm", "gpm", "opm", "nim", "nig"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        for k in ("eps", "epshk"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0000"
        PR[s] = R; rr = b + len(labs) + 2
    P.column_dimensions["A"].width = 46; P.column_dimensions["B"].width = 34
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "sbw", "fcfad", "ebitda"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現（美元），評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)；每股價值 × USD/HKD"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（US$M）"; D["B4"] = "=px/fx_hkd*sh_v"
    D["A5"] = "負債 D（借款＋租賃）"; D["B5"] = "=debt_v+lease_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金（2026-06-30）＝現金 − 借款 − 租賃 − 7 月已付股息 − 非控股權益"; D["B10"] = "=cash_v-debt_v-lease_v-div_jul-mi_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND(((px/fx_hkd*sh_v)*(rf+beta*(erp+crp))+(debt_v+lease_v)*kd*(1-tax_w))/((px/fx_hkd*sh_v)+debt_v+lease_v)+wacc_adj-wacc,9)"
    D["A12"] = "檢查：淨現金 − 資產負債表各項（應為 0）"; D["B12"] = "=ROUND(dcf_nc-(cash_v-debt_v-lease_v-div_jul-mi_v),9)"
    for rr_ in range(3, 11): D[f"B{rr_}"].number_format = "0.00%" if rr_ in (3, 6, 7, 8, 9) else "#,##0.0"
    DR = {}; rr = 14
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
                 ("＋淨現金（扣 7 月股息與非控股權益）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!F{R['fcff']}*stub_frac"),
                 ("股權價值（US$M）", f"=B{t+4}+B{t+5}+B{t+6}"),
                 ("每股價值（評價日，美元）", f"=MAX(0,B{t+7})/sh_v"),
                 ("每股價值（評價日，港元）", f"=B{t+8}*fx_hkd"),
                 ("DCF 情境價值推到 12 個月後（港元；×(1＋ke)^t − 期間股利）", f"=B{t+9}*(1+ke)^t_tgt-pay_r*PL!F{R['eps']}*fx_hkd"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+11}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=B{t+2}/G{b+1}"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（港元，推到 12 個月後）", f"=((B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6)+B{t+5}+B{t+6})/sh_v*fx_hkd*(1+ke)^t_tgt-pay_r*PL!F{R['eps']}*fx_hkd")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 10, ps=t + 9, ev=t + 4, tvs=t + 11, alt=t + 14, chk=b + 6)
        dn(f"dcf_{s}", f"DCF!$B${t+10}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 64

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價（港元）：DCF（權重 w_dcf）＋ Comparable（同業前瞻 P/E 中位數 × FY2028 EPS，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}")
        V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後（港元）", [f"=dcf_{s}" for s in SC], "DCF 頁")
    r_cu = vrow("Comparable（未調整）：同業前瞻 P/E 中位數 × FY2028 EPS × USD/HKD", [f"=pe_med*PL!H{PR[s]['eps']}*fx_hkd" for s in SC],
                "目標時點 2027-11 的 NTM 以 FY2028 代；P/E 已含淨現金利息，故不另加淨現金")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（港元）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權，港元）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權，港元）", f"=E{r_ca}")
    r_cmpu = one("Comparable 目標價（未調整，港元）", f"=E{r_cu}")
    r_px = one("現價（In 頁 px，港元）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（US$M；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級（港元）", f'=IF(B{r_up}>0.5,"—（已為最高級）",TEXT(px*(1+IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5)))),"0.00"))&"／"&TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,IF(B{r_up}>=-0.5,-0.5,-9))))),"0.00")', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 EPS（港元）", f"=B{r_tp}/PL!H{PR['B']['epshk']}")
    r_pe2 = one("現行前瞻 P/E（資料商 NTM，換算至 2026-11-20）÷ 同業中位數", "=pe_self/pe_med", "只作對照")
    r_evx = one("現行 EV ÷ FY2026 調整後 EBITDA（情境 B；顯示用）", f"=(px/fx_hkd*sh_v-dcf_nc)/PL!F{PR['B']['ebitda']}", "EV＝市值 − 淨現金（扣非控股權益與 7 月股息）")
    r_sbws = one("SBW 營收占比 FY2030（情境 B；顯示用）", f"=PL!J{PR['B']['sbwsh']}")
    for rr_ in (r_up, r_a1, r_a2, r_a3, r_sbws): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "0.00"
    for rr_ in (r_tp, r_dcft, r_cmpt, r_cmpu): V[f"B{rr_}"].number_format = "0.00"
    V.column_dimensions["A"].width = 70; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, cmpu=r_cmpu, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, pd2=r_pd2, pe=r_pe, gj=r_gj, wk=r_wk, evx=r_evx, sbws=r_sbws)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與市場共識、公司揭露對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或共識", "差", "說明"])
    rc = [("FY2026 營收（US$M）", f"=PL!F{PR['B']['rev']}", "=cons_rev26", "資料商共識 15 位（2026-11-18；Interested-party）"),
          ("FY2026 EPS（美元）", f"=PL!F{PR['B']['eps']}", "=cons_eps26", "共識取整到美分"),
          ("1H26 營收（實績） vs 模型 FY2026 營收 × 一半", f"=PL!F{PR['B']['rev']}/2", "=h_rev_1h26", "季節性：2H 營收通常略高於 1H"),
          ("FY2026 中國 SBW 營收（模型） vs 線模型 2026 市場 × 媒體份額 35%（US$M）", f"=PL!F{PR['B']['sbwap']}", "=lmk26_B*0.35*100/fx_cny",
           "差距約一個數量級：線模型 2026 滲透率 5% 與媒體份額 35% 不能同時對應耐世特 2 個量產專案的規模；本模型改用 sbw26（Assumed）"),
          ("加權目標價 vs 賣方平均目標價（港元）", "=tp", "=cons_tp", "賣方 15 位平均（Interested-party）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_nexteer.py 格點搜尋；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 64; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_nexteer.py 以 LibreOffice 重算每一案例）；結果為常數快照。估值評等只看上檔空間"
    SE["A2"] = "每案只改一個鍵（或一組同源鍵），其他輸入維持基準；情境權重與線模型讀值另列"
    SE["A3"] = "build_id"; SE["D3"] = sens["build_id"] if sens else "（基準檔）"
    SE["A4"] = "In 頁輸入指紋（敏感度重算時）"; SE["D4"] = sens["fp_sum"] if sens else 0
    SE["A5"] = "In 頁輸入指紋（本檔即時）"; SE["D5"] = f"=SUM(In!E4:E{last_in - 1})"
    SE["A6"] = "基準目標價（敏感度重算時，港元）"; SE["D6"] = sens["base"]["tp"] if sens else 0
    SE["A7"] = "模型狀態"; SE["D7"] = f'=IF(AND(ABS(D5-D4)<=1E-6*MAX(1,ABS(D4)),ABS(Val!B{r_tp}-D6)<1E-6),"CURRENT","STALE")'
    hdr(SE, 9, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    sr = 10
    if sens:
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]); SE.cell(sr, 5, f"=D{sr}/Val!$B${r_tp}-1")
            SE.cell(sr, 6, cs["dcf"]); SE.cell(sr, 7, f"=F{sr}/Val!$B${r_dcft}-1")
            SE.cell(sr, 8, f'=IF(D{sr}/px-1>0.5,"強力買進",IF(D{sr}/px-1>0.25,"買進",IF(D{sr}/px-1>=-0.25,"中立",IF(D{sr}/px-1>=-0.5,"放空","強力放空"))))')
            SE[f"E{sr}"].number_format = "+0.0%;-0.0%"; SE[f"G{sr}"].number_format = "+0.0%;-0.0%"; SE[f"D{sr}"].number_format = "0.00"; SE[f"F{sr}"].number_format = "0.00"
            sr += 1
        SE.cell(sr + 1, 1, "註：相對基準欄以本檔 Val 頁加權目標價（即時）為分母；模型狀態 STALE 表示 In 頁改過、須重跑 build_nexteer.py")
    SE.column_dimensions["A"].width = 50; SE.column_dimensions["B"].width = 36; SE.column_dimensions["C"].width = 20

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = f"對帳 — 本模型使用的線層級變數（線 2 智慧車：{LINE_REL}「下游取數」頁；建置時讀值）；與線模型不一致的自設值寫明理由"
    hdr(RE, 2, ["變數", "線模型儲存格", "線模型低", "線模型基準", "線模型高", "單位", "本模型用法", "一致性"])
    recs = []
    for nm, y, use, cons in ((MKT, 2026, "In lmk26_A／B／C（情境 A＝高、B＝基準、C＝低）→ PL 中國 SBW 市場列", "一致（直接讀值）"),
                             (MKT, 2030, "In lmk30_*；B 情境敏感度低／高＝線模型低／高", "一致"),
                             (MKT, 2032, "In lmk32_*", "一致"),
                             ("中國線控轉向（SBW）滲透率", 2026, "說明用（市場已含）", "不一致疑點：5% 對應約 90 萬套，與耐世特 2 個量產專案、同業定點數不相稱；本模型不改線值，改以份額 sbw26 吸收"),
                             ("中國線控轉向（SBW）滲透率", 2030, "說明用（市場已含）", "一致"),
                             ("中國 SBW 每套價值", 2026, "說明用；cann 抵減率依據", "一致"),
                             ("中國 SBW 每套價值", 2030, "說明用；cann 抵減率依據", "一致"),
                             ("中國 L3 滲透率", 2030, "未進公式（L3 帶動冗餘轉向，已含在 SBW 滲透率）", "—")):
        d = ln(nm, y)
        recs.append((f"{nm} {y}", f"下游取數!C{d['row']}:E{d['row']}", d["lo"], d["b"], d["hi"], d["unit"], use, cons))
    recs.append(("耐世特中國 SBW 份額約 35%", "輸入!I17（說明文字，引 36氪／車云）", None, 0.35, None, "%", "未進公式；本模型 FY2030 份額 15%（情境 B），FY2026 有效份額 2%",
                 "不一致（自設）：媒體 35% 為早期定點口徑；以 35% 乘線模型 2026 市場得約 3.8 億美元，超過耐世特全部 SBW 量產規模的合理範圍"))
    recs.append(("海外 SBW（北美、EMEASA）", "線模型未建（只建中國）", None, None, None, "US$M", "In os26、os30_*（Assumed）", "線模型無對應值"))
    for i, row in enumerate(recs, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 36; RE.column_dimensions["G"].width = 50; RE.column_dimensions["H"].width = 60

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("最大客戶 GM 占營收 34%（2025）：專案流失或北美減產", "部分（gna_C、madj_C）", "單一客戶超過一成；未另建尾端情境", "GM 北美產量、REPS／傳動軸續約公告"),
             ("美國關稅（墨西哥廠占北美營收約 48%）", "否", "關稅轉嫁條款未揭露", "USMCA 審查、客戶回收（commodity／tariff recovery）金額"),
             ("中國價格戰與國產替代（亞太 EBITDA 率約 17%）", "是（gap_C、madj_C、shr30_C）", "情境 C", "亞太分部 EBITDA 率、中國車廠訂單占比"),
             ("控股股東（中國航空工業集團體系）與地緣政治（美國對中資供應商限制）", "否", "不可量化", "美國聯網車規則、國防授權法相關名單"),
             ("SBW 安全事故或法規回退", "部分（shr30_C、os30_C）", "情境 C", "中國 GB 17675 執行、L3 試點事故"),
             ("開發費資本化減損（2024 46.9、2025 23.4）", "是（adj_x）", "Sens", "年報減損附註")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 56; RK.column_dimensions["C"].width = 40

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([AR25, AN24, IR26, PR26, PR25, SA + " statistics／forecast 頁（耐世特 2026-11-13、11-18；同業見 In 頁）", YCL, US10, LINE_REL + "（下游取數、輸入頁）"], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    bid = sens["build_id"] if sens else "（基準檔）"
    RD["A1"] = f"耐世特 Nexteer（1316.HK）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-18（財報與資料商）；成品日 2026-11-23；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.00")&" 港元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：耐世特是電動轉向（EPS）為主的轉向與傳動軸 Tier 1；實體 AI 的切入點是線控轉向（SBW）——SBW 營收＝中國 SBW 市場（線模型）× 份額 ＋ 海外 SBW，扣除替代自家 EPS 的抵減；北美、亞太、EMEASA 三個官方地區分部的傳統業務以成長路徑建模"
    RD["A4"] = (f'="在北美／亞太傳統業務 EBITDA 率維持近四季 "&TEXT(m_na,"0.0%")&"／"&TEXT(m_ap,"0.0%")&"、中國 SBW 市場依線模型基準且耐世特 FY2030 份額 "&TEXT(shr30_B,"0%")&" 的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 港元（DCF "&TEXT(Val!B{r_dcft},"0.00")&"、Comparable "&TEXT(Val!B{r_cmpt},"0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0")&" 百萬美元）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 港元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")&"。SBW 在 FY2030 占營收 "&TEXT(Val!B{r_sbws},"0.0%")&"（情境 B）"'
    RD["A6"] = "失效條件（可觀察）：北美分部 EBITDA 率連兩個半年低於 6.5%；GM 營收占比下降且無新專案遞補；亞太分部 EBITDA 率跌破 14%；SBW 量產專案到 2028 年仍不超過 4 個"
    RD["A7"] = "範圍：不讀 Siliconomics／Tokenomics／Wattonomics（算力線）；車用晶片不在本模型。HTML 展示層見 dist/ 同名 .html"
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
LO = "file://" + os.path.join(os.path.expanduser("~"), "lo_nexteer_profile")   # 呼叫端以 HOME 隔離 LibreOffice 設定檔
def recalc_many(srcs, outdir):
    os.makedirs(outdir, exist_ok=True)
    for attempt in range(2):
        try:
            subprocess.run(["soffice", f"-env:UserInstallation={LO}", "--headless", "--convert-to", "xlsx", "--outdir", outdir] + srcs,
                           check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180 + 20 * len(srcs))
            break
        except subprocess.TimeoutExpired:
            subprocess.run(["pkill", "-f", "lo_nexteer_profile"], check=False)
            if attempt: raise
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

def run_cases(ovs, with_dr=False):
    d = tempfile.mkdtemp(prefix="nex_"); paths = []; VR = DR = None
    for i, ov in enumerate(ovs):
        p = os.path.join(d, f"c{i:03d}.xlsx"); VR, _, DR = build(p, ov); paths.append(p)
    outs = []
    for j in range(0, len(paths), 20):
        outs += recalc_many(paths[j:j + 20], os.path.join(d, "o"))
    vals = [read_vals(o, VR, DR if with_dr else None) for o in outs]
    shutil.rmtree(d, ignore_errors=True); return vals

def in_val(k, col):
    for row in ROWS:
        if row[0] == k: return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)

def fingerprint():
    return hashlib.sha1(json.dumps([r for r in ROWS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]

CASES = [("g_anc", "2H26 額外年增"), ("sbw26", "2026 中國 SBW 有效份額"), ("shr30_B", "中國 SBW 份額 FY2030（情境 B）"), ("cann", "SBW 替代自家 EPS 抵減率"),
         ("os26", "海外 SBW 營收 FY2026"), ("os30_B", "海外 SBW 營收 FY2030（情境 B）"), ("os_na", "海外 SBW 北美占比"), ("os_g", "海外 SBW 年增 FY2031–32"), ("m_sbw", "SBW EBITDA 率"),
         ("gna_B", "北美傳統業務年增（情境 B）"), ("gap_B", "亞太傳統業務年增（情境 B）"), ("gem_B", "EMEASA 傳統業務年增（情境 B）"),
         ("m_na", "北美 EBITDA 率"), ("m_ap", "亞太 EBITDA 率"), ("m_em", "EMEASA EBITDA 率"), ("madj_B", "FY2028 起 EBITDA 率調整（情境 B）"), ("eb_oth", "其他與未分配 EBITDA"),
         ("da_r", "折舊攤銷 ÷ 營收"), ("adj_x", "減損與其他調整"), ("nonop", "業外"), ("tax_w", "有效稅率"), ("nci_r", "非控股權益比"), ("pay_r", "配息率"),
         ("beta", "β"), ("erp", "ERP"), ("crp", "國家風險溢酬"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"),
         ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("w_dcf", "DCF 權重"), ("rel_k", "相對倍數係數")]

def rating(u): return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"


def main():
    if "--raw" in sys.argv:
        ov = json.loads(os.environ.get("NEX_OVERRIDE", "{}"))
        build(sys.argv[sys.argv.index("--raw") + 1], ov); return
    px0 = in_val("px", "base")
    beta_b = (0.67 * 1.85 + 0.33 + 0.67 * 0.67 + 0.33) / 2
    specs = []   # (drv, set, key, ov)
    for k, drv in CASES:
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if k == "beta": v = beta_b + (-0.3 if side == "lo" else 0.3)
            if v is None: continue
            b_ = in_val(k, "base")
            if isinstance(b_, (int, float)) and abs(v - b_) < 1e-12: continue
            specs.append((drv + ("（低）" if side == "lo" else "（高）"), f"{k}＝{v:.4g}", k, {k: v}))
    lo30, hi30 = in_val("lmk30_B", "lo"), in_val("lmk30_B", "hi"); lo32, hi32 = in_val("lmk32_B", "lo"), in_val("lmk32_B", "hi")
    specs.append(("線模型中國 SBW 市場 2030、2032 取低值（情境 B）（低）", f"lmk30_B＝{lo30:.4g}、lmk32_B＝{lo32:.4g}", "lmk30_B,lmk32_B", {"lmk30_B": lo30, "lmk32_B": lo32}))
    specs.append(("線模型中國 SBW 市場 2030、2032 取高值（情境 B）（高）", f"lmk30_B＝{hi30:.4g}、lmk32_B＝{hi32:.4g}", "lmk30_B,lmk32_B", {"lmk30_B": hi30, "lmk32_B": hi32}))
    for s in SC:
        specs.append((f"只看情境 {s}", f"w_{s}＝1，其餘 0", f"w_{s}", {f"w_{x}": (1.0 if x == s else 0.0) for x in SC}))
    specs.append(("SBW 全無（中國與海外 SBW 皆為 0）", "sbw26＝0、shr30_B＝0、os26＝0、os30_B＝0", "sbw26,shr30_B,os26,os30_B", {"sbw26": 0.0, "shr30_B": 0.0, "os26": 0.0, "os30_B": 0.0}))
    vals = run_cases([{}] + [s[3] for s in specs], with_dr=True)
    base = vals[0]; print("基準", base)
    cases = []
    for (drv, st, key, ov), v in zip(specs, vals[1:]):
        v.pop("alt_dcf", None); v.pop("alt_tp", None)
        cases.append(dict(drv=drv, set=st, key=key, **v)); print(drv, round(v["tp"], 3), v["rt"])
    up_ = base["alt_tp"] / px0 - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rating(up_)))
    # 反解：格點搜尋（兩輪），找使上檔＝0、升一級、降一級的值
    up0 = base["up"]; tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    targets = [("目標價＝現價（上檔 0%）", 0.0)] + ([(f"升一級邊界（上檔 {upper:+.0%}）", upper)] if upper is not None else []) + ([(f"降一級邊界（上檔 {lower:+.0%}）", lower)] if lower is not None else [])
    revs = []
    for key, b0, lo_, hi_, note in (("shr30_B", 0.15, 0.0, 0.6, "情境 B 耐世特中國 SBW 份額 FY2030（情境 A／C 不動）"),
                                    ("m_na", None, -0.05, 0.25, "北美傳統業務調整後 EBITDA 率（三情境同動）"),
                                    ("m_ap", None, -0.05, 0.30, "亞太傳統業務調整後 EBITDA 率（三情境同動）")):
        grid = [lo_ + (hi_ - lo_) * i / 16 for i in range(17)]
        ups = [v["up"] for v in run_cases([{key: x} for x in grid])]
        fine = []
        for lab, tgt in targets:
            br = next((i for i in range(16) if (ups[i] - tgt) * (ups[i + 1] - tgt) <= 0), None)
            fine.append(br)
        reqs = []
        for br in fine:
            if br is not None: reqs += [grid[br] + (grid[br + 1] - grid[br]) * j / 8 for j in range(9)]
        fv = [v["up"] for v in run_cases([{key: x} for x in reqs])] if reqs else []
        pos = 0
        for (lab, tgt), br in zip(targets, fine):
            if br is None:
                revs.append(dict(label=lab, key=key, val=f"區間 {lo_:+.0%}～{hi_:+.0%} 內無解", base=b0 if b0 is not None else "LTM", note=note)); continue
            xs = reqs[pos:pos + 9]; us = fv[pos:pos + 9]; pos += 9
            j = next((j for j in range(8) if (us[j] - tgt) * (us[j + 1] - tgt) <= 0), 0)
            x = xs[j] + (xs[j + 1] - xs[j]) * ((tgt - us[j]) / (us[j + 1] - us[j]) if us[j + 1] != us[j] else 0)
            revs.append(dict(label=lab, key=key, val=x, base=b0 if b0 is not None else "LTM", note=note)); print("反解", key, lab, x)
    fp = fingerprint(); bid = f"nexteer-{VER}-{fp}"
    # In 頁數值指紋（SUMPRODUCT 數值格）：由基準重算檔讀
    tmp = tempfile.mkdtemp(prefix="nexf_"); raw = os.path.join(tmp, NAME)
    sens = dict(build_id=bid, fp=fp, fp_sum=0, base=base, cases=cases)
    VR, PR, DR = build(raw, {}, sens, revs)
    out = recalc_many([raw], os.path.join(tmp, "o"))[0]
    wbv = openpyxl.load_workbook(out, data_only=True); sens["fp_sum"] = wbv["Sens"]["D5"].value
    VR, PR, DR = build(raw, {}, sens, revs)
    out = recalc_many([raw], os.path.join(tmp, "o2"))[0]
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    shutil.copy(out, os.path.join(dist, NAME))
    wbv = openpyxl.load_workbook(os.path.join(dist, NAME), data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10]); assert not errs, errs
    print("模型狀態", wbv["Sens"]["D7"].value)
    json.dump(dict(build_id=bid, base=base, cases=cases, rev=revs), open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
    json.dump(dict(company="耐世特 Nexteer", ticker="1316.HK", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=v, unit="US$M", source=AR25 + " 分部資訊")
                                      for n, l, v in (("North America", "北美（North America）營收", 2290.418), ("Asia Pacific", "亞太（Asia Pacific）營收", 1468.928),
                                                      ("EMEASA", "EMEASA（歐洲、中東、非洲及南美）營收", 798.900), ("Others", "其他（Others）營收", 25.971))],
                   sublines=[dict(parent="Asia Pacific", labels=["亞太：SBW 營收（中國；子列）", "亞太：SBW 替代自家 EPS 抵減（子列）", "亞太：傳統轉向與傳動軸（子列）"]),
                             dict(parent="North America", labels=["北美：SBW 營收（子列）", "北美：SBW 替代自家 EPS 抵減（子列）", "北美：傳統轉向與傳動軸（子列）"]),
                             dict(parent="EMEASA", labels=["EMEASA：SBW 營收（子列）", "EMEASA：SBW 替代自家 EPS 抵減（子列）", "EMEASA：傳統轉向與傳動軸（子列）"])],
                   check_labels=["檢查：分部合計 − 合併營收（應為 0）", "檢查：子列合計 − 分部（三地區合計，應為 0）"],
                   note="官方應報導分部為地區（北美、亞太、EMEASA；其他為未分配）。產品線（EPS、CIS、HPS、DL）列在 Hist 頁並有合計核對；SBW 公司未單獨揭露，為命題子列"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)
    write_readme(os.path.join(dist, NAME), VR, PR, bid, sens, revs)


def write_readme(path, VR, PR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f2 = lambda x: f"{x:,.2f}"; pc = lambda x: f"{x*100:+.1f}%"
    inv = lambda k: I['E' + str([r for r in range(1, 400) if I['A' + str(r)].value == k][0])].value
    base_tp = g(f"B{VR['tp']}")
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / base_tp - 1))[:8]
    def rvv(r):
        return r["val"] if isinstance(r["val"], str) else f"{r['val']:.4f}"
    bv = lambda r: r["base"] if isinstance(r["base"], str) else f"{r['base']:.4f}"
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {rvv(r)} | {bv(r)} | {r['note']} |" for r in revs)
    yr = lambda s, k, fm="{:,.0f}": " | ".join(fm.format(P[f'{COL[y]}{PR[s][k]}'].value) if isinstance(P[f'{COL[y]}{PR[s][k]}'].value, (int, float)) else "—" for y in YRS)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f2(g(chr(66+i)+str(VR['dcfr'])))} | {f2(g(chr(66+i)+str(VR['cu'])))} | {f2(g(chr(66+i)+str(VR['tps'])))} |" for i, s in enumerate(SC))
    rts = sorted(set(rating(c["tp"] / inv("px") - 1) for c in sens["cases"]))
    nosbw = next(c for c in sens["cases"] if c["key"].startswith("sbw26,shr30_B"))
    txt = f"""# physical_ai/nexteer — 耐世特 Nexteer（1316.HK）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

港交所主板上市；財報美元（IFRS），股價與目標價港元（USD/HKD {inv('fx_hkd')}，2026-11-20）。本 README 的數字由 `build_nexteer.py` 從重算後的成品讀出產生，不手寫。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

**一句話：** 耐世特 FY2025 營收 {wb['Hist']['D3'].value/1000:,.2f} 十億美元，幾乎全是傳統電動轉向、轉向柱與傳動軸，最大客戶 GM 占 34%；SBW 在 1H26 只有 2 個量產專案、營收未揭露。股價便宜（現行前瞻 P/E 約 {inv('pe_self'):.1f} 倍、EV／EBITDA 約 {g(f"B{VR['evx']}"):.1f} 倍、淨現金約占市值 {wb['DCF']['B10'].value/wb['DCF']['B4'].value*100:.0f}%）。把 SBW 全部拿掉，加權目標價 {f2(nosbw['tp'])} 港元（{pc(nosbw['tp']/base_tp-1)}），上檔仍有 {pc(nosbw['tp']/inv('px')-1)}、評等 {rating(nosbw['tp']/inv('px')-1)}：上檔主要來自傳統業務現金流、淨現金與同業倍數，SBW 在基準情境是加分項，在情境 A（線模型高值、份額 30%）貢獻才明顯變大。

## 估值（評價日 2026-12-31；目標時點約 12 個月；港元）
| 情境 | DCF（推到 12 個月後） | Comparable（P/E） | 情境目標價 |
|---|---|---|---|
{scn}

- 加權目標價 **{f2(base_tp)} 港元**（DCF {f2(g(f"B{VR['dcf']}"))}、Comparable {f2(g(f"B{VR['cmp']}"))}），現價 {f2(inv('px'))}（2026-11-20 收盤），上檔 {pc(g(f"B{VR['up']}"))}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")}（情境 B FY2026–30 自由現金流扣股利 {g(f"C{VR['gate']}"):,.0f} 百萬美元；情境 A／C {g(f"B{VR['gate']}"):,.0f}／{g(f"D{VR['gate']}"):,.0f}）→ 投資建議 **{g(f"B{VR['rec']}")}**。
- 方法：DCF 標準版 FCFF（{g(f"B{VR['wk']}")}，為 WACC／ke／g／RONIC；美元現金流、美國 10 年期 rf）50%＋Comparable 50%（同業 5 家前瞻 P/E 中位數 {inv('pe_med'):.2f} 倍 × FY2028 EPS）。相對倍數係數 1.00（未調整：缺歷史相對倍數）；Sens 另列 0.75（現行折價持續）。
- 隱含前瞻 P/E（加權目標價 ÷ 情境 B FY2028 EPS）{g(f"B{VR['pe']}"):.1f} 倍；賣方 15 位平均目標價 {inv('cons_tp'):.2f} 港元（2026-11-18，Interested-party）。

## 營收分部（官方地區分部；PL 頁三情境，情境 B，US$M）
| 列 | {" | ".join(f"FY{y}" for y in YRS)} |
|---|{"---|" * len(YRS)}
| SBW 營收合計（命題子列加總） | {yr('B','sbw')} |
| 北美 | {yr('B','na')} |
| 亞太 | {yr('B','ap')} |
| EMEASA | {yr('B','em')} |
| 其他 | {yr('B','ot')} |
| 營收（合併） | {yr('B','rev')} |
| 調整後 EBITDA | {yr('B','ebitda')} |
| 營業利益 | {yr('B','op')} |
| 歸屬母公司淨利 | {yr('B','ni')} |
| EPS（美元） | {yr('B','eps','{:.4f}')} |

- FY2023–FY2025 實績（年報、年度業績公告）；FY2026＝1H26 實績＋2H25 ×(1＋1H26 分部年增＋g_anc)；FY2027 起為模型。
- 命題模組：中國 SBW＝線模型中國 SBW 市場（2026、2030、2032 讀值，中間年幾何內插）× 耐世特有效份額（2026 年 2% → FY2030 情境 B 15%）÷ USD/CNY；海外 SBW 由 FY2026 1,000 萬美元線性到 FY2030（情境 B 1.2 億美元）。SBW 營收的 30% 視為替代自家 EPS 的抵減（cann）。三子列＝各地區分部，核對列為 0。
- 分部調整後 EBITDA：傳統業務用近四季分部 EBITDA 率（北美 {inv('m_na')*100:.1f}%、亞太 {inv('m_ap')*100:.1f}%、EMEASA {inv('m_em')*100:.1f}%），SBW 用 {inv('m_sbw')*100:.0f}%；其他與未分配 −12 百萬美元；營業利益＝EBITDA − 折舊攤銷（營收 × {inv('da_r')*100:.1f}%）＋減損與其他調整（−12）。
- 標準損益表：營業成本、毛利、營業費用在預測期以「營業費用淨額 ÷ 營收」（FY2025＋1H26 實績 {inv('opex_r')*100:.2f}%）拆出（Derived），不影響營業利益。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 中國 SBW 份額 FY2030：情境 A／B／C 30%／15%／6%（媒體稱約 35%，為早期定點口徑，Interested-party）；海外 SBW FY2030 3 億／1.2 億／0.3 億美元；SBW EBITDA 率 14%（8–20%）。
- 傳統業務年增（A／B／C）：北美 4%／2.5%／0%、亞太 7%／4%／0%、EMEASA 8%／5%／1%；FY2028 起 EBITDA 率調整 +0.5pt／0／−1pt。
- 折現：rf 5.15%（美國 10 年期，逾期、跨模型一致）、β {inv('beta'):.3f}（耐世特 1.85 與同業中位數 0.67 的 Blume 中位數）、ERP 5%、國家風險 0.8%（Assumed）、g＝3%、RONIC＝WACC；資本支出 {inv('fc_cx')*100:.1f}%、折舊攤銷 {inv('da_r')*100:.1f}%、營運資金 6% × 營收增量；稅率 28%。
- 評價日股數 2,509.8 百萬股（無庫藏股；過期購股權 0.04% 不計）。淨現金（2026-06-30）＝現金 596.6 − 借款 51.4 − 租賃 29.6 − 7 月已付股息 45.9 − 非控股權益 53.5 ＝ {wb['DCF']['B10'].value:,.1f} 百萬美元。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| 北美傳統業務年增與 EBITDA 率（gna、m_na） | 北美分部半年 EBITDA 率（1H26 7.3%）、GM 北美產量 |
| 亞太 EBITDA 率（m_ap） | 亞太分部半年 EBITDA 率（1H26 16.9%）、中國車廠訂單占比 |
| 中國 SBW 份額（shr30） | 公司 SbW 量產專案數（1H26 2 個、累計獲 7 個）、中國 SBW 定點車型 |
| 相對倍數（rel_k） | 耐世特前瞻 P/E ÷ 同業中位數（現行 {V[f"B{VR['pe'] + 1}"].value:.2f}） |

失效條件：{RD['A6'].value.split('：', 1)[1]}。各 Sens 案例的估值評等出現「{'、'.join(rts)}」。

**區間穩健性：** 強力買進只在點值成立，**不穩健**——距下緣 {g(f"B{VR['dist']}"):.1f} 個百分點，Sens 共 {len(sens['cases'])} 案中有 {sum(1 for c in sens['cases'] if rating(c['tp']/inv('px')-1) != g(f"B{VR['rt']}"))} 案降到買進或以下（最低 {min(c['tp'] for c in sens['cases']):.2f} 港元，{min(sens['cases'], key=lambda c: c['tp'])['drv']}）；除「只看情境 C」為中立外，其餘案例都在買進以上。反解：北美 EBITDA 率降到 {next(r['val'] for r in revs if r['key']=='m_na' and '降一級' in r['label'])*100:.2f}%、或亞太降到 {next(r['val'] for r in revs if r['key']=='m_ap' and '降一級' in r['label'])*100:.2f}%（其他不變）即降為買進；中國 SBW 份額在 0–60% 內都不會使評等降級。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f2(c['tp'])} | {pc(c['tp']/base_tp-1)} | {rating(c['tp']/inv('px')-1)} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁。

## 反解（整本重算，格點搜尋；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
讀線 2 智慧車模型 `{LINE_REL}`「下游取數」：中國 SBW 市場 2026／2030／2032（D44、C45:E45、C46:E46 等；情境 A＝高、B＝基準、C＝低）→ In `lmk26_*`、`lmk30_*`、`lmk32_*`。SBW 滲透率、每套價值、L3 滲透率只作說明。**與線模型不一致的地方：** (1) 線模型 2026 滲透率 5%（約 90 萬套）× 媒體所稱耐世特份額 35% 會推得 2026 年中國 SBW 營收約 3.8 億美元，與耐世特 1H26 只有 2 個量產專案不符；本模型不改線值，以 2026 有效份額 2%（Assumed）吸收，並建議線模型複核 2026 滲透率口徑（量產裝車 vs 定點）。(2) 線模型未建海外 SBW，本模型自設 `os26`、`os30_*`。

## 限制與未取得資料
- SBW 營收、定點車型數、單一專案年營收未取得（公司未揭露；含在 EPS 產品線）。
- FY2023 地區分部 EBITDA、折舊攤銷、資本支出未取得（本版只取 2024 年度公告的比較期損益）；FY2025 工程研發、銷管費為年報管理層討論取整值。
- 1H 分部 EBITDA 取自中期簡報（百萬美元取整）；1H 其他與未分配 EBITDA 由合計減分部推得。
- 合資權益帳面值未取得，DCF 未計合資價值（年應佔利益約 5 百萬美元，估計影響低於 1%）；國家風險溢酬未查 Damodaran 當期值（Assumed 0.8%）。
- 同業前瞻 P/E 頁面日期 2026-09-18 到 11-13 不一，已用頁面報價與 2026-11-20 收盤等比例換算（Derived）；捷太格特不在收盤 csv，報價日未取得 11-20；伯特利頁面日期不明，不列；同業 5 家少於 8 家，不做迴歸。
- β 兩來源差距大（耐世特 5 年 β 1.85 vs 同業 0.67）；Sens 已列 ±0.3。
- 共識 FY2027 起逐年值未取得（資料商付費牆）。
- 未做：未參與製作的代理獨立核對（Project 規則第 10 條，待主代理安排）。

## 核對結果（建置後執行；指令見下）
in_layout 通過；sens_coverage（--require-keycol）通過；check_governance 全部通過；check_rating_gate --strict 符合；check_zero_rows 0 項；check_cached_values 通過；segment_check 符合；check_company_html 全部通過；check_layout_v13 通過（只有全頁參數代號警告，在摺疊外的表格標籤與 build_id）。

## 檔案
- `build_nexteer.py`：建置＋整本重算 Sens／反解＋產生本 README、`segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 核對指令
```
python3 tools/in_layout/check_in_layout.py physical_ai/nexteer/dist/{NAME}
python3 tools/sens_coverage/check_sens_coverage.py physical_ai/nexteer/dist/{NAME} --require-keycol
python3 tools/model_governance/check_governance.py physical_ai/nexteer/dist/{NAME} --html physical_ai/nexteer/dist/{NAME[:-5]}.html
python3 tools/model_governance/check_rating_gate.py physical_ai/nexteer/dist/{NAME} --strict
python3 tools/model_governance/check_zero_rows.py physical_ai/nexteer/dist/{NAME}
python3 tools/model_governance/check_cached_values.py physical_ai/nexteer/dist/{NAME}
python3 tools/segment_check/check_segments.py physical_ai/nexteer
python3 tools/company_html/check_company_html.py physical_ai/nexteer
python3 tools/company_html/check_layout_v13.py physical_ai/nexteer/dist/{NAME[:-5]}.html --xlsx physical_ai/nexteer/dist/{NAME}
```

## 版本紀錄
- v0.1（2026-11-23）：初版。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


if __name__ == "__main__":
    main()

# 優必選 UBTech（9880.HK）個股模型 v0.1 — AI實體應用 Project（線 4 人形機器人：整機）
# 用法：python3 physical_ai/ubtech/build_ubtech.py   → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、sens.json、segments.json，並跑核對程式
# 會計年度＝曆年（12 月決算；中國企業會計準則 CAS，港交所公告採用）。金額：十億元人民幣（¥B）；股數：百萬股；EPS：人民幣元；股價與目標價：港元。
# 結構：README、In、Hist、PL（三情境：命題模組→產品類別營收→標準損益表→現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 線層級變數（全球人形出貨、全尺寸售價路徑）讀 physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx「下游取數」頁（值抄入「對帳」頁並寫明儲存格）。
# 本工作適用《高不確定性預估工作原則》（docs/plan/高不確定性預估工作原則.md）：超出已公告訂單與實際交付的出貨、單價、毛利率一律標 Assumed 並給區間。
import sys, os, json, subprocess, shutil, hashlib, tempfile, glob, math
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261123"
NAME = f"{DATE}_優必選個股模型_{VER}.xlsx"
OUT = os.path.join(HERE, "dist", NAME)
TMP = os.path.join(tempfile.gettempdir(), "ubt_build")
LO = ["soffice", "-env:UserInstallation=file://" + os.path.join(tempfile.gettempdir(), "lo_ubt_profile"), "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML"]
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032
LINE_COL = {"A": "E", "B": "D", "C": "C"}   # 線模型欄：C 低、D 基準、E 高 → 本模型情境 C／B／A

# ---------------- 出處 ----------------
AR25 = "港交所 優必選 2025 年度業績公告（2026-03-31，CAS）https://www1.hkexnews.hk/listedco/listconews/sehk/2026/0331/2026033102607.pdf"
AR24 = "港交所 優必選 2024 年度業績公告（2025-03-30，CAS；FY2023 比較欄）https://www1.hkexnews.hk/listedco/listconews/sehk/2025/0330/2025033000189.pdf"
IR26 = "港交所 優必選 2026 中期業績公告（2026-08-28，CAS）https://www1.hkexnews.hk/listedco/listconews/sehk/2026/0828/2026082801691.pdf"
GAS25 = "蓋世汽車 2026-04-03：2025 年全尺寸人形銷量 1,079 台、年底產能逾 6,000 台 https://autonews.gasgoo.com/articles/ICV/ubtech-2025-report-card-revenue-from-full-size-humanoid-robots-grows-over-22-fold-2039900685372407808"
CD26 = "China Daily 2026-01-21：2025 年人形訂單 14 億元、2026 年工業人形產能預期逾 1 萬台 https://www.chinadaily.com.cn/a/202601/21/WS69708c5aa310d6866eb34f9a.html"
KR25 = "KrASIA 2025-11-20：2026 年產量目標 5,000 台、2027 年 1 萬台 https://kr-asia.com/pulses/158401"
TIG = "老虎證券新聞 2026-07-03：優世界 U1 系列訂單逾 13,361 台，售價 11.98／16.98 萬元起 https://www.itiger.com/news/1191180730"
EM0910 = "東方財富 2026-09-10：海外訂單逾 5,000 萬元（Walker C1、U1）；2026H1 人形總銷量 16,123 台 https://finance.eastmoney.com/a/202609103870470486.html"
FL1 = "界面新聞經騰訊新聞 2026-08-28：鋒龍股份 43% 股權（協議轉讓 29.99%＋要約 13.01%，每股 17.72 元）；鋒龍 H1 營收 2.67 億、毛利率 18.35% https://news.qq.com/rain/a/20260828A03OL300"
FL2 = "同花順經 21 世紀經濟報道 2026-10-30：鋒龍 2026 前三季營收 3.98 億元；總對價約 16.65 億元 https://news.10jqka.com.cn/20261030/c681283346.shtml"
SA = "stockanalysis.com（資料商；S&P Global 彙整；讀取 2026-11-23）"
CSV = "Yahoo Finance 日線（physical_ai/data/closes2_20261109_20261120.csv）"
TE = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與 AI半導體 勝宏科技、中際旭創 模型同一取值，跨模型一致）"
DAM = "Damodaran Country Risk Premium 表（2026-01 更新）：China CRP 0.91%（與 AI半導體 中國個股模型同值）"
LINE = "physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx"
CO = "發行公司（立場：自家業績與前景、募資需求）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方，賣方有承銷與交易誘因）"
MED = "媒體轉述公司說法（立場：公司宣傳）"

# ---------------- In 列 ----------------
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. 2026 上半年實績錨點（交付與單價；Verified／Derived）")
inp("u_h126", "全尺寸具身智能人形銷量 2026H1", "台", None, 921, None, "Verified", IR26 + "：「sales volume of 921 units」", "2025H1 約 45 台（Derived 換算：921 ÷ (1＋1946.7%)；增速出自 " + EM0910 + "）")
inp("asp_h1", "全尺寸人形每台營收 2026H1（含解決方案與服務）", "千元／台", None, "=h_fs_h126/u_h126*1000000", None, "Derived", "＝全尺寸人形營收 5.903 億元 ÷ 921 台 ＝ 約 64.1 萬元", "2025 全年：8.206 億 ÷ 1,079 台 ＝ 約 76.0 萬元（Derived）；含客製解決方案，非列表價")
inp("bl_25", "2025 年底待交付訂單（粗估）", "¥B", None, "=1.4-h_fs_25", None, "Derived", "＝2025 年人形訂單約 14 億元（" + CD26 + "，Interested-party）− 2025 全尺寸人形營收 8.21 億", "假設訂單全為全尺寸人形且全數轉營收；只作對照，不進模型")

sec("B. 命題模組：全尺寸人形＝全球人形出貨（線模型）× 優必選占比 × 每台營收（線模型售價路徑）")
for s, v, lo, hi, note in (("A", 2500, None, None, "樂觀：產能已逾 6,000 台（2025 年底，" + GAS25 + "），下半年驗收集中；約 H1 的 2.7 倍"),
                           ("B", 1400, 900, 2500, "基準：H1 921 台、2025 下半年約 1,034 台（Derived）；下半年約 H1 的 1.5 倍（年底驗收季節性）"),
                           ("C", 900, None, None, "保守：維持 H1 速度（2026 年新訂單總額未找到，2025 年底待交付粗估已在 H1 消化）")):
    inp(f"u26h2_{s}", f"情境 {s}：全尺寸人形銷量 2026H2", "台", lo, v, hi, "Assumed", note,
        "錨點：已公告訂單——2025 年約 14 億元（約 1,800–2,200 台）；2026 年訂單總額未找到（只有個別公告：Airbus、海外逾 5,000 萬元）" if s == "B" else "區間見情境 A／B／C 三列")
inp("asp_adj", "2026H2 每台營收相對 H1", "%", -0.15, 0.0, 0.10, "Assumed", "H1 每台 64.1 萬元，較 2025 全年 76.0 萬元低 16%（Derived）", "低 −15%：降價延續；高 +10%：大單含解決方案")
for s, v, lo, hi, note in (("A", 0.06, None, None, "樂觀：2027 年約 8,400 台（全球高情境 13.9 萬台 × 6%），接近公司 2027 年 1 萬台產量目標（" + KR25 + "，Interested-party）"),
                           ("B", 0.04, 0.025, 0.06, "基準：2026 年隱含占比約 4.2%（約 2,300 台 ÷ 線模型 5.5 萬台；Derived），2027 年 4%"),
                           ("C", 0.03, None, None, "保守：中國整機廠（宇樹、智元等）與 Tesla 量產擠壓；2027 年約 1,800 台")):
    inp(f"shr27_{s}", f"情境 {s}：優必選全尺寸人形占全球人形出貨（台數）FY2027", "%", lo, v, hi, "Assumed", note,
        "全球出貨含小型機（線模型口徑）；優必選只計全尺寸，占比為台數比" if s == "B" else "區間見情境 A／B／C 三列")
for s, v, lo, hi in (("A", 0.0, None, None), ("B", -0.0025, -0.005, 0.0), ("C", -0.004, None, None)):
    inp(f"dshr_{s}", f"情境 {s}：占比每年變化（FY2028 起）", "pt/年", lo, v, hi, "Assumed", "無公開份額時序；以中國競爭者擴產為方向（線模型：宇樹產能 7.5 萬台／年，Interested-party）", "")
inp("gm_fs26", "全尺寸人形毛利率 FY2026", "%", 0.45, 0.55, 0.65, "Derived",
    "殘差法：2025 年全公司毛利 7.54 億 − 其他類別營收 11.80 億 × 30% ＝ 4.00 億 ÷ 8.21 億 ≈ 49%；2026H1 同法（另扣鋒龍 1.39 億 × 18.35%）≈ 66%",
    "公司未揭露分類別毛利率（未取得）；結果依賴其他類別毛利率 30% 的假設；區間 45–65%")
for s, v, lo, hi in (("A", 0.45, None, None), ("B", 0.35, 0.25, 0.45), ("C", 0.25, None, None)):
    inp(f"gm_fslt_{s}", f"情境 {s}：全尺寸人形毛利率 FY2030 起（FY2027–FY2030 由 FY2026 線性收斂）", "%", lo, v, hi, "Assumed",
        "線模型整機定價毛利率 25%（Analogy，區間 15–45%；高值即優必選 2026H1 44.7%）", "售價與成本同隨學習曲線下降時毛利率不變；收斂反映價格競爭（中國同業價格戰）")

sec("C. 其他產品類別（財報按產品類別揭露的營收組成；成長路徑）")
inp("ne26", "非具身智能人形機器人營收 FY2026", "¥B", 0.065, 0.075, 0.09, "Assumed", "2026H1 0.0327（+22%，" + IR26 + "）；2025 0.0480", "下半年約同上半年＋U1 少量交付")
for s, v, lo, hi, note in (("A", 0.60, None, None, "樂觀：U1 訂單逾 13,361 台（" + TIG + "，Interested-party）約三成在 FY2027 交付 × 每台約 15 萬元"),
                           ("B", 0.15, 0.08, 0.30, "基準：U1 交付少量（約 1,000 台）"), ("C", 0.06, None, None, "保守：U1 交付延誤、退訂")):
    inp(f"ne27_{s}", f"情境 {s}：非具身智能人形機器人營收 FY2027（含優世界 U1）", "¥B", lo, v, hi, "Assumed", note, "U1 是否列入本類別未取得（公司未說明）" if s == "B" else "")
for s, v, lo, hi in (("A", 0.20, None, None), ("B", 0.10, 0.0, 0.20), ("C", 0.0, None, None)):
    inp(f"gne_{s}", f"情境 {s}：非具身智能人形營收年增 FY2028 起", "%/年", lo, v, hi, "Assumed", "無公司指引", "")
inp("or26", "其他智能機器人產品及服務營收 FY2026（教育、物流、服務）", "¥B", 0.48, 0.55, 0.62, "Assumed", "2025 0.6287；2026H1 0.2458（−19%，" + IR26 + "）", "下半年季節性高（2025 下半年 0.3258）")
for s, v, lo, hi in (("A", 0.03, None, None), ("B", -0.03, -0.08, 0.03), ("C", -0.10, None, None)):
    inp(f"gor_{s}", f"情境 {s}：其他智能機器人營收年增 FY2027 起", "%/年", lo, v, hi, "Assumed", "2024→2025 −16.9%；2026H1 −18.8%（Derived）", "資源轉向全尺寸人形")
inp("hw26", "其他智能硬件設備營收 FY2026（消費級）", "¥B", 0.47, 0.51, 0.55, "Assumed", "2025 0.4992；2026H1 0.2594（+2.9%）", "")
for s, v, lo, hi in (("A", 0.05, None, None), ("B", 0.02, -0.03, 0.05), ("C", -0.05, None, None)):
    inp(f"ghw_{s}", f"情境 {s}：其他智能硬件營收年增 FY2027 起", "%/年", lo, v, hi, "Assumed", "2024→2025 +6.4%（Derived）", "")
inp("fl26", "園林機械、汽車及液壓零部件（鋒龍股份並表）營收 FY2026", "¥B", None, "=h_fl_h126+(0.398-0.267)*2", None, "Derived",
    "並表約自 2026-04（H1 並表 1.392 億）；下半年＝鋒龍 Q3 單季（前三季 3.98 億 − H1 2.67 億 ＝ 1.31 億）× 2", FL1 + "；" + FL2)
inp("fl27", "鋒龍營收 FY2027（全年並表）", "¥B", None, "=0.398/0.75*1.04", None, "Derived", "＝前三季 3.98 億 ÷ 0.75 × (1＋4%)", "鋒龍前三季營收 +6.75%（" + FL2 + "）")
inp("gfl", "鋒龍營收年增 FY2028 起", "%/年", 0.0, 0.05, 0.08, "Assumed", "鋒龍 2026 前三季 +6.75%", "")
inp("ot26", "其他營收（每年持平）", "¥B", None, 0.004, None, "Assumed", "2025 0.0046；2026H1 0.0018", "金額小，不設區間")
inp("gm_leg", "非全尺寸各類別（非具身、其他智能機器人、硬件、其他）毛利率", "%", 0.25, 0.30, 0.35, "Assumed",
    "2024 年全尺寸人形營收僅 0.036 時，全公司毛利率 28.7%（" + AR24 + "；Derived）", "公司未揭露分類別毛利率")
inp("gm_fl", "鋒龍（零部件）毛利率", "%", 0.15, 0.1835, 0.22, "Verified", FL1 + "：H1 綜合毛利率 18.35%", "低／高 Assumed")

sec("D. 營業費用與損益表下半段")
inp("op26", "核心營業費用 FY2026（銷售＋管理＋研發＋稅金及附加＋其他營業項目；不含減值）", "¥B", 1.33, "=(h_gp_h126-(h_pbt_h126-h_nop_h126)-h_imp_h126)*2", 1.62, "Derived",
    "＝2026H1 核心營業費用 0.737 × 2（下半年同上半年）", "2025 全年 1.375（Derived）；低／高 ±10%（Assumed）")
for s, v, lo, hi in (("A", 0.15, None, None), ("B", 0.08, 0.03, 0.15), ("C", 0.03, None, None)):
    inp(f"gop_{s}", f"情境 {s}：核心營業費用年增 FY2027 起", "%/年", lo, v, hi, "Assumed", "2023→2025 核心費用約持平（研發 4.9→5.1 億、銷售 5.1→4.7 億）", "A：擴產與海外擴張；C：收縮")
inp("imp_r", "信用及資產減值 ÷ 營收", "%", 0.04, 0.06, 0.09, "Derived", "2025 9.5%、2026H1 7.5%、2024 15.3%（Derived）", "應收帳款占營收高（2025 年底 13.0 億 ÷ 營收 20.0 億）；基準取下降趨勢")
inp("nonop26", "業外（財務淨額＋其他收益）FY2026", "¥B", None, -0.03, None, "Assumed", "2026H1 −0.015（財務淨額 −0.023＋其他收益 0.008）", "")
inp("nonop", "業外 FY2027 起", "¥B", -0.05, 0.02, 0.08, "Assumed", "利息收入隨現金下降；2025 0.038、2024 0.044", "")
inp("tax26", "所得稅 FY2026", "¥B", None, 0.08, None, "Assumed", "2026H1 0.059（子公司與遞延所得稅）", "")
inp("tax_w", "有效稅率（只對正的稅前利益；虧損不認列稅盾）", "%", 0.15, 0.15, 0.25, "Analogy", "類比：中國高新技術企業 15%、法定 25%", "")
inp("nci_r", "少數股東損益 ÷ 稅後淨利", "%", 0.05, 0.08, 0.11, "Derived", "2025 11.0%（(7.898 − 7.032) ÷ 7.898）、2026H1 8.1%", "鋒龍 57% 少數股東與其他子公司")
inp("div_r", "股利 ÷ 歸屬母公司淨利", "%", None, 0.0, None, "Assumed", "虧損公司；派息紀錄未找到", "只用於收支關卡")

sec("E. 現金流")
inp("fc_da", "折舊攤銷 ÷ 營收", "%", 0.03, 0.04, 0.06, "Derived", "2025 4.3%（0.0865 ÷ 2.001）、2026H1 4.9%", "在建工程 18.4 億（2026-06-30）轉固後折舊上升，高案 6%")
inp("cx26", "資本支出 FY2026", "¥B", 0.45, 0.60, 0.80, "Assumed", "2025 6.135 億、2024 4.004 億（" + AR25 + "）", "在建工程 2026-06-30 18.43 億（" + IR26 + "）")
inp("fc_cx", "資本支出 ÷ 營收 FY2027 起", "%", 0.04, 0.06, 0.09, "Assumed", "2025 30.7%（擴產期）；產能逾 6,000 台後回落", "")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.30, 0.40, 0.55, "Derived",
    "2025：(13.02＋5.77 − 7.01) − (9.13＋4.61 − 5.48) ＝ 3.51 億 ÷ 營收增量 6.96 億 ＝ 50%（" + AR25 + "）", "營運資金 ÷ 營收：2024 63%、2025 59%；基準 40%（應收改善）")

sec("F. 估值輸入（評價日 2026-12-31；資產負債表 2026-06-30）")
inp("px", "股價（2026-11-20 收盤，港元）", "HK$", None, 57.900002, None, "Verified", CSV + "：9880.HK 2026-11-20 收盤 57.90", "同檔 2026-11-13 收盤 72.35（一週 −20.0%）")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("fx_hkd", "港元兌美元（HKD=X，2026-11-20）", "HKD/USD", None, 7.8426, None, "Verified", CSV + "：HKD=X 2026-11-20 7.8426", "")
inp("fx_cny", "人民幣兌美元（CNY=X，2026-11-20）", "CNY/USD", None, 6.6665, None, "Verified", CSV + "：CNY=X 2026-11-20 6.6665", "")
inp("fx", "港元 ÷ 人民幣", "HKD/CNY", None, "=fx_hkd/fx_cny", None, "Derived", "＝HKD=X ÷ CNY=X", "人民幣每股價值 × fx ＝ 港元")
inp("sh_v", "評價日股數（2026-06-30 已發行股份，含內資股與 H 股）", "百萬股", None, 503.401373, None, "Verified", IR26 + "：股本 503,401,373 元（每股面值 1 元）", "RSU／選擇權未行使數量未取得；2025 年股份支付 0.993 億、2026H1 為 0（Verified）→ 不另加")
inp("cash_v", "貨幣資金（2026-06-30）", "¥B", None, 2.381213, None, "Verified", IR26 + "：貨幣資金 2,381,213 千元", "MD&A 現金及現金等價物 2,325.6 百萬元（差額為受限資金等）")
inp("fvtpl_v", "交易性金融資產（2026-06-30）", "¥B", None, 0.270055, None, "Verified", IR26, "")
inp("debt_v", "借款（2026-06-30；短期 383,256＋長期 486,655＋一年內到期 127,416）", "¥B", None, "=0.383256+0.486655+0.127416", None, "Derived", IR26 + "（各列 Verified）", "租賃負債未取得")
inp("mi_v", "少數股東權益（2026-06-30，帳面）", "¥B", None, 0.892396, None, "Verified", IR26, "以帳面扣除；鋒龍 57% 少數股東按市值計更高（見 Risk）")
inp("stub_frac", "Stub 期（2026-06-30 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,6,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac（下半年約占一半，粗估）")
inp("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", TE, "")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Verified", DAM, "")
inp("beta", "β", "x", 1.2, 1.5, 2.0, "Assumed", SA + "：9880 5 年 β 3.18（上市 2023-12，未滿 3 年，不採用）", "以高成長虧損型硬體股 1.5 為基準，區間 1.2–2.0")
inp("kd", "稅前負債成本", "%", None, "=rf+0.02", None, "Assumed", "rf ＋ 2pt（2025 利息費用 0.026 ÷ 平均借款約 13.3 億 ≈ 2%）", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("dps_tgt", "評價日到目標時點之每股股利", "¥", None, 0.0, None, "Assumed", "虧損公司，不派息", "")
inp("w_dcf", "DCF 權重（其餘為分部 EV/營收加總）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺優必選對同業的歷史相對倍數）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業 EV／營收（CY2027E；股價 2026-11-20 收盤，股數與淨現金取資料商最近一期；Comparable 用）")
PEERS = [("2432", "越疆 Dobot 2432.HK（協作機器人、人形）", 14.4, 439.96, 2.31, 1.08, "2026-09-21", "sys"),
         ("9660", "地平線 Horizon Robotics 9660.HK（智駕晶片與方案）", 3.12, 15870.0, 7.59, 8.35, "2026-10-03", "sys"),
         ("2590", "極智嘉 Geek+ 2590.HK（倉儲 AMR）", 5.405, 1330.0, 2.45, 5.49, "2026-09-05", "cmp"),
         ("2498", "速騰聚創 RoboSense 2498.HK（LiDAR）", 13.32, 471.24, 2.30, 4.19, "2026-10-24", "cmp")]
for t, lab, p, sh, nc, rv, d, g in PEERS:
    inp(f"pxp_{t}", f"股價：{lab}", "HK$", None, p, None, "Verified", CSV + f"：{t}.HK 2026-11-20 收盤", "")
    inp(f"sh_{t}", f"股數：{lab}", "百萬股", None, sh, None, "Interested-party", f"{SA} statistics 頁（{d}）", "未回公司原文核對", DV)
    inp(f"nc_{t}", f"淨現金（港元十億）：{lab}", "HK$B", None, nc, None, "Interested-party", f"{SA} statistics 頁（{d}）", "最近一期資產負債表", DV)
    inp(f"rv_{t}", f"CY2027E 營收共識（人民幣十億）：{lab}", "¥B", None, rv, None, "Interested-party", f"{SA} forecast 頁「Revenue Next Year」", "賣方共識，未取得分析師家數以外細節", DV)
    inp(f"evs_{t}", f"EV／營收 CY2027E：{lab}", "x", None, f"=(pxp_{t}*sh_{t}/1000-nc_{t})/(rv_{t}*fx)", None, "Derived",
        "＝(股價 × 股數 − 淨現金) ÷ (營收 × 港元／人民幣)", "群組：" + ("系統與整機" if g == "sys" else "零組件與物流"))
inp("ms_sys", "系統與整機群組 EV／營收中位數（2 家；全尺寸與非具身人形用）", "x", None, "=MEDIAN(evs_2432,evs_9660)", None, "Derived", "少於 8 家，不做迴歸（第 10 節第 7 條）", "")
inp("ms_cmp", "零組件與物流群組 EV／營收中位數（2 家；其他機器人、硬件、鋒龍、其他用）", "x", None, "=MEDIAN(evs_2590,evs_2498)", None, "Derived", "", "")
inp("ms_all", "全體 4 家 EV／營收中位數（只作對照，不進模型）", "x", None, "=MEDIAN(evs_2432,evs_9660,evs_2590,evs_2498)", None, "Derived", "", "")
inp("cons_rev26", "共識 FY2026 營收（只作對照，不進模型）", "¥B", None, 3.8, None, "Interested-party", SA + " forecast 頁：2026 平均 3.8（3.1–4.5）；資料更新 2026-09-03", "中期業績後", DV)
inp("cons_rev27", "共識 FY2027 營收（只作對照，不進模型）", "¥B", None, 6.05, None, "Interested-party", SA + " forecast 頁「Revenue Next Year」6.05（+60.9%）", "", DV)
inp("cons_eps27", "共識 FY2027 EPS（只作對照，不進模型）", "¥", None, 1.21, None, "Interested-party", SA + " forecast 頁「EPS Next Year」1.21", "", DV)
inp("cons_tp", "賣方平均目標價（只作對照，不進模型）", "HK$", None, 143.28, None, "Interested-party", SA + "：11 位分析師平均 143.28（區間 118.36–177.26；2026-09-16 頁面，資料 2026-09-03）", "財報前後未更新者混雜", DV)
inp("ev_self", "優必選現行 EV／共識 FY2027 營收（只作對照，不進模型）", "x", None, "=(px*sh_v/1000/fx-cash_v-fvtpl_v+debt_v+mi_v)/cons_rev27", None, "Derived", "＝(市值 − 淨現金 ＋ 少數股東權益) ÷ 共識 FY2027 營收", "")

# ---------------- 線模型下游取數（抄入對帳頁；值與儲存格）----------------
LINE_ROWS = [  # (代號, 名稱, 低 C, 基準 D, 高 E, 單位, 儲存格)
    ("SHIP_2026", "全球人形出貨 2026", 40000, 55000, 75000, "台/年", "下游取數!C34:E34"),
    ("SHIP_2027", "全球人形出貨 2027", 59813.9512488488, 95502.4511385826, 139201.35065489, "台/年", "下游取數!C35:E35"),
    ("SHIP_2028", "全球人形出貨 2028", 89442.7190999916, 165831.23951777, 258360.213655276, "台/年", "下游取數!C36:E36"),
    ("SHIP_2029", "全球人形出貨 2029", 133748.060995284, 287950.724532662, 479521.209284006, "台/年", "下游取數!C37:E37"),
    ("SHIP_2030", "全球人形出貨 2030", 200000, 500000, 890000, "台/年", "下游取數!C38:E38"),
    ("SHIP_2031", "全球人形出貨 2031", 275945.932292243, 715484.540552628, 1324630.27448537, "台/年", "下游取數!C39:E39"),
    ("SHIP_2032", "全球人形出貨 2032", 380730.787743176, 1023836.25553961, 1971511.64503729, "台/年", "下游取數!C40:E40"),
    ("HUM_PRICE_2026", "全尺寸售價 2026（出廠）", None, 105238.935837246, None, "USD/台", "下游取數!D10"),
    ("HUM_PRICE_2028", "全尺寸售價 2028", 84882.1323312977, 71070.0223992946, 58131.3298712696, "USD/台", "下游取數!C77:E77"),
    ("HUM_PRICE_2030", "全尺寸售價 2030", 73018.3919642365, 53160.3076103043, 37509.416285559, "USD/台", "下游取數!C79:E79"),
    ("HUM_PRICE_2035", "全尺寸售價 2035", 54736.3789803759, 31704.4973784094, 17290.1172380907, "USD/台", "下游取數!C81:E81"),
    ("HUM_BOM_2026", "全尺寸每台 BOM 2026", 28008.5915492958, 60798.1549295775, 147470, "USD/台", "下游取數!C6:E6"),
    ("HUM_MFG_2026", "全尺寸每台製造成本 2026", None, 78929.2018779343, None, "USD/台", "下游取數!D9"),
    ("HUM_PROD_REQ", "現價下回本＝門檻所需相對生產力", None, 0.369281694835054, None, "倍", "下游取數!D13"),
]


def build(path, ov=None, sens=None, rev=None):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 ¥B（十億元人民幣）、股數百萬股、EPS 人民幣元、股價港元"
    IN["A2"] = "標記：Verified（附出處）／Interested-party（寫明是誰、誘因欄）／Analogy（寫明類比對象）／Assumed（給區間）／Derived（寫明換算）；查不到寫「未取得」或「未找到」"
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
            c = IN.cell(r, col, v)
            if col == 5: c.font = BLUE
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2); IN.cell(r, 10, inc)
        dn(k, f"In!$E${r}"); KR[k] = r; r += 1
    LAST_IN = r - 1
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("u26h2", "全尺寸人形銷量 2026H2（台）"), ("shr27", "優必選占全球人形出貨 FY2027"), ("dshr", "占比每年變化"), ("gm_fslt", "全尺寸人形毛利率 FY2030 起"),
           ("ne27", "非具身人形營收 FY2027"), ("gne", "非具身人形年增"), ("gor", "其他智能機器人年增"), ("ghw", "其他智能硬件年增"), ("gop", "核心營業費用年增")]
    for f, lab in FAM:
        IN.cell(r, 1, f + "_scn"); IN.cell(r, 2, lab)
        for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR[f + '_' + s]}")
        IN.cell(r, 8, f"引用 In 頁 {f}_A／{f}_B／{f}_C；{lab}"); r += 1
    IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重")
    for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR['w_' + s]}")
    IN.cell(r, 8, "引用 In 頁 w_A／w_B／w_C；合計＝1"); r += 1
    IN.column_dimensions["B"].width = 56; IN.column_dimensions["H"].width = 70

    # ---------------- Hist ----------------
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（CAS；¥B）。FY 為年度、H1 為半年度；公式只引用本頁儲存格（定義名稱 h_*）。營業費用與營業利益為本模型口徑：營業利益＝稅前利益 − 業外（財務淨額＋其他收益）"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "H1'25", "H1'26", "標記", "來源"])
    HC = {"23": "B", "24": "C", "25": "D", "h125": "E", "h126": "F"}
    HR = [  # (鍵, 標籤, FY23, FY24, FY25, H1'25, H1'26, 標記, 來源)
        ("rev", "營收", 1.055698, 1.305361, 2.000999, 0.621460, 1.269130, "Verified", AR24 + "；" + AR25 + "；" + IR26),
        ("fs", "全尺寸具身智能人形機器人產品及服務", None, 0.035619, 0.820557, 0.038206, 0.590299, "Verified", AR25 + "；" + IR26),
        ("ne", "非具身智能人形機器人產品及服務", None, 0.041606, 0.047963, 0.026846, 0.032725, "Verified", "同上"),
        ("or", "其他智能機器人產品及服務", None, 0.756580, 0.628656, 0.302814, 0.245754, "Verified", "同上"),
        ("hw", "其他智能硬件設備", None, 0.468998, 0.499207, 0.251967, 0.259377, "Verified", "同上"),
        ("fl", "園林機械、汽車及液壓零部件（鋒龍並表，2026 起）", None, 0.0, 0.0, 0.0, 0.139176, "Verified", IR26),
        ("ot", "其他", None, 0.002558, 0.004616, 0.001627, 0.001799, "Verified", "同上"),
        ("cogs", "營業成本", 0.722880, 0.931352, 1.247186, 0.404148, 0.702252, "Verified", "同上"),
        ("sell", "銷售費用", 0.506113, 0.523930, 0.471011, 0.223534, 0.238034, "Verified", "同上"),
        ("ga", "管理費用", 0.399706, 0.370006, 0.335623, 0.185102, 0.149346, "Verified", "同上"),
        ("rd", "研發費用", 0.490502, 0.478108, 0.507483, 0.218281, 0.303083, "Verified", "同上"),
        ("imp", "信用及資產減值損失", 0.171371, 0.199233, 0.190711, 0.001773, 0.094841, "Verified", "信用＋資產減值兩列相加（各列 Verified）"),
        ("oinc", "其他收益", 0.027220, 0.037635, 0.036952, 0.004585, 0.008288, "Verified", "同上"),
        ("fin", "財務收益（費用）淨額", -0.009011, 0.005983, 0.000594, 0.010229, -0.023282, "Verified", "同上"),
        ("pbt", "稅前利益（虧損）", -1.227229, -1.149259, -0.774680, -0.435132, -0.279787, "Verified", "同上"),
        ("tax", "所得稅", 0.037361, 0.010665, 0.015143, 0.004857, 0.059040, "Verified", "同上"),
        ("nia", "稅後淨利（虧損）", -1.264590, -1.159924, -0.789823, -0.439989, -0.338827, "Verified", "同上"),
        ("ni", "歸屬母公司淨利（虧損）", -1.234048, -1.123590, -0.703191, -0.413648, -0.311478, "Verified", "同上"),
        ("eps", "基本每股虧損（元）", -3.05, -2.67, -1.55, -0.94, -0.62, "Verified", "同上"),
        ("wsh", "加權平均股數（百萬股）", None, 421.322, 454.862, 438.524, 502.585, "Verified", "同上"),
        ("da", "折舊攤銷", None, 0.085147, 0.086522, 0.044370, 0.062207, "Verified", "費用性質附註"),
        ("cx", "資本支出", None, 0.4004, 0.6135, None, None, "Verified", AR25 + "（MD&A）"),
        ("sbc", "股份支付費用", None, 0.245326, 0.099300, 0.072205, 0.0, "Verified", "同上"),
        ("ar", "應收帳款（期末）", None, 0.913498, 1.302224, None, 1.679607, "Verified", "資產負債表"),
        ("inv", "存貨（期末）", None, 0.460500, 0.576526, None, 0.985473, "Verified", "同上"),
        ("ap", "應付帳款（期末）", None, 0.547808, 0.701444, None, 0.883795, "Verified", "同上"),
        ("u", "全尺寸人形銷量（台）", None, None, 1079, None, 921, "Verified", GAS25 + "；" + IR26),
    ]
    HROW = {}
    for i, row in enumerate(HR, 3):
        k, lab = row[0], row[1]
        H.cell(i, 1, lab)
        for j, v in enumerate(row[2:7]):
            c = H.cell(i, 2 + j, v if v is not None else "未取得")
            if v is not None: c.font = BLUE
            dn(f"h_{k}_{list(HC)[j]}", f"Hist!${'BCDEF'[j]}${i}")
        H.cell(i, 7, row[7]); H.cell(i, 8, row[8]); HROW[k] = i
    i = 3 + len(HR)
    H.cell(i, 1, "毛利（Derived）"); [H.cell(i, 2 + j, f"={L}{HROW['rev']}-{L}{HROW['cogs']}") for j, L in enumerate("BCDEF")]
    for j, L in enumerate("BCDEF"): dn(f"h_gp_{list(HC)[j]}", f"Hist!${L}${i}")
    HROW["gp"] = i; i += 1
    H.cell(i, 1, "業外＝財務淨額＋其他收益（Derived）"); [H.cell(i, 2 + j, f"={L}{HROW['fin']}+{L}{HROW['oinc']}") for j, L in enumerate("BCDEF")]
    for j, L in enumerate("BCDEF"): dn(f"h_nop_{list(HC)[j]}", f"Hist!${L}${i}")
    HROW["nop"] = i; i += 1
    H.cell(i, 1, "營業利益（本模型口徑＝稅前 − 業外；Derived）"); [H.cell(i, 2 + j, f"={L}{HROW['pbt']}-{L}{HROW['nop']}") for j, L in enumerate("BCDEF")]
    HROW["op"] = i; i += 1
    H.cell(i, 1, "營業費用合計（＝毛利 − 營業利益；Derived）"); [H.cell(i, 2 + j, f"={L}{HROW['gp']}-{L}{HROW['op']}") for j, L in enumerate("BCDEF")]
    HROW["opx"] = i; i += 1
    H.cell(i, 1, "其他營業項目（稅金及附加、投資收益、公允價值變動、營業外收支；差額，Derived）")
    for j, L in enumerate("BCDEF"):
        H.cell(i, 2 + j, f"={L}{HROW['opx']}-{L}{HROW['sell']}-{L}{HROW['ga']}-{L}{HROW['rd']}-{L}{HROW['imp']}")
    HROW["oth"] = i; i += 1
    H.cell(i, 1, "檢查：產品類別合計 − 營收（FY2024 起）")
    for j, L in enumerate("CDEF"):
        H.cell(i, 3 + j, f"=ROUND(SUM({L}{HROW['fs']}:{L}{HROW['ot']})-{L}{HROW['rev']},6)")
    i += 1
    H.cell(i, 1, "營運資金（應收＋存貨 − 應付）÷ 年化營收（Derived）")
    H.cell(i, 3, f"=(C{HROW['ar']}+C{HROW['inv']}-C{HROW['ap']})/C{HROW['rev']}"); H.cell(i, 4, f"=(D{HROW['ar']}+D{HROW['inv']}-D{HROW['ap']})/D{HROW['rev']}")
    H.cell(i, 6, f"=(F{HROW['ar']}+F{HROW['inv']}-F{HROW['ap']})/(F{HROW['rev']}*2)"); i += 1
    H.cell(i + 1, 1, "FY2023 產品類別營收、折舊、資本支出、營運資金：未取得（本輪只讀 FY2024 年度業績公告的比較欄損益表）。鋒龍 2026-03 起逐步並表（協議轉讓 2026-03 完成、要約 2026-04 完成）。")
    H.column_dimensions["A"].width = 56; H.column_dimensions["H"].width = 60

    # ---------------- 對帳（線層級變數）----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = "對帳 — 線 4 人形機器人：本模型讀取的線層級變數（值抄自 " + LINE + "「下游取數」頁，標明儲存格；線低／基準／高 → 本模型情境 C／B／A）"
    hdr(RE, 2, ["代號", "名稱", "線低（C 欄）", "線基準（D 欄）", "線高（E 欄）", "單位", "來源檔與儲存格", "本模型用途", "標記"])
    LR = {}
    for i, (k, lab, lo, b, hi, u, cell) in enumerate(LINE_ROWS, 3):
        RE.cell(i, 1, k); RE.cell(i, 2, lab)
        for j, v in enumerate((lo, b, hi)):
            if v is not None: RE.cell(i, 3 + j, v).font = BLUE
        RE.cell(i, 6, u); RE.cell(i, 7, LINE + " " + cell)
        use = ("PL 全球人形出貨列（情境 A 讀高、B 讀基準、C 讀低）" if k.startswith("SHIP") else
               "PL 每台營收價格指數（÷ HUM_PRICE_2026 基準）" if k.startswith("HUM_PRICE") else "說明與對照，未進公式")
        RE.cell(i, 8, use); RE.cell(i, 9, "Derived（線模型輸出）"); LR[k] = i
    i = 3 + len(LINE_ROWS) + 1
    RE.cell(i, 1, "本模型情境對應（Derived；PL 引用本區）").font = BOLD; i += 1
    hdr(RE, i, ["項目", "年度", "情境 A（線高）", "情境 B（線基準）", "情境 C（線低）", "說明"]); i += 1
    MAP = {}
    for y in range(2026, 2033):
        RE.cell(i, 1, "全球人形出貨（台）"); RE.cell(i, 2, y)
        for j, s in enumerate(SC): RE.cell(i, 3 + j, f"={LINE_COL[s]}{LR['SHIP_' + str(y)]}")
        RE.cell(i, 6, "直接讀線模型 SHIP 列"); MAP[("ship", y)] = i; i += 1
    p26 = f"$D${LR['HUM_PRICE_2026']}"
    for y in range(2026, 2033):
        RE.cell(i, 1, "全尺寸售價指數（÷ 2026 基準）"); RE.cell(i, 2, y)
        for j, s in enumerate(SC):
            L = LINE_COL[s]; r28, r30, r35 = (f"{L}{LR['HUM_PRICE_' + t]}/{p26}" for t in ("2028", "2030", "2035"))
            f = {2026: "=1", 2027: f"=SQRT({r28})", 2028: f"={r28}", 2029: f"=SQRT({r28}*{r30})", 2030: f"={r30}",
                 2031: f"={r30}*({r35}/({r30}))^(1/5)", 2032: f"={r30}*({r35}/({r30}))^(2/5)"}[y]
            RE.cell(i, 3 + j, f)
        RE.cell(i, 6, "2028、2030、2035 讀線模型；2027、2029 幾何內插；2031–2032 由 2030→2035 幾何內插"); MAP[("pidx", y)] = i; i += 1
    i += 1
    RE.cell(i, 1, "一致性對照（Derived）").font = BOLD; i += 1
    hdr(RE, i, ["項目", "本模型", "線模型", "比", "說明"]); i += 1
    RE.cell(i, 1, "2026H1 每台營收（美元）vs 線模型 2026 全尺寸出廠售價"); RE.cell(i, 2, "=asp_h1*1000/fx_cny"); RE.cell(i, 3, f"=D{LR['HUM_PRICE_2026']}"); RE.cell(i, 4, f"=B{i}/C{i}")
    RE.cell(i, 5, "優必選含解決方案與服務；線模型為出廠硬體售價"); RE_ASP = i; i += 1
    RE.cell(i, 1, "優必選 FY2026 全尺寸銷量占線模型 2026 全球出貨（情境 B）"); RE.cell(i, 2, "=PL!F{u}"); RE.cell(i, 3, f"=D{LR['SHIP_2026']}"); RE.cell(i, 4, f"=B{i}/C{i}")
    RE.cell(i, 5, "線模型全球出貨含小型機；優必選只計全尺寸"); RE_SH = i; i += 1
    RE.cell(i, 1, "與線模型不一致的自設值：無。全尺寸毛利率（FY2026 起點 55%、FY2030 起情境 B 35%）高於線模型定價毛利率 25%——線模型為整機廠出廠定價假設，優必選營收含解決方案；以殘差法推得，標 Derived／Assumed 並列區間。相對人工生產力無已驗證實例（線模型「現價下回本所需相對生產力」為門檻值，非實績），本模型不據以推需求。")
    RE.column_dimensions["A"].width = 50; RE.column_dimensions["B"].width = 34; RE.column_dimensions["G"].width = 60; RE.column_dimensions["H"].width = 50

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（全尺寸人形＝全球出貨 × 占比 × 每台營收）→ 產品類別營收 → 標準損益表 → 現金流（¥B；FY2023–FY2025 實績、FY2026 起模型；CAS）"
    PR = {}; rr = 3
    labs = [("u", "全尺寸人形：銷量（台）"), ("glob", "全尺寸人形：全球人形出貨（台；線模型）"), ("shr", "全尺寸人形：占全球出貨比例"), ("pidx", "全尺寸人形：售價指數（線模型；2026＝1）"),
            ("asp", "全尺寸人形：每台營收（千元）"),
            ("fs", "全尺寸具身智能人形機器人產品及服務 營收"), ("ne", "非具身智能人形機器人產品及服務 營收"), ("or", "其他智能機器人產品及服務 營收"),
            ("hw", "其他智能硬件設備 營收"), ("fl", "園林機械、汽車及液壓零部件 營收"), ("ot", "其他 營收"),
            ("rev", "營收（合併）"), ("chk", "檢查：分部合計 − 合併營收"), ("gr", "營收年增率"),
            ("gmfs", "全尺寸人形毛利率"), ("cogs", "營業成本"), ("gp", "毛利"), ("gm", "毛利率"),
            ("core", "核心營業費用（銷售＋管理＋研發＋其他營業項目）"), ("imp", "信用及資產減值損失"), ("opx", "營業費用（合計）"),
            ("op", "營業利益"), ("opm", "營業利益率"), ("nop", "業外（財務淨額＋其他收益）"), ("pbt", "稅前利益"), ("tax", "所得稅"),
            ("nia", "稅後淨利"), ("nci", "少數股東損益"), ("ni", "歸屬母公司淨利"), ("nim", "淨利率"), ("nig", "淨利年增率"),
            ("sh", "平均股數（百萬股）"), ("eps", "EPS（基本＝稀釋，元）"),
            ("da", "折舊攤銷"), ("cx", "資本支出"), ("dwc", "營運資金增加"), ("fcff", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）"),
            ("div", "股利"), ("fcfad", "自由現金流扣股利（收支關卡用）")]
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, (k, _) in enumerate(labs)}
        for k, l in labs: P.cell(R[k], 1, l)
        lc = LINE_COL[s]
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1)
            cell = lambda k: f"{c}{R[k]}"; prev = lambda k: f"{p}{R[k]}"
            def w(k, f): P[cell(k)] = f
            if y <= 2025:
                t = str(y)[2:]
                for k in ("fs", "ne", "or", "hw", "fl", "ot"):
                    if y >= 2024: w(k, f"=h_{k}_{t}")
                w("rev", f"=h_rev_{t}")
                if y >= 2024: w("chk", f"=ROUND(SUM({c}{R['fs']}:{c}{R['ot']})-{c}{R['rev']},6)")
                if y == 2025:
                    w("u", "=h_u_25"); w("asp", f"={cell('fs')}/{cell('u')}*1000000")
                w("cogs", f"=h_cogs_{t}"); w("imp", f"=h_imp_{t}"); w("nop", f"=h_nop_{t}"); w("pbt", f"=h_pbt_{t}")
                w("gp", f"={cell('rev')}-{cell('cogs')}"); w("op", f"={cell('pbt')}-{cell('nop')}"); w("opx", f"={cell('gp')}-{cell('op')}")
                w("core", f"={cell('opx')}-{cell('imp')}")
                w("tax", f"=h_tax_{t}"); w("nia", f"=h_nia_{t}"); w("ni", f"=h_ni_{t}"); w("nci", f"={cell('nia')}-{cell('ni')}"); w("eps", f"=h_eps_{t}")
                if y >= 2024: w("sh", f"=h_wsh_{t}"); w("da", f"=h_da_{t}"); w("cx", f"=h_cx_{t}")
                if y == 2025: w("dwc", "=(h_ar_25+h_inv_25-h_ap_25)-(h_ar_24+h_inv_24-h_ap_24)")
            elif y == 2026:
                w("u", f"=u_h126+u26h2_{s}"); w("glob", f"=對帳!{'CDE'[SC.index(s)]}{MAP[('ship', y)]}"); w("shr", f"={cell('u')}/{cell('glob')}")
                w("pidx", f"=對帳!{'CDE'[SC.index(s)]}{MAP[('pidx', y)]}")
                w("fs", f"=h_fs_h126+u26h2_{s}*asp_h1*(1+asp_adj)/1000000"); w("asp", f"={cell('fs')}/{cell('u')}*1000000")
                w("ne", "=ne26"); w("or", "=or26"); w("hw", "=hw26"); w("fl", "=fl26"); w("ot", "=ot26")
                w("gmfs", "=gm_fs26"); w("core", "=op26"); w("nop", "=nonop26"); w("tax", "=tax26"); w("sh", "=sh_v"); w("cx", "=cx26")
            else:
                w("glob", f"=對帳!{'CDE'[SC.index(s)]}{MAP[('ship', y)]}"); w("shr", f"=MAX(0,shr27_{s}+dshr_{s}*({y}-2027))")
                w("u", f"={cell('glob')}*{cell('shr')}"); w("pidx", f"=對帳!{'CDE'[SC.index(s)]}{MAP[('pidx', y)]}")
                w("asp", f"=asp_h1*(1+asp_adj)*{cell('pidx')}"); w("fs", f"={cell('u')}*{cell('asp')}/1000000")
                w("ne", f"=ne27_{s}" if y == 2027 else f"={prev('ne')}*(1+gne_{s})")
                w("or", f"={prev('or')}*(1+gor_{s})"); w("hw", f"={prev('hw')}*(1+ghw_{s})")
                w("fl", "=fl27" if y == 2027 else f"={prev('fl')}*(1+gfl)"); w("ot", "=ot26")
                w("gmfs", f"=gm_fs26+(gm_fslt_{s}-gm_fs26)*MIN(1,({y}-2026)/4)")
                w("core", f"={prev('core')}*(1+gop_{s})"); w("nop", "=nonop"); w("tax", f"=MAX(0,{cell('pbt')})*tax_w"); w("sh", "=sh_v"); w("cx", f"=fc_cx*{cell('rev')}")
            if y >= 2026:
                w("rev", f"=SUM({c}{R['fs']}:{c}{R['ot']})")
                w("chk", f"=ROUND(SUM({c}{R['fs']}:{c}{R['ot']})-{c}{R['rev']},6)")
                w("cogs", f"={cell('fs')}*(1-{cell('gmfs')})+({cell('ne')}+{cell('or')}+{cell('hw')}+{cell('ot')})*(1-gm_leg)+{cell('fl')}*(1-gm_fl)")
                w("gp", f"={cell('rev')}-{cell('cogs')}")
                w("imp", f"=imp_r*{cell('rev')}"); w("opx", f"={cell('core')}+{cell('imp')}"); w("op", f"={cell('gp')}-{cell('opx')}")
                w("pbt", f"={cell('op')}+{cell('nop')}"); w("nia", f"={cell('pbt')}-{cell('tax')}"); w("nci", f"={cell('nia')}*nci_r"); w("ni", f"={cell('nia')}-{cell('nci')}")
                w("eps", f"={cell('ni')}/{cell('sh')}*1000")
                w("da", f"=fc_da*{cell('rev')}"); w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
                w("fcff", f"={cell('op')}-MAX(0,{cell('op')})*tax_w+{cell('da')}-{cell('cx')}-{cell('dwc')}")
                w("div", f"=MAX(0,{cell('ni')})*div_r"); w("fcfad", f"={cell('fcff')}-{cell('div')}")
            if y >= 2024:
                w("gr", f"={cell('rev')}/{prev('rev')}-1")
                w("nig", f'=IF({prev("ni")}<=0,"N/M（前一年虧損）",{cell("ni")}/{prev("ni")}-1)')
            w("gm", f"={cell('gp')}/{cell('rev')}"); w("opm", f"={cell('op')}/{cell('rev')}"); w("nim", f"={cell('ni')}/{cell('rev')}")
        P.cell(R["u"], 2, "FY2026＝H1 實績 921 ＋ H2 假設；FY2027 起＝全球出貨 × 占比")
        P.cell(R["asp"], 2, "FY2027 起＝H1 每台營收 ×(1＋H2 調整)× 線模型售價指數")
        P.cell(R["gmfs"], 2, "FY2026 起點為殘差法推得（非揭露值）；FY2030 起為情境值")
        P.cell(R["core"], 2, "實績＝營業費用合計 − 減值；FY2026＝H1 × 2；之後 ×(1＋年增)")
        P.cell(R["op"], 2, "本模型口徑：稅前 − 業外（含稅金及附加、投資收益等於營業費用內）")
        P.cell(R["tax"], 2, "FY2027 起只對正的稅前利益課稅，虧損不認列稅盾")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 − MAX(0, 營業利益) × tax_w；股份支付不加回")
        for k in ("shr", "gr", "gmfs", "gm", "opm", "nim", "nig"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        for y in YRS:
            P[f"{COL[y]}{R['u']}"].number_format = "#,##0"; P[f"{COL[y]}{R['glob']}"].number_format = "#,##0"; P[f"{COL[y]}{R['asp']}"].number_format = "#,##0"
            P[f"{COL[y]}{R['pidx']}"].number_format = "0.000"
        PR[s] = R; rr = b + len(labs) + 2
    P.column_dimensions["A"].width = 52; P.column_dimensions["B"].width = 36
    # 對帳頁一致性列引用 PL（情境 B FY2026 銷量）
    RE.cell(RE_SH, 2, f"=PL!F{PR['B']['u']}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)；人民幣計，每股價值 × fx 轉港元"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（¥B；港元市值 ÷ fx）"; D["B4"] = "=px*sh_v/1000/fx"
    D["A5"] = "負債 D（借款）"; D["B5"] = "=debt_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(上限, 無風險利率)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金（2026-06-30）＝貨幣資金＋交易性金融資產 − 借款 − 少數股東權益（帳面）"; D["B10"] = "=cash_v+fvtpl_v-debt_v-mi_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC"; D["B11"] = "=ROUND(((px*sh_v/1000/fx)*(rf+beta*(erp+crp))+debt_v*kd*(1-tax_w))/((px*sh_v/1000/fx)+debt_v)+wacc_adj-wacc,9)"
    for rr_ in range(3, 11): D[f"B{rr_}"].number_format = "0.00%" if rr_ in (3, 6, 7, 8, 9) else "#,##0.00"
    DR = {}; rr = 13
    for s in SC:
        R = PR[s]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}" for y in range(2027, 2033)]); b = rr + 2
        rows = ["EBIT（營業利益）", "NOPAT", "＋折舊攤銷", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF", "折現年數（期中）", "折現因子", "FCFF 現值"]
        for i, l in enumerate(rows): D.cell(b + i, 1, l)
        for j, y in enumerate(range(2027, 2033)):
            c = "BCDEFG"[j]; pc = COL[y]
            D[f"{c}{b}"] = f"=PL!{pc}{R['op']}"; D[f"{c}{b+1}"] = f"={c}{b}-MAX(0,{c}{b})*tax_w"
            D[f"{c}{b+2}"] = f"=PL!{pc}{R['da']}"; D[f"{c}{b+3}"] = f"=-PL!{pc}{R['cx']}"; D[f"{c}{b+4}"] = f"=-PL!{pc}{R['dwc']}"
            D[f"{c}{b+5}"] = f"=PL!{pc}{R['fcff']}"; D[f"{c}{b+6}"] = f"=ROUND(SUM({c}{b+1}:{c}{b+4})-{c}{b+5},9)"
            D[f"{c}{b+7}"] = j + 0.5; D[f"{c}{b+8}"] = f"=1/(1+wacc)^{c}{b+7}"; D[f"{c}{b+9}"] = f"={c}{b+5}*{c}{b+8}"
        t = b + 11
        items = [("明確期現值合計", f"=SUM(B{b+9}:G{b+9})"),
                 ("NOPAT FY2033＝NOPAT FY2032 ×(1＋g)", f"=G{b+1}*(1+g_tv)"),
                 ("終值（FY2032 年底）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^6"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金（2026-06-30，已扣少數股東權益）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!F{R['fcff']}*stub_frac"),
                 ("股權價值（¥B）", f"=B{t+4}+B{t+5}+B{t+6}"),
                 ("每股價值（評價日，人民幣元）", f"=MAX(0,B{t+7})/sh_v*1000"),
                 ("DCF 情境價值推到 12 個月後（港元；×(1＋ke)^t − 期間股利，× fx）", f"=(B{t+8}*(1+ke)^t_tgt-dps_tgt)*fx"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M（EV≤0）\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=IF(G{b+1}>0,B{t+2}/G{b+1},\"N/M（NOPAT≤0）\")")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 9, ps=t + 8, ev=t + 4, tvs=t + 10, chk=b + 6)
        dn(f"dcf_{s}", f"DCF!$B${t+9}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 62

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = '="Val — 目標價（港元）：DCF（權重 "&TEXT(w_dcf,"0%")&"）＋ 分部 EV／營收加總（Comparable，權重 "&TEXT(1-w_dcf,"0%")&"）；三情境權重 "&TEXT(w_A,"0%")&"／"&TEXT(w_B,"0%")&"／"&TEXT(w_C,"0%")&"；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"'
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}"); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後（港元）", [f"=dcf_{s}" for s in SC], "DCF 頁")
    cf = lambda s, k, c="H": f"PL!{c}{PR[s][k]}"
    r_ev = vrow("分部 EV（FY2028 營收 × 群組 EV／營收，¥B）", [f"=({cf(s,'fs')}+{cf(s,'ne')})*ms_sys+({cf(s,'or')}+{cf(s,'hw')}+{cf(s,'fl')}+{cf(s,'ot')})*ms_cmp" for s in SC],
                "全尺寸＋非具身人形 × 系統與整機群組；其他類別 × 零組件與物流群組；目標時點（2027-11）的 NTM ≈ FY2028")
    r_nct = vrow("目標時點淨現金（¥B）＝2026-06-30 淨現金 ＋ Stub FCFF ＋ FY2027 FCFF", [f"=dcf_nc+{cf(s,'fcff','F')}*stub_frac+{cf(s,'fcff','G')}" for s in SC], "虧損期現金消耗計入")
    r_cu = vrow("Comparable（未調整）：分部 EV／營收加總（港元／股）", [f"=MAX(0,{L}{r_ev}+{L}{r_nct})/sh_v*1000*fx" for L in "BCD"], "股權價值 ÷ 評價日股數 × fx")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "相對倍數係數 1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（DCF 權重 × DCF ＋ (1 − DCF 權重) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（港元）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權；分部 EV／營收加總）", f"=E{r_ca}")
    r_px = one("現價（In 頁 px，港元）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（¥B；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_cov = one("收支缺口 ÷ 2026-06-30 淨現金（情境 B；>1＝現金不足以支應，需再融資）", f"=IF(B{r_vg}<0,-B{r_vg}/(cash_v+fvtpl_v-debt_v),0)", "收支型提示：再融資（配售）會稀釋；模型未建融資模組")
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級（港元）", f'=TEXT(px*(1+IF(B{r_up}>0.5,9,IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5))))),"#,##0.0")&"／"&IF(B{r_up}<-0.5,"無（已為最低級）",TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,-0.5)))),"#,##0.0"))', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_evs = one("隱含 EV／營收：加權目標價隱含 EV ÷ 情境 B FY2028 營收", f"=(B{r_tp}/fx*sh_v/1000-dcf_nc)/PL!H{PR['B']['rev']}", "分母：模型情境 B FY2028 營收")
    r_self = one("現行 EV／共識 FY2027 營收 ÷ 系統與整機群組中位數", "=ev_self/ms_sys", "只作對照")
    for rr_ in (r_up, r_a1, r_a2, r_a3): V[f"B{rr_}"].number_format = "+0.0%;-0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0.00"
    for rr_ in (r_ev, r_nct):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0.00"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "#,##0.00"
    V.column_dimensions["A"].width = 72; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, ca=r_ca, ev=r_ev, nct=r_nct, pd2=r_pd2, evs=r_evs, gj=r_gj, cov=r_cov, wk=r_wk, self=r_self, px=r_px)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司說法、市場共識、線模型對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司、共識或線模型", "差", "說明"])
    rc = [("FY2026 營收 vs 共識", f"=PL!F{PR['B']['rev']}", "=cons_rev26", "資料商共識（2026-09-03，中期業績後；Interested-party）"),
          ("FY2027 營收 vs 共識", f"=PL!G{PR['B']['rev']}", "=cons_rev27", "同上"),
          ("FY2027 EPS vs 共識", f"=PL!G{PR['B']['eps']}", "=cons_eps27", "同上；共識為正、模型情境 B 仍虧損"),
          ("FY2026 全尺寸人形銷量 vs 公司 2026 年產量目標 5,000 台", f"=PL!F{PR['B']['u']}", 5000, "產量目標為 Interested-party（KrASIA 2025-11-20 轉述公司）；2026 年工業人形產能預期逾 1 萬台（China Daily 2026-01-21）"),
          ("FY2027 全尺寸人形銷量 vs 公司 2027 年產量目標 1 萬台", f"=PL!G{PR['B']['u']}", 10000, "同上；目標≠銷量"),
          ("2026H1 每台營收（美元）vs 線模型 2026 出廠售價", f"=對帳!B{RE_ASP}", f"=對帳!C{RE_ASP}", "口徑不同：公司含解決方案；只看量級"),
          ("加權目標價 vs 賣方平均目標價（港元）", "=tp", "=cons_tp", "賣方 11 位平均（2026-09，Interested-party）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，建置腳本格點掃描＋線性內插；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    sv = [it for it in (rev or []) if it["key"] == "shr27_B" and isinstance(it["val"], (int, float))]
    if sv:
        i += 1
        RC.cell(i, 1, "對照：反解占比對應的情境 B FY2027 全尺寸人形銷量（台）").font = BOLD; i += 1
        hdr(RC, i, ["項目", "銷量（台）", "公司 2027 年產量目標（台）", "比", "說明"]); i += 1
        for it in sv:
            RC.cell(i, 1, "反解：" + it["label"]); RC.cell(i, 2, f"={it['val']}*對帳!D{MAP[('ship', 2027)]}"); RC.cell(i, 3, 10000); RC.cell(i, 4, f"=B{i}/C{i}")
            RC.cell(i, 5, "＝反解占比 × 線模型 2027 全球出貨（基準）；產量目標為 Interested-party（KrASIA 2025-11-20 轉述公司）"); i += 1
    RC.column_dimensions["A"].width = 62; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高或指定值），整本重建並以 LibreOffice 重算後讀加權目標價（不是閉式近似）；結果為常數快照"
    SE["A2"] = "build_id"; SE["A3"] = "模型狀態"; SE["A4"] = "In 頁輸入指紋（Sens 計算時）"; SE["A5"] = "In 頁輸入指紋（本檔即時）"
    SE["D5"] = f"=SUMPRODUCT(In!E4:E{LAST_IN},ROW(In!E4:E{LAST_IN}))"
    hdr(SE, 7, ["驅動", "設定", "鍵", "輸入值", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    if sens:
        SE["D2"] = sens["build_id"]; SE["D4"] = sens["fp"]
        SE["D3"] = f'=IF(AND(ABS(D5-D4)<1E-6,ABS(Val!B{r_tp}-{sens["base_tp"]})<0.001),"CURRENT","STALE")'
        sr = 8
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            if cs.get("val") is not None: SE.cell(sr, 4, cs["val"])
            SE.cell(sr, 5, cs["tp"]).number_format = "#,##0.00"; SE.cell(sr, 6, cs["tp"] / sens["base_tp"] - 1).number_format = "+0.0%;-0.0%"
            SE.cell(sr, 7, cs["dcf"]).number_format = "#,##0.00"
            SE.cell(sr, 8, (cs["dcf"] / sens["base_dcf"] - 1) if sens["base_dcf"] else 0).number_format = "+0.0%;-0.0%"
            SE.cell(sr, 9, cs["rt"]); sr += 1
        sr += 1
        SE.cell(sr, 1, "讀法：相對基準＝案例加權目標價 ÷ 基準加權目標價 − 1；同一案例 DCF 單法與加權目標價並列（原則第 10 節第 32 條）。情境專屬驅動以「情境 B 改為情境 A／C 值」呈現。")
        SE.cell(sr + 1, 1, "說明：折舊、資本支出、營運資金隨營收等比例變動（第 44 條）；全尺寸人形售價與成本同隨線模型學習曲線下降（毛利率另設）。相對倍數係數 rel_k 基準即為 1.00（未調整），「相對倍數係數＝1」案例＝基準。")
    SE.column_dimensions["A"].width = 56; SE.column_dimensions["B"].width = 40

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("再融資稀釋（收支缺口需配售或借款）", "否（Val 頁列缺口 ÷ 淨現金）", "模型未建融資模組；DCF 以負 FCFF 直接減值，等同按公允價格融資", "港交所配售公告、現金餘額、在建工程"),
             ("客戶集中（2026H1 單一客戶 3.085 億、約占營收 24%）", "否", "尾端 Sens 未建（第 19 條，下版補）", "中報主要客戶揭露、應收帳款逾期"),
             ("應收帳款回收（應收 16.8 億 ÷ 年化營收）", "部分（imp_r、fc_wc）", "減值率與營運資金率 Sens", "信用減值、應收周轉天數"),
             ("付費採用多為展示、教育與政府專案；相對人工生產力無已驗證實例", "部分（u26h2、shr27 區間）", "出貨與占比為 Assumed", "客戶揭露的部署規模、續購訂單、工廠單位工時產出"),
             ("中國人形價格戰（宇樹、智元等）", "是（gm_fslt、dshr）", "情境 C 與 Sens", "同業售價、每台營收"),
             ("鋒龍少數股東權益按市值計遠高於帳面（43% 持股的 A 股市值高於收購對價）", "否", "Comparable 與 DCF 以帳面扣少數股東權益；鋒龍營運價值以並表營收 × 群組倍數計", "鋒龍股價、新鋒龍注入計畫"),
             ("政策補貼與政府採購變化", "否", "無可量化資料", "地方政府人形機器人專案招標")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 56; RK.column_dimensions["C"].width = 60

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([AR25, AR24, IR26, GAS25, CD26, KR25, TIG, EM0910, FL1, FL2, SA + " statistics／forecast 頁（9880、2432、9660、2590、2498）", CSV, TE, DAM, LINE + "「下游取數」頁"], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字）----------------
    RD["A1"] = f"優必選 UBTech（9880.HK）個股模型 {VER}｜build_id {{BID}}"
    RD["A2"] = '="資料截止日 2026-11-23（財報、公告與資料商）；成品日 2026-11-23；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.00")&" 港元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：優必選的股價由全尺寸人形機器人（Walker S 系列）能否從展示與專案訂單走到規模出貨決定——全尺寸人形營收＝全球人形出貨（線模型）× 優必選占比 × 每台營收（線模型售價學習曲線）；其他產品類別（非具身人形、教育物流服務機器人、消費硬件、鋒龍零部件）以成長路徑建模"
    RD["A4"] = (f'="在 2026 下半年全尺寸人形銷量情境 B "&TEXT(u26h2_B,"#,##0")&" 台、FY2027 占全球人形出貨 "&TEXT(shr27_B,"0.0%")&"、全尺寸毛利率由 "&TEXT(gm_fs26,"0%")&" 收斂到 "&TEXT(gm_fslt_B,"0%")&" 的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 港元（DCF "&TEXT(Val!B{r_dcft},"0.00")&"、分部 EV／營收加總 "&TEXT(Val!B{r_cmpt},"0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0.00")&" 十億元）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 港元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：全尺寸人形半年銷量超過 2,500 台且每台營收不低於 50 萬元；公司揭露 2026 年新訂單總額超過 30 億元；全公司毛利率連兩個半年度超過 45% 且核心營業費用年增低於 10%；同業 EV／營收倍數重估到 5 倍以上"
    RD["A7"] = "範圍：本模型只讀線 4 人形機器人線模型（全球出貨、全尺寸售價路徑）；不讀 Siliconomics／Tokenomics／Wattonomics（算力線），自用 AI 讀法、對帳中旗標不適用。未建：融資模組、分類別毛利率（未揭露）、季度分部頁（公司只發半年報）。HTML 展示層見 dist/ 同名 .html"
    for i in range(1, 8): RD.cell(i, 1).alignment = Alignment(wrap_text=True)
    RD.column_dimensions["A"].width = 160
    wb.save(path)
    return dict(VR=VR, PR=PR, DR=DR, KR=KR, LAST_IN=LAST_IN, MAP=MAP)


# ---------------- 重算與讀值 ----------------
def recalc(paths, outdir):
    os.makedirs(outdir, exist_ok=True)
    for k in range(0, len(paths), 12):
        chunk = paths[k:k + 12]
        for attempt in (1, 2):
            try:
                subprocess.run(LO + ["--outdir", outdir] + chunk, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
                break
            except subprocess.TimeoutExpired:
                subprocess.run(["pkill", "-f", "lo_ubt_profile"]); print("soffice 逾時，重試", attempt)
                if attempt == 2: raise


def read_vals(path, meta):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; VR = meta["VR"]
    return dict(tp=V[f"B{VR['tp']}"].value, dcf=V[f"B{VR['dcf']}"].value, cmp=V[f"B{VR['cmp']}"].value, up=V[f"B{VR['up']}"].value, rt=V[f"B{VR['rt']}"].value)


def run_cases(cases, sub):
    d = os.path.join(TMP, sub); shutil.rmtree(d, ignore_errors=True); raw = os.path.join(d, "raw"); rc = os.path.join(d, "rc"); os.makedirs(raw)
    paths = []; meta = None
    for i, ov in enumerate(cases):
        p = os.path.join(raw, f"c{i:03d}.xlsx"); meta = build(p, ov); paths.append(p)
    recalc(paths, rc)
    return [read_vals(os.path.join(rc, f"c{i:03d}.xlsx"), meta) for i in range(len(cases))]


def in_values(path, last):
    wb = openpyxl.load_workbook(path, data_only=True); ws = wb["In"]; out = {}
    for r in range(4, last + 1):
        k = ws.cell(r, 1).value
        if isinstance(k, str) and ws.cell(r, 5).value is not None:
            out[k] = (ws.cell(r, 4).value, ws.cell(r, 5).value, ws.cell(r, 6).value, ws.cell(r, 2).value)
    return out


def rating(u):
    return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"


def main():
    os.makedirs(os.path.join(HERE, "dist"), exist_ok=True)
    shutil.rmtree(TMP, ignore_errors=True); os.makedirs(TMP)
    base_raw = os.path.join(TMP, "base.xlsx"); meta = build(base_raw)
    recalc([base_raw], os.path.join(TMP, "base_rc")); base_rc = os.path.join(TMP, "base_rc", "base.xlsx")
    base = read_vals(base_rc, meta); print("基準：", base)
    iv = in_values(base_rc, meta["LAST_IN"]); ivE = {k: v[1] for k, v in iv.items()}
    fp = openpyxl.load_workbook(base_rc, data_only=True)["Sens"]["D5"].value
    if "--base-only" in sys.argv: return
    if "--post" in sys.argv:   # 只重跑展示層、核對與 README（沿用 sens.json 與現有成品）
        sens = json.load(open(os.path.join(HERE, "sens.json"))); revs = sens["revs"]
        meta = build(os.path.join(TMP, "meta.xlsx"), {}, sens, revs); errs = []
        return post(meta, sens, errs)
    specs = []
    for k, (lo, e, hi, lab) in iv.items():
        for side, val in (("敏感度低", lo), ("敏感度高", hi)):
            if isinstance(val, (int, float)) and isinstance(e, (int, float)) and abs(val - e) > 1e-12:
                specs.append(dict(drv=f"{lab}（{side}）", set=f"{k}＝{val:.4g}（基準 {e:.4g}）", key=k, val=val, ov={k: val}))
    for k in ("u26h2", "shr27", "dshr", "gm_fslt", "ne27", "gop"):
        for s in ("A", "C"):
            specs.append(dict(drv=f"情境 B 的「{iv[k + '_B'][3].split('：', 1)[1]}」改為情境 {s} 值", set=f"{k}_B＝{ivE[k + '_' + s]:.4g}", key=f"{k}_B", val=ivE[k + "_" + s], ov={f"{k}_B": ivE[k + "_" + s]}))
    specs.append(dict(drv="全部分部改用系統與整機群組倍數（零組件群組＝系統群組）", set=f"ms_cmp＝{ivE['ms_sys']:.4g}", key="ms_cmp", val=ivE["ms_sys"], ov={"ms_cmp": ivE["ms_sys"]}))
    specs.append(dict(drv="全部分部改用零組件與物流群組倍數", set=f"ms_sys＝{ivE['ms_cmp']:.4g}", key="ms_sys", val=ivE["ms_cmp"], ov={"ms_sys": ivE["ms_cmp"]}))
    specs.append(dict(drv="相對倍數係數＝1（基準即為 1）", set="rel_k＝1", key="rel_k", val=1.0, ov={"rel_k": 1.0}))
    for s in SC:
        specs.append(dict(drv=f"只看情境 {s}", set=f"w_{s}＝1，其餘 0", key=f"w_{s}", val=None, ov={f"w_{x}": (1.0 if x == s else 0.0) for x in SC}))
    res = run_cases([s["ov"] for s in specs], "sens")
    cases = []
    for s, r in zip(specs, res):
        cases.append(dict(drv=s["drv"], set=s["set"], key=s["key"], val=s["val"], tp=r["tp"], dcf=r["dcf"], rt=r["rt"]))
        print(f"  {s['drv'][:44]:44s} {r['tp']:8.2f} {r['tp'] / base['tp'] - 1:+.1%} {r['rt']}")
    # 反解：格點掃描（整本重算）＋線性內插
    px = ivE["px"]; up0 = base["up"]
    tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    revs = []
    GRID = {"shr27_B": ([0.02, 0.04, 0.06, 0.08, 0.10, 0.13, 0.16, 0.20, 0.25, 0.30, 0.40, 0.50], "情境 B 的 FY2027 優必選占全球人形出貨（FY2028 起每年變化不動；情境 A／C 不動）"),
            "gm_fslt_B": ([0.15, 0.25, 0.35, 0.45, 0.55, 0.65, 0.75, 0.85], "情境 B 全尺寸人形 FY2030 起毛利率（情境 A／C 不動）"),
            "gop_B": ([-0.10, -0.05, 0.0, 0.03, 0.05, 0.08, 0.12, 0.15], "情境 B 核心營業費用年增 FY2027 起（情境 A／C 不動）")}
    for key, (grid, note) in GRID.items():
        rr = run_cases([{key: x} for x in grid], "solve_" + key); ups = [r["up"] for r in rr]
        for lab, tgt in (("目標價＝現價（上檔 0%）", 0.0), (f"升一級邊界（上檔 {upper:+.0%}）" if upper is not None else None, upper)):
            if lab is None: continue
            x = None
            for j in range(len(grid) - 1):
                a, b_ = ups[j] - tgt, ups[j + 1] - tgt
                if a == 0: x = grid[j]; break
                if a * b_ < 0: x = grid[j] + (grid[j + 1] - grid[j]) * (-a) / (b_ - a); break
            revs.append(dict(label=lab, key=key, val=x if x is not None else f"格點 {grid[0]:+.0%}～{grid[-1]:+.0%} 內無解", base=ivE[key], note=note))
            print("反解", key, lab, x)
    h = hashlib.sha1(json.dumps([list(map(str, x)) for x in ROWS], ensure_ascii=False).encode() + open(__file__, "rb").read()).hexdigest()[:8]
    bid = f"{VER}+{h}"
    sens = dict(build_id=bid, fp=fp, base_tp=base["tp"], base_dcf=base["dcf"], base=base, cases=cases, revs=revs)
    json.dump(sens, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
    fin_raw = os.path.join(TMP, "final", NAME); os.makedirs(os.path.dirname(fin_raw), exist_ok=True)
    meta = build(fin_raw, {}, sens, revs)
    wb = openpyxl.load_workbook(fin_raw); wb["README"]["A1"] = wb["README"]["A1"].value.replace("{BID}", bid); wb.save(fin_raw)
    recalc([fin_raw], os.path.join(TMP, "final_rc"))
    for f in glob.glob(os.path.join(HERE, "dist", "*.xlsx")):
        if f != OUT: print("注意：dist 另有檔案", f)
    shutil.copy(os.path.join(TMP, "final_rc", NAME), OUT)
    fin = read_vals(OUT, meta); print("成品：", OUT, fin, "build_id", bid)
    wbv = openpyxl.load_workbook(OUT, data_only=True)
    ERR = ("#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!")
    errs = [f"{ws.title}!{c.coordinate}={c.value}" for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value in ERR or c.value.startswith("Err:"))]
    print("錯誤值：", len(errs), errs[:10])
    return post(meta, sens, errs)


def post(meta, sens, errs):
    json.dump(dict(company="優必選 UBTech", ticker="9880.HK", model_sheet="PL",
                   official_segments=[dict(name=n, label=l, period="FY2025", value=v, unit="¥B（十億元人民幣）", source=AR25 + "（按產品類別的營收分拆）")
                                      for n, l, v in (("全尺寸具身智能人形機器人產品及服務", "全尺寸具身智能人形機器人產品及服務 營收", 0.820557),
                                                      ("非具身智能人形機器人產品及服務", "非具身智能人形機器人產品及服務 營收", 0.047963),
                                                      ("其他智能機器人產品及服務", "其他智能機器人產品及服務 營收", 0.628656),
                                                      ("其他智能硬件設備", "其他智能硬件設備 營收", 0.499207),
                                                      ("園林機械、汽車及液壓零部件", "園林機械、汽車及液壓零部件 營收", 0.0),
                                                      ("其他", "其他 營收", 0.004616))],
                   sublines=[], check_labels=["檢查：分部合計 − 合併營收"],
                   note="公司年報與中報按產品類別揭露營收（2026 起新增鋒龍並表的園林機械、汽車及液壓零部件）；應報導分部的正式劃分未取得，以產品類別為營收分部。FY2023 產品類別未取得（PL 頁 FY2023 只有合併營收）。Hist 頁另有 FY2024、FY2025、H1'25、H1'26 產品類別合計 − 營收核對列（皆為 0）。"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)
    # HTML 展示層
    HTML = OUT[:-5] + ".html"
    for cmd in (["physical_ai/ubtech/gen_html_json.py"], ["tools/company_html/build_company_html.py", HERE]):
        subprocess.run([sys.executable, os.path.join(ROOT, cmd[0])] + cmd[1:], cwd=ROOT, check=True)
    rc = 1 if errs else 0; checks = [("錯誤值掃描（建置腳本）", f"錯誤值 {len(errs)} 格", 0 if not errs else 1)]
    for cmd in (["tools/in_layout/check_in_layout.py", OUT], ["tools/sens_coverage/check_sens_coverage.py", OUT, "--require-keycol"],
                ["tools/model_governance/check_governance.py", OUT, "--html", HTML], ["tools/model_governance/check_rating_gate.py", OUT, "--strict"],
                ["tools/model_governance/check_zero_rows.py", OUT], ["tools/model_governance/check_cached_values.py", OUT],
                ["tools/segment_check/check_segments.py", HERE], ["tools/company_html/check_company_html.py", HERE],
                ["tools/company_html/check_layout_v13.py", HTML, "--xlsx", OUT]):
        p = subprocess.run([sys.executable, os.path.join(ROOT, cmd[0])] + cmd[1:], cwd=ROOT, capture_output=True, text=True)
        out = (p.stdout + p.stderr).strip().replace(ROOT + "/", "")
        print("==>", cmd[0], "結束碼", p.returncode); print(out[-1500:])
        lines = [l for l in out.splitlines() if l.strip()]
        summ = " ".join(lines[-3:]) if lines else ""
        checks.append((os.path.relpath(os.path.join(ROOT, cmd[0]), ROOT) + (" --require-keycol" if "--require-keycol" in cmd else "") + (" --strict" if "--strict" in cmd else ""), summ[:400], p.returncode))
        rc |= p.returncode
    import write_readme_ubt
    write_readme_ubt.write(OUT, meta, sens, checks)
    sys.exit(rc)


if __name__ == "__main__":
    sys.path.insert(0, HERE)
    main()

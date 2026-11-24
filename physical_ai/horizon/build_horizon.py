# 地平線 Horizon Robotics（9660.HK）個股模型 v0.1 — AI實體應用 Project（線 2 智慧車：ADAS／自駕晶片與方案）
# 用法：python3 physical_ai/horizon/build_horizon.py          → 建置、整本重算 Sens 與反解、輸出 dist/ 成品、README.md、segments.json、sens.json，並跑核對
#       python3 physical_ai/horizon/build_horizon.py --quick  → 沿用既有 sens.json 重建成品（不重跑敏感度）
# 會計年度＝曆年（12 月決算，IFRS）。金額：十億人民幣（RMB bn）；出貨：百萬套；股數：百萬股；EPS：人民幣；股價、目標價：港元。
# 結構：README、In、Hist、半年、PL（三情境：命題模組→營收線→標準損益表→現金流）、DCF、Val、Recon、Sens、對帳（線層級變數）、Risk、Sources。
# 線層級變數一律讀 physical_ai/lines/smart_car/dist/20261116_智慧車單位經濟_v0.1.xlsx（下游取數頁；下游取數未列者讀輸入頁並註明）。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261123"
NAME = f"{DATE}_地平線個股模型_{VER}.xlsx"
LINE_X = os.path.join(ROOT, "physical_ai/lines/smart_car/dist/20261116_智慧車單位經濟_v0.1.xlsx")
LINE_R = "physical_ai/lines/smart_car/dist/20261116_智慧車單位經濟_v0.1.xlsx"
CSV_R = "physical_ai/data/closes2_20261109_20261120.csv"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032
LO = "file://" + os.path.join(os.environ.get("HOME") or tempfile.gettempdir(), "lo_horizon_profile")

# ---------------- 出處 ----------------
IR26 = "公司 2026 中期業績公告（港交所 2026-08-31）https://www1.hkexnews.hk/listedco/listconews/sehk/2026/0831/2026083101009.pdf"
AR25 = "公司 2025 年度業績公告（港交所 2026-03-19）https://www1.hkexnews.hk/listedco/listconews/sehk/2026/0319/2026031900927.pdf"
CARI = "公司公告：CARIAD 可轉換貸款修訂協議交割、發行換股股份（港交所 2026-08-03）https://www1.hkexnews.hk/listedco/listconews/sehk/2026/0803/2026080303525.pdf"
CBN = "公司 2027 年到期 4.5 億美元零息可轉債定價（老虎證券轉述公告，2026-07）https://www.itiger.com/news/1184978776"
CNF = "中國金融信息網轉述 2026 中報（2026-08-31）https://www.cnfin.com/gs-lb/detail/20260831/4463004_1.html"
CXG = "財新 Caixin Global 2026-09-02 https://www.caixinglobal.com/2026-09-02/horizon-robotics-losses-widen-as-rd-spending-stays-high-102481105.html"
HXS = "華興證券 2026-09-25（同花順轉述）https://stock.10jqka.com.cn/20260925/c680283093.shtml"
GTJ = "國泰海通 2026-07-03（智通財經轉述）https://www.zhitongcaijing.com/content/detail/1462860.html"
GSR = "高盛 2026-09-09（同花順轉述）https://stock.10jqka.com.cn/20260909/c679742484.shtml"
HSB = "滙豐 2026-07-28（長橋轉述）https://longbridge.com/news/294004940"
TIG = "老虎證券異動快訊 2026-06-26（程式產生稿）https://www.itiger.com/news/1102538443"
ETN = "經濟通 etnet 9660 基本資料（2026-11 讀取）https://content.etnet.com.hk/content/iq/compinfo/tc/quote_ci_basicdata.php?code=9660"
SA_HZ = "stockanalysis.com 9660 statistics（2026-10-02 收盤頁）https://stockanalysis.com/quote/hkg/9660/statistics/"
SA_IS = "stockanalysis.com 9660 income statement（讀取 2026-11-23）https://stockanalysis.com/quote/hkg/9660/financials/income-statement/"
SA_MB = "stockanalysis.com MBLY statistics（2026-10-26 頁）https://stockanalysis.com/stocks/mbly/statistics/"
SA_AM = "stockanalysis.com AMBA statistics（2026-10-23 頁）https://stockanalysis.com/stocks/amba/statistics/"
SA_HS = "stockanalysis.com HSAI statistics（2026-11-20 收盤頁）https://stockanalysis.com/stocks/hsai/statistics/"
SA_RS = "stockanalysis.com 2498 statistics（2026-10-23 頁）https://stockanalysis.com/quote/hkg/2498/statistics/"
YCL = f"Yahoo Finance 日線（{CSV_R}）"
CNRF = "Trading Economics 中國 10 年期國債殖利率 2026-11-05 1.61%（與 AI半導體 勝宏科技模型同一取值，跨模型一致）"
DAM = "Damodaran Country Risk Premium（2026-01 更新）：China CRP 0.91%（與勝宏科技模型同值）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整（立場：中立彙整，數字來源為公司或賣方）"
BRK = "賣方券商（立場：承銷與交易關係，偏多）"
MED = "媒體轉述（立場：轉述公司或研究機構）"

# ---------------- 線模型取數 ----------------
LW = openpyxl.load_workbook(LINE_X, data_only=True)
LD, LI = LW["下游取數"], LW["輸入"]
SCOL = {"A": "E", "B": "D", "C": "C"}           # 下游取數：C 低／D 基準／E 高 → 情境 C／B／A
ICOL = {"A": "F", "B": "E", "C": "D"}           # 輸入：D 低情境／E 基準／F 高情境
PEN_ROW = {2026: 27, 2027: 28, 2028: 29, 2030: 30, 2032: 31}
LINE = []   # (鍵, 項目, 情境, 值, 單位, 來源格, 標記, 用在哪)
def lv(ws, cell): return ws[cell].value
for y, r in PEN_ROW.items():
    for s in SC:
        LINE.append((f"pen{y}_{s}", f"中國城市 NOA 滲透率 {y}", s, lv(LD, f"{SCOL[s]}{r}"), "%", f"下游取數!{SCOL[s]}{r}", "Assumed" if y > 2026 else "Interested-party", "PL 城市 NOA 新車數"))
LINE.append(("cars26", "中國城市 NOA 新車數 2026（基準）", "B", lv(LD, "D32"), "萬台", "下游取數!D32", "Derived", "÷ 滲透率 2026 → 國內零售 2026"))
for s in SC:
    LINE.append((f"cng_{s}", "中國乘用車國內銷量年增率 2027–2032", s, lv(LI, f"{ICOL[s]}5"), "%/年", f"輸入!{ICOL[s]}5（下游取數未列）", "Assumed", "PL 國內零售"))
    LINE.append((f"edec_{s}", "電子內容年降價率", s, lv(LI, f"{ICOL[s]}31"), "%/年", f"輸入!{ICOL[s]}31（下游取數未列；下游取數 D25→D26 城市 NOA 每車內容 780→559 隱含同一降幅）", "Assumed", "PL 每套產品方案營收"))
    LINE.append((f"lic_{s}", "城市 NOA 方案／授權費（每車）", s, lv(LI, f"{ICOL[s]}26"), "USD/車", f"輸入!{ICOL[s]}26（下游取數未列）", "Assumed", "PL 授權－城市 NOA 權利金"))
LINE.append(("cont26", "城市 NOA 每車內容 2026（基準）", "B", lv(LD, "D25"), "USD/車", "下游取數!D25", "Assumed", "說明用：地平線每套營收占比對照（Recon）"))
LINE.append(("cont30", "城市 NOA 每車內容 2030（基準）", "B", lv(LD, "D26"), "USD/車", "下游取數!D26", "Assumed", "說明用"))
LINE.append(("cars30", "中國城市 NOA 新車數 2030（基準）", "B", lv(LD, "D33"), "萬台", "下游取數!D33", "Derived", "Recon：與本模型同口徑新車數對照"))
LKEY = {k: i for i, (k, *_) in enumerate(LINE)}

# ---------------- In 列 ----------------
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. 2026 上半年實績（2026-01～06；持續經營，D-Robotics 2026-03-31 起不再併表，列停業單位）")
inp("h1_p", "產品方案營收 H1'26", "RMB bn", None, 0.925627, None, "Verified", IR26 + "：925,627 千元（H1'25 806,574）", "")
inp("h1_l", "授權及服務營收 H1'26", "RMB bn", None, 1.129357, None, "Verified", IR26 + "：1,129,357（H1'25 739,493）", "")
inp("h1_cogs", "營業成本 H1'26", "RMB bn", None, 0.698847, None, "Verified", IR26 + "：698,847（H1'25 525,643）", "")
inp("gm_p_h1", "產品方案毛利率 H1'26", "%", None, 0.362, None, "Verified", CXG + "：產品方案毛利率 36.2%（年減 8pt），授權及服務 90.4%（轉述公告）", "與公告營業成本、兩條營收相符（授權毛利率由此推得 90.4%）")
inp("h1_rnd", "研發費用 H1'26", "RMB bn", None, 2.754997, None, "Verified", IR26 + "：2,754,997（H1'25 2,259,525）", "")
inp("h1_sga", "銷售及管理費用 H1'26", "RMB bn", None, "=0.354496+0.327667", None, "Derived", IR26 + "：銷售 354,496＋管理 327,667（Verified 原文相加）", "")
inp("h1_imp", "金融資產減值 H1'26", "RMB bn", None, 0.077012, None, "Verified", IR26, "")
inp("h1_oi", "其他收入＋其他利得淨額 H1'26", "RMB bn", None, "=0.175069+0.311462", None, "Derived", IR26 + "：其他收入 175,069＋其他利得 311,462（原文相加）", "")
inp("h1_fin", "財務收益淨額 H1'26", "RMB bn", None, 0.166161, None, "Verified", IR26, "")
inp("h1_assoc", "權益法投資損益 H1'26", "RMB bn", None, -2.074286, None, "Verified", IR26 + "：主要為 D-Robotics 分拆後依權益法認列（其優先股公允價值變動）", "")
inp("h1_fv", "優先股及其他金融負債公允價值變動 H1'26", "RMB bn", None, 5.240607, None, "Verified", IR26 + "：主要為 CARIAD 可轉換貸款隨股價下跌的公允價值利得", "非現金；2026-08-03 交割後不再發生")
inp("h1_tax", "所得稅 H1'26", "RMB bn", None, 0.045848, None, "Verified", IR26, "")
inp("h1_disc", "停業單位損益 H1'26（D-Robotics）", "RMB bn", None, 2.168735, None, "Verified", IR26 + "：含分拆利得 2,778,754 與 Q1 虧損約 610,000", "一次性")
inp("h1_ni", "歸屬母公司淨利 H1'26", "RMB bn", None, 3.784345, None, "Verified", IR26 + "：3,784,345；基本 EPS 0.27、稀釋 −0.09；經調整淨虧損 1,671,232", "")
inp("h1_wsh", "加權平均股數 H1'26", "百萬股", None, 14054.233, None, "Verified", IR26 + "：14,054,233 千股", "")
inp("h1_units", "征程系列出貨 H1'26", "百萬套", None, 2.218, None, "Interested-party", CNF + "：221.8 萬套（+12.1%）", "公司中報揭露、媒體轉述", CO)
inp("h1_cityn", "中國城區 NOA（含選裝、預埋）新車 H1'26", "百萬台", None, 2.0469, None, "Interested-party", "高工智能汽車研究院 2026H1：204.69 萬輛（線模型 追蹤指標!F8 轉述）", "研究機構", "研究機構（立場：產業研究，販售資料）")
inp("h1_retail", "中國乘用車國內零售 H1'26", "百萬台", None, 8.61, None, "Interested-party", "新浪汽車 2026-08-26：H1 城市 NOA 162 萬／861 萬（線模型 追蹤指標!F8 轉述）", "分母 861 萬", MED)

sec("B. FY2026 錨點：下半年相對上半年（公司無數字指引；Assumed）")
inp("k2_p", "產品方案 H2 ÷ H1（FY2026）", "x", 0.90, 1.05, 1.20, "Assumed", "H2'25 ÷ H1'25＝(1,622.3 − 806.6) ÷ 806.6＝1.01（" + AR25 + "、" + IR26 + "，Derived 參照）", "中國乘用車 1–10 月零售 −22.3%（線模型追蹤指標）；華興：全年 HSD 距年初 30 萬套預期仍有差距")
inp("k2_l", "授權及服務 H2 ÷ H1（FY2026）", "x", 1.10, 1.35, 1.62, "Assumed", "H2'25 ÷ H1'25＝(1,934.9 − 739.5) ÷ 739.5＝1.62（Derived 參照；授權認列集中下半年）", "低 1.10：授權費隨新車型定點時點波動")
inp("k2_x", "營業費用 H2 ÷ H1（FY2026）", "x", 1.00, 1.10, 1.25, "Assumed", "FY2025 研發 5,153.7 含 D-Robotics，H2／H1 無同口徑拆分（未取得）", "")
inp("cont25", "FY2025 車載營收（持續經營口徑：產品方案＋授權及服務）", "RMB bn", None, "=1.622274+1.934913", None, "Derived", AR25 + "：產品方案 1,622,274＋授權及服務 1,934,913", "非車載 201,081 隨 D-Robotics 轉停業單位；FY2026 年增率以此為基期")

sec("C. 命題模組：產品方案＝城市 NOA 套數 × 每套營收 ＋ 其他 ADAS 套數 × 每套營收；授權＝城市 NOA 權利金 ＋ 其他授權及服務")
inp("s_city26", "地平線城區 NOA 晶片份額 2026", "%", None, 0.228, None, "Interested-party", CNF + "：城區 NOA 晶片份額 22.8%（2025 年 17.9%），次於英偉達", "線模型 追蹤指標!D11：英偉達 39.43%、地平線 22.82%（高工）", "研究機構（立場：產業研究）")
for s, v, lo, hi, note in (("A", 0.32, None, None, "樂觀：份額由 17.9%（2025）→ 22.8%（2026H1）的升勢延續；J6P／HSD、J7（2027H2 推出，華興）放量"),
                           ("B", 0.25, 0.18, 0.30, "基準：小幅上升；英偉達 Thor、華為、車廠自研（比亞迪璇璣 A3、小鵬、蔚來、理想）分食"),
                           ("C", 0.15, None, None, "保守：車廠自研晶片（比亞迪 4nm 璇璣 A3，老虎快訊 2026-06-26）擴大，份額回落")):
    inp(f"s_city30_{s}", f"情境 {s}：地平線城區 NOA 份額 2030（2026→2030 線性，之後持平）", "%", lo, v, hi, "Assumed", note, "區間見情境 A／B／C 三列（A 32%、C 15%）" if s == "B" else "")
for s, v, lo, hi in (("A", 0.01, None, None), ("B", 0.0, -0.02, 0.01), ("C", -0.02, None, None)):
    inp(f"d_soth_{s}", f"情境 {s}：其他 ADAS（非城區 NOA 新車）份額每年變化 FY2027 起", "pt/年", lo, v, hi, "Assumed",
        CNF + "：國產品牌 ADAS 份額首次突破 50%（2026H1）", "基期份額＝H1'26 出貨推得（s_oth26）；含出口與海外車型，口徑為「相當於中國非城區 NOA 新車」")
inp("p_oth", "其他 ADAS 每套產品方案營收 2026（J2／J3／J6E/M 入門）", "RMB/套", 150, 200, 260, "Assumed", HXS + "：公司稱 2026 年晶片均價 40–50 美元，低階 20–30、中階 90 以上、高階約 500（Interested-party）", "產品方案含晶片＋參考硬體；入門款以 25–35 美元×匯率估；城區 NOA 每套營收由產品方案總額減去本項校準（Derived）")
for s, v, lo, hi in (("A", 0.10, None, None), ("B", 0.05, 0.0, 0.10), ("C", 0.0, None, None)):
    inp(f"mix_{s}", f"情境 {s}：城區 NOA 每套營收的組合上移（J6P／HSD、J7）每年", "%/年", lo, v, hi, "Assumed",
        HXS + "：2027 年中高階占比提升，均價可能到 90–100 美元以上（約翻倍；Interested-party 轉述公司）", "與線模型電子內容年降價率相乘；B 5% 低於券商轉述的翻倍（2026 HSD 只 4–5 萬套）")
inp("byd_k", "比亞迪改用自研晶片：FY2028 起其他 ADAS 套數流失比例（尾端事件；基準 0）", "%", None, 0.0, 0.5, "Assumed",
    TIG + "：比亞迪為核心客戶、每年約 250 萬顆 J6（程式產生稿，未查核，不作依據）；比亞迪 4nm 璇璣 A3 量產", "原則第 10 節第 19 條：單一客戶占營收一成以上的事件型風險只進 Sens（高 50%）")
for s, v, lo, hi in (("A", 0.35, None, None), ("B", 0.25, 0.10, 0.40), ("C", 0.10, None, None)):
    inp(f"g_ls27_{s}", f"情境 {s}：其他授權及服務年增 FY2027", "%", lo, v, hi, "Assumed", "H1'26 授權及服務 +52.7%（" + IR26 + "）；FY2025 +17.4%（" + AR25 + "，Derived）", "含演算法授權、開發服務、Carizon（大眾合資）相關收入；未揭露拆分")
for s, v, lo, hi in (("A", 0.28, None, None), ("B", 0.18, 0.08, 0.28), ("C", 0.05, None, None)):
    inp(f"g_ls28_{s}", f"情境 {s}：其他授權及服務年增 FY2028–FY2030", "%/年", lo, v, hi, "Assumed", "Carizon 方案 2027 起搭載大眾三家中國合資 7 款車（" + CXG + "）", "")
for s, v, lo, hi in (("A", 0.12, None, None), ("B", 0.08, 0.04, 0.12), ("C", 0.03, None, None)):
    inp(f"g_lslt_{s}", f"情境 {s}：其他授權及服務年增 FY2031–FY2032", "%/年", lo, v, hi, "Assumed", "長期收斂", "")

sec("D. 毛利率與費用")
for s, v, lo, hi in (("A", 0.40, None, None), ("B", 0.36, 0.30, 0.42), ("C", 0.30, None, None)):
    inp(f"gm_p_{s}", f"情境 {s}：產品方案毛利率 FY2028 起（FY2026 H1 實績線性到 FY2028）", "%", lo, v, hi, "Assumed", "H1'26 36.2%、H1'25 約 44%（年減 8pt，" + CXG + "）", "C：中高階晶片價格戰；A：J6P／J7 高階占比")
inp("gm_l", "授權及服務毛利率", "%", 0.85, "=1-(h1_cogs-h1_p*(1-gm_p_h1))/h1_l", 0.93, "Derived", "＝1 −（營業成本 − 產品方案 ×(1 − 36.2%)）÷ 授權及服務＝90.4%（與財新轉述一致）", "低／高 Assumed")
inp("g_rnd27", "研發費用年增 FY2027", "%", 0.0, 0.05, 0.12, "Assumed", "H1'26 +21.9%（" + IR26 + "，Derived）；創辦人目標約 2028 年損益兩平（" + CXG + "，Interested-party）", "研發費用全期三情境相同（不隨營收情境分化：研發是爭奪份額的成本）")
inp("g_rndlt", "研發費用年增 FY2028 起", "%/年", -0.02, 0.03, 0.07, "Assumed", "同上", "")
inp("g_sga", "銷售及管理費用年增 FY2027 起", "%/年", 0.0, 0.04, 0.08, "Assumed", "H1'26 銷售 +41.4%、管理 +18.1%（Derived）", "")
inp("oi", "其他收入及利得（政府補助、投資收益）FY2027 起每年", "RMB bn", 0.15, 0.35, 0.50, "Assumed", "FY2025 其他收入 0.335＋其他利得 0.422（" + AR25 + "）；H1'26 0.487", "利得含金融資產公允價值變動，波動大；基準只計約補助水準")
inp("assoc27", "權益法投資損益 FY2027（D-Robotics、Carizon 等；FY2028 減半、FY2029 起 0）", "RMB bn", -0.8, -0.4, 0.0, "Assumed", "FY2025 −0.853、H1'26 −2.074（含 D-Robotics 優先股公允價值變動；" + IR26 + "）", "非現金，不進 FCFF；只影響 EPS")
inp("y_cash", "淨現金收益率（財務收益 ÷ 期初淨現金）", "%", 0.015, 0.02, 0.03, "Derived", "FY2025 財務收益淨額 0.390 ÷ 期初現金 15.37＝2.5%；H1'26 0.166×2 ÷ 20.8＝1.6%（Derived）", "低／高 Assumed")
inp("tax_w", "有效稅率（只對正的稅前利益）", "%", 0.10, 0.15, 0.25, "Analogy", "類比：中國高新技術企業優惠稅率 15%（法定 25%）", "累積虧損抵減未建模（偏保守）")

sec("E. 現金流")
inp("fc_da", "折舊攤銷 ÷ 營收", "%", 0.05, 0.07, 0.09, "Interested-party", SA_IS + "：FY2025 D&A 286.2 百萬 ÷ 營收 3,758＝7.6%；FY2024 13.1%", "法定現金流量表未取得（中期公告無現金流量表）", DV)
inp("fc_cx", "資本支出 ÷ 營收", "%", 0.04, 0.07, 0.10, "Assumed", "資本承擔 2026-06-30 534.9 百萬（" + IR26 + "）", "實績未取得（公告無現金流量表）；無晶圓廠（Fabless），資本支出以研發設備、測試與辦公為主")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.30, 0.50, 0.70, "Derived", "（應收＋存貨 − 應付 − 合約負債）÷ 營收：2025-12 (1.768＋1.069 − 0.234 − 0.266) ÷ 3.758＝62%；2026-06 (2.438＋1.318 − 0.173 − 0.230) ÷ TTM 4.267＝79%（" + IR26 + "）", "取增量比 50%（應收帳款天數上升趨勢若持續則偏低）")
inp("pay_r", "配息率", "%", None, 0.0, None, "Verified", IR26 + "：董事會不建議派發中期股息；FY2025 無股息", "收支關卡＝FCFF − 股利")

sec("F. 估值輸入（評價日 2026-12-31；資產負債表 2026-06-30 加計期後交易）")
inp("px", "股價（2026-11-20 收盤）", "HKD", None, 3.12, None, "Verified", YCL + "：9660.HK 2026-11-20 收盤 3.12", "同檔 11-09 收盤 3.675；etnet 市值 538.9 億港元 ÷ 15,873.5 百萬股＝3.40（較早日期，量級相符）")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("cny", "美元兌人民幣（CNY=X，2026-11-20）", "RMB/USD", None, 6.6665, None, "Verified", YCL + "：CNY=X 2026-11-20 收盤 6.6665", "")
inp("hkd", "美元兌港元（HKD=X，2026-11-20）", "HKD/USD", None, 7.8426, None, "Verified", YCL + "：HKD=X 2026-11-20 收盤 7.8426", "")
inp("fx", "人民幣 ÷ 港元（每 1 港元的人民幣）", "RMB/HKD", None, "=cny/hkd", None, "Derived", "＝CNY=X ÷ HKD=X", "")
inp("sh_out", "已發行股數（不含庫存股，2026-08-03 CARIAD 換股後）", "百萬股", None, 15873.486671, None, "Verified", CARI + "：15,873,486,671 股（A 類 2,112,766,542＋B 類 13,760,720,129，" + ETN + "）", "含 CARIAD 換股 1,301,763,486 股；2026-04～05 回購 80,161,800 股列庫存股")
inp("sh_rsu", "受限股份單位與選擇權（庫藏股法淨股數）", "百萬股", 0.0, 400.0, 800.0, "Assumed", "FY2025 股份支付 0.993 RMB bn（" + AR25 + "）", "股數未取得（年報股份激勵明細未讀）；DCF 標準規格第 3 節要求 RSU 全額、選擇權庫藏股法；本列為暫值（約 2.5%）")
inp("sh_v", "評價日股數", "百萬股", None, "=sh_out+sh_rsu", None, "Derived", "DCF 標準規格第 3 節", "可轉債 635.64 百萬股（換股價 5.55 港元，高於現價）視為負債，不計入")
inp("cash_v", "現金及約當現金＋定期存款（2026-06-30）", "RMB bn", None, "=14.866064+0.180698", None, "Derived", IR26 + "：現金 14,866,064＋定存 180,698（Verified 原文相加）", "")
inp("cb_v", "2027 年到期零息可轉債（2026-07-29 發行，4.5 億美元）", "RMB bn", None, "=450*cny/1000", None, "Derived", CBN + "：本金 4.5 億美元、零息、2027-07-27 到期、換股價 5.55 港元", "發行款入現金、本金列負債，兩者相抵")
inp("cari_cash", "支付 CARIAD 現金（可轉換貸款修訂協議）", "RMB bn", None, "=398.9*cny/1000", None, "Derived", CBN + "：可轉債所得約 3.989 億美元用於支付 CARIAD", CARI + "：交割 2026-08-03，現金金額未於公告列示（以募資用途代）")
inp("debt_v", "借款（2026-06-30；流動 790,910＋非流動 571,002）", "RMB bn", None, "=0.79091+0.571002", None, "Derived", IR26, "")
inp("lease_v", "租賃負債（2026-06-30）", "RMB bn", None, 0.167623, None, "Verified", IR26, "")
inp("fvtpl_v", "以公允價值計量的金融資產（非流動，2026-06-30）", "RMB bn", None, 4.111145, None, "Verified", IR26 + "：4,111,145（2025-12 3,502,634）", "")
inp("inv_v", "權益法投資（2026-06-30；含 D-Robotics、Carizon）", "RMB bn", None, 1.636661, None, "Verified", IR26, "")
inp("inv_k", "金融資產與權益法投資計入比例", "x", 0.5, 1.0, None, "Assumed", "基準按帳面全額；低案 50%（D-Robotics 優先股估值、Carizon 虧損）", "")
inp("stub_frac", "Stub 期（2026-06-30 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,6,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac（約下半年）")
inp("rf", "無風險利率（中國 10 年期國債）", "%", None, 0.0161, None, "Verified", CNRF, "現金流為人民幣")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "中國國家風險溢酬", "%", None, 0.0091, None, "Interested-party", DAM, "Damodaran（研究者自編表）", "研究者（立場：中立學者，方法自訂）")
inp("beta", "β（兩來源 Blume 調整後中位數）", "x", "=E{r}-0.3", "=MEDIAN(0.67*2.35+0.33,0.67*MEDIAN(1.24,2.38,1.47)+0.33)", "=E{r}+0.3", "Derived",
    SA_HZ + "：地平線 β 2.35（上市約 2 年，樣本短）；同業 5 年 β：Mobileye 1.24、Ambarella 2.38、禾賽 1.47（stockanalysis）", "Blume 0.67β＋0.33；速騰 0.14 為異常值不列；敏感度 ±0.3（Assumed）")
inp("kd", "稅前負債成本", "%", None, "=rf+0.01", None, "Assumed", "rf＋1pt（借款與零息可轉債為主）", "負債占資本小")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "人民幣 rf 1.61% 使 g 受限")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 EV/營收 Comparable）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺地平線對同業 10 年 EV/營收比值歷史，上市僅約 2 年）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業 EV／未來 12 個月營收（以 2026-11-20 收盤重算市值；淨現金、股數、前瞻營收取資料商頁面；Derived）")
PEERS = [("ev_mbly", "Mobileye MBLY（ADAS 晶片＋方案）", "=(850.19*4.57-1360)/(5660/2.86)", SA_MB + "：股數 850.19M、淨現金 1.36B、市值 5.66B 時 Forward PS 2.86 → NTM 營收 1,979M；2026-11-20 收盤 4.57（" + CSV_R + "）"),
         ("ev_amba", "Ambarella AMBA（視覺 AI 晶片）", "=(44.14*53.07-259.69)/(2940/6.33)", SA_AM + "：股數 44.14M、淨現金 259.69M、市值 2.94B 時 Forward PS 6.33；2026-11-20 收盤 53.07"),
         ("ev_hsai", "禾賽 HSAI（光達）", "=(156.44*11.83-897.23)/(1850/2.44)", SA_HS + "：股數 156.44M、淨現金 897.23M、Forward PS 2.44；2026-11-20 收盤 11.83"),
         ("ev_rs", "速騰聚創 2498.HK（光達；TTM 營收，資料商無前瞻）", "=(471.24*13.32-2300)/2510", SA_RS + "：股數 471.24M、淨現金 2.30B、TTM 營收 2.51B（港元）；2026-11-20 收盤 13.32；分母為 TTM，非 NTM")]
for k, lab, f, src in PEERS:
    inp(k, f"EV／營收：{lab}", "x", None, f, None, "Derived", src, "會計口徑：MBLY、AMBA 美國 GAAP；禾賽、速騰 IFRS（中國公司）；EV 含租賃與否依資料商（未逐家核對）", "")
inp("ev_med", "同業 EV／營收中位數（4 家；少於 8 家不做迴歸）", "x", "=MIN(ev_mbly,ev_amba,ev_hsai,ev_rs)", "=MEDIAN(ev_mbly,ev_amba,ev_hsai,ev_rs)", "=MAX(ev_mbly,ev_amba,ev_hsai,ev_rs)", "Derived", "黑芝麻智能 2533.HK 不在收盤檔（報價日未取得），未列", "同業以硬體為主，毛利率低於地平線（授權占營收過半），倍數可能偏低：不調整（rel_k＝1.00）；敏感度低＝4 家最低、高＝4 家最高（Ambarella，唯一無晶圓廠視覺晶片同業）")
inp("cons_tp", "賣方平均目標價（只作對照，不進模型）", "HKD", None, 8.14, None, "Interested-party", SA_HZ + "：18 位分析師平均 8.14、共識 Strong Buy（2026-10-02 頁）", "個別：高盛 13.12（09-09）、華興 5.80（09-25）、國泰海通 5.18（07-03）、滙豐 10.7（07-28）", BRK)
inp("cons_rev26", "國泰海通 FY2026 營收預測（只作對照，不進模型）", "RMB bn", None, 6.2, None, "Interested-party", GTJ + "：2026／2027／2028 營收 62／88／127 億元", "2026-07 財報前，只作歷史對照；滙豐 07-28 將 2026–28 營收預測下修 22%／26%／30%（" + HSB + "）", BRK)


def build(path, ov=None, sens=None, rev=None, bid="未建置"):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 RMB bn、股數百萬股、EPS 人民幣、股價港元"
    IN["A2"] = "標記：Verified（附出處）／Interested-party（寫明是誰、誘因欄）／Analogy（寫明類比對象）／Assumed（給區間）／Derived（寫明換算）；查不到寫「未取得」。線層級變數見「對帳」頁"
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
    in_last = r - 1
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("s_city30", "城區 NOA 份額 2030"), ("d_soth", "其他 ADAS 份額每年變化"), ("mix", "城區 NOA 每套營收組合上移"), ("g_ls27", "其他授權及服務年增 FY2027"),
           ("g_ls28", "其他授權及服務年增 FY2028–30"), ("g_lslt", "其他授權及服務年增 FY2031–32"), ("gm_p", "產品方案毛利率 FY2028 起")]
    for f, lab in FAM:
        IN.cell(r, 1, f + "_scn"); IN.cell(r, 2, lab)
        for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR[f + '_' + s]}")
        IN.cell(r, 8, f"引用 In 頁 {f}_A／{f}_B／{f}_C；{lab}"); r += 1
    IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重")
    for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR['w_' + s]}")
    IN.cell(r, 8, "引用 In 頁 w_A／w_B／w_C；合計＝1"); r += 1
    IN.column_dimensions["B"].width = 52; IN.column_dimensions["H"].width = 60

    # ---------------- Hist ----------------
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（年度；RMB bn；FY2023 取資料商、FY2024–FY2025 取公司年度業績公告，皆為當期申報口徑，含非車載）。公式只引用本頁（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2023", "FY2024", "FY2025", "標記", "來源"])
    HR = [("h_rev", "營收", 1.552, 2.383554, 3.758268, "Verified", "FY2024–25：" + AR25 + "；FY2023：" + SA_IS + "（資料商）"),
          ("h_p", "產品方案營收", "未取得", 0.664237, 1.622274, "Verified", AR25 + "（FY2023 分項見註）"),
          ("h_l", "授權及服務營收", "未取得", 1.647466, 1.934913, "Verified", AR25),
          ("h_na", "非車載營收", "未取得", 0.071851, 0.201081, "Verified", AR25 + "；2026 起隨 D-Robotics 列停業單位"),
          ("h_cogs", "營業成本", 0.4573, 0.5422, 1.332588, "Verified", AR25 + "；FY2023 資料商"),
          ("h_rnd", "研發費用", 2.366, 3.156055, 5.153708, "Verified", AR25 + "；FY2023 資料商"),
          ("h_sga", "銷售及管理費用", 0.77062, "=0.409853+0.637615", "=0.631924+0.725964", "Verified", AR25 + "（銷售＋管理原文相加）；FY2023 資料商 SG&A"),
          ("h_imp", "金融資產減值", "未取得", 0.051249, 0.009575, "Verified", AR25 + "（FY2023 見註，資料商併入其他）"),
          ("h_op", "營業利益（IFRS，含其他收入及利得）", -2.0, -2.14424, -3.338791, "Verified", AR25 + "；FY2023 為資料商營業利益（口徑可能不含其他收入，未核對）"),
          ("h_pbt", "稅前利益", -6.744, 2.351015, -10.466047, "Verified", AR25 + "；FY2023 資料商"),
          ("h_tax", "所得稅", -0.00508, 0.004507, 0.003319, "Verified", AR25 + "；FY2023 資料商（負值＝所得稅利益）"),
          ("h_ni", "歸屬母公司淨利", -6.739, 2.34658, -10.469012, "Verified", AR25 + "；FY2023 資料商"),
          ("h_eps", "基本 EPS（人民幣）", -2.50, 0.51, -0.81, "Verified", AR25 + "；FY2023 資料商；FY2024 稀釋 −0.17"),
          ("h_adj", "經調整淨虧損（Non-IFRS）", "未取得", -1.681155, -2.811776, "Verified", AR25),
          ("h_sbc", "股份支付", "未取得", 0.562923, 0.992844, "Verified", AR25),
          ("h_da", "折舊攤銷", 0.23672, 0.31222, 0.2862, "Interested-party", SA_IS + "（資料商；法定原文見註）"),
          ("h_units", "征程系列出貨（百萬套）", "未取得", "=4.01/1.388", 4.01, "Verified", AR25 + "：2025 年 401 萬套（+38.8%）；FY2024 由增速推得（Derived）")]
    HROW = {}
    for i, (k, lab, a, b, c_, tag, src) in enumerate(HR, 3):
        H.cell(i, 1, lab)
        for j, v in enumerate((a, b, c_)):
            cc = H.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): cc.font = BLUE
        H.cell(i, 5, tag); H.cell(i, 6, src)
        for j, y in enumerate((23, 24, 25)): dn(f"{k}_{y}", f"Hist!${'BCD'[j]}${i}")
        HROW[k] = i
    i = 3 + len(HR)
    H.cell(i, 1, "檢查：分部合計 − 營收（FY2024、FY2025；應為 0）"); H.cell(i, 3, "=ROUND(h_p_24+h_l_24+h_na_24-h_rev_24,6)"); H.cell(i, 4, "=ROUND(h_p_25+h_l_25+h_na_25-h_rev_25,6)")
    H.cell(i + 1, 1, "毛利（Derived）"); H.cell(i + 1, 2, "=h_rev_23-h_cogs_23"); H.cell(i + 1, 3, "=h_rev_24-h_cogs_24"); H.cell(i + 1, 4, "=h_rev_25-h_cogs_25")
    H.cell(i + 2, 1, "註：FY2023 產品方案／授權／非車載拆分、減值、經調整淨虧損未取得（本輪未讀 2023 年報）；折舊攤銷法定原文未取得（只有資料商）；FY2024 淨利為正來自優先股公允價值利得 4.68（非現金）")
    H.column_dimensions["A"].width = 44; H.column_dimensions["F"].width = 80

    # ---------------- 半年（營收線對帳） ----------------
    Q = wb.create_sheet("半年")
    Q["A1"] = "半年 — 營收線（官方揭露：產品方案、授權及服務、非車載；RMB bn）。H1'25 為 2026 中報重編的持續經營口徑；H2'25 ＝ FY2025 − H1'25（Derived，含非車載）"
    hdr(Q, 2, ["項目", "H1'25（重編）", "H2'25（Derived）", "FY2025", "H1'26", "說明"])
    QR = [("產品方案營收", 0.806574, "=D3-B3", "=h_p_25", "=h1_p"),
          ("授權及服務營收", 0.739493, "=D4-B4", "=h_l_25", "=h1_l"),
          ("非車載營收", 0, "=D5-B5", "=h_na_25", 0),
          ("合併營收", 1.546067, "=D6-B6", "=h_rev_25", "=h1_p+h1_l"),
          ("營業成本", 0.525643, "=D7-B7", "=h_cogs_25", "=h1_cogs")]
    for i, (lab, a, b, c_, d) in enumerate(QR, 3):
        Q.cell(i, 1, lab)
        for j, v in enumerate((a, b, c_, d)):
            cc = Q.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): cc.font = BLUE
    Q.cell(3, 6, IR26 + "（H1 兩欄）；" + AR25)
    Q.cell(5, 6, "H1'25 重編口徑不含非車載（D-Robotics 列停業單位）；原公告 H1'25 含非車載")
    Q.cell(8, 1, "檢查：分部合計 − 合併營收（應為 0）")
    for j, L in enumerate("BCDE"): Q.cell(8, 2 + j, f"=ROUND({L}3+{L}4+{L}5-{L}6,6)")
    Q.cell(9, 1, "產品方案年增（H1'26 vs H1'25）"); Q.cell(9, 5, "=E3/B3-1")
    Q.cell(10, 1, "授權及服務年增（H1'26 vs H1'25）"); Q.cell(10, 5, "=E4/B4-1")
    Q.cell(11, 1, "合併營收年增（H1'26 vs H1'25）"); Q.cell(11, 5, "=E6/B6-1")
    Q.cell(12, 1, "毛利率"); [Q.cell(12, 2 + j, f"=1-{L}7/{L}6") for j, L in enumerate("BCDE")]
    Q.cell(13, 1, "征程出貨（百萬套）"); Q.cell(13, 4, "=h_units_25"); Q.cell(13, 5, "=h1_units")
    Q.cell(14, 1, "每套產品方案營收（RMB／套，Derived）"); Q.cell(14, 4, "=D3/D13*1000"); Q.cell(14, 5, "=E3/E13*1000")
    Q.column_dimensions["A"].width = 40; Q.column_dimensions["F"].width = 70

    # ---------------- 對帳（線層級變數：值與來源格） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = f"對帳 — 本模型讀取的線層級變數（{LINE_R}；值為建置時讀出的常數，來源格見 F 欄）。情境 A＝線模型「高」、B＝「基準」、C＝「低」。與線模型不一致的自設值列在下方"
    hdr(RE, 2, ["鍵", "變數", "情境", "值", "單位", "線模型來源格", "標記", "用在哪"])
    LR = {}
    for i, (k, lab, s, v, u, src, tag, use) in enumerate(LINE, 3):
        RE.cell(i, 1, k); RE.cell(i, 2, lab); RE.cell(i, 3, s); c = RE.cell(i, 4, v); c.font = BLUE
        RE.cell(i, 5, u); RE.cell(i, 6, src); RE.cell(i, 7, tag); RE.cell(i, 8, use)
        dn("ln_" + k, f"對帳!$D${i}"); LR[k] = i
    i = 3 + len(LINE) + 1
    RE.cell(i, 1, "本模型自設、與線模型口徑不同或線模型未列的值").font = BOLD; i += 1
    hdr(RE, i, ["項目", "本模型", "線模型", "差異與理由"]); i += 1
    diffs = [("地平線城區 NOA 份額", "In s_city26 22.8%（2026）→ s_city30_A／B／C", "線模型未建份額（追蹤指標!D11 列 22.82%）", "份額屬公司層，非線層級；線模型只給市場"),
             ("每套產品方案營收", "城區 NOA 每套由產品方案總額校準（PL「城區 NOA 每套產品方案營收」）", "線模型城市 NOA 每車內容 780 美元（下游取數!D25，含感測、運算、授權）", "地平線只拿運算（晶片＋參考設計）一部分；Recon 列占比"),
             ("城市 NOA 授權費", "權利金＝地平線城區 NOA 套數 × 線模型每車授權費（2026 用基準 80 美元，線性到 2030 情境值）", "輸入!D26:F26 40／80／150", "線模型此列本身以地平線授權毛利率校準（Assumed），非獨立證據"),
             ("2026 年市場量", "三情境 2026 年皆用線模型基準（滲透率 21.4%、零售 1,830 萬）", "線模型 2026 年低／高為 18.8%／24.0%（口徑差，非情境）", "2026 年已過 10 個月；分化從 2027 開始"),
             ("2029、2031 年滲透率", "以 2028／2030、2030／2032 線性內插", "下游取數未列 2029、2031", "線模型下游取數只給 2026、27、28、30、32")]
    for row_ in diffs:
        for j, v in enumerate(row_): RE.cell(i, 1 + j, v)
        i += 1
    RE.column_dimensions["B"].width = 40; RE.column_dimensions["F"].width = 60; RE.column_dimensions["H"].width = 40

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（征程出貨 × 每套營收；城區 NOA 權利金）→ 營收線 → 標準損益表（IFRS）→ 現金流（RMB bn；FY2023–FY2025 實績（含非車載）、FY2026＝H1 實績＋H2 假設、FY2027 起模型）"
    PR = {}; rr = 3
    keys = ["retail", "pen", "cityc", "scity", "hzcity", "othc", "soth", "hzoth", "units", "pcity", "poth",
            "prod", "lic_u", "roy", "ols", "lic", "na", "rev", "chk", "gr",
            "cogs", "gp", "gm", "gmp", "rnd", "sga", "imp", "opex", "oi", "op", "opm",
            "fin", "assoc", "fv", "nonop", "pbt", "tax", "disc", "ni", "nim", "nig", "sh", "eps",
            "da", "cx", "dwc", "fcff", "div", "fcfad", "nc"]
    labs = {"retail": "中國乘用車國內零售（百萬台；線模型）", "pen": "中國城市 NOA 滲透率（線模型）", "cityc": "中國城市 NOA 新車（百萬台）",
            "scity": "地平線城區 NOA 份額", "hzcity": "地平線城區 NOA 套數（百萬套）", "othc": "其他新車（非城區 NOA，百萬台）", "soth": "地平線其他 ADAS 份額",
            "hzoth": "地平線其他 ADAS 套數（百萬套）", "units": "征程出貨合計（百萬套）", "pcity": "城區 NOA 每套產品方案營收（RMB／套）",
            "poth": "其他 ADAS 每套產品方案營收（RMB／套）", "prod": "產品方案營收", "lic_u": "城市 NOA 每車授權費（USD；線模型）",
            "roy": "授權－城區 NOA 權利金（子列）", "ols": "授權－其他授權及服務（子列）", "lic": "授權及服務營收", "na": "非車載營收",
            "rev": "營收（合併）", "chk": "檢查：分部合計 − 合併營收（應為 0）", "gr": "營收年增率",
            "cogs": "營業成本", "gp": "毛利", "gm": "毛利率", "gmp": "產品方案毛利率", "rnd": "研發費用", "sga": "銷售及管理費用", "imp": "金融資產減值",
            "opex": "營業費用（研發＋銷售及管理＋減值）", "oi": "其他收入及利得", "op": "營業利益（IFRS，含其他收入及利得）", "opm": "營業利益率",
            "fin": "財務收益淨額", "assoc": "權益法投資損益", "fv": "優先股及金融負債公允價值變動", "nonop": "業外合計（其他收入淨額：財務＋權益法＋公允價值）",
            "pbt": "稅前利益", "tax": "所得稅", "disc": "停業單位損益（D-Robotics）", "ni": "稅後淨利（歸屬母公司）", "nim": "淨利率", "nig": "淨利年增率",
            "sh": "加權平均股數（百萬股）", "eps": "EPS（基本，人民幣）",
            "da": "折舊攤銷", "cx": "資本支出", "dwc": "營運資金增加", "fcff": "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "div": "股利",
            "fcfad": "自由現金流扣股利（收支關卡用）", "nc": "期末淨現金（現金 − 借款 − 可轉債 − 租賃；不含金融資產）"}
    HK = {"rev": "h_rev", "prod": "h_p", "lic": "h_l", "na": "h_na", "cogs": "h_cogs", "rnd": "h_rnd", "sga": "h_sga", "imp": "h_imp", "op": "h_op",
          "pbt": "h_pbt", "tax": "h_tax", "ni": "h_ni", "eps": "h_eps", "da": "h_da", "units": "h_units"}
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, k in enumerate(keys)}
        for k in keys: P.cell(R[k], 1, labs[k])
        L = lambda k: f"{s}"   # noqa
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1)
            cell = lambda k: f"{c}{R[k]}"; prev = lambda k: f"{p}{R[k]}"
            def w(k, f): P[cell(k)] = f
            n = y - 2026
            if y <= 2025:
                yy = str(y)[2:]
                for k, hk in HK.items():
                    w(k, f"={hk}_{yy}")
                w("chk", f"=IF(ISNUMBER({cell('prod')}),ROUND({cell('prod')}+{cell('lic')}+{cell('na')}-{cell('rev')},6),0)")
                if y > 2023: w("gr", f"={cell('rev')}/{prev('rev')}-1")
                w("gp", f"={cell('rev')}-{cell('cogs')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
                w("opex", f"={cell('rnd')}+{cell('sga')}+N({cell('imp')})")
                w("oi", f"={cell('op')}-({cell('gp')}-{cell('opex')})")
                w("opm", f"={cell('op')}/{cell('rev')}")
                w("nonop", f"={cell('pbt')}-{cell('op')}")
                w("disc", 0)
                w("nim", f"={cell('ni')}/{cell('rev')}")
                if y > 2023: w("nig", f'=IF({prev("ni")}>0,{cell("ni")}/{prev("ni")}-1,"N/M（前期虧損）")')
                w("sh", f"={cell('ni')}/{cell('eps')}*1000")
                if y == 2025:
                    w("nc", "=20.188070+0.636922-0.527998-0.117328")
                continue
            # ---- 市場與份額（線層級） ----
            if y == 2026:
                w("retail", "=ln_cars26/ln_pen2026_B/100"); w("pen", "=ln_pen2026_B")
            else:
                w("retail", f"={prev('retail')}*(1+ln_cng_{s})")
                if y in PEN_ROW: w("pen", f"=ln_pen{y}_{s}")
                else: w("pen", f"=(ln_pen{y-1}_{s}+ln_pen{y+1}_{s})/2")
            w("cityc", f"={cell('retail')}*{cell('pen')}"); w("othc", f"={cell('retail')}-{cell('cityc')}")
            if y == 2026:
                w("scity", "=s_city26"); w("soth", "=(h1_units-s_city26*h1_cityn)/(h1_retail-h1_cityn)")
            else:
                w("scity", f"=s_city26+(s_city30_{s}-s_city26)*MIN({n},4)/4")
                w("soth", f"=MAX(0,{prev('soth')}+d_soth_{s})")
            w("hzcity", f"={cell('cityc')}*{cell('scity')}")
            w("hzoth", f"={cell('othc')}*{cell('soth')}" + ("*(1-byd_k)" if y >= 2028 else ""))
            w("units", f"={cell('hzcity')}+{cell('hzoth')}")
            # ---- 每套營收與產品方案 ----
            if y == 2026:
                w("prod", "=h1_p*(1+k2_p)"); w("poth", "=p_oth")
                w("pcity", f"=({cell('prod')}-{cell('hzoth')}*{cell('poth')}/1000)/{cell('hzcity')}*1000")
            else:
                w("poth", f"={prev('poth')}*(1-ln_edec_{s})")
                w("pcity", f"={prev('pcity')}*(1+mix_{s})*(1-ln_edec_{s})")
                w("prod", f"=({cell('hzcity')}*{cell('pcity')}+{cell('hzoth')}*{cell('poth')})/1000")
            # ---- 授權 ----
            w("lic_u", "=ln_lic_B" if y == 2026 else f"=ln_lic_B+(ln_lic_{s}-ln_lic_B)*MIN({n},4)/4")
            w("roy", f"={cell('hzcity')}*{cell('lic_u')}*cny/1000")
            if y == 2026:
                w("lic", "=h1_l*(1+k2_l)"); w("ols", f"={cell('lic')}-{cell('roy')}")
            else:
                g = f"g_ls27_{s}" if y == 2027 else (f"g_ls28_{s}" if y <= 2030 else f"g_lslt_{s}")
                w("ols", f"={prev('ols')}*(1+{g})"); w("lic", f"={cell('roy')}+{cell('ols')}")
            w("na", 0)
            w("rev", f"={cell('prod')}+{cell('lic')}+{cell('na')}")
            w("chk", f"=ROUND({cell('prod')}+{cell('lic')}+{cell('na')}-{cell('rev')},6)+ROUND({cell('roy')}+{cell('ols')}-{cell('lic')},6)")
            w("gr", f"={cell('rev')}/cont25-1" if y == 2026 else f"={cell('rev')}/{prev('rev')}-1")
            # ---- 毛利 ----
            w("gmp", "=gm_p_h1" if y == 2026 else f"=gm_p_h1+(gm_p_{s}-gm_p_h1)*MIN({n},2)/2")
            w("cogs", f"={cell('prod')}*(1-{cell('gmp')})+{cell('lic')}*(1-gm_l)")
            w("gp", f"={cell('rev')}-{cell('cogs')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
            # ---- 費用 ----
            if y == 2026:
                w("rnd", "=h1_rnd*(1+k2_x)"); w("sga", "=h1_sga*(1+k2_x)"); w("imp", "=h1_imp"); w("oi", "=h1_oi+oi/2")
            else:
                w("rnd", f"={prev('rnd')}*(1+{'g_rnd27' if y == 2027 else 'g_rndlt'})"); w("sga", f"={prev('sga')}*(1+g_sga)"); w("imp", 0); w("oi", "=oi")
            w("opex", f"={cell('rnd')}+{cell('sga')}+{cell('imp')}")
            w("op", f"={cell('gp')}-{cell('opex')}+{cell('oi')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            # ---- 業外 ----
            if y == 2026:
                w("fin", f"=h1_fin+y_cash*{prev('nc')}/2"); w("assoc", "=h1_assoc+assoc27/2"); w("fv", "=h1_fv")
                w("tax", "=h1_tax"); w("disc", "=h1_disc"); w("sh", "=(h1_wsh+sh_out)/2")
            else:
                w("fin", f"=y_cash*{prev('nc')}"); w("assoc", "=assoc27" if y == 2027 else ("=assoc27/2" if y == 2028 else 0)); w("fv", 0)
                w("tax", f"=MAX(0,{cell('pbt')})*tax_w"); w("disc", 0); w("sh", "=sh_v")
            w("nonop", f"={cell('fin')}+{cell('assoc')}+{cell('fv')}")
            w("pbt", f"={cell('op')}+{cell('nonop')}")
            w("ni", f"={cell('pbt')}-{cell('tax')}+{cell('disc')}")
            w("nim", f"={cell('ni')}/{cell('rev')}")
            w("nig", f'=IF({prev("ni")}>0,{cell("ni")}/{prev("ni")}-1,"N/M（前期虧損）")')
            w("eps", f"={cell('ni')}/{cell('sh')}*1000")
            # ---- 現金流 ----
            w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}")
            w("dwc", f"=fc_wc*({cell('rev')}-cont25)" if y == 2026 else f"=fc_wc*({cell('rev')}-{prev('rev')})")
            w("fcff", f"={cell('op')}-MAX(0,{cell('op')})*tax_w+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("div", f"=pay_r*MAX(0,{cell('ni')})")
            w("fcfad", f"={cell('fcff')}-{cell('div')}")
            if y == 2026:
                w("nc", f"=cash_v+cb_v-cari_cash-debt_v-cb_v-lease_v+{cell('fcff')}*stub_frac")
            else:
                w("nc", f"={prev('nc')}+{cell('fcfad')}+{cell('fin')}*(1-tax_w)")
        P.cell(R["retail"], 2, "FY2026＝線模型 2026 城市 NOA 新車 ÷ 滲透率；之後 ×(1＋線模型銷量年增)")
        P.cell(R["soth"], 2, "FY2026＝(H1 出貨 − 份額 × H1 城區 NOA 新車) ÷ (H1 零售 − H1 城區 NOA 新車)")
        P.cell(R["pcity"], 2, "FY2026 以產品方案總額校準；之後 ×(1＋組合上移)(1 − 線模型降價率)")
        P.cell(R["prod"], 2, "FY2026＝H1 實績 ×(1＋H2÷H1)；之後＝套數 × 每套營收")
        P.cell(R["lic"], 2, "FY2026＝H1 實績 ×(1＋H2÷H1)；之後＝權利金＋其他授權及服務")
        P.cell(R["gr"], 2, "FY2026 以持續經營 FY2025（不含非車載）為基期")
        P.cell(R["op"], 2, "公司 IFRS 營業虧損口徑（含其他收入及利得）")
        P.cell(R["ni"], 2, "FY2026 含 H1 公允價值利得 5.24 與停業單位利得 2.17（一次性、非現金）")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 − MAX(0, 營業利益) × 稅率；股份支付不加回")
        P.cell(R["nc"], 2, "FY2025 實績；FY2026＝2026-06-30 擬制（加可轉債、減 CARIAD 現金）＋下半年 FCFF")
        for k in ("pen", "scity", "soth", "gr", "gm", "gmp", "opm", "nim", "nig"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        for k in ("pcity", "poth"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "#,##0"
        PR[s] = R; rr = b + len(keys) + 2
    P.column_dimensions["A"].width = 48; P.column_dimensions["B"].width = 34
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "prod", "lic", "fcfad", "units"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)。RMB bn；每股換算港元"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（RMB bn）"; D["B4"] = "=px*fx*sh_v/1000"
    D["A5"] = "負債 D（借款＋可轉債＋租賃）"; D["B5"] = "=debt_v+cb_v+lease_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(上限, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金與投資（2026-06-30 擬制）＝現金 ＋ 可轉債所得 − 支付 CARIAD − 借款 − 可轉債 − 租賃 ＋（金融資產＋權益法投資）× inv_k"
    D["B10"] = "=cash_v+cb_v-cari_cash-debt_v-cb_v-lease_v+inv_k*(fvtpl_v+inv_v)"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"
    D["B11"] = "=ROUND(((px*fx*sh_v/1000)*(rf+beta*(erp+crp))+(debt_v+cb_v+lease_v)*kd*(1-tax_w))/((px*fx*sh_v/1000)+debt_v+cb_v+lease_v)+wacc_adj-wacc,9)"
    for rr_ in range(3, 11): D[f"B{rr_}"].number_format = "0.00%" if rr_ in (3, 6, 7, 8, 9) else "#,##0.00"
    DR = {}; rr = 13
    for s in SC:
        R = PR[s]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}" for y in range(2027, 2033)]); b = rr + 2
        rows = ["EBIT（營業利益）", "NOPAT", "＋折舊攤銷", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF（應為 0）", "折現年數（期中）", "折現因子", "FCFF 現值"]
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
                 ("＋淨現金與投資（2026-06-30 擬制）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!F{R['fcff']}*stub_frac"),
                 ("股權價值（RMB bn）", f"=B{t+4}+B{t+5}+B{t+6}"),
                 ("每股價值（評價日，港元）", f"=MAX(0,B{t+7})/sh_v*1000/fx"),
                 ("DCF 情境價值推到 12 個月後（×(1＋ke)^t，港元；無股利）", f"=B{t+8}*(1+ke)^t_tgt"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,N(B{t+10})<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("FY2032 營業利益率", f"=PL!L{R['opm']}"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後，港元）", f"=MAX(0,(B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6)+B{t+5}+B{t+6})/sh_v*1000/fx*(1+ke)^t_tgt")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 9, ps=t + 8, ev=t + 4, tvs=t + 10, alt=t + 13)
        dn(f"dcf_{s}", f"DCF!$B${t+9}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 64

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價（港元）：DCF（權重見 In 頁，基準 50%）＋ EV／營收 Comparable（其餘權重）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}"); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "DCF 頁")
    comp_f = [f"=MAX(0,ev_med*PL!H{PR[s]['rev']}+PL!G{PR[s]['nc']}+inv_k*(fvtpl_v+inv_v))/sh_v*1000/fx" for s in SC]
    r_cu = vrow("Comparable（未調整）：同業 EV／營收中位數 × FY2028 營收 ＋ FY2027 年底淨現金與投資", comp_f,
                "目標時點（2027-11）的未來 12 個月約為 FY2028；淨現金取 FY2027 年底（含 2027 年燒錢）")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（DCF 權重 × DCF ＋（1 − DCF 權重）× Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（港元）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權，港元）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權；EV／營收，港元）", f"=E{r_ca}")
    r_px = one("現價（In 頁 px，港元）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（RMB bn；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級（港元）", f'=TEXT(px*(1+IF(B{r_up}>0.5,9,IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5))))),"0.00")&"／"&IF(B{r_up}<-0.5,"無（已最低級）",TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,-0.5)))),"0.00"))', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    r_minc = one("期間流動性：情境 B 最低年末淨現金（FY2026–FY2030，RMB bn）", f"=MIN(PL!F{PR['B']['nc']}:PL!J{PR['B']['nc']})", "原則第 10 節第 55 條；不含金融資產與權益法投資")
    def _bef(s_):
        _opr = PR[s_]["op"]; _be = '"FY2032 前未轉正"'
        for y in range(2032, 2025, -1):
            _be = f'IF(PL!{COL[y]}{_opr}>0,"FY{y}",{_be})'
        return "=" + _be
    r_be = one("情境 B 營業利益轉正的第一個年度", _bef("B"))
    r_beA = one("情境 A 營業利益轉正的第一個年度", _bef("A"))
    r_beC = one("情境 C 營業利益轉正的第一個年度", _bef("C"))
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_evs = one("現行 EV／FY2027 營收（情境 B）：(市值 − 淨現金與投資) ÷ 營收", f"=(px*fx*sh_v/1000-dcf_nc)/PL!G{PR['B']['rev']}", "分母：模型情境 B FY2027 營收（約等於未來 12 個月）")
    r_evr = one("現行 EV／FY2027 營收 ÷ 同業中位數", f"=B{r_evs}/ev_med", "只作對照")
    r_ievs = one("隱含 EV／FY2028 營收（加權目標價、情境 B）", f"=(B{r_tp}*fx*sh_v/1000-dcf_nc)/PL!H{PR['B']['rev']}", "")
    for rr_ in (r_up, r_a1, r_a2, r_a3): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "0.00"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "0.00"
    V.column_dimensions["A"].width = 72; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, pd2=r_pd2, gj=r_gj, beA=r_beA, beC=r_beC, wk=r_wk, evs=r_evs, evr=r_evr, ievs=r_ievs, minc=r_minc, be=r_be)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司揭露、賣方與線模型對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或外部", "差", "說明"])
    B_ = PR["B"]
    rc = [("FY2026 營收（RMB bn）", f"=PL!F{B_['rev']}", "=cons_rev26", "國泰海通 2026-07 預測（財報前，只作歷史對照）；滙豐 07-28 已下修 22%"),
          ("FY2026 征程出貨（百萬套）", f"=PL!F{B_['units']}", "=h1_units*2", "對照：H1 實績 × 2（無公司全年指引）"),
          ("FY2026 城區 NOA 每套產品方案營收（RMB）÷ 線模型城市 NOA 每車內容（RMB）", f"=PL!F{B_['pcity']}", "=ln_cont26*cny", "隱含地平線占每車智能內容的比例＝B ÷ C（線模型內容含感測與授權，口徑不同，只看量級）"),
          ("FY2030 中國城市 NOA 新車（百萬台）", f"=PL!J{B_['cityc']}", "=ln_cars30/100", "本模型以 2026 零售 × 線模型銷量年增推得，應與線模型一致"),
          ("FY2026 授權及服務中城區 NOA 權利金占比", f"=PL!F{B_['roy']}/PL!F{B_['lic']}", 0, "權利金以線模型每車授權費推得；公司未揭露拆分（未取得）"),
          ("營業利益轉正年度（情境 B）vs 公司目標（約 2028 年損益兩平）", f"=Val!B{r_be}", "FY2028", "公司目標為 Interested-party（財新 2026-09-02 轉述創辦人；出處見 Sources）；模型輸出，非獨立證據"),
          ("加權目標價 vs 賣方平均目標價（港元）", "=tp", "=cons_tp", "資料商 18 位分析師平均（2026-10-02，Interested-party）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f'=IF(AND(ISNUMBER(B{i}),ISNUMBER(C{i})),B{i}-C{i},"—")'); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，建置程式二分法；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 64; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    sd = sens or {}
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（建置程式以 LibreOffice 重算每一案例）；結果為常數快照（港元）"
    SE["A2"] = "build_id"; SE["D2"] = bid; SE["H2"] = f"build_id {bid}"
    SE["A3"] = "模型狀態"; SE["B3"] = sd.get("fp", 0); SE["C3"] = f"=ROUND(SUM(In!E4:E{in_last}),6)"
    SE["D3"] = '=IF(ROUND(B3,6)=C3,"CURRENT","STALE")'; SE["E3"] = "B＝敏感度重算時的 In 頁基準欄指紋；C＝本檔即時指紋"
    hdr(SE, 5, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    sr = 6
    for cs in sd.get("cases", []):
        SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
        SE.cell(sr, 4, cs["tp"]); SE.cell(sr, 5, cs["tp"] / sd["base"]["tp"] - 1)
        SE.cell(sr, 6, cs["dcf"]); SE.cell(sr, 7, (cs["dcf"] / sd["base"]["dcf"] - 1) if sd["base"]["dcf"] else 0); SE.cell(sr, 8, cs["rt"])
        SE[f"E{sr}"].number_format = "0.0%"; SE[f"G{sr}"].number_format = "0.0%"; SE[f"D{sr}"].number_format = "0.00"; SE[f"F{sr}"].number_format = "0.00"
        sr += 1
    SE.cell(sr + 1, 1, "說明：相對倍數係數 rel_k 基準即為 1.00（未調整），故「相對倍數係數＝1」案例＝基準。無低／高區間的鍵不列案例；情境專屬輸入以情境 B 列做單參數案例，A／C 列即情境區間。")
    SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 34

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("比亞迪等大客戶改用自研晶片", "部分（byd_k 尾端 Sens）", "客戶集中度未取得（年報未讀）；單一客戶可能占出貨過半", "比亞迪璇璣 A3 搭載車型數、地平線半年出貨年增"),
             ("中國乘用車需求下滑（購置稅與補貼退坡）", "是（線模型銷量年增）", "情境 C 讀線模型低情境", "乘聯會月度零售"),
             ("英偉達 Thor、華為、高通在城區 NOA 的價格競爭", "是（s_city30_C、gm_p_C）", "情境 C", "高工城區 NOA 晶片份額半年報"),
             ("美國出口管制或晶圓代工限制（台積電代工）", "否", "無法量化；征程 6 系列由台積電代工（未查核）", "美國商務部規則、代工轉移公告"),
             ("合資與聯營公司（Carizon、D-Robotics）持續虧損需注資", "部分（assoc27、inv_k）", "權益法損益不進 FCFF；注資未建模", "權益法損益、對聯營公司增資公告"),
             ("2027-07 可轉債到期（4.5 億美元）", "是（列負債）", "換股價 5.55 港元高於現價，基準以現金償還", "股價相對換股價"),
             ("股份激勵稀釋", "部分（sh_rsu 暫值）", "RSU／選擇權明細未取得", "年報股份激勵附註")]
    for i, row_ in enumerate(risks, 3):
        for j, v in enumerate(row_): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 60

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([IR26, AR25, CARI, CBN, CNF, CXG, HXS, GTJ, GSR, HSB, TIG, ETN, SA_HZ, SA_IS, SA_MB, SA_AM, SA_HS, SA_RS, YCL, CNRF, DAM,
                            f"線模型：{LINE_R}（下游取數、輸入、追蹤指標）"], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"地平線 Horizon Robotics（9660.HK）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-23（財報 2026-08-31 中期業績；股價 2026-11-20）；成品日 2026-11-23；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.00")&" 港元）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：地平線的股價取決於兩件事——征程晶片在中國城區 NOA 的份額與每套價值能否上移（產品方案），以及高毛利的授權及服務能否持續高成長，讓營業虧損在研發費用仍上升下收斂；營收線＝產品方案（征程套數 × 每套營收）＋授權及服務（城區 NOA 權利金＋其他授權）"
    RD["A4"] = (f'="在城區 NOA 份額 2030 年情境 B "&TEXT(s_city30_B,"0%")&"、其他授權及服務 FY2028–30 年增 "&TEXT(g_ls28_B,"0%")&"、研發費用 FY2028 起年增 "&TEXT(g_rndlt,"0%")&" 的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"0.00")&" 港元（DCF "&TEXT(Val!B{r_dcft},"0.00")&"、EV／營收 Comparable "&TEXT(Val!B{r_cmpt},"0.00")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0.0")&" 十億人民幣）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 港元。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")&"。情境 B 營業利益轉正："&Val!B{r_be}&"；最低年末淨現金 "&TEXT(Val!B{r_minc},"0.0")&" 十億人民幣"'
    RD["A6"] = "失效條件（可觀察）：城區 NOA 晶片份額（高工半年報）連兩期下滑；征程半年出貨年增轉負；授權及服務半年年增低於 15%；半年研發費用年增高於 20% 而營收年增低於 25%"
    RD["A7"] = "範圍：只建持續經營（車載）業務；非車載（D-Robotics）2026-03-31 起列停業單位與權益法投資。車用晶片規格與產量屬 AI半導體 權威、本模型只讀；Robotaxi 試營運（2026Q3 目標）未建模。HTML 展示層見 dist/ 同名 .html"
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
def recalc_many(srcs, outdir):
    os.makedirs(outdir, exist_ok=True)
    for i in range(0, len(srcs), 15):
        for attempt in range(2):
            try:
                subprocess.run(["soffice", f"-env:UserInstallation={LO}", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", outdir] + srcs[i:i + 15],
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
                break
            except subprocess.TimeoutExpired:
                subprocess.run(["pkill", "-f", "lo_horizon_profile"]);
                if attempt: raise
    return [os.path.join(outdir, os.path.basename(s)) for s in srcs]

def read_vals(path, VR, DR=None):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]
    d = dict(tp=V[f"B{VR['tp']}"].value, dcf=V[f"B{VR['dcf']}"].value, cmp=V[f"B{VR['cmp']}"].value, up=V[f"B{VR['up']}"].value, rt=V[f"B{VR['rt']}"].value,
             fp=wb["Sens"]["C3"].value)
    if DR:
        I = wb["In"]; w = {}; wd = None
        for row in I.iter_rows(values_only=True):
            if row[0] in ("w_A", "w_B", "w_C"): w[row[0][-1]] = row[4]
            if row[0] == "w_dcf": wd = row[4]
        alt = sum(w[s] * wb["DCF"][f"B{DR[s]['alt']}"].value for s in SC)
        d["alt_dcf"] = alt; d["alt_tp"] = wd * alt + (1 - wd) * d["cmp"]
    return d

def run_cases(ovs):
    d = tempfile.mkdtemp(prefix="hz_"); paths = []; meta = None
    for i, ov in enumerate(ovs):
        p = os.path.join(d, f"c{i:03d}.xlsx"); meta = build(p, ov); paths.append(p)
    outs = recalc_many(paths, os.path.join(d, "o"))
    VR, _, DR = meta
    res = [read_vals(o, VR, DR) for o in outs]
    shutil.rmtree(d, ignore_errors=True); return res

def in_val(k, col):
    for row in ROWS:
        if row[0] == k: return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)

def fingerprint():
    return hashlib.sha1(json.dumps([ROWS, LINE], ensure_ascii=False, default=str).encode()).hexdigest()[:10]

def rating(u):
    return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"

CASES = [("k2_p", "產品方案 H2÷H1（FY2026）"), ("k2_l", "授權及服務 H2÷H1（FY2026）"), ("k2_x", "營業費用 H2÷H1（FY2026）"),
         ("s_city30_B", "城區 NOA 份額 2030（情境 B）"), ("d_soth_B", "其他 ADAS 份額每年變化（情境 B）"), ("p_oth", "其他 ADAS 每套營收"),
         ("mix_B", "城區 NOA 每套營收組合上移（情境 B）"), ("byd_k", "比亞迪改用自研晶片（尾端）"),
         ("g_ls27_B", "其他授權及服務年增 FY2027（情境 B）"), ("g_ls28_B", "其他授權及服務年增 FY2028–30（情境 B）"), ("g_lslt_B", "其他授權及服務年增 FY2031–32（情境 B）"),
         ("gm_p_B", "產品方案毛利率（情境 B）"), ("gm_l", "授權及服務毛利率"), ("g_rnd27", "研發費用年增 FY2027"), ("g_rndlt", "研發費用年增 FY2028 起"),
         ("g_sga", "銷售及管理費用年增"), ("oi", "其他收入及利得"), ("assoc27", "權益法投資損益 FY2027"), ("y_cash", "淨現金收益率"), ("tax_w", "有效稅率"),
         ("fc_da", "折舊 ÷ 營收"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("sh_rsu", "RSU／選擇權股數"), ("inv_k", "金融資產與投資計入比例"),
         ("ev_med", "同業 EV／營收倍數"), ("beta", "β"), ("erp", "ERP"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重")]
SOLVE = [("s_city30_B", 0.0, 1.0, "情境 B 城區 NOA 份額 2030（其他不變，整本重算；情境 A／C 不動）"),
         ("g_ls28_B", -0.3, 1.2, "情境 B 其他授權及服務年增 FY2028–FY2030（情境 A／C 不動）"),
         ("g_rndlt", -0.4, 0.3, "研發費用年增 FY2028 起（三情境同動）")]

def write_readme(path, VR, PR, DR, bid, sens, revs, checks):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f2 = lambda x: f"{x:,.2f}"; pc = lambda x: f"{x*100:+.1f}%"
    inv = lambda k: I['E' + str([r for r in range(1, 400) if I['A' + str(r)].value == k][0])].value
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    base = sens["base"]["tp"]
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / base - 1))[:8]
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {r['val'] if isinstance(r['val'], str) else format(r['val'], '.4f')} | {r['base']} | {r['note']} |" for r in revs)
    def yr(s, k, fmt="{:,.2f}"):
        return " | ".join(fmt.format(P[f'{COL[y]}{PR[s][k]}'].value) if isinstance(P[f'{COL[y]}{PR[s][k]}'].value, (int, float)) else "—" for y in YRS)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f2(g(chr(66+i)+str(VR['dcfr'])))} | {f2(g(chr(66+i)+str(VR['cu'])))} | {f2(g(chr(66+i)+str(VR['tps'])))} | {g(chr(66+i)+str(VR['gate'])):,.1f} |" for i, s in enumerate(SC))
    rts = sorted(set(c['rt'] for c in sens['cases']))
    chk = "\n".join(f"- `{c}`：{'通過' if rc == 0 else '未通過'}（{out}）" for c, rc, out in checks)
    txt = f"""# physical_ai/horizon — 地平線 Horizon Robotics（9660.HK）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

港交所主板上市（亞太掛牌）。報表人民幣、股價港元（匯率 CNY=X ÷ HKD=X＝{inv('fx'):.5f} 人民幣／港元）。本 README 的數字由 `build_horizon.py` 從重算後的成品讀出產生，不手寫。

**未做與縮小範圍（高不確定性原則）：** 公司仍虧損、無數字指引，FY2027 起的營收、毛利率、費用全部是 Assumed 路徑（給情境與區間）；DCF 的價值幾乎全來自終值期利潤率，Comparable 以硬體為主的同業 EV／營收，兩法都在證據邊界外。未做：客戶集中度（年報未讀）、FY2023 營收拆分、現金流量表科目（中期公告無現金流量表；折舊、資本支出用資料商與假設）、RSU／選擇權股數（暫值 400 百萬股）、未參與製作者的獨立核對。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

## 估值（評價日 2026-12-31；目標時點約 12 個月；港元）
| 情境 | DCF（推到 12 個月後） | EV／營收 Comparable | 情境目標價 | 收支關卡（FY2026–30 FCFF 扣股利，RMB bn） |
|---|---|---|---|---|
{scn}

- 加權目標價 **{f2(tp)} 港元**（DCF {f2(dcf)}、Comparable {f2(cmp_)}），現價 {f2(inv('px'))}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")} → 投資建議 **{g(f"B{VR['rec']}")}**。
- 方法：DCF 標準版 FCFF（WACC／ke／g／RONIC＝{g(f"B{VR['wk']}")}）50%＋EV／營收 Comparable 50%（同業 Mobileye、Ambarella、禾賽、速騰中位數 {inv('ev_med'):.2f} 倍 × FY2028 營收）；相對倍數係數 1.00（未調整）。
- 現行 EV／FY2027 營收（情境 B）{g(f"B{VR['evs']}"):.2f} 倍，為同業中位數的 {g(f"B{VR['evr']}"):.2f} 倍；加權目標價隱含 EV／FY2028 營收 {g(f"B{VR['ievs']}"):.2f} 倍（接近零：情境 B、C 的 DCF 股權價值為負、以零計，加權目標價約等於每股淨現金與投資）。同業以硬體為主（毛利率較低），地平線授權占營收過半，倍數差有部分是商業模式差，不是錯價（判斷題，未調整）。
- 賣方平均目標價 {inv('cons_tp'):.2f} 港元（資料商 18 位、2026-10-02），遠高於本模型：個別賣方以 2030 年獲利倍數折現（高盛 2026-09-09：29 倍 EV／EBITDA，經轉述），公司目標約 2028 年損益兩平（Interested-party）；本模型情境 B 營業利益轉正年度：{g(f"B{VR['be']}")}。

## 營收線與損益（情境 B，RMB bn）
| 列 | {" | ".join(f"FY{y}" for y in YRS)} |
|---|{"---|" * len(YRS)}
| 征程出貨（百萬套） | {yr('B','units')} |
| 產品方案 | {yr('B','prod')} |
| 授權及服務 | {yr('B','lic')} |
| 營收（合併） | {yr('B','rev')} |
| 毛利率 | {yr('B','gm', '{:.1%}')} |
| 研發費用 | {yr('B','rnd')} |
| 營業利益 | {yr('B','op')} |
| 歸屬母公司淨利 | {yr('B','ni')} |
| FCFF | {yr('B','fcff')} |

- FY2023–FY2025 為申報口徑（含非車載）；FY2026＝H1 實績＋H2 假設（淨利含 H1 一次性公允價值利得與停業單位利得）；FY2027 起模型。
- 命題模組：征程套數＝線模型中國城市 NOA 新車 × 地平線份額 ＋ 其他新車 × 其他 ADAS 份額；每套營收：城區 NOA 以 FY2026 產品方案總額校準，之後乘組合上移、除以線模型電子內容降價；授權＝城區 NOA 套數 × 線模型每車授權費＋其他授權及服務（成長路徑）。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 城區 NOA 份額 2030：情境 A／B／C {inv('s_city30_A'):.0%}／{inv('s_city30_B'):.0%}／{inv('s_city30_C'):.0%}（2026 年 22.8%，Interested-party）；其他 ADAS 份額每年 {inv('d_soth_A'):+.0%}／{inv('d_soth_B'):+.0%}／{inv('d_soth_C'):+.0%}。
- 其他授權及服務年增 FY2027 {inv('g_ls27_A'):.0%}／{inv('g_ls27_B'):.0%}／{inv('g_ls27_C'):.0%}、FY2028–30 {inv('g_ls28_A'):.0%}／{inv('g_ls28_B'):.0%}／{inv('g_ls28_C'):.0%}；產品方案毛利率 FY2028 起 {inv('gm_p_A'):.0%}／{inv('gm_p_B'):.0%}／{inv('gm_p_C'):.0%}；授權毛利率 {inv('gm_l'):.1%}（Derived）。
- 研發費用 FY2027 +{inv('g_rnd27'):.0%}、FY2028 起每年 +{inv('g_rndlt'):.0%}（三情境相同）；銷售及管理 +{inv('g_sga'):.0%}。
- 折現：rf {inv('rf'):.2%}（中國 10 年期）、β {inv('beta'):.3f}（Blume 中位數）、ERP 5%、CRP 0.91%、g＝MIN(3%, rf)、RONIC＝WACC；折舊 {inv('fc_da'):.0%}、資本支出 {inv('fc_cx'):.0%}、營運資金 {inv('fc_wc'):.0%} × 營收增量。
- 評價日股數：{inv('sh_out'):,.1f} 百萬股（CARIAD 換股後）＋RSU 暫值 {inv('sh_rsu'):,.0f}；可轉債（換股價 5.55 港元）列負債。淨現金與投資 {wb['DCF']['B10'].value:,.2f} RMB bn（2026-06-30 擬制）。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| 城區 NOA 份額（s_city30） | 高工智能汽車城區 NOA 晶片份額半年報（2026H1 22.8%） |
| 其他授權及服務成長（g_ls27、g_ls28） | 公司半年報授權及服務營收年增（H1'26 +52.7%） |
| 研發費用成長（g_rnd27、g_rndlt） | 半年報研發費用（H1'26 2.75，+21.9%） |
| 產品方案毛利率（gm_p） | 半年報產品方案毛利率（H1'26 36.2%） |

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f2(c['tp'])} | {pc(c['tp']/base-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁；各案例估值評等出現「{'、'.join(rts)}」。

## 反解（整本重算，二分法；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
讀線 2 智慧車模型 `{LINE_R}`：中國城市 NOA 滲透率 2026／27／28／30／32（下游取數 C27:E31，情境 A／B／C＝高／基準／低）、2026 城市 NOA 新車 391.62 萬（D32，推國內零售 1,830 萬）、銷量年增（輸入 D5:F5）、電子內容年降價率（輸入 D31:F31）、城市 NOA 每車授權費 40／80／150 美元（輸入 D26:F26；下游取數未列）。不一致處：2026 年三情境都用線模型基準（線模型 2026 低／高是口徑差）；2029、2031 滲透率內插；份額與每套營收為公司層自設；線模型授權費本身以地平線授權毛利率校準，與本模型的授權營收不是獨立證據。

## 限制與未取得資料
- 客戶集中度（比亞迪占比）未取得；比亞迪 250 萬顆的說法出自程式產生稿，未查核、不作依據，只用於尾端 Sens（byd_k）。
- 中期公告無現金流量表：折舊、資本支出、營業現金流未取得；FY2023 營收拆分、減值、經調整淨虧損未取得（FY2023 用資料商）。
- 支付 CARIAD 的現金以可轉債募資用途（約 3.989 億美元）代替，公告未列金額。
- 同業倍數：MBLY、AMBA、速騰的資料商頁面日期為 2026-10 下旬，市值以 2026-11-20 收盤重算，淨現金與前瞻營收沿用頁面值；速騰分母為 TTM；黑芝麻智能不在收盤檔（報價日未取得）；4 家少於 8 家不做迴歸。
- 賣方共識逐年營收、EPS 未取得（只有國泰海通 2026-07 的 62／88／127 億元，財報前）。

## 核對
{chk}

## 檔案
- `build_horizon.py`：建置＋整本重算 Sens／反解＋產生本 README、`segments.json`、`sens.json`，最後跑核對。
- `dist/{NAME}`：成品（工作表 README、In、Hist、半年、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-23）：初版；現價 {f2(inv('px'))}（2026-11-20 收盤）；加權目標價 {f2(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


def run_checks(x):
    cmds = [["tools/in_layout/check_in_layout.py", x], ["tools/sens_coverage/check_sens_coverage.py", x, "--require-keycol"],
            ["tools/model_governance/check_governance.py", x], ["tools/model_governance/check_rating_gate.py", x, "--strict"],
            ["tools/model_governance/check_zero_rows.py", x], ["tools/model_governance/check_cached_values.py", x],
            ["tools/segment_check/check_segments.py", HERE]]
    html = os.path.join(HERE, "dist", NAME[:-5] + ".html")
    for c in (["physical_ai/horizon/gen_html_json.py", x], ["tools/company_html/build_company_html.py", HERE]):
        subprocess.run([sys.executable] + c, check=True, cwd=ROOT, stdout=subprocess.DEVNULL)
    cmds += [["tools/company_html/check_company_html.py", HERE], ["tools/company_html/check_layout_v13.py", html, "--xlsx", x]]
    out = []
    for c in cmds:
        rp = subprocess.run([sys.executable] + c, capture_output=True, text=True, cwd=ROOT)
        last = (rp.stdout.strip().splitlines() or rp.stderr.strip().splitlines() or [""])
        msg = " ".join(l.strip() for l in last[:2])[:200].replace(ROOT + "/", "")
        print(c[0], rp.returncode, msg)
        out.append((c[0], rp.returncode, msg))
    return out


def main():
    quick = "--quick" in sys.argv
    if quick:
        sens = json.load(open(os.path.join(HERE, "sens.json"))); revs = sens["revs"]
    else:
        base = run_cases([{}])[0]; print("基準", base); px0 = in_val("px", "base")
        beta_b = statistics.median([0.67 * 2.35 + 0.33, 0.67 * statistics.median([1.24, 2.38, 1.47]) + 0.33])
        specs = []
        for k, drv in CASES:
            for side in ("lo", "hi"):
                v = in_val(k, side)
                if k == "beta": v = beta_b + (-0.3 if side == "lo" else 0.3)
                if v is None: continue
                b_ = in_val(k, "base")
                if isinstance(b_, (int, float)) and abs(v - b_) < 1e-12: continue
                lab_v = v.replace("=", "") if isinstance(v, str) else f"{v:.4g}"
                specs.append((dict(drv=drv + ("（低）" if side == "lo" else "（高）"), set=f"{k}＝{lab_v}", key=k), {k: v}))
        for s in SC:
            specs.append((dict(drv=f"只看情境 {s}", set=f"w_{s}＝1，其餘 0", key=f"w_{s}"), {f"w_{x}": (1.0 if x == s else 0.0) for x in SC}))
        res = run_cases([ov for _, ov in specs])
        cases = []
        for (m, _), r in zip(specs, res):
            cases.append(dict(m, tp=r["tp"], dcf=r["dcf"], cmp=r["cmp"], up=r["up"], rt=r["rt"])); print(m["key"], round(r["tp"], 3), r["rt"])
        up_ = base["alt_tp"] / px0 - 1
        cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rating(up_)))
        up0 = base["up"]; tiers = [0.5, 0.25, -0.25, -0.5]
        upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
        revs = []
        for key, lo_, hi_, note in SOLVE:
            b0 = in_val(key, "base")
            for lab, tgt in (("目標價＝現價（上檔 0%）", 0.0), (f"升一級邊界（上檔 {upper:+.0%}）" if upper is not None else None, upper),
                             (f"降一級邊界（上檔 {lower:+.0%}）" if lower is not None else None, lower)):
                if lab is None: continue
                f = lambda x: run_cases([{key: x}])[0]["up"] - tgt
                a, b = lo_, hi_; fa, fb = f(a), f(b)
                if fa * fb > 0:
                    x = f"區間 {lo_:+.0%}～{hi_:+.0%} 內無解"
                else:
                    for _ in range(10):
                        m_ = (a + b) / 2; fm = f(m_)
                        if fa * fm <= 0: b, fb = m_, fm
                        else: a, fa = m_, fm
                    x = (a + b) / 2
                revs.append(dict(label=lab, key=key, val=x, base=b0, note=note)); print("反解", key, lab, x)
        sens = dict(base=base, cases=cases, fp=base["fp"], revs=revs, src_fp=fingerprint())
        json.dump(sens, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
    bid = f"horizon-{VER}-{sens['src_fp']}"
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="hzf_"); raw = os.path.join(tmp, NAME)
    VR, PR, DR = build(raw, {}, sens, revs, bid)
    out = recalc_many([raw], os.path.join(tmp, "o"))[0]
    shutil.copy(out, os.path.join(dist, NAME))
    wbv = openpyxl.load_workbook(os.path.join(dist, NAME), data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10]); assert not errs, errs
    json.dump(dict(company="horizon", ticker="9660.HK", model_sheet="半年",
                   official_segments=[dict(name=n, label=l, period="H1'26", value=v, unit="RMB bn", source=IR26)
                                      for n, l, v in (("產品方案（Product solutions）", "產品方案營收", 0.925627), ("授權及服務（License and services）", "授權及服務營收", 1.129357),
                                                      ("非車載（Non-automotive；2026 起列停業單位）", "非車載營收", 0.0))],
                   sublines=[], check_labels=["檢查：分部合計 − 合併營收（應為 0）"],
                   note="半年頁：H1'25（重編）、H2'25、FY2025、H1'26；PL 頁三情境另有「檢查：分部合計 − 合併營收（應為 0）」（含授權兩子列合計＝授權及服務）。公司以營收線揭露（車載：產品方案、授權及服務；非車載），非 IFRS 8 多分部"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)
    checks = run_checks(os.path.join(dist, NAME))
    write_readme(os.path.join(dist, NAME), VR, PR, DR, bid, sens, revs, checks)

if __name__ == "__main__":
    main()

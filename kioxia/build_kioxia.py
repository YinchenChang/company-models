# Kioxia（285A JP）收支與估值模型 v0.1（AI 半導體 Project：記憶體——NAND）
# 用法：python3 kioxia/build_kioxia.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度 4 月到隔年 3 月：FY2026＝2026-04～2027-03 ≈ CY2026Q2–CY2027Q1。金額：十億日圓（¥B）；每股：日圓（2026-10-01 一股分三股後）。
# 結構比照 sndk/build_sndk.py（同為 NAND、讀同一個市場層）；估值主法比照記憶體樣板（Comparable 用前瞻 P/B）。
# 情境與行業一致：A／B／C 讀 markets/nand 同字母；公司特有變數（每位元成本年降）A 最有利、C 最不利，且與 SanDisk 相同（同一合資晶圓廠與 BiCS 世代）。
import sys, os, glob, json
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as GL
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import valuation_rule as VR
import fcff_dcf as FD   # v0.5（2026-11-05）：標準版 DCF（無槓桿自由現金流 FCFF＋WACC）共用模組（docs/plan/個股DCF標準規格.md）
sys.path.insert(0, os.path.join(ROOT, "siliconomics")); import reader as SIL  # noqa: E402

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC"); FADE = PatternFill("solid", fgColor="E7E6E6")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
SCN = {"A": "順利", "B": "平平", "C": "不如預期"}
def defname(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
rel = lambda f: os.path.relpath(f, ROOT)
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern)); assert len(fs) == 1, (pattern, fs); return fs[0]

# ================= 公司設定 =================
NAME, TICKER, VEN = "Kioxia", "285A", "Kioxia"
U = "¥B"; PXU = "JPY"; MF = "#,##0"; EF = "#,##0"
FYS = [2026, 2027, 2028, 2029]
def cyq(fy, k):                              # 會計年度 fy 第 k 季 → 日曆季（4 月起算）
    return f"{fy}Q{k + 1}" if k <= 3 else f"{fy + 1}Q1"
QL = [(fy, k) for fy in FYS for k in (1, 2, 3, 4)]   # 16 季：C..R
QT = {(2026, 1): "A", (2026, 2): "G"}        # A＝實績、G＝公司指引（固定）；其餘＝D（驅動）
REFQ = (2026, 2)                             # 單位成本、營業費用的校準季（第二季指引）

# ================= 上游（唯讀） =================
NAND_F = only("markets/nand/dist/*.xlsx"); MU_F = only("micron/dist/*.xlsx")
wn = openpyxl.load_workbook(NAND_F, data_only=True)
CYQS = [f"{y}Q{q}" for y in range(2025, 2031) for q in (1, 2, 3, 4)]
UP = {}
for r in wn["Mkt"].iter_rows():
    lab = str(r[0].value or "")
    for s in SC:
        if lab == f"{VEN} 營收（情境 {s}）": UP[("rev", s)] = ([r[2 + j].value for j in range(24)], r[0].row)
        if lab.startswith(f"{VEN} 位元（情境 {s}"): UP[("bit", s)] = ([r[2 + j].value for j in range(24)], r[0].row)
        if lab == f"Kioxia＋SanDisk（合資群）份額（情境 {s}）": UP[("jv", s)] = ([r[2 + j].value for j in range(24)], r[0].row)
assert len(UP) == 9, sorted(UP)
JVSH = {}
for r in wn["Ann"].iter_rows():
    if r[0].value == "Kioxia＋SanDisk（合資群）份額":
        JVSH[SC[len(JVSH)]] = ([r[3 + j].value for j in range(6)], r[0].row)
assert len(JVSH) == 3
AIS = {}                                       # 情境 → [AI 企業級, 非 AI 企業級, PC]（CY2025..CY2030）, 列
for r in wn["Ann"].iter_rows():
    lab_ = r[0].value
    if lab_ == "  需求：AI 企業級 SSD": AIS[SC[len(AIS)]] = [[r[3 + j].value for j in range(6)], None, None, r[0].row]
    if lab_ == "  需求：非 AI 企業級 SSD" and AIS: AIS[SC[len(AIS) - 1]][1] = [r[3 + j].value for j in range(6)]
    if lab_ == "  需求：PC" and AIS: AIS[SC[len(AIS) - 1]][2] = [r[3 + j].value for j in range(6)]
assert len(AIS) == 3 and all(v[1] and v[2] for v in AIS.values())
wm = openpyxl.load_workbook(MU_F, data_only=True)
MUV = wm["Val"]
def mu_hist(title):
    for i in range(1, MUV.max_row + 1):
        if MUV.cell(i, 1).value == title:
            y0 = i + 2; assert MUV.cell(y0, 1).value == "2016" and MUV.cell(y0 + 9, 1).value == "2025"
            return [MUV.cell(y0 + j, 2).value for j in range(10)], f"{rel(MU_F)}｜Val!B{y0}:B{y0 + 9}"
    raise KeyError(title)
MU_PB, MU_PB_SRC = mu_hist("P/B（主）：十年歷史倍數")
MU_PE, MU_PE_SRC = mu_hist("P/E（對照）：十年歷史倍數")
MU_PE = [v if isinstance(v, (int, float)) else None for v in MU_PE]
MU_NIM = None
for r in wm["In"].iter_rows(min_row=4):
    if r[0].value == "nim10": MU_NIM = (r[4].value, f"{rel(MU_F)}｜In!E{r[0].row}")
assert MU_NIM

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億日圓（¥B）；每股：日圓（2026-10-01 一股分三股後）；股數：百萬股。每個輸入附標記與來源；找不到者寫「未找到」"
IN["A2"] = "標記：Verified（公司文件、交易所資料）／Interested-party（寫明是誰、立場）／Analogy（寫明對照與差異）／Assumed（給低／高區間）／Derived（寫明換算）。SemiAnalysis 未使用"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
ER = "Kioxia FY2026 第一季財報（2026-07-31；SGX 轉載 Earnings Release FY2026 1Q）https://links.sgx.com/1.0.0/corporate-announcements/DOATO9EIE4EEQ2Y8/898078_Earnings-Release-FY2026-1Q-EN.pdf"
SA = "stockanalysis 季度損益（TYO:285A；2026-10-25 讀取）"
MS = "MarketScreener 預估頁（2026-10-25 擷取；頁面為 10-16 收盤；會計年度 3 月；每股數字已依分割調整，股數 1,644 百萬股）"
sec("A. 季度實績與公司指引（¥B；EPS 日圓，分割後）")
QA25 = {1: (342.8, 71.2, 44.9, 18.3), 2: (448.3, 122.1, 85.9, 40.7), 3: (543.6, 176.6, 142.8, 87.8), 4: (1002.9, 642.9, 595.4, 407.7)}
for k, (rv, gp, op, ni) in QA25.items():
    tag = "Verified" if k == 1 else "Interested-party"
    src = (ER + "（前期欄：營收 342.8、營業利益 44.9、淨利 18.3）") if k == 1 else SA
    inp(f"p{k}_rev", f"FY2025 第{k}季營收", U, None, rv, None, tag, src, "四季合計 2,337.6＝公司 FY2025 營收（storagenewsletter 轉載決算；MarketScreener 2,337.6）")
    inp(f"p{k}_gp", f"FY2025 第{k}季毛利", U, None, gp, None, "Interested-party", SA, "")
    inp(f"p{k}_op", f"FY2025 第{k}季營業利益", U, None, op, None, tag, src, "四季合計 869.0（公司 870.4，取整差）")
    inp(f"p{k}_ni", f"FY2025 第{k}季淨利", U, None, ni, None, tag, src, "四季合計 554.5（公司 554.5）")
SEGQ = {1: (217.4, 79.0, "blocksandfiles 2025-08-11（轉載 FY2025 第一季決算；分部 46.3 為其他）", "Interested-party"),
        2: (246.6, 156.9, "longportapp（轉載 FY2025 第二季決算）：SSD 與儲存占 55%、智慧裝置占 35% × 營收 448.3（Derived）", "Derived"),
        3: (300.4, 186.3, "blocksandfiles 2026-02-13（轉載 FY2025 第三季決算；其他 57.0）", "Interested-party"),
        4: (600.3, 337.3, "blocksandfiles 2026-05-21（轉載 FY2025 第四季決算；其他 65.2）", "Interested-party")}
for k, (ss_, sd_, src_, tg_) in SEGQ.items():
    inp(f"p{k}_ss", f"FY2025 第{k}季營收：SSD 與儲存", U, None, ss_, None, tg_, src_, "官方分部：SSD & Storage（企業級、資料中心、PC 用 SSD）")
    inp(f"p{k}_sd", f"FY2025 第{k}季營收：智慧裝置", U, None, sd_, None, tg_, src_, "官方分部：Smart Devices（手機、車用等嵌入式）")
inp("q1_ss", "FY2026 第一季營收：SSD 與儲存", U, None, 1174.7, None, "Verified", ER, "")
inp("q1_sd", "FY2026 第一季營收：智慧裝置", U, None, 525.7, None, "Verified", ER, "其他 66.7（零售、對 SanDisk 銷售）＝營收 − 兩主分部")
inp("q1_rev", "FY2026 第一季營收（2026-04～06）", U, None, 1767.1, None, "Verified", ER, "SSD 與儲存 1,174.7、智慧裝置 525.7、其他 66.7；平均匯率 160 日圓")
inp("q1_gp", "FY2026 第一季毛利", U, None, 1380.1, None, "Interested-party", SA + "：1,380,066 百萬日圓", "財報摘要未列毛利")
inp("q1_op", "FY2026 第一季營業利益", U, None, 1270.0, None, "Verified", ER, "含訴訟損失準備 36.6、股份報酬 19.4（non-GAAP 加回）")
inp("q1_ni", "FY2026 第一季淨利（歸屬母公司）", U, None, 842.2, None, "Verified", ER, "")
inp("q1_eps", "FY2026 第一季基本 EPS（分割後）", PXU, None, "=1539.91/3", None, "Derived", ER + "：基本 EPS ¥1,539.91（分割前）÷ 3", "2026-10-01 一股分三股（2026-07-31 董事會）")
inp("g2_rev", "FY2026 第二季（2026-07～09）營收實績", U, None, 2480.0, None, "Verified", "Kioxia FY2026 第二季財報（2026-10-30；404K Research 法說整理、PANews／TechFlow 轉載）：¥2,480.0B（指引 2,390.0；平均匯率 161）", "v0.3 前為公司指引")
inp("g2_op", "FY2026 第二季營業利益實績", U, None, 1994.7, None, "Verified", "Kioxia FY2026 第二季財報（2026-10-30；404K Research 法說整理、PANews／TechFlow 轉載）：¥1,994.7B（non-GAAP 1,988.3；指引 1,890.0）", "")
inp("g2_ni", "FY2026 第二季淨利實績", U, None, 1360.0, None, "Verified", "Kioxia FY2026 第二季財報（2026-10-30；404K Research 法說整理、PANews／TechFlow 轉載）：約 ¥1.36T（non-GAAP 1,348.6；指引 1,270.0）", "第二來源：TechFlow 2026-10-30（1.36 兆）")
inp("g2_ox", "FY2026 第二季營業費用（推估；校準用）", U, 50, 59, 70, "Derived", "毛利率約 82.8%（404K Research 依損益表推算）× 2,480.0 − 營業利益 1,994.7 ≈ 59", "第二季營業成本＝營收 − 營業利益 − 本值；公司損益表原文未取得")
inp("a25_rev", "FY2025 營收（2025-04～2026-03）", U, None, "=p1_rev+p2_rev+p3_rev+p4_rev", None, "Derived", "四季合計", "公司 2,337.6（storagenewsletter 轉載 FY25 決算，+37%）")
inp("a25_ni", "FY2025 淨利", U, None, "=p1_ni+p2_ni+p3_ni+p4_ni", None, "Derived", "四季合計", "公司 554.0（storagenewsletter）、MarketScreener 554.5")
inp("a25_eps", "FY2025 EPS（分割後）", PXU, None, 341.4, None, "Interested-party", MS + "：341.4", "≈ 554.5 ÷ 平均約 1,624 百萬股（分割後，Derived）")
inp("kx_k", "公司營收 ÷ 市場層 Kioxia 營收（TrendForce 口徑 × 匯率）", "x", None, "=q1_rev/160/Units!H4", None, "Derived", "第一季營收 ÷ 平均匯率 160 ÷ 市場層 2026Q2 Kioxia 營收（Units 頁）", "公司營收含其他（¥66.7B）")

sec("B. 成本、費用、稅、匯率（公司特有；隨情境者 A 最有利、C 最不利）")
for s, (lo, b, hi), src in zip(SC, ((0.10, 0.14, 0.18), (0.08, 0.12, 0.16), (0.06, 0.10, 0.14)),
                               ("A：BiCS10（332 層，北上 K2）FY2026 量產、密度約 +59%", "B：轉換略延後", "C：需求轉弱、稼動率下降")):
    inp(f"cd_{s}", f"情境 {s}：每位元成本年降", "%", lo, b, hi, "Assumed", src, "與 sndk/ 同值（同一合資晶圓廠與 BiCS 世代）；以第二季指引校準的單位成本為起點" if s == "A" else "同上")
inp("fx", "匯率（FY2026 第三季起）", "¥／$", 145, 160, 170, "Assumed", "第一季平均 160、第二季指引假設 162", "營收以美元計價（TrendForce 口徑 × 匯率）；成本以日圓計，不隨匯率")
inp("ox_var", "營業費用變動率（÷ 營收）", "%", 0.005, 0.01, 0.02, "Assumed", "FY2025 營業費用（毛利 − 營業利益）143.8 ÷ 營收 2,337.6 ≈ 6%，多為固定", "固定部分＝第二季營業費用 − 變動率 × 第二季營收（Derived）")
inp("ox_g", "營業費用固定部分年增", "%", 0.05, 0.10, 0.15, "Assumed", "研發（BiCS10、CBA）與人事", "")
inp("nonop_q", "每季業外淨額（利息、匯兌等；FY2026 第三季起）", U, -40, -15, 0, "Assumed", ER + "：第一季稅前 − 營業利益＝−35.3；第二季指引 −20.0", "2026 年第一季已提前償還長期借款 ¥433.2B")
inp("tax", "有效稅率", "%", 0.28, 0.32, 0.35, "Derived", ER + "：第一季淨利 842.2 ÷ 稅前 1,234.7 → 稅率 31.8%；第二季指引 1,270 ÷ 1,870 → 32.1%", "")
for y, (lo_, b_, hi_) in ((2026, (0.60, 0.66, 0.72)), (2027, (0.60, 0.68, 0.76)), (2028, (0.60, 0.70, 0.78)), (2029, (0.60, 0.70, 0.78))):
    inp(f"sssh_{y % 100}", f"FY{y} SSD 與儲存占營收" + ("（第二至四季）" if y == 2026 else ""), "%", lo_, b_, hi_, "Assumed", "FY2025 第一至四季依序 63%、55%、55%、60%；FY2026 第一季 66.5%（Verified）", "低＝回到 FY2025 水準；分部拆分不影響合併營收與 EPS")
    inp(f"otsh_{y % 100}", f"FY{y} 其他占營收", "%", 0.03, 0.04, 0.06, "Assumed", "FY2026 第一季 3.8%；FY2025 第四季 6.5%（含對 SanDisk 銷售）", "智慧裝置＝1 − SSD 與儲存 − 其他（Derived）")
inp("ncy", "淨現金收益率", "%", 0.0, 0.01, 0.02, "Assumed", "日圓利率", "v0.5 起鍵名由 rf 改為 ncy（rf 改為無風險利率，標準版 DCF 用；值不變）")

sec("C. 資本、股東回饋（2026-06-30 資產負債表；股數為分割後）")
inp("eq_j", "2026-06-30 股東權益（歸屬母公司）", U, None, 2404.3, None, "Verified", ER, "2026-03-31 為 1,399.1")
inp("nc_j", "2026-06-30 淨現金（現金 − 公司債與借款；不含租賃負債）", U, None, "=791.0-634.6", None, "Derived", ER + "：現金 791.0、公司債與借款 634.6（另有租賃負債 203.3）", "")
inp("eq25", "FY2025 年底（2026-03-31）股東權益", U, None, 1399.1, None, "Verified", ER, "")
inp("sh0", "流通股數（分割後）", "百萬股", None, "=548.014638*3", None, "Derived", ER + "：2026-06-30 發行 548,015,088 股、庫藏 450 股；× 3（2026-10-01 分割）", "MarketScreener 1,644.0 百萬股")
inp("dil", "稀釋股數 − 流通股數", "百萬股", None, "=sh0*(1539.91/1525.09-1)", None, "Derived", ER + "：基本 EPS 1,539.91 ÷ 稀釋 1,525.09", "各年固定")
inp("da_y", "折舊攤銷（年）", U, 280, 320, 380, "Assumed", ER + "：第一季 76.1（× 4 ≈ 304）", "新廠（北上 K2）投產後增加")
for y, v, lo, hi in ((2026, 450, 400, 500), (2027, 470, 400, 550), (2028, 470, 400, 550), (2029, 470, 350, 550)):
    inp(f"cap_{y}", f"資本支出 FY{y}", U, lo, v, hi, "Interested-party", "TrendForce 2026-06-03：FY2026 ¥450B、FY2026–28 年均 ¥470B（較 FY2025 ¥283.7B +66%）" if y == 2026 else ("同上" if y < 2029 else "Assumed：延續 FY2026–28 年均"), "含合資廠 Kioxia 分攤；SanDisk 分攤另計於 sndk/")
inp("payout", "股東回饋（股利＋買回）÷ 自由現金流", "%", 0.0, 0.2, 0.5, "Assumed", "FY2025 無股利；MarketScreener 共識 FY2026 每股股利 ¥266（Interested-party）", "以買回計（股數減少）")
inp("px", "股價（2026-11-04 收盤）", PXU, None, 10735, None, "Verified", "Yahoo Finance 日線收盤 2026-11-04（第三輪修正統一價格日；分割後）", "v0.5.2 為 2026-10-23 收盤 ¥12,755（stockanalysis）；10-01 為 ¥19,100")
inp("bb_px", "未來買回均價", PXU, None, "=px", None, "Assumed", "＝現價", "")
# v0.5.3（第三輪修正）：2026-08 已執行的 ¥800B 買回（7-31 決議、8-3 起、8-10 公告取得完畢）扣現金並減股數；放在 C 節最後（其後各列下移，名稱引用不受影響）
inp("bb800", "已執行買回（2026-08；FY2026 第二季）", U, None, 800, None, "Verified", "公司 2026-07-31 決議（上限 3,000 萬股（分割前）／¥800B）、2026-08-10 公告取得完畢；SBクリエイティブ ビジネス+IT「キオクシア 2026年7〜9月期決算」2026-10-30（https://www.sbbit.jp/article/st/187973，2026-11-06 讀取）：「8,000億円の自社株買い後もネット・キャッシュは8,127億円」", "PL 頁 FY2026 權益、淨現金扣除；DCF 淨現金改用 9-30 公司淨現金（已含本筆）")
inp("bb800_sh", "已執行買回減少的股數（分割後）", "百萬股", None, "=sh0-47.9/90*1000*3", None, "Derived", "9-30 登記日流通股數（分割前）＝中間配當總額 ¥47.9B ÷ 每股 ¥90＝532.2 百萬股（SBクリエイティブ ビジネス+IT「キオクシア 2026年7〜9月期決算」2026-10-30（https://www.sbbit.jp/article/st/187973，2026-11-06 讀取））；× 3 分割後＝1,596.7 百萬股；本值＝6-30 股數 sh0 − 1,596.7", "隱含買回均價約 ¥16,900（分割後；¥800B ÷ 本值）；個別取得股數公告未取得")

sec("D. 共識與對帳參考（只放 Recon，不作校準）")
inp("cons_rev26", "共識營收 FY2026（2027-03）", U, 9100, 10071.4, 11400, "Interested-party", MS + "：10,071,396 百萬日圓", "stockanalysis（15 位）平均 10.1 兆、區間 9.1–11.4 兆")
inp("cons_rev27", "共識營收 FY2027（2028-03）", U, None, 13477.2, None, "Interested-party", MS, "")
inp("cons_ni26", "共識淨利 FY2026", U, None, 5573.9, None, "Interested-party", MS, "")
inp("cons_ni27", "共識淨利 FY2027", U, None, 7730.0, None, "Interested-party", MS, "")
inp("cons_eps26", "共識 EPS FY2026", PXU, None, 3486, None, "Interested-party", MS, "")
inp("cons_eps27", "共識 EPS FY2027", PXU, None, 4871, None, "Interested-party", MS, "")
inp("cons_eps28", "共識 EPS FY2028", PXU, None, 5467, None, "Interested-party", MS, "")
inp("cons_bv26", "共識每股淨值 FY2026 年底（2027-03）", PXU, None, 3830, None, "Interested-party", MS, "")
inp("cons_bv27", "共識每股淨值 FY2027 年底（2028-03）", PXU, None, 7808, None, "Interested-party", MS, "")
inp("cons_bv28", "共識每股淨值 FY2028 年底（2029-03）", PXU, None, 11928, None, "Interested-party", MS, "")
IN_ROW = {}; r = 4
TPK = set()
def flush():
    global r
    for srow in rows:
        if srow[0] == "sec":
            IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
        key, item, unit, lo, base, hi, tag, s1, s2 = srow
        IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
        for j, v in ((4, lo), (5, base), (6, hi)):
            if v is not None:
                IN.cell(r, j, v)
                if not (isinstance(v, str) and v.startswith("=")): IN.cell(r, j).font = BLUE
        if key in TPK:
            for j in range(1, 10): IN.cell(r, j).fill = KEY
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
        defname(key, f"In!$E${r}"); IN_ROW[key] = r
        if lo is not None: defname(key + "_lo", f"In!$D${r}")
        if hi is not None: defname(key + "_hi", f"In!$F${r}")
        r += 1
    rows.clear()
flush()
sec("E. 目標價參數（黃底；Val 頁第 0 節 DCF＋Comparable 與收斂法用；Assumed 者 Andy 可改）")
TPK.update(list(FD.JUDG) + ["tp_wA", "tp_wB", "tp_wC", "tp_g", "tp_w_dcf", "nim10", "cyc_mA", "cyc_mB", "cyc_mC", "fade_mA", "fade_mB", "fade_mC", "fade_gA", "fade_gB", "fade_gC", "cons_tp"])
for s_, w_ in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s_}", f"目標價情境權重 {s_}（A 順利／B 平平／C 不如預期）", "%", None, w_, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%")
inp("tp_g", "收斂法 B 的 FY2032 營收年增（v0.5 起不再用於 DCF；DCF 永續成長率見 M2 節 g_tv）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "v0.5 刪除 tp_r、tp_roe（舊股東現金流法 DCF 參數）；本值只決定 PL 延伸期收斂路徑 fade_gA／B／C")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（2026-10-25 11:53：DCF／Comparable 各半）", "")
inp("nim10", "十年平均淨利率（Analogy：美光 FY2016–FY2025）", "%", None, MU_NIM[0], None, "Analogy", MU_NIM[1],
    "Kioxia 2024-12 上市，2018 年以前為東芝記憶體事業，十年連續的獨立財報未找到；以同為記憶體循環股的美光十年平均代替（與 sndk/ 同值）")
inp("cyc_mA", "循環中段係數：情境 A 淨利率收斂目標 ÷ 十年平均", "x", None, 1.2, None, "Assumed", "記憶體樣板預設 1.2", "Andy 可改")
inp("cyc_mC", "循環中段係數：情境 C 淨利率收斂目標 ÷ 十年平均", "x", None, 0.8, None, "Assumed", "記憶體樣板預設 0.8", "Andy 可改")
for s_, f_, t_ in (("A", "=nim10*cyc_mA", "十年平均 × cyc_mA"), ("B", "=nim10*cyc_mB", "十年平均 × cyc_mB（＝1）"), ("C", "=nim10*cyc_mC", "十年平均 × cyc_mC")):
    inp(f"fade_m{s_}", f"情境 {s_}：FY2032 淨利率收斂目標（循環中段：{t_}）", "%", None, f_, None, "Derived", "主串指示：循環產業 FY2030 後收斂用十年平均（記憶體樣板的循環中段）", "FY2030 起由 FY2029 淨利率線性收斂，FY2032 達到")
for s_, f_, t_ in (("A", "=tp_g+0.02", "tp_g＋2pt"), ("B", "=tp_g", "＝tp_g"), ("C", "=tp_g-0.02", "tp_g−2pt")):
    inp(f"fade_g{s_}", f"情境 {s_}：FY2032 營收年增收斂目標（{t_}）", "%", None, f_, None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("cons_tp", "市場共識目標價（賣方平均；低／高＝最低／最高）", PXU, 16333.33, 36929.23, 66666.67, "Interested-party", "MarketScreener 共識頁（16 位，2026-10-25 擷取，頁面收盤日 10-16）：平均 ¥36,929（分割後）", "stockanalysis（16 位，2026-09-09）：平均 ¥111,256（分割前）÷ 3 ＝ ¥37,085")
flush()
# Siliconomics v0.13（決策 S37）：Wattonomics 錨定係數，只供 Sens「Wattonomics 錨定」案例用；放在模型輸入最後
sec(f"Z. Wattonomics 錨定係數（敏感度用；Siliconomics {SIL.VERSION}）")
SIL.wta_rows(inp)
flush()
# v0.5（2026-11-05）：標準版 DCF（無槓桿自由現金流 FCFF＋WACC；tools/fcff_dcf.py）參數。放在 In 頁模型輸入最後（Z 節之後），不位移上方各列。
# 評價日＝2026-12-31（v0.5.1 起統一評價日；原 2026-06-30）；FY2026 只計評價日後一季（frac0＝0.25）；淨現金沿用最近資產負債表 2026-06-30（現金 791.0、公司債與借款 634.6、租賃 203.3）；目標時點＝評價日後 12 個月。
_SA = "stockanalysis.com Kioxia（TYO:285A）財報頁，2026-11-05 讀取；百萬日圓；該站年度以結束年標示（其 FY2026＝本模型 FY2025）"
_BSRC = ("自算產業 β：記憶體 同業 5 家中位數（美光 MU、Western Digital WDC、SK 海力士 000660、三星電子 005930、南亞科 2408；Yahoo Finance 月報酬 2021-11～2026-10，各自對本地大盤迴歸，Hamada 去槓桿、現金修正；"
         "Kioxia、SanDisk 上市不足 5 年，不入中位數）；v0.5.2 起取代 KR 半導體 4 家（工作單第二輪修正 2026-11-05 23:20 新增記憶體組）")
_US10 = ("美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；2026-11-05 讀取）",
         "FRED DGS10 同序列；與股價同日（2026-10-23）；只用於 Sens「美元口徑折現」（工作單第二輪修正，Andy 2026-11-05 23:16）")
_DIV2 = "PANews／Phemex「Kioxia Q2 operating profit hits ¥1.995 trillion」2026-10-30（https://phemex.com/news/article/kioxia-q2-operating-profit-hits-1995-trillion-on-ai-storage-demand-q3-guidance-misses-expectations-101867，2026-11-05 讀取）"
for _row in FD.in_rows(dict(
        date="2026-12-31（統一評價日；FY2026 只計評價日後一季；v0.5.3 起資產負債表日 2026-09-30 的淨現金以 stub 期滾到評價日）", amt="¥B", ccy=PXU, shu="十億股",
        rf_us=(0.05, _US10[0], _US10[1]),
        stub=dict(bs_date="2026-09-30", val_date="2026-12-31", bs_src="2026-09-30 公司淨現金 ¥812.7B（已扣 8 月 ¥800B 買回；SBクリエイティブ ビジネス+IT「キオクシア 2026年7〜9月期決算」2026-10-30（https://www.sbbit.jp/article/st/187973，2026-11-06 讀取））；9-30 現金、借款、租賃分列未取得，借款與租賃沿用 6-30（" + ER + "）",
                  div=(47.9, "Verified", "第二季財報（2026-10-30）宣告中間配當每股 ¥90（分割前）、總額 ¥47.9B、效力發生日 12-08（SBクリエイティブ ビジネス+IT「キオクシア 2026年7〜9月期決算」2026-10-30（https://www.sbbit.jp/article/st/187973，2026-11-06 讀取））", "9-30 後支付，自 9-30 淨現金扣除；8 月 ¥800B 買回已在 9-30 淨現金內，不重複扣")),
        rf=(0.02817, "財務省 國債金利情報（https://www.mof.go.jp/english/policy/jgbs/reference/interest_rate/jgbcme.csv）：2026-11-02 10 年期 2.817%（本月檔最早一日；2026-11-04 為 2.855%）",
            "Trading Economics 2026-11-05 為 2.76%（https://tradingeconomics.com/japan/government-bond-yield）；10 月逐日值未取得，取 v0.5.2 現價日（10-23）之後最近可得值；第三輪修正現價改 2026-11-04（同日 2.855%），rf 未改（不在該輪範圍）", "無風險利率（日本 10 年期公債殖利率，2026-11-02）"),
        beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Derived",
                    _BSRC + "；Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；以本公司市值 D/E 與稅率重新加槓桿（Hamada：βL＝βU ×（1＋(1 − tax_w) × D/E）；Andy 2026-11-05 21:21）",
                    "Sens 低／高（Analogy；v0.5.2 起）：β_u ∓ beta_du（0.3），各自同式重新加槓桿（tools/fcff_dcf.py v2 write_sens）；個股迴歸 β 未取得（2024-12 才上市）"),
        crp=(0.0091, "Verified", "Damodaran Country Default Spreads and Risk Premiums（2026-01-05 更新；https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html；2026-11-05 讀取）：Japan（A1）國家風險溢酬 0.91%",
             "Andy 2026-11-05 20:45：非美國市場 ERP＝5%＋CRP，不設股權成本下限"),
        kd=("=rf+0.02", "Assumed", "rf ＋ 2.0pt 信用利差（BB 級附近；借款利率未逐筆查）", "負債占資本約 4%，對 WACC 影響小"),
        tax_w=("=tax", "Derived", "＝In 頁 tax（第一季 31.8%、第二季指引 32.1%；NOPAT 與淨利同一稅率）", "只用於債務稅盾"),
        debt=(634.6, "Verified", ER + "：2026-06-30 公司債與借款 634.6（In 頁 nc_j 同一來源）", _SA + "：2026-06 Total Debt 837.8（含租賃）"),
        lease=(203.3, "Verified", ER + "：2026-06-30 租賃負債 203.3", "nc_j 不含租賃，故評價時另扣"),
        cash=("=812.7+634.6", "Derived", "2026-09-30 公司淨現金 812.7 ＋ 借款 634.6（6-30 值；9-30 分列未取得）：使「現金 − 借款」＝公司公布的 9-30 淨現金（SBクリエイティブ ビジネス+IT「キオクシア 2026年7〜9月期決算」2026-10-30（https://www.sbbit.jp/article/st/187973，2026-11-06 讀取））", "6-30 為 791.0（" + ER + "）；本模型 6-30 口徑淨現金 156.4（現金 − 公司債與借款），公司定義差異未調整"),
        inv=(0.0, "Assumed", "無重大非營業投資（合資廠為營運資產）", "—"),
        mi=(0.158, "Verified", _SA + "：2026-06 Minority Interest 158", "—"),
        sh=("=(sh0-bb800_sh+dil)/1000", "Derived", "＝(9-30 流通股數 sh0 − bb800_sh ＋ 稀釋 dil) ÷ 1000（分割後；已扣 8 月買回）", "未來買回由 FCFF 支應，不另推股數"),
        dps=(0, "Assumed", "FY2025 無股利；共識 FY2026 每股股利 ¥266（Interested-party）未採用，股東回饋以買回計", "—"),
        t_tgt=(1.0, "目標價為評價日後 12 個月（2027-12-31）"),
        fc_da=(None, "=(310.599/1706.46+311.519/2337.628)/2", None, "Derived",
               _SA + "：FY2024、FY2025（本模型口徑）折舊攤銷 ÷ 營收＝18.2%／13.3%（D&A 310,599／311,519；營收 1,706,460／2,337,628）；FY2023 無資料，只用兩年",
               "明確期 FY2026–FY2029 用 PL 頁折舊（da_y）；本值為 FY2030 起收斂的長期穩態"),
        fc_cx=(None, "=(223.847/1706.46+281.062/2337.628)/2", None, "Derived",
               _SA + "：FY2024、FY2025 資本支出 ÷ 營收＝13.1%／12.0%（223,847／281,062）",
               "明確期 FY2026–FY2029 用 PL 頁資本支出（cap_2026–cap_2029，TrendForce 年均 ¥470B）"),
        fc_cx_ss=("=MAX(fc_cx,fc_da_ss*(1+g_tv))", "Derived", "＝MAX(近兩年資本支出比例, 穩態折舊 ×（1＋g）)：穩態資本支出不得低於折舊 ×（1＋g）（規格一致性）", "近兩年資本支出比例低於折舊，故取折舊 ×（1＋g）"),
        fc_wc=(0.20, "=((238.594+352.863+43.349-86.664-24.476)/1706.46+(660.559+412.612+66.76-159.52-2.35)/2337.628)/2", 0.50, "Derived",
               _SA + "：FY2024、FY2025 營運資金 ÷ 營收＝30.7%／41.8%；營運資金＝應收＋存貨＋其他流動資產 − 應付 − 應計費用（不含其他流動負債）",
               "FY2023 的應付科目分類不一致（應付 512.5 含其他應付），未用；Sens 低／高 20%／50%（Assumed）"),
        )):
    inp(*_row) if _row[0] != "sec" else sec(_row[1])
inp("beta_u", "產業無槓桿 β（自算記憶體組，Blume 調整後；採用值）", "x", None, "=0.67*beta_u_med+0.33", None, "Derived",
    _BSRC + "；＝0.67 × 中位數 beta_u_med ＋ 0.33", "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；同 SK 海力士、美光、三星、SanDisk；用現金修正值：淨現金另計，β 只對營業資產")
inp("beta_u_med", "產業無槓桿 β 中位數（自算記憶體組，Blume 調整前；只作來源對照）", "x", None, 2.294, None, "Derived", _BSRC, "v0.5.1 為 KR 半導體 1.72")
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3", "取代 v0.5.1 的中位數 × 0.8 再 Blume／中位數未調整")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 評價用股數 sh_v)", "與 WACC 權重同一口徑")
# v0.5.3（2026-11-06，第三輪修正第 39 項）：情境 B 循環中段係數補成輸入列（＝1；原 fade_mB 直接＝nim10，情境驅動對照 B 欄為佔位 0）。放在模型輸入最後，只有其後的情境驅動對照下移
sec("E2. 目標價參數（續）：情境 B 循環中段係數（v0.5.3）")
inp("cyc_mB", "循環中段係數：情境 B 淨利率收斂目標 ÷ 十年平均", "x", None, 1, None, "Assumed", "＝1：情境 B 收斂到十年平均淨利率（fade_mB＝nim10 × 本值）", "v0.5.3 前 fade_mB 直接＝nim10，數值不變；Andy 可改")
flush()
for col, w in zip("ABCDEFGHI", [12, 60, 10, 9, 11, 9, 18, 90, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

# ---------------- Units ----------------
UN = wb.create_sheet("Units")
UN["A1"] = f"Units — 市場層 markets/nand 成品的 {VEN} 季度營收與位元（三情境；唯讀抄錄，藍字）。來源：{rel(NAND_F)}"
UN["A2"] = "C..Z＝CY2025Q1..CY2030Q4；營收為 TrendForce 口徑（$B，美元），位元為季度出貨 × 營收份額（EB）。合資群份額＝Kioxia＋SanDisk（與 sndk/ 模型讀同一列）"
hdr(UN, 3, ["項目", "單位"] + CYQS + ["來源"])
UR = {}; ur = 4
for kind, lab, unit in (("rev", "營收", "$B"), ("bit", "位元", "EB"), ("jv", "合資群份額", "%")):
    for s in SC:
        vals, srow = UP[(kind, s)]
        UN.cell(ur, 1, f"{VEN} {lab}（情境 {s}）" if kind != "jv" else f"Kioxia＋SanDisk 合資群份額（情境 {s}）"); UN.cell(ur, 2, unit)
        for j, v in enumerate(vals):
            c = UN.cell(ur, 3 + j, v); c.font = BLUE; c.number_format = "0.0%" if kind == "jv" else "#,##0.000"
        UN.cell(ur, 27, f"{rel(NAND_F)}｜Mkt!C{srow}:Z{srow}")
        UR[(kind, s)] = ur; ur += 1
UN.column_dimensions["A"].width = 36; UN.column_dimensions["AA"].width = 70
UCOL = {q: GL(3 + i) for i, q in enumerate(CYQS)}
assert UCOL["2026Q2"] == "H" and UR[("rev", "A")] == 4
UN.cell(ur + 1, 1, "AI 企業級 SSD 占（企業級＋PC）SSD 位元需求（日曆年；markets/nand Ann 頁）").font = BOLD
hdr(UN, ur + 2, ["項目", "單位"] + [f"CY{y}" for y in range(2025, 2031)] + ["來源"])
AIR = {}
for i, s in enumerate(SC):
    a_, n_, pc_, srow = AIS[s]; rr_ = ur + 3 + i
    UN.cell(rr_, 1, f"AI ÷（企業級＋PC）SSD（情境 {s}）"); UN.cell(rr_, 2, "%")
    for j in range(6):
        c = UN.cell(rr_, 3 + j, float(a_[j]) / (float(a_[j]) + float(n_[j]) + float(pc_[j]))); c.font = BLUE; c.number_format = "0.0%"
    UN.cell(rr_, 9, f"{rel(NAND_F)}｜Ann!D{srow}:I{srow} ÷（同列＋下一列＋PC 列）"); AIR[s] = rr_
AICOL = {y: GL(3 + y - 2025) for y in range(2025, 2031)}
aiw = lambda fy, s: f"(0.75*Units!{AICOL[fy]}{AIR[s]}+0.25*Units!{AICOL[fy + 1]}{AIR[s]})"

# ---------------- Q：季度實績 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 季度實績（¥B）與公司指引。會計年度 4–3 月；第一季＝4–6 月"
hdr(Q, 3, ["項目", "FY2025 第一季", "FY2025 第二季", "FY2025 第三季", "FY2025 第四季", "FY2025", "FY2026 第一季", "FY2026 第二季", "標記", "來源"])
for i, (lab, key) in enumerate((("營收", "rev"), ("毛利", "gp"), ("營業利益", "op"), ("淨利", "ni")), 4):
    Q.cell(i, 1, lab)
    for k in range(1, 5): Q.cell(i, 1 + k, f"=p{k}_{key}").number_format = "#,##0.0"
    Q.cell(i, 6, f"=SUM(B{i}:E{i})").number_format = "#,##0.0"
    Q.cell(i, 7, f"=q1_{key}").number_format = "#,##0.0"
    if key != "gp": Q.cell(i, 8, f"=g2_{key}").number_format = "#,##0.0"
    Q.cell(i, 9, "Verified／Interested-party"); Q.cell(i, 10, "公司財報（第一季前期欄）；stockanalysis 季度表；第二季為公司指引")
Q.cell(8, 1, "毛利率"); [Q.cell(8, j, f"={GL(j)}5/{GL(j)}4").__setattr__("number_format", "0.0%") for j in range(2, 8)]
Q.cell(9, 1, "營業利益率"); [Q.cell(9, j, f"={GL(j)}6/{GL(j)}4").__setattr__("number_format", "0.0%") for j in range(2, 9)]
Q.cell(10, 1, "季增：營收"); [Q.cell(10, j, f"={GL(j)}4/{GL(j - 1)}4-1").__setattr__("number_format", "0.0%") for j in (3, 4, 5)]
Q["G10"] = "=G4/E4-1"; Q["H10"] = "=H4/G4-1"; Q["G10"].number_format = Q["H10"].number_format = "0.0%"
Q.cell(12, 1, "v0.3（2026-10-31）：FY2026 第二季（2026-07～09）實績已滾入（2026-10-30：營收 2,480.0、營業利益 1,994.7、淨利 1,360；平均匯率 161）；第三季起＝市場層推導。第三季指引：營收 2,500.0、營業利益約 2,000（假設 157 日圓）")
Q.column_dimensions["A"].width = 30
for c in "BCDEFGH": Q.column_dimensions[c].width = 14
Q.column_dimensions["J"].width = 60

# ---------------- 引擎 ----------------
QC = {q: GL(3 + i) for i, q in enumerate(QL)}
FYC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
DEFP = dict(cd=None, oxv="ox_var", oxg="ox_g", tax="tax", gox="g2_ox", pay="payout", fx="fx", pxm="1", shm="1", nop="nonop_q", vq=None)
def engine(ws, r0, title, s, **over):
    p = dict(DEFP); p.update(over); cd = p["cd"] or f"cd_{s}"
    R = {}; rr = [r0]
    ws.cell(r0, 1, title).font = BOLD; rr[0] += 1
    def row(key, label, unit, fs, fmt="#,##0.0", note=""):
        r_ = rr[0]; ws.cell(r_, 1, label); ws.cell(r_, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r_}"] = f; ws[f"{col}{r_}"].number_format = fmt
        if note: ws.cell(r_, 20, note)
        R[key] = r_; rr[0] += 1
    hdr(ws, rr[0], ["A. 季度", "單位"] + [f"FY{fy % 100} 第{k}季（{cyq(fy, k)}）" for fy, k in QL] + ["", "說明"]); rr[0] += 1
    DQ = [q for q in QL if q not in QT]
    ref = QC[REFQ]
    nq = lambda q: QL.index(q) - QL.index(REFQ)
    uc = lambda q: UCOL[cyq(*q)]
    A1, G2 = QC[(2026, 1)], QC[(2026, 2)]
    row("um", "市場層營收（TrendForce 口徑，美元，同字母情境）", "$B", {QC[q]: f"=Units!{uc(q)}{UR[('rev', s)]}*{p['shm']}" + (f"*{p['pxm']}" if nq(q) > 0 else "") + (f"*{p['vq'][q]}" if p['vq'] else "") for q in QL}, fmt="#,##0.000", note="markets/nand；Sens 可乘份額係數、價格係數（第三季起）")
    row("bit", "位元（市場層；同字母）", "EB", {QC[q]: f"=Units!{uc(q)}{UR[('bit', s)]}*{p['shm']}" + (f"*{p['vq'][q]}" if p['vq'] else "") for q in QL}, fmt="#,##0.000")
    row("rev", "營收", U, {A1: "=q1_rev", G2: "=g2_rev", **{QC[q]: f"={QC[q]}{R['um']}*kx_k*{p['fx']}" for q in DQ}}, note="第一、二季實績；之後＝市場層 × 公司口徑係數 × 匯率")
    fyc = lambda q: q[0] % 100
    row("ss", "SSD 與儲存", U, {A1: "=q1_ss", **{QC[q]: f"={QC[q]}{R['rev']}*sssh_{fyc(q)}" for q in QL if q != (2026, 1)}}, note="官方分部；第一季實績，之後＝營收 × 占比（In）")
    row("ai", "  SSD 與儲存：AI 相關（＝SSD 與儲存 × AI 企業級 SSD 占比）", U, {QC[q]: f"={QC[q]}{R['ss']}*{aiw(q[0], s)}" for q in QL}, note="命題子列（Derived）：會計年度（4 月起）≈ 0.75 × 當年＋0.25 × 次年；假設公司 SSD 組合與產業（企業級＋PC）相同")
    row("ssn", "  SSD 與儲存：非 AI", U, {QC[q]: f"={QC[q]}{R['ss']}-{QC[q]}{R['ai']}" for q in QL})
    row("ot", "其他", U, {A1: "=q1_rev-q1_ss-q1_sd", **{QC[q]: f"={QC[q]}{R['rev']}*otsh_{fyc(q)}" for q in QL if q != (2026, 1)}}, note="官方分部（零售、對 SanDisk 銷售）；實績＝營收 − 兩主分部")
    row("sd", "智慧裝置", U, {A1: "=q1_sd", **{QC[q]: f"={QC[q]}{R['rev']}-{QC[q]}{R['ss']}-{QC[q]}{R['ot']}" for q in QL if q != (2026, 1)}}, note="官方分部；預測季＝營收 − SSD 與儲存 − 其他")
    row("qchk", "檢查：分部合計 − 營收（季）", U, {QC[q]: f"={QC[q]}{R['ss']}+{QC[q]}{R['sd']}+{QC[q]}{R['ot']}-{QC[q]}{R['rev']}" for q in QL}, fmt="0.000")
    rcogs = rr[0] + 1
    row("uc", "每 EB 營業成本（第二季以實績校準，之後每年 × (1 − 成本年降)）", "¥B/EB", {A1: f"={A1}{rcogs}/{A1}{R['bit']}", G2: f"={G2}{rcogs}/{G2}{R['bit']}",
        **{QC[q]: f"=${ref}{rr[0]}*(1-{cd})^({nq(q)}/4)" for q in DQ}}, fmt="0.00")
    row("cogs", "營業成本", U, {A1: "=q1_rev-q1_gp", G2: f"=g2_rev-g2_op-{p['gox']}", **{QC[q]: f"={QC[q]}{R['bit']}*{QC[q]}{R['uc']}" for q in DQ}}, note="第二季＝實績營收 − 實績營業利益 − 推估營業費用")
    row("gm", "毛利率", "%", {QC[q]: f"=1-{QC[q]}{R['cogs']}/{QC[q]}{R['rev']}" for q in QL}, fmt="0.0%")
    row("ox", "營業費用", U, {A1: "=q1_gp-q1_op", G2: f"={p['gox']}", **{QC[q]: f"=({p['gox']}-{p['oxv']}*g2_rev)*(1+{p['oxg']})^({nq(q)}/4)+{p['oxv']}*{QC[q]}{R['rev']}" for q in DQ}}, note="固定部分以第二季校準（Derived）")
    row("op", "營業利益", U, {QC[q]: f"={QC[q]}{R['rev']}-{QC[q]}{R['cogs']}-{QC[q]}{R['ox']}" for q in QL})
    rr[0] += 1
    ncref = lambda q: "nc_j" if q[0] == 2026 else f"${FYC[q[0] - 1]}{{NC}}"
    row("ni", "淨利", U, {A1: "=q1_ni", G2: "=g2_ni", **{QC[q]: f"=({QC[q]}{R['op']}+{p['nop']}+ncy*{ncref(q)}/4)*(1-{p['tax']})" for q in DQ}}, note="（營業利益＋業外＋淨現金 × 收益率 ÷ 4）×（1 − 稅率）；FY2026 用 2026-06-30 淨現金")
    row("epsq", "季 EPS（稀釋，分割後）", PXU, {QC[q]: f"={QC[q]}{R['ni']}/(sh0+dil)*1000" for q in QL}, fmt=EF)
    rr[0] += 1
    hdr(ws, rr[0], ["B. 年度（會計年度）", "單位", "FY2025A", "FY2026", "FY2027", "FY2028", "FY2029", "", "", "", "", "", "", "", "", "", "", "", "", "說明"]); rr[0] += 1
    qs = lambda fy, key: f"=SUM({QC[(fy, 1)]}{R[key]}:{QC[(fy, 4)]}{R[key]})"
    F3 = (2027, 2028, 2029)
    row("yrev", "營收", U, {"C": "=a25_rev", **{FYC[y]: qs(y, "rev") for y in FYS}}, note="FY2026＝第一、二季實績＋第三、四季推導")
    row("yss", "營收：SSD 與儲存", U, {"C": "=p1_ss+p2_ss+p3_ss+p4_ss", **{FYC[y]: qs(y, "ss") for y in FYS}})
    row("yai", "  SSD 與儲存：AI 相關", U, {"C": f"=C{rr[0] - 1}*{aiw(2025, s)}", **{FYC[y]: qs(y, "ai") for y in FYS}})
    row("ysd", "營收：智慧裝置", U, {"C": "=p1_sd+p2_sd+p3_sd+p4_sd", **{FYC[y]: qs(y, "sd") for y in FYS}})
    row("yot", "營收：其他", U, {"C": f"=a25_rev-C{R['yss']}-C{rr[0] - 1}", **{FYC[y]: qs(y, "ot") for y in FYS}}, note="FY2025＝營收 − 兩主分部（含報導值進位差約 0.1）")
    row("ychk", "檢查：分部合計 − 合併營收（年）", U, {FYC[y]: f"={FYC[y]}{R['yss']}+{FYC[y]}{R['ysd']}+{FYC[y]}{R['yot']}-{FYC[y]}{R['yrev']}" for y in FYC}, fmt="0.000")
    row("ycogs", "營業成本", U, {"C": "=a25_rev-(p1_gp+p2_gp+p3_gp+p4_gp)", **{FYC[y]: qs(y, "cogs") for y in FYS}})
    row("ygm", "毛利率", "%", {FYC[y]: f"=1-{FYC[y]}{R['ycogs']}/{FYC[y]}{R['yrev']}" for y in FYC}, fmt="0.0%")
    row("yox", "營業費用", U, {FYC[y]: qs(y, "ox") for y in FYS})
    row("yop", "營業利益", U, {"C": "=p1_op+p2_op+p3_op+p4_op", **{FYC[y]: qs(y, "op") for y in FYS}})
    row("yom", "營業利益率", "%", {FYC[y]: f"={FYC[y]}{R['yop']}/{FYC[y]}{R['yrev']}" for y in FYC}, fmt="0.0%")
    row("yni", "淨利", U, {"C": "=a25_ni", **{FYC[y]: qs(y, "ni") for y in FYS}})
    row("da", "折舊攤銷", U, {FYC[y]: "=da_y" for y in FYS})
    row("cap", "資本支出（含合資廠分攤）", U, {FYC[y]: f"=cap_{y}" for y in FYS})
    row("fcf", "自由現金流（淨利＋折舊 − 資本支出）", U, {FYC[y]: f"={FYC[y]}{R['yni']}+{FYC[y]}{R['da']}-{FYC[y]}{R['cap']}" for y in FYS})
    G3 = f"(SUM({QC[(2026, 2)]}{R['ni']}:{QC[(2026, 4)]}{R['ni']})+0.75*(D{R['da']}-D{R['cap']}))"
    row("bb", "股東回饋（以買回計）", U, {"D": f"=MAX(0,{p['pay']}*{G3})", **{FYC[y]: f"=MAX(0,{p['pay']}*{FYC[y]}{R['fcf']})" for y in F3}}, note="FY2026 只計 2026-07 以後三季的自由現金流")
    row("eq", "期末權益", U, {"C": "=eq25", "D": f"=eq_j+SUM({QC[(2026, 2)]}{R['ni']}:{QC[(2026, 4)]}{R['ni']})-D{R['bb']}-bb800",
                              **{FYC[y]: f"={FYC[y - 1]}{rr[0]}+{FYC[y]}{R['yni']}-{FYC[y]}{R['bb']}" for y in F3}}, note="FY2026＝2026-06-30 權益＋第二至四季淨利 − 回饋 − 8 月已執行買回 bb800（不含其他綜合損益）")
    row("nc", "期末淨現金", U, {"D": f"=nc_j+{G3}-D{R['bb']}-bb800", **{FYC[y]: f"={FYC[y - 1]}{rr[0]}+{FYC[y]}{R['fcf']}-{FYC[y]}{R['bb']}" for y in F3}})
    row("sh", "期末流通股數（分割後）", "百萬股", {"D": f"=sh0-bb800_sh-D{R['bb']}/bb_px*1000", **{FYC[y]: f"={FYC[y - 1]}{rr[0]}-{FYC[y]}{R['bb']}/bb_px*1000" for y in F3}}, fmt="#,##0.0")
    row("eps", "EPS（稀釋：期末流通股數＋dil；分割後）", PXU, {"C": "=a25_eps", **{FYC[y]: f"={FYC[y]}{R['yni']}/({FYC[y]}{R['sh']}+dil)*1000" for y in FYS}}, fmt=EF)
    row("bv", "每股淨值 BVPS（期末流通股數）", PXU, {FYC[y]: f"={FYC[y]}{R['eq']}/{FYC[y]}{R['sh']}*1000" for y in FYS}, fmt=EF)
    row("roe", "ROE（期末權益）", "%", {FYC[y]: f"={FYC[y]}{R['yni']}/{FYC[y]}{R['eq']}" for y in FYS}, fmt="0%")
    row("cfps", "每股自由現金流（稀釋）", PXU, {FYC[y]: f"={FYC[y]}{R['fcf']}/({FYC[y]}{R['sh']}+dil)*1000" for y in F3}, fmt=EF)
    row("chk", "檢查：FY2025 營收、淨利 − 四季合計；第一季、第二季 − 實績與指引（應為 0）", U, {"D": f"=ABS({A1}{R['rev']}-q1_rev)+ABS({G2}{R['op']}-g2_op)+ABS(C{R['yrev']}-2337.6)"}, fmt="0.000")
    for q in DQ:
        c = ws[f"{QC[q]}{R['ni']}"]; c.value = c.value.replace("{NC}", str(R["nc"]))
    R["end"] = rr[0]
    return R

PL = wb.create_sheet("PL")
PL["A1"] = f"PL — {NAME} 損益（¥B；FY2025 實績；FY2026＝第一、二季實績＋第三、四季推導；FY2027–FY2029 由市場層 markets/nand 驅動；FY2030–FY2032 收斂法見底部）"
TITLES = {s: f"情境 {s}（{SCN[s]}）：markets/nand 情境 {s}（價格路徑 {s}、2028–29 顆數係數 {s}、合資群份額 {s}）；每位元成本年降 cd_{s}" for s in SC}
PB = {}; r0 = 56; TITLER = {}
for s in SC:
    TITLER[s] = r0; PB[s] = engine(PL, r0, TITLES[s], s); r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("yrev", "營收", U, "#,##0"), ("yss", "  SSD 與儲存", U, "#,##0"), ("yai", "    SSD 與儲存：AI 相關", U, "#,##0"), ("ysd", "  智慧裝置", U, "#,##0"), ("yot", "  其他", U, "#,##0"), ("ygm", "毛利率", "%", "0.0%"), ("yop", "營業利益", U, "#,##0"), ("yom", "營業利益率", "%", "0.0%"),
                                ("yni", "淨利", U, "#,##0"), ("eps", "EPS", PXU, EF), ("bv", "每股淨值", PXU, EF), ("fcf", "自由現金流", U, "#,##0")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y, c in FYC.items():
            if PL[f"{c}{PB[s][key]}"].value is None: continue
            PL[f"{c}{sr}"] = f"={c}{PB[s][key]}"; PL[f"{c}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對：三情境 FY2025、第一季、第二季 − 實績與指引（應為 0）"); PL[f"D{sr}"] = "=" + "+".join(f"D{PB[s]['chk']}" for s in SC); PL[f"D{sr}"].number_format = "0.000"; PL_CHK = sr
assert sr < 54
PL.column_dimensions["A"].width = 58; PL.column_dimensions["B"].width = 10; PL.column_dimensions["T"].width = 60
for c in "CDEFGHIJKLMNOPQR": PL.column_dimensions[c].width = 12

p0 = PL.max_row + 3
PL.cell(p0, 1, "延伸到 FY2032：FY2025–FY2029 引用本頁頂端摘要；灰底 FY2030–FY2032 為收斂法，非驅動因子推導").font = BOLD
PL.cell(p0 + 1, 1, "收斂法：營收年增由 FY2029 年增線性收斂到 FY2032 的長期成長（A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt）；淨利率由 FY2029 線性收斂到循環中段（A＝十年平均 × 1.2、B＝十年平均、C＝十年平均 × 0.8；十年平均為美光 Analogy）；股數維持 FY2029")
EXTC = {y: GL(3 + y - 2026) for y in range(2026, 2033)}
hdr(PL, p0 + 2, ["情境（延伸）", "單位", "FY2026", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
EXT = {}; pr = p0 + 3
for s in SC:
    PL.cell(pr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD; pr += 1
    R = {k: pr + i for i, k in enumerate(("rev", "g", "nim", "ni", "eps"))}
    for k, lab in (("rev", "營收"), ("g", "  營收年增"), ("nim", "  淨利率（淨利 ÷ 營收）"), ("ni", "淨利（收斂法延伸）"), ("eps", "EPS（延伸）")):
        PL.cell(R[k], 1, lab); PL.cell(R[k], 2, {"rev": U, "ni": U, "eps": PXU}.get(k, "%"))
    for y in range(2026, 2033):
        c = EXTC[y]; pc = EXTC.get(y - 1)
        if y <= 2029:
            for k, kk in (("rev", "yrev"), ("ni", "yni"), ("eps", "eps")): PL[f"{c}{R[k]}"] = f"={FYC[y]}{SUMR[(s, kk)]}"
            PL[f"{c}{R['nim']}"] = f"={c}{R['ni']}/{c}{R['rev']}"
            PL[f"{c}{R['g']}"] = f"={c}{R['rev']}/" + (f"{pc}{R['rev']}" if y > 2026 else f"C{SUMR[(s, 'yrev')]}") + "-1"
        else:
            k = y - 2029
            PL[f"{c}{R['g']}"] = f"=$F{R['g']}+(fade_g{s}-$F{R['g']})*{k}/3"
            PL[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
            PL[f"{c}{R['nim']}"] = f"=$F{R['nim']}+(fade_m{s}-$F{R['nim']})*{k}/3"
            PL[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
            PL[f"{c}{R['eps']}"] = f"=$F{R['eps']}*{c}{R['ni']}/$F{R['ni']}"
            for kk in R: PL[f"{c}{R[kk]}"].fill = FADE
        for kk, fm in (("rev", "#,##0"), ("g", "0.0%"), ("nim", "0.0%"), ("ni", "#,##0"), ("eps", EF)):
            if PL[f"{c}{R[kk]}"].value is not None: PL[f"{c}{R[kk]}"].number_format = fm
    PL.cell(R["rev"], 10, "FY2030 起為收斂法，非驅動因子推導")
    EXT[s] = R; pr += 6

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — 估值。第 0 節為目標價（DCF 與 Comparable 加權；Comparable 依記憶體慣例用前瞻 P/B）；第 5–6 節為目前倍數、十年倍數與規則值（舊法參考）"
V["A2"] = "Kioxia 2024-12 上市，只有兩個年底：十年 P/B、P/E 以美光（同為記憶體循環股）為 Analogy，見第 6 節。每股數字皆為 2026-10-01 一股分三股後"
YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
T = {}
def tset(rr, key, label, fs, fmt, note=""):
    V.cell(rr, 1, label); T[key] = rr
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{rr}"] = f; V[f"{c}{rr}"].number_format = fmt
    if note: V[f"E{rr}"] = note
t0 = 4
V.cell(t0, 1, "0. 目標價").font = BOLD
# v0.5（2026-11-05）：0.1 改為標準版 DCF（無槓桿自由現金流 FCFF＋WACC；tools/fcff_dcf.py；明細在 DCF 頁）。
# 摘要 18 列寫在舊 0.1 節位置（舊節第 5–24 列），0.2 節以後列號不變。
EPSR, CFR = {}, {}   # 舊股東現金流法的列（v0.5 刪除）；KIOXIA_ROWS 除錯輸出保留鍵
DCFW = wb.create_sheet("DCF")
DCFW["A1"] = ("DCF — 標準版（v0.5）：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例；FY2026 只計評價日後一季）；"
              "終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py；規格 docs/plan/個股DCF標準規格.md")
DCFW["A2"] = ("EBIT：FY2026–FY2029 取 PL 推導區塊「營業利益」；FY2030–FY2032（PL 延伸表只有淨利）由淨利還原：EBIT＝淨利 ÷（1 − tax）− 4 × nonop_q − ncy × FY2029 年底淨現金（與 PL 同一稅率、同一業外口徑）。"
              "折舊（da_y）、資本支出（cap_FY）：FY2026–FY2029 取 PL 頁 ÷ 營收，FY2030 起線性收斂到 fc_da_ss、fc_cx_ss；營運資金增加＝fc_wc × 營收增量（PL 頁未計營運資金，只在 DCF 計入）。")
FY7 = list(range(2026, 2033))
def _rv(s_, y): return f"PL!{FYC[y]}{PB[s_]['yrev']}" if y <= 2029 else f"PL!{EXTC[y]}{EXT[s_]['rev']}"
def _eb(s_, y): return f"PL!{FYC[y]}{PB[s_]['yop']}" if y <= 2029 else f"=PL!{EXTC[y]}{EXT[s_]['ni']}/(1-tax)-4*nonop_q-ncy*PL!{FYC[2029]}{PB[s_]['nc']}"
DPATHS = [dict(name=s_, label=f"情境 {s_}（{SCN[s_]}）", rev={y: _rv(s_, y) for y in FY7}, rev0=f"PL!C{PB[s_]['yrev']}",
               ebit={y: _eb(s_, y) for y in FY7}, tax={y: "tax" for y in FY7}, ebit_note="FY2030 起＝淨利 ÷（1 − tax）− 4 × nonop_q − ncy × FY2029 年底淨現金",
               da={y: f"PL!{FYC[y]}{PB[s_]['da']}" for y in range(2026, 2030)}, cx={y: f"PL!{FYC[y]}{PB[s_]['cap']}" for y in range(2026, 2030)}) for s_ in SC]
# v0.5.2：折舊、資本支出改用模組 v2 的 da／cx（PL 頁逐年列；FY2030 起收斂且資本支出 ≥ 折舊 ×（1＋g）），取代 v0.5 的覆寫迴圈；stub 期淨現金滾到評價日
RDCF = FD.write_detail(DCFW, 4, DPATHS, FY7, 0.25, title="三情境 FCFF 明細（¥B；FY2026＝2026-04～2027-03，評價日 2026-12-31 後只計 2027-01 起一季）", stub=True)
RDCF["end"] = FD.write_checks(DCFW, RDCF["end"], nc_bs="812.7-203.3-0.158",
                              nc_note="2026-09-30（¥B；v0.5.3 起以 stub 期滾到評價日 2026-12-31）：公司淨現金 812.7（已扣 8 月 ¥800B 買回）− 租賃 203.3 − 少數股權 0.158（後兩者為 6-30 值）") + 1
DCFW.column_dimensions["A"].width = 58; DCFW.column_dimensions["B"].width = 44
for _c in "CDEFGHI": DCFW.column_dimensions[_c].width = 13
assert FD.summary_rows(stub=True) <= 20   # 摘要寫在第 5–24 列，0.2 節第 26 列不動
SDCF = FD.write_summary(wb, V, t0 + 1, RDCF, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt=MF, beta_note="自算產業 β 重新加槓桿",
                        title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；v0.5.2；FY2026 評價日後一季至 FY2032 明確期＋永續；明細見 DCF 頁）")
T.update(h01=t0 + 1, hd1=SDCF["h"], dcfv=SDCF["dcfv"], dcf12=SDCF["v12"], dcftp=SDCF["dcf"], tvsh=SDCF["tvsh"], mult=SDCF["mult"])
rr = t0 + 10      # 0.2 節位置同 v0.4（q＝rr＋12＝第 26 列）
q = rr + 12
V.cell(q, 1, "0.2 Comparable（市場主流同業前瞻 P/B；記憶體慣例）").font = BOLD; T["h02"] = q
hdr(V, q + 1, ["同業與前瞻 P/B", "代號", "前瞻 P/B", "日期", "來源", "標記"]); T["phd"] = q + 1
PEERS = [("美光", "MU", "=795.23/(877.55/3.35)", "2026-11-04", "11-04 收盤 795.23（Yahoo Finance 日線收盤 2026-11-04）÷ FY2027E（2027-09）每股淨值（＝10-23 收盤 877.55 ÷ 頁面 P/B 3.35；MarketScreener 共識每股淨值（2026-10-25 擷取，未更新）；micron/ 模型同值）"),
         ("SK 海力士", "000660.KS", "=1396000/(1277000/1.35)", "2026-11-04", "11-04 收盤 1,396,000（Yahoo Finance 日線收盤 2026-11-04）÷ 2027E 每股淨值（＝10-23 收盤 1,277,000 ÷ 頁面 P/B 1.35；MarketScreener 共識每股淨值（2026-10-25 擷取，未更新）；skhynix/ 模型同值）"),
         ("三星電子", "005930.KS", "=218000/(215500/1.26)", "2026-11-04", "11-04 收盤 218,000（Yahoo Finance 日線收盤 2026-11-04）÷ 2027E 每股淨值（＝10-23 收盤 215,500 ÷ 頁面 P/B 1.26；MarketScreener 共識每股淨值（2026-10-25 擷取，未更新）；samsung/ 模型同值）"),
         ("SanDisk", "SNDK", "=1037.4/((313.8+553.1)/2)", "2026-11-04", "11-04 收盤 1,037.40（Yahoo Finance 日線收盤 2026-11-04；分割前）÷ MarketScreener 約 2027 年底每股淨值（FY2027、FY2028 年底平均；sndk/ 模型同口徑）")]   # v0.5.3：報價統一 2026-11-04 收盤
for i, (nm, tk, pb, dt, src) in enumerate(PEERS, q + 2):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pb); c.number_format = "0.00"
    if not isinstance(pb, str): c.font = BLUE
    V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
T["p1"] = q + 2; T["pl"] = q + 1 + len(PEERS)
q = T["pl"] + 1
tset(q, "pmed", "同業前瞻 P/B 中位數", [f"=MEDIAN(C{T['p1']}:C{T['pl']})", None, None], "0.00",
     "記憶體同業（美光、SK 海力士、三星、SanDisk；無 A 股）。WDC（硬碟）不列入 P/B（僅有落後 P/B），P/E 並列參考見下方。Kioxia 自身 11-04 收盤 ÷ FY2027 年底（2028-03）共識每股淨值 1.37x（記憶體三家、sndk/ 同業表同值）")
tset(q + 1, "fbv", "12 個月後的前瞻每股淨值（情境加權）", ["=" + "+".join(f"tp_w{s}*(0.25*PL!{FYC[2027]}{PB[s]['bv']}+0.75*PL!{FYC[2028]}{PB[s]['bv']})" for s in SC), None, None], EF,
     "同業倍數＝現價 ÷ 約 14 個月後（約 2027 年底）每股淨值；對應 12 個月後（約 2027-10）用約 2028 年底每股淨值＝¼ × FY2027 年底（2028-03）＋¾ × FY2028 年底（2029-03）（會計年度 3 月，依月份調整）")
tset(q + 2, "comptp", "Comparable 目標價", [f"=B{q}*B{q + 1}", None, None], MF, "同業中位數 × 前瞻每股淨值")
fwd = {2027: "5/12", 2028: "7/12"}
feps = lambda s_: "(" + "+".join(f"{w}*PL!{EXTC[y]}{EXT[s_]['eps']}" for y, w in fwd.items()) + ")"   # v0.5：情境 EPS 改直接讀 PL 延伸表（Val 舊 EPS 列刪除）
tset(q + 3, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=tp_wA*" + feps("A") + "+tp_wB*" + feps("B") + "+tp_wC*" + feps("C"), None, None], EF,
     "約 2027-10 時的未來 12 月（2027-11～2028-10）≈ 5／12 × FY2027（2027-04～2028-03）＋7／12 × FY2028（會計年度 3 月，依月份調整）")
q += 5
V.cell(q, 1, "並列參考：同業前瞻 P/E（不進目標價）").font = BOLD; T["h02b"] = q
hdr(V, q + 1, ["同業與前瞻 P/E", "代號", "前瞻 P/E", "日期", "來源", "標記"]); T["ehd"] = q + 1
PEERS_E = [("美光", "MU", "=795.23/173", "2026-11-04", "11-04 收盤 795.23（Yahoo Finance 日線收盤 2026-11-04）÷ MarketScreener FY2027 共識 EPS 173（micron/ 模型 In 頁）"),
           ("SK 海力士", "000660.KS", "=1396000/475721", "2026-11-04", "11-04 收盤 1,396,000（Yahoo Finance 日線收盤 2026-11-04）÷ MarketScreener 2027 共識 EPS 475,721（skhynix/ 模型 In 頁）"),
           ("三星電子", "005930.KS", "=218000/72279", "2026-11-04", "11-04 收盤 218,000（Yahoo Finance 日線收盤 2026-11-04）÷ MarketScreener 2027 共識 EPS 72,279（samsung/ 模型 In 頁）"),
           ("SanDisk", "SNDK", "=1037.4/217.4", "2026-11-04", "11-04 收盤 1,037.40（Yahoo Finance 日線收盤 2026-11-04；分割前）÷ MarketScreener FY2027（2027-06）共識 EPS 217.4"),
           ("Western Digital", "WDC", "=323.83/(356.13/12.33)", "2026-11-04", "11-04 收盤 323.83（Yahoo Finance 日線收盤 2026-11-04）÷ 前瞻 EPS（＝stockanalysis 統計頁 10-26 收盤 356.13 ÷ Forward P/E 12.33；硬碟為主，參考）；v0.5.2 為 10-23 的 11.91")]   # v0.5.3：報價統一 2026-11-04 收盤
for i, (nm, tk, pe, dt, src) in enumerate(PEERS_E, q + 2):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.number_format = "0.00"
    if not isinstance(pe, str): c.font = BLUE
    V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
T["e1"] = q + 2; T["el"] = q + 1 + len(PEERS_E)
q = T["el"] + 1
tset(q, "emed", "同業前瞻 P/E 中位數（並列參考）", [f"=MEDIAN(C{T['e1']}:C{T['el']})", None, None], "0.00")
tset(q + 1, "ecomp", "Comparable（P/E 並列參考）＝中位數 × 12 個月後前瞻 EPS", [f"=B{q}*B{T['feps']}", None, None], MF, "不進目標價：高峰期 EPS × 低倍數，與 P/B 法互相對照")
q += 3
V.cell(q, 1, "0.3 結論").font = BOLD; T["h03"] = q
hdr(V, q + 1, ["項目", "數值", "", "", "說明"]); T["hd3"] = q + 1
q += 2
tset(q, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B{T['comptp']}", None, None], MF, "tp_w_dcf × DCF 目標價＋（1 − tp_w_dcf）× Comparable 目標價；tp_w_dcf 見 In 頁")
tset(q + 1, "comp2", "Comparable 目標價（市場主流同業法）", [f"=B{T['comptp']}", None, None], MF, "引用 0.2（前瞻 P/B）")
tset(q + 2, "cons", "市場共識目標價", ["=cons_tp", None, None], MF, "MarketScreener（16 位，2026-10-16 收盤時）：平均 ¥36,929（最低 16,333、最高 66,667；分割後）；stockanalysis（16 位，09-09）分割調整後 ¥37,085")
tset(q + 3, "upw", "上檔空間：加權目標價", [f"=B{q}/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-11-04 收盤）")
tset(q + 4, "upc", "上檔空間：Comparable 目標價", [f"=B{q + 1}/px-1", None, None], "0.0%")
x = f"B{q + 3}"
tset(q + 5, "tier", "上檔空間級距", [f'=IF({x}>0.5,"大於 50%",IF({x}>0.25,"25%–50%",IF({x}>=0,"0%–25%",IF({x}>=-0.25,"0%～−25%",IF({x}>=-0.5,"−25%～−50%","小於 −50%")))))', None, None], "@", "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）")
tset(q + 6, "rec", "投資建議", [f'=IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))', None, None], "@", "依上檔空間固定規則（Andy 2026-10-25 14:17）：>+50% 強力買進、+25～+50% 買進、−25～+25% 中立、−50～−25% 放空、<−50% 強力放空")
tset(q + 7, "dcfpx", "DCF ÷ 現價", [f"={c}{T['dcf12']}/px" for c in "BCD"], '0.00"x"', "各情境 DCF 價值（推到 12 個月後）÷ 現價（A／B／C）")
CBV = "(0.25*cons_bv27+0.75*cons_bv28)"
tset(q + 8, "gbv", "與市場共識目標價的差距：EPS 因素（記憶體以每股淨值計）", [f"=B{T['fbv']}/{CBV}-1", None, None], "0.0%",
     "本模型約 2028 年底每股淨值（情境加權）÷ 共識（¼ FY2027＋¾ FY2028 年底）− 1（Comparable 用 P/B，故以每股淨值代 EPS）")
tset(q + 9, "geps", "  對照：前瞻 EPS ÷ 共識前瞻 EPS − 1", [f"=B{T['feps']}/(5/12*cons_eps27+7/12*cons_eps28)-1", None, None], "0.0%", "共識前瞻 EPS＝5／12 × FY2027 4,871＋7／12 × FY2028 5,467（MarketScreener）")
tset(q + 10, "gmul", "與市場共識目標價的差距：倍數因素", [f"=(cons_tp/{CBV})/B{T['pmed']}-1", None, None], "0.0%")
V[f"E{q + 10}"] = f'="共識目標價隱含 P/B "&TEXT(cons_tp/{CBV},"0.00")&"x（÷ 共識約 2028 年底每股淨值），同業中位數 "&TEXT(B{T["pmed"]},"0.00")&"x"'
tset(q + 11, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D{T["dcf12"]},"#,##0")&"～"&TEXT(B{T["dcf12"]},"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
T["meth"] = q + 13
V.cell(q + 13, 1, "目標價方法")

b5 = 80
V.cell(b5, 1, "5. 目前倍數（僅供參考）").font = BOLD; hdr(V, b5 + 1, ["項目", "A／數值", "B", "C", "備註"]); CMH = b5 + 1
fr = b5 + 2; CUR = {}
for key, nm, fs, fm, note in (("px", "現價（2026-11-04 收盤）", ["=px"], EF, "Yahoo Finance 日線收盤 2026-11-04（分割後）"),
                              ("pe26", "P/E：現價 ÷ FY2025 EPS 341.4", ["=px/a25_eps"], "0.0", "落後（分割後）"),
                              ("pb26", "P/B：現價 ÷ 2026-06-30 每股淨值", ["=px/(eq_j/sh0*1000)"], "0.00", "權益 2,404.3 ÷ 1,644.0 百萬股"),
                              ("pe27", "現價 ÷ FY2027 EPS（模型）", [f"=px/PL!E{SUMR[(s, 'eps')]}" for s in SC], "0.0", "A／B／C"),
                              ("pe28", "現價 ÷ FY2028 EPS（模型）", [f"=px/PL!F{SUMR[(s, 'eps')]}" for s in SC], "0.0", "A／B／C"),
                              ("pb28", "現價 ÷ FY2028 年底每股淨值（模型）", [f"=px/PL!F{SUMR[(s, 'bv')]}" for s in SC], "0.00", "A／B／C；第 0 節歷史分位檢查用")):
    V.cell(fr, 1, nm)
    for j, f in enumerate(fs): V.cell(fr, 2 + j, f).number_format = fm
    V.cell(fr, 5, note); CUR[key] = fr; fr += 1
def _y(s, y, k): return f"=PL!{FYC[y]}{PB[s][k]}"
def _cases(bp, bt, lp, lt):
    return [{"case": "見頂", "pos": "見頂年 → 25 分位", "q": 25, "basis_label": lp, "basis": bp},
            {"case": "中段", "pos": "中段 → 中位數", "q": 50, "basis_label": lp, "basis": bp},
            {"case": "谷底", "pos": "谷底 → 75 分位", "q": 75, "basis_label": lt, "basis": bt}]
vseries = [{"id": "P/B", "name": "P/B（主）", "sources": [{"label": "美光（Analogy；會計年度 8 月底）", "vals": MU_PB, "tag": "Analogy", "note": "來源：" + MU_PB_SRC + "（macrotrends）；Kioxia 自身 2024-12 上市，2024、2025 年底倍數未找到"}],
            "cases": _cases(_y("A", 2027, "bv"), _y("C", 2028, "bv"), "FY2027 年底每股淨值（A）", "FY2028 年底每股淨值（C）")},
           {"id": "P/E", "name": "P/E（對照）", "sources": [{"label": "美光（Analogy；SEC 10-K 換算）", "vals": MU_PE, "tag": "Analogy", "note": "來源：" + MU_PE_SRC + "；虧損年不計"}],
            "cases": _cases(_y("A", 2027, "eps"), _y("C", 2028, "eps"), "FY2027 EPS（A）", "FY2028 EPS（C）")}]
nr = fr + 2
tr = VR.add_rule_block(V, nr, vseries, EF, "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底倍數依循環位置取分位。Kioxia 無十年自身歷史，以美光為 Analogy（同為記憶體循環股；差異：美光 DRAM 為主）。見頂與中段用 FY2027 A、谷底用 FY2028 C。")
PB0 = nr + 5; PB10 = f"$B${PB0}:$B${PB0 + 9}"
assert V.cell(PB0, 1).value == "2016" and V.cell(PB0 + 9, 1).value == "2025"
hq = T["band"] + 1
tset(hq, "pr", "歷史倍數檢查：現價隱含 P/B 的十年分位（A／B／C）",
     [f"=IF(px/PL!F{SUMR[(s, 'bv')]}>=MAX({PB10}),1,IF(px/PL!F{SUMR[(s, 'bv')]}<=MIN({PB10}),0,PERCENTRANK({PB10},px/PL!F{SUMR[(s, 'bv')]})))" for s in SC], "0%",
     "現價 ÷ 各情境 FY2028 年底每股淨值，在 2016–2025 年底 P/B（第 6 節，美光 Analogy）分布的分位（高於最大值記 100%、低於最小值記 0%）；只作檢查，不進目標價")
V.cell(T["meth"], 2, "DCF：股東現金流 FY2027–FY2029 用每股（淨利＋折舊－資本支出；資本支出 ¥470B／年，TrendForce），FY2030–FY2032 收斂期用 EPS ×（1 − g／ROE），r 10%、g 4%、ROE 為 FY2024–FY2025 平均；明確期＋永續，基準日 2026 年底。"
                       "Comparable：美光、SK 海力士、三星、SanDisk 前瞻 P/B 中位數 × 約 2028 年底每股淨值（情境加權）；P/E 並列參考不進目標價。"
                       "偏誤：FY2027–FY2029 為高峰期盈餘、收斂期淨利率以美光十年平均代替；營收以美元計、成本以日圓計，日圓貶值（匯率）使盈餘偏高。")
o0 = tr
V.cell(o0, 1, "舊法：規則值情境加權（參考）").font = BOLD; T["old"] = o0
hdr(V, o0 + 1, ["項目", "數值", "", "", "說明"])
AD = {V.cell(i, 1).value: i for i in range(nr, tr) if isinstance(V.cell(i, 1).value, str) and V.cell(i, 1).value.startswith("採用 ")}
t1 = o0 + 2
V.cell(t1, 1, "規則值情境加權價（參考，非目標價）"); V.cell(t1, 2, f"=tp_wA*B{t1 + 2}+tp_wB*B{t1 + 3}+tp_wC*B{t1 + 4}").number_format = MF
V.cell(t1, 5, "權重 × 各情境隱含價；約為年底價格、未折現")
V.cell(t1 + 1, 1, "相對現價（規則值情境加權價 ÷ 現價 − 1）"); V.cell(t1 + 1, 2, f"=B{t1}/px-1").number_format = "0.0%"
for k_, (s_, lab, q_) in enumerate((("A", "情境 A（順利）：75 分位 × FY2027 年底每股淨值（A）", "75 分位"), ("B", "情境 B（平平）：中位數 × FY2027 年底每股淨值（B）", "中位數"), ("C", "情境 C（不如預期）：25 分位 × FY2027 年底每股淨值（C）", "25 分位"))):
    src = AD[f"採用 {q_}（P/B）"]
    V.cell(t1 + 2 + k_, 1, lab); V.cell(t1 + 2 + k_, 2, f"=B{src}*PL!E{SUMR[(s_, 'bv')]}").number_format = MF
V.cell(t1 + 5, 1, "舊法說明"); V.cell(t1 + 5, 2, "十年（2016–2025 年底）P/B 分位（美光 Analogy）× 各情境 FY2027 年底每股淨值：A＝75 分位、B＝中位數、C＝25 分位。第 6 節見頂／中段／谷底規則值供 AI 投資地圖抽數。")
for k in ("dcftp", "comptp", "tpw"): V[f"B{T[k]}"].fill = KEY
assert T["meth"] < b5 - 1
V.column_dimensions["A"].width = 56
for c in "BCDFGH": V.column_dimensions[c].width = 15
V.column_dimensions["E"].width = 70

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A）：對 FY2027、FY2028 營收與 EPS 的影響。每個案例為一個完整推導區塊（本頁下方，與 PL 同一引擎）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收 ¥B", "FY2027 EPS", "相對基準", "FY2028 營收 ¥B", "FY2028 EPS", "相對基準", "FY2027 毛利率"])
CASES = [("市場層營收 −20%（價格；第三季起）", "價格係數 0.8", dict(pxm="0.8")), ("市場層營收 +20%（價格）", "價格係數 1.2", dict(pxm="1.2")),
         ("份額 ×0.9（營收與位元）", "份額係數 0.9", dict(shm="0.9")), ("份額 ×1.1", "份額係數 1.1", dict(shm="1.1")),
         ("匯率 145 日圓", "fx 低", dict(fx="fx_lo")), ("匯率 170 日圓", "fx 高", dict(fx="fx_hi")),
         ("每位元成本年降 10%（C 的值）", "cd_C", dict(cd="cd_C")), ("每位元成本年降 18%", "cd_A 高", dict(cd="cd_A_hi")),
         ("第二季營業費用 50（校準）", "g2_ox 低", dict(gox="g2_ox_lo")), ("第二季營業費用 70", "g2_ox 高", dict(gox="g2_ox_hi")),
         ("營業費用變動率 2%", "ox_var 高", dict(oxv="ox_var_hi")), ("稅率 35%", "tax 高", dict(tax="tax_hi")), ("稅率 28%", "tax 低", dict(tax="tax_lo")),
         ("股東回饋 ÷ FCF 50%", "payout 高", dict(pay="payout_hi")), ("股東回饋 0", "payout 低", dict(pay="payout_lo"))]
def put(i, nm, st, R, pre=""):
    SE.cell(i, 1, nm); SE.cell(i, 2, st)
    SE[f"C{i}"] = f"={pre}E{R['yrev']}"; SE[f"D{i}"] = f"={pre}E{R['eps']}"; SE[f"F{i}"] = f"={pre}F{R['yrev']}"; SE[f"G{i}"] = f"={pre}F{R['eps']}"
    SE[f"E{i}"] = f"=D{i}/D$4-1"; SE[f"H{i}"] = f"=G{i}/G$4-1"; SE[f"I{i}"] = f"={pre}E{R['ygm']}"
    for c in "CF": SE[f"{c}{i}"].number_format = "#,##0"
    for c in "DG": SE[f"{c}{i}"].number_format = EF
    for c in "EHI": SE[f"{c}{i}"].number_format = "0.0%"
put(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
base = 60
for n, (nm, st, ov) in enumerate(CASES):
    Rn = engine(SE, base + n * 50, f"案例區塊：{nm}（情境 A）", "A", **ov)
    put(5 + n, nm, st, Rn)
SN = 5 + len(CASES)
put(SN, "情境 B", "markets/nand B；成本年降 12%", PB["B"], "PL!"); put(SN + 1, "情境 C", "markets/nand C；成本年降 10%", PB["C"], "PL!")
SENS_LAST = SN + 1
SE.cell(SN + 3, 1, "讀法：第二季為公司指引（固定），價格與匯率係數只作用在第三季起；營業費用校準改變成本起點（第二季營業成本＝營收 − 營業利益 − 營業費用）。")
assert SN + 3 < base
s0 = SE.max_row + 3
# v0.5：目標價參數敏感度改為標準版 DCF 的標準案例（tools/fcff_dcf.py；每格為閉式公式，與 DCF 頁明細同式）；舊 tp_r／tp_g 案例刪除
_TS = FD.write_sens(SE, s0, RDCF, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref=f"Val!$B${T['comptp']}", base_ref=f"Val!$B${T['tpw']}", fmt=MF,
                    title="目標價參數（v0.5 標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）；基準＝三情境加權（Val 頁第 0.3 節）",
                    beta_du="beta_du", usd=True)   # v0.5.2：β＝β_u ∓ beta_du 重新加槓桿；Sens「美元口徑折現」（非美國公司）
T["sens0"] = s0; T["sens_last"] = _TS[1]
# ---- Wattonomics 錨定（三層校準原則；Siliconomics v0.13，決策 S37）：引擎區塊放在 Sens 既有內容最下方，摘要列接在案例清單（情境 C 列）之後 ----
WTAL = {}
_sc = None
for r in wn["Ann"].iter_rows():
    a_ = str(r[0].value or "")
    if a_.startswith("情境 ") and len(a_) > 3 and a_[3] in SC: _sc = a_[3]
    if _sc == "B" and a_ in ("  需求：AI 企業級 SSD", "位元需求合計", "位元供給", "位元出貨") and a_ not in WTAL:
        WTAL[a_] = ([r[4 + j].value for j in range(5)], r[0].row)      # CY2026..CY2030（Ann E..I）
assert len(WTAL) == 4, WTAL
WB0 = SE.max_row + 3
SE.cell(WB0, 1, "Wattonomics 錨定（三層校準原則）").font = BOLD
SE.cell(WB0 + 1, 1, "讀法：Wattonomics 錨定：以情境 B 為底，產業量 CY2027 起 × wta（Siliconomics SIF_WTAnchor_Base＝R4 比值）；FY2026 維持錨點。三層校準原則（Andy 2026-11-04）；個股基準暫不改，2027-01～02 雲端資本支出指引後再決定。")
SE.cell(WB0 + 2, 1, "產業量＝markets/nand 的 AI 企業級 SSD 位元需求（加速器顆數 × 每顆容量；情境 B）：CY2027 起 × wta（CY2030 沿用 wta_29）→ 需求合計 → 出貨＝MIN(需求, 供給) → 季度位元（市場層同一平滑式）；"
    f"{NAME} 營收與位元 × 季度出貨比（錨定 ÷ 情境 B），價格、份額、成本年降照情境 B；會計年度依既有日曆季對應。資料中心 AI 子列仍用情境 B 占比（不進合計）")
hdr(SE, WB0 + 3, ["日曆年（markets/nand 情境 B；藍字＝唯讀抄錄）", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "CY2030", "來源"])
WAC = {y: GL(3 + y - 2026) for y in range(2026, 2031)}
WRW = {}
def wrow_(r_, key, lab, unit, fs, fmt="#,##0.0", blue=False, note=""):
    SE.cell(r_, 1, lab); SE.cell(r_, 2, unit); WRW[key] = r_
    for col, f in fs.items():
        c = SE[f"{col}{r_}"]; c.value = f; c.number_format = fmt
        if blue: c.font = BLUE
    if note: SE.cell(r_, 8, note)
r_ = WB0 + 4
wrow_(r_, "f", "Wattonomics 錨定係數（In 頁 wta_27／28／29；CY2030 沿用 wta_29）", "x", {WAC[2026]: 1, WAC[2027]: "=wta_27", WAC[2028]: "=wta_28", WAC[2029]: "=wta_29", WAC[2030]: "=wta_29"}, "0.000"); r_ += 1
for key, lab in (("dai", "  需求：AI 企業級 SSD"), ("dem", "位元需求合計"), ("sup", "位元供給"), ("ship", "位元出貨")):
    vals, srow = WTAL[lab]
    wrow_(r_, key, f"{lab.strip()}（情境 B）", "EB", {WAC[y]: float(vals[j]) for j, y in enumerate(range(2026, 2031))}, blue=True, note=f"{rel(NAND_F)}｜Ann!E{srow}:I{srow}"); r_ += 1
wrow_(r_, "demw", "位元需求合計（錨定）＝需求 − AI 需求 ×（1 − 係數）", "EB", {WAC[y]: f"={WAC[y]}{WRW['dem']}-{WAC[y]}{WRW['dai']}*(1-{WAC[y]}{WRW['f']})" for y in range(2026, 2031)}); r_ += 1
wrow_(r_, "shipw", "位元出貨（錨定）＝MIN(需求, 供給)", "EB", {WAC[y]: f"=MIN({WAC[y]}{WRW['demw']},{WAC[y]}{WRW['sup']})" for y in range(2026, 2031)}); r_ += 2
hdr(SE, r_, ["日曆季", "單位"] + CYQS); r_ += 1
def bq_(shp, q):
    y, k = int(q[:4]), int(q[-1]); g = f"({WAC[y]}{WRW[shp]}/{WAC[y - 1]}{WRW[shp]})"
    return f"={WAC[y]}{WRW[shp]}/4*{g}^(({k}-2.5)/4)/(({g}^(-0.375)+{g}^(-0.125)+{g}^(0.125)+{g}^(0.375))/4)"
QW = [q for q in CYQS if int(q[:4]) >= 2027]
wrow_(r_, "bqb", "季度位元出貨（情境 B；市場層平滑式）", "EB", {UCOL[q]: bq_("ship", q) for q in QW}, "#,##0.00"); r_ += 1
wrow_(r_, "bqw", "季度位元出貨（錨定）", "EB", {UCOL[q]: bq_("shipw", q) for q in QW}, "#,##0.00"); r_ += 1
wrow_(r_, "vq", "季度出貨比（錨定 ÷ 情境 B；CY2026 以前＝1）", "x", {UCOL[q]: (f"={UCOL[q]}{WRW['bqw']}/{UCOL[q]}{WRW['bqb']}" if q in QW else 1) for q in CYQS}, "0.000"); r_ += 2
RW = engine(SE, r_, "案例區塊：Wattonomics 錨定（情境 B 為底；市場層營收與位元 × 季度出貨比）", "B", vq={q: f"${UCOL[cyq(*q)]}${WRW['vq']}" for q in QL})
put(SN + 2, "Wattonomics 錨定（情境 B 為底）", "產業量 CY2027 起 × wta；相對欄對情境 B 列", RW)
SE[f"E{SN + 2}"] = f"=D{SN + 2}/D{SN}-1"; SE[f"H{SN + 2}"] = f"=G{SN + 2}/G{SN}-1"
SE.column_dimensions["A"].width = 40; SE.column_dimensions["B"].width = 30
for c in "CDEFGHI": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司實績、指引與共識（共識只作對照，未用於校準）。差距＝模型 ÷ 對照 − 1（百分點列為相減）"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rcr = 4
def rec(label, ref, getter, src, note="", fmt="#,##0", pt=False):
    global rcr
    RC.cell(rcr, 1, label); RC.cell(rcr, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv = 3 + 2 * j; RC.cell(rcr, cv, getter(s)).number_format = fmt
        L_ = GL(cv); RC.cell(rcr, cv + 1, f"={L_}{rcr}-B{rcr}" if pt else f"={L_}{rcr}/B{rcr}-1").number_format = "0.0%"
    RC.cell(rcr, 9, src); RC.cell(rcr, 10, note); rcr += 1
RC.cell(rcr, 1, "A. 實績與公司指引").font = BOLD; rcr += 1
rec("FY2025 營收（四季合計）vs 公司 2,337.6", 2337.6, lambda s: f"=PL!C{SUMR[(s, 'yrev')]}", "storagenewsletter 轉載 FY25 決算；MarketScreener", "")
rec("FY2026 上半營收 vs 公司指引 4,157.1", 4157.1, lambda s: f"=PL!{QC[(2026, 1)]}{PB[s]['rev']}+PL!{QC[(2026, 2)]}{PB[s]['rev']}", "Kioxia FY2026 第一季財報（2026-07-31）", "恆等式")
rec("FY2026 第三季營收（推導）vs 第三季指引 2,500（假設 157 日圓）", "=g3_rev", lambda s: f"=PL!{QC[(2026, 3)]}{PB[s]['rev']}", "Kioxia 第二季財報（2026-10-30）第三季指引", "v0.4：市場層 2026Q4 季價 +8.1%（實現價口徑；v0.3 為合約價 +17.5%、差距 +21.5%），份額取 2026Q3 實績與第三季指引的平均。餘差：匯率假設 160 對 157 約 +1.9%；其餘為公司稱 7–9 月提前出貨、部分遞延到 1–3 月（同一會計年度內）與產品組合")
rec("FY2026 第二季毛利率（校準值）vs 第一季（百分點）", "=q1_gp/q1_rev", lambda s: f"=PL!{QC[(2026, 2)]}{PB[s]['gm']}", "Kioxia FY2026 第一季財報（2026-07-31）", "第二季營業費用為 Assumed（g2_ox）", "0.0%", True)
RC.cell(rcr, 1, "B. 共識（MarketScreener 2026-10-25 擷取；只作參考）").font = BOLD; rcr += 1
rec("FY2026 營收 vs 共識", "=cons_rev26", lambda s: f"=PL!D{SUMR[(s, 'yrev')]}", MS)
rec("FY2026 淨利 vs 共識", "=cons_ni26", lambda s: f"=PL!D{SUMR[(s, 'yni')]}", MS)
rec("FY2026 EPS vs 共識", "=cons_eps26", lambda s: f"=PL!D{SUMR[(s, 'eps')]}", MS, "差距超過 30% 見 README")
rec("FY2027 營收 vs 共識", "=cons_rev27", lambda s: f"=PL!E{SUMR[(s, 'yrev')]}", MS)
rec("FY2027 EPS vs 共識", "=cons_eps27", lambda s: f"=PL!E{SUMR[(s, 'eps')]}", MS)
rec("FY2028 EPS vs 共識", "=cons_eps28", lambda s: f"=PL!F{SUMR[(s, 'eps')]}", MS)
rec("FY2027 年底每股淨值 vs 共識", "=cons_bv27", lambda s: f"=PL!E{SUMR[(s, 'bv')]}", MS)
RC.cell(rcr, 1, "C. 市場層一致性").font = BOLD; rcr += 1
RC.cell(rcr, 1, "Kioxia 營收 ÷ 市場層（公司口徑係數 kx_k）"); RC.cell(rcr, 3, "=kx_k").number_format = "0.0000"; RC.cell(rcr, 9, "In 頁"); rcr += 1
RC.cell(rcr, 1, "Kioxia＋SanDisk 合資群份額 CY2027（市場層 Ann 頁，A／B／C）")
for j, s in enumerate(SC): RC.cell(rcr, 3 + 2 * j, JVSH[s][0][2]).number_format = "0.0%"
RC.cell(rcr, 9, f"{rel(NAND_F)}｜Ann!F{JVSH['A'][1]}（sndk/ 模型讀同一組列）"); rcr += 1
RC_LAST = rcr - 1
RC.column_dimensions["A"].width = 52
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 50; RC.column_dimensions["J"].width = 50

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源"; SO["A1"].font = BOLD
srcs = [("Kioxia FY2026 第一季財報（2026-07-31；SGX 轉載）", "https://links.sgx.com/1.0.0/corporate-announcements/DOATO9EIE4EEQ2Y8/898078_Earnings-Release-FY2026-1Q-EN.pdf"),
        ("Kioxia：第二季財報 2026-10-30 18:45（JST）公布（公司新聞頁）", "https://www.kioxia-holdings.com/en-jp/news/2026.html"),
        ("storagenewsletter：Kioxia FY25 決算（營收 2,337.6、營業利益 870、淨利 554）", "https://www.storagenewsletter.com/?p=302700"),
        ("stockanalysis：季度損益、歷史股價、共識", "https://stockanalysis.com/quote/tyo/285A/financials/?p=quarterly"),
        ("MarketScreener：Kioxia 預估與共識目標價", "https://uk.marketscreener.com/quote/stock/KIOXIA-HOLDINGS-CORPORATI-180030735/finances/"),
        ("MarketScreener：SanDisk 預估（同業 P/B、P/E）", "https://uk.marketscreener.com/quote/stock/SANDISK-CORPORATION-182801077/finances/"),
        ("TrendForce：Kioxia FY26–28 資本支出（2026-06-03）", "https://www.trendforce.com/news/2026/06/03/news-kioxia-targets-%C2%A5470b-yearly-capex-in-fy26-28-up-66-from-fy25-reportedly-weighing-third-kitakami-fab-and-ma/"),
        ("newsquawk：一股分三股公告（2026-07-31）", "https://www.newsquawk.com/headlines/kioxia-285a-jt-q1-jpy-net-profit-84217bln-exp-9736bln-revenue-177tln-prev-3428bln-yy-announces-a-3-for-1-stock-split-and-announces-a-share-buyback"),
        ("stockanalysis：WDC 統計（Forward P/E）", "https://stockanalysis.com/stocks/wdc/statistics/"),
        (f"市場層：{rel(NAND_F)}（markets/nand v0.1）", "本 repo"),
        (f"美光十年倍數與淨利率（Analogy）：{rel(MU_F)}", "本 repo"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 80; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
VA = lambda k, c="B": f"Val!{c}{T[k]}"
OBS = [
 f'="1. 加權目標價 ¥"&TEXT({VA("tpw")},"#,##0")&"，相對現價 ¥"&TEXT(px,"#,##0")&"（2026-11-04 收盤）"&TEXT({VA("upw")},"+0.0%;−0.0%")&"。主因是現價的倍數低於同業：Kioxia 現價 ÷ 2028-03 共識每股淨值 1.63x，同業中位數 "&TEXT({VA("pmed")},"0.00")&"x（SanDisk 2.86x、美光 3.35x），Comparable ¥"&TEXT({VA("comptp")},"#,##0")&"（"&TEXT({VA("upc")},"+0%;−0%")&"）；DCF（v0.5 標準版 FCFF＋WACC）¥"&TEXT({VA("dcftp")},"#,##0")&IF({VA("dcftp")}>px," 也高於現價"," 低於現價")&"（現價為 FY2027 EPS（A）的 "&TEXT(Val!B{CUR["pe27"]},"0.0")&" 倍）。股價 10-09 至 10-23 已跌 27%（17,505 → 12,755）。"',
 f'="2. 本模型 FY2026 EPS（A）¥"&TEXT(PL!D{SUMR[("A", "eps")]},"#,##0")&"，比共識 ¥3,486 "&TEXT(PL!D{SUMR[("A", "eps")]}/cons_eps26-1,"+0%;−0%")&"；上半年已由兩季實績鎖定（淨利 ¥2,202B）。v0.3 滾入第二季（營收 ¥2,480.0B、營業利益 ¥1,994.7B，皆高於指引）。v0.4 對帳：模型第三季營收 ¥"&TEXT(Recon!C7,"#,##0")&"B，比指引 ¥2,500B "&TEXT(Recon!D7,"+0.0%;−0.0%")&"（v0.3 為 +21.5%）。原因是市場層把合約價當實現價（2026Q3 實現 +25%、2026Q4 改 +8.1%，兩季累計不變），且份額只用出貨提前的單一季（改取兩季平均）；餘差為匯率假設與公司出貨季別移動（同一會計年度內）。"',
 f'="3. 營收以美元計、成本以日圓計：匯率 145／170 日圓使 FY2028 EPS（A）變動 "&TEXT(Sens!H9,"+0%;−0%")&"／"&TEXT(Sens!H10,"+0%;−0%")&"；2027 起季價 ±20% 為 "&TEXT(Sens!H5,"+0%;−0%")&"／"&TEXT(Sens!H6,"+0%;−0%")&"。情境 C 的 FY2029 毛利率降到 "&TEXT(PL!G{SUMR[("C", "ygm")]},"0%")&"。追蹤：TrendForce 每季初的 NAND 合約價、日圓匯率、SanDisk 2026-10-29 財報（合資群位元）。"']
PROP = ("Kioxia 是全球第三、四大 NAND 廠（2Q26 營收市占 13.6%），與 SanDisk 共用四日市、北上合資產能；盈餘由 NAND 合約價循環、AI 推論帶動的企業級 SSD 位元與日圓匯率決定。"
        "問題是 2026 年的高價能維持多久（FY2027 起三情境分歧），以及現價約 2.6 倍前瞻 P/E 是否已反映循環見頂。")
lines = [(f"{NAME}（{TICKER} JP）收支與估值模型 v0.1（AI 半導體 Project：記憶體——NAND）　2026-10-25", True),
         ("模型命題：" + PROP, False),
         ("重點觀察", True), (OBS[0], False), (OBS[1], False), (OBS[2], False),
         (f'="投資建議："&{VA("rec")}&"（依上檔空間固定規則；上檔空間："&TEXT({VA("upw")},"+0.0%;−0.0%")&"，級距："&{VA("tier")}&"）"', True), ("", False),
         ("驅動因子 → 推導量（因果方向：左邊決定右邊）", True),
         ("Units｜markets/nand 成品：Kioxia 季度營收（市場營收 × 份額，美元；同字母情境）與位元（季度出貨 × 份額）", False),
         ("PL 區塊｜營收＝市場層 × 公司口徑係數 × 匯率；營業成本＝位元 × 每 EB 成本（日圓；第二季以實績校準，之後每年 × (1 − 成本年降)）；營業費用＝固定＋變動", False),
         ("PL 區塊｜淨利＝（營業利益＋業外＋淨現金收益）×（1 − 稅率）；自由現金流＝淨利＋折舊 − 資本支出（¥470B／年）；股東回饋＝20% × FCF（以買回計）→ 股數 → EPS、每股淨值", False),
         ("Val｜第 0 節目標價＝DCF（v0.5 起為標準版：無槓桿自由現金流 FCFF＋WACC，FY2026 評價日後一季至 FY2032＋永續；明細在 DCF 頁）與 Comparable（同業前瞻 P/B × 約 2028 年底每股淨值）各半；第 5–6 節目前倍數、十年倍數規則值（美光 Analogy）", False), ("", False),
         ("會計年度：4 月到隔年 3 月（FY2026＝2026-04～2027-03 ≈ CY2026Q2–CY2027Q1）。前瞻 EPS 依月份調整：12 個月後（約 2027-10）的未來 12 月＝5／12 FY2027＋7／12 FY2028。每股數字為 2026-10-01 一股分三股後。", False),
         ("情境：A＝順利、B＝平平、C＝不如預期，讀 markets/nand 同字母（價格＝markets/memory NAND_S；顆數＝SCEN_UNITS；合資群份額年變動）；公司特有變數（每位元成本年降 14%／12%／10%，與 SanDisk 同值）A 最有利、C 最不利。FY2026 的 B、C 略高於 A：上游 B、C 價格路徑 2027Q1 漲幅較大。FY2030 起為收斂法，非驅動因子推導。", False),
         ("十年歷史不足：2024-12 上市，只有兩個年底；十年 P/B、P/E 與十年平均淨利率以美光為 Analogy（v0.5 起 DCF 不再用 ROE）。本版未做先驗，無 Andy 合理 P/E 對照。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 200
README_ROWS = {"thesis": 2, "obs": [4, 5, 6], "rec": 7, "drivers": [10, 11, 12, 13], "limits": 17}
assert RD.cell(17, 1).value.startswith("十年歷史不足") and RD.cell(9, 1).value.startswith("驅動因子")

order = ["README", "In", "Q", "Units", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "kioxia")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
_r = IN.max_row + 2
IN.cell(_r, 1, "v0.3 追加：FY2026 第三季公司指引（只作 Recon 對照）").font = BOLD
_r += 1
IN.cell(_r, 1, "g3_rev"); IN.cell(_r, 2, "公司指引：FY2026 第三季（2026-10～12）營收"); IN.cell(_r, 3, U); IN.cell(_r, 5, 2500.0).font = Font(color="0000FF")
IN.cell(_r, 7, "Interested-party"); IN.cell(_r, 8, "Kioxia FY2026 第二季財報（2026-10-30；404K Research 法說整理、PANews／TechFlow 轉載）：¥2,500.0B、營業利益約 ¥2,000B（假設 1 美元＝157 日圓）"); IN.cell(_r, 9, "公司：匯率不利，第三季營收與獲利與第二季大致持平")
wb.defined_names["g3_rev"] = DefinedName("g3_rev", attr_text=f"In!$E${_r}")
apply_in_layout.apply(wb, "v0.2.1", "2026-10-30")  # v0.2.1：In 頁版面（個股模型原則第 8 節；數值不變）
# v0.5.3（第三輪修正第 15 項）：情境驅動對照的收斂法列名改中性（共用工具取情境 C 的列名）
for _r in range(4, IN.max_row + 1):
    _lab = {"scn_fade_g": "FY2032 營收年增收斂目標（tp_g ± 2pt）", "scn_cyc_m": "循環中段係數：淨利率收斂目標 ÷ 十年平均", "scn_fade_m": "FY2032 淨利率收斂目標（循環中段：十年平均 × 係數）"}.get(str(IN.cell(_r, 1).value or ""))
    if _lab: IN.cell(_r, 2, _lab)
# ---------------- v0.5（2026-11-05）：HTML 新版面 v1.0（個股 HTML 版面規格 v1.0；docs/plan/個股HTML版面規格.md）＋標準版 DCF 的展示層 ----------------
# README 頁最後新增「展示層」「命題回覆」兩段（HTML 讀；數字一律由公式帶出）；原 README 各列保留不動。寫法比照 sndk/build_sndk.py v0.4.1。
# 文字中的 {鍵|格式} 轉成 TEXT(儲存格,"格式")、{鍵} 直接引用（文字格）。
import re as _re  # noqa: E402
def _srow(prefix):
    for r_ in range(4, SE.max_row + 1):
        if str(SE.cell(r_, 1).value or "").startswith(prefix):
            return r_
    raise KeyError(prefix)
_s = {k: _srow(p) for k, p in (("pdn", "市場層營收 −20%"), ("pup", "市場層營收 +20%"), ("sdn", "份額 ×0.9"), ("sup", "份額 ×1.1"),
                               ("fdn", "匯率 145 日圓"), ("fup", "匯率 170 日圓"), ("cdn", "每位元成本年降 10%"), ("cup", "每位元成本年降 18%"))}
_W = {s_: f"tp_w{s_}" for s_ in SC}
_bvye = lambda y0, y1: "(" + "+".join(f"{_W[s_]}*(0.25*PL!{FYC[y0]}{SUMR[(s_, 'bv')]}+0.75*PL!{FYC[y1]}{SUMR[(s_, 'bv')]})" for s_ in SC) + ")"
_vb = lambda k, c="B": f"Val!{c}{T[k]}"
REF = {"tpw": _vb("tpw"), "dcftp": _vb("dcftp"), "comp": _vb("comptp"), "px": "px", "upw": _vb("upw"), "upc": _vb("upc"), "rec": _vb("rec"), "tier": _vb("tier"),
       "wdcf": "tp_w_dcf", "pmed": _vb("pmed"), "fbv": _vb("fbv"), "feps": _vb("feps"), "geps": _vb("geps"), "gbv": _vb("gbv"), "cons": _vb("cons"),
       "dcfup": f"({_vb('dcftp')}/px-1)", "ceps26": "cons_eps26", "ceps27": "cons_eps27", "cfeps": "(5/12*cons_eps27+7/12*cons_eps28)",
       "cpb": f"(cons_tp/{CBV})", "cpbx": f"((cons_tp/{CBV})/{_vb('pmed')})", "bv27": _bvye(2026, 2027)}
REF["pxbv27"] = f"(px/{REF['bv27']})"
REF["bvg"] = f"({REF['fbv']}/{REF['bv27']}-1)"
REF["pmg"] = f"({REF['pmed']}/{REF['pxbv27']}-1)"
_dc, _cc = f"(tp_w_dcf*({REF['dcftp']}-px))", f"((1-tp_w_dcf)*({REF['comp']}-px))"
REF["dcfsh"], REF["compsh"] = f"({_dc}/({_dc}+{_cc}))", f"({_cc}/({_dc}+{_cc}))"
REF["src"] = (f'IF(SIGN({_dc})<>SIGN({_cc}),IF(ABS({_cc})>ABS({_dc}),"Comparable 在現價之上、DCF 在現價之下，差距淨額來自 Comparable","DCF 在現價之上、Comparable 在現價之下，差距淨額來自 DCF"),'
              f'IF({REF["dcfsh"]}>=0.6,"差距主要來自 DCF",IF({REF["dcfsh"]}<=0.4,"差距主要來自 Comparable（同業倍數法）","差距由 DCF 與 Comparable 約各半貢獻")))')
REF["cdrv"] = f'IF(ABS({REF["pmg"]})>ABS({REF["bvg"]}),"倍數差大於每股淨值成長","每股淨值成長大於倍數差")'
REF["pmdir"] = f'IF({REF["pmg"]}>=0,"高","低")'
REF["updesc"] = f'IF({REF["upw"]}>0.25,"遠高於",IF({REF["upw"]}>=0,"略高於",IF({REF["upw"]}>=-0.25,"低於","遠低於")))'
REF["verdict"] = f'IF({REF["upw"]}>0.25,"現價低於估值、有明顯上檔",IF({REF["upw"]}>=0,"現價大致合理、略有上檔",IF({REF["upw"]}>=-0.25,"現價高於估值","現價明顯高於估值")))'
_dcfpx = [f"Val!{c}{T['dcfpx']}" for c in "BCD"]
REF["ndcf"] = "(" + "+".join(f"({x}>1)" for x in _dcfpx) + ")"
REF["dcfwho"] = f'IF({REF["ndcf"]}=3,"三種情境的 DCF 價值都高於現價",IF({REF["ndcf"]}=0,"三種情境的 DCF 價值都低於現價","只有"&IF({_dcfpx[0]}>1,"情境 A ","")&IF({_dcfpx[1]}>1,"情境 B ","")&IF({_dcfpx[2]}>1,"情境 C ","")&"的 DCF 價值高於現價"))'
for s_, c in zip(SC, "BCD"):
    REF[f"dcfpx{s_}"] = f"Val!{c}{T['dcfpx']}"
    REF[f"dcf{s_}"] = f"Val!{c}{T['dcf12']}"
    REF[f"tvsh{s_}"] = f"Val!{c}{T['tvsh']}"
    for y in (2025, 2026, 2027, 2028, 2029):
        REF[f"eps{s_}{y % 100}"] = f"PL!{FYC[y]}{SUMR[(s_, 'eps')]}"
        REF[f"gm{s_}{y % 100}"] = f"PL!{FYC[y]}{SUMR[(s_, 'ygm')]}"
        REF[f"sssh{s_}{y % 100}"] = f"(PL!{FYC[y]}{SUMR[(s_, 'yss')]}/PL!{FYC[y]}{SUMR[(s_, 'yrev')]})"
        REF[f"aish{s_}{y % 100}"] = f"(PL!{FYC[y]}{SUMR[(s_, 'yai')]}/PL!{FYC[y]}{SUMR[(s_, 'yrev')]})"
    _rng = f"PL!{FYC[2026]}{SUMR[(s_, 'eps')]}:{FYC[2029]}{SUMR[(s_, 'eps')]}"
    REF[f"peak{s_}"] = f"INDEX(PL!{FYC[2026]}3:{FYC[2029]}3,MATCH(MAX({_rng}),{_rng},0))"
    REF[f"vs26{s_}"] = f'IF({REF[f"eps{s_}26"]}>cons_eps26,"高於","低於")'
    REF[f"vs27{s_}"] = f'IF({REF[f"eps{s_}27"]}>cons_eps27,"高於","低於")'
for k, r_ in _s.items():
    REF[f"l_{k}"] = f"Sens!A{r_}"; REF[f"e_{k}"] = f"Sens!H{r_}"
_ab = lambda *ks: "MAX(" + ",".join(f"ABS(Sens!H{_s[k]})" for k in ks) + ")"
REF["s_top"] = (f'IF({_ab("pdn", "pup")}>=MAX({_ab("sdn", "sup")},{_ab("fdn", "fup")},{_ab("cdn", "cup")}),"NAND 價格",'
                f'IF({_ab("fdn", "fup")}>=MAX({_ab("sdn", "sup")},{_ab("cdn", "cup")}),"日圓匯率",IF({_ab("sdn", "sup")}>={_ab("cdn", "cup")},"份額","每位元成本")))')
_t0 = _TS[0]
REF.update(wacc="wacc", ke="ke", beta="beta", betau="beta_u", rfv="rf", erp="erp", crp="crp", gtv="g_tv",
           s_wlo=f"Sens!D{_t0}", s_whi=f"Sens!F{_t0}", s_blo=f"Sens!D{_t0 + 1}", s_bhi=f"Sens!F{_t0 + 1}", s_elo=f"Sens!D{_t0 + 2}", s_ehi=f"Sens!F{_t0 + 2}", s_wclo=f"Sens!D{_t0 + 6}", s_wchi=f"Sens!F{_t0 + 6}",
           dcfnc="dcf_nc", stubB=f"Val!C{SDCF['stub']}", ncrB=f"Val!C{SDCF['ncr']}", stubdiv="stub_div", rfus="rf_us")
_usd = next(r_ for r_ in range(_TS[0], _TS[1] + 1) if SE.cell(r_, 1).value == "美元口徑折現")   # v0.5.2：Sens 美元口徑列（模組 v2 插在 DCF 權重案例之前）
REF.update(usd=f"Sens!C{_usd}", usdp=f"Sens!D{_usd}")
_tv = [f"Val!{c}{T['tvsh']}" for c in "BCD"]   # v0.5.2：終值占 EV 超過 80% 的情境（規格第 4 節；展示層標示）
REF["tvflag"] = ('IF(MAX(' + ",".join(_tv) + ')>0.8,"終值占 EV 超過 80% 的情境："&' + '&'.join(f'IF({t_}>0.8,"{s_} "&TEXT({t_},"0%")&" ","")' for s_, t_ in zip("ABC", _tv))
                 + '&"（DCF 價值多半來自終值，對 g、RONIC 與收斂期假設敏感）。","")')
_SP = "+0.0%;−0.0%"

def _fx(t):
    def lit(x):
        x = x.replace('"', '""')
        return [f'"{x[i:i + 200]}"' for i in range(0, len(x), 200)]
    parts, pos = [], 0
    for m in _re.finditer(r"\{([A-Za-z0-9_]+)(?:\|([^}]*))?\}", t):
        parts += lit(t[pos:m.start()])
        e = REF[m.group(1)]
        parts.append(e if m.group(2) is None else f'TEXT({e},"{m.group(2).replace("SP", _SP)}")')
        pos = m.end()
    parts += lit(t[pos:])
    parts = [x for x in parts if x != '""']
    return t if len(parts) == 1 and parts[0].startswith('"') else "=" + "&".join(parts)

PRES_LEAD = "命題：Kioxia 是全球第三、四大 NAND 廠，與 SanDisk 共用四日市、北上合資產能；盈餘由 NAND 合約價循環、AI 推論帶動的企業級 SSD 位元與日圓匯率決定；股價取決於目前的高價能維持多久，以及現價隱含的倍數是否已反映循環見頂。三個問題如下。"
PRES_PROPS = [
 "命題 1：盈餘由什麼決定：NAND 價格、日圓匯率、Kioxia 份額與每位元成本，哪一個對 EPS 的影響最大；SSD 與儲存（含 AI 企業級 SSD）占營收多少。",
 "命題 2：高價能維持多久：三種價格情境下，EPS 在哪一年到高點、FY2027 起分歧多大。",
 "命題 3：現價是否已反映循環見頂：用現金流折現與同業倍數估出的每股價值，和現價相比是高還是低；現價隱含的倍數與同業差多少。"]
PRES_OBS = [
 "重點 1：合理價值{updesc}現價，{src}。加權目標價約 ¥{tpw|#,##0}（DCF ¥{dcftp|#,##0} 權重 {wdcf|0%}，其餘為 Comparable ¥{comp|#,##0}），比現價 ¥{px|#,##0} {upw|SP}，投資建議{rec}（級距 {tier}）。把加權目標價與現價的差距拆成兩個來源（各自權重 ×（目標價 − 現價））：DCF 占 {dcfsh|0%}、Comparable 占 {compsh|0%}。Comparable 比現價 {upc|+0%;−0%}，可拆成兩個因素相乘：每股淨值由約 2027 年底成長到約 2028 年底（本模型，情境加權）{bvg|+0%;−0%}，記憶體同業前瞻 P/B 中位數 {pmed|0.00} 倍比現價隱含的約 2027 年底 P/B（{pxbv27|0.00} 倍）{pmdir} {pmg|0%;0%}，{cdrv}。DCF 比現價 {dcfup|+0%;−0%}：{dcfwho}（情境 DCF 價值為現價的 A {dcfpxA|0.00}／B {dcfpxB|0.00}／C {dcfpxC|0.00} 倍）。",
 "重點 2：近一年的 EPS 與共識比較，之後依情境分歧。FY2026 EPS 情境 A ¥{epsA26|#,##0}、B ¥{epsB26|#,##0}、C ¥{epsC26|#,##0}，共識 ¥{ceps26|#,##0}（三者依序{vs26A}、{vs26B}、{vs26C}共識）；FY2027 EPS 情境 A ¥{epsA27|#,##0}、B ¥{epsB27|#,##0}、C ¥{epsC27|#,##0}，共識 ¥{ceps27|#,##0}（依序{vs27A}、{vs27B}、{vs27C}共識）。同一時間窗比較：12 個月後的前瞻 EPS（約 2027-11～2028-10，情境加權）¥{feps|#,##0}，比共識同窗 ¥{cfeps|#,##0} {geps|SP}。",
 "重點 3：營業利益與 EPS 最受 {s_top} 影響。以情境 A 為基準（Sens 頁，一次改一個輸入），「{l_pdn}」「{l_pup}」使 FY2028 EPS 變動 {e_pdn|+0%;−0%}／{e_pup|+0%;−0%}；「{l_fdn}」「{l_fup}」為 {e_fdn|+0%;−0%}／{e_fup|+0%;−0%}；「{l_sdn}」「{l_sup}」為 {e_sdn|+0%;−0%}／{e_sup|+0%;−0%}；「{l_cdn}」「{l_cup}」只有 {e_cdn|SP}／{e_cup|SP}。這些是敏感度測試案例，不是門檻。",
 "重點 4：v0.5 起 DCF 改為標準版（無槓桿自由現金流折現）：DCF 目標價 ¥{dcftp|#,##0}，三種情境推到 12 個月後為 A ¥{dcfA|#,##0}、B ¥{dcfB|#,##0}、C ¥{dcfC|#,##0}（{dcfwho}）。折現率 WACC {wacc|0.0%}：股權成本 {ke|0.0%}＝rf {rfv|0.00%}（日本 10 年期公債）＋ β {beta|0.00} ×（ERP {erp|0%} ＋ 日本國家風險溢酬 {crp|0.00%}）；β 為自算記憶體產業無槓桿 β {betau|0.000}（同業中位數 Blume 調整；Kioxia 上市不足 60 個月，不入中位數）依負債比重新加槓桿，ERP 5% 與 RONIC＝WACC 為假設；永續成長率 g＝MIN(3%, rf)＝{gtv|0.00%}。相對舊法（股東現金流、折現率 10%），現金流口徑（營運資金投入、資本支出、利息不進營業現金流）、終值改用價值驅動式、折現率提高三者都使 DCF 下降（拆解見 README.md 版本說明）。情境 B 的終值現值占企業價值 {tvshB|0%}。以三情境加權目標價為基準，WACC −1pt／＋1pt 使加權目標價變動 {s_wlo|+0.0%;−0.0%}／{s_whi|+0.0%;−0.0%}，β 低／高案例（β_u ∓ 0.3 各自重新加槓桿）為 {s_blo|+0.0%;−0.0%}／{s_bhi|+0.0%;−0.0%}，ERP 4%／6% 為 {s_elo|+0.0%;−0.0%}／{s_ehi|+0.0%;−0.0%}，營運資金比例低／高為 {s_wclo|+0.0%;−0.0%}／{s_wchi|+0.0%;−0.0%}（敏感度測試，不是門檻）。"
 "淨現金以 2026-09-30 公司淨現金（已扣 8 月 ¥800B 買回）減租賃 ¥{dcfnc|#,##0}B 加 stub 期 FCFF（情境 B ¥{stubB|#,##0}B）、扣已宣告中期股利 ¥{stubdiv|#,##0}B，滾到評價日為 ¥{ncrB|#,##0}B。"
 "美元口徑折現（Sens 案例：rf 改用美國 10 年期 {rfus|0.00%}、CRP 照基準、g＝MIN(3%, {rfus|0.00%})，現金流與匯率不變）下加權目標價 ¥{usd|#,##0}，相對基準 {usdp|+0.0%;−0.0%}；基準仍採日本利率＋固定匯率。{tvflag}"]
PRES_ANS = [
 "盈餘由什麼決定？答：{s_top}。以情境 A 為基準的敏感度測試中，「{l_pdn}」「{l_pup}」使 FY2028 EPS 變動 {e_pdn|+0%;−0%}／{e_pup|+0%;−0%}，「{l_fdn}」「{l_fup}」為 {e_fdn|+0%;−0%}／{e_fup|+0%;−0%}（營收以美元計、成本以日圓計），「{l_sdn}」「{l_sup}」為 {e_sdn|+0%;−0%}／{e_sup|+0%;−0%}，「{l_cdn}」「{l_cup}」只有 {e_cdn|SP}／{e_cup|SP}。SSD 與儲存占營收由 FY2025 的 {ssshA25|0%} 到 FY2027 的 {ssshA27|0%}（情境 A；其中 AI 相關約占營收 {aishA27|0%}）；FY2026 起的分部占比是假設，分部拆分不影響合併營收與 EPS。",
 "高價能維持多久？答：依本模型三種情境的價格路徑（情境設定，不是單一預測），FY2026–FY2029 的 EPS 高點在情境 A {peakA}、情境 B {peakB}、情境 C {peakC}。FY2026 三情境相近（A ¥{epsA26|#,##0}、B ¥{epsB26|#,##0}、C ¥{epsC26|#,##0}），分歧從 FY2027 開始：FY2027 為 A ¥{epsA27|#,##0}、B ¥{epsB27|#,##0}、C ¥{epsC27|#,##0}，FY2028 為 A ¥{epsA28|#,##0}、B ¥{epsB28|#,##0}、C ¥{epsC28|#,##0}。情境 C 的 FY2029 毛利率降到 {gmC29|0%}（FY2026 為 {gmC26|0%}）。",
 "現價是否已反映循環見頂？答：依本模型{verdict}。現價 ¥{px|#,##0} 相當於約 2027 年底每股淨值（本模型，情境加權）的 {pxbv27|0.00} 倍，記憶體同業同口徑的前瞻 P/B 中位數 {pmed|0.00} 倍；Comparable 把同業倍數套在 12 個月後的前瞻每股淨值（約 2028 年底，情境加權）¥{fbv|#,##0} 上，得 ¥{comp|#,##0}（比現價 {upc|SP}）。DCF 目標價 ¥{dcftp|#,##0}（比現價 {dcfup|SP}），{dcfwho}。加權目標價 ¥{tpw|#,##0}，比現價 {upw|SP}，投資建議{rec}。市場共識目標價 ¥{cons|#,##0} 隱含約 {cpb|0.00} 倍 P/B（÷ 共識每股淨值），是同業中位數的 {cpbx|0.00} 倍；本模型約 2028 年底每股淨值比共識 {gbv|SP}。"]
RD["A1"] = RD["A1"].value + "；v0.5（2026-11-05）：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數）＋HTML 新版面 v1.0（README 新增展示層與命題回覆）；Comparable、情境權重、DCF 權重 50% 與投資建議規則不變；v0.5.1（2026-11-05）：β 改自算產業 β（KR 半導體，Blume 調整），評價日統一為 2026-12-31；v0.5.2（2026-11-05）：第二輪修正（記憶體組 β、6-30 淨現金以 stub 期滾到評價日、β 敏感度 ±0.3、Sens 美元口徑折現）；v0.5.3（2026-11-06）：第三輪修正（現價與同業報價統一 2026-11-04、資產負債表日 9-30（扣 ¥800B 買回與股數）、版本沿革固定文字）"
_r = RD.max_row + 2
RD.cell(_r, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；v0.5 新增，上方原列保留給下游）").font = BOLD; _r += 1
for t_ in [PRES_LEAD] + PRES_PROPS + PRES_OBS:
    RD.cell(_r, 1, _fx(t_)); _r += 1
_r += 1
RD.cell(_r, 1, "命題回覆（HTML 用；依本模型計算逐一回答上方三個命題）").font = BOLD; _r += 1
for t_ in PRES_ANS:
    RD.cell(_r, 1, _fx(t_)); _r += 1
_r += 1
RD.cell(_r, 1, "v0.5（2026-11-05）：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；tools/fcff_dcf.py；docs/plan/個股DCF標準規格.md v1.0）。新增 DCF 頁（三情境 FCFF 明細與檢查列；評價日 2026-06-30，FY2026 只計評價日後三季）；Val 0.1 節改為折現率與 DCF 彙總（0.2 節以後列號不變）；Sens 目標價參數改為標準案例（WACC、β、ERP、g、RONIC、資本支出、營運資金、DCF 權重、終值替代法）；In 頁新增 M2 節（日本 rf、Damodaran 產業 β 重新加槓桿、ERP＋日本 CRP、資本結構、穩態比例），刪除 tp_r、tp_roe；淨現金收益率鍵 rf 改名 ncy（值不變）。HTML 改用個股 HTML 版面規格 v1.0（README 展示層與命題回覆）。PL、Q、Units、Recon 頁與 v0.4 相同。"); _r += 1
# v0.5.3（第三輪修正第 37 項）：版本沿革列改為當版成品的固定文字（原為活公式，會顯示現值）
RD.cell(_r, 1, 'v0.5.1（2026-11-05）：β 改自算產業 β（Andy 2026-11-05 22:55）：KR 半導體同業 4 家無槓桿 β 中位數 1.72（Yahoo Finance 月報酬 2021-11～2026-10 對韓國綜合指數迴歸，Hamada 去槓桿、現金修正；Kioxia 上市不足 60 個月，以記憶體同業為準），Blume 調整後 1.482，依市值 D/E 重新加槓桿為 1.52，取代 Damodaran Global 半導體表；β 低／高案例改為中位數 × 0.8 再 Blume／中位數未調整。評價日由 2026-06-30 統一為 2026-12-31（FY2026 只計評價日後一季；目標時點 2027-12-31；淨現金沿用 2026-06-30 資產負債表）。新舊對照（v0.5 → v0.5.1）：β 1.92 → 1.52，WACC 13.8% → 11.5%，DCF 目標價 13,501 → 14,800，加權目標價 16,224 → 16,874，上檔空間 +27.2% → +32.3%，投資建議 買進 → 買進。其他模型內容不變。'); _r += 1
RD.cell(_r, 1, 'v0.5.2（2026-11-05）：工作單第二輪修正（Andy 2026-11-05 23:16 定案；共用模組 tools/fcff_dcf.py v2）。(1) β 改「記憶體」組（MU、WDC、SK 海力士、三星電子、南亞科 5 家無槓桿中位數 2.294，Blume 調整 1.867），依市值 D/E 重新加槓桿為 1.92；β 敏感度改 β_u ∓ 0.3 重新加槓桿。(2) 淨現金：資產負債表日 2026-06-30（9-30 資產負債表未取得）¥-47B，以 stub 期滾到評價日 2026-12-31＝＋stub 期 FCFF（FY2026 FCFF × stub_frac；情境 B ¥1,275B）− 已宣告中期股利 ¥49B＝¥1,179B（v0.5.1 直接沿用 6-30 淨現金，未計 7–12 月現金流）；8 月取得完畢的 ¥800B 買回因股數未取得，現金與股數都維持 6-30 口徑。(3) 明確期 FY2026–FY2029 折舊／資本支出照舊讀 PL 頁逐年列（改用模組 v2 da／cx），FY2030 起收斂且資本支出 ≥ 折舊 ×（1＋g）。(4) Sens 新增「美元口徑折現」：加權目標價 ¥16,220（−2.6%）。新舊對照（v0.5.1 → v0.5.2）：β 1.52 → 1.92，WACC 11.5% → 13.7%，DCF 目標價 14,800 → 14,355，加權目標價 16,874 → 16,651，上檔空間 +32.3% → +30.5%，投資建議 買進 → 買進。'); _r += 1
RD.cell(_r, 1, 'v0.5.3（2026-11-06）：第三輪（合併前）修正（共用模組 tools/fcff_dcf.py v2 不變）。(1) 現價改 2026-11-04 收盤 ¥10,735（Yahoo Finance 日線收盤；v0.5.2 為 10-23 的 ¥12,755）；同業報價統一 11-04 收盤：前瞻 P/B 美光 3.35 → 3.04、SK 海力士 1.35 → 1.48、三星 1.26 → 1.27、SanDisk 2.86 → 2.39，中位數 2.10 → 1.93；Kioxia 自身 1.37（與記憶體三家、SanDisk 模型同一 11-04 口徑）；前瞻 P/E 並列參考同步。(2) 資產負債表日 6-30 → 9-30：DCF 淨現金改用公司公布的 9-30 淨現金 ¥812.7B（已扣 8 月執行完畢的 ¥800B 買回；9-30 現金、借款分列未取得，借款與租賃沿用 6-30），扣 9-30 後支付的中間配當 ¥47.9B；評價用股數扣買回減少的 47.4 百萬股（分割後；由中間配當總額 ÷ 每股 ¥90 推得 9-30 流通股數）。PL 頁 FY2026 權益、淨現金同扣 ¥800B、股數同減。(3) README 版本沿革列改為當版固定文字；README 重點觀察 1 的主因改公式判定（上檔來源占比、每股淨值成長 vs 倍數差）；情境驅動對照收斂法列名中性化、新增 cyc_mB（＝1）取代情境 B 佔位 0；HTML 情境驅動表 % 列改百分比、DCF 明細只有比例列用 %、橫幅日圓取整數。新舊對照（v0.5.2 → v0.5.3）：DCF 目標價 14,355 → 14,833，Comparable 18,948 → 17,409，加權目標價 16,651 → 16,121，上檔空間 +30.5% → +50.2%，投資建議 買進 → 強力買進（上檔空間略高於 +50% 門檻）。'); _r += 1
# v0.5.3：README!A4（重點觀察 1）主因改為公式判定（上檔來源占比、每股淨值成長 vs 倍數差；同一時間窗）
assert str(RD["A4"].value).startswith('="1. 加權目標價')
RD["A4"] = _fx("1. 加權目標價 ¥{tpw|#,##0}，相對現價 ¥{px|#,##0}（2026-11-04 收盤）{upw|SP}。{src}：加權目標價與現價的差距中 DCF 占 {dcfsh|0%}、Comparable 占 {compsh|0%}（各自權重 ×（目標價 − 現價））。"
               "Comparable ¥{comp|#,##0}（{upc|+0%;−0%}）拆成兩個因素相乘（同一時間窗）：每股淨值由約 2027 年底成長到約 2028 年底（本模型，情境加權）{bvg|+0%;−0%}，同業前瞻 P/B 中位數 {pmed|0.00} 倍比現價隱含的約 2027 年底 P/B（{pxbv27|0.00} 倍）{pmg|+0%;−0%}，{cdrv}。"
               "DCF（標準版 FCFF＋WACC）¥{dcftp|#,##0}（{dcfup|+0%;−0%}）。")

wb.save(sys.argv[1])
RULER = [i for i in range(nr, tr) if isinstance(V.cell(i, 1).value, str) and V.cell(i, 1).value.startswith("規則值：")]
RC_FMT = {}
for i in range(4, RC_LAST + 1):
    lab = str(RC.cell(i, 1).value or "")
    if "毛利率" in lab: RC_FMT[str(i)] = {c: "p1" for c in "BCEG"}
    if "合資群份額" in lab: RC_FMT[str(i)] = {"C": "p1", "E": "p1", "G": "p1"}
    if "口徑係數" in lab: RC_FMT[str(i)] = {"C": "n3"}
EXTRA = dict(
    DRIVERS=[("g2_rev", "n0", ["FY2026 營收（情境 A）：", f"@PL!D{SUMR[('A', 'yrev')]}|n0"], "公司第二季指引決定 Kioxia 在市場層的份額起點；之後營收隨市場層價格與位元同向"),
             ("fx", "n0", ["FY2027 營收（情境 A）：", f"@PL!E{SUMR[('A', 'yrev')]}|n0"], "日圓貶值使營收（美元計價）上升、成本（日圓）不變，毛利率同向"),
             ("cd_A", "p0", ["FY2028 毛利率（情境 A）：", f"@PL!F{SUMR[('A', 'ygm')]}|p1"], "每位元成本下降使毛利率上升（同向；與 SanDisk 同值）"),
             ("tax", "p0", ["FY2027 淨利（情境 A）：", f"@PL!E{SUMR[('A', 'yni')]}|n0"], "稅率上升使淨利下降（反向）"),
             ("cap_2027", "n0", ["FY2027 自由現金流（情境 A）：", f"@PL!E{SUMR[('A', 'fcf')]}|n0"], "資本支出上升使自由現金流與 DCF 下降（反向）"),
             ("payout", "p0", ["FY2028 EPS（情境 A）：", f"@PL!F{SUMR[('A', 'eps')]}|n0"], "回饋以買回計，減少股數、提高 EPS")],
    QUESTION="NAND 合約價的高檔能維持多久、日圓匯率如何影響盈餘，以及現價隱含的低倍數是否已反映循環見頂",
    PL_INTRO="FY2025 為實績；FY2026 由第一季實績、第二季指引與第三、四季推導組成；FY2027 起由 NAND 市場層驅動（會計年度四月至隔年三月）。",
    TITLER=TITLER, RULEH=RULER[0] - 1, RULER=RULER, RC_FMT=RC_FMT, TORNADO="FY2028 EPS 相對基準的變動（情境 A，一次動一個輸入）")
if os.environ.get("KIOXIA_ROWS"):
    json.dump(dict(**EXTRA, IN=IN_ROW, SUMR={f"{a}|{b}": v for (a, b), v in SUMR.items()}, PB=PB, EXT=EXT, T=T, EPSR=EPSR, CFR=CFR, PL_CHK=PL_CHK, CUR=CUR, CMH=CMH,
                   SENS_LAST=SENS_LAST, RC_LAST=RC_LAST, NR=nr, TR=tr, PB0=PB0, UR={f"{a}|{b}": v for (a, b), v in UR.items()}, RD=README_ROWS, EXTC=EXTC, P0=p0),
              open(os.environ["KIOXIA_ROWS"], "w"), ensure_ascii=False)
print("saved", sys.argv[1])

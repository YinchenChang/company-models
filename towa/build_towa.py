# TOWA（6315.T）收支與估值模型 v0.1（AI 半導體 Project 步驟 5：設備與材料——日本後段設備，封膠）
# 用法：python3 towa/build_towa.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度 4 月到隔年 3 月：FY2025＝2025-04～2026-03（公司「2026年3月期」）。金額：十億日圓（¥B）。
# 結構比照 disco/build_disco.py：README、In、Q、Hist、PL、Val、Recon、Sens、Sources。
# 情境與行業一致（Andy 2026-10-25 13:13）：A／B／C 讀同字母的上游情境（HBM 位元：skhynix Mkt；CoWoS：tsmc 顆數係數＝markets/memory SCEN_UNITS；
# 設備支出：markets/equipment Mkt 頁 A／B／C）；公司特有變數（份額流失、增量利潤率調整）A 最有利、C 最不利。
import sys, os, glob, json
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
import market as MEM
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC"); FADE = PatternFill("solid", fgColor="E7E6E6")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
SCN = {"A": "順利", "B": "平平", "C": "不如預期"}
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- 讀入其他模型的成品（唯讀；不修改它們） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
TSMC_F = only("tsmc/dist/*.xlsx"); SKH_F = only("skhynix/dist/*.xlsx"); EQ_F = only("markets/equipment/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)

ws_ = openpyxl.load_workbook(SKH_F, data_only=True)["Mkt"]
SHIP = {}
for r in ws_.iter_rows():
    lab = str(r[0].value or "")
    if lab.startswith("出貨＝MIN"):
        SHIP[lab[-2]] = ([r[2 + j].value for j in range(5)], f"{rel(SKH_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert sorted(SHIP) == ["A", "B", "C"] and all(v is not None for s_ in SHIP for v in SHIP[s_][0])

wbt = openpyxl.load_workbook(TSMC_F, data_only=True)
COW = None
for r in wbt["CoWoS"].iter_rows():
    if r[0].value == "需求合計":
        COW = ([r[1 + j].value for j in range(4)], f"{rel(TSMC_F)}｜CoWoS!B{r[0].row}:E{r[0].row}")
assert COW and all(v is not None for v in COW[0])
UMSRC = {}
for r in wbt["In"].iter_rows(min_row=4):   # 台積電情境顆數係數必須等於 markets/memory 的 SCEN_UNITS（單一來源）
    k = r[0].value
    if k and k[:4] in ("um28", "um29"):
        assert abs(r[4].value - MEM.SCEN_UNITS[k[-1]][0 if k[:4] == "um28" else 1]) < 1e-12, k
        UMSRC[k] = f"{rel(TSMC_F)}｜In!E{r[0].row}"
assert len(UMSRC) == 6

ws_ = openpyxl.load_workbook(EQ_F, data_only=True)["Mkt"]
PKG = None; EQT = {}; EQAI = {}; sc_ = None
for r in ws_.iter_rows(min_row=4, max_row=40):
    a = r[0].value
    if isinstance(a, str) and a.startswith("情境 ") and len(a.strip()) == 4:
        sc_ = a.strip()[-1]; continue
    if sc_ and a == "設備總額" and sc_ not in EQT:
        EQT[sc_] = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
    if sc_ and isinstance(a, str) and a.strip() == "AI 相關設備" and sc_ not in EQAI:
        EQAI[sc_] = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
for r in ws_.iter_rows():
    if r[0].value == "SEMI 組裝與封裝設備":
        PKG = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert PKG and sorted(EQT) == list(SC) and sorted(EQAI) == list(SC)
PEAK, PEAK_SRC = CP.peak_years(ROOT)

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億日圓（¥B）；HBM 位元：EB；CoWoS：千片；設備市場：十億美元。每個輸入附標記與來源；找不到者寫「未找到」，公司未揭露者寫「公司不揭露」"
IN["A2"] = "標記：Verified（公司文件、交易所資料）／Interested-party（寫明是誰、立場）／Analogy（寫明對照與差異）／Assumed（一律給低／高區間）／Derived（寫明換算）。SemiAnalysis 未使用。公司原文（JPX 決算短信 iXBRL、towajapan.co.jp 說明資料 PDF）以網頁擷取工具讀取；圖表數字擷取錯位處未採用，待人工複核"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
FR26 = "公司 2026年3月期 決算短信（2026-04-22 登錄，JPX iXBRL 081220260422508001）"
FR27Q1 = "公司 2027年3月期 第1四半期 決算短信（2026-08-06，JPX iXBRL 081220260702586673）"
PR26 = "公司 2026年3月期 決算說明資料（2026-05，2025FY_end_presentation.pdf）"
PR26H1 = "公司 2026年3月期 第2四半期 決算說明會資料（2025-11-07）"
MATSUI = "松井證券 決算情報（finance.matsui.co.jp/stock/6315/settlement）"

sec("A. FY2025 實績（2025-04～2026-03）與事業別營收（校準錨）")
inp("a25_rev", "營收 FY2025", "¥B", None, 54.365, None, "Verified", FR26 + "：54,365 百萬日圓（+1.7%）", MATSUI + "：54,365")
inp("a25_gp", "毛利 FY2025", "¥B", None, 18.372, None, "Verified", "kabutore 財務分析（轉載公司損益表）：売上総利益 18,371,561 千日圓（毛利率 33.8%）", "edinetdb（2026-10-24 讀取）：毛利率 33.8%。決算短信摘要頁未列毛利")
inp("a25_oi", "營業利益 FY2025", "¥B", None, 6.917, None, "Verified", FR26 + "：6,917（−22.1%；營益率 12.7%）", MATSUI + "：6,917")
inp("a25_ni", "淨利（歸屬母公司）FY2025", "¥B", None, 4.593, None, "Verified", FR26 + "：4,593（−43.4%）", MATSUI + "：4,593")
inp("a25_eps", "基本 EPS FY2025", "¥", None, 61.22, None, "Verified", FR26 + "：61.22", MATSUI + "：61.22")
inp("seg_semi", "半導體事業營收 FY2025", "¥B", None, 40.39, None, "Verified", PR26 + "：403.9 億日圓（+2.2%）", "tousaiblog（2026-05-12）：403.9")
inp("seg_mold", "　其中：封膠設備與模具（モールディング装置・金型）", "¥B", None, 36.46, None, "Verified", PR26 + "：364.6 億日圓（擷取文字；版面待複核）", "第二來源：未找到。設備與模具未分開揭露（公司不揭露）")
inp("seg_osemi", "　其中：其他半導體（切割 singulation 等）", "¥B", None, "=seg_semi-seg_mold", None, "Derived", "403.9 − 364.6 ＝ 39.3", "公司：記憶體投資使切割設備大增（tousaiblog 摘要，無金額）")
inp("seg_new", "新事業營收 FY2025（含 TSS：售後、翻新、二手機、零件）", "¥B", None, 9.47, None, "Verified", PR26 + "：94.7 億日圓（+0.6%）", "tousaiblog：94.7")
inp("seg_med", "醫療器材事業營收 FY2025（BANDIC，精密塑膠成形品）", "¥B", None, 2.48, None, "Verified", PR26 + "：24.8 億日圓（+9.9%）", "tousaiblog：24.8")
inp("seg_las", "雷射事業營收 FY2025", "¥B", None, 2.00, None, "Verified", PR26 + "：20.0 億日圓（−11.0%）", "tousaiblog：20.0")
inp("seg_adj", "事業合計與營收的差額（取整與調整）", "¥B", None, "=a25_rev-seg_semi-seg_new-seg_med-seg_las", None, "Derived", "營收 − 四事業合計", "FY2025 以後設為 0")
inp("ord_25", "受注高 FY2025", "¥B", None, 59.56, None, "Verified", PR26 + "：595.6 億日圓（+25.6%，史上第二高）", "tousaiblog：595.6")
inp("bl_25", "受注殘高 2026-03 末", "¥B", None, 30.14, None, "Verified", PR26 + "：301.4 億日圓", "2025-09 末 267.4（" + PR26H1 + "）")
inp("sh", "TOWA 封膠設備全球市占（2024 年，TechInsights）", "%", 0.60, 0.648, 0.70, "Interested-party", PR26H1 + " 第 14 頁、" + PR26 + "：TechInsights 數據，世界市占 No.1；圓餅最大塊 64.8%（標籤 A–D 未明示哪塊是公司）", "本專案候選表：市占約 2/3（markets/invest_map universe，非獨立來源）。只透過「市占流失」進入模型；FY2025 水準已含在校準中")

sec("B. AI 錨點（校準用；公司不揭露 AI 或 HBM 金額）")
inp("ai_25", "FY2025 AI 相關封膠設備與模具營收（HBM＋CoWoS 等 2.5D／晶圓級）", "¥B", 7.0, 10.0, 13.0, "Assumed", "公司不揭露。推估：說明資料圓餅「移轉成形 59%／壓縮成形 41%」（標籤對應未確認）× 封膠 36.5 ≈ 壓縮成形 15.0；AI（HBM、AI 加速器、先進封裝）約占壓縮成形三分之二 → 約 10.0（約占營收 18%）", "第二來源：未找到。公司說法（Interested-party）：壓縮成形占比上升改善組合；AI／資料中心帶動第二季後的接單。低／高 ±30%")
inp("w_hbm", "AI 相關中 HBM 比重 FY2025（其餘為 CoWoS 等 2.5D／晶圓級封膠）", "%", 0.35, 0.50, 0.65, "Assumed", PR26H1 + "：HBM 封膠設備（含同業）累計約 30 台，數年內需再增 30–45 台；已取得 HBM4 客戶認證", "比重未找到；取一半。HBM 與 CoWoS 的成長形狀不同，Sens 列 35%／65%")

sec("C. HBM 封膠需求（skhynix 成品 Mkt 頁，情境 A／B／C 各一列；唯讀）")
for s in SC:
    for j, y in enumerate((25, 26, 27, 28, 29)):
        inp(f"hs{s}_{y}", f"HBM 出貨位元 20{y}（EB；{s}）", "EB", None, SHIP[s][0][j], None, "Derived", SHIP[s][1], "markets/memory：出貨＝MIN(需求 × 情境顆數係數, 供給)；2028–29 隨同字母情境" if j == 0 else "同上")
for y, (lo, b, hi) in zip((25, 26, 27, 28, 29), ((22, 24, 24), (24, 24, 26), (24, 26, 28), (24, 28, 32), (26, 30, 32))):
    inp(f"den_{y}", f"DRAM 核心晶粒容量（組合平均）20{y}", "Gb／顆", lo, b, hi, "Assumed", "沿用 disco/ 的假設（JEDEC HBM4：24Gb 或 32Gb 晶粒，TrendForce 2025-04-28）", "組合比例未找到" if y == 25 else "同上")
for y, (lo, b, hi) in zip((25, 26, 27, 28, 29), ((8, 10, 12), (10, 12, 12), (12, 12.5, 14), (12, 14, 16), (12, 15, 16))):
    inp(f"stk_{y}", f"HBM 平均堆疊層數 20{y}", "層", lo, b, hi, "Assumed", "沿用 disco/ 的假設（JEDEC HBM4：4／8／12／16 層）", "封膠以「堆疊」為單位：每個 HBM 堆疊封一次；層數愈高、同位元的堆疊數愈少" if y == 25 else "同上")
inp("dlt", "已裝機封膠產能每年汰換／升級比例（新機需求＝下一年產能增量＋汰換）", "%/年", 0.05, 0.10, 0.15, "Assumed", "公司：高性能化的汰換需求（" + PR26H1 + "，無數字）", "與校準係數一起只能識別乘積的形狀；Sens 列 5%／15%")
inp("g_tail", "CY2030–31 HBM 堆疊數與 CoWoS 片數年增（上游只到 2029；三情境同值）", "%", 0.0, 0.10, 0.20, "Assumed", "無：上游只到 2029；設備比產能早一年買，FY2029 需要 CY2030–31", "三情境同值，不自訂行業路徑")

sec("D. CoWoS 等 2.5D 封膠需求：台積電模型 CoWoS 需求（唯讀）")
inp("cw_25", "CoWoS 需求 CY2025", "千片", 600, 670, 700, "Interested-party", "36Kr／半導體產業縱橫（2025-12-04）：2025 全球 CoWoS 需求約 67 萬片（disco/ 同值）", "TrendForce 月產能 2024 年底 3.5–4 萬、2025 年底 7.5–8 萬（Derived 年均約 66 萬）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"cw_{y}", f"CoWoS 需求 CY20{y}（台積電模型，情境 A 顆數）", "千片", None, COW[0][j], None, "Derived", COW[1], "B、C 的 2028–29 乘以同字母情境顆數係數（下列 um）" if j == 0 else "同上")
for s in SC:
    for n, y in enumerate((28, 29)):
        inp(f"um{y}_{s}", f"情境 {s}：20{y} 顆數 ÷ 基準（CoWoS 片數同比例）", "x", None, MEM.SCEN_UNITS[s][n], None, "Derived", UMSRC[f"um{y}_{s}"], "markets/memory SCEN_UNITS（單一來源；建置時檢查與台積電 In 頁相同）" if (s == "A" and y == 28) else "同上")

sec("E. 非 AI 後段需求：SEMI 組裝與封裝設備 × 同字母情境的非 AI 設備比例（markets/equipment，唯讀）")
for j, y in enumerate((25, 26, 27, 28, 29)):
    inp(f"pk_{y}", f"SEMI 組裝與封裝設備 CY20{y}（情境 A 路徑）", "US$B", None, PKG[0][j], None, "Derived", PKG[1], "SEMI 2026-07；2027 幾何內插（markets/equipment）" if j == 0 else "同上")
for s in SC:
    for j, y in enumerate((25, 26, 27, 28, 29)):
        inp(f"eqt{s}_{y}", f"設備總額 CY20{y}（情境 {s}）", "US$B", None, EQT[s][0][j], None, "Derived", EQT[s][1], "markets/equipment Mkt 頁「設備總額」" if j == 0 else "同上")
    for j, y in enumerate((25, 26, 27, 28, 29)):
        inp(f"eqa{s}_{y}", f"AI 相關設備 CY20{y}（情境 {s}）", "US$B", None, EQAI[s][0][j], None, "Derived", EQAI[s][1], "非 AI 設備＝設備總額 − AI 相關設備；情境 s 的封裝設備＝SEMI 封裝 ×（非 AI 設備 s ÷ 非 AI 設備 A）" if j == 0 else "同上")
inp("g_pk_tail", "CY2030 非 AI 封裝設備與設備總額年增（上游只到 2029；三情境同值）", "%", 0.0, 0.05, 0.10, "Assumed", "無：上游只到 2029", "三情境同值")

sec("F. 其他事業（Analogy：公司自己的歷史與計畫）")
inp("g_tss", "新事業（TSS 售後等）年增", "%", 0.0, 0.03, 0.06, "Analogy", "FY2025 +0.6%、公司 FY2026 計畫 +3.1%（" + PR26 + "）", "售後隨裝機量，但含二手機與翻新（波動大）")
inp("g_med", "醫療器材事業年增", "%", 0.0, 0.03, 0.06, "Analogy", "FY2025 +9.9%、公司 FY2026 計畫 +2.5%", "與半導體循環無關")

sec("G. 公司特有變數（隨情境；方向：A 最有利、C 最不利）")
for s, (lo, b, hi) in zip(SC, ((0.0, 0.0, 0.03), (0.0, 0.03, 0.06), (0.04, 0.08, 0.12))):
    inp(f"shl_{s}", f"情境 {s}：FY2029 AI 封膠市占流失（相對 2025，線性；pt）", "pt", lo, b, hi, "Assumed", "競爭：ASMPT、Yamada（Apic Yamada）、韓國與中國本土廠的壓縮成形機；公司稱 HBM 客戶基礎擴大（Interested-party）" if s == "A" else ("B：小幅流失" if s == "B" else "C：HBM／CoWoS 降速時價格競爭加劇"), "只作用在 AI 封膠（校準已含 FY2025 水準）")
for s, (lo, b, hi) in zip(SC, ((0.01, 0.03, 0.05), (-0.02, 0.0, 0.02), (-0.08, -0.05, -0.02))):
    inp(f"ima_{s}", f"情境 {s}：增量營益率調整（相對 Hist 迴歸斜率；pt）", "pt", lo, b, hi, "Assumed", "A：壓縮成形占比上升、首台成本不再（公司說法）" if s == "A" else ("B：迴歸斜率" if s == "B" else "C：低毛利案件與首台成本延續（FY2025 營益率 12.7% 的原因）"), "增量營益率＝營業利益變動 ÷ 營收變動")

sec("H. 利潤率（由 Hist 頁 FY2016–FY2025 推得，不寫死）")
inp("im", "增量營益率（營業利益 ＝ 截距 ＋ 斜率 × 營收 的斜率）", "x", "=SLOPE(Hist!E6:E15,Hist!B6:B15)-0.05", "=SLOPE(Hist!E6:E15,Hist!B6:B15)", "=SLOPE(Hist!E6:E15,Hist!B6:B15)+0.05", "Derived", "Hist 頁 10 年（FY2016–FY2025）最小平方法；低／高 ±5pt（Assumed）", "FY2027 起：營業利益＝FY2026 錨點營業利益＋（斜率＋情境調整）×（營收 − FY2026 營收）")
inp("nioi", "淨利 ÷ 營業利益（FY2027 起）", "x", 0.65, "=g26_ni/g26_oi", 0.72, "Derived", "基準＝公司 FY2026 指引 7,000 ÷ 10,240 ＝ 0.684", "FY2025 0.664（4,593 ÷ 6,917）；FY2024 0.915（含特別利益）；低 0.65、高 0.72（Assumed）")
inp("shares", "加權平均股數", "百萬股", None, 75.029892, None, "Verified", FR27Q1 + "：第一季期中平均 75,029,892 股（不含 J-ESOP 信託持股）", "FY2025 平均 75,020,506（" + FR26 + "）")

sec("I. FY2026 實績、公司指引與對帳參考")
inp("q1_rev", "FY2026 第一季營收（2026-04～06）", "¥B", None, 16.183, None, "Verified", FR27Q1 + "：16,183（+100.3%）", MATSUI + "：16,183")
inp("q1_oi", "FY2026 第一季營業利益", "¥B", None, 2.212, None, "Verified", FR27Q1 + "：2,212（前年同期 −581）", MATSUI)
inp("q1_ni", "FY2026 第一季淨利", "¥B", None, 1.657, None, "Verified", FR27Q1 + "：1,657", MATSUI)
inp("q1_eps", "FY2026 第一季 EPS", "¥", None, 22.09, None, "Verified", FR27Q1 + "：22.09", MATSUI)
inp("h1g_rev", "公司指引：FY2026 上半營收", "¥B", None, 32.0, None, "Interested-party", FR27Q1 + "：32,000（未修正）", FR26 + " 同值")
inp("h1g_oi", "公司指引：FY2026 上半營業利益", "¥B", None, 5.12, None, "Interested-party", FR27Q1 + "：5,120", "同上")
inp("g26_rev", "公司指引：FY2026 營收", "¥B", None, 64.0, None, "Interested-party", FR27Q1 + "：64,000（+17.7%；2026-08-06 未修正）", PR26 + "：640.0 億日圓")
inp("g26_oi", "公司指引：FY2026 營業利益", "¥B", None, 10.24, None, "Interested-party", FR27Q1 + "：10,240（營益率 16.0%）", PR26)
inp("g26_ni", "公司指引：FY2026 淨利", "¥B", None, 7.0, None, "Interested-party", FR27Q1 + "：7,000；EPS 93.30", PR26)
inp("g26_eps", "公司指引：FY2026 EPS", "¥", None, 93.30, None, "Interested-party", FR27Q1 + "：93.30", "kabutan 同值")
inp("g26_semi", "公司計畫：FY2026 半導體事業營收", "¥B", None, 48.43, None, "Interested-party", PR26 + "：484.3（+19.9%）", "醫療 25.5、新事業 97.7、雷射 32.5（+61.9%）")
inp("g26_las", "公司計畫：FY2026 雷射事業營收", "¥B", None, 3.25, None, "Interested-party", PR26 + "：32.5", "")
inp("g_anc", "FY2026 錨點相對公司全年指引的調整", "%", -0.05, 0.0, 0.08, "Assumed", "基準＝公司指引（未修正）；第一季營收 16.2 已達全年指引 25.3%、營業利益 21.6%（kabuyoho 進度）", "高 +8%：接近共識 69.3 的一半距離；低 −5%：第三季記憶體投資延後的前例（FY2025 Q3 下修）")
inp("cons_rev26", "共識營收 FY2026（2027年3月期）", "¥B", 65.3, 69.3, 74.3, "Interested-party", "stockanalysis.com（S&P Global，6 位分析師，2026-09-17 更新）：平均 69.3、區間 65.3–74.3", "第二來源：未找到")
inp("cons_oi26", "共識營業利益 FY2026", "¥B", None, 12.25, None, "Interested-party", "stockanalysis.com（同上）：12.25", "")
inp("cons_ni26", "共識淨利 FY2026", "¥B", None, 8.70, None, "Interested-party", "stockanalysis.com（同上）：8.70 → EPS 約 116（÷ 75.03 百萬股，Derived）", "共識 EPS：stockanalysis 未顯示（付費）")
for s in SC:
    inp(f"pk_{s}", f"情境 {s}：設備市場見頂年度（日曆年；循環位置規則用）", "年", None, PEAK[s], None, "Derived", PEAK_SRC[s], "設備股共用規則（tools/cycle_position.py）" if s == "A" else "同上")
inp("px", "股價（2026-10-23 收盤）", "¥", None, 1968, None, "Verified", "株探 日足（s.kabutan.jp，2026-10-23 終值 1,968）", "kabutore（同值）。2026-10-09 為 2,650，兩週內 −26%")
inp("sa_fpe", "遠期 P/E（stockanalysis，2026-09-25）", "x", None, 17.54, None, "Interested-party", "stockanalysis.com statistics（股價 2,193）", "")
inp("sa_ttm", "TTM P/E（stockanalysis，2026-09-25）", "x", None, 24.07, None, "Interested-party", "stockanalysis.com（TTM EPS 90.36）", "")
sec("J. 目標價情境權重（Val 頁「0. 目標價」用；Andy 可改）")
for s_, w_ in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s_}", f"目標價情境權重 {s_}", "%", None, w_, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%")

IN_ROW = {}
r = 4
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, s1, s2 = srow
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): IN.cell(r, j).font = BLUE
    if key.startswith("tp_w"): IN.cell(r, 5).fill = KEY
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    defname(key, f"In!$E${r}"); IN_ROW[key] = r
    if lo is not None: defname(key + "_lo", f"In!$D${r}")
    if hi is not None: defname(key + "_hi", f"In!$F${r}")
    r += 1
# J 節續：估值 v2 參數（黃底）
def inrow(key, item, unit, base, tag, s1, s2, fmt=None, yellow=False):
    global r
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    c = IN.cell(r, 5, base)
    if not (isinstance(base, str) and base.startswith("=")): c.font = BLUE
    if fmt: c.number_format = fmt
    if yellow: c.fill = KEY
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    defname(key, f"In!$E${r}"); IN_ROW[key] = r; r += 1
inrow("tp_r", "折現率（股東權益要求報酬；美、台、日、韓同一值）", "%", 0.10, "Assumed", "Andy 可改（工作單 2026-10-25 預設 10%）", "TOWA 市值約 1,500 億日圓、營收波動大，風險可能高於大型設備股；依工作單不改值", "0.0%", True)
inrow("tp_g", "長期成長率（永續成長；收斂法 FY2032 營收年增的基準）", "%", 0.04, "Assumed", "Andy 可改（工作單預設 4%）", "情境 A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt（PL 頁延伸區塊）", "0.0%", True)
inrow("tp_roe", "ROE（近三年平均，2024-03～2026-03 期）", "%", "=AVERAGE(0.1215,0.1355,0.0696)", "Derived", "stockanalysis 比率頁（2026-10-25 讀取）：2024-03 期 12.15%、2025-03 期 13.55%、2026-03 期 6.96%", "決算短信：FY2025 7.0%、FY2024 13.6%（兩來源一致）", "0.0%")
inrow("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", 0.5, "Assumed", "Andy 可改（2026-10-25 11:53：DCF／Comparable 各半）", "—", "0%", True)
inrow("nim10", "十年平均淨利率（FY2016–FY2025，各年淨利 ÷ 營收的平均）", "%", "=AVERAGE(Hist!H6:H15)", "Derived", "Hist 頁（公司決算短信、四半期報告書、the-shashi 轉載有價證券報告書）", "收斂法 B、C 情境的淨利率目標", "0.0%")
inrow("cons_tp", "市場共識目標價（賣方平均）", "¥", 3300, "Interested-party", "stockanalysis：6 位分析師平均 ¥3,300（中位數 3,300；區間 2,600–4,000），2026-09-17 更新", "近期：SBI 4,150（9-17）、岡三 2,600（8-28）、Morgan Stanley 3,200（6-11）、Jefferies 4,000（5-12）；早於 10 月股價下跌", "#,##0")
inrow("cons_feps", "共識前瞻 EPS（差距拆解用）", "¥", "=cons_ni26*1000/shares", "Interested-party", "共識前瞻 EPS＝stockanalysis FY2026（2027-03 期）共識淨利 8.70 ÷ 股數（約 ¥116）；比本模型 12 個月後前瞻 EPS 早約一年半，EPS 因素因此偏大", "第二來源：未找到", "#,##0")
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 80, 80]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

# ---------------- Hist：FY2016–FY2025 年度實績 ----------------
H = wb.create_sheet("Hist")
H["A1"] = "Hist — 年度實績（十億日圓；EPS 為日圓）。來源：公司決算短信（JPX）、四半期報告書（EDINET）、松井證券（FY2021 起）、the-shashi（轉載有價證券報告書）。用途：增量營益率迴歸、十年平均淨利率、10 年 P/E 的 EPS"
H["A2"] = "股票分割：2024-10-01 一股分三股（基準日 2024-09-30；tipranks 轉載公司公告；Yahoo Finance 分割紀錄 2024-09-27）。FY2023（2024年3月期）以前公司公布的 EPS 除以 3；FY2021–FY2023 的松井、kabutan 數字已調整"
hdr(H, 5, ["會計年度", "營收", "毛利", "銷管費", "營業利益", "淨利", "EPS（公司公布）", "淨利率", "分割係數", "EPS（分割調整後）", "營業利益率", "來源1", "來源2"])
HIST = [("FY2016（2017-03）", 27.632, None, 3.831, 3.867, 154.64, 3, "決算短信 2018-05-10（前期欄）", "the-shashi：38.3／38.7"),
        ("FY2017（2018-03）", 31.010, None, 3.682, 3.026, 121.02, 3, "決算短信 2018-05-10（JPX 140120180422416748）", "決算短信 2019-05（前期欄）同值"),
        ("FY2018（2019-03）", 28.272, 6.986, 0.937, 0.877, 35.09, 3, "決算短信 2019-05（JPX 140120190511422405）", "決算短信 2020-05-07（前期欄，毛利 6,986）"),
        ("FY2019（2020-03）", 25.255, 6.588, 0.812, 0.368, 14.75, 3, "決算短信 2020-05-07（JPX 140120200507404921）", "the-shashi：8.1／3.7"),
        ("FY2020（2021-03）", 29.707, None, 3.62, 2.663, 106.49, 3, "四半期報告書 第44期第3四半期（EDINET S100NDD0，前期通期欄：營收 29,707、淨利 2,663、EPS 106.49）", "營業利益：the-shashi 36.2 億（單一來源；營收以外未見第二來源）"),
        ("FY2021（2022-03）", 50.666, None, 11.505, 8.129, 108.36, 1, MATSUI + "（分割調整後）", "the-shashi：115／81.3"),
        ("FY2022（2023-03）", 53.822, None, 10.037, 7.346, 97.90, 1, MATSUI, "kabutan：同值"),
        ("FY2023（2024-03）", 50.471, None, 8.661, 6.444, 85.90, 1, MATSUI, "kabutan：同值"),
        ("FY2024（2025-03）", 53.479, None, 8.880, 8.121, 108.28, 1, FR26 + "（前期欄）", MATSUI),
        ("FY2025（2026-03）", 54.365, 18.372, 6.917, 4.593, 61.22, 1, FR26, MATSUI)]
for i, (fy, rv, gp, oi, ni, eps, f, s1, s2) in enumerate(HIST, 6):
    H.cell(i, 1, fy)
    for j, v in ((2, rv), (3, gp), (5, oi), (6, ni), (7, eps)):
        if v is None: continue
        H.cell(i, j, v).font = BLUE; H.cell(i, j).number_format = "#,##0.000" if j < 7 else "#,##0.00"
    if gp is not None: H.cell(i, 4, f"=C{i}-E{i}").number_format = "#,##0.000"
    H.cell(i, 8, f"=F{i}/B{i}").number_format = "0.0%"
    H.cell(i, 9, f).font = BLUE; H.cell(i, 10, f"=G{i}/I{i}").number_format = "#,##0.00"
    H.cell(i, 11, f"=E{i}/B{i}").number_format = "0.0%"
    H.cell(i, 12, s1); H.cell(i, 13, s2)
H.cell(17, 1, "檢查：FY2025 列 − In 頁 a25_*（應為 0）"); H["B17"] = "=ABS(B15-a25_rev)+ABS(C15-a25_gp)+ABS(E15-a25_oi)+ABS(F15-a25_ni)+ABS(G15-a25_eps)"; H["B17"].number_format = "0.000"
H.cell(19, 1, "迴歸與統計").font = BOLD
for i, (lab, f, fm) in enumerate((("營業利益 ＝ 截距 ＋ 斜率 × 營收：斜率（增量營益率）", "=SLOPE(E6:E15,B6:B15)", "0.000"), ("截距 ¥B", "=INTERCEPT(E6:E15,B6:B15)", "0.00"), ("R²", "=RSQ(E6:E15,B6:B15)", "0.000"),
                                  ("FY2025 營業利益：迴歸值 − 實績（¥B）", "=INTERCEPT(E6:E15,B6:B15)+SLOPE(E6:E15,B6:B15)*B15-E15", "0.00"),
                                  ("十年平均淨利率", "=AVERAGE(H6:H15)", "0.0%"), ("十年平均營業利益率", "=AVERAGE(K6:K15)", "0.0%"),
                                  ("FY2021–FY2025 平均營業利益率（AI 前後的高營收期）", "=AVERAGE(K11:K15)", "0.0%"),
                                  ("毛利率 FY2018／FY2019／FY2025（其餘年度未找到）", "=C8/B8", "0.0%")), 20):
    H.cell(i, 1, lab); H.cell(i, 2, f).number_format = fm
H["C27"] = "=C9/B9"; H["D27"] = "=C15/B15"
for c in "CD": H[f"{c}27"].number_format = "0.0%"
H.column_dimensions["A"].width = 50
for c in "BCDEFGHIJK": H.column_dimensions[c].width = 12
H.column_dimensions["L"].width = 70; H.column_dimensions["M"].width = 50

# ---------------- Q：季度實績與指引 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 季度實績（十億日圓；EPS 為日圓）。會計年度 4–3 月。來源1：松井證券季度表（百萬日圓）；來源2：公司決算短信與說明資料"
hdr(Q, 3, ["項目", "FY25 Q1", "FY25 Q2", "FY25 Q3", "FY25 Q4", "FY2025", "FY26 Q1", "FY26 上半（公司指引）", "FY2026（公司指引）", "標記", "來源1", "來源2／說明"])
QD = [("營收", [8.080, 15.369, 13.481, 17.435, None, 16.183, "=h1g_rev", "=g26_rev"], "松井：8,080／15,369／13,481／17,435；FY26 Q1 16,183", "公司：上半 234.4 億（說明資料 2025-11-07）＝8.08＋15.37"),
      ("營業利益", [-0.581, 3.074, 1.192, 3.232, None, 2.212, "=h1g_oi", "=g26_oi"], "松井：−581／3,074／1,192／3,232；2,212", "公司：上半 24.9 億"),
      ("淨利", [-0.528, 2.377, 0.778, 1.966, None, 1.657, 3.5, "=g26_ni"], "松井：−528／2,377／778／1,966；1,657", "上半指引 3,500（決算短信）"),
      ("EPS（基本）", [-7.05, 31.70, 10.37, 26.20, None, 22.09, 46.65, "=g26_eps"], "松井", "上半指引 46.65"),
      ("受注高", [10.43, 14.61, None, None, None, None, None, None], "說明資料 2025-11-07：上半 250.3、第二季 146.1（第一季＝差額 104.3，Derived）", "FY2025 全年 595.6 → 下半 345.3（Derived）；FY2026 計畫每季 150–170")]
for i, (lab, vals, s1, s2) in enumerate(QD, 4):
    Q.cell(i, 1, lab)
    for col, v in zip("BCDEFGHI", vals):
        if v is None: continue
        c = Q[f"{col}{i}"]; c.value = v; c.number_format = "#,##0.00" if lab == "EPS（基本）" else "#,##0.000"
        if not isinstance(v, str): c.font = BLUE
    Q.cell(i, 10, "Verified／Interested-party（指引）"); Q.cell(i, 11, s1); Q.cell(i, 12, s2)
for rr in (4, 5, 6, 7):
    Q[f"F{rr}"] = f"=SUM(B{rr}:E{rr})"; Q[f"F{rr}"].number_format = Q[f"C{rr}"].number_format
Q["F8"] = "=ord_25"; Q["F8"].number_format = "#,##0.000"
Q.cell(9, 1, "營業利益率"); [Q.cell(9, j, f"={c}5/{c}4").__setattr__("number_format", "0.0%") for j, c in zip(range(2, 10), "BCDEFGHI")]
Q.cell(10, 1, "檢查：FY2025 四季合計 − 年度（營收、營業利益、淨利；應約為 0）")
Q["B10"] = "=F4-a25_rev"; Q["C10"] = "=F5-a25_oi"; Q["D10"] = "=F6-a25_ni"
for c in "BCD": Q[f"{c}10"].number_format = "0.000"
Q.cell(11, 1, "FY2026 第一季占全年指引（營收／營業利益）"); Q["B11"] = "=G4/I4"; Q["C11"] = "=G5/I5"
for c in "BC": Q[f"{c}11"].number_format = "0.0%"
Q.cell(13, 1, "第二季（2026-07～09）財報尚未公布（去年為 11-07 公布）；本版 FY2026 錨定公司全年指引。FY2025 第一季虧損、第三季因記憶體量產投資延後而下修（公司第三季說明資料）")
Q.column_dimensions["A"].width = 52
for c in "BCDEFGHI": Q.column_dimensions[c].width = 13
Q.column_dimensions["J"].width = 20; Q.column_dimensions["K"].width = 60; Q.column_dimensions["L"].width = 60

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
CYC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2032)}   # C＝CY2025 … I＝CY2031
FYC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2030)}   # C＝FY2025 … G＝FY2029
DEFP = dict(dlt="dlt", tail="g_tail", den="", stk="", cw25="cw_25", ai="ai_25", w="w_hbm", pkt="g_pk_tail",
            gtss="g_tss", gmed="g_med", im="im", nioi="nioi", ganc="g_anc", carry="1", shl=None, ima=None, scn="A")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    shl = P["shl"] or f"shl_{s}"; ima = P["ima"] or f"ima_{s}"
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "CY2025／FY2025", "CY2026／FY2026", "CY2027／FY2027", "CY2028／FY2028", "CY2029／FY2029", "CY2030", "CY2031", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="0.00", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 10, note); R[key] = r; r += 1
    Y5 = range(2025, 2030)
    ws.cell(r, 1, "A. 日曆年（C＝CY2025 … I＝CY2031）").font = BOLD; r += 1
    row("bits", "HBM 出貨位元（同字母情境）", "EB", {CYC[y]: f"=hs{s}_{y % 100}" for y in Y5}, note="skhynix 成品 Mkt 頁")
    row("den", "DRAM 晶粒容量", "Gb", {CYC[y]: f"=den_{y % 100}{P['den']}" for y in Y5})
    row("stk", "平均堆疊層數", "層", {CYC[y]: f"=stk_{y % 100}{P['stk']}" for y in Y5})
    row("st", "HBM 堆疊數（每個堆疊封膠一次）", "十億個", {**{CYC[y]: f"={CYC[y]}{R['bits']}*8/({CYC[y]}{R['den']}*{CYC[y]}{R['stk']})" for y in Y5},
                                                    "H": f"=G{r}*(1+{P['tail']})", "I": f"=H{r}*(1+{P['tail']})"}, fmt="0.0000", note="EB × 8 ÷（晶粒容量 × 層數）；CY2030–31 以 g_tail 外推")
    row("xh", "HBM 封膠新機需求（下一年堆疊增量＋汰換）", "十億個", {CYC[y]: f"=MAX(0,{CYC[y+1]}{R['st']}-{CYC[y]}{R['st']})+{P['dlt']}*{CYC[y]}{R['st']}" for y in range(2025, 2031)}, fmt="0.0000", note="設備比產能早一年買")
    um = {2028: f"*um28_{s}", 2029: f"*um29_{s}"}
    row("vl", "CoWoS 片數需求（同字母情境顆數係數）", "千片", {"C": f"={P['cw25']}", **{CYC[y]: f"=cw_{y % 100}{um.get(y, '')}" for y in range(2026, 2030)},
                                                 "H": f"=G{r}*(1+{P['tail']})", "I": f"=H{r}*(1+{P['tail']})"}, fmt="#,##0", note="台積電成品；2028–29 × 情境係數")
    row("xl", "2.5D 封膠新機需求（下一年片數增量＋汰換）", "千片", {CYC[y]: f"=MAX(0,{CYC[y+1]}{R['vl']}-{CYC[y]}{R['vl']})+{P['dlt']}*{CYC[y]}{R['vl']}" for y in range(2025, 2031)}, fmt="#,##0")
    row("pk", "非 AI 封裝設備指數（SEMI 封裝 × 非 AI 設備 s ÷ 非 AI 設備 A）", "US$B", {**{CYC[y]: f"=pk_{y % 100}*(eqt{s}_{y % 100}-eqa{s}_{y % 100})/(eqtA_{y % 100}-eqaA_{y % 100})" for y in Y5},
                                                                      "H": f"=G{r}*(1+{P['pkt']})"}, fmt="0.00", note="markets/equipment；CY2030 以 g_pk_tail 外推")
    row("eqt", "設備總額（同字母情境）", "US$B", {**{CYC[y]: f"=eqt{s}_{y % 100}" for y in Y5}, "H": f"=G{r}*(1+{P['pkt']})"}, fmt="0.0", note="切割與雷射用")
    ws.cell(r, 1, "B. 校準（只用 FY2025 AI 錨點；FY(t)＝3/4×CY(t)＋1/4×CY(t+1)）").font = BOLD; r += 1
    fw = lambda key, y: f"(0.75*{CYC[y]}{R[key]}+0.25*{CYC[y+1]}{R[key]})"
    row("kh", "HBM：每單位新機需求的封膠營收 kH", "¥B／十億個", {"C": f"={P['ai']}*{P['w']}/{fw('xh', 2025)}"}, fmt="0.0", note="＝AI 錨點 × HBM 比重 ÷ FY2025 新機需求")
    row("kl", "2.5D：每千片新機需求的封膠營收 kL", "¥B／千片", {"C": f"={P['ai']}*(1-{P['w']})/{fw('xl', 2025)}"}, fmt="0.00000")
    ws.cell(r, 1, "C. 會計年度驅動推導（C＝FY2025 … G＝FY2029）").font = BOLD; r += 1
    FY1 = range(2026, 2030)
    row("shf", "AI 封膠市占係數（相對 2025）", "x", {FYC[y]: f"=1-{shl}/sh*{(y - 2025) / 4}" for y in Y5}, fmt="0.000", note="公司特有：A 最有利、C 最不利")
    row("ehd", "HBM 封膠（驅動推導）", "¥B", {FYC[y]: f"=$C${R['kh']}*{fw('xh', y)}*{FYC[y]}{R['shf']}" for y in Y5}, fmt="#,##0.00")
    row("eld", "CoWoS 等 2.5D 封膠（驅動推導）", "¥B", {FYC[y]: f"=$C${R['kl']}*{fw('xl', y)}*{FYC[y]}{R['shf']}" for y in Y5}, fmt="#,##0.00")
    row("mnd", "非 AI 封膠設備與模具（驅動推導）", "¥B", {FYC[y]: f"=(seg_mold-{P['ai']})*{fw('pk', y)}/{fw('pk', 2025)}" for y in Y5}, fmt="#,##0.00", note="FY2025 基期＝封膠 − AI 錨點")
    row("osd", "其他半導體：切割等（驅動推導）", "¥B", {FYC[y]: f"=seg_osemi*{fw('eqt', y)}/{fw('eqt', 2025)}" for y in Y5}, fmt="#,##0.00", note="隨設備總額")
    row("tsd", "新事業：TSS 售後等（驅動推導）", "¥B", {"C": "=seg_new", **{FYC[y]: f"={FYC[y-1]}{r}*(1+{P['gtss']})" for y in FY1}}, fmt="#,##0.00")
    row("mdd", "醫療器材（驅動推導）", "¥B", {"C": "=seg_med", **{FYC[y]: f"={FYC[y-1]}{r}*(1+{P['gmed']})" for y in FY1}}, fmt="#,##0.00")
    row("lsd", "雷射（驅動推導）", "¥B", {FYC[y]: f"=seg_las*{fw('eqt', y)}/{fw('eqt', 2025)}" for y in Y5}, fmt="#,##0.00", note="隨設備總額")
    row("revd", "營收（驅動推導，無錨點）", "¥B", {FYC[y]: f"={FYC[y]}{R['ehd']}+{FYC[y]}{R['eld']}+{FYC[y]}{R['mnd']}+{FYC[y]}{R['osd']}+{FYC[y]}{R['tsd']}+{FYC[y]}{R['mdd']}+{FYC[y]}{R['lsd']}" + ("+seg_adj" if y == 2025 else "") for y in Y5}, fmt="#,##0.00")
    ws.cell(r, 1, "D. FY2026 錨點（公司全年指引 ×（1＋g_anc）；三情境相同）").font = BOLD; r += 1
    row("arev", "FY2026 錨點營收", "¥B", {"D": f"=g26_rev*(1+{P['ganc']})"}, fmt="#,##0.00")
    row("aoi", "FY2026 錨點營業利益（指引營益率 16.0%）", "¥B", {"D": f"=D{R['arev']}*g26_oi/g26_rev"}, fmt="#,##0.00")
    row("ani", "FY2026 錨點淨利（指引淨利 ÷ 營業利益）", "¥B", {"D": f"=D{R['aoi']}*g26_ni/g26_oi"}, fmt="#,##0.00")
    row("sc", "錨點 ÷ 驅動推導（FY2026）", "x", {"D": f"=D{R['arev']}/D{R['revd']}"}, fmt="0.000", note="FY2026 各組成同比例縮放成錨點")
    sc_ref = f"D{R['sc']}" if s == "A" else "sc_A"
    row("cf", "非 AI 各項的基期調整（FY2027 起；用情境 A 的 sc）", "x", {"D": f"=1+{P['carry']}*({sc_ref}-1)"}, fmt="0.000", note="AI 兩項 FY2027 起回到驅動推導；B、C 的 sc 差額來自 AI 驅動較低，不算非 AI")
    ws.cell(r, 1, "E. 損益（FY2025 實績；FY2026 錨點；FY2027 起驅動推導）").font = BOLD; r += 1
    mult = lambda y, ai: ("1" if y == 2025 else (f"$D${R['sc']}" if y == 2026 else ("1" if ai else f"$D${R['cf']}")))
    for k, lab, ai in (("eh", "HBM 封膠", True), ("el", "CoWoS 等 2.5D 封膠", True), ("mn", "非 AI 封膠設備與模具", False), ("os", "其他半導體（切割等）", False),
                       ("ts", "新事業（TSS 等）", False), ("md", "醫療器材", False), ("ls", "雷射", False)):
        row(k, lab, "¥B", {FYC[y]: f"={FYC[y]}{R[k + 'd']}*{mult(y, ai)}" for y in Y5}, fmt="#,##0.00")
    row("eai", "AI 封膠合計", "¥B", {FYC[y]: f"={FYC[y]}{R['eh']}+{FYC[y]}{R['el']}" for y in Y5}, fmt="#,##0.00")
    row("rev", "營收", "¥B", {FYC[y]: f"={FYC[y]}{R['eai']}+{FYC[y]}{R['mn']}+{FYC[y]}{R['os']}+{FYC[y]}{R['ts']}+{FYC[y]}{R['md']}+{FYC[y]}{R['ls']}" + ("+seg_adj" if y == 2025 else "") for y in Y5}, fmt="#,##0.00", note="FY2026＝錨點營收")
    row("aish", "AI 封膠占營收", "%", {FYC[y]: f"={FYC[y]}{R['eai']}/{FYC[y]}{R['rev']}" for y in Y5}, fmt="0.0%")
    row("oi", "營業利益", "¥B", {"C": "=a25_oi", "D": f"=D{R['aoi']}", **{FYC[y]: f"=$D{r}+({P['im']}+{ima})*({FYC[y]}{R['rev']}-$D{R['rev']})" for y in range(2027, 2030)}}, fmt="#,##0.00",
        note="FY2027 起＝FY2026 營業利益＋增量營益率 ×（營收 − FY2026 營收）")
    row("oim", "營業利益率（結果）", "%", {FYC[y]: f"={FYC[y]}{R['oi']}/{FYC[y]}{R['rev']}" for y in Y5}, fmt="0.0%")
    row("ni", "淨利", "¥B", {"C": "=a25_ni", "D": f"=D{R['ani']}", **{FYC[y]: f"={FYC[y]}{R['oi']}*{P['nioi']}" for y in range(2027, 2030)}}, fmt="#,##0.00")
    row("eps", "EPS", "¥", {"C": "=a25_eps", **{FYC[y]: f"={FYC[y]}{R['ni']}*1000/shares" for y in FY1}}, fmt="#,##0.0")
    row("chk", "檢查：FY2025 營收、營業利益 − 實績；FY2026 營收 − 錨點（應為 0）", "¥B", {"C": f"=ABS(C{R['rev']}-a25_rev)+ABS(C{R['oi']}-a25_oi)+ABS(D{R['rev']}-D{R['arev']})"}, fmt="0.000")
    R["end"] = r
    return R

# ---------------- PL：三情境 ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — TOWA 損益（¥B；FY2025 實績；FY2026 錨定公司全年指引；FY2027–FY2029 驅動推導；FY2030–FY2032 收斂法見底部）。頂端為摘要，下方為三情境各一個完整推導區塊"
TITLES = {"A": "情境 A（順利）：上游 A（HBM 位元 A、顆數係數 A、設備市場 A）；AI 封膠市占不流失；增量營益率 +3pt",
          "B": "情境 B（平平）：上游 B（2028 顆數 ×0.8、2029 ×0.85；設備市場 B 2028 見頂）；市占流失 3pt；增量營益率＝迴歸",
          "C": "情境 C（不如預期）：上游 C（2028–29 顆數 ×0.6；設備市場 C 2027 見頂）；市占流失 8pt；增量營益率 −5pt"}
PB = {}; r0 = 60
for s in SC:
    PB[s] = engine(PL, r0, TITLES[s], scn=s)
    r0 = PB[s]["end"] + 2
    if s == "A": defname("sc_A", f"PL!$D${PB['A']['sc']}")
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "¥B", "#,##0.0"), ("eai", "  AI 封膠（HBM＋2.5D）", "¥B", "#,##0.0"), ("mn", "  非 AI 封膠設備與模具", "¥B", "#,##0.0"),
                                ("os", "  其他半導體（切割等）", "¥B", "#,##0.0"), ("ts", "  新事業（TSS 等）", "¥B", "#,##0.0"), ("ls", "  雷射", "¥B", "#,##0.0"),
                                ("aish", "AI 封膠占營收", "%", "0.0%"), ("oi", "營業利益", "¥B", "#,##0.0"), ("oim", "營業利益率", "%", "0.0%"),
                                ("ni", "淨利", "¥B", "#,##0.00"), ("eps", "EPS", "¥", "#,##0.0"), ("revd", "對照：營收（只用驅動，無錨點）", "¥B", "#,##0.0")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in range(2025, 2030):
            PL[f"{FYC[y]}{sr}"] = f"={FYC[y]}{PB[s][key]}"; PL[f"{FYC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對：三情境 FY2025 − 實績、FY2026 − 錨點（應為 0）"); PL[f"C{sr}"] = "=" + "+".join(f"C{PB[s]['chk']}" for s in SC); PL[f"C{sr}"].number_format = "0.000"
PL_CHK = sr
assert sr < 58
PL.column_dimensions["A"].width = 50; PL.column_dimensions["B"].width = 14; PL.column_dimensions["J"].width = 50
for c in "CDEFGHI": PL.column_dimensions[c].width = 13

# --- PL：延伸到 FY2032（收斂法） ---
p0 = PL.max_row + 3
PL.cell(p0, 1, "延伸到 FY2032：FY2025–FY2029 引用本頁頂端摘要；灰底 FY2030–FY2032 為收斂法，非驅動因子推導").font = BOLD
PL.cell(p0 + 1, 1, "收斂法：營收年增由 FY2029 年增線性收斂到 FY2032 的長期成長（A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt）；淨利率由 FY2029 線性收斂（A＝FY2029、B＝FY2029 與十年平均的中點、C＝十年平均 nim10）；股數維持 FY2029（EPS 與淨利同比例）")
EXTC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2033)}
hdr(PL, p0 + 2, ["情境（延伸）", "單位", "FY2025A", "FY2026", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
EXT = {}; pr = p0 + 3; gadd = {"A": "+0.02", "B": "", "C": "-0.02"}
for s in SC:
    PL.cell(pr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD; pr += 1
    R = {k: pr + i for i, k in enumerate(("rev", "g", "nim", "ni", "eps"))}
    for k, lab in (("rev", "營收"), ("g", "  營收年增"), ("nim", "  淨利率（淨利 ÷ 營收）"), ("ni", "淨利（收斂法延伸）"), ("eps", "EPS（延伸）")):
        PL.cell(R[k], 1, lab); PL.cell(R[k], 2, {"rev": "¥B", "ni": "¥B", "eps": "¥"}.get(k, "%"))
    for y in range(2025, 2033):
        c = EXTC[y]; pc = EXTC.get(y - 1)
        if y <= 2029:
            PL[f"{c}{R['rev']}"] = f"={FYC[y]}{SUMR[(s, 'rev')]}"; PL[f"{c}{R['ni']}"] = f"={FYC[y]}{SUMR[(s, 'ni')]}"
            PL[f"{c}{R['eps']}"] = f"={FYC[y]}{SUMR[(s, 'eps')]}"; PL[f"{c}{R['nim']}"] = f"={c}{R['ni']}/{c}{R['rev']}"
            if y > 2025: PL[f"{c}{R['g']}"] = f"={c}{R['rev']}/{pc}{R['rev']}-1"
        else:
            k = y - 2029
            mT = {"A": f"$G{R['nim']}", "B": f"($G{R['nim']}+nim10)/2", "C": "nim10"}[s]
            PL[f"{c}{R['g']}"] = f"=$G{R['g']}+((tp_g{gadd[s]})-$G{R['g']})*{k}/3"
            PL[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
            PL[f"{c}{R['nim']}"] = f"=$G{R['nim']}+({mT}-$G{R['nim']})*{k}/3"
            PL[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
            PL[f"{c}{R['eps']}"] = f"=$G{R['eps']}*{c}{R['ni']}/$G{R['ni']}"
            for kk in R: PL[f"{c}{R[kk]}"].fill = FADE
        for kk, fm in (("rev", "#,##0.0"), ("g", "0.0%"), ("nim", "0.0%"), ("ni", "#,##0.00"), ("eps", "#,##0.0")):
            if PL[f"{c}{R[kk]}"].value is not None: PL[f"{c}{R[kk]}"].number_format = fm
    PL.cell(R["rev"], 11, "FY2030 起為收斂法，非驅動因子推導")
    EXT[s] = R; pr += 6
PL.column_dimensions["K"].width = 40

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — 估值。第 0 節為目標價（DCF 與 Comparable 加權）；第 1–6 節為十年 P/E、循環位置規則值（設備股共用規則：見頂年 25 分位、中段中位數、谷底 75 分位）與目前倍數"
V["A2"] = "P/E 定義（自算）：日曆年底收盤價 ÷ 「當期會計年度」公司公布的基本 EPS（例：2025-12 對 FY2025＝2026年3月期）。價格與 EPS 皆以 2024-10-01 一股分三股調整"
MF = "#,##0"; YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
# 第 1–6 節先寫在第 60 列起（第 0 節要引用 P/E 範圍），第 0 節放最前面（第 4 列起）
b1 = 60
V.cell(b1, 1, "1. 10 年 P/E（年底）").font = BOLD
hdr(V, b1 + 1, ["年底", "收盤價（分割調整後）", "來源2：株探 年足終值", "兩者差距", "當期 FY EPS（調整後）", "P/E 當期", "", "", "標記", "說明"])
PX = {2016: (549.00, 548), 2017: (730.00, 729), 2018: (194.33, 194), 2019: (394.00, 393), 2020: (700.33, 700), 2021: (1081.67, 1081),
      2022: (563.33, 563), 2023: (2376.67, 2376), 2024: (1546.00, 1546), 2025: (2150.00, 2150)}
p1 = b1 + 2
for i, y in enumerate(range(2016, 2026), p1):
    a, k = PX[y]
    V.cell(i, 1, f"{y}-12"); V.cell(i, 2, a).font = BLUE; V.cell(i, 3, k).font = BLUE; V.cell(i, 4, f"=B{i}/C{i}-1")
    V.cell(i, 5, f"=Hist!J{i - p1 + 6}"); V.cell(i, 6, f"=B{i}/E{i}")
    for c, fm in (("B", "#,##0.00"), ("C", "#,##0"), ("D", "0.00%"), ("E", "#,##0.00"), ("F", "0.0")): V[f"{c}{i}"].number_format = fm
    V.cell(i, 9, "Verified")
    V.cell(i, 10, "來源1：Yahoo Finance 月線 12 月 Close（已依 2024 年分割調整）；來源2：株探 年足（取整）；EPS：Hist 頁" if y == 2016 else ("FY2019 EPS 僅 4.92（淨利 3.7 億），P/E 80 倍為極端值" if y == 2019 else ""))
PE = f"F{p1}:F{p1 + 9}"
q = p1 + 11
V.cell(q, 1, "兩個價格來源可能共用同一資料商（不完全獨立）；10 年都有價格，無缺年。"); q += 2
V.cell(q, 1, "2. 分位（10 年）").font = BOLD; hdr(V, q + 1, ["統計", "P/E 當期", "用途"])
for i, (nm, p, use) in enumerate((("25 分位", 0.25, "見頂年"), ("中位數", 0.5, "中段"), ("75 分位", 0.75, "谷底")), q + 2):
    V.cell(i, 1, nm); V.cell(i, 2, f"=PERCENTILE({PE},{p})").number_format = "0.0"; V.cell(i, 3, use)
defname("pe_peak", f"Val!$B${q + 2}"); defname("pe_mid", f"Val!$B${q + 3}"); defname("pe_trough", f"Val!$B${q + 4}")
PCT_ROWS = (q + 1, q + 4)
V.cell(q + 5, 1, "年數"); V[f"B{q + 5}"] = f"=COUNT({PE})"
V.cell(q + 6, 1, "敏感度：去掉 2019 年（極端值）時 25／50／75 分位")
for j, p in enumerate((0.25, 0.5, 0.75)):
    V.cell(q + 6, 2 + j, f"=PERCENTILE(({'F%d:F%d' % (p1, p1 + 2)},{'F%d:F%d' % (p1 + 4, p1 + 9)}),{p})").number_format = "0.0"
q += 8
V.cell(q, 1, "3. FY2027 的循環位置（設備股共用規則；見頂年度讀設備市場模型，不在本模型另訂）").font = BOLD
V.cell(q + 1, 1, CP.rule_text(PEAK)); CYCR = q + 1
hdr(V, q + 2, ["情境", "設備市場見頂年度", "FY2027 對照日曆年", "判定", "倍數"]); CYH = q + 2
POS = {}
for i, s in enumerate(SC, q + 3):
    V.cell(i, 1, s); V.cell(i, 2, f"=pk_{s}"); V.cell(i, 3, 2027).font = BLUE
    V.cell(i, 4, "=" + CP.excel_rule(f"C{i}", f"pk_{s}"))
    V.cell(i, 5, f'=IF(D{i}="見頂",pe_peak,IF(D{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"; POS[s] = i
q += 7
V.cell(q, 1, "4. 隱含股價（FY2027 EPS × 依判定的倍數；其餘倍數僅供參考）").font = BOLD
hdr(V, q + 1, ["情境", "EPS 年度", "EPS", "倍數類型", "倍數", "隱含股價", "相對現價", "備註"]); VH = q + 1
vr = q + 2; VAL = {}; V4 = vr
for s in SC:
    for y, f, nm in ((2027, f"E{POS[s]}", None), (2027, "pe_peak", "參考：25 分位"), (2027, "pe_mid", "參考：中位數"), (2027, "pe_trough", "參考：75 分位"), (2028, "pe_mid", "參考：FY2028 × 中位數")):
        V.cell(vr, 1, s); V.cell(vr, 2, f"FY{y}"); V.cell(vr, 3, f"=PL!{FYC[y]}{SUMR[(s, 'eps')]}").number_format = "#,##0.0"
        V.cell(vr, 4, nm if nm else f"=\"規則值：\"&D{POS[s]}"); V.cell(vr, 5, f"={f}").number_format = "0.0"
        V.cell(vr, 6, f"=C{vr}*E{vr}").number_format = "#,##0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
        if nm is None:
            for c in range(1, 9): V.cell(vr, c).fill = KEY
            VAL[s] = vr
        vr += 1
V4E = vr - 1
q = vr + 1
V.cell(q, 1, "5. 目前倍數（僅供參考）").font = BOLD; hdr(V, q + 1, ["項目", "A／數值", "B", "C", "備註"]); CMH = q + 1
fr = q + 2; CUR = {}
for key, nm, fs, fm, note in (("px", "現價（2026-10-23 收盤）", ["=px"], "#,##0", "株探、kabutore"),
                        ("pe25", "P/E：現價 ÷ FY2025 EPS 61.22", ["=px/a25_eps"], "0.0", "上一個會計年度（營益率 12.7% 的低點年）"),
                        ("pettm", "P/E：現價 ÷ 近 12 月 EPS（FY25 Q2–Q4＋FY26 Q1）", ["=px/(Q!C7+Q!D7+Q!E7+Q!G7)"], "0.0", "近 12 月 EPS ＝ 31.70＋10.37＋26.20＋22.09 ＝ 90.36（Derived；與 stockanalysis TTM 相同）"),
                        ("peg", "P/E：現價 ÷ 公司 FY2026 指引 EPS 93.30", ["=px/g26_eps"], "0.0", "命題中的「約 21 倍」"),
                        ("sa", "遠期 P/E：stockanalysis（共識，2026-09-25）", ["=sa_fpe"], "0.0", "參考，不用於估值；當時股價 2,193"),
                        ("pe27", "現價 ÷ FY2027 EPS（模型）", [f"=px/PL!{FYC[2027]}{SUMR[(s, 'eps')]}" for s in SC], "0.0", "A／B／C"),
                        ("pe28", "現價 ÷ FY2028 EPS（模型）", [f"=px/PL!{FYC[2028]}{SUMR[(s, 'eps')]}" for s in SC], "0.0", "A／B／C")):
    V.cell(fr, 1, nm)
    for j, f in enumerate(fs): V.cell(fr, 2 + j, f).number_format = fm
    V.cell(fr, 5, note); CUR[key] = fr; fr += 1
q = fr + 1
V.cell(q, 1, "6. 舊法：規則值情境加權（參考，不是目標價）").font = BOLD; OLD = q
hdr(V, q + 1, ["項目", "數值", "", "", "說明"])
V.cell(q + 2, 1, "規則值情境加權價（12 個月；參考）"); V.cell(q + 2, 2, f"=tp_wA*F{VAL['A']}+tp_wB*F{VAL['B']}+tp_wC*F{VAL['C']}").number_format = MF
V.cell(q + 2, 5, "權重 × 各情境規則值（第 4 節）")
V.cell(q + 3, 1, "相對現價（規則值情境加權價 ÷ 現價 − 1）"); V.cell(q + 3, 2, f"=B{q + 2}/px-1").number_format = "0.0%"
V.cell(q + 4, 1, "規則值方法"); V.cell(q + 4, 2, "各情境 FY2027 EPS × 十年（2016–2025 年底）P/E 依循環位置取分位（設備股共用規則）。十年倍數多數來自 AI 前（營收 250–310 億日圓、營益率 3–14%），2019 年 80 倍為極端值；隱含價約為 FY2027 年底（2028-03）的價格、未折現")
OLDE = q + 4

# --- 第 0 節：目標價（第 4 列起） ---
T = {}
def tset(rr, key, label, fs, fmt, note=""):
    V.cell(rr, 1, label); T[key] = rr
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{rr}"] = f; V[f"{c}{rr}"].number_format = fmt
    if note: V[f"E{rr}"] = note
t0 = 4
V.cell(t0, 1, "0. 目標價").font = BOLD; T["top"] = t0
V.cell(t0 + 1, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD; T["h01"] = t0 + 1
hdr(V, t0 + 2, ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"]); T["hd1"] = t0 + 2
V.cell(t0 + 3, 1, "情境 EPS（FY2027–FY2032）"); V[f"E{t0 + 3}"] = "PL 頁延伸區塊；FY2030 起為收斂法"; T["eps_h"] = t0 + 3
EPSR = {}
for i, y in enumerate(YRS6):
    tset(t0 + 4 + i, f"eps{y}", f"  FY{y}", [f"=PL!{EXTC[y]}{EXT[s]['eps']}" for s in SC], "#,##0.0"); EPSR[y] = t0 + 4 + i
rr = t0 + 10
V.cell(rr, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）"); T["cf_h"] = rr
V[f"E{rr}"] = "模型無折舊與資本支出項目，故不用（淨利＋折舊－資本支出）；TOWA 資本支出 FY2025 41.9 億、FY2026 計畫 55.0 億（約營收 9%、指引淨利的約八成），高於再投資率 36.7%，此式可能高估現金流"
tset(rr + 1, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）")
RV = rr + 1; CFR = {}
for i, y in enumerate(YRS6):
    tset(rr + 2 + i, f"cf{y}", f"  FY{y}", [f"={c}{EPSR[y]}*(1-$B${RV})" for c in "BCD"], "#,##0.0"); CFR[y] = rr + 2 + i
c1, c6 = CFR[2027], CFR[2032]
tset(rr + 8, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=MAX(0,SUMPRODUCT({c}{c1}:{c}{c6},1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}{c6}*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6)" for c in "BCD"], MF,
     "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底，t＝1 對 FY2027；負值取 0（股東有限責任）；會計年度 4–3 月，FY2027 仍以 t＝1 折現（略高估）")
tset(rr + 9, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}{rr + 8}*(1+tp_r)" for c in "BCD"], MF)
tset(rr + 10, "dcftp", "DCF 目標價（情境加權）", [f"=tp_wA*B{rr + 9}+tp_wB*C{rr + 9}+tp_wC*D{rr + 9}", None, None], MF, "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁 tp_wA／tp_wB／tp_wC")
q = rr + 12
V.cell(q, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD; T["h02"] = q
hdr(V, q + 1, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"]); T["phd"] = q + 1
SAS = "stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左；本 repo hanmi／disco 同值）"
PEERS = [("Disco", "6146.T", 26.01, "2026-10-23", SAS), ("東京精密（Accretech）", "7729.T", 21.86, "2026-09-29", SAS),
         ("ASMPT（封膠與鍵合競爭者）", "0522.HK", 30.26, "2026-09-23", SAS), ("BE Semiconductor（BESI）", "BESI.AS", 25.82, "2026-10-23", SAS),
         ("Kulicke & Soffa", "KLIC", 12.61, "2026-10-23", SAS)]
for i, (nm, tk, pe, dt, src) in enumerate(PEERS, q + 2):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
T["p1"] = q + 2; T["pl"] = q + 1 + len(PEERS)
q = T["pl"] + 1
tset(q, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C{T['p1']}:C{T['pl']})", None, None], "0.0", "後段設備同業：日本 Disco、東京精密；封裝設備 ASMPT（封膠直接競爭者）、BESI、Kulicke & Soffa。同業名單不含中國 A 股")
fwd = {2027: "5/12", 2028: "7/12"}
feps = lambda c: "(" + "+".join(f"{w}*{c}{EPSR[y]}" for y, w in fwd.items()) + ")"
tset(q + 1, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=tp_wA*" + feps("B") + "+tp_wB*" + feps("C") + "+tp_wC*" + feps("D"), None, None], "#,##0.0",
     "約 2027-10 時的未來 12 月（2027-11～2028-10）≈ 5／12 × FY2027（2027-04～2028-03）＋7／12 × FY2028（會計年度 4–3 月，依月份調整）")
tset(q + 2, "comptp", "Comparable 目標價", [f"=B{q}*B{q + 1}", None, None], MF, "同業中位數 × 上列 EPS")
q += 4
V.cell(q, 1, "0.3 結論").font = BOLD; T["h03"] = q
hdr(V, q + 1, ["項目", "數值", "", "", "說明"]); T["hd3"] = q + 1
q += 2
tset(q, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B{T['comptp']}", None, None], MF, "tp_w_dcf × DCF 目標價＋（1 − tp_w_dcf）× Comparable 目標價；tp_w_dcf 見 In 頁")
tset(q + 1, "comp2", "Comparable 目標價（市場主流同業法）", [f"=B{T['comptp']}", None, None], MF, "引用 0.2")
tset(q + 2, "cons", "市場共識目標價", ["=cons_tp", None, None], MF, "stockanalysis：6 位分析師平均 ¥3,300（中位數 3,300；區間 2,600–4,000），2026-09-17 更新")
tset(q + 3, "upw", "上檔空間：加權目標價", [f"=B{q}/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-23 收盤）")
tset(q + 4, "upc", "上檔空間：Comparable 目標價", [f"=B{q + 1}/px-1", None, None], "0.0%")
tset(q + 5, "geps", "與市場共識目標價的差距：EPS 因素", [f"=B{T['feps']}/cons_feps-1", None, None], "0.0%", "本模型 12 個月後前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝FY2026 共識淨利 8.70 ÷ 股數，早約一年半")
tset(q + 6, "gmul", "與市場共識目標價的差距：倍數因素", [f"=(cons_tp/cons_feps)/B{T['pmed']}-1", None, None], "0.0%")
V[f"E{q + 6}"] = f'="共識目標價 ÷ 共識前瞻 EPS 隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B{T["pmed"]},"0.0")&"x"'
tset(q + 7, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D{T["dcf12"]},"#,##0")&"～"&TEXT(B{T["dcf12"]},"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(q + 8, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）",
     [f"=IF(px/{c}{EPSR[2028]}<MIN({PE}),0,IF(px/{c}{EPSR[2028]}>MAX({PE}),1,PERCENTRANK({PE},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 1 節當期口徑）的分位；只作檢查，不進目標價")
V.cell(q + 9, 1, "目標價方法"); T["meth"] = q + 9
V[f"B{q + 9}"] = ("DCF：股東現金流＝EPS ×（1 − g／ROE）（模型無折舊與資本支出項目），r 10%、g 4%、ROE 近三年平均 10.9%，FY2027–32 明確期（FY2030 起收斂法）＋永續；"
                  "Comparable：Disco、東京精密、ASMPT、BESI、Kulicke & Soffa 前瞻 P/E 中位數 × 12 個月後前瞻 EPS。偏誤：再投資率低估 TOWA 的資本支出（偏高估 DCF）；"
                  "同業中 Disco、ASMPT、BESI 的利潤率與 AI 曝險高於 TOWA（偏高估 Comparable）；十年倍數多來自 AI 前的小型股時期。")
for k in ("dcftp", "comptp", "tpw"): V[f"B{T[k]}"].fill = KEY
assert q + 9 < b1 - 1, (q + 9, b1)
V.column_dimensions["A"].width = 52
for c in "BCDEFGHI": V.column_dimensions[c].width = 15
V.column_dimensions["E"].width = 60; V.column_dimensions["J"].width = 70

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A）：對 FY2027、FY2028 營收與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方，與 PL 同一引擎；校準會跟著重算）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收 ¥B", "FY2027 EPS ¥", "相對基準", "FY2028 營收 ¥B", "FY2028 EPS ¥", "相對基準", "FY2027 AI 占營收", "FY2026 EPS ¥", "相對基準"])
cases = [("AI 錨點 7.0（−30%）", "ai_25 低", dict(ai="ai_25_lo")), ("AI 錨點 13.0（+30%）", "ai_25 高", dict(ai="ai_25_hi")),
         ("HBM 比重 35%", "w_hbm 低", dict(w="w_hbm_lo")), ("HBM 比重 65%", "w_hbm 高", dict(w="w_hbm_hi")),
         ("汰換／升級比例 5%", "dlt 低", dict(dlt="dlt_lo")), ("汰換／升級比例 15%", "dlt 高", dict(dlt="dlt_hi")),
         ("晶粒容量 低", "den 低欄", dict(den="_lo")), ("晶粒容量 高", "den 高欄", dict(den="_hi")),
         ("堆疊層數 低（同位元更多堆疊）", "stk 低欄", dict(stk="_lo")), ("堆疊層數 高", "stk 高欄", dict(stk="_hi")),
         ("CY2030–31 需求年增 0%", "g_tail 低", dict(tail="g_tail_lo")), ("CY2030–31 需求年增 20%", "g_tail 高", dict(tail="g_tail_hi")),
         ("CoWoS 2025 600 千片", "cw_25 低", dict(cw25="cw_25_lo")),
         ("CY2030 封裝設備年增 0%", "g_pk_tail 低", dict(pkt="g_pk_tail_lo")), ("CY2030 封裝設備年增 10%", "g_pk_tail 高", dict(pkt="g_pk_tail_hi")),
         ("新事業年增 0%", "g_tss 低", dict(gtss="g_tss_lo")), ("新事業年增 6%", "g_tss 高", dict(gtss="g_tss_hi")),
         ("AI 封膠市占流失 8pt（C 的值）", "shl_C", dict(shl="shl_C")), ("AI 封膠市占流失 12pt", "shl_C 高", dict(shl="shl_C_hi")),
         ("增量營益率 −5pt", "im 低", dict(im="im_lo")), ("增量營益率 +5pt", "im 高", dict(im="im_hi")),
         ("增量營益率調整取 C 的 −5pt", "ima_C", dict(ima="ima_C")),
         ("淨利÷營業利益 0.65", "nioi 低", dict(nioi="nioi_lo")), ("淨利÷營業利益 0.72", "nioi 高", dict(nioi="nioi_hi")),
         ("FY2026 錨點 −5%", "g_anc 低", dict(ganc="g_anc_lo")), ("FY2026 錨點 +8%", "g_anc 高", dict(ganc="g_anc_hi")),
         ("錨點差額不延續到 FY2027", "非 AI 基期回到驅動推導", dict(carry="0"))]
def put_case(i, name, setting, R, pre=""):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    SE[f"C{i}"] = f"={pre}{FYC[2027]}{R['rev']}"; SE[f"D{i}"] = f"={pre}{FYC[2027]}{R['eps']}"; SE[f"F{i}"] = f"={pre}{FYC[2028]}{R['rev']}"; SE[f"G{i}"] = f"={pre}{FYC[2028]}{R['eps']}"
    SE[f"E{i}"] = f"=D{i}/D$4-1"; SE[f"H{i}"] = f"=G{i}/G$4-1"; SE[f"I{i}"] = f"={pre}{FYC[2027]}{R['aish']}"
    SE[f"J{i}"] = f"={pre}{FYC[2026]}{R['eps']}"; SE[f"K{i}"] = f"=J{i}/J$4-1"
    for c in "CF": SE[f"{c}{i}"].number_format = "#,##0.0"
    for c in "DGJ": SE[f"{c}{i}"].number_format = "#,##0.0"
    for c in "EHIK": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
base_row = 70
for n, (nm, setting, ov) in enumerate(cases):
    Rn = engine(SE, base_row + n * 60, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    put_case(5 + n, nm, setting, Rn)
nn = 5 + len(cases)
put_case(nn, "情境 B", "上游 B；市占流失 3pt；增量營益率＝迴歸", PB["B"], "PL!")
put_case(nn + 1, "情境 C", "上游 C；市占流失 8pt；增量營益率 −5pt", PB["C"], "PL!")
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：校準在 FY2025 進行，所以改變 AI 錨點、HBM 比重、汰換比例等參數時，FY2025 仍等於實績，只改變之後的路徑形狀；FY2026 錨定公司指引，只受 g_anc 影響。")
assert nn + 3 < base_row
# 目標價參數
s0 = SE.max_row + 3
SE.cell(s0, 1, "目標價參數（只改 DCF 參數與權重，PL 收斂法的長期成長不隨 tp_g 變動）").font = BOLD
hdr(SE, s0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
def dcf_formula(c, r_="tp_r", g="tp_g"):
    cf = lambda y: f"Val!{c}{EPSR[y]}*(1-MIN({g}/tp_roe,0.8))"
    pv = "+".join(f"{cf(y)}/(1+{r_})^{t}" for t, y in enumerate(YRS6, 1))
    return f"MAX(0,{pv}+{cf(2032)}*(1+{g})/({r_}-{g})/(1+{r_})^6)"
tcases = [("基準", "r 10%、g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "r 11%", dict(r_="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", dict(r_="(tp_r-0.01)")),
          ("長期成長 +1pt", "g 5%（永續與再投資率）", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", dict(g="(tp_g-0.01)")),
          ("DCF 權重 0.3", "Comparable 0.7", dict(w="0.3")), ("DCF 權重 0.7", "Comparable 0.3", dict(w="0.7"))]
for k, (nm, st, kw) in enumerate(tcases):
    i = s0 + 2 + k; w = kw.pop("w", "tp_w_dcf"); rr_ = kw.get("r_", "tp_r")
    SE.cell(i, 1, nm); SE.cell(i, 2, st)
    SE[f"C{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_formula(c, **kw)}" for s, c in zip(SC, "BCD")) + f")*(1+{rr_})"
    SE[f"D{i}"] = f"=Val!B{T['comptp']}"; SE[f"E{i}"] = f"={w}*C{i}+(1-{w})*D{i}"; SE[f"F{i}"] = f"=E{i}/E${s0 + 2}-1"
    for c in "CDE": SE[f"{c}{i}"].number_format = MF
    SE[f"F{i}"].number_format = "0.0%"
T["sens0"] = s0; T["sens_last"] = s0 + 1 + len(tcases)
SE.column_dimensions["A"].width = 40; SE.column_dimensions["B"].width = 40
for c in "CDEFGHIJK": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — FY2026 錨定公司全年指引（A 節為恆等式）；A2 節列只用驅動推導的結果；B 節共識只作參考，未用於校準或錨點。差距＝模型 ÷ 對照 − 1"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.00", pt=False):
    global rr_
    RC.cell(rr_, 1, label); RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        L_ = chr(64 + cv)
        RC.cell(rr_, cd, f"={L_}{rr_}-B{rr_}" if pt else f"={L_}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y=2026: (lambda s: f"=PL!{FYC[y]}{SUMR[(s, key)]}")
pa = lambda key: (lambda s: f"=PL!D{PB[s][key]}")
RC.cell(rr_, 1, "A. FY2026 錨點（公司指引）：差距為 0 是恆等式，不是獨立檢驗").font = BOLD; rr_ += 1
rec("FY2026 營收 vs 公司指引", "=g26_rev", pl("rev"), "決算短信 2026-08-06（640.0 億）", "g_anc＝0")
rec("FY2026 營業利益 vs 公司指引", "=g26_oi", pl("oi"), "同上（102.4 億）")
rec("FY2026 EPS vs 公司指引", "=g26_eps", pl("eps"), "同上（93.30）", "股數差異造成的小差", fmt="#,##0.0")
rec("FY2026 營收 vs 第一季 × 4", "=q1_rev*4", pl("rev"), "第一季 16.2（實績）", "第一季占全年指引 25.3%；第二季起計畫每季接單 150–170 億")
rec("FY2026 營業利益率 vs 第一季（百分點）", "=q1_oi/q1_rev", pl("oim"), "第一季 13.7%", "指引隱含第二至四季營益率約 16.8%", fmt="0.0%", pt=True)
rec("FY2026 半導體＋雷射營收 vs 公司事業計畫（484.3＋32.5）", "=g26_semi+g26_las", lambda s: f"=PL!D{PB[s]['eai']}+PL!D{PB[s]['mn']}+PL!D{PB[s]['os']}+PL!D{PB[s]['ls']}", "決算說明資料 2026-05", "事業組成經錨點同比例縮放（sc），組成比例來自驅動推導")
RC.cell(rr_, 1, "A2. 只用驅動推導時的 FY2026（對照；說明錨點改變了多少）").font = BOLD; rr_ += 1
rec("驅動推導營收 vs 錨點營收", "=PL!D" + str(PB["A"]["arev"]), pa("revd"), "PL 區塊 C、D 節", "負＝驅動推導低於公司指引")
rec("驅動推導 雷射 vs 公司計畫 32.5", "=g26_las", pa("lsd"), "公司計畫 +61.9%（雷射新機）", "本模型雷射隨設備總額，低估公司計畫")
RC.cell(rr_, 1, "B. FY2026 vs 共識（參考）").font = BOLD; rr_ += 1
rec("營收 vs 共識（stockanalysis 2026-09-17，6 位）", "=cons_rev26", pl("rev"), "stockanalysis（69.3；區間 65.3–74.3）", "共識高於公司指引 8%")
rec("營業利益 vs 共識", "=cons_oi26", pl("oi"), "stockanalysis（12.25）")
rec("淨利 vs 共識", "=cons_ni26", pl("ni"), "stockanalysis（8.70）")
rec("EPS vs 共識（淨利 ÷ 股數）", "=cons_feps", pl("eps"), "Derived（約 116）", "差距約 −20%，未超過 30%：本模型 FY2026 採公司指引", fmt="#,##0.0")
RC.cell(rr_, 1, "C. 結構檢查").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "FY2025 AI 封膠占營收（校準結果；命題的「約兩成」）"); RC.cell(rr_, 3, f"=PL!C{SUMR[('A', 'aish')]}").number_format = "0.0%"; rr_ += 1
RC.cell(rr_, 1, "FY2027 AI 封膠占營收（A／B／C）")
for j, s in enumerate(SC): RC.cell(rr_, 3 + 2 * j, f"=PL!E{SUMR[(s, 'aish')]}").number_format = "0.0%"
rr_ += 1
RC.cell(rr_, 1, "FY2025 受注高 ÷ 營收（book-to-bill）"); RC.cell(rr_, 3, "=ord_25/a25_rev").number_format = "0.00"; RC.cell(rr_, 10, "1.10：接單領先營收，支持 FY2026 指引"); rr_ += 1
RC.cell(rr_, 1, "FY2026 錨點營收 ÷（期初受注殘高＋計畫接單下限 4 × 150 億）"); RC.cell(rr_, 3, f"=PL!D{SUMR[('A', 'rev')]}/(bl_25+4*15)").number_format = "0.00"; RC.cell(rr_, 10, "營收不超過可出貨量（＜1 表示指引可由殘高與接單支撐）"); rr_ += 1
RC_LAST = rr_ - 1
RC.column_dimensions["A"].width = 62
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 70

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文以網頁擷取工具讀取；說明資料圖表擷取錯位處未採用；版面與頁碼待人工複核）"; SO["A1"].font = BOLD
srcs = [("公司：2026年3月期 決算短信（JPX iXBRL，2026-04-22 登錄）", "https://www2.jpx.co.jp/disc/63150/081220260422508001_tse-acedjpsm-63150-20260422508001-ixbrl.htm"),
 ("公司：2027年3月期 第1四半期 決算短信（2026-08-06）", "https://www2.jpx.co.jp/disc/63150/081220260702586673_tse-qcedjpsm-63150-20260702586673-ixbrl.htm"),
 ("公司：2026年3月期 決算說明資料（2026-05）", "https://www.towajapan.co.jp/jp/wp-content/uploads/sites/2/2026/05/2025FY_end_presentation.pdf"),
 ("公司：2026年3月期 第2四半期 決算說明會資料（2025-11-07；HBM 台數、TechInsights 市占）", "https://www.towajapan.co.jp/jp/wp-content/uploads/sites/2/2025/11/2026_secondquarter_presentation.pdf"),
 ("公司：2026年3月期 第3四半期 決算說明資料（記憶體量產投資延後）", "https://www.towajapan.co.jp/jp/wp-content/uploads/sites/2/2026/02/2026_thirdquarter_presentation.pdf"),
 ("公司：歷年決算短信（2018-05-10、2019-05、2020-05-07）", "https://www2.jpx.co.jp/disc/63150/140120200507404921.pdf"),
 ("EDINET：四半期報告書 第44期第3四半期（FY2020 通期數字）", "https://disclosure2dl.edinet-fsa.go.jp/searchdocument/pdf/S100NDD0.pdf"),
 ("松井證券：決算情報（年度、季度）", "https://finance.matsui.co.jp/stock/6315/settlement/index"),
 ("kabutan：業績（分割調整 EPS）、日足、年足", "https://s.kabutan.jp/stocks/6315/historical_prices/daily/"),
 ("kabutore：現在株價、財務分析（毛利）", "https://www.kabutore.biz/genzaikabuka/6315.html"),
 ("the-shashi：配當與 EPS、營業利益歷史、股票分割", "https://the-shashi.com/tse/6315/financials/profit/"),
 ("tipranks：股票分割公告（三分割，2024-10-01）", "https://www.tipranks.com/news/company-announcements/towa-corporation-announces-stock-split-and-dividend-revision"),
 ("tousaiblog：FY2025 決算摘要（2026-05-12）", "https://www.tousaiblog.org/?p=3389"),
 ("edinetdb：說明資料彙整、FY2027 Q1", "https://edinetdb.jp/company/E01708/presentations/"),
 ("Yahoo Finance：月線歷史價格（分割調整）", "https://finance.yahoo.com/quote/6315.T/history/"),
 ("stockanalysis：共識、目標價、統計、比率", "https://stockanalysis.com/quote/tyo/6315/forecast/"),
 ("kabuyoho：進度率", "https://kabuyoho.jp/sp/report?bcode=6315"),
 ("36Kr／半導體產業縱橫：CoWoS 2025 67 萬片（2025-12-04）", "https://eu.36kr.com/en/p/3580962946874242"),
 ("市場層與其他模型：skhynix、tsmc、markets/equipment 成品；markets/memory/market.py（本 repo）", ""),
 ("無法開啟：irbank（權限逾時）、limo.media、kabutan 股價頁（桌面版）；kabuyoho 分析師頁（速率限制）", ""),
 ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 80; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
PROP = ("TOWA 握有封膠設備約 2/3 市占，但營收只有約兩成來自 AI（HBM 與 CoWoS 封膠），FY2025 營益率僅 12.7%；"
        "問題是 HBM 堆疊與 2.5D 封裝的封膠增量，能否在 FY2027 以後延續並把營益率拉回 16% 以上，撐住約 21 倍的指引本益比？")
lines = [("TOWA（6315.T）收支與估值模型 v0.1.1（AI 半導體 Project 步驟 5：設備與材料）　2026-10-25", True),
 ("模型命題：" + PROP, False),
 ("註：21 倍＝現價 ¥1,968（2026-10-23）÷ 公司 FY2026 指引 EPS ¥93.30；÷ FY2025 EPS ¥61.22 為 32.1 倍（Val 第 5 節）。", False), ("", False),
 ("驅動因子 → 推導量（因果方向：左邊決定右邊）", True),
 ("PL 區塊｜HBM 出貨位元（skhynix 成品，同字母情境）÷（晶粒容量 × 層數）→ HBM 堆疊數 → 下一年增量＋汰換 → 新機需求 × kH（以 FY2025 AI 錨點校準）→ HBM 封膠營收", False),
 ("PL 區塊｜CoWoS 片數（台積電成品；B、C × 同字母顆數係數）→ 下一年增量＋汰換 → 新機需求 × kL → 2.5D 封膠營收", False),
 ("PL 區塊｜SEMI 組裝與封裝設備 ×（情境 s 非 AI 設備 ÷ 情境 A 非 AI 設備）（markets/equipment）→ 非 AI 封膠設備與模具；設備總額 → 切割與雷射；新事業、醫療以自身歷史年增", False),
 ("PL（D 節）｜FY2026 錨點＝公司全年指引（營收 640、營業利益 102.4、淨利 70.0 億）×（1＋g_anc）；各組成同比例縮放，FY2027 起 AI 兩項回到驅動推導、非 AI 以錨定水準（情境 A 的 sc）為基期", False),
 ("PL（E 節）｜營業利益＝FY2026 錨點＋（Hist 迴歸增量營益率＋情境調整）×（營收 − FY2026 營收）；淨利＝營業利益 × 0.684（公司指引比例）", False),
 ("Val｜第 0 節目標價（DCF 與 Comparable 各半）；第 1–6 節十年 P/E、循環位置（設備股共用規則：見頂年度 " + "、".join(f"{s_} {PEAK[s_]}" for s_ in SC) + " → FY2027：" + CP.fy2027_text(PEAK) + "）、規則值與目前倍數", False), ("", False),
 ("會計年度：4 月到隔年 3 月（FY2025＝2025-04～2026-03，公司「2026年3月期」）。日曆年轉會計年度：FY(t)＝3/4 × CY(t)＋1/4 × CY(t+1)。金額為日圓，未建匯率。", False),
 ("情境：A＝順利、B＝平平、C＝不如預期，各讀同字母的上游情境（HBM 位元、顆數係數、設備市場）；公司特有變數（AI 封膠市占流失、增量營益率調整）A 最有利、C 最不利。FY2030 起為收斂法，非驅動因子推導。", False),
 ("可識別性：AI 錨點（FY2025 ¥10.0B）為 Assumed（公司不揭露），kH、kL 各由一點校準，只識別乘積；HBM 比重、汰換比例、晶粒容量與層數只改變路徑形狀。", False),
 ("已知限制：見 towa/README.md「待辦與限制」。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("待複核：說明資料圓餅（移轉 59%／壓縮 41%、地區別）擷取錯位，未採用為點估計；市占 64.8% 的圓餅標籤未明示公司；FY2020 營業利益單一來源。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 180


# ---------------- v0.1.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00、14:17）；新增列放在各頁現有最後一列之後，既有儲存格不變 ----------------
_U = f"B{T['upw']}"
_G = V.max_row + 2
V.cell(_G, 1, "上檔空間級距")
V[f"B{_G}"] = f'=IF({_U}>0.5,"大於 50%",IF({_U}>0.25,"25%–50%",IF({_U}>=0,"0%–25%",IF({_U}>=-0.25,"0%～−25%",IF({_U}>=-0.5,"−25%～−50%","小於 −50%")))))'
V[f"B{_G}"].number_format = "@"; V[f"E{_G}"] = "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）"
V.cell(_G + 1, 1, "投資建議")
V[f"B{_G + 1}"] = f'=IF({_U}>0.5,"強力買進",IF({_U}>0.25,"買進",IF({_U}>=-0.25,"中立",IF({_U}>=-0.5,"放空","強力放空"))))'
V[f"B{_G + 1}"].number_format = "@"
V[f"E{_G + 1}"] = "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空"
OBS = ["目標價 ¥1,726（DCF 1,015、Comparable 2,436 各半），比現價 1,968（2026-10-23）低 12.3%，級距 0%～−25%；Comparable 單獨為 +23.8%。DCF 只有現價的 0.52 倍：ROE 近三年平均 10.9% 只略高於折現率 10%，再投資率 36.7%，股東現金流只有 EPS 的六成多。",
       "本模型 12 個月後前瞻 EPS 94.4 比共識前瞻 EPS 116.0 低 18.6%；共識目標價 3,300 隱含 28.5 倍，同業中位數 25.8 倍（倍數因素 +10.2%）。情境 B、C 的 FY2027 EPS 比 A 低 12.5%、22.6%（AI 封膠市占流失 3pt／8pt 與上游下修）。",
       "FY2027 EPS 對堆疊層數（低值 → +8.9%）、淨利 ÷ 營業利益（0.65／0.72 → −4.9%／+5.3%）與 FY2026 錨點（−5%／+8% → −3.5%／+5.7%）最敏感。下一個檢驗事件是第二季財報（預計 2026-11 上旬，日期未確認），FY2026 錨點將改為上半年實績加下半年指引。"]   # 重點觀察（與 towa/README.md 相同文字）
_R = RD.max_row + 1
RD.cell(_R, 1, "v0.1.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁最後新增「上檔空間級距」「投資建議」兩列（依「上檔空間：加權目標價」固定規則）；README 頁最後新增投資建議與重點觀察三點。")
RD.cell(_R + 1, 1, f'="投資建議："&Val!B{_G + 1}&"（依上檔空間固定規則）"')
for _i, _t in enumerate(OBS, 1): RD.cell(_R + 1 + _i, 1, f"重點觀察 {_i}：{_t}")
if os.environ.get("G1_ROWS"): print("G1ROWS", "towa", "grade", _G, "rec", _G + 1, "readme_obs", _R + 2, _R + 4)

order = ["README", "In", "Q", "Hist", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(sys.argv[1])
if os.environ.get("TOWA_ROWS"):
    json.dump(dict(IN=IN_ROW, SUMR={f"{a}|{b}": v for (a, b), v in SUMR.items()}, PB=PB, EXT=EXT, T=T, EPSR=EPSR, PL_CHK=PL_CHK,
                   PCT=PCT_ROWS, POS=POS, VAL=VAL, V4=(VH, V4E), CUR=CUR, CMH=CMH, OLD=(OLD, OLDE), CYC=(CYCR, CYH), P1=p1, SENS_LAST=SENS_LAST, RC_LAST=RC_LAST),
              open(os.environ["TOWA_ROWS"], "w"), ensure_ascii=False)

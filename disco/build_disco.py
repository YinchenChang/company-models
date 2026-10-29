# Disco（6146.T）收支與估值模型 v0.2.1（AI 半導體 Project 步驟 5：設備與材料——第二家公司）
# 用法：python3 disco/build_disco.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度 4 月到隔年 3 月：FY2025＝2025-04～2026-03（公司「2026年3月期」）。金額：十億日圓（¥B）。
# 結構比照 advantest/build_advantest.py：In、Q、Hist、Drv（驅動推導）、PL、Val、Recon、Sens、Sources、README。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
import market as MEM
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
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
SHIP = {}   # HBM 出貨位元（EB），三情境各一列（markets/memory v0.4：2028–29 隨情境顆數係數變動）
for r in ws_.iter_rows():
    lab = str(r[0].value or "")
    if lab.startswith("出貨＝MIN"):
        SHIP[lab[-2]] = ([r[2 + j].value for j in range(5)], f"{rel(SKH_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert sorted(SHIP) == ["A", "B", "C"] and all(v is not None for s_ in SHIP for v in SHIP[s_][0])

ws_ = openpyxl.load_workbook(TSMC_F, data_only=True)["CoWoS"]
COW = None   # 台積電模型 CoWoS 需求合計（千片；2026–2029；2028–29 為情境 A 顆數）
for r in ws_.iter_rows():
    if r[0].value == "需求合計":
        COW = ([r[1 + j].value for j in range(4)], f"{rel(TSMC_F)}｜CoWoS!B{r[0].row}:E{r[0].row}")
assert COW and all(v is not None for v in COW[0])
ws_ = openpyxl.load_workbook(TSMC_F, data_only=True)["In"]
for r in ws_.iter_rows(min_row=4):   # 情境係數與 markets/memory 的 SCEN_UNITS 必須一致（單一來源）
    k = r[0].value
    if k and k[:4] in ("um28", "um29"):
        assert abs(r[4].value - MEM.SCEN_UNITS[k[-1]][0 if k[:4] == "um28" else 1]) < 1e-12, k

ws_ = openpyxl.load_workbook(EQ_F, data_only=True)["Mkt"]
PKG = None
for r in ws_.iter_rows():
    if r[0].value == "組裝與封裝設備":
        PKG = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert PKG
PEAK, PEAK_SRC = CP.peak_years(ROOT)   # 設備股共用循環位置規則：各情境見頂年度（markets/equipment）

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億日圓（¥B）；HBM 位元：EB；CoWoS：千片。每個輸入附標記與來源；找不到者寫「未找到」，公司未揭露者寫「公司不揭露」"
IN["A2"] = "標記：Verified（公司文件、交易所資料）／Interested-party（寫明是誰、立場）／Analogy（寫明對照與差異）／Assumed（一律給低／高區間）／Derived（寫明換算）。SemiAnalysis 未使用。公司原文（disco.co.jp）以網頁擷取工具讀取文字（建置環境的 curl 被擋）；表格欄位錯置處以算術核對，原文版面待人工複核"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t))
def inp(*a):   # Siliconomics v0.7.1 起，行業輸入（鍵加前綴 hbm_）改讀本庫快照
    rows.append((a[0], *SIL.row9("hbm_" + a[0])[1:]) if (a[0].startswith(("den_", "stk_", "hbp_"))) and ("hbm_" + a[0]) in SIL.V else a)
CO = "公司"
FR425 = "公司 2026年3月期 決算短信（2026-04-22，fr20260422.pdf）"
FR1022 = "公司 2027年3月期 第2四半期決算短信（2026-10-22，fr20261022.pdf）"
PR1022 = "公司 FY2026 2Q 決算說明資料（2026-10-22，film/20261022.pdf）"
PR422 = "公司 FY2025 4Q 決算說明資料（2026-04-22，film/20260422.pdf）"
MATSUI = "松井證券 決算情報（finance.matsui.co.jp/stock/6146/settlement）"

sec("A. FY2025 實績（2025-04～2026-03）與產品別構成（校準錨）")
inp("a25_rev", "營收 FY2025", "¥B", None, 436.889, None, "Verified", FR425 + "：436,889 百萬日圓", MATSUI + "：436,889；另見 kabuyoho 轉載之 TSE 決算短信 PDF（2026-04-22）")
inp("a25_gp", "毛利 FY2025", "¥B", None, 306.477, None, "Verified", FR425 + "：毛利 306,477（營收 − 銷貨成本 130,412）", PR422 + "：毛利 306.5、毛利率 70.1%")
inp("a25_sga", "銷管費 FY2025", "¥B", None, 121.487, None, "Verified", FR425 + "：121,487", PR422 + "：121.5")
inp("a25_oi", "營業利益 FY2025", "¥B", None, 184.989, None, "Verified", FR425 + "：184,989", MATSUI + "：184,989")
inp("a25_ni", "淨利（歸屬母公司）FY2025", "¥B", None, 135.521, None, "Verified", FR425 + "：135,521", MATSUI + "：135,521")
inp("a25_eps", "基本 EPS FY2025", "¥", None, 1249.84, None, "Verified", FR425 + "：基本 1,249.84（稀釋 1,245.90）", MATSUI + "：1,249.84")
inp("mix_eq", "精密加工裝置占營收 FY2025", "%", 0.615, 0.62, 0.625, "Verified", PR422 + "：製品別構成 精密加工裝置 62%（ダイサ 31%、グラインダ 28%、附屬裝置 3%）、精密加工ツール 21%、その他 16%（合計 99%，四捨五入）", "公司只揭露構成比（百分比取整數），不揭露金額；低／高＝±0.5 個百分點（取整誤差）。FY2024：63%／21%／16%（film/20250417）")
inp("mix_tl", "精密加工ツール（消耗品：刀片、砂輪）占營收 FY2025", "%", 0.205, 0.21, 0.215, "Verified", "同上：21%（FY2024 21%、FY2023 22%）", "同上；消耗品隨客戶稼動率變動（公司 FY26 Q1 說明：稼動率高，創新高）")
inp("eq_25", "精密加工裝置營收 FY2025（營收 × 構成比）", "¥B", None, "=a25_rev*mix_eq", None, "Derived", "436.9 × 62%", "")
inp("tl_25", "精密加工ツール營收 FY2025", "¥B", None, "=a25_rev*mix_tl", None, "Derived", "436.9 × 21%", "")
inp("ot_25", "その他（零件、服務、二手機等）FY2025", "¥B", None, "=a25_rev-eq_25-tl_25", None, "Derived", "營收 − 裝置 − 消耗品（＝17%，含取整差）", "公司構成比為 16%")
inp("r_ot", "その他 ÷（裝置＋消耗品）", "x", None, "=ot_25/(eq_25+tl_25)", None, "Derived", "FY2025 比例，各年沿用", "FY2024 16/84＝0.19，相同")

sec("B. AI 錨點（校準用；不使用共識營收或公司目標）")
inp("ai_25", "生成 AI 相關出貨額 FY2025（2026年3月期）", "¥B", 67.5, 75.0, 82.5, "Interested-party", "樂天證券 今中能夫（2026-04-27，決算レポート；券商分析師，有正面撰寫傾向）：2026年3月期 約 750 億、2025年3月期 約 550 億（出貨額）", "第二來源：未找到金額（公司不揭露 AI 金額）。低／高 ±10%＝「約」的不確定（Assumed）。出貨額視同營收：FY2025 出貨 442.6（四季合計，公司簡報）÷ 營收 436.9＝1.013（Derived）。是否含消耗品：原文不明，本模型視為裝置")
inp("w_hbm", "AI 出貨中 HBM 比重 FY2025（其餘為邏輯／先進封裝）", "%", 0.45, 0.48, 0.52, "Interested-party", "今中能夫（同上）：2025年3月期 概算 邏輯 3：HBM 7；2026年3月期 HBM 持平、邏輯大增，邏輯略高於 HBM → HBM 約 48%（Derived：38.5÷75≈51%，再依「邏輯略高」取 48%）", "第二來源：公司 FY26 2Q 法說（BigGo 摘要，2026-10-22）：全年約 HBM 一半、邏輯一半。Interested-party（公司）")
inp("ai_24", "生成 AI 相關出貨額 FY2024（只作參考，未用於校準）", "¥B", None, 55.0, None, "Interested-party", "今中能夫（同上）：約 550 億", "回測需 CY2024 的 HBM 位元與 CoWoS 片數，市場層沒有，未做（見 README 待辦）")

sec("C. HBM 研削／切割需求（markets/memory，情境 A／B／C 各一列；唯讀）")
for s in SC:
    for j, y in enumerate((25, 26, 27, 28, 29)):
        inp(f"hs{s}_{y}", f"HBM 出貨位元 20{y}（EB；{s}）", "EB", None, SHIP[s][0][j], None, "Derived", SHIP[s][1], "markets/memory v0.4（market.py @ 5fa3751）：出貨＝MIN(需求 × 情境顆數係數, 供給)；2028–29 隨情境，2025–27 三情境相同" if j == 0 else "同上")
for y, (lo, b, hi) in zip((25, 26, 27, 28, 29, 30), ((22, 24, 24), (24, 24, 26), (24, 26, 28), (24, 28, 32), (26, 30, 32), (28, 32, 32))):
    inp(f"den_{y}", f"DRAM 核心晶粒容量（組合平均）20{y}", "Gb／顆", lo, b, hi, "Assumed", "JEDEC HBM4（JESD270-4）：24Gb 或 32Gb 晶粒（TrendForce 2025-04-28）；HBM3E 以 24Gb 為主", "組合比例未找到；容量愈大，同樣位元需要的晶粒（研削＋切割次數）愈少" if y == 25 else "同上")
for y, (lo, b, hi) in zip((25, 26, 27, 28, 29, 30), ((8, 10, 12), (10, 12, 12), (12, 12.5, 14), (12, 14, 16), (12, 15, 16), (14, 16, 16))):
    inp(f"stk_{y}", f"HBM 平均堆疊層數 20{y}", "層", lo, b, hi, "Assumed", "JEDEC HBM4：4／8／12／16 層（TrendForce 2025-04-28）；HBM3E 8／12 層；Rubin Ultra 16 層 HBM4E（台積電模型 PP 表）", "組合比例未找到；每多一層 DRAM＝多一顆晶粒要薄化與切割" if y == 25 else "同上")
inp("s_stk", "每個堆疊額外的研削／切割工序（基底晶粒、模封研削、堆疊切割）", "次／堆疊", 2, 3, 4, "Assumed", "公司技術說明會 2025（2025-12）流程圖：HBM 晶粒薄化→切割→CoW 搭載→研削→切割（無數字）", "第二來源：未找到。每堆疊工序＝passes ＝ 晶粒數 ×（1＋s_stk ÷ 層數）")
inp("dlt", "已裝機產能每年汰換／升級比例（新機需求＝產能增量＋汰換）", "%/年", 0.05, 0.10, 0.15, "Assumed", "無：公司不揭露裝機量與機齡", "與校準係數 k 一起只能識別乘積的形狀；Sens 列 5%／15%")
inp("g_tail", "CY2030–31 研削需求（HBM 工序、CoWoS 片數）年增（市場層只到 2029）", "%", 0.0, 0.10, 0.20, "Assumed", "無：市場層只到 2029；設備在產能需要前一年購入，FY2029 需要 CY2030–31", "")

sec("D. 先進封裝（CoWoS 中介層等）需求：台積電模型 CoWoS 需求（唯讀）")
inp("cw_25", "CoWoS 需求 CY2025", "千片", 600, 670, 700, "Interested-party", "36Kr／半導體產業縱橫（2025-12-04）：2025 全球 CoWoS 需求約 67 萬片", "第二來源：TrendForce（2024-10-21、2025-01-02）月產能 2024 年底 3.5–4 萬、2025 年底約 7.5–8 萬 → 年均約 5.5 萬 × 12 ≈ 66 萬（Derived）。低／高 Assumed")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"cw_{y}", f"CoWoS 需求 CY20{y}（台積電模型，情境 A 顆數）", "千片", None, COW[0][j], None, "Derived", COW[1], "台積電模型 v0.2：加速器顆數 ÷ 每片封裝數＋其他用途；本模型 ÷ Morgan Stanley：2026 −21%、2027 −16%（台積電模型 Recon）" if j == 0 else "同上")
for s in SC:
    for n, y in enumerate((28, 29)):
        inp(f"um{y}_{s}", f"情境 {s}：20{y} 顆數 ÷ 基準（CoWoS 片數同比例）", "x", None, MEM.SCEN_UNITS[s][n], None, "Assumed", "markets/memory/market.py SCEN_UNITS（單一來源；台積電、Advantest、HBM 位元用同一組）", "台積電模型的 CoWoS 頁只有情境 A；B、C 以同一係數縮放（含其他 CoWoS 用途，近似）" if (s == "A" and y == 28) else "同上")

sec("E. 混合鍵合（hybrid bonding）情境開關")
inp("hb_on", "混合鍵合情境開關（1＝納入、0＝不納入）", "0/1", 0, 1, 1, "Assumed", "使用者可切換；Sens 列「開關 0」", "")
for y, (lo, b, hi) in zip((25, 26, 27, 28, 29, 30), ((0, 0, 0), (0, 0, 0), (0, 0, 0.05), (0, 0.05, 0.15), (0.05, 0.15, 0.30), (0.10, 0.25, 0.40))):
    inp(f"hbp_{y}", f"混合鍵合占 HBM 研削工序比例 20{y}", "%", lo, b, hi, "Assumed", "TrendForce（2026-03-06，引 ZDNet）：尚未用於 HBM 量產；三星最早在 16 層 HBM4E 部分採用；JEDEC 討論把高度上限放寬到 825–900µm，可能延後混合鍵合" if y == 25 else "同上", "第二來源：Referently（2026-10-06）：混合鍵合仍在導入初期" if y == 25 else "")
inp("hb_eff", "混合鍵合對每道 HBM 工序裝置支出的淨影響（＋＝增加）", "%", -0.20, 0.10, 0.30, "Assumed", "增加的方向：公司已開發混合鍵合專用研削機（超高精度、潔淨、晶圓形狀修正、機內化學清洗；技術說明會 2025-12，Interested-party）；混合鍵合需加做 CMP、清洗與檢查（Referently 2026-10-06），且需鍵合後研削、露墊、氧化膜研削（公司 2025）", "減少的方向：D2W 混合鍵合要求無微粒切割，電漿切割較適合，可能取代部分刀片／雷射切割（SPTS 撰文，3D InCites，競爭對手，Interested-party）；去除凸塊與底填可能省部分步驟。本模型假設研削精度升級大於切割被取代：淨 +10%（低 −20%、高 +30%）")

sec("F. 非 AI 後段需求（功率半導體、手機、車用、中國後段）：Analogy＝公司 AI 前的歷史 FY2015–FY2022")
inp("g_non", "非 AI 裝置營收年增（情境 A；B、C 只在 FY2028 不同）", "%", 0.0, 0.06, 0.12, "Analogy", "公司 AI 前歷史（Hist 頁）：營收 FY2015→FY2022 年複合 12.1%（含 2020–21 中國後段與功率熱潮）；FY2016→FY2019 1.7%（循環平均）。基準 6%＝兩者約中點；低 0%、高 12%", "與現在的差異：(1) 2020–22 中國 OSAT 投資高峰；FY26 Q1 中國 OSAT 再度大增、公司質疑可持續多久（BigGo 法說摘要）；(2) 功率半導體 FY2024 因 EV 放緩下滑（公司 film/20250417：其他半導體 4Q 年減 37–48%），FY26 Q1 回升，占營收不到一成（法說摘要）；(3) 手機、PC 需求 FY2022 起疲弱（公司 film/20230420）")
inp("g_nonB_28", "情境 B：FY2028 非 AI 裝置年增", "%", -0.05, 0.0, 0.05, "Assumed", "AI 資本支出消化期，非 AI 也持平", "")
inp("g_nonC_28", "情境 C：FY2028 非 AI 裝置年增（循環下行）", "%", -0.20, -0.12, -0.04, "Analogy", "公司 FY2018→FY2019（2019年3月期）營收 −11.9%（Hist 頁：167.4→147.5），記憶體與中國投資下行", "第二來源：FY2019→FY2020 −4.3%（147.5→141.1）。差異：當時 AI 尚無貢獻；現在非 AI 裝置占營收約 45%")
inp("g_tl", "非 AI 消耗品年增（隨稼動率，不隨新機）", "%", 0.0, 0.05, 0.10, "Analogy", "公司構成比推得 FY2022 0%、FY2023 +11%、FY2024 +13%、FY2025 +9%（各年說明資料的年增率，含匯率與 AI）", "基準取 0–10% 中點；情境 C FY2028 取 0%（下方）")
inp("g_tlC_28", "情境 C：FY2028 非 AI 消耗品年增", "%", -0.05, 0.0, 0.05, "Assumed", "公司：FY2019 下行期消耗品出貨仍創新高（film/20200423）", "")
inp("w_tlai", "FY2025 消耗品中屬 AI 工序的比例", "%", 0.10, "=ai_25/eq_25", 0.40, "Assumed", "公司不揭露；基準＝AI 占裝置營收的比例（約 28%）——假設裝機組合等於當年出貨組合", "AI 消耗品隨 AI 工序量（稼動）而非新機變動")

sec("G. 市占（只用於「市占變化」的情境，校準已含 FY2025 水準）")
inp("sh", "Disco 研削／切割裝置市占（2025）", "%", 0.60, 0.70, 0.80, "Interested-party", "Morningstar（2025）：60–70%（未引一手調查）", "第二來源：現代ビジネス（2025-12-30，田宮寬之，東洋經濟）：研削、切割 7–8 成（未引出處）；Desk Research Design（note）70–80%。三者都不是一手市場調查")
inp("sh_loss", "FY2029 市占下滑（相對 2025，線性）", "pt", 0.0, 0.0, 0.10, "Assumed", "基準 0；Sens 列 −10 個百分點（東京精密、中國本土廠）", "")

sec("H. 利潤率（由 Hist 頁 FY2015–FY2025 迴歸推得，不寫死）")
inp("gm_b", "毛利率斜率（毛利率 ＝ a ＋ b × ln 營收）", "x", None, "=SLOPE(Hist!H6:H16,Hist!I6:I16)", None, "Derived", "Hist 頁 11 年（FY2015–FY2025）最小平方法", "規模與產品組合（高附加價值機種）一起推高毛利率；公司 FY2023 說明：主因匯率與高附加價值產品")
inp("gm_a", "毛利率截距", "%", None, "=INTERCEPT(Hist!H6:H16,Hist!I6:I16)", None, "Derived", "同上", "")
inp("gm_b2", "毛利率斜率（只用 FY2019–FY2025，Sens 用）", "x", None, "=SLOPE(Hist!H10:H16,Hist!I10:I16)", None, "Derived", "Hist 頁 7 年", "")
inp("gm_a2", "毛利率截距（FY2019–FY2025，Sens 用）", "%", None, "=INTERCEPT(Hist!H10:H16,Hist!I10:I16)", None, "Derived", "同上", "")
inp("gm_adjC", "情境 C：FY2028–29 毛利率額外調整（稼動下滑、價格）", "pt", -0.03, -0.01, 0.0, "Assumed", "無", "")
inp("sg_v", "銷管費變動部分（每 1 日圓營收）", "x", None, "=SLOPE(Hist!D6:D16,Hist!B6:B16)", None, "Derived", "Hist 頁 11 年：銷管費 ＝ 固定 ＋ 變動 × 營收", "FY2026 研發約 36.0（公司 4 月說明）")
inp("sg_f", "銷管費固定部分", "¥B", None, "=INTERCEPT(Hist!D6:D16,Hist!B6:B16)", None, "Derived", "同上", "")
inp("etr_25", "有效稅率 FY2025（所得稅 ÷ 稅前淨利）", "%", None, 0.26226, None, "Verified", FR425 + "：所得稅 48,207 ÷ 稅前淨利 183,811 ＝ 26.2%", "FY2024 26.2%（44,070 ÷ 168,146）、FY2023 26.5%（30,364 ÷ 114,576）（fr20250417.pdf）：三年穩定")
inp("etr_h1", "有效稅率 FY2026 上半（期中）", "%", None, 0.29249, None, "Verified", FR1022 + "：所得稅 33,271 ÷ 稅前淨利 113,753 ＝ 29.2%", "FY2025 上半 28.8%（22,606 ÷ 78,571）而全年 26.2%：期中稅率歷來高於全年")
inp("def_t", "防衛特別法人稅對有效稅率的影響（2026-04-01 以後開始的事業年度）", "pt", None, "=0.232*0.04", None, "Derived", "法人稅率 23.2% × 防衛特別法人稅 4%（扣除 500 萬日圓基礎扣除額後課徵；適用 2026-04-01 以後開始的事業年度）＝ 約 +0.9 個百分點", "來源：安田會計事務所（2026-04-22）、freee 法人稅 2026 說明。公司未在短信中說明稅負變化")
inp("etr", "有效稅率（FY2026 第四季起）", "%", "=etr_25", "=etr_25+def_t", "=etr_h1", "Derived", "基準＝FY2025 26.2% ＋ 防衛特別法人稅 0.9pt ＝ 27.1%；低＝FY2025 水準；高＝FY2026 上半期中稅率 29.2%", "")
inp("pbt_oi", "稅前淨利 ÷ 營業利益（營業外＋特別損益）", "x", 0.97, 0.99, 1.00, "Derived", "FY2025 0.994（183,811 ÷ 184,989）；FY2024 1.008；FY2026 上半 0.975（含額外退職金 3,101、減損 636、匯損 486）", "上半扣除額外退職金與減損後 1.006。退職金公司未說明是否一次性；基準 0.99、低 0.97（上半水準）、高 1.00")
inp("nioi", "淨利 ÷ 營業利益（FY2026 第四季起）", "x", "=pbt_oi_lo*(1-etr_hi)", "=pbt_oi*(1-etr)", "=pbt_oi_hi*(1-etr_lo)", "Derived", "＝（稅前 ÷ 營業利益）×（1 − 有效稅率）；少數股權約 0.1B 忽略。基準 0.722；低 0.687（≈ 上半實績 0.689）；高 0.738（≈ FY2025 0.733）", "v0.1 用 FY2025 的 0.733。上半 0.689 低的原因：期中稅率偏高（29.2%）、額外退職金 3.1B、減損 0.6B；扣除後 0.712。公司第三季指引隱含 0.714（41.5 ÷ 58.1）")
inp("nioi_v01", "v0.1 的淨利 ÷ 營業利益（只用於「驅動推導」對照列）", "x", None, "=a25_ni/a25_oi", None, "Derived", "FY2025 實績 0.733", "")
inp("shares", "加權平均股數", "百萬股", None, 108.480881, None, "Verified", FR1022 + "：上半期平均 108,480,881 股", "FY2025 平均 108,431,323（" + FR425 + "）；庫藏股 16,004")

sec("I. FY2026 實績、公司指引（只給下一季）與對帳參考")
inp("q26h1_rev", "FY2026 上半營收（2026-04～09）", "¥B", None, 257.058, None, "Verified", FR1022 + "：257,058（+32.1%）", MATSUI + "：257,058")
inp("q26h1_oi", "FY2026 上半營業利益", "¥B", None, 116.639, None, "Verified", FR1022 + "：116,639（+47.9%）", MATSUI)
inp("q26h1_ni", "FY2026 上半淨利", "¥B", None, 80.366, None, "Verified", FR1022 + "：80,366（+43.7%）", MATSUI)
inp("q26h1_eps", "FY2026 上半 EPS（基本）", "¥", None, 740.84, None, "Verified", FR1022 + "：740.84", MATSUI)
inp("q26h1_gp", "FY2026 上半毛利", "¥B", None, 183.917, None, "Verified", FR1022 + "：183,917（毛利率 71.5%）", PR1022 + "：183.9")
inp("g9_rev", "公司指引：FY2026 前三季累計營收（2026-04～12）", "¥B", None, 389.5, None, "Interested-party", FR1022 + "：9 個月累計預測 389,500（+28.2%）", "BigGo 法說摘要（2026-10-22）同值；公司只給下一季指引")
inp("g9_oi", "公司指引：前三季累計營業利益", "¥B", None, 174.7, None, "Interested-party", FR1022 + "：174,700", "同上")
inp("g9_ni", "公司指引：前三季累計淨利", "¥B", None, 121.9, None, "Interested-party", FR1022 + "：121,900；EPS 1,123.43", "同上")
inp("g3_rev", "公司指引：第三季（10–12 月）營收", "¥B", None, "=g9_rev-q26h1_rev", None, "Derived", "累計 389.5 − 上半 257.1 ＝ 132.4", PR1022 + "：3Q 預測 132.4、營業利益 58.1、淨利 41.5、出貨 149.6；匯率 ¥151")
inp("g3_oi", "公司指引：第三季營業利益", "¥B", None, "=g9_oi-q26h1_oi", None, "Derived", "174.7 − 116.6 ＝ 58.1", "同上")
inp("g3_ni", "公司指引：第三季淨利", "¥B", None, "=g9_ni-q26h1_ni", None, "Derived", "121.9 − 80.4 ＝ 41.5", PR1022 + "：41.5")
inp("q26h1_sga", "FY2026 上半銷管費", "¥B", None, 67.277, None, "Verified", FR1022 + "：67,277", PR1022 + "：67.3")
inp("gm_q3", "第三季毛利率（公司說法：比第二季 71.7% 低不到 2 個百分點）", "%", 0.690, 0.698, 0.705, "Interested-party", "公司 FY26 2Q 法說（BigGo 摘要 2026-10-22）：約 1 點來自匯率、不到 1 點來自其他", "第二季 71.7%：" + PR1022 + "。第四季毛利率假設同第三季（Assumed）")
inp("q2_sga", "FY2026 第二季銷管費", "¥B", None, 34.8, None, "Verified", PR1022 + "：34.8（Q1 32.5）", "公司：第三季銷管費大致持平（法說摘要，BigGo 2026-10-22）")
inp("g_q4", "第四季營收 ÷ 第三季營收 − 1（FY2026 錨點）", "%", 0.0, 0.10, 0.22, "Assumed", "公司過去三年第四季對第三季：FY2023 +35.5%、FY2024 +29.0%、FY2025 +21.7%（Q 頁，由公司累計數推得，驗收集中在 3 月）", "打折的理由（公司說法）：第三季出貨僅比第二季 +1%、HBM 暫歇、記憶體占比由近 40% 降到約 30%、第三季營收比第二季 −7%。基準 +10%（約三年最低值的一半）；低 0%（持平，即 v0.1 對帳的假設）；高 +22%（FY2025 實績）")
inp("m_q4", "第四季營業利益率（FY2026 錨點）", "%", "=g3_oi/g3_rev-0.02", "=g3_oi/g3_rev", "=g3_oi/g3_rev+0.02", "Assumed", "基準＝第三季指引 43.9%；± 2 個百分點", "公司過去三年第四季營益率都高於第三季（+0.9 至 +4.8 個百分點，Q 頁）；基準不加季節性")
inp("anc_carry", "錨點差額延續到 FY2027 以後的非 AI 裝置與非 AI 消耗品（1＝延續、0＝不延續）", "0/1", 0, 1, 1, "Assumed", "FY2026 錨點營收 ÷ 驅動推導營收 的差額，平均分到各組成；AI 裝置與 AI 消耗品 FY2027 起回到驅動推導（不延續），非 AI 兩項以 FY2026 錨定水準為基期繼續成長", "Sens 列「不延續」")
inp("cons_rev26", "共識營收 FY2026（2027年3月期）", "¥B", 541.0, 564.96, 598.7, "Interested-party", "stockanalysis.com（2026-10-07，20 位分析師平均；低／高＝區間）", "第二來源：樂天證券 今中能夫 2026-04-27 預估 550（營業利益 240）")
inp("cons_oi26", "樂天證券預估營業利益 FY2026", "¥B", None, 240.0, None, "Interested-party", "今中能夫（2026-04-27）", "共識營業利益：未找到（stockanalysis 付費）")
inp("cons_rev27", "樂天證券預估營收 FY2027（2028年3月期）", "¥B", None, 640.0, None, "Interested-party", "今中能夫（2026-04-27）：營收 640、營業利益 290、EPS 1,962.0", "stockanalysis 下一年 671.91（2026-10-07）")
for s in SC:
    inp(f"pk_{s}", f"情境 {s}：設備市場見頂年度（日曆年；循環位置規則用）", "年", None, PEAK[s], None, "Derived", PEAK_SRC[s], "設備股共用規則（tools/cycle_position.py）：FY(t) 早於見頂年＝中段、等於＝見頂、晚於＝谷底；設備市場模型的情境係數為 Assumed" if s == "A" else "同上")
inp("px", "股價（2026-10-23 收盤）", "¥", None, 49290, None, "Verified", "株探 日足（2026-10-23 終值 49,290）", "kabutore（同值）；markets/invest_map/data/prices.csv 兩來源核對")
inp("sa_fpe", "遠期 P/E（stockanalysis，2026-10-21）", "x", None, 28.33, None, "Interested-party", "stockanalysis.com statistics（股價 53,570）；EPS 年度未標明", "隱含遠期 EPS 約 1,891（Derived）")
inp("sa_ttm", "TTM P/E（stockanalysis，2026-10-21）", "x", None, 39.95, None, "Interested-party", "stockanalysis.com（TTM EPS 1,340.86，未含 FY26 Q2）", "")
inp("semi_pk", "SEMI 組裝與封裝設備 CY2025／CY2026（markets/equipment）", "US$B", "=" + str(PKG[0][0]), PKG[0][1], None, "Interested-party", PKG[1], "SEMI 2026-07；D 欄＝2025、E 欄＝2026；只作對照")
# v0.1.4：目標價情境權重（放在 In 頁最後，不位移既有列）
sec("J. 目標價情境權重（Val 頁「0. 目標價與上檔空間」用；Andy 可改）")
for s_, w_ in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s_}", f"目標價情境權重 {s_}", "%", None, w_, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")

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
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    defname(key, f"In!$E${r}")
    if lo is not None: defname(key + "_lo", f"In!$D${r}")
    if hi is not None: defname(key + "_hi", f"In!$F${r}")
    r += 1
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 80, 80]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

# ---------------- Hist：FY2015–FY2025 年度實績（公司決算短信） ----------------
H = wb.create_sheet("Hist")
H["A1"] = "Hist — 年度實績（十億日圓；EPS 為日圓）。來源1：公司各年決算短信（英文版 fr*.pdf，以網頁擷取讀取）；來源2：松井證券（FY2021 起）。用途：利潤率迴歸（In 頁 H 節）、非 AI 成長的 Analogy、10 年 P/E 的 EPS"
H["A2"] = "股票分割：2023-04-01 一股分三股（公司 FY2022 決算短信註記；tousaiblog 第二來源）。FY2021（2022年3月期）起的 EPS 已在公司文件中依分割調整；FY2020 以前除以 3"
hdr(H, 5, ["會計年度", "營收", "毛利", "銷管費", "營業利益", "淨利", "EPS（公司公布，日圓）", "毛利率", "ln 營收", "分割係數", "EPS（分割調整後）", "來源1（公司）", "來源2"])
HIST = [("FY2015（2016-03）", 127.850, 72.298, 41.959, 30.338, 23.096, 646.09, 3, "fr20170510.pdf", "未找到（松井只到 FY2021）"),
        ("FY2016（2017-03）", 134.204, 74.495, 43.153, 31.341, 24.203, 675.82, 3, "fr20170510.pdf（說明文字寫 134,240／24,230，取財務表數字）", "未找到"),
        ("FY2017（2018-03）", 167.364, 99.125, 48.130, 50.995, 37.171, 1035.67, 3, "fr20190508.pdf", "未找到"),
        ("FY2018（2019-03）", 147.500, 86.910, 48.264, 38.645, 28.824, 802.35, 3, "fr20190508.pdf", "未找到"),
        ("FY2019（2020-03）", 141.083, 84.792, 48.341, 36.451, 27.653, 769.56, 3, "fr20210422.pdf", "未找到"),
        ("FY2020（2021-03）", 182.857, 106.883, 53.776, 53.106, 39.091, 1085.47, 3, "fr20210422.pdf（說明文字營業利益 53,160，取表內 53,106）", "未找到"),
        ("FY2021（2022-03）", 253.781, 154.011, 62.498, 91.513, 66.206, 611.67, 1, "fr20230420.pdf（分割調整後 EPS）", MATSUI),
        ("FY2022（2023-03）", 284.135, 184.506, 74.093, 110.413, 82.891, 765.47, 1, "fr20230420.pdf", MATSUI),
        ("FY2023（2024-03）", 307.554, 208.642, 87.151, 121.490, 84.205, 777.29, 1, "fr20250417.pdf", MATSUI),
        ("FY2024（2025-03）", 393.313, 277.570, 110.736, 166.834, 123.891, 1143.26, 1, "fr20250417.pdf", MATSUI),
        ("FY2025（2026-03）", 436.889, 306.477, 121.487, 184.989, 135.521, 1249.84, 1, "fr20260422.pdf", MATSUI)]
for i, (fy, rv, gp, sg, oi, ni, eps, f, s1, s2) in enumerate(HIST, 6):
    H.cell(i, 1, fy)
    for j, v in enumerate((rv, gp, sg, oi, ni, eps), 2):
        H.cell(i, j, v).font = BLUE; H.cell(i, j).number_format = "#,##0.0" if j < 7 else "#,##0.00"
    H.cell(i, 8, f"=C{i}/B{i}").number_format = "0.0%"; H.cell(i, 9, f"=LN(B{i})").number_format = "0.000"
    H.cell(i, 10, f).font = BLUE; H.cell(i, 11, f"=G{i}/J{i}").number_format = "#,##0.00"
    H.cell(i, 12, s1); H.cell(i, 13, s2)
H.cell(18, 1, "檢查：FY2025 毛利 − 銷管費 − 營業利益（應為 0）"); H["B18"] = "=C16-D16-E16"; H["B18"].number_format = "0.000"
H.cell(19, 1, "檢查：FY2025 列 − In 頁 a25_*（應為 0）"); H["B19"] = "=ABS(B16-a25_rev)+ABS(C16-a25_gp)+ABS(E16-a25_oi)+ABS(F16-a25_ni)+ABS(G16-a25_eps)"; H["B19"].number_format = "0.000"
H.cell(21, 1, "非 AI 成長的 Analogy").font = BOLD
H.cell(22, 1, "營收年複合 FY2015→FY2022（AI 前，含中國與功率熱潮）"); H["B22"] = "=(B13/B6)^(1/7)-1"; H["B22"].number_format = "0.0%"
H.cell(23, 1, "營收年複合 FY2016→FY2019（一個循環）"); H["B23"] = "=(B10/B7)^(1/3)-1"; H["B23"].number_format = "0.0%"
H.cell(24, 1, "最大單年下滑 FY2018（2019-03）"); H["B24"] = "=B9/B8-1"; H["B24"].number_format = "0.0%"
H.cell(26, 1, "迴歸結果（In 頁同值）").font = BOLD
for i, (lab, f, fm) in enumerate((("毛利率 ＝ a ＋ b × ln 營收：a", "=gm_a", "0.0000"), ("b", "=gm_b", "0.0000"), ("R²", "=RSQ(H6:H16,I6:I16)", "0.000"),
                                  ("銷管費 ＝ 固定 ＋ 變動 × 營收：固定 ¥B", "=sg_f", "0.00"), ("變動", "=sg_v", "0.000"), ("R²", "=RSQ(D6:D16,B6:B16)", "0.000"),
                                  ("FY2025 毛利率：迴歸值 − 實績", "=gm_a+gm_b*I16-H16", "0.0%"), ("FY2025 銷管費：迴歸值 − 實績（¥B）", "=sg_f+sg_v*B16-D16", "0.0")), 27):
    H.cell(i, 1, lab); H.cell(i, 2, f).number_format = fm
H.column_dimensions["A"].width = 46
for c in "BCDEFGHIJK": H.column_dimensions[c].width = 12
H.column_dimensions["L"].width = 60; H.column_dimensions["M"].width = 50

# ---------------- Q：季度實績與指引 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 季度實績（十億日圓；EPS 為日圓）。會計年度 4–3 月。來源1：公司決算說明資料（film/20260422、20261022）與決算短信；來源2：松井證券季度表（百萬日圓）"
hdr(Q, 3, ["項目", "FY25 Q1", "FY25 Q2", "FY25 Q3", "FY25 Q4", "FY2025", "FY26 Q1", "FY26 Q2", "FY26 Q3（公司指引）", "標記", "來源1", "來源2／說明"])
QD = [("營收", [89.9, 104.623, 109.291, 133.061, None, 114.308, 142.750, "=g3_rev"], "公司說明資料：89.9／104.6／109.3／133.1；FY26 114.3／142.8；3Q 指引 132.4", "松井：104,623／109,291／133,061；FY26 114,308／142,750（Q1 FY25 由全年減其他三季）"),
      ("營業利益", [34.5, 44.391, 47.341, 58.777, None, 49.033, 67.606, "=g3_oi"], "公司說明資料：34.5／44.4／47.3／58.8；FY26 49.0／67.6；3Q 58.1", "松井：44,391／47,341／58,777；49,033／67,606"),
      ("淨利", [23.8, 32.146, 36.726, 42.882, None, 34.221, 46.145, "=g9_ni-q26h1_ni"], "公司說明資料：23.8／32.1／36.7／42.9；FY26 34.2／46.1；3Q 41.5", "松井：32,146／36,726／42,882；34,221／46,145"),
      ("EPS（基本）", [None, 296.49, 338.73, 395.48, None, 315.51, 425.38, "=1123.43-q26h1_eps"], "公司決算短信：FY26 Q1 315.51、上半 740.84；9 個月指引 1,123.43", "松井：296.49／338.73／395.48；315.51／425.38"),
      ("出貨額", [111.1, 96.3, 113.6, 121.6, None, 135.9, 148.6, 149.6], "公司說明資料", "Investing.com（2026-10-22）：Q2 148.6、Q3 指引 149.6")]
for i, (lab, vals, s1, s2) in enumerate(QD, 4):
    Q.cell(i, 1, lab)
    for j, v in enumerate(vals):
        col = "BCDEGHI"[j] if False else ["B", "C", "D", "E", "F", "G", "H", "I"][j]
        if v is None: continue
        c = Q[f"{col}{i}"]; c.value = v; c.number_format = "#,##0.0" if lab != "EPS（基本）" else "#,##0.00"
        if not isinstance(v, str): c.font = BLUE
    Q.cell(i, 10, "Verified／Interested-party（指引）"); Q.cell(i, 11, s1); Q.cell(i, 12, s2)
Q["B4"] = "=a25_rev-C4-D4-E4"; Q["B5"] = "=a25_oi-C5-D5-E5"; Q["B6"] = "=a25_ni-C6-D6-E6"; Q["B7"] = "=a25_eps-C7-D7-E7"
for rr in (4, 5, 6, 7, 8):
    Q[f"F{rr}"] = f"=SUM(B{rr}:E{rr})"; Q[f"F{rr}"].number_format = Q[f"C{rr}"].number_format
Q.cell(9, 1, "營業利益率"); [Q.cell(9, j, f"={c}5/{c}4").__setattr__("number_format", "0.0%") for j, c in zip(range(2, 10), "BCDEFGHI")]
Q.cell(10, 1, "毛利率（公司說明資料；FY25 Q1、Q3 未列單季毛利額）");
for c, v in (("C", 0.695), ("E", 0.709), ("F", "=a25_gp/a25_rev"), ("G", 0.713), ("H", 0.717)):
    Q[f"{c}10"] = v; Q[f"{c}10"].number_format = "0.0%"
    if not isinstance(v, str): Q[f"{c}10"].font = BLUE
Q["L10"] = "FY25 Q3 71.4%（4Q 資料的前季對照）；3Q 指引：比 Q2 低約 2 個百分點（法說摘要）"
Q.cell(11, 1, "檢查：FY2026 上半（G+H）− 決算短信上半（營收、營業利益、淨利；應約為 0）")
Q["B11"] = "=G4+H4-q26h1_rev"; Q["C11"] = "=G5+H5-q26h1_oi"; Q["D11"] = "=G6+H6-q26h1_ni"
for c in "BCD": Q[f"{c}11"].number_format = "0.000"
Q.cell(13, 1, "製品別構成比（公司說明資料；出貨或營收基準，百分比取整）").font = BOLD
hdr(Q, 14, ["製品", "FY2023", "FY2024", "FY2025", "FY26 Q2", "", "", "", "", "標記", "來源", ""])
MIX = [("精密加工裝置", 0.64, 0.63, 0.62, 0.63), ("　ダイサ", 0.32, 0.33, 0.31, 0.33), ("　グラインダ", 0.28, 0.27, 0.28, 0.26), ("精密加工ツール（消耗品）", 0.22, 0.21, 0.21, 0.23), ("その他", 0.14, 0.16, 0.16, 0.14)]
for i, (nm, *v) in enumerate(MIX, 15):
    Q.cell(i, 1, nm)
    for j, x in enumerate(v, 2): Q.cell(i, j, x).number_format = "0%"; Q.cell(i, j).font = BLUE
    Q.cell(i, 10, "Verified"); Q.cell(i, 11, "film/20240425、20250417、20260422、20261022")
Q.cell(21, 1, "公司揭露的季度結構（法說摘要，BigGo 2026-07-23／10-22；Interested-party）：記憶體占營收 FY26 Q1 35%、Q2 略低於 40%、Q3 預估約 30%；OSAT 占 FY26 Q1 35%；功率半導體不到一成；FY26 Q2 切割機用途 IC 73%、研削機用途 IC 77%")
Q.column_dimensions["A"].width = 52
for c in "BCDEFGHI": Q.column_dimensions[c].width = 12
Q.column_dimensions["J"].width = 20; Q.column_dimensions["K"].width = 70; Q.column_dimensions["L"].width = 70

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
CYC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2032)}   # C＝CY2025 … I＝CY2031
FYC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2030)}   # C＝FY2025 … G＝FY2029
DEFP = dict(dlt="dlt", sstk="s_stk", tail="g_tail", den="", stk="", hbon="hb_on", hbp="", hbe="hb_eff", cw25="cw_25",
            ai="ai_25", w="w_hbm", gn="g_non", gnB="g_nonB_28", gnC="g_nonC_28", gtl="g_tl", gtlC="g_tlC_28", wtl="w_tlai",
            shl="sh_loss", ga="gm_a", gb="gm_b", gadj="gm_adjC", nioi="nioi", g_q4="g_q4", m_q4="m_q4", carry="anc_carry", scn="A")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "CY2025／FY2025", "CY2026／FY2026", "CY2027／FY2027", "CY2028／FY2028", "CY2029／FY2029", "CY2030", "CY2031", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="0.00", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 10, note); R[key] = r; r += 1
    ws.cell(r, 1, "A. 日曆年（C＝CY2025 … I＝CY2031）").font = BOLD; r += 1
    yy = lambda y: min(y, 2030) % 100
    row("bits", "HBM 出貨位元", "EB", {**{CYC[y]: f"=hs{s}_{y % 100}" for y in range(2025, 2030)}, "H": f"=G{r}*(1+{P['tail']})", "I": f"=H{r}*(1+{P['tail']})"}, note="markets/memory；CY2030–31 以 g_tail 外推")
    row("den", "DRAM 晶粒容量", "Gb", {CYC[y]: f"=den_{yy(y)}{P['den']}" for y in range(2025, 2032)})
    row("stk", "平均堆疊層數", "層", {CYC[y]: f"=stk_{yy(y)}{P['stk']}" for y in range(2025, 2032)})
    row("dies", "DRAM 晶粒數（每顆一次薄化＋切割）", "十億顆", {CYC[y]: f"={CYC[y]}{R['bits']}*8/{CYC[y]}{R['den']}" for y in range(2025, 2032)}, note="EB × 8×10⁹ Gb ÷ 晶粒容量")
    row("pas", "HBM 研削／切割工序數", "十億次", {CYC[y]: f"={CYC[y]}{R['dies']}*(1+{P['sstk']}/{CYC[y]}{R['stk']})" for y in range(2025, 2032)}, note="晶粒數 ×（1＋每堆疊額外工序 ÷ 層數）")
    row("hbp", "混合鍵合占比 × 開關", "%", {CYC[y]: f"=hbp_{yy(y)}{P['hbp']}*{P['hbon']}" for y in range(2025, 2032)}, fmt="0%")
    row("vh", "HBM 產能需求（精度加權工序）", "十億次", {CYC[y]: f"={CYC[y]}{R['pas']}*(1+{CYC[y]}{R['hbp']}*{P['hbe']})" for y in range(2025, 2032)}, note="混合鍵合工序 × (1＋淨影響)")
    nxh = lambda y: ("vhA_28" if (y == 2027 and s != "A") else f"{CYC[y+1]}{R['vh']}")   # v0.2.1：CY2027 的下一年（CY2028）產能讀情境 A 路徑，2027 三情境相同
    row("xh", "HBM 新機需求（下一年產能增量＋汰換；CY2027 的下一年產能讀情境 A）", "十億次", {CYC[y]: f"=MAX(0,{nxh(y)}-{CYC[y]}{R['vh']})+{P['dlt']}*{CYC[y]}{R['vh']}" for y in range(2025, 2031)}, note="設備在產能需要前一年購入；v0.2.1：CY2027 購機依 2027 年當時的共同前提（情境 A 的 CY2028 產能，具名 vhA_28），與 markets/equipment 2027 三情境相同一致；CY2028 起照讀本情境")
    um = {2028: f"*um28_{s}", 2029: f"*um29_{s}"}
    row("vl", "CoWoS 片數需求", "千片", {"C": f"={P['cw25']}", **{CYC[y]: f"=cw_{y % 100}{um.get(y, '')}" for y in range(2026, 2030)}, "H": f"=G{r}*(1+{P['tail']})", "I": f"=H{r}*(1+{P['tail']})"}, fmt="#,##0", note="台積電模型；B、C 的 2028–29 × 情境係數")
    nxl = lambda y: ("cw_28*um28_A" if y == 2027 else f"{CYC[y+1]}{R['vl']}")   # v0.2.1：CY2027 的下一年（CY2028）片數讀情境 A
    row("xl", "先進封裝新機需求（下一年片數增量＋汰換；CY2027 的下一年片數讀情境 A）", "千片", {CYC[y]: f"=MAX(0,{nxl(y)}-{CYC[y]}{R['vl']})+{P['dlt']}*{CYC[y]}{R['vl']}" for y in range(2025, 2031)}, fmt="#,##0", note="v0.2.1：CY2027 購機用 cw_28 × um28_A（情境 A），2027 三情境相同")
    ws.cell(r, 1, "B. 校準（只用 FY2025 AI 出貨錨點；FY(t)＝3/4×CY(t)＋1/4×CY(t+1)）").font = BOLD; r += 1
    fw = lambda key, y: f"(0.75*{CYC[y]}{R[key]}+0.25*{CYC[y+1]}{R[key]})"
    row("kh", "HBM：每單位新機需求的裝置營收 kH", "¥B／十億次", {"C": f"={P['ai']}*{P['w']}/{fw('xh', 2025)}"}, fmt="0.000", note="＝AI 出貨 × HBM 比重 ÷ FY2025 新機需求")
    row("kl", "先進封裝：每千片新機需求的裝置營收 kL", "¥B／千片", {"C": f"={P['ai']}*(1-{P['w']})/{fw('xl', 2025)}"}, fmt="0.0000")
    ws.cell(r, 1, "C. 會計年度（C＝FY2025 … G＝FY2029）").font = BOLD; r += 1
    FY = range(2025, 2030); FY1 = range(2026, 2030)
    row("shf", "市占係數（相對 2025）", "x", {FYC[y]: f"=1-{P['shl']}/sh*{(y - 2025) / 4}" for y in FY}, fmt="0.000")
    row("ehd", "HBM 研削／切割裝置（驅動推導）", "¥B", {FYC[y]: f"=$C${R['kh']}*{fw('xh', y)}*{FYC[y]}{R['shf']}" for y in FY}, fmt="#,##0.0")
    row("eld", "先進封裝（CoWoS 等）裝置（驅動推導）", "¥B", {FYC[y]: f"=$C${R['kl']}*{fw('xl', y)}*{FYC[y]}{R['shf']}" for y in FY}, fmt="#,##0.0")
    row("eaid", "AI 裝置合計（驅動推導）", "¥B", {FYC[y]: f"={FYC[y]}{R['ehd']}+{FYC[y]}{R['eld']}" for y in FY}, fmt="#,##0.0")
    gn = {2026: P["gn"], 2027: P["gn"], 2028: {"A": P["gn"], "B": P["gnB"], "C": P["gnC"]}[s], 2029: P["gn"]}
    row("end", "非 AI 裝置（驅動推導：功率、手機、車用、中國後段等）", "¥B", {"C": f"=eq_25-{P['ai']}", **{FYC[y]: f"={FYC[y-1]}{r}*(1+{gn[y]})*{FYC[y]}{R['shf']}/{FYC[y-1]}{R['shf']}" for y in FY1}}, fmt="#,##0.0")
    row("act", "AI 工序量指數（稼動；FY2025＝1）", "x", {FYC[y]: f"={P['w']}*{fw('vh', y)}/{fw('vh', 2025)}+(1-{P['w']})*{fw('vl', y)}/{fw('vl', 2025)}" for y in FY}, fmt="0.00", note="HBM 工序與 CoWoS 片數依 HBM 比重加權")
    row("taid", "AI 消耗品（驅動推導）", "¥B", {FYC[y]: f"=tl_25*{P['wtl']}*{FYC[y]}{R['act']}" for y in FY}, fmt="#,##0.0", note="隨工序量（稼動），不隨新機")
    gt = {2026: P["gtl"], 2027: P["gtl"], 2028: P["gtlC"] if s == "C" else P["gtl"], 2029: P["gtl"]}
    row("tnad", "非 AI 消耗品（驅動推導）", "¥B", {"C": f"=tl_25*(1-{P['wtl']})", **{FYC[y]: f"={FYC[y-1]}{r}*(1+{gt[y]})" for y in FY1}}, fmt="#,##0.0")
    row("revd", "營收（驅動推導；＝v0.1 的營收）", "¥B", {FYC[y]: f"=({FYC[y]}{R['eaid']}+{FYC[y]}{R['end']}+{FYC[y]}{R['taid']}+{FYC[y]}{R['tnad']})*(1+r_ot)" for y in FY}, fmt="#,##0.0", note="v0.1 的驅動，不含 FY2026 錨點")
    ws.cell(r, 1, "D. FY2026 錨點（公司資料：上半實績＋第三季指引＋第四季假設；三情境相同）").font = BOLD; r += 1
    row("a4r", "第四季營收（＝第三季指引 ×（1＋g_q4））", "¥B", {"D": f"=g3_rev*(1+{P['g_q4']})"}, fmt="#,##0.0")
    row("arev", "FY2026 錨點營收（上半＋第三季＋第四季）", "¥B", {"D": f"=q26h1_rev+g3_rev+D{R['a4r']}"}, fmt="#,##0.0")
    row("agp", "FY2026 錨點毛利（上半實績＋第三、四季營收 × 第三季毛利率）", "¥B", {"D": f"=q26h1_gp+(g3_rev+D{R['a4r']})*gm_q3"}, fmt="#,##0.0")
    row("aoi", "FY2026 錨點營業利益（上半＋第三季指引＋第四季營收 × m_q4）", "¥B", {"D": f"=q26h1_oi+g3_oi+D{R['a4r']}*{P['m_q4']}"}, fmt="#,##0.0")
    row("ani", "FY2026 錨點淨利（上半＋第三季指引＋第四季營業利益 × 淨利÷營業利益）", "¥B", {"D": f"=q26h1_ni+g3_ni+D{R['a4r']}*{P['m_q4']}*{P['nioi']}"}, fmt="#,##0.0")
    row("sc", "錨點 ÷ 驅動推導（FY2026）", "x", {"D": f"=D{R['arev']}/D{R['revd']}"}, fmt="0.000", note="FY2026 各組成同比例縮放成錨點")
    sc_ref = f"D{R['sc']}" if s == "A" else "sc_A"   # v0.1.3：B、C 的非 AI 基期一律用「同一組輸入下」情境 A 的 sc（具名範圍 sc_A 指向 PL 區塊 A）
    row("cf", "非 AI 兩項的基期調整（FY2027 起；用情境 A 的 sc）", "x", {"D": f"=1+{P['carry']}*({sc_ref}-1)"}, fmt="0.000", note="anc_carry＝1：非 AI 裝置與非 AI 消耗品以情境 A 的錨定水準為基期（B、C 較大的 sc 來自較低的 AI 驅動推導，不算非 AI）；Sens 案例皆為情境 A，用自己區塊的 sc")
    ws.cell(r, 1, "E. 損益（FY2025 實績；FY2026 錨點；FY2027 起驅動推導）").font = BOLD; r += 1
    mult = lambda y, ai: ("1" if y == 2025 else (f"$D${R['sc']}" if y == 2026 else ("1" if ai else f"$D${R['cf']}")))
    row("eh", "HBM 研削／切割裝置", "¥B", {FYC[y]: f"={FYC[y]}{R['ehd']}*{mult(y, True)}" for y in FY}, fmt="#,##0.0")
    row("el", "先進封裝（CoWoS 等）裝置", "¥B", {FYC[y]: f"={FYC[y]}{R['eld']}*{mult(y, True)}" for y in FY}, fmt="#,##0.0")
    row("eai", "AI 裝置合計", "¥B", {FYC[y]: f"={FYC[y]}{R['eh']}+{FYC[y]}{R['el']}" for y in FY}, fmt="#,##0.0")
    row("en", "非 AI 裝置", "¥B", {FYC[y]: f"={FYC[y]}{R['end']}*{mult(y, False)}" for y in FY}, fmt="#,##0.0")
    row("eq", "精密加工裝置營收", "¥B", {FYC[y]: f"={FYC[y]}{R['eai']}+{FYC[y]}{R['en']}" for y in FY}, fmt="#,##0.0")
    row("tai", "AI 消耗品", "¥B", {FYC[y]: f"={FYC[y]}{R['taid']}*{mult(y, True)}" for y in FY}, fmt="#,##0.0")
    row("tna", "非 AI 消耗品", "¥B", {FYC[y]: f"={FYC[y]}{R['tnad']}*{mult(y, False)}" for y in FY}, fmt="#,##0.0")
    row("tl", "精密加工ツール（消耗品）營收", "¥B", {FYC[y]: f"={FYC[y]}{R['tai']}+{FYC[y]}{R['tna']}" for y in FY}, fmt="#,##0.0")
    row("ot", "その他", "¥B", {FYC[y]: f"=r_ot*({FYC[y]}{R['eq']}+{FYC[y]}{R['tl']})" for y in FY}, fmt="#,##0.0")
    row("rev", "營收", "¥B", {FYC[y]: f"={FYC[y]}{R['eq']}+{FYC[y]}{R['tl']}+{FYC[y]}{R['ot']}" for y in FY}, fmt="#,##0.0", note="FY2026＝錨點營收")
    row("aish", "AI 占營收（AI 裝置＋AI 消耗品＋按比例的その他）", "%", {FYC[y]: f"=({FYC[y]}{R['eai']}+{FYC[y]}{R['tai']})*(1+r_ot)/{FYC[y]}{R['rev']}" for y in FY}, fmt="0.0%")
    row("aieq", "AI 裝置占營收（不含消耗品；その他按比例）", "%", {FYC[y]: f"={FYC[y]}{R['eai']}*(1+r_ot)/{FYC[y]}{R['rev']}" for y in FY}, fmt="0.0%", note="命題的「約兩成」用這一列")
    adj = {2028: P["gadj"] if s == "C" else "0", 2029: P["gadj"] if s == "C" else "0"}
    FY2 = range(2027, 2030)
    row("gm", "毛利率", "%", {"C": "=a25_gp/a25_rev", "D": f"=D{R['agp']}/D{R['rev']}", **{FYC[y]: f"={P['ga']}+{P['gb']}*LN({FYC[y]}{R['rev']})+{adj.get(y, '0')}" for y in FY2}}, fmt="0.0%", note="FY2026＝錨點；FY2027 起 a ＋ b × ln 營收（11 年迴歸）")
    row("gp", "毛利", "¥B", {FYC[y]: f"={FYC[y]}{R['rev']}*{FYC[y]}{R['gm']}" for y in FY}, fmt="#,##0.0")
    row("sga", "銷管費", "¥B", {"C": "=a25_sga", "D": f"=D{R['gp']}-D{R['aoi']}", **{FYC[y]: f"=sg_f+sg_v*{FYC[y]}{R['rev']}" for y in FY2}}, fmt="#,##0.0", note="FY2026＝錨點毛利 − 錨點營業利益；FY2027 起固定＋變動 × 營收")
    row("oi", "營業利益", "¥B", {FYC[y]: f"={FYC[y]}{R['gp']}-{FYC[y]}{R['sga']}" for y in FY}, fmt="#,##0.0")
    row("oim", "營業利益率（結果）", "%", {FYC[y]: f"={FYC[y]}{R['oi']}/{FYC[y]}{R['rev']}" for y in FY}, fmt="0.0%")
    row("ni", "淨利", "¥B", {"C": "=a25_ni", "D": f"=D{R['ani']}", **{FYC[y]: f"={FYC[y]}{R['oi']}*{P['nioi']}" for y in FY2}}, fmt="#,##0.0")
    row("eps", "EPS", "¥", {"C": "=a25_eps", **{FYC[y]: f"={FYC[y]}{R['ni']}*1000/shares" for y in FY1}}, fmt="#,##0.0")
    row("chk", "檢查：FY2025 營收、營業利益 − 實績；FY2026 營收、營業利益 − 錨點（應為 0）", "¥B", {"C": f"=ABS(C{R['rev']}-a25_rev)+ABS(C{R['oi']}-a25_oi)+ABS(D{R['rev']}-D{R['arev']})+ABS(D{R['oi']}-D{R['aoi']})"}, fmt="0.000")
    ws.cell(r, 1, "F. 對照：只用 v0.1 驅動（無錨點、淨利÷營業利益 0.733）").font = BOLD; r += 1
    row("gmd", "毛利率（驅動推導）", "%", {"C": "=a25_gp/a25_rev", **{FYC[y]: f"={P['ga']}+{P['gb']}*LN({FYC[y]}{R['revd']})+{adj.get(y, '0')}" for y in FY1}}, fmt="0.0%")
    row("oid", "營業利益（驅動推導）", "¥B", {"C": "=a25_oi", **{FYC[y]: f"={FYC[y]}{R['revd']}*{FYC[y]}{R['gmd']}-sg_f-sg_v*{FYC[y]}{R['revd']}" for y in FY1}}, fmt="#,##0.0")
    row("epsd", "EPS（驅動推導，v0.1 口徑）", "¥", {"C": "=a25_eps", **{FYC[y]: f"={FYC[y]}{R['oid']}*nioi_v01*1000/shares" for y in FY1}}, fmt="#,##0.0")
    R["end"] = r
    return R

# ---------------- PL：三情境 ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — Disco 損益（¥B；FY2025 實績；FY2026 錨定公司資料（上半實績＋第三季指引＋第四季假設）；FY2027–FY2029 驅動推導）。頂端為摘要，下方為三情境各一個完整推導區塊"
PB = {}; r0 = 70
for s in SC:
    PB[s] = engine(PL, r0, {"A": "情境 A：基準（加速器顆數續增）", "B": "情境 B：2028 顆數 ×0.8、2029 ×0.85；非 AI FY2028 持平", "C": "情境 C：2028–29 顆數 ×0.6；非 AI FY2028 −12%、毛利率下調"}[s], scn=s)
    r0 = PB[s]["end"] + 2
    if s == "A": defname("sc_A", f"PL!$D${PB['A']['sc']}"); defname("vhA_28", f"PL!$F${PB['A']['vh']}")   # vhA_28：情境 A 的 CY2028 HBM 產能需求（B、C 的 CY2027 購機讀這一格）   # B、C 區塊的 cf 讀這一格
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "¥B", "#,##0"), ("eh", "  HBM 裝置", "¥B", "#,##0"), ("el", "  先進封裝裝置", "¥B", "#,##0"), ("en", "  非 AI 裝置", "¥B", "#,##0"),
                                ("tl", "  消耗品", "¥B", "#,##0"), ("ot", "  その他", "¥B", "#,##0"), ("aish", "AI 占營收（含 AI 消耗品）", "%", "0.0%"), ("aieq", "AI 裝置占營收", "%", "0.0%"), ("gm", "毛利率", "%", "0.0%"),
                                ("oi", "營業利益", "¥B", "#,##0"), ("oim", "營業利益率", "%", "0.0%"), ("ni", "淨利", "¥B", "#,##0"), ("eps", "EPS", "¥", "#,##0"),
                                ("revd", "對照：營收（只用 v0.1 驅動）", "¥B", "#,##0"), ("oid", "對照：營業利益（只用 v0.1 驅動）", "¥B", "#,##0"), ("epsd", "對照：EPS（v0.1 驅動、淨利÷營業利益 0.733）", "¥", "#,##0")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in range(2025, 2030):
            PL[f"{FYC[y]}{sr}"] = f"={FYC[y]}{PB[s][key]}"; PL[f"{FYC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對：三情境 FY2025 − 實績、FY2026 − 錨點（營收與營業利益；應為 0）"); PL[f"C{sr}"] = "=" + "+".join(f"C{PB[s]['chk']}" for s in SC); PL[f"C{sr}"].number_format = "0.000"
PL.column_dimensions["A"].width = 50; PL.column_dimensions["B"].width = 14; PL.column_dimensions["J"].width = 50
for c in "CDEFGHI": PL.column_dimensions[c].width = 13

# ---------------- Val：P/E 估值 ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值。跨公司規則：合理倍數＝10 年（2016–2025 年底）P/E 的分位，依循環位置：見頂年 25 分位、中段中位數、谷底 75 分位"
V["A2"] = "P/E 定義（自算）：日曆年底收盤價 ÷ 「當期會計年度」公司公布的基本 EPS（例：2025-12 對 FY2025＝2026年3月期）。價格與 EPS 皆以 2023-04-01 一股分三股調整"
V.cell(4, 1, "1. 10 年 P/E（年底）").font = BOLD
hdr(V, 5, ["年底", "收盤價（未調整）", "分割係數", "收盤價（調整後）", "來源2：Yahoo Finance 調整後收盤", "兩者差距", "當期 FY EPS（調整後）", "P/E 當期", "標記", "說明"])
PX = {2016: (14160, 3, 4720.00, "2016-12-30"), 2017: (25070, 3, 8356.67, "2017-12-29"), 2018: (12850, 3, 4283.33, "2018-12-28"), 2019: (25900, 3, 8633.33, "2019-12-30"),
      2020: (34750, 3, 11583.33, "2020-12-30"), 2021: (35150, 3, 11716.67, "2021-12-30"), 2022: (37750, 3, 12583.33, "2022-12-30"), 2023: (34980, 1, 34980.00, "2023-12-29"),
      2024: (42730, 1, 42730.00, "2024-12-30"), 2025: (48170, 1, 48170.00, "2025-12-30")}
for i, y in enumerate(range(2016, 2026), 6):
    raw, f, yus, d = PX[y]
    V.cell(i, 1, f"{y}-12（{d}）")
    V.cell(i, 2, raw).font = BLUE; V.cell(i, 3, f).font = BLUE; V.cell(i, 4, f"=B{i}/C{i}"); V.cell(i, 5, yus).font = BLUE
    V.cell(i, 6, f"=D{i}/E{i}-1"); V.cell(i, 7, f"=Hist!K{i + 1}"); V.cell(i, 8, f"=D{i}/G{i}")
    for c, fm in (("B", "#,##0"), ("D", "#,##0.00"), ("E", "#,##0.00"), ("F", "0.00%"), ("G", "#,##0.00"), ("H", "0.0")): V[f"{c}{i}"].number_format = fm
    V.cell(i, 9, "Verified"); V.cell(i, 10, "來源1：Yahoo!ファイナンス 時系列（終值與調整後終值）；EPS：Hist 頁（公司決算短信）" if i == 6 else ("2023 年 P/E 偏高：FY2023 EPS 僅成長 1.5%，股價先反映 AI" if y == 2023 else ""))
V.cell(17, 1, "兩個價格來源可能共用同一資料商（不完全獨立）；10 年都有價格，無缺年。")
V.cell(19, 1, "2. 分位（10 年）").font = BOLD
hdr(V, 20, ["統計", "P/E 當期", "用途"])
for i, (nm, p, use) in enumerate((("25 分位", 0.25, "見頂年"), ("中位數", 0.5, "中段"), ("75 分位", 0.75, "谷底")), 21):
    V.cell(i, 1, nm); V.cell(i, 2, f"=PERCENTILE(H6:H15,{p})").number_format = "0.0"; V.cell(i, 3, use)
V.cell(24, 1, "年數"); V["B24"] = "=COUNT(H6:H15)"
defname("pe_peak", "Val!$B$21"); defname("pe_mid", "Val!$B$22"); defname("pe_trough", "Val!$B$23")
V.cell(25, 1, "敏感度：去掉 2023 年（極端值）時 25／50／75 分位");
for j, p in enumerate((0.25, 0.5, 0.75)):
    V.cell(25, 2 + j, f"=PERCENTILE((H6:H12,H14:H15),{p})").number_format = "0.0"
V.cell(27, 1, "3. FY2027 的循環位置（設備股共用規則；見頂年度讀設備市場模型，不在本模型另訂）").font = BOLD
V.cell(28, 1, CP.rule_text(PEAK))
hdr(V, 29, ["情境", "設備市場見頂年度", "FY2027 對照日曆年", "判定", "倍數", "v0.1 判定（營業利益路徑，對照）", "v0.1 倍數"])
for i, s in enumerate(SC, 30):
    V.cell(i, 1, s); V.cell(i, 2, f"=pk_{s}"); V.cell(i, 3, 2027).font = BLUE
    V.cell(i, 4, "=" + CP.excel_rule(f"C{i}", f"pk_{s}"))
    V.cell(i, 5, f'=IF(D{i}="見頂",pe_peak,IF(D{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 6, "中段"); V.cell(i, 7, "=pe_mid").number_format = "0.0"
V.cell(33, 1, "4. 隱含股價（FY2027 EPS × 依判定的倍數；其餘倍數僅供參考）").font = BOLD
hdr(V, 34, ["情境", "EPS 年度", "EPS", "倍數類型", "倍數", "隱含股價", "相對現價", "備註"])
vr = 35; VAL = {}
for i, s in enumerate(SC):
    for y, f, nm in ((2027, f"E{30 + i}", "規則值（依判定）"), (2027, "pe_peak", "參考：25 分位"), (2027, "pe_mid", "參考：中位數"), (2027, "pe_trough", "參考：75 分位"), (2028, "pe_mid", "參考：FY2028 × 中位數")):
        V.cell(vr, 1, s); V.cell(vr, 2, f"FY{y}"); V.cell(vr, 3, f"=PL!{FYC[y]}{SUMR[(s, 'eps')]}").number_format = "#,##0"
        V.cell(vr, 4, nm if f != f"E{30 + i}" else f"=\"規則值：\"&D{30 + i}"); V.cell(vr, 5, f"={f}").number_format = "0.0"
        V.cell(vr, 6, f"=C{vr}*E{vr}").number_format = "#,##0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
        if f == f"E{30 + i}":
            for c in range(1, 9): V.cell(vr, c).fill = KEY
            VAL[s] = vr
        vr += 1
V.cell(vr + 1, 1, "5. 目前倍數（僅供參考）").font = BOLD
hdr(V, vr + 2, ["項目", "數值", "備註"])
fr = vr + 3
for nm, f, fm, note in (("現價（2026-10-23 收盤）", "=px", "#,##0", "株探、kabutore"),
                        ("P/E：現價 ÷ FY2025 EPS 1,249.84", "=px/a25_eps", "0.0", "命題中的「約 41 倍」；這是上一個會計年度，不是嚴格的近 12 月"),
                        ("P/E：現價 ÷ 近 12 月 EPS（FY25 Q3＋Q4＋FY26 上半）", "=px/(a25_eps-515.70+q26h1_eps)", "0.0", "近 12 月 EPS ＝ 1,249.84 − 上半 515.70 ＋ 740.84 ＝ 1,474.98（Derived）"),
                        ("TTM P/E：stockanalysis（2026-10-21）", "=sa_ttm", "0.0", "其 TTM EPS 1,340.86，未含 FY26 Q2"),
                        ("遠期 P/E：stockanalysis（共識）", "=sa_fpe", "0.0", "參考，不用於估值"),
                        ("現價 ÷ 模型 FY2026 EPS（A）", f"=px/PL!{FYC[2026]}{SUMR[('A', 'eps')]}", "0.0", ""),
                        ("現價 ÷ 模型 FY2027 EPS（A／B／C）", f"=px/PL!{FYC[2027]}{SUMR[('A', 'eps')]}", "0.0", "B、C 見右")):
    V.cell(fr, 1, nm); V.cell(fr, 2, f).number_format = fm; V.cell(fr, 3, note); fr += 1
V.cell(fr - 1, 4, f"=px/PL!{FYC[2027]}{SUMR[('B', 'eps')]}").number_format = "0.0"; V.cell(fr - 1, 5, f"=px/PL!{FYC[2027]}{SUMR[('C', 'eps')]}").number_format = "0.0"
# 0. 目標價與上檔空間（v0.1.4；Andy 2026-10-25 要求；列標籤依工作單逐字）。放在 Val 既有內容之後，不位移既有列；HTML 估值章第一個區塊讀這裡
tp0 = fr + 1
V.cell(tp0, 1, "0. 目標價與上檔空間").font = BOLD
hdr(V, tp0 + 1, ["項目", "數值", "", "", "說明"])
V.cell(tp0 + 2, 1, "目標價（12 個月，情境加權）"); V.cell(tp0 + 2, 2, f"=tp_wA*B{tp0 + 4}+tp_wB*B{tp0 + 5}+tp_wC*B{tp0 + 6}").number_format = "#,##0"
V.cell(tp0 + 2, 5, "權重 × 各情境隱含價（第 4 節規則值）")
V.cell(tp0 + 3, 1, "上檔空間（目標價 ÷ 現價 − 1）"); V.cell(tp0 + 3, 2, f"=B{tp0 + 2}/px-1").number_format = "0.0%"; V.cell(tp0 + 3, 5, "現價＝In 頁 px（2026-10-23 收盤）")
for k_, (s_, lab) in enumerate((("A", "情境 A 隱含價（樂觀）"), ("B", "情境 B 隱含價（基準）"), ("C", "情境 C 隱含價（下檔）"))):
    rr_ = tp0 + 4 + k_
    V.cell(rr_, 1, lab); V.cell(rr_, 2, f"=F{VAL[s_]}").number_format = "#,##0"
    V.cell(rr_, 5, f'="FY2027 EPS × "&D{VAL[s_]}&" "&TEXT(E{VAL[s_]},"0.0")&"x（第 4 節 F{VAL[s_]}）"')
V.cell(tp0 + 7, 1, "情境權重 A／B／C")
for j_, s_ in enumerate("ABC"): V.cell(tp0 + 7, 2 + j_, f"=tp_w{s_}").number_format = "0%"
V.cell(tp0 + 7, 5, f'="合計 "&TEXT(B{tp0 + 7}+C{tp0 + 7}+D{tp0 + 7},"0%")&"；In 頁 tp_wA／tp_wB／tp_wC（Assumed，Andy 可改）"')
V.cell(tp0 + 8, 1, "目標價方法")
V.cell(tp0 + 8, 2, "P/E：各情境 FY2027 EPS × 十年（2016–2025 年底）P/E 依循環位置取分位（設備股共用規則：A、B 中段＝中位數，C 見頂＝25 分位）。"
                 "十年倍數一半以上來自 AI 前（2016–2022 年 16–34x），偏誤方向為低估；2023 年 EPS 僅成長 1.5%、股價先反映 AI 的 45.0x 則偏高（去掉後分位見第 2 節）。隱含價約為 FY2027 年底（2028-03）的價格、未折現")
V.column_dimensions["A"].width = 50
for c in "BCDEFGHI": V.column_dimensions[c].width = 15
V.column_dimensions["J"].width = 70

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A）：對 FY2027、FY2028 營收與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方，與 PL 同一引擎；校準會跟著重算）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收 ¥B", "FY2027 EPS ¥", "相對基準", "FY2028 營收 ¥B", "FY2028 EPS ¥", "相對基準", "FY2027 AI 占營收", "FY2026 EPS ¥", "相對基準"])
cases = [("汰換／升級比例 5%", "dlt 低", dict(dlt="dlt_lo")), ("汰換／升級比例 15%", "dlt 高", dict(dlt="dlt_hi")),
         ("AI 錨點 −10%（67.5）", "ai_25 低", dict(ai="ai_25_lo")), ("AI 錨點 +10%（82.5）", "ai_25 高", dict(ai="ai_25_hi")),
         ("HBM 比重 45%", "w_hbm 低", dict(w="w_hbm_lo")), ("HBM 比重 52%", "w_hbm 高", dict(w="w_hbm_hi")),
         ("混合鍵合開關 0", "不納入混合鍵合", dict(hbon="0")), ("混合鍵合淨影響 −20%", "電漿切割取代部分切割", dict(hbe="hb_eff_lo")),
         ("混合鍵合淨影響 +30%", "精度升級較大", dict(hbe="hb_eff_hi")), ("混合鍵合普及 高", "hbp 高欄", dict(hbp="_hi")),
         ("晶粒容量 低（同位元需更多晶粒）", "den 低欄", dict(den="_lo")), ("晶粒容量 高", "den 高欄", dict(den="_hi")),
         ("堆疊層數 低", "stk 低欄", dict(stk="_lo")), ("堆疊層數 高", "stk 高欄", dict(stk="_hi")),
         ("每堆疊額外工序 2", "s_stk 低", dict(sstk="s_stk_lo")), ("每堆疊額外工序 4", "s_stk 高", dict(sstk="s_stk_hi")),
         ("CY2030–31 需求年增 0%", "g_tail 低", dict(tail="g_tail_lo")), ("CY2030–31 需求年增 20%", "g_tail 高", dict(tail="g_tail_hi")),
         ("CoWoS 2025 600 千片", "cw_25 低", dict(cw25="cw_25_lo")),
         ("非 AI 裝置年增 0%", "g_non 低", dict(gn="g_non_lo")), ("非 AI 裝置年增 12%", "g_non 高", dict(gn="g_non_hi")),
         ("非 AI 消耗品年增 0%", "g_tl 低", dict(gtl="g_tl_lo")), ("非 AI 消耗品年增 10%", "g_tl 高", dict(gtl="g_tl_hi")),
         ("AI 消耗品比例 10%", "w_tlai 低", dict(wtl="w_tlai_lo")), ("AI 消耗品比例 40%", "w_tlai 高", dict(wtl="w_tlai_hi")),
         ("市占 FY2029 −10pt（線性）", "sh_loss 高", dict(shl="sh_loss_hi")),
         ("毛利率迴歸只用 FY2019–25", "gm_a2、gm_b2", dict(ga="gm_a2", gb="gm_b2")),
         ("淨利÷營業利益 低（0.687）", "稅前÷營業利益 0.97、稅率 29.2%", dict(nioi="nioi_lo")), ("淨利÷營業利益 高（0.738）", "稅前÷營業利益 1.00、稅率 26.2%", dict(nioi="nioi_hi")),
         ("第四季 ÷ 第三季 0%", "g_q4 低（持平）", dict(g_q4="g_q4_lo")), ("第四季 ÷ 第三季 +22%", "g_q4 高（FY2025 實績）", dict(g_q4="g_q4_hi")),
         ("第四季營益率 −2pt", "m_q4 低", dict(m_q4="m_q4_lo")), ("第四季營益率 +2pt", "m_q4 高", dict(m_q4="m_q4_hi")),
         ("錨點差額不延續到 FY2027", "anc_carry＝0（FY2027 起完全回到 v0.1 驅動）", dict(carry="0"))]
def put_case(i, name, setting, R, pre=""):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    SE[f"C{i}"] = f"={pre}{FYC[2027]}{R['rev']}"; SE[f"D{i}"] = f"={pre}{FYC[2027]}{R['eps']}"; SE[f"F{i}"] = f"={pre}{FYC[2028]}{R['rev']}"; SE[f"G{i}"] = f"={pre}{FYC[2028]}{R['eps']}"
    SE[f"E{i}"] = f"=D{i}/D$4-1"; SE[f"H{i}"] = f"=G{i}/G$4-1"; SE[f"I{i}"] = f"={pre}{FYC[2027]}{R['aish']}"
    SE[f"J{i}"] = f"={pre}{FYC[2026]}{R['eps']}"; SE[f"K{i}"] = f"=J{i}/J$4-1"; SE[f"J{i}"].number_format = "#,##0"; SE[f"K{i}"].number_format = "0.0%"
    for c in "CF": SE[f"{c}{i}"].number_format = "#,##0"
    for c in "DG": SE[f"{c}{i}"].number_format = "#,##0"
    for c in "EHI": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
SB = {}; base_row = 80
for n, (nm, setting, ov) in enumerate(cases):
    SB[nm] = engine(SE, base_row + n * 70, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    put_case(5 + n, nm, setting, SB[nm])
nn = 5 + len(cases)
put_case(nn, "情境 B", "顆數 2028 ×0.8、2029 ×0.85；非 AI FY2028 0%", PB["B"], "PL!")
put_case(nn + 1, "情境 C", "顆數 ×0.6；非 AI FY2028 −12%；毛利率 −1pt", PB["C"], "PL!")
SE.cell(nn + 2, 1, "對照：只用 v0.1 驅動（情境 A）"); SE.cell(nn + 2, 2, "無錨點、淨利÷營業利益 0.733")
SE[f"D{nn + 2}"] = f"=PL!E{PB['A']['epsd']}"; SE[f"G{nn + 2}"] = f"=PL!F{PB['A']['epsd']}"; SE[f"C{nn + 2}"] = f"=PL!E{PB['A']['revd']}"; SE[f"F{nn + 2}"] = f"=PL!F{PB['A']['revd']}"
SE[f"E{nn + 2}"] = f"=D{nn + 2}/D$4-1"; SE[f"H{nn + 2}"] = f"=G{nn + 2}/G$4-1"
for c in "CDFG": SE[f"{c}{nn + 2}"].number_format = "#,##0"
for c in "EH": SE[f"{c}{nn + 2}"].number_format = "0.0%"
SE.cell(nn + 4, 1, "讀法：校準在 FY2025 進行，所以改變 AI 錨點、HBM 比重、汰換比例等參數時，FY2025 仍等於實績，只改變之後的路徑形狀。汰換比例與校準係數只有乘積可識別（見 README）。")
SE.column_dimensions["A"].width = 40; SE.column_dimensions["B"].width = 40
for c in "CDEFGHIJK": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — FY2026 錨定公司資料（上半實績＋第三季指引＋第四季假設），A 節為恆等式；A2 節列只用 v0.1 驅動的結果；B 節共識只作參考，未用於校準或錨點。差距＝模型 ÷ 對照 − 1"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.0", pt=False):
    global rr_
    RC.cell(rr_, 1, label); RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        L_ = chr(64 + cv)
        RC.cell(rr_, cd, f"={L_}{rr_}-B{rr_}" if pt else f"={L_}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y=2026: (lambda s: f"=PL!{FYC[y]}{SUMR[(s, key)]}")
RC.cell(rr_, 1, "A. FY2026 錨點（公司資料）：模型＝錨點，差距為 0 是恆等式，不是獨立檢驗").font = BOLD; rr_ += 1
pa = lambda key: (lambda s: f"=PL!D{PB[s][key]}")
rec("FY2026 營收 vs 上半實績＋第三季指引＋第四季（Q3 × 1.10）", "=q26h1_rev+g3_rev+g3_rev*(1+g_q4)", pl("rev"), "決算短信＋說明資料＋g_q4（Assumed）", "257.1 ＋ 132.4 ＋ 145.7")
rec("FY2026 營業利益 vs 上半＋第三季＋第四季（Q4 營收 × 第三季營益率）", "=q26h1_oi+g3_oi+g3_rev*(1+g_q4)*m_q4", pl("oi"), "同上")
rec("第三季＋第四季平均營收 vs 第三季指引（132.4）", "=g3_rev", lambda s: f"=(PL!{FYC[2026]}{SUMR[(s, 'rev')]}-q26h1_rev)/2", "說明資料 3Q 132.4", "v0.1：−5.5%／−9.6%／−13.1%（A／B／C）；v0.1.1 由 g_q4 決定（+5%）")
rec("FY2026 營業利益率 vs 上半實績（百分點）", "=q26h1_oi/q26h1_rev", pl("oim"), "上半 45.4%", "下半營益率約 44%（第三季指引 43.9%）", fmt="0.0%", pt=True)
rec("FY2026 毛利率 vs 上半實績（百分點）", "=q26h1_gp/q26h1_rev", pl("gm"), "上半 71.5%", "第三、四季取 69.8%（公司：低不到 2 點）", fmt="0.0%", pt=True)
rec("FY2026 EPS vs 上半實績 × 2", "=q26h1_eps*2", pl("eps"), "決算短信（740.84）")
RC.cell(rr_, 1, "A2. 只用 v0.1 驅動時的 FY2026（對照；說明錨點改變了多少）").font = BOLD; rr_ += 1
rec("驅動推導營收 vs 錨點營收", "=PL!D" + str(PB["A"]["arev"]), pa("revd"), "PL 區塊 D、E 節", "負＝驅動推導低於公司資料")
rec("驅動推導營業利益 vs 錨點營業利益", "=PL!D" + str(PB["A"]["aoi"]), pa("oid"), "同上")
rec("驅動推導 EPS（淨利÷營業利益 0.733）vs 錨點 EPS", f"=PL!D{PB['A']['eps']}", pa("epsd"), "同上")
rec("驅動推導下半平均營收 vs 第三季指引", "=g3_rev", lambda s: f"=(PL!D{PB[s]['revd']}-q26h1_rev)/2", "說明資料 3Q 132.4", "＝v0.1 的對帳結果")
RC.cell(rr_, 1, "B. FY2026 vs 共識與券商（參考）").font = BOLD; rr_ += 1
rec("營收 vs 共識（stockanalysis 2026-10-07，20 位）", "=cons_rev26", pl("rev"), "stockanalysis（564.96；區間 541.0–598.7）", "共識早於 10-22 指引")
rec("營收 vs 樂天證券（2026-04-27）", 550.0, pl("rev"), "今中能夫")
rec("營業利益 vs 樂天證券", "=cons_oi26", pl("oi"), "今中能夫（240）")
rec("FY2027 營收 vs 樂天證券（640）", "=cons_rev27", pl("rev", 2027), "今中能夫")
RC.cell(rr_, 1, "C. 結構檢查").font = BOLD; rr_ += 1
rec("非 AI 裝置 FY2026 年增 vs SEMI 組裝封裝設備 CY2026 年增（百分點）", "=semi_pk/semi_pk_lo-1", lambda s: f"=PL!{FYC[2026]}{SUMR[(s, 'en')]}/PL!C{SUMR[(s, 'en')]}-1", "markets/equipment（SEMI 2026-07）", "口徑不同（SEMI 為全球封裝設備美元金額），只看方向", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "FY2025 AI 占營收（校準結果）"); RC.cell(rr_, 3, f"=PL!C{SUMR[('A', 'aish')]}").number_format = "0.0%"; RC.cell(rr_, 10, "AI 裝置 75 ＋ AI 消耗品，再按比例加その他；命題的「約兩成」"); rr_ += 1
RC.cell(rr_, 1, "FY2025 AI 裝置占營收（不含消耗品）"); RC.cell(rr_, 3, f"=PL!C{SUMR[('A', 'aieq')]}").number_format = "0.0%"; rr_ += 1
RC.cell(rr_, 1, "FY2027 AI 裝置占營收（A／B／C）")
for j, s in enumerate(SC): RC.cell(rr_, 3 + 2 * j, f"=PL!E{SUMR[(s, 'aieq')]}").number_format = "0.0%"
rr_ += 1
RC.cell(rr_, 1, "FY2027 AI 占營收（含 AI 消耗品；A／B／C）")
for j, s in enumerate(SC): RC.cell(rr_, 3 + 2 * j, f"=PL!E{SUMR[(s, 'aish')]}").number_format = "0.0%"
rr_ += 1
RC.cell(rr_, 1, "參考：公司稱 FY26 Q2 記憶體占營收略低於 40%、Q3 約 30%（含一般 DRAM／NAND，口徑不同，不可直接比較）"); rr_ += 1
RC.column_dimensions["A"].width = 62
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 70

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文以網頁擷取工具讀取；表格欄位錯置處以算術核對；版面與頁碼待人工複核）"; SO["A1"].font = BOLD
srcs = [("公司：2026年3月期 決算短信（英文，2026-04-22）", "https://www.disco.co.jp/eg/ir/library/doc/fr/fr20260422.pdf"),
 ("公司：2027年3月期 第2四半期 決算短信（英文，2026-10-22）", "https://www.disco.co.jp/eg/ir/library/doc/fr/fr20261022.pdf"),
 ("公司：2027年3月期 第1四半期 決算短信（2026-07-23）", "https://www.disco.co.jp/eg/ir/library/doc/fr/fr20260723.pdf"),
 ("公司：FY2026 2Q 決算說明資料（2026-10-22）", "https://www.disco.co.jp/eg/ir/library/doc/film/20261022.pdf"),
 ("公司：FY2025 4Q 決算說明資料（2026-04-22）", "https://www.disco.co.jp/eg/ir/library/doc/film/20260422.pdf"),
 ("公司：FY2024 4Q 說明資料（2025-04-17）／FY2023（2024-04-25）／FY2022（2023-04-20）", "https://www.disco.co.jp/eg/ir/library/doc/film/20250417.pdf"),
 ("公司：FY2026 1Q 說明筆記（2026-07-23）", "https://www.disco.co.jp/eg/ir/library/doc/film/20260723_note.pdf"),
 ("公司：歷年決算短信（fr20250417、fr20230420、fr20210422、fr20190508、fr20170510）", "https://www.disco.co.jp/eg/ir/library/fr.html"),
 ("公司：技術說明會 2025（2025-12）／2024（2024-12）", "https://www.disco.co.jp/jp/ir/movie/doc/J_Tech_Briefing_2025.pdf"),
 ("TSE 決算短信轉載（kabuyoho，2026-04-20 檔名；內容日期 2026-04-22）", "https://kabuyoho-sls-disclose.ifis.co.jp/295/140120260420506569.pdf"),
 ("松井證券：決算情報（年度、季度）", "https://finance.matsui.co.jp/stock/6146/settlement/index"),
 ("樂天證券 今中能夫：決算レポート（2026-04-27）", "https://media.rakuten-sec.net/articles/-/52260"),
 ("樂天證券 今中能夫（2025-11-04）", "https://media.rakuten-sec.net/articles/print/50557"),
 ("BigGo：FY2026 Q2 法說摘要（2026-10-22）", "https://finance.biggo.com/news/JP_6146.T_2026-10-22"),
 ("BigGo：FY2026 Q1 法說摘要（2026-07-23）", "https://finance.biggo.com/news/JP_6146.T_2026-07-23"),
 ("Investing.com：Q2 結果與指引（2026-10-22）", "https://www.investing.com/news/earnings/disco-reports-second-quarter-results-above-guidance-93CH-4965671"),
 ("EE Times Japan：FY2025 決算（2026-04-23）", "https://eetimes.itmedia.co.jp/ee/articles/2604/23/news056_2.html"),
 ("Yahoo!ファイナンス：株価時系列（年底收盤、2026-10-22）", "https://finance.yahoo.co.jp/quote/6146.T/history"),
 ("Yahoo Finance：Historical prices（調整後）", "https://finance.yahoo.com/quote/6146.T/history/"),
 ("stockanalysis：共識營收、P/E", "https://stockanalysis.com/quote/tyo/6146/forecast/"),
 ("Morningstar：市占 60–70%（2025）", "https://www.morningstar.com.au/investments/security/fra/D65/summary"),
 ("現代ビジネス（田宮寬之，2025-12-30）：市占 7–8 成", "https://gendai.media/articles/-/161100"),
 ("TrendForce：HBM4 規格（2025-04-28）", "https://www.trendforce.com/news/2025/04/28/news-following-jedecs-hbm4-standard-whats-next-for-sk-hynix-samsung-and-micron/"),
 ("TrendForce：HBM 高度上限放寬與混合鍵合（2026-03-06）", "https://www.trendforce.com/news/2026/03/06/news-industry-weigh-825-900-%CE%BCm-hbm-thickness-for-20-high-stacks-potentially-slowing-hybrid-bonding/"),
 ("3D InCites（SPTS）：電漿切割與 D2W 混合鍵合", "https://www.3dincites.com/2023/06/plasma-dicing-enables-challenging-applications-including-d2w-hybrid-bonding/"),
 ("Referently：混合鍵合說明（2026-10-06）", "https://referently.com/hybrid-bonding-explained-how-it-works-why-hbm-needs-it-and-who-supplies-the-tools/"),
 ("36Kr／半導體產業縱橫：CoWoS 2025 67 萬片（2025-12-04）", "https://eu.36kr.com/en/p/3580962946874242"),
 ("TrendForce：CoWoS 月產能（2024-10-21、2025-01-02）", "https://www.trendforce.com/news/2025/01/02/news-tsmc-set-to-expand-cowos-capacity-to-record-75000-wafers-in-2025-doubling-2024-output/"),
 ("tousaiblog：股票分割（2023-04-01 一股分三股）", "https://www.tousaiblog.org/disco-bunkatsu"),
 ("市場層：markets/memory（market.py @ 5fa3751，v0.4）、tsmc／skhynix／markets/equipment 成品（本 repo）", ""),
 ("無法開啟（建置環境被擋或逾時）：nikkei.com 株價推移、kabutan 年足、irbank、companiesmarketcap、traders.co.jp、stooq、macrotrends、gyokaidigest", ""),
 ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 80; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
PROP = "Disco 的營收由需要研削與切割的晶圓次數驅動，AI 只占約兩成；問題是 HBM 堆疊與先進封裝帶來的增量，能否抵過非 AI 後段的循環，撐住目前約 41 倍的近 12 月本益比？"
lines = [("Disco（6146.T）收支與估值模型 v0.2.2（AI 半導體 Project 步驟 5：設備與材料）　2026-10-25", True),
 ("v0.1.2（2026-10-23）：引用的上游成品檔名更新（SK 海力士 v0.4.1、美光 v0.3.1、三星 v0.2.1、台積電 v0.2.1 升版）；所有數值不變。", False),
 ("模型命題：" + PROP, False),
 ("註（v0.1 時以 2026-10-22 收盤計）：41 倍＝現價 ¥51,230 ÷ FY2025 EPS ¥1,249.84（上一個會計年度）；嚴格的近 12 月 EPS（含 FY2026 上半）為 1,474.98，對應 34.7 倍（Val 頁）。", False), ("", False),
 ("驅動因子 → 推導量（因果方向：左邊決定右邊）", True),
 ("工作表｜輸入 → 推導量 → 去向", True),
 ("Drv（PL 區塊 A）｜HBM 出貨位元（markets/memory，三情境）÷ 晶粒容量 → DRAM 晶粒數 ×（1＋每堆疊額外工序 ÷ 層數）→ HBM 研削／切割工序 → ×（1＋混合鍵合占比 × 淨影響）→ 產能需求 → 下一年增量＋汰換 → 新機需求 × kH（以 FY2025 AI 出貨校準）→ HBM 裝置營收", False),
 ("Drv（PL 區塊 A）｜CoWoS 片數（台積電模型；B、C × SCEN_UNITS）→ 下一年增量＋汰換 → 新機需求 × kL（同一錨點校準）→ 先進封裝裝置營收", False),
 ("PL｜非 AI 裝置 FY2025（裝置營收 − AI 出貨）×（1＋非 AI 年增，Analogy：公司 AI 前歷史）→ 非 AI 裝置營收；情境 B、C 只在 FY2028 不同", False),
 ("PL｜消耗品＝AI 部分 × AI 工序量指數（稼動）＋非 AI 部分 ×（1＋年增）——消耗品隨稼動、不隨新機；その他＝固定比例 ×（裝置＋消耗品）", False),
 ("Hist → PL｜毛利率＝a ＋ b × ln 營收、銷管費＝固定＋變動 × 營收（FY2015–FY2025 迴歸）→ 營業利益與營業利益率是結果，不寫死", False),
 ("PL（D 節）｜FY2026 錨點：上半實績＋第三季指引（由公司累計指引推得 132.4／58.1／41.5）＋第四季（＝第三季 ×（1＋g_q4），g_q4 基準 10%，Assumed 0–22%）；各組成同比例縮放成錨點，FY2027 起 AI 兩項回到驅動推導、非 AI 兩項以錨定水準為基期", False),
 ("Val｜10 年 P/E 分位 × FY2027 EPS；循環位置依設備股共用規則（tools/cycle_position.py：見頂年度讀設備市場模型 " + "、".join(f"{s_} {PEAK[s_]}" for s_ in SC) + " → FY2027：" + CP.fy2027_text(PEAK) + "）。Recon｜錨點恆等式、只用 v0.1 驅動的對照、共識。Sens｜一次動一個輸入", False), ("", False),
 ("v0.1.1 變更（2026-10-23）：(1) FY2026 錨定公司資料；(2) 淨利 ÷ 營業利益 0.733 → 0.722（稅前÷營業利益 0.99 ×（1 − 有效稅率 27.1%），含防衛特別法人稅 +0.9pt；區間 0.687–0.738）；(3) 循環位置改用設備股共用規則（v0.1 為各自依營業利益路徑判定，三情境皆中段）。前後對照見 disco/README.md", False), ("", False),
 ("會計年度：4 月到隔年 3 月（FY2025＝2025-04～2026-03，公司「2026年3月期」）。日曆年轉會計年度：FY(t)＝3/4 × CY(t)＋1/4 × CY(t+1)。金額為日圓，未建匯率。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司原文以網頁擷取讀取（數字已與松井證券交叉）；版面待人工複核。", False),
 ("可識別性：只有一個 AI 錨點（FY2025 ¥75B，券商估計）。kH、kL 各由一個點校準，因此「kH × 工序定義」「kL × 片數」只有乘積可識別；汰換比例、每堆疊工序數、晶粒容量與層數只改變路徑形狀，不改變 FY2025 水準，無法由資料分開，列在 Sens。AI 消耗品比例公司不揭露（Assumed）。市占只透過「變化」進入模型。", False),
 ("已知限制：見 disco/README.md「待辦與限制」。", False),
 ("v0.1.3 變更（2026-10-23）：錨點差額（sc）延續到 FY2027 以後的非 AI 兩項時，三情境一律用情境 A 的 sc（具名範圍 sc_A）。B、C 的 sc 較大（1.078、1.099 vs A 1.055）是因為其 AI 驅動推導較低，不是非 AI 較強；原寫法把這部分算成非 AI、墊高 B、C 的基期。FY2026 不變（三情境仍各自縮放到同一錨點）；情境 A 不變；Sens 各案例皆為情境 A，用自己區塊的 sc（改 g_q4、m_q4 時隨之變動）。前後對照見 disco/README.md", False),
 ("v0.1.4（2026-10-25）：Val 頁新增「0. 目標價與上檔空間」（放在 Val 既有內容之後，不位移既有列）：目標價＝In 頁權重 tp_wA／tp_wB／tp_wC（預設 A 0.3／B 0.5／C 0.2，Assumed，Andy 可改）× 三情境規則值隱含價；引用的上游成品檔名更新（SK 海力士 v0.4.2）；其餘數值不變。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 180

CFG_V2 = dict(cols=FYC, scn={"A": "順利", "B": "平平", "C": "不如預期"}, peers=[('東京威力科創（Tokyo Electron）', '8035.T', 27.09, '2026-10-21', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('Advantest', '6857.T', 37.69, '2026-10-21', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('Lasertec', '6920.T', 34.02, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('Kokusai Electric', '6525.T', 35.83, '2026-10-02', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('東京精密（Accretech）', '7729.T', 21.86, '2026-09-29', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）')], fwd={2027: "5/12", 2028: "7/12"}, fwd_note='約 2027-10 時的未來 12 月（2027-11～2028-10）≈ 5／12 × FY2027（2027-04～2028-03）＋7／12 × FY2028（會計年度 4–3 月，依月份調整）', a3='目標價（DCF 與 Comparable 加權）：見本頁第 {row} 列起「0. 目標價」（估值 v2 新增，放在既有內容之後以免位移既有列）；原「0. 目標價與上檔空間」改名為「舊法：規則值情境加權（參考）」',
    cur='¥',
    unit_rev='¥B',
    unit_ni='¥B',
    fmt='#,##0',
    fmt_rev='#,##0',
    band_fmt='#,##0',
    roe=('2024-03～2026-03 期', [0.2232, 0.2759, 0.2509], 'stockanalysis 比率頁（2026-10-25 讀取）：2024-03 期 22.32%、2025-03 期 27.59%、2026-03 期 25.09%', '單一來源'),
    nim=None,
    cons_tp=(81575, 'stockanalysis：21 位分析師平均 ¥81,575（中位數 78,850；區間 63,000–105,000），資料更新 2026-10-07', '第二來源：未找到；早於 2026-10-22 第二季財報'),
    cons_feps=(1759.08, '共識前瞻 EPS＝stockanalysis FY2026（2027-03 期）平均 1,759.08（2026-10-07）；比本模型 12 個月後前瞻 EPS 早約一年半，EPS 因素因此偏大', '第二來源：樂天證券今中能夫 FY2027 EPS 1,962.0（2026-04-27，單一分析師）'),
    peer_note='日本半導體設備同業：東京威力科創、Advantest、Lasertec、Kokusai Electric、東京精密（後兩家價格日期較舊）',
    dcf_note='；會計年度 4–3 月，FY2027＝2027-04～2028-03 仍以 t＝1 折現（略高估）',
    pe10='H6:H15',
    pe10_note='現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 1 節當期口徑）的分位；只作檢查，不進目標價',
    px_note='現價＝In 頁 px（2026-10-23 收盤）',
    method='DCF：股東現金流＝EPS ×（1 − g／ROE）（模型無折舊與資本支出項目），r 10%、g 4%、ROE 為近三年平均，FY2027–32 明確期（FY2030 起收斂法）＋永續。Comparable：東京威力科創、Advantest、Lasertec、Kokusai、東京精密前瞻 P/E 中位數 × 12 個月後前瞻 EPS。偏誤：十年平均淨利率未找到，以 FY2029 × 0.8 代替（Assumed）；同業含 Advantest 等高倍數股，偏高估。',
    readme=['v0.2（2026-10-25）估值 v2：Val 頁新增「0. 目標價」（放在既有內容之後，不位移既有列）＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，情境權重 A 0.3／B 0.5／C 0.2）＋0.5 × Comparable 目標價（同業前瞻 P/E 中位數 × 12 個月後前瞻 EPS）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解、歷史分位檢查。原「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」，不再叫目標價。', '預估期間延伸到 FY2032：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（PL 頁底部延伸區塊，灰底）。情境對應：A＝順利、B＝平平、C＝不如預期（原為顆數情境）。本版未做先驗，無 Andy 合理 P/E 對照。現價更新為 2026-10-23 收盤 ¥49,290（株探、kabutore；markets/invest_map/data/prices.csv）。十年平均淨利率未找到，收斂目標以 FY2029 淨利率 × 0.8 代替（Assumed）。'],
    nim_missing='十年淨利資料未找到（finboard.jp、irbank 在建置環境未能開啟；stockanalysis 只有五年）')

# ---------------- 估值 v2（2026-10-25，Andy 11:18／11:20／11:28／11:53）：PL 延伸到 FY2032（收斂法）＋ Val「0. 目標價」（DCF 與 Comparable 加權） ----------------
# 為不位移既有儲存格（FY2026–FY2029 既有 EPS 逐格不變），新內容放在各頁既有內容之後；HTML 估值章第一個區塊讀「0.3 結論」
def valuation_v2(CFG):
    cols = CFG["cols"]; MF = CFG["fmt"]; YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
    FADEFILL = PatternFill("solid", fgColor="E7E6E6")
    # --- In：目標價與收斂法參數（黃底；Assumed，Andy 可改） ---
    r = IN.max_row + 2
    IN.cell(r, 1, "J. 目標價（DCF 與 Comparable 加權）與 FY2030–32 收斂法參數（v2 新增；黃底＝Assumed，Andy 可改）").font = BOLD; r += 1
    def inrow(key, item, unit, base, tag, s1, s2, fmt=None, yellow=False):
        nonlocal r
        IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
        c = IN.cell(r, 5, base)
        if not (isinstance(base, str) and base.startswith("=")): c.font = BLUE
        if fmt: c.number_format = fmt
        if yellow: c.fill = KEY
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
        wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
        r += 1
    inrow("tp_r", "折現率（股東權益要求報酬；美、台、日、韓同一值）", "%", 0.10, "Assumed", "Andy 可改（工作單 2026-10-25 預設 10%）", CFG.get("r_note", "—"), "0.0%", True)
    inrow("tp_g", "長期成長率（永續成長；收斂法 FY2032 營收年增的基準）", "%", 0.04, "Assumed", "Andy 可改（工作單預設 4%）", "情境 A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt（PL 頁延伸區塊）", "0.0%", True)
    inrow("tp_roe", f"ROE（近三年平均，{CFG['roe'][0]}）", "%", "=AVERAGE(" + ",".join(str(x) for x in CFG["roe"][1]) + ")", "Derived", CFG["roe"][2], CFG["roe"][3], "0.0%")
    inrow("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", 0.5, "Assumed", "Andy 可改（2026-10-25 11:53：DCF／Comparable 各半）", "—", "0%", True)
    if CFG["nim"]:
        inrow("nim10", f"十年平均淨利率（{CFG['nim'][0]}，各年淨利 ÷ 營收的平均）", "%", "=AVERAGE(" + ",".join(f"{n}/{d}" for d, n in CFG["nim"][1]) + ")", "Derived", CFG["nim"][2], CFG["nim"][3], "0.0%")
    else:
        IN.cell(r, 1, "nim10"); IN.cell(r, 2, "十年平均淨利率"); IN.cell(r, 3, "%"); IN.cell(r, 5, "未找到"); IN.cell(r, 7, "Assumed")
        IN.cell(r, 8, CFG["nim_missing"]); IN.cell(r, 9, "依工作單：找不到十年平均時用最後一年（FY2029）淨利率 × 0.8；PL 延伸區塊逐情境計算"); r += 1
    inrow("cons_tp", "市場共識目標價（賣方平均）", CFG["cur"], CFG["cons_tp"][0], "Interested-party", CFG["cons_tp"][1], CFG["cons_tp"][2], MF)
    inrow("cons_feps", "共識前瞻 EPS（差距拆解用）", CFG["cur"], CFG["cons_feps"][0], "Interested-party", CFG["cons_feps"][1], CFG["cons_feps"][2], MF)

    # --- PL：延伸到 FY2032（收斂法） ---
    p0 = PL.max_row + 3
    PL.cell(p0, 1, "延伸到 FY2032（v2 新增）：FY2025–FY2029 引用本頁頂端摘要；灰底 FY2030–FY2032 為收斂法，非驅動因子推導").font = BOLD
    PL.cell(p0 + 1, 1, "收斂法：營收年增由 FY2029 年增線性收斂到 FY2032 的長期成長（A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt）；淨利率由 FY2029 線性收斂（A＝FY2029、B＝FY2029 與十年平均的中點、C＝十年平均" + ("" if CFG["nim"] else "；十年平均未找到，以 FY2029 × 0.8 代替（Assumed）") + "）；股數維持 FY2029（EPS 與淨利同比例）")
    ext_cols = {y: chr(ord("C") + y - 2025) for y in range(2025, 2033)}
    hdr(PL, p0 + 2, ["情境（延伸）", "單位"] + [str(PL[f"{cols[y]}3"].value) for y in range(2025, 2030)] + ["FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
    EXT = {}; pr = p0 + 3
    gadd = {"A": "+0.02", "B": "", "C": "-0.02"}
    for s in SC:
        PL.cell(pr, 1, f"情境 {s}（{CFG['scn'][s]}）").font = BOLD; pr += 1
        R = {k: pr + i for i, k in enumerate(("rev", "g", "nim", "ni", "eps"))}
        for k, lab in (("rev", "營收"), ("g", "  營收年增"), ("nim", "  淨利率（淨利 ÷ 營收）"), ("ni", "淨利（收斂法延伸）"), ("eps", "EPS（延伸）")):
            PL.cell(R[k], 1, lab); PL.cell(R[k], 2, {"rev": CFG["unit_rev"], "ni": CFG["unit_ni"], "eps": CFG["cur"]}.get(k, "%"))
        for y in range(2025, 2033):
            c = ext_cols[y]; pc = ext_cols.get(y - 1)
            if y <= 2029:
                PL[f"{c}{R['rev']}"] = f"={cols[y]}{SUMR[(s, 'rev')]}"
                PL[f"{c}{R['ni']}"] = f"={cols[y]}{SUMR[(s, 'ni')]}"
                PL[f"{c}{R['eps']}"] = f"={cols[y]}{SUMR[(s, 'eps')]}"
                PL[f"{c}{R['nim']}"] = f"={c}{R['ni']}/{c}{R['rev']}"
                if y > 2025: PL[f"{c}{R['g']}"] = f"={c}{R['rev']}/{pc}{R['rev']}-1"
            else:
                k = y - 2029
                mT = {"A": f"$G{R['nim']}", "B": f"($G{R['nim']}+{'nim10' if CFG['nim'] else '0.8*$G' + str(R['nim'])})/2", "C": "nim10" if CFG["nim"] else f"0.8*$G{R['nim']}"}[s]
                PL[f"{c}{R['g']}"] = f"=$G{R['g']}+((tp_g{gadd[s]})-$G{R['g']})*{k}/3"
                PL[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
                PL[f"{c}{R['nim']}"] = f"=$G{R['nim']}+({mT}-$G{R['nim']})*{k}/3"
                PL[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
                PL[f"{c}{R['eps']}"] = f"=$G{R['eps']}*{c}{R['ni']}/$G{R['ni']}"
                for kk in R: PL[f"{c}{R[kk]}"].fill = FADEFILL
            for kk, fm in (("rev", CFG["fmt_rev"]), ("g", "0.0%"), ("nim", "0.0%"), ("ni", CFG["fmt_rev"]), ("eps", MF)):
                if PL[f"{c}{R[kk]}"].value is not None: PL[f"{c}{R[kk]}"].number_format = fm
        PL.cell(R["rev"], 11, "FY2030 起為收斂法，非驅動因子推導")
        EXT[s] = R; pr += 6
    for c in "IJ": PL.column_dimensions[c].width = 14
    PL.column_dimensions["K"].width = 40

    # --- Val：舊法改名（不再叫目標價） ---
    REN = (("0. 目標價與上檔空間", "舊法：規則值情境加權（參考）"), ("目標價（12 個月，情境加權）", "規則值情境加權價（12 個月；參考）"),
           ("上檔空間（目標價 ÷ 現價 − 1）", "相對現價（規則值情境加權價 ÷ 現價 − 1）"), ("目標價方法", "規則值方法"))
    for row in V.iter_rows(min_col=1, max_col=1):
        v = row[0].value
        if isinstance(v, str):
            for a, b in REN:
                if v.startswith(a): row[0].value = b + v[len(a):]; break
    # --- Val：0. 目標價（列標籤依工作單逐字；B／C／D＝A 順利／B 平平／C 不如預期） ---
    t0 = V.max_row + 3
    T = {}
    def tset(rr, key, label, fs, fmt, note=""):
        V.cell(rr, 1, label); T[key] = rr
        for c, f in zip("BCD", fs):
            if f is None: continue
            V[f"{c}{rr}"] = f; V[f"{c}{rr}"].number_format = fmt
        if note: V[f"E{rr}"] = note
    V.cell(t0, 1, "0. 目標價").font = BOLD; T["top"] = t0
    V.cell(t0 + 1, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD; T["h01"] = t0 + 1
    hdr(V, t0 + 2, ["項目", f"A（{CFG['scn']['A']}）", f"B（{CFG['scn']['B']}）", f"C（{CFG['scn']['C']}）", "說明"]); T["hd1"] = t0 + 2
    V.cell(t0 + 3, 1, "情境 EPS（FY2027–FY2032）"); V[f"E{t0 + 3}"] = "PL 頁延伸區塊；FY2030 起為收斂法"; T["eps_h"] = t0 + 3
    EPSR = {}
    for i, y in enumerate(YRS6):
        tset(t0 + 4 + i, f"eps{y}", f"  FY{y}", [f"=PL!{ext_cols[y]}{EXT[s]['eps']}" for s in SC], MF); EPSR[y] = t0 + 4 + i
    rr = t0 + 10
    V.cell(rr, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）"); V[f"E{rr}"] = "模型無折舊與資本支出項目，故不用（淨利＋折舊－資本支出）；設備股資本支出低（輕資產），此式偏誤小"; T["cf_h"] = rr
    tset(rr + 1, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）")
    RV = rr + 1; CFR = {}
    for i, y in enumerate(YRS6):
        tset(rr + 2 + i, f"cf{y}", f"  FY{y}", [f"={c}{EPSR[y]}*(1-$B${RV})" for c in "BCD"], MF); CFR[y] = rr + 2 + i
    c1, c6 = CFR[2027], CFR[2032]
    tset(rr + 8, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=SUMPRODUCT({c}{c1}:{c}{c6},1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}{c6}*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6" for c in "BCD"], MF,
         "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底，t＝1 對 FY2027" + CFG["dcf_note"])
    tset(rr + 9, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}{rr + 8}*(1+tp_r)" for c in "BCD"], MF)
    tset(rr + 10, "dcftp", "DCF 目標價（情境加權）", [f"=tp_wA*B{rr + 9}+tp_wB*C{rr + 9}+tp_wC*D{rr + 9}", None, None], MF, "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁 tp_wA／tp_wB／tp_wC")
    q = rr + 12
    V.cell(q, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD; T["h02"] = q
    hdr(V, q + 1, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"]); T["phd"] = q + 1
    for i, (nm, tk, pe, dt, src) in enumerate(CFG["peers"], q + 2):
        V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
        V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
    pl_ = q + 1 + len(CFG["peers"]); T["p1"] = q + 2; T["pl"] = pl_
    q = pl_ + 1
    tset(q, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C{T['p1']}:C{pl_})", None, None], "0.0", CFG["peer_note"])
    fw = CFG["fwd"]
    def feps(c): return "(" + "+".join(f"{w}*{c}{EPSR[y]}" for y, w in fw.items()) + ")"
    tset(q + 1, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=tp_wA*" + feps("B") + "+tp_wB*" + feps("C") + "+tp_wC*" + feps("D"), None, None], MF, CFG["fwd_note"])
    tset(q + 2, "comptp", "Comparable 目標價", [f"=B{q}*B{q + 1}", None, None], MF, "同業中位數 × 上列 EPS")
    q += 4
    V.cell(q, 1, "0.3 結論").font = BOLD; T["h03"] = q
    hdr(V, q + 1, ["項目", "數值", "", "", "說明"]); T["hd3"] = q + 1
    q += 2
    tset(q, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B{T['comptp']}", None, None], MF, "tp_w_dcf × DCF 目標價＋（1 − tp_w_dcf）× Comparable 目標價；tp_w_dcf 見 In 頁")
    tset(q + 1, "comp2", "Comparable 目標價（市場主流同業法）", [f"=B{T['comptp']}", None, None], MF, "引用 0.2")
    tset(q + 2, "cons", "市場共識目標價", ["=cons_tp", None, None], MF, CFG["cons_tp"][1])
    tset(q + 3, "upw", "上檔空間：加權目標價", [f"=B{q}/px-1", None, None], "0.0%", CFG["px_note"])
    tset(q + 4, "upc", "上檔空間：Comparable 目標價", [f"=B{q + 1}/px-1", None, None], "0.0%")
    tset(q + 5, "geps", "與市場共識目標價的差距：EPS 因素", [f"=B{T['feps']}/cons_feps-1", None, None], "0.0%", "本模型 12 個月後前瞻 EPS ÷ 共識前瞻 EPS − 1；" + CFG["cons_feps"][1])
    tset(q + 6, "gmul", "與市場共識目標價的差距：倍數因素", [f"=(cons_tp/cons_feps)/B{T['pmed']}-1", None, None], "0.0%")
    V[f"E{q + 6}"] = f'="共識目標價 ÷ 共識前瞻 EPS 隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B{T["pmed"]},"0.0")&"x"'
    tset(q + 7, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D{T["dcf12"]},"{CFG["band_fmt"]}")&"～"&TEXT(B{T["dcf12"]},"{CFG["band_fmt"]}")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
    PE = CFG["pe10"]
    tset(q + 8, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）",
         [f"=IF(px/{c}{EPSR[2028]}<MIN({PE}),0,IF(px/{c}{EPSR[2028]}>MAX({PE}),1,PERCENTRANK({PE},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%", CFG["pe10_note"])
    V.cell(q + 9, 1, "目標價方法"); V[f"B{q + 9}"] = CFG["method"]; T["meth"] = q + 9
    for k in ("dcftp", "comptp", "tpw"): V[f"B{T[k]}"].fill = KEY
    if CFG.get("a3"): V["A3"] = CFG["a3"].format(row=t0)

    # --- Sens：目標價參數 ---
    s0 = SE.max_row + 3
    SE.cell(s0, 1, "目標價參數（v2 新增；只改 DCF 參數與權重，PL 收斂法的長期成長不隨 tp_g 變動）").font = BOLD
    hdr(SE, s0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
    def dcf_formula(c, r_="tp_r", g="tp_g"):
        cf = lambda y: f"Val!{c}{EPSR[y]}*(1-MIN({g}/tp_roe,0.8))"
        pv = "+".join(f"{cf(y)}/(1+{r_})^{t}" for t, y in enumerate(YRS6, 1))
        return f"({pv}+{cf(2032)}*(1+{g})/({r_}-{g})/(1+{r_})^6)"
    cases = [("基準", "r 10%、g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "r 11%", dict(r_="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", dict(r_="(tp_r-0.01)")),
             ("長期成長 +1pt", "g 5%（永續與再投資率）", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", dict(g="(tp_g-0.01)")),
             ("DCF 權重 0.4", "Comparable 0.6", dict(w="0.4")), ("DCF 權重 0.8", "Comparable 0.2", dict(w="0.8"))]
    for k, (nm, st, kw) in enumerate(cases):
        i = s0 + 2 + k; w = kw.pop("w", "tp_w_dcf")
        SE.cell(i, 1, nm); SE.cell(i, 2, st)
        rr_ = kw.get("r_", "tp_r")
        SE[f"C{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_formula(c, **kw)}" for s, c in zip(SC, "BCD")) + f")*(1+{rr_})"
        SE[f"D{i}"] = f"=Val!B{T['comptp']}"
        SE[f"E{i}"] = f"={w}*C{i}+(1-{w})*D{i}"
        SE[f"F{i}"] = f"=E{i}/E${s0 + 2}-1"
        for c in "CDE": SE[f"{c}{i}"].number_format = MF
        SE[f"F{i}"].number_format = "0.0%"
    T["sens0"] = s0; T["sens_last"] = s0 + 1 + len(cases)

    # --- README 頁 ---
    rr = RD.max_row + 1
    for t in CFG["readme"]:
        RD.cell(rr, 1, t); rr += 1
    return T, EXT

V2T, V2EXT = valuation_v2(CFG_V2)
RD.cell(RD.max_row + 1, 1, "v0.2.1（2026-10-25）情境一致性修正：CY2027 的「下一年設備增量」三情境改讀情境 A 路徑（HBM：具名 vhA_28＝情境 A 的 CY2028 產能需求；先進封裝：cw_28 × um28_A），與 markets/equipment 2027 三情境相同、markets/memory HBM 供給三情境相同的上游前提一致；CY2028 起照讀本情境。FY2027（2027-04～2028-03）含 CY2028 第一季，情境差只剩這一季（與 Advantest 相同處理）。混合鍵合占比三情境皆取基準（行業技術變數；韓美模型 CY2028 起 A 高、B 基準、C 低）。")

# ---------------- v0.2.2（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00、14:17）；新增列放在各頁現有最後一列之後，既有儲存格不變 ----------------
# 自由現金流（算力收支 2026-10-25 請求；地圖依列標籤讀取）：In 頁最後新增 da_rev、cx_rev；PL 頁最後新增一列（三情境 × FY2026–FY2032）
from openpyxl.utils import get_column_letter
_ir = IN.max_row + 1
for _k, _nm, _v, _src in (("da_rev", "折舊 ÷ 營收（自由現金流列用）", "=14.821/a25_rev", "FY2025（2026-03 期）實績：減価償却費 14,821 百萬日圓 ÷ 營收 436.889 十億日圓＝3.4%（stockanalysis 現金流量表；第二來源未找到）"),
                          ("cx_rev", "資本支出 ÷ 營收（自由現金流列用）", "=35.143/a25_rev", "FY2025 實績：設備投資 35,143 百萬日圓 ÷ 營收＝8.0%（stockanalysis 現金流量表；公司 FY2026 計畫約 330 億日圓，EE Times Japan 2026-04-23）")):
    IN.cell(_ir, 1, _k); IN.cell(_ir, 2, _nm); IN.cell(_ir, 3, "%"); IN.cell(_ir, 5, _v).number_format = "0.0%"
    IN.cell(_ir, 7, "Assumed"); IN.cell(_ir, 8, _src); IN.cell(_ir, 9, "錨＝最近完整年度 FY2025 實績 [Verified]；FY2026 起各年比率不變 [Assumed]")
    defname(_k, f"In!$E${_ir}"); _ir += 1
_fh = PL.max_row + 2
_FC = {}; _ci = 2
for _s in SC:
    for _y in range(2026, 2033):
        _FC[(_s, _y)] = get_column_letter(_ci); _ci += 1
PL.cell(_fh, 1, "項目")
for (_s, _y), _c in _FC.items(): PL[f"{_c}{_fh}"] = f"FY{_y} {_s}" + ("（收斂）" if _y >= 2030 else "")
hdr(PL, _fh, [PL.cell(_fh, _j).value for _j in range(1, _ci)])
PL.cell(_fh + 1, 1, "自由現金流（淨利＋折舊 − 資本支出；不含營運資金變動）")
for (_s, _y), _c in _FC.items():
    _e = chr(ord("C") + _y - 2025); _R2 = V2EXT[_s]
    PL[f"{_c}{_fh + 1}"] = f"=PL!{_e}{_R2['ni']}+PL!{_e}{_R2['rev']}*(da_rev-cx_rev)"; PL[f"{_c}{_fh + 1}"].number_format = "#,##0.0"
PL.cell(_fh + 2, 1, "註：上一列＝淨利 ＋ 營收 ×（折舊 ÷ 營收 − 資本支出 ÷ 營收），淨利與營收取本頁「延伸到 FY2032」區塊（FY2030 起為收斂法）；比率見 In 頁 da_rev、cx_rev（FY2025 實績）。¥B。")
_sr = SO.max_row + 1
for _a, _b in [["[Verified] 自由現金流列：Disco FY2025（2026-03 期）折舊攤銷 14,821、資本支出 35,143（百萬日圓；stockanalysis）", "https://stockanalysis.com/quote/tyo/6146/financials/cash-flow-statement/"], ["[Interested-party] 公司 FY2026 資本支出計畫約 330 億日圓（呉新工廠、羽田研發中心改建；EE Times Japan）", "https://eetimes.itmedia.co.jp/ee/spv/2604/23/news056_2.html"], ["[Assumed] 折舊 ÷ 營收、資本支出 ÷ 營收 FY2026 起各年維持 FY2025 實績比率", "本模型假設"]]:
    SO.cell(_sr, 1, _a); SO.cell(_sr, 2, _b); _sr += 1
_U = f"B{V2T['upw']}"
_G = V.max_row + 2
V.cell(_G, 1, "上檔空間級距")
V[f"B{_G}"] = f'=IF({_U}>0.5,"大於 50%",IF({_U}>0.25,"25%–50%",IF({_U}>=0,"0%–25%",IF({_U}>=-0.25,"0%～−25%",IF({_U}>=-0.5,"−25%～−50%","小於 −50%")))))'
V[f"B{_G}"].number_format = "@"; V[f"E{_G}"] = "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）"
V.cell(_G + 1, 1, "投資建議")
V[f"B{_G + 1}"] = f'=IF({_U}>0.5,"強力買進",IF({_U}>0.25,"買進",IF({_U}>=-0.25,"中立",IF({_U}>=-0.5,"放空","強力放空"))))'
V[f"B{_G + 1}"].number_format = "@"
V[f"E{_G + 1}"] = "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空"
OBS = ["目標價 ¥40,375（DCF 22,991、Comparable 57,759 各半），比現價 49,290（2026-10-23）低 18.1%，級距 0%～−25%；Comparable 單獨為 +17.2%。兩法分歧大：DCF 只有現價的 0.47 倍（A 情境 27,699 也只 0.56 倍），r 10%、g 4% 下 FY2027 EPS 1,723 只值約 16 倍，現價隱含 29.0 倍 12 個月後前瞻 EPS（1,698）；同業前瞻 P/E 中位數 34.0 倍、共識目標價 81,575 隱含 46.4 倍。",
       "本模型前瞻 EPS 1,698 比共識前瞻 EPS 1,759 低 3.5%（共識為 FY2026，早約一年半），差距主要在倍數（倍數因素 +36.3%）；情境 C 的 FY2028 EPS 1,326（A 1,851、B 1,745），來自顆數 ×0.6 與非 AI FY2028 −12%。",
       "FY2027 EPS 對非 AI 裝置年增（0%／12% → −4.8%／+4.9%）、淨利 ÷ 營業利益（0.687 → −4.8%）與市占流失（FY2029 −10pt → −4.2%）最敏感，HBM 相關參數影響都在 ±1% 內。下一個檢驗事件是 FY2026 第三季財報（約 2027-01 下旬），對照第三季營收指引 132.4 十億日圓。"]   # 重點觀察（與 disco/README.md 相同文字）
_R = RD.max_row + 1
RD.cell(_R, 1, "v0.2.2（2026-10-25）：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁最後新增「上檔空間級距」「投資建議」兩列（依「上檔空間：加權目標價」固定規則）；README 頁最後新增投資建議與重點觀察三點。另依算力收支 2026-10-25 請求，PL 頁最後新增「自由現金流（淨利＋折舊 − 資本支出；不含營運資金變動）」一列（三情境、所有模型年度；折舊、資本支出＝營收 × FY2025 實績比率 3.4%、8.0%），In 頁最後新增 「折舊 ÷ 營收」（da_rev）與「資本支出 ÷ 營收」（cx_rev）兩列；既有列不移動、既有數值不變。")
RD.cell(_R + 1, 1, f'="投資建議："&Val!B{_G + 1}&"（依上檔空間固定規則）"')
for _i, _t in enumerate(OBS, 1): RD.cell(_R + 1 + _i, 1, f"重點觀察 {_i}：{_t}")
if os.environ.get("G1_ROWS"): print("G1ROWS", "disco", "grade", _G, "rec", _G + 1, "readme_obs", _R + 2, _R + 4)

if os.environ.get("V2_ROWS"):
    print("V2", V2T, V2EXT)

order = ["README", "In", "Q", "Hist", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "disco")
wb.save(sys.argv[1])

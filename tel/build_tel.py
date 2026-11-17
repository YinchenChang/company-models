# 東京威力科創（Tokyo Electron／TEL，8035.T）個股模型 v0.1.0（AI 半導體：設備與材料，前段設備 WFE）
# 用法：python3 tel/build_tel.py raw.xlsx ；再以 LibreOffice headless 重算（-env:UserInstallation 私有設定檔）
# 會計年度 4 月到隔年 3 月：本模型 FY2025＝2025-04～2026-03（公司「2026年3月期」，公司稱 FY2026）；FY2026＝2026-04～2027-03（公司稱 FY2027）。
# 金額：十億日圓（¥B）；EPS：日圓，皆為 2026-10-01 一拆五之後（post-split）口徑。
# 結構參 accretech/build_accretech.py（日股 3 月年度、v1.3 版面）與 amat/lrcx（前段設備：終端市場需求指數 × 公司錨點）：
#   README、In、Q、Hist、HistIS、PL、Val、DCF、RevDCF、Recon、Sens、Coord、Sources。
# 設備市場讀 markets/equipment 成品 Mkt 頁 5A／5B／5C（先進邏輯、其他代工邏輯、其他前段、DRAM、NAND；同字母情境）；循環位置用 tools/cycle_position.py。
import sys, os, glob, json
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP
import fcff_dcf as FD   # 標準版 DCF（無槓桿自由現金流＋WACC）共用模組
sys.path.insert(0, os.path.join(ROOT, "siliconomics")); import reader as SIL   # 三層協調 C4 另一側讀數建置（SIL.COORD_ALT）在主引擎使用

VER, BUILT = "0.1.0", "2026-11-17"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC"); FADE = PatternFill("solid", fgColor="E7E6E6")
wb = openpyxl.Workbook()
SC = ("A", "B", "C")
SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- 讀入上游成品（唯讀；不修改） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
EQ_F = only("markets/equipment/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
ws_ = openpyxl.load_workbook(EQ_F, data_only=True)["Mkt"]
SEG_LAB = {"la": "先進邏輯（台積電設備＋其他代工中屬先進者）〔AI〕", "lo": "其他代工與邏輯", "of": "其他前段", "dr": "DRAM（含 HBM）〔AI〕", "nd": "NAND"}
SEG = {}; _cnt = {k: 0 for k in SEG_LAB}
SEMI = {}
for r in ws_.iter_rows():
    lab = r[0].value
    for k, l in SEG_LAB.items():
        if lab == l:
            s = SC[_cnt[k]]; _cnt[k] += 1
            SEG[(k, s)] = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
    for k, l in (("wfe", "SEMI 晶圓前段設備（WFE）"), ("fl", "SEMI 代工與邏輯設備"), ("dr", "SEMI DRAM 設備"), ("nd", "SEMI NAND 設備")):
        if lab == l:
            SEMI[k] = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert all(v == 3 for v in _cnt.values()) and len(SEMI) == 4, (_cnt, SEMI)
for k in SEG_LAB:   # 2025–2027 三情境相同（設備市場模型的設計）
    for j in range(3):
        assert abs(SEG[(k, "A")][0][j] - SEG[(k, "C")][0][j]) < 1e-9, k
PEAK, PEAK_SRC = CP.peak_years(ROOT)   # 設備股共用循環位置規則：各情境見頂年度（markets/equipment）
_AIS = {s: {SEG_LAB["la"]: SEG[("la", s)][0], SEG_LAB["dr"]: SEG[("dr", s)][0]} for s in SC}   # 另一側讀數建置：AI 分項（先進邏輯〔AI〕、DRAM〔AI〕）

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億日圓（¥B）；EPS：日圓（一拆五之後）；市場：十億美元。每個輸入附標記與來源；找不到者寫「未找到」，公司未揭露者寫「公司不揭露」"
IN["A2"] = ("標記：Verified（公司文件、交易所資料）／Interested-party（寫明是誰、立場）／Analogy（寫明對照與差異）／Assumed（一律給低／高區間）／Derived（寫明換算）。SemiAnalysis 未使用。"
            "公司原文（tel.com IR 的決算短信、說明資料、Q&A、逐字稿 PDF）以網頁擷取工具讀取（轉錄，數字已以算術互核；版面與頁碼待人工複核）")
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明", "來源誘因（Interested-party 列；混合來源列寫其他欄位）"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
TSQ2 = "公司 FY2027（2027年3月期）第2四半期 決算短信〔日本基準〕（連結）（2026-10-30；tel.com/ir/irta3a00000006g5-att/fy27q2tanshin-e.pdf）"
PRQ2 = "公司 FY2027 Q2 決算說明資料（2026-10-30；tel.com/ir/irta3a00000006g5-att/fy27q2presentations-e2.pdf）"
QAQ2 = "公司 FY2027 Q2 決算說明會 Q&A（2026-10-30；tel.com/ir/irta3a00000006g5-att/FY27Q2_EarningCall_QA_E.pdf）"
TRQ2 = "公司 FY2027 Q2 決算說明會逐字稿（2026-10-30；tel.com/ir/irta3a00000006g5-att/fy27q2transcript-e.pdf）"
SAI = "stockanalysis.com TYO:8035 損益表（百萬日圓；網站 FY 標籤＝截止年；2026-11-17 讀取；第三方轉錄）"
SAC = "stockanalysis.com TYO:8035 現金流量表（百萬日圓；2026-11-17 讀取；第三方轉錄）"
YHS = "Yahoo Finance 基本面快照（2026-11-17；docs/workorders/20261117_新增五家_Yahoo基本面快照.md；第三方）"
FY25N = "（本模型 FY2025＝公司 2026年3月期，公司稱 FY2026）"

sec("A. FY2025 實績（2025-04～2026-03；公司稱 FY2026；校準錨）")
inp("a25_rev", "營收 FY2025", "¥B", None, 2443.533, None, "Verified", TSQ2 + "：前期實績欄 2,443,533 百萬日圓", PRQ2 + "：2,443.5；" + SAI + "：2,443,533")
inp("a25_gp", "毛利 FY2025", "¥B", None, 1107.881, None, "Verified", PRQ2 + "：1,107.8（毛利率 45.3%）", SAI + "：1,107,881")
inp("a25_oi", "營業利益 FY2025", "¥B", None, 624.937, None, "Verified", TSQ2 + "：前期實績欄 624,936（百萬日圓，四捨五入）", SAI + "：624,937；" + PRQ2 + "：624.9")
inp("a25_sga", "銷管費 FY2025（含研發）", "¥B", None, "=a25_gp-a25_oi", None, "Derived", "＝毛利 − 營業利益（算術核對：" + PRQ2 + " 銷管費 482.9）", "研發 277.866（" + SAI + "）；研發以外 205.078")
inp("a25_rd", "研發費 FY2025", "¥B", None, 277.866, None, "Verified", SAI + "：277,866；" + PRQ2 + "：277.8", "")
inp("a25_pbt", "稅前利益 FY2025", "¥B", None, 748.180, None, "Verified", SAI + "：748,180；" + PRQ2 + "：748.1（上半 312.9＋下半 435.1）", "")
inp("a25_tax", "所得稅 FY2025", "¥B", None, 173.726, None, "Verified", SAI + "：173,726", "有效稅率 23.2%")
inp("a25_ni", "淨利（歸屬母公司）FY2025", "¥B", None, 574.454, None, "Verified", TSQ2 + "：前期實績欄 574,454", SAI + "：574,454")
inp("a25_eps", "基本 EPS FY2025（一拆五後）", "¥", None, 250.91, None, "Verified", TSQ2 + "：前期實績欄 250.91（post-split 口徑）", SAI + "：基本 250.91、稀釋 250.18")
inp("a25_da", "折舊攤銷 FY2025", "¥B", None, 81.345, None, "Verified", SAC + "：81,345", PRQ2 + "：D&A 季 21.1／23.5…")
inp("a25_cx", "資本支出 FY2025（公司口徑：有形＋無形固定資產）", "¥B", None, 216.0, None, "Verified", PRQ2 + "：資本支出趨勢 FY26（公司稱）216.0", SAC + "：現金流量表資本支出 208.984（口徑不同：不含無形或其他）")
inp("ne_25", "SPE 新設備銷售 FY2025（上半 850.1＋下半 925.3）", "¥B", None, "=850.1+925.3", None, "Derived", PRQ2 + "：SPE New Equipment Sales H1 FY26 850.1、H2 FY26 925.3（逐季 395.0／455.0／385.1／540.1）", "公司將營收分為新設備（SPE new equipment）、Field Solutions（現場解決方案＝零件、服務、二手與改裝設備）；其餘為殘差")
inp("fs_25", "Field Solutions 銷售 FY2025（四季合計）", "¥B", None, "=141.2+160.3+162.6+162.8", None, "Derived", PRQ2 + "：Field Solutions 逐季 141.2／160.3／162.6（另一處列 161.6）／162.8", "第三季數字在說明資料兩處不一致（162.6 與 161.6，差 1.0），取 162.6；只影響 FS 與殘差的拆分，不影響合計")
inp("fso_25", "Field Solutions 與其他（＝營收 − 新設備）FY2025", "¥B", None, "=a25_rev-ne_25", None, "Derived", "營收 − 新設備銷售（含 FS 與其他：FPD 設備等，公司不單列）", "FS 626.9、其他（殘差）約 41")
inp("mx25_h1_dr", "FY2025 上半（公司稱 H1 FY26）新設備銷售 DRAM 占比", "%", None, 0.61, None, "Interested-party", TRQ2 + "：slide 17 H1 FY2026 DRAM 61%／非揮發性記憶體 12%／非記憶體 27%", "圖表標籤，欄位對應為推斷（每欄合計 100%）；公司自揭 KPI")
inp("mx25_h1_nd", "FY2025 上半 非揮發性記憶體（NAND）占比", "%", None, 0.12, None, "Interested-party", TRQ2 + "：同上", "")
inp("mx25_h2_dr", "FY2025 下半（公司稱 H2 FY26）DRAM 占比", "%", None, 0.57, None, "Interested-party", TRQ2 + "：slide 17 H2 FY2026 DRAM 57%／非揮發性 8%／非記憶體 35%", "")
inp("mx25_h2_nd", "FY2025 下半 非揮發性記憶體占比", "%", None, 0.08, None, "Interested-party", TRQ2 + "：同上", "")

sec("B. FY2026 錨點：公司指引（2026-10-30 首次發布通期預想；Interested-party）與上半實績")
inp("g26_rev", "公司指引：FY2026 營收（2026-04～2027-03；公司稱 FY2027）", "¥B", None, 3700.0, None, "Interested-party", TSQ2 + "：3,700,000 百萬日圓（+51.4%）；" + PRQ2 + "：H2 隱含 2,032.5", "通期預想首次發布（7 月只發布上半預想 1,620／OP 458）。上半實績 1,667.455 → 下半隱含 2,032.5（Derived）")
inp("g26_ne", "公司指引：FY2026 新設備銷售（上半實績 1,265.7＋下半預想 1,560.0）", "¥B", None, "=ne_h1+ne_h2", None, "Derived", PRQ2 + "：新設備 H1 實績 1,265.7、H2 預想 1,560.0；通期 +59%", "")
inp("ne_h1", "FY2026 上半 新設備銷售（實績）", "¥B", None, 1265.7, None, "Verified", PRQ2 + "：H1 FY27 1,265.7（Q1 531.3、Q2 734.4）", TRQ2)
inp("ne_h2", "FY2026 下半 新設備銷售（公司預想）", "¥B", None, 1560.0, None, "Interested-party", PRQ2 + "：H2 FY27 estimate 1,560.0", "公司預想；公司稱客戶要求提前交貨，Q4 毛利率若提前認列裝機收入可到 50%（" + QAQ2 + "）")
inp("g26_gp", "公司指引：FY2026 毛利", "¥B", None, 1772.0, None, "Interested-party", PRQ2 + "：1,772.0（毛利率 47.9%；H2 48.3%）", "上半實績 790.102 → 下半隱含 981.9")
inp("g26_sga", "公司指引：FY2026 銷管費（含研發 335.0）", "¥B", None, 603.0, None, "Interested-party", PRQ2 + "：603.0（研發 335.0、其他 268.0）", "")
inp("g26_rd", "公司指引：FY2026 研發費", "¥B", None, 335.0, None, "Interested-party", PRQ2 + "：335.0", "")
inp("g26_oi", "公司指引：FY2026 營業利益", "¥B", None, 1169.0, None, "Interested-party", TSQ2 + "：1,169,000（+87.1%）；營業利益率 31.6%", "上半實績 506.776 → 下半隱含 662.2（Derived）")
inp("g26_pbt", "公司指引：FY2026 稅前利益", "¥B", None, 1172.0, None, "Interested-party", PRQ2 + "：1,172.0", "經常利益 1,179.0（" + TSQ2 + "）")
inp("g26_ni", "公司指引：FY2026 淨利（歸屬母公司）", "¥B", None, 865.0, None, "Interested-party", TSQ2 + "：865,000（+50.6%）", "上半實績 376.929 → 下半 488.0")
inp("g26_eps", "公司指引：FY2026 基本 EPS（一拆五後）", "¥", None, 381.69, None, "Interested-party", TSQ2 + "：381.69", "隱含平均股數 g26_ni ÷ g26_eps ＝ 2,266.2 百萬股")
inp("g26_cx", "公司計畫：FY2026 資本支出", "¥B", None, 190.0, None, "Interested-party", PRQ2 + "：FY27 estimate 190.0（FY26 216.0）", TRQ2 + "：FY2027 plan 190 billion yen")
inp("g26_da", "公司預想：FY2026 折舊攤銷", "¥B", None, 100.0, None, "Interested-party", PRQ2 + "：D&A FY27 estimate 100.0（H1 實績 46.8）", "")
inp("h1_rev", "FY2026 上半營收（實績）", "¥B", None, 1667.455, None, "Verified", TSQ2 + "：1,667,455（+41.3%）", PRQ2 + "：Q1 732.3＋Q2 935.0")
inp("h1_oi", "FY2026 上半營業利益（實績）", "¥B", None, 506.776, None, "Verified", TSQ2 + "：506,776（+67.2%）", PRQ2)
inp("h1_ni", "FY2026 上半淨利（實績）", "¥B", None, 376.929, None, "Verified", TSQ2 + "：376,929（+56.0%）", PRQ2)
inp("h1_eps", "FY2026 上半基本 EPS（實績）", "¥", None, 166.13, None, "Verified", TSQ2 + "：基本 166.13、稀釋 165.59", "")
inp("shares", "平均股數（FY2026 指引隱含；一拆五後）", "百萬股", None, "=g26_ni/g26_eps*1000", None, "Derived", "＝公司指引淨利 ÷ 指引 EPS（2,266.2）；上半期中平均 2,268.911（" + TSQ2 + "）", "FY2027 起每年依 bb 減少（買回）")
inp("g_anc", "FY2026 實績相對公司指引（營收各列同比例）", "%", -0.03, 0.0, 0.05, "Assumed", "基準＝公司指引（上半已為實績，偏離只會來自下半）。上半實績高於 7 月預想 營收 +2.9%、營業利益 +10.6%（" + PRQ2 + "）", "高：公司稱客戶要求提前交貨、供應鏈能見度 FY2028 前改善；低：客戶 utility 就緒延遲使裝機認列遞延（" + QAQ2 + "）。營業利益隨營收以邊際利潤率變動")

sec("C. 設備市場（markets/equipment 成品 Mkt 頁 5A／5B／5C，同字母情境；唯讀；日曆年 US$B）")
NM = {"la": "先進邏輯設備", "lo": "其他代工與邏輯設備", "of": "其他前段設備（類比、功率、光罩、廠務等）", "dr": "DRAM（含 HBM）設備", "nd": "NAND 設備"}
for s in SC:
    for k in SEG_LAB:
        for j, y in enumerate(range(2025, 2030)):
            inp(f"{k}{s}_{y % 100}", f"{NM[k]} CY{y}（{s}）", "US$B", None, SEG[(k, s)][0][j], None, "Derived", SEG[(k, s)][1] if j == 0 else "同上",
                "markets/equipment Mkt 第 5 節（公司資本支出 × 設備占比，2025 以 SEMI 實績校準）；2028–29 隨情境係數" if (j == 0 and k == "la" and s == "A") else "")
inp("g_tail", "CY2030 設備市場相對 CY2029（三情境相同；市場層只到 2029）", "%", -0.10, 0.0, 0.10, "Assumed", "無：FY2029＝3/4 × CY2029 ＋ 1/4 × CY2030，需要 CY2030；基準持平", "Sens 列 ±10%")
for k, l in (("wfe", "SEMI 晶圓前段設備（WFE）"), ("fl", "SEMI 代工與邏輯設備"), ("dr", "SEMI DRAM 設備"), ("nd", "SEMI NAND 設備")):
    inp(f"semi_{k}26", f"{l} CY2026（對帳用）", "US$B", None, SEMI[k][0][1], None, "Interested-party", SEMI[k][1], "SEMI 2026-07 年中預測（設備商協會，傾向樂觀）；只用於 Recon")
inp("co_wfe26", "公司說法：WFE 市場 CY2026", "US$B", None, 160.0, None, "Interested-party", QAQ2 + "：約 160B（邏輯與代工約 60%、記憶體約 40%；DRAM 高 40 億級、NAND 中 10 億級、邏輯代工高 80 億級、其他約 10）", "公司 2026-10-30 較 7 月上修；只用於 Recon")
inp("co_wfe27", "公司說法：WFE 市場 CY2027", "US$B", None, 210.0, None, "Interested-party", QAQ2 + "：超過 210B（DRAM 低 70 億級、NAND 高 10 億級、邏輯代工約 100、其他低 10 億級）；CY2028 無數字，客戶討論無放緩跡象", "公司稱 CY2027 為客戶資本支出計畫加總，另有上行空間；只用於 Recon")
inp("co_dr26", "公司說法：DRAM 設備 CY2026", "US$B", None, 48.0, None, "Interested-party", QAQ2 + "：high-$40B range（取 48）", "只用於 Recon")
inp("co_nd26", "公司說法：NAND 設備 CY2026", "US$B", None, 15.0, None, "Interested-party", QAQ2 + "：mid-$10B range（取 15）", "只用於 Recon")

sec("D. 新設備：FY2026 終端市場構成（公司揭露）與各終端市場需求指數")
inp("mx26_h1_dr", "FY2026 上半（公司稱 H1 FY27）新設備 DRAM 占比（實績）", "%", None, 0.57, None, "Interested-party", TRQ2 + "：slide 17 H1 FY2027 DRAM 57%／非揮發性 10%／非記憶體 33%（Q2 單季 34%／10%／56%）", "圖表標籤，欄位對應為推斷；公司自揭")
inp("mx26_h1_nd", "FY2026 上半 非揮發性記憶體占比（實績）", "%", None, 0.10, None, "Interested-party", TRQ2 + "：同上", "")
inp("mx26_h2_dr", "FY2026 下半（公司稱 H2 FY27）DRAM 占比（預想）", "%", None, 0.51, None, "Interested-party", TRQ2 + "：slide 17 H2 FY2027 forecast DRAM 51%／非揮發性 14%／非記憶體 35%", "公司預想")
inp("mx26_h2_nd", "FY2026 下半 非揮發性記憶體占比（預想）", "%", None, 0.14, None, "Interested-party", TRQ2 + "：同上", "")
inp("ne26_dr", "FY2026 新設備：DRAM", "¥B", None, "=mx26_h1_dr*ne_h1+mx26_h2_dr*ne_h2", None, "Derived", "上半占比 × 上半新設備 ＋ 下半占比 × 下半預想", "")
inp("ne26_nd", "FY2026 新設備：非揮發性記憶體（NAND）", "¥B", None, "=mx26_h1_nd*ne_h1+mx26_h2_nd*ne_h2", None, "Derived", "同上", "")
inp("ne26_nm", "FY2026 新設備：非記憶體（邏輯、代工、其他）", "¥B", None, "=g26_ne-ne26_dr-ne26_nd", None, "Derived", "新設備 − DRAM − NAND", "非記憶體對應設備市場的先進邏輯＋其他代工邏輯＋其他前段")
inp("ne25_dr", "FY2025 新設備：DRAM", "¥B", None, "=mx25_h1_dr*850.1+mx25_h2_dr*925.3", None, "Derived", "FY2025 上半、下半占比 × 新設備（850.1／925.3）", "")
inp("ne25_nd", "FY2025 新設備：非揮發性記憶體", "¥B", None, "=mx25_h1_nd*850.1+mx25_h2_nd*925.3", None, "Derived", "同上", "")
inp("ne25_nm", "FY2025 新設備：非記憶體", "¥B", None, "=ne_25-ne25_dr-ne25_nd", None, "Derived", "新設備 − DRAM − NAND", "")
inp("fyw", "（說明）會計年度對日曆年：FY(t)＝3/4 × CY(t) ＋ 1/4 × CY(t+1)", "權重", None, 0.75, None, "Derived", "FY(t)＝t 年 4 月～t+1 年 3 月，9 個月落在 CY(t)、3 個月落在 CY(t+1)", "同 accretech、advantest；不改寫成輸入（公式內固定）")

sec("E. 公司特有情境變數（方向一致：A 最有利、C 最不利；Assumed）")
for s, b_, lo, hi, note in (("A", 0.02, 0.0, 0.04, "相對設備市場的份額與組合提升：蝕刻（宮城新廠 2027 夏完工）、批次 ALD 超過 200 台出貨（FY2027、FY2028 各）、探針台出貨約 +60%（" + TRQ2 + "）、價格最適化 FY2028 全額貢獻（" + QAQ2 + "）"),
                            ("B", 0.0, -0.02, 0.02, "與市場同步"),
                            ("C", -0.02, -0.04, 0.0, "中國設備自給化、競爭（AMAT、Lam）；塗佈顯影市占約 91% 難再提高，蝕刻與沉積為競爭較強領域")):
    inp(f"exc_{s}", f"情境 {s}：新設備相對設備市場的年超額成長（份額＋組合；FY2027 起）", "%/年", lo, b_, hi, "Assumed", note, "份額水準：公司只稱塗佈顯影 >90%（" + QAQ2 + "），其他產品市占未找到；只以變化進入模型")
inp("chn", "中國占新設備銷售的比重（FY2026 上半）", "%", 0.30, 0.34, 0.40, "Interested-party", PRQ2 + "：Q1 222.9、Q2 339.9（占全社營收 Q2 36.4%、上半約 33.7%）", "以全社營收占比代替新設備占比（區域數字不分新設備與 FS）；低／高 Assumed")
for s, b_, lo, hi in (("A", 0.0, None, None), ("B", -0.05, -0.10, 0.0), ("C", -0.20, -0.30, -0.10)):
    inp(f"chh_{s}", f"情境 {s}：中國新設備營收水準變化（FY2027 一半、FY2028 起全額）", "%", lo, b_, hi, "Assumed", "公司未揭露中國管制假設；A：中國需求維持高水準；B：小幅回落；C：自給化或管制擴大" if s == "A" else "同上", "新設備 ×（1 ＋ 中國占比 × 本變數）")
for s, b_, lo, hi in (("A", 0.12, 0.08, 0.16), ("B", 0.08, 0.04, 0.12), ("C", 0.03, 0.0, 0.06)):
    inp(f"gfso_{s}", f"情境 {s}：Field Solutions 與其他年增（FY2027 起）", "%", lo, b_, hi, "Analogy", "類比 AMAT AGS 長期年增 mid-teens、Lam CSBG；TEL 上半 FS＋其他年增 21.9%（Derived：" + PRQ2 + "）、FS 季營收連續六季成長（" + TRQ2 + "）", "與現在的差異：TEL FY2026 錨點（公司指引）隱含 FS＋其他年增 30.9%，含提前認列裝機收入的成分，FY2027 起回到裝機量累積驅動的成長" if s == "A" else "同上")
inp("madjC", "情境 C：FY2028 起額外的營業利益率調整（價格最適化回吐、稼動下降）", "pt", -0.03, -0.015, 0.0, "Assumed", "無：保守情境設備市場 2028 年下滑兩成，兩點推導的邊際毛利率可能偏樂觀", "")
inp("dm", "邊際營業利益率調整（價格最適化、組合；FY2027 起各情境）", "pt", -0.03, 0.0, 0.03, "Assumed", "公司稱價格最適化是持續性措施、FY2028 全額貢獻；Q4 毛利率可達 50%（" + QAQ2 + "）", "Sens：±3 個百分點（影響營業成本變動比例）")

sec("F. 利潤率（兩點推導：FY2025 實績與 FY2026 錨點；不寫死）與淨利、股數")
inp("cogs25", "營業成本 FY2025（營收 − 毛利）", "¥B", None, "=a25_rev-a25_gp", None, "Derived", "In 頁", "")
inp("cogs26", "營業成本 FY2026 錨點（指引營收 − 指引毛利）", "¥B", None, "=g26_rev-g26_gp", None, "Derived", "公司指引", "")
inp("cv", "營業成本變動部分占營收（兩點斜率）", "%", None, "=(cogs26-cogs25)/(g26_rev-a25_rev)", None, "Derived", "（FY2026 − FY2025 營業成本）÷ 營收差", "隱含邊際毛利率約 53%（公司：一半來自量、一半來自價格最適化）")
inp("mm", "邊際營業利益率（FY2025→FY2026 錨點；兩點推導）", "%", None, "=(g26_oi-a25_oi)/(g26_rev-a25_rev)", None, "Derived", "（公司 FY2026 指引營業利益 − FY2025 實績）÷（指引營收 − 實績營收）＝ 43.3%", "FY2027 起營業利益＝FY2026 錨點 ＋（mm ＋ dm）×（營收 − FY2026 營收）；量增帶來研發與銷管費增加已含在兩點斜率內")
inp("cF", "營業成本固定部分", "¥B", None, "=cogs25-cv*a25_rev", None, "Derived", "FY2025 營業成本 − 變動部分", "")
inp("opex25", "營業費用 FY2025（銷管費含研發）", "¥B", None, "=a25_sga", None, "Derived", "In 頁", "")
inp("opex26", "營業費用 FY2026 錨點", "¥B", None, "=g26_sga", None, "Derived", "公司指引 603.0", "")
inp("ov", "營業費用變動部分占營收（兩點斜率）", "%", None, "=(opex26-opex25)/(g26_rev-a25_rev)", None, "Derived", "（FY2026 − FY2025 營業費用）÷ 營收差", "研發 FY2025 277.9 → FY2026 335.0 為主")
inp("oF", "營業費用固定部分", "¥B", None, "=opex25-ov*a25_rev", None, "Derived", "FY2025 營業費用 − 變動部分", "")
inp("nioi", "淨利 ÷ 營業利益（FY2027 起）", "x", 0.70, "=g26_ni/g26_oi", 0.78, "Derived", "基準＝公司 FY2026 指引隱含 0.740（865.0 ÷ 1,169.0；稅前 1,172.0、有效稅率約 26.2%）；高＝FY2024 實績 0.770（544.1 ÷ 697.3 為 0.780）", "低 0.70（Assumed：稅率上升）。FY2025 0.919 含 123 的營業外收益與特別利益（稅前 748.2 對營業利益 624.9），不採用")
inp("bb", "FY2027 起平均股數年減幅（淨買回）", "%", 0.0, 0.008, 0.015, "Assumed", "FY2025 稀釋股數 2,296 → FY2026 上半平均 2,268.9（−1.2%，" + TSQ2 + "）；公司再發行買回上限 150 億日圓／2,500 萬股（2026-11-01～2027-03-31；" + TSQ2 + "）", "價格 9,461 下 150 億日圓約 1.6M 股（0.07%）；FY2025、FY2026 各買回約 150 億日圓，年減 0.5–1.2%")

sec("G. 估值與對帳輸入")
for s in SC:
    inp(f"pk_{s}", f"情境 {s}：設備市場見頂年度（日曆年；循環位置規則用）", "年", None, PEAK[s], None, "Derived", PEAK_SRC[s], "設備股共用規則（tools/cycle_position.py）：FY(t) 早於見頂年＝中段、等於＝見頂、晚於＝谷底；設備市場模型的情境係數為 Assumed" if s == "A" else "同上")
inp("px", "股價（2026-11-13 收盤；一拆五後）", "¥", None, 9461, None, "Verified", "Yahoo Finance 日線收盤 2026-11-13：9,461（8035.T；markets/invest_map/data/closes_20261020_20261113.csv）", "價格只取自該 CSV（工作單第 3 節）")
inp("sa_fpe", "遠期 P/E（本 repo 同業表：stockanalysis 共識換算到 2026-11-13）", "x", None, 25.22, None, "Interested-party", "同 repo 現行模型 Val 同業表（amat、lrcx、accretech）：8035.T 25.22（2026-11-13）", "僅作並列參考（不進目標價）")
inp("cons_rev26", "共識營收 FY2026（2027年3月期）", "¥B", None, 3552.0, None, "Interested-party", YHS + "：FY3/27 營收 3.552T（23 位分析師）", "低於公司 10-30 指引 3,700（共識可能含指引前的估計）；第二來源：未找到")
inp("cons_eps26", "共識 EPS FY2026", "¥", None, 375.6, None, "Interested-party", YHS + "：FY3/27 EPS 375.6", "公司指引 381.69；第二來源：未找到")
inp("cons_rev27", "共識營收 FY2027（2028年3月期）", "¥B", None, 4487.0, None, "Interested-party", YHS + "：FY3/28 營收 4.487T", "第二來源：未找到")
inp("cons_eps27", "共識 EPS FY2027", "¥", None, 517.6, None, "Interested-party", YHS + "：FY3/28 EPS 517.6", "第二來源：未找到")

# 標準版 DCF 的 In 參數。評價日 2026-12-31（FY2026 只計 2027-01～03）；資產負債：2026-09-30（中間期決算短信）
JGB10 = ("日本財務省 國債金利情報（jgbcme.csv）10 年：2026-11-02 2.817%（https://www.mof.go.jp/english/policy/jgbs/reference/interest_rate/jgbcme.csv；沿用 accretech v0.3.7 同值，全庫日股一致）",
         "FRED IRLTLT01JPM156N（OECD 月均）：2026-09 3.050%；跨模型共用取值，維持一致、下次全面改版一併更新", "無風險利率（日圓：日本 10 年期國債殖利率，2026-11-02）")
DAM_GL = ("自算產業 β：JP 半導體設備 同業 8 家中位數（8035、6857、6146、6920、7735、6525、7729、6315；Yahoo Finance 月報酬 2021-11～2026-10，對 日經 225 ^N225 迴歸，Hamada 去槓桿、現金修正）無槓桿 β 1.91；以公司 D/E（市值）與稅率依 Hamada 重新加槓桿：β＝βU ×（1＋(1−t) × D/E）（同 accretech v0.3.7；全庫日股設備一致）",
          "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）：採用無槓桿 β 1.610；低／高：β_u ∓ beta_du（0.3），各自重新加槓桿（Analogy）")
CRP_JP = (0.0091, "Verified", "Damodaran 國家風險溢酬（2026-01-05）：日本 A1、CRP 0.91%（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html；沿用 accretech v0.3.7）", "非美國加 CRP、不設股權成本下限（2026-11-05 規則）")
RELEV = "*(1+(1-tax_w)*(debt_v+lease_v)/(px*sh_v))"
for _row in FD.in_rows(dict(
        date="2026-12-31（FY2026 只計 2027-01～03）", amt="¥B", shu="十億股", ccy="¥",
        rf=(0.02817,) + JGB10,
        rf_us=(0.05, "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（沿用 accretech v0.3.7 同值）", "FRED DGS10；Sens「美元口徑折現」案例：rf＝美國 10 年期 5.00%（CRP 照基準）、g＝MIN(3%, 5.00%)"),
        stub=dict(bs_date="2026-09-30", val_date="2026-12-31", bs_src=TSQ2 + "：2026-09-30 資產負債表與現金流量表",
                  div=("=83*2263.54876/1000", "Derived", TSQ2 + "：中間配當 415 円（一拆五前）＝83 円（一拆五後）／股，支付預定日 2026-11-27（在資產負債表日之後、評價日前）× 期末流通股數 2,263.549 百萬股（發行 2,340.164 − 庫藏 76.615）÷ 1000 ＝ 約 187.9 十億日圓",
                       "買回：2026-11-01～2027-03-31 上限 150 億日圓／2,500 萬股（公司授權，執行不保證），不屬已宣告確定金額，不扣；上半已買回 150.0（" + TSQ2 + "）已反映在 9 月底現金")),
        beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Derived") + DAM_GL,
        crp=CRP_JP,
        kd=("=rf+0.005", "Assumed", "rf ＋ 0.5 個百分點信用利差（公司無有息負債，kd 影響可忽略）", "債務 0，權重 0"),
        tax_w=("=1-g26_ni/g26_pbt", "Derived", "＝1 − 公司指引淨利 ÷ 指引稅前利益（865.0 ÷ 1,172.0）＝ 26.2%（" + PRQ2 + "）", "FY2022–FY2026 實績有效稅率 26.8%／24.5%／23.1%／22.9%／23.2%（" + SAI + "）；採指引隱含值（偏保守）；NOPAT 各年也用本稅率"),
        debt=(0, "Verified", TSQ2 + "：資產負債表無有息負債科目（yahoo 快照總債務 0；" + YHS + "）", "公司未揭露有息負債明細；視為 0"),
        lease=(0, "Assumed", TSQ2 + "：日本會計準則，經營租賃不入帳；租賃負債未單列", "—"),
        cash=("=529.762+50", "Verified", TSQ2 + "：2026-09-30 現金及存款 529,762＋有價證券 50,000 百萬日圓（現金流量表之現金及約當現金 578,601）", "沿用「現金＋有價證券」口徑；9 月底已含上半買回支出與中間配當前現金"),
        inv=(0, "Assumed", "投資有価証券（政策持股）本次未取得，維持不計入（偏保守）", "—"),
        mi=(0.0, "Derived", TSQ2 + "：純資產 2,194,666 − 歸屬母公司權益 2,166,848 − 新股預約權 27,817 ≈ 0（非支配股東持分極小）", "—"),
        dps=("=108+0.5*cons_eps27/2", "Assumed", TSQ2 + "：FY2026 期末配當預想 108 円（一拆五後；2027-06 前後支付）＋ FY2027 中間配當估計（配息率約 50%：共識 FY2027 EPS 517.6 × 50% ÷ 2＝約 129.4 円；2027-11 前後支付）",
             "公司配息政策：配息率約 50%、年配當下限 10 円（" + TRQ2 + "）；FY2027 中間配當為 Assumed"),
        sh=("=2263.54876*h1_eps/165.59/1000", "Derived", TSQ2 + "：期末流通股數 2,263.549 百萬股（發行 2,340.164 − 庫藏 76.615）× 基本 EPS 166.13 ÷ 稀釋 EPS 165.59（上半稀釋比；新股預約權以稀釋 EPS 比近似庫藏股法）÷ 1000", "股權報酬型 RSU／選擇權明細未取得；稀釋比 0.33%"),
        fc_da=(0.022, "=(52.412/1830.527+62.265/2431.568+81.345/2443.533)/3", 0.033, "Derived", SAC + "：近三年折舊攤銷 ÷ 營收：FY3/24 2.9%、FY3/25 2.6%、FY3/26 3.3%（FY2026 公司預想 100.0 ÷ 3,700＝2.7%）", "低／高 Assumed"),
        fc_cx=(0.04, "=g26_cx/g26_rev", 0.072, "Interested-party", PRQ2 + "：FY2026 資本支出計畫 190.0 ÷ 指引營收 3,700＝5.1%；高＝近三年平均 7.2%（FY3/24 6.4%、FY3/25 6.5%、FY3/26 8.5%，" + SAC + "）", "資本支出 FY2023–FY2026：74.4／121.8／162.1／216.0／190.0E（公司口徑；" + PRQ2 + "）；低 4% Assumed"),
        fc_cx_ss=("=fc_da*1.5", "Assumed", "長期穩態＝折舊 ÷ 營收 × 1.5（約 4.4%；擴產期後回落，仍高於折舊）", "一致性檢查：FY2032 資本支出 ≥ 折舊 ×（1＋g）（DCF 頁）"),
        fc_wc=(0.15, "=(246.380+35.478+44.287-37.029-3.449)/(2443.533-1399.102)", 0.40, "Derived",
               "近五年營運資金增加合計 ÷ 營收增量：營運資金變動 FY3/22 −246.4、FY3/23 −35.5、FY3/24 +37.0、FY3/25 −44.3、FY3/26 +3.4（" + SAC + "）合計 −285.7；營收 FY3/21 1,399.1 → FY3/26 2,443.5（增量 1,044.4）→ 27%", "低／高 15%／40%（Assumed）；存貨與預付款在擴產期占比高"),
        )):
    rows.append(_row)
inp("beta_u", "產業無槓桿 β（採用）：自算 JP 半導體設備 同業中位數 Blume 調整", "x", None, 1.610, None, "Derived", "自算產業 β：JP 半導體設備 同業中位數（含 8035 自身；見 In 頁 beta 列來源）", "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；與 accretech v0.3.7 同值")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務本金 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 稀釋股數 sh_v)", "與 WACC 權重同一口徑（無負債，D/E＝0）")
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）", "")

# 敏感度快照覆寫：TEL_OVERRIDE（「鍵=@lo／@hi／*倍數／數值；…」＝該鍵的基準欄改取 In 頁低／高欄、乘倍數或改數值；正式建置不設）
OVR = dict(kv.split("=", 1) for kv in os.environ.get("TEL_OVERRIDE", "").split(";") if "=" in kv)
r = 4
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, s1, s2 = srow
    if key in OVR:
        ov = OVR[key]
        if ov.startswith("*"):
            assert isinstance(base, (int, float)), (key, base); _new = base * float(ov[1:])
        else:
            _new = {"@lo": lo, "@hi": hi}[ov] if ov.startswith("@") else float(ov)
        assert _new is not None, (key, ov)
        base = _new
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
# 目標價參數（黃底；Assumed，Andy 可改）
r += 1
IN.cell(r, 1, "J. 目標價（DCF 與 Comparable 加權）與 FY2030–32 收斂法參數（黃底＝Assumed，Andy 可改）").font = BOLD; r += 1
def inrow(key, item, unit, base, tag, s1, s2, fmt=None, yellow=False):
    global r
    if key in OVR: base = float(OVR[key])   # 敏感度快照（情境權重案例）
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    c = IN.cell(r, 5, base)
    if not (isinstance(base, str) and base.startswith("=")): c.font = BLUE
    if fmt: c.number_format = fmt
    if yellow: c.fill = KEY
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    defname(key, f"In!$E${r}"); r += 1
for s_, w_ in (("A", 0.25), ("B", 0.5), ("C", 0.25)):
    inrow(f"tp_w{s_}", f"目標價情境權重 {s_}（{SCN[s_]}）", "%", w_, "Assumed", "個股 DCF 標準規格（預設 25／50／25）；Andy 可改", "三者合計應為 100%", "0%", True)
inrow("tp_g", "PL 延伸期營收年增的收斂目標（情境 B；FY2032）", "%", 0.04, "Assumed", "Andy 可改（工作單預設 4%）；DCF 永續成長率見 M2 節 g_tv", "情境 A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt（PL 頁延伸區塊）", "0.0%", True)
inrow("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", 0.5, "Assumed", "Andy 可改（2026-10-25：DCF／Comparable 各半）", "—", "0%", True)
inrow("nim10", "十年平均淨利率（FY2016–FY2025，各年淨利 ÷ 營收的平均）", "%", "=AVERAGE(Hist!F7:F16)", "Derived", "Hist 頁（stockanalysis、globalbanking 彙整的年度資料；FY2017–FY2020 營收為約數）", "收斂法 B、C 情境的淨利率目標", "0.0%")
inrow("cons_tp", "市場共識目標價（賣方平均）", "¥", 15252, "Interested-party", YHS + "：目標價均值 15,252（23 位）", "第二來源：未找到（stockanalysis 目標價頁 2026-11 未取得）", "#,##0")
inrow("cons_feps", "共識前瞻 EPS（差距拆解用）", "¥", "=cons_eps27", "Interested-party", "＝共識 FY2027（2028年3月期）EPS 517.6（Yahoo 快照）；本模型 12 個月後前瞻 EPS 約為 FY2027（5／12）與 FY2028（7／12）的混合，較共識晚約半年", "共識 FY2028 EPS：未取得", "#,##0")
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 80, 80]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

# ---------------- Hist：FY2015–FY2025 年度實績 ----------------
FB = "finboard.net JP:8035 損益表（百萬日圓；2026-11-17 讀取；第三方轉錄；網站 FY 標籤＝截止年）"
H = wb.create_sheet("Hist")
H["A1"] = ("Hist — 年度實績（十億日圓；EPS 為日圓，已調整為 2026-10-01 一拆五之後；FY2022 以前再除以 3（2023 年一拆三）＝除以 15，FY2023 起除以 5）。"
           "營收、營業利益、淨利：finboard（來源1，第三方轉錄），FY2022 起與 stockanalysis 逐年相符；FY2025 與公司決算短信相符。TEL 只有單一應報導部門，無部門歷史。用途：十年平均淨利率、Hist 迴歸、HistIS")
H["A2"] = "會計年度 FY(t)＝t 年 4 月～t+1 年 3 月（例：FY2025＝2026年3月期；公司自稱 FY2026）。EPS：finboard 原值為各期當時股數口徑，已依分割比例還原（Derived）；FY2025 為公司決算短信基本 EPS"
hdr(H, 5, ["會計年度", "營收", "營業利益", "淨利", "EPS（基本，還原後）", "淨利率", "營業利益率", "標記", "毛利", "稅前利益", "", "", "來源1", "來源2"])
HIST = [("FY2015（2016-03）", 663.948, 116.788, 77.936, "=461.10/15", 267.209, 119.399),
        ("FY2016（2017-03）", 799.719, 155.697, 115.248, "=702.26/15", 322.291, 157.549),
        ("FY2017（2018-03）", 1130.728, 281.172, 204.399, "=1245.48/15", 475.032, 280.737),
        ("FY2018（2019-03）", 1278.240, 310.571, 248.228, "=1513.58/15", 526.183, 321.662),
        ("FY2019（2020-03）", 1127.286, 237.292, 185.206, "=1170.57/15", 451.941, 244.979),
        ("FY2020（2021-03）", 1399.102, 320.685, 242.941, "=1562.20/15", 564.945, 322.103),
        ("FY2021（2022-03）", 2003.805, 599.271, 437.076, "=2807.84/15", 911.822, 601.724),
        ("FY2022（2023-03）", 2209.025, 617.724, 471.584, "=1007.82/5", 984.408, 625.185),
        ("FY2023（2024-03）", 1830.527, 456.264, 363.963, "=783.75/5", 830.270, 463.185),
        ("FY2024（2025-03）", 2431.568, 697.321, 544.133, "=1182.40/5", 1146.288, 707.727),
        ("FY2025（2026-03）", "=a25_rev", "=a25_oi", "=a25_ni", "=a25_eps", "=a25_gp", "=a25_pbt")]
for i, (fy, rv, oi, ni, eps, gp, pbt) in enumerate(HIST, 6):
    H.cell(i, 1, fy)
    for j, v in enumerate((rv, oi, ni, eps), 2):
        H.cell(i, j, v); H.cell(i, j).number_format = "#,##0.000" if j < 5 else "#,##0.00"
        if not isinstance(v, str): H.cell(i, j).font = BLUE
    H.cell(i, 6, f"=D{i}/B{i}").number_format = "0.0%"; H.cell(i, 7, f"=C{i}/B{i}").number_format = "0.0%"
    H.cell(i, 9, gp).number_format = "#,##0.000"; H.cell(i, 10, pbt).number_format = "#,##0.000"
    for j in (9, 10):
        if not isinstance(H.cell(i, j).value, str): H.cell(i, j).font = BLUE
    last = fy.startswith("FY2025")
    H.cell(i, 8, "Verified（公司決算短信，In 頁）" if last else ("Verified（finboard 與 stockanalysis 逐年相符）" if i >= 12 else "Verified（單一第三方來源 finboard；第二來源：未找到）"))
    H.cell(i, 13, "公司 FY2027 Q2 決算短信前期欄（In 頁 a25_*）" if last else FB)
    H.cell(i, 14, SAI if i >= 12 else "未找到（stockanalysis 免費頁只列 5 年；公司十年摘要未取得）")
H.cell(18, 1, "檢查：FY2025 列 − In 頁 a25_*（應為 0；FY2025 列直接引用 In 頁，此處為結構檢查）")
H["B18"] = "=ABS(B16-a25_rev)+ABS(C16-a25_oi)+ABS(D16-a25_ni)+ABS(E16-a25_eps)"; H["B18"].number_format = "0.000"
H.cell(20, 1, "迴歸與 Analogy（In 頁同值）").font = BOLD
for i, (lab_, f, fm) in enumerate((("營業利益 ＝ 固定 ＋ 邊際 × 營收：邊際（FY2015–FY2025，供對照）", "=SLOPE(C6:C16,B6:B16)", "0.000"), ("  R²", "=RSQ(C6:C16,B6:B16)", "0.000"),
                                  ("  固定（¥B）", "=INTERCEPT(C6:C16,B6:B16)", "0.00"),
                                  ("邊際（只用 FY2021–FY2025）", "=SLOPE(C11:C16,B11:B16)", "0.000"),
                                  ("兩點邊際（FY2025→FY2026 錨點；模型採用，In 頁 mm）", "=mm", "0.000"),
                                  ("營收年複合 FY2015→FY2025", "=(B16/B6)^(1/10)-1", "0.0%"), ("營收年複合 FY2020→FY2025", "=(B16/B11)^(1/5)-1", "0.0%"),
                                  ("十年平均淨利率 FY2016–FY2025", "=nim10", "0.0%"),
                                  ("最大單年營收下滑（FY2019，2020-03）", "=B10/B9-1", "0.0%"), ("FY2023（2024-03）營收下滑", "=B14/B13-1", "0.0%")), 21):
    H.cell(i, 1, lab_); H.cell(i, 2, f).number_format = fm
H.cell(32, 1, "讀法：TEL 的歷史營業利益對營收的歷史邊際約 35–40%（上表）；兩點邊際 43.3% 高於歷史，反映公司指引含價格最適化與組合改善（公司說法，Interested-party）。歷史迴歸只作對照，不進模型。")
H.column_dimensions["A"].width = 52
for c in "BCDEFGIJ": H.column_dimensions[c].width = 12
H.column_dimensions["H"].width = 40; H.column_dimensions["M"].width = 50; H.column_dimensions["N"].width = 60

# ---------------- Q：季度實績 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = ("Q — 季度實績（十億日圓；EPS 為日圓，一拆五後）。會計年度 4–3 月。損益：stockanalysis 季度表（來源1；百萬日圓）；新設備銷售：公司 FY2027 Q2 決算說明資料（來源2）。"
           "FY26 Q2＝中間期（上半）決算短信 − 第一季（Derived）。TEL 只有單一應報導部門；新設備與 Field Solutions 為公司揭露的營收構成（非應報導部門）")
hdr(Q, 3, ["項目", "FY25 Q1", "FY25 Q2", "FY25 Q3", "FY25 Q4", "FY2025", "FY26 Q1", "FY26 Q2（實績；上半 − 第一季）", "標記", "來源1", "來源2／說明", "來源誘因（Interested-party 列；混合來源列寫其他欄位）"])
QD = [("營收", [549.586, 630.082, 552.047, 711.818, None, 732.388, "=h1_rev-G4"], "#,##0.0", SAI + "：549,586／630,082／552,047／711,818；732,388", "FY26 Q2＝上半 1,667.455 − 732.388 ＝ 935.067（" + PRQ2 + "：935.0）"),
      ("營業利益", [144.695, 158.459, 116.139, 205.644, None, 211.409, "=h1_oi-G5"], "#,##0.0", SAI, "FY26 Q2＝上半 506.776 − 211.409 ＝ 295.367"),
      ("淨利", [117.801, 123.825, 118.538, 214.290, None, 164.341, "=h1_ni-G6"], "#,##0.0", SAI, "FY25 Q4 淨利 214.3 高於營業利益 205.6：含營業外與特別利益（FY2025 稅前 748.2 對營業利益 624.9）"),
      ("EPS（基本）", [51.30, 53.82, 51.52, 93.69, None, 72.03, "=h1_eps-G7"], "#,##0.00", SAI + "（每季 EPS，口徑為網站未標明；四季合計 250.33 對公司基本 250.91，差 0.2%）", "FY26 Q2＝上半 166.13 − 72.03 ＝ 94.10"),
      ("新設備銷售（SPE new equipment）", [395.0, 455.0, 385.1, 540.1, None, 531.3, "=ne_h1-G8"], "#,##0.0", PRQ2 + "：新設備銷售逐季（FY26＝公司稱 FY2026；Q1 395.0、Q2 455.0、Q3 385.1、Q4 540.1；FY27 Q1 531.3、Q2 734.4）", "公司揭露；FY26 Q2＝上半新設備 1,265.7 − 531.3 ＝ 734.4（與公司揭露相符）"),
      ("Field Solutions 與其他（＝營收 − 新設備）", ["=B4-B8", "=C4-C8", "=D4-D8", "=E4-E8", None, "=G4-G8", "=H4-H8"], "#,##0.0", "Derived：營收 − 新設備銷售", "公司揭露 Field Solutions 逐季 141.2／160.3／162.6／162.8（FY25），其餘為二手設備、FPD 設備等；殘差含 FS 以外項目")]
for i, (lab_, vals, fm, s1, s2) in enumerate(QD, 4):
    Q.cell(i, 1, lab_)
    for j, v in enumerate(vals):
        col = "BCDEFGH"[j]
        if v is None: continue
        c = Q[f"{col}{i}"]; c.value = v; c.number_format = fm
        if not isinstance(v, str): c.font = BLUE
    Q[f"F{i}"] = f"=SUM(B{i}:E{i})"; Q[f"F{i}"].number_format = fm
    Q.cell(i, 9, "Verified" if i <= 7 else ("Interested-party" if i == 8 else "Derived"))
    Q.cell(i, 10, s1 + ("；標記說明：損益為第三方轉錄，與公司決算短信算術互核；FY26 Q2＝上半 − 第一季（Derived）" if i <= 7 else "")); Q.cell(i, 11, s2)
    if i == 8: Q.cell(i, 12, "發行公司自揭營收構成（說明資料）：傾向呈現正面；以事後決算檢驗")
Q.cell(10, 1, "營業利益率"); [Q.cell(10, j, f"={c}5/{c}4").__setattr__("number_format", "0.0%") for j, c in zip(range(2, 9), "BCDEFGH")]
Q.cell(11, 1, "新設備占營收"); [Q.cell(11, j, f"={c}8/{c}4").__setattr__("number_format", "0.0%") for j, c in zip(range(2, 9), "BCDEFGH")]
Q.cell(12, 1, "檢查：FY2025 四季合計 − 年度實績（營收、營業利益、淨利；應約為 0）")
Q["B12"] = "=F4-a25_rev"; Q["C12"] = "=F5-a25_oi"; Q["D12"] = "=F6-a25_ni"
for c in "BCD": Q[f"{c}12"].number_format = "0.000"
Q.cell(13, 1, "參考：FY2025 新設備四季合計 − In 頁 ne_25（公司揭露的季數與半年數各自四捨五入，差約 0.2；不計入核對）"); Q["B13"] = "=F8-ne_25"; Q["B13"].number_format = "0.000"
Q.cell(14, 1, "上半營收占公司全年指引"); Q["B14"] = "=(G4+H4)/g26_rev"; Q["B14"].number_format = "0.0%"; Q["C14"] = "FY2025 上半占全年："; Q["D14"] = "=(B4+C4)/a25_rev"; Q["D14"].number_format = "0.0%"
Q.cell(16, 1, "公司揭露的結構（" + QAQ2 + "；Interested-party）：WFE 市場 CY2026 約 160B 美元（上修）、CY2027 超過 210B；新設備下半 1,560 對上半 1,265.7；客戶要求提前交貨，Q4 毛利率若提前認列裝機可到 50%；中國占比上半約 34%。")
Q.column_dimensions["A"].width = 44
for c in "BCDEFGH": Q.column_dimensions[c].width = 12
Q.column_dimensions["I"].width = 22; Q.column_dimensions["J"].width = 60; Q.column_dimensions["K"].width = 70

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
CYC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2031)}   # C＝CY2025 … H＝CY2030
FYC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2030)}   # C＝FY2025 … G＝FY2029
MK = ("la", "lo", "of", "dr", "nd")
MKN = {"la": "先進邏輯設備〔AI〕", "lo": "其他代工與邏輯設備", "of": "其他前段設備", "dr": "DRAM（含 HBM）設備〔AI〕", "nd": "NAND 設備"}
DEFP = dict(tail="g_tail", chn="chn", nioi="nioi", anc="g_anc", scn="A", exc=None, chh=None, gfso=None, dm="dm", madj=None, bb="bb", ovr=None, den26=None)
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    if SIL.COORD_ALT and P["ovr"] is None and s != "A":   # 三層協調 C4 另一側讀數建置（tools/coord/alt_run.py）：情境 B、C 於 CY2027 起設備市場 AI 分項（先進邏輯〔AI〕、DRAM〔AI〕）× wta
        P["ovr"] = {(k, y): f"={k}{s}_{y % 100}*wta_{y % 100}" for k in ("la", "dr") for y in (2027, 2028, 2029)}
        P["den26"] = {k: f"(0.75*{k}{s}_26+0.25*{k}{s}_27)" for k in ("la", "dr")}   # FY2026 指數分母維持未縮放（錨點）
    exc = P["exc"] or f"exc_{s}"; chh = P["chh"] or f"chh_{s}"; gfso = P["gfso"] or f"gfso_{s}"
    madj = P["madj"] or ("madjC" if s == "C" else "0")
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "CY2025／FY2025", "CY2026／FY2026", "CY2027／FY2027", "CY2028／FY2028", "CY2029／FY2029", "CY2030", "", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="0.00", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 10, note); R[key] = r; r += 1
    ws.cell(r, 1, f"A. 設備市場（日曆年；markets/equipment 情境 {s}）").font = BOLD; r += 1
    for key in MK:
        row(key, MKN[key], "US$B", {**{CYC[y]: ((P['ovr'] or {}).get((key, y)) or f"={key}{s}_{y % 100}") for y in range(2025, 2030)}, "H": f"=G{r}*(1+{P['tail']})"}, fmt="#,##0.0",
            note="CY2030＝CY2029 ×（1＋g_tail）" if key == "la" else "")
    ws.cell(r, 1, "B. 會計年度指數（FY(t)＝3/4 × CY(t)＋1/4 × CY(t+1)；FY2026＝1）").font = BOLD; r += 1
    fw = lambda key, y: f"(0.75*{CYC[y]}{R[key]}+0.25*{CYC[y+1]}{R[key]})"
    den = lambda key: (P['den26'] or {}).get(key) or fw(key, 2026)
    row("idr", "DRAM 需求指數", "x", {FYC[y]: f"={fw('dr', y)}/{den('dr')}" for y in range(2025, 2030)}, fmt="0.000")
    row("ind", "NAND 需求指數", "x", {FYC[y]: f"={fw('nd', y)}/{den('nd')}" for y in range(2025, 2030)}, fmt="0.000")
    row("inm", "非記憶體需求指數（先進邏輯＋其他代工邏輯＋其他前段，以金額加總）", "x",
        {FYC[y]: f"=({fw('la', y)}+{fw('lo', y)}+{fw('of', y)})/({den('la')}+{den('lo')}+{den('of')})" for y in range(2025, 2030)}, fmt="0.000")
    row("exf", "公司特有：超額成長係數（FY2027 起）", "x", {FYC[y]: f"=(1+{exc})^{y - 2026}" for y in range(2026, 2030)}, fmt="0.000")
    row("chf", "公司特有：中國係數（1 ＋ 中國占比 × 中國變化 × 生效比例）", "x", {"D": "=1", "E": f"=1+{P['chn']}*{chh}*0.5", "F": f"=1+{P['chn']}*{chh}", "G": f"=1+{P['chn']}*{chh}"}, fmt="0.000")
    ws.cell(r, 1, "C. 營收（FY2025 實績；FY2026 公司指引 ×（1＋g_anc）；FY2027 起驅動推導）").font = BOLD; r += 1
    def ne_row(key, label, a25, a26, idx):
        row(key, label, "¥B", {"C": f"={a25}", "D": f"={a26}*(1+{P['anc']})", **{FYC[y]: f"=$D{r}*{FYC[y]}{R[idx]}*{FYC[y]}{R['exf']}*{FYC[y]}{R['chf']}" for y in range(2027, 2030)}}, fmt="#,##0.0")
    ne_row("ne_dr", "新設備：DRAM", "ne25_dr", "ne26_dr", "idr")
    ne_row("ne_nd", "新設備：非揮發性記憶體（NAND）", "ne25_nd", "ne26_nd", "ind")
    ne_row("ne_nm", "新設備：非記憶體（邏輯、代工、其他）", "ne25_nm", "ne26_nm", "inm")
    row("ne", "新設備銷售（SPE new equipment）", "¥B", {FYC[y]: f"={FYC[y]}{R['ne_dr']}+{FYC[y]}{R['ne_nd']}+{FYC[y]}{R['ne_nm']}" for y in range(2025, 2030)}, fmt="#,##0.0",
        note="FY2025 依公司揭露的終端市場占比拆分；FY2026 依公司指引上半實績＋下半預想的占比拆分")
    row("fso", "Field Solutions 與其他（零件、服務、二手與改裝設備等）", "¥B", {"C": "=fso_25", "D": f"=(g26_rev-g26_ne)*(1+{P['anc']})", **{FYC[y]: f"={FYC[y-1]}{r}*(1+{gfso})" for y in range(2027, 2030)}}, fmt="#,##0.0")
    row("rev", "營收", "¥B", {FYC[y]: f"={FYC[y]}{R['ne']}+{FYC[y]}{R['fso']}" for y in range(2025, 2030)}, fmt="#,##0.0", note="FY2026＝公司指引（g_anc＝0 時）")
    row("nesh", "新設備占營收", "%", {FYC[y]: f"={FYC[y]}{R['ne']}/{FYC[y]}{R['rev']}" for y in range(2025, 2030)}, fmt="0.0%")
    ws.cell(r, 1, "D. 損益（營業利益＝FY2026 錨點 ＋ 邊際營業利益率 × 營收變化）").font = BOLD; r += 1
    row("oi", "營業利益", "¥B", {"C": "=a25_oi", "D": f"=g26_oi+mm*(D{R['rev']}-g26_rev)",
                              **{FYC[y]: f"=$D{r}+(mm+{P['dm']})*({FYC[y]}{R['rev']}-$D{R['rev']})" + (f"+{madj}*{FYC[y]}{R['rev']}" if y >= 2028 else "") for y in range(2027, 2030)}}, fmt="#,##0.0",
        note="mm＝兩點邊際（In 頁）；情境 C 自 FY2028 另加 madjC × 營收")
    row("oim", "營業利益率（結果）", "%", {FYC[y]: f"={FYC[y]}{R['oi']}/{FYC[y]}{R['rev']}" for y in range(2025, 2030)}, fmt="0.0%")
    row("ni", "淨利", "¥B", {"C": "=a25_ni", "D": f"=D{R['oi']}*g26_ni/g26_oi", **{FYC[y]: f"={FYC[y]}{R['oi']}*{P['nioi']}" for y in range(2027, 2030)}}, fmt="#,##0.0",
        note="FY2026＝營業利益 × 指引淨利÷營業利益；FY2027 起 × nioi")
    row("sh", "平均股數", "百萬股", {"C": "=a25_ni/a25_eps*1000", "D": "=shares", **{FYC[y]: f"=shares*(1-{P['bb']})^{y - 2026}" for y in range(2027, 2030)}}, fmt="#,##0.0", note="FY2027 起每年依 bb 減少（淨買回）")
    row("eps", "EPS", "¥", {"C": "=a25_eps", **{FYC[y]: f"={FYC[y]}{R['ni']}*1000/{FYC[y]}{R['sh']}" for y in range(2026, 2030)}}, fmt="#,##0.0")
    row("chk", "檢查：FY2025 營收、營業利益 − 實績", "¥B", {"C": f"=ABS(C{R['rev']}-a25_rev)+ABS(C{R['oi']}-a25_oi)"}, fmt="0.000")
    R["end"] = r
    return R

# ---------------- PL：三情境 ----------------
PL = wb.create_sheet("PL")
PL["A1"] = ("PL — 東京威力科創 損益（¥B；FY2025 實績；FY2026 錨定公司指引（2026-10-30 首次發布通期；上半為實績）；FY2027–FY2029 驅動推導；FY2030 起收斂法）。"
            "頂端為摘要，下方為三情境各一個完整推導區塊")
PB = {}; r0 = 60
TITLES = {"A": "情境 A（樂觀）：設備市場 A（見頂 2029）；新設備超額成長 +2%/年；中國持平；FS 與其他年增 12%",
          "B": "情境 B（基準）：設備市場 B（2028 持平、2029 −5%）；超額成長 0；中國 −5%；FS 與其他年增 8%",
          "C": "情境 C（保守）：設備市場 C（2028 −20%）；超額成長 −2%/年；中國 −20%；FS 與其他年增 3%；營益率額外 −1.5pt"}
for s in SC:
    PB[s] = engine(PL, r0, TITLES[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD; sr += 1
    for key, lab_, unit, fmt in (("rev", "營收", "¥B", "#,##0.0"), ("ne_dr", "  新設備：DRAM", "¥B", "#,##0.0"), ("ne_nd", "  新設備：NAND", "¥B", "#,##0.0"), ("ne_nm", "  新設備：非記憶體", "¥B", "#,##0.0"),
                                 ("fso", "  Field Solutions 與其他", "¥B", "#,##0.0"), ("oi", "營業利益", "¥B", "#,##0.0"), ("oim", "營業利益率", "%", "0.0%"), ("ni", "淨利", "¥B", "#,##0.0"), ("eps", "EPS", "¥", "#,##0")):
        PL.cell(sr, 1, lab_); PL.cell(sr, 2, unit)
        for y in range(2025, 2030):
            PL[f"{FYC[y]}{sr}"] = f"={FYC[y]}{PB[s][key]}"; PL[f"{FYC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對：三情境 FY2025 − 實績（營收與營業利益；應為 0）"); PL[f"C{sr}"] = "=" + "+".join(f"C{PB[s]['chk']}" for s in SC); PL[f"C{sr}"].number_format = "0.000"; sr += 1
PL.cell(sr, 1, "核對：FY2026 營收、營業利益、EPS − 公司指引（g_anc＝0 時應為 0）")
PL[f"C{sr}"] = f"=ABS(D{SUMR[('A', 'rev')]}-g26_rev)+ABS(D{SUMR[('A', 'oi')]}-g26_oi)"; PL[f"D{sr}"] = f"=D{SUMR[('A', 'eps')]}-g26_eps"
PL[f"C{sr}"].number_format = "0.000"; PL[f"D{sr}"].number_format = "0.00"; CHK26 = sr
# 延伸到 FY2032（收斂法）
p0 = PB["C"]["end"] + 3
PL.cell(p0, 1, "延伸到 FY2032：FY2025–FY2029 引用本頁頂端摘要；灰底 FY2030–FY2032 為收斂法，非驅動因子推導").font = BOLD
PL.cell(p0 + 1, 1, "收斂法：營收年增由 FY2029 年增線性收斂到 FY2032 的長期成長（A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt）；淨利率由 FY2029 線性收斂（A＝FY2029、B＝FY2029 與十年平均的中點、C＝十年平均 nim10）；股數維持（EPS 與淨利同比例）")
EC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2033)}
hdr(PL, p0 + 2, ["情境（延伸）", "單位", "FY2025A", "FY2026", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
EXT = {}; pr = p0 + 3; gadd = {"A": "+0.02", "B": "", "C": "-0.02"}
for s in SC:
    PL.cell(pr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD; pr += 1
    R = {k: pr + i for i, k in enumerate(("rev", "g", "nim", "ni", "eps"))}
    for k, lab in (("rev", "營收"), ("g", "  營收年增"), ("nim", "  淨利率（淨利 ÷ 營收）"), ("ni", "淨利（收斂法延伸）"), ("eps", "EPS（延伸）")):
        PL.cell(R[k], 1, lab); PL.cell(R[k], 2, {"rev": "¥B", "ni": "¥B", "eps": "¥"}.get(k, "%"))
    for y in range(2025, 2033):
        c = EC[y]; pc = EC.get(y - 1)
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
        for kk, fm in (("rev", "#,##0.0"), ("g", "0.0%"), ("nim", "0.0%"), ("ni", "#,##0.0"), ("eps", "#,##0")):
            if PL[f"{c}{R[kk]}"].value is not None: PL[f"{c}{R[kk]}"].number_format = fm
    PL.cell(R["rev"], 11, "FY2030 起為收斂法，非驅動因子推導")
    EXT[s] = R; pr += 6
PL.column_dimensions["A"].width = 56; PL.column_dimensions["B"].width = 10; PL.column_dimensions["J"].width = 46; PL.column_dimensions["K"].width = 36
for c in "CDEFGHI": PL.column_dimensions[c].width = 13

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — 估值（日圓）。目標價＝DCF 與 Comparable 加權（第 0 節）；跨公司規則值（十年年底 P/E 依循環位置取分位：見頂 25 分位、中段中位數、谷底 75 分位）保留為參考（第 5–6 節）"
V["A2"] = "倍數定義：本模型不自算十年年底 P/E（取不到 2016–2025 年底收盤價的可驗證來源）；跨公司規則值改用 Bloomberg 前瞻 P/E 的十年年度窗口中位數（本 repo tools/data/fwd_mult_hist.csv；與相對倍數調整同一來源）"
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
V.cell(t0 + 1, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD
hdr(V, t0 + 2, ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"])
V.cell(t0 + 3, 1, "情境 EPS（FY2027–FY2032）"); V[f"E{t0 + 3}"] = "PL 頁延伸區塊；FY2030 起為收斂法"
EPSR = {}
for i, y in enumerate(YRS6):
    tset(t0 + 4 + i, f"eps{y}", f"  FY{y}", [f"=PL!{EC[y]}{EXT[s]['eps']}" for s in SC], "#,##0"); EPSR[y] = t0 + 4 + i
rr = t0 + 10
V.cell(rr, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）"); V[f"E{rr}"] = "模型無折舊與資本支出項目，故不用（淨利＋折舊－資本支出）；公司 FY2026 資本支出指引 150 億、折舊 62 億（擴產期），近年股東現金流可能低於此式（偏高估）"
tset(rr + 1, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）")
RV = rr + 1; CFR = {}
for i, y in enumerate(YRS6):
    tset(rr + 2 + i, f"cf{y}", f"  FY{y}", [f"={c}{EPSR[y]}*(1-$B${RV})" for c in "BCD"], "#,##0"); CFR[y] = rr + 2 + i
c1, c6 = CFR[2027], CFR[2032]
tset(rr + 8, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=MAX(0,SUMPRODUCT({c}{c1}:{c}{c6},1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}{c6}*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6)" for c in "BCD"], "#,##0",
     "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底，t＝1 對 FY2027；低於 0 取 0；會計年度 4–3 月，FY2027＝2027-04～2028-03 仍以 t＝1 折現（略高估）")
tset(rr + 9, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}{rr + 8}*(1+tp_r)" for c in "BCD"], "#,##0")
tset(rr + 10, "dcftp", "DCF 目標價（情境加權）", [f"=tp_wA*B{rr + 9}+tp_wB*C{rr + 9}+tp_wC*D{rr + 9}", None, None], "#,##0", "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁 tp_wA／tp_wB／tp_wC")
q = rr + 12
V.cell(q, 1, "0.2 Comparable（市場主流同業前瞻 P/E：前段設備同業；非 A 股公司，同業不放 A 股）").font = BOLD
hdr(V, q + 1, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = [("應用材料（Applied Materials）", "AMAT", 21.02, "2026-11-13", "本 repo 現行模型 Val 同業表（lrcx、klac、asml；stockanalysis.com Forward PE 共識，換算到 2026-11-13 收盤）"),
         ("科林研發（Lam Research）", "LRCX", 21.98, "2026-11-13", "同上（amat、klac、asml 模型 Val 同業表）"),
         ("科磊（KLA）", "KLAC", 27.88, "2026-11-13", "同上"),
         ("ASML（美國 ADR）", "ASML", 23.53, "2026-11-13", "同上"),
         ("Disco", "6146.T", 25.06, "2026-11-13", "同上（accretech、advantest 模型 Val 同業表）"),
         ("Lasertec", "6920.T", 32.31, "2026-11-13", "同上（klac、accretech 模型 Val 同業表）")]
for i, (nm, tk, pe, dt, src) in enumerate(PEERS, q + 2):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
p1 = q + 2; pl_ = q + 1 + len(PEERS)
V.cell(pl_ + 1, 1, "並列參考（不進目標價）：東京威力科創自身前瞻 P/E"); V.cell(pl_ + 1, 3, "=sa_fpe").number_format = "0.0"; V.cell(pl_ + 1, 4, "2026-11-13"); V.cell(pl_ + 1, 5, "stockanalysis.com（本 repo 同業表 8035.T）；自身近十年前瞻 P/E 區間見第 2 節")
q = pl_ + 2
tset(q, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C{p1}:C{pl_})", None, None], "0.0", "前段設備同業：美國四家（應用材料、科林、科磊、ASML）與日本兩家（Disco、Lasertec）；Advantest（測試）、Kokusai（沉積）不納入以免混入測試與單一製程")
tset(q + 1, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=" + "+".join(f"tp_w{s}*(5/12*{c}{EPSR[2027]}+7/12*{c}{EPSR[2028]})" for s, c in zip(SC, "BCD")), None, None], "#,##0",
     "約 2027-10 時的未來 12 月（2027-11～2028-10）≈ 5／12 × FY2027（2027-04～2028-03）＋7／12 × FY2028（會計年度 4–3 月，依月份調整）")
tset(q + 2, "comptp", "Comparable 目標價", [f"=B{q}*B{q + 1}", None, None], "#,##0", "同業中位數 × 上列 EPS")
q += 4
V.cell(q, 1, "0.3 結論").font = BOLD; T["h03"] = q
hdr(V, q + 1, ["項目", "數值", "", "", "說明"]); T["hd3"] = q + 1
q += 2
tset(q, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B{T['comptp']}", None, None], "#,##0", "tp_w_dcf × DCF 目標價＋（1 − tp_w_dcf）× Comparable 目標價；tp_w_dcf 見 In 頁")
tset(q + 1, "comp2", "Comparable 目標價（市場主流同業法）", [f"=B{T['comptp']}", None, None], "#,##0", "引用 0.2")
tset(q + 2, "cons", "市場共識目標價", ["=cons_tp", None, None], "#,##0", "Yahoo Finance 快照（2026-11-17）：23 位分析師平均 15,252；第二來源未找到")
tset(q + 3, "upw", "上檔空間：加權目標價", [f"=B{q}/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-11-13 收盤 9,461）")
tset(q + 4, "upc", "上檔空間：Comparable 目標價", [f"=B{q + 1}/px-1", None, None], "0.0%")
tset(q + 5, "geps", "與市場共識目標價的差距：EPS 因素", [f"=B{T['feps']}/cons_feps-1", None, None], "0.0%", "本模型 12 個月後前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝FY2027（2028年3月期）共識 517.6（Recon），比本模型 12 個月後前瞻 EPS 晚約半年，EPS 因素因此偏小")
tset(q + 6, "gmul", "與市場共識目標價的差距：倍數因素", [f"=(cons_tp/cons_feps)/B{T['pmed']}-1", None, None], "0.0%")
V[f"E{q + 6}"] = f'="共識目標價 ÷ 共識前瞻 EPS 隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B{T["pmed"]},"0.0")&"x"'
tset(q + 7, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D{T["dcf12"]},"#,##0")&"～"&TEXT(B{T["dcf12"]},"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
q8 = q + 8
T["pr"] = q8  # 公式在 P/E 表建好後填入
V.cell(q8, 1, "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）")
V.cell(q + 9, 1, "目標價方法"); T["meth"] = q + 9
V[f"B{q + 9}"] = ("DCF（標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，以 WACC 折現到 2026-12-31（期中慣例；FY2026 只計 2027-01～03），FY2026–32 明確期（FY2030 起收斂法），終值＝NOPAT(FY2033) ×（1 − g ÷ RONIC）÷（WACC − g）。"
                  "Comparable：美國四家前段設備商與 Disco、Lasertec 的前瞻 P/E 中位數 × 12 個月後前瞻 EPS（再依相對倍數調整，第 0.5 節）。"
                  "偏誤：同業倍數取共識前瞻（本期為景氣上行期，共識傾向落後實際）；TEL 一拆五後股價 9,461，市場已定價 FY2027 EPS 約 520（共識）；DCF 已扣資本支出與營運資金，擴產期現金流偏低。")
for k in ("dcftp", "comptp", "tpw"): V[f"B{T[k]}"].fill = KEY
# 1. 價格與市場隱含倍數
b0 = q + 11
V.cell(b0, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, b0 + 1, ["項目", "A", "B", "C", "說明"])
fr = b0 + 2
V.cell(fr, 1, "現價（2026-11-13 收盤）"); V.cell(fr, 2, "=px").number_format = "#,##0"; V.cell(fr, 5, "Yahoo Finance 日線收盤（markets/invest_map/data/closes_20261020_20261113.csv）；一拆五後"); T["px"] = fr; fr += 1
V.cell(fr, 1, "現價 ÷ FY2026 EPS（錨點）")
for j, s in enumerate(SC): V.cell(fr, 2 + j, f"=px/PL!D{SUMR[(s, 'eps')]}").number_format = "0.0"
V.cell(fr, 5, "FY2026 三情境相同（公司指引）"); fr += 1
for y in (2027, 2028, 2029):
    V.cell(fr, 1, f"現價 ÷ FY{y} EPS（模型）")
    for j, s in enumerate(SC): V.cell(fr, 2 + j, f"=px/PL!{FYC[y]}{SUMR[(s, 'eps')]}").number_format = "0.0"
    T[f"pe{y}"] = fr; fr += 1
for nm, f, note in (("P/E：現價 ÷ FY2025 EPS 250.91", "=px/a25_eps", "上一個會計年度"),
                    ("P/E：現價 ÷ 近 12 月 EPS（FY25 Q3–Q4＋FY26 Q1–Q2）", "=px/(Q!D7+Q!E7+Q!G7+Q!H7)", "近 12 月 EPS＝Q 頁 51.52＋93.69＋72.03＋94.10（Derived；季 EPS 為第三方轉錄，合計與公司基本 EPS 有約 0.2% 差）"),
                    ("遠期 P/E：stockanalysis（共識，2026-11-13）", "=sa_fpe", "參考，不用於估值")):
    V.cell(fr, 1, nm); V.cell(fr, 2, f).number_format = "0.0"; V.cell(fr, 5, note); fr += 1
# 2. 十年前瞻 P/E（Bloomberg；取自本 repo tools/data/fwd_mult_hist.csv）
import csv as _csv
_FW = {}
with open(os.path.join(ROOT, "tools", "data", "fwd_mult_hist.csv"), encoding="utf-8") as _f:
    for _r in _csv.DictReader(_f):
        if _r["ticker"] == "8035.T" and _r["metric"] == "PE": _FW[int(_r["year"])] = (float(_r["value"]), _r["note"])
p0v = fr + 1
V.cell(p0v, 1, "2. 十年前瞻 P/E（Bloomberg BEST_PE_RATIO，未來 12 個月；各年度窗口＝前一年 11 月～當年 10 月月底值中位數）").font = BOLD
hdr(V, p0v + 1, ["年度窗口", "", "", "", "", "", "", "前瞻 P/E", "標記", "說明"])
y0 = p0v + 2
for i, y in enumerate(range(2017, 2027)):
    rr_ = y0 + i; v_, n_ = _FW[y]
    V.cell(rr_, 1, f"{y}（{y - 1}-11～{y}-10）"); c_ = V.cell(rr_, 8, v_); c_.font = BLUE; c_.number_format = "0.0"
    V.cell(rr_, 9, "Interested-party（Bloomberg 前瞻共識；本 repo 資料檔）"); V.cell(rr_, 10, n_ if i == 0 else ("2026 窗口含 2026-10-01 一拆五前後：P/E 為倍數，不受分割影響" if y == 2026 else ""))
y9 = y0 + 9; PE = f"H{y0}:H{y9}"
V.cell(y9 + 1, 1, "此處用前瞻（未來 12 個月）共識 P/E，不是年底實績 P/E；分位用於循環位置規則值（第 4–6 節，參考），不進目標價。第二來源：未找到。")
s0v = y9 + 3
V.cell(s0v, 1, "3. 分位（10 年）").font = BOLD
hdr(V, s0v + 1, ["統計", "前瞻 P/E", "用途"])
for i, (nm, p, use) in enumerate((("25 分位", 0.25, "見頂年"), ("中位數", 0.5, "中段"), ("75 分位", 0.75, "谷底")), s0v + 2):
    V.cell(i, 1, nm); V.cell(i, 2, f"=PERCENTILE({PE},{p})").number_format = "0.0"; V.cell(i, 3, use)
defname("pe_peak", f"Val!$B${s0v + 2}"); defname("pe_mid", f"Val!$B${s0v + 3}"); defname("pe_trough", f"Val!$B${s0v + 4}")
V.cell(s0v + 5, 1, "年數"); V[f"B{s0v + 5}"] = f"=COUNT({PE})"
T["pct"] = s0v
# 0.3 的分位檢查
V[f"B{q8}"], V[f"C{q8}"], V[f"D{q8}"] = [f"=IF(px/{c}{EPSR[2028]}<MIN({PE}),0,IF(px/{c}{EPSR[2028]}>MAX({PE}),1,PERCENTRANK({PE},px/{c}{EPSR[2028]})))" for c in "BCD"]
for c in "BCD": V[f"{c}{q8}"].number_format = "0%"
V[f"E{q8}"] = "現價 ÷ 各情境 FY2028 EPS 在十年前瞻 P/E（第 2 節）的分位；只作檢查，不進目標價"
# 4. 循環位置
c0 = s0v + 7
V.cell(c0, 1, "4. FY2027 的循環位置（設備股共用規則；見頂年度讀設備市場模型，不在本模型另訂）").font = BOLD
V.cell(c0 + 1, 1, CP.rule_text(PEAK))
hdr(V, c0 + 2, ["情境", "設備市場見頂年度", "FY2027 對照日曆年", "判定", "倍數"])
CYR = {}
for i, s in enumerate(SC, c0 + 3):
    V.cell(i, 1, s); V.cell(i, 2, f"=pk_{s}"); V.cell(i, 3, 2027).font = BLUE
    V.cell(i, 4, "=" + CP.excel_rule(f"C{i}", f"pk_{s}"))
    V.cell(i, 5, f'=IF(D{i}="見頂",pe_peak,IF(D{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"; CYR[s] = i
T["cyc"] = c0
# 5. 規則值隱含股價
v0 = c0 + 7
V.cell(v0, 1, "5. 規則值隱含股價（FY2027 EPS × 依判定的倍數；其餘倍數僅供參考）").font = BOLD
hdr(V, v0 + 1, ["情境", "EPS 年度", "EPS", "倍數類型", "倍數", "隱含股價", "相對現價", "備註"])
vr = v0 + 2; VAL = {}
for i, s in enumerate(SC):
    for y, f, nm in ((2027, f"E{CYR[s]}", None), (2027, "pe_peak", "參考：25 分位"), (2027, "pe_mid", "參考：中位數"), (2027, "pe_trough", "參考：75 分位"), (2028, "pe_mid", "參考：FY2028 × 中位數")):
        V.cell(vr, 1, s); V.cell(vr, 2, f"FY{y}"); V.cell(vr, 3, f"=PL!{FYC[y]}{SUMR[(s, 'eps')]}").number_format = "#,##0"
        V.cell(vr, 4, nm if nm else f"=\"規則值：\"&D{CYR[s]}"); V.cell(vr, 5, f"={f}").number_format = "0.0"
        V.cell(vr, 6, f"=C{vr}*E{vr}").number_format = "#,##0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
        if nm is None:
            for cc in range(1, 9): V.cell(vr, cc).fill = KEY
            VAL[s] = vr
        vr += 1
T["v0"] = v0
# 6. 舊法：規則值情境加權（參考）
o0 = vr + 1
V.cell(o0, 1, "6. 規則值情境加權（參考，不是目標價）").font = BOLD
hdr(V, o0 + 1, ["項目", "數值", "", "", "說明"])
V.cell(o0 + 2, 1, "規則值情境加權價（12 個月；參考）"); V.cell(o0 + 2, 2, f"=tp_wA*F{VAL['A']}+tp_wB*F{VAL['B']}+tp_wC*F{VAL['C']}").number_format = "#,##0"
V.cell(o0 + 3, 1, "相對現價（規則值情境加權價 ÷ 現價 − 1）"); V.cell(o0 + 3, 2, f"=B{o0 + 2}/px-1").number_format = "0.0%"
V.cell(o0 + 4, 1, "規則值方法"); V.cell(o0 + 4, 2, "各情境 FY2027 EPS × 十年前瞻 P/E 依循環位置取分位（前瞻倍數乘當年 EPS，口徑較寬鬆）；十年前瞻 P/E 區間 13–32x，2026 窗口 29.7x 為近十年高位")
T["old"] = o0
V.column_dimensions["A"].width = 52
for c in "BCDFGHI": V.column_dimensions[c].width = 14
V.column_dimensions["E"].width = 60; V.column_dimensions["J"].width = 60

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A）：對 FY2027、FY2028 營收與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方，與 PL 同一引擎）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收 ¥B", "FY2027 EPS ¥", "相對基準", "FY2028 營收 ¥B", "FY2028 EPS ¥", "相對基準", "FY2026 EPS ¥", "相對基準"])
cases = [("新設備超額成長 0（與市場同步；A 的低值）", "exc_A 低", dict(exc="exc_A_lo")), ("新設備超額成長 +4%/年", "exc_A 高", dict(exc="exc_A_hi")),
         ("新設備超額成長 −2%/年（情境 C 的值）", "exc＝exc_C", dict(exc="exc_C")),
         ("中國新設備 −20%（情境 C 的值）", "chh＝chh_C", dict(chh="chh_C")),
         ("中國占比 30%（搭配 chh_C）", "chn 低", dict(chn="chn_lo", chh="chh_C")), ("中國占比 40%（搭配 chh_C）", "chn 高", dict(chn="chn_hi", chh="chh_C")),
         ("FS 與其他年增 8%", "gfso_A 低", dict(gfso="gfso_A_lo")), ("FS 與其他年增 16%", "gfso_A 高", dict(gfso="gfso_A_hi")),
         ("CY2030 設備市場 −10%", "g_tail 低", dict(tail="g_tail_lo")), ("CY2030 設備市場 +10%", "g_tail 高", dict(tail="g_tail_hi")),
         ("邊際營業利益率 −3pt", "dm 低", dict(dm="dm_lo")), ("邊際營業利益率 +3pt", "dm 高", dict(dm="dm_hi")),
         ("淨利÷營業利益 0.70", "nioi 低", dict(nioi="nioi_lo")), ("淨利÷營業利益 0.78", "nioi 高", dict(nioi="nioi_hi")),
         ("股數年減 0%（無買回）", "bb 低", dict(bb="bb_lo")), ("股數年減 1.5%", "bb 高", dict(bb="bb_hi")),
         ("FY2026 實績低於指引 3%", "g_anc 低", dict(anc="g_anc_lo")), ("FY2026 實績高於指引 5%", "g_anc 高", dict(anc="g_anc_hi"))]
def put_case(i, name, setting, R, pre=""):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    SE[f"C{i}"] = f"={pre}{FYC[2027]}{R['rev']}"; SE[f"D{i}"] = f"={pre}{FYC[2027]}{R['eps']}"; SE[f"F{i}"] = f"={pre}{FYC[2028]}{R['rev']}"; SE[f"G{i}"] = f"={pre}{FYC[2028]}{R['eps']}"
    SE[f"E{i}"] = f"=D{i}/D$4-1"; SE[f"H{i}"] = f"=G{i}/G$4-1"
    SE[f"I{i}"] = f"={pre}{FYC[2026]}{R['eps']}"; SE[f"J{i}"] = f"=I{i}/I$4-1"
    for c in "CDFGI": SE[f"{c}{i}"].number_format = "#,##0"
    for c in "EHJ": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
SB = {}; base_row = 60
for n, (nm, setting, ov) in enumerate(cases):
    SB[nm] = engine(SE, base_row + n * 40, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    put_case(5 + n, nm, setting, SB[nm])
nn = 5 + len(cases)
put_case(nn, "情境 B", "設備市場 B；超額 0；中國 −5%；FS 與其他 +8%", PB["B"], "PL!")
put_case(nn + 1, "情境 C", "設備市場 C；超額 −2%；中國 −20%；FS 與其他 +3%；營益率 −1.5pt", PB["C"], "PL!")
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：FY2026 錨定公司指引，只受 g_anc 影響；驅動參數從 FY2027 起作用。FY2026 終端市場占比（mx26_*）只改變新設備拆分、不改變合計，故不列案例；g_tail 只影響 FY2029（CY2030 的 1/4）。")
# 目標價參數
s0 = base_row + len(cases) * 40 + 3
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
    i = s0 + 2 + k; w = kw.pop("w", "tp_w_dcf"); r_ = kw.get("r_", "tp_r")
    SE.cell(i, 1, nm); SE.cell(i, 2, st)
    SE[f"C{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_formula(c, **kw)}" for s, c in zip(SC, "BCD")) + f")*(1+{r_})"
    SE[f"D{i}"] = f"=Val!B{T['comptp']}"; SE[f"E{i}"] = f"={w}*C{i}+(1-{w})*D{i}"; SE[f"F{i}"] = f"=E{i}/E${s0 + 2}-1"
    for c in "CDE": SE[f"{c}{i}"].number_format = "#,##0"
    SE[f"F{i}"].number_format = "0.0%"
TS0 = s0

# ---------------- v0.3（2026-11-05）：標準版 DCF（無槓桿自由現金流 FCFF＋WACC；tools/fcff_dcf.py）取代股東現金流法 ----------------
# 依據：docs/plan/個股DCF標準規格.md v1.0、工作單「AI半導體 個股 DCF 改為標準版」（Andy 2026-11-05 20:42、20:45、21:21）。
# 做法：Val 0.1 節（原股東現金流）整段改寫為 FD 摘要（列數不變，0.2、0.3 節位置不動）；原 0.1 節的情境 EPS 表移到新 DCF 頁，
#   Comparable 的 12 個月後前瞻 EPS 與歷史倍數檢查改讀 DCF 頁（數值不變）；Sens 頁目標價參數段改為 FD 標準案例。
import re as _re3
def dcf_v03(T, paths, years, frac0, ver, a1, a2, MF, band_fmt, conv_from=None, nc_bs=None, nc_note="", detail_title=None, usd=False):
    DW = wb.create_sheet("DCF")
    DW["A1"] = a1; DW["A2"] = a2
    DW.cell(4, 1, "情境 EPS（FY2027–FY2032；v0.2 起在 Val 0.1 節，" + ver + " 移到本頁；Val 的 12 個月後前瞻 EPS 與歷史倍數檢查讀本表，數值不變）").font = BOLD
    hdr(DW, 5, ["項目", "A", "B", "C", "說明"])
    emap = {}
    for i, y in enumerate((2027, 2028, 2029, 2030, 2031, 2032)):
        rv = T[f"eps{y}"]; r = 6 + i; emap[rv] = r
        DW.cell(r, 1, f"  FY{y}")
        for c in "BCD":
            DW[f"{c}{r}"] = V[f"{c}{rv}"].value; DW[f"{c}{r}"].number_format = MF
    DW.cell(6, 5, "PL 頁延伸區塊；FY2030 起為收斂法")
    rx = _re3.compile(r"(?<![A-Za-z_!$\d])\$?([BCD])\$?(\d+)(?![\d(])")
    for k in ("feps", "pr"):
        for c in "BCD":
            v = V[f"{c}{T[k]}"].value
            if isinstance(v, str):
                V[f"{c}{T[k]}"].value = rx.sub(lambda m: f"DCF!${m.group(1)}${emap[int(m.group(2))]}" if int(m.group(2)) in emap else m.group(0), v)
    R = FD.write_detail(DW, 14, paths, years, frac0, conv_from=conv_from, fmt=MF, stub=True,
                        title=detail_title or "三情境 FCFF 明細（金額單位同 PL 頁；年度一欄；期中慣例）")
    R["end"] = FD.write_checks(DW, R["end"], nc_bs=nc_bs, nc_note=nc_note) + 1
    DW.column_dimensions["A"].width = 60; DW.column_dimensions["B"].width = 40
    for c in "CDEFGH": DW.column_dimensions[c].width = 13
    r1, r2 = T["h01"], T["dcftp"]
    assert str(V.cell(r1, 1).value).startswith("0.1 DCF"), V.cell(r1, 1).value
    _ns = FD.summary_rows(stub=True)   # 第二輪修正：模組 v2 stub 期在「＋淨現金與投資」後多兩列，用原 0.1 節之後的空白列
    for rr in range(r2 + 1, r1 + _ns):
        assert all(V.cell(rr, cc).value is None for cc in range(1, 8)), (rr, V.cell(rr, 1).value)
    assert V.cell(r1 + _ns, 1).value is None, V.cell(r1 + _ns, 1).value
    r2 = max(r2, r1 + _ns - 1)
    for rr in range(r1, r2 + 1):
        for cc in range(1, 8):
            V.cell(rr, cc).value = None; V.cell(rr, cc).fill = PatternFill(); V.cell(rr, cc).font = Font()
    S = FD.write_summary(wb, V, r1, R, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt=MF,
                         title=f"0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；{ver}；明細見 DCF 頁）", beta_note="自算產業 β 重新加槓桿")
    V[f"B{T['tpw']}"] = f"=tp_w_dcf*B{S['dcf']}+(1-tp_w_dcf)*B{T['comptp']}"
    V[f"B{T['band']}"] = f'=TEXT(D{S["v12"]},"{band_fmt}")&"～"&TEXT(B{S["v12"]},"{band_fmt}")'
    s0, s1 = T["sens0"], T["sens_last"]
    for rr in range(s0, s0 + 16):
        for cc in range(1, 10):
            if rr > s1: assert SE.cell(rr, cc).value is None, (rr, cc, SE.cell(rr, cc).value)
            SE.cell(rr, cc).value = None; SE.cell(rr, cc).fill = PatternFill(); SE.cell(rr, cc).font = Font()
    ts = FD.write_sens(SE, s0, R, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref=f"Val!$B${T['comptp']}", base_ref=f"Val!$B${T['tpw']}", fmt=MF,
                       title=f"目標價參數敏感度（{ver} 標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）", beta_du="beta_du", usd=usd)
    SE.column_dimensions["G"].width = 14
    T.update(dcf=S["dcf"], v12=S["v12"], S=S, R=R, ts=ts)
    for k in ("eps_h", "cf_h", "reinv", "dcfv", "dcf12", "dcftp") + tuple(f"eps{y}" for y in range(2027, 2033)) + tuple(f"cf{y}" for y in range(2027, 2033)):
        T.pop(k, None)
    T["dcftp"] = S["dcf"]; T["sens_last"] = ts[1]
    # 第二輪修正：模組 v2 的 In 頁 M2 標題不寫版號（原本此處改寫版號的段落刪除）
    return T
FY_X = list(range(2026, 2033))
def _ebit(s_, y):
    return f"PL!{FYC[y]}{PB[s_]['oi']}" if y <= 2029 else f"=PL!{EC[y]}{EXT[s_]['ni']}/nioi"
DPATHS = [dict(name=s_, label=f"情境 {s_}（{SCN[s_]}）", rev={y: f"PL!{EC[y]}{EXT[s_]['rev']}" for y in FY_X}, rev0=f"PL!{EC[2025]}{EXT[s_]['rev']}",
              ebit={y: _ebit(s_, y) for y in FY_X}, tax={y: "tax_w" for y in FY_X}, ebit_note="FY2030 起＝淨利 ÷ nioi（PL 延伸區塊只有淨利）") for s_ in SC]
T["h01"] = t0 + 1; T["sens0"] = TS0; T["sens_last"] = TS0 + 1 + len(tcases)
T = dcf_v03(T, DPATHS, FY_X, 0.25, "v0.1.0",
            "DCF — 標準版：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例；FY2026＝2026-04～2027-03，只計評價日後的 2027-01～03）；終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py",
            "EBIT：FY2026–FY2029 取 PL 推導區塊「營業利益」；FY2030–FY2032（PL 延伸區塊只有淨利）由淨利還原：EBIT＝淨利 ÷ nioi。稅率：公司指引隱含稅率 tax_w。"
            "業外收益不進 NOPAT，現金另在評價時加入。折舊、資本支出：近三年占營收比例，FY2027 起資本支出線性收斂到長期穩態 fc_cx_ss；營運資金：近五年營運資金增加 ÷ 營收增量 × 營收增量（存貨與預付款占比高）。",
            "#,##0", "#,##0", nc_bs="529.762+50",
            nc_note="2026-09-30 資產負債表（¥B；FY2027 Q2 決算短信 2026-10-30）：現金及存款 529.762＋有價證券 50.000；無有息負債；非支配股東持分約 0",
            detail_title="三情境 FCFF 明細（¥B；FY2026＝2026-04～2027-03，評價日後只計 2027-01～03）", usd=True)
TS0 = T["sens0"]
SE.column_dimensions["A"].width = 40; SE.column_dimensions["B"].width = 40
for c in "CDEFGHIJ": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — FY2026 錨定公司指引（A 節為恆等式，不是檢驗）；B 節共識只作參考，未用於校準或錨點；C 節為結構檢查。差距＝模型 ÷ 對照 − 1"
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
RC.cell(rr_, 1, "A. FY2026 錨點（公司指引 2026-10-30）：差距為 0 是恆等式").font = BOLD; rr_ += 1
rec("FY2026 營收 vs 公司指引", "=g26_rev", pl("rev"), "FY2027 Q2 決算短信（2026-10-30）3,700.0（+51.4%）")
rec("FY2026 營業利益 vs 公司指引", "=g26_oi", pl("oi"), "1,169.0（+87.1%）")
rec("FY2026 EPS vs 公司指引", "=g26_eps", pl("eps"), "381.69", "差異為股數捨入", fmt="#,##0.00")
rec("FY2026 營業利益率 vs FY2025 實績（百分點）", "=a25_oi/a25_rev", pl("oim"), "FY2025 25.6%", "公司稱價格最適化與組合改善（Interested-party）", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "A2. 上半實績與下半隱含（只作合理性檢查）").font = BOLD; rr_ += 1
rec("FY2026 上半營收實績 vs 7 月上半預想（1,620.0）", "=1620", lambda s: "=Q!G4+Q!H4", "公司 7 月發布的上半預想（FY2027 Q1 決算說明；上半預想 營收 1,620／營業利益 458）", "上半實績比預想高 2.9%（營收）、10.6%（營業利益）")
rec("FY2026 下半營收（指引隱含）vs FY2025 下半實績", "=Q!D4+Q!E4", lambda s: f"=PL!D{SUMR[(s, 'rev')]}-(Q!G4+Q!H4)", "FY2025 下半 Q3＋Q4（Q 頁）", "下半隱含 2,032.5 vs 1,263.9：+60.8%；公司稱客戶要求提前交貨、下半新設備 1,560 預想")
rec("FY2026 下半營業利益（指引隱含）vs FY2025 下半實績", "=Q!D5+Q!E5", lambda s: f"=PL!D{SUMR[(s, 'oi')]}-(Q!G5+Q!H5)", "FY2025 下半（Q 頁）", "下半隱含 662.2 vs 321.8")
RC.cell(rr_, 1, "B. FY2026–FY2027 vs 共識（參考）").font = BOLD; rr_ += 1
rec("FY2026 營收 vs 共識（Yahoo 快照 2026-11-17）", "=cons_rev26", pl("rev"), "Yahoo Finance 快照：3,552（FY3/27）；公司 10-30 指引 3,700 高於共識 4.2%", "共識可能含指引前的估計；第二來源：未找到")
rec("FY2026 EPS vs 共識", "=cons_eps26", pl("eps"), "Yahoo 快照：375.6；公司指引 381.69", "", fmt="#,##0")
rec("FY2027 營收 vs 共識", "=cons_rev27", pl("rev", 2027), "Yahoo 快照：FY3/28 4,487", "共識 FY2027 年增 26%：設備市場 CY2027 公司說法 >210B（+31%）")
rec("FY2027 EPS vs 共識", "=cons_eps27", pl("eps", 2027), "Yahoo 快照：FY3/28 517.6", "", fmt="#,##0")
RC.cell(rr_, 1, "C. 結構檢查（設備市場讀 markets/equipment；公司說法為 Interested-party）").font = BOLD; rr_ += 1
rec("設備市場模型 CY2026 五分項合計 vs 公司說法 WFE（約 160B）", "=co_wfe26", lambda s: f"=tot_{s}_26", "公司 FY2027 Q2 說明會 Q&A（2026-10-30）", "markets/equipment 為 2026-10-23 版，早於公司 10-30 上修；SEMI 年中預測見下列", fmt="#,##0.0")
rec("設備市場模型 CY2026 五分項合計 vs SEMI WFE 預測", "=semi_wfe26", lambda s: f"=tot_{s}_26", "SEMI 2026-07 年中預測（設備市場模型 Mkt 頁）", "", fmt="#,##0.0")
rec("設備市場模型 CY2027 五分項合計 vs 公司說法（>210B）", "=co_wfe27", lambda s: f"=tot_{s}_27", "公司 Q&A；CY2027 為客戶資本支出計畫加總", "公司稱另有上行空間；模型情境 B、C 低於公司數", fmt="#,##0.0")
rec("DRAM 設備 CY2026（模型 dr）vs 公司說法（約 48B）", "=co_dr26", lambda s: f"=dr{s}_26", "公司 Q&A：high-$40B", "", fmt="#,##0.0")
rec("NAND 設備 CY2026（模型 nd）vs 公司說法（約 15B）", "=co_nd26", lambda s: f"=nd{s}_26", "公司 Q&A：mid-$10B", "", fmt="#,##0.0")
rec("FY2026 新設備年增 vs 設備市場 CY2026 年增", f"=tot_A_26/tot_A_25-1", lambda s: f"=PL!D{PB[s]['ne']}/PL!C{PB[s]['ne']}-1", "markets/equipment（情境 A；2026 三情境相同）", "公司指引新設備 +59%（FY2026，含下半預想）；CY2026 設備市場年增較低，差距是出貨與認列時點（Assumed 之外的公司指引）", fmt="0.0%", pt=True)
rec("FY2028 新設備 ÷ FY2026（模型）vs 設備總額 CY2028 ÷ CY2026", "=0", lambda s: f"=PL!F{PB[s]['ne']}/PL!D{PB[s]['ne']}-1-(tot_{s}_28/tot_{s}_26-1)", "markets/equipment 設備總額（同字母）", "差距欄＝模型 − 設備市場；超額成長與中國係數造成的差", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "參考：中國占營收（公司，FY2026 第二季 36.4%、上半約 33.7%）"); RC.cell(rr_, 3, "=chn").number_format = "0%"; RC.cell(rr_, 10, "以全社營收占比代替新設備占比；Sens 列 chn 30%／40%"); rr_ += 1
RC.column_dimensions["A"].width = 66
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 40; RC.column_dimensions["J"].width = 70

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文以網頁擷取工具讀取，為轉錄；數字以算術互核；版面與頁碼待人工複核）"; SO["A1"].font = BOLD
srcs = [("公司：FY2027（2027年3月期）第2四半期 決算短信〔日本基準〕（連結）（2026-10-30）", "https://www.tel.com/ir/irta3a00000006g5-att/fy27q2tanshin-e.pdf"),
 ("公司：FY2027 Q2 決算說明資料（2026-10-30）", "https://www.tel.com/ir/irta3a00000006g5-att/fy27q2presentations-e2.pdf"),
 ("公司：FY2027 Q2 決算說明會 Q&A（2026-10-30）", "https://www.tel.com/ir/irta3a00000006g5-att/FY27Q2_EarningCall_QA_E.pdf"),
 ("公司：FY2027 Q2 決算說明會逐字稿（2026-10-30）", "https://www.tel.com/ir/irta3a00000006g5-att/fy27q2transcript-e.pdf"),
 ("stockanalysis.com：TYO:8035 年度與季度財報（2026-11-17；第三方轉錄）", "https://stockanalysis.com/quote/tyo/8035/financials/"),
 ("finboard.net：JP:8035 損益表（FY2012–FY2025；2026-11-17；第三方轉錄）", "https://finboard.net/companies/JP:8035/income-statement"),
 ("Yahoo Finance 基本面快照（2026-11-17；共識、目標價）", "docs/workorders/20261117_新增五家_Yahoo基本面快照.md"),
 ("Yahoo Finance 日線收盤 2026-11-13（現價）", "markets/invest_map/data/closes_20261020_20261113.csv"),
 ("Bloomberg 前瞻 P/E 十年歷史（8035.T 與同業）", "tools/data/fwd_mult_hist.csv"),
 ("本 repo 現行模型 Val 同業表（amat、lrcx、klac、asml、accretech）：同業前瞻 P/E 換算到 2026-11-13", "各模型 dist xlsx"),
 ("日本財務省：國債金利情報（10 年 2.817%，2026-11-02）", "https://www.mof.go.jp/english/policy/jgbs/reference/interest_rate/jgbcme.csv"),
 ("Damodaran：國家風險溢酬（日本 CRP 0.91%，2026-01-05）", "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html"),
 ("市場層：markets/equipment 成品（本 repo）；tools/cycle_position.py", ""),
 ("未找到：TEL 各產品市占（塗佈顯影以外）；十年年底收盤價的第二來源；共識 FY2028 EPS；RSU 與選擇權明細；有息負債明細；FY2015–FY2021 的第二來源", ""),
 ("無法開啟（建置環境被擋或逾時）：tel.com PDF 直連（HTTP 403，改以網頁擷取工具）、roic.ai、finboard 營收端點、itiger（數字與公司原文不符，棄用）", ""),
 ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 80; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
PROP = ("東京威力科創是塗佈顯影設備的全球領導者（市占約九成），並在蝕刻、沉積、清洗與探針台有完整產品線；FY2026 公司指引營收 3,700 十億日圓（+51%），"
        "新設備銷售由 DRAM（含 HBM）與先進邏輯拉動。問題是前段設備循環能撐多久，以及一拆五後股價 9,461 日圓已經反映多少。")
lines = [("東京威力科創（8035.T）收支與估值模型 v0.1.0（AI 半導體 Project：設備與材料，前段設備）　2026-11-17", True),
 ("模型命題：" + PROP, False),
 ("（價格口徑：2026-10-01 一拆五之後；現價 ¥9,461＝2026-11-13 收盤；FY2026 公司指引 EPS ¥381.69；現價 ÷ 指引 EPS ＝ 24.8 倍；見 Val 頁第 1 節）", False), ("", False),
 ("驅動因子 → 推導量（因果方向：左邊決定右邊）", True),
 ("PL 區塊｜markets/equipment 同字母情境的 DRAM、NAND、先進邏輯、其他代工邏輯、其他前段設備（CY）→ FY(t)＝3/4 × CY(t)＋1/4 × CY(t+1) → 各終端市場需求指數（FY2026＝1）× FY2026 該終端市場新設備銷售 → 新設備銷售（DRAM／NAND／非記憶體）", False),
 ("PL 區塊｜公司特有：超額成長（A +2%／B 0／C −2% 每年）、中國變化（A 0／B −5%／C −20%，× 中國占比 34%）→ 新設備銷售", False),
 ("PL 區塊｜Field Solutions 與其他（零件、服務、二手與改裝設備）：FY2026 指引隱含值 ×（1＋年增；A 12%／B 8%／C 3%，Analogy：AMAT AGS、Lam CSBG）", False),
 ("Hist → PL｜營業利益＝FY2026 錨點 ＋ 邊際營業利益率（FY2025→FY2026 兩點推導，約 43%）× 營收變化；淨利＝營業利益 × 0.74（公司指引隱含）；EPS＝淨利 ÷ 股數（每年淨買回 0.8%）", False),
 ("PL｜FY2026 錨點＝公司指引（2026-10-30 首次發布通期預想：營收 3,700、營業利益 1,169、淨利 865、EPS 381.69；上半實績 1,667.5／506.8／376.9，下半＝通期 − 上半）×（1＋g_anc，基準 0）", False),
 ("Val｜目標價＝0.5 × DCF（標準版：無槓桿自由現金流 FCFF＋WACC，FY2026–32 明確期＋永續；DCF 頁）＋ 0.5 × Comparable（前段設備同業前瞻 P/E 中位數 × 12 個月後前瞻 EPS，再依相對倍數調整）；規則值（十年前瞻 P/E 分位 × FY2027 EPS）為參考", False),
 ("會計年度：4 月到隔年 3 月（本模型 FY2025＝2025-04～2026-03，公司稱 FY2026；本模型 FY2026＝公司 FY2027）。金額為日圓，未建匯率（公司前提 1 美元＝約 150 日圓；美元計價銷貨占比高，匯率影響未建模，列入限制）。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（PL 頁底部延伸區塊，灰底）。情境：A＝樂觀、B＝基準、C＝保守（設備市場讀 markets/equipment 同字母情境）。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("可識別性：TEL 只有單一應報導部門；新設備的終端市場占比（DRAM／NAND／非記憶體）取自公司說明會圖表（欄位對應為推斷）；塗佈顯影市占約 91% 為公司說法，其他產品市占未找到；超額成長與中國變數為 Assumed，列在 Sens。", False),
 ("已知限制：見 tel/README.md「待判斷」與「建議補進 Siliconomics」。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 180

# 設備總額具名範圍（Recon 用；五分項合計）
r += 1
IN.cell(r, 1, "K. 設備總額（markets/equipment 成品 Mkt 頁五分項合計，同字母；只用於 Recon 結構檢查）").font = BOLD; r += 1
for s in SC:
    for j, y in enumerate((25, 26, 27, 28, 29)):
        IN.cell(r, 1, f"tot_{s}_{y}"); IN.cell(r, 2, f"設備總額 CY20{y}（情境 {s}；五分項合計）"); IN.cell(r, 3, "US$B")
        IN.cell(r, 5, "=" + "+".join(f"{k}{s}_{y}" for k in MK)); IN.cell(r, 7, "Derived"); IN.cell(r, 8, "In 頁 C 節五分項合計" if j == 0 else "同上")
        defname(f"tot_{s}_{y}", f"In!$E${r}"); r += 1

# ---------------- 上檔空間級距、投資建議（Val 頁最後新增，依「上檔空間：加權目標價」固定規則）----------------
_U = f"B{T['upw']}"
_G = V.max_row + 2
V.cell(_G, 1, "上檔空間級距（六級，供分組）")
V[f"B{_G}"] = f'=IF({_U}>0.5,"大於 50%",IF({_U}>0.25,"25%–50%",IF({_U}>=0,"0%–25%",IF({_U}>=-0.25,"0%～−25%",IF({_U}>=-0.5,"−25%～−50%","小於 −50%")))))'
V[f"B{_G}"].number_format = "@"; V[f"E{_G}"] = "以「上檔空間：加權目標價」分六級；六級只供分組比較；投資建議依下一列的五級固定規則"
V.cell(_G + 1, 1, "投資建議")
V[f"B{_G + 1}"] = f'=IF({_U}>0.5,"強力買進",IF({_U}>0.25,"買進",IF({_U}>=-0.25,"中立",IF({_U}>=-0.5,"放空","強力放空"))))'
V[f"B{_G + 1}"].number_format = "@"
V[f"E{_G + 1}"] = "依上檔空間固定規則：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空"
if os.environ.get("ROWS"):
    print("T", T, "SUMR", SUMR, "EXT", EXT, "PB", {s: PB[s] for s in SC}, "CHK26", CHK26, "TS0", TS0, "SENS_LAST", SENS_LAST)

# ---------------- 分部原則第 3 條：核對列「檢查：分部合計 − 合併營收」（追加於工作表最後）----------------
import os as _os2, sys as _sys2
_sys2.path.insert(0, _os2.path.join(ROOT, "tools", "segment_check"))
from addrow import add_check
add_check(Q, seg_rows=[8, 9], total_row=4, cols="BCDEFGH", rnd=6, note="分部原則第 3 條：TEL 為單一應報導部門；營收構成（新設備＋Field Solutions 與其他）合計對合併營收；分部宣告見 segments.json")

order = ["README", "In", "Q", "Hist", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
sys.path.insert(0, os.path.join(ROOT, "tools"))
import relative_multiple as RMX  # noqa: E402

# ---------------- Wattonomics 錨定（三層校準原則；Siliconomics）----------------
from openpyxl.workbook.defined_name import DefinedName as _DN  # noqa: E402
from openpyxl.styles import Font as _F  # noqa: E402
def _wta_in(wb):
    ws = wb["In"]; r = ws.max_row + 1
    ws.cell(r, 1, f"Z. Wattonomics 錨定係數（敏感度用；Siliconomics {SIL.VERSION}）").font = _F(bold=True); r += 1
    rows_ = []; SIL.wta_rows(lambda *a: rows_.append(a))
    for a in rows_:
        for j, v in enumerate(a, 1):
            if v is not None: ws.cell(r, j, v)
        ws.cell(r, 5).font = _F(color="0000FF")
        wb.defined_names[a[0]] = _DN(a[0], attr_text=f"In!$E${r}")
        r += 1
_wta_in(wb)
WTA_READ = "讀法：Wattonomics 錨定：以情境 B 為底，產業量 CY2027 起 × wta（Siliconomics SIF_WTAnchor_Base＝R4 比值）；FY2026 維持錨點。三層校準原則；個股基準暫不改，列在 Sens 與 Coord。"
def _wta_f(SE, w0, CM, desc):
    SE.cell(w0, 1, "Wattonomics 錨定（三層校準原則）").font = BOLD
    SE.cell(w0 + 1, 1, WTA_READ)
    SE.cell(w0 + 2, 1, desc)
    hdr(SE, w0 + 3, ["日曆年", "單位"] + [f"CY{y}" for y in range(2025, 2030)] + ["", "來源"])
    fr = w0 + 4
    SE.cell(fr, 1, "Wattonomics 錨定係數（In 頁 wta_27／28／29）"); SE.cell(fr, 2, "x")
    for y, f in zip(range(2025, 2030), (1, 1, "=wta_27", "=wta_28", "=wta_29")): SE[f"{CM[y]}{fr}"] = f; SE[f"{CM[y]}{fr}"].number_format = "0.000"
    return fr
WTA_AIWFE = ("產業量＝設備市場模型（markets/equipment，情境 B）中標〔AI〕的分項：先進邏輯〔AI〕、DRAM（含 HBM）〔AI〕；CY2027 起 × wta，其他分項（其他代工邏輯、其他前段、NAND）不變。"
             "注意：設備市場由公司資本支出推得，wta 是算力側（Wattonomics）對 AI 設備量的另一讀數。")
fr = _wta_f(SE, SE.max_row + 3, CYC, WTA_AIWFE + " 本模型：先進邏輯與 DRAM 兩個 AI 分項 × wta；FY2026 指數分母維持情境 B（錨點）。")
_ar = {}
for i, k in enumerate(("la", "dr"), 1):
    vals, src_ = SEG[(k, "B")]; SE.cell(fr + i, 1, f"{SEG_LAB[k]}（情境 B）"); SE.cell(fr + i, 2, "US$B"); SE.cell(fr + i, 10, src_ + "（唯讀抄錄）"); _ar[k] = fr + i
    for j, y in enumerate(range(2025, 2030)): c_ = SE[f"{CYC[y]}{fr + i}"]; c_.value = float(vals[j]); c_.font = BLUE; c_.number_format = "0.00"
ovr = {}
for y in (2027, 2028, 2029):
    for k in ("la", "dr"):
        ovr[(k, y)] = f"={k}B_{y % 100}*${CYC[y]}${fr}"
RW = engine(SE, fr + 5, "案例區塊：Wattonomics 錨定（情境 B 為底；設備市場 AI 分項 CY2027 起 × wta）", scn="B", ovr=ovr,
            den26={"la": "(0.75*laB_26+0.25*laB_27)", "dr": "(0.75*drB_26+0.25*drB_27)"})
put_case(nn + 2, "Wattonomics 錨定（情境 B 為底）", "設備市場 AI 分項 CY2027 起 × wta；相對欄對情境 B 列", RW)
for a_, b_ in (("E", "D"), ("H", "G"), ("J", "I")): SE[f"{a_}{nn + 2}"] = f"={b_}{nn + 2}/{b_}{nn}-1"

# 相對倍數調整（tools/relative_multiple.py）：tel 尚未列入 target_summary 的 universe／peers（工作單禁止改 tools/），此處在本模型內補上
_u0, _p0 = RMX._universe, RMX._peers
_TEL_PEERS = [(p[0], p[1]) for p in PEERS]
RMX._universe = lambda: {**_u0(), "tel": {"folder": "tel", "公司": "東京威力科創（Tokyo Electron）", "own_ticker": "8035.T", "代號": "8035"}}
RMX._peers = lambda folder: _TEL_PEERS if folder == "tel" else _p0(folder)
RMX.apply(wb, "tel")
def _incentive(txt):
    if "SEMI" in txt: return "設備商協會預測（會員為設備商）：傾向樂觀；只用於 Recon 對帳"
    if any(k in txt for k in ("Yahoo", "共識", "分析師", "stockanalysis", "目標價")) and "公司" not in txt[:6]:
        return "賣方分析師共識（資料商轉述）：賣方與承銷、研究業務有利益關係，傾向偏樂觀；只作對照，不進估值（除前瞻 EPS 對照）"
    if "公司" in txt: return "發行公司自揭（指引、資本支出計畫、市場說法、KPI）：傾向呈現正面；以事後決算檢驗"
    return "第三方或市場資料商轉述；立場見來源欄"
for _r in range(4, IN.max_row + 1):
    if IN.cell(_r, 7).value == "Interested-party" and not IN.cell(_r, 10).value:
        IN.cell(_r, 10, _incentive(str(IN.cell(_r, 8).value or "")))
IN.column_dimensions["J"].width = 60
sys.path.insert(0, os.path.join(ROOT, "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.1.0", "2026-11-17")
_RD = wb["README"]
# 新列放在 README 既有內容之後，不位移既有列；README 以外的工作表不動
def _rw(ws, text, exact=False, col=1, start=1):
    for r_ in range(start, ws.max_row + 1):
        v_ = ws.cell(r_, col).value
        if isinstance(v_, str) and ((v_.strip() == text) if exact else v_.strip().startswith(text)):
            return r_
    raise KeyError((ws.title, text))
def _F(*parts):   # 文字公式：字串常數每段 ≤ 120 字（Excel 公式內字串常數上限 255 字元）；("=", 運算式) 為公式片段
    out = []
    for p_ in parts:
        if isinstance(p_, tuple):
            out.append(p_[1])
        else:
            s_ = str(p_)
            out += ['"' + s_[i_:i_ + 120].replace('"', '""') + '"' for i_ in range(0, len(s_), 120)]
    return "=" + "&".join(out)
_E = lambda e: ("=", e)
_TX = lambda c, f: _E(f'TEXT({c},"{f}")')
_SG = lambda c: _TX(c, "+0.0%;-0.0%")
_PC = lambda c: _TX(c, "0.0%")
_UP = lambda c: _E(f'IF({c}>=0,"高 ","低 ")&TEXT(ABS({c}),"0.0%")')
_ws_in = wb["In"]
_IN = {_ws_in.cell(r_, 1).value: r_ for r_ in range(4, _ws_in.max_row + 1) if _ws_in.cell(r_, 1).value}
I_ = lambda k: f"In!E{_IN[k]}"
I_lo = lambda k: f"In!D{_IN[k]}"
I_hi = lambda k: f"In!F{_IN[k]}"
_V, _RC, _SE, _RD = wb["Val"], wb["Recon"], wb["Sens"], wb["README"]
vw, vd, vc = _rw(_V, "目標價（DCF 與 Comparable 加權）", exact=True), _rw(_V, "DCF 目標價（情境加權）"), _rw(_V, "Comparable 目標價", exact=True)
vca, vk, vm = _rw(_V, "Comparable 目標價（相對倍數調整後）"), _rw(_V, "相對倍數調整係數"), _rw(_V, "同業前瞻 P/E 中位數")
vup, vuc, vcons, vrec = _rw(_V, "上檔空間：加權目標價"), _rw(_V, "上檔空間：Comparable"), _rw(_V, "市場共識目標價"), _rw(_V, "投資建議", exact=True)
vpx = _rw(_V, "現價（")
vfe = _rw(_V, "12 個月後的前瞻 EPS")
vd12 = _rw(_V, "DCF 情境價值推到 12 個月後")
e0 = _rw(wb["DCF"], "情境 EPS（FY2027") + 1   # 情境 EPS 表在 DCF 頁
E27, E28, E29, E32 = e0 + 1, e0 + 2, e0 + 3, e0 + 6
W_ = I_("tp_w_dcf")
# 加權目標價相對現價的差距拆解（補充規則 B）：DCF 貢獻＝w ×（DCF − 現價），Comparable 貢獻＝(1 − w) ×（調整後 Comparable − 現價）
SH_DCF = f"{W_}*(Val!B{vd}-Val!B{vpx})/(Val!B{vw}-Val!B{vpx})"
T2 = lambda c: _TX(c, "#,##0")
T0 = lambda c: _TX(c, "#,##0")
CT_DCF = f"{W_}*(Val!B{vd}-Val!B{vpx})/Val!B{vpx}"; CT_CMP = f"(1-{W_})*(Val!B{vca}-Val!B{vpx})/Val!B{vpx}"
r_c27 = _rw(_RC, "FY2027 營收 vs 共識")
s_chh, s_chn50, s_exc0, s_b, s_c = _rw(_SE, "中國新設備 −20%"), _rw(_SE, "中國占比 40%"), _rw(_SE, "新設備超額成長 0"), _rw(_SE, "情境 B", exact=True), _rw(_SE, "情境 C", exact=True)
s_w8 = _rw(_SE, "DCF 權重 0.3／0.7"); s_wacc = _rw(_SE, "WACC ±1pt"); s_beta = _rw(_SE, "β 低／高"); s_wc = _rw(_SE, "營運資金 ÷ 營收 低／高")
v_tv, v_mx = _rw(_V, "其中：終值現值占 EV"), _rw(_V, "隱含終值 EV／EBITDA")
# ---------------- README 展示層（個股 HTML 用；公式帶數字）----------------
lead = _F("命題：東京威力科創是塗佈顯影設備的全球領導者（市占約九成），並有蝕刻、沉積、清洗與探針台產品線；FY2026 公司指引營收 ", T0("g26_rev"),
          " 十億日圓。新設備銷售中 DRAM 占比約 ", _TX("mx26_h1_dr", "0%"), "（上半實績）。問題是前段設備循環能撐多久，以及現價是否已經反映。三個問題如下。")
props = [
    "命題 1：設備循環下盈餘撐得住嗎——2028 年設備市場持平或下滑時（基準、保守情境），FY2028 每股盈餘是否仍高於 FY2027。",
    "命題 2：中國與超額成長的風險有多大——中國新設備營收下滑兩成、或新設備只跟著設備市場成長（無超額成長）時，每股盈餘少多少。",
    "命題 3：每股值多少——以 DCF 與同業倍數加權的目標價，相對現價的上檔空間是多少。",
]
_dom = _E(f'IF(ABS({CT_DCF})>ABS({CT_CMP}),"主要來自 DCF（",  "主要來自相對倍數調整後的 Comparable（")')
obs = [
    _F("重點 1：加權目標價", _E(f'IF(Val!B{vup}>=0,"高於","低於")'), "現價，投資建議", _E(f"Val!B{vrec}"), "：加權目標價約 ", T0(f"Val!B{vw}"), " 日圓，比現價 ", T0(f"Val!B{vpx}"), " ", _UP(f"Val!B{vup}"),
       "。拆解：DCF（", T0(f"Val!B{vd}"), "，權重 ", _TX(W_, "0%"), "）", _E(f'IF({CT_DCF}>=0,"拉高 ","拉低 ")&TEXT(ABS({CT_DCF}),"0.0%")'), "，相對倍數調整後的 Comparable（", T0(f"Val!B{vca}"), "）",
       _E(f'IF({CT_CMP}>=0,"拉高 ","拉低 ")&TEXT(ABS({CT_CMP}),"0.0%")'), "；差距", _dom, _E(f'IF(ABS({CT_DCF})>ABS({CT_CMP}),TEXT(ABS({CT_DCF}),"0.0%"),TEXT(ABS({CT_CMP}),"0.0%"))'), "）。"),
    _F("重點 2：設備市場情境決定 FY2028：FY2027 三情境每股盈餘樂觀／基準／保守 ", T2(f"DCF!B{E27}"), "／", T2(f"DCF!C{E27}"), "／", T2(f"DCF!D{E27}"),
       " 日圓，FY2028 為 ", T2(f"DCF!B{E28}"), "／", T2(f"DCF!C{E28}"), "／", T2(f"DCF!D{E28}"), "；共識 FY2027 EPS ", T0("cons_eps27"), "（基準情境 ", _SG(f"Recon!F{r_c27 + 1}"), "）。"),
    _F("重點 3：單一輸入中最敏感的是中國與超額成長（以樂觀情境為基準，Sens 頁）：中國占比 40%（搭配中國 −20%）使 FY2028 每股盈餘 ", _SG(f"Sens!H{s_chn50}"),
       "；新設備無超額成長（與設備市場同步）使 ", _SG(f"Sens!H{s_exc0}"), "；改用基準／保守情境使 ", _SG(f"Sens!H{s_b}"), "／", _SG(f"Sens!H{s_c}"), "。"),
    _F("重點 4：DCF 為標準版（無槓桿自由現金流折現）：推到 12 個月後的 DCF 情境價值樂觀／基準／保守為 ", T0(f"Val!B{vd12}"), "／", T0(f"Val!C{vd12}"), "／", T0(f"Val!D{vd12}"),
       "，DCF 目標價 ", T0(f"Val!B{vd}"), "。WACC ", _TX("wacc", "0.0%"), "（股權成本 ", _TX("ke", "0.0%"), "＝rf ", _TX("rf", "0.00%"),
       " ＋ β ", _TX("beta", "0.00"), " ×（ERP ", _TX("erp", "0%"), " ＋ 日本國家風險溢酬 ", _TX("crp", "0.00%"), "）；β 為自算產業 β 重新加槓桿）。公司無有息負債，WACC＝股權成本。"
       "營運資金占營收增量約 ", _TX("fc_wc", "0%"), "，營收成長要先墊營運資金。終值現值占基準情境企業價值 ", _TX(f"Val!C{v_tv}", "0%"), "、隱含終值 EV／EBITDA ",
       _TX(f"Val!C{v_mx}", "0.0"), " 倍。以情境加權目標價為基準，β 低／高時加權目標價 ", _SG(f"Sens!D{s_beta}"), "／", _SG(f"Sens!F{s_beta}"), "，WACC −1pt／＋1pt 時 ", _SG(f"Sens!D{s_wacc}"), "／", _SG(f"Sens!F{s_wacc}"),
       "；「", _E(f"Sens!A{s_w8}"), "」時 ", _SG(f"Sens!D{s_w8}"), "／", _SG(f"Sens!F{s_w8}"), "。"),
]
answers = [
    _F("設備循環下盈餘撐得住嗎？答：要看設備市場。FY2028 相對 FY2027：樂觀情境 ", _SG(f"DCF!B{E28}/DCF!B{E27}-1"), "、基準情境 ", _SG(f"DCF!C{E28}/DCF!C{E27}-1"), "、保守情境 ", _SG(f"DCF!D{E28}/DCF!D{E27}-1"),
       "（每股盈餘 ", T2(f"DCF!B{E28}"), "／", T2(f"DCF!C{E28}"), "／", T2(f"DCF!D{E28}"), " 日圓）。情境的設備市場讀 markets/equipment，超額成長與中國變數為假設。"),
    _F("中國與超額成長的風險有多大？答：以樂觀情境為基準（Sens 頁），中國新設備 −20% 使 FY2027／FY2028 每股盈餘 ", _SG(f"Sens!E{s_chh}"), "／", _SG(f"Sens!H{s_chh}"),
       "；中國占比提高到 40% 時 ", _SG(f"Sens!E{s_chn50}"), "／", _SG(f"Sens!H{s_chn50}"), "；無超額成長使 ", _SG(f"Sens!E{s_exc0}"), "／", _SG(f"Sens!H{s_exc0}"), "。下滑比例與超額成長是敏感度假設，不是預測。"),
    _F("每股值多少？答：加權目標價約 ", T0(f"Val!B{vw}"), " 日圓，比現價 ", T0(f"Val!B{vpx}"), " ", _UP(f"Val!B{vup}"), "，投資建議", _E(f"Val!B{vrec}"),
       "。加權目標價＝DCF ", T0(f"Val!B{vd}"), " 與 Comparable 各占 ", _TX(W_, "0%"), "／", _TX(f"1-{W_}", "0%"), "；進加權的 Comparable 是相對倍數調整後的 ", T0(f"Val!B{vca}"),
       "（同業前瞻本益比中位數 ", _TX(f"Val!B{vm}", "0.0"), " 倍 × 相對倍數調整係數 ", _TX(f"Val!B{vk}", "0.00"), " × 12 個月後前瞻每股盈餘 ", T0(f"Val!B{vfe}"), "）。只看未調整的 Comparable 目標價為 ", T0(f"Val!B{vc}"),
       "（上檔空間 ", _SG(f"Val!B{vuc}"), "）。市場共識目標價 ", T0(f"Val!B{vcons}"), "。"),
]
# 第二輪修正（2026-11-05）：展示層重點——淨現金滾到評價日（stub）、美元口徑折現（非美國公司）、終值占比 > 80% 標示（公式判斷；Sens 以加權目標價為基準）
def _r2pt(n, usd, ccy_fmt="#,##0.00", amt_fmt="#,##0.0", unit=""):
    _Vw, _Sw = wb["Val"], wb["Sens"]
    def _f(ws, lab, start=1):
        return next(r for r in range(start, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(lab))
    h = _f(_Vw, "0.1 DCF"); nc, st, ncr, tv = (_f(_Vw, l, h) for l in ("＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "其中：終值現值占 EV"))
    t = lambda c, f: f'TEXT({c},"{f}")'
    s = (f'="重點 {n}：淨現金由資產負債表日 "&bs_date&" 滾到評價日 2026-12-31：資產負債表淨現金與投資 "&{t(f"Val!C{nc}", amt_fmt)}&" ＋ stub 期 FCFF（情境 B 第一年 FCFF × "&TEXT(stub_frac,"0.000")&" 年）"&{t(f"Val!C{st}", amt_fmt)}'
         f'&" − 已宣告股利 "&{t("stub_div", amt_fmt)}&" ＝ "&{t(f"Val!C{ncr}", amt_fmt)}&"{unit}。"')
    if usd:
        u = _f(_Sw, "美元口徑折現")
        s += (f'&"美元口徑折現（Sens 案例：rf 改用美國 10 年期 "&TEXT(rf_us,"0.00%")&"、g＝MIN("&TEXT(g_cap,"0%")&", "&TEXT(rf_us,"0.00%")&")＝"&TEXT(MIN(g_cap,rf_us),"0.0%")&"，CRP、匯率與現金流不變）時，加權目標價 "'
              f'&{t(f"Sens!C{u}", ccy_fmt)}&"，相對基準加權目標價 "&TEXT(Sens!D{u},"+0.0%;-0.0%")&"；基準採當地 10 年期利率 "&TEXT(rf,"0.00%")&"＋固定匯率。"')
    s += (f'&IF(MAX(Val!B{tv}:D{tv})>0.8,"終值現值占 EV 超過 80% 的情境："&IF(Val!B{tv}>0.8,"A "&TEXT(Val!B{tv},"0%")&" ","")&IF(Val!C{tv}>0.8,"B "&TEXT(Val!C{tv},"0%")&" ","")&IF(Val!D{tv}>0.8,"C "&TEXT(Val!D{tv},"0%"),"")&"（DCF 價值大多來自終值，對 WACC 與 g 較敏感）。",'
          f'"三情境終值現值占 EV "&TEXT(Val!B{tv},"0%")&"／"&TEXT(Val!C{tv},"0%")&"／"&TEXT(Val!D{tv},"0%")&"（A／B／C），皆未超過 80%。")')
    return s
obs.append(_r2pt(len(obs) + 1, True, ccy_fmt="#,##0", amt_fmt="#,##0.0", unit="（¥B）"))
rr = _RD.max_row + 2
_RD.cell(rr, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；v0.3 新增，上方各列保留）").font = Font(bold=True); rr += 1
for t_ in [lead] + props + obs:
    _RD.cell(rr, 1, t_); rr += 1
rr += 1
_RD.cell(rr, 1, "命題回覆（HTML 用；依本模型計算回答開頭的命題）").font = Font(bold=True); rr += 1
for t_ in answers:
    _RD.cell(rr, 1, t_); rr += 1
rr += 1
_RD.cell(rr, 1, "HTML 橫幅附註用：Comparable 目標價（相對倍數調整後）的上檔空間（＝Val 相對倍數調整後 Comparable 目標價 ÷ 現價 − 1；進加權目標價的是該值，橫幅中間格則放未調整的只看 Comparable 目標價）")
_RD.cell(rr, 2, f"=Val!B{vca}/Val!B{vpx}-1").number_format = "0.0%"
_RD["A1"] = str(_RD["A1"].value) + "；v0.1.0（2026-11-17）：新建（工作單 20261117_新增五家）"
_RD.cell(_RD.max_row + 2, 1, "v0.1.0（2026-11-17）：新建。結構參 accretech（日股 3 月年度）與 amat／lrcx（前段設備：終端市場需求指數 × 公司錨點）；FY2026 錨定公司 2026-10-30 首次發布的通期預想；單一應報導部門，營收構成為新設備（DRAM／NAND／非記憶體）與 Field Solutions 與其他。")
FD.yellow(IN)   # M2 節判斷參數黃底

# ---------------- 治理：build_id、輸入指紋、模型狀態（Sens 頁底部；README、HTML 同步；check_governance）----------------
import hashlib
def _build_id():
    h = hashlib.sha256(); here = os.path.dirname(os.path.abspath(__file__))
    for fn in ("build_tel.py", "tel_post.py"):
        h.update(fn.encode()); h.update(open(os.path.join(here, fn), "rb").read())
    return f"tel-v{VER}-{h.hexdigest()[:10]}"
BUILD_ID = _build_id()
_snap_f = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sens.json")
_SN = json.load(open(_snap_f, encoding="utf-8"))["_base"] if (os.path.exists(_snap_f) and not os.environ.get("TEL_OVERRIDE") and not os.environ.get("TEL_SNAP")) else {}
_r = SE.max_row + 3
SE.cell(_r, 1, "治理：輸入指紋、模型狀態與 build_id（Sens 為建置快照；改 In 頁後須重跑 python3 tel/sens_tel.py 與建置）").font = BOLD; _r += 1
SE.cell(_r, 1, "輸入指紋（本檔即時重算）")
_FP = f"In!$D$4:$F${IN.max_row}"
SE.cell(_r, 4, f"=SUMPRODUCT({_FP},ROW({_FP})*3+COLUMN({_FP}))").number_format = "0.000000"; _fp1 = _r; _r += 1
SE.cell(_r, 1, "輸入指紋（敏感度重算時的基準檔）")
if "fp" in _SN: c_ = SE.cell(_r, 4, _SN["fp"]); c_.number_format = "0.000000"; c_.font = BLUE
_fp2 = _r; _r += 1
SE.cell(_r, 1, "敏感度快照基準加權目標價（敏感度重算時）")
if "tp" in _SN: c_ = SE.cell(_r, 4, _SN["tp"]); c_.number_format = "#,##0.00"; c_.font = BLUE
_tp2 = _r; _r += 1
SE.cell(_r, 1, "模型狀態（CURRENT＝Sens 快照、HTML、README 與本檔輸入一致；STALE＝In 頁已改、須重跑 tel/sens_tel.py）").font = BOLD
SE.cell(_r, 4, f'=IF(AND(ISNUMBER(D{_fp2}),ISNUMBER(D{_tp2}),ABS(Val!B{vw}-D{_tp2})<0.005,ABS(D{_fp1}-D{_fp2})<=0.000000001*MAX(1,ABS(D{_fp2}))),"CURRENT","STALE")'); _r += 1
SE.cell(_r, 1, "build_id（README、Sens、HTML 三處相同）"); SE.cell(_r, 4, BUILD_ID); _r += 1

# ---------------- 三層協調 C4 兩讀數（Coord 頁；快照 tel/coord_alt.json）、個股 HTML 版面 v1.3、C0（公司實績）----------------
sys.path.insert(0, os.path.join(ROOT, "tools", "coord"))
import coord as COORD  # noqa: E402
_CR = COORD.add_sheet(wb, os.path.dirname(os.path.abspath(__file__)), SIL)
if _CR:
    COORD.stamp(wb, "v0.1.0", "v0.1.0", "2026-11-17")
    if not os.environ.get("TEL_OVERRIDE") and not os.environ.get("TEL_SNAP"):
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import tel_post  # noqa: E402
        print(tel_post.apply(wb, os.path.dirname(os.path.abspath(__file__)), SIL, _CR, PB, EQ_F))
_RD = wb["README"]; _RD["A1"] = str(_RD["A1"].value) + f"；build_id {BUILD_ID}"
wb.save(sys.argv[1])

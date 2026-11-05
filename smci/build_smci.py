# 美超微 Super Micro Computer（SMCI US）收支與估值模型 v0.2（AI 半導體 Project：伺服器組裝次產業，樣板＝wiwynn/、quanta/）
# 用法：python3 smci/build_smci.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度 FY(t)＝t−1 年 7 月～t 年 6 月（6 月底結束）。上游市場層為日曆年：FY(t) 市場＝½ × CY(t−1)＋½ × CY(t)（半年加權）。
# 金額：十億美元（US$B）；EPS：美元（GAAP 稀釋）。口徑：歸屬母公司淨利＝扣除特別股股利後、歸屬普通股的 GAAP 淨利。
# FY2026（2025-07～2026-06）已結束：FY2026 為實績錨點；FY2027–FY2029 由驅動因子推導；FY2030–FY2032 為收斂法。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
import cycle_position as CP
import valuation_rule as VR
import fcff_dcf as FD   # v0.3：標準版 DCF（無槓桿自由現金流＋WACC）共用模組（tools/fcff_dcf.py，原樣取自 claude/html-v1-nvidia）

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")

# ---------------- 讀入其他模型的成品（唯讀；不修改它們） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
TSMC_F = only("tsmc/dist/*.xlsx"); BR_F = only("markets/chip_bridge/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and (str(k).startswith("um") or str(k).startswith("u_") or str(k).startswith("g29_")):
        TS[k] = (r[4].value, r[0].row)
PLATS = ("bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth")
assert all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()
assert all(f"u_{p}_{y}" in TS for p in PLATS for y in (26, 27, 28)) and all(f"g29_{p}" in TS for p in PLATS), TS.keys()
def units(p, y):
    return TS[f"u_{p}_{y}"][0] if y < 29 else TS[f"u_{p}_28"][0] * (1 + TS[f"g29_{p}"][0])
# 商用 GPU（美超微可銷售的平台）：NVIDIA Blackwell、Rubin、Rubin Ultra、Hopper 與 AMD Instinct；不含 Google TPU、Trainium、其他 ASIC
# 世代分組：GB300 世代（Blackwell、Hopper；AMD 2026 的 MI355）、Rubin 世代（Rubin、Rubin Ultra；AMD 2027 起的 MI450 整櫃，Analogy）
OLD = {y: sum(units(p, y) for p in ("bw", "hp")) + (units("amd", y) if y == 26 else 0) for y in (26, 27, 28, 29)}
NEW = {y: sum(units(p, y) for p in ("rb", "ru")) + (units("amd", y) if y >= 27 else 0) for y in (26, 27, 28, 29)}
ALL = {y: sum(units(p, y) for p in PLATS) for y in (26, 27, 28, 29)}
U_SRC = f"{rel(TSMC_F)}｜In!E{TS['u_bw_26'][1]}:E{TS['g29_oth'][1]}"
BR = {}
for r in openpyxl.load_workbook(BR_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k in ("rack", "gpu_n"):
        BR[k] = (r[4].value, r[7].value, f"{rel(BR_F)}｜In!E{r[0].row}（GB300）、H{r[0].row}（VR200）")
assert len(BR) == 2, BR.keys()
# CF10：商用 GPU 系統市場只有一個值（Siliconomics Rack 頁）；本模型的世代分組與每顆系統價值必須重現同一市場（情境 A）
for j, y in enumerate((26, 27, 28, 29)):
    _m = OLD[y] * BR["rack"][0] / BR["gpu_n"][0] * 1000 + NEW[y] * BR["rack"][1] / BR["gpu_n"][1] * 1000
    _m *= {28: SIL.base("um28_B"), 29: SIL.base("um29_B")}.get(y, 1)   # v0.15：本庫只給基準（原始顆數 × 基準係數）
    assert abs(_m - SIL.NAMES["SIF_GPUSysMkt_Base"][j]) < 1e-6, ("商用 GPU 系統市場與 Siliconomics 基準不一致", y, _m, SIL.NAMES["SIF_GPUSysMkt_Base"][j])

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億美元（US$B）；每顆系統價值：千美元；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個。EPS 為 GAAP 稀釋、美元"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間，不當點估計（放情境或 Sens）。SemiAnalysis 未使用。公司原文＝美超微新聞稿（8-K）、財報說明稿、10-K（經二手摘要）；公司的 AI 占比、指引與展望為 Interested-party；券商與資料商估計皆為 Interested-party。商用 GPU 顆數與情境係數讀台積電模型、每顆系統價值讀橋接表（唯讀）"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
PR4Q = "https://ir.supermicro.com/news/news-details/2026/Supermicro-Announces-Fourth-Quarter-and-Full-Fiscal-Year-2026-Financial-Results/default.aspx"
PREP = "https://s204.q4cdn.com/707617056/files/doc_financials/2026/q4/V2/FQ426-SMCI-Prepared-Remarks.pdf"
MB0811 = "https://www.marketbeat.com/instant-alerts/super-micro-computer-q4-earnings-call-highlights-2026-08-11/"
PRQ2 = "https://www.sec.gov/Archives/edgar/data/1375365/000137536526000006/exhibit991_20251231.htm"
PRQ1U = "https://ir.supermicro.com/news/news-details/2026/Supermicro-Provides-First-Quarter-of-Fiscal-Year-2027-Preliminary-Business-Update/default.aspx"
PRQ1 = "https://www.sec.gov/Archives/edgar/data/1375365/000137536526000026/exhibit991_20260930.htm"   # v0.2：1Q FY27 新聞稿（8-K Ex.99.1，2026-11-03）
INVEST = "https://www.supermicro.com/en/pressreleases/supermicro-announces-completion-independent-investigation-and-continued-enhancement"
ST0820 = "https://stocktwits.com/news-articles/markets/equity/smci-stock-gains-supermicro-clears-senior-management-export-control-case/cZYIq3bRJZ5"
FIN0611 = "https://ir.supermicro.com/news/news-details/2026/Supermicro-Announces-Pricing-of-Equity-and-Equity-Linked-Financing-Transactions-To-Fund-AI-Orders/default.aspx"
K10 = "https://www.sec.gov/Archives/edgar/data/0001375365/000137536526000022/smci-20260630.htm"
STK10 = "https://stockti.com/super-micro-revenue-surges-78-but-margin-pressures-and-control-gaps-linger"
FUT = "https://futurumgroup.com/insights/supermicro-q2-fy2026-delivers-breakout-ai-gpu-platform-revenue"
SA = "https://stockanalysis.com/stocks/smci/statistics/"
SAF = "https://stockanalysis.com/stocks/smci/forecast/"
SAR = "https://stockanalysis.com/stocks/smci/financials/ratios/"
SAI = "https://stockanalysis.com/stocks/smci/financials/"
YF = "https://finance.yahoo.com/quote/SMCI/analysis/"
TICK = "https://tickflow.io/stock/SMCI/forecast"
SWS = "https://simplywall.st/stocks/us/tech/nasdaq-smci/super-micro-computer/future"
MT = "https://www.macrotrends.net/stocks/charts/SMCI/super-micro-computer/pe-ratio"
MTNI = "https://www.macrotrends.net/stocks/charts/SMCI/super-micro-computer/net-income"
MTREV = "https://www.macrotrends.net/stocks/charts/SMCI/super-micro-computer/revenue"
FR = "https://fullratio.com/stocks/nasdaq-smci/pe-ratio"
P_DELL = "https://stockanalysis.com/stocks/dell/statistics/"
P_HPE = "https://stockanalysis.com/stocks/hpe/statistics/"
P_WY = "https://stockanalysis.com/quote/tpe/6669/statistics/"
P_QC = "https://stockanalysis.com/quote/tpe/2382/statistics/"
P_WS = "https://stockanalysis.com/quote/tpe/3231/statistics/"
URLNAME = {PR4Q: "公司 FY2026 第四季新聞稿（2026-08-11）", PREP: "公司 4Q FY26 財報說明稿", MB0811: "MarketBeat 2026-08-11（法說重點）",
           PRQ2: "公司 2Q FY26 新聞稿（8-K）", PRQ1U: "公司 1Q FY27 初步營運更新（2026-10-20）", PRQ1: "公司 1Q FY27 新聞稿（8-K，2026-11-03）", INVEST: "公司 2026-08 獨立調查完成公告",
           ST0820: "Stocktwits 2026-08-20", FIN0611: "公司 2026-06-11 股權與特別股定價公告", K10: "公司 FY2026 10-K（2026-08-31）", STK10: "StockTi（10-K 摘要）",
           FUT: "Futurum（2Q FY26）", SA: "stockanalysis 統計頁", SAF: "stockanalysis 預測頁", SAR: "stockanalysis 比率頁", SAI: "stockanalysis 損益表",
           YF: "Yahoo Finance 分析頁", TICK: "tickflow 預測頁", SWS: "Simply Wall St", MT: "Macrotrends P/E", MTNI: "Macrotrends 淨利", MTREV: "Macrotrends 營收",
           FR: "fullratio P/E", P_DELL: "stockanalysis（Dell）", P_HPE: "stockanalysis（HPE）", P_WY: "stockanalysis（緯穎）", P_QC: "stockanalysis（廣達）", P_WS: "stockanalysis（緯創）"}

sec("A. 商用 GPU 顆數與每顆系統價值（讀 tsmc、markets/chip_bridge 成品；情境 A（原始顆數），B／C 依顆數係數縮放）")
for y in (26, 27, 28, 29):
    inp(f"uo_{y}", f"GB300 世代商用 GPU 顆數 CY20{y}（Blackwell＋Hopper" + ("＋AMD MI355）" if y == 26 else "）"), "百萬顆", None, round(OLD[y], 6), None, "Derived",
        U_SRC if y == 26 else "同上", "台積電模型八類加速器中 NVIDIA Blackwell、Hopper（2026 另含 AMD）；2029＝2028 × (1＋各類成長)" if y == 26 else "同上")
for y in (26, 27, 28, 29):
    inp(f"un_{y}", f"Rubin 世代商用 GPU 顆數 CY20{y}（Rubin＋Rubin Ultra" + ("）" if y == 26 else "＋AMD MI450）"), "百萬顆", None, round(NEW[y], 6), None, "Derived",
        U_SRC if y == 26 else "同上", "AMD 自 2027 年起改計 Rubin 世代價值（MI450 整櫃；Analogy）" if y == 27 else "同上")
for y in (26, 27, 28, 29):
    inp(f"ut_{y}", f"全部 AI 加速器顆數 CY20{y}（對照）", "百萬顆", None, round(ALL[y], 6), None, "Derived", U_SRC if y == 26 else "同上", "八類合計；商用 GPU 以外為 TPU、Trainium、其他 ASIC（美超微不銷售）" if y == 26 else "同上")
inp("v_gb", "每顆系統價值：GB300 世代（整櫃買方價 ÷ 每架 GPU 封裝數）", "千美元/顆", None, f"={BR['rack'][0]}/{BR['gpu_n'][0]}*1000", None, "Derived", BR["rack"][2],
    f"Tokenomics v20.1 機架單價 {BR['rack'][0]} 百萬美元 ÷ {BR['gpu_n'][0]} 顆（經橋接表；唯讀）；HGX 八卡伺服器與 AMD MI355 以同值近似（Analogy）")
inp("v_vr", "每顆系統價值：Rubin 世代（VR200 整櫃買方價 ÷ 每架 GPU 封裝數）", "千美元/顆", None, f"={BR['rack'][1]}/{BR['gpu_n'][1]}*1000", None, "Derived", BR["rack"][2],
    f"機架單價 {BR['rack'][1]} 百萬美元 ÷ {BR['gpu_n'][1]} 顆；Rubin Ultra 與 AMD MI450 未另訂（Analogy：同 VR200）")
inp("vpg_28", "每顆系統價值 2028 ÷ 2027（Rubin 世代）", "x", 0.85, 1.0, 1.2, "Assumed", "Rubin Ultra（Kyber）整櫃價格：未找到", "與廣達、緯穎模型的每 GW 價值變動同區間")
inp("vpg_29", "每顆系統價值 2029 ÷ 2027（Rubin 世代）", "x", 0.85, 1.0, 1.2, "Assumed", "同上", "—")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}（個股自訂）：2028 顆數 ÷ 原始顆數", "x", None, TS[f"um28_{s}"][0], None, "Assumed", f"{rel(TSMC_F)}｜In!E{TS[f'um28_{s}'][1]}", SIL.scen_note(s, 28, TS[f"um28_{s}"][0]))
    inp(f"um29_{s}", f"情境 {s}（個股自訂）：2029 顆數 ÷ 原始顆數", "x", None, TS[f"um29_{s}"][0], None, "Assumed", f"{rel(TSMC_F)}｜In!E{TS[f'um29_{s}'][1]}", SIL.scen_note(s, 29, TS[f"um29_{s}"][0]))
inp("mf", "AI 部署係數（Sens 用；CY2027 起）", "x", 0.8, 1.0, 1.2, "Assumed", "台積電模型 CY2027 商用 GPU 顆數約 1.4 倍（8.35 → 11.2 百萬顆）", "±20%")

sec("B. 營收結構、份額（FY2027 以公司指引校準）")
inp("ais_25", "FY2025 AI GPU 平台 ÷ 營收", "%", 0.65, 0.70, 0.75, "Interested-party", "公司：FY2025 AI GPU 平台占營收逾七成（法說；原文未逐字取得）", "量化第二來源：未找到")
inp("ais_1h", "1H FY26（2025-07～12）AI GPU 平台 ÷ 營收", "%", 0.75, 0.85, 0.90, "Assumed", FUT + "：2Q FY26 營收由 GB300 NVL72 與 AMD MI355 量產帶動（未給比例）", "公司揭露比例：未找到")
inp("ais_q3", "3Q FY26（2026-01～03）AI ÷ 營收", "%", 0.78, 0.80, 0.85, "Interested-party", PREP + "：3Q 逾 80%（公司）", MB0811 + "：同")
inp("ais_q4", "4Q FY26（2026-04～06）AI ÷ 營收", "%", 0.55, 0.60, 0.65, "Interested-party", PREP + "：約 60%，專案出貨時點使比重下降（公司）", MB0811 + "：同；公司預期未來逾 80% 營收與 AI 相關、純 AI 約 60–70%")
inp("g27", "FY2027 營收指引（情境 A＝上緣、B＝中點、C＝下緣）", "US$B", 72.0, 77.0, 82.0, "Interested-party", PRQ1 + "：FY2027 營收 720～820 億美元（公司，2026-11-03 上修；原 2026-08-11 為 650～720 億）", PRQ1 + "：1Q FY27 實績 199.46 億美元＋2Q 指引 195～215 億（1H 約 404 億；全年指引隱含 2H 316～416 億）")
inp("rel_28", "美超微份額 FY2028 ÷ FY2027（A＝高、B＝基準、C＝低）", "x", 0.85, 0.95, 1.0, "Assumed", "Dell、HPE 與 ODM 直供（廣達、緯創、鴻海）爭取 neocloud 與企業訂單；4Q FY26 逾 600 億美元新訂單（公司）", "最大客戶占 FY2026 營收 28%（公司說明稿）：大客戶轉單為下檔")
inp("rel_29", "美超微份額 FY2029 ÷ FY2027", "x", 0.80, 0.92, 1.0, "Assumed", "同上", "—")
inp("g_nai27", "非 AI（企業、儲存、邊緣）營收年增 FY2027", "%", 0.0, 0.15, 0.30, "Assumed", PREP + "：企業與通路營收 4Q FY26 年增 172%（含 AI，公司）", "非 AI 營收分項：未找到；FY2026 非 AI 為殘差（Q 頁）")
inp("g_nai28", "非 AI 營收年增 FY2028", "%", -0.05, 0.08, 0.20, "Assumed", "同上", "—")
inp("g_nai29", "非 AI 營收年增 FY2029", "%", -0.05, 0.06, 0.15, "Assumed", "同上", "—")

sec("C. 毛利率（A 最有利、C 最不利）")
inp("gm_nai", "非 AI 營收毛利率", "%", 0.13, 0.15, 0.18, "Assumed", "FY2023（AI 占比低）GAAP 毛利率約 18%（公司年報，stockanalysis 損益表：毛利 1,283 ÷ 營收 7,123）", "部門別毛利率：未找到；公司 DCBBS 目標毛利率 20%＋（Futurum 轉述）")
inp("dgm_27", "FY2027 AI 毛利率 − FY2026 校準值（A＝高、B＝基準、C＝低）", "pt", -0.005, 0.005, 0.015, "Interested-party", PRQ1 + "：1Q FY27 實績 GAAP 毛利率 12.7%（non-GAAP 12.8%；初步 12.0%–12.5%，客戶與產品組合，公司）；v0.2 未改情境值（待 Andy 判斷）", PREP + "：4Q FY26 17.6%（non-GAAP）約四分之三來自組合與遞延至 FY27 的合約、四分之一來自關稅與存貨跌價回轉（可能不重複）")
inp("dgm_28", "FY2028–FY2029 AI 毛利率 − FY2026 校準值", "pt", -0.015, 0.003, 0.015, "Assumed", "C：Dell、ODM 價格競爭與 GPU 轉手比重上升；A：液冷與 DCBBS 整體解決方案", "FY2024 13.8% → FY2025 11.1% → FY2026 10.8%（GAAP；stockanalysis 損益表）")

sec("D. 1Q FY27 實績（Recon 對帳用；v0.2 由初步數字改為 2026-11-03 實績）")
inp("q1_rev", "1Q FY27（2026-07～09）營收（實績）", "US$B", None, 19.945832, None, "Verified", PRQ1 + "：19,945,832 千美元（公司，未經審計）", PRQ1U + "：10-20 初步 185～195 億；原指引 145～155 億（4Q FY26 財報說明稿）")
inp("q1_gm", "1Q FY27 毛利率（GAAP，實績）", "%", None, "=2.541043/19.945832", None, "Verified", PRQ1 + "：毛利 2,541,043 ÷ 營收 19,945,832 千美元＝12.7%（公司）", "初步 12.0%–12.5%（2026-10-20）；原指引 10.4%–10.8%（4Q FY26 財報說明稿）")

sec("E. 營業費用、淨利係數、特別股與股數（FY2027 起）")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.3, 0.4, 0.6, "Derived", PREP + "：1Q FY27 GAAP 營業費用約 4.53 億美元（年化約 18.1 億，較 FY2026 14.57 億 +24%）", "FY2026 實測：營業費用 +24%、營收 +78%（Q 頁），彈性約 0.3；1Q FY27 實績營業費用 4.766 億（年化 19.1 億，+31%；公司 2026-11-03）")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.002, 0.0, 0.002, "Assumed", "無", "營益率約 8%，±0.2 個百分點約 ±2.5% 的營業利益")
inp("nio_f", "淨利（特別股股利前）÷ 營業利益", "%", 0.68, 0.75, 0.80, "Derived", "FY2026 2,230.5 ÷ 2,770.5＝0.805；稅率 FY2026 19.9%、1Q FY27 指引 20.1%（公司）", "1Q FY27 指引隱含約 0.68；1Q FY27 實績 1.5942 ÷ 2.0645＝0.772（利息淨支出 0.28 億、稅率 21.4%；公司 2026-11-03）；v0.2 基準值不變")
inp("pref_div", "強制可轉換特別股股利（年）", "US$B", None, "=86.25*50/1000*0.07", None, "Derived", FIN0611 + "：7,500 萬股存託股＋超額配售 1,125 萬股、每股 50 美元、股利率 7.0%，2029-06-01 強制轉換（公司）", PR4Q + "：4Q FY26 募得淨額 56 億美元，其中強制可轉換特別股 42 億（公司）；超額配售全數執行為 Assumed")
inp("sh27", "FY2027 GAAP 稀釋股數（1Q 實績與 2Q 指引平均）", "十億股", None, "=(0.749351+0.753)/2", None, "Derived", PRQ1 + "：1Q FY27 GAAP 稀釋加權股數 7.4935 億、2Q 指引 7.53 億（公司）", "v0.1 為 1Q 指引 7.45 億（" + PR4Q + "）")
inp("sh_g", "股數年變動（員工股票報酬、可轉債）", "%", 0.0, 0.01, 0.03, "Assumed", "FY2025 6.284 → FY2026 6.973 億股（GAAP 稀釋，含 6 月增資；公司）", "1Q FY27 指引股票報酬費用約 1.06 億美元／季（公司）")
inp("conv_sh", "特別股轉換新增股數（FY2030 起；最低～最高轉換率中點）", "十億股", "=86.25*1.5152/1000", "=86.25*(1.5152+1.8182)/2/1000", "=86.25*1.8182/1000", "Derived", FIN0611 + "：每股存託股轉換 1.5152～1.8182 股普通股（公司）", "轉換後特別股股利消失；FY2029 股利以全年計（實際 2029-06 轉換，差約一個月）")
inp("sbc_ps", "GAAP → non-GAAP EPS 調整（每股；Recon 對共識用）", "US$", None, "=(0.100223-0.020048)*4/sh27", None, "Derived", PRQ1 + "：1Q FY27 股票報酬 1.002 億美元、稅效果 0.200 億（公司實績；v0.1 用 1Q 指引 1.06／0.32 億）", "1Q FY27 non-GAAP 1.81 − GAAP 1.75＝0.06（含參與證券分配與股數差）；FY2026 0.37（公司）")

sec("F. 共識（只作 Recon 對帳，不作校準；共識 EPS 為 non-GAAP）")
inp("c_rev27", "共識 FY2027 營收", "US$B", 63.5, 70.0, 82.0, "Interested-party", SAF + "：平均 700 億美元（15 位，低 635、高 820；2026-10-22）", YF + "：同 700 億（15 位）")
inp("c_rev28", "共識 FY2028 營收", "US$B", 70.84, 81.4, 92.0, "Interested-party", YF + "：平均 814 億美元（15 位）", "stockanalysis FY2028 需付費：未找到")
inp("c_eps27", "共識 FY2027 EPS（non-GAAP）", "US$", 3.50, 4.67, 6.40, "Interested-party", SAF + "：4.67（15 位，低 3.50、高 6.40）", YF + "：4.67（18 位）")
inp("c_eps28", "共識 FY2028 EPS（non-GAAP）", "US$", None, 5.59, None, "Interested-party", YF + "：5.59（17 位）", "第二來源：未找到（tickflow 季度數字前後不一致，不採用）")

sec("G. 價格與市場倍數")
inp("px", "股價（2026-10-23 收盤）", "US$", None, 32.94, None, "Interested-party", SA + "：10-23 收盤 32.94", FR + "：10-23 價格 32.94、近 12 月 EPS 3.65")
inp("pe_ttm", "近 12 月 P/E（stockanalysis，10-23）", "x", None, 10.12, None, "Interested-party", SA, "fullratio 9.02（以 non-GAAP EPS）")
inp("fpe", "前瞻 P/E（stockanalysis，10-23）", "x", None, 7.06, None, "Interested-party", SA, "以 non-GAAP 共識 EPS")
inp("cons_feps", "共識前瞻 12 月 EPS（10-23 價格 ÷ 前瞻 P/E；non-GAAP）", "US$", None, "=px/fpe", None, "Derived", "stockanalysis 前瞻 P/E 7.06（2026-10-23）", "≈ 共識 FY2027 EPS 4.67")
inp("cons_tp", "市場共識目標價（賣方平均）", "US$", 15.0, 43.27, 60.0, "Interested-party", SAF + "：18 位平均 43.27、中位數 45、最低 15、最高 60（S&P Global，2026-10-22）", TICK + "：25 位平均 41.44（2026-10-24）；Simply Wall St 08-12：41.75（16 位）")
REV10 = [2225, 2485, 3360, 3500, 3339, 3557, 5196, 7123, 14989, 21972]
NI10 = [72, 67, 46, 72, 84, 112, 285, 640, 1153, 1049]
inp("nim10", "十年平均淨利率（FY2016–FY2025，淨利 ÷ 營收）", "%", None, "=AVERAGE(" + ",".join(f"{n}/{v}" for n, v in zip(NI10, REV10)) + ")", None, "Derived",
    MTREV + "、" + MTNI + "：各年營收與淨利（百萬美元；FY2016 2,225／72 … FY2023 7,123／640）", SAI + "：FY2024 14,989／1,153、FY2025 21,972／1,049")

sec("H. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_g", "PL 延伸期營收年增的收斂目標（情境 B；FY2032；v0.3 起不再用於 DCF，DCF 永續成長率見 H2 節 g_tv）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "v0.3 刪除 tp_r、tp_roe（舊股東現金流法 DCF 參數）；本值只決定 FY2030–FY2032 營收收斂路徑")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（Andy 2026-10-25 11:53 定 DCF／Comparable 各半）", "—")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 特別股股利前淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）；FY2030 起特別股已轉換，股利消失", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（情境 B FY2029 與十年平均的中點）", "%", None, 0, None, "Derived", "PL 頁情境 B FY2029、In 頁 nim10", "—")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（＝十年平均淨利率）", "%", None, "=nim10", None, "Derived", "In 頁 nim10", "—")
# v0.3：標準版 DCF（無槓桿自由現金流 FCFF＋WACC；tools/fcff_dcf.py；docs/plan/個股DCF標準規格.md v1.0）。評價日沿用 v0.2 DCF 基準日 2026-12-31（FY2027 只計 2027-01～06 半年）；資產負債為 1Q FY27（2026-09-30）
SA_BS = "stockanalysis（S&P Global）Supermicro 季資產負債表 1Q FY27（2026-09-30，US$ 百萬）https://stockanalysis.com/stocks/smci/financials/balance-sheet/?p=quarterly（2026-11-05 讀取）"
SA_CF = "stockanalysis（S&P Global）Supermicro 年度現金流量表 FY2024–FY2026（US$ 百萬）https://stockanalysis.com/stocks/smci/financials/cash-flow-statement/（2026-11-05 讀取）"
SA_IS = "stockanalysis（S&P Global）Supermicro 年度損益表 FY2024–FY2026（GAAP，US$ 百萬）https://stockanalysis.com/stocks/smci/financials/income-statement/（2026-11-05 讀取）"
SA_BY = "stockanalysis（S&P Global）Supermicro 年度資產負債表 FY2024–FY2026 https://stockanalysis.com/stocks/smci/financials/balance-sheet/（2026-11-05 讀取）"
NO2 = "第二來源：無（公司 10-Q 本次未直接讀取）"
_FD = FD.in_rows(dict(
    date="2026-12-31（同 v0.2 DCF 基準日）", ccy="US$", amt="US$B", shu="十億股",
    rf=(0.0500, "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；2026-11-05 讀取；同 NVIDIA v0.3）",
        "FRED DGS10（https://fred.stlouisfed.org/series/DGS10）：10-28～10-30 5.05%／4.98%／4.94%；與股價同日（2026-10-23）", "無風險利率（美國 10 年期公債殖利率，2026-10-23）"),
    beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Derived",
                "自算產業 β：US 伺服器／EMS 同業 8 家中位數（DELL SMCI HPE HPQ NTAP CLS FLEX JBL；Yahoo Finance 月報酬 2021-11～2026-10，對 S&P 500（^GSPC） 迴歸，Hamada 去槓桿、現金修正）無槓桿 β 1.48，Blume 調整後 βU 1.322（In 頁 beta_u）；依美超微 D/E（市值；In 頁 beta_de）與 tax_w 以 Hamada 重新加槓桿：βL＝βU ×（1＋(1−t) × D/E）",
                "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）。Sens 低／高（Analogy；v0.3.2 起）：β_u ∓ beta_du（0.3），各自同式重新加槓桿（工作單第二輪修正，Andy 2026-11-05 23:16 定案；tools/fcff_dcf.py v2 write_sens）"),
    kd=("=0.19457/((5.059+9.26)/2)", "Derived", SA_IS + "：FY2026 利息費用 194.57 ÷ 平均總債務（FY2025 5,059、FY2026 9,260；" + SA_BY + "）", "可轉債票息低；含租賃"),
    tax_w=("=(0.0521+0.1301+0.1996)/3", "Derived", SA_IS + "：有效稅率 FY2024 5.21%、FY2025 13.01%、FY2026 19.96%（所得稅 63.29／156.85／556.33 ÷ 稅前 1,216／1,206／2,787）",
           "FY2026、1Q FY27 指引約 20%，近三年平均偏低（NOPAT 偏高）；美國法定 21%"),
    debt=("=1.963+6.683", "Verified", SA_BS + "：短期借款 1,963＋長期債務 6,683（含可轉債；總債務 9,234 含租賃 588）", NO2),
    lease=(0.58805, "Verified", SA_BS + "：長期租賃負債 588.05", NO2),
    cash=(6.897, "Verified", SA_BS + "：現金及約當現金 6,897（含 2026-06 增資與特別股募資）", NO2),
    inv=(0.0, "Verified", SA_BS + "：未列長期投資", NO2),
    mi=("=0.00016+pref_div*29/12", "Derived", SA_BS + "：非控制權益 0.16；另加強制可轉換特別股在 2029-06 轉換前應付的股利（pref_div × 29 個月，未折現；特別股本身以轉換股數計入 sh_v）", "In 頁 pref_div、conv_sh（公司 2026-06-11 定價公告）"),
    stub=dict(bs_date="2026-09-30", val_date="2026-12-31", bs_src=SA_BS + "（最新一季 1Q FY27，已公布）",
              div=("=0", "Derived", "普通股不配息（stockanalysis 統計頁 2026-11-05）；強制可轉換特別股 2027-01 起至 2029-06 轉換前的股利已計入 mi_v（pref_div × 29 個月）",
                   "2026-10～12 的特別股股利（約一季 0.076）是否已宣告未查得，本版未扣；買回無已宣告金額")),
    dps=(0, "Verified", "普通股不配息（stockanalysis 統計頁 2026-11-05）", "—"),
    sh=("=0.65813+conv_sh", "Derived", SA_BS + "：普通股流通股數 658.13 百萬股（2026-09-30）＋特別股轉換新增股數 conv_sh（中點）", "可轉債視為債務（已在 debt_v），不另計稀釋；GAAP 稀釋股數 FY2026 697 百萬股（含員工股票報酬與可轉債）"),
    fc_da=("=0.09027/39.063", "=(0.03869/14.989+0.05834/21.972+0.09027/39.063)/3", "=0.05834/21.972", "Derived",
           SA_CF + "：折舊攤銷 38.69／58.34／90.27 ÷ 營收 14,989／21,972／39,063（" + SA_IS + "）近三年平均", "低＝FY2026、高＝FY2025"),
    fc_cx=("=0.162/39.063", "=(0.12428/14.989+0.12721/21.972+0.162/39.063)/3", "=0.12428/14.989", "Derived",
           SA_CF + "：資本支出 124.28／127.21／162 ÷ 同年營收，近三年平均", "Sens 用 fc_cx_d ±1pt"),
    fc_cx_ss=("=MAX(fc_cx,fc_da_ss*(1+g_tv))", "Derived", "長期穩態：資本支出 ≥ 折舊 ×（1＋g）（規格一致性檢查）", "—"),
    fc_wc=(0.20, "=((2.749+4.333-1.472-0.19305)/14.989+(2.359+4.68-1.282-0.36874)/21.972+(6.263+12.896-2.247-1.578)/39.063)/3", 0.40, "Derived",
           SA_BY + "：營運資金＝應收＋存貨 − 應付 − 合約負債（預收）；FY2024 2,749＋4,333 − 1,472 − 193、FY2025 2,359＋4,680 − 1,282 − 369、FY2026 6,263＋12,896 − 2,247 − 1,578，各 ÷ 同年營收後平均（36.1%、24.5%、39.3%）",
           "1Q FY27 應收 9,685＋存貨 15,054 − 應付 2,727 − 預收 4,329（季資產負債表）；Sens 低／高 20%／40%（Assumed）"),
))
_FD[0] = ("sec", "H2. 折現率與無槓桿自由現金流（標準版 DCF；tools/fcff_dcf.py；評價日 2026-12-31；黃底＝Andy 可改）")
rows += _FD
inp("beta_u", "產業無槓桿 β（採用）：自算 US 伺服器／EMS 同業中位數 Blume 調整（0.67 × 中位數＋0.33）", "x", None, 1.322, None, "Analogy",
    "自算產業 β：US 伺服器／EMS 同業 8 家（DELL SMCI HPE HPQ NTAP CLS FLEX JBL）中位數（Yahoo Finance 含股利調整月收盤價 2021-11～2026-10 共 60 個月，對 S&P 500（^GSPC） 迴歸，Hamada 去槓桿、現金修正）1.48 → Blume 調整 1.322",
    "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；本組不變（工作單第二輪修正：只改 Sens 規則）")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務本金 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 稀釋股數 sh_v)", "與 WACC 權重同一口徑")
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3", "取代上一版的低＝中位數 × 0.8 再 Blume、高＝未調整中位數")
TP_KEYS += ["tp_g", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]

r = 4; INROW = {}
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, s1, s2 = srow
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): IN.cell(r, j).font = BLUE
    for u_, nm_ in URLNAME.items():          # 網址只列在 Sources 頁；In 頁寫來源名稱
        s1 = s1.replace(u_, nm_); s2 = s2.replace(u_, nm_)
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    INROW[key] = r; r += 1
for k in TP_KEYS:
    for c in range(1, 10): IN.cell(INROW[k], c).fill = KEY
FD.yellow(IN)
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億美元；EPS 為美元、GAAP 稀釋；股數為十億股）。藍字＝輸入；黑字＝公式。會計年度 FY2026＝2025-07～2026-06。實績取自美超微新聞稿（8-K）與財報說明稿"
Q["A2"] = "FY2026 錨點＝全年實績（公司 2026-08-11 公布，10-K 2026-08-31 申報；新聞稿標示未經審計，需待審計確認）。3Q 營收、營業費用、營業利益、淨利＝全年 − 1H − 4Q（Derived）；毛利＝營業利益＋營業費用。股數＝淨利 ÷ 公司揭露稀釋 EPS（可轉債加回使其與揭露稀釋股數略有差異）"
hdr(Q, 3, ["項目", "FY2025A", "1H FY26（2025-07～12）", "3Q FY26（2026-01～03）", "4Q FY26（2026-04～06）", "FY2026A（錨點）", "標記", "來源1（公司）", "來源2／說明"])
S1 = "公司 FY2026 第四季新聞稿（2026-08-11）；1H：公司 2Q FY26 新聞稿（8-K）"
S2 = "stockanalysis 損益表：FY2026 營收 39,063、毛利 4,227、營業利益 2,770（百萬美元）"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2=S2, ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"F{r}"] = ann or f"=SUM(C{r}:E{r})"; Q[f"F{r}"].number_format = fmt
    Q.cell(r, 7, tag); Q.cell(r, 8, s1); Q.cell(r, 9, s2)
qrow(4, "營收", {"B": 21.972, "C": 17.7003, "D": "=39.063-C4-E4", "E": 11.12}, tag="Verified／Derived（3Q）", s2="1Q FY26 5.0178、2Q 12.6825（8-K）；FY2026 39.063（stockanalysis）")
qrow(5, "毛利（＝營業利益＋營業費用）", {c: f"={c}8+{c}7" for c in "BCDE"}, tag="Derived", s2="GAAP 毛利率：FY2025 11.1%、1H 7.2%、3Q 9.9%、4Q 17.5%、FY2026 10.8%（公司）")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEF"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2="—")
qrow(7, "營業費用（GAAP）", {"B": 1.1769, "C": 0.6094, "D": "=1.4568-C7-E7", "E": 0.4546}, tag="Verified／Derived（3Q）", s2="FY2026 1.4568（公司）")
qrow(8, "營業利益", {"B": 1.253, "C": 0.6566, "D": "=2.7705-C8-E8", "E": 1.4881}, tag="Verified／Derived（3Q）", s2="FY2026 2.7705（公司）")
qrow(9, "營業利益率", {c: f"={c}8/{c}4" for c in "BCDEF"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2="—")
qrow(10, "歸屬母公司淨利（GAAP）", {"B": 1.0489, "C": 0.5688, "D": "=2.2305-C10-E10", "E": 1.1782}, tag="Verified／Derived（3Q）", s2="3Q 4.83 億（公司說明稿）；stockanalysis FY2026 淨利 2,196（可能已扣特別股股利，未採用）")
qrow(11, "淨利 ÷ 營業利益（含業外、稅）", {c: f"={c}10/{c}8" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(14, "  公司揭露稀釋 EPS（GAAP）", {"B": 1.68, "C": 0.86, "D": 0.72, "E": 1.62}, fmt="0.00", ann="=3.26", tag="Verified", s2="non-GAAP：FY2025 2.06、FY2026 3.63、4Q 1.70（公司）")
qrow(12, "股數（十億股；＝淨利 ÷ 揭露稀釋 EPS）", {c: f"={c}10/{c}14" for c in "BCDE"}, fmt="0.0000", ann="=F10/F14", tag="Derived", s2="淨利 ÷ 揭露 EPS；FY2026 揭露稀釋股數 6.973 億（第 15 列）")
qrow(13, "EPS（GAAP 稀釋）", {c: f"={c}10/{c}12" for c in "BCDE"}, fmt="0.00", ann="=F10/F12", tag="Derived", s1="淨利 ÷ 股數", s2="—")
qrow(15, "  公司揭露稀釋股數（對照）", {"B": 0.6284, "C": 0.6741, "D": 0.692, "E": 0.705}, fmt="0.0000", ann="=0.6973", tag="Verified", s2="4Q 7.050 億、3Q 6.92 億（公司說明稿）")
qrow(16, "AI GPU 平台營收（＝營收 × AI 占比）", {"B": "=B4*ais_25", "C": "=C4*ais_1h", "D": "=D4*ais_q3", "E": "=E4*ais_q4"}, tag="Derived／Interested-party", s1="公司比例（In 頁 B 節）", s2="AI 營收金額未揭露：未找到")
qrow(17, "非 AI（企業、儲存、邊緣等）", {c: f"={c}4-{c}16" for c in "BCDE"}, tag="Derived", s1="—", s2="—")
Q.cell(19, 1, "FY2026 錨點的檢查（參考；校準未使用券商數字）").font = BOLD
Q.cell(20, 1, "FY2026 營收年增"); Q["B20"] = "=F4/B4-1"
Q.cell(21, 1, "FY2026 營業費用年增"); Q["B21"] = "=F7/B7-1"
Q.cell(22, 1, "營業費用彈性（營業費用年增 ÷ 營收年增；In 頁 phi 的對照）"); Q["B22"] = "=B21/B20"
Q.cell(23, 1, "FY2026 AI 占營收（Derived；1H、3Q、4Q 加權）"); Q["B23"] = "=F16/F4"
Q.cell(24, 1, "FY2026 毛利率（不含 4Q；1H＋3Q）"); Q["B24"] = "=(C5+D5)/(C4+D4)"
Q.cell(25, 1, "2026-01～06（2H FY26）AI GPU 平台營收（份額校準用）"); Q["B25"] = "=D16+E16"
Q.cell(26, 1, "FY2026 EPS 自算 − 公司揭露（應為 0）"); Q["B26"] = "=F13-F14"
for rr in range(20, 27): Q[f"B{rr}"].number_format = "0.0%" if rr not in (22, 25, 26) else "0.000"
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 7), ("oi", 8), ("ni", 10), ("sh", 12), ("eps", 13), ("ai", 16), ("nai", 17)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$F${rr}")
defname("a26_aih2", "Q!$B$25")
Q.column_dimensions["A"].width = 60
for c in "BCDEF": Q.column_dimensions[c].width = 15
Q.column_dimensions["G"].width = 26; Q.column_dimensions["H"].width = 70; Q.column_dimensions["I"].width = 70
Q.freeze_panes = "B4"

# ---------------- 引擎 ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G", 2030: "H", 2031: "I", 2032: "J"}
FADE = (2030, 2031, 2032); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
CYS = (2026, 2027, 2028, 2029); F3 = (2027, 2028, 2029)
SCN_P = {"A": dict(g27="g27_hi", rel="rel_{y}_hi", d27="dgm_27_hi", d28="dgm_28_hi"),
         "B": dict(g27="g27", rel="rel_{y}", d27="dgm_27", d28="dgm_28"),
         "C": dict(g27="g27_lo", rel="rel_{y}_lo", d27="dgm_27_lo", d28="dgm_28_lo")}
_mfy = lambda m, y: m[y] if isinstance(m, dict) else m   # Siliconomics v0.13：mf 可為年別 dict（Sens「Wattonomics 錨定」案例）；預設字串不變
DEFP = dict(scn="A", vpg="vpg_{y}", mf="mf", gn="g_nai{y}", gmn="gm_nai", opxa="0", nio="nio_f", phi="phi", fg="fade_g{s}", fm="fade_m{s}", rel_mult="1")
def engine(ws, r0, title, **over):
    s = over.get("scn", "A")
    P = dict(DEFP); P.update(SCN_P[s]); P.update(over)
    p = lambda k, y=None: P[k].format(y=(y % 100 if y else ""), s=s)
    um = lambda y: {2026: "", 2027: "", 2028: f"*um28_{s}", 2029: f"*um29_{s}"}[y]
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026A（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.00", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 11, note); R[key] = r; r += 1
    def sec_(t):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
    c_ = lambda y: FC[y]
    mfx = lambda y: "" if y == 2026 else f"*{_mfy(P['mf'], y)}"
    sec_("A. 商用 GPU 系統市場（D–G 欄的「CY」列為日曆年 CY2026–CY2029；「FY」列為美超微會計年度）")
    row("uo", "GB300 世代商用 GPU 顆數（CY）", "百萬顆", {c_(y): f"=uo_{y % 100}{um(y)}{mfx(y)}" for y in CYS}, fmt="0.00",
        note="Blackwell＋Hopper（2026 另含 AMD MI355）；台積電模型顆數 × 情境係數")
    row("un", "Rubin 世代商用 GPU 顆數（CY）", "百萬顆", {c_(y): f"=un_{y % 100}{um(y)}{mfx(y)}" for y in CYS}, fmt="0.00",
        note="Rubin＋Rubin Ultra（2027 起另含 AMD MI450）")
    row("vo", "每顆系統價值：GB300 世代（CY）", "千美元/顆", {c_(y): "=v_gb" for y in CYS}, fmt="0.0")
    row("vn", "每顆系統價值：Rubin 世代（CY）", "千美元/顆", {"D": "=v_vr", "E": "=v_vr", **{c_(y): f"=v_vr*{p('vpg', y)}" for y in (2028, 2029)}}, fmt="0.0",
        note="整櫃買方價 ÷ 每架 GPU 封裝數（橋接表）")
    row("mcy", "商用 GPU 系統市場（CY）", "US$B", {c_(y): f"=({c_(y)}{R['uo']}*{c_(y)}{R['vo']}+{c_(y)}{R['un']}*{c_(y)}{R['vn']})" for y in CYS}, fmt="#,##0.0",
        note="顆數 × 每顆系統價值；含 NVIDIA 與 AMD，不含 TPU、Trainium 與其他 ASIC")
    row("mfy", "商用 GPU 系統市場（FY＝½ 前一 CY＋½ 當年 CY）", "US$B", {c_(y): f"=0.5*{c_(y - 1)}{R['mcy']}+0.5*{c_(y)}{R['mcy']}" for y in F3}, fmt="#,##0.0",
        note="會計年度 7 月～6 月：各取前後日曆年一半（Derived；未考慮年內出貨季節性）")
    row("mfy0", "FY2027 市場（部署係數 1.0；份額校準用）", "US$B", {"E": f"=0.5*D{R['mcy']}+0.5*(uo_27*v_gb+un_27*v_vr)"}, fmt="#,##0.0",
        note="份額以基準市場校準，Sens 的部署係數才會改變 FY2027 營收")
    row("sh0", "美超微有效份額（2H FY26 校準）＝ 2026-01～06 AI 營收 ÷ ½ CY2026 市場", "%", {"D": f"=a26_aih2/(0.5*D{R['mcy']})"}, fmt="0.00%",
        note="Derived：CY2026 上下半年各半（Rubin 下半年放量，上半年市場偏小，此份額偏低估）；只作對照")
    row("g27", "FY2027 營收（公司指引；A 上緣、B 中點、C 下緣）", "US$B", {"E": f"={p('g27')}"}, fmt="#,##0.0", note="Interested-party：公司 2026-11-03 指引 720～820 億美元（原 2026-08-11 為 650～720 億）")
    row("rel", "份額相對值（÷ FY2027）", "x", {"E": "=1", **{c_(y): f"={p('rel', y)}*{P['rel_mult']}" for y in (2028, 2029)}}, fmt="0.00",
        note="A 最有利、C 最不利（Assumed）；治理事件案例另見 Sens")
    row("sh", "美超微有效份額（FY2027 以指引校準）", "%", {"E": f"=(E{R['g27']}-E{{nai}})/E{R['mfy0']}", **{c_(y): f"=$E{r}*{c_(y)}{R['rel']}" for y in (2028, 2029)}}, fmt="0.00%",
        note="FY2027：（指引營收 − 非 AI 營收）÷ 基準市場；4Q FY26 逾 600 億美元新訂單使份額高於 2H FY26")
    row("shr", "FY2027 份額 ÷ 2H FY26 份額", "x", {"E": f"=E{R['sh']}/D{R['sh0']}"}, fmt="0.00", note="Derived：指引隱含的份額上升幅度")
    row("ai", "AI GPU 平台營收", "US$B", {"C": "=a25_ai", "D": "=a26_ai", **{c_(y): f"={c_(y)}{R['mfy']}*{c_(y)}{R['sh']}" for y in F3}}, fmt="#,##0.00")
    sec_("B. 非 AI 與合計")
    row("nai", "非 AI（企業、儲存、邊緣等）", "US$B", {"C": "=a25_nai", "D": "=a26_nai", **{c_(y): f"={c_(y - 1)}{r}*(1+{p('gn', y)})" for y in F3}}, fmt="#,##0.00")
    ws[f"E{R['sh']}"] = ws[f"E{R['sh']}"].value.replace("{nai}", str(R["nai"]))
    row("rev", "營收", "US$B", {c_(y): f"={c_(y)}{R['ai']}+{c_(y)}{R['nai']}" for y in FYS}, fmt="#,##0.00")
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("aish", "AI GPU 平台 ÷ 營收", "%", {c_(y): f"={c_(y)}{R['ai']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("ugpu", "隱含美超微商用 GPU 出貨（AI 營收 ÷ 每顆平均系統價值）", "百萬顆",
        {c_(y): f"={c_(y)}{R['ai']}/((0.5*({c_(y - 1)}{R['uo']}*{c_(y - 1)}{R['vo']}+{c_(y - 1)}{R['un']}*{c_(y - 1)}{R['vn']})+0.5*({c_(y)}{R['uo']}*{c_(y)}{R['vo']}+{c_(y)}{R['un']}*{c_(y)}{R['vn']}))"
                                    f"/(0.5*({c_(y - 1)}{R['uo']}+{c_(y - 1)}{R['un']})+0.5*({c_(y)}{R['uo']}+{c_(y)}{R['un']})))" for y in F3}, fmt="0.00",
        note="對照用：份額 × 會計年度商用 GPU 顆數")
    sec_("C. 損益")
    row("gai0", "AI 毛利率 FY2026 校準值 ＝（毛利 − 非 AI 毛利率 × 非 AI 營收）÷ AI 營收", "%", {"D": f"=(a26_gp-{P['gmn']}*a26_nai)/a26_ai"}, fmt="0.00%",
        note="Derived：FY2026 含 4Q 一次性利益（遞延合約、關稅、存貨跌價回轉），校準值偏高")
    row("gmai", "AI 毛利率", "%", {"D": f"=D{R['gai0']}", "E": f"=$D{R['gai0']}+{p('d27')}", **{c_(y): f"=$D{R['gai0']}+{p('d28')}" for y in (2028, 2029)}}, fmt="0.00%",
        note="A 最高、C 最低（In 頁 C 節）")
    row("gp", "毛利", "US$B", {"C": "=a25_gp", "D": "=a26_gp", **{c_(y): f"={c_(y)}{R['gmai']}*{c_(y)}{R['ai']}+{P['gmn']}*{c_(y)}{R['nai']}" for y in F3}}, fmt="#,##0.00")
    row("gm", "毛利率", "%", {c_(y): f"={c_(y)}{R['gp']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.00%")
    row("opb", "營業費用（基準路徑）", "US$B", {"C": "=a25_opx", "D": "=a26_opx", **{c_(y): f"={c_(y - 1)}{r}*(1+MAX(0,{P['phi']}*{c_(y)}{R['g']}))" for y in F3}}, fmt="#,##0.00")
    row("opx", "營業費用", "US$B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", **{c_(y): f"={c_(y)}{R['opb']}+{P['opxa']}*{c_(y)}{R['rev']}" for y in F3}}, fmt="#,##0.00")
    row("oi", "營業利益", "US$B", {c_(y): f"={c_(y)}{R['gp']}-{c_(y)}{R['opx']}" for y in FYS}, fmt="#,##0.00")
    row("oim", "營業利益率", "%", {c_(y): f"={c_(y)}{R['oi']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.00%")
    row("nio", "淨利（特別股股利前）÷ 營業利益", "%", {"C": "=a25_ni/a25_oi", "D": "=a26_ni/a26_oi", **{c_(y): f"={P['nio']}" for y in F3}}, fmt="0.0%")
    row("ni0", "淨利（特別股股利前）", "US$B", {"C": "=a25_ni", "D": "=a26_ni", **{c_(y): f"={c_(y)}{R['oi']}*{c_(y)}{R['nio']}" for y in F3}}, fmt="#,##0.00")
    row("pd", "特別股股利", "US$B", {"C": "=0", "D": "=0", **{c_(y): "=pref_div" for y in F3}}, fmt="#,##0.000",
        note="FY2026 只有約 3 週，已含在實績；2029-06 強制轉換後為 0")
    row("ni", "歸屬母公司淨利", "US$B", {c_(y): f"={c_(y)}{R['ni0']}-{c_(y)}{R['pd']}" for y in FYS}, fmt="#,##0.00",
        note="扣除特別股股利後、歸屬普通股（GAAP）")
    row("nim0", "淨利率（特別股股利前）", "%", {c_(y): f"={c_(y)}{R['ni0']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.00%", note="收斂法的起點（FY2030 起特別股已轉換）")
    row("sh_", "股數（GAAP 稀釋）", "十億股", {"C": "=a25_sh", "D": "=a26_sh", "E": "=sh27", **{c_(y): f"={c_(y - 1)}{r}*(1+sh_g)" for y in (2028, 2029)}}, fmt="0.0000")
    row("eps", "EPS（GAAP 稀釋）", "US$", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['sh_']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("nim", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.00%")
    # FY2030–32：收斂法（上游 CY2029 為最後一年，FY2030 需 CY2030；非驅動因子推導）
    for y in FADE:
        c, pc, k = FC[y], FC[y - 1], y - 2029
        ws[f"{c}{R['g']}"] = f"=$G{R['g']}+({p('fg')}-$G{R['g']})*{k}/3"
        ws[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
        ws[f"{c}{R['nim']}"] = f"=$G{R['nim0']}+({p('fm')}-$G{R['nim0']})*{k}/3"
        ws[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
        ws[f"{c}{R['sh_']}"] = f"={pc}{R['sh_']}*(1+sh_g)" + ("+conv_sh" if y == 2030 else "")
        ws[f"{c}{R['eps']}"] = f"={c}{R['ni']}/{c}{R['sh_']}"
        ws[f"{c}{R['epsg']}"] = f"={c}{R['eps']}/{pc}{R['eps']}-1"
        for key, fm in (("g", "0.0%"), ("rev", "#,##0.00"), ("nim", "0.00%"), ("ni", "#,##0.00"), ("sh_", "0.0000"), ("eps", "0.00"), ("epsg", "0.0%")):
            ws[f"{c}{R[key]}"].number_format = fm; ws[f"{c}{R[key]}"].fill = FADEFILL
    ws.cell(R["rev"], 11, "FY2030 起為收斂法：營收年增由 FY2029 線性收斂到 fade_g、淨利率由 FY2029 特別股股利前淨利率收斂到 fade_m（FY2032 達到），非驅動因子推導；FY2030 股數加特別股轉換股數")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 美超微損益（十億美元；EPS 美元、GAAP 稀釋；歸屬母公司＝扣除特別股股利後）。FY2025A、FY2026A 實績（FY2026 為錨點），FY2027 起為模型。會計年度 7 月～6 月：FY(t) 市場＝½ CY(t−1)＋½ CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A（順利）：顆數係數 1.0；FY2027 營收＝指引上緣；份額維持；AI 毛利率 +1.5pt",
         "B": "情境 B（平平）：CY2028 顆數 ×0.8、CY2029 ×0.85；FY2027 營收＝指引中點；份額 ×0.95／×0.92；AI 毛利率 +0.5pt／+0.3pt",
         "C": "情境 C（不如預期）：CY2028–29 顆數 ×0.6；FY2027 營收＝指引下緣；份額 ×0.85／×0.80；AI 毛利率 −0.5pt／−1.5pt"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026A（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
sr = 4; SUMR = {}
SUMK = (("rev", "營收", "US$B", "#,##0.00"), ("g", "  營收年增", "%", "0.0%"), ("ai", "  AI GPU 平台營收", "US$B", "#,##0.00"),
        ("nai", "  非 AI（企業、儲存、邊緣等）", "US$B", "#,##0.00"), ("aish", "  AI GPU 平台 ÷ 營收", "%", "0.0%"),
        ("gm", "毛利率", "%", "0.00%"), ("oim", "營業利益率", "%", "0.00%"), ("ni", "歸屬母公司淨利", "US$B", "#,##0.00"), ("nim", "  淨利率", "%", "0.00%"),
        ("eps", "EPS", "US$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%"))
LONGK = ("rev", "g", "ni", "nim", "eps", "epsg")
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in SUMK:
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS + (FADE if key in LONGK else ()):
            if key in ("g", "epsg") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
            if y in FADE: PL[f"{FC[y]}{sr}"].fill = FADEFILL
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
IN.cell(INROW["fade_mA"], 5, f"=PL!G{PB['A']['nim0']}")
IN.cell(INROW["fade_mB"], 5, f"=(PL!G{PB['B']['nim0']}+nim10)/2")
for k in ("fade_mA", "fade_mB"): IN.cell(INROW[k], 5).font = Font(color="000000")
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型 − 實績").font = BOLD; sr += 1
CHK = []
for lab, f in (("  三情境：營收差（US$B）", "+".join(f"ABS(D{PB[s]['rev']}-a26_rev)" for s in SC)),
               ("  三情境：EPS 差（US$）", "+".join(f"ABS(D{PB[s]['eps']}-a26_eps)" for s in SC)),
               ("  情境 A：毛利差（US$B；以 FY2027 公式套用 FY2026）", f"D{PB['A']['gai0']}*D{PB['A']['ai']}+gm_nai*D{PB['A']['nai']}-a26_gp")):
    PL.cell(sr, 1, lab); PL[f"D{sr}"] = "=" + f; PL[f"D{sr}"].number_format = "0.0000"; CHK.append(sr); sr += 1
assert sr < 60, sr
PL.cell(2, 1, "灰底＝FY2030–FY2032 收斂法（上游只到 CY2029，FY2030 需 CY2030）：營收年增線性收斂到 fade_g、淨利率線性收斂到 fade_m（In 頁 H 節），非驅動因子推導")
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 12; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 13

# ---------------- Units ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 商用 GPU 系統市場、美超微份額與隱含出貨（情境 A）。前五列為日曆年（CY），其後為美超微會計年度（FY，7 月～6 月）。顆數讀台積電模型、每顆系統價值讀橋接表（唯讀）"
A_ = PB["A"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "2026", "2027", "2028", "2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("uo", "CY：GB300 世代商用 GPU 顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['uo']}", "0.00", "tsmc 成品"),
        ("un", "CY：Rubin 世代商用 GPU 顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['un']}", "0.00", "tsmc 成品"),
        ("ut", "CY：全部 AI 加速器顆數（對照）", "百萬顆", lambda y: f"=ut_{y % 100}", "0.00", "含 TPU、Trainium、其他 ASIC"),
        ("vn", "CY：每顆系統價值（Rubin 世代）", "千美元/顆", lambda y: f"=PL!{FC[y]}{A_['vn']}", "0.0", "GB300 世代為 In 頁 v_gb"),
        ("mcy", "CY：商用 GPU 系統市場", "US$B", lambda y: f"=PL!{FC[y]}{A_['mcy']}", "#,##0.0", "顆數 × 每顆系統價值"),
        ("mfy", "FY：商用 GPU 系統市場（½ 前一 CY＋½ 當年 CY）", "US$B", lambda y: f"=PL!{FC[y]}{A_['mfy']}" if y >= 2027 else None, "#,##0.0", "FY2026 需 CY2025（上游無），不列"),
        ("sh", "FY：美超微有效份額", "%", lambda y: f"=PL!{FC[y]}{A_['sh']}" if y >= 2027 else f"=PL!D{A_['sh0']}", "0.00%", "FY2026 欄＝2H FY26 校準值；FY2027 起以指引校準"),
        ("ai", "FY：美超微 AI GPU 平台營收", "US$B", lambda y: f"=PL!{FC[y]}{A_['ai']}", "#,##0.0", ""),
        ("ugpu", "FY：隱含美超微商用 GPU 出貨", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['ugpu']}" if y >= 2027 else None, "0.00", "AI 營收 ÷ 每顆平均系統價值"),
        ("gmai", "FY：AI 毛利率", "%", lambda y: f"=PL!{FC[y]}{A_['gmai']}", "0.00%", "FY2026 為校準值"),
        ("gm", "FY：毛利率", "%", lambda y: f"=PL!{FC[y]}{A_['gm']}", "0.00%", "")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in CYS:
        v = f(y)
        if v is None: continue
        UN[f"{cyc(y)}{ur}"] = v; UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); UR[key] = ur; ur += 1
ur += 1
UN.cell(ur, 1, "讀法：美超微賣的是含 GPU 的整機與整櫃（品牌商，買賣模式），營收隨商用 GPU 顆數 × 每顆系統價值放大；份額在 2H FY26 約數個百分點，FY2027 指引隱含份額明顯上升（大量新訂單）。FY2028 起份額相對值是 Assumed，大客戶（FY2026 占 28%）轉單是最大下檔。").font = BOLD
UN_NOTE = ur
UN.column_dimensions["A"].width = 60; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 14
UN.column_dimensions["G"].width = 60

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（美元，GAAP EPS）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，會計年度年底落後 P/E）只作參考"
V["A2"] = "循環位置：美超微不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照廣達、緯穎）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
O = 51
V.cell(O, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, O + 1, ["項目", "A", "B", "C", "說明"])
V.cell(O + 2, 1, "現價（2026-10-23 收盤）"); V[f"B{O + 2}"] = "=px"; V[f"B{O + 2}"].number_format = "0.00"
for k, y, lab in ((3, 2026, "現價 ÷ FY2026 EPS（實績）"), (4, 2027, "現價 ÷ FY2027 EPS（模型）"), (5, 2028, "現價 ÷ FY2028 EPS（模型）"), (6, 2029, "現價 ÷ FY2029 EPS（模型）")):
    V.cell(O + k, 1, lab)
    for j, s in enumerate(SC):
        c = "BCD"[j]; V[f"{c}{O + k}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{O + k}"].number_format = "0.0"
V.cell(O + 7, 1, "未來 12 月 EPS（模型 B）≈ 8／12 × FY2027＋4／12 × FY2028"); V[f"B{O + 7}"] = f"=8/12*PL!E{SUMR[('B', 'eps')]}+4/12*PL!F{SUMR[('B', 'eps')]}"; V[f"B{O + 7}"].number_format = "0.00"
V.cell(O + 8, 1, "  現價 ÷ 未來 12 月 EPS（模型 B）"); V[f"B{O + 8}"] = f"=px/B{O + 7}"; V[f"B{O + 8}"].number_format = "0.0"
V.cell(O + 9, 1, "近 12 月 P/E（stockanalysis，10-23）"); V[f"B{O + 9}"] = "=pe_ttm"; V[f"B{O + 9}"].number_format = "0.0"
V.cell(O + 10, 1, "前瞻 P/E（stockanalysis，10-23；non-GAAP）"); V[f"B{O + 10}"] = "=fpe"; V[f"B{O + 10}"].number_format = "0.0"
V[f"E{O + 3}"] = "FY2026 三情境相同（實績）"; V[f"E{O + 7}"] = "2026-11～2027-10：FY2027（7 月～6 月）8 個月＋FY2028 4 個月"
CY0 = O + 12
V.cell(CY0, 1, "2. 循環位置（本模型營收最高的會計年度）").font = BOLD
hdr(V, CY0 + 1, ["情境", "營收最高年（FY2026–FY2029）", "FY2027 判定", "規則倍數（P/E）", "見頂年是否未知", "說明"])
CYR = {}
for i, s in enumerate(SC, CY0 + 2):
    rr = f"PL!D{PB[s]['rev']}:G{PB[s]['rev']}"
    V.cell(i, 1, s); CYR[s] = i
    V.cell(i, 2, f"=INDEX({{2026,2027,2028,2029}},MATCH(MAX({rr}),{rr},0))")
    V.cell(i, 3, '=IF(B{0}=2029,"中段",'.format(i) + CP.excel_rule("2027", f"B{i}") + ")")
    V.cell(i, 4, f'=IF(C{i}="見頂",pe_peak,IF(C{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 5, f'=IF(B{i}=2029,"是：營收到觀測期末仍在增加","否")')
    V.cell(i, 6, "tools/cycle_position.py excel_rule；營收最高年取代設備市場見頂年" if s == "A" else "同上")
R30 = CY0 + 6
V.cell(R30, 1, "3. 規則值隱含股價（美元；約 FY2027 年底＝2027-06，未折現）＝ FY2027 EPS × 規則倍數").font = BOLD
hdr(V, R30 + 1, ["情境", "FY2027 EPS", "（同左）", "規則倍數", "隱含股價", "相對現價", "現價 ÷ FY2027 EPS"])
VAL = {}
for j, s in enumerate(SC):
    i = R30 + 2 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 3, f"=B{i}").number_format = "0.00"
    V.cell(i, 4, f"=D{CYR[s]}").number_format = "0.0"
    V.cell(i, 5, f"=C{i}*D{i}").number_format = "#,##0.00"
    V.cell(i, 6, f"=E{i}/px-1").number_format = "0%"
    V.cell(i, 7, f"=px/C{i}").number_format = "0.0"
    for c in range(1, 8): V.cell(i, c).fill = KEY
    VAL[s] = i
W0 = R30 + 6
V.cell(W0, 1, "4. What must be true：FY2027 年底價 ＝ 現價 ×（1＋報酬）所需的倍數（FY2028 EPS；2027-06 的價格看下一會計年度 EPS）").font = BOLD
hdr(V, W0 + 1, ["FY2027 年底價 ÷ 現價", "所需倍數（÷ FY2028 EPS，A）", "所需倍數（÷ FY2028 EPS，B）", "所需倍數（÷ FY2028 EPS，C）", "對照：十年中位數"])
vr = W0 + 2
for x in (1.0, 1.1, 1.21, 1.5, 2.0):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    for j, s in enumerate(SC):
        V.cell(vr, 2 + j, f"=px*A{vr}/PL!F{SUMR[(s, 'eps')]}").number_format = "0.0"
    V.cell(vr, 5, "=pe_mid").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數低於歷史中位數時，現價相對便宜；美超微的低倍數同時反映治理、會計與營運資金風險（見 README）。本版未做先驗，無 Andy 合理 P/E 對照。")
WMT_NOTE = vr
# 十年 P/E：會計年度年底（6 月底）落後 P/E；Macrotrends（GAAP TTM EPS）與 fullratio（年度值與 6 月季度 TTM 相同）
PE_MT = [16.35, 18.40, None, None, 26.05, 16.75, 7.63, 21.77, 41.13, 29.52]
PE_FR = [16.57, 17.61, 26.28, 13.82, 16.7, 15.99, 7.34, 20.6, 39.58, 27.69]
T10 = vr + 2
V.cell(T10, 1, "5. 十年 P/E（FY2016–FY2025 會計年度年底＝6 月底；兩來源平均，分位排名用；Macrotrends 2018–2019 因延遲申報缺值）").font = BOLD
hdr(V, T10 + 1, ["年底", "兩來源平均 P/E", "Macrotrends", "fullratio"])
RB0 = T10 + 16
for i, y in enumerate(range(2016, 2026)):
    rr = T10 + 2 + i
    V.cell(rr, 1, str(y))
    V.cell(rr, 2, f"=AVERAGE(C{rr}:D{rr})").number_format = "0.00"
    V.cell(rr, 3, f'=IF(ISNUMBER(B{RB0 + 5 + i}),B{RB0 + 5 + i},"")').number_format = "0.00"
    V.cell(rr, 4, f"=C{RB0 + 5 + i}").number_format = "0.00"
PE10 = f"$B${T10 + 2}:$B${T10 + 11}"
V.cell(T10 + 12, 1, "年數"); V[f"B{T10 + 12}"] = f"=COUNT({PE10})"
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "Macrotrends（6 月底 P/E）", "vals": PE_MT, "tag": "Interested-party（資料商；GAAP 近 12 月 EPS；2018–2019 延遲申報期間 EPS 記 0，缺值）"},
    {"label": "fullratio（年度 P/E）", "vals": PE_FR, "tag": "Interested-party（資料商；年度值與 6 月季度相同，推定為會計年度年底；EPS 口徑可能為調整後）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位，兩個來源分別取分位再平均。美超微會計年度 6 月底結束，取會計年度年底。兩來源 2020 年相差約 9 個倍數點（Macrotrends 的 FY2020 EPS 受重編影響）；FY2024 因股價高峰達約 40x。")
assert V.cell(RB0 + 5, 1).value == "2016" and V.cell(RB0 + 14, 1).value == "2025", (V.cell(RB0 + 5, 1).value, V.cell(RB0 + 14, 1).value)
rule_rows = [rw for rw in range(RB0, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")

# 0. 目標價（Andy 2026-10-25 11:28／11:53：DCF 與 Comparable 各半；列標籤依工作單逐字；B／C／D＝A 順利／B 平平／C 不如預期）
YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
V.cell(4, 1, "0. 目標價").font = BOLD
T = {}
def tset(r, key, label, fs, fmt, note=""):
    V.cell(r, 1, label); T[key] = r
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{r}"] = f; V[f"{c}{r}"].number_format = fmt
    if note: V[f"E{r}"] = note
# v0.3：0.1 節改為標準版 DCF（FCFF＋WACC）。原 0.1 節占第 5–24 列，新摘要 18 列寫在第 5–22 列，0.2、0.3 節列號不變
DCFW = wb.create_sheet("DCF")
DCFW["A1"] = "DCF — 標準版（v0.3）：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例；FY2027 只計 2027-01～06 半年）；終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py"
DCFW["A2"] = ("EBIT（GAAP，含股權報酬費用）：FY2027–FY2029 取 PL 推導區塊「營業利益」；FY2030–FY2032（收斂期只有淨利；特別股已轉換）由淨利 ÷ nio_f 還原。"
              "NOPAT＝EBIT ×（1 − tax_w，近三年有效稅率）。折舊、資本支出、營運資金：近三年（FY2024–FY2026）實績占營收比例（In 頁 H2 節）。強制可轉換特別股以轉換股數計入股數，轉換前股利列為類債扣除。")
FY_X = list(range(2027, 2033))
_ebit = lambda s_, y: f"PL!{FC[y]}{PB[s_]['oi']}" if y <= 2029 else f"=PL!{FC[y]}{PB[s_]['ni']}/nio_f"
DPATHS = [dict(name=s, label=f"情境 {s}（{ {'A': '順利', 'B': '平平', 'C': '不如預期'}[s] }）", rev={y: f"PL!{FC[y]}{PB[s]['rev']}" for y in FY_X}, rev0=f"PL!{FC[2026]}{PB[s]['rev']}",
              ebit={y: _ebit(s, y) for y in FY_X}, tax={y: "tax_w" for y in FY_X}, ebit_note="FY2030 起＝淨利 ÷ nio_f") for s in SC]
RD_ = FD.write_detail(DCFW, 4, DPATHS, FY_X, 0.5, title="三情境 FCFF 明細（US$B；FY2027＝2026-07～2027-06，評價日後只計 2027-01～06）", stub=True,
                      note="淨現金滾到評價日（v0.3.2）：資產負債表日 2026-09-30 → 評價日 2026-12-31；stub FCFF＝各情境 FY2027（2026-07～2027-06，含 stub 期）FCFF × stub_frac（模組預設），再扣已宣告股利 stub_div")
RD_["end"] = FD.write_checks(DCFW, RD_["end"], nc_bs="6.897+0-(1.963+6.683)-0.58805-mi_v",
                             nc_note="1Q FY27 資產負債表（US$B）：現金 6.897 − 短期借款 1.963 − 長期債務 6.683 − 租賃 0.588 − 非控制權益與特別股轉換前股利（mi_v）") + 1
DCFW.column_dimensions["A"].width = 58; DCFW.column_dimensions["B"].width = 40
for _c in "CDEFGH": DCFW.column_dimensions[_c].width = 13
assert FD.summary_rows(stub=True) <= 20   # 原 0.1 節列數；0.2 節起列號不變
SDCF = FD.write_summary(wb, V, 5, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt="#,##0.00",
                        title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；明細見 DCF 頁；B／C／D＝情境 A／B／C）", beta_note="自算產業 β 重新加槓桿")
T.update(dcfv=SDCF["dcfv"], dcf12=SDCF["v12"], dcftp=SDCF["dcf"], tvsh=SDCF["tvsh"], mult=SDCF["mult"], wacc=SDCF["wacc"])
V[f"E{SDCF['dcf']}"] = "Σ 權重 × 情境價值（已推到 12 個月後）；權重見 In 頁"
EPSR, CFR = {}, {}
_E = lambda s_, y: f"PL!{FC[y]}{SUMR[(s_, 'eps')]}"
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = [("Dell Technologies", "DELL", 16.22, "2026-10-23"), ("Hewlett Packard Enterprise", "HPE", 12.19, "2026-10-23"),
         ("緯穎（Wiwynn）", "6669.TW", 11.73, "2026-10-21"), ("廣達（Quanta）", "2382.TW", 11.52, "2026-10-21"), ("緯創（Wistron）", "3231.TW", 9.21, "2026-10-19")]
for i, (nm, tk, pe, dt) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左）"); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", "AI 伺服器品牌與 ODM 同業：Dell、HPE、緯穎、廣達、緯創（不含 A 股）")
tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=" + "+".join(f"tp_w{s_}*(8/12*{_E(s_, 2028)}+4/12*{_E(s_, 2029)})" for s_ in SC), None, None], "0.00",
     "會計年度 6 月底結束：約 2027-10 時的未來 12 月（2027-11～2028-10）≈ 8／12 × FY2028＋4／12 × FY2029")
tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], "0.00", "同業中位數 × 前瞻 EPS")
V.cell(37, 1, "0.3 結論").font = BOLD
hdr(V, 38, ["項目", "數值", "", "", "說明"])
tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B35", None, None], "0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁（0.5）")
tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], "0.00")
tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], "0.00", "stockanalysis（S&P Global）：18 位分析師平均 43.27（中位數 45、最低 15、最高 60），2026-10-22；tickflow 25 位 41.44")
tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-23 收盤）")
tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS（GAAP）÷ 共識前瞻 EPS（non-GAAP，＝價格 ÷ 前瞻 P/E）− 1；期間約晚一年、口徑不同（GAAP 較低）")
tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
tset(46, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D{T["dcf12"]},"0.0")&"～"&TEXT(B{T["dcf12"]},"0.0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [(lambda e: f"=IF(px/{e}<=MIN({PE10}),0,IF(px/{e}>=MAX({PE10}),1,PERCENTRANK({PE10},px/{e})))")(_E(s_, 2028)) for s_ in SC], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 FY2016–FY2025 年底 P/E（第 5 節兩來源平均）的分位；低於歷史最小值記 0%、高於最大值記 100%；只作檢查，不進目標價")
V.cell(48, 1, "目標價方法")
V["B48"] = ("DCF（v0.3 標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加（GAAP 營業利益，股權報酬視為費用），以 WACC 折現到 2026 年底（期中慣例；FY2027 只計半年），終值＝NOPAT(FY2033) ×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, 美國 10 年期公債殖利率)、RONIC＝WACC；"
            "EV 減淨負債與特別股轉換前股利 ÷ 股數（含特別股轉換），以股權成本推到 12 個月後（DCF 頁、tools/fcff_dcf.py）。β 為自算產業 β（美國伺服器／EMS 同業 8 家無槓桿 β 中位數，Blume 調整）依 Supermicro D/E 重新加槓桿；治理與會計風險不另加溢酬。營運資金約占營收三分之一，營收高成長年度 FCFF 為負。"
            "Comparable：Dell、HPE、緯穎、廣達、緯創前瞻 P/E 中位數（不含 A 股）；同業倍數未反映美超微的治理與會計折價（自身前瞻 P/E 約 7x），偏誤向上。"
            "FY2030 起為收斂法；FY2030 股數加特別股轉換股數。治理風險不改折現率，以 Sens 案例呈現。"
            "情境對應：A＝順利（原始顆數）、B＝平平（顆數 ×0.8／0.85）、C＝不如預期（顆數 ×0.6）")
for r_ in (T["dcftp"], 35, 39):
    V[f"B{r_}"].fill = KEY
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；FY2027 份額在每個案例內依指引重新校準，部署係數與每顆價值除外）：對 FY2027、FY2028 營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("AI 部署 ×0.8", "CY2027 起商用 GPU 顆數 ×0.8", dict(mf="mf_lo")), ("AI 部署 ×1.2", "×1.2", dict(mf="mf_hi")),
         ("FY2027 營收＝指引下緣", "720 億美元（情境 A 為上緣 820 億）", dict(g27="g27_lo")),
         ("份額基準（B 值）", "份額相對值 FY2028 0.95、FY2029 0.92", dict(rel="rel_{y}")),
         ("份額流失（C 值）", "份額相對值 0.85／0.80", dict(rel="rel_{y}_lo")),
         ("治理事件：大客戶流失", "FY2027 指引下緣、FY2028 起份額 ×0.8（出口管制、會計或內控事件使客戶轉向 Dell、ODM）", dict(g27="g27_lo", rel_mult="0.8")),
         ("每顆系統價值 低", "2028–29 Rubin 世代 ×0.85", dict(vpg="vpg_{y}_lo")), ("每顆系統價值 高", "×1.2", dict(vpg="vpg_{y}_hi")),
         ("AI 毛利率 基準（B 值）", "FY2027 +0.5pt、FY2028 起 +0.3pt", dict(d27="dgm_27", d28="dgm_28")),
         ("AI 毛利率 低（C 值）", "FY2027 −0.5pt、FY2028 起 −1.5pt", dict(d27="dgm_27_lo", d28="dgm_28_lo")),
         ("非 AI 毛利率 13%", "基準 15%", dict(gmn="gm_nai_lo")), ("非 AI 毛利率 18%", "基準 15%", dict(gmn="gm_nai_hi")),
         ("非 AI 成長 低", "FY2027 0%、FY2028–29 −5%", dict(gn="g_nai{y}_lo")), ("非 AI 成長 高", "FY2027 +30%、FY2028 +20%、FY2029 +15%", dict(gn="g_nai{y}_hi")),
         ("營業費用率 +0.2pt", "營業費用額外 +0.2% × 營收", dict(opxa="opx_adj_hi")), ("營業費用彈性 0.6", "基準 0.4", dict(phi="phi_hi")),
         ("淨利 ÷ 營業利益 0.68", "利息（營運資金借款）、合規與法律費用上升；1Q FY27 指引隱含值", dict(nio="nio_f_lo")),
         ("淨利 ÷ 營業利益 0.80", "FY2026 實測約 0.805", dict(nio="nio_f_hi"))]
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.0"), ("D", "gm", 2027, "0.00%"), ("E", "eps", 2027, "0.00"),
                           ("G", "rev", 2028, "#,##0.0"), ("H", "gm", 2028, "0.00%"), ("I", "eps", 2028, "0.00")):
        SE[f"{c}{i}"] = f"={pre}{FC[y]}{R[key]}"; SE[f"{c}{i}"].number_format = fmt
    SE[f"F{i}"] = f"=E{i}/E$4-1"; SE[f"J{i}"] = f"=I{i}/I$4-1"
    for c in "FJ": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
base_row = 40; SB = {}
for nm, setting, ov in cases:
    SB[nm] = engine(SE, base_row, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    base_row = SB[nm]["end"] + 2
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm], "")
nn = 5 + len(cases)
put_case(nn, "情境 B", "CY2028 顆數 ×0.8、CY2029 ×0.85；指引中點；份額、毛利率 B 值", PB["B"], "PL!")
put_case(nn + 1, "情境 C", "CY2028–29 顆數 ×0.6；指引下緣；份額、毛利率 C 值", PB["C"], "PL!")
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：FY2027 營收以指引校準，所以 FY2027 主要受指引區間、毛利率與淨利係數影響；FY2028 起顆數情境、份額與每顆價值才發揮作用。美超微淨利率只有約 5%–6%，AI 毛利率每 1 個百分點約影響 EPS 10%。治理事件以份額與淨利係數案例呈現，不改情境結構。")
SENS_NOTE = nn + 3
# v0.3：目標價參數敏感度改為標準版 DCF 案例（tools/fcff_dcf.py write_sens；舊 tp_r／tp_roe 案例刪除），表格放在案例區塊之後（原位置只有 13 列空間不足）
TPS0 = SE.max_row + 3
SE.cell(nn + 4, 1, f"v0.3：目標價參數表（標準版 DCF）移到本頁第 {TPS0} 列")
_tvs = FD.write_sens(SE, TPS0, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref=f"Val!$B${T['comptp']}", base_ref=f"Val!$B${T['tpw']}", fmt="#,##0.00", beta_du="beta_du",
                     title="目標價參數敏感度（標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）")
TPCASES = FD.SENS_CASES
TPS_LAST = _tvs[1]
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 48
for c in "CDEFGHIJ": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司指引、共識與份額口徑。只作事後對照，不為對上數字而調參數。共識 EPS 為 non-GAAP；本模型為 GAAP，另列「模型＋股票報酬調整」對照"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.00", pt=False, gap=True):
    global rr_
    RC.cell(rr_, 1, label)
    if ref is not None: RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        if ref is not None and gap:
            L = chr(64 + cv)
            RC.cell(rr_, cd, f"={L}{rr_}-B{rr_}" if pt else f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y: (lambda s: f"=PL!{FC[y]}{PB[s][key]}")
RC.cell(rr_, 1, "A. 共識（校準未使用）").font = BOLD; rr_ += 1
rec("FY2027 營收 vs 共識", "=c_rev27", pl("rev", 2027), "stockanalysis／Yahoo（15 位）", "本模型 FY2027 營收＝公司指引（A 上緣、B 中點、C 下緣）")
rec("FY2028 營收 vs 共識", "=c_rev28", pl("rev", 2028), "Yahoo（15 位）", "")
rec("FY2027 EPS（GAAP）vs 共識（non-GAAP）", "=c_eps27", pl("eps", 2027), "stockanalysis／Yahoo", "口徑不同：GAAP 較 non-GAAP 低約股票報酬")
rec("FY2027 EPS（模型＋股票報酬調整）vs 共識", "=c_eps27", lambda s: f"=PL!E{PB[s]['eps']}+sbc_ps", "同上", "調整＝每季股票報酬（稅後）× 4 ÷ 股數（In 頁 sbc_ps）")
rec("FY2028 EPS（模型＋股票報酬調整）vs 共識", "=c_eps28", lambda s: f"=PL!F{PB[s]['eps']}+sbc_ps", "Yahoo（17 位）", "")
RC.cell(rr_, 1, "B. 公司指引與初步數字").font = BOLD; rr_ += 1
rec("FY2027 營收 vs 指引中點（Interested-party）", "=g27", pl("rev", 2027), "公司 2026-11-03", "依情境取指引上緣／中點／下緣，差距即指引區間")
rec("FY2027 毛利率 vs 1Q FY27 實績毛利率（百分點）", "=q1_gm", pl("gm", 2027), "公司 2026-11-03", "1Q 實績 12.7%（GAAP），含 4Q FY26 遞延合約的組合利益", fmt="0.00%", pt=True)
rec("1Q FY27 實績營收 × 4 vs FY2027 營收", "=q1_rev*4", pl("rev", 2027), "公司 2026-11-03", "1Q 年化 798 億，落在新全年指引 720～820 億之內")
RC.cell(rr_, 1, "C. 份額與市場").font = BOLD; rr_ += 1
rec("2H FY26 有效份額（商用 GPU 系統市場）", None, lambda s: f"=PL!D{PB[s]['sh0']}", "Derived", "", fmt="0.00%")
rec("FY2027 有效份額（指引校準）", None, pl("sh", 2027), "Derived", "", fmt="0.00%")
rec("FY2027 份額 ÷ 2H FY26 份額", None, pl("shr", 2027), "Derived", "大於 1：指引隱含份額上升（4Q FY26 逾 600 億美元新訂單）", fmt="0.00")
rec("FY2027 隱含商用 GPU 出貨（百萬顆）", None, pl("ugpu", 2027), "Derived", "", fmt="0.00")
RC.cell(rr_, 1, "D. 校準").font = BOLD; rr_ += 1
rec("AI 毛利率 FY2026 校準值（Derived）", None, lambda s: f"=PL!D{PB[s]['gai0']}", "FY2026 實績", "含 4Q 一次性利益；不含 4Q 的 FY2026 毛利率見 Q 頁 B24", fmt="0.00%")
RC.cell(rr_, 1, "E. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以部署路徑取代「台積電顆數 × 情境係數」"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RECON_LAST = rr_ - 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文為新聞稿與財報說明稿；10-K 內容只經二手摘要，原文待人工複核；券商與資料商皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("公司：FY2026 第四季與全年新聞稿（2026-08-11；營收、毛利率、營業費用、淨利、EPS、股數、1Q FY27 與 FY2027 指引、董事會調查）", PR4Q),
        ("公司：4Q FY26 財報說明稿（AI 占比、客戶與地區結構、最大客戶 28%、毛利率說明、新訂單）", PREP),
        ("公司（摘要）：MarketBeat 2026-08-11 法說重點", MB0811),
        ("公司：2Q FY26 新聞稿（8-K；1Q、2Q FY26 數字）", PRQ2),
        ("公司：1Q FY27 初步營運更新（2026-10-20；營收 185～195 億、毛利率 12.0–12.5%、11-03 財報）", PRQ1U),
        ("公司：2026-06-11 普通股與強制可轉換特別股定價（股利率 7%、轉換率、2029-06-01 轉換）", FIN0611),
        ("公司：2026-08 獨立調查完成公告（出口管制）", INVEST),
        ("媒體：Stocktwits 2026-08-20（調查結論、2026-03 起訴、共同創辦人辭任董事）", ST0820),
        ("公司：FY2026 10-K（2026-08-31）", K10),
        ("媒體：StockTi（10-K 摘要：內控仍未有效、BDO 否定意見、單一供應商 63.1%）", STK10),
        ("媒體：Futurum（2Q FY26 平台）", FUT),
        ("資料商：stockanalysis 統計頁（現價、股數、P/E、目標價）", SA),
        ("資料商：stockanalysis 預測頁（共識營收、EPS、目標價）", SAF),
        ("資料商：stockanalysis 比率頁（ROE、年度 P/E）", SAR),
        ("資料商：stockanalysis 損益表（FY2022–FY2026）", SAI),
        ("資料商：Yahoo Finance 分析頁（FY2027–28 共識）", YF),
        ("資料商：tickflow 預測頁（目標價第二來源）", TICK),
        ("資料商：Simply Wall St（目標價歷史、ROE 預測、1Q 指引上修）", SWS),
        ("資料商：Macrotrends P/E（6 月底）", MT), ("資料商：Macrotrends 營收", MTREV), ("資料商：Macrotrends 淨利", MTNI),
        ("資料商：fullratio P/E（年度、現價）", FR),
        ("同業前瞻 P/E：stockanalysis（Dell）", P_DELL), ("同業前瞻 P/E：stockanalysis（HPE）", P_HPE), ("同業前瞻 P/E：stockanalysis（緯穎）", P_WY),
        ("同業前瞻 P/E：stockanalysis（廣達）", P_QC), ("同業前瞻 P/E：stockanalysis（緯創）", P_WS),
        ("市場層：tsmc、markets/chip_bridge 成品（本 repo，唯讀）", f"{rel(TSMC_F)}；{rel(BR_F)}"),
        ("公司：1Q FY27 新聞稿（8-K Ex.99.1，2026-11-03；營收、毛利、營業費用、淨利、EPS、股數、2Q FY27 與 FY2027 指引）", PRQ1),
        ("待補（人工）：1Q FY27 10-Q；10-K 原文（客戶占比、內控）；各季 AI GPU 占比的公司原文；AI 與非 AI 毛利率；可轉債轉換價與稀釋；特別股超額配售執行數；前瞻 P/E 歷史（Bloomberg）", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 100; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("美超微 Super Micro Computer（SMCI US）收支與估值模型 v0.1.1（AI 半導體 Project：伺服器組裝次產業，樣板＝wiwynn/、quanta/）　2026-10-25", True),
 ("版本：v0.1.1（2026-10-25：補上檔空間級距、投資建議、重點觀察（數值不變））。v0.1（初版）。美超微是 AI 伺服器品牌商（整機、整櫃、直接液冷），客戶含 neocloud、大型資料中心與企業，以買賣模式銷售含 NVIDIA、AMD GPU 的系統。商用 GPU 顆數與情境係數讀台積電模型，每顆系統價值讀橋接表（Tokenomics 整櫃價格），本模型只決定美超微的份額、毛利率與非 AI 營收。", False),
 ("模型命題：美超微 AI 營收＝商用 GPU 顆數（台積電模型，同字母情境）× 每顆系統價值（整櫃買方價 ÷ 每架 GPU）× 美超微有效份額（FY2027 以公司指引校準，FY2028 起依情境變動）；非 AI 營收以成長率推；毛利率＝AI 與非 AI 加權，AI 毛利率依情境調整。", False),
 ("核心問題：FY2027 營收指引 720～820 億美元（2026-11-03 上修；1Q 實績 199.5 億、GAAP EPS 1.75）能否轉成 GAAP EPS 約 6 美元、並守住約 11% 毛利率；現價約 33 美元、約 7 倍前瞻 P/E 的治理折價是否過深。", False), ("", False),
 ("會計年度：FY(t)＝t−1 年 7 月～t 年 6 月；FY2026 已結束（實績錨點）。上游為日曆年：FY(t) 市場＝½ CY(t−1)＋½ CY(t)。金額：十億美元；EPS：GAAP 稀釋（美元）。口徑：歸屬母公司淨利＝扣除強制可轉換特別股股利後、歸屬普通股；2029-06 特別股轉換後（FY2030 起）股利消失、股數增加。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("AI 營收：tsmc 成品的商用 GPU 顆數（Blackwell、Hopper、Rubin、Rubin Ultra、AMD）× 情境顆數係數 × 每顆系統價值（橋接表：GB300 機架 430 萬 ÷ 72、VR200 機架 840 萬 ÷ 72 美元）＝商用 GPU 系統市場；× 美超微有效份額（FY2027＝（指引營收 − 非 AI）÷ 基準市場，Derived；FY2028 起 × 份額相對值，Assumed）。", False),
 ("非 AI：FY2026 殘差 × 成長率（Assumed）。毛利率：非 AI 15%（Assumed）；AI 毛利率＝FY2026 校準值＋情境調整（Interested-party／Assumed）。營業費用彈性 phi（Derived）。淨利：營業利益 × 淨利係數（稅與利息）− 特別股股利。", False),
 ("校準：FY2026 ＝ 全年實績（公司 2026-08-11）；FY2027 營收＝公司指引（A 上緣、B 中點、C 下緣）。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，三情境權重 0.3／0.5／0.2）＋0.5 × Comparable 目標價（Dell、HPE、緯穎、廣達、緯創前瞻 P/E 中位數 × 12 個月後前瞻 EPS（8／12 × FY2028＋4／12 × FY2029）；不含 A 股）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解。十年 P/E 規則值 × FY2027 EPS 保留為參考（不是目標價）。情境對應：A＝順利、B＝平平、C＝不如預期（原為顆數情境）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 fade_g、淨利率由 FY2029 特別股股利前淨利率收斂到 fade_m，FY2032 達到；PL 頁灰底）。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司原文為新聞稿與財報說明稿；10-K 只經二手摘要；券商與資料商為 Interested-party。", False),
 ("", False),
 ("已知限制：(1) 治理風險：2024 年會計師辭任與延遲申報、2026-03 前員工與共同創辦人因出口管制被起訴、FY2026 10-K 內控仍未有效（BDO 否定意見）；不改情境結構，以 Sens 案例（大客戶流失、淨利係數）呈現；(2) FY2027 份額以公司指引校準，FY2028 起份額、AI 毛利率為 Assumed；(3) 最大客戶占 FY2026 營收 28%、單一供應商占採購 63.1%；(4) 每顆系統價值以 NVIDIA 整櫃買方價近似，HGX 與 AMD 為 Analogy；(5) EPS 推現金流未扣營運資金（FY2026 營業現金流 −68 億美元），DCF 偏高；(6) 共識 EPS 為 non-GAAP，本模型為 GAAP；(7) 1Q FY27 財報（2026-11-03）已於 v0.2 滾入；EPS 口徑未扣兩類法下分配予特別股的盈餘（1Q 模型口徑 2.03 vs 揭露 1.75）；R4 產業部署路徑未建。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ================= v0.1.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（工作單；Andy 2026-10-25 14:00、14:17） =================
# 全部放在 Val、README 現有最後一列之後，不位移既有儲存格
OBS = [
       "目標價 US$86.9（DCF 88.3、Comparable 85.5 各半），比現價 32.94（2026-10-23）高 163.7%，級距大於 50%；Comparable 單獨為 159.5%。主因是倍數與 EPS 兩者：同業（Dell、HPE、緯穎、廣達、緯創）前瞻 P/E 中位數 11.7 倍，現價隱含只有 7.1 倍、共識目標價 43.27 隱含 9.3 倍（比同業低 21%），Comparable 未計美超微的治理與會計折價；本模型 12 個月後前瞻 EPS 7.29（GAAP）也比共識前瞻 EPS 4.67（non-GAAP，期間約早一年、尚未反映 11-03 上修）高 56.2%。",
       "v0.2 滾入 1Q FY27：營收 199.5 億美元（原指引中點 150 億，+33%）、GAAP 毛利率 12.7%、GAAP 稀釋 EPS 1.75；公司把 FY2027 營收指引由 650～720 億上修為 720～820 億，模型 FY2027 EPS（B）5.14 → 5.82。但模型 EPS 口徑〔（淨利 − 特別股股利）÷ 稀釋股數〕套在 1Q 得 2.03，比公司揭露的 1.75 高 15.8%：公司依兩類法另把約 2.5 億美元盈餘分配給強制可轉換特別股，模型未扣，FY2027–29 EPS 因此偏高約一成半（方法未改，待 Andy 判斷）。",
       "AI 毛利率仍是最敏感的變數：取 C 值時 FY2028 EPS −30.2%、B 值 −12.1%；1Q 實績 12.7% 高於模型 FY2027 B 的 10.7%。情境分歧大：FY2028 EPS A／B／C＝9.89／6.87／3.97（C 比 A 低 59.9%），DCF 情境價值 146.9／74.6／34.3；治理事件（大客戶流失）案例 FY2028 EPS −31.0%。下一個檢驗點：2Q FY27 財報（約 2027-02）營收是否落在 195～215 億、毛利率能否守在 12% 以上。"
      ]   # 重點觀察（與 README.md 相同文字；數字取自重算後的 xlsx）
_U = f"B{T['upw']}"
GR = V.max_row + 3
V.cell(GR - 1, 1, "0.4 上檔空間級距與投資建議（v0.1.1 新增；放在本頁最後以免位移既有儲存格；依「上檔空間：加權目標價」第 " + str(T["upw"]) + " 列）").font = BOLD
tset(GR, "grade", "上檔空間級距", [f'=IF({_U}>0.5,"大於 50%",IF({_U}>0.25,"25%–50%",IF({_U}>=0,"0%–25%",IF({_U}>=-0.25,"0%～−25%",IF({_U}>=-0.5,"−25%～−50%","小於 −50%")))))', None, None], "General",
     "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）；負值區間同樣重要")
tset(GR + 1, "rec", "投資建議", [f'=IF({_U}>0.5,"強力買進",IF({_U}>0.25,"買進",IF({_U}>=-0.25,"中立",IF({_U}>=-0.5,"放空","強力放空"))))', None, None], "General",
     "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空（剛好 +25% 為中立、+50% 為買進、−25% 為中立、−50% 為放空）")
rn = RD.max_row + 1
RD.cell(rn, 1, f'="投資建議："&Val!B{GR + 1}&"（依上檔空間固定規則）"').font = BOLD
OBS_ROW = rn + 1
for i_, t_ in enumerate(OBS, 1):
    RD.cell(rn + i_, 1, f"重點觀察 {i_}：" + t_)


# ---------------- 分部原則第 3 條（2026-10-28）：核對列與宣告（追加於工作表最後，不位移既有列，數值不變） ----------------
import os as _os2, sys as _sys2
_sys2.path.insert(0, _os2.path.join(_os2.path.dirname(_os2.path.dirname(_os2.path.abspath(__file__))), "tools", "segment_check"))
from addrow import add_check
add_check(Q, seg_rows=[16, 17], total_row=4, cols="BCDEF", rnd=3, note="分部原則第 3 條（2026-10-28 追加）；分部宣告見 segments.json")
RD.cell(RD.max_row + 1, 1, "v0.1.2（2026-10-28）：分部原則重建（數值不變）。Q 頁最後追加核對列「檢查：分部合計 − 合併營收」；官方分部與子列宣告見 segments.json。")
if isinstance(RD.cell(1, 1).value, str): RD.cell(1, 1).value = RD.cell(1, 1).value.replace("v0.1.1", "v0.1.2")

# ================= v0.2（2026-11-04）：滾入 1Q FY27 實績、2Q FY27 指引、FY2027 全年指引上修（工作單 brief_earn_20261104） =================
# 只改輸入與錨點（In 頁 g27、q1_*、sh27、sbc_ps）；方法、情境結構、估值規則不變。以下區塊追加在 Q、Recon、README 頁最後，不位移既有儲存格
SQ1 = "公司 1Q FY27 新聞稿（8-K Ex.99.1，2026-11-03；未經審計）"
q0 = Q.max_row + 2
Q.cell(q0, 1, "v0.2：1Q FY27（2026-07～09）實績與 2Q FY27（2026-10～12）指引（十億美元；EPS 美元；2026-11-03 公布）").font = BOLD
hdr(Q, q0 + 1, ["項目", "1Q FY27A（實績）", "2Q FY27 指引低", "2Q FY27 指引高", "1H FY27（1Q 實績＋2Q 指引中點）", "", "標記", "來源1（公司）", "來源2／說明"])
QV = {}
def q2row(r, key, label, vals, fmt="#,##0.000", tag="Verified", s1=SQ1, s2=""):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    Q.cell(r, 7, tag); Q.cell(r, 8, s1); Q.cell(r, 9, s2); QV[key] = r
r = q0 + 2
q2row(r, "rev", "營收", {"B": 19.945832, "C": 19.5, "D": 21.5, "E": f"=B{r}+(C{r}+D{r})/2"}, tag="Verified／Interested-party（指引）", s2="10-20 初步 185～195 億；原指引 145～155 億"); r += 1
q2row(r, "gp", "毛利", {"B": 2.541043}, s2="non-GAAP 毛利 2.549（股票報酬 0.008）"); r += 1
q2row(r, "gm", "毛利率", {"B": f"=B{QV['gp']}/B{QV['rev']}"}, fmt="0.00%", tag="Derived", s2="GAAP 12.7%、non-GAAP 12.8%（公司）；4Q FY26 17.5%、1Q FY26 9.3%"); r += 1
q2row(r, "opx", "營業費用（研發 0.2273、銷售 0.1325、管理 0.1167）", {"B": 0.47655}, s2="non-GAAP 0.3844"); r += 1
q2row(r, "oi", "營業利益", {"B": 2.064493}); r += 1
q2row(r, "oim", "營業利益率", {"B": f"=B{QV['oi']}/B{QV['rev']}"}, fmt="0.00%", tag="Derived"); r += 1
q2row(r, "nonop", "業外淨額（利息收入 0.0542 − 利息費用 0.0825 − 其他 0.0055）", {"B": "=0.054188-0.082539-0.005467"}, s2="稅前 2.0307"); r += 1
q2row(r, "tax", "所得稅（稅率 21.4%）", {"B": 0.435491}, s2="2Q 指引稅率 21.2%（GAAP）"); r += 1
q2row(r, "ni0", "淨利（特別股股利前；含權益法損失 0.001）", {"B": 1.594205}); r += 1
q2row(r, "nio", "淨利 ÷ 營業利益（In 頁 nio_f 的對照）", {"B": f"=B{QV['ni0']}/B{QV['oi']}"}, fmt="0.0%", tag="Derived", s2="In 頁 nio_f 基準 0.75"); r += 1
q2row(r, "pd", "特別股股利", {"B": 0.075796}, s2="In 頁 pref_div 年 0.302（季約 0.075）"); r += 1
q2row(r, "part", "分配予參與證券（強制可轉換特別股，兩類法）", {"B": 0.251561}, s2="模型 EPS 口徑未扣此項（見下兩列）"); r += 1
q2row(r, "nicd", "歸屬普通股淨利（稀釋，公司揭露）", {"B": 1.308286}, s2="＝基本 1.2668＋可轉債利息 0.0180＋稀釋重分配 0.0232"); r += 1
q2row(r, "sh", "GAAP 稀釋加權股數（十億股）", {"B": 0.749351, "C": 0.753, "D": 0.753}, fmt="0.0000", tag="Verified／Interested-party（指引）", s2="基本 0.6581；non-GAAP 稀釋 0.7644（2Q 指引 0.769）"); r += 1
q2row(r, "eps", "EPS（GAAP 稀釋，公司揭露）", {"B": 1.75, "C": 1.59, "D": 1.83, "E": f"=B{r}+(C{r}+D{r})/2"}, fmt="0.00", tag="Verified／Interested-party（指引）", s2="基本 1.92"); r += 1
q2row(r, "epsn", "EPS（non-GAAP 稀釋）", {"B": 1.81, "C": 1.66, "D": 1.90, "E": f"=B{r}+(C{r}+D{r})/2"}, fmt="0.00", tag="Verified／Interested-party（指引）", s2="1Q FY26 0.35"); r += 1
q2row(r, "epsm", "EPS 依模型口徑＝（淨利 − 特別股股利）÷ 稀釋股數", {"B": f"=(B{QV['ni0']}-B{QV['pd']})/B{QV['sh']}"}, fmt="0.00", tag="Derived", s1="模型 PL 頁的 EPS 公式套用 1Q 實績", s2="比公司揭露高，差額來自兩類法下分配予特別股的盈餘（待 Andy 判斷）"); r += 1
q2row(r, "epsd", "  模型口徑 ÷ 公司揭露 − 1", {"B": f"=B{QV['epsm']}/B{QV['eps']}-1"}, fmt="0.0%", tag="Derived", s1="—"); r += 1
q2row(r, "ocf", "營業現金流", {"B": -0.434021}, s2="應收 +35.6 億、存貨 +24.8 億、遞延收入 +30.5 億"); r += 1
q2row(r, "fyg", "FY2027 全年營收指引（低／高；中點見 In 頁 g27）", {"C": 72.0, "D": 82.0, "E": f"=(C{r}+D{r})/2"}, tag="Interested-party", s2="原 2026-08-11 為 650～720 億"); r += 1
q2row(r, "h2", "全年指引隱含 2H FY27 營收（全年 − 1H）", {"C": f"=C{QV['fyg']}-E{QV['rev']}", "D": f"=D{QV['fyg']}-E{QV['rev']}", "E": f"=E{QV['fyg']}-E{QV['rev']}"}, tag="Derived", s1="—", s2="低端隱含 2H 低於 1H"); r += 1
for nm, k in (("q1a_rev", "rev"), ("q1a_gm", "gm"), ("q1a_oi", "oi"), ("q1a_ni0", "ni0"), ("q1a_eps", "eps"), ("q1a_epsm", "epsm")):
    defname(nm, f"Q!$B${QV[k]}")
defname("h1_eps", f"Q!$E${QV['eps']}"); defname("h1_rev", f"Q!$E${QV['rev']}")
# Recon：模型原估 vs 實績（追加於最後）
RC.cell(rr_ + 1, 1, "F. v0.2：1Q FY27 實績 vs 模型原估（v0.1.4）與公司指引").font = BOLD; rr_ += 2
hdr(RC, rr_, ["項目", "1Q FY27 實績", "原指引（2026-08-11）中點", "10-20 初步中點", "v0.1.4 模型 FY2027（B）", "v0.2 模型 FY2027（B）", "", "", "來源", "說明"]); rr_ += 1
for lab, b, c, d, e, f, note in (
        ("營收（US$B）", "=q1a_rev", 15.0, 19.0, 68.5, f"=PL!E{PB['B']['rev']}", "原模型 FY2027 取全年指引中點 685 億；1Q 實績 199.5 億比原指引中點高 33%"),
        ("毛利率", "=q1a_gm", 0.106, 0.1225, 0.108, f"=PL!E{PB['B']['gm']}", "原指引 10.4%–10.8%"),
        ("EPS（GAAP 稀釋，US$）", "=q1a_eps", 0.935, None, 5.14, f"=PL!E{PB['B']['eps']}", "原指引 0.89–0.98；1H FY27（1Q 實績＋2Q 指引中點）＝" + "3.46")):
    RC.cell(rr_, 1, lab)
    for col, v in zip("BCDEF", (b, c, d, e, f)):
        if v is None: continue
        cc = RC[f"{col}{rr_}"]; cc.value = v; cc.number_format = "0.00%" if "毛利率" in lab else "#,##0.00"
        if not (isinstance(v, str) and v.startswith("=")): cc.font = BLUE
    RC.cell(rr_, 9, "公司 8-K（2026-08-11、10-20、11-03）；v0.1.4 成品"); RC.cell(rr_, 10, note); rr_ += 1
RC.cell(rr_, 1, "v0.2 模型 FY2027 EPS（B）− 1H FY27（1Q 實績＋2Q 指引中點）＝隱含 2H EPS"); RC[f"B{rr_}"] = f"=PL!E{PB['B']['eps']}-h1_eps"; RC[f"B{rr_}"].number_format = "0.00"
RC.cell(rr_, 10, "1H 為公司揭露口徑（兩類法）；模型口徑未扣參與證券分配，見 Q 頁"); rr_ += 1
RC.cell(rr_, 1, "1Q FY27：模型 EPS 口徑 ÷ 公司揭露 − 1"); RC[f"B{rr_}"] = "=q1a_epsm/q1a_eps-1"; RC[f"B{rr_}"].number_format = "0.0%"
RC.cell(rr_, 10, "模型 EPS＝（淨利 − 特別股股利）÷ 稀釋股數；公司另扣分配予參與證券（強制可轉換特別股）。方法未改，待 Andy 判斷"); rr_ += 1
RD.cell(RD.max_row + 1, 1, "v0.2（2026-11-04）：滾入 1Q FY27 實績（營收 199.46 億美元、GAAP 毛利率 12.7%、GAAP 稀釋 EPS 1.75、non-GAAP 1.81）與 2Q FY27 指引（營收 195～215 億、GAAP EPS 1.59～1.83）；FY2027 營收指引由 650～720 億上修為 720～820 億（情境 A 上緣、B 中點、C 下緣）；FY2027 稀釋股數 7.51 億（1Q 實績與 2Q 指引平均）。方法、情境結構與估值規則不變；股價沿用 2026-10-23 收盤。1Q 明細見 Q 頁最後、模型原估 vs 實績見 Recon 頁 F 節。")
if isinstance(RD.cell(1, 1).value, str): RD.cell(1, 1).value = RD.cell(1, 1).value.replace("v0.1.2", "v0.2").replace("2026-10-25", "2026-11-04")
order = ["README", "In", "Q", "Units", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "smci")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.1.3", "2026-10-30", split=True, split_skip=[])  # v0.1.3：In 頁版面；v0.1.4（2026-10-31）：第 3 點情境專屬列拆分（數值不變）
# ---------------- Siliconomics v0.13（2026-11-04；決策 S37）：Sens「Wattonomics 錨定」案例 ----------------
# In 頁最後追加 Z 節（wta_27／28／29，數值取 SIL.WTA）；Sens 頁既有內容最下方另起小表與案例區塊。既有儲存格不位移、數值不變。
WTA_NOTE = "Wattonomics 錨定：以情境 B 為底，產業量 CY2027 起 × wta（Siliconomics SIF_WTAnchor_Base＝R4 比值）；FY2026 維持錨點。三層校準原則（Andy 2026-11-04）；個股基準暫不改，2027-01～02 雲端資本支出指引後再決定。"
def _wta_in():
    r = IN.max_row + 2
    IN.cell(r, 1, "Z. Wattonomics 錨定係數（敏感度用；Siliconomics v0.13）").font = BOLD; r += 1
    def w(k, lab, unit, lo, b, hi, tag, s1, s2):
        nonlocal r
        for j, v in enumerate((k, lab, unit, lo, b, hi, tag, s1, s2), 1):
            if v is not None: IN.cell(r, j, v)
        IN.cell(r, 5).font = BLUE; IN.cell(r, 5).number_format = "0.000"
        wb.defined_names[k] = DefinedName(k, attr_text=f"In!$E${r}"); r += 1
    SIL.wta_rows(w)
def _wta_sens(mf, setting, how):
    r0 = SE.max_row + 3
    SE.cell(r0, 1, "Wattonomics 錨定（三層校準原則）").font = BOLD
    SE.cell(r0 + 1, 1, WTA_NOTE)
    SE.cell(r0 + 2, 1, how)
    ncol = max(j for j in range(1, 30) if SE.cell(3, j).value is not None)
    H = [SE.cell(3, j).value for j in range(1, ncol + 1)]
    hdr(SE, r0 + 3, H)
    base_b = "情境 B" in str(SE.cell(4, 1).value)
    rb, rB = r0 + 4, r0 + 5
    rw = rB + (0 if base_b else 1)
    rd = rw + 1
    WR = engine(SE, rd + 4, "案例區塊：Wattonomics 錨定（情境 B × wta）", scn="B", mf=mf)
    SE.cell(rb, 1, "基準（主表第 4 列）"); SE.cell(rb, 2, SE.cell(4, 1).value)
    for j in range(3, ncol + 1):
        L = SE.cell(4, j).column_letter
        SE[f"{L}{rb}"] = f"={L}4"; SE[f"{L}{rb}"].number_format = SE[f"{L}4"].number_format
    if not base_b:
        put_case(rB, "情境 B（PL 區塊 B）", "Siliconomics 基準係數（未乘 wta）", PB["B"], "PL!")
    put_case(rw, "Wattonomics 錨定", setting, WR, "")
    for k, (rr, lab) in enumerate(((rb, "錨定 − 基準（主表第 4 列）"), (rB, "錨定 − 情境 B")) if not base_b else ((rb, "錨定 − 基準（主表第 4 列）＝情境 B"),)):
        ri = rd + k
        SE.cell(ri, 1, lab); SE.cell(ri, 2, "差異%（毛利率、營益率等百分比欄為百分點差）")
        for j in range(3, ncol + 1):
            if "相對基準" in str(H[j - 1]): continue
            L = SE.cell(4, j).column_letter
            pct = "%" in SE[f"{L}4"].number_format
            SE[f"{L}{ri}"] = f"={L}{rw}-{L}{rr}" if pct else f"={L}{rw}/{L}{rr}-1"
            SE[f"{L}{ri}"].number_format = "0.0%"
    return dict(r0=r0, base=rb, case=rw, diff=rd, block=WR)
_wta_in()
WTA_SENS = _wta_sens({2027: "wta_27", 2028: "wta_28", 2029: "wta_29"}, '情境 B；商用 GPU 顆數（日曆年）CY2027 起 × wta_27／28／29；會計年度市場＝½ 前一 CY＋½ 當年 CY', '做法：引擎 mf（預設為 In 頁 mf＝1）改為年別 dict（日曆年 wta_27／28／29），乘在日曆年 GPU 顆數列；會計年度（7 月～6 月）沿用模型既有換算 FY＝½ 前一 CY＋½ 當年 CY，所以 FY2027 只有一半受 CY2027 影響。FY2030–32 沿用引擎延伸表的收斂法。目標價：Sens 目標價表只有 DCF 參數案例，本案例不重算目標價。')
# ================= v0.3（2026-11-05）：展示層改為個股 HTML 版面規格 v1.0（與標準版 DCF 同一版；docs/plan/個股HTML版面規格.md） =================
def _lab(ws, prefix, start=1):
    for r_ in range(start, ws.max_row + 1):
        v_ = ws.cell(r_, 1).value
        if isinstance(v_, str) and v_.strip().startswith(prefix.strip()): return r_
    raise KeyError((ws.title, prefix))
def _in(k, c="D"):
    for r_ in range(4, IN.max_row + 1):
        if IN.cell(r_, 1).value == k: return f"In!{c}{r_}"
    raise KeyError(k)
_v = lambda p, c="B": f"Val!{c}{_lab(V, p)}"
_s = lambda p, c: f"Sens!{c}{_lab(SE, p)}"
_r = lambda p, c: f"Recon!{c}{_lab(RC, p)}"
_pa, _pb, _pc = PB["A"], PB["B"], PB["C"]
TPW, DCF, CMP, PX, UP, REC = _v("目標價（DCF 與 Comparable 加權）"), _v("DCF 目標價（情境加權）"), _v("Comparable 目標價（相對倍數調整後）"), _v("現價（"), _v("上檔空間：加權目標價"), _v("投資建議")
FEPS, MED, CONS = _v("12 個月後的前瞻 EPS"), _v("同業前瞻 P/E 中位數"), _v("市場共識目標價")
NTM, COEF, CMP0, UPC0 = _v("未來 12 月 EPS（模型 B）"), _v("相對倍數調整係數"), _v("Comparable 目標價（市場主流同業法）"), _v("上檔空間：Comparable 目標價")
PE12 = _v("  現價 ÷ 未來 12 月 EPS")
D12 = lambda c: _v("DCF 情境價值推到 12 個月後", c)
_SG = '"+0.0%;-0.0%"'
_e1, _e2, _m = f"({NTM}/cons_feps)", f"({FEPS}/{NTM})", f"({MED}*{COEF}/({PX}/cons_feps))"
_wd, _wc = f"tp_w_dcf*({DCF}-{PX})/({TPW}-{PX})", f"(1-tp_w_dcf)*({CMP}-{PX})/({TPW}-{PX})"
_le = f"(LN({_e1})+LN({_e2}))/LN({CMP}/{PX})"
_main = f'IF(ABS({_wd}-{_wc})<0.1,"由 DCF 與 Comparable 各貢獻約一半",IF({_wd}>{_wc},"主要來自 DCF","主要來自 Comparable"))'
RD["A1"] = RD["A1"].value + "；v0.3（2026-11-05）：README 新增展示層、命題回覆與本版要點各段，HTML 改用個股版面規格 v1.0；模型改版：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、In 頁 H2 節、Sens 目標價參數），Comparable、情境權重、DCF 權重 50% 與投資建議規則不變；v0.3.1（2026-11-05）β 改為自算產業 β（US 伺服器／EMS 同業中位數，Blume 調整；Andy 2026-11-05 22:55），評價日 2026-12-31 不變，其他模型參數不變；v0.3.2（2026-11-05）第二輪修正：換用共用模組 tools/fcff_dcf.py v2（淨現金由資產負債表日 2026-09-30 滾到評價日、β 敏感度改 β_u ± 0.3）"
r = RD.max_row + 2
RD.cell(r, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；v0.3 新增，原 A3、A4、重點觀察各列保留）").font = BOLD; r += 1
RD.cell(r, 1, "命題：美超微是 AI 伺服器品牌商，以買賣模式賣含 GPU 的整機與整櫃，營收跟著商用 GPU 顆數與自身份額走、毛利率約一成。要回答的是 FY2027 營收指引能否轉成每股盈餘、守住毛利率，以及約 7 倍前瞻本益比的治理折價是否過深。"); r += 1
RD.cell(r, 1, "命題 1｜AI 營收跟著商用 GPU 顆數成長：AI 營收＝商用 GPU 顆數（讀台積電模型）× 每顆系統價值（橋接表整櫃價）× 美超微有效份額（FY2027 以公司營收指引校準，FY2028 起依情境變動）。"); r += 1
RD.cell(r, 1, "命題 2｜毛利率守在約一成：AI 毛利率依情境調整（A 最有利、C 最不利），非 AI 毛利率與成長率另列；淨利 ÷ 營業利益反映利息與合規費用。"); r += 1
RD.cell(r, 1, f'="命題 3｜治理折價可能過深：現價只有本模型目前未來 12 個月 EPS 的 "&TEXT({PE12},"0.0")&" 倍（GAAP），低於同業前瞻本益比中位數 "&TEXT({MED},"0.0")&" 倍（同業為共識口徑，多為 non-GAAP；同為目前的未來 12 個月）；若盈餘路徑成立、治理事件不擴大，每股價值應高於現價。"'); r += 1
RD.cell(r, 1, f'="重點 1｜目標價比現價高 "&TEXT({UP},"0%")&"，上檔"&{_main}&"，兩者都建立在本模型盈餘遠高於現價隱含的水準：加權目標價每股約 "&TEXT({TPW},"0.00")&" 美元（DCF "&TEXT({DCF},"0.00")&"、相對倍數調整後的 Comparable "&TEXT({CMP},"0.00")&"，權重 "&TEXT(tp_w_dcf,"0%")&"／"&TEXT(1-tp_w_dcf,"0%")&"），現價 "&TEXT({PX},"0.00")&" 美元，投資建議"&{REC}&"。加權上檔約 "&TEXT({TPW}-{PX},"0.00")&" 美元中，DCF 貢獻 "&TEXT({_wd},"0%")&"、Comparable 貢獻 "&TEXT({_wc},"0%")&"。Comparable 比現價高 "&TEXT({CMP}/{PX}-1,"0%")&"，拆成三個因素：同一期間（目前的未來 12 個月）本模型 EPS "&TEXT({NTM},"0.00")&" 美元（GAAP）對共識前瞻 EPS "&TEXT(cons_feps,"0.00")&" 美元（non-GAAP）"&TEXT({_e1}-1,{_SG})&"；到約一年後的未來 12 個月再成長 "&TEXT({_e2}-1,{_SG})&"（"&TEXT({FEPS},"0.00")&" 美元）；目標倍數 "&TEXT({MED}*{COEF},"0.0")&" 倍（同業前瞻本益比中位數 "&TEXT({MED},"0.0")&" 倍 × 相對倍數調整 "&TEXT({COEF},"0.00")&"）對現價隱含的共識前瞻本益比 "&TEXT({PX}/cons_feps,"0.0")&" 倍 "&TEXT({_m}-1,{_SG})&"。盈餘因素約占 Comparable 上檔的 "&TEXT({_le},"0%")&"、倍數因素約占 "&TEXT(1-{_le},"0%")&"（對數拆解）。只看市場主流同業法（不做相對倍數調整）的 Comparable 為 "&TEXT({CMP0},"0.00")&" 美元、上檔 "&TEXT({UPC0},"0%")&"。"'); r += 1
RD.cell(r, 1, f'="重點 2｜FY2027 以公司營收指引校準：營收樂觀、基準、保守為 "&TEXT(PL!E{_pa["rev"]}*10,"#,##0")&"、"&TEXT(PL!E{_pb["rev"]}*10,"#,##0")&"、"&TEXT(PL!E{_pc["rev"]}*10,"#,##0")&" 億美元（指引上緣、中點、下緣），GAAP EPS "&TEXT(PL!E{_pa["eps"]},"0.00")&"、"&TEXT(PL!E{_pb["eps"]},"0.00")&"、"&TEXT(PL!E{_pc["eps"]},"0.00")&" 美元。敏感度測試（以情境 A 為基準）：AI 毛利率低值時 FY2027 EPS "&TEXT({_s("AI 毛利率 低（C 值）", "F")},{_SG})&"；淨利 ÷ 營業利益 0.68 時 "&TEXT({_s("淨利 ÷ 營業利益 0.68", "F")},{_SG})&"；營收落在指引下緣時 "&TEXT({_s("FY2027 營收＝指引下緣", "F")},{_SG})&"。"'); r += 1
RD.cell(r, 1, f'="重點 3｜治理事件是最大的下檔：敏感度測試（以情境 A 為基準）中「大客戶流失」（FY2027 指引下緣、FY2028 起份額 ×0.8）時 FY2028 EPS "&TEXT({_s("治理事件：大客戶流失", "J")},{_SG})&"；份額流失（C 值）時 "&TEXT({_s("份額流失（C 值）", "J")},{_SG})&"。FY2028 EPS 樂觀、基準、保守約 "&TEXT(PL!F{_pa["eps"]},"0.00")&"、"&TEXT(PL!F{_pb["eps"]},"0.00")&"、"&TEXT(PL!F{_pc["eps"]},"0.00")&" 美元。治理與會計風險不另加折現率。"'); r += 1
_sv = lambda p, c: f"Sens!{c}{_lab(SE, p, TPS0)}"
RD.cell(r, 1, f'="重點 4｜v0.3 起 DCF 改為標準版（無槓桿自由現金流折現；GAAP 營業利益，股權報酬視為費用），DCF 目標價為 "&TEXT({DCF},"0.00")&" 美元（舊法為股東現金流、折現率 10%；新舊對照見本版要點）。WACC "&TEXT(wacc,"0.0%")&"（股權成本 "&TEXT(ke,"0.0%")&"＝美國 10 年期公債 "&TEXT(rf,"0.00%")&" ＋ β "&TEXT(beta,"0.00")&" × ERP "&TEXT(erp,"0%")&"；β 為自算產業 β（美國伺服器／EMS 同業 8 家無槓桿 β 中位數，Blume 調整）依美超微負債比重新加槓桿，ERP 5% 與 RONIC＝WACC 為假設），永續成長率 "&TEXT(g_tv,"0.0%")&"。營運資金約占營收增量 "&TEXT(fc_wc,"0%")&"：FY2027 營運資金增加 "&TEXT(DCF!C{RD_["paths"]["B"]["dwc"]}*10,"#,##0")&" 億美元，情境 B 的 FY2027 無槓桿自由現金流 "&TEXT(DCF!C{RD_["paths"]["B"]["fcff"]}*10,"#,##0")&" 億美元（評價日後半年按比例計入）。強制可轉換特別股以轉換股數計入股數。DCF 三情境（推到 12 個月後）"&TEXT({D12("B")},"0.00")&"／"&TEXT({D12("C")},"0.00")&"／"&TEXT({D12("D")},"0.00")&" 美元；終值現值占情境 B 企業價值 "&TEXT(Val!C{SDCF["tvsh"]},"0%")&"、情境 A "&TEXT(Val!B{SDCF["tvsh"]},"0%")&IF(Val!B{SDCF["tvsh"]}>0.8,"（超過 80%）","")&"。敏感度（以情境加權目標價為基準）：WACC −1pt／＋1pt 時 "&TEXT({_sv("WACC ±1pt", "D")},{_SG})&"／"&TEXT({_sv("WACC ±1pt", "F")},{_SG})&"；營運資金比例 "&TEXT(fc_wc_lo,"0%")&"／"&TEXT(fc_wc_hi,"0%")&" 時 "&TEXT({_sv("營運資金 ÷ 營收", "D")},{_SG})&"／"&TEXT({_sv("營運資金 ÷ 營收", "F")},{_SG})&"。"'); r += 1
_tvs3 = [f"Val!{c}{SDCF['tvsh']}" for c in "BCD"]
_tvflag = "&".join(f'IF({t}>0.8,"情境 {n}（{l}）"&TEXT({t},"0%")&"；","")' for t, n, l in zip(_tvs3, "ABC", ("順利", "平平", "不如預期")))
_bsv = lambda c: f"Sens!{c}{_lab(SE, 'β 低／高', TPS0)}"
RD.cell(r, 1, f'="重點 5｜淨現金滾到評價日、β 敏感度與終值占比：資產負債表為 "&bs_date&"（最近一季），淨現金與投資 "&TEXT(dcf_nc,"#,##0.00")&" US$B，加上 stub 期無槓桿自由現金流（FY2027 全年 FCFF × "&TEXT(stub_frac,"0.000")&" 年）"&TEXT(Val!C{SDCF["stub"]},"#,##0.00")&" US$B、扣除已宣告股利 "&TEXT(stub_div,"#,##0.00")&" US$B（普通股不配息），滾到評價日為 "&TEXT(Val!C{SDCF["ncr"]},"#,##0.00")&" US$B（情境 B）。β 改用產業無槓桿 β "&TEXT(beta_u,"0.000")&" ∓ "&TEXT(beta_du,"0.0")&" 各自重新加槓桿（β "&TEXT(beta_lo,"0.00")&"／"&TEXT(beta_hi,"0.00")&"）時，加權目標價變動 "&TEXT({_bsv("D")},{_SG})&"／"&TEXT({_bsv("F")},{_SG})&"（敏感度，以情境加權目標價為基準）。美國公司以美元計價，「美元口徑折現」即基準，不另列。"&IF(MAX({",".join(_tvs3)})>0.8,"終值占比標示（終值現值占企業價值）："&{_tvflag}&"超過 80%，這些情境的價值主要取決於永續期假設（g、RONIC＝WACC）。","三情境終值現值占企業價值都未超過 80%。")'); r += 1
r += 1
RD.cell(r, 1, "命題回覆（HTML 用；依本模型計算逐一回答展示層的三個命題）").font = BOLD; r += 1
RD.cell(r, 1, f'="AI 營收能跟著商用 GPU 顆數成長嗎？答：FY2027 以指引校準、FY2028 起放慢。AI 營收 FY2027 約 "&TEXT(Units!D{UR["ai"]}*10,"#,##0")&" 億美元、FY2028 約 "&TEXT(Units!E{UR["ai"]}*10,"#,##0")&" 億美元（情境 A）；FY2028 營收年增樂觀、基準、保守為 "&TEXT(PL!F{_pa["g"]},"0.0%")&"、"&TEXT(PL!F{_pb["g"]},"0.0%")&"、"&TEXT(PL!F{_pc["g"]},"0.0%")&"。份額相對值與每顆系統價值的變動為假設。"'); r += 1
RD.cell(r, 1, f'="毛利率守得住約一成嗎？答：本模型假設守住。毛利率 FY2027 樂觀、基準、保守為 "&TEXT(PL!E{_pa["gm"]},"0.0%")&"、"&TEXT(PL!E{_pb["gm"]},"0.0%")&"、"&TEXT(PL!E{_pc["gm"]},"0.0%")&"（AI 毛利率依情境調整，皆為假設）；AI 毛利率每 1 個百分點約影響 EPS 一成（Sens 頁讀法）。"'); r += 1
RD.cell(r, 1, f'="治理折價是否過深、每股值多少？答："&IF({DCF}>{PX},"依本模型偏深。","不一定。")&"現價 "&TEXT({PX},"0.00")&" 美元只有本模型目前未來 12 個月 EPS 的 "&TEXT({PE12},"0.0")&" 倍。每股約 "&TEXT({TPW},"0.00")&" 美元（DCF "&TEXT({DCF},"0.00")&"、相對倍數調整後的 Comparable "&TEXT({CMP},"0.00")&"；只看市場主流同業法為 "&TEXT({CMP0},"0.00")&"），比現價高 "&TEXT({UP},"0.0%")&"，投資建議"&{REC}&"；DCF 三情境價值（推到 12 個月後）樂觀、基準、保守約 "&TEXT({D12("B")},"0.00")&"、"&TEXT({D12("C")},"0.00")&"、"&TEXT({D12("D")},"0.00")&" 美元，"&IF(MIN({D12("B")},{D12("C")},{D12("D")})>{PX},"都高於現價","並非都高於現價")&"。市場共識目標價 "&TEXT({CONS},"0.00")&" 美元，本模型加權目標價比它 "&TEXT({TPW}/{CONS}-1,{_SG})&"；這建立在本模型 GAAP 盈餘路徑與治理事件不擴大的假設上。"'); r += 1
r += 1
RD.cell(r, 1, "HTML 輔助列（只供 HTML 抽數，不是模型輸出；B 欄為公式）").font = BOLD; r += 1
RD.cell(r, 1, "上檔空間：Comparable 目標價（相對倍數調整後）"); RD.cell(r, 2, f"={CMP}/{PX}-1").number_format = "0.0%"; r += 1
RD.cell(r, 1, f'="估值說明：加權目標價＝DCF 目標價 × "&TEXT(tp_w_dcf,"0%")&"＋Comparable 目標價（相對倍數調整後）× "&TEXT(1-tp_w_dcf,"0%")&"，投資建議依它的上檔空間判定。DCF 為標準版：無槓桿自由現金流（稅後營業利益＋折舊 − 資本支出 − 營運資金增加）以 WACC "&TEXT(wacc,"0.0%")&" 折現，終值用永續成長率 "&TEXT(g_tv,"0.0%")&"、新投資報酬率＝WACC，企業價值減淨負債與特別股轉換前股利（資產負債表日的淨現金加 stub 期無槓桿自由現金流、扣已宣告股利，滾到評價日）後除以股數（含特別股轉換），三情境價值以股權成本推到 12 個月後，再依情境權重加權；Comparable＝同業前瞻本益比中位數 × 本公司長期相對同業的倍數比 × 約一年後的未來 12 個月 EPS（GAAP，情境加權）。十年本益比規則值是舊法，只作參考，不進目標價。"'); r += 1
r += 1
RD.cell(r, 1, "本版要點（版本沿革；HTML 附錄用）").font = BOLD; r += 1
RD.cell(r, 1, f'="v0.3.2（2026-11-05）：第二輪修正（Andy 2026-11-05 23:16 定案），換用共用模組 tools/fcff_dcf.py v2。(1) 淨現金滾到評價日：資產負債表日 "&bs_date&"（1Q FY27 已公布）→ 評價日 2026-12-31，淨現金與投資 "&TEXT(dcf_nc,"0.00")&"＋stub 期 FCFF（FY2027 全年 FCFF × "&TEXT(stub_frac,"0.000")&" 年；情境 B "&TEXT(Val!C{SDCF["stub"]},"0.00")&"）− 已宣告股利 "&TEXT(stub_div,"0.00")&"＝"&TEXT(Val!C{SDCF["ncr"]},"0.00")&"（US$B；情境 B）。(2) β 敏感度改 β_u ± 0.3 各自重新加槓桿（β "&TEXT(beta_lo,"0.00")&"／"&TEXT(beta_hi,"0.00")&"）；基準 β_u 1.322 不變（β "&TEXT(beta,"0.00")&"）。(3) 美國公司不加「美元口徑折現」（＝基準）。(4) RONIC＝WACC 不變；非控制權益（含特別股轉換前股利）不變；資本支出、折舊：本模型沒有擴產排程列，維持近三年占營收比例。β 1.73 → "&TEXT(beta,"0.00")&"、WACC 10.7% → "&TEXT(wacc,"0.0%")&"、DCF 目標價 72.40 → "&TEXT({DCF},"0.00")&"、加權目標價 76.01 → "&TEXT({TPW},"0.00")&"、上檔空間 +130.8% → "&TEXT({UP},"+0.0%;-0.0%")&"、投資建議 強力買進 → "&{REC}&"。"'); r += 1
RD.cell(r, 1, "v0.3.1（2026-11-05）：β 改為自算產業 β（US 伺服器／EMS 同業 8 家無槓桿 β 中位數 1.48，Blume 調整後 1.322，依美超微負債比重新加槓桿；Andy 2026-11-05 22:55，取代 Damodaran 產業 β）。β 1.72 → 1.73、WACC 10.7% → 10.7%、DCF 目標價 72.48 → 72.40、加權目標價 76.05 → 76.01、上檔空間 +130.9% → +130.8%、投資建議 強力買進 → 強力買進。評價日 2026-12-31、目標時點 12 個月後與其他模型參數不變。"); r += 1
RD.cell(r, 1, "v0.3（2026-11-05）：模型改版，DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC，docs/plan/個股DCF標準規格.md）。DCF 目標價 88.26 → 72.48、加權目標價 83.94 → 76.05、上檔空間 +154.8% → +130.9%、投資建議 強力買進 → 強力買進。DCF 差異拆解見 PR 內文。Comparable、情境權重、DCF 權重 50% 與投資建議規則不變。"); r += 1   # 舊版沿革：固定為該版成品的計算結果
RD.cell(r, 1, "v0.3（2026-11-05）展示層：README 新增展示層、命題回覆、HTML 輔助列與本版要點；HTML 改用個股 HTML 版面規格 v1.0（投資建議橫幅、命題分點與命題回覆、驅動因子依業務分塊、估值與對帳只列主要項目）。"); r += 1
RD.cell(r, 1, "=A2"); r += 1
wb.save(sys.argv[1])
if os.environ.get("SMCI_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE, "UN_NOTE": UN_NOTE,
               "T": T, "T10": T10, "TPS0": TPS0, "TPS_LAST": TPS_LAST, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0, "NCASES": len(cases)},
              open(os.environ["SMCI_ROWMAP"], "w"), ensure_ascii=False, indent=1)

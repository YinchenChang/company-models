# 緯創資通（3231）收支與估值模型 v0.1.1（AI 半導體 Project：伺服器組裝次產業，樣板＝quanta/）
# 用法：python3 wistron/build_wistron.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；整櫃價格與市場以美元輸入、以匯率 fx 換算；EPS：新台幣。
# 口徑：營收、毛利、營業利益為合併報表（含子公司緯穎 6669）；EPS 用歸屬母公司淨利＝緯創本體淨利＋持股比例 × 緯穎淨利。
# 拆成兩塊：緯穎（合併子公司，持股約 37.68%）與緯創本體（合併 − 緯穎）。本體再分 AI 系統（L10／L11）、GPU 基板與運算板、一般伺服器與交換器、非伺服器。
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
NET_F = only("markets/network/dist/*.xlsx"); TSMC_F = only("tsmc/dist/*.xlsx"); BR_F = only("markets/chip_bridge/dist/*.xlsx")
QU_F = only("quanta/dist/*.xlsx"); FX_F = only("foxconn/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
ITC = None
for r in openpyxl.load_workbook(NET_F, data_only=True)["Mkt"].iter_rows(min_row=4):
    if r[0].value == "AI IT 資本支出（加速器 × 每顆 IT capex）":
        ITC = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
assert ITC and all(v is not None for v in ITC[0]), ITC
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and (str(k).startswith("u_") or str(k).startswith("g29_") or str(k).startswith("um")):
        TS[k] = (r[4].value, r[0].row)
GPLAT = ("bw", "rb", "ru", "hp", "amd")      # GPU 基板適用：NVIDIA 各世代＋AMD Instinct（ASIC 不計）
assert all(f"u_{p}_{y}" in TS for p in GPLAT for y in (26, 27, 28)) and all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()
GU = {y: sum(TS[f"u_{p}_{y}"][0] for p in GPLAT) for y in (26, 27, 28)}
GU[29] = sum(TS[f"u_{p}_28"][0] * (1 + TS[f"g29_{p}"][0]) for p in GPLAT)
GU_SRC = f"{rel(TSMC_F)}｜In!E{TS['u_bw_26'][1]}:E{TS['g29_amd'][1]}（NVIDIA Blackwell、Rubin、Rubin Ultra、Hopper 與 AMD Instinct 加總；2029＝2028 ×（1＋各類 2029 成長））"
BR = {}
for r in openpyxl.load_workbook(BR_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k in ("racks", "capit", "rack"):
        BR[k] = (r[4].value, r[7].value, f"{rel(BR_F)}｜In!E{r[0].row}（GB300）、H{r[0].row}（VR200）")
assert len(BR) == 3, BR.keys()
# 廣達模型（組裝線權威，規則第 13 條）：ODM 每櫃價格、Rubin 占比、匯率、每 GW 價值變動、客供比例；份額只作合計檢查
QWB = openpyxl.load_workbook(QU_F, data_only=True)
QU = {}
for r in QWB["In"].iter_rows(min_row=4):
    k = r[0].value
    if k in ("odm_gb", "odm_vr", "vr26", "fx", "vpg_28", "vpg_29", "cs_27", "cs_28", "cs_29", "keep"):
        QU[k] = (r[3].value, r[4].value, r[5].value, f"{rel(QU_F)}｜In!D{r[0].row}:F{r[0].row}")
assert len(QU) == 10, QU.keys()
def units_row(wbk, label, f):
    for r in wbk["Units"].iter_rows(min_row=4):
        if r[0].value == label:
            return ([r[2 + j].value for j in range(4)], f"{rel(f)}｜Units!C{r[0].row}:F{r[0].row}")
    raise KeyError(label)
QSH = units_row(QWB, "廣達有效份額", QU_F)
FSH = units_row(openpyxl.load_workbook(FX_F, data_only=True), "鴻海有效份額（金額口徑）", FX_F)

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B），整櫃價格與市場為美元；顆數：百萬顆；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間，不當點估計（放情境或 Sens）。SemiAnalysis 未使用。公司原文＝緯創、緯穎法說（經鉅亨、優分析、富果摘要）與月營收公告；公司展望為 Interested-party；券商與資料商為 Interested-party。組裝市場與 ODM 每櫃價格讀廣達模型（唯讀），口徑一致"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
CNY2Q = "https://news.cnyes.com/news/id/6559382"
UAN2Q = "https://uanalyze.com.tw/articles/5392152784"
UAN25 = "https://uanalyze.com.tw/articles/6054848921"
SA_FIN = "https://stockanalysis.com/quote/tpe/3231/financials/"
KGI05 = "https://www.kgi.com.hk/zh-cn/-/media/files/kgishk/research-reports/tw-reports/wistron_3231-tt_08052026e.pdf"
KGIW = "https://www.kgi.com.hk/zh-hk/-/media/files/kgishk/research-reports/tw-reports/wiwynn-6669-tt20260728en.pdf"
FUGW = "https://blog.fugle.tw/post/earnings-call-6669-2026-02-26-2"
UANW2Q = "https://uanalyze.com.tw/articles/3734153022"
TN07 = "https://finance.technews.tw/2026/08/07/wistron-3231-202607-financial-report/"
TN08 = "https://finance.technews.tw/2026/09/09/wistron-3231-202608-financial-report/"
TN09 = "https://finance.technews.tw/2026/10/08/wistron-3231-202609-financial-report/"
UAN09 = "https://uanalyze.com.tw/articles/5406456758"
TNW09 = "https://finance.technews.tw/2026/10/08/wiwynn-6669-202609-financial-report/"
SD_GDR = "https://statementdog.com/news/17618"
UDN0910 = "https://money.udn.com/money/amp/story/5607/9745501"
SA_Q = "https://stockanalysis.com/quote/tpe/3231/"
SA_F = "https://stockanalysis.com/quote/tpe/3231/forecast/"
SA_R = "https://stockanalysis.com/quote/tpe/3231/financials/ratios/"
SD = "https://statementdog.com/analysis/3231"
GI = "https://goodinfo.tw/tw/StockBzPerformance.asp?STOCK_ID=3231&RPT_CAT=M_YEAR"
WR0810 = "https://wantrich.chinatimes.com/news/20260810900002-420101"
CTEE0805 = "https://www.ctee.com.tw/news/20260805700302-430203"
CM2024 = "https://www.cmoney.tw/notes/note-detail.aspx?nid=851483"

sec("A. AI 部署規模與加速器顆數（讀 markets/network、markets/chip_bridge、tsmc 成品；情境 A 基準，B／C 依顆數係數縮放）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}", "US$B", None, ITC[0][j], None, "Derived", ITC[1],
        "加速器顆數 × 每顆 IT capex（含 ASIC）；除以每 GW IT 資本支出得到部署 GW（與廣達、鴻海模型同口徑）" if y == 26 else "同上")
inp("itpg_gb", "每 GW IT 資本支出（GB300）", "US$B", None, BR["capit"][0], None, "Derived", BR["capit"][2], "Tokenomics v20.1 IF_CapexIT（經橋接表；唯讀）")
inp("itpg_vr", "每 GW IT 資本支出（VR200；2027 起）", "US$B", None, BR["capit"][1], None, "Derived", BR["capit"][2], "同上；Rubin Ultra 與 ASIC 機架未另訂（Assumed：同 VR200）")
inp("vr26", "2026 年部署中 Rubin 世代占比（混合用；廣達模型）", "%", QU["vr26"][0], QU["vr26"][1], QU["vr26"][2], "Derived", QU["vr26"][3], "廣達模型輸入（TrendForce 2026 組合 Rubin 22%）；緯創：Vera Rubin 已開始出貨（2026-09，優分析）")
inp("racks_gb", "每 GW 機架數（GB300）", "架", None, BR["racks"][0], None, "Derived", BR["racks"][2], "Tokenomics IF_RacksPerGW（經橋接表）")
inp("racks_vr", "每 GW 機架數（VR200）", "架", None, BR["racks"][1], None, "Derived", BR["racks"][2], "同上")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", f"{rel(TSMC_F)}｜In!E{TS[f'um28_{s}'][1]}", "A＝基準；B＝AI 資本支出消化；C＝明顯下修（與台積電、廣達、鴻海模型相同）")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", f"{rel(TSMC_F)}｜In!E{TS[f'um29_{s}'][1]}", "同上")
for y in (26, 27, 28, 29):
    inp(f"gu_{y}", f"GPU 顆數 20{y}（NVIDIA＋AMD，情境 A 基準）", "百萬顆", None, round(GU[y], 6), None, "Derived", GU_SRC if y == 26 else "同上",
        "GPU 基板（HGX）與運算板（Bianca／Strata）只對應 GPU；Google TPU、Trainium 等 ASIC 不計" if y == 26 else "同上")
inp("fx", "匯率（新台幣 ÷ 美元；廣達模型）", "NT$/US$", QU["fx"][0], QU["fx"][1], QU["fx"][2], "Derived", QU["fx"][3], "與廣達、鴻海、台達電模型相同")

sec("B. 組裝價格與客供（讀廣達模型，唯讀）、緯創本體與緯穎的份額")
inp("odm_gb", "ODM 每櫃價格 GB300（廣達模型）", "US$M/架", QU["odm_gb"][0], QU["odm_gb"][1], QU["odm_gb"][2], "Derived", QU["odm_gb"][3], "廣達模型標記 Interested-party（工商時報、FTNN 引述法人）；本模型不另訂")
inp("odm_vr", "ODM 每櫃價格 VR200（廣達模型）", "US$M/架", QU["odm_vr"][0], QU["odm_vr"][1], QU["odm_vr"][2], "Derived", QU["odm_vr"][3], "同上")
inp("vpg_28", "每 GW ODM 整櫃價值 2028 ÷ 2027（Rubin Ultra；廣達模型）", "x", QU["vpg_28"][0], QU["vpg_28"][1], QU["vpg_28"][2], "Assumed", QU["vpg_28"][3], "廣達模型的 Assumed：Rubin Ultra 每櫃價格與每 GW 架數未找到")
inp("vpg_29", "每 GW ODM 整櫃價值 2029 ÷ 2027（廣達模型）", "x", QU["vpg_29"][0], QU["vpg_29"][1], QU["vpg_29"][2], "Assumed", QU["vpg_29"][3], "同上")
for y in (27, 28, 29):
    k = f"cs_{y}"
    inp(k, f"客供比例 20{y}（緯創本體 AI 系統；廣達模型）", "%", QU[k][0], QU[k][1], QU[k][2], "Assumed", QU[k][3],
        "緯創：與新客戶議定較好的價格條件，GB／Blackwell 新計價下營益率維持合理（2Q26 法說，鉅亨）；客供比例的量化：未找到" if y == 27 else "同上")
inp("keep", "客供部分保留的毛利比例（相對買賣模式；廣達模型）", "x", QU["keep"][0], QU["keep"][1], QU["keep"][2], "Assumed", QU["keep"][3], "同廣達模型")
inp("rel_27", "緯創本體 AI 系統有效份額 2027 ÷ 2026（相對值）", "x", *SIL.vals("asm_wirel_27"), "Assumed", CNY2Q + "：2H26 起新企業與 CSP 客戶，含 L10 與 L11 整櫃（公司）", UDN0910 + "：美系券商 8 月調查 鴻海約 2,600 櫃、廣達 2,500–2,550、緯創 2,300–2,400（Interested-party，單一來源）")
inp("rel_28", "緯創本體份額 2028 ÷ 2026", "x", *SIL.vals("asm_wirel_28"), "Assumed", "同上；鴻海、廣達同時擴產（廣達模型 0.95、鴻海模型 1.0）", "—")
inp("rel_29", "緯創本體份額 2029 ÷ 2026", "x", *SIL.vals("asm_wirel_29"), "Assumed", "同上", "—")
inp("wrel_27", "緯穎 AI 有效份額 2027 ÷ 2026（相對值）", "x", 0.85, 1.0, 1.2, "Assumed", FUGW + "：AI 伺服器約九成為 ASIC 專案；2H26 新 ASIC 與 GPU 平台開始出貨（公司）", KGIW + "：凱基 2027F 緯穎營收 +74%（Interested-party）")
inp("wrel_28", "緯穎 AI 份額 2028 ÷ 2026", "x", 0.75, 1.0, 1.2, "Assumed", "同上（ASIC 顆數占 tsmc 成品的比重 2026–28 上升）", "—")
inp("wrel_29", "緯穎 AI 份額 2029 ÷ 2026", "x", 0.70, 0.95, 1.2, "Assumed", "同上", "—")

sec("C. GPU 基板與運算板（緯創本體：GPU 顆數 × 每顆基板金額 × 份額；FY2026 以券商「AI 伺服器零組件超過千億」校準）")
inp("brd_26", "FY2026 GPU 基板與運算板營收（AI 伺服器零組件）", "NT$B", 100, 110, 140, "Interested-party", KGI05 + "：2026 AI 伺服器零組件營收超過 NT$1,000 億，由運算板（Bianca／Strata）帶動（凱基 2026-05-08）", "公司：零組件（L10）出貨量逐季增加（2Q26 法說，鉅亨）；金額公司未揭露：未找到")
inp("bsh_26", "緯創占 NVIDIA＋AMD GPU 基板的份額 2026（只用於換算每顆金額）", "%", 0.35, 0.50, 0.70, "Assumed", "緯創為 NVIDIA HGX 基板與運算板主要供應商（多家媒體）；份額量化：未找到", "FY2026 營收由 brd_26 固定，份額只影響每顆金額的顯示")
for y, v, lo, hi in ((27, 1.15, 1.0, 1.4), (28, 1.2, 1.0, 1.5), (29, 1.2, 1.0, 1.5)):
    inp(f"bv_{y}", f"每顆 GPU 基板金額 20{y} ÷ 2026（Rubin 世代）", "x", lo, v, hi, "Assumed", "Rubin 每顆功耗、HBM 與 I/O 增加，基板層數與面積上升（Assumed）" if y == 27 else "同上",
        "每顆基板金額：未找到" if y == 27 else "—")
for y, v, lo, hi in ((27, 1.0, 0.8, 1.1), (28, 0.95, 0.7, 1.1), (29, 0.9, 0.65, 1.1)):
    inp(f"brel_{y}", f"GPU 基板份額 20{y} ÷ 2026", "x", lo, v, hi, "Assumed", KGI05 + "：AMD 需求帶動零組件營收；NVIDIA LPU 基板專案（凱基）" if y == 27 else "同上",
        "第二供應商（鴻海、廣達）分食：未找到量化" if y == 27 else "—")

sec("D. 營收結構錨點與非 AI 成長")
inp("nsv_26", "2026 年非伺服器（筆電、桌機、顯示器等）÷ 合併營收", "%", 0.13, 0.15, 0.18, "Interested-party", KGI05 + "：2026F 伺服器占營收 85%（凱基）", CNY2Q + "：2H26 PC 出貨約減 15%（1H 提前拉貨），營收降幅較小（公司）")
inp("gen_26", "FY2026 緯創本體一般伺服器與交換器營收", "NT$B", 80, 150, 250, "Assumed", KGI05 + "：交換器 2026 營收成長十倍以上至 NT$160–180 億；一般伺服器需求優於預期（凱基）", "本體一般伺服器金額：未找到；本體 AI 系統＝本體營收 − 基板 − 一般伺服器 − 非伺服器（殘差）")
inp("wai_26", "2026 年緯穎 AI 伺服器 ÷ 緯穎營收", "%", *SIL.vals("asm_wyai_26"), "Interested-party", "siliconomics｜SIF_asm_wyai_26（CF09：與緯穎模型同一值，以緯穎模型的公司說明與凱基 46% 為準）", FUGW + "：2025 AI 約 50%、2026 略高於 50%（緯穎 4Q25 法說；v0.6 前本模型取 55%）")
inp("g_gen27", "本體一般伺服器與交換器年增 2027", "%", 0.0, 0.20, 0.40, "Assumed", CNY2Q + "：交換器 2027 年新客戶、持續成長（公司）", "—")
inp("g_gen28", "本體一般伺服器與交換器年增 2028", "%", -0.05, 0.10, 0.25, "Assumed", "同上", "—")
inp("g_gen29", "本體一般伺服器與交換器年增 2029", "%", -0.05, 0.08, 0.20, "Assumed", "同上", "—")
inp("g_pc27", "非伺服器（筆電等）年增 2027", "%", -0.15, -0.03, 0.05, "Assumed", CNY2Q + "：記憶體漲價主要影響 PC 與手機（公司）", "—")
inp("g_pc28", "非伺服器年增 2028", "%", -0.10, 0.0, 0.05, "Assumed", "同上", "—")
inp("g_pc29", "非伺服器年增 2029", "%", -0.10, 0.0, 0.05, "Assumed", "同上", "—")
inp("g_wn27", "緯穎非 AI（一般伺服器）年增 2027", "%", 0.0, 0.10, 0.25, "Assumed", FUGW + "：2026 一般伺服器出貨 +20–30%（緯穎）", "—")
inp("g_wn28", "緯穎非 AI 年增 2028", "%", -0.05, 0.06, 0.20, "Assumed", "同上", "—")
inp("g_wn29", "緯穎非 AI 年增 2029", "%", -0.05, 0.05, 0.15, "Assumed", "同上", "—")

sec("E. 毛利率與緯穎利潤率")
inp("d_ai", "本體 AI 系統毛利率（毛額基礎）減非伺服器毛利率", "pt", -0.02, -0.01, 0.0, "Assumed", "本體毛利率 2Q26 約 3.9%（合併 − 緯穎推得）；整櫃含轉手 GPU，毛利率低", "部門別毛利率未揭露：未找到")
inp("d_brd", "GPU 基板毛利率減非伺服器毛利率", "pt", 0.01, 0.03, 0.05, "Assumed", "基板多為客供 GPU（不含 GPU 金額），加工價值占比較高（Assumed）", "未找到")
inp("d_gen", "一般伺服器與交換器毛利率減非伺服器毛利率", "pt", 0.0, 0.015, 0.03, "Assumed", "交換器與一般伺服器不含轉手 GPU（Assumed）", "未找到")
inp("vrd", "VR 世代起本體 AI 系統毛利率變動（FY2027 起）", "pt", -0.005, -0.002, 0.0, "Assumed", "與鴻海模型同值；廣達模型 −0.3 個百分點", "公司：營益率可望逐季、逐年改善（2Q26 法說）——為 Interested-party，本版不採上修")
inp("wgm", "緯穎毛利率（FY2027 起）", "%", 0.075, 0.085, 0.095, "Interested-party", FUGW + "：FY2025 8.3%；" + UANW2Q + "：1H26 8.41%、2Q26 9.26%（緯穎）", "只影響合併毛利率顯示；EPS 用營益率 wopm")
inp("wopm", "緯穎營業利益率（FY2027 起）", "%", 0.055, 0.065, 0.075, "Assumed", "FY2025 6.7%、1H26 6.79%、2Q26 7.27%（緯穎）", KGIW + "：凱基隱含 2027F 淨利率 4.4%（＝營益率約 5.7%）")
inp("wnio", "緯穎淨利 ÷ 營業利益", "%", 0.74, 0.78, 0.82, "Derived", "FY2025 51.118 ÷ 63.918＝0.80；1H26 29.084 ÷ 37.66＝0.77（緯穎）", "—")

sec("F. FY2026 錨點（1H26 實績＋3Q26 營收實績（月營收）＋假設）")
inp("q3_rev", "3Q26 合併營收（7–9 月月營收合計）", "NT$B", None, "=308.219+460.065+380.806", None, "Verified", TN07 + "：7 月 3,082.19 億；" + TN08 + "：8 月 4,600.65 億；" + TN09 + "：9 月 3,808.06 億（公司月營收）", UAN09 + "：3Q26 約 1.15 兆、前三季 2.89 兆（年增 97%）")
inp("q4g", "4Q26 合併營收季增", "%", -0.08, 0.05, 0.15, "Interested-party", SA_F + "：共識 FY2026 營收 4.1 兆（4.0–4.2 兆，7 家）隱含 4Q 約 +5%", UAN09 + "：AI 伺服器逐季成長、4Q 為今年高峰；筆電季減（公司）")
inp("q3_wrev", "3Q26 緯穎營收（前 9 月 − 1H26）", "NT$B", None, "=989.278-554.66", None, "Verified", TNW09 + "：前 9 月 9,892.78 億、9 月 1,726.20 億（緯穎月營收）", UANW2Q + "：1H26 5,546.6 億（緯穎）")
inp("q4gw", "4Q26 緯穎營收季增", "%", -0.10, 0.05, 0.15, "Assumed", "9 月 1,726 億高於 3Q 月均約 1,449 億（緯穎月營收）", "—")
inp("q_wgm", "3Q、4Q26 緯穎毛利率", "%", 0.080, 0.088, 0.095, "Interested-party", "1H26 8.41%、2Q26 9.26%（緯穎）", "—")
inp("q_wopm", "3Q、4Q26 緯穎營業利益率", "%", 0.060, 0.068, 0.075, "Interested-party", "1H26 6.79%、2Q26 7.27%（緯穎）", "新平台放量初期利潤率較低（Assumed）")
inp("q_wnio", "3Q、4Q26 緯穎淨利 ÷ 營業利益", "%", 0.74, 0.77, 0.80, "Derived", "1H26 0.77（緯穎）", "—")
inp("q_sgm", "3Q、4Q26 緯創本體毛利率", "%", 0.036, 0.039, 0.042, "Derived", "2Q26 約 3.89%、1H26 約 3.91%（合併 − 緯穎推得）", "公司：營益率可望逐季改善（2Q26 法說）")
inp("q_sopx", "3Q、4Q26 緯創本體營業費用 ÷ 本體營收", "%", 0.016, 0.0175, 0.019, "Derived", "2Q26 約 1.74%（本體毛利 − 本體營業利益）", "凱基：1Q26 本體營益率約 2%")
inp("q_snio", "3Q、4Q26 緯創本體淨利 ÷ 本體營業利益（含業外、稅、其他少數股權）", "%", 0.45, 0.56, 0.65, "Derived", "1H26 約 0.54（1Q26 匯損 13 億，凱基）、FY2025 約 0.56（Q 頁推得）", "業外與稅前利益：未找到（公開摘要未列）")
inp("stake", "緯創持有緯穎比例（2026 起）", "%", 0.35, 0.3768, 0.43, "Interested-party", KGIW + "：緯創持有緯穎 37.68%（凱基 2026-07-28）", CM2024 + "：2024 年為 42.66%（CMoney）；2026 年緯穎現金增資最多 1,800 萬股（富果）")
inp("stake25", "緯創持有緯穎比例（FY2025 拆分用）", "%", 0.3768, 0.40, 0.4266, "Assumed", "介於 2024 年 42.66% 與 2026 年 37.68% 之間", "只影響 FY2025 本體淨利的拆分")
inp("q_sh3", "3Q26 加權股數（GDR 9-10 發行 2.5 億股）", "十億股", None, "=3.18+0.25*21/92", None, "Derived", SD_GDR + "：GDR 2.5 億股、每股約 NT$186.24、募資約 NT$460 億、稀釋 7.29%、9 月 10 日發行", SA_Q + "：發行前流通股數 3.18B（stockanalysis）")
inp("q_sh4", "4Q26 起股數", "十億股", None, "=3.18+0.25", None, "Derived", "同上", "—")

sec("G. 營業費用、淨利係數、股數（FY2027 起）")
inp("phi", "緯創本體營業費用成長對本體營收成長的彈性", "x", 0.15, 0.3, 0.6, "Derived", "FY2026 錨點實測（Q 頁）", "本體營業費用率 FY2025 約 3.3% → 2Q26 約 1.74%")
inp("opx_adj", "本體營業費用額外占本體營收比（Sens 用；基準 0）", "pt", -0.002, 0.0, 0.002, "Assumed", "無", "本體營益率約 2%，±0.2 個百分點約 ±10% 本體營業利益")
inp("snio", "緯創本體淨利 ÷ 本體營業利益（FY2027 起）", "%", 0.45, 0.56, 0.65, "Derived", "同 q_snio", "—")
inp("sh_g", "股數年變動（FY2027 起）", "%", 0.0, 0.0, 0.01, "Assumed", "GDR 後 3.43B（Q 頁）", "—")
inp("mf", "AI 部署係數（Sens 用；FY2027 起）", "x", 0.8, 1.0, 1.2, "Assumed", "網通模型 2027 年 AI IT 資本支出約 2 倍", "±20%")

sec("H. 共識、市場資料（只作 Recon 對帳，不作校準）")
inp("cons_rev26", "共識 FY2026 營收", "NT$B", 4000, 4100, 4200, "Interested-party", SA_F + "：7 家平均 4.1 兆（區間 4.0–4.2 兆；2026-09-16 擷取）", "—")
inp("cons_eps26", "共識 FY2026 EPS", "NT$", 16.14, 17.09, 18.14, "Interested-party", SA_F + "：7 家平均 17.09（2026-09-16 擷取）", WR0810 + "：美系外資與本土投顧 16.67–17.65（2026-08-10）")
inp("b_eps27", "法人 FY2027 EPS（下限說法）", "NT$", None, 20.0, None, "Interested-party", WR0810 + "：法人報告估 2027 EPS 可望超過 20 元（未具名）", "較樂觀說法 23.3–26.4（同文，未具名）")
inp("b_eps27h", "法人 FY2027 EPS（樂觀說法中點）", "NT$", 23.3, 24.85, 26.4, "Interested-party", WR0810 + "（未具名）", "共識 FY2027（付費）：未找到")
inp("k_eps26", "凱基 FY2026 EPS（2026-05-08，2Q 前）", "NT$", None, 13.96, None, "Interested-party", KGI05, "目標價 225＝15x 2026–27 平均 EPS")
inp("k_eps27", "凱基 FY2027 EPS（2026-05-08）", "NT$", None, 18.29, None, "Interested-party", KGI05, "—")
inp("k_racks", "凱基：2026 年 GB 與 VR 出貨櫃數", "千櫃", None, 15, None, "Interested-party", KGI05 + "：AI 伺服器系統（L10）營收倍增以上，GB 與 VR 約 1.5 萬櫃", "1Q26 約 3.9–4.0 千櫃、2Q26F 4.0–4.2 千櫃（凱基）")
inp("rk_sh", "三家 ODM 中緯創的 NVL72 櫃數份額（8 月調查）", "%", None, "=2350/(2600+2525+2350)", None, "Interested-party", "經濟日報 2026-09-10 引述美系券商（網址見 Sources）", "櫃數口徑；單月、三家合計（不含 Dell、Supermicro 等）")
inp("kw_rev26", "凱基緯穎 FY2026 營收（07-28）", "NT$B", None, 1361.998, None, "Interested-party", KGIW, "—")
inp("kw_rev27", "凱基緯穎 FY2027 營收（07-28）", "NT$B", None, 2368.079, None, "Interested-party", KGIW, "—")
inp("kw_ni27", "凱基緯穎 FY2027 淨利（07-28）", "NT$B", None, 105.263, None, "Interested-party", KGIW, "—")

sec("I. 價格與市場倍數")
inp("px", "股價（2026-10-22 收盤）", "NT$", None, 190, None, "Interested-party", SD + "：10-22 收盤 190（−2.56%）（財報狗）", "財報狗：10-14 收盤 187；Goodinfo 頁面同 190")
inp("pe_ttm", "近 12 月 P/E（財報狗，10-22）", "x", None, 15.09, None, "Interested-party", SD + "：近四季 EPS 12.59", "Goodinfo 目前 P/E 14.93")
inp("fpe", "前瞻 P/E（stockanalysis，09-29）", "x", None, 9.93, None, "Interested-party", SA_Q + "：價格 185.5（2026-09-29 擷取）", "—")
inp("cons_feps", "共識前瞻 12 月 EPS（09-29 價格 ÷ 前瞻 P/E）", "NT$", None, "=185.5/fpe", None, "Derived", "stockanalysis 前瞻 P/E 9.93、價格 185.5（2026-09-29）", "共識前瞻 EPS 的直接數字：未找到")
inp("cons_tp", "市場共識目標價（賣方平均）", "NT$", 165, 268.71, 400, "Interested-party", SA_F + "：17 位分析師平均 268.71、中位數 275、區間 165–400（2026-09-16 擷取）", CTEE0805 + "：本國投顧目標價 300；" + WR0810 + "：240–243")
inp("nim_5y", "近五年平均淨利率（FY2021–25，歸屬母公司淨利 ÷ 合併營收；只作參考）", "%", None, "=AVERAGE(10.468/862.083,11.162/984.619,11.472/867.057,17.446/1049.256,27.408/2186.523)", None, "Derived",
    SA_FIN + "：FY2021–25 營收與淨利（stockanalysis）", "十年平均（2016–2025）：未找到；收斂法改用最後一年 × 0.8（工作單規則）")

sec("J. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_r", "折現率／要求報酬", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%；美、台、日、韓同一值）", "DCF 折現；情境價值推到 12 個月後")
inp("tp_g", "長期成長率", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "DCF 永續成長與 FY2032 收斂目標（情境 B）")
inp("tp_roe", "ROE（近三年平均，FY2023–25）", "%", None, "=AVERAGE(0.1477,0.1942,0.2626)", None, "Derived", SA_R + "：FY2023 14.77%、FY2024 19.42%、FY2025 26.26%（stockanalysis）", GI + "：Goodinfo 同（14.8%、19.4%、26.3%）")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（Andy 2026-10-25 11:53 定 50／50）", "—")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（FY2029 與「十年平均替代值」的中點＝FY2029 × 0.9）", "%", None, 0, None, "Assumed", "十年平均未找到，以 FY2029 × 0.8 代替（工作單規則）；PL 頁情境 B FY2029", "近五年平均 nim_5y 只作參考")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（十年平均未找到：FY2029 × 0.8）", "%", None, 0, None, "Assumed", "工作單規則：找不到十年平均用最後一年 × 0.8", "PL 頁情境 C FY2029")
TP_KEYS += ["tp_r", "tp_g", "tp_roe", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]

URLNAME = {CNY2Q: "緯創 2Q26 法說（鉅亨）", UAN2Q: "緯創 2Q26 法說（優分析）", UAN25: "緯創 2025 財報（優分析）", KGI05: "凱基緯創報告 2026-05-08",
           KGIW: "凱基緯穎報告 2026-07-28", FUGW: "緯穎 4Q25 法說（富果）", UANW2Q: "緯穎 2Q26 財報（優分析）", TN07: "TechNews 緯創 7 月營收",
           TN08: "TechNews 緯創 8 月營收", TN09: "TechNews 緯創 9 月營收", UAN09: "優分析 2026-10-08", TNW09: "TechNews 緯穎 9 月營收",
           SD_GDR: "財報狗 GDR 新聞", UDN0910: "經濟日報 2026-09-10", SA_Q: "stockanalysis 報價頁", SA_F: "stockanalysis 預估頁",
           SA_R: "stockanalysis 比率頁", SA_FIN: "stockanalysis 財報頁", SD: "財報狗個股頁", GI: "Goodinfo 經營績效", WR0810: "旺得富 2026-08-10",
           CTEE0805: "工商時報 2026-08-05", CM2024: "CMoney 2024-08-01"}
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
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億新台幣；EPS 為新台幣；股數為十億股）。藍字＝輸入；黑字＝公式。合併＝緯創本體＋緯穎；本體＝合併 − 緯穎（含其他子公司，集團內交易未扣除）"
Q["A2"] = "FY2026 錨點＝1H26 實績＋3Q26 營收實績（緯創、緯穎月營收）＋3Q、4Q 假設（緯穎毛利率與營益率、本體毛利率與費用率、淨利係數；4Q 營收＝3Q ×（1＋季增））。1Q26 營收＝2Q26 ÷（1＋季增 5.81%）；1Q 毛利率、營益率由 2Q 與季增百分點推得。歸屬母公司淨利＝本體淨利＋持股 × 緯穎淨利；非控制權益＝（1 − 持股）× 緯穎淨利（其他子公司的少數股權併入本體係數）"
hdr(Q, 3, ["項目", "FY2025A", "1H26（實績）", "3Q26（營收實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1（公司）", "來源2／說明"])
S1 = "FY2025：stockanalysis 財報頁、優分析 2026-03-12；1H26：鉅亨、優分析 2026-08-04（2Q26 法說）"
S2 = "3Q26：月營收公告（TechNews 2026-08-07、09-09、10-08）"
SW = "緯穎：富果 FY2025 法說、優分析 2Q26 財報、TechNews 9 月營收"
Q1R = "895.443/(1+0.0581)"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2=S2, ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"F{r}"] = ann or f"=SUM(C{r}:E{r})"; Q[f"F{r}"].number_format = fmt
    Q.cell(r, 7, tag); Q.cell(r, 8, s1); Q.cell(r, 9, s2)
qrow(4, "合併營收", {"B": 2186.523, "C": f"={Q1R}+895.443", "D": "=q3_rev", "E": "=D4*(1+q4g)"}, tag="Verified／Interested-party（4Q 季增）", s2="FY2025 2,186.523（stockanalysis；公司 2.19 兆）；2Q26 8,954.43 億、季增 5.81%（公司）")
qrow(5, "合併毛利", {"B": 133.983, "C": f"={Q1R}*(0.0556-0.0045)+895.443*0.0556", "D": "=D10*q_wgm+D13*q_sgm", "E": "=E10*q_wgm+E13*q_sgm"}, tag="Verified／Derived（毛利率 × 營收）",
     s2="FY2025 毛利率 6.13%；2Q26 5.56%、季增 0.45 個百分點（公司）")
qrow(6, "合併毛利率", {c: f"={c}5/{c}4" for c in "BCDEF"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2="—")
qrow(7, "合併營業利益", {"B": 78.553, "C": f"={Q1R}*(0.0374-0.0029)+895.443*0.0374", "D": "=D12+D15", "E": "=E12+E15"}, tag="Verified／Derived",
     s2="FY2025 營益率 3.59%；2Q26 3.74%、季增 0.29 個百分點（公司）")
qrow(8, "合併營業利益率", {c: f"={c}7/{c}4" for c in "BCDEF"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2="—")
qrow(9, "合併營業費用", {c: f"={c}5-{c}7" for c in "BCDE"}, tag="Derived", s1="毛利 − 營業利益", s2="—")
qrow(10, "緯穎營收", {"B": 950.663, "C": 554.66, "D": "=q3_wrev", "E": "=D10*(1+q4gw)"}, tag="Verified／Assumed（4Q）", s1=SW, s2="FY2025 9,506.63 億；1H26 5,546.6 億（緯穎）")
qrow(11, "緯穎毛利", {"B": 78.454, "C": "=554.66*0.0841", "D": "=D10*q_wgm", "E": "=E10*q_wgm"}, tag="Verified／Derived", s1=SW, s2="FY2025 784.54 億；1H26 毛利率 8.41%")
qrow(12, "緯穎營業利益", {"B": 63.918, "C": "=554.66*0.0679", "D": "=D10*q_wopm", "E": "=E10*q_wopm"}, tag="Verified／Derived", s1=SW, s2="FY2025 639.18 億；1H26 營益率 6.79%")
qrow(13, "緯創本體營收（合併 − 緯穎）", {c: f"={c}4-{c}10" for c in "BCDE"}, tag="Derived", s1="—", s2="集團內交易：未找到（假設為 0）")
qrow(14, "緯創本體毛利", {c: f"={c}5-{c}11" for c in "BCDE"}, tag="Derived", s1="—", s2="—")
qrow(15, "緯創本體營業利益", {"B": "=B7-B12", "C": "=C7-C12", "D": "=D14-D13*q_sopx", "E": "=E14-E13*q_sopx"}, tag="Derived", s1="—", s2="凱基：1Q26 本體營益率約 2%、緯穎 6.3%")
qrow(16, "緯創本體營業費用", {c: f"={c}14-{c}15" for c in "BCDE"}, tag="Derived", s1="—", s2="—")
qrow(17, "緯穎淨利", {"B": 51.118, "C": 29.084, "D": "=D12*q_wnio", "E": "=E12*q_wnio"}, tag="Verified／Derived", s1=SW, s2="FY2025 511.18 億；1H26 290.84 億（緯穎）")
qrow(18, "緯創持有緯穎比例", {"B": "=stake25", "C": "=stake", "D": "=stake", "E": "=stake"}, fmt="0.00%", ann="=stake", tag="Interested-party／Assumed（FY2025）", s1="凱基 2026-07-28：37.68%", s2="2024 年 42.66%（CMoney）")
qrow(19, "非控制權益（≈（1 − 持股）× 緯穎淨利）", {c: f"=(1-{c}18)*{c}17" for c in "BCDE"}, tag="Derived", s1="其他子公司的少數股權併入本體係數", s2="公司揭露的非控制權益金額：未找到")
qrow(20, "歸屬母公司淨利", {"B": 27.408, "C": "=9.631+14.827", "D": "=D21+D18*D17", "E": "=E21+E18*E17"}, tag="Verified／Derived（3Q、4Q）", s2="FY2025 274.08 億；1Q26 96.31 億、2Q26 148.27 億（公司稅後純益）")
qrow(21, "緯創本體淨利（歸屬母公司淨利 − 持股 × 緯穎淨利）", {"B": "=B20-B18*B17", "C": "=C20-C18*C17", "D": "=D15*q_snio", "E": "=E15*q_snio"}, tag="Derived", s1="—", s2="—")
qrow(22, "合併淨利（含非控制權益）", {c: f"={c}20+{c}19" for c in "BCDE"}, tag="Derived", s1="—", s2="—")
qrow(23, "股數（十億股；加權）", {"B": "=27.408/9.04", "C": "=24.458/7.78", "D": "=q_sh3", "E": "=q_sh4"}, ann="=(2*C23+D23+E23)/4", tag="Derived", s2="淨利 ÷ EPS；3Q 起含 GDR 2.5 億股")
qrow(24, "EPS", {c: f"={c}20/{c}23" for c in "BCDE"}, fmt="0.00", ann="=F20/F23", tag="Derived", s1="淨利 ÷ 股數", s2="—")
qrow(25, "  公司揭露 EPS", {"B": 9.04, "C": 7.78}, fmt="0.00", ann="", tag="Verified", s2="1Q26 3.06、2Q26 4.72")
Q.cell(26, 1, "FY2026 營收拆分（緯創本體與緯穎；FY2025 拆分未找到）").font = BOLD
qrow(27, "緯創本體：GPU 基板與運算板", {}, ann="=brd_26", tag="Interested-party", s1="凱基：AI 伺服器零組件超過千億", s2="—")
qrow(28, "緯創本體：一般伺服器與交換器", {}, ann="=gen_26", tag="Assumed", s1="—", s2="—")
qrow(29, "緯創本體：非伺服器（筆電、桌機、顯示器等）", {}, ann="=F4*nsv_26", tag="Interested-party", s1="凱基：伺服器 85%", s2="—")
qrow(30, "緯創本體：AI 系統（L10／L11；殘差）", {}, ann="=F13-F27-F28-F29", tag="Derived", s1="本體營收 − 基板 − 一般伺服器 − 非伺服器", s2="公司未揭露：未找到")
qrow(31, "緯穎：AI 伺服器", {}, ann="=F10*wai_26", tag="Interested-party", s1="緯穎：略高於 50%", s2="—")
qrow(32, "緯穎：非 AI（一般伺服器）", {}, ann="=F10-F31", tag="Derived", s1="—", s2="—")
Q.cell(34, 1, "FY2026 錨點的檢查（參考；校準未使用券商數字）").font = BOLD
checks = [("FY2026 合併營收年增", "=F4/B4-1", "0.0%"), ("FY2026 緯穎營收年增", "=F10/B10-1", "0.0%"), ("FY2026 本體營收年增", "=F13/B13-1", "0.0%"),
          ("FY2026 本體營業費用年增", "=F16/B16-1", "0.0%"), ("本體營業費用彈性（費用年增 ÷ 營收年增；In 頁 phi 的依據）", "=B38/B37", "0.00"),
          ("FY2026 營收 ÷ 共識 − 1", "=F4/cons_rev26-1", "0.0%"), ("FY2026 EPS ÷ 共識 − 1", "=F24/cons_eps26-1", "0.0%"),
          ("1H26 EPS 自算 − 公司 7.78（應約為 0）", "=C24-C25", "0.00"), ("FY2026 合併營業利益 ÷ 緯穎營業利益（緯穎占比的倒數）", "=F7/F12", "0.00"),
          ("隱含本體 AI 系統出貨櫃數（千櫃；FY2026 ODM 混合價格）", "=F30/fx/(odm_gb*(1-vr26)+odm_vr*vr26)", "#,##0.0")]
for i, (lab, f, fm) in enumerate(checks, 35):
    Q.cell(i, 1, lab); Q[f"B{i}"] = f; Q[f"B{i}"].number_format = fm
assert Q.cell(37, 1).value.startswith("FY2026 本體營收") and Q.cell(38, 1).value.startswith("FY2026 本體營業費用")
QCHK_LAST = 34 + len(checks)
for nm, rr in (("rev", 4), ("gp", 5), ("oi", 7), ("wrev", 10), ("wgp", 11), ("woi", 12), ("srev", 13), ("sgp", 14), ("soi", 15), ("sopx", 16),
               ("wni", 17), ("nci", 19), ("ni", 20), ("sni", 21), ("cni", 22), ("sh", 23), ("eps", 24)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$F${rr}")
for nm, rr in (("brd", 27), ("gen", 28), ("pc", 29), ("ais", 30), ("wai", 31), ("wna", 32)):
    defname(f"a26_{nm}", f"Q!$F${rr}")
Q.column_dimensions["A"].width = 60
for c in "BCDEF": Q.column_dimensions[c].width = 15
Q.column_dimensions["G"].width = 28; Q.column_dimensions["H"].width = 60; Q.column_dimensions["I"].width = 70
Q.freeze_panes = "B4"

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G", 2030: "H", 2031: "I", 2032: "J"}
FADE = (2030, 2031, 2032); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
BU = (2026, 2027, 2028, 2029); F3 = (2027, 2028, 2029)
DEFP = dict(scn="A", rel="rel_{y}", wrel="wrel_{y}", cs="cs_{y}", keep="keep", vpg="vpg_{y}", mf="1", bv="bv_{y}", brel="brel_{y}",
            ggen="g_gen{y}", gpc="g_pc{y}", gwn="g_wn{y}", d="d_ai", dbrd="d_brd", dgen="d_gen", vrd="vrd", opxa="0",
            wgm="wgm", wopm="wopm", wnio="wnio", snio="snio", stake="stake", fg="fade_g{s}", fm="fade_m{s}")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    p = lambda k, y=None: P[k].format(y=(y % 100 if y else ""), s=s)
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.0", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 11, note); R[key] = r; r += 1
    def sec_(t):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
    c_ = lambda y: FC[y]
    umx = lambda y: {2026: "1", 2027: "1", 2028: f"um28_{s}", 2029: f"um29_{s}"}[y]
    sec_("A. AI 部署與整櫃組裝市場（D 欄＝FY2026 校準年；E–G＝FY2027–FY2029；與廣達模型同口徑）")
    row("itc", "AI IT 資本支出 × 情境顆數係數", "US$B", {c_(y): f"=itc_{y % 100}*{umx(y)}" + ("" if y == 2026 else f"*{P['mf']}") for y in BU})
    row("itpg", "每 GW IT 資本支出", "US$B", {"D": "=itpg_gb*(1-vr26)+itpg_vr*vr26", **{c_(y): "=itpg_vr" for y in F3}})
    row("gw", "部署 GW（IT 電力）", "GW", {c_(y): f"={c_(y)}{R['itc']}/{c_(y)}{R['itpg']}" for y in BU}, fmt="0.0")
    row("vpg", "每 GW ODM 整櫃價值", "US$B/GW", {"D": "=(racks_gb*odm_gb*(1-vr26)+racks_vr*odm_vr*vr26)/1000", "E": "=racks_vr*odm_vr/1000",
                                                  **{c_(y): f"=$E{r}*{p('vpg', y)}" for y in (2028, 2029)}}, fmt="0.00", note="架數 × ODM 每櫃價格（讀廣達模型）")
    row("mku", "AI 伺服器組裝市場（ODM 價，美元）", "US$B", {c_(y): f"={c_(y)}{R['gw']}*{c_(y)}{R['vpg']}" for y in BU})
    row("mkn", "AI 伺服器組裝市場（新台幣）", "NT$B", {c_(y): f"={c_(y)}{R['mku']}*fx" for y in BU})
    sec_("B. 緯創本體 AI 系統（L10／L11）")
    row("sh0", "本體 AI 系統有效份額（FY2026 校準）＝ AI 系統營收 ÷ 市場", "%", {"D": f"=a26_ais/D{R['mkn']}"}, fmt="0.0%",
        note="Derived：分子為本體營收殘差；分母含 ASIC、HGX 機架與其他組裝廠，所以低於櫃數份額")
    row("rel", "份額相對值（÷ FY2026）", "x", {"D": "=1", **{c_(y): f"={p('rel', y)}" for y in F3}}, fmt="0.00")
    row("sh", "本體 AI 系統有效份額", "%", {c_(y): f"=$D${R['sh0']}*{c_(y)}{R['rel']}" for y in BU}, fmt="0.0%")
    row("aiv", "本體 AI 系統毛額價值（買賣模式下的營收）", "NT$B", {c_(y): f"={c_(y)}{R['mkn']}*{c_(y)}{R['sh']}" for y in BU})
    row("cs", "客供比例（FY2027 起新增部分）", "%", {"D": "=0", **{c_(y): f"={p('cs', y)}" for y in F3}}, fmt="0.0%", note="FY2026 的客供已含在錨點與校準份額內，記為 0")
    row("ais", "本體 AI 系統營收", "NT$B", {c_(y): f"={c_(y)}{R['aiv']}*(1-{c_(y)}{R['cs']})" for y in BU})
    sec_("C. 緯創本體 GPU 基板與運算板（GPU 顆數 × 每顆基板金額 × 份額）")
    row("gu", "GPU 顆數（NVIDIA＋AMD，含情境係數）", "百萬顆", {c_(y): f"=gu_{y % 100}*{umx(y)}" + ("" if y == 2026 else f"*{P['mf']}") for y in BU}, fmt="0.00")
    row("bsh", "緯創基板份額", "%", {"D": "=bsh_26", **{c_(y): f"=bsh_26*{p('brel', y)}" for y in F3}}, fmt="0.0%")
    row("bvu", "每顆 GPU 基板金額（緯創營收 ÷ 顆數 ÷ 份額）", "US$／顆", {"D": f"=a26_brd/fx/D{R['gu']}/D{R['bsh']}*1000", **{c_(y): f"=$D{r}*{p('bv', y)}" for y in F3}}, fmt="#,##0",
        note="FY2026 校準（Derived）；之後 × 每顆金額指數 bv（Assumed）")
    row("brd", "GPU 基板與運算板營收", "NT$B", {c_(y): f"={c_(y)}{R['gu']}*{c_(y)}{R['bsh']}*{c_(y)}{R['bvu']}*fx/1000" for y in BU})
    sec_("D. 其他事業與合計")
    row("gen", "本體一般伺服器與交換器", "NT$B", {"D": "=a26_gen", **{c_(y): f"={c_(y - 1)}{r}*(1+{p('ggen', y)})" for y in F3}})
    row("pc", "本體非伺服器（筆電、桌機、顯示器等）", "NT$B", {"D": "=a26_pc", **{c_(y): f"={c_(y - 1)}{r}*(1+{p('gpc', y)})" for y in F3}})
    row("srev", "緯創本體營收", "NT$B", {"C": "=a25_srev", **{c_(y): f"={c_(y)}{R['ais']}+{c_(y)}{R['brd']}+{c_(y)}{R['gen']}+{c_(y)}{R['pc']}" for y in BU}})
    row("wsh0", "緯穎 AI 有效份額（FY2026 校準）", "%", {"D": f"=a26_wai/D{R['mkn']}"}, fmt="0.0%", note="Derived：緯穎營收 × AI 占比 ÷ 市場（緯穎 AI 約九成為 ASIC）")
    row("wsh", "緯穎 AI 有效份額", "%", {"D": f"=D{R['wsh0']}", **{c_(y): f"=$D${R['wsh0']}*{p('wrel', y)}" for y in F3}}, fmt="0.0%")
    row("wai", "緯穎 AI 伺服器營收", "NT$B", {c_(y): f"={c_(y)}{R['mkn']}*{c_(y)}{R['wsh']}" for y in BU})
    row("wna", "緯穎非 AI（一般伺服器）", "NT$B", {"D": "=a26_wna", **{c_(y): f"={c_(y - 1)}{r}*(1+{p('gwn', y)})" for y in F3}})
    row("wrev", "緯穎營收", "NT$B", {"C": "=a25_wrev", **{c_(y): f"={c_(y)}{R['wai']}+{c_(y)}{R['wna']}" for y in BU}})
    row("rev", "合併營收", "NT$B", {c_(y): f"={c_(y)}{R['srev']}+{c_(y)}{R['wrev']}" for y in FYS})
    row("srv", "伺服器營收（本體 AI 系統＋GPU 基板＋本體一般伺服器與交換器＋緯穎）", "NT$B", {c_(y): f"={c_(y)}{R['ais']}+{c_(y)}{R['brd']}+{c_(y)}{R['gen']}+{c_(y)}{R['wrev']}" for y in BU}, note="公司產品別揭露：伺服器（凱基：2026F 伺服器約占 85%）")
    row("npc", "非伺服器營收（筆電、桌機、顯示器等）", "NT$B", {c_(y): f"={c_(y)}{R['pc']}" for y in BU}, note="公司產品別揭露：非伺服器")
    row("segchk", "檢查：分部合計 − 合併營收（FY2026 錨點起）", "NT$B", {c_(y): f"={c_(y)}{R['srv']}+{c_(y)}{R['npc']}-{c_(y)}{R['rev']}" for y in BU}, fmt="0.000")
    row("srvsh", "伺服器 ÷ 合併營收", "%", {c_(y): f"={c_(y)}{R['srv']}/{c_(y)}{R['rev']}" for y in BU}, fmt="0.0%")
    row("g", "合併營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("aish", "AI（本體 AI 系統＋基板＋緯穎 AI）÷ 合併營收", "%", {c_(y): f"=({c_(y)}{R['ais']}+{c_(y)}{R['brd']}+{c_(y)}{R['wai']})/{c_(y)}{R['rev']}" for y in BU}, fmt="0.0%")
    row("wshr", "緯穎 ÷ 合併營收", "%", {c_(y): f"={c_(y)}{R['wrev']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("rk", "隱含本體出貨櫃數（毛額價值 ÷ ODM 每櫃價格；NVL72 等值）", "千櫃",
        {"D": f"=D{R['aiv']}/fx/(odm_gb*(1-vr26)+odm_vr*vr26)", **{c_(y): f"={c_(y)}{R['aiv']}/fx/odm_vr" for y in F3}}, note="十億美元 ÷ 百萬美元 ＝ 千櫃")
    sec_("E. 緯創本體損益")
    row("g0", "非伺服器毛利率（FY2026 校準，其他部門＝此值＋差距）", "%",
        {"D": f"=(a26_sgp-{P['d']}*D{R['ais']}-{P['dbrd']}*D{R['brd']}-{P['dgen']}*D{R['gen']})/D{R['srev']}"}, fmt="0.00%", note="Derived：讓 FY2026 本體毛利等於錨點")
    row("gai", "本體 AI 系統毛利 ＝ 毛利率 × 毛額價值 ×（1 − 客供 ×（1 − 保留比例））", "NT$B",
        {c_(y): f"=($D${R['g0']}+{P['d']}" + ("" if y == 2026 else f"+{P['vrd']}") + f")*{c_(y)}{R['aiv']}*(1-{c_(y)}{R['cs']}*(1-{P['keep']}))" for y in BU}, fmt="#,##0.00")
    row("sgp", "本體毛利", "NT$B", {"C": "=a25_sgp", "D": "=a26_sgp", **{c_(y): f"=$D${R['g0']}*{c_(y)}{R['pc']}+($D${R['g0']}+{P['dbrd']})*{c_(y)}{R['brd']}+($D${R['g0']}+{P['dgen']})*{c_(y)}{R['gen']}+{c_(y)}{R['gai']}" for y in F3}})
    row("sgm", "本體毛利率", "%", {c_(y): f"={c_(y)}{R['sgp']}/{c_(y)}{R['srev']}" for y in FYS}, fmt="0.00%")
    row("sgs", "本體營業費用成長（基準路徑）", "%", {c_(y): f"=MAX(0,phi*({c_(y)}{R['srev']}/{c_(y - 1)}{R['srev']}-1))" for y in F3}, fmt="0.0%")
    row("sopx", "本體營業費用", "NT$B", {"C": "=a25_sopx", "D": "=a26_sopx", **{c_(y): f"={c_(y - 1)}{r}*(1+{c_(y)}{R['sgs']})+{P['opxa']}*{c_(y)}{R['srev']}" for y in F3}})
    row("soi", "本體營業利益", "NT$B", {c_(y): f"={c_(y)}{R['sgp']}-{c_(y)}{R['sopx']}" for y in FYS})
    row("soim", "本體營業利益率", "%", {c_(y): f"={c_(y)}{R['soi']}/{c_(y)}{R['srev']}" for y in FYS}, fmt="0.00%")
    row("sni", "本體淨利（歸屬緯創，不含緯穎持股收益）", "NT$B", {"C": "=a25_sni", "D": "=a26_sni", **{c_(y): f"={c_(y)}{R['soi']}*{P['snio']}" for y in F3}})
    sec_("F. 緯穎（合併子公司）")
    row("wgp", "緯穎毛利", "NT$B", {"C": "=a25_wgp", "D": "=a26_wgp", **{c_(y): f"={c_(y)}{R['wrev']}*{P['wgm']}" for y in F3}})
    row("woi", "緯穎營業利益", "NT$B", {"C": "=a25_woi", "D": "=a26_woi", **{c_(y): f"={c_(y)}{R['wrev']}*{P['wopm']}" for y in F3}})
    row("wni", "緯穎淨利", "NT$B", {"C": "=a25_wni", "D": "=a26_wni", **{c_(y): f"={c_(y)}{R['woi']}*{P['wnio']}" for y in F3}})
    row("stk", "緯創持有緯穎比例", "%", {"C": "=stake25", **{c_(y): f"={P['stake']}" for y in BU}}, fmt="0.00%")
    sec_("G. 合併與歸屬母公司")
    row("gp", "合併毛利", "NT$B", {c_(y): f"={c_(y)}{R['sgp']}+{c_(y)}{R['wgp']}" for y in FYS})
    row("gm", "合併毛利率", "%", {c_(y): f"={c_(y)}{R['gp']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.00%")
    row("oi", "合併營業利益", "NT$B", {c_(y): f"={c_(y)}{R['soi']}+{c_(y)}{R['woi']}" for y in FYS})
    row("oim", "合併營業利益率", "%", {c_(y): f"={c_(y)}{R['oi']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.00%")
    row("cni", "合併淨利（含非控制權益）", "NT$B", {c_(y): f"={c_(y)}{R['sni']}+{c_(y)}{R['wni']}" for y in FYS})
    row("nci", "非控制權益（（1 − 持股）× 緯穎淨利）", "NT$B", {c_(y): f"=(1-{c_(y)}{R['stk']})*{c_(y)}{R['wni']}" for y in FYS})
    row("ni", "歸屬母公司淨利", "NT$B", {c_(y): f"={c_(y)}{R['cni']}-{c_(y)}{R['nci']}" for y in FYS})
    row("wcon", "緯穎對歸屬母公司淨利的貢獻 ÷ 歸屬母公司淨利", "%", {c_(y): f"={c_(y)}{R['stk']}*{c_(y)}{R['wni']}/{c_(y)}{R['ni']}" for y in FYS}, fmt="0.0%")
    row("shr", "股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", "E": "=q_sh4*(1+sh_g)", **{c_(y): f"={c_(y - 1)}{r}*(1+sh_g)" for y in (2028, 2029)}}, fmt="0.000",
        note="FY2027 起以 GDR 後股數 3.43B 為基礎")
    row("eps", "EPS", "NT$", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['shr']}" for y in FYS}, fmt="0.00")
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("nim", "淨利率（歸屬母公司淨利 ÷ 合併營收）", "%", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.00%")
    # FY2030–32：收斂法（上游只到 2029；非驅動因子推導）
    for y in FADE:
        c, pc, k = FC[y], FC[y - 1], y - 2029
        ws[f"{c}{R['g']}"] = f"=$G{R['g']}+({p('fg')}-$G{R['g']})*{k}/3"
        ws[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
        ws[f"{c}{R['nim']}"] = f"=$G{R['nim']}+({p('fm')}-$G{R['nim']})*{k}/3"
        ws[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
        ws[f"{c}{R['shr']}"] = f"={pc}{R['shr']}*(1+sh_g)"
        ws[f"{c}{R['eps']}"] = f"={c}{R['ni']}/{c}{R['shr']}"
        ws[f"{c}{R['epsg']}"] = f"={c}{R['eps']}/{pc}{R['eps']}-1"
        for key, fm in (("g", "0.0%"), ("rev", "#,##0.0"), ("nim", "0.00%"), ("ni", "#,##0.0"), ("shr", "0.000"), ("eps", "0.00"), ("epsg", "0.0%")):
            ws[f"{c}{R[key]}"].number_format = fm; ws[f"{c}{R[key]}"].fill = FADEFILL
    ws.cell(R["rev"], 11, "FY2030 起為收斂法：營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m（FY2032 達到），非驅動因子推導")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 緯創損益（十億新台幣；EPS 新台幣）。營收、毛利、營業利益為合併（含緯穎）；EPS 用歸屬母公司淨利。FY2025A 實績、FY2026 錨點（三情境相同），FY2027 起為模型。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A（順利）：顆數基準（係數 1.0）", "B": "情境 B（平平）：CY2028 顆數 ×0.8、CY2029 ×0.85", "C": "情境 C（不如預期）：CY2028–29 顆數 ×0.6"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
sr = 4; SUMR = {}
SUMK = (("rev", "合併營收", "NT$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("ais", "  本體 AI 系統", "NT$B", "#,##0.0"), ("brd", "  GPU 基板與運算板", "NT$B", "#,##0.0"),
        ("wrev", "  緯穎", "NT$B", "#,##0.0"), ("gm", "合併毛利率", "%", "0.00%"), ("oim", "合併營業利益率", "%", "0.00%"),
        ("cni", "合併淨利（含非控制權益）", "NT$B", "#,##0.0"), ("ni", "歸屬母公司淨利", "NT$B", "#,##0.0"), ("nim", "  淨利率", "%", "0.00%"),
        ("eps", "EPS", "NT$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%"))
LONGK = ("rev", "g", "ni", "nim", "eps", "epsg")
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in SUMK:
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS + (FADE if key in LONGK else ()):
            if key in ("g", "epsg") and y == 2025: continue
            if key in ("ais", "brd") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
            if y in FADE: PL[f"{FC[y]}{sr}"].fill = FADEFILL
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
IN.cell(INROW["fade_mA"], 5, f"=PL!G{SUMR[('A', 'nim')]}")
IN.cell(INROW["fade_mB"], 5, f"=PL!G{SUMR[('B', 'nim')]}*0.9")
IN.cell(INROW["fade_mC"], 5, f"=PL!G{SUMR[('C', 'nim')]}*0.8")
for k in ("fade_mA", "fade_mB", "fade_mC"): IN.cell(INROW[k], 5).font = Font(color="000000")
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型 − 錨點").font = BOLD; sr += 1
CHK = []
for lab, f in (("  三情境：合併營收差（NT$B）", "+".join(f"ABS(D{PB[s]['rev']}-a26_rev)" for s in SC)),
               ("  三情境：EPS 差（NT$）", "+".join(f"ABS(D{PB[s]['eps']}-a26_eps)" for s in SC)),
               ("  情境 A：本體毛利差（NT$B；以 FY2027 公式套用 FY2026）", f"D{PB['A']['g0']}*D{PB['A']['pc']}+(D{PB['A']['g0']}+d_brd)*D{PB['A']['brd']}+(D{PB['A']['g0']}+d_gen)*D{PB['A']['gen']}+D{PB['A']['gai']}-a26_sgp")):
    PL.cell(sr, 1, lab); PL[f"D{sr}"] = "=" + f; PL[f"D{sr}"].number_format = "0.0000"; CHK.append(sr); sr += 1
assert sr < 60, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 10; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 13
PL.cell(2, 1, "灰底＝FY2030–FY2032 收斂法（上游只到 2029）：營收年增線性收斂到 fade_g、淨利率（歸屬母公司淨利 ÷ 合併營收）線性收斂到 fade_m（In 頁 J 節），非驅動因子推導")

# ---------------- Units ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — AI 伺服器組裝市場（與廣達模型同口徑）、緯創本體與緯穎份額、GPU 基板（情境 A）。部署 GW＝AI IT 資本支出（markets/network）÷ 每 GW IT 資本支出（Tokenomics，經橋接表）；ODM 每櫃價格讀廣達模型"
A_ = PB["A"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
QSH_V = dict(zip(BU, QSH[0])); FSH_V = dict(zip(BU, FSH[0]))
UNITS_ROWS = (
    ("gw", "部署 GW（情境 A）", "GW", lambda y: f"=PL!{FC[y]}{A_['gw']}", "0.0", "AI IT 資本支出 ÷ 每 GW IT 資本支出"),
    ("mku", "AI 伺服器組裝市場（ODM 價）", "US$B", lambda y: f"=PL!{FC[y]}{A_['mku']}", "#,##0.0", "含 ASIC 機架（假設每 GW 價值相同；廣達模型口徑）"),
    ("sh", "緯創本體 AI 系統有效份額（金額口徑）", "%", lambda y: f"=PL!{FC[y]}{A_['sh']}", "0.0%", "FY2026 以本體營收殘差校準"),
    ("wsh", "緯穎 AI 有效份額", "%", lambda y: f"=PL!{FC[y]}{A_['wsh']}", "0.0%", "緯穎 AI 約九成為 ASIC（AWS Trainium 等）"),
    ("csh", "緯創合併 AI 系統份額（本體＋緯穎）", "%", lambda y: f"={cyc(y)}6+{cyc(y)}7", "0.0%", ""),
    ("qsh", "廣達有效份額（廣達模型 Units，唯讀）", "%", lambda y: QSH_V[y], "0.0%", "來源：" + QSH[1]),
    ("fsh", "鴻海有效份額（鴻海模型 Units，唯讀）", "%", lambda y: FSH_V[y], "0.0%", "來源：" + FSH[1]),
    ("tsh", "鴻海＋廣達＋緯創（含緯穎）合計有效份額", "%", lambda y: f"={cyc(y)}8+{cyc(y)}9+{cyc(y)}10", "0.0%", "其餘為英業達、Dell、Supermicro、ASIC 自有組裝等"),
    ("ais", "緯創本體 AI 系統營收", "US$B", lambda y: f"=PL!{FC[y]}{A_['ais']}/fx", "#,##0.0", "扣除客供"),
    ("rk", "隱含本體出貨櫃數（NVL72 等值）", "千櫃", lambda y: f"=PL!{FC[y]}{A_['rk']}", "#,##0.0", "毛額價值 ÷ ODM 每櫃價格"),
    ("gu", "GPU 顆數（NVIDIA＋AMD）", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['gu']}", "0.00", "tsmc 成品 In 頁"),
    ("bvu", "每顆 GPU 基板金額", "US$／顆", lambda y: f"=PL!{FC[y]}{A_['bvu']}", "#,##0", "FY2026 校準；份額 50% 為 Assumed"),
    ("brd", "GPU 基板與運算板營收", "US$B", lambda y: f"=PL!{FC[y]}{A_['brd']}/fx", "#,##0.00", "顆數 × 份額 × 每顆金額"),
    ("wai", "緯穎 AI 伺服器營收", "US$B", lambda y: f"=PL!{FC[y]}{A_['wai']}/fx", "#,##0.0", ""))
for key, lab, unit, f, fmt, note in UNITS_ROWS:
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
        if key in ("qsh", "fsh"): UN[f"{cyc(y)}{ur}"].font = BLUE
    UN.cell(ur, 7, note); UR[key] = ur; ur += 1
assert (UR["sh"], UR["wsh"], UR["csh"], UR["qsh"], UR["fsh"]) == (6, 7, 8, 9, 10), UR
ur += 1
UN.cell(ur, 1, "讀法：有效份額為金額口徑（分母為全部 AI IT 資本支出換算的組裝市場，含 ASIC、HGX 與其他組裝廠），低於櫃數份額。GPU 基板只對應 NVIDIA 與 AMD 的 GPU 顆數；基板份額與每顆金額是 Assumed，FY2026 金額取券商說法。緯穎以 ASIC 機架為主，所以份額跟著 ASIC 部署走。").font = BOLD
UN_NOTE = ur
UN.column_dimensions["A"].width = 52; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 13
UN.column_dimensions["G"].width = 60

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣；EPS 為歸屬母公司）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，年底落後 P/E）只作參考"
V["A2"] = "循環位置：緯創不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照廣達、鴻海、NVIDIA；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
O = 51
V.cell(O, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, O + 1, ["項目", "A", "B", "C", "說明"])
V.cell(O + 2, 1, "現價（2026-10-22 收盤）"); V[f"B{O + 2}"] = "=px"; V[f"B{O + 2}"].number_format = "#,##0.0"
for k, y, lab in ((3, 2026, "現價 ÷ FY2026 EPS（錨點）"), (4, 2027, "現價 ÷ FY2027 EPS（模型）"), (5, 2028, "現價 ÷ FY2028 EPS（模型）"), (6, 2029, "現價 ÷ FY2029 EPS（模型）")):
    V.cell(O + k, 1, lab)
    for j, s in enumerate(SC):
        c = "BCD"[j]; V[f"{c}{O + k}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{O + k}"].number_format = "0.0"
V.cell(O + 7, 1, "未來 12 月 EPS（模型 B）≈ 0.2 × FY2026＋0.8 × FY2027"); V[f"B{O + 7}"] = f"=0.2*PL!D{SUMR[('B', 'eps')]}+0.8*PL!E{SUMR[('B', 'eps')]}"; V[f"B{O + 7}"].number_format = "0.00"
V.cell(O + 8, 1, "  現價 ÷ 未來 12 月 EPS（模型 B）"); V[f"B{O + 8}"] = f"=px/B{O + 7}"; V[f"B{O + 8}"].number_format = "0.0"
V.cell(O + 9, 1, "近 12 月 P/E（財報狗，10-22）"); V[f"B{O + 9}"] = "=pe_ttm"; V[f"B{O + 9}"].number_format = "0.0"
V.cell(O + 10, 1, "前瞻 P/E（stockanalysis，09-29）"); V[f"B{O + 10}"] = "=fpe"; V[f"B{O + 10}"].number_format = "0.0"
V[f"E{O + 3}"] = "FY2026 三情境相同（錨點）"; V[f"E{O + 4}"] = "FY2027 三情境相同（顆數係數從 CY2028 開始）"
CY0 = O + 12
V.cell(CY0, 1, "2. 循環位置（本模型營收最高年；FY(t)＝CY(t)）").font = BOLD
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
V.cell(R30, 1, "3. 規則值隱含股價（新台幣；約 FY2027 年底，未折現）＝ FY2027 EPS × 規則倍數").font = BOLD
hdr(V, R30 + 1, ["情境", "FY2027 EPS", "（同左）", "規則倍數", "隱含股價", "相對現價", "現價 ÷ FY2027 EPS"])
VAL = {}
for j, s in enumerate(SC):
    i = R30 + 2 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 3, f"=B{i}").number_format = "0.00"
    V.cell(i, 4, f"=D{CYR[s]}").number_format = "0.0"
    V.cell(i, 5, f"=C{i}*D{i}").number_format = "#,##0"
    V.cell(i, 6, f"=E{i}/px-1").number_format = "0%"
    V.cell(i, 7, f"=px/C{i}").number_format = "0.0"
    for c in range(1, 8): V.cell(i, c).fill = KEY
    VAL[s] = i
W0 = R30 + 6
V.cell(W0, 1, "4. What must be true：FY2027 年底價 ＝ 現價 ×（1＋報酬）所需的倍數（FY2028 EPS；約 2027 年底的價格看下一年 EPS）").font = BOLD
hdr(V, W0 + 1, ["FY2027 年底價 ÷ 現價", "所需倍數（÷ FY2028 EPS，A）", "所需倍數（÷ FY2028 EPS，B）", "所需倍數（÷ FY2028 EPS，C）", "對照：十年中位數"])
vr = W0 + 2
for x in (1.0, 1.1, 1.21, 1.5, 2.0):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    for j, s in enumerate(SC):
        V.cell(vr, 2 + j, f"=px*A{vr}/PL!F{SUMR[(s, 'eps')]}").number_format = "0.0"
    V.cell(vr, 5, "=pe_mid").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數低於歷史中位數時，現價相對便宜；倍數是否合理，另看前瞻 P/E 歷史（Bloomberg）。本版未做先驗，無 Andy 合理 P/E 對照。")
WMT_NOTE = vr
# 十年 P/E：Goodinfo 年底收盤 ÷ EPS（2017–2025；2016 EPS 未找到）；stockanalysis 年度 P/E（FY2021–25）
PE_GI = [None, "=23.95/1.48", "=19.1/1.76", "=28.35/2.40", "=31.0/3.10", "=29.15/3.76", "=29.4/4.01", "=98.6/4.08", "=104/6.11", "=150.5/9.04"]
PE_SA = [None, None, None, None, None, 7.92, 7.50, 24.42, 17.22, 17.46]
T10 = vr + 2
V.cell(T10, 1, "5. 十年 P/E（2016–2025 年底；兩來源平均，分位排名用；無負值或 >200 需剔除；2016 兩來源皆未找到）").font = BOLD
hdr(V, T10 + 1, ["年底", "兩來源平均 P/E", "Goodinfo（年底收盤 ÷ EPS）", "stockanalysis（年度 P/E）"])
RB0 = T10 + 16
for i, y in enumerate(range(2016, 2026)):
    rr = T10 + 2 + i
    V.cell(rr, 1, str(y))
    has = []
    if PE_GI[i] is not None: V.cell(rr, 3, f"=B{RB0 + 5 + i}").number_format = "0.00"; has.append("C")
    if PE_SA[i] is not None: V.cell(rr, 4, f"=C{RB0 + 5 + i}").number_format = "0.00"; has.append("D")
    if has: V.cell(rr, 2, f"=AVERAGE(C{rr}:D{rr})").number_format = "0.00"
PE10 = f"$B${T10 + 2}:$B${T10 + 11}"
V.cell(T10 + 12, 1, "年數"); V[f"B{T10 + 12}"] = f"=COUNT({PE10})"
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "Goodinfo（年底收盤 ÷ EPS）", "vals": PE_GI, "tag": "Derived（資料商年底收盤與公司 EPS；未還原除息；2016 EPS 未找到）"},
    {"label": "stockanalysis（年度 P/E）", "vals": PE_SA, "tag": "Interested-party（資料商；只有 FY2021–25）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "#,##0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位，兩個來源分別取分位再平均。Goodinfo 有 2017–2025 共 9 年、stockanalysis 只有 FY2021–25；2017–2022 年緯創以筆電代工為主（約 7–16x），2023 年起 AI 伺服器重新定價（17–24x），規則值可能低估 AI 期的合理倍數。")
assert V.cell(RB0 + 5, 1).value == "2016" and V.cell(RB0 + 14, 1).value == "2025"
rule_rows = [rw for rw in range(RB0, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")

# 0. 目標價（Andy 2026-10-25 11:28 規格、11:53 修訂：DCF 與 Comparable 各半；列標籤依工作單逐字；B／C／D＝A 順利／B 平平／C 不如預期）
YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
V.cell(4, 1, "0. 目標價").font = BOLD
V.cell(5, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD
hdr(V, 6, ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"])
T = {}
def tset(r, key, label, fs, fmt, note=""):
    V.cell(r, 1, label); T[key] = r
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{r}"] = f; V[f"{c}{r}"].number_format = fmt
    if note: V[f"E{r}"] = note
V.cell(7, 1, "情境 EPS（FY2027–FY2032）"); V["E7"] = "PL 頁各情境（歸屬母公司 EPS）；FY2030 起為收斂法"; T["eps_h"] = 7
EPSR = {}
for i, y in enumerate(YRS6):
    tset(8 + i, f"eps{y}", f"  FY{y}", [f"=PL!{FC[y]}{SUMR[(s, 'eps')]}" for s in SC], "0.00"); EPSR[y] = 8 + i
V.cell(14, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）"); T["cf_h"] = 14
tset(15, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）")
for i, y in enumerate(YRS6):
    tset(16 + i, f"cf{y}", f"  FY{y}", [f"={c}{EPSR[y]}*(1-$B$15)" for c in "BCD"], "0.00")
def dcf_formula(c, r="tp_r", g="tp_g", reinv="$B$15"):
    """c＝情境欄；r、g 可替換（Sens 用）。基準日 2026 年底，t＝1 對 FY2027。"""
    cf = lambda y: f"Val!{c}{EPSR[y]}*(1-{reinv})"
    pv = "+".join(f"{cf(y)}/(1+{r})^{t}" for t, y in enumerate(YRS6, 1))
    return f"({pv}+{cf(2032)}*(1+{g})/({r}-{g})/(1+{r})^6)"
tset(22, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=SUMPRODUCT({c}16:{c}21,1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}21*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6" for c in "BCD"], "0.00",
     "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底")
tset(23, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}22*(1+tp_r)" for c in "BCD"], "0.00")
tset(24, "dcftp", "DCF 目標價（情境加權）", ["=tp_wA*B23+tp_wB*C23+tp_wC*D23", None, None], "0.00", "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁")
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = [("廣達（Quanta）", "2382.TW", 11.31, "2026-10-14"), ("鴻海（Hon Hai）", "2317.TW", 11.37, "2026-10-19"),
         ("英業達（Inventec）", "2356.TW", 18.00, "2026-08-14"), ("緯穎（Wiwynn；緯創子公司）", "6669.TW", 13.87, "2026-10-02"),
         ("美超微（Supermicro）", "SMCI", 7.06, "2026-10-23")]
for i, (nm, tk, pe, dt) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左）"); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
assert PEER_LAST == 32
tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", "伺服器 ODM 同業：廣達、鴻海、英業達、緯穎、美超微（券商常用比較組；不含 A 股）")
tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）", [f"=tp_wA*(2/12*B{EPSR[2027]}+10/12*B{EPSR[2028]})+tp_wB*(2/12*C{EPSR[2027]}+10/12*C{EPSR[2028]})+tp_wC*(2/12*D{EPSR[2027]}+10/12*D{EPSR[2028]})", None, None], "0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], "0.00", "同業中位數 × 前瞻 EPS")
V.cell(37, 1, "0.3 結論").font = BOLD
hdr(V, 38, ["項目", "數值", "", "", "說明"])
tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B24+(1-tp_w_dcf)*B35", None, None], "0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁（0.5）")
tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], "0.00")
tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], "0.00", "stockanalysis：17 位分析師平均 268.71（中位數 275、區間 165–400），2026-09-16 擷取")
tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-22 收盤）")
tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝09-29 價格 ÷ 前瞻 P/E（In 頁 cons_feps）")
tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
tset(46, "band", "價格帶（DCF 不如預期～順利）", ['=TEXT(D23,"0.0")&"～"&TEXT(B23,"0.0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=IF(px/{c}{EPSR[2028]}<MIN({PE10}),0,IF(px/{c}{EPSR[2028]}>MAX({PE10}),1,PERCENTRANK({PE10},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 5 節兩來源平均；2016 缺）的分位（低於十年最小值記 0%、高於最大值記 100%）；只作檢查，不進目標價")
V.cell(48, 1, "目標價方法")
V["B48"] = ("DCF：股東現金流＝歸屬母公司 EPS ×（1 − g／ROE），r＝10%、g＝4%、ROE 為近三年平均約 20%（再投資率約 20%），FY2027–32 明確期＋永續；"
            "組裝廠資本支出相對淨利不高，未改用淨利＋折舊－資本支出。FY2027–28 的 EPS 高於法人說法，DCF 偏高估。"
            "Comparable：廣達、鴻海、英業達、緯穎、美超微前瞻 P/E 中位數（不含 A 股）；英業達價格日期為 8 月、緯穎為子公司，同業倍數低（伺服器組裝的低利潤率）。"
            "情境對應：A＝順利（顆數基準）、B＝平平（顆數 ×0.8／0.85）、C＝不如預期（顆數 ×0.6）")
for r_ in (24, 35, 39):
    V[f"B{r_}"].fill = KEY
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；FY2026 恆等於錨點，份額與本體毛利率在每個案例內依錨點重算）：對 FY2027、FY2028 合併營收、合併毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("AI 部署 ×0.8", "FY2027 起 AI IT 資本支出與 GPU 顆數 ×0.8", dict(mf="0.8")), ("AI 部署 ×1.2", "×1.2", dict(mf="1.2")),
         ("本體份額流失", "相對值 2027 0.85、2028 0.75、2029 0.70", dict(rel="rel_{y}_lo")), ("本體份額上升", "相對值 1.15", dict(rel="rel_{y}_hi")),
         ("緯穎份額流失", "相對值 2027 0.85、2028 0.75、2029 0.70", dict(wrel="wrel_{y}_lo")), ("緯穎份額上升", "相對值 1.2", dict(wrel="wrel_{y}_hi")),
         ("客供 高", "客供比例 25%／40%／50%", dict(cs="cs_{y}_hi")), ("客供保留毛利 0.7", "客供部分只保留 70% 毛利", dict(keep="keep_lo")),
         ("每 GW 價值 低", "2028–29 每 GW ODM 價值 ×0.85", dict(vpg="vpg_{y}_lo")), ("每 GW 價值 高", "×1.2", dict(vpg="vpg_{y}_hi")),
         ("基板份額 低", "2027 0.8、2028 0.7、2029 0.65", dict(brel="brel_{y}_lo")), ("基板每顆金額 高", "2027 1.4、2028–29 1.5", dict(bv="bv_{y}_hi")),
         ("本體 AI 毛利率差距 −2pt", "基準 −1pt", dict(d="d_ai_lo")), ("緯穎營益率 5.5%", "基準 6.5%", dict(wopm="wopm_lo")), ("緯穎營益率 7.5%", "", dict(wopm="wopm_hi")),
         ("本體淨利係數 0.45", "基準 0.56", dict(snio="snio_lo")), ("本體淨利係數 0.65", "", dict(snio="snio_hi")),
         ("緯穎持股 35%", "基準 37.68%", dict(stake="stake_lo")), ("本體營業費用率 +0.2pt", "營業費用額外 +0.2% × 本體營收", dict(opxa="opx_adj_hi"))]
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0"), ("D", "gm", 2027, "0.00%"), ("E", "eps", 2027, "0.00"),
                           ("G", "rev", 2028, "#,##0"), ("H", "gm", 2028, "0.00%"), ("I", "eps", 2028, "0.00")):
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
put_case(nn, "情境 B", "CY2028 顆數 ×0.8、CY2029 ×0.85", PB["B"], "PL!")
put_case(nn + 1, "情境 C", "CY2028–29 顆數 ×0.6", PB["C"], "PL!")
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：FY2027 三情境相同；顆數情境 B、C 的差異從 FY2028 起。緯穎持股比例只影響歸屬母公司淨利（合併營收與毛利率不變）；客供案例同時改變營收與毛利率，EPS 只受保留毛利比例影響。")
SENS_NOTE = nn + 3
TPS0 = nn + 5
SE.cell(TPS0, 1, "目標價參數（只改 DCF 參數與權重；PL 收斂法的 fade_g 不隨 tp_g 變動）").font = BOLD
hdr(SE, TPS0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
def tp_case(i, name, setting, r="tp_r", g="tp_g", w="tp_w_dcf"):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    reinv = f"MIN({g}/tp_roe,0.8)"
    SE[f"C{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_formula(c, r, g, reinv)}" for s, c in zip(SC, "BCD")) + f")*(1+{r})"
    SE[f"D{i}"] = f"=Val!B{T['comptp']}"
    SE[f"E{i}"] = f"={w}*C{i}+(1-{w})*D{i}"
    SE[f"F{i}"] = f"=E{i}/E${TPS0 + 2}-1"
    for c in "CDE": SE[f"{c}{i}"].number_format = "0.00"
    SE[f"F{i}"].number_format = "0.0%"
TPCASES = [("基準", "r 10%、g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "r 11%", dict(r="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", dict(r="(tp_r-0.01)")),
           ("長期成長 +1pt", "g 5%", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", dict(g="(tp_g-0.01)")),
           ("DCF 權重 0.3", "", dict(w="0.3")), ("DCF 權重 0.7", "", dict(w="0.7"))]
for k, (nm, st, kw) in enumerate(TPCASES):
    tp_case(TPS0 + 2 + k, nm, st, **kw)
TPS_LAST = TPS0 + 1 + len(TPCASES)
assert TPS_LAST < 40, TPS_LAST
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 40
for c in "CDEFGHIJ": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 共識、券商估計與櫃數調查。只作事後對照，不為對上數字而調參數。券商多未具名或單一來源；共識 FY2027（付費）未找到"
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
RC.cell(rr_, 1, "A. 共識與券商估計（校準未使用）").font = BOLD; rr_ += 1
rec("FY2026 營收 vs 共識", "=cons_rev26", pl("rev", 2026), "stockanalysis（7 家，09-16）", "", fmt="#,##0.0")
rec("FY2026 EPS vs 共識", "=cons_eps26", pl("eps", 2026), "stockanalysis（7 家，09-16）", "3Q 營收創高後，共識可能上修")
rec("FY2026 EPS vs 凱基（05-08，2Q 前）", "=k_eps26", pl("eps", 2026), "凱基 2026-05-08", "")
rec("FY2027 EPS vs 法人下限說法（20 元以上）", "=b_eps27", pl("eps", 2027), "旺得富 2026-08-10（未具名）", "")
rec("FY2027 EPS vs 法人樂觀說法中點（23.3–26.4）", "=b_eps27h", pl("eps", 2027), "旺得富 2026-08-10（未具名）", "")
rec("FY2027 EPS vs 凱基（05-08）", "=k_eps27", pl("eps", 2027), "凱基 2026-05-08", "2Q 財報前，偏舊")
rec("FY2027 營收年增（模型）", None, pl("g", 2027), "—", "", fmt="0.0%")
RC.cell(rr_, 1, "B. 緯穎（子公司）vs 凱基緯穎報告").font = BOLD; rr_ += 1
rec("FY2026 緯穎營收 vs 凱基（07-28）", "=kw_rev26", pl("wrev", 2026), "凱基 2026-07-28", "3Q26 緯穎月營收大增，高於凱基", fmt="#,##0.0")
rec("FY2027 緯穎營收 vs 凱基", "=kw_rev27", pl("wrev", 2027), "凱基 2026-07-28", "", fmt="#,##0.0")
rec("FY2027 緯穎淨利 vs 凱基", "=kw_ni27", pl("wni", 2027), "凱基 2026-07-28", "", fmt="#,##0.0")
rec("FY2027 緯穎對歸屬母公司淨利的貢獻比重（模型）", None, pl("wcon", 2027), "—", "持股 37.68% × 緯穎淨利 ÷ 歸屬母公司淨利", fmt="0.0%")
RC.cell(rr_, 1, "C. 櫃數與份額（機制檢查）").font = BOLD; rr_ += 1
rec("FY2026 隱含本體出貨櫃數 vs 凱基 GB＋VR 約 1.5 萬櫃（千櫃）", "=k_racks", pl("rk", 2026), "凱基 2026-05-08", "本體 AI 系統營收為殘差；凱基口徑（L10 是否含 GPU）不明", fmt="#,##0.0")
rec("FY2026 本體有效份額 vs 三家 ODM 中緯創的 NVL72 櫃數份額（百分點）", "=rk_sh", lambda s: f"=PL!D{PB[s]['sh0']}", "經濟日報 2026-09-10（美系券商）", "口徑不同：本模型分母為全部 AI 組裝市場", fmt="0.0%", pt=True)
rec("8 月調查年化櫃數（2,300–2,400 × 12；千櫃）vs 隱含本體櫃數", "=2.35*12", pl("rk", 2026), "經濟日報 2026-09-10", "調查可能含緯穎或運算托盤口徑", fmt="#,##0.0")
rec("FY2026 GPU 基板營收 vs 凱基「超過千億」（NT$B）", 100, pl("brd", 2026), "凱基 2026-05-08", "校準值（brd_26 取 110）", fmt="#,##0.0")
rec("FY2026 本體一般伺服器與交換器 vs 凱基交換器 160–180 億（NT$B）", 17, pl("gen", 2026), "凱基 2026-05-08", "本模型含一般伺服器（Assumed）", fmt="#,##0.0")
RC.cell(rr_, 1, "D. 合併與歸母").font = BOLD; rr_ += 1
rec("FY2026 非控制權益 ÷ 合併淨利", None, lambda s: f"=PL!D{PB[s]['nci']}/PL!D{PB[s]['cni']}", "—", "合併淨利約四成歸緯穎的少數股東", fmt="0.0%")
rec("FY2027 緯穎 ÷ 合併營收", None, pl("wshr", 2027), "—", "", fmt="0.0%")
rec("4Q26 EPS（模型）", None, lambda s: "=Q!E24", "—", "", fmt="0.00")
RC.cell(rr_, 1, "E. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以部署路徑取代「AI IT 資本支出 ÷ 每 GW IT 資本支出」"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60
RECON_LAST = rr_ - 1

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文只經法說摘要與新聞轉述，原始簡報待人工複核；券商與資料商皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("公司（摘要）：鉅亨 2Q26 法說（2Q26 營收、毛利率、營益率、淨利、EPS；AI 伺服器展望、L10／L11、交換器、GDR）", CNY2Q),
        ("公司（摘要）：優分析 2Q26 法說（1H26 營收、毛利率、營益率、淨利 244.58 億、EPS 7.78；資本支出）", UAN2Q),
        ("公司（摘要）：優分析 2025 財報（FY2025 營收 2.19 兆、毛利率 6.13%、營益率 3.59%、淨利 274.08 億、EPS 9.04）", UAN25),
        ("公司：月營收公告（TechNews 7 月）", TN07), ("公司：月營收公告（TechNews 8 月）", TN08), ("公司：月營收公告（TechNews 9 月）", TN09),
        ("公司（摘要）：優分析 2026-10-08（3Q26 約 1.15 兆、4Q 展望）", UAN09),
        ("緯穎（摘要）：富果 4Q25 法說（FY2025 營收、毛利、營業利益、淨利、AI 占比、2026 展望）", FUGW),
        ("緯穎（摘要）：優分析 2Q26 財報（1H26、2Q26 營收、毛利率、營益率、淨利）", UANW2Q),
        ("緯穎：月營收公告（TechNews 9 月；前 9 月 9,892.78 億）", TNW09),
        ("券商：凱基緯創報告 2026-05-08（2026–27 EPS、伺服器占比、零組件、櫃數、交換器）", KGI05),
        ("券商：凱基緯穎報告 2026-07-28（緯穎 2025A–2027F、緯創持股 37.68%）", KGIW),
        ("資料商：CMoney 2024-08-01（緯創持有緯穎 42.66%）", CM2024),
        ("媒體：財報狗 GDR 新聞（2.5 億股、NT$186.24、稀釋 7.29%）", SD_GDR),
        ("券商：經濟日報 2026-09-10（美系券商 NVL72 櫃數調查）", UDN0910),
        ("券商：旺得富 2026-08-10（法人 2026–27 EPS、目標價）", WR0810), ("券商：工商時報 2026-08-05（本國投顧目標價 300）", CTEE0805),
        ("資料商：財報狗（10-22 收盤 190、近 12 月 P/E）", SD), ("資料商：stockanalysis 報價頁（前瞻 P/E、股數）", SA_Q),
        ("資料商：stockanalysis 預估頁（共識營收、EPS、目標價）", SA_F), ("資料商：stockanalysis 財報與比率頁（FY2021–25 損益、ROE、P/E）", SA_FIN),
        ("資料商：stockanalysis 比率頁", SA_R), ("資料商：Goodinfo 歷年經營績效（年底收盤、EPS、ROE）", GI),
        ("同業前瞻 P/E：stockanalysis（廣達、鴻海、英業達、緯穎、美超微各報價頁）", "https://stockanalysis.com/quote/tpe/2382/；https://stockanalysis.com/quote/tpe/2317/；https://stockanalysis.com/quote/tpe/2356/；https://stockanalysis.com/quote/tpe/6669/；https://stockanalysis.com/stocks/smci/"),
        ("市場層：markets/network、markets/chip_bridge、tsmc、quanta、foxconn 成品（本 repo，唯讀）", f"{rel(NET_F)}；{rel(BR_F)}；{rel(TSMC_F)}；{rel(QU_F)}；{rel(FX_F)}"),
        ("待補（人工）：3Q26 財報與法說（約 11-10）；緯創本體 AI 系統、GPU 基板、一般伺服器的金額（公司未揭露）；非控制權益金額；業外與稅前利益；客供比例的量化；2016 EPS；共識 FY2027（付費）；前瞻 P/E 十年歷史（Bloomberg）", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 100; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("緯創資通（3231）收支與估值模型 v0.1.1（AI 半導體 Project：伺服器組裝次產業，樣板＝quanta/）　2026-10-25", True),
 ("版本：v0.1.1（2026-10-25：補上檔空間級距、投資建議、重點觀察（數值不變））。v0.1（初版）。伺服器組裝次產業第三家（廣達、鴻海之後）；AI 伺服器組裝市場、ODM 每櫃價格、客供比例讀廣達模型（規則第 13 條）。另加 GPU 基板驅動（讀 tsmc 成品顆數），並把子公司緯穎拆開以區分合併與歸母口徑。", False),
 ("模型命題：緯創的合併營收＝本體 AI 系統（部署 GW × 每 GW 架數 × 每櫃價格 × 份額 ×（1 − 客供））＋GPU 基板與運算板（GPU 顆數 × 每顆基板金額 × 份額）＋一般伺服器與非伺服器＋子公司緯穎（ASIC 機架為主）。緯穎占合併營收約三分之一、營益率約為本體的三倍，但緯創只持有 37.68%，合併淨利約四成歸少數股東；EPS 看歸屬母公司淨利＝本體淨利＋37.68% × 緯穎淨利。", False),
 ("核心問題：本體約 2% 的營益率能不能隨 AI 系統與基板放量撐住，以及緯穎的成長有多少真正流到緯創股東；約 9–10 倍前瞻 P/E 隱含什麼路徑。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣；整櫃價格與市場以美元計、以匯率換算。口徑：營收、毛利、營業利益為合併（含緯穎）；淨利與 EPS 為歸屬母公司。3Q26 財報與法說約 2026-11-10，本版只納入 3Q26 月營收。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("本體 AI 系統：markets/network 的 AI IT 資本支出 × 情境顆數係數 ÷ 每 GW IT 資本支出（Tokenomics，經橋接表）＝部署 GW；× 每 GW 架數 × ODM 每櫃價格（廣達模型）＝組裝市場；× 本體有效份額（FY2026 以本體營收殘差校準）× 份額相對值（Assumed）×（1 − 客供比例，廣達模型）。", False),
 ("GPU 基板與運算板：tsmc 成品的 NVIDIA＋AMD GPU 顆數 × 緯創基板份額（Assumed）× 每顆基板金額（FY2026 以券商「AI 伺服器零組件超過千億」校準，之後 × 金額指數，Assumed）。緯穎：AI＝組裝市場 × 緯穎 AI 份額（FY2026 以緯穎 AI 占比略高於五成校準）；非 AI 以成長率推。", False),
 ("損益：本體毛利＝各部門營收 × 毛利率（非伺服器毛利率 FY2026 校準，AI 系統、基板、一般伺服器加減差距，Assumed），營業費用彈性 phi，本體淨利＝本體營業利益 × 淨利係數；緯穎營業利益＝營收 × 營益率、淨利＝營業利益 × 係數。合併淨利＝本體＋緯穎；非控制權益＝（1 − 持股）× 緯穎淨利；歸屬母公司淨利＝合併淨利 − 非控制權益。", False),
 ("校準：FY2026 ＝ 1H26 實績（公司與緯穎）＋3Q26 營收實績（兩家月營收）＋假設（4Q 季增、利潤率）。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，三情境權重 0.3／0.5／0.2）＋0.5 × Comparable 目標價（廣達、鴻海、英業達、緯穎、美超微前瞻 P/E 中位數 × 12 個月後前瞻 EPS；不含 A 股）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解。十年 P/E 規則值 × FY2027 EPS 保留為參考（不是目標價）。情境對應：A＝順利、B＝平平、C＝不如預期（即顆數情境）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m，FY2032 達到；PL 頁灰底）。十年平均淨利率未找到，C 用 FY2029 × 0.8、B 用 FY2029 × 0.9（Assumed）。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司原文只經法說摘要與新聞；券商多未具名或單一來源。", False),
 ("", False),
 ("已知限制：(1) 本體 AI 系統、GPU 基板、一般伺服器的金額公司都沒有揭露，AI 系統是殘差，基板取單一券商說法；(2) 緯穎持股只有單一近期來源（凱基 37.68%），FY2025 用 40%（Assumed）拆分，非控制權益金額未找到；(3) 集團內交易（緯創與緯穎互供）假設為 0；(4) 客供比例、份額相對值、基板份額與每顆金額是 Assumed；(5) 業外與稅前利益未找到，以淨利 ÷ 營業利益係數涵蓋；(6) 十年 P/E 只有 2017–2025（Goodinfo）與 2021–25（stockanalysis），2016 缺；(7) 3Q26 財報（約 11-10）後更新錨點；R4 產業部署路徑未建。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ================= v0.1.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（工作單；Andy 2026-10-25 14:00、14:17） =================
# 全部放在 Val、README 現有最後一列之後，不位移既有儲存格
OBS = [
       "目標價 NT$355.1（DCF 397.0、Comparable 313.3 各半），比現價 190（2026-10-22）高 86.9%，級距大於 50%；Comparable 單獨為 64.9%。主因是 EPS 不是倍數：本模型 12 個月後前瞻 EPS 27.56 比共識前瞻 EPS 18.68 高 47.5%（兩者期間差約一年；FY2027 EPS 27.58 也比法人樂觀說法中點 24.85 高 11.0%），同業（廣達、鴻海、英業達、緯穎、美超微）前瞻 P/E 中位數 11.4 倍與現價隱含 10.2 倍相近；DCF（ROE 20.2%）比 Comparable 高 26.7%，再把目標價往上拉。",
       "獲利的兩個槓桿是本體淨利係數與緯穎（持股 37.68%）：本體淨利係數 0.45（基準 0.56）時 FY2027 EPS −11.1%，緯穎營益率 5.5%（基準 6.5%）時 −6.7%，僅次於 AI 部署 ±17.2%。3Q26 財報與法說（約 2026-11-10，緯穎約同期）可檢驗本體獲利率。",
       "情境分歧從 FY2028 起：FY2028 EPS A／B／C＝33.13／26.96／20.64（B、C 比 A 低 18.6%、37.7%）；現價只有 FY2027 EPS 的 6.9 倍，情境 C 的 FY2028 也只有 9.2 倍，低於十年 P/E 中位數 14.5 倍。"
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

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "wistron")
wb.save(sys.argv[1])
if os.environ.get("WISTRON_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE, "UN_NOTE": UN_NOTE,
               "T": T, "T10": T10, "TPS0": TPS0, "TPS_LAST": TPS_LAST, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0, "QCHK_LAST": QCHK_LAST},
              open(os.environ["WISTRON_ROWMAP"], "w"), ensure_ascii=False, indent=1)

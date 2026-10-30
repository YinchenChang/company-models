# 光寶科（2301.TW）個股模型 v0.2（AI 半導體 Project：電源與散熱次產業；AI 伺服器電源 PSU、電源櫃、BBU）
# 用法：python3 liteon/build_liteon.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；AI 電源市場以美元計、以匯率 fx 換算；EPS：新台幣。
# 口徑：歸屬母公司淨利（IFRS）。樣板：delta/（台達電：部署 GW × 每 kW 電源內容 × 份額）。
# 每 kW 電源內容口徑與台達電一致：唯讀 delta/dist 的 In 頁 pkw（本專案 AI 電源內容權威在台達電模型）；Units、Recon 頁核對光寶＋台達電份額合計。
# 情境與行業一致（Andy 2026-10-25 13:13）：A／B／C 讀 tsmc 同字母顆數係數（＝markets/memory SCEN_UNITS）；公司特有變數（份額相對值、AI 毛利率額外改善）A 取高值、B 基準、C 低值。
# FY2030–FY2032 為收斂法（上游只到 2029）；目標價＝DCF 與 Comparable 加權（Val 第 0 節）。本版未做先驗。
import sys, os, glob
import openpyxl
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
import cycle_position as CP
import valuation_rule as VR
import market as MEM

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
DLT_F = only("delta/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
ITC = None
for r in openpyxl.load_workbook(NET_F, data_only=True)["Mkt"].iter_rows(min_row=4):
    if r[0].value == "AI IT 資本支出（加速器 × 每顆 IT capex）":
        ITC = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
assert ITC and all(v is not None for v in ITC[0]), ITC
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and str(k).startswith("um"):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
assert all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()
for s_ in SC:   # 情境與行業一致：tsmc 顆數係數必須等於 markets/memory 的 SCEN_UNITS（同一字母）
    assert abs(TS[f"um28_{s_}"][0] - MEM.SCEN_UNITS[s_][0]) < 1e-9 and abs(TS[f"um29_{s_}"][0] - MEM.SCEN_UNITS[s_][1]) < 1e-9, s_
BR = {}
for r in openpyxl.load_workbook(BR_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k in ("racks", "capit"):
        BR[k] = (r[4].value, r[7].value, f"{rel(BR_F)}｜In!E{r[0].row}（GB300）、H{r[0].row}（VR200）")
assert len(BR) == 2, BR.keys()
DWB = openpyxl.load_workbook(DLT_F, data_only=True)
DPK = {}
for r in DWB["In"].iter_rows(min_row=4):
    k = r[0].value
    if k in ("fx",):
        DPK[k] = ((r[3].value, r[4].value, r[5].value), r[6].value, f"{rel(DLT_F)}｜In!D{r[0].row}:F{r[0].row}")
assert len(DPK) == 1 and all(v is not None for k in DPK for v in DPK[k][0]), DPK
DU = {}
for r in DWB["Units"].iter_rows(min_row=4):
    lab = r[0].value
    for key, want in (("sh", "台達電有效份額"), ("pmk", "AI 電源市場"), ("aip", "台達電 AI 電源營收"), ("gw", "部署 GW（情境 A）")):
        if lab == want:
            DU[key] = ([r[2 + j].value for j in range(4)], f"{rel(DLT_F)}｜Units!C{r[0].row}:F{r[0].row}")
assert len(DU) == 4 and all(v is not None for k in DU for v in DU[k][0]), DU
# Siliconomics M17 的台達電 FY2026 有效份額是台達電模型校準值的快照：兩者必須相同（台達電錨點更新時，先更新本庫）
assert abs(DU["sh"][0][0] - SIL.base("pwr_dsh_26")) < 1e-9, ("Siliconomics pwr_dsh_26 與 delta/dist 不一致", DU["sh"][0][0], SIL.base("pwr_dsh_26"))

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B），AI 電源市場與每 kW 內容為美元；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文＝光寶科財報與法說會（經媒體、資料商轉述）、月營收公告；公司的 AI 占比、產品時程說法為 Interested-party；券商與媒體估計皆為 Interested-party"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
UA2Q = "https://uanalyze.com.tw/articles/3402052563"
FUG2Q = "https://blog.fugle.tw/post/earnings-call-2301-2026-07-31"
UAFY = "https://uanalyze.com.tw/articles/1580848095"
FUGFY = "https://blog.fugle.tw/post/earnings-call-2301-2026-02-25"
TN9 = "https://finance.technews.tw/2026/10/08/ltc-2301-202609-financial-report/"
NST = "https://www.nstock.tw/2301?status=2"
SA = "https://stockanalysis.com/quote/tpe/2301/"
SAQ = "https://stockanalysis.com/quote/tpe/2301/financials/?p=quarterly"
SAY = "https://stockanalysis.com/quote/tpe/2301/financials/"
SAF = "https://stockanalysis.com/quote/tpe/2301/forecast/"
MSV = "https://www.marketscreener.com/quote/stock/LITE-ON-TECHNOLOGY-CORPOR-6492334/valuation/"
WEARN = "https://stock.wearn.com/financial.asp?kind=2301"
DIGRIN = "https://www.digrin.com/stocks/detail/2301.TW/price"
HIS = "https://histock.tw/stock/2301/%E6%9C%AC%E7%9B%8A%E6%AF%94"
KGI = "https://www.kgi.com.hk/zh-hk/-/media/files/kgishk/research-reports/tw-reports/power-supply-sector_20250926e.pdf"

sec("A. AI 部署規模（讀 markets/network、markets/chip_bridge、tsmc 成品；情境 A 基準，B／C 依同字母顆數係數縮放）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}", "US$B", None, ITC[0][j], None, "Derived", ITC[1],
        "加速器顆數 × 每顆 IT capex（含 ASIC）；本模型以此除以每 GW IT 資本支出得到部署 GW（與台達電模型相同）" if y == 26 else "同上")
inp("itpg_gb", "每 GW IT 資本支出（GB300）", "US$B", None, BR["capit"][0], None, "Derived", BR["capit"][2], "Tokenomics IF_CapexIT（經橋接表；唯讀）")
inp("itpg_vr", "每 GW IT 資本支出（VR200；2027 起）", "US$B", None, BR["capit"][1], None, "Derived", BR["capit"][2], "同上；Rubin Ultra 與 ASIC 機架未另訂（Assumed：同 VR200）")
inp("vr26", "2026 年部署中 Rubin 世代占比（混合每 GW IT 資本支出用）", "%", *SIL.vals("vr26"), "Interested-party", "TrendForce 2026 組合 Rubin 22%（經 markets/memory；與台達電模型相同）", "—")
inp("racks_vr", "每 GW 機架數（VR200）", "架", None, BR["racks"][1], None, "Derived", BR["racks"][2], "Units 頁每架內容換算用")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝順利（基準）；B＝平平；C＝不如預期；與台積電、台達電模型相同（＝markets/memory SCEN_UNITS，建置時核對）")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")
inp("fx", "匯率（新台幣 ÷ 美元）", "NT$/US$", *DPK["fx"][0], "Derived", DPK["fx"][2], "與台達電模型相同（唯讀）；10-23 收盤：未找到")

sec("B. AI 電源內容（口徑與台達電一致：唯讀 delta/dist In 頁）與光寶份額")
PKTAG = "Siliconomics 快照 In 頁 M17（本專案 AI 電源每 kW 內容唯一值；含機架電源櫃、PSU、BBU、HVDC 電源櫃；口徑與台達電模型相同）"
for y, lab in ((26, "2026"), (27, "2027（VR200）"), (28, "2028（800V HVDC 放量）"), (29, "2029")):
    inp(f"pkw_{y}", f"AI 電源內容 {lab}（台達電模型）", "US$/kW", *SIL.vals(f"pwr_pkw_{y}"), SIL.V[f"pwr_pkw_{y}"]["tag"], f"siliconomics｜SIF_pwr_pkw_{y}",
        PKTAG if y == 26 else "同上；低／高亦取自本庫（Sens 用）")
inp("x_np", "AI 相關營收中非電源部分（液冷 CDU、Sidecar、AI 光電等）占比", "%", 0.0, 0.10, 0.25, "Assumed",
    "公司：液冷 120kW CDU、Sidecar 2Q26 起出貨（富果 2026-07-31 法說摘要）；凱基 2025-09：光寶 2026 年液冷營收目標約 5%（Interested-party）", "只影響份額（分子）與光寶＋台達電份額合計，不影響營收（非電源部分隨 AI 電源同比例成長，Assumed）")
for y in (27, 28, 29):
    inp(f"rel_{y}", *SIL.row9(f"pwr_lrel_{y}")[1:])   # Siliconomics M17

sec("C. 營收結構錨點（公司揭露比例 × 營收）")
inp("ai_25", "2025 年 AI 相關營收 ÷ 營收", "%", 0.20, 0.21, 0.23, "Interested-party", UAFY + "：AI 相關營收超過 20%（公司，達標）", FUGFY + "：2025 年 AI 占比超過 20%；雲端與物聯網 45%、光電 17%、資訊與消費電子 38%")
inp("ai_26", "2026 年 AI 相關營收 ÷ 營收", "%", 0.30, 0.31, 0.35, "Interested-party", FUG2Q + "：AI 相關產品 2026 年超過 30%（公司）；BBU 約占雲端營收 20–25%", FUGFY + "：2026 年目標 30% 以上、長期 50% 以上")
sec("C2. 營收分部（官方三大事業群：雲端及物聯網、光電、資訊及消費性電子；v0.2 依個股模型原則重建）")
SL = "Investing.com 2026-07-31（Lite-On Q2 2026 slides 摘要）：2Q26 Cloud and AIoT 289 億（55%，季增 26%）、Opto-electronics 80 億（季增 12%）、IT and Consumer Electronics 158 億（季增 18%）"
inp("s25_cl", "2025 年雲端及物聯網 ÷ 營收", "%", None, 0.45, None, "Interested-party", FUGFY + "：2025 年雲端與物聯網 45%、光電 17%、資訊與消費電子 38%（公司法說）", "月營收新聞稿 2025-12：雲端及物聯網 46%、光電 14%、資訊及消費性電子 40%（單月，旺得富 2026-01-08）")
inp("s25_op", "2025 年光電 ÷ 營收", "%", None, 0.17, None, "Interested-party", "同上", "—")
inp("s25_ic", "2025 年資訊及消費性電子 ÷ 營收", "%", None, "=1-s25_cl-s25_op", None, "Derived", "1 − 雲端及物聯網 − 光電（公司 38%）", "公司比例為四捨五入值，合計 100%")
inp("q2_cl", "2Q26 雲端及物聯網營收", "NT$B", None, 28.9, None, "Verified", SL, "公司：2Q26 營收 527 億，雲端及物聯網占 55%（富果 2026-07-31）")
inp("q2_op", "2Q26 光電營收", "NT$B", None, 8.0, None, "Verified", SL, "—")
inp("q2_ic", "2Q26 資訊及消費性電子營收", "NT$B", None, 15.8, None, "Verified", SL, "三事業群合計 527 億＝合併營收 52.703（差 0）")
inp("g1_cl", "2Q26 雲端及物聯網季增", "%", None, 0.26, None, "Verified", SL, "1Q26＝2Q26 ÷（1＋季增）（Derived）")
inp("g1_op", "2Q26 光電季增", "%", None, 0.12, None, "Verified", SL, "同上")
inp("g1_ic", "2Q26 資訊及消費性電子季增", "%", None, 0.18, None, "Verified", SL, "同上；1Q26 三者合計 434.7 億 vs 公司 434.07 億（季增四捨五入，差額併入資訊及消費性電子）")
inp("s2h_cl", "3Q、4Q26 雲端及物聯網 ÷ 營收", "%", 0.53, 0.55, 0.58, "Assumed", "沿用 2Q26 55%；3Q26 各事業群營收：未找到（月營收新聞稿未取得）", "公司：雲端運算全年年增超過 70%（富果 2026-07-31）")
inp("s2h_op", "3Q、4Q26 光電 ÷ 營收", "%", 0.14, 0.152, 0.16, "Assumed", "沿用 2Q26 15.2%（80 ÷ 527）", "—")
for y, v in ((27, (0.0, 0.05, 0.10)), (28, (0.0, 0.05, 0.10)), (29, (0.0, 0.05, 0.10))):
    inp(f"g_cn{y}", f"雲端及物聯網：非 AI 部分年增 20{y}（一般伺服器與網通電源、車用、工控與 IoT）", "%", *v, "Assumed",
        "公司 2026：工控、標準電源、邊緣運算擴充；車用電源（富果 2026-02-25）" if y == 27 else "同上", "FY2026 由錨點反推（雲端及物聯網 − AI 相關）" if y == 27 else "—")
for y, v in ((27, (0.0, 0.05, 0.10)), (28, (0.0, 0.05, 0.10)), (29, (0.0, 0.05, 0.10))):
    inp(f"g_op{y}", f"光電年增 20{y}（光耦合器、感測、AI 伺服器與交換器指示燈）", "%", *v, "Assumed",
        "公司：高階光耦合器用於工控、新能源與 AI 電能（technews 2026-03-10）；2Q26 季增 12%" if y == 27 else "同上",
        "Denselight（新加坡 InP，持股 21%，權益法）不併入營收" if y == 27 else "—")
for y, v in ((27, (-0.05, 0.0, 0.05)), (28, (-0.05, 0.0, 0.05)), (29, (-0.05, 0.0, 0.05))):
    inp(f"g_ic{y}", f"資訊及消費性電子年增 20{y}（PC 與遊戲機電源、低軌衛星與高階電源、影像）", "%", *v, "Assumed",
        "公司 2026 指引：資訊產品 −6～−8%、手機 −4～−5%（記憶體短缺）；下半年持平（富果 2026-07-31）" if y == 27 else "同上",
        "低軌衛星高階電源 2026 年增一倍以上（Investing.com 2026-07-31，Interested-party）" if y == 27 else "—")

sec("D. 毛利率")
inp("d_ai", "AI 相關毛利率高於其他營收的幅度", "pt", 0.08, 0.12, 0.16, "Assumed",
    UA2Q + "：2Q26 雲端與物聯網營收 288.9 億、營業利益 69 億（營益率約 24%）vs 公司 15.6%；其他部門合計約 5–6%（Derived）", "部門別毛利率未揭露：未找到；毛利率 FY2025 22.9% → 2Q26 27.2%（AI 占比上升）")
for y, v in ((27, (-0.01, 0.0, 0.01)), (28, (-0.015, 0.0, 0.015)), (29, (-0.02, 0.0, 0.02))):
    inp(f"gmu_{y}", f"AI 相關毛利率額外改善 20{y}（800V 電源櫃、組合；情境 A 取高、C 取低）", "pt", *v, "Assumed",
        "公司：2026 年毛利率目標約 25%（富果 2026-07-31）；800V HVDC 電源櫃與 110kW Power Shelf 2027 放量（Interested-party）" if y == 27 else "同上",
        "下檔：零組件漲價、新廠（高雄二期、越南、德州）爬坡" if y == 27 else "—")

sec("E. FY2026 錨點（1H26 財報＋3Q26 月營收實績＋假設；3Q26 財報與法說預計 2026-10-30 前後，截至 2026-10-25 尚未公布）")
inp("ytd9", "2026 年 1–9 月營收（月營收累計，實績）", "NT$B", None, 153.559, None, "Verified", TN9 + "：9 月 188.48 億元、前 9 月 1,535.59 億元（年增 26.1%）", NST + "：9 月 188.48 億、前 9 月 1,535.59 億（同值）")
inp("q3_rev", "3Q26 營收（＝1–9 月累計 − 1H26）", "NT$B", None, "=ytd9-(43.407+52.703)", None, "Derived", "1H26 營收 961.1 億（優分析 2026-07-31）", "7、8 月個別營收：未找到")
inp("q4g", "4Q26 營收季增", "%", -0.03, 0.03, 0.08, "Assumed", "公司：下半年「非常樂觀」；800V HVDC 電源櫃 11 月小量（富果 2026-07-31）", "9 月月減 3.8%（technews）；資訊與消費電子下半年與上半年持平；共識 FY2026 營收 2,198 億（MarketScreener，只作 Recon）")
inp("q_gm", "3Q、4Q26 毛利率", "%", 0.25, 0.26, 0.275, "Assumed", "1Q26 21.7%、2Q26 27.2%、1H26 24.7%（stockanalysis、優分析）", "公司 2026 年毛利率目標約 25%（富果 2026-07-31，Interested-party）")
inp("q_opx", "3Q、4Q26 營業費用 ÷ 營收", "%", 0.112, 0.118, 0.125, "Assumed", "1Q26 12.2%、2Q26 11.7%（毛利 − 營業利益；Q 頁）", "FY2025 12.8%；研發約占營收 5–5.5%（富果 2026-07-31）")
inp("q_nop", "3Q、4Q26 營業外淨額（每季）", "NT$B", 0.5, 0.9, 1.2, "Assumed", "1H26 稅前 − 營業利益＝約 20.4 億（stockanalysis 稅前淨利率推算）", "FY2025 約 29.7 億（同上）")
inp("q_nir", "淨利 ÷ 稅前（1H26 實績；3Q、4Q26 沿用）", "%", 0.73, "=10.891/14.341", 0.78, "Derived", "1H26：淨利 108.91 ÷ 稅前約 143.41 億（stockanalysis 季度淨利、稅前淨利率推算）", "FY2025：151.16 ÷ 197.1＝0.767（同上）")
inp("q_sh", "股數（十億股；EPS 換算）", "十億股", None, "=10.891/4.80", None, "Derived", "1H26 淨利 108.91 億 ÷ EPS 4.80（優分析 2026-07-31）", "stockanalysis 流通股數 22.7 億股（2026-10-22）")

sec("F. 營業費用、業外、稅與股數（FY2027 起）")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.5, 0.7, 1.0, "Assumed", "FY2026 錨點實測約 0.65（Q 頁 B26）", "FY2025：營業費用 +17% ÷ 營收 +21%（stockanalysis 推算）")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.01, 0.0, 0.01, "Assumed", "無", "Sens 用 ±1 個百分點")
inp("nop_f", "營業外淨額（每年）", "NT$B", 2.0, 3.5, 5.0, "Assumed", "FY2025 約 29.7 億、1H26 約 20.4 億（stockanalysis 推算）", "含利息收入、投資收益；Denselight 權益法損益未另計")
inp("nir_f", "淨利 ÷ 稅前（FY2027 起）", "%", 0.73, 0.76, 0.78, "Derived", "同 q_nir（1H26 0.759；FY2025 0.767）", "—")
inp("sh_g", "股數年變動", "%", 0.0, 0.0, 0.005, "Assumed", "FY2025 推導 22.8 億股 → 1H26 22.7 億股（淨利 ÷ EPS）", "—")
inp("mf", "AI 部署係數（Sens 用；FY2027 起）", "x", 0.8, 1.0, 1.2, "Assumed", "網通模型 2027 年 AI IT 資本支出約 2 倍", "±20%")

sec("G. 共識、券商與台達電模型（只作 Recon 對帳，不作校準）")
for y, rv, ep in ((26, 219.778, 10.81), (27, 292.383, 14.99), (28, 366.110, 20.74)):
    inp(f"cons_rev{y}", f"共識 FY20{y} 營收（MarketScreener）", "NT$B", None, rv, None, "Interested-party", MSV + "（參考價 255.50，2026-10-22）", "stockanalysis FY2026 2,193 億（區間 2,101–2,297）" if y == 26 else "—")
    inp(f"cons_eps{y}", f"共識 FY20{y} EPS（MarketScreener）", "NT$", None, ep, None, "Interested-party", MSV, "stockanalysis FY2026 10.81（11 家，區間 10.06–11.79）" if y == 26 else "FY2027 以後第二來源：未找到")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"d_sh{y}", f"台達電模型：台達電 AI 電源有效份額 20{y}", "%", None, DU["sh"][0][j], None, "Derived", DU["sh"][1], "唯讀台達電成品（三情境相同：份額相對值不隨情境變動）" if y == 26 else "同上")
    inp(f"d_pmk{y}", f"台達電模型：AI 電源市場 20{y}（情境 A）", "US$B", None, DU["pmk"][0][j], None, "Derived", DU["pmk"][1], "唯讀台達電成品（Recon：與本模型情境 A 應相同）" if y == 26 else "同上")
    inp(f"d_aip{y}", f"台達電模型：台達電 AI 電源營收 20{y}（情境 A）", "US$B", None, DU["aip"][0][j], None, "Derived", DU["aip"][1], "唯讀台達電成品" if y == 26 else "同上")
inp("kgi_dsh", *SIL.row9("pwr_kgi_0926")[1:])   # Siliconomics M17（CF11 待 Andy）

sec("H. 價格與估值")
inp("px", "股價（2026-10-22 收盤，臺灣證交所）", "NT$", None, 255.5, None, "Interested-party", SA + "：10-22 255.50（−3.04%）、前收 263.50", "nstock（網址見 Sources 頁）：10-22 255.5（成交 7,003 張；HiStock 同值）；10-23 收盤：未找到")
inp("pe_ttm", "近 12 月 P/E（stockanalysis，10-22）", "x", None, 30.09, None, "Interested-party", SA, "近 12 月 EPS 8.49")
inp("fpe", "前瞻 P/E（stockanalysis，10-22）", "x", None, 19.70, None, "Interested-party", SA, "口徑：頁面未明示（約未來 12 月）")
inp("cons_feps", "共識前瞻 12 月 EPS（現價 ÷ 前瞻 P/E）", "NT$", None, "=px/fpe", None, "Derived", "stockanalysis 前瞻 P/E 19.70（2026-10-22）", "共識前瞻 EPS 的直接數字：未找到")
inp("cons_tp", "市場共識目標價（賣方平均）", "NT$", None, 314.25, None, "Interested-party", SA + "：12 位分析師（S&P Global）平均 314.25（頁面時間 2026-10-22）",
    SAF + "：2026-09-23 頁面平均 302.67、中位數 292.50、區間 176–400")
inp("nim_avg", "十年平均淨利率（2016–2025）", "%", None, "=AVERAGE(0.041,0.0123,0.0385,0.0529,0.0626,0.0845,0.0818,0.0984,0.0872,0.091)", None, "Interested-party",
    WEARN + "：稅後淨利率 2016–2025：4.10%、1.23%、3.85%、5.29%、6.26%、8.45%、8.18%、9.84%、8.72%、9.10%（資料商）", "單一來源；2017 年 1.23% 含一次性損失（商譽減損）")

sec("J. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_r", "折現率／要求報酬", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%，各市場同一值）", "DCF 折現；情境價值推到 12 個月後")
inp("tp_g", "長期成長率", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "DCF 永續成長與 FY2032 收斂目標（情境 B）")
inp("tp_roe", "ROE（近三年平均，FY2023–25）", "%", None, "=AVERAGE(0.1748,0.1359,0.1671)", None, "Derived", WEARN + "：股東權益報酬率 FY2023 17.48%、FY2024 13.59%、FY2025 16.71%", "單一來源")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（2026-10-25 11:53 定 DCF／Comparable 各半）", "—")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（情境 B FY2029 與十年平均的中點）", "%", None, 0, None, "Derived", "PL 頁情境 B FY2029、In 頁 nim_avg", "—")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（＝十年平均淨利率 nim_avg）", "%", None, "=nim_avg", None, "Derived", "In 頁 nim_avg（wearn 2016–2025）", "—")
TP_KEYS += ["tp_r", "tp_g", "tp_roe", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]

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
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    if key in TP_KEYS:
        for j in range(1, 10): IN.cell(r, j).fill = KEY
        IN.cell(r, 7, tag + "（Andy 可改）" if tag == "Assumed" else tag)
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    INROW[key] = r; r += 1
for col, w in zip("ABCDEFGHI", [12, 60, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)
def alias(new, base, suf, years=(27, 28, 29), sep="_"):
    for y in years:
        wb.defined_names[f"{new}_{y}"] = DefinedName(f"{new}_{y}", attr_text=wb.defined_names[f"{base}{sep}{y}{suf}"].attr_text)
for nm, base, suf in (("RLO", "rel", "_lo"), ("RHI", "rel", "_hi"), ("GULO", "gmu", "_lo"), ("GUHI", "gmu", "_hi")):
    alias(nm, base, suf)
for nm, suf in (("PKLO", "_lo"), ("PKHI", "_hi")):
    alias(nm, "pkw", suf, years=(26, 27, 28, 29))
for seg in ("cn", "op", "ic"):   # 分部成長的低／高別名（Sens 用）：g_cnLO27 等
    for nm, suf in (("LO", "_lo"), ("HI", "_hi")):
        for y in (27, 28, 29):
            wb.defined_names[f"g_{seg}{nm}{y}"] = DefinedName(f"g_{seg}{nm}{y}", attr_text=wb.defined_names[f"g_{seg}{y}{suf}"].attr_text)

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億新台幣；EPS 為新台幣；股數為十億股）。藍字＝輸入；黑字＝公式。實績取自光寶科財報與月營收公告（經 stockanalysis、優分析、technews 轉述）"
Q["A2"] = ("FY2026 錨點＝1H26 財報＋3Q26 月營收實績（1–9 月累計 − 1H26）＋假設（3Q、4Q 毛利率、營業費用率、營業外；4Q 營收季增）。稅前利益＝營收 × stockanalysis 稅前淨利率（Derived，非公司原文）。"
           "3Q26 財報與法說預計 2026-10-30 前後（stockanalysis 估計；公司公告日期：未找到），截至 2026-10-25 尚未公布，公布後更新")
hdr(Q, 3, ["項目", "FY2025A", "1H26（實績）", "3Q26（營收實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1", "來源2／說明"])
S1 = "FY2025：" + SAY + "（營收 1,660.85 億、毛利 380.39 億、營業利益 167.38 億、淨利 151.16 億）；1H26：" + SAQ + "（1Q26、2Q26）"
S2 = "FY2025：優分析 2026-02-25（營收 1,660.9 億、毛利率 22.9%、營益率 10.1%、淨利 151.2 億、EPS 6.64）；1H26：優分析 2026-07-31（營收 961.1 億、毛利率 24.7%、營業利益 123.0 億、淨利 108.9 億、EPS 4.80）"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2=S2, ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"F{r}"] = ann or f"=SUM(C{r}:E{r})"; Q[f"F{r}"].number_format = fmt
    Q.cell(r, 7, tag); Q.cell(r, 8, s1); Q.cell(r, 9, s2)
qrow(4, "營收", {"B": 166.085, "C": "=43.407+52.703", "D": "=q3_rev", "E": "=D4*(1+q4g)"}, tag="Verified／Assumed（4Q 季增）", s2=S2 + "；3Q26＝1–9 月月營收累計 − 1H26")
qrow(5, "毛利", {"B": 38.039, "C": "=9.403+14.348", "D": "=D4*q_gm", "E": "=E4*q_gm"}, tag="Verified／Assumed")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="FY2025 22.9%；1H26 24.7%（公司）")
qrow(7, "營業費用", {"B": "=B5-B8", "C": "=C5-C8", "D": "=D4*q_opx", "E": "=E4*q_opx"}, tag="Derived／Assumed", s2="毛利 − 營業利益")
qrow(8, "營業利益", {"B": 16.738, "C": "=4.099+8.203", "D": "=D5-D7", "E": "=E5-E7"}, tag="Verified／Derived")
qrow(9, "營業利益率", {c: f"={c}8/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="FY2025 10.1%；2Q26 15.6%（公司）")
qrow(10, "營業外淨額", {"B": "=B11-B8", "C": "=C11-C8", "D": "=q_nop", "E": "=q_nop"}, tag="Derived／Assumed", s1="稅前 − 營業利益", s2="—")
qrow(11, "稅前利益", {"B": "=166.085*0.1187", "C": "=4.970+9.371", "D": "=D8+D10", "E": "=E8+E10"}, tag="Derived",
     s1="FY2025＝營收 × 稅前淨利率 11.87%（stockanalysis）；1H26＝1Q26 約 49.70 億＋2Q26 約 93.71 億（stockanalysis 季度稅前淨利率推算）", s2="公司揭露稅前：未找到")
qrow(12, "歸屬母公司淨利", {"B": 15.116, "C": "=3.775+7.116", "D": "=D11*q_nir", "E": "=E11*q_nir"}, tag="Verified／Derived（3Q、4Q）", s2="1Q26 37.75 億、2Q26 71.16 億（stockanalysis）；1H26 108.9 億（優分析）")
qrow(13, "淨利 ÷ 稅前（含稅與非控制權益）", {c: f"={c}12/{c}11" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(14, "股數（十億股）", {"B": "=15.116/6.64", "C": "=q_sh", "D": "=q_sh", "E": "=q_sh"}, ann="=q_sh", tag="Derived", s2="淨利 ÷ EPS")
qrow(15, "EPS", {c: f"={c}12/{c}14" for c in "BCDE"}, fmt="0.00", ann="=F12/F14", tag="Derived", s1="淨利 ÷ 股數", s2="—")
qrow(16, "  公司揭露 EPS", {"B": 6.64, "C": 4.80}, fmt="0.00", ann="", tag="Verified", s2="FY2025 6.64（wearn 同值；stockanalysis 列 6.59，口徑未載明）；2Q26 3.14（優分析）")
qrow(17, "AI 相關營收（＝營收 × AI 相關占比）", {"B": "=B4*ai_25"}, ann="=F4*ai_26", tag="Derived／Interested-party", s1="公司 AI 相關占比（In 頁 C 節）", s2="各季 AI 營收金額未揭露：未找到")
qrow(18, "  其中 AI 電源（＝AI 相關 ×（1 − 非電源占比））", {"B": "=B17*(1-x_np)"}, ann="=F17*(1-x_np)", tag="Derived／Assumed", s1="In 頁 x_np", s2="AI 電源含 PSU、Power Shelf、BBU、HVDC 電源櫃（與台達電模型的每 kW 內容口徑相同）")
qrow(19, "非 AI 營收（＝營收 − AI 相關）", {"B": "=B4-B17"}, ann="=F4-F17", tag="Derived", s1="—", s2="一般雲端電源、光電、資訊與消費電子、車用")
Q.cell(23, 1, "FY2026 錨點的檢查（參考；校準未使用券商數字）").font = BOLD
Q.cell(24, 1, "FY2026 營收年增"); Q["B24"] = "=F4/B4-1"
Q.cell(25, 1, "FY2026 營業費用年增"); Q["B25"] = "=F7/B7-1"
Q.cell(26, 1, "營業費用彈性（營業費用年增 ÷ 營收年增；In 頁 phi 的參考）"); Q["B26"] = "=B25/B24"
Q.cell(27, 1, "FY2026 營收 ÷ 共識（MarketScreener）− 1"); Q["B27"] = "=F4/cons_rev26-1"
Q.cell(28, 1, "FY2026 EPS ÷ 共識（MarketScreener）− 1"); Q["B28"] = "=F15/cons_eps26-1"
Q.cell(29, 1, "非 AI 營收 FY2026 年增"); Q["B29"] = "=F19/B19-1"
for rr in range(24, 30): Q[f"B{rr}"].number_format = "0.0%" if rr != 26 else "0.00"
Q.cell(31, 1, "營收分部（官方三大事業群；v0.2）").font = BOLD
qrow(32, "雲端及物聯網", {"B": "=B4*s25_cl", "C": "=q2_cl/(1+g1_cl)+q2_cl", "D": "=D4*s2h_cl", "E": "=E4*s2h_cl"}, tag="Interested-party／Verified／Assumed",
     s1="FY2025＝營收 × 公司比例；1H26＝2Q26 實績＋2Q26 ÷（1＋季增）", s2="3Q、4Q26 比例 Assumed（In 頁 s2h_cl）")
qrow(33, "光電", {"B": "=B4*s25_op", "C": "=q2_op/(1+g1_op)+q2_op", "D": "=D4*s2h_op", "E": "=E4*s2h_op"}, tag="Interested-party／Verified／Assumed", s1="同上", s2="同上")
qrow(34, "資訊及消費性電子", {"B": "=B4*s25_ic", "C": "=C4-C32-C33", "D": "=D4-D32-D33", "E": "=E4-E32-E33"}, tag="Derived",
     s1="FY2025＝營收 × 公司比例；1H26 與 3Q、4Q＝營收 − 其餘兩事業群（1H26 公司值 158＋季增 18% 推得 291.9 億，差 0.6 億為季增四捨五入）", s2="—")
qrow(35, "檢查：分部合計 − 合併營收", {c: f"={c}32+{c}33+{c}34-{c}4" for c in "BCDE"}, tag="Derived", s1="應為 0", s2="—")
qrow(36, "雲端及物聯網：AI 相關（＝第 17 列）", {"B": "=B17"}, ann="=F17", tag="Derived", s1="AI 相關全部歸雲端及物聯網（Assumed：AI 光電金額小，未拆）", s2="—")
qrow(37, "雲端及物聯網：非 AI（＝雲端及物聯網 − AI 相關）", {"B": "=B32-B36"}, ann="=F32-F36", tag="Derived", s1="—", s2="一般伺服器與網通電源、車用、工控、IoT")
for nm, rr in (("cl", 32), ("op", 33), ("ic", 34), ("cn", 37)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$F${rr}")
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 7), ("oi", 8), ("nop", 10), ("pt", 11), ("ni", 12), ("sh", 14), ("eps", 15), ("ai", 17), ("aip", 18), ("nai", 19)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$F${rr}")
Q.column_dimensions["A"].width = 60
for c in "BCDEF": Q.column_dimensions[c].width = 15
Q.column_dimensions["G"].width = 26; Q.column_dimensions["H"].width = 90; Q.column_dimensions["I"].width = 70
Q.freeze_panes = "B4"

# ---------------- 引擎 ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G", 2030: "H", 2031: "I", 2032: "J"}
FADE = (2030, 2031, 2032); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
BU = (2026, 2027, 2028, 2029)
umx = lambda y, s: {2026: "1", 2027: "1", 2028: f"um28_{s}", 2029: f"um29_{s}"}[y]
# 情境與行業一致：A＝份額相對值、AI 毛利率額外改善取高值；B＝基準；C＝取低值（每 kW 內容、非 AI 成長各情境相同）
SCN_P = {"A": dict(rel="RHI", gmu="GUHI"), "B": dict(), "C": dict(rel="RLO", gmu="GULO")}
DEFP = dict(scn="B", pkw="pkw", rel="rel", mf="1", gsuf="", d="d_ai", gmu="gmu", opxa="0", phi="phi", nir="nir_f")
YH = ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"]
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, YH); r += 1
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
    sec_("A. AI 部署與 AI 電源市場（D 欄＝FY2026 校準年；每 kW 內容取自台達電模型）")
    row("itc", "AI IT 資本支出 × 情境顆數係數", "US$B", {c_(y): f"=itc_{y % 100}*{umx(y, s)}" + ("" if y == 2026 else f"*{P['mf']}") for y in BU})
    row("itpg", "每 GW IT 資本支出", "US$B", {"D": "=itpg_gb*(1-vr26)+itpg_vr*vr26", **{c_(y): "=itpg_vr" for y in BU[1:]}})
    row("gw", "部署 GW（IT 電力）", "GW", {c_(y): f"={c_(y)}{R['itc']}/{c_(y)}{R['itpg']}" for y in BU}, fmt="0.0")
    row("pkw", "AI 電源內容（台達電模型）", "US$/kW", {c_(y): f"={P['pkw']}_{y % 100}" for y in BU}, fmt="#,##0")
    row("pmk", "AI 電源市場（美元）", "US$B", {c_(y): f"={c_(y)}{R['gw']}*{c_(y)}{R['pkw']}/1000" for y in BU})
    row("pmn", "AI 電源市場（新台幣）", "NT$B", {c_(y): f"={c_(y)}{R['pmk']}*fx" for y in BU}, fmt="#,##0.0")
    row("sh0", "光寶有效份額（FY2026 校準）＝ AI 電源營收 ÷ 市場", "%", {"D": f"=a26_aip/D{R['pmn']}"}, fmt="0.0%",
        note="Derived：分子為公司 AI 相關占比 × 營收 ×（1 − 非電源占比）；分母為與台達電模型相同的 AI 電源市場（含 ASIC 機架）")
    row("rel", "份額相對值（÷ FY2026）", "x", {"D": "=1", **{c_(y): f"={P['rel']}_{y % 100}" for y in BU[1:]}}, fmt="0.00")
    row("sh", "光寶有效份額", "%", {c_(y): f"=$D${R['sh0']}*{c_(y)}{R['rel']}" for y in BU}, fmt="0.0%")
    row("aip", "AI 電源營收", "NT$B", {"C": "=a25_aip", **{c_(y): f"={c_(y)}{R['pmn']}*{c_(y)}{R['sh']}" for y in BU}}, fmt="#,##0.0")
    row("ai", "AI 相關營收（＝AI 電源 ÷（1 − 非電源占比））", "NT$B", {"C": "=a25_ai", **{c_(y): f"={c_(y)}{R['aip']}/(1-x_np)" for y in BU}}, fmt="#,##0.0",
        note="非電源部分（液冷、AI 光電）隨 AI 電源同比例成長（Assumed）")
    sec_("B. 營收分部（官方三大事業群）與合計")
    gs = lambda seg, y: f"g_{seg}{P['gsuf']}{y % 100}"
    row("cn", "  雲端及物聯網：非 AI", "NT$B", {"C": "=a25_cn", "D": "=a26_cn", **{c_(y): f"={c_(y - 1)}{r}*(1+{gs('cn', y)})" for y in BU[1:]}}, fmt="#,##0.0",
        note="一般伺服器與網通電源、車用、工控、IoT；AI 相關（上列）為本分部的命題子列")
    row("cl", "雲端及物聯網", "NT$B", {c_(y): f"={c_(y)}{R['ai']}+{c_(y)}{R['cn']}" for y in FYS}, fmt="#,##0.0", note="＝AI 相關＋非 AI（子列合計＝分部）")
    row("op", "光電", "NT$B", {"C": "=a25_op", "D": "=a26_op", **{c_(y): f"={c_(y - 1)}{r}*(1+{gs('op', y)})" for y in BU[1:]}}, fmt="#,##0.0")
    row("ic", "資訊及消費性電子", "NT$B", {"C": "=a25_ic", "D": "=a26_ic", **{c_(y): f"={c_(y - 1)}{r}*(1+{gs('ic', y)})" for y in BU[1:]}}, fmt="#,##0.0")
    row("nai", "非 AI 營收（雲端非 AI＋光電＋資訊及消費性電子）", "NT$B", {c_(y): f"={c_(y)}{R['cn']}+{c_(y)}{R['op']}+{c_(y)}{R['ic']}" for y in FYS}, fmt="#,##0.0")
    row("rev", "營收", "NT$B", {c_(y): f"={c_(y)}{R['cl']}+{c_(y)}{R['op']}+{c_(y)}{R['ic']}" for y in FYS}, fmt="#,##0.0", note="＝三大事業群合計")
    row("chk", "檢查：分部合計 − 合併營收（FY2025 實績、FY2026 錨點）", "NT$B", {"C": f"=C{R['rev']}-a25_rev", "D": f"=D{R['rev']}-a26_rev"}, fmt="0.000")
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("aish", "AI 相關 ÷ 營收", "%", {c_(y): f"={c_(y)}{R['ai']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("pgw", "每 GW 光寶 AI 相關營收", "US$M/GW", {c_(y): f"={c_(y)}{R['ai']}/fx/{c_(y)}{R['gw']}*1000" for y in BU}, fmt="#,##0")
    sec_("C. 損益")
    row("gmn", "其他營收毛利率 ＝ 錨點毛利率 − 差距 × AI 占比（FY2026 校準）", "%", {"D": f"=a26_gp/D{R['rev']}-{P['d']}*D{R['aish']}"}, fmt="0.0%",
        note="Derived：讓 FY2026 毛利率等於錨點；AI 相關毛利率＝其他＋差距")
    row("gma", "AI 相關毛利率 ＝ 校準值 ＋ 額外改善", "%", {"D": f"=D{R['gmn']}+{P['d']}", **{c_(y): f"=$D${R['gmn'] + 1}+{P['gmu']}_{y % 100}" for y in BU[1:]}}, fmt="0.0%",
        note="額外改善為 800V 電源櫃與組合（Assumed；情境 A 高、C 低）")
    row("gp", "毛利", "NT$B", {"C": "=a25_gp", "D": "=a26_gp", **{c_(y): f"={c_(y)}{R['ai']}*{c_(y)}{R['gma']}+{c_(y)}{R['nai']}*$D${R['gmn']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("gm", "毛利率", "%", {c_(y): f"={c_(y)}{R['gp']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("opb", "營業費用（基準路徑）", "NT$B", {"C": "=a25_opx", "D": "=a26_opx", **{c_(y): f"={c_(y - 1)}{r}*(1+MAX(0,{P['phi']}*{c_(y)}{R['g']}))" for y in BU[1:]}}, fmt="#,##0.0")
    row("opx", "營業費用", "NT$B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", **{c_(y): f"={c_(y)}{R['opb']}+{P['opxa']}*{c_(y)}{R['rev']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("oi", "營業利益", "NT$B", {c_(y): f"={c_(y)}{R['gp']}-{c_(y)}{R['opx']}" for y in FYS}, fmt="#,##0.0")
    row("oim", "營業利益率", "%", {c_(y): f"={c_(y)}{R['oi']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("nop", "營業外淨額", "NT$B", {"C": "=a25_nop", "D": "=a26_nop", **{c_(y): "=nop_f" for y in BU[1:]}}, fmt="#,##0.0")
    row("nir", "歸屬母公司淨利 ÷ 稅前", "%", {"C": "=a25_ni/a25_pt", "D": "=a26_ni/a26_pt", **{c_(y): f"={P['nir']}" for y in BU[1:]}}, fmt="0.0%")
    row("ni", "歸屬母公司淨利", "NT$B", {"C": "=a25_ni", "D": "=a26_ni", **{c_(y): f"=({c_(y)}{R['oi']}+{c_(y)}{R['nop']})*{c_(y)}{R['nir']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("sh_", "股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", **{c_(y): f"={c_(y - 1)}{r}*(1+sh_g)" for y in BU[1:]}}, fmt="0.0000")
    row("eps", "EPS", "NT$", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['sh_']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("nim", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    # FY2030–32：收斂法（上游只到 2029；非驅動因子推導）
    fg, fm = f"fade_g{s}", f"fade_m{s}"
    for y in FADE:
        c, pc, k = FC[y], FC[y - 1], y - 2029
        ws[f"{c}{R['g']}"] = f"=$G{R['g']}+({fg}-$G{R['g']})*{k}/3"
        ws[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
        ws[f"{c}{R['nim']}"] = f"=$G{R['nim']}+({fm}-$G{R['nim']})*{k}/3"
        ws[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
        ws[f"{c}{R['sh_']}"] = f"={pc}{R['sh_']}*(1+sh_g)"
        ws[f"{c}{R['eps']}"] = f"={c}{R['ni']}/{c}{R['sh_']}"
        ws[f"{c}{R['epsg']}"] = f"={c}{R['eps']}/{pc}{R['eps']}-1"
        for key, fmt_ in (("g", "0.0%"), ("rev", "#,##0.0"), ("nim", "0.0%"), ("ni", "#,##0.0"), ("sh_", "0.0000"), ("eps", "#,##0.00"), ("epsg", "0.0%")):
            ws[f"{c}{R[key]}"].number_format = fmt_; ws[f"{c}{R[key]}"].fill = FADEFILL
    ws.cell(R["rev"], 11, "FY2030 起為收斂法：營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m（FY2032 達到），非驅動因子推導")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 光寶科損益（十億新台幣；EPS 新台幣；歸屬母公司）。FY2025A 實績、FY2026 錨點（1H26 財報＋3Q26 月營收＋假設），FY2027 起為模型。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要。顆數係數從 CY2028 開始；份額相對值、AI 毛利率額外改善從 FY2027 起依情境取高／基準／低值"
PL["A2"] = "灰底＝FY2030–FY2032 收斂法（上游只到 2029）：營收年增線性收斂到 fade_g、淨利率線性收斂到 fade_m（In 頁 J 節），非驅動因子推導"
PB = {}; r0 = 75
SCN_T = {"A": "情境 A（順利）：顆數基準（tsmc 係數 A＝1.0）；份額相對值、AI 毛利率額外改善取高值",
         "B": "情境 B（平平）：tsmc 係數 B（CY2028 ×0.8、CY2029 ×0.85）；公司特有變數取基準",
         "C": "情境 C（不如預期）：tsmc 係數 C（CY2028–29 ×0.6）；份額相對值、AI 毛利率額外改善取低值"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s, **SCN_P[s])
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目"] + YH[1:])
sr = 4; SUMR = {}
SUMK = (("rev", "營收", "NT$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("ai", "  AI 相關營收（AI 電源、BBU、液冷）", "NT$B", "#,##0.0"),
        ("cl", "  雲端及物聯網", "NT$B", "#,##0.0"), ("cn", "    其中非 AI", "NT$B", "#,##0.0"), ("op", "  光電", "NT$B", "#,##0.0"), ("ic", "  資訊及消費性電子", "NT$B", "#,##0.0"),
        ("nai", "  非 AI 營收", "NT$B", "#,##0.0"), ("aish", "  AI 相關 ÷ 營收", "%", "0.0%"), ("gm", "毛利率", "%", "0.0%"), ("oim", "營業利益率", "%", "0.0%"),
        ("ni", "歸屬母公司淨利", "NT$B", "#,##0.0"), ("nim", "  淨利率", "%", "0.0%"), ("eps", "EPS", "NT$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%"))
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
IN.cell(INROW["fade_mA"], 5, f"=PL!G{SUMR[('A', 'nim')]}")
IN.cell(INROW["fade_mB"], 5, f"=(PL!G{SUMR[('B', 'nim')]}+nim_avg)/2")
for k in ("fade_mA", "fade_mB"): IN.cell(INROW[k], 5).font = Font(color="000000")
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型 − 錨點").font = BOLD; sr += 1
CHK = []
for lab, f in (("  三情境：營收差（NT$B）", "+".join(f"ABS(D{PB[s]['rev']}-a26_rev)" for s in SC)),
               ("  情境 B：AI 相關差（NT$B）", f"D{PB['B']['ai']}-a26_ai"),
               ("  情境 B：毛利差（NT$B）", f"D{PB['B']['gma']}*D{PB['B']['ai']}+D{PB['B']['gmn']}*D{PB['B']['nai']}-a26_gp")):
    PL.cell(sr, 1, lab); PL[f"D{sr}"] = "=" + f; PL[f"D{sr}"].number_format = "0.0000"; CHK.append(sr); sr += 1
assert sr < 75, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 10; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 13

# ---------------- Units：AI 電源市場（每 kW 內容取自台達電模型）、光寶份額與台達電份額合計 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — AI 電源市場（每 kW 內容口徑與台達電模型一致，唯讀 delta/dist；本專案權威在台達電模型）、光寶有效份額與光寶＋台達電份額合計（情境 B）。部署 GW＝AI IT 資本支出（markets/network）÷ 每 GW IT 資本支出（Tokenomics，經橋接表）"
B_ = PB["B"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("gw", "部署 GW（情境 B）", "GW", lambda y: f"=PL!{FC[y]}{B_['gw']}", "0.0", "AI IT 資本支出 × 顆數係數 B ÷ 每 GW IT 資本支出"),
        ("pkw", "AI 電源內容（台達電模型）", "US$/kW", lambda y: f"=PL!{FC[y]}{B_['pkw']}", "#,##0", "唯讀 delta/dist In 頁（2026–27 Derived，2028 起 Assumed）"),
        ("prk", "每架 AI 電源內容（以 VR200 每 GW 架數換算）", "US$K/架", lambda y: f"=PL!{FC[y]}{B_['pkw']}*1000000/racks_vr/1000", "#,##0", ""),
        ("pmk", "AI 電源市場", "US$B", lambda y: f"=PL!{FC[y]}{B_['pmk']}", "#,##0.0", "含 ASIC 機架（假設每 kW 內容相同）"),
        ("sh", "光寶有效份額", "%", lambda y: f"=PL!{FC[y]}{B_['sh']}", "0.0%", "FY2026 以公司 AI 相關占比 ×（1 − 非電源占比）校準"),
        ("dsh", "台達電有效份額（台達電模型）", "%", lambda y: f"=d_sh{y % 100}", "0.0%", "唯讀 delta/dist Units 頁；台達電份額相對值不隨情境變動"),
        ("tsh", "光寶＋台達電份額合計", "%", lambda y: f"=PL!{FC[y]}{B_['sh']}+d_sh{y % 100}", "0.0%", "應低於 100%；其餘為 Flex、群電、康舒、Vertiv、麥格米特等"),
        ("osh", "其他廠商份額（＝1 − 合計）", "%", lambda y: f"=1-PL!{FC[y]}{B_['sh']}-d_sh{y % 100}", "0.0%", ""),
        ("aip", "光寶 AI 電源營收", "US$B", lambda y: f"=PL!{FC[y]}{B_['aip']}/fx", "#,##0.00", ""),
        ("ai", "光寶 AI 相關營收（含液冷等非電源）", "US$B", lambda y: f"=PL!{FC[y]}{B_['ai']}/fx", "#,##0.00", ""),
        ("rat", "光寶 AI 電源營收 ÷ 台達電 AI 電源營收（份額比）", "x", lambda y: f"=PL!{FC[y]}{B_['sh']}/d_sh{y % 100}", "0.00", "同一市場，故等於份額比"),
        ("pgw", "每 GW 光寶 AI 相關營收", "US$M/GW", lambda y: f"=PL!{FC[y]}{B_['pgw']}", "#,##0", "")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); UR[key] = ur; ur += 1
ur += 1
UN.cell(ur, 1, "讀法：光寶與台達電讀同一個 AI 電源市場（部署 GW × 台達電模型每 kW 內容）。FY2026 光寶有效份額約一成八（AI 相關扣除一成非電源）、台達電約五成八，合計約七成六，與凱基「台達電 60–65%、光寶與 Flex 第二梯隊」相符；台達電模型的份額逐年下降（光寶、Flex、中國廠追趕），本模型光寶份額相對值在情境 A 上升、B 持平、C 下降，CY2029 合計約六成三（C）到七成二（A），其餘留給其他廠商，合計不超過 100%。").font = BOLD
UNOTE = ur
UN.column_dimensions["A"].width = 66; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 14
UN.column_dimensions["G"].width = 60

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，年底落後 P/E）只作參考。本版未做先驗，無 Andy 合理 P/E 對照"
V["A2"] = "循環位置：光寶科不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照台達電；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
O = 51
V.cell(O, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, O + 1, ["項目", "A", "B", "C", "說明"])
V.cell(O + 2, 1, "現價（2026-10-22 收盤）"); V[f"B{O + 2}"] = "=px"; V[f"B{O + 2}"].number_format = "0.00"
for k, y, lab in ((3, 2026, "現價 ÷ FY2026 EPS（錨點）"), (4, 2027, "現價 ÷ FY2027 EPS（模型）"), (5, 2028, "現價 ÷ FY2028 EPS（模型）"), (6, 2029, "現價 ÷ FY2029 EPS（模型）")):
    V.cell(O + k, 1, lab)
    for j, s in enumerate(SC):
        c = "BCD"[j]; V[f"{c}{O + k}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{O + k}"].number_format = "0.0"
V.cell(O + 7, 1, "未來 12 月 EPS（模型 B）≈ 0.2 × FY2026＋0.8 × FY2027"); V[f"B{O + 7}"] = f"=0.2*PL!D{SUMR[('B', 'eps')]}+0.8*PL!E{SUMR[('B', 'eps')]}"; V[f"B{O + 7}"].number_format = "0.00"
V.cell(O + 8, 1, "  現價 ÷ 未來 12 月 EPS（模型 B）"); V[f"B{O + 8}"] = f"=px/B{O + 7}"; V[f"B{O + 8}"].number_format = "0.0"
V.cell(O + 9, 1, "近 12 月 P/E（stockanalysis，10-22）"); V[f"B{O + 9}"] = "=pe_ttm"; V[f"B{O + 9}"].number_format = "0.0"
V.cell(O + 10, 1, "前瞻 P/E（stockanalysis，10-22）"); V[f"B{O + 10}"] = "=fpe"; V[f"B{O + 10}"].number_format = "0.0"
V[f"E{O + 3}"] = "FY2026 三情境相同（錨點）"; V[f"E{O + 4}"] = "FY2027 三情境差異只來自份額相對值、AI 毛利率額外改善（顆數係數從 CY2028 開始）"
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
V.cell(R30, 1, "3. 規則值隱含股價（新台幣；約 FY2027 年底，未折現）＝ FY2027 EPS × 規則倍數（參考，不是目標價）").font = BOLD
hdr(V, R30 + 1, ["情境", "FY2027 EPS", "（同左）", "規則倍數", "隱含股價", "相對現價", "現價 ÷ FY2027 EPS"])
VAL = {}
for j, s in enumerate(SC):
    i = R30 + 2 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 3, f"=B{i}").number_format = "0.00"
    V.cell(i, 4, f"=D{CYR[s]}").number_format = "0.0"
    V.cell(i, 5, f"=C{i}*D{i}").number_format = "0.00"
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
RB0 = vr + 3
# 十年 P/E：自算＝digrin 12 月未調整價格 ÷ 年度 EPS（wearn）；HiStock 12 月 P/E（2021–2025）只作核對
PX_DEC = [48.60, 40.60, 40.60, 49.35, 49.80, 63.80, 63.80, 117.00, 99.50, 163.50]
EPS_Y = [4.05, 1.13, 3.42, 4.03, 4.31, 6.01, 6.19, 6.36, 5.21, 6.64]
PE_SELF = [f"={p}/{e}" for p, e in zip(PX_DEC, EPS_Y)]
PE_HIS = [None] * 5 + [10.89, 11.39, 18.03, 18.53, 25.95]
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "自算（digrin 12 月未調整價格 ÷ 年度 EPS）", "vals": PE_SELF, "tag": "Derived（EPS：wearn 2016–2025；10 年）"},
    {"label": "HiStock（12 月 P/E）", "vals": PE_HIS, "tag": "Interested-party（資料商；只有 2021–2025，只作核對，不採用）", "use": False}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位。採用自算（10 年）；HiStock 只有 2021–2025，與自算同量級，只作核對；companiesmarketcap 光寶頁無法取得（網站錯誤）。2016–2022 年光寶為低成長消費電子與電源零組件公司（約 10–12x，2017 年 EPS 1.13 含一次性損失使 P/E 偏高），AI 前後倍數口徑不同，規則值可能低估 AI 期的合理倍數。")
assert V.cell(RB0 + 5, 1).value == "2016" and V.cell(RB0 + 14, 1).value == "2025", (V.cell(RB0 + 5, 1).value, V.cell(RB0 + 14, 1).value)
PE10 = f"$B${RB0 + 5}:$B${RB0 + 14}"
rule_rows = [rw for rw in range(RB0, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")

# 0. 目標價（Andy 2026-10-25 11:28／11:53：DCF 與 Comparable 加權；列標籤依工作單逐字；B／C／D＝A 順利／B 平平／C 不如預期）
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
V.cell(7, 1, "情境 EPS（FY2027–FY2032）"); V["E7"] = "PL 頁各情境；FY2030 起為收斂法"; T["eps_h"] = 7
EPSR = {}
for i, y in enumerate(YRS6):
    tset(8 + i, f"eps{y}", f"  FY{y}", [f"=PL!{FC[y]}{SUMR[(s, 'eps')]}" for s in SC], "0.00"); EPSR[y] = 8 + i
V.cell(14, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）"); T["cf_h"] = 14
tset(15, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）")
CFR = {}
for i, y in enumerate(YRS6):
    tset(16 + i, f"cf{y}", f"  FY{y}", [f"={c}{EPSR[y]}*(1-$B$15)" for c in "BCD"], "0.00"); CFR[y] = 16 + i
def dcf_formula(c, r="tp_r", g="tp_g", reinv="$B$15"):
    """c＝情境欄；r、g 可替換（Sens 用）。基準日 2026 年底，t＝1 對 FY2027。負值取 0（股東有限責任）。"""
    cf = lambda y: f"Val!{c}{EPSR[y]}*(1-{reinv})"
    pv = "+".join(f"{cf(y)}/(1+{r})^{t}" for t, y in enumerate(YRS6, 1))
    return f"MAX(0,{pv}+{cf(2032)}*(1+{g})/({r}-{g})/(1+{r})^6)"
tset(22, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=MAX(0,SUMPRODUCT({c}16:{c}21,1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}21*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6)" for c in "BCD"], "0.00",
     "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底；負值取 0（Andy 12:22）")
tset(23, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}22*(1+tp_r)" for c in "BCD"], "0.00")
tset(24, "dcftp", "DCF 目標價（情境加權）", ["=tp_wA*B23+tp_wB*C23+tp_wC*D23", None, None], "0.00", "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁")
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E；非 A 股公司，同業不放 A 股）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = [("台達電（Delta Electronics）", "2308.TW", 30.51, "2026-10-22"), ("Vertiv", "VRT", 24.53, "2026-10-23"),
         ("Eaton", "ETN", 24.17, "2026-10-23"), ("Flex", "FLEX", 14.74, "2026-10-23")]
for i, (nm, tk, pe, dt) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左）"); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
V.cell(33, 1, "未列入（資料過舊）：群電 6412.TW 前瞻 P/E 13.6x（stockanalysis 快取 2026-09-04）；康舒 6282.TW 前瞻 P/E：未找到。麥格米特 002851.SZ 為 A 股，依規則不列")
tset(34, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", "AI 資料中心電源與電氣同業：台達電、Vertiv、Eaton、Flex（券商常用比較組；台達電模型的同業名單也含光寶）")
tset(35, "feps", "12 個月後的前瞻 EPS（情境加權）", [f"=tp_wA*(2/12*B{EPSR[2027]}+10/12*B{EPSR[2028]})+tp_wB*(2/12*C{EPSR[2027]}+10/12*C{EPSR[2028]})+tp_wC*(2/12*D{EPSR[2027]}+10/12*D{EPSR[2028]})", None, None], "0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(36, "comptp", "Comparable 目標價", ["=B34*B35", None, None], "0.00", "同業中位數 × 前瞻 EPS")
V.cell(38, 1, "0.3 結論").font = BOLD
hdr(V, 39, ["項目", "數值", "", "", "說明"])
tset(40, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B24+(1-tp_w_dcf)*B36", None, None], "0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁")
tset(41, "comp2", "Comparable 目標價（市場主流同業法）", ["=B36", None, None], "0.00")
tset(42, "cons", "市場共識目標價", ["=cons_tp", None, None], "0.00", "stockanalysis（S&P Global）：12 位分析師平均 314.25（頁面時間 2026-10-22）")
tset(43, "upw", "上檔空間：加權目標價", ["=B40/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-22 收盤）")
tset(44, "upc", "上檔空間：Comparable 目標價", ["=B41/px-1", None, None], "0.0%")
tset(45, "geps", "與市場共識目標價的差距：EPS 因素", ["=B35/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝現價 ÷ 前瞻 P/E（In 頁 cons_feps；期間為現在起算 12 個月，與本模型不同）")
tset(46, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B34-1", None, None], "0.0%")
V["E46"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B34,"0.0")&"x"'
tset(47, "band", "價格帶（DCF 不如預期～順利）", ['=TEXT(D23,"0.00")&"～"&TEXT(B23,"0.00")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(48, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=IF(px/{c}{EPSR[2028]}<MIN({PE10}),0,IF(px/{c}{EPSR[2028]}>MAX({PE10}),1,PERCENTRANK({PE10},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 6 節自算）的分位；只作檢查，不進目標價")
V.cell(49, 1, "目標價方法")
V["B49"] = ("DCF：股東現金流＝歸母 EPS ×（1 − g／ROE），r＝10%、g＝4%、ROE 為近三年平均約 16%（再投資率約 25%），FY2027–32 明確期＋永續；"
            "光寶 2026 年資本支出約 180 億元（約營收的 8%，含德州廠），本模型未建資本支出與折舊，未改用淨利＋折舊－資本支出，擴產期現金流可能偏高估。"
            "Comparable：台達電、Vertiv、Eaton、Flex 前瞻 P/E 中位數（不放 A 股）；台達電與 Vertiv 的 AI 占比與市占較高、歷來倍數高於光寶，光寶十年 P/E 中位數約 12x，Comparable 可能偏高估。"
            "情境對應：A＝順利（顆數基準、份額與 AI 毛利率高值）、B＝平平（顆數 ×0.8／0.85、基準）、C＝不如預期（顆數 ×0.6、低值）")
for r_ in (24, 36, 40):
    V[f"B{r_}"].fill = KEY
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 B＝基準輸入；有效份額、其他營收毛利率在每個案例內依 FY2026 錨點重算，FY2026 恆等於錨點）：對 FY2027、FY2028 的營收、AI 相關營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 AI 相關", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 AI 相關", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("份額流失（低）", "份額相對值 2027 0.85、2028 0.80、2029 0.75（台達電、Flex、中國廠競爭）", dict(rel="RLO")),
         ("份額提升（高）", "份額相對值 2027 1.10、2028 1.20、2029 1.25（新 CSP 客戶、800V 電源櫃）", dict(rel="RHI")),
         ("電源內容 低", "每 kW：2026 380、2027 500、2028–29 550（台達電模型低值）", dict(pkw="PKLO")),
         ("電源內容 高", "每 kW：2026 520、2027 700、2028 800、2029 850（台達電模型高值）", dict(pkw="PKHI")),
         ("AI 部署 ×0.8", "FY2027 起 AI IT 資本支出 ×0.8", dict(mf="mf_lo")), ("AI 部署 ×1.2", "×1.2", dict(mf="mf_hi")),
         ("AI 毛利率差距 8pt", "AI 相關毛利率只高 8 個百分點", dict(d="d_ai_lo")), ("AI 毛利率差距 16pt", "高 16 個百分點", dict(d="d_ai_hi")),
         ("AI 毛利率額外改善 低", "2027 −1pt、2028 −1.5pt、2029 −2pt", dict(gmu="GULO")), ("AI 毛利率額外改善 高", "2027 +1pt、2028 +1.5pt、2029 +2pt", dict(gmu="GUHI")),
         ("非 AI 分部成長 低", "雲端非 AI、光電 0%；資訊及消費性電子 −5%（2027–29）", dict(gsuf="LO")), ("非 AI 分部成長 高", "雲端非 AI、光電 +10%；資訊及消費性電子 +5%", dict(gsuf="HI")),
         ("營業費用率 +1pt", "營業費用額外 +1% × 營收", dict(opxa="opx_adj_hi")),
         ("營業費用彈性 1.0", "基準 0.7", dict(phi="phi_hi")),
         ("淨利 ÷ 稅前 0.73", "基準 0.76", dict(nir="nir_f_lo")), ("淨利 ÷ 稅前 0.78", "", dict(nir="nir_f_hi"))]
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.0"), ("D", "ai", 2027, "#,##0.0"), ("E", "gm", 2027, "0.0%"), ("F", "eps", 2027, "0.00"),
                           ("H", "rev", 2028, "#,##0.0"), ("I", "ai", 2028, "#,##0.0"), ("J", "gm", 2028, "0.0%"), ("K", "eps", 2028, "0.00")):
        SE[f"{c}{i}"] = f"={pre}{FC[y]}{R[key]}"; SE[f"{c}{i}"].number_format = fmt
    SE[f"G{i}"] = f"=F{i}/F$4-1"; SE[f"L{i}"] = f"=K{i}/K$4-1"
    for c in "GL": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 B）", "PL 區塊 B", PB["B"], "PL!")
base_row = 50; SB = {}
for nm, setting, ov in cases:
    SB[nm] = engine(SE, base_row, f"案例區塊：{nm}（情境 B）", scn="B", **ov)
    base_row = SB[nm]["end"] + 2
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm], "")
nn = 5 + len(cases)
put_case(nn, "情境 A（全部）", "顆數係數 A＋份額相對值、AI 毛利率額外改善高值", PB["A"], "PL!")
put_case(nn + 1, "情境 C（全部）", "顆數係數 C＋份額相對值、AI 毛利率額外改善低值", PB["C"], "PL!")
SENS_LAST = nn + 1; SENS_NOTE = nn + 3
SE.cell(nn + 3, 1, "讀法：基準為情境 B（顆數係數 B 從 CY2028 起）。每個案例的 FY2026 有效份額與其他營收毛利率依該案例重算，所以 FY2026 恆等於錨點；「電源內容」案例的 FY2026 份額會反向調整（內容低 → 份額高），所以只影響 FY2027 起的成長路徑。AI 相關中非電源占比（x_np）只影響份額，不影響營收與 EPS，未列入。")
TPS0 = nn + 5
SE.cell(TPS0, 1, "目標價參數（只改 DCF 參數；PL 收斂法的 fade_g 不隨 tp_g 變動）").font = BOLD
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
assert TPS_LAST < 50, TPS_LAST
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 56
for c in "CDEFGHIJKL": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 共識（MarketScreener、stockanalysis）、台達電模型的 AI 電源市場與份額、凱基份額。只作事後對照，不為對上數字而調參數。FY2027 起共識只有一個來源"
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
rec("FY2026 營收 vs 共識（MarketScreener）", "=cons_rev26", pl("rev", 2026), "MarketScreener（參考價 2026-10-22）", "stockanalysis 2,193 億", fmt="#,##0.0")
rec("FY2026 EPS vs 共識（MarketScreener）", "=cons_eps26", pl("eps", 2026), "MarketScreener", "stockanalysis 10.81（11 家）；3Q26 財報（10-30 前後）前")
rec("FY2027 營收 vs 共識（MarketScreener）", "=cons_rev27", pl("rev", 2027), "MarketScreener", "", fmt="#,##0.0")
rec("FY2027 EPS vs 共識（MarketScreener）", "=cons_eps27", pl("eps", 2027), "MarketScreener", "差距說明見 README（部署 GW 2027 年約 1.8 倍）")
rec("FY2028 營收 vs 共識（MarketScreener）", "=cons_rev28", pl("rev", 2028), "MarketScreener", "", fmt="#,##0.0")
rec("FY2028 EPS vs 共識（MarketScreener）", "=cons_eps28", pl("eps", 2028), "MarketScreener", "")
RC.cell(rr_, 1, "B. AI 電源市場與份額（與台達電模型一致性檢查）").font = BOLD; rr_ += 1
rec("CY2026 AI 電源市場（US$B）vs 台達電模型（情境 A）", "=d_pmk26", pl("pmk", 2026), "delta/dist Units 頁（唯讀）", "情境 A 應為 0%（同一部署 GW × 同一每 kW 內容）", fmt="#,##0.0")
rec("CY2027 AI 電源市場（US$B）vs 台達電模型（情境 A）", "=d_pmk27", pl("pmk", 2027), "同上", "", fmt="#,##0.0")
rec("CY2028 AI 電源市場（US$B）vs 台達電模型（情境 A）", "=d_pmk28", pl("pmk", 2028), "同上", "情境 B、C 低於台達電情境 A：顆數係數", fmt="#,##0.0")
for y in BU:
    rec(f"CY{y} 光寶＋台達電 AI 電源有效份額合計", None, (lambda yy: (lambda s: f"=PL!{FC[yy]}{PB[s]['sh']}+d_sh{yy % 100}"))(y), "本模型＋delta/dist Units 頁",
        "應低於 100%；其餘為 Flex、群電、康舒、Vertiv、麥格米特等" if y == 2026 else "—", fmt="0.0%")
rec("FY2026 台達電模型份額 vs 凱基台達電市占（中點）", "=kgi_dsh", lambda s: "=d_sh26", "凱基 2025-09（Interested-party）", "口徑不同：模型分母含 ASIC 機架、HVDC、BBU", fmt="0.0%", pt=True)
rec("CY2027 光寶 AI 電源營收 ÷ 台達電 AI 電源營收", None, lambda s: f"=PL!E{PB[s]['sh']}/d_sh27", "本模型＋delta/dist", "同一市場，等於份額比", fmt="0.00")
rec("CY2027 每 GW 光寶 AI 相關營收（US$M）", None, pl("pgw", 2027), "—", "台達電模型 CY2027 每 GW AI 營收（電源＋液冷）約 470", fmt="#,##0")
RC.cell(rr_, 1, "C. 校準").font = BOLD; rr_ += 1
rec("其他營收毛利率（Derived）", None, lambda s: f"=PL!D{PB[s]['gmn']}", "FY2026 錨點", "AI 相關毛利率＝此值＋d_ai", fmt="0.0%")
rec("FY2027 毛利率 vs FY2026（百分點）", "=a26_gp/a26_rev", pl("gm", 2027), "錨點", "AI 占比上升使毛利率上升", fmt="0.0%", pt=True)
rec("FY2027 營業利益率 vs 2Q26 實績 15.6%（百分點）", "=8.203/52.703", pl("oim", 2027), "stockanalysis 2Q26", "", fmt="0.0%", pt=True)
rec("4Q26 EPS（模型）", None, lambda s: "=Q!E15", "—", "3Q26 財報 10-30 前後公布", fmt="0.00")
RC.cell(rr_, 1, "D. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以部署路徑取代「AI IT 資本支出 ÷ 每 GW IT 資本支出」"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.cell(rr_, 1, "E. 目標價對帳（模型欄為情境加權值，只填 A 欄）").font = BOLD; rr_ += 1
for lab, ref, mdl, src in (("12 個月後的前瞻 EPS（情境加權）vs 共識前瞻 EPS", "=cons_feps", f"=Val!B{T['feps']}", "stockanalysis 前瞻 P/E（10-22）"),
                           ("目標價（DCF 與 Comparable 加權）vs 市場共識目標價", "=cons_tp", f"=Val!B{T['tpw']}", "stockanalysis（S&P Global）12 位分析師"),
                           ("Comparable 目標價 vs 市場共識目標價", "=cons_tp", f"=Val!B{T['comptp']}", "同上")):
    RC.cell(rr_, 1, lab); RC.cell(rr_, 2, ref).number_format = "0.00"; RC.cell(rr_, 3, mdl).number_format = "0.00"
    RC.cell(rr_, 4, f"=C{rr_}/B{rr_}-1").number_format = "0.0%"; RC.cell(rr_, 9, src); RC.cell(rr_, 10, "共識只作對帳，不作校準"); rr_ += 1
RC.cell(rr_, 1, "F. 營收分部（v0.2：官方三大事業群；共識的分部預估：未找到）").font = BOLD; rr_ += 1
rec("FY2026 雲端及物聯網年增 vs 公司「雲端運算全年年增超過 70%」", 0.70, lambda s: f"=PL!D{PB[s]['cl']}/PL!C{PB[s]['cl']}-1", "富果 2026-07-31（公司法說，Interested-party）",
    "口徑不同：公司說的是雲端運算（雲端及物聯網的一部分）；FY2026 雲端及物聯網年增由錨點推得", fmt="0.0%", pt=True)
for lab, v, note in (("2Q26 雲端及物聯網營業利益率（實績）", "=6.9/28.9", "營業利益 69 億、季增 102%"), ("2Q26 光電營業利益率（實績）", "=0.729/8.0", "營業利益 7.29 億"),
                     ("2Q26 資訊及消費性電子營業利益率（實績）", "=1.4/15.8", "營業利益 14 億"),
                     ("2Q26 三事業群營業利益合計 − 合併營業利益（未分攤，NT$B）", "=6.9+0.729+1.4-8.203", "負向未分攤費用；本模型營業利益仍由合併毛利率與費用推導，分部利益只作對帳")):
    RC.cell(rr_, 1, lab); RC.cell(rr_, 2, v).number_format = "0.0%" if "率" in lab else "0.00"
    RC.cell(rr_, 9, "Investing.com 2026-07-31（Lite-On Q2 2026 slides 摘要）"); RC.cell(rr_, 10, note); rr_ += 1
RECON_LAST = rr_ - 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝光寶科財報、月營收公告與法說會，經資料商與媒體轉述；券商與媒體估計皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("二手：優分析 2026-07-31（2Q26、1H26 損益、雲端與物聯網營收與營業利益）", UA2Q),
        ("二手：富果 2026-07-31 法說摘要（AI 占比超過 30%、BBU 占雲端 20–25%、800V 時程、資本支出 180 億、產能）", FUG2Q),
        ("二手：優分析 2026-02-25（FY2025、4Q25 損益、AI 占比超過 20%）", UAFY),
        ("二手：富果 2026-02-25 法說摘要（FY2025 部門占比、2026 指引、資本支出）", FUGFY),
        ("二手：technews 2026-10-08（9 月營收 188.48 億、前 9 月 1,535.59 億）", TN9),
        ("資料商：nstock（10-22 收盤、9 月營收）", NST),
        ("資料商：stockanalysis（現價、近 12 月與前瞻 P/E、股數、共識目標價；同業前瞻 P/E）", SA),
        ("資料商：stockanalysis 季度與年度損益（營收、毛利、營業利益、淨利、稅前淨利率）", SAQ),
        ("資料商：stockanalysis forecast（2026-09-23 目標價區間、FY2026 營收與 EPS）", SAF),
        ("資料商：MarketScreener（FY2026–28 共識營收、淨利、EPS）", MSV),
        ("資料商：wearn（2012–2025 年度 EPS、淨利率、ROE）", WEARN),
        ("資料商：digrin（2014–2025 年 12 月價格，未調整）", DIGRIN),
        ("資料商：HiStock（12 月 P/E 2021–2025）", HIS),
        ("券商：凱基 Power supply sector 2025-09-26（台達電市占 60–65%、光寶第二梯隊、每架電源內容）", KGI),
        ("市場層：markets/network、markets/chip_bridge、tsmc 成品；delta 成品（AI 電源每 kW 內容、台達電份額，本專案權威）（本 repo，唯讀）", ""),
        ("無法取得：companiesmarketcap 光寶 P/E 頁（網站錯誤）；10-23 收盤；7、8 月個別營收", "未找到"),
        ("待補（人工）：3Q26 財報與法說（10-30 前後）；10 月營收（11-10 前後）；部門別毛利率；AI 電源與 BBU 的季度金額；光寶 AI 電源市占的量化來源；資本支出與折舊路徑；前瞻 P/E 十年歷史（Bloomberg）；FY2027 起共識第二來源", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 100; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("光寶科（2301.TW）個股模型 v0.2（AI 半導體 Project：電源與散熱次產業；AI 伺服器電源 PSU、電源櫃、BBU）　2026-10-28", True),
 ("版本：v0.2（2026-10-28：依個股模型原則重建營收層——營收列改為官方三大事業群〔雲端及物聯網、光電、資訊及消費性電子〕，AI 相關為雲端及物聯網的子列，分部合計＝合併營收；非 AI 單一成長率改為三條分部成長率；每 kW 內容改讀 Siliconomics）；v0.1.1（2026-10-25：補上檔空間級距、投資建議、重點觀察，數值不變）；v0.1（初版）。AI 電源每 kW 內容沿用台達電模型（delta/dist In 頁，本專案權威，口徑一致），本模型只建光寶份額；Units 與 Recon 頁核對光寶＋台達電份額合計。", False),
 ("模型命題：光寶科的 AI 營收＝部署 GW（AI IT 資本支出 ÷ 每 GW IT 資本支出）× AI 電源每 kW 內容（台達電模型）× 光寶份額 ÷（1 − 非電源占比）；非 AI（一般雲端電源、光電、資訊與消費電子、車用）約占七成，另設低成長。AI 相關營收 2025 年約兩成、2026 年約三成，毛利率高於其他業務，使 AI 部署成為 EPS 的主要驅動。", False),
 ("核心問題：光寶在台達電主導的 AI 電源市場（800V HVDC 電源櫃、BBU、新 CSP 客戶）能不能把有效份額由約一成八往上拉；消費電子衰退與零組件成本會吃掉多少；現價約 20 倍前瞻 P/E 隱含什麼路徑。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029 由驅動因子推導，FY2030–FY2032 為收斂法。金額：十億新台幣；市場以美元計、以匯率 32.0 換算（與台達電模型相同）。口徑：歸屬母公司淨利（IFRS）。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("AI 營收：markets/network 的 AI IT 資本支出 × 同字母顆數係數（tsmc；＝markets/memory SCEN_UNITS）÷ 每 GW IT 資本支出（Tokenomics，經橋接表）＝部署 GW；× 每 kW 電源內容（台達電模型 450 → 600 → 680 → 720 美元）× 匯率 × 光寶有效份額（FY2026 以公司 AI 相關占比校準）× 份額相對值（Assumed）÷（1 − 非電源占比）。", False),
 ("營收分部：雲端及物聯網＝AI 相關（命題模組）＋非 AI（FY2026 錨點 × 年增 5%，Assumed）；光電＝FY2026 錨點 × 年增 5%（Assumed）；資訊及消費性電子＝FY2026 錨點 × 年增 0%（Assumed，公司 2026 資訊產品 −6～−8%、低軌衛星電源倍增）。FY2026 各事業群＝1H26 實績（2Q26 簡報＋季增）＋3Q、4Q 依 2Q26 比例（Assumed）。毛利率：AI 相關比其他營收高 d_ai 12pt（Assumed，依 2Q26 雲端部門營益率約 24% vs 其他約 5–6%），其他毛利率為 FY2026 殘差（Derived）；另加 AI 毛利率額外改善（情境 A 高、C 低）。營業費用：彈性 phi 0.7（Assumed）。淨利 ÷ 稅前 0.76（1H26 實績）。", False),
 ("校準：FY2026 ＝ 1H26 財報＋3Q26 月營收實績（1–9 月累計 − 1H26）＋假設（3Q、4Q 毛利率 26%、營業費用率 11.8%、4Q 季增 3%）。3Q26 財報與法說預計 2026-10-30 前後，截至 2026-10-25 尚未公布。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，三情境權重 0.3／0.5／0.2，負值取 0）＋0.5 × Comparable 目標價（台達電、Vertiv、Eaton、Flex 前瞻 P/E 中位數 × 12 個月後前瞻 EPS；不放 A 股）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解。十年 P/E 規則值 × FY2027 EPS 保留為參考（不是目標價）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m，FY2032 達到；PL 頁灰底）。十年平均淨利率 6.5%（wearn 2016–2025，單一來源）。", False),
 ("情境與行業一致：A＝順利（tsmc 顆數係數 A、份額相對值與 AI 毛利率額外改善取高值）；B＝平平（係數 B、取基準）；C＝不如預期（係數 C、取低值）；每 kW 電源內容（台達電模型）與非 AI 成長各情境相同。", False),
 ("份額合計：光寶與台達電讀同一個 AI 電源市場。FY2026 光寶有效份額約一成八、台達電模型約五成八，合計約七成六；之後台達電模型份額逐年下降（光寶、Flex 追趕），光寶情境 B 持平、A 上升、C 下降，CY2029 合計約六成三（C）到七成二（A）（Units 頁、Recon B 節）。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司數字取自財報與月營收（經資料商與媒體轉述）；券商與資料商為 Interested-party。", False),
 ("", False),
 ("已知限制：(1) AI 相關營收只有公司年度占比（2025 年超過 20%、2026 年超過 30%），季度金額未揭露；(2) 份額相對值與非電源占比為 Assumed；(3) 每 kW 內容沿用台達電模型（2028 起 Assumed）；(4) 部門別毛利率未揭露，AI 毛利率差距由部門營益率推估；(5) 未建資本支出與折舊（2026 年約 180 億），DCF 現金流可能偏高估；(6) 十年 P/E 只有一個採用來源，2016–2022 年為 AI 前低倍數；(7) Comparable 同業（台達電、Vertiv）倍數高於光寶歷史；(8) FY2027 起共識只有 MarketScreener 一個來源；(9) 3Q26 財報後更新錨點；R4 產業部署路徑未建。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ---------------- v0.1.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 14:00、14:17；既有數值與儲存格位置不變） ----------------
_UPW = [r_ for r_ in range(1, 60) if V.cell(r_, 1).value == "上檔空間：加權目標價"]
assert _UPW == [43], _UPW
assert V.max_row == 117, V.max_row   # 0.3 結論後只有一列空白，改放在工作表現有最後一列之後
GR_ROW, RC_ROW = 118, 119
for _r in (GR_ROW, RC_ROW):
    assert all(V.cell(_r, _c).value is None for _c in range(1, V.max_column + 1)), _r
V.cell(GR_ROW, 1, "上檔空間級距")
V.cell(GR_ROW, 2, '=IF(B43>0.5,"大於 50%",IF(B43>0.25,"25%–50%",IF(B43>=0,"0%–25%",IF(B43>=-0.25,"0%～−25%",IF(B43>=-0.5,"−25%～−50%","小於 −50%")))))').number_format = "@"
V.cell(GR_ROW, 5, "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）")
V.cell(RC_ROW, 1, "投資建議")
V.cell(RC_ROW, 2, '=IF(B43>0.5,"強力買進",IF(B43>0.25,"買進",IF(B43>=-0.25,"中立",IF(B43>=-0.5,"放空","強力放空"))))').number_format = "@"
V.cell(RC_ROW, 5, "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
OBS = [
       "目標價 NT$424（DCF 319、Comparable 529 各半），比現價 255.5（2026-10-22）高 65.9%，級距 大於 50%；Comparable 單獨為 +107.1%。主因是 EPS 不是倍數：本模型 12 個月後前瞻 EPS 21.73 比共識前瞻 EPS 12.97 高 67.6%（共識為現在起 12 個月，早約一年；本模型情境 B 的 FY2027 EPS 18.74 仍高 44.5%），同業（台達電、Vertiv、Eaton、Flex）前瞻 P/E 中位數 24.4 倍與共識目標價 314.25 隱含的 24.2 倍相當；差距來自 AI 電源份額與 800V 電源櫃放量假設（Assumed）。DCF 單獨也有 +24.8%。",
       "情境分歧大：情境 A／C（顆數、份額、AI 毛利率全部）FY2028 EPS 比 B 高 43.9%／低 34.8%，DCF 價格帶 121～531；單一變數中份額流失 −14.8%、AI 部署 ×0.8 −14.7%。",
       "追蹤 3Q26 財報與法說（2026-10-30 前後）與 10 月營收（11-10 前後）：本版 3Q、4Q 毛利率 26% 為假設；AI 相關營收只有年度占比（2026 年超過 30%），季度金額未揭露。"
      ]   # 重點觀察（與 README.md 相同文字）
_n0 = RD.max_row
RD.cell(_n0 + 1, 1, "v0.1.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（Val 頁 0.3 結論加「上檔空間級距」「投資建議」兩列、README 頁加重點觀察三點；數值不變，既有儲存格位置不變）。")
RD.cell(_n0 + 2, 1, f'="投資建議："&Val!B{RC_ROW}&"（依上檔空間固定規則）"').font = BOLD
for _i, _t in enumerate(OBS, 1): RD.cell(_n0 + 2 + _i, 1, f"重點觀察 {_i}：" + _t)

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "liteon")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.2.1", "2026-10-31")  # v0.2.1：In 頁版面（個股模型原則第 8 節；數值不變）
wb.save(sys.argv[1])
if os.environ.get("LITEON_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE,
               "UNOTE": UNOTE, "T": T, "TPS0": TPS0, "TPS_LAST": TPS_LAST, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0, "block_end": block_end},
              open(os.environ["LITEON_ROWMAP"], "w"), ensure_ascii=False, indent=1)

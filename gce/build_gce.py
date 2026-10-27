# 金像電（2368.TW）收支與估值模型 v0.1.1（AI 半導體 Project：載板與 PCB 次產業；伺服器與交換器 PCB、高層數板）
# 用法：python3 gce/build_gce.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；每 GW PCB 金額以美元計（取自勝宏模型成品）、以匯率 fx 換算；EPS：新台幣。
# 口徑：歸屬母公司淨利（IFRS）。樣板：victory_giant/（勝宏：每 GW PCB 金額 × 份額 × 價格，產能上限）。
# 每 GW PCB 金額口徑與勝宏一致：唯讀 victory_giant/dist 的 Units 頁（本專案權威），本模型只加金像電的參與係數；Recon 頁對照台光電每 GW CCL 金額。
# 情境與行業一致（Andy 2026-10-25 13:13）：A／B／C 讀 tsmc 同字母顆數係數（＝markets/memory SCEN_UNITS）；公司特有變數（份額、價格、AI 毛利率改善）A 取高值、B 基準、C 低值。
# FY2030–FY2032 為收斂法（上游只到 2029）；目標價＝DCF 與 Comparable 加權（Val 第 0 節）。
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
EMC_F = only("emc/dist/*.xlsx"); VG_F = only("victory_giant/dist/*.xlsx")
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
EMC = {}
for r in openpyxl.load_workbook(EMC_F, data_only=True)["Units"].iter_rows(min_row=4):
    lab = r[0].value
    for key, want in (("ccl", "每 GW 高速 CCL 金額（合計，不含價格指數）"), ("emcpg", "每 GW 台光電 AI 相關營收")):
        if lab == want:
            EMC[key] = ([r[2 + j].value for j in range(4)], f"{rel(EMC_F)}｜Units!C{r[0].row}:F{r[0].row}")
assert len(EMC) == 2 and all(v is not None for v in EMC["ccl"][0]), EMC
VG = {}
for r in openpyxl.load_workbook(VG_F, data_only=True)["Units"].iter_rows(min_row=4):
    lab = r[0].value
    for key, want in (("cmp", "每 GW PCB：加速器卡／運算板"), ("sw", "每 GW PCB：交換器板"), ("mid", "每 GW PCB：中板／背板／其他"),
                      ("pgw", "每 GW AI 高速 PCB 金額（合計，不含價格指數）"), ("pge", "每 GW 勝宏 AI 相關營收"), ("sh", "勝宏有效份額")):
        if lab == want:
            VG[key] = ([r[2 + j].value for j in range(4)], f"{rel(VG_F)}｜Units!C{r[0].row}:F{r[0].row}")
assert len(VG) == 6 and all(v is not None for k in VG for v in VG[k][0]), VG

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B），每 GW PCB 金額為百萬美元；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文＝金像電財報（公開資訊觀測站）、月營收公告與法說會（經媒體轉述）；公司的訂單、產品組合與產能說法為 Interested-party；券商與媒體估計皆為 Interested-party"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
UA2Q = "https://uanalyze.com.tw/articles/7465753217"
UAAMZ = "https://uanalyze.com.tw/articles/9313556134"
UA8M = "https://uanalyze.com.tw/articles/6187254928"
UAFY = "https://uanalyze.com.tw/articles/8548248766"
UDNFY = "https://money.udn.com/money/amp/story/5607/9371693"
TN9 = "https://finance.technews.tw/2026/10/08/gce-2368-202609-financial-report/"
MW10 = "https://www.moneyweekly.com.tw/_Article?AID=255441"
MW9 = "https://www.moneyweekly.com.tw/_Article?AID=250910"
FUG = "https://blog.fugle.tw/post/earnings-call-2368-2025-11-28"
CNY = "https://m.cnyes.com/news/id/6169890"
TNFB = "https://finance.technews.tw/2026/05/13/three-pillars/"
SA = "https://stockanalysis.com/quote/tpe/2368/"
SAQ = "https://stockanalysis.com/quote/tpe/2368/financials/?p=quarterly"
SAF = "https://stockanalysis.com/quote/tpe/2368/forecast/"
MSV = "https://www.marketscreener.com/quote/stock/GOLD-CIRCUIT-ELECTRONICS--6496480/valuation/"
WEARN = "https://stock.wearn.com/financial.asp?kind=2368"
DIGRIN = "https://www.digrin.com/stocks/detail/2368.TW/price"

sec("A. AI 部署規模（讀 markets/network、markets/chip_bridge、tsmc 成品；情境 A 基準，B／C 依同字母顆數係數縮放）")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}", "US$B", None, ITC[0][j], None, "Derived", ITC[1],
        "加速器顆數 × 每顆 IT capex（含 ASIC）；本模型以此除以每 GW IT 資本支出得到部署 GW（與勝宏、台光電模型相同）" if y == 26 else "同上")
inp("itpg_gb", "每 GW IT 資本支出（GB300）", "US$B", None, BR["capit"][0], None, "Derived", BR["capit"][2], "Tokenomics IF_CapexIT（經橋接表；唯讀）")
inp("itpg_vr", "每 GW IT 資本支出（VR200；2027 起）", "US$B", None, BR["capit"][1], None, "Derived", BR["capit"][2], "同上；Rubin Ultra 與 ASIC 機架未另訂（Assumed：同 VR200）")
inp("vr26", "2026 年部署中 Rubin 世代占比（混合每 GW IT 資本支出用）", "%", *SIL.vals("vr26"), "Interested-party", "TrendForce 2026 組合 Rubin 22%（經 markets/memory；與勝宏、台光電模型相同）", "—")
inp("racks_gb", "每 GW 機架數（GB300）", "架", None, BR["racks"][0], None, "Derived", BR["racks"][2], "Recon 用")
inp("racks_vr", "每 GW 機架數（VR200）", "架", None, BR["racks"][1], None, "Derived", BR["racks"][2], "Units 頁每架等值金額用")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝順利（基準）；B＝平平；C＝不如預期；與台積電、勝宏模型相同（＝markets/memory SCEN_UNITS，建置時核對）")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")
inp("fx", "匯率（新台幣 ÷ 美元）", "NT$/US$", 30.5, 32.0, 33.0, "Derived", "與台光電、台達電、廣達模型相同（emc/dist In 頁 fx＝32.0）", "10-23 收盤：未找到")

sec("B. 每 GW AI 高速 PCB 金額（口徑與勝宏一致：唯讀 victory_giant/dist Units 頁）與金像電份額")
VGTAG = "唯讀勝宏模型成品（本專案每 GW PCB 金額權威；2026 Assumed、之後 × 世代指數 Analogy／Assumed）"
for nm, lab in (("cmp", "加速器卡／運算板（GPU／ASIC 模組、運算匣主板）"), ("sw", "交換器板（NVLink 交換匣＋scale-out 800G／1.6T 交換器）"), ("mid", "中板／正交背板／CPX 與其他")):
    for j, y in enumerate((26, 27, 28, 29)):
        inp(f"vg_{nm}{y}", f"每 GW PCB：{lab} 20{y}（勝宏模型）", "US$M/GW", None, VG[nm][0][j], None, "Derived", VG[nm][1], VGTAG if y == 26 else "同上")
inp("w_mid", "金像電對中板／背板／其他的參與係數（1＝與加速器卡、交換器板相同）", "x", 0.0, 0.5, 1.0, "Assumed",
    "金像電以 ASIC 伺服器主板（Trainium、TPU）與 800G 交換器板為主（優分析 2026-09-09；Interested-party）；NVIDIA VR200 中板、Kyber 正交背板的參與度：未找到", "低＝不參與、高＝與其他兩項相同")
inp("pgwf", "每 GW 金額成長倍率（套在勝宏模型 2027–29 每 GW 金額相對 2026 的增量上；Sens 用）", "x", 0.7, 1.0, 1.3, "Assumed",
    "勝宏模型的世代升級主要來自 NVIDIA VR200 中板與 CPX、Kyber 背板；ASIC 主板改 HDI＋HLC 混壓也提高每 GW 金額（優分析 2026-09-09）", "基準＝與勝宏模型相同")
for y, v in ((27, (0.85, 1.00, 1.08)), (28, (0.75, 0.95, 1.08)), (29, (0.70, 0.90, 1.05))):
    inp(f"rel_{y}", f"金像電份額 20{y} ÷ 2026（相對值；情境 A 取高、B 基準、C 取低）", "x", *v, "Assumed",
        "風險：ASIC 主板設計由高層數（HLC）轉 HDI＋HLC 混壓（Trainium4、TPU 新世代），2026-06 股價修正逾五成；公司以 3–4 階 HDI 混壓因應（優分析 2026-09-09，Interested-party）" if y == 27 else "同上",
        "利多：另一家美國雲端業者 ASIC 主板訂單、Amazon 私募入股（2026-09-29）；對手健鼎、TTM、勝宏、沪电擴產" if y == 27 else "—")
for y, v in ((27, (0.96, 1.00, 1.03)), (28, (0.92, 0.97, 1.02)), (29, (0.88, 0.94, 1.00))):
    inp(f"pp_{y}", f"PCB 價格指數 20{y}（2026＝1；同規格降價或漲價，不含世代升級；情境 A 取高、C 取低）", "x", *v, "Assumed",
        "分析師：3Q26 成長動能含價格改善、產能吃緊（優分析 2026-09-09，Interested-party）" if y == 27 else "同上",
        "新產能（中壢、蘇州二廠、泰國二廠）2027 年下半年起開出後價格回落（Assumed）" if y == 27 else "—")

sec("C. 產能（全產全銷：營收上限＝FY2026 營收 × 產值產能指數 × 價格指數；各情境相同）")
CAPS = "公司：2026 年資本支出約 190 億元（2025 年約 50–60 億）；台灣、蘇州、常熟、泰國同步擴產；台灣 2026 下半年產能較去年增約五成、泰國產能年增 5–6 倍（理財周刊 2026-09-23，Interested-party）"
CAPD = "蘇州二廠、泰國二廠 2027 年下半年量產；中壢新廠投資 138 億元（2026-09-04 董事會）；2027–29 每年一座新廠（優分析 2026-09-09）；2026-09-29 再通過 79 億元土地與資本支出"
for y, v in ((27, (1.30, 1.45, 1.60)), (28, (1.60, 1.90, 2.15)), (29, (1.90, 2.30, 2.60))):
    inp(f"cx_{y}", f"平均產值產能 20{y} ÷ FY2026 營收（產能指數；以 FY2026 產品價格計）", "x", *v, "Assumed",
        CAPS if y == 27 else ("同上；每年一座新廠" if y == 29 else "同上"),
        CAPD if y == 27 else ("4Q26 年化營收約為 FY2026 的 1.2 倍（Q 頁），FY2027 指數下限據此" if y == 28 else "—"))

sec("D. 營收結構錨點（公司與媒體揭露比例 × 營收）")
inp("ai_25", "2025 年 AI 相關營收（AI 伺服器 GPU／ASIC 主板、800G 交換器）÷ 營收", "%", 0.20, 0.30, 0.40, "Interested-party",
    UA8M + "：AI 占比由約兩成升至約五成（2026-09-09）", "公司法說 2025-11：3Q25 伺服器（一般＋AI）占 80%、網通 10%、筆電 7%、其他 3%，AI 未單獨揭露（富果整理）")
inp("g_nai26", "2026 年非 AI 營收年增（一般伺服器、網通、筆電、其他）", "%", 0.10, 0.20, 0.30, "Assumed",
    "一般伺服器平台換代與價格改善（Assumed）；FY2026 AI 相關營收＝錨點營收 − 非 AI（Derived）", "理財周刊 2026-10-14：分析師稱 2Q26 AI 相關約占 40–50%、伺服器約 78%")
for y, v in ((27, (-0.05, 0.05, 0.10)), (28, (-0.05, 0.03, 0.08)), (29, (-0.05, 0.03, 0.08))):
    inp(f"g_nai{y}", f"非 AI 營收年增 20{y}", "%", *v, "Assumed",
        "產能優先給 AI 與高階板（Assumed）" if y == 27 else "同上", "—")

sec("E. 毛利率")
inp("d_ai", "AI 相關（高層數 AI 伺服器板、800G 交換器板）毛利率高於其他營收的幅度", "pt", 0.06, 0.10, 0.14, "Assumed",
    "部門別毛利率未揭露：未找到；毛利率 FY2024 29.3% → FY2025 32.8% → 1H26 34.7%，同期 AI 占比由約兩成升至約五成", "泰國廠由一般伺服器板轉 AI 高單價板（優分析 2026-09-09）")
for y, v in ((27, (-0.01, 0.01, 0.02)), (28, (-0.01, 0.015, 0.03)), (29, (-0.01, 0.02, 0.035))):
    inp(f"gmu_{y}", f"AI 相關毛利率額外改善 20{y}（新廠爬坡完成、組合；不含價格；情境 A 取高、C 取低）", "pt", *v, "Assumed",
        "分析師：3Q26 起毛利率逐季上升（優分析 2026-09-09，Interested-party）；新廠（楊梅租賃廠 3Q26、中壢、泰國二廠）爬坡期折舊與良率壓力" if y == 27 else "同上", "—")

sec("F. FY2026 錨點（1H26 財報＋3Q26 月營收實績＋假設；3Q26 財報與法說預計 2026-11 上旬，截至 2026-10-25 尚未公布）")
inp("q1_rev", "1Q26 營收（實績）", "NT$B", None, 19.313, None, "Verified", SAQ + "：1Q26 營收 19,313、毛利 6,722、營業利益 5,185、淨利 3,484 百萬元、稅前淨利率 26.34%", "優分析 2026-08-11：1H26 營收 435.92 億、2Q26 242.79 億（相減得 1Q26 193.13 億）")
inp("q3_rev", "3Q26 營收（7–9 月月營收合計，實績）", "NT$B", None, "=10.015+10.384+9.320", None, "Verified",
    TN9 + "：9 月 93.20 億、前 9 月 733.79 億；7 月 100.15 億、8 月 103.84 億（優分析）", "理財周刊 2026-10-14：3Q26 營收 297.18 億（同值）")
inp("q4g", "4Q26 營收季增", "%", 0.00, 0.08, 0.15, "Assumed", "新 AI ASIC 產品線 4Q 開始出貨、楊梅租賃廠 3Q26 開出（優分析 2026-09-09、2026-03-10）", "9 月營收月減 10.2%（technews）；共識 FY2026 營收 1,063 億（stockanalysis 8 家）隱含 4Q 約 327 億（只作 Recon）")
inp("q_gm", "3Q、4Q26 毛利率", "%", 0.34, 0.35, 0.365, "Assumed", "1Q26 34.81%、2Q26 34.65%、1H26 34.72%（優分析 2026-08-11）", "分析師：3Q26 起逐季上升（Interested-party）")
inp("q_opx", "3Q、4Q26 營業費用 ÷ 營收", "%", 0.075, 0.080, 0.085, "Assumed", "1H26：毛利 151.35 − 營業利益 116.02 億＝35.33 億、8.1%（Q 頁）", "2Q26 營業利益率 26.43%（理財周刊 2026-10-14）")
inp("q_nop", "3Q、4Q26 營業外淨額（每季）", "NT$B", -0.2, 0.0, 0.3, "Assumed", "1Q26 稅前 − 營業利益＝−0.10 億（stockanalysis 稅前淨利率推算）；2Q26 由淨利 ÷ q_nir 推算約 +0.57 億（Derived）", "可轉債利息、匯兌損益不外推")
inp("q_nir", "淨利 ÷ 稅前（1Q26 實績；3Q、4Q26 沿用）", "%", 0.65, None, 0.75, "Derived", "1Q26：淨利 34.84 ÷（營收 193.13 × 稅前淨利率 26.34%）＝0.685（stockanalysis）", "FY2025：96.07 ÷ 141.73＝0.678（stockanalysis）；含中國子公司盈餘匯回與未分配盈餘稅（Assumed）")
inp("sh_1h", "1H26 加權股數（十億股）", "十億股", None, "=8.267/16.14", None, "Derived", "1H26 淨利 82.67 億 ÷ EPS 16.14（優分析 2026-08-11）", "理財周刊引 1H26 EPS 16.17（差異未說明）")
inp("sh_h2", "下半年股數（十億股；含 Amazon 私募 185.6 萬股）", "十億股", None, "=0.51190+0.001856", None, "Derived", "stockanalysis 流通股數 511.90 百萬股（2026-10-22）", "Amazon 私募 1,856,308 股、每股 856 元（理財周刊 2026-10-14）")

sec("G. 營業費用、業外、稅與股數（FY2027 起）")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.40, 0.60, 1.00, "Assumed", "營業費用率 FY2025 9.4% → 1H26 8.1%（營收 +68%）（stockanalysis、Q 頁）", "外資：成長三支柱含嚴格費用控制（technews 2026-05-13，Interested-party）")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.01, 0.0, 0.01, "Assumed", "無", "Sens 用 ±1 個百分點")
inp("nop_f", "營業外淨額（每年）", "NT$B", -1.0, 0.0, 1.0, "Assumed", "FY2025 +1.30 億（stockanalysis 推算）", "2025 年發行逾 90 億元無擔保可轉債（鉅亨 2025-09-26）")
inp("nir_f", "淨利 ÷ 稅前（FY2027 起）", "%", 0.65, 0.685, 0.75, "Derived", "同 q_nir（1Q26 0.685；FY2025 0.678）", "—")
inp("sh_g", "股數年變動（可轉債轉換、員工酬勞）", "%", 0.0, 0.01, 0.03, "Assumed", "2025 年逾 90 億元可轉債（鉅亨 2025-09-26）；轉換價：未找到", "理財周刊 2026-10-14 提示可轉債稀釋風險")
inp("mf", "AI 部署係數（Sens 用；FY2027 起）", "x", 0.8, 1.0, 1.2, "Assumed", "網通模型 2027 年 AI IT 資本支出約 2 倍", "±20%")

sec("H. 共識、券商與上游（只作 Recon 對帳，不作校準）")
for y, rv, ep in ((26, 106.371, 41.22), (27, 159.967, 62.80), (28, 214.339, 95.35)):
    inp(f"cons_rev{y}", f"共識 FY20{y} 營收（MarketScreener）", "NT$B", None, rv, None, "Interested-party", MSV + "（2026-10-25 擷取，參考價 1,115）", "stockanalysis FY2026 1,063 億（8 家）" if y == 26 else "—")
    inp(f"cons_eps{y}", f"共識 FY20{y} EPS（MarketScreener）", "NT$", None, ep, None, "Interested-party", MSV, "stockanalysis FY2026 41.45（8 家，區間 37.24–46.54）" if y == 26 else "—")
for y, ep in ((26, 41.57), (27, 69.10), (28, 95.07)):
    inp(f"fs_eps{y}", f"FactSet 中位數 FY20{y} EPS（2026-09）", "NT$", None, ep, None, "Interested-party", MW10 + "（理財周刊轉引）", "理財周刊 2026-09-23：2027 年 67.32" if y == 27 else "—")
inp("fb_eps27", "外資 FY2027 EPS（2026-05-13，目標價 1,310）", "NT$", None, 65.39, None, "Interested-party", TNFB, "券商名未載明")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"emc_ccl{y}", f"台光電模型：每 GW 高速 CCL 金額 20{y}（合計）", "US$M/GW", None, EMC["ccl"][0][j], None, "Derived", EMC["ccl"][1], "唯讀台光電成品（CCL 口徑；只在 Recon 對照）" if y == 26 else "同上")
    inp(f"vg_pgw{y}", f"勝宏模型：每 GW AI 高速 PCB 金額 20{y}（合計）", "US$M/GW", None, VG["pgw"][0][j], None, "Derived", VG["pgw"][1], "唯讀勝宏成品（Recon 用）" if y == 26 else "同上")
for j, y in enumerate((26, 27, 28)):
    inp(f"emc_pg{y}", f"台光電模型：每 GW 台光電 AI 相關營收 20{y}", "US$M/GW", None, EMC["emcpg"][0][j], None, "Derived", EMC["emcpg"][1], "唯讀台光電成品")
    inp(f"vg_pge{y}", f"勝宏模型：每 GW 勝宏 AI 相關營收 20{y}", "US$M/GW", None, VG["pge"][0][j], None, "Derived", VG["pge"][1], "唯讀勝宏成品")
inp("vg_sh26", "勝宏模型：勝宏有效份額 2026", "%", None, VG["sh"][0][0], None, "Derived", VG["sh"][1], "唯讀勝宏成品")

sec("I. 價格與估值")
inp("px", "股價（2026-10-22 收盤，臺灣證交所）", "NT$", None, 1115, None, "Interested-party", "stockanalysis（網址見 Sources 頁）：10-22 收盤 1,115（−6.69%）、前收 1,195", "nstock、聚財網 10-22：1,115；10-23 收盤：未找到")
inp("pe_ttm", "近 12 月 P/E（stockanalysis，10-22）", "x", None, 39.79, None, "Interested-party", SA, "近 12 月 EPS 28.02")
inp("fpe", "前瞻 P/E（stockanalysis，10-22）", "x", None, 19.57, None, "Interested-party", SA, "口徑：頁面未明示（約未來 12 月）")
inp("cons_feps", "共識前瞻 12 月 EPS（現價 ÷ 前瞻 P/E）", "NT$", None, "=px/fpe", None, "Derived", "stockanalysis 前瞻 P/E 19.57（2026-10-22）", "共識前瞻 EPS 的直接數字：未找到")
inp("cons_tp", "市場共識目標價（賣方平均）", "NT$", None, 1554, None, "Interested-party", SAF + "：9 位分析師（S&P Global）平均 1,554、中位數 1,610、區間 850–1,955（頁面時間 2026-09-04）", "外資 1,310（technews 2026-05-13）")
inp("nim_avg", "十年平均淨利率替代值（2016–2020 營收與淨利未找到；依工作單以情境最後一年 × 0.8）", "%", None, 0, None, "Assumed",
    "工作單 2026-10-25：找不到十年平均時用最後一年 × 0.8；本值＝情境 C FY2029 淨利率 × 0.8（公式於建置時連結）", "參考：FY2021–25 五年平均 13.4%（stockanalysis：11.0%、13.9%、11.7%、14.4%、16.0%）")

sec("J. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_r", "折現率／要求報酬", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%，各市場同一值）", "DCF 折現；情境價值推到 12 個月後")
inp("tp_g", "長期成長率", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "DCF 永續成長與 FY2032 收斂目標（情境 B）")
inp("tp_roe", "ROE（近三年平均，FY2023–25）", "%", None, "=AVERAGE(0.2265,0.2945,0.3507)", None, "Derived", "stockanalysis：FY2023 22.65%、FY2024 29.45%、FY2025 35.07%", "stockanalysis 目前 ROE 45.95%（2026-09）")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（2026-10-25 11:53 定 DCF／Comparable 各半）", "—")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（情境 B FY2029 與十年平均替代值的中點）", "%", None, 0, None, "Assumed", "PL 頁情境 B FY2029、In 頁 nim_avg（十年平均未找到）", "—")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（＝十年平均替代值 nim_avg）", "%", None, "=nim_avg", None, "Assumed", "In 頁 nim_avg（十年平均未找到，用情境 C FY2029 × 0.8）", "—")
TP_KEYS += ["tp_r", "tp_g", "tp_roe", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]

r = 4; INROW = {}
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, s1, s2 = srow
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    if key == "q_nir": base = "=3.484/(19.313*0.2634)"
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
for nm, base, suf in (("RLO", "rel", "_lo"), ("RHI", "rel", "_hi"), ("PPLO", "pp", "_lo"), ("PPHI", "pp", "_hi"), ("CAPLO", "cx", "_lo"), ("CAPHI", "cx", "_hi"),
                      ("GULO", "gmu", "_lo"), ("GUHI", "gmu", "_hi")):
    alias(nm, base, suf)
for nm, suf in (("GNLO", "_lo"), ("GNHI", "_hi")):
    for y in (27, 28, 29):
        wb.defined_names[f"{nm}{y}"] = DefinedName(f"{nm}{y}", attr_text=wb.defined_names[f"g_nai{y}{suf}"].attr_text)

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億新台幣；EPS 為新台幣；股數為十億股）。藍字＝輸入；黑字＝公式。實績取自金像電財報與月營收公告（經 stockanalysis、優分析、technews 轉述）"
Q["A2"] = ("FY2026 錨點＝1H26 財報＋3Q26 月營收實績（7–9 月）＋假設（3Q 毛利率、營業費用率、營業外；4Q 營收季增）。1H26 毛利與營業利益＝1Q26（stockanalysis）＋2Q26（營收 × 毛利率 34.65%、營業利益率 26.43%）；"
           "1H26 稅前＝1Q26 稅前＋2Q26 淨利 ÷ q_nir（Derived）。3Q26 財報與法說預計 2026-11 上旬（stockanalysis 估 11-11；截至 2026-10-25 尚未公布），公布後更新")
hdr(Q, 3, ["項目", "FY2025A", "1H26（實績）", "3Q26（營收實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1", "來源2／說明"])
S1 = "FY2025：stockanalysis 年度損益（營收 600.04 億、毛利 196.83 億、營業利益 140.43 億、淨利 96.07 億）；1H26：優分析 2026-08-11（營收 435.92 億、毛利率 34.72%、淨利 82.67 億、EPS 16.14）"
S2 = "FY2025：優分析 2026-03-10（營收 600.04 億、毛利率 32.8%、淨利 96.07 億、EPS 19.47）；經濟日報 2026-03-11（淨利 96.06 億）"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2=S2, ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"F{r}"] = ann or f"=SUM(C{r}:E{r})"; Q[f"F{r}"].number_format = fmt
    Q.cell(r, 7, tag); Q.cell(r, 8, s1); Q.cell(r, 9, s2)
qrow(4, "營收", {"B": 60.004, "C": 43.592, "D": "=q3_rev", "E": "=D4*(1+q4g)"}, tag="Verified／Assumed（4Q 季增）", s2="3Q26＝7–9 月月營收合計（technews、優分析）；月營收與財報營收口徑可能略有差異")
qrow(5, "毛利", {"B": 19.683, "C": "=6.722+24.279*0.3465", "D": "=D4*q_gm", "E": "=E4*q_gm"}, tag="Verified／Assumed", s2="2Q26 毛利率 34.65%（優分析）；1Q26 毛利 67.22 億（stockanalysis）")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="1H26 34.72%（優分析）；FY2025 32.8%")
qrow(7, "營業費用", {"B": "=B5-B8", "C": "=C5-C8", "D": "=D4*q_opx", "E": "=E4*q_opx"}, tag="Derived／Assumed", s2="毛利 − 營業利益")
qrow(8, "營業利益（毛利 − 營業費用）", {"B": 14.043, "C": "=5.185+24.279*0.2643", "D": "=D5-D7", "E": "=E5-E7"}, tag="Interested-party（stockanalysis）／Derived",
     s2="1Q26 營業利益 51.85 億（stockanalysis）；2Q26 營業利益率 26.43%（理財周刊 2026-10-14）")
qrow(9, "營業利益率", {c: f"={c}8/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(10, "營業外淨額", {"B": "=B11-B8", "C": "=C11-C8", "D": "=q_nop", "E": "=q_nop"}, tag="Derived／Assumed", s1="稅前 − 營業利益", s2="—")
qrow(11, "稅前利益", {"B": "=60.004*0.2362", "C": "=19.313*0.2634+4.783/q_nir", "D": "=D8+D10", "E": "=E8+E10"}, tag="Derived",
     s1="FY2025＝營收 × 稅前淨利率 23.62%（stockanalysis）；1H26＝1Q26 營收 × 26.34%＋2Q26 淨利 47.83 億 ÷ q_nir", s2="2Q26 稅前：未找到")
Q["F10"] = "=SUM(C10:E10)"; Q["F11"] = "=SUM(C11:E11)"
qrow(12, "歸屬母公司淨利", {"B": 9.607, "C": 8.267, "D": "=D11*q_nir", "E": "=E11*q_nir"}, tag="Verified／Derived（3Q、4Q）", s2="FY2025 96.07 億（年增 71%）；1H26 82.67 億（2Q26 47.83 億）")
qrow(13, "淨利 ÷ 營業利益", {c: f"={c}12/{c}8" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(14, "股數（十億股；加權）", {"B": "=9.607/19.47", "C": "=sh_1h", "D": "=sh_h2", "E": "=sh_h2"}, ann="=(sh_1h+sh_h2)/2", tag="Derived", s2="淨利 ÷ EPS；FY2026＝上半年與下半年平均")
qrow(15, "EPS", {c: f"={c}12/{c}14" for c in "BCDE"}, fmt="0.00", ann="=F12/F14", tag="Derived", s1="淨利 ÷ 股數", s2="—")
qrow(16, "  公司揭露 EPS", {"B": 19.47, "C": 16.14}, fmt="0.00", ann="", tag="Verified", s2="FY2025 19.47（stockanalysis 列 19.00，口徑未載明）；1H26 16.14")
qrow(17, "AI 相關營收（FY2025＝營收 × AI 占比；FY2026＝營收 − 非 AI）", {"B": "=B4*ai_25"}, ann="=F4-F18", tag="Interested-party／Derived", s1="AI 占比（In 頁 D 節）", s2="各季 AI 營收金額未揭露：未找到")
qrow(18, "非 AI 營收（FY2026＝FY2025 ×（1＋g_nai26））", {"B": "=B4-B17"}, ann="=B18*(1+g_nai26)", tag="Derived／Assumed", s1="—", s2="一般伺服器、網通、筆電、其他")
Q.cell(23, 1, "FY2026 錨點的檢查（參考；校準未使用券商數字）").font = BOLD
Q.cell(24, 1, "FY2026 營收年增"); Q["B24"] = "=F4/B4-1"
Q.cell(25, 1, "FY2026 AI 相關 ÷ 營收"); Q["B25"] = "=F17/F4"
Q.cell(26, 1, "FY2026 營業費用 ÷ 營收"); Q["B26"] = "=F7/F4"
Q.cell(27, 1, "FY2026 營收 ÷ 共識（MarketScreener）− 1"); Q["B27"] = "=F4/cons_rev26-1"
Q.cell(28, 1, "FY2026 EPS ÷ 共識（MarketScreener）− 1"); Q["B28"] = "=F15/cons_eps26-1"
Q.cell(29, 1, "4Q26 年化營收 ÷ FY2026 營收"); Q["B29"] = "=E4*4/F4"
for rr in range(24, 29): Q[f"B{rr}"].number_format = "0.0%"
Q["B29"].number_format = "0.00"
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 7), ("oi", 8), ("ni", 12), ("sh", 14), ("eps", 15), ("ai", 17), ("nai", 18)):
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
# 情境與行業一致：A＝份額、價格、AI 毛利率改善取高值；B＝基準；C＝取低值（產能、非 AI 各情境相同）
SCN_P = {"A": dict(rel="RHI", pp="PPHI", gmu="GUHI"), "B": dict(), "C": dict(rel="RLO", pp="PPLO", gmu="GULO")}
DEFP = dict(scn="B", pgwf="pgwf", wmid="w_mid", rel="rel", pp="pp", cx="cx", mf="1", gn="g_nai", d="d_ai", gmu="gmu", opxa="0", phi="phi", nir="nir_f")
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
    sec_("A. AI 部署與每 GW AI 高速 PCB 金額（需求面；D 欄＝FY2026 校準年；每 GW 金額取自勝宏模型）")
    row("itc", "AI IT 資本支出 × 情境顆數係數", "US$B", {c_(y): f"=itc_{y % 100}*{umx(y, s)}" + ("" if y == 2026 else f"*{P['mf']}") for y in BU})
    row("itpg", "每 GW IT 資本支出", "US$B", {"D": "=itpg_gb*(1-vr26)+itpg_vr*vr26", **{c_(y): "=itpg_vr" for y in BU[1:]}})
    row("gw", "部署 GW（IT 電力）", "GW", {c_(y): f"={c_(y)}{R['itc']}/{c_(y)}{R['itpg']}" for y in BU}, fmt="0.0")
    for k, lab, nm, w in (("vcmp", "每 GW PCB：加速器卡／運算板", "cmp", ""), ("vsw", "每 GW PCB：交換器板", "sw", ""),
                          ("vmid", "每 GW PCB：中板／背板／其他 × 金像電參與係數", "mid", f"*{P['wmid']}")):
        row(k, lab, "US$M/GW", {"D": f"=vg_{nm}26{w}", **{c_(y): f"=(vg_{nm}26+(vg_{nm}{y % 100}-vg_{nm}26)*{P['pgwf']}){w}" for y in BU[1:]}}, fmt="#,##0")
    row("pgw", "每 GW 金像電可參與 AI 高速 PCB 金額（三項合計）", "US$M/GW", {c_(y): f"={c_(y)}{R['vcmp']}+{c_(y)}{R['vsw']}+{c_(y)}{R['vmid']}" for y in BU}, fmt="#,##0")
    row("mku", "可參與 AI 高速 PCB 市場（美元；不含價格指數）", "US$B", {c_(y): f"={c_(y)}{R['gw']}*{c_(y)}{R['pgw']}/1000" for y in BU})
    row("mkn", "可參與 AI 高速 PCB 市場（新台幣）", "NT$B", {c_(y): f"={c_(y)}{R['mku']}*fx" for y in BU}, fmt="#,##0.0")
    row("pp", "PCB 價格指數（FY2026＝1）", "x", {"D": "=1", **{c_(y): f"={P['pp']}_{y % 100}" for y in BU[1:]}}, fmt="0.00")
    row("sh0", "金像電有效份額（FY2026 校準）＝ AI 相關營收 ÷ 可參與市場", "%", {"D": f"=a26_ai/D{R['mkn']}"}, fmt="0.0%",
        note="Derived：分子為 FY2026 錨點營收 − 非 AI；分母為本模型的可參與市場（含 ASIC）")
    row("rel", "份額相對值（÷ FY2026）", "x", {"D": "=1", **{c_(y): f"={P['rel']}_{y % 100}" for y in BU[1:]}}, fmt="0.00")
    row("sh", "金像電有效份額", "%", {c_(y): f"=$D${R['sh0']}*{c_(y)}{R['rel']}" for y in BU}, fmt="0.0%")
    row("aid", "AI 相關營收（需求面）＝ 市場 × 份額 × 價格指數", "NT$B", {c_(y): f"={c_(y)}{R['mkn']}*{c_(y)}{R['sh']}*{c_(y)}{R['pp']}" for y in BU}, fmt="#,##0.0")
    sec_("B. 產能上限（全產全銷）與營收")
    row("cx", "平均產值產能指數（FY2026 營收＝1）", "x", {"D": "=1", **{c_(y): f"={P['cx']}_{y % 100}" for y in BU[1:]}}, fmt="0.00",
        note="FY2026 校準：公司稱產能吃緊、產線滿載，FY2026 營收視為產能（Derived）")
    row("rcap", "產能上限營收 ＝ FY2026 營收 × 產能指數 × 價格指數", "NT$B", {c_(y): f"=a26_rev*{c_(y)}{R['cx']}*{c_(y)}{R['pp']}" for y in BU}, fmt="#,##0.0")
    row("nai", "非 AI 營收", "NT$B", {"C": "=a25_nai", "D": "=a26_nai", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gn']}{y % 100})" for y in BU[1:]}}, fmt="#,##0.0")
    row("ai", "AI 相關營收 ＝ MIN（需求面、產能上限 − 非 AI）", "NT$B", {"C": "=a25_ai", "D": f"=D{R['aid']}",
        **{c_(y): f"=MIN({c_(y)}{R['aid']},{c_(y)}{R['rcap']}-{c_(y)}{R['nai']})" for y in BU[1:]}}, fmt="#,##0.0", note="產能不足時，非 AI 照成長率出貨、AI 取剩餘產能（Assumed）")
    row("bind", "產能是否受限", "", {c_(y): f'=IF({c_(y)}{R["aid"]}>{c_(y)}{R["rcap"]}-{c_(y)}{R["nai"]}+0.0001,"受限","未受限")' for y in BU[1:]}, fmt="General")
    row("rev", "營收", "NT$B", {c_(y): f"={c_(y)}{R['ai']}+{c_(y)}{R['nai']}" for y in FYS}, fmt="#,##0.0")
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("aish", "AI 相關 ÷ 營收", "%", {c_(y): f"={c_(y)}{R['ai']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("util", "需求面 ÷ 產能上限（AI 需求面＋非 AI）", "%", {c_(y): f"=({c_(y)}{R['aid']}+{c_(y)}{R['nai']})/{c_(y)}{R['rcap']}" for y in BU}, fmt="0.0%",
        note="大於 100%＝需求超過產能（營收由產能決定）")
    row("pgwe", "每 GW 金像電 AI 相關營收", "US$M/GW", {c_(y): f"={c_(y)}{R['ai']}/fx/{c_(y)}{R['gw']}*1000" for y in BU}, fmt="#,##0")
    sec_("C. 損益")
    row("gmn", "其他營收毛利率 ＝ 錨點毛利率 − 差距 × AI 占比（FY2026 校準）", "%", {"D": f"=a26_gp/D{R['rev']}-{P['d']}*D{R['aish']}"}, fmt="0.0%",
        note="Derived：讓 FY2026 毛利率等於錨點；AI 相關毛利率＝其他＋差距")
    row("gma", "AI 相關毛利率 ＝ 1 −（1 − 校準值）÷ 價格指數 ＋ 額外改善", "%", {"D": f"=D{R['gmn']}+{P['d']}", **{c_(y): f"=1-(1-$D${R['gmn'] + 1})/{c_(y)}{R['pp']}+{P['gmu']}_{y % 100}" for y in BU[1:]}}, fmt="0.0%",
        note="Derived：同規格成本不變，價格升降全部落在毛利；額外改善為新廠爬坡完成與組合（Assumed）")
    row("gp", "毛利", "NT$B", {"C": "=a25_gp", "D": "=a26_gp", **{c_(y): f"={c_(y)}{R['ai']}*{c_(y)}{R['gma']}+{c_(y)}{R['nai']}*$D${R['gmn']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("gm", "毛利率", "%", {c_(y): f"={c_(y)}{R['gp']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("opb", "營業費用（基準路徑）", "NT$B", {"C": "=a25_opx", "D": "=a26_opx", **{c_(y): f"={c_(y - 1)}{r}*(1+MAX(0,{P['phi']}*{c_(y)}{R['g']}))" for y in BU[1:]}}, fmt="#,##0.0")
    row("opx", "營業費用", "NT$B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", **{c_(y): f"={c_(y)}{R['opb']}+{P['opxa']}*{c_(y)}{R['rev']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("oi", "營業利益", "NT$B", {c_(y): f"={c_(y)}{R['gp']}-{c_(y)}{R['opx']}" for y in FYS}, fmt="#,##0.0")
    row("oim", "營業利益率", "%", {c_(y): f"={c_(y)}{R['oi']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("nop", "營業外淨額", "NT$B", {c_(y): "=nop_f" for y in BU[1:]}, fmt="#,##0.0", note="FY2025、FY2026 含在淨利錨點內")
    row("ni", "歸屬母公司淨利", "NT$B", {"C": "=a25_ni", "D": "=a26_ni", **{c_(y): f"=({c_(y)}{R['oi']}+{c_(y)}{R['nop']})*{P['nir']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("nioi", "淨利 ÷ 營業利益", "%", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['oi']}" for y in FYS}, fmt="0.0%")
    row("sh_", "股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", "E": "=sh_h2*(1+sh_g)", **{c_(y): f"={c_(y - 1)}{r}*(1+sh_g)" for y in (2028, 2029)}}, fmt="0.0000",
        note="FY2027 起＝期末股數 ×（1＋sh_g）（FY2026 為加權）")
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
PL["A1"] = "PL — 金像電損益（十億新台幣；EPS 新台幣；歸屬母公司）。FY2025A 實績、FY2026 錨點（1H26 財報＋3Q26 月營收＋假設），FY2027 起為模型。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要。顆數係數從 CY2028 開始；份額、價格、AI 毛利率改善從 FY2027 起依情境取高／基準／低值"
PL["A2"] = "灰底＝FY2030–FY2032 收斂法（上游只到 2029）：營收年增線性收斂到 fade_g、淨利率線性收斂到 fade_m（In 頁 J 節），非驅動因子推導"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A（順利）：顆數基準（tsmc 係數 A＝1.0）；份額、價格、AI 毛利率改善取高值",
         "B": "情境 B（平平）：tsmc 係數 B（CY2028 ×0.8、CY2029 ×0.85）；公司特有變數取基準",
         "C": "情境 C（不如預期）：tsmc 係數 C（CY2028–29 ×0.6）；份額、價格、AI 毛利率改善取低值"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s, **SCN_P[s])
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目"] + YH[1:])
sr = 4; SUMR = {}
SUMK = (("rev", "營收", "NT$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("ai", "  AI 相關營收（AI 伺服器主板、800G 交換器）", "NT$B", "#,##0.0"),
        ("nai", "  非 AI 營收", "NT$B", "#,##0.0"), ("util", "  需求面 ÷ 產能上限", "%", "0.0%"), ("gm", "毛利率", "%", "0.0%"), ("oim", "營業利益率", "%", "0.0%"),
        ("ni", "歸屬母公司淨利", "NT$B", "#,##0.0"), ("nim", "  淨利率", "%", "0.0%"), ("eps", "EPS", "NT$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%"))
LONGK = ("rev", "g", "ni", "nim", "eps", "epsg")
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in SUMK:
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS + (FADE if key in LONGK else ()):
            if key in ("g", "epsg", "util") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
            if y in FADE: PL[f"{FC[y]}{sr}"].fill = FADEFILL
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
IN.cell(INROW["fade_mA"], 5, f"=PL!G{SUMR[('A', 'nim')]}")
IN.cell(INROW["fade_mB"], 5, f"=(PL!G{SUMR[('B', 'nim')]}+nim_avg)/2")
IN.cell(INROW["nim_avg"], 5, f"=PL!G{SUMR[('C', 'nim')]}*0.8")
for k in ("fade_mA", "fade_mB", "nim_avg"): IN.cell(INROW[k], 5).font = Font(color="000000")
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型 − 錨點").font = BOLD; sr += 1
CHK = []
for lab, f in (("  三情境：營收差（NT$B）", "+".join(f"ABS(D{PB[s]['rev']}-a26_rev)" for s in SC)),
               ("  情境 B：AI 相關差（NT$B）", f"D{PB['B']['ai']}-a26_ai"),
               ("  情境 B：毛利差（NT$B）", f"D{PB['B']['gma']}*D{PB['B']['ai']}+D{PB['B']['gmn']}*D{PB['B']['nai']}-a26_gp"),
               ("  情境 B：產能上限 − 營收（NT$B；FY2026 校準恆等）", f"D{PB['B']['rcap']}-D{PB['B']['rev']}")):
    PL.cell(sr, 1, lab); PL[f"D{sr}"] = "=" + f; PL[f"D{sr}"].number_format = "0.0000"; CHK.append(sr); sr += 1
assert sr < 60, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 10; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 13

# ---------------- Units：每 GW PCB 金額（取自勝宏模型）與金像電份額 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 每 GW AI 高速 PCB 金額（口徑與勝宏模型一致，唯讀 victory_giant/dist；本專案權威在勝宏模型）、金像電參與係數、份額與產能。部署 GW＝AI IT 資本支出（markets/network）÷ 每 GW IT 資本支出（Tokenomics，經橋接表）"
B_ = PB["B"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("gw", "部署 GW（情境 B）", "GW", lambda y: f"=PL!{FC[y]}{B_['gw']}", "0.0", "AI IT 資本支出 × 顆數係數 B ÷ 每 GW IT 資本支出"),
        ("vgp", "勝宏模型：每 GW AI 高速 PCB 金額（合計）", "US$M/GW", lambda y: f"=vg_pgw{y % 100}", "#,##0", "唯讀 victory_giant/dist（口徑來源）"),
        ("vcmp", "每 GW PCB：加速器卡／運算板", "US$M/GW", lambda y: f"=PL!{FC[y]}{B_['vcmp']}", "#,##0", "勝宏模型值；成長倍率 pgwf＝1 時相同"),
        ("vsw", "每 GW PCB：交換器板", "US$M/GW", lambda y: f"=PL!{FC[y]}{B_['vsw']}", "#,##0", "同上"),
        ("vmid", "每 GW PCB：中板／背板／其他 × 參與係數", "US$M/GW", lambda y: f"=PL!{FC[y]}{B_['vmid']}", "#,##0", "參與係數 w_mid（Assumed）"),
        ("pgw", "每 GW 金像電可參與 AI 高速 PCB 金額（不含價格指數）", "US$M/GW", lambda y: f"=PL!{FC[y]}{B_['pgw']}", "#,##0", ""),
        ("prk", "每架等值可參與金額（以 VR200 每 GW 架數換算）", "US$K/架", lambda y: f"=PL!{FC[y]}{B_['pgw']}*1000/racks_vr", "#,##0.0", ""),
        ("pcr", "每 GW 可參與 PCB 金額 ÷ 台光電模型每 GW CCL 金額", "x", lambda y: f"=PL!{FC[y]}{B_['pgw']}/emc_ccl{y % 100}", "0.00", "見 Recon B 節"),
        ("mku", "可參與 AI 高速 PCB 市場（含價格指數）", "US$B", lambda y: f"=PL!{FC[y]}{B_['mku']}*PL!{FC[y]}{B_['pp']}", "#,##0.0", "含 ASIC 部署（假設每 GW 金額相同）"),
        ("sh", "金像電有效份額", "%", lambda y: f"=PL!{FC[y]}{B_['sh']}", "0.0%", "FY2026 以錨點 AI 相關營收校準"),
        ("aid", "金像電 AI 相關營收（需求面）", "US$B", lambda y: f"=PL!{FC[y]}{B_['aid']}/fx", "#,##0.00", ""),
        ("ai", "金像電 AI 相關營收（實際＝MIN 需求面、產能）", "US$B", lambda y: f"=PL!{FC[y]}{B_['ai']}/fx", "#,##0.00", ""),
        ("pgwe", "每 GW 金像電 AI 相關營收", "US$M/GW", lambda y: f"=PL!{FC[y]}{B_['pgwe']}", "#,##0", "對照勝宏、台光電（Recon）"),
        ("cx", "平均產值產能指數（FY2026＝1）", "x", lambda y: f"=PL!{FC[y]}{B_['cx']}", "0.00", "公司擴產計畫（Assumed）"),
        ("rcap", "產能上限營收（含價格指數）", "NT$B", lambda y: f"=PL!{FC[y]}{B_['rcap']}", "#,##0.0", "FY2026 校準"),
        ("util", "需求面 ÷ 產能上限", "%", lambda y: f"=PL!{FC[y]}{B_['util']}", "0.0%", "大於 100%＝產能受限"),
        ("oth", "其他廠商（健鼎、TTM、勝宏、沪电、華通等）＝ 可參與市場 − 金像電", "US$B", lambda y: f"=PL!{FC[y]}{B_['mku']}*PL!{FC[y]}{B_['pp']}-PL!{FC[y]}{B_['ai']}/fx", "#,##0.0", "產能受限時，未滿足需求計入其他廠商")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); UR[key] = ur; ur += 1
ur += 1
UN.cell(ur, 1, "讀法：每 GW 金額的絕對水準只影響 FY2026 校準出的份額（金額高 → 份額低），不影響營收；影響營收的是金額成長（勝宏模型世代指數 × pgwf）、參與係數與份額相對值。產能受限的年度，營收由產值產能指數決定；擴產趕上需求後，份額與每 GW 金額重新成為主要變數。").font = BOLD
UNOTE = ur
UN.column_dimensions["A"].width = 66; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 14
UN.column_dimensions["G"].width = 60

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，年底落後 P/E）只作參考。本版未做先驗，無 Andy 合理 P/E 對照"
V["A2"] = "循環位置：金像電不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照勝宏、台光電；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
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
V[f"E{O + 3}"] = "FY2026 三情境相同（錨點）"; V[f"E{O + 4}"] = "FY2027 三情境差異只來自份額、價格、AI 毛利率改善（顆數係數從 CY2028 開始）"
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
# 十年 P/E：自算＝digrin 12 月未調整收盤 ÷ 年度 EPS（2016–2020 聚財網；2021–2025 stockanalysis／公司）；MarketScreener 年度 P/E（2021–2025）只作核對
PX_DEC = [10.40, 10.60, 13.10, 16.65, 50.50, 76.0, 86.80, 218.0, 241.50, 687.0]
EPS_Y = [0.26, -0.50, 0.42, 0.24, 3.82, 5.98, 8.78, 7.25, 11.54, 19.47]
PE_SELF = [None if e <= 0 else f"={p}/{e}" for p, e in zip(PX_DEC, EPS_Y)]
PE_MS = [None] * 5 + [14.1, 9.89, 30.1, 20.9, 35.3]
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "自算（digrin 12 月未調整收盤 ÷ 年度 EPS）", "vals": PE_SELF, "tag": "Derived（EPS：2016–2020 聚財網、2021–2025 stockanalysis／公司；2017 年 EPS −0.50 為虧損，P/E 無意義，留空；9 年）"},
    {"label": "MarketScreener（年度 P/E）", "vals": PE_MS, "tag": "Interested-party（資料商；只有 2021–2025，只作核對，不採用）", "use": False}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位。採用自算（9 年；2017 年虧損留空）；MarketScreener 只有 2021–2025，與自算同量級，只作核對；HiStock 12 月 P/E 2021–2025：16.7、9.5、31.6、21.6、43.2。2016–2019 年 EPS 低於 0.5 元，P/E 被低獲利放大（2019 年 69x），十年分布偏高。")
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
PEERS = [("健鼎（Tripod Technology）", "3044.TW", 13.96, "2026-09-10"), ("華通（Compeq）", "2313.TW", 20.76, "2026-08-20"),
         ("臻鼎-KY（Zhen Ding）", "4958.TW", 25.13, "2026-08-31"), ("博智（Allied Circuit）", "8155.TWO", 18.49, "2026-08-14"),
         ("TTM Technologies", "TTMI", 10.31, "2026-10-23")]
for i, (nm, tk, pe, dt) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左，臺股頁面為快取日期）"); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
V.cell(33, 1, "並列參考（A 股，不進目標價）：勝宏科技 300476.SZ 前瞻 P/E 14.8x（10-23）、沪电股份 002463.SZ 20.7x（10-22）；stockanalysis，Interested-party")
tset(34, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", "臺美伺服器、網通 PCB 同業：健鼎、華通、臻鼎-KY、博智、TTM（臺股媒體與券商常用比較組）；定穎投控、瀚宇博前瞻 P/E 未找到")
tset(35, "feps", "12 個月後的前瞻 EPS（情境加權）", [f"=tp_wA*(2/12*B{EPSR[2027]}+10/12*B{EPSR[2028]})+tp_wB*(2/12*C{EPSR[2027]}+10/12*C{EPSR[2028]})+tp_wC*(2/12*D{EPSR[2027]}+10/12*D{EPSR[2028]})", None, None], "0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(36, "comptp", "Comparable 目標價", ["=B34*B35", None, None], "0.00", "同業中位數 × 前瞻 EPS")
V.cell(38, 1, "0.3 結論").font = BOLD
hdr(V, 39, ["項目", "數值", "", "", "說明"])
tset(40, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B24+(1-tp_w_dcf)*B36", None, None], "0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁")
tset(41, "comp2", "Comparable 目標價（市場主流同業法）", ["=B36", None, None], "0.00")
tset(42, "cons", "市場共識目標價", ["=cons_tp", None, None], "0.00", "stockanalysis（S&P Global）：9 位分析師平均 1,554、中位數 1,610（頁面時間 2026-09-04）")
tset(43, "upw", "上檔空間：加權目標價", ["=B40/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-22 收盤）")
tset(44, "upc", "上檔空間：Comparable 目標價", ["=B41/px-1", None, None], "0.0%")
tset(45, "geps", "與市場共識目標價的差距：EPS 因素", ["=B35/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝現價 ÷ 前瞻 P/E（In 頁 cons_feps；期間為現在起算 12 個月，與本模型不同）")
tset(46, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B34-1", None, None], "0.0%")
V["E46"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B34,"0.0")&"x"'
tset(47, "band", "價格帶（DCF 不如預期～順利）", ['=TEXT(D23,"0.00")&"～"&TEXT(B23,"0.00")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(48, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=IF(px/{c}{EPSR[2028]}<MIN({PE10}),0,IF(px/{c}{EPSR[2028]}>MAX({PE10}),1,PERCENTRANK({PE10},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 6 節自算）的分位；只作檢查，不進目標價")
V.cell(49, 1, "目標價方法")
V["B49"] = ("DCF：股東現金流＝歸母 EPS ×（1 − g／ROE），r＝10%、g＝4%、ROE 為近三年平均約 29%（再投資率約 14%），FY2027–32 明確期＋永續；"
            "金像電 2026 年資本支出約 190 億元（約營收的 18%，折舊約 14 億），本模型未建資本支出與折舊，未改用淨利＋折舊－資本支出，擴產期現金流偏高估。"
            "Comparable：健鼎、華通、臻鼎-KY、博智、TTM 前瞻 P/E 中位數（臺美同業，不放 A 股；臺股同業價格多為 8–9 月快取，TTM 倍數低），方向不確定。"
            "情境對應：A＝順利（顆數基準、份額價格高值）、B＝平平（顆數 ×0.8／0.85、基準）、C＝不如預期（顆數 ×0.6、低值）")
for r_ in (24, 36, 40):
    V[f"B{r_}"].fill = KEY
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 B＝基準輸入；有效份額、其他營收毛利率在每個案例內依 FY2026 錨點重算，FY2026 恆等於錨點）：對 FY2027、FY2028 的營收、AI 相關營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 AI 相關", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 AI 相關", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("每 GW 金額成長 低", "勝宏模型每 GW 金額增量 ×0.7", dict(pgwf="pgwf_lo")),
         ("每 GW 金額成長 高", "增量 ×1.3", dict(pgwf="pgwf_hi")),
         ("中板／背板不參與", "參與係數 0", dict(wmid="w_mid_lo")),
         ("中板／背板全參與", "參與係數 1", dict(wmid="w_mid_hi")),
         ("份額流失（低）", "份額相對值 2027 0.85、2028 0.75、2029 0.70（ASIC 改 HDI 混壓情境）", dict(rel="RLO")),
         ("份額提升（高）", "份額相對值 2027–28 1.08、2029 1.05", dict(rel="RHI")),
         ("價格回落快", "價格指數 2027 0.96、2028 0.92、2029 0.88", dict(pp="PPLO")),
         ("價格延續", "價格指數 2027 1.03、2028 1.02、2029 1.00", dict(pp="PPHI")),
         ("擴產延遲", "產能指數 2027 1.30、2028 1.60、2029 1.90", dict(cx="CAPLO")),
         ("擴產提前", "產能指數 2027 1.60、2028 2.15、2029 2.60", dict(cx="CAPHI")),
         ("AI 毛利率未改善", "額外改善 2027–29 皆 −1pt", dict(gmu="GULO")),
         ("AI 毛利率改善 快", "額外改善 2027 +2pt、2028 +3pt、2029 +3.5pt", dict(gmu="GUHI")),
         ("AI 部署 ×0.8", "FY2027 起 AI IT 資本支出 ×0.8", dict(mf="mf_lo")), ("AI 部署 ×1.2", "×1.2", dict(mf="mf_hi")),
         ("AI 毛利率差距 6pt", "AI 相關毛利率只高 6 個百分點", dict(d="d_ai_lo")), ("AI 毛利率差距 14pt", "高 14 個百分點", dict(d="d_ai_hi")),
         ("非 AI 成長 低", "2027–29 −5%", dict(gn="GNLO")), ("非 AI 成長 高", "2027 +10%、2028–29 +8%", dict(gn="GNHI")),
         ("營業費用彈性 1.0", "基準 0.6", dict(phi="phi_hi")),
         ("淨利 ÷ 稅前 0.65", "基準 0.685", dict(nir="nir_f_lo")), ("淨利 ÷ 稅前 0.75", "", dict(nir="nir_f_hi"))]
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
put_case(nn, "情境 A（全部）", "顆數係數 A＋份額、價格、AI 毛利率改善高值", PB["A"], "PL!")
put_case(nn + 1, "情境 C（全部）", "顆數係數 C＋份額、價格、AI 毛利率改善低值", PB["C"], "PL!")
SENS_LAST = nn + 1; SENS_NOTE = nn + 3
SE.cell(nn + 3, 1, "讀法：基準為情境 B（顆數係數 B 從 CY2028 起）。產能受限的年度，需求面輸入（每 GW 金額、份額、AI 部署）對營收的影響被產能上限截斷，只有產能與價格會改變營收；價格指數同時改變產能上限營收與 AI 毛利率（成本不變），所以對 EPS 的影響大於對營收。")
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
RC["A1"] = "Recon — 本模型 vs 共識、券商估計與勝宏模型的每 GW PCB 金額、台光電模型的每 GW CCL 金額。只作事後對照，不為對上數字而調參數。共識取自 MarketScreener 與 FactSet（理財周刊轉引）；券商多為單一來源"
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
rec("FY2026 營收 vs 共識（MarketScreener）", "=cons_rev26", pl("rev", 2026), "MarketScreener（2026-10-25 擷取）", "stockanalysis 8 家 1,063 億；共識隱含 4Q26 約 327 億", fmt="#,##0.0")
rec("FY2026 EPS vs 共識（MarketScreener）", "=cons_eps26", pl("eps", 2026), "MarketScreener", "3Q26 財報（11 月上旬）前")
rec("FY2026 EPS vs FactSet 中位數（2026-09）", "=fs_eps26", pl("eps", 2026), "FactSet（理財周刊 2026-10-14 轉引）", "")
rec("FY2027 營收 vs 共識（MarketScreener）", "=cons_rev27", pl("rev", 2027), "MarketScreener", "", fmt="#,##0.0")
rec("FY2027 EPS vs 共識（MarketScreener）", "=cons_eps27", pl("eps", 2027), "MarketScreener", "")
rec("FY2027 EPS vs FactSet 中位數", "=fs_eps27", pl("eps", 2027), "FactSet（理財周刊轉引）", "")
rec("FY2027 EPS vs 外資（2026-05-13）", "=fb_eps27", pl("eps", 2027), "technews 2026-05-13（外資，名稱未載明）", "目標價 1,310")
rec("FY2028 營收 vs 共識（MarketScreener）", "=cons_rev28", pl("rev", 2028), "MarketScreener", "", fmt="#,##0.0")
rec("FY2028 EPS vs 共識（MarketScreener）", "=cons_eps28", pl("eps", 2028), "MarketScreener", "")
rec("FY2028 EPS vs FactSet 中位數", "=fs_eps28", pl("eps", 2028), "FactSet（理財周刊轉引）", "")
RC.cell(rr_, 1, "B. 每 GW PCB 金額 vs 勝宏模型、台光電每 GW CCL 金額（機制檢查）").font = BOLD; rr_ += 1
for y in (2026, 2027, 2028):
    rec(f"CY{y} 每 GW 可參與 PCB 金額（本模型）vs 勝宏模型每 GW PCB 金額", f"=vg_pgw{y % 100}", pl("pgw", y), "victory_giant/dist Units 頁（唯讀）",
        "差距只來自中板／背板參與係數 w_mid（與成長倍率 pgwf）" if y == 2026 else "—", fmt="#,##0")
for y in (2026, 2027):
    rec(f"CY{y} 每 GW 可參與 PCB 金額（本模型）vs 台光電模型每 GW CCL 金額", f"=emc_ccl{y % 100}", pl("pgw", y), "emc/dist Units 頁（唯讀）",
        "差距＝PCB ÷ CCL − 1；勝宏模型已註明 PCB ÷ CCL 只有約 1.3 倍，低於 Morgan Stanley 每架口徑 5.2 倍（範圍不同）" if y == 2026 else "—", fmt="#,##0")
rec("CY2027 每 GW 金額指數（÷ CY2026）：本模型 vs 台光電 CCL", "=emc_ccl27/emc_ccl26", lambda s: f"=PL!E{PB[s]['pgw']}/PL!D{PB[s]['pgw']}", "emc/dist", "世代升級幅度的對照", fmt="0.00")
for y in (2026, 2027):
    rec(f"CY{y} 每 GW 金像電 AI 營收 vs 每 GW 勝宏 AI 營收（US$M）", f"=vg_pge{y % 100}", pl("pgwe", y), "victory_giant/dist", "同為 AI 高速 PCB 廠；勝宏較偏 NVIDIA GPU 加速卡與 HDI" if y == 2026 else "—", fmt="#,##0")
    rec(f"CY{y} 每 GW 金像電 AI 營收 vs 每 GW 台光電 AI 營收（US$M）", f"=emc_pg{y % 100}", pl("pgwe", y), "emc/dist", "下游 PCB 廠 ÷ 上游 CCL 廠", fmt="#,##0")
rec("FY2026 金像電有效份額 vs 勝宏有效份額（勝宏模型）", "=vg_sh26", lambda s: f"=PL!D{PB[s]['sh0']}", "victory_giant/dist", "分母不同：本模型為可參與市場（中板 × w_mid）", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "C. 產能與校準").font = BOLD; rr_ += 1
rec("FY2027 需求面 ÷ 產能上限", None, pl("util", 2027), "—", "大於 100%＝產能受限", fmt="0.0%")
rec("FY2028 需求面 ÷ 產能上限", None, pl("util", 2028), "—", "", fmt="0.0%")
rec("其他營收毛利率（Derived）", None, lambda s: f"=PL!D{PB[s]['gmn']}", "FY2026 錨點", "AI 相關毛利率＝此值＋d_ai", fmt="0.0%")
rec("FY2027 毛利率 vs FY2026（百分點）", "=a26_gp/a26_rev", pl("gm", 2027), "錨點", "", fmt="0.0%", pt=True)
rec("FY2027 營業利益率 vs 2Q26 實績 26.43%（百分點）", "=0.2643", pl("oim", 2027), "理財周刊 2026-10-14", "", fmt="0.0%", pt=True)
rec("4Q26 EPS（模型）", None, lambda s: "=Q!E15", "—", "3Q26 財報 11 月上旬公布", fmt="0.00")
RC.cell(rr_, 1, "D. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以部署路徑取代「AI IT 資本支出 ÷ 每 GW IT 資本支出」"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RECON_LAST = rr_ - 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝金像電財報、月營收公告與法說會，經資料商與媒體轉述；券商與媒體估計皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("二手：優分析 2026-08-11（1H26、2Q26 營收、毛利率、淨利、EPS；7 月營收）", UA2Q),
        ("資料商：stockanalysis 季度損益（1Q25–1Q26 營收、毛利、營業利益、稅前淨利率、淨利）", SAQ),
        ("二手：technews 2026-10-08（9 月營收 93.20 億、前 9 月 733.79 億）", TN9),
        ("二手：優分析 2026-09-09（8 月營收、AI 占比約五成、泰國廠產值、2026 資本支出、中壢新廠、ASIC 設計風險）", UA8M),
        ("二手：優分析 2026-10-05（Amazon 私募、79 億元資本支出）", UAAMZ),
        ("二手：理財周刊 2026-10-14（3Q26 營收、2Q26 營業利益率與淨利率、FactSet EPS 中位數、Amazon 私募股數）", MW10),
        ("二手：理財周刊 2026-09-23（2026 資本支出 190 億、各廠擴產時程、產能年增）", MW9),
        ("二手：優分析 2026-03-10（FY2025 營收、毛利率、淨利、EPS 19.47、股利、擴產）", UAFY),
        ("二手：經濟日報 2026-03-11（FY2025 淨利、4Q25）", UDNFY),
        ("公司法說（富果整理 2025-11-28；3Q25 產品組合、資本支出、生產基地）", FUG),
        ("二手：鉅亨 2025-09-26（可轉債逾 90 億元、泰國一二期）", CNY),
        ("券商：外資（technews 2026-05-13；FY2027 EPS 65.39、目標價 1,310）", TNFB),
        ("資料商：stockanalysis（現價、前瞻與近 12 月 P/E、股數、共識目標價、FY2021–25 損益、ROE；同業前瞻 P/E）", SA),
        ("資料商：stockanalysis forecast（共識目標價 9 位、FY2026 營收與 EPS 8 家）", SAF),
        ("資料商：MarketScreener（FY2026–28 共識營收、EPS；FY2021–25 P/E）", MSV),
        ("資料商：聚財網（2014–2020 年度 EPS）", WEARN),
        ("資料商：digrin（2015–2025 年 12 月收盤）", DIGRIN),
        ("資料商：HiStock（12 月 P/E 2021–2025；年度 EPS 2018–2025）", "https://histock.tw/stock/2368/%E6%9C%AC%E7%9B%8A%E6%AF%94"),
        ("市場層：markets/network、markets/chip_bridge、tsmc 成品；victory_giant 成品（每 GW PCB 金額）、emc 成品（台光電每 GW CCL，Recon 用）（本 repo，唯讀）", ""),
        ("無法取得：companiesmarketcap 金像電 P/E 頁（網站錯誤）；10-23 收盤", "未找到"),
        ("待補（人工）：3Q26 財報與法說（11 月上旬）；10 月營收（11-10 前後）；分產品（AI 伺服器、交換器）營收與毛利率；2016–2020 營收與淨利；產能（萬㎡）與產值的公司原文；可轉債轉換價；資本支出與折舊路徑；前瞻 P/E 十年歷史（Bloomberg）", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 100; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("金像電（2368.TW）收支與估值模型 v0.1（AI 半導體 Project：載板與 PCB 次產業；伺服器與交換器 PCB、高層數板）　2026-10-25", True),
 ("版本：v0.1（初版）。每 GW AI 高速 PCB 金額沿用勝宏模型（victory_giant/dist Units 頁，本專案權威，口徑一致），本模型只加金像電對中板／背板的參與係數；金像電份額與產值產能在本模型 Units 頁建立；Recon 頁對照勝宏模型與台光電每 GW CCL 金額。", False),
 ("模型命題：金像電的 AI 營收＝部署 GW（AI IT 資本支出 ÷ 每 GW IT 資本支出）× 每 GW 可參與 AI 高速 PCB 金額（勝宏模型三項，中板／背板 × 參與係數）× 金像電份額 × 價格指數，但受產值產能（FY2026 營收 × 擴產指數）限制（全產全銷）。非 AI（一般伺服器、網通、筆電）另設成長率。", False),
 ("核心問題：2026 年約 190 億元資本支出（台灣、蘇州、泰國二廠 2027 下半年開出）能讓產能在 2027–29 年擴大多少；ASIC 主板改 HDI＋HLC 混壓後份額守不守得住；現價約 20 倍前瞻 P/E 隱含什麼路徑。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029 由驅動因子推導，FY2030–FY2032 為收斂法。金額：十億新台幣；每 GW 金額以美元計、以匯率 32.0 換算。口徑：歸屬母公司淨利（IFRS）。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("需求面：markets/network 的 AI IT 資本支出 × 同字母顆數係數（tsmc；＝markets/memory SCEN_UNITS）÷ 每 GW IT 資本支出（Tokenomics，經橋接表）＝部署 GW；× 每 GW PCB 金額（勝宏模型，US$450M → 682M）× 匯率 × 有效份額（FY2026 以錨點 AI 營收校準）× 份額相對值 × 價格指數（Assumed）。", False),
 ("產能面：FY2026 營收（產能吃緊、滿載）× 產值產能指數（2027 1.45、2028 1.90、2029 2.30；Assumed，依 2026 年約 190 億元資本支出與各廠時程）× 價格指數。AI 相關營收＝MIN（需求面、產能上限 − 非 AI）。", False),
 ("毛利率：AI 相關比其他營收高 d_ai 10pt（Assumed），其他毛利率為 FY2026 殘差（Derived）；價格指數變動全部落在 AI 毛利；另加新廠爬坡完成的額外改善（Assumed）。營業費用：彈性 phi 0.6（Assumed）。淨利 ÷ 稅前 0.685（1Q26 實績）。", False),
 ("校準：FY2026 ＝ 1H26 財報＋3Q26 月營收實績（297.19 億）＋假設（3Q、4Q 毛利率 35%、營業費用率 8%、4Q 季增 8%）。3Q26 財報與法說預計 2026-11 上旬，截至 2026-10-25 尚未公布。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，三情境權重 0.3／0.5／0.2，負值取 0）＋0.5 × Comparable 目標價（臺美同業健鼎、華通、臻鼎-KY、博智、TTM 前瞻 P/E 中位數 × 12 個月後前瞻 EPS；不放 A 股）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解。十年 P/E 規則值 × FY2027 EPS 保留為參考（不是目標價）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m，FY2032 達到；PL 頁灰底）。十年平均淨利率未找到（2016–2020 營收與淨利缺），依工作單以情境 C FY2029 淨利率 × 0.8 替代（Assumed）。", False),
 ("情境與行業一致：A＝順利（tsmc 顆數係數 A、份額相對值／價格指數／AI 毛利率改善取高值）；B＝平平（係數 B、取基準）；C＝不如預期（係數 C、取低值）；產能、非 AI 成長、每 GW 金額各情境相同。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司數字取自財報與月營收（經資料商與媒體轉述）；券商與資料商為 Interested-party。", False),
 ("", False),
 ("已知限制：(1) 每 GW PCB 金額沿用勝宏模型（2026 水準 Assumed、世代指數 Analogy／Assumed），其升級主要來自 NVIDIA VR200 中板與 Kyber 背板；金像電以 ASIC 與交換器為主，參與係數 w_mid 為 Assumed；(2) 產值產能指數為 Assumed，公司未揭露產能（萬㎡）；(3) AI 占比 FY2025 取媒體轉述（約兩成到三成），FY2026 由非 AI 成長假設反推；(4) 未建資本支出與折舊（2026 年資本支出約 190 億、折舊約 14 億），DCF 現金流在擴產期偏高估；(5) 淨利 ÷ 稅前約 0.68（稅負偏高，含中國子公司），未拆分；(6) 十年 P/E 只有一個採用來源，2016–2019 年低獲利使 P/E 偏高；(7) 臺股同業前瞻 P/E 為 8–9 月快取值；(8) 3Q26 財報後更新錨點；R4 產業部署路徑未建。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ---------------- v0.1.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00／14:17）----------------
# 新列一律放在 Val、README 工作表既有內容之後，不位移既有列；既有數值不變
VER_PREV, VER_NOW = 'v0.1', 'v0.1.1'
OBS = [
 "目標價 NT$1,267（DCF 1,203、Comparable 1,331 各半），比現價 1,115（2026-10-22）高 13.7%，級距 0%–25%；Comparable 單獨為 +19.4%。兩法接近；與共識目標價 1,554 的差距在倍數：本模型前瞻 EPS 72.00 比共識前瞻 EPS 56.97 高 26.4%（期間約晚一年），但同業（健鼎、華通、臻鼎-KY、博智、TTM）前瞻 P/E 中位數只有 18.5 倍，共識目標價隱含 27.3 倍。",
 "擴產進度最敏感：擴產延遲（2028 年產能指數 1.60）使 FY2028 EPS −18.3%、價格回落快 −13.5%、份額流失 −10.7%；情境 C 的 FY2028 EPS 45.31 比情境 B 76.56 低 40.8%，DCF 價格帶 559～1,723。",
 "本模型情境 B 的 FY2027 EPS 60.15 比共識 62.80 低 4.2%、FY2028 EPS 76.56 比共識 95.35 低 19.7%。下一個檢驗點是 3Q26 財報與法說（預計 2026-11 上旬，stockanalysis 估 11-11）與 10 月營收（11-10 前後）。"
]   # 重點觀察（數字取自重算後的 Val／PL／Sens；與 README.md 相同文字）
UPW_ROW = next(r_ for r_ in range(1, V.max_row + 1) if V.cell(r_, 1).value == "上檔空間：加權目標價")
GRADE_ROW = V.max_row + 1; REC_ROW = GRADE_ROW + 1
_u = f"B{UPW_ROW}"
V.cell(GRADE_ROW, 1, "上檔空間級距")
V[f"B{GRADE_ROW}"] = f'=IF({_u}>0.5,"大於 50%",IF({_u}>0.25,"25%–50%",IF({_u}>=0,"0%–25%",IF({_u}>=-0.25,"0%～−25%",IF({_u}>=-0.5,"−25%～−50%","小於 −50%")))))'
V[f"B{GRADE_ROW}"].number_format = "@"
V[f"E{GRADE_ROW}"] = f"以「上檔空間：加權目標價」（第 {UPW_ROW} 列）分六級（Andy 2026-10-25 14:00）"
V.cell(REC_ROW, 1, "投資建議")
V[f"B{REC_ROW}"] = f'=IF({_u}>0.5,"強力買進",IF({_u}>0.25,"買進",IF({_u}>=-0.25,"中立",IF({_u}>=-0.5,"放空","強力放空"))))'
V[f"B{REC_ROW}"].number_format = "@"
V[f"E{REC_ROW}"] = "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空（剛好 +25% 為中立、+50% 為買進、−25% 為中立、−50% 為放空）"
_t = RD["A1"].value
assert VER_PREV in _t and VER_NOW not in _t, _t
RD["A1"] = _t.replace(VER_PREV, VER_NOW, 1)
for r_ in range(2, RD.max_row + 1):
    _v = RD.cell(r_, 1).value
    if isinstance(_v, str) and _v.startswith("版本："):
        RD.cell(r_, 1, "版本：" + VER_NOW + "（2026-10-25：補上檔空間級距、投資建議、重點觀察，數值不變）。前版：" + _v[3:])
        break
rn = RD.max_row + 1
RD.cell(rn, 1, VER_NOW + "：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁最後新增「上檔空間級距」「投資建議」兩列（第 " + str(GRADE_ROW) + "、" + str(REC_ROW) + " 列，依「上檔空間：加權目標價」的固定規則）；重點觀察見下方三列。")
RD.cell(rn + 1, 1, f'="投資建議："&Val!B{REC_ROW}&"（依上檔空間固定規則）"')
for k_, o_ in enumerate(OBS, 1):
    RD.cell(rn + 1 + k_, 1, f"重點觀察 {k_}：" + o_)

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(sys.argv[1])
if os.environ.get("GCE_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE,
               "UNOTE": UNOTE, "T": T, "TPS0": TPS0, "TPS_LAST": TPS_LAST, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0, "block_end": block_end},
              open(os.environ["GCE_ROWMAP"], "w"), ensure_ascii=False, indent=1)

# Amkor Technology（AMKR）收支與估值模型 v0.2.1（AI 半導體 Project：封測次產業）
# 用法：python3 amkr/build_amkr.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：百萬美元（US$M）；EPS：美元（稀釋）。
# 口徑：歸屬母公司淨利（GAAP）。結構比照 quanta/build_quanta.py：README、In、Q、Units、PL、Val、Recon、Sens、Sources。
# 加速器顆數、情境係數與 CoWoS 片數口徑唯讀取自 tsmc/dist；AI IT 資本支出取自 markets/network/dist；每 GW IT 資本支出取自 markets/chip_bridge/dist。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
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
TSMC_F = only("tsmc/dist/*.xlsx"); NET_F = only("markets/network/dist/*.xlsx"); BR_F = only("markets/chip_bridge/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
tw = openpyxl.load_workbook(TSMC_F, data_only=True)
TS = {}
for r in tw["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and (str(k).startswith("u_") or str(k).startswith("g29_") or str(k).startswith("um") or k == "osat"):
        TS[k] = (r[4].value, r[0].row)
PLATS = ("bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth")
assert all(f"u_{p}_{y}" in TS for p in PLATS for y in (26, 27, 28)) and all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()
UNITS = {y: sum(TS[f"u_{p}_{y}"][0] for p in PLATS) for y in (26, 27, 28)}
UNITS[29] = sum(TS[f"u_{p}_28"][0] * (1 + TS[f"g29_{p}"][0]) for p in PLATS)
U_SRC = f"{rel(TSMC_F)}｜In!E{TS['u_bw_26'][1]}:E{TS['g29_oth'][1]}（八類加速器加總；2029＝2028 × (1＋各類 2029 成長)，與台積電模型 CoWoS 頁同口徑）"
CW = {}
for r in tw["CoWoS"].iter_rows(min_row=4, max_row=30):
    lab = str(r[0].value or "").strip()
    if lab in ("需求合計", "其他 CoWoS 用途", "需求 ÷ 可用產能"):
        CW[lab] = ([r[1 + j].value for j in range(4)], f"{rel(TSMC_F)}｜CoWoS!B{r[0].row}:E{r[0].row}")
assert len(CW) == 3, CW.keys()
WPU = {y: (CW["需求合計"][0][j] - CW["其他 CoWoS 用途"][0][j]) / UNITS[y] for j, y in enumerate((26, 27, 28, 29))}
AREA = {y: WPU[y] / WPU[26] for y in WPU}
ITC = None
for r in openpyxl.load_workbook(NET_F, data_only=True)["Mkt"].iter_rows(min_row=4):
    if r[0].value == "AI IT 資本支出（加速器 × 每顆 IT capex）":
        ITC = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
assert ITC and all(v is not None for v in ITC[0]), ITC
CAPIT = None
for r in openpyxl.load_workbook(BR_F, data_only=True)["In"].iter_rows(min_row=4):
    if r[0].value == "capit":
        CAPIT = (r[4].value, r[7].value, f"{rel(BR_F)}｜In!E{r[0].row}（GB300）、H{r[0].row}（VR200）")
assert CAPIT, "capit"

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：百萬美元（US$M）；顆數：百萬顆；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間，不當點估計（放情境或 Sens）。SemiAnalysis 未使用。公司原文＝SEC 8-K 新聞稿、10-K、投資人日簡報；法說內容經 Motley Fool 逐字稿與 MarketBeat 摘要；公司指引與目標為 Interested-party；券商與資料商為 Interested-party"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
PR2Q = "https://www.sec.gov/Archives/edgar/data/0001047127/000104712726000043/amkr6302026erex-991.htm"
PR4Q = "https://www.sec.gov/Archives/edgar/data/1047127/000104712726000007/amkr123125erex-991.htm"
K10 = "https://www.sec.gov/Archives/edgar/data/1047127/000104712726000014/amkr-20251231.htm"
IDAY = "https://www.sec.gov/Archives/edgar/data/0001047127/000104712726000036/ex-991amkorinvestorprese.htm"
FOOL4Q = "https://www.fool.com/earnings/call-transcripts/2026/02/09/amkor-amkr-q4-2025-earnings-call-transcript/"
FOOL1Q = "https://www.fool.com/earnings/call-transcripts/2026/04/27/amkor-amkr-q1-2026-earnings-call-transcript/"
MB2Q = "https://www.marketbeat.com/instant-alerts/amkor-technology-q2-earnings-call-highlights-2026-07-27/"
SA = "https://stockanalysis.com/stocks/amkr/"

sec("A. AI 加速器顆數、封裝面積與部署規模（讀 tsmc、markets/network、markets/chip_bridge 成品；唯讀，不自行重估）")
for y in (26, 27, 28, 29):
    inp(f"units_{y}", f"AI 加速器顆數 20{y}（情境 A 基準）", "百萬顆", None, round(UNITS[y], 6), None, "Derived", U_SRC if y == 26 else "同上",
        "NVIDIA、AMD、Google TPU、AWS Trainium、其他 ASIC 合計；台積電模型讀 markets/memory 的同一份顆數" if y == 26 else "同上")
for s in SC:
    for y in (28, 29):
        inp(f"um{y}_{s}", f"情境 {s}：20{y} 顆數 ÷ 基準", "x", None, TS[f"um{y}_{s}"][0], None, "Assumed", f"{rel(TSMC_F)}｜In!E{TS[f'um{y}_{s}'][1]}",
            "A＝基準；B＝AI 資本支出消化；C＝明顯下修（與台積電、廣達、NVIDIA 模型相同）" if (s == "A" and y == 28) else "同上")
for y in (26, 27, 28, 29):
    inp(f"ar_{y}", f"封裝面積指數 20{y}（加速器用 CoWoS 片數 ÷ 顆數，2026＝1）", "x", None, round(AREA[y], 6), None, "Derived",
        f"{CW['需求合計'][1]} 減 {CW['其他 CoWoS 用途'][1]}，除以顆數" if y == 26 else "同上",
        "每片封裝數下降＝單顆封裝面積變大（Rubin、Rubin Ultra）；台積電模型的每片封裝數為 Derived／Assumed" if y == 26 else "同上")
inp("cow_util27", "台積電 CoWoS 需求 ÷ 可用產能 2027（Recon 用）", "%", None, CW["需求 ÷ 可用產能"][0][1], None, "Derived", CW["需求 ÷ 可用產能"][1], "台積電模型：可用產能＝台積電產能 ×（1＋委外補充）")
inp("osat_ts", "台積電模型：委外封測補充產能 ÷ 台積電（Recon 用）", "%", None, TS["osat"][0], None, "Assumed", f"{rel(TSMC_F)}｜In!E{TS['osat'][1]}", "台積電模型的全流程委外口徑；本模型的委外比例為後段（oS、測試）口徑，兩者不同")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}（Recon 用：部署 GW）", "US$B", None, ITC[0][j], None, "Derived", ITC[1] if j == 0 else "同上", "加速器顆數 × 每顆 IT capex" if j == 0 else "同上")
inp("capit_gb", "每 GW IT 資本支出（GB300；2026 用）", "US$B", None, CAPIT[0], None, "Derived", CAPIT[2], "Tokenomics v20.1 IF_CapexIT（經橋接表；唯讀）")
inp("capit_vr", "每 GW IT 資本支出（VR200；2027 起）", "US$B", None, CAPIT[1], None, "Derived", CAPIT[2], "同上")
inp("tam_gw", "OSAT 先進封裝可服務市場 ÷ 每 GW 部署（Recon 用）", "US$B/GW", None, 0.5, None, "Interested-party", IDAY + "：約 US$0.5B／GW installed（公司投資人日 2026-05-21）", "第二來源：未找到")

sec("B. AI 先進封裝（2.5D、HDFO、AI 測試）：營收＝顆數 × 後段委外比例 × Amkor 份額 × 每顆封測金額（FY2026 以公司「三倍」說法校準）")
inp("tri", "FY2026 AI 先進封裝營收 ÷ FY2025（公司：約三倍）", "x", 2.5, 3.0, 3.2, "Interested-party", FOOL1Q + "：Advanced AI packaging revenue still on track to triple in 2026（1Q26 法說）", FOOL4Q + "：2.5D and HDFO … nearly triple over the course of this year（4Q25 法說）；低＝供應受限時")
inp("g_pc26", "FY2026 非 AI 運算（PC、儲存、基礎設施）年增", "%", -0.10, -0.03, 0.05, "Assumed", "公司：PC／筆電疲弱（1Q26、4Q25 法說）；運算成長由資料中心帶動", "非 AI 運算金額未揭露：未找到")
for y, v, lo, hi in ((26, 0.50, 0.40, 0.60), (27, 0.55, 0.45, 0.65), (28, 0.60, 0.45, 0.70), (29, 0.60, 0.45, 0.75)):
    inp(f"out_{y}", f"AI 加速器後段（oS、測試）委外比例 20{y}", "%", lo if lo is not None else None, v, hi, "Assumed",
        "台積電 CoWoS 2027 需求 ÷ 產能約 96%（台積電模型）；台積電與 Amkor 十年合作（2026-06-16）、NVIDIA 與 Amkor 多年合作 US$1.5B（2026-07-23）" if y == 26 else "同上",
        "FY2026 只影響拆分（每顆金額由校準吸收）；FY2027 起的相對變化才影響營收；委外比例的量化：未找到" if y == 26 else "Rubin 世代 CoWoS 緊、美國產能需求 → 委外上升（Assumed）")
for y, v, lo, hi in ((26, 0.35, 0.28, 0.40), (27, 0.35, 0.28, 0.40), (28, 0.37, 0.28, 0.45), (29, 0.38, 0.28, 0.45)):
    inp(f"sh_{y}", f"Amkor 占委外後段的份額 20{y}", "%", lo, v, hi, "Assumed",
        "Amkor 為全球第二大封測廠；競爭者日月光（含矽品）、京元電（測試）" if y == 26 else "同上",
        "份額的量化：未找到（日月光模型待 10-29 法說後建立）；FY2026 只影響拆分" if y == 26 else "NVIDIA 合作與美國產能（Assumed）")
inp("el", "每顆封測金額對封裝面積的彈性", "x", 0.3, 0.5, 1.0, "Assumed", "封裝面積大 → 基板上封裝、測試時間與良率成本上升；後段金額不與面積等比例", "第二來源：未找到")

sec("C. 其他終端市場成長（FY2027–FY2029）")
for s, a, b, lo, hi, txt in (("A", 0.03, 0.04, -0.02, 0.06, "Apple 新機與高階組合；SiP 移轉越南後回到正常季節性"),
                              ("B", 0.0, 0.03, -0.05, 0.04, "通訊逆風延續到 1H27（公司：一項逆風到 4Q 與明年上半年）"),
                              ("C", -0.06, 0.0, -0.12, 0.02, "Apple 分散封測供應或自行轉移（客戶集中風險）")):
    inp(f"g_cm27_{s}", f"通訊年增 FY2027（情境 {s}）", "%", lo, a, hi, "Assumed", txt, "Apple 占 FY2025 營收 29.8%、Qualcomm 11.1%（10-K）；投資人日通訊 2025–30 CAGR 約 4%" if s == "A" else "同上")
    inp(f"g_cmL_{s}", f"通訊年增 FY2028–29（情境 {s}）", "%", lo, b, hi, "Assumed", "同上", "同上")
inp("g_au", "車用與工業年增 FY2027–29", "%", 0.04, 0.10, 0.14, "Interested-party", IDAY + "：車用與工業 2025–30 CAGR 約 10%（公司）", "公司：ADAS 與資訊娛樂帶動先進車用強勁成長（1Q26、2Q26 法說）")
inp("g_cs", "消費年增 FY2027–29", "%", -0.03, 0.03, 0.06, "Interested-party", IDAY + "：消費 2025–30 CAGR 約 3%（公司）", "第二來源：未找到")
inp("g_pc", "非 AI 運算年增 FY2027–29", "%", -0.05, 0.03, 0.06, "Assumed", "投資人日：運算 2025–30 CAGR 約 24%（含 AI；公司）", "非 AI 運算拆分：未找到")

sec("D. 亞利桑那廠（Peoria；第一期約 US$3.5B，2027 完工、2028 投產；第二期不在本模型）")
for s, a, b in (("A", 150, 500), ("B", 100, 400), ("C", 50, 250)):
    inp(f"az_rev28_{s}", f"亞利桑那廠營收 FY2028（情境 {s}）", "US$M", None, a, None, "Assumed", IDAY + "：第一期滿載約 US$1B 營收、2029 損益兩平、2030 滿載（公司）" if s == "A" else "同上",
        FOOL1Q + "：meaningful revenue by late 2029、full impact by 2030（1Q26 法說）" if s == "A" else "同上")
    inp(f"az_rev29_{s}", f"亞利桑那廠營收 FY2029（情境 {s}）", "US$M", None, b, None, "Assumed", "同上", "同上")
inp("az_cm", "亞利桑那廠現金邊際貢獻率（營收 − 變動現金成本）", "%", 0.45, 0.55, 0.60, "Assumed", "公司：滿載毛利率 30% 以上（投資人日）；以本參數與固定成本、折舊推得滿載毛利率約 30%", "第二來源：未找到")
for y, v, lo, hi in ((27, 120, 80, 170), (28, 100, 70, 150), (29, 100, 70, 150)):
    inp(f"az_fix{y}", f"亞利桑那廠固定現金成本（啟動、人力）FY20{y}", "US$M", lo, v, hi, "Assumed",
        FOOL1Q + "：啟動成本自 2027 起稀釋營業利益率約 1–2%，2028 改善（1Q26 法說）" if y == 27 else "同上",
        "MarketBeat 2Q26 摘要：Q3 之後毛利率承受美國低稼動產能（含亞利桑那折舊）負擔" if y == 27 else "同上")
inp("az_capex1", "亞利桑那第一期投資", "US$M", 3000, 3500, 4000, "Interested-party", FOOL4Q + "：Phase One（約 US$7B 美國計畫的一半）施工中（4Q25 法說）", "第二期擴大至合計約 US$12B（公司 2026-09-08），2027 年底動工、2029 年底完工，不在本模型")
inp("az_inc", "政府補助與投資抵減 ÷ 投資（降低折舊基礎）", "%", 0.20, 0.35, 0.40, "Interested-party", FOOL1Q + "：CHIPS 補助約 US$400M＋35% 投資抵減，合計約 US$2.8B 對應 US$7B（1Q26 法說）", IDAY + "：投資抵減約 US$2B、CHIPS 約 US$400M；會計上是否抵減資產成本：未找到（Assumed）")
inp("az_life", "亞利桑那資產平均折舊年限（廠房與設備混合）", "年", 8, 12, 15, "Assumed", "10-K：測試設備年限 2024 年起由 5 年延長為 7 年（投資人日附錄）；廠房年限較長", "混合年限：未找到")
inp("az_svc28", "FY2028 折舊期間占比（2028 年初投產）", "x", 0.25, 0.5, 0.75, "Assumed", "公司：2027 完工、2028 投產（投資人日）", "—")

sec("E. 毛利率、折舊、營業費用與淨利（FY2027 起）")
inp("ft", "本業現金毛利增量 ÷ 營收增量（現金流通率）", "%", 0.30, 0.40, 0.50, "Assumed", "FY2025→FY2026 錨點實測見 Q 頁（稼動率由 50% 區間升到 70% 區間，屬一次性改善）", FOOL4Q + "：公司預期 2026 年增量流通率約 30%（含折舊、扣除資產出售）")
inp("capi", "本業資本支出 ÷ 本業營收（FY2027 起，不含亞利桑那）", "%", 0.10, 0.14, 0.18, "Assumed", "FY2025 資本支出 904.6 ÷ 營收 6,708＝13.5%（公司）", "2026 年 30–35% 用於 HDFO、測試與先進封裝產能（公司）")
inp("cap_m26", "FY2026 本業資本支出（不含亞利桑那）", "US$M", 900, 1200, 1500, "Assumed", PR2Q + "：FY2026 資本支出約 US$2.5–3.0B（公司）；65–70% 為廠房（含亞利桑那、韓國松島）", "亞利桑那 2026 年支出：未找到")
inp("life", "本業資產平均折舊年限（新增資本支出）", "年", 5, 7, 9, "Interested-party", IDAY + "：測試設備年限 7 年（2024 年起）", "其他設備年限：未找到")
inp("phi", "營業費用成長對營收成長的彈性（FY2027 起）", "x", 0.3, 0.5, 1.0, "Assumed", "FY2026 實測見 Q 頁（含激勵費用上升，偏高）", "營業費用率 FY2025 7.0% → 2Q26 6.3%（扣除不動產處分利益前）")
inp("opx_gain", "2Q26 不動產處分利益（從營業費用扣除；FY2027 基期加回）", "US$M", None, 21, None, "Interested-party", MB2Q + "：營業費用 119 含 21 不動產處分利益（2Q26 法說）", FOOL1Q + "：2Q26 指引含約 20 處分利益")
for y, v in ((27, -20), (28, -40), (29, -60)):
    inp(f"nonop{y}", f"營業外淨額（其他收益 − 利息費用）FY20{y}", "US$M", v - 40, v, v + 40, "Assumed",
        "FY2025 −22.7、1H26 −0.7（公司）；2026-05 發行 US$1.15B 零息可轉債（公司）" if y == 27 else "同上",
        "亞利桑那第二期動工後現金下降、負債上升（Assumed）" if y == 27 else "同上")
inp("tax", "有效稅率（FY2027 起）", "%", 0.15, 0.20, 0.25, "Interested-party", IDAY + "：規劃稅率 20%（公司）", FOOL1Q + "：全年假設 20%；1Q26 12.8%、2Q26 14%（含一次性利益）")
inp("nci", "非控制權益（每年）", "US$M", None, 3, None, "Derived", "FY2025 2.3、1H26 1.7（公司損益表推得）", "—")
inp("sh_g", "稀釋股數年變動", "%", -0.01, 0.0, 0.01, "Assumed", "FY2025 248.5 → 2Q26 250.0 百萬股（公司）；買回授權 US$300M（投資人日）", "零息可轉債的轉換稀釋：轉換價未找到")

sec("F. FY2026 錨點（1H26 實績＋3Q26 公司指引＋4Q26 假設）")
for k, nm, v, lo, hi, s1 in (("g3_cm", "通訊", -0.08, -0.10, -0.06, "high-single-digit 季減"), ("g3_cp", "運算", 0.30, 0.25, 0.30, "近 30% 季增（AI 資料中心、HDFO 資料中心 CPU）"),
                             ("g3_au", "車用與工業", 0.05, 0.03, 0.07, "mid-single-digit 季增"), ("g3_cs", "消費", 0.15, 0.13, 0.17, "mid-teens 季增")):
    inp(k, f"3Q26 {nm}季增（公司說法）", "%", lo, v, hi, "Interested-party", MB2Q + "：" + s1 + "（2Q26 法說）", "Benzinga 2026-07-28 同（運算近 30%、通訊 high-single-digit 季減）" if k in ("g3_cm", "g3_cp") else "第二來源：未找到")
inp("q3_gm", "3Q26 毛利率（指引 18.5–19.5% 中點）", "%", 0.185, 0.19, 0.195, "Interested-party", PR2Q + "：3Q26 毛利率 18.5–19.5%（公司）", MB2Q + "：同；主因產品組合")
inp("q3_opx", "3Q26 營業費用（指引）", "US$M", None, 140, None, "Interested-party", MB2Q + "：約 US$140M（2Q26 法說）", "第二來源：未找到")
for k, nm, v, lo, hi in (("q4_cm", "通訊", 0.05, -0.05, 0.15), ("q4_cp", "運算", 0.05, 0.0, 0.12), ("q4_au", "車用與工業", 0.02, -0.03, 0.05), ("q4_cs", "消費", -0.05, -0.10, 0.0)):
    inp(k, f"4Q26 {nm}季增", "%", lo, v, hi, "Assumed", "公司未給 4Q 指引（3Q26 財報 2026-10-26 公布）" if k == "q4_cm" else "同上",
        "通訊：公司稱一項逆風延續到 4Q 與明年上半年（2Q26 法說），低於 4Q 正常季節性" if k == "q4_cm" else "—")
inp("q4_gm", "4Q26 毛利率", "%", 0.175, 0.185, 0.195, "Assumed", "3Q26 指引中點 19%；公司稱 3Q 之後毛利率看稼動與組合，並承受美國產能低稼動負擔（2Q26 法說）", "—")
inp("q4_opx", "4Q26 營業費用", "US$M", 130, 140, 150, "Assumed", "同 3Q26 指引", "—")
inp("q_nonop", "3Q、4Q26 營業外淨額（每季）", "US$M", -10, 0, 10, "Assumed", "1Q26 −4.0、2Q26 +3.3（公司）", "—")
inp("q_tax", "3Q、4Q26 有效稅率", "%", 0.15, 0.20, 0.22, "Interested-party", "公司全年規劃 20%（1Q26 法說）", "3Q26 指引淨利 180–205 與營業利益推得約 20%")
inp("q_nci", "3Q、4Q26 非控制權益（每季）", "US$M", None, 1.0, None, "Derived", "1Q26 0.6、2Q26 1.1（公司損益表推得）", "—")
inp("q_sh", "3Q、4Q26 稀釋股數", "百萬股", None, 250.0, None, "Derived", "3Q26 指引淨利 180–205 ÷ EPS 0.72–0.82 ≈ 250（公司）", "2Q26 249.978（公司）")
inp("q_da3", "3Q26 折舊攤提", "US$M", 180, 190, 200, "Assumed", "1Q26 171、2Q26 180（公司）；資本支出 70% 在下半年（1Q26 法說）", "—")
inp("q_da4", "4Q26 折舊攤提", "US$M", 185, 200, 215, "Assumed", "同上", "—")
inp("gain25", "4Q25 毛利中的資產出售利益（一次性）", "US$M", None, 30, None, "Interested-party", FOOL4Q + "：4Q25 毛利率 16.7% 含約 US$30M 資產出售（4Q25 法說）", "Futurum 2026-02 同")

sec("G. 公司指引與長期目標、共識（只作 Recon 對帳，不作校準）")
inp("g3_rev", "3Q26 營收指引（1,950–2,050 中點）", "US$M", 1950, 2000, 2050, "Interested-party", PR2Q, "MarketBeat 2Q26 摘要同")
inp("g3_ni", "3Q26 淨利指引（180–205 中點）", "US$M", 180, 192.5, 205, "Interested-party", PR2Q, "—")
inp("g3_eps", "3Q26 EPS 指引（0.72–0.82 中點）", "US$", 0.72, 0.77, 0.82, "Interested-party", PR2Q, "—")
inp("capex26", "FY2026 資本支出指引（2.5–3.0B 中點）", "US$M", 2500, 2750, 3000, "Interested-party", PR2Q, PR4Q + "（2026-02 首次給出，維持不變）")
inp("t28_rev", "公司 2028 營收目標（±US$0.5B）", "US$M", 8500, 9000, 9500, "Interested-party", IDAY + "（投資人日 2026-05-21）", "第二來源：未找到")
inp("t28_gm", "公司 2028 毛利率目標（±1 個百分點）", "%", 0.165, 0.175, 0.185, "Interested-party", IDAY, "—")
inp("t28_eps", "公司 2028 EPS 目標（±0.25）", "US$", 2.25, 2.50, 2.75, "Interested-party", IDAY, "含亞利桑那第一期；不含第二期")
inp("t30_rev", "公司 2030 營收目標（約 11B 以上）", "US$M", None, 11000, None, "Interested-party", IDAY, "—")
inp("t30_gm", "公司 2030 毛利率目標（約 22% 以上）", "%", None, 0.22, None, "Interested-party", IDAY, "—")
inp("t30_eps", "公司 2030 EPS 目標（約 5 以上）", "US$", None, 5.0, None, "Interested-party", IDAY, "—")
inp("cons_rev26", "共識營收 FY2026", "US$M", 7500, 7670, 8000, "Interested-party", "stockanalysis.com（11 位分析師；2026-10-23 擷取）", "tickflow 2026-10-24：7.47B（7 位）")
inp("cons_eps26", "共識 EPS FY2026（non-GAAP）", "US$", 2.45, 2.57, 2.73, "Interested-party", "stockanalysis.com（11 位；2026-10-23）", "MarketBeat：2.60（年度標示不明）")
inp("cons_eps27", "共識 EPS FY2027", "US$", None, 2.72, None, "Interested-party", "MarketBeat：明年 2.72（頁面未明確標示年度；單一來源）", "stockanalysis、Yahoo、Zacks 的 FY2027 數字：未找到（付費或無法開啟）")
inp("cons_q3rev", "共識 3Q26 營收", "US$M", None, 2040, None, "Interested-party", "MarketBeat（2026-10）", "Benzinga 2026-07-28：指引前共識 2.09B")
inp("apple", "Apple 占 FY2025 營收", "%", None, 0.298, None, "Verified", K10 + "：Apple 29.8%、Qualcomm 11.1%；前十大 72%（10-K）", "2Q26 前十大 66%（公司新聞稿）；第二來源：未找到")
sec("H. 價格與市場倍數")
inp("px", "股價（2026-10-23 收盤，Nasdaq）", "US$", None, 39.30, None, "Interested-party", SA + "：10-23 收盤 39.30、前收 39.05", "stocksguide 同 39.30；aaii：10-22 收盤 39.05，10 月以來 −26%（原因未找到）")
inp("pe_ttm", "近 12 月 P/E（stockanalysis，10-23）", "x", None, 17.64, None, "Interested-party", SA, "companiesmarketcap 17.64")
inp("fpe", "前瞻 P/E（stockanalysis，10-23）", "x", None, 14.02, None, "Interested-party", SA, "—")
inp("cons_feps", "共識前瞻 12 月 EPS（現價 ÷ 前瞻 P/E）", "US$", None, "=px/fpe", None, "Derived", "stockanalysis 前瞻 P/E 14.02（2026-10-23）", "共識前瞻 EPS 的直接數字：未找到")
inp("cons_tp", "市場共識目標價（賣方平均）", "US$", 65, 76.40, 92, "Interested-party", "stockanalysis.com：11 位分析師平均 76.40、中位數 73.50、區間 65–92（2026-10-23 擷取）", "stocksguide：13 位平均 71.40（2026-10-23）；多數目標價訂於 10 月股價下跌前")
REV10 = [3894, 4186, 4316, 4053, 5051, 6138, 7092, 6503, 6318, 6708]
NI10 = [164, 261, 127, 121, 338, 643, 766, 360, 354, 373.9]
inp("nim10", "十年平均淨利率（2016–2025，歸屬母公司淨利 ÷ 營收）", "%", None, "=AVERAGE(" + ",".join(f"{n}/{v}" for n, v in zip(NI10, REV10)) + ")", None, "Derived",
    "公司各年第四季新聞稿（SEC 8-K）：營收與淨利（2016 3,894／164 … 2025 6,708／373.9）", "stockanalysis FY2021–25 同；2017 含稅改一次性利益 42")

sec("I. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_r", "折現率／要求報酬", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%；各市場同一值）", "DCF 與 Sens")
inp("tp_g", "長期成長率", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "成熟倍數與 FY2032 收斂目標（情境 B）")
inp("tp_roe", "ROE（近三年平均，FY2023–25）", "%", None, "=AVERAGE(0.0941,0.0870,0.0866)", None, "Derived", "stockanalysis 比率頁：FY2023 9.41%、FY2024 8.70%、FY2025 8.66%", "單一來源；近 12 月 12.51%（同站）")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（工作單 2026-10-25 11:53 定 DCF／Comparable 各半；v0.1 為 0.6）", "—")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（情境 B FY2029 與十年平均的中點）", "%", None, 0, None, "Derived", "PL 頁情境 B FY2029、In 頁 nim10", "—")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（＝十年平均淨利率）", "%", None, "=nim10", None, "Derived", "In 頁 nim10", "—")
TP_KEYS += ["tp_r", "tp_g", "tp_roe", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]

URLNAME = {PR2Q: "公司 2Q26 新聞稿（SEC 8-K）", PR4Q: "公司 4Q25 新聞稿（SEC 8-K）", K10: "公司 FY2025 10-K", IDAY: "公司投資人日簡報（SEC 8-K）",
           FOOL4Q: "4Q25 法說逐字稿（Motley Fool）", FOOL1Q: "1Q26 法說逐字稿（Motley Fool）", MB2Q: "2Q26 法說摘要（MarketBeat）", SA: "stockanalysis.com"}
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

# ---------------- Q：實績與 FY2026 錨點（US$M） ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（百萬美元；EPS 為美元、稀釋；股數為百萬股）。藍字＝輸入（SEC 8-K 新聞稿）；黑字＝公式。1Q26＝1H26 − 2Q26（Derived）"
Q["A2"] = "FY2026 錨點＝1Q、2Q26 實績＋3Q26 公司指引（營收依各終端市場季增說法、毛利率與營業費用指引）＋4Q26 假設。3Q26 財報 2026-10-26 公布，本版尚未納入，屆時以實績取代 E 欄"
hdr(Q, 3, ["項目", "FY2025A", "1Q26（實績）", "2Q26（實績）", "3Q26（指引）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1（公司）", "來源2／說明"])
S1 = "FY2025：4Q25 新聞稿（SEC 8-K）；1H26、2Q26：2Q26 新聞稿（SEC 8-K）"
def qrow(r, label, vals, fmt="#,##0.0", tag="Verified", s1=S1, s2="—", ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"G{r}"] = ann or f"=SUM(C{r}:F{r})"; Q[f"G{r}"].number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, s1); Q.cell(r, 10, s2)
qrow(4, "營收", {"B": 6708.0, "C": "=3582.7-D4", "D": 1898.0, "E": "=SUM(E5:E8)", "F": "=SUM(F5:F8)"}, tag="Verified／Derived", s2="stockanalysis：FY2025 6,708；1H26 3,582.7、2Q26 1,898.0（公司）")
qrow(5, "  通訊（手機、平板）", {"B": "=B4*0.46", "C": "=C4*0.44", "D": "=D4*0.42", "E": "=D5*(1+g3_cm)", "F": "=E5*(1+q4_cm)"}, tag="Derived（占比 × 營收）／Interested-party", s2="占比：FY2025 46%、1Q26 44%、2Q26 42%（公司）")
qrow(6, "  運算（資料中心、基礎設施、PC、儲存）", {"B": "=B4*0.20", "C": "=C4*0.21", "D": "=D4*0.22", "E": "=D6*(1+g3_cp)", "F": "=E6*(1+q4_cp)"}, tag="Derived／Interested-party", s2="占比：20%、21%、22%（公司）")
qrow(7, "  車用、工業與其他", {"B": "=B4*0.19", "C": "=C4*0.21", "D": "=D4*0.22", "E": "=D7*(1+g3_au)", "F": "=E7*(1+q4_au)"}, tag="Derived／Interested-party", s2="占比：19%、21%、22%（公司）")
qrow(8, "  消費", {"B": "=B4*0.15", "C": "=C4*0.14", "D": "=D4*0.14", "E": "=D8*(1+g3_cs)", "F": "=E8*(1+q4_cs)"}, tag="Derived／Interested-party", s2="占比：15%、14%、14%（公司）")
qrow(9, "毛利", {"B": 938.6, "C": "=557.6-D9", "D": 318.6, "E": "=E4*q3_gm", "F": "=F4*q4_gm"}, tag="Verified／Interested-party（3Q）／Assumed（4Q）", s2="FY2025 含 4Q25 資產出售約 30（公司法說）")
qrow(10, "毛利率", {c: f"={c}9/{c}4" for c in "BCDEFG"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2="公司：FY2025 14.0%、1Q26 14.2%、2Q26 16.8%")
qrow(11, "營業費用（銷管＋研發）", {"B": "=304.5+166.7", "C": "=(175.9+81.6)-D11", "D": "=78.9+39.8", "E": "=q3_opx", "F": "=q4_opx"}, tag="Verified／Interested-party", s2="2Q26 含不動產處分利益 21（降低費用）")
qrow(12, "營業利益", {c: f"={c}9-{c}11" for c in "BCDEF"}, tag="Derived", s2="公司：FY2025 467.4、1H26 300.1、2Q26 199.9（核對見第 31 列）")
qrow(13, "營業外淨額（其他收益 − 利息費用）", {"B": "=52.7-75.4", "C": "=(34.0-34.7)-D13", "D": "=20.3-17.0", "E": "=q_nonop", "F": "=q_nonop"}, tag="Verified／Assumed", s2="其他收益含利息收入與匯兌（公司）")
qrow(14, "稅前利益", {c: f"={c}12+{c}13" for c in "BCDEF"}, tag="Derived", s1="—")
qrow(15, "所得稅", {"B": 68.5, "C": "=40.6-D15", "D": 28.3, "E": "=E14*q_tax", "F": "=F14*q_tax"}, tag="Verified／Interested-party", s2="FY2025 有效稅率 15.4%（含遞延稅資產一次性利益）")
qrow(16, "非控制權益", {"B": "=B14-B15-B17", "C": "=C14-C15-C17", "D": "=D14-D15-D17", "E": "=q_nci", "F": "=q_nci"}, tag="Derived", s1="稅後 − 歸屬母公司淨利")
qrow(17, "歸屬母公司淨利", {"B": 373.9, "C": "=257.1-D17", "D": 173.8, "E": "=E14-E15-E16", "F": "=F14-F15-F16"}, tag="Verified／Derived", s2="1Q26 83（公司）")
qrow(18, "稀釋股數（百萬股）", {"B": 248.5, "C": "=2*249.83-D18", "D": 249.978, "E": "=q_sh", "F": "=q_sh"}, ann="=AVERAGE(C18:F18)", tag="Verified／Derived", s2="1Q26＝2 × 1H26 加權 − 2Q26（Derived）")
qrow(19, "EPS（稀釋）", {c: f"={c}17/{c}18" for c in "BCDEF"}, fmt="0.00", ann="=G17/G18", tag="Derived", s1="淨利 ÷ 股數")
qrow(20, "  公司揭露 EPS", {"B": 1.50, "C": 0.33, "D": 0.70}, fmt="0.00", ann="", tag="Verified")
qrow(21, "折舊攤提", {"B": 642.0, "C": "=350.6-D21", "D": 180.0, "E": "=q_da3", "F": "=q_da4"}, tag="Verified／Assumed", s2="投資人日：LTM 1Q26 659")
qrow(22, "現金毛利（毛利＋折舊攤提）", {c: f"={c}9+{c}21" for c in "BCDEF"}, tag="Derived", s1="假設折舊攤提全數在銷貨成本（近似）")
Q.cell(24, 1, "FY2026 錨點的檢查（參考；校準未使用券商數字）").font = BOLD
checks = [("FY2026 營收年增", "=G4/B4-1", "0.0%"), ("FY2026 運算年增（公司 2026-02 說法：超過 20%）", "=G6/B6-1", "0.0%"),
          ("現金毛利增量 ÷ 營收增量（FY2025 扣除資產出售利益；In 頁 ft 的對照）", "=(G22-(B22-gain25))/(G4-B4)", "0.0%"),
          ("營業費用彈性（扣除 2Q26 處分利益；In 頁 phi 的對照）", "=((G11+opx_gain)/B11-1)/(G4/B4-1)", "0.00"),
          ("3Q26 營收 ÷ 指引中點 − 1", "=E4/g3_rev-1", "0.0%"), ("3Q26 EPS ÷ 指引中點 − 1", "=E19/g3_eps-1", "0.0%"),
          ("FY2025 營業利益 自算 − 公司公布 467.4（應為 0）", "=B12-467.4", "0.0"), ("1H26 營業利益 自算 − 公司公布 300.1（應為 0）", "=C12+D12-300.1", "0.0")]
for i, (lab, f, fm) in enumerate(checks, 25):
    Q.cell(i, 1, lab); Q[f"B{i}"] = f; Q[f"B{i}"].number_format = fm
for nm, rr in (("rev", 4), ("cm", 5), ("comp", 6), ("au", 7), ("cs", 8), ("gp", 9), ("opx", 11), ("oi", 12), ("nonop", 13), ("tax", 15), ("nci", 16), ("ni", 17), ("sh", 18), ("eps", 19), ("da", 21)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$G${rr}")
Q.column_dimensions["A"].width = 62
for c in "BCDEFG": Q.column_dimensions[c].width = 14
Q.column_dimensions["H"].width = 30; Q.column_dimensions["I"].width = 60; Q.column_dimensions["J"].width = 60
Q.freeze_panes = "B4"

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G", 2030: "H", 2031: "I", 2032: "J"}
FADE = (2030, 2031, 2032); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
BU = (2026, 2027, 2028, 2029); F3 = (2027, 2028, 2029)
DEFP = dict(scn="A", out="out_{y}", sh="sh_{y}", el="el", tri="tri", gpc26="g_pc26", gpc="g_pc", gau="g_au", gcs="g_cs",
            gcm27="g_cm27_{s}", gcmL="g_cmL_{s}", mf="1", ft="ft", capi="capi", capm26="cap_m26", life="life",
            azrev="az_rev{y}_{s}", azcm="az_cm", azfix="az_fix{y}", azinc="az_inc", azlife="az_life", azsvc="az_svc28",
            phi="phi", nonop="nonop{y}", tax="tax", fg="fade_g{s}", fm="fade_m{s}")
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
    um =lambda y: {2026: "", 2027: "", 2028: f"*um28_{s}", 2029: f"*um29_{s}"}[y]
    sec_("A. AI 加速器 → Amkor AI 先進封裝（D 欄＝FY2026 校準年）")
    row("units", "AI 加速器顆數（含情境係數）", "百萬顆", {FC[y]: f"=units_{y % 100}{um(y)}" + ("" if y == 2026 else f"*{P['mf']}") for y in BU}, "0.00")
    row("area", "封裝面積指數（2026＝1）", "x", {FC[y]: f"=ar_{y % 100}" for y in BU}, "0.000")
    row("out", "後段委外比例", "%", {"D": "=out_26", **{FC[y]: f"={p('out', y)}" for y in F3}}, "0.0%", note="FY2026 固定基準值（Sens 只動 FY2027 起）")
    row("sh", "Amkor 份額（占委外後段）", "%", {"D": "=sh_26", **{FC[y]: f"={p('sh', y)}" for y in F3}}, "0.0%")
    rai = r
    row("ai", "AI 先進封裝營收（2.5D、HDFO、AI 測試）", "US$M",
        {"C": f"=(a26_comp-a25_comp*(1+{P['gpc26']}))/({P['tri']}-(1+{P['gpc26']}))", "D": f"=C{rai}*{P['tri']}",
         **{FC[y]: f"={FC[y]}{rai - 4}*{FC[y]}{rai - 2}*{FC[y]}{rai - 1}*{FC[y]}{rai + 1}" for y in F3}},
        note="FY2025、FY2026：由運算營收與公司「約三倍」說法反推（Derived）；FY2027 起＝顆數 × 委外 × 份額 × 每顆金額")
    row("v", "每顆封測金額（Amkor 營收 ÷ 委外 ÷ 份額）", "US$／顆", {"D": f"=D{rai}/(D{rai - 4}*D{rai - 2}*D{rai - 1})",
        **{FC[y]: f"=$D{rai + 1}*{FC[y]}{rai - 3}^{P['el']}" for y in F3}}, "#,##0",
        note="FY2026 校準（Derived）；之後 × 封裝面積指數 ^ 彈性")
    row("pu", "每顆加速器的 Amkor AI 營收", "US$／顆", {FC[y]: f"={FC[y]}{rai}/{FC[y]}{rai - 4}" for y in BU}, "#,##0.0", note="含資料中心 CPU 的 HDFO 等連帶營收（拆分未揭露）")
    sec_("B. 終端市場與合計")
    row("cpn", "運算（非 AI：PC、儲存、基礎設施）", "US$M", {"C": f"=a25_comp-C{rai}", "D": f"=a26_comp-D{rai}", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['gpc']})" for y in F3}})
    row("comp", "運算合計", "US$M", {FC[y]: f"={FC[y]}{rai}+{FC[y]}{R['cpn']}" for y in FYS})
    row("cm", "通訊（Apple 占 FY2025 營收 29.8%）", "US$M", {"C": "=a25_cm", "D": "=a26_cm", "E": f"=D{r}*(1+{p('gcm27')})", "F": f"=E{r}*(1+{p('gcmL')})", "G": f"=F{r}*(1+{p('gcmL')})"})
    row("au", "車用、工業與其他", "US$M", {"C": "=a25_au", "D": "=a26_au", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['gau']})" for y in F3}})
    row("cs", "消費", "US$M", {"C": "=a25_cs", "D": "=a26_cs", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['gcs']})" for y in F3}})
    row("az", "亞利桑那廠營收（第一期）", "US$M", {"C": "=0", "D": "=0", "E": "=0", "F": f"={p('azrev', 2028)}", "G": f"={p('azrev', 2029)}"},
        note="第一期滿載約 US$1B（公司）；與既有客戶的亞洲產能可能部分替代（限制）")
    row("rev", "營收", "US$M", {FC[y]: f"={FC[y]}{R['comp']}+{FC[y]}{R['cm']}+{FC[y]}{R['au']}+{FC[y]}{R['cs']}+{FC[y]}{R['az']}" for y in FYS})
    row("g", "營收年增", "%", {FC[y]: f"={FC[y]}{R['rev']}/{FC[y - 1]}{R['rev']}-1" for y in FYS[1:]}, "0.0%")
    row("aish", "AI 先進封裝 ÷ 營收", "%", {FC[y]: f"={FC[y]}{rai}/{FC[y]}{R['rev']}" for y in FYS}, "0.0%")
    row("cmsh", "通訊 ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['cm']}/{FC[y]}{R['rev']}" for y in FYS}, "0.0%")
    sec_("C. 折舊、毛利與損益")
    row("mrev", "本業營收（不含亞利桑那）", "US$M", {FC[y]: f"={FC[y]}{R['rev']}-{FC[y]}{R['az']}" for y in FYS})
    row("capm", "本業資本支出（不含亞利桑那）", "US$M", {"D": f"={P['capm26']}", **{FC[y]: f"={P['capi']}*{FC[y]}{R['mrev']}" for y in F3}},
        note="FY2026 為 Assumed（總額 2.5–3.0B 扣除亞利桑那）；之後＝本業營收 × 資本支出率")
    row("dam", "本業折舊攤提", "US$M", {"C": "=a25_da", "D": "=a26_da", **{FC[y]: f"={FC[y - 1]}{r}*(1-1/{P['life']})+{FC[y - 1]}{R['capm']}/{P['life']}" for y in F3}},
        note="前一年 ×（1 − 1／年限）＋前一年資本支出 ÷ 年限")
    row("azda", "亞利桑那折舊（第一期，扣除補助與抵減）", "US$M", {"E": "=0", "F": f"=az_capex1*(1-{P['azinc']})/{P['azlife']}*{P['azsvc']}", "G": f"=az_capex1*(1-{P['azinc']})/{P['azlife']}"})
    row("da", "折舊攤提合計", "US$M", {"C": f"=C{R['dam']}", "D": f"=D{R['dam']}", **{FC[y]: f"={FC[y]}{R['dam']}+{FC[y]}{R['azda']}" for y in F3}})
    row("gpm", "本業毛利 ＝ FY2026 ＋ 流通率 × 本業營收增量 − 本業折舊增量", "US$M",
        {"C": "=a25_gp", "D": "=a26_gp", **{FC[y]: f"=$D{r}+{P['ft']}*({FC[y]}{R['mrev']}-$D{R['mrev']})-({FC[y]}{R['dam']}-$D{R['dam']})" for y in F3}},
        note="流通率為現金毛利口徑；折舊另計")
    row("azgp", "亞利桑那毛利 ＝ 營收 × 邊際貢獻率 − 固定現金成本 − 折舊", "US$M",
        {"E": f"=E{R['az']}*{P['azcm']}-{p('azfix', 2027)}", **{FC[y]: f"={FC[y]}{R['az']}*{P['azcm']}-{p('azfix', y)}-{FC[y]}{R['azda']}" for y in (2028, 2029)}},
        note="2027 為投產前啟動成本")
    row("gp", "毛利", "US$M", {"C": "=a25_gp", "D": "=a26_gp", **{FC[y]: f"={FC[y]}{R['gpm']}+{FC[y]}{R['azgp']}" for y in F3}})
    row("gm", "毛利率", "%", {FC[y]: f"={FC[y]}{R['gp']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("azimp", "亞利桑那對毛利率的影響（亞利桑那毛利 ÷ 營收）", "pt", {FC[y]: f"={FC[y]}{R['azgp']}/{FC[y]}{R['rev']}" for y in F3}, "0.00%",
        note="公司：啟動成本使營業利益率降約 1–2%（2027 起、2028 改善）")
    row("opx", "營業費用（銷管＋研發）", "US$M", {"C": "=a25_opx", "D": "=a26_opx", "E": f"=(D{r}+opx_gain)*(1+{P['phi']}*E{R['g']})",
        **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['phi']}*{FC[y]}{R['g']})" for y in (2028, 2029)}}, note="FY2027 基期加回 2Q26 不動產處分利益")
    row("oi", "營業利益", "US$M", {FC[y]: f"={FC[y]}{R['gp']}-{FC[y]}{R['opx']}" for y in FYS})
    row("oim", "營業利益率", "%", {FC[y]: f"={FC[y]}{R['oi']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("nonop", "營業外淨額", "US$M", {"C": "=a25_nonop", "D": "=a26_nonop", **{FC[y]: f"={p('nonop', y)}" for y in F3}})
    row("pti", "稅前利益", "US$M", {FC[y]: f"={FC[y]}{R['oi']}+{FC[y]}{R['nonop']}" for y in FYS})
    row("tax", "所得稅", "US$M", {"C": "=a25_tax", "D": "=a26_tax", **{FC[y]: f"={FC[y]}{R['pti']}*{P['tax']}" for y in F3}})
    row("nci", "非控制權益", "US$M", {"C": "=a25_nci", "D": "=a26_nci", **{FC[y]: "=nci" for y in F3}})
    row("ni", "歸屬母公司淨利", "US$M", {FC[y]: f"={FC[y]}{R['pti']}-{FC[y]}{R['tax']}-{FC[y]}{R['nci']}" for y in FYS})
    row("shr", "稀釋股數", "百萬股", {"C": "=a25_sh", "D": "=a26_sh", **{FC[y]: f"={FC[y - 1]}{r}*(1+sh_g)" for y in F3}}, "0.0")
    row("eps", "EPS（稀釋）", "US$", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['shr']}" for y in FYS}, "0.00")
    row("epsg", "EPS 年增", "%", {FC[y]: f"={FC[y]}{R['eps']}/{FC[y - 1]}{R['eps']}-1" for y in FYS[1:]}, "0.0%")
    row("nim", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
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
        for key, fm in (("g", "0.0%"), ("rev", "#,##0.0"), ("nim", "0.00%"), ("ni", "#,##0.0"), ("shr", "0.0"), ("eps", "0.00"), ("epsg", "0.0%")):
            ws[f"{c}{R[key]}"].number_format = fm; ws[f"{c}{R[key]}"].fill = FADEFILL
    ws.cell(R["rev"], 11, "FY2030 起為收斂法：營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m（FY2032 達到），非驅動因子推導")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — Amkor 損益（百萬美元；EPS 美元、稀釋；歸屬母公司）。FY2025A 實績、FY2026 錨點（1H26 實績＋3Q26 指引＋4Q26 假設；三情境相同），FY2027 起為模型。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A：顆數基準；通訊 FY2027 +3%、之後 +4%；亞利桑那 FY2028 營收 150、FY2029 500",
         "B": "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85；通訊 FY2027 0%、之後 +3%；亞利桑那 100／400",
         "C": "情境 C：CY2028–29 顆數 ×0.6；通訊 FY2027 −6%、之後 0%（Apple 分散供應）；亞利桑那 50／250"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
sr = 4; SUMR = {}
SUMK = (("rev", "營收", "US$M", "#,##0"), ("g", "  營收年增", "%", "0.0%"), ("ai", "  AI 先進封裝", "US$M", "#,##0"), ("comp", "  運算合計", "US$M", "#,##0"),
        ("cm", "  通訊", "US$M", "#,##0"), ("au", "  車用、工業與其他", "US$M", "#,##0"), ("az", "  亞利桑那廠", "US$M", "#,##0"),
        ("gm", "毛利率", "%", "0.00%"), ("oim", "營業利益率", "%", "0.00%"), ("ni", "歸屬母公司淨利", "US$M", "#,##0"), ("nim", "  淨利率", "%", "0.00%"),
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
IN.cell(INROW["fade_mA"], 5, f"=PL!G{SUMR[('A', 'nim')]}")
IN.cell(INROW["fade_mB"], 5, f"=(PL!G{SUMR[('B', 'nim')]}+nim10)/2")
for k in ("fade_mA", "fade_mB"): IN.cell(INROW[k], 5).font = Font(color="000000")
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型 − 錨點").font = BOLD; sr += 1
CHK = []
for lab, f in (("  三情境：營收差（US$M）", "+".join(f"ABS(D{PB[s]['rev']}-a26_rev)" for s in SC)),
               ("  三情境：EPS 差（US$）", "+".join(f"ABS(D{PB[s]['eps']}-a26_eps)" for s in SC)),
               ("  情境 A：AI＋非 AI 運算 − Q 頁運算（US$M）", f"D{PB['A']['ai']}+D{PB['A']['cpn']}-a26_comp")):
    PL.cell(sr, 1, lab); PL[f"D{sr}"] = "=" + f; PL[f"D{sr}"].number_format = "0.0000"; CHK.append(sr); sr += 1
assert sr < 60, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 10; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 13
PL.cell(2, 1, "灰底＝FY2030–FY2032 收斂法（上游只到 2029）：營收年增線性收斂到 fade_g、淨利率線性收斂到 fade_m（In 頁 I 節），非驅動因子推導")

# ---------------- Units：AI 加速器 → Amkor AI 先進封裝 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — AI 加速器顆數（tsmc 成品）→ Amkor AI 先進封裝營收（情境 A）。部署 GW＝AI IT 資本支出（markets/network）÷ 每 GW IT 資本支出（Tokenomics，經橋接表），只作 Recon"
A_ = PB["A"]; cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("units", "AI 加速器顆數（情境 A）", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['units']}", "0.00", "tsmc 成品 In 頁（台積電模型與 HBM、網通同一份）"),
        ("area", "封裝面積指數（2026＝1）", "x", lambda y: f"=PL!{FC[y]}{A_['area']}", "0.000", "加速器用 CoWoS 片數 ÷ 顆數（tsmc 成品 CoWoS 頁）"),
        ("out", "後段委外比例", "%", lambda y: f"=PL!{FC[y]}{A_['out']}", "0.0%", "Assumed"),
        ("sh", "Amkor 份額（占委外後段）", "%", lambda y: f"=PL!{FC[y]}{A_['sh']}", "0.0%", "Assumed"),
        ("v", "每顆封測金額", "US$／顆", lambda y: f"=PL!{FC[y]}{A_['v']}", "#,##0", "FY2026 校準；之後隨封裝面積"),
        ("ai", "Amkor AI 先進封裝營收", "US$M", lambda y: f"=PL!{FC[y]}{A_['ai']}", "#,##0", "顆數 × 委外 × 份額 × 每顆金額"),
        ("pu", "每顆加速器的 Amkor AI 營收", "US$／顆", lambda y: f"=PL!{FC[y]}{A_['pu']}", "#,##0.0", ""),
        ("gw", "部署 GW（AI IT 資本支出 ÷ 每 GW IT 資本支出）", "GW", lambda y: f"=itc_{y % 100}/" + ("capit_gb" if y == 2026 else "capit_vr"), "0.0", "2026 用 GB300、2027 起用 VR200"),
        ("tam", "OSAT 先進封裝市場（GW × US$0.5B）", "US$M", lambda y: f"={cyc(y)}{{gw}}*tam_gw*1000", "#,##0", "公司投資人日口徑（Interested-party）"),
        ("tsh", "Amkor AI 營收 ÷ OSAT 先進封裝市場", "%", lambda y: f"={cyc(y)}{{ai}}/{cyc(y)}{{tam}}", "0.0%", "隱含市占（對照用）")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    UR[key] = ur
    for y in BU:
        fx = f(y).replace("{gw}", str(UR.get("gw", 0))).replace("{ai}", str(UR.get("ai", 0))).replace("{tam}", str(UR.get("tam", 0)))
        UN[f"{cyc(y)}{ur}"] = fx; UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); ur += 1
ur += 1
UN.cell(ur, 1, "讀法：AI 先進封裝 FY2026 以公司「約三倍」說法與運算營收反推，不是公司揭露數字；委外比例與份額是 Assumed，只有相對變化影響 FY2027 起的營收。每顆封測金額隨封裝面積上升（彈性 0.5），所以 Rubin 世代單顆金額高於 Blackwell。").font = BOLD
UN_NOTE = ur
UN.column_dimensions["A"].width = 52; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 13
UN.column_dimensions["G"].width = 56

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（美元）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，年底落後 P/E）只作參考"
V["A2"] = "循環位置：Amkor 不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照廣達、NVIDIA、Arista；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
O = 51                     # 第 1 節起始列（第 0 節在上方）
V.cell(O, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, O + 1, ["項目", "A", "B", "C", "說明"])
V.cell(O + 2, 1, "現價（2026-10-23 收盤）"); V[f"B{O + 2}"] = "=px"; V[f"B{O + 2}"].number_format = "0.00"
for k, y, lab in ((3, 2026, "現價 ÷ FY2026 EPS（錨點）"), (4, 2027, "現價 ÷ FY2027 EPS（模型）"), (5, 2028, "現價 ÷ FY2028 EPS（模型）"), (6, 2029, "現價 ÷ FY2029 EPS（模型）")):
    V.cell(O + k, 1, lab)
    for j, s in enumerate(SC):
        c = "BCD"[j]; V[f"{c}{O + k}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{O + k}"].number_format = "0.0"
V.cell(O + 7, 1, "未來 12 月 EPS（模型 B）≈ 0.2 × FY2026＋0.8 × FY2027"); V[f"B{O + 7}"] = f"=0.2*PL!D{SUMR[('B', 'eps')]}+0.8*PL!E{SUMR[('B', 'eps')]}"; V[f"B{O + 7}"].number_format = "0.00"
V.cell(O + 8, 1, "  現價 ÷ 未來 12 月 EPS（模型 B）"); V[f"B{O + 8}"] = f"=px/B{O + 7}"; V[f"B{O + 8}"].number_format = "0.0"
V.cell(O + 9, 1, "近 12 月 P/E（stockanalysis，10-23）"); V[f"B{O + 9}"] = "=pe_ttm"; V[f"B{O + 9}"].number_format = "0.0"
V.cell(O + 10, 1, "前瞻 P/E（stockanalysis，10-23）"); V[f"B{O + 10}"] = "=fpe"; V[f"B{O + 10}"].number_format = "0.0"
V[f"E{O + 3}"] = "FY2026 三情境相同（錨點）"; V[f"E{O + 4}"] = "FY2027 三情境只差通訊與亞利桑那啟動成本"
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
V.cell(R30, 1, "3. 規則值隱含股價（美元；約 FY2027 年底，未折現）＝ FY2027 EPS × 規則倍數").font = BOLD
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
# 十年 P/E：companiesmarketcap 年底 P/E；自算＝digrin 12 月未調整股價 ÷ 公司稀釋 EPS
PE_CMC = [14.4, 8.61, 11.3, 24.3, 10.1, 8.85, 7.30, 21.9, 17.6, 26.1]
PX_DEC = [10.55, 10.05, 6.56, 13.00, 15.08, 24.79, 23.98, 33.27, 25.69, 39.48]
EPS_D = [0.69, 1.09, 0.53, 0.50, 1.40, 2.62, 3.11, 1.46, 1.43, 1.50]
PE_SELF = [f"={p_}/{e_}" for p_, e_ in zip(PX_DEC, EPS_D)]
T10 = vr + 2               # 十年 P/E（兩來源平均）範圍：PERCENTRANK 用
V.cell(T10, 1, "5. 十年 P/E（2016–2025 年底；兩來源平均，分位排名用；無負值或 >200 需剔除）").font = BOLD
hdr(V, T10 + 1, ["年底", "兩來源平均 P/E", "companiesmarketcap", "自算"])
RB0 = T10 + 16
for i, y in enumerate(range(2016, 2026)):
    rr = T10 + 2 + i
    V.cell(rr, 1, str(y))
    V.cell(rr, 2, f"=AVERAGE(C{rr}:D{rr})").number_format = "0.00"
    V.cell(rr, 3, f"=B{RB0 + 5 + i}").number_format = "0.00"
    V.cell(rr, 4, f"=C{RB0 + 5 + i}").number_format = "0.00"
PE10 = f"$B${T10 + 2}:$B${T10 + 11}"
V.cell(T10 + 12, 1, "年數"); V[f"B{T10 + 12}"] = f"=COUNT({PE10})"
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "companiesmarketcap（年底 P/E）", "vals": PE_CMC, "tag": "Interested-party（資料商；2016–2025 共 10 年）"},
    {"label": "自算（digrin 12 月未調整股價 ÷ 公司稀釋 EPS）", "vals": PE_SELF, "tag": "Derived（價格為月資料的 12 月值，非逐日年底收盤；EPS 取自公司各年第四季新聞稿）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位，兩個來源分別取分位再平均。Amkor 的落後 P/E 與盈餘反向：盈餘低谷年（2019、2023–2025）P/E 22–26x，高峰年（2021–2022）7–9x；套用在成長年的 EPS 時，中位數偏向高估。2017 EPS 含美國稅改一次性利益 0.17。")
assert V.cell(RB0 + 5, 1).value == "2016" and V.cell(RB0 + 14, 1).value == "2025", (V.cell(RB0 + 5, 1).value, V.cell(RB0 + 14, 1).value)
rule_rows = [rw for rw in range(RB0, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")

# 0. 目標價（Andy 2026-10-25 11:28 最終版：DCF 與 Comparable 加權；列標籤依工作單逐字；B／C／D＝A 順利／B 平平／C 不如預期）
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
# v0.2：非 A 股公司的同業不放 A 股（Andy 2026-10-25 11:53）——移除長電、通富，補矽格、南茂
PEERS = [("日月光投控（ASE Technology）", "ASX／3711.TW", 27.37, "2026-10-23"), ("力成（Powertech）", "6239.TW", 20.60, "2026-10-02"),
         ("京元電子（KYEC）", "2449.TW", 26.13, "2026-10-06"), ("矽格（Sigurd）", "6257.TW", 23.30, "2026-10-02"), ("南茂（ChipMOS）", "IMOS／8150.TW", 22.08, "2026-10-08")]
for i, (nm, tk, pe, dt) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左）"); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", "台美封測同業：日月光、力成、京元電、矽格、南茂（不放 A 股：長電、通富有結構性溢價）")
tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）", [f"=tp_wA*(2/12*B{EPSR[2027]}+10/12*B{EPSR[2028]})+tp_wB*(2/12*C{EPSR[2027]}+10/12*C{EPSR[2028]})+tp_wC*(2/12*D{EPSR[2027]}+10/12*D{EPSR[2028]})", None, None], "0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], "0.00", "同業中位數 × 前瞻 EPS")
V.cell(37, 1, "0.3 結論").font = BOLD
hdr(V, 38, ["項目", "數值", "", "", "說明"])
tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B24+(1-tp_w_dcf)*B35", None, None], "0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁")
tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], "0.00")
tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], "0.00", "stockanalysis：11 位分析師平均 76.40（中位數 73.50），2026-10-23 擷取；stocksguide 13 位 71.40")
tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-23 收盤）")
tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝現價 ÷ 前瞻 P/E（In 頁 cons_feps）")
tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
tset(46, "band", "價格帶（DCF 不如預期～順利）", ['=TEXT(D23,"0.00")&"～"&TEXT(B23,"0.00")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=PERCENTRANK({PE10},px/{c}{EPSR[2028]})" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 5 節兩來源平均）的分位；只作檢查，不進目標價")
V.cell(48, 1, "目標價方法")
V["B48"] = ("DCF：股東現金流＝GAAP 稀釋 EPS ×（1 − g／ROE），r＝10%、g＝4%、ROE 為近三年平均約 9%（再投資率約 45%），FY2027–32 明確期＋永續；"
            "Amkor 資本密集，但 FY2030 起只有收斂法、亞利桑那第二期支出未建，未改用淨利＋折舊－資本支出，現金流偏高估。"
            "Comparable：台美封測同業（日月光、力成、京元電、矽格、南茂）前瞻 P/E 中位數；不放 A 股（長電、通富約 35–50x，結構性溢價）；京元電以 AI 測試為主、倍數較高，部分價格為 10 月上旬。"
            "情境對應：A＝順利（顆數基準）、B＝平平（顆數 ×0.8／0.85）、C＝不如預期（顆數 ×0.6、通訊下滑）")
for r_ in (24, 35, 39):
    V[f"B{r_}"].fill = KEY
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；FY2026 恆等於錨點，AI 拆分與每顆金額在每個案例內重算）：對 FY2027、FY2028 營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("AI 部署 ×0.8", "FY2027 起顆數 ×0.8", dict(mf="0.8")), ("AI 部署 ×1.2", "×1.2", dict(mf="1.2")),
         ("委外比例 低", "2027 45%、2028–29 45%", dict(out="out_{y}_lo")), ("委外比例 高", "2027 65%、2028 70%、2029 75%", dict(out="out_{y}_hi")),
         ("Amkor 份額 低", "28%（日月光、矽品搶單）", dict(sh="sh_{y}_lo")), ("Amkor 份額 高", "2027 40%、2028–29 45%", dict(sh="sh_{y}_hi")),
         ("每顆金額彈性 0.3", "封裝面積帶動較少", dict(el="el_lo")), ("每顆金額彈性 1.0", "與面積等比例", dict(el="el_hi")),
         ("AI 三倍 → 2.5 倍", "FY2026 AI 拆分較大、基期較高", dict(tri="tri_lo")),
         ("通訊 情境 C 路徑", "FY2027 −6%、之後 0%", dict(gcm27="g_cm27_C", gcmL="g_cmL_C")),
         ("車用與工業 +4%", "基準 +10%", dict(gau="g_au_lo")), ("車用與工業 +14%", "", dict(gau="g_au_hi")),
         ("現金流通率 30%", "基準 40%", dict(ft="ft_lo")), ("現金流通率 50%", "", dict(ft="ft_hi")),
         ("本業資本支出率 18%", "折舊上升較快", dict(capi="capi_hi")), ("本業折舊年限 5 年", "基準 7 年", dict(life="life_lo")),
         ("亞利桑那固定成本 高", "2027 170、2028–29 150", dict(azfix="az_fix{y}_hi")), ("亞利桑那補助 20%", "基準 35%", dict(azinc="az_inc_lo")),
         ("營業費用彈性 1.0", "基準 0.5", dict(phi="phi_hi")), ("稅率 25%", "基準 20%", dict(tax="tax_hi")),
         ("營業外 −40", "每年再 −40", dict(nonop="nonop{y}_lo"))]
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
put_case(nn, "情境 B", "CY2028 顆數 ×0.8、CY2029 ×0.85；通訊較弱", PB["B"], "PL!")
put_case(nn + 1, "情境 C", "CY2028–29 顆數 ×0.6；通訊 −6%", PB["C"], "PL!")
SENS_LAST = nn + 1
assert nn + 3 < 40
SE.cell(nn + 3, 1, "讀法：顆數情境 B、C 的差異從 FY2028 起；FY2027 只差通訊與亞利桑那啟動成本。委外比例、份額、彈性只影響 AI 先進封裝（FY2026 約占營收一成），所以單項影響小於流通率與通訊；毛利率對流通率最敏感。")
SENS_NOTE = nn + 3
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
           ("DCF 權重 0.3", "Comparable 0.7", dict(w="0.3")), ("DCF 權重 0.7", "Comparable 0.3", dict(w="0.7"))]
for k, (nm, st, kw) in enumerate(TPCASES):
    tp_case(TPS0 + 2 + k, nm, st, **kw)
TPS_LAST = TPS0 + 1 + len(TPCASES)
assert TPS_LAST < 40, TPS_LAST
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 40
for c in "CDEFGHIJ": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司指引、公司長期目標、共識與市場口徑。只作事後對照，不為對上數字而調參數。共識 FY2027 只有單一來源且年度標示不明"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0", pt=False, gap=True):
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
RC.cell(rr_, 1, "A. FY2026 錨點 vs 公司指引與共識（錨點用 3Q 指引，非獨立檢驗）").font = BOLD; rr_ += 1
rec("3Q26 營收（各終端市場季增說法加總）vs 指引中點", "=g3_rev", lambda s: "=Q!E4", "2Q26 新聞稿：1,950–2,050", "共識 3Q26 2,040（MarketBeat）")
rec("3Q26 EPS vs 指引中點", "=g3_eps", lambda s: "=Q!E19", "2Q26 新聞稿：0.72–0.82", "", fmt="0.00")
rec("FY2026 營收 vs 共識", "=cons_rev26", pl("rev", 2026), "stockanalysis（11 位）", "tickflow 7.47B（7 位）")
rec("FY2026 EPS vs 共識（non-GAAP）", "=cons_eps26", pl("eps", 2026), "stockanalysis（11 位）", "模型為 GAAP", fmt="0.00")
rec("FY2027 EPS vs 共識（MarketBeat，年度標示不明）", "=cons_eps27", pl("eps", 2027), "MarketBeat（單一來源）", "", fmt="0.00")
rec("FY2026 運算年增 vs 公司「超過 20%」（2026-02）（百分點）", 0.20, lambda s: "=Q!G6/Q!B6-1", "4Q25 法說", "3Q26 指引運算近 +30% 季增，已高於年初說法", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "B. vs 公司長期目標（投資人日 2026-05-21；2Q26 財報前訂定）").font = BOLD; rr_ += 1
rec("FY2028 營收 vs 目標 9.0B（±0.5B）", "=t28_rev", pl("rev", 2028), "投資人日", "")
rec("FY2028 毛利率 vs 目標 17.5%（±1pt）（百分點）", "=t28_gm", pl("gm", 2028), "投資人日", "", fmt="0.0%", pt=True)
rec("FY2028 EPS vs 目標 2.50（±0.25）", "=t28_eps", pl("eps", 2028), "投資人日", "FY2026 錨點已高於 2028 目標", fmt="0.00")
rec("FY2029 營收 vs 2030 目標 11B（目標年晚一年）", "=t30_rev", pl("rev", 2029), "投資人日", "")
rec("FY2029 毛利率 vs 2030 目標 22%（百分點）", "=t30_gm", pl("gm", 2029), "投資人日", "", fmt="0.0%", pt=True)
rec("FY2029 EPS vs 2030 目標 5.0", "=t30_eps", pl("eps", 2029), "投資人日", "", fmt="0.00")
rec("運算 FY2025–29 年均成長 vs 投資人日 2025–30 約 24%（百分點）", 0.24, lambda s: f"=(PL!G{PB[s]['comp']}/PL!C{PB[s]['comp']})^(1/4)-1", "投資人日", "", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "C. AI 先進封裝的市場口徑").font = BOLD; rr_ += 1
rec("FY2026 Amkor AI 營收 ÷ OSAT 先進封裝市場（GW × 0.5B）", None, lambda s: "=Units!C13", "投資人日 0.5B／GW；部署 GW 讀網通模型與橋接表", "隱含市占", fmt="0.0%")
rec("FY2027 Amkor AI 營收 ÷ OSAT 先進封裝市場", None, lambda s: "=Units!D13", "同上", "情境 A", fmt="0.0%")
rec("台積電 CoWoS 2027 需求 ÷ 可用產能（tsmc 成品）", None, lambda s: "=cow_util27", "tsmc/dist CoWoS 頁", "接近 100% → 後段委外需求上升（本模型委外比例 Assumed 上升的依據之一）", fmt="0%")
rec("FY2026 每顆加速器 Amkor AI 營收（US$）", None, pl("pu", 2026), "—", "對照：台積電模型每顆台積電內容 1.1–6.0 千美元（邏輯＋CoWoS）", fmt="#,##0.0")
RC.cell(rr_, 1, "D. 亞利桑那").font = BOLD; rr_ += 1
for y in (2027, 2028, 2029):
    rec(f"FY{y} 亞利桑那對毛利率的影響 vs 公司對營業利益率 −1～−2%（取 −1.5%）（百分點）", -0.015, pl("azimp", y), "1Q26 法說：2027 起 1–2%、2028 改善；2029 損益兩平（投資人日）", "", fmt="0.00%", pt=True)
RC.cell(rr_, 1, "E. 客戶集中").font = BOLD; rr_ += 1
rec("FY2027 通訊 ÷ 營收 vs FY2025 Apple 占營收 29.8%（百分點）", "=apple", pl("cmsh", 2027), "10-K", "通訊含 Android 與 Qualcomm；只看方向", fmt="0.0%", pt=True)
RC.column_dimensions["A"].width = 72
for c in "BCDEFGH": RC.column_dimensions[c].width = 11
RC.column_dimensions["I"].width = 40; RC.column_dimensions["J"].width = 60
RECON_LAST = rr_ - 1

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（SEC 申報為公司原文；法說內容經逐字稿與摘要；券商與資料商皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("公司：2Q26 新聞稿（SEC 8-K，2026-07-27；1H26、2Q26 損益、終端市場占比、3Q26 指引、資本支出）", PR2Q),
        ("公司：4Q25 新聞稿（SEC 8-K，2026-02-09；FY2025 損益、終端市場占比、先進／主流產品）", PR4Q),
        ("公司：FY2025 10-K（Apple 29.8%、Qualcomm 11.1%、前十大 72%、2.5D 與 HDFO 說明）", K10),
        ("公司：投資人日簡報（SEC 8-K，2026-05-21；2028、2030 目標、終端市場 CAGR、亞利桑那、0.5B／GW）", IDAY),
        ("公司：4Q25 法說逐字稿（Motley Fool；2.5D 與 HDFO 近三倍、亞利桑那第一期約一半、資產出售 30）", FOOL4Q),
        ("公司：1Q26 法說逐字稿（Motley Fool；AI 封裝三倍、亞利桑那稀釋 1–2%、CHIPS 400M＋35% 抵減）", FOOL1Q),
        ("公司（摘要）：MarketBeat 2Q26 法說重點（3Q 各終端市場、營業費用 140、NVIDIA、台積電合作）", MB2Q),
        ("公司：亞利桑那第二期、投資擴大至約 US$12B（2026-09-08）", "https://ir.amkor.com/node/24686"),
        ("公司：亞利桑那 US$7B 計畫（Business Wire 2025-10-06）", "https://www.businesswire.com/news/home/20251006108489/en"),
        ("公司：與 NVIDIA 策略合作（Business Wire 2026-07-23）；Reuters 經 The Star：US$1.5B", "https://www.thestar.com.my/tech/tech-news/2026/07/24/nvidia-amkor-strike-15-billion-chip-packaging-deal"),
        ("公司：與台積電長期合作（Business Wire 2026-06-16）", "https://www.businesswire.com/news/home/20260616574153/en/"),
        ("媒體：Manufacturing Dive（亞利桑那 2027 年中完工、2028 年初投產、CHIPS 407M）", "https://www.manufacturingdive.com/news/amkor-arizona-7-billion-semiconductor-tsmc-apple-nvidia/802297/"),
        ("媒體：Benzinga 經 Webull（2Q26 後股價 −12%、3Q 指引前共識 2.09B）", "https://www.webull.com/news/15305145166619648"),
        ("公司：歷年第四季新聞稿（EPS 2016–2022）", "https://www.sec.gov/Archives/edgar/data/1047127/000104712723000004/amkr123122erex-991.htm"),
        ("公司：4Q19、4Q20、4Q17 新聞稿（EPS 2016–2020）", "https://www.sec.gov/Archives/edgar/data/1047127/000104712720000003/amkr123119erex-991.htm"),
        ("資料商：stockanalysis（現價、近 12 月與前瞻 P/E、共識 FY2026、FY2021–25 財務）", SA),
        ("資料商：companiesmarketcap 年底 P/E", "https://companiesmarketcap.com/amkor-technology/pe-ratio/"),
        ("資料商：digrin 月股價（12 月未調整值）", "https://www.digrin.com/stocks/detail/AMKR/price"),
        ("資料商：MarketBeat 盈餘頁（共識 3Q26、FY2027）", "https://www.marketbeat.com/stocks/NASDAQ/AMKR/earnings/"),
        ("資料商：aaii（10-22 收盤 39.05、10 月 −26%、10-22 一家降評）", "https://www.aaii.com/investingideas/article/541165-why-amkor-technology-inc8217s-amkr-stock-is-down-804"),
        ("資料商：stocksguide（10-23 價格 39.30、目標價均值 71.40）", "https://stocksguide.com/en/forecast/Amkor-Technology-Inc-US0316521006"),
        ("市場層：tsmc、markets/network、markets/chip_bridge 成品（本 repo，唯讀）", f"{rel(TSMC_F)}；{rel(NET_F)}；{rel(BR_F)}"),
        ("待補（人工）：3Q26 財報（2026-10-26）；AI 先進封裝與 2.5D 營收金額；委外比例與份額的量化；亞利桑那支出時程；可轉債轉換價；共識 FY2027（付費）；前瞻 P/E 十年歷史（Bloomberg）", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 100; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("Amkor Technology（AMKR）收支與估值模型 v0.2.1（AI 半導體 Project：封測次產業）　2026-10-25", True),
 ("版本：v0.2.1（2026-10-25：補上檔空間級距、投資建議、重點觀察，數值不變）；v0.2（2026-10-25：tp_w_dcf 改 0.5；Comparable 同業移除 A 股；v0.1 初版）。Amkor 為全球第二大封測廠，承接台積電 CoWoS 後段（oS、測試）與 2.5D／HDFO 委外；亞利桑那廠為美國第一座量產先進封裝廠。日月光模型待 10-29 法說後建立，本版未讀。", False),
 ("模型命題：Amkor 的 AI 先進封裝營收＝AI 加速器顆數（tsmc 成品）× 後段委外比例 × Amkor 份額 × 每顆封測金額（隨封裝面積上升），FY2026 約占營收一成；其餘九成是通訊（Apple 占 29.8%）、車用與工業、消費與非 AI 運算。毛利率只有 14–19%，資本支出 2026 年約 US$2.75B（營收的三成以上），亞利桑那的啟動成本與折舊是 2027–29 毛利率的關鍵。", False),
 ("核心問題：AI 封裝的高成長能不能抵過通訊（Apple）停滯與亞利桑那折舊，讓 EPS 在 FY2027–28 持續成長；約 14 倍前瞻 P/E 隱含什麼路徑。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029（三年）。金額：百萬美元；EPS：美元（稀釋）。口徑：歸屬母公司淨利（GAAP）。3Q26 財報 2026-10-26 公布，本版未納入。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("AI 先進封裝：tsmc 成品的加速器顆數 × 情境係數（2028–29）× 後段委外比例（Assumed）× Amkor 份額（Assumed）× 每顆封測金額（FY2026 校準，之後 × 封裝面積指數 ^ 0.5；面積指數＝tsmc 成品 CoWoS 片數 ÷ 顆數）。FY2026 的 AI 金額由運算營收與公司「AI 封裝約三倍」說法反推。", False),
 ("其他終端：通訊依情境（A +3%／B 0%／C −6% 於 FY2027），車用與工業 +10%、消費 +3%（公司投資人日 CAGR），非 AI 運算 +3%。亞利桑那第一期營收 FY2028–29 依情境（Assumed）。", False),
 ("毛利：本業毛利＝FY2026 錨點＋現金流通率 40%（Assumed）× 營收增量 − 本業折舊增量（資本支出率 14%、年限 7 年）；亞利桑那毛利＝營收 × 55% − 固定現金成本 − 折舊（第一期 US$3.5B × (1 − 補助 35%) ÷ 12 年）。營業費用彈性 0.5；營業外、稅率 20%、股數持平。", False),
 ("校準：FY2026 ＝ 1H26 實績＋3Q26 公司指引（終端市場季增說法、毛利率 19%、營業費用 140）＋4Q26 假設。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，三情境權重 0.3／0.5／0.2）＋0.5 × Comparable 目標價（日月光、力成、京元電、矽格、南茂前瞻 P/E 中位數 × 12 個月後前瞻 EPS；不放 A 股）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解。十年 P/E 規則值 × FY2027 EPS 保留為參考（不是目標價）。情境對應：A＝順利、B＝平平、C＝不如預期（原為顆數情境）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m，FY2032 達到；PL 頁灰底）。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司數字取自 SEC 8-K 與 10-K；法說內容經逐字稿與摘要；券商與資料商為 Interested-party。", False),
 ("", False),
 ("已知限制：(1) AI 先進封裝金額公司未揭露，FY2026 由「約三倍」說法反推，委外比例與份額是 Assumed；(2) 每顆封測金額只隨封裝面積變動，未分 2.5D、HDFO、測試；(3) 亞利桑那營收可能部分替代亞洲既有產能（Apple 晶片在美國封裝），本模型當作增量；(4) 亞利桑那支出時程、補助的會計處理未找到；(5) 本業毛利用單一現金流通率，未建稼動率；(6) 十年 P/E 與盈餘反向（低谷年高倍數），規則值偏向高估；(7) 共識 FY2027 只有單一來源且年度標示不明；(8) 3Q26 財報（10-26）後更新錨點。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ---------------- v0.2.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 14:00、14:17；既有數值與儲存格位置不變） ----------------
_UPW = [r_ for r_ in range(1, 60) if V.cell(r_, 1).value == "上檔空間：加權目標價"]
assert _UPW == [42], _UPW
GR_ROW, RC_ROW = 49, 50   # 0.3 結論後的空白預留列（不移動其他列）
for _r in (GR_ROW, RC_ROW):
    assert all(V.cell(_r, _c).value is None for _c in range(1, V.max_column + 1)), _r
V.cell(GR_ROW, 1, "上檔空間級距")
V.cell(GR_ROW, 2, '=IF(B42>0.5,"大於 50%",IF(B42>0.25,"25%–50%",IF(B42>=0,"0%–25%",IF(B42>=-0.25,"0%～−25%",IF(B42>=-0.5,"−25%～−50%","小於 −50%")))))').number_format = "@"
V.cell(GR_ROW, 5, "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）")
V.cell(RC_ROW, 1, "投資建議")
V.cell(RC_ROW, 2, '=IF(B42>0.5,"強力買進",IF(B42>0.25,"買進",IF(B42>=-0.25,"中立",IF(B42>=-0.5,"放空","強力放空"))))').number_format = "@"
V.cell(RC_ROW, 5, "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
OBS = [
       "目標價 US$55.7（DCF 36.0、Comparable 75.3 各半），比現價 39.30（2026-10-23）高 41.7%，級距 25%–50%；Comparable 單獨為 +91.7%。主因是倍數：本模型 12 個月後前瞻 EPS 3.23 比共識前瞻 EPS 2.80 高 15.3%（期間差約一年），台美封測同業（日月光、力成、京元電、矽格、南茂）前瞻 P/E 中位數 23.3 倍，現價只隱含 14.0 倍、共識目標價 76.40 隱含 27.3 倍；DCF 36.0 低於現價（−8.4%），因近三年 ROE 8.9% 低於折現率 10%（可發放比例只 55%）。",
       "最敏感的是現金流通率、委外比例與份額：現金流通率 30% 時 FY2028 EPS −16.9%、委外比例低 −16.2%、Amkor 份額低（28%）−15.7%、AI 部署 ×0.8 −12.8%；情境 C FY2028 EPS 2.32，比情境 A 的 4.00 低 42.0%。",
       "追蹤 3Q26 財報（2026-10-26）：本版 3Q26 用公司指引（毛利率 19%），AI 先進封裝金額公司未揭露、由「約三倍」說法反推；亞利桑那廠固定成本與折舊從 FY2028 起影響毛利率（固定成本高案例 FY2028 EPS −4.0%）。"
      ]   # 重點觀察（與 README.md 相同文字）
_n0 = RD.max_row
RD.cell(_n0 + 1, 1, "v0.2.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（Val 頁 0.3 結論加「上檔空間級距」「投資建議」兩列、README 頁加重點觀察三點；數值不變，既有儲存格位置不變）。")
RD.cell(_n0 + 2, 1, f'="投資建議："&Val!B{RC_ROW}&"（依上檔空間固定規則）"').font = BOLD
for _i, _t in enumerate(OBS, 1): RD.cell(_n0 + 2 + _i, 1, f"重點觀察 {_i}：" + _t)

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(sys.argv[1])
if os.environ.get("AMKR_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE, "UN_NOTE": UN_NOTE, "T": T, "T10": T10, "TPS0": TPS0, "TPS_LAST": TPS_LAST, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0},
              open(os.environ["AMKR_ROWMAP"], "w"), ensure_ascii=False, indent=1)

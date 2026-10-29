# 健策精密（3653.TW）收支與估值模型 v0.1.1（AI 半導體 Project：電源與散熱次產業；均熱片（IHS／lid）、加強環、液冷冷板元件）
# 用法：python3 jentech/build_jentech.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；每顆均熱片金額與液冷市場以美元計、以匯率 fx 換算；EPS：新台幣。
# 口徑：歸屬母公司淨利（IFRS）。樣板：avc/（奇鋐：散熱同次產業）；Val 第 0 節結構同 gce/（新版 v0.1 寫法：目標價在 Val 最前面）。
# 均熱片：加速器顆數（markets/network Units 頁，情境 A 基準）× tsmc 同字母情境顆數係數 × 每顆均熱片金額（Assumed）× 健策份額（FY2026 校準）。
# 液冷冷板元件：讀 avc/dist PL 頁同字母情境的「AI 液冷市場（奇鋐可做部分）」× 健策份額（Assumed）；奇鋐份額同頁讀入，只作份額合計檢查。
# 情境與行業一致（Andy 2026-10-25 13:13）：A／B／C 讀 tsmc 同字母顆數係數（＝markets/memory SCEN_UNITS）與 avc 同字母液冷市場；
# 可拆式上蓋（TR lid）滲透屬技術替代，與份額、液冷份額同方向：A 取高值、B 基準、C 取低值。
# FY2030–FY2032 為收斂法（上游只到 2029）；目標價＝DCF 與 Comparable 加權（Val 第 0 節）。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（液冷行業變數 v0.6.1 起單一來源）
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
NET_F = only("markets/network/dist/*.xlsx"); TSMC_F = only("tsmc/dist/*.xlsx"); AVC_F = only("avc/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
# 加速器顆數（百萬顆；情境 A 基準）
UNITS = {}
GRP = (("bw", "NVIDIA Blackwell（B200／B300）"), ("rb", "NVIDIA Rubin"), ("ru", "NVIDIA Rubin Ultra"), ("hp", "NVIDIA Hopper（H200 等）"),
       ("amd", "AMD Instinct"), ("tpu", "Google TPU"), ("trn", "AWS Trainium"), ("oth", "其他 ASIC（MTIA、Maia、Ascend 等）"))
for r in openpyxl.load_workbook(NET_F, data_only=True)["Units"].iter_rows(min_row=4):
    for k, lab in GRP:
        if r[0].value == lab:
            UNITS[k] = ([r[1 + j].value for j in range(4)], f"{rel(NET_F)}｜Units!B{r[0].row}:E{r[0].row}")
assert len(UNITS) == len(GRP) and all(v is not None for k in UNITS for v in UNITS[k][0]), UNITS.keys()
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and str(k).startswith("um"):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
assert all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()
for s_ in SC:   # 情境與行業一致：tsmc 顆數係數必須等於 markets/memory 的 SCEN_UNITS（同一字母）
    assert abs(TS[f"um28_{s_}"][0] - MEM.SCEN_UNITS[s_][0]) < 1e-9 and abs(TS[f"um29_{s_}"][0] - MEM.SCEN_UNITS[s_][1]) < 1e-9, s_
# 奇鋐模型：同字母情境的 AI 液冷市場（美元）與奇鋐有效份額（依 PL 頁情境區塊出現順序 A、B、C）
AVC = {"lmk": [], "sh": []}
for r in openpyxl.load_workbook(AVC_F, data_only=True)["PL"].iter_rows(min_row=60):
    if r[0].value == "AI 液冷市場（奇鋐可做部分；美元）":
        AVC["lmk"].append(([r[3 + j].value for j in range(4)], f"{rel(AVC_F)}｜PL!D{r[0].row}:G{r[0].row}"))
    if r[0].value == "奇鋐有效份額":
        AVC["sh"].append(([r[3 + j].value for j in range(4)], f"{rel(AVC_F)}｜PL!D{r[0].row}:G{r[0].row}"))
assert len(AVC["lmk"]) == 3 and len(AVC["sh"]) == 3, {k: len(v) for k, v in AVC.items()}
assert all(v is not None for k in AVC for x in AVC[k] for v in x[0]), AVC
# AI 液冷市場只有一個值（Siliconomics Rack 頁）：奇鋐成品三情境必須與本庫相同
for _s, (_v, _src) in zip(SC, AVC["lmk"]):
    assert all(abs(a - b) < 1e-6 for a, b in zip(_v, SIL.NAMES[f"SIF_LiqMkt_{_s}"])), ("液冷市場與 Siliconomics 不一致", _s, _v, SIL.NAMES[f"SIF_LiqMkt_{_s}"])

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B）；每顆均熱片金額為美元、液冷市場為十億美元；比例：%。每個輸入盡量附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文＝健策財報、月營收與自結損益公告（經資料商與媒體轉述）；公司的產品與出貨說法為 Interested-party；券商與媒體估計皆為 Interested-party"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
PKT = "https://www.pocket.tw/school/articles/6275/"
PKT2 = "https://www.pocket.tw/school/articles/6817/"
MW = "https://www.moneyweekly.com.tw/_Article?AID=241110"
CT0807 = "https://www.ctee.com.tw/news/20260807700227-439901"
UA1005 = "https://uanalyze.com.tw/articles/8953956478"
DW = "https://money.udn.com/money/amp/story/5607/9786948"
UDN1008 = "https://money.udn.com/money/amp/story/5607/9802684"
TN9 = "https://finance.technews.tw/2026/10/08/jentech-3653-202609-financial-report/"
TN7 = "https://finance.technews.tw/2026/08/07/jentech-3653-202607-financial-report/"
FS05 = "https://news.cnyes.com/news/id/6468499"
SA = "https://stockanalysis.com/quote/tpe/3653/"
SAQ = "https://stockanalysis.com/quote/tpe/3653/financials/?p=quarterly"
SAR = "https://stockanalysis.com/quote/tpe/3653/financials/ratios/"
SAF = "https://stockanalysis.com/quote/tpe/3653/forecast/"
HISQ = "https://histock.tw/stock/3653/%E6%90%8D%E7%9B%8A%E8%A1%A8"
HISE = "https://histock.tw/stock/3653/%E6%AF%8F%E8%82%A1%E7%9B%88%E9%A4%98"
GI = "https://goodinfo.tw/tw/StockBzPerformance.asp?STOCK_ID=3653"
WEARN = "https://stock.wearn.com/financial.asp?kind=3653"
SD = "https://statementdog.com/analysis/3653"

sec("A. AI 加速器顆數（讀 markets/network Units 頁，情境 A 基準；B／C 依 tsmc 同字母顆數係數縮放）")
for k, lab in GRP:
    for j, y in enumerate((26, 27, 28, 29)):
        inp(f"u_{k}{y}", f"{lab} 顆數 20{y}（網通模型）", "百萬顆", None, UNITS[k][0][j], None, "Derived", UNITS[k][1],
            "唯讀 markets/network（顆數取自 markets/memory；2029 為成長率）" if j == 0 else "同上")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝順利（基準）；B＝平平；C＝不如預期；與台積電、奇鋐模型相同（＝markets/memory SCEN_UNITS，建置時核對）")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")
inp("fx", "匯率（新台幣 ÷ 美元）", "NT$/US$", *SIL.vals("fx_twd"), "Derived", "與奇鋐、台達電、金像電模型相同（avc/dist In 頁 fx＝32.0）", "鉅亨 2026-05-22 頁面 32.508（只作參考）")

sec("B. 每顆加速器均熱片／上蓋／加強環金額（美元；可拆式上蓋滲透為技術替代：A 取高值＝快、C 取低值＝慢）")
inp("asp_bw", "Blackwell（B200／B300）每顆均熱片＋加強環", "US$/顆", 30, 35, 40, "Interested-party", MW + "：Blackwell Ultra 均熱片約 US$30（理財周刊 2026-08-05）", "加強環另計 US$5（Assumed）；第二來源：未找到")
inp("asp_hp", "Hopper 每顆均熱片", "US$/顆", 10, 15, 20, "Assumed", "未找到（Hopper 單價低於 Blackwell）", "2027 起顆數為 0，只影響 FY2026 校準")
for y, v in ((26, (None, 100, None)), (27, (110, 150, 200)), (28, (120, 180, 220)), (29, (120, 180, 220))):
    inp(f"asp_vr_{y}", f"Rubin（VR200）每顆均熱片／上蓋 20{y}（單片式 → 可拆式三件式上蓋）", "US$/顆", *v, "Interested-party" if y == 26 else "Assumed",
        MW + "：Rubin 均熱片約 US$100、較前代高約六成（理財周刊 2026-08-05）" if y == 26 else
        (UDN1008 + "：本國投顧假設 2027 年 Rubin／Rubin Ultra 800–900 萬顆採可拆式上蓋、ASP 至少 US$200（經濟日報 2026-10-08）" if y == 27 else "同上"),
        "可拆式上蓋 2027 年第一季量產（大和、本國投顧，Interested-party）；基準＝約一半出貨採可拆式（Assumed）" if y == 27 else "—")
inp("asp_ru", "Rubin Ultra 每顆可拆式上蓋", "US$/顆", 150, 200, 250, "Interested-party", MW + "：Rubin Ultra 約 US$200（理財周刊 2026-08-05）", "大和 2026-10-01：Vera Rubin、Rubin Ultra 自 2027 年第一季採用（經濟日報）")
for y, v in ((26, (None, 40, None)), (27, (45, 60, 80)), (28, (50, 80, 110)), (29, (50, 80, 110))):
    inp(f"asp_amd_{y}", f"AMD Instinct 每顆均熱片／上蓋 20{y}", "US$/顆", *v, "Assumed",
        PKT2 + "：健策為 AMD、Intel 長期 IHS 核心供應商（口袋學堂 2025-12-15）" if y == 26 else (UDN1008 + "：AMD GPU 與 ASIC 採可拆式上蓋、單價更高（本國投顧）" if y == 27 else "同上"),
        "單價：未找到" if y == 26 else "—")
for y, v in ((26, (None, 30, None)), (27, (30, 35, 45)), (28, (30, 40, 55)), (29, (30, 40, 55))):
    inp(f"asp_asic_{y}", f"ASIC（TPU、Trainium、其他）每顆均熱片＋加強環 20{y}", "US$/顆", *v, "Assumed",
        CT0807 + "：多項 ASIC 產品採用加強環（stiffener），ASP 較高（工商時報 2026-08-07）" if y == 26 else "同上",
        "AWS Trainium3 為客戶（口袋學堂 2026-08-26）；單價：未找到" if y == 26 else "—")

sec("C. 健策份額（FY2026 以 AI 均熱片營收校準）與液冷冷板元件份額")
inp("k_non", "非 NVIDIA（AMD、ASIC）份額 ÷ NVIDIA 份額", "x", 0.25, 0.40, 0.55, "Assumed",
    PKT2 + "：健策為全球前兩大 IHS 供應商（口袋學堂，Interested-party）；份額數字：未找到", "ASIC 客戶多、供應商分散（Assumed）；影響 FY2026 校準出的 NVIDIA 份額")
for y, v in ((27, (0.92, 0.97, 1.00)), (28, (0.85, 0.93, 1.00)), (29, (0.80, 0.90, 0.98))):
    inp(f"rel_{y}", f"健策均熱片份額 20{y} ÷ 2026（相對值；情境 A 取高、B 基準、C 取低）", "x", *v, "Assumed",
        "可拆式上蓋（TR lid）結構複雜、單價數倍，新供應商可能切入（Assumed）；4Q26 小量送樣（優分析 2026-10-05）" if y == 27 else "同上",
        "對手份額：未找到" if y == 27 else "—")
for y in (27, 28, 29):
    inp(f"lcs_{y}", *SIL.row9(f"lc_jtsh_{y}")[1:])   # Siliconomics M19
for s, (vals, src) in zip(SC, AVC["lmk"]):
    for j, y in enumerate((26, 27, 28, 29)):
        inp(f"lmk{s}_{y}", f"情境 {s}：AI 液冷市場 20{y}（奇鋐模型，同字母情境）", "US$B", None, vals[j], None, "Derived", src,
            "唯讀 avc/dist（機架內冷板＋分歧管；不含 CDU）；雙鴻模型（auras/）未在 main，依工作單讀奇鋐模型" if (j == 0 and s == "A") else "同上")
for s, (vals, src) in zip(SC, AVC["sh"]):
    for j, y in enumerate((26, 27, 28, 29)):
        inp(f"avsh{s}_{y}", f"情境 {s}：奇鋐液冷有效份額 20{y}（奇鋐模型；份額合計檢查用）", "%", None, vals[j], None, "Derived", src, "只作 Recon 份額合計檢查" if (j == 0 and s == "A") else "同上")

sec("D. 營收結構錨點（公司與媒體揭露比例 × 營收）")
inp("th_25", "2025 年散熱產品（均熱片、ILM、車用水冷模組等）÷ 營收", "%", None, 0.726, None, "Interested-party", PKT + "：2025 年散熱產品占營收 72.6%（口袋學堂 2026-08-26）", "第二來源：未找到")
inp("th_26", "2026 年散熱產品 ÷ 營收", "%", 0.66, 0.75, 0.80, "Interested-party", PKT + "：2Q26 散熱產品約 77%", MW + "：1H26 散熱 65.54%、導線架 13.64%（同文另稱約七成，口徑不一）；本國投顧：均熱片維持八成以上（經濟日報 2026-10-08）")
inp("ai_25", "2025 年 AI 加速器均熱片（GPU、ASIC）÷ 營收", "%", 0.22, 0.30, 0.38, "Assumed", "未找到（公司未拆 AI 與伺服器 CPU、PC 均熱片）", "只影響 FY2025 拆分顯示")
inp("ai_26", "2026 年 AI 加速器均熱片 ÷ 營收", "%", 0.32, 0.40, 0.48, "Assumed",
    "未找到；推論：VR 均熱片 2026 年約 200 萬片 × US$100（工商時報 2026-08-07）＋Blackwell 約 500 萬顆 × US$35 ≈ US$0.37B ≈ 營收的 36%", "其餘散熱＝伺服器與 PC CPU 均熱片、交換器加強環、ILM、車用水冷；影響 FY2026 份額校準")
for y, v in ((27, (0.05, 0.12, 0.20)), (28, (0.03, 0.10, 0.15)), (29, (0.02, 0.08, 0.12))):
    inp(f"g_oth_{y}", f"其他散熱營收年增 20{y}（CPU 均熱片、交換器加強環、CPO、ILM、車用水冷）", "%", *v, "Assumed",
        CT0807 + "：Spectrum-X、Quantum-X 交換器加強環 2027 年貢獻（工商時報 2026-08-07）" if y == 27 else "同上",
        UDN1008 + "：CPO 均熱片、Intel CPU 載具與插槽成長（本國投顧）" if y == 27 else "—")
for y in (27, 28, 29):
    inp(f"g_lf_{y}", f"導線架、精密沖壓與其他營收年增 20{y}", "%", -0.05, 0.03, 0.08, "Assumed", "800V 車用平台帶動導線架需求（口袋學堂，Interested-party）" if y == 27 else "同上", "—")

sec("E. 毛利率")
inp("d_ai", "AI 加速器均熱片毛利率高於其他營收的幅度", "pt", 0.06, 0.12, 0.18, "Assumed",
    "產品別毛利率：未找到；毛利率 FY2025 41.6% → 1Q26 41.8% → 2Q26 46.8%（stockanalysis），同期 VR 均熱片出貨", "大和：3Q26 毛利率約 51%；券商：2027–28 年毛利率逾 50%（Interested-party）")

sec("F. FY2026 錨點（1H26 財報＋3Q26 月營收實績＋假設；3Q26 財報預計 2026-11-09，截至 2026-10-25 尚未公布）")
inp("q3_rev", "3Q26 營收（7–9 月月營收合計，實績）", "NT$B", None, "=3.213+3.226+3.054", None, "Verified",
    TN9 + "：9 月 30.54 億、前 9 月 220.73 億（− 1H26 125.80 ＝ 94.93）", TN7 + "：7 月 32.13 億；8 月 32.26 億（財報狗）")
inp("q_gm", "3Q、4Q26 毛利率", "%", 0.47, 0.49, 0.51, "Assumed", "2Q26 46.81%（公司）；8 月自結淨利 10.79 億 ÷ 營收 32.26 億＝33.4%（優分析 2026-10-05），與本假設推得的 3Q 淨利率相近", DW + "：大和估 3Q26 毛利率約 51%、營益率至 42%（Interested-party）")
inp("q_opx", "3Q、4Q26 營業費用 ÷ 營收", "%", 0.080, 0.085, 0.090, "Derived", "1Q26 9.8%、2Q26 9.0%（HiStock：毛利 − 營業利益）", "營收放大、費用率下降（Assumed）")
inp("q4g", "4Q26 營收季增", "%", 0.03, 0.10, 0.15, "Interested-party", UDN1008 + "：本國投顧估 4Q26 季增約 10%", "VR 均熱片 4Q 月出貨 50–60 萬片（工商時報 2026-08-07）")
inp("q_nop", "3Q、4Q26 業外收益（每季）", "NT$B", -0.10, 0.08, 0.20, "Derived", "1Q26 +0.079、2Q26 +0.092（HiStock：稅前 − 營業利益）", "FY2025 −0.37～+0.24（匯兌）")
inp("q_nir", "3Q、4Q26 歸屬母公司淨利 ÷ 稅前", "%", 0.78, 0.805, 0.82, "Derived", "1H26：歸母淨利 37.16 ÷ 稅前 46.16＝0.805（stockanalysis、HiStock）", "FY2025：52.77 ÷ 66.07＝0.799")
inp("q_sh", "股數（十億股；EPS 換算）", "十億股", None, "=3.716/25.33", None, "Derived", "1H26 歸母淨利 37.16 億 ÷ EPS 25.33（HiStock）", SA + "：流通股數 146.73 百萬股")

sec("G. 營業費用、業外、稅與股數（FY2027 起）")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.4, 0.6, 0.9, "Assumed", "FY2026 錨點實測約 0.8（Q 頁）", "營收倍增期費用率下降（Assumed）")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.02, 0.0, 0.02, "Assumed", "無", "Sens 用 ±2 個百分點")
inp("nop_f", "業外收益（每年）", "NT$B", -0.5, 0.3, 1.0, "Assumed", "FY2025 +0.11、1H26 +0.17（HiStock）", "匯兌損益波動大")
inp("nir_f", "歸屬母公司淨利 ÷ 稅前", "%", 0.78, 0.805, 0.82, "Derived", "同 q_nir", "—")
inp("sh_g", "股數年變動", "%", 0.0, 0.0, 0.01, "Assumed", "FY2025 推導 0.1436 → 1H26 0.1467 十億股（公司 EPS 與淨利）", "可轉債、現金增資：未找到")
inp("mf", "AI 部署係數（Sens 用；FY2027 起顆數 ×）", "x", 0.8, 1.0, 1.2, "Assumed", "網通模型 2027 顆數 23.7M（2026 15.2M）", "±20%")

sec("H. 共識、券商（只作 Recon 對帳，不作校準）")
inp("cons_rev26", "共識 FY2026 營收（stockanalysis，7 家）", "NT$B", None, 32.7, None, "Interested-party", SAF + "（頁面 2026-10-22；區間 32.4–33.3B）", "—")
inp("cons_eps26", "共識 FY2026 EPS（stockanalysis，7 家）", "NT$", None, 69.96, None, "Interested-party", SAF + "（區間 68.78–71.65）", "—")
for y, rv, ep in ((26, 28.940, 57.77), (27, 51.836, 117.18), (28, 80.811, 204.29)):
    inp(f"fs_rev{y}", f"FactSet 中位數 FY20{y} 營收（2026-05-22，9 家）", "NT$B", None, rv, None, "Interested-party", FS05, "較舊（5 月）；較新 FactSet 調查：未找到" if y == 26 else "—")
    inp(f"fs_eps{y}", f"FactSet 中位數 FY20{y} EPS（2026-05-22）", "NT$", None, ep, None, "Interested-party", FS05, "平均 59.33／121.01／191.02" if y == 26 else "—")
for y, ep in ((26, 71.0), (27, 151.0), (28, 262.0)):
    inp(f"ab_eps{y}", f"亞系券商 FY20{y} EPS（2026-08 下旬）", "NT$", None, ep, None, "Interested-party", PKT + "（口袋學堂 2026-08-26 轉引，券商名未載明）", "—")
inp("mw_eps27", "分析師 FY2027 EPS（理財周刊 2026-08-05）", "NT$", None, 123.83, None, "Interested-party", MW, "2026 年 60.98")
inp("dw_q3rev", "大和 3Q26 營收（2026-10-01）", "NT$B", None, 9.65, None, "Interested-party", DW, "目標價 10,000（73x 未來 12 月 EPS）")
inp("dw_q3eps", "大和 3Q26 EPS（2026-10-01）", "NT$", None, 22.77, None, "Interested-party", DW, "Bloomberg 共識 20.37")
inp("bb_q3eps", "Bloomberg 共識 3Q26 EPS（經大和轉引）", "NT$", None, 20.37, None, "Interested-party", DW, "—")
inp("ld_g27", "本國投顧 FY2027 營收年增（2026-10-08）", "%", None, 1.47, None, "Interested-party", UDN1008, "大和：2027–28 EPS 高於 Bloomberg 共識 14–21%")

sec("I. 價格與估值")
inp("px", "股價（2026-10-22 收盤，臺灣證交所）", "NT$", None, 6070, None, "Interested-party", SD + "：10/22 收盤 6,070（−465、−7.12%）", "stockanalysis 個股與 forecast 頁 2026-10-22 價格 6,070；10-23 收盤：未找到")
inp("pe_ttm", "近 12 月 P/E（財報狗，10-22）", "x", None, 136.99, None, "Interested-party", SD, "近 4 季 EPS 44.31")
inp("fpe", "前瞻 P/E（stockanalysis，2026-08-27 價格 5,645）", "x", None, 59.27, None, "Interested-party", SA, "較新的前瞻 P/E：未找到（頁面快取）")
inp("cons_feps", "共識前瞻 12 月 EPS（≈ 0.2 × 共識 FY2026＋0.8 × FactSet FY2027）", "NT$", None, "=0.2*cons_eps26+0.8*fs_eps27", None, "Derived",
    "stockanalysis FY2026 69.96（10-22）、FactSet FY2027 中位數 117.18（05-22，較舊）", "共識前瞻 EPS 的直接數字：未找到；大和稱其 2027–28 EPS 高於 Bloomberg 共識 14–21%")
inp("cons_tp", "市場共識目標價（賣方平均）", "NT$", None, 7204, None, "Interested-party", SAF + "：11 位分析師平均 7,204、中位數 7,300、區間 4,500–10,000（頁面 2026-10-22）", "大和 10,000（2026-10-01）；FactSet 中位數 4,920（2026-05-22，較舊）")
inp("nim_avg", "十年平均淨利率替代值（2016–2020 營收與淨利未找到；依工作單以情境最後一年 × 0.8）", "%", None, 0, None, "Assumed",
    "工作單 2026-10-25：找不到十年平均時用最後一年 × 0.8；本值＝情境 C FY2029 淨利率 × 0.8（公式於建置時連結）", "參考：FY2021–25 五年平均 21.6%（stockanalysis：13.6%、21.7%、19.1%、24.0%、26.0%）")

sec("J. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_r", "折現率／要求報酬", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%，各市場同一值）", "DCF 折現；情境價值推到 12 個月後")
inp("tp_g", "長期成長率", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "DCF 永續成長與 FY2032 收斂目標（情境 B）")
inp("tp_roe", "ROE（近三年平均，FY2023–25）", "%", None, "=AVERAGE(0.1975,0.2530,0.2816)", None, "Derived", SAR + "：FY2023 19.75%、FY2024 25.30%、FY2025 28.16%", "Goodinfo：19.8%、25.3%、28.2%（同值）")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（2026-10-25 11:53 定 DCF／Comparable 各半）", "—")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（情境 B FY2029 與十年平均替代值的中點）", "%", None, 0, None, "Assumed", "PL 頁情境 B FY2029、In 頁 nim_avg（十年平均未找到）", "—")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（＝十年平均替代值 nim_avg）", "%", None, "=nim_avg", None, "Assumed", "In 頁 nim_avg（十年平均未找到，用情境 C FY2029 × 0.8）", "—")
TP_KEYS += ["tp_r", "tp_g", "tp_roe", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]

URLN = {PKT: "口袋學堂 2026-08-26", PKT2: "口袋學堂 2025-12-15", MW: "理財周刊 2026-08-05", CT0807: "工商時報 2026-08-07", UA1005: "優分析 2026-10-05",
        DW: "經濟日報 2026-10-01（大和）", UDN1008: "經濟日報 2026-10-08", TN9: "technews 2026-10-08", TN7: "technews 2026-08-07", FS05: "鉅亨 2026-05-22（FactSet）",
        SA: "stockanalysis 個股頁", SAQ: "stockanalysis 季度損益", SAR: "stockanalysis 比率頁", SAF: "stockanalysis forecast 頁", HISQ: "HiStock 損益表",
        HISE: "HiStock 每股盈餘", GI: "Goodinfo", WEARN: "聚財網", SD: "財報狗"}
def nourl(t):   # In、Q 頁不放網址（HTML 展示層須自含）；網址列在 Sources 頁
    if not isinstance(t, str): return t
    for u, n in sorted(URLN.items(), key=lambda x: -len(x[0])):
        t = t.replace(u + "：", n + "：").replace(u, n)
    return t
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
    IN.cell(r, 7, tag); IN.cell(r, 8, nourl(s1)); IN.cell(r, 9, nourl(s2))
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
def alias(new, base, suf, years=(27, 28, 29)):
    for y in years:
        wb.defined_names[f"{new}_{y}"] = DefinedName(f"{new}_{y}", attr_text=wb.defined_names[f"{base}_{y}{suf}"].attr_text)
for nm, base, suf in (("VRLO", "asp_vr", "_lo"), ("VRHI", "asp_vr", "_hi"), ("AMLO", "asp_amd", "_lo"), ("AMHI", "asp_amd", "_hi"),
                      ("ASLO", "asp_asic", "_lo"), ("ASHI", "asp_asic", "_hi"), ("RLO", "rel", "_lo"), ("RHI", "rel", "_hi"),
                      ("LCLO", "lcs", "_lo"), ("LCHI", "lcs", "_hi"), ("GOLO", "g_oth", "_lo"), ("GOHI", "g_oth", "_hi"), ("GLLO", "g_lf", "_lo"), ("GLHI", "g_lf", "_hi")):
    alias(nm, base, suf)

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億新台幣；EPS 為新台幣；股數為十億股）。藍字＝輸入；黑字＝公式。實績取自健策財報與月營收公告（經 stockanalysis、HiStock、technews 轉述）"
Q["A2"] = ("FY2026 錨點＝1H26 財報＋3Q26 月營收實績（7–9 月 94.93 億）＋假設（3Q、4Q 毛利率、營業費用率、業外、淨利 ÷ 稅前；4Q 營收＝3Q × (1＋季增)）。"
           "3Q26 財報預計 2026-11-09（stockanalysis），截至 2026-10-25 尚未公布；公司每月公告自結損益（8 月淨利 10.79 億、EPS 7.35），7、9 月自結數：未找到")
hdr(Q, 3, ["項目", "FY2025A", "1H26（實績）", "3Q26（營收實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1", "來源2／說明"])
S1 = "stockanalysis 季度損益（" + SAQ + "）：FY2025 營收 202.76 億、毛利 84.34 億、營業利益 64.98 億、歸母淨利 52.77 億；1Q26、2Q26 各季"
S2 = "HiStock 季度損益（" + HISQ + "）：營收、毛利、營業利益、稅前、稅後同值（稅後含非控制權益）"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2=S2, ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"F{r}"] = ann or f"=SUM(C{r}:E{r})"; Q[f"F{r}"].number_format = fmt
    Q.cell(r, 7, tag); Q.cell(r, 8, nourl(s1)); Q.cell(r, 9, nourl(s2))
qrow(4, "營收", {"B": 20.276, "C": "=5.305+7.275", "D": "=q3_rev", "E": "=D4*(1+q4g)"}, tag="Verified／Interested-party（4Q 季增）", s2="1H26 125.80 億（口袋學堂、理財周刊同值）")
qrow(5, "毛利", {"B": 8.434, "C": "=2.217+3.405", "D": "=D4*q_gm", "E": "=E4*q_gm"}, tag="Verified／Assumed")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="FY2025 41.6%；1Q26 41.8%、2Q26 46.81%、1H26 44.69%（口袋學堂）")
qrow(7, "營業費用", {"B": "=B5-B8", "C": "=C5-C8", "D": "=D4*q_opx", "E": "=E4*q_opx"}, tag="Derived（毛利 − 營業利益）／Assumed")
qrow(8, "營業利益", {"B": 6.498, "C": "=1.697+2.749", "D": "=D5-D7", "E": "=E5-E7"}, tag="Verified", s2="2Q26 營益率 37.8%（口袋學堂）")
qrow(9, "營業利益率", {c: f"={c}8/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(10, "業外收益", {"B": "=B11-B8", "C": "=C11-C8", "D": "=q_nop", "E": "=q_nop"}, tag="Derived")
qrow(11, "稅前利益", {"B": "=1.7405+1.3271+1.7227+1.8167", "C": "=1.7756+2.8409", "D": "=D8+D10", "E": "=E8+E10"}, tag="Verified（HiStock 各季）／Derived", s1=HISQ, s2="各季稅前：1Q25 17.41、2Q25 13.27、3Q25 17.23、4Q25 18.17、1Q26 17.76、2Q26 28.41 億")
qrow(12, "歸屬母公司淨利", {"B": 5.277, "C": "=1.398+2.318", "D": "=D11*q_nir", "E": "=E11*q_nir"}, tag="Verified／Derived（3Q、4Q）", s2="稅後（含非控制權益）FY2025 53.10、1H26 37.33 億（HiStock）")
qrow(13, "淨利 ÷ 稅前（含稅與非控制權益）", {c: f"={c}12/{c}11" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(14, "股數（十億股）", {"B": "=5.277/36.75", "C": "=q_sh", "D": "=q_sh", "E": "=q_sh"}, ann="=q_sh", tag="Derived", s2="歸母淨利 ÷ 公司 EPS")
qrow(15, "EPS", {c: f"={c}12/{c}14" for c in "BCDE"}, fmt="0.00", ann="=F12/F14", tag="Derived", s1="淨利 ÷ 股數", s2="—")
qrow(16, "  公司揭露 EPS", {"B": 36.75, "C": 25.33}, fmt="0.00", ann="", tag="Verified", s1=HISE, s2="FY2025 36.75（HiStock、Goodinfo、聚財網；stockanalysis 列 36.14，口徑未載明）；1Q26 9.53、2Q26 15.80")
SPL = [(17, "AI 加速器均熱片營收（＝營收 × AI 占比）", "=B4*ai_25", "=F4*ai_26"),
       (18, "其他散熱營收（＝營收 ×（散熱占比 − AI 占比））", "=B4*(th_25-ai_25)", "=F4*(th_26-ai_26)"),
       (19, "導線架、精密沖壓與其他（＝營收 ×（1 − 散熱占比））", "=B4*(1-th_25)", "=F4*(1-th_26)")]
for rr, lab, fb, ff in SPL:
    qrow(rr, lab, {"B": fb}, ann=ff, tag="Assumed／Interested-party", s1="In 頁 D 節", s2="各產品季度金額未揭露：未找到" if rr == 17 else "—")
Q.cell(22, 1, "FY2026 錨點的檢查（參考；校準未使用券商數字）").font = BOLD
Q.cell(23, 1, "FY2026 營收年增"); Q["B23"] = "=F4/B4-1"
Q.cell(24, 1, "FY2026 營業費用年增"); Q["B24"] = "=F7/B7-1"
Q.cell(25, 1, "營業費用彈性（營業費用年增 ÷ 營收年增；In 頁 phi 的參考）"); Q["B25"] = "=B24/B23"
Q.cell(26, 1, "FY2026 營收 ÷ 共識（stockanalysis）− 1"); Q["B26"] = "=F4/cons_rev26-1"
Q.cell(27, 1, "FY2026 EPS ÷ 共識（stockanalysis）− 1"); Q["B27"] = "=F15/cons_eps26-1"
Q.cell(28, 1, "3Q26 淨利率（模型）vs 8 月自結淨利率 33.4%（百分點）"); Q["B28"] = "=D12/D4-10.79/32.26"
for rr in range(23, 29): Q[f"B{rr}"].number_format = "0.0%" if rr != 25 else "0.00"
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 7), ("oi", 8), ("nop", 10), ("ni", 12), ("sh", 14), ("eps", 15), ("ail", 17), ("oth", 18), ("lf", 19)):
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
# 情境與行業一致：A＝可拆式上蓋滲透快（VR、AMD、ASIC 單價高值）、份額與液冷份額高值；B＝基準；C＝低值。其他散熱、導線架各情境相同
SCN_P = {"A": dict(vr="VRHI", amd="AMHI", asic="ASHI", rel="RHI", lcs="LCHI"), "B": dict(), "C": dict(vr="VRLO", amd="AMLO", asic="ASLO", rel="RLO", lcs="LCLO")}
DEFP = dict(scn="B", vr="asp_vr", amd="asp_amd", asic="asp_asic", ru="asp_ru", bw="asp_bw", rel="rel", lcs="lcs", k="k_non", ai26="ai_26", mf="1",
            go="g_oth", gl="g_lf", d="d_ai", opxa="0", phi="phi", nir="nir_f")
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
    mfx = lambda y: "" if y == 2026 else f"*{P['mf']}"
    asp = lambda base, nm, y: f"{base}_26" if y == 2026 else f"{nm}_{y % 100}"
    sec_("A. AI 加速器顆數與均熱片市場（D 欄＝FY2026 校準年；E–G＝FY2027–FY2029；顆數 × 同字母情境係數）")
    for key, lab, us in (("ubw", "Blackwell 顆數", ("bw",)), ("uhp", "Hopper 顆數", ("hp",)), ("urb", "Rubin 顆數", ("rb",)), ("uru", "Rubin Ultra 顆數", ("ru",)),
                         ("uamd", "AMD Instinct 顆數", ("amd",)), ("uasic", "ASIC 顆數（TPU＋Trainium＋其他）", ("tpu", "trn", "oth"))):
        row(key, lab, "百萬顆", {c_(y): "=(" + "+".join(f"u_{u}{y % 100}" for u in us) + f")*{umx(y, s)}{mfx(y)}" for y in BU}, fmt="0.00")
    row("pvr", "Rubin 每顆均熱片／上蓋", "US$/顆", {c_(y): "=" + asp("asp_vr", P["vr"], y) for y in BU}, fmt="#,##0")
    row("pamd", "AMD 每顆均熱片／上蓋", "US$/顆", {c_(y): "=" + asp("asp_amd", P["amd"], y) for y in BU}, fmt="#,##0")
    row("pasic", "ASIC 每顆均熱片＋加強環", "US$/顆", {c_(y): "=" + asp("asp_asic", P["asic"], y) for y in BU}, fmt="#,##0")
    row("mnv", "NVIDIA 加速器均熱片市場（美元）", "US$B", {c_(y): f"=({c_(y)}{R['ubw']}*{P['bw']}+{c_(y)}{R['uhp']}*asp_hp+{c_(y)}{R['urb']}*{c_(y)}{R['pvr']}+{c_(y)}{R['uru']}*{P['ru']})/1000" for y in BU}, fmt="#,##0.000",
        note="百萬顆 × 美元 ÷ 1000＝十億美元")
    row("mnon", "非 NVIDIA（AMD、ASIC）均熱片市場（美元）", "US$B", {c_(y): f"=({c_(y)}{R['uamd']}*{c_(y)}{R['pamd']}+{c_(y)}{R['uasic']}*{c_(y)}{R['pasic']})/1000" for y in BU}, fmt="#,##0.000")
    row("mkt", "AI 加速器均熱片市場合計", "US$B", {c_(y): f"={c_(y)}{R['mnv']}+{c_(y)}{R['mnon']}" for y in BU}, fmt="#,##0.000")
    row("sh0", "健策 NVIDIA 份額（FY2026 校準）＝ AI 均熱片營收 ÷（NVIDIA 市場＋k × 非 NVIDIA 市場）", "%",
        {"D": f"=a26_rev*{P['ai26']}/fx/(D{R['mnv']}+{P['k']}*D{R['mnon']})"}, fmt="0.0%", note="Derived：分子為 FY2026 錨點營收 × AI 占比（Assumed）；k＝非 NVIDIA 份額 ÷ NVIDIA 份額")
    row("rel", "份額相對值（÷ FY2026）", "x", {"D": "=1", **{c_(y): f"={P['rel']}_{y % 100}" for y in BU[1:]}}, fmt="0.00")
    row("shnv", "健策 NVIDIA 均熱片份額", "%", {c_(y): f"=$D${R['sh0']}*{c_(y)}{R['rel']}" for y in BU}, fmt="0.0%")
    row("shnon", "健策非 NVIDIA 均熱片份額", "%", {c_(y): f"={c_(y)}{R['shnv']}*{P['k']}" for y in BU}, fmt="0.0%")
    row("ail", "AI 加速器均熱片營收", "NT$B", {"C": "=a25_ail", **{c_(y): f"=({c_(y)}{R['mnv']}*{c_(y)}{R['shnv']}+{c_(y)}{R['mnon']}*{c_(y)}{R['shnon']})*fx" for y in BU}}, fmt="#,##0.0")
    sec_("B. 液冷冷板元件（奇鋐模型同字母情境的液冷市場 × 健策份額）")
    row("lmk", "AI 液冷市場（奇鋐模型；美元）", "US$B", {c_(y): f"=lmk{s}_{y % 100}{mfx(y)}" for y in BU}, fmt="#,##0.0", note="唯讀 avc/dist PL 頁同字母情境；FY2026 健策液冷含在其他散熱內")
    row("lcs", "健策液冷份額", "%", {"D": "=0", **{c_(y): f"={P['lcs']}_{y % 100}" for y in BU[1:]}}, fmt="0.0%")
    row("lc", "液冷冷板元件營收（MCL、MLCP）", "NT$B", {"C": "=0", **{c_(y): f"={c_(y)}{R['lmk']}*{c_(y)}{R['lcs']}*fx" for y in BU}}, fmt="#,##0.0")
    sec_("C. 其他散熱、導線架與合計")
    row("oth", "其他散熱營收（CPU 均熱片、交換器加強環、CPO、ILM、車用水冷）", "NT$B",
        {"C": "=a25_oth", "D": f"=a26_rev*(th_26-{P['ai26']})", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['go']}_{y % 100})" for y in BU[1:]}}, fmt="#,##0.0")
    row("lf", "導線架、精密沖壓與其他營收", "NT$B", {"C": "=a25_lf", "D": "=a26_lf", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gl']}_{y % 100})" for y in BU[1:]}}, fmt="#,##0.0")
    row("rev", "營收", "NT$B", {c_(y): f"={c_(y)}{R['ail']}+{c_(y)}{R['lc']}+{c_(y)}{R['oth']}+{c_(y)}{R['lf']}" for y in FYS}, fmt="#,##0.0")
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("aish", "AI 加速器均熱片 ÷ 營收", "%", {c_(y): f"={c_(y)}{R['ail']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("pgu", "每顆加速器健策 AI 均熱片營收（全部加速器平均）", "US$/顆",
        {c_(y): f"={c_(y)}{R['ail']}/fx*1000/({c_(y)}{R['ubw']}+{c_(y)}{R['uhp']}+{c_(y)}{R['urb']}+{c_(y)}{R['uru']}+{c_(y)}{R['uamd']}+{c_(y)}{R['uasic']})" for y in BU}, fmt="#,##0.0")
    sec_("D. 損益")
    row("gmn", "其他營收毛利率 ＝ 錨點毛利率 − 差距 × AI 占比（FY2026 校準）", "%", {"D": f"=a26_gp/D{R['rev']}-{P['d']}*D{R['aish']}"}, fmt="0.0%",
        note="Derived：讓 FY2026 毛利率等於錨點；AI 均熱片毛利率＝其他＋差距；液冷與其他營收同毛利率")
    row("gp", "毛利", "NT$B", {"C": "=a25_gp", "D": "=a26_gp", **{c_(y): f"=$D${R['gmn']}*{c_(y)}{R['rev']}+{P['d']}*{c_(y)}{R['ail']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("gm", "毛利率", "%", {c_(y): f"={c_(y)}{R['gp']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("opb", "營業費用（基準路徑）", "NT$B", {"C": "=a25_opx", "D": "=a26_opx", **{c_(y): f"={c_(y - 1)}{r}*(1+MAX(0,{P['phi']}*{c_(y)}{R['g']}))" for y in BU[1:]}}, fmt="#,##0.0")
    row("opx", "營業費用", "NT$B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", **{c_(y): f"={c_(y)}{R['opb']}+{P['opxa']}*{c_(y)}{R['rev']}" for y in BU[1:]}}, fmt="#,##0.0")
    row("oi", "營業利益", "NT$B", {c_(y): f"={c_(y)}{R['gp']}-{c_(y)}{R['opx']}" for y in FYS}, fmt="#,##0.0")
    row("oim", "營業利益率", "%", {c_(y): f"={c_(y)}{R['oi']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("nop", "業外收益", "NT$B", {"C": "=a25_nop", "D": "=a26_nop", **{c_(y): "=nop_f" for y in BU[1:]}}, fmt="#,##0.00")
    row("ni", "歸屬母公司淨利", "NT$B", {"C": "=a25_ni", "D": "=a26_ni", **{c_(y): f"=({c_(y)}{R['oi']}+{c_(y)}{R['nop']})*{P['nir']}" for y in BU[1:]}}, fmt="#,##0.0")
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
PL["A1"] = "PL — 健策損益（十億新台幣；EPS 新台幣；歸屬母公司）。FY2025A 實績、FY2026 錨點（1H26 財報＋3Q26 月營收＋假設），FY2027 起為模型。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要。顆數係數從 CY2028 開始；可拆式上蓋單價、份額、液冷份額從 FY2027 起依情境取高／基準／低值"
PL["A2"] = "灰底＝FY2030–FY2032 收斂法（上游只到 2029）：營收年增線性收斂到 fade_g、淨利率線性收斂到 fade_m（In 頁 J 節），非驅動因子推導"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A（順利）：顆數基準（tsmc 係數 A＝1.0）；奇鋐模型液冷市場 A；可拆式上蓋單價、份額、液冷份額取高值",
         "B": "情境 B（平平）：tsmc 係數 B（CY2028 ×0.8、CY2029 ×0.85）；液冷市場 B；公司特有變數取基準",
         "C": "情境 C（不如預期）：tsmc 係數 C（CY2028–29 ×0.6）；液冷市場 C；可拆式上蓋單價、份額、液冷份額取低值"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s, **SCN_P[s])
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目"] + YH[1:])
sr = 4; SUMR = {}
SUMK = (("rev", "營收", "NT$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("ail", "  AI 加速器均熱片營收", "NT$B", "#,##0.0"),
        ("lc", "  液冷冷板元件營收", "NT$B", "#,##0.0"), ("oth", "  其他散熱營收", "NT$B", "#,##0.0"), ("lf", "  導線架與其他營收", "NT$B", "#,##0.0"),
        ("gm", "毛利率", "%", "0.0%"), ("oim", "營業利益率", "%", "0.0%"),
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
IN.cell(INROW["nim_avg"], 5, f"=PL!G{SUMR[('C', 'nim')]}*0.8")
for k in ("fade_mA", "fade_mB", "nim_avg"): IN.cell(INROW[k], 5).font = Font(color="000000")
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型 − 錨點").font = BOLD; sr += 1
CHK = []
for lab, f in (("  三情境：營收差（NT$B）", "+".join(f"ABS(D{PB[s]['rev']}-a26_rev)" for s in SC)),
               ("  情境 B：AI 均熱片差（NT$B）", f"D{PB['B']['ail']}-a26_ail"),
               ("  情境 B：毛利差（NT$B）", f"D{PB['B']['gmn']}*D{PB['B']['rev']}+d_ai*D{PB['B']['ail']}-a26_gp")):
    PL.cell(sr, 1, lab); PL[f"D{sr}"] = "=" + f; PL[f"D{sr}"].number_format = "0.0000"; CHK.append(sr); sr += 1
assert sr < 60, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 10; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 13

# ---------------- Units：加速器顆數、每顆均熱片金額、份額；液冷份額合計 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — AI 加速器均熱片市場（顆數 × 每顆金額）、健策份額與每顆營收；液冷冷板元件份額與奇鋐份額合計（情境 B）。顆數取自 markets/network（× tsmc 同字母係數）；液冷市場取自奇鋐模型 avc/dist（同字母情境）"
B_ = PB["B"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("utot", "AI 加速器顆數合計（情境 B）", "百萬顆", lambda y: "=" + "+".join(f"PL!{FC[y]}{B_[k]}" for k in ("ubw", "uhp", "urb", "uru", "uamd", "uasic")), "0.0", "markets/network Units × tsmc 係數 B"),
        ("unv", "  其中 NVIDIA（Blackwell、Hopper、Rubin、Rubin Ultra）", "百萬顆", lambda y: "=" + "+".join(f"PL!{FC[y]}{B_[k]}" for k in ("ubw", "uhp", "urb", "uru")), "0.0", ""),
        ("pvr", "Rubin 每顆均熱片／上蓋", "US$/顆", lambda y: f"=PL!{FC[y]}{B_['pvr']}", "#,##0", "2027 起可拆式上蓋滲透（In 頁 B 節）"),
        ("mnv", "NVIDIA 加速器均熱片市場", "US$B", lambda y: f"=PL!{FC[y]}{B_['mnv']}", "0.000", ""),
        ("mnon", "非 NVIDIA（AMD、ASIC）均熱片市場", "US$B", lambda y: f"=PL!{FC[y]}{B_['mnon']}", "0.000", ""),
        ("pavg", "市場平均每顆金額", "US$/顆", lambda y: f"=PL!{FC[y]}{B_['mkt']}*1000/{cyc(y)}4", "#,##0.0", ""),
        ("shnv", "健策 NVIDIA 份額", "%", lambda y: f"=PL!{FC[y]}{B_['shnv']}", "0.0%", "FY2026 以 AI 均熱片營收校準"),
        ("shnon", "健策非 NVIDIA 份額", "%", lambda y: f"=PL!{FC[y]}{B_['shnon']}", "0.0%", "＝ NVIDIA 份額 × k_non"),
        ("ail", "健策 AI 加速器均熱片營收", "US$B", lambda y: f"=PL!{FC[y]}{B_['ail']}/fx", "0.000", ""),
        ("shall", "健策 AI 均熱片 ÷ 市場合計", "%", lambda y: f"=PL!{FC[y]}{B_['ail']}/fx/PL!{FC[y]}{B_['mkt']}", "0.0%", ""),
        ("othm", "其他廠商均熱片（市場 − 健策）", "US$B", lambda y: f"=PL!{FC[y]}{B_['mkt']}-PL!{FC[y]}{B_['ail']}/fx", "0.000", "其他 IHS 供應商；奇鋐、雙鴻不以均熱片為主，與本項無重疊"),
        ("lmk", "AI 液冷市場（奇鋐模型 B）", "US$B", lambda y: f"=PL!{FC[y]}{B_['lmk']}", "0.0", "機架內冷板＋分歧管"),
        ("lcs", "健策液冷份額", "%", lambda y: f"=PL!{FC[y]}{B_['lcs']}", "0.0%", "Assumed（MCL、MLCP）"),
        ("avsh", "奇鋐液冷有效份額（奇鋐模型 B）", "%", lambda y: "=avshB_" + str(y % 100), "0.0%", "唯讀 avc/dist"),
        ("shsum", "奇鋐＋健策液冷份額合計", "%", lambda y: f"={cyc(y)}17+{cyc(y)}16", "0.0%", "其餘為雙鴻、酷碼、台達電、Boyd、CoolIT 等；雙鴻模型未在 main，未納入合計"),
        ("lcu", "健策液冷冷板元件營收", "US$B", lambda y: f"=PL!{FC[y]}{B_['lc']}/fx", "0.000", "")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); UR[key] = ur; ur += 1
assert UR["lcs"] == 16 and UR["avsh"] == 17, UR
ur += 1
UN.cell(ur, 1, "讀法：VR 世代每顆均熱片金額約為 Blackwell 的 3 倍，2027 年起可拆式上蓋再提高單價；健策營收成長主要來自每顆金額，而非顆數。FY2026 的 AI 占比（Assumed）決定校準份額；可拆式上蓋單價與份額相對值是最需要每季檢查的輸入。液冷冷板元件份額為新產品 Assumed，與奇鋐份額合計遠低於 100%。").font = BOLD
UNOTE = ur
UN.column_dimensions["A"].width = 66; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 14
UN.column_dimensions["G"].width = 60

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，年底落後 P/E）只作參考。本版未做先驗，無 Andy 合理 P/E 對照"
V["A2"] = "循環位置：健策不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照奇鋐、台達電；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
O = 54
V.cell(O, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, O + 1, ["項目", "A", "B", "C", "說明"])
V.cell(O + 2, 1, "現價（2026-10-22 收盤）"); V[f"B{O + 2}"] = "=px"; V[f"B{O + 2}"].number_format = "#,##0"
for k, y, lab in ((3, 2026, "現價 ÷ FY2026 EPS（錨點）"), (4, 2027, "現價 ÷ FY2027 EPS（模型）"), (5, 2028, "現價 ÷ FY2028 EPS（模型）"), (6, 2029, "現價 ÷ FY2029 EPS（模型）")):
    V.cell(O + k, 1, lab)
    for j, s in enumerate(SC):
        c = "BCD"[j]; V[f"{c}{O + k}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{O + k}"].number_format = "0.0"
V.cell(O + 7, 1, "未來 12 月 EPS（模型 B）≈ 0.2 × FY2026＋0.8 × FY2027"); V[f"B{O + 7}"] = f"=0.2*PL!D{SUMR[('B', 'eps')]}+0.8*PL!E{SUMR[('B', 'eps')]}"; V[f"B{O + 7}"].number_format = "0.00"
V.cell(O + 8, 1, "  現價 ÷ 未來 12 月 EPS（模型 B）"); V[f"B{O + 8}"] = f"=px/B{O + 7}"; V[f"B{O + 8}"].number_format = "0.0"
V.cell(O + 9, 1, "近 12 月 P/E（財報狗，10-22）"); V[f"B{O + 9}"] = "=pe_ttm"; V[f"B{O + 9}"].number_format = "0.0"
V.cell(O + 10, 1, "前瞻 P/E（stockanalysis，08-27）"); V[f"B{O + 10}"] = "=fpe"; V[f"B{O + 10}"].number_format = "0.0"
V[f"E{O + 3}"] = "FY2026 三情境相同（錨點）"; V[f"E{O + 4}"] = "FY2027 三情境差異只來自可拆式上蓋單價、份額、液冷份額（顆數係數從 CY2028 開始）"
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
RB0 = vr + 3
# 十年年底 P/E：Goodinfo 年底收盤 ÷ 當年 EPS（2016 年 EPS 取聚財網）；stockanalysis 年度 P/E（2021–2025）只作核對
GI_PX = [47.5, 66.7, 78.2, 234.5, 247.5, 408, 376.5, 769, 1525, 2745]
GI_EPS = [2.32, 2.19, 4.61, 6.55, 8.27, 9.88, 19.54, 16.59, 24.15, 36.75]
PE_GI = [f"={p}/{e}" for p, e in zip(GI_PX, GI_EPS)]
PE_SA = [None] * 5 + [41.68, 19.49, 47.12, 63.57, 74.34]
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "Goodinfo（年底收盤 ÷ 當年 EPS）", "vals": PE_GI, "tag": "Derived（Goodinfo 年底收盤與 EPS；2016 年 EPS 2.32 取聚財網；10 年）"},
    {"label": "stockanalysis（年度 P/E）", "vals": PE_SA, "tag": "Interested-party（資料商；只有 2021–2025，只作核對、不採用）", "use": False,
     "note": "stockanalysis 2021–2025 與 Goodinfo 推導差距在 1x 內；2016–2020 年：未找到第二來源"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "#,##0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位。採用 Goodinfo 年底收盤 ÷ 當年 EPS（10 年；2016 年 EPS 取聚財網）；stockanalysis 只有 2021–2025，只作核對。健策 2019 年起已是 CPU 均熱片龍頭、2023 年起 AI 均熱片放量，十年倍數由約 17x 升到約 75x，分布本身已含 AI 溢價。")
assert V.cell(RB0 + 5, 1).value == "2016" and V.cell(RB0 + 14, 1).value == "2025", (V.cell(RB0 + 5, 1).value, V.cell(RB0 + 14, 1).value)
PE10 = f"$B${RB0 + 5}:$B${RB0 + 14}"
rule_rows = [rw for rw in range(RB0, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")

# 0. 目標價（Andy 2026-10-25 11:28／11:53／12:22／14:00：DCF 與 Comparable 各半；DCF 負值取 0；列標籤依工作單逐字；B／C／D＝A 順利／B 平平／C 不如預期）
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
tset(22, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=MAX(0,SUMPRODUCT({c}16:{c}21,1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}21*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6)" for c in "BCD"], "#,##0",
     "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底；負值取 0（Andy 12:22）")
tset(23, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}22*(1+tp_r)" for c in "BCD"], "#,##0")
tset(24, "dcftp", "DCF 目標價（情境加權）", ["=tp_wA*B23+tp_wB*C23+tp_wC*D23", None, None], "#,##0", "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁")
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E；非 A 股公司，同業不放 A 股）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = [("奇鋐（AVC）", "3017.TW", 27.04, "2026-10-05"), ("雙鴻（Auras）", "3324.TWO", 17.27, "2026-08-28"),
         ("Vertiv", "VRT", 24.53, "2026-10-23"), ("台達電（Delta）", "2308.TW", 30.51, "2026-10-22")]
for i, (nm, tk, pe, dt) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左，臺股頁面為快取日期）"); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
V.cell(33, 1, "並列參考（不進目標價）：英維克（002837.SZ）為 A 股，依規則不列；均熱片專業同業（上市可比）：未找到")
tset(34, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", "散熱與液冷同業：奇鋐、雙鴻、Vertiv、台達電（與奇鋐模型同一比較組，以奇鋐取代健策）")
tset(35, "feps", "12 個月後的前瞻 EPS（情境加權）", [f"=tp_wA*(2/12*B{EPSR[2027]}+10/12*B{EPSR[2028]})+tp_wB*(2/12*C{EPSR[2027]}+10/12*C{EPSR[2028]})+tp_wC*(2/12*D{EPSR[2027]}+10/12*D{EPSR[2028]})", None, None], "0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(36, "comptp", "Comparable 目標價", ["=B34*B35", None, None], "#,##0", "同業中位數 × 前瞻 EPS")
V.cell(38, 1, "0.3 結論").font = BOLD
hdr(V, 39, ["項目", "數值", "", "", "說明"])
tset(40, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B24+(1-tp_w_dcf)*B36", None, None], "#,##0", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁")
tset(41, "comp2", "Comparable 目標價（市場主流同業法）", ["=B36", None, None], "#,##0")
tset(42, "cons", "市場共識目標價", ["=cons_tp", None, None], "#,##0", "stockanalysis（S&P Global）：11 位分析師平均 7,204、中位數 7,300（頁面 2026-10-22）")
tset(43, "upw", "上檔空間：加權目標價", ["=B40/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-22 收盤）")
tset(44, "upc", "上檔空間：Comparable 目標價", ["=B41/px-1", None, None], "0.0%")
tset(45, "grade", "上檔空間級距", ['=IF(B43>0.5,"大於 50%",IF(B43>0.25,"25%–50%",IF(B43>=0,"0%–25%",IF(B43>=-0.25,"0%～−25%",IF(B43>=-0.5,"−25%～−50%","小於 −50%")))))', None, None], "@",
     "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）；負值區間同樣重要")
tset(46, "rec", "投資建議", ['=IF(B43>0.5,"強力買進",IF(B43>0.25,"買進",IF(B43>=-0.25,"中立",IF(B43>=-0.5,"放空","強力放空"))))', None, None], "@",
     "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
tset(47, "geps", "與市場共識目標價的差距：EPS 因素", ["=B35/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS（In 頁 cons_feps）為現在起算 12 個月，期間與本模型不同")
tset(48, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B34-1", None, None], "0.0%")
V["E48"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B34,"0.0")&"x"'
tset(49, "band", "價格帶（DCF 不如預期～順利）", ['=TEXT(D23,"#,##0")&"～"&TEXT(B23,"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(50, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=IF(px/{c}{EPSR[2028]}<MIN({PE10}),0,IF(px/{c}{EPSR[2028]}>MAX({PE10}),1,PERCENTRANK({PE10},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 6 節 Goodinfo）的分位；只作檢查，不進目標價")
V.cell(51, 1, "目標價方法")
V["B51"] = ("DCF：股東現金流＝歸母 EPS ×（1 − g／ROE），r＝10%、g＝4%、ROE 為近三年平均約 24%（再投資率約 16%），FY2027–29 驅動因子、FY2030–32 收斂法＋永續；"
            "本模型未建資本支出與折舊（大園新廠 2026-04 投產），未改用淨利＋折舊－資本支出。"
            "Comparable：奇鋐、雙鴻、Vertiv、台達電前瞻 P/E 中位數（散熱與液冷主流比較組，不放 A 股）；健策現價前瞻 P/E 約 59x，遠高於同業（均熱片近乎寡占、毛利率約 47%），Comparable 未給寡占溢價，方向偏低。"
            "情境對應：A＝順利（顆數基準、可拆式上蓋與份額高值）、B＝平平（顆數 ×0.8／0.85、基準）、C＝不如預期（顆數 ×0.6、低值）")
for r_ in (24, 36, 40):
    V[f"B{r_}"].fill = KEY
V.cell(40, 1).font = BOLD
defname("tp_final", "Val!$B$40"); defname("tp_dcf", "Val!$B$24"); defname("tp_comp", "Val!$B$36")
V["A3"] = "目標價：見第 0 節（DCF 與 Comparable 加權）；第 3 節規則值不是目標價，只作參考"
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 B＝基準輸入；健策份額、其他營收毛利率在每個案例內依 FY2026 錨點重算，FY2026 恆等於錨點）：對 FY2027、FY2028 的營收、AI 均熱片營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 AI 均熱片", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 AI 均熱片", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("可拆式上蓋 慢（VR 單價 低）", "Rubin 每顆 2027 US$110、2028–29 120", dict(vr="VRLO")),
         ("可拆式上蓋 快（VR 單價 高）", "Rubin 每顆 2027 US$200、2028–29 220", dict(vr="VRHI")),
         ("AMD、ASIC 單價 低", "AMD 45／50／50、ASIC 30", dict(amd="AMLO", asic="ASLO")),
         ("AMD、ASIC 單價 高", "AMD 80／110／110、ASIC 45／55／55", dict(amd="AMHI", asic="ASHI")),
         ("份額流失（低）", "份額相對值 0.92／0.85／0.80（可拆式上蓋新供應商切入）", dict(rel="RLO")),
         ("份額守住（高）", "份額相對值 1.00／1.00／0.98", dict(rel="RHI")),
         ("非 NVIDIA 份額比 0.25", "k_non 0.25（FY2026 NVIDIA 份額校準變高）", dict(k="k_non_lo")),
         ("非 NVIDIA 份額比 0.55", "k_non 0.55", dict(k="k_non_hi")),
         ("2026 AI 占比 32%（讀法）", "FY2026 AI 均熱片＝營收 × 32%；份額變低、其他散熱變大", dict(ai26="ai_26_lo")),
         ("2026 AI 占比 48%（讀法）", "FY2026 AI 均熱片＝營收 × 48%", dict(ai26="ai_26_hi")),
         ("液冷份額 低", "0.3%／0.6%／1.0%", dict(lcs="LCLO")), ("液冷份額 高", "1.5%／3.0%／4.5%", dict(lcs="LCHI")),
         ("AI 部署 ×0.8", "FY2027 起顆數與液冷市場 ×0.8", dict(mf="mf_lo")), ("AI 部署 ×1.2", "×1.2", dict(mf="mf_hi")),
         ("AI 毛利率差距 6pt", "AI 均熱片毛利率只高 6 個百分點", dict(d="d_ai_lo")), ("AI 毛利率差距 18pt", "高 18 個百分點", dict(d="d_ai_hi")),
         ("其他散熱成長 低", "+5％／+3％／+2％", dict(go="GOLO")), ("其他散熱成長 高", "+20％／+15％／+12％", dict(go="GOHI")),
         ("營業費用彈性 0.9", "基準 0.6", dict(phi="phi_hi")), ("營業費用率 +2pt", "營業費用額外 +2% × 營收", dict(opxa="opx_adj_hi"))]
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.0"), ("D", "ail", 2027, "#,##0.0"), ("E", "gm", 2027, "0.0%"), ("F", "eps", 2027, "0.00"),
                           ("H", "rev", 2028, "#,##0.0"), ("I", "ail", 2028, "#,##0.0"), ("J", "gm", 2028, "0.0%"), ("K", "eps", 2028, "0.00")):
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
put_case(nn, "情境 A（全部）", "顆數係數 A＋液冷市場 A＋可拆式上蓋、份額、液冷份額高值", PB["A"], "PL!")
put_case(nn + 1, "情境 C（全部）", "顆數係數 C＋液冷市場 C＋低值", PB["C"], "PL!")
SENS_LAST = nn + 1; SENS_NOTE = nn + 3
SE.cell(nn + 3, 1, "讀法：基準為情境 B（顆數係數 B 從 CY2028 起）。每個案例的 FY2026 份額與其他營收毛利率依該案例重算，所以 FY2026 恆等於錨點。「2026 AI 占比」與「非 NVIDIA 份額比」只改校準：AI 占比高 → 份額高、其他散熱小，FY2027 起成長更集中在均熱片。")
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
    for c in "CDE": SE[f"{c}{i}"].number_format = "#,##0"
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
RC["A1"] = "Recon — 本模型 vs 共識、券商估計與奇鋐模型液冷份額。只作事後對照，不為對上數字而調參數。共識取自 stockanalysis（S&P Global）與 FactSet（鉅亨轉引，5 月）；券商多為單一來源"
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
rec("FY2026 營收 vs 共識（stockanalysis）", "=cons_rev26", pl("rev", 2026), "stockanalysis 2026-10-22（7 家）", "", fmt="#,##0.0")
rec("FY2026 EPS vs 共識（stockanalysis）", "=cons_eps26", pl("eps", 2026), "stockanalysis 2026-10-22（7 家）", "3Q26 財報（11-09）前")
rec("FY2026 EPS vs 亞系券商（08 月）", "=ab_eps26", pl("eps", 2026), "口袋學堂 2026-08-26 轉引", "")
rec("3Q26 營收 vs 大和（10-01）", "=dw_q3rev", lambda s: "=Q!D4", "大和（經濟日報 2026-10-01）", "實績（月營收）低於大和", fmt="#,##0.00")
rec("3Q26 EPS（模型）vs 大和（10-01）", "=dw_q3eps", lambda s: "=Q!D15", "大和（經濟日報 2026-10-01）", "Bloomberg 共識 20.37")
rec("FY2027 營收 vs FactSet 中位數（05-22）", "=fs_rev27", pl("rev", 2027), "FactSet（鉅亨 2026-05-22）", "較舊", fmt="#,##0.0")
rec("FY2027 EPS vs FactSet 中位數（05-22）", "=fs_eps27", pl("eps", 2027), "FactSet（鉅亨 2026-05-22）", "平均 121.01")
rec("FY2027 EPS vs 亞系券商（08 月）", "=ab_eps27", pl("eps", 2027), "口袋學堂 2026-08-26 轉引", "")
rec("FY2027 EPS vs 分析師（理財周刊 08-05）", "=mw_eps27", pl("eps", 2027), "理財周刊 2026-08-05", "")
rec("FY2027 營收年增 vs 本國投顧（10-08）（百分點）", "=ld_g27", pl("g", 2027), "經濟日報 2026-10-08", "", fmt="0.0%", pt=True)
rec("FY2028 營收 vs FactSet 中位數（05-22）", "=fs_rev28", pl("rev", 2028), "FactSet（鉅亨 2026-05-22）", "", fmt="#,##0.0")
rec("FY2028 EPS vs FactSet 中位數（05-22）", "=fs_eps28", pl("eps", 2028), "FactSet（鉅亨 2026-05-22）", "平均 191.02")
rec("FY2028 EPS vs 亞系券商（08 月）", "=ab_eps28", pl("eps", 2028), "口袋學堂 2026-08-26 轉引", "")
RC.cell(rr_, 1, "B. 均熱片與液冷份額（機制檢查）").font = BOLD; rr_ += 1
rec("FY2026 健策 NVIDIA 均熱片份額（校準）", None, lambda s: f"=PL!D{PB[s]['sh0']}", "—", "份額第二來源：未找到；公司稱全球前兩大 IHS 供應商", fmt="0.0%")
rec("CY2026 健策 AI 均熱片 ÷ 市場合計", None, lambda s: f"=PL!D{PB[s]['ail']}/fx/PL!D{PB[s]['mkt']}", "—", "", fmt="0.0%")
for y in (2027, 2028, 2029):
    rec(f"CY{y} 奇鋐＋健策液冷份額合計", None, lambda s, y=y: f"=avsh{s}_{y % 100}+PL!{FC[y]}{PB[s]['lcs']}", "avc/dist PL 頁（唯讀）",
        "遠低於 100%：其餘為雙鴻、酷碼、台達電、Boyd、CoolIT 等；雙鴻模型未在 main，未納入" if y == 2027 else "—", fmt="0.0%")
rec("CY2027 Rubin 每顆均熱片 vs 本國投顧假設（≥ US$200）", "=200", lambda s: f"=PL!E{PB[s]['pvr']}", "經濟日報 2026-10-08", "本國投顧假設 800–900 萬顆全採可拆式", fmt="#,##0")
rec("FY2027 每顆加速器健策 AI 均熱片營收（US$）", None, pl("pgu", 2027), "—", "", fmt="#,##0.0")
RC.cell(rr_, 1, "C. 校準").font = BOLD; rr_ += 1
rec("其他營收毛利率（Derived）", None, lambda s: f"=PL!D{PB[s]['gmn']}", "FY2026 錨點", "AI 均熱片毛利率＝此值＋d_ai", fmt="0.0%")
rec("FY2027 毛利率 vs 券商（逾 50%）（百分點）", "=0.5", pl("gm", 2027), "大和、本國投顧（Interested-party）", "", fmt="0.0%", pt=True)
rec("4Q26 EPS（模型）", None, lambda s: "=Q!E15", "—", "3Q26 財報 11-09 公布", fmt="0.00")
RC.cell(rr_, 1, "D. 目標價對帳（模型欄為情境加權值）").font = BOLD; rr_ += 1
RECV2 = rr_
for lab, ref, mdl, src in (("12 個月後的前瞻 EPS（情境加權）vs 共識前瞻 EPS", "=cons_feps", "=Val!B35", "In 頁 cons_feps（Derived）"),
                           ("目標價（DCF 與 Comparable 加權）vs 市場共識目標價", "=cons_tp", "=Val!B40", "stockanalysis 2026-10-22（11 位）"),
                           ("Comparable 目標價 vs 市場共識目標價", "=cons_tp", "=Val!B36", "同上")):
    RC.cell(rr_, 1, lab); RC.cell(rr_, 2, ref).number_format = "#,##0.00"; RC.cell(rr_, 3, mdl).number_format = "#,##0.00"
    RC.cell(rr_, 4, f"=C{rr_}/B{rr_}-1").number_format = "0.0%"; RC.cell(rr_, 9, src); RC.cell(rr_, 10, "共識只作對帳，不作校準"); rr_ += 1
RC.cell(rr_, 1, "E. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以部署路徑取代網通模型顆數"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RECON_LAST = rr_ - 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝健策財報、月營收與自結損益公告，經資料商與媒體轉述；券商與媒體估計皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("資料商：stockanalysis 季度損益（1Q25–2Q26 營收、毛利、營業利益、歸母淨利、EPS）", SAQ),
        ("資料商：HiStock 季度損益（稅前、稅後）與每股盈餘（2018–2026）", HISQ),
        ("二手：口袋學堂 2026-08-26（2Q26、1H26、2025 散熱占比 72.6%、2Q26 約 77%、7 月營收、亞系券商 EPS）", PKT),
        ("二手：口袋學堂 2025-12-15（全球前兩大 IHS 供應商、AMD／Intel 核心供應商、水冷模組出貨）", PKT2),
        ("二手：理財周刊 2026-08-05（各世代均熱片單價、1H26 產品比重、MCL 驗證）", MW),
        ("二手：工商時報 2026-08-07（6 月自結、VR 均熱片出貨量、加強環）", CT0807),
        ("二手：優分析 2026-10-05（8 月自結淨利 10.79 億、EPS 7.35；可拆式上蓋 4Q26 送樣）", UA1005),
        ("二手：technews 2026-10-08（9 月營收 30.54 億、前 9 月 220.73 億）", TN9),
        ("二手：technews 2026-08-07（7 月營收 32.13 億）", TN7),
        ("券商：大和（經濟日報 2026-10-01；目標價 10,000、3Q26 估計）", DW),
        ("券商：本國投顧（經濟日報 2026-10-08；2027 營收 +147%、可拆式上蓋 ASP、均熱片占八成以上）", UDN1008),
        ("資料商：FactSet（鉅亨 2026-05-22；EPS 與營收中位數、目標價 4,920）", FS05),
        ("資料商：stockanalysis（股數、前瞻 P/E、ROE、年度 P/E、共識目標價與 FY2026 預估；同業前瞻 P/E）", SAF),
        ("資料商：stockanalysis 比率頁（ROE、年度 P/E 2021–2025）", SAR),
        ("資料商：Goodinfo（年底收盤、EPS 2017–2025、ROE）", GI),
        ("資料商：聚財網（2016 年 EPS 2.32）", WEARN),
        ("資料商：財報狗（10-22 收盤 6,070、近 12 月 P/E）", SD),
        ("市場層：markets/network（加速器顆數）、tsmc（情境顆數係數）成品；奇鋐模型 avc/dist（液冷市場與奇鋐份額）（本 repo，唯讀）", ""),
        ("待補（人工）：3Q26 財報（11-09）；10 月營收與自結（11-10 前後）；7、9 月自結損益；AI 均熱片占比與各客戶份額；可拆式上蓋實際單價；液冷營收；資本支出與折舊；2016–2020 營收與淨利；前瞻 P/E 十年歷史（Bloomberg）；較新的 FactSet 共識", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 100; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("健策精密（3653.TW）收支與估值模型 v0.1.1（AI 半導體 Project：電源與散熱次產業；均熱片、加強環、液冷冷板元件；樣板＝奇鋐）　2026-10-25", True),
 ("版本：v0.1.1（2026-10-25：上游奇鋐 v0.2.1 液冷滲透率改依情境字母（奇鋐），數值變動：本模型讀奇鋐同字母情境液冷市場，情境 A、C 的液冷冷板元件營收與 EPS 隨之變動，情境 B 不變）。v0.1（初版）。加速器顆數取自 markets/network（× tsmc 同字母情境係數）；每顆均熱片金額、健策份額為本模型 Assumed（Units 頁）；液冷市場讀奇鋐模型 avc/dist 同字母情境，健策液冷份額 Assumed。", False),
 ("模型命題：健策的 AI 營收＝Σ 加速器顆數（Blackwell、Rubin、Rubin Ultra、AMD、ASIC）× 每顆均熱片／上蓋金額 × 健策份額，加上液冷冷板元件（奇鋐模型液冷市場 × 健策份額）。VR 世代每顆金額約為 Blackwell 的 3 倍，2027 年可拆式上蓋再提高單價；成長主要來自每顆金額，而非顆數。", False),
 ("核心問題：可拆式上蓋的滲透速度與單價、以及健策在新結構下的份額，能不能支撐 FY2027–28 EPS 倍增；現價約 59 倍前瞻 P/E（同業中位數約 26 倍）隱含的路徑合不合理。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029 由驅動因子推導，FY2030–FY2032 為收斂法。金額：十億新台幣；每顆金額與液冷市場以美元計、以匯率 32.0 換算。口徑：歸屬母公司淨利（IFRS）。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("AI 均熱片：markets/network Units 頁加速器顆數 × tsmc 同字母顆數係數（＝markets/memory SCEN_UNITS）× 每顆金額（Blackwell US$35、Rubin 2026 US$100 → 2027 150 → 2028 180、Rubin Ultra 200、AMD 40 → 80、ASIC 30 → 40；Interested-party／Assumed）× 健策份額（NVIDIA 份額 FY2026 校準，非 NVIDIA＝NVIDIA × 0.4）× 份額相對值。", False),
 ("液冷冷板元件：奇鋐模型 avc/dist PL 頁同字母情境的 AI 液冷市場（美元）× 健策份額（2027 0.8%、2028 1.5%、2029 2.5%；Assumed）。其他散熱（CPU 均熱片、交換器加強環、CPO、ILM、車用水冷）與導線架另設成長率。毛利率：AI 均熱片比其他營收高 d_ai 12pt（Assumed），其他毛利率為 FY2026 殘差（Derived）。", False),
 ("校準：FY2026 ＝ 1H26 財報＋3Q26 月營收實績（94.93 億）＋假設（3Q、4Q 毛利率 49%、營業費用率 8.5%、4Q 季增 10%）。3Q26 財報預計 2026-11-09，截至 2026-10-25 尚未公布。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，三情境權重 0.3／0.5／0.2，負值取 0）＋0.5 × Comparable 目標價（奇鋐、雙鴻、Vertiv、台達電前瞻 P/E 中位數 × 12 個月後前瞻 EPS；不放 A 股）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解、上檔空間級距。十年 P/E 規則值 × FY2027 EPS 保留為參考（不是目標價）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m，FY2032 達到；PL 頁灰底）。十年平均淨利率未找到，依工作單以情境 C FY2029 淨利率 × 0.8 替代（Assumed）。", False),
 ("情境與行業一致：A＝順利（tsmc 顆數係數 A、奇鋐液冷市場 A；可拆式上蓋單價、份額相對值、液冷份額取高值）；B＝平平（係數 B、液冷市場 B、基準）；C＝不如預期（係數 C、液冷市場 C、低值）；其他散熱、導線架各情境相同。可拆式上蓋滲透屬技術替代，依規則與行業同字母（A＝快）。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司數字取自財報與月營收（經資料商與媒體轉述）；券商與資料商為 Interested-party。", False),
 ("", False),
 ("已知限制：(1) AI 加速器均熱片占營收比（2026 40%）與各客戶份額未揭露，為 Assumed；(2) 每顆金額只有媒體轉述的單一來源（理財周刊），可拆式上蓋單價為 Assumed；(3) 液冷份額為新產品 Assumed；(4) 未建資本支出與折舊；(5) 共識前瞻 EPS 以 FactSet 5 月 FY2027 中位數推算，較舊；(6) 十年 P/E 2016–2020 只有一個來源；(7) 3Q26 財報後更新錨點；R4 產業部署路徑未建。", False)]
OBS = [
    "1. 加權目標價 NT$3,903，上檔空間 −35.7%（級距 −25%～−50%），主因是倍數而非 EPS：本模型 12 個月後前瞻 EPS 176.8 元比共識前瞻 EPS 107.7 元高 64%（EPS 因素 +64%），但同業前瞻 P/E 中位數只有 25.8x（奇鋐 27.0x、雙鴻 17.3x、Vertiv 24.5x、台達電 30.5x），市場共識目標價 7,204 元隱含 66.9x（倍數因素 +159%）；Comparable 未給健策均熱片寡占溢價，DCF（3,247 元，ROE 24% 使再投資率 16%）也低於 Comparable（4,558 元）。",
    "2. 可拆式上蓋（TR lid）單價是 EPS 最大的公司變數：基準假設 Rubin 每顆 2027 年 US$150（2026 年 US$100），若達本國投顧假設的 US$200，FY2027 EPS 由 147.8 升到 173.7 元（+17.5%），若只有 US$110 則降到 127.1 元（−14.0%）；追蹤 4Q26 送樣與 2027 年第一季量產（大和、本國投顧）以及 3Q26 財報（2026-11-09）的毛利率是否達本模型假設的 49%。",
    "3. 份額與液冷是次要變數：FY2026 以 AI 占比 40%（Assumed）校準出健策 NVIDIA 均熱片份額 88%；液冷冷板元件份額基準 2027 年 0.8%、2029 年 2.5%，與奇鋐模型份額（47%→42%）合計不到 50%，其餘為雙鴻、酷碼、台達電等；MCL 驗證與 MLCP 認證（市場傳言）是否轉為營收，可由 2027 年月營收與法說檢驗。",
]
lines[4:4] = [("投資建議：放空（依上檔空間固定規則；加權目標價 NT$3,903、上檔空間 −35.7%、級距 −25%～−50%；見 Val 第 0 節）", True)]
lines[5:5] = [("重點觀察（Andy 2026-10-25 14:00；數字見 Val 第 0 節）", True)] + [(t, False) for t in OBS] + [("", False)]
OBS_ROW = 6
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ---------------- 分部原則第 3 條（2026-10-28）：核對列與宣告（追加於工作表最後，不位移既有列，數值不變） ----------------
import os as _os2, sys as _sys2
_sys2.path.insert(0, _os2.path.join(_os2.path.dirname(_os2.path.dirname(_os2.path.abspath(__file__))), "tools", "segment_check"))
from addrow import add_check
add_check(Q, seg_rows=[17, 18, 19], total_row=4, cols="B", rnd=6, note="分部原則第 3 條（2026-10-28 追加）；分部宣告見 segments.json")
RD.cell(RD.max_row + 1, 1, "v0.1.2（2026-10-28）：分部原則重建（數值不變）。Q 頁最後追加核對列「檢查：分部合計 − 合併營收」；官方分部與子列宣告見 segments.json。")
if isinstance(RD.cell(1, 1).value, str): RD.cell(1, 1).value = RD.cell(1, 1).value.replace("v0.1.1", "v0.1.2")

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "jentech")
wb.save(sys.argv[1])
if os.environ.get("JT_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE,
               "UNOTE": UNOTE, "T": T, "TPS0": TPS0, "TPS_LAST": TPS_LAST, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0,
               "block_end": block_end, "OBS_ROW": OBS_ROW, "RECV2": RECV2},
              open(os.environ["JT_ROWMAP"], "w"), ensure_ascii=False, indent=1)

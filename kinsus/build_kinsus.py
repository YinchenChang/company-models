# 景碩（3189.TW）收支與估值模型 v0.1（AI 半導體 Project：載板與 PCB 次產業；ABF／BT 載板第二梯隊；樣板 nanya_pcb/、unimicron/）
# 用法：python3 kinsus/build_kinsus.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；EPS：新台幣（基本，歸屬母公司）。
# 景碩合併晶碩（Pegavision，6491，隱形眼鏡），晶碩淨利約 69% 屬非控制權益；歸屬母公司淨利＝合併淨利 − 非控制權益。
# AI 加速器載板的份額口徑與 Ibiden、欣興、南電一致：每顆 NVIDIA 載板金額、Ibiden 份額與非 NVIDIA 權重唯讀取自 ibiden/dist；欣興份額唯讀取自 unimicron/dist；南電份額唯讀取自 nanya_pcb/dist。
# 加速器顆數、AI IT 資本支出、交換晶片與交換器市場唯讀取自 markets/network/dist；情境顆數係數取自 tsmc/dist（與 markets/memory SCEN_UNITS 核對）；不修改上游。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
import cycle_position as CP
import valuation_rule as VR
import market as MEM
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）

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
NET_F = only("markets/network/dist/*.xlsx"); TSMC_F = only("tsmc/dist/*.xlsx"); IB_F = only("ibiden/dist/*.xlsx")
UN_F = only("unimicron/dist/*.xlsx"); NP_F = only("nanya_pcb/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
UNITS = {}
for r in openpyxl.load_workbook(NET_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and (str(k).startswith("u_") or str(k).startswith("g29_")):
        UNITS[k] = (r[4].value, r[1].value, f"{rel(NET_F)}｜In!E{r[0].row}")
ITC = None; SWM = None
for r in openpyxl.load_workbook(NET_F, data_only=True)["Mkt"].iter_rows(min_row=4):
    if r[0].value == "AI IT 資本支出（加速器 × 每顆 IT capex）":
        ITC = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
    if r[0].value == "交換晶片與交換器（scale-up＋scale-out）":
        SWM = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
assert ITC and all(v for v in ITC[0]), ITC
assert SWM and all(v for v in SWM[0]), SWM
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and str(k).startswith("um"):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
assert all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()
for s_ in SC:   # 情境與行業一致：tsmc 顆數係數必須等於 markets/memory 的 SCEN_UNITS（同字母）
    assert abs(TS[f"um28_{s_}"][0] - MEM.SCEN_UNITS[s_][0]) < 1e-9 and abs(TS[f"um29_{s_}"][0] - MEM.SCEN_UNITS[s_][1]) < 1e-9, s_
NV = ("bw", "rb", "ru", "hp"); OT = ("amd", "tpu", "trn", "oth")
assert all(f"u_{g}_{y}" in UNITS for g in NV + OT for y in (26, 27, 28)), sorted(UNITS)
ibw = openpyxl.load_workbook(IB_F, data_only=True)
IB = {}
for r in ibw["In"].iter_rows(min_row=4):
    k = r[0].value
    if k in ("w_ot", "g_asp", "rel_27", "rel_28", "rel_29", "pp_27", "pp_28", "pp_29"):
        IB[k] = (r[3].value, r[4].value, r[5].value, f"{rel(IB_F)}｜In!E{r[0].row}", r[1].value)
assert len(IB) == 8, IB.keys()
IB_V26 = None
for r in ibw["Units"].iter_rows(min_row=4, max_row=20):
    if str(r[0].value or "").startswith("每顆 NVIDIA 載板金額"):
        IB_V26 = (r[2].value, f"{rel(IB_F)}｜Units!C{r[0].row}", r[0].value)
assert IB_V26 and "70%" in IB_V26[2], IB_V26
UNS = {}
for r in openpyxl.load_workbook(UN_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and (str(k).startswith("sh_nv_") or str(k).startswith("sh_ot_")):
        UNS[k] = (r[4].value, f"{rel(UN_F)}｜In!E{r[0].row}")
assert all(f"sh_{t}_{y}" in UNS for t in ("nv", "ot") for y in (26, 27, 28, 29)), UNS.keys()
NPS = {}
for r in openpyxl.load_workbook(NP_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and (str(k).startswith("sh_nv_") or str(k).startswith("sh_ot_")):
        NPS[k] = (r[4].value, f"{rel(NP_F)}｜In!E{r[0].row}")
assert all(f"sh_{t}_{y}" in NPS for t in ("nv", "ot") for y in (26, 27, 28, 29)), NPS.keys()

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B）；顆數：百萬顆；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間，不當點估計（放情境或 Sens）。SemiAnalysis 未使用。公司數字經資料商（stockanalysis）與媒體轉載；公司說法為 Interested-party；券商、媒體、資料商為 Interested-party"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
SA = "stockanalysis.com（3189）"
SAQ = "stockanalysis 季損益表（公司財報轉載）"
SA6491 = "stockanalysis 晶碩（6491）季損益表"
TN9 = "TechNews 2026-10-06（9 月營收、前 9 月累計）"
CTEE8 = "工商時報 2026-08-06（7 月營收、資本支出 196 億）"
GS8 = "TechNews 2026-08-04（高盛；2Q26 法說重點）"
UDN8 = "經濟日報 2026-08-07（2Q26 獲利、產能規劃）"
UDN9 = "經濟日報 2026-09-16（投顧法人：Vera CPU、Rubin 份額、產能）"
UDN10 = "經濟日報 2026-10-12（美系券商 EPS、3Q／4Q26 毛利率）"
WR10 = "旺得富 2026-10-14（永豐、玉山投顧；8 月自結、4Q26 預估）"
UDN4 = "經濟日報 2026-04-23（投顧法人：CPU 份額、稼動率）"
MDJ1 = "MoneyDJ 2026-01-19（產品組合、2026 年資本支出）"
FG25 = "富果 2Q25 法說摘要（2025-08-19，AI 摘要）"
BRO = "brodoginvestment 景碩分析 2026-04-07（引述券商；Interested-party）"
CNYES_FS = "鉅亨 2026-06-20（FactSet 13 位分析師）"

sec("A. AI 加速器顆數、AI IT 資本支出與交換器市場（讀 markets/network、tsmc 成品；唯讀，不自行重估）")
for g in NV + OT:
    for y in (26, 27, 28):
        v, lab, src = UNITS[f"u_{g}_{y}"]
        inp(f"u_{g}_{y}", lab, "百萬顆", None, v, None, "Derived", src, "唯讀上游；與 Ibiden、欣興、南電、台積電模型同一份顆數")
    v, lab, src = UNITS[f"g29_{g}"]
    inp(f"g29_{g}", lab, "%", None, v, None, "Derived", src, "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}", "US$B", None, ITC[0][j], None, "Derived", ITC[1], "Units 頁每 GW 換算" if y == 26 else "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"swm_{y}", f"交換晶片與交換器市場 20{y}（scale-up＋scale-out）", "US$B", None, SWM[0][j], None, "Derived", SWM[1], "其他 AI 載板（網通 ASIC、交換器、FPGA）的成長依據" if y == 26 else "同上")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝順利（基準）；B＝平平；C＝不如預期；與台積電、Ibiden、欣興、南電模型相同（＝markets/memory SCEN_UNITS，建置時核對）" if s == "A" else "同上")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")

sec("B. 每顆加速器載板金額與份額（口徑與 Ibiden、欣興、南電一致；三者數字唯讀取自各自 dist）")
inp("v26", "FY2026 每顆 NVIDIA 加速器載板金額（全市場；Ibiden 模型校準）", "US$", None, round(IB_V26[0], 6), None, "Derived", IB_V26[1], "Ibiden 模型：Ibiden 每顆營收 ÷ 份額 70%（BofA，Interested-party）")
inp("g_asp", "每顆載板金額年增（面積、層數；Ibiden 模型）", "%", IB["g_asp"][0], IB["g_asp"][1], IB["g_asp"][2], "Analogy", IB["g_asp"][3], "Ibiden 路線圖 2026 年 90×90 mm → 2028 年 110×110 mm（經 Ibiden 模型）")
for y in (27, 28, 29):
    k = f"pp_{y}"
    inp(f"ibpp_{y}", f"ABF 價格溢價指數 20{y}（FY2026＝1；Ibiden 模型）", "x", IB[k][0], IB[k][1], IB[k][2], "Assumed", IB[k][3], "全市場短缺溢價；新產能到位後回落（Ibiden 模型；欣興、南電模型同值）" if y == 27 else "同上")
    k = f"rel_{y}"
    inp(f"ibrel_{y}", f"Ibiden NVIDIA 份額 20{y} ÷ 2026（Ibiden 模型）", "x", None, IB[k][1], None, "Assumed", IB[k][3], "Units 頁份額合計檢查用")
inp("ib_sh26", "Ibiden 占 NVIDIA 加速器載板 FY2026（Ibiden 模型假設）", "%", *SIL.vals("abf_ibsh26"), "Interested-party", IB_V26[1] + "（列標籤：假設 FY2026 份額 70%）", "BofA：Ibiden 70–80%")
inp("ib_w_ot", "Ibiden 非 NVIDIA 權重（份額 × 每顆金額，相對 NVIDIA；Ibiden 模型）", "x", IB["w_ot"][0], IB["w_ot"][1], IB["w_ot"][2], "Assumed", IB["w_ot"][3], "Units 頁：÷ v_ot 即 Ibiden 非 NVIDIA 份額")
inp("v_ot", "非 NVIDIA 加速器每顆載板金額 ÷ NVIDIA（欣興、南電模型同值）", "x", *SIL.vals("abf_vot"), "Assumed", "TPU、Trainium、MI 系列封裝面積接近 GPU；量化：未找到", "與欣興、南電模型 v_ot 相同")
for y in (26, 27, 28, 29):
    inp(f"un_nv_{y}", f"欣興占 NVIDIA 加速器載板 20{y}（欣興模型）", "%", None, UNS[f"sh_nv_{y}"][0], None, "Assumed", UNS[f"sh_nv_{y}"][1], "Units 頁份額合計檢查用")
    inp(f"un_ot_{y}", f"欣興占非 NVIDIA 加速器載板（金額）20{y}（欣興模型）", "%", None, UNS[f"sh_ot_{y}"][0], None, "Assumed", UNS[f"sh_ot_{y}"][1], "同上")
    inp(f"np_nv_{y}", f"南電占 NVIDIA 加速器載板 20{y}（南電模型）", "%", None, NPS[f"sh_nv_{y}"][0], None, "Assumed", NPS[f"sh_nv_{y}"][1], "同上")
    inp(f"np_ot_{y}", f"南電占非 NVIDIA 加速器載板（金額）20{y}（南電模型）", "%", None, NPS[f"sh_ot_{y}"][0], None, "Assumed", NPS[f"sh_ot_{y}"][1], "同上")
for y in (26, 27, 28, 29):
    inp(f"sh_nv_{y}", *SIL.row9(f"abf_ks_nv_{y}")[1:])
for y in (26, 27, 28, 29):
    inp(f"sh_ot_{y}", *SIL.row9(f"abf_ks_ot_{y}")[1:])
inp("cpu_r", "NVIDIA 自家 CPU（Grace／Vera）顆數 ÷ NVIDIA 加速器顆數", "x", 0.3, 0.4, 0.5, "Analogy",
    "GB200／VR200 NVL72 機櫃：每 2 顆 GPU 封裝配 1 顆 CPU（0.5）；HGX 系統用 x86 CPU（0）", "NVIDIA 機櫃與 HGX 出貨比例：未找到")
inp("v_cpu", "每顆 NVIDIA CPU 載板金額 ÷ 每顆 NVIDIA 加速器載板金額", "x", 0.20, 0.30, 0.45, "Assumed", "Vera CPU 為單晶片（88 核），封裝面積小於 GPU；量化：未找到", "—")
for y, v, lo, hi in ((26, 0.35, 0.25, 0.50), (27, 0.70, 0.50, 0.90), (28, 0.75, 0.50, 1.00), (29, 0.70, 0.45, 1.00)):
    inp(f"sh_cpu_{y}", f"景碩占 NVIDIA CPU（Grace／Vera）載板 20{y}", "%", lo, v, hi, "Interested-party" if y == 26 else "Assumed",
        UDN4 + "：美系 AI 客戶 CPU 供應份額由 10% 升到 50%" if y == 26 else (GS8 + "：公司法說表達目標拿下 Vera CPU 載板 100% 份額；" + UDN9 + "：2027 年 80–100%" if y == 27 else "同上；2028 年後第二供應商（Ibiden）可能回補，基準略降"),
        "年平均（上半年 10% 起跳）" if y == 26 else ("公司與券商說法 80–100%（Interested-party），基準 70% 為打折後的模型取值（2026-10-28 來源複核：改標 Assumed）" if y == 27 else "—"))

sec("C. FY2026 產品組合與 AI 占比（公司未按季揭露；券商與法說摘要推全年）")
inp("mx_abf", "FY2026 ABF 載板 ÷ 營收", "%", 0.36, 0.42, 0.46, "Interested-party", MDJ1 + "：ABF 約 30–40%、BT 約 35–40%、隱形眼鏡 15–20%", FG25 + "：2Q25 ABF 約 40%；2026 年 ABF 漲幅大於 BT（" + GS8 + "）")
inp("ai_abf", "FY2026 AI（加速器、NVIDIA CPU、網通 ASIC）÷ ABF 營收", "%", 0.18, 0.25, 0.32, "Interested-party", GS8 + "：2026 年 AI 相關約占總營收 10%、2027 年 15–20%", "10% ÷ ABF 42%≈24%（Derived）")

sec("D. 需求與產能（FY2027 起）")
inp("el_o", "其他 AI 載板（網通 ASIC、交換器、FPGA）對交換晶片與交換器市場的彈性", "x", 0.3, 0.5, 0.8, "Assumed", "其他 AI 載板＝FY2026 AI 載板 − 加速器與 NVIDIA CPU 載板（校準殘差）；隨交換器市場成長但低於等比例（單價另由溢價計）", "與南電模型同值")
inp("g_nai", "非 AI ABF（伺服器與 PC CPU、基地台 FPGA）年增（量）", "%", 0.02, 0.06, 0.10, "Assumed", BRO + "：景碩 CPU 覆晶載板市占長期約 50%（Intel、AMD）", UDN4 + "：AI 代理帶動一般伺服器 CPU 需求延伸到 2028 年")
for y in (27, 28, 29):
    inp(f"cap_{y}", *SIL.row9(f"abfcap_ks_{y}")[1:])
inp("g_mix", "每單位產能營收年增（層數、尺寸組合）", "%", 0.0, 0.05, 0.08, "Assumed", "AI 與伺服器 CPU 載板面積大、層數高；組合往 AI 移動", "與欣興、南電模型同值")
inp("carry", "ABF 漲價年度化（CY2026 基期含上半年低價；FY2027 起 ×(1＋此值)）", "%", 0.08, 0.15, 0.22, "Assumed",
    WR10 + "：4Q26 ABF 漲幅由 5–10% 上修到 15–20%、BT 每季 5–10%", GS8 + "：3Q、4Q26 ABF 與 BT 平均每季漲約 8%；4Q26 年化 ÷ FY2026 營收約 1.20（Q 頁，Derived；含量增）")
for s, vs in (("A", (0.28, 0.10, 0.05)), ("B", (0.20, 0.05, 0.03)), ("C", (0.12, 0.0, -0.03))):
    for y, v in zip((27, 28, 29), vs):
        inp(f"g_bt{y}_{s}", f"BT 與其他載板年增 FY20{y}（情境 {s}；公司特有，A 最有利）", "%", None, v, None, "Assumed",
            WR10 + "：BT 每季漲 5–10%；" + UDN9 + "：記憶體占 BT 需求逾五成、光模組為成長動能" if (s == "A" and y == 27) else "同上",
            CTEE8 + "：BT 漲價幅度大於 ABF，帶動組合改善" if (s == "A" and y == 27) else "—")
inp("g_lens", "晶碩（隱形眼鏡）營收年增", "%", 0.03, 0.10, 0.15, "Derived", SA6491 + "：1H26 40.0 億，年增 24%；FY2025 70.4 億，年增 3%", "三情境相同（與 AI 無關）")

sec("E. 毛利、折舊、費用與淨利（FY2027 起）")
inp("ft_abf", "ABF 營收增量的現金毛利率（扣除漲價溢價後）", "%", 0.40, 0.50, 0.60, "Assumed", "與欣興、南電模型同值", GS8 + "：2027–28 年毛利率每年約 +5pt（組合與稼動率）")
inp("ft_bt", "BT 與其他載板營收增量的現金毛利率", "%", 0.30, 0.45, 0.60, "Assumed", "BT 成長以漲價為主（漲價幾乎全數落在毛利）；量增部分毛利率較低", CTEE8)
inp("ft_lens", "晶碩營收增量的現金毛利率", "%", 0.35, 0.45, 0.55, "Assumed", "晶碩毛利率：未找到（本串未取得）", "—")
for y, v, lo, hi in ((27, 17.0, 10.0, 22.0), (28, 16.0, 9.0, 22.0), (29, 12.0, 7.0, 16.0)):
    inp(f"cx{y}", f"資本支出 FY20{y}（三情境相同；公司特有，含晶碩）", "NT$B", lo, v, hi, "Assumed",
        CTEE8 + "：K6 等新專案 196 億（三年，廠房）＋原三年設備計畫 235 億＝431 億" if y == 27 else "同上",
        GS8 + "：新廠設備由 3–4 家核心客戶直接贊助（若屬實，景碩帳上資本支出偏低值）；" + BRO + "：原計畫 2026–28 年 80／100／85 億" if y == 27 else "—")
inp("capex26", "FY2026 資本支出", "NT$B", 8.0, 11.0, 13.0, "Assumed", "1H26 實際 35.2 億（stockanalysis 現金流量表：1Q 13.6、2Q 21.6）", MDJ1 + "：2026 年約 80 億；8 月新增 196 億專案（三年）")
inp("da26", "FY2026 折舊攤提（1H26 實績＋下半年每季 21.5 億）", "NT$B", None, "=1.974+2.046+2*2.15", None, "Derived", "stockanalysis 現金流量表：1Q26 19.74、2Q26 20.46 億", "下半年每季 21.5 億（Assumed，延續季增）")
inp("life", "新增資產平均折舊年限", "年", 6, 8, 10, "Assumed", "載板設備與廠房混合；與欣興、南電模型同值", "會計年限：未找到")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.3, 0.6, 1.0, "Assumed", "1Q26 營業費用 14.59 億、2Q26 15.82 億（營收季增 13%）", "—")
for y in (27, 28, 29):
    inp(f"nonop{y}", f"營業外淨額 FY20{y}", "NT$B", 0.0, 0.4, 0.8, "Assumed",
        "1Q26 0.74 億、2Q26 1.05 億（stockanalysis 推得；現增後利息支出下降）" if y == 27 else "同上", "利息收入、匯兌（Assumed）" if y == 27 else "—")
inp("tax", "所得稅 ÷ 稅前利益（FY2027 起）", "%", 0.08, 0.12, 0.20, "Assumed", "FY2025 12.0%（stockanalysis：所得稅 3.70 ÷ 稅前 30.87 億）", "1H26 6.9%；營所稅 20%，獲利升高後投資抵減占比下降")
inp("sh_g", "股數年變動", "%", 0.0, 0.0, 0.01, "Assumed", "流通股數 5.2691 億股（stockanalysis；2026 年初現增 7,000 萬股，每股 145 元）", "股本 52.69 億（nstock）")
inp("nm_lens", "晶碩淨利率（非控制權益推算用）", "%", 0.21, 0.25, 0.27, "Derived", SA6491 + "：1H26 淨利 10.17 億 ÷ 營收 40.01 億＝25.4%", "FY2025 16.28 ÷ 70.39＝23.1%")
inp("nci_sh", "非控制權益 ÷ 晶碩淨利（景碩對晶碩的有效持股約 31%）", "%", None, "=(0.32657+0.37289)/(0.47228+0.54425)", None, "Derived",
    SAQ + "：1Q26、2Q26 非控制權益 3.27、3.73 億", SA6491 + "：1Q26、2Q26 晶碩淨利 4.72、5.44 億；持股比例原文：未找到")

sec("F. FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設）")
inp("q3_rev", "3Q26 營收（月營收實績：前 9 月累計 − 1H26）", "NT$B", None, "=38.141-11.105-12.496", None, "Verified", TN9 + "：前 9 月 381.41 億、9 月 51.81 億", CTEE8 + "：7 月 45.49 億")
inp("q3_gm", "3Q26 毛利率", "%", 0.278, 0.31, 0.33, "Assumed", WR10 + "：8 月自結 EPS 1.55、淨利率 17%；永豐 3Q26 EPS 4.29", UDN10 + "：美系券商 3Q26 毛利率 27.8%")
inp("q3_opx", "3Q26 營業費用", "NT$B", 1.60, 1.70, 1.80, "Assumed", "1Q26 1.459、2Q26 1.582（stockanalysis）", "—")
inp("q4_g", "4Q26 營收季增", "%", 0.08, 0.12, 0.19, "Assumed", GS8 + "：3Q、4Q26 各季增約 10%", WR10 + "：4Q26 營收 173.14 億（季增 19.1%）")
inp("q4_gm", "4Q26 毛利率", "%", 0.293, 0.32, 0.33, "Assumed", WR10 + "：4Q26 毛利率 33%", UDN10 + "：美系券商 29.3%")
inp("q4_opx", "4Q26 營業費用", "NT$B", 1.70, 1.80, 1.95, "Assumed", "同上", "—")
inp("q_nonop", "3Q、4Q26 營業外淨額（每季）", "NT$B", 0.0, 0.10, 0.20, "Assumed", "1Q26 0.074、2Q26 0.105（stockanalysis 推得）", "—")
inp("q_tax", "3Q、4Q26 所得稅 ÷ 稅前利益", "%", 0.06, 0.09, 0.12, "Derived", "1H26：所得稅 1.90 ÷ 稅前 27.49 億＝6.9%（stockanalysis）", "FY2025 12.0%")
inp("q3_lens", "3Q26 晶碩營收", "NT$B", 2.10, 2.20, 2.30, "Assumed", SA6491 + "：2Q26 2.115、1Q26 1.886", "3Q26 月營收：未找到")
inp("q4_lens", "4Q26 晶碩營收", "NT$B", 2.00, 2.10, 2.25, "Assumed", SA6491 + "：4Q25 1.930", "—")
inp("q_sh", "3Q、4Q26 加權股數", "十億股", None, 0.52691, None, "Interested-party", SA + "：流通股數 526.91M", "nstock：股本 52.69 億")

sec("G. 共識與券商（只作 Recon 與 Val 對照，不作校準）")
inp("cons_eps26", "共識 EPS FY2026（主串補位排序表，2026-10-22）", "NT$", None, 10.42, None, "Derived", "主串提供（補位排序表）；原始口徑本串未查證", CNYES_FS + "：平均 9.73（2026-06）")
inp("cons_eps27", "共識 EPS FY2027（主串補位排序表，2026-10-22）", "NT$", None, 26.47, None, "Derived", "主串提供（補位排序表）；原始口徑本串未查證", CNYES_FS + "：平均 19.93（2026-06）")
inp("fs_eps26", "FactSet 共識 EPS FY2026（平均，2026-06）", "NT$", 7.75, 9.73, 13.6, "Interested-party", CNYES_FS, "中位數 9.33")
inp("fs_eps27", "FactSet 共識 EPS FY2027（平均，2026-06）", "NT$", 13.39, 19.93, 33.58, "Interested-party", CNYES_FS, "中位數 18.45")
inp("fs_eps28", "FactSet 共識 EPS FY2028（平均，2026-06）", "NT$", 18.87, 34.39, 49.15, "Interested-party", CNYES_FS, "中位數 36.45")
inp("fs_rev27", "FactSet 共識營收 FY2027（平均，2026-06）", "NT$B", 61.862, 70.348, 87.849, "Interested-party", CNYES_FS, "—")
inp("fs_rev28", "FactSet 共識營收 FY2028（平均，2026-06）", "NT$B", 79.292, 93.549, 114.962, "Interested-party", CNYES_FS, "—")
inp("bk_eps27", "券商 EPS FY2027（永豐 36.03、玉山 42.41、美系 27.3 平均）", "NT$", 27.3, "=AVERAGE(36.03,42.41,27.3)", 42.41, "Interested-party", WR10, UDN10)
inp("bk_eps28", "券商 EPS FY2028（永豐 62.22、玉山 89.75、美系 51.5 平均）", "NT$", 51.5, "=AVERAGE(62.22,89.75,51.5)", 89.75, "Interested-party", WR10, UDN10)

sec("H. 價格與市場資料")
inp("px", "股價（2026-10-22 收盤）", "NT$", None, 986, None, "Interested-party", SA + "：986（2026-10-22 14:39）", "nstock：986")
inp("fpe", "前瞻 P/E（stockanalysis，價格日期 2026-10-22）", "x", None, 44.70, None, "Interested-party", SA, "口徑：頁面未明示（約未來 12 月）")
inp("pe_ttm", "近 12 月 P/E（stockanalysis，10-22）", "x", None, 166.14, None, "Interested-party", SA, "近 12 月 EPS 5.93")
inp("cons_feps", "共識前瞻 12 月 EPS（stockanalysis 頁面價格 986 ÷ 前瞻 P/E）", "NT$", None, "=986/fpe", None, "Derived", SA, "直接數字：未找到")
inp("cons_tp", "市場共識目標價（賣方平均）", "NT$", None, 1092.15, None, "Interested-party", SA + "：13 位分析師平均 1,092.15（頁面價格日期 2026-10-22）", "10 月：永豐 1,280、玉山 1,270（" + WR10 + "）；高盛 1,125（2026-08-04）")
inp("fx", "匯率", "NT$／US$", None, SIL.base("fx_twd"), None, "Assumed", "Siliconomics SIF_fx_twd（全庫共用；v0.8.2 前本模型取 32.5）", "—")
REV5 = [33.336, 42.441, 26.832, 30.535, 39.351]
NI5 = [3.859, 6.977, 0.04752, 0.04889, 1.596]
inp("nim5", "近五年平均淨利率（FY2021–25；十年平均的參考）", "%", None, "=AVERAGE(" + ",".join(f"{n}/{v}" for n, v in zip(NI5, REV5)) + ")", None, "Derived",
    SA + "：營收與歸屬母公司淨利 FY2021–25", "十年（2016–2025）營收：未找到（FY2016–20 營收未取得），收斂法依工作單改用 FY2029 × 0.8")

sec("I. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_r", "折現率／要求報酬", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%，美台日韓同一值）", "—")
inp("tp_g", "長期成長率", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "永續價值與 FY2032 收斂目標（情境 B）")
inp("tp_roe", "ROE（近三年平均，FY2023–25）", "%", None, "=AVERAGE(0.03,0.0337,0.0669)", None, "Derived", SA + " 比率頁：FY2023 3.00%、FY2024 3.37%、FY2025 6.69%", "單一來源；本模型 DCF 改用（淨利＋折舊－資本支出），此值只作參考列")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（工作單 2026-10-25 11:53 定 0.5，DCF／Comparable 各半）", "—")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（FY2029 與十年平均的中點）", "%", None, 0, None, "Assumed", "十年平均淨利率：未找到，以情境 B FY2029 × 0.8 代替（工作單規則）；本列＝（FY2029＋FY2029 × 0.8）÷ 2", "近五年平均見 nim5（只作參考）")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（＝十年平均）", "%", None, 0, None, "Assumed", "十年平均淨利率：未找到，以情境 C FY2029 淨利率 × 0.8 代替（工作單規則）", "—")
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
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    INROW[key] = r; r += 1
for k in TP_KEYS:
    for c in range(1, 10): IN.cell(INROW[k], c).fill = KEY
for col, w in zip("ABCDEFGHI", [12, 60, 10, 9, 11, 9, 18, 80, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — FY2025 實績、1Q26／2Q26 實績、3Q26 營收實績與 FY2026 錨點（十億新台幣；EPS 新台幣、基本；股數十億股）。藍字＝輸入；黑字＝公式"
Q["A2"] = "FY2026 錨點＝1Q、2Q26 財報實績＋3Q26 營收實績（月營收）與 3Q26 毛利率、費用假設＋4Q26 假設。3Q26 財報約 10 月底～11 月初公布（stockanalysis 估 2026-11-06），屆時以實績取代 E 欄（毛利率、費用、業外、淨利）"
hdr(Q, 3, ["項目", "FY2025A", "1Q26（實績）", "2Q26（實績）", "3Q26（營收實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1（公司）", "來源2／說明"])
S1 = "stockanalysis 年度與季損益表（公司財報轉載）"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2="—", ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"G{r}"] = ann or f"=SUM(C{r}:F{r})"; Q[f"G{r}"].number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, s1); Q.cell(r, 10, s2)
qrow(4, "營收", {"B": 39.351, "C": 11.105, "D": 12.496, "E": "=q3_rev", "F": "=E4*(1+q4_g)"}, tag="Verified／Assumed（4Q）", s2="FY2025 393.51 億（nstock 同）；1Q26 111.05 億（brodog 同）；1H26＝前 9 月 381.41 − 3Q（TechNews）")
qrow(5, "毛利", {"B": 8.314, "C": 2.349, "D": 3.261, "E": "=E4*q3_gm", "F": "=F4*q4_gm"}, tag="Verified／Assumed", s2="2Q26 毛利率 26.09%（nstock 同；經濟日報 26.1%）")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEFG"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2="FY2025 21.1%（brodog 同）")
qrow(7, "營業費用", {c: f"={c}5-{c}8" for c in "BCD"} | {"E": "=q3_opx", "F": "=q4_opx"}, tag="Derived／Assumed", s1="毛利 − 營業利益")
qrow(8, "營業利益", {"B": 2.670, "C": 0.89064, "D": 1.679, "E": "=E5-E7", "F": "=F5-F7"}, tag="Verified／Derived", s2="2Q26 營業利益率 13.43%（nstock 同）")
qrow(9, "營業外淨額", {c: f"={c}10-{c}8" for c in "BCD"} | {"E": "=q_nonop", "F": "=q_nonop"}, tag="Derived／Assumed", s1="稅前利益 − 營業利益")
qrow(10, "稅前利益", {"B": 3.087, "C": 0.96446, "D": 1.784, "E": "=E8+E9", "F": "=F8+F9"}, tag="Verified／Derived")
qrow(11, "所得稅", {"B": 0.36988, "C": 0.08866, "D": 0.10087, "E": "=E10*q_tax", "F": "=F10*q_tax"}, tag="Verified／Assumed")
qrow(12, "合併淨利（含非控制權益）", {c: f"={c}10-{c}11" for c in "BCDEF"}, tag="Derived", s1="稅前 − 所得稅", s2="stockanalysis：FY2025 27.17、1Q26 8.76、2Q26 16.83 億")
qrow(13, "非控制權益（主要為晶碩）", {c: f"={c}12-{c}14" for c in "BCD"} | {"E": "=E19*nm_lens*nci_sh", "F": "=F19*nm_lens*nci_sh"}, tag="Derived", s1="合併淨利 − 歸屬母公司淨利；3Q、4Q26＝晶碩營收 × 晶碩淨利率 × 非控制比例",
     s2="stockanalysis：FY2025 11.21、1Q26 3.27、2Q26 3.73 億")
qrow(14, "歸屬母公司淨利", {"B": 1.596, "C": 0.54923, "D": 1.311, "E": "=E12-E13", "F": "=F12-F13"}, tag="Verified／Derived", s2="WR10：8 月自結 EPS 1.55".replace("WR10", WR10))
qrow(15, "加權股數（十億股）", {"B": "=B14/B16", "C": "=C14/C16", "D": "=D14/D16", "E": "=q_sh", "F": "=q_sh"}, ann="=AVERAGE(C15:F15)", tag="Derived", s1="淨利 ÷ EPS", s2="2026 年初現增 7,000 萬股（每股 145 元）")
qrow(16, "EPS（基本）", {"B": 3.51, "C": 1.17, "D": 2.57, "E": "=E14/E15", "F": "=F14/F15"}, fmt="0.00", ann="=G14/G15", tag="Verified／Derived", s2="FY2025 3.51（wantgoo 同）；1H26 3.74（nstock、經濟日報同）")
qrow(17, "折舊攤提（年度）", {"B": "=1.709+1.772+1.787+1.890", "G": "=da26"}, ann="", tag="Verified／Derived", s1="stockanalysis 季現金流量表（四季加總）", s2="EBITDA 口徑 71.08 億")
qrow(18, "資本支出（年度）", {"B": "=2.368+1.277+1.080+1.391", "G": "=capex26"}, ann="", tag="Verified／Assumed", s1="FY2025：stockanalysis 季現金流量表（四季加總）；FY2026：In 頁 capex26", s2="—")
qrow(19, "晶碩營收（合併於景碩）", {"B": 7.039, "C": 1.886, "D": 2.115, "E": "=q3_lens", "F": "=q4_lens"}, tag="Verified／Assumed", s1=SA6491, s2="FY2025 70.39 億（優分析同）")
Q.cell(21, 1, "FY2026 產品別（錨點營收 × In 頁 C 節占比；晶碩取實績＋假設）").font = BOLD
for i, (lab, f) in enumerate((("ABF 載板", "=G4*mx_abf"), ("BT 與其他載板（含類載板）", "=G4-G22-G24"), ("晶碩（隱形眼鏡）", "=G19")), 22):
    Q.cell(i, 1, lab); Q[f"G{i}"] = f; Q[f"G{i}"].number_format = "#,##0.000"; Q.cell(i, 8, "Derived")
Q.cell(26, 1, "檢查（參考；校準未使用券商數字）").font = BOLD
checks = [("1H26 營收 自算 − （前 9 月累計 38.141 − 3Q26）（應為 0）", "=C4+D4-(38.141-q3_rev)", "0.000"),
          ("FY2026 營收年增", "=G4/B4-1", "0.0%"),
          ("FY2026 歸屬母公司淨利 ÷ 合併淨利", "=G14/G12", "0.0%"),
          ("FY2026 EPS ÷ 永豐 13.21 − 1", "=G16/13.21-1", "0.0%"),
          ("3Q26 EPS ÷ 永豐估 4.29 − 1", "=E16/4.29-1", "0.0%"),
          ("4Q26 EPS ÷ 永豐估 5.39 − 1", "=F16/5.39-1", "0.0%"),
          ("4Q26 營收年化 ÷ FY2026 營收（漲價年度化 carry 的對照）", "=F4*4/G4", "0.00")]
for i, (lab, f, fm) in enumerate(checks, 27):
    Q.cell(i, 1, lab); Q[f"B{i}"] = f; Q[f"B{i}"].number_format = fm
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 7), ("oi", 8), ("nonop", 9), ("nic", 12), ("nci", 13), ("ni", 14), ("sh", 15), ("eps", 16), ("da", 17), ("cx", 18), ("lens", 19)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$G${rr}")
for nm, rr in (("abf", 22), ("bt", 23)):
    defname(f"a26_{nm}", f"Q!$G${rr}")
defname("a26_shend", "Q!$F$15")
Q.column_dimensions["A"].width = 64
for c in "BCDEFG": Q.column_dimensions[c].width = 14
Q.column_dimensions["H"].width = 26; Q.column_dimensions["I"].width = 60; Q.column_dimensions["J"].width = 64
Q.freeze_panes = "B4"

# ---------------- 引擎 ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G", 2030: "H", 2031: "I", 2032: "J"}
FADE = (2030, 2031, 2032); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
BU = (2026, 2027, 2028, 2029); F3 = (2027, 2028, 2029)
def units_f(g, y):
    return f"u_{g}_{y % 100}" if y <= 2028 else f"u_{g}_28*(1+g29_{g})"
DEFP = dict(scn="A", ga="g_asp", pp="ibpp_{y}", vot="v_ot", shnv="sh_nv_{y}", shot="sh_ot_{y}", shcpu="sh_cpu_{y}", cpur="cpu_r", vcpu="v_cpu",
            aiabf="ai_abf", elo="el_o", gnai="g_nai", cap="cap_{y}", gmix="g_mix", carry="carry", gbt="g_bt{y}_{s}", glens="g_lens",
            ftabf="ft_abf", ftbt="ft_bt", ftlens="ft_lens", cx="cx{y}", life="life", phi="phi", nonop="nonop{y}", tax="tax", nml="nm_lens",
            mf="1", fg="fade_g{s}", fm="fade_m{s}")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    p = lambda k, y=None: P[k].format(y=(y % 100 if y else ""), s=s)
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.00", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 11, note); R[key] = r; r += 1
    def sec_(t):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
    um = lambda y: {2026: "1", 2027: "1", 2028: f"um28_{s}", 2029: f"um29_{s}"}[y]
    mf = lambda y: "" if y == 2026 else f"*{P['mf']}"
    sec_("A. AI 加速器與 NVIDIA CPU 載板（口徑與 Ibiden、欣興、南電一致；D 欄＝FY2026 校準年）")
    row("unv", "NVIDIA 加速器顆數 × 情境係數", "百萬顆", {FC[y]: "=(" + "+".join(units_f(g, y) for g in NV) + f")*{um(y)}{mf(y)}" for y in BU}, note="markets/network In：Blackwell、Rubin、Rubin Ultra、Hopper")
    row("uot", "非 NVIDIA 加速器顆數 × 情境係數", "百萬顆", {FC[y]: "=(" + "+".join(units_f(g, y) for g in OT) + f")*{um(y)}{mf(y)}" for y in BU}, note="AMD、TPU、Trainium、其他 ASIC")
    row("ucpu", "NVIDIA CPU（Grace／Vera）顆數 ＝ NVIDIA 加速器 × CPU 比例", "百萬顆", {FC[y]: f"={FC[y]}{R['unv']}*{P['cpur']}" for y in BU}, note="In 頁 cpu_r（Analogy：NVL72 每 2 顆 GPU 配 1 顆 CPU）")
    row("idx", "每顆載板面積與層數指數（FY2026＝1）", "x", {"D": "=1", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['ga']})" for y in F3}}, "0.000", note="Ibiden 模型 g_asp（Analogy）")
    row("pp", "ABF 價格溢價指數（全市場；FY2026＝1）", "x", {"D": "=1", **{FC[y]: f"={p('pp', y)}" for y in F3}}, "0.00", note="Ibiden 模型 pp（Assumed）")
    row("cy", "景碩漲價年度化（CY2026 基期；FY2026＝1）", "x", {"D": "=1", **{FC[y]: f"=1+{P['carry']}" for y in F3}}, "0.00", note="In 頁 carry（Assumed）")
    row("v", "每顆 NVIDIA 加速器載板金額（全市場）＝ FY2026 金額 × 面積層數 × 溢價", "US$／顆", {FC[y]: f"=v26*{FC[y]}{R['idx']}*{FC[y]}{R['pp']}" for y in BU}, "#,##0.0", note="FY2026 金額取自 Ibiden 模型 Units 頁（權威）")
    row("shnv", "景碩份額：NVIDIA 加速器（GPU）載板", "%", {FC[y]: f"={p('shnv', y)}" for y in BU}, "0.0%")
    row("shot", "景碩份額：非 NVIDIA 加速器載板（金額）", "%", {FC[y]: f"={p('shot', y)}" for y in BU}, "0.0%")
    row("shcpu", "景碩份額：NVIDIA CPU 載板", "%", {FC[y]: f"={p('shcpu', y)}" for y in BU}, "0.0%")
    row("accu", "景碩加速器與 CPU 載板營收 ＝（NVIDIA × 金額 × 份額 ＋ 非 NVIDIA × 金額 × v_ot × 份額 ＋ CPU × 金額 × v_cpu × 份額）× 漲價年度化", "US$M",
        {FC[y]: f"=({FC[y]}{R['unv']}*{FC[y]}{R['v']}*{FC[y]}{R['shnv']}+{FC[y]}{R['uot']}*{FC[y]}{R['v']}*{P['vot']}*{FC[y]}{R['shot']}+{FC[y]}{R['ucpu']}*{FC[y]}{R['v']}*{P['vcpu']}*{FC[y]}{R['shcpu']})*{FC[y]}{R['cy']}" for y in BU}, "#,##0")
    row("acc", "景碩加速器與 CPU 載板營收（新台幣）", "NT$B", {FC[y]: f"={FC[y]}{R['accu']}*fx/1000" for y in BU}, "#,##0.00")
    sec_("B. ABF 需求與產能（營收＝MIN（需求, 產能））")
    row("ai0", "AI 載板需求（FY2026＝ABF × AI 占比）", "NT$B", {"D": f"=a26_abf*{P['aiabf']}"}, "#,##0.00", note="高盛：2026 年 AI 約占總營收 10%")
    row("oai", "其他 AI 載板需求（網通 ASIC、交換器、FPGA）", "NT$B", {"D": f"=D{R['ai0']}-D{R['acc']}",
        **{FC[y]: f"=$D{r}*(swm_{y % 100}/swm_26)^{P['elo']}*{um(y)}{mf(y)}*{FC[y]}{R['pp']}*{FC[y]}{R['cy']}" for y in F3}}, "#,##0.00",
        note="FY2026＝校準殘差（Derived）；之後 × 交換晶片與交換器市場成長 ^ 彈性 × 情境係數 × 溢價")
    row("nai", "非 AI ABF 需求（伺服器與 PC CPU、基地台 FPGA；含漲價年度化）", "NT$B", {"D": f"=a26_abf-D{R['ai0']}", **{FC[y]: f"=$D{r}*(1+{P['gnai']})^{y - 2026}*{FC[y]}{R['cy']}" for y in F3}}, "#,##0.00",
        note="景碩 4Q26 ABF 漲價為全面性（非只 AI），故非 AI 也乘漲價年度化；不乘全市場溢價")
    row("dem", "ABF 需求合計", "NT$B", {FC[y]: f"={FC[y]}{R['acc']}+{FC[y]}{R['oai']}+{FC[y]}{R['nai']}" for y in BU}, "#,##0.00")
    row("cap", "ABF 產能上限（FY2026 營收 × 產能指數 × 組合 × 溢價 × 漲價年度化）", "NT$B", {"D": "=a26_abf", **{FC[y]: f"=$D{R['dem']}*{p('cap', y)}*(1+{P['gmix']})^{y - 2026}*{FC[y]}{R['pp']}*{FC[y]}{R['cy']}" for y in F3}}, "#,##0.00",
        note="2026 年 ABF 稼動率約 85–95%，以 FY2026 營收為產能基準")
    row("abf", "ABF 載板營收 ＝ MIN（需求, 產能）", "NT$B", {"D": "=a26_abf", **{FC[y]: f"=MIN({FC[y]}{R['dem']},{FC[y]}{R['cap']})" for y in F3}}, "#,##0.00")
    row("fill", "需求滿足率（營收 ÷ 需求；< 100% 表示產能受限）", "%", {FC[y]: f"={FC[y]}{R['abf']}/{FC[y]}{R['dem']}" for y in BU}, "0.0%")
    row("aisub", "AI 載板營收（加速器、CPU＋其他 AI，乘滿足率）", "NT$B", {FC[y]: f"=({FC[y]}{R['acc']}+{FC[y]}{R['oai']})*{FC[y]}{R['fill']}" for y in BU}, "#,##0.00")
    row("accsh", "加速器與 CPU 載板 ÷ AI 載板", "%", {FC[y]: f"={FC[y]}{R['acc']}/({FC[y]}{R['acc']}+{FC[y]}{R['oai']})" for y in BU}, "0.0%")
    sec_("C. 其他產品與合計")
    row("bt", "BT 與其他載板（情境別年增；公司特有）", "NT$B", {"D": "=a26_bt", **{FC[y]: f"={FC[y - 1]}{r}*(1+{p('gbt', y)})" for y in F3}}, "#,##0.00", note="BT 成長以漲價為主；記憶體占 BT 需求逾五成")
    row("lens", "晶碩（隱形眼鏡）", "NT$B", {"C": "=a25_lens", "D": "=a26_lens", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['glens']})" for y in F3}}, "#,##0.00", note="與 AI 無關；三情境相同")
    row("rev", "營收", "NT$B", {"C": "=a25_rev", **{FC[y]: f"={FC[y]}{R['abf']}+{FC[y]}{R['bt']}+{FC[y]}{R['lens']}" for y in BU}}, "#,##0.0")
    row("segchk", "檢查：分部合計（ABF＋BT 與其他＋晶碩）− 合併營收（FY2026 錨點起）", "NT$B", {FC[y]: f"={FC[y]}{R['abf']}+{FC[y]}{R['bt']}+{FC[y]}{R['lens']}-{FC[y]}{R['rev']}" for y in BU}, "0.000")
    row("g", "營收年增", "%", {FC[y]: f"={FC[y]}{R['rev']}/{FC[y - 1]}{R['rev']}-1" for y in FYS[1:]}, "0.0%")
    row("abfsh", "ABF ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['abf']}/{FC[y]}{R['rev']}" for y in BU}, "0.0%")
    row("aish", "AI 載板 ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['aisub']}/{FC[y]}{R['rev']}" for y in BU}, "0.0%", note="高盛：2026 年約 10%、2027 年 15–20%（公司口徑）")
    sec_("D. 折舊、毛利與損益")
    row("cx", "資本支出", "NT$B", {"C": "=a25_cx", "D": "=capex26", **{FC[y]: f"={p('cx', y)}" for y in F3}}, "#,##0.0")
    row("da", "折舊攤提 ＝ 前一年 ×（1 − 1／年限）＋ 前一年資本支出 ÷ 年限", "NT$B", {"C": "=a25_da", "D": "=da26",
        **{FC[y]: f"={FC[y - 1]}{r}*(1-1/{P['life']})+{FC[y - 1]}{R['cx']}/{P['life']}" for y in F3}}, "#,##0.0", note="新產能前一年投入、當年開始折舊（近似）")
    row("dar", "折舊 ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['da']}/{FC[y]}{R['rev']}" for y in FYS}, "0.0%")
    row("prem", "ABF 漲價溢價營收（相對 FY2026 價格；全數落在毛利）", "NT$B", {FC[y]: f"={FC[y]}{R['fill']}*(({FC[y]}{R['acc']}+{FC[y]}{R['oai']})*(1-1/({FC[y]}{R['pp']}*{FC[y]}{R['cy']}))+{FC[y]}{R['nai']}*(1-1/{FC[y]}{R['cy']}))" for y in F3}, "#,##0.00",
        note="AI 載板 ×（1 − 1／（溢價 × 漲價年度化））＋ 非 AI ABF ×（1 − 1／漲價年度化），乘滿足率")
    row("gp", "毛利 ＝ FY2026 ＋ ABF 現金毛利增量 ＋ BT 現金毛利增量 ＋ 晶碩現金毛利增量 ＋ 溢價 − 折舊增量", "NT$B", {"C": "=a25_gp", "D": "=a26_gp",
        **{FC[y]: (f"=$D{r}+{P['ftabf']}*({FC[y]}{R['abf']}-{FC[y]}{R['prem']}-$D{R['abf']})+{P['ftbt']}*({FC[y]}{R['bt']}-$D{R['bt']})+{P['ftlens']}*({FC[y]}{R['lens']}-$D{R['lens']})"
                   f"+{FC[y]}{R['prem']}-({FC[y]}{R['da']}-$D{R['da']})") for y in F3}}, "#,##0.00", note="流通率為現金毛利口徑；折舊另計")
    row("gm", "毛利率", "%", {FC[y]: f"={FC[y]}{R['gp']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("opx", "營業費用", "NT$B", {"C": "=a25_opx", "D": "=a26_opx", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['phi']}*{FC[y]}{R['g']})" for y in F3}}, "#,##0.00")
    row("oi", "營業利益", "NT$B", {FC[y]: f"={FC[y]}{R['gp']}-{FC[y]}{R['opx']}" for y in FYS}, "#,##0.00")
    row("oim", "營業利益率", "%", {FC[y]: f"={FC[y]}{R['oi']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("nonop", "營業外淨額", "NT$B", {"C": "=a25_nonop", "D": "=a26_nonop", **{FC[y]: f"={p('nonop', y)}" for y in F3}}, "#,##0.00")
    row("pti", "稅前利益", "NT$B", {FC[y]: f"={FC[y]}{R['oi']}+{FC[y]}{R['nonop']}" for y in FYS}, "#,##0.00")
    row("tax", "所得稅", "NT$B", {FC[y]: f"={FC[y]}{R['pti']}*{P['tax']}" for y in F3}, "#,##0.00")
    row("nic", "合併淨利（含非控制權益）", "NT$B", {"C": "=a25_nic", "D": "=a26_nic", **{FC[y]: f"={FC[y]}{R['pti']}-{FC[y]}{R['tax']}" for y in F3}}, "#,##0.00")
    row("nci", "非控制權益 ＝ 晶碩營收 × 晶碩淨利率 × 非控制比例", "NT$B", {"C": "=a25_nci", "D": "=a26_nci", **{FC[y]: f"={FC[y]}{R['lens']}*{P['nml']}*nci_sh" for y in F3}}, "#,##0.00",
        note="景碩對晶碩的有效持股約 31%（In 頁 nci_sh，Derived）")
    row("ni", "歸屬母公司淨利", "NT$B", {"C": "=a25_ni", "D": "=a26_ni", **{FC[y]: f"={FC[y]}{R['nic']}-{FC[y]}{R['nci']}" for y in F3}}, "#,##0.00")
    row("shr", "加權股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", "E": "=a26_shend*(1+sh_g)", **{FC[y]: f"={FC[y - 1]}{r}*(1+sh_g)" for y in (2028, 2029)}}, "0.0000")
    row("eps", "EPS（基本）", "NT$", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['shr']}" for y in FYS}, "0.00")
    row("epsg", "EPS 年增", "%", {FC[y]: f"={FC[y]}{R['eps']}/{FC[y - 1]}{R['eps']}-1" for y in FYS[1:]}, "0.0%")
    row("nim", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("cxr", "資本支出 ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['cx']}/{FC[y]}{R['rev']}" for y in FYS}, "0.0%")
    row("fcf", "股東現金流 ＝ 淨利 ＋ 折舊 − 資本支出", "NT$B", {FC[y]: f"={FC[y]}{R['ni']}+{FC[y]}{R['da']}-{FC[y]}{R['cx']}" for y in BU}, "#,##0.00",
        note="不含營運資金變動；折舊與資本支出為合併數（含晶碩），淨利為歸屬母公司")
    row("fps", "每股股東現金流", "NT$", {FC[y]: f"={FC[y]}{R['fcf']}/{FC[y]}{R['shr']}" for y in BU}, "0.00")
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
        ws[f"{c}{R['dar']}"] = f"=$G{R['dar']}"
        ws[f"{c}{R['da']}"] = f"={c}{R['rev']}*{c}{R['dar']}"
        ws[f"{c}{R['cxr']}"] = f"=$G{R['cxr']}+($G{R['dar']}*(1+tp_g)^({P['life']}/2)-$G{R['cxr']})*{k}/3"
        ws[f"{c}{R['cx']}"] = f"={c}{R['rev']}*{c}{R['cxr']}"
        ws[f"{c}{R['fcf']}"] = f"={c}{R['ni']}+{c}{R['da']}-{c}{R['cx']}"
        ws[f"{c}{R['fps']}"] = f"={c}{R['fcf']}/{c}{R['shr']}"
        for key, fm in (("g", "0.0%"), ("rev", "#,##0.0"), ("nim", "0.00%"), ("ni", "#,##0.00"), ("shr", "0.0000"), ("eps", "0.00"), ("epsg", "0.0%"),
                        ("dar", "0.0%"), ("da", "#,##0.0"), ("cxr", "0.0%"), ("cx", "#,##0.0"), ("fcf", "#,##0.00"), ("fps", "0.00")):
            ws[f"{c}{R[key]}"].number_format = fm; ws[f"{c}{R[key]}"].fill = FADEFILL
    ws.cell(R["rev"], 11, "FY2030 起為收斂法：營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m（FY2032 達到），非驅動因子推導")
    ws.cell(R["cxr"], 11, "FY2030 起：資本支出率由 FY2029 線性收斂到 折舊率 ×（1＋tp_g）^（年限／2）（穩態更新）；折舊率維持 FY2029")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 景碩損益（十億新台幣；EPS 新台幣、基本；歸屬母公司）。FY2025A 實績、FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設；三情境相同），FY2027–FY2029 由驅動因子推導，FY2030 起為收斂法。三情境各一個完整推導區塊（本頁下方），頂端為摘要"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A（順利）：tsmc 顆數係數 A（基準）；BT 與其他載板年增 28%／10%／5%",
         "B": "情境 B（平平）：tsmc 顆數係數 B（CY2028 ×0.8、CY2029 ×0.85）；BT 與其他載板年增 20%／5%／3%",
         "C": "情境 C（不如預期）：tsmc 顆數係數 C（CY2028–29 ×0.6）；BT 與其他載板年增 12%／0%／−3%"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
sr = 4; SUMR = {}
SUMK = (("rev", "營收", "NT$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("abf", "  ABF 載板", "NT$B", "#,##0.0"), ("aisub", "  AI 載板", "NT$B", "#,##0.0"),
        ("fill", "  ABF 需求滿足率", "%", "0.0%"), ("gm", "毛利率", "%", "0.00%"), ("da", "  折舊攤提", "NT$B", "#,##0.0"), ("oim", "營業利益率", "%", "0.00%"),
        ("ni", "歸屬母公司淨利", "NT$B", "#,##0.0"), ("nim", "  淨利率", "%", "0.00%"), ("eps", "EPS", "NT$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%"))
LONGK = ("rev", "g", "ni", "nim", "eps", "epsg", "da")
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in SUMK:
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS + (FADE if key in LONGK else ()):
            if key in ("g", "epsg") and y == 2025: continue
            if key in ("abf", "aisub", "fill") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
            if y in FADE: PL[f"{FC[y]}{sr}"].fill = FADEFILL
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
IN.cell(INROW["fade_mA"], 5, f"=PL!G{SUMR[('A', 'nim')]}")
IN.cell(INROW["fade_mB"], 5, f"=(PL!G{SUMR[('B', 'nim')]}+PL!G{SUMR[('B', 'nim')]}*0.8)/2")
IN.cell(INROW["fade_mC"], 5, f"=PL!G{SUMR[('C', 'nim')]}*0.8")
for k in ("fade_mA", "fade_mB", "fade_mC"): IN.cell(INROW[k], 5).font = Font(color="000000")
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型 − 錨點").font = BOLD; sr += 1
CHK = []
for lab, f in (("  三情境：營收差（NT$B）", "+".join(f"ABS(D{PB[s]['rev']}-a26_rev)" for s in SC)),
               ("  三情境：EPS 差（NT$）", "+".join(f"ABS(D{PB[s]['eps']}-a26_eps)" for s in SC)),
               ("  情境 A：其他 AI 載板 FY2026 應 > 0（1＝是）", f"IF(D{PB['A']['oai']}>0,0,1)")):
    PL.cell(sr, 1, lab); PL[f"D{sr}"] = "=" + f; PL[f"D{sr}"].number_format = "0.0000"; CHK.append(sr); sr += 1
assert sr < 60, sr
PL.column_dimensions["A"].width = 66; PL.column_dimensions["B"].width = 10; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 13
PL.cell(2, 1, "灰底＝FY2030–FY2032 收斂法（上游顆數只到 2029）：營收年增線性收斂到 fade_g、淨利率線性收斂到 fade_m（In 頁 I 節），非驅動因子推導")
# 自由現金流（算力收支 2026-10-25 請求；地圖依列標籤讀取）：表頭列＋一列（三情境 × FY2026–FY2032）
FCF_YRS = (2026, 2027, 2028, 2029, 2030, 2031, 2032)
fh = max(PB[s]["end"] for s in SC) + 2; FCF_ROW = fh + 1
PL.cell(fh, 1, "項目（自由現金流；欄＝年度 × 情境，FY2026 為錨點年、FY2030 起為收斂法）").font = BOLD
PL.cell(fh, 2, "單位").font = BOLD
_col = 3
for s_ in SC:
    for y_ in FCF_YRS:
        PL.cell(fh, _col, f"FY{y_} {s_}").font = BOLD
        c_ = PL.cell(FCF_ROW, _col, f"={FC[y_]}{PB[s_]['fcf']}"); c_.number_format = "#,##0.0"
        if y_ >= 2030: c_.fill = FADEFILL
        _col += 1
PL.cell(fh, _col, "說明").font = BOLD
PL.cell(FCF_ROW, 1, "自由現金流（淨利＋折舊 − 資本支出；不含營運資金變動）"); PL.cell(FCF_ROW, 2, "NT$B")
PL.cell(FCF_ROW, _col, "取自本頁各情境推導區塊的「股東現金流」列（歸屬母公司淨利＋合併折舊 − 合併資本支出；FY2030 起資本支出率收斂到穩態更新）；不進 EPS")

# ---------------- Units：份額合計檢查（與 Ibiden、欣興、南電同口徑） ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 景碩在 AI 加速器與 NVIDIA CPU 載板的份額與每顆金額（口徑與 Ibiden、欣興、南電模型一致；情境 A）。份額合計＝Ibiden（ibiden/dist）＋欣興（unimicron/dist）＋南電（nanya_pcb/dist）＋景碩，其餘為三星電機、AT&S 等"
A_ = PB["A"]; cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ibrel = lambda y: "1" if y == 2026 else f"ibrel_{y % 100}"
UKEYS = ["unv", "uot", "ucpu", "v", "shnv", "ibnv", "unnv", "npnv", "sumnv", "shot", "ibot", "unot", "npot", "sumot", "shcpu", "pu", "accu", "aisub", "itr", "pgw"]
UR = {k: 4 + i for i, k in enumerate(UKEYS)}
UDEF = {
    "unv": ("NVIDIA 加速器顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['unv']}", "#,##0.00", "markets/network"),
    "uot": ("非 NVIDIA 加速器顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['uot']}", "#,##0.00", ""),
    "ucpu": ("NVIDIA CPU（Grace／Vera）顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['ucpu']}", "#,##0.00", "NVIDIA 加速器 × cpu_r（Analogy）"),
    "v": ("每顆 NVIDIA 加速器載板金額（全市場）", "US$", lambda y: f"=PL!{FC[y]}{A_['v']}", "#,##0", "Ibiden 模型 Units 頁同值（FY2026 校準 × 面積層數 × 溢價）"),
    "shnv": ("景碩份額：NVIDIA（GPU）", "%", lambda y: f"=PL!{FC[y]}{A_['shnv']}", "0.0%", "Assumed（券商稱 Rubin 15–20%，與上游合計衝突）"),
    "ibnv": ("Ibiden 份額：NVIDIA（70% × Ibiden 份額相對值）", "%", lambda y: f"=ib_sh26*{ibrel(y)}", "0.0%", "Ibiden 模型"),
    "unnv": ("欣興份額：NVIDIA", "%", lambda y: f"=un_nv_{y % 100}", "0.0%", "欣興模型"),
    "npnv": ("南電份額：NVIDIA", "%", lambda y: f"=np_nv_{y % 100}", "0.0%", "南電模型"),
    "sumnv": ("合計：Ibiden＋欣興＋南電＋景碩（NVIDIA）", "%", lambda y: "=" + "+".join(f"{cyc(y)}{UR[k]}" for k in ("shnv", "ibnv", "unnv", "npnv")), "0.0%", "應 ≤ 100%；其餘為三星電機等"),
    "shot": ("景碩份額：非 NVIDIA（金額）", "%", lambda y: f"=PL!{FC[y]}{A_['shot']}", "0.0%", "Assumed"),
    "ibot": ("Ibiden 份額：非 NVIDIA（Ibiden 權重 × 份額相對值 ÷ v_ot）", "%", lambda y: f"=ib_w_ot*{ibrel(y)}/v_ot", "0.0%", "與欣興、南電模型同算法"),
    "unot": ("欣興份額：非 NVIDIA（金額）", "%", lambda y: f"=un_ot_{y % 100}", "0.0%", "欣興模型"),
    "npot": ("南電份額：非 NVIDIA（金額）", "%", lambda y: f"=np_ot_{y % 100}", "0.0%", "南電模型"),
    "sumot": ("合計：Ibiden＋欣興＋南電＋景碩（非 NVIDIA）", "%", lambda y: "=" + "+".join(f"{cyc(y)}{UR[k]}" for k in ("shot", "ibot", "unot", "npot")), "0.0%", "應 ≤ 100%；其餘為三星電機、AT&S"),
    "shcpu": ("景碩份額：NVIDIA CPU（Grace／Vera）", "%", lambda y: f"=PL!{FC[y]}{A_['shcpu']}", "0.0%", "公司與券商說法（Interested-party）；其餘主要為 Ibiden"),
    "pu": ("每顆 NVIDIA 加速器的景碩營收（GPU＋CPU 載板）", "US$", lambda y: f"=PL!{FC[y]}{A_['v']}*(PL!{FC[y]}{A_['shnv']}+cpu_r*v_cpu*PL!{FC[y]}{A_['shcpu']})*PL!{FC[y]}{A_['cy']}", "#,##0.0", "Ibiden、欣興、南電對照值見各自 dist Units 頁"),
    "accu": ("景碩加速器與 CPU 載板營收", "US$M", lambda y: f"=PL!{FC[y]}{A_['accu']}", "#,##0", ""),
    "aisub": ("景碩 AI 載板營收（含網通 ASIC 等其他 AI）", "NT$B", lambda y: f"=PL!{FC[y]}{A_['aisub']}", "#,##0.0", ""),
    "itr": ("景碩 AI 載板營收 ÷ AI IT 資本支出", "%", lambda y: f"=PL!{FC[y]}{A_['aisub']}/fx/itc_{y % 100}", "0.000%", "× 每 GW IT capex 即每 GW 營收"),
    "pgw": ("每 GW 景碩 AI 載板營收（VR200 每 GW IT capex 37.59B）", "US$M", lambda y: f"=PL!{FC[y]}{A_['aisub']}/fx/itc_{y % 100}*37.59*1000", "#,##0", "Tokenomics IF_CapexIT（VR200 基準，經橋接表；與 Ibiden、欣興、南電模型相同係數）")}
for key in UKEYS:
    lab, unit, f, fmt, note = UDEF[key]; ur = UR[key]
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note)
UN_NOTE = 4 + len(UKEYS) + 1
UN.cell(UN_NOTE, 1, "讀法：景碩的 AI 曝險以 NVIDIA CPU（Grace→Vera）載板為主，GPU 與 ASIC 加速器載板份額小；券商稱 Rubin GPU 份額 15–20%，但 Ibiden、欣興、南電模型的份額已合計 94–97%，本模型取剩餘空間，兩者不能同時成立（見 README）。").font = BOLD
UN.column_dimensions["A"].width = 60; UN.column_dimensions["B"].width = 10
for c in "CDEF": UN.column_dimensions[c].width = 13
UN.column_dimensions["G"].width = 66

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，年底落後 P/E）只作參考"
V["A2"] = "循環位置：景碩不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照 Ibiden、欣興、南電；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
O = 53
V.cell(O, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, O + 1, ["項目", "A", "B", "C", "說明"])
V.cell(O + 2, 1, "現價（2026-10-22 收盤）"); V[f"B{O + 2}"] = "=px"; V[f"B{O + 2}"].number_format = "#,##0"
for k, y, lab in ((3, 2026, "現價 ÷ FY2026 EPS（錨點）"), (4, 2027, "現價 ÷ FY2027 EPS（模型）"), (5, 2028, "現價 ÷ FY2028 EPS（模型）"), (6, 2029, "現價 ÷ FY2029 EPS（模型）")):
    V.cell(O + k, 1, lab)
    for j, s in enumerate(SC):
        c = "BCD"[j]; V[f"{c}{O + k}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{O + k}"].number_format = "0.0"
V.cell(O + 7, 1, "未來 12 月 EPS（模型 B）≈ 0.2 × FY2026＋0.8 × FY2027"); V[f"B{O + 7}"] = f"=0.2*PL!D{SUMR[('B', 'eps')]}+0.8*PL!E{SUMR[('B', 'eps')]}"; V[f"B{O + 7}"].number_format = "0.00"
V.cell(O + 8, 1, "  現價 ÷ 未來 12 月 EPS（模型 B）"); V[f"B{O + 8}"] = f"=px/B{O + 7}"; V[f"B{O + 8}"].number_format = "0.0"
V.cell(O + 9, 1, "近 12 月 P/E（stockanalysis，10-22）"); V[f"B{O + 9}"] = "=pe_ttm"; V[f"B{O + 9}"].number_format = "0.0"
V.cell(O + 10, 1, "前瞻 P/E（stockanalysis，價格日期 10-22）"); V[f"B{O + 10}"] = "=fpe"; V[f"B{O + 10}"].number_format = "0.0"
V[f"E{O + 3}"] = "FY2026 三情境相同（錨點）"; V[f"E{O + 4}"] = "FY2027 情境差只來自 BT 與其他載板年增（顆數係數從 CY2028 開始）"
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
for x in (0.75, 1.0, 1.1, 1.21, 1.5):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    for j, s in enumerate(SC):
        V.cell(vr, 2 + j, f"=px*A{vr}/PL!F{SUMR[(s, 'eps')]}").number_format = "0.0"
    V.cell(vr, 5, "=pe_mid").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數愈高，現價愈依賴「Vera CPU 載板份額、ABF 漲價延續、K6 新產能如期開出」。本版未做先驗，無 Andy 合理 P/E 對照。")
WMT_NOTE = vr
# 十年 P/E：stockanalysis 年度 P/E（FY2021–25）；自算＝digrin 12 月未調整股價 ÷ 年度 EPS（wantgoo）；2019 EPS 為負，不計
PE_SA = [None, None, None, None, None, 27.22, 6.78, 953.58, 943.41, 45.51]
PX_DEC = [71.20, 54.20, 43.60, 51.80, 80.90, 233.0, 104.50, 99.70, 101.0, 159.0]
EPS_H = [5.01, 1.10, 0.78, -4.52, 1.21, 8.56, 15.47, 0.11, 0.11, 3.51]
PE_SELF = [None if e_ <= 0 else f"={p_}/{e_}" for p_, e_ in zip(PX_DEC, EPS_H)]
T10 = vr + 2
V.cell(T10, 1, "5. 十年 P/E（2016–2025 年底；兩來源平均，分位排名用；2019 年 EPS 為負，P/E 無意義，不計）").font = BOLD
hdr(V, T10 + 1, ["年底", "兩來源平均 P/E", "stockanalysis", "自算"])
RB0 = T10 + 16
for i, y in enumerate(range(2016, 2026)):
    rr = T10 + 2 + i
    V.cell(rr, 1, str(y))
    if PE_SELF[i] is not None:
        V.cell(rr, 2, f"=AVERAGE(C{rr}:D{rr})").number_format = "0.00"
        V.cell(rr, 3, f'=IF(ISNUMBER(B{RB0 + 5 + i}),B{RB0 + 5 + i},"")').number_format = "0.00"
        V.cell(rr, 4, f'=IF(ISNUMBER(C{RB0 + 5 + i}),C{RB0 + 5 + i},"")').number_format = "0.00"
    else:
        V.cell(rr, 5, "不計（EPS 為負，P/E 無意義）")
PE10 = f"$B${T10 + 2}:$B${T10 + 11}"
V.cell(T10 + 12, 1, "年數"); V[f"B{T10 + 12}"] = f"=COUNT({PE10})"
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "stockanalysis（年度 P/E）", "vals": PE_SA, "tag": "Interested-party（資料商；只有 2021–2025 共 5 年）"},
    {"label": "自算（digrin 12 月未調整股價 ÷ 年度 EPS）", "vals": PE_SELF, "tag": "Derived（價格為月資料的 12 月值；EPS 取自 wantgoo 歷年股利表；2019 EPS 為負不計，共 9 年）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位，兩個來源分別取分位再平均。景碩的落後 P/E 與盈餘反向：盈餘低谷年（2023、2024 年 EPS 0.11）P/E 約 900 倍，高峰年（2022）約 7 倍；2019 年虧損不計，十年多為 AI 載板需求前的循環，套用在 FY2027 EPS 時偏誤方向不明確。")
assert V.cell(RB0 + 5, 1).value == "2016" and V.cell(RB0 + 14, 1).value == "2025", (V.cell(RB0 + 5, 1).value, V.cell(RB0 + 14, 1).value)
rule_rows = [rw for rw in range(RB0, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")

# 0. 目標價（Andy 2026-10-25 11:28／11:53／12:22：DCF 與 Comparable 加權；DCF 負值取 0；列標籤依工作單逐字；B／C／D＝A 順利／B 平平／C 不如預期）
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
V["E14"] = "方法：景碩資本密集且模型有資本支出與折舊，改用每股（淨利＋折舊－資本支出）（PL 頁「每股股東現金流」；FY2030 起資本支出率收斂到穩態更新）"
tset(15, "reinv", "  參考：再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）；本模型不採用，只作對照（第 25 列）")
CFR = {}
for i, y in enumerate(YRS6):
    tset(16 + i, f"cf{y}", f"  FY{y}", [f"=PL!{FC[y]}{PB[s]['fps']}" for s in SC], "0.00"); CFR[y] = 16 + i
tset(22, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=MAX(0,SUMPRODUCT({c}16:{c}21,1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}21*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6)" for c in "BCD"], "0.00",
     "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底；< 0 時取 0（股東有限責任）")
tset(23, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}22*(1+tp_r)" for c in "BCD"], "0.00")
tset(24, "dcftp", "DCF 目標價（情境加權）", ["=tp_wA*B23+tp_wB*C23+tp_wC*D23", None, None], "0.00", "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁")
tset(25, "dcfref", "  參考：EPS ×（1 − g／ROE）法的 DCF 目標價（情境加權）",
     ["=(tp_wA*MAX(0,SUMPRODUCT(B8:B13,1/(1+tp_r)^{1;2;3;4;5;6})+B13*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6)+tp_wB*MAX(0,SUMPRODUCT(C8:C13,1/(1+tp_r)^{1;2;3;4;5;6})+C13*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6)"
      "+tp_wC*MAX(0,SUMPRODUCT(D8:D13,1/(1+tp_r)^{1;2;3;4;5;6})+D13*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6))*(1-B15)*(1+tp_r)", None, None], "0.00",
     "近三年 ROE 約 4%（遠低於 g 的 4% 所需）→ 再投資率上限 80%；不進目標價")
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = [("欣興（Unimicron）", "3037.TW", 42.28, "2026-10-22"), ("南電（Nan Ya PCB）", "8046.TW", 33.73, "2026-09-07"), ("Ibiden（揖斐電）", "4062.T", 53.97, "2026-10-16"),
         ("三星電機（Samsung Electro-Mechanics）", "009150.KS", 43.40, "2026-08-28"), ("AT&S", "ATS.VI", 14.91, "2026-10-23")]
for i, (nm, tk, pe, dt) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左；與欣興、南電、Ibiden 模型同一組數字）"); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", "ABF 載板同業：欣興、南電、Ibiden、三星電機、AT&S（無中國 A 股）；南電、三星電機價格日期較舊")
tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）", [f"=tp_wA*(2/12*B{EPSR[2027]}+10/12*B{EPSR[2028]})+tp_wB*(2/12*C{EPSR[2027]}+10/12*C{EPSR[2028]})+tp_wC*(2/12*D{EPSR[2027]}+10/12*D{EPSR[2028]})", None, None], "0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], "0.00", "同業中位數 × 前瞻 EPS")
V.cell(37, 1, "0.3 結論").font = BOLD
hdr(V, 38, ["項目", "數值", "", "", "說明"])
tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B24+(1-tp_w_dcf)*B35", None, None], "0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁（0.5）")
tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], "0.00")
tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], "0.00", "stockanalysis：13 位分析師平均 1,092.15（頁面價格日期 2026-10-22）；10 月永豐 1,280、玉山 1,270")
tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-22 收盤）")
tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝頁面價格 ÷ 前瞻 P/E（In 頁 cons_feps）；兩者期間差約一年")
tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
tset(46, "band", "價格帶（DCF 不如預期～順利）", ['=TEXT(D23,"#,##0")&"～"&TEXT(B23,"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=IF(px/{c}{EPSR[2028]}>=MAX({PE10}),1,IF(px/{c}{EPSR[2028]}<=MIN({PE10}),0,PERCENTRANK({PE10},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 5 節兩來源平均，2019 虧損不計）的分位（超出區間時記 0% 或 100%）；只作檢查，不進目標價")
tset(48, "grade", "上檔空間級距", ['=IF(B42>0.5,"大於 50%",IF(B42>0.25,"25%–50%",IF(B42>=0,"0%–25%",IF(B42>=-0.25,"0%～−25%",IF(B42>=-0.5,"−25%～−50%","小於 −50%")))))', None, None], "@",
     "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）")
tset(49, "rec", "投資建議", ['=IF(B42>0.5,"強力買進",IF(B42>0.25,"買進",IF(B42>=-0.25,"中立",IF(B42>=-0.5,"放空","強力放空"))))', None, None], "@",
     "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空（剛好 +25% 為中立、+50% 為買進、−25% 為中立、−50% 為放空）")
V.cell(50, 1, "目標價方法")
V["B50"] = ("DCF：股東現金流＝每股（淨利＋折舊－資本支出），r＝10%、g＝4%，FY2027–32 明確期＋永續，情境價值 < 0 取 0；景碩 2026–28 年新廠（K6）與設備資本支出合計約 430 億，"
            "折舊與資本支出含晶碩（淨利只計歸屬母公司），現金流偏低估約晶碩非控制部分；客戶贊助設備若屬實，資本支出偏高估。ABF 溢價沿用 Ibiden 模型（2027 持平、之後回落），比券商保守。"
            "Comparable：欣興、南電、Ibiden、三星電機、AT&S 前瞻 P/E 中位數（無 A 股）；同業同處 ABF 短缺溢價期，倍數偏高；景碩含晶碩（隱形眼鏡）未另估。"
            "情境對應：A＝順利（顆數基準）、B＝平平（顆數 ×0.8／0.85）、C＝不如預期（顆數 ×0.6），皆讀 tsmc 同字母係數；BT 與其他載板年增 A 最高、C 最低")
for r_ in (24, 35, 39):
    V[f"B{r_}"].fill = KEY
V.column_dimensions["A"].width = 60
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；FY2026 恆等於錨點，其他 AI 載板在每個案例內重新校準）：對 FY2027、FY2028 營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("AI 部署 ×0.8", "FY2027 起顆數與其他 AI ×0.8", dict(mf="0.8")), ("AI 部署 ×1.2", "×1.2", dict(mf="1.2")),
         ("ABF 產能 低", "產能指數 1.10／1.30／1.55", dict(cap="cap_{y}_lo")), ("ABF 產能 高", "1.25／1.62／2.00", dict(cap="cap_{y}_hi")),
         ("ABF 溢價 低（Ibiden 低值）", "0.95／0.85／0.75", dict(pp="ibpp_{y}_lo")), ("ABF 溢價 高（Ibiden 高值）", "1.05／1.05／1.00", dict(pp="ibpp_{y}_hi")),
         ("漲價年度化 8%", "基準 15%", dict(carry="carry_lo")), ("漲價年度化 22%", "", dict(carry="carry_hi")),
         ("ABF 現金毛利率 40%", "基準 50%", dict(ftabf="ft_abf_lo")), ("ABF 現金毛利率 60%", "", dict(ftabf="ft_abf_hi")),
         ("BT 現金毛利率 30%", "基準 45%", dict(ftbt="ft_bt_lo")), ("BT 現金毛利率 60%", "", dict(ftbt="ft_bt_hi")),
         ("BT 年增 走 C 路徑", "12%／0%／−3%", dict(gbt="g_bt{y}_C")),
         ("非 AI ABF 年增 2%", "基準 6%", dict(gnai="g_nai_lo")), ("非 AI ABF 年增 10%", "", dict(gnai="g_nai_hi")),
         ("折舊年限 6 年", "基準 8 年", dict(life="life_lo")), ("折舊年限 10 年", "", dict(life="life_hi")),
         ("資本支出 低", "100／90／70 億", dict(cx="cx{y}_lo")), ("資本支出 高", "220／220／160 億", dict(cx="cx{y}_hi")),
         ("景碩 NVIDIA GPU 份額 低", "1%→2%", dict(shnv="sh_nv_{y}_lo")), ("景碩 NVIDIA GPU 份額 高（券商說法）", "10%→15%", dict(shnv="sh_nv_{y}_hi")),
         ("景碩 Vera CPU 份額 低", "50%→45%", dict(shcpu="sh_cpu_{y}_lo")), ("景碩 Vera CPU 份額 高", "90%→100%", dict(shcpu="sh_cpu_{y}_hi")),
         ("景碩 ASIC 份額 低", "2%→3%", dict(shot="sh_ot_{y}_lo")), ("景碩 ASIC 份額 高", "8%→10%", dict(shot="sh_ot_{y}_hi")),
         ("CPU 載板金額比 0.20", "基準 0.30", dict(vcpu="v_cpu_lo")), ("CPU 載板金額比 0.45", "", dict(vcpu="v_cpu_hi")),
         ("營業費用彈性 1.0", "基準 0.6", dict(phi="phi_hi")), ("稅率 20%", "基準 12%", dict(tax="tax_hi")), ("業外 0", "每年 0（基準 4 億）", dict(nonop="nonop{y}_lo")),
         ("晶碩年增 3%", "基準 10%", dict(glens="g_lens_lo"))]
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.0"), ("D", "gm", 2027, "0.00%"), ("E", "eps", 2027, "0.00"),
                           ("G", "rev", 2028, "#,##0.0"), ("H", "gm", 2028, "0.00%"), ("I", "eps", 2028, "0.00")):
        SE[f"{c}{i}"] = f"={pre}{FC[y]}{R[key]}"; SE[f"{c}{i}"].number_format = fmt
    SE[f"F{i}"] = f"=ROUND(E{i}/E$4-1,8)"; SE[f"J{i}"] = f"=ROUND(I{i}/I$4-1,8)"   # 取到 1e-8，避免產能受限時的浮點殘差（約 1e-16）
    for c in "FJ": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
base_row = 60; SB = {}
for nm, setting, ov in cases:
    SB[nm] = engine(SE, base_row, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    base_row = SB[nm]["end"] + 2
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm], "")
nn = 5 + len(cases)
put_case(nn, "情境 B", "tsmc 顆數係數 B；BT 年增較低", PB["B"], "PL!")
put_case(nn + 1, "情境 C", "tsmc 顆數係數 C；BT 年增最低", PB["C"], "PL!")
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：漲價年度化與 ABF 溢價直接決定毛利（溢價全數落在毛利）；ABF 產能受限時，份額與部署的上修只部分轉成營收；FY2027 情境差只來自 BT 年增。")
SENS_NOTE = nn + 3
TPS0 = nn + 5
SE.cell(TPS0, 1, "目標價參數（只改 DCF 參數；PL 收斂法的 fade_g 與資本支出率路徑不隨 tp_g 變動）").font = BOLD
hdr(SE, TPS0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
def dcf_f(c, r_, g_):
    return f"MAX(0,SUMPRODUCT(Val!{c}16:{c}21,1/(1+{r_})^{{1;2;3;4;5;6}})+Val!{c}21*(1+{g_})/({r_}-{g_})/(1+{r_})^6)"
def tp_case(i, name, setting, r_="tp_r", g_="tp_g", w="tp_w_dcf"):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    SE[f"C{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_f(c, r_, g_)}" for s, c in zip(SC, "BCD")) + f")*(1+{r_})"
    SE[f"D{i}"] = f"=Val!B{T['comptp']}"
    SE[f"E{i}"] = f"={w}*C{i}+(1-{w})*D{i}"
    SE[f"F{i}"] = f"=E{i}/E${TPS0 + 2}-1"
    for c in "CDE": SE[f"{c}{i}"].number_format = "#,##0.0"
    SE[f"F{i}"].number_format = "0.0%"
TPCASES = [("基準", "r 10%、g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "r 11%", dict(r_="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", dict(r_="(tp_r-0.01)")),
           ("長期成長 +1pt", "g 5%", dict(g_="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", dict(g_="(tp_g-0.01)")),
           ("DCF 權重 0.3", "", dict(w="0.3")), ("DCF 權重 0.7", "", dict(w="0.7"))]
for k, (nm, st, kw) in enumerate(TPCASES):
    tp_case(TPS0 + 2 + k, nm, st, **kw)
TPS_LAST = TPS0 + 1 + len(TPCASES)
assert TPS_LAST < 60, TPS_LAST
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 40
for c in "CDEFGHIJ": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 券商、共識與市場口徑。只作事後對照，不為對上數字而調參數。FactSet 共識為 2026-06（2Q26 財報前）；10 月券商數字取自媒體轉載（永豐、玉山、美系）"
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
RC.cell(rr_, 1, "A. EPS 與營收 vs 共識、券商").font = BOLD; rr_ += 1
rec("FY2026 EPS vs 共識 10.42（主串補位排序表）", "=cons_eps26", pl("eps", 2026), "主串提供", "永豐 13.21、玉山 12.93、美系 10.6；FactSet 2026-06 平均 9.73")
rec("FY2026 EPS vs 永豐 13.21（2026-10-14）", 13.21, pl("eps", 2026), "旺得富（永豐）", "")
rec("FY2027 EPS vs 共識 26.47（主串補位排序表）", "=cons_eps27", pl("eps", 2027), "主串提供", "")
rec("FY2027 EPS vs 券商平均（永豐 36.03、玉山 42.41、美系 27.3）", "=bk_eps27", pl("eps", 2027), "旺得富、經濟日報", "差距超過 30% 的原因見 README")
rec("FY2027 EPS vs FactSet 平均 19.93（2026-06）", "=fs_eps27", pl("eps", 2027), "鉅亨（FactSet 13 位）", "")
rec("FY2028 EPS vs FactSet 平均 34.39（2026-06）", "=fs_eps28", pl("eps", 2028), "鉅亨（FactSet）", "")
rec("FY2028 EPS vs 券商平均（永豐 62.22、玉山 89.75、美系 51.5）", "=bk_eps28", pl("eps", 2028), "旺得富、經濟日報", "")
rec("FY2027 營收 vs FactSet 平均（2026-06）", "=fs_rev27", pl("rev", 2027), "鉅亨（FactSet）", "", fmt="#,##0.0")
rec("FY2028 營收 vs FactSet 平均（2026-06）", "=fs_rev28", pl("rev", 2028), "鉅亨（FactSet）", "", fmt="#,##0.0")
rec("模型前瞻 12 月 EPS（0.2 × FY2026＋0.8 × FY2027）vs 頁面價格 ÷ 前瞻 P/E", "=cons_feps", lambda s: f"=0.2*PL!D{PB[s]['eps']}+0.8*PL!E{PB[s]['eps']}", "stockanalysis（推導）", "口徑不同，只作量級比較")
RC.cell(rr_, 1, "B. 錨點與毛利率").font = BOLD; rr_ += 1
rec("3Q26 毛利率（假設）vs 美系券商 27.8%（百分點）", 0.278, lambda s: "=Q!E6", "經濟日報 2026-10-12", "3Q26 財報約 10 月底～11 月初", fmt="0.0%", pt=True)
rec("3Q26 EPS（模型）vs 永豐估 4.29", 4.29, lambda s: "=Q!E16", "旺得富（永豐）", "", fmt="0.00")
rec("4Q26 營收（模型）vs 永豐估 173.14 億", 17.314, lambda s: "=Q!F4", "旺得富（永豐）", "", fmt="#,##0.00")
rec("4Q26 毛利率（假設）vs 永豐估 33%（百分點）", 0.33, lambda s: "=Q!F6", "旺得富（永豐）", "美系 29.3%", fmt="0.0%", pt=True)
rec("FY2027 毛利率 vs 高盛「每年約 +5pt」推算（FY2026 錨點＋5pt）", "=Q!G6+0.05", pl("gm", 2027), "TechNews（高盛）", "", fmt="0.0%", pt=True)
rec("FY2028 毛利率 vs 同上（FY2026 錨點＋10pt）", "=Q!G6+0.10", pl("gm", 2028), "TechNews（高盛）", "", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "C. 產能、份額與 AI 口徑").font = BOLD; rr_ += 1
rec("FY2027 ABF 需求滿足率（< 100%＝產能受限）", None, pl("fill", 2027), "PL 頁", "高盛：2027–28 年 ABF 可能短缺", fmt="0.0%")
rec("FY2028 ABF 需求滿足率", None, pl("fill", 2028), "PL 頁", "", fmt="0.0%")
rec("FY2026 AI 載板占營收 vs 高盛 10%（百分點）", 0.10, pl("aish", 2026), "TechNews（高盛）", "", fmt="0.0%", pt=True)
rec("FY2027 AI 載板占營收 vs 高盛 15–20% 中點（百分點）", 0.175, pl("aish", 2027), "TechNews（高盛）", "", fmt="0.0%", pt=True)
rec("FY2027 份額合計：Ibiden＋欣興＋南電＋景碩（NVIDIA）", None, lambda s: f"=Units!D{UR['sumnv']}", "Units 頁", "應 ≤ 100%", fmt="0.0%")
rec("FY2027 份額合計：Ibiden＋欣興＋南電＋景碩（非 NVIDIA）", None, lambda s: f"=Units!D{UR['sumot']}", "Units 頁", "應 ≤ 100%", fmt="0.0%")
rec("FY2027 每顆 NVIDIA 加速器的景碩營收（US$；GPU＋CPU 載板）", None, lambda s: f"=Units!D{UR['pu']}", "Units 頁", "Ibiden 同期約 169（ibiden/dist Units 頁）", fmt="#,##0.0")
rec("FY2027 每 GW 景碩 AI 載板營收（US$M）", None, lambda s: f"=Units!D{UR['pgw']}", "Units 頁", "", fmt="#,##0")
RC.cell(rr_, 1, "D. 資本支出").font = BOLD; rr_ += 1
rec("FY2026–FY2028 資本支出合計 vs 公司 ABF 擴產計畫 431 億（235＋196）", 43.1, lambda s: f"=PL!D{PB[s]['cx']}+PL!E{PB[s]['cx']}+PL!F{PB[s]['cx']}", "工商時報 2026-08-06", "模型含晶碩；計畫期間與年度未完全對齊", fmt="#,##0.0")
RC.column_dimensions["A"].width = 72
for c in "BCDEFGH": RC.column_dimensions[c].width = 11
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60
RECON_LAST = rr_ - 1

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司財報經資料商與媒體轉載；法說內容經媒體摘要；券商與資料商皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("資料商：stockanalysis（現價 986（2026-10-22）、股數、前瞻與近 12 月 P/E、目標價 1,092.15（13 位）、下次財報日估 2026-11-06）", "https://stockanalysis.com/quote/tpe/3189/"),
        ("資料商：stockanalysis 年度與季損益表（營收、毛利、營業利益、稅前、所得稅、非控制權益、EPS、股數）", "https://stockanalysis.com/quote/tpe/3189/financials/income-statement/?p=quarterly"),
        ("資料商：stockanalysis 季現金流量表（折舊攤提、資本支出、現金增資 101.59 億）", "https://stockanalysis.com/quote/tpe/3189/financials/cash-flow-statement/?p=quarterly"),
        ("資料商：stockanalysis 比率頁（年度 P/E、ROE FY2021–25）", "https://stockanalysis.com/quote/tpe/3189/financials/ratios/"),
        ("資料商：stockanalysis 晶碩（6491）季損益表（營收、淨利）", "https://stockanalysis.com/quote/tpe/6491/financials/?p=quarterly"),
        ("資料商：nstock（股本 52.69 億、2Q26 毛利率、9 月營收、現價核對）", "https://www.nstock.tw/%E6%99%AF%E7%A2%A9%E5%81%9A%E4%BB%80%E9%BA%BC.html"),
        ("媒體：TechNews 2026-10-06（9 月營收 51.81 億、前 9 月 381.41 億）", "https://finance.technews.tw/2026/10/06/kinsus-3189-202609-financial-report/"),
        ("媒體：工商時報 2026-08-06（7 月營收、資本支出 196 億＋235 億）", "https://www.ctee.com.tw/news/20260806701841-430503"),
        ("券商（媒體）：TechNews 2026-08-04 高盛（目標價 1,125；Vera CPU 100% 目標、AI 占營收、毛利率、客戶贊助設備、產能倍增）", "https://finance.technews.tw/2026/08/04/vera-cpu/"),
        ("券商（媒體）：經濟日報 2026-08-07（2Q26 EPS、毛利率、ABF 產能 4,000→8,000 萬顆）", "https://money.udn.com/money/story/5607/9676971"),
        ("券商（媒體）：經濟日報 2026-09-16（Vera CPU 80–100%、Rubin 15–20%、BT 記憶體逾五成）", "https://money.udn.com/money/amp/story/5607/9757810"),
        ("券商（媒體）：經濟日報 2026-10-12（美系券商 EPS 2026–28、3Q／4Q26 毛利率）", "https://money.udn.com/money/story/5607/9808557"),
        ("券商（媒體）：旺得富 2026-10-14（永豐 1,280、玉山 1,270；EPS 2026–28；8 月自結；4Q26 預估）", "https://wantrich.chinatimes.com/news/20261014900438-420101"),
        ("券商（媒體）：經濟日報 2026-04-23（CPU 份額 10%→50%、稼動率）", "https://money.udn.com/money/amp/story/5607/9459664"),
        ("媒體：MoneyDJ 2026-01-19（產品組合 ABF 30–40%、BT 35–40%、隱形眼鏡 15–20%；2026 年資本支出約 80 億）", "https://tw.stock.yahoo.com/news/ai-%E8%A8%98%E6%86%B6%E9%AB%94%E9%9C%80%E6%B1%82%E6%8E%A8%E5%8B%95-%E6%99%AF%E7%A2%A9q1%E6%B7%A1%E5%AD%A3%E6%9C%89%E6%92%90-042600728.html"),
        ("公司（摘要）：富果 2Q25 法說摘要（應用別、ABF 約 40%、晶碩 18%）", "https://blog.fugle.tw/post/earnings-call-3189-2025-08-19"),
        ("研究文章：brodoginvestment 景碩分析 2026-04-07（引述券商；現增、份額、資本支出）", "https://brodoginvestment.substack.com/p/3db"),
        ("共識：鉅亨 2026-06-20 FactSet（13 位分析師 EPS、營收）", "https://news.cnyes.com/news/id/6505998"),
        ("媒體：優分析（晶碩 2025 年營收 70.39 億）", "https://uanalyze.com.tw/articles/8838744153"),
        ("資料商：digrin 月股價（12 月未調整值）", "https://www.digrin.com/stocks/detail/3189.TW/price"),
        ("資料商：wantgoo 歷年股利表（EPS 2014–2025）", "https://www.wantgoo.com/stock/3189/profitability/dividend-payout-ratio"),
        ("資料商：stockanalysis 同業前瞻 P/E（欣興、南電、Ibiden、三星電機、AT&S；與欣興、南電、Ibiden 模型同一組）", "https://stockanalysis.com/quote/tpe/3037/"),
        ("主串提供：共識 EPS FY2026 10.42、FY2027 26.47（補位排序表，Derived；原始口徑本串未查證）", "—"),
        ("本 repo 上游成品（唯讀）：ibiden、unimicron、nanya_pcb、markets/network、tsmc；情境係數與 markets/memory/market.py SCEN_UNITS 核對", f"{rel(IB_F)}；{rel(UN_F)}；{rel(NP_F)}；{rel(NET_F)}；{rel(TSMC_F)}"),
        ("待補（人工）：3Q26 財報與法說（約 10 月底～11 月初）；公司原文（公開資訊觀測站財報、法說簡報）；FY2016–20 營收（十年平均淨利率）；ABF／BT 季度拆分；加速器、CPU 與其他 AI 載板拆分；份額量化；晶碩持股比例原文與毛利率；晶碩 3Q26 營收；折舊會計年限；客戶贊助設備的金額與會計處理；10 月後 FactSet 共識", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 110; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
OBS = ["目標價 NT$908（DCF 529、Comparable 1,287 各半），比現價 986（2026-10-22）低 7.9%，級距 0%～−25%（中立）；Comparable 單獨為 +30.5%，市場共識目標價 1,092。兩法差距來自 DCF：2026–28 年資本支出合計 440 億（K6 新廠與 ABF 擴產），情境 B 每股股東現金流 FY2027 只有 10.37（EPS 26.21 的 40%）；"
       "本模型 12 個月後前瞻 EPS 30.44 比共識前瞻 EPS 22.06 高 38%（期間差約一年），共識目標價隱含 49.5 倍、比同業中位數 42.3 倍高 17%。",
       "FY2027–28 ABF 需求滿足率約 86%／88%（產能受限），EPS 由價格與產能決定、而非 AI 份額：漲價年度化 8%／22%（基準 15%）使 FY2027 EPS −11.4%／+11.4%，ABF 溢價取 Ibiden 低／高值使 FY2028 EPS −15.2%／+15.7%，ABF 產能低值 −13.2%；"
       "AI 部署 ×0.8、Vera CPU 或 GPU 份額高低對 FY2027–28 EPS 的影響都在 0.3% 以內。",
       "份額說法與上游衝突：券商稱景碩 Rubin GPU 份額 15–20%、Vera CPU 80–100%，但 Ibiden、欣興、南電模型的 NVIDIA 份額已合計約 94–97%，本模型 GPU 取 2%→6%。下一個檢驗點是 3Q26 財報與法說（stockanalysis 估 2026-11-06）：本模型 3Q26 毛利率 31%、EPS 4.30（永豐 4.29、美系毛利率 27.8%），"
       "FY2026 EPS 13.69 比主串共識 10.42 高 31%。"]   # 重點觀察（數字取自重算後的 Val／PL／Sens；與 README.md 相同文字）
lines = [("景碩（Kinsus，3189.TW）收支與估值模型 v0.1（AI 半導體 Project：載板與 PCB 次產業——ABF／BT 載板第二梯隊；樣板＝nanya_pcb/、unimicron/）　2026-10-26", True),
 ("版本：v0.1（初版）。AI 加速器載板的每顆金額、Ibiden 份額與非 NVIDIA 權重讀 ibiden/dist，欣興份額讀 unimicron/dist，南電份額讀 nanya_pcb/dist（本專案權威）；四家份額合計在 Units 頁檢查。", False),
 ("模型命題：景碩的 ABF 載板營收＝MIN（需求, 產能）；需求＝NVIDIA CPU（Grace→Vera）載板（加速器顆數 × CPU 比例 × 每顆金額 × 景碩份額，公司目標 2027 年 80–100%）＋GPU 與 ASIC 加速器載板（與 Ibiden、欣興、南電同口徑，份額個位數）＋其他 AI＋非 AI ABF（伺服器與 PC CPU）；產能依公司規劃 4,000→8,000 萬顆／月（2029 年）。BT 與其他載板（約四成營收）以漲價為主；晶碩（隱形眼鏡）約 15% 營收、約 69% 淨利屬非控制權益。", False),
 ("核心問題：Vera CPU 載板份額、ABF 與 BT 漲價與 K6 新產能，能不能讓 EPS 從 FY2025 的 3.51 元升到 FY2027 二十多元；約 986 元的股價（近 12 月 P/E 166 倍、前瞻 P/E 45 倍）隱含什麼路徑。", False),
 ("INVEST", False),
 ("重點觀察 1：" + OBS[0], False), ("重點觀察 2：" + OBS[1], False), ("重點觀察 3：" + OBS[2], False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（延伸到 FY2032；PL 頁灰底）。金額：十億新台幣；EPS：新台幣（基本、歸屬母公司）。3Q26 財報約 10 月底～11 月初公布，本版只納入 3Q26 營收實績。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("加速器與 CPU 載板：markets/network 的 NVIDIA 與非 NVIDIA 顆數 × tsmc 同字母情境係數 × 每顆 NVIDIA 載板金額（Ibiden 模型 US$215 × 面積層數 +18%／年 × Ibiden 溢價指數）× 景碩份額（GPU 2%→6%、非 NVIDIA 金額 3%→6%、NVIDIA CPU 35%→70–75%；CPU 顆數＝NVIDIA × 0.4、CPU 載板金額＝GPU × 0.3）× 漲價年度化 1.15。", False),
 ("其他 AI 載板＝FY2026 AI 載板（ABF × 25%）− 加速器與 CPU 載板（校準殘差），之後 × 交換晶片與交換器市場成長 ^ 0.5 × 情境係數 × 溢價 × 漲價年度化；非 AI ABF 量 +6%／年 × 漲價年度化。ABF 產能指數 1.20／1.50／1.85（公司規劃，Interested-party）× 組合 +5%／年。BT 與其他載板年增依情境（A 28%／10%／5%、B 20%／5%／3%、C 12%／0%／−3%）；晶碩 +10%／年。", False),
 ("毛利＝FY2026 錨點＋ABF 現金毛利增量（50%）＋BT 增量（45%）＋晶碩增量（45%）＋漲價溢價營收 − 折舊增量；折舊＝前一年 ×（1 − 1／8 年）＋前一年資本支出 ÷ 8 年；資本支出 170／160／120 億（三情境相同）。營業費用彈性 0.6；業外每年 4 億；稅率 12%；非控制權益＝晶碩營收 × 25% × 68.8%；股數持平 5.27 億股。", False),
 ("校準：FY2026 ＝ 1Q26、2Q26 財報實績（EPS 1.17、2.57）＋3Q26 營收實績 145.4 億＋3Q26 毛利率 31%、4Q26 營收季增 12%、毛利率 32%（Assumed）。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（股東現金流＝每股（淨利＋折舊－資本支出），FY2027–32 明確期＋永續，r 10%、g 4%，三情境權重 0.3／0.5／0.2，負值取 0）＋0.5 × Comparable 目標價（欣興、南電、Ibiden、三星電機、AT&S 前瞻 P/E 中位數 × 12 個月後前瞻 EPS；無 A 股）；另列 Comparable 單獨目標價、市場共識目標價、差距拆解與上檔空間級距。十年 P/E 規則值 × FY2027 EPS 保留為參考（不是目標價）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("情境與行業一致：A／B／C 讀 tsmc 同字母顆數係數（＝markets/memory SCEN_UNITS，建置時核對）；交換器市場讀 markets/network（單一路徑 × 同字母係數）；公司特有變數只有 BT 與其他載板年增隨情境，A 最有利、C 最不利。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司數字經資料商與媒體轉載（公司原文待人工複核）；券商與資料商為 Interested-party。", False),
 ("", False),
 ("已知限制：(1) 公司不揭露 ABF／BT 與 AI 拆分，產品組合取媒體與法說摘要、AI 占比取高盛轉述；(2) 券商稱景碩 Rubin GPU 份額 15–20%，與 Ibiden、欣興、南電模型的份額合計（94–97%）衝突，本模型取剩餘空間（Sens 列高值）；(3) Vera CPU 份額與 CPU 載板金額比為公司說法與 Assumed；(4) 毛利用現金流通率近似，未建分廠稼動率；(5) 客戶贊助設備的金額與會計處理未知，資本支出可能高估；(6) 晶碩持股比例由非控制權益推得（約 31%），未取得原文；(7) 十年 P/E 2023–24 年約 900 倍、2019 年虧損不計；(8) FactSet 共識為 2026-06，主串提供的共識 EPS 口徑未查證；(9) 3Q26 財報（約 10 月底～11 月初）後更新錨點。", False)]
for i, (t, b) in enumerate(lines, 1):
    if t == "INVEST":
        RD.cell(i, 1, f'="投資建議："&Val!B{T["rec"]}&"（依上檔空間固定規則；加權目標價上檔空間 "&TEXT(Val!B{T["upw"]},"0.0%")&"，級距 "&Val!B{T["grade"]}&"）"')
    else:
        RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "kinsus")
wb.save(sys.argv[1])
if os.environ.get("KINSUS_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE, "UN_NOTE": UN_NOTE, "T": T, "T10": T10,
               "TPS0": TPS0, "TPS_LAST": TPS_LAST, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0, "FCF_ROW": FCF_ROW, "NCASES": len(cases)},
              open(os.environ["KINSUS_ROWMAP"], "w"), ensure_ascii=False, indent=1)

# 南電（8046.TW）收支與估值模型 v0.1.1（AI 半導體 Project：載板與 PCB 次產業；樣板 unimicron/）
# 用法：python3 nanya_pcb/build_nanya_pcb.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；EPS：新台幣（基本，歸屬母公司）。
# AI 加速器載板的份額口徑與 Ibiden、欣興一致：每顆 NVIDIA 載板金額、Ibiden 份額與非 NVIDIA 權重唯讀取自 ibiden/dist；欣興份額唯讀取自 unimicron/dist。
# 加速器顆數、AI IT 資本支出、交換晶片與交換器市場唯讀取自 markets/network/dist；情境顆數係數取自 tsmc/dist（與 markets/memory SCEN_UNITS 核對）；不修改上游。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
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
NET_F = only("markets/network/dist/*.xlsx"); TSMC_F = only("tsmc/dist/*.xlsx"); IB_F = only("ibiden/dist/*.xlsx"); UN_F = only("unimicron/dist/*.xlsx")
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

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B）；顆數：百萬顆；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間，不當點估計（放情境或 Sens）。SemiAnalysis 未使用。公司數字經資料商（stockanalysis）與媒體轉載；公司說法為 Interested-party；券商、媒體、資料商為 Interested-party"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
SA = "stockanalysis.com（8046）"
SAQ = "stockanalysis 季損益表（公司財報轉載）"
FG1Q = "富果 1Q26 法說摘要（2026-06-12，AI 摘要）"
VC4Q = "vocus 4Q25 法說摘要（2026-03-17）"
SETN7 = "三立 2026-07-27（2Q26 營收、法人預估、產能）"
KGI = "凱基 PCB 產業報告 2026-07-20（Interested-party）"
BOFA = "TechNews 2026-10-02（美銀報告）"
DAIWA = "TechNews 2026-10-05（大和報告；8 月自結獲利）"
TN9 = "TechNews 2026-10-03（9 月營收、前 9 月累計）"
WV7 = "winvest 2026-08（7 月營收、1Q26 EPS）"
BRO = "brodoginvestment 南電深度分析（引述券商報告；Interested-party）"
CNYES_FS = "鉅亨 2026-06-24（FactSet 14 位分析師）"

sec("A. AI 加速器顆數、AI IT 資本支出與交換器市場（讀 markets/network、tsmc 成品；唯讀，不自行重估）")
for g in NV + OT:
    for y in (26, 27, 28):
        v, lab, src = UNITS[f"u_{g}_{y}"]
        inp(f"u_{g}_{y}", lab, "百萬顆", None, v, None, "Derived", src, "唯讀上游；與 Ibiden、欣興、台積電模型同一份顆數")
    v, lab, src = UNITS[f"g29_{g}"]
    inp(f"g29_{g}", lab, "%", None, v, None, "Derived", src, "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}", "US$B", None, ITC[0][j], None, "Derived", ITC[1], "Units 頁每 GW 換算" if y == 26 else "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"swm_{y}", f"交換晶片與交換器市場 20{y}（scale-up＋scale-out）", "US$B", None, SWM[0][j], None, "Derived", SWM[1], "南電 AI 載板以網通交換器為主（800G／1.6T 交換晶片）；其他 AI 載板的成長依據" if y == 26 else "同上")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝順利（基準）；B＝平平；C＝不如預期；與台積電、Ibiden、欣興模型相同（＝markets/memory SCEN_UNITS，建置時核對）" if s == "A" else "同上")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")

sec("B. 每顆加速器載板金額與份額（口徑與 Ibiden、欣興一致；兩者數字唯讀取自各自 dist）")
inp("v26", "FY2026 每顆 NVIDIA 加速器載板金額（全市場；Ibiden 模型校準）", "US$", None, round(IB_V26[0], 6), None, "Derived", IB_V26[1], "Ibiden 模型：Ibiden 每顆營收 ÷ 份額 70%（BofA，Interested-party）")
inp("g_asp", "每顆載板金額年增（面積、層數；Ibiden 模型）", "%", IB["g_asp"][0], IB["g_asp"][1], IB["g_asp"][2], "Analogy", IB["g_asp"][3], "Ibiden 路線圖 2026 年 90×90 mm → 2028 年 110×110 mm（經 Ibiden 模型）")
for y in (27, 28, 29):
    k = f"pp_{y}"
    inp(f"ibpp_{y}", f"ABF 價格溢價指數 20{y}（FY2026＝1；Ibiden 模型）", "x", IB[k][0], IB[k][1], IB[k][2], "Assumed", IB[k][3], "全市場短缺溢價；新產能到位後回落（Ibiden 模型；欣興模型同值）" if y == 27 else "同上")
    k = f"rel_{y}"
    inp(f"ibrel_{y}", f"Ibiden NVIDIA 份額 20{y} ÷ 2026（Ibiden 模型）", "x", None, IB[k][1], None, "Assumed", IB[k][3], "Units 頁份額合計檢查用")
inp("ib_sh26", "Ibiden 占 NVIDIA 加速器載板 FY2026（Ibiden 模型假設）", "%", *SIL.vals("abf_ibsh26"), "Interested-party", IB_V26[1] + "（列標籤：假設 FY2026 份額 70%）", "BofA：Ibiden 70–80%")
inp("ib_w_ot", "Ibiden 非 NVIDIA 權重（份額 × 每顆金額，相對 NVIDIA；Ibiden 模型）", "x", IB["w_ot"][0], IB["w_ot"][1], IB["w_ot"][2], "Assumed", IB["w_ot"][3], "Units 頁：÷ v_ot 即 Ibiden 非 NVIDIA 份額")
inp("v_ot", "非 NVIDIA 加速器每顆載板金額 ÷ NVIDIA（欣興模型同值）", "x", *SIL.vals("abf_vot"), "Assumed", "TPU、Trainium、MI 系列封裝面積接近 GPU；量化：未找到", "與欣興模型 v_ot 相同")
for y in (26, 27, 28, 29):
    inp(f"un_nv_{y}", f"欣興占 NVIDIA 加速器載板 20{y}（欣興模型）", "%", None, UNS[f"sh_nv_{y}"][0], None, "Assumed", UNS[f"sh_nv_{y}"][1], "Units 頁份額合計檢查用")
    inp(f"un_ot_{y}", f"欣興占非 NVIDIA 加速器載板（金額）20{y}（欣興模型）", "%", None, UNS[f"sh_ot_{y}"][0], None, "Assumed", UNS[f"sh_ot_{y}"][1], "同上")
for y in (26, 27, 28, 29):
    inp(f"sh_nv_{y}", *SIL.row9(f"abf_np_nv_{y}")[1:])
for y in (26, 27, 28, 29):
    inp(f"sh_ot_{y}", *SIL.row9(f"abf_np_ot_{y}")[1:])

sec("C. FY2026 產品組合與 AI 占比（公司未按季揭露；券商與法說摘要推全年）")
inp("mx_abf", "FY2026 ABF 載板 ÷ 營收", "%", 0.53, 0.58, 0.62, "Interested-party", KGI + "：1Q26 ABF 約 55–60%、BT 約 30–35%", VC4Q + "：2025 年 ABF 50–55%、BT 約 30%、PCB 10–15%")
inp("mx_bt", "FY2026 BT 載板 ÷ 營收（其餘為 PCB）", "%", 0.28, 0.32, 0.35, "Interested-party", KGI, BRO + "：ABF 約 60%、BT 約 30%")
inp("ai_abf", "FY2026 AI 與資料中心 ÷ ABF 營收", "%", 0.55, 0.65, 0.75, "Interested-party", KGI + "：AI 與資料中心（800G 交換器）約占 ABF 營收 60%", FG1Q + "：1Q26 應用別網通 49%、AI 與 HPC 16%（全公司）")

sec("D. 需求與產能（FY2027 起）")
inp("el_o", "其他 AI 載板（交換器、伺服器 CPU、網通 ASIC）對交換晶片與交換器市場的彈性", "x", 0.3, 0.5, 0.8, "Assumed", "其他 AI 載板＝FY2026 AI 載板 − 加速器載板（校準殘差）；隨交換器市場成長但低於等比例（單價已由溢價另計）", "研究報告 A：南電占高階 800G／1.6T 交換晶片 ABF 逾 70%（" + BRO + "）")
inp("g_nai", "非 AI ABF（PC、一般伺服器、車用）年增", "%", -0.02, 0.03, 0.06, "Assumed", FG1Q + "：電腦 17%（零組件缺貨限制）、車用 6%（需求弱）", "—")
for y in (27, 28, 29):
    inp(f"cap_{y}", *SIL.row9(f"abfcap_np_{y}")[1:])
inp("g_mix", "每單位產能營收年增（層數、尺寸組合）", "%", 0.0, 0.05, 0.08, "Assumed", "公司：產品朝 140 mm 以上大尺寸、30 層以上（2028 年後）發展（" + VC4Q + "）", "與欣興模型同值")
inp("carry", "南電漲價年度化（CY2026 基期含上半年低價；FY2027 起 ×(1＋此值)）", "%", 0.08, 0.15, 0.22, "Assumed",
    "FY2026 毛利率由 1Q26 15.8% 升到 4Q26 約 32%（美銀估），季營收由 111.8 億升到約 180 億；4Q26 年化 ÷ FY2026 營收約 1.22（Q 頁，Derived；含量增）", "研究報告：ABF 每平方呎單價 2026 +35%、2027 +57%（" + BRO + "，Interested-party）")
for s, vs in (("A", (0.30, 0.10, 0.05)), ("B", (0.25, 0.05, 0.03)), ("C", (0.15, 0.0, 0.0))):
    for y, v in zip((27, 28, 29), vs):
        inp(f"g_bt{y}_{s}", f"BT 載板年增 FY20{y}（情境 {s}；公司特有，A 最有利）", "%", None, v, None, "Assumed",
            BOFA + "：BT 4Q26 漲價個位數、1Q27 再漲 20% 以上" if (s == "A" and y == 27) else "同上",
            DAIWA + "：BT 1Q27 季增最高 20%；台、韓部分廠商縮減 BT 產能（" + KGI + "）" if (s == "A" and y == 27) else "—")
inp("g_pcb", "PCB（一般電路板）年增", "%", -0.05, 0.0, 0.03, "Assumed", KGI + "：PCB 以消費型為主且虧損，公司計畫產能轉向 BT", "—")

sec("E. 毛利、折舊、費用與淨利（FY2027 起）")
inp("ft_abf", "ABF 營收增量的現金毛利率（扣除漲價溢價後）", "%", 0.40, 0.50, 0.60, "Assumed", BOFA + "：2027 年 ABF 毛利率約 50%", "與欣興模型同值")
inp("ft_oth", "其他產品（BT、PCB）營收增量的現金毛利率", "%", 0.25, 0.40, 0.55, "Assumed", "BT 成長以漲價為主（漲價幾乎全數落在毛利）；PCB 虧損", "—")
for y, v, lo, hi in ((27, 14.0, 6.0, 20.0), (28, 18.0, 6.0, 24.0), (29, 12.0, 6.0, 20.0)):
    inp(f"cx{y}", f"資本支出 FY20{y}（三情境相同；公司特有）", "NT$B", lo, v, hi, "Assumed",
        KGI + "：2025–28 年資本支出約 US$12.8 億（約 416 億新台幣；七家載板廠第六）" if y == 27 else "同上",
        "公司：2026 年約 40 億，新廠待董事會決議（" + FG1Q + "）" if y == 27 else "—")
inp("capex26", "FY2026 資本支出（公司計畫）", "NT$B", None, 4.0, None, "Interested-party", FG1Q + "：2026 年約 40 億", "TTM 至 2026-06：19.3 億（stockanalysis 現金流量表）")
inp("da26", "FY2026 折舊攤提", "NT$B", None, 7.1, None, "Interested-party", "stockanalysis 現金流量表：TTM 至 2026-06 71.0 億", "研究報告 C：2026 年 71 億、2028 年 58 億（" + BRO + "）")
inp("life", "新增資產平均折舊年限", "年", 6, 8, 10, "Assumed", "載板設備與廠房混合；與欣興模型同值", "會計年限：未找到")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.3, 0.6, 1.0, "Assumed", "1Q26 營業費用 4.19 億、2Q26 4.61 億（營收季增 21%；Q 頁推得）", "—")
for y in (27, 28, 29):
    inp(f"nonop{y}", f"營業外淨額 FY20{y}", "NT$B", 0.0, 0.6, 1.2, "Assumed",
        "1Q26 2.5 億、2Q26 0.8 億（stockanalysis 推得；含匯兌）" if y == 27 else "同上", "利息收入（現金部位）（Assumed）" if y == 27 else "—")
inp("tax", "所得稅 ÷ 稅前利益", "%", 0.18, 0.21, 0.24, "Derived", "1H26：稅前 45.7、淨利 35.7 億 → 21.9%（stockanalysis 推得）", "2Q26 24.0%、1Q26 18.1%；營所稅 20%")
inp("sh_g", "股數年變動", "%", 0.0, 0.0, 0.01, "Assumed", "流通股數 6.4617 億股（stockanalysis）", "無可轉債資訊：未找到")

sec("F. FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設）")
inp("q3_rev", "3Q26 營收（月營收實績加總）", "NT$B", None, 16.379, None, "Verified", TN9 + "：前 9 月 411.31 億 − 1H26 247.52 億", WV7 + "：7 月 54.40 億；" + DAIWA + "：8 月 53.53 億；9 月 55.86 億")
inp("q3_gm", "3Q26 毛利率", "%", 0.28, 0.30, 0.315, "Assumed", DAIWA + "：8 月自結稅前淨利率 27.8%、毛利率推算約 30–31%", SETN7 + "：法人估 28.1%")
inp("q3_opx", "3Q26 營業費用", "NT$B", 0.46, 0.50, 0.55, "Assumed", "1Q26 0.419、2Q26 0.461（Q 頁推得）", "—")
inp("q4_g", "4Q26 營收季增", "%", 0.03, 0.10, 0.15, "Assumed", BOFA + "：4Q26 季增約 10%", DAIWA + "：法人估 4Q26 184.1 億（季增 12%）")
inp("q4_gm", "4Q26 毛利率", "%", 0.29, 0.325, 0.34, "Assumed", BOFA + "：4Q26 毛利率 32–33%", "—")
inp("q4_opx", "4Q26 營業費用", "NT$B", 0.48, 0.52, 0.58, "Assumed", "同上", "—")
inp("q_nonop", "3Q、4Q26 營業外淨額（每季）", "NT$B", 0.0, 0.15, 0.40, "Assumed", "1Q26 0.246、2Q26 0.077（stockanalysis 推得）", "3Q26 財報約 10 月底（stockanalysis 估 2026-11-03）")
inp("q_sh", "3Q、4Q26 加權股數", "十億股", None, 0.64617, None, "Interested-party", SA + "：流通股數 646.17M", KGI + "：約 6.46 億股")

sec("G. 共識與券商（只作 Recon 與 Val 對照，不作校準）")
inp("cons_eps26", "共識 EPS FY2026（FactSet 平均，2026-06）", "NT$", 8.56, 13.19, 20.48, "Interested-party", CNYES_FS, "中位數 12.57；凱基 2026-07：17.27；美銀 2026-10：18.44")
inp("cons_eps27", "共識 EPS FY2027（FactSet 平均，2026-06）", "NT$", 21.66, 32.12, 52.88, "Interested-party", CNYES_FS, "中位數 28.9；凱基 38.88；美銀 43.80")
inp("cons_eps28", "共識 EPS FY2028（FactSet 平均，2026-06）", "NT$", 42.03, 59.21, 80.8, "Interested-party", CNYES_FS, "中位數 54.25；凱基 61.35；美銀 70.09")
inp("cons_rev27", "共識營收 FY2027（FactSet 平均，2026-06）", "NT$B", 69.31, 90.26, 127.93, "Interested-party", CNYES_FS, "美銀 90.33；凱基 88.41")
inp("cons_rev28", "共識營收 FY2028（FactSet 平均，2026-06）", "NT$B", 109.36, 133.54, 178.55, "Interested-party", CNYES_FS, "美銀 120.65；凱基 123.31")
inp("bk_eps27", "券商 EPS FY2027（美銀 2026-10 43.80、凱基 2026-07 38.88 平均）", "NT$", 38.88, 41.34, 43.80, "Interested-party", BOFA, KGI)
inp("bk_eps28", "券商 EPS FY2028（美銀 70.09、凱基 61.35 平均）", "NT$", 61.35, 65.72, 70.09, "Interested-party", BOFA, KGI)

sec("H. 價格與市場資料")
inp("px", "股價（2026-10-22 收盤；10-23 收盤未取得）", "NT$", None, 1265, None, "Interested-party", SA + " 歷史價格：10-22 收盤 1,265、10-21 1,340", "財報狗、HiStock 2026-10-22：1,265")
inp("fpe", "前瞻 P/E（stockanalysis，價格日期 2026-09-07）", "x", None, 33.73, None, "Interested-party", SA + "：頁面價格 1,050（2026-09-07）", "口徑：頁面未明示（約未來 12 月）；10 月更新值未取得")
inp("pe_ttm", "近 12 月 P/E（財報狗，10-22）", "x", None, 148.7, None, "Interested-party", "財報狗 2026-10-22", "HiStock 148.65")
inp("cons_feps", "共識前瞻 12 月 EPS（stockanalysis 頁面價格 1,050 ÷ 前瞻 P/E）", "NT$", None, "=1050/fpe", None, "Derived", SA, "以頁面同日價格推算（不用 10-22 現價）；直接數字：未找到")
inp("cons_tp", "市場共識目標價（賣方平均）", "NT$", None, 1651.31, None, "Interested-party", SA + "：14 位分析師平均 1,651.31（頁面價格日期 2026-09-07）", "10 月：美銀 2,000、大和 2,460；凱基 2026-07：1,480")
inp("fx", "匯率", "NT$／US$", None, SIL.base("fx_twd"), None, "Assumed", "Siliconomics SIF_fx_twd（全庫共用；v0.8.2 前本模型取 32.5）", "—")
REV5 = [52.228, 64.647, 42.253, 32.283, 40.173]
NI5 = [10.582, 19.416, 5.817, 0.204, 1.947]
inp("nim5", "近五年平均淨利率（FY2021–25；十年平均的參考）", "%", None, "=AVERAGE(" + ",".join(f"{n}/{v}" for n, v in zip(NI5, REV5)) + ")", None, "Derived",
    SA + "：營收與淨利 FY2021–25", "十年（2016–2025）營收：未找到（FY2016–20 營收未取得），收斂法依工作單改用 FY2029 × 0.8")

sec("I. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_r", "折現率／要求報酬", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%，美台日韓同一值）", "—")
inp("tp_g", "長期成長率", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "永續價值與 FY2032 收斂目標（情境 B）")
inp("tp_roe", "ROE（近三年平均，FY2023–25）", "%", None, "=AVERAGE(0.1145,0.0044,0.0425)", None, "Derived", SA + " 比率頁：FY2023 11.45%、FY2024 0.44%、FY2025 4.25%", "單一來源；本模型 DCF 改用（淨利＋折舊－資本支出），此值只作參考列")
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
Q["A2"] = "FY2026 錨點＝1Q、2Q26 財報實績＋3Q26 營收實績（月營收）與 3Q26 毛利率、費用假設＋4Q26 假設。3Q26 財報約 10 月底公布，屆時以實績取代 E 欄（毛利率、費用、業外、淨利）"
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
qrow(4, "營收", {"B": 40.173, "C": 11.177, "D": 13.575, "E": "=q3_rev", "F": "=E4*(1+q4_g)"}, tag="Verified／Assumed（4Q）", s2="FY2025 401.73 億（vocus 4Q25 法說、FactSet 同）；1Q26 111.77 億（富果同）；2Q26 135.75 億（大和、三立同）")
qrow(5, "毛利", {"B": 3.647, "C": 1.771, "D": 3.360, "E": "=E4*q3_gm", "F": "=F4*q4_gm"}, tag="Verified／Assumed", s2="1Q26 毛利率 15.8%（三立同；富果摘要 16.8% 與此不符，不採用）")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEFG"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2="FY2025 9.08%（凱基、美銀 9.1%）")
qrow(7, "營業費用", {c: f"={c}5-{c}8" for c in "BCD"} | {"E": "=q3_opx", "F": "=q4_opx"}, tag="Derived／Assumed", s1="毛利 − 營業利益")
qrow(8, "營業利益", {"B": 1.982, "C": 1.352, "D": 2.899, "E": "=E5-E7", "F": "=F5-F7"}, tag="Verified／Derived", s2="FY2025 19.82 億（vocus 同）；1Q26 13.52 億（富果同）")
qrow(9, "營業外淨額", {"B": 0.364, "C": 0.246, "D": 0.077, "E": "=q_nonop", "F": "=q_nonop"}, tag="Derived／Assumed", s2="稅前利益 − 營業利益；稅前利益由 stockanalysis 稅前利益率 × 營收推得（近似）")
qrow(10, "稅前利益", {c: f"={c}8+{c}9" for c in "BCDEF"}, tag="Derived", s1="—", s2="")
qrow(11, "歸屬母公司淨利", {"B": 1.947, "C": 1.309, "D": 2.262, "E": "=E10*(1-tax)", "F": "=F10*(1-tax)"}, tag="Verified／Derived", s2="FY2025 19.47 億（FactSet 表同）；2Q26 22.61 億（大和同）")
qrow(12, "所得稅", {c: f"={c}10-{c}11" for c in "BCDEF"}, tag="Derived", s1="—")
qrow(13, "加權股數（十億股）", {"B": "=B11/B14", "C": "=C11/C14", "D": "=D11/D14", "E": "=q_sh", "F": "=q_sh"}, ann="=AVERAGE(C13:F13)", tag="Derived", s1="淨利 ÷ EPS")
qrow(14, "EPS（基本）", {"B": 3.01, "C": 2.03, "D": 3.50, "E": "=E11/E13", "F": "=F11/F13"}, fmt="0.00", ann="=G11/G13", tag="Verified／Derived", s2="FY2025 3.01（wantgoo 同）；1Q26 2.03（winvest 同）；2Q26 3.50（大和、HiStock 同）")
qrow(15, "折舊攤提（年度）", {"B": 6.894, "G": "=da26"}, ann="", tag="Verified／Interested-party", s1="stockanalysis 現金流量表", s2="FY2026：TTM 至 2026-06 71.0 億")
qrow(16, "資本支出（年度）", {"B": 2.422, "G": "=capex26"}, ann="", tag="Verified／Interested-party", s1="FY2025：stockanalysis 現金流量表；FY2026：公司計畫約 40 億", s2="富果摘要稱 2025 年實際 43.14 億（口徑可能含預付設備款；不採用）")
Q.cell(18, 1, "FY2026 產品別（錨點營收 × In 頁 C 節占比）").font = BOLD
for i, (lab, f) in enumerate((("ABF 載板", "=G4*mx_abf"), ("BT 載板", "=G4*mx_bt"), ("PCB（一般電路板）", "=G4*(1-mx_abf-mx_bt)")), 19):
    Q.cell(i, 1, lab); Q[f"G{i}"] = f; Q[f"G{i}"].number_format = "#,##0.000"; Q.cell(i, 8, "Derived")
Q.cell(23, 1, "檢查（參考；校準未使用券商數字）").font = BOLD
checks = [("1H26 營收 自算 − 前 9 月累計 41.131 + 3Q26（應為 0）", "=C4+D4-(41.131-q3_rev)", "0.000"),
          ("FY2026 營收年增", "=G4/B4-1", "0.0%"),
          ("FY2025→FY2026 現金毛利增量 ÷ 營收增量（含漲價；In 頁 ft_abf、ft_oth 的對照）", "=((G5+G15)-(B5+B15))/(G4-B4)", "0.0%"),
          ("FY2026 EPS ÷ 凱基 17.27 − 1", "=G14/17.27-1", "0.0%"),
          ("3Q26 EPS ÷ 大和估 5.62 − 1", "=E14/5.62-1", "0.0%"),
          ("4Q26 營收年化 ÷ FY2026 營收（漲價年度化 carry 的對照）", "=F4*4/G4", "0.00")]
for i, (lab, f, fm) in enumerate(checks, 24):
    Q.cell(i, 1, lab); Q[f"B{i}"] = f; Q[f"B{i}"].number_format = fm
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 7), ("oi", 8), ("nonop", 9), ("ni", 11), ("sh", 13), ("eps", 14), ("da", 15), ("cx", 16)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$G${rr}")
for nm, rr in (("abf", 19), ("bt", 20), ("pcb", 21)):
    defname(f"a26_{nm}", f"Q!$G${rr}")
defname("a26_shend", "Q!$F$13")
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
DEFP = dict(scn="A", ga="g_asp", pp="ibpp_{y}", vot="v_ot", shnv="sh_nv_{y}", shot="sh_ot_{y}", aiabf="ai_abf", elo="el_o", gnai="g_nai",
            cap="cap_{y}", gmix="g_mix", carry="carry", gbt="g_bt{y}_{s}", gpcb="g_pcb", ftabf="ft_abf", ftoth="ft_oth",
            cx="cx{y}", life="life", phi="phi", nonop="nonop{y}", tax="tax", mf="1", fg="fade_g{s}", fm="fade_m{s}")
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
    sec_("A. AI 加速器載板（口徑與 Ibiden、欣興一致；D 欄＝FY2026 校準年）")
    row("unv", "NVIDIA 加速器顆數 × 情境係數", "百萬顆", {FC[y]: "=(" + "+".join(units_f(g, y) for g in NV) + f")*{um(y)}{mf(y)}" for y in BU}, note="markets/network In：Blackwell、Rubin、Rubin Ultra、Hopper")
    row("uot", "非 NVIDIA 加速器顆數 × 情境係數", "百萬顆", {FC[y]: "=(" + "+".join(units_f(g, y) for g in OT) + f")*{um(y)}{mf(y)}" for y in BU}, note="AMD、TPU、Trainium、其他 ASIC")
    row("idx", "每顆載板面積與層數指數（FY2026＝1）", "x", {"D": "=1", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['ga']})" for y in F3}}, "0.000", note="Ibiden 模型 g_asp（Analogy）")
    row("pp", "ABF 價格溢價指數（全市場；FY2026＝1）", "x", {"D": "=1", **{FC[y]: f"={p('pp', y)}" for y in F3}}, "0.00", note="Ibiden 模型 pp（Assumed）")
    row("cy", "南電漲價年度化（CY2026 基期；FY2026＝1）", "x", {"D": "=1", **{FC[y]: f"=1+{P['carry']}" for y in F3}}, "0.00", note="In 頁 carry（Assumed）")
    row("v", "每顆 NVIDIA 加速器載板金額（全市場）＝ FY2026 金額 × 面積層數 × 溢價", "US$／顆", {FC[y]: f"=v26*{FC[y]}{R['idx']}*{FC[y]}{R['pp']}" for y in BU}, "#,##0.0", note="FY2026 金額取自 Ibiden 模型 Units 頁（權威）")
    row("shnv", "南電份額：NVIDIA 加速器載板", "%", {FC[y]: f"={p('shnv', y)}" for y in BU}, "0.0%")
    row("shot", "南電份額：非 NVIDIA 加速器載板（金額）", "%", {FC[y]: f"={p('shot', y)}" for y in BU}, "0.0%")
    row("accu", "南電加速器載板營收 ＝（NVIDIA × 金額 × 份額 ＋ 非 NVIDIA × 金額 × v_ot × 份額）× 漲價年度化", "US$M",
        {FC[y]: f"=({FC[y]}{R['unv']}*{FC[y]}{R['v']}*{FC[y]}{R['shnv']}+{FC[y]}{R['uot']}*{FC[y]}{R['v']}*{P['vot']}*{FC[y]}{R['shot']})*{FC[y]}{R['cy']}" for y in BU}, "#,##0")
    row("acc", "南電加速器載板營收（新台幣）", "NT$B", {FC[y]: f"={FC[y]}{R['accu']}*fx/1000" for y in BU}, "#,##0.00")
    sec_("B. ABF 需求與產能（營收＝MIN（需求, 產能））")
    row("ai0", "AI 載板需求（FY2026＝ABF × AI 占比）", "NT$B", {"D": f"=a26_abf*{P['aiabf']}"}, "#,##0.00", note="凱基：AI 與資料中心（800G 交換器）約占 ABF 60%")
    row("oai", "其他 AI 載板需求（交換器為主，另含伺服器 CPU、網通 ASIC）", "NT$B", {"D": f"=D{R['ai0']}-D{R['acc']}",
        **{FC[y]: f"=$D{r}*(swm_{y % 100}/swm_26)^{P['elo']}*{um(y)}{mf(y)}*{FC[y]}{R['pp']}*{FC[y]}{R['cy']}" for y in F3}}, "#,##0.00",
        note="FY2026＝校準殘差（Derived）；之後 × 交換晶片與交換器市場成長 ^ 彈性 × 情境係數 × 溢價")
    row("nai", "非 AI ABF 需求（PC、一般伺服器、車用）", "NT$B", {"D": f"=a26_abf-D{R['ai0']}", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['gnai']})" for y in F3}}, "#,##0.00")
    row("dem", "ABF 需求合計", "NT$B", {FC[y]: f"={FC[y]}{R['acc']}+{FC[y]}{R['oai']}+{FC[y]}{R['nai']}" for y in BU}, "#,##0.00")
    row("cap", "ABF 產能上限（FY2026 營收 × 產能指數 × 組合 × 溢價）", "NT$B", {"D": "=a26_abf", **{FC[y]: f"=$D{R['dem']}*{p('cap', y)}*(1+{P['gmix']})^{y - 2026}*{FC[y]}{R['pp']}*{FC[y]}{R['cy']}" for y in F3}}, "#,##0.00",
        note="1H26 ABF 稼動率約 80–85%（昆山約 50%），以 FY2026 營收為產能基準")
    row("abf", "ABF 載板營收 ＝ MIN（需求, 產能）", "NT$B", {"D": "=a26_abf", **{FC[y]: f"=MIN({FC[y]}{R['dem']},{FC[y]}{R['cap']})" for y in F3}}, "#,##0.00")
    row("fill", "需求滿足率（營收 ÷ 需求；< 100% 表示產能受限）", "%", {FC[y]: f"={FC[y]}{R['abf']}/{FC[y]}{R['dem']}" for y in BU}, "0.0%")
    row("aisub", "AI 載板營收（加速器＋其他 AI，乘滿足率）", "NT$B", {FC[y]: f"=({FC[y]}{R['acc']}+{FC[y]}{R['oai']})*{FC[y]}{R['fill']}" for y in BU}, "#,##0.00")
    row("accsh", "加速器載板 ÷ AI 載板", "%", {FC[y]: f"={FC[y]}{R['acc']}/({FC[y]}{R['acc']}+{FC[y]}{R['oai']})" for y in BU}, "0.0%")
    sec_("C. 其他產品與合計")
    row("bt", "BT 載板（情境別年增；公司特有）", "NT$B", {"D": "=a26_bt", **{FC[y]: f"={FC[y - 1]}{r}*(1+{p('gbt', y)})" for y in F3}}, "#,##0.00", note="BT 成長以漲價為主")
    row("pcb", "PCB（一般電路板）", "NT$B", {"D": "=a26_pcb", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['gpcb']})" for y in F3}}, "#,##0.00")
    row("rev", "營收", "NT$B", {"C": "=a25_rev", **{FC[y]: f"={FC[y]}{R['abf']}+{FC[y]}{R['bt']}+{FC[y]}{R['pcb']}" for y in BU}}, "#,##0.0")
    row("g", "營收年增", "%", {FC[y]: f"={FC[y]}{R['rev']}/{FC[y - 1]}{R['rev']}-1" for y in FYS[1:]}, "0.0%")
    row("abfsh", "ABF ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['abf']}/{FC[y]}{R['rev']}" for y in BU}, "0.0%")
    row("aish", "AI 載板 ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['aisub']}/{FC[y]}{R['rev']}" for y in BU}, "0.0%", note="公司：1Q26 網通 49%、AI 與 HPC 16%（應用別口徑不同）")
    sec_("D. 折舊、毛利與損益")
    row("cx", "資本支出", "NT$B", {"C": "=a25_cx", "D": "=capex26", **{FC[y]: f"={p('cx', y)}" for y in F3}}, "#,##0.0")
    row("da", "折舊攤提 ＝ 前一年 ×（1 − 1／年限）＋ 前一年資本支出 ÷ 年限", "NT$B", {"C": "=a25_da", "D": "=da26",
        **{FC[y]: f"={FC[y - 1]}{r}*(1-1/{P['life']})+{FC[y - 1]}{R['cx']}/{P['life']}" for y in F3}}, "#,##0.0", note="新產能前一年投入、當年開始折舊（近似）")
    row("dar", "折舊 ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['da']}/{FC[y]}{R['rev']}" for y in FYS}, "0.0%")
    row("prem", "漲價溢價營收（相對 FY2026 價格；全數落在毛利）", "NT$B", {FC[y]: f"={FC[y]}{R['aisub']}*(1-1/({FC[y]}{R['pp']}*{FC[y]}{R['cy']}))" for y in F3}, "#,##0.00",
        note="AI 載板營收 ×（1 − 1／價格指數）")
    row("gp", "毛利 ＝ FY2026 ＋ ABF 現金毛利增量 ＋ 其他產品現金毛利增量 ＋ 溢價 − 折舊增量", "NT$B", {"C": "=a25_gp", "D": "=a26_gp",
        **{FC[y]: (f"=$D{r}+{P['ftabf']}*({FC[y]}{R['abf']}-{FC[y]}{R['prem']}-$D{R['abf']})+{P['ftoth']}*(({FC[y]}{R['rev']}-{FC[y]}{R['abf']})-($D{R['rev']}-$D{R['abf']}))"
                   f"+{FC[y]}{R['prem']}-({FC[y]}{R['da']}-$D{R['da']})") for y in F3}}, "#,##0.00", note="流通率為現金毛利口徑；折舊另計")
    row("gm", "毛利率", "%", {FC[y]: f"={FC[y]}{R['gp']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("opx", "營業費用", "NT$B", {"C": "=a25_opx", "D": "=a26_opx", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['phi']}*{FC[y]}{R['g']})" for y in F3}}, "#,##0.00")
    row("oi", "營業利益", "NT$B", {FC[y]: f"={FC[y]}{R['gp']}-{FC[y]}{R['opx']}" for y in FYS}, "#,##0.00")
    row("oim", "營業利益率", "%", {FC[y]: f"={FC[y]}{R['oi']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("nonop", "營業外淨額", "NT$B", {"C": "=a25_nonop", "D": "=a26_nonop", **{FC[y]: f"={p('nonop', y)}" for y in F3}}, "#,##0.00")
    row("pti", "稅前利益", "NT$B", {FC[y]: f"={FC[y]}{R['oi']}+{FC[y]}{R['nonop']}" for y in FYS}, "#,##0.00")
    row("tax", "所得稅", "NT$B", {FC[y]: f"={FC[y]}{R['pti']}*{P['tax']}" for y in F3}, "#,##0.00")
    row("ni", "歸屬母公司淨利", "NT$B", {"C": "=a25_ni", "D": "=a26_ni", **{FC[y]: f"={FC[y]}{R['pti']}-{FC[y]}{R['tax']}" for y in F3}}, "#,##0.00")
    row("shr", "加權股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", "E": "=a26_shend*(1+sh_g)", **{FC[y]: f"={FC[y - 1]}{r}*(1+sh_g)" for y in (2028, 2029)}}, "0.0000")
    row("eps", "EPS（基本）", "NT$", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['shr']}" for y in FYS}, "0.00")
    row("epsg", "EPS 年增", "%", {FC[y]: f"={FC[y]}{R['eps']}/{FC[y - 1]}{R['eps']}-1" for y in FYS[1:]}, "0.0%")
    row("nim", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("cxr", "資本支出 ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['cx']}/{FC[y]}{R['rev']}" for y in FYS}, "0.0%")
    row("fcf", "股東現金流 ＝ 淨利 ＋ 折舊 − 資本支出", "NT$B", {FC[y]: f"={FC[y]}{R['ni']}+{FC[y]}{R['da']}-{FC[y]}{R['cx']}" for y in F3}, "#,##0.00",
        note="不含營運資金變動")
    row("fps", "每股股東現金流", "NT$", {FC[y]: f"={FC[y]}{R['fcf']}/{FC[y]}{R['shr']}" for y in F3}, "0.00")
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
PL["A1"] = "PL — 南電損益（十億新台幣；EPS 新台幣、基本；歸屬母公司）。FY2025A 實績、FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設；三情境相同），FY2027–FY2029 由驅動因子推導，FY2030 起為收斂法。三情境各一個完整推導區塊（本頁下方），頂端為摘要"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A（順利）：tsmc 顆數係數 A（基準）；BT 年增 30%／10%／5%",
         "B": "情境 B（平平）：tsmc 顆數係數 B（CY2028 ×0.8、CY2029 ×0.85）；BT 年增 25%／5%／3%",
         "C": "情境 C（不如預期）：tsmc 顆數係數 C（CY2028–29 ×0.6）；BT 年增 15%／0%／0%"}
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

# ---------------- Units：份額合計檢查（與 Ibiden、欣興同口徑） ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 南電在 AI 加速器載板的份額與每顆金額（口徑與 Ibiden、欣興模型一致；情境 A）。份額合計＝Ibiden（ibiden/dist）＋欣興（unimicron/dist）＋南電，其餘為三星電機、景碩、AT&S 等"
A_ = PB["A"]; cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ibrel = lambda y: "1" if y == 2026 else f"ibrel_{y % 100}"
UKEYS = ["unv", "uot", "v", "shnv", "ibnv", "unnv", "sumnv", "shot", "ibot", "unot", "sumot", "pu", "accu", "aisub", "itr", "pgw"]
UR = {k: 4 + i for i, k in enumerate(UKEYS)}
UDEF = {
    "unv": ("NVIDIA 加速器顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['unv']}", "#,##0.00", "markets/network"),
    "uot": ("非 NVIDIA 加速器顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['uot']}", "#,##0.00", ""),
    "v": ("每顆 NVIDIA 加速器載板金額（全市場）", "US$", lambda y: f"=PL!{FC[y]}{A_['v']}", "#,##0", "Ibiden 模型 Units 頁同值（FY2026 校準 × 面積層數 × 溢價）"),
    "shnv": ("南電份額：NVIDIA", "%", lambda y: f"=PL!{FC[y]}{A_['shnv']}", "0.0%", "Assumed"),
    "ibnv": ("Ibiden 份額：NVIDIA（70% × Ibiden 份額相對值）", "%", lambda y: f"=ib_sh26*{ibrel(y)}", "0.0%", "Ibiden 模型"),
    "unnv": ("欣興份額：NVIDIA", "%", lambda y: f"=un_nv_{y % 100}", "0.0%", "欣興模型"),
    "sumnv": ("合計：Ibiden＋欣興＋南電（NVIDIA）", "%", lambda y: f"={cyc(y)}{UR['shnv']}+{cyc(y)}{UR['ibnv']}+{cyc(y)}{UR['unnv']}", "0.0%", "應 ≤ 100%；其餘為三星電機等"),
    "shot": ("南電份額：非 NVIDIA（金額）", "%", lambda y: f"=PL!{FC[y]}{A_['shot']}", "0.0%", "Assumed"),
    "ibot": ("Ibiden 份額：非 NVIDIA（Ibiden 權重 × 份額相對值 ÷ v_ot）", "%", lambda y: f"=ib_w_ot*{ibrel(y)}/v_ot", "0.0%", "與欣興模型同算法"),
    "unot": ("欣興份額：非 NVIDIA（金額）", "%", lambda y: f"=un_ot_{y % 100}", "0.0%", "欣興模型"),
    "sumot": ("合計：Ibiden＋欣興＋南電（非 NVIDIA）", "%", lambda y: f"={cyc(y)}{UR['shot']}+{cyc(y)}{UR['ibot']}+{cyc(y)}{UR['unot']}", "0.0%", "應 ≤ 100%；其餘為三星電機、景碩、AT&S"),
    "pu": ("每顆 NVIDIA 加速器的南電營收", "US$", lambda y: f"=PL!{FC[y]}{A_['v']}*PL!{FC[y]}{A_['shnv']}*PL!{FC[y]}{A_['cy']}", "#,##0.0", "Ibiden、欣興對照值見各自 dist Units 頁"),
    "accu": ("南電加速器載板營收", "US$M", lambda y: f"=PL!{FC[y]}{A_['accu']}", "#,##0", ""),
    "aisub": ("南電 AI 載板營收（含交換器等其他 AI）", "NT$B", lambda y: f"=PL!{FC[y]}{A_['aisub']}", "#,##0.0", ""),
    "itr": ("南電 AI 載板營收 ÷ AI IT 資本支出", "%", lambda y: f"=PL!{FC[y]}{A_['aisub']}/fx/itc_{y % 100}", "0.000%", "× 每 GW IT capex 即每 GW 營收"),
    "pgw": ("每 GW 南電 AI 載板營收（VR200 每 GW IT capex 37.59B）", "US$M", lambda y: f"=PL!{FC[y]}{A_['aisub']}/fx/itc_{y % 100}*37.59*1000", "#,##0", "Tokenomics IF_CapexIT（VR200 基準，經橋接表；與 Ibiden、欣興模型相同係數）")}
for key in UKEYS:
    lab, unit, f, fmt, note = UDEF[key]; ur = UR[key]
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note)
UN_NOTE = 4 + len(UKEYS) + 1
UN.cell(UN_NOTE, 1, "讀法：南電在 GPU、ASIC 加速器載板的份額小（個位數），AI 載板以網通交換器（800G／1.6T 交換晶片）為主；加速器載板只占南電 AI 載板一小部分，其餘為校準殘差（公司未揭露）。三家合計留給三星電機、景碩、AT&S。").font = BOLD
UN.column_dimensions["A"].width = 60; UN.column_dimensions["B"].width = 10
for c in "CDEF": UN.column_dimensions[c].width = 13
UN.column_dimensions["G"].width = 66

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，年底落後 P/E）只作參考"
V["A2"] = "循環位置：南電不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照 Ibiden、欣興；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
O = 51
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
V.cell(O + 10, 1, "前瞻 P/E（stockanalysis，價格日期 09-07）"); V[f"B{O + 10}"] = "=fpe"; V[f"B{O + 10}"].number_format = "0.0"
V[f"E{O + 3}"] = "FY2026 三情境相同（錨點）"; V[f"E{O + 4}"] = "FY2027 情境差只來自 BT 年增（顆數係數從 CY2028 開始）"
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
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數愈高，現價愈依賴「ABF、BT 漲價延續、交換器載板份額守住」。本版未做先驗，無 Andy 合理 P/E 對照。")
WMT_NOTE = vr
# 十年 P/E：stockanalysis 年度 P/E（FY2021–25）；自算＝digrin 12 月未調整股價 ÷ 年度 EPS（wantgoo）；2016–2018 EPS 為負，不計
PE_SA = [None, None, None, None, None, 34.93, 7.56, 27.94, 413.91, 79.99]
PX_DEC = [24.65, 26.10, 33.80, 45.95, 182.0, 572.0, 227.0, 251.5, 130.5, 241.0]
EPS_H = [-1.07, -3.03, -0.90, 0.48, 5.67, 16.38, 30.05, 9.00, 0.32, 3.01]
PE_SELF = [None if e_ <= 0 else f"={p_}/{e_}" for p_, e_ in zip(PX_DEC, EPS_H)]
T10 = vr + 2
V.cell(T10, 1, "5. 十年 P/E（2016–2025 年底；兩來源平均，分位排名用；2016–2018 年 EPS 為負，P/E 無意義，不計）").font = BOLD
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
    {"label": "自算（digrin 12 月未調整股價 ÷ 年度 EPS）", "vals": PE_SELF, "tag": "Derived（價格為月資料的 12 月值；EPS 取自 wantgoo 歷年股利表；2016–2018 EPS 為負不計，2019–2025 共 7 年）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位，兩個來源分別取分位再平均。南電的落後 P/E 與盈餘反向：盈餘低谷年（2019、2024）P/E 96–414x，高峰年（2022）約 8x；2016–2018 年虧損不計，十年多為 AI 載板需求前的循環，套用在 FY2027 EPS 時偏誤方向不明確。")
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
V["E14"] = "方法：南電模型有資本支出與折舊，改用每股（淨利＋折舊－資本支出）（PL 頁「每股股東現金流」；FY2030 起資本支出率收斂到穩態更新）"
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
     "近三年 ROE 約 5%（2024 年接近 0）→ 再投資率約 74%；不進目標價")
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = [("欣興（Unimicron）", "3037.TW", 42.28, "2026-10-22"), ("景碩（Kinsus）", "3189.TW", 44.70, "2026-10-22"), ("Ibiden（揖斐電）", "4062.T", 53.97, "2026-10-16"),
         ("三星電機（Samsung Electro-Mechanics）", "009150.KS", 43.40, "2026-08-28"), ("AT&S", "ATS.VI", 14.91, "2026-10-23")]
for i, (nm, tk, pe, dt) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左；與欣興、Ibiden 模型同一組數字）"); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", "ABF 載板同業：欣興、景碩、Ibiden、三星電機、AT&S（無中國 A 股）；三星電機價格日期較舊")
tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）", [f"=tp_wA*(2/12*B{EPSR[2027]}+10/12*B{EPSR[2028]})+tp_wB*(2/12*C{EPSR[2027]}+10/12*C{EPSR[2028]})+tp_wC*(2/12*D{EPSR[2027]}+10/12*D{EPSR[2028]})", None, None], "0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], "0.00", "同業中位數 × 前瞻 EPS")
V.cell(37, 1, "0.3 結論").font = BOLD
hdr(V, 38, ["項目", "數值", "", "", "說明"])
tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B24+(1-tp_w_dcf)*B35", None, None], "0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁（0.5）")
tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], "0.00")
tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], "0.00", "stockanalysis：14 位分析師平均 1,651.31（頁面價格日期 2026-09-07）；10 月美銀 2,000、大和 2,460")
tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-22 收盤）")
tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝頁面價格 ÷ 前瞻 P/E（In 頁 cons_feps）")
tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
tset(46, "band", "價格帶（DCF 不如預期～順利）", ['=TEXT(D23,"#,##0")&"～"&TEXT(B23,"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=IF(px/{c}{EPSR[2028]}>=MAX({PE10}),1,IF(px/{c}{EPSR[2028]}<=MIN({PE10}),0,PERCENTRANK({PE10},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2019–2025 年底 P/E（第 5 節兩來源平均，2016–2018 虧損不計）的分位（超出區間時記 0% 或 100%）；只作檢查，不進目標價")
V.cell(48, 1, "目標價方法")
V["B48"] = ("DCF：股東現金流＝每股（淨利＋折舊－資本支出），r＝10%、g＝4%，FY2027–32 明確期＋永續，情境價值 < 0 取 0；南電資本支出低（2026 年約 40 億）、折舊約 70 億，"
            "現金流高於淨利；新廠資本支出未定，若擴產加速，現金流偏高估；ABF 溢價沿用 Ibiden 模型（2027 持平、之後回落），比券商保守。"
            "Comparable：欣興、景碩、Ibiden、三星電機、AT&S 前瞻 P/E 中位數（無 A 股）；同業同處 ABF 短缺溢價期，倍數偏高。"
            "情境對應：A＝順利（顆數基準）、B＝平平（顆數 ×0.8／0.85）、C＝不如預期（顆數 ×0.6），皆讀 tsmc 同字母係數；BT 年增 A 最高、C 最低")
for r_ in (24, 35, 39):
    V[f"B{r_}"].fill = KEY
V.column_dimensions["A"].width = 60
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；FY2026 恆等於錨點，其他 AI 載板在每個案例內重新校準）：對 FY2027、FY2028 營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("AI 部署 ×0.8", "FY2027 起顆數與其他 AI ×0.8", dict(mf="0.8")), ("AI 部署 ×1.2", "×1.2", dict(mf="1.2")),
         ("ABF 產能 低", "產能指數 1.10／1.20／1.35", dict(cap="cap_{y}_lo")), ("ABF 產能 高", "1.30／1.50／1.85", dict(cap="cap_{y}_hi")),
         ("ABF 溢價 低（Ibiden 低值）", "0.95／0.85／0.75", dict(pp="ibpp_{y}_lo")), ("ABF 溢價 高（Ibiden 高值）", "1.05／1.05／1.00", dict(pp="ibpp_{y}_hi")),
         ("漲價年度化 8%", "基準 15%", dict(carry="carry_lo")), ("漲價年度化 22%", "", dict(carry="carry_hi")),
         ("ABF 現金毛利率 40%", "基準 50%", dict(ftabf="ft_abf_lo")), ("ABF 現金毛利率 60%", "", dict(ftabf="ft_abf_hi")),
         ("其他產品現金毛利率 25%", "基準 40%", dict(ftoth="ft_oth_lo")), ("其他產品現金毛利率 55%", "", dict(ftoth="ft_oth_hi")),
         ("BT 年增 走 C 路徑", "15%／0%／0%", dict(gbt="g_bt{y}_C")),
         ("折舊年限 6 年", "基準 8 年", dict(life="life_lo")), ("折舊年限 10 年", "", dict(life="life_hi")),
         ("資本支出 低", "每年 60 億", dict(cx="cx{y}_lo")), ("資本支出 高", "200／240／200 億", dict(cx="cx{y}_hi")),
         ("南電 NVIDIA 份額 低", "1%", dict(shnv="sh_nv_{y}_lo")), ("南電 NVIDIA 份額 高", "3%→6%", dict(shnv="sh_nv_{y}_hi")),
         ("南電 ASIC 份額 低", "3%", dict(shot="sh_ot_{y}_lo")), ("南電 ASIC 份額 高", "8%→10%", dict(shot="sh_ot_{y}_hi")),
         ("交換器載板彈性 0.3", "基準 0.5", dict(elo="el_o_lo")), ("交換器載板彈性 0.8", "", dict(elo="el_o_hi")),
         ("營業費用彈性 1.0", "基準 0.6", dict(phi="phi_hi")), ("稅率 24%", "基準 21%", dict(tax="tax_hi")), ("業外 0", "每年 0（基準 6 億）", dict(nonop="nonop{y}_lo"))]
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.0"), ("D", "gm", 2027, "0.00%"), ("E", "eps", 2027, "0.00"),
                           ("G", "rev", 2028, "#,##0.0"), ("H", "gm", 2028, "0.00%"), ("I", "eps", 2028, "0.00")):
        SE[f"{c}{i}"] = f"={pre}{FC[y]}{R[key]}"; SE[f"{c}{i}"].number_format = fmt
    SE[f"F{i}"] = f"=E{i}/E$4-1"; SE[f"J{i}"] = f"=I{i}/I$4-1"
    for c in "FJ": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
base_row = 50; SB = {}
for nm, setting, ov in cases:
    SB[nm] = engine(SE, base_row, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    base_row = SB[nm]["end"] + 2
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm], "")
nn = 5 + len(cases)
put_case(nn, "情境 B", "tsmc 顆數係數 B；BT 年增較低", PB["B"], "PL!")
put_case(nn + 1, "情境 C", "tsmc 顆數係數 C；BT 年增最低", PB["C"], "PL!")
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：漲價年度化與 ABF 溢價直接決定毛利（溢價全數落在毛利）；ABF 產能受限時，份額與部署的上修只部分轉成營收；南電加速器份額小，NVIDIA／ASIC 份額的影響有限；FY2027 情境差只來自 BT 年增。")
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
assert TPS_LAST < 50, TPS_LAST
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 40
for c in "CDEFGHIJ": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 券商、共識與市場口徑。只作事後對照，不為對上數字而調參數。FactSet 共識為 2026-06（2Q26 財報前），10 月券商數字只取得美銀一家完整年度預估"
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
rec("FY2026 EPS vs 美銀 18.44（2026-10-02）", 18.44, pl("eps", 2026), "TechNews（美銀）", "FactSet 2026-06 平均 13.19 為 2Q26 財報之前")
rec("FY2026 EPS vs 凱基 17.27（2026-07-20）", 17.27, pl("eps", 2026), "凱基", "")
rec("FY2027 EPS vs 券商平均（美銀 43.80、凱基 38.88）", "=bk_eps27", pl("eps", 2027), "TechNews（美銀）、凱基", "差距超過 30% 的原因見 README")
rec("FY2027 EPS vs FactSet 平均 32.12（2026-06）", "=cons_eps27", pl("eps", 2027), "鉅亨（FactSet 14 位）", "")
rec("FY2028 EPS vs FactSet 平均 59.21（2026-06）", "=cons_eps28", pl("eps", 2028), "鉅亨（FactSet）", "")
rec("FY2028 EPS vs 券商平均（美銀 70.09、凱基 61.35）", "=bk_eps28", pl("eps", 2028), "TechNews（美銀）、凱基", "")
rec("FY2027 營收 vs FactSet 平均（2026-06）", "=cons_rev27", pl("rev", 2027), "鉅亨（FactSet）", "美銀 90.33", fmt="#,##0.0")
rec("FY2028 營收 vs FactSet 平均（2026-06）", "=cons_rev28", pl("rev", 2028), "鉅亨（FactSet）", "美銀 120.65", fmt="#,##0.0")
rec("模型前瞻 12 月 EPS（0.2 × FY2026＋0.8 × FY2027）vs 頁面價格 ÷ 前瞻 P/E", "=cons_feps", lambda s: f"=0.2*PL!D{PB[s]['eps']}+0.8*PL!E{PB[s]['eps']}", "stockanalysis（推導）", "口徑不同，只作量級比較")
RC.cell(rr_, 1, "B. 錨點與毛利率").font = BOLD; rr_ += 1
rec("3Q26 毛利率（假設）vs 法人估 28.1%（百分點）", 0.281, lambda s: "=Q!E6", "三立 2026-07-27（法人）", "3Q26 財報約 10 月底", fmt="0.0%", pt=True)
rec("3Q26 EPS（模型）vs 大和估 5.62", 5.62, lambda s: "=Q!E14", "TechNews（大和）", "", fmt="0.00")
rec("4Q26 毛利率（假設）vs 美銀估 32.5%（百分點）", 0.325, lambda s: "=Q!F6", "TechNews（美銀）：32–33%", "", fmt="0.0%", pt=True)
rec("FY2027 毛利率 vs 美銀 39.9%（百分點）", 0.399, pl("gm", 2027), "TechNews（美銀）", "凱基 37.4%", fmt="0.0%", pt=True)
rec("FY2028 毛利率 vs 美銀 46.6%（百分點）", 0.466, pl("gm", 2028), "TechNews（美銀）", "凱基 42.1%", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "C. 產能、份額與 AI 口徑").font = BOLD; rr_ += 1
rec("FY2027 ABF 需求滿足率（< 100%＝產能受限）", None, pl("fill", 2027), "PL 頁", "凱基：ABF 供需缺口 2027 年 22%、2028 年 29%", fmt="0.0%")
rec("FY2028 ABF 需求滿足率", None, pl("fill", 2028), "PL 頁", "", fmt="0.0%")
rec("FY2026 AI 載板占營收", None, pl("aish", 2026), "PL 頁", "公司應用別：網通 49%、AI 與 HPC 16%（1Q26）", fmt="0.0%")
rec("FY2027 份額合計：Ibiden＋欣興＋南電（NVIDIA）", None, lambda s: f"=Units!D{UR['sumnv']}", "Units 頁", "應 ≤ 100%", fmt="0.0%")
rec("FY2027 份額合計：Ibiden＋欣興＋南電（非 NVIDIA）", None, lambda s: f"=Units!D{UR['sumot']}", "Units 頁", "應 ≤ 100%", fmt="0.0%")
rec("FY2027 每顆 NVIDIA 加速器的南電營收（US$）", None, lambda s: f"=Units!D{UR['pu']}", "Units 頁", "Ibiden 同期約 169（ibiden/dist Units 頁）", fmt="#,##0.0")
rec("FY2027 每 GW 南電 AI 載板營收（US$M）", None, lambda s: f"=Units!D{UR['pgw']}", "Units 頁", "Ibiden 同期約 62（只含加速器載板）", fmt="#,##0")
RC.cell(rr_, 1, "D. 資本支出").font = BOLD; rr_ += 1
rec("FY2027＋FY2028 資本支出 vs 凱基 2025–28 合計約 416 億扣除 2025–26（約 64 億）", 35.2, lambda s: f"=PL!E{PB[s]['cx']}+PL!F{PB[s]['cx']}", "凱基（US$12.8 億 × 32.5）", "", fmt="#,##0.0")
RC.column_dimensions["A"].width = 72
for c in "BCDEFGH": RC.column_dimensions[c].width = 11
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60
RECON_LAST = rr_ - 1

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司財報經資料商與媒體轉載；法說內容經媒體摘要；券商與資料商皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("資料商：stockanalysis（現價、股數、前瞻與近 12 月 P/E、目標價、FY2021–25 與季損益、現金流、ROE、年度 P/E）", "https://stockanalysis.com/quote/tpe/8046/"),
        ("資料商：stockanalysis 歷史價格（2026-10-22 收盤 1,265）", "https://stockanalysis.com/quote/tpe/8046/history/"),
        ("資料商：財報狗（10-22 收盤、近 12 月 P/E、近 4 季 ROE）", "https://statementdog.com/analysis/8046"),
        ("資料商：HiStock（現價核對、2Q26 EPS）", "https://histock.tw/stock/8046"),
        ("公司（摘要）：富果 1Q26 法說摘要（應用別、資本支出、擴產）", "https://blog.fugle.tw/post/earnings-call-8046-2026-06-12"),
        ("公司（摘要）：vocus 4Q25 法說摘要（FY2025、產品別 2025）", "https://vocus.cc/article/6a12e069fd897800013a5f84"),
        ("媒體：TechNews 2026-10-03（9 月營收、前 9 月累計）", "https://finance.technews.tw/2026/10/03/n-p-c-8046-202609-financial-report/"),
        ("媒體：winvest（7 月營收、1Q26 EPS）", "https://winvest.tw/News/Detail/90850"),
        ("券商（媒體）：三立 2026-07-27（2Q26 營收、法人 3Q26 預估、產能）", "https://www.mirrormedia.mg/external/setn_1879204"),
        ("券商（媒體）：TechNews 2026-10-02 美銀（目標價 2,000、EPS 2026–28、毛利率、BT 漲價、AMD／博通）", "https://finance.technews.tw/2026/10/02/bank-of-america-gives-nanyang-power-a-buy-rating-with-a-target-price-of-nt2000/"),
        ("券商（媒體）：TechNews 2026-10-05 大和（目標價 2,460、8 月自結獲利、3Q／4Q26 預估）", "https://finance.technews.tw/2026/10/05/daiwa-securities-strongly-supports-nan-ya-powers-target-price-of-nt2460/"),
        ("券商：凱基 PCB 產業報告 2026-07-20（EPS、營收、毛利率、產品組合、產能、資本支出、ABF 缺口、同業 P/E）", "https://www.kgi.com.hk/zh-cn/-/media/files/kgishk/research-reports/tw-reports/pcb-sector_20260720e.pdf"),
        ("研究文章：brodoginvestment 南電深度分析（引述券商報告；交換器份額、折舊、單價）", "https://brodoginvestment.substack.com/p/8046-tt"),
        ("共識：鉅亨 2026-06-24 FactSet（14 位分析師 EPS、營收）", "https://news.cnyes.com/news/id/6509588"),
        ("資料商：digrin 月股價（12 月未調整值）", "https://www.digrin.com/stocks/detail/8046.TW/price"),
        ("資料商：wantgoo 歷年股利表（EPS 2014–2025）", "https://www.wantgoo.com/stock/8046/profitability/dividend-payout-ratio"),
        ("資料商：stockanalysis 同業前瞻 P/E（欣興、景碩、Ibiden、三星電機、AT&S；與欣興、Ibiden 模型同一組）", "https://stockanalysis.com/quote/tpe/3189/"),
        ("本 repo 上游成品（唯讀）：ibiden、unimicron、markets/network、tsmc；情境係數與 markets/memory/market.py SCEN_UNITS 核對", f"{rel(IB_F)}；{rel(UN_F)}；{rel(NET_F)}；{rel(TSMC_F)}"),
        ("待補（人工）：3Q26 財報與法說（約 10 月底）；公司原文（公開資訊觀測站財報、法說簡報）；FY2016–20 營收（十年平均淨利率）；ABF／BT 季度拆分；加速器與交換器載板拆分；份額量化；折舊會計年限；新廠資本支出；10 月後共識（FactSet 更新）；companiesmarketcap 年底 P/E（擷取失敗）", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 110; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("南電（8046.TW）收支與估值模型 v0.1（AI 半導體 Project：載板與 PCB 次產業）　2026-10-25", True),
 ("版本：v0.1（初版）。樣板：unimicron/。AI 加速器載板的每顆金額、Ibiden 份額與非 NVIDIA 權重讀 ibiden/dist，欣興份額讀 unimicron/dist（本專案權威）；三家份額合計在 Units 頁檢查。", False),
 ("模型命題：南電的 ABF 載板營收＝MIN（需求, 產能）；需求以網通交換器載板為主（隨 markets/network 的交換晶片與交換器市場），加上 AI 加速器顆數 × 每顆載板金額（與 Ibiden、欣興同口徑）× 南電份額（個位數）與非 AI ABF；產能 2026–28 年靠去瓶頸與昆山廠爬坡，新廠約 2029 年。BT 載板（約三成營收）以漲價為主。資本支出低（2026 年約 40 億）、折舊約 70 億。", False),
 ("核心問題：2026 年下半年的漲價與交換器載板需求，能不能讓毛利率從 FY2026 約 27% 再升、EPS 在 FY2027–28 倍增；約 1,265 元的股價隱含什麼路徑。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（延伸到 FY2032；PL 頁灰底）。金額：十億新台幣；EPS：新台幣（基本、歸屬母公司）。3Q26 財報約 10 月底公布，本版只納入 3Q26 營收實績。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("加速器載板：markets/network 的 NVIDIA 與非 NVIDIA 顆數 × tsmc 同字母情境係數 × 每顆 NVIDIA 載板金額（Ibiden 模型 US$215 × 面積層數 +18%／年 × Ibiden 溢價指數）× 南電份額（NVIDIA 2%→4%、非 NVIDIA 金額 5%→7%，Assumed）× 漲價年度化 1.15。", False),
 ("其他 AI 載板（交換器為主）＝FY2026 AI 載板（ABF × 65%）− 加速器載板（校準殘差），之後 × 交換晶片與交換器市場成長 ^ 0.5 × 情境係數 × 溢價 × 漲價年度化；非 AI ABF +3%。ABF 產能指數 1.20／1.35／1.60（Assumed）× 組合 +5%／年。BT 年增依情境（A 30%／10%／5%、B 25%／5%／3%、C 15%／0%／0%），PCB 持平。", False),
 ("毛利＝FY2026 錨點＋ABF 現金毛利增量（50%）＋其他產品現金毛利增量（40%）＋漲價溢價營收 − 折舊增量；折舊＝前一年 ×（1 − 1／8 年）＋前一年資本支出 ÷ 8 年；資本支出 140／180／120 億（三情境相同）。營業費用彈性 0.6；業外每年 6 億；稅率 21%；股數持平 6.46 億股。", False),
 ("校準：FY2026 ＝ 1Q26、2Q26 財報實績（EPS 2.03、3.50）＋3Q26 營收實績 163.8 億＋3Q26 毛利率 30%、4Q26 營收季增 10%、毛利率 32.5%（Assumed）。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（股東現金流＝每股（淨利＋折舊－資本支出），FY2027–32 明確期＋永續，r 10%、g 4%，三情境權重 0.3／0.5／0.2，負值取 0）＋0.5 × Comparable 目標價（欣興、景碩、Ibiden、三星電機、AT&S 前瞻 P/E 中位數 × 12 個月後前瞻 EPS；無 A 股）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解。十年 P/E 規則值 × FY2027 EPS 保留為參考（不是目標價）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("情境與行業一致：A／B／C 讀 tsmc 同字母顆數係數（＝markets/memory SCEN_UNITS，建置時核對）；交換器市場讀 markets/network（單一路徑 × 同字母係數）；公司特有變數只有 BT 年增隨情境，A 最有利、C 最不利。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司數字經資料商與媒體轉載（公司原文待人工複核）；券商與資料商為 Interested-party。", False),
 ("", False),
 ("已知限制：(1) 加速器、交換器與 CPU 載板的拆分公司未揭露，其他 AI 載板為校準殘差；(2) 份額、產能指數、漲價年度化皆為 Assumed，與 Ibiden、欣興合計只做上限檢查；(3) 毛利用現金流通率近似，未建分廠稼動率；(4) AMD 伺服器 CPU 認證（美銀：成功則 2027 年逾兩成營收）未建入；(5) 新廠資本支出未定，DCF 現金流可能偏高；(6) 十年 P/E 2016–2018 年虧損不計，其餘與盈餘反向；(7) FactSet 共識為 2026-06；(8) 3Q26 財報（約 10 月底）後更新錨點。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ---------------- v0.1.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00／14:17）----------------
# 新列一律放在 Val、README 工作表既有內容之後，不位移既有列；既有數值不變
VER_PREV, VER_NOW = 'v0.1', 'v0.1.1'
OBS = [
 "目標價 NT$907（DCF 466、Comparable 1,347 各半），比現價 1,265（2026-10-22）低 28.3%，級距 −25%～−50%；Comparable 單獨為 +6.5%。主因是 DCF 不是 EPS：本模型前瞻 EPS 31.04 與共識前瞻 EPS 31.13 只差 −0.3%，同業前瞻 P/E 中位數 43.4 倍、現價 ÷ 共識前瞻 EPS 40.6 倍；DCF 用每股（淨利＋折舊 − 資本支出），FY2028 情境 B 每股現金流 14.55 只有 EPS 30.61 的 48%，DCF 比現價低 63.1%（改用 EPS ×（1 − g／ROE）只有 137，因 ROE 5.4% 遠低於折現率 10%）。",
 "ABF 漲價路徑最敏感：漲價年度化 8%（基準 15%）時 FY2027 EPS −9.7%；ABF 溢價取 Ibiden 低值時 FY2028 EPS −17.4%、高值 +17.8%；ABF 產能低值 −8.1%（FY2027）。",
 "FY2028 EPS 三情境 35.90／30.61／23.33，比 2026-06 FactSet 共識 FY2028 EPS 59.21 低 39%～61%（共識在 2Q26 財報前）。下一個檢驗點是 3Q26 財報與法說（約 10 月底，stockanalysis 估 2026-11-03）。"
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
if os.environ.get("NANYA_PCB_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE, "UN_NOTE": UN_NOTE, "T": T, "T10": T10,
               "TPS0": TPS0, "TPS_LAST": TPS_LAST, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0},
              open(os.environ["NANYA_PCB_ROWMAP"], "w"), ensure_ascii=False, indent=1)

# 欣興電子（3037.TW）收支與估值模型 v0.1.1（AI 半導體 Project：載板與 PCB 次產業；樣板 ibiden/）
# 用法：python3 unimicron/build_unimicron.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；EPS：新台幣（基本，歸屬母公司）。
# AI 加速器載板的份額口徑與 Ibiden 一致：每顆 NVIDIA 載板金額、Ibiden 份額、非 NVIDIA 權重唯讀取自 ibiden/dist（本專案權威，規則第 13 條）。
# 加速器顆數、AI IT 資本支出唯讀取自 markets/network/dist；情境顆數係數取自 tsmc/dist；不修改上游。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP
import valuation_rule as VR
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
rel = lambda f: os.path.relpath(f, ROOT)
UNITS = {}
for r in openpyxl.load_workbook(NET_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and (str(k).startswith("u_") or str(k).startswith("g29_")):
        UNITS[k] = (r[4].value, r[1].value, f"{rel(NET_F)}｜In!E{r[0].row}")
ITC = None
for r in openpyxl.load_workbook(NET_F, data_only=True)["Mkt"].iter_rows(min_row=4):
    if r[0].value == "AI IT 資本支出（加速器 × 每顆 IT capex）":
        ITC = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
assert ITC and all(v for v in ITC[0]), ITC
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and str(k).startswith("um"):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
assert all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()
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

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B）；顆數：百萬顆；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間，不當點估計（放情境或 Sens）。SemiAnalysis 未使用。公司原文＝公開資訊觀測站財報與法說會（經工商時報、Quartr 轉載）；公司說法與指引為 Interested-party；券商、媒體、資料商為 Interested-party"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
CT2Q = "工商時報 2026-07-29（2Q26 法說：產品組合、稼動率、AI 占比、資本支出）"
CT2QF = "工商時報 2026-07-29（2Q26 財報：營收、營業利益、業外、淨利）"
QR2Q = "Quartr 2Q26 摘要（公司財報轉載）"
FG4Q = "富果 4Q25 法說摘要（2026-02-25，AI 摘要）"
SINO1Q = "永豐 豐學 1Q26（2026-04-28）"
UA_GDR = "優分析 2026-04-10（GDR 董事會決議）"
UA_GDR2 = "優分析 2026-07-06（GDR 5,000 萬單位、詢價）"
TN9 = "TechNews 2026-10-08（9 月營收、前 9 月累計）"
SETN = "三立 2026-10-19（本土法人；Q3 營收、Q4 預估、產能時程）"
UDN10 = "經濟日報 2026-10-12（美系券商 EPS 預估）"
CNYES_FS = "鉅亨 2026-06-24（FactSet 22 位分析師）"
SA = "stockanalysis.com（3037）"

sec("A. AI 加速器顆數與 AI IT 資本支出（讀 markets/network、tsmc 成品；唯讀，不自行重估）")
for g in NV + OT:
    for y in (26, 27, 28):
        v, lab, src = UNITS[f"u_{g}_{y}"]
        inp(f"u_{g}_{y}", lab, "百萬顆", None, v, None, "Derived", src, "唯讀上游；與 Ibiden、台積電模型同一份顆數")
    v, lab, src = UNITS[f"g29_{g}"]
    inp(f"g29_{g}", lab, "%", None, v, None, "Derived", src, "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}", "US$B", None, ITC[0][j], None, "Derived", ITC[1], "其他 AI 載板（交換器、CPU、網通）的成長依據；Recon 每 GW 換算" if y == 26 else "同上")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝順利（基準）；B＝平平（AI 資本支出消化）；C＝不如預期（明顯下修）；與台積電、Ibiden 模型相同" if s == "A" else "同上")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")

sec("B. 每顆加速器載板金額與份額（口徑與 Ibiden 一致；Ibiden 數字唯讀取自 ibiden/dist）")
inp("v26", "FY2026 每顆 NVIDIA 加速器載板金額（全市場；Ibiden 模型校準）", "US$", None, round(IB_V26[0], 6), None, "Derived", IB_V26[1], "Ibiden 模型：Ibiden 每顆營收 ÷ 份額 70%（BofA：Ibiden 占 GPU 載板 70–80%，Interested-party）")
inp("g_asp", "每顆載板金額年增（面積、層數；Ibiden 模型）", "%", IB["g_asp"][0], IB["g_asp"][1], IB["g_asp"][2], "Analogy", IB["g_asp"][3], "Ibiden 路線圖 2026 年 90×90 mm → 2028 年 110×110 mm（Tom's Hardware，經 Ibiden 模型）")
for y in (27, 28, 29):
    k = f"pp_{y}"
    inp(f"ibpp_{y}", f"ABF 價格溢價指數 20{y}（FY2026＝1；Ibiden 模型）", "x", IB[k][0], IB[k][1], IB[k][2], "Assumed", IB[k][3], "全市場短缺溢價；新產能到位後回落（Ibiden 模型）" if y == 27 else "同上")
    k = f"rel_{y}"
    inp(f"ibrel_{y}", f"Ibiden NVIDIA 份額 20{y} ÷ 2026（Ibiden 模型）", "x", None, IB[k][1], None, "Assumed", IB[k][3], "Units 頁份額合計檢查用")
inp("ib_sh26", "Ibiden 占 NVIDIA 加速器載板 FY2026（Ibiden 模型假設）", "%", *SIL.vals("abf_ibsh26"), "Interested-party", IB_V26[1] + "（列標籤：假設 FY2026 份額 70%）", "BofA（經 thebuildout 2026-03-26）：Ibiden 70–80%")
inp("ib_w_ot", "Ibiden 非 NVIDIA 權重（份額 × 每顆金額，相對 NVIDIA；Ibiden 模型）", "x", IB["w_ot"][0], IB["w_ot"][1], IB["w_ot"][2], "Assumed", IB["w_ot"][3], "Units 頁：÷ v_ot 即 Ibiden 非 NVIDIA 份額")
inp("v_ot", "非 NVIDIA 加速器每顆載板金額 ÷ NVIDIA", "x", *SIL.vals("abf_vot"), "Assumed", "TPU、Trainium、MI 系列封裝面積接近 GPU，但層數與 CoWoS 規格不同；量化：未找到", "Ibiden 權重 0.4 ÷ 0.8 ＝ Ibiden 非 NVIDIA 份額 50%（Units 頁合計檢查）")
for y in (26, 27, 28, 29):
    inp(f"sh_nv_{y}", *SIL.row9(f"abf_un_nv_{y}")[1:])
for y in (26, 27, 28, 29):
    inp(f"sh_ot_{y}", *SIL.row9(f"abf_un_ot_{y}")[1:])

sec("C. FY2026 產品組合與 AI 占比（公司說法；Q2 組合推全年）")
inp("mx_abf", "FY2026 ABF 載板 ÷ 營收", "%", 0.50, 0.53, 0.56, "Interested-party", CT2Q + "：2Q26 ABF 52%（季增 22%）→ 1Q26 約 49%；下半年 AI 占比升", FG4Q + "：4Q25 載板（ABF＋BT）59%")
inp("mx_bt", "FY2026 BT 載板 ÷ 營收", "%", 0.08, 0.09, 0.10, "Interested-party", CT2Q + "：2Q26 BT 9%", "—")
inp("mx_pcb", "FY2026 HDI＋PCB ÷ 營收", "%", 0.29, 0.31, 0.33, "Interested-party", CT2Q + "：2Q26 HDI 22%、PCB 9%", FG4Q + "：4Q25 HDI 27%、PCB 11%")
inp("ai_abf", "FY2026 AI 載板 ÷ ABF 營收", "%", 0.65, 0.75, 0.85, "Assumed", CT2Q + "：載板 AI 占比上半年逾 60%、下半年 70%（ABF＋BT 口徑）", "BT 多為記憶體與手機，AI 集中在 ABF：0.66 × 62% ÷ 53% ≈ 0.77（Derived）")
inp("ai_pcb", "FY2026 HDI＋PCB 的 AI 占比", "%", 0.55, 0.65, 0.75, "Interested-party", CT2Q + "：PCB AI 占比上半年逾 60%、下半年 65–70%", FG4Q + "：2025 年 55–60%")

sec("D. 需求與產能（FY2027 起）")
inp("el_o", "其他 AI 載板（交換器、伺服器 CPU、網通）對 AI IT 資本支出的彈性", "x", 0.3, 0.5, 0.8, "Assumed", "其他 AI 載板＝FY2026 AI 載板 − 加速器載板（校準殘差）；隨部署規模成長但低於等比例", "量化：未找到")
inp("g_nai", "非 AI ABF（PC、一般伺服器）年增", "%", -0.02, 0.03, 0.06, "Assumed", "PC、一般伺服器緩步成長", "—")
for y in (27, 28, 29):
    inp(f"cap_{y}", *SIL.row9(f"abfcap_un_{y}")[1:])
inp("g_mix", "每單位產能營收年增（層數、尺寸組合）", "%", 0.0, 0.05, 0.08, "Assumed", "高層數載板每片金額高，但也占用較多製程時間；淨效果 Assumed", "—")
inp("carry", "欣興漲價年度化（CY2026 只反映約半年漲價；FY2027 起 ×(1＋此值)）", "%", 0.0, 0.05, 0.10, "Assumed",
    "Ibiden 溢價指數以 FY2026（2026-04～2027-03）為基期，已含 2026 下半年漲價；欣興 CY2026 基期含上半年較低價格", "媒體與法人：2026 年 ABF 漲價約 15–20%，外資估最高 30%（口袋學堂，非公司指引）")
inp("g_bt", "BT 載板年增", "%", 0.0, 0.05, 0.10, "Assumed", CT2Q + "：BT 稼動率 80–85%；2Q26 年減 4%", "記憶體循環上行（Assumed）")
for y, v in ((27, 0.20), (28, 0.15), (29, 0.10)):
    inp(f"g_pai_{y}", f"AI HDI／PCB 年增 20{y}（顆數情境係數另乘）", "%", v - 0.10, v, v + 0.10, "Assumed",
        CT2Q + "：HDI 2Q26 年增 29%；HDI／PCB 稼動率 85–90%" if y == 27 else "同上", "AI 伺服器加速器板、交換器板（Assumed）" if y == 27 else "—")
inp("g_pcbn", "非 AI HDI／PCB 年增（手機、消費、車用）", "%", -0.03, 0.02, 0.05, "Assumed", CT2Q + "：IoT 年減 3%、車用年減 26%（2Q26 應用別）", "—")
inp("g_fpc", "軟板與其他年增", "%", -0.05, 0.0, 0.03, "Assumed", CT2Q + "：軟板 2Q26 年減 14%", "—")

sec("E. 毛利、折舊、費用與淨利（FY2027 起）：擴產折舊是毛利率關鍵")
inp("ft_abf", "ABF 營收增量的現金毛利率（扣除漲價溢價後）", "%", 0.40, 0.50, 0.60, "Assumed", "FY2025→FY2026 全公司現金毛利增量 ÷ 營收增量見 Q 頁（含漲價，偏高）", "法人：ABF 毛利率可能接近 40%（三立，Interested-party）；加回折舊約 50%")
inp("ft_oth", "其他產品（BT、HDI、PCB、軟板）營收增量的現金毛利率", "%", 0.15, 0.25, 0.35, "Assumed", "FY2025 全公司毛利率 13.9%、折舊 ÷ 營收 14%（現金毛利率約 28%）", "—")
for s, vs in (("A", (55, 50, 45)), ("B", (48, 42, 38)), ("C", (40, 30, 28))):
    for y, v in zip((27, 28, 29), vs):
        inp(f"cx{y}_{s}", f"資本支出 FY20{y}（情境 {s}）", "NT$B", None, v, None, "Assumed",
            UDN10 + "：美系券商估 2027、2028 年資本支出均逾 400 億（Interested-party）" if (s == "A" and y == 27) else "同上",
            "2026 年預算 537 億（第三次上修；公司）；光復二廠設備 300–400 億（" + FG4Q + "）" if (s == "A" and y == 27) else "—")
inp("capex26", "FY2026 資本支出（公司預算）", "NT$B", None, 53.7, None, "Interested-party", CT2QF + "：董事會由 340 億上修至 537 億", "口袋學堂同；長交期設備另追加 176 億")
inp("da26", "FY2026 折舊攤提（公司預估）", "NT$B", None, 20.8, None, "Interested-party", FG4Q + "：2025 年 188 億、2026 年約 208 億", "資本支出上修後可能更高（未找到更新）")
inp("life", "新增資產平均折舊年限", "年", 6, 8, 10, "Assumed", "載板設備與廠房混合；FY2025 折舊 188 ÷ 近 8 年資本支出約合 8 年", "會計年限：未找到")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.3, 0.6, 1.0, "Assumed", "1Q26 營業費用 39.7 億、2Q26 39.9 億（營收季增 15%；Q 頁推得）", "—")
for y in (27, 28, 29):
    inp(f"nonop{y}", f"營業外淨額 FY20{y}", "NT$B", 0.0, 2.0, 4.0, "Assumed",
        "2Q26 業外 89.2 億含金融資產評價利益 92.1 億等一次性項目（口袋學堂、工商時報）；不外推" if y == 27 else "同上",
        "利息收入（GDR 募資約 US$14 億）＋權益法（Assumed）" if y == 27 else "—")
inp("tax", "所得稅與少數股權 ÷ 稅前利益", "%", 0.15, 0.18, 0.22, "Derived", "1H26：稅前 218.75、歸屬母公司 181.57 → 17.0%（工商時報）", "2Q26 15.8%；營所稅 20%")
inp("sh_g", "股數年變動（可轉債轉換、員工酬勞）", "%", 0.0, 0.0, 0.01, "Assumed", "可轉債 40 億（永豐 1Q26）；轉換價：未找到", "—")

sec("F. FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設）")
inp("q3_rev", "3Q26 營收（月營收實績加總）", "NT$B", None, 52.65, None, "Verified", TN9 + "：前 9 月 1,329.86 億 − 1H26 803.36 億", SETN + "：3Q26 526.5 億（季增 22.8%）")
inp("q3_gm", "3Q26 毛利率", "%", 0.255, 0.27, 0.285, "Assumed", "公司：3Q 毛利率可望優於 2Q 的 24.8%（2Q26 法說，Interested-party）", SETN + "：法人估 27.2%")
inp("q3_opx", "3Q26 營業費用", "NT$B", 4.0, 4.3, 4.6, "Assumed", "1Q26 3.97、2Q26 3.99（Q 頁推得）", "—")
inp("q4_g", "4Q26 營收季增", "%", 0.0, 0.05, 0.10, "Assumed", SETN + "：法人估 4Q26 579.2 億（季增 10%）", "公司未給指引")
inp("q4_gm", "4Q26 毛利率", "%", 0.265, 0.285, 0.305, "Assumed", SETN + "：法人估 30.5%", "—")
inp("q4_opx", "4Q26 營業費用", "NT$B", 4.1, 4.4, 4.8, "Assumed", "同上", "—")
inp("q_nonop", "3Q、4Q26 營業外淨額（每季）", "NT$B", 0.0, 1.0, 3.0, "Assumed", "1Q26 3.54、2Q26 8.92（含一次性評價利益）", "3Q26 財報 2026-10-29 前後公布（stockanalysis 估計）")
inp("q_sh", "3Q、4Q26 加權股數（GDR 後）", "十億股", None, 1.64, None, "Interested-party", SA + "：流通股數 1.64B（2026-10-22）", UA_GDR2 + "：5,000 萬單位（每單位 1 股，以每單位 US$27 ÷ 台股約 NT$880 推得）")

sec("G. 共識與券商（只作 Recon 與 Val 對照，不作校準）")
inp("cons_eps26", "共識 EPS FY2026（FactSet 中位數，2026-06）", "NT$", 10.5, 15.17, 18.55, "Interested-party", CNYES_FS, "2Q26 一次性業外（EPS 約 5 元）之前；本土法人 2026-10：23.42（三立）")
inp("cons_eps27", "共識 EPS FY2027（FactSet 平均，2026-06）", "NT$", 21.27, 30.37, 41.37, "Interested-party", CNYES_FS, "美系券商 2026-10：37（" + UDN10 + "）；本土法人 41.10（三立）")
inp("cons_eps28", "共識 EPS FY2028（FactSet 平均，2026-06）", "NT$", 36.59, 48.52, 60.83, "Interested-party", CNYES_FS, "美系券商 2026-10：64.8（經濟日報）")
inp("cons_rev27", "共識營收 FY2027（FactSet 平均，2026-06）", "NT$B", 193.5, 255.2, 320.2, "Interested-party", CNYES_FS, "—")
inp("cons_rev28", "共識營收 FY2028（FactSet 平均，2026-06）", "NT$B", 259.7, 326.5, 402.5, "Interested-party", CNYES_FS, "—")
inp("bk_eps27", "券商 EPS FY2027（2026-10 兩家平均：美系 37、本土 41.10）", "NT$", 37.0, 39.05, 41.10, "Interested-party", UDN10, SETN)

sec("H. 價格與市場資料")
inp("px", "股價（2026-10-22 收盤；10-23 收盤未取得）", "NT$", None, 1085, None, "Interested-party", SA + "：10-22 收盤 1,085、10-21 1,100", "HiStock 2026-10-25 頁面：1,085")
inp("fpe", "前瞻 P/E（stockanalysis，10-22）", "x", None, 42.28, None, "Interested-party", SA, "口徑：頁面未明示（約未來 12 月）")
inp("pe_ttm", "近 12 月 P/E（stockanalysis，10-22）", "x", None, 70.74, None, "Interested-party", SA, "HiStock 72.24")
inp("cons_feps", "共識前瞻 12 月 EPS（現價 ÷ 前瞻 P/E）", "NT$", None, "=px/fpe", None, "Derived", SA, "直接數字：未找到")
inp("cons_tp", "市場共識目標價（賣方平均）", "NT$", None, 1317.37, None, "Interested-party", SA + "：19 位分析師平均 1,317.37（2026-10-22 擷取）", "本土法人 1,500（三立 2026-10-19）；FactSet 2026-06：1,000")
inp("fx", "匯率", "NT$／US$", None, SIL.base("fx_twd"), None, "Assumed", "Siliconomics SIF_fx_twd（全庫共用；v0.8.2 前本模型取 32.5）", "鉅亨報價框：USD/TWD 32.508（2026-10）")
inp("tgt_abf_gap27", "ABF 供需缺口 2027（外資報告）", "%", None, 0.27, None, "Interested-party", SINO1Q + "：2026 年 8%、2027 年 27%、2028 年 35%", "Tom's Hardware（經 Ibiden 模型）：2027 約 21%")
REV5 = [104.563, 140.489, 104.036, 115.373, 131.241]
NI5 = [13.222, 29.619, 11.980, 5.082, 6.673]
inp("nim5", "近五年平均淨利率（FY2021–25；十年平均的參考）", "%", None, "=AVERAGE(" + ",".join(f"{n}/{v}" for n, v in zip(NI5, REV5)) + ")", None, "Derived",
    SA + "：營收與淨利 FY2021–25", "十年（2016–2025）營收：未找到（FY2016–20 營收未取得），收斂法依工作單改用 FY2029 × 0.8")

sec("I. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_r", "折現率／要求報酬", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%，美台日韓同一值）", "—")
inp("tp_g", "長期成長率", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "永續價值與 FY2032 收斂目標（情境 B）")
inp("tp_roe", "ROE（近三年平均，FY2023–25）", "%", None, "=AVERAGE(0.1294,0.0567,0.0731)", None, "Derived", SA + " 比率頁：FY2023 12.94%、FY2024 5.67%、FY2025 7.31%", "單一來源；本模型 DCF 改用（淨利＋折舊－資本支出），此值只作參考列")
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
Q["A1"] = "Q — FY2025 實績、1H26 實績、3Q26 營收實績與 FY2026 錨點（十億新台幣；EPS 新台幣、基本；股數十億股）。藍字＝輸入；黑字＝公式"
Q["A2"] = "FY2026 錨點＝1Q、2Q26 財報實績＋3Q26 營收實績（月營收）與 3Q26 毛利率、費用假設＋4Q26 假設。3Q26 財報約 2026-10-29 公布，屆時以實績取代 E 欄（毛利率、費用、業外、淨利）"
hdr(Q, 3, ["項目", "FY2025A", "1Q26（實績）", "2Q26（實績）", "3Q26（營收實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1（公司）", "來源2／說明"])
S1 = "FY2025：stockanalysis 損益表（公司財報）；1H26、2Q26：工商時報 2026-07-29（公司財報）"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2="—", ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"G{r}"] = ann or f"=SUM(C{r}:F{r})"; Q[f"G{r}"].number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, s1); Q.cell(r, 10, s2)
qrow(4, "營收", {"B": 131.241, "C": "=80.336-D4", "D": 42.889, "E": "=q3_rev", "F": "=E4*(1+q4_g)"}, tag="Verified／Assumed（4Q）", s2="1Q26 374.46 億（永豐 1Q26）；2Q26 428.89 億（Quartr 同）")
qrow(5, "毛利", {"B": 18.292, "C": "=C4*0.1796", "D": "=D4*0.248", "E": "=E4*q3_gm", "F": "=F4*q4_gm"}, tag="Verified（毛利率 × 營收）／Assumed", s2="毛利率：1Q26 17.96%、2Q26 24.8%（公司；Quartr 同）")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEFG"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2="FY2025 13.94%")
qrow(7, "營業費用", {c: f"={c}5-{c}8" for c in "BCD"} | {"E": "=q3_opx", "F": "=q4_opx"}, tag="Derived／Assumed", s1="毛利 − 營業利益")
qrow(8, "營業利益", {"B": 6.674, "C": "=9.408-D8", "D": 6.651, "E": "=E5-E7", "F": "=F5-F7"}, tag="Verified／Derived", s2="1H26 94.08 億（口袋學堂；工商時報 94.07）")
qrow(9, "營業外淨額（FY2025＝淨利 − 營業利益，近似）", {"B": "=B11-B8", "C": "=12.466-D9", "D": 8.924, "E": "=q_nonop", "F": "=q_nonop"}, tag="Verified／Assumed", s2="2Q26 含金融資產評價利益 92.1 億、湖口廠處分利益 5.6 億、欣興同泰減損 13.7 億（口袋學堂）")
qrow(10, "稅前利益", {c: f"={c}8+{c}9" for c in "CDEF"}, tag="Derived", s1="—", s2="1H26 公布 218.75 億（核對見第 24 列）")
qrow(11, "歸屬母公司淨利", {"B": 6.673, "C": "=18.157-D11", "D": 13.114, "E": "=E10*(1-tax)", "F": "=F10*(1-tax)"}, tag="Verified／Derived", s2="1Q26 50.43 億（永豐）；FY2025 業外與稅後淨利接近營業利益")
qrow(12, "所得稅與少數股權", {c: f"={c}10-{c}11" for c in "CDEF"}, tag="Derived", s1="—")
qrow(13, "加權股數（十億股）", {"B": "=B11/B14", "C": "=C11/C14", "D": "=D11/D14", "E": "=q_sh", "F": "=q_sh"}, ann="=AVERAGE(C13:F13)", tag="Derived", s1="淨利 ÷ EPS；3Q 起含 GDR")
qrow(14, "EPS（基本）", {"B": 4.38, "C": 3.28, "D": 8.45, "E": "=E11/E13", "F": "=F11/F13"}, fmt="0.00", ann="=G11/G13", tag="Verified／Derived", s2="1H26 11.70（工商時報）")
qrow(15, "折舊攤提（年度）", {"B": 18.8, "G": "=da26"}, ann="", tag="Interested-party", s1=FG4Q, s2="stockanalysis 現金流量表 FY2025 18.46")
qrow(16, "資本支出（年度）", {"B": 25.617, "G": "=capex26"}, ann="", tag="Verified／Interested-party", s1="FY2025：stockanalysis 現金流量表；FY2026：公司預算 537 億", s2="—")
Q.cell(18, 1, "FY2026 產品別（錨點營收 × In 頁 C 節占比）").font = BOLD
for i, (lab, f) in enumerate((("ABF 載板", "=G4*mx_abf"), ("BT 載板", "=G4*mx_bt"), ("HDI＋PCB", "=G4*mx_pcb"), ("軟板與其他", "=G4*(1-mx_abf-mx_bt-mx_pcb)")), 19):
    Q.cell(i, 1, lab); Q[f"G{i}"] = f; Q[f"G{i}"].number_format = "#,##0.000"; Q.cell(i, 8, "Derived")
Q.cell(23, 1, "檢查（參考；校準未使用券商數字）").font = BOLD
checks = [("1H26 稅前利益 自算 − 公司公布 21.875（應接近 0）", "=C10+D10-21.875", "0.000"),
          ("FY2026 營收年增", "=G4/B4-1", "0.0%"),
          ("FY2025→FY2026 現金毛利增量 ÷ 營收增量（含漲價；In 頁 ft_abf、ft_oth 的對照）", "=((G5+G15)-(B5+B15))/(G4-B4)", "0.0%"),
          ("FY2026 EPS ÷ 本土法人 23.42 − 1", "=G14/23.42-1", "0.0%"),
          ("4Q26 營收 ÷ 本土法人 57.92 − 1", "=F4/57.92-1", "0.0%")]
for i, (lab, f, fm) in enumerate(checks, 24):
    Q.cell(i, 1, lab); Q[f"B{i}"] = f; Q[f"B{i}"].number_format = fm
for nm, rr in (("rev", 4), ("gp", 5), ("opx", 7), ("oi", 8), ("nonop", 9), ("ni", 11), ("sh", 13), ("eps", 14), ("da", 15), ("cx", 16)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$G${rr}")
for nm, rr in (("abf", 19), ("bt", 20), ("pcb", 21), ("fpc", 22)):
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
DEFP = dict(scn="A", ga="g_asp", pp="ibpp_{y}", vot="v_ot", shnv="sh_nv_{y}", shot="sh_ot_{y}", aiabf="ai_abf", aipcb="ai_pcb", elo="el_o", gnai="g_nai",
            cap="cap_{y}", gmix="g_mix", carry="carry", gbt="g_bt", gpai="g_pai_{y}", gpcbn="g_pcbn", gfpc="g_fpc", ftabf="ft_abf", ftoth="ft_oth",
            cx="cx{y}_{s}", life="life", phi="phi", nonop="nonop{y}", tax="tax", mf="1", fg="fade_g{s}", fm="fade_m{s}")
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
    sec_("A. AI 加速器載板（口徑與 Ibiden 一致；D 欄＝FY2026 校準年）")
    row("unv", "NVIDIA 加速器顆數 × 情境係數", "百萬顆", {FC[y]: "=(" + "+".join(units_f(g, y) for g in NV) + f")*{um(y)}{mf(y)}" for y in BU}, note="markets/network In：Blackwell、Rubin、Rubin Ultra、Hopper")
    row("uot", "非 NVIDIA 加速器顆數 × 情境係數", "百萬顆", {FC[y]: "=(" + "+".join(units_f(g, y) for g in OT) + f")*{um(y)}{mf(y)}" for y in BU}, note="AMD、TPU、Trainium、其他 ASIC")
    row("idx", "每顆載板面積與層數指數（FY2026＝1）", "x", {"D": "=1", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['ga']})" for y in F3}}, "0.000", note="Ibiden 模型 g_asp（Analogy）")
    row("pp", "ABF 價格溢價指數（全市場；FY2026＝1）", "x", {"D": "=1", **{FC[y]: f"={p('pp', y)}" for y in F3}}, "0.00", note="Ibiden 模型 pp（Assumed）")
    row("cy", "欣興漲價年度化（CY2026 基期；FY2026＝1）", "x", {"D": "=1", **{FC[y]: f"=1+{P['carry']}" for y in F3}}, "0.00", note="In 頁 carry（Assumed）")
    row("v", "每顆 NVIDIA 加速器載板金額（全市場）＝ FY2026 金額 × 面積層數 × 溢價", "US$／顆", {FC[y]: f"=v26*{FC[y]}{R['idx']}*{FC[y]}{R['pp']}" for y in BU}, "#,##0.0", note="FY2026 金額取自 Ibiden 模型 Units 頁（權威）")
    row("shnv", "欣興份額：NVIDIA 加速器載板", "%", {FC[y]: f"={p('shnv', y)}" for y in BU}, "0.0%")
    row("shot", "欣興份額：非 NVIDIA 加速器載板（金額）", "%", {FC[y]: f"={p('shot', y)}" for y in BU}, "0.0%")
    row("accu", "欣興加速器載板營收 ＝（NVIDIA × 金額 × 份額 ＋ 非 NVIDIA × 金額 × v_ot × 份額）× 漲價年度化", "US$M",
        {FC[y]: f"=({FC[y]}{R['unv']}*{FC[y]}{R['v']}*{FC[y]}{R['shnv']}+{FC[y]}{R['uot']}*{FC[y]}{R['v']}*{P['vot']}*{FC[y]}{R['shot']})*{FC[y]}{R['cy']}" for y in BU}, "#,##0")
    row("acc", "欣興加速器載板營收（新台幣）", "NT$B", {FC[y]: f"={FC[y]}{R['accu']}*fx/1000" for y in BU}, "#,##0.00")
    sec_("B. ABF 需求與產能（營收＝MIN（需求, 產能））")
    row("ai0", "AI 載板需求（FY2026＝ABF × AI 占比）", "NT$B", {"D": f"=a26_abf*{P['aiabf']}"}, "#,##0.00", note="公司：載板 AI 占比上半年逾 60%、下半年 70%")
    row("oai", "其他 AI 載板需求（交換器、伺服器 CPU、網通 ASIC）", "NT$B", {"D": f"=D{R['ai0']}-D{R['acc']}",
        **{FC[y]: f"=$D{r}*(itc_{y % 100}/itc_26)^{P['elo']}*{um(y)}{mf(y)}*{FC[y]}{R['pp']}*{FC[y]}{R['cy']}" for y in F3}}, "#,##0.00",
        note="FY2026＝校準殘差（Derived）；之後 × AI IT 資本支出成長 ^ 彈性 × 溢價")
    row("nai", "非 AI ABF 需求（PC、一般伺服器）", "NT$B", {"D": f"=a26_abf-D{R['ai0']}", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['gnai']})" for y in F3}}, "#,##0.00")
    row("dem", "ABF 需求合計", "NT$B", {FC[y]: f"={FC[y]}{R['acc']}+{FC[y]}{R['oai']}+{FC[y]}{R['nai']}" for y in BU}, "#,##0.00")
    row("cap", "ABF 產能上限（FY2026 營收 × 產能指數 × 組合 × 溢價）", "NT$B", {"D": "=a26_abf", **{FC[y]: f"=$D{R['dem']}*{p('cap', y)}*(1+{P['gmix']})^{y - 2026}*{FC[y]}{R['pp']}*{FC[y]}{R['cy']}" for y in F3}}, "#,##0.00",
        note="FY2026 稼動率逾 90%，以 FY2026 營收為產能基準")
    row("abf", "ABF 載板營收 ＝ MIN（需求, 產能）", "NT$B", {"D": "=a26_abf", **{FC[y]: f"=MIN({FC[y]}{R['dem']},{FC[y]}{R['cap']})" for y in F3}}, "#,##0.00")
    row("fill", "需求滿足率（營收 ÷ 需求；< 100% 表示產能受限）", "%", {FC[y]: f"={FC[y]}{R['abf']}/{FC[y]}{R['dem']}" for y in BU}, "0.0%")
    row("aisub", "AI 載板營收（加速器＋其他 AI，乘滿足率）", "NT$B", {FC[y]: f"=({FC[y]}{R['acc']}+{FC[y]}{R['oai']})*{FC[y]}{R['fill']}" for y in BU}, "#,##0.00")
    row("accsh", "加速器載板 ÷ AI 載板", "%", {FC[y]: f"={FC[y]}{R['acc']}/({FC[y]}{R['acc']}+{FC[y]}{R['oai']})" for y in BU}, "0.0%")
    sec_("C. 其他產品與合計")
    row("bt", "BT 載板", "NT$B", {"D": "=a26_bt", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['gbt']})" for y in F3}}, "#,##0.00")
    rp = r
    row("pai", "AI HDI／PCB", "NT$B", {"D": f"=a26_pcb*{P['aipcb']}", "E": f"=D{rp}*(1+{p('gpai', 2027)})",
        "F": f"=D{rp}*(1+{p('gpai', 2027)})*(1+{p('gpai', 2028)})*{um(2028)}", "G": f"=D{rp}*(1+{p('gpai', 2027)})*(1+{p('gpai', 2028)})*(1+{p('gpai', 2029)})*{um(2029)}"}, "#,##0.00",
        note="基準路徑 × 顆數情境係數")
    row("pcbn", "非 AI HDI／PCB", "NT$B", {"D": f"=a26_pcb*(1-{P['aipcb']})", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['gpcbn']})" for y in F3}}, "#,##0.00")
    row("fpc", "軟板與其他", "NT$B", {"D": "=a26_fpc", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['gfpc']})" for y in F3}}, "#,##0.00")
    row("rev", "營收", "NT$B", {"C": "=a25_rev", **{FC[y]: f"={FC[y]}{R['abf']}+{FC[y]}{R['bt']}+{FC[y]}{R['pai']}+{FC[y]}{R['pcbn']}+{FC[y]}{R['fpc']}" for y in BU}}, "#,##0.0")
    row("g", "營收年增", "%", {FC[y]: f"={FC[y]}{R['rev']}/{FC[y - 1]}{R['rev']}-1" for y in FYS[1:]}, "0.0%")
    row("abfsh", "ABF ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['abf']}/{FC[y]}{R['rev']}" for y in BU}, "0.0%")
    row("aish", "AI（載板＋HDI／PCB）÷ 營收", "%", {FC[y]: f"=({FC[y]}{R['aisub']}+{FC[y]}{R['pai']})/{FC[y]}{R['rev']}" for y in BU}, "0.0%", note="公司：2H26 約 70%")
    sec_("D. 擴產折舊、毛利與損益")
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
    row("nonop", "營業外淨額", "NT$B", {"C": "=a25_nonop", "D": "=a26_nonop", **{FC[y]: f"={p('nonop', y)}" for y in F3}}, "#,##0.00", note="FY2026 含 2Q26 一次性評價利益")
    row("pti", "稅前利益", "NT$B", {FC[y]: f"={FC[y]}{R['oi']}+{FC[y]}{R['nonop']}" for y in FYS}, "#,##0.00")
    row("tax", "所得稅與少數股權", "NT$B", {FC[y]: f"={FC[y]}{R['pti']}*{P['tax']}" for y in F3}, "#,##0.00")
    row("ni", "歸屬母公司淨利", "NT$B", {"C": "=a25_ni", "D": "=a26_ni", **{FC[y]: f"={FC[y]}{R['pti']}-{FC[y]}{R['tax']}" for y in F3}}, "#,##0.00")
    row("shr", "加權股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", "E": "=a26_shend*(1+sh_g)", **{FC[y]: f"={FC[y - 1]}{r}*(1+sh_g)" for y in (2028, 2029)}}, "0.0000")
    row("eps", "EPS（基本）", "NT$", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['shr']}" for y in FYS}, "0.00")
    row("epsg", "EPS 年增", "%", {FC[y]: f"={FC[y]}{R['eps']}/{FC[y - 1]}{R['eps']}-1" for y in FYS[1:]}, "0.0%")
    row("nim", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("cxr", "資本支出 ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['cx']}/{FC[y]}{R['rev']}" for y in FYS}, "0.0%")
    row("fcf", "股東現金流 ＝ 淨利 ＋ 折舊 − 資本支出", "NT$B", {FC[y]: f"={FC[y]}{R['ni']}+{FC[y]}{R['da']}-{FC[y]}{R['cx']}" for y in F3}, "#,##0.00",
        note="不含客戶預付款（光復二廠、楊梅二廠由客戶全額預付）與營運資金變動")
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
PL["A1"] = "PL — 欣興損益（十億新台幣；EPS 新台幣、基本；歸屬母公司）。FY2025A 實績、FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設；三情境相同），FY2027–FY2029 由驅動因子推導，FY2030 起為收斂法。三情境各一個完整推導區塊（本頁下方），頂端為摘要"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A（順利）：顆數基準；資本支出 FY2027–29 為 550／500／450 億",
         "B": "情境 B（平平）：CY2028 顆數 ×0.8、CY2029 ×0.85；資本支出 480／420／380 億",
         "C": "情境 C（不如預期）：CY2028–29 顆數 ×0.6；資本支出 400／300／280 億"}
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

# ---------------- Units：份額合計檢查（與 Ibiden 同口徑） ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 欣興在 AI 加速器載板的份額與每顆金額（口徑與 Ibiden 模型一致；情境 A）。份額合計＝Ibiden（ibiden/dist）＋欣興，其餘為三星電機、景碩、南電、AT&S 等"
A_ = PB["A"]; cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
ibrel = lambda y: "1" if y == 2026 else f"ibrel_{y % 100}"
for key, lab, unit, f, fmt, note in (
        ("unv", "NVIDIA 加速器顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['unv']}", "#,##0.00", "markets/network"),
        ("uot", "非 NVIDIA 加速器顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['uot']}", "#,##0.00", ""),
        ("v", "每顆 NVIDIA 加速器載板金額（全市場）", "US$", lambda y: f"=PL!{FC[y]}{A_['v']}", "#,##0", "Ibiden 模型 Units 頁同值（FY2026 校準 × 面積層數 × 溢價）"),
        ("shnv", "欣興份額：NVIDIA", "%", lambda y: f"=PL!{FC[y]}{A_['shnv']}", "0.0%", "Assumed"),
        ("ibnv", "Ibiden 份額：NVIDIA（70% × Ibiden 份額相對值）", "%", lambda y: f"=ib_sh26*{ibrel(y)}", "0.0%", "Ibiden 模型"),
        ("sumnv", "合計：Ibiden＋欣興（NVIDIA）", "%", lambda y: f"={cyc(y)}{{shnv}}+{cyc(y)}{{ibnv}}", "0.0%", "應 ≤ 100%；其餘為三星電機等"),
        ("shot", "欣興份額：非 NVIDIA（金額）", "%", lambda y: f"=PL!{FC[y]}{A_['shot']}", "0.0%", "Assumed"),
        ("ibot", "Ibiden 份額：非 NVIDIA（Ibiden 權重 × 份額相對值 ÷ v_ot）", "%", lambda y: f"=ib_w_ot*{ibrel(y)}/v_ot", "0.0%", "Ibiden 模型的非 NVIDIA 權重假設份額相對值與 NVIDIA 同步"),
        ("sumot", "合計：Ibiden＋欣興（非 NVIDIA）", "%", lambda y: f"={cyc(y)}{{shot}}+{cyc(y)}{{ibot}}", "0.0%", "應 ≤ 100%；其餘為景碩、南電、三星電機、AT&S"),
        ("pu", "每顆 NVIDIA 加速器的欣興營收", "US$", lambda y: f"=PL!{FC[y]}{A_['v']}*PL!{FC[y]}{A_['shnv']}*PL!{FC[y]}{A_['cy']}", "#,##0", "Ibiden 對照值見 ibiden/dist Units 頁"),
        ("accu", "欣興加速器載板營收", "US$M", lambda y: f"=PL!{FC[y]}{A_['accu']}", "#,##0", ""),
        ("aisub", "欣興 AI 載板營收（含其他 AI）", "NT$B", lambda y: f"=PL!{FC[y]}{A_['aisub']}", "#,##0.0", ""),
        ("itr", "欣興 AI 載板營收 ÷ AI IT 資本支出", "%", lambda y: f"=PL!{FC[y]}{A_['aisub']}/fx/itc_{y % 100}", "0.000%", "× 每 GW IT capex 即每 GW 營收"),
        ("pgw", "每 GW 欣興 AI 載板營收（VR200 每 GW IT capex 37.59B）", "US$M", lambda y: f"=PL!{FC[y]}{A_['aisub']}/fx/itc_{y % 100}*37.59*1000", "#,##0", "Tokenomics IF_CapexIT（VR200 基準，經橋接表；與 Ibiden 模型相同係數）")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit); UR[key] = ur
    for y in BU:
        fx_ = f(y)
        for kk in ("shnv", "ibnv", "shot", "ibot"):
            fx_ = fx_.replace("{" + kk + "}", str(UR.get(kk, 0)))
        UN[f"{cyc(y)}{ur}"] = fx_; UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); ur += 1
ur += 1
UN.cell(ur, 1, "讀法：欣興在 NVIDIA 加速器載板約占兩到三成、在 ASIC 約三成；加速器載板只占欣興 AI 載板約三成，其餘是交換器、伺服器 CPU 與網通 ASIC 載板（校準殘差，公司未揭露）。份額合計留給三星電機、景碩、南電、AT&S。").font = BOLD
UN_NOTE = ur
UN.column_dimensions["A"].width = 60; UN.column_dimensions["B"].width = 10
for c in "CDEF": UN.column_dimensions[c].width = 13
UN.column_dimensions["G"].width = 66

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，年底落後 P/E）只作參考"
V["A2"] = "循環位置：欣興不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照 Ibiden、NVIDIA；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
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
V.cell(O + 9, 1, "近 12 月 P/E（stockanalysis，10-22）"); V[f"B{O + 9}"] = "=pe_ttm"; V[f"B{O + 9}"].number_format = "0.0"
V.cell(O + 10, 1, "前瞻 P/E（stockanalysis，10-22）"); V[f"B{O + 10}"] = "=fpe"; V[f"B{O + 10}"].number_format = "0.0"
V[f"E{O + 3}"] = "FY2026 三情境相同（錨點；含 2Q26 一次性業外）"; V[f"E{O + 4}"] = "FY2027 三情境相同（顆數係數從 CY2028 開始；資本支出差異從 FY2028 折舊起）"
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
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數愈高，現價愈依賴「ABF 短缺與漲價延續、擴產折舊被營收吸收」。本版未做先驗，無 Andy 合理 P/E 對照。")
WMT_NOTE = vr
# 十年 P/E：companiesmarketcap 年底 P/E；自算＝digrin 12 月未調整股價 ÷ 公司 EPS（wantgoo 歷年 EPS）；2016 EPS 0.01，P/E 無意義，兩來源都不計
PE_CMC = [None, 49.0, 17.4, 16.5, 21.0, 23.5, 5.57, 21.9, 41.8, 50.1]
PX_DEC = [None, 16.30, 22.30, 42.00, 87.40, 231.0, 120.0, 176.0, 141.0, None]
EPS_H = [0.01, 0.28, 1.15, 2.24, 3.74, 8.98, 20.08, 7.88, 3.34, 4.38]
PE_SELF = [None if p_ is None else f"={p_}/{e_}" for p_, e_ in zip(PX_DEC, EPS_H)]
PE_SELF[9] = 50.40     # 2025：digrin 無 12 月價格，改用 stockanalysis FY2025 年底 P/E（標記於說明）
T10 = vr + 2
V.cell(T10, 1, "5. 十年 P/E（2016–2025 年底；兩來源平均，分位排名用；2016 年 EPS 0.01 元，P/E 逾 1,000 倍無意義，不計）").font = BOLD
hdr(V, T10 + 1, ["年底", "兩來源平均 P/E", "companiesmarketcap", "自算"])
RB0 = T10 + 16
for i, y in enumerate(range(2016, 2026)):
    rr = T10 + 2 + i
    V.cell(rr, 1, str(y))
    if i > 0:
        V.cell(rr, 2, f"=AVERAGE(C{rr}:D{rr})").number_format = "0.00"
        V.cell(rr, 3, f"=B{RB0 + 5 + i}").number_format = "0.00"
        V.cell(rr, 4, f"=C{RB0 + 5 + i}").number_format = "0.00"
    else:
        V.cell(rr, 5, "不計（EPS 0.01 元，P/E 無意義）")
PE10 = f"$B${T10 + 2}:$B${T10 + 11}"
V.cell(T10 + 12, 1, "年數"); V[f"B{T10 + 12}"] = f"=COUNT({PE10})"
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "companiesmarketcap（年底 P/E）", "vals": PE_CMC, "tag": "Interested-party（資料商；2017–2025 共 9 年，2016 逾 1,000 倍不計）"},
    {"label": "自算（digrin 12 月未調整股價 ÷ 年度 EPS）", "vals": PE_SELF, "tag": "Derived（價格為月資料的 12 月值；EPS 取自 wantgoo 歷年股利表；2025 年 digrin 無資料，改用 stockanalysis 年底 P/E 50.40）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位，兩個來源分別取分位再平均。欣興的落後 P/E 與盈餘反向：盈餘低谷年（2017、2024–2025）P/E 42–58x，高峰年（2022）約 6x；十年多為 AI 載板需求前的 PCB 循環，套用在 FY2027 EPS 時偏誤方向不明確。")
assert V.cell(RB0 + 5, 1).value == "2016" and V.cell(RB0 + 14, 1).value == "2025", (V.cell(RB0 + 5, 1).value, V.cell(RB0 + 14, 1).value)
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
V["E14"] = "方法：欣興資本密集、模型有資本支出與折舊，改用每股（淨利＋折舊－資本支出）（PL 頁「每股股東現金流」；FY2030 起資本支出率收斂到穩態更新）"
tset(15, "reinv", "  參考：再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）；本模型不採用，只作對照（第 25 列）")
CFR = {}
for i, y in enumerate(YRS6):
    tset(16 + i, f"cf{y}", f"  FY{y}", [f"=PL!{FC[y]}{PB[s]['fps']}" for s in SC], "0.00"); CFR[y] = 16 + i
tset(22, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=SUMPRODUCT({c}16:{c}21,1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}21*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6" for c in "BCD"], "0.00",
     "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底")
tset(23, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}22*(1+tp_r)" for c in "BCD"], "0.00")
tset(24, "dcftp", "DCF 目標價（情境加權）", ["=tp_wA*B23+tp_wB*C23+tp_wC*D23", None, None], "0.00", "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁")
tset(25, "dcfref", "  參考：EPS ×（1 − g／ROE）法的 DCF 目標價（情境加權）",
     ["=(tp_wA*(SUMPRODUCT(B8:B13,1/(1+tp_r)^{1;2;3;4;5;6})+B13*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6)+tp_wB*(SUMPRODUCT(C8:C13,1/(1+tp_r)^{1;2;3;4;5;6})+C13*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6)"
      "+tp_wC*(SUMPRODUCT(D8:D13,1/(1+tp_r)^{1;2;3;4;5;6})+D13*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6))*(1-B15)*(1+tp_r)", None, None], "0.00",
     "近三年 ROE 約 9% → 再投資率約 46%；不進目標價")
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = [("Ibiden（揖斐電）", "4062.T", 53.97, "2026-10-16"), ("景碩（Kinsus）", "3189.TW", 44.70, "2026-10-22"), ("南電（Nan Ya PCB）", "8046.TW", 33.73, "2026-09-07"),
         ("三星電機（Samsung Electro-Mechanics）", "009150.KS", 43.40, "2026-08-28"), ("AT&S", "ATS.VI", 14.91, "2026-10-23")]
for i, (nm, tk, pe, dt) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左）"); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", "ABF 載板同業：Ibiden、景碩、南電、三星電機、AT&S（無中國 A 股）；南電、三星電機價格日期較舊")
tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）", [f"=tp_wA*(2/12*B{EPSR[2027]}+10/12*B{EPSR[2028]})+tp_wB*(2/12*C{EPSR[2027]}+10/12*C{EPSR[2028]})+tp_wC*(2/12*D{EPSR[2027]}+10/12*D{EPSR[2028]})", None, None], "0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], "0.00", "同業中位數 × 前瞻 EPS")
V.cell(37, 1, "0.3 結論").font = BOLD
hdr(V, 38, ["項目", "數值", "", "", "說明"])
tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B24+(1-tp_w_dcf)*B35", None, None], "0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁（0.5）")
tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], "0.00")
tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], "0.00", "stockanalysis：19 位分析師平均 1,317.37（2026-10-22 擷取）；本土法人 1,500（2026-10-19）")
tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-22 收盤）")
tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝現價 ÷ 前瞻 P/E（In 頁 cons_feps）")
tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
tset(46, "band", "價格帶（DCF 不如預期～順利）", ['=TEXT(D23,"#,##0")&"～"&TEXT(B23,"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=IF(px/{c}{EPSR[2028]}>=MAX({PE10}),1,IF(px/{c}{EPSR[2028]}<=MIN({PE10}),0,PERCENTRANK({PE10},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2017–2025 年底 P/E（第 5 節兩來源平均，2016 不計）的分位（超出十年區間時記 0% 或 100%）；只作檢查，不進目標價")
V.cell(48, 1, "目標價方法")
V["B48"] = ("DCF：股東現金流＝每股（淨利＋折舊－資本支出），r＝10%、g＝4%，FY2027–32 明確期＋永續；欣興 2026–28 年資本支出 400–550 億，"
            "明確期現金流被擴產壓低，且未計客戶預付款（光復二廠、楊梅二廠全額預付），DCF 偏低估。"
            "Comparable：Ibiden、景碩、南電、三星電機、AT&S 前瞻 P/E 中位數（無 A 股）；同業同處 ABF 短缺溢價期，倍數偏高、部分價格為 8–9 月。"
            "情境對應：A＝順利（顆數基準）、B＝平平（顆數 ×0.8／0.85）、C＝不如預期（顆數 ×0.6）")
for r_ in (24, 35, 39):
    V[f"B{r_}"].fill = KEY
V.column_dimensions["A"].width = 60
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；FY2026 恆等於錨點，其他 AI 載板在每個案例內重新校準）：對 FY2027、FY2028 營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("AI 部署 ×0.8", "FY2027 起顆數與其他 AI ×0.8", dict(mf="0.8")), ("AI 部署 ×1.2", "×1.2", dict(mf="1.2")),
         ("ABF 產能 低", "產能指數 1.25／1.50／1.70", dict(cap="cap_{y}_lo")), ("ABF 產能 高", "1.45／1.90／2.30", dict(cap="cap_{y}_hi")),
         ("ABF 溢價 低（Ibiden 低值）", "0.95／0.85／0.75", dict(pp="ibpp_{y}_lo")), ("ABF 溢價 高（Ibiden 高值）", "1.05／1.05／1.00", dict(pp="ibpp_{y}_hi")),
         ("漲價年度化 0%", "基準 5%", dict(carry="carry_lo")), ("漲價年度化 10%", "", dict(carry="carry_hi")),
         ("ABF 現金毛利率 40%", "基準 50%", dict(ftabf="ft_abf_lo")), ("ABF 現金毛利率 60%", "", dict(ftabf="ft_abf_hi")),
         ("其他產品現金毛利率 15%", "基準 25%", dict(ftoth="ft_oth_lo")), ("其他產品現金毛利率 35%", "", dict(ftoth="ft_oth_hi")),
         ("折舊年限 6 年", "基準 8 年", dict(life="life_lo")), ("折舊年限 10 年", "", dict(life="life_hi")),
         ("資本支出 走 C 路徑", "400／300／280 億（折舊較低）", dict(cx="cx{y}_C")),
         ("欣興 NVIDIA 份額 低", "20%→22%", dict(shnv="sh_nv_{y}_lo")), ("欣興 NVIDIA 份額 高", "30%→36%", dict(shnv="sh_nv_{y}_hi")),
         ("欣興 ASIC 份額 低", "22%", dict(shot="sh_ot_{y}_lo")), ("欣興 ASIC 份額 高", "40%", dict(shot="sh_ot_{y}_hi")),
         ("其他 AI 彈性 0.3", "基準 0.5", dict(elo="el_o_lo")), ("其他 AI 彈性 0.8", "", dict(elo="el_o_hi")),
         ("AI HDI／PCB 成長 低", "每年 −10pt", dict(gpai="g_pai_{y}_lo")), ("AI HDI／PCB 成長 高", "每年 +10pt", dict(gpai="g_pai_{y}_hi")),
         ("營業費用彈性 1.0", "基準 0.6", dict(phi="phi_hi")), ("稅率 22%", "基準 18%", dict(tax="tax_hi")), ("業外 0", "每年 0（基準 20 億）", dict(nonop="nonop{y}_lo"))]
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
put_case(nn, "情境 B", "CY2028 顆數 ×0.8、CY2029 ×0.85；資本支出較低", PB["B"], "PL!")
put_case(nn + 1, "情境 C", "CY2028–29 顆數 ×0.6；資本支出最低", PB["C"], "PL!")
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：ABF 產能受限時，顆數與份額的上修只部分轉成營收（滿足率 < 100%），產能與價格的影響較大；折舊年限與資本支出路徑直接改變毛利率；FY2027 三情境相同，差異從 FY2028 起。")
SENS_NOTE = nn + 3
TPS0 = nn + 5
SE.cell(TPS0, 1, "目標價參數（只改 DCF 參數；PL 收斂法的 fade_g 與資本支出率路徑不隨 tp_g 變動）").font = BOLD
hdr(SE, TPS0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
def dcf_f(c, r_, g_):
    return f"(SUMPRODUCT(Val!{c}16:{c}21,1/(1+{r_})^{{1;2;3;4;5;6}})+Val!{c}21*(1+{g_})/({r_}-{g_})/(1+{r_})^6)"
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
RC["A1"] = "Recon — 本模型 vs 券商、共識與市場口徑。只作事後對照，不為對上數字而調參數。FactSet 共識為 2026-06（2Q26 財報前），10 月券商數字只有兩家"
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
rec("FY2026 EPS vs 本土法人 23.42（2026-10-19）", 23.42, pl("eps", 2026), "三立（本土法人）", "FactSet 2026-06 中位數 15.17 為 2Q26 一次性業外之前")
rec("FY2026 EPS vs 美系券商 21（2026-10-12）", 21.0, pl("eps", 2026), "經濟日報（美系券商）", "")
rec("FY2027 EPS vs 2026-10 券商平均（美系 37、本土 41.10）", "=bk_eps27", pl("eps", 2027), "經濟日報、三立", "差距超過 30% 的原因見 README")
rec("FY2027 EPS vs FactSet 平均 30.37（2026-06）", "=cons_eps27", pl("eps", 2027), "鉅亨（FactSet 22 位）", "")
rec("FY2028 EPS vs FactSet 平均 48.52（2026-06）", "=cons_eps28", pl("eps", 2028), "鉅亨（FactSet）", "美系券商 2026-10：64.8")
rec("FY2028 EPS vs 美系券商 64.8（2026-10）", 64.8, pl("eps", 2028), "經濟日報", "")
rec("FY2027 營收 vs FactSet 平均（2026-06）", "=cons_rev27", pl("rev", 2027), "鉅亨（FactSet）", "", fmt="#,##0.0")
rec("FY2028 營收 vs FactSet 平均（2026-06）", "=cons_rev28", pl("rev", 2028), "鉅亨（FactSet）", "", fmt="#,##0.0")
rec("模型前瞻 12 月 EPS（0.2 × FY2026＋0.8 × FY2027）vs 現價 ÷ 前瞻 P/E", "=cons_feps", lambda s: f"=0.2*PL!D{PB[s]['eps']}+0.8*PL!E{PB[s]['eps']}", "stockanalysis（推導）", "口徑不同，只作量級比較")
RC.cell(rr_, 1, "B. 錨點與毛利率").font = BOLD; rr_ += 1
rec("3Q26 毛利率（假設）vs 本土法人估 27.2%（百分點）", 0.272, lambda s: "=Q!E6", "三立（本土法人）", "3Q26 財報約 10-29", fmt="0.0%", pt=True)
rec("4Q26 營收（假設）vs 本土法人估 579.2 億", 57.92, lambda s: "=Q!F4", "三立", "", fmt="#,##0.0")
rec("FY2027 毛利率 vs 4Q26 毛利率（假設）（百分點）", "=Q!F6", pl("gm", 2027), "Q 頁", "擴產折舊增量抵銷部分規模效益", fmt="0.0%", pt=True)
rec("FY2028 毛利率 vs FY2026 錨點（百分點）", "=Q!G6", pl("gm", 2028), "Q 頁", "", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "C. 產能、份額與 AI 口徑").font = BOLD; rr_ += 1
rec("FY2027 ABF 需求滿足率（< 100%＝產能受限）", None, pl("fill", 2027), "PL 頁", "外資：ABF 供需缺口 2027 年 27%（永豐 1Q26 轉述）", fmt="0.0%")
rec("FY2028 ABF 需求滿足率", None, pl("fill", 2028), "PL 頁", "", fmt="0.0%")
rec("FY2026 AI 占營收 vs 公司 2H26 約 70%（百分點）", 0.70, pl("aish", 2026), "工商時報 2Q26 法說", "全年口徑低於下半年", fmt="0.0%", pt=True)
rec("FY2027 份額合計：Ibiden＋欣興（NVIDIA）", None, lambda s: f"=Units!D{UR['sumnv']}", "Units 頁", "應 ≤ 100%", fmt="0.0%")
rec("FY2027 份額合計：Ibiden＋欣興（非 NVIDIA）", None, lambda s: f"=Units!D{UR['sumot']}", "Units 頁", "應 ≤ 100%", fmt="0.0%")
rec("FY2027 每顆 NVIDIA 加速器的欣興營收（US$）", None, lambda s: f"=Units!D{UR['pu']}", "Units 頁", "Ibiden 同期約 169（ibiden/dist Units 頁）", fmt="#,##0")
rec("FY2027 每 GW 欣興 AI 載板營收（US$M）", None, lambda s: f"=Units!D{UR['pgw']}", "Units 頁", "Ibiden 同期約 62（只含加速器載板）", fmt="#,##0")
RC.cell(rr_, 1, "D. 資本支出").font = BOLD; rr_ += 1
rec("FY2027 資本支出 vs 美系券商「逾 400 億」", 40.0, pl("cx", 2027), "經濟日報", "", fmt="#,##0.0")
rec("FY2028 資本支出 vs 美系券商「逾 400 億」", 40.0, pl("cx", 2028), "經濟日報", "", fmt="#,##0.0")
RC.column_dimensions["A"].width = 72
for c in "BCDEFGH": RC.column_dimensions[c].width = 11
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60
RECON_LAST = rr_ - 1

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司財報經媒體與資料商轉載；法說內容經媒體摘要；券商與資料商皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("公司（轉載）：工商時報 2026-07-29 2Q26 財報（營收、營業利益、業外、淨利、資本支出）", "https://www.ctee.com.tw/news/20260729700548-430201"),
        ("公司（轉載）：工商時報 2026-07-29 2Q26 法說（產品組合、應用別、稼動率、AI 占比、產能時程）", "https://www.ctee.com.tw/news/20260729701506-430201"),
        ("公司（轉載）：Quartr 2Q26 摘要（營收、毛利率、營業利益、EBITDA、1H26）", "https://quartr.com/events/unimicron-technology-corp-3037-q2-2026_F9DCOIJX"),
        ("公司（摘要）：富果 4Q25 法說摘要（FY2025、折舊 188 億 → 208 億、4Q25 產品組合、光復二廠）", "https://blog.fugle.tw/post/earnings-call-3037-2026-02-25"),
        ("媒體：口袋學堂 2Q26 法說（業外明細、ABF 52%、Blackwell 約 30%）", "https://www.pocket.tw/school/articles/7708/"),
        ("媒體：優分析 2Q26（資本支出 537 億）", "https://uanalyze.com.tw/articles/4246852366"),
        ("媒體：永豐 豐學 1Q26（1Q26 營收、毛利率、EPS、ABF 供需缺口、私募與可轉債）", "https://www.sinotrade.com.tw/richclub/hotstock/欣興-3037--1Q26-EPS繳出3-28元改寫13季新高-T-glass缺料推升ABF漲價循環的預期差再放大-豐學PRIME盤後精選--69f16461f5cc3b01237ecb00"),
        ("媒體：優分析 2026-04-10（GDR 董事會決議）", "https://uanalyze.com.tw/articles/7364850419"),
        ("媒體：優分析 2026-07（GDR 5,000 萬單位、每單位 US$26.96–27.76）", "https://uanalyze.com.tw/articles/9347152010"),
        ("媒體：TechNews 2026-10-08（9 月營收 186.51 億、前 9 月 1,329.86 億）", "https://finance.technews.tw/2026/10/08/unimicron-3037-202609-financial-report/"),
        ("券商（媒體）：三立 2026-10-19 本土法人（目標價 1,500、EPS 2026 23.42／2027 41.10、Q3、Q4 預估、產能時程）", "https://www.mirrormedia.mg/external/setn_1923649"),
        ("券商（媒體）：經濟日報 2026-10-12 美系券商（EPS 2026 21／2027 37／2028 64.8、資本支出逾 400 億）", "https://money.udn.com/money/story/5607/9808557"),
        ("共識：鉅亨 2026-06-24 FactSet（22 位分析師 EPS、營收）", "https://news.cnyes.com/news/id/6510190"),
        ("資料商：stockanalysis（現價、前瞻與近 12 月 P/E、股數、目標價、FY2021–25 損益、ROE、現金流）", "https://stockanalysis.com/quote/tpe/3037/"),
        ("資料商：stockanalysis 同業前瞻 P/E（Ibiden、景碩、南電、三星電機、AT&S）", "https://stockanalysis.com/quote/tpe/3189/"),
        ("資料商：companiesmarketcap 年底 P/E", "https://companiesmarketcap.com/unimicron/pe-ratio/"),
        ("資料商：digrin 月股價（12 月未調整值）", "https://www.digrin.com/stocks/detail/3037.TW/price"),
        ("資料商：wantgoo 歷年股利表（EPS 2016–2025）", "https://www.wantgoo.com/stock/3037/profitability/dividend-payout-ratio"),
        ("資料商：HiStock（現價核對）", "https://histock.tw/stock/3037"),
        ("本 repo 上游成品（唯讀）：ibiden、markets/network、tsmc", f"{rel(IB_F)}；{rel(NET_F)}；{rel(TSMC_F)}"),
        ("待補（人工）：3Q26 財報（約 2026-10-29）；公司原文（公開資訊觀測站財報、法說簡報）；FY2016–20 營收（十年平均淨利率）；ABF／BT 年度拆分；加速器與其他 AI 載板拆分；份額量化；折舊會計年限；客戶預付款金額；可轉債轉換價；10 月後共識（FactSet 更新）", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 110; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("欣興電子（3037.TW）收支與估值模型 v0.1（AI 半導體 Project：載板與 PCB 次產業）　2026-10-25", True),
 ("版本：v0.1（初版）。樣板：ibiden/。AI 加速器載板的每顆金額、Ibiden 份額與非 NVIDIA 權重讀 ibiden/dist（本專案權威），欣興份額與 Ibiden 合計在 Units 頁檢查。", False),
 ("模型命題：欣興的 ABF 載板營收＝MIN（需求, 產能）；需求＝AI 加速器顆數 × 每顆載板金額（面積與層數 × 溢價，與 Ibiden 同口徑）× 欣興份額，加上交換器、伺服器 CPU 等其他 AI 載板（隨 AI IT 資本支出）與非 AI ABF；產能依光復二廠、楊梅二廠時程擴張。BT、HDI／PCB、軟板另推。2026 年資本支出 537 億、2027–28 年仍在 400–550 億，擴產折舊是毛利率的關鍵。", False),
 ("核心問題：ABF 短缺與漲價能不能抵過擴產折舊，讓毛利率從 FY2026 約 25% 持續上升、EPS 在 FY2027–28 倍增；約 42 倍前瞻 P/E 隱含什麼路徑。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（延伸到 FY2032；PL 頁灰底）。金額：十億新台幣；EPS：新台幣（基本、歸屬母公司）。3Q26 財報約 2026-10-29 公布，本版只納入 3Q26 營收實績。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("加速器載板：markets/network 的 NVIDIA 與非 NVIDIA 顆數 × 情境係數（tsmc）× 每顆 NVIDIA 載板金額（Ibiden 模型 FY2026 校準 US$215 × 面積層數 +18%／年 × Ibiden 溢價指數）× 欣興份額（NVIDIA 25%→30%、非 NVIDIA 金額 30%→33%，Assumed；非 NVIDIA 每顆金額 ＝ 0.8 × NVIDIA）× 漲價年度化 1.05。", False),
 ("其他 AI 載板＝FY2026 AI 載板（ABF × 75%）− 加速器載板（校準殘差），之後 × AI IT 資本支出成長 ^ 0.5；非 AI ABF +3%。ABF 產能指數 1.35／1.70／2.00（Assumed）× 組合 +5%／年 × 溢價；營收＝MIN（需求, 產能）。BT +5%、AI HDI／PCB +20%／+15%／+10%（× 顆數情境係數）、非 AI HDI／PCB +2%、軟板與其他持平。", False),
 ("毛利＝FY2026 錨點＋ABF 現金毛利增量（50%）＋其他產品現金毛利增量（25%）＋漲價溢價營收 − 折舊增量；折舊＝前一年 ×（1 − 1／8 年）＋前一年資本支出 ÷ 8 年；資本支出依情境。營業費用彈性 0.6；業外每年 20 億；稅與少數股權 18%；股數持平（GDR 後 16.4 億股）。", False),
 ("校準：FY2026 ＝ 1H26 財報實績（EPS 11.70，含 2Q26 一次性業外約 89 億）＋3Q26 營收實績 526.5 億＋3Q26 毛利率 27%、4Q26 營收季增 5%、毛利率 28.5%（Assumed）。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（股東現金流＝每股（淨利＋折舊－資本支出），FY2027–32 明確期＋永續，r 10%、g 4%，三情境權重 0.3／0.5／0.2）＋0.5 × Comparable 目標價（Ibiden、景碩、南電、三星電機、AT&S 前瞻 P/E 中位數 × 12 個月後前瞻 EPS；無 A 股）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解。十年 P/E 規則值 × FY2027 EPS 保留為參考（不是目標價）。情境對應：A＝順利、B＝平平、C＝不如預期（原為顆數情境，另含資本支出路徑）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m，FY2032 達到；十年平均淨利率未找到，依工作單以 FY2029 × 0.8 代替）。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司數字經媒體與資料商轉載（公司原文待人工複核）；券商與資料商為 Interested-party。", False),
 ("", False),
 ("已知限制：(1) 加速器載板與其他 AI 載板的拆分公司未揭露，其他 AI 載板為校準殘差；(2) 份額、產能指數、漲價年度化皆為 Assumed，與 Ibiden 合計只做上限檢查；(3) 毛利用現金流通率近似，未建分廠稼動率與爬坡良率；(4) DCF 未計客戶預付款（光復二廠、楊梅二廠全額預付）與營運資金，現金流偏低估；(5) FY2026 錨點含 2Q26 一次性業外，FY2026 P/E 偏低；(6) 十年 P/E 多為 AI 前的 PCB 循環、與盈餘反向；(7) 共識 FactSet 為 2026-06，10 月只有兩家券商；(8) 3Q26 財報（約 10-29）後更新錨點。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ---------------- v0.1.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00／14:17）----------------
# 新列一律放在 Val、README 工作表既有內容之後，不位移既有列；既有數值不變
VER_PREV, VER_NOW = 'v0.1', 'v0.1.1'
OBS = [
 "目標價 NT$727（DCF 389、Comparable 1,065 各半），比現價 1,085（2026-10-22）低 33.0%，級距 −25%～−50%；Comparable 單獨為 −1.8%。主因是 DCF 不是 EPS：本模型前瞻 EPS 24.54 只比共識前瞻 EPS 25.66 低 4.4%，同業前瞻 P/E 中位數 43.4 倍與現價 ÷ 共識前瞻 EPS 42.3 倍相近；DCF 用每股（淨利＋折舊 − 資本支出），擴產期 FY2027 情境 A 每股現金流 9.16 只有 EPS 27.51 的 33%，DCF 比現價低 64.2%（改用 EPS ×（1 − g／ROE）只有 249，因 ROE 8.6% 低於折現率 10%）。",
 "ABF 漲價路徑最敏感：ABF 溢價取 Ibiden 低值（2028 年 0.85）時 FY2028 EPS −24.1%、取高值 +24.1%；漲價年度化 0%（基準 5%）時 FY2027 EPS −11.1%，AI 部署 ×0.8 時 −19.4%。",
 "情境分歧從 FY2028 開始：FY2027 EPS 三情境同為 27.51（比 2026-06 FactSet 共識 30.37 低 9.4%），FY2028 為 30.44／23.21／16.05。下一個檢驗點是 3Q26 財報與法說（約 2026-10-29）：ABF 報價與 2027 年資本支出。"
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
if os.environ.get("UNIMICRON_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE, "UN_NOTE": UN_NOTE, "T": T, "T10": T10,
               "TPS0": TPS0, "TPS_LAST": TPS_LAST, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0},
              open(os.environ["UNIMICRON_ROWMAP"], "w"), ensure_ascii=False, indent=1)

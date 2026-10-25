# KLA（KLAC）收支與估值模型 v0.2（AI 半導體 Project 步驟 5：設備與材料——量測與檢測）
# 用法：python3 klac/build_klac.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度 7 月到隔年 6 月：FY2027＝2026-07～2027-06（一半在 CY2026、一半在 CY2027）。金額：百萬美元（US$M）；EPS 為美元（2026-06 的 10 比 1 分割調整後）。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- 讀入設備市場模型成品（唯讀；不修改） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
EQ_F = only("markets/equipment/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
ws_ = openpyxl.load_workbook(EQ_F, data_only=True)["Mkt"]
EQK = {"先進邏輯（台積電設備＋其他代工中屬先進者）〔AI〕": "L1", "其他代工與邏輯": "L2", "其他前段": "L3",
       "DRAM（含 HBM）〔AI〕": "M1", "NAND": "M2", "先進封裝〔AI〕": "P"}
EQ = {}; blk = None; SEMI_WFE = None
for r in ws_.iter_rows():
    lab = str(r[0].value or "")
    for s in SC:
        if lab.startswith(f"5{s}. "): blk = s
    if lab.startswith("6. "): blk = None
    if blk and lab in EQK:
        EQ[(blk, EQK[lab])] = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
    if blk and lab == f"設備總額（情境 {blk}）":
        EQ[(blk, "T")] = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
    if lab == "SEMI 晶圓前段設備（WFE）":
        SEMI_WFE = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert len(EQ) == 21 and SEMI_WFE, sorted(EQ)
PEAK, PEAK_SRC = CP.peak_years(ROOT)   # 設備股共用循環位置規則

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：百萬美元（US$M）；設備市場：十億美元（US$B）。每個輸入附兩個來源（來源2 找不到者寫「未找到」）；公司原文（SEC 8-K／10-K）算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間，不當點估計用。SemiAnalysis 未使用。共識只放 Recon，不作校準"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
CALL = "KLA FQ4-26 法說（2026-07-28；MarketBeat 逐字稿）"
sec("A. CY2026 錨點（上半年實績＋公司對下半年的說法）")
inp("h2up", "CY2026 下半年營收較上半年增幅", "%", 0.15, 0.20, 0.25, "Interested-party", CALL + "：下半年約比上半年高 20%", "TIKR 2026-07-29 法說摘要（同）；FQ1-27 指引 US$4.0B±0.2B 與共識 FQ2-27 US$4.49B 合計 8.49B＝上半年 7.07B×1.20（Recon）；低／高為 Assumed 區間")
inp("g_svc_h2", "CY2026 下半年服務營收較上半年增幅", "%", 0.03, 0.06, 0.09, "Assumed", "無：公司稱服務 2026 年成長在長期目標 13–15% 的下緣", "6%≈年增 13% 的半年版；只影響錨點中系統與服務的拆分")
inp("g_oth26", "SPTS＋PCB 兩事業 CY2026 年增", "%", 0.15, 0.25, 0.30, "Interested-party", CALL + "：Orbotech 合計產品 CY2026 成長 25% 以上", "第二來源：未找到。CY2025 基期取 Q 頁四季合計")
inp("ap26", "先進封裝製程控制營收 CY2026", "US$M", 1000, 1100, 1400, "Interested-party", CALL + "：約 US$1.1B，年增 70% 以上", "handsoff substack（2026-05-31，引公司 FQ3-26）：FY2026 1.3–1.4B（會計年度口徑，與 1.1B 的日曆年口徑不同，未解）；高＝1.4B")
inp("ap_g26", "先進封裝製程控制營收 CY2026 年增", "%", None, 0.70, None, "Interested-party", CALL + "：年增 70% 以上（取 70%）", "用於推 CY2025 基期＝ap26÷1.7")
inp("k_ap", "先進封裝製程控制營收成長 ÷ 先進封裝設備市場成長（倍數）", "x", 1.0, 2.0, 2.5, "Interested-party", CALL + "：成長率接近封裝市場的 2 倍", "低＝與市場同速、高＝2.5 倍（Assumed）；下滑時同樣放大（β 解讀，Assumed）")
sec("B. 製程控制強度（系統營收 ÷ 加權前段 WFE；以 CY2026 錨點校準）")
inp("fl26", "CY2026 半導體製程控制系統營收中代工邏輯占比", "%", 0.60, 0.63, 0.66, "Interested-party", CALL + "：2026 年代工邏輯「low to mid-60s」；FQ1-27 為 73%／記憶體 27%", "第二來源：未找到。取 63%；用來推記憶體相對強度")
inp("m_mem", "記憶體 WFE 的製程控制強度 ÷ 代工邏輯（相對倍數）", "x", "=(1-fl26_hi)/fl26_hi*Units!D8/Units!D11", "=(1-fl26)/fl26*Units!D8/Units!D11", "=(1-fl26_lo)/fl26_lo*Units!D8/Units!D11", "Derived", "（1−代工邏輯占比）÷ 代工邏輯占比 × CY2026 代工邏輯 WFE ÷ 記憶體 WFE（設備模型情境 A）", "代工邏輯 WFE＝先進邏輯＋其他代工邏輯＋其他前段（含光罩，KLA 光罩檢測歸此）；低欄用 fl26 高值")
inp("kla_wfe25", "CY2025 WFE（公司口徑，含先進封裝）", "US$B", None, 120, None, "Interested-party", CALL + "：2025 年約 US$120B", "設備模型 2025 前段＋先進封裝＝" + f"{EQ[('A','L1')][0][0]+EQ[('A','L2')][0][0]+EQ[('A','L3')][0][0]+EQ[('A','M1')][0][0]+EQ[('A','M2')][0][0]+EQ[('A','P')][0][0]:.1f}（相近）")
inp("share_up", "KLA 占 WFE 比重 2025→2030 的提升", "pt", 0.0, 0.015, 0.03, "Interested-party", CALL + "：公司 2030 模型稱市占較 2025 提升 150bp（±25）", "第二來源：未找到。低＝不提升、高＝300bp（Assumed）；公司未說明分母，本模型解讀為「半導體製程控制營收 ÷ WFE」")
inp("pc_u", "製程控制強度年增（先進製程與先進封裝檢測強度上升）", "%/年", "=((s25+share_up_lo)/s25)^(1/5)-1", "=((s25+share_up)/s25)^(1/5)-1", "=((s25+share_up_hi)/s25)^(1/5)-1", "Derived", "＝[(CY2025 占比＋提升)÷CY2025 占比]^(1/5)−1；CY2025 占比＝半導體製程控制營收 ÷ 公司 WFE（Units）", "自 CY2027 起套用於 CY2026 校準的強度")
inp("kla_sh_pc", "KLA 製程控制市占（顯示用：換算製程控制占 WFE）", "%", 0.53, 0.58, 0.63, "Interested-party", "handsoff substack 2026-05-31（引 BofA CY2025）：製程控制市占約 58%、晶圓檢測 81.9%（賣方）", "第二來源：未找到。只用於顯示，不影響營收")
sec("C. 中國（重要變數；基準不另扣，風險放 Sens）")
inp("chn_sh26", "中國占營收 FY2026", "%", None, 0.30, None, "Verified", "KLA FY2026 10-K：中國 30%（FY2025 33%、FY2024 43%）", "handsoff substack：FQ3-26 單季 24%（約 US$820M）")
inp("chn_sh25", "中國占營收 FY2025", "%", None, 0.33, None, "Verified", "KLA FY2026 10-K", "—")
inp("chn_sh24", "中國占營收 FY2024", "%", None, 0.43, None, "Verified", "KLA FY2026 10-K", "—")
inp("chn_cut", "CY2027 起中國系統營收下降比例（出口管制）", "%", 0.25, 0.0, 0.50, "Assumed", "無：10-K 稱出口許可未獲准已傷及訂單；BIS 子公司規則暫停一年至 2026-11", "基準 0（設備模型的 WFE 已含中國）；Sens 列 25%／50%")
sec("D. 服務與 SPTS＋PCB（會計年度，前一年 × 成長率）")
for s, v, o in (("A", 0.14, 0.12), ("B", 0.12, 0.08), ("C", 0.10, 0.00)):
    inp(f"g_svc_{s}", f"服務營收 FY2027–29 年增（{s}）", "%", v - 0.02, v, v + 0.02, "Assumed", "公司長期目標 13–15%，2027 年預期在上緣；約 80% 為合約制（" + CALL + "）", "FY2026 實績 +16.5%（Q 頁）；C 取裝機量成長放緩")
    inp(f"g_oth_{s}", f"SPTS＋PCB 營收 FY2027–29 年增（{s}）", "%", o - 0.04, o, o + 0.04, "Assumed", "公司：EPC／PCB 長期成長 high single digit；CY2026 Orbotech 合計 >25%", "FY2026 實績 +11.5%（Q 頁）")
sec("E. 毛利與營業費用（不寫死利益率）")
inp("gm_inc", "增量毛利率（毛利變動 ÷ 營收變動；相對 FY2026）", "%", 0.60, "=(a26_gp-a25_gp)/(a26_rev-a25_rev)", 0.65, "Derived", "兩點：FY2025、FY2026 實績（Q 頁）", "公司：增量毛利率 60–65% 至 2027（非 GAAP；" + CALL + "）＝低／高；記憶體定價逆風約 100bp")
inp("opex_q1", "營業費用 FQ1-27（R&D＋SG&A）", "US$M", None, 690, None, "Interested-party", CALL + "：約 US$690M", "第二來源：未找到")
inp("opex_step", "之後每季營業費用增加", "US$M／季", 15, 17.5, 20, "Interested-party", CALL + "：每季增加約 US$15–20M", "取中點")
inp("opex27", "營業費用 FY2027（4×FQ1＋6×每季增額）", "US$M", None, "=4*opex_q1+6*opex_step", None, "Derived", "690＋707.5＋725＋742.5", "三情境相同（FY2027 已大致決定）")
for s, g in (("A", 0.08), ("B", 0.05), ("C", 0.02)):
    inp(f"g_opex_{s}", f"營業費用 FY2028–29 年增（{s}）", "%", g - 0.03, g, g + 0.03, "Assumed", "無：FY2026 +11.4%、FY2027 指引隱含 +7.6%", "—")
sec("F. 營業外、稅、股數")
inp("oth_exp", "營業外淨額（利息與其他）每年", "US$M", -200, -100, 0, "Interested-party", CALL + "：其他收支約每季 US$25M 費用", "FY2026 實績 −54.9（利息 −284.4＋其他 +229.6）；低／高 Assumed")
inp("tax", "有效稅率", "%", 0.13, 0.145, 0.17, "Interested-party", CALL + "：規劃稅率 14.5%", "FY2026 實績 13.8%（Q 頁）")
inp("shares27", "稀釋股數 FY2027", "百萬股", None, 1312, None, "Interested-party", "KLA FQ4-26 新聞稿（SEC 8-K）：FQ1-27 指引稀釋股數 1,312M", "FY2026 稀釋 1,319.6M、FQ4-26 1,315.0M")
inp("bb", "稀釋股數年減（買回）FY2028–29", "%", 0.0, "=1-a26_sh/a25_sh", 0.025, "Derived", "FY2025→FY2026 稀釋股數變化（Q 頁）", "低／高 Assumed")
sec("G. 公司指引、共識（只供 Recon）、股價、循環位置、目標價權重")
inp("g_q1_rev", "FQ1-27 營收指引", "US$M", 3800, 4000, 4200, "Interested-party", "KLA FQ4-26 新聞稿（SEC 8-K，2026-07-28）", "TIKR 2026-07-29")
inp("g_q1_gm", "FQ1-27 GAAP 毛利率指引", "%", 0.606, 0.616, 0.626, "Interested-party", "同上", "非 GAAP 62.5%")
inp("g_q1_eps", "FQ1-27 GAAP 稀釋 EPS 指引", "US$", 1.04, 1.14, 1.24, "Interested-party", "同上", "非 GAAP 1.16±0.10")
inp("cons_q2_rev", "共識 FQ2-27 營收", "US$M", None, 4490, None, "Interested-party", "Yahoo Finance（2026-10-25 讀取；21 位分析師）", "第二來源：未找到")
inp("cons_rev27", "共識營收 FY2027", "US$M", None, 18140, None, "Interested-party", "Yahoo Finance（24 位）：18.14B", "stockanalysis（2026-10-20）：18.1B")
inp("cons_eps27", "共識 EPS FY2027（非 GAAP）", "US$", None, 5.44, None, "Interested-party", "Yahoo Finance（25 位）：5.44", "stockanalysis（2026-10-20，24 位）：5.46")
inp("cons_rev28", "共識營收 FY2028", "US$M", None, 21570, None, "Interested-party", "Yahoo Finance（21 位）：21.57B", "第二來源：未找到（stockanalysis 需付費）")
inp("cons_eps28", "共識 EPS FY2028（非 GAAP）", "US$", None, 6.74, None, "Interested-party", "Yahoo Finance（22 位）：6.74", "第二來源：未找到")
for s in SC:
    inp(f"pk_{s}", f"情境 {s}：設備市場見頂年度（日曆年；循環位置規則用）", "年", None, PEAK[s], None, "Derived", PEAK_SRC[s], "設備股共用規則（tools/cycle_position.py）：FY(t) 早於見頂年＝中段、等於＝見頂、晚於＝谷底；KLA 的 FY(t) 對照日曆年 t（理由見 Val 第 4 節）" if s == "A" else "同上")
inp("px", "股價（stockanalysis 2026-10-23 收盤）", "US$", None, 145.61, None, "Verified", "stockanalysis.com 歷史股價：2026-10-23 收盤 145.61", "digrin：145.61；tipranks：2026-10-22 收盤 147.83（10-22 單日 −9.5%，原因未查明）")
inp("tp_wA", "目標價情境權重 A", "%", None, 0.3, None, "Assumed", "Andy 可改", "三個權重合計應為 100%（Val 第 0 節檢查）")
inp("tp_wB", "目標價情境權重 B", "%", None, 0.5, None, "Assumed", "Andy 可改", "—")
inp("tp_wC", "目標價情境權重 C", "%", None, 0.2, None, "Assumed", "Andy 可改", "—")

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
    defname(key, f"In!$E${r}"); INROW[key] = r
    if lo is not None: defname(key + "_lo", f"In!$D${r}")
    if hi is not None: defname(key + "_hi", f"In!$F${r}")
    r += 1
for col, w in zip("ABCDEFGHI", [12, 58, 10, 10, 11, 10, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

# ---------------- Q：實績（百萬美元；EPS 為分割調整後美元） ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績（百萬美元；EPS 為美元，2026-06-12 除權的 10 比 1 分割調整後）。會計年度 7–6 月：FY2026＝2025-07～2026-06。藍字＝輸入（SEC 8-K 新聞稿）；黑字＝公式（Derived）"
Q["A2"] = "FY2027 第一季（2026-07～09）：預定 2026-10-28 盤後公布（tipranks 財報行事曆；KLA 去年同季為 10-29）；建置日 2026-10-25 尚未公布，故不納入，指引見下方。公布後更新本頁 I 欄"
hdr(Q, 3, ["項目", "FQ3-25（2025-01～03）", "FQ4-25（2025-04～06）", "FY26 上半（2025-07～12，FQ1＋FQ2）", "FQ3-26（2026-01～03）", "FQ4-26（2026-04～06）", "FY2025 全年", "FY2026 全年", "FQ1-27（未公布）", "標記", "來源1（公司）", "來源2／說明"])
S_Q326 = "FQ3-26 新聞稿（SEC 8-K，2026-04-29）"; S_Q426 = "FQ4-26 新聞稿（SEC 8-K，2026-07-28）"
QR = {}
def qrow(r, key, label, vals, tag, s1, s2, fmt="#,##0.0"):
    QR[key] = r
    Q.cell(r, 1, label)
    for j, v in zip("BCDEFGH", vals):
        if v is None: continue
        c = Q[f"{j}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    Q.cell(r, 10, tag); Q.cell(r, 11, s1); Q.cell(r, 12, s2)
H1 = lambda a, b: f"={a}-{b}"   # 九個月 − FQ3
qrow(4, "rev", "營收", [3063.0, 3174.741, H1(9921.9, 3415.1), 3415.1, 3657.556, 12156.162, 13579.476], "Verified", "FQ3-25：" + S_Q326 + "前年同期欄；FQ4-25、全年：" + S_Q426 + "；上半＝九個月 9,921.9 − FQ3-26", "stockanalysis 年度：FY2025 12,156、FY2026 13,579；FQ2-26 單季 3,297（FQ3-26 新聞稿）")
qrow(5, "prod", "產品營收", [2393.8, 2472.182, H1(7616.4, 2640.3), 2640.3, 2837.151, 9472.854, 10453.537], "Verified", "同上", "第二來源：未找到（年度服務占比 23%，10-K）")
qrow(6, "svc", "服務營收", [669.2, 702.559, H1(2305.5, 774.8), 774.8, 820.405, 2683.308, 3125.939], "Verified", "同上", "FY2026 10-K：服務約占營收 23%（3,125.9÷13,579.5＝23.0%）")
qrow(7, "spc", "  事業別：半導體製程控制（含服務）", [2738.8, 2877.647, H1(8988.0, 3083.9), 3083.9, 3256.781, 10947.359, 12244.733], "Verified", "同上（事業別表）", "—")
qrow(8, "spts", "  事業別：特殊半導體製程（SPTS）", [156.5, 141.866, H1(424.4, 164.0), 164.0, 159.704, 587.107, 584.064], "Verified", "同上", "—")
qrow(9, "pcb", "  事業別：PCB 與元件檢測", [168.6, 154.106, H1(509.3, 167.6), 167.6, 241.110, 621.721, 750.415], "Verified", "同上", "—")
qrow(10, "fx", "  匯率影響等（營收 − 三事業，Derived）", [f"={c}4-{c}7-{c}8-{c}9" for c in "BCDEFGH"], "Derived", "—", "公司事業別表另列 FX；四捨五入差")
qrow(11, "sys", "半導體製程控制系統（產品營收 − SPTS − PCB；Derived）", [f"={c}5-{c}8-{c}9" for c in "BCDEFGH"], "Derived", "產品營收扣兩事業全部營收（其服務部分未揭露，故系統略為低估、與服務合計不變）", "系統＋服務＋兩事業＝營收（恆等）")
qrow(12, "oth", "SPTS＋PCB", [f"={c}8+{c}9" for c in "BCDEFGH"], "Derived", "—", "—")
qrow(13, "cost", "營業成本", [1175.7, 1207.286, H1(3842.0, 1327.7), 1327.7, 1413.108, 4751.867, 5255.060], "Verified", "同上", "—")
qrow(14, "gp", "毛利", [f"={c}4-{c}13" for c in "BCDEFGH"], "Derived", "—", "stockanalysis 年度毛利：FY2025 7,404、FY2026 8,324")
qrow(15, "gmr", "毛利率（GAAP）", [f"={c}14/{c}4" for c in "BCDEFGH"], "Derived", "—", "公司非 GAAP：FQ4-26 62.4%", "0.0%")
qrow(16, "rd", "研發費用", [338.0, 352.989, H1(1133.1, 388.8), 388.8, 399.023, 1360.334, 1532.118], "Verified", "同上", "—")
qrow(17, "sga", "銷管費用", [248.9, 262.706, H1(840.0, 291.1), 291.1, 291.477, 1029.734, 1131.518], "Verified", "同上", "法說稱 FQ4-26 SG&A 283（非 GAAP）")
qrow(18, "imp", "商譽與無形資產減損", [0, 0, 0, 0, 0, 239.1, 0], "Verified", "FY2025 減損 239.1（FQ3-26 新聞稿九個月欄、FQ4-26 新聞稿）", "—")
qrow(19, "opx", "營業費用（研發＋銷管；不含減損）", [f"={c}16+{c}17" for c in "BCDEFGH"], "Derived", "—", "—")
qrow(20, "oi", "營業利益（毛利 − 營業費用 − 減損）", [f"={c}14-{c}19-{c}18" for c in "BCDEFGH"], "Derived", "—", "stockanalysis 年度營業利益：FY2026 5,661（相符）；FY2025 5,014（不含減損）")
qrow(21, "nonop", "營業外淨額（其他收入 − 利息）", [None, None, None, None, None, "=171.487-302.166", "=229.585-284.44"], "Verified", S_Q426 + "：利息 FY25 302.2／FY26 284.4；其他收入 171.5／229.6", "—")
qrow(22, "taxv", "所得稅", [None, None, None, None, None, 582.805, 775.154], "Verified", "同上", "—")
qrow(23, "taxr", "有效稅率", [None, None, None, None, None, "=G22/(G20+G21)", "=H22/(H20+H21)"], "Derived", "—", "—", "0.0%")
qrow(24, "ni", "淨利（GAAP）", [1088.4, 1202.849, H1(3467.7, 1201.0), 1201.0, 1363.059, 4061.643, 4830.771], "Verified", "同上", "stockanalysis：FY2025 4,062、FY2026 4,831")
qrow(25, "ck_ni", "檢查：稅前 − 稅 − 淨利（年度，應約 0）", [None, None, None, None, None, "=G20+G21-G22-G24", "=H20+H21-H22-H24"], "—", "—", "—", "0.0")
qrow(26, "eps", "EPS（GAAP 稀釋，分割調整後）", [0.816, 0.91, "=(26.26-9.12)/10", 0.912, 1.04, 3.04, 3.66], "Verified／Derived", "FQ4-26 新聞稿（分割後）：FQ4-25 0.91、FY2025 3.04、FY2026 3.66、FQ4-26 1.04；FQ3-26 新聞稿（分割前 ÷10）：FQ3-25 8.16、FQ3-26 9.12、九個月 26.26", "SEC XBRL（10-K）：FY2025 3.04、FY2026 3.66；stockanalysis 同", "0.000")
qrow(27, "sh", "稀釋加權股數（百萬股，分割調整後）", [1333.0, 1327.341, None, 1317.5, 1314.986, 1337.502, 1319.633], "Verified", S_Q426 + "；FQ3：分割前 133.30／131.75 ×10", "—", "#,##0.0")
qrow(28, "ck", "檢查：FY2026 − (上半＋FQ3＋FQ4)（營收、產品、服務、淨利依序 B–E 欄）", ["=H4-D4-E4-F4", "=H5-D5-E5-F5", "=H6-D6-E6-F6", "=H24-D24-E24-F24"], "—", "九個月與單季皆取新聞稿四捨五入值，差額為 0.1 以內的捨入差", "—", "0.00")
Q["I4"] = "預定 2026-10-28 盤後公布，尚未公布"
for nm in ("rev", "prod", "svc", "spc", "spts", "pcb", "sys", "oth", "gp", "rd", "sga", "imp", "opx", "oi", "nonop", "taxv", "ni", "eps", "sh"):
    defname(f"a25_{nm}", f"Q!$G${QR[nm]}"); defname(f"a26_{nm}", f"Q!$H${QR[nm]}")
Q.cell(30, 1, "FQ1-27 公司指引（2026-07-28）").font = BOLD
hdr(Q, 31, ["項目", "指引", "低", "高", "", "", "", "", "", "標記", "來源1（公司）", "來源2／說明"])
for i, (nm, k, f) in enumerate((("營收（US$M）", "g_q1_rev", "#,##0"), ("GAAP 毛利率", "g_q1_gm", "0.0%"), ("GAAP 稀釋 EPS（US$）", "g_q1_eps", "0.00")), 32):
    Q.cell(i, 1, nm); Q.cell(i, 2, f"={k}").number_format = f; Q.cell(i, 3, f"={k}_lo").number_format = f; Q.cell(i, 4, f"={k}_hi").number_format = f
    Q.cell(i, 10, "Interested-party"); Q.cell(i, 11, S_Q426); Q.cell(i, 12, "非 GAAP：毛利率 62.5%、EPS 1.16；稀釋股數 1,312M")
Q.column_dimensions["A"].width = 56
for c in "BCDEFGHI": Q.column_dimensions[c].width = 15
Q.column_dimensions["J"].width = 18; Q.column_dimensions["K"].width = 80; Q.column_dimensions["L"].width = 60
Q.freeze_panes = "B4"

# ---------------- Units：設備市場（設備模型成品）與 KLA 日曆年實績 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 設備市場（十億美元，日曆年；讀 markets/equipment 成品三情境）與 KLA 日曆年實績（百萬美元；由 Q 頁各季加總）"
hdr(UN, 3, ["項目", "單位", "CY2025", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ccy = lambda y: chr(ord("C") + y - 2025)   # C＝2025 … G＝2029
UR = {}; ur = 4
NM = {"L1": "先進邏輯〔AI〕", "L2": "其他代工與邏輯", "L3": "其他前段（類比、功率、光罩等）", "M1": "DRAM（含 HBM）〔AI〕", "M2": "NAND", "P": "先進封裝〔AI〕", "T": "設備總額"}
for s in SC:
    UN.cell(ur, 1, f"情境 {s}：設備市場（markets/equipment 5{s} 區塊）").font = BOLD; ur += 1
    for k in ("L1", "L2", "L3", "M1", "M2", "P", "T"):
        if k == "L3":
            pass
        vals, src = EQ[(s, k)]
        UN.cell(ur, 1, "  " + NM[k]); UN.cell(ur, 2, "US$B")
        for j, v in enumerate(vals):
            c = UN.cell(ur, 3 + j, v); c.font = BLUE; c.number_format = "0.00"
        UN.cell(ur, 8, src); UR[(s, k)] = ur; ur += 1
        if k == "L3":
            UN.cell(ur, 1, "  代工邏輯 WFE 小計（L）"); UN.cell(ur, 2, "US$B")
            for y in range(2025, 2030): UN[f"{ccy(y)}{ur}"] = f"={ccy(y)}{ur-3}+{ccy(y)}{ur-2}+{ccy(y)}{ur-1}"; UN[f"{ccy(y)}{ur}"].number_format = "0.00"
            UR[(s, "L")] = ur; ur += 1
        if k == "M2":
            UN.cell(ur, 1, "  記憶體 WFE 小計（M）"); UN.cell(ur, 2, "US$B")
            for y in range(2025, 2030): UN[f"{ccy(y)}{ur}"] = f"={ccy(y)}{ur-2}+{ccy(y)}{ur-1}"; UN[f"{ccy(y)}{ur}"].number_format = "0.00"
            UR[(s, "M")] = ur; ur += 1
    ur += 1
# m_mem 公式引用 Units!D8（A 的 L，2026）與 D11（A 的 M，2026）——確認列號
assert UR[("A", "L")] == 8 and UR[("A", "M")] == 11, (UR[("A", "L")], UR[("A", "M")])
UN.cell(ur, 1, "SEMI 晶圓前段設備（WFE；對帳用）"); UN.cell(ur, 2, "US$B")
for j, v in enumerate(SEMI_WFE[0]): c = UN.cell(ur, 3 + j, v); c.font = BLUE; c.number_format = "0.00"
UN.cell(ur, 8, SEMI_WFE[1] + "（SEMI 2026-07；設備商協會，傾向樂觀）"); UR["semi"] = ur; ur += 2
UN.cell(ur, 1, "KLA 日曆年實績（百萬美元；CY2025＝FQ3-25＋FQ4-25＋FY26 上半；CY2026 上半＝FQ3-26＋FQ4-26）").font = BOLD; ur += 1
hdr(UN, ur, ["項目", "單位", "CY2025", "CY2026 上半", "", "", "", "說明"]); ur += 1
for nm, lab in (("rev", "營收"), ("prod", "產品"), ("svc", "服務"), ("spc", "半導體製程控制（事業別，含服務）"), ("oth", "SPTS＋PCB"), ("sys", "半導體製程控制系統（產品 − SPTS − PCB）")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, "US$M")
    UN[f"C{ur}"] = f"=Q!B{QR[nm]}+Q!C{QR[nm]}+Q!D{QR[nm]}"; UN[f"D{ur}"] = f"=Q!E{QR[nm]}+Q!F{QR[nm]}"
    for c in "CD": UN[f"{c}{ur}"].number_format = "#,##0.0"
    defname(f"cy25_{nm}", f"Units!$C${ur}"); defname(f"h1_{nm}", f"Units!$D${ur}"); ur += 1
UN.cell(ur, 1, "CY2025 半導體製程控制營收 ÷ 公司 WFE（s25）"); UN.cell(ur, 2, "%"); UN[f"C{ur}"] = "=cy25_spc/(kla_wfe25*1000)"; UN[f"C{ur}"].number_format = "0.00%"
UN.cell(ur, 8, "pc_u（強度年增）的基期；公司稱 2030 年較 2025 提升 150bp"); defname("s25", f"Units!$C${ur}"); ur += 2
UN.cell(ur, 1, "KLA 會計年度與日曆年：FY(t)＝CY(t−1) 下半＋CY(t) 上半。CY2027 起每年上下半各 1/2（Assumed，忽略年內斜率）；CY2026 的下半＝錨點 − 上半實績。")
UN.column_dimensions["A"].width = 60; UN.column_dimensions["B"].width = 10; UN.column_dimensions["H"].width = 80
for c in "CDEFG": UN.column_dimensions[c].width = 12

# ---------------- 引擎（PL 與 Sens 共用） ----------------
FYC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2030)}   # FY2025＝C … FY2029＝G
DEFP = dict(m="m_mem", u="pc_u", kap="k_ap", ap26="ap26", h2="h2up", cut="chn_cut", gmi="gm_inc", tax="tax", bb="bb", oexp="oth_exp",
            gsh2="g_svc_h2", go26="g_oth26", opq="opex27", scn="A", gsvc=None, goth=None, gopx=None)
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    gsvc = P["gsvc"] or f"g_svc_{s}"; goth = P["goth"] or f"g_oth_{s}"; gopx = P["gopx"] or f"g_opex_{s}"
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "CY2025／FY2025A", "CY2026／FY2026A", "CY2027／FY2027", "CY2028／FY2028", "CY2029／FY2029", "說明"]); r += 1
    ws.cell(r, 1, "A. 日曆年（C＝CY2025 … G＝CY2029）").font = BOLD; r += 1
    def row(key, label, unit, fs, fmt="#,##0.0", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 8, note); R[key] = r; r += 1
    CYS = range(2025, 2030); c = ccy
    row("W", "加權前段 WFE（代工邏輯＋m × 記憶體）", "US$M", {c(y): f"=(Units!{c(y)}{UR[(s,'L')]}+{P['m']}*Units!{c(y)}{UR[(s,'M')]})*1000" for y in CYS}, note=f"設備模型情境 {s}；m＝記憶體相對強度")
    row("Peq", "先進封裝設備市場", "US$M", {c(y): f"=Units!{c(y)}{UR[(s,'P')]}*1000" for y in CYS})
    row("tot", "營收 CY2026 錨點（上半實績 ×（2＋下半增幅））", "US$M", {"D": f"=h1_rev*(2+{P['h2']})"}, note="上半實績＋公司說法（下半約高 20%）")
    row("svc", "服務 CY2026 錨點", "US$M", {"D": f"=h1_svc*(2+{P['gsh2']})"})
    row("oth", "SPTS＋PCB CY2026 錨點", "US$M", {"D": f"=cy25_oth*(1+{P['go26']})"})
    rs = r + 3  # 系統列位置（ap、fe、r 之後）
    row("ap", "先進封裝製程控制系統營收", "US$M", {"C": f"={P['ap26']}/(1+ap_g26)", "D": f"={P['ap26']}", **{c(y): f"={c(y-1)}{r}*(1+{P['kap']}*({c(y)}{R['Peq']}/{c(y-1)}{R['Peq']}-1))" for y in range(2027, 2030)}}, note="CY2027 起：前一年 ×（1＋k × 封裝設備市場成長）")
    row("fe", "前段製程控制系統營收", "US$M", {"C": f"=C{rs}-C{r-1}", "D": f"=D{rs}-D{r-1}", **{c(y): f"={c(y)}{r+1}*{c(y)}{R['W']}" for y in range(2027, 2030)}}, note="CY2025／2026＝系統 − 先進封裝；之後＝強度 × 加權 WFE")
    row("ri", "製程控制強度（前段系統 ÷ 加權前段 WFE）", "%", {"C": f"=C{r-1}/C{R['W']}", "D": f"=D{r-1}/D{R['W']}", **{c(y): f"=$D{r}*(1+{P['u']})^{y-2026}" for y in range(2027, 2030)}}, fmt="0.00%", note="CY2026 校準；之後每年 ×（1＋pc_u）")
    assert r == rs
    row("sys", "半導體製程控制系統營收 CY", "US$M", {"C": "=cy25_sys", "D": f"=D{R['tot']}-D{R['svc']}-D{R['oth']}", **{c(y): f"=({c(y)}{R['fe']}+{c(y)}{R['ap']})*(1-chn_sh26*{P['cut']})" for y in range(2027, 2030)}}, note="CY2026＝錨點 − 服務 − 兩事業；CY2027 起 ×（1 − 中國占比 × 中國下降比例）")
    row("pcs", "顯示：製程控制占前段 WFE（強度 ÷ KLA 市占）", "%", {c(y): f"={c(y)}{R['ri']}/kla_sh_pc" for y in CYS}, fmt="0.0%", note="KLA 市占 58% 為賣方數字，只供顯示")
    ws.cell(r, 1, "B. 會計年度（C＝FY2025A … G＝FY2029；FY(t)＝CY(t−1) 下半＋CY(t) 上半）").font = BOLD; r += 1
    fy = FYC
    row("fsys", "半導體製程控制系統營收 FY", "US$M", {"C": "=a25_sys", "D": "=a26_sys", "E": f"=(D{R['sys']}-h1_sys)+0.5*E{R['sys']}", "F": f"=0.5*E{R['sys']}+0.5*F{R['sys']}", "G": f"=0.5*F{R['sys']}+0.5*G{R['sys']}"}, note="FY2027＝CY2026 下半（錨點 − 上半實績）＋1/2 CY2027")
    row("fsvc", "服務營收", "US$M", {"C": "=a25_svc", "D": "=a26_svc", **{fy[y]: f"={fy[y-1]}{r}*(1+{gsvc})" for y in range(2027, 2030)}})
    row("foth", "SPTS＋PCB 營收", "US$M", {"C": "=a25_oth", "D": "=a26_oth", **{fy[y]: f"={fy[y-1]}{r}*(1+{goth})" for y in range(2027, 2030)}})
    row("rev", "營收", "US$M", {fy[y]: f"={fy[y]}{R['fsys']}+{fy[y]}{R['fsvc']}+{fy[y]}{R['foth']}" for y in range(2025, 2030)}, note="FY2025、FY2026＝實績（恆等）")
    row("gp", "毛利", "US$M", {"C": "=a25_gp", "D": "=a26_gp", **{fy[y]: f"=a26_gp+{P['gmi']}*({fy[y]}{R['rev']}-a26_rev)" for y in range(2027, 2030)}}, note="FY2026 毛利＋增量毛利率 ×（營收 − FY2026 營收）")
    row("gm", "毛利率", "%", {fy[y]: f"={fy[y]}{R['gp']}/{fy[y]}{R['rev']}" for y in range(2025, 2030)}, fmt="0.0%")
    row("opx", "營業費用（研發＋銷管）", "US$M", {"C": "=a25_opx", "D": "=a26_opx", "E": f"={P['opq']}", **{fy[y]: f"={fy[y-1]}{r}*(1+{gopx})" for y in range(2028, 2030)}})
    row("oi", "營業利益", "US$M", {"C": "=a25_oi", **{fy[y]: f"={fy[y]}{R['gp']}-{fy[y]}{R['opx']}" for y in range(2026, 2030)}}, note="FY2025 含減損 239.1")
    row("oim", "營業利益率", "%", {fy[y]: f"={fy[y]}{R['oi']}/{fy[y]}{R['rev']}" for y in range(2025, 2030)}, fmt="0.0%")
    row("noi", "營業外淨額", "US$M", {"C": "=a25_nonop", "D": "=a26_nonop", **{fy[y]: f"={P['oexp']}" for y in range(2027, 2030)}})
    row("ni", "淨利", "US$M", {"C": "=a25_ni", "D": "=a26_ni", **{fy[y]: f"=({fy[y]}{R['oi']}+{fy[y]}{R['noi']})*(1-{P['tax']})" for y in range(2027, 2030)}})
    row("sh", "稀釋股數", "百萬股", {"C": "=a25_sh", "D": "=a26_sh", "E": "=shares27", **{fy[y]: f"={fy[y-1]}{r}*(1-{P['bb']})" for y in range(2028, 2030)}})
    row("eps", "EPS（GAAP 稀釋）", "US$", {"C": "=a25_eps", "D": "=a26_eps", **{fy[y]: f"={fy[y]}{R['ni']}/{fy[y]}{R['sh']}" for y in range(2027, 2030)}}, fmt="0.00")
    row("chn", "顯示：中國營收（× FY2026 中國占比；CY2027 起扣下降）", "US$M", {"D": "=a26_rev*chn_sh26", **{fy[y]: f"={fy[y]}{R['rev']}*chn_sh26*(1-{P['cut']}*IF({y}=2027,0.5,1))" for y in range(2027, 2030)}}, fmt="#,##0", note="只供顯示")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — KLA 損益（US$M；EPS 為美元）。FY2025、FY2026 為實績，FY2027–FY2029 為模型。三情境各一個完整推導區塊（見下方），本頁頂端為摘要"
PB = {}; r0 = 45
TITLE = {"A": "情境 A：設備市場續增（2029 見頂）", "B": "情境 B：設備市場 2028 持平、2029 −5%", "C": "情境 C：設備市場 2028 −20%（2027 見頂）"}
for s in SC:
    PB[s] = engine(PL, r0, TITLE[s], scn=s); r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026A", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; PL.cell(sr, 8, TITLE[s]); sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "US$M", "#,##0"), ("fsys", "  半導體製程控制系統", "US$M", "#,##0"), ("fsvc", "  服務", "US$M", "#,##0"), ("foth", "  SPTS＋PCB", "US$M", "#,##0"),
                                ("gm", "毛利率", "%", "0.0%"), ("oi", "營業利益", "US$M", "#,##0"), ("oim", "營業利益率", "%", "0.0%"), ("ni", "淨利", "US$M", "#,##0"), ("eps", "EPS（GAAP 稀釋）", "US$", "0.00")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in range(2025, 2030):
            PL[f"{FYC[y]}{sr}"] = f"={FYC[y]}{PB[s][key]}"; PL[f"{FYC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對：FY2025、FY2026 模型營收 − 實績（應為 0）").font = BOLD; sr += 1
for s in SC:
    PL.cell(sr, 1, f"  情境 {s}"); PL[f"C{sr}"] = f"=ABS(C{SUMR[(s,'rev')]}-a25_rev)+ABS(D{SUMR[(s,'rev')]}-a26_rev)"; PL[f"C{sr}"].number_format = "0.000"; sr += 1
PL.column_dimensions["A"].width = 54; PL.column_dimensions["B"].width = 12; PL.column_dimensions["H"].width = 60
for c in "CDEFG": PL.column_dimensions[c].width = 13

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值。跨公司規則：合理倍數＝10 年（2016–2025 年底）P/E 的分位，依循環位置：見頂年 25 分位、中段中位數、谷底 75 分位。循環位置依設備股共用規則（第 4 節）。EPS 一律 GAAP 稀釋"
V["A2"] = "P/E 定義（自算）：日曆年底收盤價 ÷ 「當期會計年度」GAAP 稀釋 EPS（年底所在會計年度，例如 2025-12 對 FY2026＝2026-06 期）；另列「前期」（上一個已結束會計年度）。價格與 EPS 皆為 2026-06 的 10 比 1 分割調整後"
# 0. 目標價（列號固定，引用第 5 節）
V.cell(4, 1, "0. 目標價與上檔空間").font = BOLD
hdr(V, 5, ["項目", "數值", "說明"])
TP0 = 6
# 第 1 節以後
V.cell(16, 1, "1. 公司公布 EPS（會計年度，GAAP 稀釋，美元，分割調整後）— 兩個來源").font = BOLD
hdr(V, 17, ["會計年度（6 月結束）", "來源1：SEC XBRL（10-K；分割前 ÷10）", "來源2：macrotrends（6 月 TTM，÷10）", "兩者差距", "標記", "說明"])
EPS1 = {2016: 0.449, 2017: 0.588, 2018: 0.510, 2019: 0.749, 2020: 0.770, 2021: 1.337, 2022: 2.192, 2023: 2.415, 2024: 2.03, 2025: 3.04, 2026: 3.66}
EPS2 = {2016: 0.449, 2017: 0.588, 2018: 0.510, 2019: 0.749, 2020: 0.770, 2021: 1.337, 2022: 2.192, 2023: 2.415, 2024: 2.028, 2025: 3.037, 2026: 3.66}
NOTE = {2018: "美國稅改一次性費用使 FY2018 偏低", 2024: "SEC 為 FY2026 10-K 重編的分割後 2.03（分割前 20.28）", 2025: "分割前 30.37", 2026: "來源2 為 stockanalysis（macrotrends 2026-06 列異常，未用）；Q 頁"}
EP = {}
for i, y in enumerate(range(2016, 2027), 18):
    V.cell(i, 1, f"FY{y}（{y}-06）")
    V.cell(i, 2, EPS1[y]).font = BLUE; V.cell(i, 3, EPS2[y]).font = BLUE
    V.cell(i, 4, f"=B{i}/C{i}-1").number_format = "0.0%"
    for c in "BC": V[f"{c}{i}"].number_format = "0.000"
    V.cell(i, 5, "Verified"); V.cell(i, 6, NOTE.get(y, "SEC companyconcept EarningsPerShareDiluted（10-K 年度值）")); EP[y] = i
V.cell(29, 1, "FY2026 EPS 改由 Q 頁實績帶入第 2 節（＝a26_eps）；其餘年度用來源1。兩來源差距皆來自四捨五入").font = BOLD
V.cell(30, 1, "2. 10 年 P/E（年底）").font = BOLD
hdr(V, 31, ["年底", "收盤價 US$（分割調整）", "當期 FY EPS", "前期 FY EPS", "P/E 當期（自算）", "P/E 前期（自算）", "macrotrends TTM P/E", "自算當期 − macrotrends", "收盤價來源2：macrotrends（÷10，含股利調整）", "標記", "來源／說明"])
PX = {2016: 78.68, 2017: 105.07, 2018: 89.49, 2019: 178.17, 2020: 258.91, 2021: 430.11, 2022: 377.03, 2023: 581.30, 2024: 630.12, 2025: 1215.08}
MT_PX = {2016: 77.51, 2017: 103.76, 2018: 88.61, 2019: 176.85, 2020: 257.45, 2021: 428.18, 2022: 375.83, 2023: 580.14, 2024: 629.42, 2025: 1214.77}
MT_PE = {2016: 14.09, 2017: 25.00, 2018: 9.69, 2019: 24.94, 2020: 29.36, 2021: 22.06, 2022: 15.44, 2023: 29.37, 2024: 26.49, 2025: 35.36}
for i, y in enumerate(range(2016, 2026), 32):
    V.cell(i, 1, f"{y}-12")
    V.cell(i, 2, f"={PX[y]}/10").font = BLUE
    V.cell(i, 3, "=a26_eps" if y == 2025 else f"=B{EP[y+1]}"); V.cell(i, 4, f"=B{EP[y]}")
    V.cell(i, 5, f"=B{i}/C{i}"); V.cell(i, 6, f"=B{i}/D{i}")
    V.cell(i, 7, MT_PE[y]).font = BLUE; V.cell(i, 8, f"=E{i}-G{i}"); V.cell(i, 9, f"={MT_PX[y]}/10").font = BLUE
    for c in "BCDI": V[f"{c}{i}"].number_format = "#,##0.000"
    for c in "EFGH": V[f"{c}{i}"].number_format = "0.0"
    V.cell(i, 10, "Verified／Derived")
    V.cell(i, 11, {2018: "當期 P/E 11.9 為十年最低（2018Q4 記憶體循環下修）", 2025: "當期 EPS＝FY2026 實績 3.66；最高 33.2"}.get(y, "收盤價：digrin「Real price」（未調整）12 月收盤 ÷10"))
V.cell(43, 1, "收盤價來源1：digrin（Real price，年底收盤，未做股利調整）；來源2：macrotrends（含股利調整，故略低，I 欄）。macrotrends P/E 為 TTM 口徑（年底已公布的近四季 EPS），與自算的當期口徑不同，只作對照。")
V.cell(45, 1, "3. 分位").font = BOLD
hdr(V, 46, ["統計", "P/E 當期（自算）", "P/E 前期（自算）", "macrotrends P/E（參考）", "用途"])
for i, (nm, p, use) in enumerate((("25 分位", 0.25, "見頂年"), ("中位數", 0.5, "中段"), ("75 分位", 0.75, "谷底")), 47):
    V.cell(i, 1, nm); V.cell(i, 5, use)
    V.cell(i, 2, f"=PERCENTILE(E32:E41,{p})"); V.cell(i, 3, f"=PERCENTILE(F32:F41,{p})"); V.cell(i, 4, f"=PERCENTILE(G32:G41,{p})")
    for c in "BCD": V[f"{c}{i}"].number_format = "0.0"
V.cell(50, 1, "年數"); V["B50"] = "=COUNT(E32:E41)"; V["C50"] = "=COUNT(F32:F41)"; V["D50"] = "=COUNT(G32:G41)"
defname("pe_peak", "Val!$B$47"); defname("pe_mid", "Val!$B$48"); defname("pe_trough", "Val!$B$49")
V.cell(51, 1, "偏誤方向：2016–2018 年（AI 需求出現前）當期 P/E 13.4、20.6、11.9 偏低，拉低 25 分位與中位數；2023、2025 年（28.7、33.2）為 AI 期高倍數。照規則計算，不另換倍數。")
# 4. 循環位置
V.cell(53, 1, "4. 循環位置（設備股共用規則；見頂年度讀設備市場模型，不在本模型另訂）").font = BOLD
V.cell(54, 1, CP.rule_text(PEAK) + "。KLA 的 FY(t)＝t−1 年 7 月～t 年 6 月，上下半各一半；以 FY(t) 對照日曆年 t（結束年）。驗證：以 FY 加權的設備總額（1/2 CY(t−1)＋1/2 CY(t)）的見頂會計年度與此判定一致（右表）")
hdr(V, 55, ["情境", "設備市場見頂年度", "FY2027 對照日曆年", "判定", "倍數", "FY 加權設備總額 FY2027", "FY 加權設備總額 FY2028", "FY 加權見頂（FY2027–29）"])
for i, s in enumerate(SC, 56):
    V.cell(i, 1, s); V.cell(i, 2, f"=pk_{s}"); V.cell(i, 3, 2027).font = BLUE
    V.cell(i, 4, "=" + CP.excel_rule(f"C{i}", f"pk_{s}"))
    V.cell(i, 5, f'=IF(D{i}="見頂",pe_peak,IF(D{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    t = UR[(s, "T")]
    V.cell(i, 6, f"=0.5*Units!D{t}+0.5*Units!E{t}").number_format = "0.0"; V.cell(i, 7, f"=0.5*Units!E{t}+0.5*Units!F{t}").number_format = "0.0"
    V.cell(i, 8, f'=IF(AND(F{i}>=G{i},F{i}>=0.5*Units!F{t}+0.5*Units!G{t}),"FY2027",IF(G{i}>=0.5*Units!F{t}+0.5*Units!G{t},"FY2028","FY2029 或之後"))')
# 5. 隱含股價
V.cell(60, 1, "5. 隱含股價（美元）：規則值＝FY2027 EPS × 依循環位置的倍數；其餘倍數僅供參考").font = BOLD
hdr(V, 61, ["情境", "EPS 年度", "EPS", "倍數類型", "倍數", "隱含股價", "相對現價", "備註"])
vr = 62; VAL = {}
for s in SC:
    for y, typ, nm in ((2027, f"E{56 + SC.index(s)}", "rule"), (2027, "pe_peak", "參考：25 分位"), (2027, "pe_mid", "參考：中位數"), (2027, "pe_trough", "參考：75 分位"),
                       (2028, "pe_peak", "參考：25 分位"), (2028, "pe_mid", "參考：中位數"), (2028, "pe_trough", "參考：75 分位")):
        V.cell(vr, 1, s); V.cell(vr, 2, f"FY{y}"); V.cell(vr, 3, f"=PL!{FYC[y]}{SUMR[(s,'eps')]}").number_format = "0.00"
        V.cell(vr, 4, nm if nm != "rule" else f'="規則值："&D{56 + SC.index(s)}'); V.cell(vr, 5, f"={typ}").number_format = "0.0"
        V.cell(vr, 6, f"=C{vr}*E{vr}").number_format = "#,##0.0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
        if nm == "rule":
            for cc in range(1, 9): V.cell(vr, cc).fill = KEY
            VAL[s] = vr
        vr += 1
assert vr == 83
V.cell(84, 1, "6. 目前遠期 P/E").font = BOLD
hdr(V, 85, ["項目", "A", "B", "C", "備註"])
fr = 86
for nm, f, fm, note in (("現價（stockanalysis 2026-10-23 收盤）", ["=px"], "0.00", "2026-10 月內由 206.89（10-02）跌至 145.61"),
                        ("遠期 P/E：現價 ÷ 共識 FY2027 EPS（非 GAAP）", ["=px/cons_eps27"], "0.0", "stockanalysis 顯示 26.65x"),
                        ("現價 ÷ FY2026 EPS（實績）", ["=px/a26_eps"], "0.0", ""),
                        ("現價 ÷ FY2027 EPS（模型）", [f"=px/PL!E{SUMR[(s,'eps')]}" for s in SC], "0.0", "A／B／C"),
                        ("現價 ÷ FY2028 EPS（模型）", [f"=px/PL!F{SUMR[(s,'eps')]}" for s in SC], "0.0", "A／B／C"),
                        ("現價隱含倍數 ÷ 規則倍數（情境 A）— 以 FY2027 A EPS", [f"=B{fr+3}/E56"], "0.00", ">1＝現價高於規則給的倍數")):
    V.cell(fr, 1, nm)
    for j, ff in enumerate(f): V.cell(fr, 2 + j, ff).number_format = fm
    V.cell(fr, 5, note); fr += 1
# 0. 目標價內容
TPR = [("目標價（12 個月，情境加權）", f"=tp_wA*B{TP0+2}+tp_wB*B{TP0+3}+tp_wC*B{TP0+4}", "0.00", "權重 × 各情境隱含價"),
       ("上檔空間（目標價 ÷ 現價 − 1）", f"=B{TP0}/px-1", "0.0%", "現價見第 6 節"),
       ("情境 A 隱含價（樂觀）", f"=F{VAL['A']}", "0.00", "第 5 節規則值（設備市場續增）"),
       ("情境 B 隱含價（基準）", f"=F{VAL['B']}", "0.00", "第 5 節規則值（設備市場 2028 持平）"),
       ("情境 C 隱含價（下檔）", f"=F{VAL['C']}", "0.00", "第 5 節規則值（設備市場 2028 −20%）"),
       ("情境權重 A／B／C", "=tp_wA", "0%", None),
       ("目標價方法", "P/E 規則值：FY2027（2027-06 期）GAAP 稀釋 EPS × 2016–2025 年底當期 P/E 依循環位置取分位（A、B 中段＝中位數，C 見頂＝25 分位；設備股共用規則）。十年倍數含 AI 前（2016–2018）低倍數期，偏誤方向為偏低（壓低目標價）；照規則計算，未另換倍數", None, None),
       ("檢查：權重合計（應為 100%）", "=tp_wA+tp_wB+tp_wC", "0%", None)]
for i, (lab, f, fm, note) in enumerate(TPR, TP0):
    V.cell(i, 1, lab); V.cell(i, 2, f)
    if fm: V.cell(i, 2).number_format = fm
    if note: V.cell(i, 3, note)
    if i == TP0:
        for cc in (1, 2, 3): V.cell(i, cc).fill = KEY
V.cell(TP0 + 5, 3, "=tp_wB").number_format = "0%"; V.cell(TP0 + 5, 4, "=tp_wC").number_format = "0%"
V.column_dimensions["A"].width = 50
for c in "BCDEFGHIJ": V.column_dimensions[c].width = 16
V.column_dimensions["K"].width = 70

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A）：對 FY2027、FY2028 營收與 EPS 的影響。每個案例是一個完整的推導區塊（本頁下方，與 PL 同一引擎）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收 US$M", "FY2027 EPS US$", "相對基準", "FY2028 營收 US$M", "FY2028 EPS US$", "相對基準"])
cases = [("中國系統營收 −25%（CY2027 起）", "出口管制擴大（Assumed）", dict(cut="chn_cut_lo")), ("中國系統營收 −50%（CY2027 起）", "", dict(cut="chn_cut_hi")),
         ("製程控制強度不再提升", "pc_u＝0", dict(u="pc_u_lo")), ("製程控制強度提升加倍", "2030 較 2025 +300bp", dict(u="pc_u_hi")),
         ("下半年增幅 15%", "CY2026 錨點較低", dict(h2="h2up_lo")), ("下半年增幅 25%", "", dict(h2="h2up_hi")),
         ("記憶體相對強度 低（代工邏輯占 66%）", "", dict(m="m_mem_lo")), ("記憶體相對強度 高（代工邏輯占 60%）", "", dict(m="m_mem_hi")),
         ("先進封裝倍數 1.0x", "與封裝市場同速", dict(kap="k_ap_lo")), ("先進封裝倍數 2.5x", "", dict(kap="k_ap_hi")),
         ("先進封裝 CY2026 1.4B", "取 substack 會計年度口徑上緣", dict(ap26="ap26_hi")),
         ("增量毛利率 60%", "公司 60–65% 下緣", dict(gmi="gm_inc_lo")), ("增量毛利率 65%", "上緣", dict(gmi="gm_inc_hi")),
         ("營業費用 FY2028–29 年增 5%", "", dict(gopx="g_opex_A_lo")), ("營業費用 FY2028–29 年增 11%", "", dict(gopx="g_opex_A_hi")),
         ("服務年增 12%", "", dict(gsvc="g_svc_A_lo")), ("服務年增 16%", "", dict(gsvc="g_svc_A_hi")),
         ("稅率 13%", "", dict(tax="tax_lo")), ("稅率 17%", "", dict(tax="tax_hi")),
         ("營業外 −200", "", dict(oexp="oth_exp_lo")), ("股數不減", "買回停止", dict(bb="bb_lo")), ("股數年減 2.5%", "", dict(bb="bb_hi"))]
def put_case(i, name, setting, R, pre=""):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    SE[f"C{i}"] = f"={pre}E{R['rev']}"; SE[f"D{i}"] = f"={pre}E{R['eps']}"; SE[f"F{i}"] = f"={pre}F{R['rev']}"; SE[f"G{i}"] = f"={pre}F{R['eps']}"
    SE[f"E{i}"] = f"=D{i}/D$4-1"; SE[f"H{i}"] = f"=G{i}/G$4-1"
    for c in "CF": SE[f"{c}{i}"].number_format = "#,##0"
    for c in "DG": SE[f"{c}{i}"].number_format = "0.00"
    for c in "EH": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
SB = {}
for n, (nm, setting, ov) in enumerate(cases):
    SB[nm] = engine(SE, 50 + n * 36, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm])
nn = 5 + len(cases)
put_case(nn, "設備市場情境 B", "2028 持平、2029 −5%；服務、兩事業、費用成長 B", PB["B"], "PL!")
put_case(nn + 1, "設備市場情境 C", "2028 −20%；服務、兩事業、費用成長 C", PB["C"], "PL!")
SE.cell(nn + 3, 1, "讀法：FY2027 的一半是 CY2026 下半（錨點，三情境相同），另一半是 CY2027（三情境的設備市場相同），所以情境差異到 FY2028 才明顯；中國與強度假設自 CY2027 起作用，FY2027 只吃一半。")
SE.column_dimensions["A"].width = 40; SE.column_dimensions["B"].width = 40
for c in "CDEFGH": SE.column_dimensions[c].width = 15

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司指引、共識、公司的 WFE 說法、SEMI。只作事後對照，不為對上數字而調整參數。差距 > 10% 的原因見 README"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.0", pt=False):
    global rr_
    RC.cell(rr_, 1, label); RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        L = chr(64 + cv)
        RC.cell(rr_, cd, f"={L}{rr_}-B{rr_}" if pt else f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y: (lambda s: f"=PL!{FYC[y]}{SUMR[(s, key)]}")
eg = lambda key, col: (lambda s: f"=PL!{col}{PB[s][key]}")
RC.cell(rr_, 1, "A. CY2026 下半錨點 vs 公司指引＋共識").font = BOLD; rr_ += 1
rec("CY2026 下半營收（錨點 − 上半實績）vs FQ1-27 指引＋共識 FQ2-27", "=g_q1_rev+cons_q2_rev", lambda s: f"=PL!D{PB[s]['tot']}-h1_rev", "公司 8-K（FQ1 4.0B）＋Yahoo 共識（FQ2 4.49B）", "錨點用公司「下半約高 20%」，非共識")
RC.cell(rr_, 1, "B. FY2027、FY2028 vs 共識（Yahoo Finance；共識 EPS 為非 GAAP，模型為 GAAP）").font = BOLD; rr_ += 1
rec("FY2027 營收", "=cons_rev27", pl("rev", 2027), "Yahoo（18.14B）；stockanalysis 18.1B")
rec("FY2027 EPS", "=cons_eps27", pl("eps", 2027), "Yahoo 5.44；stockanalysis 5.46", "GAAP 比非 GAAP 低約 2–3%（FY2026：3.66 vs 3.76）", fmt="0.00")
rec("FY2028 營收", "=cons_rev28", pl("rev", 2028), "Yahoo（21.57B）")
rec("FY2028 EPS", "=cons_eps28", pl("eps", 2028), "Yahoo 6.74", "", fmt="0.00")
RC.cell(rr_, 1, "C. WFE：設備模型 vs 公司說法、SEMI（US$B）").font = BOLD; rr_ += 1
rec("CY2025 WFE（前段＋先進封裝）vs 公司約 120", "=kla_wfe25", lambda s: f"=Units!C{UR[(s,'L')]}+Units!C{UR[(s,'M')]}+Units!C{UR[(s,'P')]}", CALL, "設備模型 2025 以 SEMI 實績校準", fmt="0.0")
rec("CY2026 WFE（前段＋先進封裝）vs 公司「low 150s」（取 152）", 152, lambda s: f"=Units!D{UR[(s,'L')]}+Units!D{UR[(s,'M')]}+Units!D{UR[(s,'P')]}", CALL, "設備模型 2026 比公司說法高；本模型以 KLA 自身的 CY2026 錨點校準強度，所以這個差異不進入營收", fmt="0.0")
RC.cell(rr_ - 1, 2).font = BLUE
rec("CY2026 前段 WFE vs SEMI WFE", f"=Units!D{UR['semi']}", lambda s: f"=Units!D{UR[(s,'L')]}+Units!D{UR[(s,'M')]}", "SEMI 2026-07（設備模型成品）", "", fmt="0.0")
rec("CY2027 前段 WFE vs SEMI WFE（年增連乘）", f"=Units!E{UR['semi']}", lambda s: f"=Units!E{UR[(s,'L')]}+Units!E{UR[(s,'M')]}", "同上", "公司引述 2027 共識約 190B（含封裝；未確認）", fmt="0.0")
RC.cell(rr_, 1, "D. 強度與市占（顯示；非獨立檢驗）").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "製程控制強度 CY2025（前段系統 ÷ 加權前段 WFE）"); RC.cell(rr_, 3, f"=PL!C{PB['A']['ri']}").number_format = "0.00%"; RC.cell(rr_, 10, "CY2026 校準值見下一列；兩者差距反映設備模型 2026 WFE 高於公司說法"); rr_ += 1
RC.cell(rr_, 1, "製程控制強度 CY2026（校準）"); RC.cell(rr_, 3, f"=PL!D{PB['A']['ri']}").number_format = "0.00%"; rr_ += 1
RC.cell(rr_, 1, "製程控制占前段 WFE CY2026（強度 ÷ KLA 市占 58%）"); RC.cell(rr_, 3, f"=PL!D{PB['A']['pcs']}").number_format = "0.0%"; RC.cell(rr_, 10, "業界常見說法約 11–13%（未找到可引用來源）"); rr_ += 1
RC.cell(rr_, 1, "E. 其他資訊（非對帳）").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "FQ1-27 營收指引 ÷ 模型 FY2027（A）"); RC.cell(rr_, 3, f"=g_q1_rev/PL!E{SUMR[('A','rev')]}").number_format = "0.0%"; RC.cell(rr_, 10, "營收逐季上升，Q1 低於 25% 屬正常"); rr_ += 1
RC.cell(rr_, 1, "模型 FY2027 中國營收（A；× FY2026 占比 30%）"); RC.cell(rr_, 3, f"=PL!E{PB['A']['chn']}").number_format = "#,##0"; rr_ += 1
RC.column_dimensions["A"].width = 60
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 40; RC.column_dimensions["J"].width = 80

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（SEC 原文可開啟；法說數字取自逐字稿網站）"; SO["A1"].font = BOLD
srcs = [("KLA FQ4-26 與 FY2026 新聞稿（SEC 8-K Exhibit 99.1，2026-07-28）", "https://www.sec.gov/Archives/edgar/data/0000319201/000031920126000024/exhibit991earningsrelease7.htm"),
 ("KLA FQ3-26 新聞稿（SEC 8-K Exhibit 99.1，2026-04-29）", "https://www.sec.gov/Archives/edgar/data/0000319201/000031920126000014/exhibit991earningsrelease3.htm"),
 ("KLA FY2026 10-K（中國占比、服務占比、訂單、出口管制）", "https://www.sec.gov/Archives/edgar/data/0000319201/000031920126000027/klac-20260630.htm"),
 ("SEC XBRL companyconcept：EarningsPerShareDiluted（歷年 EPS）", "https://data.sec.gov/api/xbrl/companyconcept/CIK0000319201/us-gaap/EarningsPerShareDiluted.json"),
 ("KLA FQ4-26 法說逐字稿（MarketBeat）", "https://www.marketbeat.com/earnings/reports/2026-7-28-kla-co-stock/"),
 ("TIKR：KLA FQ4-26 摘要（2026-07-29）", "https://www.tikr.com/blog/kla-corporations-q4-earnings-hit-a-record-3-66-billion-the-guide-points-higher-still"),
 ("MIAX：KLAC 10 比 1 分割通知（除權 2026-06-12）", "https://www.miaxglobal.com/sites/default/files/alert-files/KLAC_Split_58941.pdf"),
 ("tipranks：財報日 2026-10-28（FQ1-27）、10 月收盤價", "https://www.tipranks.com/stocks/klac/earnings"),
 ("stockanalysis：股價歷史（2026-10-23 收盤 145.61）", "https://stockanalysis.com/stocks/klac/history/"),
 ("stockanalysis：年度財報（FY2022–26）與共識（2026-10-20）", "https://stockanalysis.com/stocks/klac/forecast/"),
 ("Yahoo Finance：共識（FY2027、FY2028、FQ1、FQ2）", "https://finance.yahoo.com/quote/KLAC/analysis/"),
 ("digrin：年底收盤價（Real price）", "https://www.digrin.com/stocks/detail/KLAC/price"),
 ("macrotrends：P/E 與 TTM EPS 歷史", "https://www.macrotrends.net/stocks/charts/KLAC/kla/pe-ratio"),
 ("handsoff substack（2026-05-31；引 BofA）：製程控制市占 58%、FQ3-26 中國 24%", "https://handsoff.substack.com/p/kla-beat-across-the-board-process"),
 ("市場層：markets/equipment 成品（本 repo）", rel(EQ_F)),
 ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 70; SO.column_dimensions["B"].width = 120

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("KLA（KLAC）收支與估值模型 v0.2（AI 半導體 Project 步驟 5：設備與材料——量測與檢測）　2026-10-25", True),
 ("模型命題：KLA 的系統營收跟著晶圓廠設備支出（WFE）走，再乘上製程控制強度——先進製程與先進封裝每片晶圓需要更多量測與檢測。股價取決於設備市場在 2027 年之後放緩時，強度上升與服務能否撐住盈餘，以及中國（FY2026 占營收 30%）的出口管制風險。", False), ("", False),
 ("會計年度：7 月到隔年 6 月（FY2027＝2026-07～2027-06，一半在 CY2026、一半在 CY2027）。FY(t)＝CY(t−1) 下半＋CY(t) 上半。金額百萬美元；EPS 為 GAAP 稀釋、2026-06-12 除權的 10 比 1 分割調整後。FY2027 第一季預定 2026-10-28 公布，建置日（2026-10-25）尚未公布，故不納入。", False), ("", False),
 ("驅動因子 → 推導量", True),
 ("系統營收（日曆年）＝製程控制強度 × 加權前段 WFE（設備模型三情境；代工邏輯＋m × 記憶體）＋先進封裝製程控制（封裝設備成長 × 2 倍）；強度以 CY2026 錨點（上半實績＋公司「下半約高 20%」）校準，之後每年 +pc_u（公司 2030 模型：占 WFE 比重 +150bp）", False),
 ("服務、SPTS＋PCB：前一年 × 成長率（Assumed，錨在公司長期目標）。毛利＝FY2026 毛利＋增量毛利率（兩點推導 64.6%，公司說 60–65%）× 營收增量；營業費用 FY2027 取公司說法，之後 × 成長率；不寫死利益率", False),
 ("情境 A／B／C：設備市場模型的三情境（A 續增、B 2028 持平、C 2028 −20%），服務、兩事業與費用成長隨情境；中國出口管制只在 Sens", False),
 ("Val：P/E 規則——10 年（2016–2025 年底）當期 P/E 分位，見頂年 25 分位、中段中位數、谷底 75 分位；循環位置依設備股共用規則（各情境見頂年度讀設備市場模型：" + "、".join(f"{s_} {PEAK[s_]}" for s_ in SC) + " → FY2027：" + CP.fy2027_text(PEAK) + "）。第 0 節為情境加權目標價。本版未做先驗，無 Andy 合理 P/E 對照", False), ("", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。", False),
 ("已知限制：(1) 製程控制強度只有一個校準點（CY2026 錨點），強度上升率來自公司 2030 模型（Interested-party）；(2) 記憶體相對強度由公司「代工邏輯 low-to-mid 60s」一句推得；(3) 系統營收＝產品 − 兩事業（兩事業的服務未揭露）；(4) CY2027 起上下半各 1/2，忽略年內斜率；(5) 中國不在基準內另扣（設備模型 WFE 已含），出口管制只在 Sens；(6) 共識 EPS 為非 GAAP，模型為 GAAP；(7) 2026-10 股價由約 207 跌到 146，原因未查明。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

CFG_V2 = dict(cols=FYC, scn={"A": "順利", "B": "平平", "C": "不如預期"}, peers=[('應用材料（Applied Materials）', 'AMAT', 21.39, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('科林研發（Lam Research）', 'LRCX', 22.39, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('ASML（美國 ADR）', 'ASML', 24.74, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('東京威力科創（Tokyo Electron）', '8035.T', 27.09, '2026-10-21', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('Lasertec', '6920.T', 34.02, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）')], fwd={2028: "8/12", 2029: "4/12"}, fwd_note='約 2027-10 時的未來 12 月（2027-11～2028-10）≈ 8／12 × FY2028（2027-07～2028-06）＋4／12 × FY2029（會計年度 7–6 月，依月份調整）', a3='目標價（DCF 與 Comparable 加權）：見本頁第 {row} 列起「0. 目標價」（估值 v2 新增，放在既有內容之後以免位移既有列）；原「0. 目標價與上檔空間」改名為「舊法：規則值情境加權（參考）」',
    cur='US$',
    unit_rev='US$M',
    unit_ni='US$M',
    fmt='0.00',
    fmt_rev='#,##0',
    band_fmt='0.00',
    roe=('FY2024–FY2026', [0.8784, 1.0077, 0.875], 'stockanalysis 比率頁（2026-10-25 讀取）：FY2024 87.84%、FY2025 100.77%、FY2026 87.50%', '單一來源；高 ROE 來自大量回購使股東權益偏低'),
    nim=('FY2016–FY2025，GAAP', [(2.98, 0.70442), (3.48, 0.92608), (4.04, 0.80226), (4.57, 1.18), (5.81, 1.22), (6.92, 2.08), (9.21, 3.32), (10.5, 3.39), (9.81, 2.76), (12.16, 4.06)], 'vcpscanner.com 年度損益（US$B）', '第二來源：FY2025 與 Q 頁實績核對；其餘年度單一來源（FY2026 淨利率 35.6%，未納入）'),
    cons_tp=(232.38, 'stockanalysis：30 位分析師平均 US$232.38（中位數 228.50；區間 175–325），資料更新 2026-10-20（10 比 1 分割後）', '第二來源：未找到'),
    cons_feps=('=cons_eps28', '共識前瞻 EPS＝共識 FY2028 非 GAAP EPS（In 頁 cons_eps28，Yahoo 6.74）；本模型前瞻 EPS 為 8／12 FY2028＋4／12 FY2029（GAAP），約晚四個月', 'Derived'),
    peer_note='前段設備與量測檢測同業：應用材料、科林研發、ASML、東京威力科創、Lasertec',
    dcf_note='；會計年度 7–6 月，FY2027＝2026-07～2027-06 以 t＝1 折現（略低估）',
    pe10='E32:E41',
    pe10_note='現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（當期口徑自算）的分位；只作檢查，不進目標價',
    px_note='現價＝In 頁 px（2026-10-23 收盤）',
    method='DCF：股東現金流＝GAAP EPS ×（1 − g／ROE）（模型無折舊與資本支出項目），r 10%、g 4%、ROE 為近三年平均（回購使 ROE 偏高、再投資率偏低），FY2027–32 明確期（FY2030 起收斂法）＋永續。Comparable：應用材料、科林研發、ASML、東京威力科創、Lasertec 前瞻 P/E 中位數 × 12 個月後前瞻 EPS。偏誤：科磊製程控制份額高、歷史上享有溢價，用同業中位數偏低估。',
    readme=['v0.2（2026-10-25）估值 v2：Val 頁新增「0. 目標價」（放在既有內容之後，不位移既有列）＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，情境權重 A 0.3／B 0.5／C 0.2）＋0.5 × Comparable 目標價（同業前瞻 P/E 中位數 × 12 個月後前瞻 EPS）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解、歷史分位檢查。原「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」，不再叫目標價。', '預估期間延伸到 FY2032：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（PL 頁底部延伸區塊，灰底）。情境對應：A＝順利、B＝平平、C＝不如預期（原為顆數情境）。本版未做先驗，無 Andy 合理 P/E 對照。'])

# ---------------- 估值 v2（2026-10-25，Andy 11:18／11:20／11:28／11:53）：PL 延伸到 FY2032（收斂法）＋ Val「0. 目標價」（DCF 與 Comparable 加權） ----------------
# 為不位移既有儲存格（FY2026–FY2029 既有 EPS 逐格不變），新內容放在各頁既有內容之後；HTML 估值章第一個區塊讀「0.3 結論」
def valuation_v2(CFG):
    cols = CFG["cols"]; MF = CFG["fmt"]; YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
    FADEFILL = PatternFill("solid", fgColor="E7E6E6")
    # --- In：目標價與收斂法參數（黃底；Assumed，Andy 可改） ---
    r = IN.max_row + 2
    IN.cell(r, 1, "J. 目標價（DCF 與 Comparable 加權）與 FY2030–32 收斂法參數（v2 新增；黃底＝Assumed，Andy 可改）").font = BOLD; r += 1
    def inrow(key, item, unit, base, tag, s1, s2, fmt=None, yellow=False):
        nonlocal r
        IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
        c = IN.cell(r, 5, base)
        if not (isinstance(base, str) and base.startswith("=")): c.font = BLUE
        if fmt: c.number_format = fmt
        if yellow: c.fill = KEY
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
        wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
        r += 1
    inrow("tp_r", "折現率（股東權益要求報酬；美、台、日、韓同一值）", "%", 0.10, "Assumed", "Andy 可改（工作單 2026-10-25 預設 10%）", CFG.get("r_note", "—"), "0.0%", True)
    inrow("tp_g", "長期成長率（永續成長；收斂法 FY2032 營收年增的基準）", "%", 0.04, "Assumed", "Andy 可改（工作單預設 4%）", "情境 A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt（PL 頁延伸區塊）", "0.0%", True)
    inrow("tp_roe", f"ROE（近三年平均，{CFG['roe'][0]}）", "%", "=AVERAGE(" + ",".join(str(x) for x in CFG["roe"][1]) + ")", "Derived", CFG["roe"][2], CFG["roe"][3], "0.0%")
    inrow("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", 0.5, "Assumed", "Andy 可改（2026-10-25 11:53：DCF／Comparable 各半）", "—", "0%", True)
    if CFG["nim"]:
        inrow("nim10", f"十年平均淨利率（{CFG['nim'][0]}，各年淨利 ÷ 營收的平均）", "%", "=AVERAGE(" + ",".join(f"{n}/{d}" for d, n in CFG["nim"][1]) + ")", "Derived", CFG["nim"][2], CFG["nim"][3], "0.0%")
    else:
        IN.cell(r, 1, "nim10"); IN.cell(r, 2, "十年平均淨利率"); IN.cell(r, 3, "%"); IN.cell(r, 5, "未找到"); IN.cell(r, 7, "Assumed")
        IN.cell(r, 8, CFG["nim_missing"]); IN.cell(r, 9, "依工作單：找不到十年平均時用最後一年（FY2029）淨利率 × 0.8；PL 延伸區塊逐情境計算"); r += 1
    inrow("cons_tp", "市場共識目標價（賣方平均）", CFG["cur"], CFG["cons_tp"][0], "Interested-party", CFG["cons_tp"][1], CFG["cons_tp"][2], MF)
    inrow("cons_feps", "共識前瞻 EPS（差距拆解用）", CFG["cur"], CFG["cons_feps"][0], "Interested-party", CFG["cons_feps"][1], CFG["cons_feps"][2], MF)

    # --- PL：延伸到 FY2032（收斂法） ---
    p0 = PL.max_row + 3
    PL.cell(p0, 1, "延伸到 FY2032（v2 新增）：FY2025–FY2029 引用本頁頂端摘要；灰底 FY2030–FY2032 為收斂法，非驅動因子推導").font = BOLD
    PL.cell(p0 + 1, 1, "收斂法：營收年增由 FY2029 年增線性收斂到 FY2032 的長期成長（A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt）；淨利率由 FY2029 線性收斂（A＝FY2029、B＝FY2029 與十年平均的中點、C＝十年平均" + ("" if CFG["nim"] else "；十年平均未找到，以 FY2029 × 0.8 代替（Assumed）") + "）；股數維持 FY2029（EPS 與淨利同比例）")
    ext_cols = {y: chr(ord("C") + y - 2025) for y in range(2025, 2033)}
    hdr(PL, p0 + 2, ["情境（延伸）", "單位"] + [str(PL[f"{cols[y]}3"].value) for y in range(2025, 2030)] + ["FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
    EXT = {}; pr = p0 + 3
    gadd = {"A": "+0.02", "B": "", "C": "-0.02"}
    for s in SC:
        PL.cell(pr, 1, f"情境 {s}（{CFG['scn'][s]}）").font = BOLD; pr += 1
        R = {k: pr + i for i, k in enumerate(("rev", "g", "nim", "ni", "eps"))}
        for k, lab in (("rev", "營收"), ("g", "  營收年增"), ("nim", "  淨利率（淨利 ÷ 營收）"), ("ni", "淨利（收斂法延伸）"), ("eps", "EPS（延伸）")):
            PL.cell(R[k], 1, lab); PL.cell(R[k], 2, {"rev": CFG["unit_rev"], "ni": CFG["unit_ni"], "eps": CFG["cur"]}.get(k, "%"))
        for y in range(2025, 2033):
            c = ext_cols[y]; pc = ext_cols.get(y - 1)
            if y <= 2029:
                PL[f"{c}{R['rev']}"] = f"={cols[y]}{SUMR[(s, 'rev')]}"
                PL[f"{c}{R['ni']}"] = f"={cols[y]}{SUMR[(s, 'ni')]}"
                PL[f"{c}{R['eps']}"] = f"={cols[y]}{SUMR[(s, 'eps')]}"
                PL[f"{c}{R['nim']}"] = f"={c}{R['ni']}/{c}{R['rev']}"
                if y > 2025: PL[f"{c}{R['g']}"] = f"={c}{R['rev']}/{pc}{R['rev']}-1"
            else:
                k = y - 2029
                mT = {"A": f"$G{R['nim']}", "B": f"($G{R['nim']}+{'nim10' if CFG['nim'] else '0.8*$G' + str(R['nim'])})/2", "C": "nim10" if CFG["nim"] else f"0.8*$G{R['nim']}"}[s]
                PL[f"{c}{R['g']}"] = f"=$G{R['g']}+((tp_g{gadd[s]})-$G{R['g']})*{k}/3"
                PL[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
                PL[f"{c}{R['nim']}"] = f"=$G{R['nim']}+({mT}-$G{R['nim']})*{k}/3"
                PL[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
                PL[f"{c}{R['eps']}"] = f"=$G{R['eps']}*{c}{R['ni']}/$G{R['ni']}"
                for kk in R: PL[f"{c}{R[kk]}"].fill = FADEFILL
            for kk, fm in (("rev", CFG["fmt_rev"]), ("g", "0.0%"), ("nim", "0.0%"), ("ni", CFG["fmt_rev"]), ("eps", MF)):
                if PL[f"{c}{R[kk]}"].value is not None: PL[f"{c}{R[kk]}"].number_format = fm
        PL.cell(R["rev"], 11, "FY2030 起為收斂法，非驅動因子推導")
        EXT[s] = R; pr += 6
    for c in "IJ": PL.column_dimensions[c].width = 14
    PL.column_dimensions["K"].width = 40

    # --- Val：舊法改名（不再叫目標價） ---
    REN = (("0. 目標價與上檔空間", "舊法：規則值情境加權（參考）"), ("目標價（12 個月，情境加權）", "規則值情境加權價（12 個月；參考）"),
           ("上檔空間（目標價 ÷ 現價 − 1）", "相對現價（規則值情境加權價 ÷ 現價 − 1）"), ("目標價方法", "規則值方法"))
    for row in V.iter_rows(min_col=1, max_col=1):
        v = row[0].value
        if isinstance(v, str):
            for a, b in REN:
                if v.startswith(a): row[0].value = b + v[len(a):]; break
    # --- Val：0. 目標價（列標籤依工作單逐字；B／C／D＝A 順利／B 平平／C 不如預期） ---
    t0 = V.max_row + 3
    T = {}
    def tset(rr, key, label, fs, fmt, note=""):
        V.cell(rr, 1, label); T[key] = rr
        for c, f in zip("BCD", fs):
            if f is None: continue
            V[f"{c}{rr}"] = f; V[f"{c}{rr}"].number_format = fmt
        if note: V[f"E{rr}"] = note
    V.cell(t0, 1, "0. 目標價").font = BOLD; T["top"] = t0
    V.cell(t0 + 1, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD; T["h01"] = t0 + 1
    hdr(V, t0 + 2, ["項目", f"A（{CFG['scn']['A']}）", f"B（{CFG['scn']['B']}）", f"C（{CFG['scn']['C']}）", "說明"]); T["hd1"] = t0 + 2
    V.cell(t0 + 3, 1, "情境 EPS（FY2027–FY2032）"); V[f"E{t0 + 3}"] = "PL 頁延伸區塊；FY2030 起為收斂法"; T["eps_h"] = t0 + 3
    EPSR = {}
    for i, y in enumerate(YRS6):
        tset(t0 + 4 + i, f"eps{y}", f"  FY{y}", [f"=PL!{ext_cols[y]}{EXT[s]['eps']}" for s in SC], MF); EPSR[y] = t0 + 4 + i
    rr = t0 + 10
    V.cell(rr, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）"); V[f"E{rr}"] = "模型無折舊與資本支出項目，故不用（淨利＋折舊－資本支出）；設備股資本支出低（輕資產），此式偏誤小"; T["cf_h"] = rr
    tset(rr + 1, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）")
    RV = rr + 1; CFR = {}
    for i, y in enumerate(YRS6):
        tset(rr + 2 + i, f"cf{y}", f"  FY{y}", [f"={c}{EPSR[y]}*(1-$B${RV})" for c in "BCD"], MF); CFR[y] = rr + 2 + i
    c1, c6 = CFR[2027], CFR[2032]
    tset(rr + 8, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=SUMPRODUCT({c}{c1}:{c}{c6},1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}{c6}*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6" for c in "BCD"], MF,
         "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底，t＝1 對 FY2027" + CFG["dcf_note"])
    tset(rr + 9, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}{rr + 8}*(1+tp_r)" for c in "BCD"], MF)
    tset(rr + 10, "dcftp", "DCF 目標價（情境加權）", [f"=tp_wA*B{rr + 9}+tp_wB*C{rr + 9}+tp_wC*D{rr + 9}", None, None], MF, "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁 tp_wA／tp_wB／tp_wC")
    q = rr + 12
    V.cell(q, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD; T["h02"] = q
    hdr(V, q + 1, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"]); T["phd"] = q + 1
    for i, (nm, tk, pe, dt, src) in enumerate(CFG["peers"], q + 2):
        V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
        V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
    pl_ = q + 1 + len(CFG["peers"]); T["p1"] = q + 2; T["pl"] = pl_
    q = pl_ + 1
    tset(q, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C{T['p1']}:C{pl_})", None, None], "0.0", CFG["peer_note"])
    fw = CFG["fwd"]
    def feps(c): return "(" + "+".join(f"{w}*{c}{EPSR[y]}" for y, w in fw.items()) + ")"
    tset(q + 1, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=tp_wA*" + feps("B") + "+tp_wB*" + feps("C") + "+tp_wC*" + feps("D"), None, None], MF, CFG["fwd_note"])
    tset(q + 2, "comptp", "Comparable 目標價", [f"=B{q}*B{q + 1}", None, None], MF, "同業中位數 × 上列 EPS")
    q += 4
    V.cell(q, 1, "0.3 結論").font = BOLD; T["h03"] = q
    hdr(V, q + 1, ["項目", "數值", "", "", "說明"]); T["hd3"] = q + 1
    q += 2
    tset(q, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B{T['comptp']}", None, None], MF, "tp_w_dcf × DCF 目標價＋（1 − tp_w_dcf）× Comparable 目標價；tp_w_dcf 見 In 頁")
    tset(q + 1, "comp2", "Comparable 目標價（市場主流同業法）", [f"=B{T['comptp']}", None, None], MF, "引用 0.2")
    tset(q + 2, "cons", "市場共識目標價", ["=cons_tp", None, None], MF, CFG["cons_tp"][1])
    tset(q + 3, "upw", "上檔空間：加權目標價", [f"=B{q}/px-1", None, None], "0.0%", CFG["px_note"])
    tset(q + 4, "upc", "上檔空間：Comparable 目標價", [f"=B{q + 1}/px-1", None, None], "0.0%")
    tset(q + 5, "geps", "與市場共識目標價的差距：EPS 因素", [f"=B{T['feps']}/cons_feps-1", None, None], "0.0%", "本模型 12 個月後前瞻 EPS ÷ 共識前瞻 EPS − 1；" + CFG["cons_feps"][1])
    tset(q + 6, "gmul", "與市場共識目標價的差距：倍數因素", [f"=(cons_tp/cons_feps)/B{T['pmed']}-1", None, None], "0.0%")
    V[f"E{q + 6}"] = f'="共識目標價 ÷ 共識前瞻 EPS 隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B{T["pmed"]},"0.0")&"x"'
    tset(q + 7, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D{T["dcf12"]},"{CFG["band_fmt"]}")&"～"&TEXT(B{T["dcf12"]},"{CFG["band_fmt"]}")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
    PE = CFG["pe10"]
    tset(q + 8, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）",
         [f"=IF(px/{c}{EPSR[2028]}<MIN({PE}),0,IF(px/{c}{EPSR[2028]}>MAX({PE}),1,PERCENTRANK({PE},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%", CFG["pe10_note"])
    V.cell(q + 9, 1, "目標價方法"); V[f"B{q + 9}"] = CFG["method"]; T["meth"] = q + 9
    for k in ("dcftp", "comptp", "tpw"): V[f"B{T[k]}"].fill = KEY
    if CFG.get("a3"): V["A3"] = CFG["a3"].format(row=t0)

    # --- Sens：目標價參數 ---
    s0 = SE.max_row + 3
    SE.cell(s0, 1, "目標價參數（v2 新增；只改 DCF 參數與權重，PL 收斂法的長期成長不隨 tp_g 變動）").font = BOLD
    hdr(SE, s0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
    def dcf_formula(c, r_="tp_r", g="tp_g"):
        cf = lambda y: f"Val!{c}{EPSR[y]}*(1-MIN({g}/tp_roe,0.8))"
        pv = "+".join(f"{cf(y)}/(1+{r_})^{t}" for t, y in enumerate(YRS6, 1))
        return f"({pv}+{cf(2032)}*(1+{g})/({r_}-{g})/(1+{r_})^6)"
    cases = [("基準", "r 10%、g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "r 11%", dict(r_="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", dict(r_="(tp_r-0.01)")),
             ("長期成長 +1pt", "g 5%（永續與再投資率）", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", dict(g="(tp_g-0.01)")),
             ("DCF 權重 0.4", "Comparable 0.6", dict(w="0.4")), ("DCF 權重 0.8", "Comparable 0.2", dict(w="0.8"))]
    for k, (nm, st, kw) in enumerate(cases):
        i = s0 + 2 + k; w = kw.pop("w", "tp_w_dcf")
        SE.cell(i, 1, nm); SE.cell(i, 2, st)
        rr_ = kw.get("r_", "tp_r")
        SE[f"C{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_formula(c, **kw)}" for s, c in zip(SC, "BCD")) + f")*(1+{rr_})"
        SE[f"D{i}"] = f"=Val!B{T['comptp']}"
        SE[f"E{i}"] = f"={w}*C{i}+(1-{w})*D{i}"
        SE[f"F{i}"] = f"=E{i}/E${s0 + 2}-1"
        for c in "CDE": SE[f"{c}{i}"].number_format = MF
        SE[f"F{i}"].number_format = "0.0%"
    T["sens0"] = s0; T["sens_last"] = s0 + 1 + len(cases)

    # --- README 頁 ---
    rr = RD.max_row + 1
    for t in CFG["readme"]:
        RD.cell(rr, 1, t); rr += 1
    return T, EXT

V2T, V2EXT = valuation_v2(CFG_V2)
if os.environ.get("V2_ROWS"):
    print("V2", V2T, V2EXT)

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(sys.argv[1])

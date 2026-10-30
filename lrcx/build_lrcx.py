# Lam Research（LRCX）收支與估值模型 v0.2.1（AI 半導體 Project 步驟 5：設備與材料）
# 用法：python3 lrcx/build_lrcx.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度 7 月到隔年 6 月（以 6 月最後一個星期日結束）：FY2027＝2026-07～2027-06。金額：十億美元（US$B）；EPS 美元。
# 樣板：advantest/（工作表結構、三情境、Val 列標籤、設備股共用循環規則）。
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

# ---------------- 讀入設備市場模型成品（唯讀；不修改） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
EQ_F = only("markets/equipment/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
ws_ = openpyxl.load_workbook(EQ_F, data_only=True)["Mkt"]
EQ = {}      # (情境, 分項) -> ([CY2025..CY2029], 來源)
LABS = {"先進邏輯（台積電設備＋其他代工中屬先進者）〔AI〕": "adv", "其他代工與邏輯": "ofl", "DRAM（含 HBM）〔AI〕": "dram", "NAND": "nand", "晶圓前段設備": "wfe"}
cur = None
for r in ws_.iter_rows():
    lab = r[0].value
    if isinstance(lab, str) and lab.startswith("5A."): cur = "A"
    elif isinstance(lab, str) and lab.startswith("5B."): cur = "B"
    elif isinstance(lab, str) and lab.startswith("5C."): cur = "C"
    elif isinstance(lab, str) and lab.startswith("6."): cur = None
    if cur and lab in LABS:
        EQ[(cur, LABS[lab])] = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert len(EQ) == 15 and all(v is not None for k in EQ for v in EQ[k][0]), sorted(EQ)
for k in ("adv", "ofl", "dram", "nand"):   # 2025–2027 三情境相同（設備模型的設計）
    for j in range(3):
        assert abs(EQ[("A", k)][0][j] - EQ[("C", k)][0][j]) < 1e-9
PEAK, PEAK_SRC = CP.peak_years(ROOT)

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億美元（US$B）；EPS 美元。每個輸入附兩個來源（來源2 找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司財報（8-K 新聞稿、10-K、10-Q）取自 SEC EDGAR 原文；法說內容取自逐字稿摘要（Motley Fool、MarketBeat、TIKR）"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)

SEGN = {"fl": "代工與邏輯（先進邏輯＋其他代工與邏輯）", "dram": "DRAM（含 HBM）", "nand": "NAND"}
sec("A. 晶圓前段設備市場（讀 markets/equipment 成品 Mkt 頁 5A／5B／5C；三情境；日曆年 US$B；唯讀，不另估）")
for s in SC:
    for k, nm in (("adv", "先進邏輯〔AI〕"), ("ofl", "其他代工與邏輯"), ("dram", "DRAM（含 HBM）〔AI〕"), ("nand", "NAND"), ("wfe", "晶圓前段設備合計")):
        for j, y in enumerate(range(2025, 2030)):
            inp(f"{k}{s}_{y % 100}", f"{nm} CY{y}（情境 {s}）", "US$B", None, EQ[(s, k)][0][j], None, "Derived", EQ[(s, k)][1] if j == 0 else "同上",
                "markets/equipment：公司資本支出 × 設備占比＋其他廠；2028 起 B、C 乘情境係數（Assumed）" if j == 0 else "")
for s in SC:
    inp(f"pk_{s}", f"情境 {s}：設備市場見頂年度（日曆年；循環位置規則用）", "年", None, PEAK[s], None, "Derived", PEAK_SRC[s],
        "設備股共用規則（tools/cycle_position.py）：FY(t) 早於見頂年＝中段、等於＝見頂、晚於＝谷底" if s == "A" else "同上")

sec("B. 系統營收的情境與份額參數")
inp("h1s", "上半年（1–6 月）占日曆年系統營收與 CSBG 的比例（CY→FY 轉換用）", "%", 0.45, 0.50, 0.55, "Assumed", "無：FY(t)＝CY(t−1) 下半＋CY(t) 上半；基準各半", "CY2026 實績＋指引：上半 8.0、下半 11.8（系統）——成長年下半較高，故區間 45–55%")
inp("drift", "相對 WFE 分項的份額漂移（每年，CY2027 起）", "%/年", -0.02, 0.0, 0.02, "Assumed", "無：基準為 0（CY2026 的份額錨點延續）",
    "公司稱 CY2026 為連續第三年優於 WFE、SAM 占 WFE 比重上升約 4 個百分點（Interested-party）；不放進基準，只在 Sens 列 ±2%")
inp("cn_sh", "中國占系統營收比重（以 FY2026 中國占總營收代用）", "%", 0.26, 0.34, 0.43, "Verified", "Lam FY2026 10-K：中國占營收 34%（FY2025 34%）",
    "季別：FY26 Q1 43%、Q2 35%、Q3 34%、Q4 26%、FY27 Q1 35%（8-K）；低／高＝季別最小／最大。系統營收的中國比重未揭露（Assumed 同總營收）")
inp("cutC", "情境 C：出口管制加嚴，中國系統營收折損（CY2028 起）", "%", 0.0, 0.20, 0.40, "Assumed",
    "Analogy：2022-10 美國出口管制，公司估 2023 年營收影響 US$2.0–2.5B（techwireasia 轉述；當時中國約占營收 30%），約為中國營收的三到四成",
    "基準 20%＝類比區間的一半；高 40%＝類比上緣。CSBG（含 Reliant 成熟製程設備）未套用")
inp("cut_sens", "Sens：中國系統營收折損（情境 A，CY2027 起）", "%", 0.20, 0.0, 0.40, "Assumed", "同上：基準 0（現行管制已反映在 CY2026 錨點）", "低欄＝20%、高欄＝40%（Sens 用）")
inp("csq2_g", "FY2027 Q2 CSBG 相對 Q1 的變化（拆分指引營收用）", "%", -0.05, 0.0, 0.10, "Assumed", "無：公司只給總營收指引 8.6",
    "結果檢查：CY2026 CSBG 年增＝33%（公司稱逾 30%，Recon）")

sec("C. 客戶支援事業群 CSBG（服務、備件、升級、Reliant 成熟製程設備）")
inp("g_cs27", "CSBG CY2027 年增", "%", 0.10, 0.15, 0.20, "Interested-party", "Lam FQ1-27 法說（2026-10-21）：CY2027 成長逾 15%（Motley Fool 逐字稿）", "第二來源：未找到。低／高 Assumed")
for s, g in (("A", 0.12), ("B", 0.06), ("C", -0.03)):
    inp(f"g_cs_{s}", f"CSBG CY2028–29 年增（{s}）", "%", g - 0.04, g, g + 0.04, "Assumed",
        "無：裝機量（公司稱逾 10 萬腔體）持續增加、記憶體稼動率隨情境；C 含 NAND 升級潮結束（公司稱 2027 年結束）", "—")

sec("D. 毛利率與營業費用（兩點推導：FY2026 實績與 FY2027 上半〔Q1 實績＋Q2 指引〕年化；不寫死利益率）")
inp("rev_h1", "FY2027 上半營收年化（Q1 實績＋Q2 指引）× 2", "US$B", None, "=2*(q1_rev+g_q2rev)", None, "Derived", "Q 頁", "—")
inp("gm_h1", "FY2027 上半毛利率（Q1 實績＋Q2 指引 52.5%）", "%", None, "=(q1_gp+g_q2gm*g_q2rev)/(q1_rev+g_q2rev)", None, "Derived", "Q 頁", "—")
inp("gm_b", "毛利率斜率（每單位 ln 營收）", "pt", None, "=(gm_h1-a26_gp/a26_rev)/(LN(rev_h1)-LN(a26_rev))", None, "Derived",
    "兩點：FY2026（營收 23.2、毛利率 50.5%）、FY2027 上半年化（33.6、52.4%）", "對數形式：規模效益遞減；FY2025→FY2026 的斜率較大（7.8），未採")
inp("gm_a", "毛利率截距", "%", None, "=a26_gp/a26_rev-gm_b*LN(a26_rev)", None, "Derived", "同上", "公司長期目標毛利率 mid-50s（Interested-party）；對照見 Recon")
inp("opx_h1", "FY2027 上半營業費用年化（Q1 實績＋Q2 指引〔毛利率 − 營益率〕× 營收）× 2", "US$B", None, "=2*(q1_opx+(g_q2gm-g_q2om)*g_q2rev)", None, "Derived", "Q 頁", "—")
inp("opex_var", "營業費用變動部分占營收（兩點斜率）", "%", None, "=(opx_h1-a26_opx)/(rev_h1-a26_rev)", None, "Derived", "（上半年化 − FY2026）÷（營收差）", "—")
inp("opex_fix", "營業費用固定部分", "US$B", None, "=a26_opx-opex_var*a26_rev", None, "Derived", "FY2026 營業費用 − 變動 × 營收", "—")
inp("phi", "營業費用中變動部分占比（FY2026；Sens 用）", "%", 0.15, "=opex_var*a26_rev/a26_opx", 0.50, "Derived", "基準＝兩點推導值；低／高 Assumed", "—")

sec("E. 業外、稅率、股數")
inp("oth_inc", "業外收益淨額（每年）", "US$B", 0.0, 0.10, 0.20, "Assumed", "FY2026 0.063、FY2025 0.057、FY27 Q1 0.028（8-K）", "含遞延薪酬資產評價，波動大")
inp("tax", "有效稅率", "%", 0.13, 0.165, 0.19, "Interested-party", "Lam FQ1-27 法說：非 GAAP 稅率 16.5%，12 月季「及之後」預期在此範圍（Motley Fool 逐字稿）",
    "GAAP：FY27 Q1 16.9%（559.1÷3,312.2，8-K）；FY2026 12.1%、FY2025 10.1%（偏低，不當基準）")
inp("sh_27", "FY2027 稀釋股數", "十億股", None, 1.253, None, "Interested-party", "公司 12 月季指引：稀釋股數 1.253B（8-K 2026-10-21）", "FY27 Q1 實績 1.2558B（10-Q）")
inp("bb", "FY2028 起每年股數淨減少（庫藏股）", "%/年", 0.0, 0.015, 0.03, "Assumed", "FY2025→FY2026 稀釋股數 1.2901→1.2611（−2.2%，10-K）", "公司長期目標：自由現金流 85% 回饋股東（Interested-party）")

sec("F. 公司指引（FY2027 Q2＝2026-10～12 月季；2026-10-21 公布）、法說說法、共識與股價")
inp("g_q2rev", "指引營收 FY27 Q2", "US$B", 8.2, 8.6, 9.0, "Interested-party", "Lam 8-K 新聞稿（2026-10-21）：8.60 ± 0.40", "Motley Fool 逐字稿同")
inp("g_q2gm", "指引毛利率 FY27 Q2", "%", 0.515, 0.525, 0.535, "Interested-party", "同上：52.5% ± 1%", "同上")
inp("g_q2om", "指引營業利益率 FY27 Q2", "%", 0.395, 0.405, 0.415, "Interested-party", "同上：40.5% ± 1%", "同上")
inp("g_q2eps", "指引稀釋 EPS FY27 Q2", "US$", 2.18, 2.33, 2.48, "Interested-party", "同上：2.33 ± 0.15", "同上")
inp("co_wfe26", "公司估 CY2026 WFE（mid-$150B）", "US$B", None, 155, None, "Interested-party", "Lam FQ1-27 法說：mid-$150B（年初 mid-$130B）", "GuruFocus／BigGo 法說摘要；公司 WFE 口徑與 SEMI 不同")
inp("co_wfe25", "公司估 CY2025 WFE（約 $110B）", "US$B", None, 110, None, "Interested-party", "Lam FQ2-26 法說（2026-01-28）：約 110（TIKR 摘要）", "第二來源：未找到")
inp("co_wfe27", "公司初步看 CY2027 WFE（逾 $200B）", "US$B", None, 200, None, "Interested-party", "Lam FQ1-27 法說：more than $200B；DRAM 成長最多，其次代工邏輯、NAND", "GuruFocus 法說摘要")
inp("co_sys26", "公司：CY2026 系統營收年增（well over 40%）", "%", None, 0.40, None, "Interested-party", "Lam FQ1-27 法說（含 12 月季指引中點）", "下限值；Recon 用")
inp("co_cs26", "公司：CY2026 CSBG 年增（逾 30%）", "%", None, 0.30, None, "Interested-party", "Lam FQ1-27 法說", "下限值；Recon 用")
inp("cons_eps27", "共識 EPS FY2027（Yahoo，29 位分析師）", "US$", None, 9.63, None, "Interested-party", "Yahoo Finance analysis（2026-10-23 收盤後頁面）", "stockanalysis 9.48（30 位，2026-10-22；非 GAAP）")
inp("cons_eps27b", "共識 EPS FY2027（stockanalysis）", "US$", None, 9.48, None, "Interested-party", "stockanalysis forecast（2026-10-22 更新）", "—")
inp("cons_rev27", "共識營收 FY2027（Yahoo）", "US$B", None, 35.39, None, "Interested-party", "Yahoo Finance analysis", "stockanalysis 34.9")
inp("cons_eps28", "共識 EPS FY2028（Yahoo，28 位）", "US$", None, 12.15, None, "Interested-party", "Yahoo Finance analysis（低 10.41／高 15.12）", "第二來源：未找到（stockanalysis 需付費）")
inp("cons_rev28", "共識營收 FY2028（Yahoo）", "US$B", None, 42.68, None, "Interested-party", "Yahoo Finance analysis", "第二來源：未找到")
inp("px", "股價（stockanalysis 2026-10-23 收盤）", "US$", None, 230.10, None, "Interested-party", "stockanalysis.com（2026-10-23 16:00 EDT）", "digrin 月資料 2026-10：230.10")
inp("sa_fpe", "遠期 P/E（stockanalysis）", "x", None, 22.39, None, "Interested-party", "stockanalysis.com", "自算見 Val")
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}", "%", None, w, None, "Assumed", "Andy 可改", "三者合計應為 100%（Val 第 0 節檢查）")

r = 4
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
    r += 1
for col, w in zip("ABCDEFGHI", [12, 64, 10, 9, 11, 9, 18, 80, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
INROW = {}
for rr in range(4, r):
    if IN.cell(rr, 3).value is not None: INROW[IN.cell(rr, 1).value] = rr

def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績（FY2025 全年、FY2026 四季與全年、FY2027 Q1、Q2 指引） ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績（十億美元；EPS 美元；GAAP）。會計年度 7–6 月：FY2026＝2025-07～2026-06。藍字＝輸入；黑字＝公式（Derived）。來源：Lam 8-K 新聞稿（Exhibit 99.1）、10-K、10-Q（SEC EDGAR 原文）"
Q["A2"] = "FY2027 第一季（2026-07～09）已於 2026-10-21 公布並納入（G 欄）；H 欄為 12 月季指引（2026-10-21）。下一次財報：FY2027 第二季，預計 2027-01 下旬（公司尚未公告日期）"
hdr(Q, 3, ["項目", "FY2025 全年", "FY26 Q1（7–9月）", "FY26 Q2（10–12月）", "FY26 Q3（1–3月）", "FY26 Q4（4–6月）", "FY27 Q1（7–9月）", "FY27 Q2（指引）", "FY2026 全年", "標記", "來源1（公司）", "來源2／說明"])
QR = {}
def qrow(r, key, label, vals, tag, s1, s2, fmt="#,##0.000"):
    QR[key] = r
    Q.cell(r, 1, label)
    for j, v in zip("BCDEFGHI", vals):
        if v is None: continue
        c = Q[f"{j}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    Q.cell(r, 10, tag); Q.cell(r, 11, s1); Q.cell(r, 12, s2)
K10 = "FY2026 10-K（2026-08-07）"; PR = "各季 8-K 新聞稿（2025-07-30、2025-10-22、2026-01-28、2026-04-22、2026-07-29、2026-10-21）"
qrow(4, "sys", "系統營收（新的先進製程設備）", [11.491280, 3.547565, 3.357493, 3.730582, 4.249848, 5.732907, "=H6-H5", 14.885488], "Verified", PR + "：Systems revenue", K10 + "：FY2026 14,885.5、FY2025 11,491.3（百萬美元）；四季合計＝全年（見檢查列）")
qrow(5, "cs", "客戶支援與其他（CSBG：服務、備件、升級、Reliant 成熟製程設備）", [6.944311, 1.776608, 1.987298, 2.110906, 2.472390, 2.486511, "=G5*(1+csq2_g)", 8.347202], "Verified", PR + "：Customer support-related revenue and other", K10 + "；Q2 指引欄＝Q1 × (1＋csq2_g)（Assumed）")
qrow(6, "rev", "營收", [18.435591, 5.324173, 5.344791, 5.841488, 6.722238, 8.219418, "=g_q2rev", 23.232690], "Verified", PR, K10 + "：23,232.7、18,435.6")
qrow(7, "chk1", "檢查：系統＋CSBG − 營收（應為 0）", [f"={c}4+{c}5-{c}6" for c in "BCDEFGHI"], "—", "—", "—", "0.000")
qrow(8, "gp", "毛利", [8.979059, 2.684879, 2.651162, 2.910527, 3.478740, 4.298853, "=H6*g_q2gm", 11.725308], "Verified", PR, K10)
qrow(9, "gmr", "毛利率", [f"={c}8/{c}6" for c in "BCDEFGHI"], "Derived", "—", "公司：FY2026 50.5%、FY27 Q1 52.3%（GAAP）", "0.0%")
qrow(10, "rd", "研發費用", [2.096387, 0.576446, 0.573305, 0.583200, 0.642922, 0.666398, None, 2.375873], "Verified", PR, K10)
qrow(11, "sga", "銷管費用", [0.981704, 0.279345, 0.267654, 0.280311, 0.322330, 0.348728, None, 1.149640], "Verified", PR, K10)
qrow(12, "opx", "營業費用", ["=B10+B11", "=C10+C11", "=D10+D11", "=E10+E11", "=F10+F11", "=G10+G11", "=H8-H13", "=I10+I11"], "Derived", "—", "Q2 指引欄＝毛利 − 營業利益")
qrow(13, "oi", "營業利益", [5.900968, 1.829088, 1.810203, 2.047016, 2.513488, 3.283727, "=H6*g_q2om", 8.199795], "Verified", PR, K10)
qrow(14, "oim", "營業利益率", [f"={c}13/{c}6" for c in "BCDEFGHI"], "Derived", "—", "公司：FY2026 35.3%、FY27 Q1 40.0%", "0.0%")
qrow(15, "oth", "業外收益淨額", [0.057161, 0.030074, 0.026410, -0.035460, 0.041654, 0.028493, None, 0.062678], "Verified", PR, K10)
qrow(16, "tx", "所得稅", [0.599912, 0.290502, 0.242619, 0.186096, 0.277860, 0.559082, None, 0.997077], "Verified", PR, K10)
qrow(17, "ni", "淨利", [5.358217, 1.568660, 1.593994, 1.825460, 2.277282, 2.753138, "=H18*H19", 7.265396], "Verified", PR, K10 + "；Q2 指引欄＝指引 EPS × 指引股數（Derived）")
qrow(18, "eps", "EPS（稀釋，美元）", [4.15, 1.24, 1.26, 1.45, 1.81, 2.19, "=g_q2eps", 5.76], "Verified", PR, K10 + "；非 GAAP FY2026 5.82（各季 1.26＋1.27＋1.47＋1.82）", "0.00")
qrow(19, "dsh", "稀釋股數（十億股）", [1.290142, 1.269313, 1.261739, 1.257325, 1.256032, 1.255768, "=sh_27", 1.261102], "Verified", PR, K10, "0.0000")
qrow(20, "epsb", "EPS（基本，美元；Val 的十年 P/E 用）", [4.17, 1.24, 1.27, 1.46, 1.82, 2.20, None, 5.79], "Verified", PR, K10, "0.00")
qrow(21, "ck2", "檢查：四季合計 − 全年（系統、CSBG、營收、毛利、營業利益、淨利依序在 B–G 欄）",
     ["=SUM(C4:F4)-I4", "=SUM(C5:F5)-I5", "=SUM(C6:F6)-I6", "=SUM(C8:F8)-I8", "=SUM(C13:F13)-I13", "=SUM(C17:F17)-I17"], "—", "—", "—", "0.000")
Q.cell(23, 1, "地區與客戶別（占比）").font = BOLD
qrow(24, "cn", "中國占營收", [0.34, 0.43, 0.35, 0.34, 0.26, 0.35, None, 0.34], "Verified", PR + "：Region revenue", K10 + "：FY2026、FY2025 皆 34%；管理層稱 12 月季中國占比將「明顯下降」（FQ1-27 法說）", "0%")
qrow(25, "pf", "代工（Foundry）占系統營收", [0.45, 0.60, 0.59, 0.54, 0.44, 0.46, "=G25", 0.54], "Verified", "各季 10-Q 與 FY2026 10-K：設備與升級營收的客戶別占比", "法說稱占「系統營收」，數字與 10-Q 相同；Q2 指引欄沿用 Q1（Assumed）", "0%")
qrow(26, "pl_", "邏輯／IDM 占系統營收", [0.13, 0.06, 0.07, 0.07, 0.10, 0.08, "=G26", 0.07], "Verified", "同上", "同上", "0%")
qrow(27, "pd", "DRAM 占系統營收", [None, 0.16, 0.23, 0.27, 0.23, 0.20, "=G27", None], "Interested-party", "法說逐字稿摘要（TIKR FQ2-26；MarketBeat FQ4-26；Motley Fool FQ1-27）：Q1 16%、Q2 23%、Q3 27%、Q4 23%、FY27 Q1 20%",
     "10-Q 只揭露記憶體合計；DRAM＋NAND＝記憶體（見檢查列）。FY2025 全年 DRAM／NAND 拆分：未找到", "0%")
qrow(28, "pn", "NAND（非揮發性記憶體）占系統營收", [None, 0.18, 0.11, 0.12, 0.23, 0.26, "=G28", None], "Interested-party", "同上：Q1 18%、Q2 11%、Q3 12%、Q4 23%、FY27 Q1 26%", "同上", "0%")
qrow(29, "pm", "記憶體占系統營收（10-Q）", [0.42, 0.34, 0.34, 0.39, 0.46, 0.46, "=G29", 0.39], "Verified", "10-Q／10-K：Memory", "—", "0%")
qrow(30, "ck3", "檢查：代工＋邏輯＋記憶體 − 1；DRAM＋NAND − 記憶體（C–H 欄）", ["=B25+B26+B29-1", "=C27+C28-C29", "=D27+D28-D29", "=E27+E28-E29", "=F27+F28-F29", "=G27+G28-G29", "=H27+H28-H29", "=I25+I26+I29-1"], "—", "—", "B、I 欄為三項合計 − 1", "0.00")
Q.cell(32, 1, "CY2025 對照用：FY2025 下半（2025-01～06）").font = BOLD
hdr(Q, 33, ["項目", "FY25 Q3（1–3月）", "FY25 Q4（4–6月）", "", "", "", "", "", "", "標記", "來源1（公司）", "來源2／說明"])
for i, (nm, a, b) in enumerate((("系統營收", 3.035276, 3.437625), ("CSBG", 1.684899, 1.733768), ("營收", 4.720175, 5.171393)), 34):
    Q.cell(i, 1, nm); Q.cell(i, 2, a).font = BLUE; Q.cell(i, 3, b).font = BLUE
    for c in "BC": Q[f"{c}{i}"].number_format = "#,##0.000"
    Q.cell(i, 10, "Verified"); Q.cell(i, 11, "8-K 2025-04-23、2025-07-30；FY26 Q3 與 Q1 新聞稿的前期比較欄"); Q.cell(i, 12, "—")
for nm, rr in (("sys", 4), ("cs", 5), ("rev", 6), ("gp", 8), ("opx", 12), ("oi", 13), ("ni", 17), ("eps", 18), ("dsh", 19), ("epsb", 20)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$I${rr}"); defname(f"q1_{nm}", f"Q!$G${rr}")
Q.column_dimensions["A"].width = 62
for c in "BCDEFGHI": Q.column_dimensions[c].width = 14
Q.column_dimensions["J"].width = 18; Q.column_dimensions["K"].width = 80; Q.column_dimensions["L"].width = 80
Q.freeze_panes = "B4"

# ---------------- Units：CY2026 系統營收錨點（客戶別）與對 WFE 分項的比例 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 日曆年系統營收依客戶別的錨點：CY2026＝FY26 Q3＋Q4（實績）＋FY27 Q1（實績）＋FY27 Q2（指引）；各季系統營收 × 客戶別占比；再除以設備市場模型的 CY2026 分項，得到份額係數 k"
hdr(UN, 3, ["項目", "單位", "FY26 Q3", "FY26 Q4", "FY27 Q1", "FY27 Q2（指引）", "CY2026", "說明"])
QC = {"FY26 Q3": "E", "FY26 Q4": "F", "FY27 Q1": "G", "FY27 Q2": "H"}
UR = {}
def urow(r, key, lab, unit, fs, total, fmt="0.000", note=""):
    UN.cell(r, 1, lab); UN.cell(r, 2, unit)
    for j, f in enumerate(fs):
        if f is None: continue
        UN.cell(r, 3 + j, f).number_format = fmt
    if total is not None: UN.cell(r, 7, total).number_format = fmt
    UN.cell(r, 8, note); UR[key] = r
urow(4, "sys", "系統營收", "US$B", [f"=Q!{c}4" for c in QC.values()], "=SUM(C4:F4)", note="Q2 指引欄＝指引營收 − CSBG 推估")
urow(5, "fl", "  代工與邏輯", "US$B", [f"=Q!{c}4*(Q!{c}25+Q!{c}26)" for c in QC.values()], "=SUM(C5:F5)")
urow(6, "dram", "  DRAM", "US$B", [f"=Q!{c}4*Q!{c}27" for c in QC.values()], "=SUM(C6:F6)")
urow(7, "nand", "  NAND", "US$B", [f"=Q!{c}4*Q!{c}28" for c in QC.values()], "=SUM(C7:F7)")
urow(8, "chk", "  檢查：三項 − 系統營收（應為 0）", "US$B", ["=C5+C6+C7-C4", "=D5+D6+D7-D4", "=E5+E6+E7-E4", "=F5+F6+F7-F4"], "=G5+G6+G7-G4")
urow(9, "cs", "CSBG", "US$B", [f"=Q!{c}5" for c in QC.values()], "=SUM(C9:F9)")
for k in ("sys", "fl", "dram", "nand", "cs"): defname(f"cy26_{k}", f"Units!$G${UR[k]}")
defname("h2_26_sys", "Units!$E$4:$F$4"); defname("h2_26_cs", "Units!$E$9:$F$9")
UN.cell(11, 1, "CY2025 與 CY2026 年增（對照公司說法）").font = BOLD
UN.cell(12, 1, "CY2025 系統營收（FY25 Q3＋Q4＋FY26 Q1＋Q2）"); UN.cell(12, 2, "US$B"); UN["G12"] = "=Q!B34+Q!C34+Q!C4+Q!D4"
UN.cell(13, 1, "CY2025 CSBG"); UN.cell(13, 2, "US$B"); UN["G13"] = "=Q!B35+Q!C35+Q!C5+Q!D5"
UN.cell(14, 1, "CY2026 系統營收年增"); UN.cell(14, 2, "%"); UN["G14"] = "=G4/G12-1"; UN.cell(14, 8, "公司：well over 40%（含 12 月季指引中點）")
UN.cell(15, 1, "CY2026 CSBG 年增"); UN.cell(15, 2, "%"); UN["G15"] = "=G9/G13-1"; UN.cell(15, 8, "公司：逾 30%")
UN.cell(16, 1, "CY2026 下半系統營收（FY27 Q1＋Q2）"); UN.cell(16, 2, "US$B"); UN["G16"] = "=E4+F4"
UN.cell(17, 1, "CY2026 下半 CSBG"); UN.cell(17, 2, "US$B"); UN["G17"] = "=E9+F9"
for rr, f in ((12, "0.000"), (13, "0.000"), (14, "0.0%"), (15, "0.0%"), (16, "0.000"), (17, "0.000")): UN[f"G{rr}"].number_format = f
defname("cy25_sys", "Units!$G$12"); defname("cy25_cs", "Units!$G$13"); defname("h2s26", "Units!$G$16"); defname("h2c26", "Units!$G$17")
UN.cell(19, 1, "份額係數 k＝CY2026 Lam 系統營收（客戶別）÷ 設備市場模型 CY2026 分項（三情境 CY2026 相同）").font = BOLD
hdr(UN, 20, ["分項", "單位", "Lam CY2026", "WFE 分項 CY2026", "k", "", "", "說明"])
for i, (k, nm, wf) in enumerate((("fl", "代工與邏輯", "advA_26+oflA_26"), ("dram", "DRAM（含 HBM）", "dramA_26"), ("nand", "NAND", "nandA_26")), 21):
    UN.cell(i, 1, nm); UN.cell(i, 2, "US$B"); UN[f"C{i}"] = f"=cy26_{k}"; UN[f"D{i}"] = f"={wf}"; UN[f"E{i}"] = f"=C{i}/D{i}"
    for c in "CD": UN[f"{c}{i}"].number_format = "0.00"
    UN[f"E{i}"].number_format = "0.0%"; defname(f"k_{k}", f"Units!$E${i}")
    UN.cell(i, 8, "Derived：只用公司實績與指引，未用共識；k 是「系統營收 ÷ 模型 WFE 分項」，不是嚴格市占（口徑：系統營收不含升級與 Reliant；WFE 模型為設備市場模型由下而上值）")
UN.cell(24, 1, "合計：Lam 系統營收 ÷ 模型晶圓前段設備"); UN["C24"] = "=cy26_sys"; UN["D24"] = "=wfeA_26"; UN["E24"] = "=C24/D24"; UN["E24"].number_format = "0.0%"
UN.cell(25, 1, "對照：Lam（系統＋CSBG）÷ 公司估 CY2026 WFE"); UN["C25"] = "=cy26_sys+cy26_cs"; UN["D25"] = "=co_wfe26"; UN["E25"] = "=C25/D25"; UN["E25"].number_format = "0.0%"
UN.cell(25, 8, "CSBG 不屬 WFE，僅供量級參考")
UN.column_dimensions["A"].width = 60; UN.column_dimensions["H"].width = 90
for c in "CDEFG": UN.column_dimensions[c].width = 13

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
COL = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}   # FY 與 CY 共用欄：D＝CY2026／FY2026 … G＝CY2029／FY2029
DEFP = dict(h1="h1s", drift="drift", cut27="0", cut28=None, cs27="g_cs27", csg=None, F="opex_fix", v="opex_var", tax="tax", bb="bb", oth="oth_inc", scn="A")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    if P["cut28"] is None: P["cut28"] = "cutC" if s == "C" else P["cut27"]
    if P["csg"] is None: P["csg"] = f"g_cs_{s}"
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025", "CY2026／FY2026", "CY2027／FY2027", "CY2028／FY2028", "CY2029／FY2029", "說明"]); r += 1
    ws.cell(r, 1, "A. 日曆年（D＝CY2026 … G＝CY2029）").font = BOLD; r += 1
    def row(key, label, unit, fs, fmt="0.000", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for y, f in fs.items():
            c = ws[f"{COL[y]}{r}"]; c.value = f; c.number_format = fmt
        ws.cell(r, 8, note); R[key] = r; r += 1
    wfl = lambda y: f"(adv{s}_{y % 100}+ofl{s}_{y % 100})"
    W = {"fl": wfl, "dram": lambda y: f"dram{s}_{y % 100}", "nand": lambda y: f"nand{s}_{y % 100}"}
    for k, nm in (("fl", "代工與邏輯"), ("dram", "DRAM"), ("nand", "NAND")):
        row(f"s_{k}", f"系統營收：{nm}", "US$B", {2026: f"=cy26_{k}", **{y: f"=k_{k}*{W[k](y)}*(1+{P['drift']})^{y - 2026}" for y in (2027, 2028, 2029)}},
            note="CY2026＝錨點（實績＋指引）；之後＝k × 設備市場模型分項 × (1＋份額漂移)^年")
    cut = {2027: P["cut27"], 2028: P["cut28"], 2029: P["cut28"]}
    row("s_cut", "出口管制：中國系統營收折損", "US$B", {y: f"=-cn_sh*{cut[y]}*({COL[y]}{R['s_fl']}+{COL[y]}{R['s_dram']}+{COL[y]}{R['s_nand']})" for y in (2027, 2028, 2029)},
        note="−中國占比 × 折損比例 × 系統營收；情境 C 自 CY2028 起 cutC，其餘為 0（Sens 另設）")
    row("s_cy", "系統營收 CY", "US$B", {y: f"=SUM({COL[y]}{R['s_fl']}:{COL[y]}{R['s_cut']})" for y in (2026, 2027, 2028, 2029)})
    row("c_cy", "CSBG CY", "US$B", {2026: "=cy26_cs", 2027: f"=D{r}*(1+{P['cs27']})", 2028: f"=E{r}*(1+{P['csg']})", 2029: f"=F{r}*(1+{P['csg']})"},
        note="CY2026＝錨點；CY2027 公司稱逾 15%；之後依情境")
    ws.cell(r, 1, "B. 會計年度（C＝FY2025 … G＝FY2029；FY(t)＝CY(t−1) 下半＋CY(t) 上半）").font = BOLD; r += 1
    def fyc(key):
        cy = R[key]
        return {2027: f"={'h2s26' if key == 's_cy' else 'h2c26'}+{P['h1']}*E{cy}",
                2028: f"=(1-{P['h1']})*E{cy}+{P['h1']}*F{cy}", 2029: f"=(1-{P['h1']})*F{cy}+{P['h1']}*G{cy}"}
    row("sys", "系統營收", "US$B", {2025: "=a25_sys", 2026: "=a26_sys", **fyc("s_cy")}, note="FY2027＝CY2026 下半（FY27 Q1 實績＋Q2 指引）＋ h1 × CY2027")
    row("cs", "CSBG", "US$B", {2025: "=a25_cs", 2026: "=a26_cs", **fyc("c_cy")})
    row("rev", "營收", "US$B", {y: f"={COL[y]}{R['sys']}+{COL[y]}{R['cs']}" for y in COL})
    row("gm", "毛利率", "%", {2025: "=a25_gp/a25_rev", 2026: "=a26_gp/a26_rev", **{y: f"=gm_a+gm_b*LN({COL[y]}{R['rev']})" for y in (2027, 2028, 2029)}}, fmt="0.0%",
        note="截距＋斜率 × ln 營收（兩點推導）")
    row("gp", "毛利", "US$B", {y: f"={COL[y]}{R['rev']}*{COL[y]}{R['gm']}" for y in COL})
    row("opx", "營業費用", "US$B", {2025: "=a25_opx", 2026: "=a26_opx", **{y: f"={P['F']}+{P['v']}*{COL[y]}{R['rev']}" for y in (2027, 2028, 2029)}}, note="固定＋變動 × 營收")
    row("oi", "營業利益", "US$B", {y: f"={COL[y]}{R['gp']}-{COL[y]}{R['opx']}" for y in COL})
    row("oim", "營業利益率", "%", {y: f"={COL[y]}{R['oi']}/{COL[y]}{R['rev']}" for y in COL}, fmt="0.0%")
    row("ni", "淨利", "US$B", {2025: "=a25_ni", 2026: "=a26_ni", **{y: f"=({COL[y]}{R['oi']}+{P['oth']})*(1-{P['tax']})" for y in (2027, 2028, 2029)}})
    row("dsh", "稀釋股數", "十億股", {2025: "=a25_dsh", 2026: "=a26_dsh", 2027: "=sh_27", 2028: f"=E{r}*(1-{P['bb']})", 2029: f"=F{r}*(1-{P['bb']})"}, fmt="0.0000")
    row("eps", "EPS（稀釋）", "US$", {2025: "=a25_eps", 2026: "=a26_eps", **{y: f"={COL[y]}{R['ni']}/{COL[y]}{R['dsh']}" for y in (2027, 2028, 2029)}}, fmt="0.00")
    row("ssh", "系統營收占營收", "%", {y: f"={COL[y]}{R['sys']}/{COL[y]}{R['rev']}" for y in COL}, fmt="0.0%")
    R["end"] = r
    return R

# ---------------- PL：三情境 ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — Lam Research 損益（US$B；FY2025、FY2026 為實績，FY2027 起為模型；FY2027 上半錨定 Q1 實績＋Q2 指引）。三情境各一個完整推導區塊（見下方），本頁頂端為摘要"
PB = {}; r0 = 40
for s in SC:
    PB[s] = engine(PL, r0, {"A": "情境 A：設備市場續增（markets/equipment A：2029 見頂）",
                            "B": "情境 B：設備市場 2028 持平、2029 −5%（markets/equipment B）",
                            "C": "情境 C：設備市場 2028 −20%（markets/equipment C）；另加出口管制中國系統營收 −20%（CY2028 起）"}[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026A", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "US$B", "#,##0.00"), ("sys", "  系統營收", "US$B", "#,##0.00"), ("cs", "  CSBG", "US$B", "#,##0.00"),
                                ("gm", "毛利率", "%", "0.0%"), ("oi", "營業利益", "US$B", "#,##0.00"), ("oim", "營業利益率", "%", "0.0%"), ("ni", "淨利", "US$B", "#,##0.00"), ("eps", "EPS（稀釋）", "US$", "0.00")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in COL:
            PL[f"{COL[y]}{sr}"] = f"={COL[y]}{PB[s][key]}"; PL[f"{COL[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對：FY2025、FY2026 模型＝實績（營收、EPS 差的絕對值合計，應為 0）").font = BOLD; sr += 1
PL[f"C{sr}"] = "=" + "+".join(f"ABS({c}{SUMR[(s, 'rev')]}-{a}_rev)+ABS({c}{SUMR[(s, 'eps')]}-{a}_eps)" for s in SC for c, a in (("C", "a25"), ("D", "a26"))); PL[f"C{sr}"].number_format = "0.0000"
PL.column_dimensions["A"].width = 56; PL.column_dimensions["B"].width = 12; PL.column_dimensions["H"].width = 70
for c in "CDEFG": PL.column_dimensions[c].width = 14

# ---------------- Val：P/E 估值 ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值。跨公司規則：合理倍數＝10 年（2016–2025 年底）P/E 的分位，依循環位置：見頂年 25 分位、中段中位數、谷底 75 分位。循環位置依設備股共用規則（見第 4 節）"
V["A2"] = "P/E 定義（自算）：日曆年底收盤價 ÷ 「當期會計年度」公司公布的基本 EPS（年底所在會計年度，例如 2025-12 對 FY2026＝2025-07～2026-06）。價格與 EPS 皆為 2024-10 的 10 比 1 分割調整後"
# 0. 目標價（列號固定在 4–13）
V.cell(4, 1, "0. 目標價與上檔空間").font = BOLD
hdr(V, 5, ["項目", "數值", "說明"])
VAL_ROW = {"A": None}
# 先保留，第 5 節建好後回填
# 1. EPS
V.cell(16, 1, "1. 公司公布 EPS（會計年度，美元，分割調整後）— 兩個來源").font = BOLD
hdr(V, 17, ["會計年度（6 月結束）", "來源1：10-K 基本 EPS（分割調整）", "來源2：10-K 稀釋 EPS（分割調整）", "基本 ÷ 稀釋 − 1", "原始基本 EPS（分割前）", "標記", "說明"])
EPSD = [("FY2016", 5.75, 5.22, True), ("FY2017", 10.47, 9.24, True), ("FY2018", 14.73, 13.17, True), ("FY2019", 14.37, 13.70, True), ("FY2020", 15.55, 15.10, True),
        ("FY2021", 27.22, 26.90, True), ("FY2022", 32.92, 32.75, True), ("FY2023", 3.33, 3.32, False), ("FY2024", 2.91, 2.90, False), ("FY2025", 4.17, 4.15, False), ("FY2026", 5.79, 5.76, False)]
EP = {}
for i, (fy_, b, d, pre) in enumerate(EPSD, 18):
    V.cell(i, 1, fy_)
    if pre:
        V.cell(i, 5, b).font = BLUE; V.cell(i, 2, f"=E{i}/10"); V.cell(i, 3, d / 10).font = BLUE
        note = "SEC XBRL（各年 10-K）：分割前基本 EPS ÷ 10；稀釋同理"
    else:
        V.cell(i, 2, b).font = BLUE; V.cell(i, 3, d).font = BLUE; V.cell(i, 5, f"=B{i}*10")
        note = "FY2025／FY2026 10-K（分割後重編）；E 欄為 ×10 還原（Derived）"
    V.cell(i, 4, f"=B{i}/C{i}-1")
    for c in "BCE": V[f"{c}{i}"].number_format = "0.000"
    V[f"D{i}"].number_format = "0.0%"
    V.cell(i, 6, "Verified"); V.cell(i, 7, note); EP[fy_] = i
V.cell(29, 1, "全部年度統一用基本 EPS（Basic，比照樣板）。來源2 為同一份 10-K 的稀釋 EPS（非獨立來源，只作口徑對照）；獨立的第二個 EPS 來源：未找到（macrotrends、stockanalysis 頁面未取得年度表）")
V.cell(30, 1, "2. 10 年 P/E（年底）").font = BOLD
hdr(V, 31, ["年底", "收盤價 US$（分割調整，未還原股利）", "當期 FY EPS（基本）", "前期 FY EPS（基本）", "P/E 當期（自算）", "P/E 前期（自算）", "companiesmarketcap P/E", "自算當期 − cmc", "macrotrends 年底價（還原股利）", "標記", "來源／說明"])
PX = {2016: 105.73, 2017: 184.07, 2018: 136.17, 2019: 292.40, 2020: 472.27, 2021: 719.15, 2022: 420.30, 2023: 783.26, 2024: 72.23, 2025: 171.18}
MT = {2016: 9.3223, 2017: 16.4265, 2018: 12.4183, 2019: 27.2375, 2020: 44.7187, 2021: 68.7159, 2022: 40.7327, 2023: 76.8336, 2024: 71.5745, 2025: 171.18}
CMC = {2016: 15.1, 2017: 15.9, 2018: 6.83, 2019: 19.4, 2020: 21.9, 2021: 21.3, 2022: 11.0, 2023: 29.7, 2024: 21.7, 2025: 34.9}
for i, y in enumerate(range(2016, 2026), 32):
    V.cell(i, 1, f"{y}-12")
    if y <= 2023: V.cell(i, 2, f"={PX[y]}/10")
    else: V.cell(i, 2, PX[y]).font = BLUE
    V.cell(i, 3, f"=B{EP[f'FY{y+1}']}"); V.cell(i, 4, f"=B{EP[f'FY{y}']}")
    V.cell(i, 5, f"=B{i}/C{i}"); V.cell(i, 6, f"=B{i}/D{i}")
    V.cell(i, 7, CMC[y]).font = BLUE; V.cell(i, 8, f"=E{i}-G{i}"); V.cell(i, 9, MT[y]).font = BLUE
    for c in "BCDI": V[f"{c}{i}"].number_format = "#,##0.00"
    for c in "EFGH": V[f"{c}{i}"].number_format = "0.0"
    V.cell(i, 10, "Verified／Derived")
    V.cell(i, 11, "收盤價：digrin「Real price」（分割前原價 ÷ 10）" if y <= 2023 else "收盤價：digrin（分割後）")
V.cell(42, 1, "收盤價來源1：digrin 年底原價（2016–2023 為分割前，公式 ÷ 10）；來源2：macrotrends 年底價（分割與股利還原，故早年較低；2025 相同）。companiesmarketcap 為 TTM 口徑，只作對照（差距見 H 欄）。")
V.cell(44, 1, "3. 分位").font = BOLD
hdr(V, 45, ["統計", "P/E 當期（自算）", "P/E 前期（自算）", "cmc P/E（參考）", "用途"])
for i, (nm, p, use) in enumerate((("25 分位", 0.25, "見頂年"), ("中位數", 0.5, "中段"), ("75 分位", 0.75, "谷底")), 46):
    V.cell(i, 1, nm); V.cell(i, 5, use)
    V.cell(i, 2, f"=PERCENTILE(E32:E41,{p})"); V.cell(i, 3, f"=PERCENTILE(F32:F41,{p})"); V.cell(i, 4, f"=PERCENTILE(G32:G41,{p})")
    for c in "BCD": V[f"{c}{i}"].number_format = "0.0"
V.cell(49, 1, "年數"); V["B49"] = "=COUNT(E32:E41)"; V["C49"] = "=COUNT(F32:F41)"; V["D49"] = "=COUNT(G32:G41)"
defname("pe_peak", "Val!$B$46"); defname("pe_mid", "Val!$B$47"); defname("pe_trough", "Val!$B$48")
V.cell(50, 1, "讀法：當期口徑的分母是年底時仍在進行中的會計年度（已過一半），接近遠期 P/E；2018、2022 年底記憶體下行使倍數偏低，2023、2025 年底 AI 預期使倍數偏高。")
V.cell(52, 1, "4. 循環位置（設備股共用規則；見頂年度讀設備市場模型，不在本模型另訂）").font = BOLD
V.cell(53, 1, CP.rule_text(PEAK) + "。Lam 的 FY(t)＝t−1 年 7 月～t 年 6 月，日曆年 t 只占一半；比照樣板以會計年度結束的年份 t 對照（FY2027 對照 2027）")
hdr(V, 54, ["情境", "設備市場見頂年度", "FY2027 對照日曆年", "判定", "倍數", "對照：以 2026 對照（lag 1）的判定", "規則"])
for i, s in enumerate(SC, 55):
    V.cell(i, 1, s); V.cell(i, 2, f"=pk_{s}"); V.cell(i, 3, 2027).font = BLUE
    V.cell(i, 4, "=" + CP.excel_rule(f"C{i}", f"pk_{s}"))
    V.cell(i, 5, f'=IF(D{i}="見頂",pe_peak,IF(D{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 6, "=" + CP.excel_rule_lag(f"C{i}", f"pk_{s}", 1)); V.cell(i, 7, "tools/cycle_position.py")
V.cell(59, 1, "5. 隱含股價（美元）：規則值＝FY2027 EPS × 依循環位置的倍數；其餘倍數僅供參考").font = BOLD
hdr(V, 60, ["情境", "EPS 年度", "EPS", "倍數類型", "倍數", "隱含股價", "相對現價", "備註"])
vr = 61; VAL = {}
for s in SC:
    for y, typ, nm in ((2027, f"E{55 + SC.index(s)}", "rule"), (2027, "pe_peak", "參考：25 分位"), (2027, "pe_mid", "參考：中位數"), (2027, "pe_trough", "參考：75 分位"),
                       (2028, "pe_peak", "參考：25 分位"), (2028, "pe_mid", "參考：中位數"), (2028, "pe_trough", "參考：75 分位")):
        V.cell(vr, 1, s); V.cell(vr, 2, f"FY{y}"); V.cell(vr, 3, f"=PL!{COL[y]}{SUMR[(s, 'eps')]}").number_format = "0.00"
        V.cell(vr, 4, nm if nm != "rule" else f'="規則值："&D{55 + SC.index(s)}'); V.cell(vr, 5, f"={typ}").number_format = "0.0"
        V.cell(vr, 6, f"=C{vr}*E{vr}").number_format = "#,##0.0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
        if nm == "rule":
            for c in range(1, 9): V.cell(vr, c).fill = KEY
            VAL[s] = vr
        vr += 1
V.cell(vr + 1, 1, "6. 目前遠期 P/E").font = BOLD
hdr(V, vr + 2, ["項目", "A", "B", "C", "備註"])
fr = vr + 3; FR = {}
for key, nm, fa, fb, fc, fm, note in (
        ("px", "現價（stockanalysis 2026-10-23 收盤）", "=px", None, None, "#,##0.00", ""),
        ("sa", "遠期 P/E：stockanalysis", "=sa_fpe", None, None, "0.0", "stockanalysis 顯示 22.39x（遠期 EPS 口徑未註明）"),
        ("cons", "遠期 P/E：現價 ÷ 共識 FY2027 EPS（Yahoo 9.63）", "=px/cons_eps27", None, None, "0.0", ""),
        ("p26", "現價 ÷ FY2026 EPS（實績 5.76）", "=px/a26_eps", None, None, "0.0", "已結束的會計年度"),
        ("p27", "現價 ÷ FY2027 EPS（模型）", *[f"=px/PL!E{SUMR[(s, 'eps')]}" for s in SC], "0.0", "A／B／C"),
        ("p28", "現價 ÷ FY2028 EPS（模型）", *[f"=px/PL!F{SUMR[(s, 'eps')]}" for s in SC], "0.0", "A／B／C"),
        ("rel", "現價隱含倍數 ÷ 規則倍數（情境 A）— 以 FY2027 A EPS", None, None, None, "0.00", ">1＝現價高於規則給的倍數")):
    V.cell(fr, 1, nm)
    for j, f in enumerate((fa, fb, fc)):
        if f is not None: V.cell(fr, 2 + j, f).number_format = fm
    V.cell(fr, 5, note); FR[key] = fr; fr += 1
V.cell(FR["rel"], 2, f"=B{FR['p27']}/E55").number_format = "0.00"
# 回填第 0 節
TP = [("目標價（12 個月，情境加權）", f"=tp_wA*B8+tp_wB*B9+tp_wC*B10", "#,##0.0", "＝權重 A × 情境 A 隱含價＋權重 B × 情境 B 隱含價＋權重 C × 情境 C 隱含價"),
      ("上檔空間（目標價 ÷ 現價 − 1）", "=B6/px-1", "0.0%", "現價：stockanalysis 2026-10-23 收盤"),
      ("情境 A 隱含價（樂觀）", f"=F{VAL['A']}", "#,##0.0", f"＝第 5 節規則值（FY2027 EPS × 依循環位置的倍數）"),
      ("情境 B 隱含價（基準）", f"=F{VAL['B']}", "#,##0.0", "同上"),
      ("情境 C 隱含價（下檔）", f"=F{VAL['C']}", "#,##0.0", "同上"),
      ("情境權重 A／B／C", "=tp_wA", "0%", "In 頁 tp_wA／tp_wB／tp_wC（Assumed；Andy 可改）"),
      ("目標價方法", "P/E：模型 FY2027 稀釋 EPS（A／B／C 各自）× 依設備股共用循環規則的 10 年（2016–2025 年底）當期 P/E 分位（A、B 中段＝中位數，C 見頂＝25 分位）；樣本含 2018、2022 記憶體下行年的低倍數與 2023、2025 年底的 AI 高倍數，未調整", None, ""),
      ("檢查：權重合計 − 1（應為 0）", "=tp_wA+tp_wB+tp_wC-1", "0.00", "")]
for i, (lab, f, fm, note) in enumerate(TP, 6):
    V.cell(i, 1, lab); c = V.cell(i, 2, f)
    if fm: c.number_format = fm
    V.cell(i, 3, note)
V["C11"] = "=tp_wB"; V["D11"] = "=tp_wC"; V["C11"].number_format = "0%"; V["D11"].number_format = "0%"; V["E11"] = "In 頁 tp_wA／tp_wB／tp_wC（Assumed；Andy 可改）"
for c in range(1, 4): V.cell(6, c).fill = KEY
V.column_dimensions["A"].width = 52
for c in "BCDEFGHIJ": V.column_dimensions[c].width = 16
V.column_dimensions["K"].width = 60

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A）：對 FY2027、FY2028 營收與 EPS 的影響。每個案例是一個完整的推導區塊（本頁下方，與 PL 同一引擎）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收 US$B", "FY2027 EPS US$", "相對基準", "FY2028 營收 US$B", "FY2028 EPS US$", "相對基準"])
cases = [("份額漂移 −2%／年", "CY2027 起每年相對 WFE 分項 −2%", dict(drift="drift_lo")), ("份額漂移 +2%／年", "每年 +2%", dict(drift="drift_hi")),
         ("出口管制：中國系統營收 −20%（CY2027 起）", "cn_sh 34% × 20%", dict(cut27="cut_sens_lo")), ("出口管制：中國系統營收 −40%（CY2027 起）", "cn_sh 34% × 40%（2022 管制類比上緣）", dict(cut27="cut_sens_hi")),
         ("上半年占比 45%", "FY(t) 中 CY(t) 上半的權重", dict(h1="h1s_lo")), ("上半年占比 55%", "", dict(h1="h1s_hi")),
         ("CSBG CY2027 +10%", "公司稱逾 15%", dict(cs27="g_cs27_lo")), ("CSBG CY2027 +20%", "", dict(cs27="g_cs27_hi")),
         ("CSBG CY2028–29 +8%", "", dict(csg="g_cs_A_lo")), ("CSBG CY2028–29 +16%", "", dict(csg="g_cs_A_hi")),
         ("營業費用變動比例 低（15%）", "FY2026 營業費用中變動部分 15%", dict(F="(1-phi_lo)*a26_opx", v="phi_lo*a26_opx/a26_rev")),
         ("營業費用變動比例 高（50%）", "變動 50%", dict(F="(1-phi_hi)*a26_opx", v="phi_hi*a26_opx/a26_rev")),
         ("稅率 13%", "", dict(tax="tax_lo")), ("稅率 19%", "", dict(tax="tax_hi")),
         ("庫藏股 0%／年", "", dict(bb="bb_lo")), ("庫藏股 3%／年", "", dict(bb="bb_hi")),
         ("業外收益 0", "", dict(oth="oth_inc_lo")), ("業外收益 0.2", "", dict(oth="oth_inc_hi"))]
def put_case(i, name, setting, R, pre=""):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    SE[f"C{i}"] = f"={pre}E{R['rev']}"; SE[f"D{i}"] = f"={pre}E{R['eps']}"; SE[f"F{i}"] = f"={pre}F{R['rev']}"; SE[f"G{i}"] = f"={pre}F{R['eps']}"
    SE[f"E{i}"] = f"=D{i}/D$4-1"; SE[f"H{i}"] = f"=G{i}/G$4-1"
    for c in "CF": SE[f"{c}{i}"].number_format = "#,##0.00"
    for c in "DG": SE[f"{c}{i}"].number_format = "0.00"
    for c in "EH": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
SB = {}
for n, (nm, setting, ov) in enumerate(cases):
    SB[nm] = engine(SE, 50 + n * 30, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    put_case(5 + n, nm, setting, SB[nm])
nn = 5 + len(cases)
put_case(nn, "設備市場情境 B", "WFE 分項讀 markets/equipment B；CSBG 成長 B", PB["B"], "PL!")
put_case(nn + 1, "設備市場情境 C", "WFE 分項讀 C；出口管制 −20%（CY2028 起）；CSBG 成長 C", PB["C"], "PL!")
SE.cell(nn + 3, 1, "讀法：FY2027 上半已錨定（Q1 實績＋Q2 指引），只有下半（CY2027 上半）隨輸入變動，所以 FY2027 的敏感度小、FY2028 才明顯；設備市場 B、C 的係數從 CY2028 開始。")
SENS_LAST = nn + 1
SE.column_dimensions["A"].width = 44; SE.column_dimensions["B"].width = 46
for c in "CDEFGH": SE.column_dimensions[c].width = 15

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司說法、共識、設備市場模型。只作事後對照，不為對上數字而調整參數。差距 > 10% 的原因見 README"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.00", pt=False):
    global rr_
    RC.cell(rr_, 1, label); RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        L = chr(64 + cv)
        RC.cell(rr_, cd, f"={L}{rr_}-B{rr_}" if pt else f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y: (lambda s: f"=PL!{COL[y]}{SUMR[(s, key)]}")
RC.cell(rr_, 1, "A. FY2027、FY2028 vs 共識（校準未使用共識）").font = BOLD; rr_ += 1
rec("FY2027 營收", "=cons_rev27", pl("rev", 2027), "Yahoo（29 位）", "stockanalysis 34.9")
rec("FY2027 EPS（Yahoo）", "=cons_eps27", pl("eps", 2027), "Yahoo（9.63）", "共識多為非 GAAP；Lam 非 GAAP 與 GAAP 差約 0.01／季")
rec("FY2027 EPS（stockanalysis）", "=cons_eps27b", pl("eps", 2027), "stockanalysis（9.48）")
rec("FY2028 營收", "=cons_rev28", pl("rev", 2028), "Yahoo（28 位；區間 36.5–49.2）")
rec("FY2028 EPS", "=cons_eps28", pl("eps", 2028), "Yahoo（12.15；區間 10.41–15.12）")
RC.cell(rr_, 1, "B. 公司說法（法說，Interested-party）").font = BOLD; rr_ += 1
rec("CY2026 系統營收年增（公司：well over 40%）", "=co_sys26", lambda s: "=Units!G14", "FQ1-27 法說", "模型值＝實績＋指引（含 Q2 拆分假設），非獨立檢驗", fmt="0.0%", pt=True)
rec("CY2026 CSBG 年增（公司：逾 30%）", "=co_cs26", lambda s: "=Units!G15", "FQ1-27 法說", "同上；Q2 CSBG＝Q1 持平時", fmt="0.0%", pt=True)
rec("FY2027 毛利率 vs Q2 指引 52.5%", "=g_q2gm", pl("gm", 2027), "8-K 2026-10-21", "公司稱近期維持目前水準；長期 mid-50s", fmt="0.0%", pt=True)
rec("FY2029 毛利率 vs 公司長期目標（mid-50s，取 55%）", "=0.55", pl("gm", 2029), "FQ4-26／FQ1-27 法說", "", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "C. 設備市場（WFE，US$B）").font = BOLD; rr_ += 1
rec("CY2026 晶圓前段（公司 mid-$150B vs 設備市場模型）", "=co_wfe26", lambda s: f"=wfe{s}_26", "FQ1-27 法說", "口徑不同（公司 CY2025 約 110、SEMI 116.9）")
rec("CY2027 晶圓前段（公司 >$200B vs 設備市場模型）", "=co_wfe27", lambda s: f"=wfe{s}_27", "FQ1-27 法說", "下限值")
rec("CY2026／CY2025 晶圓前段成長：公司（155÷110）vs 模型", "=co_wfe26/co_wfe25-1", lambda s: f"=wfe{s}_26/wfe{s}_25-1", "FQ2-26、FQ1-27 法說", "", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "D. 其他資訊（非對帳）").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "FY2027 上半（Q1 實績＋Q2 指引）÷ 模型 FY2027 全年營收（A）"); RC.cell(rr_, 3, f"=(q1_rev+g_q2rev)/PL!E{SUMR[('A', 'rev')]}").number_format = "0.0%"; rr_ += 1
RC.cell(rr_, 1, "FY2027 上半 EPS（Q1 實績＋Q2 指引）÷ 模型 FY2027 EPS（A）"); RC.cell(rr_, 3, f"=(q1_eps+g_q2eps)/PL!E{SUMR[('A', 'eps')]}").number_format = "0.0%"; rr_ += 1
RC.cell(rr_, 1, "Lam 系統營收 ÷ 模型晶圓前段（CY2026）"); RC.cell(rr_, 3, "=Units!E24").number_format = "0.0%"; rr_ += 1
RECON_LAST = rr_ - 1
RC.column_dimensions["A"].width = 58
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 30; RC.column_dimensions["J"].width = 70

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司財報取自 SEC EDGAR 原文；法說取自逐字稿摘要）"; SO["A1"].font = BOLD
E = "https://www.sec.gov/Archives/edgar/data/707549/"
srcs = [("公司：FY27 Q1 8-K 新聞稿（2026-10-21）", E + "000070754926000052/lrcx_exhibitx991xq1x2027.htm"),
 ("公司：FY27 Q1 10-Q（2026-10-22）", E + "000070754926000054/lrcx-20260927.htm"),
 ("公司：FY2026 10-K（2026-08-07）", E + "000070754926000037/lrcx-20260628.htm"),
 ("公司：FY26 Q4 8-K 新聞稿（2026-07-29）", E + "000070754926000033/lrcx_exhibitx991xq4x2026.htm"),
 ("公司：FY26 Q3 8-K 新聞稿（2026-04-22）", E + "000070754926000020/lrcx_exhibitx991xq3x2026.htm"),
 ("公司：FY26 Q2 8-K 新聞稿（2026-01-28）", E + "000070754926000006/lrcx_exhibitx991xq2x2026.htm"),
 ("公司：FY26 Q1 8-K 新聞稿（2025-10-22）", E + "000070754925000082/lrcx_exhibitx991xq1x2026.htm"),
 ("公司：FY25 Q4 8-K 新聞稿（2025-07-30）", E + "000070754925000068/lrcx_exhibitx991xq4x2025.htm"),
 ("公司：10-Q（FY26 Q1、Q2、Q3：客戶別占比）", E + "000070754926000022/lrcx-20260329.htm"),
 ("SEC XBRL：EPS 歷史（基本、稀釋）", "https://data.sec.gov/api/xbrl/companyconcept/CIK0000707549/us-gaap/EarningsPerShareBasic.json"),
 ("Motley Fool：FQ1-27 法說逐字稿（WFE、客戶別、CSBG、稅率）", "https://www.fool.com/earnings/call-transcripts/2026/10/21/lam-research-lrcx-q1-2027-earnings-call-transcript/"),
 ("MarketBeat：FQ4-26 法說重點（6 月季客戶別、WFE low-$150B）", "https://www.marketbeat.com/instant-alerts/lam-research-q4-earnings-call-highlights-2026-07-29/"),
 ("TIKR：FQ2-26 法說重點（12 月季客戶別、CY2025 WFE 約 110）", "https://www.tikr.com/blog/lam-research-stock-after-q2-eps-beats-foundry-surges-and-nand-waits"),
 ("GuruFocus：FQ1-27 法說重點（$200B WFE）", "https://www.gurufocus.com/news/9136007/lam-research-corp-lrcx-q1-2027-earnings-call-highlights-record-revenue-firstever-40-operating-margin-and-a-bullish-200-billion-wfe-outlook"),
 ("techwireasia：2022-10 出口管制影響 US$2.0–2.5B（類比）", "https://techwireasia.com/2022/10/us-chip-ban-on-china-may-cost-lam-research-a-staggering-us2-5bn/"),
 ("stockanalysis：現價、遠期 P/E", "https://stockanalysis.com/stocks/lrcx/"),
 ("stockanalysis：共識（FY2027）", "https://stockanalysis.com/stocks/lrcx/forecast/"),
 ("Yahoo Finance：共識（FY2027、FY2028）", "https://finance.yahoo.com/quote/LRCX/analysis/"),
 ("digrin：年底收盤價（原價與還原價）", "https://www.digrin.com/stocks/detail/LRCX/price"),
 ("macrotrends：年底收盤價（還原股利）", "https://www.macrotrends.net/stocks/charts/LRCX/lam-research/stock-price-history"),
 ("companiesmarketcap：年底 P/E（TTM 口徑）", "https://companiesmarketcap.com/lam-research/pe-ratio/"),
 ("市場層：markets/equipment 成品（本 repo，唯讀）", rel(EQ_F)),
 ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 70; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("Lam Research（LRCX）收支與估值模型 v0.2（AI 半導體 Project 步驟 5：設備與材料）　2026-10-25", True),
 ("模型命題：Lam 的系統營收跟著晶圓前段設備（WFE）的記憶體與代工邏輯支出走，記憶體占系統營收近一半（DRAM 與 NAND 升級）、中國占營收約三分之一。核心問題：WFE 在 2027 年後放緩、或對中國的出口管制加嚴時，FY2027 以後的 EPS 還撐得住現在約 25 倍的遠期本益比嗎？", False), ("", False),
 ("會計年度：7 月到隔年 6 月（FY2027＝2026-07～2027-06）；FY(t)＝CY(t−1) 下半＋CY(t) 上半（上半占比 h1，基準 50%）。FY2026 已結束（實績）；FY2027 第一季（2026-09 季）已於 2026-10-21 公布並納入，12 月季用公司指引。金額十億美元。", False), ("", False),
 ("驅動因子 → 推導量", True),
 ("系統營收（日曆年）＝ k（份額係數）× 設備市場模型的 WFE 分項（代工與邏輯、DRAM、NAND；三情境；markets/equipment）；k 由 CY2026（FY26 Q3–Q4 實績＋FY27 Q1 實績＋Q2 指引）× 客戶別占比校準", False),
 ("CSBG（服務、備件、升級、Reliant）另推：CY2026 錨點 × 年增（CY2027 公司稱逾 15%；之後依情境 Assumed）", False),
 ("毛利率＝截距＋斜率 × ln 營收（FY2026 與 FY2027 上半年化兩點推導）；營業費用＝固定＋變動 × 營收（同兩點）；稅率 16.5%（公司說法）；股數 1.253B 起每年 −1.5%", False),
 ("情境 A／B／C：WFE 分項讀設備市場模型 A（2029 見頂）／B（2028 持平）／C（2028 −20%）；情境 C 另加出口管制：中國系統營收 −20%（CY2028 起，Assumed；2022 管制類比約三到四成）", False),
 ("Val：P/E 規則——10 年（2016–2025 年底）當期 P/E 分位，見頂年 25 分位、中段中位數、谷底 75 分位；循環位置依設備股共用規則（各情境見頂年度：" + "、".join(f"{s_} {PEAK[s_]}" for s_ in SC) + " → FY2027：" + CP.fy2027_text(PEAK) + "）。第 0 節：情境加權目標價。本版未做先驗，無 Andy 合理 P/E 對照。", False), ("", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。", False),
 ("已知限制：(1) k 只在 CY2026 一點校準，之後份額不變（漂移只在 Sens）；公司稱連續三年優於 WFE，若延續則模型偏低；(2) 客戶別占比為「設備與升級營收」口徑（10-Q），套用於系統營收；DRAM／NAND 拆分來自法說摘要；(3) 中國以總營收占比代用系統營收占比，出口管制折損為 Assumed；(4) 設備市場模型的 WFE 以 SEMI 2025 實績校準，與公司 WFE 口徑不同，只用成長率；(5) CY→FY 以上下半各半轉換，成長年下半較高（h1 區間 45–55%）；(6) CSBG 成長無由下而上（裝機量 × 每腔體支出）推導；(7) NAND 升級潮公司稱 2027 年結束，模型的 NAND 隨設備市場模型 NAND 支出成長，可能高估 CY2028 後的 NAND。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

CFG_V2 = dict(cols=COL, scn={"A": "順利", "B": "平平", "C": "不如預期"}, peers=[('應用材料（Applied Materials）', 'AMAT', 21.39, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('科磊（KLA）', 'KLAC', 26.65, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('ASML（美國 ADR）', 'ASML', 24.74, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('東京威力科創（Tokyo Electron）', '8035.T', 27.09, '2026-10-21', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）')], fwd={2028: "8/12", 2029: "4/12"}, fwd_note='約 2027-10 時的未來 12 月（2027-11～2028-10）≈ 8／12 × FY2028（2027-07～2028-06）＋4／12 × FY2029（會計年度 7–6 月，依月份調整）', a3='目標價（DCF 與 Comparable 加權）：見本頁第 {row} 列起「0. 目標價」（估值 v2 新增，放在既有內容之後以免位移既有列）；原「0. 目標價與上檔空間」改名為「舊法：規則值情境加權（參考）」',
    cur='US$',
    unit_rev='US$B',
    unit_ni='US$B',
    fmt='0.00',
    fmt_rev='#,##0.00',
    band_fmt='0.00',
    roe=('FY2024–FY2026', [0.4571, 0.5824, 0.6507], 'stockanalysis 比率頁（2026-10-25 讀取）：FY2024 45.71%、FY2025 58.24%、FY2026 65.07%', '單一來源；近 12 月 70.12%（同站）'),
    nim=('FY2016–FY2025，GAAP', [(5.89, 0.91405), (8.01, 1.7), (11.08, 2.38), (9.65, 2.19), (10.04, 2.25), (14.63, 3.91), (17.23, 4.61), (17.43, 4.51), (14.91, 3.83), (18.44, 5.36)], 'vcpscanner.com 年度損益（US$B）', '第二來源：FY2025 與 Q 頁實績核對；其餘年度單一來源'),
    cons_tp=(364.03, 'stockanalysis：34 位分析師平均 US$364.03（中位數 352.50；區間 255–500），資料更新 2026-10-22', '第二來源：未找到'),
    cons_feps=('=cons_eps28', '共識前瞻 EPS＝共識 FY2028 EPS（In 頁 cons_eps28，Yahoo 12.15）；本模型前瞻 EPS 為 8／12 FY2028＋4／12 FY2029，約晚四個月', 'Derived'),
    peer_note='前段設備同業：應用材料、科磊、ASML、東京威力科創',
    dcf_note='；會計年度 7–6 月，FY2027＝2026-07～2027-06 以 t＝1 折現（略低估）',
    pe10='E32:E41',
    pe10_note='現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（當期口徑自算）的分位；只作檢查，不進目標價',
    px_note='現價＝In 頁 px（2026-10-23 收盤）',
    method='DCF：股東現金流＝EPS ×（1 − g／ROE）（模型無折舊與資本支出項目），r 10%、g 4%、ROE 為近三年平均，FY2027–32 明確期（FY2030 起收斂法）＋永續。Comparable：應用材料、科磊、ASML、東京威力科創前瞻 P/E 中位數 × 12 個月後前瞻 EPS。偏誤：會計年度 7–6 月以 t＝1 折現略低估；記憶體與中國循環風險只在情境 C 反映。',
    readme=['v0.2（2026-10-25）估值 v2：Val 頁新增「0. 目標價」（放在既有內容之後，不位移既有列）＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，情境權重 A 0.3／B 0.5／C 0.2）＋0.5 × Comparable 目標價（同業前瞻 P/E 中位數 × 12 個月後前瞻 EPS）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解、歷史分位檢查。原「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」，不再叫目標價。', '預估期間延伸到 FY2032：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（PL 頁底部延伸區塊，灰底）。情境對應：A＝順利、B＝平平、C＝不如預期（原為顆數情境）。本版未做先驗，無 Andy 合理 P/E 對照。三情境在 FY2028 之前完全相同（FY2027 只差顆數係數、而係數從 CY2028 才分歧），DCF 情境仍照算，差異來自 FY2028 之後。'])

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

# ---------------- v0.2.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00／14:17）----------------
# 新列一律放在 Val、README 工作表既有內容之後，不位移既有列；既有數值不變
VER_PREV, VER_NOW = 'v0.2', 'v0.2.1'
OBS = [
 "目標價 US$203（DCF 147、Comparable 259 各半），比現價 230.10（2026-10-23）低 11.8%，級距 0%～−25%；Comparable 單獨為 +12.5%。DCF 比現價低 36.2%，因情境 B、C 的 EPS 在 FY2029 後依收斂法下滑（C 的 FY2032 EPS 5.25 只有 FY2027 9.42 的 56%），不是 ROE（近三年 56.3%）。",
 "與市場的主要差異在 EPS：本模型前瞻 EPS 10.07 比共識 FY2028 EPS 12.15 低 17.1%（期間約晚四個月），情境 A 的 FY2028 EPS 10.68 也比共識低 12.1%；共識目標價 364.03 隱含 30.0 倍，比同業前瞻 P/E 中位數 25.7 倍高 16.6%。",
 "出口管制與設備市場情境最敏感：中國系統營收 −40% 使 FY2028 EPS −11.8%、設備情境 C 使 −18.2%。下一個檢驗點是 FY2027 第二季（12 月季）財報（預計 2027-01 下旬），檢驗 CY2027 上半年的權重（45%～55% 時 FY2027 EPS ±6.6%）。"
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
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "lrcx")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.2.2", "2026-10-31")  # v0.2.2：In 頁版面（個股模型原則第 8 節；數值不變）
wb.save(sys.argv[1])
if os.environ.get("LRCX_ROWS"):
    print("IN", {k: INROW[k] for k in ("h1s", "drift", "cn_sh", "cutC", "g_cs27", "gm_b", "tax", "bb", "px", "pk_A", "dramA_26", "tp_wA")})
    print("PB", PB["A"], "SUMR", SUMR, "VAL", VAL, "FR", FR, "SENS_LAST", SENS_LAST, "RECON_LAST", RECON_LAST)

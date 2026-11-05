# 泰瑞達（Teradyne，TER US）收支與估值模型 v0.4.1（AI 半導體 Project 步驟 5：設備與材料——測試機第二家）
# 用法：python3 ter/build_ter.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年（FY2026＝2026-01～12）。金額：百萬美元（US$M）；EPS：美元（GAAP 稀釋）。
# 樣板：advantest/（工作表結構、三情境、Val 列標籤、html.json 七章）。
import sys, os, glob, json
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
import cycle_position as CP
import fcff_dcf as FD   # v0.4：標準版 DCF（無槓桿自由現金流＋WACC）共用模組

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
YRS = (2026, 2027, 2028, 2029)

# ---------------- 讀入其他模型的成品（唯讀；不修改它們） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
ADV_F = only("advantest/dist/*.xlsx"); TSMC_F = only("tsmc/dist/*.xlsx"); EQ_F = only("markets/equipment/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)

# Advantest 模型：SoC／記憶體測試機市場（CY，三情境）、Advantest SoC 市占（FY 口徑，三情境）——同一份市場，份額合計可核對
wa = openpyxl.load_workbook(ADV_F, data_only=True)
ws_ = wa["PL"]
ADV = {"soc": {}, "mem": {}, "sh": {}}; n_ = {"soc": 0, "mem": 0, "sh": 0}
for r in ws_.iter_rows(min_row=40):
    lab = r[0].value
    k = {"SoC 測試機市場 CY": "soc", "記憶體測試機市場 CY": "mem", "Advantest SoC 市占": "sh"}.get(lab)
    if k and n_[k] < 3:
        s = SC[n_[k]]; n_[k] += 1
        if k == "sh":   # FY2026..FY2029 在 D..G
            ADV[k][s] = ([r[3 + j].value for j in range(4)], f"{rel(ADV_F)}｜PL!D{r[0].row}:G{r[0].row}")
        else:           # CY2026..CY2029 在 C..F
            ADV[k][s] = ([r[2 + j].value for j in range(4)], f"{rel(ADV_F)}｜PL!C{r[0].row}:F{r[0].row}")
assert all(n_[k] == 3 for k in n_), n_
ADVAI = {}   # v0.13：Advantest 模型 AI SoC 測試機市場（CY；Wattonomics 錨定案例只對 AI 部分乘係數）
for r in ws_.iter_rows(min_row=40):
    if r[0].value == "AI SoC 測試機市場" and len(ADVAI) < 3:
        ADVAI[SC[len(ADVAI)]] = ([r[2 + j].value for j in range(4)], f"{rel(ADV_F)}｜PL!C{r[0].row}:F{r[0].row}")
assert len(ADVAI) == 3, ADVAI
# CF15：SoC 測試機市場只有一個值（Siliconomics Equip 頁）；Advantest 成品三情境 CY2026–29 必須與本庫相同
assert all(abs(a - b) < 1e-6 for a, b in zip(ADV["soc"]["B"][0], SIL.NAMES["SIF_TestSoCMkt_Base"][1:])), ("SoC 測試機市場與 Siliconomics 基準不一致", ADV["soc"]["B"][0], SIL.NAMES["SIF_TestSoCMkt_Base"])   # v0.15：只核對基準
ADVIN = {}
for r in wa["In"].iter_rows(min_row=4):
    if r[0].value: ADVIN[r[0].value] = (r[4].value, r[3].value, r[5].value, f"{rel(ADV_F)}｜In!E{r[0].row}")
# 台積電模型：情境顆數係數（與 Advantest 讀的同一組，建置時核對）
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and (k.startswith("um28_") or k.startswith("um29_")):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
for s in SC:
    for y in ("28", "29"):
        assert abs(TS[f"um{y}_{s}"][0] - ADVIN[f"um{y}_{s}"][0]) < 1e-12, (s, y)
# 設備市場模型：各情境晶圓前段設備（WFE；測試對 WFE 比的對帳用）
WFE = {}; _n = 0
for r in openpyxl.load_workbook(EQ_F, data_only=True)["Mkt"].iter_rows(min_row=4, max_row=40):
    if r[0].value == "  其中晶圓前段":
        s = SC[_n]; _n += 1
        WFE[s] = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert _n == 3
PEAK, PEAK_SRC = CP.peak_years(ROOT)

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：百萬美元（US$M），市場規模為十億美元（US$B）。每個輸入附兩個來源（來源2 找不到者寫「未找到」）；公司 SEC 原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文（SEC 8-K 新聞稿、10-Q、10-K）在建置環境可開啟"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t))
def inp(*a):   # v0.7 起，泰瑞達測試 ÷ WFE 區間（Siliconomics M21：ts_tw_lo／ts_tw_hi）改讀本庫快照
    rows.append((a[0], *SIL.row9("ts_" + a[0])[1:]) if a[0] in ("tw_lo", "tw_hi") else a)

sec("A. 測試機市場（讀 Advantest 模型成品；與 Advantest 同一份市場，不另估）")
inp("mk_soc25", "SoC 測試機市場 CY2025", "US$B", None, ADVIN["mk_soc25"][0], None, "Interested-party", ADVIN["mk_soc25"][3], "Advantest（測試機廠）自估；與 Advantest 模型同一值")
inp("mk_mem25", "記憶體測試機市場 CY2025", "US$B", None, ADVIN["mk_mem25"][0], None, "Interested-party", ADVIN["mk_mem25"][3], "同上")
for s in SC:
    for j, y in enumerate(YRS):
        inp(f"ms_{s}_{y}", f"SoC 測試機市場 CY{y}（{s}）", "US$B", None, ADV["soc"][s][0][j], None, "Derived", ADV["soc"][s][1] if j == 0 else "同上",
            "Advantest 模型：加速器顆數（markets/memory UNITS）× 每顆測試時間 × 每加權顆支出（以 CY2026 公司市場規模校準）＋非 AI；情境係數讀台積電模型" if (s, j) == ("A", 0) else "同上")
for s in SC:
    for j, y in enumerate(YRS):
        inp(f"mm_{s}_{y}", f"記憶體測試機市場 CY{y}（{s}）", "US$B", None, ADV["mem"][s][0][j], None, "Derived", ADV["mem"][s][1] if j == 0 else "同上",
            "Advantest 模型：CY2026 公司市場規模 × [HBM 出貨位元指數（markets/memory，同字母）與 DRAM 設備支出指數（markets/equipment，同字母）加權]" if (s, j) == ("A", 0) else "同上")
for s in SC:
    for j, y in enumerate(YRS):
        inp(f"as_{s}_{y}", f"Advantest SoC 市占 FY{y}（{s}；Advantest 會計年度 4–3 月）", "%", None, ADV["sh"][s][0][j], None, "Interested-party", ADV["sh"][s][1] if j == 0 else "同上",
            "Advantest 自估 CY2025 66%；情境 C 假設泰瑞達搶回（FY2028 61%、FY2029 56%）。本模型以同年度日曆年對照（FY(t) 有 3/4 落在 CY(t)）" if (s, j) == ("A", 0) else "同上")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}（個股自訂）：2028 加速器顆數 ÷ 原始顆數（說明用）", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], SIL.scen_note(s, 28, TS[f"um28_{s}"][0]))
    inp(f"um29_{s}", f"情境 {s}（個股自訂）：2029 加速器顆數 ÷ 原始顆數（說明用）", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], SIL.scen_note(s, 29, TS[f"um29_{s}"][0]))
for s in SC:
    for j, y in enumerate((2025, 2026, 2027, 2028, 2029)):
        inp(f"wfe_{s}_{y}", f"晶圓前段設備（WFE）CY{y}（{s}）", "US$B", None, WFE[s][0][j], None, "Derived", WFE[s][1] if j == 0 else "同上", "markets/equipment Mkt 頁（只用於 Recon 的測試對 WFE 比）" if (s, j) == ("A", 0) else "同上")

sec("B. 泰瑞達市占（FY2026 錨點推得；A、B 情境維持，C 情境與 Advantest 互補）")
inp("s_soc", "泰瑞達 SoC 測試機市占（FY2026 錨點 SoC 營收 ÷ CY2026 市場）", "%", "=s_soc-0.05", "=f26_soc/(ms_A_2026*1000)", "=s_soc+0.05", "Derived", "Q 頁 FY2026 錨點 SoC 營收 ÷ Advantest 模型 CY2026 SoC 市場（Siliconomics v0.9：11.5）", "低／高＝±5 個百分點（Assumed）。分析師在 2Q26 法說估泰瑞達測試份額約 37%（公司未證實；口徑不同）")
inp("s_mem", "泰瑞達記憶體測試機市占（FY2026 錨點 ÷ CY2026 市場）", "%", "=s_mem-0.05", "=f26_mem/(mm_A_2026*1000)", "=s_mem+0.05", "Derived", "Q 頁 FY2026 錨點記憶體營收 ÷ Advantest 模型 CY2026 記憶體市場 2.75", "FY2025 隱含 24.0%（504.9 ÷ 2,100）；2Q26 公司：DRAM 強、NAND 最終測試回溫，記憶體營收創新高")
inp("adv_sh26", "Advantest SoC 市占 FY2026（三情境相同；互補計算的基準）", "%", None, "=as_A_2026", None, "Derived", "In 頁 as_A_2026", "C 情境泰瑞達市占＝s_soc＋(adv_sh26 − Advantest 當年市占)，使兩家合計與其他廠份額不變")

sec("C. FY2026 錨點（上半年實績＋3Q26 實績＋4Q26 指引中點；三情境相同）")
inp("q4_rev", "4Q26 營收指引中點（1,250–1,350）", "US$M", 1250, 1300, 1350, "Interested-party", "公司 3Q26 新聞稿（8-K Ex.99.1，2026-10-27）：4Q26 營收 1,250–1,350", "3Q26 法說（2026-10-28）：全年營收約 5.2B")
inp("q4_eps", "4Q26 GAAP 稀釋 EPS 指引中點（1.88–2.18）", "US$", 1.88, 2.03, 2.18, "Interested-party", "同上：GAAP 1.88–2.18；non-GAAP 1.95–2.25", "—")
inp("co_rev26", "公司口述 FY2026 全年營收（3Q26 法說）", "US$M", None, 5200, None, "Interested-party", "公司 3Q26 法說（2026-10-28，Motley Fool 逐字稿）：全年營收約 5.2B", "只作 Recon 對照，不進推導")
inp("h2g", "4Q26 非 SoC 事業（記憶體、IST、Product Test、機器人）÷ 3Q26 − 1", "%", -0.10, 0.0, 0.10, "Assumed", "公司 3Q26 法說（2026-10-28）：記憶體 4Q 維持強勁、可能在創新高後消化；Product Test、機器人續成長；運算 4Q 回升", "SoC＝4Q26 營收指引 − 非 SoC（殘差，Derived）")
inp("rob_h2", "下半年機器人稅前利益", "US$M", 0, 5, 15, "Assumed", "公司 3Q26 新聞稿：機器人 3Q26 單季轉盈（金額未揭露）；上半年 −3.5", "公司目標 2026 年兩平（4Q25 法說）")

sec("D. 其他事業（前一年 × 成長率；機器人不隨 AI 情境）")
for s, g in (("A", 0.08), ("B", 0.05), ("C", 0.02)):
    inp(f"g_pt_{s}", f"Product Test 年成長 FY2027–29（{s}）", "%", g - 0.04, g, g + 0.04, "Assumed", "FY2025 +8.1%（331.1→358.0，10-K）；2Q26 +26%（AI 相關與國防航太）", "—")
inp("g_rob", "機器人營收年成長 FY2027–29（三情境相同）", "%", 0.0, 0.10, 0.20, "Assumed", "公司 2Q26：中期機器人隨公司整體成長；FY2025 −15.5%（364.8→308.3，10-K）、2Q26 +33%", "機器人另列，不隨 AI 情境")
for y, m in ((2027, 0.02), (2028, 0.04), (2029, 0.05)):
    inp(f"rm_{y}", f"機器人稅前利益率 FY{y}", "%", m - 0.07, m, m + 0.05, "Assumed", "FY2025 −32.2%（−99.4 ÷ 308.3，含重組）；上半年 FY2026 −1.8%（−3.5 ÷ 191.2）", "公司目標 2026 年兩平（Interested-party）")

sec("E. 獲利結構（兩點推導：FY2025 實績與 FY2026 錨點；不寫死利益率）")
inp("t_inc", "測試事業（半導體測試＋Product Test）增量稅前利益率", "%", 0.40, "=(f26_ttest-a25_ttest)/(f26_rtest-a25_rtest)", 0.56, "Derived", "（FY2026 錨點測試稅前 − FY2025）÷（營收差）；Q 頁", "低／高 Assumed（固定成本重新推導，使 FY2026 錨點不變）")
inp("t_fix", "測試事業固定成本（＝增量利益率 × FY2026 營收 − FY2026 稅前）", "US$M", "=t_inc_lo*f26_rtest-f26_ttest", "=t_inc*f26_rtest-f26_ttest", "=t_inc_hi*f26_rtest-f26_ttest", "Derived", "同上", "以 FY2026 錨點定錨（與 FY2025 兩點同一直線）")
inp("corp", "公司與沖銷（利息、匯兌、併購費用等）每年", "US$M", None, "=f26_corp", None, "Derived", "FY2026 錨點（上半年 −5.2 × 2）", "FY2025 −8.8（10-K）")
inp("etr", "有效稅率", "%", 0.12, "=(h1_tax+Q!I16)/(h1_pt+Q!I15)", 0.17, "Derived", "前三季 FY2026：（128.9＋59.8）÷（909.3＋404.3）（10-Q、3Q26 新聞稿）", "FY2025 12.1%；3Q26 14.8%")
inp("aff", "權益法投資損益（Technoprobe，含投資攤銷）每年", "US$M", None, "=2*h1_aff", None, "Derived", "上半年 FY2026 −6.6 × 2（10-Q）", "FY2025 −19.9")
inp("sh26", "稀釋股數 FY2026", "百萬股", None, 157.5, None, "Verified", "公司 3Q26 新聞稿：3Q26 稀釋加權股數 157.5", "前三季 157.6")
inp("bb", "稀釋股數每年減少（回購）", "%", 0.0, 0.01, 0.02, "Assumed", "FY2024→FY2025 −2.2%（163.3→159.7）、FY2025→2Q26 −1.3%", "上半年 FY2026 回購僅 74.2（FY2025 702.1）")
inp("gm_adj_C28", "情境 C：FY2028 測試事業利益率額外調整（價格競爭）", "pt", ADVIN["gm_adj_C28"][1], ADVIN["gm_adj_C28"][0], ADVIN["gm_adj_C28"][2], "Assumed", ADVIN["gm_adj_C28"][3], "與 Advantest 情境 C 毛利率調整同值（兩家同受價格競爭）")
inp("gm_adj_C29", "情境 C：FY2029 測試事業利益率額外調整", "pt", ADVIN["gm_adj_C29"][1], ADVIN["gm_adj_C29"][0], ADVIN["gm_adj_C29"][2], "Assumed", ADVIN["gm_adj_C29"][3], "同上")
inp("m_soc", "SoC 測試機市場乘數（2027 起；Sens 用）", "x", 0.85, 1.0, 1.0, "Assumed", "基準 1.0＝Advantest 模型市場；低 0.85＝測試對 WFE 比回到公司說的 7–9%（見 Recon）", "—")

sec("F. 公司指引、目標模型與共識（只作對帳）")
inp("h1sh_lo", "上半年占全年營收：下緣", "%", None, 0.50, None, "Interested-party", "公司 2Q26 法說：上半年約占 50–52%（earningswhispers 摘要）", "第二來源：未找到")
inp("h1sh_hi", "上半年占全年營收：上緣", "%", None, 0.52, None, "Interested-party", "同上", "—")
inp("cons_rev26", "共識營收 FY2026", "US$M", None, 5100, None, "Interested-party", "stockanalysis：US$5.1B（16 位分析師，資料更新 2026-10-20）", "第二來源：MarketBeat 3Q26 共識營收 1.26B（單季）")
inp("cons_eps26", "共識 EPS FY2026（non-GAAP）", "US$", None, 9.09, None, "Interested-party", "stockanalysis：9.09（區間 8.82–9.35；2026-10-20）", "MarketBeat：9.11")
inp("cons_eps27", "共識 EPS FY2027（non-GAAP）", "US$", None, 11.30, None, "Interested-party", "MarketBeat（2026-10-20）：11.30", "stockanalysis 需付費，未取得")
inp("tm_rev", "公司目標模型：營收（ATE 市場 US$12–14B 時，未給年度）", "US$M", None, 6000, None, "Interested-party", "公司 4Q25 法說（2026-02-03，fintool 摘要）：營收 6.0B、營業利益率 30–34%、non-GAAP EPS 9.50–11.00", "公司稱 4Q26 法說更新目標模型")
inp("tm_eps_lo", "公司目標模型：non-GAAP EPS 下緣", "US$", None, 9.50, None, "Interested-party", "同上", "—")
inp("tm_eps_hi", "公司目標模型：non-GAAP EPS 上緣", "US$", None, 11.00, None, "Interested-party", "同上", "—")
inp("tw_lo", "公司：測試對 WFE 比長期區間 下緣", "%", None, 0.07, None, "Interested-party", "公司 2Q26 法說：2023 約 4%、2025 約 7%、2026 前五月約 8%；長期 7–9%，2027 可能回到 7%（earningswhispers）", "gurufocus 2Q26 摘要（同數字）")
inp("tw_hi", "公司：測試對 WFE 比長期區間 上緣", "%", None, 0.09, None, "Interested-party", "同上", "—")
inp("an_share", "分析師估泰瑞達測試份額（2Q26 法說提問）", "%", None, 0.37, None, "Interested-party", "earningswhispers 2Q26 逐字稿摘要：分析師估約 37%，公司未證實", "口徑未明（可能為 SoC 或 ATE 合計）")
for s in SC:
    inp(f"pk_{s}", f"情境 {s}：設備市場見頂年度（日曆年；循環位置規則用）", "年", None, PEAK[s], None, "Derived", PEAK_SRC[s], "設備股共用規則（tools/cycle_position.py）：FY(t) 早於見頂年＝中段、等於＝見頂、晚於＝谷底；泰瑞達會計年度＝日曆年" if s == "A" else "同上")
inp("px", "股價（2026-10-23 收盤）", "US$", None, 308.85, None, "Verified", "stockanalysis（2026-10-23 收盤 308.85）", "equibles（同值）")
inp("sa_fpe", "遠期 P/E（stockanalysis，2026-10-23 價格）", "x", None, 32.52, None, "Interested-party", "stockanalysis statistics（2026-10-24 更新）", "自算見 Val")
sec("I. 規則值情境權重（Val 頁「舊法：規則值情境加權（參考）」與目標價 DCF 共用；Andy 可改）")
for s_, w_ in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s_}", f"目標價情境權重 {s_}", "%", None, w_, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%")

# v0.4：標準版 DCF（無槓桿自由現金流＋WACC；tools/fcff_dcf.py）的 In 參數。評價日 2026-12-31（會計年度＝日曆年）；股價 2026-10-23；資產負債為 3Q26（2026-09-27）
UST10 = ("美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610；2026-11-05 讀取）",
         "FRED DGS10（同一序列，https://fred.stlouisfed.org/series/DGS10）：10-28～10-30 5.05%／4.98%／4.94% 與財政部表一致；與股價同日（2026-10-23）；同 NVIDIA v0.3", "無風險利率（美國 10 年期公債殖利率，2026-10-23）")
DAM_US = ("自算產業 β：US 半導體設備 同業 7 家中位數（AMAT、LRCX、KLAC、TER、FORM、COHU、KLIC；Yahoo Finance 月報酬 2021-11～2026-10，對 S&P 500 ^GSPC 迴歸，Hamada 去槓桿、現金修正）無槓桿 β 1.75；以公司 D/E（市值）與稅率依 Hamada 重新加槓桿：β＝βU ×（1＋(1−t) × D/E）",
          "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）：採用無槓桿 β 1.503。低／高（Analogy；第二輪修正起）：β_u ∓ beta_du（0.3），各自依公司 D/E 與稅率重新加槓桿（工作單第二輪修正，Andy 2026-11-05 23:16 定案；tools/fcff_dcf.py v2 write_sens）；取代 Damodaran 美國產業表")
CRP_US = (0.0, "Assumed", "美國等成熟市場：0（同 NVIDIA v0.3；Damodaran 2026-01-05 表美國 Aa1、CRP 0.23%，本模型不加）", "非美國市場：Damodaran 最新 Country Risk Premium 表（Andy 2026-11-05 20:45）")
SA = "stockanalysis.com TER 財報（US$M，2026-11-05 讀取）"
SAQ = "stockanalysis.com TER 季度資產負債表（3Q26，2026-09-27；US$M；2026-11-05 讀取）https://stockanalysis.com/stocks/ter/financials/balance-sheet/?p=quarterly"
RELEV = "*(1+(1-tax_w)*(debt_v+lease_v)/(px*sh_v))"
for _row in FD.in_rows(dict(
        date="2026-12-31（會計年度＝日曆年；同 v0.3 DCF 基準日）", amt="US$M", shu="百萬股", ccy="US$",
        rf=(0.0500,) + UST10,
        stub=dict(bs_date="2026-09-27", val_date="2026-12-31", bs_src=SAQ + "；Teradyne 10-Q（3Q26，最新已公布一季）",
                  div=(0, "Verified", "Teradyne 股利紀錄：最近一次每股 0.13 美元（登記日 2026-09-04、2026-09-25 發放，在資產負債表日之前）；2026-12 發放的下一次建置日 2026-11-05 尚未宣告（去年同期 2025-11-10 宣告），不扣（https://www.quantisnow.com/dividends/TER）",
                       "買回：無已宣告的 stub 期金額，不扣（股數沿用評價用稀釋股數，與 FCFF 一致）")),
        beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Derived") + DAM_US,
        crp=CRP_US,
        kd=("=rf+0.01", "Assumed", "rf ＋ 1 個百分點信用利差（泰瑞達無金融負債，只有租賃）", "負債占資本不到 1%，對 WACC 影響可忽略"),
        tax_w=("=etr", "Interested-party", "＝In 頁 etr（模型稅率）", "NOPAT 各年也用 etr（與 PL 所得稅＝稅前利益 × etr 同一稅率）"),
        debt=(0, "Verified", SAQ + "：無借款", "Teradyne 10-Q（3Q26）"),
        lease=("=17.06+120.05", "Verified", SAQ + "：租賃負債 流動 17.06＋非流動 120.05", "Teradyne 10-Q（3Q26）"),
        cash=("=460.74+5.85", "Verified", SAQ + "：現金及約當現金 460.74＋短期投資 5.85", "Teradyne 10-Q（3Q26）"),
        inv=(676.84, "Verified", SAQ + "：長期投資 676.84（含 Technoprobe 權益法投資與退休金相關投資）", "帳面價值計入；Sens 測不計"),
        mi=(0, "Verified", SAQ + "：無少數股權", "—"),
        dps=("=76.31/sh26", "Derived", SA + "：FY2025 已付股利 76.31 ÷ 稀釋股數 sh26（每股約 0.48 美元）", "評價日到目標時點之間假設同此水準（Assumed）"),
        sh=("=sh26", "Verified", "＝In 頁 sh26（3Q26 稀釋加權股數 157.5M）", "回購由 FCFF 支應，不另推未來股數"),
        fc_da=(None, "=(110.89/2676.3+119.74/2819.9+127.98/a25_rev)/3", None, "Derived", SA + "：近三年折舊攤銷 ÷ 營收：FY2023 110.89／2,676.3、FY2024 119.74／2,819.9、FY2025 127.98／FY2025 營收（Q 頁 a25_rev）", "Teradyne 10-K 現金流量表"),
        fc_cx=(None, "=(159.64/2676.3+198.1/2819.9+224.01/a25_rev)/3", None, "Derived", SA + "：近三年資本支出 ÷ 營收：FY2023 6.0%、FY2024 7.0%、FY2025 7.0%", "Sens 用 fc_cx_d ±1pt"),
        fc_wc=(0.05, "=((422.12+309.97+46.71-180.13-290.76-99.8-37.28)/2676.3+(471.43+298.49+63.21-134.79-282.81-107.71-40.51)/2819.9+(786.91+379.55+63.39-269.19-347.67-153.12-50.89)/a25_rev)/3", 0.20, "Derived",
               "近三年營運資金 ÷ 營收平均：營運資金＝應收＋存貨＋預付 − 應付 − 應計費用 − 流動與長期預收款（stockanalysis 季度資產負債表 12 月底欄；其他流動資產不計）", "Sens 低／高 5%／20%（Assumed）"),
        )):
    rows.append(_row)
# 第二輪修正（2026-11-05）：β 以 β_u 重新加槓桿寫成公式；β 敏感度＝β_u ± 0.3（tools/fcff_dcf.py v2 write_sens 讀 beta_u、beta_de、beta_du）
inp("beta_u", "產業無槓桿 β（採用）：自算 US 半導體設備 同業中位數 Blume 調整", "x", None, 1.503, None, "Derived", "自算產業 β：US 半導體設備 同業中位數（Yahoo Finance 月報酬 2021-11～2026-10，對本地大盤迴歸，Hamada 去槓桿、現金修正）；見 In 頁 beta 列來源", "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務本金 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 稀釋股數 sh_v)", "與 WACC 權重同一口徑")
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3", "取代上一版的「中位數 × 0.8 再 Blume／中位數未調整」")

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
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
INROW = {IN.cell(i, 1).value: i for i in range(4, IN.max_row + 1) if IN.cell(i, 1).value}

def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績（FY2025 四季、FY2026 上半年）與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績（百萬美元；EPS 為美元，GAAP 稀釋）。會計年度＝日曆年。藍字＝輸入（SEC 8-K 新聞稿、10-Q、10-K）；黑字＝公式（Derived）"
Q["A2"] = "v0.3（2026-10-31）：I 欄滾入 3Q26（2026-10-27 新聞稿 8-K Ex.99.1；SoC、記憶體、IST 分項取自 2026-10-28 法說）；FY2026 錨點改為上半年＋3Q26 實績＋4Q26 指引中點"
hdr(Q, 3, ["項目", "1Q25", "2Q25", "3Q25", "4Q25", "FY2025", "1Q26", "2Q26", "3Q26", "上半年 FY2026", "標記", "來源1（公司 SEC 文件）", "來源2／說明"])
QR = {}
def qrow(r, key, label, vals, tag, s1, s2, fmt="#,##0.0"):
    QR[key] = r
    Q.cell(r, 1, label)
    for j, v in zip("BCDEFGHIJ", vals):
        if v is None: continue
        c = Q[f"{j}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    Q.cell(r, 11, tag); Q.cell(r, 12, s1); Q.cell(r, 13, s2)
PR = "公司各季新聞稿（8-K Ex.99.1）"; DQ = "10-Q／10-K 收入拆分表（Disaggregation of Revenue）"
qrow(4, "soc", "SoC 測試機（System on-a-Chip）", [406.4, 396.6, 440.2, "=F4-SUM(B4:D4)", 1889.7, "=J4-H4", 843.0, 676.0, 1724.8], "Verified／Derived",
     DQ + "：2Q25 396.6、3Q25 440.2、FY2025 1,889.7、2Q26 843.0、上半年 1,724.8；1Q25＝上半年 803.0 − 2Q25", "gurufocus 2Q26 法說摘要：SoC 843、記憶體 212、IST 67；3Q26 法說（2026-10-28，Motley Fool／Benzinga 逐字稿）：SoC 676、記憶體 301、IST 95")
qrow(5, "mem", "記憶體測試機", [109.4, 60.9, 128.1, "=F5-SUM(B5:D5)", 504.9, "=J5-H5", 212.3, 301.0, 414.8], "Verified／Derived", DQ + "：1Q25＝上半年 170.4 − 2Q25；4Q25＝全年 − 前三季", "同上（記憶體 212，創新高）")
qrow(6, "ist", "IST（整合系統測試）", [26.7, 34.3, 37.6, "=F6-SUM(B6:D6)", 129.2, "=J6-H6", 66.5, 95.0, 93.1], "Verified／Derived", DQ, "同上（IST 67，+94%）")
qrow(7, "semi", "半導體測試（SoC＋記憶體＋IST）", [f"={c}4+{c}5+{c}6" for c in "BCDEFGHI"] + ["=J4+J5+J6"], "Derived", "公司新聞稿：1Q25 543、2Q25 492、3Q25 606、4Q25 883、1Q26 1,111、2Q26 1,122；FY2025 2,523.7（10-K）", "檢查列見下")
qrow(8, "pt", "Product Test", [74.2, 85.1, 88.3, "=F8-SUM(B8:D8)", 358.0, "=J8-H8", 107.2, 138.0, 187.7], "Verified／Derived", PR + "與 10-K（FY2025 358.0）", "MD&A 各季比較表")
qrow(9, "rob", "機器人（Robotics）", [69.0, 74.9, 75.1, "=F9-SUM(B9:D9)", 308.3, "=J9-H9", 99.9, 96.0, 191.2], "Verified／Derived", PR + "與 10-K（FY2025 308.3）", "MD&A 各季比較表")
qrow(10, "rev", "營收", [685.7, 651.8, 769.2, 1083.3, 3190.0, 1282.5, 1329.0, 1306.0, 2611.5], "Verified", PR + "：FY2025 3,190.0（+13%）；2Q26 1,329.0（+104%）", "stockanalysis／MarketBeat（2Q26 1.33B）")
qrow(11, "chk", "檢查：半導體測試＋Product Test＋機器人 − 營收（應約為 0；單季分項四捨五入）", [f"={c}7+{c}8+{c}9-{c}10" for c in "BCDEFGHI"] + ["=J7+J8+J9-J10"], "—", "—", "—")
qrow(12, "gp", "毛利", [415.3, 373.0, 449.3, 619.7, 1857.3, 780.9, 794.6, 771.3, 1575.6], "Verified", PR, "—")
qrow(13, "gm", "毛利率", [f"={c}12/{c}10" for c in "BCDEFGHI"] + ["=J12/J10"], "Derived", "—", "公司：FY2025 58.3%（non-GAAP）；3Q26 指引 58–59%；全年約 59%（2Q26 法說）", "0.0%")
qrow(14, "oi", "營業利益（GAAP）", [120.8, 90.7, 145.3, 293.2, 650.1, 473.0, 437.8, 402.0, 910.8], "Verified／Derived", PR + "：1Q25＝上半年 211.5 − 2Q25", "—")
qrow(15, "pti", "稅前利益（含權益法前）", [119.0, 96.6, 148.1, 289.6, 653.3, 465.7, 443.6, 404.3, 909.3], "Verified／Derived", PR + "：1Q25＝上半年 215.6 − 2Q25", "—")
qrow(16, "tax", "所得稅", [14.5, 12.3, 23.3, 29.2, 79.3, 62.2, 66.8, 59.8, 128.9], "Verified／Derived", PR, "—")
qrow(17, "aff", "權益法投資損益（Technoprobe）；3Q26 起含少數股東權益", [-5.6, -5.9, -5.2, -3.2, -19.9, -4.6, -1.9, -2.1, -6.6], "Verified／Derived", PR + "；3Q26＝權益法 +3.3 − 少數股東權益 5.3（新聞稿損益表）", "—")
qrow(18, "ni", "淨利（歸屬泰瑞達）", [98.9, 78.4, 119.6, 257.2, 554.0, 398.9, 374.5, 342.4, 773.4], "Verified", PR, "—")
qrow(19, "eps", "EPS（GAAP 稀釋）", [0.61, 0.49, 0.75, 1.63, 3.47, 2.53, 2.38, 2.17, 4.91], "Verified", PR + "（四季加總 3.48 與全年 3.47 差為股數加權）", "macrotrends 年度 EPS（FY2024 3.32 等）", "0.00")
qrow(20, "dsh", "稀釋加權股數（百萬股）", [162.0, 160.1, 159.1, 157.7, 159.7, 157.6, 157.7, 157.5, 157.7], "Verified", PR, "—")
Q.cell(22, 1, "事業別稅前利益（10-K／10-Q 分部資訊；機器人另列）").font = BOLD
hdr(Q, 23, ["事業", "FY2024", "FY2025", "上半年 FY2026", "2Q26", "", "", "", "", "", "標記", "來源1", "來源2／說明"])
SEG = [("ttsemi", "半導體測試", 558.2, 700.7, 895.4, 427.3), ("ttpt", "Product Test", 65.7, 60.7, 22.7, 18.0),
       ("ttrob", "機器人", -77.6, -99.4, -3.5, -2.6), ("ttcorp", "公司與沖銷", 62.7, -8.8, -5.2, 0.9)]
for i, (k, nm, a, b, c, d) in enumerate(SEG, 24):
    Q.cell(i, 1, nm)
    for j, v in zip("BCDE", (a, b, c, d)):
        Q[f"{j}{i}"] = v; Q[f"{j}{i}"].font = BLUE; Q[f"{j}{i}"].number_format = "#,##0.0"
    Q.cell(i, 11, "Verified"); Q.cell(i, 12, "10-K FY2025、10-Q 2Q26：Income (Loss) Before Income Taxes 分部表"); Q.cell(i, 13, "FY2024 公司與沖銷含出售 DIS 事業利益 57.1")
    QR[k] = i
Q.cell(28, 1, "合計（應＝稅前利益）")
for j in "BCDE": Q[f"{j}28"] = f"=SUM({j}24:{j}27)"; Q[f"{j}28"].number_format = "#,##0.0"
Q.cell(28, 13, "FY2025 應＝653.3；上半年應＝909.3（四捨五入差）")

# FY2026 錨點
Q.cell(30, 1, "FY2026 錨點（上半年實績＋下半年推估；三情境相同）").font = BOLD
hdr(Q, 31, ["項目", "上半年（實績）", "下半年（3Q26 實績＋4Q26 推估）", "FY2026", "", "", "", "", "", "", "標記", "說明", ""])
AN = [("soc", "SoC 測試機", "=J4", "=C40-C33-C34-C35-C36", "Derived：下半年營收 − 非 SoC 事業（殘差）"),
      ("mem", "記憶體測試機", "=J5", "=I5*(2+h2g)", "3Q26 實績＋4Q26（＝3Q26 × (1＋h2g)）"),
      ("ist", "IST", "=J6", "=I6*(2+h2g)", "同上"),
      ("pt", "Product Test", "=J8", "=I8*(2+h2g)", "同上"),
      ("rob", "機器人", "=J9", "=I9*(2+h2g)", "同上；不隨 AI 情境")]
for i, (k, nm, h1, h2, note) in enumerate(AN, 32):
    Q.cell(i, 1, nm); Q.cell(i, 2, h1); Q.cell(i, 3, h2); Q.cell(i, 4, f"=B{i}+C{i}"); Q.cell(i, 12, note)
    for c in "BCD": Q[f"{c}{i}"].number_format = "#,##0.0"
    defname(f"f26_{k}", f"Q!$D${i}")
Q.cell(37, 1, "測試事業營收（半導體測試＋Product Test）"); Q["B37"] = "=B32+B33+B34+B35"; Q["C37"] = "=C32+C33+C34+C35"; Q["D37"] = "=B37+C37"; defname("f26_rtest", "Q!$D$37")
Q.cell(38, 1, "（保留）")
Q.cell(39, 1, "營收：下半年＝3Q26 實績＋4Q26 指引中點"); Q["C39"] = "=I10+q4_rev"; Q.cell(39, 12, "Verified（3Q26）＋Interested-party（4Q26 指引）")
Q.cell(40, 1, "營收"); Q["B40"] = "=J10"; Q["C40"] = "=C39"; Q["D40"] = "=B40+C40"; defname("f26_rev", "Q!$D$40")
Q.cell(41, 1, "上半年占全年（公司 50–52%）"); Q["D41"] = "=B40/D40"; Q["D41"].number_format = "0.0%"
Q.cell(42, 1, "淨利（下半年＝3Q26 實績＋4Q26 GAAP EPS 指引中點 × 稀釋股數）"); Q["B42"] = "=J18"; Q["C42"] = "=I18+q4_eps*sh26"; Q["D42"] = "=B42+C42"; defname("f26_ni", "Q!$D$42")
Q.cell(43, 1, "權益法投資損益與少數股東權益（4Q26＝3Q26）"); Q["B43"] = "=J17"; Q["C43"] = "=2*I17"; Q["D43"] = "=B43+C43"; defname("f26_aff", "Q!$D$43")
Q.cell(44, 1, "稅前利益（下半年＝3Q26 實績＋4Q26（淨利 − 權益法）÷（1 − 稅率））"); Q["B44"] = "=J15"; Q["C44"] = "=I15+(C42-I18-(C43-I17))/(1-etr)"; Q["D44"] = "=B44+C44"; defname("f26_pti", "Q!$D$44")
Q.cell(45, 1, "所得稅"); Q["B45"] = "=J16"; Q["C45"] = "=I16+(C44-I15)*etr"; Q["D45"] = "=B45+C45"; defname("f26_tax", "Q!$D$45")
Q.cell(46, 1, "機器人稅前利益"); Q["B46"] = "=D26"; Q["C46"] = "=rob_h2"; Q["D46"] = "=B46+C46"; defname("f26_trob", "Q!$D$46")
Q.cell(47, 1, "公司與沖銷（下半年＝上半年）"); Q["B47"] = "=D27"; Q["C47"] = "=B47"; Q["D47"] = "=B47+C47"; defname("f26_corp", "Q!$D$47")
Q.cell(48, 1, "測試事業稅前利益（稅前 − 機器人 − 公司與沖銷）"); Q["B48"] = "=B44-B46-B47"; Q["C48"] = "=C44-C46-C47"; Q["D48"] = "=B48+C48"; defname("f26_ttest", "Q!$D$48")
Q.cell(49, 1, "EPS（GAAP 稀釋）"); Q["D49"] = "=D42/sh26"; Q["D49"].number_format = "0.00"; defname("f26_eps", "Q!$D$49")
Q.cell(50, 1, "毛利率（上半年實績；下半年＝3Q26 實績＋4Q26 同 3Q26 毛利率 59.1%）"); Q["B50"] = "=J13"; Q["C50"] = "=(I12+q4_rev*I13)/C40"; Q["D50"] = "=(B40*B50+C40*C50)/D40"
for c in "BCD": Q[f"{c}50"].number_format = "0.0%"
defname("f26_gm", "Q!$D$50")
for rr in range(39, 49):
    for c in "BCD":
        if Q[f"{c}{rr}"].value is not None and rr != 41: Q[f"{c}{rr}"].number_format = "#,##0.0"
Q.cell(51, 1, "對照：FY2025 測試事業營收、測試事業稅前利益（10-K）")
Q["B51"] = "=F7+F8"; Q["C51"] = "=C24+C25"; defname("a25_rtest", "Q!$B$51"); defname("a25_ttest", "Q!$C$51")
for c in "BC": Q[f"{c}51"].number_format = "#,##0.0"
Q.cell(52, 1, "上半年 FY2026：稅前、所得稅、權益法（10-Q）")
Q["B52"] = "=J15"; Q["C52"] = "=J16"; Q["D52"] = "=J17"; defname("h1_pt", "Q!$B$52"); defname("h1_tax", "Q!$C$52"); defname("h1_aff", "Q!$D$52")
for nm, rr in (("soc", 4), ("mem", 5), ("ist", 6), ("semi", 7), ("pt", 8), ("rob", 9), ("rev", 10), ("gp", 12), ("pti", 15), ("tax", 16), ("aff", 17), ("ni", 18), ("eps", 19), ("dsh", 20)):
    defname(f"a25_{nm}", f"Q!$F${rr}")
defname("a25_trob", "Q!$C$26"); defname("a25_corp", "Q!$C$27")
Q.column_dimensions["A"].width = 58
for c in "BCDEFGHIJ": Q.column_dimensions[c].width = 13
Q.column_dimensions["K"].width = 18; Q.column_dimensions["L"].width = 80; Q.column_dimensions["M"].width = 60
Q.freeze_panes = "B4"

# ---------------- Units：市場與份額一致性 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 測試機市場（十億美元；日曆年；讀 Advantest 模型成品）與 SoC 市占合計。泰瑞達 A、B 情境市占維持 FY2026 水準；C 情境與 Advantest 互補（Advantest 讓出的份額全由泰瑞達取得），三情境的「其他廠」份額相同"
hdr(UN, 3, ["項目", "單位", "CY2025", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ycol = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
UR = {}; ur = 4
def tshare(s, y, sh="s_soc"):
    return sh if (y <= 2026 or s != "C") else f"({sh}+adv_sh26-as_C_{y})"
for s in SC:
    UN.cell(ur, 1, f"情境 {s}").font = BOLD; ur += 1
    for key, lab, unit, f25, fy, fmt in (
            ("soc", "SoC 測試機市場", "US$B", "=mk_soc25", lambda y: f"=ms_{s}_{y}", "0.00"),
            ("mem", "記憶體測試機市場", "US$B", "=mk_mem25", lambda y: f"=mm_{s}_{y}", "0.00"),
            ("adv", "Advantest SoC 市占（Advantest 模型，FY 口徑）", "%", "=as_A_2026", lambda y: f"=as_{s}_{y}", "0.0%"),
            ("ter", "泰瑞達 SoC 市占", "%", "=a25_soc/(mk_soc25*1000)", lambda y: "=" + tshare(s, y), "0.0%")):
        UN.cell(ur, 1, "  " + lab); UN.cell(ur, 2, unit); UN[f"C{ur}"] = f25
        for y in YRS: UN[f"{ycol[y]}{ur}"] = fy(y)
        for y in (2025,) + YRS: UN[f"{ycol[y]}{ur}"].number_format = fmt
        UR[(s, key)] = ur; ur += 1
    UN.cell(ur, 1, "  其他廠 SoC 市占（1 − Advantest − 泰瑞達）"); UN.cell(ur, 2, "%")
    for y in (2025,) + YRS: UN[f"{ycol[y]}{ur}"] = f"=1-{ycol[y]}{UR[(s,'adv')]}-{ycol[y]}{UR[(s,'ter')]}"; UN[f"{ycol[y]}{ur}"].number_format = "0.0%"
    UR[(s, "oth")] = ur
    UN.cell(ur, 8, "CY2025 用 Advantest 自估 66% 對照；三情境 CY2026 起其他廠份額相同（互補設計）"); ur += 1
    UN.cell(ur, 1, "  SoC＋記憶體測試機市場 ÷ 晶圓前段設備（WFE）"); UN.cell(ur, 2, "%")
    for y in (2025,) + YRS:
        a_ = "mk_soc25+mk_mem25" if y == 2025 else f"ms_{s}_{y}+mm_{s}_{y}"
        UN[f"{ycol[y]}{ur}"] = f"=({a_})/wfe_{s}_{y}"; UN[f"{ycol[y]}{ur}"].number_format = "0.0%"
    UR[(s, "tw")] = ur
    UN.cell(ur, 8, "公司說法：測試對 WFE 比 2025 約 7%、2026 約 8%，長期 7–9%（口徑為 ATE 市場，含其他測試機；見 Recon）"); ur += 2
UN.column_dimensions["A"].width = 56; UN.column_dimensions["B"].width = 10; UN.column_dimensions["H"].width = 90
for c in "CDEFG": UN.column_dimensions[c].width = 12

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
FYC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2030)}   # FY2025＝C … FY2029＝G
DEFP = dict(sh="s_soc", shm="s_mem", msoc="m_soc", gpt=None, grob="g_rob", rm="", inc="t_inc", fix="t_fix", etr="etr", bb="bb", scn="A", wm=None)
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    gpt = P["gpt"] or f"g_pt_{s}"
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.0", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 8, note); R[key] = r; r += 1
    fy = lambda y: FYC[y]
    F = (2027, 2028, 2029)
    row("msoc", "SoC 測試機市場 CY（Advantest 模型）", "US$B", {"C": "=mk_soc25", "D": f"=ms_{s}_2026", **{fy(y): f"=ms_{s}_{y}*{P['msoc']}" + (f"+({P['wm'][y]}-1)*adv_ai{s}_{y % 100}" if P["wm"] else "") for y in F}}, fmt="0.00")   # v0.13：wm＝Wattonomics 錨定（只對 AI SoC 部分乘係數；預設 None＝公式不變）
    row("mmem", "記憶體測試機市場 CY（Advantest 模型）", "US$B", {"C": "=mk_mem25", "D": f"=mm_{s}_2026", **{fy(y): f"=mm_{s}_{y}" for y in F}}, fmt="0.00")
    row("sh", "泰瑞達 SoC 市占", "%", {"C": f"=a25_soc/(C{R['msoc']}*1000)", "D": f"=f26_soc/(D{R['msoc']}*1000)", **{fy(y): "=" + tshare(s, y, P["sh"]) for y in F}}, fmt="0.0%",
        note="FY2025、FY2026＝實績／錨點隱含；C 情境與 Advantest 互補" if s == "C" else "FY2025、FY2026＝實績／錨點隱含；之後維持")
    row("rsoc", "SoC 測試機營收", "US$M", {"C": "=a25_soc", "D": "=f26_soc", **{fy(y): f"={fy(y)}{R['msoc']}*1000*{fy(y)}{R['sh']}" for y in F}}, note="市場 × 市占")
    row("shm", "泰瑞達記憶體測試機市占", "%", {"C": f"=a25_mem/(C{R['mmem']}*1000)", "D": f"=f26_mem/(D{R['mmem']}*1000)", **{fy(y): f"={P['shm']}" for y in F}}, fmt="0.0%")
    row("rmem", "記憶體測試機營收", "US$M", {"C": "=a25_mem", "D": "=f26_mem", **{fy(y): f"={fy(y)}{R['mmem']}*1000*{fy(y)}{R['shm']}" for y in F}})
    row("rist", "IST 營收（隨 SoC 市場）", "US$M", {"C": "=a25_ist", "D": "=f26_ist", **{fy(y): f"=$D{r}*{fy(y)}{R['msoc']}/$D{R['msoc']}" for y in F}}, note="FY2026 × SoC 市場相對 CY2026")
    row("rsemi", "半導體測試營收", "US$M", {fy(y): f"={fy(y)}{R['rsoc']}+{fy(y)}{R['rmem']}+{fy(y)}{R['rist']}" for y in range(2025, 2030)})
    row("rpt", "Product Test 營收", "US$M", {"C": "=a25_pt", "D": "=f26_pt", **{fy(y): f"={fy(y-1)}{r}*(1+{gpt})" for y in F}})
    row("rtest", "測試事業營收（半導體測試＋Product Test）", "US$M", {fy(y): f"={fy(y)}{R['rsemi']}+{fy(y)}{R['rpt']}" for y in range(2025, 2030)})
    adj = {2027: "0", 2028: "gm_adj_C28" if s == "C" else "0", 2029: "gm_adj_C29" if s == "C" else "0"}
    row("ttest", "測試事業稅前利益", "US$M", {"C": "=a25_ttest", "D": "=f26_ttest", **{fy(y): f"={P['inc']}*{fy(y)}{R['rtest']}-{P['fix']}+{adj[y]}*{fy(y)}{R['rtest']}" for y in F}},
        note="增量利益率 × 營收 − 固定成本（兩點推導）；C 情境另加價格競爭調整")
    row("mtest", "測試事業稅前利益率", "%", {fy(y): f"={fy(y)}{R['ttest']}/{fy(y)}{R['rtest']}" for y in range(2025, 2030)}, fmt="0.0%")
    row("rrob", "機器人營收（不隨 AI 情境）", "US$M", {"C": "=a25_rob", "D": "=f26_rob", **{fy(y): f"={fy(y-1)}{r}*(1+{P['grob']})" for y in F}})
    row("trob", "機器人稅前利益", "US$M", {"C": "=a25_trob", "D": "=f26_trob", **{fy(y): f"={fy(y)}{R['rrob']}*rm_{y}{P['rm']}" for y in F}})
    row("corp", "公司與沖銷", "US$M", {"C": "=a25_corp", "D": "=f26_corp", **{fy(y): "=corp" for y in F}})
    row("rev", "營收", "US$M", {fy(y): f"={fy(y)}{R['rtest']}+{fy(y)}{R['rrob']}" for y in range(2025, 2030)})
    row("chk", "檢查：分部合計 − 合併營收（FY2025 實績、FY2026 錨點；四捨五入到百萬美元）", "US$M", {"C": f"=ROUND(C{R['rsemi']}+C{R['rpt']}+C{R['rrob']}-a25_rev,0)", "D": f"=ROUND(D{R['rsemi']}+D{R['rpt']}+D{R['rrob']}-f26_rev,0)"}, fmt="#,##0.0",
        note="FY2025 分部合計 3,190.1 vs 公司合併營收 3,190.0：季度分部數字四捨五入差 0.1")
    row("pti", "稅前利益", "US$M", {fy(y): f"={fy(y)}{R['ttest']}+{fy(y)}{R['trob']}+{fy(y)}{R['corp']}" for y in range(2025, 2030)})
    row("pm", "稅前利益率", "%", {fy(y): f"={fy(y)}{R['pti']}/{fy(y)}{R['rev']}" for y in range(2025, 2030)}, fmt="0.0%")
    row("tax", "所得稅", "US$M", {"C": "=a25_tax", "D": "=f26_tax", **{fy(y): f"={fy(y)}{R['pti']}*{P['etr']}" for y in F}})
    row("aff", "權益法投資損益（Technoprobe）", "US$M", {"C": "=a25_aff", "D": "=f26_aff", **{fy(y): "=aff" for y in F}})
    row("ni", "淨利", "US$M", {"C": "=a25_ni", "D": "=f26_ni", **{fy(y): f"={fy(y)}{R['pti']}-{fy(y)}{R['tax']}+{fy(y)}{R['aff']}" for y in F}}, note="FY2026＝錨點（上半年實績＋下半年指引）")
    row("sh_", "稀釋股數", "百萬股", {"C": "=a25_dsh", "D": "=sh26", **{fy(y): f"=sh26*(1-{P['bb']})^{y-2026}" for y in F}})
    row("eps", "EPS（GAAP 稀釋）", "US$", {"C": "=a25_eps", **{fy(y): f"={fy(y)}{R['ni']}/{fy(y)}{R['sh_']}" for y in (2026,) + F}}, fmt="0.00")
    row("ssh", "SoC 占營收", "%", {fy(y): f"={fy(y)}{R['rsoc']}/{fy(y)}{R['rev']}" for y in range(2025, 2030)}, fmt="0.0%")
    gadj = {2027: "", 2028: "+gm_adj_C28" if s == "C" else "", 2029: "+gm_adj_C29" if s == "C" else ""}
    row("gm", "毛利率（參考；不進 EPS）", "%", {"C": "=a25_gp/a25_rev", "D": "=f26_gm", **{fy(y): f"=f26_gm+(J13-a25_gp/a25_rev)/(Q!J4/Q!J10-a25_soc/a25_rev)*({fy(y)}{R['ssh']}-D{R['ssh']}){gadj[y]}".replace("J13", "Q!J13") for y in F}}, fmt="0.0%",
        note="FY2026＋斜率 ×（SoC 占營收 − FY2026）；斜率由 FY2025 與上半年 FY2026 兩點推得；EPS 由事業別稅前利益推導，本列只供參考")
    R["end"] = r
    return R

# ---------------- PL：三情境 ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 泰瑞達損益（US$M；FY2025 為實績，FY2026 為錨點（上半年實績＋下半年指引），FY2027–FY2029 為模型）。三情境各一個完整推導區塊（見下方），本頁頂端為摘要"
PB = {}; r0 = 44
TITLE = {"A": "情境 A（順利）：加速器顆數續增；市占維持", "B": "情境 B（平平）：2028 顆數 ×0.8、2029 ×0.85（Advantest 模型的 B 市場）",
         "C": "情境 C（不如預期）：2028–29 顆數 ×0.6；泰瑞達自 Advantest 搶回份額但價格壓力"}
for s in SC:
    PB[s] = engine(PL, r0, TITLE[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "US$M", "#,##0"), ("rsoc", "  SoC 測試機", "US$M", "#,##0"), ("rmem", "  記憶體測試機", "US$M", "#,##0"), ("rrob", "  機器人（不隨 AI 情境）", "US$M", "#,##0"),
                                ("gm", "毛利率（參考）", "%", "0.0%"), ("pti", "稅前利益", "US$M", "#,##0"), ("pm", "稅前利益率", "%", "0.0%"), ("ni", "淨利", "US$M", "#,##0"), ("eps", "EPS", "US$", "0.00")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in range(2025, 2030):
            PL[f"{FYC[y]}{sr}"] = f"={FYC[y]}{PB[s][key]}"; PL[f"{FYC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對：FY2026 EPS − 錨點（應為 0）").font = BOLD
PL[f"C{sr}"] = f"=ABS(D{SUMR[('A','eps')]}-f26_eps)+ABS(D{SUMR[('C','eps')]}-f26_eps)"; PL[f"C{sr}"].number_format = "0.0000"
PL.column_dimensions["A"].width = 50; PL.column_dimensions["B"].width = 12; PL.column_dimensions["H"].width = 60
for c in "CDEFG": PL.column_dimensions[c].width = 13

# ---------------- Val：P/E 估值 ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — 估值。目標價＝DCF 與 Comparable 加權（第 0 節，放在本頁後段）；規則值（十年年底 P/E 依循環位置取分位：見頂年 25 分位、中段中位數、谷底 75 分位）保留為參考。循環位置依設備股共用規則（第 4 節）"
V["A2"] = "P/E 定義（自算）：年底收盤價（macrotrends，股利調整後）÷ 當年度 GAAP 稀釋 EPS（會計年度＝日曆年）；FY2016 為虧損（EPS −0.21），P/E 無意義，不納入分位"
V.cell(4, 1, "1. 公司公布 EPS（GAAP 稀釋，美元）— 兩個來源").font = BOLD
hdr(V, 5, ["會計年度", "來源1：SEC XBRL（10-K）", "來源2：macrotrends", "兩者差距", "淨利 US$M（10-K）", "營收 US$M（10-K）", "股東權益 US$M（年底）", "說明"])
EPSD = [(2015, 0.97, 0.97, 206.5, 1639.6, 1965.8), (2016, -0.21, -0.21, -43.4, 1753.3, 1828.7), (2017, 1.28, 1.28, 257.7, 2136.6, 1953.6), (2018, 2.35, 2.35, 451.8, 2100.8, 1522.4),
        (2019, 2.60, 2.60, 467.5, 2295.0, 1480.2), (2020, 4.28, 4.28, 784.1, 3121.5, 2207.0), (2021, 5.53, 5.53, 1014.6, 3702.9, 2562.4), (2022, 4.22, 4.22, 715.5, 3155.0, 2451.3),
        (2023, 2.73, 2.73, 448.8, 2676.3, 2525.9), (2024, 3.32, 3.32, 542.4, 2819.9, 2819.3), (2025, 3.47, None, 554.0, 3190.0, 2795.8)]
EP = {}
for i, (y, a, b, ni, rv, eq) in enumerate(EPSD, 6):
    V.cell(i, 1, y)
    for c, v in zip("BCEFG", (a, b, ni, rv, eq)):
        if v is not None: V[f"{c}{i}"] = v; V[f"{c}{i}"].font = BLUE
    V[f"D{i}"] = f'=IF(C{i}="","",B{i}-C{i})'
    for c in "BCD": V[f"{c}{i}"].number_format = "0.00"
    for c in "EFG": V[f"{c}{i}"].number_format = "#,##0.0"
    V.cell(i, 8, "macrotrends 年表只到 2024；FY2025 第二來源：公司 4Q25 新聞稿 3.47（同一公司來源）" if y == 2025 else ("虧損年（Wireless Test 商譽減損）" if y == 2016 else ""))
    EP[y] = i
V.cell(17, 1, "全部年度用 GAAP 稀釋 EPS。淨利、營收、股東權益取自 SEC XBRL（10-K），用於十年平均淨利率與近三年 ROE").font = BOLD
V.cell(18, 1, "2. 10 年 P/E（年底）").font = BOLD
hdr(V, 19, ["年底", "收盤價 US$（股利調整）", "當年 EPS", "", "P/E 當年（自算）", "", "companiesmarketcap P/E", "自算 − cmc", "", "標記", "來源／說明"])
PXH = {2016: 24.1472, 2017: 40.1225, 2018: 30.3431, 2019: 66.4325, 2020: 117.3990, 2021: 160.6100, 2022: 86.1704, 2023: 107.5220, 2024: 125.2740, 2025: 193.4060}
CMC = {2016: -115, 2017: 31.2, 2018: 12.5, 2019: 23.6, 2020: 24.9, 2021: 26.0, 2022: 19.1, 2023: 36.9, 2024: 37.0, 2025: 55.5}
for i, y in enumerate(range(2016, 2026), 20):
    V.cell(i, 1, f"{y}-12")
    V.cell(i, 2, PXH[y]).font = BLUE; V.cell(i, 3, f"=B{EP[y]}")
    V.cell(i, 5, f'=IF(C{i}<=0,"",B{i}/C{i})')
    V.cell(i, 7, CMC[y]).font = BLUE
    V.cell(i, 8, f'=IF(E{i}="","",E{i}-G{i})')
    for c in "BC": V[f"{c}{i}"].number_format = "#,##0.00"
    for c in "EGH": V[f"{c}{i}"].number_format = "0.0"
    V.cell(i, 10, "Verified／Derived")
    V.cell(i, 11, {2016: "EPS 為負，P/E 不納入分位（cmc 列 −115）", 2025: "statmuse 2025-12-31 收盤 193.33（未調整，第二來源）；macrotrends 調整後 193.41"}.get(y, "收盤價：macrotrends 年表（年底收盤，股利調整後；泰瑞達 2016 年後無分割）"))
V.cell(31, 1, "收盤價來源1：macrotrends（股利調整後年底收盤）；來源2：statmuse（僅 2025 年底 193.33）。股利調整使早年價格略低（每季 0.06–0.13 美元股利），P/E 偏低約 1–3%。companiesmarketcap（TTM 口徑）僅作對照，差距見 H 欄。")
V.cell(33, 1, "3. 分位（只含正 EPS 年度）").font = BOLD
hdr(V, 34, ["統計", "P/E 當年（自算）", "", "cmc P/E（參考，正值）", "用途"])
for i, (nm, p, use) in enumerate((("25 分位", 0.25, "見頂年"), ("中位數", 0.5, "中段"), ("75 分位", 0.75, "谷底")), 35):
    V.cell(i, 1, nm); V.cell(i, 5, use)
    V.cell(i, 2, f"=PERCENTILE(E20:E29,{p})"); V.cell(i, 4, f"=PERCENTILE(G21:G29,{p})")
    for c in "BD": V[f"{c}{i}"].number_format = "0.0"
V.cell(38, 1, "年數"); V["B38"] = "=COUNT(E20:E29)"; V["D38"] = "=COUNT(G21:G29)"
defname("pe_peak", "Val!$B$35"); defname("pe_mid", "Val!$B$36"); defname("pe_trough", "Val!$B$37")
V.cell(40, 1, "4. 循環位置（設備股共用規則；見頂年度讀設備市場模型，不在本模型另訂）").font = BOLD
V.cell(41, 1, CP.rule_text(PEAK).replace("會計年度 FY(t) 對照日曆年 t（FY(t) 有 3/4 落在 t 年）", "會計年度 FY(t)＝日曆年 t"))
hdr(V, 42, ["情境", "設備市場見頂年度", "FY2027 對照日曆年", "判定", "倍數", "規則"])
for i, s in enumerate(SC, 43):
    V.cell(i, 1, s); V.cell(i, 2, f"=pk_{s}"); V.cell(i, 3, 2027).font = BLUE
    V.cell(i, 4, "=" + CP.excel_rule(f"C{i}", f"pk_{s}"))
    V.cell(i, 5, f'=IF(D{i}="見頂",pe_peak,IF(D{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 6, "tools/cycle_position.py")
V.cell(47, 1, "5. 隱含股價（美元）：規則值＝FY2027 EPS × 依循環位置的倍數；其餘倍數僅供參考").font = BOLD
hdr(V, 48, ["情境", "EPS 年度", "EPS", "倍數類型", "倍數", "隱含股價", "相對現價", "備註"])
vr = 49; VAL = {}
for s in SC:
    for y, typ, nm in ((2027, f"E{43 + SC.index(s)}", "rule"), (2027, "pe_peak", "參考：25 分位"), (2027, "pe_mid", "參考：中位數"), (2027, "pe_trough", "參考：75 分位"),
                       (2028, "pe_peak", "參考：25 分位"), (2028, "pe_mid", "參考：中位數"), (2028, "pe_trough", "參考：75 分位")):
        V.cell(vr, 1, s); V.cell(vr, 2, f"FY{y}"); V.cell(vr, 3, f"=PL!{FYC[y]}{SUMR[(s,'eps')]}").number_format = "0.00"
        V.cell(vr, 4, nm if nm != "rule" else f'="規則值："&D{43 + SC.index(s)}'); V.cell(vr, 5, f"={typ}").number_format = "0.0"
        V.cell(vr, 6, f"=C{vr}*E{vr}").number_format = "#,##0.00"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
        if nm == "rule":
            for c in range(1, 9): V.cell(vr, c).fill = KEY
            VAL[s] = vr
        vr += 1
V.cell(vr + 1, 1, "6. 目前遠期 P/E").font = BOLD
hdr(V, vr + 2, ["項目", "A（順利）", "B（平平）", "C（不如預期）", "備註"])
fr = vr + 3; FR = {}
for key, nm, fs, fm, note in (("px", "現價（2026-10-23 收盤）", ["=px"], "#,##0.00", "stockanalysis、equibles"),
                              ("sa", "遠期 P/E：stockanalysis", ["=sa_fpe"], "0.0", "分析師共識前瞻 EPS（non-GAAP）"),
                              ("g26", "現價 ÷ FY2026 EPS（錨點）", ["=px/f26_eps"], "0.0", "FY2026 三情境相同"),
                              ("pe27", "現價 ÷ FY2027 EPS（模型）", [f"=px/PL!{FYC[2027]}{SUMR[(s,'eps')]}" for s in SC], "0.0", ""),
                              ("pe28", "現價 ÷ FY2028 EPS（模型）", [f"=px/PL!{FYC[2028]}{SUMR[(s,'eps')]}" for s in SC], "0.0", ""),
                              ("c27", "現價 ÷ 共識 FY2027 EPS（non-GAAP）", ["=px/cons_eps27"], "0.0", "MarketBeat 11.30")):
    V.cell(fr, 1, nm)
    for j, f in enumerate(fs): V.cell(fr, 2 + j, f).number_format = fm
    V.cell(fr, 5, note); FR[key] = fr; fr += 1
# 舊法：規則值情境加權（參考；不是目標價）
tp0 = fr + 1
V.cell(tp0, 1, "舊法：規則值情境加權（參考）").font = BOLD
hdr(V, tp0 + 1, ["項目", "數值", "", "", "說明"])
V.cell(tp0 + 2, 1, "規則值情境加權價（12 個月；參考）"); V.cell(tp0 + 2, 2, f"=tp_wA*B{tp0 + 4}+tp_wB*B{tp0 + 5}+tp_wC*B{tp0 + 6}").number_format = "#,##0.00"
V.cell(tp0 + 2, 5, "權重 × 各情境隱含價（第 5 節規則值）")
V.cell(tp0 + 3, 1, "相對現價（規則值情境加權價 ÷ 現價 − 1）"); V.cell(tp0 + 3, 2, f"=B{tp0 + 2}/px-1").number_format = "0.0%"; V.cell(tp0 + 3, 5, "現價＝In 頁 px（2026-10-23 收盤）")
for k_, (s_, lab) in enumerate((("A", "情境 A 隱含價（順利）"), ("B", "情境 B 隱含價（平平）"), ("C", "情境 C 隱含價（不如預期）"))):
    rr_ = tp0 + 4 + k_
    V.cell(rr_, 1, lab); V.cell(rr_, 2, f"=F{VAL[s_]}").number_format = "#,##0.00"
    V.cell(rr_, 5, f'="FY2027 EPS × "&D{VAL[s_]}&" "&TEXT(E{VAL[s_]},"0.0")&"x（第 5 節 F{VAL[s_]}）"')
V.cell(tp0 + 7, 1, "情境權重 A／B／C")
for j_, s_ in enumerate("ABC"): V.cell(tp0 + 7, 2 + j_, f"=tp_w{s_}").number_format = "0%"
V.cell(tp0 + 7, 5, f'="合計 "&TEXT(B{tp0 + 7}+C{tp0 + 7}+D{tp0 + 7},"0%")&"；In 頁 tp_wA／tp_wB／tp_wC（Assumed，Andy 可改）"')
V.cell(tp0 + 8, 1, "規則值方法")
V.cell(tp0 + 8, 2, "P/E：各情境 FY2027 EPS × 十年（2016–2025 年底）P/E 依循環位置取分位（設備股共用規則：A、B 中段＝中位數，C 見頂＝25 分位；FY2016 虧損不納入，實為 9 年）。隱含價約為 FY2027 年底的價格、未折現")
TP0 = tp0
V.column_dimensions["A"].width = 46
for c in "BCDEFGHIJ": V.column_dimensions[c].width = 16
V.column_dimensions["K"].width = 80

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A）：對 FY2027、FY2028 營收與 EPS 的影響。每個案例是一個完整的推導區塊（本頁下方，與 PL 同一引擎）；FY2026 錨點不變"
hdr(SE, 3, ["案例", "設定", "FY2027 營收 US$M", "FY2027 EPS US$", "相對基準", "FY2028 營收 US$M", "FY2028 EPS US$", "相對基準"])
cases = [("泰瑞達 SoC 市占 −5 個百分點", "A 情境 FY2027 起", dict(sh="s_soc_lo")), ("泰瑞達 SoC 市占 +5 個百分點", "", dict(sh="s_soc_hi")),
         ("SoC 測試機市場 −15%（2027 起）", "測試對 WFE 比回到公司說的區間", dict(msoc="m_soc_lo")),
         ("記憶體市占 −5 個百分點", "", dict(shm="s_mem_lo")), ("記憶體市占 +5 個百分點", "", dict(shm="s_mem_hi")),
         ("測試事業增量利益率 40%", "固定成本重新推導（FY2026 錨點不變）", dict(inc="t_inc_lo", fix="t_fix_lo")), ("測試事業增量利益率 56%", "同上", dict(inc="t_inc_hi", fix="t_fix_hi")),
         ("Product Test 成長 低", "A 情境 4%", dict(gpt="g_pt_A_lo")), ("Product Test 成長 高", "A 情境 12%", dict(gpt="g_pt_A_hi")),
         ("機器人成長 0%", "三情境相同", dict(grob="g_rob_lo")), ("機器人成長 20%", "", dict(grob="g_rob_hi")),
         ("機器人利益率 低", "每年 −7 個百分點", dict(rm="_lo")), ("機器人利益率 高", "每年 +5 個百分點", dict(rm="_hi")),
         ("稅率 12%", "", dict(etr="etr_lo")), ("稅率 17%", "", dict(etr="etr_hi")),
         ("回購 0%／年", "", dict(bb="bb_lo")), ("回購 2%／年", "", dict(bb="bb_hi"))]
def put_case(i, name, setting, R, pre=""):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    SE[f"C{i}"] = f"={pre}{FYC[2027]}{R['rev']}"; SE[f"D{i}"] = f"={pre}{FYC[2027]}{R['eps']}"; SE[f"F{i}"] = f"={pre}{FYC[2028]}{R['rev']}"; SE[f"G{i}"] = f"={pre}{FYC[2028]}{R['eps']}"
    SE[f"E{i}"] = f"=D{i}/D$4-1"; SE[f"H{i}"] = f"=G{i}/G$4-1"
    for c in "CF": SE[f"{c}{i}"].number_format = "#,##0"
    for c in "DG": SE[f"{c}{i}"].number_format = "0.00"
    for c in "EH": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
SB = {}
base_row = 40
for n, (nm, setting, ov) in enumerate(cases):
    SB[nm] = engine(SE, base_row + n * 32, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm])
nn = 5 + len(cases)
put_case(nn, "情境 B（平平）", "Advantest 模型的 B 市場；Product Test 成長 B", PB["B"], "PL!")
put_case(nn + 1, "情境 C（不如預期）", "C 市場；市占自 Advantest 搶回、利益率下調", PB["C"], "PL!")
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：測試機市場（Advantest 模型）在 CY2027 已比 CY2026 高約四成，泰瑞達的 FY2027 主要取決於市場是否真的長到這個水準與市占；情境 B、C 的顆數係數從 2028 才生效，所以對 FY2027 無影響、對 FY2028 才明顯。")
SE.column_dimensions["A"].width = 40; SE.column_dimensions["B"].width = 46
for c in "CDEFGH": SE.column_dimensions[c].width = 15

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司指引、目標模型、共識、Advantest 份額、測試對 WFE 比。只作事後對照，不為對上數字而調整參數。差距 > 10% 的原因見 README"
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
pl = lambda key, y=2026: (lambda s: f"=PL!{FYC[y]}{SUMR[(s, key)]}")
RC.cell(rr_, 1, "A. FY2026 vs 公司指引與共識").font = BOLD; rr_ += 1
rec("營收：公司口述全年約 5.2B（3Q26 法說）", "=co_rev26", pl("rev"), "公司 3Q26 法說", "模型 FY2026＝前三季實績＋4Q26 指引中點；與公司口述一致性檢查")
rec("營收（共識）", "=cons_rev26", pl("rev"), "stockanalysis 2026-10-20")
rec("EPS（共識，non-GAAP）", "=cons_eps26", pl("eps"), "stockanalysis 2026-10-20", "模型為 GAAP；non-GAAP 另加權益法投資攤銷約 0.20、無形資產攤銷約 0.08／年（2Q26 新聞稿調節表）", fmt="0.00")
RC.cell(rr_, 1, "B. FY2027 vs 共識").font = BOLD; rr_ += 1
rec("EPS（共識，non-GAAP）", "=cons_eps27", pl("eps", 2027), "MarketBeat 2026-10-20", "差距大於 10% 的原因見 README", fmt="0.00")
RC.cell(rr_, 1, "C. 公司目標模型（ATE 市場 US$12–14B 時；未給年度）").font = BOLD; rr_ += 1
rec("營收 6.0B：模型 FY2027", "=tm_rev", pl("rev", 2027), "公司 4Q25 法說（fintool）", "目標模型未給年度；本列只看量級")
rec("non-GAAP EPS 上緣 11.00：模型 FY2027（GAAP）", "=tm_eps_hi", pl("eps", 2027), "同上", fmt="0.00")
RC.cell(rr_, 1, "D. 市占與市場規模").font = BOLD; rr_ += 1
rec("CY2025 SoC：泰瑞達隱含市占 vs 1 − Advantest 自估 66%", "=1-as_A_2026", lambda s: f"=PL!C{PB[s]['sh']}", "Advantest 自估（Interested-party）", "差額＝其他廠份額（CY2025 約 6.6%）", fmt="0.0%", pt=True)
rec("CY2026 SoC 市占 vs 分析師估約 37%", "=an_share", lambda s: f"=PL!D{PB[s]['sh']}", "2Q26 法說分析師提問（公司未證實）", "口徑不同：Advantest 的 SoC 市場含其自身所有 SoC 測試機；分析師口徑未明", fmt="0.0%", pt=True)
for y in (2026, 2027, 2028):
    rec(f"CY{y}：SoC＋記憶體測試機 ÷ WFE vs 公司長期上緣 9%", "=tw_hi", (lambda yy: (lambda s: f"=Units!{ycol[yy]}{UR[(s,'tw')]}"))(y), "公司 2Q26 法說（Interested-party）",
        "公司口徑為 ATE 市場（含類比等其他測試機）÷ WFE；本模型分子只有 SoC＋記憶體，口徑更窄卻已高於上緣" if y == 2027 else "", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "E. 其他資訊（非對帳）").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "上半年 FY2026 營收 ÷ 模型 FY2026 營收"); RC.cell(rr_, 3, f"=Q!J10/PL!D{SUMR[('A','rev')]}").number_format = "0.0%"; RC.cell(rr_, 10, "公司：上半年 50–52%"); rr_ += 1
RC.cell(rr_, 1, "指引紀錄：1Q26、2Q26 營收皆高於指引上緣（1,250）"); RC.cell(rr_, 3, "=(Q!G10+Q!H10)/2/1250-1").number_format = "0.0%"; RC.cell(rr_, 10, "指引偏保守 → FY2026 錨點可能偏低（下半年以指引中點推估）"); rr_ += 1
RECON_LAST = rr_ - 1
RC.column_dimensions["A"].width = 56
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 38; RC.column_dimensions["J"].width = 80

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司 SEC 原文在建置環境可開啟；法說數字取自逐字稿摘要網站，待以公司原始逐字稿複核）"; SO["A1"].font = BOLD
srcs = [("公司：2Q26 新聞稿（8-K Ex.99.1，2026-07-28）", "https://www.sec.gov/Archives/edgar/data/97210/000119312526321933/ter-ex99_1.htm"),
 ("公司：2Q26 10-Q（收入拆分、分部稅前利益）", "https://www.sec.gov/Archives/edgar/data/97210/000119312526327715/ter-20260628.htm"),
 ("公司：1Q26 新聞稿（2026-04-28）", "https://www.sec.gov/Archives/edgar/data/97210/000119312526188706/ter-ex99_1.htm"),
 ("公司：1Q26 10-Q", "https://www.sec.gov/Archives/edgar/data/97210/000119312526201058/ter-20260329.htm"),
 ("公司：4Q25 與 FY2025 新聞稿（2026-02-02）", "https://www.sec.gov/Archives/edgar/data/97210/000119312526034348/ter-ex99_1.htm"),
 ("公司：FY2025 10-K（收入拆分 2023–2025、分部稅前利益）", "https://www.sec.gov/Archives/edgar/data/97210/000119312526059002/ter-20251231.htm"),
 ("公司：3Q25 新聞稿與 10-Q", "https://www.sec.gov/Archives/edgar/data/97210/000119312525258477/ter-20250928.htm"),
 ("公司：1Q25、2Q25 新聞稿", "https://www.sec.gov/Archives/edgar/data/97210/000095017025099876/ter-ex99_1.htm"),
 ("SEC XBRL company facts（年度 EPS、淨利、營收、股東權益）", "https://data.sec.gov/api/xbrl/companyfacts/CIK0000097210.json"),
 ("公司：3Q26 財報日期公告（2026-10-27 盤後）", "https://investingnews.com/teradyne-to-announce-third-quarter-2026-results/"),
 ("earningswhispers：2Q26 法說逐字稿（測試對 WFE 比、份額、上半年占比）", "https://earningswhispers.com/transcript/TER/Q22026"),
 ("gurufocus：2Q26 法說摘要（SoC 843、記憶體 212、IST 67）", "https://www.gurufocus.com/news/8987636/teradyne-inc-ter-q2-2026-earnings-call-highlights-record-revenue-and-robust-ai-demand-propel-growth"),
 ("fintool：4Q25 法說摘要（目標模型、VIP 份額約 50%）", "https://fintool.com/news/teradyne-q4-ai-demand-surge"),
 ("stockanalysis：統計（現價、遠期 P/E、流通股）", "https://stockanalysis.com/stocks/ter/statistics/"),
 ("stockanalysis：預測（共識目標價、FY2026 共識）", "https://stockanalysis.com/stocks/ter/forecast/"),
 ("MarketBeat：3Q26 財報前瞻（FY2026／FY2027 共識 EPS、共識目標價 396.80）", "https://www.marketbeat.com/instant-alerts/upcoming-teradyne-ter-set-to-post-earnings-on-tuesday-2026-10-20/"),
 ("equibles：現價（第二來源）", "https://equibles.com/stocks/ter/price"),
 ("macrotrends：年底收盤價（股利調整）", "https://www.macrotrends.net/stocks/charts/TER/teradyne/stock-price-history"),
 ("macrotrends：年度稀釋 EPS（第二來源）", "https://www.macrotrends.net/stocks/charts/TER/teradyne/eps-earnings-per-share-diluted"),
 ("statmuse：2025-12-31 收盤 193.33", "https://www.statmuse.com/money/ask/teradyne-stock-price-at-end-of-each-year"),
 ("companiesmarketcap：年底 P/E（對照）", "https://companiesmarketcap.com/teradyne/pe-ratio/"),
 ("stockanalysis：同業遠期 P/E（Cohu、FormFactor；Advantest、KLA、AMAT 同站）", "https://stockanalysis.com/stocks/cohu/statistics/"),
 ("本 repo：advantest／tsmc／markets/equipment 成品（唯讀）", ""),
 ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 70; SO.column_dimensions["B"].width = 120

# ---------------- README 頁 ----------------
RD = wb.create_sheet("README", 0)
lines = [("泰瑞達（Teradyne，TER US）收支與估值模型 v0.3（AI 半導體 Project 步驟 5：設備與材料）　2026-10-31", True),
 ("模型命題：泰瑞達的盈餘由 AI 加速器與 HBM 的測試需求決定：SoC 測試機約三成份額、記憶體測試機約三成份額，乘上與 Advantest 同一份測試機市場；機器人事業另列、不隨 AI 情境。股價要看兩件事：測試機市場在 2027 年是否真的再長四成（公司自己說測試對 WFE 比 2027 年可能回落），以及 GPU 與第二家雲端客戶的份額能否在 2027 年起明顯增加。", False),
 ("", False),
 ("會計年度＝日曆年。金額為百萬美元，EPS 為 GAAP 稀釋（共識為 non-GAAP，差約每年 0.3）。FY2026 為錨點：上半年與 3Q26 實績（v0.3 滾入，2026-10-27 公布）＋4Q26 指引中點。", False),
 ("", False),
 ("驅動因子 → 推導量", True),
 ("SoC 測試機營收＝SoC 測試機市場（讀 Advantest 模型，CY，三情境同字母）× 泰瑞達市占（FY2026 錨點推得，約三成）；記憶體測試機＝記憶體測試機市場（Advantest 模型，HBM 位元與 DRAM 設備支出同字母）× 市占；IST 隨 SoC 市場；Product Test 前一年 × 成長率", False),
 ("測試事業稅前利益＝增量利益率 × 測試營收 − 固定成本（FY2025 實績與 FY2026 錨點兩點推導）；機器人營收與稅前利益另列（Assumed，三情境相同）；公司與沖銷、稅率、權益法投資損益沿用 FY2026", False),
 ("情境 A／B／C（順利／平平／不如預期）：測試機市場讀 Advantest 模型同字母（其顆數係數讀台積電模型：2028 ×1.0／0.8／0.6，2029 ×1.0／0.85／0.6）；泰瑞達 A、B 市占維持，C 與 Advantest 互補（Advantest FY2028 61%、FY2029 56% → 泰瑞達 +5／+10 個百分點）並承受同樣的價格壓力", False),
 ("Val：目標價＝DCF 與 Comparable 加權（第 0 節）；規則值（十年 P/E 分位依循環位置）保留參考。Recon：公司指引、目標模型、共識、Advantest 份額、測試對 WFE 比。Sens：市占、市場、增量利益率、機器人等", False), ("", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司 SEC 原文可開啟；法說數字取自逐字稿摘要網站，待複核。", False),
 ("已知限制：(1) 測試機市場整份取自 Advantest 模型（以 Advantest 自估的 CY2026 市場校準，Interested-party），CY2027 比 CY2026 高約四成，SoC＋記憶體 ÷ WFE 達一成，高於泰瑞達自己說的 7–9%；(2) 泰瑞達市占由 FY2026 錨點推得，錨點的 4Q26 用營收與 EPS 指引中點、非 SoC 事業 4Q＝3Q（Assumed）；(3) 利益率只有兩點，增量利益率 48% 偏高時 FY2027 以後高估；(4) 機器人營收與利益率全為 Assumed；(5) C 情境泰瑞達份額上升是與 Advantest 互補的結果，與「C 最不利」的公司變數方向相反（見 README）；(6) 模型為 GAAP，共識為 non-GAAP。", False),
 ("本版未做先驗，無 Andy 合理 P/E 對照。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

PEERS = [("Advantest", "6857.T", 37.69, "2026-10-21", "stockanalysis statistics：Forward PE（分析師共識；價格日期如左）"),
         ("科磊（KLA）", "KLAC", 26.65, "2026-10-23", "stockanalysis statistics：Forward PE（分析師共識；價格日期如左）"),
         ("Cohu", "COHU", 33.89, "2026-10-23", "stockanalysis statistics：Forward PE（近 12 月虧損，前瞻倍數偏高）"),
         ("FormFactor", "FORM", 28.17, "2026-10-23", "stockanalysis statistics：Forward PE（分析師共識；價格日期如左）"),
         ("應用材料（Applied Materials）", "AMAT", 21.39, "2026-10-23", "stockanalysis statistics：Forward PE（分析師共識；價格日期如左）")]
CFG_V2 = dict(cols=FYC, scn={"A": "順利", "B": "平平", "C": "不如預期"}, peers=PEERS, fwd={2027: "2/12", 2028: "10/12"},
    fwd_note="約 2027-10 時的未來 12 月（2027-11～2028-10）≈ 2／12 × FY2027 ＋ 10／12 × FY2028（會計年度＝日曆年）",
    cur="US$", unit_rev="US$M", unit_ni="US$M", fmt="#,##0.00", fmt_rev="#,##0", band_fmt="#,##0.00",
    roe=("FY2023–FY2025", "=AVERAGE(Val!E14/AVERAGE(Val!G13:G14),Val!E15/AVERAGE(Val!G14:G15),Val!E16/AVERAGE(Val!G15:G16))",
         "SEC XBRL（10-K）：淨利 ÷ 平均股東權益：FY2023 18.0%、FY2024 20.3%、FY2025 19.7%（Val 頁第 1 節）", "stockanalysis 近 12 月 ROE 36.46%（AI 期間，非三年平均）"),
    nim=("FY2016–FY2025", "=AVERAGE(" + ",".join(f"Val!E{EP[y]}/Val!F{EP[y]}" for y in range(2016, 2026)) + ")",
         "SEC XBRL（10-K）淨利 ÷ 營收（Val 頁第 1 節；含 FY2016 虧損）", "macrotrends EPS 逐年相同（第二來源）"),
    cons_tp=(446.47, "stockanalysis：17 位分析師平均 US$446.47（中位數 450；區間 350–550），資料更新 2026-10-20", "MarketBeat：396.80（11 買進、5 持有）；早於 2026-10-27 的 3Q26 財報"),
    cons_feps=("=cons_eps27", "共識前瞻 EPS＝MarketBeat FY2027 共識 11.30（non-GAAP）；比本模型 12 個月後前瞻 EPS（約 FY2028）早約一年，且為 non-GAAP，EPS 因素因此偏大", "Interested-party"),
    peer_note="測試與設備同業（券商常用比較組）：Advantest（直接對手）、KLA、Cohu、FormFactor、應用材料；不含 A 股",
    dcf_note="", pe10="E20:E29",
    pe10_note="現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 2 節，FY2016 虧損不納入）的分位；只作檢查，不進目標價",
    px_note="現價＝In 頁 px（2026-10-23 收盤）",
    method="DCF（v0.4 標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加（EBIT 以稅前利益代，股權報酬視為費用），以 WACC 折現到 2026-12-31（期中慣例），FY2027–32 明確期（FY2030 起收斂法），終值＝NOPAT(FY2033) ×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, rf)、RONIC＝WACC；EV 加淨現金與長期投資 ÷ 稀釋股數，以 ke 推到 12 個月後並扣除期間股利（DCF 頁、tools/fcff_dcf.py）；β 為自算產業 β（美國半導體設備同業中位數，Blume 調整）重新加槓桿（股權成本與 WACC 見 0.1 節）。Comparable：Advantest、KLA、Cohu、FormFactor、應用材料前瞻 P/E 中位數 × 12 個月後前瞻 EPS。偏誤：測試機市場取自 Advantest 模型，CY2027 起高於泰瑞達自己說的測試對 WFE 比，EPS 偏高估；十年平均淨利率（約一成八）低於 AI 期間，情境 C 收斂偏保守；同業前瞻 P/E 為 AI 熱度下的倍數。",
    readme=["估值 v2（2026-10-25）：Val 頁「0. 目標價」＝0.5 × DCF 目標價（FY2027–32 明確期＋永續，情境權重 A 0.3／B 0.5／C 0.2，情境價值下限 0）＋0.5 × Comparable 目標價；另列 Comparable 單獨目標價、市場共識目標價與差距拆解、歷史分位檢查。規則值情境加權保留為參考，不叫目標價。",
            "預估期間延伸到 FY2032：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（PL 頁底部延伸區塊，灰底）。"])

# ---------------- 估值 v2：PL 延伸到 FY2032（收斂法）＋ Val「0. 目標價」（DCF 與 Comparable 加權） ----------------
def valuation_v2(CFG):
    cols = CFG["cols"]; MF = CFG["fmt"]; YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
    FADEFILL = PatternFill("solid", fgColor="E7E6E6")
    r = IN.max_row + 2
    IN.cell(r, 1, "J. 目標價（DCF 與 Comparable 加權）與 FY2030–32 收斂法參數（黃底＝Assumed，Andy 可改）").font = BOLD; r += 1
    J = {}
    def inrow(key, item, unit, base, tag, s1, s2, fmt=None, yellow=False):
        nonlocal r
        IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
        c = IN.cell(r, 5, base)
        if not (isinstance(base, str) and base.startswith("=")): c.font = BLUE
        if fmt: c.number_format = fmt
        if yellow: c.fill = KEY
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
        wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}"); J[key] = r
        r += 1
    # v0.4：刪除 tp_r、tp_roe（舊股東現金流法 DCF 參數）；DCF 折現率見 In 頁 M2 節（tools/fcff_dcf.py）
    inrow("tp_g", "PL 延伸期營收年增的收斂目標（情境 B；FY2032；v0.4 起不再用於 DCF，DCF 永續成長率見 M2 節 g_tv）", "%", 0.04, "Assumed", "Andy 可改（工作單預設 4%）", "情境 A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt（PL 頁延伸區塊）", "0.0%", True)
    inrow("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", 0.5, "Assumed", "Andy 可改（2026-10-25 11:53：DCF／Comparable 各半）", "—", "0%", True)
    inrow("nim10", f"十年平均淨利率（{CFG['nim'][0]}，各年淨利 ÷ 營收的平均）", "%", CFG["nim"][1], "Derived", CFG["nim"][2], CFG["nim"][3], "0.0%")
    inrow("cons_tp", "市場共識目標價（賣方平均）", CFG["cur"], CFG["cons_tp"][0], "Interested-party", CFG["cons_tp"][1], CFG["cons_tp"][2], MF)
    inrow("cons_feps", "共識前瞻 EPS（差距拆解用）", CFG["cur"], CFG["cons_feps"][0], "Interested-party", CFG["cons_feps"][1], CFG["cons_feps"][2], MF)

    p0 = PL.max_row + 3
    PL.cell(p0, 1, "延伸到 FY2032：FY2025–FY2029 引用本頁頂端摘要；灰底 FY2030–FY2032 為收斂法，非驅動因子推導").font = BOLD
    PL.cell(p0 + 1, 1, "收斂法：營收年增由 FY2029 年增線性收斂到 FY2032 的長期成長（A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt）；淨利率由 FY2029 線性收斂（A＝FY2029、B＝FY2029 與十年平均的中點、C＝十年平均）；股數維持 FY2029（EPS 與淨利同比例）")
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
                mT = {"A": f"$G{R['nim']}", "B": f"($G{R['nim']}+nim10)/2", "C": "nim10"}[s]
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

    t0 = V.max_row + 3
    T = {"p0": p0}
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
    V.cell(rr, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）"); V[f"E{rr}"] = "模型無折舊與資本支出項目，故不用（淨利＋折舊－資本支出）；FY2025 資本支出 224.0、折舊 111.4（10-K 現金流量表），資本支出約為營收 7%，偏誤小"; T["cf_h"] = rr
    tset(rr + 1, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）")
    RV = rr + 1; CFR = {}
    for i, y in enumerate(YRS6):
        tset(rr + 2 + i, f"cf{y}", f"  FY{y}", [f"={c}{EPSR[y]}*(1-$B${RV})" for c in "BCD"], MF); CFR[y] = rr + 2 + i
    c1, c6 = CFR[2027], CFR[2032]
    tset(rr + 8, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=MAX(0,SUMPRODUCT({c}{c1}:{c}{c6},1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}{c6}*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6)" for c in "BCD"], MF,
         "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底，t＝1 對 FY2027；低於 0 時取 0（股東有限責任）" + CFG["dcf_note"])
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
    V["A3"] = f"目標價（DCF 與 Comparable 加權）：見本頁第 {t0} 列起「0. 目標價」"

    s0 = SE.max_row + 3
    SE.cell(s0, 1, "目標價參數（只改 DCF 參數與權重，PL 收斂法的長期成長不隨 tp_g 變動）").font = BOLD
    hdr(SE, s0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
    def dcf_formula(c, r_="tp_r", g="tp_g"):
        cf = lambda y: f"Val!{c}{EPSR[y]}*(1-MIN({g}/tp_roe,0.8))"
        pv = "+".join(f"{cf(y)}/(1+{r_})^{t}" for t, y in enumerate(YRS6, 1))
        return f"MAX(0,{pv}+{cf(2032)}*(1+{g})/({r_}-{g})/(1+{r_})^6)"
    cases = [("基準", "r 10%、g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "r 11%", dict(r_="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", dict(r_="(tp_r-0.01)")),
             ("長期成長 +1pt", "g 5%（永續與再投資率）", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", dict(g="(tp_g-0.01)")),
             ("DCF 權重 0.3", "Comparable 0.7", dict(w="0.3")), ("DCF 權重 0.7", "Comparable 0.3", dict(w="0.7"))]
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
    rr = RD.max_row + 1
    for t in CFG["readme"]:
        RD.cell(rr, 1, t); rr += 1
    return T, EXT, J

V2T, V2EXT, V2J = valuation_v2(CFG_V2)

# ---------------- v0.4（2026-11-05）：標準版 DCF（無槓桿自由現金流 FCFF＋WACC；tools/fcff_dcf.py）取代股東現金流法 ----------------
# 依據：docs/plan/個股DCF標準規格.md v1.0、工作單「AI半導體 個股 DCF 改為標準版」（Andy 2026-11-05 20:42、20:45、21:21）。
# 做法：Val 0.1 節（原股東現金流）整段改寫為 FD 摘要（列數不變，0.2、0.3 節位置不動）；原 0.1 節的情境 EPS 表移到新 DCF 頁，
#   Comparable 的 12 個月後前瞻 EPS 與歷史倍數檢查改讀 DCF 頁（數值不變）；Sens 頁目標價參數段改為 FD 標準案例。
import re as _re3
def dcf_v03(T, paths, years, frac0, ver, a1, a2, MF, band_fmt, conv_from=None, nc_bs=None, nc_note="", detail_title=None, usd=False):
    DW = wb.create_sheet("DCF")
    DW["A1"] = a1; DW["A2"] = a2
    DW.cell(4, 1, "情境 EPS（FY2027–FY2032；v0.2 起在 Val 0.1 節，" + ver + " 移到本頁；Val 的 12 個月後前瞻 EPS 與歷史倍數檢查讀本表，數值不變）").font = BOLD
    hdr(DW, 5, ["項目", "A", "B", "C", "說明"])
    emap = {}
    for i, y in enumerate((2027, 2028, 2029, 2030, 2031, 2032)):
        rv = T[f"eps{y}"]; r = 6 + i; emap[rv] = r
        DW.cell(r, 1, f"  FY{y}")
        for c in "BCD":
            DW[f"{c}{r}"] = V[f"{c}{rv}"].value; DW[f"{c}{r}"].number_format = MF
    DW.cell(6, 5, "PL 頁延伸區塊；FY2030 起為收斂法")
    rx = _re3.compile(r"(?<![A-Za-z_!$\d])\$?([BCD])\$?(\d+)(?![\d(])")
    for k in ("feps", "pr"):
        for c in "BCD":
            v = V[f"{c}{T[k]}"].value
            if isinstance(v, str):
                V[f"{c}{T[k]}"].value = rx.sub(lambda m: f"DCF!${m.group(1)}${emap[int(m.group(2))]}" if int(m.group(2)) in emap else m.group(0), v)
    R = FD.write_detail(DW, 14, paths, years, frac0, conv_from=conv_from, fmt=MF, stub=True,
                        title=detail_title or "三情境 FCFF 明細（金額單位同 PL 頁；年度一欄；期中慣例）")
    R["end"] = FD.write_checks(DW, R["end"], nc_bs=nc_bs, nc_note=nc_note) + 1
    DW.column_dimensions["A"].width = 60; DW.column_dimensions["B"].width = 40
    for c in "CDEFGH": DW.column_dimensions[c].width = 13
    r1, r2 = T["h01"], T["dcftp"]
    assert str(V.cell(r1, 1).value).startswith("0.1 DCF"), V.cell(r1, 1).value
    _ns = FD.summary_rows(stub=True)   # 第二輪修正：模組 v2 stub 期在「＋淨現金與投資」後多兩列，用原 0.1 節之後的空白列
    for rr in range(r2 + 1, r1 + _ns):
        assert all(V.cell(rr, cc).value is None for cc in range(1, 8)), (rr, V.cell(rr, 1).value)
    assert V.cell(r1 + _ns, 1).value is None, V.cell(r1 + _ns, 1).value
    r2 = max(r2, r1 + _ns - 1)
    for rr in range(r1, r2 + 1):
        for cc in range(1, 8):
            V.cell(rr, cc).value = None; V.cell(rr, cc).fill = PatternFill(); V.cell(rr, cc).font = Font()
    S = FD.write_summary(wb, V, r1, R, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt=MF,
                         title=f"0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；{ver}；明細見 DCF 頁）", beta_note="自算產業 β 重新加槓桿")
    V[f"B{T['tpw']}"] = f"=tp_w_dcf*B{S['dcf']}+(1-tp_w_dcf)*B{T['comptp']}"
    V[f"B{T['band']}"] = f'=TEXT(D{S["v12"]},"{band_fmt}")&"～"&TEXT(B{S["v12"]},"{band_fmt}")'
    s0, s1 = T["sens0"], T["sens_last"]
    for rr in range(s0, s0 + 16):
        for cc in range(1, 10):
            if rr > s1: assert SE.cell(rr, cc).value is None, (rr, cc, SE.cell(rr, cc).value)
            SE.cell(rr, cc).value = None; SE.cell(rr, cc).fill = PatternFill(); SE.cell(rr, cc).font = Font()
    ts = FD.write_sens(SE, s0, R, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref=f"Val!$B${T['comptp']}", base_ref=f"Val!$B${T['tpw']}", fmt=MF,
                       title=f"目標價參數敏感度（{ver} 標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）", beta_du="beta_du", usd=usd)
    SE.column_dimensions["G"].width = 14
    T.update(dcf=S["dcf"], v12=S["v12"], S=S, R=R, ts=ts)
    for k in ("eps_h", "cf_h", "reinv", "dcfv", "dcf12", "dcftp") + tuple(f"eps{y}" for y in range(2027, 2033)) + tuple(f"cf{y}" for y in range(2027, 2033)):
        T.pop(k, None)
    T["dcftp"] = S["dcf"]; T["sens_last"] = ts[1]
    # 第二輪修正：模組 v2 的 In 頁 M2 標題不寫版號（原本此處改寫版號的段落刪除）
    return T
_EXTC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2033)}   # PL 延伸區塊欄（valuation_v2 同式）
FY_X = list(range(2027, 2033))
def _ebit(s_, y):
    return f"PL!{FYC[y]}{PB[s_]['pti']}" if y <= 2029 else f"=(PL!{_EXTC[y]}{V2EXT[s_]['ni']}-aff)/(1-etr)"
DPATHS = [dict(name=s_, label=f"情境 {s_}（{CFG_V2['scn'][s_]}）", rev={y: f"PL!{_EXTC[y]}{V2EXT[s_]['rev']}" for y in FY_X}, rev0=f"PL!{_EXTC[2026]}{V2EXT[s_]['rev']}",
              ebit={y: _ebit(s_, y) for y in FY_X}, tax={y: "etr" for y in FY_X}, ebit_note="稅前利益（事業別稅前利益合計，不含權益法損益）；FY2030 起＝（淨利 − aff）÷（1 − etr）") for s_ in SC]
V2T = dcf_v03(V2T, DPATHS, FY_X, 1.0, "v0.4",
              "DCF — 標準版（v0.4）：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例；會計年度＝日曆年）；終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py",
              "EBIT：FY2027–FY2029 取 PL 推導區塊「稅前利益」（事業別稅前利益＋公司與沖銷；模型無營業利益列，以此代 EBIT，不含 Technoprobe 權益法損益）；FY2030–FY2032（PL 延伸區塊只有淨利）由淨利還原：EBIT＝（淨利 − aff）÷（1 − etr）。稅率：etr。"
              "權益法損益不進 NOPAT，長期投資另在評價時加入。折舊、資本支出：近三年占營收比例（In 頁 fc_da、fc_cx，各年不變）；營運資金：近三年平均占營收 × 營收增量。",
              CFG_V2["fmt"], CFG_V2["band_fmt"], nc_bs="460.74+5.85+676.84-(17.06+120.05)",
              nc_note="3Q26 資產負債表（US$M）：現金 460.74＋短期投資 5.85＋長期投資 676.84 − 租賃（17.06＋120.05）；無借款",
              detail_title="三情境 FCFF 明細（US$M；會計年度＝日曆年）")
if os.environ.get("TER_ROWS"):
    json.dump(dict(V2T=V2T, EXT=V2EXT, J=V2J, SUMR={f"{a}|{b}": v for (a, b), v in SUMR.items()}, PB=PB, VAL=VAL, FR=FR, TP0=TP0, INROW=INROW,
                   UR={f"{a}|{b}": v for (a, b), v in UR.items()}, SENS_LAST=SENS_LAST, RECON_LAST=RECON_LAST, EP=EP),
              open(os.environ["TER_ROWS"], "w"), ensure_ascii=False, indent=0)


# ---------------- v0.1.1：上檔空間級距、投資建議、重點觀察（Andy 2026-10-25 14:00、14:17）；各頁只在現有最後一列之後追加，不位移既有列 ----------------
_up = next(r for r in range(1, V.max_row + 1) if V.cell(r, 1).value == "上檔空間：加權目標價")
GRADE_ROW = V.max_row + 2; REC_ROW = GRADE_ROW + 1
V.cell(GRADE_ROW, 1, "上檔空間級距")
V.cell(GRADE_ROW, 2, f'=IF(B{_up}>0.5,"大於 50%",IF(B{_up}>0.25,"25%–50%",IF(B{_up}>=0,"0%–25%",IF(B{_up}>=-0.25,"0%～−25%",IF(B{_up}>=-0.5,"−25%～−50%","小於 −50%")))))')
V.cell(GRADE_ROW, 5, f"以「上檔空間：加權目標價」（B{_up}）分六級（Andy 2026-10-25 14:00）")
V.cell(REC_ROW, 1, "投資建議")
V.cell(REC_ROW, 2, f'=IF(B{_up}>0.5,"強力買進",IF(B{_up}>0.25,"買進",IF(B{_up}>=-0.25,"中立",IF(B{_up}>=-0.5,"放空","強力放空"))))')
V.cell(REC_ROW, 5, "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
OBS = ['目標價 US$374（DCF 230、Comparable 相對倍數調整後 518 各半），比現價 308.85（2026-10-23）高 21.0%，級距 0%–25%，投資建議維持中立。v0.3 滾入 3Q26（營收 1,306、SoC 676、記憶體 301、IST 95、GAAP EPS 2.17）與 4Q26 指引（營收 1,250–1,350、GAAP EPS 1.88–2.18），並改用 Siliconomics v0.9 的 CY2026 SoC 測試機市場 11.5；加權目標價 354 → 374（+5.5%）。DCF 仍比現價低 25.7%，Comparable 單獨 +67.7%，兩法分歧大。', 'SoC 市占由 29.5% 修為 26.7%：FY2026 SoC 營收 3,071（三季實績 2,400.8＋4Q26 推估 670），除以較大的市場 11.5。與 Advantest 合計：Advantest CY2026 SoC 市占 71.4%＋泰瑞達 26.7%＝98.1%（v0.2.1 為 66%＋29.5%＝95.5%，且舊市場 11.0 低於兩家實績合計 11.3）。公司法說：運算營收 3Q26 年增近五成，但運算市場擴大、對手在運算強，稀釋整體份額；第一家雲端廠下一波量產在 2027 年中、第二家 2027 下半年。', 'FY2026 GAAP EPS 9.11（共識 non-GAAP 9.09；公司全年營收約 5.2B，模型 5,218）。FY2027 EPS A 14.07，比共識 11.30 高 24.5%，主因仍是測試機市場 CY2027 比 CY2026 高約四成（16.2 對 11.5）。情境 C 的泰瑞達市占與 Advantest 互補（Advantest FY2028 61%、FY2029 56% 為絕對值，由 71.4% 起算降幅變大），FY2028 EPS A／B／C 為 17.08／14.59／14.48。下一個檢驗點是 4Q26 法說（2027-01 下旬，公司將更新 ATE 市場模型）。']   # 重點觀察（與 README.md 相同文字；數字取自重算後的 Val／PL／Sens 頁）
_rr = RD.max_row + 1
RD.cell(_rr, 1, 'v0.1.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（數值不變）。Val 頁末新增「上檔空間級距」「投資建議」兩列（Andy 2026-10-25 14:00、14:17 固定規則），README 頁末新增投資建議與重點觀察三點；既有儲存格位置與數值不變。')
RD.cell(_rr + 1, 1, f'="投資建議："&Val!B{REC_ROW}&"（依上檔空間固定規則）"').font = Font(bold=True)
for _j, _t in enumerate(OBS, 1): RD.cell(_rr + 1 + _j, 1, f"重點觀察 {_j}：" + _t)
RD.cell(RD.max_row + 2, 1, "v0.3（2026-10-31）：滾入 3Q26（Q 頁 I 欄；8-K 2026-10-27、法說 2026-10-28）與 4Q26 指引；FY2026 錨點改為上半年＋3Q26 實績＋4Q26 指引中點（原 2 × 3Q26 指引）。與 Advantest 合併對帳 SoC 測試機後，Siliconomics v0.9 的 CY2026 SoC 測試機市場 11.0 → 11.5，Advantest 市占改為 CY2026 推得 71.4%。泰瑞達 SoC 市占 29.5% → 26.7%、FY2026 EPS 8.78 → 9.11；加權目標價 354 → 374（+5.5%），投資建議維持中立。")
if os.environ.get("RATING_ROWMAP"):   # 除錯用：輸出新增列號（供撰寫 html.json）；CI 不設此變數
    print("ROWMAP", GRADE_ROW, REC_ROW, _rr, _rr + 1, _rr + 2)


order = ["README", "In", "Q", "Units", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "ter")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.2.1", "2026-10-31")  # v0.2.1：In 頁版面（個股模型原則第 8 節；數值不變）
# ---------------- v0.13（2026-11-04）：Sens「Wattonomics 錨定」案例（三層校準原則；決策 S37）。In 頁最後三列、Sens 頁最下方小表與案例區塊；既有格不動 ----------------
WTA_NOTE = "Wattonomics 錨定：以情境 B 為底，產業量 CY2027 起 × wta（Siliconomics SIF_WTAnchor_Base＝R4 比值）；FY2026 維持錨點。三層校準原則（Andy 2026-11-04）；個股基準暫不改，2027-01～02 雲端資本支出指引後再決定。"
WTA_MF = {2027: "wta_27", 2028: "wta_28", 2029: "wta_29"}
def wta_in_rows():
    r = IN.max_row + 2
    IN.cell(r, 1, f"Z. Wattonomics 錨定係數（敏感度用；Siliconomics {SIL.VERSION}）").font = BOLD; r += 1
    def _inp(key, item, unit, lo, base, hi, tag, s1, s2):
        nonlocal r
        for j, v in enumerate((key, item, unit, lo, base, hi, tag, s1, s2), 1):
            if v is not None: IN.cell(r, j, v)
        IN.cell(r, 5).number_format = "0.000"
        wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
        r += 1
    SIL.wta_rows(_inp)
    for j, y in enumerate((2027, 2028, 2029), 1):   # Advantest 模型情境 B 的 AI SoC 測試機市場（CY；唯讀）
        _inp(f"adv_aiB_{y % 100}", f"Advantest 模型 AI SoC 測試機市場 CY{y}（B；Wattonomics 錨定案例用）", "US$B", None, ADVAI["B"][0][j], None, "Derived", ADVAI["B"][1] if j == 1 else "同上", "只用於 Sens 頁 Wattonomics 錨定案例")
def wta_sens():
    t0 = SE.max_row + 3; b0 = t0 + 11
    RW = engine(SE, b0, "案例區塊：Wattonomics 錨定（情境 B；FY2027 起 AI SoC 測試機市場 × wta_27／wta_28／wta_29）", scn="B", wm=WTA_MF)
    SE.cell(t0, 1, "Wattonomics 錨定（三層校準原則）").font = BOLD
    for c in range(1, 8 + 1): SE.cell(t0 + 1, c, SE.cell(3, c).value).font = BOLD
    put_case(t0 + 2, "基準（情境 A；同第 4 列）", "PL 區塊 A", PB["A"], "PL!")
    put_case(t0 + 3, "情境 B（本案例的底）", "Advantest 模型的 B 市場；Product Test 成長 B", PB["B"], "PL!")
    put_case(t0 + 4, "Wattonomics 錨定", "情境 B；FY2027 起 SoC 測試機市場中的 AI 部分 × wta_27／wta_28／wta_29（非 AI 部分不變）", RW)
    SE.cell(t0 + 5, 1, "差異%：Wattonomics 錨定 ÷ 基準 − 1"); SE.cell(t0 + 6, 1, "差異%：Wattonomics 錨定 ÷ 情境 B − 1")
    for c in "CDFG":
        for rr, rb in ((t0 + 5, t0 + 2), (t0 + 6, t0 + 3)):
            SE[f"{c}{rr}"] = f"=IFERROR({c}{t0 + 4}/{c}{rb}-1,\"\")"; SE[f"{c}{rr}"].number_format = "0.0%"
    SE.cell(t0 + 7, 1, WTA_NOTE)
    SE.cell(t0 + 8, 1, "SoC 測試機市場＝Advantest 模型（含 AI 與非 AI）；本案例只把 AI 部分（Advantest 模型 AI SoC 測試機市場，In 頁 adv_aiB_*）× wta：市場＋（wta−1）× AI 部分。FY(t)＝CY(t)；記憶體測試機市場不乘（HBM 出貨受供給限制）。")
    SE.cell(t0 + 9, 1, "目標價：本模型 Sens 目標價參數表只動 DCF 參數，不逐案例重算 DCF，本案例不列目標價。")
    return RW
wta_in_rows(); WTA_R = wta_sens()   # v0.13：放在 In 版面套用之後，確保在 In 頁最後
OLDV = (("dcf", "v0.3 DCF 目標價（舊法：股東現金流）", 232.48770315055, "#,##0.00"), ("tp", "v0.3 目標價（DCF 與 Comparable 加權）", 377.863103090163, "#,##0.00"),
        ("up", "v0.3 上檔空間：加權目標價", 0.223451847466936, "0.0%"), ("rec", "v0.3 投資建議", "中立", "@"))
A1_FROM = "v0.3"
# ---------------- v0.4（2026-11-05）：README 新增「展示層」「命題回覆」兩段（個股 HTML 版面規格 v1.0；公式帶數字）----------------
# 新列放在 README 既有內容之後，不位移既有列；README 以外的工作表不動
def _rw(ws, text, exact=False, col=1, start=1):
    for r_ in range(start, ws.max_row + 1):
        v_ = ws.cell(r_, col).value
        if isinstance(v_, str) and ((v_.strip() == text) if exact else v_.strip().startswith(text)):
            return r_
    raise KeyError((ws.title, text))
def _F(*parts):   # 文字公式：字串常數每段 ≤ 120 字（Excel 公式內字串常數上限 255 字元）；("=", 運算式) 為公式片段
    out = []
    for p_ in parts:
        if isinstance(p_, tuple):
            out.append(p_[1])
        else:
            s_ = str(p_)
            out += ['"' + s_[i_:i_ + 120].replace('"', '""') + '"' for i_ in range(0, len(s_), 120)]
    return "=" + "&".join(out)
_E = lambda e: ("=", e)
_TX = lambda c, f: _E(f'TEXT({c},"{f}")')
_SG = lambda c: _TX(c, "+0.0%;-0.0%")
_PC = lambda c: _TX(c, "0.0%")
_UP = lambda c: _E(f'IF({c}>=0,"高 ","低 ")&TEXT(ABS({c}),"0.0%")')
_ws_in = wb["In"]
_IN = {_ws_in.cell(r_, 1).value: r_ for r_ in range(4, _ws_in.max_row + 1) if _ws_in.cell(r_, 1).value}
I_ = lambda k: f"In!E{_IN[k]}"
I_lo = lambda k: f"In!D{_IN[k]}"
I_hi = lambda k: f"In!F{_IN[k]}"
_V, _RC, _SE, _RD = wb["Val"], wb["Recon"], wb["Sens"], wb["README"]
vw, vd, vc = _rw(_V, "目標價（DCF 與 Comparable 加權）", exact=True), _rw(_V, "DCF 目標價（情境加權）"), _rw(_V, "Comparable 目標價", exact=True)
vca, vk, vm = _rw(_V, "Comparable 目標價（相對倍數調整後）"), _rw(_V, "相對倍數調整係數"), _rw(_V, "同業前瞻 P/E 中位數")
vup, vuc, vcons, vrec = _rw(_V, "上檔空間：加權目標價"), _rw(_V, "上檔空間：Comparable"), _rw(_V, "市場共識目標價"), _rw(_V, "投資建議", exact=True)
vpx = _rw(_V, "現價（")
vfe = _rw(_V, "12 個月後的前瞻 EPS")
vd12 = _rw(_V, "DCF 情境價值推到 12 個月後")
e0 = _rw(wb["DCF"], "情境 EPS（FY2027") + 1   # 情境 EPS 表在 DCF 頁
E27, E28, E29, E32 = e0 + 1, e0 + 2, e0 + 3, e0 + 6
W_ = I_("tp_w_dcf")
# 加權目標價相對現價的差距拆解（補充規則 B）：DCF 貢獻＝w ×（DCF − 現價），Comparable 貢獻＝(1 − w) ×（調整後 Comparable − 現價）
SH_DCF = f"{W_}*(Val!B{vd}-Val!B{vpx})/(Val!B{vw}-Val!B{vpx})"
T2 = lambda c: _TX(c, "0.00")
T0 = lambda c: _TX(c, "#,##0.00")
CT_DCF = f"{W_}*(Val!B{vd}-Val!B{vpx})/Val!B{vpx}"; CT_CMP = f"(1-{W_})*(Val!B{vca}-Val!B{vpx})/Val!B{vpx}"
r_e27 = _rw(_RC, "EPS（共識，non-GAAP）", start=_rw(_RC, "B. FY2027")); r_tm = _rw(_RC, "non-GAAP EPS 上緣"); r_w27 = _rw(_RC, "CY2027：SoC＋記憶體測試機 ÷ WFE")
s_shm, s_shp, s_mk, s_b, s_c = _rw(_SE, "泰瑞達 SoC 市占 −5 個百分點"), _rw(_SE, "泰瑞達 SoC 市占 +5 個百分點"), _rw(_SE, "SoC 測試機市場 −15%"), _rw(_SE, "情境 B（平平）"), _rw(_SE, "情境 C（不如預期）")
s_w8 = _rw(_SE, "DCF 權重 0.3／0.7"); s_wacc = _rw(_SE, "WACC ±1pt"); s_beta = _rw(_SE, "β 低／高")
v_tv, v_mx = _rw(_V, "其中：終值現值占 EV"), _rw(_V, "隱含終值 EV／EBITDA")
_ro = _RD.max_row + 2
_RD.cell(_ro, 1, "v0.3 → v0.4 對照：舊法（股東現金流 DCF，折現率 10%）的目標價（常數，取自 v0.3 重算後 xlsx 的 Val 頁；只作對照）").font = Font(bold=True)
OLD = {}
for k_, (key_, lab_, v_, f_) in enumerate(OLDV, 1):
    _RD.cell(_ro + k_, 1, lab_); _c = _RD.cell(_ro + k_, 2, v_); _c.font = BLUE; _c.number_format = f_; OLD[key_] = _ro + k_
lead = _F("命題：泰瑞達的盈餘由 AI 加速器與 HBM 的測試需求決定——與 Advantest 同一份測試機市場，乘上泰瑞達的 SoC 市占約 ", _TX("s_soc", "0%"), "、記憶體市占約 ", _TX("s_mem", "0%"),
          "；機器人事業另列、不隨 AI 情境。問題是測試機市場在 2027 年能否再長約四成、份額能否提升，以及現價是否已經反映。三個問題如下。")
props = [
    "命題 1：2027 年的測試機市場撐得起模型的盈餘嗎——模型 FY2027 每股盈餘相對共識與公司目標模型高多少，測試機占晶圓前段設備的比例是否超過公司說的長期上緣。",
    "命題 2：份額與市場規模的風險有多大——SoC 市占少 5 個百分點、或 SoC 測試機市場少 15% 時，每股盈餘少多少。",
    "命題 3：每股值多少——以 DCF 與同業倍數加權的目標價，相對現價的上檔空間是多少。",
]
obs = [
    _F("重點 1：加權目標價高於現價，投資建議", _E(f"Val!B{vrec}"), "：加權目標價約 ", T0(f"Val!B{vw}"), " 美元，比現價 ", T0(f"Val!B{vpx}"), " ", _UP(f"Val!B{vup}"),
       _E(f'IF(ABS(Val!B{vup})<=0.25,"（在正負 25% 之間為中立）",IF(Val!B{vup}<0,"（−25%～−50% 為放空）","（+25%～+50% 為買進）"))'),
       "。拆解：相對倍數調整後的 Comparable（", T0(f"Val!B{vca}"), "，權重 ", _TX(f"1-{W_}", "0%"), "）", _E(f'IF({CT_CMP}>=0,"拉高 ","拉低 ")&TEXT(ABS({CT_CMP}),"0.0%")'),
       "，DCF（", T0(f"Val!B{vd}"), "）", _E(f'IF({CT_DCF}>=0,"拉高 ","拉低 ")&TEXT(ABS({CT_DCF}),"0.0%")'), "：上檔全部來自 Comparable，標準版 DCF 低於現價（見重點 4）。"),
    _F("重點 2：模型盈餘明顯高於共識：FY2027 每股盈餘順利／平平／不如預期 ", T2(f"DCF!B{E27}"), "／", T2(f"DCF!C{E27}"), "／", T2(f"DCF!D{E27}"), " 美元，比共識 ", T2(f"Recon!B{r_e27}"),
       " 高 ", _PC(f"Recon!D{r_e27}"), "／", _PC(f"Recon!F{r_e27}"), "／", _PC(f"Recon!H{r_e27}"), "（本模型為 GAAP、共識為 non-GAAP，同一年度比較），也高於公司目標模型的 non-GAAP 每股盈餘上緣 ",
       T2(f"Recon!B{r_tm}"), "。差距來自模型讀的測試機市場：CY2027 測試機占晶圓前段設備 ", _PC(f"Recon!C{r_w27}"), "，高於公司說的長期上緣 ", _PC(f"Recon!B{r_w27}"), "。"),
    _F("重點 3：單一輸入中最敏感的是市占與市場規模（以順利情境為基準，Sens 頁）：SoC 市占 −5 個百分點使 FY2027 每股盈餘 ", _SG(f"Sens!E{s_shm}"), "；SoC 測試機市場 −15%（2027 起）使 ",
       _SG(f"Sens!E{s_mk}"), "；情境 B、C 的差異從 FY2028 才出現（FY2028 每股盈餘 ", _SG(f"Sens!H{s_b}"), "／", _SG(f"Sens!H{s_c}"), "）。"),
    _F("重點 4：DCF 改為標準版（無槓桿自由現金流折現）：推到 12 個月後的 DCF 情境價值順利／平平／不如預期為 ", T0(f"Val!B{vd12}"), "／", T0(f"Val!C{vd12}"), "／", T0(f"Val!D{vd12}"),
       "，DCF 目標價 ", T0(f"Val!B{vd}"), "（舊法股東現金流 ", T0(f"README!B{OLD['dcf']}"), "）。WACC ", _TX("wacc", "0.0%"), "（股權成本 ", _TX("ke", "0.0%"), "＝rf ", _TX("rf", "0.00%"),
       " ＋ β ", _TX("beta", "0.00"), " × ERP ", _TX("erp", "0%"), "；β 為自算產業 β（美國半導體設備同業中位數，Blume 調整）重新加槓桿，ERP 5% 與 RONIC＝WACC 為假設），高於舊法的 10%。終值現值占平平情境企業價值 ",
       _TX(f"Val!C{v_tv}", "0%"), "、隱含終值 EV／EBITDA ", _TX(f"Val!C{v_mx}", "0.0"), " 倍。以情境加權目標價為基準，WACC −1pt／＋1pt 時加權目標價 ", _SG(f"Sens!D{s_wacc}"), "／", _SG(f"Sens!F{s_wacc}"),
       "；「", _E(f"Sens!A{s_w8}"), "」時 ", _SG(f"Sens!D{s_w8}"), "／", _SG(f"Sens!F{s_w8}"), "。"),
]
answers = [
    _F("2027 年的測試機市場撐得起模型的盈餘嗎？答：模型假設撐得起，但比公司自己的說法樂觀。FY2027 每股盈餘三情境 ", T2(f"DCF!B{E27}"), "／", T2(f"DCF!C{E27}"), "／", T2(f"DCF!D{E27}"),
       " 美元，比共識 ", T2(f"Recon!B{r_e27}"), " 高 ", _PC(f"Recon!D{r_e27}"), "～", _PC(f"Recon!H{r_e27}"), "（GAAP 對 non-GAAP）；模型 CY2027 的 SoC＋記憶體測試機市場占晶圓前段設備 ", _PC(f"Recon!C{r_w27}"),
       "，高於公司說的長期上緣 ", _PC(f"Recon!B{r_w27}"), "（公司稱 2027 年此比例可能回落）。測試機市場讀 Advantest 模型，是假設。"),
    _F("份額與市場規模的風險有多大？答：以順利情境為基準（Sens 頁），SoC 市占 −5 個百分點使 FY2027／FY2028 每股盈餘 ", _SG(f"Sens!E{s_shm}"), "／", _SG(f"Sens!H{s_shm}"),
       "，+5 個百分點使 ", _SG(f"Sens!E{s_shp}"), "／", _SG(f"Sens!H{s_shp}"), "；SoC 測試機市場 −15% 使 ", _SG(f"Sens!E{s_mk}"), "／", _SG(f"Sens!H{s_mk}"), "。市占 ±5 個百分點與市場 −15% 是敏感度假設。"),
    _F("每股值多少？答：加權目標價約 ", T0(f"Val!B{vw}"), " 美元，比現價 ", T0(f"Val!B{vpx}"), " ", _UP(f"Val!B{vup}"), "，投資建議", _E(f"Val!B{vrec}"),
       "。加權目標價＝DCF ", T0(f"Val!B{vd}"), " 與 Comparable 各占 ", _TX(W_, "0%"), "／", _TX(f"1-{W_}", "0%"), "；進加權的 Comparable 是相對倍數調整後的 ", T0(f"Val!B{vca}"),
       "（同業前瞻本益比中位數 ", _TX(f"Val!B{vm}", "0.0"), " 倍 × 相對倍數調整係數 ", _TX(f"Val!B{vk}", "0.00"), " × 12 個月後前瞻每股盈餘 ", T2(f"Val!B{vfe}"), "，上檔空間 ", _SG(f"Val!B{vca}/Val!B{vpx}-1"),
       "）。只看 Comparable 目標價（市場主流同業法，不做相對倍數調整）為 ", T0(f"Val!B{vc}"), "（上檔空間 ", _SG(f"Val!B{vuc}"), "），頁首橫幅中間格放這個值，用來對照市場主流的目標價。市場共識目標價 ",
       T0(f"Val!B{vcons}"), "。"),
]
# 第二輪修正（2026-11-05）：展示層重點——淨現金滾到評價日（stub）、美元口徑折現（非美國公司）、終值占比 > 80% 標示（公式判斷；Sens 以加權目標價為基準）
def _r2pt(n, usd, ccy_fmt="#,##0.00", amt_fmt="#,##0.0", unit=""):
    _Vw, _Sw = wb["Val"], wb["Sens"]
    def _f(ws, lab, start=1):
        return next(r for r in range(start, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(lab))
    h = _f(_Vw, "0.1 DCF"); nc, st, ncr, tv = (_f(_Vw, l, h) for l in ("＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "其中：終值現值占 EV"))
    t = lambda c, f: f'TEXT({c},"{f}")'
    s = (f'="重點 {n}：淨現金由資產負債表日 "&bs_date&" 滾到評價日 2026-12-31：資產負債表淨現金與投資 "&{t(f"Val!C{nc}", amt_fmt)}&" ＋ stub 期 FCFF（情境 B 第一年 FCFF × "&TEXT(stub_frac,"0.000")&" 年）"&{t(f"Val!C{st}", amt_fmt)}'
         f'&" − 已宣告股利 "&{t("stub_div", amt_fmt)}&" ＝ "&{t(f"Val!C{ncr}", amt_fmt)}&"{unit}。"')
    if usd:
        u = _f(_Sw, "美元口徑折現")
        s += (f'&"美元口徑折現（Sens 案例：rf 改用美國 10 年期 "&TEXT(rf_us,"0.00%")&"、g＝MIN("&TEXT(g_cap,"0%")&", "&TEXT(rf_us,"0.00%")&")＝"&TEXT(MIN(g_cap,rf_us),"0.0%")&"，CRP、匯率與現金流不變）時，加權目標價 "'
              f'&{t(f"Sens!C{u}", ccy_fmt)}&"，相對基準加權目標價 "&TEXT(Sens!D{u},"+0.0%;-0.0%")&"；基準採當地 10 年期利率 "&TEXT(rf,"0.00%")&"＋固定匯率。"')
    s += (f'&IF(MAX(Val!B{tv}:D{tv})>0.8,"終值現值占 EV 超過 80% 的情境："&IF(Val!B{tv}>0.8,"A "&TEXT(Val!B{tv},"0%")&" ","")&IF(Val!C{tv}>0.8,"B "&TEXT(Val!C{tv},"0%")&" ","")&IF(Val!D{tv}>0.8,"C "&TEXT(Val!D{tv},"0%"),"")&"（DCF 價值大多來自終值，對 WACC 與 g 較敏感）。",'
          f'"三情境終值現值占 EV "&TEXT(Val!B{tv},"0%")&"／"&TEXT(Val!C{tv},"0%")&"／"&TEXT(Val!D{tv},"0%")&"（A／B／C），皆未超過 80%。")')
    return s
obs.append(_r2pt(len(obs) + 1, False, amt_fmt="#,##0", unit="（US$M）"))
rr = _RD.max_row + 2
_RD.cell(rr, 1, "展示層：命題摘要、命題分項與白話重點（HTML 用；v0.4 新增，上方各列保留）").font = Font(bold=True); rr += 1
for t_ in [lead] + props + obs:
    _RD.cell(rr, 1, t_); rr += 1
rr += 1
_RD.cell(rr, 1, "命題回覆（HTML 用；依本模型計算回答開頭的命題）").font = Font(bold=True); rr += 1
for t_ in answers:
    _RD.cell(rr, 1, t_); rr += 1
rr += 1
_RD.cell(rr, 1, "HTML 橫幅附註用：Comparable 目標價（相對倍數調整後）的上檔空間（＝Val 相對倍數調整後 Comparable 目標價 ÷ 現價 − 1；進加權目標價的是該值，橫幅中間格則放未調整的只看 Comparable 目標價）")
_RD.cell(rr, 2, f"=Val!B{vca}/Val!B{vpx}-1").number_format = "0.0%"
_RD["A1"] = (str(_RD["A1"].value).replace(A1_FROM, "v0.4", 1) + "；v0.4（2026-11-05）：DCF 改為標準版（無槓桿自由現金流 FCFF＋WACC；DCF 頁、Val 0.1 節、Sens 目標價參數）＋HTML 新版面 v1.0（README 展示層與命題回覆）；"
             "Comparable、情境權重、DCF 權重 50% 與投資建議規則不變")
_RD.cell(_RD.max_row + 2, 1, "v0.4（2026-11-05）：(1) DCF 改為標準版（docs/plan/個股DCF標準規格.md v1.0；tools/fcff_dcf.py）：Val 0.1 節改為 FCFF＋WACC 摘要（列數不變）；新增 DCF 頁（三情境 FCFF 明細與檢查列；原 Val 0.1 節的情境 EPS 表移入，Comparable 讀此表，數值不變）；"
         "In 頁新增 M2 節（rf 美國 10 年期公債、Damodaran 美國半導體設備產業 β 重新加槓桿、ERP 5%、g＝MIN(3%, rf)、RONIC＝WACC、資本結構、折舊／資本支出／營運資金比例）；刪除 tp_r、tp_roe；Sens 頁目標價參數段改為標準案例。"
         "(2) HTML 新版面 v1.0：README 新增「展示層」「命題回覆」兩段（公式帶數字）。新舊目標價對照見本頁「v0.3 → v0.4 對照」列。")
# v0.4.1（2026-11-05）：β 改自算產業 β（US 半導體設備同業中位數、Blume 調整；Andy 2026-11-05 22:55），並依 Advantest v0.7.1 成品重建，其他模型不變
_RD["A1"] = str(_RD["A1"].value) + "；v0.4.1（2026-11-05）β 改自算產業 β（US 半導體設備同業中位數、Blume 調整；Andy 2026-11-05 22:55），並依 Advantest v0.7.1 成品重建，其他模型不變"
_RD.cell(_RD.max_row + 2, 1, "v0.4.1（2026-11-05）：β 由 Damodaran 美國產業表改為自算產業 β——US 半導體設備同業 7 家（AMAT、LRCX、KLAC、TER、FORM、COHU、KLIC）Yahoo Finance 月報酬 2021-11～2026-10 對 ^GSPC 迴歸、Hamada 去槓桿、現金修正後的中位數，Blume 調整 0.67β＋0.33，再依本公司 D/E 與稅率重新加槓桿（In 頁 beta）；Sens β 低／高改為中位數 × 0.8 再 Blume／中位數未調整。上游 Advantest 成品改為 v0.7.1（只改 β；本模型讀的市場與市占數值不變）。評價日 2026-12-31 不變，其他模型不變")
# 第二輪修正（2026-11-05；工作單 2026-11-05 23:20，Andy 23:16 定案）：共用模組 tools/fcff_dcf.py v2
_RD["A1"] = str(_RD["A1"].value) + "；v0.4.2（2026-11-05）第二輪修正：換用 tools/fcff_dcf.py v2——淨現金由 3Q26 資產負債表日 2026-09-27 滾到評價日 2026-12-31（stub 期 FCFF；stub 期無已宣告股利）、β 敏感度改 β_u ± 0.3 重新加槓桿；上游 Advantest 改讀 v0.7.2 成品"
_RD.cell(_RD.max_row + 2, 1, "v0.4.2（2026-11-05）：第二輪修正（工作單 2026-11-05 23:20；Andy 2026-11-05 23:16 定案），換用共用模組 tools/fcff_dcf.py v2。(1) Stub 期：資產負債表日 2026-09-27（3Q26 10-Q，最新已公布一季）≠ 評價日 2026-12-31，淨現金滾到評價日＝資產負債表淨現金與投資＋各情境 FY2027 FCFF（第一個明確期年度，代表 2026 第四季）× stub_frac − 已宣告股利（0：最近一次 2026-09-25 已發放，12 月發放的下一次建置日尚未宣告；買回無已宣告金額，不扣）；In 頁新增 bs_date、stub_frac、stub_div，Val 0.1 節多「＋Stub 期 FCFF」「淨現金滾到評價日」兩列（0.2 節以後位置不動）。(2) β：In 頁新增 beta_u（1.503）、beta_de、beta_du（0.3），beta＝beta_u 重新加槓桿；Sens β 低／高改為 β_u ∓ 0.3 各自重新加槓桿。(3) ke 說明格由模組參數提供；In 頁 M2 標題不寫版號。美國公司不加「美元口徑折現」案例（＝基準）。(4) 上游 Advantest 成品改讀 v0.7.2（第二輪修正；本模型讀的 PL／In 市場數值不變）。資本支出、折舊：PL 沒有逐年列，維持占營收比例。展示層新增重點 5（stub 淨現金、終值占比是否超過 80%，公式帶數字）。PL、Comparable、情境權重不變。")

FD.yellow(IN)   # v0.4：M2 節判斷參數黃底
wb.save(sys.argv[1])

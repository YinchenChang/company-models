# Advantest（6857.T）收支與估值模型 v0.3（AI 半導體 Project 步驟 5：設備與材料——第一家公司）
# 用法：python3 advantest/build_advantest.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度 4 月到隔年 3 月：FY2026＝2026-04～2027-03。金額：日圓十億（¥B）為主，另列美元（US$B）。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
import market as MEM
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP

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
TSMC_F = only("tsmc/dist/*.xlsx"); SKH_F = only("skhynix/dist/*.xlsx"); EQ_F = only("markets/equipment/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)

ws_ = openpyxl.load_workbook(TSMC_F, data_only=True)["In"]
TS = {}   # 台積電模型 In 頁：情境顆數係數與 2029 成長
for r in ws_.iter_rows(min_row=4):
    k = r[0].value
    if k and (k.startswith("um2") or k.startswith("g29_") or k.startswith("u_")):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
# 與 markets/memory 的 UNITS 必須一致（同一份）
for k, nm, u, gb, tag, src in MEM.UNITS:
    for i, y in enumerate((26, 27, 28)):
        assert abs(TS[f"u_{k}_{y}"][0] - u[i]) < 1e-9, (k, y)

ws_ = openpyxl.load_workbook(SKH_F, data_only=True)["Mkt"]
SHIP = {}   # HBM 出貨位元（EB），三情境各一列（markets/memory v0.4 起隨情境顆數係數變動）
for r in ws_.iter_rows():
    lab = str(r[0].value or "")
    if lab.startswith("出貨＝MIN"):
        sc_ = lab[-2]
        SHIP[sc_] = ([r[2 + j].value for j in range(5)], f"{rel(SKH_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert sorted(SHIP) == ["A", "B", "C"] and all(v is not None for sc_ in SHIP for v in SHIP[sc_][0])

ws_ = openpyxl.load_workbook(EQ_F, data_only=True)["Mkt"]
DRAM = {}; TEST = None; _n = 0
for r in ws_.iter_rows():
    lab = r[0].value
    if lab == "DRAM（含 HBM）〔AI〕":
        s = SC[_n]; _n += 1
        DRAM[s] = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
    if lab == "測試設備":
        TEST = ([r[2 + j].value for j in range(5)], f"{rel(EQ_F)}｜Mkt!C{r[0].row}:G{r[0].row}")
assert _n == 3 and TEST
PEAK, PEAK_SRC = CP.peak_years(ROOT)   # 設備股共用循環位置規則：各情境見頂年度（markets/equipment）

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億日圓（¥B）、十億美元（US$B）；顆數：百萬顆。每個輸入附兩個來源（來源2 找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文（advantest.com）在建置環境無法開啟，數字取自搜尋引擎對公司文件的摘錄，並與第二來源交叉；原文待人工複核"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)

CO = "公司"
sec("A. 加速器顆數（markets/memory 與 HBM、台積電模型同一份；情境係數讀台積電模型成品，不另訂）")
for k, nm, u, gb, tag, src in MEM.UNITS:
    for i, y in enumerate((26, 27, 28)):
        inp(f"u_{k}_{y}", f"{nm} 顆數 20{y}", "百萬顆", None, u[i], None, tag, "markets/memory/market.py UNITS：" + src if i == 0 else "同上", f"台積電模型 {TS[f'u_{k}_{y}'][1]}（值相同，建置時檢查）" if i == 0 else "同上")
    inp(f"g29_{k}", f"{nm} 2029 顆數成長", "%", None, TS[f"g29_{k}"][0], None, "Assumed", TS[f"g29_{k}"][1], "markets/memory 無 2029；沿用台積電模型")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝基準；B＝AI 資本支出消化；C＝明顯下修（與台積電模型相同）")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")
inp("cy30g", "2030 年（日曆年）顆數與記憶體測試機相對 2029 的變化（只用於 FY2029 的 1／4）", "%", -0.10, 0.0, 0.10, "Assumed", "無：市場層只到 2029", "FY2029＝3/4×CY2029＋1/4×CY2030，需要 CY2030；基準持平")

sec("B. SoC 測試機市場（日曆年）與校準：AI 市場＝加權顆數 × 每加權顆測試支出；另加非 AI SoC")
inp("mk_soc25", "SoC 測試機市場 CY2025", "US$B", None, 6.9, None, "Interested-party", "Advantest（測試機廠，有誇大市場傾向）：CY2025 約 6.9（FY2025 簡報；經二手轉述，搜尋摘錄）", "第二來源：未找到")
inp("mk_mem25", "記憶體測試機市場 CY2025", "US$B", None, 2.1, None, "Interested-party", "同上：約 2.1", "第二來源：未找到")
inp("mk_soc26", "SoC 測試機市場 CY2026（公司上修後）", "US$B", 10.5, 11.0, 11.5, "Interested-party", "Advantest FY2025 簡報（2026-04-27）：10.5–11.5（二手轉述，搜尋摘錄；公司原文未開啟）", "第二來源：未找到。SoC＋記憶體合計 13.0–14.5 的二手說法與 10.5+2.5、11.5+3.0 一致")
inp("mk_mem26", "記憶體測試機市場 CY2026", "US$B", 2.5, 2.75, 3.0, "Interested-party", "Advantest：2.5–3.0（公司原文未開啟，二手摘錄）", "合計 13.0–14.5 的二手說法間接支持；第二來源：未找到")
inp("nonai26", "非 AI SoC 測試機市場 CY2026（手機、車用、PC、工業等）", "US$B", 3.7, 4.3, 4.9, "Analogy", "AI 需求出現前的 SoC 測試機市場：Teradyne 4Q22 法說（marketbeat 逐字稿）CY2022 約 4.6、較 2021 −6%（2021 約 4.9，Derived）；Teradyne 2Q23 法說 CY2023 3.7–4.1（https://www.marketbeat.com/earnings/transcripts/86996 、/94281）。低＝CY2023 下緣 3.7，高＝CY2021 推得 4.9，基準＝兩者中點 4.3", "第二來源：Advantest FY2023 4Q Investor Note 列 SoC 測試機市場 4.3／4.0／3.3（年份標籤未能從摘錄確認，僅作數量級佐證）。搜尋過：Teradyne 4Q20／4Q22／2Q23 法說、Advantest 各季 Note 與 Investors Guide、Technavio／ResearchAndMarkets（只有增量或不一致的估計，未採）。與現在的差異：(1) 2021–22 已含早期 AI／運算（Teradyne 稱運算部分歷來僅 US$0.5–0.6B，2020）；(2) 2023 後手機大幅走弱（2020 年手機測試市場約 1.7）、車用 ADAS 與工業走強；(3) Teradyne 與 Advantest 對 TAM 的定義不同（兩家 CY2024 分別約 4.0–4.2 與 4.1，相近）；(4) 基準假設 2026 非 AI 回到 2021–23 區間的中點")
inp("g_nonai", "非 AI SoC 市場年增", "%", 0.0, 0.04, 0.08, "Assumed", "無", "—")
inp("eff_par", "平行測試效率提升（每年降低每加權顆所需測試支出）", "%/年", 0.05, 0.0, 0.10, "Assumed", "無：基準為 0，不納入基準（納入等於用 SEMI 預測校準）；只在 Sens 列 5%／10%", "低欄＝5%、高欄＝10%（Sens 用）；公司未揭露平行度提升")
inp("asp_m", "高階 SoC 測試機平均單價", "US$M／台", 1.0, 1.5, 2.5, "Assumed", "無：公司不揭露台數與單價", "只影響「每百萬顆機台數」的顯示（k＝每顆支出 ÷ 單價），不影響營收與 EPS")
for k, nm, tt, lo, hi, src in (
    ("bw", "NVIDIA Blackwell", 1.0, 0.9, 1.1, "基準世代（定義為 1.0）"),
    ("rb", "NVIDIA Rubin", 1.3, 1.1, 1.6, "晶粒數、HBM4 堆疊數、介面速度上升；Assumed"),
    ("ru", "NVIDIA Rubin Ultra", 1.7, 1.3, 2.2, "4 顆光罩尺寸晶粒＋16 疊 HBM4E（台積電模型 PP 表）；Assumed"),
    ("hp", "NVIDIA Hopper", 0.6, 0.5, 0.8, "單晶粒、HBM3E；上一世代"),
    ("amd", "AMD Instinct", 1.1, 0.9, 1.4, "MI455X 12 疊 HBM4；Assumed"),
    ("tpu", "Google TPU", 1.0, 0.8, 1.3, "TPU 7→8；Assumed"),
    ("trn", "AWS Trainium", 0.8, 0.6, 1.1, "較 Blackwell 簡單；Assumed"),
    ("oth", "其他 ASIC", 0.7, 0.5, 1.0, "中小型 ASIC；Assumed")):
    inp(f"tt_{k}", f"每顆測試時間係數（Blackwell＝1.0）：{nm}", "x", lo, tt, hi, "Assumed", "無實證：公司稱每顆測試時間隨複雜度與世代上升，未給倍數；" + src, "第二來源：未找到。校準在 2026 年進行，所以係數只決定世代間的相對權重")
sec("C. 市占")
inp("sh_soc", "Advantest SoC 測試機市占（CY2025 公司自估；FY 口徑沿用）", "%", 0.61, 0.66, 0.71, "Interested-party", "Advantest FY2025 簡報：CY2024 56%→CY2025 66%（二手轉述，X：jukan05）", "第二來源：未找到。低／高＝±5 個百分點（Assumed 區間）")
inp("shC_28", "情境 C：FY2028 市占（Teradyne 搶回）", "%", 0.56, 0.61, 0.66, "Assumed", "無：假設 2 年內回到 CY2024 的 56%", "Analogy：CY2024 市占 56%（Advantest 簡報）")
inp("shC_29", "情境 C：FY2029 市占", "%", 0.51, 0.56, 0.61, "Analogy", "CY2024 市占 56%", "區間 ±5 個百分點")
inp("sh_mem", "Advantest 記憶體測試機市占（FY2025 實績 ÷ FY2025 市場，Derived）", "%", "=sh_mem-0.05", "=a25_mem/(fmk_mem25*fx)", "=sh_mem+0.05", "Derived", "FY2025 實績（Q 頁）÷（3/4×CY2025＋1/4×CY2026 市場）÷ 匯率", "公司稱 CY2025 記憶體市占略低於前一年（二手轉述）；±5 個百分點為 Assumed")
inp("w_hbm", "記憶體測試機需求中隨 HBM 位元成長的比重（其餘隨 DRAM 設備支出）", "%", 0.3, 0.5, 0.7, "Assumed", "無：公司稱 HBM 測試機為記憶體成長主力", "第二來源：未找到")
sec("D. 記憶體（讀 markets/memory 與 markets/equipment 成品；唯讀）")
for s in SC:
    for j, y in enumerate((25, 26, 27, 28, 29)):
        inp(f"hs{s}_{y}", f"HBM 出貨位元 20{y}（EB；{s}）", "EB", None, SHIP[s][0][j], None, "Derived", SHIP[s][1], "markets/memory Mkt：出貨＝MIN(需求 × 情境顆數係數, 供給)；2028–29 隨情境，2026–27 三情境相同")
for s in SC:
    for j, y in enumerate((25, 26, 27, 28, 29)):
        inp(f"dr{s}_{y}", f"DRAM 設備支出（含 HBM）20{y}（{s}）", "US$B", None, DRAM[s][0][j], None, "Derived", DRAM[s][1], "markets/equipment Mkt（以公司資本支出 × 設備占比推得）")
inp("semi_test26", "SEMI 測試設備 2026", "US$B", None, TEST[0][1], None, "Interested-party", TEST[1], "SEMI 2026-07 年中預測；SEMI 為設備商協會（傾向樂觀）")
inp("semi_test27", "SEMI 測試設備 2027（equipment 內插）", "US$B", None, TEST[0][2], None, "Derived", TEST[1], "2026、2028 為新聞稿值，2027 幾何內插")
inp("semi_test28", "SEMI 測試設備 2028", "US$B", None, TEST[0][3], None, "Interested-party", TEST[1], "同上")
sec("E. 其他系統與服務（前一年 × 成長率）")
inp("g_oth_26", "其他系統 FY2026 成長（Q1 FY26 年化 ÷ FY2025 − 1）", "%", 0.10, "=q1_oth*4/a25_oth-1", 0.35, "Derived", "Q 頁：Q1 FY26 其他系統 26.0×4÷80.5", "低／高 Assumed；Q1 年化僅一季")
inp("g_svc_26", "服務與其他 FY2026 成長（Q1 FY26 年化 ÷ FY2025 − 1）", "%", 0.10, "=q1_svc*4/a25_svc-1", 0.30, "Derived", "Q 頁：Q1 FY26 服務與其他 33.9×4÷109.2", "低／高 Assumed")
for s, o, v in (("A", 0.15, 0.12), ("B", 0.10, 0.10), ("C", 0.00, 0.06)):
    inp(f"g_oth_{s}", f"其他系統 FY2027–29 年成長（{s}）", "%", o - 0.05, o, o + 0.05, "Assumed", "無：FY2025 其他系統無前年可比（公司僅揭露合計）", "—")
    inp(f"g_svc_{s}", f"服務與其他 FY2027–29 年成長（{s}）", "%", v - 0.04, v, v + 0.04, "Assumed", "FY2025 +12.7%（公司：裝機量成長）", "—")
sec("F. 毛利與營業費用（兩點推導：FY2025 實績與 FY2026 指引；不寫死利益率）")
inp("gm_g", "FY2026 指引毛利率（66–67% 中點）", "%", 0.66, 0.665, 0.67, "Interested-party", "Advantest Q1 FY26 法說（2026-07-29）：全年毛利率 66–67%（Investing.com 逐字稿摘要）", "第二來源：未找到")
inp("gm_a", "毛利率截距（＝FY2025 毛利率 − 斜率 × FY2025 SoC 占比）", "%", None, "=a25_gp/a25_rev-gm_b*a25_soc/a25_rev", None, "Derived", "兩點：FY2025（64.3%，SoC 占 68.0%）、FY2026 指引（66.5%，SoC 占 72.6%）", "毛利率隨 SoC 占比上升（公司：產品組合改善）；只有兩點，等同線性外推")
inp("gm_b", "毛利率斜率（每增加 1 個百分點 SoC 占比的毛利率變化）", "x", None, "=(gm_g-a25_gp/a25_rev)/(g_soc/g_rev-a25_soc/a25_rev)", None, "Derived", "同上", "—")
inp("gm_adj_C28", "情境 C：FY2028 毛利率額外調整（價格競爭）", "pt", -0.04, -0.01, 0.0, "Assumed", "無", "Teradyne 搶回份額的價格壓力")
inp("gm_adj_C29", "情境 C：FY2029 毛利率額外調整", "pt", -0.06, -0.02, 0.0, "Assumed", "無", "—")
inp("opex25", "營業費用 FY2025（毛利 − 營業利益）", "¥B", None, "=a25_gp-a25_oi", None, "Derived", "Q 頁", "—")
inp("opex_g", "營業費用 FY2026 指引（指引毛利率 × 指引營收 − 指引營業利益）", "¥B", None, "=gm_g*g_rev-g_oi", None, "Derived", "指引", "公司：R&D 全年預測 110.0（較 4 月 +10.0）")
inp("opex_var", "營業費用變動部分占營收（兩點斜率）", "%", None, "=(opex_g-opex25)/(g_rev-a25_rev)", None, "Derived", "（FY2026 指引營業費用 − FY2025）÷（營收差）", "—")
inp("opex_fix", "營業費用固定部分", "¥B", None, "=opex25-opex_var*a25_rev", None, "Derived", "FY2025 營業費用 − 變動部分 × FY2025 營收", "—")
inp("phi", "營業費用中變動部分占比（FY2025；Sens 用）", "%", 0.35, "=opex_var*a25_rev/opex25", 0.75, "Derived", "基準＝兩點推導值；低／高 Assumed", "—")
sec("G. 淨利、股數、匯率")
inp("nioi", "淨利 ÷ 營業利益", "x", 0.72, "=a25_ni/a25_oi", "=g_ni/g_oi", "Derived", "基準＝FY2025 實績 0.752；高＝FY2026 指引隱含 0.780（含戰略投資評價益，屬一次性，不當基準，只在 Sens 列示；處理方式同台積電模型的一次性利益）；低 Assumed（稅率正常化）", "FY2025 Q4 含選擇權評價益約 ¥17.3B")
inp("shares", "股數（指引淨利 ÷ 指引 EPS）", "百萬股", None, "=g_ni/g_eps*1000", None, "Derived", "660.0B÷911.64＝約 724.0M", "FY2025 加權平均約 728.6M（淨利 375.4B÷515.15）")
inp("fx", "匯率", "¥/US$", 140, 150, 160, "Interested-party", "公司 FY2026 假設：第二季起 ¥150（Q1 FY26 法說；公司原文未開啟）", "第二來源：未找到。FY2025 實績不換匯；只用於把美元市場換成日圓營收")
sec("H. 公司指引（FY2026；2026-07-29 上修）與對帳來源")
inp("g_rev", "指引營收 FY2026", "¥B", None, 1714.0, None, "Interested-party", "Advantest 2027年3月期 第1四半期決算短信（2026-07-29，J_FR_FY2026_1Q.pdf）", "Investing.com Q1 FY26；kabutan（4 月版 1,420.0）")
inp("g_oi", "指引營業利益 FY2026", "¥B", None, 846.0, None, "Interested-party", "同上", "kabutan／gyokaidigest（4 月版 627.5）")
inp("g_ni", "指引淨利 FY2026（歸屬母公司）", "¥B", None, 660.0, None, "Interested-party", "同上", "kabutan／gyokaidigest（4 月版 465.5）")
inp("g_soc", "指引 SoC 測試機 FY2026", "¥B", None, 1243.5, None, "Interested-party", "公司 Q1 FY26 簡報；Investing.com：由 999.5 上修", "—")
inp("g_mem", "指引記憶體測試機 FY2026", "¥B", None, 227.0, None, "Interested-party", "公司 Q1 FY26 指引（原文未開啟）；4 月版 201.0（+17%，日經摘錄）", "第二來源：未找到")
inp("g_eps", "指引 EPS FY2026（基本）", "¥", None, 911.64, None, "Interested-party", "公司決算短信（Yahoo Japan 摘要）", "kabutan 記 913.83（差 0.2%，未解）")
inp("cons_rev26", "共識營收 FY2026", "¥B", None, 1760.0, None, "Interested-party", "stockanalysis.com 2026-10-07：¥1.76T", "第二來源：未找到")
inp("cons_oi26", "共識營業利益 FY2026", "¥B", None, 886.4, None, "Interested-party", "stockanalysis.com 2026-10-07", "第二來源：未找到")
for s in SC:
    inp(f"pk_{s}", f"情境 {s}：設備市場見頂年度（日曆年；循環位置規則用）", "年", None, PEAK[s], None, "Derived", PEAK_SRC[s], "設備股共用規則（tools/cycle_position.py）：FY(t) 早於見頂年＝中段、等於＝見頂、晚於＝谷底；設備市場模型 v0.1.2 的情境係數為 Assumed" if s == "A" else "同上")
inp("px", "股價（2026-10-23 收盤）", "¥", None, 32340, None, "Verified", "Yahoo!ファイナンス 時系列（2026-10-23 終值 32,340）", "株探 日足（同值）；markets/invest_map/data/prices.csv 兩來源核對")
inp("sa_fpe", "遠期 P/E（stockanalysis，2026-10-21 價格 ¥34,440）", "x", None, 37.7, None, "Interested-party", "stockanalysis.com", "自算：股價 ÷ 指引 EPS，見 Val")
# v0.2.3：目標價情境權重（放在 In 頁最後，不位移既有列）
sec("I. 目標價情境權重（Val 頁「0. 目標價與上檔空間」用；Andy 可改）")
for s_, w_ in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s_}", f"目標價情境權重 {s_}", "%", None, w_, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")

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
IN_LAST = r
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績（FY2025 四季、FY2026 第一季） ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績（十億日圓；EPS 為日圓）。會計年度 4–3 月：FY2025＝2025-04～2026-03。藍字＝輸入；黑字＝公式（Derived）。公司原文（決算短信、Q 別簡報）在建置環境無法開啟，數字取自搜尋引擎對公司文件的摘錄，並與第二來源交叉"
Q["A2"] = "FY2026 第二季（2026-07～09）：預定 2026-10-28 15:30 JST 公布（公司 IR 行事曆）；建置日 2026-10-22 尚未公布，故不納入，指引維持 2026-07-29 上修版。公布後更新本頁 H 欄與 In 頁指引"
hdr(Q, 3, ["項目", "FY25 Q1（4–6月）", "FY25 Q2（7–9月）", "FY25 Q3（10–12月）", "FY25 Q4（1–3月）", "FY2025 全年", "FY26 Q1（4–6月）", "FY26 Q2（未公布）", "標記", "來源1（公司）", "來源2／說明"])
QR = {}
def qrow(r, key, label, vals, tag, s1, s2, fmt="#,##0.0"):
    QR[key] = r
    Q.cell(r, 1, label)
    for j, v in zip("BCDEFG", vals):
        if v is None: continue
        c = Q[f"{j}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    Q.cell(r, 9, tag); Q.cell(r, 10, s1); Q.cell(r, 11, s2)
N1 = "Advantest 各季 Investor Note（advantest.com/document/en/investors/ir-library/result/）"
qrow(4, "soc", "SoC 測試機", [191.3, 173.7, 165.2, 237.2, 767.4, 259.6], "Verified", N1 + "：Q1 191.3（Q2 −17.6 → 173.7；Q4 +72.1 → 237.2）；Q3 以 Q4 反推 165.1–165.2", "Investing.com 各季簡報報導（Q3 165.2、Q4 237.2、FY 767.4）；四季合計＝767.4（見檢查列）")
qrow(5, "mem", "記憶體測試機", [33.5, 43.9, 57.3, 36.8, 171.5, 48.0], "Verified", N1 + "：Q1 33.5、Q2 43.9、Q4 36.8（較 Q3 −20.6）", "Investing.com：Q3 57.3、FY 171.5、FY26 Q1 48.0；四季合計＝171.5")
qrow(6, "oth", "其他系統", [15.8, None, 22.6, None, 80.5, 26.0], "Q1 Derived／其餘 Verified", "Q1：測試系統事業 240.6 − SoC − 記憶體（公司 FY2025 Q1 短信：測試系統 240.6）；FY26 Q1：測試系統 333.6 − SoC − 記憶體", "Investing.com：Q3 22.6、FY 80.5；Q2、Q4 單季未找到（見下一列合計）")
qrow(7, "svc", "服務、支援與其他", [23.2, None, 28.7, None, 109.2, 33.9], "Derived／Verified", "FY2025：公司事業別表「Services and Others」109.2（+12.7%）；單季＝營收 − 測試系統", "MarketScreener（公司 FY2025 決算）：109.2 vs 96.9；Q2、Q4 單季未找到")
qrow(8, "othsvc", "其他系統＋服務與其他（Q2、Q4 單季未找到分項，列合計）", ["=B6+B7", "=C9-C4-C5", "=D6+D7", "=E9-E4-E5", "=F6+F7", "=G6+G7"], "Derived", "營收 − SoC − 記憶體", "—")
qrow(9, "rev", "營收", [263.8, 262.9, 273.8, 328.1, 1128.6, 367.5], "Verified", "公司決算：FY2025 1,128.6（+44.7%）；上半 526.7、前三季 800.5；FY26 Q1 367.5（+39.3%）", "Investing.com／kabutan；Q1 FY26 3,675 億日圓")
qrow(10, "chk1", "檢查：SoC＋記憶體＋其他＋服務 − 營收（應為 0）", [f"={c}4+{c}5+{c}8-{c}9" for c in "BCDEFG"], "—", "—", "—", "0.0")
qrow(11, "gmr", "毛利率", [0.651, None, 0.620, 0.674, 0.643, 0.695], "Verified／Interested-party", "公司簡報：FY26 Q1 69.5%（前年同期 65.1%）；Q4 67.4%；Q2 僅稱「高於 60%」", "Investing.com：Q3 62.0%、FY2025 64.3%（FY2024 57.1%）", "0.0%")
qrow(12, "gp", "毛利", ["=B9*B11", "=F12-B12-D12-E12", "=D9*D11", "=E9*E11", "=F9*F11", "=G9*G11"], "Derived", "營收 × 毛利率；Q2＝全年 − 其他三季（殘差，與公司「高於 60%」一致）", "Investing.com：Q3 毛利 169.7")
qrow(13, "oi", "營業利益", ["=232.4-C13", 108.4, 113.6, 153.1, 499.1, 190.0], "Q1 Derived／其餘 Verified", "公司：上半營業利益 232.4、前三季 346.0、FY 499.1（+118.8%）、Q4 營益率 46.7%；FY26 Q1 190.0（+53.3%）", "Investing.com／kabutan；Q2 108.4 為單一二手（Desk Research Design）")
qrow(14, "oim", "營業利益率", [f"={c}13/{c}9" for c in "BCDEFG"], "Derived", "—", "公司：FY2025 44.2%、Q4 46.7%、FY26 Q1 51.7%", "0.0%")
qrow(15, "opx", "營業費用（毛利 − 營業利益）", [f"={c}12-{c}13" for c in "BCDEFG"], "Derived", "—", "—")
qrow(16, "ni", "淨利（歸屬母公司）", ["=169.8-C16", 79.6, 78.7, 126.9, 375.4, 174.8], "Q1 Derived／其餘 Verified", "公司：上半 169.8、FY 375.4（+132.9%）；FY26 Q1 174.8（+93.8%，含戰略投資評價益）", "Investing.com；Q2 79.6 單一二手")
qrow(17, "eps", "EPS（基本，日圓）", [123.14, "=232.19-B17", "=F17-232.19-E17", 174.06, 515.15, 241.27], "Verified／Derived", "公司決算短信：Q1 123.14、上半 232.19、FY 515.15、Q4 174.06、FY26 Q1 241.27", "松井證券決算一覽（FY 515.15、Q4 174.06）；Q2、Q3 為差額（Derived）")
qrow(18, "ck2", "檢查：四季合計 − 全年（SoC、記憶體、營收、營業利益、淨利、EPS 依序在 B–G 欄）", [f"=SUM(B4:E4)-F4", f"=SUM(B5:E5)-F5", f"=SUM(B9:E9)-F9", f"=SUM(B13:E13)-F13", f"=SUM(B16:E16)-F16", f"=SUM(B17:E17)-F17"], "—", "—", "—", "0.00")
Q["H4"] = "預定 2026-10-28 15:30 JST，尚未公布"
for nm, rr in (("soc", 4), ("mem", 5), ("oth", 6), ("svc", 7), ("rev", 9), ("gp", 12), ("oi", 13), ("ni", 16), ("eps", 17)):
    defname(f"a25_{nm}", f"Q!$F${rr}")
for nm, rr in (("soc", 4), ("mem", 5), ("oth", 6), ("svc", 7), ("rev", 9), ("oi", 13), ("ni", 16), ("eps", 17)):
    defname(f"q1_{nm}", f"Q!$G${rr}")
Q.cell(21, 1, "FY2026 公司指引（2026-07-29 上修；4 月版供對照）").font = BOLD
hdr(Q, 22, ["項目", "4 月版", "7 月版（上修）", "上修幅度", "FY2025 實績", "7 月版年增", "", "", "標記", "來源1（公司）", "來源2／說明"])
G = [("營收", 1420.0, "=g_rev", "=a25_rev"), ("SoC 測試機", 999.5, "=g_soc", "=a25_soc"), ("記憶體測試機", 201.0, "=g_mem", "=a25_mem"),
     ("其他系統＋服務與其他（營收 − SoC − 記憶體，Derived）", "=B23-B24-B25", "=C23-C24-C25", "=a25_oth+a25_svc"), ("營業利益", 627.5, "=g_oi", "=a25_oi"),
     ("淨利", 465.5, "=g_ni", "=a25_ni"), ("EPS（基本，日圓）", None, "=g_eps", "=a25_eps")]
for i, (nm, a, b, c_) in enumerate(G, 23):
    Q.cell(i, 1, nm)
    if a is not None:
        Q.cell(i, 2, a).number_format = "#,##0.0"
        if not isinstance(a, str): Q.cell(i, 2).font = BLUE
        Q.cell(i, 4, f"=C{i}/B{i}-1").number_format = "0.0%"
    Q.cell(i, 3, b).number_format = "#,##0.0"; Q.cell(i, 5, c_).number_format = "#,##0.0"; Q.cell(i, 6, f"=C{i}/E{i}-1").number_format = "0.0%"
    Q.cell(i, 9, "Interested-party"); Q.cell(i, 10, "4 月版：公司 FY2025 決算（2026-04-27）；7 月版：Q1 FY26 決算短信（2026-07-29）"); Q.cell(i, 11, "kabutan、gyokaidigest、Investing.com；7 月 EPS 公司 911.64、kabutan 913.83")
Q.cell(30, 1, "營業利益率（指引）"); Q.cell(30, 2, "=B27/B23").number_format = "0.0%"; Q.cell(30, 3, "=C27/C23").number_format = "0.0%"; Q.cell(30, 5, "=a25_oi/a25_rev").number_format = "0.0%"
Q.cell(30, 11, "公司：49.4%（FY2025 44.2%）")
Q.cell(32, 1, "FY2025 四項合計與營收差約 ¥44B 的查明").font = BOLD
Q.cell(33, 1, "SoC＋記憶體＋其他系統＋服務（65.0，二手說法）"); Q.cell(33, 2, "=a25_soc+a25_mem+a25_oth+65").number_format = "#,##0.0"
Q.cell(34, 1, "營收 − 上列"); Q.cell(34, 2, "=a25_rev-B33").number_format = "#,##0.0"
Q.cell(35, 1, "「服務、支援與其他」全年（公司事業別表）− 65.0"); Q.cell(35, 2, "=a25_svc-65").number_format = "#,##0.0"
Q.cell(36, 1, "SoC＋記憶體＋其他系統＋「服務、支援與其他」109.2 − 營收（應為 0）"); Q.cell(36, 2, "=a25_soc+a25_mem+a25_oth+a25_svc-a25_rev").number_format = "0.0"
Q.cell(37, 1, "結論：差額 44.2 ＝ 109.2 − 65.0。65.0 是「服務／支援」的子項（二手說法），公司事業別表的第四項是「服務、支援與其他」109.2（含耗材等其他收入）；四項改用 109.2 後與營收完全相符（差 0）。65.0 未在公司原文核到。")
Q.column_dimensions["A"].width = 58
for c in "BCDEFGH": Q.column_dimensions[c].width = 14
Q.column_dimensions["I"].width = 22; Q.column_dimensions["J"].width = 90; Q.column_dimensions["K"].width = 70
Q.freeze_panes = "B4"

# ---------------- Units：加速器顆數（三情境）、測試時間加權、AI 每顆支出的校準 ----------------
PLAT = [k for k, *_ in MEM.UNITS]; PNAME = {k: nm for k, nm, *_ in MEM.UNITS}
CYS = (2026, 2027, 2028, 2029, 2030)
cyc = lambda y: chr(ord("C") + y - 2026)          # Units／引擎的日曆年欄：C＝2026 … G＝2030
def u(k, y, s):
    if y <= 2027: return f"u_{k}_{y % 100}"
    if y == 2028: return f"u_{k}_28*um28_{s}"
    return f"u_{k}_28*(1+g29_{k})*um29_{s}"

UN = wb.create_sheet("Units")
UN["A1"] = "Units — 加速器顆數（百萬顆；日曆年；與 HBM、台積電模型同一份）× 每顆測試時間係數 → 加權顆數；2026 年用公司 CY2026 市場規模校準每加權顆的測試支出"
hdr(UN, 3, ["項目", "單位"] + [f"CY{y}" for y in CYS] + ["說明"])
UR = {}; ur = 4
for s in SC:
    UN.cell(ur, 1, f"情境 {s}：加速器顆數").font = BOLD; ur += 1
    first = ur
    for k in PLAT:
        UN.cell(ur, 1, f"  {PNAME[k]}"); UN.cell(ur, 2, "百萬顆")
        for y in CYS[:4]: UN[f"{cyc(y)}{ur}"] = "=" + u(k, y, s); UN[f"{cyc(y)}{ur}"].number_format = "0.00"
        UN[f"G{ur}"] = f"=F{ur}*(1+cy30g)"; UN[f"G{ur}"].number_format = "0.00"
        UR[(s, k)] = ur; ur += 1
    UN.cell(ur, 1, "  合計").font = BOLD; UN.cell(ur, 2, "百萬顆")
    for y in CYS: UN[f"{cyc(y)}{ur}"] = f"=SUM({cyc(y)}{first}:{cyc(y)}{ur-1})"; UN[f"{cyc(y)}{ur}"].number_format = "0.00"
    UN.cell(ur, 8, "2026＝15.2、2027＝23.7、2028＝26.5（A）；與 markets/memory 的 UNITS 相同" if s == "A" else "2028–29 乘以情境係數 um28／um29（台積電模型）")
    UR[(s, "tot")] = ur; ur += 2
UN.cell(ur, 1, "測試時間加權顆數（Σ 顆數 × tt）").font = BOLD; ur += 1
for s in SC:
    UN.cell(ur, 1, f"  情境 {s}"); UN.cell(ur, 2, "百萬顆（Blackwell 等值）")
    for y in CYS:
        UN[f"{cyc(y)}{ur}"] = "=" + "+".join(f"{cyc(y)}{UR[(s, k)]}*tt_{k}" for k in PLAT); UN[f"{cyc(y)}{ur}"].number_format = "0.00"
    UR[(s, "wu")] = ur; ur += 1
ur += 1
UN.cell(ur, 1, "校準（只用公司 CY2026 市場規模，不用共識）").font = BOLD; ur += 1
cal = [("CY2026 SoC 測試機市場（公司）", "US$B", "=mk_soc26", "0.00", "Interested-party"),
       ("− 非 AI SoC（Analogy：AI 前的 SoC 測試機市場）", "US$B", "=nonai26", "0.00", "Analogy，3.7–4.9；見 In 頁"),
       ("＝ AI SoC 測試機市場 CY2026", "US$B", "=C{a}-C{b}", "0.00", "Derived"),
       ("÷ CY2026 加權顆數（情境 A＝B＝C，2026 年無情境係數）", "百萬顆", f"=C{UR[('A','wu')]}", "0.00", ""),
       ("＝ 每加權顆測試支出", "US$／顆", "=C{c}/C{d}*1000", "0", "Derived：每顆 Blackwell 等值的測試機支出"),
       ("÷ 測試機單價（Assumed）＝ 每百萬顆所需機台數 k", "台／百萬顆", "=C{e}/asp_m", "0", "Derived（單價為 Assumed 1.0–2.5 US$M；k 與單價只有乘積可識別）")]
cr0 = ur
refs = {}
for i, (lab, unit, f, fmt, note) in enumerate(cal):
    refs["abcdef"[i]] = cr0 + i
for i, (lab, unit, f, fmt, note) in enumerate(cal):
    UN.cell(cr0 + i, 1, lab); UN.cell(cr0 + i, 2, unit)
    UN[f"C{cr0+i}"] = f.format(**refs) if "{" in f else f; UN[f"C{cr0+i}"].number_format = fmt; UN.cell(cr0 + i, 8, note)
defname("cw_unit", f"Units!$C${refs['e']}"); defname("k_mach", f"Units!$C${refs['f']}")
ur = cr0 + len(cal) + 1
UN.cell(ur, 1, "FY2025 市場（3/4×CY2025＋1/4×CY2026；用於 FY2025 隱含市占與記憶體市占）").font = BOLD; ur += 1
UN.cell(ur, 1, "SoC 測試機"); UN.cell(ur, 2, "US$B"); UN[f"B{ur}"] = "=0.75*mk_soc25+0.25*mk_soc26"; UN[f"B{ur}"].number_format = "0.00"; defname("fmk_soc25", f"Units!$B${ur}"); ur += 1
UN.cell(ur, 1, "記憶體測試機"); UN.cell(ur, 2, "US$B"); UN[f"B{ur}"] = "=0.75*mk_mem25+0.25*mk_mem26"; UN[f"B{ur}"].number_format = "0.00"; defname("fmk_mem25", f"Units!$B${ur}"); ur += 2
UN.cell(ur, 1, "測試時間係數（Assumed；Blackwell＝1.0）").font = BOLD; ur += 1
hdr(UN, ur, ["平台", "", "低", "基準", "高"]); ur += 1
for k in PLAT:
    UN.cell(ur, 1, PNAME[k]); UN[f"C{ur}"] = f"=tt_{k}_lo"; UN[f"D{ur}"] = f"=tt_{k}"; UN[f"E{ur}"] = f"=tt_{k}_hi"; ur += 1
UN.column_dimensions["A"].width = 56; UN.column_dimensions["B"].width = 22; UN.column_dimensions["H"].width = 70
UN.cell(ur + 1, 1, "核對（PL 引擎的加權顆數 − 本頁，應為 0）在 PL 頁底部。")
UNITS_CHECK_ROW = ur + 2

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用，確保一致） ----------------
FYC = {y: chr(ord("C") + y - 2025) for y in range(2025, 2030)}   # FY2025＝C … FY2029＝G
DEFP = dict(fx="fx", eff="0", sh="sh_soc", shC28="shC_28", shC29="shC_29", shm="sh_mem", tt="", mk26="mk_soc26", nonai="nonai26",
            w="w_hbm", F="opex_fix", v="opex_var", nioi="nioi", mkm26="mk_mem26", scn="A")
def engine(ws, r0, title, P0=None, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "CY2026／FY2025", "CY2027／FY2026", "CY2028／FY2027", "CY2029／FY2028", "CY2030／FY2029", "說明"]); r += 1
    ws.cell(r, 1, "A. 日曆年（C＝CY2026 … G＝CY2030）").font = BOLD; r += 1
    def row(key, label, unit, fs, fmt="0.00", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 8, note); R[key] = r; r += 1
    row("wu", "測試時間加權顆數", "百萬顆", {**{cyc(y): "=" + "+".join(f"{u(k, y, s)}*tt_{k}{P['tt']}" for k in PLAT) for y in CYS[:4]}, "G": f"=F{r}*(1+cy30g)"}, note="Σ 顆數 × tt；CY2030＝CY2029×(1+cy30g)")
    row("cw", "每加權顆測試支出（以 CY2026 校準）", "US$B／百萬顆", {"C": f"=({P['mk26']}-{P['nonai']})/C{R['wu']}"}, fmt="0.0000", note="（公司 CY2026 SoC 市場 − 非 AI）÷ 加權顆數")
    row("ai", "AI SoC 測試機市場", "US$B", {cyc(y): f"=$C${R['cw']}*{cyc(y)}{R['wu']}*(1-{P['eff']})^{y-2026}" for y in CYS})
    row("nai", "非 AI SoC 測試機市場", "US$B", {cyc(y): f"={P['nonai']}*(1+g_nonai)^{y-2026}" for y in CYS})
    row("soc", "SoC 測試機市場 CY", "US$B", {cyc(y): f"={cyc(y)}{R['ai']}+{cyc(y)}{R['nai']}" for y in CYS}, note="CY2026 ＝ 公司規模（校準）")
    dr = lambda y: f"dr{s}_{y % 100}"
    row("mem", "記憶體測試機市場 CY", "US$B", {**{cyc(y): f"={P['mkm26']}*({P['w']}*hs{s}_{y%100}/hs{s}_26+(1-{P['w']})*{dr(y)}/{dr(2026)})" for y in CYS[:4]}, "G": f"=F{r}*(1+cy30g)"}, note="CY2026 錨定公司規模；之後＝w×HBM 出貨位元指數＋(1−w)×DRAM 設備支出指數")
    ws.cell(r, 1, "B. 會計年度（C＝FY2025 … G＝FY2029；FY(t)＝3/4×CY(t)＋1/4×CY(t+1)）").font = BOLD; r += 1
    fy = lambda y: FYC[y]
    row("fsoc", "SoC 測試機市場 FY", "US$B", {"C": f"=fmk_soc25", **{fy(y): f"=0.75*{cyc(y)}{R['soc']}+0.25*{cyc(y+1)}{R['soc']}" for y in range(2026, 2030)}})
    row("fmem", "記憶體測試機市場 FY", "US$B", {"C": f"=fmk_mem25", **{fy(y): f"=0.75*{cyc(y)}{R['mem']}+0.25*{cyc(y+1)}{R['mem']}" for y in range(2026, 2030)}})
    shf = {2026: P["sh"], 2027: P["sh"], 2028: P["shC28"] if s == "C" else P["sh"], 2029: P["shC29"] if s == "C" else P["sh"]}
    row("sh", "Advantest SoC 市占", "%", {"C": f"=a25_soc/(C{R['fsoc']}*{P['fx']})", **{fy(y): f"={shf[y]}" for y in range(2026, 2030)}}, fmt="0.0%", note="FY2025＝實績隱含（對照公司自估 66%）")
    row("rsoc", "SoC 測試機營收", "¥B", {"C": "=a25_soc", **{fy(y): f"={fy(y)}{R['fsoc']}*{fy(y)}{R['sh']}*{P['fx']}" for y in range(2026, 2030)}}, fmt="#,##0.0", note="市場 × 市占 × 匯率")
    row("rmem", "記憶體測試機營收", "¥B", {"C": "=a25_mem", **{fy(y): f"={fy(y)}{R['fmem']}*{P['shm']}*{P['fx']}" for y in range(2026, 2030)}}, fmt="#,##0.0")
    g_o = {2026: "g_oth_26", **{y: f"g_oth_{s}" for y in (2027, 2028, 2029)}}; g_s = {2026: "g_svc_26", **{y: f"g_svc_{s}" for y in (2027, 2028, 2029)}}
    row("roth", "其他系統營收", "¥B", {"C": "=a25_oth", **{fy(y): f"={fy(y-1)}{r}*(1+{g_o[y]})" for y in range(2026, 2030)}}, fmt="#,##0.0")
    row("rsvc", "服務、支援與其他營收", "¥B", {"C": "=a25_svc", **{fy(y): f"={fy(y-1)}{r}*(1+{g_s[y]})" for y in range(2026, 2030)}}, fmt="#,##0.0")
    row("rev", "營收", "¥B", {fy(y): f"={fy(y)}{R['rsoc']}+{fy(y)}{R['rmem']}+{fy(y)}{R['roth']}+{fy(y)}{R['rsvc']}" for y in range(2025, 2030)}, fmt="#,##0.0")
    row("ssh", "SoC 占營收", "%", {fy(y): f"={fy(y)}{R['rsoc']}/{fy(y)}{R['rev']}" for y in range(2025, 2030)}, fmt="0.0%")
    adj = {2026: "0", 2027: "0", 2028: "gm_adj_C28" if s == "C" else "0", 2029: "gm_adj_C29" if s == "C" else "0"}
    row("gm", "毛利率", "%", {"C": "=a25_gp/a25_rev", **{fy(y): f"=gm_a+gm_b*{fy(y)}{R['ssh']}+{adj[y]}" for y in range(2026, 2030)}}, fmt="0.0%", note="截距＋斜率×SoC 占比（兩點推導）")
    row("gp", "毛利", "¥B", {fy(y): f"={fy(y)}{R['rev']}*{fy(y)}{R['gm']}" for y in range(2025, 2030)}, fmt="#,##0.0")
    row("opx", "營業費用", "¥B", {"C": "=opex25", **{fy(y): f"={P['F']}+{P['v']}*{fy(y)}{R['rev']}" for y in range(2026, 2030)}}, fmt="#,##0.0", note="固定＋變動×營收")
    row("oi", "營業利益", "¥B", {fy(y): f"={fy(y)}{R['gp']}-{fy(y)}{R['opx']}" for y in range(2025, 2030)}, fmt="#,##0.0")
    row("oim", "營業利益率", "%", {fy(y): f"={fy(y)}{R['oi']}/{fy(y)}{R['rev']}" for y in range(2025, 2030)}, fmt="0.0%")
    row("ni", "淨利", "¥B", {"C": "=a25_ni", **{fy(y): f"={fy(y)}{R['oi']}*{P['nioi']}" for y in range(2026, 2030)}}, fmt="#,##0.0")
    row("eps", "EPS", "¥", {"C": "=a25_eps", **{fy(y): f"={fy(y)}{R['ni']}*1000/shares" for y in range(2026, 2030)}}, fmt="#,##0.0")
    row("usd", "營收（美元）", "US$B", {fy(y): f"={fy(y)}{R['rev']}/{P['fx']}" for y in range(2025, 2030)}, fmt="0.0", note="FY2025 以基準匯率換算（顯示用）")
    R["end"] = r
    return R

# ---------------- PL：三情境 ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — Advantest 損益（¥B；FY2025 為實績，FY2026–FY2029 為模型）。三情境各一個完整推導區塊（見下方），本頁頂端為摘要"
PB = {}; r0 = 40
for s in SC:
    PB[s] = engine(PL, r0, {"A": "情境 A：基準（顆數續增）", "B": "情境 B：2028 顆數 ×0.8、2029 ×0.85", "C": "情境 C：2028–29 顆數 ×0.6；Teradyne 搶回市占、價格壓力"}[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "¥B", "#,##0"), ("usd", "營收", "US$B", "0.0"), ("rsoc", "  SoC 測試機", "¥B", "#,##0"), ("rmem", "  記憶體測試機", "¥B", "#,##0"),
                                ("gm", "毛利率", "%", "0.0%"), ("oi", "營業利益", "¥B", "#,##0"), ("oim", "營業利益率", "%", "0.0%"), ("ni", "淨利", "¥B", "#,##0"), ("eps", "EPS", "¥", "#,##0.0")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in range(2025, 2030):
            PL[f"{FYC[y]}{sr}"] = f"={FYC[y]}{PB[s][key]}"; PL[f"{FYC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對：PL 引擎加權顆數 − Units（應為 0）").font = BOLD; sr += 1
for s in SC:
    PL.cell(sr, 1, f"  情境 {s}")
    PL[f"C{sr}"] = "=" + "+".join(f"ABS({cyc(y)}{PB[s]['wu']}-Units!{cyc(y)}{UR[(s,'wu')]})" for y in CYS); PL[f"C{sr}"].number_format = "0.0000"; sr += 1
PL.column_dimensions["A"].width = 50; PL.column_dimensions["B"].width = 18; PL.column_dimensions["H"].width = 60
for c in "CDEFG": PL.column_dimensions[c].width = 13

# ---------------- Val：P/E 估值 ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值。跨公司規則：合理倍數＝10 年（2016–2025 年底）P/E 的分位，依循環位置：見頂年 25 分位、中段中位數、谷底 75 分位。循環位置依設備股共用規則（見第 4 節；v0.2 以前三情境皆把 FY2027 歸為見頂年）"
V["A2"] = "P/E 定義（自算）：日曆年底收盤價 ÷ 「當期會計年度」公司公布的基本 EPS（年底所在會計年度，例如 2025-12 對 FY2025＝2026-03 期）；另列「前期」（上一個已結束會計年度）。所有價格與 EPS 為 2023-10 的 4 比 1 分割調整後"
V.cell(4, 1, "1. 公司公布 EPS（會計年度，日圓，分割調整後）— 兩個來源").font = BOLD
hdr(V, 5, ["會計年度（3 月結束）", "來源1：irbank（分割調整後；2025-03 改用公司決算書）", "來源2：公司原始 EPS（分割前）", "來源2 ÷ 4", "兩者差距", "標記", "說明"])
EPSD = [("2016-03", 9.59, 38.35, "公司 74th 年報：基本 38.35"), ("2017-03", 20.27, 81.07, "公司 FY2017 決算附註：基本 81.07（摘要表另列 148.11，口徑不同，未採）"),
        ("2018-03", 25.49, 101.94, "公司 FY2018 決算附註：101.94（摘要表 256.28，未採）"), ("2019-03", 75.59, 302.35, "公司 FY2019 決算附註：302.35"),
        ("2020-03", 67.53, 270.12, "公司 FY2020 決算：270.12"), ("2021-03", 88.47, 353.87, "公司 FY2020 決算 353.87；株探 353.9"),
        ("2022-03", 112.39, 449.6, "株探 449.6（公司值未直接取得）"), ("2023-03", 174.35, 697.4, "株探 697.4（公司值未直接取得）"),
        ("2024-03", 84.45, 337.8, "公司 FY2024 決算：84.45（分割後）；右欄為 84.45×4 的還原（Derived，非獨立來源）"),
        ("2025-03", 226.24, 904.96, "公司 FY2025 決算書（FY2024 Consolidated Financial Results，E_FR_FY2024_FN.pdf）：基本 226.24、希薄化後 225.77；右欄為 ×4 還原。irbank／松井證券列 218.67，口徑未解，不採"),
        ("2026-03", 515.15, 2060.6, "公司 FY2025 決算 515.15（Q 頁）；右欄為 ×4 還原")]
EP = {}
for i, (fy_, a, raw, note) in enumerate(EPSD, 6):
    V.cell(i, 1, fy_)
    V.cell(i, 2, a).font = BLUE; V.cell(i, 3, raw).font = BLUE
    V.cell(i, 4, f"=C{i}/4"); V.cell(i, 5, f"=B{i}/D{i}-1")
    for c in "BD": V[f"{c}{i}"].number_format = "0.00"
    V[f"E{i}"].number_format = "0.0%"
    V.cell(i, 6, "Verified" if i < 11 else "Verified（單一來源＋還原）"); V.cell(i, 7, note); EP[fy_] = i
V.cell(17, 1, "全部年度統一用基本 EPS（Basic）。公司決算書均區分基本與稀釋；2016–2020 年 3 月期取自公司決算附註的基本 EPS，2021-03 與 2024-03 以後取自公司決算的基本 EPS，2022-03、2023-03 為株探值（未直接取得公司值）").font = BOLD
V.cell(18, 1, "2. 10 年 P/E（年底）").font = BOLD
hdr(V, 19, ["年底", "收盤價 ¥（分割調整）", "當期 FY EPS", "前期 FY EPS", "P/E 當期（自算）", "P/E 前期（自算）", "companiesmarketcap P/E", "自算當期 − cmc", "自算前期 − cmc", "標記", "來源／說明"])
PX = {2016: None, 2017: 522.8, 2018: 561.0, 2019: 1540.0, 2020: 1932.5, 2021: 2725.0, 2022: 2120.0, 2023: 4797.0, 2024: 9198.0, 2025: 19635.0}
CMC = {2016: 23.9, 2017: 22.8, 2018: 19.6, 2019: 8.10, 2020: 15.1, 2021: 26.2, 2022: 21.1, 2023: 16.8, 2024: 76.3, 2025: None}
for i, y in enumerate(range(2016, 2026), 20):
    V.cell(i, 1, f"{y}-12")
    if PX[y] is not None: V.cell(i, 2, PX[y]).font = BLUE
    V.cell(i, 3, f"=B{EP[f'{y+1}-03']}"); V.cell(i, 4, f"=B{EP[f'{y}-03']}")
    V.cell(i, 5, f'=IF(B{i}="","",B{i}/C{i})'); V.cell(i, 6, f'=IF(B{i}="","",B{i}/D{i})')
    if CMC[y] is not None: V.cell(i, 7, CMC[y]).font = BLUE
    V.cell(i, 8, f'=IF(OR(B{i}="",G{i}=""),"",E{i}-G{i})'); V.cell(i, 9, f'=IF(OR(B{i}="",G{i}=""),"",F{i}-G{i})')
    for c in "BCD": V[f"{c}{i}"].number_format = "#,##0.00"
    for c in "EFGHI": V[f"{c}{i}"].number_format = "0.0"
    V.cell(i, 10, "Verified／Derived")
    V.cell(i, 11, {2016: "收盤價未找到（檢索缺口，不是不存在）：不納入分位；補入的敏感度見下", 2023: "當期 EPS 84.45 偏低（FY2023 淨利 62.3B，Q 頁外）→ P/E 56.8 為極端值", 2025: "companiesmarketcap 2025 年底值未取得（搜尋摘錄缺列）；Yahoo Finance 2025-12-31 TTM P/E 55.18 可作參考"}.get(y, "收盤價：日經新聞 過去 10 年股價推移（年底收盤）"))
V.cell(31, 1, "收盤價來源1：日經新聞 過去10年株價（nikkei.com/nkd/company/history/yprice/?scode=6857）；來源2：未找到（Yahoo Japan、kabutan 歷史頁未開啟）。EPS 來源1：irbank；來源2：公司決算（見上表）。companiesmarketcap（美元計價 TTM 口徑，與日圓自算不可直接比較）僅作對照，差距見 H、I 欄。")
V.cell(33, 1, "3. 分位（只含有價格的年度）").font = BOLD
hdr(V, 34, ["統計", "P/E 當期（自算）", "P/E 前期（自算）", "cmc P/E（參考）", "用途"])
for i, (nm, p, use) in enumerate((("25 分位", 0.25, "見頂年"), ("中位數", 0.5, "中段"), ("75 分位", 0.75, "谷底")), 35):
    V.cell(i, 1, nm); V.cell(i, 5, use)
    V.cell(i, 2, f"=PERCENTILE(E20:E29,{p})"); V.cell(i, 3, f"=PERCENTILE(F20:F29,{p})"); V.cell(i, 4, f"=PERCENTILE(G20:G29,{p})")
    for c in "BCD": V[f"{c}{i}"].number_format = "0.0"
V.cell(38, 1, "年數"); V["B38"] = "=COUNT(E20:E29)"; V["C38"] = "=COUNT(F20:F29)"; V["D38"] = "=COUNT(G20:G29)"
defname("pe_peak", "Val!$B$35"); defname("pe_mid", "Val!$B$36"); defname("pe_trough", "Val!$B$37")
V.cell(40, 1, "敏感度：補入 2016 年（價格未找到）。以假設的 2016 P/E 重算 25 分位（當期口徑）").font = BOLD
hdr(V, 41, ["2016 P/E 假設", "25 分位", "中位數", "75 分位"])
for i, x in enumerate((10, 20, 30), 42):
    V.cell(i, 1, x).font = BLUE
V.cell(46, 1, "輔助：10 個值（2016 以假設取代）").font = BOLD
for j, x in enumerate((10, 20, 30)):
    col = "BCD"[j]
    V.cell(47, 2 + j, f"=A{42+j}")
    for k in range(10):
        rr = 48 + k
        V[f"{col}{rr}"] = f'=IF(E{20+k}="",{col}$47,E{20+k})'; V[f"{col}{rr}"].number_format = "0.0"
    V[f"B{42+j}"] = f"=PERCENTILE({col}48:{col}57,0.25)"; V[f"C{42+j}"] = f"=PERCENTILE({col}48:{col}57,0.5)"; V[f"D{42+j}"] = f"=PERCENTILE({col}48:{col}57,0.75)"
    for c in "BCD": V[f"{c}{42+j}"].number_format = "0.0"

V.cell(59, 1, "4. 循環位置（設備股共用規則；見頂年度讀設備市場模型，不在本模型另訂）").font = BOLD
V.cell(60, 1, CP.rule_text(PEAK))
hdr(V, 61, ["情境", "設備市場見頂年度", "FY2027 對照日曆年", "判定", "倍數", "v0.2 判定（對照）", "v0.2 倍數", "規則"])
for i, s in enumerate(SC, 62):
    V.cell(i, 1, s); V.cell(i, 2, f"=pk_{s}"); V.cell(i, 3, 2027).font = BLUE
    V.cell(i, 4, "=" + CP.excel_rule(f"C{i}", f"pk_{s}"))
    V.cell(i, 5, f'=IF(D{i}="見頂",pe_peak,IF(D{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 6, "見頂"); V.cell(i, 7, "=pe_peak").number_format = "0.0"; V.cell(i, 8, "tools/cycle_position.py")
V.cell(66, 1, "5. 隱含股價（日圓）：規則值＝FY2027 EPS × 依循環位置的倍數；其餘倍數僅供參考").font = BOLD
hdr(V, 67, ["情境", "EPS 年度", "EPS", "倍數類型", "倍數", "隱含股價", "相對現價", "備註"])
vr = 68; VAL = {}
for s in SC:
    for y, typ, nm in ((2027, f"E{62 + SC.index(s)}", "rule"), (2027, "pe_peak", "參考：25 分位（v0.2 規則值）"), (2027, "pe_mid", "參考：中位數"), (2027, "pe_trough", "參考：75 分位"),
                       (2028, "pe_peak", "參考：25 分位"), (2028, "pe_mid", "參考：中位數"), (2028, "pe_trough", "參考：75 分位")):
        V.cell(vr, 1, s); V.cell(vr, 2, f"FY{y}"); V.cell(vr, 3, f"=PL!{FYC[y]}{SUMR[(s,'eps')]}").number_format = "#,##0"
        V.cell(vr, 4, nm if nm != "rule" else f'="規則值："&D{62 + SC.index(s)}'); V.cell(vr, 5, f"={typ}").number_format = "0.0"
        V.cell(vr, 6, f"=C{vr}*E{vr}").number_format = "#,##0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
        if nm == "rule":
            for c in range(1, 9): V.cell(vr, c).fill = KEY
            VAL[s] = vr
        vr += 1
V.cell(vr + 1, 1, "6. 目前遠期 P/E").font = BOLD
hdr(V, vr + 2, ["項目", "數值", "備註"])
fr = vr + 3
for nm, f, fm, note in (("現價（2026-10-23 收盤）", "=px", "#,##0", "Yahoo!ファイナンス、株探"),
                        ("遠期 P/E：stockanalysis", "=sa_fpe", "0.0", "stockanalysis 顯示 37.7x"),
                        ("遠期 P/E：現價 ÷ 公司指引 EPS（911.64）", "=px/g_eps", "0.0", "自算：現價 ÷ 公司指引 EPS"),
                        ("  兩者差距", f"=B{fr+2}-B{fr+1}", "0.0", "stockanalysis 遠期 P/E 為 2026-10-21 價格 ¥34,440 所計，與現價日期不同；其遠期 EPS 為共識，非公司指引"),
                        ("現價 ÷ 模型 FY2026 EPS（A）", f"=px/PL!{FYC[2026]}{SUMR[('A','eps')]}", "0.0", ""),
                        ("現價 ÷ 模型 FY2027 EPS（A／B／C）", f"=px/PL!{FYC[2027]}{SUMR[('A','eps')]}", "0.0", "B、C 見右"),
                        ("現價隱含倍數 ÷ 規則倍數（情境 A）— 以 FY2027 A EPS", f"=B{fr+5}/E62", "0.00", ">1＝現價高於規則給的倍數")):
    V.cell(fr, 1, nm); V.cell(fr, 2, f).number_format = fm; V.cell(fr, 3, note); fr += 1
V.cell(fr - 2, 3, f"=px/PL!{FYC[2027]}{SUMR[('B','eps')]}").number_format = "0.0"; V.cell(fr - 2, 4, f"=px/PL!{FYC[2027]}{SUMR[('C','eps')]}").number_format = "0.0"
# 0. 目標價與上檔空間（v0.2.3；Andy 2026-10-25 要求；列標籤依工作單逐字）。放在 Val 既有內容之後，不位移既有列；HTML 估值章第一個區塊讀這裡
tp0 = fr + 1
V.cell(tp0, 1, "0. 目標價與上檔空間").font = BOLD
hdr(V, tp0 + 1, ["項目", "數值", "", "", "說明"])
V.cell(tp0 + 2, 1, "目標價（12 個月，情境加權）"); V.cell(tp0 + 2, 2, f"=tp_wA*B{tp0 + 4}+tp_wB*B{tp0 + 5}+tp_wC*B{tp0 + 6}").number_format = "#,##0"
V.cell(tp0 + 2, 5, "權重 × 各情境隱含價（第 5 節規則值）")
V.cell(tp0 + 3, 1, "上檔空間（目標價 ÷ 現價 − 1）"); V.cell(tp0 + 3, 2, f"=B{tp0 + 2}/px-1").number_format = "0.0%"; V.cell(tp0 + 3, 5, "現價＝In 頁 px（2026-10-23 收盤）")
for k_, (s_, lab) in enumerate((("A", "情境 A 隱含價（樂觀）"), ("B", "情境 B 隱含價（基準）"), ("C", "情境 C 隱含價（下檔）"))):
    rr_ = tp0 + 4 + k_
    V.cell(rr_, 1, lab); V.cell(rr_, 2, f"=F{VAL[s_]}").number_format = "#,##0"
    V.cell(rr_, 5, f'="FY2027 EPS × "&D{VAL[s_]}&" "&TEXT(E{VAL[s_]},"0.0")&"x（第 5 節 F{VAL[s_]}）"')
V.cell(tp0 + 7, 1, "情境權重 A／B／C")
for j_, s_ in enumerate("ABC"): V.cell(tp0 + 7, 2 + j_, f"=tp_w{s_}").number_format = "0%"
V.cell(tp0 + 7, 5, f'="合計 "&TEXT(B{tp0 + 7}+C{tp0 + 7}+D{tp0 + 7},"0%")&"；In 頁 tp_wA／tp_wB／tp_wC（Assumed，Andy 可改）"')
V.cell(tp0 + 8, 1, "目標價方法")
V.cell(tp0 + 8, 2, "P/E：各情境 FY2027 EPS × 十年（2016–2025 年底）P/E 依循環位置取分位（設備股共用規則：A、B 中段＝中位數，C 見頂＝25 分位；2016 年底價格未找到，實為 9 年）。"
                 "十年倍數多數來自 AI 前（2017–2022 年 7–24x），偏誤方向為低估；但 2023 年 EPS 低谷使該年 56.8x 偏高、拉高 75 分位（不影響中位數與 25 分位）。隱含價約為 FY2027 年底（2028-03）的價格、未折現")
V.column_dimensions["A"].width = 46
for c in "BCDEFGHIJ": V.column_dimensions[c].width = 16
V.column_dimensions["G"].width = 22; V.column_dimensions["K"].width = 80

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A）：對 FY2027、FY2028 營收與 EPS 的影響。每個案例是一個完整的推導區塊（本頁下方，與 PL 同一引擎）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收 ¥B", "FY2027 EPS ¥", "相對基準", "FY2028 營收 ¥B", "FY2028 EPS ¥", "相對基準"])
cases = [("市占 −5 個百分點", "SoC 市占 61%", dict(sh="sh_soc_lo")), ("市占 +5 個百分點", "SoC 市占 71%", dict(sh="sh_soc_hi")),
         ("平行測試效率每年 −5%（Assumed）", "每加權顆測試支出每年 −5%（2027 起）", dict(eff="eff_par_lo")), ("平行測試效率每年 −10%（Assumed）", "每年 −10%", dict(eff="eff_par_hi")),
         ("每顆測試時間係數 低", "各平台 tt 取低值", dict(tt="_lo")), ("每顆測試時間係數 高", "各平台 tt 取高值", dict(tt="_hi")),
         ("營業費用變動比例 低（35%）", "FY2025 營業費用中變動部分 35%（固定 65%）", dict(F="(1-phi_lo)*opex25", v="phi_lo*opex25/a25_rev")),
         ("營業費用變動比例 高（75%）", "變動 75%（固定 25%）", dict(F="(1-phi_hi)*opex25", v="phi_hi*opex25/a25_rev")),
         ("CY2026 SoC 市場 低（10.5）", "校準錨 10.5", dict(mk26="mk_soc26_lo")), ("CY2026 SoC 市場 高（11.5）", "校準錨 11.5", dict(mk26="mk_soc26_hi")),
         ("非 AI SoC 低（3.7）", "AI 占比較高 → 每顆支出較高", dict(nonai="nonai26_lo")), ("非 AI SoC 高（4.9）", "AI 占比較低 → 每顆支出較低", dict(nonai="nonai26_hi")),
         ("記憶體市占 −5pt", "", dict(shm="sh_mem_lo")), ("記憶體市占 +5pt", "", dict(shm="sh_mem_hi")),
         ("記憶體 HBM 比重 低（30%）", "", dict(w="w_hbm_lo")), ("記憶體 HBM 比重 高（70%）", "", dict(w="w_hbm_hi")),
         ("匯率 ¥140", "", dict(fx="fx_lo")), ("匯率 ¥160", "", dict(fx="fx_hi")),
         ("淨利÷營業利益 低（0.72）", "", dict(nioi="nioi_lo")), ("淨利÷營業利益 0.780（指引隱含）", "含戰略投資評價益（一次性），不當基準", dict(nioi="nioi_hi"))]
SE.cell(4, 1, "基準（情境 A）"); SE.cell(4, 2, "PL 區塊 A")
def put_case(i, name, setting, R):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    SE[f"C{i}"] = f"=PL!{FYC[2027]}{R['rev']}" if R is PB["A"] or R is PB["B"] or R is PB["C"] else f"={FYC[2027]}{R['rev']}"
    pre = "PL!" if (R is PB["A"] or R is PB["B"] or R is PB["C"]) else ""
    SE[f"C{i}"] = f"={pre}{FYC[2027]}{R['rev']}"; SE[f"D{i}"] = f"={pre}{FYC[2027]}{R['eps']}"; SE[f"F{i}"] = f"={pre}{FYC[2028]}{R['rev']}"; SE[f"G{i}"] = f"={pre}{FYC[2028]}{R['eps']}"
    SE[f"E{i}"] = f"=D{i}/D$4-1"; SE[f"H{i}"] = f"=G{i}/G$4-1"
    for c in "CF": SE[f"{c}{i}"].number_format = "#,##0"
    for c in "DG": SE[f"{c}{i}"].number_format = "#,##0.0"
    for c in "EH": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"])
SB = {}
base_row = 50
for n, (nm, setting, ov) in enumerate(cases):
    r0_ = base_row + n * 32
    SB[nm] = engine(SE, r0_, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm])
nn = 5 + len(cases)
put_case(nn, "顆數情境 B", "2028 ×0.8、2029 ×0.85；記憶體 B 路徑；服務／其他系統成長 B", PB["B"])
put_case(nn + 1, "顆數情境 C", "2028–29 ×0.6；市占回落、毛利率下調", PB["C"])
SE.cell(nn + 3, 1, "讀法：FY2027 以上各案例的差異主要來自 CY2027 市場（FY2027＝3/4×CY2027＋1/4×CY2028）；顆數情境 B、C 的 2028 係數只影響 FY2027 的 1／4，所以對 FY2027 影響小、對 FY2028 才明顯。")
SE.column_dimensions["A"].width = 40; SE.column_dimensions["B"].width = 46
for c in "CDEFGH": SE.column_dimensions[c].width = 15

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司指引、共識、公司公布的市場規模、SEMI。只作事後對照，不為對上數字而調整參數。差距 > 10% 的原因見 README"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.0", pt=False):
    global rr_
    RC.cell(rr_, 1, label); RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        L, M = chr(64 + cv), "B"
        RC.cell(rr_, cd, f"={L}{rr_}-{M}{rr_}" if pt else f"={L}{rr_}/{M}{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y=2026: (lambda s: f"=PL!{FYC[y]}{SUMR[(s, key)]}")
RC.cell(rr_, 1, "A. FY2026 vs 公司指引（2026-07-29 上修；校準時未使用指引）").font = BOLD; rr_ += 1
rec("營收", "=g_rev", pl("rev"), "公司 Q1 FY26 決算短信")
rec("SoC 測試機", "=g_soc", pl("rsoc"), "同上")
rec("記憶體測試機", "=g_mem", pl("rmem"), "同上（公司原文未開啟）")
rec("營業利益", "=g_oi", pl("oi"), "同上")
rec("營業利益率（差距以百分點計）", "=g_oi/g_rev", pl("oim"), "同上（49.4%）", fmt="0.0%", pt=True)
rec("淨利", "=g_ni", pl("ni"), "同上")
rec("EPS（日圓）", "=g_eps", pl("eps"), "同上（911.64）")
RC.cell(rr_, 1, "B. FY2026 vs 共識（stockanalysis 2026-10-07）").font = BOLD; rr_ += 1
rec("營收", "=cons_rev26", pl("rev"), "stockanalysis（¥1.76T）")
rec("營業利益", "=cons_oi26", pl("oi"), "stockanalysis（¥886.4B）")
RC.cell(rr_, 1, "C. 市場規模與市占").font = BOLD; rr_ += 1
cy = lambda key, y: (lambda s: f"=PL!{cyc(y)}{PB[s][key]}")
rec("SoC 測試機市場 CY2026（US$B）", "=mk_soc26", cy("soc", 2026), "公司：10.5–11.5（中點）", "校準恆等式：模型 CY2026 以公司中點為錨，差距必為 0，不是獨立檢驗", fmt="0.00")
rec("記憶體測試機市場 CY2026（US$B）", "=mk_mem26", cy("mem", 2026), "公司：2.5–3.0（中點）", "同上，錨定", fmt="0.00")
rec("Advantest FY2025 隱含 SoC 市占（FY2025 實績 ÷ 市場）", "=sh_soc", lambda s: f"=PL!C{PB[s]['sh']}", "公司自估 CY2025 66%", "獨立檢驗：實績與公司市場規模推得的 FY 口徑市占，對照公司自估的 CY 口徑", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "D. 測試設備合計（SoC＋記憶體測試機）vs SEMI 測試設備（markets/equipment 的 Mkt 頁，SEMI 2026-07）").font = BOLD; rr_ += 1
for y, nm in ((2026, "semi_test26"), (2027, "semi_test27"), (2028, "semi_test28")):
    rec(f"CY{y}：SoC＋記憶體測試機市場（US$B）", f"={nm}", (lambda yy: (lambda s: f"=PL!{cyc(yy)}{PB[s]['soc']}+PL!{cyc(yy)}{PB[s]['mem']}"))(y),
        "SEMI 2026-07（equipment 成品）", "SEMI 測試設備含探針台、分選機與其他 ATE，口徑較寬；2027 為 SEMI 內插值" if y == 2026 else "", fmt="0.00")
RC.cell(rr_, 1, "E. 其他資訊（非對帳）").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "FY26 Q1 實績營收 ÷ 模型 FY2026 全年（A）"); RC.cell(rr_, 3, f"=q1_rev/PL!{FYC[2026]}{SUMR[('A','rev')]}").number_format = "0.0%"; RC.cell(rr_, 10, "AI 需求逐季上升，Q1 低於 25% 屬正常"); rr_ += 1
RC.cell(rr_, 1, "FY26 Q1 SoC×4 ÷ 模型 FY2026 SoC（A）"); RC.cell(rr_, 3, f"=q1_soc*4/PL!{FYC[2026]}{SUMR[('A','rsoc')]}").number_format = "0.0%"; rr_ += 1
RC.column_dimensions["A"].width = 56
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 38; RC.column_dimensions["J"].width = 80

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文頁面在建置環境無法開啟；數字取自搜尋引擎對這些文件的摘錄，並與第二來源交叉；原文待人工複核）"; SO["A1"].font = BOLD
srcs = [("公司：FY2025 決算（英文，2026-04-27）", "https://www.advantest.com/en/news/2026/a81o6o0000000hgw-att/E_FR_FY2025_FN.pdf"),
 ("公司：FY2025 決算短信（日文）", "https://www.advantest.com/ja/news/2026/a81o6o0000000heh-att/J_FR_FY2025_FN.pdf"),
 ("公司：FY2025 Q4 Investor Note（2026-04-27）", "https://www.advantest.com/document/en/investors/ir-library/result/JE_BIZ_260427_note.pdf"),
 ("公司：FY2025 Q3 Investor Note（2026-01-28）", "https://www.advantest.com/document/en/investors/ir-library/result/JE_BIZ_260128_note.pdf"),
 ("公司：FY2025 Q2 Investor Note（2025-10-28）", "https://www.advantest.com/document/en/investors/ir-library/result/E_BIZ_251028_note.pdf"),
 ("公司：FY2025 Q1 Investor Note（2025-07-29）", "https://www.advantest.com/document/en/investors/ir-library/result/E_BIZ_250729_note.pdf"),
 ("公司：FY2026 Q1 Investor Note（2026-07-29）", "https://www.advantest.com/document/en/investors/ir-library/result/JE_BIZ_260729_note.pdf"),
 ("公司：FY2026 Q1 決算短信（日文，2026-07-29）", "https://www.advantest.com/ja/news/2026/qnpuno0000000cnr-att/J_FR_FY2026_1Q.pdf"),
 ("公司：FY2025 Q2 決算短信（日文）", "https://www.advantest.com/ja/news/2025/g2l53r0000000oc1-att/J_FR_FY2025_2Q.pdf"),
 ("公司：IR 行事曆（FY2026 Q2 預定 2026-10-28 15:30 JST）", "https://www.advantest.com/en/investors/ir-calendar/"),
 ("Investing.com：FY25 簡報摘要（營收分項、毛利率 64.3%）", "https://www.investing.com/news/company-news/advantest-fy25-slides-record-margins-as-ai-testing-demand-surges-93CH-4639589"),
 ("Investing.com：Q3 FY25 簡報摘要", "https://www.investing.com/news/company-news/advantest-q3-fy25-slides-record-sales-drive-guidance-upgrade-on-ai-demand-93CH-4470612"),
 ("Investing.com：Q1 FY26 簡報摘要（SoC 259.6、記憶體 48.0、指引）", "https://www.investing.com/news/company-news/advantest-q1-fy26-slides-record-profit-raised-forecast-on-ai-demand-93CH-4821153"),
 ("Investing.com：Q1 FY26 法說逐字稿（毛利率 66–67%、R&D 110.0）", "https://www.investing.com/news/transcripts/earnings-call-transcript-advantest-lifts-fy2026-outlook-after-record-q1-93CH-4820868"),
 ("MarketScreener：公司 FY2025 決算（Services and Others 109.2）", "https://www.marketscreener.com/news/advantest-financial-results-e-fr-fy2025-fn-ce7f59dcdc8df223"),
 ("kabutan：FY2026 指引上修", "https://kabutan.jp/news/?b=k202607290036"),
 ("gyokaidigest：FY2026 Q1 分析", "https://www.gyokaidigest.com/companies/advantest/report/2027-Q1"),
 ("松井證券：決算情報（EPS 515.15）", "https://finance.matsui.co.jp/stock/6857/settlement/index"),
 ("irbank：EPS 推移（來源1）", "https://irbank.net/E01950/valuation"),
 ("日經新聞：過去 10 年株價推移（年底收盤）", "https://www.nikkei.com/nkd/company/history/yprice/?scode=6857"),
 ("companiesmarketcap：P/E（年底，2016–2024）", "https://companiesmarketcap.com/advantest/pe-ratio/"),
 ("Yahoo Finance：key statistics（2025-12-31 TTM P/E 55.18）", "https://finance.yahoo.com/quote/6857.T/key-statistics/"),
 ("stockanalysis：Advantest（現價、遠期 P/E）", "https://stockanalysis.com/quote/tyo/6857/"),
 ("公司：FY2017–FY2020 決算（EPS 第二來源）", "https://www.advantest.com/document/en/investors/ir-library/result/E_FR_FY2020_FN.pdf"),
 ("X（jukan05）：SoC 市占 56%→66%（二手轉述）", "https://x.com/jukan05/status/2048660802736140321"),
 ("市場層：markets/memory/market.py（UNITS）；tsmc／skhynix／equipment 成品（本 repo）", ""),
 ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 70; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("Advantest（6857.T）收支與估值模型 v0.3（AI 半導體 Project 步驟 5：設備與材料）　2026-10-25", True),
 ("v0.2.2（2026-10-23）：引用的上游成品檔名更新（SK 海力士 v0.4.1、美光 v0.3.1、三星 v0.2.1、台積電 v0.2.1 升版）；所有數值不變。", False),
 ("模型命題：Advantest 的盈餘由 AI 加速器的 SoC 測試需求決定。每顆晶片的測試時間隨複雜度上升，Advantest 握有約三分之二的市占。股價取決於加速器顆數的成長在 FY2027 之後放緩時，測試機的投資會不會跟著回落。", False), ("", False),
 ("會計年度：4 月到隔年 3 月（FY2026＝2026-04～2027-03）。金額以日圓（¥B）為主，另列美元；匯率基準 ¥150（公司第二季起的假設）放 In 頁。FY2026 第二季預定 2026-10-28 公布，建置日（2026-10-22）尚未公布，故不納入。", False), ("", False),
 ("驅動因子 → 推導量", True),
 ("SoC 測試機市場（日曆年）＝加速器顆數（markets/memory）× 每顆測試時間係數 × 每加權顆測試支出（以公司 CY2026 市場規模校準）＋非 AI SoC；日曆年轉會計年度：3/4、1/4 加權 → × Advantest 市占（基準 66%）× 匯率 → SoC 營收", False),
 ("記憶體測試機：公司 CY2026 市場規模為錨 × [w × HBM 出貨位元指數（markets/memory）＋(1−w) × DRAM 設備支出指數（markets/equipment）] × 市占（FY2025 實績推得）", False),
 ("其他系統、服務與其他：前一年 × 成長率（Assumed）。毛利率＝截距＋斜率 × SoC 占比（FY2025 與 FY2026 指引兩點推導）；營業費用＝固定＋變動 × 營收（兩點推導）；不寫死利益率", False),
 ("情境 A／B／C：顆數係數讀台積電模型成品（2028 ×1.0／0.8／0.6，2029 ×1.0／0.85／0.6）；情境 C 另加 Teradyne 搶回市占（FY2028 61%、FY2029 56%）與價格壓力", False),
 ("Val：P/E 規則——10 年（2016–2025 年底）P/E 分位，見頂年 25 分位、中段中位數、谷底 75 分位；循環位置依設備股共用規則（各情境見頂年度讀設備市場模型：" + "、".join(f"{s_} {PEAK[s_]}" for s_ in SC) + " → FY2027：" + CP.fy2027_text(PEAK) + "）。Recon：對公司指引、共識、公司市場規模、SEMI。Sens：市占、測試時間係數、顆數情境、營業費用變動比例等", False), ("", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司原文（advantest.com）在建置環境無法開啟，數字取自搜尋引擎摘錄，原文待人工複核。", False),
 ("已知限制：(1) 每加權顆測試支出由「公司 CY2026 市場規模 − 非 AI 市場」校準，市場規模為測試機廠自估（Interested-party），非 AI 以 AI 前（CY2021–23）Teradyne 公布的 SoC 測試機市場為參照（Analogy，3.7–4.9）；平行測試效率提升不在基準，只在 Sens 列示；(2) 每顆測試時間係數無實證，只決定世代間相對權重；(3) HBM 出貨位元自 markets/memory v0.4 起隨情境顆數係數變動（2028–29）；供給三情境相同、單價假設未重估，B、C 的 HBM 位元與單價同時下修；記憶體測試機的情境差異來自 HBM 出貨位元與 DRAM 設備支出兩條線；(4) 無測試機效率（平行測試）提升的建模，可能高估 CY2027 之後市場；(5) CY2030 無市場層資料，持平假設；(6) 10 年 P/E 缺 2016 年底收盤價。", False),
 ("v0.2.3（2026-10-25）：Val 頁新增「0. 目標價與上檔空間」（放在 Val 既有內容之後，不位移既有列）：目標價＝In 頁權重 tp_wA／tp_wB／tp_wC（預設 A 0.3／B 0.5／C 0.2，Assumed，Andy 可改）× 三情境規則值隱含價；引用的上游成品檔名更新（SK 海力士 v0.4.2）；其餘數值不變。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

CFG_V2 = dict(cols=FYC, scn={"A": "順利", "B": "平平", "C": "不如預期"}, peers=[('泰瑞達（Teradyne）', 'TER', 32.52, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('東京威力科創（Tokyo Electron）', '8035.T', 27.09, '2026-10-21', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('Lasertec', '6920.T', 34.02, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('Disco', '6146.T', 26.01, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）'),('科磊（KLA）', 'KLAC', 26.65, '2026-10-23', 'stockanalysis.com statistics：Forward PE（分析師共識；價格日期如左）')], fwd={2027: "5/12", 2028: "7/12"}, fwd_note='約 2027-10 時的未來 12 月（2027-11～2028-10）≈ 5／12 × FY2027（2027-04～2028-03）＋7／12 × FY2028（會計年度 4–3 月，依月份調整）', a3='目標價（DCF 與 Comparable 加權）：見本頁第 {row} 列起「0. 目標價」（估值 v2 新增，放在既有內容之後以免位移既有列）；原「0. 目標價與上檔空間」改名為「舊法：規則值情境加權（參考）」',
    cur='¥',
    unit_rev='¥B',
    unit_ni='¥B',
    fmt='#,##0',
    fmt_rev='#,##0',
    band_fmt='#,##0',
    roe=('2024-03～2026-03 期', [0.1557, 0.3438, 0.5765], 'stockanalysis 比率頁（2026-10-25 讀取）：2024-03 期 15.57%、2025-03 期 34.38%、2026-03 期 57.65%', '單一來源；近 12 月 56.89%（同站）'),
    nim=('FY2016–FY2025＝2017-03～2026-03 期', [(155916, 14201), (207223, 18103), (282456, 56993), (275894, 53532), (312789, 69787), (416901, 87301), (560191, 130400), (486507, 62290), (779707, 161177), (1128610, 375353)], 'finboard.jp 損益計算書（百萬日圓；營收、純益；2026-10-25 讀取）', '第二來源：the-shashi.com 營收逐年相同（純益未列）；2026-03 期與 Q 頁實績一致'),
    cons_tp=(42686, 'stockanalysis：21 位分析師平均 ¥42,686（中位數 42,900；區間 23,000–55,000），資料更新 2026-10-07', '第二來源：未找到；早於 2026-10-28 第二季財報'),
    cons_feps=('=34440/sa_fpe', '共識前瞻 EPS＝stockanalysis 當時價格 ¥34,440（2026-10-21）÷ stockanalysis 遠期 P/E（In 頁 sa_fpe）的隱含值，口徑約 FY2026；比本模型 12 個月後前瞻 EPS 早約一年半，EPS 因素因此偏大', 'Derived'),
    peer_note='測試與日本半導體設備同業：泰瑞達（直接對手）、東京威力科創、Lasertec、Disco、KLA',
    dcf_note='；會計年度 4–3 月，FY2027＝2027-04～2028-03 仍以 t＝1 折現（略高估）',
    pe10='E20:E29',
    pe10_note='現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 2 節當期口徑，2016 缺價格）的分位；只作檢查，不進目標價',
    px_note='現價＝In 頁 px（2026-10-23 收盤）',
    method='DCF：股東現金流＝EPS ×（1 − g／ROE）（模型無折舊與資本支出項目），r 10%、g 4%、ROE 為近三年平均，FY2027–32 明確期（FY2030 起收斂法）＋永續。Comparable：泰瑞達、東京威力科創、Lasertec、Disco、KLA 前瞻 P/E 中位數 × 12 個月後前瞻 EPS。偏誤：FY2028 起情境 B、C 的顆數係數只是 Assumed；十年平均淨利率（約一成九）遠低於 AI 期間，情境 C 收斂偏保守；同業前瞻 P/E 為 AI 熱度下的倍數，偏高估。',
    readme=['v0.3（2026-10-25）估值 v2：Val 頁新增「0. 目標價」（放在既有內容之後，不位移既有列）＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，情境權重 A 0.3／B 0.5／C 0.2）＋0.5 × Comparable 目標價（同業前瞻 P/E 中位數 × 12 個月後前瞻 EPS）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解、歷史分位檢查。原「0. 目標價與上檔空間」改名「舊法：規則值情境加權（參考）」，不再叫目標價。', '預估期間延伸到 FY2032：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（PL 頁底部延伸區塊，灰底）。情境對應：A＝順利、B＝平平、C＝不如預期（原為顆數情境）。本版未做先驗，無 Andy 合理 P/E 對照。現價更新為 2026-10-23 收盤 ¥32,340（Yahoo!ファイナンス、株探；markets/invest_map/data/prices.csv）。'])

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

wb.move_sheet("In", offset=0)
order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(sys.argv[1])

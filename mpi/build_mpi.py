# 旺矽（6223.TWO；MPI）個股模型 v0.1（AI 半導體 Project：測試介面）
# 用法：python3 mpi/build_mpi.py raw.xlsx ；再以 LibreOffice headless 重算
# 樣板：kyec/build_kyec.py（測試介面／測試服務；加速器顆數 × 每顆測試時間），工作表結構與 Val 頁列標籤同 amkr/、ase/。
# 會計年度＝日曆年，FY(t)＝CY(t)。金額：十億新台幣（NT$B）；EPS：新台幣。口徑：歸屬母公司淨利。
# 行業變數（加速器顆數、情境係數、每顆測試時間係數、非 AI SoC 成長、SEMI 測試設備、設備情境係數、CPO 滲透率）唯讀取自 Siliconomics 快照。
import sys, os
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP
import valuation_rule as VR

VER, DATE = "v0.1", "2026-10-31"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
SILV = f"Siliconomics 快照 {SIL.VERSION}（{SIL.SNAP['source']['file']}）"

# ---------------- Siliconomics：平台別顆數 × 每顆測試時間係數 → 晶圓測試量 ----------------
PLATS = ("bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth")
NV = ("bw", "rb", "ru", "hp")
UP = {p: SIL.NAMES[f"SIF_UnitsPlat_{p}"] for p in PLATS}          # [None, 2026, 2027, 2028, 2029]
UT = SIL.NAMES["SIF_UnitsPlat_Total_A"]
UNITS = {26: UT[1], 27: UT[2], 28: UT[3], 29: UT[4]}
TT = {p: SIL.base(f"ts_tt_{p}") for p in PLATS}
EXO = {"amd": 0.3, "tpu": 1.0, "trn": 0.8, "oth": 1.0}              # 旺矽曝險（NVIDIA 另以 ex_nv 輸入）；Assumed
NVTT = {y: sum(UP[p][j] * TT[p] for p in NV) / UT[j] for j, y in enumerate((26, 27, 28, 29), 1)}
ASTT = {y: sum(UP[p][j] * TT[p] * EXO[p] for p in EXO) / UT[j] for j, y in enumerate((26, 27, 28, 29), 1)}
SEMI = SIL.NAMES["SIF_SEMI_Test"]       # 2025–2029
SOC = {s: SIL.NAMES[f"SIF_TestSoCMkt_{s}"] for s in SC}

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B）；顆數：百萬顆；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = ("標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間（放 Sens），情境專屬列不給區間。SemiAnalysis 未使用。"
            "公司原文＝公開資訊觀測站月營收與季報；法說內容經富果、口袋學堂等摘要；公司指引與說法為 Interested-party；券商與媒體為 Interested-party")
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
FG1 = "https://blog.fugle.tw/post/earnings-call-6223-2026-05-15"
FG4 = "https://blog.fugle.tw/post/earnings-call-6223-2026-03-03"
PK2 = "https://www.pocket.tw/school/articles/6433/"
TN9 = "https://finance.technews.tw/2026/10/08/mpi-6223-202609-financial-report/"
MW = "https://www.moneyweekly.com.tw/_Article?AID=254951"
UDN8 = "https://money.udn.com/money/amp/story/5607/9703462"
BN = "https://www.bnext.com.tw/article/91242/mpi-probe-card-ten-thousand"
FUBON = "https://fubon-ebrokerdj.fbs.com.tw/Z/ZC/ZCX/ZCX_6223.djhtm"
WG = "https://www.wantgoo.com/stock/6223/enterprise-value/price-to-earning-river"

sec(f"A. AI 加速器顆數、情境係數與晶圓測試量（讀 {SILV}；唯讀，不自行重估）")
for y in (26, 27, 28, 29):
    inp(f"units_{y}", f"AI 加速器顆數 20{y}（情境 A 基準）", "百萬顆", None, UNITS[y], None, "Derived",
        "Siliconomics SIF_UnitsPlat_Total_A（八類加速器合計）" if y == 26 else "同上", "與台積電、日月光、京元電模型同一份" if y == 26 else "同上")
for s in SC:
    for y in (28, 29):
        inp(f"um{y}_{s}", SIL.V[f"um{y}_{s}"]["label"], "x", None, SIL.base(f"um{y}_{s}"), None, "Assumed", f"Siliconomics um{y}_{s}（同字母行業情境）",
            "A＝基準；B＝AI 資本支出消化；C＝明顯下修" if (s == "A" and y == 28) else "同上")
for y in (26, 27, 28, 29):
    inp(f"nvtt_{y}", f"NVIDIA 加速器的測試時間加權顆數 ÷ 總顆數 20{y}", "x", None, round(NVTT[y], 6), None, "Derived",
        "Siliconomics 平台顆數 SIF_UnitsPlat_bw／rb／ru／hp × 每顆測試時間係數 ts_tt_*（Blackwell＝1）" if y == 26 else "同上", "晶圓測試量的代理：測試時間越長、探針卡耗用越多（Analogy：京元電、Advantest 模型同係數）" if y == 26 else "同上")
    inp(f"astt_{y}", f"ASIC 與 AMD 的測試時間 × 旺矽曝險加權顆數 ÷ 總顆數 20{y}", "x", None, round(ASTT[y], 6), None, "Derived",
        "Siliconomics 平台顆數 × ts_tt_*；旺矽曝險 AMD 0.3、TPU 1.0、Trainium 0.8、其他 ASIC 1.0（Assumed，本模型）" if y == 26 else "同上",
        "媒體：旺矽為多數 CSP 委外 AI ASIC 探針卡主要供應商，平均份額逾六成（傳聞；財訊 2026-10-12）" if y == 26 else "同上")
inp("ex_nv", "旺矽對 NVIDIA 加速器的曝險（ASIC＝1）", "x", 0.2, 0.4, 0.7, "Assumed", MW + "：客戶含 Broadcom、Marvell、NVIDIA（媒體）", "NVIDIA 探針卡主力為 FormFactor、Technoprobe（產業常識；量化：未找到）")
inp("g_nonai", "非 AI SoC 市場年增（其他探針卡的行業成長）", "%", *SIL.vals("ts_g_nonai"), "Assumed", f"Siliconomics ts_g_nonai（{SIL.VERSION}）", "其他探針卡＝CPC、手機 AP、網通、車用")
for s in SC:
    for y in (27, 28, 29):
        k = f"cpo_{s}_{y}"
        inp(k, SIL.V[k]["label"], "%", None, SIL.base(k), None, "Assumed", f"Siliconomics {k}（技術替代屬行業情境，同字母）", "旺矽 CPO 測試設備營收依此滲透率相對 2027 情境 A 縮放" if (s == "A" and y == 27) else "同上")
inp("cpo_A_26", SIL.V["cpo_A_26"]["label"] + "（三情境相同）", "%", None, SIL.base("cpo_A_26"), None, "Assumed", "Siliconomics cpo_A_26（B、C 同值）", "—")
for y, i in ((27, 2), (28, 3), (29, 4)):
    inp(f"semi_{y}", f"SEMI 測試設備 20{y} ÷ 2026（情境 A 設備路徑）", "x", None, round(SEMI[i] / SEMI[1], 6), None, "Derived", "Siliconomics SIF_SEMI_Test（2025–2029）" if y == 27 else "同上", "設備分部的行業成長（SoC、記憶體測試設備合計）" if y == 27 else "同上")
for s in SC:
    for y in (28, 29):
        if s == "A":
            inp(f"eqk{y}_A", f"情境 A：設備水準 20{y} ÷ 2027（SEMI 路徑）", "x", None, round(SEMI[y - 25] / SEMI[2], 6), None, "Derived", "Siliconomics SIF_SEMI_Test 推得", "—")
        else:
            inp(f"eqk{y}_{s}", f"情境 {s}：設備水準 20{y} ÷ 2027", "x", None, SIL.base(f"eq_ck_{s}_{y}"), None, "Assumed", f"Siliconomics eq_ck_{s}_{y}（設備情境係數，套用於 2027 水準；與 markets/equipment 同口徑）", "—")

sec("B. 探針卡分部：AI／HPC 探針卡（VPC、MEMS）＝ AI_26 × 晶圓測試量指數 × 旺矽份額與價格相對值 × 每顆內容成長；其他探針卡依非 AI SoC 成長")
inp("ai_sh25", "AI／HPC 探針卡 ÷ 探針卡營收 FY2025", "%", None, 0.55, None, "Assumed", "公司：VPC 需求主要來自 AI／HPC；VPC 產能 2H25 起滿載（2025-03 法說，富果）", "AI 占比公司未揭露：未找到")
inp("ai_sh26", "AI／HPC 探針卡 ÷ 探針卡營收 FY2026（校準）", "%", 0.45, 0.60, 0.70, "Assumed", MW + "：法人預估 VPC 逾探針卡營收六成、高階探針卡約占營收五成（媒體，2026-10-12）", "FY2026 只影響拆分；AI 與其他的成長差異使 FY2027 起受影響")
for s, a, b, c in (("A", 1.05, 1.08, 1.10), ("B", 1.00, 1.00, 1.00), ("C", 0.92, 0.88, 0.85)):
    for y, v in ((27, a), (28, b), (29, c)):
        inp(f"rel{y}_{s}", f"情境 {s}：旺矽份額與價格相對值 20{y}（2026＝1）", "x", None, v, None, "Assumed",
            BN + "：AI／HPC 客戶的垂直 MEMS 探針卡單價可高出平均五成以上；ALETHEIA 估旺矽需擴產約四倍（媒體轉述券商，2026-06-13）" if (s == "A" and y == 27) else "同上",
            "A＝新客戶（美系 CPU、NVIDIA）與價格；B＝維持；C＝Technoprobe、FormFactor 兩年內針產能倍增（同上）後份額與價格下滑" if (s == "A" and y == 27) else "同上")
inp("g_ct", "每顆加速器探針卡內容年增（針數、晶粒面積）", "%", 0.0, 0.05, 0.10, "Assumed", "晶粒面積與 HBM 堆疊增加 → 每片晶圓針數上升（產業常識；量化：未找到）", "公司：2026 年底 VPC 月產能 200 萬針、MEMS 350 萬針（富果 2026-05）")
inp("d_npc", "其他探針卡相對非 AI SoC 市場的超額成長", "%", 0.0, 0.03, 0.06, "Assumed", "公司：CPC 訂單出貨比大於 1（2025-03 法說）；非 AI ASIC 新客戶 4Q26 小量（經濟日報 2026-08-20）", "—")

sec("C. 設備分部（AST、溫控、光電／CPO 測試）")
inp("k_cpo", "CPO 測試設備營收 FY2027（情境 A 滲透率下）", "NT$B", 0.5, 1.2, 2.0, "Assumed", UDN8 + "：CPO 測試設備為 2027 年設備事業主要成長動能；湖口新廠一期供 CPO 機台量產（媒體轉述法人）", BN + "：每增加 10 億元 CPO 相關營收約增 EPS 4–5 元（券商）")
inp("g_oth", "其他營收年增", "%", None, 0.05, None, "Assumed", "其他約占營收 2%（公司 1Q26）", "—")

sec("D. 毛利率、營業費用與淨利（FY2027 起）")
for s, a, b in (("A", 0.600, 0.610), ("B", 0.595, 0.595), ("C", 0.580, 0.565)):
    inp(f"gm27_{s}", f"情境 {s}：毛利率 FY2027", "%", None, a, None, "Assumed", "高盛估 2025 57.4%、2026 59.1%（財訊 2026-10-12）；公司：2027 毛利率續升（經濟日報 2026-08-20）" if s == "A" else "同上",
        "A＝高階 MEMS 比重與價格；C＝擴產折舊與價格競爭" if s == "A" else "同上")
    inp(f"gmL_{s}", f"情境 {s}：毛利率 FY2028–29", "%", None, b, None, "Assumed", "同上", "同上")
inp("gm_adj", "毛利率平移（三情境同時；敏感度用）", "%", -0.02, 0.0, 0.02, "Assumed", "基準 0；只在 Sens 動", "—")
inp("phi", "營業費用成長對營收成長的彈性（FY2027 起）", "x", 0.4, 0.6, 0.9, "Assumed", "FY2026 實測見 Q 頁；公司：研發費用占營收至少 18%（2025-03 法說）", "—")
inp("ni_oi", "歸屬母公司淨利 ÷ 營業利益（FY2027 起）", "x", 0.80, 0.85, 0.90, "Derived", "1Q26 0.96、2Q26 0.85（公司；業外含匯兌）", "稅率與業外未分別揭露：未找到")
inp("sh_g", "股數年變動", "%", None, 0.0, None, "Assumed", "FY2025 95.0 → 2Q26 98.0 百萬股（淨利 ÷ EPS 推得）", "—")

sec("E. FY2026 錨點（1Q、2Q26 實績＋3Q26 營收實績＋假設）")
inp("q3_pcsh", "3Q26 探針卡 ÷ 營收", "%", None, 0.75, None, "Assumed", "1Q26 71.1%、2Q26 74.7%（公司法說）", "3Q26 事業別：未找到（11 月法說）")
inp("q3_oth", "3Q、4Q26 其他 ÷ 營收", "%", None, 0.02, None, "Derived", "1Q26、2Q26 皆 2.0%（公司）", "—")
inp("q3_gm", "3Q26 毛利率", "%", None, 0.594, None, "Interested-party", PK2 + "：公司稱 3Q26 毛利率回到 1Q26 水準（59.4%）、4Q26 再小升", "3Q26 財報：未公布（約 11 月中）")
inp("q4_gm", "4Q26 毛利率", "%", None, 0.600, None, "Interested-party", "同上（4Q26 小升）", "—")
inp("q4_g", "4Q26 營收季增", "%", None, 0.075, None, "Interested-party", PK2 + "：4Q26 營收季增 5–10%（公司 2Q26 法說）；取中點", "3Q26 實績季增 12.8%，高於指引 0–5%")
inp("q4_pcsh", "4Q26 探針卡 ÷ 營收", "%", None, 0.75, None, "Assumed", "同 3Q26", "公司：2026 年設備占比與 1Q26 26.9% 差異不大（2H CPO 出貨小）")
inp("q_opx", "3Q、4Q26 營業費用季增", "%", None, 0.06, None, "Assumed", "1Q26 1.07 → 2Q26 1.25（+18%）", "—")
inp("q_nioi", "3Q、4Q26 淨利 ÷ 營業利益", "x", None, 0.87, None, "Derived", "1H26 2.751 ÷ 3.074＝0.895（公司）；匯兌利益不延續", "—")

sec("F. 公司指引、券商預估（只作 Recon 對帳，不作校準）")
inp("cons_eps26", "法人 EPS FY2026", "NT$", None, 64.79, None, "Interested-party", MW + "：法人預估 2026 EPS 64.79、2027 128.95（財訊 2026-10-12）", "口袋學堂 2026-09-08：券商 60.89–70.84（8/17 報告）")
inp("cons_eps27", "法人 EPS FY2027", "NT$", None, 128.95, None, "Interested-party", "同上", "—")
inp("al_eps27", "ALETHEIA EPS FY2027", "NT$", None, 186.5, None, "Interested-party", BN + "：ALETHEIA 2026 至少 78.1、2027 186.5、2028 286.5（數位時代 2026-06-13）", "單一券商")
inp("al_eps28", "ALETHEIA EPS FY2028", "NT$", None, 286.5, None, "Interested-party", "同上", "—")
inp("cons_rev26", "券商營收 FY2026（區間中點）", "NT$B", None, 22.09, None, "Interested-party", PK2 + "：券商 2026 營收 20.38–23.80（8/17 報告）", "—")

sec("G. 價格與市場倍數")
inp("px", "股價（2026-10-30 收盤，櫃買中心）", "NT$", None, 4610.0, None, "Interested-party", FUBON + "：10-30 收盤 4,610.00（−230.00）", WG + "：10 月收盤 4,610.00（玩股網）")
inp("pe_ttm", "近 12 月 P/E（玩股網，10-30）", "x", None, 103.64, None, "Interested-party", WG, "—")
inp("fpe", "前瞻 P/E（現價 ÷ 法人 FY2027 EPS）", "x", None, "=px/cons_eps27", None, "Derived", "Bloomberg 前瞻倍數待取（fwd_mult_hist.csv 無 6223）", "—")
inp("cons_feps", "共識前瞻 12 月 EPS（ALETHEIA：2／12 × FY2027＋10／12 × FY2028；單一券商）", "NT$", None, "=2/12*al_eps27+10/12*al_eps28", None, "Derived", "與市場共識目標價同一券商（ALETHEIA）；與本模型 12 個月後前瞻 EPS 同口徑", "賣方共識前瞻 EPS：未找到")
inp("cons_tp", "市場共識目標價（以 ALETHEIA 單一券商代替）", "NT$", None, 10000.0, None, "Interested-party", BN + "：ALETHEIA 目標價 4,000 → 10,000（2026-06-13）", "賣方平均目標價：未找到")

sec("H. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_r", "折現率／要求報酬", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%；各市場同一值）", "DCF 與 Sens")
inp("tp_g", "長期成長率", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "成熟倍數與 FY2032 收斂目標（情境 B）")
inp("tp_roe", "ROE（近兩年平均，FY2024–25；FY2023 未找到）", "%", None, "=AVERAGE(2.30/9.2,3.18/AVERAGE(9.2,14.5))", None, "Derived",
    FG4 + "：每股淨值 2024 年底 98 → 2025 年底 148；權益 2025 年底 14.5（富果 2026-05）；淨利 2024 約 2.30（EPS 24.42 × 約 94 百萬股）、2025 3.18", "2024 年以年底權益計（期初未找到）")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（工作單 2026-10-25 11:53 定 DCF／Comparable 各半）", "—")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（情境 B FY2029 與其 × 0.8 的中點）", "%", None, 0, None, "Assumed", "十年平均淨利率未找到：工作單規則以 FY2029 × 0.8 代替", "PL 頁情境 B FY2029（公式於建置時連結）")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（＝情境 C FY2029 × 0.8）", "%", None, 0, None, "Assumed", "同上（十年平均未找到）", "PL 頁情境 C FY2029（公式於建置時連結）")
TP_KEYS += ["tp_r", "tp_g", "tp_roe", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]

URLNAME = {FG1: "2026-05 法說摘要（富果）", FG4: "2026-03 法說摘要（富果）", PK2: "2Q26 財報與法說摘要（口袋學堂 2026-09-08）", TN9: "9 月營收（科技新報 2026-10-08）",
           MW: "財訊 2026-10-12", UDN8: "經濟日報 2026-08-20", BN: "數位時代 2026-06-13（ALETHEIA）", FUBON: "富邦證券 e 點通（6223 日線）", WG: "玩股網本益比河流圖"}
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
    for u_, nm_ in URLNAME.items():
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

# ---------------- Q：實績與 FY2026 錨點（NT$B） ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億新台幣；EPS 為新台幣；股數為百萬股）。藍字＝輸入（公司月營收、季報與法說）；黑字＝公式。事業別＝公司法說揭露的探針卡、設備、其他占比 × 營收"
Q["A2"] = "v0.1（2026-10-31）：FY2026 錨點＝1Q、2Q26 實績＋3Q26 營收實績（月營收：前三季 15.068 − 1H26 9.166）＋3Q26 損益依公司說法（毛利率回到 59.4%）＋4Q26 依指引（季增 5–10% 取 7.5%）。3Q26 財報約 11 月中公布"
hdr(Q, 3, ["項目", "FY2025A", "1Q26（實績）", "2Q26（實績）", "3Q26（營收實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1（公司）", "來源2／說明"])
S1 = "FY2025、1Q26：公司法說（富果 2026-03、2026-05）；2Q26：口袋學堂 2026-09-08；3Q26 營收：公司月營收（科技新報 2026-10-08）"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2="—", ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"G{r}"] = ann or f"=SUM(C{r}:F{r})"; Q[f"G{r}"].number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, s1); Q.cell(r, 10, s2)
qrow(4, "營收", {"B": 13.377, "C": 3.933, "D": 5.233, "E": "=15.068-9.166", "F": "=E4*(1+q4_g)"}, tag="Verified／Interested-party（4Q）", s2="FY2025 13.377（+31.5%）；1H26 9.166；前三季 15.068（+58.0%）")
qrow(5, "探針卡", {"B": "=B4*0.738", "C": "=C4*0.711", "D": "=D4*0.747", "E": "=E4*q3_pcsh", "F": "=F4*q4_pcsh"}, tag="Derived（占比 × 營收）", s2="占比：FY2025 73.8%（富果；口袋學堂另載 72.2%）、1Q26 71.1%、2Q26 74.7%")
qrow(6, "  AI／HPC 探針卡（VPC、MEMS；命題子列）", {"B": "=B5*ai_sh25", "C": "=C5*ai_sh26", "D": "=D5*ai_sh26", "E": "=E5*ai_sh26", "F": "=F5*ai_sh26"}, tag="Assumed（占比）", s2="AI 占比公司未揭露")
qrow(7, "  其他探針卡（CPC 等）", {c: f"={c}5-{c}6" for c in "BCDEF"}, tag="Derived", s1="—")
qrow(8, "設備", {"B": "=B4*0.246", "C": "=C4*0.269", "D": "=D4*0.233", "E": "=E4*(1-q3_pcsh-q3_oth)", "F": "=F4*(1-q4_pcsh-q3_oth)"}, tag="Derived（占比 × 營收）", s2="占比：FY2025 24.6%、1Q26 26.9%、2Q26 23.3%（含 AST、溫控、光電）")
qrow(9, "其他", {"B": "=B4-B5-B8", "C": "=C4*0.020", "D": "=D4-D5-D8", "E": "=E4*q3_oth", "F": "=F4*q3_oth"}, tag="Derived", s2="FY2025 約 1.6%、1Q26 2.0%")
qrow(10, "毛利", {"B": "=B4*0.574", "C": 2.338, "D": "=D4*0.5837", "E": "=E4*q3_gm", "F": "=F4*q4_gm"}, tag="Verified／Interested-party", s2="FY2025 毛利率 57.4%（高盛，財訊 2026-10-12）；2Q26 58.37%（公司）")
qrow(11, "毛利率", {c: f"={c}10/{c}4" for c in "BCDEFG"}, fmt="0.00%", ann="", tag="Derived", s1="—")
qrow(12, "營業費用（毛利 − 營業利益）", {"C": "=C10-C13", "D": "=D10-D13", "E": "=D12*(1+q_opx)", "F": "=E12*(1+q_opx)"}, tag="Derived", s1="—", s2="FY2025 營業利益：未找到")
qrow(13, "營業利益", {"C": 1.273, "D": 1.801, "E": "=E10-E12", "F": "=F10-F12"}, tag="Verified／Derived")
qrow(14, "淨利 ÷ 營業利益", {"C": "=C15/C13", "D": "=D15/D13", "E": "=q_nioi", "F": "=q_nioi", "G": "=G15/G13"}, fmt="0.000", ann="", tag="Derived", s1="—")
qrow(15, "歸屬母公司淨利", {"B": 3.18, "C": 1.227, "D": 1.524, "E": "=E13*E14", "F": "=F13*F14"}, tag="Verified／Interested-party（FY2025）", s2="FY2025 3.18（+38.1%；富果摘要）；1H26 2.751")
qrow(16, "股數（百萬股）", {"B": "=B15/B18*1000", "C": "=C15/C18*1000", "D": "=D15/D18*1000", "E": "=D16", "F": "=D16"}, fmt="#,##0.00", ann="=AVERAGE(C16:F16)", tag="Derived", s1="淨利 ÷ 公司揭露 EPS")
qrow(17, "EPS", {c: f"={c}15/{c}16*1000" for c in "BCDEF"}, fmt="0.00", ann="=SUM(C17:F17)", tag="Derived", s1="淨利 ÷ 股數；FY2026＝四季加總")
qrow(18, "  公司揭露 EPS", {"B": 33.49, "C": 12.53, "D": 15.55}, fmt="0.00", ann="", tag="Verified", s2="1H26 28.08；4Q25 9.81")
Q.cell(20, 1, "FY2026 錨點的檢查（參考；校準未使用券商數字）").font = BOLD
checks = [("FY2026 營收年增", "=G4/B4-1", "0.0%"), ("3Q26 營收季增（指引 0–5%）", "=E4/D4-1", "0.0%"), ("FY2026 探針卡年增", "=G5/B5-1", "0.0%"), ("FY2026 設備年增", "=G8/B8-1", "0.0%"),
          ("FY2026 營收 ÷ 券商區間中點 − 1", "=G4/cons_rev26-1", "0.0%"), ("1H26 EPS 自算 − 公司 28.08（應約 0）", "=C17+D17-28.08", "0.000")]
for i, (lab, f, fm) in enumerate(checks, 21):
    Q.cell(i, 1, lab); Q[f"B{i}"] = f; Q[f"B{i}"].number_format = fm
for nm, rr in (("rev", 4), ("pc", 5), ("ai", 6), ("npc", 7), ("eq", 8), ("oth", 9), ("gp", 10), ("opx", 12), ("oi", 13), ("ni", 15), ("sh", 16), ("eps", 17)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$G${rr}")
Q.column_dimensions["A"].width = 62
for c in "BCDEFG": Q.column_dimensions[c].width = 14
Q.column_dimensions["H"].width = 26; Q.column_dimensions["I"].width = 60; Q.column_dimensions["J"].width = 60
Q.freeze_panes = "B4"

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G", 2030: "H", 2031: "I", 2032: "J"}
FADE = (2030, 2031, 2032); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
BU = (2026, 2027, 2028, 2029); F3 = (2027, 2028, 2029)
DEFP = dict(scn="A", mf="1", exnv="ex_nv", rel="rel{y}_{s}", gct="g_ct", gna="g_nonai", dnpc="d_npc", kcpo="k_cpo", cpo="cpo_{s}_{y}",
            eqk="eqk{y}_{s}", ai26="a26_ai", npc26="a26_npc", gm27="gm27_{s}", gmL="gmL_{s}", gmadj="gm_adj", phi="phi", nioi="ni_oi", fg="fade_g{s}", fm="fade_m{s}")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    p = lambda k, y=None: P[k].format(y=(y % 100 if y else ""), s=s)
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.000", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 11, note); R[key] = r; r += 1
    def sec_(t):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
    um = lambda y: {2026: "", 2027: "", 2028: f"*um28_{s}", 2029: f"*um29_{s}"}[y]
    sec_("A. AI 加速器 → 晶圓測試量 → AI／HPC 探針卡（D 欄＝FY2026 校準年）")
    row("units", "AI 加速器顆數（含情境係數）", "百萬顆", {FC[y]: f"=units_{y % 100}{um(y)}" + ("" if y == 2026 else f"*{P['mf']}") for y in BU}, "0.00")
    row("k", "每顆加權係數（測試時間 × 旺矽曝險）", "x", {FC[y]: f"={P['exnv']}*nvtt_{y % 100}+astt_{y % 100}" for y in BU}, "0.000", note="NVIDIA 部分 × ex_nv＋ASIC 與 AMD 部分（Siliconomics 測試時間係數）")
    row("w", "晶圓測試量指數（2026＝1）", "x", {FC[y]: f"={FC[y]}{R['units']}*{FC[y]}{R['k']}/($D{R['units']}*$D{R['k']})" for y in BU}, "0.000")
    row("rel", "旺矽份額與價格相對值（2026＝1）", "x", {"D": "=1", **{FC[y]: f"={p('rel', y)}" for y in F3}}, "0.00")
    row("ct", "每顆探針卡內容指數（2026＝1）", "x", {"D": "=1", **{FC[y]: f"=(1+{P['gct']})^{y - 2026}" for y in F3}}, "0.000")
    sec_("B. 營收分部（十億新台幣）")
    row("ai", "  AI／HPC 探針卡（VPC、MEMS；命題子列）", "NT$B", {"C": "=a25_ai", "D": f"={P['ai26']}", **{FC[y]: f"=$D{r}*{FC[y]}{R['w']}*{FC[y]}{R['rel']}*{FC[y]}{R['ct']}" for y in F3}},
        note="FY2027 起＝FY2026 × 晶圓測試量指數 × 份額與價格相對值 × 內容指數")
    row("npc", "  其他探針卡（CPC 等）", "NT$B", {"C": "=a25_npc", "D": f"={P['npc26']}", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['gna']}+{P['dnpc']})" for y in F3}}, note="非 AI SoC 市場年增（Siliconomics）＋超額成長")
    row("pc", "探針卡", "NT$B", {FC[y]: f"={FC[y]}{R['ai']}+{FC[y]}{R['npc']}" for y in FYS})
    row("cpoeq", "  CPO 測試設備", "NT$B", {"D": f"={P['kcpo']}*cpo_A_26/cpo_A_27", **{FC[y]: f"={P['kcpo']}*{p('cpo', y)}/cpo_A_27" for y in F3}},
        note="2027 情境 A＝k_cpo；依 Siliconomics 同字母 CPO 滲透率縮放")
    row("beq", "  設備（不含 CPO：AST、溫控、其他光電）", "NT$B", {"C": "=a25_eq", "D": f"=a26_eq-D{r - 1}", "E": f"=D{r}*semi_27",
        **{FC[y]: f"=$E{r}*{p('eqk', y)}" for y in (2028, 2029)}}, note="FY2027＝FY2026 × SEMI 測試設備成長；FY2028–29＝FY2027 × 設備情境係數（Siliconomics）")
    row("eq", "設備", "NT$B", {"C": "=a25_eq", **{FC[y]: f"={FC[y]}{R['cpoeq']}+{FC[y]}{R['beq']}" for y in BU}})
    row("oth", "其他", "NT$B", {"C": "=a25_oth", "D": "=a26_oth", **{FC[y]: f"={FC[y - 1]}{r}*(1+g_oth)" for y in F3}})
    row("rev", "營收", "NT$B", {FC[y]: f"={FC[y]}{R['pc']}+{FC[y]}{R['eq']}+{FC[y]}{R['oth']}" for y in FYS})
    row("g", "營收年增", "%", {FC[y]: f"={FC[y]}{R['rev']}/{FC[y - 1]}{R['rev']}-1" for y in FYS[1:]}, "0.0%")
    row("aish", "AI／HPC 探針卡 ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['ai']}/{FC[y]}{R['rev']}" for y in FYS}, "0.0%")
    row("pcsh", "探針卡 ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['pc']}/{FC[y]}{R['rev']}" for y in FYS}, "0.0%")
    sec_("C. 毛利與損益")
    row("gm", "毛利率", "%", {"C": "=a25_gp/a25_rev", "D": "=a26_gp/a26_rev", "E": f"={p('gm27')}+{P['gmadj']}", **{FC[y]: f"={p('gmL')}+{P['gmadj']}" for y in (2028, 2029)}}, "0.00%")
    row("gp", "毛利", "NT$B", {FC[y]: f"={FC[y]}{R['rev']}*{FC[y]}{R['gm']}" for y in FYS})
    row("opx", "營業費用", "NT$B", {"D": "=a26_opx", **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['phi']}*{FC[y]}{R['g']})" for y in F3}}, note="FY2025 營業利益未找到，自 FY2026 起")
    row("oi", "營業利益", "NT$B", {FC[y]: f"={FC[y]}{R['gp']}-{FC[y]}{R['opx']}" for y in BU})
    row("oim", "營業利益率", "%", {FC[y]: f"={FC[y]}{R['oi']}/{FC[y]}{R['rev']}" for y in BU}, "0.00%")
    row("ni", "歸屬母公司淨利", "NT$B", {"C": "=a25_ni", "D": "=a26_ni", **{FC[y]: f"={FC[y]}{R['oi']}*{P['nioi']}" for y in F3}})
    row("shr", "股數", "百萬股", {"C": "=a25_sh", "D": "=a26_sh", **{FC[y]: f"={FC[y - 1]}{r}*(1+sh_g)" for y in F3}}, "#,##0.00")
    row("eps", "EPS", "NT$", {"C": "=a25_eps", "D": "=a26_eps", **{FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['shr']}*1000" for y in F3}}, "0.00")
    row("epsg", "EPS 年增", "%", {FC[y]: f"={FC[y]}{R['eps']}/{FC[y - 1]}{R['eps']}-1" for y in FYS[1:]}, "0.0%")
    row("nim", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    for y in FADE:
        c, pc, k = FC[y], FC[y - 1], y - 2029
        ws[f"{c}{R['g']}"] = f"=$G{R['g']}+({p('fg')}-$G{R['g']})*{k}/3"
        ws[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
        ws[f"{c}{R['nim']}"] = f"=$G{R['nim']}+({p('fm')}-$G{R['nim']})*{k}/3"
        ws[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
        ws[f"{c}{R['shr']}"] = f"={pc}{R['shr']}*(1+sh_g)"
        ws[f"{c}{R['eps']}"] = f"={c}{R['ni']}/{c}{R['shr']}*1000"
        ws[f"{c}{R['epsg']}"] = f"={c}{R['eps']}/{pc}{R['eps']}-1"
        for key, fm in (("g", "0.0%"), ("rev", "#,##0.000"), ("nim", "0.00%"), ("ni", "#,##0.000"), ("shr", "#,##0.00"), ("eps", "0.00"), ("epsg", "0.0%")):
            ws[f"{c}{R[key]}"].number_format = fm; ws[f"{c}{R[key]}"].fill = FADEFILL
    ws.cell(R["rev"], 11, "FY2030 起為收斂法：營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m（FY2032 達到），非驅動因子推導")
    R["end"] = r
    return R

def dcf_expr(cells, r="tp_r", g="tp_g"):
    """一個情境的 DCF 每股價值（股東現金流＝EPS ×（1 − g／ROE）；下限 0）。cells：FY2027–FY2032 EPS 儲存格（6 個）。"""
    reinv = f"MIN({g}/tp_roe,0.8)"
    cf = lambda x: f"{x}*(1-{reinv})"
    pv = "+".join(f"{cf(x)}/(1+{r})^{t}" for t, x in enumerate(cells, 1))
    return f"MAX(0,{pv}+{cf(cells[-1])}*(1+{g})/({r}-{g})/(1+{r})^6)"

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 旺矽損益（十億新台幣；EPS 新台幣；歸屬母公司）。FY2025A 實績、FY2026 錨點（1Q、2Q26 實績＋3Q26 營收實績＋假設；三情境相同），FY2027 起為模型。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要"
PB = {}; r0 = 70
SCN_T = {"A": "情境 A：顆數基準；份額與價格相對值 1.05／1.08／1.10；毛利率 60.0%／61.0%；CPO 與設備依 Siliconomics 情境 A",
         "B": "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85；相對值 1.00；毛利率 59.5%；設備 2028 持平",
         "C": "情境 C：CY2028–29 顆數 ×0.6；相對值 0.92／0.88／0.85；毛利率 58.0%／56.5%；設備 2028 下滑兩成"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
sr = 4; SUMR = {}
SUMK = (("rev", "營收", "NT$B", "#,##0.00"), ("g", "  營收年增", "%", "0.0%"), ("pc", "  探針卡", "NT$B", "#,##0.00"), ("ai", "    AI／HPC 探針卡", "NT$B", "#,##0.00"),
        ("aish", "    AI／HPC 探針卡 ÷ 營收", "%", "0.0%"), ("eq", "  設備", "NT$B", "#,##0.00"), ("cpoeq", "    CPO 測試設備", "NT$B", "#,##0.00"),
        ("gm", "毛利率", "%", "0.00%"), ("oim", "營業利益率", "%", "0.00%"), ("ni", "歸屬母公司淨利", "NT$B", "#,##0.00"), ("nim", "  淨利率", "%", "0.00%"),
        ("eps", "EPS", "NT$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%"))
LONGK = ("rev", "g", "ni", "nim", "eps", "epsg")
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in SUMK:
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS + (FADE if key in LONGK else ()):
            if key in ("g", "epsg") and y == 2025: continue
            if key in ("oim",) and y == 2025: continue
            if key == "cpoeq" and y == 2025: continue
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
               ("  情境 A：AI＋其他探針卡 − Q 頁探針卡（NT$B）", f"D{PB['A']['ai']}+D{PB['A']['npc']}-a26_pc")):
    PL.cell(sr, 1, lab); PL[f"D{sr}"] = "=" + f; PL[f"D{sr}"].number_format = "0.0000"; CHK.append(sr); sr += 1
assert sr < 70, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 10; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 13
PL.cell(2, 1, "灰底＝FY2030–FY2032 收斂法（上游只到 2029）：營收年增線性收斂到 fade_g、淨利率線性收斂到 fade_m（In 頁 H 節），非驅動因子推導")

# ---------------- Units：AI 加速器 → 晶圓測試量 → AI 探針卡 ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — AI 加速器顆數（Siliconomics）× 每顆測試時間 → 晶圓測試量 → 旺矽 AI／HPC 探針卡（情境 A）。SoC 測試機市場（Siliconomics SIF_TestSoCMkt_A）只作 Recon"
A_ = PB["A"]; cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("units", "AI 加速器顆數（情境 A）", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['units']}", "0.00", "Siliconomics（台積電、日月光、京元電模型同一份）"),
        ("k", "每顆加權係數（測試時間 × 曝險）", "x", lambda y: f"=PL!{FC[y]}{A_['k']}", "0.000", "Siliconomics ts_tt_*；曝險為本模型 Assumed"),
        ("w", "晶圓測試量指數（2026＝1）", "x", lambda y: f"=PL!{FC[y]}{A_['w']}", "0.000", "顆數 × 係數"),
        ("rel", "份額與價格相對值", "x", lambda y: f"=PL!{FC[y]}{A_['rel']}", "0.00", "Assumed（情境 A）"),
        ("ct", "每顆探針卡內容指數", "x", lambda y: f"=PL!{FC[y]}{A_['ct']}", "0.000", "Assumed"),
        ("ai", "旺矽 AI／HPC 探針卡營收", "NT$B", lambda y: f"=PL!{FC[y]}{A_['ai']}", "0.00", "FY2026 × 指數 × 相對值 × 內容"),
        ("pu", "每顆加速器的旺矽 AI 探針卡營收", "NT$／顆", lambda y: f"=PL!{FC[y]}{A_['ai']}/PL!{FC[y]}{A_['units']}*1000", "#,##0", "十億元 ÷ 百萬顆 × 1,000"),
        ("soc", "SoC 測試機市場（Siliconomics SIF_TestSoCMkt_A）", "US$B", lambda y: SOC["A"][y - 2025], "0.00", "Siliconomics 推導值（唯讀）"),
        ("aig", "旺矽 AI 探針卡年增", "%", lambda y: f"={cyc(y)}{{ai}}/{chr(ord(cyc(y)) - 1)}{{ai}}-1" if y > 2026 else None, "0.0%", "對照下列 SoC 測試機市場年增"),
        ("socg", "SoC 測試機市場年增", "%", lambda y: f"={cyc(y)}{{soc}}/{chr(ord(cyc(y)) - 1)}{{soc}}-1" if y > 2026 else SOC["A"][1] / SOC["A"][0] - 1, "0.0%", "測試機（設備）與探針卡（耗材）同受晶圓測試量驅動")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    UR[key] = ur
    for y in BU:
        fx_ = f(y)
        if fx_ is None: continue
        if isinstance(fx_, str):
            fx_ = fx_.replace("{ai}", str(UR.get("ai", 0))).replace("{soc}", str(UR.get("soc", 0)))
        UN[f"{cyc(y)}{ur}"] = fx_; UN[f"{cyc(y)}{ur}"].number_format = fmt
        if key in ("soc",) or (key == "socg" and y == 2026): UN[f"{cyc(y)}{ur}"].font = BLUE
    UN.cell(ur, 7, note); ur += 1
ur += 1
UN.cell(ur, 1, "讀法：AI／HPC 探針卡 FY2026 以 AI 占探針卡六成（Assumed）校準，不是公司揭露數字；晶圓測試量用 Siliconomics 的每顆測試時間係數加權，NVIDIA 部分再乘旺矽曝險 ex_nv。份額與價格相對值是 Assumed，只有相對變化影響 FY2027 起的營收。").font = BOLD
UN_NOTE = ur
UN.column_dimensions["A"].width = 52; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 13
UN.column_dimensions["G"].width = 56

OBS_MPI = [
    "目標價 NT$2,641（DCF 1,995、Comparable 3,288 各半），比現價 4,610（2026-10-30）低 42.7%，級距 −25%～−50%，投資建議放空。Comparable 單獨為 −28.7%。",
    "上檔空間低於 −25% 的主因是 EPS，不是倍數：本模型 12 個月後前瞻 EPS 108.4 比 ALETHEIA 同口徑 269.8 低 59.8%（FY2027 情境 B 102.3 也比法人 128.95 低 21%），而 ALETHEIA 目標價 10,000 隱含 37.1 倍、只比同業中位數 30.3 倍高 22%；現價 ÷ FY2027 EPS 42.3／45.1／49.9 倍（A／B／C），情境 B 的 FY2028 EPS 只比 FY2027 多 3%（顆數 ×0.8）。",
    "AI／HPC 探針卡 FY2027 需較 FY2026 成長 85%／77%／63%（A／B／C；晶圓測試量指數 1.68 × 份額與價格相對值 × 內容），已高於公司「2027 產能增幅不低於 2026 約五成」的計畫，產能是上檔的瓶頸；最敏感的是 AI 部署 ±20%（加權目標價 ±12.4%）與每顆內容成長（−7.1%／+7.8%）。下一個檢驗點是 3Q26 財報與法說（約 11 月中）的毛利率是否回到 59.4%，以及 10 月營收（11-10 前）。",
]
# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，年底落後 P/E）只作參考"
V["A2"] = "循環位置：旺矽以探針卡（耗材）為主、設備約占四分之一，不套設備市場見頂年；比照京元電，以本模型營收最高的會計年度為見頂年（FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
O = 51
V.cell(O, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, O + 1, ["項目", "A", "B", "C", "說明"])
V.cell(O + 2, 1, "現價（2026-10-30 收盤）"); V[f"B{O + 2}"] = "=px"; V[f"B{O + 2}"].number_format = "0.00"
for k, y, lab in ((3, 2026, "現價 ÷ FY2026 EPS（錨點）"), (4, 2027, "現價 ÷ FY2027 EPS（模型）"), (5, 2028, "現價 ÷ FY2028 EPS（模型）"), (6, 2029, "現價 ÷ FY2029 EPS（模型）")):
    V.cell(O + k, 1, lab)
    for j, s in enumerate(SC):
        c = "BCD"[j]; V[f"{c}{O + k}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{O + k}"].number_format = "0.0"
V.cell(O + 7, 1, "未來 12 月 EPS（模型 B）≈ 0.2 × FY2026＋0.8 × FY2027"); V[f"B{O + 7}"] = f"=0.2*PL!D{SUMR[('B', 'eps')]}+0.8*PL!E{SUMR[('B', 'eps')]}"; V[f"B{O + 7}"].number_format = "0.00"
V.cell(O + 8, 1, "  現價 ÷ 未來 12 月 EPS（模型 B）"); V[f"B{O + 8}"] = f"=px/B{O + 7}"; V[f"B{O + 8}"].number_format = "0.0"
V.cell(O + 9, 1, "近 12 月 P/E（玩股網，10-30）"); V[f"B{O + 9}"] = "=pe_ttm"; V[f"B{O + 9}"].number_format = "0.0"
V.cell(O + 10, 1, "前瞻 P/E（現價 ÷ 法人 FY2027 EPS）"); V[f"B{O + 10}"] = "=fpe"; V[f"B{O + 10}"].number_format = "0.0"
V[f"E{O + 3}"] = "FY2026 三情境相同（錨點）"; V[f"E{O + 4}"] = "FY2027 三情境只差份額與價格相對值、毛利率與 CPO 滲透率"
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
    V.cell(i, 6, "tools/cycle_position.py excel_rule；本模型營收最高年取代設備市場見頂年（比照京元電）" if s == "A" else "同上")
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
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數低於歷史中位數時，現價相對便宜；倍數是否合理，另看前瞻 P/E 歷史（Bloomberg，Val 第 0.5 節）。本版未做先驗，無 Andy 合理 P/E 對照。")
WMT_NOTE = vr
PE_MT = [None, None, None, None, None, "=116.5/6.38", "=112.5/12.67", "=217.5/13.55", "=926/19.76", "=2250/31.01"]
T10 = vr + 2
V.cell(T10, 1, "5. 十年 P/E（2021–2025 年底；2016–2020 未找到；單一來源，分位排名用；無負值或 >200 需剔除）").font = BOLD
hdr(V, T10 + 1, ["年底", "採用 P/E", "玩股網（12 月收盤 ÷ 近四季 EPS）", "（第二來源：未找到）"])
RB0 = T10 + 16
for i, y in enumerate(range(2016, 2026)):
    rr = T10 + 2 + i
    V.cell(rr, 1, str(y))
    V.cell(rr, 2, f'=IF(ISNUMBER(C{rr}),C{rr},"")').number_format = "0.00"
    V.cell(rr, 3, f'=IF(ISNUMBER(B{RB0 + 5 + i}),B{RB0 + 5 + i},"")').number_format = "0.00"
PE10 = f"$B${T10 + 2}:$B${T10 + 11}"
V.cell(T10 + 12, 1, "年數"); V[f"B{T10 + 12}"] = f"=COUNT({PE10})"
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "玩股網（12 月收盤 ÷ 近四季 EPS）", "vals": PE_MT, "tag": "Interested-party（資料商；2021–2025 共 5 年，本模型以頁面 12 月收盤與近四季 EPS 換算；2016–2020 未找到）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位。旺矽只取得 2021–2025 五年，且 2024–2025 年為 AI 重新評價期（47x、73x），分位偏高；玩股網的近四季 EPS 與公司年度 EPS 口徑不同（2025 31.01 vs 33.49）。")
assert V.cell(RB0 + 5, 1).value == "2016" and V.cell(RB0 + 14, 1).value == "2025"
rule_rows = [rw for rw in range(RB0, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")

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
for i, y in enumerate(YRS6):
    tset(16 + i, f"cf{y}", f"  FY{y}", [f"={c}{EPSR[y]}*(1-$B$15)" for c in "BCD"], "0.00")
tset(22, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=MAX(0,SUMPRODUCT({c}16:{c}21,1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}21*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6)" for c in "BCD"], "0.00",
     "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底；下限 0")
tset(23, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}22*(1+tp_r)" for c in "BCD"], "0.00")
tset(24, "dcftp", "DCF 目標價（情境加權）", ["=tp_wA*B23+tp_wB*C23+tp_wC*D23", None, None], "0.00", "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁")
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = [("FormFactor", "FORM", 28.17, "2026-10-23"), ("Advantest", "6857.T", 37.69, "2026-10-21"), ("泰瑞達（Teradyne）", "TER", 32.52, "2026-10-23"), ("京元電子（KYEC）", "2449.TW", 26.13, "2026-10-06")]
for i, (nm, tk, pe, dt) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, "stockanalysis.com forward PE（分析師共識；價格日期如左；與泰瑞達、Amkor 模型同一組數字）"); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", "測試介面與測試同業：FormFactor（探針卡）、Advantest、泰瑞達（測試機）、京元電（測試）；Technoprobe、精測不在 Bloomberg 檔，未納入")
tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）", [f"=tp_wA*(2/12*B{EPSR[2027]}+10/12*B{EPSR[2028]})+tp_wB*(2/12*C{EPSR[2027]}+10/12*C{EPSR[2028]})+tp_wC*(2/12*D{EPSR[2027]}+10/12*D{EPSR[2028]})", None, None], "0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], "0.00", "同業中位數 × 前瞻 EPS")
V.cell(37, 1, "0.3 結論").font = BOLD
hdr(V, 38, ["項目", "數值", "", "", "說明"])
tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B24+(1-tp_w_dcf)*B35", None, None], "0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable（相對倍數調整後，Val 第 0.5 節）；tp_w_dcf 見 In 頁")
tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], "0.00")
tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], "0.00", "賣方平均：未找到；以 ALETHEIA 10,000（2026-06-13，單一券商）代替")
tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", "現價＝In 頁 px（2026-10-30 收盤）")
tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", "本模型 12 個月後前瞻 EPS ÷ ALETHEIA 同口徑前瞻 EPS − 1（In 頁 cons_feps）")
tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
tset(46, "band", "價格帶（DCF 不如預期～順利）", ['=TEXT(D23,"0.00")&"～"&TEXT(B23,"0.00")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=IF(px/{c}{EPSR[2028]}>=MAX({PE10}),1,IF(px/{c}{EPSR[2028]}<=MIN({PE10}),0,PERCENTRANK({PE10},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 5 節）的分位；只作檢查，不進目標價")
V.cell(48, 1, "目標價方法")
V["B48"] = ("DCF：股東現金流＝EPS ×（1 − g／ROE），r＝10%、g＝4%、ROE 為 FY2024–25 平均約 26%（再投資率約 15%），FY2027–32 明確期＋永續；旺矽資本支出小（2026 年約 27 億元），本式適用。"
            "Comparable：測試介面與測試同業（FormFactor、Advantest、泰瑞達、京元電）前瞻 P/E 中位數；旺矽本身沒有 Bloomberg 前瞻倍數，相對倍數調整係數＝1（不調整）。"
            "情境對應：A＝順利（顆數基準）、B＝平平（顆數 ×0.8／0.85）、C＝不如預期（顆數 ×0.6）；設備分部讀 Siliconomics 設備情境係數")
for r_ in (24, 35, 39):
    V[f"B{r_}"].fill = KEY
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 16
GR_ROW, RC_ROW = 49, 50
V.cell(GR_ROW, 1, "上檔空間級距")
V.cell(GR_ROW, 2, '=IF(B42>0.5,"大於 50%",IF(B42>0.25,"25%–50%",IF(B42>=0,"0%–25%",IF(B42>=-0.25,"0%～−25%",IF(B42>=-0.5,"−25%～−50%","小於 −50%")))))').number_format = "@"
V.cell(GR_ROW, 5, "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）")
V.cell(RC_ROW, 1, "投資建議")
V.cell(RC_ROW, 2, '=IF(B42>0.5,"強力買進",IF(B42>0.25,"買進",IF(B42>=-0.25,"中立",IF(B42>=-0.5,"放空","強力放空"))))').number_format = "@"
V.cell(RC_ROW, 5, "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")

# ---------------- Sens（每案三情境重算，對加權目標價；鍵欄） ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（三情境同時重算，再依情境權重求 DCF 與 Comparable 目標價）：對 FY2027、FY2028 EPS（情境 B）與加權目標價的影響。每個案例的三情境推導區塊在本頁下方"
SE["A2"] = "註：情境專屬列（份額與價格相對值、毛利率、CPO 滲透率、設備情境係數、顆數係數）不給敏感度區間，差異見情境 B、C。ai_sh26（AI 占探針卡）改的是 FY2026 錨點拆分，案例區塊以 FY2026 探針卡 × 低／高占比重算拆分（錨點合計不變）；Comparable 欄已乘相對倍數調整係數（rm_factor）。"
hdr(SE, 3, ["案例", "鍵", "設定", "FY2027 EPS（B）", "FY2028 EPS（B）", "FY2028 EPS（B）相對基準", "DCF 目標價", "Comparable 目標價", "加權目標價", "加權目標價相對基準"])
CASES = [("AI 部署 ×0.8", "mf", "FY2027 起顆數 ×0.8（三情境）", dict(mf="0.8")), ("AI 部署 ×1.2", "mf", "×1.2", dict(mf="1.2")),
         ("NVIDIA 曝險 0.2", "ex_nv", "基準 0.4（ASIC＝1）", dict(exnv="ex_nv_lo")), ("NVIDIA 曝險 0.7", "ex_nv", "", dict(exnv="ex_nv_hi")),
         ("AI 占探針卡 45%", "ai_sh26", "FY2026 校準；基準 60%", dict(ai26="a26_pc*ai_sh26_lo", npc26="a26_pc*(1-ai_sh26_lo)")), ("AI 占探針卡 70%", "ai_sh26", "", dict(ai26="a26_pc*ai_sh26_hi", npc26="a26_pc*(1-ai_sh26_hi)")),
         ("每顆內容年增 0%", "g_ct", "基準 5%", dict(gct="g_ct_lo")), ("每顆內容年增 10%", "g_ct", "", dict(gct="g_ct_hi")),
         ("非 AI SoC 年增 0%", "g_nonai", "Siliconomics ts_g_nonai 低", dict(gna="g_nonai_lo")), ("非 AI SoC 年增 8%", "g_nonai", "Siliconomics 高", dict(gna="g_nonai_hi")),
         ("其他探針卡超額成長 0%", "d_npc", "基準 3%", dict(dnpc="d_npc_lo")), ("其他探針卡超額成長 6%", "d_npc", "", dict(dnpc="d_npc_hi")),
         ("CPO 測試設備 0.5", "k_cpo", "FY2027 情境 A；基準 1.2", dict(kcpo="k_cpo_lo")), ("CPO 測試設備 2.0", "k_cpo", "", dict(kcpo="k_cpo_hi")),
         ("毛利率 −2pt", "gm_adj", "三情境、FY2027 起", dict(gmadj="gm_adj_lo")), ("毛利率 +2pt", "gm_adj", "", dict(gmadj="gm_adj_hi")),
         ("營業費用彈性 0.4", "phi", "基準 0.6", dict(phi="phi_lo")), ("營業費用彈性 0.9", "phi", "", dict(phi="phi_hi")),
         ("淨利 ÷ 營業利益 0.80", "ni_oi", "基準 0.85", dict(nioi="ni_oi_lo")), ("淨利 ÷ 營業利益 0.90", "ni_oi", "", dict(nioi="ni_oi_hi"))]
BASE_ENG = 120
SB = {}; brow = BASE_ENG
SE.cell(BASE_ENG - 2, 1, "敏感度「逐案推導區塊」（以下每案三個情境各一區塊；不是案例列）").font = BOLD
for nm, k, st, ov in CASES:
    SB[nm] = {}
    for s in SC:
        SB[nm][s] = engine(SE, brow, f"案例區塊：{nm}（情境 {s}）", scn=s, **ov)
        brow = SB[nm][s]["end"] + 2
YC = [FC[y] for y in YRS6]
def put_case(i, name, key, setting, blocks, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, key); SE.cell(i, 3, setting)
    SE[f"D{i}"] = f"={pre}E{blocks['B']['eps']}"; SE[f"E{i}"] = f"={pre}F{blocks['B']['eps']}"
    SE[f"F{i}"] = f"=E{i}/E$4-1"
    SE[f"G{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_expr([pre + c + str(blocks[s]['eps']) for c in YC])}" for s in SC) + ")*(1+tp_r)"
    feps = "+".join(f"tp_w{s}*(2/12*{pre}E{blocks[s]['eps']}+10/12*{pre}F{blocks[s]['eps']})" for s in SC)
    SE[f"H{i}"] = f"=Val!$B${T['pmed']}*({feps})*rm_factor"
    SE[f"I{i}"] = f"=tp_w_dcf*G{i}+(1-tp_w_dcf)*H{i}"
    SE[f"J{i}"] = f"=I{i}/I$4-1"
    for c, fm in (("D", "0.00"), ("E", "0.00"), ("F", "0.0%"), ("G", "#,##0.0"), ("H", "#,##0.0"), ("I", "#,##0.0"), ("J", "0.0%")):
        SE[f"{c}{i}"].number_format = fm
put_case(4, "基準", "—", "PL 三情境", PB, "PL!")
for n, (nm, k, st, ov) in enumerate(CASES):
    put_case(5 + n, nm, k, st, SB[nm], "")
nn = 5 + len(CASES)
TPS0 = nn
def tp_case(i, name, key, setting, r="tp_r", g="tp_g", w="tp_w_dcf"):
    SE.cell(i, 1, name); SE.cell(i, 2, key); SE.cell(i, 3, setting)
    SE[f"D{i}"] = "=D4"; SE[f"E{i}"] = "=E4"; SE[f"F{i}"] = "=E{0}/E$4-1".format(i)
    SE[f"G{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_expr(['Val!' + c + str(EPSR[y]) for y in YRS6], r, g)}" for s, c in zip(SC, "BCD")) + f")*(1+{r})"
    SE[f"H{i}"] = "=H4"
    SE[f"I{i}"] = f"={w}*G{i}+(1-{w})*H{i}"
    SE[f"J{i}"] = f"=I{i}/I$4-1"
    for c, fm in (("D", "0.00"), ("E", "0.00"), ("F", "0.0%"), ("G", "#,##0.0"), ("H", "#,##0.0"), ("I", "#,##0.0"), ("J", "0.0%")):
        SE[f"{c}{i}"].number_format = fm
TPCASES = [("折現率 +1pt", "tp_r", "r 11%", dict(r="(tp_r+0.01)")), ("折現率 −1pt", "tp_r", "r 9%", dict(r="(tp_r-0.01)")),
           ("長期成長 +1pt", "tp_g", "g 5%", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "tp_g", "g 3%", dict(g="(tp_g-0.01)")),
           ("DCF 權重 0.3", "tp_w_dcf", "Comparable 0.7", dict(w="0.3")), ("DCF 權重 0.7", "tp_w_dcf", "Comparable 0.3", dict(w="0.7"))]
for k_, (nm, key, st, kw) in enumerate(TPCASES):
    tp_case(TPS0 + k_, nm, key, st, **kw)
TPS_LAST = TPS0 + len(TPCASES) - 1
put_case(TPS_LAST + 1, "情境 B 全權重", "—", "只看情境 B（對照）", {"A": PB["B"], "B": PB["B"], "C": PB["B"]}, "PL!")
put_case(TPS_LAST + 2, "情境 C 全權重", "—", "只看情境 C（對照）", {"A": PB["C"], "B": PB["C"], "C": PB["C"]}, "PL!")
SENS_LAST = TPS_LAST + 2
SE.cell(SENS_LAST + 2, 1, "讀法：顆數情境 B、C 的差異從 FY2028 起；FY2027 只差份額與價格相對值、毛利率與 CPO。AI／HPC 探針卡約占營收四成五，晶圓測試量指數（顆數 × 測試時間）與毛利率是 EPS 最敏感的兩項；NVIDIA 曝險越低，Rubin 世代測試時間加長的好處越少。")
SENS_NOTE = SENS_LAST + 2
assert SENS_NOTE < BASE_ENG - 3, SENS_NOTE
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 12; SE.column_dimensions["C"].width = 34
for c in "DEFGHIJ": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司指引、券商預估與市場口徑。只作事後對照，不為對上數字而調參數。賣方共識未找到，以法人（財訊 2026-10-12 轉述）與 ALETHEIA（單一券商）對照"
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
RC.cell(rr_, 1, "A. FY2026 錨點 vs 公司指引與券商").font = BOLD; rr_ += 1
rec("3Q26 營收季增 vs 公司指引中點 2.5%（百分點）", 0.025, lambda s: "=Q!E4/Q!D4-1", "2Q26 法說（口袋學堂）：0–5%", "5 月法說曾為 +10–15%", fmt="0.0%", pt=True)
rec("FY2026 營收 vs 券商區間中點", "=cons_rev26", pl("rev", 2026), "券商 20.38–23.80（8/17）", "", fmt="#,##0.00")
rec("FY2026 EPS vs 法人", "=cons_eps26", pl("eps", 2026), "財訊 2026-10-12", "口袋學堂：券商 60.89–70.84")
rec("FY2027 EPS vs 法人", "=cons_eps27", pl("eps", 2027), "財訊 2026-10-12", "")
rec("FY2027 EPS vs ALETHEIA", "=al_eps27", pl("eps", 2027), "數位時代 2026-06-13", "看多券商；差距超過 30% 時 README 說明")
rec("FY2028 EPS vs ALETHEIA", "=al_eps28", pl("eps", 2028), "同上", "")
RC.cell(rr_, 1, "B. 晶圓測試量與 SoC 測試機市場（Siliconomics）").font = BOLD; rr_ += 1
for y in (2027, 2028, 2029):
    soc = SOC
    rec(f"FY{y} AI 探針卡年增 vs SoC 測試機市場年增（同字母情境，百分點）", None, lambda s, y=y: f"=PL!{FC[y]}{PB[s]['ai']}/PL!{FC[y - 1]}{PB[s]['ai']}-1-({SOC[s][y - 2025]}/{SOC[s][y - 2026]}-1)",
        "Siliconomics SIF_TestSoCMkt_A／B／C", "正值＝探針卡成長快於測試機市場", fmt="0.0%", gap=False)
rec("FY2026 每顆加速器的旺矽 AI 探針卡營收（NT$）", None, lambda s: f"=Units!C{UR['pu']}", "—", "對照：京元電每顆 Blackwell 等值測試營收 US$46.9（京元電模型）", fmt="#,##0")
RC.cell(rr_, 1, "C. 產能").font = BOLD; rr_ += 1
rec("FY2027 AI 探針卡 ÷ FY2026（公司：2027 產能增幅不低於 2026 的約五成）", 1.5, lambda s: f"=PL!E{PB[s]['ai']}/PL!D{PB[s]['ai']}", "2026-05 法說（富果）：VPC 月產能 200 萬針、MEMS 350 萬針（2026 年底）", "比值高於 1.5 表示需要比公司計畫更多的產能或漲價", fmt="0.00")
RC.column_dimensions["A"].width = 72
for c in "BCDEFGH": RC.column_dimensions[c].width = 11
RC.column_dimensions["I"].width = 40; RC.column_dimensions["J"].width = 60
RECON_LAST = rr_ - 1

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司月營收與法說為公司原文或摘要；券商、媒體與資料商皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("公司（摘要）：2026-05 法說（富果；1Q26 損益、事業別、產能、資本支出、季度指引）", FG1),
        ("公司（摘要）：2026-03 法說（富果；FY2025 營收、EPS、事業別、每股淨值）", FG4),
        ("公司（摘要）：2Q26 財報與法說（口袋學堂 2026-09-08；2Q26 損益、事業別、3Q、4Q26 指引、券商區間）", PK2),
        ("公司：9 月營收（科技新報 2026-10-08；前三季 150.68 億元）", TN9),
        ("媒體：財訊 2026-10-12（客戶、份額傳聞、產能、高盛毛利率、法人 EPS）", MW),
        ("媒體：經濟日報 2026-08-20（CPO 測試設備、Insertion 2／3、2027 展望）", UDN8),
        ("媒體：數位時代 2026-06-13（ALETHEIA 目標價與 EPS、CPO 探針卡市場、擴產倍數）", BN),
        ("價格：富邦證券 e 點通 6223 日線（10-30 收盤 4,610.00）", FUBON),
        ("資料商：玩股網本益比河流圖（2021–2025 年 12 月收盤與近四季 EPS、近 12 月 P/E）", WG),
        (f"行業層：{SILV}（加速器顆數、情境係數、每顆測試時間係數、非 AI SoC 成長、SEMI 測試設備、設備情境係數、CPO 滲透率、SoC 測試機市場）", "siliconomics/data/" + SIL.FILE),
        ("待補（人工）：3Q26 財報（約 11 月中）；FY2025 營業利益與年報營運部門；AI 營收占比；VPC／MEMS 份額；賣方共識與平均目標價；Bloomberg 前瞻倍數；2016–2020 年底 P/E", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 100; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
OBS = OBS_MPI
lines = [(f"旺矽（6223.TWO；MPI）個股模型 {VER}（AI 半導體 Project：測試介面）　{DATE}", True),
 (f"版本：{VER}（{DATE}：初版）。探針卡（懸臂 CPC、垂直 VPC、MEMS）約占營收七成五，設備（AST 先進半導體測試、溫控、光電／CPO 測試）約兩成五。樣板 kyec/；行業變數讀 {SILV}。", False),
 ("模型命題：AI 加速器（GPU、ASIC、HBM 相關）用的垂直與 MEMS 探針卡需求＝晶圓測試量（Siliconomics 加速器顆數 × 每顆測試時間係數，NVIDIA 部分再乘旺矽曝險）× 每顆探針卡內容 × 旺矽份額與價格相對值；其他探針卡依 Siliconomics 非 AI SoC 成長；設備分部依 Siliconomics SEMI 測試設備路徑與設備情境係數，另加 CPO 測試設備（Siliconomics 同字母 CPO 滲透率）。", False),
 ("核心問題：AI 探針卡的量（晶圓測試量）與價（MEMS 溢價、份額）能否支撐約 36 倍法人 FY2027 本益比；Technoprobe、FormFactor 擴產後份額與價格是否下滑。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)。預測期 FY2027–FY2029（三年）由驅動因子推導，FY2030–32 為收斂法。金額：十億新台幣；EPS：新台幣。口徑：歸屬母公司淨利。現價 4,610.00（2026-10-30 收盤）；其他模型仍用 10-22／10-23 收盤。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("AI／HPC 探針卡：FY2026 ＝ 探針卡 × 60%（Assumed）；之後 × 晶圓測試量指數（顆數 × 測試時間 × 曝險，Siliconomics）× 份額與價格相對值（Assumed，情境專屬）×（1＋每顆內容年增 5%）^年數。其他探針卡：非 AI SoC 年增 4%（Siliconomics）＋超額 3%。", False),
 ("設備：不含 CPO 的設備 FY2027 隨 SEMI 測試設備成長（Siliconomics），FY2028–29 依設備情境係數（A＝SEMI 路徑、B 持平、C 下滑兩成）；CPO 測試設備 FY2027 情境 A 為 12 億元（Assumed），依 Siliconomics 同字母 CPO 滲透率縮放。", False),
 ("毛利率：情境專屬（A 60%／61%、B 59.5%、C 58%／56.5%）；營業費用彈性 0.6；淨利 ÷ 營業利益 0.85；股數持平。", False),
 ("校準：FY2026 ＝ 1Q、2Q26 實績＋3Q26 營收實績（月營收）＋3Q26 毛利率 59.4%（公司說法）＋4Q26 營收季增 7.5%（指引 5–10%）。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，三情境權重 0.3／0.5／0.2）＋0.5 × Comparable 目標價（FormFactor、Advantest、泰瑞達、京元電前瞻 P/E 中位數 × 12 個月後前瞻 EPS；旺矽 Bloomberg 前瞻倍數待取，相對倍數係數＝1）。十年 P/E 規則值保留為參考。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m，FY2032 達到；PL 頁灰底）。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司數字取自月營收與法說摘要；券商、媒體與資料商為 Interested-party。", False),
 ("", False),
 ("已知限制：(1) AI 營收、VPC／MEMS 拆分與份額公司未揭露，AI 占探針卡六成為 Assumed；(2) 3Q26 損益尚未公布（約 11 月中），3Q26 只有營收是實績；(3) FY2025 營業利益未找到；(4) 產能（針數）未建模，情境 A 的成長可能超過公司擴產計畫；(5) CPO 測試設備金額為 Assumed；(6) 賣方共識未找到，以法人與 ALETHEIA 對照；(7) 十年 P/E 只有 2021–2025 五年、單一來源；(8) Bloomberg 前瞻倍數待取。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170
_n0 = RD.max_row
RD.cell(_n0 + 2, 1, f'="投資建議："&Val!B{RC_ROW}&"（依上檔空間固定規則）"').font = BOLD
for _i, _t in enumerate(OBS, 1): RD.cell(_n0 + 2 + _i, 1, f"重點觀察 {_i}：" + _t)
OBS_ROW0 = _n0 + 3

# ---------------- 分部原則第 3 條：核對列 ----------------
sys.path.insert(0, os.path.join(ROOT, "tools", "segment_check"))
from addrow import add_check
CHK_SEG = add_check(Q, seg_rows=[5, 8, 9], total_row=4, cols="BCDEFG", rnd=6, note="分部原則第 3 條：事業別＝探針卡、設備、其他（公司法說揭露占比 × 營收）；分部宣告見 segments.json")
Q.cell(CHK_SEG + 2, 1, "檢查：AI＋其他探針卡 − 探針卡（命題子列合計）").font = BOLD
for c in "BCDEFG":
    Q[f"{c}{CHK_SEG + 2}"] = f"=ROUND({c}6+{c}7-{c}5,6)"; Q[f"{c}{CHK_SEG + 2}"].number_format = "0.000"

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "mpi")
for _row in RD.iter_rows():
    for _c in _row:
        if isinstance(_c.value, str) and _c.value.startswith("2026-10-29：加權目標價中的 Comparable"):
            _c.value = ("相對倍數調整（Val 第 0.5 節）：旺矽本身沒有 Bloomberg 前瞻倍數（tools/data/fwd_mult_hist.csv 無 6223），調整係數自動取 1.0（不調整）；Bloomberg 前瞻倍數待取。"
                        "重點觀察 1 的第一句為公式。")
sys.path.insert(0, os.path.join(ROOT, "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, VER, DATE)
wb.save(sys.argv[1])
if os.environ.get("MPI_ROWMAP"):
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW, "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE,
               "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE, "UN_NOTE": UN_NOTE, "T": T, "T10": T10, "TPS0": TPS0, "PEER_LAST": PEER_LAST,
               "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0, "OBS_ROW0": OBS_ROW0, "CHK_SEG": CHK_SEG, "README_N": len(lines)},
              open(os.environ["MPI_ROWMAP"], "w"), ensure_ascii=False, indent=1)

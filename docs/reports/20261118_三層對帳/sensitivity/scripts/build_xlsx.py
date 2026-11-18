# 產生 20261118_2027差距_單因子敏感度.xlsx：In（藍字輸入，附等級與來源）→ Calc（每欄一個情境，全為公式）→ Tornado／Joint／Must／Check
import sys, json
sys.path.insert(0, "/home/claude/out/sens/scripts")
import replica as R
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL
RES = json.load(open("/home/claude/out/sens/scripts/results.json"))
UNO = json.load(open("/home/claude/out/sens/scripts/uno_results.json"))
D, I = R.D, R.I
OUT = "/home/claude/out/sens/20261118_2027差距_單因子敏感度.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); KEY = PatternFill("solid", fgColor="FFF2CC")
wb = Workbook(); IN = wb.active; IN.title = "In"
IN["A1"] = ("In — 輸入（藍字＝輸入；黑字＝由輸入推得的公式）。基準＝兩層現行值（Siliconomics v0.20、Wattonomics v0.3.6、協調橋接 v0.3.6、Tokenomics v22.0 快照）；"
            "低／高＝本報告的單因子端點，依據寫在說明欄。不含 NVIDIA FY2027 第三季實績。")
IN.append([]); IN.append(["名稱", "項目", "單位", "低", "基準", "高", "等級", "來源／端點依據"])
for c in range(1, 9): IN.cell(3, c).font = BOLD
NAMES = {}
def nm(key, row, col):
    ref = f"In!${CL(col)}${row}"
    wb.defined_names[key] = DefinedName(key, attr_text=ref); NAMES[key] = ref
def add(key, label, unit, base, lo=None, hi=None, grade="", src="", formula=False):
    r = IN.max_row + 1
    IN.cell(r, 1, key); IN.cell(r, 2, label); IN.cell(r, 3, unit)
    for c, v in ((4, lo), (5, base), (6, hi)):
        if v is None: continue
        cell = IN.cell(r, c, v)
        if not (isinstance(v, str) and v.startswith("=")): cell.font = BLUE
    IN.cell(r, 7, grade); IN.cell(r, 8, src)
    nm(key, r, 5)
    if lo is not None: nm(key + "_lo", r, 4)
    if hi is not None: nm(key + "_hi", r, 6)
    return r
def sec(t):
    r = IN.max_row + 2; IN.cell(r, 1, t).font = BOLD
TK = I["tk"]; GENS = ["H100", "GB200", "GB300", "VR200", "RU"]
SRC_TK = "Tokenomics v22.0 Interface（經 Siliconomics／Wattonomics TK_In 鏡像，兩層逐值相同）"
sec("A. Tokenomics（兩層共用）")
for g in GENS:
    add(f"i_CIT_{g}", f"每 GW IT 資本支出 {g}", "US$B/GW", TK["CapexIT"][g][1], TK["CapexIT"][g][0], TK["CapexIT"][g][2], "Derived（機架價 Interested-party）", SRC_TK + "；低／高＝成本情境")
for g in GENS:
    add(f"i_FAC_{g}", f"每 GW 廠房資本支出 {g}", "US$B/GW", TK["CapexFacility"][g][1], TK["CapexFacility"][g][0], TK["CapexFacility"][g][2], "Derived", SRC_TK)
for g in GENS:
    add(f"i_GPG_{g}", f"每 GW 封裝數 {g}", "顆/GW", TK["GPUsPerGW"][g][1], None, None, "Derived", SRC_TK + "（IT 關鍵電力、機架設計功率情境 2、機架外 8%）")
KW = {"H100": (38, 41, 44), "GB200": (120, 132, 140), "GB300": (132, 136, 140), "VR200": (188, 227, 230), "RU": (300, 340, 454)}
for g in GENS:
    add(f"i_KW_{g}", f"機架設計功率 {g}", "kW/架", KW[g][1], KW[g][0], KW[g][2], "Interested-party", "Tokenomics v22.0 Spec_Rack 第 8、10 列（analysis-R 第五節功率低／高端）")
add("i_offrack", "機架外 IT 開銷占比", "比", 0.08, 0.05, 0.14, "Assumed／Analogy", "基準＝Tokenomics；低 5% Assumed；高 14%＝Epoch IT 開銷 1.14（peers 報告；兩者定義是否相同未驗證）")
add("i_pue", "PUE（只進設施電力，不進任何一層的 GW）", "x", 1.2, 1.1, 1.3, "Assumed", "Tokenomics IF_FacilityGW；兩層 GW 都是 IT 關鍵電力（analysis-R 第五節）")
for h in ["AMD", "TPU", "Trn", "Maia", "Ascend", "OAI"]:
    add(f"i_HOLD_{h}", f"非 NVIDIA 持有成本比 {h}（每 GW 持有成本 ÷ VR200）", "比", TK["hold"][h][1], TK["hold"][h][0], TK["hold"][h][2], "Assumed", "Tokenomics v22.0 IF_NonNV_Hold_*")
sec("B. Siliconomics 2027 供給（原始顆數＝截頂前，百萬封裝）")
RAW = {"bw": (2.4, "NVIDIA Blackwell", "AC Research（原文未取得）", "Interested-party"), "rb": (7.0, "NVIDIA Rubin", "Morgan Stanley 2026-07「Rubin 及 Rubin Ultra 接近 700 萬顆」（二手轉述）", "Interested-party"),
       "ru": (0.3, "NVIDIA Rubin Ultra", "時程假設", "Assumed"), "hp": (0.0, "NVIDIA Hopper", "TrendForce", "Interested-party"),
       "amd": (1.5, "AMD Instinct", "MS MI455X 1.0M＋MI450 0.5M", "Interested-party／Assumed"), "tpu": (5.5, "Google TPU", "MS 5.0、GF 8.84 取中間偏 MS", "Interested-party"),
       "trn": (2.5, "AWS Trainium", "年出貨未找到", "Assumed"), "oth": (4.5, "其他 ASIC", "JPMorgan ASIC 12.5M − TPU − Trainium", "Derived／Assumed")}
LOHI = {"amd": (1.2, 1.8), "tpu": (5.0, 8.84), "trn": (2.0, 3.5), "oth": (3.0, 6.0)}
for k, (v, lab, src, gr) in RAW.items():
    lo, hi = LOHI.get(k, (None, None))
    add(f"i_raw_{k}", f"{lab} 2027 原始顆數", "百萬封裝", v, lo, hi, gr, f"Siliconomics v0.20 In u_{k}_27 ÷ 截頂係數；{src}" + ("；低／高＝In 低／高欄（TPU 用 MS 5.0／GF 8.84）" if lo else ""))
add("i_cap", "HBM 截頂係數（v0.18 S48）", "x", 2.256909 / 2.4, None, None, "Derived", "Siliconomics v0.20 In u_bw_27 ÷ 2.4（＝0.94038）")
HG = {"bw": 280, "rb": 288, "ru": 1024, "hp": 141, "amd": 360, "tpu": 250, "trn": 144, "oth": 200}
for k, v in HG.items(): add(f"i_hbmgb_{k}", f"每封裝 HBM {k}", "GB", v, None, None, "Interested-party／Assumed", "Siliconomics v0.20 In g_*_27")
add("i_hbm_gross", "HBM 出貨 ÷ 裝載", "x", 1.15, 1.0, 1.3, "Assumed", "Siliconomics In hbm_gross")
add("i_hbm_now", "2027 HBM 位元供給（現行，本庫自算＝截頂後需求）", "EB", "=i_hbm_gross*(" + "+".join(f"i_raw_{k}*i_cap*i_hbmgb_{k}" for k in HG) + ")/1000", None, None, "Derived", "Siliconomics v0.20 Cap 頁 6.67 EB")
add("i_hbm_ind", "2027 HBM 位元供給（獨立估計：晶圓投片 × 每片位元 ÷ 換算比）", "EB", 3.82, 2.06, 6.97, "Analogy", "docs/reports/20261118_三層對帳/hbm/20261118_HBM獨立供給_查證.md（媒體晶圓估計；報告建議不入模型）")
IT = {"amd": (100, 80, 120), "tpu": (25, 15, 35), "trn": (20, 12, 30), "oth": (20, 12, 30)}
for k, (b, lo, hi) in IT.items(): add(f"i_it_{k}", f"每顆 IT 資本支出 {k}", "千美元/顆", b, lo, hi, "Assumed", "Siliconomics v0.20 In it_AMDR／it_TPU／it_TRN／it_ASIC")
RC = {"amd": 1.1819181351, "tpu": 1.2415106461, "trn": 1.6100380200, "oth": 1.5875444934}
for k, v in RC.items(): add(f"i_rconv_{k}", f"非 NVIDIA 每 GW 換算修正 {k}（本庫隱含每 GW 顆數 ÷ Tokenomics 證據每 GW 顆數）", "x", 1.0, None, v, "Derived（功率 Interested-party）", "analysis-R 第五節非 NVIDIA 表；高欄＝改用證據值；Maia 無證據，借用『其他 ASIC』")
sec("C. Wattonomics 參數")
add("i_lag_on", "付款（出貨）→上線延遲", "月", 4.5, 3, 6, "Assumed", "Wattonomics In 第 94 列；橋接同參數")
add("i_inv_on", "未上線庫存 ÷ 年出貨", "比", 0.1, 0.05, 0.2, "Assumed", "Wattonomics In 第 95 列")
add("i_lag_bldg", "廠房資本支出提前入帳", "月", 18, 12, 24, "Assumed", "Wattonomics In 第 10 列")
add("i_oth_gmult", "「其他」年增率 ÷ 市場年增率", "x", 1, 0, 1.5, "Analogy", "Wattonomics In 第 96 列")
add("i_ai_sh", "資本支出 AI 占比（未覆寫者）", "比", 0.85, 0.75, 0.95, "Assumed", "Wattonomics In 第 13 列")
add("i_ai_amzn", "Amazon AI 占比覆寫", "比", 0.68, 0.60, 0.75, "Derived", "Wattonomics In 第 14 列")
add("i_cov", "站點為主買方的覆蓋率倒數（真值 ÷ 站點讀數）", "x", 1.0, "=1/0.77", "=1/0.44", "Analogy", "Epoch 自估站點資料庫覆蓋全球已部署算力約 44%（90% 區間 29–77%）；只作類比（peers 報告）")
add("i_rev", "七家 2027–28 AI 資本支出事後上修倍數", "x", 1.0, None, 1.65, "Analogy／Interested-party", "Coatue：2026 共識資本支出約 5,180 億美元，比年初預期高 65%（頁面 2026-11-18 讀取；頁面日期未核）")
add("i_rev_mid", "（中間點，只供參考）", "x", "=SQRT(1.65)", None, None, "Assumed", "對數中點；不入龍捲風")
sec("D. NVIDIA 顆數單因子端點")
add("i_nv_k1", "NVIDIA 2027 顆數乘數：以 2026 供給側高估幅度類比（K1 季節權重）", "x", 1.0, "=164.2/(448.36*0.424)", None, "Analogy", "K1：2026 隱含 448.36 × 上半年占比高端 42.4% 對實績上半年 164.2（FY27 Q1＋Q2，Verified）；基準 41.4% 時 0.885；線性 50% 時 0.733（不採：高成長年偏誤）")
add("i_nv_guid", "NVIDIA 2027 顆數乘數：FY2028 指引營收恆等式", "x", 1.0, "=1/1.2", None, "Interested-party／Derived", "unit-basis 4.3(c)：基準單價下供給側比指引高 13%–20%；取 20%")
add("i_rb_dup", "Rubin Ultra 重複計算：Rubin 原始顆數扣除", "百萬封裝", 0, 0.3, None, "Interested-party", "#662（A）：Rubin 7.0 → 6.7；MS 700 萬含 Rubin Ultra")
add("i_bw_die", "Blackwell 2027 若為晶粒口徑：乘數", "x", 1.0, 0.5, None, "Interested-party（原文未取得）", "unit-basis 第一節：最壞情況")
add("i_cw_cap", "CoWoS 可用產能 2027（KGI 年底 17 萬片／月）÷ 需求", "x", 1.0, "=MIN(1,(130+170)/2*12*1.1/1982.6161)", None, "Interested-party", "Siliconomics v0.20 Cap 頁：需求 1,982.6 千片；現行產能 2,178")
sec("E. Wattonomics 已算好的替代值（不重做，直接代入橋接）")
add("i_w2a_27", "W2_A：2027 Blackwell 加回後市場年新增 2027（b＝10%／16.2%）", "GW", 25.5005334169982, 25.59, 25.65, "Assumed→Analogy／Interested-party", "wattonomics/docs/reports/20261118_W2_A_2027不計Blackwell重審.md 第 3 節（兩位小數）")
add("i_w2a_28", "W2_A：同上 2028", "GW", 29.3111548208607, 29.39, 29.45, "同上", "同上")
add("i_fac_hold", "非 NVIDIA 廠房成本乘持有比（1＝是）", "旗標", 1, None, 0, "Assumed", "wattonomics/docs/reports/20261118_W2_B_非NVIDIA廠房持有比口徑.md")
# ---- 逐家逐年 ----
sec("F. Wattonomics 逐家逐年（七家＋xAI；CapexGW 頁現值，2025–2028）")
PL = R.PL; YR = R.YR
HOLDN = R.HOLDN
for q in PL:
    d = D[q]
    for i, y in enumerate(YR):
        if y < 2025: continue
        add(f"i_ai_{q}_{y}", f"{q} {y} 曆年 AI 資本支出", "US$B", d["ai"][i] if d["ai"][i] is not None else 0, None, None, "Derived", f"Wattonomics v0.3.6 CapexGW {q}_ai")
        add(f"i_a1_{q}_{y}", f"{q} {y} 廠房提前 t−1", "US$B", d["a1"][i] if d["a1"][i] is not None else 0, None, None, "Derived", f"CapexGW {q}_a1")
        add(f"i_a2_{q}_{y}", f"{q} {y} 廠房提前 t−2", "US$B", d["a2"][i] if d["a2"][i] is not None else 0, None, None, "Derived", f"CapexGW {q}_a2")
        own = (d["own"][i] if "own" in d and d["own"][i] is not None else 1.0)
        add(f"i_own_{q}_{y}", f"{q} {y} 廠房自付比", "比", own, None, None, "Derived", f"CapexGW {q}_own")
        n = 0.0 if q in ("crwv", "nbis", "xai") else R.nnv(q, y)
        add(f"i_n_{q}_{y}", f"{q} {y} 非 NVIDIA 占比（換算用）", "比", n, None, None, "Derived（由 Siliconomics Units 推得）", "Wattonomics Players 頁")
        add(f"i_x_{q}_{y}", f"{q} {y} 付款年 GW 現值（缺值代入用）", "GW", d["x"][i], None, None, "Derived", f"CapexGW {q}_x")
        add(f"i_main_{q}_{y}", f"{q} {y} 主路徑年新增現值", "GW", d["main"][i] or 0, None, None, "Derived", f"CapexGW {q}_main")
        add(f"i_has_{q}_{y}", f"{q} {y} 有資本支出值（1）", "旗標", 0 if d["ai"][i] is None else 1, None, None, "Derived", "")
    add(f"i_ye_{q}_2025", f"{q} 2025 主路徑年底", "GW", d["ye"][1] if "ye" in d else 0, None, None, "Derived", f"CapexGW {q}_ye")
# Oracle 財年資料（ai_sh 只作用 FY2025–26）
fy = D["orcl"]["fyai"]
add("i_orcl_fyai26", "Oracle FY2026 AI 資本支出", "US$B", fy[2], None, None, "Derived", "CapexGW orcl_fyai"); add("i_orcl_fyai27", "Oracle FY2027 AI 資本支出", "US$B", fy[3], None, None, "Derived", "同上（AI 比讀個股模型 0.946）")
add("i_orcl_fyai28", "Oracle FY2028 AI 資本支出", "US$B", fy[4], None, None, "Derived", "同上")
sec("G. 市場其餘（上線年）")
for y in (2027, 2028):
    i = YR.index(y)
    add(f"i_pipe_{y}", f"Pipeline 年新增 {y}", "GW", D["pipeadd"][i], None, None, "Derived", "Wattonomics Deploy 第 4 列")
    add(f"i_site_{y}", f"十一家（七家、xAI、新雲、主權、其他）站點年新增合計 {y}", "GW", sum((D[q]["site"][i] or 0) for q in R.SITE11), None, None, "Derived", "CapexGW *_site")
    add(f"i_neo_{y}", f"其他新雲 主路徑 {y}", "GW", D["neo"]["main"][i], None, None, "Derived／Assumed", "CapexGW neo_main（X6）")
    add(f"i_sov_{y}", f"主權 AI 主路徑 {y}", "GW", D["sov"]["main"][i], None, None, "Assumed", "CapexGW sov_main")
for y in (2025, 2026, 2027, 2028):
    add(f"i_mx_{y}", f"市場（不含其他）年底 現值 {y}", "GW", D["oth"]["mx"][YR.index(y)], None, None, "Derived", "CapexGW oth_mx")
add("i_othye_2025", "其他 2025 年底", "GW", D["oth"]["ye"][1], None, None, "Assumed", "CapexGW oth_ye")
MIX = R.base_params()["mix"]
for y in (2025, 2026, 2027, 2028):
    for g in GENS:
        add(f"i_mix_{g}_{y}", f"NVIDIA 世代占比 {g} {y}", "比", MIX[y].get(g, 0.0), None, None, "Derived／Assumed", "Wattonomics Players／In mix_*（2027 Blackwell＝0 為 v0.1 簡化）")
add("i_crwv_q3", "CoreWeave 2026 主路徑下限（Q3 實績 − 2025 年底）", "GW", 0.85, None, None, "Verified／Derived", "CapexGW crwv_main 2026 ＝MAX(上線年, w1_crwv_q3/1000 − 2025 年底)")
sec("H. 對照值（什麼必須為真用）")
add("i_sil_nvrev", "Siliconomics 2027 隱含 NVIDIA 資料中心營收（凍結）", "US$B", 743.78, None, None, "Derived", "adjudication/20261118_凍結紀錄_2027年隱含值.md")
add("i_wat_nvrev", "Wattonomics 2027 隱含 NVIDIA 資料中心營收（凍結）", "US$B", 408.76, None, None, "Derived", "同上")
add("i_guid", "NVIDIA FY2028 資料中心營收（指引 +70% 推得）", "US$B", 633.7, None, None, "Interested-party", "Siliconomics v0.16.1；unit-basis 4.0")
add("i_nv26", "NVIDIA CY2026 資料中心營收錨點（Q1、Q2 實績＋Q3 指引中點）", "US$B", 375.46, 363.4, 387.5, "Verified／Interested-party", "Wattonomics In nv_dc_cy26（Q2 法說時點）")
add("i_site7", "七家 2027 站點實體讀數", "GW", 7.50, None, None, "Derived", "交接檔第 4 節")
add("i_cap7", "七家 2027 資本支出路徑", "GW", 18.16, None, None, "Derived", "同上")
add("i_capex8_27", "七家＋xAI 2027 曆年 AI 資本支出合計", "US$B", "=" + "+".join(f"i_ai_{q}_2027" for q in PL), None, None, "Derived", "")
add("i_tk_tot_vr", "Tokenomics VR200 每 GW 總資本支出（IT＋廠房）", "US$B/GW", "=i_CIT_VR200+i_FAC_VR200", None, None, "Derived", "")
IN.column_dimensions["A"].width = 22; IN.column_dimensions["B"].width = 58; IN.column_dimensions["H"].width = 90

# ================= Calc =================
CA = wb.create_sheet("Calc")
CA["A1"] = ("Calc — 每一欄是一個情境；上半部是該情境的參數（只有被動的因子改用端點，其餘＝基準），下半部是公式鏈："
            "供給（原始顆數 × 截頂 → 成本等值 GW、實體 GW）；需求（逐家 AI 資本支出 → 付款年 GW → 上線年 → 市場年新增 → 平滑法出貨 s2）；橋接 x2＝s2 ÷（實體 ÷ 成本等值）；差距＝供給 − x2。")
ROWS = {}; ORDER = []
def row(key, label):
    ROWS[key] = len(ORDER) + 4; ORDER.append((key, label))
# 參數列
PARAM = []
for g in GENS: PARAM += [f"CIT_{g}", f"FAC_{g}", f"GPG_{g}"]
PARAM += [f"HOLD_{h}" for h in ["AMD", "TPU", "Trn", "Maia", "Ascend", "OAI"]]
PARAM += [f"raw_{k}" for k in RAW] + ["cap", "recap", "hbm_sup"] + [f"it_{k}" for k in IT] + [f"rconv_{k}" for k in RC]
PARAM += ["lag_on", "inv_on", "lag_bldg", "oth_gmult", "cov", "fac_hold", "dwt27", "dwt28", "readF"]
PARAM += [f"m_{q}_{y}" for q in PL for y in YR]
DEFAULT = {**{f"CIT_{g}": f"i_CIT_{g}" for g in GENS}, **{f"FAC_{g}": f"i_FAC_{g}" for g in GENS}, **{f"GPG_{g}": f"i_GPG_{g}" for g in GENS},
           **{f"HOLD_{h}": f"i_HOLD_{h}" for h in ["AMD", "TPU", "Trn", "Maia", "Ascend", "OAI"]}, **{f"raw_{k}": f"i_raw_{k}" for k in RAW},
           "cap": "i_cap", "recap": "0", "hbm_sup": "i_hbm_now", **{f"it_{k}": f"i_it_{k}" for k in IT}, **{f"rconv_{k}": f"i_rconv_{k}" for k in RC},
           "lag_on": "i_lag_on", "inv_on": "i_inv_on", "lag_bldg": "i_lag_bldg", "oth_gmult": "i_oth_gmult", "cov": "i_cov", "fac_hold": "i_fac_hold",
           "dwt27": "0", "dwt28": "0", "readF": "0", **{f"m_{q}_{y}": "1" for q in PL for y in YR}}
for k in PARAM: row("p_" + k, "參數 " + k)
# 公式鏈列
for y in (2025, 2026, 2027, 2028): row(f"F7_{y}", f"NVIDIA 每 GW IT capex（世代加權）{y}"); row(f"F8_{y}", f"NVIDIA 每 GW 廠房 capex {y}")
row("w0", "廠房權重 w0"); row("w1", "w1"); row("w2", "w2"); row("k", "k＝lag_on/12＋inv_on")
for q in PL:
    for y in (2025, 2026, 2027, 2028):
        for kk in ("pit", "pfac", "pfo", "ai", "a1", "a2", "pay"): row(f"{kk}_{q}_{y}", f"{q} {y} {kk}")
    for y in (2026, 2027, 2028): row(f"on_{q}_{y}", f"{q} {y} 上線年新增")
    for y in (2026, 2027, 2028): row(f"dye_{q}_{y}", f"{q} {y} 年底差（對現值）")
for y in (2026, 2027, 2028): row(f"mx_{y}", f"市場（不含其他）年底 {y}")
row("othye_2025", "其他年底 2025")
for y in (2026, 2027, 2028): row(f"othye_{y}", f"其他年底 {y}"); row(f"oth_{y}", f"其他年新增 {y}")
for y in (2027, 2028): row(f"wt_{y}", f"Wattonomics 市場年新增（上線年）{y}")
row("s2", "平滑法出貨 s2"); row("othon27", "七家與 xAI 以外 上線 2027"); row("othon28", "同上 2028"); row("pay8_27", "七家＋xAI 付款年 GW 2027"); row("shipF", "付款年讀法出貨（Compare F 同式）")
row("hbm_dem", "HBM 需求（原始顆數）EB"); row("c", "截頂係數（採用）")
for k in RAW: row(f"u_{k}", f"截頂後顆數 {k}")
row("it_bw", "it GB300"); row("it_rb", "it VR200"); row("it_ru", "it RU"); row("it_hp", "it H100")
row("S0", "供給 成本等值 GW"); row("nvphys", "NVIDIA 實體 GW"); row("nnphys", "非 NVIDIA 實體 GW"); row("phys", "供給 實體 GW"); row("rt", "實體 ÷ 成本等值")
row("x2", "需求側換算值 x2（依讀法）"); row("G", "差距（成本等值）＝S0 − x2"); row("Gphys", "差距（實體）＝實體 − 出貨"); row("nvunits", "NVIDIA 顆數（百萬封裝）")
for key, lab in ORDER: CA.cell(ROWS[key], 1, key); CA.cell(ROWS[key], 2, lab)
CA.cell(2, 1, "情境"); CA.cell(3, 1, "說明")
def chain(col):
    c = CL(col); r = lambda k: f"{c}{ROWS[k]}"
    F = {}
    for y in (2025, 2026, 2027, 2028):
        F[f"F7_{y}"] = "=" + "+".join(f"i_mix_{g}_{y}*{r('p_CIT_' + g)}" for g in GENS)
        F[f"F8_{y}"] = "=" + "+".join(f"i_mix_{g}_{y}*{r('p_FAC_' + g)}" for g in GENS)
    L = f"{r('p_lag_bldg')}/12"
    F["w0"] = f"=MAX(0,1-{L})"; F["w1"] = f"=MAX(0,MIN({L},2-{L}))"; F["w2"] = f"=MAX(0,MIN(1,{L}-1))"
    F["k"] = f"={r('p_lag_on')}/12+{r('p_inv_on')}"
    yi = {y: YR.index(y) for y in YR}
    for q in PL:
        h = HOLDN[q]; rk = {"Maia": "oth", "TPU": "tpu", "Trn": "trn", "AMD": "amd"}[h]
        for y in (2025, 2026, 2027, 2028):
            m0, m1, m2 = r(f"p_m_{q}_{y}"), r(f"p_m_{q}_{y - 1}"), r(f"p_m_{q}_{max(y - 2, 2024)}")
            F[f"ai_{q}_{y}"] = f"=i_ai_{q}_{y}*{m0}"; F[f"a1_{q}_{y}"] = f"=i_a1_{q}_{y}*{m1}"; F[f"a2_{q}_{y}"] = f"=i_a2_{q}_{y}*{m2}"
            n = f"i_n_{q}_{y}"
            F[f"pit_{q}_{y}"] = f"=(1-{n})*{r('F7_' + str(y))}+{n}*{r('p_HOLD_' + h)}*{r('p_CIT_VR200')}/{r('p_rconv_' + rk)}"
            F[f"pfac_{q}_{y}"] = f"=(1-{n})*{r('F8_' + str(y))}+{n}*IF({r('p_fac_hold')}=1,{r('p_HOLD_' + h)},1)*{r('p_FAC_VR200')}"
            F[f"pfo_{q}_{y}"] = f"=i_own_{q}_{y}*{r(f'pfac_{q}_{y}')}"
            pit, pfo = r(f"pit_{q}_{y}"), r(f"pfo_{q}_{y}"); s = f"({pit}+{pfo})"
            F[f"pay_{q}_{y}"] = (f"=IF(i_has_{q}_{y}=1,({pit}/{s}*{r(f'ai_{q}_{y}')}+{pfo}/{s}*({r('w0')}*{r(f'ai_{q}_{y}')}+{r('w1')}*{r(f'a1_{q}_{y}')}+{r('w2')}*{r(f'a2_{q}_{y}')}))/{s},i_x_{q}_{y})")
        cum = []
        for y in (2026, 2027, 2028):
            p, pp = r(f"pay_{q}_{y}"), r(f"pay_{q}_{y - 1}")
            base = f"{p}-{r('k')}*({p}-{pp})"
            if q == "xai" and y == 2026: f = f"={p}"
            elif q == "crwv" and y == 2026: f = f"=MAX({base},i_crwv_q3)"
            else: f = "=" + base
            F[f"on_{q}_{y}"] = f
            cum.append(f"({r(f'on_{q}_{y}')}-i_main_{q}_{y})")
            F[f"dye_{q}_{y}"] = "=" + "+".join(cum)
    for y in (2026, 2027, 2028): F[f"mx_{y}"] = f"=i_mx_{y}+" + "+".join(r(f"dye_{q}_{y}") for q in PL)
    F["othye_2025"] = "=i_othye_2025"
    for y in (2026, 2027, 2028):
        mxp = "i_mx_2025" if y == 2026 else r(f"mx_{y - 1}")
        F[f"othye_{y}"] = f"={r(f'othye_{y - 1}')}*(1+({r(f'mx_{y}')}/{mxp}-1)*{r('p_oth_gmult')})"
        F[f"oth_{y}"] = f"={r(f'othye_{y}')}-{r(f'othye_{y - 1}')}"
    for y, dk in ((2027, "p_dwt27"), (2028, "p_dwt28")):
        F[f"wt_{y}"] = (f"=(i_pipe_{y}-i_site_{y}+i_neo_{y})*{r('p_cov')}+i_sov_{y}+{r(f'oth_{y}')}+" + "+".join(r(f"on_{q}_{y}") for q in PL) + f"+{r(dk)}")
    F["s2"] = f"={r('wt_2027')}+{r('k')}*({r('wt_2028')}-{r('wt_2027')})"
    F["othon27"] = f"={r('wt_2027')}-(" + "+".join(r(f"on_{q}_2027") for q in PL) + ")"
    F["othon28"] = f"={r('wt_2028')}-(" + "+".join(r(f"on_{q}_2028") for q in PL) + ")"
    F["pay8_27"] = "=" + "+".join(r(f"pay_{q}_2027") for q in PL)
    F["shipF"] = f"={r('pay8_27')}+{r('othon27')}+{r('k')}*({r('othon28')}-{r('othon27')})"
    F["hbm_dem"] = "=i_hbm_gross*(" + "+".join(f"{r('p_raw_' + k)}*i_hbmgb_{k}" for k in RAW) + ")/1000"
    F["c"] = f"=IF({r('p_recap')}=1,MIN(1,{r('p_hbm_sup')}/{r('hbm_dem')}),{r('p_cap')})"
    for k in RAW: F[f"u_{k}"] = f"={r('p_raw_' + k)}*{r('c')}"
    for k, g in (("bw", "GB300"), ("rb", "VR200"), ("ru", "RU"), ("hp", "H100")):
        F[f"it_{k}"] = f"={r('p_CIT_' + g)}/{r('p_GPG_' + g)}*1000000"
    F["S0"] = "=(" + "+".join(f"{r('u_' + k)}*{r('it_' + k)}" for k in ("bw", "rb", "ru", "hp")) + "+" + "+".join(f"{r('u_' + k)}*{r('p_it_' + k)}" for k in IT) + f")/{r('p_CIT_VR200')}"
    F["nvphys"] = "=" + "+".join(f"{r('u_' + k)}*1000000/{r('p_GPG_' + g)}" for k, g in (("bw", "GB300"), ("rb", "VR200"), ("ru", "RU"), ("hp", "H100")))
    holdx = {"amd": r("p_HOLD_AMD"), "tpu": r("p_HOLD_TPU"), "trn": r("p_HOLD_Trn"), "oth": f"(({r('p_HOLD_Maia')}+{r('p_HOLD_Ascend')}+{r('p_HOLD_OAI')})/3)"}
    F["nnphys"] = "=" + "+".join(f"{r('u_' + k)}*{r('p_it_' + k)}/({r('p_CIT_VR200')}*{holdx[k]})*{r('p_rconv_' + k)}" for k in IT)
    F["phys"] = f"={r('nvphys')}+{r('nnphys')}"; F["rt"] = f"={r('phys')}/{r('S0')}"
    F["x2"] = f"=IF({r('p_readF')}=1,{r('shipF')},{r('s2')})/{r('rt')}"
    F["G"] = f"={r('S0')}-{r('x2')}"; F["Gphys"] = f"={r('phys')}-IF({r('p_readF')}=1,{r('shipF')},{r('s2')})"
    F["nvunits"] = f"={r('u_bw')}+{r('u_rb')}+{r('u_ru')}+{r('u_hp')}"
    return F
# ---- 情境 ----
def mpow(j):   # 功率端點：GPG、CIT 同乘 基準 kW ÷ 端點 kW
    o = {}
    for g in GENS:
        f = f"i_KW_{g}/i_KW_{g}_{j}"
        o[f"CIT_{g}"] = f"i_CIT_{g}*{f}"; o[f"GPG_{g}"] = f"i_GPG_{g}*{f}"
    return o
def moff(j):
    o = {}
    for g in GENS:
        f = f"(1-i_offrack_{j})/(1-i_offrack)"
        o[f"CIT_{g}"] = f"i_CIT_{g}*{f}"; o[f"GPG_{g}"] = f"i_GPG_{g}*{f}"
    return o
def mcost(j): return {**{f"CIT_{g}": f"i_CIT_{g}*(i_CIT_{g}_{j}/i_CIT_{g})" for g in GENS}, **{f"FAC_{g}": f"i_FAC_{g}_{j}" for g in GENS}}
def mnv(mult): return {f"raw_{k}": f"i_raw_{k}*{mult}" for k in ("bw", "rb", "ru")}
def maish(j):
    f = f"(i_ai_sh_{j}/i_ai_sh)"; fa = f"(i_ai_amzn_{j}/i_ai_amzn)"
    o = {}
    for q in ("msft", "googl", "meta", "crwv", "nbis"):
        for y in YR: o[f"m_{q}_{y}"] = f
    for y in YR: o[f"m_amzn_{y}"] = fa
    o["m_orcl_2024"] = f; o["m_orcl_2025"] = f
    o["m_orcl_2026"] = f"(5/12*i_orcl_fyai26*{f}+7/12*i_orcl_fyai27)/(5/12*i_orcl_fyai26+7/12*i_orcl_fyai27)"
    return o
CONS = {"googl": ("300.979/290.3651", "339.451/317.3988"), "amzn": ("275.063/253.0801", "292.927/266.9716"), "crwv": ("46.476/42.5812", "50.129/34.2227"), "nbis": ("27.243/31.9887", "38.39/32.9622")}
def mcons():
    o = {}
    for q, (a, b) in CONS.items(): o[f"m_{q}_2027"] = a; o[f"m_{q}_2028"] = b
    o["m_orcl_2027"] = "(5/12*i_orcl_fyai27+7/12*100.962*0.938578542167978)/i_ai_orcl_2027"
    o["m_orcl_2028"] = "(5/12*100.962*0.938578542167978+7/12*84.848*0.938578542167978)/i_ai_orcl_2028"
    return o
def mrev(x): return {f"m_{q}_{y}": x for q in PL if q != "xai" for y in (2027, 2028)}
def mrev8(x): return {f"m_{q}_{y}": x for q in PL for y in (2027, 2028)}
FX = {
 "S1": ({"cap": "MIN(1,i_cap*i_hbm_ind/i_hbm_now)"}, {"cap": "1"}),
 "S2": (mnv("i_nv_k1_lo"), {}), "S3": (mnv("i_nv_guid_lo"), {}),
 "S4": ({"raw_rb": "i_raw_rb-i_rb_dup_lo"}, {}), "S5": ({"raw_bw": "i_raw_bw*i_bw_die_lo"}, {}),
 "S6": ({"cap": "i_cap*i_cw_cap_lo"}, {}),
 "S7": ({f"raw_{k}": f"i_raw_{k}_lo" for k in LOHI}, {f"raw_{k}": f"i_raw_{k}_hi" for k in LOHI}),
 "S8": ({f"it_{k}": f"i_it_{k}_lo" for k in IT}, {f"it_{k}": f"i_it_{k}_hi" for k in IT}),
 "S9": ({}, {f"rconv_{k}": f"i_rconv_{k}_hi" for k in RC}),
 "T1": (mpow("lo"), mpow("hi")), "T2": (moff("lo"), moff("hi")), "T3": ({}, {}), "T4": (mcost("lo"), mcost("hi")),
 "D1": (maish("lo"), maish("hi")), "D2": ({}, mcons()), "D3": ({}, mrev("i_rev_hi")),
 "D4": ({"lag_bldg": "i_lag_bldg_lo"}, {"lag_bldg": "i_lag_bldg_hi"}), "D5": ({"lag_on": "i_lag_on_lo"}, {"lag_on": "i_lag_on_hi"}),
 "D6": ({"inv_on": "i_inv_on_lo"}, {"inv_on": "i_inv_on_hi"}), "D7": ({"cov": "i_cov_lo"}, {"cov": "i_cov_hi"}),
 "D8": ({"oth_gmult": "i_oth_gmult_lo"}, {"oth_gmult": "i_oth_gmult_hi"}), "D9": ({}, {}),
 "D10": ({"dwt27": "i_w2a_27_lo-i_w2a_27", "dwt28": "i_w2a_28_lo-i_w2a_28"}, {"dwt27": "i_w2a_27_hi-i_w2a_27", "dwt28": "i_w2a_28_hi-i_w2a_28"}),
 "D11": ({}, {"fac_hold": "i_fac_hold_hi"}), "D12": ({}, {"readF": "1"}),
}
SC = [("base", "基準", {})]
for fid, (lo, hi) in FX.items():
    SC.append((f"{fid}_lo", f"{fid} 低端", lo)); SC.append((f"{fid}_hi", f"{fid} 高端", hi))
for fid in ("S2", "S3", "S4", "S5", "S7"):
    lo, hi = FX[fid]
    SC.append((f"{fid}_lo_rc", f"{fid} 低端（截頂依 HBM 重算）", {**lo, "recap": "1"})); SC.append((f"{fid}_hi_rc", f"{fid} 高端（截頂重算）", {**hi, "recap": "1"}))
MT = RES["must"]
JO = {"J_sup": ["S3", "S4", "S7", "S8"], "J_dem": ["D1", "D2", "D4", "D7", "D8"], "J_phy": ["T1", "T2", "T4", "S9"], "J_ev": ["S3", "S4", "D2"]}
CLOSE = {r["id"]: ("lo" if r["d_lo"] <= r["d_hi"] else "hi") for r in RES["factors"]}
def compose(ids):
    o = {}
    for i in ids:
        ov = FX[i][0] if CLOSE[i] == "lo" else FX[i][1]
        for k, v in ov.items():
            if k in o and k.startswith("m_"): o[k] = f"({o[k]})*({v})"
            elif k in o and k.startswith(("CIT_", "GPG_")): o[k] = f"({o[k]})*" + v[len("i_" + k + "*"):]
            elif k in o and k == "raw_rb": o[k] = v.replace("i_raw_rb", f"({o[k]})")
            elif k in o and k.startswith("raw_") and "i_nv_guid_lo" in o[k]: o[k] = f"({o[k]})"   # 同參數只取一次
            else: o[k] = v
    return o
for jid, ids in JO.items(): SC.append((jid, "聯合：" + "、".join(ids), compose(ids)))
# 交互作用
SC.append(("X_S3T1", "S3 低 × T1 低", {**mnv("i_nv_guid_lo"), **mpow("lo")}))
SC.append(("X_D1D3", "D1 高 × D3 高", {k: f"({maish('hi').get(k, '1')})*({mrev('i_rev_hi').get(k, '1')})" for k in set(maish("hi")) | set(mrev("i_rev_hi"))}))
SC.append(("X_S1S3", "S1 低 × S3 低", {"cap": "MIN(1,i_cap*i_hbm_ind/i_hbm_now)", **mnv("i_nv_guid_lo")}))
# 什麼必須為真（解出的值放 Must 頁輸入，這裡驗證差距＝0）
SC.append(("M_nv", "需求側正確：NVIDIA 顆數乘數", mnv("Must!C5")))
SC.append(("M_all", "需求側正確：全部顆數乘數", {"cap": "i_cap*Must!C6"}))
SC.append(("M_hbm", "需求側正確：HBM 供給（截頂重算）", {"recap": "1", "hbm_sup": "Must!C7"}))
SC.append(("M_cap8", "供給側正確：七家＋xAI 2027–28 AI 資本支出倍數（平滑法）", mrev8("Must!C9")))
SC.append(("M_cap8F", "供給側正確：同上（付款年讀法）", {**mrev8("Must!C10"), "readF": "1"}))
SC.append(("M_extra", "供給側正確：七家以外每年多上線 GW", {"dwt27": "Must!C11", "dwt28": "Must!C11"}))
SC.append(("M_cov", "供給側正確：站點為主買方覆蓋率倒數", {"cov": "Must!C12"}))
COLS = {}
for j, (sid, desc, ov) in enumerate(SC):
    col = 3 + j; COLS[sid] = col; c = CL(col)
    CA.cell(2, col, sid).font = BOLD; CA.cell(3, col, desc)
    for k in PARAM:
        v = ov.get(k, DEFAULT[k])
        CA.cell(ROWS["p_" + k], col, "=" + v)
    for k, f in chain(col).items(): CA.cell(ROWS[k], col, f)
    for k in ("S0", "x2", "G", "Gphys"): CA.cell(ROWS[k], col).fill = KEY
CA.column_dimensions["A"].width = 16; CA.column_dimensions["B"].width = 34
CA.freeze_panes = "C4"
G = lambda sid, k="G": f"Calc!{CL(COLS[sid])}{ROWS[k]}"

# ================= Tornado =================
TO = wb.create_sheet("Tornado", 0)
TO["A1"] = "龍捲風表：2027 年差距（Siliconomics 成本等值 − Wattonomics 換算值 x2），一次只動一個因子。依影響幅度排序（排序以本次重算值決定；『排名』欄為公式）。負＝差距縮小。"
TO["A2"] = "基準差距（GW）"; TO["C2"] = "=" + G("base"); TO["C2"].fill = KEY
TO["D2"] = "基準：供給 S0"; TO["E2"] = "=" + G("base", "S0"); TO["F2"] = "x2"; TO["G2"] = "=" + G("base", "x2")
hdrs = ["排名（公式）", "編號", "類", "側", "因子", "等級", "低端（依據）", "高端（依據）", "低端時差距", "高端時差距", "Δ低", "Δ高", "幅度", "縮小差距最多的一端 Δ", "占基準差距",
        "Δ低（截頂重算）", "Δ高（截頂重算）", "Δ低（付款年讀法）", "Δ高（付款年讀法）", "實體 GW 差距 Δ低", "實體 Δ高", "狀態", "出處"]
for i, h in enumerate(hdrs): TO.cell(4, i + 1, h).font = BOLD
for n, f in enumerate(RES["factors"]):
    r = 5 + n; fid = f["id"]
    TO.cell(r, 2, fid); TO.cell(r, 3, f["cls"]); TO.cell(r, 4, f["side"]); TO.cell(r, 5, f["name"]); TO.cell(r, 6, f["grade"])
    TO.cell(r, 7, f["lo_lab"]); TO.cell(r, 8, f["hi_lab"])
    TO.cell(r, 9, "=" + G(f"{fid}_lo")); TO.cell(r, 10, "=" + G(f"{fid}_hi"))
    TO.cell(r, 11, f"=I{r}-$C$2"); TO.cell(r, 12, f"=J{r}-$C$2"); TO.cell(r, 13, f"=MAX(ABS(K{r}),ABS(L{r}))"); TO.cell(r, 14, f"=MIN(K{r},L{r})")
    TO.cell(r, 15, f"=N{r}/$C$2")
    if f"{fid}_lo_rc" in COLS:
        TO.cell(r, 16, f"={G(fid + '_lo_rc')}-$C$2"); TO.cell(r, 17, f"={G(fid + '_hi_rc')}-$C$2")
    if fid != "D12":
        TO.cell(r, 18, f"=({G(fid + '_lo', 'S0')}-{G(fid + '_lo', 'shipF')}/{G(fid + '_lo', 'rt')})-({G('base', 'S0')}-{G('base', 'shipF')}/{G('base', 'rt')})")
        TO.cell(r, 19, f"=({G(fid + '_hi', 'S0')}-{G(fid + '_hi', 'shipF')}/{G(fid + '_hi', 'rt')})-({G('base', 'S0')}-{G('base', 'shipF')}/{G('base', 'rt')})")
    TO.cell(r, 20, f"={G(fid + '_lo', 'Gphys')}-{G('base', 'Gphys')}"); TO.cell(r, 21, f"={G(fid + '_hi', 'Gphys')}-{G('base', 'Gphys')}")
    TO.cell(r, 22, f["status"]); TO.cell(r, 23, f["src"])
    TO.cell(r, 1, f"=RANK(M{r},$M$5:$M${4 + len(RES['factors'])})")
    for c in (9, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 21): TO.cell(r, c).number_format = "+0.00;-0.00;0.00"
    TO.cell(r, 15).number_format = "0%"
for c, w in zip("ABCDEFGHIJKLMNOPQRSTUVW", (8, 6, 4, 6, 50, 22, 36, 36, 10, 10, 8, 8, 8, 10, 9, 10, 10, 10, 10, 9, 9, 30, 60)): TO.column_dimensions[c].width = w
# ================= Joint =================
JS = wb.create_sheet("Joint", 1)
JS["A1"] = "殘差與交互作用：每組因子各取『縮小差距』那一端，同時設定（聯合）對照單因子加總。交互＝聯合 − 加總；剩餘差距＝基準 + 聯合。不硬湊到 100%。"
for i, h in enumerate(["組", "因子", "單因子加總 Δ", "聯合 Δ", "交互作用", "聯合後剩餘差距", "剩餘 ÷ 基準"]): JS.cell(3, i + 1, h).font = BOLD
TROW = {f["id"]: 5 + n for n, f in enumerate(RES["factors"])}
JN = {"J_sup": "供給側：NVIDIA 顆數取指引（S3）＋Rubin Ultra（S4）＋非 NVIDIA 顆數低（S7）＋每顆 IT capex 低（S8）",
      "J_dem": "需求側：AI 占比高（D1）＋共識（D2）＋廠房 12 月（D4）＋覆蓋 44%（D7）＋其他 1.5（D8）",
      "J_phy": "共用物理：機架功率（T1）＋機架外（T2）＋成本情境（T4）＋非 NVIDIA 換算（S9），各取縮小端",
      "J_ev": "只取有外部證據端點者：指引（S3）＋Rubin Ultra（S4）＋共識（D2）"}
for n, (jid, ids) in enumerate(JO.items()):
    r = 4 + n
    JS.cell(r, 1, JN[jid]); JS.cell(r, 2, "、".join(ids))
    JS.cell(r, 3, "=" + "+".join(f"Tornado!N{TROW[i]}" for i in ids)); JS.cell(r, 4, f"={G(jid)}-Tornado!$C$2")
    JS.cell(r, 5, f"=D{r}-C{r}"); JS.cell(r, 6, f"={G(jid)}"); JS.cell(r, 7, f"=F{r}/Tornado!$C$2")
r = 10; JS.cell(r, 1, "兩因子交互作用").font = BOLD
for i, h in enumerate(["組合", "A 單獨 Δ", "B 單獨 Δ", "A＋B 同時 Δ", "交互作用"]): JS.cell(r + 1, i + 1, h).font = BOLD
for n, (sid, a, b, lab) in enumerate([("X_S3T1", "S3", "T1", "NVIDIA 顆數（指引）× 機架低功率：每 GW 顆數變多時，少掉的顆數換成的 GW 也變少"),
                                       ("X_D1D3", "D1", "D3", "AI 占比 × 資本支出上修：相乘"),
                                       ("X_S1S3", "S1", "S3", "HBM 獨立估計 × 指引：截頂後同比例再縮")]):
    rr = r + 2 + n
    JS.cell(rr, 1, lab); JS.cell(rr, 2, f"=MIN(Tornado!K{TROW[a]},Tornado!L{TROW[a]})"); JS.cell(rr, 3, f"=MIN(Tornado!K{TROW[b]},Tornado!L{TROW[b]})")
    JS.cell(rr, 4, f"={G(sid)}-Tornado!$C$2"); JS.cell(rr, 5, f"=D{rr}-B{rr}-C{rr}")
JS.cell(16, 1, "NVIDIA 顆數（S3）在截頂係數固定 vs 依 HBM 重算"); JS.cell(16, 2, f"=Tornado!K{TROW['S3']}"); JS.cell(16, 3, f"=Tornado!P{TROW['S3']}")
JS.column_dimensions["A"].width = 90; JS.column_dimensions["B"].width = 22
for rr in range(4, 17):
    for cc in range(2, 8): JS.cell(rr, cc).number_format = "+0.00;-0.00;0.00"
JS["G4"].number_format = JS["G5"].number_format = JS["G6"].number_format = JS["G7"].number_format = "0%"
# ================= Must =================
MU = wb.create_sheet("Must", 2)
MU["A1"] = "什麼必須為真：讓 2027 差距歸零所需的單一參數值。C 欄藍字＝Python 二分法解出（scripts/sens.py），D 欄＝以該值在 Calc 重算的差距（應約 0），E 欄＝對照的獨立證據。"
for i, h in enumerate(["前提", "需要", "值", "以此值重算的差距", "對照"]): MU.cell(3, i + 1, h).font = BOLD
musts = [
 (5, "需求側正確", "NVIDIA 2027 顆數乘數（截頂固定）", MT["s_nv_mult_fixed"], "M_nv", "=C5*i_sil_nvrev", "對應 NVIDIA 2027 資料中心營收（凍結值等比，未含價格變動）；指引 i_guid、2026 錨點 i_nv26"),
 (6, "需求側正確", "全部平台顆數乘數", MT["s_all_mult"], "M_all", "", ""),
 (7, "需求側正確", "2027 HBM 位元供給（EB；截頂重算）", MT["s_hbm_sup_recap"], "M_hbm", "", "對照：現行 6.67、獨立估計 3.82（2.06–6.97）"),
 (9, "供給側正確", "七家＋xAI 2027–28 AI 資本支出倍數（平滑法 x2）", MT["d_capex_mult_8"], "M_cap8", "", "對照：Coatue 2026 共識上修 1.65 倍"),
 (10, "供給側正確", "同上（付款年讀法）", MT["d_capex_mult_8_pay"], "M_cap8F", "", ""),
 (11, "供給側正確", "七家與 xAI 以外，每年多上線 GW（2027、2028 各加）", MT["d_extra_on_gw"], "M_extra", "=C11*i_tk_tot_vr", "換成資本支出（US$B／年，Tokenomics VR200 每 GW 總資本支出）；對照七家＋xAI 2027 AI 資本支出 i_capex8_27"),
 (12, "供給側正確", "站點為主買方的覆蓋率倒數（真值 ÷ 讀數）", MT["d_cov"], "M_cov", "=1/C12", "隱含覆蓋率；對照 Epoch 90% 區間 29%–77%"),
]
for r, a, b, v, sid, e, note in musts:
    MU.cell(r, 1, a); MU.cell(r, 2, b); MU.cell(r, 3, v).font = BLUE; MU.cell(r, 4, "=" + G(sid)); MU.cell(r, 5, e); MU.cell(r, 6, note)
MU.cell(14, 1, "需求側每 GW 成本（IT＋廠房）只在需求側縮放、供給側不動時須為 TK 的倍數"); MU.cell(14, 3, MT["d_cost_mult"]).font = BLUE
MU.cell(14, 6, "Python 計算（Calc 無對應欄：兩層共用 Tokenomics，Excel 鏈內無法只改一側）；對照 TK 低成本情境 42.21 ÷ 50.26＝0.84")
MU.cell(16, 1, "NVIDIA 2027 隱含營收：供給側 ÷ 指引"); MU.cell(16, 3, "=i_sil_nvrev/i_guid")
MU.cell(17, 1, "NVIDIA 2027 隱含營收：需求側 ÷ 指引"); MU.cell(17, 3, "=i_wat_nvrev/i_guid")
MU.cell(18, 1, "需求側正確時 NVIDIA 2027 營收 ÷ 2026 錨點"); MU.cell(18, 3, "=E5/i_nv26")
MU.cell(19, 1, "需求側正確時，2027 年出貨但未上線的 NVIDIA 封裝（百萬；若供給側顆數仍為真）"); MU.cell(19, 3, f"={G('base', 'nvunits')}*(1-C5)")
MU.cell(20, 1, "七家 2027 站點讀數 ÷ 資本支出路徑"); MU.cell(20, 3, "=i_site7/i_cap7")
MU.column_dimensions["A"].width = 60; MU.column_dimensions["B"].width = 52; MU.column_dimensions["F"].width = 70
# ================= Check =================
CK = wb.create_sheet("Check")
CK["A1"] = "核對：(1) 本活頁簿公式 vs Python 重算（scripts/replica.py、sens.py）；(2) Python 重算 vs 兩層活頁簿 LibreOffice 實算（scripts/uno_runs.py）。"
for i, h in enumerate(["項目", "本活頁簿", "Python（藍字）", "差"]): CK.cell(3, i + 1, h).font = BOLD
r = 4
for f in RES["factors"]:
    for e in ("lo", "hi"):
        CK.cell(r, 1, f"{f['id']}_{e} 差距"); CK.cell(r, 2, "=" + G(f"{f['id']}_{e}")); CK.cell(r, 3, f[f"G_{e}"]).font = BLUE; CK.cell(r, 4, f"=B{r}-C{r}"); r += 1
CK.cell(r, 1, "基準差距"); CK.cell(r, 2, "=" + G("base")); CK.cell(r, 3, RES["G0"]).font = BLUE; CK.cell(r, 4, f"=B{r}-C{r}"); r += 1
last = r - 1
CK.cell(r, 1, "最大絕對差"); CK.cell(r, 4, f"=MAX(MAX(D4:D{last}),-MIN(D4:D{last}))"); r += 2
CK.cell(r, 1, "兩層活頁簿實算（LibreOffice）對照 Python").font = BOLD; r += 1
UNOROWS = [("Wattonomics 基準 2027 市場年新增", UNO["wat"]["base"]["WIF_MktGW_Add_2027"][0], f"=Calc!{CL(COLS['base'])}{ROWS['wt_2027']}"),
           ("Wattonomics ai_sh 0.95：2027", UNO["wat"]["ai_hi"]["WIF_MktGW_Add_2027"][0], f"=Calc!{CL(COLS['D1_hi'])}{ROWS['wt_2027']}"),
           ("Wattonomics lag_on 3：2027", UNO["wat"]["lag_on_lo"]["WIF_MktGW_Add_2027"][0], f"=Calc!{CL(COLS['D5_lo'])}{ROWS['wt_2027']}"),
           ("Wattonomics inv_on 20%：2028", UNO["wat"]["inv_hi"]["WIF_MktGW_Add_2028"][0], f"=Calc!{CL(COLS['D6_hi'])}{ROWS['wt_2028']}"),
           ("Wattonomics lag_bldg 12：2027", UNO["wat"]["bldg_lo"]["WIF_MktGW_Add_2027"][0], f"=Calc!{CL(COLS['D4_lo'])}{ROWS['wt_2027']}"),
           ("Wattonomics oth_gmult 0：2027", UNO["wat"]["oth_lo"]["WIF_MktGW_Add_2027"][0], f"=Calc!{CL(COLS['D8_lo'])}{ROWS['wt_2027']}"),
           ("Wattonomics TK 低成本：2027", UNO["wat"]["cost_lo"]["WIF_MktGW_Add_2027"][0], None),
           ("Siliconomics 基準 成本等值", UNO["sil"]["base"]["SIF_DeployGW_Base"][1], f"=Calc!{CL(COLS['base'])}{ROWS['S0']}"),
           ("Siliconomics 基準 實體", UNO["sil"]["base"]["SIF_DeployGWPhys_Base"][1], f"=Calc!{CL(COLS['base'])}{ROWS['phys']}"),
           ("Siliconomics NVIDIA × 0.884（S2 類）", UNO["sil"]["nv_x0884"]["SIF_DeployGW_Base"][1], None),
           ("Siliconomics 低功率（T1 低）成本等值", UNO["sil"]["pow_lo"]["SIF_DeployGW_Base"][1], f"=Calc!{CL(COLS['T1_lo'])}{ROWS['S0']}"),
           ("Siliconomics TK 低成本 成本等值", UNO["sil"]["cost_lo"]["SIF_DeployGW_Base"][1], None),
           ("Siliconomics it_TPU 35 成本等值", UNO["sil"]["it_tpu_hi"]["SIF_DeployGW_Base"][1], None)]
for i, h in enumerate(["項目", "兩層活頁簿實算（藍字）", "本活頁簿", "差"]): CK.cell(r, i + 1, h).font = BOLD
r += 1
for lab, v, f in UNOROWS:
    CK.cell(r, 1, lab); CK.cell(r, 2, v).font = BLUE
    if f: CK.cell(r, 3, f); CK.cell(r, 4, f"=C{r}-B{r}")
    else: CK.cell(r, 3, "（只與 Python 對過，見報告）")
    r += 1
CK.column_dimensions["A"].width = 44
wb.save(OUT)
print("saved", OUT, "scenarios", len(SC), "rows", len(ORDER))

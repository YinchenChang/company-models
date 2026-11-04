#!/usr/bin/env python3
# 用法：python3 wattonomics/tools/oracle_gap_diag.py build/raw/20261104_Oracle園區與資本支出差異診斷.xlsx；再以 soffice --headless --calc --convert-to xlsx 重算到 wattonomics/docs/verify/
"""產生 20261104_Oracle園區與資本支出差異診斷.xlsx（全部計算為 Excel 公式；LibreOffice 重算）。
不改任何模型；輸入取自 Wattonomics v0.1.9 與 Oracle 個股模型 v0.4.2（數值寫死並標來源）。"""
import sys, itertools
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL

out = sys.argv[1]
wb = Workbook()
BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------- README ----------------
ws = wb.active; ws.title = "README"
readme = [
 "Oracle 2028 園區 GW 與資本支出推得 GW 差異診斷（2026-11-04；診斷用，不改任何模型）",
 "對象：Wattonomics v0.1.9（wattonomics/model/20261102_Wattonomics_v0.1.9.xlsx）Oracle 2028：園區年新增 4.912（Players!I112＝CapexGW!I137）vs 甲2 上線年新增 1.237（CapexGW!I136）。",
 "同口徑比較：年新增對年新增（4.912 vs 1.237）、年底存量對年底存量（Players!I106 8.605 vs CapexGW!I139 4.873）。",
 "頁面：In（輸入，藍字，含標記與來源）→ Site（園區上線時點與分期）→ Calc（甲2 公式逐變體重算：23 組）→ Bridge（年新增與年底存量的橋接；Shapley 分攤，與順序無關）→ Sens（敏感度）→ Check（重現 Wattonomics 數字的核對）。",
 "甲2 公式與 Wattonomics CapexGW 第 119–139 列相同：曆年資本支出＝財年 × 5/12＋次財年 × 7/12；AI＝× ai_sh；資金＝IT 占比 × 當年＋廠房占比 ×（0.5 前一年＋0.5 前兩年）；÷（每 IT GW 的 IT＋廠房成本）；上線年＝(1−lag/12)·當年＋lag/12·前一年 − inv·(當年 − 前一年)。",
 "四個轉換因子：A＝只承擔部分（承租園區廠房由房東出資，Oracle 只付 IT；L＝承租比例）；S＝AI 占比改用 Oracle 個股模型；C＝每 GW IT 成本改用 Oracle 個股模型世代組合；T＝資本支出到上線時差改用在建工程推得的 7.7 個月。",
 "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高。單位：IT GW（設施 ÷ IF_FacilityGW）；金額 US$B。",
 "重算：soffice --headless --calc --convert-to xlsx；錯誤格須為 0（Check 頁）。",
]
for i, t in enumerate(readme, 1):
    ws.cell(i, 1, t).alignment = WRAP
ws.column_dimensions["A"].width = 160

# ---------------- In ----------------
wi = wb.create_sheet("In")
wi.append(["鍵", "項目", "單位", "基準", "低", "高", "標記", "來源／說明"])
for c in wi[1]: c.font = BOLD; c.fill = HDR
R = {}  # key -> row
def scal(key, item, unit, base, lo=None, hi=None, tag="", src=""):
    wi.append([key, item, unit, base, lo, hi, tag, src])
    r = wi.max_row; R[key] = r
    for col in (4, 5, 6):
        wi.cell(r, col).font = BLUE
def ref(key, col="D"):
    return f"In!${col}${R[key]}"

scal("ai_sh", "Wattonomics 資本支出 AI 占比", "%", 0.85, 0.75, 0.95, "Assumed", "Wattonomics In 第 13 列 ai_sh")
scal("lag_on", "Wattonomics 付款到上線延遲", "月", 4.5, 3, 6, "Assumed", "Wattonomics In 第 101 列")
scal("inv_on", "Wattonomics 未上線庫存 ÷ 年出貨", "%", 0.10, 0.05, 0.20, "Assumed", "Wattonomics In 第 102 列")
scal("w1", "廠房提前入帳權重 w1（前一年）", "—", 0.5, 1.0, 0.0, "Assumed", "Wattonomics CapexGW!E4（w0＝0、w2＝0.5）；低／高＝lag_bldg_lo（0,1,0）／lag_bldg_hi（0,0,1），CapexGW 第 5、6 列")
scal("w2", "廠房提前入帳權重 w2（前兩年）", "—", 0.5, 0.0, 1.0, "Assumed", "Wattonomics CapexGW!F4；低／高同上")
scal("fye", "Oracle 財年結束月", "月", 5, None, None, "Verified", "Wattonomics In 第 95 列 fye_orcl")
scal("ye2025", "Oracle 2025 年底 IT GW（晶片換算）", "GW", 1.203, None, None, "Derived", "Wattonomics Players!F106（In chip_orcl 650.2 MW × 1.85）")
scal("obs", "Oracle 已連網 IT GW（2026-08-31）", "GW", 2.543, None, None, "Derived", "Wattonomics In 第 35 列 obs_orcl：FY26 末 1,835＋Q1 交付 850 ÷ 1.2（Oracle 自述 MW，Interested-party）")
scal("obs_t", "觀測時點（小數年）", "年", 2026 + 8/12, None, None, "Verified", "2026-08-31")
scal("cip", "在建工程（FY27Q1 末）", "US$B", 48.546, None, None, "Interested-party", "Oracle FY27 Q1 10-Q（repo oracle/data/oracle_facts_20261007.json bs.cip.fy27q1）")
scal("ltm_cx", "近四季資本支出（FY26Q2–FY27Q1）", "US$B", 75.66, None, None, "Derived", "同上 cf.capex.ltm_fy27q1（12.033＋18.635＋16.493＋28.499）")
scal("lag_O", "資本支出到上線時差（在建工程 ÷ 近四季資本支出 × 12）", "月", None, 6, 9, "Derived", "公式＝cip ÷ ltm_cx × 12；低／高為判斷區間（在建工程含尚未裝機的設備與少量自有建物）")
wi.cell(R["lag_O"], 4, f"={ref('cip')}/{ref('ltm_cx')}*12")
scal("L", "承租園區比例（廠房由房東出資，Oracle 只付 IT）", "%", 1.0, 0.8, 1.0, "Interested-party／Analogy", "Oracle 財務長 Kehring（2025-12-10 法說，經 The Stack、Fortune 轉述）：資本支出絕大部分為設備，土地、建物、電力以租賃取得；九站開發商皆為第三方（E#7–E#9、E#12–E#14）；Kehring 引文 E#5（E#6 為同一發言的另一篇報導，單一來源）")
scal("byoh", "客戶自帶硬體／供應商融資使 Oracle 每 GW 資本支出下降比例", "%", 0.0, 0.0, 0.10, "Assumed", "Oracle Investor Day 2026-10-28：供應商融資、客戶自帶硬體、預付降低前期現金需求（E#3）；客戶自帶晶片 E#16；預付與客戶自帶硬體合約部分 750 億（E#19）；金額未揭露，基準 0")
scal("R", "園區分期上線期間（完工日前均勻上線）", "月", 18, 12, 24, "Analogy", "Abilene 8 棟約 2025 中至 2026Q4（約 18 個月）；Frontier 首棟 2026 下半年、全站 2028Q4（約 27 個月；E#7）")
scal("s_shk", "Shackelford 容量縮放（Vantage 1.4 GW vs T2 2.0 GW 設施）", "x", 1.4/(2000/1200), 1.0, 1.4/2.0, "Analogy（Vantage，Interested-party）", "DCD：Frontier 全站 1.4 GW（口徑未說明）。基準＝視為 IT；低＝不調整；高＝視為設施")
scal("iff", "IF_FacilityGW（設施 ÷ IT）", "x", 1.2, 1.1, 1.3, "Derived（TK）", "Wattonomics 共用參數")
scal("tgt_gw", "Oracle 目標：FY2028 年底資料中心容量（口徑未說明）", "GW", 10, None, None, "Interested-party", "Investor Day 2026-10-28 Maxson：兩年增約 7 GW、FY28 年底逾 10 GW（E#1、E#2）；只作交叉驗證")
scal("tgt_add", "Oracle 目標：兩年新增容量", "GW", 7, None, None, "Interested-party", "同上")
scal("lease_unc", "尚未起租租賃承諾（未折現）", "US$B", 288, None, None, "Interested-party", "Oracle FY27 Q1 10-Q 附註 6（repo 事實總帳）；FY27Q2–FY2029 起租、15–19 年（E#10、E#11）；只作備忘")
scal("ai_dep", "Oracle Q1 FY27 在役 AI 設備折舊年率", "US$B／年", 9.671, None, None, "Derived", "Oracle 個股模型 Thesis!C26（Q1 折舊 3.2 − 非 AI 0.782）× 4")
scal("life_ai", "AI 設備折舊年限（推回在役毛額用）", "年", 6, 5, 6, "Derived／Assumed", "Oracle 10-K 伺服器耐用年限 6 年；個股模型 In life_ai 低 5")
scal("ai_inserv", "Oracle 在役 AI 設備毛額（FY27Q1＝折舊年率 × 年限）", "US$B", None, None, None, "Derived（含 Assumed 年限）", "公式；個股模型 Thesis!C36＝58.026；低＝年限 5")
wi.cell(R["ai_inserv"], 4, f"={ref('ai_dep')}*{ref('life_ai')}")
wi.cell(R["ai_inserv"], 5, f"={ref('ai_dep')}*{ref('life_ai','E')}")
wi.cell(R["ai_inserv"], 6, f"={ref('ai_dep')}*{ref('life_ai','F')}")
scal("sh_m25", "sh_m FY2025 區間（年度表 D 欄）", "%", None, 0.75, 0.95, "Assumed", "個股模型 FY25 無 AI 拆分；基準＝ai_sh")
scal("sh_m26", "sh_m FY2026 區間（年度表 E 欄）", "%", None, 0.85, 0.95, "Assumed", "個股模型 FY26 無 AI 拆分；基準＝FY27 值")
scal("itm2526", "itm FY2025–26 區間（年度表 D、E 欄）", "US$B／GW", None, 24.0152, 32.2372, "Assumed", "GB200 與 GB300 平均；低＝全 GB200、高＝全 GB300")
scal("mw_fy26", "Oracle 已連網 AI MW-IT（FY2026 年底 2026-05）", "GW", 1.835, None, None, "Derived", "Oracle 個股模型 In mw_fy26（FY26 交付 1.2 GW ÷ 1.2 等）")
# year table
wi.append([]); wi.append(["年度表", "", "", 2025, 2026, 2027, 2028, 2029, "標記", "來源／說明"])
yr_hdr = wi.max_row
for c in wi[yr_hdr]: c.font = BOLD; c.fill = HDR
YC = {2025: "D", 2026: "E", 2027: "F", 2028: "G", 2029: "H"}
def yrow(key, item, vals, tag, src):
    wi.append([key, item, ""] + [vals.get(y) for y in range(2025, 2030)] + [tag, src])
    r = wi.max_row; R[key] = r
    for y in range(2025, 2030):
        wi.cell(r, 4 + y - 2025).font = BLUE
def yref(key, y):
    return f"In!${YC[y]}${R[key]}"
yrow("fy9", "Oracle 財年資本支出（採用：In Y9；FY2025 用 Y2）", {2025: 21.215, 2026: 55.663, 2027: 92.5, 2028: 86.398, 2029: 38.006},
     "Derived（個股模型基準 B）", "Wattonomics CapexGW!F121–J121（In 第 200 列；Oracle 個股模型 v0.4.2 CF!C58–F58）")
yrow("fy2", "Oracle 財年資本支出（In Y2：FY27 指引、FY28–29 共識）", {2025: 21.215, 2026: 55.663, 2027: 92.5, 2028: 100.962, 2029: 84.848},
     "Interested-party", "Wattonomics CapexGW 第 120 列（MarketScreener 共識 2026-10-07）")
yrow("fyai", "Oracle 個股模型 AI 資本支出（財年；FY2025、FY2026 無拆分）", {2027: 87.531, 2028: 81.091, 2029: 32.359},
     "Derived（個股模型基準 B）", "Wattonomics In 第 198 列；Oracle 個股模型 CF!D56–F56")
yrow("sh_m", "AI 占比：Oracle 個股模型（FY27–29＝AI ÷ 合計；FY25 取 ai_sh、FY26 取 FY27 值）", {}, "Derived／Assumed", "FY2025、FY2026 個股模型未拆，Assumed")
r = R["sh_m"]
wi.cell(r, 4, f"={ref('ai_sh')}")
wi.cell(r, 6, f"={yref('fyai',2027)}/{yref('fy9',2027)}")
wi.cell(r, 7, f"={yref('fyai',2028)}/{yref('fy9',2028)}")
wi.cell(r, 8, f"={yref('fyai',2029)}/{yref('fy9',2029)}")
wi.cell(r, 5, f"=F{r}")
yrow("pit", "Wattonomics 每 IT GW IT 資本支出（曆年；含非 NVIDIA 占比）", {2025: 27.424, 2026: 33.065, 2027: 36.595, 2028: 38.802, 2029: 38.778},
     "Derived（TK IF_CapexIT 世代加權）", "Wattonomics CapexGW 第 128 列")
yrow("pfac", "Wattonomics 每 IT GW 廠房資本支出（曆年）", {2025: 12.217, 2026: 12.217, 2027: 12.151, 2028: 12.158, 2029: 12.153},
     "Derived（TK IF_CapexFacility）", "Wattonomics CapexGW 第 129 列")
yrow("itm", "Oracle 個股模型每 MW IT 資本支出（財年新增產能組合；US$B／GW）", {2025: (24.0152+32.2372)/2, 2026: (24.0152+32.2372)/2, 2027: 32.2372, 2028: 0.5*32.2372+0.5*37.5864, 2029: 37.5864},
     "Derived（TK v21.4）／Assumed（FY25–26）", "Oracle 個股模型 In 第 38–40 列（GB200 24.02、GB300 32.24、VR200 37.59）× In 第 26–28 列新增組合；FY25–26 取 GB200／GB300 平均 [Assumed]")
wi.cell(R["sh_m25"], 4, f"={yref('sh_m',2025)}"); wi.cell(R["sh_m26"], 4, f"={yref('sh_m',2026)}"); wi.cell(R["itm2526"], 4, f"={yref('itm',2025)}")
yrow("chk_on", "Wattonomics 甲2 上線年新增（核對用）", {2025: 0.465, 2026: 1.086, 2027: 1.348, 2028: 1.237}, "Derived", "Wattonomics CapexGW!F136–I136")
yrow("chk_ye", "Wattonomics 主路徑年底 IT GW（核對用）", {2025: 1.203, 2026: 2.289, 2027: 3.637, 2028: 4.873}, "Derived", "Wattonomics CapexGW!F139–I139")
yrow("chk_site", "Wattonomics 園區年底 IT GW（核對用）", {2025: 1.203, 2026: 2.749, 2027: 3.693, 2028: 8.605}, "Derived", "Wattonomics Players!F106–I106")
yrow("mdl_ye", "Oracle 個股模型已連網 MW-IT（財年末，基準 B；GW）", {2027: 4.66775, 2028: 7.50075, 2029: 8.778}, "Derived（個股模型）", "Oracle 個股模型 Thesis!E41–G41（FY27 末 2027-05、FY28 末 2028-05、FY29 末 2029-05；上限 cap_B 8,778）")
# parks
wi.append([]); wi.append(["園區鍵", "名稱", "口徑", "MW", "上線（小數年；月底）", "Shackelford", "使用者", "擁有者／開發商（房東）", "Wattonomics 列", "來源"])
pk_hdr = wi.max_row
for c in wi[pk_hdr]: c.font = BOLD; c.fill = HDR
parks = [
 ("cp_orcl_abl", "Abilene 餘 2 棟（824 − 618）", "未說明", 206, 2027.0, 0, "OpenAI", "Crusoe（Blue Owl、Primary Digital 合資）", "Pipeline 18／In 270", "A#69；Oracle 法說 6 of 8 棟"),
 ("cp_orcl_shk1", "Shackelford（Frontier）首棟", "設施", 600, 2027 + 11/12, 1, "OpenAI", "Vantage Data Centers（銀行團 380 億債務含本站）", "Pipeline 19／In 271", "J#26、E#7、E#8"),
 ("cp_orcl_shk2", "Shackelford 其餘 9 棟", "設施", 1400, 2028 + 11/12, 1, "OpenAI", "Vantage Data Centers", "Pipeline 20／In 272", "A#72、B#10、E#7、E#8"),
 ("cp_orcl_jup", "Doña Ana（Project Jupiter）", "設施", 2200, 2029.0, 0, "OpenAI", "STACK／BorderPlex（開發商為 Blue Owl 旗下單位）", "Pipeline 21／In 273", "A#73、E#12、E#13"),
 ("cp_orcl_pw", "Port Washington（Lighthouse）", "設施", 1300, 2029.0, 0, "OpenAI", "Vantage（土地 Cloverleaf）", "Pipeline 22／In 274", "A#74、E#8、E#9"),
 ("cp_orcl_sal", "Saline Township（The Barn）其餘", "設施", 922, 2029.0, 0, "OpenAI", "Related Digital（Blackstone 股權、PIMCO 債務）", "Pipeline 23／In 275", "A#75、E#6"),
 ("cp_orcl_sal_b1", "Saline Township 第 1 棟", "設施", 461, 2028.0, 0, "OpenAI", "Related Digital", "Pipeline 94／In 346", "J#20"),
 ("cp_orcl_redoak", "DataBank Red Oak DFW10", "IT", 60, 2027 + 1/12, 0, "未揭露", "DataBank", "Pipeline 95／In 347", "J#21（二手）"),
 ("cp_orcl_redoak2", "DataBank Red Oak DFW11", "IT", 60, 2028 + 1/12, 0, "未揭露", "DataBank", "Pipeline 96／In 348", "J#21b（二手）"),
]
PK0 = wi.max_row + 1
for p in parks:
    wi.append(list(p))
    r = wi.max_row
    wi.cell(r, 4).font = BLUE; wi.cell(r, 5).font = BLUE
PK1 = wi.max_row
wi.append(["", "註：上線欄＝Pipeline t 基準（月底小數年；12 月＝次年 1 月 0 日）。「未說明」口徑依 Wattonomics 基準讀為 IT。2028 年新增＝Shackelford 其餘、Jupiter、Port Washington、Saline 其餘（11–12 月）＋DFW11（1 月）。"])
for col, w in zip("ABCDEFGHIJ", (16, 46, 10, 10, 14, 12, 10, 34, 18, 60)):
    wi.column_dimensions[col].width = w

# ---------------- Site ----------------
wsS = wb.create_sheet("Site")
wsS["A1"] = "Site — 園區 IT GW 與分期上線（每變體四欄：到 2027.0、2028.0、2029.0、2028-05 底已上線 IT GW）。分期＝完工日前 R 個月均勻上線，不早於觀測時點。"
site_vars = [  # name, R expr, s_shk expr, iff expr
 ("S0 重現（R＝0、不縮放、1.2）", "0", "1", "1.2"),
 ("S1 Shackelford 縮放（R＝0）", "0", ref("s_shk"), ref("iff")),
 ("S2 分期 R 基準", ref("R"), ref("s_shk"), ref("iff")),
 ("S3 分期 R 低", ref("R", "E"), ref("s_shk"), ref("iff")),
 ("S4 分期 R 高", ref("R", "F"), ref("s_shk"), ref("iff")),
 ("S5 Shackelford 不縮放（R＝0）", "0", ref("s_shk", "E"), ref("iff")),
 ("S6 Shackelford 視為設施（R＝0）", "0", ref("s_shk", "F"), ref("iff")),
 ("S7 IF_FacilityGW 低（R＝0）", "0", ref("s_shk"), ref("iff", "E")),
 ("S8 IF_FacilityGW 高（R＝0）", "0", ref("s_shk"), ref("iff", "F")),
 ("S9 分期 R 基準、Shackelford 不縮放", ref("R"), ref("s_shk", "E"), ref("iff")),
]
wsS.append([])
hdr = ["園區鍵", "名稱"]
for v in site_vars:
    hdr += [v[0] + "｜IT GW", "到 2027.0", "到 2028.0", "到 2029.0", "到 2028-05 底"]
wsS.append(hdr)
for c in wsS[3]: c.font = BOLD; c.fill = HDR; c.alignment = WRAP
S0 = 4
for i, p in enumerate(parks):
    r = S0 + i; ir = PK0 + i
    wsS.cell(r, 1, f"=In!A{ir}"); wsS.cell(r, 2, f"=In!B{ir}")
    for j, (nm, Rx, sx, fx) in enumerate(site_vars):
        c0 = 3 + j * 5
        g = f"In!D{ir}/1000/IF(In!C{ir}=\"設施\",{fx},1)*IF(In!F{ir}=1,{sx},1)"
        wsS.cell(r, c0, f"={g}")
        gcell = f"{CL(c0)}{r}"
        te = f"In!E{ir}"
        ts = f"MAX({te}-({Rx})/12,{ref('obs_t')})"
        for k, t in enumerate((2027.0, 2028.0, 2029.0, 2028 + 5/12)):
            f = f"=IF({t}>={te}-1E-9,{gcell},IF(AND(({Rx})>0,{t}>{ts}),{gcell}*({t}-{ts})/MAX({te}-{ts},1E-9),0))"
            wsS.cell(r, c0 + 1 + k, f)
SE = S0 + len(parks) - 1
# summary rows
sr = SE + 2
wsS.cell(sr, 1, "變體"); wsS.cell(sr, 2, "年底 2026"); wsS.cell(sr, 3, "年底 2027"); wsS.cell(sr, 4, "年底 2028"); wsS.cell(sr, 5, "2028 年新增"); wsS.cell(sr, 6, "2027 年新增"); wsS.cell(sr, 7, "2028-05 底")
for c in wsS[sr]: c.font = BOLD; c.fill = HDR
SITE = {}
for j, v in enumerate(site_vars):
    r = sr + 1 + j
    c0 = 3 + j * 5
    wsS.cell(r, 1, v[0])
    for k in range(3):
        col = CL(c0 + 1 + k)
        wsS.cell(r, 2 + k, f"={ref('obs')}+SUM({col}{S0}:{col}{SE})")
    wsS.cell(r, 5, f"=D{r}-C{r}"); wsS.cell(r, 6, f"=C{r}-B{r}")
    col = CL(c0 + 4); wsS.cell(r, 7, f"={ref('obs')}+SUM({col}{S0}:{col}{SE})")
    SITE[j] = r
wsS.column_dimensions["A"].width = 34; wsS.column_dimensions["B"].width = 30

# ---------------- Calc ----------------
wc = wb.create_sheet("Calc")
wc["A1"] = "Calc — 甲2 公式逐變體重算（欄 C–F＝2025–2028）。旗標：A＝只承擔部分（廠房成本 ×（1−L））；S＝AI 占比用個股模型；C＝IT 成本用個股模型；T＝時差用 lag_O。"
YRS = [2025, 2026, 2027, 2028]
YCOL = {2025: "C", 2026: "D", 2027: "E", 2028: "F"}
blocks = {}
def block(name, L, smode, cmode, lag, byoh, fysrc, aish):
    r0 = wc.max_row + 2
    wc.cell(r0, 1, name).font = BOLD
    wc.cell(r0, 2, "參數：L＝"); wc.cell(r0, 3, f"={L}")
    wc.cell(r0, 4, "lag＝"); wc.cell(r0, 5, f"={lag}")
    wc.cell(r0, 6, "byoh＝"); wc.cell(r0, 7, f"={byoh}")
    Lc, lagc, byc = f"$C${r0}", f"$E${r0}", f"$G${r0}"
    rows = ["曆年資本支出", "曆年 AI 資本支出", "每 GW IT 成本", "每 GW 廠房成本（Oracle 承擔）", "t−1 AI", "t−2 AI",
            "當年 GW 資金", "年新增 IT GW（付款年）", "年新增 IT GW（上線年）", "年底 IT GW"]
    rr = {nm: r0 + 1 + i for i, nm in enumerate(rows)}
    for nm, r in rr.items():
        wc.cell(r, 1, nm)
    fye = ref("fye")
    for y in YRS:
        c = YCOL[y]
        fy_t, fy_n = yref(fysrc, y), yref(fysrc, y + 1)
        wc[f"{c}{rr['曆年資本支出']}"] = f"={fy_t}*{fye}/12+{fy_n}*(12-{fye})/12"
        if smode:
            wc[f"{c}{rr['曆年 AI 資本支出']}"] = f"={fy_t}*{yref('sh_m',y)}*{fye}/12+{fy_n}*{yref('sh_m',y+1)}*(12-{fye})/12"
        else:
            wc[f"{c}{rr['曆年 AI 資本支出']}"] = f"={c}{rr['曆年資本支出']}*{aish}"
        itb = f"({yref('itm',y)}*{fye}/12+{yref('itm',y+1)}*(12-{fye})/12)" if cmode else yref('pit', y)
        wc[f"{c}{rr['每 GW IT 成本']}"] = f"={itb}*(1-{byc})"
        wc[f"{c}{rr['每 GW 廠房成本（Oracle 承擔）']}"] = f"={yref('pfac',y)}*(1-{Lc})"
        ai = f"{c}{rr['曆年 AI 資本支出']}"
        if y == 2025:
            wc[f"{c}{rr['t−1 AI']}"] = f"={ai}"; wc[f"{c}{rr['t−2 AI']}"] = f"={ai}"
        elif y == 2026:
            p = YCOL[y-1]
            wc[f"{c}{rr['t−1 AI']}"] = f"={p}{rr['曆年 AI 資本支出']}"; wc[f"{c}{rr['t−2 AI']}"] = f"={c}{rr['t−1 AI']}"
        else:
            p, pp = YCOL[y-1], YCOL[y-2]
            wc[f"{c}{rr['t−1 AI']}"] = f"={p}{rr['曆年 AI 資本支出']}"; wc[f"{c}{rr['t−2 AI']}"] = f"={pp}{rr['曆年 AI 資本支出']}"
        it = f"{c}{rr['每 GW IT 成本']}"; fc = f"{c}{rr['每 GW 廠房成本（Oracle 承擔）']}"
        wc[f"{c}{rr['當年 GW 資金']}"] = (f"={it}/({it}+{fc})*{ai}+{fc}/({it}+{fc})*({ref('w1')}*{c}{rr['t−1 AI']}+{ref('w2')}*{c}{rr['t−2 AI']})")
        wc[f"{c}{rr['年新增 IT GW（付款年）']}"] = f"={c}{rr['當年 GW 資金']}/({it}+{fc})"
        g = f"{c}{rr['年新增 IT GW（付款年）']}"
        gp = "0" if y == 2025 else f"{YCOL[y-1]}{rr['年新增 IT GW（付款年）']}"
        wc[f"{c}{rr['年新增 IT GW（上線年）']}"] = f"=(1-{lagc}/12)*{g}+{lagc}/12*{gp}-{ref('inv_on')}*({g}-{gp})"
        if y == 2025:
            wc[f"{c}{rr['年底 IT GW']}"] = f"={ref('ye2025')}"
        else:
            wc[f"{c}{rr['年底 IT GW']}"] = f"={YCOL[y-1]}{rr['年底 IT GW']}+{c}{rr['年新增 IT GW（上線年）']}"
    blocks[name] = {"add": f"Calc!$F${rr['年新增 IT GW（上線年）']}", "ye": f"Calc!$F${rr['年底 IT GW']}",
                    "add27": f"Calc!$E${rr['年新增 IT GW（上線年）']}", "ye27": f"Calc!$E${rr['年底 IT GW']}"}
    return name

F = ["A", "S", "C", "T"]
def subset_name(s):
    return "V_" + ("".join(sorted(s)) if s else "base")
for k in range(5):
    for s in itertools.combinations(F, k):
        s = set(s)
        block(subset_name(s),
              L=ref("L") if "A" in s else "0", smode="S" in s, cmode="C" in s,
              lag=ref("lag_O") if "T" in s else ref("lag_on"), byoh="0", fysrc="fy9", aish=ref("ai_sh"))
ALL = subset_name(set(F))
block("X_base_ai_lo", "0", False, False, ref("lag_on"), "0", "fy9", ref("ai_sh", "E"))
block("X_base_ai_hi", "0", False, False, ref("lag_on"), "0", "fy9", ref("ai_sh", "F"))
block("X_base_Y2", "0", False, False, ref("lag_on"), "0", "fy2", ref("ai_sh"))
block("X_all_L_lo", ref("L", "E"), True, True, ref("lag_O"), "0", "fy9", ref("ai_sh"))
block("X_all_lag_lo", ref("L"), True, True, ref("lag_O", "E"), "0", "fy9", ref("ai_sh"))
block("X_all_lag_hi", ref("L"), True, True, ref("lag_O", "F"), "0", "fy9", ref("ai_sh"))
block("X_all_byoh_hi", ref("L"), True, True, ref("lag_O"), ref("byoh", "F"), "fy9", ref("ai_sh"))
block("X_all_lagon_cur", ref("L"), True, True, ref("lag_on"), "0", "fy9", ref("ai_sh"))
wc.column_dimensions["A"].width = 34

# ---------------- Bridge ----------------
wbB = wb.create_sheet("Bridge")
from math import factorial
def shap_formula(f, metric):
    terms = []
    others = [x for x in F if x != f]
    n = len(F)
    for k in range(len(others) + 1):
        for s in itertools.combinations(others, k):
            s = set(s)
            w = factorial(len(s)) * factorial(n - len(s) - 1) / factorial(n)
            terms.append(f"{w:.10f}*({blocks[subset_name(s | {f})][metric]}-{blocks[subset_name(s)][metric]})")
    return "=" + "+".join(terms)
wbB["A1"] = "Bridge — 園區 → 資本支出推得 GW 的橋接（IT GW）。轉換段四因子以 Shapley 平均分攤（24 種順序平均，與順序無關）；低／高＝單獨作用與最後作用的較小／較大值，或輸入低／高。負號＝由園區往資本支出方向扣減。"
hdrB = ["項目", "分類", "2028 年新增：基準", "低", "高", "年底 2028：基準", "低", "高", "說明"]
def bridge_table(r0, title):
    wbB.cell(r0, 1, title).font = BOLD
    for i, h in enumerate(hdrB):
        c = wbB.cell(r0 + 1, i + 1, h); c.font = BOLD; c.fill = HDR
    return r0 + 2
r = bridge_table(3, "一、橋接表")
site_r0, site_r1, site_r2 = SITE[0], SITE[1], SITE[2]
rows_B = {}
def brow(key, item, cat, a_b, a_l, a_h, y_b, y_l, y_h, note):
    global r
    wbB.cell(r, 1, item); wbB.cell(r, 2, cat)
    for i, v in enumerate((a_b, a_l, a_h, y_b, y_l, y_h)):
        if v is not None:
            wbB.cell(r, 3 + i, v)
    wbB.cell(r, 9, note)
    rows_B[key] = r
    r += 1
def S(j, col):  # site summary cell
    return f"Site!${col}${SITE[j]}"
brow("site", "園區（Wattonomics 甲；T2 公告完工日）", "起點", f"={S(0,'E')}", None, None, f"={S(0,'D')}", None, None, "重現 Players!I112（4.912）、I106（8.605）")
brow("c1", "(C1) Shackelford 容量口徑（Vantage 1.4 GW）", "其他：口徑", f"={S(1,'E')}-{S(0,'E')}", f"={S(5,'E')}-{S(0,'E')}", f"={S(6,'E')}-{S(0,'E')}",
     f"={S(1,'D')}-{S(0,'D')}", f"={S(5,'D')}-{S(0,'D')}", f"={S(6,'D')}-{S(0,'D')}", "T2 用 Epoch 2.0 GW 設施；開發商稱 1.4 GW。園區端 C1 與 B1 依序計算（先 C1 後 B1）；若先分期，C1 年新增為 −0.11（見交叉驗證）")
brow("c2", "(C2) 重複計算（Oracle 園區另入 OpenAI／SPV）", "其他：重複", "=0", "=0", "=0", "=0", "=0", "=0", "已查：Oracle 九站只在 orcl；oai 列為 Milam、Lordstown（SoftBank）；4.5 GW 合約只在 S4 對照（GL01-13）")
brow("b1", "(B1) 園區清單把整站放在完工月（分期上線）", "期間錯配：園區時點", f"={S(2,'E')}-{S(1,'E')}", f"={S(4,'E')}-{S(1,'E')}", f"={S(3,'E')}-{S(1,'E')}",
     f"={S(2,'D')}-{S(1,'D')}", f"={S(2,'D')}-{S(1,'D')}", f"={S(2,'D')}-{S(1,'D')}", "2028 年 4.85 GW 落在 11–12 月；R＝12／18／24 個月。年底存量不受影響")
brow("x", "＝分期後園區（同一批站點）", "小計", f"=C{rows_B['site']}+C{rows_B['c1']}+C{rows_B['c2']}+C{rows_B['b1']}", None, None,
     f"=F{rows_B['site']}+F{rows_B['c1']}+F{rows_B['c2']}+F{rows_B['b1']}", None, None, "")
brow("res", "(D) 未解釋殘差（含客戶自帶硬體、供應商融資、起點差異）", "殘差",
     f"={blocks[ALL]['add']}-C{rows_B['x']}", f"={blocks['X_all_byoh_hi']['add']}-C{rows_B['x']}", None,
     f"={blocks[ALL]['ye']}-F{rows_B['x']}", f"={blocks['X_all_byoh_hi']['ye']}-F{rows_B['x']}", None,
     "＝Oracle 一致的資本支出推得 GW − 分期後園區；低欄＝byoh 10% 時")
brow("allc", "＝以 Oracle 一致假設由 Y9 資本支出推得（A＋S＋C＋T 全用）", "小計", f"={blocks[ALL]['add']}", None, None, f"={blocks[ALL]['ye']}", None, None, "")
labels = {"T": ("(B2) 資本支出到上線時差（7.7 → 4.5 月）", "期間錯配：資本支出時差"),
          "C": ("(C3) 每 GW IT 成本（個股模型世代 → Wattonomics）", "其他：成本組合"),
          "S": ("(C4) AI 占比（個股模型約 0.94 → 0.85）", "其他：AI 占比"),
          "A": ("(A1) 廠房由房東出資（只承擔 IT → 全額 IT＋廠房）", "只承擔部分")}
for f in ["T", "C", "S", "A"]:
    others = set(F) - {f}
    single_a = f"({blocks[subset_name({f})]['add']}-{blocks['V_base']['add']})"
    last_a = f"({blocks[ALL]['add']}-{blocks[subset_name(others)]['add']})"
    single_y = f"({blocks[subset_name({f})]['ye']}-{blocks['V_base']['ye']})"
    last_y = f"({blocks[ALL]['ye']}-{blocks[subset_name(others)]['ye']})"
    brow("f" + f, labels[f][0], labels[f][1], "=-(" + shap_formula(f, "add")[1:] + ")", f"=-MAX({single_a},{last_a})", f"=-MIN({single_a},{last_a})",
         "=-(" + shap_formula(f, "ye")[1:] + ")", f"=-MAX({single_y},{last_y})", f"=-MIN({single_y},{last_y})", "Shapley；低／高＝單獨／最後作用")
brow("end", "＝Wattonomics 甲2（CapexGW）", "終點", f"=C{rows_B['allc']}+C{rows_B['fT']}+C{rows_B['fC']}+C{rows_B['fS']}+C{rows_B['fA']}", None, None,
     f"=F{rows_B['allc']}+F{rows_B['fT']}+F{rows_B['fC']}+F{rows_B['fS']}+F{rows_B['fA']}", None, None, "重現 CapexGW!I136（1.237）、I139（4.873）")
brow("gap", "總差距（園區 − 甲2）", "核對", f"=C{rows_B['site']}-C{rows_B['end']}", None, None, f"=F{rows_B['site']}-F{rows_B['end']}", None, None, "")
brow("sum", "各項合計（應＝總差距）", "核對",
     f"=-(C{rows_B['c1']}+C{rows_B['c2']}+C{rows_B['b1']}+C{rows_B['res']}+C{rows_B['fT']}+C{rows_B['fC']}+C{rows_B['fS']}+C{rows_B['fA']})", None, None,
     f"=-(F{rows_B['c1']}+F{rows_B['c2']}+F{rows_B['b1']}+F{rows_B['res']}+F{rows_B['fT']}+F{rows_B['fC']}+F{rows_B['fS']}+F{rows_B['fA']})", None, None, "")
# share by category
r += 1
wbB.cell(r, 1, "二、依 Andy 的三類歸併（占總差距）").font = BOLD; r += 1
for i, h in enumerate(["類別", "", "2028 年新增（GW）", "占比", "", "年底 2028（GW）", "占比", "", "包含"]):
    c = wbB.cell(r, i + 1, h); c.font = BOLD; c.fill = HDR
r += 1
cats = [("只承擔部分（A）", ["fA"], "A1"),
        ("期間錯配：園區清單完工月（B1）", ["b1"], "B1"),
        ("期間錯配：資本支出到上線時差（B2）", ["fT"], "B2"),
        ("其他：AI 占比、成本組合、口徑、重複（C）", ["fS", "fC", "c1", "c2"], "C1–C4"),
        ("未解釋殘差（D）", ["res"], "D")]
CAT = {}
for nm, keys, inc in cats:
    wbB.cell(r, 1, nm)
    wbB.cell(r, 3, "=-(" + "+".join(f"C{rows_B[k]}" for k in keys) + ")")
    wbB.cell(r, 4, f"=C{r}/C{rows_B['gap']}")
    wbB.cell(r, 6, "=-(" + "+".join(f"F{rows_B[k]}" for k in keys) + ")")
    wbB.cell(r, 7, f"=F{r}/F{rows_B['gap']}")
    wbB.cell(r, 9, inc)
    CAT[nm] = r; r += 1
# cross-checks
r += 1
wbB.cell(r, 1, "三、交叉驗證（獨立於 Wattonomics 公式）").font = BOLD; r += 1
xc = [
 ("Oracle 目標隱含每 IT GW 資本支出：FY27＋FY28 Y9 ÷（7 GW ÷ IF_FacilityGW）", f"=({yref('fy9',2027)}+{yref('fy9',2028)})/({ref('tgt_add')}/{ref('iff')})", "US$B／IT GW", "Derived；偏誤：(a) 分子為全部資本支出；(b) 資本支出與上線有時差（部分容量用 FY26 的錢、部分錢 FY29 才上線）；(c) 7 GW 可能含客戶自帶硬體；7 GW 口徑未說明、只有二手轉述"),
 ("同上，分子改用 AI 資本支出（個股模型 FY27＋FY28）", f"=({yref('fyai',2027)}+{yref('fyai',2028)})/({ref('tgt_add')}/{ref('iff')})", "US$B／IT GW", "Derived；偏誤 (a) 修正後"),
 ("對照：未分期園區清單 FY27–FY28（2026-06 至 2028-05）新增 IT GW", f"=Site!$G${SITE[0]}-{ref('mw_fy26')}", "GW", "Derived；遠低於 7 ÷ 1.2＝5.8，只有接受分期（B1）才大致相容"),
 ("對照：分期後園區（R 18）FY27–FY28 新增 IT GW", f"=Site!$G${SITE[2]}-{ref('mw_fy26')}", "GW", "Derived"),
 ("同上，若 7 GW 為 IT 口徑", f"=({yref('fy9',2027)}+{yref('fy9',2028)})/{ref('tgt_add')}", "US$B／IT GW", "Derived"),
 ("對照：Wattonomics 全額（IT＋廠房）2028", f"={yref('pit',2028)}+{yref('pfac',2028)}", "US$B／IT GW", "CapexGW!I128＋I129"),
 ("對照：Wattonomics 只 IT 2028", f"={yref('pit',2028)}", "US$B／IT GW", "CapexGW!I128"),
 ("Oracle 在役 AI 設備毛額 ÷ 已連網 IT GW（FY27Q1）", f"={ref('ai_inserv')}/{ref('obs')}", "US$B／IT GW", "Derived；低於 TK 只 IT 成本 → 不含廠房，且可能含客戶自帶硬體或舊世代"),
 ("Oracle 個股模型 2028 年底 IT GW（FY29 內依速度 2,833／年至上限）", f"=MIN(In!$H${R['mdl_ye']},In!$G${R['mdl_ye']}+2.833*7/12)", "GW", "Derived；個股模型 MW＝MIN(cap_B 已簽站點上限, mw_q1＋pace_B × 期間)（Thesis!D41–G41；build_oracle.py:258），由交付速度與已簽站點決定，不是由資本支出推得 → 與園區 8.61 一致近於循環，不是獨立檢查"),
 ("差距：Oracle 一致資本支出推得 2028 年底 − 個股模型", f"={blocks[ALL]['ye']}-MIN(In!$H${R['mdl_ye']},In!$G${R['mdl_ye']}+2.833*7/12)", "GW", "未調和；可能原因：客戶自帶硬體、供應商融資、個股模型的在建工程起點、交付速度假設"),
 ("Oracle 個股模型 2027 年底 IT GW（FY28 內線性）", f"=In!$F${R['mdl_ye']}+(In!$G${R['mdl_ye']}-In!$F${R['mdl_ye']})*7/12", "GW", "Derived"),
 ("Oracle 目標 FY28 年底（2028-05）IT GW（10 GW ÷ 1.2）", f"={ref('tgt_gw')}/{ref('iff')}", "GW", "Interested-party；口徑未說明，可能含非 AI 舊區域；二手"),
 ("分期後園區 2028-05 底（R 基準）", f"=Site!$G${SITE[2]}", "GW", "Derived"),
 ("差距：Oracle 目標 − 分期後園區（2028-05）", f"={ref('tgt_gw')}/{ref('iff')}-Site!$G${SITE[2]}", "GW", "未調和；可能原因：10 GW 含非 AI 舊區域或設施口徑不同、清單漏列站點、目標偏樂觀"),
 ("分期後園區 2027 年底（R 基準）", f"={S(2,'C')}", "GW", "對照 Wattonomics 園區 3.693"),
 ("Oracle 一致資本支出推得 2027 年底", f"={blocks[ALL]['ye27']}", "GW", "對照 甲2 3.637"),
 ("個股模型 2027 年底（FY28 內線性）− Oracle 一致資本支出推得", f"=In!$F${R['mdl_ye']}+(In!$G${R['mdl_ye']}-In!$F${R['mdl_ye']})*7/12-{blocks[ALL]['ye27']}", "GW", "未調和（個股模型 6.32 vs 5.24）"),
 ("未起租租賃承諾 ÷ 平均期限 17 年（年租金量級）", f"={ref('lease_unc')}/17", "US$B／年", "備忘：房東出資的廠房以租金回收，不在 Oracle 資本支出"),
]
xc.append(("順序：先分期再縮放 Shackelford 時 C1 的年新增", f"={S(2,'E')}-{S(9,'E')}", "GW", "Derived；C1 依順序 −0.11 至 −0.19，B1 相應變動"))
XC = {}
for nm, f, u, note in xc:
    wbB.cell(r, 1, nm); wbB.cell(r, 3, f); wbB.cell(r, 4, u); wbB.cell(r, 9, note)
    XC[nm] = r; r += 1
r += 1
wbB.cell(r, 1, "四、單獨套用效果（給建議用；與 Shapley 分攤值不同）").font = BOLD; r += 1
for i, h in enumerate(["做法", "", "2028 年新增", "", "", "年底 2028", "", "", "說明"]):
    c = wbB.cell(r, i + 1, h); c.font = BOLD; c.fill = HDR
r += 1
SA = {}
for nm, key, note in [("只套 A（Oracle 只除 IT 成本）", "V_A", "建議 1 單獨執行"), ("只套 S（AI 占比改用個股模型）", "V_S", "建議 2 單獨執行"),
                      ("A＋S 同時", "V_AS", "建議 1＋2；甲2 年底"), ]:
    wbB.cell(r, 1, nm)
    wbB.cell(r, 3, f"={blocks[key]['add']}-{blocks['V_base']['add']}"); wbB.cell(r, 6, f"={blocks[key]['ye']}-{blocks['V_base']['ye']}")
    wbB.cell(r, 9, note); SA[key] = r; r += 1
wbB.cell(r, 1, "A＋S 同時：甲2 年底值"); wbB.cell(r, 6, f"={blocks['V_AS']['ye']}"); r += 1
for col, w in zip("ABCDEFGHI", (58, 26, 14, 10, 10, 14, 10, 10, 70)):
    wbB.column_dimensions[col].width = w

# ---------------- Sens ----------------
wsn = wb.create_sheet("Sens")
wsn["A1"] = "Sens — 敏感度（先看敏感度再看基準）。每列只動一個輸入；欄位＝受影響的量。"
wsn.append([])
wsn.append(["輸入", "設定", "受影響的量", "2028 年新增", "年底 2028", "相對基準（年新增）", "相對基準（年底）", "說明"])
for c in wsn[3]: c.font = BOLD; c.fill = HDR
def sline(inp, setting, what, add, ye, base_add, base_ye, note=""):
    wsn.append([inp, setting, what, add, ye, f"=D{wsn.max_row+1}-{base_add}", f"=E{wsn.max_row+1}-{base_ye}", note])
b_add, b_ye = blocks["V_base"]["add"], blocks["V_base"]["ye"]
a_add, a_ye = blocks[ALL]["add"], blocks[ALL]["ye"]
sline("基準", "Wattonomics 甲2", "甲2", f"={b_add}", f"={b_ye}", b_add, b_ye)
sline("ai_sh", "低 0.75", "甲2", f"={blocks['X_base_ai_lo']['add']}", f"={blocks['X_base_ai_lo']['ye']}", b_add, b_ye)
sline("ai_sh", "高 0.95", "甲2", f"={blocks['X_base_ai_hi']['add']}", f"={blocks['X_base_ai_hi']['ye']}", b_add, b_ye)
sline("資本支出來源", "In Y2（FY28–29 共識 101.0／84.8）", "甲2", f"={blocks['X_base_Y2']['add']}", f"={blocks['X_base_Y2']['ye']}", b_add, b_ye, "Y9 FY29 38.0 低於共識，但與 Investor Day『FY29 下降』方向一致")
sline("基準", "Oracle 一致（A＋S＋C＋T）", "資本支出推得", f"={a_add}", f"={a_ye}", a_add, a_ye)
sline("L", "0.8", "資本支出推得", f"={blocks['X_all_L_lo']['add']}", f"={blocks['X_all_L_lo']['ye']}", a_add, a_ye)
sline("lag_O", "6 月", "資本支出推得", f"={blocks['X_all_lag_lo']['add']}", f"={blocks['X_all_lag_lo']['ye']}", a_add, a_ye)
sline("lag_O", "9 月", "資本支出推得", f"={blocks['X_all_lag_hi']['add']}", f"={blocks['X_all_lag_hi']['ye']}", a_add, a_ye)
sline("lag_O", "維持 4.5 月", "資本支出推得", f"={blocks['X_all_lagon_cur']['add']}", f"={blocks['X_all_lagon_cur']['ye']}", a_add, a_ye)
sline("byoh", "10%", "資本支出推得", f"={blocks['X_all_byoh_hi']['add']}", f"={blocks['X_all_byoh_hi']['ye']}", a_add, a_ye)
s_add, s_ye = S(2, 'E'), S(2, 'D')
sline("基準", "分期後園區（R 18、Shackelford 縮放）", "園區", f"={s_add}", f"={s_ye}", s_add, s_ye)
sline("R", "12 月", "園區", f"={S(3,'E')}", f"={S(3,'D')}", s_add, s_ye)
sline("R", "24 月", "園區", f"={S(4,'E')}", f"={S(4,'D')}", s_add, s_ye)
sline("s_shk", "不縮放（2.0 GW 設施）", "園區", f"={S(9,'E')}", f"={S(9,'D')}", s_add, s_ye)
sline("IF_FacilityGW", "1.1（R＝0）", "園區", f"={S(7,'E')}", f"={S(7,'D')}", S(1,'E'), S(1,'D'), "相對 R＝0 縮放後")
sline("IF_FacilityGW", "1.3（R＝0）", "園區", f"={S(8,'E')}", f"={S(8,'D')}", S(1,'E'), S(1,'D'), "相對 R＝0 縮放後")
for col, w in zip("ABCDEFGH", (18, 36, 14, 12, 12, 16, 16, 60)):
    wsn.column_dimensions[col].width = w

# ---------------- Check ----------------
wk = wb.create_sheet("Check")
wk.append(["核對項", "模型值", "本檔重算", "差", "通過（|差| < 0.002）"])
for c in wk[1]: c.font = BOLD; c.fill = HDR
def chk(nm, a, b):
    rr = wk.max_row + 1
    wk.append([nm, a, b, f"=C{rr}-B{rr}", f"=IF(ABS(D{rr})<0.002,1,0)"])
for y, c in YCOL.items():
    chk(f"甲2 上線年新增 {y}", f"={yref('chk_on',y)}", f"=Calc!{c}{int(blocks['V_base']['add'].split('$')[-1])}")
for y, c in YCOL.items():
    chk(f"甲2 年底 {y}", f"={yref('chk_ye',y)}", f"=Calc!{c}{int(blocks['V_base']['ye'].split('$')[-1])}")
chk("園區年底 2026", f"={yref('chk_site',2026)}", f"={S(0,'B')}")
chk("園區年底 2027", f"={yref('chk_site',2027)}", f"={S(0,'C')}")
chk("園區年底 2028", f"={yref('chk_site',2028)}", f"={S(0,'D')}")
chk("橋接合計＝總差距（年新增）", f"=Bridge!C{rows_B['gap']}", f"=Bridge!C{rows_B['sum']}")
chk("橋接合計＝總差距（年底）", f"=Bridge!F{rows_B['gap']}", f"=Bridge!F{rows_B['sum']}")
chk("Shapley 四因子合計＝Oracle 一致 − 基準（年新增）", f"={a_add}-{b_add}", f"=-(Bridge!C{rows_B['fA']}+Bridge!C{rows_B['fS']}+Bridge!C{rows_B['fC']}+Bridge!C{rows_B['fT']})")
rr = wk.max_row + 2
wk.cell(rr, 1, "錯誤格數（全活頁簿；由核對程式另計）"); wk.cell(rr + 1, 1, "全部通過"); wk.cell(rr + 1, 2, f"=IF(SUM(E2:E{rr-2})={rr-3},1,0)")
wk.column_dimensions["A"].width = 46

wb.save(out)
import json
# map 檔只供除錯
json.dump({"rows_B": rows_B, "CAT": CAT, "XC": XC, "SITE": SITE, "blocks": blocks}, open(out + ".map.json", "w"), ensure_ascii=False, indent=1)
print("saved", out)

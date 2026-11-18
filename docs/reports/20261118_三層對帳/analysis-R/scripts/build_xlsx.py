# 建 R1R2 活頁簿：輸入頁（值＋來源＋等級）＋計算頁（公式連到輸入頁），之後以 LibreOffice 重算
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL

D = json.load(open("/home/claude/out/analysis-R/scripts/inputs.json"))
R = json.load(open("/home/claude/out/analysis-R/scripts/results.json"))
S, W, T, B, E, C = D["sil"], D["wat"], D["tk"], D["br"], D["ev"], D["cb"]
OUT = "/home/claude/out/analysis-R/20261118_R1R2_平台別差距拆解.xlsx"
YRS = [2026, 2027, 2028, 2029]
YC = {2026: 4, 2027: 5, 2028: 6, 2029: 7}   # D..G
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); GRN = Font(color="008000")
HDR = PatternFill("solid", fgColor="DDEBF7"); KEY = PatternFill("solid", fgColor="FFF2CC"); BAD = PatternFill("solid", fgColor="F8CBAD")
wb = openpyxl.Workbook()
REF = {}

SILF = "siliconomics/model/20261116_Siliconomics_v0.20.xlsx"
SILJ = "siliconomics/data/siliconomics_snapshot_v0.20.json"
WATF = "wattonomics/model/20261116_Wattonomics_v0.3.6.xlsx"
TKF = "Tokenomics model/20261105_Tokenomics_v22.0.xlsx"


def header(ws, title, cols):
    ws["A1"] = title; ws["A1"].font = BOLD
    for j, h in enumerate(cols, 1):
        c = ws.cell(3, j, h); c.font = BOLD; c.fill = HDR


# ======================= 說明 =======================
rd = wb.active; rd.title = "說明"
lines = [
    "R1／R2 平台別差距拆解（analysis-R，2026-11-18）— 唯讀分析，未改任何模型",
    "資料來源：company-models main d04f591（Siliconomics v0.20、Wattonomics v0.3.6、協調橋接 v0.3.6、markets/chip_bridge v0.2.3 dist）；Tokenomics 79f6b30（v22.0）",
    "單位：GW＝IT 關鍵電力（Tokenomics 口徑），出貨年。實體 GW＝顆數 ÷ 每 GW 顆數（NVIDIA）或 金額 ÷（VR200 每 GW IT capex × 持有比）（非 NVIDIA，兩層同一慣例）。成本等值 GW＝IT 資本支出 ÷ 當年 NVIDIA 每 GW IT capex（Siliconomics Deploy 頁口徑）",
    "頁面：In_SIL／In_WAT／In_CG／In_TK＝輸入（藍字＝自模型或快照抄入的值，H 欄等級、I 欄來源名稱）；R1_SIL、R1_WAT、R1_並排、R1_歸因、R2_顆數＝公式頁（逐格可點回輸入頁）；自我核對＝公式結果與 Python 獨立重算之差",
    "等級：Verified／Interested-party／Analogy／Assumed／Derived（同 COMMON.md 第 7 條）。本檔所有計算結果皆為 [Derived]",
    "Wattonomics 的 NVIDIA／非 NVIDIA 拆分只存在於 Compare 頁 F 節（2025–2027）；2028–2029 模型未拆，本檔不以比例硬拆（列為結構限制）",
    "非 NVIDIA 的平台別（AMD／TPU／Trainium／其他）在 Wattonomics 是依各買方換算份額時所用的持有比名稱歸屬（Maia→其他、TPU→TPU、Trn→Trainium、AMD→AMD、OAI→其他、Ascend→其他）；這是模型隱含歸屬，不是模型輸出的平台拆分",
    "樣本外：NVIDIA FY2027 Q3（107.6）未進入本檔任何計算或參數",
]
for i, t in enumerate(lines, 1):
    rd.cell(i, 1, t)
rd.column_dimensions["A"].width = 160

# ======================= In_SIL =======================
ws = wb.create_sheet("In_SIL")
header(ws, "In_SIL — Siliconomics v0.20 輸入（藍字）", ["鍵", "項目", "單位", "2026", "2027", "2028", "2029", "等級", "來源名稱"])
r = 4


def inrow(ws, key, label, unit, vals, tag, src, scalar=False):
    global r
    ws.cell(r, 1, key); ws.cell(r, 2, label); ws.cell(r, 3, unit)
    if scalar:
        c = ws.cell(r, 4, vals); c.font = BLUE
        REF[key] = f"{ws.title}!$D${r}"
    else:
        for y, v in zip(YRS, vals):
            if v is None:
                continue
            c = ws.cell(r, YC[y], v); c.font = BLUE
            REF[(key, y)] = f"{ws.title}!${CL(YC[y])}${r}"
    ws.cell(r, 8, tag); ws.cell(r, 9, src)
    r += 1


ug = {g["key"]: g for g in json.load(open("/home/claude/company-models/siliconomics/data/siliconomics_snapshot_v0.20.json"))["unit_groups"]}
for p in ["bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth"]:
    inrow(ws, f"u_{p}", f"原始顆數：{ug[p]['name']}", "百萬顆（封裝）", S[f"u_{p}"], ug[p]["tag"], f"{SILJ} SIF_UnitsPlat_{p}（{ug[p]['src']}）")
inrow(ws, "coef", "基準顆數係數 crd_k（S33：2028 0.8；2029 依連續法）", "x", S["coef"], "Assumed（2028）／Derived（2029）", f"{SILJ} SIF_UnitsCoef_Base；{SILF} Coord!C37:F37")
inrow(ws, "denom", "成本等值分母：當年 NVIDIA 每 GW IT capex", "$B/GW", S["deploy_denom"], "Derived", f"{SILF} Deploy!C6:F6（2026＝GB300×(1−vr26)＋VR200×vr26；2027 起 VR200）")
for n, lab, tag in [("it_GB", "每顆 IT capex：Blackwell（以 GB300 NVL72 計）", "Derived"), ("it_VR", "每顆 IT capex：Rubin（VR200）", "Derived"), ("it_RU", "每顆 IT capex：Rubin Ultra", "Derived"),
                    ("it_HP", "每顆 IT capex：Hopper", "Derived"), ("it_AMD8", "每顆 IT capex：AMD MI355（2026 用）", "Assumed"), ("it_AMDR", "每顆 IT capex：AMD Helios（2027 起用）", "Assumed"),
                    ("it_TPU", "每顆 IT capex：TPU", "Assumed"), ("it_TRN", "每顆 IT capex：Trainium", "Assumed"), ("it_ASIC", "每顆 IT capex：其他 ASIC", "Assumed")]:
    lohi = S.get(n + "_lohi")
    extra = f"；低／高 {lohi[0]}／{lohi[1]}" if lohi and lohi[0] is not None else ""
    inrow(ws, n, lab, "$K/顆", S[n], tag, f"{SILF} In 頁 {n}{extra}", scalar=True)
inrow(ws, "sil_phys", "SIF_DeployGWPhys_Base（實體部署 GW）", "GW", S["phys"], "Derived", f"{SILJ} SIF_DeployGWPhys_Base")
inrow(ws, "sil_ceq", "SIF_DeployGW_Base（成本等值部署 GW）", "GW", S["deploy"], "Derived", f"{SILJ} SIF_DeployGW_Base")
inrow(ws, "sil_nvphys_m", "Deploy 頁 實體 GW：NVIDIA（模型值）", "GW", S["deploy_nv_phys"], "Derived", f"{SILF} Deploy!C8:F8")
inrow(ws, "sil_nnvphys_m", "Deploy 頁 實體 GW：非 NVIDIA（模型值）", "GW", S["deploy_nnv_phys"], "Derived", f"{SILF} Deploy!C9:F9")
for c_, w_ in zip("ABCI", [14, 52, 14, 90]):
    ws.column_dimensions[c_].width = w_

# ======================= In_TK =======================
ws = wb.create_sheet("In_TK")
header(ws, "In_TK — Tokenomics v22.0 與三處鏡像（藍字）；欄 D＝Tokenomics Interface 基準、E＝Siliconomics TK_In、F＝Wattonomics TK_In、G＝chip_bridge dist（v21.7 快照）", ["鍵", "項目", "單位", "TK Interface", "SIL TK_In", "WAT TK_In", "chip_bridge", "等級", "來源名稱"])
r = 4
GN = {"H100": "Hopper H100", "GB200": "GB200 NVL72", "GB300": "GB300 NVL72", "VR200": "VR200 NVL72", "RU": "Rubin Ultra（MGX 單架 72 GPU）"}
for g in ["H100", "GB200", "GB300", "VR200", "RU"]:
    cb = {"GB300": C["gpus"][1], "VR200": C["gpus"][4]}.get(g)
    ws.cell(r, 1, f"gpg_{g}"); ws.cell(r, 2, f"每 GW GPU 封裝數：{GN[g]}"); ws.cell(r, 3, "顆/GW")
    for j, v in zip([4, 5, 6, 7], [T["GPUsPerGW"][g][1], S["tk_in"][f"TK_GPUsPerGW_{g}"][1], W["tk_in"][f"TK_GPUsPerGW_{g}"][1], cb]):
        if v is not None:
            ws.cell(r, j, v).font = BLUE
            REF[(f"gpg_{g}", j)] = f"In_TK!${CL(j)}${r}"
    ws.cell(r, 8, "Derived（TK：機架設計功率與封裝數推得）"); ws.cell(r, 9, f"{TKF} Interface 第 7 列（IF_GPUsPerGW）；{SILF} TK_In；{WATF} TK_In；markets/chip_bridge dist In!gpus")
    r += 1
for g in ["H100", "GB200", "GB300", "VR200", "RU"]:
    ws.cell(r, 1, f"cit_{g}"); ws.cell(r, 2, f"每 GW IT capex：{GN[g]}"); ws.cell(r, 3, "$B/GW")
    cb = {"GB300": C["capit"][1], "VR200": C["capit"][4]}.get(g)
    for j, v in zip([4, 5, 6, 7], [T["CapexIT"][g][1], S["tk_in"][f"TK_CapexIT_{g}"][1], W["tk_in"][f"TK_CapexIT_{g}"][1], cb]):
        if v is not None:
            ws.cell(r, j, v).font = BLUE
            REF[(f"cit_{g}", j)] = f"In_TK!${CL(j)}${r}"
    ws.cell(r, 8, "Derived"); ws.cell(r, 9, f"{TKF} Interface 第 9 列（IF_CapexIT）及三處鏡像")
    r += 1
for h in ["AMD", "TPU", "Trn", "OAI", "Ascend", "Maia"]:
    ws.cell(r, 1, f"hold_{h}"); ws.cell(r, 2, f"非 NVIDIA 持有成本比（÷ VR200）：{h}"); ws.cell(r, 3, "x")
    for j, v in zip([4, 5, 6], [T["hold"][h][1], S["tk_in"][f"TK_NonNV_Hold_{h}"][1], W["tk_in"][f"TK_NonNV_Hold_{h}"][1]]):
        ws.cell(r, j, v).font = BLUE
        REF[(f"hold_{h}", j)] = f"In_TK!${CL(j)}${r}"
    ws.cell(r, 8, "Assumed（低／基準／高 " + "／".join(str(x) for x in T["hold"][h]) + "）"); ws.cell(r, 9, f"{TKF} NonNV 第 5–11 列 F 欄；Interface P 節")
    r += 1
r += 1
ws.cell(r, 1, "TK 機架規格（R2 重算用）").font = BOLD; r += 1
for key, lab, unit, val, tag, src in [("it_mw", "IT 電力規模", "GW", T["it_gw"], "Decision", f"{TKF} Inputs!E5"),
                                      ("offrack", "機架外 IT 電力占比", "%", T["offrack"], "Assumed（5–12%）", f"{TKF} Inputs!E8"),
                                      ("pue", "PUE（只用於設施電力，不進每 GW 顆數）", "x", T["pue"][1], "Analogy（1.1／1.2／1.3）", f"{TKF} Inputs!E9"),
                                      ("util", "平均用電 ÷ 配電設計功率（只用於能耗，不進每 GW 顆數）", "%", T["util"][1], "Assumed（0.7／0.8／0.9）", f"{TKF} Inputs!E10"),
                                      ("power_case", "功率情境（1 低／2 基準／3 高）", "選擇", T["power_case"], "Decision", f"{TKF} Inputs!E6")]:
    ws.cell(r, 1, key); ws.cell(r, 2, lab); ws.cell(r, 3, unit); ws.cell(r, 4, val).font = BLUE; ws.cell(r, 8, tag); ws.cell(r, 9, src)
    REF[key] = f"In_TK!$D${r}"; r += 1
for g in ["H100", "GB200", "GB300", "VR200", "RU"]:
    ws.cell(r, 1, f"kw_{g}"); ws.cell(r, 2, f"機架配電設計功率 低／基準／高：{GN[g]}（D／E／F 欄）"); ws.cell(r, 3, "kW")
    for j, v in zip([4, 5, 6], T["rack_kw"][g]):
        ws.cell(r, j, v).font = BLUE
        REF[(f"kw_{g}", j)] = f"In_TK!${CL(j)}${r}"
    ws.cell(r, 7, T["pkg_per_rack"][g]).font = BLUE; REF[f"pkg_{g}"] = f"In_TK!$G${r}"
    ws.cell(r, 8, "Interested-party／Assumed（Spec_Rack 標記）"); ws.cell(r, 9, f"{TKF} Spec_Rack 第 8–10 列（功率）、第 6 列（G 欄＝每架封裝數）")
    r += 1
r += 1
ws.cell(r, 1, "TK 證據工作簿（不在 Interface；非 NVIDIA 每 GW 顆數）").font = BOLD; r += 1
for key, lab, unit, val, tag, src in [
        ("ev_rackmw", "VR200 每 GW 機架 IT 功率", "MW", E["A7"]["vals"][1], "Derived", "Tokenomics docs/evidence/20261025_NonNV_TPU_Trainium/NonNV_TPU_Trainium.xlsx 推導!A7"),
        ("ev_tpu", "TPU v7 每 GW 顆數（每顆所攤 1.085 kW）", "顆/GW", E["C18"]["vals"][1], "Derived（功率 Interested-party）", "同上 推導!C18"),
        ("ev_tpu_kw_lo", "TPU v7 每顆所攤 kW 低端（1.30）", "kW", E["C4"]["vals"][0], "Derived", "同上 推導!C4"),
        ("ev_tpu_kw_hi", "TPU v7 每顆所攤 kW 高端（1.03）", "kW", E["C4"]["vals"][2], "Derived", "同上 推導!C4"),
        ("ev_trn", "Trainium3 每 GW 顆數（72 顆／架）", "顆/GW", E["D15"]["vals"][1], "Derived（功率 Interested-party）", "同上 推導!D15"),
        ("ev_amd_kw", "AMD Helios 機架功率（取 245 kW；AMD 說法 225–245）", "kW", 245.0, "Interested-party", "Tokenomics docs/evidence/20261025_NonNV_AMD/README.md（AMD、ServeTheHome、Schneider 參考設計）"),
        ("ev_amd_kw_lo", "AMD Helios 機架功率 低端", "kW", 225.0, "Interested-party", "同上"),
        ("ev_oai_kw", "OpenAI／Broadcom Jalapeño 每顆所攤 kW", "kW", 1.26, "Analogy（由 Broadcom 說法反推，口徑不明）", "Tokenomics docs/evidence/20261025_NonNV_Cerebras_OpenAI/README.md"),
        ("ev_s", "持有比公式中 IT 以外部分占 VR200 持有成本比 s（持有比＝1 − s ×（1 − IT capex 比））", "—", E["A20"]["vals"][1], "Derived", "TPU_Trainium 推導!A20"),
        ("ev_uplift", "IT 資本支出 ÷（機架數 × 機架價）網路等加成", "x", E["A10"]["vals"][1], "Derived", "TPU_Trainium 推導!A10")]:
    ws.cell(r, 1, key); ws.cell(r, 2, lab); ws.cell(r, 3, unit); ws.cell(r, 4, val).font = BLUE; ws.cell(r, 8, tag); ws.cell(r, 9, src)
    REF[key] = f"In_TK!$D${r}"; r += 1
for c_, w_ in zip("ABCI", [14, 60, 10, 100]):
    ws.column_dimensions[c_].width = w_

# ======================= In_WAT =======================
ws = wb.create_sheet("In_WAT")
header(ws, "In_WAT — Wattonomics v0.3.6 與協調橋接 v0.3.6 輸入（藍字）", ["鍵", "項目", "單位", "2026", "2027", "2028", "2029", "等級", "來源名稱"])
r = 4
for b in W["buyers"]:
    inrow(ws, f"wsh_{b}", f"出貨年 GW：{W['label_' + b].split('：')[0].strip()}", "GW", W[f"wsh_{b}"] + [None, None], "Derived", f"{WATF} Compare F2a wsh_{b}")
for b in W["buyers"]:
    inrow(ws, f"wnv_{b}", f"NVIDIA 出貨 GW：{b}", "GW", W[f"wnv_{b}"] + [None, None], "Derived（份額讀 SIF_nvsh：Interested-party／Assumed）", f"{WATF} Compare F2a wnv_{b}")
for b in W["buyers"]:
    ws.cell(r, 1, f"holdmap_{b}"); ws.cell(r, 2, f"換算份額用持有比（模型隱含平台）：{b}"); ws.cell(r, 4, W[f"hold_{b}"]).font = BLUE
    REF[f"holdmap_{b}"] = f"In_WAT!$D${r}"
    ws.cell(r, 8, "Derived"); ws.cell(r, 9, f"{WATF} Compare wgs_{b} 公式中的 TK_NonNV_Hold_*"); r += 1
inrow(ws, "w4_lo", "NVIDIA 出貨 GW：逐買方份額低值", "GW", W["w4_lo"], "Derived", f"{WATF} Compare F2a w4_lo")
inrow(ws, "w4_hi", "NVIDIA 出貨 GW：逐買方份額高值", "GW", W["w4_hi"], "Derived", f"{WATF} Compare F2a w4_hi")
inrow(ws, "recon_nvgw", "WIF_Recon_NV_NVGW（模型值）", "GW", [W["recon_NVGW_2026"], W["recon_NVGW_2027"], None, None], "Derived", "wattonomics_snapshot_v0.3.6 WIF_Recon_NV_NVGW_<年>")
inrow(ws, "recon_ship", "WIF_Recon_NV_ShipGW（模型值）", "GW", [W["recon_ShipGW_2026"], W["recon_ShipGW_2027"], None, None], "Derived", "wattonomics_snapshot_v0.3.6 WIF_Recon_NV_ShipGW_<年>")
for g in ["H100", "GB200", "GB300", "VR200", "RU"]:
    inrow(ws, f"mix_{g}", f"NVIDIA 世代占比（新增 GW）：{g}", "%", W["mix"][g], "Derived／Assumed（2027 起不計 Blackwell：工作單）", f"{WATF} Players mix_{g}")
inrow(ws, "s2", "平滑法實體出貨 GW（全市場；橋接 s2＝Compare capship）", "GW", B["s2"], "Derived（lag_on、inv_on Assumed）", f"{WATF} Compare capship；siliconomics/docs/reports/20261116_協調橋接_v0.3.6.xlsx Bridge s2")
inrow(ws, "x2", "換算值 x2（成本等值；Coord 需求側讀數）", "GW", B["x2"], "Derived", "協調橋接 v0.3.6 Bridge x2（＝s2 ÷ rt）；Siliconomics In crd_x2")
inrow(ws, "rt", "Siliconomics 實體 ÷ 成本等值（橋接用）", "x", B["rt"], "Derived", "協調橋接 v0.3.6 Bridge rt（v0.19；v0.20 同值）")
inrow(ws, "g2", "橋接 ② 時差與庫存（成本等值）", "GW", B["g2"], "Derived", "協調橋接 v0.3.6 Bridge g2")
for k, lab in [("sx_sl", "lag_on 低"), ("sx_sh", "lag_on 高"), ("sx_il", "inv_on 低"), ("sx_ih", "inv_on 高")]:
    inrow(ws, k, f"換算值敏感度：{lab}", "GW", B[k], "Derived", f"協調橋接 v0.3.6 Bridge {k}")
inrow(ws, "ai_base", "七家甲2 付款年 GW 合計（基準 ai_sh 0.85）", "GW", W["cg_ai_base"], "Derived", f"{WATF} CapexGW 第 267 列")
inrow(ws, "ai_lo", "同上（ai_sh 低 0.75）", "GW", W["cg_ai_lo"], "Assumed 敏感度", f"{WATF} CapexGW 第 271 列")
inrow(ws, "ai_hi", "同上（ai_sh 高 0.95）", "GW", W["cg_ai_hi"], "Assumed 敏感度", f"{WATF} CapexGW 第 272 列")
inrow(ws, "nnv_mkt", "市場非 NVIDIA 金額占比 nnv_mkt（Siliconomics Units 推得；CapexGW 換算用）", "%", W["nnv_mkt"], "Derived（取自 Siliconomics）", f"{WATF} Players 第 12 列")
inrow(ws, "spend_it", "Spend IT 設備小計（上線年）", "US$B", W["spend_it"], "Derived", f"{WATF} Spend 第 16 列")
for c_, w_ in zip("ABCI", [14, 56, 10, 90]):
    ws.column_dimensions[c_].width = w_

# ======================= In_CG（Wattonomics 付款年換算要素） =======================
ws = wb.create_sheet("In_CG")
header(ws, "In_CG — Wattonomics CapexGW 七家＋xAI 付款年換算要素（藍字；2026–2029）＋ 反事實重算（公式）", ["玩家", "項目", "單位", "2026", "2027", "2028", "2029", "等級", "來源名稱"])
r = 4
ws.cell(r, 1, "w"); ws.cell(r, 2, "廠房提前權重 w0／w1／w2（D／E／F）")
for j, v in zip([4, 5, 6], W["cg_w"]):
    ws.cell(r, j, v).font = BLUE
REF["w0"], REF["w1"], REF["w2"] = "In_CG!$D$4", "In_CG!$E$4", "In_CG!$F$4"
ws.cell(r, 8, "Assumed（lag_bldg）"); ws.cell(r, 9, f"{WATF} CapexGW!D4:F4"); r += 1
inrow(ws, "F7", "NVIDIA 每 IT GW IT capex（Wattonomics 世代加權）", "$B/GW", W["cg_F7"], "Derived", f"{WATF} CapexGW 第 7 列")
inrow(ws, "F8", "NVIDIA 每 IT GW 廠房 capex", "$B/GW", W["cg_F8"], "Derived", f"{WATF} CapexGW 第 8 列")
inrow(ws, "facVR", "VR200 每 GW 廠房 capex（TK）", "$B/GW", [S["tk_in"]["TK_CapexFacility_VR200"][1]] * 4, "Derived", f"{SILF} TK_In TK_CapexFacility_VR200")
HM = {"msft": "Maia", "googl": "TPU", "amzn": "Trn"}
cg_rows = {}
for p, d in W["cg"].items():
    start = r
    inrow(ws, f"{p}_ai", f"{p}：曆年 AI 資本支出", "US$B", d["ai"], "Derived（個股模型基準）", f"{WATF} CapexGW {p}_ai")
    inrow(ws, f"{p}_a1", f"{p}：t−1 AI 資本支出", "US$B", d["a1"], "Derived", f"{WATF} CapexGW {p}_a1")
    inrow(ws, f"{p}_a2", f"{p}：t−2 AI 資本支出", "US$B", d["a2"], "Derived", f"{WATF} CapexGW {p}_a2")
    inrow(ws, f"{p}_pit", f"{p}：每 IT GW IT capex（模型）", "$B/GW", d["pit"], "Derived", f"{WATF} CapexGW {p}_pit")
    inrow(ws, f"{p}_pfac", f"{p}：每 IT GW 廠房 capex（全額，模型）", "$B/GW", d["pfac"], "Derived", f"{WATF} CapexGW {p}_pfac")
    inrow(ws, f"{p}_pfo", f"{p}：每 IT GW 廠房 capex（帳上）", "$B/GW", d["pfo"], "Derived", f"{WATF} CapexGW {p}_pfo（無此列者＝pfac）")
    inrow(ws, f"{p}_gw", f"{p}：付款年 GW（模型值）", "GW", d["x"], "Derived", f"{WATF} CapexGW {p}_x")
    inrow(ws, f"{p}_nnv", f"{p}：CapexGW 換算用非 NVIDIA 占比", "%", W["players_nnv"][p], "Derived／Derived（逐家）", f"{WATF} Players {p}_nnv")
    hn = HM.get(p, "AMD")

    def fr(name, label, ftxt, tag="Derived（本檔公式）"):
        global r
        ws.cell(r, 1, name); ws.cell(r, 2, label)
        for y in YRS:
            ftx = ftxt(y)
            if ftx:
                ws.cell(r, YC[y], ftx)
                REF[(name, y)] = f"In_CG!${CL(YC[y])}${r}"
        ws.cell(r, 8, tag); r += 1

    g = lambda k, y: REF[(k, y)]
    fund = lambda y, pit, pfo: f"({pit}/({pit}+{pfo})*{g(p + '_ai', y)}+{pfo}/({pit}+{pfo})*({REF['w0']}*{g(p + '_ai', y)}+{REF['w1']}*{g(p + '_a1', y)}+{REF['w2']}*{g(p + '_a2', y)}))/({pit}+{pfo})"
    fr(f"{p}_gwchk", f"{p}：付款年 GW 以公式重算 − 模型值（應為 0）", lambda y: "=" + fund(y, g(p + '_pit', y), g(p + '_pfo', y)) + f"-{g(p + '_gw', y)}")
    # 反事實 B：NVIDIA 每 GW IT capex 改用 Siliconomics 當年世代組合（R1_SIL 的 F7_sil）
    fr(f"{p}_pitB", f"{p}：反事實每 IT GW IT capex（NVIDIA 世代組合＝Siliconomics）", lambda y: f"=(1-{g(p + '_nnv', y)})*R1_SIL!{CL(YC[y])}$40+{g(p + '_nnv', y)}*{REF[('hold_' + hn, 5)]}*{REF[('cit_VR200', 5)]}")
    fr(f"{p}_gwB", f"{p}：反事實付款年 GW（世代組合）", lambda y: "=" + fund(y, g(p + '_pitB', y), g(p + '_pfo', y)))
    # 反事實 C：換算用非 NVIDIA 占比改為對帳頁逐買方 GW 份額（1 − wgs；只 2026–27）
    fr(f"{p}_nC", f"{p}：對帳頁非 NVIDIA GW 占比（1 − wnv ÷ wsh）", lambda y: f"=1-{REF[('wnv_' + p, y)]}/{REF[('wsh_' + p, y)]}" if y in (2026, 2027) and W[f"wsh_{p}"][y - 2026] else None)
    fr(f"{p}_pitC", f"{p}：反事實每 IT GW IT capex（份額＝對帳頁）", lambda y: f"=(1-{g(p + '_nC', y)})*{g('F7', y)}+{g(p + '_nC', y)}*{REF[('hold_' + hn, 5)]}*{REF[('cit_VR200', 5)]}" if (p + '_nC', y) in REF else None)
    fr(f"{p}_pfoC", f"{p}：反事實每 IT GW 廠房 capex（帳上；份額＝對帳頁）", lambda y: f"=((1-{g(p + '_nC', y)})*{g('F8', y)}+{g(p + '_nC', y)}*{REF[('hold_' + hn, 5)]}*{g('facVR', y)})*IF({g(p + '_pfac', y)}=0,1,{g(p + '_pfo', y)}/{g(p + '_pfac', y)})" if (p + '_nC', y) in REF else None)
    fr(f"{p}_gwC", f"{p}：反事實付款年 GW（份額）", lambda y: ("=" + fund(y, g(p + '_pitC', y), g(p + '_pfoC', y))) if (p + '_nC', y) in REF else f"={g(p + '_gw', y)}")
    cg_rows[p] = True
    r += 1
for c_, w_ in zip("ABCI", [12, 56, 10, 60]):
    ws.column_dimensions[c_].width = w_

# ======================= R1_SIL =======================
ws = wb.create_sheet("R1_SIL")
header(ws, "R1_SIL — Siliconomics 平台別實體 GW 與成本等值 GW（公式）。實體：NVIDIA＝顆數 ÷ TK 每 GW 顆數；非 NVIDIA＝顆數 × 每顆 IT capex ÷（VR200 每 GW IT capex × 持有比）。全部 × crd_k", ["鍵", "項目", "每 GW 換算用", "2026", "2027", "2028", "2029", "等級", "說明"])
SIL_NV = [("bw", "GB300", "it_GB"), ("rb", "VR200", "it_VR"), ("ru", "RU", "it_RU"), ("hp", "H100", "it_HP")]
SIL_NN = [("amd", "AMD"), ("tpu", "TPU"), ("trn", "Trn"), ("oth", None)]
r = 4
RS = {}


def frow(ws, key, label, conv, ftxt, tag="Derived", note="", fill=None, store=None):
    global r
    ws.cell(r, 1, key); ws.cell(r, 2, label); ws.cell(r, 3, conv)
    for y in YRS:
        t = ftxt(y)
        if t is not None:
            c = ws.cell(r, YC[y], t); c.number_format = "0.000"
            if fill:
                c.fill = fill
            (store if store is not None else RS)[(key, y)] = f"{ws.title}!${CL(YC[y])}${r}"
    ws.cell(r, 8, tag); ws.cell(r, 9, note)
    r += 1


u = lambda p, y: REF[(f"u_{p}", y)]
k = lambda y: REF[("coef", y)]
den = lambda y: REF[("denom", y)]
ws.cell(r, 1, "A. 實體 GW").font = BOLD; r += 1
for p, g_, itn in SIL_NV:
    frow(ws, f"ph_{p}", f"實體 GW：{ug[p]['name']}", f"TK 每 GW 顆數 {g_}", lambda y, p=p, g_=g_: f"={u(p, y)}/{REF[(f'gpg_{g_}', 5)]}*1000000*{k(y)}")
for p, h in SIL_NN:
    if h:
        frow(ws, f"ph_{p}", f"實體 GW：{ug[p]['name']}", f"VR200 × 持有比 {h}", lambda y, p=p, h=h: f"={u(p, y)}*{REF['it_AMD8'] if (p == 'amd' and y == 2026) else REF['it_AMDR'] if p == 'amd' else REF['it_' + p.upper()]}/({REF[('cit_VR200', 5)]}*{REF[('hold_' + h, 5)]})*{k(y)}")
    else:
        frow(ws, f"ph_{p}", f"實體 GW：{ug[p]['name']}", "VR200 × 持有比（Maia、Ascend、OAI 平均）", lambda y, p=p: f"={u(p, y)}*{REF['it_ASIC']}/({REF[('cit_VR200', 5)]}*(({REF[('hold_Maia', 5)]}+{REF[('hold_Ascend', 5)]}+{REF[('hold_OAI', 5)]})/3))*{k(y)}")
frow(ws, "ph_NV", "實體 GW：NVIDIA 小計", "", lambda y: f"=SUM({RS[('ph_bw', y)].split('!')[1]}:{RS[('ph_hp', y)].split('!')[1]})", fill=KEY)
frow(ws, "ph_NNV", "實體 GW：非 NVIDIA 小計", "", lambda y: f"=SUM({RS[('ph_amd', y)].split('!')[1]}:{RS[('ph_oth', y)].split('!')[1]})", fill=KEY)
frow(ws, "ph_TOT", "實體 GW：合計", "", lambda y: f"={RS[('ph_NV', y)].split('!')[1]}+{RS[('ph_NNV', y)].split('!')[1]}", fill=KEY)
frow(ws, "ph_res", "殘差：SIF_DeployGWPhys_Base − 平台加總（應為 0）", "", lambda y: f"={REF[('sil_phys', y)]}-{RS[('ph_TOT', y)].split('!')[1]}")
r += 1
ws.cell(r, 1, "B. 成本等值 GW（＝顆數 × 每顆 IT capex × crd_k ÷ 當年 NVIDIA 每 GW IT capex）").font = BOLD; r += 1
for p, g_, itn in SIL_NV:
    frow(ws, f"ce_{p}", f"成本等值 GW：{ug[p]['name']}", itn, lambda y, p=p, itn=itn: f"={u(p, y)}*{REF[itn]}*{k(y)}/{den(y)}")
for p, h in SIL_NN:
    itn = {"amd": None, "tpu": "it_TPU", "trn": "it_TRN", "oth": "it_ASIC"}[p]
    frow(ws, f"ce_{p}", f"成本等值 GW：{ug[p]['name']}", itn or "it_AMD8（2026）／it_AMDR", lambda y, p=p, itn=itn: f"={u(p, y)}*{(REF['it_AMD8'] if y == 2026 else REF['it_AMDR']) if p == 'amd' else REF[itn]}*{k(y)}/{den(y)}")
frow(ws, "ce_NV", "成本等值 GW：NVIDIA 小計", "", lambda y: f"=SUM({RS[('ce_bw', y)].split('!')[1]}:{RS[('ce_hp', y)].split('!')[1]})", fill=KEY)
frow(ws, "ce_NNV", "成本等值 GW：非 NVIDIA 小計", "", lambda y: f"=SUM({RS[('ce_amd', y)].split('!')[1]}:{RS[('ce_oth', y)].split('!')[1]})", fill=KEY)
frow(ws, "ce_TOT", "成本等值 GW：合計", "", lambda y: f"={RS[('ce_NV', y)].split('!')[1]}+{RS[('ce_NNV', y)].split('!')[1]}", fill=KEY)
frow(ws, "ce_res", "殘差：SIF_DeployGW_Base − 平台加總（應為 0）", "", lambda y: f"={REF[('sil_ceq', y)]}-{RS[('ce_TOT', y)].split('!')[1]}")
r += 1
ws.cell(r, 1, "C. 輔助").font = BOLD; r += 1
frow(ws, "nvpk", "NVIDIA 封裝數（× crd_k；百萬顆）", "", lambda y: f"=({u('bw', y)}+{u('rb', y)}+{u('ru', y)}+{u('hp', y)})*{k(y)}")
frow(ws, "nvit", "NVIDIA 系統 IT 資本支出（× crd_k；US$B）", "", lambda y: f"=({u('bw', y)}*{REF['it_GB']}+{u('rb', y)}*{REF['it_VR']}+{u('ru', y)}*{REF['it_RU']}+{u('hp', y)}*{REF['it_HP']})*{k(y)}")
# F7_sil 必須在第 40 列（In_CG 反事實引用）
while r < 40:
    r += 1
assert r == 40
frow(ws, "F7sil", "NVIDIA 每 GW IT capex（Siliconomics 當年世代組合，實體 GW 加權）", "", lambda y: f"=({RS[('ph_bw', y)].split('!')[1]}*{REF[('cit_GB300', 5)]}+{RS[('ph_rb', y)].split('!')[1]}*{REF[('cit_VR200', 5)]}+{RS[('ph_ru', y)].split('!')[1]}*{REF[('cit_RU', 5)]}+{RS[('ph_hp', y)].split('!')[1]}*{REF[('cit_H100', 5)]})/{RS[('ph_NV', y)].split('!')[1]}")
for c_, w_ in zip("ABCI", [10, 52, 30, 60]):
    ws.column_dimensions[c_].width = w_

# ======================= R1_WAT =======================
ws = wb.create_sheet("R1_WAT")
header(ws, "R1_WAT — Wattonomics 實體出貨 GW：NVIDIA 依世代（Players 組合）、非 NVIDIA 依買方隱含平台；只有 2026–2027 可拆（Compare F 節）", ["鍵", "項目", "歸屬", "2026", "2027", "2028", "2029", "等級", "說明"])
r = 4
RW = {}
SPLIT = (2026, 2027)
ws.cell(r, 1, "A. NVIDIA 依世代（＝Σ逐買方 NVIDIA 出貨 GW × 世代占比）").font = BOLD; r += 1
frow(ws, "nv", "NVIDIA 出貨 GW（Σ逐買方）", "", lambda y: "=" + "+".join(REF[(f"wnv_{b}", y)] for b in W["buyers"]) if y in SPLIT else None, store=RW, fill=KEY)
frow(ws, "nvchk", "檢查：WIF_Recon_NV_NVGW − 上列（應為 0）", "", lambda y: f"={REF[('recon_nvgw', y)]}-{RW[('nv', y)].split('!')[1]}" if y in SPLIT else None, store=RW)
for pk, gens, lab in [("hp", ["H100"], "Hopper"), ("bw", ["GB200", "GB300"], "Blackwell（GB200＋GB300）"), ("rb", ["VR200"], "Rubin（VR200）"), ("ru", ["RU"], "Rubin Ultra")]:
    frow(ws, f"w_{pk}", f"實體 GW：NVIDIA {lab}", "Players 世代占比", lambda y, gens=gens: f"={RW[('nv', y)].split('!')[1]}*(" + "+".join(REF[(f'mix_{g}', y)] for g in gens) + ")" if y in SPLIT else None, store=RW)
r += 1
ws.cell(r, 1, "B. 非 NVIDIA 依買方（出貨 − NVIDIA），歸到模型換算份額所用持有比的平台").font = BOLD; r += 1
for b in W["buyers"]:
    frow(ws, f"nn_{b}", f"非 NVIDIA 出貨 GW：{b}", f"持有比 {W['hold_' + b]} → {({'TPU': 'TPU', 'Trn': 'Trainium', 'AMD': 'AMD', 'Maia': '其他', 'OAI': '其他', 'Ascend': '其他'})[W['hold_' + b]]}",
         lambda y, b=b: f"={REF[(f'wsh_{b}', y)]}-{REF[(f'wnv_{b}', y)]}" if y in SPLIT else None, store=RW)
PL = {"amd": ["AMD"], "tpu": ["TPU"], "trn": ["Trn"], "oth": ["Maia", "OAI", "Ascend"]}
for pk, hs in PL.items():
    bs = [b for b in W["buyers"] if W["hold_" + b] in hs]
    frow(ws, f"w_{pk}", f"實體 GW：{ {'amd': 'AMD', 'tpu': 'TPU', 'trn': 'Trainium', 'oth': '其他 ASIC'}[pk]}（模型隱含歸屬）", "＋".join(bs),
         lambda y, bs=bs: "=" + "+".join(RW[(f"nn_{b}", y)].split("!")[1] for b in bs) if y in SPLIT else None, store=RW)
frow(ws, "w_NV", "實體 GW：NVIDIA 小計", "", lambda y: "=" + "+".join(RW[(f"w_{p}", y)].split("!")[1] for p in ["hp", "bw", "rb", "ru"]) if y in SPLIT else None, store=RW, fill=KEY)
frow(ws, "w_NNV", "實體 GW：非 NVIDIA 小計", "", lambda y: "=" + "+".join(RW[(f"w_{p}", y)].split("!")[1] for p in ["amd", "tpu", "trn", "oth"]) if y in SPLIT else None, store=RW, fill=KEY)
frow(ws, "w_TOT", "實體出貨 GW 合計（Compare F (2)：七家與 xAI 付款年＋其餘鏡像）", "", lambda y: f"={RW[('w_NV', y)].split('!')[1]}+{RW[('w_NNV', y)].split('!')[1]}" if y in SPLIT else None, store=RW, fill=KEY)
frow(ws, "w_chk", "檢查：WIF_Recon_NV_ShipGW − 合計（應為 0）", "", lambda y: f"={REF[('recon_ship', y)]}-{RW[('w_TOT', y)].split('!')[1]}" if y in SPLIT else None, store=RW)
frow(ws, "w_s2", "對照：平滑法實體出貨 GW（全部鏡像；橋接 s2；2028–29 只有此讀數，不能拆平台）", "", lambda y: f"={REF[('s2', y)]}", store=RW)
frow(ws, "w_Fs2", "Compare F 讀法 − 平滑法", "", lambda y: f"={RW[('w_TOT', y)].split('!')[1]}-{RW[('w_s2', y)].split('!')[1]}" if y in SPLIT else None, store=RW)
r += 1
ws.cell(r, 1, "C. 成本等值（同 Siliconomics 定義：IT 資本支出 ÷ 當年 NVIDIA 每 GW IT capex；用 TK 係數）").font = BOLD; r += 1
frow(ws, "c_NV", "成本等值 GW：NVIDIA", "Σ 世代 GW × TK CapexIT ÷ 分母", lambda y: f"=({RW[('nv', y)].split('!')[1]}*({REF[('mix_H100', y)]}*{REF[('cit_H100', 5)]}+{REF[('mix_GB200', y)]}*{REF[('cit_GB200', 5)]}+{REF[('mix_GB300', y)]}*{REF[('cit_GB300', 5)]}+{REF[('mix_VR200', y)]}*{REF[('cit_VR200', 5)]}+{REF[('mix_RU', y)]}*{REF[('cit_RU', 5)]}))/{den(y)}" if y in SPLIT else None, store=RW)
frow(ws, "c_NNV", "成本等值 GW：非 NVIDIA", "Σ 買方 GW × 持有比 × VR200 ÷ 分母", lambda y: "=(" + "+".join(f"{RW[(f'nn_{b}', y)].split('!')[1]}*{REF[('hold_' + W['hold_' + b], 5)]}" for b in W["buyers"]) + f")*{REF[('cit_VR200', 5)]}/{den(y)}" if y in SPLIT else None, store=RW)
frow(ws, "c_TOT", "成本等值 GW：合計（本檔 TK 係數口徑）", "", lambda y: f"={RW[('c_NV', y)].split('!')[1]}+{RW[('c_NNV', y)].split('!')[1]}" if y in SPLIT else None, store=RW, fill=KEY)
frow(ws, "c_x2", "對照：協調用換算值 x2（＝平滑法 ÷ Siliconomics 的 rt）", "", lambda y: f"={REF[('x2', y)]}", store=RW)
r += 1
frow(ws, "nvpk", "隱含 NVIDIA 封裝數（百萬顆）＝Σ 世代 GW × TK 每 GW 顆數", "", lambda y: f"={RW[('nv', y)].split('!')[1]}*({REF[('mix_H100', y)]}*{REF[('gpg_H100', 5)]}+{REF[('mix_GB200', y)]}*{REF[('gpg_GB200', 5)]}+{REF[('mix_GB300', y)]}*{REF[('gpg_GB300', 5)]}+{REF[('mix_VR200', y)]}*{REF[('gpg_VR200', 5)]}+{REF[('mix_RU', y)]}*{REF[('gpg_RU', 5)]})/1000000" if y in SPLIT else None, store=RW)
frow(ws, "nvit", "NVIDIA 系統 IT 資本支出（US$B）＝Σ 世代 GW × TK CapexIT", "", lambda y: f"={RW[('c_NV', y)].split('!')[1]}*{den(y)}" if y in SPLIT else None, store=RW)
for c_, w_ in zip("ABCI", [10, 58, 34, 60]):
    ws.column_dimensions[c_].width = w_

# ======================= R1_並排 =======================
ws = wb.create_sheet("R1_並排")
header(ws, "R1_並排 — 逐平台、逐年兩層並排（實體 GW，出貨年）；差距＝Siliconomics − Wattonomics。2028–2029 Wattonomics 未拆平台（結構限制）", ["鍵", "平台", "層", "2026", "2027", "2028", "2029", "等級", "說明"])
r = 4
RP = {}
PLATS = [("bw", "NVIDIA Blackwell"), ("rb", "NVIDIA Rubin"), ("ru", "NVIDIA Rubin Ultra"), ("hp", "NVIDIA Hopper"), ("NV", "NVIDIA 小計"),
         ("amd", "AMD（Wattonomics 為隱含歸屬）"), ("tpu", "TPU（同上）"), ("trn", "Trainium（同上）"), ("oth", "其他 ASIC（同上）"), ("NNV", "非 NVIDIA 小計"), ("TOT", "合計")]
for pk, lab in PLATS:
    fill = KEY if pk in ("NV", "NNV", "TOT") else None
    frow(ws, f"s_{pk}", lab, "Siliconomics", lambda y, pk=pk: f"={RS[('ph_' + pk, y)]}", store=RP, fill=fill)
    frow(ws, f"w_{pk}", lab, "Wattonomics", lambda y, pk=pk: (f"={RW[('w_' + pk, y)]}" if y in SPLIT else ("模型未拆" if pk != "TOT" else f"={REF[('s2', y)]}")), store=RP, fill=fill,
         note="2028–29 合計＝平滑法 s2" if pk == "TOT" else "")
    frow(ws, f"g_{pk}", lab, "差距", lambda y, pk=pk: (f"={RP[('s_' + pk, y)].split('!')[1]}-{RP[('w_' + pk, y)].split('!')[1]}" if (y in SPLIT or pk == "TOT") else None), store=RP, fill=fill)
r += 1
ws.cell(r, 1, "NVIDIA 占差距比例").font = BOLD; r += 1
frow(ws, "nvsh", "NVIDIA 差距 ÷ 總差距（實體；Wattonomics Compare F 讀法）", "", lambda y: f"={RP[('g_NV', y)].split('!')[1]}/{RP[('g_TOT', y)].split('!')[1]}" if y in SPLIT else None, store=RP, fill=KEY)
frow(ws, "nvsh_lo", "同上，Wattonomics NVIDIA 用逐買方份額高值（w4_hi）", "", lambda y: f"=({RP[('s_NV', y)].split('!')[1]}-{REF[('w4_hi', y)]})/{RP[('g_TOT', y)].split('!')[1]}" if y in SPLIT else None, store=RP)
frow(ws, "nvsh_hi", "同上，Wattonomics NVIDIA 用逐買方份額低值（w4_lo）", "", lambda y: f"=({RP[('s_NV', y)].split('!')[1]}-{REF[('w4_lo', y)]})/{RP[('g_TOT', y)].split('!')[1]}" if y in SPLIT else None, store=RP)
frow(ws, "nvsh_s2", "同上，總差距改用平滑法 s2（NVIDIA 仍用 Compare F）", "", lambda y: f"={RP[('g_NV', y)].split('!')[1]}/({RP[('s_TOT', y)].split('!')[1]}-{REF[('s2', y)]})" if y in SPLIT else None, store=RP)
frow(ws, "nvsh_ce", "成本等值口徑：NVIDIA 差距 ÷ 總差距（Wattonomics 用 TK 係數換算）", "", lambda y: f"=({RS[('ce_NV', y)]}-{RW[('c_NV', y)]})/({RS[('ce_TOT', y)]}-{RW[('c_TOT', y)]})" if y in SPLIT else None, store=RP)
r += 1
ws.cell(r, 1, "成本等值並排（合計）").font = BOLD; r += 1
frow(ws, "ce_s", "Siliconomics 成本等值（＝SIF_DeployGW_Base）", "", lambda y: f"={RS[('ce_TOT', y)]}", store=RP)
frow(ws, "ce_w", "Wattonomics 換算值 x2", "", lambda y: f"={REF[('x2', y)]}", store=RP)
frow(ws, "ce_g", "差距（＝橋接殘差 ③）", "", lambda y: f"={RP[('ce_s', y)].split('!')[1]}-{RP[('ce_w', y)].split('!')[1]}", store=RP, fill=KEY)
frow(ws, "ce_wtk", "Wattonomics 成本等值（本檔 TK 係數、自身組成）", "", lambda y: f"={RW[('c_TOT', y)]}" if y in SPLIT else None, store=RP)
for c_, w_ in zip("ABCI", [10, 34, 14, 40]):
    ws.column_dimensions[c_].width = w_

# ======================= R1_歸因 =======================
ws = wb.create_sheet("R1_歸因")
header(ws, "R1_歸因 — 2026、2027 實體 GW 總差距（Wattonomics Compare F 讀法）逐項；2028–29 只到合計。正值＝可解釋 Siliconomics 高於 Wattonomics 的部分", ["鍵", "原因類別", "項目", "2026", "2027", "2028", "2029", "等級", "可否唯一歸因／說明"])
r = 4
RA = {}
frow(ws, "tot", "—", "總差距（實體；2026–27 Compare F，2028–29 平滑法 s2）", lambda y: f"={RP[('g_TOT', y)]}", store=RA, fill=KEY)
frow(ws, "a_gpg", "每 GW 顆數或內容", "NVIDIA 每 GW 顆數兩層不同（兩層都讀 TK v22.0 同值）", lambda y: f"=({REF[('gpg_VR200', 5)]}-{REF[('gpg_VR200', 6)]})+({REF[('gpg_GB300', 5)]}-{REF[('gpg_GB300', 6)]})+({REF[('gpg_RU', 5)]}-{REF[('gpg_RU', 6)]})", store=RA,
     tag="Verified（repo 內逐值相同）", note="0：Wattonomics 不用顆數算 GW；換成顆數時與 Siliconomics 用同一 TK 值")
frow(ws, "a_mix", "每 GW 顆數或內容", "NVIDIA 世代組合造成的每 GW 成本差（Wattonomics 2027 不計 Blackwell）", lambda y: "=" + "+".join(f"({REF[(p + '_gwB', y)]}-{REF[(p + '_gw', y)]})" for p in W["cg"]), store=RA,
     tag="Derived", note="可機械算出：七家與 xAI 付款年換算改用 Siliconomics 世代組合後 Wattonomics 增加的 GW（2026 為負＝Wattonomics 反而變少）")
frow(ws, "a_nnv", "NVIDIA 份額假設", "Wattonomics 內部兩套非 NVIDIA 占比不一致（CapexGW 換算 vs Compare 對帳）對總量的影響", lambda y: ("=" + "+".join(f"({REF[(p + '_gwC', y)]}-{REF[(p + '_gw', y)]})" for p in W["cg"])) if y in SPLIT else None, store=RA,
     tag="Derived", note="可機械算出；只影響每 GW 成本，因此只改總量一點點")
frow(ws, "a_share", "NVIDIA 份額假設", "NVIDIA 占比差異（Shapley 組成效果；不改總差距，只決定差距落在 NVIDIA 或非 NVIDIA）", lambda y: f"=({RS[('ph_NV', y)]}/{RS[('ph_TOT', y)]}-{RW[('w_NV', y)]}/{RW[('w_TOT', y)]})*({RS[('ph_TOT', y)]}+{RW[('w_TOT', y)]})/2" if y in SPLIT else None, store=RA,
     tag="Derived", note="不計入解釋總差距（組成）；區間見報告（兩種順序 3.25–5.01）")
frow(ws, "a_timeF", "年度歸屬", "Wattonomics 出貨年兩種讀法差（Compare F 付款年 − 平滑法）", lambda y: f"={RW[('w_Fs2', y)]}" if y in SPLIT else None, store=RA,
     tag="Derived", note="讀法選擇，不是差距原因；用平滑法時總差距再大這麼多")
frow(ws, "a_lag_lo", "年度歸屬", "出貨→上線參數（lag_on、inv_on）敏感度 低端（實體）", lambda y: f"=(MIN({REF[('sx_sl', y)]},{REF[('sx_sh', y)]},{REF[('sx_il', y)]},{REF[('sx_ih', y)]})-{REF[('x2', y)]})*{REF[('rt', y)]}", store=RA,
     tag="Assumed 敏感度", note="基準下為 0；只列區間，不歸因")
frow(ws, "a_lag_hi", "年度歸屬", "出貨→上線參數敏感度 高端（實體）", lambda y: f"=(MAX({REF[('sx_sl', y)]},{REF[('sx_sh', y)]},{REF[('sx_il', y)]},{REF[('sx_ih', y)]})-{REF[('x2', y)]})*{REF[('rt', y)]}", store=RA, tag="Assumed 敏感度")
frow(ws, "a_ai_hi", "買方清單／覆蓋", "七家 AI 占比 ai_sh 0.95（高）時 Wattonomics 增加的付款年 GW", lambda y: f"={REF[('ai_hi', y)]}-{REF[('ai_base', y)]}", store=RA,
     tag="Assumed 敏感度", note="候選解釋，無證據支持 ai_sh 偏低；不歸因")
r += 1
frow(ws, "expl", "—", "可唯一歸因合計（每 GW 成本：世代組合＋份額不一致）", lambda y: f"={RA[('a_mix', y)].split('!')[1]}" + (f"+{RA[('a_nnv', y)].split('!')[1]}" if ('a_nnv', y) in RA else ""), store=RA, fill=KEY)
frow(ws, "unexp", "未解釋", "未解釋（總差距 − 可唯一歸因；含買方覆蓋、買方資本支出口徑、供給側顆數高估、通路與供應商庫存等，模型內無法區分）", lambda y: f"={RA[('tot', y)].split('!')[1]}-{RA[('expl', y)].split('!')[1]}", store=RA, fill=BAD)
frow(ws, "unexp_p", "未解釋", "未解釋 ÷ 總差距", lambda y: f"={RA[('unexp', y)].split('!')[1]}/{RA[('tot', y)].split('!')[1]}", store=RA)
r += 1
frow(ws, "cum_s", "年度歸屬（累計檢查）", "Siliconomics 實體 GW 2026 起累計", lambda y: "=" + "+".join(RS[('ph_TOT', yy)] for yy in YRS if yy <= y), store=RA)
frow(ws, "cum_w", "年度歸屬（累計檢查）", "Wattonomics 平滑法 s2 2026 起累計", lambda y: "=" + "+".join(REF[('s2', yy)] for yy in YRS if yy <= y), store=RA)
frow(ws, "cum_g", "年度歸屬（累計檢查）", "累計差距（若只是年度歸屬，累計應收斂；實際逐年擴大）", lambda y: f"={RA[('cum_s', y)].split('!')[1]}-{RA[('cum_w', y)].split('!')[1]}", store=RA, fill=KEY)
for c_, w_ in zip("ABCI", [10, 22, 70, 70]):
    ws.column_dimensions[c_].width = w_

# ======================= R2_顆數 =======================
ws = wb.create_sheet("R2_顆數")
ws["A1"] = "R2_顆數 — 每 GW 顆數共同口徑：同一晶片世代在 Tokenomics、Siliconomics、Wattonomics（與 chip_bridge）的顆／GW；通過條件＝換算差 ≤ 1% 且口徑明示"; ws["A1"].font = BOLD
hd = ["世代", "TK Interface", "SIL TK_In", "WAT TK_In", "chip_bridge（v21.7）", "最大相對差", "≤1%？", "TK 規格重算（FLOOR）", "重算 − Interface", "功率低端顆數", "功率高端顆數", "低端相對", "高端相對", "每顆所攤 IT kW"]
for j, h in enumerate(hd, 1):
    c = ws.cell(3, j, h); c.font = BOLD; c.fill = HDR
r = 4
R2R = {}
for g in ["H100", "GB200", "GB300", "VR200", "RU"]:
    ws.cell(r, 1, GN[g])
    for j in [4, 5, 6, 7]:
        if (f"gpg_{g}", j) in REF:
            ws.cell(r, j - 2, f"={REF[(f'gpg_{g}', j)]}")
    vals = [f"C{r}", f"D{r}"] + ([f"E{r}"] if (f"gpg_{g}", 7) in REF else [])
    ws.cell(r, 6, "=MAX(" + ",".join(f"ABS({v}/B{r}-1)" for v in vals) + ")").number_format = "0.00%"
    ws.cell(r, 7, f'=IF(F{r}<=0.01,"是","否")')
    rk = lambda col: f"FLOOR({REF['it_mw']}*1000*(1-{REF['offrack']})*1000/{REF[(f'kw_{g}', col)]},1)*{REF[f'pkg_{g}']}"
    ws.cell(r, 8, "=" + rk(5)); ws.cell(r, 9, f"=H{r}-B{r}")
    ws.cell(r, 10, "=" + rk(4)); ws.cell(r, 11, "=" + rk(6))
    ws.cell(r, 12, f"=J{r}/B{r}-1").number_format = "0.0%"; ws.cell(r, 13, f"=K{r}/B{r}-1").number_format = "0.0%"
    ws.cell(r, 14, f"=1000000/B{r}").number_format = "0.000"
    R2R[g] = r; r += 1
r += 1
ws.cell(r, 1, "非 NVIDIA：三層都沒有每 GW 顆數；Siliconomics 與 Wattonomics 由金額換 GW（持有比當 IT capex 比）。隱含顆數對照 Tokenomics 證據（不在 Interface）").font = BOLD; r += 1
hd2 = ["平台", "Siliconomics 隱含顆／GW＝持有比×VR200 IT capex÷每顆 IT capex", "Wattonomics 隱含顆／GW", "TK 證據顆／GW", "相對差（SIL ÷ TK 證據 − 1）", "≤1%？", "TK 持有比", "持有比反推 IT capex 比 k＝1−(1−持有比)÷s", "2027 SIL 實體 GW（模型）", "2027 SIL 實體 GW（改用 TK 證據顆數）", "差", "證據來源"]
for j, h in enumerate(hd2, 1):
    c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
r += 1
for p, h, itref, evf, src in [("tpu", "TPU", REF["it_TPU"], f"={REF['ev_tpu']}", R["r2_nn"]["tpu"]["src"]),
                              ("trn", "Trn", REF["it_TRN"], f"={REF['ev_trn']}", R["r2_nn"]["trn"]["src"]),
                              ("amd", "AMD", REF["it_AMDR"], f"={REF['ev_rackmw']}*1000/{REF['ev_amd_kw']}*72", R["r2_nn"]["amd"]["src"]),
                              ("oth", None, REF["it_ASIC"], f"={REF['ev_rackmw']}*1000/{REF['ev_oai_kw']}", R["r2_nn"]["oth"]["src"])]:
    ws.cell(r, 1, {"tpu": "Google TPU", "trn": "AWS Trainium", "amd": "AMD Instinct（Helios）", "oth": "其他 ASIC（Siliconomics 用 Maia、Ascend、OAI 持有比平均；證據只有 OpenAI 客製）"}[p])
    hold_expr = REF[(f"hold_{h}", 5)] if h else f"(({REF[('hold_Maia', 5)]}+{REF[('hold_Ascend', 5)]}+{REF[('hold_OAI', 5)]})/3)"
    ws.cell(r, 2, f"={hold_expr}*{REF[('cit_VR200', 5)]}*1000000000/({itref}*1000)")
    ws.cell(r, 3, "無每顆價格 → 無法換顆數（只有金額）")
    ws.cell(r, 4, evf)
    ws.cell(r, 5, f"=B{r}/D{r}-1").number_format = "0.0%"
    ws.cell(r, 6, f'=IF(ABS(E{r})<=0.01,"是","否")')
    ws.cell(r, 7, f"={hold_expr}")
    ws.cell(r, 8, f"=1-(1-G{r})/{REF['ev_s']}")
    ws.cell(r, 9, f"={RS[('ph_' + p, 2027)]}")
    ws.cell(r, 10, f"={REF[('u_' + p, 2027)]}*{REF[('coef', 2027)]}*1000000/D{r}")
    ws.cell(r, 11, f"=J{r}-I{r}")
    ws.cell(r, 12, src)
    R2R[p] = r; r += 1
ws.cell(r, 1, "合計"); ws.cell(r, 9, f"=SUM(I{r-4}:I{r-1})"); ws.cell(r, 10, f"=SUM(J{r-4}:J{r-1})"); ws.cell(r, 11, f"=J{r}-I{r}"); R2R["nn_sum"] = r; r += 2
ws.cell(r, 1, "口徑明示檢查").font = BOLD; r += 1
hd3 = ["處", "每 GW 的電力口徑", "功率基礎", "PUE", "利用率", "明示？"]
for j, h in enumerate(hd3, 1):
    c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR
r += 1
for row in [["Tokenomics Interface", "IT 關鍵電力（Interface 第 1 列、Inputs!E5）", "機架配電設計功率（Spec_Rack 第 9 列，功率情境 2）；扣機架外 8%", "1.2，只用於設施電力（IF_FacilityGW）", "0.8 只用於能耗，不進顆數", "是"],
            ["Siliconomics TK_In／Deploy", "IT 電力（Deploy A7『部署 GW（IT 電力）』）；TK_In 表頭未寫電力口徑", "沿用 TK；非 NVIDIA 無功率基礎（金額換算）", "未用", "未用", "NVIDIA 部分：是（經 TK）；非 NVIDIA：否"],
            ["Wattonomics TK_In", "IT 關鍵電力（TK_In A2 明寫）", "沿用 TK；晶片功耗換 IT 用 k_chip2it 1.65／1.85／2.08（Analogy，與 TK 機架規格無連結）", "IF_FacilityGW 1.2", "未用", "是（但 k_chip2it 與 TK 規格不同源）"],
            ["markets/chip_bridge", "沿用 TK（只有 GB300、VR200）", "沿用 TK", "未用", "未用", "是；來源標示 v21.7"]]:
    for j, v in enumerate(row, 1):
        ws.cell(r, j, v)
    r += 1
r += 1
ws.cell(r, 1, "chip_bridge 讀 Tokenomics v21.7 而非 v22.0 的影響").font = BOLD; r += 1
cbr = R["r2_cb"]
for t in [f"chip_bridge 讀的 8 個 IF_ 名稱共 {cbr['ncells_cb']} 格：v21.7 與 v22.0 不同的格數＝{len(cbr['diffs_cb'])}（換算差 0%）",
          f"兩版快照全部 22 個名稱中只有 {len(cbr['diffs_all'])} 格不同，全在 IF_NonNV_Out_TPU、IF_NonNV_Out_Trn（產出比；chip_bridge 不讀）",
          "chip_bridge dist（v0.2.3）In 頁 GB300、VR200 的顆數、機架數、IT／廠房／合計 capex 與 v22.0 逐格相減皆為 0",
          "另：Siliconomics TK_In 第 2 列仍寫『markets/chip_bridge v20.1』，與 build_bridge.py 實際讀 v21.7 不一致（標示問題，不影響數值）"]:
    ws.cell(r, 1, t); r += 1
for c_, w_ in zip("ABCDEFGHIJKL", [44, 22, 18, 18, 16, 10, 12, 18, 14, 16, 12, 70]):
    ws.column_dimensions[c_].width = w_

# ======================= 自我核對 =======================
ws = wb.create_sheet("自我核對")
header(ws, "自我核對 — 公式結果（LibreOffice 重算）與 Python 獨立重算（scripts/compute.py）之差；|差| < 1e-6 為通過", ["項目", "公式儲存格", "Python 值", "公式值", "差", "通過"])
checks = []
for y in YRS:
    s = R["sil"][str(y)]
    for pk in ["bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth"]:
        checks.append((f"SIL 實體 {pk} {y}", RS[(f"ph_{pk}", y)], s[pk]["phys"]))
        checks.append((f"SIL 成本等值 {pk} {y}", RS[(f"ce_{pk}", y)], s[pk]["ceq"]))
    checks.append((f"SIL 實體合計 {y}", RS[("ph_TOT", y)], s["TOT"]["phys"]))
    checks.append((f"SIL 成本等值合計 {y}", RS[("ce_TOT", y)], s["TOT"]["ceq"]))
    checks.append((f"SIL F7 世代組合 {y}", RS[("F7sil", y)], R["att"][str(y)]["F7_sil"]))
    checks.append((f"歸因 世代組合 {y}", RA[("a_mix", y)], R["att"][str(y)]["mix_effect"]))
    checks.append((f"總差距 {y}", RA[("tot", y)], R["gap"][str(y)]["tot_phys_F"] if y in SPLIT else R["gap"][str(y)]["tot_phys_s2"]))
    checks.append((f"累計差距 {y}", RA[("cum_g", y)], sum(R["sil"][str(yy)]["TOT"]["phys"] - R["wat"][str(yy)]["s2"] for yy in YRS if yy <= y)))
    if y in SPLIT:
        w = R["wat"][str(y)]; g = R["gap"][str(y)]
        for pk in ["hp", "bw", "rb", "ru", "amd", "tpu", "trn", "oth"]:
            checks.append((f"WAT 實體 {pk} {y}", RW[(f"w_{pk}", y)], w["plat"][pk]))
        checks += [(f"WAT NVIDIA {y}", RW[("w_NV", y)], w["NV_phys"]), (f"WAT 合計 {y}", RW[("w_TOT", y)], w["tot_phys"]),
                   (f"WAT 成本等值 {y}", RW[("c_TOT", y)], w["tot_ceq_tk"]), (f"WAT NVIDIA 封裝數 {y}", RW[("nvpk", y)], w["nv_pkgs"]),
                   (f"NVIDIA 占差距 {y}", RP[("nvsh", y)], g["NV_share_phys"]), (f"NVIDIA 占差距（份額高） {y}", RP[("nvsh_lo", y)], g["NV_share_phys_lo"]),
                   (f"NVIDIA 占差距（份額低） {y}", RP[("nvsh_hi", y)], g["NV_share_phys_hi"]), (f"NVIDIA 占差距（成本等值） {y}", RP[("nvsh_ce", y)], g["NV_share_ceq"]),
                   (f"份額組成效果 {y}", RA[("a_share", y)], g["share_eff"]), (f"歸因 份額不一致 {y}", RA[("a_nnv", y)], R["att"][str(y)]["nnv_effect"])]
for g_ in ["H100", "GB200", "GB300", "VR200", "RU"]:
    checks.append((f"R2 最大相對差 {g_}", f"R2_顆數!$F${R2R[g_]}", R["r2_nv"][g_]["max_rel"]))
    checks.append((f"R2 規格重算 {g_}", f"R2_顆數!$H${R2R[g_]}", R["r2_nv"][g_]["recalc_base"]))
    checks.append((f"R2 功率低端 {g_}", f"R2_顆數!$J${R2R[g_]}", R["r2_nv"][g_]["pow_lo_cnt"]))
for p in ["tpu", "trn", "amd", "oth"]:
    checks.append((f"R2 非 NVIDIA 隱含顆數 {p}", f"R2_顆數!$B${R2R[p]}", R["r2_nn"][p]["sil_implied"]))
    checks.append((f"R2 非 NVIDIA 相對差 {p}", f"R2_顆數!$E${R2R[p]}", R["r2_nn"][p]["rel"]))
    checks.append((f"R2 2027 改用證據顆數 {p}", f"R2_顆數!$J${R2R[p]}", R["r2_alt_nnv_phys"]["2027"][p]))
r = 4
for lab, cell, pv in checks:
    ws.cell(r, 1, lab); ws.cell(r, 2, cell); ws.cell(r, 3, pv)
    ws.cell(r, 4, f"={cell}")
    ws.cell(r, 5, f"=D{r}-C{r}")
    ws.cell(r, 6, f'=IF(ABS(E{r})<0.000001,"通過","不符")')
    r += 1
ws.cell(r + 1, 1, "不符項數"); ws.cell(r + 1, 4, f'=COUNTIF(F4:F{r-1},"不符")')
ws.column_dimensions["A"].width = 40; ws.column_dimensions["B"].width = 26
wb.save(OUT)
print("saved", OUT, "checks", len(checks))

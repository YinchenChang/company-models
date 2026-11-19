# 唯讀收集：從三層模型與快照抽出 R1、R2 需要的輸入，寫成 inputs.json（不改任何 repo 檔案）
import json, os, hashlib
import openpyxl

import sys
CM = sys.argv[1] if len(sys.argv) > 1 else "/home/claude/company-models"
TKR = sys.argv[2] if len(sys.argv) > 2 else "/home/claude/tokenomics"
OUT = sys.argv[3] if len(sys.argv) > 3 else "inputs.json"
YRS = [2026, 2027, 2028, 2029]


def _cur(p):
    return open(p, encoding="utf-8").read().strip()


_WX = _cur(f"{CM}/wattonomics/model/CURRENT")
_WV = __import__("re").search(r"_(v[\d.]+)\.xlsx$", _WX).group(1)
SIL_X = f"{CM}/siliconomics/model/" + _cur(f"{CM}/siliconomics/model/CURRENT")
SIL_J = f"{CM}/siliconomics/data/" + _cur(f"{CM}/siliconomics/data/CURRENT_SNAPSHOT")
WAT_X = f"{CM}/wattonomics/model/{_WX}"
WAT_J = f"{CM}/wattonomics/data/wattonomics_snapshot_{_WV}.json"
TK_X = f"{TKR}/model/" + _cur(f"{TKR}/model/CURRENT")
TK_SIL22 = f"{CM}/siliconomics/data/tokenomics_snapshot_v22.0.json"
TK_SIL217 = f"{CM}/siliconomics/data/tokenomics_snapshot_v21.7.json"
TK_WAT22 = f"{CM}/wattonomics/data/tokenomics_snapshot_v22.0.json"
CB_X = sorted(__import__("glob").glob(f"{CM}/markets/chip_bridge/dist/*.xlsx"))[-1]
BR_X = sorted(__import__("glob").glob(f"{CM}/siliconomics/docs/reports/*_協調橋接_{_WV}*.xlsx"))[-1]
TPU_EV = f"{TKR}/docs/evidence/20261025_NonNV_TPU_Trainium/NonNV_TPU_Trainium.xlsx"


def sha(p):
    return hashlib.sha256(open(p, "rb").read()).hexdigest()


D = {"files": {}, "sil": {}, "wat": {}, "tk": {}, "cb": {}, "br": {}, "ev": {}}
for k, p in [("SIL_X", SIL_X), ("SIL_J", SIL_J), ("WAT_X", WAT_X), ("WAT_J", WAT_J), ("TK_X", TK_X), ("TK_SIL22", TK_SIL22),
             ("TK_SIL217", TK_SIL217), ("TK_WAT22", TK_WAT22), ("CB_X", CB_X), ("BR_X", BR_X), ("TPU_EV", TPU_EV)]:
    D["files"][k] = {"path": p, "sha256": sha(p)}

# ---------------- Siliconomics ----------------
sj = json.load(open(SIL_J))["names"]
S = D["sil"]
for k in ["bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth"]:
    S[f"u_{k}"] = sj[f"SIF_UnitsPlat_{k}"][1:]          # 2026–29（快照首格為 2025 空白）
S["coef"] = sj["SIF_UnitsCoef_Base"][1:]
S["deploy"] = sj["SIF_DeployGW_Base"]
S["phys"] = sj["SIF_DeployGWPhys_Base"]
S["alt"] = [sj[f"SIF_DeployGW_Alt_{y}"] for y in YRS]
sv = openpyxl.load_workbook(SIL_X, data_only=True)


def nm(wb, name):
    sh, ref = list(wb.defined_names[name].destinations)[0]
    return wb[sh][ref.replace("$", "")].value


for n in ["it_GB", "it_VR", "it_RU", "it_HP", "it_AMD8", "it_AMDR", "it_TPU", "it_TRN", "it_ASIC", "vr26"]:
    S[n] = nm(sv, n)
for n in ["it_AMD8", "it_AMDR", "it_TPU", "it_TRN", "it_ASIC"]:
    # In 頁 D、F 欄＝低、高
    sh, ref = list(sv.defined_names[n].destinations)[0]
    r = sv[sh][ref.replace("$", "")].row
    S[n + "_lohi"] = [sv[sh].cell(r, 4).value, sv[sh].cell(r, 6).value]
dep = sv["Deploy"]
S["deploy_denom"] = [dep.cell(6, c).value for c in range(3, 7)]
S["deploy_nv_phys"] = [dep.cell(8, c).value for c in range(3, 7)]
S["deploy_nnv_phys"] = [dep.cell(9, c).value for c in range(3, 7)]
S["deploy_itcapex_raw"] = [dep.cell(4, c).value for c in range(3, 7)]
# TK_In 鏡像（Siliconomics 版）
tki = sv["TK_In"]
S["tk_in"] = {}
for r in range(5, 59):
    key = tki.cell(r, 3).value
    if key:
        S["tk_in"][key] = [tki.cell(r, 4).value, tki.cell(r, 5).value, tki.cell(r, 6).value]
S["tk_in_header"] = tki["A1"].value
S["tk_in_note_J"] = tki["A2"].value
S["deploy_note"] = dep["A1"].value

# ---------------- Wattonomics ----------------
wv = openpyxl.load_workbook(WAT_X, data_only=True)
W = D["wat"]
cmp_ = wv["Compare"]
col = {2025: 6, 2026: 7, 2027: 8, 2028: 9, 2029: 10}
buyers = ["msft", "googl", "amzn", "meta", "orcl", "crwv", "nbis", "wyfi", "xai", "oai", "ant", "neo", "sov", "cn", "oth"]
rowkey = {}
for r in range(25, 200):
    b = cmp_.cell(r, 2).value
    if b:
        rowkey[b] = r
W["buyers"] = buyers
for b in buyers:
    for pre in ["wsh", "wgs", "wgs_lo", "wgs_hi", "wnv"]:
        r = rowkey[f"{pre}_{b}"]
        W[f"{pre}_{b}"] = [cmp_.cell(r, col[y]).value for y in [2026, 2027]]
    W[f"label_{b}"] = cmp_.cell(rowkey[f"wsh_{b}"], 1).value
    # c 的持有比名稱（從公式字串抽出）
wf = openpyxl.load_workbook(WAT_X)
import re
for b in buyers:
    f = wf["Compare"].cell(rowkey[f"wgs_{b}"], 8).value
    m = re.search(r"TK_NonNV_Hold_(\w+?)\*", f)
    W[f"hold_{b}"] = m.group(1)
for k in ["w2", "w4", "w4_lo", "w4_hi", "w3", "w5", "wdsh", "wgsh", "capship", "capfx"]:
    r = rowkey[k]
    W[k] = [cmp_.cell(r, col[y]).value for y in YRS]
pl = wv["Players"]
W["mix"] = {}
for r, g in zip(range(5, 10), ["H100", "GB200", "GB300", "VR200", "RU"]):
    W["mix"][g] = [pl.cell(r, col[y]).value for y in YRS]
W["nnv_mkt"] = [pl.cell(12, col[y]).value for y in YRS]
W["players_nnv"] = {}
for r in range(1, pl.max_row + 1):
    b = pl.cell(r, 2).value
    if isinstance(b, str) and b.endswith("_nnv"):
        W["players_nnv"][b[:-4]] = [pl.cell(r, col[y]).value for y in YRS]
W["mktadd"] = [wv["Deploy"].cell(14, col[y]).value for y in YRS] + [wv["Deploy"].cell(14, 11).value]
# CapexGW：七家與 xAI 的付款年換算要素（2026–2029）
cg = wv["CapexGW"]
ck = {}
for r in range(1, cg.max_row + 1):
    b = cg.cell(r, 2).value
    if b:
        ck[b] = r
W["cg_w"] = [cg["D4"].value, cg["E4"].value, cg["F4"].value]
W["cg_F7"] = [cg.cell(7, col[y]).value for y in YRS]
W["cg_F8"] = [cg.cell(8, col[y]).value for y in YRS]
W["cg"] = {}
for p in ["msft", "googl", "amzn", "meta", "orcl", "crwv", "nbis", "xai"]:
    d = {}
    aik = f"{p}_ai"
    if p == "xai":
        r_ai = ck["xai_pit"] - 1
        d["ai"] = [cg.cell(r_ai, col[y]).value for y in YRS]
        d["ai_label"] = cg.cell(r_ai, 1).value
    else:
        d["ai"] = [cg.cell(ck[aik], col[y]).value for y in YRS]
    for f_ in ["pit", "pfac", "a1", "a2", "gw" if p != "xai" else "gw", "x"]:
        key = f"{p}_{f_}"
        if key in ck:
            d[f_] = [cg.cell(ck[key], col[y]).value for y in YRS]
    pfo_key = f"{p}_pfo"
    d["pfo"] = [cg.cell(ck[pfo_key], col[y]).value for y in YRS] if pfo_key in ck else d["pfac"]
    W["cg"][p] = d
for _r in (267, 271, 272): print("CapexGW row", _r, cg.cell(_r, 1).value, cg.cell(_r, 2).value)
W["cg_ai_lo"] = [cg.cell(271, col[y]).value for y in YRS]
W["cg_ai_base"] = [cg.cell(267, col[y]).value for y in YRS]
W["cg_ai_hi"] = [cg.cell(272, col[y]).value for y in YRS]
wi = wv["In"]
for r in range(1, wi.max_row + 1):
    a = wi.cell(r, 1).value
    if a in ("lag_on", "inv_on", "ai_sh", "k_chip2it", "spend_nnv_hold", "lag_bldg"):
        W["in_" + a] = [wi.cell(r, 8).value, wi.cell(r, 9).value, wi.cell(r, 10).value]
wj = json.load(open(WAT_J))["names"]
for y in [2025, 2026, 2027]:
    for k in ["NVGW", "ShipGW", "ContentPerGW"]:
        W[f"recon_{k}_{y}"] = wj[f"WIF_Recon_NV_{k}_{y}"]
wtk = wv["TK_In"]
W["tk_in"] = {}
for r in range(5, 64):
    key = wtk.cell(r, 3).value
    if key:
        W["tk_in"][key] = [wtk.cell(r, 4).value, wtk.cell(r, 5).value, wtk.cell(r, 6).value]
W["tk_in_header"] = wtk["A1"].value + " " + (wtk["A2"].value or "")
W["spend_it"] = [wv["Spend"].cell(16, col[y]).value for y in YRS]

# ---------------- 協調橋接（rt、x2、敏感度） ----------------
bv = openpyxl.load_workbook(BR_X, data_only=True)["Bridge"]
B = D["br"]
for r in range(4, bv.max_row + 1):
    k = bv.cell(r, 1).value
    if k in ("s2", "x2", "x0", "rt", "g", "g2", "g3", "sx_sl", "sx_sh", "sx_il", "sx_ih", "s0"):
        B[k] = [bv.cell(r, c).value for c in range(5, 9)]

# ---------------- Tokenomics ----------------
tv = openpyxl.load_workbook(TK_X, data_only=True)
itf = tv["Interface"]
T = D["tk"]
gens = ["H100", "GB200", "GB300", "VR200", "RU"]
base_col = {"H100": 4, "GB200": 7, "GB300": 10, "VR200": 13, "RU": 16}
for lab, r in [("RacksPerGW", 6), ("GPUsPerGW", 7), ("FacilityGW", 8), ("CapexIT", 9), ("CapexFacility", 10)]:
    T[lab] = {g: [itf.cell(r, base_col[g] - 1).value, itf.cell(r, base_col[g]).value, itf.cell(r, base_col[g] + 1).value] for g in gens}
dc = tv["DC_Cost"]
spec = tv["Spec_Rack"]
T["rack_kw"] = {g: [spec.cell(8, c).value, spec.cell(9, c).value, spec.cell(10, c).value] for g, c in zip(gens, range(3, 8))}
T["pkg_per_rack"] = {g: spec.cell(6, c).value for g, c in zip(gens, range(3, 8))}
inp = tv["Inputs"]
T["it_gw"] = inp["E5"].value
T["power_case"] = inp["E6"].value
T["offrack"] = inp["E8"].value
T["pue"] = [inp["D9"].value, inp["E9"].value, inp["F9"].value]
T["util"] = [inp["D10"].value, inp["E10"].value, inp["F10"].value]
nn = tv["NonNV"]
T["hold"] = {}
for r, key in zip(range(5, 12), ["AMD", "TPU", "Trn", "CS", "OAI", "Ascend", "Maia"]):
    T["hold"][key] = [nn.cell(r, 5).value, nn.cell(r, 6).value, nn.cell(r, 7).value]

# 三份 TK 快照（Siliconomics v22.0、v21.7；Wattonomics v22.0）
for key, p in [("snap_sil22", TK_SIL22), ("snap_sil217", TK_SIL217), ("snap_wat22", TK_WAT22)]:
    T[key] = json.load(open(p))

# chip_bridge 實際建檔值（dist）
cbv = openpyxl.load_workbook(CB_X, data_only=True)["In"]
C = D["cb"]
C["A1"] = cbv["A1"].value
for r in range(5, 14):
    k = cbv.cell(r, 1).value
    if k:
        C[k] = [cbv.cell(r, c).value for c in range(4, 10)]

# Tokenomics 證據工作簿（非 Interface）：TPU v7、Trainium3 每 GW 顆數
ev = openpyxl.load_workbook(TPU_EV, data_only=True)["推導"]
E = D["ev"]
for row in ev.iter_rows(values_only=True):
    if row and row[0] in ("A7", "A10", "A20", "C4", "C17", "C18", "D3", "D11", "D15", "F4"):
        E[row[0]] = {"label": row[1], "unit": row[2], "vals": [row[3], row[4], row[5]]}

json.dump(D, open(OUT, "w"), ensure_ascii=False, indent=1, default=str)
print("ok", OUT)

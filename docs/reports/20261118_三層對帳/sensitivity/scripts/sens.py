# LBNL 式單因子敏感度：2027 年供給側（Siliconomics）− 需求側（Wattonomics 換算值 x2）差距。
# 一次只動一個因子；用 replica.py（已對兩層活頁簿 LibreOffice 重算逐項核對）。輸出 results.json。
import sys, json, copy
sys.path.insert(0, "/home/claude/out/sens/scripts")
import replica as R
I = R.I; TK = I["tk"]
B = R.base_params(); BASE = R.gap(B)
G0 = BASE["G"]

def P(): return R.base_params()
KW = {"H100": (38, 41, 44), "GB200": (120, 132, 140), "GB300": (132, 136, 140), "VR200": (188, 227, 230), "RU": (300, 340, 454)}
def power(p, j):              # j＝0 低功率端（每 GW 顆數較多）、2 高功率端
    for g, kw in KW.items():
        f = kw[1] / kw[j]; p["GPG"][g] *= f; p["CIT"][g] *= f
    return p
def offrack(p, off):
    f = (1 - off) / (1 - 0.08)
    for g in KW: p["GPG"][g] *= f; p["CIT"][g] *= f
    return p
def tkcost(p, j):   # 相對倍數（可與功率、機架外疊乘）
    for g in KW: p["CIT"][g] *= TK["CapexIT"][g][j] / TK["CapexIT"][g][1]; p["FAC"][g] = TK["CapexFacility"][g][j]
    return p
def nvmult(p, m, keys=("bw", "rb", "ru")):
    for k in keys: p["raw"][k] *= m
    return p
HBM_NOW = R.C_CAP * 1.15 * sum(R.base_params()["raw"][k] * R.base_params()["hbm_gb"][k] for k in R.base_params()["raw"]) / 1000; HBM_IND = (2.06, 3.82, 6.97)
def hbmcap(p, sup):
    p["cap"] = min(1.0, R.C_CAP * sup / HBM_NOW); return p
NN_LO = {"amd": 1.2, "tpu": 5.0, "trn": 2.0, "oth": 3.0}; NN_HI = {"amd": 1.8, "tpu": 8.84, "trn": 3.5, "oth": 6.0}
IT_LO = {"amd": 80, "tpu": 15, "trn": 12, "oth": 12}; IT_HI = {"amd": 120, "tpu": 35, "trn": 30, "oth": 30}
RCONV = {"tpu": 1.2415106461, "trn": 1.6100380200, "amd": 1.1819181351, "oth": 1.5875444934}
def conv(p):
    for k, v in RCONV.items(): p["r_conv"][k] = v
    p["r_conv"].update({"TPU": RCONV["tpu"], "Trn": RCONV["trn"], "AMD": RCONV["amd"], "Maia": RCONV["oth"]})
    return p
# 七家 2027、2028 AI 資本支出：市場共識（In Y2）÷ 個股模型（Y9）；Microsoft Y2 不含融資租賃（口徑不同）不換；Meta 無 Y2
CONS = {"googl": (300.979 / 290.3651, 339.451 / 317.3988), "amzn": (275.063 / 253.0801, 292.927 / 266.9716),
        "crwv": (46.476 / 42.5812, 50.129 / 34.2227), "nbis": (27.243 / 31.9887, 38.39 / 32.9622)}
def mulm(p, q, v):
    p["mAI"][q] = [a * b for a, b in zip(p["mAI"][q], v)]
def cons(p):
    for q, (m7, m8) in CONS.items(): mulm(p, q, [1, 1, 1, m7, m8])
    fy = R.D["orcl"]["fyai"]; ai = R.D["orcl"]["ai"]
    fy28c = 100.962 * 0.938578542167978; fy29c = 84.848 * 0.938578542167978   # Oracle FY28、FY29 共識 × 模型 AI 比（FY29 比沿用 FY28）
    fy29m = (ai[4] - 5 / 12 * fy[4]) * 12 / 7
    mulm(p, "orcl", [1, 1, 1, (5 / 12 * fy[3] + 7 / 12 * fy28c) / ai[3], (5 / 12 * fy28c + 7 / 12 * fy29c) / ai[4]])
    return p
def rev(p, m):
    for q in R.PL:
        if q != "xai": mulm(p, q, [1, 1, 1, m, m])
    return p
def aish(p, j):
    f = (0.75, 0.85, 0.95)[j] / 0.85; fa = (0.60, 0.68, 0.75)[j] / 0.68
    mm = R.aish_mult(f, fa)
    for q in R.PL: mulm(p, q, mm[q])
    return p
def setv(p, **kv): p.update(kv); return p
SHIPF = 28.337257833714467   # Wattonomics Compare F 出貨年 GW（七家與 xAI 付款年＋其餘鏡像）

# 因子表：(id, 類, 側, 名稱, 等級, (低端說明, 低端函式), (高端說明, 高端函式), 狀態, 出處)
F = [
 ("S1", "P", "供給", "HBM 位元供給（截頂係數）", "Analogy（獨立估計）／Derived（現行）",
  ("獨立估計基準 3.82 EB → 係數 0.539", lambda p: hbmcap(p, HBM_IND[1])), ("不截頂（係數 1；＝券商原始顆數）", lambda p: setv(p, cap=1.0)),
  "新算", "docs/reports/20261118_三層對帳/hbm/20261118_HBM獨立供給_查證.md；Siliconomics v0.18 S48"),
 ("S2", "I", "供給", "NVIDIA 2027 顆數：券商上限型用語當點值（以 2026 供給側高估幅度類比）", "Analogy",
  ("× 0.864（K1：2026 上半年季節權重高端 42.4% 下供給側高估 15.8%）", lambda p: nvmult(p, 164.2 / (448.36 * 0.424))), ("× 1.00（上限型用語：不會更高）", lambda p: p),
  "新算", "analysis-K/20261118_K1_歷史回測.md 第 2、5 節；unit-basis 第六節第 3 點"),
 ("S3", "A", "供給", "NVIDIA 2027 顆數：營收＝顆數×單價，對 FY2028 指引", "Interested-party（指引）／Derived（單價）",
  ("× 0.833（基準單價下供給側高指引 20%）", lambda p: nvmult(p, 1 / 1.20)), ("× 1.00（單價取低端時指引容得下現值）", lambda p: p),
  "新算", "unit-basis/20261118_顆數來源口徑查核.md 4.3(c)"),
 ("S4", "P", "供給", "Rubin Ultra 重複計算（Rubin 7.0 → 6.7）", "Interested-party（MS 轉述）",
  ("Rubin 原始 6.7（截頂係數固定）", lambda p: (p["raw"].__setitem__("rb", p["raw"]["rb"] - 0.3), p)[1]), ("—（現值）", lambda p: p),
  "已確認（#662：截頂重算後 39.16→38.74）", "summary/20261118_三層對帳_第二波匯總報告.md 第 0 節；#662"),
 ("S5", "P", "供給", "封裝／晶粒口徑（最壞：2027 Blackwell 240 萬為晶粒）", "Interested-party（AC Research，原文未取得）",
  ("Blackwell × 0.5", lambda p: nvmult(p, 0.5, ("bw",))), ("—（兩層都是封裝）", lambda p: p),
  "已排除（只留最壞情況上限）", "unit-basis 第一節"),
 ("S6", "P", "供給", "CoWoS 產能上限（2027 年底月產能 KGI 17 萬片）", "Interested-party",
  ("產能 1,980 千片 ÷ 需求 1,982.6 → × 0.9987", lambda p: setv(p, cap=R.C_CAP * 1980 / 1982.6161)), ("—（未觸頂）", lambda p: p),
  "新算", "Siliconomics v0.20 Cap 頁、In cw_capend27"),
 ("S7", "A", "供給", "非 NVIDIA 2027 顆數（In 低／高；TPU 用 MS 5.0／GF 8.84）", "Interested-party／Assumed",
  ("AMD 1.2、TPU 5.0、Trn 2.0、其他 3.0（原始）", lambda p: (p["raw"].update(NN_LO), p)[1]), ("AMD 1.8、TPU 8.84、Trn 3.5、其他 6.0", lambda p: (p["raw"].update(NN_HI), p)[1]),
  "新算", "Siliconomics v0.20 In 第 70、76、82、88 列"),
 ("S8", "A", "供給", "非 NVIDIA 每顆 IT 資本支出（it_AMDR、it_TPU、it_TRN、it_ASIC）", "Assumed",
  ("80／15／12／12 千美元", lambda p: (p["it_nn"].update(IT_LO), p)[1]), ("120／35／30／30 千美元", lambda p: (p["it_nn"].update(IT_HI), p)[1]),
  "新算", "Siliconomics v0.20 In 第 222–225 列"),
 ("S9", "P", "共用", "非 NVIDIA 每 GW 換算（持有比當 IT capex 比；兩層同改用 Tokenomics 證據每 GW 顆數）", "Derived（功率 Interested-party）",
  ("—（現行慣例）", lambda p: p), ("改用證據值（TPU ×1.24、Trn ×1.61、AMD ×1.18、其他 ×1.59）", conv),
  "新算", "analysis-R 第五節非 NVIDIA 表"),
 ("T1", "P", "共用", "NVIDIA 機架設計功率（Tokenomics 功率情境；每 GW 顆數與每 GW IT capex 同比）", "Interested-party",
  ("低功率（VR200 188 kW、RU 300、GB300 132）", lambda p: power(p, 0)), ("高功率（VR200 230、RU 454、GB300 140）", lambda p: power(p, 2)),
  "新算", "Tokenomics v22.0 Spec_Rack 第 8、10 列（analysis-R 第五節）"),
 ("T2", "P", "共用", "機架外 IT 開銷（現值 8%）", "Assumed／Analogy",
  ("5%", lambda p: offrack(p, 0.05)), ("14%（Epoch IT 開銷 1.14；定義是否相同未驗證）", lambda p: offrack(p, 0.14)),
  "新算", "Tokenomics Spec_Rack；peers 報告 Epoch"),
 ("T3", "P", "共用", "PUE（設施 ÷ IT）", "Verified（口徑）",
  ("1.1", lambda p: p), ("1.3", lambda p: p), "新算（兩層都是 IT 關鍵電力，PUE 不進 GW）", "analysis-R 第五節口徑"),
 ("T4", "A", "共用", "Tokenomics 成本情境（每 GW IT 與廠房資本支出）", "Derived（機架價 Interested-party）",
  ("低成本（VR200 IT 32.0、廠房 10.2）", lambda p: tkcost(p, 0)), ("高成本（VR200 IT 42.6、廠房 16.2）", lambda p: tkcost(p, 2)),
  "新算", "Tokenomics v22.0 IF_CapexIT、IF_CapexFacility 低／高"),
 ("D1", "A", "需求", "買方資本支出 AI 占比 ai_sh（Amazon ai_amzn）", "Assumed（Amazon Derived）",
  ("0.75（Amazon 0.60）", lambda p: aish(p, 0)), ("0.95（Amazon 0.75）", lambda p: aish(p, 2)),
  "新算", "Wattonomics v0.3.6 In 第 13、14 列"),
 ("D2", "A", "需求", "七家 2027–28 資本支出水準：改用市場共識（In Y2）", "Interested-party",
  ("—（個股模型）", lambda p: p), ("共識（Alphabet、Amazon、Oracle、CoreWeave、Nebius；Microsoft 口徑不同不換）", cons),
  "新算", "Wattonomics v0.3.6 In 第 121–127 列"),
 ("D3", "I", "需求", "買方資本支出共識事後上修（Coatue：2026 共識比年初高 65%）", "Analogy／Interested-party",
  ("× 1.00（不上修）", lambda p: p), ("七家 2027–28 × 1.65", lambda p: rev(p, 1.65)),
  "新算", "peers 報告第四節；Coatue 頁面（WebFetch 2026-11-18）"),
 ("D4", "A", "需求", "廠房提前入帳 lag_bldg", "Assumed",
  ("12 月", lambda p: setv(p, lag_bldg=12)), ("24 月", lambda p: setv(p, lag_bldg=24)), "新算", "Wattonomics In 第 10 列"),
 ("D5", "P", "共用", "出貨（付款）→上線時差 lag_on（Wattonomics 與橋接同參數）", "Assumed",
  ("3 月", lambda p: setv(p, lag_on=3)), ("6 月", lambda p: setv(p, lag_on=6)), "新算", "Wattonomics In 第 94 列；橋接 In lag_on"),
 ("D6", "P", "共用", "未上線庫存 ÷ 年出貨 inv_on（同上）", "Assumed",
  ("5%", lambda p: setv(p, inv_on=0.05)), ("20%", lambda p: setv(p, inv_on=0.2)), "新算", "Wattonomics In 第 95 列"),
 ("D7", "I", "需求", "站點覆蓋率（七家與 xAI 以外、以站點為主的買方：OpenAI 自建、Anthropic、中國、新雲）", "Analogy（Epoch 自估）",
  ("覆蓋 77% → × 1.30", lambda p: setv(p, cov=1 / 0.77)), ("覆蓋 44% → × 2.27", lambda p: setv(p, cov=1 / 0.44)),
  "新算", "peers 報告（Epoch 44%，90% 區間 29–77%）"),
 ("D8", "A", "需求", "「其他（企業、學研）」成長倍數 oth_gmult", "Analogy",
  ("0（持平）", lambda p: setv(p, oth_gmult=0)), ("1.5", lambda p: setv(p, oth_gmult=1.5)), "新算", "Wattonomics In 第 96 列"),
 ("D9", "I", "需求", "需求側 NVIDIA 份額（讀供給側買方份額 SIF_nvsh）", "Interested-party／Assumed",
  ("份額低值", lambda p: p), ("份額高值", lambda p: p), "新算（總量 0；只改 NVIDIA 占差距 66%–107%）", "analysis-R 第一、二節"),
 ("D10", "P", "需求", "Wattonomics 2027 不計 Blackwell（Blackwell 占 NVIDIA 新增 GW）", "Assumed→Analogy／Interested-party",
  ("10%", lambda p: setv(p, dwt=[25.59 - 25.5005334169982, 29.39 - 29.3111548208607])), ("16.2%", lambda p: setv(p, dwt=[25.65 - 25.5005334169982, 29.45 - 29.3111548208607])),
  "已排除（W2_A；本表只把其市場年新增代入橋接）", "wattonomics/docs/reports/20261118_W2_A_2027不計Blackwell重審.md 第 3 節"),
 ("D11", "P", "需求", "非 NVIDIA 廠房成本乘持有比（拿掉）", "Assumed",
  ("—（現行）", lambda p: p), ("廠房全額", lambda p: setv(p, fac_hold=False)), "已排除（W2_B：方向相反）", "wattonomics/docs/reports/20261118_W2_B_非NVIDIA廠房持有比口徑.md 第 2 節"),
 ("D12", "A", "需求", "出貨年讀法：七家與 xAI 改用付款年（Compare F 28.34）而非平滑法 s2", "Derived",
  ("—（平滑法）", lambda p: p), ("付款年讀法", "SHIPF"), "新算（讀法選擇，不是原因）", "analysis-R 第三節；橋接 v0.3.6"),
]

def run(fn):
    p = P()
    if fn == "SHIPF":
        g = dict(R.gap(p)); g["x2"] = g["shipF"] / g["rt"]; g["G"] = g["Gf"]; g["Gphys"] = g["phys"] - g["shipF"]; return g
    return R.gap(fn(p))
def run_recap(fn):
    p = P(); p = fn(p) if fn != "SHIPF" else p; p["recap"] = True
    return R.gap(p)
res = []
SUPPLY_UNITS = {"S2", "S3", "S4", "S5", "S7"}
for fid, cls, side, name, grade, lo, hi, status, src in F:
    glo = run(lo[1]); ghi = run(hi[1])
    row = {"id": fid, "cls": cls, "side": side, "name": name, "grade": grade, "lo_lab": lo[0], "hi_lab": hi[0], "status": status, "src": src,
           "G_lo": glo["G"], "G_hi": ghi["G"], "d_lo": glo["G"] - G0, "d_hi": ghi["G"] - G0,
           "df_lo": glo["Gf"] - BASE["Gf"], "df_hi": ghi["Gf"] - BASE["Gf"],
           "Gp_lo": glo["Gphys"], "Gp_hi": ghi["Gphys"], "S0_lo": glo["S0"], "S0_hi": ghi["S0"], "x2_lo": glo["x2"], "x2_hi": ghi["x2"]}
    if fid in SUPPLY_UNITS:
        rlo = run_recap(lo[1]); rhi = run_recap(hi[1])
        row.update({"rc_d_lo": rlo["G"] - G0, "rc_d_hi": rhi["G"] - G0})
    row["swing"] = max(abs(row["d_lo"]), abs(row["d_hi"]))
    row["close"] = min(row["d_lo"], row["d_hi"])       # 最能縮小差距的一端（負＝縮小）
    res.append(row)
res.sort(key=lambda r: -r["swing"])
out = {"base": BASE, "G0": G0, "factors": res}

# ---- 聯合（殘差）：每個參數取「縮小差距」那一端，同參數的替代證據只取一個 ----
def joint(ids_fns):
    p = P()
    for fn in ids_fns: p = fn(p)
    return R.gap(p)
FN = {r[0]: r for r in F}
def closing_fn(fid):
    r = next(x for x in res if x["id"] == fid)
    f = FN[fid]
    return f[5][1] if r["d_lo"] <= r["d_hi"] else f[6][1]
sets = {
 "供給側（NVIDIA 顆數取指引 S3、Rubin Ultra S4、非 NVIDIA 顆數 S7 低、每顆 IT capex S8 低）": ["S3", "S4", "S7", "S8"],
 "需求側（D1 AI 占比高、D2 共識、D4 廠房 12 月、D7 覆蓋 44%、D8 其他 1.5）": ["D1", "D2", "D4", "D7", "D8"],
 "共用物理（T1 低功率、T2 機架外 5%…取縮小端、T4 成本取縮小端、S9 換算）": ["T1", "T2", "T4", "S9"],
 "有外部證據端點者（S3 指引、S4 重複計算、D2 共識）": ["S3", "S4", "D2"],
}
jo = {}
for nm, ids in sets.items():
    fns = [closing_fn(i) for i in ids]
    g = joint(fns)
    ssum = sum(next(x for x in res if x["id"] == i)["close"] for i in ids)
    jo[nm] = {"ids": ids, "sum_single": ssum, "joint": g["G"] - G0, "interaction": (g["G"] - G0) - ssum, "G_after": g["G"]}
out["joint"] = jo

# ---- 交互作用示例 ----
def two(f1, f2):
    a = R.gap(f1(P()))["G"] - G0; b = R.gap(f2(P()))["G"] - G0; ab = R.gap(f2(f1(P())))["G"] - G0
    return {"a": a, "b": b, "ab": ab, "inter": ab - a - b}
out["inter"] = {
 "NVIDIA 顆數 S3 × 機架功率 T1 低": two(lambda p: nvmult(p, 1 / 1.2), lambda p: power(p, 0)),
 "AI 占比 D1 高 × 共識上修 D3": two(lambda p: aish(p, 2), lambda p: rev(p, 1.65)),
 "NVIDIA 顆數 S3 × HBM 截頂重算": {"fixed": R.gap(nvmult(P(), 1 / 1.2))["G"] - G0, "recap": R.gap(setv(nvmult(P(), 1 / 1.2), recap=True))["G"] - G0},
 "HBM 獨立估計 S1 × 指引 S3": two(lambda p: hbmcap(p, HBM_IND[1]), lambda p: nvmult(p, 1 / 1.2)),
}

# ---- 什麼必須為真 ----
def solve(f, lo, hi, tol=1e-9):
    a, b = lo, hi
    fa = f(a)
    for _ in range(200):
        m = (a + b) / 2; fm = f(m)
        if (fm > 0) == (fa > 0): a, fa = m, fm
        else: b = m
    return (a + b) / 2
mt = {}
# 供給側正確 → 需求側要多少
mt["d_capex_mult_8"] = solve(lambda m: R.gap((lambda p: ([mulm(p, q, [1, 1, 1, m, m]) for q in R.PL], p)[1])(P()))["G"], 1, 4)
mt["d_capex_mult_8_pay"] = solve(lambda m: R.gap((lambda p: ([mulm(p, q, [1, 1, 1, m, m]) for q in R.PL], p)[1])(P()))["Gf"], 1, 4)
mt["d_capex_mult_7_pay"] = solve(lambda m: R.gap(rev(P(), m))["Gf"], 1, 4)
mt["d_capex_mult_7"] = solve(lambda m: R.gap(rev(P(), m))["G"], 1, 4)
mt["d_extra_on_gw"] = solve(lambda x: R.gap(setv(P(), dwt=[x, x]))["G"], 0, 40)
def demcost(p, f):
    # 只改需求側每 GW 成本（IT＋廠房）：在 replica 中，需求側與供給側共用 CIT；這裡只縮放需求側，供給側另算
    return p
def g_demcost(f):
    p = P(); s = R.supply(p)
    q = P()
    for g in KW: q["CIT"][g] *= f; q["FAC"][g] *= f
    d = R.demand(q); rt = s["phys"] / s["S0"]
    return s["S0"] - d["s2"] / rt
mt["d_cost_mult"] = solve(g_demcost, 0.2, 1.0)
mt["d_cov"] = solve(lambda c: R.gap(setv(P(), cov=c))["G"], 1, 20)
# 需求側正確 → 供給側要多少
mt["s_nv_mult_fixed"] = solve(lambda m: R.gap(nvmult(P(), m))["G"], 0, 1)
mt["s_all_mult"] = solve(lambda m: R.gap(setv(P(), cap=R.C_CAP * m))["G"], 0.01, 1)
mt["s_hbm_sup_recap"] = solve(lambda s: R.gap(setv(P(), recap=True, hbm_sup=s))["G"], 0.5, 6.6727)
mt["base_nv_units"] = BASE["nv_units"]; mt["base_nv_phys"] = BASE["nv_phys"]
# 隱含 NVIDIA 2027 資料中心營收（凍結值 743.8 依 NVIDIA 顆數等比；未含價格變動）
mt["sil_nv_rev_frozen"] = 743.78; mt["wat_nv_rev_frozen"] = 408.76; mt["guid_fy28_dc"] = 633.7
mt["nv_rev_if_demand_right"] = 743.78 * mt["s_nv_mult_fixed"]
out["must"] = mt
json.dump(out, open("/home/claude/out/sens/scripts/results.json", "w"), ensure_ascii=False, indent=1, default=float)
print(f"G0={G0:.4f}")
for r in res:
    extra = f" | 截頂重算 {r.get('rc_d_lo', float('nan')):+.2f}/{r.get('rc_d_hi', float('nan')):+.2f}" if "rc_d_lo" in r else ""
    print(f"{r['id']:4s} {r['cls']} {r['side']} 低 {r['d_lo']:+7.2f} 高 {r['d_hi']:+7.2f} 付款年 {r['df_lo']:+6.2f}/{r['df_hi']:+6.2f}  phys {r['Gp_lo']-BASE['Gphys']:+6.2f}/{r['Gp_hi']-BASE['Gphys']:+6.2f}{extra}  {r['name'][:40]}")
for k, v in jo.items(): print(k, {a: (round(b, 3) if isinstance(b, float) else b) for a, b in v.items()})
for k, v in out["inter"].items(): print(k, {a: round(b, 3) for a, b in v.items()})
print({k: round(v, 4) for k, v in mt.items()})

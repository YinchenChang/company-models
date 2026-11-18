# 2027 年差距的縮減式複製（Siliconomics 供給鏈＋Wattonomics 需求鏈＋協調橋接 x2）。
# 只讀 repo 抽出的值（wat_extract.json、analysis-R inputs.json）；不改任何模型。
import json, copy
D = json.load(open("/home/claude/out/sens/scripts/wat_extract.json"))
I = json.load(open("/home/claude/wt/sens/docs/reports/20261118_三層對帳/analysis-R/scripts/inputs.json"))
YR = [2024, 2025, 2026, 2027, 2028]          # wat_extract 欄序
PL = ["msft", "googl", "amzn", "meta", "orcl", "crwv", "nbis", "xai"]
HOLDN = {"msft": "Maia", "googl": "TPU", "amzn": "Trn", "meta": "AMD", "orcl": "AMD", "crwv": "AMD", "nbis": "AMD", "xai": "AMD"}
SITE11 = PL + ["neo", "sov", "oth"]
C_CAP = 2.256909 / 2.4

def base_params():
    tk = I["tk"]
    p = {
        # ---- Tokenomics（兩層共用） ----
        "CIT": {g: tk["CapexIT"][g][1] for g in ["H100", "GB200", "GB300", "VR200", "RU"]},
        "FAC": {g: tk["CapexFacility"][g][1] for g in ["H100", "GB200", "GB300", "VR200", "RU"]},
        "GPG": {g: tk["GPUsPerGW"][g][1] for g in ["H100", "GB200", "GB300", "VR200", "RU"]},
        "HOLD": {h: tk["hold"][h][1] for h in tk["hold"]},
        # ---- Siliconomics 2027 原始顆數（截頂前，百萬封裝）與截頂 ----
        "raw": {"bw": 2.4, "rb": 7.0, "ru": 0.3, "hp": 0.0, "amd": 1.5, "tpu": 5.5, "trn": 2.5, "oth": 4.5},
        "cap": C_CAP, "recap": False,
        "hbm_gb": {"bw": 280, "rb": 288, "ru": 1024, "hp": 141, "amd": 360, "tpu": 250, "trn": 144, "oth": 200},
        "hbm_gross": 1.15, "hbm_sup": None,          # None＝現行（＝截頂後需求）
        "it_nn": {"amd": 100.0, "tpu": 25.0, "trn": 20.0, "oth": 20.0},
        "r_conv": {"amd": 1.0, "tpu": 1.0, "trn": 1.0, "oth": 1.0, "Maia": 1.0, "Trn": 1.0, "TPU": 1.0, "AMD": 1.0},  # 非 NVIDIA 每 GW 換算修正（R2）
        # ---- Wattonomics ----
        "lag_on": 4.5, "inv_on": 0.1, "lag_bldg": 18.0, "oth_gmult": 1.0,
        "mAI": {q: [1, 1, 1, 1, 1] for q in PL},     # 各家各年 AI 資本支出乘數（2024–2028）
        "cov": 1.0, "fac_hold": True, "dwt": [0.0, 0.0],
        "mix": {2024: {"H100": 1.0}, 2025: {"H100": 0.1, "GB200": 0.7, "GB300": 0.2}, 2026: {"GB300": 0.78, "VR200": 0.22},
                2027: {"VR200": I["wat"]["mix"]["VR200"][1], "RU": I["wat"]["mix"]["RU"][1]},
                2028: {"VR200": I["wat"]["mix"]["VR200"][2], "RU": I["wat"]["mix"]["RU"][2]}},
    }
    return p

def nnv(q, y):
    pn = I["wat"]["players_nnv"][q]
    return pn[max(0, y - 2026)] if y <= 2029 else pn[-1]

def supply(p):
    CIT, GPG, H = p["CIT"], p["GPG"], p["HOLD"]
    den = CIT["VR200"]
    nvg = {"bw": "GB300", "rb": "VR200", "ru": "RU", "hp": "H100"}
    hold = {"amd": H["AMD"], "tpu": H["TPU"], "trn": H["Trn"], "oth": (H["Maia"] + H["Ascend"] + H["OAI"]) / 3}
    raw = p["raw"]
    if p["recap"]:
        sup = p["hbm_sup"] if p["hbm_sup"] is not None else C_CAP * p["hbm_gross"] * sum(base_params()["raw"][k] * p["hbm_gb"][k] for k in raw) / 1000
        dem = p["hbm_gross"] * sum(raw[k] * p["hbm_gb"][k] for k in raw) / 1000
        c = min(1.0, sup / dem)
    else:
        c = p["cap"]
    u = {k: raw[k] * c for k in raw}
    it = {k: CIT[g] / GPG[g] * 1e6 for k, g in nvg.items()}
    it.update(p["it_nn"])
    S0 = sum(u[k] * it[k] for k in u) / den
    nv_phys = sum(u[k] * 1e6 / GPG[g] for k, g in nvg.items())
    nn_phys = sum(u[k] * it[k] / (den * hold[k]) * p["r_conv"][k] for k in hold)
    nv_ceq = sum(u[k] * it[k] for k in nvg) / den
    return {"S0": S0, "phys": nv_phys + nn_phys, "nv_phys": nv_phys, "nn_phys": nn_phys, "nv_ceq": nv_ceq, "c": c,
            "nv_units": sum(u[k] for k in nvg)}

def demand(p):
    CIT, FAC, H = p["CIT"], p["FAC"], p["HOLD"]
    L = p["lag_bldg"] / 12; w = (max(0, 1 - L), max(0, min(L, 2 - L)), max(0, min(1, L - 1)))
    k = p["lag_on"] / 12 + p["inv_on"]
    F7 = {y: sum(v * CIT[g] for g, v in p["mix"][y].items()) for y in YR}
    F8 = {y: sum(v * FAC[g] for g, v in p["mix"][y].items()) for y in YR}
    pay, main, dye = {}, {}, {}
    for q in PL:
        d = D[q]; pay[q] = {}; main[q] = {}
        for i, y in enumerate(YR):
            if y < 2025: continue
            m = p["mAI"][q]
            ai = d["ai"][i]; a1 = d["a1"][i]; a2 = d["a2"][i]
            if ai is None: pay[q][y] = None; continue
            ai *= m[i]; a1 *= m[i - 1]; a2 *= m[max(i - 2, 0)]
            n = 0.0 if q in ("crwv", "nbis", "xai") else nnv(q, y)
            h = H[HOLDN[q]]; r = p["r_conv"][HOLDN[q]]
            pit = (1 - n) * F7[y] + n * h * CIT["VR200"] / r
            pfac = (1 - n) * F8[y] + n * (h if p["fac_hold"] else 1.0) * FAC["VR200"]
            own = d["own"][i] if "own" in d and d["own"][i] is not None else 1.0
            pfo = own * pfac
            fund = pit / (pit + pfo) * ai + pfo / (pit + pfo) * (w[0] * ai + w[1] * a1 + w[2] * a2)
            pay[q][y] = fund / (pit + pfo)
        for y in (2026, 2027, 2028):
            prev = pay[q].get(y - 1)
            if prev is None:   # 前一年無模型值（xAI 2026、CoreWeave 2025 以 甲 代入）
                prev = D[q]["x"][YR.index(y - 1)]
                on = pay[q][y] if q == "xai" else pay[q][y] - k * (pay[q][y] - prev)
            else:
                on = pay[q][y] - k * (pay[q][y] - prev)
            if q == "crwv" and y == 2026: on = max(on, 0.85)
            main[q][y] = on
        dye[q] = {}
        cum = 0.0
        for y in (2026, 2027, 2028):
            cum += main[q][y] - D[q]["main"][YR.index(y)]
            dye[q][y] = cum
    # 「其他」隨市場成長
    mx = {y: D["oth"]["mx"][YR.index(y)] + sum(dye[q][y] for q in PL) for y in (2026, 2027, 2028)}
    mx[2025] = D["oth"]["mx"][YR.index(2025)]
    ye = {2025: D["oth"]["ye"][YR.index(2025)]}; oth = {}
    for y in (2026, 2027, 2028):
        g = mx[y] / mx[y - 1] - 1
        ye[y] = ye[y - 1] * (1 + g * p["oth_gmult"]); oth[y] = ye[y] - ye[y - 1]
    wt = {}
    for j, y in enumerate((2027, 2028)):
        i = YR.index(y)
        R = D["pipeadd"][i] - sum((D[q]["site"][i] or 0) for q in SITE11)
        wt[y] = (R + D["neo"]["main"][i]) * p["cov"] + D["sov"]["main"][i] + oth[y] + sum(main[q][y] for q in PL) + p["dwt"][j]
    s2 = wt[2027] + k * (wt[2028] - wt[2027])
    oth_on = {y: wt[y] - sum(main[q][y] for q in PL) for y in (2027, 2028)}   # 七家與 xAI 以外（上線年）
    shipF = sum(pay[q][2027] for q in PL) + oth_on[2027] + k * (oth_on[2028] - oth_on[2027])   # Compare F：七家與 xAI 付款年＋其餘鏡像
    return {"wt27": wt[2027], "wt28": wt[2028], "s2": s2, "shipF": shipF, "k": k, "pay27": sum(pay[q][2027] for q in PL), "oth27": oth[2027]}

def gap(p):
    s = supply(p); d = demand(p)
    rt = s["phys"] / s["S0"]
    x2 = d["s2"] / rt
    return {**s, **d, "rt": rt, "x2": x2, "G": s["S0"] - x2, "Gphys": s["phys"] - d["s2"], "Gf": s["S0"] - d["shipF"] / rt}

if __name__ == "__main__":
    b = gap(base_params())
    print({k: round(v, 6) for k, v in b.items() if isinstance(v, float)})

def aish_mult(f_gen, f_amzn):
    """ai_sh（各家未覆寫者）與 ai_amzn 改值時，各家各年 AI 資本支出乘數。Oracle FY2027 起讀個股模型 AI 比，只有 FY2025–26 受 ai_sh 影響。xAI 不受影響。"""
    m = {q: [1.0] * 5 for q in PL}
    for q in ("msft", "googl", "meta", "crwv", "nbis"): m[q] = [f_gen] * 5
    m["amzn"] = [f_amzn] * 5
    fy = D["orcl"]["fyai"]   # 2024..2028 財年 AI 資本支出（FY 結束於 5 月；曆年＝5/12 FY_y＋7/12 FY_y+1）
    cy26 = 5 / 12 * fy[2] + 7 / 12 * fy[3]
    m["orcl"] = [f_gen, f_gen, (5 / 12 * fy[2] * f_gen + 7 / 12 * fy[3]) / cy26, 1.0, 1.0]
    return m

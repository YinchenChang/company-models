# Python 獨立重算：R1 平台別拆解、歸因、R2 每 GW 顆數對照。只讀 inputs.json，輸出 results.json
import json, math

import sys
D = json.load(open(sys.argv[1] if len(sys.argv) > 1 else "inputs.json"))
S, W, T, B, E, C = D["sil"], D["wat"], D["tk"], D["br"], D["ev"], D["cb"]
YRS = [2026, 2027, 2028, 2029]
R = {}

tk = lambda n: S["tk_in"][n][1]           # Siliconomics TK_In 基準
VR = tk("TK_CapexIT_VR200")
CIT = {g: tk(f"TK_CapexIT_{g}") for g in ["H100", "GB200", "GB300", "VR200", "RU"]}
GPG = {g: tk(f"TK_GPUsPerGW_{g}") for g in ["H100", "GB200", "GB300", "VR200", "RU"]}
HOLD = {h: tk(f"TK_NonNV_Hold_{h}") for h in ["AMD", "TPU", "Trn", "CS", "OAI", "Ascend", "Maia"]}
HOLD_ASIC = (HOLD["Maia"] + HOLD["Ascend"] + HOLD["OAI"]) / 3

# ---------------- R1 Siliconomics 側 ----------------
NVMAP = {"bw": "GB300", "rb": "VR200", "ru": "RU", "hp": "H100"}
IT_NV = {"bw": S["it_GB"], "rb": S["it_VR"], "ru": S["it_RU"], "hp": S["it_HP"]}
sil = {}
for i, y in enumerate(YRS):
    k = S["coef"][i]; den = S["deploy_denom"][i]
    it_amd = S["it_AMD8"] if y == 2026 else S["it_AMDR"]
    row = {}
    for p, g in NVMAP.items():
        u = S[f"u_{p}"][i]
        row[p] = {"phys": u / GPG[g] * 1e6 * k, "ceq": u * IT_NV[p] * k / den, "units": u * k}
    for p, it, h in [("amd", it_amd, HOLD["AMD"]), ("tpu", S["it_TPU"], HOLD["TPU"]), ("trn", S["it_TRN"], HOLD["Trn"]), ("oth", S["it_ASIC"], HOLD_ASIC)]:
        u = S[f"u_{p}"][i]
        row[p] = {"phys": u * it / (VR * h) * k, "ceq": u * it * k / den, "units": u * k}
    nv_p = sum(row[p]["phys"] for p in NVMAP); nn_p = sum(row[p]["phys"] for p in ["amd", "tpu", "trn", "oth"])
    nv_c = sum(row[p]["ceq"] for p in NVMAP); nn_c = sum(row[p]["ceq"] for p in ["amd", "tpu", "trn", "oth"])
    row["NV"] = {"phys": nv_p, "ceq": nv_c}; row["NNV"] = {"phys": nn_p, "ceq": nn_c}
    row["TOT"] = {"phys": nv_p + nn_p, "ceq": nv_c + nn_c}
    row["resid_phys"] = S["phys"][i] - row["TOT"]["phys"]
    row["resid_ceq"] = S["deploy"][i] - row["TOT"]["ceq"]
    row["nv_units"] = sum(row[p]["units"] for p in NVMAP)
    row["nv_itcapex"] = sum(S[f"u_{p}"][i] * IT_NV[p] * k for p in NVMAP)   # $B（顆數百萬 × $K）
    row["tot_itcapex"] = row["nv_itcapex"] + sum(S[f"u_{p}"][i] * it * k for p, it in [("amd", it_amd), ("tpu", S["it_TPU"]), ("trn", S["it_TRN"]), ("oth", S["it_ASIC"])])
    sil[y] = row
R["sil"] = sil

# ---------------- R1 Wattonomics 側 ----------------
GEN_BUCKET = {"H100": "hp", "GB200": "bw", "GB300": "bw", "VR200": "rb", "RU": "ru"}
PLAT_OF_HOLD = {"TPU": "tpu", "Trn": "trn", "AMD": "amd", "Maia": "oth", "OAI": "oth", "Ascend": "oth"}
wat = {}
for i, y in enumerate(YRS):
    den = S["deploy_denom"][i]
    row = {"split": y in (2026, 2027)}
    row["s2"] = B["s2"][i]; row["x2"] = B["x2"][i]; row["rt"] = B["rt"][i]
    if row["split"]:
        j = i
        nv = W["w4"][j]; tot = W["w2"][j]
        row["tot_phys"] = tot; row["NV_phys"] = nv; row["NNV_phys"] = tot - nv
        row["NV_lo"] = W["w4_lo"][j]; row["NV_hi"] = W["w4_hi"][j]
        gen = {}
        for g in ["H100", "GB200", "GB300", "VR200", "RU"]:
            gen[g] = nv * W["mix"][g][i]
        row["gen"] = gen
        plat = {p: 0.0 for p in ["hp", "bw", "rb", "ru", "amd", "tpu", "trn", "oth"]}
        ceq = {p: 0.0 for p in plat}
        for g, v in gen.items():
            plat[GEN_BUCKET[g]] += v; ceq[GEN_BUCKET[g]] += v * CIT[g] / den
        nnvb = {}
        for b in W["buyers"]:
            x = W[f"wsh_{b}"][j] - W[f"wnv_{b}"][j]
            h = W[f"hold_{b}"]
            nnvb[b] = x
            plat[PLAT_OF_HOLD[h]] += x; ceq[PLAT_OF_HOLD[h]] += x * HOLD[h] * VR / den
        row["plat"] = plat; row["plat_ceq"] = ceq; row["nnv_by_buyer"] = nnvb
        row["NV_ceq"] = sum(ceq[p] for p in ["hp", "bw", "rb", "ru"])
        row["NNV_ceq"] = sum(ceq[p] for p in ["amd", "tpu", "trn", "oth"])
        row["tot_ceq_tk"] = row["NV_ceq"] + row["NNV_ceq"]
        row["chk_nv"] = W[f"recon_NVGW_{y}"] - nv
        row["chk_tot"] = W[f"recon_ShipGW_{y}"] - tot
        row["nv_pkgs"] = sum(gen[g] * GPG[g] for g in gen) / 1e6      # 百萬顆封裝
        row["nv_itcapex"] = sum(gen[g] * CIT[g] for g in gen)
    wat[y] = row
R["wat"] = wat

# ---------------- 並排與差距 ----------------
gap = {}
for y in YRS:
    s, w = sil[y], wat[y]
    g = {"tot_phys_F": (s["TOT"]["phys"] - w["tot_phys"]) if w["split"] else None,
         "tot_phys_s2": s["TOT"]["phys"] - w["s2"],
         "tot_ceq_x2": s["TOT"]["ceq"] - w["x2"]}
    if w["split"]:
        g["NV_phys"] = s["NV"]["phys"] - w["NV_phys"]
        g["NNV_phys"] = s["NNV"]["phys"] - w["NNV_phys"]
        g["NV_share_phys"] = g["NV_phys"] / g["tot_phys_F"]
        g["NV_share_phys_lo"] = (s["NV"]["phys"] - w["NV_hi"]) / g["tot_phys_F"]   # Wattonomics NVIDIA 用份額高值 → 差距占比低
        g["NV_share_phys_hi"] = (s["NV"]["phys"] - w["NV_lo"]) / g["tot_phys_F"]
        g["NV_ceq"] = s["NV"]["ceq"] - w["NV_ceq"]
        g["NNV_ceq"] = s["NNV"]["ceq"] - w["NNV_ceq"]
        g["tot_ceq_tk"] = s["TOT"]["ceq"] - w["tot_ceq_tk"]
        g["NV_share_ceq"] = g["NV_ceq"] / g["tot_ceq_tk"]
        g["NV_share_vs_s2"] = g["NV_phys"] / g["tot_phys_s2"]
        g["plat"] = {p: s[p]["phys"] - w["plat"][p] for p in ["hp", "bw", "rb", "ru", "amd", "tpu", "trn", "oth"]}
        # Shapley 兩因子：總量 × NVIDIA 占比
        Ts, Tw = s["TOT"]["phys"], w["tot_phys"]; hs, hw = s["NV"]["phys"] / Ts, w["NV_phys"] / Tw
        g["sh_s"], g["sh_w"] = hs, hw
        g["vol_eff"] = (Ts - Tw) * (hs + hw) / 2; g["share_eff"] = (hs - hw) * (Ts + Tw) / 2
        g["vol_eff_rng"] = sorted([(Ts - Tw) * hw, (Ts - Tw) * hs]); g["share_eff_rng"] = sorted([(hs - hw) * Ts, (hs - hw) * Tw])
    gap[y] = g
R["gap"] = gap

# ---------------- 歸因：機械可算項（Wattonomics 付款年換算重算） ----------------
def gw_player(ai, a1, a2, pit, pfo, w):
    fund = pit / (pit + pfo) * ai + pfo / (pit + pfo) * (w[0] * ai + w[1] * a1 + w[2] * a2)
    return fund / (pit + pfo)

att = {}
for i, y in enumerate(YRS):
    F7 = W["cg_F7"][i]; F8 = W["cg_F8"][i]
    # (B) NVIDIA 世代組合改用 Siliconomics 當年組合（實體 GW 權重）
    s = sil[y]
    sh = {p: s[p]["phys"] / s["NV"]["phys"] for p in NVMAP}
    F7_sil = sum(sh[p] * CIT[NVMAP[p]] for p in NVMAP)
    base = cf_mix = cf_nnv = 0.0
    det = {}
    for p, d in W["cg"].items():
        hname = {"msft": "Maia", "googl": "TPU", "amzn": "Trn"}.get(p, "AMD")
        n = W["players_nnv"][p][i]
        pit, pfo = d["pit"][i], d["pfo"][i]
        own = pfo / d["pfac"][i] if d["pfac"][i] else 1.0
        g0 = gw_player(d["ai"][i], d["a1"][i], d["a2"][i], pit, pfo, W["cg_w"])
        # 反算檢查
        det[p] = {"gw_model": d["x"][i], "gw_py": g0}
        pit_m = (1 - n) * F7_sil + n * HOLD[hname] * VR
        g1 = gw_player(d["ai"][i], d["a1"][i], d["a2"][i], pit_m, pfo, W["cg_w"])
        # (C) 換算用非 NVIDIA 占比改用對帳頁逐買方 GW 份額（只 2026–27 有）
        if y in (2026, 2027):
            n2 = 1 - W[f"wgs_{p}"][i]
            pit_n = (1 - n2) * F7 + n2 * HOLD[hname] * VR
            pfac_n = (1 - n2) * F8 + n2 * HOLD[hname] * 12.67 if False else (1 - n2) * F8 + n2 * HOLD[hname] * S["tk_in"]["TK_CapexFacility_VR200"][1]
            g2 = gw_player(d["ai"][i], d["a1"][i], d["a2"][i], pit_n, pfac_n * own, W["cg_w"])
        else:
            g2 = g0
        base += g0; cf_mix += g1; cf_nnv += g2
        det[p].update({"cf_mix": g1, "cf_nnv": g2, "nnv_capex": n, "nnv_recon": (1 - W[f"wgs_{p}"][i]) if y in (2026, 2027) else None})
    att[y] = {"F7_wat": F7, "F7_sil": F7_sil, "mix_effect": cf_mix - base, "nnv_effect": cf_nnv - base, "base7x": base, "det": det,
              "ai_sh_lo": W["cg_ai_lo"][i] - W["cg_ai_base"][i], "ai_sh_hi": W["cg_ai_hi"][i] - W["cg_ai_base"][i],
              "lag_rng_ceq": [min(B[k][i] for k in ["sx_sl", "sx_sh", "sx_il", "sx_ih"]) - B["x2"][i], max(B[k][i] for k in ["sx_sl", "sx_sh", "sx_il", "sx_ih"]) - B["x2"][i]],
              "timing_bridge_ceq": B["g2"][i], "F_vs_s2": (wat[y]["tot_phys"] - wat[y]["s2"]) if wat[y]["split"] else None}
    att[y]["lag_rng_phys"] = [v * B["rt"][i] for v in att[y]["lag_rng_ceq"]]
R["att"] = att

# 累計（年度歸屬能否解釋）：2026–29 實體
R["cum"] = {"sil_phys": sum(sil[y]["TOT"]["phys"] for y in YRS), "wat_s2": sum(wat[y]["s2"] for y in YRS)}
R["cum"]["gap"] = R["cum"]["sil_phys"] - R["cum"]["wat_s2"]

# ---------------- R2 每 GW 顆數 ----------------
gens = ["H100", "GB200", "GB300", "VR200", "RU"]
snap = lambda key, n, g: T[key]["items"][n]["values"][{"H100": "Hopper H100", "GB200": "GB200 NVL72", "GB300": "GB300 NVL72", "VR200": "VR200 NVL72", "RU": "Rubin Ultra（MGX NVL 單架，72 GPU）"}[g]]["基準"]
r2 = {}
for g in gens:
    v = {"tk_if": T["GPUsPerGW"][g][1], "sil_tkin": S["tk_in"][f"TK_GPUsPerGW_{g}"][1], "wat_tkin": W["tk_in"][f"TK_GPUsPerGW_{g}"][1],
         "snap_sil22": snap("snap_sil22", "IF_GPUsPerGW", g), "snap_wat22": snap("snap_wat22", "IF_GPUsPerGW", g),
         "snap_217": snap("snap_sil217", "IF_GPUsPerGW", g),
         "cb": {"GB300": C["gpus"][1], "VR200": C["gpus"][4]}.get(g),
         "racks_tk": T["RacksPerGW"][g][1], "pkg": T["pkg_per_rack"][g], "kw": T["rack_kw"][g]}
    # 由 TK 規格重算：FLOOR(1000 MW ×(1−8%)／機架 kW)×封裝數
    rk = lambda kw: math.floor(T["it_gw"] * 1000 * (1 - T["offrack"]) * 1000 / kw)
    v["recalc_base"] = rk(v["kw"][1]) * v["pkg"]
    v["pow_lo_cnt"] = rk(v["kw"][0]) * v["pkg"]; v["pow_hi_cnt"] = rk(v["kw"][2]) * v["pkg"]
    v["max_rel"] = max(abs(x / v["tk_if"] - 1) for x in [v["sil_tkin"], v["wat_tkin"], v["snap_sil22"], v["snap_wat22"], v["snap_217"]] + ([v["cb"]] if v["cb"] else []))
    v["kw_per_pkg_it"] = 1e6 / v["tk_if"]       # 每顆所攤 IT kW（含機架外 8%）
    r2[g] = v
R["r2_nv"] = r2
# 非 NVIDIA：Siliconomics 隱含 vs Tokenomics 證據
ev_tpu = E["C18"]["vals"][1]; ev_trn = E["D15"]["vals"][1]
rack_it_mw = E["A7"]["vals"][1]
amd_kw_rack = 245.0   # TK docs/evidence/20261025_NonNV_AMD README：Helios 依負載 225–245 kW（AMD）；與 VR200 227 kW 同口徑比較取 245
amd_lo_kw = 225.0
ev_amd = rack_it_mw * 1000 / amd_kw_rack * 72
ev_amd_hi = rack_it_mw * 1000 / amd_lo_kw * 72
oai_kw = 1.26   # TK docs/evidence/20261025_NonNV_Cerebras_OpenAI：Jalapeño 每顆所攤 1.26 kW（r1；由 Broadcom 說法反推，口徑不明）
ev_oai = rack_it_mw * 1000 / oai_kw
nn = {}
for p, it, h, evv, evlab in [("tpu", S["it_TPU"], HOLD["TPU"], ev_tpu, "TK 證據 NonNV_TPU_Trainium 推導 C18（1.085 kW／顆）"),
                             ("trn", S["it_TRN"], HOLD["Trn"], ev_trn, "TK 證據 NonNV_TPU_Trainium 推導 D15（72 顆／架、9,889 架／GW）"),
                             ("amd", S["it_AMDR"], HOLD["AMD"], ev_amd, "TK 證據 20261025_NonNV_AMD README：Helios 245 kW／72 顆，以 VR200 機架 IT 919.8 MW 換算"),
                             ("oth", S["it_ASIC"], HOLD_ASIC, ev_oai, "TK 證據 20261025_NonNV_Cerebras_OpenAI：Jalapeño 每顆約 1.26 kW（只代表 OpenAI 客製）")]:
    sil_imp = h * VR * 1e9 / (it * 1e3)
    nn[p] = {"sil_implied": sil_imp, "tk_ev": evv, "rel": sil_imp / evv - 1, "hold": h, "it": it, "src": evlab}
nn["amd"]["tk_ev_225"] = ev_amd_hi
nn["tpu"]["tk_ev_lohi_power"] = [rack_it_mw * 1000 / E["C4"]["vals"][0], rack_it_mw * 1000 / E["C4"]["vals"][2]]
R["r2_nn"] = nn
# 若用 TK 證據每 GW 顆數，Siliconomics 非 NVIDIA 實體 GW 會變多少（2027）
alt = {}
for i, y in enumerate(YRS):
    k = S["coef"][i]
    alt[y] = {p: S[f"u_{p}"][i] * k * 1e6 / nn[p]["tk_ev"] for p in ["tpu", "trn", "amd", "oth"]}
    alt[y]["sum"] = sum(alt[y][p] for p in ["tpu", "trn", "amd", "oth"])
R["r2_alt_nnv_phys"] = alt
# 持有比 → IT 資本支出比（TK 公式 持有比＝1 − s×(1 − k)）
s_ = E["A20"]["vals"][1]
R["r2_hold_to_k"] = {h: 1 - (1 - HOLD[h]) / s_ for h in HOLD}
R["r2_hold_s"] = s_
# chip_bridge v21.7 vs v22.0：差異格數
def flat(x, p=""):
    if isinstance(x, dict):
        for kk, vv in x.items():
            yield from flat(vv, p + "/" + kk)
    elif isinstance(x, list):
        for ii, vv in enumerate(x):
            yield from flat(vv, p + f"[{ii}]")
    else:
        yield p, x
a = T["snap_sil217"]["items"]; b = T["snap_sil22"]["items"]
cb_names = ["IF_RacksPerGW", "IF_GPUsPerGW", "IF_CapexIT", "IF_CapexFacility", "IF_CapexTotal", "IF_RackPrice", "IF_ScaleOutShare", "IF_StoragePerGPU"]
diffs_all = []; diffs_cb = []; ncells_cb = 0
for n in b:
    fa = dict(flat(a[n]["values"])); fb = dict(flat(b[n]["values"]))
    for kk in fb:
        if n in cb_names:
            ncells_cb += 1
        if fa.get(kk) != fb[kk]:
            diffs_all.append((n, kk, fa.get(kk), fb[kk]))
            if n in cb_names:
                diffs_cb.append((n, kk))
R["r2_cb"] = {"diffs_all": diffs_all, "diffs_cb": diffs_cb, "ncells_cb": ncells_cb,
              "dist_vs_v22": {k: [C[k][j] - [T["snap_sil22"]["items"][nmap]["values"][gg]["低成本" if jj == 0 else "基準" if jj == 1 else "高成本"]
                                              for gg in ["GB300 NVL72", "VR200 NVL72"] for jj in range(3)][j] for j in range(6)]
                              for k, nmap in [("gpus", "IF_GPUsPerGW"), ("racks", "IF_RacksPerGW"), ("capit", "IF_CapexIT"), ("capfac", "IF_CapexFacility"), ("captot", "IF_CapexTotal")]}}

# Wattonomics 隱含 NVIDIA 封裝數 vs Siliconomics
R["nv_pkgs"] = {y: {"sil": sil[y]["nv_units"], "wat": wat[y].get("nv_pkgs")} for y in YRS}

json.dump(R, open(sys.argv[2] if len(sys.argv) > 2 else "results.json", "w"), ensure_ascii=False, indent=1)

# ------------- 印出摘要 -------------
def f(x, n=2):
    return "—" if x is None else f"{x:.{n}f}"
for y in YRS:
    s, w, g = sil[y], wat[y], gap[y]
    print(y, "SIL phys", f(s["TOT"]["phys"]), "NV", f(s["NV"]["phys"]), "NNV", f(s["NNV"]["phys"]), "resid", f(s["resid_phys"], 6), "| ceq", f(s["TOT"]["ceq"]), "NV", f(s["NV"]["ceq"]), "resid", f(s["resid_ceq"], 6))
    print("   plat phys", {p: round(s[p]["phys"], 3) for p in ["bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth"]})
    if w["split"]:
        print("   WAT phys", f(w["tot_phys"]), "NV", f(w["NV_phys"]), "NNV", f(w["NNV_phys"]), "chk", f(w["chk_nv"], 9), f(w["chk_tot"], 9), "ceq_tk", f(w["tot_ceq_tk"]), "NVceq", f(w["NV_ceq"]))
        print("   WAT plat", {p: round(v, 3) for p, v in w["plat"].items()})
        print("   gap", {k: (round(v, 3) if isinstance(v, float) else v) for k, v in g.items() if k != "plat"})
        print("   gap plat", {p: round(v, 3) for p, v in g["plat"].items()})
    print("   s2", f(w["s2"]), "x2", f(w["x2"]), "gap s2", f(g["tot_phys_s2"]), "gap x2", f(g["tot_ceq_x2"]))
    a_ = att[y]; print("   att mix", f(a_["mix_effect"], 3), "nnv", f(a_["nnv_effect"], 3), "F7", f(a_["F7_wat"]), f(a_["F7_sil"]), "ai_sh", f(a_["ai_sh_lo"]), f(a_["ai_sh_hi"]), "lag phys", [round(v, 3) for v in a_["lag_rng_phys"]], "F-s2", f(a_["F_vs_s2"]))
    print("   chk player gw", {p: round(d["gw_model"] - d["gw_py"], 9) for p, d in a_["det"].items()})
print("cum", R["cum"])
print("nv pkgs", R["nv_pkgs"])
print("r2 nv", {g: (v["tk_if"], v["recalc_base"], v["max_rel"], v["pow_lo_cnt"], v["pow_hi_cnt"]) for g, v in r2.items()})
print("r2 nn", {p: (round(v["sil_implied"]), round(v["tk_ev"]), round(v["rel"], 3)) for p, v in nn.items()})
print("alt nnv", {y: {k: round(v, 2) for k, v in alt[y].items()} for y in YRS})
print("hold->k", R["r2_hold_to_k"])
print("cb diffs", R["r2_cb"]["diffs_all"], R["r2_cb"]["diffs_cb"], R["r2_cb"]["ncells_cb"], R["r2_cb"]["dist_vs_v22"])

import json, copy, sys
sys.path.insert(0, "/home/claude/out/sens/scripts")
import replica as R
U = json.load(open("/home/claude/out/sens/scripts/uno_results.json"))
def P(): return R.base_params()
def scale_ai(p, amzn, oth):
    p["mAI"] = R.aish_mult(oth, amzn); return p
def _old(p, amzn, oth):
    for q in R.PL:
        if q == "xai": continue
        f = amzn if q == "amzn" else oth
        p["mAI"][q] = [f] * 5
    return p
sc = {}
p = P(); sc["ai_hi"] = scale_ai(p, 0.75 / 0.68, 0.95 / 0.85)
p = P(); sc["ai_lo"] = scale_ai(p, 0.60 / 0.68, 0.75 / 0.85)
for nm, kv in [("lag_on_lo", {"lag_on": 3}), ("lag_on_hi", {"lag_on": 6}), ("inv_lo", {"inv_on": 0.05}), ("inv_hi", {"inv_on": 0.2}),
               ("bldg_lo", {"lag_bldg": 12}), ("bldg_hi", {"lag_bldg": 24}), ("oth_lo", {"oth_gmult": 0}), ("oth_hi", {"oth_gmult": 1.5})]:
    p = P(); p.update(kv); sc[nm] = p
p = P(); p["CIT"]["VR200"] *= 227 / 188; p["CIT"]["RU"] *= 340 / 300; p["CIT"]["GB300"] *= 136 / 132; sc["pow_lo_w"] = p
p = P(); tk = R.I["tk"]
for g in ["GB300", "VR200", "RU"]: p["CIT"][g] = tk["CapexIT"][g][0]; p["FAC"][g] = tk["CapexFacility"][g][0]
sc["cost_lo_w"] = p
rows = []
for nm, p in sc.items():
    key = nm.replace("_w", "")
    d = R.demand(p)
    u = U["wat"][key]
    rows.append((nm, d["wt27"] - u["WIF_MktGW_Add_2027"][0], d["wt28"] - u["WIF_MktGW_Add_2028"][0]))
for r in rows: print("WAT", r[0], f"Δwt27={r[1]:+.5f} Δwt28={r[2]:+.5f}")
# Siliconomics
p = P(); p["raw"]["rb"] = 6.7; print("SIL rb_67", R.supply(p)["S0"] - U["sil"]["rb_67"]["SIF_DeployGW_Base"][1], R.supply(p)["phys"] - U["sil"]["rb_67"]["SIF_DeployGWPhys_Base"][1])
p = P(); [p["raw"].__setitem__(k, p["raw"][k] * 0.884) for k in ("bw", "rb", "ru")]; s = R.supply(p); print("SIL nv0884", s["S0"] - U["sil"]["nv_x0884"]["SIF_DeployGW_Base"][1], s["phys"] - U["sil"]["nv_x0884"]["SIF_DeployGWPhys_Base"][1])
p = P(); p["CIT"]["VR200"] *= 227 / 188; p["CIT"]["RU"] *= 340 / 300; p["CIT"]["GB300"] *= 136 / 132; p["GPG"]["VR200"] *= 227 / 188; p["GPG"]["RU"] *= 340 / 300; p["GPG"]["GB300"] *= 136 / 132
s = R.supply(p); print("SIL pow_lo", s["S0"] - U["sil"]["pow_lo"]["SIF_DeployGW_Base"][1], s["phys"] - U["sil"]["pow_lo"]["SIF_DeployGWPhys_Base"][1])
p = P(); tk = R.I["tk"]
for g in ["GB300", "VR200", "RU"]: p["CIT"][g] = tk["CapexIT"][g][0]
s = R.supply(p); print("SIL cost_lo", s["S0"] - U["sil"]["cost_lo"]["SIF_DeployGW_Base"][1], s["phys"] - U["sil"]["cost_lo"]["SIF_DeployGWPhys_Base"][1])
p = P(); p["it_nn"]["tpu"] = 35; s = R.supply(p); print("SIL it_tpu", s["S0"] - U["sil"]["it_tpu_hi"]["SIF_DeployGW_Base"][1], s["phys"] - U["sil"]["it_tpu_hi"]["SIF_DeployGWPhys_Base"][1])

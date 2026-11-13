"""CoreWeave v5.0 反向 DCF：整本重算反解 k_adj，使加權目標價＝現價、＝升一級邊界、＝降一級邊界 → data/reverse_v50.json。
方法：先在 k_adj 0.5–3.0 的格點整本重算，找出夾住目標的兩點，再在其間細分重算一次，線性內插。"""
import datetime, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import runner_v50 as RN

JOBS = int(sys.argv[sys.argv.index("--jobs") + 1]) if "--jobs" in sys.argv else 2
grid = [round(0.5 + 0.1 * i, 3) for i in range(26)]
res = RN.run([("base", {}, {})] + [(f"k{k}", {"k_adj": k}, {}) for k in grid], JOBS)
b = res["base"]
px = None
import openpyxl  # 現價取自基準檔 In 頁
rg = RN.in_ranges(); px = rg["px"][1]
up = b["tp"] / px - 1
bounds = [0.5, 0.25, -0.25, -0.5]
upb = min([x for x in bounds if x > up], default=None); dnb = max([x for x in bounds if x <= up], default=None)
targets = [("px", "加權目標價＝現價", px)]
if upb is not None: targets.append(("up", f"加權目標價＝升一級邊界（上檔 {upb:+.0%}）", px * (1 + upb)))
if dnb is not None: targets.append(("dn", f"加權目標價＝降一級邊界（上檔 {dnb:+.0%}）", px * (1 + dnb)))
pts = sorted((k, res[f"k{k}"]["tp"]) for k in grid)
def bracket(t):
    for (k0, v0), (k1, v1) in zip(pts, pts[1:]):
        if (v0 - t) * (v1 - t) <= 0: return k0, k1
    return None
fine = {}
for tid, lab, t in targets:
    br = bracket(t)
    if br: fine[tid] = [round(br[0] + (br[1] - br[0]) * j / 5, 4) for j in range(1, 5)]
res2 = RN.run([(f"{tid}_{k}", {"k_adj": k}, {}) for tid, ks in fine.items() for k in ks], JOBS) if fine else {}
solves = []
for tid, lab, t in targets:
    if tid not in fine:
        solves.append({"id": tid, "label": lab + "（在 k_adj 0.5–3.0 內無解）", "key": "k_adj", "target": t, "value": None, "pm30": None}); continue
    br = bracket(t)
    p2 = sorted([br[0], br[1]] + fine[tid])
    vals = [(k, res[f"k{k}"]["tp"] if f"k{k}" in res else res2[f"{tid}_{k}"]["tp"]) for k in p2]
    for (k0, v0), (k1, v1) in zip(vals, vals[1:]):
        if (v0 - t) * (v1 - t) <= 0:
            k = k0 + (t - v0) * (k1 - k0) / (v1 - v0) if v1 != v0 else k0; break
    solves.append({"id": tid, "label": lab, "key": "k_adj", "target": t, "value": k, "pm30": b["pm30"] * k})
out = {"date": datetime.date.today().isoformat(), "base_tp": b["tp"], "px": px, "pm30_base": b["pm30"], "solves": solves,
       "grid": [[k, v] for k, v in pts]}
json.dump(out, open(os.path.join(RN.HERE, "data", "reverse_v50.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(solves, ensure_ascii=False))

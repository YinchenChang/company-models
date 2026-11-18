# 讀一個情境複本的重算後活頁簿與診斷檔 → json。用法：python3 metrics.py <情境名>
import json, sys, glob, os, openpyxl
N = sys.argv[1]; D = f"/tmp/pp_{N}"
F = open(f"{D}/siliconomics/model/CURRENT").read().strip()
wb = openpyxl.load_workbook(f"{D}/siliconomics/model/{F}", data_only=True)
def nm(n):
    sh, ref = list(wb.defined_names[n].destinations)[0]
    r = wb[sh][ref.replace("$", "")]
    if isinstance(r, tuple):
        return [c.value for row in r for c in (row if isinstance(row, tuple) else (row,))]
    return r.value
YR = [2026, 2027, 2028, 2029]
co = wb["Coord"]; rows = {co.cell(r, 1).value: r for r in range(1, co.max_row + 1)}
cv = lambda k: [co.cell(rows[k], c).value for c in range(3, 7)]
out = {"file": F}
out["base"] = nm("SIF_DeployGW_Base")[-4:]
out["phys"] = nm("SIF_DeployGWPhys_Base")[-4:]
for k in ("sup", "dem", "st", "lab", "k", "alt", "altr", "es", "ew", "cur"):
    out[k] = cv(k)
# 顆數（截頂後、入基準鏈前）
inn = wb["In"]; ir = {inn.cell(r, 1).value: r for r in range(1, inn.max_row + 1)}
u = lambda k: inn.cell(ir[k], 5).value
out["units"] = {y: {p: u(f"u_{p}_{y%100}") for p in ("bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth")} for y in (2026, 2027)}
for y in (2026, 2027):
    out["units"][y]["NV"] = sum(out["units"][y][p] for p in ("bw", "rb", "ru", "hp"))
sup27 = u("hbm_sup26") * (1 + u("hbm_sup27")); g = {p: u(f"g_{p}_27") for p in ("bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth")}
dem27 = sum(out["units"][2027][p] * g[p] for p in g) * u("hbm_gross") / 1000
out["hbm27"] = {"supply_EB": sup27, "demand_EB_after_units": dem27, "util": dem27 / sup27}
for n in ("SIF_HBM_ShipEB_Base", "SIF_HBM_Mkt_Base", "SIF_HBM_GapEB_Base", "SIF_HBMUtil_Base", "SIF_UnitsCoef_Base"):
    out[n] = nm(n)
out["SGOV_Errors"] = nm("SGOV_Errors"); out["SGOV_Warn"] = nm("SGOV_Warn")
rc = wb["Recon"]
out["reconF"] = {"implied_rev": [rc.cell(88, 4).value, rc.cell(88, 5).value], "act_h1": rc.cell(89, 4).value,
                 "main_reading": rc.cell(92, 4).value, "second_reading": rc.cell(93, 4).value, "status": rc.cell(95, 4).value,
                 "company_ref": [rc.cell(96, 4).value, rc.cell(96, 5).value], "ratio_ref": [rc.cell(97, 4).value, rc.cell(97, 5).value],
                 "seasonal": [rc.cell(r, 4).value for r in (109, 110, 111)]}
# 錯誤格
err = 0
for ws in wb:
    for row in ws.iter_rows(values_only=True):
        for v in row:
            if isinstance(v, str) and v.startswith("#") and v.rstrip("!?/0").upper() in ("#REF", "#DIV", "#VALUE", "#NAME", "#N/A", "#NUM", "#NULL") or (isinstance(v, str) and v in ("#DIV/0!", "#N/A", "#NAME?", "#REF!", "#VALUE!", "#NUM!", "Err:502", "Err:504")):
                err += 1
out["err_cells"] = err
# 診斷：收斂換算值
try:
    b = openpyxl.load_workbook(f"{D}/build/bridge/bridge_raw.xlsx", data_only=True)["Bridge"]
    br = {b.cell(r, 1).value: r for r in range(1, b.max_row + 1)}
    out["bridge"] = {k: [b.cell(br[k], c).value for c in range(5, 9)] for k in ("x2", "rt", "s0", "g3", "s2")}
except Exception as e:
    out["bridge"] = str(e)
try:
    e = openpyxl.load_workbook(f"{D}/build/ev/ev_raw.xlsx", data_only=True)["Calc"]
    out["ev"] = {k: [e.cell(r, c).value for c in range(3, 7)] for k, r in (("den", 10), ("es", 41), ("lo", 49), ("hi", 50), ("st50", 56))}
except Exception as x:
    out["ev"] = str(x)
json.dump(out, open(f"/home/claude/out/proposal/runs/m_{N}.json", "w"), ensure_ascii=False, indent=1, default=str)
print(N, "base", [round(x, 3) for x in out["base"]], "k", out["k"], "st", out["st"], "err", err, "SGOV", out["SGOV_Errors"])

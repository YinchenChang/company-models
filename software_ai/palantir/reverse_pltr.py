# 主導變數整本重算反解（工作單 G2-9）：二分法找使加權目標價＝現價所需的值；每一步都是整本重算（build_palantir.py --set＋LibreOffice）。
# 用法：python3 software_ai/palantir/reverse_pltr.py <工作資料夾> <輸出 reverse_solve.json>
import json, os, sys
sys.argv = [sys.argv[0]] + sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
WD, OUT = sys.argv[1], sys.argv[2]
sys.path.insert(0, HERE)
os.makedirs(WD, exist_ok=True)
import importlib.util, subprocess, shutil, tempfile
import openpyxl
BUILD = os.path.join(HERE, "build_palantir.py")
n = [0]
def tp(sets):
    n[0] += 1; name = f"r{n[0]:03d}"
    raw = os.path.join(WD, name + ".xlsx")
    env = dict(os.environ); env["ROWMAP"] = os.path.join(WD, name + ".rm.json")
    subprocess.run([sys.executable, BUILD, raw] + sum((["--set", f"{k}={v}"] for k, v in sets), []), check=True, stdout=subprocess.DEVNULL, env=env)
    out = os.path.join(WD, "out"); os.makedirs(out, exist_ok=True)
    prof = tempfile.mkdtemp(prefix="lo_rev_"); e2 = dict(os.environ); e2["LC_ALL"] = "C.UTF-8"
    for _try in range(4):   # LibreOffice 首次啟動偶發 255，重試
        p_ = subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, raw], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=e2, timeout=900)
        if p_.returncode == 0 and os.path.exists(os.path.join(out, name + ".xlsx")): break
    else: raise SystemExit("soffice 重算失敗")
    shutil.rmtree(prof, ignore_errors=True)
    rm = json.load(open(env["ROWMAP"], encoding="utf-8"))
    wb = openpyxl.load_workbook(os.path.join(out, name + ".xlsx"), data_only=True)
    return wb["Val"].cell(rm["VR"]["tp"], 2).value, wb["Val"].cell(rm["VR"]["px"], 2).value, wb["In"].cell(rm["IN_ROW"]["wl27_B"], 5).value
base_tp, px, wlb = tp([])
def solve(name, mk, lo, hi, incr, iters=14):
    flo, fhi = tp(mk(lo))[0] - px, tp(mk(hi))[0] - px
    if (flo > 0) == (fhi > 0): return {"value": None, "tp": None, "note": f"搜尋範圍 {lo}～{hi} 內無解（端點目標價 {tp(mk(lo))[0]:.2f}／{tp(mk(hi))[0]:.2f}，現價 {px:.2f}）"}
    for _ in range(iters):
        mid = (lo + hi) / 2; fm = tp(mk(mid))[0] - px
        if (fm > 0) == (flo > 0): lo, flo = mid, fm
        else: hi, fhi = mid, fm
    v = (lo + hi) / 2
    return {"value": v + incr, "tp": tp(mk(v))[0], "note": None}
res = {"wacc_adj": solve("wacc_adj", lambda v: [("wacc_adj", v)], -0.085, 0.0, 0.0),
       "wl27": solve("wl27", lambda d: [("wl27_A", 0.42 + d), ("wl27_B", 0.32 + d), ("wl27_C", 0.20 + d)], 0.0, 6.0, 0.32)}
res["_base"] = {"tp": base_tp, "px": px}
json.dump(res, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(res)

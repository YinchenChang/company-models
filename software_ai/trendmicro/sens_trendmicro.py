# 敏感度與反解：一次改一個 In 輸入，整本以 LibreOffice 重算，讀 Val 加權目標價。
# 用法：python3 sens_trendmicro.py <工作資料夾> <輸出 sens.json>（JOBS＝平行數，預設 4；KREQ=0 略過反解）
import json, os, subprocess, sys, shutil, tempfile
from concurrent.futures import ThreadPoolExecutor
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__)); WD, OUT = sys.argv[1], sys.argv[2]; os.makedirs(WD, exist_ok=True)
BUILD = os.path.join(HERE, "build_trendmicro.py")
def run(name, sets):
    raw = os.path.join(WD, name + ".xlsx"); rmf = os.path.join(WD, name + ".rowmap.json")
    env = dict(os.environ, ROWMAP=rmf)
    subprocess.run([sys.executable, BUILD, raw] + sum((["--set", f"{k}={v}"] for k, v in sets), []), check=True, stdout=subprocess.DEVNULL, env=env)
    out = os.path.join(WD, "out_" + name); prof = tempfile.mkdtemp(prefix="lo_tm_")
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, raw],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"), timeout=600)
    rm = json.load(open(rmf)); wb = openpyxl.load_workbook(os.path.join(out, name + ".xlsx"), data_only=True)
    V = wb["Val"]; vr = rm["VR"]
    o = {"tp": V.cell(vr["tp"], 2).value, "tpdcf": V.cell(vr["tpdcf"], 2).value, "rec": V.cell(vr["rec"], 2).value, "gate": V.cell(vr["gate"], 2).value,
         "fp": wb["Sens"].cell(rm["SE"]["FPR"] + 1, 4).value}
    shutil.rmtree(out, ignore_errors=True); shutil.rmtree(prof, ignore_errors=True)
    for f in (raw, rmf): os.remove(f)
    return o
base = run("base", []); res = {"base|base": base}
rm0 = json.load(open(os.path.join(HERE, "data", "rowmap.json")))
jobs = {f"{k}|{sd}": st for k, sd, st in rm0["CASES"]}
with ThreadPoolExecutor(max_workers=int(os.environ.get("JOBS", "4"))) as ex:
    futs = {k: ex.submit(run, "c%02d" % i, st) for i, (k, st) in enumerate(jobs.items())}
    for k, f in futs.items():
        res[k] = {**f.result(), "set": jobs[k]}; print(k, round(res[k]["tp"], 1), flush=True)
kreq = {}
if os.environ.get("KREQ", "1") == "1":
    px = 7214
    BND = {"s_b26": (-0.2, 0.6), "tpm": (0.0, 40.0), "fx": (80.0, 300.0)}
    for var, (a, b) in BND.items():
        kreq[var] = {}
        for tag, mult in (("px", 1.0), ("up", 1.5), ("down", 1.25)):
            tgt = px * mult; lo, hi = a, b
            flo = run("k_lo", [(var, lo)])["tp"] - tgt; fhi = run("k_hi", [(var, hi)])["tp"] - tgt
            if flo * fhi > 0: kreq[var][tag] = None; continue
            for _ in range(14):
                mid = (lo + hi) / 2; fm = run("k_m", [(var, mid)])["tp"] - tgt
                if fm * flo <= 0: hi = mid
                else: lo, flo = mid, fm
            kreq[var][tag] = (lo + hi) / 2; print("kreq", var, tag, kreq[var][tag], flush=True)
res["kreq"] = kreq
json.dump(res, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

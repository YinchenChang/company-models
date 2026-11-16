#!/usr/bin/env python3
"""HubSpot 敏感度與反解：一次改一個 In 輸入（低／高，另加聯動與路由案），整本以 LibreOffice 重算，讀 Val 頁加權目標價。
用法：python3 sens_hubspot.py <工作資料夾> [--kreq]（環境變數 JOBS＝平行數，預設 4）
輸出：data/sens_v01.json（敏感度）；--kreq 時另輸出 data/kreq_v01.json（主導變數反解：非 GAAP 營業費用率 ng_opex，使加權目標價＝現價；另解翻轉評等所需值）。
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor

import openpyxl

import build_hubspot as B

HERE = os.path.dirname(os.path.abspath(__file__))
WD = sys.argv[1]
os.makedirs(WD, exist_ok=True)


def recalc(raw):
    out = os.path.join(WD, "out")
    os.makedirs(out, exist_ok=True)
    prof = tempfile.mkdtemp(prefix="lo_hs_")
    env = dict(os.environ)
    env["LC_ALL"] = "C.UTF-8"
    for attempt in range(4):
        rc = subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, raw],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env, timeout=900).returncode
        if rc == 0:
            break
    else:
        raise RuntimeError(f"soffice 失敗：{raw}")
    shutil.rmtree(prof, ignore_errors=True)
    return os.path.join(out, os.path.basename(raw))


def run(name, sets):
    raw = os.path.join(WD, f"{name}.xlsx")
    os.environ["ROWMAP"] = os.path.join(WD, f"{name}.rowmap.json")
    B.build(raw, ov=dict(sets))
    wb = openpyxl.load_workbook(recalc(raw), data_only=True)
    v = wb["Val"]
    return {"tp": v["B8"].value, "dcf": v["B6"].value, "cmp": v["B7"].value, "m27": wb["PL"][f"G{B.PB['B']['ngopm']}"].value}


def sens():
    base = run("base", [])
    jobs = B.case_jobs()
    res = {}
    with ThreadPoolExecutor(max_workers=int(os.environ.get("JOBS", "4"))) as ex:
        futs = [(j, ex.submit(run, f"c{i:03d}", j[4])) for i, j in enumerate(jobs)]
        cases = []
        for (key, side, label, setting, sets), f in futs:
            o = f.result()
            cases.append({"key": key, "side": side, "label": label, "setting": setting, "tp": o["tp"], "dcf": o["dcf"], "cmp": o["cmp"]})
            print(key, side, round(o["tp"], 2), flush=True)
    json.dump({"base": base, "cases": cases}, open(os.path.join(HERE, "data", "sens_v01.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return base


def solve(key, target, x0, x1, tol=0.005):
    """割線法：找 key 使加權目標價＝target"""
    f0, f1 = run("k0", [(key, x0)])["tp"] - target, run("k1", [(key, x1)])["tp"] - target
    for i in range(12):
        x2 = x1 - f1 * (x1 - x0) / (f1 - f0)
        o = run("k2", [(key, x2)])
        f2 = o["tp"] - target
        print("solve", key, round(x2, 5), round(o["tp"], 3), flush=True)
        if abs(f2) < tol:
            return x2, o
        x0, f0, x1, f1 = x1, f1, x2, f2
    return x2, o


def kreq():
    wb = openpyxl.load_workbook(os.path.join(WD, "out", "base.xlsx"), data_only=True)
    ing = wb["In"]
    base = next(ing.cell(r, 5).value for r in range(4, ing.max_row) if ing.cell(r, 1).value == "inc_m")
    px = next(ing.cell(r, 5).value for r in range(4, ing.max_row) if ing.cell(r, 1).value == "px")
    x, o = solve("inc_m", px, base, base + 0.1)
    y, o2 = solve("inc_m", 0.75 * px, base, base - 0.05)
    z, o3 = solve("inc_m", 1.25 * px, base, base + 0.1)
    json.dump({"key": "inc_m", "base": base, "required": x, "margin27": o["m27"], "tp": o["tp"], "flip_m25": y, "flip_tp": o2["tp"], "flip_p25": z, "flip_p25_tp": o3["tp"], "px": px},
              open(os.path.join(HERE, "data", "kreq_v01.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)


if __name__ == "__main__":
    b = sens()
    print("base", b)
    if "--kreq" in sys.argv:
        kreq()

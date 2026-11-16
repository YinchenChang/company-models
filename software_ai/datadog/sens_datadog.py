#!/usr/bin/env python3
"""Datadog 個股模型 v0.1 敏感度：In 頁每個有低／高區間的鍵一次改一個，整本以 LibreOffice 重算，記錄加權目標價；另做主導變數（rel_m）反解。
用法：python3 software_ai/datadog/sens_datadog.py <工作目錄> <輸出 sens.json> [JOBS]"""
import json, os, subprocess, sys, concurrent.futures as cf
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WD, OUT = sys.argv[1], sys.argv[2]
JOBS = int(sys.argv[3]) if len(sys.argv) > 3 else 3
os.makedirs(WD, exist_ok=True)
rm_path = os.path.join(WD, "rowmap_base.json"); base_raw = os.path.join(WD, "base_raw.xlsx")
subprocess.run([sys.executable, os.path.join(HERE, "build_datadog.py"), base_raw], check=True, env=dict(os.environ, ROWMAP=rm_path), stdout=subprocess.DEVNULL)
rm = json.load(open(rm_path, encoding="utf-8")); VR = rm["VR"]


def recalc(raw, tag):
    od = os.path.join(WD, "o_" + tag); os.makedirs(od, exist_ok=True)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{WD}/lo_{tag}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", od, raw],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"))
    return os.path.join(od, os.path.basename(raw))


def read(x):
    wb = openpyxl.load_workbook(x, data_only=True); V = wb["Val"]
    return {"tp": V.cell(VR["tp"], 2).value, "dcf": V.cell(VR["tpdcf"], 2).value, "gate": V.cell(VR["gate_in"], 2).value, "rec": V.cell(VR["rec"], 2).value, "fp": wb["Sens"].cell(rm["SE_FP"], 4).value}


def run(tag, sets):
    raw = os.path.join(WD, f"c_{tag}.xlsx")
    cmd = [sys.executable, os.path.join(HERE, "build_datadog.py"), raw] + sum((["--set", f"{k}={v}"] for k, v in sets), [])
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL)
    return read(recalc(raw, tag))


def parse(vv):
    return [(p.split("=")[0], float(p.split("=")[1])) for p in vv.split(";")]


base = read(recalc(base_raw, "base"))
cases = rm["SENS_DEF"]
res = {}
def job(c):
    k, side, v = c
    tag = (k + "_" + side).replace("@", "x").replace("|", "_")
    sets = parse(v) if k.startswith("@") else [(k, float(v))]
    return f"{k}|{side}", run(tag, sets)
with cf.ThreadPoolExecutor(JOBS) as ex:
    for key, r in ex.map(job, cases):
        res[key] = r; print(key, round(r["tp"], 2), flush=True)
# 反解：rel_m（相對倍數調整係數；DCF 一腿無法由單一營運變數支持現價，故解 Comparable 需要的相對倍數）
x0, x1 = 1.0, 2.5
f0 = base["tp"] - 0.0
px = openpyxl.load_workbook(os.path.join(WD, "o_base", "base_raw.xlsx"), data_only=True)["In"]
pxv = None
for row in px.iter_rows(min_row=4):
    if row[0].value == "px": pxv = row[4].value
ft = lambda x: run(f"rs_{x:.4f}".replace(".", "p"), [("rel_m", x)])["tp"] - pxv
y0, y1 = ft(x0), ft(x1)
for _ in range(4):
    x2 = x1 - y1 * (x1 - x0) / (y1 - y0)
    x0, y0, x1 = x1, y1, x2
    y1 = ft(x1)
    if abs(y1) < 0.005: break
rev = {"x": x1, "tp": y1 + pxv, "note": "sens_datadog.py 以割線法整本重算 rel_m（Comparable 相對同業倍數調整係數），使加權目標價＝現價；DCF 一腿（股票薪酬視為費用）無單一營運變數可支持現價"}
json.dump({"base_tp": base["tp"], "fp": base["fp"], "cases": res, "reverse": rev}, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("base", base["tp"], "reverse", rev)

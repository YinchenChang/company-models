#!/usr/bin/env python3
"""LY 個股模型 v0.1 敏感度：In 頁每個有低／高區間的鍵一次改一個，整本以 LibreOffice 重算，記錄加權目標價。
用法：python3 software_ai/lycorp/sens_lycorp.py <工作目錄> <輸出 sens.json> [JOBS]
"""
import json, os, subprocess, sys, tempfile, concurrent.futures as cf
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WD, OUT = sys.argv[1], sys.argv[2]
JOBS = int(sys.argv[3]) if len(sys.argv) > 3 else 3
os.makedirs(WD, exist_ok=True)
rm_path = os.path.join(WD, "rowmap_base.json")
base_raw = os.path.join(WD, "base_raw.xlsx")
env = dict(os.environ, ROWMAP=rm_path)
subprocess.run([sys.executable, os.path.join(HERE, "build_lycorp.py"), base_raw], check=True, env=env, stdout=subprocess.DEVNULL)
rm = json.load(open(rm_path, encoding="utf-8"))
VR = rm["VR"]


def recalc(raw, tag):
    od = os.path.join(WD, "o_" + tag)
    os.makedirs(od, exist_ok=True)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{WD}/lo_{tag}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", od, raw],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"))
    return os.path.join(od, os.path.basename(raw))


def read(x):
    wb = openpyxl.load_workbook(x, data_only=True)
    V = wb["Val"]
    return {"tp": V.cell(VR["tp"], 2).value, "dcf": V.cell(VR["tpdcf"], 2).value, "gate": V.cell(VR["fcf_gate"], 2).value, "rec": V.cell(VR["rec"], 2).value}


def run(case):
    k, side, v = case
    tag = f"{k}_{side}".replace("|", "_").replace("@", "c_")
    raw = os.path.join(WD, f"c_{tag}.xlsx")
    sets = []
    for kv in (v.split(";") if k.startswith("@") else [f"{k}={v}"]):
        sets += ["--set", kv]
    subprocess.run([sys.executable, os.path.join(HERE, "build_lycorp.py"), raw] + sets, check=True, stdout=subprocess.DEVNULL)
    return f"{k}|{side}", read(recalc(raw, tag))


bx = recalc(base_raw, "base")
base = read(bx)
wbk = openpyxl.load_workbook(bx, data_only=True)
fp = wbk["Sens"].cell(rm["SE_FP"], 4).value
cases = [(k, s, v) for k, s, v in rm["SENS_DEF"]]
res = {}
with cf.ThreadPoolExecutor(JOBS) as ex:
    for key, r in ex.map(run, cases):
        res[key] = r
        print(key, round(r["tp"], 2), flush=True)
# 反解：主導變數＝無 AI 基礎營業利益率（opm26＝opm32，全期持平），割線法求加權目標價＝現價
wbi = wbk["In"]
px = wbi.cell(rm["IN"]["px"], 5).value
x0 = wbi.cell(rm["IN"]["opm26"], 5).value


def tp_at(x, tag):
    raw = os.path.join(WD, f"rv_{tag}.xlsx")
    subprocess.run([sys.executable, os.path.join(HERE, "build_lycorp.py"), raw, "--set", f"opm26={x}", "--set", f"opm32={x}"], check=True, stdout=subprocess.DEVNULL)
    return read(recalc(raw, "rv" + tag))["tp"]


xa, fa = x0, base["tp"] - px
xb = x0 - 0.02
fb = tp_at(xb, "b") - px
xs, tps = xb, fb + px
for i in range(3):
    xn = xb - fb * (xb - xa) / (fb - fa)
    tn = tp_at(xn, f"n{i}")
    xa, fa, xb, fb = xb, fb, xn, tn - px
    xs, tps = xn, tn
    print("reverse", i, round(xn, 6), round(tn, 4), flush=True)
    if abs(tn - px) < 0.01:
        break
json.dump({"base_tp": base["tp"], "fp": fp, "cases": res, "reverse": {"x": xs, "tp": tps, "px": px, "var": "opm26,opm32"}}, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("base", base["tp"], "fp", fp)

#!/usr/bin/env python3
"""Naver 個股模型 v0.1 敏感度：In 頁每個有低／高區間的鍵（含聯動與組合案）一次改，整本以 LibreOffice 重算；並對敏感度最大的營運輸入做整本重算反解（G2 第 9 條）。
用法：python3 software_ai/naver/sens_naver.py <工作目錄> <輸出 sens.json> [JOBS]
"""
import json, os, subprocess, sys, concurrent.futures as cf
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WD, OUT = sys.argv[1], sys.argv[2]
JOBS = int(sys.argv[3]) if len(sys.argv) > 3 else 3
os.makedirs(WD, exist_ok=True)
rm_path = os.path.join(WD, "rowmap_base.json")
base_raw = os.path.join(WD, "base_raw.xlsx")
subprocess.run([sys.executable, os.path.join(HERE, "build_naver.py"), base_raw], check=True, env=dict(os.environ, ROWMAP=rm_path), stdout=subprocess.DEVNULL)
rm = json.load(open(rm_path, encoding="utf-8"))
VR = rm["VR"]
NONOP = {"beta", "erp", "kd_sp", "g_tv", "rel_adj", "w_dcf", "pe_inv", "dil", "fx", "ptok_idx", "tok_mult", "route", "nci_v"}


def recalc(raw, tag):
    od = os.path.join(WD, "o_" + tag)
    os.makedirs(od, exist_ok=True)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{WD}/lo_{tag}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", od, raw],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"))
    return os.path.join(od, os.path.basename(raw))


def read(x):
    V = openpyxl.load_workbook(x, data_only=True)["Val"]
    return {"tp": V.cell(VR["tp"], 2).value, "dcf": V.cell(VR["tpdcf"], 2).value, "gate": V.cell(VR["gate_B"], 2).value, "rec": V.cell(VR["rec"], 2).value}


def build_run(ov, tag):
    raw = os.path.join(WD, f"c_{tag}.xlsx")
    cmd = [sys.executable, os.path.join(HERE, "build_naver.py"), raw]
    for k, v in ov.items():
        cmd += ["--set", f"{k}={v}"]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL)
    return read(recalc(raw, tag))


def run(case):
    k, side, ov, lab = case
    tag = (k + "_" + side).replace(",", "_")
    return f"{k}|{side}", build_run(ov, tag)


bx = recalc(base_raw, "base")
base = read(bx)
fp = openpyxl.load_workbook(bx, data_only=True)["Sens"].cell(rm["SE_FP"], 4).value
res = {}
with cf.ThreadPoolExecutor(JOBS) as ex:
    for key, r in ex.map(run, [tuple(c) for c in rm["SENS_DEF"]]):
        res[key] = r
        print(key, round(r["tp"], 1), flush=True)
# 反解：營運輸入中相對基準絕對值最大者
best = None
inrows = {x[0]: x for x in rm["IN_ROWS"]}
for c in rm["SENS_DEF"]:
    k, side, ov, lab = c
    if "," in k or k in NONOP or k not in inrows: continue
    d = res[f"{k}|{side}"]["tp"] / base["tp"] - 1
    if best is None or abs(d) > abs(best[0]): best = (d, k, side, ov[k])
d, k, side, v1 = best
px = openpyxl.load_workbook(bx, data_only=True)["In"].cell(rm["IN"]["px"], 5).value
b0 = inrows[k][3] if isinstance(inrows[k][3], (int, float)) else None
if b0 is None:  # 公式基準：由基準活頁簿讀值
    b0 = openpyxl.load_workbook(bx, data_only=True)["In"].cell(rm["IN"][k], 5).value
x0, f0 = b0, base["tp"] - px
x1, f1 = v1, res[f"{k}|{side}"]["tp"] - px
sol = None
for it in range(7):
    if abs(f1 - f0) < 1e-9: break
    x2 = x1 - f1 * (x1 - x0) / (f1 - f0)
    r2 = build_run({k: x2}, f"solve{it}")
    x0, f0, x1, f1 = x1, f1, x2, r2["tp"] - px
    print("solve", k, x2, r2["tp"], flush=True)
    sol = {"key": k, "x": x2, "tp": r2["tp"]}
    if abs(f1) < 0.002 * px: break
fmt = "0.0%" if abs(b0) < 1 else "0.00"
lab = [x for x in rm["SENS_DEF"] if x[0] == k][0][3]
sol.update({"base": b0, "label": lab.split("（")[0], "fmt": fmt})
json.dump({"base_tp": base["tp"], "fp": fp, "cases": res, "solve": sol}, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("base", base["tp"], "solve", sol)

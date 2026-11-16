#!/usr/bin/env python3
"""敏感度：In 頁有區間（D 或 F 與基準 E 不同）的鍵，一次動一個、整本 LibreOffice 重算；另有聯動案例。
用法：python3 sens_nri.py <raw.xlsx（未含敏感度）> <rowmap.json> <out.json> [JOBS]"""
import json, os, shutil, subprocess, sys, tempfile
from concurrent.futures import ThreadPoolExecutor
import openpyxl
RAW, RM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
JOBS = int(sys.argv[4]) if len(sys.argv) > 4 else 4
rm = json.load(open(RM, encoding="utf-8")); IR, VR = rm["IR"], rm["VR"]
os.makedirs("/home/claude/scratch/nri", exist_ok=True)
WD = tempfile.mkdtemp(prefix="sens_", dir="/home/claude/scratch/nri")
def recalc(path, prof):
    out = os.path.join(os.path.dirname(path), "out"); os.makedirs(out, exist_ok=True)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, path],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"), check=True)
    return openpyxl.load_workbook(os.path.join(out, os.path.basename(path)), data_only=True)
def tp_of(wbv): V = wbv["Val"]; return V.cell(VR["tp"], 2).value, V.cell(VR["tpdcf"], 2).value, V.cell(VR["rec"], 2).value
b = os.path.join(WD, "base"); os.makedirs(b); shutil.copy(RAW, os.path.join(b, "case.xlsx"))
wbase = recalc(os.path.join(b, "case.xlsx"), os.path.join(WD, "lo0"))
base_tp, base_dcf, base_rec = tp_of(wbase)
Ib = wbase["In"]
cases = []
for k, r in IR.items():
    if k in ("px_date", "tp_date", "work_date", "rf_date", "data_date", "val_date"): continue
    e, d, f = (Ib.cell(r, c).value for c in (5, 4, 6))
    for side, v in (("低", d), ("高", f)):
        if isinstance(v, (int, float)) and not isinstance(v, bool) and isinstance(e, (int, float)) and abs(v - e) > 1e-12:
            lab = Ib.cell(r, 2).value
            if k in ("tk_tier", "inf_mult") and side == "高": lab += "（上界：只動成本一邊，價格保留率不同步）"
            if k == "tk_tier" and side == "低": lab += "（模型組合路由到便宜模型：只動成本一邊）"
            cases.append(([(k, v)], side, k, lab, f"{k}＝{v:g}（基準 {e:g}）"))
rets = {x: Ib.cell(IR["ret_" + x], 5).value for x in "ABC"}
cases.append(([("tk_tier", 3), ("inf_mult", 50)] + [("ret_" + x, min(1.0, rets[x] + 0.05)) for x in "ABC"], "高", "tk_tier,inf_mult,ret_A,ret_B,ret_C", "價格與成本聯動：推論成本取高端（最高層級、額外倍數 50）且成本轉嫁客戶（價格保留率各情境 +5 個百分點；同步變動）", "聯動高"))
cases.append(([("tk_tier", 1), ("inf_mult", 1)] + [("ret_" + x, rets[x] - 0.05) for x in "ABC"], "低", "tk_tier,inf_mult,ret_A,ret_B,ret_C", "模型組合路由到便宜模型：推論成本取低端（最低層級、額外倍數 1），客戶同步要求分享（價格保留率各情境 −5 個百分點）", "聯動低"))
cases.append(([("rm_adj", Ib.cell(IR["rm_adj"], 5).value)], "基準", "rm_adj", "相對倍數係數＝1（基準，原則第 47 條）", "rm_adj＝1"))
def run(i):
    edits, side, key, lab, setting = cases[i]
    d = os.path.join(WD, f"c{i}"); os.makedirs(d)
    wb = openpyxl.load_workbook(RAW)
    for k, v in edits: wb["In"].cell(IR[k], 5).value = v
    p = os.path.join(d, "case.xlsx"); wb.save(p)
    tp, dcf, rec = tp_of(recalc(p, os.path.join(WD, f"lo{i % JOBS}")))
    return {"key": key, "side": side, "label": lab, "setting": setting, "tp": tp, "dcf": dcf, "rec": rec}
from collections import defaultdict
groups = defaultdict(list); res = {}
for i in range(len(cases)): groups[i % JOBS].append(i)
def worker(g):
    for i in groups[g]: res[i] = run(i)
with ThreadPoolExecutor(JOBS) as ex: list(ex.map(worker, range(JOBS)))
json.dump({"fingerprint": rm["FP"], "build_id": rm["BUILD_ID"], "base_tp": base_tp, "base_dcf": base_dcf, "base_rec": base_rec, "cases": [res[i] for i in range(len(cases))]}, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
shutil.rmtree(WD, ignore_errors=True)
print("cases", len(cases), "base", base_tp)

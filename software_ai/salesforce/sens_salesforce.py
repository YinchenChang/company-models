#!/usr/bin/env python3
"""敏感度：In 頁有區間（D 或 F 與 E 不同）的鍵，一次動一個、整本 LibreOffice 重算。
用法：python3 sens_salesforce.py <raw.xlsx（base；未含敏感度）> <rowmap.json> <out.json> [JOBS]"""
import json, os, shutil, subprocess, sys, tempfile
from concurrent.futures import ThreadPoolExecutor
import openpyxl
RAW, RM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
JOBS = int(sys.argv[4]) if len(sys.argv) > 4 else 4
rm = json.load(open(RM, encoding="utf-8"))
IR, VR = rm["IR"], rm["VR"]
wbf = openpyxl.load_workbook(RAW)
ws = wbf["In"]
cases = []
for k, r in IR.items():
    if k in ("px_date",): continue
    e, d, f = ws.cell(r, 5).value, ws.cell(r, 4).value, ws.cell(r, 6).value
    for side, v in (("低", d), ("高", f)):
        if isinstance(v, (int, float)) and isinstance(e, (int, float)) and v != e:
            lab = ws.cell(r, 2).value
            if k == "ptok" and side == "高": lab += "（上界：只動成本一邊，售價不同步）"
            if k == "ptok" and side == "低": lab += "（模型組合路由到便宜模型）"
            if k == "cm_adj": lab += "（sens.%s）" % ("low" if side == "低" else "high")
            if k == "asr_sh": lab += "（ASR 最終交割列入敏感度）"
            cases.append((k, side, v, e, lab))
COMBO = [("ptok,p0", "高", {"ptok": "F", "p0": "F"}, "價格與成本聯動：token 價格與每任務單價同取高端（同步變動）"),
         ("ptok,p0", "低", {"ptok": "D", "p0": "D"}, "價格與成本聯動：token 價格與每任務單價同取低端（同步變動）"),
         ("cm_adj", "PS", {"cm_adj": ("in", "rel_ps")}, "相對倍數僅 PS 讀數（ps_only）"),
         ("cm_adj", "PSG", {"cm_adj": ("in", "rel_psg")}, "相對倍數僅 PSG 讀數（psg_only）")]
WD = tempfile.mkdtemp(prefix="sfsens_", dir="/home/claude/scratch/salesforce")


def recalc(path, prof):
    out = os.path.join(os.path.dirname(path), "out")
    os.makedirs(out, exist_ok=True)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, path],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"), check=True)
    wb = openpyxl.load_workbook(os.path.join(out, os.path.basename(path)), data_only=True)
    V = wb["Val"]
    return V.cell(VR["tp"], 2).value, V.cell(VR["dcf_tp"], 2).value, V.cell(VR["rec"], 2).value


def run(i):
    if i >= len(cases):
        k, side, cmap, lab = COMBO[i - len(cases)]
        d = os.path.join(WD, f"c{i}"); os.makedirs(d)
        wb = openpyxl.load_workbook(RAW)
        for kk, col in cmap.items():
            wb["In"].cell(IR[kk], 5).value = wb["In"].cell(IR[col[1]], 5).value if isinstance(col, tuple) else wb["In"].cell(IR[kk], 4 if col == "D" else 6).value
        p = os.path.join(d, "case.xlsx"); wb.save(p)
        tp, dcf, rec = recalc(p, os.path.join(WD, f"lo{i % JOBS}"))
        return {"key": k, "side": side, "label": lab, "setting": lab, "tp": tp, "dcf": dcf, "rec": rec}
    k, side, v, e, lab = cases[i]
    d = os.path.join(WD, f"c{i}"); os.makedirs(d)
    wb = openpyxl.load_workbook(RAW)
    wb["In"].cell(IR[k], 5).value = v
    p = os.path.join(d, "case.xlsx"); wb.save(p)
    tp, dcf, rec = recalc(p, os.path.join(WD, f"lo{i % JOBS}"))
    return {"key": k, "side": side, "label": lab, "setting": f"{k}＝{v:g}（基準 {e:g}）", "tp": tp, "dcf": dcf, "rec": rec}


b = os.path.join(WD, "base"); os.makedirs(b)
shutil.copy(RAW, os.path.join(b, "case.xlsx"))
base_tp, base_dcf, base_rec = recalc(os.path.join(b, "case.xlsx"), os.path.join(WD, "lo0"))
# 以 JOBS 個 profile 平行（同一 profile 不可同時跑）
from collections import defaultdict
groups = defaultdict(list)
for i in range(len(cases) + len(COMBO)): groups[i % JOBS].append(i)
res = {}
def worker(g):
    for i in groups[g]: res[i] = run(i)
with ThreadPoolExecutor(JOBS) as ex: list(ex.map(worker, range(JOBS)))
out = {"fingerprint": rm["FP"], "build_id": rm["BUILD_ID"], "base_tp": base_tp, "base_dcf": base_dcf, "base_rec": base_rec, "cases": [res[i] for i in range(len(cases) + len(COMBO))]}
json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
shutil.rmtree(WD, ignore_errors=True)
print("cases", len(cases), "base", base_tp)

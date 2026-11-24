#!/usr/bin/env python3
"""主導變數反解（原則第 4 條）：整本 LibreOffice 重算，割線法求使加權目標價＝現價所需的輸入值。
用法：python3 kreq_salesforce.py <raw.xlsx> <rowmap.json> <out.json>"""
import json, os, shutil, subprocess, sys, tempfile
import openpyxl
RAW, RM, OUT = sys.argv[1:4]
rm = json.load(open(RM, encoding="utf-8")); IR, VR = rm["IR"], rm["VR"]
WD = tempfile.mkdtemp(prefix="sfkreq_")  # 外部稽核優化 W3-9（2026-11-23）：寫死的容器路徑改為系統暫存資料夾（換機器也能跑）
base = openpyxl.load_workbook(RAW)["In"]
px = base.cell(IR["px"], 5).value
n = [0]
def tp_of(edits):
    n[0] += 1
    d = os.path.join(WD, f"c{n[0]}"); os.makedirs(d)
    wb = openpyxl.load_workbook(RAW)
    for k, v in edits.items(): wb["In"].cell(IR[k], 5).value = v
    p = os.path.join(d, "c.xlsx"); wb.save(p)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{WD}/lo", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", d + "/o", p],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"), check=True)
    return openpyxl.load_workbook(d + "/o/c.xlsx", data_only=True)["Val"].cell(VR["tp"], 2).value
def solve(mk, x0, x1, lo, hi):
    f = lambda x: tp_of(mk(x)) - px
    a, b = x0, x1; fa, fb = f(a), f(b)
    for _ in range(8):
        if fb == fa: break
        c = b - fb * (b - a) / (fb - fa)
        c = min(max(c, lo), hi)
        a, fa, b = b, fb, c; fb = f(b)
        if abs(fb) < 0.05: return b, "解"
    return (b, "解") if abs(fb) < 0.5 else (None, "區間內無解（%.2f 時目標價差 %.2f）" % (b, fb))
rows = []
g = lambda k: base.cell(IR[k], 5).value
b_rho = {s: g("rho_" + s) for s in "ABC"}
spec = [("tok_task", "每任務 token 數（token）", lambda x: {"tok_task": x}, g("tok_task"), "token/任務", g("tok_task") * 2, 1, 5e5),
        ("ptok", "每百萬 token 市場價（US$/M）", lambda x: {"ptok": x}, g("ptok"), "US$/M", g("ptok") * 2, 0.01, 200),
        ("rho", "席位年縮減率（三情境同時平移；基準為情境 B）", lambda x: {("rho_" + s): b_rho[s] + (x - b_rho["B"]) for s in "ABC"}, b_rho["B"], "比例/年", b_rho["B"] + 0.01, 0.0, 0.5),
        ("p0", "每任務實現單價 FY2027（US$）", lambda x: {"p0": x}, g("p0"), "US$/任務", g("p0") * 0.9, 0.001, 5)]
for k, lab, mk, b0, unit, x1, lo, hi in spec:
    v, st = solve(mk, b0, x1, lo, hi)
    rows.append({"label": lab, "base": b0, "solved": v if v is not None else st, "unit": unit, "note": "使加權目標價＝現價的值；" + st + "；其他輸入維持基準"})
json.dump({"fingerprint": rm["FP"], "px": px, "rows": rows}, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
shutil.rmtree(WD, ignore_errors=True)
print(json.dumps(rows, ensure_ascii=False, indent=1))

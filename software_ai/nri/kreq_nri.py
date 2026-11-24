#!/usr/bin/env python3
"""主導變數反解（原則第 4 條；G2-9）：整本 LibreOffice 重算，割線法求使加權目標價＝現價所需的輸入值。
用法：python3 kreq_nri.py <raw.xlsx> <rowmap.json> <out.json>"""
import json, os, shutil, subprocess, sys, tempfile
import openpyxl
RAW, RM, OUT = sys.argv[1:4]
rm = json.load(open(RM, encoding="utf-8")); IR, VR = rm["IR"], rm["VR"]
os.makedirs(os.path.join(tempfile.gettempdir(), "scratch", "nri"), exist_ok=True)
WD = tempfile.mkdtemp(prefix="kreq_", dir=os.path.join(tempfile.gettempdir(), "scratch", "nri"))
subprocess.run(["soffice", f"-env:UserInstallation=file://{WD}/lo", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", WD + "/b", RAW], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"), check=True)
base = openpyxl.load_workbook(WD + "/b/" + os.path.basename(RAW), data_only=True)["In"]
g = lambda k: base.cell(IR[k], 5).value
px = g("px"); n = [0]
def tp_of(edits):
    n[0] += 1; d = os.path.join(WD, f"c{n[0]}"); os.makedirs(d)
    wb = openpyxl.load_workbook(RAW)
    for k, v in edits.items(): wb["In"].cell(IR[k], 5).value = v
    p = os.path.join(d, "c.xlsx"); wb.save(p)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{WD}/lo", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", d + "/o", p], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"), check=True)
    return openpyxl.load_workbook(d + "/o/c.xlsx", data_only=True)["Val"].cell(VR["tp"], 2).value
def solve(mk, x0, x1, lo, hi):
    f = lambda x: tp_of(mk(x)) - px
    a, b = x0, x1; fa, fb = f(a), f(b)
    for _ in range(8):
        if fb == fa: break
        c = min(max(b - fb * (b - a) / (fb - fa), lo), hi)
        a, fa, b = b, fb, c; fb = f(b)
        if abs(fb) < 0.5: return b, "解"
    return (b, "解") if abs(fb) < 5 else (None, "區間內無解（%.4g 時目標價差 %.1f）" % (b, fb))
rets = {s: g("ret_" + s) for s in "ABC"}
spec = [("rm_adj", "相對倍數調整係數（Comparable 需重估倍數）", lambda x: {"rm_adj": x}, g("rm_adj"), "倍", g("rm_adj") * 1.2, 0.3, 3.0),
        ("beta_raw", "原始 β（影響 ke 與 WACC）", lambda x: {"beta_raw": x}, g("beta_raw"), "倍", g("beta_raw") + 0.3, -0.5, 3.0),
        ("ret", "價格保留率（三情境同時平移；基準為情境 B）", lambda x: {("ret_" + s): rets[s] + (x - rets["B"]) for s in "ABC"}, rets["B"], "比例", rets["B"] + 0.1, 0.0, 1.2),
        ("gv_28", "AI 前需求量年增 FY2028", lambda x: {"gv_28": x}, g("gv_28"), "比例/年", g("gv_28") + 0.02, -0.1, 0.3)]
rows = []
for k, lab, mk, b0, unit, x1, lo, hi in spec:
    v, st = solve(mk, b0, x1, lo, hi)
    rows.append({"label": lab, "base": b0, "solved": v if v is not None else st, "unit": unit, "note": "使加權目標價＝現價的值；" + st + "；其他輸入維持基準"})
json.dump({"fingerprint": rm["FP"], "px": px, "rows": rows}, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
shutil.rmtree(WD, ignore_errors=True)
print(json.dumps(rows, ensure_ascii=False, indent=1))

#!/usr/bin/env python3
"""敏感度與反解：整本 LibreOffice 重算，一次動一個（或一組聯動）輸入。
用法：python3 sens_snowflake.py <raw.xlsx（base；未含敏感度）> <rowmap.json> <out.json> [JOBS]
      python3 sens_snowflake.py --kreq <raw.xlsx> <rowmap.json> <out.json>   # 反解：支持現價所需的 opx_dr"""
import json, os, shutil, subprocess, sys, tempfile
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
import openpyxl
KREQ = sys.argv[1] == "--kreq"
a = sys.argv[2:] if KREQ else sys.argv[1:]
RAW, RM, OUT = a[0], a[1], a[2]
JOBS = int(a[3]) if len(a) > 3 else 4
rm = json.load(open(RM, encoding="utf-8"))
IR, VR = rm["IR"], rm["VR"]
wbf = openpyxl.load_workbook(RAW)
ws = wbf["In"]
os.makedirs("/home/claude/scratch/snowflake", exist_ok=True)
WD = tempfile.mkdtemp(prefix="snsens_", dir="/home/claude/scratch/snowflake")


def recalc(path, prof):
    out = os.path.join(os.path.dirname(path), "out")
    os.makedirs(out, exist_ok=True)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, path],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"), check=True)
    wb = openpyxl.load_workbook(os.path.join(out, os.path.basename(path)), data_only=True)
    V = wb["Val"]
    return V.cell(VR["tp"], 2).value, V.cell(VR["dcf_tp"], 2).value, V.cell(VR["rec"], 2).value


def run_case(i, sets, prof):
    d = os.path.join(WD, f"c{i}"); os.makedirs(d, exist_ok=True)
    wb = openpyxl.load_workbook(RAW)
    for k, v in sets.items(): wb["In"].cell(IR[k], 5).value = v
    p = os.path.join(d, "case.xlsx"); wb.save(p)
    return recalc(p, prof)


def val(k): return ws.cell(IR[k], 5).value


if KREQ:
    px = ws.cell(IR["px"], 5).value
    outs = []
    for key, label, hi0 in (("opx_dr", "非 GAAP 營業費用率逐年下降幅度（opx_dr，百分點/年；FY2028 起）", 0.2), ("cm_adj", "Comparable 倍數調整係數（cm_adj；對同業前瞻 P/S 中位數的倍數）", 5.0)):
        lo, hi = (0.0 if key == "opx_dr" else 0.5), hi0
        f = lambda x: run_case(0, {key: x}, os.path.join(WD, "lo0"))[0]
        base_tp = f(val(key)); top = f(hi)
        if top < px:
            req, note = "無解", f"上限 {hi:g} 仍不足以使加權目標價達現價（上限時 {top:.1f}）；其餘輸入取基準"
        else:
            for _ in range(14):
                mid = (lo + hi) / 2
                if f(mid) < px: lo = mid
                else: hi = mid
            req, note = round((lo + hi) / 2, 3), "二分法，14 次整本重算；其餘輸入取基準；反解值是「加權目標價＝現價」的點，不是評等門檻"
        outs.append({"key": key, "label": label, "base": val(key), "px": px, "base_tp": base_tp, "req": req, "note": note})
    json.dump(outs, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(outs); shutil.rmtree(WD, ignore_errors=True); sys.exit(0)

cases = []  # (key, side, sets, setting, label)
NOTE = {("g_px", "高"): "｜上界：只動成本端價格，售價不聯動", ("g_px", "低"): "｜只動成本端價格，售價不聯動", ("U_tier", "低"): "｜路由到較便宜模型",
        ("U_tier", "高"): "｜路由到較貴模型"}
for k, r in IR.items():
    if k in ("px_date",): continue
    e, d, f = ws.cell(r, 5).value, ws.cell(r, 4).value, ws.cell(r, 6).value
    for side, v in (("低", d), ("高", f)):
        if isinstance(v, (int, float)) and isinstance(e, (int, float)) and v != e:
            cases.append((k, side, {k: v}, f"{k}＝{v:g}（基準 {e:g}）", ws.cell(r, 2).value + NOTE.get((k, side), "")))
# 聯動與組合案例（原則第 44 條）
cases.append(("g_px", "高", {"g_px": 0.0, "g_px_rs": 0.0}, "g_px 與 g_px_rs 同步＝0（基準 −0.25）", "價格與成本聯動（成本端與轉售售價同步不降價）"))
cases.append(("g_px", "低", {"g_px": -0.5, "g_px_rs": -0.5}, "g_px 與 g_px_rs 同步＝−0.5（基準 −0.25）", "價格與成本聯動（成本端與轉售售價同步每年降一半）"))
cases.append(("sh_dil", "低", {"sh_dil": 0.0}, "sh_dil＝0（原口徑：潛在稀釋全不計；基準 18）", "原口徑（潛在稀釋全不計）"))
for pre, lab, dl in (("sbc32", "股票薪酬占營收 FY2032（三情境同步）", 0.04), ("U_w32", "代理工作負載乘數 FY2032（三情境同步）", 0.10), ("U_e32", "工作負載被替代 FY2032（三情境同步）", 0.04), ("U_mk", "AI 功能轉售加成（三情境同步）", 0.10)):
    for side, sg in (("低", -1), ("高", 1)):
        sets = {f"{pre}_{s}": ws.cell(IR[f"{pre}_{s}"], 5).value + sg * dl for s in "ABC"}
        cases.append((pre, side, sets, f"{pre} 三情境同步 {'+' if sg > 0 else '−'}{dl:g}", lab + ("（高＝數值加）" if sg > 0 else "（低＝數值減）")))
groups = defaultdict(list)
for i in range(len(cases)): groups[i % JOBS].append(i)
res = {}
b = os.path.join(WD, "base"); os.makedirs(b)
shutil.copy(RAW, os.path.join(b, "case.xlsx"))
base_tp, base_dcf, base_rec = recalc(os.path.join(b, "case.xlsx"), os.path.join(WD, "lo0"))


def worker(g):
    for i in groups[g]:
        k, side, sets, setting, lab = cases[i]
        tp, dcf, rec = run_case(i, sets, os.path.join(WD, f"lo{g}"))
        res[i] = {"key": k, "side": side, "label": lab, "setting": setting, "tp": tp, "dcf": dcf, "rec": rec}


with ThreadPoolExecutor(JOBS) as ex: list(ex.map(worker, range(JOBS)))
out = {"fingerprint": rm["FP"], "build_id": rm["BUILD_ID"], "base_tp": base_tp, "base_dcf": base_dcf, "base_rec": base_rec, "cases": [res[i] for i in range(len(cases))]}
json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
shutil.rmtree(WD, ignore_errors=True)
print("cases", len(cases), "base", base_tp)

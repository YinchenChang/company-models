#!/usr/bin/env python3
"""敏感度與反解：In 頁有區間（D 或 F 與 E 不同）的鍵，一次動一個、整本 LibreOffice 重算；另列聯動案與模型路由案；
再對主導變數做整本重算反解（支持現價所需值）。
用法：python3 sens_fujitsu.py <raw.xlsx（未含敏感度）> <rowmap.json> <sens.json> <rev.json> [JOBS]"""
import json, os, shutil, subprocess, sys, tempfile
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
import openpyxl

RAW, RM, OUT, REVOUT = sys.argv[1:5]
JOBS = int(sys.argv[5]) if len(sys.argv) > 5 else 4
SLUG = os.environ.get("SLUG", "fujitsu")
rm = json.load(open(RM, encoding="utf-8"))
IR, VR = rm["IR"], rm["VR"]
wbf = openpyxl.load_workbook(RAW)
ws = wbf["In"]
base = {k: ws.cell(r, 5).value for k, r in IR.items()}
cases = []
for k, r in IR.items():
    e, d, f = ws.cell(r, 5).value, ws.cell(r, 4).value, ws.cell(r, 6).value
    for side, v in (("低", d), ("高", f)):
        if isinstance(v, (int, float)) and isinstance(e, (int, float)) and v != e:
            lab = ws.cell(r, 2).value + ("（上界：只動成本一邊）" if k == "cps" and side == "高" else "")
            cases.append(dict(key=k, side=side, label=lab, setting=f"{k}＝{v:g}（基準 {e:g}）", sets=[(k, v)]))
# G5 第 1 條：歷史相對倍數 ps_only／psg_only（psg_only 缺：PSG 有效點不足；sens.low／high 已是 cm_adj 的低高端）
_rm = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "relmult_hist_fujitsu.json"), encoding="utf-8"))
for _k, _lab in (("ps_only", "只用 PS 讀數"), ("psg_only", "只用 PSG 讀數")):
    if _rm["sens"].get(_k) is not None:
        cases.append(dict(key=f"cm_adj:{_k}", side="相對倍數", label=f"歷史相對倍數：{_lab}（Comparable 倍數調整係數 cm_adj）", setting=f"cm_adj＝{_rm['sens'][_k]:g}（基準 {base['cm_adj']:g}）", sets=[("cm_adj", _rm["sens"][_k])]))
if "ps_use_db" in IR:
    cases.append(dict(key="ps_use_db:db_quarter", side="同業 P/S", label="同業 P/S＝資料庫最近一季中位（對照：v0.2／v0.2.1 的值）", setting="ps_use_db＝1（ps_med_db＝%g）" % base["ps_med_db"], sets=[("ps_use_db", 1)]))
# 聯動與模型路由（原則第 44 條）
a_keys = [k for k in IR if k.startswith("a32_")]
if "cps" in IR:
    cps_hi, cps_lo = ws.cell(IR["cps"], 6).value, ws.cell(IR["cps"], 4).value
    cases.append(dict(key="joint:cps_hi+a_lo", side="聯動", label="價格與用量聯動：token 價格取高端（Astra 層級）且自動化比例減半（貴則少用）", setting=f"cps＝{cps_hi:g}；a26、a32_A/B/C 各乘 0.5",
                      sets=[("cps", cps_hi), ("a26", base["a26"] * 0.5)] + [(k, base[k] * 0.5) for k in a_keys]))
    cases.append(dict(key="joint:route_terra", side="模型組合", label="模型路由：多數任務路由到便宜模型（Terra 層級成本）且自動化比例乘 1.25（便宜則多用）", setting=f"cps＝{cps_lo:g}；a26、a32_A/B/C 各乘 1.25",
                      sets=[("cps", cps_lo), ("a26", base["a26"] * 1.25)] + [(k, base[k] * 1.25) for k in a_keys]))
os.makedirs(f"/home/claude/scratch/{SLUG}", exist_ok=True)
WD = tempfile.mkdtemp(prefix=f"{SLUG}sens_", dir=f"/home/claude/scratch/{SLUG}")


def recalc(path, prof):
    out = os.path.join(os.path.dirname(path), "out"); os.makedirs(out, exist_ok=True)
    for attempt in range(4):
        rc = subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, path],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8")).returncode
        if rc == 0 and os.path.exists(os.path.join(out, os.path.basename(path))): break
        import time; time.sleep(3)
    else:
        raise RuntimeError("soffice 重算失敗：" + path)
    wb = openpyxl.load_workbook(os.path.join(out, os.path.basename(path)), data_only=True)
    V = wb["Val"]
    return V.cell(VR["tp"], 2).value, V.cell(VR["dcf_tp"], 2).value, V.cell(VR["rec"], 2).value


def run(tag, sets, prof):
    d = os.path.join(WD, tag); os.makedirs(d, exist_ok=True)
    wb = openpyxl.load_workbook(RAW)
    for k, v in sets: wb["In"].cell(IR[k], 5).value = v
    p = os.path.join(d, "case.xlsx"); wb.save(p)
    return recalc(p, prof)


b = os.path.join(WD, "base"); os.makedirs(b); shutil.copy(RAW, os.path.join(b, "case.xlsx"))
base_tp, base_dcf, base_rec = recalc(os.path.join(b, "case.xlsx"), os.path.join(WD, "lo0"))
groups = defaultdict(list)
for i in range(len(cases)): groups[i % JOBS].append(i)
res = {}


def worker(g):
    for i in groups[g]:
        tp, dcf, rec = run(f"c{i}", cases[i]["sets"], os.path.join(WD, f"lo{g}"))
        res[i] = (tp, dcf, rec)


with ThreadPoolExecutor(JOBS) as ex: list(ex.map(worker, range(JOBS)))
out_cases = []
for i, c in enumerate(cases):
    tp, dcf, rec = res[i]
    out_cases.append({"key": c["key"], "side": c["side"], "label": c["label"], "setting": c["setting"], "tp": tp, "dcf": dcf, "rec": rec})
json.dump({"fingerprint": rm["FP"], "build_id": rm["BUILD_ID"], "base_tp": base_tp, "base_dcf": base_dcf, "base_rec": base_rec, "cases": out_cases}, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
# ---- 反解：主導變數（AI 命題與營運參數中，低高案加權目標價差最大者）
px = openpyxl.load_workbook(RAW, data_only=False)["In"].cell(IR["px"], 5).value
CAND = os.environ.get("REV_CAND", "hg,rg,cs,a26,n26,gm,corp_dx").split(",")
span = {}
for c, (tp, dcf, rec) in zip(cases, [res[i] for i in range(len(cases))]):
    if c["key"] in CAND: span.setdefault(c["key"], []).append(tp)
dom = max(span, key=lambda k: max(span[k]) - min(span[k]))
e = base[dom]
lo_tp = [x for c, x in zip(cases, [res[i][0] for i in range(len(cases))]) if c["key"] == dom]
vals = [(c["sets"][0][1], res[i][0]) for i, c in enumerate(cases) if c["key"] == dom]
vals.append((e, base_tp))
(x0, y0), (x1, y1) = sorted(vals, key=lambda t: t[0])[0], sorted(vals, key=lambda t: t[0])[-1]
slope = (y1 - y0) / (x1 - x0)
x = e + (px - base_tp) / slope
hist = []
for it in range(6):
    tp, _, _ = run(f"rev{it}", [(dom, x)], os.path.join(WD, "lo0"))
    hist.append((x, tp))
    if abs(tp - px) < 0.05: break
    s2 = (tp - base_tp) / (x - e) if x != e else slope
    x = e + (px - base_tp) / s2
label = openpyxl.load_workbook(RAW)["In"].cell(IR[dom], 2).value
ok = abs(hist[-1][1] - px) < 0.5
rng = [v for v in (openpyxl.load_workbook(RAW)["In"].cell(IR[dom], 4).value, openpyxl.load_workbook(RAW)["In"].cell(IR[dom], 6).value) if v is not None]
outside = "" if (not rng or min(rng) <= hist[-1][0] <= max(rng)) else f"；解落在敏感度區間 {min(rng):g}～{max(rng):g} 之外（區間內無解）"
json.dump({"key": dom, "label": f"{dom}：{label}", "base": e, "solved": hist[-1][0], "tp_at_solved": hist[-1][1], "px": px, "converged": ok,
           "note": f"整本重算割線法：將 {dom} 由基準 {e:g} 改為 {hist[-1][0]:.4g}，加權目標價 {hist[-1][1]:.2f} 對現價 {px:.2f}（{'收斂' if ok else '未收斂，值僅供參考'}）{outside}；其他輸入維持基準；{dom} 為 Sens 頁 AI 命題與營運參數中影響最大者。",
           "history": hist}, open(REVOUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
shutil.rmtree(WD, ignore_errors=True)
print("cases", len(cases), "base", base_tp, "dominant", dom, hist[-1])

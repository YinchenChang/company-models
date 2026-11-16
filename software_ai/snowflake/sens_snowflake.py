#!/usr/bin/env python3
"""敏感度與反解：整本 LibreOffice 重算，一次動一個（或一組聯動）輸入。
用法：python3 sens_snowflake.py <raw.xlsx（base；未含敏感度）> <rowmap.json> <out.json> [JOBS]
      python3 sens_snowflake.py --kreq <raw.xlsx> <rowmap.json> <out.json>   # 反解：支持現價所需的 opx_v 與 rel_mult"""
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


# In 頁 D／E／F 有公式（標準版 DCF 列、迴歸導出列）：以整本重算後的數值為準
b0 = os.path.join(WD, "base"); os.makedirs(b0, exist_ok=True)
shutil.copy(RAW, os.path.join(b0, "case.xlsx"))
subprocess.run(["soffice", f"-env:UserInstallation=file://{os.path.join(WD, 'lo0')}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", os.path.join(b0, "out"), os.path.join(b0, "case.xlsx")],
               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"), check=True)
wbv = openpyxl.load_workbook(os.path.join(b0, "out", "case.xlsx"), data_only=True)
ws = wbv["In"]; QV = wbv["Q"]; RS = rm["RS"]
qv = lambda k: QV.cell(RS[k], 2).value


def val(k): return ws.cell(IR[k], 5).value


if KREQ:
    px = ws.cell(IR["px"], 5).value
    outs = []
    for key, label, hi0 in (("opx_v", "非 GAAP 營業費用變動部分（opx_v；每增加 1 美元營收增加的營業費用；越低越有利）", 0.0), ("rel_mult", "Comparable 倍數 rel_mult（SNOW 相對同業前瞻 P/S 中位數的倍數）", 6.0)):
        lo, hi = (0.5, 6.0) if key == "rel_mult" else (0.0, 0.46)
        f = lambda x: run_case(0, {key: x}, os.path.join(WD, "lo0"))[0]
        base_tp = f(val(key))
        if key == "opx_v":   # 方向相反：越低越有利，從基準往 0 找 f(x)=px 的點
            top = f(0.0)
            if top < px:
                req, note = "無解", f"變動部分降到 0 仍不足以使加權目標價達現價（為 0 時 {top:.1f}）；其餘輸入取基準（opx_a 固定部分不變）"
            else:
                lo2, hi2 = 0.0, val(key)
                for _ in range(14):
                    mid = (lo2 + hi2) / 2
                    if f(mid) >= px: lo2 = mid
                    else: hi2 = mid
                req, note = round((lo2 + hi2) / 2, 3), "二分法，14 次整本重算；其餘輸入取基準；反解值是「加權目標價＝現價」的點，不是評等門檻"
            outs.append({"key": key, "label": label, "base": val(key), "px": px, "base_tp": base_tp, "req": req, "note": note}); continue
        top = f(hi)
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
    if k in ("U_gm0",) and False: continue
    for side, v in (("低", d), ("高", f)):
        if isinstance(v, (int, float)) and isinstance(e, (int, float)) and v != e:
            cases.append((k, side, {k: v}, f"{k}＝{v:g}（基準 {e:g}）", ws.cell(r, 2).value + NOTE.get((k, side), "")))
# 聯動與組合案例（原則第 44 條）
cases.append(("g_px", "高", {"g_px": 0.0, "g_px_rs": 0.0}, "g_px 與 g_px_rs 同步＝0（基準 −0.25）", "價格與成本聯動（成本端與轉售售價同步不降價）"))
cases.append(("g_px", "低", {"g_px": -0.5, "g_px_rs": -0.5}, "g_px 與 g_px_rs 同步＝−0.5（基準 −0.25）", "價格與成本聯動（成本端與轉售售價同步每年降一半）"))
cases.append(("rsu", "低", {"rsu": 0.0, "opt_n": 0.0}, "rsu＝0、opt_n＝0（只計流通股與可轉債）", "不計 RSU 與選擇權稀釋"))
for pre, lab, dl in (("sbc_v", "股票薪酬變動部分（三情境同步）", 0.03), ("U_w32", "代理工作負載乘數 FY2032（三情境同步）", 0.10), ("U_e32", "工作負載被替代 FY2032（三情境同步）", 0.04), ("U_mk", "轉售加成 FY2032（三情境同步）", 0.10)):
    for side, sg in (("低", -1), ("高", 1)):
        sets = {f"{pre}_{s}": ws.cell(IR[f"{pre}_{s}"], 5).value + sg * dl for s in "ABC"}
        cases.append((pre, side, sets, f"{pre} 三情境同步 {'+' if sg > 0 else '−'}{dl:g}", lab + ("（高＝數值加）" if sg > 0 else "（低＝數值減）")))
# 「全期持平」案例（原則 G3 第 1 條：持平只放 Sens 下緣）：占營收比固定為最近四季（LTM）實績，不計規模經濟
ol, sl = qv("opxltm"), qv("sbcltm")
cases.append(("opx_v", "低", {"opx_a": 0.0, "opx_v": ol}, f"opx_a＝0、opx_v＝LTM 營業費用率 {ol:.3f}（全期持平）", "營業費用占營收全期持平於 LTM（Sens 下緣；無規模經濟）"))
cases.append(("sbc_v", "低", {"sbc_a": 0.0, "sbc_v_A": sl, "sbc_v_B": sl, "sbc_v_C": sl}, f"sbc_a＝0、sbc_v 三情境＝LTM 股票薪酬率 {sl:.3f}（全期持平）", "股票薪酬占營收全期持平於 LTM（Sens 下緣；無規模經濟）"))
cases.append(("opx_a", "低", {"opx_a": 0.0, "opx_v": ol, "sbc_a": 0.0, "sbc_v_A": sl, "sbc_v_B": sl, "sbc_v_C": sl}, "營業費用率與股票薪酬率同時全期持平於 LTM", "兩者同時全期持平（工作單原 G2 第 3 條口徑；Sens 下緣）"))
cases.append(("opx_v", "高", {"opx_a": qv("opxa6"), "opx_v": qv("opxv6"), "sbc_a": qv("sbca6"), "sbc_v_B": qv("sbcv6")}, "六點全樣本斜率（FY2022–LTM）取代最近三點", "全樣本規模經濟（營業費用與股票薪酬皆用六點迴歸；情境 A、C 變動部分不變）"))
groups = defaultdict(list)
for i in range(len(cases)): groups[i % JOBS].append(i)
res = {}
Vb = wbv["Val"]
base_tp, base_dcf, base_rec = Vb.cell(VR["tp"], 2).value, Vb.cell(VR["dcf_tp"], 2).value, Vb.cell(VR["rec"], 2).value


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

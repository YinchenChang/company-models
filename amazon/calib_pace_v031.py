# C10 速度重解（v0.3.1）：解 In 頁 pace_B，使情境 B 的 Q4 26 公式資本支出（AI＝新增總 MW ×（IT＋自建機房）＋非 AI）等於指引隱含的季速。
# 預設視窗＝下半年：（2026 指引 − 上半年實績）÷ 2（Thesis 頁 c10_gap＝0）。判斷替代（Sens）：只校準 Q4（指引 − 9M 實績）、全年平均（指引 ÷ 4）、指引 210／230。
# 另解舊 Trainium 口徑（cf_chip＝cf_cx＝1.0）的速度 pace_trn1（變動拆解 (c)／(d) 用）。這是把支出換算成容量（實體換算），不是由營收倒推價格（工作單第 2 節）。
# 用法：python3 amazon/calib_pace_v031.py [工作資料夾]  → 寫 amazon/data/pace_v031.json
import json, os, subprocess, sys, tempfile, shutil
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WD = sys.argv[1] if len(sys.argv) > 1 else tempfile.mkdtemp(prefix="calib_amzn_")
os.makedirs(WD, exist_ok=True)
BUILD = os.path.join(HERE, "build_amazon.py")
N = [0]


def resid(pace, extra, target):
    N[0] += 1
    name = f"p{N[0]:03d}"
    raw = os.path.join(WD, name + ".xlsx"); rmf = os.path.join(WD, name + ".json")
    cmd = [sys.executable, BUILD, raw, "--set", f"pace_B={pace}"] + sum((["--set", e] for e in extra), [])
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, env=dict(os.environ, AMZN_ROWMAP=rmf))
    out = os.path.join(WD, "out"); prof = tempfile.mkdtemp(prefix="lo_cal_")
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, raw],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"), timeout=600)
    shutil.rmtree(prof, ignore_errors=True)
    rm = json.load(open(rmf, encoding="utf-8"))
    REG = {tuple(k.split("|")): v for k, v in rm["REG"].items()}
    wb = openpyxl.load_workbook(os.path.join(out, name + ".xlsx"), data_only=True)
    g = lambda key: wb[REG[("BT", key)][0]][f"F{REG[('BT', key)][1]}"].value
    inn, q = wb["In"], wb["Q"]
    capg = inn.cell(rm["INROW"]["capex_g26"], 5).value
    c9 = q[f"O{rm['QR']['cxn']}"].value; c1h = q[f"P{rm['QR']['cxn']}"].value
    q4 = g("cx_fml") + g("nonai_f")
    if target == "h2": return (capg - c1h) / 2 - q4
    if target == "q4": return (capg - c9) - q4
    if target == "fy": return capg / 4 - q4
    raise ValueError(target)


def solve(extra, target="h2", x0=3634.3, x1=4200.0):
    f0 = resid(x0, extra, target); f1 = resid(x1, extra, target)
    for _ in range(8):
        if abs(f1) < 1e-3: break
        x2 = x1 - f1 * (x1 - x0) / (f1 - f0)
        x0, f0 = x1, f1
        x1 = round(x2, 2); f1 = resid(x1, extra, target)
        print(f"  {target} {extra} pace {x1:.2f} → 落差 {f1:+.4f}", flush=True)
    x = round(x1, 1)
    return x, round(resid(x, extra, target), 4)


res = {}
for k, extra, tg in (("pace", [], "h2"), ("pace_trn1", ["cf_chip=1", "cf_cx=1"], "h2"), ("pace_q4", [], "q4"), ("pace_fy", [], "fy"),
                     ("pace_g210", ["capex_g26=210"], "h2"), ("pace_g230", ["capex_g26=230"], "h2")):
    x, gap = solve(extra, tg)
    res[k] = x; res[k + "_gap"] = gap
    print(k, x, gap, flush=True)
res["note"] = ("C10 v0.3.1：情境 B 的 Q4 26 公式資本支出（AI＋非 AI）＝（2026 指引 − 上半年實績）÷ 2（下半年視窗）；pace_q4＝只校準 Q4（指引 − 9M 實績）；pace_fy＝指引 ÷ 4；"
               "pace_g210／g230＝指引 210／230；pace_trn1＝舊 Trainium 口徑 1.0。*_gap＝四捨五入到 0.1 MW／年後的落差（US$bn）")
json.dump(res, open(os.path.join(HERE, "data", "pace_v031.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(res)

# C10 速度重解（v0.3.1）：解 In 頁 pace，使 FY2026 資本支出公式值（9M 實績＋Q4 26 的 AI 與非 AI 公式）＝2026 指引中點 capex_g，
# 即 Thesis 頁「Q4 26 指引下限補足」的未截斷落差＝0。這是把支出換算成容量（實體換算），不是由營收倒推價格（工作單第 2 節）。
# 同時解 TPU 係數 1.0（舊 C2 口徑）時的速度 pace_tpu1（只供變動拆解與敏感度）。
# 用法：python3 alphabet/calib_pace_v031.py [工作資料夾]  → 寫 alphabet/data/pace_v031.json
import json, os, subprocess, sys, tempfile, shutil
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WD = sys.argv[1] if len(sys.argv) > 1 else tempfile.mkdtemp(prefix="calib_googl_")
os.makedirs(WD, exist_ok=True)
BUILD = os.path.join(HERE, "build_alphabet.py")


def resid(pace, extra, target=None):
    name = f"p{abs(hash((pace, tuple(extra)))) % 10**8}"
    raw = os.path.join(WD, name + ".xlsx"); rmf = os.path.join(WD, name + ".json")
    cmd = [sys.executable, BUILD, raw, "--set", f"pace={pace}"] + sum((["--set", e] for e in extra), [])
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, env=dict(os.environ, ROWMAP=rmf))
    out = os.path.join(WD, "out"); prof = tempfile.mkdtemp(prefix="lo_cal_")
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, raw],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"), timeout=600)
    shutil.rmtree(prof, ignore_errors=True)
    rm = json.load(open(rmf, encoding="utf-8"))
    wb = openpyxl.load_workbook(os.path.join(out, name + ".xlsx"), data_only=True)
    th, sg, inn = wb["Thesis"], wb["Seg"], wb["In"]
    g = inn.cell(rm["IN_ROW"]["capex_g"], 5).value; c9 = inn.cell(rm["IN_ROW"]["capex_1h"], 5).value
    q4 = th[f"F{rm['T']['cx_ai_f']}"].value + sg[f"F{rm['SR']['B']['cxn']}"].value
    if target == "h2":                     # 下半年視窗：Q4 公式 × 2＝指引 − 上半年實績（80.598），即以 2H 平均季速校準
        return (g - 80.598) / 2 - q4, q4
    return g - c9 - q4, q4


def solve(extra, x0=4159.67, x1=6500.0, target=None):
    f0, _ = resid(x0, extra, target); f1, _ = resid(x1, extra, target)
    for _ in range(8):
        if abs(f1) < 1e-4: break
        x2 = x1 - f1 * (x1 - x0) / (f1 - f0)
        x0, f0 = x1, f1
        x1 = round(x2, 2); f1, q4 = resid(x1, extra, target)
        print(f"  pace {x1:.2f} → 落差 {f1:+.4f}", flush=True)
    return round(x1, 1)


res = {}
res["pace"] = solve([])
res["pace_tpu1"] = solve(["tpu_f=1", "tpu_cf=1"])
for k, extra in (("pace", []), ("pace_tpu1", ["tpu_f=1", "tpu_cf=1"])):
    f, q4 = resid(res[k], extra)
    res[k + "_check"] = {"gap": round(f, 4), "q4_formula": round(q4, 3)}
# C10 的判斷替代（Sens 頁）：下半年視窗、指引區間兩端
res["pace_h2"] = solve([], target="h2")
res["pace_g195"] = solve(["capex_g=195"])
res["pace_g205"] = solve(["capex_g=205"])
res["note"] = "C10 v0.3.1：FY2026＝9M 實績＋Q4 公式＝指引中點（alphabet/calib_pace_v031.py；四捨五入到 0.1 MW／年後的落差見 *_check.gap）"
json.dump(res, open(os.path.join(HERE, "data", "pace_v031.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(res)

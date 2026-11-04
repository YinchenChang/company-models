# 台燿 Sens 產生器：一次把 In 頁一個鍵的基準值換成敏感度低／高值，整本重建、LibreOffice 重算，讀加權目標價（tp_final）與情境 A FY2028 EPS，
# 寫入 tuc/sens.json（建置時 build_tuc.py 讀入 Sens 頁）。In 頁或模型改動後重跑一次，再正式建置。
# 用法：python3 tuc/sens_tuc.py [暫存資料夾，預設 /tmp/work_tuc/sens]
import json, os, subprocess, sys, glob, shutil
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_tuc/sens"
shutil.rmtree(TMP, ignore_errors=True); os.makedirs(os.path.join(TMP, "raw")); os.makedirs(os.path.join(TMP, "rc"))
BUILD = os.path.join(HERE, "build_tuc.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("TUC_OVERRIDE", "TUC_ROWMAP")}
base_raw = os.path.join(TMP, "raw", "base.xlsx")
subprocess.run([sys.executable, BUILD, base_raw], check=True, env={**env0, "TUC_OVERRIDE": "_none=0"})
ws = openpyxl.load_workbook(base_raw)["In"]
jobs = {}
for r in range(4, ws.max_row + 1):
    k, lo, hi = ws.cell(r, 1).value, ws.cell(r, 4).value, ws.cell(r, 6).value
    if k and isinstance(lo, (int, float)) and isinstance(hi, (int, float)) and lo != hi:
        if "情境驅動對照" in str(ws.cell(r - 1, 1).value or ""): break
        jobs[k] = (lo, hi)
for k, (lo, hi) in jobs.items():
    for side, v in (("lo", lo), ("hi", hi)):
        subprocess.run([sys.executable, BUILD, os.path.join(TMP, "raw", f"{k}__{side}.xlsx")], check=True, env={**env0, "TUC_OVERRIDE": f"{k}={v}"})
subprocess.run([sys.executable, BUILD, os.path.join(TMP, "raw", "_wta.xlsx")], check=True, env={**env0, "TUC_WTA": "1", "TUC_ROWMAP": os.path.join(TMP, "wta_rowmap.json")})   # v0.13：Wattonomics 錨定
files = sorted(glob.glob(os.path.join(TMP, "raw", "*.xlsx")))
subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", os.path.join(TMP, "rc")] + files,
               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
def read(f):
    wb = openpyxl.load_workbook(f, data_only=True)
    dn = wb.defined_names["tp_final"].attr_text.split("!")[1].replace("$", "")
    tp = wb["Val"][dn].value
    eps = wb["Sens"]["G4"].value
    return tp, eps
out = {}
b = read(os.path.join(TMP, "rc", "base.xlsx"))
for k in jobs:
    out[k] = {"低": read(os.path.join(TMP, "rc", f"{k}__lo.xlsx")), "高": read(os.path.join(TMP, "rc", f"{k}__hi.xlsx"))}
_wb = openpyxl.load_workbook(os.path.join(TMP, "rc", "_wta.xlsx"), data_only=True); _sm = json.load(open(os.path.join(TMP, "wta_rowmap.json")))["SUMR"]
_tp, _ = read(os.path.join(TMP, "rc", "_wta.xlsx"))
out["_wta"] = {"tp": _tp, "epsA28": _wb["PL"][f"F{_sm['A|eps']}"].value, "epsB27": _wb["PL"][f"E{_sm['B|eps']}"].value, "epsB28": _wb["PL"][f"F{_sm['B|eps']}"].value}   # v0.13
json.dump(out, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
print("基準", b, "；案例", len(out))
print("Wattonomics 錨定", out["_wta"])
for k, v in sorted(((k, v) for k, v in out.items() if k != "_wta"), key=lambda kv: -max(abs(kv[1]["低"][0] / b[0] - 1), abs(kv[1]["高"][0] / b[0] - 1))):
    print(f"{k:10s} 低 {v['低'][0] / b[0] - 1:+.1%}  高 {v['高'][0] / b[0] - 1:+.1%}")

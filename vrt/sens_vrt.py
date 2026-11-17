# Vertiv 敏感度快照：每個案例以環境變數 VRT_OVERRIDE 覆寫一個輸入（低／高）後整本重建、LibreOffice 重算，
# 讀加權目標價與 DCF 目標價，寫入 vrt/sens.json；正式建置（build_vrt.py）讀入 Sens 頁「業務主因」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁模型狀態顯示 STALE）。
# 用法：LO_PROFILE=/tmp/lo_n5_vrt python3 vrt/sens_vrt.py [暫存資料夾，預設 /tmp/work_vrt/sens]（一次只開一個 LibreOffice）
import glob, json, os, shutil, subprocess, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_vrt/sens"
BUILD = os.path.join(HERE, "build_vrt.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("VRT_OVERRIDE", "VRT_ROWMAP")}
env0["VRT_SENS_RUN"] = "1"
src = open(os.path.join(HERE, "build_part2.py"), encoding="utf-8").read()
import re
KEYS = re.findall(r'\("([a-z0-9_]+)", "[^"]*", "[^"]*"\)', src.split("BIZ = [")[1].split("]\nBR0")[0])

def lab(ws, prefix):
    return next(r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix))

shutil.rmtree(TMP, ignore_errors=True); os.makedirs(os.path.join(TMP, "raw")); os.makedirs(os.path.join(TMP, "rc"))
jobs = {"base": ""}
for k in KEYS:
    jobs[f"{k}__lo"] = f"{k}=@lo"; jobs[f"{k}__hi"] = f"{k}=@hi"
CASES = {"share_back": "rel_27=1.20;rel_28=1.45;rel_29=1.45"}  # 與 build_part2.py 的 CASES 同步
for ck, cov in CASES.items():
    jobs[f"case_{ck}"] = cov
for nm, ov in jobs.items():
    subprocess.run([sys.executable, BUILD, os.path.join(TMP, "raw", f"{nm}.xlsx")], check=True, env={**env0, "VRT_OVERRIDE": ov}, stdout=subprocess.DEVNULL)
subprocess.run(["soffice"] + ([f"-env:UserInstallation=file://{os.environ['LO_PROFILE']}"] if os.environ.get("LO_PROFILE") else []) + ["--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", os.path.join(TMP, "rc")]
               + sorted(glob.glob(os.path.join(TMP, "raw", "*.xlsx"))), check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
def read(nm):
    wb = openpyxl.load_workbook(os.path.join(TMP, "rc", f"{nm}.xlsx"), data_only=True)
    V, S = wb["Val"], wb["Sens"]
    return dict(tp=V.cell(lab(V, "目標價（DCF 與 Comparable 加權）"), 2).value, dcf=V.cell(lab(V, "DCF 目標價"), 2).value, fp=S.cell(lab(S, "輸入指紋（本檔即時重算"), 4).value)
b = read("base")
out = {"_base": {"fp": b["fp"], "tp": b["tp"], "dcf": b["dcf"]}, "biz": {}}
print("基準", round(b["tp"], 4), "DCF", round(b["dcf"], 4))
for k in KEYS:
    d = {}
    for side, nm in (("低", f"{k}__lo"), ("高", f"{k}__hi")):
        r = read(nm); d[side] = {"tp": r["tp"], "dcf": r["dcf"]}
    out["biz"][k] = d
    print(f"{k:10s} 低 {d['低']['tp'] / b['tp'] - 1:+.2%} 高 {d['高']['tp'] / b['tp'] - 1:+.2%}")
out["cases"] = {}
for ck in CASES:
    r = read(f"case_{ck}"); out["cases"][ck] = {"tp": r["tp"], "dcf": r["dcf"]}
    print(f"案例 {ck}: 加權 {r['tp']:.2f}（{r['tp'] / b['tp'] - 1:+.2%}）DCF {r['dcf']:.2f}")
json.dump(out, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)

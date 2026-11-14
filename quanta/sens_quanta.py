# 廣達敏感度快照（v0.4.7 起；樣板 nvidia/sens_nvidia.py）：每個業務主因案例以環境變數 QUANTA_OVR 把 In 頁輸入改為敏感度低／高值後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 quanta/sens.json；正式建置（build_quanta.py → v047_post.py）讀入 Sens 頁「目標價敏感度：業務主因」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 quanta/sens_quanta.py [暫存資料夾，預設 /tmp/work_quanta/sens]；並行時設 LO_PROFILE＝file:///tmp/lo_quanta
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_quanta/sens"
BUILD = os.path.join(HERE, "build_quanta.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("QUANTA_OVR", "QUANTA_NOSENS", "SIL_COORD_ALT", "COORD_RUN")}
env0["QUANTA_NOSENS"] = "1"; env0["PYTHONDONTWRITEBYTECODE"] = "1"
sys.path.insert(0, HERE)
from v047_post import BIZ  # noqa: E402


def lab(ws, prefix):
    return next(r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix))


def run(jobs, sub):
    raw, rc = os.path.join(TMP, sub, "raw"), os.path.join(TMP, sub, "rc")
    shutil.rmtree(os.path.join(TMP, sub), ignore_errors=True); os.makedirs(raw); os.makedirs(rc)
    for nm, env in jobs.items():
        subprocess.run([sys.executable, BUILD, os.path.join(raw, f"{nm}.xlsx")], check=True, env={**env0, **env}, stdout=subprocess.DEVNULL)
    subprocess.run(["soffice"] + ([f"-env:UserInstallation={os.environ['LO_PROFILE']}"] if os.environ.get("LO_PROFILE") else [])
                   + ["--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", rc] + sorted(glob.glob(os.path.join(raw, "*.xlsx"))),
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return rc


ERR = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "Err:502")


def read(rc, nm):
    wb = openpyxl.load_workbook(os.path.join(rc, f"{nm}.xlsx"), data_only=True)
    V = wb["Val"]
    err = sum(1 for ws in wb for row in ws.iter_rows() for x in row if isinstance(x.value, str) and x.value in ERR)
    assert err == 0, (nm, err)
    return dict(tp=V.cell(lab(V, "目標價（DCF 與 Comparable 加權）"), 2).value, dcf=V.cell(lab(V, "DCF 目標價（情境加權）"), 2).value)


JOBS = {"base": {}}
for _, lo, hi, key in BIZ:
    JOBS[f"{key}__lo"] = {"QUANTA_OVR": lo}
    JOBS[f"{key}__hi"] = {"QUANTA_OVR": hi}
rc = run(JOBS, "cases")
R = {nm: read(rc, nm) for nm in JOBS}
b = R["base"]
out = {"_base": {"tp": b["tp"], "dcf": b["dcf"]}, "biz": {}}
print("基準", round(b["tp"], 4), "DCF", round(b["dcf"], 4))
for nm, r in R.items():
    if "__" in nm:
        k, side = nm.split("__")
        out["biz"].setdefault(k, {})["低" if side == "lo" else "高"] = {"tp": r["tp"], "dcf": r["dcf"]}
        print(f"{nm:20s} {r['tp']:9.3f} {r['tp'] / b['tp'] - 1:+.2%}  DCF {r['dcf'] - b['dcf']:+.2f}")
json.dump(out, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)

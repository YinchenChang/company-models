# AppLovin 敏感度（v0.1）：一次改一個（或聯動改數個）In 頁輸入，整本以 LibreOffice 重算，讀 Val 頁加權目標價、DCF 目標價、投資建議、CF 收支關卡 B、FY2027 營收 B 與輸入指紋，
# 寫成 JSON 供 build_applovin.py --sens 使用（鍵＝案例 id）。案例清單來自 build_applovin.py --cases（同一份定義，不手寫）。
# 用法：python3 software_ai/applovin/sens_v01.py <工作資料夾> <輸出 sens.json>（環境變數 JOBS＝平行數，預設 2）
import json, os, subprocess, sys, shutil, tempfile
from concurrent.futures import ThreadPoolExecutor
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WD, OUT = sys.argv[1], sys.argv[2]
os.makedirs(WD, exist_ok=True)
BUILD = os.path.join(HERE, "build_applovin.py")


def build(name, sets):
    raw = os.path.join(WD, f"{name}.xlsx")
    cmd = [sys.executable, BUILD, raw] + sum((["--set", f"{k}={v}"] for k, v in sets), [])
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, env={**os.environ, "ROWMAP": os.path.join(WD, f"{name}.rowmap.json")})
    return raw


def recalc(raw):
    out = os.path.join(WD, "out"); os.makedirs(out, exist_ok=True)
    prof = tempfile.mkdtemp(prefix="lo_app_")
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, raw],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env={**os.environ, "LC_ALL": "C.UTF-8"}, timeout=900)
    shutil.rmtree(prof, ignore_errors=True)
    return os.path.join(out, os.path.basename(raw))


def read(xlsx, rowmap):
    rm = json.load(open(rowmap, encoding="utf-8"))
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    g = lambda k: wb["Val"].cell(rm["VR"][k], 2).value
    return {"tp": g("tp"), "tpdcf": g("tpdcf"), "rec": g("rec"), "gateB": wb["CF"].cell(rm["GATE"], 3).value,
            "rev27B": wb["PL"][f"G{rm['PLR']['B']['rev']}"].value, "fp": wb["Sens"].cell(rm["SE_FP"], 4).value}


def run(name, sets):
    raw = build(name, sets)
    return read(recalc(raw), os.path.join(WD, f"{name}.rowmap.json"))


cases = json.loads(subprocess.check_output([sys.executable, BUILD, "--cases"]))
jobs = {"base|base": []}
for cs in cases: jobs[cs["id"]] = cs["sets"]
res = {}
with ThreadPoolExecutor(max_workers=int(os.environ.get("JOBS", "2"))) as ex:
    futs = {k: ex.submit(run, f"c{i:02d}", sets) for i, (k, sets) in enumerate(jobs.items())}
    for k, f in futs.items():
        o = f.result(); res[k] = {**o, "set": jobs[k]}
        print(k, jobs[k], round(o["tp"], 2), flush=True)
json.dump(res, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

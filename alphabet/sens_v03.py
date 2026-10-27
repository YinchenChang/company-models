# 敏感度與拆解：一次改一個 In 頁輸入（低／高值），整本以 LibreOffice 重算，讀 Val 頁加權目標價，寫成 JSON 供 build_alphabet.py --sens 使用。
# 用法：python3 alphabet/sens_v03.py <工作資料夾> <輸出 sens.json>（環境變數 JOBS＝平行數，預設 4）
import json, os, subprocess, sys, shutil, tempfile
from concurrent.futures import ThreadPoolExecutor
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WD, OUT = sys.argv[1], sys.argv[2]
os.makedirs(WD, exist_ok=True)
BUILD = os.path.join(HERE, "build_alphabet.py")


def build(name, sets):
    raw = os.path.join(WD, f"{name}.xlsx")
    cmd = [sys.executable, BUILD, raw] + sum((["--set", f"{k}={v}"] for k, v in sets), [])
    env = dict(os.environ); env["ROWMAP"] = os.path.join(WD, f"{name}.rowmap.json")
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, env=env)
    return raw


def recalc(raw):
    out = os.path.join(WD, "out"); os.makedirs(out, exist_ok=True)
    prof = tempfile.mkdtemp(prefix="lo_googl_")
    env = dict(os.environ); env["LC_ALL"] = "C.UTF-8"
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                    "--outdir", out, raw], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env, timeout=600)
    shutil.rmtree(prof, ignore_errors=True)
    return os.path.join(out, os.path.basename(raw))


def tp_of(xlsx, rowmap):
    rm = json.load(open(rowmap, encoding="utf-8"))
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    return wb["Val"].cell(rm["VR"]["tp"], 2).value


def run(name, sets):
    raw = build(name, sets)
    return tp_of(recalc(raw), os.path.join(WD, f"{name}.rowmap.json"))


base_raw = build("base", [])
rm = json.load(open(os.path.join(WD, "base.rowmap.json"), encoding="utf-8"))
IN = openpyxl.load_workbook(base_raw)["In"]
row = rm["IN_ROW"]


def lohi(key, which):
    return IN.cell(row[key], 4 if which == "lo" else 6).value


jobs = {"base|base": []}
for key, side, val, _r in rm["SE_ROWS"]:
    v = lohi(key, val) if val in ("lo", "hi") else val
    sets = [(key, v)]
    if key == "tpu_f":
        sets.append(("tpu_cf", lohi("tpu_cf", "lo")))
    jobs[f"{key}|{side}"] = sets
jobs["att_w|att"] = [("w_dcf", 0.45)]
jobs["att_px|att"] = [("px", 350.5)]
res = {}
with ThreadPoolExecutor(max_workers=int(os.environ.get("JOBS", "4"))) as ex:
    futs = {k: ex.submit(run, f"c{i:02d}", sets) for i, (k, sets) in enumerate(jobs.items())}
    for k, f in futs.items():
        res[k] = {"tp": f.result(), "set": jobs[k]}
        print(k, jobs[k], round(res[k]["tp"], 2), flush=True)
json.dump(res, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

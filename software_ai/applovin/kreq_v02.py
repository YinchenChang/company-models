# AppLovin 反解（G2-9）：整本 LibreOffice 重算，解出使加權目標價＝現價（req）與上檔空間＝+25%（flip；買進邊界）的主導變數值。
# 變數：beta（折現率）、u0_B（情境 B AI 提升）、e0_B（情境 B 新垂直增量）。割線法；只作展示、不進目標價。
# 用法：python3 software_ai/applovin/kreq_v02.py <工作資料夾> <輸出 json>
import json, os, subprocess, sys, shutil, tempfile
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
WD, OUT = sys.argv[1], sys.argv[2]
os.makedirs(WD, exist_ok=True)
N = [0]; CACHE = {}


def run(key, v):
    ck = (key, round(v, 7))
    if ck in CACHE: return CACHE[ck]
    N[0] += 1
    raw = os.path.join(WD, f"k{N[0]}.xlsx"); rmf = raw + ".rowmap.json"
    subprocess.run([sys.executable, os.path.join(HERE, "build_applovin.py"), raw, "--set", f"{key}={v}"], check=True, stdout=subprocess.DEVNULL, env={**os.environ, "ROWMAP": rmf})
    out = os.path.join(WD, "out"); prof = tempfile.mkdtemp(prefix="lo_kr_", dir=WD)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, raw],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env={**os.environ, "LC_ALL": "C.UTF-8"}, timeout=900)
    shutil.rmtree(prof, ignore_errors=True)
    rm = json.load(open(rmf, encoding="utf-8")); V = openpyxl.load_workbook(os.path.join(out, os.path.basename(raw)), data_only=True)["Val"]
    CACHE[ck] = (V.cell(rm["VR"]["tp"], 2).value, V.cell(rm["VR"]["px"], 2).value)
    return CACHE[ck]


def solve(key, target_up, a, b):
    f = lambda v: run(key, v)[0] / run(key, v)[1] - 1 - target_up
    fa, fb = f(a), f(b)
    for _ in range(25):
        if fb == fa: break
        c = b - fb * (b - a) / (fb - fa)
        a, fa, b, fb = b, fb, c, f(c)
        print(key, round(b, 5), round(fb, 6), flush=True)
        if abs(fb) < 1e-5: break
    return {"solved": abs(fb) < 1e-4, "value": b, "tp": run(key, b)[0]}


res = {"base_tp": run("beta", 2.0251)[0]}
for key, a, b in (("pe_rel", 1.0, 0.8), ("beta", 2.5, 3.0), ("u0_B", 0.08, 0.03), ("e0_B", 0.10, 0.05)):
    res[key] = {"req": solve(key, 0.0, a, b), "flip": solve(key, 0.25, a, b)}
json.dump(res, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps({k: v for k, v in res.items()}, ensure_ascii=False))

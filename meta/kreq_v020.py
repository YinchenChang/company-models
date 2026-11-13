# Meta v0.2.0：k 梯度與反解（原則第 10 節第 52、4 條；採用規劃書 A1）。整本 LibreOffice 重算，只作展示、不進目標價。
#   ladder：k_ai＝0、0.25、0.50、0.75、1.00、1.25 的加權目標價、DCF、Comparable（分部加總）目標價。
#   kreq：加權目標價＝現價的 k；kflip：上檔空間＝+25%（買進／中立邊界）的 k；kflip2：上檔空間＝+50%（買進／強力買進邊界）的 k。割線法。
# 做法同 spacex/kreq_v02.py。用法：python3 meta/kreq_v020.py <工作資料夾> <輸出 json>（make_v02.sh 以 KREQ=1 呼叫）
import json, os, subprocess, sys, shutil, tempfile
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
WD, OUT = sys.argv[1], sys.argv[2]
os.makedirs(WD, exist_ok=True)
N = [0]; CACHE = {}


def run(k):
    k = round(k, 6)
    if k in CACHE: return CACHE[k]
    N[0] += 1
    raw = os.path.join(WD, f"k{N[0]}.xlsx"); rmf = raw + ".rowmap.json"
    subprocess.run([sys.executable, os.path.join(HERE, "build_meta.py"), raw, "--set", f"k_ai={k}"], check=True, stdout=subprocess.DEVNULL, env={**os.environ, "ROWMAP": rmf})
    out = os.path.join(WD, "out"); prof = tempfile.mkdtemp(prefix="lo_kr_", dir=WD)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, raw],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env={**os.environ, "LC_ALL": "C.UTF-8"}, timeout=900)
    shutil.rmtree(prof, ignore_errors=True)
    rm = json.load(open(rmf, encoding="utf-8")); V = openpyxl.load_workbook(os.path.join(out, os.path.basename(raw)), data_only=True)["Val"]
    g = lambda key: V.cell(rm["VR"][key], 2).value
    CACHE[k] = {"tp": g("tp"), "tpdcf": g("tpdcf"), "tpcomp": g("tpcomp"), "px": g("px")}
    return CACHE[k]


def solve(target_up, a, b):
    f = lambda k: run(k)["tp"] / run(k)["px"] - 1 - target_up
    fa, fb = f(a), f(b)
    for _ in range(30):
        if fb == fa: break
        c = b - fb * (b - a) / (fb - fa)
        a, fa, b, fb = b, fb, c, f(c)
        print(round(b, 5), round(fb, 6), flush=True)
        if abs(fb) < 1e-6: break
    return {"solved": abs(fb) < 1e-4, "k": b, "tp": run(b)["tp"], "px": run(b)["px"], "target_up": target_up}


res = {"ladder": {f"{k:.2f}": run(k) for k in (0, 0.25, 0.5, 0.75, 1.0, 1.25)}}
res["kreq"] = solve(0.0, 0.25, 0.5)
res["kflip"] = solve(0.25, 0.75, 1.0)
res["kflip2"] = solve(0.5, 1.0, 1.25)
res["note"] = "k_ai 同乘 DCF 影子效益與分部加總的讀法 2 加回；整本 LibreOffice 重算反解；只作展示、不進基準（原則第 52 條）"
json.dump(res, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps({k: (v if k != "ladder" else {kk: round(vv["tp"], 2) for kk, vv in v.items()}) for k, v in res.items()}, ensure_ascii=False))

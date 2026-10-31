# 敏感度與變動拆解（v0.3.1）：一次改一個 In 頁輸入（Sens 頁案例，定義在 build_amazon.py 的 SCASES，經 rowmap 的 SE_JOBS 傳來），
# 整本以 LibreOffice 重算，讀 Val 頁加權目標價、投資建議、情境 B 目標價、FY2028 AWS 營業利益率、收支關卡，寫成 JSON 供 build_amazon.py --sens 使用。
# 另跑 v0.3 → v0.3.1 拆解鏈（Recon D 節）：S4＝本版但 (f) 其他改回 v0.3；S3＝S4 再改回舊 Trainium 口徑（速度依舊口徑重解）；
# S2＝S4＋舊口徑＋v0.3 的速度與 v0.3 隱含 Q4 營收（(c) 改回）；S1＝S2 再以 v0.3 對 2026Q3 的模型值取代實績（q3cf_v03.py；(b) 改回）。
# 用法：python3 amazon/sens_v031.py <工作資料夾> <輸出 sens.json>（環境變數 JOBS＝平行數，預設 2）
import json, os, subprocess, sys, shutil, tempfile
from concurrent.futures import ThreadPoolExecutor
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WD, OUT = sys.argv[1], sys.argv[2]
os.makedirs(WD, exist_ok=True)
BUILD = os.path.join(HERE, "build_amazon.py")
Q3 = os.path.join(HERE, "data", "q3cf_v03_2026q3.json")


def build(name, sets):
    raw = os.path.join(WD, f"{name}.xlsx")
    cmd = [sys.executable, BUILD, raw]
    for k, v in sets:
        cmd += ["--q3v03", Q3] if k == "__q3v03" else ["--set", f"{k}={v}"]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, env=dict(os.environ, AMZN_ROWMAP=os.path.join(WD, f"{name}.rowmap.json")))
    return raw


def recalc(raw):
    out = os.path.join(WD, "out"); os.makedirs(out, exist_ok=True)
    dst = os.path.join(out, os.path.basename(raw))
    for _ in range(3):                                   # LibreOffice 偶爾沒有輸出檔 → 重試
        prof = tempfile.mkdtemp(prefix="lo_amzn_")
        subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                        "--outdir", out, raw], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=dict(os.environ, LC_ALL="C.UTF-8"), timeout=600)
        shutil.rmtree(prof, ignore_errors=True)
        if os.path.exists(dst): break
    return dst


def read(xlsx, rmf):
    rm = json.load(open(rmf, encoding="utf-8"))
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    V, cf = wb["Val"], wb["CF"]
    sh, cell = rm["CX27B"].split("!")
    return {"tp": V["B19"].value, "rec": V["B24"].value, "tpB": V["C7"].value, "awsm28B": V["C10"].value, "gateB": cf["C6"].value, "gateW": cf["E6"].value,
            "tp1": V["B33"].value, "cx27B": wb[sh][cell].value}


def run(name, sets):
    raw = build(name, sets)
    return read(recalc(raw), os.path.join(WD, f"{name}.rowmap.json"))


base_raw = build("base", [])
rm = json.load(open(os.path.join(WD, "base.rowmap.json"), encoding="utf-8"))
jobs = {"base|base": []}
for k, sets in rm["SE_JOBS"].items():
    jobs[k] = [tuple(x) for x in sets]
PACE = json.load(open(os.path.join(HERE, "data", "pace_v031.json"), encoding="utf-8"))
Q3J = json.load(open(Q3, encoding="utf-8"))
S4 = [("rf", 0.05), ("px_val", 249.09), ("oth_q4", "=oth_r")]
S3 = S4 + [("cf_chip", 1), ("cf_cx", 1), ("pace_B", PACE["pace_trn1"])]
S2 = S4 + [("cf_chip", 1), ("cf_cx", 1), ("pace_B", Q3J["v03_2H26"]["pace_B"]), ("g26q4_lo", round(Q3J["q4rev"], 6)), ("g26q4_hi", round(Q3J["q4rev"], 6))]
S1 = S2 + [("__q3v03", 1)]
for k, v in (("S1", S1), ("S2", S2), ("S3", S3), ("S4", S4)):
    jobs[f"chain|{k}"] = v
if os.environ.get("ONLY"):
    keep = os.environ["ONLY"].split(",")
    jobs = {k: v for k, v in jobs.items() if k == "base|base" or any(k.startswith(x) for x in keep)}
res = {}
with ThreadPoolExecutor(max_workers=int(os.environ.get("JOBS", "2"))) as ex:
    futs = {k: ex.submit(run, f"c{i:03d}", sets) for i, (k, sets) in enumerate(jobs.items())}
    for k, f in futs.items():
        o = f.result()
        res[k] = {**o, "set": [[a, b] for a, b in jobs[k]]}
        print(k, round(o["tp"], 4), o["rec"], round(o["gateB"], 1), flush=True)
json.dump(res, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

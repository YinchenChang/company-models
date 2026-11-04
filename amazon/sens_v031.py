# 敏感度與變動拆解（v0.3.1）：一次改一個 In 頁輸入（Sens 頁案例，定義在 build_amazon.py 的 SCASES，經 rowmap 的 SE_JOBS 傳來），
# 整本以 LibreOffice 重算，讀 Val 頁加權目標價、投資建議、情境 B 目標價、FY2028 AWS 營業利益率、收支關卡，寫成 JSON 供 build_amazon.py --sens 使用。
# 另跑 v0.3.1 → v0.3.2 拆解鏈（Recon D 節；v0.3.2 起，10-Q 補入）：把 10-Q 值依序改回 v0.3.1 的推導或沿用值（--q3v03 事實覆寫）：
# S3＝本版但分部折舊 Q3 改回推導；S2＝S3＋融資義務改回 6/30（含利息）；S1＝S2＋融資租賃改回 6/30；S0＝S1＋債務面額改回推導（應＝v0.3.1）。
# （v0.3 → v0.3.1 的拆解鏈見 git 歷史與 data/bridge_v031.json。）
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
        if k.startswith("__fact:"):                      # 事實覆寫（拆解鏈：10-Q 值改回 v0.3.1）
            fj = os.path.join(WD, f"{name}.{k[7:]}.facts.json")
            json.dump({"facts": {k[7:]: v}}, open(fj, "w", encoding="utf-8"))
            cmd += ["--q3v03", fj]
        else:
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
# v0.3.1 的推導或沿用值（data/amazon_facts_2026q3.json 各 10-Q 事實的 note「原（v0.3.1）：」；分部折舊依 v0.3.1 公式重算）
_F = {f["id"]: f for f in json.load(open(os.path.join(HERE, "data", "amazon_facts_2026q3.json"), encoding="utf-8"))["items"]}
_O = {f["id"]: f for g in json.load(open(os.path.join(HERE, "data", "amazon_facts_20261008.json"), encoding="utf-8"))["groups"].values() for f in g}
_r = _F["cf.da.2026q3"]["value"] / _O["cf.da.2026q2"]["value"]
S3 = [(f"__fact:is.{c}.da.2026q3", _O[f"is.{c}.da.2026q2"]["value"] * _r) for c in ("NA", "INTL", "AWS")]
S2 = S3 + [("__fact:lease.finOblPrin.2026q3", _O["lease.finObl.2026q2"]["value"])]
S1 = S2 + [("__fact:lease.finPV.2026q3", _O["lease.finPV.2026q2"]["value"])]
S0 = S1 + [("__fact:debt.face.20260930", _O["debt.face.2026q2"]["value"] + _F["cf.ltdProceeds.2026q3"]["value"] - _F["cf.ltdRepay.2026q3"]["value"])]
for k, v in (("S0", S0), ("S1", S1), ("S2", S2), ("S3", S3)):
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
# 變動拆解鏈的中間值（v0.3.2 報告第 2 節；amazon/data/bridge_v032.json）
if all(f"chain|{k}" in res for k in ("S0", "S1", "S2", "S3")):
    v031 = 217.650781868569
    ch = [("v0.3.1", v031)] + [(k, res[f"chain|{k}"]["tp"]) for k in ("S0", "S1", "S2", "S3")] + [("v0.3.2", res["base|base"]["tp"])]
    labs = ["殘差：10-Q 值全部改回後 − v0.3.1（應為 0）", "(a) 債務面額 181.09 → 180.89", "(b) 融資租賃 13.45 → 14.50",
            "(c) 融資義務 11.07（含利息）→ 9.73（本金）", "(d) 2026Q3 分部折舊：推導 → 10-Q"]
    steps = [{"item": labs[i], "from": ch[i][0], "to": ch[i + 1][0], "tp_from": ch[i][1], "tp_to": ch[i + 1][1], "delta": ch[i + 1][1] - ch[i][1]} for i in range(5)]
    other = {}
    for kk in ("tp1", "tpB", "gateB", "gateW", "awsm28B", "cx27B"):
        other[kk] = {k: res[f"chain|{k}"][kk] for k in ("S0", "S1", "S2", "S3")}
        other[kk]["v0.3.2"] = res["base|base"][kk]
    bo = {"_readme": "Amazon v0.3.1 → v0.3.2 加權目標價變動拆解（10-Q 補入；逐步重建鏈，amazon/sens_v031.py chain|S0–S3 的整本重算值）。"
                     "S0＝10-Q 值全部改回 v0.3.1；S1＝(a) 後；S2＝(b) 後；S3＝(c) 後；(d)＝本版 − S3。",
          "chain": dict(ch), "chain_sets": {k: res[f"chain|{k}"]["set"] for k in ("S0", "S1", "S2", "S3")}, "steps": steps, "other_outputs": other,
          "total": ch[-1][1] - v031, "sum_check": sum(x["delta"] for x in steps) - (ch[-1][1] - v031)}
    json.dump(bo, open(os.path.join(HERE, "data", "bridge_v032.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

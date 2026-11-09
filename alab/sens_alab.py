# Astera Labs 敏感度快照（v0.3.6 起；驗收 A6、A9 與原則 1）：每個案例以環境變數覆寫輸入後整本重建、LibreOffice 重算，
# 讀加權目標價與 DCF 目標價，寫入 alab/sens.json；正式建置（build_alab.py → v036_post.py）讀入 Sens 頁「業務主因」表、Recon 反解與 README 目標價變動拆解。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁模型狀態顯示 STALE）。
# 用法：python3 alab/sens_alab.py [暫存資料夾，預設 /tmp/work_alab/sens]；ALAB_STAGE 同正式建置（預設 0.3.7）。
import json, os, subprocess, sys, glob, shutil
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_alab/sens"
STAGE = os.environ.get("ALAB_STAGE", "0.3.7")
shutil.rmtree(TMP, ignore_errors=True); os.makedirs(os.path.join(TMP, "raw")); os.makedirs(os.path.join(TMP, "rc"))
BUILD = os.path.join(HERE, "build_alab.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("ALAB_OVERRIDE", "ALAB_V37_PARTS", "ALAB_ROWMAP")}
env0["ALAB_STAGE"] = STAGE
JOBS = {"base": {}}
GK_LO, GK_HI = 0.173423 / 0.334406, 0.179056 / 0.261646   # FY2025、3Q26 單季（公司新聞稿；與 Hist 頁同值）
if STAGE >= "0.3.7":
    JOBS.update({
        "relx__lo": {"ALAB_OVERRIDE": "_relx=XLO"}, "relx__hi": {"ALAB_OVERRIDE": "_relx=XHI"},
        "relp__lo": {"ALAB_OVERRIDE": "_relp=PLO"}, "relp__hi": {"ALAB_OVERRIDE": "_relp=PHI"},
        "gaap_k__lo": {"ALAB_OVERRIDE": f"gaap_k={GK_LO}"}, "gaap_k__hi": {"ALAB_OVERRIDE": f"gaap_k={GK_HI}"},
        "cl__lo": {"ALAB_OVERRIDE": "_cl=1;cl_loss=0.15"}, "cl__hi": {"ALAB_OVERRIDE": "_cl=1;cl_loss=0.5"},
        "cl30__lo": {"ALAB_OVERRIDE": "_cl=1;cl_loss=0.3"},
        "shorig__lo": {"ALAB_OVERRIDE": "_shorig=1"},
    })
    XK = (1.25, 1.5, 2.0, 3.0)
    for k in XK: JOBS[f"xk_{k}"] = {"ALAB_OVERRIDE": f"_xk={k}"}
    STEPS = [("v0.3.6 口徑", ""), ("目標時點改現價日＋12 個月", "tt"), ("評價股數改流通股＋RSU＋選擇權", "tt,sh"),
             ("資產負債改 10-Q（租賃、現金）", "tt,sh,bs"), ("GAAP 轉換改最近四季（J2）", "tt,sh,bs,gaap"), ("情境 A／C 帶份額高／低路徑（J1）", "tt,sh,bs,gaap,j1")]
    for i, (_, parts) in enumerate(STEPS): JOBS[f"step{i}"] = {"ALAB_V37_PARTS": parts}
for nm, env in JOBS.items():
    subprocess.run([sys.executable, BUILD, os.path.join(TMP, "raw", f"{nm}.xlsx")], check=True, env={**env0, **env}, stdout=subprocess.DEVNULL)
files = sorted(glob.glob(os.path.join(TMP, "raw", "*.xlsx")))
subprocess.run(["soffice"] + ([f"-env:UserInstallation={os.environ['LO_PROFILE']}"] if os.environ.get("LO_PROFILE") else [])
               + ["--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", os.path.join(TMP, "rc")] + files,
               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def lab(ws, prefix):
    return next(r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.startswith(prefix))


def read(nm):
    wb = openpyxl.load_workbook(os.path.join(TMP, "rc", f"{nm}.xlsx"), data_only=True)
    V = wb["Val"]; SE = wb["Sens"]
    px = V.cell(lab(V, "現價（"), 2).value
    return dict(tp=V.cell(lab(V, "目標價（DCF 與 Comparable 加權）"), 2).value, dcf=V.cell(lab(V, "DCF 目標價（情境加權）"), 2).value,
                fp=SE.cell(lab(SE, "輸入指紋（本檔即時重算"), 4).value, px=px)


b = read("base")
out = {"_stage": STAGE, "_base": {"fp": b["fp"], "tp": b["tp"], "dcf": b["dcf"]}}
print("基準", round(b["tp"], 4), "DCF", round(b["dcf"], 4))
if STAGE >= "0.3.7":
    biz = {}
    for nm in JOBS:
        if "__" in nm:
            k, side = nm.split("__"); r = read(nm)
            biz.setdefault(k, {})["低" if side == "lo" else "高"] = {"tp": r["tp"], "dcf": r["dcf"]}
            print(f"{nm:12s} {r['tp'] / b['tp'] - 1:+.1%}  DCF {r['dcf'] - b['dcf']:+.2f}")
    out["biz"] = biz
    pts = [{"k": 1.0, "tp": b["tp"]}] + [{"k": k, "tp": read(f"xk_{k}")["tp"]} for k in XK]

    def solve(target):
        for p0, p1 in zip(pts, pts[1:]):
            if (p0["tp"] - target) * (p1["tp"] - target) <= 0 and p1["tp"] != p0["tp"]:
                return p0["k"] + (target - p0["tp"]) * (p1["k"] - p0["k"]) / (p1["tp"] - p0["tp"])
        return None
    out["xk"] = {"pts": pts, "half": solve(0.5 * b["px"]), "px": solve(b["px"])}
    print("xk", out["xk"])
    st = [(lab_, read(f"step{i}")["tp"]) for i, (lab_, _) in enumerate(STEPS)]
    parts = "；".join(f"{l} {t1 - t0:+.2f}" for (_, t0), (l, t1) in zip(st, st[1:]))
    out["attr"] = (f"v0.3.7 目標價變動拆解（逐項累加，alab/sens_alab.py 整本重算）：v0.3.6 加權目標價 {st[0][1]:.2f} 美元 → {parts} → v0.3.7 {st[-1][1]:.2f} 美元"
                   f"（合計 {st[-1][1] - st[0][1]:+.2f}，{st[-1][1] / st[0][1] - 1:+.1%}）。")
    print(out["attr"])
json.dump(out, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)

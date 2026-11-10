# AMD 敏感度快照（v0.4.5 起；驗收 A2、B3、B4、B5、J1 與原則 1、4、32、39）：每個案例以環境變數覆寫輸入後整本重建、LibreOffice 重算，
# 讀加權目標價與 DCF 目標價，寫入 amd/sens.json；正式建置（build_amd.py → v045_post.py）讀入 Sens 頁「業務主因」表、Recon 反解與 README 目標價變動拆解。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁模型狀態顯示 STALE）。
# 用法：python3 amd/sens_amd.py [暫存資料夾，預設 /tmp/work_amd/sens]；AMD_STAGE 同正式建置（預設 0.4.6）。
import json, os, subprocess, sys, glob, shutil
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_amd/sens"
STAGE = os.environ.get("AMD_STAGE", "0.4.6")
BUILD = os.path.join(HERE, "build_amd.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("AMD_OVERRIDE", "AMD_V46_PARTS", "AMD_ROWMAP")}
env0["AMD_STAGE"] = STAGE


def lab(ws, prefix):
    return next(r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.startswith(prefix))


def run(jobs, sub):
    raw, rc = os.path.join(TMP, sub, "raw"), os.path.join(TMP, sub, "rc")
    shutil.rmtree(os.path.join(TMP, sub), ignore_errors=True); os.makedirs(raw); os.makedirs(rc)
    for nm, env in jobs.items():
        subprocess.run([sys.executable, BUILD, os.path.join(raw, f"{nm}.xlsx")], check=True, env={**env0, **env}, stdout=subprocess.DEVNULL)
    subprocess.run(["soffice"] + ([f"-env:UserInstallation={os.environ['LO_PROFILE']}"] if os.environ.get("LO_PROFILE") else [])
                   + ["--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", rc] + sorted(glob.glob(os.path.join(raw, "*.xlsx"))),
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return rc


def read(rc, nm):
    wb = openpyxl.load_workbook(os.path.join(rc, f"{nm}.xlsx"), data_only=True)
    V = wb["Val"]; SE = wb["Sens"]
    try: ctx = wb.defined_names["cash_tx"].attr_text.split("!"); ctax = wb[ctx[0]][ctx[1].replace("$", "")].value
    except KeyError: ctax = None
    return dict(tp=V.cell(lab(V, "目標價（DCF 與 Comparable 加權）"), 2).value, dcf=V.cell(lab(V, "DCF 目標價（情境加權）"), 2).value,
                fp=SE.cell(lab(SE, "輸入指紋（本檔即時重算"), 4).value, px=V.cell(lab(V, "現價（"), 2).value, ctax=ctax)


rc0 = run({"base": {}}, "base")
b = read(rc0, "base")
out = {"_stage": STAGE, "_base": {"fp": b["fp"], "tp": b["tp"], "dcf": b["dcf"]}}
print("基準", round(b["tp"], 4), "DCF", round(b["dcf"], 4), "現金稅率", b["ctax"])
if STAGE >= "0.4.6":
    O = lambda s: {"AMD_OVERRIDE": s}
    JOBS = {"uf__lo": O("uf=0.8"), "uf__hi": O("uf=1.2"), "pg__hi": O("_pg=hi"), "abs__lo": O("_abs=lo"), "abs__hi": O("_abs=hi"),
            "gm_gpu__lo": O("gm_gpu=0.4"), "gm_gpu__hi": O("gm_gpu=0.55"), "cpu__lo": O("_cpu=lo"), "cpu__hi": O("_cpu=hi"),
            "gsh__lo": O("_gsh=0.4"), "gsh__hi": O("_gsh=0.6"), "oseg1__lo": O("_oseg=-1"), "oseg1__hi": O("_oseg=1"),
            "oseg05__lo": O("_oseg=-0.5"), "oseg05__hi": O("_oseg=0.5"), "wf__lo": O("wf=0"), "wf__hi": O("wf=0.5"),
            "rmhi__hi": O("_rmhi=3"), "tx__lo": O(f"tx_f={b['ctax']}"),
            "sbc_rev__lo": O(f"sbc_rev={1.513 / 35.166}"), "sbc_rev__hi": O(f"sbc_rev={(1.407 / 25.785 + 1.638 / 34.639 + 1.513 / 35.166) / 3}"),
            "gw_30__lo": O("gw_30=6"), "gw_30__hi": O("gw_30=10")}
    PK = (0.25, 0.5, 0.8, 1.25, 1.5, 2.0, 3.0, 4.0)
    for k in PK: JOBS[f"pk_{k}"] = O(f"_pgk={k}")
    STEPS = [("v0.4.5 口徑", ""), ("目標時點改現價日＋12 個月", "tt"), ("評價股數改流通股＋庫藏股法", "tt,sh"), ("股票薪酬比率改最近四季", "tt,sh,sbc"),
             ("認股權證閘門式歸屬（J1a）", "tt,sh,sbc,w"), ("情境 A／C 帶 AMD 顆數與 CPU 路徑（J2b）", "tt,sh,sbc,w,j2")]
    for i, (_, parts) in enumerate(STEPS): JOBS[f"step{i}"] = {"AMD_V46_PARTS": parts}
    rc = run(JOBS, "cases")
    biz = {}
    for nm in JOBS:
        if "__" in nm:
            k, side = nm.split("__"); r = read(rc, nm)
            biz.setdefault(k, {})["低" if side == "lo" else "高"] = {"tp": r["tp"], "dcf": r["dcf"]}
            print(f"{nm:12s} {r['tp'] / b['tp'] - 1:+.1%}  DCF {r['dcf'] - b['dcf']:+.2f}")
    out["biz"] = biz
    pts = sorted([{"k": 1.0, "tp": b["tp"]}] + [{"k": k, "tp": read(rc, f"pk_{k}")["tp"]} for k in PK], key=lambda p: p["k"])

    def solve(target):
        for p0, p1 in zip(pts, pts[1:]):
            if (p0["tp"] - target) * (p1["tp"] - target) <= 0 and p1["tp"] != p0["tp"]:
                return p0["k"] + (target - p0["tp"]) * (p1["k"] - p0["k"]) / (p1["tp"] - p0["tp"])
        return None
    u = b["tp"] / b["px"] - 1
    e_up = next((e for e in (-0.5, -0.25, 0.25, 0.5) if e > u), None); e_dn = next((e for e in (0.5, 0.25, -0.25, -0.5) if e < u), None)
    out["pk"] = {"pts": pts, "up": solve(b["px"] * (1 + e_up)) if e_up is not None else None,
                 "dn": solve(b["px"] * (1 + e_dn)) if e_dn is not None else None, "px": solve(b["px"])}
    print("pk", out["pk"])
    st = [(lab_, read(rc, f"step{i}")["tp"]) for i, (lab_, _) in enumerate(STEPS)]
    parts = "；".join(f"{l} {t1 - t0:+.2f}" for (_, t0), (l, t1) in zip(st, st[1:]))
    out["attr"] = (f"v0.4.6 目標價變動拆解（逐項累加，amd/sens_amd.py 整本重算）：v0.4.5 加權目標價 {st[0][1]:.2f} 美元 → {parts} → v0.4.6 {st[-1][1]:.2f} 美元"
                   f"（合計 {st[-1][1] - st[0][1]:+.2f}，{st[-1][1] / st[0][1] - 1:+.1%}）。")
    print(out["attr"])
json.dump(out, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)

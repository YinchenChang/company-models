# NVIDIA 敏感度快照（v0.3.5 起；驗收原則 1、22、32、39、48、50）：每個案例以環境變數覆寫輸入後整本重建、LibreOffice 重算，
# 讀加權目標價、DCF 目標價與 FY2028 毛利率（情境 B），寫入 nvidia/sens.json；正式建置（build_nvidia.py → v035_post.py）讀入
# Sens 頁「業務主因」表、Recon G 節（毛利率差距分解、abs_mem 反解）與 README 目標價變動拆解。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁模型狀態顯示 STALE）。
# 用法：python3 nvidia/sens_nvidia.py [暫存資料夾，預設 /tmp/work_nvidia/sens]；NV_STAGE 同正式建置（預設 0.3.6）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_nvidia/sens"
STAGE = os.environ.get("NV_STAGE", "0.3.6")
BUILD = os.path.join(HERE, "build_nvidia.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("NV_OVERRIDE", "NV_V36_PARTS", "NV_ROWMAP")}
env0["NV_STAGE"] = STAGE
sys.path.insert(0, HERE)
from v035_post import BIZ, BIZ36  # noqa: E402


def lab(ws, prefix, start=1):
    return next(r for r in range(start, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix))


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
    V = wb["Val"]; SE = wb["Sens"]; PL = wb["PL"]; IN = wb["In"]
    rb = lab(PL, "情境 B"); rg = lab(PL, "毛利率", rb)
    g28 = next(IN.cell(r, 5).value for r in range(1, IN.max_row + 1) if IN.cell(r, 1).value == "g28_gm")
    return dict(tp=V.cell(lab(V, "目標價（DCF 與 Comparable 加權）"), 2).value, dcf=V.cell(lab(V, "DCF 目標價（情境加權）"), 2).value,
                fp=SE.cell(lab(SE, "輸入指紋（本檔即時重算"), 4).value, px=V.cell(lab(V, "現價（"), 2).value,
                gm28=PL.cell(rg, 5).value, gm27=PL.cell(rg, 4).value, g28=g28)


rc0 = run({"base": {}}, "base")
b = read(rc0, "base")
out = {"_stage": STAGE, "_base": {"fp": b["fp"], "tp": b["tp"], "dcf": b["dcf"]}}
print("基準", round(b["tp"], 4), "DCF", round(b["dcf"], 4), "FY2028 毛利率 B", b["gm28"])
O = lambda s: {"NV_OVERRIDE": s}
SPEC = {"uk": ("_P:uf=uk_lo", "_P:uf=uk_hi"), "cn_amt": ("_P:cn=cn_amt_lo", "_P:cn=cn_amt_hi"), "opx_adj": ("_P:opxa=opx_adj_lo", "_P:opxa=opx_adj_hi"),
        "krdecay": (None, "_krdecay=1"), "sh_v": ("sh_v==h_shout", None)}
JOBS = {}
for _, key, _ in BIZ + (BIZ36 if STAGE >= "0.3.6" else []):
    lo, hi = SPEC.get(key, (f"{key}=@lo", f"{key}=@hi"))
    if lo: JOBS[f"{key}__lo"] = O(lo)
    if hi: JOBS[f"{key}__hi"] = O(hi)
if STAGE >= "0.3.6":
    JOBS["gm_abs0"] = O("abs_mem=0"); JOBS["gm_abs0_h"] = O("abs_mem=0;_hbmflat=1"); JOBS["gm_abs0_hc"] = O("abs_mem=0;_hbmflat=1;_cdxflat=1")
    STEPS = [("v0.3.5 口徑", ""), ("評價股數改流通股＋RSU（原則 11）", "sh"), ("目標時點改現價日＋12 個月（原則 26）", "sh,tt"),
             ("折舊與資本支出比率改最近四季（原則 35）", "sh,tt,ratio"), ("FY↔CY 月數加權、k_rec 與 c_nh 重算（原則 49）", "sh,tt,ratio,mw")]
    for i, (_, parts) in enumerate(STEPS): JOBS[f"step{i}"] = {"NV_V36_PARTS": parts}
rc = run(JOBS, "cases")
R = {nm: read(rc, nm) for nm in JOBS}
biz = {}
for nm, r in R.items():
    if "__" in nm:
        k, side = nm.split("__")
        biz.setdefault(k, {})["低" if side == "lo" else "高"] = {"tp": r["tp"], "dcf": r["dcf"]}
        print(f"{nm:16s} {r['tp'] / b['tp'] - 1:+.2%}  DCF {r['dcf'] - b['dcf']:+.2f}")
out["biz"] = biz
if STAGE >= "0.3.6":
    g0, ga, gh, gc = b["gm28"], R["gm_abs0"]["gm28"], R["gm_abs0_h"]["gm28"], R["gm_abs0_hc"]["gm28"]
    out["gm"] = {"total": g0 - b["g28"], "abs": g0 - ga, "hbm": ga - gh, "lp": gh - gc, "mix": gc - b["gm27"], "anchor": b["gm27"] - b["g28"]}
    print("毛利率分解", {k: round(v * 100, 3) for k, v in out["gm"].items()})
    pts = sorted([(0.0, R["abs_mem__lo"]["gm28"]), (0.5, g0), (1.0, R["abs_mem__hi"]["gm28"])])
    tgt = 0.72
    (x0, y0), (x1, y1) = (pts[0], pts[1]) if tgt >= pts[1][1] else (pts[1], pts[2])
    out["abs72"] = x0 + (tgt - y0) * (x1 - x0) / (y1 - y0)
    print("abs_mem for 72%:", out["abs72"], pts)
    st = [(l, R[f"step{i}"]["tp"]) for i, (l, _) in enumerate(STEPS)]
    px = b["px"]
    rec = lambda u: "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"
    parts = "；".join(f"{l} {t1 - t0:+.2f}" for (_, t0), (l, t1) in zip(st, st[1:]))
    u0, u1 = st[0][1] / px - 1, st[-1][1] / px - 1
    edge = min((0.5, 0.25, -0.25, -0.5), key=lambda e: abs(u1 - e))
    out["steps"] = [{"label": l, "tp": t} for l, t in st]
    out["attr"] = (f"v0.3.6 目標價變動拆解（逐項累加，nvidia/sens_nvidia.py 整本重算）：v0.3.5 加權目標價 {st[0][1]:.2f} 美元 → {parts} → v0.3.6 {st[-1][1]:.2f} 美元"
                   f"（合計 {st[-1][1] - st[0][1]:+.2f}，{st[-1][1] / st[0][1] - 1:+.1%}）；上檔空間 {u0:+.1%} → {u1:+.1%}，評等 {rec(u0)} → {rec(u1)}（{'未翻轉' if rec(u0) == rec(u1) else '翻轉'}）；"
                   f"距最近級距邊界 {edge:+.0%} 為 {abs(u1 - edge) * 100:.1f} 個百分點（加權目標價 {px * (1 + edge):.2f} 美元時翻轉）。")
    print(out["attr"])
json.dump(out, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)

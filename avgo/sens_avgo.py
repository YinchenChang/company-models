# Broadcom 敏感度快照（v0.5.6）：每個案例以環境變數 AVGO_P 覆寫三情境主引擎的案例參數（同 Sens 頁既有「一次動一個輸入」案例的覆寫），整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 avgo/sens.json；正式建置（build_avgo.py → v056_post.py）讀入 Sens 頁「加權目標價相對基準」表。
# 這與 Sens 頁既有案例的差別：既有案例只動情境 A 並看 FY2027／FY2028 EPS；本快照三情境同時套用，並重算 DCF 與 Comparable，得到加權目標價。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 avgo/sens_avgo.py [暫存資料夾，預設 /tmp/work_avgo/sens]；並行時設 LO_PROFILE＝file:///tmp/lo_avgo
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_avgo/sens"
BUILD = os.path.join(HERE, "build_avgo.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("AVGO_P", "AVGO_SENS")}
env0["COORD_RUN"] = "1"; env0["AVGO_SENS"] = "1"; env0["PYTHONDONTWRITEBYTECODE"] = "1"   # 不加 Coord 頁、不跑 v056_post
# (鍵, 白話標籤, 低案覆寫, 高案覆寫)；覆寫與 build_avgo.py Sens 頁 cases 相同
CASES = [
    ("uk", "Broadcom XPU 顆數 ×0.9／×1.1（FY2027 起）", dict(uf="uk_lo"), dict(uf="uk_hi")),
    ("sh_tpu", "Broadcom 占 Google TPU 顆數份額 −15／＋10 個百分點（FY2027 起）", dict(sht="(-0.15)"), dict(sht="0.10")),
    ("sh_oth", "Broadcom 占其他 ASIC 顆數份額 −15／＋15 個百分點（FY2027 起）", dict(sho="(-0.15)"), dict(sho="0.15")),
    ("core", "XPU 核心售價 ×0.8／×1.2（FY2027 起）", dict(pfc="pfc_lo"), dict(pfc="pfc_hi")),
    ("f_hbm", "代購 HBM 比例 30%／100%", dict(fh="f_hbm_lo"), dict(fh="f_hbm_hi")),
    ("m_hbm", "代購 HBM 毛利率 0%／20%", dict(mh="m_hbm_lo"), dict(mh="m_hbm_hi")),
    ("net", "AI 網通營收 ×0.8／×1.2（FY2027 起）", dict(nf="0.8"), dict(nf="1.2")),
    ("sw", "基礎架構軟體成長 低／高（FY2027 +8%、之後 +4%；+18%、+12%）", dict(gs27="g_sw27_lo", gs="g_sw_lo"), dict(gs27="g_sw27_hi", gs="g_sw_hi")),
    ("opx", "營業費用額外占營收比 −2／＋2 個百分點", dict(opxa="opx_adj_lo"), dict(opxa="opx_adj_hi")),
]


def lab(ws, prefix):
    return next(r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix))


def run(jobs, sub):
    raw, rc = os.path.join(TMP, sub, "raw"), os.path.join(TMP, sub, "rc")
    shutil.rmtree(os.path.join(TMP, sub), ignore_errors=True); os.makedirs(raw); os.makedirs(rc)
    for nm, ov in jobs.items():
        env = dict(env0)
        if ov: env["AVGO_P"] = json.dumps(ov)
        subprocess.run([sys.executable, BUILD, os.path.join(raw, f"{nm}.xlsx")], check=True, env=env, stdout=subprocess.DEVNULL)
    subprocess.run(["soffice"] + ([f"-env:UserInstallation={os.environ['LO_PROFILE']}"] if os.environ.get("LO_PROFILE") else [])
                   + ["--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", rc] + sorted(glob.glob(os.path.join(raw, "*.xlsx"))),
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return rc


def read(rc, nm):
    wb = openpyxl.load_workbook(os.path.join(rc, f"{nm}.xlsx"), data_only=True)
    V = wb["Val"]
    errs = sum(1 for ws in wb for row in ws.iter_rows() for x in row if isinstance(x.value, str) and x.value.startswith(("#", "Err:")))
    assert errs == 0, (nm, errs)
    return dict(tp=V.cell(lab(V, "目標價（DCF 與 Comparable 加權）"), 2).value, dcf=V.cell(lab(V, "DCF 目標價（情境加權）"), 2).value,
                comp=V.cell(lab(V, "Comparable 目標價（相對倍數調整後）"), 2).value)


jobs = {"base": {}}
for k, _, lo, hi in CASES:
    jobs[f"{k}__lo"] = lo; jobs[f"{k}__hi"] = hi
rc = run(jobs, "cases")
R = {nm: read(rc, nm) for nm in jobs}
b = R["base"]
out = {"_說明": "加權目標價敏感度整本重算快照（avgo/sens_avgo.py）；正式建置讀入 Sens 頁「加權目標價相對基準」表", "_base": b, "biz": {}}
for k, label, _, _ in CASES:
    out["biz"][k] = {"label": label, "低": R[f"{k}__lo"], "高": R[f"{k}__hi"]}
    print(f"{k:7s} 低 {R[f'{k}__lo']['tp'] / b['tp'] - 1:+.2%}  高 {R[f'{k}__hi']['tp'] / b['tp'] - 1:+.2%}")
print("基準", b)
json.dump(out, open(os.path.join(HERE, "sens.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

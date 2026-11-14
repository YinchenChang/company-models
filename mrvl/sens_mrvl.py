# Marvell 敏感度快照（v0.3.5）：每個案例以環境變數 MRVL_P 覆寫三情境主引擎的案例參數（同 Sens 頁既有「一次動一個輸入」案例的覆寫），整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 mrvl/sens.json；正式建置（build_mrvl.py → v035_post.py）讀入 Sens 頁「加權目標價相對基準」表。
# 與 Sens 頁既有案例的差別：既有案例只動情境 A 並看 FY2028／FY2029 EPS；本快照三情境同時套用，並重算 DCF 與 Comparable，得到加權目標價。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。樣板 avgo/sens_avgo.py。
# 用法：python3 mrvl/sens_mrvl.py [暫存資料夾，預設 /tmp/work_mrvl/sens]；並行時設 LO_PROFILE＝file:///tmp/lo_mrvl
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_mrvl/sens"
BUILD = os.path.join(HERE, "build_mrvl.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("MRVL_P", "MRVL_SENS", "SIL_COORD_ALT")}
env0["COORD_RUN"] = "1"; env0["MRVL_SENS"] = "1"; env0["PYTHONDONTWRITEBYTECODE"] = "1"   # 不加 Coord 頁、不跑 v035_post
# (鍵, 白話標籤, 低案覆寫, 高案覆寫)；覆寫與 build_mrvl.py Sens 頁 cases 相同；None＝該側無案例
CASES = [
    ("uk", "Marvell XPU 顆數係數 ×0.8／×1.2（CY2027 起）", dict(uf="uk_lo"), dict(uf="uk_hi")),
    ("sho", "Marvell 占其他 ASIC 份額 −3／＋6 個百分點（CY2027 起）", dict(sho="(-0.03)"), dict(sho="0.06")),
    ("sht", "Marvell 占 AWS Trainium 份額 −5 個百分點（CY2027 起；只有低）", dict(sht="(-0.05)"), None),
    ("shp", "Marvell 占 Google TPU 份額 ＋5 個百分點（CY2027 起；只有高）", None, dict(shp="0.05")),
    ("pfc", "XPU 核心售價 ×0.8／×1.2（CY2027 起）", dict(pfc="pfc_lo"), dict(pfc="pfc_hi")),
    ("fh", "HBM 代購比例 低／高", dict(fh="f_hbm_lo"), dict(fh="f_hbm_hi")),
    ("shd", "光晶片組份額 −10／＋5 個百分點（CY2027 起）", dict(shd="(-0.10)"), dict(shd="0.05")),
    ("dspf", "每模組光晶片組金額 ×0.8（CY2027 起；只有低）", dict(dspf="0.8"), None),
    ("lpoa", "線性直驅占比 ＋10 個百分點（CY2027 起；只有低）", dict(lpoa="0.10"), None),
    ("of", "Marvell AI 光晶片組 ×1.2（CY2027 起；只有高）", None, dict(of="of_hi")),
    ("a28", "XPU attach FY2028 低／高", dict(a28="att28_lo"), dict(a28="att28_hi")),
    ("gsw28", "交換器 FY2028 年增 低（只有低）", dict(gsw28="g_sw28_lo"), None),
    ("gco28", "通訊與其他 FY2028 年增 低（只有低）", dict(gco28="g_co28_lo"), None),
    ("opxa", "營業費用額外占營收 ＋2 個百分點（只有低）", dict(opxa="opx_adj_hi"), None),
]


def lab(ws, prefix):
    return next(r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix))


def run(jobs, sub):
    raw, rc = os.path.join(TMP, sub, "raw"), os.path.join(TMP, sub, "rc")
    shutil.rmtree(os.path.join(TMP, sub), ignore_errors=True); os.makedirs(raw); os.makedirs(rc)
    for nm, ov in jobs.items():
        env = dict(env0)
        if ov: env["MRVL_P"] = json.dumps(ov)
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
    if lo is not None: jobs[f"{k}__lo"] = lo
    if hi is not None: jobs[f"{k}__hi"] = hi
rc = run(jobs, "cases")
R = {nm: read(rc, nm) for nm in jobs}
b = R["base"]
out = {"_說明": "加權目標價敏感度整本重算快照（mrvl/sens_mrvl.py）；正式建置讀入 Sens 頁「加權目標價相對基準」表", "_base": b, "biz": {}}
for k, label, lo, hi in CASES:
    out["biz"][k] = {"label": label, "低": R.get(f"{k}__lo"), "高": R.get(f"{k}__hi")}
    f = lambda nm: f"{R[nm]['tp'] / b['tp'] - 1:+.2%}" if nm in R else "—"
    print(f"{k:7s} 低 {f(k + '__lo')}  高 {f(k + '__hi')}")
print("基準", b)
json.dump(out, open(os.path.join(HERE, "sens.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

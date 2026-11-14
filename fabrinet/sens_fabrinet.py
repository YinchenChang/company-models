# Fabrinet 敏感度快照（v0.3.5；樣板 cohr/sens_cohr.py、lite/sens_lite.py）：每個案例以環境變數 CO_OVERRIDE 把 In 頁一組輸入由基準欄換成低欄或高欄（三情境主引擎同時套用；公司變數 rel、cpo、gm 只動情境 B）；
# 「P.參數=值」改主引擎參數），整本重建、LibreOffice 重算，讀加權目標價、DCF 目標價與 Comparable 目標價，寫入 fabrinet/sens.json；
# 正式建置（build_fabrinet.py → v045_post.py）讀入 Sens 頁「加權目標價相對基準」表。
# 與 Sens 頁既有案例的差別：既有案例只動情境 A 並看 FY2027／FY2028 EPS；本快照三情境同時套用，並重算 DCF 與 Comparable，得到加權目標價。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 fabrinet/sens_fabrinet.py [暫存資料夾，預設 /tmp/work_fabrinet/sens]；並行時設 LO_PROFILE＝file:///tmp/lo_fabrinet
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_fabrinet/sens"
BUILD = os.path.join(HERE, "build_fabrinet.py")
env0 = {k: v for k, v in os.environ.items() if k != "CO_OVERRIDE"}
env0["COORD_RUN"] = "1"; env0["PYTHONDONTWRITEBYTECODE"] = "1"   # 不加 Coord 頁、不跑 v035_post
Y3 = lambda k: [f"{k}_{y}" for y in (27, 28, 29)]
LH = lambda keys: ({x: "@lo" for x in keys}, {x: "@hi" for x in keys})
# (鍵, 白話標籤, 低案覆寫, 高案覆寫)；低／高欄取 In 頁 D／F 欄（與 Sens 頁既有案例相同）；None＝該側不做
CASES = [
    ("mf", "AI 光學市場 CY2027 起 ×0.8／×1.2", *LH(["mf"])),
    ("rel", "Fabrinet 光模組有效份額 CY2027–CY2029 相對值 0.80／0.70／0.60 對 1.15／1.25／1.30（只動情境 B；A、C 本身取高、低值）", *LH(Y3("rel"))),
    ("pxr", "單價係數 CY2027–CY2029 0.90／0.80／0.72 對 1.05", *LH(Y3("pxr"))),
    ("gm", "non-GAAP 毛利率 FY2027–FY2029 11.8／11.5／11.3% 對 12.7／12.8／12.8%（只動情境 B）", *LH(Y3("gm"))),
    ("cpo", "Fabrinet 占 CPO 光學金額 CY2027–CY2029 低／高（只動情境 B）", *LH(Y3("cpo"))),
    ("hw", "日曆年上半年占全年 42%／50%（會計年度換算）", *LH(["hw"])),
    ("dci", "DCI 年增 FY2028–FY2029 低／高", *LH(["g_dci_28", "g_dci_29"])),
    ("hpc", "HPC 年增 FY2028–FY2029 低／高", *LH(["g_hpc_28", "g_hpc_29"])),
    ("ci", "通訊基礎設施年增 FY2028–FY2029 低／高", *LH(["g_ci_28", "g_ci_29"])),
    ("aio", "車用、工業與其他年增 FY2028–FY2029 低／高", *LH(["g_aio_28", "g_aio_29"])),
    ("phi", "營業費用成長對營收成長的彈性 0.1／0.6", *LH(["phi"])),
    ("opx", "營業費用額外占營收比 −0.2／＋0.5 個百分點", *LH(["opx_adj"])),
    ("int", "業外淨收益 每年 0.02／0.04 十億美元", *LH(["int_f"])),
    ("tx", "non-GAAP 稅率 4%／6%", *LH(["tx"])),
]


def lab(ws, prefix):
    return next(r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix))


def run(jobs, sub):
    raw, rc = os.path.join(TMP, sub, "raw"), os.path.join(TMP, sub, "rc")
    shutil.rmtree(os.path.join(TMP, sub), ignore_errors=True); os.makedirs(raw); os.makedirs(rc)
    for nm, ov in jobs.items():
        env = dict(env0)
        if ov: env["CO_OVERRIDE"] = ";".join(f"{k}={v}" for k, v in ov.items())
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
out = {"_說明": "加權目標價敏感度整本重算快照（fabrinet/sens_fabrinet.py）；正式建置讀入 Sens 頁「加權目標價相對基準」表", "_base": b, "biz": {}}
for k, label, lo, hi in CASES:
    out["biz"][k] = {"label": label, "低": R.get(f"{k}__lo"), "高": R.get(f"{k}__hi")}
    f = lambda s: f"{R[f'{k}__{s}']['tp'] / b['tp'] - 1:+.2%}" if f"{k}__{s}" in R else "—"
    print(f"{k:5s} 低 {f('lo')}  高 {f('hi')}")
print("基準", b)
json.dump(out, open(os.path.join(HERE, "sens.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

# Coherent 敏感度快照（v0.4.5）：每個案例以環境變數 CO_OVERRIDE 把 In 頁一組輸入由基準欄換成低欄或高欄（三情境主引擎同時套用，與 Sens 頁既有「一次動一個輸入」案例的低／高設定相同），
# 整本重建、LibreOffice 重算，讀加權目標價、DCF 目標價與 Comparable 目標價，寫入 cohr/sens.json；正式建置（build_cohr.py → v045_post.py）讀入 Sens 頁「加權目標價相對基準」表。
# 這與 Sens 頁既有案例的差別：既有案例只動情境 A 並看 FY2027／FY2028 EPS；本快照三情境同時套用，並重算 DCF 與 Comparable，得到加權目標價。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 cohr/sens_cohr.py [暫存資料夾，預設 /tmp/work_cohr/sens]；並行時設 LO_PROFILE＝file:///tmp/lo_cohr
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_cohr/sens"
BUILD = os.path.join(HERE, "build_cohr.py")
env0 = {k: v for k, v in os.environ.items() if k != "CO_OVERRIDE"}
env0["COORD_RUN"] = "1"; env0["PYTHONDONTWRITEBYTECODE"] = "1"   # 不加 Coord 頁、不跑 v045_post
Y3 = lambda k: [f"{k}_{y}" for y in (27, 28, 29)]
# (鍵, 白話標籤, 低案的 In 鍵, 高案的 In 鍵)；低／高欄取 In 頁 D／F 欄（與 Sens 頁既有案例相同）
CASES = [
    ("rel", "Coherent 可插拔有效份額 CY2027–CY2029 相對值 0.90／0.85／0.80 對 1.25／1.35／1.40", Y3("rel"), Y3("rel")),
    ("pxr", "單價係數 CY2027–CY2029 0.90／0.80／0.72 對 1.05", Y3("pxr"), Y3("pxr")),
    ("gm", "non-GAAP 毛利率 FY2027–FY2029 38.5／37.5／36.5% 對 42.5／44／45%", Y3("gm"), Y3("gm")),
    ("cpo", "Coherent 占 CPO 光學金額 CY2027–CY2029 0% 對 25／30／35%", Y3("cpo"), Y3("cpo")),
    ("mf", "AI 光模組市場 CY2027 起 ×0.8／×1.2", ["mf"], ["mf"]),
    ("hw", "日曆年上半年占會計年度的比重 42%／50%（資料中心由日曆年換會計年度）", ["hw"], ["hw"]),
    ("com", "通訊營收年增 FY2027–FY2029 10／0／−5% 對 30／20／15%", Y3("g_com"), Y3("g_com")),
    ("ind", "工業營收年增 FY2027–FY2029 低／高（FY2027 −5%／＋8%；之後 −5%／＋10%）", Y3("g_ind"), Y3("g_ind")),
    ("phi", "營業費用成長對營收成長的彈性 0.4／0.9", ["phi"], ["phi"]),
    ("opx", "營業費用額外占營收比 −1／＋2 個百分點", ["opx_adj"], ["opx_adj"]),
    ("tx", "non-GAAP 稅率 18%／20%", ["tx"], ["tx"]),
    ("q2", "FY2027 第二季資料中心季增 11%／27%（第二季營收對指引的落點）", ["q2_dg"], ["q2_dg"]),
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
    jobs[f"{k}__lo"] = {x: "@lo" for x in lo}; jobs[f"{k}__hi"] = {x: "@hi" for x in hi}
rc = run(jobs, "cases")
R = {nm: read(rc, nm) for nm in jobs}
b = R["base"]
out = {"_說明": "加權目標價敏感度整本重算快照（cohr/sens_cohr.py）；正式建置讀入 Sens 頁「加權目標價相對基準」表", "_base": b, "biz": {}}
for k, label, _, _ in CASES:
    out["biz"][k] = {"label": label, "低": R[f"{k}__lo"], "高": R[f"{k}__hi"]}
    print(f"{k:5s} 低 {R[f'{k}__lo']['tp'] / b['tp'] - 1:+.2%}  高 {R[f'{k}__hi']['tp'] / b['tp'] - 1:+.2%}")
print("基準", b)
json.dump(out, open(os.path.join(HERE, "sens.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

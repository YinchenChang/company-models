# BESI 敏感度快照（v0.2.4 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖）：每個案例以環境變數 BESI_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 besi/sens.json；正式建置（build_besi.py → v024_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 besi/sens_besi.py [暫存資料夾，預設 /tmp/besi_w/sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_besi）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/besi_w/sens"
BUILD = os.path.join(HERE, "build_besi.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("BESI_OVERRIDE", "BESI_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄，*倍數＝基準乘倍數（三情境同時受影響）
_3 = lambda k, v: ";".join(f"{k}_{s}={v}" for s in "ABC")
_ATT = lambda m: ";".join(f"att{s}_{y}=*{m}" for s in "ABC" for y in (28, 29, 30))
CASES = [
    ("att", "邏輯混合鍵合導入比例 ×0.7／×1.3（2028 起）", _ATT(0.7), _ATT(1.3), "三情境 2028–2030 的混合鍵合占 AI 先進封裝晶圓比例同乘 0.7／1.3（導入比例無直接資料，Assumed）"),
    ("hbsh", "FY2026 混合鍵合占營收 18%／35%", "hb_sh26=@lo", "hb_sh26=@hi", "基準 25%（公司不揭露，Assumed）；同時改變邏輯校準係數與主流業務基期"),
    ("asp", "HBM 混合鍵合機單價 €2.0M／€3.5M", "asp=@lo", "asp=@hi", "基準 €2.5M（公司不揭露，Assumed）"),
    ("shm", "BESI 占 HBM 混合鍵合機採購 −10／＋10 個百分點", _3("shm", "@lo"), _3("shm", "@hi"), "基準 A 60%／B 50%／C 35%（Assumed）"),
    ("shl", "邏輯混合鍵合相對份額 −0.1／＋0.1", _3("shl", "@lo"), _3("shl", "@hi"), "基準 A 1.0／B 0.95／C 0.8（相對 FY2026；Assumed）"),
    ("uph", "混合鍵合機有效產出 600／1,500 顆／時", "uph=@lo", "uph=@hi", "基準 1,000（公司乾循環最高 2,000，Assumed）"),
    ("dlt", "已裝機產能汰換比例 5%／15%", "dlt=@lo", "dlt=@hi", "基準 10%（Assumed）"),
    ("el", "主流業務對設備市場年增的彈性 0.5／1.5", "el=@lo", "el=@hi", "基準 1.0（Assumed）"),
    ("prem", "混合鍵合毛利率溢價 0／15 個百分點", "gm_prem=@lo", "gm_prem=@hi", "基準 8 個百分點（Assumed）"),
    ("phi", "營業費用變動比例 25%／60%", "phi=@lo", "phi=@hi", "基準 40%（Assumed）"),
    ("q4g", "第四季營收相對第三季 −5%／＋5%", "q4g=@lo", "q4g=@hi", "公司指引持平 ±5%；改變 FY2026 錨點、校準係數與主流業務基期"),
    ("wts", "情境權重（全給保守 C／全給樂觀 A）", "tp_wA=0;tp_wB=0;tp_wC=1", "tp_wA=1;tp_wB=0;tp_wC=0", "基準權重 A 0.3／B 0.5／C 0.2"),
]


def lab(ws, prefix, start=1):
    return next(r for r in range(start, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip() == prefix)


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
    V = wb["Val"]
    return dict(tp=V.cell(lab(V, "目標價（DCF 與 Comparable 加權）"), 2).value, dcf=V.cell(lab(V, "DCF 目標價（情境加權）"), 2).value)


jobs0 = {"base": {"BESI_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"BESI_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"BESI_OVERRIDE": hi}
rc = run(JOBS, "cases")
biz = {}
for nm in JOBS:
    k, side = nm.split("__")
    r = read(rc, nm)
    biz.setdefault(k, {})["低" if side == "lo" else "高"] = r
    print(f"{nm:14s} 加權 {r['tp']:.2f}（{r['tp'] / b['tp'] - 1:+.2%}）DCF {r['dcf']:.2f}")
out = {"_base": b, "cases": [{"key": k, "label": l, "setting": s} for k, l, _, _, s in CASES], "biz": biz}
json.dump(out, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
print("基準", b)

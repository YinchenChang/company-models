# Kioxia 敏感度快照（v0.5.5 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖）：每個案例以環境變數 KIOXIA_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 kioxia/sens.json；正式建置（build_kioxia.py → v055_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 kioxia/sens_kioxia.py [暫存資料夾，預設 /tmp/kioxia_w/sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_kioxia）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/kioxia_w/sens"
BUILD = os.path.join(HERE, "build_kioxia.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("KIOXIA_OVERRIDE", "KIOXIA_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄（三情境同時受影響）；pxm／shm＝主引擎市場層營收價格係數（第三季起）／份額係數
_CD = lambda s: ";".join(f"cd_{x}={s}" for x in "ABC")
_CAP = lambda s: ";".join(f"cap_{y}={s}" for y in (2026, 2027, 2028, 2029))
CASES = [
    ("px", "NAND 價格（第三季起市場層營收 ×0.8／×1.2）", "pxm=0.8", "pxm=1.2", "三情境的市場層營收同乘（位元與成本不變）"),
    ("sh", "Kioxia 份額 ×0.9／×1.1（營收與位元）", "shm=0.9", "shm=1.1", "三情境的市場層營收與位元同乘"),
    ("fx", "匯率 145／170 日圓", "fx=@lo", "fx=@hi", "基準 160；營收以美元計、成本以日圓計"),
    ("cd", "每位元成本年降 −4／+4 個百分點", _CD("@lo"), _CD("@hi"), "三情境各取 In 頁低／高欄（A 14%、B 12%、C 10% 各 ±4pt）"),
    ("oxv", "營業費用變動率 0.5%／2%", "ox_var=@lo", "ox_var=@hi", "基準 1%"),
    ("tax", "有效稅率 28%／35%", "tax=@lo", "tax=@hi", "基準 32%"),
    ("da", "折舊攤銷 2,800／3,800 億日圓／年", "da_y=@lo", "da_y=@hi", "基準 3,200 億日圓"),
    ("cap", "資本支出 低／高（FY2026–FY2029）", _CAP("@lo"), _CAP("@hi"), "In 頁 cap_2026–cap_2029 低／高欄"),
    ("pay", "股東回饋 0／50% 自由現金流", "payout=@lo", "payout=@hi", "基準 20%（以買回計）"),
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


jobs0 = {"base": {"KIOXIA_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"KIOXIA_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"KIOXIA_OVERRIDE": hi}
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

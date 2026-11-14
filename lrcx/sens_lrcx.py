# 科林研發 敏感度快照（v0.3.5 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖）：每個案例以環境變數 LRCX_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 lrcx/sens.json；正式建置（build_lrcx.py → v035_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 lrcx/sens_lrcx.py [暫存資料夾，預設 /tmp/b3eq/lrcx_sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_lrcx）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/b3eq/lrcx_sens"
BUILD = os.path.join(HERE, "build_lrcx.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("LRCX_OVERRIDE", "LRCX_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄，*倍數＝基準乘倍數，=公式＝改公式（三情境同時受影響）
_E = lambda m: ";".join(f"{k}{s}_{y}=*{m}" for k in ("adv", "ofl", "dram", "nand") for s in "ABC" for y in (27, 28, 29))
_3 = lambda k, v: ";".join(f"{k}_{s}={v}" for s in "ABC")
CASES = [
    ("wfe", "晶圓前段設備市場 −15%／＋15%（CY2027 起）", _E(0.85), _E(1.15), "設備市場模型三情境的先進邏輯、其他代工邏輯、DRAM（含 HBM）、NAND 設備 CY2027–CY2029 同乘 0.85／1.15"),
    ("drift", "份額漂移每年 −2%／＋2%（CY2027 起）", "drift=@lo", "drift=@hi", "相對 WFE 分項的份額漂移（基準 0）"),
    ("h1", "上半年占日曆年比例 45%／55%", "h1s=@lo", "h1s=@hi", "日曆年換算會計年度（基準 50%）"),
    ("cut", "情境 C 出口管制折損 0%／40%", "cutC=@lo", "cutC=@hi", "情境 C 自 CY2028 起中國系統營收折損（基準 20%）"),
    ("cs27", "CSBG CY2027 年增 10%／20%", "g_cs27=@lo", "g_cs27=@hi", "基準 15%（公司：逾 15%）"),
    ("csg", "CSBG CY2028–29 年增 −4／＋4 個百分點", _3("g_cs", "@lo"), _3("g_cs", "@hi"), "基準 A 12%／B 6%／C −3%"),
    ("phi", "營業費用變動比例 15%／50%", "opex_var==phi_lo*a26_opx/a26_rev;opex_fix==a26_opx*(1-phi_lo)", "opex_var==phi_hi*a26_opx/a26_rev;opex_fix==a26_opx*(1-phi_hi)", "基準由兩點推得"),
    ("tax", "稅率 13%／19%", "tax=@lo", "tax=@hi", "基準 16.5%（公司非 GAAP 指引）；DCF 稅率同步"),
    ("oth", "業外收益每年 0／0.2 十億美元", "oth_inc=@lo", "oth_inc=@hi", "基準 0.1"),
    ("bb", "庫藏股每年 0%／3%", "bb=@lo", "bb=@hi", "FY2028 起稀釋股數每年減少（基準 1.5%）"),
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


jobs0 = {"base": {"LRCX_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"LRCX_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"LRCX_OVERRIDE": hi}
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

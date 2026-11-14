# 應用材料 敏感度快照（v0.4.5 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖）：每個案例以環境變數 AMAT_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 amat/sens.json；正式建置（build_amat.py → v045_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 amat/sens_amat.py [暫存資料夾，預設 /tmp/b3eq/amat_sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_amat）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/b3eq/amat_sens"
BUILD = os.path.join(HERE, "build_amat.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("AMAT_OVERRIDE", "AMAT_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄，*倍數＝基準乘倍數，=公式＝改公式（三情境同時受影響）
_E = lambda m: ";".join(f"{k}{s}_{y}=*{m}" for k in ("la", "lo", "of", "dr", "nd") for s in "ABC" for y in (27, 28, 29))
_3 = lambda k, v: ";".join(f"{k}_{s}={v}" for s in "ABC")
CASES = [
    ("wfe", "晶圓廠設備市場 −15%／＋15%（CY2027 起）", _E(0.85), _E(1.15), "設備市場模型三情境的五個分項（先進邏輯、其他代工邏輯、其他前段、DRAM 含 HBM、NAND）CY2027–CY2029 同乘 0.85／1.15"),
    ("sf", "份額係數 0.95／1.05", "sf=@lo", "sf=@hi", "FY2027 起應用材料份額 −5%／+5%（基準 1.00）"),
    ("dr", "第四季 DRAM 占半導體系統 26%／34%", "q4_dr=@lo", "q4_dr=@hi", "改變 FY2026 終端市場權重（基準 30%）"),
    ("fyw", "會計年度落在前一年的比例 1/12／3/12", "fyw=@lo", "fyw=@hi", "基準 2/12"),
    ("cv", "營業成本變動比例 75%／95%", "cv==phic_lo*cogs25/a25_rev;cF==(1-phic_lo)*cogs25", "cv==phic_hi*cogs25/a25_rev;cF==(1-phic_hi)*cogs25", "基準由兩點推得（FY2025 與 FY2026 錨點）"),
    ("ov", "營業費用變動比例 35%／75%", "ov==phi_lo*opex25/a25_rev;oF==(1-phi_lo)*opex25", "ov==phi_hi*opex25/a25_rev;oF==(1-phi_hi)*opex25", "基準由兩點推得"),
    ("ags", "AGS 年成長 −4／＋4 個百分點", _3("g_ags", "@lo"), _3("g_ags", "@hi"), "基準 A 15%／B 11%／C 7%"),
    ("oth", "其他（顯示器）年成長 −10／＋10 個百分點", _3("g_oth", "@lo"), _3("g_oth", "@hi"), "基準 A 0%／B −5%／C −10%"),
    ("aff", "關係企業規則三情境都恢復（每年 US$600M）", "aff_A=1;aff_B=1", None, "基準只有情境 C 恢復；只有低案"),
    ("nioi", "淨利 ÷ 營業利益 0.85／FY2026 錨點值", "nioi=@lo", "nioi=@hi", "基準＝FY2025 實績比例"),
    ("bb", "淨回購 0%／2.5%／年", "bb=@lo", "bb=@hi", "稀釋股數每年減少（基準 1%）"),
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


jobs0 = {"base": {"AMAT_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"AMAT_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"AMAT_OVERRIDE": hi}
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

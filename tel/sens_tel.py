# 東京威力科創 敏感度快照（v0.3.5 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖）：每個案例以環境變數 TEL_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 tel/sens.json；正式建置（build_tel.py → tel_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 tel/sens_tel.py [暫存資料夾，預設 /tmp/tel_w/sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_n5_tel）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/tel_w/sens"
BUILD = os.path.join(HERE, "build_tel.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("TEL_OVERRIDE", "TEL_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄（三情境同時受影響）
_3 = lambda k, v, ss="ABC": ";".join(f"{k}_{s}={v}" for s in ss)
CASES = [
    ("exc", "新設備相對設備市場的年超額成長 低／高", _3("exc", "@lo"), _3("exc", "@hi"), "A 0%／4%、B −2%／+2%、C −4%／0%（基準 2%／0%／−2%；Assumed）"),
    ("chh", "中國新設備營收水準變化 低／高（情境 B、C）", _3("chh", "@lo", "BC"), _3("chh", "@hi", "BC"), "B −10%／0%、C −30%／−10%（基準 −5%／−20%；情境 A 中國持平）"),
    ("chn", "中國占新設備營收 30%／40%", "chn=@lo", "chn=@hi", "基準 34%（公司：FY2026 上半中國占營收約 34%；Interested-party）"),
    ("gfso", "Field Solutions 與其他年增 低／高", _3("gfso", "@lo"), _3("gfso", "@hi"), "A 8%／16%、B 4%／12%、C 0%／6%（基準 12%／8%／3%；Analogy）"),
    ("tail", "CY2030 設備市場相對 CY2029 −10%／＋10%", "g_tail=@lo", "g_tail=@hi", "基準 0%（Assumed）"),
    ("dm", "邊際營業利益率 −3／＋3 個百分點", "dm=@lo", "dm=@hi", "基準 0（Assumed；價格最適化、組合）"),
    ("madj", "情境 C 的營業利益率額外調整 −3／0 個百分點", "madjC=@lo", "madjC=@hi", "基準 −1.5 個百分點（只影響情境 C）"),
    ("nioi", "淨利 ÷ 營業利益 0.70／0.78", "nioi=@lo", "nioi=@hi", "基準＝公司 FY2026 指引隱含 0.740"),
    ("bb", "平均股數年減幅 0%／1.5%", "bb=@lo", "bb=@hi", "基準 0.8%（Assumed）"),
    ("anc", "FY2026 實績相對公司指引 −3%／＋5%", "g_anc=@lo", "g_anc=@hi", "基準＝公司指引（10-30 版）"),
    ("wts", "情境權重（全給保守 C／全給樂觀 A）", "tp_wA=0;tp_wB=0;tp_wC=1", "tp_wA=1;tp_wB=0;tp_wC=0", "基準權重 A 0.25／B 0.5／C 0.25"),
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
    S = wb["Sens"]; fr = next((r for r in range(1, S.max_row + 1) if S.cell(r, 1).value == "輸入指紋（本檔即時重算）"), None)
    return dict(tp=V.cell(lab(V, "目標價（DCF 與 Comparable 加權）"), 2).value, dcf=V.cell(lab(V, "DCF 目標價（情境加權）"), 2).value, fp=S.cell(fr, 4).value if fr else None)


jobs0 = {"base": {"TEL_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"TEL_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"TEL_OVERRIDE": hi}
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

# 科磊 敏感度快照（v0.4.5 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖）：每個案例以環境變數 KLAC_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 klac/sens.json；正式建置（build_klac.py → v045_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 klac/sens_klac.py [暫存資料夾，預設 /tmp/b3eq/klac_sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_klac）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/b3eq/klac_sens"
BUILD = os.path.join(HERE, "build_klac.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("KLAC_OVERRIDE", "KLAC_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄，*倍數＝基準乘倍數，=公式＝改公式（三情境同時受影響）
_3 = lambda k, v: ";".join(f"{k}_{s}={v}" for s in "ABC")
CASES = [
    ("wfe", "晶圓前段設備市場 −15%／＋15%（CY2027 起，含先進封裝）", "_eqm=0.85", "_eqm=1.15", "設備市場模型三情境各分項（代工邏輯、記憶體、先進封裝）CY2027–CY2029 同乘 0.85／1.15"),
    ("pcu", "製程控制強度年增 低／高", "pc_u=@lo", "pc_u=@hi", "基準約 2.9%／年（公司 2030 年占 WFE 提升目標推得）"),
    ("mem", "記憶體強度相對倍數 低／高", "m_mem=@lo", "m_mem=@hi", "記憶體 WFE 的製程控制強度 ÷ 代工邏輯（基準由 CY2026 錨點推得）"),
    ("kap", "先進封裝成長倍數 1.0／2.5", "k_ap=@lo", "k_ap=@hi", "先進封裝製程控制成長 ÷ 封裝設備市場成長（基準 2.0）"),
    ("ap26", "CY2026 先進封裝營收 1.0／1.4 十億美元", "ap26=@lo", "ap26=@hi", "基準 1.1（公司說法）"),
    ("h2", "CY2026 下半營收增幅 低／高（第二季指引區間）", "h2up=@lo", "h2up=@hi", "錨點 CY2026 隨第二季指引低／高"),
    ("cut", "CY2027 起中國系統營收下降 25%／50%", "chn_cut=@lo", "chn_cut=@hi", "出口管制（基準 0）"),
    ("svc", "服務營收年增 −2／＋2 個百分點", _3("g_svc", "@lo"), _3("g_svc", "@hi"), "基準 A 14%／B 12%／C 10%"),
    ("gmi", "增量毛利率 60%／65%", "gm_inc=@lo", "gm_inc=@hi", "毛利變動 ÷ 營收變動（基準由兩點推得約 64.6%）"),
    ("opx", "營業費用 FY2028–29 年增 −3／＋3 個百分點", _3("g_opex", "@lo"), _3("g_opex", "@hi"), "基準 A 8%／B 5%／C 2%"),
    ("tax", "有效稅率 13%／17%", "tax=@lo", "tax=@hi", "基準 14.5%"),
    ("bb", "買回 0%／2.5%／年", "bb=@lo", "bb=@hi", "FY2028–29 稀釋股數年減"),
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


jobs0 = {"base": {"KLAC_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"KLAC_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"KLAC_OVERRIDE": hi}
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

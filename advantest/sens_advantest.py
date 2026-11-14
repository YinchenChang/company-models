# Advantest 敏感度快照（v0.7.5 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖）：每個案例以環境變數 ADV_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 advantest/sens.json；正式建置（build_advantest.py → v045_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 advantest/sens_advantest.py [暫存資料夾，預設 /tmp/work_advantest/sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_advantest）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_advantest/sens"
BUILD = os.path.join(HERE, "build_advantest.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("ADV_OVERRIDE", "ADV_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄（三情境同時受影響）
_T = lambda s: ";".join(f"tt_{k}={s}" for k in ("bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth"))
_G = lambda s: ";".join(f"{k}_{x}={s}" for k in ("g_oth", "g_svc") for x in "ABC")
CASES = [
    ("soc_mk", "SoC 測試機市場 CY2026 取 11.0／12.0 十億美元", "mk_soc26=@lo", "mk_soc26=@hi", "校準錨（Advantest 自估區間；市占 sh_soc 隨之重算）"),
    ("nonai", "非 AI SoC 市場 3.7／4.9 十億美元", "nonai26=@lo", "nonai26=@hi", "AI 占比較高（低）或較低（高）；每加權顆測試支出隨之校準"),
    ("soc_sh", "Advantest SoC 市占 ±5 個百分點", "sh_soc=@lo", "sh_soc=@hi", "FY2026 起三情境（情境 C 另有 FY2028–29 自有市占）"),
    ("mem_sh", "Advantest 記憶體測試機市占 ±5 個百分點", "sh_mem=@lo", "sh_mem=@hi", "三情境"),
    ("hbm", "記憶體 HBM 比重 30%／70%", "w_hbm=@lo", "w_hbm=@hi", "記憶體測試機需求隨 HBM 位元的比重（基準 50%）"),
    ("fx", "匯率 ¥140／¥160", "fx=@lo", "fx=@hi", "基準 ¥150；市占 sh_soc 為公式（含匯率）"),
    ("nioi", "淨利÷營業利益 0.72／指引隱含 0.78", "nioi=@lo", "nioi=@hi", "高案含戰略投資評價益（一次性）"),
    ("tt", "每顆測試時間係數 低／高", _T("@lo"), _T("@hi"), "各平台 tt 同時取低／高"),
    ("nonai_g", "非 AI SoC 市場年增 0%／8%", "g_nonai=@lo", "g_nonai=@hi", "基準 4%"),
    ("oth_g", "其他系統與服務成長 低／高", _G("@lo"), _G("@hi"), "FY2027–29 三情境各自低／高"),
    ("cy30", "2030 年顆數與記憶體變化 −10%／+10%", "cy30g=@lo", "cy30g=@hi", "只影響 FY2029 的 1/4 與延伸起點"),
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


jobs0 = {"base": {"ADV_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"ADV_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"ADV_OVERRIDE": hi}
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

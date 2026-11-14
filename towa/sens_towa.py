# TOWA 敏感度快照（v0.3.5 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖）：每個案例以環境變數 TOWA_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 towa/sens.json；正式建置（build_towa.py → v035_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 towa/sens_towa.py [暫存資料夾，預設 /tmp/towa_w/sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_towa）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/towa_w/sens"
BUILD = os.path.join(HERE, "build_towa.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("TOWA_OVERRIDE", "TOWA_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄（三情境同時受影響）
_3 = lambda k, v: ";".join(f"{k}_{s}={v}" for s in "ABC")
CASES = [
    ("ai", "FY2025 AI 相關封膠營收錨點 7／13 十億日圓", "ai_25=@lo", "ai_25=@hi", "基準 10（公司未揭露金額，Assumed；校準錨）"),
    ("w", "AI 相關中 HBM 比重 35%／65%", "w_hbm=@lo", "w_hbm=@hi", "基準 50%（Assumed）"),
    ("dlt", "已裝機封膠產能汰換比例 5%／15%", "dlt=@lo", "dlt=@hi", "基準 10%（Assumed）"),
    ("tail", "CY2030–31 HBM 堆疊與 CoWoS 片數年增 0%／20%", "g_tail=@lo", "g_tail=@hi", "基準 10%（Assumed）"),
    ("pkt", "CY2030 非 AI 封裝設備年增 0%／10%", "g_pk_tail=@lo", "g_pk_tail=@hi", "基準 5%（Assumed）"),
    ("tss", "新事業（TSS 售後等）年增 0%／6%", "g_tss=@lo", "g_tss=@hi", "基準 3%"),
    ("shl", "AI 封膠市占流失 低／高", _3("shl", "@lo"), _3("shl", "@hi"), "A 0／3pt、B 0／6pt、C 4／12pt（基準 0／3／8pt）"),
    ("ima", "增量營益率調整 低／高", _3("ima", "@lo"), _3("ima", "@hi"), "A +1／+5pt、B −2／+2pt、C −8／−2pt（基準 +3／0／−5pt）"),
    ("im", "增量營益率（迴歸斜率）±5 個百分點", "im=@lo", "im=@hi", "基準 27.0%（Hist 迴歸）"),
    ("nioi", "淨利 ÷ 營業利益 0.65／0.72", "nioi=@lo", "nioi=@hi", "基準＝公司 FY2026 指引隱含 0.684"),
    ("anc", "FY2026 錨點相對公司指引 −5%／＋8%", "g_anc=@lo", "g_anc=@hi", "基準＝公司全年指引"),
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


jobs0 = {"base": {"TOWA_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"TOWA_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"TOWA_OVERRIDE": hi}
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

# 泰瑞達 敏感度快照（v0.4.5 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖）：每個案例以環境變數 TER_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 ter/sens.json；正式建置（build_ter.py → v045_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 ter/sens_ter.py [暫存資料夾，預設 /tmp/work_ter/sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_ter）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_ter/sens"
BUILD = os.path.join(HERE, "build_ter.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("TER_OVERRIDE", "TER_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄（三情境同時受影響）
_G = lambda s: ";".join(f"g_pt_{x}={s}" for x in "ABC")
_RM = lambda s: ";".join(f"rm_{y}={s}" for y in (2027, 2028, 2029))
CASES = [
    ("soc_mk", "SoC 測試機市場 −15%（2027 起）", "m_soc=@lo", None, "測試對 WFE 比回到公司說的 7–9% 區間（市場乘數 0.85；三情境）；只有低案"),
    ("soc_sh", "泰瑞達 SoC 市占 ±5 個百分點", "s_soc=@lo", "s_soc=@hi", "A、B 情境 FY2027 起市占 −5／＋5 個百分點（C 情境與 Advantest 互補，不變）"),
    ("mem_sh", "泰瑞達記憶體市占 ±5 個百分點", "s_mem=@lo", "s_mem=@hi", "三情境 FY2027 起"),
    ("inc", "測試事業增量利益率 40%／56%", "t_inc=@lo", "t_inc=@hi", "固定成本重新推導（FY2026 錨點不變；基準由兩點推得）"),
    ("pt", "Product Test 成長 低／高", _G("@lo"), _G("@hi"), "A 4%／12%、B 1%／9%、C −2%／6%（基準 8%／5%／2%）"),
    ("rob", "機器人成長 0%／20%", "g_rob=@lo", "g_rob=@hi", "三情境相同（基準 10%）"),
    ("rob_m", "機器人利益率 低／高", _RM("@lo"), _RM("@hi"), "FY2027–29 稅前利益率低 −5%／−3%／−2%、高 7%／9%／10%（基準 2%／4%／5%）"),
    ("etr", "稅率 12%／17%", "etr=@lo", "etr=@hi", "有效稅率（基準由前三季實績推得）"),
    ("bb", "回購 0%／2%／年", "bb=@lo", "bb=@hi", "稀釋股數每年減少（基準 1%）"),
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


jobs0 = {"base": {"TER_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"TER_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"TER_OVERRIDE": hi}
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

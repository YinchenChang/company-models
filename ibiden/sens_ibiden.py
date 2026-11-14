# Ibiden 敏感度快照（v0.5.5 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖）：每個案例以環境變數 IB_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 ibiden/sens.json；正式建置（build_ibiden.py → v055_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 ibiden/sens_ibiden.py [暫存資料夾，預設 /tmp/work_ibiden/sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_ibiden）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_ibiden/sens"
BUILD = os.path.join(HERE, "build_ibiden.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("IB_OVERRIDE", "IB_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄
_PP = lambda s: ";".join(f"pp_{y}={s}" for y in (27, 28, 29))
_RL = lambda s: ";".join(f"rel_{y}={s}" for y in (27, 28, 29))
CASES = [
    ("premium", "短缺溢價路徑（快速回落／延續）", _PP("@lo"), _PP("@hi"), "價格溢價指數 2027–2029：低 0.95／0.85／0.75、高 1.05／1.05／1.00"),
    ("share", "Ibiden 份額相對值（流失／持平）", _RL("@lo"), _RL("@hi"), "份額相對值 2027–2029：低 0.85／0.75／0.70、高 1.00"),
    ("asp", "每顆載板金額年增（10%／22%）", "g_asp=@lo", "g_asp=@hi", "面積、層數成長較慢／較快（基準 18%）"),
    ("w_ot", "非 NVIDIA 加速器權重（0.2／0.6）", "w_ot=@lo", "w_ot=@hi", "Ibiden 在 ASIC 的份額 × 金額相對 NVIDIA（基準 0.4）"),
    ("ai_e", "AI 占電子事業比率（55%／75%）", "ai_e=@lo", "ai_e=@hi", "FY2026 切分（公司未揭露；基準 65%；校準跟著變）"),
    ("m_ne", "電子事業非 AI 營益率（10%／20%）", "m_ne=@lo", "m_ne=@hi", "非 AI 部分營益率（基準 15%；低值使 AI 載板校準營益率較高）"),
    ("g_ne", "電子事業非 AI 營收年增（0%／10%）", "g_ne=@lo", "g_ne=@hi", "通用伺服器、PC 載板等（基準 5%）"),
    ("g_cer", "陶瓷事業營收年增（−5%／＋3%）", "g_cer=@lo", "g_cer=@hi", "基準 0%"),
    ("m_cer", "陶瓷事業營益率（5%／10%）", "m_cer=@lo", "m_cer=@hi", "基準 7.7%"),
    ("g_oth", "其他事業營收年增（0%／5%）", "g_oth=@lo", "g_oth=@hi", "基準 3%"),
    ("m_oth", "其他事業營益率（9%／13%）", "m_oth=@lo", "m_oth=@hi", "基準 11.5%"),
    ("ni_oi", "淨利 ÷ 營業利益（0.62／0.70）", "ni_oi=@lo", "ni_oi=@hi", "稅、非營業與少數股權（基準 0.661）"),
    ("sh_g", "CB 轉換稀釋（股數每年 +3%）", None, "sh_g=@hi", "2031 年到期 CB 轉換（基準 0%）；只有高案"),
    ("wts", "情境權重（全給保守 C／全給樂觀 A）", "tp_wA=0;tp_wB=0;tp_wC=1", "tp_wA=1;tp_wB=0;tp_wC=0", "基準權重 A 0.3／B 0.5／C 0.2；顆數情境 C＝CY2028–29 ×0.6、A＝原始顆數"),
]


def lab(ws, prefix, start=1):
    return next(r for r in range(start, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix))


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


jobs0 = {"base": {"IB_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"IB_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"IB_OVERRIDE": hi}
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

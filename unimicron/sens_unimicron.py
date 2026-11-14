# 欣興敏感度快照（v0.4.4 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖；樣板 ibiden/sens_ibiden.py）：每個案例以環境變數 UN_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 unimicron/sens.json；正式建置（build_unimicron.py → v044_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 unimicron/sens_unimicron.py [暫存資料夾，預設 /tmp/work_unimicron/sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_unimicron）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_unimicron/sens"
BUILD = os.path.join(HERE, "build_unimicron.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("UN_OVERRIDE", "UN_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄
_M = lambda pre, s, ys=(27, 28, 29): ";".join(f"{pre}{y}={s}" for y in ys)
CASES = [
    ("premium", "ABF 價格溢價路徑（快速回落／延續）", _M("ibpp_", "@lo"), _M("ibpp_", "@hi"), "價格溢價指數 2027–2029（讀 Ibiden 模型）：低 0.95／0.85／0.75、高 1.05／1.05／1.00"),
    ("share", "欣興加速器載板份額（流失／提升）", _M("sh_nv_", "@lo") + ";" + _M("sh_ot_", "@lo"), _M("sh_nv_", "@hi") + ";" + _M("sh_ot_", "@hi"), "NVIDIA 與非 NVIDIA 加速器載板份額 2027–2029 同取低值或高值"),
    ("asp", "每顆載板金額年增（10%／22%）", "g_asp=@lo", "g_asp=@hi", "面積、層數成長較慢／較快（基準 18%）"),
    ("cap", "ABF 產能擴充（較慢／較快）", _M("cap_", "@lo"), _M("cap_", "@hi"), "ABF 產能指數 2027–2029：低 1.25／1.50／1.70、高 1.45／1.90／2.30"),
    ("carry", "欣興漲價年度化（0%／10%）", "carry=@lo", "carry=@hi", "CY2026 只反映約半年漲價（基準 5%）"),
    ("ft_abf", "ABF 增量現金毛利率（40%／60%）", "ft_abf=@lo", "ft_abf=@hi", "基準 50%"),
    ("ft_oth", "其他產品增量現金毛利率（15%／35%）", "ft_oth=@lo", "ft_oth=@hi", "BT、HDI、PCB、軟板（基準 25%）"),
    ("el_o", "其他 AI 載板對 AI 資本支出的彈性（0.3／0.8）", "el_o=@lo", "el_o=@hi", "交換器、伺服器 CPU、網通（基準 0.5）"),
    ("g_pai", "AI HDI／PCB 年增（各年 −10／＋10 個百分點）", _M("g_pai_", "@lo"), _M("g_pai_", "@hi"), "基準 2027 20%、2028 15%、2029 10%"),
    ("ai_abf", "FY2026 AI 占 ABF 營收（65%／85%）", "ai_abf=@lo", "ai_abf=@hi", "起點切分（基準 75%；其他 AI 載板為校準殘差）"),
    ("life", "新增資產折舊年限（6／10 年）", "life=@lo", "life=@hi", "基準 8 年"),
    ("phi", "營業費用對營收的彈性（0.3／1.0）", "phi=@lo", "phi=@hi", "基準 0.6"),
    ("nonop", "營業外淨額 FY2027–2029（0／40 億元）", _M("nonop", "@lo"), _M("nonop", "@hi"), "基準每年 20 億元"),
    ("tax", "所得稅與少數股權占稅前利益（15%／22%）", "tax=@lo", "tax=@hi", "基準 18%"),
    ("q4", "4Q26 營收季增與毛利率（0%、28%／10%、31%）", "q4_g=@lo;q4_gm=@lo", "q4_g=@hi;q4_gm=@hi", "FY2026 錨點的第四季假設（基準季增 5%、毛利率 29.5%）"),
    ("wts", "情境權重（全給保守 C／全給樂觀 A）", "tp_wA=0;tp_wB=0;tp_wC=1", "tp_wA=1;tp_wB=0;tp_wC=0", "基準權重 A 0.3／B 0.5／C 0.2"),
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


jobs0 = {"base": {"UN_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"UN_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"UN_OVERRIDE": hi}
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

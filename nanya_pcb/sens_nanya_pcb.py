# 南電敏感度快照（v0.3.5 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖；樣板 ibiden/sens_ibiden.py）：每個案例以環境變數 NP_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 nanya_pcb/sens.json；正式建置（build_nanya_pcb.py → v044_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 nanya_pcb/sens_nanya_pcb.py [暫存資料夾，預設 /tmp/work_nanya_pcb/sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_nanya_pcb）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_nanya_pcb/sens"
BUILD = os.path.join(HERE, "build_nanya_pcb.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("NP_OVERRIDE", "NP_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄
_M = lambda pre, s, ys=(27, 28, 29): ";".join(f"{pre}{y}={s}" for y in ys)
CASES = [
    ("premium", "ABF 價格溢價路徑（快速回落／延續）", _M("ibpp_", "@lo"), _M("ibpp_", "@hi"), "價格溢價指數 2027–2029（讀 Ibiden 模型）：低 0.95／0.85／0.75、高 1.05／1.05／1.00"),
    ("share", "南電加速器載板份額（流失／提升）", _M("sh_nv_", "@lo") + ";" + _M("sh_ot_", "@lo"), _M("sh_nv_", "@hi") + ";" + _M("sh_ot_", "@hi"), "NVIDIA 與非 NVIDIA 加速器載板份額 2027–2029 同取低值或高值"),
    ("asp", "每顆載板金額年增（10%／22%）", "g_asp=@lo", "g_asp=@hi", "面積、層數成長較慢／較快（基準 18%）"),
    ("cap", "ABF 產能擴充（較慢／較快）", _M("cap_", "@lo"), _M("cap_", "@hi"), "ABF 產能指數 2027–2029：低 1.10／1.20／1.35、高 1.30／1.50／1.85"),
    ("carry", "南電漲價年度化（8%／22%）", "carry=@lo", "carry=@hi", "CY2026 基期含上半年低價（基準 15%）"),
    ("ft_abf", "ABF 增量現金毛利率（40%／60%）", "ft_abf=@lo", "ft_abf=@hi", "基準 50%"),
    ("ft_oth", "BT、一般電路板增量現金毛利率（25%／55%）", "ft_oth=@lo", "ft_oth=@hi", "基準 40%"),
    ("el_o", "交換器等其他 AI 載板對交換器市場的彈性（0.3／0.8）", "el_o=@lo", "el_o=@hi", "基準 0.5"),
    ("ai_abf", "FY2026 AI 與資料中心占 ABF 營收（55%／75%）", "ai_abf=@lo", "ai_abf=@hi", "起點切分（基準 65%；其他 AI 載板為校準殘差）"),
    ("cx", "資本支出 FY2027–2029（低／高）", _M("cx", "@lo"), _M("cx", "@hi"), "低 60／60／60 億、高 200／240／200 億（基準 140／180／120 億）"),
    ("life", "新增資產折舊年限（6／10 年）", "life=@lo", "life=@hi", "基準 8 年"),
    ("phi", "營業費用對營收的彈性（0.3／1.0）", "phi=@lo", "phi=@hi", "基準 0.6"),
    ("nonop", "營業外淨額 FY2027–2029（0／12 億元）", _M("nonop", "@lo"), _M("nonop", "@hi"), "基準每年 6 億元"),
    ("tax", "所得稅占稅前利益（18%／24%）", "tax=@lo", "tax=@hi", "基準為最近四季實績 22.7%"),
    ("q4", "4Q26 營收季增與毛利率（3%、29%／15%、34%）", "q4_g=@lo;q4_gm=@lo", "q4_g=@hi;q4_gm=@hi", "FY2026 錨點的第四季假設（基準季增 10%、毛利率 32.5%）"),
    ("g_pcb", "一般電路板年增（−5%／＋3%）", "g_pcb=@lo", "g_pcb=@hi", "基準 0%"),
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


jobs0 = {"base": {"NP_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"NP_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"NP_OVERRIDE": hi}
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

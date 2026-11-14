# 景碩敏感度快照（v0.4.4 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖；樣板 ibiden/sens_ibiden.py）：每個案例以環境變數 KS_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 kinsus/sens.json；正式建置（build_kinsus.py → v044_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 kinsus/sens_kinsus.py [暫存資料夾，預設 /tmp/work_kinsus/sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_kinsus）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_kinsus/sens"
BUILD = os.path.join(HERE, "build_kinsus.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("KS_OVERRIDE", "KS_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄
_M = lambda pre, s, ys=(27, 28, 29): ";".join(f"{pre}{y}={s}" for y in ys)
CASES = [
    ("premium", "ABF 價格溢價路徑（快速回落／延續）", _M("ibpp_", "@lo"), _M("ibpp_", "@hi"), "價格溢價指數 2027–2029（讀 Ibiden 模型）：低 0.95／0.85／0.75、高 1.05／1.05／1.00"),
    ("share", "景碩 GPU 與 ASIC 加速器載板份額（流失／提升）", _M("sh_nv_", "@lo") + ";" + _M("sh_ot_", "@lo"), _M("sh_nv_", "@hi") + ";" + _M("sh_ot_", "@hi"), "NVIDIA 與非 NVIDIA 加速器載板份額 2027–2029 同取低值或高值"),
    ("cpu", "景碩 NVIDIA CPU 載板份額（流失／提升）", _M("sh_cpu_", "@lo"), _M("sh_cpu_", "@hi"), "Grace／Vera CPU 載板份額 2027–2029：低 50%／50%／45%、高 90%／100%／100%"),
    ("asp", "每顆載板金額年增（10%／22%）", "g_asp=@lo", "g_asp=@hi", "面積、層數成長較慢／較快（基準 18%）"),
    ("cap", "ABF 產能擴充（較慢／較快）", _M("cap_", "@lo"), _M("cap_", "@hi"), "ABF 產能指數 2027–2029：低 1.10／1.30／1.55、高 1.25／1.62／2.00"),
    ("carry", "ABF 漲價年度化（8%／22%）", "carry=@lo", "carry=@hi", "CY2026 基期含上半年低價（基準 15%）"),
    ("ft_abf", "ABF 增量現金毛利率（40%／60%）", "ft_abf=@lo", "ft_abf=@hi", "基準 50%"),
    ("ft_bt", "BT 與其他載板增量現金毛利率（30%／60%）", "ft_bt=@lo", "ft_bt=@hi", "基準 45%"),
    ("el_o", "其他 AI 載板對交換器市場的彈性（0.3／0.8）", "el_o=@lo", "el_o=@hi", "網通 ASIC、交換器、FPGA（基準 0.5）"),
    ("g_nai", "非 AI ABF 年增（2%／10%）", "g_nai=@lo", "g_nai=@hi", "伺服器與 PC CPU、基地台 FPGA（基準 6%）"),
    ("g_lens", "晶碩營收年增（3%／15%）", "g_lens=@lo", "g_lens=@hi", "隱形眼鏡（基準 10%）"),
    ("cx", "資本支出 FY2027–2029（低／高）", _M("cx", "@lo"), _M("cx", "@hi"), "低 100／90／70 億、高 220／220／160 億（基準 170／160／120 億）"),
    ("phi", "營業費用對營收的彈性（0.3／1.0）", "phi=@lo", "phi=@hi", "基準 0.6"),
    ("nonop", "營業外淨額 FY2027–2029（0／8 億元）", _M("nonop", "@lo"), _M("nonop", "@hi"), "基準每年 4 億元"),
    ("tax", "所得稅占稅前利益（8%／20%）", "tax=@lo", "tax=@hi", "基準 12%"),
    ("q4", "4Q26 營收季增與毛利率（8%、29.3%／19%、33%）", "q4_g=@lo;q4_gm=@lo", "q4_g=@hi;q4_gm=@hi", "FY2026 錨點的第四季假設（基準季增 12%、毛利率 32%）"),
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


jobs0 = {"base": {"KS_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"KS_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"KS_OVERRIDE": hi}
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

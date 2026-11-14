# 台積電 敏感度快照（v0.5.5 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖）：每個案例以環境變數 TSMC_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 tsmc/sens.json；正式建置（build_tsmc.py → v055_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 tsmc/sens_tsmc.py [暫存資料夾，預設 /tmp/work_tsmc/sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_tsmc）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_tsmc/sens"
BUILD = os.path.join(HERE, "build_tsmc.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("TSMC_OVERRIDE", "SIL_COORD_ALT", "COORD_RUN")}
env0["COORD_RUN"] = "1"   # 不加 Coord 頁與 v055_post（只讀 Val 頁）

SC3 = "ABC"
_ALL = lambda f: ";".join(f(k) for k in ("gm27_A gm27_B gm27_C gm28_A gm28_B gm28_C gm29_A gm29_B gm29_C".split()))
_GM = lambda d: _ALL(lambda k: f"{k}=d{d}")
_NA = lambda d: ";".join(f"na{y}_{s}=d{d}" for y in (27, 28, 29) for s in SC3)
_CT = lambda sd: ";".join(f"ct_{k}={sd}" for k in ("GB", "VR", "RU", "HP", "AMD8", "AMDR", "TPU7", "TPU8", "TRN", "ASIC"))
_PX = lambda v: ";".join(f"px{y}={v}" for y in (27, 28, 29))
_CX = lambda m: ";".join(f"capex{y}=x{m}" for y in (27, 28, 29))
# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；低案＝輸入取低值（不一定使目標價較低）
CASES = [
    ("units", "AI 加速器顆數 ×0.8／×1.2（2027 起，三情境）", "_units_x=0.8", "_units_x=1.2", "2027 年起各平台顆數同乘 0.8／1.2（2029 由 2028 成長，隨之縮放；2026 為錨點不動）"),
    ("content", "每顆台積電內容 −20%／+20%", _CT("@lo"), _CT("@hi"), "In 頁各平台每顆台積電內容取低／高欄（基準 ±20%）"),
    ("price", "AI 內容價格調升：不漲／基準的 1.5 倍", _PX("0"), _PX("x1.5"), "2027–2029 價格調升 0%／基準 ×1.5（基準 12%／8%／5%）"),
    ("ai_oth", "其他 AI 相關營收：0%／40%", "ai_oth=@lo", "ai_oth=@hi", "網通 ASIC、HBM 基底晶粒、CPU 等加成（基準 20%）"),
    ("gm", "毛利率 −3／+3 個百分點（2027–2029，三情境）", _GM("-0.03"), _GM("+0.03"), "FY2027–FY2029 各情境毛利率同加減；FY2026 為前三季實績與 4Q 指引錨點，不動"),
    ("na", "非 AI 營收成長 −5／+5 個百分點（2027–2029，三情境）", _NA("-0.05"), _NA("+0.05"), "非 AI 基準成長（In 頁各情境各年）同加減；平台差距不變"),
    ("opex", "營業費用率 −2／+2 個百分點（2027 起）", "opex=d-0.02", "opex=d+0.02", "基準 7.5%（2026 用實績推得值）"),
    ("ni_op", "淨利 ÷ 營業利益 0.864／0.909（2027 起）", "ni_op=@lo", "ni_op=@hi", "低：2026 年扣除一次性利益後的內含水準；高：2023 年年度值（NI_OI 頁）；基準 0.875"),
    ("fx", "新台幣兌美元 30／33（2027 起）", "fx=@lo", "fx=@hi", "基準 32；2026 年用實績與指引匯率，不動"),
    ("capex", "資本支出 −10%／+10%（2027–2029）", _CX("0.9"), _CX("1.1"), "只影響現金流折現（DCF）；同業比較法不變"),
    ("wts", "情境權重（全給保守 C／全給樂觀 A）", "tp_wA=0;tp_wB=0;tp_wC=1", "tp_wA=1;tp_wB=0;tp_wC=0", "基準權重 A 0.3／B 0.5／C 0.2"),
]


def lab(ws, prefix):
    return next(r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip() == prefix)


def run(jobs, sub):
    raw, rc = os.path.join(TMP, sub, "raw"), os.path.join(TMP, sub, "rc")
    shutil.rmtree(os.path.join(TMP, sub), ignore_errors=True); os.makedirs(raw); os.makedirs(rc)
    for nm, env in jobs.items():
        subprocess.run([sys.executable, BUILD, os.path.join(raw, f"{nm}.xlsx")], check=True, env={**env0, **env}, stdout=subprocess.DEVNULL)
    subprocess.run(["soffice"] + ([f"-env:UserInstallation={os.environ['LO_PROFILE']}"] if os.environ.get("LO_PROFILE") else [])
                   + ["--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", rc] + sorted(glob.glob(os.path.join(raw, "*.xlsx"))),
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return rc


ERR = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "Err:502")


def read(rc, nm):
    wb = openpyxl.load_workbook(os.path.join(rc, f"{nm}.xlsx"), data_only=True)
    V = wb["Val"]
    errs = sum(1 for ws in wb for row in ws.iter_rows() for x in row if isinstance(x.value, str) and x.value in ERR)
    assert errs == 0, (nm, errs)
    return dict(tp=V.cell(lab(V, "目標價（DCF 與 Comparable 加權）"), 2).value, dcf=V.cell(lab(V, "DCF 目標價（情境加權）"), 2).value)


rc0 = run({"base": {}}, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"TSMC_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"TSMC_OVERRIDE": hi}
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

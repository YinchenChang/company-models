# Amkor 敏感度快照（v0.4.4）：每個營運案例以環境變數 AMKR_P 覆寫三情境主引擎（PL 頁）的案例參數（覆寫鍵同 build_amkr.py Sens 頁「一次動一個輸入」案例），
# 整本重建、LibreOffice 重算，讀加權目標價、DCF 目標價與 Comparable（相對倍數調整後），寫入 amkr/sens.json；正式建置（build_amkr.py → v044_post.py）讀入 Sens 頁「加權目標價相對基準」表。
# 與 Sens 頁既有案例的差別：既有案例只動情境 A 並看 FY2027／FY2028 EPS；本快照三情境同時套用，並重算 DCF、Comparable 與加權，得到加權目標價。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 amkr/sens_amkr.py [暫存資料夾，預設 /tmp/work_amkr/sens]；並行時設 LO_PROFILE＝file:///tmp/lo_amkr
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_amkr/sens"
BUILD = os.path.join(HERE, "build_amkr.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("AMKR_P", "AMKR_SENS", "SIL_COORD_ALT")}
env0["COORD_RUN"] = "1"; env0["AMKR_SENS"] = "1"; env0["PYTHONDONTWRITEBYTECODE"] = "1"   # 不加 Coord 頁、不跑 v044_post
# (鍵, 白話標籤, 低案覆寫, 高案覆寫)；None＝該案只有單向
CASES = [
    ("mf", "AI 部署顆數 ×0.8／×1.2（FY2027 起）", dict(mf="0.8"), dict(mf="1.2")),
    ("out", "AI 後段委外比例 低／高（2027 年 45%／65%，2028–29 年 45%／70–75%）", dict(out="out_{y}_lo"), dict(out="out_{y}_hi")),
    ("sh", "Amkor 占委外後段份額 低／高（28%；2027 年 40%、2028–29 年 45%）", dict(sh="sh_{y}_lo"), dict(sh="sh_{y}_hi")),
    ("el", "每顆封測金額對封裝面積的彈性 0.3／1.0", dict(el="el_lo"), dict(el="el_hi")),
    ("tri", "FY2026 AI 先進封裝營收為 FY2025 的 2.5 倍／3.2 倍（基準 3 倍）", dict(tri="tri_lo"), dict(tri="tri_hi")),
    ("cm", "通訊走保守情境路徑（FY2027 −6%、之後 0%；只有低案）", dict(gcm27="g_cm27_C", gcmL="g_cmL_C"), None),
    ("au", "車用與工業年增 +4%／+14%（基準 +10%）", dict(gau="g_au_lo"), dict(gau="g_au_hi")),
    ("ft", "現金流通率 30%／50%（基準 40%）", dict(ft="ft_lo"), dict(ft="ft_hi")),
    ("capi", "本業資本支出占營收 10%／18%（基準 14%）", dict(capi="capi_lo"), dict(capi="capi_hi")),
    ("life", "本業折舊年限 5 年／9 年（基準 7 年）", dict(life="life_lo"), dict(life="life_hi")),
    ("azfix", "亞利桑那固定現金成本 低／高（2027 年 80／170，2028–29 年 70／150）", dict(azfix="az_fix{y}_lo"), dict(azfix="az_fix{y}_hi")),
    ("azinc", "亞利桑那補助與抵減占投資 20%／40%（基準 35%）", dict(azinc="az_inc_lo"), dict(azinc="az_inc_hi")),
    ("azcm", "亞利桑那邊際貢獻率 45%／60%（基準 55%）", dict(azcm="az_cm_lo"), dict(azcm="az_cm_hi")),
    ("phi", "營業費用對營收成長的彈性 0.3／1.0（基準 0.5）", dict(phi="phi_lo"), dict(phi="phi_hi")),
    ("tax", "有效稅率 15%／25%（基準 20%）", dict(tax="tax_lo"), dict(tax="tax_hi")),
    ("nonop", "營業外淨額 低／高（2027 年 −60／＋20 百萬美元，2028–29 年再各 −20／＋40）", dict(nonop="nonop{y}_lo"), dict(nonop="nonop{y}_hi")),
]


def lab(ws, prefix):
    return next(r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix))


def run(jobs, sub):
    raw, rc = os.path.join(TMP, sub, "raw"), os.path.join(TMP, sub, "rc")
    shutil.rmtree(os.path.join(TMP, sub), ignore_errors=True); os.makedirs(raw); os.makedirs(rc)
    for nm, ov in jobs.items():
        env = dict(env0)
        if ov: env["AMKR_P"] = json.dumps(ov)
        subprocess.run([sys.executable, BUILD, os.path.join(raw, f"{nm}.xlsx")], check=True, env=env, stdout=subprocess.DEVNULL)
    subprocess.run(["soffice"] + ([f"-env:UserInstallation={os.environ['LO_PROFILE']}"] if os.environ.get("LO_PROFILE") else [])
                   + ["--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", rc] + sorted(glob.glob(os.path.join(raw, "*.xlsx"))),
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return rc


def read(rc, nm):
    wb = openpyxl.load_workbook(os.path.join(rc, f"{nm}.xlsx"), data_only=True)
    V = wb["Val"]
    # 容許 Val「歷史倍數檢查」列的 PERCENTRANK 超出歷史範圍（只作檢查、不進目標價）
    errs = sum(1 for ws in wb for row in ws.iter_rows() for x in row if isinstance(x.value, str) and x.value.startswith(("#", "Err:"))
               and not (ws.title == "Val" and str(ws.cell(x.row, 1).value or "").startswith("歷史倍數檢查")))
    assert errs == 0, (nm, errs)
    return dict(tp=V.cell(lab(V, "目標價（DCF 與 Comparable 加權）"), 2).value, dcf=V.cell(lab(V, "DCF 目標價（情境加權）"), 2).value,
                comp=V.cell(lab(V, "Comparable 目標價（相對倍數調整後）"), 2).value)


if __name__ == "__main__":
    jobs = {"base": {}}
    for k, _, lo, hi in CASES:
        if lo: jobs[f"{k}__lo"] = lo
        if hi: jobs[f"{k}__hi"] = hi
    rc = run(jobs, "cases")
    R = {nm: read(rc, nm) for nm in jobs}
    b = R["base"]
    out = {"_說明": "加權目標價敏感度整本重算快照（amkr/sens_amkr.py）；正式建置讀入 Sens 頁「加權目標價相對基準」表", "_base": b, "biz": {}}
    for k, label, lo, hi in CASES:
        out["biz"][k] = {"label": label, "低": R.get(f"{k}__lo"), "高": R.get(f"{k}__hi")}
        print(f"{k:6s} 低 {(R[f'{k}__lo']['tp'] / b['tp'] - 1) if lo else float('nan'):+.2%}  高 {(R[f'{k}__hi']['tp'] / b['tp'] - 1) if hi else float('nan'):+.2%}")
    print("基準", b)
    json.dump(out, open(os.path.join(HERE, "sens.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

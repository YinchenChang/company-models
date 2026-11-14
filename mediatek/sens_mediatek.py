# 聯發科敏感度快照（v0.3.4）：每個營運案例以環境變數 MTK_P 覆寫三情境主引擎（PL 頁）的案例參數（覆寫鍵同 build_mediatek.py 引擎參數與 Sens 頁「一次動一個輸入」案例），
# 整本重建、LibreOffice 重算，讀加權目標價、DCF 目標價與 Comparable（相對倍數調整後），寫入 mediatek/sens.json；正式建置（build_mediatek.py → v034_post.py）讀入 Sens 頁「加權目標價相對基準」表。
# 與 Sens 頁既有案例的差別：既有案例只動情境 A 並看 FY2027／FY2028 EPS；本快照三情境同時套用，並重算 DCF、Comparable 與加權，得到加權目標價。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 mediatek/sens_mediatek.py [暫存資料夾，預設 /tmp/work_mediatek/sens]；並行時設 LO_PROFILE＝file:///tmp/lo_mediatek
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_mediatek/sens"
BUILD = os.path.join(HERE, "build_mediatek.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("MTK_P", "MTK_SENS", "SIL_COORD_ALT")}
env0["COORD_RUN"] = "1"; env0["MTK_SENS"] = "1"; env0["PYTHONDONTWRITEBYTECODE"] = "1"   # 不加 Coord 頁、不跑 v034_post
# (鍵, 白話標籤, 低案覆寫, 高案覆寫)；None＝該案只有單向
CASES = [
    ("uf", "聯發科 ASIC 顆數 ×0.8／×1.2（FY2027 起）", dict(uf="uf_lo"), dict(uf="uf_hi")),
    ("sh", "聯發科占 TPU 份額 −5 個百分點／取上限 20%（CY2027 起）", dict(shp="(-0.05)"), dict(shs="_hi")),
    ("pfc", "第二顆 ASIC 核心售價 ×0.8／×2（CY2028 起；基準 ×1）", dict(pfc="pfc_lo"), dict(pfc="pfc_hi")),
    ("oth", "其他 ASIC 份額 +5 個百分點（第二客戶；基準 0；只有高案）", None, dict(sho="0.05")),
    ("fh", "HBM 由聯發科代購的比例 30%／100%（基準 60%）", dict(fh="f_hbm_lo"), dict(fh="f_hbm_hi")),
    ("mc", "ASIC 核心毛利率 35%／55%（基準 45%）", dict(mc="m_core_lo"), dict(mc="m_core_hi")),
    ("mob", "手機走保守情境路徑（FY2027 −8%、之後 −3%；只有低案）", dict(gmob27="g_mob27_C", gmobL="g_mobL_C"), None),
    ("sep", "智慧邊緣平台走保守情境路徑（FY2027 −3%、之後 +1%；只有低案）", dict(gsep27="g_sep27_C", gsepL="g_sepL_C"), None),
    ("gma", "非 ASIC 毛利率 −2／＋2 個百分點", dict(gma="gmadj_lo"), dict(gma="gmadj_hi")),
    ("phi", "營業費用對營收成長的彈性 0.2／0.6（基準 0.3）", dict(phi="phi_lo"), dict(phi="phi_hi")),
    ("opxa", "營業費用額外占營收 −2／＋2 個百分點", dict(opxa="opx_adj_lo"), dict(opxa="opx_adj_hi")),
    ("fx", "匯率 30／34（ASIC 美元營收換算；基準 32）", dict(fx="fx_lo"), dict(fx="fx_hi")),
    ("tax", "有效稅率 12%／17%（基準 14%）", dict(tax="tax_lo"), dict(tax="tax_hi")),
    ("nonop", "營業外淨額每年 10／24 十億新台幣（基準 18）", dict(nonop="nonop_f_lo"), dict(nonop="nonop_f_hi")),
]


def lab(ws, prefix):
    return next(r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix))


def run(jobs, sub):
    raw, rc = os.path.join(TMP, sub, "raw"), os.path.join(TMP, sub, "rc")
    shutil.rmtree(os.path.join(TMP, sub), ignore_errors=True); os.makedirs(raw); os.makedirs(rc)
    for nm, ov in jobs.items():
        env = dict(env0)
        if ov: env["MTK_P"] = json.dumps(ov)
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
    out = {"_說明": "加權目標價敏感度整本重算快照（mediatek/sens_mediatek.py）；正式建置讀入 Sens 頁「加權目標價相對基準」表", "_base": b, "biz": {}}
    for k, label, lo, hi in CASES:
        out["biz"][k] = {"label": label, "低": R.get(f"{k}__lo"), "高": R.get(f"{k}__hi")}
        print(f"{k:6s} 低 {(R[f'{k}__lo']['tp'] / b['tp'] - 1) if lo else float('nan'):+.2%}  高 {(R[f'{k}__hi']['tp'] / b['tp'] - 1) if hi else float('nan'):+.2%}")
    print("基準", b)
    json.dump(out, open(os.path.join(HERE, "sens.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

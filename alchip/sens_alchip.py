# 世芯-KY 敏感度快照（v0.2.4）：每個營運案例以環境變數 ALCHIP_P 覆寫三情境主引擎（PL 頁）的案例參數（覆寫鍵同 build_alchip.py Sens 頁「一次動一個輸入」案例），
# 整本重建、LibreOffice 重算，讀加權目標價、DCF 目標價與 Comparable（相對倍數調整後），寫入 alchip/sens.json；正式建置（build_alchip.py → v024_post.py）讀入 Sens 頁「加權目標價相對基準」表。
# 與 Sens 頁既有案例的差別：既有案例只動情境 A 並看 FY2027／FY2028 EPS；本快照三情境同時套用，並重算 DCF、Comparable 與加權，得到加權目標價。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 alchip/sens_alchip.py [暫存資料夾，預設 /tmp/work_alchip/sens]；並行時設 LO_PROFILE＝file:///tmp/lo_designsvc
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_alchip/sens"
BUILD = os.path.join(HERE, "build_alchip.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("ALCHIP_P", "ALCHIP_SENS", "SIL_COORD_ALT")}
env0["COORD_RUN"] = "1"; env0["ALCHIP_SENS"] = "1"; env0["PYTHONDONTWRITEBYTECODE"] = "1"   # 不加 Coord 頁、不跑 v044_post
# (鍵, 白話標籤, 低案覆寫, 高案覆寫)；None＝該案只有單向；覆寫鍵同 build_alchip.py engine() 的 DEFP（三情境同時套用）
CASES = [
    ("shp", "Trainium 份額 −10／＋10 個百分點（CY2027 起）", dict(shp="(-0.10)"), dict(shp="0.10")),
    ("shs", "Trainium 份額取低／高值（CY2027 起；低 40%、30%，高 75%、85%，高值受 1 − Marvell 上限）", dict(shs="_lo"), dict(shs="_hi")),
    ("uf", "世芯 Trainium 顆數 ×0.8／×1.2（CY2027 起）", dict(uf="uf_lo"), dict(uf="uf_hi")),
    ("pfc", "2 奈米每顆營收係數 ×0.8／×2.5（CY2028 起；基準 B 1.2）", dict(pfc="pfc_A_lo"), dict(pfc="pfc_A_hi")),
    ("fh", "代購 HBM 並計入營收的比例 0.6（基準 0；只有高案）", None, dict(fh="f_al_hi")),
    ("mc", "Trainium 量產毛利率 10%／16%（基準 13%）", dict(mc="m_trn_lo"), dict(mc="m_trn_hi")),
    ("oth", "其他營收走保守（C）／樂觀（A）情境路徑（FY2027 起）", dict(go27="g_o27_C", goL="g_oL_C"), dict(go27="g_o27_A", goL="g_oL_A")),
    ("gma", "其他營收毛利率 −3／＋3 個百分點（FY2027 起）", dict(gma="gmadj_lo"), dict(gma="gmadj_hi")),
    ("phi", "營業費用彈性 0.2／0.6（基準 0.3）", dict(phi="phi_lo"), dict(phi="phi_hi")),
    ("opxa", "營業費用率 ＋2 個百分點（只有高案）", None, dict(opxa="opx_adj_hi")),
    ("fx", "匯率 30／34（基準 32；Trainium 美元營收換算）", dict(fx="fx_lo"), dict(fx="fx_hi")),
    ("tax", "有效稅率 17%／25%（基準 20%）", dict(tax="tax_lo"), dict(tax="tax_hi")),
    ("nonop", "營業外淨額 1.0／2.2（基準 1.6）", dict(nonop="nonop_f_lo"), dict(nonop="nonop_f_hi")),
    ("fcwc", "營運資金強度改 9 月底實績（營運資金 ÷ 3Q26 年化營收，約 21%；基準 4.5%；只有低案）", dict(fcwc="1"), None),
    ("stubwc", "4Q26 營運資金依 FY2026 全年假設反推（釋出 22.28 十億元；基準 4Q 變動 0；只有高案）", None, dict(stubwc="1")),
]


def lab(ws, prefix):
    return next(r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix))


def run(jobs, sub):
    raw, rc = os.path.join(TMP, sub, "raw"), os.path.join(TMP, sub, "rc")
    shutil.rmtree(os.path.join(TMP, sub), ignore_errors=True); os.makedirs(raw); os.makedirs(rc)
    for nm, ov in jobs.items():
        env = dict(env0)
        if ov: env["ALCHIP_P"] = json.dumps(ov)
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
    out = {"_說明": "加權目標價敏感度整本重算快照（alchip/sens_alchip.py）；正式建置讀入 Sens 頁「加權目標價相對基準」表", "_base": b, "biz": {}}
    for k, label, lo, hi in CASES:
        out["biz"][k] = {"label": label, "低": R.get(f"{k}__lo"), "高": R.get(f"{k}__hi")}
        print(f"{k:6s} 低 {(R[f'{k}__lo']['tp'] / b['tp'] - 1) if lo else float('nan'):+.2%}  高 {(R[f'{k}__hi']['tp'] / b['tp'] - 1) if hi else float('nan'):+.2%}")
    print("基準", b)
    json.dump(out, open(os.path.join(HERE, "sens.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

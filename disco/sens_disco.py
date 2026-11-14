# Disco 敏感度快照（v0.3.5 起；版面規格 v1.3 H 節「加權目標價相對基準」主圖）：每個案例以環境變數 DISCO_OVERRIDE 覆寫 In 頁輸入後整本重建、
# LibreOffice 重算，讀加權目標價與 DCF 目標價，寫入 disco/sens.json；正式建置（build_disco.py → v035_post.py）讀入 Sens 頁末「加權目標價相對基準」表。
# In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁快照狀態顯示 STALE）。
# 用法：python3 disco/sens_disco.py [暫存資料夾，預設 /tmp/disco_w/sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_disco）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/disco_w/sens"
BUILD = os.path.join(HERE, "build_disco.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("DISCO_OVERRIDE", "DISCO_SNAP", "SIL_COORD_ALT", "COORD_RUN")}

# 案例：(代號, 白話標籤, 低案覆寫, 高案覆寫, 設定說明)；覆寫值 @lo／@hi＝In 頁敏感度低／高欄（三情境同時受影響）
CASES = [
    ("ai", "生成 AI 出貨錨點 FY2025 −10%／＋10%", "ai_25=@lo", "ai_25=@hi", "基準 75.0 十億日圓（公司說明資料；校準錨，改變 AI 裝置每單位新機需求的營收）"),
    ("w", "AI 出貨中 HBM 比重 45%／52%", "w_hbm=@lo", "w_hbm=@hi", "基準 48%"),
    ("dlt", "已裝機產能汰換比例 5%／15%", "dlt=@lo", "dlt=@hi", "基準 10%（Assumed）"),
    ("hbe", "混合鍵合對每道 HBM 工序裝置支出 −20%／＋30%", "hb_eff=@lo", "hb_eff=@hi", "基準 +10%（Assumed）"),
    ("sstk", "每個堆疊額外研削／切割工序 2／4", "s_stk=@lo", "s_stk=@hi", "基準 3（Assumed）"),
    ("tail", "CY2030–31 研削需求年增 0%／20%", "g_tail=@lo", "g_tail=@hi", "基準 10%（Assumed）"),
    ("gn", "非 AI 裝置營收年增 0%／12%", "g_non=@lo", "g_non=@hi", "基準 6%（Assumed）"),
    ("gtl", "非 AI 消耗品年增 0%／10%", "g_tl=@lo", "g_tl=@hi", "基準 5%（Assumed）"),
    ("wtl", "消耗品中 AI 工序比例 10%／40%", "w_tlai=@lo", "w_tlai=@hi", "基準 27.7%（Derived）"),
    ("shl", "市占 FY2029 下滑 10 個百分點（只有低案）", "sh_loss=@hi", None, "基準不下滑"),
    ("pbt", "稅前 ÷ 營業利益 0.97／1.00", "pbt_oi=@lo", "pbt_oi=@hi", "基準 0.99；改變淨利 ÷ 營業利益"),
    ("q4", "FY2026 第四季營收 ÷ 第三季 持平／＋22%", "g_q4=@lo", "g_q4=@hi", "基準 +10%（Assumed）"),
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


jobs0 = {"base": {"DISCO_SNAP": "1"}}
rc0 = run(jobs0, "base")
b = read(rc0, "base")
JOBS = {}
for key, _, lo, hi, _ in CASES:
    if lo: JOBS[f"{key}__lo"] = {"DISCO_OVERRIDE": lo}
    if hi: JOBS[f"{key}__hi"] = {"DISCO_OVERRIDE": hi}
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

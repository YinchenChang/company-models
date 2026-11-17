# ABB 敏感度快照：每個案例以環境變數 ABB_OVERRIDE 覆寫 In 頁輸入後整本重建、LibreOffice 重算，讀加權目標價（美元）與 DCF，寫入 abb/sens.json；
# 正式建置（build_abb.py）讀入 Sens 頁下段「經營輸入的整本重算快照」。In 頁或模型改動後重跑一次，再正式建置（否則 Sens 頁模型狀態顯示 STALE）。
# 用法：python3 abb/sens_abb.py [暫存資料夾，預設 /tmp/abb_w/sens]；LO_PROFILE 指定 LibreOffice 私有設定檔（file:///tmp/lo_n5_abb）。
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/abb_w/sens"
BUILD = os.path.join(HERE, "build_abb.py")
env0 = {k: v for k, v in os.environ.items() if k not in ("ABB_OVERRIDE", "ABB_SNAP")}
grp = lambda ks, side: ";".join(f"{k}={side}" for k in ks)
G = lambda ks: (ks, grp(ks, "@lo"), grp(ks, "@hi"))
Q4K = ["q4g_el", "q4g_mo", "q4g_au", "q4cnet", "q4m_el", "q4m_mo", "q4m_au", "q4co", "br_q4", "fin_q4", "rob_q4"]
CASES = [
    ("el_g", "Electrification 非資料中心成長 低／高", *G(["g_el27", "g_el32"]), "FY2027 6%／11%、FY2032 3.5%／6.5%（基準 9%／5%）"),
    ("mo_g", "Motion 成長 低／高", *G(["g_mo27", "g_mo32"]), "FY2027 4%／8%、FY2032 2.5%／5.5%（基準 6%／4%）"),
    ("au_g", "Automation 成長 低／高", *G(["g_au27", "g_au32"]), "同 Motion"),
    ("sp_g", "情境 A／C 的成長差距（加幅、減幅）低／高", *G(["dg_up", "dg_dn"]), "A 加 1%／3%、C 減 2%／4%（基準 2%／3%）；基準情境 B 不受影響，只動 A、C"),
    ("dc_sh", "ABB 占 Wattonomics 配電＋UPS 設施支出的份額 低／高", *G(["shr27", "shr32"]), "FY2027 6%／9%、FY2032 5%／9%（基準 7.5%／7%）"),
    ("dc_sp", "情境 A／C 的份額差距 低／高", *G(["dsh_up", "dsh_dn"]), "只動 A、C"),
    ("dc_cal", "資料中心占 Electrification 的校準（FY2025、FY2026）低／高", *G(["dc25_pct", "dc26_pct"]), "FY2025 10%／18%、FY2026 17%／27%（基準 14%／22%）；只改非資料中心的起點"),
    ("wif", "Wattonomics 配電＋UPS 設施支出係數 ×0.8／×1.2", *G(["wif_k"]), "三情境同乘；測 ABB 對 Wattonomics 的敏感度"),
    ("m_el", "Electrification Operational EBITA 率 低／高", *G(["m_el27", "m_el32"]), "FY2027 24.5%／27%、FY2032 24.5%／28%（基準 26%／26.5%）"),
    ("m_mo", "Motion Operational EBITA 率 低／高", *G(["m_mo27", "m_mo32"]), "FY2027 18%／20.5%（基準 19.3%）"),
    ("m_au", "Automation Operational EBITA 率 低／高", *G(["m_au27", "m_au32"]), "FY2027 14%／16.5%（基準 15.2%）"),
    ("dm", "情境 A／C 的利潤率差距 低／高", *G(["dm_up", "dm_dn"]), "只動 A、C"),
    ("corp", "公司與其他成本、沖銷 低／高", *G(["co_pct", "cnet"]), "成本占 1.8%／0.8%（基準 1.3%）；沖銷 1%／1.8%（基準 1.4%）"),
    ("gm", "毛利率 39.5%／41.5%", *G(["gm"]), "只影響損益表分項（營業利益由 Operational EBITA 推得），目標價不變為預期"),
    ("br", "Operational EBITA → 營業利益的橋 −1.7%／−0.7%（占營收）", *G(["br_pct"]), "基準 −1.2%"),
    ("fin", "業外淨額、非控制權益、股數變動 低／高", *G(["fin_f", "nci_pct", "sh_g"]), "業外 0／+0.2；非控制權益 1.5%／3%；股數 −1.5%／0%（基準 +0.12、2%、−1%）；低／高方向不一致，以各鍵低／高欄原值同時帶入"),
    ("tax", "有效稅率 23%／26%", *G(["tax_f"]), "基準 24.5%；低＝23%（目標價較高）"),
    ("q4", "FY2026 第四季假設 低／高（全部 Q4 輸入）", *G(Q4K), "Q4 營收成長、利潤率、公司成本、橋、業外與 Robotics 淨利同時取低／高"),
    ("rob", "Robotics 價值（出售淨款或保留事業）35／54 億美元", *G(["inv_v"]), "基準 50 億（企業價值約 54 億美元，扣稅與分離成本假設）；低 35 億＝中國禁止收購其中國業務（2026-11-05 報導）、交易失敗且保留事業價值打折；高 54 億＝以企業價值全額計"),
    ("mi", "非控制權益 低／高", *G(["mi_v"]), "低＝Yahoo 2026-09 帳面 0.483、高＝3.0 十億美元（ABB India 少數股權市價假設，Assumed；基準帳面 0.559）"),
    ("wts", "情境權重：全給保守 C／全給樂觀 A", "tp_wA=0;tp_wB=0;tp_wC=1", "tp_wA=1;tp_wB=0;tp_wC=0", "基準權重 A 0.3／B 0.5／C 0.2"),
]
# CASES 元素：(代號, 白話標籤, 鍵清單, 低案覆寫, 高案覆寫, 設定說明)；wts 無鍵清單
CASES = [(c if len(c) == 6 else (c[0], c[1], ["tp_wA", "tp_wB", "tp_wC"], c[2], c[3], c[4])) for c in CASES]


def lab(ws, prefix):
    return next(r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix))


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
    return dict(tp=V.cell(lab(V, "加權目標價（美元）"), 2).value, tp_chf=V.cell(lab(V, "目標價（DCF 與 Comparable 加權）"), 2).value,
                dcf=V.cell(lab(V, "DCF 目標價（情境加權；美元）"), 2).value)


b = None
rc0 = run({"base": {"ABB_SNAP": "1"}}, "base")
b0 = read(rc0, "base"); b = dict(tp_usd=b0["tp"], tp_chf=b0["tp_chf"], dcf_usd=b0["dcf"])
JOBS = {}
for key, _, ks, lo, hi, _ in CASES:
    JOBS[f"{key}__lo"] = {"ABB_OVERRIDE": lo}; JOBS[f"{key}__hi"] = {"ABB_OVERRIDE": hi}
rc = run(JOBS, "cases")
biz = {}
for nm in JOBS:
    k, side = nm.split("__")
    r = read(rc, nm)
    biz.setdefault(k, {})["低" if side == "lo" else "高"] = dict(tp=r["tp"], dcf=r["dcf"])
    print(f"{nm:12s} 加權 {r['tp']:.2f}（{r['tp'] / b['tp_usd'] - 1:+.2%}）DCF {r['dcf']:.2f}")
out = {"_base": b, "cases": [{"key": k, "label": l, "keys": "、".join(ks), "setting": s} for k, l, ks, _, _, s in CASES], "biz": biz}
json.dump(out, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
print("基準", b)

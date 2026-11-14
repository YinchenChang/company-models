"""WhiteFiber v0.4 敏感度（CoreWeave v5.0 樣板）：一次動一個輸入（In 頁低／高），整本重建＋LibreOffice 重算 → data/sens_v50.json（建置時寫入 Sens 頁）。
用法：python3 whitefiber/scripts/sens_v50.py [--jobs 2]"""
import datetime, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import runner_v50 as RN

JOBS = int(sys.argv[sys.argv.index("--jobs") + 1]) if "--jobs" in sys.argv else 2
RG = RN.in_ranges()
GEN = ("h100", "gb200", "gb300", "vr200")
# (鍵欄, 標籤, 設定說明, 低 ovr, 高 ovr)
def one(k, lab, note=None):
    lo, b, hi = RG[k]
    f = lambda v: None if v is None else {k: v}
    fmt = lambda v: "—" if v is None else (f"{v:.4g}" if isinstance(v, (int, float)) else str(v))
    return (k, lab, note or f"{fmt(lo)}／{fmt(hi)}（基準 {fmt(b)}）", f(lo), f(hi))
CASES = [
    ("he_all", "IF_HoldEcon 四世代同時（Tokenomics 低／高成本）", "Tokenomics 低成本／高成本", {f"he_{g}": RG[f"he_{g}"][0] for g in GEN}, {f"he_{g}": RG[f"he_{g}"][2] for g in GEN}),
    one("k_adj", "定價倍數 k 調整"), one("p_long", "長約價格證據（IREN–Microsoft，Analogy）"), one("od_share", "隨需占計費產能比例"), one("u_bill", "計費比例"),
    one("p_new_ovr", "新增 MW 每 MW 年收入改用公司合約換算值（Interested-party）"), one("mw_q3", "2026-09-30 雲端在役 MW 假設"),
    one("rent_mw", "雲端承租機房每 MW 年租金"),
    one("sga_r", "營業費用率（一般及管理，扣 SBC）"), one("sbc_r", "股票薪酬 ÷ 營收"), one("lam", "成長型資本支出提前比例 λ"),
    one("life_ai", "AI 設備折舊年限"), one("cip_life", "在建工程折舊年限（全以 AI 設備年限）"),
    one("om_h100", "2026-06-30 世代占比：H100（差額由 GB300 吸收）"), one("om_gb200", "2026-06-30 世代占比：GB200（差額由 GB300 吸收）"),
    one("nm27_gb300", "2027 年新增 GB300 占比"), one("nm28_gb300", "2028 年起新增 GB300 占比"),
    one("q4cx_ovr", "Q4 資本支出改用其他值"), one("pp_q4", "Q4 額外客戶預付"),
    one("prepay_cov", "客戶預付覆蓋比"), one("nwc_r", "營運資金增加 ÷ 營收增量"), one("eq_cap", "每年新股上限占市值（原則 62）"),
    one("r_junk", "高息債利率"), one("lev_max", "投資級債上限（借款 ÷ EBITDA）"), one("kd", "稅前債務成本 kd（WACC 與新債利率）"),
    one("tax_lt", "長期現金稅率"), one("nol0", "NOL 可用餘額"), one("rsu_net", "RSU＋選擇權淨股數"),
    one("colo_core", "託管核心營收（NC-1＋MTL）"), one("mgd_27", "代管服務營收 FY2027"), one("mgd_28", "代管服務營收 FY2028 起"), one("colo_m", "託管與代管 EBITDA 率"), one("cxo_27", "託管資本支出 FY2027（含 NC-2／NC-3 初期工程）"),
    one("beta_u", "無槓桿 β（自算產業 β ±0.3）"), one("erp", "股權風險溢酬"), one("wacc_adj", "WACC ±1pt"), one("ronic_adj", "RONIC − WACC"), one("tp_g", "永續成長率 g"),
    one("w_dcf", "DCF 權重"), one("peer_nd_mode", "同業淨負債改用共識年底預估（替代口徑）"), one("mult_adj", "Comparable 倍數 ±20%"),
    ("mwB_27 mwB_28 mwB_29 mwB_30", "融資受限：建置延後一年（情境 B 路徑後移；原則 9）", "B 路徑 25／35／43／51", {"mwB_27": 25, "mwB_28": 35, "mwB_29": 43, "mwB_30": 51}, None),
]
runs = [("base", {}, {})]
for i, (k, lab, st, lo, hi) in enumerate(CASES):
    if lo: runs.append((f"{i}lo", lo, {}))
    if hi: runs.append((f"{i}hi", hi, {}))
GRID = [(w, g) for w in (-0.01, 0.0, 0.01) for g in (0.02, 0.03, 0.04)]
for w, g in GRID: runs.append((f"g{w}_{g}", {"wacc_adj": w, "tp_g": g}, {}))
res = RN.run(runs, JOBS)
b = res["base"]
out = {"date": datetime.date.today().isoformat(), "base": {"tp": b["tp"], "fp": b["fp"], "dcf": b["dcf"]}, "cases": [], "grid": []}
out["cases"].append({"key": "rel_mult", "label": "相對倍數係數＝1（原則 47；基準即為 1，未調整）", "setting": "1.00／—", "lo": {k: b[k] for k in ("tp", "dcf", "gate", "rec")}})
for i, (k, lab, st, lo, hi) in enumerate(CASES):
    c = {"key": k, "label": lab, "setting": st}
    for side in ("lo", "hi"):
        x = res.get(f"{i}{side}")
        if x: c[side] = {kk: x[kk] for kk in ("tp", "dcf", "gate", "rec")}
    out["cases"].append(c)
out["grid"] = [[res[f"g{w}_{g}"]["tp"] for g in (0.02, 0.03, 0.04)] for w in (-0.01, 0.0, 0.01)]
json.dump(out, open(os.path.join(RN.HERE, "data", "sens_v50.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("sens done: base", b["tp"], "cases", len(CASES))

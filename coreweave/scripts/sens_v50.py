"""CoreWeave v5.0 敏感度：一次動一個輸入（In 頁低／高），整本重建＋LibreOffice 重算 → data/sens_v50.json（建置時寫入 Sens 頁）。
用法：python3 coreweave/scripts/sens_v50.py [--jobs 2]"""
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
    one("p_new_ovr", "J4：新增 MW 每 MW 年收入改用公司宣稱中點（Interested-party）"), one("rent_mw", "每 MW 年租金（高＝Q2 實際租賃費用比率，A11）"),
    one("sga_r", "管銷費用率（扣 SBC）"), one("sbc_r", "股票薪酬 ÷ 營收"), one("lam", "成長型資本支出提前比例 λ"),
    one("life_ai", "AI 設備折舊年限（會計，Tokenomics）"),
    ("life_ai", "AI 設備經濟汰換壓力（低於 Tokenomics 區間；Assumed，待 Tokenomics 登錄）", "4 年／—", {"life_ai": 4}, None),
    one("om_h100", "2026-06-30 世代占比：H100（差額由 GB300 吸收）"), one("om_gb200", "2026-06-30 世代占比：GB200（差額由 GB300 吸收）"),
    one("nm27_gb300", "2027 年新增 GB300 占比"), one("nm28_gb300", "2028 年起新增 GB300 占比"),
    one("q4cx_ovr", "J5：Q4 資本支出改用指引隱含值"), one("pp_q4", "Q4 額外客戶預付（公司稱約 9）"),
    one("prepay_cov", "客戶預付覆蓋比"), one("nwc_r", "營運資金增加 ÷ 營收增量（A12）"), one("eq_cap", "每年新股上限占市值（A12、原則 62）"),
    one("r_junk", "高息債利率（A12）"), one("lev_max", "投資級債上限（借款 ÷ EBITDA）"), one("kd", "稅前債務成本 kd（WACC 與新債利率）"),
    one("tax_lt", "長期現金稅率（B3）"), one("rsu_net", "RSU＋選擇權淨股數（待 Q3 10-Q）"),
    one("beta_u", "無槓桿 β（自算產業 β ±0.3）"), one("erp", "股權風險溢酬"), one("wacc_adj", "WACC ±1pt"), one("ronic_adj", "RONIC − WACC"), one("tp_g", "永續成長率 g"),
    one("w_dcf", "DCF 權重（A1）"), one("mult_adj", "Comparable 倍數 ±20%"),
    ("mwB_27 mwB_28 mwB_29 mwB_30", "融資受限：建置延後一年（情境 B 路徑後移；原則 9）", "B 路徑 2,475／3,100／4,000／4,800", {"mwB_27": 2475, "mwB_28": 3100, "mwB_29": 4000, "mwB_30": 4800}, None),
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

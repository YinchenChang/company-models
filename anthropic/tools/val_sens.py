#!/usr/bin/env python3
"""Val 敏感度表（龍捲風；工作單 20261111 F9）：一次改一個 Inputs → LibreOffice 重算整本活頁簿 → 讀加權目標價 VAL_TP_W → 寫 data/val_sens.yaml。

情境 A／C 的營運快照（data/val_scenarios.yaml）在各案中不變：F9 各案皆為估值層參數（折現率、延伸期、股數、淨現金、倍數調整）
或只影響現金與股數的發行條件，不改營運現金流；所以各案的 A／B／C 三情境 DCF 都是活公式重算的結果。
單邊案例（開關）：低值欄＝替代讀法、高值欄＝基準（標「基準」）。
用法：python3 tools/val_sens.py model/<現行>.xlsx
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
from val_recalc import ROOT, read_names, run_many  # noqa: E402

# key、案例名稱、低值設定、高值設定（None＝基準）、標籤
CASES = [
    ("wacc", "WACC ±1pt", {"val_wacc_adj": 0.01}, {"val_wacc_adj": -0.01}, "WACC＋1pt", "WACC−1pt"),
    ("g", "永續成長率 g 2%／4%", {"val_g": 0.02}, {"val_g": 0.04}, "g 2%", "g 4%"),
    ("beta", "β 1.6／1.0", {"val_beta_override": 1.6}, {"val_beta_override": 1.0}, "β 1.6", "β 1.0"),
    ("erp", "ERP 6%／4%", {"val_erp": 0.06}, {"val_erp": 0.04}, "ERP 6%", "ERP 4%"),
    ("ronic", "RONIC＝WACC＋5pt", None, {"val_ronic_add": 0.05}, "基準（RONIC＝WACC）", "RONIC＝WACC＋5pt"),
    ("extg", "延伸期 2033 營收年增 5%／15%", {"ext_g33": 0.05}, {"ext_g33": 0.15}, "ext_g33 5%", "ext_g33 15%"),
    ("extm", "延伸期 EBIT 利益率加碼 +5pt／+10pt（至 2033）", {"ext_m_add": 0.05}, {"ext_m_add": 0.10}, "+5pt", "+10pt"),
    ("px", "IPO 價格 70／78（淨募資、私募股數、選擇權庫藏股與現價連動）", {"ipo_px": 70}, {"ipo_px": 78}, "$70", "$78"),
    ("pbi", "首次公益發行 846.84M 股不計", None, {"val_pbi_switch": 0}, "基準（計入）", "不計"),
    ("vmo", "捐贈配對已歸屬義務 553.44M 股不計", None, {"val_vmo_switch": 0}, "基準（計入）", "不計"),
    ("nol", "2026Q4 以前 NOL 視為無限（明確期不繳稅）", None, {"val_nol_inf_switch": 1}, "基準", "NOL 無限"),
    ("wc", "營運資金 ±5% 營收增量", {"val_wc_rate": 0.05}, {"val_wc_rate": -0.05}, "＋5%（投入）", "−5%（釋出）"),
    ("comp", "Comparable 不調整（倍數不乘 FCF 利益率比例）", None, {"val_comp_adj_switch": 0}, "基準（調整）", "不調整"),
    ("oa", "超額配售計入（現金＋9.1、股數＋123.75M）", None, {"p_overallot_switch": 1}, "基準（不計）", "計入"),
    ("q4p", "2026Q4 EBIT 前九月代理改工作單字面（調整後營業損益 4.081）", {"val_q4_proxy_switch": 0}, None, "字面讀法", "基準（GAAP −1.260）"),
]


def reader(path):
    v, _ = read_names(path, ["VAL_TP_W", "VAL_Up_W", "VAL_Rating"])
    return v


def main():
    model = Path(sys.argv[1]).resolve()
    jobs = {"base": {}}
    for k, _, lo, hi, *_ in CASES:
        if lo:
            jobs[f"{k}_lo"] = lo
        if hi:
            jobs[f"{k}_hi"] = hi
    res = run_many(model, jobs, reader, workers=4)
    base = res["base"]["VAL_TP_W"]
    cases = []
    for k, name, lo, hi, ll, hl in CASES:
        lo_r = res[f"{k}_lo"] if lo else res["base"]
        hi_r = res[f"{k}_hi"] if hi else res["base"]
        cases.append(dict(key=k, name=name, tp_lo=float(lo_r["VAL_TP_W"]), tp_hi=float(hi_r["VAL_TP_W"]), lo_label=ll, hi_label=hl,
                          rating_lo=lo_r["VAL_Rating"], rating_hi=hi_r["VAL_Rating"], up_lo=float(lo_r["VAL_Up_W"]), up_hi=float(hi_r["VAL_Up_W"]),
                          overrides_lo={a: b for a, b in (lo or {}).items()}, overrides_hi={a: b for a, b in (hi or {}).items()}))
    out = dict(note="tools/val_sens.py 產生；LibreOffice 重算整本活頁簿後的 VAL_TP_W；勿手改", model=model.name, date=dt.date.today().isoformat(),
               base_tp=float(base), base_rating=res["base"]["VAL_Rating"], cases=cases,
               gap=["EV／NTM 毛利（毛利＝營收淨額 − 算力）：同業毛利率或 EV／毛利倍數未提供（工作單 F6 只給 EV／遠期營收與 FCF 利益率）→ 資料缺口，未計算"])
    (ROOT / "data" / "val_sens.yaml").write_text(yaml.safe_dump(out, allow_unicode=True, sort_keys=False, width=200), encoding="utf-8")
    print("base", round(base, 3), res["base"]["VAL_Rating"])
    for cs in cases:
        print(f"{cs['name']}: {cs['tp_lo']:.2f}（{cs['rating_lo']}） / {cs['tp_hi']:.2f}（{cs['rating_hi']}）")


if __name__ == "__main__":
    main()

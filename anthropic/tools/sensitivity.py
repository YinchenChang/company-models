#!/usr/bin/env python3
"""v0.1 完成報告：關鍵驅動敏感度（engine 重算現行活頁簿；不另寫計算）。

用法：python3 tools/sensitivity.py [--json build_out/sensitivity.json] [--md build_out/sensitivity.md]
每個驅動把 Inputs（或 TK 快照）的值改為該列自己的「低」「高」欄（讀活頁簿；本檔不寫任何參數值），以 engine 重算後讀：
2030 每 VR 等值 GW 差額（COST_PropGap_VR）、2030 營收淨額（REV_NetCapped）、累計外部資金需求 2030（FND_ExtNeedCum）、
只靠已到位融資的現金谷底（FND_CashNoExtTrough）、2030 年底現金（FND_CashEnd）。排序：|Δ 2030 差額| 最大者在前（同分看 |Δ 累計外部資金需求|、|Δ 現金谷底|）。
"""
import argparse
import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
from engine import Engine, current_model_path  # noqa: E402

# （驅動名稱, 群組, [Inputs key 或 TK 名稱]）——「低」「高」取各列自己的區間欄
DRIVERS = [
    ("API 任務數成長 2027–30", "營收", ["api_gn_2027", "api_gn_2028", "api_gn_2029", "api_gn_2030"]),
    ("API 每任務 token 成長 2026–30", "營收", ["api_gk_2026", "api_gk_2027", "api_gk_2028", "api_gk_2029", "api_gk_2030"]),
    ("價格彈性（任務數）", "營收", ["api_elast_n"]),
    ("API 牌價年變動", "營收／容量", ["price_chg"]),
    ("雲端平台抽成率", "營收", ["partner_fee_rate"]),
    ("每 GW 年合約價", "算力", ["contract_price"]),
    ("NonNV 產出比（TPU、Trainium、AMD）", "算力", ["TK:TPUv7_Out", "TK:Trn3_Out", "TK:MI455X_Out"]),
    ("AMD 合約期間", "算力", ["amd_term"]),
    ("「最高可達」GW 計入比例（AMD）", "算力", ["upto_share"]),
    ("爬坡年數", "算力", ["ramp_years"]),
    ("自建資本支出均攤年數", "算力", ["fluid_years"]),
    ("員工人數年增率 2026–30", "非算力", ["hc_g_2026", "hc_g_2027", "hc_g_2028", "hc_g_2029", "hc_g_2030"]),
    ("每人年成本年變動", "非算力", ["pp_cost_g"]),
    ("最低現金月數", "融資", ["min_cash_months"]),
]
OUT = [("gap30", "COST_PropGap_VR", 5), ("rev30", "REV_NetCapped", 5), ("cum30", "FND_ExtNeedCum", 5), ("trough", "FND_CashNoExtTrough", None),
       ("cash30", "FND_CashEnd", 5), ("gap27", "COST_PropGap_VR", 2)]


def read(eng):
    out = {}
    for k, n, i in OUT:
        v = eng.get_name(n)
        out[k] = v[i] if i is not None else v
    return out


def inp_cells(eng, wb_inputs, key):
    """回傳（值格名稱, 低, 高）：Inputs 以 key 欄（N）找列；TK NonNV 以具名範圍 TK_NNV_*_Out／OutLo／OutHi。"""
    if key.startswith("TK:"):
        base = "TK_NNV_" + key[3:]
        return base, eng.get_name(base + "Lo"), eng.get_name(base + "Hi")
    r = wb_inputs[key]
    return r["id"], r["lo"], r["hi"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", type=Path, default=REPO / "build_out" / "sensitivity.json")
    ap.add_argument("--md", type=Path, default=REPO / "build_out" / "sensitivity.md")
    a = ap.parse_args()
    eng = Engine(current_model_path())
    eng.evaluate_all()
    rows = eng.get("Inputs", "A5:N600")
    inputs = {r[13]: dict(id=r[0], lo=r[5], hi=r[6], val=r[4]) for r in rows if r[0]}
    base = read(eng)
    res = []
    for name, grp, keys in DRIVERS:
        cells = [inp_cells(eng, inputs, k) for k in keys]
        orig = {c: eng.get_name(c) for c, _l, _h in cells}
        one = dict(driver=name, group=grp, inputs=[c for c, *_ in cells])
        for side, idx in (("lo", 1), ("hi", 2)):
            for c in cells:
                eng.set_name(c[0], c[idx])
            eng.evaluate_all()
            one[side] = read(eng)
            one[side + "_set"] = {c[0]: c[idx] for c in cells}
            for c in cells:
                eng.set_name(c[0], orig[c[0]])
        eng.evaluate_all()
        one["swing_gap30"] = max(abs(one[s]["gap30"] - base["gap30"]) for s in ("lo", "hi"))
        one["swing_cum30"] = max(abs(one[s]["cum30"] - base["cum30"]) for s in ("lo", "hi"))
        one["swing_trough"] = max(abs(one[s]["trough"] - base["trough"]) for s in ("lo", "hi"))
        res.append(one)
    res.sort(key=lambda d: (-d["swing_gap30"], -d["swing_cum30"], -d["swing_trough"]))
    a.json.parent.mkdir(parents=True, exist_ok=True)
    json.dump(dict(base=base, drivers=res), open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    f = lambda x: f"{x:,.2f}"  # noqa: E731
    md = [f"基準：2030 每 VR 等值 GW 差額 {f(base['gap30'])}、2030 營收淨額 {f(base['rev30'])}、累計外部資金需求 {f(base['cum30'])}、"
          f"只靠已到位融資的現金谷底 {f(base['trough'])}、2030 年底現金 {f(base['cash30'])}",
          "", "| 排名 | 驅動（群組） | 設定（低／高＝各列區間欄） | 2030 營收淨額 | 2027 差額 | 2030 差額 | 累計外部資金需求 | 現金谷底 |", "|---|---|---|---|---|---|---|---|"]
    for i, d in enumerate(res, 1):
        sets = "／".join(",".join(f"{v:g}" for v in d[s + "_set"].values()) for s in ("lo", "hi"))
        md.append(f"| {i} | {d['driver']}（{d['group']}） | {sets} | {f(d['lo']['rev30'])}／{f(d['hi']['rev30'])} | {f(d['lo']['gap27'])}／{f(d['hi']['gap27'])} | "
                  f"{f(d['lo']['gap30'])}／{f(d['hi']['gap30'])} | {f(d['lo']['cum30'])}／{f(d['hi']['cum30'])} | {f(d['lo']['trough'])}／{f(d['hi']['trough'])} |")
    a.md.write_text("\n".join(md) + "\n", encoding="utf-8")
    print("\n".join(md))


if __name__ == "__main__":
    main()

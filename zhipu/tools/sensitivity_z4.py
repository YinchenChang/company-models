#!/usr/bin/env python3
"""Z4 敏感度與翻轉條件（工作單 Z4 第 3 步；報告用）。

以 engine（pycel 執行現行 Excel）逐一改寫 Inputs，讀取命題輸出；不在 Python 另算任何模型數值（只排序、找轉折點）。
用法：python3 tools/sensitivity_z4.py [--yaml tools/sensitivity_z4.yaml] [--json out.json]
輸出：每個情境的 2030 每 VR 等值 GW 差額、首次轉正年、覆蓋率、累計外部資金需求（基準／轉股）、2030 年底現金、首次缺口年；
      以及「翻轉條件」：對單一輸入二分搜尋，找出使 2030 差額＝0（或首次出現外部資金需求）的值。
"""
import argparse
import json
import sys
from pathlib import Path

import yaml

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
from engine import Engine, current_model_path, resolve_key  # noqa: E402

YEARS = [2025, 2026, 2027, 2028, 2029, 2030]


def _one(v):
    return v[0] if isinstance(v, list) else v


def metrics(e):
    gap = e.get_name("COST_PropGap_VR")
    first_pos = next((y for y, g in zip(YEARS, gap) if g > 0), "無")
    cum = e.get_name("FND_ExtNeedCum")
    return dict(
        gap30=gap[5], first_pos=first_pos, cov30=e.get_name("COST_Coverage")[5],
        gapcash30=e.get_name("COST_GapCash")[5], sup30=e.get_name("CMP_SupplyGW")[5],
        cum30=cum[5], cumconv30=e.get_name("FND_ExtNeedCumConv")[5], cash30=e.get_name("FND_CashEnd")[5],
        first_gap=_one(e.get_name("FND_FirstGapYear")), peak=_one(e.get_name("FND_PeakExtNeed")),
        fcf=e.get_name("FND_FCF"), rvs30=e.get_name("RVS_ExtNeedCum")[5])


class Runner:
    def __init__(self, path):
        self.e = Engine(path)
        self.orig = {}

    def apply(self, inputs):
        for k, v in inputs.items():
            sheet, coord = resolve_key(self.e.names, k)
            if k not in self.orig:
                self.orig[k] = self.e.get(sheet, coord)
            self.e.set_input(sheet, coord, v)

    def reset(self):
        for k, v in self.orig.items():
            self.e.set_input(*resolve_key(self.e.names, k), v)

    def run(self, inputs):
        self.apply(inputs)
        m = metrics(self.e)
        self.reset()
        return m

    def solve(self, key, lo, hi, target, fixed=None, tol=1e-4, it=40):
        """二分搜尋 key ∈ [lo, hi] 使 target(metrics) 由負轉正（回傳轉折值；端點同號回傳 None）。"""
        fixed = fixed or {}
        f = lambda x: target(self.run({**fixed, key: x}))  # noqa: E731
        flo, fhi = f(lo), f(hi)
        if (flo > 0) == (fhi > 0):
            return None
        for _ in range(it):
            mid = (lo + hi) / 2
            fm = f(mid)
            if (fm > 0) == (flo > 0):
                lo, flo = mid, fm
            else:
                hi, fhi = mid, fm
            if abs(hi - lo) < tol:
                break
        return (lo + hi) / 2


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--yaml", type=Path, default=REPO / "tools" / "sensitivity_z4.yaml")
    ap.add_argument("--json", type=Path)
    a = ap.parse_args()
    cfg = yaml.safe_load(a.yaml.read_text(encoding="utf-8"))
    R = Runner(current_model_path())
    out = {"scenarios": [], "flips": []}
    base = R.run({})
    for s in cfg["scenarios"]:
        m = R.run(s.get("inputs", {}))
        out["scenarios"].append(dict(id=s["id"], label=s["label"], group=s.get("group", ""), inputs=s.get("inputs", {}), **m,
                                     d_gap30=m["gap30"] - base["gap30"], d_cash30=m["cash30"] - base["cash30"]))
        print(f"{s['id']:<24} gap30 {m['gap30']:>9.1f}  first+ {m['first_pos']!s:>5}  cum30 {m['cum30']:>8.2f}  cash30 {m['cash30']:>8.1f}  gap {m['first_gap']}")
    for fl in cfg.get("flips", []):
        tgt = {"gap30": lambda m: m["gap30"], "cum30": lambda m: m["cum30"] - 1e-9}[fl["target"]]
        x = R.solve(fl["key"], fl["lo"], fl["hi"], tgt, fl.get("fixed"))
        out["flips"].append(dict(fl, value=x))
        print("flip", fl["id"], x)
    out["base"] = base
    if a.json:
        a.json.write_text(json.dumps(out, ensure_ascii=False, indent=1, default=str), encoding="utf-8")


if __name__ == "__main__":
    main()

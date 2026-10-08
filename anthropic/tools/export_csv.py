#!/usr/bin/env python3
"""匯出 SRC_OAI、Inputs、TK_Link、Checks → data/export/*.csv，並檢查 CHK_Errors（CI governance job 用）。
值由 engine（pycel）重算 model/CURRENT 指向的活頁簿；不寫回 Excel。CHK_Errors ≠ 0 時以非零狀態結束。"""
import argparse
import csv
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
from engine import Engine  # noqa: E402

SHEETS = {"SRC_OAI": "X", "Inputs": "M", "TK_Link": "X", "Checks": "F"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", type=Path, default=REPO / "data" / "export")
    ap.add_argument("--check-only", action="store_true")
    a = ap.parse_args()
    eng = Engine()
    errors = eng.get_name("CHK_Errors")
    print(f"{eng.path.name}: CHK_Errors={errors}")
    if not a.check_only:
        a.out.mkdir(parents=True, exist_ok=True)
        for s, last in SHEETS.items():
            grid = eng.get(s, f"A4:{last}700")
            rows = [r for r in grid if any(x != "" for x in r)]
            with open(a.out / f"{s}.csv", "w", newline="", encoding="utf-8-sig") as f:
                csv.writer(f).writerows(rows)
            print(f"  {s}.csv：{len(rows)} 列")
    if errors != 0:
        print("CHK_Errors 不為 0", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Oracle v0.4.7 反解（個股模型原則第 10 節第 4 條）：只動 p_long，以整本重算（build_oracle.py ＋ LibreOffice）
求 (a) 加權目標價＝現價、(b) 上檔空間＝−25%（評等由放空翻為中立）所需的 p_long。結果寫入 oracle/data/reverse_v047.json。
用法：python3 oracle/scripts/reverse_v047.py
"""
import json, os, sys, tempfile, shutil, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sens_v04 import run, HERE   # noqa: E402

PX = 115.16          # In 頁 px（與 build_oracle.py 同值；改現價時同步）


def solve(target, tmp, tag, x0=9.70, x1=12.7247, tol=0.02):
    f = lambda x, i: run({"p_long": x}, f"{tag}_{i}", tmp)["tp"] - target
    y0, y1 = f(x0, 0), f(x1, 1)
    for i in range(2, 12):
        x2 = x1 - y1 * (x1 - x0) / (y1 - y0)
        y2 = f(x2, i)
        print(f"{tag}: p_long {x2:.4f} → 差 {y2:+.4f}")
        if abs(y2) < tol:
            return x2
        x0, y0, x1, y1 = x1, y1, x2, y2
    raise SystemExit(f"{tag} 未收斂")


def main():
    tmp = tempfile.mkdtemp(prefix="orcl_rev_")
    try:
        out = {"date": datetime.date.today().isoformat(), "px": PX,
               "eq_px": round(solve(PX, tmp, "eq_px"), 4),
               "neutral": round(solve(PX * 0.75, tmp, "neutral"), 4)}
        json.dump(out, open(os.path.join(HERE, "data", "reverse_v047.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print(out)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()

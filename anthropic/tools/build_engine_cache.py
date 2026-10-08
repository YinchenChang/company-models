#!/usr/bin/env python3
"""建置 engine 的 pycel 序列化快取（CI 的 engine-cache job 用；2026-10-04 提速指令第 2 點）。

用法：python3 tools/build_engine_cache.py engine-cache/engine.pkl
快取只存 pycel 的計算圖與重算後的值，不含任何另寫的公式；載入時仍強制全簿重算（見 engine/core.py）。
指紋（活頁簿、engine 原始碼、pycel／Python 版本）不符時載入會直接報錯。
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from engine import Engine  # noqa: E402

out = Path(sys.argv[1])
out.parent.mkdir(parents=True, exist_ok=True)
t = time.perf_counter()
eng = Engine()
print(f"{eng.path.name}: 建圖＋重算 {time.perf_counter() - t:.1f} 秒")
t = time.perf_counter()
eng.save_cache(out)
print(f"序列化 {time.perf_counter() - t:.1f} 秒 → {out}（{out.stat().st_size / 1e6:.1f} MB）")

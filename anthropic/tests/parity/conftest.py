import json
import os
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))  # repo 根目錄
sys.path.insert(0, str(Path(__file__).resolve().parent))

# 由 .gitignore 排除；供每輪報告取數。CI 分片時各 job 以 PARITY_RESULTS 指到各自的檔案。
RESULTS = Path(os.environ.get("PARITY_RESULTS") or Path(__file__).parent / "_results.json")


def pytest_collection_modifyitems(config, items):
    """CI 分片（只切分，不改任何測試或斷言）：PARITY_SHARD="i/N" 時，test_scenario_parity 與 test_cache_matches_fresh 只留第 i 片
    （情境依 scenarios.yaml 順序輪流分配，i 從 0 起算）。其他測試不受影響。
    未設定 PARITY_SHARD 時全部照跑。"""
    spec = os.environ.get("PARITY_SHARD")
    if not spec:
        return
    i, n = (int(x) for x in spec.split("/"))
    keep, drop, k = [], [], {}
    for it in items:
        if it.originalname in ("test_scenario_parity", "test_cache_matches_fresh"):   # 兩者依同一順序分片，同一情境落在同一片
            j = k[it.originalname] = k.get(it.originalname, -1) + 1
            (keep if j % n == i else drop).append(it)
        else:
            keep.append(it)
    items[:] = keep
    config.hook.pytest_deselected(items=drop)


@pytest.fixture(scope="session")
def results_store():
    data = json.loads(RESULTS.read_text()) if RESULTS.exists() else {}
    yield data
    RESULTS.write_text(json.dumps(data, ensure_ascii=False, indent=1, default=str))

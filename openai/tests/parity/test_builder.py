"""builder 與 v0.5 遷移的結構性測試（不需 LibreOffice）。"""
import sys
from collections import Counter
from pathlib import Path

import pytest
import yaml

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "builder"))
from v05_map import build_map, load  # noqa: E402
from engine import Engine, current_model_path  # noqa: E402

EXPECT = yaml.safe_load((Path(__file__).parent / "scenarios.yaml").read_text(encoding="utf-8"))["workbook_expectations"]


@pytest.fixture(scope="module")
def R():
    return build_map()


def test_every_v05_leaf_has_a_destination(R):
    assert len(R.leaves) == EXPECT["v05_leaves"]
    assert [p for p in R.leaves if p not in R.dest] == []


def test_tagged_objects_count():
    n = 0

    def w(o):
        nonlocal n
        if isinstance(o, dict):
            n += "tag" in o
            [w(v) for v in o.values()]
        elif isinstance(o, list):
            [w(v) for v in o]
    w(load())
    assert n == EXPECT["v05_tagged_objects"]


def _same(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= 1e-12 * max(abs(a), abs(b), 1)       # LibreOffice 存檔保留 15 位有效數字
    return a == b


def test_src_and_inputs_values_equal_json(R):
    """SRC_OAI／Inputs 內每個有值的列，Excel 具名範圍的值必須等於 builder 由 v0.5 JSON 讀出的值。"""
    eng = Engine(current_model_path())
    for r in R.src_rows + R.inp_rows:
        if r["value"] != "":
            assert _same(eng.get_name(r["id"]), r["value"]), r["id"]


def test_classification_rules(R):
    for r in R.src_rows:
        assert not str(r["tag"]).startswith(("Assumed", "Analogy")), r["id"]
    kinds = Counter(k for k, _, _ in R.dest.values())
    assert kinds["SRC"] > 0 and kinds["INP"] > 0
    for p, (k, ref, why) in R.dest.items():
        if k == "TK":
            assert "TK_Link" in ref


def test_pending_list_resolved(R):
    """r3 E8 已逐項處理 v0.5 分類不明的 A 類項目；日後新增不明項目才會出現在 pending。"""
    assert isinstance(R.pending, list) and R.pending == []


def _e6_scan(ws):
    import re
    seen = 0
    for row in ws.iter_rows(min_row=5):
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                seen += 1
                f = re.sub(r"'?[A-Za-z_]+'?!\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?", "", c.value)    # 跨頁參照
                f = re.sub(r"\$?\b[A-Z]{1,3}\$?\d+\b", "", f)                          # 儲存格參照
                f = re.sub(r"\b[A-Za-z_][A-Za-z_0-9]*\b", "", f)                          # 具名範圍與函數
                f = f.replace("(1-", "(").replace("+1)", ")").replace("(1+", "(")   # 恆等式：1−比例；年數含頭尾的 +1；1＋成長率（P2）
                assert not re.search(r"\d", f), f"{ws.title}!{c.coordinate} 公式含常數：{c.value}"
    return seen


def test_e6_no_constants_in_v9_formulas():
    """E6：Derived_V9 的公式不得內含常數（算式中的『1-比例』恆等式除外）；換算係數與比例必須在 Inputs。"""
    import openpyxl
    wb = openpyxl.load_workbook(current_model_path())
    assert _e6_scan(wb["Derived_V9"]) >= 9


def test_e6_no_constants_in_p2_formulas():
    """E6（P2）：Demand、Revenue 兩頁的公式不得內含常數；恆等式只容許 1−比例、1＋成長率；定義常數（天數、月數、單位換算）在 Inputs。"""
    import openpyxl
    wb = openpyxl.load_workbook(current_model_path())
    assert _e6_scan(wb["Demand"]) >= 200
    assert _e6_scan(wb["Revenue"]) >= 200


def test_v12_pro_constraints_in_src(R):
    """V12：SRC_OAI 只登 pro 的兩個約束（倍數、占比上限），不再登 pro 的用戶人數點值。"""
    pro = [r for r in R.src_rows if r["metric"].startswith("Pro ")]
    assert len(pro) == 2
    assert {r["unit"] for r in pro} == {"倍", "比例"}


def test_e6_no_constants_in_p3_formulas():
    """E6（P3）：Compute 頁的公式不得內含常數；恆等式只容許 1−比例、年數 +1、1＋成長率；定義常數（T→M、等差級數除數）在 Inputs。
    世代名稱與成本情境為文字格（SUMIFS 的鍵），不在公式內。"""
    import openpyxl
    wb = openpyxl.load_workbook(current_model_path())
    assert _e6_scan(wb["Compute"]) >= 600


def test_e6_no_constants_in_p4_formulas():
    """E6（P4）：Cost 頁的公式不得內含常數；恆等式只容許 1−比例、年數 +1、1＋成長率；定義常數（$M→$B、$→$B、等差級數除數）在 Inputs。
    世代名稱與成本情境為文字格（SUMIFS 的鍵），不在公式內。"""
    import openpyxl
    wb = openpyxl.load_workbook(current_model_path())
    assert _e6_scan(wb["Cost"]) >= 500
    assert _e6_scan(wb["Compute"]) >= 620

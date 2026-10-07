"""Excel（LibreOffice 重算）與 engine 的一致性測試（CLAUDE.md 第 3 節）。骨架沿用 Tokenomics：
- 結構性：頁、名稱、CURRENT、SRC_OAI 筆數；
- 情境：改輸入 → LibreOffice 重算作基準 → 與 engine 比對全部公式格；任何情境不得出現錯誤值（檢查頁 ERR 是字串，不是錯誤值）；
- 效能：增量與全簿重算各 < 2 秒；快取等價。"""
import os
import time
from pathlib import Path

import pytest
import yaml

from engine import Engine, current_model_path, read_defined_names_xml, resolve_key
from parity_lib import compare, excel_values, format_mismatches, lo_recalc, set_inputs

HERE = Path(__file__).parent
CFG = yaml.safe_load((HERE / "scenarios.yaml").read_text(encoding="utf-8"))
SCENARIOS = CFG["scenarios"]
EXPECT = CFG["workbook_expectations"]
FULL_RECALC_LIMIT_S = 2.0
INCREMENTAL_LIMIT_S = 2.0


@pytest.fixture(scope="module")
def model():
    return current_model_path()


def new_engine(model, fresh=False):
    cache = None if fresh else os.environ.get("OPENAI_ENGINE_CACHE")
    return Engine(model, cache=cache or None)


@pytest.fixture(scope="module")
def base_engine(model):
    eng = new_engine(model, fresh=True)
    return eng, eng.evaluate_all()


def test_model_current_pointer(model):
    assert (model.parent / "CURRENT").read_text(encoding="utf-8").strip() == model.name


def test_workbook_expectations(model):
    eng = new_engine(model)
    assert eng.sheetnames == EXPECT["sheets"]
    assert len(eng.formula_cells) == EXPECT["formula_cells"]
    assert len(eng.names) == EXPECT["defined_names"]
    assert sum(n.startswith("SRC_OAI_") for n in eng.names) == EXPECT["src_named"]
    assert sum(n.startswith("INP_") for n in eng.names) == EXPECT["inp_rows"]
    assert sum(n.startswith(("TK_IF_", "TK_SRC_", "TK_L1_")) for n in eng.names) == EXPECT["tk_values"]
    assert eng.get_name("CHK_Errors") == 0


def test_row_counts_match_report(model):
    """SRC_OAI 筆數、Inputs 筆數、TK 待合併數：與報告所列一致（期望值在 scenarios.yaml，報告引用同一數字）。"""
    eng = new_engine(model)
    ids = [r[0] for r in eng.get("SRC_OAI", "A5:B400") if r[0] != ""]
    assert len(ids) == EXPECT["src_rows"] and len(set(ids)) == len(ids)
    inp = [r[0] for r in eng.get("Inputs", "A5:B600") if r[0] != ""]
    assert len(inp) == EXPECT["inp_rows"] and len(set(inp)) == len(inp)
    pending = [r for r in eng.get("TK_Link", "E10:F200") if r[1] == "待 Tokenomics 提供"]
    assert len(pending) == EXPECT["tk_pending"]


def test_named_ranges(model):
    eng = new_engine(model)
    xml_names = read_defined_names_xml(model)
    assert set(eng.names) == set(xml_names)
    for n in eng.names:
        assert eng.names[n] == xml_names[n]
        v = eng.get_name(n)
        flat = v if isinstance(v, list) else [v]
        assert all(not (isinstance(x, str) and x.startswith("#")) for x in flat), f"{n} 含錯誤值"
    for n in ("TK_File", "TK_Version", "TK_Commit", "TK_ReadDate"):
        assert eng.get_name(n) not in ("", None)
    assert len(eng.get_name("TK_HdrGen")) == len(eng.get_name("TK_HdrCost")) == len(eng.get_name("TK_IF_TokGW_Luna")) == 15


def test_named_ranges_vs_libreoffice(model, tmp_path):
    src = tmp_path / model.name
    set_inputs(model, src, {})
    out = lo_recalc(src, tmp_path / "lo")
    import openpyxl
    lo = {k: v.attr_text.replace("$", "") for k, v in openpyxl.load_workbook(out).defined_names.items()}
    ours = {k: v.replace("$", "") for k, v in Engine(model).names.items()}
    assert lo == ours


FRESH_STATE = {}


@pytest.mark.parametrize("sc", SCENARIOS, ids=[s["id"] for s in SCENARIOS])
def test_scenario_parity(sc, model, base_engine, tmp_path, results_store):
    scen_xlsx = tmp_path / model.name
    set_inputs(model, scen_xlsx, sc["inputs"])
    ref = excel_values(lo_recalc(scen_xlsx, tmp_path / "lo"), base_engine[0].formula_cells)
    eng = new_engine(model, fresh=True)
    for key, v in sc["inputs"].items():
        eng.set_key(key, v)
    t0 = time.perf_counter()
    got = eng.evaluate_all()
    elapsed = time.perf_counter() - t0
    FRESH_STATE[sc["id"]] = (got, {n: eng.get_name(n) for n in eng.names})
    res = compare(got, ref)
    changed = [k for k in base_engine[1] if base_engine[1][k] != got[k]]
    results_store[sc["id"]] = {k: v for k, v in res.items() if k != "mismatches"} | {
        "n_mismatch": len(res["mismatches"]), "eval_all_seconds_after_change": round(elapsed, 2), "changed_cells": len(changed)}
    assert elapsed < INCREMENTAL_LIMIT_S
    assert res["error_value_cells"] == 0, f"[{sc['id']}] 錯誤值 {res['error_value_cells']} 格"
    if sc["id"] != "base":
        assert changed, f"[{sc['id']}] 未改變任何公式格（測試空轉）"
    assert not res["mismatches"], f"[{sc['id']}] " + format_mismatches(res["mismatches"])


@pytest.fixture(scope="session")
def cache_path(tmp_path_factory):
    env = os.environ.get("OPENAI_ENGINE_CACHE")
    if env:
        return env
    p = tmp_path_factory.mktemp("cache") / "engine.pkl"
    Engine(current_model_path()).save_cache(p)
    return str(p)


@pytest.mark.parametrize("sc", SCENARIOS, ids=[s["id"] for s in SCENARIOS])
def test_cache_matches_fresh(sc, model, cache_path, results_store):
    if sc["id"] in FRESH_STATE:
        want_cells, want_names = FRESH_STATE[sc["id"]]
    else:
        fresh = new_engine(model, fresh=True)
        for key, v in sc["inputs"].items():
            fresh.set_key(key, v)
        want_cells = fresh.evaluate_all()
        want_names = {n: fresh.get_name(n) for n in fresh.names}
    t0 = time.perf_counter()
    eng = Engine(model, cache=cache_path)
    load_s = time.perf_counter() - t0
    for key, v in sc["inputs"].items():
        eng.set_key(key, v)
    got_cells = eng.evaluate_all()
    bad_cells = [k for k in want_cells if want_cells[k] != got_cells[k] or type(want_cells[k]) is not type(got_cells[k])]
    bad_names = [n for n in want_names if want_names[n] != eng.get_name(n)]
    results_store.setdefault("_cache_equiv", {})[sc["id"]] = {
        "cells": len(want_cells), "names": len(want_names), "cell_mismatch": len(bad_cells),
        "name_mismatch": len(bad_names), "load_seconds": round(load_s, 1)}
    assert not bad_cells and not bad_names


def test_scenarios_actually_change_outputs(model, base_engine):
    base = base_engine[1]
    for sc in SCENARIOS[1:]:
        eng = new_engine(model)
        for key, v in sc["inputs"].items():
            eng.set_key(key, v)
        cur = eng.evaluate_all()
        assert [k for k in base if base[k] != cur[k]], f"{sc['id']} 未改變任何公式格"


def test_scenario_expected_results(model):
    """情境的預期結果：被改的輸入必須讓對應檢查列轉為 ERR（保證檢查有牙齒）。"""
    want = {"a_equity_sum_broken": "C12", "c_util_out_of_range": "C10", "d_oracle_term": "C13", "b_input_out_of_range": "C06", "e_pro_ratio_high": "C22",
            "i_event_order_broken": "C34", "m_usage_mix_broken": "C30", "u_fleet_broken": "C40"}
    eng0 = new_engine(model, fresh=True)
    assert eng0.get_name("CHK_Errors") == 0
    for sc in SCENARIOS[1:]:
        eng = new_engine(model, fresh=True)
        for k, v in sc["inputs"].items():
            eng.set_key(k, v)
        errs = {eng.get("Checks", f"A{r}") for r in range(5, 120) if eng.get("Checks", f"E{r}") == "ERR"}
        if want.get(sc["id"]):
            assert want[sc["id"]] in errs, (sc["id"], errs)
        assert eng.get_name("CHK_Errors") == len(errs)


def test_incremental_recalc_matches_fresh_and_is_fast(model):
    eng = new_engine(model, fresh=True)
    base = eng.evaluate_all()
    names = eng.names
    orig = {a: eng.get(*resolve_key(names, a)) for sc in SCENARIOS[1:] for a in sc["inputs"]}
    for sc in SCENARIOS[1:]:
        t0 = time.perf_counter()
        for key, v in sc["inputs"].items():
            eng.set_key(key, v)
        eng.evaluate_all()
        assert time.perf_counter() - t0 < INCREMENTAL_LIMIT_S
        for key in sc["inputs"]:
            eng.set_key(key, orig[key])
    res = compare(eng.evaluate_all(), base)
    assert not res["mismatches"], format_mismatches(res["mismatches"])


def test_full_recalc_time(model, results_store):
    eng = new_engine(model, fresh=True)
    eng.evaluate_all()
    t0 = time.perf_counter()
    eng._xl.recalculate()
    full = time.perf_counter() - t0
    results_store["_full_recalc_seconds"] = round(full, 2)
    assert full < FULL_RECALC_LIMIT_S

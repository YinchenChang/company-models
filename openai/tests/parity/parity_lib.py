"""Parity 測試工具：LibreOffice headless 重算（基準）與比對。"""
from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

import openpyxl

from engine import resolve_key

REL_TOL = 1e-9
ABS_TOL = 1e-12


def set_inputs(src: Path, dst: Path, inputs: dict[str, object]) -> None:
    """複製活頁簿並改寫輸入格。openpyxl 存檔不帶快取值，LibreOffice 載入時必須重算。"""
    wb = openpyxl.load_workbook(src)
    names = {k: v.attr_text for k, v in wb.defined_names.items()}
    for key, v in inputs.items():          # 鍵：'Sheet!A1'、'NAME' 或 'NAME[k]'（具名範圍優先）
        sheet, coord = resolve_key(names, key)
        cell = wb[sheet][coord]
        if isinstance(cell.value, str) and cell.value.startswith("="):
            raise ValueError(f"{key} 是公式格，不是輸入格")
        cell.value = v
    wb.save(dst)


def lo_recalc(xlsx: Path, outdir: Path) -> Path:
    """以 LibreOffice headless 重算並另存 xlsx，回傳輸出路徑。"""
    outdir.mkdir(parents=True, exist_ok=True)
    profile = Path(tempfile.mkdtemp(prefix="lo_profile_"))
    try:
        cmd = [
            "soffice", f"-env:UserInstallation=file://{profile}",
            "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(outdir), str(xlsx),
        ]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    out = outdir / xlsx.name
    if not out.exists():
        raise RuntimeError(f"LibreOffice 重算失敗：{r.stdout}\n{r.stderr}")
    return out


def excel_values(recalculated: Path, cells: list[tuple[str, str]]) -> dict[tuple[str, str], object]:
    wb = openpyxl.load_workbook(recalculated, data_only=True)
    return {(s, c): ("" if wb[s][c].value is None else wb[s][c].value) for s, c in cells}


def _is_num(x) -> bool:
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def compare(engine_vals: dict, excel_vals: dict) -> dict:
    """比對；數值：相對誤差 ≤1e-9 或絕對誤差 ≤1e-12；其他型別必須完全相等。"""
    mismatches, max_rel, max_rel_all, max_abs, n_num, n_err = [], 0.0, 0.0, 0.0, 0, 0
    for key, e in engine_vals.items():
        x = excel_vals[key]
        if isinstance(e, str) and e.startswith("#") or isinstance(x, str) and x.startswith("#"):
            n_err += 1
        if _is_num(e) and _is_num(x):
            n_num += 1
            d = abs(e - x)
            scale = max(abs(e), abs(x))
            rel = d / scale if scale else 0.0
            max_rel_all = max(max_rel_all, rel)  # 全部數值格（含絕對誤差極小者）
            if d > ABS_TOL:  # 絕對誤差已在容差內者不計入 max_rel_err
                max_rel = max(max_rel, rel)
            max_abs = max(max_abs, d)
            ok = d <= ABS_TOL or rel <= REL_TOL
        else:
            ok = type(e) is type(x) and e == x
        if not ok:
            mismatches.append((key[0], key[1], e, x))
    return {
        "cells": len(engine_vals), "numeric_cells": n_num, "mismatches": mismatches,
        "max_rel_err": max_rel, "max_rel_err_all_cells": max_rel_all, "max_abs_err": max_abs, "error_value_cells": n_err,
    }


def format_mismatches(mm: list, limit: int = 20) -> str:
    lines = [f"不符 {len(mm)} 格；前 {min(limit, len(mm))} 個：", "頁名 | 座標 | 引擎值 | Excel 值"]
    lines += [f"{s} | {c} | {e!r} | {x!r}" for s, c, e, x in mm[:limit]]
    return "\n".join(lines)

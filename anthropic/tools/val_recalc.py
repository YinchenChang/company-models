#!/usr/bin/env python3
"""估值頁（Val）建置快照的共用工具：改 Inputs（藍字）→ LibreOffice 重算整本活頁簿 → 讀值。

只執行與讀取 Excel（CLAUDE.md 第 1 節：Python 不另寫計算）。tools/val_snapshots.py（情境 A／C）與 tools/val_sens.py（敏感度表）共用。
"""
from __future__ import annotations

import shutil
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "builder"))
from common import lo_recalc  # noqa: E402


def inputs_rows(wb):
    ws = wb["Inputs"]
    return {ws.cell(r, 14).value: r for r in range(5, ws.max_row + 1) if ws.cell(r, 14).value}


def recalc_with(model: Path, overrides: dict, tag: str) -> Path:
    """overrides：{Inputs key: 值 或 'lo'／'hi'（取該列低／高欄）}。回傳重算後的 xlsx 路徑（暫存目錄）。"""
    wb = openpyxl.load_workbook(model)
    rows = inputs_rows(wb)
    ws = wb["Inputs"]
    for k, v in overrides.items():
        r = rows[k]
        if v == "lo":
            v = ws[f"F{r}"].value
        elif v == "hi":
            v = ws[f"G{r}"].value
        ws[f"E{r}"].value = v
    d = Path(tempfile.mkdtemp(prefix=f"val_{tag}_"))
    raw = d / f"{tag}.xlsx"
    wb.save(raw)
    out = lo_recalc(raw, d / "out")
    return out


def read_names(path: Path, names):
    wb = openpyxl.load_workbook(path, data_only=True)
    res = {}
    for n in names:
        sh, ref = wb.defined_names[n].attr_text.split("!")
        c = wb[sh.strip("'")][ref.replace("$", "")]
        res[n] = [x.value for row in c for x in row] if isinstance(c, tuple) else c.value
    return res, wb


def run_many(model: Path, jobs: dict, reader, workers: int = 4):
    """jobs：{tag: overrides}；reader(path) → 結果。平行跑 LibreOffice（各自獨立設定檔）。"""
    def one(item):
        tag, ov = item
        p = recalc_with(model, ov, tag)
        try:
            return tag, reader(p)
        finally:
            shutil.rmtree(p.parent.parent, ignore_errors=True)
    with ThreadPoolExecutor(workers) as ex:
        return dict(ex.map(one, jobs.items()))

#!/usr/bin/env python3
"""兩份活頁簿逐格比對（範圍外變動核對；CLAUDE.md 第 8 節第 7 步）。

比對方法：以 openpyxl 讀兩份 xlsx 的「LibreOffice 重算後快取值」（data_only=True），逐頁、逐格（A1 到兩份檔的最大列與最大欄的聯集）比較；
數值以相對誤差 ≤1e-9（絕對 ≤1e-12）視為相同，文字須完全相同。可選擇只比數值格（--numeric-only）或排除某些頁（--skip）。
另可用 --ignore 指定「預期會變動」的格（頁!座標，或 頁!列號 表整列）；輸出：比對格數、不同格數、預期內／範圍外的明細。
本程式只讀取與比較，不做任何模型計算。
--existing-only：只比對 A 有值的格（B 新增列只計數）。
用法：python3 tools/diff_xlsx.py A.xlsx B.xlsx [--skip _Defaults,README] [--numeric-only] [--existing-only] [--ignore-file f.txt] [--json out.json] [--max 50]
"""
from __future__ import annotations

import argparse
import json
import math
from pathlib import Path

import openpyxl


def same(a, b):
    if a is None or a == "":
        return b is None or b == ""
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return math.isclose(float(a), float(b), rel_tol=1e-9, abs_tol=1e-12)
    return a == b


def diff(pa: Path, pb: Path, skip=(), numeric_only=False, ignore=(), existing_only=False):
    wa = openpyxl.load_workbook(pa, data_only=True)
    wb = openpyxl.load_workbook(pb, data_only=True)
    ign_cells = {x for x in ignore if not x.split("!")[1].isdigit()}
    ign_rows = {(x.split("!")[0], int(x.split("!")[1])) for x in ignore if x.split("!")[1].isdigit()}
    out = dict(sheets_only_a=[s for s in wa.sheetnames if s not in wb.sheetnames], sheets_only_b=[s for s in wb.sheetnames if s not in wa.sheetnames],
               compared=0, numeric_compared=0, diff=[], ignored=[], added=0)
    for s in wa.sheetnames:
        if s in skip or s not in wb.sheetnames:
            continue
        A, B = wa[s], wb[s]
        nr, nc = max(A.max_row, B.max_row), max(A.max_column, B.max_column)
        ra = {(c.row, c.column): c.value for row in A.iter_rows(min_row=1, max_row=nr, max_col=nc) for c in row}
        rb = {(c.row, c.column): c.value for row in B.iter_rows(min_row=1, max_row=nr, max_col=nc) for c in row}
        for k in set(ra) | set(rb):
            va, vb = ra.get(k), rb.get(k)
            isnum = isinstance(va, (int, float)) or isinstance(vb, (int, float))
            if numeric_only and not isnum:
                continue
            if va is None and vb is None:
                continue
            if existing_only and (va is None or va == ""):
                out["added"] += 1          # 只在 B 有值的格（新增列）：不比對，只計數
                continue
            out["compared"] += 1
            out["numeric_compared"] += int(isnum)
            if not same(va, vb):
                coord = f"{s}!{openpyxl.utils.get_column_letter(k[1])}{k[0]}"
                rec = (coord, va, vb)
                (out["ignored"] if (coord in ign_cells or (s, k[0]) in ign_rows) else out["diff"]).append(rec)
    out["diff"].sort()
    out["ignored"].sort()
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("a", type=Path)
    ap.add_argument("b", type=Path)
    ap.add_argument("--skip", default="_Defaults")
    ap.add_argument("--numeric-only", action="store_true")
    ap.add_argument("--ignore-file", type=Path)
    ap.add_argument("--json", type=Path)
    ap.add_argument("--max", type=int, default=50)
    ap.add_argument("--existing-only", action="store_true", help="只比對 A 有值的格；B 新增的格只計數（新列附加在頁尾時用）")
    a = ap.parse_args()
    ign = [x.strip() for x in a.ignore_file.read_text(encoding="utf-8").splitlines() if x.strip() and not x.startswith("#")] if a.ignore_file else []
    r = diff(a.a, a.b, skip=tuple(x for x in a.skip.split(",") if x), numeric_only=a.numeric_only, ignore=ign, existing_only=a.existing_only)
    print(f"只在 A 的頁：{r['sheets_only_a']}；只在 B 的頁：{r['sheets_only_b']}")
    print(f"比對格數 {r['compared']}（其中數值 {r['numeric_compared']}）；預期內變動 {len(r['ignored'])}；範圍外（未預期）變動 {len(r['diff'])}；B 新增格 {r['added']}")
    for x in r["diff"][: a.max]:
        print("  DIFF", *x)
    if a.json:
        a.json.write_text(json.dumps(r, ensure_ascii=False, indent=1, default=str), encoding="utf-8")


if __name__ == "__main__":
    main()

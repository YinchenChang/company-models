#!/usr/bin/env python3
"""xlsx 逐格比對與錯誤格檢查（本機與 CI 共用）。

用法：
  python3 tools/xlsx_diff.py 新.xlsx 基準.xlsx [--tol 1e-9] [--limit 20]
      逐工作表、逐格比對「快取值」（openpyxl data_only=True）：
      數值以 |a-b| <= tol × max(1, |a|, |b|) 判定相同；其他型別須完全相等；
      工作表名單、儲存格範圍內任一格不同皆算差異。同時檢查兩檔的錯誤格數。
  python3 tools/xlsx_diff.py 新.xlsx --errors-only
      只檢查錯誤格（#DIV/0!、#NAME?、#REF!、#VALUE!、#N/A、#NUM!、#NULL!、Err:xxx）。

結束碼：0＝無差異且錯誤格為 0；1＝有差異或錯誤格；2＝用法或讀檔錯誤。
輸出：每個檔案的錯誤格數、差異總數，以及前 --limit 個差異（預設 20）。
注意：兩個檔案都須含快取值（openpyxl 直接寫出的檔案沒有，須先以 LibreOffice 重算存檔）。
"""
import argparse
import sys

import openpyxl

ERR_PREFIX = ('#DIV/0!', '#NAME?', '#REF!', '#VALUE!', '#N/A', '#NUM!', '#NULL!', 'Err:')


def is_error(v):
    return isinstance(v, str) and v.startswith(ERR_PREFIX)


def load(path):
    try:
        return openpyxl.load_workbook(path, data_only=True)
    except Exception as e:  # 檔案不存在或不是 xlsx
        print(f'讀檔失敗：{path}（{e}）', file=sys.stderr)
        sys.exit(2)


def cells(wb):
    """回傳 {(工作表, 座標): 值}，不含空格。"""
    out = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    out[(ws.title, c.coordinate)] = c.value
    return out


def same(a, b, tol):
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= tol * max(1.0, abs(a), abs(b))
    return a == b


def error_cells(cs):
    return [(k, v) for k, v in cs.items() if is_error(v)]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('new')
    ap.add_argument('ref', nargs='?')
    ap.add_argument('--tol', type=float, default=1e-9)
    ap.add_argument('--limit', type=int, default=20)
    ap.add_argument('--errors-only', action='store_true')
    a = ap.parse_args()
    if not a.errors_only and not a.ref:
        ap.error('需要基準檔，或加 --errors-only')

    wb_new = load(a.new)
    cs_new = cells(wb_new)
    errs = error_cells(cs_new)
    bad = bool(errs)
    print(f'{a.new}：{len(wb_new.worksheets)} 張工作表、{len(cs_new)} 格、錯誤格 {len(errs)}')
    for (sh, co), v in errs[:a.limit]:
        print(f'  錯誤格 {sh}!{co} = {v}')
    if a.errors_only:
        return 1 if bad else 0

    wb_ref = load(a.ref)
    cs_ref = cells(wb_ref)
    ref_errs = error_cells(cs_ref)
    print(f'{a.ref}：{len(wb_ref.worksheets)} 張工作表、{len(cs_ref)} 格、錯誤格 {len(ref_errs)}')
    diffs = []
    s_new, s_ref = [w.title for w in wb_new.worksheets], [w.title for w in wb_ref.worksheets]
    if s_new != s_ref:
        diffs.append(f'工作表名單／順序不同：新 {s_new} ≠ 基準 {s_ref}')
    for k in sorted(set(cs_new) | set(cs_ref), key=lambda k: (s_ref.index(k[0]) if k[0] in s_ref else 99, k[0], k[1])):
        n, r = cs_new.get(k), cs_ref.get(k)
        if k not in cs_new:
            diffs.append(f'{k[0]}!{k[1]}：新檔為空，基準 {r!r}')
        elif k not in cs_ref:
            diffs.append(f'{k[0]}!{k[1]}：基準為空，新檔 {n!r}')
        elif not same(n, r, a.tol):
            diffs.append(f'{k[0]}!{k[1]}：新 {n!r} ≠ 基準 {r!r}')
    print(f'差異 {len(diffs)} 格（容差 {a.tol:g}）')
    for d in diffs[:a.limit]:
        print('  ', d)
    if len(diffs) > a.limit:
        print(f'   … 另有 {len(diffs) - a.limit} 項未列出')
    return 1 if (diffs or bad) else 0


if __name__ == '__main__':
    sys.exit(main())

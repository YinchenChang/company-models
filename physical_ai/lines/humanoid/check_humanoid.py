# -*- coding: utf-8 -*-
"""核對程式：LibreOffice 重算 → 0 錯誤值、核對頁全部 PASS、印出主要結果。

python3 check_humanoid.py [xlsx]   （預設 dist/ 下最新檔）
"""
import glob
import os
import shutil
import subprocess
import sys
import tempfile

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
ERRS = ("#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "Err:")


def recalc(src):
    tmp = tempfile.mkdtemp()
    dst = os.path.join(tmp, "in.xlsx")
    shutil.copy(src, dst)
    out = os.path.join(tmp, "out")
    os.makedirs(out)
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", out, dst],
                   check=True, capture_output=True, timeout=300)
    return os.path.join(out, "in.xlsx")


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))[-1]
    wb = load_workbook(recalc(src), data_only=True)
    bad = []
    nform = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and any(c.value.startswith(e) for e in ERRS):
                    bad.append(f"{ws.title}!{c.coordinate}={c.value}")
    wf = load_workbook(src)
    for ws in wf.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    nform += 1
    k = wb["核對"]
    fails = [(k.cell(r, 2).value, k.cell(r, 3).value, k.cell(r, 4).value)
             for r in range(4, k.max_row + 1) if k.cell(r, 5).value == "FAIL"]
    print(f"檔案：{src}")
    print(f"公式數：{nform}；錯誤值：{len(bad)}；核對 FAIL：{len(fails)}")
    for b in bad[:30]:
        print("  ERR", b)
    for f in fails:
        print("  FAIL", f)
    d = wb["下游取數"]
    print("\n下游取數（代號｜低｜基準｜高｜單位）")
    for r in range(5, d.max_row + 1):
        if d.cell(r, 1).value and d.cell(r, 4).value is not None:
            def f(v):
                return f"{v:,.2f}" if isinstance(v, (int, float)) else str(v)
            print(f"  {d.cell(r,1).value}｜{f(d.cell(r,3).value)}｜{f(d.cell(r,4).value)}｜{f(d.cell(r,5).value)}｜{d.cell(r,6).value}")
    s = wb["敏感度"]
    print("\n敏感度排序（前 10）")
    for r in range(6, 16):
        print("  ", [s.cell(r, j).value if not isinstance(s.cell(r, j).value, float) else round(s.cell(r, j).value, 2) for j in range(1, 9)])
    print("\n基準與絲杠情形")
    for r in range(1, s.max_row + 1):
        if s.cell(r, 1).value == "C. 基準結果與兩個特別表":
            for rr in range(r + 1, r + 10):
                print("  ", [s.cell(rr, j).value if not isinstance(s.cell(rr, j).value, float) else round(s.cell(rr, j).value, 3) for j in range(1, 9)])
    c = wb["計算"]
    print("\n計算頁")
    for r in range(1, c.max_row + 1):
        v = c.cell(r, 2).value
        if c.cell(r, 1).value and v is not None and not isinstance(v, str) or (isinstance(v, str) and v in ("已跨過", "未跨過")):
            print(f"  {c.cell(r,1).value}: {v if not isinstance(v,float) else round(v,3)} | {c.cell(r,3).value} | {c.cell(r,4).value}")
    sys.exit(1 if (bad or fails) else 0)


if __name__ == "__main__":
    main()

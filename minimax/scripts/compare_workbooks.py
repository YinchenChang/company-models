# 逐格比對兩份 xlsx（值與公式）。用法：python3 scripts/compare_workbooks.py 舊.xlsx 新.xlsx [--sheet 頁名 ...略過頁]
# 以「舊檔非空格」為基準：列出不同格、統計新增格（舊檔空、新檔有）。結束碼恆為 0。
import sys
from openpyxl import load_workbook
a, b = sys.argv[1:3]
skip = set(sys.argv[4:]) if len(sys.argv) > 3 and sys.argv[3] == '--skip' else set()
av, bv = load_workbook(a, data_only=True), load_workbook(b, data_only=True)
af, bf = load_workbook(a), load_workbook(b)
n = d = new = 0
for ws in av.worksheets:
    if ws.title in skip or ws.title not in bv.sheetnames: continue
    for row in ws.iter_rows():
        for c in row:
            if c.value is None: continue
            n += 1
            k = c.coordinate
            if bv[ws.title][k].value != c.value or bf[ws.title][k].value != af[ws.title][k].value:
                d += 1; print('DIFF', ws.title, k, str(c.value)[:60], '|', str(bv[ws.title][k].value)[:60])
    for row in bv[ws.title].iter_rows():
        for c in row:
            if c.value is not None and ws[c.coordinate].value is None: new += 1
print(f'舊檔非空格 {n}｜不同 {d}｜新增格 {new}｜頁序相同 {av.sheetnames == bv.sheetnames}')

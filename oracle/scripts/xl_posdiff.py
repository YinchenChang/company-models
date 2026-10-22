# 一次性「對位比對」（Oracle v0.3.1 A 段步驟 1）：兩份 Excel 依工作表＋列＋欄的位置逐格比對「值」與「公式」，
# 只允許 --allow 指定的儲存格不同（例如改名的表頭格）。用於 A 欄改名使 xl_diff --by-label 無法以列名稱配對時。
# 用法：python3 scripts/xl_posdiff.py 舊.xlsx 新.xlsx [--allow=工作表!儲存格,...]   結果：第一行為「0 differences」才算通過。
import sys, openpyxl
a, b = sys.argv[1], sys.argv[2]
allow = {x for o in sys.argv if o.startswith('--allow=') for x in o.split('=', 1)[1].split(',') if x}
def norm(v):
    return v.__class__.__name__ + str(sorted((k, str(x)) for k, x in vars(v).items())) if hasattr(v, 'ref') and not isinstance(v, (int, float)) else v
def same(x, y): return x == y or (isinstance(x, float) and isinstance(y, float) and abs(x - y) < 1e-9)
diffs, allowed, n = [], [], 0
for mode in (False, True):  # False＝公式，True＝值
    wa, wb = openpyxl.load_workbook(a, data_only=mode), openpyxl.load_workbook(b, data_only=mode)
    if wa.sheetnames != wb.sheetnames: diffs.append(('sheets', wa.sheetnames, wb.sheetnames))
    for s in wa.sheetnames:
        A, B = wa[s], wb[s]
        if (A.max_row, A.max_column) != (B.max_row, B.max_column): diffs.append((s, 'size', (A.max_row, A.max_column), (B.max_row, B.max_column)))
        for r in range(1, max(A.max_row, B.max_row) + 1):
            for c in range(1, max(A.max_column, B.max_column) + 1):
                x, y = norm(A.cell(r, c).value), norm(B.cell(r, c).value); n += 1
                if not same(x, y):
                    k = f'{s}!{A.cell(r, c).coordinate}'
                    (allowed if k in allow else diffs).append((('值' if mode else '公式'), k, x, y))
print(f'{len(diffs)} differences' if diffs else '0 differences')
print(f'比對 {n:,} 格（公式與值各一輪）；允許的差異 {len(allowed)} 格：')
for d in allowed: print('  允許', d)
for d in diffs[:40]: print('  差異', d)
sys.exit(1 if diffs else 0)

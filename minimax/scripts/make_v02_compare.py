# 產生 v0.1 → v0.2 對照表（①變動列 ②關鍵輸出前後值 ③已套用的預設）。
# 用法：python3 scripts/make_v02_compare.py v0.1.xlsx M1基準.xlsx v0.2資料更新後_I31=0.xlsx v0.2成品.xlsx 輸出.xlsx
#   M1 基準＝M1 完成時的程式（commit 40ce39a）建置；I31=0＝以現行程式 `build_xlsx.py --set I31=0` 建置（僅含資料更新、不含 M3.1 降價）。
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from xlvals import load, val, Y

v01, m1, v02a, v02, out = sys.argv[1:6]
D = [load(p) for p in (v01, m1, v02a, v02)]
HDR = PatternFill('solid', fgColor='DDE6F0')
wb = Workbook()

# ① 變動列：以 ID 為鍵比對 M1 基準與 v0.2 成品的公式；列出新增與公式改變
ws = wb.active; ws.title = '①變動列'
for c, h in enumerate(['分類', 'ID／儲存格', '項目', 'M1 基準', 'v0.2', '說明'], 1):
    x = ws.cell(1, c, h); x.font = Font(bold=True); x.fill = HDR
f1, f2 = load_workbook(m1), load_workbook(v02)
rows = []
for sh in ('Demand', 'Revenue', 'Compute', 'PnL', 'PxQ', 'Funding', 'Valuation'):
    a = {}
    if sh in f1.sheetnames:
        for r in f1[sh].iter_rows(min_row=6):
            if r[0].value and len(str(r[0].value)) <= 4: a[r[0].value] = r
    for r in f2[sh].iter_rows(min_row=6):
        k = r[0].value
        if not (k and len(str(k)) <= 4): continue
        if k not in a:
            rows.append(('新增列', f'{sh}!{k}', r[1].value, '—', '（新列）', ''))
        else:
            for i, y in enumerate(Y):
                fa, fb = a[k][3 + i].value, r[3 + i].value
                if fa != fb:
                    rows.append(('公式改變', f'{sh}!{k}@{y}', r[1].value, str(fa)[:80], str(fb)[:80], ''))
for sh, lab in (('SRC_MM', 'SRC_MM'), ('Inputs', 'Inputs')):
    a = {r[0].value: r for r in f1[sh].iter_rows(min_row=5) if r[0].value}
    for r in f2[sh].iter_rows(min_row=5):
        k = r[0].value
        if not k: continue
        if k not in a: rows.append(('新增列', f'{sh}!{k}', r[1].value, '—', str(r[4].value if sh == 'SRC_MM' else r[3].value), ''))
        elif str(a[k][1].value) != str(r[1].value): rows.append(('標註改變', f'{sh}!{k}', r[1].value, str(a[k][1].value)[:60], str(r[1].value)[:60], '舊列保留並標「已更新」'))
for rr in rows: ws.append(list(rr))
for col, w in zip('ABCDEF', [10, 18, 60, 40, 40, 30]): ws.column_dimensions[col].width = w

# ② 關鍵輸出前後值
ws = wb.create_sheet('②關鍵輸出前後值')
heads = ['輸出', '年', 'v0.1（2026-10-08）', 'M1（TK 同步 v19.0）', 'v0.2 資料更新（不含 M3.1 降價；I31=0）', 'v0.2 成品（基準，含 M3.1 已觀測降價）', '新機制（PxQ／D2／η 對照）對此值的影響']
for c, h in enumerate(heads, 1):
    x = ws.cell(1, c, h); x.font = Font(bold=True); x.fill = HDR
KEY = [('C11', 2025, 'η 2025'), ('C11', 2026, 'η 1H26'), ('D01', 2026, 'API 全年平均單價 $/M'), ('R09', 2026, '營收 2026 $M'), ('R09', 2030, '營收 2030 $M'),
       ('C13', 2026, 'API 推論 MW 2026'), ('C13', 2030, 'API 推論 MW 2030'), ('C21', 2026, '總 MW 2026'), ('C21', 2030, '總 MW 2030'),
       ('P21', 2030, '每 MW 營收 2030 $M'), ('P22', 2030, '每 MW 全成本 2030 $M'), ('P23', 2027, '覆蓋率 2027'), ('P23', 2030, '覆蓋率 2030'),
       ('F11', 2030, '累計需新融資 2030 $M'), ('F19', 2026, '首次需新融資年'), ('V02', 2026, '市值 $M'), ('F14', 2030, '累計稀釋（占市值）'), ('V13', 2030, '反向檢驗 折現價值÷市值'),
       ('X05', 2025, '回本比值 2025（新）'), ('X05', 2026, '回本比值 2026（新）'), ('X05', 2030, '回本比值 2030（新）'), ('X28', 2030, '比值＝1 所需合約價 $M（新）')]
for k, y, lab in KEY:
    def g(d):
        try: return val(d, k, y)
        except Exception: return None
    r = [g(d) for d in D]
    note = '無（新輸出）' if k.startswith('X') else ('0' if (r[3] is not None and r[2] is not None and abs(r[3] - r[2]) < 1e-9) else '見資料更新→成品差')
    ws.append([lab, y, r[0], r[1], r[2], r[3], '無（新輸出）' if k.startswith('X') else '0（新機制不改既有數值）'])
for col, w in zip('ABCDEFG', [30, 8, 20, 20, 30, 30, 30]): ws.column_dimensions[col].width = w

# ③ 已套用的預設
ws = wb.create_sheet('③已套用的預設')
for c, h in enumerate(['#', '問題', '預設', '替代', '影響方向／量化'], 1):
    x = ws.cell(1, c, h); x.font = Font(bold=True); x.fill = HDR
import json
for row in json.load(open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'docs', 'reports', '20261022_minimax_v0.2_defaults.json'), encoding='utf-8')):
    ws.append(row)
for col, w in zip('ABCDE', [5, 40, 60, 40, 70]): ws.column_dimensions[col].width = w
wb.save(out); print(out)

# Tokenomics 同步歸因：以舊快照為底，一次只換一組名稱（其餘維持舊值），各自建置、LibreOffice 重算後讀關鍵輸出。
# 用法：python3 scripts/tk_attribution.py 舊快照.json 新快照.json   → out/attribution.json
import sys, os, json, subprocess
from openpyxl import load_workbook
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from recalc import recalc

GROUPS = [('①產能 IF_TokGW_*', ['IF_TokGW_Luna', 'IF_TokGW_Sol']), ('②利用率 IF_Util', ['IF_Util']),
          ('③持有成本與全成本 IF_HoldEcon／IF_FullCost_*／IF_PriceRef_*', ['IF_HoldEcon', 'IF_FullCost_Luna', 'IF_FullCost_Sol', 'IF_PriceRef_Luna', 'IF_PriceRef_Sol'])]
# (列 ID, 年, 標籤)；年＝None 表示取整列
OUT = [('C12', y, f'η {y}') for y in (2025, 2026)] + [('C13', y, f'API 推論 MW {y}') for y in range(2025, 2031)] + \
      [('C21', y, f'總 MW {y}') for y in range(2025, 2031)] + [('P21', y, f'每 MW 營收 {y}') for y in range(2025, 2031)] + \
      [('P22', y, f'每 MW 全成本 {y}') for y in range(2025, 2031)] + [('P23', y, f'覆蓋率 {y}') for y in range(2025, 2031)] + \
      [('F11', 2030, '累計需新融資 2030'), ('V13', 2030, '折現價值÷市值'), ('C23', 2030, 'VR 等值係數 2030'), ('C29', 2030, 'TK 持有成本 Hopper'), ('C30', 2030, 'API 每 M token 成本 2030')]


def read(path):
    wb = load_workbook(path, data_only=True)
    loc = {}
    for ws in wb.worksheets:
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if isinstance(v, str): loc.setdefault(v, (ws, r))
    res = {}
    for rid, y, lab in OUT:
        ws, r = loc[rid]
        res[lab] = ws.cell(r, 4 + y - 2025).value
    res['Checks 合計'] = wb['Checks'].cell([r for r in range(1, 30) if wb['Checks'].cell(r, 1).value == 'CHK_Sum'][0], 3).value
    return res


def run(tag, snap):
    d = os.path.join(ROOT, 'out', 'attr'); os.makedirs(d, exist_ok=True)
    sp = os.path.join(d, f'{tag}.json'); json.dump(snap, open(sp, 'w'), ensure_ascii=False)
    x = os.path.join(d, f'{tag}.xlsx')
    subprocess.run(['python3', os.path.join(ROOT, 'build_xlsx.py'), x, '--snapshot', sp], check=True, capture_output=True)
    r = recalc(x)
    assert r['status'] == 'success', r
    return read(x)


old, new = (json.load(open(p, encoding='utf-8')) for p in sys.argv[1:3])
res = {'舊（基準）': run('base', old)}
for i, (name, names) in enumerate(GROUPS):
    s = json.loads(json.dumps(old))
    for n in names: s['names'][n] = new['names'][n]
    res[name] = run(f'g{i}', s)
res['全部換新'] = run('all', new)
json.dump(res, open(os.path.join(ROOT, 'out', 'attribution.json'), 'w'), ensure_ascii=False, indent=1)
print('done')

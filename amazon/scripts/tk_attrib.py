# MAG v0.2.1：Tokenomics 同步的目標價變動拆解（工作單 step 5）。數值一律取自 Excel（LibreOffice 開啟重算；三情境以「情境選擇」切換）。
# 四步依序累加，加總＝總變動：
#   ① 換快照（k 固定為舊值、C10 速度舊值）　② k 隨成本重算　③ C10 重解　④ 上限檢查（新增 P×Q 第五列等；應為 0）
# 起點（v0.2）＝舊成品 xlsx；終點＝本 repo 目前資料。步驟用暫存副本的 company.json 變體建 Excel（不動 repo 檔案）。
# 用法：python3 scripts/tk_attrib.py --old dist/舊版.xlsx --k-old low,base,high:... （見下）--json out.json
#   --old       舊版成品 .xlsx
#   --kold      舊 k 的 JSON：{"kLong":{"low":..,"base":..,"high":..},"kSpot":{...}}
#   --pace-old  舊 C10 速度（MW／年，基準情境）
import sys, os, json, shutil, argparse, tempfile, subprocess
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT); sys.path.insert(0, os.path.join(ROOT, 'scripts'))
from uno_q import Workbook
IN, VA, AI = '輸入與假設', '評價_DCF與目標價', 'AI增量報酬'
ap = argparse.ArgumentParser(); ap.add_argument('--old', required=True); ap.add_argument('--kold', required=True); ap.add_argument('--pace-old', type=float, required=True); ap.add_argument('--json')
a = ap.parse_args()


def variant(tmp, name, edit):
    d = os.path.join(tmp, name)
    shutil.copytree(ROOT, d, ignore=shutil.ignore_patterns('out', 'dist', '.git', '__pycache__', 'node_modules'))
    p = os.path.join(d, 'company.json'); co = json.load(open(p, encoding='utf-8')); edit(co, d)
    json.dump(co, open(p, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    x = os.path.join(d, 'step.xlsx')
    r = subprocess.run(['python3', 'build_xlsx.py', x], cwd=d, capture_output=True, text=True, env=dict(os.environ, LC_ALL='C.UTF-8', LANG='C.UTF-8'))
    if r.returncode: sys.exit(f'建 Excel 失敗（{name}）：\n{(r.stdout + r.stderr)[-1500:]}')
    return x


def read(x):
    out = {}
    with Workbook(x) as wb:
        sel = wb.cell(IN, '情境選擇（', prefix=True); s0 = wb.get(sel)
        t1, t2 = wb.cell(VA, '加權目標價'), wb.cell(VA, '讀法 2｜加權目標價')
        for s in (1, 2, 3):
            wb.set(sel, s)
            try: roic = wb.get(wb.cell(AI, '對外 AI ROIC（主值）', col='F')); be = wb.get(wb.cell(AI, '打平 k（', col='C', prefix=True)); k = wb.get(wb.cell(AI, '目前採用 k'))
            except KeyError: roic = be = k = None
            out[s] = {'t1': wb.get(t1), 't2': wb.get(t2), 'roic': roic, 'be': be, 'k': k}
        wb.set(sel, s0)
    return out


def strip_pq(co, d):  # 移除 P×Q 名稱（快照與清單），Excel 即無第五列
    pq = [n for n in co['tokenomics']['names'] if n.startswith('L1_RevGW_PQ_')]
    for key in ('names', 'optional'): co['tokenomics'][key] = [n for n in co['tokenomics'][key] if n not in pq]
    sp = os.path.join(d, co['tokenomics']['snapshotFile']); sn = json.load(open(sp, encoding='utf-8'))
    for n in pq: sn['items'].pop(n, None)
    sn['source']['optional'] = [n for n in sn['source'].get('optional', []) if n not in pq]
    json.dump(sn, open(sp, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)


def set_pace(co, v):
    co['scenarios']['mwPath']['pace']['base'] = v; co['scenarios']['mw31']['base'] = v; co['defaults']['mw31'] = v


def fix_k(co):
    ko = json.loads(a.kold)
    for sd in ('kLong', 'kSpot'): co['pricing'][sd] = ko[sd]


with tempfile.TemporaryDirectory() as tmp:
    final = json.load(open(os.path.join(ROOT, 'company.json'), encoding='utf-8'))['scenarios']['mwPath']['pace']['base']
    steps = [('v0.2', None),
             ('① 換快照（k、C10 固定）', lambda: variant(tmp, 's1', lambda co, d: (fix_k(co), set_pace(co, a.pace_old), strip_pq(co, d)))),
             ('② k 隨成本重算', lambda: variant(tmp, 's2', lambda co, d: (set_pace(co, a.pace_old), strip_pq(co, d)))),
             ('③ C10 重解', lambda: variant(tmp, 's3', lambda co, d: strip_pq(co, d))),
             ('④ 上限檢查層（新增第五列）', lambda: variant(tmp, 's4', lambda co, d: None))]
    res = []
    for nm, f in steps:
        x = a.old if f is None else f()
        res.append((nm, read(x))); print('讀取', nm, flush=True)
    rows = []
    for s, sn in ((1, '保守'), (2, '基準'), (3, '積極')):
        for key, rd in (('t1', '讀法 1'), ('t2', '讀法 2')):
            vals = [r[1][s][key] for r in res]
            d = [vals[i + 1] - vals[i] for i in range(4)]
            rows.append({'scenario': sn, 'reading': rd, 'v02': vals[0], 'steps': d, 'new': vals[4], 'total': vals[4] - vals[0], 'residual': (vals[4] - vals[0]) - sum(d)})
    meta = {sn: {'old': {k: res[0][1][s][k] for k in ('roic', 'be', 'k')}, 'new': {k: res[4][1][s][k] for k in ('roic', 'be', 'k')}} for s, sn in ((1, '保守'), (2, '基準'), (3, '積極'))}
    out = {'steps': [r[0] for r in res], 'rows': rows, 'meta': meta, 'paceNew': final}
    for r in rows: print(f"{r['scenario']} {r['reading']}: {r['v02']:.2f} → {r['new']:.2f}（{r['total']:+.2f}）｜" + '｜'.join(f'{x:+.2f}' for x in r['steps']))
    if a.json: json.dump(out, open(a.json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

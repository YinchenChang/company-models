# v4.5（5a-1）：期間滾動測試（暫存副本，不動 repo）。
# A. 滾動後第一屏不得出現過期的日期與期間字樣（已決定事項 9，第 3、4 輪）：
#    把副本的 calendar 改為下一個已申報季度（ytdActual 只改截止季與月數，數字不動——只測文字，不測數字），
#    建 HTML 與 Excel（重算），擷取第一屏（scripts/first_screen.py），舊日曆特有的字樣（評價日、年初至今／首期標籤等）不得出現。
#    例外：版本紀錄與來源（有日期的紀錄）；以及字樣整段落在 company.json／共識檔的資料字串內（資料隨資料更新，不是程式寫死）。
# B. 日曆推算：12 月財年首期 0.25／0.5／0.75／1、5 月財年（Oracle 型）、無已申報季度（OpenAI 型），檢查期間長度、折現年數與標籤。
# C. 滾動檢查：只滾日曆、未更新 company.json → asOf 時建置必須失敗（首期一次性金額與期初餘額不得沿用舊季度）。
# 用法：python3 scripts/test_rolling.py（verify.sh 步驟 8b）
import os, sys, json, re, shutil, subprocess
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT); sys.path.insert(0, os.path.join(ROOT, 'scripts'))
import calendar_q
from first_screen import html_first_screens, xlsx_first_screens

fails = []

# ---------- B. 日曆推算 ----------
base = json.load(open(os.path.join(ROOT, 'company.json'), encoding='utf-8'))
CASES = [  # (說明, 日曆覆寫, 期待：首期長度, 評價日, 年初至今標籤, 期間第一格)
    ('12 月財年 Q1 已申報', dict(fiscalYearEndMonth=12, latestQuarterFiled='FY27Q1'), 0.75, '2027-03-31', '1Q27', 'FY27'),
    ('12 月財年 Q2 已申報', dict(fiscalYearEndMonth=12, latestQuarterFiled='FY26Q2'), 0.5, '2026-06-30', '1H26', 'FY26'),
    ('12 月財年 Q3 已申報', dict(fiscalYearEndMonth=12, latestQuarterFiled='FY26Q3'), 0.25, '2026-09-30', '9M26', 'FY26'),
    ('12 月財年 Q4 已申報', dict(fiscalYearEndMonth=12, latestQuarterFiled='FY26Q4'), 1, '2026-12-31', None, 'FY27'),
    ('5 月財年 Q1 已申報（Oracle 型）', dict(fiscalYearEndMonth=5, latestQuarterFiled='FY27Q1'), 0.75, '2026-08-31', '1Q27', 'FY27'),
    ('無已申報季度（OpenAI 型）', dict(fiscalYearEndMonth=12, latestQuarterFiled=None, firstModelFY='FY26'), 1, '2025-12-31', None, 'FY26'),
]
for name, cal, stub, vd, ytd, p0 in CASES:
    for hz in ('firstFullYearEnd', 'valuationPlus12m'):
        co = dict(base); co['calendar'] = {**base['calendar'], **cal, 'targetHorizon': hz}
        d = calendar_q.derive(co)
        got = (d['periodYears'][0], d['valuationDate'], d['ytdLabel'], d['periods'][0])
        if got != (stub, vd, ytd, p0): fails.append(f'B {name}／{hz}：{got} ≠ {(stub, vd, ytd, p0)}')
        te = d['tEnd']
        if abs(te[0] - stub) > 1e-12 or any(abs(te[i + 1] - te[i] - 1) > 1e-12 for i in range(len(te) - 1)): fails.append(f'B {name}：折現年數 {te}')
        ex = [k - d['evOffset'] for k in range(1, len(te))]      # EV/EBITDA 腿各錨定期的折回年數
        if min(ex) < 0: fails.append(f'B {name}／{hz}：錨定期折回年數出現負值（需要內插口徑）：{ex}')
print(f'B 日曆推算：{len(CASES)} 種 × 2 種目標價時點，失敗 {len(fails)}')

# ---------- A. 滾動後第一屏 ----------
d = os.path.join(ROOT, 'out', 'rtest')
shutil.rmtree(d, ignore_errors=True)
shutil.copytree(ROOT, d, ignore=shutil.ignore_patterns('out', 'dist', '.git', '__pycache__'))
co = json.load(open(os.path.join(d, 'company.json'), encoding='utf-8'))
old = calendar_q.derive(co)
fy, q = calendar_q._parse_q(co['calendar']['latestQuarterFiled'])
nq = f'FY{(fy + (q == 4)) % 100:02d}Q{q % 4 + 1}'
co['calendar']['latestQuarterFiled'] = co['calendar']['latestQuarterReported'] = nq
new = calendar_q.derive(co)
co['ytdActual'].update(throughQuarter=nq, months=new['ytdMonths'], label=co['ytdActual']['label'].replace(old['ytdLabel'], new['ytdLabel'] or ''))
co['historicalPL'][-1]['year'] = new['ytdLabel']  # 資料更新時一併更新的標籤（數字不動）
# C. 滾動檢查（首期一次性金額與期初餘額的所屬季度）：只滾日曆、未更新 asOf 時必須失敗，且逐項列出全部欄位
try:
    calendar_q.apply(json.loads(json.dumps(co))); fails.append('C 滾動後 asOf 未更新卻沒有失敗')
except AssertionError as e:
    n = sum(1 for ln in str(e).splitlines() if '尚未依最新已申報季度更新' in ln)
    _nr = sum(1 for _, p, _ in calendar_q.ROLL_FIELDS if not (p.split('.')[0] in calendar_q.OPT_TOP and p.split('.')[0] not in co))  # 選用區段不存在時不列
    if n != _nr: fails.append(f'C 滾動檢查只列出 {n} 項，應為 {_nr} 項')
    print(f'C 滾動檢查：只滾日曆、未更新 asOf → 建置失敗並列出 {n} 項（清單 {_nr} 項）')
co['asOf'].update({p: nq for _, p, _ in calendar_q.ROLL_FIELDS if not (p.split('.')[0] in calendar_q.OPT_TOP and p.split('.')[0] not in co)})  # 本測試只測文字：數值不動、只標為已檢視
json.dump(co, open(os.path.join(d, 'company.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
os.makedirs(os.path.join(d, 'out'), exist_ok=True)
env = dict(os.environ, LC_ALL='C.UTF-8', LANG='C.UTF-8', CRWV_SKIP_QCHECK='1')  # 只改日曆、不改數字：季度層一致性檢查必然不符，測試副本略過
run = lambda *a: subprocess.run(a, cwd=d, env=env, capture_output=True, text=True)
r1 = run('python3', 'build_html_portable.py', '0.0', 'out/r.html', '2026-01-01')
r2 = run('python3', 'build_xlsx.py', 'out/r.xlsx')
r3 = run('python3', 'scripts/recalc.py', 'out/r.xlsx', '180') if r2.returncode == 0 else r2
for r, what in ((r1, '建 HTML'), (r2, '建 Excel'), (r3, '重算')):
    if r.returncode: fails.append(f'A {what}失敗：{(r.stdout + r.stderr)[-400:]}')

def toks(c):  # 舊日曆特有的字樣（新日曆沒有者才算過期）
    out = {c['valuationDate'], c['valuationMD']}
    for k in ('ytdLabel', 'stubLabel'):
        if c[k]: out.add(c[k])
    for k in ('ytdShort', 'ytdShortAlt', 'stubShort', 'ytdWord', 'stubWord', 'filedQLabel'):
        if c.get(k): out.add(c[k])
    return out
stale = sorted(toks(old) - toks(new), key=len, reverse=True)
data_str = set()
def walk(x):
    if isinstance(x, dict): [walk(v) for v in x.values()]
    elif isinstance(x, list): [walk(v) for v in x]
    elif isinstance(x, str): data_str.add(x)
QDATA = set()
def qwalk(x):
    if isinstance(x, dict): [qwalk(v) for v in x.values()]
    elif isinstance(x, list): [qwalk(v) for v in x]
    elif isinstance(x, str): QDATA.add(x)
qwalk(co.get('quarterly', {}))
walk(co); walk(json.load(open(os.path.join(d, co['meta']['consensusFile']), encoding='utf-8')))
SKIP_VIEWS = ('版本紀錄', '來源')

def scan(screens, kind):
    hits = []
    for view, text in screens.items():
        if any(s in view for s in SKIP_VIEWS): continue
        for line in text.split('\n'):
            for t in stale:
                for m in re.finditer(re.escape(t), line):
                    # 字樣在較長的英數字詞中（例如 1H 是 1H26 的一部分另計）或整段落在資料字串內者不算
                    if any(t in s and s in line for s in data_str if len(s) >= len(t) + 2 or s in QDATA): continue
                    if re.match(r'[0-9A-Za-z]', line[m.end():m.end() + 1] or ' ') or re.match(r'[0-9A-Za-z]', line[m.start() - 1:m.start()] if m.start() else ' '): continue
                    hits.append((line, f'{kind}［{view}］「{t}」'))
    seen = {}
    for line, where in hits: seen.setdefault(line, []).append(where)
    return [f'{"、".join(sorted(set(w))[:3])}{"…" if len(set(w)) > 3 else ""}：{line[:140]}' for line, w in seen.items()]

if not any(f.startswith('A ') for f in fails):
    hits = scan(html_first_screens(os.path.join(d, 'out', 'r.html')), 'HTML') + scan(xlsx_first_screens(os.path.join(d, 'out', 'r.xlsx')), 'Excel')
    print(f'A 滾動 {co["calendar"]["latestQuarterFiled"]}（評價日 {old["valuationDate"]} → {new["valuationDate"]}）：過期字樣 {stale}；第一屏命中 {len(hits)} 行')
    for h in hits: print('   ' + h)
    if hits: fails.append(f'A 第一屏出現過期字樣 {len(hits)} 行')
shutil.rmtree(d, ignore_errors=True)
print('期間滾動測試：' + ('通過' if not fails else '不通過\n  ' + '\n  '.join(fails)))
sys.exit(1 if fails else 0)

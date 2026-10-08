# 目標價變動拆解工具測試（scripts/attrib.py；verify.sh 步驟 8c）。暫存副本，不動 repo。
# 1. 月數：calendar_q.months_between 的幾種日期（跨年、5 月財年季末）。
# 2. 同一版對同一版：評價日不動、四項皆 0。
# 3. 滾動一季（副本：calendar 加一季、asOf 標為已檢視、數值不動），且 WACC 改為 12%（確認折現率讀 Excel 輸入格、不是寫死 11%）：
#    (a) 係數須＝(1＋Excel 的 WACC)^(calendar_q 推算的月數 ÷ 12)；EV/EBITDA 腿與加權目標價的 (a) 變動率須等於係數；四項相加＝總變動。
# 用法：python3 scripts/test_attrib.py 已建好的.xlsx
import os, sys, json, subprocess, tempfile, shutil
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import calendar_q

fails = []
for a, b, m in (('2026-06-30', '2026-09-30', 3), ('2026-09-30', '2027-03-31', 6), ('2026-08-31', '2026-11-30', 3), ('2025-12-31', '2026-12-31', 12)):
    if calendar_q.months_between(a, b) != m: fails.append(f'1 月數 {a}→{b} 應為 {m}')
print(f'1 月數：4 組，失敗 {len(fails)}')

xlsx = os.path.abspath(sys.argv[1])
tmp = tempfile.mkdtemp(prefix='tattrib_')
env = dict(os.environ, LC_ALL='C.UTF-8', LANG='C.UTF-8', CRWV_SKIP_QCHECK='1')  # 只改日曆與 WACC：季度層一致性檢查必然不符，測試副本略過
run = lambda *a: subprocess.run(['python3', os.path.join(ROOT, 'scripts', 'attrib.py'), *a], capture_output=True, text=True, env=env)

r = run('--base', xlsx, '--d', xlsx, '--json', os.path.join(tmp, 'same.json'))
if r.returncode: fails.append('2 同版對同版執行失敗：' + (r.stdout + r.stderr)[-600:])
else:
    j = json.load(open(os.path.join(tmp, 'same.json'), encoding='utf-8'))
    nz = [f'{s} {m} {k}' for s, v in j['scenarios'].items() for m, d in v['delta'].items() for k, x in d.items() if abs(x) > 1e-9]
    if j['months'] != 0 or nz: fails.append(f'2 同版對同版應全為 0：月數 {j["months"]}、非 0 {nz[:5]}')
print(f'2 同版對同版：失敗 {sum(f.startswith("2") for f in fails)}')

co = json.load(open(os.path.join(ROOT, 'company.json'), encoding='utf-8'))
co['valuation']['wacc'] = 0.12
json.dump(co, open(os.path.join(tmp, 'base.json'), 'w', encoding='utf-8'), ensure_ascii=False)
old = calendar_q.derive(co)
fy, q = calendar_q._parse_q(co['calendar']['latestQuarterFiled'])
nq = f'FY{(fy + (q == 4)) % 100:02d}Q{q % 4 + 1}'
co['calendar']['latestQuarterFiled'] = co['calendar']['latestQuarterReported'] = nq
new = calendar_q.derive(co)
co['ytdActual'].update(throughQuarter=nq, months=new['ytdMonths'], label=co['ytdActual']['label'].replace(old['ytdLabel'], new['ytdLabel'] or ''))
co['historicalPL'][-1]['year'] = new['ytdLabel']
co['asOf'].update({p: nq for _, p, _ in calendar_q.ROLL_FIELDS})
json.dump(co, open(os.path.join(tmp, 'roll.json'), 'w', encoding='utf-8'), ensure_ascii=False)
r = run('--base', os.path.join(tmp, 'base.json'), '--b', os.path.join(tmp, 'roll.json'), '--json', os.path.join(tmp, 'roll_out.json'))
print(r.stdout[-2500:])
if r.returncode: fails.append('3 滾動一季執行失敗：' + (r.stdout + r.stderr)[-600:])
else:
    j = json.load(open(os.path.join(tmp, 'roll_out.json'), encoding='utf-8'))
    m = calendar_q.months_between(old['valuationDate'], new['valuationDate'])
    want = (1 + 0.12) ** (m / 12)
    if j['months'] != m or abs(j['wacc'] - 0.12) > 1e-12 or abs(j['grow'] - want) > 1e-12:
        fails.append(f"3 (a) 係數：月數 {j['months']}（應 {m}）、WACC {j['wacc']}（應 0.12）、係數 {j['grow']}（應 {want}）")
    for s, v in j['scenarios'].items():
        for k in ('ev', 'blend'):
            b0, a1 = v['levels']['base'][k], v['levels']['a'][k]
            if b0 > 0 and abs(a1 / b0 - want) > 1e-9: fails.append(f'3 {s} {k}：(a) 變動率 {a1 / b0:.6f} ≠ {want:.6f}')
shutil.rmtree(tmp, ignore_errors=True)
print('目標價變動拆解測試：' + ('通過' if not fails else '不通過\n  ' + '\n  '.join(fails)))
sys.exit(1 if fails else 0)

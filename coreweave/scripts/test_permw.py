# W2（2026-10-07）：每 MW 改寫的暫存副本測試（verify.sh 步驟 8d）。
# 測試 A｜MW 設施口徑：副本設 meta.mwBasis＝facility，Tokenomics 每 MW 值須 ÷ IF_FacilityGW（基準 1.2）：
#   每 MW 建置成本、每 MW 電費、每 MW GPU 數、經濟持有成本＝IT 口徑值 ÷ 1.2；IT MW＝公司 MW ÷ 1.2；cmp31 基準全部一致。
# 測試 B｜GPU 小時價格路線：副本設 methodology.perMw.revenue＝gpuHr、pricing.gpuHr 填虛構價格（測試用，不入 repo、非真實數字），
#   「每 MW 年收入」須＝Σ 平均在役占比 × 每 MW GPU 數 × 價格 × 8,760；價格低／高情境改變目標價；cmp31 基準全部一致。
import os, sys, json, shutil, subprocess, re
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'out')
FAIL = []


def run(cmd, cwd=None, ok=(0,)):
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if r.returncode not in ok:
        raise RuntimeError(' '.join(cmd) + '\n' + r.stdout[-2000:] + r.stderr[-2000:])
    return r


def copy_repo(name, edit):
    d = os.path.join(OUT, name)
    shutil.rmtree(d, ignore_errors=True)
    shutil.copytree(ROOT, d, ignore=shutil.ignore_patterns('out', 'dist', '.git', '__pycache__'))
    os.makedirs(os.path.join(d, 'out'))
    p = os.path.join(d, 'company.json'); co = json.load(open(p, encoding='utf-8')); edit(co)
    open(p, 'w', encoding='utf-8').write(json.dumps(co, indent=1, ensure_ascii=False) + '\n')
    return d


def build(d, label):
    x = os.path.join(d, 'out', 't.xlsx')
    run(['python3', os.path.join(d, 'build_xlsx.py'), x]); run(['python3', os.path.join(d, 'scripts', 'recalc.py'), x, '120'])
    if run(['python3', os.path.join(d, 'scripts', 'permw_sens.py'), x], ok=(0, 3)).returncode == 3:  # 副本的敏感度快照隨設定改變：重建
        run(['python3', os.path.join(d, 'build_xlsx.py'), x]); run(['python3', os.path.join(d, 'scripts', 'recalc.py'), x, '120'])
    run(['python3', os.path.join(d, 'xlx.py'), x, '2', os.path.join(d, 'out', 'xl17_2.json')])
    log = run(['node', os.path.join(d, 'cmp31.js'), 'base'], cwd=os.path.join(d, 'out')).stdout
    nok = len(re.findall(r'^OK ', log, re.M)); nbad = len(re.findall(r'^(XX |MISSING)|MISSING XL KEY', log, re.M))
    if nbad or not nok: FAIL.append(f'{label}：cmp31 基準 OK {nok}、不一致 {nbad}：' + ' / '.join(l for l in log.splitlines() if l.startswith(('XX', 'MISSING')))[:600])
    print(f'  {label}：cmp31 基準 {nok} 項、不一致 {nbad}')
    return json.load(open(os.path.join(d, 'out', 'xl17_2.json'), encoding='utf-8'))


near = lambda a, b, tol=1e-6: isinstance(a, (int, float)) and isinstance(b, (int, float)) and abs(a - b) <= tol * max(1, abs(b))
CO = json.load(open(os.path.join(ROOT, 'company.json'), encoding='utf-8'))
print('=== 基準（目前 company.json）')
X0 = json.load(open(os.path.join(OUT, 'xl17_2.json'), encoding='utf-8')) if os.path.exists(os.path.join(OUT, 'xl17_2.json')) else build(copy_repo('permw_base', lambda c: None), '基準')

print('=== 測試 A｜MW 設施口徑（meta.mwBasis＝facility）')
XA = build(copy_repo('permw_facility', lambda c: c['meta'].update(mwBasis='facility')), '測試 A')
for k in ('每MW經濟性|每 MW 電費（世代加權）', '每MW經濟性|每 MW GPU 數（世代加權）', '每MW經濟性|每 MW 經濟持有成本（不賠錢下限）',
          '每MW經濟性|Tokenomics IT 設備（IF_CapexIT）', '每MW經濟性|平均在役 MW（IT 關鍵電力）'):
    a, b = XA[k], X0[k]
    if not all(near(x, y / 1.2) for x, y in zip(a, b)): FAIL.append(f'測試 A：{k} 應為 IT 口徑 ÷ 1.2：{a} vs {b}')
if CO['methodology']['perMw']['capex'] == 'tokenomics' and not all(near(x, y / 1.2) for x, y in zip(XA['輸入與假設|每 MW 建置成本'], X0['輸入與假設|每 MW 建置成本'])):
    FAIL.append('測試 A：每 MW 建置成本未 ÷ 1.2')
if not all(near(x, y) for x, y in zip(XA['每MW經濟性|平均在役 MW（設施＝IT × IF_FacilityGW）'], X0['每MW經濟性|平均在役 MW 合計'])):
    FAIL.append('測試 A：設施口徑下「設施 MW」應等於公司 MW')
print('  設施口徑：Tokenomics 每 MW 值 ÷ 1.2、IT MW＝公司 MW ÷ 1.2 已核對')

print('=== 測試 B｜GPU 小時價格路線（虛構價格，只在暫存副本）')
def _gpu(c):
    c['methodology']['perMw']['revenue'] = 'gpuHr'
    c['pricing']['gpuHr'] = {g: {'base': 3.0 + 0.5 * j, 'low': 2.5 + 0.5 * j, 'high': 3.5 + 0.5 * j, 'source': '測試用虛構價格', 'date': '2026-10-07', 'tag': '[Assumed]'}
                             for j, g in enumerate(c['fleet']['generations'])}
XB = build(copy_repo('permw_gpuhr', _gpu), '測試 B')
rv, gp = XB['輸入與假設|每 MW 年收入'], XB['每MW經濟性|GPU 小時價格路線：每 MW 年收入']
if not all(near(a * 1000, b) for a, b in zip(rv, gp)): FAIL.append(f'測試 B：每 MW 年收入 ≠ GPU 小時路線：{rv} vs {gp}')
mix = [XB[f'每MW經濟性|平均在役占比｜{g}'] for g in CO['fleet']['generations']]
gpu = [XB[f'TK 收入參考｜{g}'.join(['輸入與假設|', ''])][0] for g in CO['fleet']['generations']]
ind = [sum(mix[j][i] * gpu[j] * (3.0 + 0.5 * j) for j in range(len(gpu))) * 8760 / 1e6 for i in range(5)]
if not all(near(a, b, 1e-9) for a, b in zip(gp, ind)): FAIL.append(f'測試 B：Python 獨立計算 {ind} ≠ Excel {gp}')
sn = json.load(open(os.path.join(OUT, 'permw_gpuhr', 'permw_sens.json'), encoding='utf-8'))['scenarios']['base']
if not (sn['pxLow'] and sn['pxHigh'] and sn['pxLow'][0] < sn['base'][0] < sn['pxHigh'][0]): FAIL.append(f'測試 B：價格低／高情境未改變目標價：{sn}')
print(f"  GPU 小時路線：每 MW 年收入 FY30 {gp[4]:.2f} US$m（Python 獨立計算一致）；目標價 低／基準／高 {sn['pxLow'][0]:.2f}／{sn['base'][0]:.2f}／{sn['pxHigh'][0]:.2f}")

print('=== 結果')
if FAIL:
    print('test_permw：失敗'); [print('  ' + f) for f in FAIL]; sys.exit(1)
print('test_permw：全部通過（測試 A 設施口徑、測試 B GPU 小時價格路線）')

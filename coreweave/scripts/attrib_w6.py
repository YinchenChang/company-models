# W6（2026-10-09）：v4.7 → v4.8（接 Tokenomics v5.31）目標價變動拆解（已決定事項 12(3)；工作單 W6 第 5 步）。
# (a) 時間推移＝0（評價日不變）、(b) 實際數更新＝0（無新財報）；
# (c) 假設變更＝Tokenomics v5.27 → v5.31 的數值（GB300 機架價格 5.0 → 4.3 $M、IT 維護等值費率 3% → 1.57%），依序：
#   ① GB300 機架價格（資本支出、折舊、稅險）：IF_CapexIT、IF_CapexTotal、IF_DeprIT、IF_TaxIns 換 v5.31（IT 維護仍為 v5.27 等值費率）
#   ② 收入錨（持有成本）：IF_HoldEcon、IF_GPUhrEcon、IF_HoldAcct、IF_OpexGW、L1_ 名稱換 v5.31（持有成本同時含機架價格與維護費率的變動；k 輸入不動）
# (d) 方法變更：
#   ③ IT 維護依機齡兩段（methodology.perMw.maint＝age：保固期內 MW × IF_MaintITWarr＋保固期滿 MW × IF_MaintITPost；全部名稱＝v5.31）＝v4.8
#   ④ MW 口徑：查證後未改（meta.mwBasis 維持 IT），定義為 0
# 數值取自 Excel（scripts/attrib.py 的 read／legs）；另讀融資缺口＝MAX(0, −FY30 融資前累積現金)。
# 用法：python3 scripts/attrib_w6.py --v47 v4.7成品.xlsx [--v48 v4.8.xlsx] [--json 輸出.json]
#   v5.27 快照自 git 歷史（V527_REF，預設 8671a11）取回；檢查「依序各步相加＝v4.8 − v4.7」（誤差 < 0.01 美元／股）、④＝0。
import sys, os, json, shutil, argparse, tempfile, subprocess
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'scripts'))
from attrib import read, legs, SCN
from attrib_w5 import gap

V527_REF = os.environ.get('V527_REF', '8671a11')
G1 = ['IF_CapexIT', 'IF_CapexTotal', 'IF_DeprIT', 'IF_TaxIns']
G2 = ['IF_HoldEcon', 'IF_GPUhrEcon', 'IF_HoldAcct', 'IF_OpexGW', 'L1_GPUhr_GB200_vsCW', 'L1_GPUhr_GB300_vsBE']


def old_snap():
    s = subprocess.run(['git', '-C', ROOT, 'show', f'{V527_REF}:coreweave/data/tokenomics_snapshot_v5.27.json'], capture_output=True, text=True, check=True).stdout
    return json.loads(s)


def build(tag, mod, tmp, new_names=()):
    d = os.path.join(tmp, tag)
    shutil.copytree(ROOT, d, ignore=shutil.ignore_patterns('out', 'dist', '.git', '__pycache__', 'docs'))
    p = os.path.join(d, 'company.json'); co = json.load(open(p, encoding='utf-8'))
    if new_names is not None:  # 混合快照：v5.31 中 new_names 以外的名稱換回 v5.27 值；IT 維護等值費率口徑
        sp = os.path.join(d, co['tokenomics']['snapshotFile']); sn = json.load(open(sp, encoding='utf-8')); old = old_snap()['items']
        for n in list(sn['items']):
            if n in old and n not in new_names: sn['items'][n] = old[n]
        json.dump(sn, open(sp, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
        co['methodology']['perMw']['maint'] = 'flat'
    mod(co)
    open(p, 'w', encoding='utf-8').write(json.dumps(co, indent=1, ensure_ascii=False) + '\n')
    x = os.path.join(d, f'{tag}.xlsx')
    r = subprocess.run(['python3', 'build_xlsx.py', x], cwd=d, capture_output=True, text=True, env=dict(os.environ, LC_ALL='C.UTF-8', LANG='C.UTF-8'))
    if r.returncode: sys.exit(f'建 Excel 失敗（{tag}）：\n{(r.stdout + r.stderr)[-1500:]}')
    return x


STEPS = [('rack', '① (c) GB300 機架價格：資本支出、折舊、稅險（IF_CapexIT 等換 v5.31）'),
         ('anchor', '② (c) 收入錨：Tokenomics 持有成本 v5.31（IF_HoldEcon、IF_GPUhrEcon；含機架價格與維護費率；k 輸入不動）'),
         ('age', '③ (d) IT 維護依機齡兩段（保固期內 IF_MaintITWarr、期滿 IF_MaintITPost）'),
         ('mw', '④ (d) MW 口徑（查證後未改，維持 IT）')]
CLS = {'rack': 'c', 'anchor': 'c', 'age': 'd', 'mw': 'd'}


def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--v47', required=True); ap.add_argument('--v48'); ap.add_argument('--json')
    a = ap.parse_args()
    tmp = tempfile.mkdtemp(prefix='attribw6_')
    try:
        x1 = build('rack', lambda co: None, tmp, G1)
        x2 = build('anchor', lambda co: None, tmp, G1 + G2)
        x48 = os.path.abspath(a.v48) if a.v48 else build('v48', lambda co: None, tmp, None)
        X = {'v47': read(os.path.abspath(a.v47)), 'rack': read(x1), 'anchor': read(x2), 'age': read(x48)}
        G = {'v47': gap(os.path.abspath(a.v47)), 'rack': gap(x1), 'anchor': gap(x2), 'age': gap(x48)}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    X['mw'], G['mw'], X['v48'], G['v48'] = X['age'], G['age'], X['age'], G['age']
    res, fails = {'steps': STEPS, 'scenarios': {}}, []
    for s in (1, 2, 3):
        L = {k: legs(X[k], s) for k in X}
        o = {'v47': L['v47'], 'v48': L['v48'], 'gap': {k: G[k][s] for k in ('v47', 'rack', 'anchor', 'v48')}, 'seq': {}}
        p = 'v47'
        for k, _ in STEPS:
            o['seq'][k] = {m: L[k][m] - L[p][m] for m in ('dcf', 'ev', 'blend')}; o['seq'][k]['gap'] = G[k][s] - G[p][s]; p = k
        tot = L['v48']['blend'] - L['v47']['blend']; ssum = sum(v['blend'] for v in o['seq'].values())
        o['total'] = tot
        o['abcd'] = {'a': 0.0, 'b': 0.0, 'c': sum(o['seq'][k]['blend'] for k in o['seq'] if CLS[k] == 'c'), 'd': sum(o['seq'][k]['blend'] for k in o['seq'] if CLS[k] == 'd')}
        if abs(ssum - tot) >= 0.01: fails.append(f'{SCN[s]}：依序各步相加 {ssum:.4f} ≠ 總變動 {tot:.4f}')
        if abs(o['seq']['mw']['blend']) >= 1e-9: fails.append(f'{SCN[s]}：④ MW 口徑不為 0')
        for k in ('v47', 'v48'):
            if X[k]['dcfMode'] == 1 and abs(L[k]['blend'] - X[k][s]['blendXl']) > 1e-6: fails.append(f'{SCN[s]} {k}：重組加權 ≠ Excel')
        res['scenarios'][SCN[s]] = o
        print(f"\n{SCN[s]}：v4.7 ${L['v47']['blend']:.2f} → v4.8 ${L['v48']['blend']:.2f}（總變動 {tot:+.2f}；(a) 0、(b) 0、(c) {o['abcd']['c']:+.2f}、(d) {o['abcd']['d']:+.2f}）；融資缺口 {G['v47'][s]:.1f} → {G['v48'][s]:.1f}")
        print('| 逐項 | 加權目標價 | DCF 腿 | EV/EBITDA 腿 | 融資缺口 |'); print('|---|---|---|---|---|')
        for k, nm in STEPS:
            q = o['seq'][k]; print(f"| {nm} | {q['blend']:+.2f} | {q['dcf']:+.2f} | {q['ev']:+.2f} | {q['gap']:+.1f} |")
        print(f'| 合計 | {ssum:+.2f} | | | |')
    if a.json: json.dump(res, open(a.json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('\n拆解檢查：' + ('通過（三情境：依序各步相加＝v4.8 − v4.7，誤差 < 0.01 美元／股；④＝0；加權目標價與 Excel 一致）' if not fails else '不通過\n  ' + '\n  '.join(fails)))
    return 1 if fails else 0


if __name__ == '__main__':
    sys.exit(main())

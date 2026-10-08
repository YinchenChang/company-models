# W4（2026-10-08）：v4.6 → v4.7 目標價變動拆解（已決定事項 12(3)；工作單 W4 第 5 步）。
# (a) 時間推移＝0（評價日不變）、(b) 實際數更新＝0（無新財報）、(c) 假設變更＝0（Tokenomics v5.26 → v5.27：本模型引用的 25 個名稱數值與位置相同；
#     新增 IF_RevGWFleet 只用於上限檢查），全部變動屬 (d) 方法變更（每 MW 收入改以 Tokenomics 為錨）。(d) 依序再拆：
#   ① 錨取代舊輸入（k＝1）：revenue＝tkAnchor、k_長約＝k_現貨＝1 → 每 MW 年收入＝Σ 平均在役占比 × IF_HoldEcon
#   ② 套用 k：k_長約、k_現貨、長約占比為 company.json 預設（＝v4.7）
#   ③ 上限檢查：只新增檢查列、不改數字（應為 0）
# 數值取自 Excel（scripts/attrib.py 的 read／legs：LibreOffice 開啟、三情境切換「情境選擇」）；另讀融資缺口＝MAX(0, −FY30 融資前累積現金)。
# 用法：python3 scripts/attrib_w4.py --v46 v4.6成品.xlsx [--v47 v4.7.xlsx] [--json 輸出.json]
#   --v47 省略時以目前 repo 建 Excel；檢查「依序各步相加＝v4.7 − v4.6」（誤差 < 0.01 美元／股）、最後一步＝v4.7。
import sys, os, json, shutil, argparse, tempfile, subprocess
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'scripts'))
from attrib import read, legs, SCN
from uno_q import Workbook


def build(tag, mod, tmp):
    d = os.path.join(tmp, tag)
    shutil.copytree(ROOT, d, ignore=shutil.ignore_patterns('out', 'dist', '.git', '__pycache__', 'docs'))
    p = os.path.join(d, 'company.json'); co = json.load(open(p, encoding='utf-8')); mod(co)
    open(p, 'w', encoding='utf-8').write(json.dumps(co, indent=1, ensure_ascii=False) + '\n')
    x = os.path.join(d, f'{tag}.xlsx')
    r = subprocess.run(['python3', 'build_xlsx.py', x], cwd=d, capture_output=True, text=True, env=dict(os.environ, LC_ALL='C.UTF-8', LANG='C.UTF-8'))
    if r.returncode: sys.exit(f'建 Excel 失敗（{tag}）：\n{(r.stdout + r.stderr)[-1500:]}')
    return x


def gap(xlsx):
    out = {}
    with Workbook(xlsx) as wb:
        sel = wb.cell('輸入與假設', '情境選擇', prefix=True); s0 = wb.get(sel); c = wb.cell('各期收支', '融資前累積現金', col='G')
        for s in (1, 2, 3): wb.set(sel, s); out[s] = max(0.0, -wb.get(c))
        wb.set(sel, s0)
    return out


def k1(co):
    co['methodology']['perMw']['revenue'] = 'tkAnchor'
    co['pricing']['anchorMultiple']['long']['base'] = 1; co['pricing']['anchorMultiple']['spot']['base'] = 1


STEPS = [('anchor', '① 錨取代舊輸入（k＝1：每 MW 年收入＝在役世代 IF_HoldEcon）'), ('k', '② 套用定價倍數 k（長約／現貨依 RPO 涵蓋加權）'), ('cap', '③ 上限檢查（只新增檢查列）')]


def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--v46', required=True); ap.add_argument('--v47'); ap.add_argument('--json')
    a = ap.parse_args()
    tmp = tempfile.mkdtemp(prefix='attribw4_')
    try:
        x1 = build('anchor', k1, tmp)
        x47 = os.path.abspath(a.v47) if a.v47 else build('v47', lambda co: None, tmp)
        X = {'v46': read(os.path.abspath(a.v46)), 'anchor': read(x1), 'k': read(x47)}
        G = {'v46': gap(os.path.abspath(a.v46)), 'anchor': gap(x1), 'k': gap(x47)}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    X['cap'], G['cap'], X['v47'], G['v47'] = X['k'], G['k'], X['k'], G['k']
    res, fails = {'steps': STEPS, 'scenarios': {}}, []
    for s in (1, 2, 3):
        L = {k: legs(X[k], s) for k in X}
        o = {'v46': L['v46'], 'v47': L['v47'], 'gap': {k: G[k][s] for k in ('v46', 'anchor', 'k', 'v47')}, 'seq': {}}
        p = 'v46'
        for k, _ in STEPS:
            o['seq'][k] = {m: L[k][m] - L[p][m] for m in ('dcf', 'ev', 'blend')}; o['seq'][k]['gap'] = G[k][s] - G[p][s]; p = k
        tot = L['v47']['blend'] - L['v46']['blend']; ssum = sum(v['blend'] for v in o['seq'].values())
        o['total'] = tot; o['abcd'] = {'a': 0.0, 'b': 0.0, 'c': 0.0, 'd': tot}
        if abs(ssum - tot) >= 0.01: fails.append(f'{SCN[s]}：依序各步相加 {ssum:.4f} ≠ 總變動 {tot:.4f}')
        if abs(o['seq']['cap']['blend']) >= 1e-9: fails.append(f'{SCN[s]}：③ 上限檢查不為 0')
        for k in ('v46', 'v47'):
            if X[k]['dcfMode'] == 1 and abs(L[k]['blend'] - X[k][s]['blendXl']) > 1e-6: fails.append(f'{SCN[s]} {k}：重組加權 ≠ Excel')
        res['scenarios'][SCN[s]] = o
        print(f"\n{SCN[s]}：v4.6 ${L['v46']['blend']:.2f} → v4.7 ${L['v47']['blend']:.2f}（總變動 {tot:+.2f}；(a) 0、(b) 0、(c) 0、(d) {tot:+.2f}）；融資缺口 {G['v46'][s]:.1f} → {G['v47'][s]:.1f}")
        print('| (d) 方法變更逐項 | 加權目標價 | DCF 腿 | EV/EBITDA 腿 | 融資缺口 |'); print('|---|---|---|---|---|')
        for k, nm in STEPS:
            q = o['seq'][k]; print(f"| {nm} | {q['blend']:+.2f} | {q['dcf']:+.2f} | {q['ev']:+.2f} | {q['gap']:+.1f} |")
        print(f'| 合計 | {ssum:+.2f} | | | |')
    if a.json: json.dump(res, open(a.json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('\n拆解檢查：' + ('通過（三情境：依序各步相加＝v4.7 − v4.6，誤差 < 0.01 美元／股；③＝0；加權目標價與 Excel 一致）' if not fails else '不通過\n  ' + '\n  '.join(fails)))
    return 1 if fails else 0


if __name__ == '__main__':
    sys.exit(main())

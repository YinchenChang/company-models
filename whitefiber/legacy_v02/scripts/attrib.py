# 目標價變動拆解（已決定事項 12；(a) 的算法 2026-09-26 Andy 決定）：把兩版之間的加權目標價變動拆成
#   (a) 時間推移＝定義值（純時間價值，以 WACC 累積）：前一版兩條腿（0 截斷前）×（1＋WACC）^（評價日後移月數 ÷ 12），再做 0 截斷與加權。
#       等同「離開模型期的季度以模型值當作暫定實際數、淨負債與股數推到新評價日、其餘路徑不變」，但淨負債以 WACC 而非債務成本累積。
#       WACC 讀前一版 Excel 的 WACC 輸入格；月數由 calendar_q.months_between 以兩版 Excel 的評價日（calendar_q 推算、建置時寫入）計算。
#   (b) 實際數更新：填入 10-Q 的年初至今實際、期初餘額與到期表後重算，減去 (a)。含兩個來源：該季實際與模型的差異；
#       以及融資時點殘差（(a) 以 WACC 累積淨負債，實際以債務成本累積；該季實際發股）。
#   (c) 假設變更、(d) 方法變更：依序重算的差額。四項相加＝總變動（逐項檢查）。
# 只拆解目標價採用的 0 截斷口徑；選擇權模式的 DCF 腿只列前後值、不拆解（Andy 決定）。
# 數值一律取自 Excel（LibreOffice 開啟、重算；已決定事項 8）：三情境以「情境選擇」輸入格切換，其餘輸入不動。
# 用法：python3 scripts/attrib.py --base 前一版 [--b 步驟] [--c 步驟] --d 新版 [--json 輸出.json]
#   各步驟可為：已建好的 .xlsx（例如 dist/ 成品）；company.json（以目前 repo 的程式建 Excel）；或 repo 目錄（以該目錄的程式與資料建 Excel）。
#   省略的中間步驟＝與前一步相同（變動 0）；--d 省略時，最後一個給定的步驟即為新版。
import sys, os, json, shutil, argparse, tempfile, subprocess
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT); sys.path.insert(0, os.path.join(ROOT, 'scripts'))
import calendar_q
from uno_q import Workbook

IN, VA = '輸入與假設', '評價_DCF與目標價'
SCN = {1: '保守', 2: '基準', 3: '積極'}


def build(spec, tmp):
    """步驟 → 可開啟的 .xlsx 路徑。"""
    if spec.endswith('.xlsx'): return os.path.abspath(spec)
    if os.path.isdir(spec): src, cj = spec, None
    elif spec.endswith('.json'): src, cj = ROOT, spec
    else: sys.exit(f'無法辨識的步驟：{spec}（應為 .xlsx、company.json 或 repo 目錄）')
    d = tempfile.mkdtemp(prefix='attrib_', dir=tmp)
    shutil.copytree(src, d, dirs_exist_ok=True, ignore=shutil.ignore_patterns('out', 'dist', '.git', '__pycache__'))
    if cj: shutil.copy(cj, os.path.join(d, 'company.json'))
    x = os.path.join(d, 'step.xlsx')
    r = subprocess.run(['python3', 'build_xlsx.py', x], cwd=d, capture_output=True, text=True,
                       env=dict(os.environ, LC_ALL='C.UTF-8', LANG='C.UTF-8'))
    if r.returncode: sys.exit(f'建 Excel 失敗（{spec}）：\n{(r.stdout + r.stderr)[-1500:]}')
    return x


def read(xlsx):
    """三情境的兩條腿（0 截斷前後）、權重、加權目標價，以及 WACC、評價日、目標價時點年數。"""
    out = {}
    with Workbook(xlsx) as wb:
        g = lambda sh, lab, **k: wb.get(wb.cell(sh, lab, **k))
        sel = wb.cell(IN, '情境選擇', prefix=True)
        sel0 = wb.get(sel)
        vd = wb.cell(IN, '評價日（已申報季度的季末）'); vd = vd[0].getCellRangeByName(vd[1]).getString()
        out['wacc'], out['valuationDate'], out['targetT'] = g(IN, 'WACC'), vd, g(IN, '目標價時點距評價日')
        out['wdcf'], out['dcfMode'] = g(IN, '權重：DCF'), g(IN, 'DCF 下限方式', prefix=True)
        for s in (1, 2, 3):
            wb.set(sel, s)
            bad = g(VA, 'DCF 失效？', prefix=True)
            fac = g(VA, '折回 ', prefix=True)
            evRaw = (g(VA, '錨定年度企業價值（倍數 × EBITDA）') - g(VA, '錨定年度末淨負債（總債務 − 現金）')) / g(VA, '錨定年度末股數（含瀑布新股）') * fac
            out[s] = {'dcfRaw': g(VA, '目標價區間｜未截斷 DCF 每股'), 'dcfBad': bad, 'evRaw': evRaw,
                      'ev': g(VA, 'EV/EBITDA 每股（融資後）'), 'dcfOpt': g(VA, 'DCF 每股：選擇權（Merton）'),
                      'dcfUsedT': g(VA, 'DCF 每股（推到目標價時點）'), 'blendXl': g(VA, '加權目標價')}
        wb.set(sel, sel0)
    return out


def legs(x, s, grow=1.0):
    """0 截斷口徑的兩條腿與加權（grow＝(a) 的時間推移係數，套在 0 截斷前的值上）。"""
    v = x[s]; T = (1 + x['wacc']) ** x['targetT']
    dcf = 0.0 if v['dcfBad'] else max(0.0, v['dcfRaw'] * grow) * T
    ev = max(0.0, v['evRaw'] * grow)
    wd = 0.0 if v['dcfBad'] else x['wdcf']
    return {'dcf': dcf, 'ev': ev, 'blend': wd * dcf + (1 - wd) * ev}


def main():
    ap = argparse.ArgumentParser(description='目標價變動拆解 (a)(b)(c)(d)')
    ap.add_argument('--base', required=True); ap.add_argument('--b'); ap.add_argument('--c'); ap.add_argument('--d')
    ap.add_argument('--json')
    a = ap.parse_args()
    specs = [('b', a.b), ('c', a.c), ('d', a.d)]
    if not any(v for _, v in specs): sys.exit('至少要給 --b、--c 或 --d 其中一步')
    tmp = tempfile.mkdtemp(prefix='attrib_')
    try:
        X = {'base': read(build(a.base, tmp))}
        prev = 'base'
        for k, v in specs:
            X[k] = read(build(v, tmp)) if v else X[prev]
            prev = k
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    B, D = X['base'], X['d']
    # 新評價日取最後一步；中間步驟的評價日須一致（(b) 起就是新日曆）
    for k in ('b', 'c'):
        if X[k] is not B and X[k]['valuationDate'] != D['valuationDate']:
            sys.exit(f"步驟 ({k}) 的評價日 {X[k]['valuationDate']} ≠ 新版 {D['valuationDate']}：(b) 起應已滾動日曆")
    months = calendar_q.months_between(B['valuationDate'], D['valuationDate'])
    grow = (1 + B['wacc']) ** (months / 12)
    res, fails = {'valuationDate': [B['valuationDate'], D['valuationDate']], 'months': months, 'wacc': B['wacc'], 'grow': grow, 'scenarios': {}}, []
    print(f"評價日 {B['valuationDate']} → {D['valuationDate']}（{months} 個月）；(a) 係數＝(1＋WACC {B['wacc']:.2%})^({months}/12)＝{grow:.5f}")
    for s in (1, 2, 3):
        L = {'base': legs(B, s), 'a': legs(B, s, grow), 'b': legs(X['b'], s), 'c': legs(X['c'], s), 'd': legs(D, s)}
        for k in ('base', 'b', 'c', 'd'):  # 讀值一致性：0 截斷模式下，以兩條腿重組的加權目標價須等於 Excel 的加權目標價
            x = X[k]
            if x['dcfMode'] == 1 and abs(L[k]['blend'] - x[s]['blendXl']) > 1e-6:
                fails.append(f"{SCN[s]}（{k}）：重組加權 {L[k]['blend']:.4f} ≠ Excel {x[s]['blendXl']:.4f}")
        steps = ['base', 'a', 'b', 'c', 'd']
        dl = {m: {st: L[st][m] - L[p][m] for p, st in zip(steps, steps[1:])} for m in ('dcf', 'ev', 'blend')}
        for m in dl:
            tot = L['d'][m] - L['base'][m]
            if abs(sum(dl[m].values()) - tot) > 1e-9: fails.append(f'{SCN[s]} {m}：四項相加 ≠ 總變動')
        res['scenarios'][SCN[s]] = {'levels': L, 'delta': dl,
                                     'optionDcf': [B[s]['dcfOpt'] * (1 + B['wacc']) ** B['targetT'], D[s]['dcfOpt'] * (1 + D['wacc']) ** D['targetT']]}
        pc = lambda x, y: f'{(y / x - 1) * 100:+.1f}%' if abs(x) > 1e-9 else '—'
        print(f"\n{SCN[s]}（0 截斷口徑）")
        print('| 項目 | 前一版 | (a) 時間推移 | (b) 實際數 | (c) 假設 | (d) 方法 | 新版 | 總變動 |')
        print('|---|---|---|---|---|---|---|---|')
        for m, nm in (('dcf', 'DCF 腿（推到目標價時點）'), ('ev', 'EV/EBITDA 腿'), ('blend', '加權目標價')):
            d = dl[m]
            print(f"| {nm} | ${L['base'][m]:.2f} | {d['a']:+.2f}（{pc(L['base'][m], L['a'][m])}） | {d['b']:+.2f} | {d['c']:+.2f} | {d['d']:+.2f} | ${L['d'][m]:.2f} | {L['d'][m] - L['base'][m]:+.2f} |")
        o = res['scenarios'][SCN[s]]['optionDcf']
        print(f"選擇權模式 DCF 腿（只列前後值，不拆解）：${o[0]:.2f} → ${o[1]:.2f}")
    if a.json: json.dump(res, open(a.json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('\n拆解檢查：' + ('通過（三情境 × 三項，四項相加＝總變動；加權目標價與 Excel 一致）' if not fails else '不通過\n  ' + '\n  '.join(fails)))
    return 1 if fails else 0


if __name__ == '__main__':
    sys.exit(main())

# Oracle v0.3.1 B 段目標價變動拆解（已決定事項 12：(b) 實際數／市場資料更新再拆 ① 現價 → ② 同業倍數 → ③ CAPM 輸入；(a)(c)(d)＝0；④ 其餘更新檔＝殘差，應為 0）。
# 數值取自 Excel（LibreOffice 重算；scripts/attrib.py 的 build／read／legs）；每一步在前一步 company.json 上只換該步的欄位：
#   ① 現價：callFacts.priceLast、defaults.eqPx、valuation.price（WACC 股權權重、賣出門檻、ATM 發行價）
#   ② 同業倍數：peers 整段（軟體同業 NTM EV/EBITDA → 傳統事業倍數中位數；CRWV／NBIS 市值只作畫面比較）
#   ③ CAPM 輸入：valuation.rf（beta、kd 無可驗證的新值，未動）
#   ④ 其餘（共識檔、說明文字、總帳參照）＝殘差，應為 0
# 用法：python3 scripts/attrib_v031.py --base 前一版.xlsx --a-json 前一版company.json --b-json 新版company.json [--json 輸出.json]
import sys, os, json, shutil, argparse, tempfile, copy
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import attrib
from attrib import build, read, legs, SCN


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--base', required=True); ap.add_argument('--a-json', required=True); ap.add_argument('--b-json', required=True); ap.add_argument('--json')
    a = ap.parse_args()
    A, B = (json.load(open(p, encoding='utf-8')) for p in (a.a_json, a.b_json))
    S = [('①', '現價', lambda x: (x['callFacts'].update(priceLast=B['callFacts']['priceLast']), x['defaults'].update(eqPx=B['defaults']['eqPx']), x['valuation'].update(price=B['valuation']['price']))),
         ('②', '同業倍數', lambda x: x.update(peers=copy.deepcopy(B['peers']))),
         ('③', 'CAPM 輸入', lambda x: x['valuation'].update(rf=B['valuation']['rf'], capm=copy.deepcopy(B['valuation']['capm']))),
         ('④', '其餘（共識檔、文字）殘差', lambda x: x.update(copy.deepcopy(B)))]
    tmp = tempfile.mkdtemp(prefix='attrib31_')
    try:
        cur, X, names = copy.deepcopy(A), [read(build(a.base, tmp))], ['前一版']
        for i, (m, nm, fn) in enumerate(S):
            fn(cur); p = os.path.join(tmp, f's{i}.json'); json.dump(cur, open(p, 'w', encoding='utf-8'), ensure_ascii=False)
            X.append(read(build(p, tmp))); names.append(nm)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    res, fails = {'stages': names, 'wacc': [x['wacc'] for x in X], 'scenarios': {}}, []
    print('WACC：' + ' → '.join(f"{w:.4%}" for w in res['wacc']))
    for s in (1, 2, 3):
        L = [legs(x, s) for x in X]
        for k, x in enumerate(X):
            if x['dcfMode'] == 1 and abs(L[k]['blend'] - x[s]['blendXl']) > 1e-6: fails.append(f"{SCN[s]}（{names[k]}）：重組加權 {L[k]['blend']:.4f} ≠ Excel {x[s]['blendXl']:.4f}")
        dl = {m: [L[k + 1][m] - L[k][m] for k in range(len(S))] for m in ('dcf', 'ev', 'blend')}
        tot = {m: L[-1][m] - L[0][m] for m in dl}
        for m in dl:
            if abs(sum(dl[m]) - tot[m]) > 0.01: fails.append(f'{SCN[s]} {m}：各步相加 ≠ 總變動')
        res['scenarios'][SCN[s]] = {'levels': L, 'delta': dl, 'total': tot, 'optionDcf': [X[0][s]['dcfOpt'] * (1 + X[0]['wacc']) ** X[0]['targetT'], X[-1][s]['dcfOpt'] * (1 + X[-1]['wacc']) ** X[-1]['targetT']]}
        print(f"\n{SCN[s]}（0 截斷口徑）\n| 項目 | 前一版 | ① 現價 | ② 同業倍數 | ③ CAPM | ④ 殘差 | 新版 | 總變動 |\n|---|---|---|---|---|---|---|---|")
        for m, nm in (('dcf', 'DCF 腿（推到目標價時點）'), ('ev', 'EV/EBITDA 腿'), ('blend', '加權目標價')):
            d = dl[m]; print(f"| {nm} | ${L[0][m]:.2f} | {d[0]:+.2f} | {d[1]:+.2f} | {d[2]:+.2f} | {d[3]:+.2f} | ${L[-1][m]:.2f} | {tot[m]:+.2f} |")
    if a.json: json.dump(res, open(a.json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('\n拆解檢查：' + ('通過（三情境 × 三項，各步相加＝總變動，誤差 < 0.01；加權目標價與 Excel 一致）' if not fails else '不通過\n  ' + '\n  '.join(fails)))
    return 1 if fails else 0


if __name__ == '__main__':
    sys.exit(main())

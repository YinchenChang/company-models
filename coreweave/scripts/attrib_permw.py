# W3（2026-10-08）：v4.5 → v4.6 目標價變動拆解（已決定事項 12(3)；工作單 W3 第 3 步）。
# (a) 時間推移＝0（評價日不變）、(b) 實際數更新＝0、(c) 假設變更＝0（共用輸入未改），全部變動屬 (d) 方法變更（每 MW 改接 Tokenomics）。
# (d) 依序逐項切換（每步只改一項，前一步結果為下一步起點）：
#   ① 每 MW 資本支出 capex＝tokenomics（Σ 新增世代占比 × IF_CapexIT；折舊年限仍用舊輸入 defaults.gpuLife）
#   ② 折舊年限改用 IF_DeprLifeIT（期初在役世代加權、取整數）
#   ③ 營運成本改由下而上 cost＝bottomUp（Tokenomics 電費、IT 維護、人員軟體、稅險＋管銷率）
#   ④ 每 MW 收入改 GPU 小時價格（本次採備案 revenue＝legacy：不切換，＝0）
#   ⑤ MW 口徑換算（meta.mwBasis＝IT：無換算，＝0）
# 另列「單獨切換」：每項只改自己、其他維持 v4.5（顯示順序依賴）。
# 數值取自 Excel（scripts/attrib.py 的 build／read／legs：LibreOffice 開啟、三情境切換「情境選擇」）；各步以 attrib.py 拆解 (a)–(d) 並檢查四項相加＝總變動。
# 用法：python3 scripts/attrib_permw.py --v45 v4.5成品.xlsx [--v46 v4.6.xlsx] [--json 輸出.json]
#   --v46 省略時以目前 repo 建 Excel；並檢查「依序各步相加＝v4.6 − v4.5」（誤差 < 0.01 美元／股）、最後一步＝v4.6。
import sys, os, json, shutil, argparse, tempfile, subprocess
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'scripts'))
from attrib import read, legs, SCN

CO = json.load(open(os.path.join(ROOT, 'company.json'), encoding='utf-8'))
LEG = {'capex': 'legacy', 'cost': 'ebitdaPct', 'revenue': 'legacy'}
NEW = {k: CO['methodology']['perMw'][k] for k in LEG}


def build(tag, pmw, life_tk, tmp):
    """以目前 repo 的程式建一個變體 Excel：methodology.perMw＝pmw；life_tk＝False 時把快照的 IF_DeprLifeIT 標為缺漏（折舊年限沿用舊輸入）。"""
    d = os.path.join(tmp, tag)
    shutil.copytree(ROOT, d, ignore=shutil.ignore_patterns('out', 'dist', '.git', '__pycache__', 'docs'))
    co = json.load(open(os.path.join(d, 'company.json'), encoding='utf-8'))
    co['methodology']['perMw'].update(pmw)
    open(os.path.join(d, 'company.json'), 'w', encoding='utf-8').write(json.dumps(co, indent=1, ensure_ascii=False) + '\n')
    if not life_tk:
        sp = os.path.join(d, co['tokenomics']['snapshotFile'])
        s = json.load(open(sp, encoding='utf-8')); s['items']['IF_DeprLifeIT'] = {'missing': True}
        json.dump(s, open(sp, 'w', encoding='utf-8'), ensure_ascii=False)
    x = os.path.join(d, f'{tag}.xlsx')
    r = subprocess.run(['python3', 'build_xlsx.py', x], cwd=d, capture_output=True, text=True,
                       env=dict(os.environ, LC_ALL='C.UTF-8', LANG='C.UTF-8'))
    if r.returncode: sys.exit(f'建 Excel 失敗（{tag}）：\n{(r.stdout + r.stderr)[-1500:]}')
    return x


def life_of(xlsx):
    from uno_q import Workbook
    with Workbook(xlsx) as wb: return wb.get(wb.cell('輸入與假設', 'GPU 經濟壽命（年）'))


STEPS = [  # (代碼, 名稱, 依序切換後的 perMw, 依序：折舊年限用 Tokenomics？, 單獨切換的 perMw, 單獨：折舊年限用 Tokenomics？)
    ('capex', '① 每 MW 資本支出（Tokenomics IF_CapexIT × 世代組合）', {**LEG, 'capex': NEW['capex']}, False, {**LEG, 'capex': NEW['capex']}, False),
    ('life', '② 折舊年限（IF_DeprLifeIT）', {**LEG, 'capex': NEW['capex']}, True, None, None),
    ('cost', '③ 營運成本改由下而上（Tokenomics＋管銷率）', {**LEG, 'capex': NEW['capex'], 'cost': NEW['cost']}, True, {**LEG, 'cost': NEW['cost']}, False),
    ('revenue', '④ 每 MW 收入改 GPU 小時價格（採備案 legacy：不切換）', {**LEG, 'capex': NEW['capex'], 'cost': NEW['cost'], 'revenue': NEW['revenue']}, True, None, None),
    ('mw', '⑤ MW 口徑換算（IT 口徑：無換算）', None, None, None, None),
]


def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--v45', required=True); ap.add_argument('--v46'); ap.add_argument('--json')
    a = ap.parse_args()
    tmp = tempfile.mkdtemp(prefix='attribpm_')
    try:
        X = {'v45': read(os.path.abspath(a.v45))}
        prev, seq, lifes, built, pkey = 'v45', [], {}, {}, json.dumps([LEG, False], sort_keys=True)
        for k, nm, pmw, ltk, spmw, sltk in STEPS:
            key = pkey if pmw is None else json.dumps([pmw, ltk], sort_keys=True)
            if key == pkey: X[k] = X[prev]  # 與前一步相同設定（④ 採備案、⑤ IT 口徑）：變動 0
            else:
                if key not in built:
                    x = build(k, pmw, ltk, tmp); built[key] = (read(x), x)
                X[k] = built[key][0]; lifes[k] = life_of(built[key][1])
            seq.append(k); prev, pkey = k, key
        X['v46'] = read(os.path.abspath(a.v46)) if a.v46 else X[prev]
        # 單獨切換：只改該項，其他維持 v4.5
        S = {}
        for k, nm, pmw, ltk, spmw, sltk in STEPS:
            if spmw is None: S[k] = X['v45']; continue
            key = json.dumps([spmw, sltk], sort_keys=True)
            if key not in built:
                x = build('alone_' + k, spmw, sltk, tmp); built[key] = (read(x), x)
            S[k] = built[key][0]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    life45 = life_of(os.path.abspath(a.v45))
    res, fails = {'steps': [[k, nm] for k, nm, *_ in STEPS], 'life': {'v45': life45, **lifes}, 'scenarios': {}}, []
    for s in (1, 2, 3):
        L = {k: legs(X[k], s) for k in ['v45'] + seq + ['v46']}
        Ls = {k: legs(S[k], s) for k in seq}
        o = {'v45': L['v45'], 'v46': L['v46'], 'seq': {}, 'alone': {}}
        p = 'v45'
        for k in seq:
            o['seq'][k] = {m: L[k][m] - L[p][m] for m in ('dcf', 'ev', 'blend')}; p = k
            o['alone'][k] = {m: Ls[k][m] - L['v45'][m] for m in ('dcf', 'ev', 'blend')}
        tot = L['v46']['blend'] - L['v45']['blend']
        ssum = sum(v['blend'] for v in o['seq'].values())
        o['total'] = tot; o['abcd'] = {'a': 0.0, 'b': 0.0, 'c': 0.0, 'd': tot}
        if abs(ssum - tot) >= 0.01: fails.append(f'{SCN[s]}：依序各步相加 {ssum:.4f} ≠ 總變動 {tot:.4f}')
        if abs(L[seq[-1]]['blend'] - L['v46']['blend']) >= 0.01: fails.append(f'{SCN[s]}：最後一步 {L[seq[-1]]["blend"]:.4f} ≠ v4.6 {L["v46"]["blend"]:.4f}')
        for k in ('v45', 'v46'):
            if X[k]['dcfMode'] == 1 and abs(L[k]['blend'] - X[k][s]['blendXl']) > 1e-6: fails.append(f'{SCN[s]} {k}：重組加權 ≠ Excel')
        res['scenarios'][SCN[s]] = o
        print(f"\n{SCN[s]}：v4.5 ${L['v45']['blend']:.2f} → v4.6 ${L['v46']['blend']:.2f}（總變動 {tot:+.2f}；(a) 0、(b) 0、(c) 0、(d) {tot:+.2f}）")
        print('| (d) 方法變更逐項 | 依序切換 | 單獨切換 | 順序依賴 |'); print('|---|---|---|---|')
        for k, nm, *_ in STEPS:
            q, al = o['seq'][k]['blend'], o['alone'][k]['blend']
            print(f'| {nm} | {q:+.2f} | {al:+.2f} | {q - al:+.2f} |')
        print(f'| 合計 | {ssum:+.2f} | {sum(v["blend"] for v in o["alone"].values()):+.2f} | |')
    print(f"\n折舊年限：v4.5 {life45:g} 年；各步 {lifes}")
    if a.json: json.dump(res, open(a.json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('\n拆解檢查：' + ('通過（三情境：依序各步相加＝v4.6 − v4.5，誤差 < 0.01 美元／股；最後一步＝v4.6；加權目標價與 Excel 一致）' if not fails else '不通過\n  ' + '\n  '.join(fails)))
    return 1 if fails else 0


if __name__ == '__main__':
    sys.exit(main())

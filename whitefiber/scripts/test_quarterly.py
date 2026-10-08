# v4.4：季度層測試（verify.sh 步驟 8；只用暫存副本，repo 內的 company.json 與共識檔不變）
# 測試 A｜假設實際數：在 out/qtest_actual/ 的副本填入一組假設的 2026Q3 實際數，重建 HTML 與 Excel，
#        以 Python 獨立計算「實際 vs 模型／共識／指引」差距，與 HTML 引擎及 Excel 逐項比對；cmp31 基準情境須全部一致。
# 測試 B｜可移植性（Oracle／OpenAI 型）：副本設 driver.type＝none、刪除季度指引與期間指引、共識檔刪除季度共識，
#        確認建置、重算（0 公式錯誤）、cmp31 皆通過，畫面無錯誤且缺資料欄位顯示「不適用」。
# 用法：python3 scripts/test_quarterly.py；任何一項失敗即以代碼 1 結束。
import json, os, re, shutil, subprocess, sys
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'out')
FAIL = []


def run(cmd, cwd=None):
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if r.returncode:
        raise RuntimeError(' '.join(cmd) + '\n' + r.stdout[-2000:] + r.stderr[-2000:])
    return r.stdout


def copy_repo(name):
    d = os.path.join(OUT, name)
    shutil.rmtree(d, ignore_errors=True)
    shutil.copytree(ROOT, d, ignore=shutil.ignore_patterns('out', 'dist', '.git', '__pycache__'))
    os.makedirs(os.path.join(d, 'out'))
    return d


def build_and_compare(d, label):
    run(['node', os.path.join(d, 'scripts', 'check_quarterly.js'), d])
    run(['python3', os.path.join(d, 'build_html_portable.py'), '4.4', os.path.join(d, 'out', 't.html'), '2026-09-26'])
    x = os.path.join(d, 'out', 't.xlsx')
    run(['python3', os.path.join(d, 'build_xlsx.py'), x])
    rc = json.loads(run(['python3', os.path.join(d, 'scripts', 'recalc.py'), x, '120']).strip().splitlines()[-1])
    if rc.get('total_errors'):
        FAIL.append(f'{label}：重算公式錯誤 {rc["total_errors"]}')
    run(['python3', os.path.join(d, 'xlx.py'), x, '2', os.path.join(d, 'out', 'xl17_2.json')])
    log = run(['node', os.path.join(d, 'cmp31.js'), 'base'], cwd=os.path.join(d, 'out'))
    nok = len(re.findall(r'^OK ', log, re.M)); nbad = len(re.findall(r'^(XX |MISSING)|MISSING XL KEY', log, re.M))
    (FAIL.append if nbad or not nok else (lambda s: None))(f'{label}：cmp31 基準 OK {nok}、不一致 {nbad}')
    print(f'  {label}：check_quarterly 通過；Excel 重算 {rc.get("total_formulas")} 個公式、錯誤 {rc.get("total_errors")}；cmp31 基準 {nok} 項、不一致 {nbad}')
    return json.load(open(os.path.join(d, 'out', 'xl17_2.json'), encoding='utf-8'))


def html_quarterly(d):  # HTML 引擎（基準情境）的季度層結果
    js = ("require(process.argv[1]+'/load_engine.js')(process.argv[1]);const q=structuredClone(DEFAULTS);const d=runFunding(q),p=runValuation(d,q,VAL_DEFAULTS);"
          "const v=quarterlyView(d,p,q);console.log(JSON.stringify({rows:v.rows.map(m=>({key:m.key,label:m.label,q:m.q})),line:v.rows.map((m,i)=>v.line(m,v.fi)),keyLines:v.keyLines,fi:v.fi}))")
    return json.loads(run(['node', '-e', js, d]))


def page_text(html):  # 開啟「資金模型 → 各期收支 → 季度追蹤」與總結頁，回傳（文字, 頁面錯誤）
    from playwright.sync_api import sync_playwright
    errs = []
    with sync_playwright() as p:
        try:
            b = p.chromium.launch()
        except Exception:
            b = p.chromium.launch(executable_path=os.environ.get('CHROMIUM_PATH', '/opt/pw-browsers/chromium'))
        pg = b.new_page(viewport={'width': 1440, 'height': 1000})
        pg.on('pageerror', lambda e: errs.append(str(e)))
        pg.goto('file://' + html); pg.wait_for_timeout(2500)
        summ = pg.inner_text('body')
        click = lambda t: (pg.locator('button:visible').filter(has_text=re.compile('^\\s*' + re.escape(t) + '\\s*$')).first.click(timeout=3000), pg.wait_for_timeout(500))
        pg.get_by_role('button', name='資金模型', exact=True).click(); pg.wait_for_timeout(700)
        click('各期收支'); click('季度追蹤')
        txt = pg.inner_text('body')
        b.close()
    return summ, txt, errs


print('=== 測試 A｜假設 Q3 實際數（暫存副本；repo 內實際數維持空白）')
ACT = {"revenue": 0.046, "adjEbitda": 0.006, "adjOpInc": -0.004, "capex": 0.15, "mw": 12, "source": "假設測試（非真實數字）", "date": "2026-11-12", "tag": "Assumed"}  # WhiteFiber v0.1b：WhiteFiber 量級（US$bn）
RTX = f"{ACT['revenue']:,.2f}bn"
d = copy_repo('qtest_actual')
co = json.load(open(os.path.join(d, 'company.json'), encoding='utf-8'))
QK = co['quarterly']['focus']  # v0.1b：焦點季（原寫死 2026Q3）
co['quarterly']['actuals'][QK] = ACT
json.dump(co, open(os.path.join(d, 'company.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
X = build_and_compare(d, '測試 A')
H = html_quarterly(d)
sys.path.insert(0, d); import calendar_q as _cq  # WhiteFiber v0.1b：共識檔經正規化讀取；無季度共識時為空（差距為不適用）
cons = (_cq.load_consensus(d, co).get('quarterlyEstimates') or {}).get(QK) or {}
G = co['quarterly'].get('guidance', {}).get(QK, {})  # v0.1b：公司未給季度指引時為空（差距與位置皆為不適用）
# 獨立計算（Python）：模型值取 Excel「季度追蹤」焦點列（Q3），共識讀共識檔、指引讀 company.json、實際數為上面的假設值
exp = {}
for m in co['quarterly']['metrics']:
    k = m['key']; u = m['unit']; row = X[f"季度追蹤|焦點｜{m['label']}"]
    model = row[0]
    a = ACT['adjEbitda'] / ACT['revenue'] if k == 'ebitdaMargin' else ACT.get(k)
    c = {'revenue': cons.get('revenue'), 'adjEbitda': cons.get('ebitda'), 'ebitdaMargin': (cons['ebitda'] / cons['revenue']) if cons.get('ebitda') and cons.get('revenue') else None}.get(k)
    g = G.get(k)
    kind = m.get('gap') or ('pt' if u == '%' else 'ratio')  # v4.4 第 4 輪：營收、CapEx＝比例；利潤類與 MW＝差額；EBITDA 率＝百分點
    rel = (lambda x, y: x / y - 1) if kind == 'ratio' else (lambda x, y: x - y)
    exp[k] = dict(kind=kind, am=rel(a, model), ac=rel(a, c) if c is not None else None, ag=rel(a, (g[0] + g[1]) / 2) if g else None,
                  posA=(None if not g else '高於上緣' if a > g[1] else '低於下緣' if a < g[0] else '區間上半部' if a >= (g[0] + g[1]) / 2 else '區間下半部'), actual=a)
    hq = next(r for r in H['rows'] if r['key'] == k)['q'][H['fi']]
    for f in ('am', 'ac', 'ag'):
        e = exp[k][f]; hv = hq[f]; xv = row[{'am': 7, 'ac': 8, 'ag': 9}[f]]
        ok = (e is None and hv is None and not isinstance(xv, (int, float))) or (e is not None and abs(hv - e) < 1e-9 and abs(xv - e) < 1e-9)
        if not ok:
            FAIL.append(f'測試 A {m["label"]} {f}：獨立計算 {e}、HTML {hv}、Excel {xv}')
    if not (isinstance(row[6], (int, float)) and abs(row[6] - a) < 1e-12):
        FAIL.append(f'測試 A {m["label"]} 實際：Excel {row[6]} ≠ {a}')
    if exp[k]['posA'] and hq['posA'] != exp[k]['posA']:
        FAIL.append(f'測試 A {m["label"]} 指引位置：HTML {hq["posA"]} ≠ {exp[k]["posA"]}')
fmt = lambda x: ('−' if x < 0 else '+') + f'{abs(x) * 100:,.1f}%'
want = f"｜實際 ${RTX}（較模型 {fmt(exp['revenue']['am'])}" + (f"、較共識 {fmt(exp['revenue']['ac'])}" if exp['revenue']['ac'] is not None else "") + (f"、{exp['revenue']['posA']}" if exp['revenue']['posA'] else "") + "）"
if want not in H['keyLines'][0]:
    FAIL.append(f'測試 A 驗證點句缺少「{want}」：{H["keyLines"][0]}')
summ, txt, errs = page_text(os.path.join(d, 'out', 't.html'))
if errs or RTX not in txt or want not in summ:
    FAIL.append(f'測試 A 畫面：錯誤 {errs}；季度追蹤含 {RTX}：{RTX in txt}；一頁摘要含實際數句：{want in summ}')
for m in co['quarterly']['metrics']:
    e = exp[m['key']]
    fm = (lambda x: fmt(x)[:-1] + 'pt') if e['kind'] == 'pt' else (lambda x: ('−' if x < 0 else '+') + (f'{abs(x):,.0f} MW' if m['unit'] == 'MW' else f'${abs(x):,.2f}bn')) if e['kind'] == 'diff' else fmt
    print(f"  {m['label']}：實際 {e['actual']:.4g}；較模型 {fm(e['am'])}" + (f"、較共識 {fm(e['ac'])}" if e['ac'] is not None else '') +
          (f"、較指引中點 {fm(e['ag'])}（{e['posA']}）" if e['ag'] is not None else ''))
print('  驗證點句：' + H['keyLines'][0])

print('=== 測試 B｜可移植性：無 MW、無季度指引、無季度共識（暫存副本）')
d = copy_repo('qtest_port')
co = json.load(open(os.path.join(d, 'company.json'), encoding='utf-8'))
Q = co['quarterly']
Q['driver'] = {'type': 'none'}
Q.pop('guidance', None); Q.pop('periodGuidance', None)
Q['consistency'] = [x for x in Q['consistency'] if not x[0].startswith(('quarterly.guidance', 'quarterly.periodGuidance', 'quarterly.driver'))]
json.dump(co, open(os.path.join(d, 'company.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
cf = os.path.join(d, co['meta']['consensusFile'])
cc = json.load(open(cf, encoding='utf-8')); cc.pop('quarterlyEstimates', None)
json.dump(cc, open(cf, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
X = build_and_compare(d, '測試 B')
summ, txt, errs = page_text(os.path.join(d, 'out', 't.html'))
row = X['季度追蹤|焦點｜營收']
na = sum(1 for v in row if v == '不適用')
if errs or '不適用' not in txt or na < 3:
    FAIL.append(f'測試 B 畫面：錯誤 {errs}；季度追蹤含「不適用」：{"不適用" in txt}；Excel 焦點營收列「不適用」{na} 格')
print(f'  焦點營收列（Excel）：{row}')
print(f'  畫面錯誤：{errs or "無"}')

print('=== 結果')
if FAIL:
    print('test_quarterly：失敗 ' + str(len(FAIL)) + ' 項'); [print('  ' + x) for x in FAIL]; sys.exit(1)
print('test_quarterly：全部通過（測試 A 假設實際數、測試 B 可移植性）')

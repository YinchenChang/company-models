# MAG v0.1b：共用引擎 C8 (a)–(f) 機制測試（verify.sh 步驟 8d；只用暫存副本 out/mtest/，repo 內 company.json 不變，不進成品）。
# Amazon 的這些輸入為 0 或空清單；本測試把副本的輸入改為非零（虛構測試值，不是任何公司的真實數字），確認 HTML 與 Excel 兩邊都能運作且 cmp31 一致：
#   (a) 未分攤公司層費用：legacyBiz 新增 explicit 線（營收 0、EBITDA 負值陣列）
#   (b) 硬體銷售：legacyBiz 新增 explicit 線（各期營收輸入 × EBITDA 率；不走 MW × 每 MW）
#   (c) 多筆持股：valuation.holdings 新增一筆上市持股（不折價）
#   (d) 租用算力排程：leases.rentedCompute 一筆（起租、年期、年租金）
#   (e) 強制轉換特別股：debt.convertibles 新增一檔 mand＝true（Oracle 模組，評價日後到期轉股）＋特別股股利
#   (f) 未起租租賃的營業／融資拆分：leases.uncommenced.opShare 改 0.5
#   另：回購基準非零（瀑布「減少回購」步驟）
# 檢查：建置、重算 0 錯誤、反向 DCF／3 × 3 求解、cmp31 基準全部一致，且 HTML 引擎中上列機制的數字確實非零。
# 用法：python3 scripts/test_mag_mechanisms.py；任何一項失敗即以代碼 1 結束。
import json, os, re, shutil, subprocess, sys
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'out')
FAIL = []


def run(cmd, cwd=None, ok=(0,)):
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if r.returncode not in ok:
        raise RuntimeError(' '.join(cmd) + '\n' + r.stdout[-2000:] + r.stderr[-2000:])
    return r


_h0 = json.loads(run(['node', '-e', "require(process.argv[1]+'/load_engine.js')(process.argv[1]);console.log(JSON.stringify(holdValQ(VAL_DEFAULTS)))", ROOT]).stdout)  # 原始持股價值（測試前）
d = os.path.join(OUT, 'mtest')
shutil.rmtree(d, ignore_errors=True)
shutil.copytree(ROOT, d, ignore=shutil.ignore_patterns('out', 'dist', '.git', '__pycache__'))
os.makedirs(os.path.join(d, 'out'))
P = os.path.join(d, 'company.json')
co = json.load(open(P, encoding='utf-8'))
LB = co['defaults']['legacyBiz']['lines']
LB.append({'key': 'corpTest', 'label': '測試｜未分攤公司層費用', 'kind': 'explicit', 'peer': 'cloud', 'ytd': 0, 'rev': [0, 0, 0, 0, 0],
           'ebitda': [-1.5, -3.2, -3.4, -3.6, -3.8], 'm0': 0, 'mLT': None, 'oa': 0, 'cx': 0})
LB.append({'key': 'hwTest', 'label': '測試｜硬體銷售', 'kind': 'explicit', 'peer': 'cloud', 'ytd': 0.8, 'rev': [1.2, 3.0, 4.5, 5.5, 6.0],
           'm0': 0.3, 'mLT': 0.25, 'oa': 0, 'cx': 0.05})
co['valuation']['holdings'].append(['測試｜上市持股', 120.0, 0.05, '虛構測試值', True])
co['leases']['rentedCompute'] = [{'name': '測試租約', 'start': '2026-10', 'years': 5, 'annualRent': 2.4, 'mw': 300, 'use': 'internal'}]
co['debt']['convertibles'].append(['測試｜強制轉換特別股', 5.0, 0.06, '2029-05', 1.0, 300.0, '虛構測試值', True])
co['defaults']['dividend']['preferred'] = [0.15, 0.3, 0.3, 0.13, 0]
co['leases']['uncommenced']['opShare'] = 0.5
co['defaults']['buyback'] = {'annual': 40.0, 'floorShare': 0.25, 'note': '虛構測試值'}
co['pricing']['customCapexFactor'] = 0.8  # v0.1 交付前修訂：自研晶片 IT 資本支出係數
co['capexModel']['rentedExt'] = {'open': 200, 'path': [260, 320, 380, 380, 380], 'rentMW': 9.0, 'note': '虛構測試值'}  # MAG v0.1b r3（C20）：租用對外 MW
co['defaults']['dividend']['perShareQ'] = 0.05; co['defaults']['dividend']['growth'] = 0.08  # MAG v0.1b r3（C19）：股利每股成長
json.dump(co, open(P, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

run(['node', os.path.join(d, 'scripts', 'check_quarterly.js'), d])
run(['python3', os.path.join(d, 'build_html_portable.py'), '0.1', os.path.join(d, 'out', 't.html'), '2026-10-08'])
x = os.path.join(d, 'out', 't.xlsx')
for k in range(2):
    run(['python3', os.path.join(d, 'build_xlsx.py'), x])
    rc = json.loads(run(['python3', os.path.join(d, 'scripts', 'recalc.py'), x, '120']).stdout.strip().splitlines()[-1])
    if rc.get('total_errors'): FAIL.append(f'重算公式錯誤 {rc["total_errors"]}')
    r = run(['python3', os.path.join(d, 'scripts', 'rv_solve.py'), x], ok=(0, 3))
    if r.returncode == 0: break
run(['python3', os.path.join(d, 'xlx.py'), x, '2', os.path.join(d, 'out', 'xl17_2.json')])
log = run(['node', os.path.join(d, 'cmp31.js'), 'base'], cwd=os.path.join(d, 'out')).stdout
nok = len(re.findall(r'^OK ', log, re.M)); bad = re.findall(r'^(?:XX |MISSING).*|.*MISSING XL KEY.*', log, re.M)
if bad or not nok: FAIL.append(f'cmp31 基準 OK {nok}、不一致 {len(bad)}：' + '；'.join(b[:120] for b in bad[:5]))
js = ("require(process.argv[1]+'/load_engine.js')(process.argv[1]);const q=structuredClone(DEFAULTS);const d=runFunding(q),p=runValuation(d,q,VAL_DEFAULTS);"
      "const L=k=>d.lg.lines.find(x=>x.key===k);console.log(JSON.stringify({corp:L('corpTest').ebitda,hw:L('hwTest').rev,hwEb:L('hwTest').ebitda,rent:d.years.map(y=>y.rentedCompute),"
      "bbPlan:d.years.map(y=>y.buybackPlan),bb:d.years.map(y=>y.buyback),bbCut:d.years.map(y=>y.buybackCut),hold:holdValQ(VAL_DEFAULTS),mand:CVN.filter(n=>n.mand).length,"
      "pref:d.years.map(y=>y.divPref),leaseOff:d.years.map(y=>y.leaseCashOff),offAll:d.years.map(y=>y.offLeaseAll),tgt:p.call.blended}))")
h = json.loads(run(['node', '-e', js, d]).stdout)
chk = [('(a) 公司層費用 EBITDA', h['corp'][1] < 0), ('(b) 硬體銷售營收', h['hw'][1] > 0 and h['hwEb'][1] > 0),
       ('(c) 上市持股不折價（持股價值＝原值＋120 × 5%）', abs(h['hold'] - (_h0 + 6.0)) < 1e-9),
       ('(d) 租用算力租金', sum(h['rent']) > 0 and h['rent'][0] > 0), ('(e) 強制轉換特別股與特別股股利', h['mand'] >= 1 and sum(h['pref']) > 0),
       ('(f) 未起租只扣融資部分（50%）', all(abs(a - 0.5 * b) < 1e-9 for a, b in zip(h['leaseOff'], h['offAll'])) and sum(h['offAll']) > 0),
       ('回購（計畫非零）', sum(h['bbPlan']) > 0)]
for nm, ok in chk:
    print(f"  {'OK ' if ok else 'XX '}{nm}")
    if not ok: FAIL.append(nm)
print(f"  cmp31 基準 {nok} 項、不一致 {len(bad)}；加權目標價 {h['tgt']:.2f}；回購 {[round(v, 2) for v in h['bb']]}（被迫減少 {[round(v, 2) for v in h['bbCut']]}）；租用租金 {[round(v, 2) for v in h['rent']]}")
if FAIL:
    print('test_mag_mechanisms：失敗 ' + str(len(FAIL)) + ' 項'); [print('  ' + f) for f in FAIL]; sys.exit(1)
print('test_mag_mechanisms：C8 (a)–(f)＋回購 非零時 HTML 與 Excel 一致（全部通過）')

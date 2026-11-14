# WhiteFiber v0.2（自 Oracle v0.3）：tkAnchor 模式下，把預設情境（defaults.scenario）的錨定每 MW 年收入、期初可計費 MW 與可計費 MW 路徑回寫 company.json →
# defaults.m.revMW、defaults.billableOpen、defaults.m.billable（HTML 開啟時的預設輸入；build_xlsx.py 會檢查三者等於 tkanchor.py 的計算值）。
# 用法：python3 scripts/tkanchor_sync.py [--write]（不帶參數只比對；不一致時以代碼 1 結束）
import json, os, sys
REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
import calendar_q, tkanchor
co = calendar_q.load(REPO)
if (co.get('methodology', {}).get('perMw') or {}).get('revenue') != 'tkAnchor':
    print('legacy：不需回寫'); sys.exit(0)
_, tk = tkanchor.load(co)
sc = co['defaults']['scenario']
R = tkanchor.compute(co, tk, sc, co['periodYears'])
rev = [round(x, 10) for x in R['rev']]; bo = tkanchor.bopen(co, R['rev'][0]); bil = tkanchor.billable_path(co, sc, co['periodYears'], bo)
cur = co['defaults']
same = all(abs(a - b) < 1e-9 for a, b in zip(cur['m']['revMW'], rev)) and cur['billableOpen'] == bo and cur['m']['billable'] == bil
print(f"錨定每 MW 年收入（{sc}）：{[round(x * 1e3 / tkanchor.uf(co), 3) for x in rev]}；期初可計費 MW {bo}；可計費路徑 {bil}；company.json {'一致' if same else '不一致'}")
if same: sys.exit(0)
if '--write' not in sys.argv: sys.exit(1)
p = os.path.join(REPO, 'company.json')
raw = json.load(open(p, encoding='utf-8'))
raw['defaults']['m']['revMW'] = rev; raw['defaults']['billableOpen'] = bo; raw['defaults']['m']['billable'] = bil
json.dump(raw, open(p, 'w', encoding='utf-8'), ensure_ascii=False, indent=1); open(p, 'a').write('')
print('已回寫 company.json')

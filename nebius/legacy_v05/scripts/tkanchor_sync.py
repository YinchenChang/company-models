# v0.2a：tkAnchor 模式下，把預設情境（defaults.scenario）的錨定每 MW 年收入與期初可計費 MW 回寫 company.json →
# defaults.m.revMW、defaults.billableOpen（HTML 開啟時的預設輸入；build_xlsx.py 會檢查兩者等於 tkanchor.py 的計算值）。
# 用法：python3 scripts/tkanchor_sync.py [--write]（不帶參數只比對；不一致時以代碼 1 結束）
import json, os, re, sys
REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
import calendar_q, tkanchor
co = calendar_q.load(REPO)
if (co.get('methodology', {}).get('perMw') or {}).get('revenue') != 'tkAnchor':
    print('legacy：不需回寫'); sys.exit(0)
_, tk = tkanchor.load(co)
R = tkanchor.compute(co, tk, co['defaults']['scenario'], co['periodYears'])
rev = [round(x, 10) for x in R['rev']]; bo = tkanchor.billable_open(co, R['rev'][0])
cur = co['defaults']
same = all(abs(a - b) < 1e-9 for a, b in zip(cur['m']['revMW'], rev)) and cur['billableOpen'] == bo
print(f"錨定每 MW 年收入（{co['defaults']['scenario']}）：{[round(x * 1e3, 3) for x in rev]}；期初可計費 MW {bo}；company.json {'一致' if same else '不一致'}")
if same: sys.exit(0)
if '--write' not in sys.argv: sys.exit(1)
p = os.path.join(REPO, 'company.json'); s = open(p, encoding='utf-8').read()
i = s.index('\n "defaults": {'); j = s.index('\n "valuation": {', i)
d = s[i:j]
d = re.sub(r'("billableOpen": )\d+', lambda m: m.group(1) + str(bo), d, count=1)
k = d.index('\n  "m": {'); a = d.index('"revMW": [', k); b = d.index(']', a)
ind = '\n    '
d = d[:a] + '"revMW": [' + ind + (',' + ind).join(repr(x) for x in rev) + '\n   ' + d[b:]
s = s[:i] + d + s[j:]
json.loads(s); open(p, 'w', encoding='utf-8').write(s); print('已回寫 company.json')

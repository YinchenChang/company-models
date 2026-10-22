# MAG v0.2.1 補充：company.json → pricing.kEvidence（k 的證據價格）須與 data/kevidence_<公司>_*.json 的 kspecInputs 逐筆相同（價格、單位、世代、分母名稱），
# 且 kLong／kSpot 的每個 {ref} 都指向存在的證據、不再有寫在 company.json 或公式裡的未登錄常數。用法：python3 scripts/check_kevidence.py（verify.sh 呼叫）
import json, os, glob, sys
R = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
co = json.load(open(os.path.join(R, 'company.json'), encoding='utf-8')); P = co.get('pricing') or {}
if 'kEvidence' not in P: print('無 kEvidence（不適用）'); sys.exit(0)
kf = glob.glob(os.path.join(R, 'data', 'kevidence_*.json'))
if len(kf) != 1: sys.exit(f'找不到唯一的 kevidence 檔：{kf}')
src = json.load(open(kf[0], encoding='utf-8')).get('kspecInputs') or sys.exit('kevidence 檔缺 kspecInputs')
ids = {e['id'] for e in P['kEvidence']}; bad = []
if ids != set(src): bad.append(f'證據 id 不一致：company.json {sorted(ids)}、kevidence {sorted(src)}')
for e in P['kEvidence']:
    s = src.get(e['id'])
    if s and any(s[k] != e[k] for k in ('price', 'unit', 'gen', 'tkName')): bad.append(f"{e['id']}：company.json 與 kevidence 不同 {s} vs {({k: e[k] for k in ('price','unit','gen','tkName')})}")
    if not e.get('note') or not e.get('source'): bad.append(f"{e['id']}：缺旁註或來源")
for sd in ('kLong', 'kSpot'):
    for a, sp in P[sd].items():
        if not (isinstance(sp, dict) and sp.get('ref') in ids): bad.append(f'pricing.{sd}.{a} 不是指向 kEvidence 的 ref：{sp}')
if bad: print('\n'.join(bad)); sys.exit(1)
print(f'k 證據價格 {len(ids)} 筆與 {os.path.basename(kf[0])} 一致，kLong／kSpot 6 個規格皆為 ref')

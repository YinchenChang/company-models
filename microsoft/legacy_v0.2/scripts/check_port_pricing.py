# MAG v0.2.1 補充：以 scripts/port_company_json.py 重產 company.json，其 pricing 區塊須與現行 company.json 完全相同（k 規格、證據價格、說明文字）。
# 用法：python3 scripts/check_port_pricing.py [Amazon company.json，預設 ../amazon/company.json]
import json, os, subprocess, sys, tempfile
R = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
amz = sys.argv[1] if len(sys.argv) > 1 else os.path.join(R, '..', 'amazon', 'company.json')
cur = json.load(open(os.path.join(R, 'company.json'), encoding='utf-8'))
snap = os.path.join(R, cur['tokenomics']['snapshotFile'])
with tempfile.TemporaryDirectory() as t:
    out = os.path.join(t, 'company.json')
    r = subprocess.run(['python3', os.path.join(R, 'scripts', 'port_company_json.py'), amz, os.path.join(R, 'company.json'), snap, out], capture_output=True, text=True)
    if r.returncode: sys.exit('port_company_json.py 失敗：\n' + r.stderr[-800:])
    new = json.load(open(out, encoding='utf-8'))
bad = [k for k in sorted(set(new['pricing']) | set(cur['pricing'])) if new['pricing'].get(k) != cur['pricing'].get(k)]
if bad: sys.exit('重產後 pricing 與現行 company.json 不同：' + '、'.join(bad))
print('port_company_json.py 重產的 pricing 與現行 company.json 完全相同（%d 個欄位）' % len(cur['pricing']))

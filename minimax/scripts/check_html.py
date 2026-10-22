# 驗收：HTML 與 Excel 一致、無外部引用。
# 用法：python3 scripts/check_html.py 已重算.xlsx 成品.html
# ①每個 <span class="x" data-x="鍵|年或欄序|格式">…</span> 的內文必須等於由 Excel 重算值經同一格式函數得到的字串；
# ②HTML 不得含外部引用（http(s)://、src=、<link、@import、url(）；③標記數字至少 N 個；④版本區的 Tokenomics 版本／提交／SHA-256 與快照一致。
import json, os, re, sys, html as H
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'scripts'))
from xlvals import load, val, FM

xl, page = sys.argv[1:3]
D = load(xl)
txt = open(page, encoding='utf-8').read()
bad = []
spans = re.findall(r'<span class="x" data-x="([^"]*)">(.*?)</span>', txt, flags=re.S)
for dx, inner in spans:
    key, y, fm = dx.split('|')
    got = H.unescape(inner)
    try:
        exp = FM[fm](val(D, key, y or None))
    except Exception as e:
        bad.append(f'{dx}: 取值失敗 {e}'); continue
    if got != exp: bad.append(f'{dx}: HTML「{got}」≠ Excel「{exp}」')
if len(spans) < 150: bad.append(f'標記數字只有 {len(spans)} 個（至少 150）')
for pat in ('http://', 'https://', ' src=', '<link', '@import', 'url('):
    if pat in txt: bad.append(f'含外部引用樣式：{pat}')
cur = open(os.path.join(ROOT, 'data', 'tokenomics_snapshot_CURRENT'), encoding='utf-8').read().strip()
snap = json.load(open(os.path.join(ROOT, 'data', cur), encoding='utf-8'))
for k in ('version', 'commit', 'file_sha256'):
    if snap[k] not in H.unescape(txt): bad.append(f'版本區缺 Tokenomics {k}：{snap[k]}')
print(f'HTML 標記數字 {len(spans)} 個｜不符 {len(bad)}')
if bad:
    print('::error::' + '；'.join(bad[:20])); sys.exit(1)
print('OK')

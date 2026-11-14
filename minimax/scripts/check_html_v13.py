# 個股模型頁（v1.3）驗收：每個帶 data-ref 的數字（含圖上數值標籤）與 xlsx 重算值一致；文字標籤與儲存格文字相同；無外部引用。
# 用法：python3 scripts/check_html_v13.py 已重算.xlsx 個股模型.html
import html as H
import re
import sys

import openpyxl

xl, page = sys.argv[1:3]
wb = openpyxl.load_workbook(xl, data_only=True)
txt = open(page, encoding='utf-8').read()
scan = txt.replace('xmlns="http://www.w3.org/2000/svg"', '')   # inline SVG 命名空間不是外部引用
bad, n = [], 0
for m in re.finditer(r'<(span|text)([^>]*)>(.*?)</\1>', txt, flags=re.S):
    a = dict(re.findall(r'([\w-]+)="([^"]*)"', m.group(2)))
    if 'data-ref' not in a:
        continue
    sh, c = H.unescape(a['data-ref']).split('!')
    v = wb[sh][c].value
    n += 1
    if 'data-v' in a:
        dv = float(H.unescape(a['data-v']))
        if not isinstance(v, (int, float)) or abs(dv - v) > 1e-9 * max(1, abs(v)):
            bad.append(f'{a["data-ref"]}: HTML {dv} ≠ xlsx {v}')
    else:
        if H.unescape(re.sub(r'<[^>]+>', '', m.group(3))) != str(v):
            bad.append(f'{a["data-ref"]}: 文字不符')
for pat in ('http://', 'https://', ' src=', '<link', '@import', '<script'):
    if pat in scan:
        bad.append(f'含外部引用樣式：{pat}')
print(f'個股模型頁 data-ref {n} 個｜不符 {len(bad)}')
if bad or n < 100:
    print('::error::' + '；'.join(bad[:20] or ['data-ref 少於 100']))
    sys.exit(1)
print('OK')

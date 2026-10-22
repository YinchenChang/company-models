# v4.1：LibreOffice 重算後，把模擬運算表還原為 Excel 的 dataTable 公式。
# LibreOffice 能正確計算 Excel 的模擬運算表（t="dataTable"），但存檔時會把每一格改寫成 TABLE(公式列, 輸入格, 輸入值) 一般公式，
# Excel 不認得這個函式。這裡把整塊 TABLE() 儲存格合併回一個 <f t="dataTable" ref=… r1=…/>（只放在左上格），保留 LibreOffice 算出的值。
# 只支援單變數、輸入值排成一欄的運算表（本模型的情境區間運算表）；其他形式會直接報錯，不猜。
# 用法：python3 fix_datatable.py 檔案.xlsx
import sys, zipfile, re, shutil
from openpyxl.utils import column_index_from_string as ci, get_column_letter as cl

CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)>(.*?)</c>', re.S)
TABLE = re.compile(r'<f[^>]*>TABLE\(\$?([A-Z]+)\$?(\d+),\$?([A-Z]+)\$?(\d+),\$?([A-Z]+)\$?(\d+)\)</f>')


def fix(x):
    hits = []
    for m in CELL.finditer(x):
        t = TABLE.search(m.group(4))
        if t: hits.append((m, t))
    if not hits:
        return x, None
    cols = [ci(m.group(1)) for m, _ in hits]; rows = [int(m.group(2)) for m, _ in hits]
    c0, c1, r0, r1 = min(cols), max(cols), min(rows), max(rows)
    assert len(hits) == (c1 - c0 + 1) * (r1 - r0 + 1), 'TABLE() 儲存格不是一個完整的矩形區塊'
    inputs = {(t.group(3), t.group(4)) for _, t in hits}
    assert len(inputs) == 1, f'輸入格不唯一：{inputs}'
    for m, t in hits:  # 單變數、輸入值排成一欄：公式列在區塊上一列、輸入值在區塊左一欄
        assert int(t.group(2)) == r0 - 1 and t.group(1) == m.group(1), '公式列位置不符（只支援輸入值排成一欄）'
        assert ci(t.group(5)) == c0 - 1 and t.group(6) == m.group(2), '輸入值位置不符（只支援輸入值排成一欄）'
    ref = f'{cl(c0)}{r0}:{cl(c1)}{r1}'
    inp = ''.join(inputs.pop())
    top = f'{cl(c0)}{r0}'
    out, last = [], 0
    for m, t in hits:
        body = TABLE.sub('', m.group(4))
        if f'{m.group(1)}{m.group(2)}' == top:
            body = f'<f t="dataTable" ref="{ref}" dt2D="0" dtr="0" r1="{inp}"/>' + body
        out.append(x[last:m.start()] + f'<c r="{m.group(1)}{m.group(2)}"{m.group(3)}>{body}</c>')
        last = m.end()
    out.append(x[last:])
    return ''.join(out), (ref, inp)


src = sys.argv[1]; tmp = src + '.tmp'
zin = zipfile.ZipFile(src); zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
done = []
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename.startswith('xl/worksheets/sheet') and b'TABLE(' in data:
        x, info = fix(data.decode())
        if info: done.append((item.filename, *info)); data = x.encode()
    zout.writestr(item, data)
zout.close(); zin.close(); shutil.move(tmp, src)
if not done: sys.exit('找不到 TABLE() 儲存格：模擬運算表遺失（建置或重算有問題）')
for f, ref, inp in done: print(f'datatable fixed: {f} {ref} 輸入格 {inp}')

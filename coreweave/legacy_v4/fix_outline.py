# LibreOffice 重算後，補回 Excel 群組按鈕位置（outlinePr）與收合狀態。
# v3.4 修正：<sheetPr> 子元素必須依 OOXML 順序 tabColor → outlinePr → pageSetUpPr，
# 否則 Excel 會回報「檔案損毀，無法開啟」（LibreOffice／openpyxl 不檢查此順序）。
import sys, zipfile, re, json, shutil, os
src = sys.argv[1]; tmp = src + '.tmp'
# outline.json 由 build_xlsx.py 寫在 repo 根目錄的 out/；可用第 2 個參數另指定
OUT = json.load(open(sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out', 'outline.json')))
OUTLINE_PR = '<outlinePr summaryBelow="0" summaryRight="1"/>'
ORDER = ['tabColor', 'outlinePr', 'pageSetUpPr']

def fix_sheetpr(x):
    """移除既有 outlinePr，再依 schema 順序重排 sheetPr 的子元素。"""
    m = re.search(r'<sheetPr([^>]*?)(/>|>(.*?)</sheetPr>)', x, re.S)
    if not m:  # 沒有 sheetPr：在 <worksheet ...> 後新增
        return re.sub(r'(<worksheet[^>]*>)', r'\1<sheetPr>' + OUTLINE_PR + '</sheetPr>', x, count=1)
    attrs, inner = m.group(1), (m.group(3) or '')
    kids = re.findall(r'<(\w+)\b[^>]*?(?:/>|>.*?</\1>)', inner, re.S)
    full = re.findall(r'(<(\w+)\b[^>]*?(?:/>|>.*?</\2>))', inner, re.S)
    parts = {k: el for el, k in full if k != 'outlinePr'}
    parts['outlinePr'] = OUTLINE_PR
    unknown = [k for k in parts if k not in ORDER]
    assert not unknown, f'sheetPr 有未預期子元素：{unknown}'
    new = '<sheetPr' + attrs + '>' + ''.join(parts[k] for k in ORDER if k in parts) + '</sheetPr>'
    return x[:m.start()] + new + x[m.end():]

zin = zipfile.ZipFile(src)
wbx = zin.read('xl/workbook.xml').decode()
rels = zin.read('xl/_rels/workbook.xml.rels').decode()
sheets = re.findall(r'<sheet name="([^"]+)" sheetId="\d+"[^>]*r:id="([^"]+)"', wbx)
rmap = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
rmap.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
name2file = {n: 'xl/' + rmap[rid].lstrip('/').replace('xl/', '') for n, rid in sheets}
file2name = {f: n for n, f in name2file.items()}

zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
# [Content_Types].xml 放第一個（Excel 慣例）
items = sorted(zin.infolist(), key=lambda i: i.filename != '[Content_Types].xml')
for item in items:
    data = zin.read(item.filename)
    n = file2name.get(item.filename)
    if n in OUT:
        x = fix_sheetpr(data.decode())
        for (hr, lv, col) in OUT[n]:
            if col:
                x = re.sub(r'(<row r="%d"[^>]*?)collapsed="false"' % hr, r'\1collapsed="true"', x, count=1)
                if not re.search(r'<row r="%d"[^>]*collapsed="true"' % hr, x):
                    x = re.sub(r'(<row r="%d")' % hr, r'\1 collapsed="true"', x, count=1)
        data = x.encode()
    zout.writestr(item, data)
zout.close(); zin.close(); shutil.move(tmp, src); print('outline fixed')

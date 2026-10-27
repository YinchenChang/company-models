# MAG v0.2.1：由舊、新 Excel 的 xl_diff 結果自動產生「預期差異清單」（--expect 用），原因欄＝所屬工作表與列名稱。
# 用法：python3 scripts/gen_expect.py 舊.xlsx 新.xlsx 輸出.txt "原因前綴" ["表頭註解…"]
# 產生後仍須人工確認：清單只是把差異逐格列出，不判斷差異是否合理；合理性看報告的目標價變動拆解與不變性驗收（步驟 ①＝0）。
import sys, subprocess, re, os, openpyxl
old, new, out, why = sys.argv[1:5]; hdr = sys.argv[5:]
here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
def run(*flags):
    r = subprocess.run(['python3', 'xl_diff.py', old, new, '--by-label', '--all', *flags], cwd=here, capture_output=True, text=True).stdout.split('\n')
    return r
wb = openpyxl.load_workbook(new)
cells, renames = {}, []
for flags in (('--values',), ()):
    for ln in run(*flags):
        if '新版缺少此列：' in ln:
            sh = ln.split(' ')[0]; continue
        m = re.match(r'^(\S+) ([A-Z]{1,3}\d+) ', ln)
        if m and m.group(1) in wb.sheetnames:
            sh, c = m.group(1), m.group(2); ws = wb[sh]
            lab = ws.cell(row=ws[c].row, column=1).value
            cells[f'{sh}!{c}'] = f'{why}｜{str(lab or "").strip()[:50]}'
# 版本字串列（A1／A2 內含版本、日期）：舊版標籤 → 新版標籤
wo = openpyxl.load_workbook(old)
for sh, c in (('導覽', 'A1'), ('Tokenomics_取數', 'A2')):
    if wo[sh][c].value != wb[sh][c].value: renames.append(f'列 {wo[sh][c].value} => {wb[sh][c].value}'); cells[f'{sh}!{c}'] = f'{why}｜版本字串'
with open(out, 'w', encoding='utf-8') as f:
    for h in hdr: f.write('# ' + h + '\n')
    f.write('# 列改名（同一區段內）\n'); f.write('\n'.join(renames) + '\n')
    f.write('# 逐格（工作表!儲存格  來源｜列名稱）\n')
    for k, v in cells.items(): f.write(f'{k}  {v}\n')
print(len(cells), '格；列改名', len(renames))

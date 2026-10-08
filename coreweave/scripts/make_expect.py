# W3（2026-10-08）：升版驗收的預期差異清單產生器（已決定事項 12(4)）。
# 對前一版成品與新版 Excel 跑 xl_diff.py --by-label --all（值與公式兩種），把每個差異格依「工作表＋列名稱」歸類：
#   (1) 下游工作表（RULES['downstream']）：每 MW 資本支出／營運成本改變後的連動結果，整張工作表允許變動（列名稱寫入原因）；
#   (2) 「輸入與假設」等輸入頁：只允許 RULES['inputRows'] 列名的列（每 MW 方法直接改動的格），其餘一律不列入清單（保持為差異，verify 失敗）；
#   (3) 同業頁：只允許列名稱含 RULES['peerRowKey'] 的列（本公司那一列的倍數隨模型 EBITDA 改變）。
#   列名稱因數字改變而改名者（例如反向 DCF 矩陣的「$34.0m（目前）」），依同區段、同一尾碼配對，寫成「列 舊 => 新」。
# 不在規則內的差異：列在 stdout 的「未歸類」並以代碼 1 結束（不寫入清單；交給 verify 判定失敗）。
# 用法：python3 scripts/make_expect.py 前一版.xlsx 新版.xlsx 規則.json 輸出清單.txt
import sys, os, re, json, subprocess
from openpyxl import load_workbook
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
old, new, rules, out = sys.argv[1:5]
R = json.load(open(rules, encoding='utf-8'))


def run(*a, exp=None):
    r = subprocess.run(['python3', os.path.join(ROOT, 'xl_diff.py'), old, new, '--by-label', '--all', *a] + ([f'--expect={exp}'] if exp else []),
                       capture_output=True, text=True)
    return r.stdout.splitlines()


wb = load_workbook(new)
lab = lambda sh, coord: str(wb[sh][f"A{re.sub(r'[A-Z]+', '', coord)}"].value or '')
cells, unk = {}, []
# 第 1 輪：找出因數字改變而改名的列（新版缺少的列），依同工作表、同一尾碼配對新增列
renames = []
for ln in run('--values'):
    m = re.match(r'^(\S+) 第 (\d+) 列 新版缺少此列：(.*)$', ln.strip())
    if m: renames.append((m.group(1), m.group(3).strip()))
ren_lines = []
if renames:
    added = [l.strip() for l in run('--values') if re.match(r'^  \S+ 第 \d+ 列 ', l)]
    for sh, o in renames:
        suf = re.search(r'（[^（）]*）$', o)
        cand = [re.sub(r'^\S+ 第 \d+ 列 ', '', a) for a in added if a.startswith(sh + ' ') and suf and a.endswith(suf.group(0))]
        if len(cand) == 1: ren_lines.append(f"列 {o} => {cand[0]}")
        else: unk.append(f"{sh}：列「{o}」改名無法唯一配對（候選 {cand}）")
tmp = out + '.renames'
open(tmp, 'w', encoding='utf-8').write('\n'.join(ren_lines) + '\n')
# 第 2 輪：套用列改名後，值與公式的全部差異逐格歸類
for mode in (['--values'], []):
    lines = run(*mode, exp=tmp)
    for ln in lines[1:]:
        if ln.startswith(('預期差異', '清單中', '新增工作表', '以列名稱', '新增列', '  ', '常數改為', '文字改為', '略過的')): break
        p = ln.split(' ', 2)
        if len(p) < 2 or not re.fullmatch(r'[A-Z]+\d+', p[1]): unk.append(ln); continue
        sh, c = p[0], p[1]; L = lab(sh, c)
        if sh in R['downstream']: why = f"下游（{R['downstream'][sh]}）：{L[:40]}"
        elif sh in R['inputRows'] and any(L.startswith(k) for k in R['inputRows'][sh]):
            why = next(v for k, v in R['inputRows'][sh].items() if L.startswith(k))
        elif sh in R.get('peerSheets', []) and any(k in L for k in R['peerRowKey']): why = f"本公司列（倍數隨模型 EBITDA 改變）：{L[:30]}"
        else: unk.append(f"{sh}!{c}「{L[:40]}」：{ln[:120]}"); continue
        cells.setdefault(f"{sh}!{c}", why)
os.remove(tmp)
with open(out, 'w', encoding='utf-8') as f:
    f.write(f"# {R['title']}（scripts/make_expect.py 依 {os.path.relpath(rules, ROOT)} 產生；勿手改）\n")
    f.write("# 規則：下游工作表整張允許變動；輸入頁只允許每 MW 方法直接改動的列；其餘須 0 差異。\n")
    for l in ren_lines: f.write(l + '\n')
    for k in sorted(cells): f.write(f"{k} {cells[k]}\n")
print(f"預期差異清單：{len(cells)} 格、列改名 {len(ren_lines)}；未歸類 {len(unk)}")
for u in unk[:40]: print('  未歸類', u)
sys.exit(1 if unk else 0)

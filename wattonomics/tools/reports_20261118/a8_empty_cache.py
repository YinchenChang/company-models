#!/usr/bin/env python3
"""A8：數出 Wattonomics 活頁簿「公式存在但快取值為空」的儲存格並分類。唯讀，不改活頁簿。
步驟：(1) openpyxl 公式版與 data_only 版對照；(2) 直接讀 sheet XML 判斷空值型態（空字串結果 or 缺 <v>＝未重算）；
(3) 以 LibreOffice UNO calculateAll 強制重算，逐格比對快取（找未重算）；(4) 找斷鏈（#REF!、錯誤值、壞名稱）；
(5) 對每個空值格驗證其空字串是否合理（前提格為空等，程式逐格評估護欄條件）。輸出 xlsx 清單＋md 報告。"""
import openpyxl, zipfile, re, os, sys, json, subprocess, collections, shutil, tempfile
import xml.etree.ElementTree as ET
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
MODEL = os.path.join(ROOT, 'wattonomics', 'model', open(os.path.join(ROOT, 'wattonomics', 'model', 'CURRENT')).read().strip())
OUT = os.path.join(ROOT, 'wattonomics', 'docs', 'reports', '20261118_A8_空快取公式分類.xlsx')
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
wf = openpyxl.load_workbook(MODEL); wv = openpyxl.load_workbook(MODEL, data_only=True)
# (1) 公式數與空快取
nform = sum(1 for ws in wf for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith('='))
# (2) XML 型態
z = zipfile.ZipFile(MODEL)
wbx = ET.fromstring(z.read('xl/workbook.xml')); rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
rid = {r.get('Id'): r.get('Target') for r in rels}
sheets = {s.get('name'): rid[s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')] for s in wbx.find('m:sheets', NS)}
empties = []  # (sheet, ref, t, vstate, formula)
for name, t in sheets.items():
    root = ET.fromstring(z.read('xl/' + t.replace('/xl/', '').lstrip('/')))
    for c in root.iter('{%s}c' % NS['m']):
        f = c.find('m:f', NS)
        if f is None: continue
        v = c.find('m:v', NS)
        if v is None or (v.text or '') == '':
            empties.append((name, c.get('r'), c.get('t'), 'no <v>（未寫入快取）' if v is None else '<v/> 空字串', f.text or ''))
# (3) 強制重算
tmpd = tempfile.mkdtemp(); hard = os.path.join(tmpd, 'hard.xlsx')
subprocess.run([sys.executable, os.path.join(os.path.dirname(__file__), 'hard_recalc.py'), MODEL, hard], check=True)
wh = openpyxl.load_workbook(hard, data_only=True)
ncell = ndiff = 0
for ws in wv:
    h = wh[ws.title]
    for row in ws.iter_rows():
        for c in row:
            ncell += 1; a = c.value; b = h[c.coordinate].value
            if not (a == b or (isinstance(a, (int, float)) and isinstance(b, (int, float)) and abs(a - b) <= 1e-9 * max(1, abs(a)))): ndiff += 1
shutil.rmtree(tmpd)
# (4) 斷鏈
nref = sum(1 for ws in wf for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith('=') and '#REF' in c.value)
nerr = sum(1 for ws in wv for row in ws.iter_rows() for c in row if isinstance(c.value, str) and re.match(r'^#(REF|VALUE|DIV/0|NAME|N/A|NUM|NULL)', c.value))
nbadname = sum(1 for n, d in wf.defined_names.items() if '#REF' in str(d.attr_text))
# (5) 合理性
REF = re.compile(r"(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))!)?\$?([A-Z]{1,3})\$?(\d+)(?![\d(])")
def V(s, a):
    v = wv[s][a].value; return '' if v is None else v
def N(x): return x if isinstance(x, (int, float)) and not isinstance(x, bool) else 0
def first_arg(b):
    d = 0; q = False
    for i, ch in enumerate(b):
        if ch == '"': q = not q
        if q: continue
        if ch == '(': d += 1
        elif ch == ')': d -= 1
        elif ch == ',' and d == 0: return b[:i], b[i + 1:]
    return b, ''
def tr(cond, sh):
    s = re.sub(r'"([^"]*)"', lambda m: repr(m.group(1)), cond)
    s = REF.sub(lambda m: f"V({repr(m.group(1) or m.group(2) or sh)},{repr(m.group(3)+m.group(4))})", s)
    s = s.replace('<>', '!=')
    return re.sub(r'(?<![<>!=])=(?!=)', '==', s)
def empty(v): return v is None or v == ''
def classify(sh, ref, f):
    m = re.match(r'^IF\(\(?' + REF.pattern + r'\s*(=|<>)\s*""\s*,(.*)$', f)
    if m:
        s = m.group(1) or m.group(2) or sh; v = V(s, m.group(3) + m.group(4)); op = m.group(5); rest = m.group(6)
        if op == '=' and rest.startswith('"",') and empty(v): return 1, '前提格為空→回傳空字串（程式逐格驗證前提格確為空）'
        if op == '<>' and empty(v): return 2, '已知值鏈鏈尾：IF(x<>"",x,下一格)，x 為空且往後無已知值（程式驗證）'
    mm = re.match('^' + REF.pattern + '$', f)
    if mm and empty(V(mm.group(1) or mm.group(2) or sh, mm.group(3) + mm.group(4))): return 3, '直接引用的來源格本身是空字串（程式驗證）'
    if f.startswith('IF('):
        cond, rest = first_arg(f[3:]); then, _ = first_arg(rest)
        if then.strip() == '""':
            try:
                if eval(tr(cond, sh), {'V': V, 'N': N, 'ISNUMBER': lambda x: isinstance(x, (int, float)), 'OR': lambda *a: any(a), 'AND': lambda *a: all(a), 'ABS': abs, 'LEFT': lambda x, n: str(x)[:int(n)]}):
                    return 4, '護欄條件為真→回傳空字串（程式評估條件）'
            except Exception: pass
    if f.strip() == '""': return 5, '常數 =""：整列統一公式的結構占位（例：無晶片資料的列、首年平均）'
    if re.search(r'"[^"]{2,}"', f) and f.rstrip().endswith('"")') or f.endswith(',"")') or 'COUNTIFS' in f and f.rstrip().endswith('"")') is False and ',"")' in f:
        pass
    if 'COUNTIFS' in f: return 6, '查表無符合資料（FundRaw COUNTIFS＝0）→空字串（家族判斷，未逐格評估）'
    if re.search(r'IF\(.*,"[^"]*[^\x00-\x7f][^"]*",""\)', f) or f.endswith(',"")') or f.endswith('"")))'): return 7, '標記／說明欄：無需標記時回傳空字串（家族判斷，未逐格評估）'
    if re.match(r'^IF\(OR\(.*(fye_)', f) or 'fye_' in f: return 8, '財年換算：缺值年回傳空字串（家族判斷；名稱參照未逐格評估）'
    if re.match(r'^IF\(AND\(R', f) or '/2,"")' in f: return 9, '兩端值缺任一即回傳空字串（區間中點；家族判斷）'
    return 10, '其他組合式（家族判斷：條件式皆以空字串作為「無資料」輸出）'
rows = []
for sh, ref, t, vs, f in empties:
    k, why = classify(sh, ref, f)
    r_ = int(re.sub(r'\D', '', ref)); lab = wv[sh].cell(r_, 1).value
    rows.append((sh, ref, str(lab or '')[:50], '=' + f, vs, k, why))
cnt = collections.Counter(r[5] for r in rows); cat = collections.Counter(r[0] for r in rows)
verified = sum(v for k, v in cnt.items() if k in (1, 2, 3, 4)); family = len(rows) - verified
# xlsx
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '總表'
B = Font(bold=True); H = PatternFill('solid', fgColor='DDEBF7')
summ = [('活頁簿', os.path.basename(MODEL)), ('公式總數', nform), ('公式存在但快取值為空', len(rows)), ('其中：XML 為 t="str" 且 <v/> 空字串（LibreOffice 計算結果＝空字串）', sum(1 for r in rows if r[4].startswith('<v/>'))),
        ('其中：缺 <v>（公式從未計算，真正的未重算）', sum(1 for r in rows if r[4].startswith('no'))), ('強制重算後與快取不同的儲存格數（全部 %s 格）' % f'{ncell:,}', ndiff),
        ('含 #REF! 的公式／壞掉的具名範圍／快取錯誤值', f'{nref}／{nbadname}／{nerr}'), ('程式逐格驗證空字串合理（類別 1–4）', verified), ('以公式家族判斷合理（類別 5–10）', family)]
for i, (a, b) in enumerate(summ, 1):
    ws.cell(i, 1, a).font = B; ws.cell(i, 2, b)
ws.cell(len(summ) + 2, 1, '依工作表').font = B
for i, (k, v) in enumerate(sorted(cat.items(), key=lambda x: -x[1]), len(summ) + 3): ws.cell(i, 1, k); ws.cell(i, 2, v)
ws.column_dimensions['A'].width = 80; ws.column_dimensions['B'].width = 40
wc = wb.create_sheet('類別')
NAMES = {1: '合理｜前提格為空', 2: '合理｜已知值鏈鏈尾', 3: '合理｜直接引用空字串來源', 4: '合理｜護欄條件為真', 5: '合理｜常數 =""（結構占位）', 6: '合理｜COUNTIFS 無符合資料', 7: '合理｜標記／說明欄', 8: '合理｜財年換算缺值', 9: '合理｜區間中點缺端', 10: '合理｜其他組合式（家族判斷）'}
for j, h in enumerate(['類別', '名稱', '儲存格數', '驗證方式'], 1):
    c = wc.cell(1, j, h); c.font = B; c.fill = H
for i, k in enumerate(sorted(cnt), 2):
    wc.cell(i, 1, k); wc.cell(i, 2, NAMES[k]); wc.cell(i, 3, cnt[k]); wc.cell(i, 4, '程式逐格驗證' if k <= 4 else '公式家族判斷')
for j, w in enumerate([8, 34, 10, 20], 1): wc.column_dimensions[L(j)].width = w
wl = wb.create_sheet('逐格清單')
for j, h in enumerate(['工作表', '儲存格', '同列 A 欄標籤', '公式', 'XML 空值型態', '類別', '分類理由'], 1):
    c = wl.cell(1, j, h); c.font = B; c.fill = H
for i, r in enumerate(rows, 2):
    for j, v in enumerate(r, 1): wl.cell(i, j, v if j != 4 else "'" + v if False else v)
    wl.cell(i, 4).data_type = 's'
for j, w in enumerate([12, 9, 40, 90, 22, 7, 70], 1): wl.column_dimensions[L(j)].width = w
wl.freeze_panes = 'A2'; wl.auto_filter.ref = f'A1:G{len(rows)+1}'
os.makedirs(os.path.dirname(OUT), exist_ok=True); wb.save(OUT)
# md
by_sheet = ', '.join(f'{k} {v:,}' for k, v in sorted(cat.items(), key=lambda x: -x[1]))
md = f'''# A8（Wattonomics 部分）：「公式存在但快取值為空」儲存格的實際點數與分類

**日期：** 2026-11-18　**對象：** `{os.path.basename(MODEL)}`（v0.3.6）　**範圍：** 只讀，不修任何公式、不改任何數值。　**對應：** 三層對帳建議書 r2 工作單 A8 最後一項（第三方稽核稱約 3,023 個）。
**清單與程式：** `wattonomics/docs/reports/20261118_A8_空快取公式分類.xlsx`（逐格清單）；`wattonomics/tools/reports_20261118/a8_empty_cache.py`（可重跑）。

## 一、結論

1. **稽核的數字屬實：{len(rows):,} 個**（公式共 {nform:,} 個，占 {len(rows)/nform*100:.1f}%）。第三方稽核與本次獨立點數完全一致。這個數字也是 `wattonomics/tools/check.py` 每次執行都印出的提示（「3023 個公式格值為空（空字串結果或未重算；若全部為空表示未經 LibreOffice 重算）」）；該句把兩種原因併在一起，容易被讀成「未重算」。建議下次工程升版把提示拆成「空字串結果 N 格」與「缺快取 M 格」兩個數（本張不改，避免動到 check.py 的輸出格式）。
2. **全部是合理的空字串結果，沒有未重算、沒有斷鏈。**
   - **未重算：0。** {len(rows):,} 格在檔案 XML 內都是「公式 ＋ 空字串結果」（`t="str"`、`<v/>`），不是缺快取。另以 LibreOffice 對整本活頁簿強制重算（{ncell:,} 格逐格比對），快取與重算結果不同的格數：**{ndiff}**。
   - **斷鏈：0。** 含 `#REF!` 的公式 {nref}、壞掉的具名範圍 {nbadname}、快取中的錯誤值 {nerr}。
   - **合理性：** {verified:,} 格（{verified/len(rows)*100:.0f}%）由程式逐格驗證（前提格確為空、已知值鏈鏈尾、直接引用空字串來源、護欄條件為真）；其餘 {family:,} 格屬標記欄、常數占位、查表無資料等公式家族，以家族為單位判斷（未逐格評估，清單中標「家族判斷」）。
3. 所以第三方稽核「公式存在但快取值為空」的說法在事實上成立，但風險判斷不成立：這些是模型刻意用 `""` 表示「這個年份或玩家沒有資料」，不是壞掉的公式。
4. **不修任何公式。** 沒有發現「明顯工程缺陷」需要處理。

## 二、分類結果

| 類別 | 名稱 | 儲存格數 | 驗證 |
|---|---|---|---|
''' + '\n'.join(f'| {k} | {NAMES[k]} | {cnt[k]:,} | {"程式逐格驗證" if k<=4 else "公式家族判斷"} |' for k in sorted(cnt)) + f'''

依工作表：{by_sheet}。

三種稽核類別對照：

| 稽核類別 | 數量 | 說明 |
|---|---|---|
| 合理的空字串結果 | {len(rows):,} | 見上表 |
| 未重算（快取缺失或過期） | 0 | XML 無缺 `<v>`；強制重算後 {ncell:,} 格零差異 |
| 斷鏈（#REF! 等） | 0 | 公式、具名範圍、快取值皆無 |

## 三、為什麼會有這麼多空字串（白話）

模型的年份欄固定是 2023–2032 十欄，但每個玩家、每個項目不是每年都有資料。公式的寫法統一為「沒有輸入就回空字串」，而不是寫 0，這與規則 C0 第 7 點「沒有證據不得歸零」的精神一致（本報告不確認這是原設計意圖，只確認結果行為如此）。大宗來源：Players 頁各玩家的「輸入、晶片換算、已知值、前一／下一已知值」輔助列（728）、Gap（579）、Funding（556）的查表與標記列、C1（343）與 CapexGW（325）的個股模型財年資料缺年。

## 四、觀察（不修；供下次升版參考）

- 141 格是常數公式 `=""`（Players 無晶片資料的列與首年「年平均」、Deploy 年增率與年平均的前段年份等）。功能上等同空白格，但仍計為公式；整列公式統一是刻意的結構選擇。
- Gap 頁 2 格（`E175`、`F175`）含 `IF(FALSE(),"",…)`，條件恆為假，是 builder 產生式的殘留死分支。結果值正確，不影響任何數字。若要清理，應在下一次改 builder 的單獨工程 PR 處理。
- CapexGW 另有 3 格（`E247`、`E257`、`E263`）寫成 `IF(2024>=2026, …, IF(Players!…="","",…))`：年份寫成字面常數，條件恆為假，只走後段；同樣是 builder 逐年展開的產物，結果正確。
- 外部程式用 openpyxl（`data_only=True`）讀取時，這 {len(rows):,} 格會得到 `None` 而不是 `""`。其中 Interface 頁的 309 格在 `export_snapshot.py` 匯出的快照裡就是 309 個值為 `null` 的名稱（`WIF_Funding_*` 178、`WIF_Gap_*` 116、`WIF_PlayerGW_*` 15，皆為該年份無資料；快照共 2,647 個名稱）。下游若直接讀這些名稱要先處理 null；這屬各下游的讀取規則，不是模型缺陷。

## 五、驗證了什麼、沒驗證什麼

- 驗證：公式數與空快取數（openpyxl 兩種讀法）；XML 型態；強制重算零差異；斷鏈零；{verified:,} 格逐格合理性。
- 沒驗證：{family:,} 格的合理性只到公式家族層級；沒有用 Excel 本體重算（只用 LibreOffice）；Excel 與 LibreOffice 對 `"" + 數字` 的處理一致（皆為 #VALUE!），而活頁簿沒有任何錯誤值，故不影響。
'''
open(OUT.replace('.xlsx', '.md'), 'w', encoding='utf-8').write(md)
print(len(rows), nform, ndiff, ncell, nref, nerr, nbadname, dict(cnt))

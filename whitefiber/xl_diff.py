# Excel 逐格比對（公式與常數；另可比對重算後的值）。用法：python3 xl_diff.py a.xlsx b.xlsx [--values] [--by-label] [--ignore=工作表,工作表]
# --ignore（v4.3）：兩檔都略過這些工作表（例如內容刻意改寫的「來源」、新增的「摘要」），其餘照常比對；略過的工作表另外列出。
# --by-label（v4.2）：列數或 A 欄不同的工作表（例如新增輸入列造成位移），改以「區段＋A 欄列名稱」配對列再逐格比對；
#   區段＝該列之上最近的第一層區段標題列（淺藍底 D9E2F3），同名的表頭列（例如各區的「項目」）因此可區分。
#   配對鍵重複、或 A 欄空白的列有內容時直接報錯並列出，不自行配對。A、B 兩檔結構相同的工作表仍逐格比對。
#   新版多出的列列為「新增列」，不計為差異；舊版有、新版沒有的列計為差異。
#   公式比對時，舊版公式中的列號（含跨工作表引用）先依列對照換算成新位置再比；
#   舊版是文字、新版改為文字公式且重算後的值與舊版文字相同者列為「文字改為活公式」，不計為差異（值由 --values 另行核對）。
#   新版公式把對「新增列」儲存格的引用換回該格的常數後（「+-」併為「-」）與舊版公式相同者，列為「常數改為引用輸入格」，不計為差異。
import sys, re, openpyxl
a, b = sys.argv[1], sys.argv[2]; vals = '--values' in sys.argv; by_label = '--by-label' in sys.argv
IGN = [x for o in sys.argv if o.startswith('--ignore=') for x in o.split('=', 1)[1].split(',') if x]


def norm(v):
    if isinstance(v, str): return re.sub(r'v\d\.\d', 'vX', v)
    if hasattr(v, 'ref') and not isinstance(v, (int, float)):  # 模擬運算表／陣列公式物件（v4.1）：依屬性比對，不比物件身分
        return f'{type(v).__name__}{sorted((k, str(x)) for k, x in vars(v).items())}'
    return v


# --expect=檔案（v4.5）：預期差異清單，每行「工作表!儲存格  原因」（# 開頭為註解）。清單內的格有差異時另列「預期差異」、不計入差異數；
# 清單內的格若沒有差異也會列出（提醒清單過期）。每個任務只列經說明的格，升版後清空。
EXP, SEC_RENAME, ROW_RENAME = {}, {}, {}  # SEC_RENAME：新版區段標題 → 舊版（「區段 舊標題 => 新標題」；該區段的列仍以列名稱配對比對）
for o in sys.argv:
    if o.startswith('--expect='):
        for ln in open(o.split('=', 1)[1], encoding='utf-8'):
            ln = ln.strip()
            if ln.startswith('列 ') and ' => ' in ln:  # 列名稱改名（同一區段內）
                o_, n_ = ln[2:].split(' => ', 1); ROW_RENAME[norm(n_.strip())] = norm(o_.strip()); continue
            if ln.startswith('區段 ') and ' => ' in ln:
                o_, n_ = ln[3:].split(' => ', 1); SEC_RENAME[norm(n_.split('  ')[0].strip())] = norm(o_.strip()); continue
            if ln and not ln.startswith('#'):
                k, _, why = ln.partition(' '); EXP[k] = why.strip()
SEC = 'FFD9E2F3'  # build_xlsx.py FILL_SEC（第一層區段標題）


def same(x, y):
    return x == y or (isinstance(x, float) and isinstance(y, float) and abs(x - y) < 1e-9)


def keys(ws, new=False):
    """每列的配對鍵 (區段, A 欄名稱)；回傳 {鍵: 列號} 與錯誤清單。"""
    out, errs, sec, seen = {}, [], '', {}
    for r in range(1, ws.max_row + 1):
        c = ws.cell(r, 1); lab = c.value
        if lab is None or str(lab).strip() == '':
            if any(ws.cell(r, j).value is not None for j in range(2, ws.max_column + 1)):
                errs.append(f'{ws.title} 第 {r} 列：A 欄空白但有內容，無法以列名稱配對')
            continue
        lab = str(lab); lab = lab if re.fullmatch(r'v\d+\.\d+', lab) else norm(lab)  # 句中的版本號正規化（版本紀錄頁的列名稱本身就是版本號，不動）
        if c.fill is not None and c.fill.fgColor is not None and c.fill.fgColor.rgb == SEC:
            sec = SEC_RENAME.get(lab, lab) if new else lab; k = ('§', sec)  # v4.5：區段標題列另立鍵，不與區段內同名的資料列（例如「加權目標價」）重複
        else:
            k = (sec, ROW_RENAME.get(lab, lab) if new else lab)
        seen.setdefault(k, []).append(r)
        out[k] = r
    for k, rs in seen.items():
        if len(rs) > 1:
            errs.append(f'{ws.title}：列名稱重複「{k[1]}」（區段「{k[0][:30]}」；第 {"、".join(map(str, rs))} 列）')
    return out, errs


wa, wb = openpyxl.load_workbook(a, data_only=vals), openpyxl.load_workbook(b, data_only=vals)
vb = openpyxl.load_workbook(b, data_only=True) if (by_label and not vals) else None
for _w in (wa, wb) + ((vb,) if vb else ()):
    for _n in IGN:
        if _n in _w.sheetnames: del _w[_n]
diffs, added, conv, errs, inl = [], [], [], [], []
if wa.sheetnames != wb.sheetnames: diffs.append(('sheets', wa.sheetnames, wb.sheetnames))
col_a = lambda ws: [ws.cell(r, 1).value for r in range(1, ws.max_row + 1)]
rowmap, mode = {}, {}  # rowmap[工作表][舊列號] = 新列號（label 模式）
for n in wa.sheetnames:
    if n not in wb.sheetnames: continue
    A, B = wa[n], wb[n]
    if by_label and (A.max_row != B.max_row or col_a(A) != col_a(B)):
        ka, ea = keys(A); kb, eb = keys(B, new=True); errs += ea + [e + '（新版）' for e in eb]
        mode[n] = 'label'; rowmap[n] = {ra: kb[k] for k, ra in ka.items() if k in kb}
        added += [(n, rb, k[1]) for k, rb in sorted(kb.items(), key=lambda t: t[1]) if k not in ka]
        diffs += [(n, f'第 {ra} 列', f'新版缺少此列：{k[1][:60]}', '') for k, ra in ka.items() if k not in kb]
    else:
        mode[n] = 'pos'
if errs:
    print('ERROR：無法以列名稱配對（不自行配對）'); [print('  ' + e) for e in errs]; sys.exit(2)

REF = re.compile(r"(?:('(?:[^']|'')+'|[^\W\d][\w.]*)!)?(\$?[A-Z]{1,3}\$?)(\d+)(?![\w(])")


def remap(f, sheet):
    """把舊版公式中的列號依 rowmap 換成新列號；字串常數內的文字不動。"""
    if not isinstance(f, str) or not f.startswith('='): return f
    parts = re.split(r'("(?:[^"]|"")*")', f)
    def sub(m):
        sh = m.group(1); tgt = sh.strip("'").replace("''", "'") if sh else sheet
        if m.start() > 0 and (m.string[m.start() - 1].isalnum() or m.string[m.start() - 1] in '_.'): return m.group(0)
        if not sh and m.start() > 0 and m.string[m.start() - 1] == ':':  # v4.5：範圍終點（'頁'!C61:G61 的 G61）沿用起點的工作表
            pm = re.search(r"(?:('(?:[^']|'')+'|[^\W\d][\w.]*)!)?\$?[A-Z]{1,3}\$?\d+:$", m.string[:m.start()])
            if pm and pm.group(1): tgt = pm.group(1).strip("'").replace("''", "'")
        rm = rowmap.get(tgt)
        if rm is None: return m.group(0)
        r = int(m.group(3)); return f"{sh + '!' if sh else ''}{m.group(2)}{rm.get(r, r)}"
    return ''.join(p if i % 2 else REF.sub(sub, p) for i, p in enumerate(parts))


ADDED = set()  # 新版新增列 (工作表, 列號)


def inline(f, sheet):
    """新版公式中引用新增列的儲存格換成該格常數（僅常數格）。"""
    if not isinstance(f, str) or not f.startswith('='): return f
    parts = re.split(r'("(?:[^"]|"")*")', f)
    def sub(m):
        sh = m.group(1); tgt = sh.strip("'").replace("''", "'") if sh else sheet
        if (tgt, int(m.group(3))) not in ADDED: return m.group(0)
        v = wb[tgt][m.group(2).replace('$', '') + m.group(3)].value
        return repr(v) if isinstance(v, (int, float)) else m.group(0)
    return ''.join(p if i % 2 else REF.sub(sub, p).replace('+-', '-') for i, p in enumerate(parts))


def rng(s, sheet):  # 模擬運算表屬性（ref＝C143:E145、r1＝C5）亦依列對照換算
    return re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)', lambda m: f'{m.group(1)}{rowmap.get(sheet, {}).get(int(m.group(2)), int(m.group(2)))}', s)


def norm_old(v, sheet):
    if vals: return norm(v)
    if isinstance(v, str): return norm(remap(v, sheet))
    if hasattr(v, 'ref') and not isinstance(v, (int, float)):
        return f'{type(v).__name__}{sorted((k, rng(str(x), sheet) if k in ("ref", "r1", "r2") else remap(str(x), sheet)) for k, x in vars(v).items())}'
    return v


ADDED.update((n, r) for n, r, _ in added)
for n in wa.sheetnames:
    if n not in wb.sheetnames: continue
    A, B = wa[n], wb[n]
    if mode[n] == 'label':
        pairs = sorted(rowmap[n].items())
    else:
        pairs = [(r, r) for r in range(1, max(A.max_row, B.max_row) + 1)]
    for ra, rb in pairs:
        for c in range(1, max(A.max_column, B.max_column) + 1):
            x, y = norm_old(A.cell(ra, c).value, n) if by_label else norm(A.cell(ra, c).value), norm(B.cell(rb, c).value)
            if same(x, y): continue
            ox = A.cell(ra, c).value
            if vb is not None and isinstance(ox, str) and not ox.startswith('=') and isinstance(y, str) and y.startswith('=') \
                    and norm(vb[n].cell(rb, c).value) == norm(ox):
                conv.append((n, B.cell(rb, c).coordinate, ox[:50])); continue
            if by_label and not vals and isinstance(y, str) and norm(inline(B.cell(rb, c).value, n)) == x:
                inl.append((n, B.cell(rb, c).coordinate)); continue
            diffs.append((n, B.cell(rb, c).coordinate, str(x)[:70], str(y)[:70]))
hit = [d for d in diffs if f'{d[0]}!{d[1]}' in EXP]; diffs = [d for d in diffs if f'{d[0]}!{d[1]}' not in EXP]
print(len(diffs), 'differences'); [print(*d) for d in (diffs if '--all' in sys.argv else diffs[:15])]  # --all（v4.5）：列出全部差異
if EXP:
    print(f'預期差異 {len(hit)}（--expect 清單，不計入）：'); [print(f'  {n}!{c}：{x} → {y}（{EXP[n + "!" + c]}）') for n, c, x, y in hit]
    miss = sorted(set(EXP) - {f'{n}!{c}' for n, c, *_ in hit})
    if miss: print(f'清單中沒有差異的格 {len(miss)}：' + '、'.join(miss))
if IGN: print(f'略過的工作表（未比對）：{"、".join(IGN)}')
if by_label:
    lab = [n for n, m in mode.items() if m == 'label']
    print(f'以列名稱配對的工作表：{"、".join(lab) if lab else "無（全部逐格）"}')
    print(f'新增列 {len(added)}：'); [print(f'  {n} 第 {r} 列 {t[:60]}') for n, r, t in added]
    if inl: print(f'常數改為引用輸入格 {len(inl)}（引用換回常數後與舊版公式相同）：' + '、'.join(f'{n}!{c}' for n, c in inl))
    if conv: print(f'文字改為活公式 {len(conv)}（重算後值與舊版文字相同）：'); [print(f'  {n} {c} {t}') for n, c, t in conv]

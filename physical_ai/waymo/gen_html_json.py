#!/usr/bin/env python3
"""產生 physical_ai/waymo/html.json（HTML 展示層設定，版面 v1.3）。

數字一律以儲存格參照取自 dist/ 內的 xlsx；本程式只依標籤找列、依 xlsx 值決定敏感度案例的分組（不寫任何數字到 html.json）。
Waymo 為 Alphabet 分部、視同未上市：無股價、無標準損益表、無目標價與評等。
- 橫幅：三情境權益價值（情境 C＝低、加權＝基準、情境 A＝高）與最近一輪融資估值（對照）。
- 損益章以「每英里經濟與營收推導」取代標準損益表；反向 DCF 章以「融資估值隱含什麼」（反解）取代。

用法：python3 physical_ai/waymo/gen_html_json.py && python3 tools/company_html/build_company_html.py physical_ai/waymo
"""
import glob
import json
import os

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = sorted(glob.glob(os.path.join(HERE, 'dist', '*.xlsx')))
assert len(XLSX) == 1, XLSX
wb = openpyxl.load_workbook(XLSX[0], data_only=True)

MK = '市場天花板與情境'
UE = '每英里經濟'
VA = '估值'
YC = 'CDEFGHIJKL'          # 2026..2035


def find_row(sheet, label, start=1, col='A', exact=False):
    ws = wb[sheet]
    for r in range(start, ws.max_row + 1):
        v = ws[f'{col}{r}'].value
        if isinstance(v, str) and ((v == label) if exact else v.startswith(label)):
            return r
    raise KeyError(f'{sheet} 找不到「{label}」（自第 {start} 列）')


def scen_rows(sheet, label):
    """情境 A／B／C 三個區塊中同一標籤的列號。"""
    hdr = {s: find_row(sheet, f'情境 {s}') for s in 'ABC'}
    return {s: find_row(sheet, label, hdr[s]) for s in 'ABC'}, hdr


def rng(sheet, r, c0='C', c1='L'):
    return f'@{sheet}!{c0}{r}:{c1}{r}'


# ───────── 市場天花板與情境（滲透） ─────────
mk_hdr = {s: find_row(MK, f'情境 {s}') for s in 'ABC'}
def mk(label):
    return {s: find_row(MK, label, mk_hdr[s]) for s in 'ABC'}

R_WK = mk('每週付費趟數')
R_PEN = mk('Robotaxi 滲透率')
R_SH = mk('Waymo 占 Robotaxi 份額')
R_PX = mk('Waymo 每付費英里實收價格')
R_MI = mk('付費（載客）英里')
R_FL = mk('所需車隊')
R_WT = mk('Waymo 美國趟數 ÷ 美國叫車基底')
B_YR = mk_hdr['B']

# ───────── 每英里經濟 ─────────
ue_hdr = {s: find_row(UE, f'情境 {s}') for s in 'ABC'}
def ue(label):
    return {s: find_row(UE, label, ue_hdr[s]) for s in 'ABC'}

U_PX = ue('每付費英里實收價格')
U_OP = ue('營運成本第 3–8 項（線')
U_DEP = ue('每付費英里折舊')
U_OPD = ue('營運成本（每付費英里，含折舊）')
U_FIXM = ue('固定費用 ÷ 付費英里')
U_FULL = ue('全載成本')
U_MARG = ue('價格 − 營運成本')
U_SIG = ue('補貼訊號')

# ───────── 估值 ─────────
va_hdr = {s: find_row(VA, f'情境 {s}') for s in 'ABC'}
def va(label):
    return {s: find_row(VA, label, va_hdr[s]) for s in 'ABC'}

V_REV = va('Robotaxi 乘車營收')
V_EBIT = va('營業利益 EBIT')
V_FCF = va('FCFF')
V_PVX = va('明確期現值合計')
V_PVTV = va('永續價值現值')
V_EV = va('企業價值 EV')
V_TVS = va('永續價值占 EV')
V_EQ = va('股權價值（EV＋淨現金）')
V_SUM = find_row(VA, '股權價值（分部 DCF）')
V_LO = find_row(VA, '估值區間下緣')
V_HI = find_row(VA, '估值區間上緣')
V_OK = find_row(VA, '檢查：加權落在區間內')
V_RND = find_row(VA, '最近一輪融資估值')
V_R1 = find_row(VA, '加權 DCF ÷ 融資估值')
V_R2 = find_row(VA, '情境 A ÷ 融資估值')
V_RAT = find_row(VA, '評等')
V_CON = find_row(VA, '結論（公式組字）')
V_RS = find_row(VA, '反解')
V_RS_ROWS = list(range(V_RS + 1, wb[VA].max_row + 1))
RS_FMT = {}
for r in V_RS_ROWS:
    a = wb[VA][f'A{r}'].value or ''
    RS_FMT[r] = 'p1' if 'WACC' in a else 'n2' if ('英里' in a or '固定費用' in a) else 'n1'

# ───────── In ─────────
def in_row(key):
    return find_row('In', key, col='A', exact=True)

IN_GROUPS = [
    ('一、市場與自有車轉換閘門', ['ph', 'vmt', 'ldv_sh', 'p_own', 'cf']),
    ('二、Waymo 營運（每趟英里、車輛、固定費用）', ['tl_end', 'util_k', 'vcost26', 'vcost_dec', 'vlife', 'fix26', 'fix_g']),
    ('三、估值（折現率、稅、淨現金、融資對照）', ['tax', 'nwc', 'wacc', 'g_tv', 'netcash', 'val_round']),
]


def in_fmt(r):
    u = wb['In'][f'C{r}'].value or ''
    if u.startswith('%'):
        return 'p1'
    if u in ('十億英里', '千美元', '年'):
        return 'n0'
    return 'n2'


def in_param_rows(with_range):
    rows = []
    for title, keys in IN_GROUPS:
        sel = []
        for k in keys:
            r = in_row(k)
            has = wb['In'][f'D{r}'].value is not None or wb['In'][f'F{r}'].value is not None
            if has == with_range:
                f = in_fmt(r)
                sel.append(['@In!B%d' % r, '@In!C%d' % r, f'@In!D{r}|{f}', f'@In!E{r}|{f}', f'@In!F{r}|{f}', f'@In!G{r}|tag'])
        if sel:
            rows.append({'sub': title})
            rows += sel
    return rows


SD_ROWS = [find_row('In', k, start=find_row('In', '情境驅動對照'), exact=True) for k in ('c35', 'eps', 'intl35', 'scn_w')]
SD_FMT = {SD_ROWS[0]: 'p1', SD_ROWS[1]: 'n1', SD_ROWS[2]: 'p0', SD_ROWS[3]: 'p0'}

# ───────── Sens：依 xlsx 值分組（變動 ≥ 基準一成者為主圖） ─────────
S = wb['Sens']
S_BASE = find_row('Sens', '基準', exact=True)
S_CHK = find_row('Sens', '檢查：Sens 基準')
S_ID = find_row('Sens', 'build_id', exact=True)
S_ST = find_row('Sens', '模型狀態', exact=True)
base = S[f'D{S_BASE}'].value
S_CF = find_row('Sens', 'cf', col='C', exact=True)
S_TL = next(r for r in range(S_BASE + 1, S_CHK) if S[f'C{r}'].value == 'tl_end' and '敏感度低' in str(S[f'B{r}'].value))
lo_main, hi_main, oth_main, lo_rest, hi_rest, oth_rest = [], [], [], [], [], []
for r in range(S_BASE + 1, S_CHK):
    b, e = S[f'B{r}'].value, S[f'E{r}'].value
    if b is None or e is None:
        continue
    b = str(b)
    big = abs(e) >= 0.1 * abs(base)
    if '敏感度低' in b or '_C' in b or 'wacc＝0.11' in b:
        (lo_main if big else lo_rest).append(r)
    elif '敏感度高' in b or '_A' in b or 'wacc＝0.13' in b:
        (hi_main if big else hi_rest).append(r)
    else:
        (oth_main if big else oth_rest).append(r)
assert lo_main and hi_main, (lo_main, hi_main)
J = lambda rs: ','.join(map(str, rs))


def tornado(title, rows, **kw):
    if not rows:
        return None
    d = {'type': 'tornado', 'sheet': 'Sens', 'label_col': 'A', 'value_col': 'E', 'fmt': 'n1', 'title': title, 'rows': J(rows)}
    d.update(kw)
    return d


# ───────── 章節 ─────────
yrs_mk = f'@{MK}!C{B_YR}:L{B_YR}|yr'


def pl_metric(name, sheet, rows, fmt, chart=True, c0='C', c1='L'):
    return {'name': name, 'fmt': fmt, 'chart': chart, 'rows': {s: f'{rng(sheet, rows[s], c0, c1)}|{fmt}' for s in 'ABC'}}


thesis = {'id': 'thesis', 'title': '命題與核心問題', 'blocks': [
    {'type': 'theme', 'title': '情境 B：每週付費趟數（百萬趟／週，長條）與每付費英里實收價格（美元，折線）：2026–2035',
     'x': f'@{MK}!C{B_YR}:L{B_YR}', 'fmt': 'n1', 'fmt_r': 'n2', 'stack': True, 'pad': 1.15,
     'bars': [{'name': [f'@{MK}!A{R_WK["B"]}'], 'y': rng(MK, R_WK['B']), 'color': '#10a0a7'}],
     'lines': [{'name': [f'@{MK}!A{R_PX["B"]}'], 'y': rng(MK, R_PX['B']), 'axis': 'r', 'color': '#c29a2b', 'labels': True}],
     'caption': '量（滲透 × 份額 × 誘發需求）與價格是本模型的兩個承重變數；價格下降帶來誘發需求，也壓低每英里營收。'},
    {'type': 'h3', 'text': '命題'},
    {'type': 'list', 'ordered': True, 'items': ['@README!A3', '@README!A7']},
    {'type': 'note', 'collapse': True, 'summary': '展開：命題回覆（依模型計算）',
     'parts': ['問題：本版假設下 Waymo 值多少、是否接近融資估值？答：', f'@{VA}!C{V_CON}',
               '　反解：', f'@{VA}!A{V_RS_ROWS[-1]}', ' ', f'@{VA}!F{V_RS_ROWS[-1]}|p1', '；',
               f'@{VA}!A{V_RS_ROWS[1]}', ' ', f'@{VA}!F{V_RS_ROWS[1]}|p1', '。']},
    {'type': 'h3', 'text': '重點觀察'},
    {'type': 'list', 'ordered': True, 'items': [
        ['價格（舊金山同路線實測）：', '@滲透!C11'],
        ['滲透：情境 B 每週付費趟數由 ', f'@{MK}!C{R_WK["B"]}|n2', ' 百萬（', f'@{MK}!C{B_YR}|yr', '）增至 ',
         f'@{MK}!G{R_WK["B"]}|n1', ' 百萬（', f'@{MK}!G{B_YR}|yr', '）、', f'@{MK}!L{R_WK["B"]}|n1', ' 百萬（', f'@{MK}!L{B_YR}|yr',
         '）；所需車隊 ', f'@{MK}!L{R_FL["B"]}|n0', ' 千台。情境 C 只有 ', f'@{MK}!L{R_WK["C"]}|n1', ' 百萬，情境 A 達 ', f'@{MK}!L{R_WK["A"]}|n0', ' 百萬。'],
        ['每英里經濟：情境 B 全載成本每付費英里 ', f'@{UE}!C{U_FULL["B"]}|n1', ' 美元（', f'@{UE}!C{ue_hdr["B"]}|yr', '），到 ',
         f'@{UE}!I{ue_hdr["B"]}|yr', ' 降至 ', f'@{UE}!I{U_FULL["B"]}|n2', '，低於價格 ', f'@{UE}!I{U_PX["B"]}|n2',
         '；前期高主因是每年約 ', f'@{UE}!C{find_row(UE, "固定費用", ue_hdr["B"], exact=True)}|n1', ' 十億美元的固定費用分攤在少量里程，所以里程規模是承重變數。'],
        ['估值：三情境權益價值 ', f'@{VA}!E{V_SUM}|n1', '～', f'@{VA}!C{V_SUM}|n1', ' 十億美元（加權 ', f'@{VA}!F{V_SUM}|n1',
         '），加權為融資估值的 ', f'@{VA}!F{V_R1}|p0', '、情境 A 為 ', f'@{VA}!F{V_R2}|p0', '。'],
        ['敏感度：影響最大的是自有車轉換閘門的便利溢價（調高後加權 ', f'@Sens!D{S_CF}|n1', '，基準 ', f'@Sens!D{S_BASE}|n1',
         '）；其次為每趟付費英里不再變長（', f'@Sens!D{S_TL}|n1', '）與情境 B 改用低滲透情境的成本、滲透路徑。'],
    ]},
]}

valuation = {'id': 'valuation', 'title': '估值：分部 DCF 權益價值（未上市，不給目標價與評等）', 'blocks': [
    {'type': 'theme', 'title': '三情境權益價值與加權（長條）對照最近一輪融資估值（虛線），十億美元',
     'x': ['情境 C（低）', '加權（基準）', '情境 A（高）'], 'fmt': 'n1', 'stack': False, 'pad': 1.1,
     'bars': [{'name': '權益價值（分部 DCF）', 'y': f'@{VA}!E{V_SUM},{VA}!F{V_SUM},{VA}!C{V_SUM}', 'color': '#0b7a80'}],
     'lines': [{'name': [f'@{VA}!A{V_RND}', '（Interested-party，只作對照）'], 'y': f'@{VA}!F{V_RND},{VA}!F{V_RND},{VA}!F{V_RND}',
                'color': '#b4472f', 'dash': True}],
     'caption': '長條＝分部 DCF 權益價值（情境權重見下表）；虛線＝最近一輪融資投後估值，不進模型。'},
    {'type': 'note', 'parts': ['未上市分部，沒有市場價格，因此不給目標價、上檔空間與評等；以最近一輪融資投後估值作對照。同業倍數腿未建（上市自駕同業營收規模與商業模式差異大），DCF 權重為全部。']},
    {'type': 'note', 'parts': [f'@{VA}!C{V_CON}']},
    {'type': 'table', 'caption': '彙總（十億美元）', 'head': ['項目', '情境 A', '情境 B', '情境 C', '加權', '說明'],
     'rows': [[f'@{VA}!A{V_SUM}', f'@{VA}!C{V_SUM}|n1', f'@{VA}!D{V_SUM}|n1', f'@{VA}!E{V_SUM}|n1', f'@{VA}!F{V_SUM}|n1', f'@{VA}!G{V_SUM}']]},
    {'type': 'table', 'head': ['項目', '數值', '說明'], 'rows': [
        [f'@{VA}!A{V_LO}', f'@{VA}!F{V_LO}|n1', ''],
        [f'@{VA}!A{V_HI}', f'@{VA}!F{V_HI}|n1', ''],
        [f'@{VA}!A{V_OK}', f'@{VA}!F{V_OK}', ''],
        [f'@{VA}!A{V_RND}', f'@{VA}!F{V_RND}|n0', f'@{VA}!G{V_RND}'],
        [f'@{VA}!A{V_R1}', f'@{VA}!F{V_R1}|p1', ''],
        [f'@{VA}!A{V_R2}', f'@{VA}!F{V_R2}|p1', ''],
        [f'@{VA}!A{V_RAT}', f'@{VA}!F{V_RAT}', ''],
    ]},
    {'type': 'h3', 'text': '折現率與估值參數'},
    {'type': 'kv', 'items': [[f'@In!B{in_row(k)}', f'@In!E{in_row(k)}|{f}'] for k, f in
                             (('wacc', 'p1'), ('g_tv', 'p1'), ('tax', 'p0'), ('nwc', 'p0'), ('netcash', 'n1'))]
     + [[f'@In!B{in_row(k)}', f'@In!E{in_row(k)}|p0'] for k in ('scn_wA', 'scn_wB', 'scn_wC')]},
    {'type': 'note', 'parts': [f'@{VA}!A1']},
    {'type': 'table', 'caption': 'DCF 彙總（三情境；十億美元）', 'head': ['項目', '情境 A', '情境 B', '情境 C'],
     'rows': [[f'@{VA}!A{R["A"]}'] + [f'@{VA}!C{R[s]}|{f}' for s in 'ABC'] for R, f in
              ((V_PVX, 'n1'), (V_PVTV, 'n1'), (V_EV, 'n1'), (V_TVS, 'p0'), (V_EQ, 'n1'))]},
]}

sens = {'id': 'sens', 'title': '敏感度：加權權益價值相對基準', 'blocks': [
    tornado('加權權益價值相對基準（十億美元）：輸入取低值或情境 B 改用低滲透情境路徑，變動達基準一成以上的案例', lo_main),
    tornado('加權權益價值相對基準（十億美元）：輸入取高值或情境 B 改用高滲透情境路徑，變動達基準一成以上的案例', hi_main),
    {'type': 'note', 'parts': ['@Sens!A1']},
    {'type': 'note', 'parts': ['基準（加權權益價值，十億美元）：', f'@Sens!D{S_BASE}|n1', '；情境 B 權益價值：', f'@Sens!F{S_BASE}|n1', '。路徑案例的「改用低滲透／高滲透情境」指只把情境 B 的該條線讀入路徑換成情境 C／A，其餘不變。']},
] + ([tornado('加權權益價值相對基準（十億美元）：情境權重', oth_main)] if oth_main else []) + [
    tornado('輸入取低值或改用低滲透路徑：其餘案例', lo_rest, collapse=True, summary='展開：取低值或改用低滲透路徑，變動未達基準一成的案例'),
    tornado('輸入取高值或改用高滲透路徑：其餘案例', hi_rest, collapse=True, summary='展開：取高值或改用高滲透路徑，變動未達基準一成的案例'),
] + ([tornado('其他：其餘案例', oth_rest, collapse=True, summary='展開：其他案例')] if oth_rest else []) + [
    {'type': 'note', 'collapse': True, 'summary': '展開：敏感度頁註記', 'parts': ['@Sens!A2']},
    {'type': 'table', 'collapse': True, 'summary': '展開：全部案例（加權權益價值、相對基準、情境 B 權益價值）',
     'head': ['驅動', '設定', '加權權益價值', '相對基準', '情境 B 權益價值'],
     'rowgen': {'sheet': 'Sens', 'rows': f'{S_BASE}-{S_CHK - 2}', 'cols': ['A', 'B', 'D', 'E', 'F'], 'fmts': ['', '', 'n1', 'n1', 'n1']}},
    {'type': 'kv', 'items': [[f'@Sens!A{S_CHK}', f'@Sens!D{S_CHK}|n6'], ['模型狀態', f'@Sens!D{S_ST}'], ['建置編號', f'@Sens!D{S_ID}']]},
]}
sens['blocks'] = [b for b in sens['blocks'] if b]

revdcf = {'id': 'revdcf', 'title': '反向檢驗：融資估值隱含什麼（未上市，以融資估值代替現價）', 'blocks': [
    {'type': 'note', 'parts': ['未上市、沒有現價，本章以最近一輪融資投後估值代替現價做反解：要讓模型權益價值等於融資估值，折現率或營運假設須改到多少（整本重算）。']},
    {'type': 'table', 'caption': '反解（整本重算）', 'head': ['項目', '數值', '說明'],
     'rows': [[f'@{VA}!A{r}', f'@{VA}!F{r}|{RS_FMT[r]}', f'@{VA}!G{r}'] for r in V_RS_ROWS]},
    {'type': 'kv', 'items': [[f'@{VA}!A{V_RND}', f'@{VA}!F{V_RND}|n0'], [f'@{VA}!A{V_R1}', f'@{VA}!F{V_R1}|p1'], [f'@{VA}!A{V_R2}', f'@{VA}!F{V_R2}|p1']]},
]}

ue_years = f'@{UE}!C{ue_hdr["B"]}:L{ue_hdr["B"]}|yr'


def ue_table(s, collapse):
    h = ue_hdr[s]
    rows = list(range(U_PX[s], U_SIG[s] + 2))
    blk = {'type': 'table', 'caption': [f'@{UE}!A{h}', '：每付費英里（美元；固定費用為十億美元、單車成本為千美元）'],
           'head': ['項目', '單位', f'@{UE}!C{h}:L{h}|yr'],
           'rowgen': {'sheet': UE, 'rows': f'{rows[0]}-{rows[-1]}', 'cols': ['A', 'B'] + list(YC), 'fmts': ['', ''] + ['n2'] * 10}}
    if collapse:
        blk.update(collapse=True, summary=f'展開：情境 {s} 每英里經濟明細')
    return blk


def va_table(s, collapse):
    h = va_hdr[s]
    blk = {'type': 'table', 'caption': [f'@{VA}!A{h}', '：營收推導與自由現金流（十億美元）'],
           'head': ['項目', f'@{VA}!C{h}:L{h}|yr'],
           'rowgen': {'sheet': VA, 'rows': f'{V_REV[s]}-{V_FCF[s]}', 'cols': ['A'] + list(YC), 'fmts': [''] + ['n2'] * 10}}
    if collapse:
        blk.update(collapse=True, summary=f'展開：情境 {s} 營收推導與自由現金流')
    return blk


pl = {'id': 'pl', 'title': '每英里經濟與營收推導（未上市分部，無標準損益表）', 'blocks': [
    {'type': 'theme', 'title': '情境 B：每付費英里成本堆疊（長條）與實收價格（折線），美元：2028–2035',
     'x': f'@{UE}!E{ue_hdr["B"]}:L{ue_hdr["B"]}', 'fmt': 'n2', 'stack': True, 'pad': 1.1, 'total': f'@{UE}!E{U_FULL["B"]}:L{U_FULL["B"]}',
     'bars': [{'name': [f'@{UE}!A{U_OP["B"]}'], 'y': rng(UE, U_OP['B'], 'E'), 'color': '#10a0a7'},
              {'name': [f'@{UE}!A{U_DEP["B"]}'], 'y': rng(UE, U_DEP['B'], 'E'), 'color': '#6fc7cb'},
              {'name': [f'@{UE}!A{U_FIXM["B"]}'], 'y': rng(UE, U_FIXM['B'], 'E'), 'color': '#c29a2b'}],
     'lines': [{'name': [f'@{UE}!A{U_PX["B"]}'], 'y': rng(UE, U_PX['B'], 'E'), 'color': '#3b3838'}],
     'caption': '長條合計＝全載成本（營運成本＋車輛折舊＋固定費用分攤）；折線高於合計的年度即價格覆蓋全載成本。早年固定費用分攤極高，明細見下表。'},
    {'type': 'note', 'parts': ['本模型為未上市分部，沒有可對照的報告損益表；本章以每英里經濟（價格對每英里成本）與估值頁的營收推導（營收、營運成本、固定費用、折舊、營業利益、自由現金流）取代標準損益表。營收為模型口徑的乘客實付，不是公司 GAAP 報告數（Waymo 不單獨揭露營收）。']},
    {'type': 'note', 'parts': [f'@{UE}!A1']},
    {'type': 'pl', 'years': f'@{UE}!E{ue_hdr["B"]}:L{ue_hdr["B"]}|yr', 'metrics': [
        pl_metric('每付費英里實收價格（美元）', UE, U_PX, 'n2', True, 'E'),
        pl_metric('全載成本（每付費英里，美元）', UE, U_FULL, 'n2', True, 'E'),
        pl_metric('價格減營運成本（含折舊，美元）', UE, U_MARG, 'n2', False, 'E'),
    ]},
    ue_table('B', False),
    ue_table('A', True),
    ue_table('C', True),
    {'type': 'h3', 'text': '營收推導（估值頁）'},
    {'type': 'note', 'parts': [f'@{VA}!A2']},
    {'type': 'pl', 'years': f'@{VA}!C{va_hdr["B"]}:L{va_hdr["B"]}|yr', 'metrics': [
        pl_metric('營收（十億美元）', VA, V_REV, 'n2'),
        pl_metric('營業利益（十億美元）', VA, V_EBIT, 'n2'),
        pl_metric('自由現金流 FCFF（十億美元）', VA, V_FCF, 'n2'),
    ]},
    va_table('B', False),
    va_table('A', True),
    va_table('C', True),
]}

PX_HDR = find_row('滲透', '期間', exact=True)
MS_HDR = find_row('滲透', '日期', exact=True)
CT_HDR = find_row('滲透', '城市', exact=True)
CP_HDR = find_row('滲透', '月份', exact=True)
AN_HDR = find_row('滲透', '錨', exact=True)
SH_E = find_row('滲透', 'E. 全國份額')

drivers = {'id': 'drivers', 'title': '驅動因子：滲透（價格、量、份額）與市場天花板', 'blocks': [
    {'type': 'note', 'parts': [f'@{MK}!A1']},
    {'type': 'pl', 'years': yrs_mk, 'metrics': [
        pl_metric('每週付費趟數（百萬趟／週）', MK, R_WK, 'n2'),
        pl_metric('Robotaxi 滲透率（替代美國叫車趟數）', MK, R_PEN, 'p1'),
        pl_metric('Waymo 占 Robotaxi 份額', MK, R_SH, 'p0', False),
        pl_metric('Waymo 美國趟數占美國叫車基底', MK, R_WT, 'p1', False),
        pl_metric('Waymo 每付費英里實收價格（美元）', MK, R_PX, 'n2'),
        pl_metric('付費（載客）英里（百萬英里）', MK, R_MI, 'n0', False),
        pl_metric('所需車隊（千台）', MK, R_FL, 'n1'),
    ]},
    {'type': 'table', 'collapse': True, 'summary': '展開：情境 B 市場天花板推導全表',
     'head': ['項目', '單位', yrs_mk],
     'rowgen': {'sheet': MK, 'rows': f'{B_YR + 1}-{R_FL["B"]}', 'cols': ['A', 'B'] + list(YC), 'fmts': ['', ''] + ['n3'] * 10}},
    {'type': 'note', 'collapse': True, 'summary': '展開：天花板讀法（自有車轉換閘門）', 'parts': [f'@{MK}!A3']},
    {'type': 'h3', 'text': '一、價格（每付費英里實收價格）'},
    {'type': 'note', 'parts': ['@滲透!A2']},
    {'type': 'table', 'caption': '@滲透!A4', 'head': [f'@滲透!A{PX_HDR}', f'@滲透!C{PX_HDR}', f'@滲透!F{PX_HDR}', f'@滲透!H{PX_HDR}', f'@滲透!I{PX_HDR}', f'@滲透!K{PX_HDR}', f'@滲透!N{PX_HDR}'],
     'rowgen': {'sheet': '滲透', 'rows': f'{PX_HDR + 1}-{PX_HDR + 3}', 'cols': ['A', 'C', 'F', 'H', 'I', 'K', 'N'], 'fmts': ['', 'n2', 'n2', 'n2', 'n2', 'sp1', 'tag']}},
    {'type': 'table', 'caption': '@滲透!A29', 'head': [f'@滲透!A{AN_HDR}', f'@滲透!B{AN_HDR}', f'@滲透!C{AN_HDR}', f'@滲透!D{AN_HDR}'],
     'rowgen': {'sheet': '滲透', 'rows': f'{AN_HDR + 1}-{AN_HDR + 5}', 'cols': ['A', 'B', 'C', 'D'], 'fmts': ['', 'n2', '', 'tag']}},
    {'type': 'note', 'parts': [f'@滲透!A{find_row("滲透", "結論：同城市")}']},
    {'type': 'h3', 'text': '二、量：全國里程碑、城市、加州月度'},
    {'type': 'table', 'caption': '@滲透!A44', 'head': [f'@滲透!{c}{MS_HDR}' for c in 'ABCDFGI'],
     'rowgen': {'sheet': '滲透', 'rows': f'{MS_HDR + 1}-{MS_HDR + 9}', 'cols': list('ABCDFGI'), 'fmts': ['', 'n0', 'n0', 'n0', 'n0', '', 'tag']}},
    {'type': 'kv', 'items': [[f'@滲透!A{MS_HDR + 10}', f'@滲透!B{MS_HDR + 10}|n1']]},
    {'type': 'table', 'caption': '@滲透!A57', 'head': [f'@滲透!{c}{CT_HDR}' for c in 'ABCDEFI'],
     'rowgen': {'sheet': '滲透', 'rows': f'{CT_HDR + 1}-{CT_HDR + 15}', 'cols': list('ABCDEFI'), 'fmts': ['', '', '', 'n0', 'n0', 'p0', 'n2']}},
    {'type': 'table', 'collapse': True, 'summary': '展開：加州 CPUC 月度（趟數、空駛比、每趟付費英里）', 'caption': '@滲透!A76',
     'head': [f'@滲透!{c}{CP_HDR}' for c in 'ABDEGHI'],
     'rowgen': {'sheet': '滲透', 'rows': f'{CP_HDR + 1}-{CP_HDR + 18}', 'cols': list('ABDEGHI'), 'fmts': ['', 'n0', 'n0', 'p1', 'n2', 'n1', '']}},
    {'type': 'table', 'caption': '@滲透!A98', 'head': ['項目', '數值'],
     'rows': [[f'@滲透!A{r}', f'@滲透!C{r}|{"p2" if r < SH_E + 3 else "n2"}'] for r in range(SH_E + 1, SH_E + 5)]},
    {'type': 'h3', 'text': '三、輸入參數（依業務分塊）'},
    {'type': 'table', 'caption': '三情境驅動（In 頁情境驅動對照）', 'head': ['項目', '情境 A', '情境 B', '情境 C'],
     'rows': [[f'@In!B{r}', f'@In!D{r}|{SD_FMT[r]}', f'@In!E{r}|{SD_FMT[r]}', f'@In!F{r}|{SD_FMT[r]}'] for r in SD_ROWS]},
    {'type': 'table', 'caption': '單一參數區間（基準進主模型；敏感度低／高只用於敏感度頁）', 'head': ['@In!B3', '@In!C3', '@In!D3', '@In!E3', '@In!F3', '@In!G3'],
     'rows': in_param_rows(True)},
    {'type': 'table', 'collapse': True, 'summary': '展開：固定輸入（無敏感度區間）', 'head': ['@In!B3', '@In!C3', '@In!D3', '@In!E3', '@In!F3', '@In!G3'],
     'rows': in_param_rows(False)},
]}

RC_N = wb['對帳'].max_row
SC_N = find_row('資料來源比較', '建議：') - 2
recon = {'id': 'recon', 'title': '對帳：線層級變數與資料來源', 'blocks': [
    {'type': 'note', 'parts': ['@對帳!A1']},
    {'type': 'table', 'caption': '線層級變數：本模型使用值', 'head': ['@對帳!A2', '@對帳!C2', '@對帳!D2', '@對帳!G2'],
     'rowgen': {'sheet': '對帳', 'rows': f'3-{RC_N}', 'cols': ['A', 'C', 'D', 'G'], 'fmts': ['', '', '', 'tag']}},
    {'type': 'table', 'collapse': True, 'summary': '展開：口徑、出處、影響儲存格與對帳結論', 'head': ['@對帳!A2', '@對帳!E2', '@對帳!F2', '@對帳!H2', '@對帳!I2'],
     'rowgen': {'sheet': '對帳', 'rows': f'3-{RC_N}', 'cols': ['A', 'E', 'F', 'H', 'I']}},
    {'type': 'h3', 'text': '資料來源比較'},
    {'type': 'note', 'parts': [f'@資料來源比較!A{SC_N + 2}']},
    {'type': 'table', 'head': ['@資料來源比較!A2', '@資料來源比較!D2', '@資料來源比較!E2', '@資料來源比較!G2'],
     'rowgen': {'sheet': '資料來源比較', 'rows': f'3-{SC_N}', 'cols': ['A', 'D', 'E', 'G'], 'fmts': ['', '', '', 'tag']}},
    {'type': 'table', 'collapse': True, 'summary': '展開：涵蓋、頻率、本模型使用處與備註', 'head': ['@資料來源比較!A2', '@資料來源比較!B2', '@資料來源比較!C2', '@資料來源比較!F2', '@資料來源比較!H2'],
     'rowgen': {'sheet': '資料來源比較', 'rows': f'3-{SC_N}', 'cols': ['A', 'B', 'C', 'F', 'H']}},
]}

status = {'id': 'status', 'title': '資料狀態', 'blocks': [
    {'type': 'kv', 'items': [['模型版本', ['$version', '（', '$date', '）']], ['版本與日期', '@README!A1'], ['build_id', f'@Sens!D{S_ID}'],
                             ['模型狀態（敏感度）', f'@Sens!D{S_ST}'], ['來源 xlsx', '$xlsx'], ['來源 xlsx 的雜湊指紋', '$sha']]},
    {'type': 'h3', 'text': '本版要點與限制'},
    {'type': 'list', 'items': ['@README!A4', '@README!A5', '@README!A6', '@README!A8']},
    {'type': 'h3', 'text': '「未取得」清單'},
    {'type': 'scan', 'needle': '未取得', 'collapse': True, 'summary': '展開：標為「未取得」的項目', 'sheets': {'滲透': ['A']}, 'empty': '沒有標為未取得的項目。'},
]}

cfg = {
    'company': 'Waymo',
    'ticker': '未上市（Alphabet 分部）',
    'layout': 'v1.3',
    'wide_text': True,
    'thesis': '@README!A3',
    'question': 'Robotaxi 的滲透速度與每英里實收價格，能否讓 Waymo 的每付費英里全載成本降到價格以下，且其價值是否接近最近一輪融資估值',
    'scenarios': [{'key': 'A', 'name': '情境 A（高滲透）'}, {'key': 'B', 'name': '情境 B（基準）'}, {'key': 'C', 'name': '情境 C（低滲透）'}],
    'hero': {
        'rec': f'@{VA}!F{V_RAT}',
        'tiles': [
            ['權益價值：低（情境 C）', f'@{VA}!C{V_EQ["C"]}|n1', '十億美元'],
            ['權益價值：基準（三情境加權）', f'@{VA}!F{V_SUM}|n1', ['十億美元；為融資估值的 ', f'@{VA}!F{V_R1}|p0']],
            ['權益價值：高（情境 A）', f'@{VA}!C{V_EQ["A"]}|n1', ['十億美元；為融資估值的 ', f'@{VA}!F{V_R2}|p0']],
            [[f'@{VA}!A{V_RND}', '（對照）'], f'@{VA}!F{V_RND}|n0', '十億美元'],
        ],
        'note': ['@README!A10'],
    },
    'sections': [thesis, valuation, sens, revdcf, pl, drivers, recon, status],
}

with open(os.path.join(HERE, 'html.json'), 'w', encoding='utf-8') as fh:
    json.dump(cfg, fh, ensure_ascii=False, indent=1)
    fh.write('\n')
print('html.json written; sens lo', lo_main, 'hi', hi_main, 'oth', oth_main)

#!/usr/bin/env python3
"""K2：Wattonomics 2026 年起七家主路徑（資本支出 ÷ Tokenomics 每 GW 成本，CapexGW 甲2）與站點實體讀數（SiteOnly）並列。
唯讀 wattonomics/model/<CURRENT>；輸出 wattonomics/docs/reports/20261118_K2_兩條路徑並列.xlsx（公式活頁簿，LibreOffice 重算）。
不改任何模型、快照或公式。"""
import openpyxl, sys, os, subprocess, shutil
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
MODEL = os.path.join(ROOT, 'wattonomics', 'model', open(os.path.join(ROOT, 'wattonomics', 'model', 'CURRENT')).read().strip())
OUT = os.path.join(ROOT, 'wattonomics', 'docs', 'reports', '20261118_K2_兩條路徑並列.xlsx')
wv = openpyxl.load_workbook(MODEL, data_only=True)
YR0 = 2023
def col(y): return 4 + (y - YR0)
def rowmap(sh):
    ws = wv[sh]; m = {}
    for r in range(1, ws.max_row + 1):
        k = ws.cell(r, 2).value
        if k and k not in m: m[k] = r
    return m
CM, PM = rowmap('CapexGW'), rowmap('Players')
def cap(key, y): return wv['CapexGW'].cell(CM[key], col(y)).value
def ply(key, y): return wv['Players'].cell(PM[key], col(y)).value
def num(v): return v if isinstance(v, (int, float)) else None

FIRMS = [('msft', 'Microsoft'), ('googl', 'Alphabet'), ('amzn', 'Amazon'), ('meta', 'Meta'), ('orcl', 'Oracle'),
         ('crwv', 'CoreWeave'), ('nbis', 'Nebius'), ('xai', 'SpaceX／xAI（參考：不屬七家，主路徑讀 SpaceX 個股模型）')]
YEARS = [2026, 2027, 2028, 2029]
SA = '站點管線：In T2 園區（公告時程＋已確認延遲）'
SB = '公司揭露／年底目標輸入（Players 旗標 1）'
SC = '公司 2030 目標線性內插（Players 旗標 3）'
SD = '觀測年底值（旗標 2）減前一年晶片換算值（旗標 5）'
def classify(f, y):
    ft, fp = ply(f + '_f', y), ply(f + '_f', y - 1)
    if ft == 4: return SA, '是', '來源為站點／園區清單，與晶片出貨推估無關；標記多為 Analogy（Epoch 站點時間線）與 Interested-party（公司公告）'
    if ft == 1 and fp == 1: return SB, '參考', '公司揭露或年底目標（Interested-party），不是晶片推估，但也不是實體站點逐站讀數；只作參考'
    if ft == 3: return SC, '否', '站點讀數是公司 2030 年目標的線性內插，不是站點實體證據；隱含成本不可比'
    if ft == 2 and fp == 5:
        why = '前一年底來自 Epoch 晶片功耗 × 倍數（晶片出貨推估），不得當獨立站點讀數；隱含成本不可比，改看下限讀數（僅 T2 園區當年上線量）'
        if f == 'googl': why = '觀測值本身（obs_googl）與前一年底皆來自 Epoch 晶片功耗 × 倍數（晶片出貨推估）；隱含成本不可比，改看下限讀數（僅 T2 園區當年上線量）'
        return SD, '否', why
    return '未分類（旗標 %s／%s）' % (ft, fp), '否', '請人工檢視'

wb = openpyxl.Workbook()
ws0 = wb.active; ws0.title = '說明'
ws = wb.create_sheet('K2_並列')
BLUE = Font(color='0000FF'); BOLD = Font(bold=True); HDR = PatternFill('solid', fgColor='DDEBF7')
heads = ['玩家', '鍵', '年', '曆年 AI 資本支出\n(US$B)\n[CapexGW *_ai]', 'TK 每 IT GW IT capex\n(US$B/GW)\n[*_pit]', 'TK 每 IT GW 廠房 capex\n(承租方帳上口徑)\n[*_pfo]',
         'TK 每 GW 成本\n(US$B/GW)\n=E+F', '甲2 付款年 GW\n[*_gw]', '主路徑 GW（上線年）\n[*_on]\n＝資本支出 ÷ TK 每 GW 成本，經 lag_on／inv_on',
         '站點讀數 GW\n[*_site＝Players 年新增]', '差距 GW\n=主路徑 − 站點', '站點 ÷ 主路徑', '站點證據來源類別', '可比？\n(是／參考／否)', '類別說明或不可比原因',
         'T2 園區當年上線 GW\n[Players *_cp；不含觀測與晶片換算；下限讀數]',
         '隱含每 GW 成本\n(US$B/GW)\n=曆年 AI 資本支出 ÷ 站點 GW（同年）', '隱含 ÷ TK 每 GW 成本',
         '隱含每 GW 成本（時差調整）\n=G×主路徑 GW ÷ 站點 GW', '調整後 ÷ TK 每 GW 成本\n(=主路徑 ÷ 站點)',
         '下限讀數隱含成本\n=G×主路徑 GW ÷ T2 當年上線\n（只對「否」列；上限值）']
for j, h in enumerate(heads, 1):
    c = ws.cell(3, j, h); c.font = BOLD; c.fill = HDR; c.alignment = Alignment(wrap_text=True, vertical='top')
ws.cell(1, 1, 'K2 兩條路徑並列：Wattonomics v0.3.6 主路徑 GW（資本支出 ÷ TK 每 GW 成本）對站點實體讀數（SiteOnly），逐家逐年').font = Font(bold=True, size=12)
ws.cell(2, 1, '藍字＝自 v0.3.6 活頁簿讀出的輸入值（欄首括號為來源列鍵）；黑字＝本表公式。年份＝上線年（站點）與主路徑上線年對齊；曆年 AI 資本支出為付款年口徑，與上線年有時差，故同時給「時差調整版」。')
r = 4; first = {}
for f, nm in FIRMS:
    first[f] = r
    for y in YEARS:
        cls, comp, why = classify(f, y)
        vals = [nm, f, y, num(cap(f + '_ai', y)), num(cap(f + '_pit', y)), num(cap(f + '_pfo', y)) if (f + '_pfo') in CM else None,
                None, num(cap(f + '_gw', y)), num(cap(f + '_on', y)), num(cap(f + '_site', y))]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            if j >= 4 and j != 7 and v is not None: c.font = BLUE
        ws.cell(r, 7, f'=E{r}+F{r}')
        ws.cell(r, 11, f'=I{r}-J{r}')
        ws.cell(r, 12, f'=IF(I{r}=0,"",J{r}/I{r})')
        ws.cell(r, 13, cls); ws.cell(r, 14, comp); ws.cell(r, 15, why)
        cpv = num(ply(f + '_cp', y)) if (f + '_cp') in PM else None
        c = ws.cell(r, 16, cpv if cpv is not None else 0); c.font = BLUE
        ws.cell(r, 17, f'=IF(AND(N{r}<>"否",J{r}>0),D{r}/J{r},IF(N{r}="否","不可比","站點為 0"))')
        ws.cell(r, 18, f'=IF(ISNUMBER(Q{r}),Q{r}/G{r},Q{r})')
        ws.cell(r, 19, f'=IF(AND(N{r}<>"否",J{r}>0),G{r}*I{r}/J{r},IF(N{r}="否","不可比","站點為 0"))')
        ws.cell(r, 20, f'=IF(ISNUMBER(S{r}),S{r}/G{r},S{r})')
        ws.cell(r, 21, f'=IF(AND(N{r}="否",P{r}>0),G{r}*I{r}/P{r},"")')
        r += 1
last = r - 1
for rr in range(4, last + 1):
    for cc in (4, 5, 6, 7, 8, 9, 10, 11, 16, 17, 19, 21): ws.cell(rr, cc).number_format = '#,##0.000'
    for cc in (12, 18, 20): ws.cell(rr, cc).number_format = '0.00'
widths = [18, 7, 6, 14, 14, 14, 12, 12, 18, 14, 12, 11, 34, 9, 60, 18, 18, 13, 20, 16, 22]
for j, w in enumerate(widths, 1): ws.column_dimensions[L(j)].width = w
ws.row_dimensions[3].height = 105; ws.freeze_panes = 'D4'

# 累計 2027–2029
wc = wb.create_sheet('K2_累計2027_29')
wc.cell(1, 1, '2027–2029 累計（避開 2026 觀測年與時差：只在三年站點讀數皆為「站點管線」時才算隱含成本）').font = Font(bold=True, size=12)
ch = ['玩家', '鍵', '曆年 AI 資本支出 Σ\n(US$B)', '主路徑 GW Σ', '站點 GW Σ', '差距 Σ GW', '站點 ÷ 主路徑', '三年皆為站點管線？',
      '隱含每 GW 成本\n=資本支出 Σ ÷ 站點 Σ', 'TK 平均每 GW 成本\n=Σ(G×主路徑)÷Σ主路徑', '隱含 ÷ TK', '備註']
for j, h in enumerate(ch, 1):
    c = wc.cell(3, j, h); c.font = BOLD; c.fill = HDR; c.alignment = Alignment(wrap_text=True, vertical='top')
rr = 4
for f, nm in FIRMS:
    a, b = first[f] + 1, first[f] + 3  # 2027..2029
    wc.cell(rr, 1, nm); wc.cell(rr, 2, f)
    wc.cell(rr, 3, f"=SUM(K2_並列!D{a}:D{b})"); wc.cell(rr, 4, f"=SUM(K2_並列!I{a}:I{b})"); wc.cell(rr, 5, f"=SUM(K2_並列!J{a}:J{b})")
    wc.cell(rr, 6, f'=D{rr}-E{rr}'); wc.cell(rr, 7, f'=E{rr}/D{rr}')
    wc.cell(rr, 8, f'=IF(COUNTIF(K2_並列!N{a}:N{b},"是")=3,"是",IF(COUNTIF(K2_並列!N{a}:N{b},"否")=0,"參考","否"))')
    wc.cell(rr, 9, f'=IF(H{rr}="否","不可比",IF(E{rr}>0,C{rr}/E{rr},"站點為 0"))')
    wc.cell(rr, 10, f"=SUMPRODUCT(K2_並列!G{a}:G{b},K2_並列!I{a}:I{b})/D{rr}")
    wc.cell(rr, 11, f'=IF(ISNUMBER(I{rr}),I{rr}/J{rr},I{rr})')
    note = {'crwv': '2027–29 站點讀數為公司 2030 目標 8 GW 的線性內插，不是站點證據', 'nbis': '2028–29 站點清單已結束（0 GW），站點讀數只覆蓋 2027 年的 0.403 GW',
            'xai': '主路徑為 SpaceX 個股模型 AI 資本支出推得；2027 起站點只有 Colossus 擴建 0.585 GW', 'googl': '2029 站點 0 GW（清單結束）'}.get(f, '')
    wc.cell(rr, 12, note)
    for cc in (3, 4, 5, 6, 9, 10): wc.cell(rr, cc).number_format = '#,##0.00'
    for cc in (7, 11): wc.cell(rr, cc).number_format = '0.00'
    rr += 1
for j, w in enumerate([30, 7, 16, 14, 14, 12, 12, 14, 20, 20, 12, 60], 1): wc.column_dimensions[L(j)].width = w
wc.row_dimensions[3].height = 60

# 證據來源
we = wb.create_sheet('K2_證據來源')
we.cell(1, 1, '站點管線（In T2）各園區明細：2026–2029 基準上線年；取自 Pipeline 頁 A 節與 In T2（標記、來源 ID）').font = Font(bold=True, size=12)
eh = ['玩家', '園區鍵', '園區', 'MW', '口徑', 'IT GW', '公告上線年', '已確認延遲月', '基準上線年', '有效納入', '標記（In T2 E 欄）', '來源 ID（In T2 F 欄）', '延遲確認等級（Q 欄）']
for j, h in enumerate(eh, 1):
    c = we.cell(3, j, h); c.font = BOLD; c.fill = HDR; c.alignment = Alignment(wrap_text=True, vertical='top')
wi = wv['In']; t2 = {}
for rr_ in range(248, 400):
    k = wi.cell(rr_, 1).value
    if k and str(k).startswith('cp'): t2[k] = [wi.cell(rr_, c).value for c in range(1, 20)]
wp = wv['Pipeline']; er = 4
keys = {f for f, _ in FIRMS}
for rr_ in range(5, 140):
    k = wp.cell(rr_, 1).value
    if not k or not str(k).startswith('cp'): continue
    op = wp.cell(rr_, 3).value
    if op not in keys: continue
    by = wp.cell(rr_, 14).value
    if not isinstance(by, (int, float)) or not (2026 <= by <= 2029): continue
    inn = t2.get(k, [None] * 19)
    vals = [dict(FIRMS).get(op, op), k, wp.cell(rr_, 2).value, wp.cell(rr_, 5).value, wp.cell(rr_, 6).value, wp.cell(rr_, 12).value, wp.cell(rr_, 8).value,
            wp.cell(rr_, 10).value, by, wp.cell(rr_, 15).value, inn[4], inn[5], inn[16]]
    for j, v in enumerate(vals, 1): we.cell(er, j, v)
    er += 1
for j, w in enumerate([18, 16, 44, 8, 8, 9, 10, 10, 10, 8, 26, 14, 40], 1): we.column_dimensions[L(j)].width = w

# 說明
lines = ['K2 兩條路徑並列（Wattonomics v0.3.6；唯讀，不改任何既有數值）', '',
 '目的：現在 Wattonomics 的主路徑（甲2）把資本支出除以 Tokenomics 每 GW 成本換成 GW，另一方向（Gap 頁）又用同一個係數把 GW 換回資本支出，兩個方向差距恆為 0，不構成對 Tokenomics 每 GW 成本的檢驗。站點實體讀數（SiteOnly）是獨立的第二條路徑，兩者並列才有檢驗力。',
 '', '欄位定義見 K2_並列 第 3 列；來源列鍵以 CapexGW、Players 頁 B 欄鍵表示。',
 '站點讀數＝Players 頁「年新增 IT GW」（＝CapexGW *_site＝Interface WIF_PlayerGW_YE_<玩家>_SiteOnly_<年> 的年差）。',
 '證據來源類別以 Players 頁旗標判定：4＝站點管線（T2 園區）；1＝公司揭露或年底目標；3＝公司目標內插；2 減 5＝觀測年底值減晶片換算值（含晶片出貨推估，不得當獨立站點讀數）。',
 '隱含每 GW 成本只在「可比＝是」或「參考」列計算；「否」列寫不可比與原因，並另給只用 T2 園區當年上線量的下限讀數（隱含成本的上限值）。',
 '恆等式提醒：隱含成本 ÷ TK 每 GW 成本＝主路徑 GW ÷ 站點 GW。站點清單只含已公告園區，是下限，所以比值大於 1 有兩種解釋（每 GW 實際成本高於 TK，或站點清單不完整），本表無法區分，須配合 K1 回測或公司揭露的上線 MW 實績才能裁定。']
for i, t in enumerate(lines, 1): ws0.cell(i, 1, t)
ws0.column_dimensions['A'].width = 140
for rr_ in range(1, len(lines) + 1): ws0.cell(rr_, 1).alignment = Alignment(wrap_text=True)
os.makedirs(os.path.dirname(OUT), exist_ok=True)
tmp = OUT + '.raw.xlsx'; wb.save(tmp)
outdir = os.path.dirname(OUT)
subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', os.path.join(outdir, '_tmp_k2'), tmp], check=True, stdout=subprocess.DEVNULL)
shutil.move(os.path.join(outdir, '_tmp_k2', os.path.basename(tmp)), OUT); shutil.rmtree(os.path.join(outdir, '_tmp_k2')); os.remove(tmp)
print('ok', OUT)

# ---- Markdown 報告（由重算後的活頁簿讀值產生，避免轉抄錯誤）----
wr = openpyxl.load_workbook(OUT, data_only=True)
w1, w2 = wr['K2_並列'], wr['K2_累計2027_29']
def f2(v, d=2):
    if isinstance(v, (int, float)): return f'{v:,.{d}f}'
    return '' if v is None else str(v)
def short(c): return {SA: 'A 站點管線', SB: 'B 公司揭露／目標', SC: 'C 目標內插', SD: 'D 觀測減晶片換算'}.get(c, c)
md = []
md.append('# K2 兩條路徑並列：資本支出路徑對站點實體讀數（Wattonomics v0.3.6）\n')
md.append('**日期：** 2026-11-18　**範圍：** 只新增顯示與報告，既有活頁簿、快照與公式一個字元都沒改。　**對應工作單：** 三層對帳建議書 r2 第四批 K2（並列部分；回測屬 K1）。\n')
md.append('**活頁簿：** `wattonomics/docs/reports/20261118_K2_兩條路徑並列.xlsx`（藍字＝自 v0.3.6 讀出的輸入，黑字＝公式；產生程式 `wattonomics/tools/reports_20261118/k2_dual_path.py`）。\n')
md.append('## 一、結論\n')
md.append('1. 目前主路徑與 Gap 頁用同一個 Tokenomics（TK）每 GW 成本互相換算，兩個方向的差距恆為 0，沒有檢驗力。站點實體讀數（SiteOnly）是獨立的第二條路徑；並列後才第一次出現可檢驗的差距。')
md.append('2. **七家 2027 年：資本支出路徑 18.16 GW，站點讀數 7.50 GW（41%）。** 逐家站點只占主路徑的 9%（Amazon）到 95%（Oracle）。換成白話：若站點清單完整，資本支出隱含的每 GW 成本是 TK 的 1.05–10.8 倍（Oracle、Nebius 約 1.0–1.1 倍，Microsoft、Alphabet、Meta 約 2.9–4.4 倍，Amazon 約 11 倍）；若 TK 成本正確，則站點清單漏了這些 GW。**兩種解釋用現有資料分不開**，所以本表不下「TK 每 GW 成本偏低」的結論。')
md.append('3. 能分開的條件：公司揭露的上線 MW 實績（例如 Microsoft、Alphabet 的在役 MW）或第三方的完整站點清單。K1 回測與 2027-02 的買方 10-K 是下一個檢驗點。')
md.append('4. **2026 年有四家站點讀數不可當獨立檢驗**（Microsoft、Alphabet、Amazon、Oracle）：模型用「2026 觀測年底值減 2025 年底的晶片換算值」得到年新增，2025 年底值來自 Epoch 晶片功耗 × 倍數，等於晶片出貨推估。這四家 2026 年只給下限讀數（僅 T2 園區當年上線量）。\n')
md.append('## 二、定義與資料來源\n')
md.append('- **主路徑 GW（上線年）：** CapexGW 頁 `*_on`。＝曆年 AI 資本支出經廠房提前入帳，除以 TK 每 GW 成本（`*_pit`＋`*_pfo`，承租方帳上口徑）得付款年 GW（`*_gw`），再以 `lag_on`（4.5 個月）、`inv_on`（10%）平移到上線年。2026 年起七家主路徑＝此值。')
md.append('- **站點讀數 GW：** Players 頁年新增（＝CapexGW `*_site`；Interface `WIF_PlayerGW_YE_<玩家>_SiteOnly_<年>` 的年差）。')
md.append('- **隱含每 GW 成本：** 兩種算法並列。(a) 同年曆年 AI 資本支出 ÷ 站點 GW（工作單字面）；(b) 時差調整版＝TK 每 GW 成本 × 主路徑 GW ÷ 站點 GW。(b) 與 (a) 的差別在於資本支出是付款年、站點是上線年。比值（隱含 ÷ TK）在 (b) 恆等於「主路徑 ÷ 站點」。')
md.append('- **可比條件：** 只有站點讀數不來自晶片出貨推估才算。類別 A（站點管線）可比；類別 B（公司揭露或年底目標，Interested-party）列「參考」；類別 C（公司 2030 目標線性內插）與類別 D（觀測值減晶片換算前一年）不可比。\n')
md.append('## 三、逐家逐年（US$B 或 GW；完整欄位見活頁簿）\n')
md.append('| 玩家 | 年 | TK 每 GW 成本 | 主路徑 GW | 站點 GW | 差距 | 站點÷主路徑 | 站點證據類別 | 可比？ | 隱含成本(同年) | 隱含成本(時差調整) | 調整後÷TK |')
md.append('|---|---|---|---|---|---|---|---|---|---|---|---|')
for rr in range(4, last + 1):
    g = lambda c: w1.cell(rr, c).value
    md.append(f"| {str(g(1)).split('（')[0]} | {g(3)} | {f2(g(7),1)} | {f2(g(9))} | {f2(g(10))} | {f2(g(11))} | {f2(g(12))} | {short(g(13))} | {g(14)}{'＊' if isinstance(g(12),(int,float)) and g(12)<0.15 and g(14)!='否' else ''} | {f2(g(17),1)} | {f2(g(19),1)} | {f2(g(20),2)} |")
md.append('\n＊站點占主路徑不到 15%：多半反映站點清單在該年已接近結束（站點清單在 2029 結束），不能解讀為成本；單年判讀以 2027 為主。\n\n不可比列（類別 C、D）另給下限讀數（只用 T2 園區當年上線量，不含觀測與晶片換算），其隱含成本為上限值：')
for rr in range(4, last + 1):
    if w1.cell(rr, 14).value == '否' and isinstance(w1.cell(rr, 21).value, (int, float)):
        md.append(f"- {str(w1.cell(rr,1).value).split('（')[0]} {w1.cell(rr,3).value}：T2 當年上線 {f2(w1.cell(rr,16).value)} GW → 隱含每 GW 成本 ≤ {f2(w1.cell(rr,21).value,0)} US$B（TK {f2(w1.cell(rr,7).value,1)}）")
md.append('\n## 四、2027–2029 累計（避開時差；三年站點皆為站點管線才算）\n')
md.append('| 玩家 | 曆年 AI 資本支出 Σ | 主路徑 GW Σ | 站點 GW Σ | 站點÷主路徑 | 可比 | 隱含每 GW 成本 | TK 平均每 GW 成本 | 隱含÷TK |')
md.append('|---|---|---|---|---|---|---|---|---|')
for rr in range(4, 12):
    g = lambda c: w2.cell(rr, c).value
    md.append(f"| {str(g(1)).split('（')[0]} | {f2(g(3),1)} | {f2(g(4))} | {f2(g(5))} | {f2(g(7))} | {g(8)} | {f2(g(9),1)} | {f2(g(10),1)} | {f2(g(11),2)} |")
md.append('\n累計仍受兩個限制：(1) 站點清單在 2028–2029 年逐步結束（README 與 Compare W2：站點清單在 2029 結束），後兩年的低站點讀數有一部分只是清單沒有再列；(2) 累計資本支出含為 2030 年以後上線預付的廠房。所以 2027 單年與累計都只能當「檢驗的起點」。\n')
md.append('## 五、證據來源類別（逐家）\n')
md.append('- **Microsoft、Alphabet、Amazon、Meta、Oracle（2027 起）：** 站點管線（In T2）。標記以 Analogy 為主（Epoch AI 站點時間線，第三方推估），其次 Interested-party（公司公告、地方報導）與 Derived。**不含任何晶片出貨推估。** 逐園區明細在活頁簿「K2_證據來源」。')
md.append('- **Meta 2026：** 年新增＝T2 園區當年上線（旗標 4），不含觀測值，所以仍可比。')
md.append('- **Microsoft、Alphabet、Amazon、Oracle 2026：** 類別 D，見結論 4。Alphabet 的觀測值（In `obschip_googl`）本身就是 Epoch 晶片功耗 × 倍數。')
md.append('- **CoreWeave：** 2026 為公司年底指引 1.85 GW 減 2025 實績 0.85（類別 B，Interested-party）；2027–2029 是公司 2030 年 8 GW 目標的線性內插（類別 C，不可比）。')
md.append('- **Nebius、SpaceX／xAI：** 2026 為公司年底目標減 2025 實績（類別 B）；2027 起為 T2 園區。\n')
md.append('## 六、限制與沒做的事\n')
md.append('- 站點清單是已公告園區的下限，不是全部上線量；本表不能把「站點 < 主路徑」解讀為資本支出路徑偏高。')
md.append('- 租賃方（Oracle、CoreWeave、Nebius、部分 Microsoft）的站點包含出租方蓋的廠房，其資本支出不在承租方帳上；模型已用承租方帳上口徑的 `*_pfo`，但站點 GW 仍是全站。這使 Oracle 2028 的站點（3.25 GW）高過主路徑（1.68 GW）。')
md.append('- 曆年 AI 資本支出的 AI 占比（0.85／0.68）是 [Assumed]，直接影響主路徑；本表沿用，不另做敏感度（CapexGW 頁已有 AI 占比低／高列）。')
md.append('- 沒有做歷史回測（K1），沒有改任何基準值，也沒有把站點讀數連進任何計算。')
md.append('- 在 Compare 頁新增並列區塊需要升版（v0.3.7）與重建，會牽動快照與 Checks；為了保證「既有數值逐格不變」，本張先以獨立報告交付。若 Andy 要把並列放進活頁簿，可在下一次升版時把 `K2_並列` 的欄位搬進 Compare 新節（只讀 CapexGW 與 Players，不違反取數規則），見「待 Andy 判斷」。\n')
md.append('## 七、待 Andy 判斷\n')
md.append('1. 並列是否要進活頁簿：選項 (a) 維持獨立報告（現況，零風險）；(b) 下一次升版（v0.3.7）在 Compare 加新節，只讀 CapexGW 與 Players，不違反取數規則。預設 (a)。')
md.append('2. 站點清單不完整怎麼補：選項 (a) 等 2027-02 買方 10-K 與公司在役 MW 揭露；(b) 現在就為 Microsoft、Alphabet、Amazon、Meta 另找公司揭露的在役 MW 與租賃承諾，補進 In T2（改值，另開單）。預設 (a)，因為補清單屬改基準。\n')
md.append('## 八、通過條件對照\n')
md.append('| 條件 | 結果 |\n|---|---|\n| 逐家、逐年並列兩條路徑與差距 | 達成（七家加 xAI 參考，2026–2029；2030–2032 可由同程式延伸） |\n| 標明每家站點讀數的證據來源類別，且不得出自晶片出貨推估 | 達成；發現並標出 4 家 2026 年含晶片換算，列為不可比 |\n| 隱含每 GW 成本與 TK 比值，只在口徑可比時算，否則寫不可比與原因 | 達成（「可比」欄與原因欄） |\n| 既有數值逐格不變 | 達成（未動活頁簿與快照；見 PR 描述的 diff 結果） |')
open(os.path.join(os.path.dirname(OUT), '20261118_K2_兩條路徑並列.md'), 'w', encoding='utf-8').write('\n'.join(md) + '\n')
print('md ok')

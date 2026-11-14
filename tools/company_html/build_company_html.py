#!/usr/bin/env python3
"""公司模型 HTML 展示層產生器（AI 半導體 Project）。

Excel 是計算事實的來源；本程式只呈現。每個數字與標籤都由產生器從 xlsx 的重算值（LibreOffice 存檔後的快取值）
抽出，並標記來源「工作表!儲存格」。HTML 內不含手寫數字（設定檔的說明文字不得含數字，年度與日期除外，建置時檢查）。

用法：
  python3 tools/company_html/build_company_html.py <公司資料夾> [--out 輸出.html] [--xlsx 指定.xlsx]
    讀 <公司資料夾>/html.json 與 <公司資料夾>/dist/ 內唯一的 xlsx，
    輸出 <公司資料夾>/dist/<xlsx 同名>.html（版本號與 xlsx 相同）。
輸出為單一自含 HTML（內嵌 CSS 與 inline SVG；無外部資源、無 JavaScript）；相同輸入逐位元組相同。

html.json 的格式見 tools/company_html/README.md。
"""
import argparse
import datetime
import hashlib
import html
import json
import math
import os
import re
import sys

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

MINUS = '−'
# 設定檔文字（非 xlsx 抽出）允許的數字樣式：年度、日期、版本；其餘一律不得含阿拉伯數字
ALLOWED_PROSE_DIGITS = re.compile(r'\b(?:FY|CY)?(?:19|20)\d{2}[A-Z]?\b|\b\d{4}-\d{2}-\d{2}\b')
REF_RE = re.compile(r"^(?P<sheet>[^!]+)!(?P<a>[A-Z]+\d+)(?::(?P<b>[A-Z]+\d+))?$")
CELL_RE = re.compile(r'^([A-Z]+)(\d+)$')


# ───────────────────────── 數字格式 ─────────────────────────
def date_text(v):
    """日期格的顯示文字（YYYY-MM-DD）；非日期原樣回傳。"""
    return v.strftime('%Y-%m-%d') if isinstance(v, datetime.datetime) else v


def _sign(v, s, plus=False):
    """負值加 U+2212；四捨五入後為零不加號。"""
    if re.fullmatch(r'[0-9,]*\.?0*[%x]?', s) and not re.search(r'[1-9]', s):
        return s
    if v < 0:
        return MINUS + s
    return ('+' + s) if plus else s


def fmt_value(v, f):
    """依格式碼把數值轉成顯示字串。格式碼：n0–n3 數字、p0–p2 百分比（值×100）、sp0–sp2 帶正負號百分比、
    x1 倍數、yr 年（無千分位）、g 通用（≥1 取兩位小數、否則四位）。"""
    if f == 'yr':
        return str(int(round(v)))
    m = re.fullmatch(r'(s?)(n|p|x)(\d)', f)
    if m:
        plus, kind, d = m.group(1) == 's', m.group(2), int(m.group(3))
        if kind == 'n':
            return _sign(v, f'{abs(v):,.{d}f}', plus)
        if kind == 'p':
            return _sign(v, f'{abs(v) * 100:,.{d}f}%', plus)
        return _sign(v, f'{abs(v):,.{d}f}x', plus)
    if f in ('', 'g'):
        if float(v).is_integer():
            return _sign(v, f'{abs(v):,.0f}')
        d = 2 if abs(v) >= 1 else 4
        s = f'{abs(v):,.{d}f}'.rstrip('0').rstrip('.')
        return _sign(v, s)
    raise ValueError(f'未知格式碼：{f}')


def parse_display(s):
    """把顯示字串還原成 (數值, 顯示精度容差)；檢核用，與 fmt_value 獨立實作。"""
    t = s.strip().replace(MINUS, '-').replace(',', '')
    scale = 1.0
    if t.endswith('%'):
        t, scale = t[:-1], 100.0
    elif t.endswith('x'):
        t = t[:-1]
    if t.startswith('+'):
        t = t[1:]
    dec = len(t.split('.')[1]) if '.' in t else 0
    return float(t) / scale, 0.5 * 10 ** (-dec) / scale


# ───────────────────────── 活頁簿與參照 ─────────────────────────
class Book:
    """xlsx 的重算值（LibreOffice 存檔後的快取值）。"""

    def __init__(self, path):
        self.path = path
        self.wb = openpyxl.load_workbook(path, data_only=True)

    def cell(self, sheet, coord):
        return self.wb[sheet][coord].value

    def expand(self, refs):
        """'PL!C5:G5,PL!B5' → [(sheet, coord), …]；範圍依列優先展開。"""
        out = []
        for part in refs.split(','):
            m = REF_RE.match(part.strip())
            if not m:
                raise ValueError(f'無法解析儲存格參照：{part!r}')
            sh, a, b = m.group('sheet'), m.group('a'), m.group('b')
            if sh not in self.wb.sheetnames:
                raise ValueError(f'xlsx 沒有工作表 {sh}（參照 {part!r}）')
            if not b:
                out.append((sh, a))
                continue
            ca, ra = CELL_RE.match(a).groups()
            cb, rb = CELL_RE.match(b).groups()
            for r in range(int(ra), int(rb) + 1):
                for c in range(column_index_from_string(ca), column_index_from_string(cb) + 1):
                    out.append((sh, f'{get_column_letter(c)}{r}'))
        return out


def split_spec(spec):
    """'@PL!C5:G5|n0' → ('PL!C5:G5', 'n0')；非參照字串回傳 (None, 文字)。"""
    if isinstance(spec, str) and spec.startswith('@'):
        body = spec[1:]
        refs, _, f = body.partition('|')
        return refs, f
    return None, spec


def check_prose(text, where):
    """設定檔文字不得含手寫數字（年度、日期除外）。"""
    rest = ALLOWED_PROSE_DIGITS.sub('', text)
    if re.search(r'\d', rest):
        raise ValueError(f'html.json 的文字含手寫數字（{where}）：{text!r}；數字須用 @工作表!儲存格 取自 xlsx')


class Ctx:
    """建置脈絡：活頁簿、記錄用過的儲存格。"""

    def __init__(self, book, meta):
        self.b = book
        self.used = set()
        self.meta = meta

    # 單一儲存格 → HTML 片段
    def span(self, sheet, coord, f):
        v = self.b.cell(sheet, coord)
        ref = f'{sheet}!{coord}'
        self.used.add(ref)
        if v is None:
            return f'<span class="na" data-ref="{ref}" title="{ref}">—</span>'
        if isinstance(v, datetime.datetime):   # 可選：日期格以 YYYY-MM-DD 文字標籤顯示（檢核比對同一格式）
            v = date_text(v)
        if isinstance(v, str):
            cls = 'lbl' + tag_class(v, f)
            return f'<span class="{cls}" data-ref="{ref}" title="{ref}">{esc(v)}</span>'
        if isinstance(v, bool):
            raise ValueError(f'{ref} 是布林值')
        txt = fmt_value(v, f)
        return f'<span class="x" data-ref="{ref}" data-v="{v!r}" data-f="{f}" title="{ref}">{txt}</span>'

    def part(self, spec):
        """一個片段：參照（可為多格範圍，以空白連接）、中繼資料（$version、$date、$sha、$xlsx）或文字。"""
        if isinstance(spec, str) and spec in self.meta:
            return f'<code class="meta" data-meta="{spec[1:]}">{esc(self.meta[spec])}</code>'
        refs, f = split_spec(spec)
        if refs is None:
            check_prose(spec, '文字片段')
            return esc(spec)
        return ' '.join(self.span(s, c, f) for s, c in self.b.expand(refs))

    def cell_html(self, spec):
        if isinstance(spec, list):
            return ''.join(self.part(p) for p in spec)
        return self.part(spec)

    def text_of(self, spec):
        """純文字（給 data-l、圖例用）：參照取 xlsx 值，數字以其格式顯示。"""
        parts = spec if isinstance(spec, list) else [spec]
        out = []
        for p in parts:
            refs, f = split_spec(p)
            if refs is None:
                check_prose(p, '文字片段')
                out.append(p)
            else:
                for s, c in self.b.expand(refs):
                    v = self.b.cell(s, c)
                    self.used.add(f'{s}!{c}')
                    out.append('' if v is None else v if isinstance(v, str) else fmt_value(v, f))
        return ''.join(out)

    def values(self, refs):
        out = []
        for s, c in self.b.expand(refs):
            v = self.b.cell(s, c)
            self.used.add(f'{s}!{c}')
            out.append(v)
        return out


def esc(s):
    return html.escape(str(s), quote=True)


def tag_class(v, f):
    """標記欄（Verified／Analogy／Assumed／Derived／Interested-party）加底色，其他文字不加。"""
    if f != 'tag':
        return ''
    for k, c in (('Assumed', 'ta'), ('Analogy', 'tn'), ('Derived', 'td'), ('Interested', 'ti'), ('Verified', 'tv')):
        if k in v:
            return ' tag ' + c
    return ' tag'


def expand_rows(spec):
    """'5-8,10' → [5,6,7,8,10]"""
    out = []
    for p in str(spec).split(','):
        p = p.strip()
        if '-' in p:
            a, b = p.split('-')
            out += list(range(int(a), int(b) + 1))
        else:
            out.append(int(p))
    return out


# ───────────────────────── 區塊 ─────────────────────────
def render_table(cx, blk):
    """通用表格。head：欄標題規格；rows：每列為規格清單（範圍自動展開）；或 rowgen：依工作表列號產生。
    grid=true：數字格表（手機仍為表格）；否則手機改為逐列堆疊。"""
    head = []
    for h in blk.get('head', []):
        head.append(h)
    rows = []
    if 'rowgen' in blk:
        g = blk['rowgen']
        cols = g['cols']
        fmts = g.get('fmts', [''] * len(cols))
        for r in expand_rows(g['rows']):
            cells = []
            ov = g.get('row_fmt', {}).get(str(r), {})
            for c, f in zip(cols, fmts):
                f = ov.get(c, f)
                cells.append(f'@{g["sheet"]}!{c}{r}|{f}')
            vals = [cx.b.cell(g['sheet'], f'{c}{r}') for c in cols]
            sub = len(cols) > 1 and vals[0] is not None and all(v is None for v in vals[1:])
            rows.append(('sub', cells[:1]) if sub else ('row', cells))
    else:
        for r in blk['rows']:
            if isinstance(r, dict):   # {"sub": 規格}：分組小標列
                rows.append(('sub', [r['sub']]))
            else:
                rows.append(('row', r))
    ncols = max([len(head)] + [len(r) if k == 'row' else 1 for k, r in rows])
    # 範圍展開後的欄數
    def expanded(r):
        out = []
        for c in r:
            if isinstance(c, str) and c.startswith('@') and ('!' in c) and (':' in c.split('|')[0] or ',' in c.split('|')[0]):
                refs, f = split_spec(c)
                out += [f'@{s}!{co}|{f}' for s, co in cx.b.expand(refs)]
            else:
                out.append(c)
        return out
    heads = expanded(head)
    labels = [cx.text_of(h) for h in heads]
    cls = 'grid' if blk.get('grid') else 'rt'
    out = []
    if blk.get('caption'):
        out.append(f'<p class="cap">{cx.cell_html(blk["caption"])}</p>')
    out.append(f'<div class="tw"><table class="{cls}">')
    if heads:
        out.append('<thead><tr>' + ''.join(f'<th scope="col">{cx.cell_html(h)}</th>' for h in heads) + '</tr></thead>')
    out.append('<tbody>')
    for kind, r in rows:
        if kind == 'sub':
            out.append(f'<tr class="sub"><th colspan="{len(heads) or ncols}" scope="colgroup">{cx.cell_html(r[0])}</th></tr>')
            continue
        cells = expanded(r)
        tds = []
        for i, c in enumerate(cells):
            lab = esc(labels[i]) if i < len(labels) else ''
            if i == 0:
                tds.append(f'<th scope="row" data-l="{lab}"><span class="cv">{cx.cell_html(c)}</span></th>')
            else:
                tds.append(f'<td data-l="{lab}"><span class="cv">{cx.cell_html(c)}</span></td>')
        out.append('<tr>' + ''.join(tds) + '</tr>')
    out.append('</tbody></table></div>')
    return ''.join(out)


def nice_ticks(lo, hi, n=4):
    span = (hi - lo) or abs(hi) or 1.0
    step = 10 ** math.floor(math.log10(span / n))
    for m in (1, 2, 2.5, 5, 10):
        if span / (step * m) <= n:
            step *= m
            break
    start = math.floor(lo / step + 1e-9) * step
    ticks, v = [], start
    while v < hi + step * 0.999:
        ticks.append(v)
        v += step
        if len(ticks) > 12:
            break
    return ticks


def render_lines(cx, blk, scen):
    """折線圖（inline SVG）。series：[{name, y, scenario}]；點的 data-ref／data-v 供檢核。"""
    xs_lab = blk['x'] if isinstance(blk['x'], list) else cx.values(blk['x'])
    for lab in (xs_lab if isinstance(blk['x'], list) else []):
        check_prose(lab, 'x 軸標籤')
    series = blk['series']
    f = blk.get('fmt', 'n0')
    W, H, L, R, T, B = 360, 236, 46, 12, 10, 30
    allv = [v for s in series for v in cx.values(s['y']) if v is not None]
    ymin = 0.0 if blk.get('ymin', 'zero') == 'zero' else min(allv)
    ymax = max(allv)
    ticks = nice_ticks(ymin, ymax)
    lo, hi = ticks[0], ticks[-1]
    n = len(xs_lab)
    px = lambda i: L + i * (W - L - R) / max(n - 1, 1)
    py = lambda v: T + (H - T - B) * (1 - (v - lo) / (hi - lo))
    o = [f'<svg viewBox="0 0 {W} {H}" class="chart" role="img" aria-label="{esc(cx.text_of(blk["title"]))}">']
    for t in ticks:
        y = py(t)
        o.append(f'<line x1="{L}" x2="{W - R}" y1="{y:.1f}" y2="{y:.1f}" class="grid"/>'
                 f'<text x="{L - 5}" y="{y + 3.5:.1f}" class="ax" text-anchor="end">{fmt_value(t, f)}</text>')
    for i, lab in enumerate(xs_lab):
        anc = 'end' if i == n - 1 else 'middle'
        o.append(f'<text x="{px(i):.1f}" y="{H - 10}" class="ax" text-anchor="{anc}">{esc(lab)}</text>')
    for si, s in reversed(list(enumerate(series))):          # C、B、A 的順序畫，A 在最上層
        refs = cx.b.expand(s['y'])
        vals = [cx.b.cell(a, c) for a, c in refs]
        pts = ' '.join(f'{px(i):.1f},{py(v):.1f}' for i, v in enumerate(vals))
        o.append(f'<polyline points="{pts}" class="ln s{si}"/>')
        nm = cx.text_of(s['name'])
        for i, ((sh, co), v) in enumerate(zip(refs, vals)):
            ref = f'{sh}!{co}'
            cx.used.add(ref)
            o.append(f'<circle cx="{px(i):.1f}" cy="{py(v):.1f}" r="3.2" class="pt s{si}" data-ref="{ref}" data-v="{v!r}" data-f="{f}">'
                     f'<title>{esc(nm)}｜{esc(xs_lab[i])}：{fmt_value(v, f)}（{ref}）</title></circle>')
    o.append('</svg>')
    legend = ''.join(f'<li><svg viewBox="0 0 28 8" class="sw" aria-hidden="true"><line x1="1" x2="27" y1="4" y2="4" class="ln s{i}"/></svg>'
                     f'{cx.cell_html(s["name"])}</li>' for i, s in enumerate(series))
    return (f'<figure class="fig"><figcaption>{cx.cell_html(blk["title"])}</figcaption>' + ''.join(o) +
            f'<ul class="legend">{legend}</ul></figure>')


def render_pl(cx, blk, cfg):
    """三情境損益：每個指標一張表（列＝情境、欄＝年度）；chart=true 的指標另畫折線圖。"""
    years = blk['years']
    scen = cfg['scenarios']
    out, charts = [], []
    for m in blk['metrics']:
        f = m.get('fmt', 'n0')
        rows = []
        for i, sc in enumerate(scen):
            rng = m['rows'][sc['key']]
            rows.append([sc['name'], rng if '|' in rng else f'{rng}|{f}'])
        out.append(render_table(cx, {
            'grid': True, 'caption': m['name'], 'head': [m.get('corner', '情境')] + (years if isinstance(years, list) else [years]), 'rows': rows}))
        if m.get('chart'):
            charts.append(render_lines(cx, {
                'title': m['name'], 'x': years if isinstance(years, list) else split_spec(years)[0], 'fmt': f, 'ymin': m.get('ymin', 'zero'),
                'series': [{'name': sc['name'], 'y': split_spec(m['rows'][sc['key']])[0]} for sc in scen]}, scen))
    return ''.join(out) + '<div class="charts">' + ''.join(charts) + '</div>'


def render_bars(cx, blk):
    """水平長條（CSS）：各情境隱含價與現價標線。items：[{label, value, series}]；ref：{label, value}。"""
    items = blk['items']
    f = blk.get('fmt', 'n0')
    vals = [cx.b.cell(*cx.b.expand(split_spec(it['value'])[0])[0]) for it in items]
    rs, rc = cx.b.expand(split_spec(blk['ref']['value'])[0])[0]
    rv = cx.b.cell(rs, rc)
    top = max(vals + [rv]) * 1.08
    o = [f'<figure class="fig bars"><figcaption>{cx.cell_html(blk["title"])}</figcaption>']
    for it, v in zip(items, vals):
        sh, co = cx.b.expand(split_spec(it['value'])[0])[0]
        cx.used.add(f'{sh}!{co}')
        o.append(f'<div class="br"><span class="bl">{cx.cell_html(it["label"])}</span>'
                 f'<span class="bt"><i class="s{it.get("series", 0)}" style="width:{v / top * 100:.2f}%"></i>'
                 f'<b style="left:{rv / top * 100:.2f}%"></b></span>'
                 f'<span class="bv">{cx.span(sh, co, f)}</span></div>')
    o.append(f'<p class="refline">{cx.cell_html(blk["ref"]["label"])}：{cx.span(rs, rc, f)}</p></figure>')
    return ''.join(o)


def render_tornado(cx, blk):
    """敏感度龍捲風圖（CSS 對稱長條）：每個案例一條，依 |值| 由大到小排序。"""
    sh = blk['sheet']
    f = blk.get('fmt', 'sp1')
    items = []
    for r in expand_rows(blk['rows']):
        v = cx.b.cell(sh, f'{blk["value_col"]}{r}')
        if v is None:
            continue
        items.append((abs(v), r, v))
    items.sort(key=lambda t: (-t[0], t[1]))
    top = max(a for a, _, _ in items) or 1.0
    o = [f'<figure class="fig tor"><figcaption>{cx.cell_html(blk["title"])}</figcaption>']
    for a, r, v in items:
        if a == 0:
            continue
        w = a / top * 50
        pos = 'p' if v > 0 else 'n'
        side = f'left:50%;width:{w:.2f}%' if v > 0 else f'left:{50 - w:.2f}%;width:{w:.2f}%'
        o.append(f'<div class="trw"><span class="tlb">{cx.span(sh, f"{blk["label_col"]}{r}", "")}</span>'
                 f'<span class="ttr"><i class="{pos}" style="{side}"></i></span>'
                 f'<span class="tvv">{cx.span(sh, f"{blk["value_col"]}{r}", f)}</span></div>')
    o.append('</figure>')
    return ''.join(o)


def render_note(cx, blk):
    return f'<p class="note">{cx.cell_html(blk["parts"])}</p>'


def render_kv(cx, blk):
    """鍵值列表：[[標籤規格, 值規格], …]"""
    o = ['<dl class="kv">']
    for k, v in blk['items']:
        o.append(f'<div><dt>{cx.cell_html(k)}</dt><dd>{cx.cell_html(v)}</dd></div>')
    o.append('</dl>')
    return ''.join(o)


def render_scan(cx, blk):
    """列出指定工作表中含某關鍵字的儲存格（例如「未找到」），附列標籤與位置。"""
    needle = blk['needle']
    hits = []
    for sh, labcols in blk['sheets'].items():
        ws = cx.b.wb[sh]
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and needle in c.value:
                    lab = ' '.join(str(ws[f'{lc}{c.row}'].value) for lc in labcols if ws[f'{lc}{c.row}'].value is not None)
                    hits.append((sh, c.coordinate, c.row, lab))
    o = [f'<details class="more"><summary>{esc(blk["summary"])}</summary>' if blk.get('collapse') else '']
    o.append('<div class="tw"><table class="rt"><thead><tr><th>項目</th><th>內容</th></tr></thead><tbody>')
    for sh, co, r, lab in hits:
        labcols = blk['sheets'][sh]
        lab_cells = ' '.join(cx.span(sh, f'{lc}{r}', '') for lc in labcols if cx.b.cell(sh, f'{lc}{r}') is not None)
        o.append(f'<tr><th scope="row" data-l="項目">{lab_cells}</th><td data-l="內容">{cx.span(sh, co, "")}</td></tr>')
    o.append('</tbody></table></div>')
    if blk.get('collapse'):
        o.append('</details>')
    if not hits:
        o = [f'<p class="note">{esc(blk.get("empty", ""))}</p>']
    return ''.join(o)


UPSTREAM_RE = re.compile(r'(?:[\w\-]+/)+dist/[^\s｜；、，。（）()]+?\.xlsx')


def render_upstream(cx, blk, root):
    """上游模型版本：掃描指定工作表中引用的 dist/*.xlsx，列出首次出現位置、檔案是否存在、上游 README 標題與 Tokenomics 版本。"""
    seen = {}
    for sh in blk['sheets']:
        for row in cx.b.wb[sh].iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    for m in UPSTREAM_RE.finditer(c.value):
                        seen.setdefault(m.group(0), (sh, c.coordinate))
    o = ['<div class="tw"><table class="rt"><thead><tr><th>引用的上游成品</th><th>首次出現</th><th>repo 內現況</th><th>上游標題（README 頁首）</th></tr></thead><tbody>']
    for path in sorted(seen):
        sh, co = seen[path]
        cx.used.add(f'{sh}!{co}')
        full = os.path.join(root, path)
        title, state = '—', '不存在（引用已過時或路徑錯誤）'
        if os.path.exists(full):
            state = '存在'
            try:
                t = openpyxl.load_workbook(full, data_only=True)['README']['A1'].value
                title = str(t) if t else '—'
            except Exception:
                title = '（無法讀取）'
        else:
            d = os.path.dirname(full)
            if os.path.isdir(d):
                others = sorted(f for f in os.listdir(d) if f.endswith('.xlsx'))
                if others:
                    state = '不存在；現有：' + '、'.join(others)
        o.append(f'<tr><th scope="row" data-l="引用的上游成品"><code class="meta" data-meta="upstream">{esc(path)}</code></th>'
                 f'<td data-l="首次出現"><span class="meta" data-meta="upstream-ref" data-ref="{sh}!{co}" title="{sh}!{co}">{sh}!{co}</span></td>'
                 f'<td data-l="repo 內現況"><span class="meta" data-meta="upstream-state">{esc(state)}</span></td>'
                 f'<td data-l="上游標題（README 頁首）"><span class="meta" data-meta="upstream-title">{esc(title)}</span></td></tr>')
    o.append('</tbody></table></div>')
    return ''.join(o)


def render_list(cx, blk):
    tag = 'ol' if blk.get('ordered') else 'ul'   # ordered=true：編號清單
    return f'<{tag} class="items">' + ''.join(f'<li>{cx.cell_html(i)}</li>' for i in blk['items']) + f'</{tag}>'


# 只在橫幅有 ai_note 時輸出（HTML 版面規格 v1.1），既有成品逐位元組不變
AI_CSS = ".hero-ai{font-size:14px;color:#1f2328;font-weight:600}"


def render_hero(cx, h):
    """投資建議橫幅（html.json 頂層 hero，可選；Andy 2026-11-05）：全寬、大字、依建議著色。
    hero：{"rec": 投資建議規格, "tiles": [[標籤, 主值規格, 副值規格或空字串], …], "note": 規格或省略,
           "ai_note": 規格或省略（自用 AI 讀法 2 的依賴度一行，HTML 版面規格 v1.1；原則第 10 節第 52 條）}"""
    rec_txt = cx.text_of(h['rec'])
    tone = 'neg' if ('放空' in rec_txt or '收支不足' in rec_txt) else 'pos' if '買進' in rec_txt else 'neu'
    o = [f'<div class="hero hero-{tone}"><div class="hero-rec"><span class="hero-k">投資建議</span>'
         f'<span class="hero-v">{cx.cell_html(h["rec"])}</span></div><div class="hero-tiles">']
    for lab, main, sub in h['tiles']:
        o.append(f'<div class="hero-t"><div class="hero-l">{cx.cell_html(lab)}</div><div class="hero-n">{cx.cell_html(main)}</div>'
                 + (f'<div class="hero-s">{cx.cell_html(sub)}</div>' if sub else '') + '</div>')
    o.append('</div>')
    if h.get('ai_note'):
        o.append(f'<p class="hero-note hero-ai">{cx.cell_html(h["ai_note"])}</p>')
    if h.get('note'):
        o.append(f'<p class="hero-note">{cx.cell_html(h["note"])}</p>')
    o.append('</div>')
    return ''.join(o)


THEME_COLORS = ['#0b7a80', '#10a0a7', '#6fc7cb', '#b7e3e4', '#c29a2b', '#6f6bae', '#b4472f', '#8a8c8c', '#d9b65c', '#a3a1d0']


def render_theme(cx, blk):
    """分區主題圖（大圖，Andy 2026-11-11）：長條（stack＝true 堆疊，否則並列）＋折線；折線可用右軸（axis＝'r'）。
    每個長條段與點帶 data-ref／data-v（class pt）供檢核。"""
    xs = blk['x'] if isinstance(blk['x'], list) else [str(v) for v in cx.values(split_spec(blk['x'])[0])]
    for lab in (blk['x'] if isinstance(blk['x'], list) else []):
        check_prose(lab, 'x 軸標籤')
    n = len(xs)
    bars = blk.get('bars', [])
    lines = blk.get('lines', [])
    f, fr = blk.get('fmt', 'n0'), blk.get('fmt_r', blk.get('fmt', 'n0'))
    stack = blk.get('stack', True)
    W, H, L, R, T, B = 900, 360, 60, 60 if any(l.get('axis') == 'r' for l in lines) else 20, 16, 34
    bvals = [[cx.b.cell(a, c) for a, c in cx.b.expand(split_spec(b['y'])[0])] for b in bars]
    bvals = [[(v or 0) for v in row] for row in bvals]
    lvals = [[cx.b.cell(a, c) for a, c in cx.b.expand(split_spec(l['y'])[0])] for l in lines]
    left = [v for row in bvals for v in row] + [v for l, row in zip(lines, lvals) if l.get('axis') != 'r' for v in row if v is not None]
    if bars and stack:
        left += [sum(row[i] for row in bvals) for i in range(n)]
    right = [v for l, row in zip(lines, lvals) if l.get('axis') == 'r' for v in row if v is not None]
    tl = nice_ticks(min(0.0, min(left)) if left else 0.0, (max(left) if left else 1.0) * blk.get('pad', 1.0), 5)   # pad：左軸上方留白（讓右軸折線不壓到長條）
    tr = nice_ticks(min(0.0, min(right)), max(right), 5) if right else None
    lo, hi = tl[0], tl[-1]
    py = lambda v: T + (H - T - B) * (1 - (v - lo) / (hi - lo))
    pyr = (lambda v: T + (H - T - B) * (1 - (v - tr[0]) / (tr[-1] - tr[0]))) if tr else None
    cw = (W - L - R) / n
    px = lambda i: L + cw * (i + 0.5)
    o = [f'<svg viewBox="0 0 {W} {H}" class="chart big" role="img" aria-label="{esc(cx.text_of(blk["title"]))}">']
    for t in tl:
        o.append(f'<line x1="{L}" x2="{W - R}" y1="{py(t):.1f}" y2="{py(t):.1f}" class="grid"/>'
                 f'<text x="{L - 6}" y="{py(t) + 3.5:.1f}" class="ax" text-anchor="end">{fmt_value(t, f)}</text>')
    if tr:
        for t in tr:
            o.append(f'<text x="{W - R + 6}" y="{pyr(t) + 3.5:.1f}" class="ax">{fmt_value(t, fr)}</text>')
    for i, lab in enumerate(xs):
        o.append(f'<text x="{px(i):.1f}" y="{H - 12}" class="ax" text-anchor="middle">{esc(lab)}</text>')
    bw = cw * 0.62
    for i in range(n):
        base = 0.0
        for k, (b, row) in enumerate(zip(bars, bvals)):
            refs = cx.b.expand(split_spec(b['y'])[0])
            sh, co = refs[i]
            raw = cx.b.cell(sh, co)
            v = row[i]
            if stack:
                x0, w, y0, y1 = px(i) - bw / 2, bw, base + v, base
                base += v
            else:
                w = bw / len(bars)
                x0, y0, y1 = px(i) - bw / 2 + k * w, v, 0.0
            ref = f'{sh}!{co}'
            cx.used.add(ref)
            col = b.get('color', THEME_COLORS[k % len(THEME_COLORS)])
            dv = '' if raw is None else f' data-ref="{ref}" data-v="{raw!r}" data-f="{f}"'
            o.append(f'<rect x="{x0:.1f}" y="{py(max(y0, y1)):.1f}" width="{w:.1f}" height="{abs(py(y0) - py(y1)):.1f}" fill="{col}" class="pt bar"{dv}>'
                     f'<title>{esc(cx.text_of(b["name"]))}｜{esc(xs[i])}：{fmt_value(v, f)}（{ref}）</title></rect>')
        if stack and blk.get('total') and bars:      # 合計標籤須取自 xlsx 的合計列（不在產生器內加總）
            tsh, tco = cx.b.expand(split_spec(blk['total'])[0])[i]
            tv = cx.b.cell(tsh, tco)
            cx.used.add(f'{tsh}!{tco}')
            o.append(f'<text x="{px(i):.1f}" y="{py(base) - 5:.1f}" class="tot x" text-anchor="middle" data-ref="{tsh}!{tco}" data-v="{tv!r}" data-f="{f}">{fmt_value(tv, f)}</text>')
    for k, (l, row) in enumerate(zip(lines, lvals)):
        yf = pyr if l.get('axis') == 'r' else py
        ff = fr if l.get('axis') == 'r' else f
        refs = cx.b.expand(split_spec(l['y'])[0])
        col = l.get('color', ['#3b3838', '#c29a2b', '#6f6bae', '#b4472f'][k % 4])
        dash = ' stroke-dasharray="6 4"' if l.get('dash') else ''
        pts = [(px(i + l.get('offset', 0)), yf(v)) for i, v in enumerate(row) if v is not None]
        o.append(f'<polyline points="{" ".join(f"{a:.1f},{b:.1f}" for a, b in pts)}" fill="none" stroke="{col}" stroke-width="2.4"{dash}/>')
        nm = cx.text_of(l['name'])
        j = 0
        for i, ((sh, co), v) in enumerate(zip(refs, row)):
            if v is None:
                continue
            ref = f'{sh}!{co}'
            cx.used.add(ref)
            x = px(i + l.get('offset', 0))
            o.append(f'<circle cx="{x:.1f}" cy="{yf(v):.1f}" r="4" fill="{col}" class="pt" data-ref="{ref}" data-v="{v!r}" data-f="{ff}">'
                     f'<title>{esc(nm)}｜{esc(xs[i + l.get("offset", 0)])}：{fmt_value(v, ff)}（{ref}）</title></circle>')
            if l.get('labels'):
                o.append(f'<text x="{x:.1f}" y="{yf(v) + (18 if l.get('labels') == 'below' else -8):.1f}" class="vl x" text-anchor="middle" fill="{col}" data-ref="{ref}" data-v="{v!r}" data-f="{ff}">{fmt_value(v, ff)}</text>')
    o.append('</svg>')
    leg = ''.join(f'<li><i style="background:{b.get("color", THEME_COLORS[k % len(THEME_COLORS)])}"></i>{cx.cell_html(b["name"])}</li>' for k, b in enumerate(bars))
    leg += ''.join(f'<li><i class="ln" style="border-color:{l.get("color", ["#3b3838", "#c29a2b", "#6f6bae", "#b4472f"][k % 4])}{";border-top-style:dashed" if l.get("dash") else ""}"></i>'
                   f'{cx.cell_html(l["name"])}{"（右軸）" if l.get("axis") == "r" else ""}</li>' for k, l in enumerate(lines))
    cap = f'<p class="tcap">{cx.cell_html(blk["caption"])}</p>' if blk.get('caption') else ''
    return (f'<figure class="theme"><figcaption>{cx.cell_html(blk["title"])}</figcaption><div class="tsc">{"".join(o)}</div>'
            f'<ul class="tleg">{leg}</ul>{cap}</figure>')


def render_compare(cx, blk):
    """對照長條（每組自有尺度）：groups＝[{name, fmt, items:[{name, v, ref:true?}]}]；標記 ref 的項目（指引、共識）以灰色顯示。"""
    out = [f'<figure class="theme"><figcaption>{cx.cell_html(blk["title"])}</figcaption><div class="cmp">']
    for g in blk['groups']:
        f = g.get('fmt', 'n0')
        vals = [cx.values(split_spec(it['v'])[0])[0] for it in g['items']]
        mx = max(abs(v) for v in vals if isinstance(v, (int, float))) or 1
        out.append(f'<div class="cg"><div class="cgh">{cx.cell_html(g["name"])}</div>')
        for it, v in zip(g['items'], vals):
            w = abs(v) / mx * 100
            cls = 'r' if it.get('ref') else 'm'
            refs, _ = split_spec(it['v'])
            out.append(f'<div class="br"><span>{cx.cell_html(it["name"])}</span><div class="bt"><i class="{cls}" style="width:{w:.2f}%"></i></div>'
                       f'<span class="bv">{cx.cell_html(refs and "@" + refs + "|" + f)}</span></div>')
        out.append('</div>')
    cap = f'<p class="tcap">{cx.cell_html(blk["caption"])}</p>' if blk.get('caption') else ''
    out.append(f'</div>{cap}</figure>')
    return ''.join(out)


def render_block(cx, blk, cfg, root):
    if blk.get('collapse') and blk['type'] != 'scan':
        inner = render_block(cx, {k: v for k, v in blk.items() if k not in ('collapse', 'summary')}, cfg, root)
        check_prose(blk['summary'], 'summary')
        return f'<details class="more"><summary>{esc(blk["summary"])}</summary>{inner}</details>'
    t = blk['type']
    if t == 'table':
        return render_table(cx, blk)
    if t == 'pl':
        return render_pl(cx, blk, cfg)
    if t == 'bars':
        return render_bars(cx, blk)
    if t == 'tornado':
        return render_tornado(cx, blk)
    if t == 'note':
        return render_note(cx, blk)
    if t == 'kv':
        return render_kv(cx, blk)
    if t == 'scan':
        return render_scan(cx, blk)
    if t == 'upstream':
        return render_upstream(cx, blk, root)
    if t == 'list':
        return render_list(cx, blk)
    if t == 'lines':
        return render_lines(cx, blk, cfg['scenarios'])
    if t == 'theme':
        return render_theme(cx, blk)
    if t == 'compare':
        return render_compare(cx, blk)
    if t == 'h3':
        return f'<h3>{cx.cell_html(blk["text"])}</h3>'
    raise ValueError(f'未知區塊型別：{t}')


# ───────────────────────── 頁面 ─────────────────────────
CSS = """
:root{--bg:#fff;--surface:#fff;--soft:#f7fbfb;--ink:#3b3838;--muted:#5f605f;--line:#dce3e3;--grid:#ecefef;--axis:#8a8c8c;
--accent:#0b7a80;--teal:#10a0a7;--c0:#10a0a7;--c1:#c29a2b;--c2:#6f6bae;--pos:#0b7a80;--neg:#b4472f;
--tv:#d5ecea;--ti:#fbf8ee;--ta:#f6e3de;--tn:#f1f0f8;--td:#f2f9f9;color-scheme:light}
*{box-sizing:border-box}html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.65 'Noto Sans TC','Microsoft JhengHei','微軟正黑體','PingFang TC',Arial,sans-serif;overflow-wrap:anywhere}
.wrap{max-width:78rem;margin:0 auto;padding:32px 24px 40px}
header.top{display:grid;gap:8px}
h1{font-size:28px;line-height:1.3;font-weight:400;margin:0}
h2{font-size:20px;font-weight:600;margin:0 0 8px;counter-increment:sec}
h2::before{content:counter(sec) ". ";color:var(--teal)}
h3{font-size:15px;font-weight:600;margin:18px 0 4px}
main{counter-reset:sec;display:grid;gap:24px;margin-top:20px}
section{background:var(--surface);border:1px solid var(--line);border-radius:6px;padding:20px;min-width:0}
.meta{color:var(--muted);font-size:13px}
.thesis{font-size:18px;line-height:1.6;margin:4px 0 8px;max-width:70ch}
.q{border-left:3px solid var(--teal);padding:4px 12px;margin:10px 0;background:var(--soft);max-width:70ch}
nav.toc{display:flex;flex-wrap:wrap;gap:2px;border-bottom:1px solid var(--line);margin-top:10px;font-size:14px}
nav.toc a{color:var(--muted);text-decoration:none;padding:6px 12px;border-bottom:2px solid transparent;margin-bottom:-1px}
nav.toc a:hover{color:var(--ink);border-bottom-color:var(--teal)}
.refs-toggle{position:absolute;opacity:0;pointer-events:none}
label.refs{justify-self:start;font-size:13px;color:var(--accent);cursor:pointer;border:1px solid var(--teal);border-radius:999px;padding:2px 14px;background:#fff}
.refs-toggle:focus-visible+.wrap label.refs{outline:2px solid var(--accent)}
.refs-toggle:checked+.wrap label.refs{background:var(--teal);color:#fff}
.refs-toggle:checked+.wrap .x::after,.refs-toggle:checked+.wrap .lbl::after,.refs-toggle:checked+.wrap .na::after{content:" [" attr(data-ref) "]";font-size:10px;color:var(--muted);font-weight:400;white-space:nowrap}
p.cap{margin:14px 0 4px;font-weight:600}
.note{margin:8px 0;max-width:80ch}
.tw{max-width:100%;overflow-x:auto}
table{border-collapse:collapse;width:100%;font-size:14px}
th,td{border-bottom:1px solid var(--line);padding:6px 8px;text-align:left;vertical-align:top}
thead th{font-size:13px;color:var(--muted);font-weight:600}
table.grid td,table.grid thead th:not(:first-child){text-align:right;font-variant-numeric:tabular-nums}
table.grid th,table.grid td{padding:6px 4px;font-size:13px}
tr.sub th{background:var(--soft);font-size:13px;color:var(--accent);padding-top:10px}
.x,.meta[data-ref]{font-variant-numeric:tabular-nums;white-space:nowrap}
.tag{display:inline-block;white-space:nowrap;padding:0 8px;border-radius:9px;font-size:12px;background:var(--grid)}
.tv{background:var(--tv)}.ti{background:var(--ti)}.ta{background:var(--ta)}.tn{background:var(--tn)}.td{background:var(--td)}
.na{color:var(--axis)}
.charts{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,17rem),1fr));gap:16px;margin:14px 0}
figure.fig{margin:0;padding:12px 14px;background:var(--surface);border:1px solid var(--line);border-radius:6px}
figcaption{font-weight:600;font-size:14px;margin-bottom:4px}
svg.chart{width:100%;height:auto;display:block;max-width:560px}
.chart text{font-size:10px;fill:var(--muted)}
.chart .grid{stroke:var(--grid);stroke-width:1}
.ln{fill:none;stroke-width:2.2;stroke-linejoin:round}
.s0.ln,.legend .s0{stroke:var(--c0)}.s1.ln,.legend .s1{stroke:var(--c1)}.s2.ln,.legend .s2{stroke:var(--c2)}
.ln.s1{stroke-dasharray:6 3}.ln.s2{stroke-dasharray:2 3}.legend .s1{stroke-dasharray:6 3}.legend .s2{stroke-dasharray:2 3}
.pt.s0{fill:var(--c0)}.pt.s1{fill:var(--c1)}.pt.s2{fill:var(--c2)}.pt{stroke:#fff;stroke-width:1.5}
.legend{display:flex;gap:4px 14px;flex-wrap:wrap;list-style:none;margin:2px 0 0;padding:0;font-size:13px;color:var(--muted)}
.legend li{display:flex;align-items:center;gap:5px}.sw{width:28px;height:8px}.sw .ln{stroke-width:2}
.br,.trw{display:grid;gap:8px;align-items:center;margin:6px 0}
.br{grid-template-columns:minmax(70px,150px) 1fr 72px}
.bt{position:relative;height:18px;background:var(--grid);border-radius:3px}
.bt i{position:absolute;left:0;top:0;bottom:0;border-radius:3px}
.bt i.s0{background:var(--c0)}.bt i.s1{background:var(--c1)}.bt i.s2{background:var(--c2)}
.bt b{position:absolute;top:-3px;bottom:-3px;width:2px;background:var(--ink)}
.bv{text-align:right;font-variant-numeric:tabular-nums}
.refline{font-size:13px;color:var(--muted);margin:4px 0 0}
.trw{grid-template-columns:minmax(110px,230px) 1fr 62px;font-size:13px}
.ttr{position:relative;height:14px;background:var(--grid);border-radius:3px}
.ttr::before{content:"";position:absolute;left:50%;top:0;bottom:0;width:1px;background:var(--axis)}
.ttr i{position:absolute;top:2px;bottom:2px;border-radius:2px}.ttr i.p{background:var(--teal)}.ttr i.n{background:var(--neg)}
.tvv{text-align:right;font-variant-numeric:tabular-nums}
dl.kv{margin:8px 0}dl.kv div{display:flex;gap:10px;border-bottom:1px solid var(--line);padding:4px 0;flex-wrap:wrap}
dl.kv dt{color:var(--muted);flex:1 1 180px}dl.kv dd{margin:0;flex:2 1 220px}
ul.items{margin:6px 0;padding-left:20px}
details.more{margin:8px 0}summary{cursor:pointer;color:var(--accent)}
code.meta{font:12px/1.4 ui-monospace,Menlo,monospace}
footer{margin:24px 0 0;padding-top:10px;border-top:1px solid var(--line);font-size:12px;color:var(--muted)}
@media (max-width:640px){
 body{font-size:14.5px}h1{font-size:23px}.wrap{padding:16px}section{padding:14px}
 table.rt thead{display:none}
 table.rt,table.rt tbody,table.rt tr,table.rt th,table.rt td{display:block;width:100%}
 table.rt tr{border:1px solid var(--line);border-radius:6px;margin:8px 0;padding:4px 8px;background:var(--surface)}
 table.rt th,table.rt td{border:0;padding:3px 0;display:flex;gap:10px;justify-content:space-between;text-align:right}
 table.rt th[scope=row]{font-weight:700;text-align:left;display:block}
 table.rt td::before{content:attr(data-l);color:var(--muted);font-size:12px;text-align:left;flex:0 0 38%}
 table.rt td .cv{max-width:62%;text-align:right}
 tr.sub{border:0!important;background:none!important;padding:0!important}tr.sub th{display:block}
 .br{grid-template-columns:70px 1fr 58px}.trw{grid-template-columns:96px 1fr 54px;font-size:12px}
}
"""

# html.json 頂層 wide_text=true：說明文字（命題、核心問題、各章說明、note）延伸到頁面全寬（Andy 2026-11-05，SpaceX 版面）
THEME_CSS = """
.theme{margin:6px 0 16px;padding:14px 16px;border:1px solid var(--line);border-radius:6px;background:var(--surface)}.theme figcaption{font-weight:600;font-size:15px;margin-bottom:6px}.tsc{overflow-x:auto}svg.chart.big{max-width:900px;min-width:560px;margin:0 auto}svg.chart .tot{font-size:11px;fill:var(--ink);font-weight:600}svg.chart .vl{font-size:11px;font-weight:600}.tleg{display:flex;flex-wrap:wrap;gap:4px 16px;list-style:none;margin:6px 0 0;padding:0;font-size:13px;color:var(--muted)}.tleg i{display:inline-block;width:12px;height:12px;border-radius:2px;margin-right:6px;vertical-align:-1px}.tleg i.ln{width:22px;height:0;border-top:2.4px solid;border-radius:0;vertical-align:4px}.tcap{font-size:13px;color:var(--muted);margin:6px 0 0}.cmp{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,22rem),1fr));gap:12px 28px}.cgh{font-weight:600;font-size:14px;margin:4px 0}.bt i.r{background:var(--axis)}.bt i.m{background:var(--c0)}
"""

WIDE_CSS = """
.thesis,.q,.note{max-width:none}
ol.items{margin:6px 0;padding-left:24px}ol.items li,ul.items li{margin:2px 0}
.hero{margin-top:18px;border-radius:8px;padding:18px 22px;border:2px solid var(--line);display:grid;gap:12px}
.hero-neg{background:#fbece8;border-color:#b4472f}.hero-pos{background:#e6f4f3;border-color:#0b7a80}.hero-neu{background:#fbf6e8;border-color:#c29a2b}
.hero-rec{display:flex;align-items:baseline;gap:14px;flex-wrap:wrap}
.hero-k{font-size:16px;color:var(--muted);font-weight:600}
.hero-v{font-size:34px;font-weight:700;line-height:1.2}
.hero-neg .hero-v,.hero-neg .hero-s{color:#b4472f}.hero-pos .hero-v,.hero-pos .hero-s{color:#0b7a80}.hero-neu .hero-v,.hero-neu .hero-s{color:#8a6d16}
.hero-tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,15rem),1fr));gap:12px}
.hero-t{background:#fff;border:1px solid var(--line);border-radius:6px;padding:10px 14px}
.hero-l{font-size:14px;color:var(--muted)}
.hero-n{font-size:28px;font-weight:700;font-variant-numeric:tabular-nums;line-height:1.3}
.hero-s{font-size:17px;font-weight:600;font-variant-numeric:tabular-nums}
.hero-note{margin:0;font-size:13px;color:var(--muted)}
@media (max-width:640px){table.rt tr.sub th{justify-content:flex-start;text-align:left}.hero-v{font-size:28px}.hero-n{font-size:24px}}
"""

REQUIRED_IDS = ['thesis', 'drivers', 'pl', 'valuation', 'recon', 'sens', 'status']
# 版面規格 v1.3（2026-11-14；可選，頂層 "layout": "v1.3"）：命題→估值→敏感度→反向 DCF→損益→驅動→個別模組（0 到多個，id 自訂）→對帳→資料狀態
V13_HEAD = ['thesis', 'valuation', 'sens', 'revdcf', 'pl', 'drivers']
V13_TAIL = ['recon', 'status']


def v13_order_error(ids):
    """v1.3 章節順序檢查；回傳錯誤訊息或 None。"""
    if ids[:len(V13_HEAD)] != V13_HEAD or ids[-len(V13_TAIL):] != V13_TAIL or len(ids) != len(set(ids)):
        return f'layout v1.3 的 sections 須為 {V13_HEAD} → 個別模組 → {V13_TAIL}，實際 {ids}'
    return None
OPTIONAL_IDS = ['conclusion', 'funding', 'revdcf']   # conclusion＝命題回覆（可選；放在 sens 與 status 之間，Andy 2026-11-05）；funding＝收支型個股的資金缺口與融資（緊接 thesis 之後）、revdcf＝反向 DCF（緊接 valuation 之後）（可選；Andy 2026-11-10，首例 Oracle v0.5.0）


def find_xlsx(cdir, override=None):
    if override:
        return override
    import glob
    fs = sorted(glob.glob(os.path.join(cdir, 'dist', '*.xlsx')))
    if len(fs) != 1:
        raise SystemExit(f'{cdir}/dist 內應剛好一個 xlsx，實際 {len(fs)} 個')
    return fs[0]


def sha12(path):
    h = hashlib.sha256()
    with open(path, 'rb') as fh:
        h.update(fh.read())
    return h.hexdigest()[:12]


def parse_name(path):
    """20261023_Advantest模型_v0.2.1.xlsx → ('2026-10-23', '0.2.1')"""
    m = re.match(r'^(\d{4})(\d{2})(\d{2})_.*_v([\d.]+)\.xlsx$', os.path.basename(path))
    if not m:
        raise SystemExit(f'成品檔名不符「日期_名稱_v版本.xlsx」：{path}')
    return f'{m.group(1)}-{m.group(2)}-{m.group(3)}', m.group(4)


def build(cdir, xlsx=None, root=None):
    cfg = json.load(open(os.path.join(cdir, 'html.json'), encoding='utf-8'))
    xlsx = find_xlsx(cdir, xlsx)
    root = root or os.path.dirname(os.path.abspath(cdir))
    date, ver = parse_name(xlsx)
    cx = Ctx(Book(xlsx), {'$version': 'v' + ver, '$date': date, '$sha': sha12(xlsx), '$xlsx': os.path.basename(xlsx)})
    ids = [s['id'] for s in cfg['sections'] if s['id'] not in OPTIONAL_IDS]
    allids = [s['id'] for s in cfg['sections']]
    if cfg.get('layout') == 'v1.3':
        err = v13_order_error(allids)
        if err:
            raise SystemExit(err)
    elif 'conclusion' in allids and allids[allids.index('conclusion') - 1:allids.index('conclusion') + 2] != ['sens', 'conclusion', 'status']:
        raise SystemExit(f'conclusion 須放在 sens 與 status 之間，實際 {allids}')
    if cfg.get('layout') != 'v1.3':
        for opt, prev in (('funding', 'thesis'), ('revdcf', 'valuation')):
            if opt in allids and allids[allids.index(opt) - 1] != prev:
                raise SystemExit(f'{opt} 須緊接在 {prev} 之後，實際 {allids}')
        if ids[:len(REQUIRED_IDS)] != REQUIRED_IDS:
            raise SystemExit(f'sections 的 id 須依序為 {REQUIRED_IDS}，實際 {ids}')
    check_prose(cfg['question'], 'question')
    for s in cfg['scenarios']:
        check_prose(s['name'], 'scenarios.name')
    secs = []
    for s in cfg['sections']:
        check_prose(s['title'], f'sections.{s["id"]}.title')
        body = []
        if s['id'] == 'thesis':
            body.append(f'<p class="thesis" id="thesis-line">{cx.cell_html(cfg["thesis"])}</p>'
                        f'<p class="q"><strong>核心問題：</strong>{esc(cfg["question"])}</p>')
        if s.get('intro'):
            body.append(f'<p class="note">{cx.cell_html(s["intro"])}</p>')
        for blk in s.get('blocks', []):
            body.append(render_block(cx, blk, cfg, root))
        secs.append(f'<section id="{s["id"]}"><h2>{esc(s["title"])}</h2>{"".join(body)}</section>')
    toc = ''.join(f'<a href="#{s["id"]}">{esc(s["title"])}</a>' for s in cfg['sections'])
    title = f'{cfg["company"]}模型'
    page = f"""<!doctype html>
<html lang="zh-Hant-TW">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light">
<title>{esc(title)}</title>
<style>{CSS}{WIDE_CSS if (cfg.get('wide_text') or cfg.get('hero')) else ''}{AI_CSS if (cfg.get('hero') or {}).get('ai_note') else ''}{THEME_CSS if any(b.get('type') in ('theme', 'compare') for sc in cfg['sections'] for b in sc['blocks']) else ''}</style>
</head>
<body>
<input type="checkbox" id="refs" class="refs-toggle" aria-label="顯示每個數字的來源儲存格">
<div class="wrap">
<header class="top">
<h1>{esc(title)}</h1>
<div class="meta"><span class="meta" data-meta="ticker">{esc(cfg["ticker"])}</span>｜版本 {cx.part('$version')}｜成品日期 {cx.part('$date')}</div>
<label class="refs" for="refs">顯示來源儲存格（工作表!儲存格）</label>
<nav class="toc" aria-label="章節">{toc}</nav>
</header>
{render_hero(cx, cfg['hero']) if cfg.get('hero') else ''}<main>
{"".join(secs)}
</main>
<footer>
<p>本頁由 <code class="meta">tools/company_html/build_company_html.py</code> 自 Excel 成品的重算值產生；數字的計算事實以 Excel 為準。每個數字的來源「工作表!儲存格」標在滑鼠懸停文字，或按上方按鈕顯示。</p>
<p>來源 xlsx：{cx.part('$xlsx')}｜雜湊指紋：{cx.part('$sha')}</p>
</footer>
</div>
</body>
</html>
"""
    return page, xlsx


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('company_dir')
    ap.add_argument('--xlsx')
    ap.add_argument('--out')
    a = ap.parse_args()
    page, xlsx = build(a.company_dir, a.xlsx)
    out = a.out or os.path.splitext(xlsx)[0] + '.html'
    with open(out, 'w', encoding='utf-8', newline='\n') as fh:
        fh.write(page)
    print(f'{out}（{len(page.encode("utf-8")):,} bytes）')


if __name__ == '__main__':
    main()

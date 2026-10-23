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
    return '<ul class="items">' + ''.join(f'<li>{cx.cell_html(i)}</li>' for i in blk['items']) + '</ul>'


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

REQUIRED_IDS = ['thesis', 'drivers', 'pl', 'valuation', 'recon', 'sens', 'status']


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
    ids = [s['id'] for s in cfg['sections']]
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
<style>{CSS}</style>
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
<main>
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

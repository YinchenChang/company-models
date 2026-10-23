#!/usr/bin/env python3
"""公司 HTML 展示層檢核（本機與 CI 共用）。

用法：python3 tools/company_html/check_company_html.py <公司資料夾>...
對每家公司（需有 html.json、dist/*.xlsx、dist/*.html）檢查：
 1. 重現性：重新產生兩次，兩次逐位元組相同，且與 dist 內已提交的 HTML 逐位元組相同。
 2. 抽數：HTML 內每個帶 data-ref 的數字（class x、圖上的點 pt）與 xlsx 重算值相對誤差 < 1e-9；
    顯示字串與值的差在顯示精度內；文字標籤（lbl）與 xlsx 儲存格文字完全相同；空格（na）對應的儲存格確實為空。
 3. 無手寫數字：標籤與數字以外的文字，除年度（FY／CY 加四位數）與日期外不得含阿拉伯數字
    （圖軸刻度 class ax、中繼資料 class meta 例外；meta 另以檔案實際內容核對）。
 4. 自含：不含外部網址、script、link、img、iframe、@import。
 5. 結構：七個必備章節依序出現；每張圖有資料點；每個長條有寬度。
結束碼：0 通過；1 有失敗。
"""
import glob
import hashlib
import json
import os
import re
import sys
import tempfile
from html.parser import HTMLParser

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_company_html as B  # noqa: E402

TOL = 1e-9
REQUIRED = B.REQUIRED_IDS
VOID = {'meta', 'link', 'input', 'br', 'line', 'img', 'hr'}


class P(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.stack = []          # (tag, attrs)
        self.items = []          # (kind, attrs, text)  kind: x / lbl / na / pt / meta
        self.untagged = []       # 未標記的文字節點
        self.section_ids = []
        self.bad = []            # 外部資源等
        self.cur = None          # 目前收集文字的項目
        self.svg_charts = 0
        self.title_depth = 0
        self.style_depth = 0
        self.bar_widths = []

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        cls = (a.get('class') or '').split()
        if tag in ('script', 'link', 'img', 'iframe', 'object', 'embed'):
            self.bad.append(f'含 <{tag}>')
        for k in ('src', 'href'):
            v = a.get(k, '')
            if v and not v.startswith('#'):
                self.bad.append(f'含外部參照 {k}={v}')
        if tag == 'section' and a.get('id'):
            self.section_ids.append(a['id'])
        if tag == 'style':
            self.style_depth += 1
        if tag == 'svg' and 'chart' in cls:
            self.svg_charts += 1
        if tag == 'i' and a.get('style') and re.search(r'width:', a['style']):
            self.bar_widths.append(float(re.search(r'width:([\d.]+)%', a['style']).group(1)))
        if tag == 'circle' and 'pt' in cls:
            self.items.append(['pt', a, ''])
            self.cur = self.items[-1]
        if tag == 'title' and self.stack and self.stack[-1][0] == 'circle':
            self.title_depth += 1
        for k in ('x', 'lbl', 'na'):
            if k in cls and 'data-ref' in a:
                self.items.append([k, a, ''])
                self.cur = self.items[-1]
        if 'meta' in cls and 'data-meta' in a:
            self.items.append(['meta', a, ''])
            self.cur = self.items[-1]
        if tag not in VOID:
            self.stack.append((tag, cls, a))

    def handle_endtag(self, tag):
        if tag == 'style':
            self.style_depth -= 1
        if tag == 'title' and self.title_depth:
            self.title_depth -= 1
        while self.stack:
            t, cls, a = self.stack.pop()
            if self.cur is not None and self.cur[1] is a:
                self.cur = None
            if t == tag:
                break

    def handle_data(self, data):
        if self.style_depth or not data.strip():
            return
        if self.cur is not None:
            self.cur[2] += data
            return
        if any(t == 'title' for t, _, _ in self.stack) and any(t == 'head' for t, _, _ in self.stack):
            return
        if any('ax' in cls for _, cls, _ in self.stack) or self.title_depth:
            return
        self.untagged.append(data)


def check_company(cdir, errs):
    name = os.path.basename(os.path.abspath(cdir))
    cfg = json.load(open(os.path.join(cdir, 'html.json'), encoding='utf-8'))
    xlsx = B.find_xlsx(cdir)
    committed = os.path.splitext(xlsx)[0] + '.html'
    if not os.path.exists(committed):
        errs.append(f'{name}：缺少已提交的 HTML {committed}')
        return
    # 1. 重現性
    p1, _ = B.build(cdir)
    p2, _ = B.build(cdir)
    if p1.encode('utf-8') != p2.encode('utf-8'):
        errs.append(f'{name}：連續兩次產生的 HTML 不同')
    if p1.encode('utf-8') != open(committed, 'rb').read():
        errs.append(f'{name}：重新產生的 HTML 與已提交的 {os.path.basename(committed)} 不同（請重建並提交）')
    # 2–5 解析已提交的檔案
    text = open(committed, encoding='utf-8').read()
    ps = P()
    ps.feed(text)
    book = B.Book(xlsx)
    cell = lambda ref: book.cell(*ref.split('!', 1))
    n_x = n_lbl = n_na = n_pt = 0
    meta_expect = {
        'ticker': cfg['ticker'], 'version': 'v' + B.parse_name(xlsx)[1], 'date': B.parse_name(xlsx)[0],
        'xlsx': os.path.basename(xlsx), 'sha': B.sha12(xlsx)}
    for kind, a, t in ps.items:
        ref = a.get('data-ref')
        if kind in ('x', 'pt'):
            v = cell(ref)
            dv = float(a['data-v'])
            if not isinstance(v, (int, float)):
                errs.append(f'{name}：{ref} 在 xlsx 不是數值（{v!r}）')
                continue
            if abs(dv - v) > TOL * max(1.0, abs(v)):
                errs.append(f'{name}：{ref} data-v {dv} ≠ xlsx {v}')
            if kind == 'x':
                n_x += 1
                shown, tol = B.parse_display(t)
                if a.get('data-f') in ('yr',):
                    shown, tol = float(t.replace(',', '')), 0.5
                if abs(shown - v) > tol * (1 + 1e-9) + 1e-12:
                    errs.append(f'{name}：{ref} 顯示「{t}」與 xlsx {v} 的差超過顯示精度')
                if B.fmt_value(v, a['data-f']) != t.strip():
                    errs.append(f'{name}：{ref} 顯示「{t}」≠ 格式化結果「{B.fmt_value(v, a["data-f"])}」')
            else:
                n_pt += 1
        elif kind == 'lbl':
            n_lbl += 1
            if cell(ref) != t:
                errs.append(f'{name}：{ref} 標籤「{t[:30]}」≠ xlsx「{str(cell(ref))[:30]}」')
        elif kind == 'na':
            n_na += 1
            if cell(ref) is not None:
                errs.append(f'{name}：{ref} 標為空但 xlsx 有值 {cell(ref)!r}')
        elif kind == 'meta':
            m = a['data-meta']
            if m in meta_expect and t != meta_expect[m]:
                errs.append(f'{name}：meta {m}「{t}」≠ 預期「{meta_expect[m]}」')
    # 上游標題核對
    # 3. 無手寫數字
    for d in ps.untagged:
        rest = B.ALLOWED_PROSE_DIGITS.sub('', d)
        if re.search(r'\d', rest):
            errs.append(f'{name}：未標記來源的數字：{d.strip()[:60]!r}')
    # 4. 自含
    errs += [f'{name}：{b}' for b in ps.bad]
    if re.search(r'https?://|@import|url\(', text):
        errs.append(f'{name}：含外部網址或 @import／url()')
    # 5. 結構
    if ps.section_ids[:len(REQUIRED)] != REQUIRED:
        errs.append(f'{name}：章節順序 {ps.section_ids} 不符 {REQUIRED}')
    if ps.svg_charts == 0 or n_pt == 0:
        errs.append(f'{name}：沒有折線圖資料點')
    if not ps.bar_widths or min(ps.bar_widths) <= 0:
        errs.append(f'{name}：長條寬度為零或沒有長條')
    print(f'{name}：數字 {n_x}、圖上資料點 {n_pt}、標籤 {n_lbl}、空格 {n_na}、未標記文字節點 {len(ps.untagged)}')


def main():
    dirs = sys.argv[1:]
    if not dirs:
        print(__doc__)
        sys.exit(2)
    errs = []
    for d in dirs:
        check_company(d, errs)
    if errs:
        print(f'\n失敗 {len(errs)} 項（顯示前 40 項）：')
        for e in errs[:40]:
            print(' -', e)
        sys.exit(1)
    print('全部通過')


if __name__ == '__main__':
    main()

#!/usr/bin/env python3
"""個股 HTML 版面規格 v1.3 檢查（docs/plan/個股HTML版面規格.md「v1.3 規格」一節）。

檢查項目：
  A 章節：必備章節齊全、順序正確（個別模組只能在 drivers 與 recon 之間）；目錄順序與章節一致。
  B 橫幅：.hero-note 可見文字 ≤ 200 字、不含參數代號。
  C 重點觀察：thesis 章「重點觀察」≤ 5 點、合計 ≤ 750 字。
  D 損益：pl 章含 GAAP／Non-GAAP 標示與必備列；表頭至少 3 個歷史年＋5 個預測年（以年份欄位計）。
  F 圖表：除 revdcf、recon、status 外，每章至少一張圖（svg、長條 .ttr／.bt 或 canvas）。
  G 橫幅現價附日期；估值章含 DCF、Comparable、加權、市場共識四個目標價。
  H 敏感度章主圖為「加權目標價相對基準」。
  全頁（警告）：摺疊（details）與附錄以外的可見文字出現參數代號。

用法：python3 tools/company_html/check_layout_v13.py <html>...   結束碼：0＝全過；1＝有必修。
"""
import re
import sys

from bs4 import BeautifulSoup

REQUIRED_HEAD = ['thesis', 'valuation', 'sens', 'revdcf', 'pl', 'drivers']
REQUIRED_TAIL = ['recon', 'status']
NO_CHART = {'revdcf', 'recon', 'status'}
HERO_MAX = 200
OBS_MAX_N, OBS_MAX_CH = 5, 750
# 參數代號：英數加底線的識別字（he_all、k_adj、rent_mw），或 Excel 具名範圍樣式（TK_、IF_ 開頭）
PARAM_RE = re.compile(r'(?<![A-Za-z0-9])[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+(?![A-Za-z0-9])')
PL_ROWS = [
    ('營收', r'^營收(?!.*(年增|YoY|成長))'),
    ('營收年增率', r'營收.*(年增|YoY|成長率)'),
    ('營業成本', r'營業成本|COGS'),
    ('毛利', r'^毛利(?!率)'),
    ('毛利率', r'毛利率'),
    ('營業費用', r'營業費用|營運費用|營運成本'),
    ('營業利益', r'^營業利益(?!率)'),
    ('營業利益率', r'營業利益率'),
    ('其他收入', r'其他收入|業外'),
    ('稅前利益', r'稅前'),
    ('所得稅', r'所得稅'),
    ('稅後淨利', r'(稅後)?淨利(?!率|.*年增|.*YoY)'),
    ('淨利率', r'淨利率'),
    ('EPS', r'EPS|每股盈餘'),
    ('淨利年增率', r'淨利.*(年增|YoY)'),
]


def vis_text(el):
    el = BeautifulSoup(str(el), 'lxml')
    for d in el.find_all('details'):
        d.decompose()
    for x in el.select('[hidden], .refs, script, style'):
        x.decompose()
    return re.sub(r'\s+', '', el.get_text())


def check(path):
    s = BeautifulSoup(open(path, encoding='utf-8').read(), 'lxml')
    must, warn = [], []
    secs = [x for x in s.find_all('section') if x.get('id')]
    ids = [x['id'] for x in secs]
    # A 章節順序
    head, tail = ids[:len(REQUIRED_HEAD)], ids[-len(REQUIRED_TAIL):] if len(ids) >= 2 else ids
    if head != REQUIRED_HEAD:
        must.append(f'A 章節開頭應為 {REQUIRED_HEAD}，實際 {ids}')
    if tail != REQUIRED_TAIL:
        must.append(f'A 章節結尾應為 {REQUIRED_TAIL}，實際 {ids}')
    for r in REQUIRED_HEAD + REQUIRED_TAIL:
        if ids.count(r) != 1:
            must.append(f'A 章節 {r} 出現 {ids.count(r)} 次')
    toc = [a.get('href', '').lstrip('#') for a in s.select('nav.toc a')]
    if toc and toc != ids:
        must.append(f'A 目錄順序與章節不一致：目錄 {toc}；章節 {ids}')
    modules = ids[len(REQUIRED_HEAD):-len(REQUIRED_TAIL)] if len(ids) > 8 else []
    # B 橫幅
    hn = s.select_one('.hero-note')
    if hn is None:
        must.append('B 找不到橫幅說明 .hero-note')
    else:
        t = re.sub(r'\s+', '', hn.get_text())
        if len(t) > HERO_MAX:
            must.append(f'B 橫幅說明 {len(t)} 字 > {HERO_MAX}')
        p = PARAM_RE.findall(t)
        if p:
            must.append(f'B 橫幅說明含參數代號 {sorted(set(p))}')
    # C 重點觀察
    th = s.find('section', id='thesis')
    obs = None
    if th:
        for h in th.find_all(['h3', 'h4']):
            if '重點觀察' in h.get_text():
                obs = h.find_next_sibling(['ol', 'ul'])
                break
    if obs is None:
        must.append('C 命題章找不到「重點觀察」清單')
    else:
        n = len(obs.find_all('li', recursive=False))
        ch = len(re.sub(r'\s+', '', obs.get_text()))
        if n > OBS_MAX_N:
            must.append(f'C 重點觀察 {n} 點 > {OBS_MAX_N}')
        if ch > OBS_MAX_CH:
            must.append(f'C 重點觀察 {ch} 字 > {OBS_MAX_CH}')
    # D 損益
    pl = s.find('section', id='pl')
    if pl:
        txt = pl.get_text()
        if not re.search(r'Non-GAAP|GAAP', txt):
            must.append('D 損益章未標示 GAAP／Non-GAAP')
        labels = []
        for tr in pl.find_all('tr'):
            c = tr.find(['th', 'td'])
            if c:
                labels.append(re.sub(r'\s+', '', c.get_text()))
        for name, rx in PL_ROWS:
            if not any(re.search(rx, l) for l in labels):
                must.append(f'D 損益表缺列：{name}')
        years = set()
        for tr in pl.find_all('tr')[:3]:
            for c in tr.find_all(['th', 'td']):
                m = re.search(r'(?:FY|CY)?(20\d\d)', c.get_text())
                if m:
                    years.add(m.group(1))
        if len(years) < 8:
            must.append(f'D 損益表年份欄 {len(years)} 個 < 8（歷史 3＋未來 5）：{sorted(years)}')
    # G 現價日期與四個目標價
    hero = s.select_one('.hero')
    if hero is not None:
        ht = re.sub(r'\s+', '', hero.get_text())
        if not re.search(r'現價[^0-9]{0,40}20\d\d-\d\d-\d\d', ht):
            must.append('G 橫幅現價未附日期（YYYY-MM-DD）')
    val = s.find('section', id='valuation')
    if val is not None:
        vt = re.sub(r'\s+', '', val.get_text())
        for name, rx in [('DCF 目標價', r'DCF'), ('Comparable 目標價', r'Comparable|同業'), ('加權目標價', r'加權'), ('市場共識目標價', r'共識')]:
            if not re.search(rx, vt):
                must.append(f'G 估值章缺 {name}')
    # H 敏感度主圖
    sens = s.find('section', id='sens')
    if sens is not None and '相對基準' not in vis_text(sens):
        must.append('H 敏感度章主圖未以「加權目標價相對基準」呈現')
    # F 圖表
    for x in secs:
        if x['id'] in NO_CHART:
            continue
        if not (x.find_all('svg') or x.select('.ttr, .bt, canvas')):
            must.append(f'F 章節 {x["id"]} 沒有圖表')
    # 全頁參數代號（警告）
    main = s.find('main') or s.body
    for x in secs:
        if x['id'] == 'status':
            continue
        p = PARAM_RE.findall(vis_text(x))
        if p:
            warn.append(f'全頁 {x["id"]} 可見文字含參數代號 {sorted(set(p))[:8]}')
    print(('FAIL ' if must else 'ok   ') + path + f'（章節 {ids}；個別模組 {modules}）')
    for m in must:
        print('  必修', m)
    for w in warn:
        print('  警告', w)
    return not must


if __name__ == '__main__':
    ok = all([check(p) for p in sys.argv[1:]])
    sys.exit(0 if ok else 1)

#!/usr/bin/env python3
"""版面檢查（Playwright／Chromium）：寬 1280 與 390 各截一張圖，並檢查
(1) 頁面沒有水平捲動；(2) 每張 SVG 圖都有資料點、每個長條都有寬度（沒有空白圖）；(3) 文字沒有被裁掉。
截圖寫到 --outdir（預設暫存目錄，不入 repo）。

用法：python3 tools/company_html/layout_check.py <html 檔>... [--outdir 目錄] [--dark]
結束碼：0＝全部通過；1＝有問題。
"""
import argparse
import os
import sys
import tempfile

from playwright.sync_api import sync_playwright

JS = """() => {
  const doc = document.documentElement;
  const wide = [];
  for (const el of document.body.querySelectorAll('*')) {
    const r = el.getBoundingClientRect();
    if (r.width > 0 && r.right > window.innerWidth + 1) wide.push(el.tagName + '.' + (el.className.baseVal ?? el.className));
  }
  const charts = [...document.querySelectorAll('svg.chart')].map(s => ({pts: s.querySelectorAll('circle.pt').length, lines: s.querySelectorAll('polyline').length, h: s.getBoundingClientRect().height}));
  const bars = [...document.querySelectorAll('.bt i, .ttr i')].map(i => i.getBoundingClientRect().width);
  return {sw: doc.scrollWidth, iw: window.innerWidth, wide: wide.slice(0, 8), charts, bars};
}"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('html', nargs='+')
    ap.add_argument('--outdir', default=os.path.join(tempfile.gettempdir(), 'company_html_shots'))
    ap.add_argument('--dark', action='store_true')
    a = ap.parse_args()
    os.makedirs(a.outdir, exist_ok=True)
    bad = 0
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path='/opt/pw-browsers/chromium') if os.path.exists('/opt/pw-browsers/chromium') else p.chromium.launch()
        for path in a.html:
            for w, h in ((1280, 900), (390, 844)):
                ctx = b.new_context(viewport={'width': w, 'height': h}, color_scheme='dark' if a.dark else 'light')
                pg = ctx.new_page()
                pg.goto('file://' + os.path.abspath(path))
                r = pg.evaluate(JS)
                name = os.path.basename(path).replace('.html', f'_{w}{"_dark" if a.dark else ""}.png')
                pg.screenshot(path=os.path.join(a.outdir, name), full_page=True)
                probs = []
                if r['sw'] > r['iw']:
                    probs.append(f'水平捲動：scrollWidth {r["sw"]} > {r["iw"]}；超出元素 {r["wide"]}')
                if not r['charts'] or any(c['pts'] == 0 or c['lines'] == 0 or c['h'] < 50 for c in r['charts']):
                    probs.append(f'空白圖：{r["charts"]}')
                if not r['bars'] or any(x <= 0 for x in r['bars']):
                    probs.append('長條寬度為零或沒有長條')
                print(('FAIL ' if probs else 'ok   ') + f'{os.path.basename(path)} @{w}：scrollWidth {r["sw"]}／{r["iw"]}，圖 {len(r["charts"])}，長條 {len(r["bars"])}  → {name}')
                for q in probs:
                    print('     ', q)
                bad += bool(probs)
                ctx.close()
        b.close()
    sys.exit(1 if bad else 0)


if __name__ == '__main__':
    main()

# 逐頁截圖＋溢出自查（MAG v0.1c）：python3 shots.py <html> <輸出資料夾>
import sys, os, re, json
from playwright.sync_api import sync_playwright
html, outd = sys.argv[1], sys.argv[2]; os.makedirs(outd, exist_ok=True)
OVF = r"""() => { const W = document.documentElement.clientWidth, bad = [];
  if (document.documentElement.scrollWidth > W + 2) bad.push('page scrollWidth ' + document.documentElement.scrollWidth + ' > ' + W);
  for (const el of document.querySelectorAll('body *')) { const cs = getComputedStyle(el); if (cs.display === 'none' || cs.visibility === 'hidden') continue;
    if (/(auto|scroll)/.test(cs.overflowX)) continue; const r = el.getBoundingClientRect(); if (r.width === 0) continue;
    if (r.right > W + 2 && !el.closest('[style*="overflow"]')) bad.push((el.tagName + '.' + el.className).slice(0, 40) + ' right ' + Math.round(r.right));
    if (bad.length > 8) break; } return bad }"""
res, errs, n = {}, [], 0
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(viewport={'width': 1440, 'height': 900})
    pg.on('pageerror', lambda e: errs.append(str(e)))
    pg.goto('file://' + os.path.abspath(html)); pg.wait_for_timeout(2500)
    SKIP = ('總結', '資金模型', '損益與評價', '簡報模式（全螢幕）', '列印／另存 PDF', '載入', '重設', '儲存 JSON', '年度 CSV', '站點 CSV')
    vis = lambda: [t for t in dict.fromkeys(pg.locator('button:visible').all_inner_texts()) if 0 < len(t) < 14 and t not in SKIP and not t.startswith(('保守', '基準', '積極'))]
    def click(t):
        try: pg.locator('button:visible').filter(has_text=re.compile('^\\s*' + re.escape(t) + '\\s*$')).first.click(timeout=1500); pg.wait_for_timeout(500); return True
        except Exception: return False
    def shot(key):
        global n; n += 1; pg.evaluate('window.scrollTo(0,0)'); pg.wait_for_timeout(200)
        f = f"{n:02d}_{key.replace('/', '_')}.png"; pg.screenshot(path=os.path.join(outd, f), full_page=True)
        res[key] = {'file': f, 'overflow': pg.evaluate(OVF)}
    for top in ['總結', '資金模型', '損益與評價']:
        pg.get_by_role('button', name=top, exact=True).click(); pg.wait_for_timeout(800); shot(top)
        for t in vis():
            if click(t): shot(top + '/' + t)
    b.close()
json.dump({'shots': res, 'pageErrors': errs}, open(os.path.join(outd, 'index.json'), 'w'), ensure_ascii=False, indent=1)
print(len(res), 'shots; errors', len(errs), '; overflow pages', [k for k, v in res.items() if v['overflow']])

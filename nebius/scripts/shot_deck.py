# 簡報逐頁截圖（v0.2；工作單 v0.1d 步驟 4）：頁首一張（1280×720 視窗、畫面模式）＋總結頁每張投影片一張（列印版面，1280×720）。
# 同時自查：每張投影片內的元素是否超出投影片邊界（溢出）、文字是否被截斷（scrollWidth > clientWidth）、是否出現 NaN／undefined／Infinity。
# 用法：python3 scripts/shot_deck.py 檔案.html 輸出資料夾
import sys, os, json
from playwright.sync_api import sync_playwright

JS_CHECK = r"""
(sl) => {
  const R = sl.getBoundingClientRect(), out = [];
  for (const el of sl.querySelectorAll('*')) {
    const r = el.getBoundingClientRect();
    if (!r.width || !r.height) continue;
    const t = (el.innerText || '').trim().slice(0, 40);
    if (r.right > R.right + 1 || r.bottom > R.bottom + 1 || r.left < R.left - 1 || r.top < R.top - 1)
      out.push(`超出投影片：${el.tagName} ${Math.round(r.right - R.right)}/${Math.round(r.bottom - R.bottom)} 「${t}」`);
    const cs = getComputedStyle(el);
    if (el.children.length === 0 && t && (el.scrollWidth > el.clientWidth + 2) && cs.overflow !== 'visible')
      out.push(`文字被截斷：${el.tagName}「${t}」`);
  }
  // 正式簡報頁（非「附錄｜」）：說明文字字級 ≥ 18px（頁尾除外）、標題 ≥ 36px
  const kick = (sl.firstElementChild || {}).innerText || '';
  if (!kick.startsWith('附錄')) {
    const foot = sl.lastElementChild, h2 = sl.querySelector('h2');
    if (h2 && parseFloat(getComputedStyle(h2).fontSize) < 36) out.push(`標題字級 ${getComputedStyle(h2).fontSize} < 36px`);
    for (const el of sl.querySelectorAll('*')) {
      if (foot && foot.contains(el)) continue;
      const own = [...el.childNodes].some(n => n.nodeType === 3 && n.textContent.trim());
      const fs = parseFloat(getComputedStyle(el).fontSize);
      if (own && fs < 18) out.push(`字級 ${fs}px < 18：「${el.innerText.trim().slice(0, 20)}」`);
    }
  }
  const txt = sl.innerText;
  for (const w of ['NaN', 'undefined', 'Infinity']) if (txt.includes(w)) out.push(`出現 ${w}`);
  return { issues: [...new Set(out)].slice(0, 12), title: (sl.querySelector('h2') || {}).innerText || '', kicker: (sl.firstElementChild || {}).innerText || '' };
}
"""

def main(html, out):
    os.makedirs(out, exist_ok=True)
    res = []
    with sync_playwright() as p:
        try: b = p.chromium.launch()
        except Exception: b = p.chromium.launch(executable_path=os.environ.get('CHROMIUM_PATH', '/opt/pw-browsers/chromium'))
        pg = b.new_page(viewport={'width': 1280, 'height': 720})
        errs = []; pg.on('pageerror', lambda e: errs.append(str(e)))
        pg.goto('file://' + os.path.abspath(html)); pg.wait_for_timeout(3000)
        pg.screenshot(path=os.path.join(out, '00_頁首.png'))
        # 簡報模式：第 1 頁進場動畫播完後截一張（確認動畫結束時內容完整可見）
        pg.get_by_role('button', name='簡報模式（全螢幕）').click(); pg.wait_for_timeout(1800)
        pg.screenshot(path=os.path.join(out, '00_簡報模式.png'))
        pg.keyboard.press('Escape'); pg.wait_for_timeout(300)
        pg.emulate_media(media='print'); pg.wait_for_timeout(800)
        slides = pg.locator('.sumQ-slide')
        n = slides.count()
        for i in range(n):
            el = slides.nth(i)
            info = el.evaluate(JS_CHECK)
            fn = f'{i + 1:02d}.png'
            el.screenshot(path=os.path.join(out, fn))
            res.append({'file': fn, **info})
        b.close()
    json.dump({'slides': res, 'errors': errs}, open(os.path.join(out, 'check.json'), 'w'), ensure_ascii=False, indent=1)
    bad = 0
    for r in res:
        print(f"{r['file']}｜{r['kicker']}｜{r['title']}")
        for x in r['issues']: print('   ! ' + x); bad += 1
    print(f'{len(res)} 張投影片；問題 {bad} 項；頁面錯誤 {len(errs)}', errs[:3])
    return 1 if bad or errs else 0

if __name__ == '__main__':
    sys.exit(main(sys.argv[1], sys.argv[2]))

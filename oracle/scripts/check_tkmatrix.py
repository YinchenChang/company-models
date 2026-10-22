# v0.3.1：HTML 3 × 3 矩陣次層檢查——未動過任何輸入的 HTML：評價 → 目標價頁可展開「3 × 3 容量 × 價格矩陣」，快照檢查顯示「通過」，
# 且表內九格加權目標價與 Excel「每MW收入_錨定」G 區同一份（tk_sens.json）。用法：python3 scripts/check_tkmatrix.py <HTML>
import sys, os, re, json
from playwright.sync_api import sync_playwright
H = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TS = json.load(open(os.path.join(H, 'tk_sens.json'), encoding='utf-8'))
with sync_playwright() as p:
    try: b = p.chromium.launch()
    except Exception: b = p.chromium.launch(executable_path=os.environ.get('CHROMIUM_PATH', '/opt/pw-browsers/chromium'))
    pg = b.new_page(viewport={'width': 1440, 'height': 900}); errs = []
    pg.on('pageerror', lambda e: errs.append(str(e)))
    pg.goto('file://' + os.path.abspath(sys.argv[1])); pg.wait_for_timeout(2500)
    pg.get_by_role('button', name='損益與評價', exact=True).click(); pg.wait_for_timeout(600)
    pg.locator('button:visible').filter(has_text=re.compile(r'^\s*評價\s*$')).first.click(); pg.wait_for_timeout(400)
    pg.locator('button:visible').filter(has_text=re.compile(r'^\s*目標價\s*$')).first.click(); pg.wait_for_timeout(600)
    pg.get_by_text('3 × 3 容量 × 價格矩陣（建置快照）').first.click(); pg.wait_for_timeout(400)
    t = pg.inner_text('body'); b.close()
bad = []
if '快照檢查：目前情境 × 價格基準的快照 − 即時加權目標價 → 通過' not in t: bad.append('快照檢查未顯示「通過」')
for sc in ('low', 'base', 'high'):
    for px in ('low', 'base', 'high'):
        v = f"${TS['matrix'][sc][px]['tgt']:.1f}"
        if v not in t: bad.append(f'矩陣缺 {sc}×{px} 加權目標價 {v}')
if errs: bad.append(f'頁面錯誤 {errs}')
print('3 × 3 矩陣次層檢查：' + ('通過（快照檢查「通過」、九格目標價與 tk_sens.json 一致）' if not bad else '失敗 ' + '；'.join(bad)))
sys.exit(1 if bad else 0)

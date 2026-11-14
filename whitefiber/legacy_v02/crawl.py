# 逐頁逐分頁擷取畫面文字（正規化版本號），用來比對兩版 HTML 的顯示結果是否完全相同
import sys, os, json, re
from playwright.sync_api import sync_playwright
def crawl(path):
    out = {}; errs = []
    with sync_playwright() as p:
        try: b = p.chromium.launch()
        except Exception:  # 雲端環境預裝的 Chromium 版本可能與 playwright 套件不符：改用 CHROMIUM_PATH 或 /opt/pw-browsers/chromium
            b = p.chromium.launch(executable_path=os.environ.get('CHROMIUM_PATH', '/opt/pw-browsers/chromium'))
        pg = b.new_page(viewport={'width': 1440, 'height': 900})
        pg.on('pageerror', lambda e: errs.append(str(e)))
        pg.goto('file://' + os.path.abspath(path)); pg.wait_for_timeout(2500)
        norm = lambda t: re.sub(r'v\d\.\d', 'vX', t)
        # 擷取整頁可見文字：<main> 只是資金模型頁的一個區塊，總結頁與損益與評價頁的內容不在其中
        def snap(key): out[key] = norm(pg.inner_text('body'))
        SKIP = ('總結', '資金模型', '損益與評價', '簡報模式（全螢幕）', '列印／另存 PDF', '載入', '重設', '儲存 JSON', '年度 CSV', '站點 CSV')
        def visible():  # 可點的次級按鈕（排除頁籤、工具列與情境按鈕）
            return [t for t in dict.fromkeys(pg.locator('button:visible').all_inner_texts())
                    if 0 < len(t) < 14 and t not in SKIP and not t.startswith(('保守', '基準', '積極'))]
        def click(t):  # 以可見按鈕的文字精確比對（同名按鈕可能有隱藏副本；部分按鈕以 role 名稱查不到）
            try:
                pg.locator('button:visible').filter(has_text=re.compile('^\\s*' + re.escape(t) + '\\s*$')).first.click(timeout=1500)
                pg.wait_for_timeout(400); return True
            except Exception: return False
        for top in ['總結', '資金模型', '損益與評價']:
            pg.get_by_role('button', name=top, exact=True).click(); pg.wait_for_timeout(700); snap(top)
            # 兩層選單：先逐一點擊第一層並記錄當時可見的按鈕；某項獨有（不在各狀態共同可見集合內）的按鈕即其次級分頁
            level1, states = visible(), {}
            for t in level1:
                if click(t): states[t] = visible(); snap(top + '/' + t)
            common = set.intersection(*(set(v) for v in states.values())) if states else set()
            for t, vis in states.items():
                subs = [u for u in vis if u not in common and top + '/' + u not in out]
                for u in subs:
                    if click(t) and click(u): snap(top + '/' + t + '/' + u)
        b.close()
    return out, errs
if __name__ == '__main__':
    o, e = crawl(sys.argv[1]); json.dump({'pages': o, 'errors': e}, open(sys.argv[2], 'w'), ensure_ascii=False); print(len(o), 'views; errors', e)

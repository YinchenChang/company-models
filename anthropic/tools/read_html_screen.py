#!/usr/bin/env python3
"""HTML 一頁摘要的『逐節讀畫面』檢查（v0.1.3 B）：以 Playwright 開啟 HTML，對每一節（section）：
1. 捲動到該節並對該節元素截圖（桌機 1280、手機 390 寬）；
2. 讀出畫面上實際顯示的數字文字（innerText，非 data-v），與該元素的 data-v 以同一格式化規則比對；
3. 記錄該節標題、數字個數、是否有「—」（空值）、元素是否在可視範圍內、頁面是否水平捲動。
離線檢查：任何非 file:/data: 請求都記錄。輸出 JSON（stdout）與截圖。
用法：[CHROMIUM_PATH=<chrome 路徑>] python3 tools/read_html_screen.py dist/<檔名>.html --out docs/reports/img --prefix 20261022_v0.1.3"""
import argparse
import json
import os
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
from build_html import fmt  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("html", type=Path)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--prefix", required=True)
    a = ap.parse_args()
    a.out.mkdir(parents=True, exist_ok=True)
    url = a.html.resolve().as_uri()
    bad, report = [], {}
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path=os.environ.get("CHROMIUM_PATH") or None)       # 雲端環境：/opt/pw-browsers/chromium-*/chrome-linux/chrome
        for w, tag in ((1280, "desktop"), (390, "mobile")):
            ctx = b.new_context(viewport={"width": w, "height": 900}, color_scheme="light", device_scale_factor=1 if w > 600 else 2)
            page = ctx.new_page()
            page.on("request", lambda r: bad.append(r.url) if not r.url.startswith(("file:", "data:", "about:")) else None)
            page.goto(url, wait_until="load")
            sw = page.evaluate("document.documentElement.scrollWidth")
            secs = page.query_selector_all("main > section, main > header")
            out = []
            for el in secs:
                sid = el.get_attribute("id") or "header"
                el.scroll_into_view_if_needed()
                shot = a.out / f"{a.prefix}_{tag}_節{sid}.png"
                el.screenshot(path=str(shot))
                head = (el.query_selector("h2") or el.query_selector("h1")).inner_text().strip().replace("\n", " ")
                nums = el.query_selector_all("[data-ref]")
                shown, mism, empty = [], [], 0
                for n in nums:
                    txt = n.evaluate("e => { const c = e.cloneNode(true); c.querySelectorAll('title').forEach(t => t.remove()); return c.textContent.trim(); }")      # HTML span 與 SVG text 皆適用（略過 SVG title）
                    kind = n.get_attribute("data-fmt")
                    dv = n.get_attribute("data-v")
                    try:
                        v = eval(dv, {"__builtins__": {}}, {})
                        want = fmt(v, kind)
                    except Exception:
                        want = None
                    if want is not None and txt != want:
                        mism.append((n.get_attribute("data-ref"), txt, want))
                    if txt in ("", "—"):
                        empty += 1
                    shown.append(txt)
                box = el.bounding_box()
                out.append(dict(section=sid, heading=head, numbers=len(nums), empty=empty, mismatch=mism, first=shown[:10],
                                width=round(box["width"]) if box else None, screenshot=str(shot)))
            report[tag] = dict(viewport=w, scrollWidth=sw, horizontal_scroll=sw > w, sections=out)
            full = a.out / f"{a.prefix}_{tag}_整頁.png"
            page.screenshot(path=str(full), full_page=True)
            ctx.close()
        b.close()
    report["external_requests"] = bad
    print(json.dumps(report, ensure_ascii=False, indent=1))
    ok = not bad and all(not r["horizontal_scroll"] and all(not s["mismatch"] for s in r["sections"]) for k, r in report.items() if k in ("desktop", "mobile"))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()

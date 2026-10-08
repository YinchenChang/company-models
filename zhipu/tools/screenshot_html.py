#!/usr/bin/env python3
"""HTML 一頁摘要的版面截圖（S6）：桌面 1280 寬、手機 390 寬 × 淺色、深色，整頁截圖。
離線檢查：攔截所有網路請求，任何非 file: 請求都記錄並以非零狀態結束；同時檢查水平捲動。
用法：python3 tools/screenshot_html.py dist/<檔名>.html --out docs/reports/img"""
import argparse
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("html", type=Path)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--prefix", default="20261008_v0.6_成品")
    a = ap.parse_args()
    a.out.mkdir(parents=True, exist_ok=True)
    url = a.html.resolve().as_uri()
    bad, problems = [], []
    with sync_playwright() as p:
        b = p.chromium.launch()
        for w, tag in ((1280, "desktop"), (390, "mobile")):
            for scheme in ("light", "dark"):
                ctx = b.new_context(viewport={"width": w, "height": 900}, color_scheme=scheme,
                                    device_scale_factor=1 if w > 600 else 2)
                page = ctx.new_page()
                page.on("request", lambda r: bad.append(r.url) if not r.url.startswith(("file:", "data:", "about:")) else None)
                page.goto(url, wait_until="load")
                sw = page.evaluate("document.documentElement.scrollWidth")
                if sw > w:
                    problems.append(f"{tag}-{scheme}: 水平捲動 scrollWidth={sw} > {w}")
                f = a.out / f"{a.prefix}_{tag}_{scheme}.png"
                page.screenshot(path=str(f), full_page=True)
                print(f"{f}  scrollWidth={sw}")
                ctx.close()
        b.close()
    if bad:
        problems.append(f"外部請求 {len(bad)} 個：{bad[:5]}")
    for x in problems:
        print("問題：", x, file=sys.stderr)
    sys.exit(1 if problems else 0)


if __name__ == "__main__":
    main()

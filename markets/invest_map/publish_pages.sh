#!/usr/bin/env bash
# 把 dist/index.html 發布到 gh-pages 分支（GitHub Pages：https://yinchenchang.github.io/company-models/）
# 另發布 Siliconomics 展示層：https://yinchenchang.github.io/company-models/siliconomics/
# 每次以只含網站檔案的全新歷史覆寫 gh-pages（成品分支，不保留歷史；原始碼歷史在 main）。
# 用法：bash markets/invest_map/publish_pages.sh
set -euo pipefail
ROOT=$(git rev-parse --show-toplevel)
SHA=$(git -C "$ROOT" rev-parse --short HEAD)
URL=$(git -C "$ROOT" remote get-url origin)
TMP=$(mktemp -d)
cp "$ROOT/markets/invest_map/dist/index.html" "$TMP/index.html"
touch "$TMP/.nojekyll"
# 「我們的模型」分頁：複製 data.json 列出的各公司模型 HTML 到 models/<資料夾>/
python3 - "$ROOT" "$TMP" <<'PY'
import json, os, shutil, sys
root, tmp = sys.argv[1], sys.argv[2]
d = json.load(open(os.path.join(root, "markets/invest_map/dist/data.json"), encoding="utf-8"))
for tk, m in d.get("models", {}).items():
    src = os.path.join(root, m["folder"], "dist", m["file"])
    dst = os.path.join(tmp, "models", m["folder"])
    os.makedirs(dst, exist_ok=True)
    shutil.copy2(src, os.path.join(dst, m["file"]))
    x = os.path.join(root, m["folder"], "dist", m["xlsx"])
    if os.path.exists(x):
        shutil.copy2(x, os.path.join(dst, m["xlsx"]))
print("模型 HTML：", len(d.get("models", {})), "份")
PY
# Siliconomics 展示層：發布時由 model/CURRENT 的 xlsx 直接產生（不另存 HTML 副本，避免與 xlsx 不同步）
python3 "$ROOT/siliconomics/tools/build_html.py" "$TMP/siliconomics/index.html"
# Wattonomics（瓦帳）互動頁：wattonomics/index.html（Andy 2026-10-30 要求公司內可開；與地圖同一 Pages 站，地圖重發時一併帶上）
WV=$(head -1 "$ROOT/wattonomics/model/CURRENT" | sed -nE 's/.*_v([0-9.]+)\.xlsx$/\1/p')
WX="$ROOT/wattonomics/data/explorer_v$WV.json"
[ -n "$WV" ] && [ -f "$WX" ] || { echo "停止：找不到 Wattonomics explorer（CURRENT 版本＝「$WV」）；不發布，避免覆寫後少了 wattonomics/ 頁" >&2; exit 1; }
mkdir -p "$TMP/wattonomics"
python3 "$ROOT/wattonomics/explorer/build_page.py" "$WX" "$TMP/wattonomics/index.html"
# build_page.py 產出的是 artifact 片段；補上完整文件骨架（doctype、charset、viewport），公司瀏覽器才不會亂碼
python3 - "$TMP/wattonomics/index.html" <<'PY'
import sys
p = sys.argv[1]; body = open(p, encoding="utf-8").read()
if not body.lstrip().lower().startswith("<!doctype"):
    body = ('<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1">'
            '<style>body{margin:0}</style></head><body>\n' + body + '\n</body></html>\n')
    open(p, "w", encoding="utf-8").write(body)
PY
echo "Wattonomics 互動頁：v$WV"
printf '# AI 投資地圖（GitHub Pages 發布分支）\n\n本分支只放網站成品，由 main 的 `markets/invest_map/publish_pages.sh` 產生（來源 %s）。研究整理，非投資建議。\n' "$SHA" > "$TMP/README.md"
cd "$TMP"
git init -q -b gh-pages
git add -A
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -q -m "發布 AI 投資地圖（來源 $SHA）"
git push -q -f "$URL" gh-pages:gh-pages
cd / && rm -rf "$TMP"
echo "已發布到 gh-pages（來源 $SHA）"

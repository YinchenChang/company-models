#!/usr/bin/env bash
# 把 dist/index.html 發布到 gh-pages 分支（GitHub Pages：https://yinchenchang.github.io/company-models/）
# 每次以只含網站檔案的全新歷史覆寫 gh-pages（成品分支，不保留歷史；原始碼歷史在 main）。
# 用法：bash markets/invest_map/publish_pages.sh
set -euo pipefail
ROOT=$(git rev-parse --show-toplevel)
SHA=$(git -C "$ROOT" rev-parse --short HEAD)
URL=$(git -C "$ROOT" remote get-url origin)
TMP=$(mktemp -d)
cp "$ROOT/markets/invest_map/dist/index.html" "$TMP/index.html"
touch "$TMP/.nojekyll"
# 各公司完整收支模型（HTML＋Excel），以資料夾名為穩定網址：models/<資料夾>.html／.xlsx
mkdir -p "$TMP/models"
for d in "$ROOT"/*/dist; do
  f=$(basename "$(dirname "$d")")
  h=$(ls "$d"/*.html 2>/dev/null | head -1); x=$(ls "$d"/*.xlsx 2>/dev/null | head -1)
  [ -n "$h" ] && cp "$h" "$TMP/models/$f.html"
  [ -n "$x" ] && cp "$x" "$TMP/models/$f.xlsx"
done
printf '# AI 投資地圖（GitHub Pages 發布分支）\n\n本分支只放網站成品，由 main 的 `markets/invest_map/publish_pages.sh` 產生（來源 %s）。研究整理，非投資建議。\n' "$SHA" > "$TMP/README.md"
cd "$TMP"
git init -q -b gh-pages
git add -A
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -q -m "發布 AI 投資地圖（來源 $SHA）"
git push -q -f "$URL" gh-pages:gh-pages
cd / && rm -rf "$TMP"
echo "已發布到 gh-pages（來源 $SHA）"

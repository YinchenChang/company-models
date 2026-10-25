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
print("模型 HTML：", len(d.get("models", {})), "份")
PY
printf '# AI 投資地圖（GitHub Pages 發布分支）\n\n本分支只放網站成品，由 main 的 `markets/invest_map/publish_pages.sh` 產生（來源 %s）。研究整理，非投資建議。\n' "$SHA" > "$TMP/README.md"
cd "$TMP"
git init -q -b gh-pages
git add -A
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -q -m "發布 AI 投資地圖（來源 $SHA）"
git push -q -f "$URL" gh-pages:gh-pages
cd / && rm -rf "$TMP"
echo "已發布到 gh-pages（來源 $SHA）"

#!/usr/bin/env bash
# 發布實體 AI 地圖到 gh-pages 的 physical/ 子資料夾：在現有 gh-pages 上「追加一個 commit」，不 force、只動 physical/。
# 用法（在含 physical_ai/map/dist 成品的工作目錄）：bash physical_ai/map/publish_physical_map.sh
# 需要 Andy 明確授權發布（AI實體應用 Project 規則第 11 條；2026-11-17 08:10 已授權 v0.1）。
set -euo pipefail
ROOT=$(git rev-parse --show-toplevel)
SHA=$(git -C "$ROOT" rev-parse --short HEAD)
SRC="$ROOT/physical_ai/map/dist"
[ -f "$SRC/index.html" ] && [ -f "$SRC/data.json" ] || { echo "停止：缺 dist/index.html 或 data.json" >&2; exit 1; }
python3 "$ROOT/physical_ai/map/check_map.py" >/dev/null || { echo "停止：check_map.py 未通過" >&2; exit 1; }
VER=$(python3 -c "import json;print(json.load(open('$SRC/data.json',encoding='utf-8'))['meta'].get('version','v?'))")
WT=$(mktemp -d)
git -C "$ROOT" fetch -q origin gh-pages
git -C "$ROOT" worktree add -q --detach "$WT" origin/gh-pages
trap 'git -C "$ROOT" worktree remove --force "$WT" >/dev/null 2>&1 || true' EXIT
mkdir -p "$WT/physical"
cp "$SRC/index.html" "$SRC/data.json" "$WT/physical/"
cd "$WT"
git add -A physical
if git diff --cached --quiet; then echo "無變更，不發布"; exit 0; fi
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -q -m "發布實體 AI 地圖 $VER（來源 $SHA；只更新 physical/）"
git push -q origin HEAD:gh-pages
echo "已發布實體 AI 地圖 $VER 到 gh-pages/physical（來源 $SHA）：$(git rev-parse --short HEAD)"

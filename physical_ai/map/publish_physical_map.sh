#!/usr/bin/env bash
# 發布實體 AI 地圖到 gh-pages 的 physical_ai/ 子資料夾（v0.2 起；與 AI 算力地圖、AI 軟體地圖約定的網址），舊位置 physical/ 放同一份頁面（<base> 指向 physical_ai/，不轉址）。
# 在現有 gh-pages 上「追加一個 commit」，不 force、只動 physical_ai/（含 models/）與 physical/。
# 用法（在含 physical_ai/map/dist 成品的工作目錄）：bash physical_ai/map/publish_physical_map.sh
# 需要 Andy 明確授權發布（AI實體應用 Project 規則第 11 條；2026-11-17 08:10 授權發布；13:58 要求改版）。
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
mkdir -p "$WT/physical_ai" "$WT/physical"
cp "$SRC/index.html" "$SRC/data.json" "$WT/physical_ai/"
# v0.3：本專案個股模型 HTML 與同版 xlsx 複製到 physical_ai/models/<資料夾>/（整個 models/ 重建，舊版不留）
rm -rf "$WT/physical_ai/models"
python3 - "$SRC/data.json" "$ROOT" "$WT" <<'PY'
import json, os, shutil, sys
d = json.load(open(sys.argv[1], encoding="utf-8")); root, wt = sys.argv[2], sys.argv[3]
for k, m in d["models"].items():
    if not m["url"].startswith("models/"):
        continue
    dst = os.path.join(wt, "physical_ai", "models", m["folder"]); os.makedirs(dst, exist_ok=True)
    for fn in (m["file"], m["xlsx"]):
        shutil.copy2(os.path.join(root, "physical_ai", m["folder"], "dist", fn), os.path.join(dst, fn))
PY
rm -f "$WT/physical/data.json"
# v0.3：舊位置 physical/ 不再用轉址頁（換頁時會多跳一頁），改放同一份頁面並以 <base> 指向 physical_ai/（相對連結、模型頁都解析到 physical_ai/）。
# AI 算力地圖與 AI 軟體地圖的連結改直連 physical_ai/ 之後（PR #713），這一份只服務舊書籤。
python3 - "$SRC/index.html" "$WT/physical/index.html" <<'PY'
import sys
t = open(sys.argv[1], encoding="utf-8").read()
k = '<meta charset="utf-8">'
assert t.count(k) == 1
open(sys.argv[2], "w", encoding="utf-8").write(t.replace(k, k + '\n<base href="../physical_ai/"><link rel="canonical" href="../physical_ai/">'))
PY
cd "$WT"
git add -A physical_ai physical
if git diff --cached --quiet; then echo "無變更，不發布"; exit 0; fi
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -q -m "發布實體 AI 地圖 $VER（來源 $SHA；只更新 physical_ai/，physical/ 同步）"
git push -q origin HEAD:gh-pages
echo "已發布實體 AI 地圖 $VER 到 gh-pages/physical_ai（來源 $SHA）：$(git rev-parse --short HEAD)"

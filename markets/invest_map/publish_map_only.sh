#!/usr/bin/env bash
# 只發布 AI 投資地圖（index.html 與 models/）到 gh-pages：在現有 gh-pages 上「追加一個 commit」，不 force、不動 wattonomics/、siliconomics/。
# 先檢查 data.json 列出的每一份模型 HTML／xlsx 都存在，缺任何一份就停止（2026-11-04：v0.6.22 發布時 Anthropic 檔已被新版取代，複製中斷仍推送，已回復）。
# 用法（在 main 的工作目錄）：bash markets/invest_map/publish_map_only.sh
# 需要 Andy 在對話中明確說「發布」（Project Instructions 第 6 條）。
set -euo pipefail
ROOT=$(git rev-parse --show-toplevel)
SHA=$(git -C "$ROOT" rev-parse --short HEAD)
VER=$(python3 -c "import json;print(json.load(open('$ROOT/markets/invest_map/dist/data.json',encoding='utf-8'))['meta']['version'])")

# 1) 先檢查：缺檔就停，不建任何 commit
python3 - "$ROOT" <<'PY'
import json, os, sys
root = sys.argv[1]
d = json.load(open(os.path.join(root, "markets/invest_map/dist/data.json"), encoding="utf-8"))
missing = []
for tk, m in d.get("models", {}).items():
    for key in ("file", "xlsx"):
        if m.get(key) and not os.path.exists(os.path.join(root, m["folder"], "dist", m[key])):
            missing.append(f"{tk}: {m['folder']}/dist/{m[key]}")
if missing:
    print("停止：data.json 指向的模型檔不存在（請先重抽 extract_models.py 並重建）：", *missing, sep="\n  ", file=sys.stderr)
    sys.exit(1)
print("模型檔檢查通過：", len(d.get("models", {})), "家")
PY

# 2) 在 gh-pages 最新版上更新地圖檔
WT=$(mktemp -d)
git -C "$ROOT" fetch -q origin gh-pages
git -C "$ROOT" worktree add -q --detach "$WT" origin/gh-pages
trap 'git -C "$ROOT" worktree remove --force "$WT" >/dev/null 2>&1 || true' EXIT
cp "$ROOT/markets/invest_map/dist/index.html" "$WT/index.html"
rm -rf "$WT/models"
python3 - "$ROOT" "$WT" <<'PY'
import json, os, shutil, sys
root, wt = sys.argv[1], sys.argv[2]
d = json.load(open(os.path.join(root, "markets/invest_map/dist/data.json"), encoding="utf-8"))
for tk, m in d["models"].items():
    dst = os.path.join(wt, "models", m["folder"])
    os.makedirs(dst, exist_ok=True)
    shutil.copy2(os.path.join(root, m["folder"], "dist", m["file"]), dst)
    shutil.copy2(os.path.join(root, m["folder"], "dist", m["xlsx"]), dst)
print("已複製模型：", len(d["models"]), "家")
PY
cd "$WT"
git add -A
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -q -m "發布 AI 投資地圖 $VER（來源 main $SHA；只更新 index.html 與 models/）"
git push -q origin HEAD:gh-pages
echo "已發布 AI 投資地圖 $VER 到 gh-pages（來源 $SHA）：$(git rev-parse --short HEAD)"

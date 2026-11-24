#!/usr/bin/env bash
# 發布 AI 軟體地圖到 gh-pages：在現有 gh-pages 上「追加一個 commit」，只動站上 software_ai/（index.html＋models/<公司>.html），
# 不 force、不動站根與 wattonomics/、siliconomics/、physical_ai/。比照 markets/invest_map/publish_map_only.sh。
# 先檢查：dist/index.html 存在、每個有連結的模型 HTML 來源檔都在，缺任何一份就停止。
# 用法（在 main 的工作目錄）：bash software_ai/map/publish_software_map.sh
# 需要 Andy 在對話中明確說「發布」；本檔只是腳本，建置者不得執行。
set -euo pipefail
ROOT=$(git rev-parse --show-toplevel)
SHA=$(git -C "$ROOT" rev-parse --short HEAD)
M="$ROOT/software_ai/map"
[ -f "$M/dist/index.html" ] || { echo "停止：找不到 dist/index.html，先跑 build_html.py 與 check_map.py" >&2; exit 1; }
python3 "$M/check_map.py" >/dev/null || { echo "停止：check_map.py 未通過" >&2; exit 1; }
VER=$(python3 -c "import json;print(json.load(open('$M/dist/data.json',encoding='utf-8'))['meta']['version'])")

python3 - "$ROOT" <<'PY'
import json, os, sys
root = sys.argv[1]
d = json.load(open(os.path.join(root, "software_ai/map/dist/data.json"), encoding="utf-8"))
x = json.load(open(os.path.join(root, "software_ai/map/data/models_extract.json"), encoding="utf-8"))["models"]
missing = [tk for tk, m in d["models"].items() if m.get("link") and not os.path.exists(os.path.join(root, x[tk]["html"]["file"]))]
if missing:
    print("停止：模型 HTML 來源檔不存在：", *missing, sep="\n  ", file=sys.stderr); sys.exit(1)
print("模型檔檢查通過")
PY

WT=$(mktemp -d)
git -C "$ROOT" fetch -q --depth=1 origin +gh-pages:refs/remotes/origin/gh-pages
git -C "$ROOT" worktree add -q --detach "$WT" origin/gh-pages
trap 'git -C "$ROOT" worktree remove --force "$WT" >/dev/null 2>&1 || true' EXIT
rm -rf "$WT/software_ai"
mkdir -p "$WT/software_ai/models"
cp "$M/dist/index.html" "$WT/software_ai/index.html"
python3 - "$ROOT" "$WT" <<'PY'
import json, os, shutil, sys
root, wt = sys.argv[1], sys.argv[2]
d = json.load(open(os.path.join(root, "software_ai/map/dist/data.json"), encoding="utf-8"))
x = json.load(open(os.path.join(root, "software_ai/map/data/models_extract.json"), encoding="utf-8"))["models"]
n = 0
for tk, m in d["models"].items():
    if m.get("link"):
        shutil.copy2(os.path.join(root, x[tk]["html"]["file"]), os.path.join(wt, "software_ai", m["link"]))   # models/<資料夾>.html
        n += 1
print("已複製模型 HTML：", n, "家")
PY
cd "$WT"
git add -A
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -q -m "發布 AI 軟體地圖 $VER（來源 main $SHA；只更新 software_ai/）" -m "${PUBLISH_TRAILER:-Co-Authored-By: Claude <noreply@anthropic.com>}"
git push -q origin HEAD:gh-pages
echo "已發布 AI 軟體地圖 $VER 到 gh-pages software_ai/（來源 $SHA）：$(git rev-parse --short HEAD)"

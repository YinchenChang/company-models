#!/usr/bin/env bash
# 一鍵建置：Excel → 重算 → 敏感度（寫入 Sensitivity 頁）→ 再重算 → HTML。成品寫到 dist/。
set -euo pipefail
cd "$(dirname "$0")/.."
D=${1:-20261008}; V=${2:-v0_1}
X=out/${D}_MiniMax收支模型_${V}.xlsx; H=out/${D}_MiniMax收支模型_${V}.html
mkdir -p out dist
python3 build_xlsx.py "$X" >/dev/null
python3 scripts/recalc.py "$X"
python3 scripts/sensitivity.py "$X" >/dev/null
python3 scripts/recalc.py "$X"
python3 build_html.py "$X" out/sensitivity.json "$H"
cp "$X" "$H" dist/

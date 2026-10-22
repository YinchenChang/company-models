#!/usr/bin/env bash
# 一鍵建置：Excel → 重算 → 敏感度（寫入 Sensitivity 頁）→ 再重算 → HTML。成品寫到 dist/。
set -euo pipefail
cd "$(dirname "$0")/.."
D=${1:-20261008}; V=${2:-v0_1}   # 用法：build_all.sh [日期] [版本] [dist]
X=out/${D}_MiniMax收支模型_${V}.xlsx; H=out/${D}_MiniMax收支模型_${V}.html
mkdir -p out dist
python3 build_xlsx.py "$X" >/dev/null
python3 scripts/recalc.py "$X"
python3 scripts/sensitivity.py "$X" >/dev/null
python3 scripts/recalc.py "$X"
python3 build_html.py "$X" out/sensitivity.json "$H"
# 驗收：HTML 數字逐一對照 Excel、無外部引用（M3 起）；Checks 合計與錯誤值
python3 scripts/check_html.py "$X" "$H"
python3 scripts/ci_check.py "$X"
# 第三個參數給 dist 才複製成品到 dist/（M1 起預設只留在 out/，不覆蓋既有成品）
if [ "${3:-}" = "dist" ]; then cp "$X" "$H" dist/; fi

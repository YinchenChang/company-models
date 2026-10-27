#!/bin/bash
# Siliconomics 升版後的本機下游批次重建（與 CI ai-semis-build 同順序、同比對）。
# 用法：bash siliconomics/tools/cascade.sh <工作資料夾>
#   依 .github/workflows/ai-semis-build.yml 的 targets 清單，逐一：建檔 → LibreOffice 重算 → 與 dist 逐格比對（容差 1e-9）；
#   有差異時把新檔覆寫到 dist（同檔名），讓後面的下游讀到新上游。輸出每個目標 SAME／CHANGED，最後印 DONE。
#   應跑兩輪：第二輪必須全部 SAME（確認順序依賴已收斂）。HTML 展示層另以 tools/company_html/build_company_html.py 重產。
set -u
cd "$(dirname "$0")/../.."
export PYTHONDONTWRITEBYTECODE=1
B=${1:?需要工作資料夾}; mkdir -p "$B/raw" "$B/out" "$B/diff"
grep -oE '^\s+"[a-z0-9_]+\|python3 [^"]+\|[^"]+"' .github/workflows/ai-semis-build.yml | sed -E 's/^\s+"//; s/"$//' > "$B/targets.txt"
while IFS='|' read -r name cmd dist; do
  [ "$name" = siliconomics ] && continue
  $cmd "$B/raw/$name.xlsx" > /dev/null 2>"$B/$name.err" || { echo "BUILD_FAIL $name"; continue; }
  soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$B/out" "$B/raw/$name.xlsx" >/dev/null 2>&1 || { echo "CALC_FAIL $name"; continue; }
  ref=$(ls $dist/*.xlsx)
  if python3 tools/xlsx_diff.py "$B/out/$name.xlsx" "$ref" --tol 1e-9 --limit 400 > "$B/diff/$name.txt" 2>&1; then echo "SAME $name"
  else echo "CHANGED $name $(grep -m1 '差異' "$B/diff/$name.txt")"; cp "$B/out/$name.xlsx" "$ref"; fi
done < "$B/targets.txt"
echo DONE

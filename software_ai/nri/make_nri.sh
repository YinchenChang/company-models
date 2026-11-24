#!/bin/bash
# NRI 個股模型 v0.2.3：建置 → 敏感度與反解 → 重算 → 全部核對。需要 python3＋openpyxl、LibreOffice。
#   bash software_ai/nri/make_nri.sh            # 沿用已存的敏感度與反解
#   SENS=1 bash software_ai/nri/make_nri.sh     # 改了輸入、公式或 facts 後重跑敏感度與反解
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"; cd "$ROOT"
WD="${WD:-${TMPDIR:-/tmp}/scratch/nri/make}"; mkdir -p "$WD/out" software_ai/nri/dist; export ROWMAP="$WD/rowmap.json"
N=software_ai/nri; S=$N/data/sens_v023_20261124.json; K=$N/data/kreq_v023.json
OUT=$N/dist/20261124_野村總研個股模型_v0.2.3.xlsx
python3 $N/facts_nri.py
python3 $N/build_nri.py "$WD/raw0.xlsx"
if [ "${SENS:-0}" = "1" ]; then
  python3 $N/sens_nri.py "$WD/raw0.xlsx" "$ROWMAP" "$S" 4
  python3 $N/kreq_nri.py "$WD/raw0.xlsx" "$ROWMAP" "$K"
fi
python3 $N/build_nri.py "$WD/raw.xlsx" --sens "$S" --kreq "$K"
LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" >/dev/null 2>&1
cp "$WD/out/raw.xlsx" "$OUT"
python3 $N/gen_html_json.py "$OUT" "$ROWMAP" $N/html.json
python3 tools/segment_check/check_segments.py $N
python3 tools/in_layout/check_in_layout.py "$OUT"
python3 tools/sens_coverage/check_sens_coverage.py "$OUT" --require-keycol
python3 tools/model_governance/check_rating_gate.py "$OUT" --strict
python3 tools/company_html/build_company_html.py $N
python3 tools/company_html/check_company_html.py $N
python3 tools/company_html/check_layout_v13.py "${OUT%.xlsx}.html" --xlsx "$OUT"
python3 tools/model_governance/version_table.py $N --update
python3 tools/model_governance/check_governance.py "$OUT" --html "${OUT%.xlsx}.html" --display $N

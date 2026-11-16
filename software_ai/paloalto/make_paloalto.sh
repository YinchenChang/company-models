#!/bin/bash
# Palo Alto Networks 個股模型 v0.1：建置（含敏感度整本重算）→ LibreOffice 重算 → html.json → HTML → 全部核對。
#   bash software_ai/paloalto/make_paloalto.sh      # 約 1.5 分鐘；需要 python3＋openpyxl、LibreOffice
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"; cd "$ROOT"
P=software_ai/paloalto
python3 $P/build_paloalto.py
OUT="$P/dist/20261116_Palo Alto Networks個股模型_v0.1.xlsx"
python3 $P/gen_html_json.py "$OUT" $P/html.json
python3 tools/segment_check/check_segments.py $P
python3 tools/in_layout/check_in_layout.py "$OUT"
python3 tools/sens_coverage/check_sens_coverage.py "$OUT" --require-keycol
python3 tools/company_html/build_company_html.py $P
python3 tools/company_html/check_company_html.py $P
python3 tools/company_html/check_layout_v13.py "${OUT%.xlsx}.html" --xlsx "$OUT"
python3 tools/model_governance/check_governance.py "$OUT" --html "${OUT%.xlsx}.html"
python3 tools/model_governance/check_rating_gate.py "$OUT" --strict
python3 tools/tokenomics/import_tokenomics.py --check $P/../unit_economics/data/tk_snapshot.json

#!/bin/bash
# Salesforce 個股模型 v0.1：建置 → LibreOffice 重算 → 核對。需要 python3＋openpyxl、LibreOffice。
#   bash software_ai/salesforce/make_v01.sh            # 沿用已提交的敏感度 data/sens_v01_20261116.json
#   SENS=1 bash software_ai/salesforce/make_v01.sh     # 改了輸入或公式後重跑敏感度（約 71 案，約 40 秒）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"; cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD/out" software_ai/salesforce/dist
S=software_ai/salesforce/data/sens_v01_20261116.json
OUT=software_ai/salesforce/dist/20261116_Salesforce個股模型_v0.1.xlsx
python3 software_ai/salesforce/build_salesforce.py "$WD/raw0.xlsx" --rowmap "$WD/rowmap.json"
if [ "${SENS:-0}" = "1" ]; then python3 software_ai/salesforce/sens_salesforce.py "$WD/raw0.xlsx" "$WD/rowmap.json" "$S" 4; fi
python3 software_ai/salesforce/build_salesforce.py "$WD/raw.xlsx" --rowmap "$WD/rowmap.json" --sens "$S"
LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" >/dev/null 2>&1
cp "$WD/out/raw.xlsx" "$OUT"
python3 software_ai/salesforce/gen_html_json.py "$OUT" "$WD/rowmap.json" software_ai/salesforce/html.json
python3 tools/segment_check/check_segments.py software_ai/salesforce
python3 tools/in_layout/check_in_layout.py "$OUT"
python3 tools/sens_coverage/check_sens_coverage.py "$OUT" --require-keycol
python3 tools/company_html/build_company_html.py software_ai/salesforce
python3 tools/company_html/check_company_html.py software_ai/salesforce
python3 tools/company_html/check_layout_v13.py "${OUT%.xlsx}.html" --xlsx "$OUT"
python3 tools/model_governance/check_governance.py "$OUT" --html "${OUT%.xlsx}.html"

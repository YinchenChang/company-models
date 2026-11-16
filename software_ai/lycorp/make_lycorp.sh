#!/bin/bash
# LY 個股模型 v0.1：建置 → LibreOffice 重算 → 核對 → HTML。
#   bash software_ai/lycorp/make_lycorp.sh              # 沿用 data/sens_v010_20261116.json
#   SENS=1 bash software_ai/lycorp/make_lycorp.sh       # 重跑敏感度（約 1 分鐘）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"; cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
VDATE="${VDATE:-20261116}"; VER="${VER:-0.1.0}"
OUTNAME="${VDATE}_LY個股模型_v${VER}.xlsx"
SJ="software_ai/lycorp/data/sens_v${VER//./}_${VDATE}.json"
python3 tools/tokenomics/import_tokenomics.py --check software_ai/lycorp/data/tokenomics_snapshot_v22.0.json
if [ "${SENS:-0}" = "1" ]; then python3 software_ai/lycorp/sens_lycorp.py "$WD/sens" "$SJ" 3; fi
ROWMAP="$WD/rowmap.json" python3 software_ai/lycorp/build_lycorp.py "$WD/raw.xlsx" --sens "$SJ"
mkdir -p "$WD/out" software_ai/lycorp/dist
LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "software_ai/lycorp/dist/$OUTNAME"
python3 software_ai/lycorp/check_lycorp.py "software_ai/lycorp/dist/$OUTNAME"
python3 tools/in_layout/check_in_layout.py "software_ai/lycorp/dist/$OUTNAME"
python3 tools/sens_coverage/check_sens_coverage.py "software_ai/lycorp/dist/$OUTNAME" --require-keycol
python3 tools/segment_check/check_segments.py software_ai/lycorp
python3 software_ai/lycorp/gen_html_json.py "software_ai/lycorp/dist/$OUTNAME" "$WD/rowmap.json" software_ai/lycorp/html.json
python3 tools/company_html/build_company_html.py software_ai/lycorp
python3 tools/company_html/check_company_html.py software_ai/lycorp
python3 tools/company_html/check_layout_v13.py "software_ai/lycorp/dist/${OUTNAME%.xlsx}.html"
python3 tools/model_governance/check_governance.py "software_ai/lycorp/dist/$OUTNAME" --html "software_ai/lycorp/dist/${OUTNAME%.xlsx}.html"

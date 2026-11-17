#!/bin/bash
# Datadog 個股模型 v0.2：事實總帳 → Tokenomics 快照核對 → （敏感度）→ 建置 → LibreOffice 重算 → 核對 → HTML → 治理檢查。
#   bash software_ai/datadog/make_datadog.sh              # 沿用 data/sens_v022_20261117.json
#   SENS=1 bash software_ai/datadog/make_datadog.sh       # 重跑敏感度（約 80 次整本重算）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"; cd "$ROOT"
D=software_ai/datadog
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
VDATE="${VDATE:-20261117}"; VER="${VER:-0.2.2}"
OUTNAME="${VDATE}_Datadog個股模型_v${VER}.xlsx"
SJ="$D/data/sens_v022_20261117.json"
python3 $D/facts_datadog.py
if [ -n "${TOKENOMICS_DIR:-}" ]; then python3 tools/tokenomics/import_tokenomics.py --check $D/data/tokenomics_snapshot_v22.0.json --tokenomics "$TOKENOMICS_DIR"; fi
if [ "${SENS:-0}" = "1" ]; then python3 $D/sens_datadog.py "$WD/sens" "$SJ" "${JOBS:-4}"; fi
ROWMAP="$WD/rowmap.json" python3 $D/build_datadog.py "$WD/raw.xlsx" --sens "$SJ"
mkdir -p "$WD/out" $D/dist
LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "$D/dist/$OUTNAME"
python3 $D/check_datadog.py "$D/dist/$OUTNAME"
python3 tools/in_layout/check_in_layout.py "$D/dist/$OUTNAME"
python3 tools/sens_coverage/check_sens_coverage.py "$D/dist/$OUTNAME" --require-keycol
python3 tools/segment_check/check_segments.py $D
python3 tools/model_governance/check_rating_gate.py "$D/dist/$OUTNAME" --strict
python3 $D/gen_html_json.py "$D/dist/$OUTNAME" "$WD/rowmap.json" $D/html.json
python3 tools/company_html/build_company_html.py $D
python3 tools/company_html/check_company_html.py $D
python3 tools/company_html/check_layout_v13.py "$D/dist/${OUTNAME%.xlsx}.html"
python3 tools/model_governance/version_table.py $D --update
python3 tools/model_governance/check_governance.py "$D/dist/$OUTNAME" --html "$D/dist/${OUTNAME%.xlsx}.html" --display $D

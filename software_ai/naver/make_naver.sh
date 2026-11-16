#!/bin/bash
# Naver 個股模型 v0.2：建置 → LibreOffice 重算 → 核對 → HTML。
#   bash software_ai/naver/make_naver.sh              # 沿用 data/sens_v02_20261117.json
#   SENS=1 bash software_ai/naver/make_naver.sh       # 重跑敏感度與反解（約 10 分鐘）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"; cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
VDATE="${VDATE:-20261117}"; VER="${VER:-0.2}"
N=software_ai/naver
OUTNAME="${VDATE}_Naver個股模型_v${VER}.xlsx"
SJ="$N/data/sens_v${VER//./}_${VDATE}.json"
python3 tools/tokenomics/import_tokenomics.py --check $N/data/tokenomics_snapshot_v22.0.json --tokenomics "${TOKENOMICS:-/home/claude/yinchenchang/tokenomics}"
python3 $N/facts_naver.py
if [ "${SENS:-0}" = "1" ]; then python3 $N/sens_naver.py "$WD/sens" "$SJ" 4; fi
ROWMAP="$WD/rowmap.json" python3 $N/build_naver.py "$WD/raw.xlsx" --sens "$SJ"
mkdir -p "$WD/out" $N/dist
LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "$N/dist/$OUTNAME"
python3 $N/check_naver.py "$N/dist/$OUTNAME"
python3 tools/in_layout/check_in_layout.py "$N/dist/$OUTNAME"
python3 tools/sens_coverage/check_sens_coverage.py "$N/dist/$OUTNAME" --require-keycol
python3 tools/segment_check/check_segments.py $N
python3 tools/model_governance/check_rating_gate.py "$N/dist/$OUTNAME" --strict
python3 $N/gen_html_json.py "$N/dist/$OUTNAME" "$WD/rowmap.json" $N/html.json
python3 tools/company_html/build_company_html.py $N
python3 tools/company_html/check_company_html.py $N
python3 tools/company_html/check_layout_v13.py "$N/dist/${OUTNAME%.xlsx}.html" --xlsx "$N/dist/$OUTNAME"
python3 tools/model_governance/check_governance.py "$N/dist/$OUTNAME" --html "$N/dist/${OUTNAME%.xlsx}.html"

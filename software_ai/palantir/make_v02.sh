#!/bin/bash
# Palantir 個股模型 v0.2：一次跑完
#   事實總帳 → Tokenomics 快照 --check → （反解）→ （敏感度）→ 建置 → LibreOffice 重算 → 核對 → In 頁版面 → 敏感度覆蓋 → 分部檢查 → HTML → 治理、展示層、評等關卡。
#   bash software_ai/palantir/make_v02.sh                  # 沿用已提交的敏感度 data/sens_v02_<VDATE>.json 與反解 data/reverse_solve.json
#   SENS=1 REV=1 bash software_ai/palantir/make_v02.sh     # 重跑敏感度（約 75 次整本重算；JOBS＝平行數，預設 3）與反解（約 36 次）
#   TOKENOMICS_DIR=<Tokenomics clone>                      # 有設定時 --check 快照與 Tokenomics 原檔一致
# 需要：python3＋openpyxl、LibreOffice（soffice）、bs4。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
cd "$ROOT"
P=software_ai/palantir
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
VER="${VER:-0.2.2}"; VDATE="${VDATE:-20261117}"
OUTNAME="${VDATE}_Palantir個股模型_v${VER}.xlsx"
SJ="$P/data/sens_v${VER//./}_${VDATE}.json"
python3 $P/data/facts_palantir.py
if [ -n "${TOKENOMICS_DIR:-}" ]; then python3 tools/tokenomics/import_tokenomics.py --check $P/data/tokenomics_snapshot_v22.0.json --tokenomics "$TOKENOMICS_DIR"; fi
if [ "${REV:-0}" = "1" ]; then python3 $P/reverse_pltr.py "$WD/rev" $P/data/reverse_solve.json; fi
if [ "${SENS:-0}" = "1" ]; then JOBS="${JOBS:-3}" python3 $P/sens_pltr.py "$WD/sens" "$SJ"; fi
ROWMAP="$WD/rowmap.json" python3 $P/build_palantir.py "$WD/raw.xlsx" --sens "$SJ"
mkdir -p "$WD/out" $P/dist
LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "$P/dist/$OUTNAME"
python3 $P/check_pltr.py "$P/dist/$OUTNAME"
python3 tools/in_layout/check_in_layout.py "$P/dist/$OUTNAME"
python3 tools/sens_coverage/check_sens_coverage.py "$P/dist/$OUTNAME" --require-keycol
python3 tools/segment_check/check_segments.py $P
python3 $P/gen_html_json.py "$P/dist/$OUTNAME" "$WD/rowmap.json" $P/html.json
python3 tools/company_html/build_company_html.py $P
python3 tools/company_html/check_company_html.py $P
python3 tools/company_html/check_layout_v13.py "$P/dist/${OUTNAME%.xlsx}.html" --xlsx "$P/dist/$OUTNAME"
python3 tools/model_governance/version_table.py $P --update
python3 tools/model_governance/check_governance.py "$P/dist/$OUTNAME" --html "$P/dist/${OUTNAME%.xlsx}.html" --display $P
python3 tools/model_governance/check_rating_gate.py "$P/dist/$OUTNAME" --strict

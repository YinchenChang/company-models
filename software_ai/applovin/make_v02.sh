#!/bin/bash
# AppLovin 個股模型：一次跑完
#   事實總帳（data/facts_applovin.csv，已提交）→ Tokenomics 快照 --check → （k 反解）→ （敏感度）→ 建置 → LibreOffice 重算 → 核對 → In 頁版面 → 敏感度覆蓋 → 分部檢查 → 收支關卡 → HTML → 治理與展示層一致性。
#   bash software_ai/applovin/make_v02.sh                  # 沿用已提交的反解 data/kreq_v02.json 與敏感度 data/sens_v02_<VDATE>.json
#   KREQ=1 SENS=1 bash software_ai/applovin/make_v02.sh    # 重跑反解與敏感度（JOBS＝平行數，預設 2）
#   TOKENOMICS_DIR=<Tokenomics clone>                      # 有設定時 --check 快照與 Tokenomics 原檔一致
# 需要：python3＋openpyxl＋beautifulsoup4、LibreOffice（soffice）。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
cd "$ROOT"
D=software_ai/applovin
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
VER="${VER:-0.2.3}"; VDATE="${VDATE:-20261117}"
OUTNAME="${VDATE}_AppLovin個股模型_v${VER}.xlsx"
SJ="$D/data/sens_v${VER//./}_${VDATE}.json"; KJ="$D/data/kreq_v${VER//./}.json"
if [ -n "${TOKENOMICS_DIR:-}" ]; then python3 tools/tokenomics/import_tokenomics.py --check $D/data/tokenomics_snapshot.json --tokenomics "$TOKENOMICS_DIR"; fi
if [ "${KREQ:-0}" = "1" ]; then python3 $D/kreq_v02.py "$WD/kreq" "$KJ"; fi
if [ "${SENS:-0}" = "1" ]; then python3 $D/sens_v02.py "$WD/sens" "$SJ"; fi
ROWMAP="$WD/rowmap.json" python3 $D/build_applovin.py "$WD/raw.xlsx" --sens "$SJ" --kreq "$KJ"
mkdir -p "$WD/out" $D/dist
LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "$D/dist/$OUTNAME"
python3 $D/check_v02.py "$D/dist/$OUTNAME"
python3 tools/in_layout/check_in_layout.py "$D/dist/$OUTNAME"
python3 tools/sens_coverage/check_sens_coverage.py "$D/dist/$OUTNAME" --require-keycol
python3 tools/segment_check/check_segments.py $D
python3 tools/model_governance/check_rating_gate.py "$D/dist/$OUTNAME" --strict
python3 $D/gen_html_json.py "$D/dist/$OUTNAME" "$WD/rowmap.json" $D/html.json
python3 tools/company_html/build_company_html.py $D
python3 tools/company_html/check_company_html.py $D
python3 tools/company_html/check_layout_v13.py "$D/dist/${OUTNAME%.xlsx}.html" --xlsx "$D/dist/$OUTNAME"
python3 tools/model_governance/version_table.py $D --update
python3 tools/model_governance/check_governance.py "$D/dist/$OUTNAME" --html "$D/dist/${OUTNAME%.xlsx}.html" --display $D

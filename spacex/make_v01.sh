#!/bin/bash
# SpaceX 個股模型 v0.1：一次跑完 事實總帳 → 敏感度重算（可略）→ 建置 → LibreOffice 重算 → 核對 → In 頁版面 → 敏感度覆蓋 → segment_check → HTML。
#   bash spacex/make_v01.sh            # 沿用已提交的敏感度結果 spacex/data/sens_v01_20261030.json
#   SENS=1 bash spacex/make_v01.sh     # 重新跑敏感度（約 59 次整本重算；JOBS＝平行數，預設 2）
# 需要：python3＋openpyxl、LibreOffice（soffice）。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
OUTNAME="20261030_SpaceX個股模型_v0.1.xlsx"
SJ="spacex/data/sens_v01_20261030.json"
python3 spacex/facts_2026q3.py > /dev/null           # 事實總帳（兩讀核對；不一致即失敗）
if [ "${SENS:-0}" = "1" ]; then
  python3 spacex/sens_v01.py "$WD/sens" "$SJ"
fi
ROWMAP="$WD/rowmap.json" python3 spacex/build_spacex.py "$WD/raw.xlsx" --sens "$SJ"
mkdir -p "$WD/out" spacex/dist
LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "spacex/dist/$OUTNAME"
python3 spacex/check_v01.py "spacex/dist/$OUTNAME"
python3 tools/in_layout/check_in_layout.py "spacex/dist/$OUTNAME"
python3 tools/sens_coverage/check_sens_coverage.py "spacex/dist/$OUTNAME"  # 敏感度覆蓋（原則第 9 節）
python3 tools/segment_check/check_segments.py spacex
python3 spacex/gen_html_json.py "spacex/dist/$OUTNAME" "$WD/rowmap.json" spacex/html.json
python3 tools/company_html/build_company_html.py spacex
python3 tools/company_html/check_company_html.py spacex

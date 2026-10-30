#!/bin/bash
# SpaceX 個股模型 v0.2：一次跑完 事實總帳 → 敏感度重算（可略）→ 建置 → LibreOffice 重算 → 核對 → In 頁版面 → 敏感度覆蓋 → segment_check → HTML。
#   bash spacex/make_v02.sh            # 沿用已提交的敏感度結果 spacex/data/sens_v02_20261030.json
#   SENS=1 bash spacex/make_v02.sh     # 重新跑敏感度（約 95 次整本重算；JOBS＝平行數，預設 2）
# 需要：python3＋openpyxl、LibreOffice（soffice）。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
OUTNAME="20261030_SpaceX個股模型_v0.2.xlsx"
SJ="spacex/data/sens_v02_20261030.json"
python3 spacex/facts_2026q3.py > /dev/null           # 事實總帳（兩讀核對；不一致即失敗）
python3 spacex/facts_v02.py                          # v0.2 增補：424B4、2026Q2 10-Q、法說指引（兩讀；FY2025 分部對帳 0）
if [ "${SENS:-0}" = "1" ]; then
  python3 spacex/sens_v02.py "$WD/sens" "$SJ"
fi
ROWMAP="$WD/rowmap.json" python3 spacex/build_spacex.py "$WD/raw.xlsx" --sens "$SJ"
mkdir -p "$WD/out" spacex/dist
LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "spacex/dist/$OUTNAME"
python3 spacex/check_v02.py "spacex/dist/$OUTNAME"
python3 tools/in_layout/check_in_layout.py "spacex/dist/$OUTNAME"
python3 tools/sens_coverage/check_sens_coverage.py "spacex/dist/$OUTNAME" --require-keycol  # 敏感度覆蓋（原則第 9 節）
python3 tools/segment_check/check_segments.py spacex
python3 spacex/gen_html_json.py "spacex/dist/$OUTNAME" "$WD/rowmap.json" spacex/html.json
python3 tools/company_html/build_company_html.py spacex
python3 tools/company_html/check_company_html.py spacex

#!/bin/bash
# Meta 個股模型 v0.1：一次跑完 事實總帳 → （敏感度）→ 建置 → LibreOffice 重算 → 核對 → In 頁版面 → 敏感度覆蓋 → 分部檢查 → HTML。
#   bash meta/make_v01.sh            # 沿用已提交的敏感度結果 meta/data/sens_v01_20261030.json
#   SENS=1 bash meta/make_v01.sh     # 重跑敏感度（約 80 次整本重算；JOBS＝平行數，預設 2）
# 需要：python3＋openpyxl、LibreOffice（soffice）。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
OUTNAME="20261030_Meta個股模型_v0.1.xlsx"
SJ="meta/data/sens_v01_20261030.json"
python3 meta/facts_2026q3.py                           # 事實總帳（兩讀核對；不一致即失敗；逐字稿不在本機時短句標「未核對」）
if [ "${SENS:-0}" = "1" ]; then python3 meta/sens_v01.py "$WD/sens" "$SJ"; fi
ROWMAP="$WD/rowmap.json" python3 meta/build_meta.py "$WD/raw.xlsx" --sens "$SJ"
mkdir -p "$WD/out" meta/dist
LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "meta/dist/$OUTNAME"
python3 meta/check_v01.py "meta/dist/$OUTNAME"
python3 tools/in_layout/check_in_layout.py "meta/dist/$OUTNAME"
python3 tools/sens_coverage/check_sens_coverage.py "meta/dist/$OUTNAME" --require-keycol  # 敏感度覆蓋（原則第 9 節）
python3 tools/segment_check/check_segments.py meta
python3 meta/gen_html_json.py "meta/dist/$OUTNAME" "$WD/rowmap.json" meta/html.json
python3 tools/company_html/build_company_html.py meta
python3 tools/company_html/check_company_html.py meta

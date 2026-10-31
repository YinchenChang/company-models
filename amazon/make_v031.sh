#!/bin/bash
# Amazon 個股模型 v0.3.1（2026Q3 季度滾動）：一次跑完 事實兩讀核對 →（CALIB=1）C10 重解 →（SENS=1）敏感度與拆解鏈 → 建置 → 重算 → 核對 → HTML。
#   bash amazon/make_v031.sh            # 沿用已提交的 amazon/data/pace_v031.json 與 amazon/data/sens_v031_20261031.json
#   CALIB=1 SENS=1 bash amazon/make_v031.sh   # 重解 C10、重跑敏感度（約 90 次整本重算；JOBS＝平行數）
# 需要：python3＋openpyxl、LibreOffice（soffice）、git（q3cf_v03.py 讀 main 1eec35e9 的 v0.3 建置程式）。任何一步失敗即停止。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
OUTNAME="20261031_Amazon個股模型_v0.3.1.xlsx"
SJ="amazon/data/sens_v031_20261031.json"
python3 amazon/facts_2026q3.py                                   # 2026Q3 事實總帳（兩讀核對；不一致即失敗）
if [ "${CALIB:-0}" = "1" ]; then python3 amazon/calib_pace_v031.py "$WD/calib"; fi
if [ "${SENS:-0}" = "1" ]; then
  python3 amazon/q3cf_v03.py > /dev/null                          # v0.3 對 2026Q3 的模型值（拆解 (a)／(b)）
  python3 amazon/sens_v031.py "$WD/sens" "$SJ"
fi
AMZN_WRITE_HTMLJSON=1 AMZN_ROWMAP="$WD/rowmap.json" python3 amazon/build_amazon.py "$WD/raw.xlsx" --sens "$SJ"
mkdir -p "$WD/out"
LC_ALL=C.UTF-8 soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "amazon/dist/$OUTNAME"
python3 amazon/check_v03.py "amazon/dist/$OUTNAME"
python3 tools/in_layout/check_in_layout.py "amazon/dist/$OUTNAME"
if [ -f mag/check_in_layout.py ]; then python3 mag/check_in_layout.py "amazon/dist/$OUTNAME"; fi
python3 tools/sens_coverage/check_sens_coverage.py "amazon/dist/$OUTNAME" --require-keycol
python3 tools/segment_check/check_segments.py amazon
python3 tools/company_html/build_company_html.py amazon
python3 tools/company_html/check_company_html.py amazon
echo "make_v031：全部通過（$OUTNAME）"

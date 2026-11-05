#!/bin/bash
# Amazon 個股模型 v0.3.1–v0.3.5（v0.3.5＝只改展示層，HTML 套用個股 HTML 版面規格 v1.0；2026Q3 季度滾動；v0.3.2＝Tokenomics 快照 v21.9；v0.3.3＝10-Q 補入；v0.3.4＝Tokenomics 快照 v21.10；腳本沿用 v0.3.1 名稱）：一次跑完 事實兩讀核對 →（CALIB=1）C10 重解 →（SENS=1）敏感度與拆解鏈 → 建置 → 重算 → 核對 → HTML。
#   bash amazon/make_v031.sh            # 沿用已提交的 amazon/data/pace_v031.json 與 amazon/data/sens_v035_20261105.json
#   CALIB=1 SENS=1 bash amazon/make_v031.sh   # 重解 C10、重跑敏感度（約 100 次整本重算；JOBS＝平行數）
#   v0.3.3：10-Q 只改債務、租賃、融資義務、分部折舊，不影響 C10（資本支出與 MW），故只需 SENS=1；拆解鏈為 v0.3.2 → v0.3.3（sens_v031.py chain|S0–S3）
# 需要：python3＋openpyxl、LibreOffice（soffice）。任何一步失敗即停止。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
OUTNAME="20261105_Amazon個股模型_v0.3.5.xlsx"
SJ="amazon/data/sens_v035_20261105.json"
python3 amazon/facts_2026q3.py                                   # 2026Q3 事實總帳（兩讀核對；不一致即失敗）
if [ "${CALIB:-0}" = "1" ]; then python3 amazon/calib_pace_v031.py "$WD/calib"; fi
if [ "${SENS:-0}" = "1" ]; then
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
python3 tools/company_html/relayout.py amazon "amazon/dist/$OUTNAME"   # v0.3.5：新版面（amazon/layout_spec.py）
python3 tools/company_html/build_company_html.py amazon
python3 tools/company_html/check_company_html.py amazon
echo "make_v031：全部通過（$OUTNAME）"

#!/bin/bash
# Alphabet 個股模型 v0.3.1（季度滾動；腳本沿用 v0.3 名稱）：一次跑完 建置 → 敏感度重算（可略）→ 成品重算 → 核對 → HTML。
#   bash alphabet/make_v03.sh            # 沿用已提交的敏感度結果 alphabet/data/sens_v031_20261029.json
#   SENS=1 bash alphabet/make_v03.sh     # 重新跑敏感度（約 40 次整本重算；JOBS＝平行數）
# 需要：python3＋openpyxl、LibreOffice（soffice）；tools/segment_check/check_segments.py 尚未在 main 時以 SEGCHECK 指定路徑。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
OUTNAME="20261029_Alphabet個股模型_v0.3.1.xlsx"
SJ="alphabet/data/sens_v031_20261029.json"
python3 alphabet/facts_2026q3.py                       # 2026Q3 事實總帳（兩讀核對；不一致即失敗）
if [ "${CALIB:-0}" = "1" ]; then python3 alphabet/calib_pace_v031.py "$WD/calib"; fi   # C10 速度重解（改了資本支出相關輸入時）
if [ "${SENS:-0}" = "1" ]; then
  python3 alphabet/sens_v03.py "$WD/sens" "$SJ"
fi
ROWMAP="$WD/rowmap.json" python3 alphabet/build_alphabet.py "$WD/raw.xlsx" --sens "$SJ"
mkdir -p "$WD/out"
LC_ALL=C.UTF-8 soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "alphabet/dist/$OUTNAME"
python3 alphabet/check_v03.py "alphabet/dist/$OUTNAME"
SEGCHECK="${SEGCHECK:-tools/segment_check/check_segments.py}"
if [ -f "$SEGCHECK" ]; then python3 "$SEGCHECK" alphabet; else echo "（略過 segment_check：$SEGCHECK 不存在）"; fi
python3 alphabet/gen_html_json.py "alphabet/dist/$OUTNAME" "$WD/rowmap.json" alphabet/html.json
python3 tools/company_html/build_company_html.py alphabet
python3 tools/company_html/check_company_html.py alphabet

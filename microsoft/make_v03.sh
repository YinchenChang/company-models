#!/bin/bash
# Microsoft 個股模型 v0.3.2：一次跑完 建置 → LibreOffice 重算 → 核對 → segment_check → HTML。
#   bash microsoft/make_v03.sh            # 用已提交的變動拆解 microsoft/data/attrib_v031_20261029.json（Recon 第 5、6 節）
#   SENS=1 bash microsoft/make_v03.sh     # 先重跑 C10 重解與變動拆解（microsoft/attrib_v031.py --solve，約 15 次整本重算）
# 敏感度（Sens 頁）在同一本 xlsx 內以完整引擎區塊計算，不需另跑。需要：python3＋openpyxl、python-docx、LibreOffice（soffice）。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
OUTNAME="20261029_Microsoft個股模型_v0.3.2.xlsx"
python3 microsoft/facts_fy27q1.py
if [ "${SENS:-0}" = "1" ]; then
  python3 microsoft/attrib_v031.py --solve
fi
python3 microsoft/build_microsoft.py "$WD/raw.xlsx"
mkdir -p "$WD/out"
LC_ALL=C.UTF-8 soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "microsoft/dist/$OUTNAME"
python3 microsoft/check_v03.py "microsoft/dist/$OUTNAME"
python3 mag/check_in_layout.py "microsoft/dist/$OUTNAME"
python3 tools/segment_check/check_segments.py microsoft
python3 tools/company_html/build_company_html.py microsoft
python3 tools/company_html/check_company_html.py microsoft

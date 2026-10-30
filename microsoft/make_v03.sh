#!/bin/bash
# Microsoft 個股模型 v0.3.3：一次跑完 建置 → LibreOffice 重算 → 核對 → segment_check → HTML。
#   bash microsoft/make_v03.sh            # 用已提交的變動拆解 microsoft/data/attrib_v031_20261029.json（Recon 第 5、6 節）
#   SENS=1 bash microsoft/make_v03.sh     # 先重跑 C10 重解與變動拆解（microsoft/attrib_v031.py --solve，約 15 次整本重算），
#                                          再重跑 Sens 整本重算案例（microsoft/sens_v033.py，約 60 次整本重算 → data/sens_v033_20261030.json）
# 敏感度（Sens 頁）：原有案例在同一本 xlsx 內以完整引擎區塊計算；v0.3.3 補的案例（原則第 9 節）為整本重算的數值。需要：python3＋openpyxl、python-docx、LibreOffice（soffice）。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
OUTNAME="20261030_Microsoft個股模型_v0.3.3.xlsx"
python3 microsoft/facts_fy27q1.py
if [ "${SENS:-0}" = "1" ]; then
  python3 microsoft/attrib_v031.py --solve
  python3 microsoft/sens_v033.py
fi
python3 microsoft/build_microsoft.py "$WD/raw.xlsx"
mkdir -p "$WD/out"
LC_ALL=C.UTF-8 soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "microsoft/dist/$OUTNAME"
python3 microsoft/check_v03.py "microsoft/dist/$OUTNAME"
python3 mag/check_in_layout.py "microsoft/dist/$OUTNAME"
python3 tools/sens_coverage/check_sens_coverage.py "microsoft/dist/$OUTNAME" --require-keycol   # v0.3.3：敏感度覆蓋（原則第 9 節）
python3 tools/segment_check/check_segments.py microsoft
python3 tools/company_html/build_company_html.py microsoft
python3 tools/company_html/check_company_html.py microsoft

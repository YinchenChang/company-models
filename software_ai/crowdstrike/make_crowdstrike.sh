#!/bin/bash
# CrowdStrike 個股模型 v0.2.2：建置（含敏感度整本重算）→ 反解 → 建置 → HTML → 全部核對。需要 python3＋openpyxl、LibreOffice。
#   bash software_ai/crowdstrike/make_crowdstrike.sh            # 沿用已存的敏感度與反解（--quick）
#   SENS=1 bash software_ai/crowdstrike/make_crowdstrike.sh     # 改了輸入或公式後重跑敏感度與反解
# 注意：建置會把 dist/ 內既有成品（含同版號）搬進 dist/_待刪除/；重跑後須刪掉 _待刪除/ 內多出的同版複本。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"; cd "$ROOT"
N=software_ai/crowdstrike; SCR=/home/claude/scratch/crowdstrike; mkdir -p $SCR
if [ "${SENS:-0}" = "1" ]; then
  python3 $N/build_crowdstrike.py > /dev/null
  python3 $N/build_crowdstrike.py --reverse > /dev/null
  python3 $N/build_crowdstrike.py
else
  python3 $N/build_crowdstrike.py --quick
fi
OUT=$(ls $N/dist/*_CrowdStrike個股模型_v0.2.2.xlsx)
rm -f "$N/dist/_待刪除/$(basename "$OUT")" "$N/dist/_待刪除/$(basename "${OUT%.xlsx}.html")"
python3 $N/gen_html_json.py "$OUT" $SCR/layout.json $N/html.json
python3 tools/company_html/build_company_html.py $N
python3 tools/segment_check/check_segments.py $N
python3 tools/in_layout/check_in_layout.py "$OUT"
python3 tools/sens_coverage/check_sens_coverage.py "$OUT" --require-keycol
python3 tools/model_governance/check_rating_gate.py "$OUT" --strict
python3 tools/company_html/check_company_html.py $N
python3 tools/company_html/check_layout_v13.py "${OUT%.xlsx}.html" --xlsx "$OUT"
python3 tools/model_governance/check_governance.py "$OUT" --html "${OUT%.xlsx}.html" --display $N
! grep -qE 'https?://' "${OUT%.xlsx}.html" || { echo "HTML 含外部網址"; exit 1; }

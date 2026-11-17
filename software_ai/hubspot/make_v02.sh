#!/bin/bash
# HubSpot 個股模型 v0.2.1：建置 →（敏感度與反解）→ LibreOffice 重算 → 核對。需要 python3＋openpyxl、LibreOffice。
#   bash software_ai/hubspot/make_v02.sh            # 沿用已提交的 data/sens_v02.json 與 data/kreq_v02.json
#   SENS=1 bash software_ai/hubspot/make_v02.sh     # 改了輸入或公式後重跑敏感度與反解（約 75 次整本重算，JOBS 控制並行）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"; cd "$ROOT"
F=software_ai/hubspot
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD/out" $F/dist
OUT=$F/dist/20261117_HubSpot個股模型_v0.2.1.xlsx
if [ "${SENS:-0}" = "1" ]; then JOBS="${JOBS:-4}" python3 $F/sens_hubspot.py "$WD/sens" --kreq; fi
ROWMAP="$WD/rowmap.json" python3 $F/build_hubspot.py "$WD/raw.xlsx" --sens $F/data/sens_v02.json --kreq $F/data/kreq_v02.json
LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" >/dev/null 2>&1
cp "$WD/out/raw.xlsx" "$OUT"
python3 - "$OUT" <<'PY'
import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
bad = [(w.title, c.coordinate, c.value) for w in wb for r in w.iter_rows() for c in r if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
print("錯誤值格數", len(bad), bad[:5]); sys.exit(1 if bad else 0)
PY
python3 $F/gen_html_json.py "$OUT" "$WD/rowmap.json" $F/html.json
python3 tools/segment_check/check_segments.py $F
python3 tools/in_layout/check_in_layout.py "$OUT"
python3 tools/sens_coverage/check_sens_coverage.py "$OUT" --require-keycol
python3 tools/company_html/build_company_html.py $F
python3 tools/company_html/check_company_html.py $F
python3 tools/company_html/check_layout_v13.py "${OUT%.xlsx}.html" --xlsx "$OUT"
python3 tools/model_governance/version_table.py $F --update
python3 tools/model_governance/check_governance.py "$OUT" --html "${OUT%.xlsx}.html" --display $F
python3 tools/model_governance/check_rating_gate.py "$OUT" --strict

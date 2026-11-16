#!/bin/bash
# Fujitsu 個股模型 v0.1：建置 → （敏感度與反解）→ LibreOffice 重算 → 核對。需要 python3＋openpyxl、LibreOffice。
#   bash software_ai/fujitsu/make_v01.sh            # 沿用已提交的 data/sens_v01_20261116.json 與 data/rev_solve_v01.json
#   SENS=1 bash software_ai/fujitsu/make_v01.sh     # 改了輸入或公式後重跑敏感度與反解（約 70 次整本重算）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"; cd "$ROOT"
F=software_ai/fujitsu
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD/out" $F/dist
S=$F/data/sens_v01_20261116.json; RV=$F/data/rev_solve_v01.json
OUT=$F/dist/20261116_Fujitsu個股模型_v0.1.xlsx
python3 $F/build_fujitsu.py "$WD/raw0.xlsx" --rowmap "$WD/rowmap.json"
if [ "${SENS:-0}" = "1" ]; then python3 $F/sens_fujitsu.py "$WD/raw0.xlsx" "$WD/rowmap.json" "$S" "$RV" 4; fi
python3 $F/build_fujitsu.py "$WD/raw.xlsx" --rowmap "$WD/rowmap.json" --sens "$S" --rev "$RV"
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

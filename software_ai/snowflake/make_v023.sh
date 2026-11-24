#!/bin/bash
# Snowflake 個股模型 v0.2.3：建置 → LibreOffice 重算 → 核對。需要 python3＋openpyxl、LibreOffice。
#   bash software_ai/snowflake/make_v023.sh            # 沿用已存的敏感度與反解 data/sens_v022_20261117.json、data/kreq_v022.json
#   SENS=1 bash software_ai/snowflake/make_v023.sh     # 改了輸入或公式後重跑敏感度與反解（約 90 案＋14 次反解，約 4 分鐘）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"; cd "$ROOT"
D=software_ai/snowflake
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD/out" $D/dist /home/claude/scratch/snowflake
S=$D/data/sens_v022_20261117.json
K=$D/data/kreq_v022.json
OUT=$D/dist/20261124_Snowflake個股模型_v0.2.3.xlsx
python3 $D/build_snowflake.py "$WD/raw0.xlsx" --rowmap "$WD/rowmap.json"
if [ "${SENS:-0}" = "1" ]; then
  python3 $D/sens_snowflake.py "$WD/raw0.xlsx" "$WD/rowmap.json" "$S" 4
  python3 $D/sens_snowflake.py --kreq "$WD/raw0.xlsx" "$WD/rowmap.json" "$K"
fi
python3 $D/build_snowflake.py "$WD/raw.xlsx" --rowmap "$WD/rowmap.json" --sens "$S"
LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" >/dev/null 2>&1
cp "$WD/out/raw.xlsx" "$OUT"
python3 - "$OUT" <<'PY'
import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True); n = 0
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and (c.value.startswith("#") or "Err:" in c.value): n += 1; print("錯誤值", ws.title, c.coordinate, c.value)
print("錯誤值格數", n); sys.exit(1 if n else 0)
PY
python3 $D/gen_html_json.py "$OUT" "$WD/rowmap.json" $D/html.json
python3 tools/segment_check/check_segments.py $D
python3 tools/in_layout/check_in_layout.py "$OUT"
python3 tools/sens_coverage/check_sens_coverage.py "$OUT" --require-keycol
python3 tools/company_html/build_company_html.py $D
python3 tools/company_html/check_company_html.py $D
python3 tools/company_html/check_layout_v13.py "${OUT%.xlsx}.html" --xlsx "$OUT"
python3 tools/model_governance/check_governance.py "$OUT" --html "${OUT%.xlsx}.html"
python3 tools/model_governance/check_rating_gate.py "$OUT" --strict

#!/bin/bash
# 趨勢科技個股模型 v0.1：事實總帳 → Tokenomics --check → （敏感度與反解）→ 建置 → LibreOffice 重算 → 核對 → HTML → 治理。
#   bash software_ai/trendmicro/make_v01.sh          # 沿用 data/sens_v01.json
#   SENS=1 bash software_ai/trendmicro/make_v01.sh   # 重跑敏感度與反解
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"; cd "$ROOT"; T=software_ai/trendmicro
WD="${WD:-/home/claude/scratch/trendmicro}"; mkdir -p "$WD" $T/dist
VDATE="${VDATE:-20261116}"; OUT="${VDATE}_趨勢科技個股模型_v0.1.xlsx"
python3 $T/facts_trendmicro.py
python3 tools/tokenomics/import_tokenomics.py --check $T/data/tk_snapshot.json --tokenomics "${TOKENOMICS_DIR:-/home/claude/yinchenchang/tokenomics}"
python3 $T/build_trendmicro.py "$WD/probe.xlsx" >/dev/null
if [ "${SENS:-0}" = "1" ]; then JOBS="${JOBS:-4}" python3 $T/sens_trendmicro.py "$WD/sens" $T/data/sens_v01.json; fi
ROWMAP="$WD/rowmap.json" python3 $T/build_trendmicro.py "$WD/raw.xlsx" --sens $T/data/sens_v01.json
rm -rf "$WD/out"; LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" >/dev/null 2>&1
cp "$WD/out/raw.xlsx" "$T/dist/$OUT"
python3 - "$T/dist/$OUT" <<'PY'
import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True); bad = []
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:")): bad.append(f"{ws.title}!{c.coordinate}={c.value}")
print("重算後錯誤格：", len(bad), bad[:5]); sys.exit(1 if bad else 0)
PY
python3 tools/in_layout/check_in_layout.py "$T/dist/$OUT"
python3 tools/sens_coverage/check_sens_coverage.py "$T/dist/$OUT" --require-keycol
python3 tools/segment_check/check_segments.py $T
python3 tools/model_governance/check_rating_gate.py "$T/dist/$OUT" --strict
python3 $T/gen_html_json.py "$T/dist/$OUT" "$WD/rowmap.json" $T/html.json
python3 tools/company_html/build_company_html.py $T
python3 tools/company_html/check_company_html.py $T
python3 tools/company_html/check_layout_v13.py "$T/dist/${OUT%.xlsx}.html" --xlsx "$T/dist/$OUT"
python3 tools/model_governance/check_governance.py "$T/dist/$OUT" --html "$T/dist/${OUT%.xlsx}.html"

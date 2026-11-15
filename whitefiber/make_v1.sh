#!/bin/bash
# WhiteFiber 個股模型 v0.4（CoreWeave v5.0 新雲樣板）：一次跑完 （敏感度、反解、拆解）→ 建置 → LibreOffice 重算 → 核對 → HTML。
#   bash whitefiber/make_v1.sh                    # 沿用已提交的 data/sens_v50.json、reverse_v50.json、decomp_v50.json
#   SENS=1 REV=1 DEC=1 bash whitefiber/make_v1.sh  # 改了公式或輸入時重跑（JOBS＝平行數，預設 2）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
WD="${WD:-$(mktemp -d /tmp/wyfi_mk.XXXX)}"; mkdir -p "$WD"
OUTNAME="20261115_WhiteFiber個股模型_v0.4.2.xlsx"
if [ "${SENS:-0}" = "1" ]; then python3 whitefiber/scripts/sens_v50.py --jobs "${JOBS:-2}"; fi
if [ "${REV:-0}" = "1" ]; then python3 whitefiber/scripts/reverse_v50.py --jobs "${JOBS:-2}"; fi
if [ "${DEC:-0}" = "1" ]; then python3 whitefiber/scripts/decomp_v50.py; fi
WYFI_ROWMAP="$WD/rowmap.json" python3 whitefiber/build_whitefiber.py "$WD/raw.xlsx"
mkdir -p "$WD/out"
soffice -env:UserInstallation=file:///tmp/lo_whitefiber35 --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "whitefiber/dist/$OUTNAME"
python3 - "whitefiber/dist/$OUTNAME" <<'PY'
import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
err = [f"{ws.title}!{c.coordinate}" for ws in wb for row in ws.iter_rows() for c in row
       if isinstance(c.value, str) and c.value.startswith(("#REF", "#DIV", "#VALUE", "#NAME", "#N/A", "#NUM", "Err:"))]
print(f"公式錯誤：{len(err)} 格")
if err: sys.exit(f"失敗：{err[:10]}")
PY
python3 whitefiber/scripts/check_v50.py "whitefiber/dist/$OUTNAME"
python3 tools/in_layout/check_in_layout.py "whitefiber/dist/$OUTNAME"
python3 tools/sens_coverage/check_sens_coverage.py "whitefiber/dist/$OUTNAME" --require-keycol
python3 tools/segment_check/check_segments.py whitefiber
python3 whitefiber/scripts/make_html_json.py "$WD/rowmap.json"
python3 tools/company_html/build_company_html.py whitefiber
python3 tools/company_html/check_company_html.py whitefiber
python3 tools/model_governance/check_governance.py "whitefiber/dist/$OUTNAME" --html "whitefiber/dist/${OUTNAME%.xlsx}.html" --display whitefiber
python3 tools/model_governance/check_rating_gate.py "whitefiber/dist/$OUTNAME" --strict

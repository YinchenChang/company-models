#!/bin/bash
# CoreWeave 個股模型 v5.0（Oracle 收支型樣板）：一次跑完 （敏感度、反解、拆解）→ 建置 → LibreOffice 重算 → 核對 → HTML。
#   bash coreweave/make_v5.sh                    # 沿用已提交的 data/sens_v50.json、reverse_v50.json、decomp_v50.json
#   SENS=1 REV=1 DEC=1 bash coreweave/make_v5.sh  # 改了公式或輸入時重跑（約 2 分鐘；JOBS＝平行數，預設 2）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
OUTNAME="20261113_CoreWeave個股模型_v5.0.1.xlsx"
if [ "${SENS:-0}" = "1" ]; then python3 coreweave/scripts/sens_v50.py --jobs "${JOBS:-2}"; fi
if [ "${REV:-0}" = "1" ]; then python3 coreweave/scripts/reverse_v50.py --jobs "${JOBS:-2}"; fi
if [ "${DEC:-0}" = "1" ]; then python3 coreweave/scripts/decomp_v50.py; fi
CRWV_ROWMAP="$WD/rowmap.json" python3 coreweave/build_coreweave.py "$WD/raw.xlsx"
mkdir -p "$WD/out"
soffice -env:UserInstallation=file://"$WD/lo" --headless --calc --convert-to xlsx --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "coreweave/dist/$OUTNAME"
python3 - "coreweave/dist/$OUTNAME" <<'PY'
import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
err = [f"{ws.title}!{c.coordinate}" for ws in wb for row in ws.iter_rows() for c in row
       if isinstance(c.value, str) and c.value.startswith(("#REF", "#DIV", "#VALUE", "#NAME", "#N/A", "#NUM", "Err:"))]
print(f"公式錯誤：{len(err)} 格")
if err: sys.exit(f"失敗：{err[:10]}")
PY
python3 coreweave/scripts/check_v50.py "coreweave/dist/$OUTNAME"          # 檢查列、Sens／RevDCF 快照、收支恆等式
python3 tools/in_layout/check_in_layout.py "coreweave/dist/$OUTNAME"
python3 tools/sens_coverage/check_sens_coverage.py "coreweave/dist/$OUTNAME" --require-keycol
python3 tools/segment_check/check_segments.py coreweave
python3 coreweave/scripts/make_html_json.py "$WD/rowmap.json"
python3 tools/company_html/build_company_html.py coreweave
python3 tools/company_html/check_company_html.py coreweave
python3 tools/model_governance/check_governance.py "coreweave/dist/$OUTNAME" --html "coreweave/dist/${OUTNAME%.xlsx}.html" --display coreweave
python3 tools/model_governance/check_rating_gate.py "coreweave/dist/$OUTNAME" --strict        # 原則 63：收支關卡不通過不得為買進

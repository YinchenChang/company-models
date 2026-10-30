#!/bin/bash
# Oracle 個股模型 v0.4.2：一次跑完 （敏感度）→ 建置 → LibreOffice 重算 → 核對（In 版面、敏感度覆蓋、segment_check）→ HTML。
#   bash oracle/make_v04.sh            # 沿用已提交的敏感度結果 oracle/data/sens_v04.json
#   SENS=1 bash oracle/make_v04.sh     # 重跑敏感度（約 100 次整本重算；JOBS＝平行數，預設 4）
# v0.4.2 起新增本腳本（原 README「建置」各步驟）；敏感度覆蓋核對依個股模型原則第 9 節（--require-keycol）。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
OUTNAME="20261030_Oracle個股模型_v0.4.2.xlsx"
if [ "${SENS:-0}" = "1" ]; then python3 oracle/scripts/sens_v04.py --jobs "${JOBS:-4}"; fi
ORCL_ROWMAP="$WD/rowmap.json" python3 oracle/build_oracle.py "$WD/raw.xlsx"
mkdir -p "$WD/out"
soffice -env:UserInstallation=file://"$WD/lo" --headless --calc --convert-to xlsx --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "oracle/dist/$OUTNAME"
python3 - "oracle/dist/$OUTNAME" <<'PY'
# 公式錯誤 0、Sens 快照未過期
import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
err = [f"{ws.title}!{c.coordinate}" for ws in wb for row in ws.iter_rows() for c in row
       if isinstance(c.value, str) and c.value.startswith(("#REF", "#DIV", "#VALUE", "#NAME", "#N/A", "#NUM", "Err:"))]
se = wb["Sens"]
st = [r for r in range(1, se.max_row + 1) if str(se.cell(r, 1).value or "").startswith("快照檢查")]
d = se.cell(st[0], 3).value if st else None
print(f"公式錯誤：{len(err)} 格；Sens 快照差 {d}")
if err or not isinstance(d, (int, float)) or abs(d) > 1e-6:
    sys.exit(f"失敗：{err[:10]}；Sens 快照差 {d}（重跑 SENS=1 bash oracle/make_v04.sh）")
PY
python3 tools/in_layout/check_in_layout.py "oracle/dist/$OUTNAME"
python3 tools/sens_coverage/check_sens_coverage.py "oracle/dist/$OUTNAME" --require-keycol   # 原則第 9 節
python3 tools/segment_check/check_segments.py oracle
python3 oracle/scripts/make_html_json.py "$WD/rowmap.json"
python3 tools/company_html/build_company_html.py oracle
python3 tools/company_html/check_company_html.py oracle

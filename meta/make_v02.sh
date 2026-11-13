#!/bin/bash
# Meta 個股模型 v0.2.x（v0.2.0：v0.1.6 驗收採用＋原則第 10 節第 1–55 條；腳本由 make_v01.sh 改名）：一次跑完
#   事實總帳 → Tokenomics 快照 --check → （k 梯度反解）→ （敏感度）→ 建置 → LibreOffice 重算 → 核對 → In 頁版面 → 敏感度覆蓋 → 分部檢查 → HTML → 治理與展示層一致性。
#   bash meta/make_v02.sh                  # 沿用已提交的 k 反解 meta/data/kreq_v020.json 與敏感度 meta/data/sens_v020_<VDATE>.json
#   KREQ=1 SENS=1 bash meta/make_v02.sh    # 重跑 k 梯度與反解（約 20 次整本重算）與敏感度（約 110 次；JOBS＝平行數，預設 2）
#   TOKENOMICS_DIR=<Tokenomics clone>     # 有設定時 --check 快照與 Tokenomics 原檔一致
# 需要：python3＋openpyxl、LibreOffice（soffice）。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
WD="${WD:-$(mktemp -d)}"; mkdir -p "$WD"
VER="${VER:-0.2.0}"; VDATE="${VDATE:-20261111}"
OUTNAME="${VDATE}_Meta個股模型_v${VER}.xlsx"
SJ="meta/data/sens_v${VER//./}_${VDATE}.json"; KJ="meta/data/kreq_v${VER//./}.json"
python3 meta/facts_2026q3.py                           # 事實總帳（兩讀核對；不一致即失敗）
if [ -n "${TOKENOMICS_DIR:-}" ]; then python3 tools/tokenomics/import_tokenomics.py --check meta/data/tokenomics_snapshot_v22.0.json --tokenomics "$TOKENOMICS_DIR"; fi
if [ "${KREQ:-0}" = "1" ]; then python3 meta/kreq_v020.py "$WD/kreq" "$KJ"; fi
if [ "${SENS:-0}" = "1" ]; then python3 meta/sens_v02.py "$WD/sens" "$SJ"; fi
ROWMAP="$WD/rowmap.json" python3 meta/build_meta.py "$WD/raw.xlsx" --sens "$SJ" --kreq "$KJ"
mkdir -p "$WD/out" meta/dist
LC_ALL=C.UTF-8 soffice -env:UserInstallation=file://$WD/lo --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WD/out" "$WD/raw.xlsx" > /dev/null 2>&1
cp "$WD/out/raw.xlsx" "meta/dist/$OUTNAME"
python3 meta/check_v02.py "meta/dist/$OUTNAME"
python3 tools/in_layout/check_in_layout.py "meta/dist/$OUTNAME"
python3 tools/sens_coverage/check_sens_coverage.py "meta/dist/$OUTNAME" --require-keycol  # 敏感度覆蓋（原則第 9 節）
python3 tools/segment_check/check_segments.py meta
python3 meta/gen_html_json.py "meta/dist/$OUTNAME" "$WD/rowmap.json" meta/html.json
python3 tools/company_html/build_company_html.py meta
python3 tools/company_html/check_company_html.py meta
python3 tools/model_governance/version_table.py meta --update
python3 tools/model_governance/check_governance.py "meta/dist/$OUTNAME" --html "meta/dist/${OUTNAME%.xlsx}.html" --display meta   # 原則第 1、2、51 條

#!/usr/bin/env bash
# W2（2026-10-07）：舊方法組合的回歸驗收。把 coreweave/ 與 tools/ 複製到暫存目錄，company.json → methodology.perMw 改為
# capex=legacy、cost=ebitdaPct、revenue=legacy（全為舊方法），在副本跑 scripts/verify.sh --vs-dist：
# 畫面文字、Excel 值與公式都必須與 dist/ v4.5 成品 0 差異（新增列、新增工作表不計），證明新方法沒有改到舊邏輯。
# W3（v4.6 起）：dist/ 已是新版成品，副本的 dist/ 換成 v4.5 成品（自 git 歷史 LEGACY_DIST_REF 取出，預設 9af51ca），
# 並把 vlog.py／tail.js 的版本紀錄截到 v4.5、以 VER=4.5、DATE=2026-09-26 建置（舊方法組合須重現 v4.5 成品）。
# 用法：scripts/verify_legacy.sh（結果寫在 out/verify_legacy.log；暫存副本在 out/legacy_copy/，下次執行時覆蓋）
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DST="$ROOT/out/legacy_copy"
rm -rf "$DST"; mkdir -p "$DST/coreweave"
( cd "$ROOT" && tar --exclude=./out --exclude=./__pycache__ --exclude=./scripts/__pycache__ -cf - . ) | ( cd "$DST/coreweave" && tar -xf - )
cp -r "$ROOT/../tools" "$DST/tools"
mkdir -p "$DST/coreweave/out"
python3 - "$DST/coreweave/company.json" <<'PY'
import json, sys
p = sys.argv[1]; co = json.load(open(p, encoding='utf-8'))
co['methodology']['perMw'].update({'capex': 'legacy', 'cost': 'ebitdaPct', 'revenue': 'legacy'})
open(p, 'w', encoding='utf-8').write(json.dumps(co, indent=1, ensure_ascii=False) + '\n')
PY
# v4.5 成品（git 歷史）放進副本的 dist/；版本紀錄截到 v4.5
REF="${LEGACY_DIST_REF:-9af51ca}"
rm -f "$DST/coreweave/dist/"*
for f in $(git -C "$ROOT/.." -c core.quotepath=false ls-tree --name-only "$REF" coreweave/dist/); do git -C "$ROOT" show "$REF:$f" > "$DST/coreweave/dist/$(basename "$f")"; done
python3 - "$DST/coreweave" <<'PY'
import sys, os, re, json
d = sys.argv[1]
p = os.path.join(d, 'vlog.py'); s = open(p, encoding='utf-8').read()
i = s.index(' ("v4.6"'); j = s.rindex(']'); open(p, 'w', encoding='utf-8').write(s[:i] + s[j:])
p = os.path.join(d, 'tail.js'); s = open(p, encoding='utf-8').read()
k = s.index('var VLOG = ['); e = s.index(']];', k); i = s.index(', ["v4.6"', k)
open(p, 'w', encoding='utf-8').write(s[:i] + s[e + 1:])
PY
export VER=4.5 DATE=2026-09-26
# README 欄位表的「目前數值」隨 company.json 改變：副本內重新產生（只為通過 0b，不影響比對）
(cd "$DST/coreweave" && python3 scripts/fields_doc.py --write >/dev/null)
set +e
"$DST/coreweave/scripts/verify.sh" --vs-dist > "$ROOT/out/verify_legacy.log" 2>&1; rc=$?
set -e
tail -40 "$ROOT/out/verify_legacy.log"
exit $rc

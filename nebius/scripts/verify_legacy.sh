#!/usr/bin/env bash
# v0.2a（2026-10-08，比照 CoreWeave W2／W4 的 verify_legacy.sh）：舊方法的回歸驗收。
# 把 nebius/ 與 tools/ 複製到暫存目錄，company.json → methodology.perMw.revenue 改為 legacy，並把 defaults.m.revMW、defaults.billableOpen
# 還原為 scenarios.revMW 的預設情境值與其校準值（v0.2 的預設輸入），在副本跑 scripts/verify.sh --vs-dist：
# 畫面文字、Excel 值與公式都必須與 v0.2 成品 0 差異（新增工作表「Tokenomics_取數」與新增列不計），證明新方法沒有改到舊邏輯。
# dist/：副本的 dist/ 換成 v0.2 成品（自 git 歷史 LEGACY_DIST_REF 取出，預設 6760147＝v0.2a 開工時的 main）；
# 版本紀錄截到 v0.2、以 VER=0.2、DATE=2026-10-07 建置。
# 用法：scripts/verify_legacy.sh（結果寫在 out/verify_legacy.log；暫存副本在 out/legacy_copy/，下次執行時覆蓋）
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DST="$ROOT/out/legacy_copy"
rm -rf "$DST"; mkdir -p "$DST/nebius"
( cd "$ROOT" && tar --exclude=./out --exclude=./__pycache__ --exclude=./scripts/__pycache__ -cf - . ) | ( cd "$DST/nebius" && tar -xf - )
cp -r "$ROOT/../tools" "$DST/tools"
mkdir -p "$DST/nebius/out"
python3 - "$DST/nebius/company.json" <<'PY'
import json, sys, re
p = sys.argv[1]; s = open(p, encoding='utf-8').read(); co = json.loads(s)
sc = co['defaults']['scenario']; rev = co['scenarios']['revMW'][sc]
bo = int(co['latestQuarter']['revenue'] * 4 / rev[0] + 0.5)
co['methodology']['perMw']['revenue'] = 'legacy'
co['defaults']['m']['revMW'] = rev; co['defaults']['billableOpen'] = bo
open(p, 'w', encoding='utf-8').write(json.dumps(co, indent=1, ensure_ascii=False) + '\n')
PY
REF="${LEGACY_DIST_REF:-6760147}"
rm -f "$DST/nebius/dist/"*
for f in $(git -C "$ROOT/.." -c core.quotepath=false ls-tree --name-only "$REF" nebius/dist/); do git -C "$ROOT" show "$REF:$f" > "$DST/nebius/dist/$(basename "$f")"; done
git -C "$ROOT" show "$REF:nebius/rv_snap.json" > "$DST/nebius/rv_snap.json"
python3 - "$DST/nebius" <<'PY'
import sys, os
d = sys.argv[1]
p = os.path.join(d, 'vlog.py'); s = open(p, encoding='utf-8').read()
if ' ("v0.3"' in s:
    i = s.index(' ("v0.3"'); j = s.rindex(']'); open(p, 'w', encoding='utf-8').write(s[:i] + s[j:])
p = os.path.join(d, 'tail.js'); s = open(p, encoding='utf-8').read()
k = s.index('var VLOG = [')
if ', ["v0.3"' in s[k:]:
    e = s.index(']];', k); i = s.index(', ["v0.3"', k)
    open(p, 'w', encoding='utf-8').write(s[:i] + s[e + 1:])
PY
export VER=0.2 DATE=2026-10-07
# README 欄位表的「目前數值」隨 company.json 改變：副本內重新產生（只為通過 0b，不影響比對）
(cd "$DST/nebius" && python3 scripts/fields_doc.py --write >/dev/null)
set +e
LEGACY= "$DST/nebius/scripts/verify.sh" --vs-dist > "$ROOT/out/verify_legacy.log" 2>&1; rc=$?
set -e
tail -30 "$ROOT/out/verify_legacy.log"
exit $rc

#!/usr/bin/env bash
# 舊方法組合的回歸驗收：把 coreweave/ 與 tools/ 複製到暫存目錄，改 company.json → methodology.perMw，在副本跑 scripts/verify.sh --vs-dist：
# 畫面文字、Excel 值與公式都必須與前一版成品 0 差異（新增列、新增工作表不計），證明新方法沒有改到舊邏輯。
# 基準（LEGACY_BASE）：
#   v4.6（預設；W4 起）：只把收入改回 revenue=legacy（capex=tokenomics、cost=bottomUp 維持 v4.6；W5 起另移除 companyAdjust），Tokenomics 快照換回 v4.6 用的 v5.26
#        （自 git 歷史 LEGACY_DIST_REF 取 company.json 的 tokenomics 區段與快照檔；v5.27 原 25 名數值與 v5.26 相同），
#        副本 dist/ 換成 v4.6 成品（git 歷史 LEGACY_DIST_REF，預設 f373885），版本紀錄截到 v4.6，以 VER=4.6、DATE=2026-10-08 建置。
#   v4.5（W2／W3 原用法）：capex=legacy、cost=ebitdaPct、revenue=legacy，對 v4.5 成品（LEGACY_DIST_REF 預設 9af51ca），版本紀錄截到 v4.5。
# 用法：scripts/verify_legacy.sh（結果寫在 out/verify_legacy.log；暫存副本在 out/legacy_copy/，下次執行時覆蓋）
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
BASE="${LEGACY_BASE:-v4.6}"
case "$BASE" in
  v4.6) REF="${LEGACY_DIST_REF:-f373885}"; CUT="v4.7"; export VER=4.6 DATE=2026-10-08 ;;
  v4.5) REF="${LEGACY_DIST_REF:-9af51ca}"; CUT="v4.6"; export VER=4.5 DATE=2026-09-26 ;;
  *) echo "LEGACY_BASE 只接受 v4.6 或 v4.5"; exit 2 ;;
esac
DST="$ROOT/out/legacy_copy"
rm -rf "$DST"; mkdir -p "$DST/coreweave"
( cd "$ROOT" && tar --exclude=./out --exclude=./__pycache__ --exclude=./scripts/__pycache__ -cf - . ) | ( cd "$DST/coreweave" && tar -xf - )
cp -r "$ROOT/../tools" "$DST/tools"
mkdir -p "$DST/coreweave/out"
git -C "$ROOT" show "$REF:coreweave/company.json" > "$DST/ref_company.json"
python3 - "$DST/coreweave/company.json" "$BASE" "$DST/ref_company.json" <<'PY'
import json, sys
p, base, refp = sys.argv[1:4]; co = json.load(open(p, encoding='utf-8')); ref = json.load(open(refp, encoding='utf-8'))
if base == 'v4.6':
    co['methodology']['perMw']['revenue'] = 'legacy'
    co.pop('companyAdjust', None)                 # W5：公司調整只在新方法（tkAnchor）使用；舊方法回歸不含
    co['tokenomics'] = ref['tokenomics']          # v4.6 用的快照（v5.26）
else:
    co['methodology']['perMw'].update({'capex': 'legacy', 'cost': 'ebitdaPct', 'revenue': 'legacy'})
open(p, 'w', encoding='utf-8').write(json.dumps(co, indent=1, ensure_ascii=False) + '\n')
print('TKSNAP', co['tokenomics']['snapshotFile'])
PY
SNAP="$(python3 -c "import json,sys; print(json.load(open(sys.argv[1], encoding='utf-8'))['tokenomics']['snapshotFile'])" "$DST/coreweave/company.json")"
[[ -f "$DST/coreweave/$SNAP" ]] || git -C "$ROOT" show "$REF:coreweave/$SNAP" > "$DST/coreweave/$SNAP"
# 前一版成品（git 歷史）放進副本的 dist/；版本紀錄截到前一版
rm -f "$DST/coreweave/dist/"*
for f in $(git -C "$ROOT/.." -c core.quotepath=false ls-tree --name-only "$REF" coreweave/dist/); do git -C "$ROOT" show "$REF:$f" > "$DST/coreweave/dist/$(basename "$f")"; done
python3 - "$DST/coreweave" "$CUT" <<'PY'
import sys, os
d, cut = sys.argv[1:3]
p = os.path.join(d, 'vlog.py'); s = open(p, encoding='utf-8').read()
if f' ("{cut}"' in s:
    i = s.index(f' ("{cut}"'); j = s.rindex(']'); open(p, 'w', encoding='utf-8').write(s[:i] + s[j:])
p = os.path.join(d, 'tail.js'); s = open(p, encoding='utf-8').read()
k = s.index('var VLOG = ['); e = s.index(']];', k)
if f', ["{cut}"' in s[k:e + 3]:
    i = s.index(f', ["{cut}"', k); open(p, 'w', encoding='utf-8').write(s[:i] + s[e + 1:])
PY
# README 欄位表的「目前數值」隨 company.json 改變：副本內重新產生（只為通過 0b，不影響比對）
(cd "$DST/coreweave" && python3 scripts/fields_doc.py --write >/dev/null)
set +e
"$DST/coreweave/scripts/verify.sh" --vs-dist > "$ROOT/out/verify_legacy.log" 2>&1; rc=$?
set -e
tail -40 "$ROOT/out/verify_legacy.log"
exit $rc

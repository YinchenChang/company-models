#!/usr/bin/env bash
# WhiteFiber v0.2（自 Oracle v0.3 verify_legacy.sh；比照 CoreWeave W2／W4）：舊方法的回歸驗收。
# 把 whitefiber/ 與 tools/ 複製到暫存目錄，company.json → methodology.perMw.revenue 改為 legacy，並把 defaults.m.revMW、defaults.billableOpen
# 還原為 scenarios.revMW 的預設情境值與其校準值（v0.1 的預設輸入），在副本跑 scripts/verify.sh --vs-dist：
# 畫面文字、Excel 值與公式都必須與 v0.1 成品 0 差異（新增工作表「Tokenomics_取數」與新增列不計），證明新方法沒有改到舊邏輯。
# dist/：副本的 dist/ 換成 v0.1 成品（自 git 歷史 LEGACY_DIST_REF 取出，預設 1079d5c＝v0.2 開工時的 main（v0.1 已合併））；
# 版本紀錄截到 v0.1、以 VER=0.1、DATE=2026-10-08 建置。
# 用法：scripts/verify_legacy.sh（結果寫在 out/verify_legacy.log；暫存副本在 out/legacy_copy/，下次執行時覆蓋）
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DST="$ROOT/out/legacy_copy"
rm -rf "$DST"; mkdir -p "$DST/whitefiber"
( cd "$ROOT" && tar --exclude=./out --exclude=./__pycache__ --exclude=./scripts/__pycache__ -cf - . ) | ( cd "$DST/whitefiber" && tar -xf - )
cp -r "$ROOT/../tools" "$DST/tools"
mkdir -p "$DST/whitefiber/out"
git -C "$ROOT" show "${LEGACY_DIST_REF:-1079d5c}:whitefiber/company.json" > "$DST/ref_company.json"
python3 - "$DST/whitefiber/company.json" "$DST/ref_company.json" <<'PY'
import json, sys, re
p = sys.argv[1]; s = open(p, encoding='utf-8').read(); co = json.loads(s)
ref = json.load(open(sys.argv[2], encoding='utf-8'))
co['texts'] = ref['texts']; co['texts']['scenarioTipTk'] = ref['texts']['scenarioTip']; co['varianceReasons'] = ref['varianceReasons']; co['meta']['updateDate'] = ref['meta']['updateDate']  # 畫面文字（副標、差異原因、更新日）還原為 v0.1
sc = co['defaults']['scenario']; rev = co['scenarios']['revMW'][sc]
dp = co['scenarios']['billableRatio'].get('roundDp', 0)
rq = lambda x: int(x * 10 ** dp + 0.5) / 10 ** dp if dp else int(x + 0.5)  # WhiteFiber：可計費 MW 取 roundDp 位
bo = rq(co['scenarios']['billableRatio']['openAnnualRevenue'] / rev[0])  # converge 模式的校準公式
co['methodology']['perMw']['revenue'] = 'legacy'
if 'capex' in co['methodology']['perMw']: co['methodology']['perMw']['capex'] = 'legacy'  # 步驟 4 起：每 MW 建置成本也還原為 v0.1 的寫死值（非 tokenomics 即用 scenarios.capexTemplate.costMW）
co['methodology']['perMw']['deprLife'] = 'legacy'; co['methodology']['perMw']['gpuConv'] = 'legacy'  # v0.1：汰換年限為公司會計年限（gpuLife 5），不核對 GPU 數換算
co['defaults']['gpuLife'] = ref['defaults']['gpuLife']; co['methodology']['checks']['capexPerMwBand'] = ref['methodology']['checks']['capexPerMwBand']
co['defaults']['ebitdarAdj'] = ref['defaults']['ebitdarAdj']  # EBITDAR 率校準值依收入而定：legacy 還原為 v0.1 的校準值
co['defaults']['m']['revMW'] = rev; co['defaults']['billableOpen'] = bo
import os; sys.path.insert(0, os.path.dirname(p)); import calendar_q
a = []
for i, L in enumerate(calendar_q.derive(co)['periodYears']):
    a.append(min(co['scenarios']['mwPath']['contracted'][sc][i], co['scenarios']['mwPath']['connectedStart'] if i == 0 else a[i - 1] + co['scenarios']['mwPath']['pace'][sc] * L))
co['defaults']['m']['billable'] = [rq(bo + (x - bo) * b) for x, b in zip(a, co['scenarios']['billableRatio']['ratio'])]
open(p, 'w', encoding='utf-8').write(json.dumps(co, indent=1, ensure_ascii=False) + '\n')
PY
REF="${LEGACY_DIST_REF:-1079d5c}"
rm -f "$DST/whitefiber/dist/"*
for f in $(git -C "$ROOT/.." -c core.quotepath=false ls-tree --name-only "$REF" whitefiber/dist/); do git -C "$ROOT" show "$REF:$f" > "$DST/whitefiber/dist/$(basename "$f")"; done
git -C "$ROOT" show "$REF:whitefiber/rv_snap.json" > "$DST/whitefiber/rv_snap.json"
python3 - "$DST/whitefiber" <<'PY'
import sys, os
d = sys.argv[1]
p = os.path.join(d, 'vlog.py'); s = open(p, encoding='utf-8').read()
if ' ("v0.2"' in s:
    i = s.index(' ("v0.2"'); j = s.rindex(']'); open(p, 'w', encoding='utf-8').write(s[:i] + s[j:])
p = os.path.join(d, 'tail.js'); s = open(p, encoding='utf-8').read()
k = s.index('var VLOG = [')
if ', ["v0.2"' in s[k:]:
    e = s.index(']];', k); i = s.index(', ["v0.2"', k)
    open(p, 'w', encoding='utf-8').write(s[:i] + s[e + 1:])
PY
export VER=0.1 DATE=2026-10-08
# README 欄位表的「目前數值」隨 company.json 改變：副本內重新產生（只為通過 0b，不影響比對）
(cd "$DST/whitefiber" && python3 scripts/fields_doc.py --write >/dev/null)
set +e
LEGACY= EXPECT=scripts/expect/legacy_vs_v0_1.txt "$DST/whitefiber/scripts/verify.sh" --vs-dist > "$ROOT/out/verify_legacy.log" 2>&1; rc=$?
set -e
tail -30 "$ROOT/out/verify_legacy.log"
exit $rc

#!/usr/bin/env bash
# 一次跑完建置與核對：季度層檢查 → Tokenomics 快照 --check（W1）→ 建 HTML → 建 Excel → 重算 → 反向 DCF 與每 MW 敏感度快照（W2）→ fix_outline → fix_datatable → verify_ooxml → 快照值＝Excel 分頁值（W1）→ 離線開啟檢查 → 三情境 xlx＋cmp31 → FY27 錨定 cmp31 → 季度層測試（v4.4）→ 期間滾動測試 → 目標價變動拆解工具測試。
# 任何一項失敗即以非零代碼結束。產物放在 repo 根目錄的 out/（不納入版控）。
# 用法：scripts/verify.sh [--vs-dist]
#   --vs-dist  另與 dist/ 內現行成品比對：crawl.py 畫面文字與 xl_diff.py --by-label（值與公式）皆須 0 差異（純結構修改的驗收）。
#              新增列、常數改為引用輸入格、文字改為活公式會另列清單，不計為差異（規則見 xl_diff.py 開頭）。
# 版本號與更新日：預設取 vlog.py 最後一列與 dist/ 成品檔名的日期；可用環境變數 VER、DATE 覆寫。
# EXPECT＝預期差異清單（v4.5 起）：傳給 xl_diff.py --expect（升版時對舊版成品做 --vs-dist，列出預期變動的格子，其餘須 0 差異）；W3 起有 EXPECT 時 crawl 畫面文字差異只要求無頁面錯誤、無缺頁（明細存 out/crawl_upgrade_diff.txt）。
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUT="$ROOT/out"
mkdir -p "$OUT"
cd "$ROOT"
export LC_ALL=C.UTF-8 LANG=C.UTF-8

VS_DIST=0
[[ "${1:-}" == "--vs-dist" ]] && VS_DIST=1

VER="${VER:-$(python3 -c "import vlog; print(vlog.VLOG[-1][0].lstrip('v'))")}"
DIST_HTML="$(ls dist/*_v*.html | head -1)"
DIST_XLSX="$(ls dist/*_v*.xlsx | head -1)"
DATE_TAG="$(basename "$DIST_HTML" | cut -d_ -f1)"           # 20260924
DATE="${DATE:-${DATE_TAG:0:4}-${DATE_TAG:4:2}-${DATE_TAG:6:2}}"  # 2026-09-24
NAME="${DATE//-/}_CoreWeave收支模型_v${VER//./_}"
HTML="$OUT/$NAME.html"
XLSX="$OUT/$NAME.xlsx"

PASS=(); FAIL=()
step() { echo; echo "=== $1"; }
ok()   { PASS+=("$1"); echo "[PASS] $1"; }
bad()  { FAIL+=("$1"); echo "[FAIL] $1"; }
finish() {
  echo; echo "================ 結果 ================"
  for s in "${PASS[@]}"; do echo "PASS  $s"; done
  for s in "${FAIL[@]:-}"; do [[ -n "$s" ]] && echo "FAIL  $s"; done
  if (( ${#FAIL[@]} )); then echo "verify.sh：失敗 ${#FAIL[@]} 項"; exit 1; fi
  echo "verify.sh：全部通過（${#PASS[@]} 項）"; exit 0
}

# cmp31 輸出中 OK 行數、XX／MISSING 行數；有任何不一致即失敗
run_cmp() {  # $1=標籤 $2..=cmp31 參數
  local label="$1"; shift
  local log="$OUT/cmp31_${label}.log"
  if ! (cd "$OUT" && node "$ROOT/cmp31.js" "$@") > "$log" 2>&1; then bad "cmp31 $label（執行錯誤，見 $log）"; return; fi
  local nok nbad
  nok=$(grep -c '^OK ' "$log" || true)
  nbad=$(grep -cE '^(XX |MISSING)|\| MISSING XL KEY' "$log" || true)
  if (( nbad == 0 && nok > 0 )); then ok "cmp31 $label：$nok 項全部 OK"
  else bad "cmp31 $label：OK $nok、不一致 $nbad（見 $log）"; grep -E '^(XX |MISSING)|MISSING XL KEY' "$log" | head -10; fi
}

step "0. 季度層檢查（季度加總＝年度、指引一致性、超過門檻的差距都有原因；v4.4）"
if node scripts/check_quarterly.js "$ROOT"; then ok "check_quarterly：季度加總＝年度、差異原因齊全"; else bad "check_quarterly"; finish; fi

step "0b. README 欄位說明與 company.json 一致（scripts/fields_doc.py --check；5a）"
if python3 scripts/fields_doc.py --check; then ok "README 欄位說明與 company.json 一致"; else bad "README 欄位說明（執行 python3 scripts/fields_doc.py --write）"; fi

step "0c. Tokenomics 快照可重現（tools/tokenomics/import_tokenomics.py --check；W1）"
TK_SNAP="$(python3 -c "import json; print(json.load(open('company.json', encoding='utf-8')).get('tokenomics', {}).get('snapshotFile', ''))")"
if [[ -z "$TK_SNAP" ]]; then ok "Tokenomics 快照：company.json 無 tokenomics 區段（不適用）"
else
  set +e; r=$(python3 "$ROOT/../tools/tokenomics/import_tokenomics.py" --check "$TK_SNAP" ${TOKENOMICS_DIR:+--tokenomics "$TOKENOMICS_DIR"} 2>&1); rc=$?; set -e
  echo "$r" | grep -v '^警告：Tokenomics 尚無名稱' || true
  if (( rc != 0 )); then bad "Tokenomics 快照 --check"
  elif grep -q '略過 --check' <<<"$r"; then ok "Tokenomics 快照 --check：找不到 Tokenomics clone，略過（警告）"
  else ok "Tokenomics 快照 --check：$(grep -- '--check 通過' <<<"$r" | sed 's/^--check 通過：//')"; fi
fi

step "1. 建 HTML（v$VER · $DATE）"
if python3 build_html_portable.py "$VER" "$HTML" "$DATE" docs/template_v3_3.html; then ok "建 HTML"; else bad "建 HTML"; finish; fi

step "2. 建 Excel"
if python3 build_xlsx.py "$XLSX"; then ok "建 Excel"; else bad "建 Excel"; finish; fi

step "3. 重算（LibreOffice headless）"
if python3 scripts/recalc.py "$XLSX" 120; then ok "重算：0 公式錯誤"; else bad "重算"; finish; fi

step "3b. 反向 DCF：以 Excel 求解（scripts/rv_solve.py；v4.5 取代 JS 快照）"
set +e; python3 scripts/rv_solve.py "$XLSX"; rc=$?; set -e
if (( rc == 3 )); then  # 反解結果有變動：rv_snap.json 已更新，重建 Excel 後再解一次須無變動
  echo "rv_snap.json 已更新：重建 Excel"
  python3 build_xlsx.py "$XLSX" && python3 scripts/recalc.py "$XLSX" 120 >/dev/null || { bad "重建 Excel（反向 DCF 更新後）"; finish; }
  set +e; python3 scripts/rv_solve.py "$XLSX"; rc=$?; set -e
fi
if (( rc == 0 )); then ok "反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致"; else bad "反向 DCF 求解（代碼 $rc）"; finish; fi

step "3c. 每 MW 敏感度快照：以 Excel 求值（scripts/permw_sens.py；W2）"
set +e; python3 scripts/permw_sens.py "$XLSX"; rc=$?; set -e
if (( rc == 3 )); then  # 快照有變動：permw_sens.json 已更新，重建 Excel 後再求一次須無變動
  echo "permw_sens.json 已更新：重建 Excel"
  python3 build_xlsx.py "$XLSX" && python3 scripts/recalc.py "$XLSX" 120 >/dev/null || { bad "重建 Excel（敏感度快照更新後）"; finish; }
  set +e; python3 scripts/permw_sens.py "$XLSX"; rc=$?; set -e
fi
if (( rc == 0 )); then ok "每 MW 敏感度快照：Excel 求值，permw_sens.json 與 Excel 一致"; else bad "每 MW 敏感度快照（代碼 $rc）"; finish; fi

step "4. fix_outline"
if python3 fix_outline.py "$XLSX"; then ok "fix_outline"; else bad "fix_outline"; finish; fi

step "4b. fix_datatable（模擬運算表還原為 Excel 格式）"
if python3 fix_datatable.py "$XLSX"; then ok "fix_datatable"; else bad "fix_datatable"; finish; fi

step "5. verify_ooxml"
if python3 verify_ooxml.py "$XLSX"; then ok "verify_ooxml：OOXML OK"; else bad "verify_ooxml"; fi

step "5d. 快照值＝Excel「Tokenomics_取數」分頁值（scripts/check_tokenomics_tab.py；W1）"
if [[ -n "$TK_SNAP" ]]; then
  if r=$(python3 scripts/check_tokenomics_tab.py "$XLSX"); then echo "$r"; ok "$r"; else echo "$r"; bad "快照值＝Excel 分頁值"; fi
else ok "Tokenomics_取數：不適用（無 tokenomics 區段）"; fi

step "5c. 離線開啟檢查（已決定事項 11：單一檔案、無網路請求、console 無錯誤、關鍵數字正常顯示）"
if python3 scripts/check_offline.py "$HTML" "$XLSX"; then ok "離線開啟：新建 HTML"; else bad "離線開啟：新建 HTML"; fi
if python3 scripts/check_offline.py "$DIST_HTML" "$DIST_XLSX"; then ok "離線開啟：dist/ 成品"; else bad "離線開啟：dist/ 成品"; fi

step "6. 三情境 xlx＋cmp31（預設錨定）"
declare -A SC=([1]=low [2]=base [3]=high)
for s in 1 2 3; do
  if python3 xlx.py "$XLSX" "$s" "$OUT/xl17_$s.json"; then run_cmp "${SC[$s]}" "${SC[$s]}"
  else bad "xlx 情境 $s"; fi
done

step "7. FY27 錨定 cmp31（基準）"
if python3 xlx.py "$XLSX" 2 "$OUT/xl17_2_a1.json" 1; then run_cmp "base_FY27" base 1
else bad "xlx FY27 錨定"; fi

step "8. 季度層測試（暫存副本：假設 Q3 實際數、可移植性；v4.4）"
if python3 scripts/test_quarterly.py > "$OUT/test_quarterly.log" 2>&1; then ok "test_quarterly：假設實際數與可移植性測試通過"; tail -3 "$OUT/test_quarterly.log"
else bad "test_quarterly（見 $OUT/test_quarterly.log）"; tail -20 "$OUT/test_quarterly.log"; fi

step "8b. 期間滾動測試（暫存副本：日曆推算 6 種情況、滾動後第一屏無過期日期與期間字樣；v4.5）"
if python3 scripts/test_rolling.py > "$OUT/test_rolling.log" 2>&1; then ok "test_rolling：日曆推算與滾動後第一屏"; tail -3 "$OUT/test_rolling.log"
else bad "test_rolling（見 $OUT/test_rolling.log）"; tail -30 "$OUT/test_rolling.log"; fi

step "8c. 目標價變動拆解工具測試（scripts/attrib.py：(a)＝(1＋WACC)^(月數÷12)，WACC 讀 Excel、月數讀 calendar_q；四項相加＝總變動）"
if python3 scripts/test_attrib.py "$XLSX" > "$OUT/test_attrib.log" 2>&1; then ok "test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）"; tail -1 "$OUT/test_attrib.log"
else bad "test_attrib（見 $OUT/test_attrib.log）"; tail -20 "$OUT/test_attrib.log"; fi

step "8d. 每 MW 改寫測試（暫存副本：MW 設施口徑換算、GPU 小時價格路線；W2）"
if python3 scripts/test_permw.py > "$OUT/test_permw.log" 2>&1; then ok "test_permw：設施口徑換算與 GPU 小時價格路線"; tail -1 "$OUT/test_permw.log"
else bad "test_permw（見 $OUT/test_permw.log）"; tail -20 "$OUT/test_permw.log"; fi

if (( VS_DIST )); then
  step "9. 與 dist/ 成品比對"
  python3 crawl.py "$DIST_HTML" "$OUT/crawl_dist.json"
  python3 crawl.py "$HTML" "$OUT/crawl_new.json"
  if python3 - "$OUT/crawl_dist.json" "$OUT/crawl_new.json" <<'EOF'
import sys, json, difflib
a, b = (json.load(open(p)) for p in sys.argv[1:3])
pa, pb = a['pages'], b['pages']
n = sum(len(v) for v in pa.values())
diff = [k for k in sorted(set(pa) | set(pb)) if pa.get(k) != pb.get(k)]
print(f'畫面：dist {len(pa)} 個、新版 {len(pb)} 個頁面／分頁，約 {n:,} 字；差異 {len(diff)} 個；頁面錯誤 dist {len(a["errors"])}、新版 {len(b["errors"])}')
for k in diff[:5]:
    print('  ', k); print('\n'.join(list(difflib.unified_diff((pa.get(k) or '').splitlines(), (pb.get(k) or '').splitlines(), lineterm='', n=0))[:12]))
sys.exit(1 if diff or b['errors'] or not pa else 0)
EOF
  then ok "crawl 畫面文字 0 差異"
  elif [[ -n "${EXPECT:-}" ]]; then  # W3：升版驗收（有 EXPECT 清單）時，數字改變使畫面文字必然不同：只要求頁面無錯誤，差異頁另存 out/crawl_upgrade_diff.txt 供人工檢視
    if python3 - "$OUT/crawl_dist.json" "$OUT/crawl_new.json" "$OUT/crawl_upgrade_diff.txt" <<'EOF2'
import sys, json, difflib
a, b = (json.load(open(p)) for p in sys.argv[1:3])
pa, pb = a['pages'], b['pages']
diff = [k for k in sorted(set(pa) | set(pb)) if pa.get(k) != pb.get(k)]
with open(sys.argv[3], 'w') as f:
    for k in diff:
        f.write(f'== {k}\n' + '\n'.join(difflib.unified_diff((pa.get(k) or '').splitlines(), (pb.get(k) or '').splitlines(), lineterm='', n=0)) + '\n')
print(f'升版驗收：畫面文字差異 {len(diff)} 頁（新增頁 {len(set(pb) - set(pa))}、缺少頁 {len(set(pa) - set(pb))}），明細 {sys.argv[3]}')
sys.exit(1 if b['errors'] or set(pa) - set(pb) else 0)
EOF2
    then ok "crawl 畫面文字：升版預期差異（頁面無錯誤、無缺頁；明細見 out/crawl_upgrade_diff.txt）"; else bad "crawl 畫面文字（頁面錯誤或缺頁）"; fi
  else bad "crawl 畫面文字"; fi
  # v4.2：--by-label——列數或 A 欄不同的工作表以「區段＋列名稱」配對（新增列另列清單；列名稱重複即報錯），其餘逐格
  r=$(python3 xl_diff.py "$DIST_XLSX" "$XLSX" --values --by-label ${EXPECT:+--expect=$EXPECT} || true); echo "$r"
  if [[ "$(head -1 <<<"$r")" == "0 differences" ]]; then ok "xl_diff --values --by-label 0 差異"; else bad "xl_diff --values --by-label"; fi
  r=$(python3 xl_diff.py "$DIST_XLSX" "$XLSX" --by-label ${EXPECT:+--expect=$EXPECT} || true); echo "$r"
  if [[ "$(head -1 <<<"$r")" == "0 differences" ]]; then ok "xl_diff 公式 --by-label 0 差異"; else bad "xl_diff 公式 --by-label"; fi
fi

finish

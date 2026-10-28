#!/usr/bin/env bash
# 把 Wattonomics API 工具的輸出（out/ 與 requests/）commit 到結果分支（不寫 main）。
# 用法：wattonomics_commit.sh <工具資料夾> <結果分支> <訊息>
#   - 結果分支＝目前分支時：直接 commit＋push（mag/extract 慣例）。
#   - 否則：把工具資料夾的 out/、requests/ 複製到結果分支（不存在則從目前 HEAD 開）後 push。
set -uo pipefail
DIR="$1"; BR="$2"; MSG="$3"
git config user.name "github-actions[bot]"
git config user.email "41898282+github-actions[bot]@users.noreply.github.com"
CUR="$(git rev-parse --abbrev-ref HEAD 2>/dev/null || echo detached)"
if [ "$CUR" = "$BR" ]; then
  git add -A "$DIR"
  git diff --cached --quiet || (git commit -qm "$MSG" && git pull -q --rebase -X theirs && git push -q)
  exit 0
fi
TMP="$(mktemp -d)"; mkdir -p "$TMP/x"
for sub in out requests; do [ -d "$DIR/$sub" ] && cp -r "$DIR/$sub" "$TMP/x/"; done
git stash -q -u || true
if git fetch -q origin "$BR" 2>/dev/null; then git checkout -q -B "$BR" "origin/$BR"; else git checkout -q -B "$BR"; fi
mkdir -p "$DIR"
for sub in out requests; do [ -d "$TMP/x/$sub" ] && { rm -rf "$DIR/$sub"; cp -r "$TMP/x/$sub" "$DIR/"; }; done
git add -A "$DIR"
git diff --cached --quiet || (git commit -qm "$MSG" && git push -q origin "$BR")

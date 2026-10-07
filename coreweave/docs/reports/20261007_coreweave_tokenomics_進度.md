# CoreWeave × Tokenomics 改造進度（接手用；W1–W3 共用）

## 目前狀態（每次 push 前覆寫）
- 已完成：W0 遷移（chat 端）；W1 步驟 0–3（快照 v5.24、company.json `tokenomics` 區段；取數工具 `tools/tokenomics/import_tokenomics.py`；分支 `claude/coreweave-w1-tokenomics` 自 `claude/coreweave-w0-migrate` a29db91 開出、draft PR、Tokenomics 唯讀副本 master 098873a／v5.24）。
- 下一步：W1 步驟 4——Excel「Tokenomics_取數」分頁（build_xlsx.py）、具名範圍 TK_*、verify.sh 新增 --check 與「快照值＝Excel 分頁值」檢查（步驟 3 已併入步驟 2：10 個 v5.25 名稱列為 optional）。
- 未解問題：Tokenomics v5.25（IF_DeprLifeIT 等 10 個名稱）尚未合併到 Tokenomics master；本張以 v5.24 產生快照，10 個名稱列為 optional（記為 missing）。

## 工作單總覽
| 工作單 | 分支 | PR | 狀態 |
|---|---|---|---|
| W0 遷移 | `claude/coreweave-w0-migrate` | | chat 端完成 |
| W1 Tokenomics 取數層 | `claude/coreweave-w1-tokenomics`（疊加於 W0 分支） | #6 | 進行中 |
| W2 每 MW 改寫 | `claude/coreweave-w2-permw` | | 未開始 |
| W3 v4.6 成品與對照 | `claude/coreweave-w3-v4.6` | | 未開始 |

<!-- 各工作單在下方新增自己的段落：「## Wx」＋步驟紀錄表（步驟｜狀態｜commit｜備註） -->

## W1 Tokenomics 取數層

Tokenomics 版本：`model/CURRENT`＝`20261007_Tokenomics_v5.24.xlsx`，master HEAD `098873a3c6855d1ef5e3414d5a0f12e39b2644ed`（2026-10-07 讀取）。

| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 開分支、draft PR #6、Tokenomics 唯讀副本 | 完成 | d4775c1 | W0 未合併：自 `origin/claude/coreweave-w0-migrate` a29db91 開分支，PR base＝W0 分支 |
| 1 取數工具＋根目錄 README 段落 | 完成 | （本 commit） | 測試：產生 v5.24 快照 25 名（missing 10）；`--check` 通過；`--force-recalc`（LibreOffice 重算副本）值＝快取值；IF_Hdr*／不存在名稱報錯；竄改快照 1e-6 時 `--check` 失敗；無 clone 時警告略過 |
| 2 名稱清單、快照、company.json `tokenomics`、README 欄位表 | 完成 | （本 commit） | `coreweave/data/tokenomics_names.txt`（15 個現有名稱＋10 個 optional；另加 `IF_RacksPerGW`、`IF_Util` 供對照）；`coreweave/data/tokenomics_snapshot_v5.24.json`；`fields_doc.py --write` |
| 3 v5.25 名稱列 optional | 完成（以 v5.24 產快照） | 同上 | Tokenomics master 仍為 v5.24（098873a）；10 名記為 `{"missing": true}`。**W2 第 0 步：若 master 已是 v5.25，重抓快照（檔名改 v5.25）並刪 v5.24 快照** |

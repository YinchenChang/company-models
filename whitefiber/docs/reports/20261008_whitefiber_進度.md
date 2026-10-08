# WhiteFiber v0.1 分段建置進度（接手用；三張工作單共用）

## 目前狀態（每次 push 前覆寫）
- 已完成：v0.1a 步驟 0（基線 verify.sh 19 項全過；進度檔建立）。
- 下一步：v0.1a 步驟 1：SEC EDGAR 一手文件（FY2025 10-K、2026 Q1／Q2 10-Q、8-K 財報新聞稿、S-1／424B、2026-08 可轉債 8-K、XBRL companyfacts）→ 寫入 `whitefiber/data/whitefiber_facts_20261008.json` 群組 1。
- 未解問題：無。

## 工作單總覽
| 工作單 | 分支 | PR | 狀態 |
|---|---|---|---|
| v0.1a 資料蒐集 | `claude/whitefiber-v0.1` | #23（draft） | 進行中 |
| v0.1b 模型改寫 | `claude/whitefiber-v0.1` | 同一 PR | 未開始 |
| v0.1c 驗證與成品 | `claude/whitefiber-v0.1` | 同一 PR | 未開始 |

<!-- 各工作單在下方新增自己的段落：「## v0.1x」＋步驟紀錄表（步驟｜狀態｜commit｜備註） -->

## v0.1a 資料蒐集

- 分支：`claude/whitefiber-v0.1`（自 af4af27）；PR #23。

### 步驟紀錄
| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 基線 verify、進度檔 | 完成 | （本 commit） | verify.sh 19 項全過（約 2 分 38 秒；cmp31 每情境 398 項）。soffice 24.2、openpyxl、playwright 1.56.0 已在環境；補裝 Playwright Chromium（`python3 -m playwright install chromium`）。複製無遺漏（company.json 引用的 data 檔齊全）。 |

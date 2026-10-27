# MAG 財報萃取（API 版，2026-10-27 起）

把 10-Q／10-K 與財報新聞稿（SEC 8-K 附件 99.1）中 MAG 模型要更新的欄位萃取成**候選值** Excel。
**不改 company.json**；候選值須逐項對原文核對後，才由建置代理依 MAG 工作單填入。

## 文件來源（2026-10-27 定）
- SEC 會擋 GitHub 雲端 IP（實測 403），**不得繞過**。Andy 10-27 選定：每季由 Andy 手動下載 10-Q 與財報新聞稿到 Drive `02-算力收支/MAG財報下載/`，其餘自動處理。
- 請求檔用 `files`（repo 內路徑，例 `mag/extract/docs/microsoft_FY27Q1_10-Q.pdf`）或 `drive_ids`（Drive 檔案 id；需該資料夾開放「知道連結者可檢視」）。兩者皆無時才走 EDGAR 直連（目前不可用）。
- PDF 用 `pdftotext -layout` 轉文字；HTML 轉文字；其餘當純文字。

## 怎麼跑
1. 在分支 `claude/mag-extract` 的 `mag/extract/requests/` 放一個 JSON 請求檔並 push：
   ```json
   {"company": "microsoft", "period_end": "2026-09-30", "target_q": "FY27Q1", "mode": "live",
    "files": ["mag/extract/docs/microsoft_FY27Q1_10-Q.pdf", "mag/extract/docs/microsoft_FY27Q1_新聞稿.pdf"]}
   ```
   - `company`：company-models 中有 `company.json`＋`calendar_q.py` 的公司資料夾（MAG 三家，亦可 coreweave、oracle 等同架構模型）；`period_end`：該季期末日（EDGAR reportDate）；
   - `target_q`：`quarterly.actuals` 的季度鍵（Microsoft 用 FYyyQn，另兩家用 yyyyQn）；
   - `mode`：`live`（新季）或 `backtest`（已申報季，與 company.json 現值比對準確率）。
2. 工作流程 `mag-extract`（push 觸發）自動執行，結果 commit 回同一分支：
   `mag/extract/out/<請求名>.xlsx`（候選值頁：紅＝兩模型不同、黃＝只一個找到）與 `.md` 摘要；請求檔移到 `requests/done/`。

## 設計
- 欄位清單由 `<公司>/company.json` 自動產生：`latestQuarter`、`ytdActual`、`callFacts`、`quarterly.actuals[target_q]`、`calendar_q.py → ROLL_FIELDS`；說明取 `scripts/fields_doc.py`。
- 同一份文件由 claude-opus-5-5 與 claude-sonnet-5-5 各萃取一次（Batch API，五折），逐欄比對。
- 費用：每家每季約 US$0.5–1（Max 方案每月 API credit 支付）。
- 限制：法說會內容（多數 `callFacts`）不在 SEC 文件中，會標「皆未找到」；ROLL_FIELDS 多為推導值，只回報原始數字與出處。
- 對 SEC 的請求以 GitHub 免回覆地址 `YinchenChang@users.noreply.github.com` 作聯絡方式（SEC 要求；Andy 2026-10-27 同意，不用個人信箱）；每秒少於 10 次。

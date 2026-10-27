# MAG 財報萃取（API 版，2026-10-27 起）

把 10-Q／10-K 與財報新聞稿（SEC 8-K 附件 99.1）中 MAG 模型要更新的欄位萃取成**候選值** Excel。
**不改 company.json**；候選值須逐項對原文核對後，才由建置代理依 MAG 工作單填入。

## 怎麼跑
1. 在分支 `claude/mag-extract` 的 `mag/extract/requests/` 放一個 JSON 請求檔並 push：
   ```json
   {"company": "microsoft", "period_end": "2026-09-30", "target_q": "FY27Q1", "mode": "live"}
   ```
   - `company`：microsoft／alphabet／amazon；`period_end`：該季期末日（EDGAR reportDate）；
   - `target_q`：`quarterly.actuals` 的季度鍵（Microsoft 用 FYyyQn，另兩家用 yyyyQn）；
   - `mode`：`live`（新季）或 `backtest`（已申報季，與 company.json 現值比對準確率）。
2. 工作流程 `mag-extract`（push 觸發）自動執行，結果 commit 回同一分支：
   `mag/extract/out/<請求名>.xlsx`（候選值頁：紅＝兩模型不同、黃＝只一個找到）與 `.md` 摘要；請求檔移到 `requests/done/`。

## 設計
- 欄位清單由 `<公司>/company.json` 自動產生：`latestQuarter`、`ytdActual`、`callFacts`、`quarterly.actuals[target_q]`、`calendar_q.py → ROLL_FIELDS`；說明取 `scripts/fields_doc.py`。
- 同一份文件由 claude-opus-5-5 與 claude-sonnet-5-5 各萃取一次（Batch API，五折），逐欄比對。
- 費用：每家每季約 US$0.5–1（Max 方案每月 API credit 支付）。
- 限制：法說會內容（多數 `callFacts`）不在 SEC 文件中，會標「皆未找到」；ROLL_FIELDS 多為推導值，只回報原始數字與出處。
- 對 SEC 的請求不帶使用者 email；每秒少於 10 次。

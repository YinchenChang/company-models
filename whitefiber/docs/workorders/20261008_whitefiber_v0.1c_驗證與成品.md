# 工作單 WhiteFiber v0.1c｜驗證、報告、交接檔、成品（2026-10-08，r0，chat 端）

- 前置：v0.1b 已完成並經 chat 端審查（審查留言在 PR，與本工作單同等效力）。
- 必讀：`20261008_whitefiber_共同規則.md`、v0.1a 資料報告、v0.1b 模型報告、進度檔、對照表 r1。

## 步驟（每步完成即 commit＋push＋更新進度檔）
0. 進度檔新增「v0.1c」段落；先處理 v0.1b 審查留言交代的事項（步驟 0.5，若有；改數字者列前後對照）。
1. **版本與命名**：`vlog.py` 與 `tail.js` 的 VLOG 重設為 WhiteFiber 版本紀錄，第一列 `v0.1`（日期＝建置日；說明「由 CRWV v4.5／Nebius v0.1／Oracle v0.2 模板建立；雲端 MW × 每 MW 收入、託管分部、GPU 租賃、預付、可轉債、專案融資瀑布、建設延誤、分部 EV/EBITDA」）；CRWV／Nebius／Oracle 歷史一句話帶過。成品名稱由 `meta.company` 組成（`YYYYMMDD_WhiteFiber收支模型_v0_1`）；`meta.updateDate` 改建置日。
2. **成品**：`whitefiber/dist/` 移除 Oracle 成品，放入 `YYYYMMDD_WhiteFiber收支模型_v0_1.html` 與 `.xlsx`；HTML 單一檔案、離線可開。
3. **完整驗證**：`verify.sh` 全過；再跑 `verify.sh --vs-dist`（對剛放入的成品，應 0 差異）。`crawl.py`：描述本公司的 Oracle／CoreWeave／Nebius 為 0（同業、對照、來源、版本紀錄除外，逐項列出）。完整輸出存入報告。
4. **交接檔與文件**：`whitefiber/docs/handoff/` 以 `YYYYMMDD_WhiteFiber收支模型_交接檔_v0_1.md` 取代 Oracle 版（只留一個檔）：命題、結構與因果、目前結果、資料來源與標記、已套用的預設、資料缺口、已知限制、待辦（含 2026 Q3 財報〔約 2026-11 中〕後的季度更新步驟）、如何重建。`whitefiber/README.md`、`whitefiber/CLAUDE.md` 開頭改為 WhiteFiber（守則條文保留，Oracle 待辦改為 WhiteFiber 待辦）。repo 根目錄 README 表格新增 WhiteFiber 一列：`whitefiber/`｜WhiteFiber（WYFI）｜`oracle/` @ `e4540c3`（Oracle v0.2 分支）｜v0.1。
5. **最終報告** `whitefiber/docs/reports/20261008_whitefiber_v0.1.md`（結論先行）：
   - 一句話結論：命題（長約託管能否成為可融資現金底座、支應擴張而不反覆稀釋）目前的答案與信心。
   - 三情境目標價、區間、評等；與共識目標價差距及原因。
   - 雲端：實現單價 vs 正向推導單價 vs 合約隱含單價，說明原因與對命題的意義。
   - 託管：NC-1 每 MW 經濟性（租金、建置成本、EBITDA、回收年數、專案融資可借額度）。
   - 敏感度：雲端每 MW、合約單價、託管租金與建置成本、Nscale 違約、NC-1 融資未完成、延誤月數、債務上限。
   - 已套用的預設（v0.1a–c 彙整）、資料缺口、已知限制（必含：續約價格衰退未做；Nscale 對手方信用未建模；雲端合約到期後以 Tokenomics 單價續作；上市時間短、beta 與共識樣本小）。
   - verify.sh 完整輸出（附錄）。
6. **PR 回報**並轉為 ready（`gh api -X POST repos/YinchenChang/company-models/pulls/<n>/ccr/ready_for_review`；不可用則在留言註明「已完成，可審查」）。

## 驗收
- `verify.sh` 與 `--vs-dist` 全過；`whitefiber/dist/` 只有 WhiteFiber v0.1 兩個成品，數字一致。
- 離線開啟檢查通過；一頁摘要在無網路下正常顯示。
- 交接檔可讓沒看過本對話的人重建與更新模型。
- 報告每個關鍵數字可追溯到 `company.json`／`data/` 與來源。

## 停止條件
依共同規則第 6 節。

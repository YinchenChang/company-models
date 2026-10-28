# 來源複核（API 版）

用途：核對各模型中標記為 Verified／Interested-party 的輸入值，是否與「來源／說明」欄所列來源一致。只產出結果，不改模型。

- 程式：`audit.py`；工作流程：`.github/workflows/source-audit.yml`（手動執行，可設累計花費上限 `max_usd`，預設 40 美元）。
- 對象：Siliconomics 快照列，加上各公司 dist 最新 xlsx 的 In 工作表；來源寫 Siliconomics／markets／Tokenomics 的繼承列不重複查。
- 模型：claude-sonnet-5-5（Batch API），工具 web_search、web_fetch。頁面打不開一律判「來源打不開」，不使用快取、鏡像或其他繞道。
- 判定：相符／部分相符／不符／來源打不開／找不到來源。
- 結果在分支 `source-audit-results` 的 `tools/source_audit/out/`：`source_audit.xlsx`（摘要、明細）、`source_audit.md`、`state.jsonl`（進度，重跑時接續）。
- 判定由模型產生，「不符」與「部分相符」須由維護串打開來源確認後才改模型。

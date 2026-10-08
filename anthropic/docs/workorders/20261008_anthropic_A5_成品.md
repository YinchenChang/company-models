# 工作單 Anthropic A5｜成品：HTML 一頁摘要、命題與驅動對照表、交接檔（2026-10-08，chat 端）

- 分支 `claude/anthropic-v0.1`；A4 經 chat 端審查後開工。參考 OpenAI v0.6 `tools/build_html.py`、`tests/test_html.py`、`docs/plan/20261008_OpenAI_命題與驅動對照表.md`。
1. **命題與驅動對照表** `docs/plan/YYYYMMDD_Anthropic_命題與驅動對照表.md`（格式同 OpenAI 版：一句話命題、主因果鏈、各頁輸入／衍生量／因果方向／是否繼承 Tokenomics、繼承 vs 公司專屬、與 OpenAI 模型的關係、三個關鍵驅動）。
2. **HTML 一頁摘要** `dist/YYYYMMDD_Anthropic收支模型_v0_1.html`：只讀 Excel 重算後的數值；單一檔案、離線雙擊可開、無任何外部資源（圖表內嵌 SVG）；內容同 OpenAI 版＋**與 OpenAI v0.6 並排**一節（每 VR 等值 GW 營收、全成本、差額、覆蓋率、累計外部資金需求）；每個數字附 Excel 位置；測試：HTML 數字與 Excel 一致、無外部 URL。版面風格與 OpenAI 版一致。
3. 複製 Excel 到 `dist/YYYYMMDD_Anthropic收支模型_v0_1.xlsx`；交接檔 `docs/handoff/Anthropic_handoff.md` 更新為 v0.1 完成狀態；repo 根目錄 README 的 Anthropic 列。
4. 驗收：tests 全過；報告 `docs/reports/YYYYMMDD_v0.1_成品.md`；PR 留言、轉 ready（`gh api -X POST repos/YinchenChang/company-models/pulls/<n>/ccr/ready_for_review`）。

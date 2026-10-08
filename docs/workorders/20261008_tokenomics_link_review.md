# 工作單（各公司共用）｜Tokenomics 連接審視（2026-10-08，chat 端；只審視回報，不改模型）

- 發出：chat 端（Claude，Project「算力收支」）。Andy 2026-10-08：「叫他們都再審視一下跟這個 Tokenomics 連接了嗎」「統一去管理他們的更新」。
- 適用：anthropic、coreweave、minimax、nebius、openai、oracle、whitefiber、zhipu 八個資料夾，各自執行一次，報告放各自的 `docs/reports/YYYYMMDD_<公司>_tokenomics_link_review.md`。
- 依據：Tokenomics repo `docs/plan/Tokenomics_downstream_contract.md`（下游資料契約 v0.1）；本 repo `docs/reports/20261008_tokenomics_link_audit.md`（工具初步盤點，含誤判可能）。
- **前提**：Tokenomics v5.29 已合併（`model/CURRENT`＝`*_Tokenomics_v5.29.xlsx`；`IFW_`、`IFC_` 名稱存在）。未合併時只做第 1–3 節，第 4 節標「待 v5.29」。
- **本單不改模型**：不改 company.json、builder、Excel。改動由 chat 端依報告另開工作單。

## 1. 取數點清單（必做）
每一個取自 Tokenomics 的數值一列：

| 本模型的格或參數 | Tokenomics 名稱 | 合規（IF_／IFW_／IFC_／L1_／SRC_ Active）或違規（其他前綴、工作表直接引用、寫死數字） | 四層瀑布的哪一層（100%／IF_Util／Prod／Life／不適用） | 用途是否符合 `IFC_Use`（「上限」列不得當收入預測） | 所用 Tokenomics 版本（檔名、合併雜湊） |

先跑 `python3 tools/tokenomics/link_audit.py --tokenomics <Tokenomics clone>` 取得候選清單，再逐筆人工確認（工具會把公司自己的同名工作表如 `Inputs!` 誤判為違規，須標明）。

## 2. 口徑問題（必答）
1. 本模型收入按 token、GPU-hour 還是 MW-year 計價？對應契約第 3 節，利用率與簽約率各用哪一個？有無在收入端重複扣減？
2. 神雲／雲平台模型：每 MW 收入是否仍用「成本加成與市場價 50/50 平均」？容量交付情境與單價情境是否同一個選擇器同向綁定？（契約第 4 節：應分離；CoreWeave W4 為範本）
3. 模型商模型：`IF_Alloc*` 六格（Q1、Q2、服務 GW、需求 D、隱含 N×k）取用處；v5.29 這六格因每則提示 token 數與推論支出口徑的判斷而改變，列出受影響的輸出。
4. 公司特有變數（WACC、合約價、價值捕獲率、簽約率、交付進度）是否只在本模型覆寫、沒有回寫 Tokenomics？

## 3. 缺口與版本
- Tokenomics 缺的產業級數據：列出（名稱、用途、目前以什麼代替），不自行建表；由 chat 端走 G2 進 DB_Evidence。
- 本模型應改用官方快照工具的名稱清單草稿：`data/tokenomics_names.txt`（只列 IF_／IFW_／IFC_／L1_），附在報告，不執行。

## 4. v5.29 重取影響（前提成立時）
- 以 `import_tokenomics.py` 產生 v5.29 快照到暫存路徑（不提交），用 `--check` 對本模型現用值比較，列出變動的名稱與幅度；特別列出 `IF_Alloc*` 與 `IFW_*_Prod`／`_Life` 相對目前所用層的差異。

## 5. 報告格式
繁體中文、非工程語言；開頭列分支、SHA、報告路徑；同內容貼成 PR comment；末節「待 chat 端判斷」。合併前由 chat 端審查。

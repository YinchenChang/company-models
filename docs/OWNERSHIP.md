# 路徑歸屬（OWNERSHIP）

本 repo 由三個研究專案共用。每個路徑只有一個負責專案；共用區由 AI供應鏈 維護。CI `ownership-check`（`tools/ownership/check_ownership.py`）依下表分類每個 PR 的變更檔。

## 路徑 → 負責專案

依下表由上而下比對，第一個符合的規則決定歸屬。

| 路徑 | 歸屬群組 | 負責專案 | 說明 |
|---|---|---|---|
| `software_ai/**` | AI軟體應用 | AI軟體應用 | AI 軟體應用個股模型、AI 對軟體單位經濟 |
| `physical_ai/**` | AI實體應用 | AI實體應用 | 實體 AI 個股模型與實體 AI 地圖（目前只在分支 `physical-ai-v0.1`） |
| `tools/**` | 共用 | AI供應鏈 維護 | 共用工具（HTML 展示層、Tokenomics 取數、治理檢查、Drive 同步清單等） |
| `markets/invest_map/**` | 共用 | AI供應鏈 維護 | AI 算力地圖（GitHub Pages 站根；`publish_pages.sh` 須保留其他專案的站內子目錄） |
| `docs/plan/**` | 共用 | AI供應鏈 維護 | 跨模型規劃與原則文件 |
| `.github/**` | 共用 | AI供應鏈 維護 | CI workflow |
| `README.md`（根目錄） | 共用 | AI供應鏈 維護 | 多專案索引與 AI供應鏈 總表 |
| `docs/OWNERSHIP.md` | 共用 | AI供應鏈 維護 | 本檔 |
| 其他所有路徑 | AI供應鏈 | AI供應鏈 | 各公司資料夾、`siliconomics/`、`wattonomics/`、`markets/`（`invest_map/` 除外）、`mag/`、`docs/`（`plan/` 除外）等 |

Tokenomics 在另一個 repo（`YinchenChang/Tokenomics`），不適用本表。

## 規則

1. **一個 PR 只改自己專案的路徑。** 「共用」視為獨立的一個群組：只改共用區的 PR 也算單一群組。
2. **跨群組須標示。** PR 同時改到兩個以上群組（例如 AI供應鏈 ＋ 共用、AI軟體應用 ＋ 共用）時，PR 標題必須以「[跨專案]」開頭，並在 PR 說明寫明對應的交流訊息（交流資料夾的訊息檔名或日期）或理由。未標示時 `ownership-check` 失敗。
3. **共用區由 AI供應鏈 維護。** 其他專案需要改共用區（例如共用工具加功能、地圖加導覽連結）時，先以交流訊息告知 AI供應鏈；由 AI供應鏈 修改，或由提出方開「[跨專案]」PR 並在說明引用該訊息。
4. **不改他人的成品。** 不得改動其他專案的 `dist/` 成品、模型程式與資料；發現問題時以交流訊息通知負責專案。
5. **CI 隔離。** 各專案的建置檢查只在自己的路徑有變動時執行：`ai-semis-build` 排除 `software_ai/**` 與 `physical_ai/**`；`software-ai-build` 只在 `software_ai/**` 有變動時執行。

## 本機檢查

```
python3 tools/ownership/check_ownership.py --files software_ai/x/a.py tools/y.py --title "[跨專案] 範例"
python3 tools/ownership/check_ownership.py --base origin/main --title "<PR 標題>"
```

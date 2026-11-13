# tools/model_governance — 個股模型治理核對

起因：2026-11-09 SpaceX v0.2.5 驗收報告（M1 Sens 失同步、M8 標記混用）。原則見 `docs/plan/個股模型原則_命題加營收分部.md` 第 10 節。

`python3 tools/model_governance/check_governance.py <重算後 xlsx> [--html <成品 html>]`

| 檢查 | 失敗條件 |
|---|---|
| build_id | Sens 頁無「build_id」列；README!A1 或 HTML 不含同一字串 |
| 模型狀態 | Sens 頁「模型狀態」不是 CURRENT |
| 單一標記 | In 頁 G 欄、Q 頁「標記」欄不是 Verified／Interested-party／Analogy／Assumed／Derived／— 之一；來源以「…；Derived）」或「…；Interested-party）」結尾卻標 Verified |
| 來源誘因 | 沒有「來源誘因」欄，或 Interested-party 列該欄空白 |

已套用：SpaceX v0.2.6。其他個股於下次改版時加入建置腳本（做法照 `spacex/build_spacex.py` 的 `BUILD_ID`、`norm_tag`、Sens 指紋三段）。

## 展示層數字一致性（原則第 10 節第 51 條；2026-11-11 起，首例 NVIDIA v0.3.5）

`python3 tools/model_governance/check_display.py <公司資料夾>`（或 `check_governance.py <xlsx> --html <html> --display <公司資料夾>`）

| 檢查 | 失敗條件 |
|---|---|
| README 頁 | 關鍵數字（`display_check.json` 的 items：目標價、上檔空間、EPS、倍數等）的格式化值未出現在非沿革列；pattern 擷取到與來源儲存格不同的數字（手寫舊數字） |
| HTML | 指向來源儲存格的數字 data-v 與來源不同；關鍵數字未出現；HTML 內 README 頁文字與 xlsx 快取值不同 |
| repo README.md | 版本表（`<!-- version-table: … -->` 標記）與成品重算不一致；沒有以目前 dist 成品為最後一欄的版本表 |

版本表由 `python3 tools/model_governance/version_table.py <公司資料夾> --update` 從 dist（及 `_待刪除/`）的成品產生，不手寫。設定檔 `<公司>/display_check.json`（items、history_re、table_rows）。其他個股於下次改版時加入。

## 收支關卡評等檢查（原則第 10 節第 63 條）

`python3 tools/model_governance/check_rating_gate.py <成品.xlsx>`：基準情境收支合計為負而「投資建議」為買進或強力買進時失敗（應為「收支不足」）；沒有「收支關卡」列的模型回報「不適用」。`--map markets/invest_map/dist/data.json` 盤點地圖判定關卡不通過、估值評等卻為買進以上的公司。新建與改版的個股模型在建置腳本內必跑。

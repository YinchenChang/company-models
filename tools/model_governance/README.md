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

## 核對列與錯誤值掃描（2026-11-20 起；CI `zero-rows-check`）

`python3 tools/model_governance/check_zero_rows.py --all --baseline tools/model_governance/zero_rows_baseline.json`

| 檢查 | 失敗條件 |
|---|---|
| 核對列 | 標籤含「應為 0」的列（或「應為 0」節標題下以「核對：」「檢查：」開頭的列）有數值的絕對值超過容差（預設 0.005；標籤寫明「捨入」者 0.2） |
| 錯誤值 | 任何工作表出現 `#REF!`、`#NAME?`、`#VALUE!`、`#DIV/0!`、`#N/A`、`#NUM!`、`#NULL!` 或 `Err:` |

不判的列：以「差額」開頭（刻意揭露）、「應接近 0」（逆向 DCF 搜尋殘差）、只有部分欄或部分情境應為 0、Sens 頁。

基準清單是已知待修項目，只減不增。寫新的核對列時，請讓標籤含「應為 0」，並用具名輸入或來源常數產生核對式，不要手寫會過期的數字（起因見 `docs/reports/20261120_核對列掃描與查因.md`）。

容差：標籤寫「捨入」或「四捨五入」的列用 `--tol-round`；標籤寫明數字（例：「四捨五入差 < 0.01」）時用標籤的數字（2026-11-20 第二波 W2-3）。

淨現金核對列的寫法：建置腳本存檔前呼叫 `tools/fcff_dcf.py` 的 `sync_nc_check(wb)`，核對式由 In 頁 `cash_v`、`inv_v`、`debt_v`、`lease_v`、`mi_v` 當下的內容產生（台積電、智邦、Astera Labs 已改；其餘模型仍是 `write_checks(nc_bs=…)` 的手寫數字，改版時一併換）。

## 公式格快取值檢查（`check_cached_values.py`；2026-11-20 第二波 W2-6）

`python3 tools/model_governance/check_cached_values.py --all`（CI `zero-rows-check` 的第二步）

成品的公式格沒有快取值即失敗。openpyxl 直接寫出的檔沒有計算結果；存進 `dist/` 前要經 LibreOffice 重算。直接由腳本存檔的活頁簿（兩本地圖活頁簿）在存檔後呼叫 `tools/xlsx_recalc.py` 的 `recalc_inplace(路徑)`。公式重算後回傳空字串的格不算失敗。

## 其他檢查遇到不適用的檔案時（2026-11-20 第二波 W2-7、W2-9）

| 檢查 | 情況 | 行為 |
|---|---|---|
| `check_display.py` | `dist/` 內不是剛好一個 xlsx、找不到 README.md | 回報失敗項目，不丟例外 |
| `check_display.py`、`version_table.py` | README 版本表引用的舊成品已不在工作樹（`dist/` 只留現行版） | 該欄凍結：沿用 README 既有文字，只檢查還在樹上的欄；整張表都不在時略過並列出數量 |
| `check_display.py` | 路徑分隔符 | 一律以「/」比對，Windows 上也對得上 |
| `check_governance.py` | 活頁簿沒有 In 頁（OpenAI、Anthropic、MiniMax、智譜的模型商結構） | 回報「不適用」，結束碼 0 |

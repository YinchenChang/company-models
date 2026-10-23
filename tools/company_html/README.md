# tools/company_html — 公司模型 HTML 展示層

Excel 是計算事實的來源；HTML 只負責呈現，給 Andy 與投資處同事閱讀。每家公司一個 `html.json`，由共用產生器讀取 `dist/*.xlsx`（LibreOffice 重算後的值）輸出單一自含 HTML。

| 檔案 | 內容 |
|---|---|
| `build_company_html.py` | 產生器：`python3 tools/company_html/build_company_html.py <公司資料夾>`，輸出 `<公司>/dist/<xlsx 同名>.html`（版本號與 xlsx 相同） |
| `check_company_html.py` | 檢核（CI 的 `ai-semis-build` 執行）：重現性、抽數、無手寫數字、自含、結構 |
| `layout_check.py` | 版面檢查（Playwright／Chromium）：寬 1280 與 390 截圖，檢查無水平捲動、無空白圖；截圖不入 repo |

## 頁面七章（`html.json` 的 `sections` 須依此順序）

`thesis` 命題與核心問題 → `drivers` 驅動因子→推導量 → `pl` 三情境損益（表格＋折線圖）→ `valuation` 估值 → `recon` 對帳 → `sens` 敏感度 → `status` 資料狀態。

## `html.json` 格式

頂層：`company`、`ticker`、`thesis`（儲存格參照，取自 xlsx README 頁的命題）、`question`（核心問題，純文字）、`scenarios`（`[{key, name}]`）、`sections`。每章：`id`、`title`、可選 `intro`、`blocks`。

儲存格規格字串：`@工作表!儲存格[:範圍][,…][|格式]`，例如 `@PL!C5:G5|n0`。範圍在表格列內自動展開成多格；其他字串是純文字。格式碼：`n0–n3` 數字、`p0–p2` 百分比（值×100）、`sp0–sp2` 帶正負號百分比、`x1` 倍數、`yr` 年、`tag`（標記欄加底色）、空白＝依型別（文字原樣、數字通用）。`@ds:資料集名!列鍵.欄位|格式`（例如 `@ds:valuation!美光 見頂.multiple|x2`）讀 `markets/overview/datasets/<資料集名>.json`，供 xlsx 內沒有、但規則需要的值（十年分位倍數）使用；檢核同樣回頭比對該 JSON。`$version`、`$date`、`$sha`、`$xlsx` 是中繼資料（取自檔名與 xlsx 雜湊）。

區塊型別：`table`（`head`＋`rows`，或 `rowgen`：依工作表列號與欄產生，可用 `row_fmt` 逐列覆蓋格式；`collapse`＋`summary` 可摺疊）、`pl`（每個指標一張情境×年度表，`chart:true` 另畫折線圖）、`bars`（隱含價長條與現價標線）、`tornado`（敏感度長條，依絕對值排序）、`kv`、`note`、`list`、`h3`、`scan`（掃描指定工作表中含關鍵字，如「未找到」的儲存格）、`upstream`（掃描 In 頁引用的上游 `dist/*.xlsx`，列出 repo 內現況與上游標題）。

## 規格與檢核

- **不得手寫數字**：設定檔的文字除年度（FY／CY＋四位數）與日期外不得含阿拉伯數字，建置時即報錯；所有數字與標籤取自 xlsx，並帶 `data-ref="工作表!儲存格"`（滑鼠懸停顯示，或按頁首按鈕顯示在數字旁）。
- **抽數檢核**：每個數字（含圖上的點）的 `data-v` 與 xlsx 重算值相對誤差 < 1e-9；顯示字串與值的差在顯示精度內；文字標籤與儲存格文字完全相同。
- **重現性**：重跑兩次輸出逐位元組相同，且與已提交的 HTML 相同；因此 xlsx 或設定檔變動後，須重建 HTML 一併提交。
- **自含**：內嵌 CSS 與 inline SVG；無外部網址、無 JavaScript；風格比照 AI 半導體價值鏈儀表板（白底、青綠主色；Andy 2026-10-23 指定，不做深色模式）；手機（390）與桌機（1280）無水平捲動。
- **建置 commit 的替代**：HTML 無法寫入自己的 commit（寫入即改變 commit），且 squash 合併會改 hash。改記**來源 xlsx 的雜湊指紋**（內容決定、可重算）與上游成品清單。

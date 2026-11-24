你在 repo YinchenChang/company-models 的工作樹根目錄。任務：外部稽核優化 W3-6——替 `{FOLDER}/` 這一家個股模型補齊治理項目，讓下面兩項檢查通過，而且**不改任何既有數值**。先讀 `docs/workorders/20261120_外部稽核優化_工作單.md` 的 W3-6 與 J4、`docs/reports/20261123_W3-6_治理項目盤點.md` 裡 `{FOLDER}` 那一列，再讀 `tools/model_governance/stamp_governance.py` 的檔頭說明。已做完的範例：`cohr/build_cohr.py`、`cohr/sens_cohr.py`（有整本重算快照）、`mpi/build_mpi.py`（沒有快照）——搜尋「W3-6」。

要通過的檢查（對重建、重算後的成品）：
- `python3 tools/model_governance/check_governance.py <xlsx> --html <html>`
- `python3 tools/sens_coverage/check_sens_coverage.py <xlsx> --require-keycol`

規則：
1. 只改 `{FOLDER}/` 內的檔案。不得改其他公司資料夾、`tools/`、`markets/`、`docs/`。不得 git commit、push 或開 PR（外層流程會處理）。不呼叫任何外部網路服務。
2. **不改值。** 目標價、評等、EPS、In 頁所有輸入、既有的每一個數值格與公式都不得改變。允許的改動只有：新增列、新增欄、標記文字、Sens 頁 B 欄（設定）文字後面附鍵、README!A1 末尾加 build_id。不升版號、不改成品檔名、不改日期。
3. 治理四項用共用函式，不要自己另寫：在建置腳本最後一個 `wb.save(...)` 之前呼叫 `stamp_governance.apply(wb, here=…, slug="{FOLDER}", ver="<現行版號>", snapshot_tp=…, sens_keys=…, exempt=…)`。
   - `ver` 用現行成品檔名的版號；README!A1 開頭的版號若與檔名一致可用正則讀出，否則寫死並加註解「升版時一起改」。
   - 有整本重算快照（`sens.json`、`sens_*.py`）時，`snapshot_tp` 給快照的基準加權目標價；敏感度重算、協調試算等「不是正式建置」的路徑（環境變數覆寫輸入時）不要呼叫。
   - 這一家若已有自己的 build_id／模型狀態／來源誘因段落（少數幾家），不要重複加；只補缺的。
4. Sens 鍵：
   - 先跑覆蓋檢查，看哪些鍵未涵蓋。對每個未涵蓋的鍵，回到建置腳本與 `sens_*.py` 找它的案例。
   - 已有案例、只是那一列沒有鍵：用 `sens_keys={案例列 A 欄文字: "鍵1 鍵2"}` 補上。鍵必須是該案例**真的會改到**的 In 鍵（看案例的覆寫內容，不要用標籤猜）。有快照的模型，建議像 `cohr/sens_cohr.py` 一樣把每案覆寫的鍵存進 `sens.json` 的 `keys`，建置時讀出來。
   - 確實沒有任何目標價案例的鍵：有整本重算腳本的模型，把它加進腳本的案例清單並重跑快照（既有案例與基準值不得改變；重跑後比對）。沒有整本重算腳本的模型，不要為此新寫一套重算流程。
   - 只作對照、結構上不影響加權目標價的鍵（例如只進 Recon 的共識數字）：用 `exempt={鍵: 理由}`。理由要寫出「為什麼對加權目標價恆為 0」，並且你要用公式引用鏈確認過；不確定就不要豁免。
   - 目標價敏感度表不在 Sens 頁前 40 列、檢查找不到表頭的模型：**不要搬動既有的表**。在回報寫明「表頭在第幾列」，鍵的對照照樣補好，這一項留給主控統一處理。
   - 做完仍未涵蓋的鍵：不要硬湊。逐一列在回報裡（鍵、In 頁項目文字、缺的是案例還是判斷）。
5. 主要標記：共用函式把混標（例「Verified／Interested-party（指引）」）改成一個主要標記，預設取最近一期那一段、有公司指引時取 Interested-party。你認為某一列預設取法不對時，用 `tag_override` 指定，並在回報寫出列與理由；不要改共用函式。
6. 依該公司慣例重建（建檔指令見 `.github/workflows/ai-semis-build.yml` 的 targets；LibreOffice 用私有設定檔 `-env:UserInstallation=file:///tmp/lo_{FOLDER}`），重建 xlsx 與 HTML（`{FOLDER}/gen_html_json.py`、`tools/company_html/build_company_html.py {FOLDER}`），成品放回 `{FOLDER}/dist/`（檔名不變，覆蓋現行檔）。
7. 驗證並把結果寫進回報：
   - 與改動前的成品比：所有原有的數值格相同（容差 1e-9；Sens 頁因新增列而位移時，按 A 欄標籤比）。`python3 tools/xlsx_diff.py` 的差異只應該是新增的列與欄、標記文字、附鍵文字、README!A1。
   - 兩項檢查的輸出；另跑 `tools/company_html/check_company_html.py {FOLDER}`、`tools/in_layout/check_in_layout.py <xlsx>`、`tools/model_governance/check_zero_rows.py <xlsx>`。改動前就沒過的項目不要求，但要寫明。
8. 做不完或遇到無法判斷的事：不要硬做。把已完成的部分留在可重建的狀態，回報卡在哪裡，然後結束。
9. 文字用台灣繁體中文，直述、不用行話。程式註解寫「外部稽核優化 W3-6（2026-11-23）」與原因。

完成後輸出一段不超過 12 行的摘要：build_id；兩項檢查結果；數值格是否全部相同；補了幾列鍵、新增幾個案例、寫了幾個豁免（各附鍵名）；仍未涵蓋的鍵；用了 `tag_override` 的列；改了哪些檔。

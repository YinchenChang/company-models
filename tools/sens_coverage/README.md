# tools/sens_coverage：敏感度覆蓋核對

規格：`docs/plan/個股模型原則_命題加營收分部.md` 第 9 節（程式開頭註解與第 9 節一致）。

- `python3 tools/sens_coverage/check_sens_coverage.py <成品 xlsx> [--require-keycol] [--groups 規則.json]`：單檔核對；有未涵蓋（或找不到 Sens 表頭、`--require-keycol` 下無鍵欄）即失敗（結束碼 1）；無 In 頁低／基準／高三欄＝不適用（通過）。建置腳本必跑，新模型加 `--require-keycol`。
- `python3 tools/sens_coverage/check_sens_coverage.py --all [--md 報告.md]`：掃描全部 `*/dist/*.xlsx`（每資料夾取檔名排序最後一檔），只報告；含「已排除鍵（不進模型）」與「警告」欄。
- 案例列：Sens 表頭（前 40 列、含「…目標價」欄名）以下、目標價欄有數值、A 欄不是「讀法：」「註：」等（開頭字後接冒號）的列。涵蓋：鍵欄含該鍵，或 A、B 欄（驅動、設定）寫出該鍵——含「不變／不改／維持／另見／不重解」的子句除外；說明欄提到的不算。
- IF_HoldEcon：只認鍵欄——寫 `he_all` 或全部有區間的成員鍵（A、B 欄與說明欄文字一律不算；無鍵欄的舊 Sens 頁只認 A、B 欄明寫 `he_all`）、且無否定字（待補／尚未／未定／未做／未列／不得／缺少／缺案例）；不得豁免。IF_HoldEcon 有基準無區間＝警告。
- 豁免：一句內含鍵＋「不列／豁免／免列／不另列」＋「：」後至少 4 字理由，且無否定字；否則列「豁免不合格」。
- 不進模型：In 頁說明欄完整片語「只作對照，不進模型」；追引用鏈（加引號工作表名、範圍、整欄、定義名稱；INDIRECT／OFFSET 視為引用），只可落在 Recon、README、Sens（下游也須留在這三頁）或 Val 顯示用格，否則失敗。
- 「候選案例列」提示只以項目文字前綴比對，僅供參考、列號可能錯指。
- 群組：程式內 `GROUPS_DEFAULT`；`--groups` 規則檔格式 `{"groups":[{"name":…,"key_re":…,"label_re":…,"aliases":[…],"strict":false}]}`。
- 測試：`python3 tools/sens_coverage/test_check_sens_coverage.py`（含 2026-10-30 獨立核對反例）。

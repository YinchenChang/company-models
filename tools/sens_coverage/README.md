# tools/sens_coverage：敏感度覆蓋核對

規格：`docs/plan/個股模型原則_命題加營收分部.md` 第 9 節（In 頁有低／高區間的鍵都要有 Sens 案例或寫明豁免理由；IF_HoldEcon 四世代以一案同時涵蓋，不得豁免）。

- `python3 tools/sens_coverage/check_sens_coverage.py <成品 xlsx> [--groups 規則.json]`：單檔核對；有未涵蓋即失敗（結束碼 1）；無 In 頁低／基準／高三欄＝不適用（通過）。建置腳本必跑。
- `python3 tools/sens_coverage/check_sens_coverage.py --all [--md 報告.md]`：掃描全部 `*/dist/*.xlsx`（每資料夾取檔名排序最後一檔），只報告。
- 涵蓋判定：Sens 頁「鍵」欄含該鍵，或案例區任一文字格以獨立識別字提到該鍵。無「鍵」欄的 Sens 頁只能靠文字，報告另列「疑似有案例（項目文字相符、未寫鍵）」供人工判讀。
- 豁免：Sens 頁表頭以上註記（A1／A2）或案例區文字中，有一句同時含該鍵與「不列／豁免／免列／不另列」。In 頁該列註明「不進模型」者（低／高只作對照）不列入。
- 群組：程式內 `GROUPS_DEFAULT`（IF_HoldEcon：he_h／he_b2／he_b3／he_vr，亦認 he_h100 等與項目文字「IF_HoldEcon｜」）；`--groups` 規則檔格式 `{"groups":[{"name":…,"key_re":…,"label_re":…,"aliases":[…],"text_all":[[…]],"strict":false}]}`。
- 測試：`python3 tools/sens_coverage/test_check_sens_coverage.py`。

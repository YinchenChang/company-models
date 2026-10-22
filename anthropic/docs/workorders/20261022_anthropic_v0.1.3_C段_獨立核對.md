# 工作單 Anthropic v0.1.3 C 段｜獨立核對（唯讀；2026-10-22，chat 端）

- 執行者必須是**未參與 A、B 段**的新工作階段。只讀：不得 commit、push、修改任何追蹤檔；暫存檔放 `/tmp/verify`。核對分支 `claude/anthropic-v0.1.3` 最新提交（用 worktree）。不得在對外請求中使用使用者 email。
- 檢查（逐項 PASS／FAIL＋證據）：
  1. LibreOffice 強制全簿重算（OOXMLRecalcMode＝always；先移除快取值的副本）錯誤 0、與快取值差 0；`CHK_Errors`＝0。
  2. 自行重算：U2 新法下 Nscale／Lambda 年付款與 GW；每 VR 等值 GW 差額 2025–2030 與覆蓋率；資金三層各自的累計外部資金需求與現金谷底；擴張情境 3 檔的差額與資金結果；η 拆解的乘積恆等式（因子乘積 × 殘差＝η）。
  3. 前後對照：把 U2 開關切回 v0.1.2 做法，模型應重現 v0.1.2（`model/archive/20261021_Anthropic_v0.1.2.xlsx`）的所有基準輸出；TK v19.0 標記不改任何數值。
  4. Bridge 頁不被任何計算頁引用（掃公式）；η 值未因拆解改變。
  5. 報告與對照 Excel 每個數字對 Excel；HTML：以 Playwright 打開，逐節讀畫面數字（至少 60 個，含三個新小節與「v0.1.3 更新了什麼」）對重算 Excel；無外部 URL／資源；dist xlsx 與 model xlsx 逐位元組相同。
  6. 文字：結論與「要成立必須」是否為條件式；外部驗收報告逐項處理表與共同說明一致。
  7. `python3 -m pytest tests/test_html.py tests/parity/test_builder.py -q`（在 anthropic/ 內）。
- 結果寫成 PR 留言（第一行 `[Anthropic 核對] v0.1.3｜<OK／OK 附註／須修正>｜日期`）：FAIL 逐項列位置與正確值。只留言，不改檔。

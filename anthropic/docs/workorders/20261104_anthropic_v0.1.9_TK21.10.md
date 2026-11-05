# 工作單 Anthropic v0.1.9｜Tokenomics v21.7 → v21.10 同步（2026-11-04，chat 端）

- **來由：** Tokenomics 已連升三版。
  - v21.8：高成本情境 IT 折舊 4 → 5 年，只動高成本欄。
  - v21.9：TPU／Trainium3 產出比改用原生精度。
  - v21.10：產出比改 FP8 對齊作基準，原生精度降為低端（Tokenomics PR #193）。
  - 三版合計，本模型讀的 NonNV 產出比變動如下，基準欄會變：

    | 晶片 | 產出比（低／基準／高） | v21.7 | v21.10 |
    |---|---|---|---|
    | TPU v7 | `TK_NNV_TPUv7_Out*` | 0.6／0.8／1.1 | **0.35／0.66／0.84** |
    | Trainium3 | `TK_NNV_Trn3_Out*` | 0.3／0.5／0.8 | **0.14／0.28／0.50** |

  - 持有比自 v21.7 起沒有變。
  - 以上數值以 TK 實檔為準；chat 端只依交接檔記載，**建置代理要先跑 `check_tk_snapshot` 列出全部 DIFF**。
- **預期影響：**
  - TPU、Trainium 每 GW 產能下降 → η 重校（2025 推論支出校準）。
  - 有效推論／研發 GW 分攤、差額、覆蓋率、資金、現金谷底都可能改變。
  - 合約付款（GW × 持有比）應不變。
  - 任何「應不變」的格子若改變，停下列入「待 Project 判斷」，不自行處理。
- **分支：** `claude/anthropic-v0.1.9`（從 main 開）。開 PR 到 main，**不得合併**（chat 端驗收後合併）。
- **必讀：**
  - `anthropic/CLAUDE.md`
  - `docs/workorders/20261008_anthropic_共同規則.md`
  - 上一張工作單 `docs/workorders/20261102_anthropic_v0.1.8_S1A滾動.md`
  - 報告 `docs/reports/20261103_v0.1.8.md`
  - `docs/handoff/Anthropic_handoff.md`
  - 進度檔
- **規則同前：**
  - 新列加在尾端；公式不含常數。
  - 每步 commit＋push＋更新進度檔。
  - pytest 依檔分開跑。
  - 不得合併、不得 force push、不得改 main。
  - 不得在任何對外請求中使用使用者 email。
- **Tokenomics：** `GIT_LFS_SKIP_SMUDGE=1 git clone --depth 1 https://github.com/YinchenChang/tokenomics /tmp/tk`。驗證 `model/CURRENT`＝v21.10，並記錄 master SHA。

## 步驟（每步 commit＋push＋進度檔）
1. **快照 v21.10。**
   - TK_Link 重建；產出 `model/<日期>_Anthropic_v0.1.9.xlsx`。v0.1.8 檔 `git mv` 到 archive，並更新 CURRENT。
   - 列出全部 DIFF 名稱、新舊值，以及各名稱在本模型的使用位置（進計算或只進對照）。
2. **差異拆解（只動該項）：**
   - 依序只換 (a) TPU 產出比、(b) Trainium3 產出比、(c) 其他 DIFF。
   - 每步列 η、差額 2026–2030、2030 覆蓋率、累計外部資金、現金谷底（年、最低現金）。
3. **敏感度（不進基準，只寫進報告）：**
   - (i) 產出比取 TK 低端（＝v21.9 原生精度）。
   - (ii) 產出比取 TK 高端。
   - (iii) v21.7 舊值（對照）。
   - 每列同上欄位，看命題 1（持續轉正年）、命題 2（是否需外部資金）有沒有翻轉。
4. **重算與驗收：**
   - LibreOffice 錯誤 0；`CHK_Errors`＝0。
   - parity 期望值有意識地更新，只限 TK 連帶。
   - pytest 依檔分開全過。
   - 逐格比較 v0.1.8 → v0.1.9，範圍外差異列表。
   - η 各版本小表加 v21.10 欄。
5. **HTML 與 dist：**
   - 產出 `dist/<日期>_Anthropic收支模型_v0_1_9.html`／`.xlsx`（v0.1.8 dist 檔 git rm）。
   - 前面加 v0.1.9 摘要句。
   - 「資料與版本」改 TK v21.10。
   - Playwright 逐節讀畫面（桌機與手機寬度）。
6. **報告 `docs/reports/<日期>_v0.1.9.md`（繁體中文，白話）：**
   - ① TK 名稱新舊值與使用位置。
   - ② 結論表：v0.1.8 vs v0.1.9，格式同 v0.1.8 交接第 2 節。
   - ③ 拆解（步驟 2）。
   - ④ 敏感度（步驟 3）。
   - ⑤ 逐格比較表 2。
   - ⑥ 待 Project 判斷。
   - 另更新 CHANGELOG、`docs/handoff/Anthropic_handoff.md`、進度檔。
   - 開 PR（ready），PR 留言 `[Anthropic 回報] v0.1.9｜完成｜日期`。
   - **代理回覆中附報告全文**。
- **通知：** Wattonomics 的 `wattonomics/tools/import_gap_models.py` 若指向 v0.1.8 dist，**不要改**（屬 Wattonomics 串），只在報告 ⑥ 列出新路徑與 F07 新值。

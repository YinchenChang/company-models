# Anthropic 收支模型 交接檔（權威版；放 repo）

> 新對話串：先讀本檔、規格 `docs/plan/20261008_Anthropic_v0.1_規格.md` 與進度檔 `docs/reports/20261008_anthropic_進度.md`，再以 git 確認狀態。

## 1. 命題與範圍
- 命題（與 OpenAI v0.6 同一句）：Anthropic 每 VR 等值 GW 的年營收能否覆蓋每 GW 年全成本；若不能，缺口要多少外部資金、由誰以什麼條件提供。FY2025–FY2030，曆年制，2025 為實際校準年。
- 來由：Andy 2026-10-08「請比照 OpenAI，做 Anthropic」；依 2026-10-07 授權，Claude 套用預設（規格 D1–D23）不先確認，v0.1 完成時彙總審查。
- 分層：算力物理取自 Tokenomics（TK_Link 快照）；公司財務原始數據在 `SRC_ANT`；假設在 `Inputs`；OpenAI v0.6 命題輸出在 `OAI_Link`（只並排）。

## 2. 位置
- repo `YinchenChang/company-models`，資料夾 `anthropic/`；建置分支 `claude/anthropic-v0.1`，PR #22（base `main`；A5 完成後轉 ready，待 chat 端審查合併）。
- 規格 `docs/plan/20261008_Anthropic_v0.1_規格.md`（r1）；共同規則與 A1–A5 工作單 `docs/workorders/20261008_anthropic_*`。
- 成品：`dist/`；Drive `02-算力收支/Anthropic/`（chat 端放置）。

## 3. 模型狀態：**v0.1 完成**（2026-10-08；A0–A5）
| 項目 | 現況 |
|---|---|
| Excel | `model/20261008_Anthropic_v0.1.xlsx`（`model/CURRENT`；SHA-256 `127f3ea7…`）；成品副本 `dist/20261008_Anthropic收支模型_v0_1.xlsx`（位元組相同，測試檢查） |
| HTML 一頁摘要 | `dist/20261008_Anthropic收支模型_v0_1.html`（產生器 `tools/build_html.py`；敏感度情境 `tools/html_scenarios.yaml` 20 個；測試 `tests/test_html.py`；截圖 `docs/reports/img/20261008_v0.1_成品_*.png`） |
| 命題與驅動對照表 | `docs/plan/20261008_Anthropic_命題與驅動對照表.md` |
| Tokenomics | v5.26（`20261007_Tokenomics_v5.26.xlsx`），建置時 master `70d859e`（CURRENT 與 `3dd1216` 相同）；68 名＋NonNV 18 格＋PUE 3 格，check_tk_snapshot OK 89 |
| OpenAI 對照 | `OAI_Link`：`20261008_OpenAI_v0.6.xlsx`（SHA-256 `34fc34a7…`，提交 `e91df57`），17 名 |
| 規模 | 13 頁；SRC_ANT 379；Inputs 175；具名範圍 1,068；公式格 5,491；Checks C01–C105（CHK_Errors＝0、CHK_Warnings＝1：C89 AMD 逐家對帳，預期）；parity 情境 34 |
| 頁 | README、SRC_ANT、TK_Link、OAI_Link、Inputs、Demand、Revenue、Compute、Cost、Funding、Reverse、Checks、_Defaults |

**主要結果（基準）**：每 VR 等值 GW 差額（含股權報酬、現金口徑）2026–2030 −6.5／−5.2／−2.9／＋5.1／＋11.2（2025 −162 為分母效應）；覆蓋率 0.87／0.76／0.84／1.35／1.78。2026 已交割 $100B（Series G 30、Amazon 特別股 5、Series H 65）覆蓋 2026–2028 缺口，2026–2030 累計外部資金需求 **0**；只靠已到位融資的現金谷底 2028 年 80.2。OpenAI v0.6 同期 442.8。反向（管理層舊目標、同一支出）累計 25.2。翻轉需兩個需求驅動同時落在下緣（2030 差額 −1.7、外部資金 12.7）。

**歷程**：A0 骨架 `aa925dd` → A1 資料 `8d955cf` → A2 需求與營收 `fd5db5b`（v0.1-A2）→ A3 算力與成本 `c9d2478`（v0.1-A3）→ A4 融資與反向 `07565be`＋報告 `7842e3f`（＝v0.1）→ A5 成品（見 PR #22 完成留言）。報告 `docs/reports/20261008_v0.1-A1_資料.md`、`-A2.md`、`-A3.md`、`20261008_v0.1.md`（v0.1 完成報告，④A1–A4 預設彙總）、`20261008_v0.1_成品.md`。

## 4. 已定決定摘要（規格 D1–D23；r1 為 A1 審查後改寫）
D1 命題與命名同 OpenAI；D2 營收線方案別＋API 層級別；D3 廣告＝0；D4 Claude Code 不另立營收線；D5 r1 報導為總額、扣雲端平台抽成 16%（16–30%）後淨額進命題；D6 r1 產能一律取 Tokenomics，TPU／Trainium／AMD＝VR200 × NonNV 比例；D7 組合由逐合約 GW 加權；D8 η 以 2025 推論支出校準；D9 合約價 12 × 持有比；D10 r1 只計算力合約、金額 ↔ GW 互推、機房租約不另計；D11 自建 Fluidstack $50B 3 年均攤；D12 研發 GW 殘差；D13 非算力 2025 說明書、之後人數 × 每人成本，股權報酬單列；D14 2025 認列營收校準；D15 r1 2026 半校準；D16 r1 來源順序、條件式與 IPO 列情境；D17 最低現金＝次年非算力 6 個月；D18 回流對照只列不沖銷；D19 反向只作對照；D20 OAI_Link 只並排；D21 2025 $34B 非現金費用不計；D22 2026 Q2 營業利益只對照；D23 run-rate 里程碑只對照。

## 5. Andy 待審預設（v0.1 里程碑；全表 A2 24＋A3 30＋A4 21＝75 項與 D1–D23 見 `docs/reports/20261008_v0.1.md` ④、`20261008_v0.1_對照.xlsx` 第 ④ 頁；HTML 第 ⑧ 節）
1. **只計已簽合約**（規格未列的慣例，與 OpenAI 同）：使 2029–2030 轉正（供給 GW 2028 年 11.1 → 2030 年 10.4）；新增合約機制需 Andy 決定。上限：2029 新增支出 ≤ 39.7、2030 ≤ 86.5（當年差額），累計 ≤ 211.8（2030 現金餘裕）。
2. **API 任務數成長**（A2-13；與 OpenAI v0.6 同值）：最大不確定；2030 差額 2.0–31.0、現金谷底 41–117。
3. **2026 非算力成本可能低估約 $7.0B**（A4-17／D22）：補足後 2026 差額 −13.1、谷底 52.3，結論不變。
4. **合約價 12 對 VR 世代偏低**（D9）：雲端毛利率 2027–2030 −6% 至 −10%（兩家同一參數）。
5. **2026 已到位 $100B 的組成與避免重複**（A4-2～6）、**逐家錨定招股書**（A4-15；S4 谷底 52.3）。

## 6. 待辦（backlog；不在 v0.1 範圍，需 Andy 或 chat 端決定）
1. **新增合約機制**（最重要）：2029 起新簽算力與營收成長或 VR 等值 GW 目標連動；兩家同步。
2. **合約價分世代**：改為 TK `IF_HoldEcon` ×（1＋雲端毛利率 Inputs），兩家同步。
3. **公開版 S-1**：上 EDGAR 後把招股書轉述數字改 Verified、以 2026 Q1–Q2 實際營業費用校準 D22、補逐年付款排程與股權報酬、2026 年中現金。
4. **AMD 合約**：「>$20B」與「最多 2 GW」不相容（Checks C89 WARN），決定以金額或 GW 為準。
5. Fluidstack 自建、Hut 8／Cipher 機房與 Broadcom TPU 的對應（重複計算風險）。
6. Tokenomics 缺口：NonNV 比例與 PUE 改具名範圍；NonNV 表補 TPU v6e、Trainium2、MI355X、RTX PRO 6000；研發 GW 外部錨點。
7. 敏感度若要在 Excel 內直接看：v0.2 以 Inputs 情境選擇器＋結果表實作（目前由 yaml＋重算保證一致，同 OpenAI v0.6）。
8. GitHub Actions 因帳號付款／用量上限未啟動；目前以本地 pytest 驗收。

# Anthropic v0.1.7 工作單：正式 S-1 全面套用＋自建資本支出改 GW 法＋快照同步（2026-10-31，chat 端）

**背景：** v0.1.6 只把正式 S-1 的 2026 上半年營收換進基準，其餘正式數字（承諾、2025 校準、季度營業損益）只進對照列。Andy 10-31 13:06 問「Anthropic 底層模型是不是已更新完成、S-1 正式版有沒有整個更新進去」；chat 端回報未完整套用並提出判斷題，Andy 13:46「同意，請繼續」。本版把正式 S-1 套進基準。

**Andy 已定（原話）：**
- 10:22（Wattonomics 串，同一原則）：Anthropic 自建資本支出「以自建 GW 增量 × TK 每 GW 建置成本為基準，不用 Fluidstack $50B 公告金額；S-1 資本支出只作對照」。
- 13:46「同意，請繼續」＝同意 chat 端預設：S-1 揭露的正式承諾（含 6-30 後新簽 245）**進基準**，並列「不進基準」（＝v0.1.6 合約口徑）讀法。

**原則：** 共同規則、CLAUDE.md 照舊（Excel 為唯一事實來源、Python 不旁路、ID 穩定、舊列不刪）。每一項都要有 Inputs 開關或情境，可回到 v0.1.6 口徑，並在報告列「只動該項」的前後值。

## 工作項目
- **W1 承諾進基準（合約供給與成本）：**
  - (a) 逐家改讀正式 S-1：Google ≥114.5（2026-04～2033-07）、Amazon ≥110.0（2026-05～2036-04）、Microsoft ≥44.9（2026-11～2033-05）、SpaceXAI ≤120.5（～2029）、Broadcom 161.2（2026–2031）、Nscale ≤44.6、Lambda 約 41.3、AMD >20（2027–）。付款→GW 用既有方法（Compute 第三、四節、Cost 第十一、十二節），期間以 S-1 期間為準；上限型（SpaceXAI、Nscale）沿用模型對上限型的既有處理。
  - (b) 未歸屬餘額：S-1 三段合計（321.4＋265.6＋245＝832）− 模型逐家合約總額（W1(a) 後）＝「S-1 承諾未歸屬餘額」，以通用合約價（INP_114）換 GW，付款期間預設 2027–2033 平均攤提［Assumed；區間 2027–2031（快）／2028–2036（慢）放 Sens］。三段與 Note 16 的 417 重疊，不相加。
  - (c) Google 額外 TPU 125–150（不在承諾表）不進基準，只作情境。
  - (d) Inputs 開關 `s1f_commit_switch`：0＝v0.1.6 合約口徑。報告並列兩讀的命題輸出（每 VR 等值 GW 差額、覆蓋率、轉正年、累計外部資金需求、現金谷底）。
- **W2 自建資本支出改 GW 法：** 自建資本支出＝自建 GW 增量 ×（TK `IF_CapexTotal`，廠房依既有 lag 提前，與 Wattonomics Gap 頁同一套）。自建 GW 以 Wattonomics v0.1.7 為權威（`wattonomics/data/wattonomics_snapshot_v0.1.7.json`：`WIF_Gap_Player_ant_<yyyy>_CapexAIGW` 目前各年 0，即 Wattonomics 無 Anthropic 自有 GW 證據）；若查到 Anthropic 自有站點 GW 證據，列出但只進對照、在報告提請 chat 端轉交 Wattonomics。Fluidstack 站點視為 Fluidstack 擁有、經承諾租回（已在 W1 的 S-1 承諾內），不再另計資本支出（避免重複計算）。對照列：S-1 2026H1 資本支出 0.482（年化 0.96）vs 模型 2026 資本支出。舊法（$50B 三年均攤＝16.67／年）保留為開關 `own_capex_method`（0＝舊法）。
- **W3 2025 校準改正式值：** 2025 營收 4.6 → 4.585（SRC_ANT_412）、算力支出 7.33 → 7.326（SRC_ANT_424）；η 隨之重校。
- **W4 非算力成本（D22）改用正式季度營業損益校準：** 用 S-1 季度營業損益（2025Q1–2026Q2；2026Q2 +0.974）與 2026H1 營收、算力費用，回推 2026 非算力成本（營收 − 算力 − 營業損益 − 股權報酬口徑，報告寫明口徑）；取代 v0.1.1 的報導推估。若推得值與現值差 >30%，先報告、照推得值入基準並列舊值。
- **W5 快照同步：** OAI_Link 改 OpenAI v0.6.4（`openai/model/20261029_OpenAI_v0.6.4.xlsx`）；TK_Link 改 Tokenomics v21.7（clone 於 `/home/claude/yinchenchang/tokenomics`，master `90f7deb`，`model/20261025_Tokenomics_v21.7.xlsx`）。列 TK 名稱變動表（只列實際讀取者）。
- **W6 Amazon 實際投入（10-Q 發現）：** Amazon 2026Q3 10-Q 顯示對 Anthropic 2026 實際投入 10.0（Wattonomics 證據 F#81；模型 Funding 若以 18.0 計入已到位融資，改 10.0 並列舊值）。查 `wattonomics/evidence_raw/Q_review_2026q3.md` 與 Funding 頁現行組成。
- **W7 成品：** HTML 沿用 #282 版面（① 後直接接 ②；各版更新紀錄在頁尾），新增「①之五 v0.1.7 更新了什麼」放頁尾紀錄區最前；① 的一句話結論依新結果改寫（條件式）。頁首說明改為「正式 S-1 已全面套用」，並寫明仍為二手轉載的列。報告 `docs/reports/20261031_v0.1.7.md`、對照 Excel 照共同規則；報告第一節回答「全面套用正式 S-1 後結論是否翻轉」，附 W1–W6 瀑布（從 v0.1.6 逐項加入，各步前後值）。

## 預設（未回應照此；報告「已套用的預設」逐項列出）
| 題 | 預設 |
|---|---|
| 未歸屬餘額付款期 | 2027–2033 平均（Assumed；快／慢兩讀進 Sens） |
| 上限型合約（SpaceXAI ≤120.5、Nscale ≤44.6） | 沿用模型對上限型的既有取值規則 |
| Google 額外 TPU 125–150 | 不進基準 |
| 自建 GW 無證據 | 0（未找到≠不存在：報告列為缺口） |
| 仍為二手轉載的 S-1 列 | 維持 Interested-party；不改值 |
| 非算力成本推得值與現值差 >30% | 照推得值入基準，報告首節說明 |

## 停止條件
共同規則第 6 節。另加：任何一項使 CHK_Errors ≠ 0 且無法以工程方式修正時，停下報告。

## 驗收（chat 端）
parity 全套、`tests/test_html.py`、`tests/parity/test_builder.py` 全過；`tools/sens_coverage/check_sens_coverage.py <成品> --require-keycol`、`tools/in_layout/check_in_layout.py`（若適用）跑過；未參與代理獨立核對瀑布與首節數字。

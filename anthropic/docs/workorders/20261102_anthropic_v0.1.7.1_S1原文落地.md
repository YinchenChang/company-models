# Anthropic v0.1.7.1 工作單：S-1 原文核對落地＋已到位融資更正（2026-11-02，chat 端）

**背景：** 11-02 S-1 原文核對（`docs/reports/20261102_S1原文核對.md`，PR #293 已合併）後，chat 端續查 v0.1.8 工作單 R3，發現一項會翻轉命題 2 的事實錯誤：**2026 上半年股權募資原文為 110.4（現金流量表「特別股發行收入」110,432 百萬），模型「已到位融資」只算 100。** 原文拆分（Note 10、Note 4、Note 9）：
- Q1：Series G-1／G-2 共 30.2（每股 259.1364）。
- Q2：Series G-2 共 **15.0**＝Google 38.6 百萬股（2026-03 Google 2026 NPA 承諾，約 10.0）＋Amazon 19.3 百萬股（2026-04 Amazon 2026 NPA 承諾，約 5.0）。
- Q2：Series H-1／H-2 共 65.2（每股 589.0095）；Note 9：2026-05 Google 另投 10.0、Amazon 另投 5.0 於 Series H（使兩家 2026 NPA 剩餘額度降為 20.0、15.0）。
- 模型：F12 Series G 30、F13 Amazon Series G 5（＝上面 G-2 的 Amazon 5）、F14 Series H 65；**F17 Google 10 的開關 `googl_upfront_switch`＝0，理由是「視為 Series H 15 的一部分」——原文顯示 Google 的 G-2 10.0 與 Series H 分開，故少算 10。**

粗估：2029 現金谷底 11.86 → 約 21.9（最低現金 13.05）→ 累計外部資金需求 1.19 → 0。

v0.1.8（S-1/A）工作單照舊；本版先把可由現行原文確定的項目落地，S-1/A 出來後 v0.1.8 以本版為底。

**原則：** 共同規則、CLAUDE.md 照舊（第 8 節替換程序：新增 SRC 列、舊列標取代、不刪不覆寫；ID 穩定；Python 不旁路）。每項有開關或可回 v0.1.7。

## 工作項目
- **V1 來源等級（＝v0.1.8 R1）：** `data/anthropic_src.yaml` 中原文核對報告第 1 節判「一致」的 46 列：tag 改 Verified、party 寫「Anthropic（正式 S-1；期中未經查核）」、url 改 `https://www.sec.gov/Archives/edgar/data/2130357/000162828026067168/anthropicpbc-sx1.htm`、retrieved 2026-11-02、備註加原文位置（頁碼或附註）與「2026-11-02 chat 端瀏覽器核對」。SRC_ANT_421 備註改「原文表內 11,555、正文 11.6；本列為轉載者捨入，只對照」（tag 不改）。附註出處更正：446「Note 16」→ Note 13（F-59）；445 出處改 Note 7（F-36）；459 頁碼 56 → 57。427–431 值改原文精度 0.944／1.390／2.047／2.945／3.387（只對照列；建置後確認命題輸出不變）。逐列位置以核對報告第 1 節表格為準。
- **V2 已到位融資更正（基準變動）：**
  - 新增 SRC 列（Verified，正式 S-1）：(a) 2026 上半年特別股發行收入（淨額）110.432（合併現金流量表，F-13）；(b) Q1 Series G 30.2（Note 10）；(c) Q2 Series G-2 15.0，其中 Google 38.6 百萬股、Amazon 19.3 百萬股，每股 259.1364（Note 10、Note 4）；(d) Q2 Series H 65.2（Note 10）；(e) Google 2026-05 投資 Series H 10.0、Amazon 5.0（Note 9）。
  - `googl_upfront_switch` 0 → **1**；依據改為新 SRC 列（c）；名稱改「Google 2026 Q2 Series G-2 $10B（S-1 Note 4、10：與 Series H 分開；1＝計入）」。舊依據列 SRC_ANT_328 備註加「已由 SRC_ANT_nnn 證實為另一筆（2026-11-02）」。Funding F14、F15、F17 的標籤與說明依原文改寫（Series H 65.2 含 Google 10、Amazon 5；G-2 15.0 另計）。`fund_series_h_amount` 是否改 65.2、Series G 是否改 30.2：改用原文值（新 SRC 列），報告列前後值。
  - 新增 Checks（INFO 或 WARN）：FND_Committed 2026 − S-1 特別股收入 110.432，|差| ≤ 0.5 為 OK。
  - Funding 第十三節 6-30 現金對照：保留，新增一列「其中：Broadcom 租約保證金（受限現金）約 2.5（2026-04 存入；S-1 Note 16）」只對照；報告寫明差額拆解（股權收入、上半年自由現金流時點、受限現金）。
- **V3 條件式融資額度（只動情境 S2／S3，不動基準）：** Google 條件式 30 → 20（S-1 MD&A：2026-06-30 剩餘可發行 20.0；另有 IPO 同時認購最多 10.0 的選擇權，行使則 NPA 等額減少，不相加）；Amazon 15 維持（原文 15.0）。新增 SRC 列、舊列標取代。
- **V4 最低現金依據（只對照）：** 新增 SRC 列：A&R 循環額度 15.0（2031-09 到期）、最低流動性條款 5.0（自 2026-12-31 季末起）（MD&A 第 163–164 頁；Note 17）。Funding 第四節說明加一句：模型最低現金（D17，次年非算力成本 × 6 個月，Assumed 3–12）高於契約下限 5.0；循環額度照 D16 只列或有。報告列 `min_cash_months` 3／6／12 三點的累計外部資金需求與谷底。
- **V5 報告文字更正（＝v0.1.8 R2）：** v0.1.7 報告與 repo 交接中「上半年股權報酬未單獨揭露」「約 0.111 未說明」改為「股權報酬 2.185＋非現金捐贈 0.126（2026Q1）＝2.311」。
- **V6 成品：** HTML 沿用 #282 版面；頁尾紀錄區最前加「①之六 v0.1.7.1 更新了什麼」；① 的一句話結論依新結果改寫（條件式，寫明最低現金假設）。報告 `docs/reports/20261102_v0.1.7.1.md`、對照 Excel；CHANGELOG；`model/CURRENT`；dist；repo 交接檔第 2b／3 節。

## 預設（未回應照此）
| 題 | 預設 |
|---|---|
| Google G-2 10 | 計入基準（原文事實；舊讀法 0 保留為開關，報告並列） |
| Series G／H 金額 | 改原文 30.2／65.2（差 +0.4；報告列影響） |
| 受限現金 2.5 | 不另計（Broadcom 租金保證金，之後抵付租金）；只對照 |
| N1 SpaceXAI、N2 期後短期 21、N3 有條件融資、既有六題 | 照舊 |

## 停止條件
共同規則第 6 節；CHK_Errors ≠ 0 且無法以工程方式修正時停下報告。範圍外變動必須為 0（V1、V3、V4、V5 不得動基準輸出）。

## 驗收（chat 端）
parity 全套（分片 `PARITY_SHARD`／`PARITY_LOOP_SHARD`；CI 增量重算 2 秒門檻偶發超時屬計時，非數值）、`tests/test_html.py`、`tests/parity/test_builder.py` 全過；`tools/sens_coverage/check_sens_coverage.py <成品> --require-keycol`；瀑布：V2 單獨的前後值；未參與代理獨立核對 V2 數字與命題輸出。

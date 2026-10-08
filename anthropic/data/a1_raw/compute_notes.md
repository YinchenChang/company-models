# A1-C 算力合約與算力支出：研究筆記（規格第 4 節 D）

- 截止：2026-10-08；檢索日 2026-10-08；共 91 筆（`compute.yaml`），not_found 11 項。
- 標記原則：公司新聞稿、合作方新聞稿、外流招股書／投資人資料一律 Interested-party；Broadcom 8-K、SpaceX IPO 申報為 Verified；Epoch AI、Measured AI 等第三方估計為 Analogy；自行相減或倍數換算為 Derived。
- **結論先講**
  1. 招股書草稿（Reuters 2026-09-28 取得，尚未公開於 EDGAR）是目前最完整的口徑：六家合作方合計約 **$518B**，約 80% 不可取消。分別為 Google 111.1、Amazon 110、Microsoft 31.4、Broadcom 161.2、xAI/SpaceX ≤84.5、AMD >20（單位 $B）。
  2. 2025 實際「算力與基礎設施」支出為 **$7.33B**（招股書草稿）。其中訓練約 $4.1B（The Information 引管理層預測），推論等其餘約 $3.2B 為推導值。
  3. 容量目標：2025 年底約 1.4 GW；2026 年底約 5 GW（NYT 引投資人簡報）；2027 年底約 10 GW（「再倍增」）。2025-10 以來新簽合約約 14.8 GW 以上（The Information）。
  4. 口徑問題嚴重：除 TeraWulf 明示「critical IT」、New Carlisle 2.2 GW 為用電之外，其餘 GW 數字都沒有說明是 IT 還是設施口徑。另外多數數字是「最高」（up to），不代表已承諾。

## 1. 合約總表

| # | 對象 | 公布日 | 金額 | 容量（GW／MW） | 晶片 | 期間／起始 | GW 口徑 | 主要來源 |
|---|---|---|---|---|---|---|---|---|
| 1 | Google Cloud（TPU） | 2025-10-23 | 「tens of billions」 | 2026 年上線遠超過 1 GW；TPU 最多 100 萬顆 | TPU v7 Ironwood | 2026 起；期間未揭露 | 未明 | Anthropic／Google 新聞稿 |
| 2 | Broadcom（Google TPU 硬體） | 2025-12-15 | 首筆 $10B＋追加 $11B（2026 交貨） | — | Ironwood | 2026 | — | Broadcom 法說（DCD） |
| 3 | Google／Broadcom 次世代 TPU | 2026-04-06 | 8-K 未給金額；The Information 稱 $200B／5 年；招股書：Google ≥$111.1B（2026-04～2033-07）＋Broadcom 設備租賃 $161.2B（其中五年 TPU 租賃 $125.2B） | 8-K：約 3.5 GW；Anthropic 2026-05 稱 5 GW | 次世代 TPU | 2027 起；Broadcom 對 Google 供應至 2031 | 未明 | Anthropic 新聞稿、Broadcom 8-K、Reuters |
| 4 | AWS／Amazon（Project Rainier） | 2025-10-29 | AWS 園區資本支出 $11B（AWS 出資） | 約 50 萬顆 Trn2，2025 年底 >100 萬顆；園區用電最高 2.2 GW | Trainium2 | 2025-10 啟用 | 設施（2.2 GW）；Epoch 估 IT 910 MW | AWS（Converge Digest 轉述）、Epoch |
| 5 | AWS／Amazon 擴大協議 | 2026-04-20 | 十年 >$100B；招股書 $110B（2026-05～2036-04） | 最多 5 GW；2026 年底新增近 1 GW | Trn2、Trn3、Trn4 | 十年 | 未明 | Anthropic 新聞稿 |
| 6 | Microsoft Azure＋NVIDIA | 2025-11-18 | $30B；招股書 $31.4B（2026-11～2033-05） | 最多 1 GW | Grace Blackwell、Vera Rubin | 未揭露 | 未明 | Anthropic 新聞稿 |
| 7 | Fluidstack 自建 | 2025-11-12 | $50B | 未揭露（DCD 推測單站 168–360 MW） | （第三方：Google TPU） | 2026 年內陸續上線 | 未明 | Anthropic 新聞稿 |
| 7a | Hut 8 River Bend（LA，經 Fluidstack） | 2025-12-17 | $7B（15 年；含選擇權 $17.7B） | 245 MW；Hut 8 為 Anthropic 開發最多 2.295 GW | — | 2027Q2 起 | 未明 | DCD |
| 7b | Cipher Barber Lake（TX，經 Fluidstack） | —（Epoch 歸屬） | — | IT 168→207 MW | TPU v7 約 12.4 萬顆 | 2026-12／2027-04 | IT | Epoch（Analogy） |
| 8 | CoreWeave | 2026-04-10 | 未揭露 | 未揭露 | 未揭露 | 2026 下半年起；多年 | — | CoreWeave 新聞稿 |
| 9 | Nexus Hubbard（TX，Google 擔保） | 2026-03-30 | 融資最高 $5B（FT），其後約 $15B（WSJ） | 首期 500 MW；遠期約 7.7 GW；自建燃氣電廠 1.6 GW | — | 2026 年內首期 | 未明 | DCD、Global DC Hub |
| 10 | SpaceX／xAI Colossus | 2026-05-06 | $1.25B／月至 2029-05（S-1）；招股書上限 $84.5B | Colossus 1 >300 MW、>22 萬顆 GPU | NVIDIA | 2026-05～2029-05；任一方 90 天可退 | 未明 | Anthropic 新聞稿、SpaceX S-1、Reuters |
| 11 | Akamai | 2026-05-08 | $1.8B／7 年 | 未揭露 | RTX PRO 6000 Blackwell | 2026–2033 | — | Akamai 財報＋Bloomberg 指認 |
| 12 | TeraWulf Hawesville（KY，直接承租） | 2026-07-06／08-05 | 約 $19B／20 年 | 401 MW | （第三方：Google TPU） | 2027H2 起、2028 初全數 | **IT** | TeraWulf 新聞稿 |
| 13 | AMD | 2026-07-22 | 招股書 >$20B；AMD 另投資最多 $5B | 最多 2 GW | MI450（MI455X／Helios） | 首 1 GW 2027H1 起 | 未明 | AMD 新聞稿、Reuters |
| 14 | Volta Infra（挪威 Tydal，Bitdeer） | 2026-08-05 | $10B／6 年 | 133 MW | Vera Rubin | 至 2027-03 分兩期交付 | 未明 | Bloomberg（轉述） |
| 15 | Riot Rockdale（TX） | 2026-08-11 | $9.1B／20 年（含續約 $16.5B） | 191 MW（2027-12 先上 96 MW） | （第三方：AMD MI450） | 2027-12～2028-06 | 未明 | DCD 引 Bloomberg |
| 16 | Nscale Monarch（WV） | 2026-08-26 | $45B／6 年 | 460 MW | Vera Rubin | 2027 年底起 | 未明 | DCD 引 Bloomberg |
| 17 | Lambda（Hut 8 Beacon Point，TX） | 2026-09-01 | $35B（多年） | 350 MW（NVIDIA 承租） | NVIDIA | 未揭露 | 未明 | VKTR 引 WSJ |

說明：7a、7b、12、15 屬機房租賃（不含晶片），硬體多來自 Google TPU（Broadcom 租賃）或 AMD。因此在金額上可能與 #3、#13 重疊，不可直接加總。

## 2. 時間線

- 2024–2025：AWS 為主要訓練夥伴。外流帳單顯示 2024 年付 AWS $1.36B，2025 年 1–9 月付 $2.66B（2025-09 單月 $519M）。
- 2025-10-23：Google TPU 合約（最多 100 萬顆 Ironwood；2026 年 >1 GW）。
- 2025-10-29：Project Rainier 啟用（約 50 萬顆 Trn2；園區 2.2 GW；年底 >100 萬顆）。
- 2025-11-12：與 Fluidstack 宣布 $50B 美國自建（TX、NY；2026 年上線）。
- 2025-11-18：Azure $30B＋NVIDIA 最多 1 GW。
- 2025-12-15：Broadcom 揭露 Anthropic 的 $10B＋$11B TPU 訂單。
- 2025-12-17：Hut 8 River Bend 245 MW（最多 2.295 GW）。
- 2026-01：The Information 報導預測：2026 年訓練 $12B、推論 $7B；2029 年前訓練 >$100B（或「至多」$100B，兩種報導相反）。
- 2026-02-18：預測 2029 年前付三大雲推論費用 ≥$80B。
- 2026-03：Nexus Hubbard 租約（Google 擔保）。
- 2026-04-06：Google／Broadcom 多 GW（8-K：3.5 GW，2027 起）。
- 2026-04-10：CoreWeave。
- 2026-04-20：Amazon 最多 5 GW、十年 >$100B。
- 2026-05-06：SpaceX Colossus 1（>300 MW）；SpaceX S-1 揭露 $1.25B／月。
- 2026-05-08：Akamai $1.8B。
- 2026-07：TeraWulf 401 MW IT；AMD 最多 2 GW。
- 2026-08：Volta $10B、Riot $9.1B、Nscale $45B；Hubbard 融資 $15B。
- 2026-09-01：Lambda $35B。
- 2026-09-06：The Information 彙整「≥14.8 GW、最多 $517B」。
- 2026-09-18：NYT 報導 2026 年底 5 GW、2027 年底再倍增。
- 2026-09-28／29：Reuters 取得招股書草稿（$518B，80% 不可取消；2025 算力 $7.33B）。
- 2026-10-01：Broadcom 最多 $42B 可轉債融資；五年 TPU 租賃 $125.2B。

## 3. 逐年 GW（可推導範圍）

| 年底 | 公司／投資人口徑 | 第三方 | 已公布新增來源（上限值，口徑混雜） |
|---|---|---|---|
| 2025 | 約 1.4 GW（媒體轉述） | Measured AI 1.4 | Rainier 約 0.9 GW IT（Epoch） |
| 2026 | 約 5 GW（NYT） | Measured AI 5.0 | Google >1、AWS 新增近 1、SpaceX >0.3、MSFT/NVDA ≤1、Nexus 0.5、Fluidstack 站點 |
| 2027 | 約 10 GW（NYT「倍增」，Derived） | Measured AI 9.5 | Google/Broadcom 3.5–5、AMD 1、TeraWulf 0.4、Nscale 0.46、Hut 8 0.245、Riot 0.096、Volta 0.133 |
| 2028 | — | Measured AI >15 | AWS 至 5 GW（十年內）、AMD 第 2 GW、Riot／TeraWulf 全數 |

## 4. 矛盾點

1. **Google／Broadcom 容量**：Broadcom 8-K 寫約 3.5 GW；Anthropic 2026-05 新聞稿寫「5 GW agreement」。差額 1.5 GW 可能是把 2026 年的 1 GW 算進去，原因未說明。
2. **Google 金額**：The Information 稱 $200B／5 年（Google＋Broadcom）；招股書寫 Google ≥$111.1B 加 Broadcom 設備租賃 $161.2B，合計 $272B。The Information 的 $200B 可能只是其中一部分。
3. **SpaceX**：S-1 寫月費到 2029-05，約 $45B；招股書寫上限 $84.5B，約為兩倍，顯示合約已擴大。Musk 則稱是「180-day lease」，90 天可退。
4. **訓練成本**：The Information（2026-01）標題稱 2029 年前「exceed $100B」，Tiger 轉述同一來源為「as high as $100B」；WSJ 外流文件稱訓練成本 2028 年高峰約 $30B／年。年度或累計口徑不明。
5. **2026 推論 $7B（2026-01 預測）**：這個數字明顯低於目前的合約節奏。單是 SpaceX 一年就約 $15B，應視為過時。
6. **AWS 晶片**：The Information 表列「5 GW (Trainium3)」，公司新聞稿則說 Trn2 到 Trn4 都有。
7. **$517B 與 $518B**：前者是 The Information 依公開與知情人士報導彙整（含 Nscale、Lambda、Fluidstack 等），後者是招股書的六家合作方。兩者組成不同，數字相近是巧合。招股書沒有單列 Fluidstack、Nscale、Lambda、Volta、Riot、TeraWulf，這些可能歸在 Google／Broadcom 或其他科目之下，也可能尚未計入。

## 5. 缺口（詳見 yaml `not_found`）

- 晶片占比（TPU／Trainium／GPU）：公司只說用三個平台、Amazon 是主要訓練夥伴，沒有給比例。判斷為可能未揭露。
- 招股書逐年付款排程、2027–2030 逐年算力支出：找不到。
- 2025 推論成本金額：找不到，只能推導約 $3.2B。
- 主要合約的 GW 是 IT 還是設施口徑：可能未揭露。建模時建議統一視為「未明」，並以 Inputs 的 PUE 給區間。
- CoreWeave 金額與 MW、Fluidstack $50B 的 GW、Trainium3 顆數：未揭露或找不到。
- 無法開啟的頁面：aboutamazon.com、blogs.microsoft.com（需授權）、CNBC（403）、Broadcom 8-K PDF（二進位）。這幾個來源改用 Anthropic 新聞稿或轉載替代。

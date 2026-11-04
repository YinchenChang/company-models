---
ticker: Anthropic
name: Anthropic
name_full: Anthropic PBC
listing: 未上市
hq: 美國加州舊金山 [外部・未經本專案檢核]
founded: 2021 [外部・未經本專案檢核]
employees: 約 3,500（2026-06，公司口述）[本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E299]
fiscal_year_end: 12 月底（本專案模型以曆年計）
asof: 2026-11-04
---

## 一句話定位
AI 模型公司（位於 AI 金流的最上游）：自研 Claude 系列大模型，收入以企業與開發者按 token 付費的 API 為主（FY2025 82.9%，模型推估 FY2026 91.0%）[本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E30、E416；Revenue!E58、E67；Derived]，算力向 Google、Amazon、Microsoft 等租用；本專案模型（已納入公開 S-1 與 S-1/A 的算力承諾）推估自由現金流 2026 年為正、2027–2029 年為負、2030 年起持續轉正，只靠已到位融資不需再募資 [本專案：Anthropic收支模型_v0_1_8.xlsx Funding!E14–I14、D254、I43]，但若 SpaceXAI 合約照 S-1/A 月費付滿上限，需要外部資金 22.0bn [本專案：anthropic/docs/reports/20261103_v0.1.8.md ⑤]。

## 公司在做什麼
- 模型與 API：Claude 分 Opus（頂層）、Sonnet（中層）、Haiku（低層）三個價位，按「每百萬 token」計價（token 是模型處理文字的最小單位）；2026-09 Opus 5.5 牌價輸入 4、輸出 20 美元，2026-10 Haiku 5.5 降至 0.1／0.5 美元 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E115–E116、E119–E120]。
- 訂閱：個人 Pro 每月 20 美元、Max 100／200 美元，另有 Team 與 Enterprise 企業席位 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E145、E147–E148、E150–E156]。
- 程式代理 Claude Code（讓 AI 直接讀寫程式碼的工具）：公司稱 2026-02 年化營收 2.5bn；第三方估 2026-05 約 8bn [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E62、E64]（[Interested-party]／[Analogy]）。本專案模型不另立營收線，併入訂閱與 API [本專案：anthropic/docs/handoff/Anthropic_handoff.md §4 D4]。
- 客戶以企業為主：2025-10 商業客戶約 300,000 家、企業與商業客戶約占營收 80%；2026-04 年化逾 1M 美元客戶約 1,000 家 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E166、E40、E171]（[Interested-party]）；公開 S-1：2026-09-30 近十二個月營收逾 10 萬美元的客戶逾 11,900 家、淨營收留存率約 900% [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E468、E467]（[Verified]）。
- 營收規模：2025 年認列 4.585bn；2026 上半年 16.286bn、Q3 16.956bn [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E416、E418、E492]（公開 S-1 與 S-1/A，[Verified]，期中數字未經查核）。

## 主要產品與營收占比
| 產品／部門 | 內容（白話） | 營收占比 | 期間 | 標記 |
|---|---|---|---|---|
| API（用量計價） | 開發者、企業與雲端市集按 token 付費 | 82.9%（3.8 ÷ 4.585bn） | FY2025 | [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E30、E416]（[Derived]） |
| 訂閱 | Pro／Max 個人方案＋Team／Enterprise 席位 | 17.2%（0.789 ÷ 4.585bn） | FY2025 | [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E29、E416]（[Derived]） |
| API（用量計價） | 同上 | 91.0%（45.67 ÷ 50.2bn） | FY2026（模型推估） | [本專案：Anthropic收支模型_v0_1_8.xlsx Revenue!E58、E67]（[Derived]） |
| 訂閱 | 同上 | 9.0%（4.53 ÷ 50.2bn） | FY2026（模型推估） | [本專案：Anthropic收支模型_v0_1_8.xlsx Revenue!E55、E67]（[Derived]） |

- 2025 年 API 與訂閱拆分來自外流招股書草稿，公開 S-1 未拆；2026 年 API 是由總額倒推的唯一項，不是公司揭露 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E30；Revenue!E57]；第三方估 2026 API 72.5%、訂閱 12.5% [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E65–E66]（[Analogy]）。
- 雲端通路：2025 年 47% 營收經 Amazon／Google 雲端市集銷售（草稿轉述），公開 S-1 寫 2026 上半年約 48% [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E31、E466]；模型以總額認列後再扣平台抽成 16%（2026 抽成 4.14bn、淨額 46.05bn）[本專案：Anthropic收支模型_v0_1_8.xlsx Revenue!E70–E72]。
- 客戶集中：公開 S-1 寫 2026 前九個月前五大客戶約占 25%、沒有單一客戶超過 10%；上半年最大揭露客戶（Customer C）占 11% [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E463、G464、E465]（[Verified]）。
- 獲利：2026 Q2 首度 GAAP 營業獲利 0.974bn；Q3 營業損失 0.567bn，主因股權報酬 1.07bn 與一次性非現金捐贈配對 1.96bn，排除後調整後營業利益率 15% [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E445、E494–E496、E498]（S-1／S-1/A，[Verified]）；草稿轉述 2025 淨損 42bn（其中 34bn 為非現金重估）、2026 上半年淨損 80.7bn（約 72bn 非現金）[本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E284–E285、E395–E396]（[Interested-party]）。

## 市場定位與產業位置
- 在 AI 金流最上游：模型 2026 營收淨額 46.05bn、算力成本（現金）28.18bn [本專案：Anthropic收支模型_v0_1_8.xlsx Funding!E8–E9]；公司稱 Q3 每 1 美元算力支出帶來 1.7 美元營收 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E391]（[Interested-party]），以 S-1/A Q3 營收 16.956bn ÷ 算力與基礎設施費用 9.863bn 計約 1.72 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E492–E493]（[Derived]）。
- 每 VR 等值 GW（把各種晶片換算成 NVIDIA VR200 的算力單位）：2026 營收淨額 32.28bn、全成本 33.22bn，覆蓋率 0.97；2030 升至 1.20 [本專案：Anthropic收支模型_v0_1_8.xlsx Cost!E109、E113、E117、I117]。本模型沒有 OpenAI 那樣的 P×Q 回本比值頁：本專案未建立。
- 晶片最分散的模型商：模型 2030 供給 12.94 GW 中，Google TPU v7／次世代 5.43 GW、Trainium3 1.58、AMD 1.6、NVIDIA 合計約 4.33（其中 2.79 GW 是 S-1 承諾中尚未對應到具名合約的餘額，以 VR200 族換算）[本專案：Anthropic收支模型_v0_1_8.xlsx Compute!I91–I100、I255]（[Derived] 加總）。
- 算力承諾（公開 S-1/A，截至 2026-09-30）：總額約 895.4bn，含表內義務 320.0bn、已簽未起租租賃 451.2bn（加權租期約 10 年）、9-30 後新簽 100.2bn、短期承諾 17.4＋6.6bn；其中 213.8bn 附終止權或開工前條件 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E499–E500、E503–E505、E507–E508、E510–E511]（[Verified]）。表內義務付款：2026 Q4 5.2、2027 29.2、2028 44.4、2029 47.8、2030 45.1bn、之後 148.4bn；未來 12 個月預計支付 69.8bn [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E501–E506、E513]。
- 逐家合約（S-1／S-1/A）：Google Cloud 121.3bn、Broadcom TPU 設備租賃 161.2bn（約 3.5 GW）、Amazon 110bn、Microsoft Azure 44.9bn、Nscale 最高 44.6bn、Lambda 41.3bn、SpaceXAI 最高 120.5bn（滿載每月約 3.1bn）[本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E517、E458、E192、E455、E518、E459–E460、E457、E519]；模型基準對 SpaceX 仍用每月 1.25bn、合計 45bn [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E231；Cost!D143]。模型把 S-1/A 三段承諾 871.4bn 扣掉逐家合約 580.1bn，餘額 291.3bn 依 S-1/A 付款結構計入成本 [本專案：Anthropic收支模型_v0_1_8.xlsx Cost!D442、D368、D370]。
- 投資人兼供應商：Google、Amazon、Microsoft、NVIDIA、AMD 同時是股東與算力賣方 [本專案：Anthropic收支模型_v0_1_8.xlsx Funding!B20–B34]；本專案推估 Anthropic 約占 Google Cloud backlog 22%（下限）[本專案：alphabet_facts_20261008.json rp.anthropic.rpoShare]。

## 管理層與經營風格
| 職位 | 姓名 | 任職起 | 背景 |
|---|---|---|---|
| 執行長（共同創辦人） | Dario Amodei | 2021 | 前 OpenAI 研究副總 [外部・未經本專案檢核] |
| 總裁兼董事長（共同創辦人） | Daniela Amodei | 2021 | 前 OpenAI 安全與政策副總 [外部・未經本專案檢核] |
| 財務長 | Krishna Rao | 2024 | 前 Fanatics Commerce 財務長 [外部・未經本專案檢核] |
| 技術長 | Rahul Patil | 2025 | 前 Stripe 技術長 [外部・未經本專案檢核] |
| 首席科學官（共同創辦人） | Jared Kaplan | 2021 | [外部・未經本專案檢核] |

- 經營風格：本專案尚未研究。
- 股權與治理（本專案可見部分）：共同創辦人經 Founder LLC 掌握 50.1% 投票權 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E315]。
- 資本募集：公開 S-1 列 2026 年 Series G-1／G-2 30.2bn、Series H-1／H-2 65.2bn，另 Google 與 Amazon 於 Q2 認購 Series G-2 共 15.0bn；上半年特別股發行收入 110.4bn，Q3 無新股權 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E477–E478、E482、E476、E525]；模型 2026 已到位融資 110.4bn [本專案：Anthropic收支模型_v0_1_8.xlsx Funding!E26]。新聞稿口徑投後估值 Series G 380bn、Series H 965bn，成立以來私募累計 >125bn [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E307、E303、E308]（[Interested-party]）。
- 現金：2025 年底 20.28bn、2026-06-30 131.65bn、2026-09-30 131.9bn（另受限現金 2.6bn）[本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E309、E448、E520–E521]；循環信用額度 2026-09 擴為 15.0bn [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E489]。
- IPO：報導預計 2026-11 募資約 100bn、目標估值 2,000bn [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E345、E314]（[Interested-party]）；S-1/A 發行價區間與股數仍留白 [本專案：anthropic/docs/reports/20261103_S1A原文摘錄.md]；Google 可在 IPO 時以發行價認購最多 10bn，與其剩餘 20bn 認購額度互抵 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E486、G487]。

## 同業與競爭態勢
| 同業 | 重疊領域 | 市占或規模 | 標記 |
|---|---|---|---|
| OpenAI | API、訂閱、程式代理 | 2026 Q2 營收 6.7bn；9 月底 run-rate 約 50bn（淨額口徑） | [本專案：OpenAI收支模型_v0_6_3.xlsx SRC_OAI!C9、C142] |
| Google（Gemini） | API、企業助理；也是 Anthropic 股東與 TPU 供應商 | Gemini 營收：本專案未建立 | — |
| MiniMax（00100.HK） | API、消費型 App | 1H26 營收 116.6m 美元 | [本專案：MiniMax收支模型_v0_2.xlsx SRC_MM!E8] |
| 智譜（02513.HK） | API、Coding 方案 | 1H26 營收 RMB 954m | [本專案：MiniMax收支模型_v0_2.xlsx SRC_MM!E75] |

- 份額（第三方）：企業 LLM API 使用份額 32%（Menlo，2025-07）；OpenRouter 文字請求份額 2.5%（2026-09 底）[本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E186、E185]（[Analogy]）。兩者差距大，反映 Anthropic 偏企業、OpenRouter 偏開發者長尾 [判斷]。
- 營收口徑：Anthropic 雲端市集營收以總額認列 [本專案：Anthropic收支模型_v0_1_8.xlsx Revenue!B67]，與 OpenAI 淨額口徑不能直接比大小。
- 價格競爭：Opus 5.5 降價 20%、Haiku 5.5 較前代降 90% [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E162–E163]。

## 主要客戶
- 財報口徑：公開 S-1 寫 2026 前九個月前五大客戶約占營收 25%、無單一客戶超過 10%，上半年 Customer C 占 11%、雲端市集約占 48% [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E463、G464、E465–E466]；S-1 未具名客戶。報導稱 2025 年約 47% 營收經由 Amazon 與 Google（兩者同時是投資人與供應商）。 [外部・未經本專案檢核：Revenue Memo：Anthropic funding，2026]
- 具名關係（多為算力供應商，同時是通路／股東）：
- Amazon／AWS：2026-04-20 承諾十年逾 1,000 億美元、最高 5 GW Trainium 等；Bedrock 為 Claude 通路。 [外部・未經本專案檢核] 公開 S-1 列 Amazon 合約最低 110bn（2026-05 至 2036-04）[本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E455]。
- Google：2026-05 報導承諾約 2,000 億美元 Google Cloud 與 TPU 容量。 [外部・未經本專案檢核：Idaho Business Review：Anthropic commits $200B Google Cloud，2026-05-06] S-1/A 列 Google Cloud 合約最低 121.3bn，另有 2028–2031 額外 TPU 估 125–150bn 不在承諾表內 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E517、F462–G462]。
- Microsoft／Nvidia：2025-11 分別承諾投資最高 50 億、100 億美元。 [外部・未經本專案檢核：Revenue Memo：Anthropic funding，2026] S-1/A 列 Microsoft Azure 合約最低 44.9bn（2026-11 至 2033-06）[本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E518]。
- 終端客戶（企業開發者等）未找到具名且有來源的清單。 [外部・未經本專案檢核]

## 前十大股東
| 排名 | 股東 | 持股比例 | 標記 |
|---|---|---|---|
| 1 | Google | 約 14%（2026-04 交易前；上限 15%；現值未公開） | [外部・未經本專案檢核：資料日 2025 初法院文件／2026-04；Allocations：Anthropic S-1 cap table…，2026-06；Revenue Memo：Anthropic funding，2026] |
| 2 | Amazon | 比例未公開；累計投資約 80–130 億美元（來源不一），另有最高 200 億里程碑；2026-04 申報文件估其持股價值 742 億美元 | [外部・未經本專案檢核：資料日 2026-04；Allocations：Anthropic S-1 cap table…，2026-06；Revenue Memo：Anthropic funding，2026] |
| 3 | 其餘：Series H 共同領投 Altimeter、Dragoneer、Greenoaks、Sequoia；Series G 的 GIC、Coatue、D.E. Shaw；Series F 的 ICONIQ、Fidelity、Lightspeed；Menlo Ventures、Spark Capital；Microsoft、Nvidia（承諾最高 50 億、100 億美元） | 均未公開 | [外部・未經本專案檢核：資料日 2025–2026；Allocations：Anthropic S-1 cap table…，2026-06；Revenue Memo：Anthropic funding，2026] |

- 口徑：未上市，無股東名冊公開；以下為媒體報導，且未提供完整持股比例。創辦人（Dario、Daniela Amodei）持股從未公開；另有 Long-Term Benefit Trust 持有特殊類別股份選任部分董事。 [外部・未經本專案檢核：Revenue Memo：Anthropic funding，2026；Allocations：Anthropic S-1 cap table…，2026-06]
- 融資：Series G 2026-02 300 億美元／估值 3,800 億；Series H 2026-05-28 650 億美元／估值 9,650 億美元（post-money）；累計逾 1,200 億美元。IPO 目標估值報導逾 2 兆美元。 [外部・未經本專案檢核：Revenue Memo：Anthropic funding，2026；Futurum：Anthropic files for IPO，2026-06-02；Mitrade：Anthropic has committed more than $100 billion to AW，2026-09-07]
- 公開 S-1 的 2026 認購：Google 認購 Series G-2 38.6 百萬股（約 10.0bn）與 Series H 10bn；Amazon 認購 Series G-2 19.3 百萬股（約 5.0bn）與 Series H 5bn [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E479–E481、E484–E485；Funding!E21、E25]；S-1 未揭露各股東持股比例。
- 註：Amazon 累計金額各來源不一（80 億 vs 約 130 億），未能調和。 [外部・未經本專案檢核]

## 近期重大事件
- 2026-11-02｜公開 S-1/A（第 1 號修訂）：Q3 營收 16.956bn、營業損失 0.567bn、調整後營業利益率 15%；首次揭露算力承諾總額約 895.4bn 與付款表；發行價區間仍留白，路演最早 11-10 之後 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E492、E494、E498、E499；anthropic/docs/reports/20261103_S1A原文摘錄.md]
- 2026-10-26｜向 SEC 公開提交 S-1：2026 上半年營收 16.286bn、Q2 首度 GAAP 營業獲利 0.974bn、6-30 現金與短期投資 131.65bn；承諾總額留白 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E418、E445、E448、D453；anthropic/docs/reports/20261102_S1原文核對.md]
- 2026-10-20｜WSJ 轉述草擬 S-1：上半年營收 16.3bn、淨損 80.7bn（約 72bn 非現金）、已鎖定算力承諾 >600bn、談判中再約 150bn；兩家客戶占上半年營收約兩成 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E384、E395–E396、E406–E407、E409]；[外部新聞：investingLive 引 WSJ]
- 2026-10-14｜告訴投資人 Q3 營收 16.8bn（季增約 45%），連續第二季獲利（S-1/A 實際 Q3 營收 16.956bn）[本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E388–E389、E492]；[外部新聞：The Star 引 Reuters／FT]
- 2026-10-07｜推出 Haiku 5.5（0.1／0.5 美元），Sonnet 5.5 快取讀取價減半 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E119–E120、E132]
- 2026-10-01｜Bloomberg：最快 11 月中上市，11-09 當週起路演，感恩節（11-26）前掛牌 [外部新聞：Bloomberg Law]
- 2026-09-29｜外流招股書草稿：2025 營收 4.6bn、淨損 42bn、算力承諾約 518bn（公開 S-1 的 2025 營收為 4.585bn）[本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E5、E284、E249、E416]
- 2026-09-22｜Opus 5.5 上市，牌價降 20% [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E162]

## 需要觀察的事項
- 收支結果｜本專案模型自由現金流 2026–2030：+7.0、−42.3、−49.6、−5.6、+42.8bn，2030 年起持續轉正 [本專案：Anthropic收支模型_v0_1_8.xlsx Funding!E14–I14、D254]；計入已到位 110.4bn 股權後累計外部資金需求為 0，現金谷底 40.2bn（2029 年，當年最低現金 12.0bn）；若完全不融資則需 102.5bn [本專案：Anthropic收支模型_v0_1_8.xlsx Funding!E26、I43、D54、H38、I127]；每 VR 等值 GW 差額 2026–2030 為 −0.94、−6.64、−5.48、−1.29、+3.27bn，2030 年才持續轉正 [本專案：Anthropic收支模型_v0_1_8.xlsx Cost!E115–I115、D412]｜S-1/A 第 2 號修訂（價格區間）、IPO 後首份季報
- SpaceXAI 讀法｜模型基準仍用每月 1.25bn × 36 個月＝45bn；S-1/A 寫滿載每月約 3.1bn、上限 120.5bn、2027 Q1 後可 90 天通知取消 [本專案：Anthropic收支模型_v0_1_8.xlsx Cost!D143；SRC_ANT!E519、E457]。若照 S-1 路徑付滿上限，累計外部資金需求由 0 變 22.0bn（首次 2029 年），是 v0.1.8 唯一單獨就讓結論翻轉的讀法 [本專案：anthropic/docs/reports/20261103_v0.1.8.md ⑤、⑪]｜2027 Q1 取消權生效前
- 承諾與附條件｜模型算力成本（現金）2028 年 146.7bn、2029 年 141.9bn，其中包含 S-1/A 三段承諾 871.4bn 裡未對應到具名合約的餘額，依付款表攤付（2029 年 34.4bn）[本專案：Anthropic收支模型_v0_1_8.xlsx Funding!G9–H9；Cost!D442、H372]；附條件 213.8bn 若全數自承諾扣除（上緣讀法），每 VR 等值 GW 差額提早到 2029 年持續轉正、現金谷底升至 92.25bn [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E512；anthropic/docs/reports/20261103_v0.1.8.md ②]；草稿轉述的談判中 150bn 若全數成交，2030 每 VR 等值 GW 差額由 3.27 降到 0.89bn，並需外部資金 6.4bn（2029 年）[本專案：Anthropic收支模型_v0_1_8.xlsx Cost!I115、I262；Funding!H116、I118]｜新合約公告、S-1/A 承諾表更新
- 2026 Q4 與一次性捐贈｜模型假設 Q4＝Q3，全年營收 50.2bn [本專案：Anthropic收支模型_v0_1_8.xlsx Revenue!E67]；2026 全成本把 Q3 非現金捐贈 1.96bn 照 Q4＝Q3 再計一次（合計 3.92bn），若只計一次，2026 每 VR 等值 GW 差額由 −0.94 轉為 +0.44bn（非現金，不影響現金與外部資金）[本專案：Anthropic收支模型_v0_1_8.xlsx Cost!E431；anthropic/docs/reports/20261103_v0.1.8.md ⑤]｜2027 年初 Q4 實績
- 雲端平台抽成率｜基準 16%，區間至 30%，直接決定營收淨額 [本專案：Anthropic收支模型_v0_1_8.xlsx Revenue!E70、K70]；S-1 寫上半年約 48% 營收經雲端市集，但未揭露抽成率 [本專案：Anthropic收支模型_v0_1_8.xlsx SRC_ANT!E466]｜IPO 後首份季報

## 來源
1. 本專案：`/home/claude/company-models/anthropic/dist/20261103_Anthropic收支模型_v0_1_8.xlsx`（SRC_ANT、Revenue、Compute、Cost、Funding）與 `anthropic/dist/20261103_Anthropic收支模型_v0_1_8.html`；`anthropic/docs/handoff/Anthropic_handoff.md`；`anthropic/docs/reports/20261103_v0.1.8.md`、`20261103_S1A原文摘錄.md`、`20261102_S1原文核對.md`、`20261102_v0.1.7.1.md`、`20261031_v0.1.7.md`、`20261031_v0.1.6.md`
2. 本專案：`/home/claude/company-models/openai/dist/20261025_OpenAI收支模型_v0_6_3.xlsx`（SRC_OAI）；`minimax/dist/20261022_MiniMax收支模型_v0_2.xlsx`（SRC_MM）；`alphabet/data/alphabet_facts_20261008.json`
3. [Anthropic, PBC Form S-1/A（Amendment No. 1）](https://www.sec.gov/Archives/edgar/data/2130357/000162828026068676/anthropicpbc-sx1ax1.htm)（SEC EDGAR，2026-11-02）
4. [Anthropic, PBC Form S-1](https://www.sec.gov/Archives/edgar/data/2130357/000162828026067168/anthropicpbc-sx1.htm)（SEC EDGAR，2026-10-26）
5. [Anthropic IPO draft shows nearly 15-fold revenue jump and $600 billion in compute deals](https://investinglive.com/stocks/anthropic-ipo-draft-shows-nearly-15-fold-revenue-jump-and-600-billion-in-compute-deals/)（investingLive，2026-10-20）
6. [Anthropic revenue jumps to $16.8 billion in third quarter, FT reports](https://www.thestar.com.my/tech/tech-news/2026/10/15/anthropic-revenue-jumps-to-168-billion-in-third-quarter-ft-reports)（The Star 轉 Reuters，2026-10-15）
7. [Anthropic Targets Mega-IPO Before Thanksgiving Holiday](https://news.bgov.com/banking-law/anthropic-said-to-target-mega-ipo-before-thanksgiving-holiday)（Bloomberg Law，2026-10-01）
8. [Factbox-Anthropic IPO prospectus shows key executives and board](https://www.thestar.com.my/tech/tech-news/2026/09/29/factbox-anthropic-ipo-prospectus-shows-key-executives-and-board)（The Star 轉 Reuters，2026-09-29；管理層）

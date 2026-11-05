---
ticker: META
name: Meta
name_full: Meta Platforms, Inc.
listing: Nasdaq｜META（Class A）
hq: 美國加州門洛帕克 [外部・未經本專案檢核]
founded: 2004 [外部・未經本專案檢核：Meta DEF 14A 2026，2026-04-16]
employees: 約 66,775（2026Q3 季末）[本專案：meta_facts_2026q3.json txt.headcount.2026q3]
fiscal_year_end: 12 月底
asof: 2026-11-04
---

## 一句話定位
全球最大的社群廣告平台之一（Facebook、Instagram、WhatsApp、Messenger），靠賣廣告版位賺錢；AI 算力全部自用，用來改善推薦與廣告投放，不對外出租 [外部・未經本專案檢核]，本專案模型也以「全部自用」處理 [本專案：20261104_Meta個股模型_v0.1.2.xlsx In!I48]。本專案個股模型 v0.1.2 的命題是：AI 資本支出能否由廣告現金流支撐，以及自用算力能否帶來足以回本的廣告變現提升 [本專案：meta/README.md 命題]。

## 公司在做什麼
- Family of Apps（應用程式家族）：Facebook、Instagram、Messenger、WhatsApp 等免費服務，靠廣告主購買版位付費；2026Q3 每日活躍人數（DAP，至少用一個 App 的人）36.4 億 [本專案：20261104_Meta個股模型_v0.1.4.xlsx Q!I49]。
- 廣告營收可拆成「曝光數 × 每則廣告平均價格」：2026Q3 曝光數年增 10%、每則價格年增 14% [本專案：20261104_Meta個股模型_v0.1.4.xlsx Q!I47、Q!I48]。
- Reality Labs：VR 頭盔、AI 眼鏡與元宇宙相關研發，營收小、長期虧損；公司指引 2026 年虧損 180–190 億美元、2027 年約 150 億美元 [本專案：meta_facts_2026q3.json txt.guid_rl_loss2026.2026q3、call.rlLoss2027]。
- AI 為什麼需要它：廣告排序、內容推薦與生成式 AI 助手（Meta AI、Muse 個人代理）都要在大量 GPU 機房上訓練與推論；Meta 自建機房為主，近期以第三方雲端合約補足容量 [本專案：meta_facts_2026q3.json fu.ownInfra、fu.cloudAug]。

## 主要產品與營收占比
| 產品／部門 | 內容（白話） | 營收占比 | 期間 | 標記 |
|---|---|---|---|---|
| **Family of Apps**（分部） | 四大 App 的廣告與其他收入 | 98.90%（1,987.59 億美元）；9M26 99.30% | FY2025；2026 年前三季 | [本專案：20261104_Meta個股模型_v0.1.4.xlsx Q!K7、Q!L7；Derived＝÷Q!K10、Q!L10] |
| 　Advertising | App 內的廣告版位 | 97.62%（1,961.75 億美元）；9M26 97.65% | FY2025；9M26 | [本專案：20261104_Meta個股模型_v0.1.4.xlsx Q!K5、Q!L5；Derived] |
| 　Other revenue | WhatsApp 商業訊息、Meta One 訂閱等 | 1.29%（25.84 億美元）；9M26 1.64% | FY2025；9M26 | [本專案：20261104_Meta個股模型_v0.1.4.xlsx Q!K6、Q!L6；Derived] |
| **Reality Labs**（分部） | VR／AR 硬體與軟體 | 1.10%（22.07 億美元）；9M26 0.70% | FY2025；9M26 | [本專案：20261104_Meta個股模型_v0.1.4.xlsx Q!K9、Q!L9；Derived] |

- 口徑：兩分部合計 2,009.66 億美元（FY2025）、1,817.13 億美元（9M26）＝合併營收，核對列為 0 [本專案：20261104_Meta個股模型_v0.1.4.xlsx Q!K10、Q!L10、Q!K11]。
- 分部營業利益：FY2025 Family of Apps 1,024.69 億、Reality Labs −191.93 億美元 [本專案：20261104_Meta個股模型_v0.1.4.xlsx Q!K14、Q!K15]；Family of Apps 營業利益率 FY2025 51.5%、9M26 37.4%（9M26 含法律一次性費用）[本專案：20261104_Meta個股模型_v0.1.4.xlsx PL!D39、PL!E39]。
- 成長：近四季（2025Q4–2026Q3）廣告營收年增 27.5%，Other revenue 年增 64.6% [本專案：20261104_Meta個股模型_v0.1.4.xlsx Q!B61、Q!B66]。
- AI 相關營收：公司未揭露，本專案未建立（模型以「自用 AI 影子效益」代替，屬評價口徑，非財報營收）[本專案：20261104_Meta個股模型_v0.1.4.xlsx Thesis!N70]。

## 市場定位與產業位置
- 數位廣告市占：本專案未建立。
- 算力全部自用：本專案以資本支出換算，評價日（2026-09-30）自用 AI 容量約 5,038 MW（支出口徑、含在建），其中在建工程約 1,873 MW [本專案：20261104_Meta個股模型_v0.1.4.xlsx Thesis!C20、Thesis!C21]；在建工程 894.99 億美元 [本專案：20261104_Meta個股模型_v0.1.2.xlsx In!E80]。
- 表外算力承諾大：尚未起租的租賃義務 3,557.3 億、不可取消合約承諾 4,545.2 億美元，另有最多 147.2 億美元的或有雲端算力採購 [本專案：20261104_Meta個股模型_v0.1.4.xlsx Recon!B38、Recon!B39、Recon!B45]；模型未把這些列入基準 MW，可能低估自用算力 [本專案：20261104_Meta個股模型_v0.1.4.xlsx Recon!A36]。
- 也是 GPU 雲的大客戶：與 Nebius 第二份協議專用容量 120 億美元（另有未售容量承購上限 150 億）、第一份 29 億美元 [本專案：nebius_facts_20261007.json ct.meta2.dedicated、ct.meta2.backstop、ct.meta1]；與 CoreWeave 2026-04 擴約約 210 億美元至 2032-12 [本專案：coreweave/data/permw_inputs_20261007.json gen-vr-meta]。
- 自研晶片 MTIA：完整 10-Q、法說與 follow-up call 皆未提及，占比未找到 [本專案：20261104_Meta個股模型_v0.1.4.xlsx Recon!B46]。

## 管理層與經營風格
| 職位 | 姓名 | 任職起 | 背景 |
|---|---|---|---|
| 創辦人、董事長兼執行長 | Mark Zuckerberg | 執行長 2004；董事長 2012 | 創辦人 [外部・未經本專案檢核：Meta DEF 14A 2026，2026-04-16] |
| 總裁兼副董事長 | Dina Powell McCormick | 2026（2025–2026 任顧問） | [外部・未經本專案檢核：Meta DEF 14A 2026，2026-04-16] |
| 營運長 | Javier Olivan | 2022 | [外部・未經本專案檢核：Meta DEF 14A 2026，2026-04-16] |
| 財務長 | Susan Li | 2022 | [外部・未經本專案檢核：Meta DEF 14A 2026，2026-04-16] |
| 技術長 | Andrew Bosworth | 2022 | 曾主管 Reality Labs [外部・未經本專案檢核] |
| 產品長 | Chris Cox | 2020（2014–2019 曾任） | [外部・未經本專案檢核：Meta DEF 14A 2026，2026-04-16] |
| 法務長 | C.J. Mahoney | 2026 | [外部・未經本專案檢核：Meta DEF 14A 2026，2026-04-16] |

- 資本配置：FY2025 買回 262.48 億、配息 53.24 億美元；2026 年前三季買回為 0 [本專案：20261104_Meta個股模型_v0.1.4.xlsx Q!K43、Q!K44、Q!L43]。2026 資本支出指引（含融資租賃本金）由年初 1,150–1,350 億逐季上調到 1,350–1,450 億美元 [本專案：meta_facts_2026q3.json txt.guid_capex2026.2025q4、txt.guid_capex2026.2026q3]。
- 融資轉向：票據面額 840 億美元（2026-09-30）[本專案：20261104_Meta個股模型_v0.1.4.xlsx In!E12]；2026-10 起首度對外募股，強制轉換特別股淨額 124.58 億、Class A 遠期出售淨額 249.17 億美元 [本專案：20261104_Meta個股模型_v0.1.2.xlsx In!E16、In!E24]。
- 治理：雙層股權，Zuckerberg 持有約 60.6% 投票權（2026-10-01）[外部・未經本專案檢核：Meta DEF 14C，2026-11-02]；新設無表決權 Class C 股，以 1:1 股票股利發放 [本專案：meta_facts_2026q3.json call.nonvoting]。
- 其餘經營風格：本專案尚未研究。

## 同業與競爭態勢
| 同業 | 重疊領域 | 市占或規模 | 標記 |
|---|---|---|---|
| Alphabet（GOOGL） | 搜尋與 YouTube 廣告、AI 助手 | 市占本專案未建立；NTM EV/EBITDA 14.71 倍 | [本專案：20261104_Meta個股模型_v0.1.2.xlsx In!E52] |
| AppLovin（APP） | 行動 App 廣告 | NTM EV/EBITDA 11.38 倍 | [本專案：20261104_Meta個股模型_v0.1.2.xlsx In!E55] |
| Reddit（RDDT） | 社群廣告 | NTM EV/EBITDA 13.63 倍 | [本專案：20261104_Meta個股模型_v0.1.2.xlsx In!E56] |
| Pinterest（PINS） | 社群／視覺廣告 | NTM EV/EBITDA 6.91 倍 | [本專案：20261104_Meta個股模型_v0.1.2.xlsx In!E53] |
| Snap（SNAP） | 社群廣告、AR 眼鏡 | NTM EV/EBITDA 7.43 倍 | [本專案：20261104_Meta個股模型_v0.1.2.xlsx In!E57] |
| The Trade Desk（TTD） | 程式化廣告採購 | NTM EV/EBITDA 6.53 倍 | [本專案：20261104_Meta個股模型_v0.1.2.xlsx In!E54] |

- 六家中位數 9.41 倍，Meta 自身 8.99 倍（2026-10-28）；除 GOOGL 外五家為 2026-10-07 基準，未更新 [本專案：20261104_Meta個股模型_v0.1.2.xlsx Val!B110、In!E58；meta/README.md 日期與口徑註記]。
- 競爭風險：AI 答案與代理分流廣告曝光，以及 Alphabet、TikTok 的 AI 廣告工具追上，可能讓 Meta 的價格溢價消退，即模型情境 C [本專案：20261104_Meta個股模型_v0.1.2.xlsx In!H112]。
- 規模不同、商業模式不同（Meta 無流量取得成本），倍數只能當量級參考 [判斷]。

## 主要客戶
- 財報口徑：10-K 稱「substantially all of our revenue」來自廣告主在 App 上購買廣告版位，且廣告主沒有長期採購承諾；本次擷取段落未見單一客戶占營收比重的揭露，標示為「未找到」（查過 FY2025 10-K 風險因子段）。 [外部・未經本專案檢核：Meta FY2025 Form 10-K，2026-01]

| 客戶 | 關係 | 依據 | 標記 |
|---|---|---|---|
| 廣告主（未具名） | 購買 Facebook、Instagram 廣告版位 | 10-K 稱多數廣告主只把一小部分預算放在 Meta | [外部・未經本專案檢核：Meta FY2025 Form 10-K，2026-01] |
| 具名大客戶 | 未找到 | 10-K 擷取段未具名任何廣告主 | [外部・未經本專案檢核：Meta FY2025 Form 10-K，2026-01] |

## 前十大股東
| 排名 | 股東 | 持股比例 | 標記 |
|---|---|---|---|
| — | Mark Zuckerberg | 總投票權約 60.6%；Class A 約 122 萬股、Class B 約 3.41 億股 | [外部・未經本專案檢核：Meta DEF 14C，2026-11-02；資料日 2026-10-01] |
| 1 | BlackRock, Inc. | 6.82% | [外部・未經本專案檢核：VCP Scanner META institutional holders（13F），資料期 2026Q2] |
| 2 | Vanguard Capital Management LLC | 5.64% | [外部・未經本專案檢核：VCP Scanner META institutional holders（13F），資料期 2026Q2] |
| 3 | FMR LLC | 3.77% | [外部・未經本專案檢核：VCP Scanner META institutional holders（13F），資料期 2026Q2] |
| 4 | State Street Corp | 3.62% | [外部・未經本專案檢核：VCP Scanner META institutional holders（13F），資料期 2026Q2] |
| 5 | Geode Capital Management, LLC | 2.16% | [外部・未經本專案檢核：VCP Scanner META institutional holders（13F），資料期 2026Q2] |
| 6 | JPMorgan Chase & Co | 1.81% | [外部・未經本專案檢核：VCP Scanner META institutional holders（13F），資料期 2026Q2] |
| 7 | Capital World Investors | 1.61% | [外部・未經本專案檢核：VCP Scanner META institutional holders（13F），資料期 2026Q2] |
| 8 | Capital Research Global Investors | 1.56% | [外部・未經本專案檢核：VCP Scanner META institutional holders（13F），資料期 2026Q2] |
| 9 | Morgan Stanley | 1.55% | [外部・未經本專案檢核：VCP Scanner META institutional holders（13F），資料期 2026Q2] |
| 10 | Vanguard Portfolio Management LLC | 1.51% | [外部・未經本專案檢核：VCP Scanner META institutional holders（13F），資料期 2026Q2] |

- 口徑：機構為 13F 申報（2026Q2），網站未說明分母是 Class A 或全部股數；Vanguard 因內部重組分兩個申報實體列示。Zuckerberg 為投票權口徑，另列不排名，兩者不可直接合併排序。 [外部・未經本專案檢核：VCP Scanner META institutional holders，資料期 2026Q2；Meta DEF 14C，2026-11-02]
- Zuckerberg 經濟權約 13.4%（上列 A＋B 股數 ÷ 2026-10-16 流通股 25.56 億股）[Derived：外部股數 ÷ 本專案 20261104_Meta個股模型_v0.1.2.xlsx In!E7]。

## 近期重大事件
- 2026-11-02｜寄發 DEF 14C：Class C 無表決權股的重述章程 2027-01-19 生效，每股 Class A／B 配發一股 Class C；同時揭露與 Zuckerberg 的 Founder Alignment Agreement｜股數約翻倍，模型每股數字為股票股利前口徑 [外部新聞：SEC EDGAR（Meta DEF 14C）；本專案：meta/README.md 每股口徑]
- 2026-10-26～29｜首度對外募股：Class A 遠期出售 4,000 萬股（淨額 249.17 億美元，預計 FY2027 交割），6.50% 強制轉換特別股淨額 124.58 億美元於 10-29 交割｜資本支出開始需要外部股權資金 [本專案：20261104_Meta個股模型_v0.1.2.xlsx In!E23、In!E24、In!E16]
- 2026-10-21｜Q3 財報：營收 646.01 億美元、稀釋 EPS 3.83 美元（排除法律費用為 7.40）；資本支出 349.06 億、單季自由現金流 −21.2 億美元｜利潤率被 AI 成本與一次性費用壓低 [本專案：20261104_Meta個股模型_v0.1.2.xlsx Q!I10、Q!I32、In!E125、Q!I40、Q!I41]
- 2026-10-21｜Q4 營收指引 730–760 億美元；2026 資本支出指引 1,350–1,450 億、總費用 1,750–1,790 億美元；2027 資本支出待 Q4 法說才給｜模型 Q4 營收 763.4 億略高於指引上緣 [本專案：meta_facts_2026q3.json txt.guid_q4rev.2026q3、txt.guid_capex2026.2026q3、txt.guid_exp2026.2026q3、call.capex2027；20261104_Meta個股模型_v0.1.4.xlsx Recon!B27]
- 2026Q3｜多州檢察長協議：Q3 認列法律費用 98.9 億美元，約 127 億美元分 10 年支付，另 53 億或有金額未認列｜模型自 FY2027 起每年扣協議付款 12.7 億美元 [本專案：meta_facts_2026q3.json txt.legal_charge.2026q3、legal.agSettle.total、legal.agContingent；20261104_Meta個股模型_v0.1.4.xlsx CF!H17]
- 2026-07-29｜Q2 財報：法律費用 24.0 億、遣散費用 11.8 億美元；2026 資本支出指引調為 1,300–1,450 億美元；季末員工 75,472 人，Q3 季末降到 66,775 人｜成本結構調整與資本支出加碼同時發生 [本專案：meta_facts_2026q3.json txt.legal_charge.2026q2、txt.severance.2026q2、txt.guid_capex2026.2026q2、txt.headcount.2026q2、txt.headcount.2026q3]

## 需要觀察的事項
- 自用 AI 的讀法（評等關鍵）｜讀法 2（自用 AI 價值中性）加權目標價 844.10 美元、評等「買進」；讀法 1（只計成本）428.70 美元、「放空」，DCF 只剩 134.92 美元 [本專案：20261104_Meta個股模型_v0.1.4.xlsx Val!B20、Val!D20、Val!B21、Val!D21、Val!B22]｜每次財報對 AI 投資回報（廣告轉換率、價格）的說法
- 每 MW 持有成本（IF_HoldEcon）｜四世代同時取 Tokenomics 低／高成本欄時，加權目標價 765.15／1,000.38 美元；低值已低於買進下限 765.61 [本專案：20261104_Meta個股模型_v0.1.2.xlsx Sens!D28、Sens!D29；meta/README.md 投資建議]｜Tokenomics 下一版快照
- FY2027 資本支出｜模型情境 A／B／C 為 2,365.6／2,057.1／1,748.5 億美元，市場共識（財報前、不含租賃本金）1,930.6 億 [本專案：20261104_Meta個股模型_v0.1.2.xlsx In!E72、In!E71、In!E73、Recon!F15]｜Q4 法說（2027 年初）首次給 2027 指引
- 現金流與收支關卡｜情境 B 自由現金流 FY2027 −382.5 億、FY2028 −208.8 億美元；五年（FY2026–FY2030）自由現金流扣股利 A／B／C 為 −195.6／+149.8／+490.1 億美元 [本專案：20261104_Meta個股模型_v0.1.4.xlsx CF!H55、CF!I55、CF!B6、CF!C6、CF!D6]｜公司債發行與遠期出售交割
- AI 折舊壓低利潤率｜情境 B 合併營業利益率由 FY2025 41.4% 降到 FY2028 24.5%；FY2027／FY2028 EPS 30.05／27.90 美元，低於共識 34.07／39.53 [本專案：20261104_Meta個股模型_v0.1.4.xlsx PL!D84、PL!I84、PL!H99、PL!I99、Recon!F12、Recon!F13]｜Q4 財報的 2027 費用指引
- 每則廣告價格｜2026Q3 價格年增 14%、曝光年增 10%；長期價格年增取 1%／5% 時加權目標價 810.18／879.96 美元 [本專案：20261104_Meta個股模型_v0.1.4.xlsx Q!I48、Q!I47、Sens!D10、Sens!D11]｜各季新聞稿的曝光與價格年增

## 來源
1. 本專案：/home/claude/wt23/meta/dist/20261104_Meta個股模型_v0.1.2.xlsx（In、Q、Thesis、PL、CF、Val、Recon、Sens）、meta/README.md、meta/docs/reports/20261104_meta_v0.1.2_TK快照v21.9.md（2026-11-04）
2. 本專案：/home/claude/wt23/meta/data/meta_facts_2026q3.json（2026-10-30 擷取）
3. 本專案：/home/claude/wt23/nebius/data/nebius_facts_20261007.json、/home/claude/wt23/coreweave/data/permw_inputs_20261007.json（2026-10-07）
4. [Meta Platforms DEF 14A 2026（管理層名單）](https://www.sec.gov/Archives/edgar/data/1326801/000162828026025532/meta-20260416.htm)（2026-04-16）
5. [Meta Leadership and Governance](https://investor.atmeta.com/leadership-and-governance/default.aspx)（擷取 2026-11-04）
6. [Meta Platforms DEF 14C（Class C 股票股利、投票權）](https://www.sec.gov/Archives/edgar/data/0001326801/000162828026068935/meta-20261102.htm)（2026-11-02）
7. [Meta Platforms Form 10-K FY2025](https://www.sec.gov/Archives/edgar/data/1326801/000162828026003942/meta-20251231.htm)（2026-01）
8. [VCP Scanner：META institutional holders（13F）](https://vcpscanner.com/stock/meta/institutional-holders)（資料期 2026Q2，擷取 2026-11-04）

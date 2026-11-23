---
ticker: PANW
name: Palo Alto Networks
name_full: Palo Alto Networks, Inc.
listing: Nasdaq｜PANW
hq: 美國加州聖塔克拉拉 [外部・未經本專案檢核]
founded: 2005 [外部・未經本專案檢核]
employees: 21,921（FY2026 年底，含 CyberArk 併入約 4,223 人）[本專案：research_paloalto.md §6]
fiscal_year_end: 7 月底（FY2026＝2025-08-01 至 2026-07-31）
asof: 2026-11-23
---

## 一句話定位
企業資安平台廠商：從防火牆起家，一路擴到雲端、SOC（資安監控中心）、端點與身分安全，位在軟體價值鏈的「安全控制層」；靠訂閱與支援按防護範圍（身分、端點、工作負載）乘上模組數計價賺錢 [本專案：research_paloalto.md §1]。AI 是增量還是侵蝕的爭點：AI 代理（非人類身分）讓需要保護的對象變多，但價格壓力與併購整合成本也同時存在，公司並未揭露 AI 營收 [本專案：paloalto/README.md 命題；research_paloalto.md §5]。

## 公司在做什麼
- 防火牆與網路安全：在企業網路進出口檢查並阻擋攻擊流量，早期以硬體設備銷售為主，現在多為軟體與訂閱 [外部・未經本專案檢核]。
- 雲端安全（Prisma）、SOC 與端點（Cortex）：替客戶監控雲端工作負載與主機，並用分析自動找出威脅 [本專案：research_paloalto.md §1]。
- 身分安全：2026 年 2 月完成併入 CyberArk，新增特權帳號與機器身分管理，品牌改名 Idira [本專案：research_paloalto.md §1]。
- 補 AI 能力的小型併購：Protect AI、Koi、Portkey、Embrace Mobile、Console Systems 等，目標是保護 AI 模型與 AI 代理的工作負載 [本專案：research_paloalto.md §5]。
- 為什麼 AI 需要它：AI 代理會自己登入系統、呼叫資料，每多一個代理就多一個需要驗證與監控的身分 [本專案：paloalto/README.md 命題]。

## 主要產品與營收占比
| 產品／部門 | 內容（白話） | 營收占比 | 期間 | 標記 |
|---|---|---|---|---|
| Product | 防火牆硬體與軟體授權（含 CyberArk 地端身分軟體授權） | 19.9%（US$2,280M） | FY2026 | [本專案：research_paloalto.md §2、§3；占比為 Derived，2,280÷11,480] |
| Subscription | 威脅防護、雲端交付安全、SASE（雲端網路加安全的整合服務）、SOC 等訂閱 | 54.3%（US$6,239M） | FY2026 | [本專案：research_paloalto.md §3；占比為 Derived，6,239÷11,480] |
| Support | 硬體維護與技術支援 | 25.8%（US$2,961M） | FY2026 | [本專案：research_paloalto.md §3；占比為 Derived，2,961÷11,480] |
| 合計 | — | 100%（US$11,480M） | FY2026 | [本專案：research_paloalto.md §3] |

- 官方只有一個應報導分部，所以沒有更細的分部營收；地區為美洲 66.9%、歐非中東 21.1%、亞太 12.0% [本專案：research_paloalto.md §2；占比為 Derived，用 7,679／2,428／1,373 除以 11,480]。
- 公司自訂指標 NGS ARR（新世代安全的年化經常性營收）FY2026 年底 US$9.1B，年增 63%，其中含併購；RPO（已簽約但尚未認列的營收）US$21.2B [本專案：research_paloalto.md §4]。
- 毛利率 70.4%，比上年的 73.4% 低；Subscription and support 毛利率 69.2% [本專案：research_paloalto.md §6]。
- AI 營收占比：本專案未建立（公司不揭露）[本專案：research_paloalto.md §9]。

## 市場定位與產業位置
- 同時覆蓋網路、雲端、SOC、身分四條線的平台型廠商，賣點是「整合採購」；資安支出屬合規剛性，本專案給的不可替代性分數為 1（最高）[本專案：map/dist/data.json 資安次產業]。市占數字：本專案未建立。
- FY2026 有機營收年增 14.4%，合併年增 24.5%；差距來自 CyberArk 與 Chronosphere 併入，併入貢獻營收 US$930M [本專案：research_paloalto.md §3；Derived，11,480−930 後與 FY2025 比]。
- 每單位年價、受保護身分與端點數、模組採用分布：公司不揭露，本專案模型只能用比例尺假設 [本專案：research_paloalto.md §9；paloalto/README.md]。
- 推論成本：公司向模型商購買，不自建算力；模型推算 FY2032 推論成本增量只占營業成本約 4%，量級小於 AI 增量與價格侵蝕 [本專案：20261117_Palo Alto Networks個股模型_v0.2.1.xlsx README!A9、README!A16]。
- 通路集中度見「主要客戶」節。

## 管理層與經營風格
| 職位 | 姓名 | 任職起 | 背景 |
|---|---|---|---|
| 董事長暨執行長 | Nikesh Arora | 2018 | 公司官網管理層頁列示 [外部・未經本專案檢核] |
| 總裁 | BJ Jenkins | 官網未載 | 公司官網管理層頁列示 [外部・未經本專案檢核] |
| 財務長 | Dipak Golechha | 2020 | 公司官網管理層頁列示 [外部・未經本專案檢核] |
| 產品暨技術長（CPTO） | Lee Klarich | 2006 年加入；2025-08 任現職 | FY2026 新任董事並主持安全委員會 [外部・未經本專案檢核] |
| 新世代安全總裁 | Karim Temsamani | 2024-09 | 公司官網管理層頁列示 [外部・未經本專案檢核] |
| Idira 總裁 | Sonny Singh | 官網未載 | 負責併入的身分安全業務 [外部・未經本專案檢核] |

- 經營風格：本專案尚未研究。
- 本專案記錄到的資本配置：以併購擴張，單筆如 Console Systems 約 US$500M、Protect AI 約 US$635M；CyberArk 發行約 112 百萬股 [本專案：research_paloalto.md §5、§8]。
- 股票薪酬占營收 15.8%（FY2021 為 21.8%）；FY2028 調整後自由現金流率目標 40% [本專案：research_paloalto.md §6]。
- 近期人事：Klarich 於 FY2026 進董事會，審計委員會另有新任獨立董事 Mark Goodburn [外部・未經本專案檢核]。

## 同業與競爭態勢
| 同業 | 重疊領域 | 市占或規模 | 標記 |
|---|---|---|---|
| Fortinet | 防火牆、SASE | 本專案未建立 | — |
| Zscaler | 雲端安全、零信任網路 | 本專案未建立 | — |
| Check Point | 防火牆、雲端安全 | 本專案未建立 | — |
| CrowdStrike | 端點、SOC、身分 | 本專案未建立 | — |
| Okta、Microsoft Entra | 身分安全（與 CyberArk 重疊） | 本專案未建立 | — |

- 本專案的比較同業為 Fortinet、Zscaler、Check Point、Cloudflare 等；規模與市占需另行建立 [本專案：paloalto/README.md；20261117_Palo Alto Networks個股模型_v0.2.1.xlsx Val!A47:A58]。
- 競爭格局的方向：資安廠商都在往「單一平台」收斂，並各自加碼 AI 代理安全產品 [判斷]。
- 身分線有大型作業系統與雲端廠商的內建功能做替代；可觀測性（Chronosphere）市場也已擁擠 [本專案：research_paloalto.md §8 空方第 2 點，第三方意見]。

## 主要客戶
- 財報揭露：FY2026、FY2025、FY2024 都沒有任何單一終端客戶占營收超過 10% [外部・未經本專案檢核：PANW 10-K FY2026 Note 1 集中度，2026-09-10]。
- FY2026 有兩家經銷商各占營收 10% 以上，合計 30%；2026-07-31 有一家經銷商占應收帳款毛額 19% [外部・未經本專案檢核：PANW 10-K FY2026，2026-09-10]。
| 客戶 | 關係 | 依據 | 標記 |
|---|---|---|---|
| 經銷商兩家（未具名） | 通路，買貨後轉售給企業 | 10-K 只寫「兩家經銷商」，未載名稱與各自比例 | [外部・未經本專案檢核：PANW 10-K FY2026，2026-09-10] |
| NTT DATA | 全球策略聯盟（共同推 AI 安全轉型） | 公司新聞稿；屬合作夥伴，非財報客戶集中度 | [外部・未經本專案檢核：Palo Alto Networks 新聞稿，2026-08-20] |
| AT&T Business | 聯合提供分散式企業安全服務 | 公司新聞稿；屬合作夥伴 | [外部・未經本專案檢核：Palo Alto Networks 新聞稿，2026-07-29] |
- 未找到：兩家經銷商的名稱與各自占比。查過 10-K FY2026 的集中度附註與風險因素、2026 委託書；皆未具名。

## 前十大股東
| 排名 | 股東 | 持股比例 | 標記 |
|---|---|---|---|
| 1 | Vanguard Capital Management（60,785,782 股） | 7.4% | [外部・未經本專案檢核：PANW 2026 委託書，資料日期 2026-09-18] |
| 2 | BlackRock, Inc.（49,244,010 股） | 6.0% | [外部・未經本專案檢核：PANW 2026 委託書，資料日期 2026-09-18] |
| 3–10 | 未找到 | 未找到 | 委託書只列 5% 以上持有人 |

- 口徑：委託書「5% 以上實益持有人」表，分母為 2026-09-18 流通股 818,443,112 股。其餘持股人未找到：查過委託書（無 5% 以下名單）；Yahoo Finance 持股頁的比例欄與委託書明顯不一致、不予採用；Simply Wall St 頁面資料日期為 2025-07，過舊不採用。不推估 [外部・未經本專案檢核：2026 委託書，2026-09-18]。

## 近期重大事件
- 2026-11-18｜FY2027 第一季財報：營收 US$3,445M、年增 39%；NGS ARR US$9.73B；公司上調 FY2027 全年指引，營收 US$14.52–14.57B、NGS ARR US$11.475–11.525B。同日宣布擬收購 Applied Compute（開放權重模型的後訓練與推論平台）。本專案模型建於 2026-11-17，尚未納入這次實績與指引 [外部新聞：Palo Alto Networks 新聞稿（Placera 轉載）]。
- 2026-10-30｜公司建議股東不要接受 Tutanota LLC 的小額公開收購要約（mini-tender，常以低於市價收購少量股份）[外部新聞：Palo Alto Networks 新聞稿標題]。
- 2026-09-22｜Unit 42 Continuous Frontier AI Defense 上線，提供 Anthropic 與 OpenAI 最新模型給客戶用於資安防禦；這是把 AI 模型商當供應商與合作對象的例子 [外部新聞：Palo Alto Networks 新聞稿標題]。
- 2026-09-01｜FY2026 第四季財報：全年營收 US$11.48B；NGS ARR US$9.10B；FY2030 NGS ARR 目標由 US$15B 調到 US$20B [外部新聞：Palo Alto Networks 新聞稿；本專案：research_paloalto.md §8]。同日宣布收購 Console（AI 原生的代理式工作流程平台，併入 Cortex）。
- 2026-08-27｜Embrace Mobile 收購完成，金額約 US$325M；2026-09-01 Console Systems 完成，約 US$500M [本專案：research_paloalto.md §5]。
- 2026-08-20｜與 NTT DATA 簽署全球策略聯盟 [外部新聞：Palo Alto Networks 新聞稿標題]。

## 需要觀察的事項
- 代理身分數與模組加購（n32、m32）｜模型最大的假設，決定 AI 增量；公司若揭露非人類身分產品營收或採用數即可檢驗 [本專案：paloalto/README.md 承重 Assumed]｜下一次財報與法說（FQ2 日期未公告）。
- 價格侵蝕（e32）與 Subscription and support 毛利率｜FY2026 為 69.2%，是本專案的可證偽指標 [本專案：research_paloalto.md §6；paloalto/README.md]｜每季財報。
- 有機成長 vs 併購貢獻｜FY2026 有機 14.4%、合併 24.5%；公司未拆分有機 NGS ARR，需要看後續是否揭露 [本專案：research_paloalto.md §3、§4]｜FQ2 財報。
- 費用率路徑｜Non-GAAP 營業利益率 FY2026 29.2%、FY2027 原指引 29.5%，股票薪酬占營收 15.8%；模型假設逐年改善 [本專案：research_paloalto.md §6；paloalto/README.md G3 診斷]｜每季營業利益率與股票薪酬實績。
- 通路集中度｜兩家經銷商合計占營收 30%，其中任一家的條件變動都會影響收入認列與應收帳款 [外部・未經本專案檢核：PANW 10-K FY2026，2026-09-10]｜下一份 10-K（約 2027-09）。
- 模型更新｜2026-11-18 的實績與上調指引尚未進模型，原模型 FY2027 營收指引區間為 US$14.10–14.20B [本專案：research_paloalto.md §7]｜下一版模型。

## 來源
1. 本專案：`software_ai/paloalto/data/research_paloalto.md`；`software_ai/paloalto/data/facts_paloalto.csv`；`software_ai/paloalto/README.md`；`software_ai/paloalto/dist/20261117_Palo Alto Networks個股模型_v0.2.1.xlsx`（README、Val）；`software_ai/map/dist/data.json`
2. [Palo Alto Networks 管理層頁](https://paloaltonetworks.com/about-us/management)（2026-11-23 查閱）
3. [PANW Form 10-K FY2026](https://www.sec.gov/Archives/edgar/data/1327567/000132756726000023/panw-20260731.htm)（2026-09-10）
4. [PANW DEF 14A（2026 委託書）](https://www.sec.gov/Archives/edgar/data/0001327567/000130817926000432/panw-20261027.htm)（2026-10-27）
5. [Palo Alto Networks 2026 新聞稿列表](https://www.paloaltonetworks.com/company/press/2026)（2026-11-23 查閱）
6. [Palo Alto Networks Reports Fiscal First Quarter 2027 Financial Results（Placera 轉載）](https://www.placera.se/pressmeddelanden/palo-alto-networks-reports-fiscal-first-quarter-2027-financial-results-20261118)（2026-11-18）
7. [Palo Alto Networks Reports Fiscal Fourth Quarter and Fiscal Year 2026 Financial Results](https://www.paloaltonetworks.com/company/press/2026/palo-alto-networks-reports-fiscal-fourth-quarter-and-fiscal-year-2026-financial-results)（2026-09-01）

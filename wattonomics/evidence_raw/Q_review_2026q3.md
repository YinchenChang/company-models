# 獨立核對：F_funding.csv 第 228–246 列（MSFT FY27Q1、GOOGL 2026Q3）

核對者：未參與製作的獨立代理｜日期：2026-10-29｜CSV 只讀，沒有修改。
核對對象：`/home/claude/company-models/wattonomics/evidence_raw/F_funding.csv` 第 228–246 列（第 1 列為表頭）。

原始文件位置的寫法：
- 10-Q（docx 沒有頁碼）：以章節／附註／表格標示，例如「現金流量表」「Note 12 補充現金流量表」「Part II Item 2」。
- 法說逐字稿：以段落序號標示（python-docx 讀出的第 n 段）。
- Alphabet 新聞稿：以 PDF 實體頁碼標示。
- MSFT xlsx：以工作表名稱標示。

## 一、逐列核對

| 列號 | 項目 | CSV 值 | 文件所載（≤25 字引文與位置） | 判定 | 備註 |
|---|---|---|---|---|---|
| 228 | MSFT FY27Q1 ocf | 58.0 | "Net cash from operations … 58,000"（10-Q 現金流量表，三個月至 2026-09-30） | 屬實 | MD&A 寫 "increased $12.9 billion to $58.0 billion"，數字一致。 |
| 229 | MSFT FY27Q1 capex_cash | 39.367 | "Additions to property and equipment (39,367)"；去年同期 (19,394)（10-Q 現金流量表 Investing） | 屬實 | 這是現金基礎的 PP&E 增添，與 capex_cash 定義相符。 |
| 230 | MSFT FY27Q1 fin_lease_new | 5.192 | "Right-of-use assets obtained … Finance leases 5,192"；去年同期 9,147（10-Q Note 12 補充現金流量表） | 屬實 | 數值與去年同期都正確。CFO 所說的 7.4 是另一個口徑，見第二節。fin_lease_new 應保留 5.192。 |
| 231 | MSFT FY27Q1 capex_incl_fl | 45.2 | "Capital expenditures were $45.2 billion"（MSFT 逐字稿第 129 段）；xlsx〈CapEx〉表 Q1-27 = 45.2 | 屬實 | 「約四分之三為短期資產（CPU、GPU）」與「第三方產能交付延後」都有文件支持（第 129 段）。另有公司 xlsx 佐證，可以補為 second_source。 |
| 232 | MSFT FY27Q1 fcf | 18.633 | 58,000 − 39,367 = 18,633；"free cash flow was $18.6 billion"（逐字稿第 131 段） | 屬實 | 算術正確。引文出自逐字稿，但 url 欄寫的是 10-Q，10-Q 裡沒有 "free cash flow" 一詞；建議把逐字稿列為 second_source。 |
| 233 | MSFT FY27Q1 buyback_div | 11.136 | "Common stock repurchased (4,379)"；"cash dividends paid (6,757)"（10-Q 現金流量表 Financing） | 小誤〔只更正 detail〕 | 4.379 + 6.757 = 11.136，正確。detail 有兩處要改：(1) 10.2 不是出自新聞稿，而是出自 10-Q Part II Item 2 "We returned $10.2 billion"，以及逐字稿第 132 段；(2) 口徑差是可以解釋的：10.2 = 回購計畫內買回 2.925 + 宣告股利 7.279（xlsx〈Return of Cash to Shareholders〉Q1-27）。現金流量表的 4.379 另含員工代扣稅買回 1.5（Note 14），股利 6.757 是實際付現。建議 detail 改為「…（公司稱回饋 10.2＝計畫內買回 2.925＋宣告股利 7.279；現金口徑另含代扣稅買回 1.5）」。 |
| 234 | MSFT FY27Q1 bond_issue | 0.0 | 融資段只列 "Repayments of debt (4,850)"，沒有發債收入列（10-Q 現金流量表） | 屬實 | MD&A 也寫 "$4.9 billion increase in cash used for repayments of debt"。 |
| 235 | MSFT 2026-09-30 lease_not_commenced | 364.6 | "had not yet commenced of $364.6 billion"（10-Q Note 12 末段） | 屬實（detail 一處無法核對） | 「FY2027–FY2034 起租、租期 1–20 年」「主要資料中心」都相符。「6 月底 329.1」不在本次提供的文件中（應出自 FY26 10-K），無法核對。 |
| 236 | MSFT 2026-09-30 fin_lease_liab | 70.973 | "Total finance lease liabilities $70,973"；6/30 為 66,594（10-Q Note 12 補充資產負債表） | 屬實 | 「不在帳面借款內」有文件支持：融資租賃負債列在 Other current liabilities 5,005 與 Other long-term liabilities 65,968，不在 Long-term debt 31,090 之內。 |
| 237 | MSFT CY2026 capex_incl_fl | 170 | "spend should be approximately $170 billion due to the 3P capacity delays"（逐字稿第 173 段） | 小誤〔只更正 detail〕 | 值正確。另做了內部一致性檢查：xlsx 的 FY26Q3 31.9 + FY26Q4 41.0 + FY27Q1 45.2 = 118.1，170 − 118.1 = 51.9，與 Q2「over $50B」相符。detail 有兩處問題：(1)「原約 175」在本次文件中找不到，無法核對；(2) 逐字稿說第三方產能 "shifted to H2"，指的是財年下半年，也就是 2027 年 1–6 月，不是曆年下半年。建議改為「延到財年下半年（2027 年上半）」，並刪除或另外標注「原約 175」的來源。 |
| 238 | MSFT FY27Q2 capex_incl_fl | 50 | "In Q2, we expect CapEx spend will be over $50 billion"（逐字稿第 173 段） | 小誤〔只更正 detail〕 | 50 是下限，detail 已寫 over 50，可以接受，建議再註明「下限」。「下半年加速」的原文是 "second half of the fiscal year"，應改寫為「財年下半年」，以免誤讀為 2026 下半年。其餘驅動因素（CPU、GPU、記憶體與儲存成本、延後的 3P 產能）都相符。 |
| 239 | GOOGL 9M2026 ocf | 127.149 | "Net cash provided by operating activities … 127,149"（新聞稿 PDF 第 6 頁現金流量表，YTD 欄） | 屬實 | Q3 單季 42,290 正確。各季加總也吻合：Q1 45,790 + Q2 39,069 + Q3 42,290 = 127,149（PDF 第 8 頁）。 |
| 240 | GOOGL 9M2026 capex_cash | 131.953 | "Purchases of property and equipment … (131,953)"（PDF 第 6 頁） | 屬實 | Q3 51,355 正確。各季加總：35,674 + 44,924 + 51,355 = 131,953。 |
| 241 | GOOGL 9M2026 fcf | −4.804 | 127,149 − 131,953 = −4,804；"Free cash flow … $ (9,065)"（PDF 第 8 頁，Q3 欄） | 屬實 | 算術正確；用各季 FCF 加總 10,116 − 5,855 − 9,065 = −4,804 也吻合。新聞稿只給單季與 TTM，沒有 9M；9M 是推導值，標 [Derived] 正確。「TPU 庫存、歐盟罰款 5.2」出自逐字稿第 61 段 "$5.2 billion cash payment for a 2018 EC fine"，不是新聞稿，建議補 second_source。 |
| 242 | GOOGL 9M2026 buyback_div | 8.167 | "Repurchases of stock … 0"；"Dividend payments on preferred and common stock (8,167)"（PDF 第 6 頁） | 屬實 | 8.167 已含特別股股利，建議 detail 寫成「普通股＋特別股股利 8.167」。 |
| 243 | GOOGL 9M2026 bond_issue | 85.222 | "Proceeds from issuance of debt, net of costs … 85,222"（PDF 第 6 頁） | 屬實 | 這是扣除發行成本後的發債總收入，不是淨借款。還款 (8,879) 與 Q3 單季 28,996 都正確。新聞稿另寫 Q3 優先無擔保票據淨收入 28.7（PDF 第 4 頁），差額是其他債務。 |
| 244 | GOOGL 9M2026 equity_round | 61.933 | "Proceeds from issuance of common stock … 42,870"；"mandatory convertible preferred … 19,063"（PDF 第 6 頁） | 屬實 | 42,870 + 19,063 = 61,933，是現金流量表上的實收金額。Q3 單季普通股 12,371 ≈ ATM 4.3 + 8.1（PDF 第 4 頁）。強制轉換特別股在 Q3 單季為 0，全數發生在上半年。 |
| 245 | GOOGL CY2026 capex_cash | 200 | "full year 2026 CapEx guidance range of $195 billion to $205 billion remain unchanged"（GOOGL 逐字稿第 71 段） | 屬實 | 200 是區間中點。Alphabet 的 CapEx 等於 PP&E 購置（新聞稿 FCF 定義），所以掛 capex_cash 正確。「2027 significantly 增加、Q4 法說給細節」與第 71 段相符。 |
| 246 | GOOGL 2026-09-30 cloud_backlog | 561 | "Google Cloud's backlog was $561 billion in the third quarter"（逐字稿第 69 段） | 屬實 | 原文是 "just over 55%" 於 24 個月內認列，detail 寫「約 55%」可以接受。 |

## 二、MSFT 融資租賃口徑：7.4 與 5.192，以及 45.2 ≈ 39.4 ＋ 融資租賃 是否成立

**文件裡的數字**

| 數字 | 內容 | 出處 |
|---|---|---|
| 45.2 | 公司口徑資本支出，含融資租賃 | 逐字稿第 129 段；xlsx〈CapEx〉"Capital expenditures including assets acquired under capital leases" Q1-27 |
| 7.4 | "This quarter, finance leases were $7.4 billion. … we recognize the full principal value of finance leases in capex at the time of lease commencement." | 逐字稿第 130 段 |
| 39.4 | "cash paid for P, P, and E was $39.4 billion" | 逐字稿第 130 段；與 10-Q 39,367 一致 |
| 5,192 | 以租賃義務換得的融資租賃 ROU 資產 | 10-Q Note 12 補充現金流量表 |
| 82,712 → 89,779 | 融資租賃 PP&E 原值，增加 7,067 | 10-Q Note 12 補充資產負債表 |
| 66,594 → 70,973 | 融資租賃負債，增加 4,379；同期融資租賃本金付現 1,000 | 10-Q Note 12 |
| 26.7 → 26.5 | PP&E 應付帳款餘額 | 10-Q Note 6 |

**等式不成立**

| 組合 | 結果 | 與 45.2 的差距 |
|---|---|---|
| 39.4 + 7.4 | 46.8 | 多 1.6 |
| 39.367 + 5.192 | 44.559（四捨五入 44.6） | 少 0.6 |
| 推回的融資租賃（45.2 − 39.367） | 5.833 | 與 7.4、5.192 都不相等 |
| 改用應計基礎 PP&E（39.367 − 0.2 應付變動）約 39.17，再加 7.4 | 約 46.6 | 仍多約 1.4 |

去年同期也對不上：xlsx 的 Q1-26 公司口徑 34.9，而 19.394（現金）+ 9.147（ROU）= 28.5。可見「現金 PP&E ＋ 10-Q ROU 融資租賃」本來就不等於公司口徑 capex。

**可能的解釋（只能部分解釋）**

- **7.4 與 5.192 的差距**：CFO 的 7.4 是「起租時的全額本金價值」，屬於公司自訂的管理口徑。10-Q 的 5,192 是 GAAP 補充揭露，可能已扣除租賃修改、提前終止或重新衡量。用 10-Q 數字交叉驗證：負債增加 4,379 + 本金付現 1,000 ≈ 5,379，接近 5,192，代表 10-Q 內部是一致的。融資租賃資產原值增加 7,067，比較接近 7.4，差距可能來自預付租金或原始直接成本進入資產端、但沒有進入負債端；不過原值增加是扣除報廢後的淨數，無法確認這個推論。
- **45.2 與兩種加總都不吻合**：可能原因包括：公司口徑 capex 用應計基礎而不是現金基礎；7.4 與 45.2 的租賃計入時點不同；或逐字稿有轉錄誤差。文件都沒有說明，這一點列為**未解**。
- **巧合提醒**：10-Q MD&A 的 "Cash used in investing increased $7.4 billion" 與 CFO 的 7.4 金額相同，但內容無關，不要混用。

**建議**

- fin_lease_new 保留 **5.192**（10-Q GAAP 補充揭露、可以稽核，與負債變動自洽）。detail 可以補一句「CFO 稱 7.4 為起租全額本金口徑；10-Q 負債增 4.379＋本金付現 1.0≈5.4 支持 5.192」。
- capex_incl_fl 45.2 照公司數值使用即可，已有 xlsx 第二來源。
- 模型裡**不要**用「45.2 − 39.4」反推融資租賃，也**不要**用「39.4 + 7.4」重建 capex。

## 三、第二批：Amazon 2026Q3（F_funding 第 247–254 列）

- 核對者：獨立核對代理（未參與製作）；日期：2026-10-30
- 原文：Amazon Q3 2026 財報新聞稿 PDF（s2.q4cdn.com …/AMZN-Q3-2026-Earnings-Release.pdf，以 WebFetch 讀取，分兩次詢問現金流量表 3M／9M 欄與融資活動段）；IR 新聞稿頁（ir.aboutamazon.com …/Amazon-com-Announces-Third-Quarter-Results，讀取現金流量表 3M／9M／12M 六欄、FCF 調節表、資產負債表）。兩處數字一致。
- 原文位置寫法：PDF 經 WebFetch 轉文字後沒有頁碼，下表以「現金流量表〔欄〕」「補充現金流資訊」「FCF 調節表」「資產負債表」「正文 Guidance／FCF 段」標示位置。單位：原文百萬美元，CSV 為十億美元（÷1,000）。

| 列號 | 項目 | CSV 值 | 文件所載（≤25 字引文與位置） | 判定 | 備註 |
|---|---|---|---|---|---|
| 247 | 9M2026 ocf | 113.841 | "operating activities" 113,841（現金流量表 9M 2026） | 屬實 | Q3 單季 42,422、9M 2025 為 85,055，detail 兩數相符。標記 [Interested-party] 正確。 |
| 248 | 9M2026 capex_cash | 172.188 | "Purchases of property and equipment" (175,670)；"sales and incentives" 3,482（9M 2026） | 小誤（更正標記，值不變） | 重算 175.670 − 3.482 = 172.188，相符；Q3 單季購置 77,259 相符。但這是推導值，依規則應標 [Derived]（H1 同類第 70 列即標 [Derived]），CSV 標 [Interested-party]。 |
| 249 | 9M2026 fin_lease_new | 3.6 | "acquired under finance leases, net of remeasurements…" 3,600（補充現金流資訊 9M） | 屬實 | Q3 單季 1,472 相符；9M − Q3 = 2,128 與 H1 第 71 列相符。 |
| 250 | 9M2026 capex_incl_fl | 175.788 | 由 172,188 與 3,600 加總（無單一原文列） | 屬實 | 重算 172.188 + 3.600 = 175.788。標記 [Derived] 正確。 |
| 251 | 9M2026 fcf | -58.347 | "decreased to an outflow of $42.4 billion"（正文 FCF 段） | 屬實 | 重算 113.841 − 172.188 = −58.347。TTM：FCF 調節表 OCF 168,300 − 淨購置 210,657 = (42,357)，與 detail −42.357、公司 42.4 相符。標記 [Derived] 正確。 |
| 252 | 9M2026 buyback_div | 0.0 | 融資活動段僅有短債、長債、融資租賃本金、融資義務本金各列 | 屬實 | 融資活動段沒有買回或股利列，9M 填 0 有依據。引文是轉述，可接受。 |
| 253 | 9M2026 bond_issue | 115.1 | "Proceeds from long-term debt" 115,100（現金流量表 9M 2026） | 屬實 | Q3 單季 48,102、長債還款 (2,759)、短債收入 20,693、還款 (19,609)、9-30 長期借款帳面 173,125（資產負債表）全部相符。「含 3 月 53.8、7 月 25」原文沒有逐筆說明，屬合理推論，見下節。 |
| 254 | CY2026 capex_cash | 220 | "expect 2026 cash capital expenditures to be approximately $220 billion"（正文 Guidance 段） | 屬實 | 原文是 "We continue to expect"，detail 寫「不變」有依據。 |

### 重複計算檢查

**H1 ＋ Q3 單季 ＝ 9M（以新聞稿 Q3 單季欄核對）**

| 項目 | H1（既有列） | Q3 單季（新聞稿） | 加總 | 9M（新聞稿） | 結果 |
|---|---|---|---|---|---|
| OCF | 71.419（第 69 列） | 42.422 | 113.841 | 113.841 | 一致 |
| 資本支出淨額 | 96.31（第 70 列） | 77.259 − 1.381 = 75.878 | 172.188 | 175.670 − 3.482 = 172.188 | 一致 |
| 融資租賃取得 | 2.128（第 71 列） | 1.472 | 3.600 | 3.600 | 一致 |
| FCF | −24.891（第 72 列） | 42.422 − 75.878 = −33.456 | −58.347 | −58.347 | 一致 |

另外，FCF 調節表的 TTM 至 2026-06 為 (7,604)，與第 72 列 detail「−7.6」相符。

**長期借款收入 115.1 是否已涵蓋 3 月與 7 月發債**

- 9M 115,100 減去 Q3 48,102，得 H1 = 66,998，足以涵蓋 3 月的 53.8（Q1）。Q3 的 48,102 足以涵蓋 7 月 8 日的 25.0。因此 115.1 已含第 74、75 列，兩者**不可再與 115.1 相加**。
- 2025-11 的 15.0（第 73 列）落在 2025Q4：12M 130,027 − 9M 115,100 = 14,927 ≈ 15.0。它不在 9M2026 範圍內，與 115.1 不重複。
- 待查（原文未解釋）：Q3 長債收入 48.1 比 7 月 25.0 多約 23.1；H1 的 67.0 比 3 月 53.8 多約 13.2。第 75 列 detail 寫「公司稱 2026 不再發債」，但 Q3 仍有約 23 的其他長期借款收入，可能是定期貸款、其他幣別債或其他借款。Q3 新聞稿沒有明細，需要看 10-Q 債務附註。另外，現金收入是扣除折價與費用後的淨額，第 74 列的 EUR 部分也有匯率換算，所以逐筆金額本來就不會與現金流量表完全相等。

### 取代建議

| 既有列 | 內容 | 建議 |
|---|---|---|
| 69 | Amazon 2026H1 ocf 71.419 | 標「由 Q 段取代」（第 247 列） |
| 70 | Amazon 2026H1 capex_cash 96.31 | 標「由 Q 段取代」（第 248 列） |
| 71 | Amazon 2026H1 fin_lease_new 2.128 | 標「由 Q 段取代」（第 249 列） |
| 72 | Amazon 2026H1 fcf −24.891 | 標「由 Q 段取代」（第 251 列） |
| 74 | Amazon 2026-03 bond_issue 53.8 | 標「由 Q 段取代」（已含於第 253 列 115.1）；可保留作事件紀錄，但不計入加總 |
| 75 | Amazon 2026-07-08 bond_issue 25.0 | 同上 |
| 79 | Amazon CY2026 capex_cash 220（A_us_cloud 既有） | 標「由 Q 段取代」（第 254 列，來源較新、值相同） |
| 73 | Amazon 2025-11 bond_issue 15.0 | 不取代（屬 2025Q4，不在 9M2026 範圍內） |

### 判定統計

共 8 列：屬實 7、小誤 1（第 248 列，標記 [Interested-party] 應改為 [Derived]）、錯誤 0、無法核對 0。

## 四、第三批：Amazon、Meta 2026Q3 逐字稿（F_funding 第 255–257 列）

- 核對者：獨立核對代理（未參與製作）；日期：2026-10-30
- 原文：Amazon 2026Q3 法說逐字稿（/tmp/amzn_tr.txt，非空白段 158 段）、Meta 2026Q3 法說逐字稿（/tmp/meta_tr.txt，非空白段 221 段），都由 Andy 下載的 docx 轉成純文字。段落位置以 python 去掉空白行後重新編號（0 起算），沒有沿用 CSV 引文。
- 通話日期：Amazon 第 7 段寫 "as of today, October 29, 2026"，與第 255、256 列 doc_date 2026-10-29 相符。Meta 逐字稿全文**沒有通話日期**：第 9 段只寫 "assumptions as of today"，全文唯一的日期線索是第 45 段「August 2026」、第 125 段 "Earlier this month"、第 179 段 "late September"、第 75 段 "late January 2027"。第 257 列 doc_date 留空、來源欄註「逐字稿未載日期」，處理正確。

| 列號 | 項目 | CSV 值 | 文件所載（≤25 字引文與段落） | 判定 | 備註 |
|---|---|---|---|---|---|
| 255 | Amazon 2026-09-30 cloud_backlog | 578 | "Backlog is $578 billion, growing triple digits"（Amazon 第 13 段） | 屬實 | 「年增三位數、分散」都有原文（同段 "well-diversified"）。Q3 現金資本支出 75.9 出自第 61 段 "cash CapEx, which was $75.9 billion in Q3"，CSV 引文沒有列這段，建議補上。77.259 − 1.381 = 75.878 ≈ 75.9，算術正確。標記 [Interested-party] 正確。 |
| 256 | Amazon CY2027 capex_cash | 220 | "expect to spend meaningfully more in 2027"（Amazon 第 31 段） | 小誤（更正值與 detail） | 引文與段落相符，同段有 "still expect 2026 cash CapEx of approximately $220 billion"。AWS 成長：第 33 段 "grow north of 40% in each of 2027 and in 2028"，第 79 段也有相同說法，相符。容量：第 89 段 "hundreds of megawatts or sometimes up to a gigawatt"；"very good visibility into 2027 and 2028 and some of 2029"。detail 寫「2027–2029 能見度高」，對 2029 說得太滿（原文是 "some of 2029"）。值 220 的寫法容易被誤讀，見下節。 |
| 257 | Meta 2026-09-30 net_debt_change | 83.7 | "$82.4 billion in cash … $83.7 billion in debt"（Meta 第 47 段） | 小誤（更正 item 與 detail；值不變） | 83.7、82.4、"including principal payments on finance leases, were $34.9 billion"、"Free cash flow was negative $2.1 billion" 都在第 47 段，相符；FCF −2.1 與 F#95 的「Q3 −2.12」一致。無表決權股：第 73 段 "today we announced plans to create a new class of non-voting capital stock"，第 127 段 "usable for capital raising, employee compensation, acquisitions"、"do not have any specific financing plans to announce"，相符。2027 資本支出：第 142 段 "until we get to the Q4 call"，相符（另見第 71 段 "still developing our budget and capacity plans for next year"）。問題有三：①值 83.7 是**存量**，item 卻寫成 net_debt_change（流量），見下節；②「含短期」逐字稿沒有說，屬推論，無法由本文件核對；③「無具體融資計畫」只是通話當時的說法，見下節的時序分析。 |

### 第 256 列：「220 只作下限」是否會被誤讀

- **會被誤讀。** `tools/import_funding.py` 的 `ym()` 把 CY2027 解析成 2027-01～2027-12，因此這列會落在 post 區（期後事件與指引）。讀者在 post 區會看到「Amazon CY2027 現金資本支出 220」，與第 254 列「CY2026 220」數字相同，很容易讀成「2027 持平」，意思正好與原文 "meaningfully more" 相反。下限的意思只寫在 detail 括號內，任何只讀 value 欄的程式或圖表都看不到。
- 第 238 列也用「over 50（下限）」的寫法，但那是公司**自己給的數字**的下限；第 256 列的 220 是 2026 的指引，被借來當 2027 的下限，性質不同。
- **建議**（擇一，以甲為優先）：
  - 甲、仿照第 182 列，value 改為「（空白，未量化）」，detail 改為：「指引：2027 現金資本支出將 meaningfully more（較 2026 約 220 明顯增加，未給數字；> 220 為推論下限，不入數值）；另稱 AWS 2027、2028 各年營收成長可望 40% 以上；對 2027、2028 及部分 2029 的大塊容量（數百 MW，有時達 1 GW）能見度很好」。
  - 乙、若一定要放數值，item 改為 `capex_cash_floor`（不在 FITEMS 內，不會被加總），value 220，detail 開頭寫「下限（非估計）：> 2026 指引 220」。

### 第 257 列：item、淨負債算術與第 100 列

- **算術**：83.7 − 82.4 = 1.3，「淨負債約 1.3」正確。
- **與第 100 列比較**：第 100 列以長期借款 80.922 對照現金約 82，得到淨現金約 +1.1（「約 0」）；第 257 列以 83.7 對照 82.4，得到淨負債約 1.3。兩者都是「約 0」，量級上不矛盾，但**正負號相反**。差異來自債務口徑：83.7 − 80.922 = 2.778。CSV 把這個差額解釋為「含短期」，但逐字稿沒有說明 83.7 的組成，需要以 10-Q 資產負債表（短期債或一年內到期長期債）核對。在核對之前，建議 detail 寫「口徑差 2.8 待 10-Q 核對」，不寫「含短期」。
- **item 錯誤會造成重複加總**：83.7 是 9-30 的債務存量，不是變化量。period 2026-09-30 經 `ym()` 解析為 2026-09，會落在 Meta 2026 區（9M2026，2026-01～2026-09）。一旦採用，Funding 表「淨負債變化」欄會把第 100 列的 22.178 與這列的 83.7 加總成 105.9，數字錯誤。目前因為未判定所以不採用，但採用前必須先更正 item。建議 item 改為 `debt_level`（或與既有的 `cash_level` 並列的存量 item；兩者都不在 FITEMS 內，不會被加總）。
- **建議的 detail 全文**：「9-30 債務 83.7（CFO 口徑；10-Q 長期借款 80.922 見 F#100，差 2.8 的組成待 10-Q 核對）；現金與有價證券 82.4 → 淨負債約 1.3（F#100 口徑為淨現金約 1.1，兩者都約 0）；Q3 單季含融資租賃本金資本支出 34.9、FCF −2.1；宣布新設無表決權股（股票股利 1:1，預計 2027-01 下旬交易），用途含募資、員工薪酬、併購；通話當時稱『目前無具體融資計畫』，其後 10-26 FWP 即定價 37.5（F#106、F#107）；2027 資本支出指引待 Q4 法說」。引文欄建議補上「today we announced plans to create a new class（第 73 段）；until we get to the Q4 call（第 142 段）」。

### Meta 逐字稿與 F#106、F#107（2026-10-26 FWP）的時間先後

- 逐字稿沒有日期，所以無法直接排序，只能用內容推算。
- FWP（F#107）已提到「2027Q1 起最多 60 的 Class C ATM」。Class C 就是逐字稿第 73 段「today we announced」的新股類別，因此 FWP 不可能早於宣布日。另外，F#97 的 2026 資本支出指引 135–145 也以 d116268dfwp 為來源，與逐字稿第 69 段 "narrowed from our prior outlook of $130 billion-$145 billion" 相同，可見 10-26 時 Q3 指引已經公開。
- 逐字稿第 127 段說 "do not have any specific financing plans to announce at this time"。這句話只有在通話**早於** FWP 啟動與定價時才說得通（定價結果：遠期出售 25、特別股 12.5，合計 37.5）。
- **結論**：時序一致，條件是 Meta 法說日**不晚於 2026-10-26**，而且在募資啟動之前。如果法說其實在 10-26 之後（例如與 10-Q 同為 10-29），「今天宣布新股類別」和「無具體融資計畫」都會與已定價的 37.5 矛盾，那就要懷疑這份逐字稿的版本或來源。另外要注意：F#92–100 的 10-Q doc_date 是 10-29，F#95 的第二來源 beancount 網址日期是 2026/10/21。法說、新聞稿、10-Q 的先後目前無法由現有文件確認。
- 建議：第 257 列 doc_date 維持空白，detail 或來源欄加註「通話日期未載；依 FWP 10-26 已引用 Class C，推定不晚於 2026-10-26」。請 Andy 由 Meta IR 新聞稿或 8-K（Item 2.02）的日期補上確切日期。本次核對無法取得 FWP 原文，SEC 網址的讀取請求逾時，所以 F#106、F#107 的日期沿用 G 段已核對的 2026-10-26。

### 判定統計

共 3 列：屬實 1（第 255 列）、小誤 2（第 256 列：值 220 改為空白或改 item；detail「2027–2029」改為「2027、2028 及部分 2029」。第 257 列：item 改為存量 item 以免與 F#100 重複加總；「含短期」改為「待 10-Q 核對」；補註融資時序）、錯誤 0、無法核對 0（第 257 列「含短期」一處由逐字稿無法核對）。

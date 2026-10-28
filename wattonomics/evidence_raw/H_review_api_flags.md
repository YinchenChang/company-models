# H 段：API 回查旗標（不符 5／口徑不同 41）人工複核｜2026-10-28

- **核對者：** 未參與製作的代理。
- **核對範圍：** `tools/wattonomics_api/evidence_audit/out/full_20261028.csv` 中判定為「不符」或「口徑不同」的 46 列。
- **抓取方式：** 每列都用 WebFetch 打開原文，只用 CSV 中出現過的網址（含 second_source_url），沒有自行拼網址。
- **被擋的網址（未繞過）：** cryptobriefing.com（A#56，client_error）。A#56 改用第二來源 theregister 核對。
- **Excel 依據：** `wattonomics/model/20261028_Wattonomics_v0.1.1.xlsx`，看 In 頁與 FundRaw 頁。
  - FundRaw 的「採用」欄（O）為 1，才會進 Funding 頁。
  - Funding 頁 S 欄（供應商融資／策略投資）沒有任何公式引用，只供顯示，不進任何比例。
  - Funding 頁 N 欄（淨負債變化）也沒有下游引用。
- **列號：** 檔案列號，表頭算第 1 行，與 API 輸出相同。A＝A_us_cloud.csv，B＝B_labs_china_totals.csv，D＝B_delay.csv，F＝F_funding.csv，P＝P_facility_power.csv。
- **本段沒有改 repo 任何檔案。**

## 0. 統計

| 判定 | 列數 | 列 |
|---|---|---|
| 維持 | 6 | A#7、A#29、D#5、B#20、F#227、P#44 |
| 更正數值 | 8 | A#18、D#12、B#37、B#64、F#179、F#182、P#50、P#68 |
| 更正口徑或標記 | 32 | 其餘各列 |
| 刪除 | 0 | — |

### API 判定的準確度

- API 判定基本正確：39 列。
- API 判定部分錯誤：7 列。
  - **B#20：** 8-K 寫的是 "through May 31, 2031"，所以「2029–2031」有依據。
  - **B#34：** Colossus II 在第二來源 Yahoo（Reuters）裡有寫。
  - **P#44：** 2026-04-13 在第二來源有寫。
  - **A#19、B#33：** 1 GW 確實是「額外」的，但它本身也是上限（"up to"）。
  - **A#7：** 只是措辭問題。
  - **F#199：** 判定口徑不同是對的，但沒有指出真正的問題：這筆是採購，不是融資。

### 與先前人工複核衝突、經打開原文確認後的結論

| 列 | 先前判定 | API 判定 | 誰對 | 依據 |
|---|---|---|---|---|
| A#18 Abilene 900 MW | C 段：屬實（「含現場電廠」） | 口徑不同 | **API** | W.Media 引 Crusoe 新聞稿與 CEO 的說法："a dedicated 900 MW on-site power plant"、"900 MW of new on-site power generation"。只有記者導言寫 "Microsoft will now add some 900 MW of capacity"。第二來源（Shopifreaks 引 Bloomberg 2026-03-24）："Microsoft has agreed to lease a 700-megawatt data center"。 |
| A#19／B#33 Azure 1 GW | C 段：A#19 小誤，寫「全案上限就是 1 GW」；B#33 屬實 | 口徑不同（「額外」選擇權） | **API 對了大半，C 段錯** | Anthropic 原文："committed to purchase $30 billion of Azure compute capacity **and to contract additional** compute capacity up to one gigawatt"。NVIDIA 段寫 "initially be up to one gigawatt"。所以 1 GW 是 $30B 以外的額外承諾，而且是上限（up to），不是確定量。CSV 寫「1000 上限，另有選擇權再 1 GW」是重複計算；C 段說「全案上限 1 GW」則把 $30B 錯併進 1 GW 裡。原文也沒有寫 $30B 對應幾 GW。 |
| A#37 | C 段：屬實 | 口徑不同 | **API** | Fortune："AWS remains on pace to double its power capacity…, Jassy said"，是轉述，不是引語。範圍是 AWS，不是全 Amazon。 |
| B#64 | C 段：屬實 | 口徑不同 | **API** | Epoch：16M 是「一季落後」情境，兩季落後是 12M；"tripled in 2025" 指的是 2025 年，不是 2026 年。 |
| F#199 | G 段：小誤，刪除「至 2029-05」，理由是不在來源 | 口徑不同 | **G 段錯在這一點** | Yahoo（Reuters 2026-05-20）原文："agreed to pay it $1.25 billion per month through May 2029"。另有 "fees would be reduced during the capacity ramp-up"，所以年化 15 是 [Derived] 上限。 |
| D#23 | C 段：屬實（引述） | 口徑不同 | **兩者都對** | 引述屬實；但 6／9／12 月是由 "several quarters" 推出來的，tag 應標 [Derived]。 |
| P#5、P#27、P#29 | G 段：屬實 | 口徑不同 | **API** | 這三列的數值屬實，但 basis 或區間是自己加上去的：原文沒有寫 per IT MW；13／15、0.8／1.2、0.5 這些數字原文都沒有。 |
| P#44 | G 段：屬實 | 口徑不同（as_of） | **G 段** | 第二來源 jaredwatkins 寫 "announced… on April 13, 2026"。只有 tag 要從 Bloom 本人說改成「報導（IWR）」。 |

## 1. 承重列

| 檔案:列 | API | 我的判定 | 正確值／口徑 | Excel 是否採用 | 對 Excel 的影響 |
|---|---|---|---|---|---|
| A#18 Microsoft Abilene | 口徑不同 | **更正數值** | **700 MW**（Bloomberg：Microsoft 承租，口徑未說明）。900 MW 是 Crusoe 現場發電廠容量（發電側），不是 Microsoft 的容量。全站 2.1 GW、首棟 2027 年中屬實。metric 改為 contracted（報導）。 | **是**：In T2 `cp_msft_abl`（In!J189＝900，口徑「未說明」）→ Pipeline A8 | **In!J189：900 → 700**。G189 改為：「Bloomberg 700 MW 承租（口徑未說明）；900 MW 為 Crusoe 現場電廠」。2027 年 Microsoft 新增少 0.2 GW-IT（另讀為設施時少 0.17）。 |
| A#70 Oracle Abilene 擴建取消 | 口徑不同 | 更正口徑 | TechRadar：原計畫把園區擴到約 2 GW，目前上限 1.2 GW，所以取消的增量約 0.8 GW，不是「2 GW 擴建」。B#8 引 Bloomberg 寫 600 MW，兩個數字要並列。expected_online「Microsoft 承租 900 MW」改為 700 MW（第二來源 Bloomberg）。as_of 改 2026-03（TechRadar 2026-03-09）。原因：融資與電力延誤屬實。 | 否 | 無 |
| A#19 Anthropic×Azure | 口徑不同 | 更正口徑 | 拆成兩部分：$30B Azure 確定採購（MW 未揭露），加上**額外**至多 1 GW（上限，NVIDIA GB／VR）。value 1000 保留，但 unit 改成「MW（額外，上限）」，metric 改成 option／planned。刪除「另有選擇權再 1 GW」。tag 標第三方彙整（ComputePrices）；一手頁現在已可讀。 | 間接：In!I57 `ctr_msft_ant`＝1（S4 對照，0 個公式引用） | 值不變。**In!B57 改名**為「Anthropic–Azure：$30B 確定（GW 未揭露）＋額外至多 1 GW」；G57 刪除「C 段更正：上限 1 GW」。不影響計算。 |
| B#33 同上 | 口徑不同 | 更正口徑 | 同 A#19。site_or_contract 改為「Azure $30B ＋ 額外 NVIDIA GB／VR 至多 1 GW」。expected_online「2026-11 至 2033-05 付款」和 $31.4B 的出處是 InvestmentNews（2026-09-29，屬實），應移到 second_source。 | 同上 | 同上 |
| A#29 Anthropic×Google TPU | 口徑不同 | **維持** | 原文："well over a gigawatt of capacity online in 2026"、"up to one million TPUs"。1000 MW 只能當下限讀，CSV 的 unit 已寫「下限」，不必改。 | 是：In!L172:M172 `al_googl_ant`＝1 GW；In!I58 `ctr_googl_ant`＝1 | 無（保守下限）。G172 可補一句「>1 GW 的下限」。 |
| A#37 Amazon 容量翻倍 | 口徑不同 | 更正口徑 | site_or_scope 改為「**AWS** 電力容量目標」。quote 改為「記者轉述 Jassy（非逐字）」。 | 否（ai_amzn 用的是 AWS 資本支出占比，與此列無關） | 無 |
| B#16 Ports Pike | 口徑不同 | 更正口徑 | metric 改為「首期（NVIDIA 信用支持）」，不是 contracted。4.25 IT-GW 屬實。OpenAI 是全園 8 IT-GW 的客戶，20 年租約；SB Energy S-1（經 ledger 轉述）稱 OpenAI 約 8.0 GW-IT 是 binding lease。3.75 IT-GW 是 NVIDIA 的選擇權。新聞稿沒有「首期 800MW」。 | 是：In T2 `cp_oai_pp1–3`（1417＋1417＋1416，2028–30） | 數值不變（4.25 是實體首期，三年均分 [Assumed]）。G210:G212 補「OpenAI 租約 8.0 GW-IT；4.25 為 NVIDIA 信用支持首期；未簽 3.75 屬 NVIDIA 選擇權」。 |
| B#64 Epoch 運行中算力 | 口徑不同 | **更正數值** | 區間 12–16M H100e：1,600 萬是「一季落後」，1,200 萬是「兩季落後」。value 建議寫 1,600（高）／1,200（低）。expected_online「2026 年約 3 倍」改為「2025 年約 3 倍（tripled in 2025）」。 | 否（tp_epoch 用的是 B#63 的 30 GW） | 無 |
| F#179 Stargate 股本 | 不符 | **更正數值** | **52 → 45**。19＋19＋約 7＝45（Seeking Alpha 引 The Information："roughly $45B"）。DCD 的 55 是首期 100 中尚缺的金額。持股：DCD 寫 OpenAI 40%，SA 寫 OpenAI＋SoftBank 合計約 40%，兩者衝突，要並列。DCD 沒有日期，SA 為 2025-01-23。 | 否（O181＝0） | 無。F.md 第 191 行、第 330 行的「52」要改成 45。 |
| F#196 Anthropic 淨負債變化 | 口徑不同 | 更正口徑 | 原文只有期末現金與短投 **+20.28**（存量），沒有負債數字。item 不能寫 net_debt_change，value 也不能是 −20.28。改 item＝cash_level、value＝20.28，或刪除。 | 否（O198＝0） | 無。若日後採用，不可放進 Funding 頁 N 欄。 |
| F#226 NVIDIA 51.157 | 口徑不同 | 更正口徑 | 51.157 是 NVIDIA 2026-07-26 的**非上市證券帳面存量**，屬於投資組合，不等於供應商融資流量。其中非上市權益 47.898，上市權益 42.783。detail「非上市 51.2＋上市 42.8」是兩個不同口徑相加（＝93.9，不是 99）。item 改為 investment_stock_nonmkt。Artiverse 自稱內容多由 AI 生成，應改看 10-Q。 | 否（O228＝0） | 無。**連帶問題：** F#222（43.364，同為非上市權益存量）已採用，O224＝1，桶 other、vendor_financing，口徑同樣錯。建議改 item 或改採用＝0。Funding 的 S 欄沒有下游引用，所以只影響顯示。 |
| F#199 Anthropic→SpaceX | 口徑不同 | 更正口徑 | 1.25／月**至 2029-05**屬實（Reuters 經 Yahoo），90 天可終止，爬坡期費用減免。年化 15 是 [Derived] 上限。關鍵：這是 Anthropic **支付**的算力採購，不是 Anthropic 取得的融資，所以不應放進 vendor_financing。範圍為 Colossus 與 Colossus II。tag 改 [Derived]（Reuters 轉述 S-1）。 | **是**（O201＝1，ant、2026、vendor_financing） | **FundRaw O201：1 → 0**（或 item 改為 compute_purchase）。Funding!S30：57 → 42。S 欄沒有下游引用，比例不變。N201 的 G 段更正「刪除至 2029-05」要撤回。 |
| F#182 AMD 認股權證 | 口徑不同 | 更正數值 | value 0.0 → 空白（未量化）。歸屬條件：1 GW 起，最多到 6 GW，另綁股價目標與技術商業里程碑。tag [Verified] → [Interested-party]（AMD 新聞稿）。 | 否（O184＝0） | 無 |

## 2. B_delay（D1／D2 依據）

| 檔案:列 | API | 我的判定 | 正確值／口徑 | Excel 是否採用 | 對 Excel 的影響 |
|---|---|---|---|---|---|
| D#5 發電併網 | 口徑不同 | 維持（補註） | LBNL：2018–24 年那批案子 >4 年，圖上標約 4.3 年＝52 月；2025 年完工案 >5 年＝60 月（下限）；PJM 2025 年進場案平均 >7 年＝84 月，與 LBNL 是不同統計。這些都是**發電側**數字。tag 改為「LBNL（第三方轉述）；84 為 PJM 平均」。 | 是：In!G24 `dly_d2_bind` 依據（「併網 48–60 月」） | 無。值 0.3／0.4／0.5 不變。D2 的依據仍成立：D#2、D#3 大負載併網 48 月以上，C 段判屬實。 |
| D#10 燃機全產業 | 口徑不同 | 更正口徑 | 60、84 屬實（EPRI，2025 年平均 5 年、部分 7 年）。low 36 是「兩年前最強機型最多約 3 年」，屬歷史值，應註明；或改引 D#7（廠商新單約 3 年）當作現行低值。metric 改為「最強機型等待（EPRI）」。 | 否（D2 引的是 D#7、D#8） | 無 |
| D#12 電力變壓器 | 不符 | 更正數值 | 中值 30（128 週）屬實。原文沒有 24 和 60，也沒有「2020 年前 24–30 個月」。原文唯一的低值是經紀商自稱 12–14 個月（Interested-party）。最高的是 GSU 144 週，約 33 月。建議 low 24 改標 [Assumed]，**high 60 → 36**（[Assumed]，略高於 GSU 33）。note 刪除「2020 年前 24–30 個月」。 | 否（D2 引的是 D#11 GSU） | 無 |
| D#18 Denton | 不符 | 更正口徑或標記 | DCD 頁面沒有 MW、原指引、年中全園區這些內容，也沒有月數；只寫「第三方開發商落後、客戶同意修改合約日期」，以及 12 月底已有 >16,000 GPU。as_of 不是 2026-03-02，頁面沒有日期。月數 3／6／6 改標 [Assumed]，或留空。 | 是：In!G23 `dly_d1_m` 依據（已註「無法核對」） | 值不變。**G23 刪除 Denton**，改為「D#16 Oracle 部分站 12、D#19 Nebius NJ 6–7；low 3 [Assumed]」。 |
| D#19 Nebius NJ | 口徑不同 | 更正口徑 | metric「首期 50MW 工期」→「首期 50MW 延遲（實際 10–11 月 − 目標 4 月）」。6／7／7 是延遲月數，數值本身正確。供應商延誤造成 4→6 月。 | 是：In!G23 依據 | 無。仍支持 D1 基準 6 月。 |
| D#21 UAE 第一期 | 口徑不同 | 更正標記 | metric 改為「Q3 2026 目標，至 9-12 無試運轉公告」。原文明說**不視為延遲**；Reuters 分散多站指的是 5 GW 園區，第一期繼續施工（有修改）；Khazna 稱 2026 年內交付。月數 1／3／6 標 [Assumed]。 | In!G224 `cpx_sov_uae` 註（納入＝0） | 無。G224「D#21 未見上線」改為「至 9-12 無試運轉公告；官方未改期」。 |
| D#23 AWS | 口徑不同 | 更正標記 | 原話是 "several quarters"，對 "a couple of quarters" 的回應。6／9／12 是 [Derived]（2–4 季），tag 要改，不能標成本人說的月數。 | 否 | 無 |

### D1／D2 依據結論

- **數值不需要改：**
  - D1：比例 0.25／0.35／0.45，月數 3／6／12。
  - D2：比例 0.3／0.4／0.5，再延 12 月。
- **只需要修 In!G23 的說明：**
  - 刪除 Denton。
  - 月數依據改為「D#16 Oracle 部分站 12、D#19 Nebius NJ 延遲 6–7（實際 10–11 對目標 4）；low 3 [Assumed]」。
- **In!G24 可以補一句：** D#5 是發電側併網，大負載併網看 D#2、D#3。

## 3. P 段（per IT MW 口徑等）

| 檔案:列 | API | 我的判定 | 正確值／口徑 | Excel 是否採用 | 影響 |
|---|---|---|---|---|---|
| P#2 JLL | 口徑不同 | 更正口徑 | 11.3、>15、+6% 屬實；10.7 是 [Derived]。basis「per IT MW」→「未說明」，metric「shell-and-core」→「整體建造成本（範圍未說明）」，quote 刪除「不含 IT」。 | 否（12.67 拆分用 McKinsey 比例） | 無 |
| P#3 T&T | 口徑不同 | 更正口徑 | 12.08、14–16、約貴 10% 屬實。basis 改為「未說明（典型 60MW 設計）」。 | 否 | 無 |
| P#5 Epoch 經 CorAdvisors | 口徑不同 | 更正口徑 | 14.1 [Derived] 算對。basis 改為「未說明（1 GW 園區）」。low 13、high 15 標 [Assumed]。as_of／doc_date 改 2026-07（原文 2026-07-24）。 | 否（In!G69 冷卻寫的「Epoch 約 1.9」與此同源，數值正確） | 無 |
| P#16 Axis 電氣 | 口徑不同 | 更正口徑 | 4.8–7.2、40–48% 屬實。basis 改為「未說明（建造預算每 MW；100MW 設施）」。quote 刪除「作者自稱與總數未對齊」。 | 否（Interface 的 "AP16" 是誤匹配，不是 P#16） | 無 |
| P#17 Axis 冷卻 | 口徑不同 | 更正口徑 | 3.5–5.0、29–33% 屬實；basis 同 P#16。 | 否（In!G69 的高值 3.5 引 G 段，不是 P#17） | 無 |
| P#18 Axis 殼體 | 口徑不同 | 更正口徑 | 1.8–2.3、約 15% 屬實；basis 同 P#16。 | 否 | 無 |
| P#27 genset 單價 | 口徑不同 | 更正標記 | base 1.0（約 $1,000/kW，NREL 2025-06）屬實。0.8／1.2 是自設區間，應標 [Assumed]。第二來源 Groww（2026-08-01）："2-3 MW each… $500,000 and $2 million"，換算 0.17–1.0 $M/MW。這是單機設備價，不含安裝，可作為低值佐證。 | **是**：In!H67:J67 `fac_gen_px` 0.8／1.0／1.2 | 基準不變。建議 **In!H67 低值：0.8 → 0.5**（Groww 設備價下緣加上安裝）；E67 改為「Analogy（基準）／Assumed（區間）」。只在 fac_gen_px 單獨取低值時，影響 Facility 的 genset 列和其他專業服務列；拆分合計仍對齊 TK。屬選擇性修改。 |
| P#29 Fairwater 不設 genset | 口徑不同 | 更正標記 | 原文："forgo traditional backup infrastructure like on-site generators for the GPU fleet"、"99.99% uptime"。high 0.5 是 [Assumed]（非 GPU 負載仍可能設發電機，原文沒說）。tag 改為記者轉述 Microsoft。 | 是：In!H66 `fac_gen_mult` 低值 0 | 無 |
| P#44 Bloom×Oracle | 口徑不同 | 維持（tag 小修） | 1.2 GW 確定、最多 2.8 GW 屬實（IWR 2026-10-23）。2026-04-13 是公告日（第二來源 jaredwatkins）。tag 改為「報導（IWR）」。 | 否 | 無 |
| P#50 主變 | 口徑不同 | **更正數值** | ModulEdge 原文 60 MW 是 **IT 負載**。設計負載 76.5／88.4 MVA。一台退出仍符合需求的組態是 4×40（160 MVA）→ 2.67 MVA／IT MW，或 3×63（189 MVA）→ 3.15。3×40＝120 MVA 不符一台退出（88.4 的情境），所以 base 2.0 沒有依據。ComEd：225 MW 是全場需求（設施），首批 4×100＝400 MVA；÷（225÷1.2）IT ≈ 2.13。建議 **low 1.8 → 2.1、base 2.0 → 2.67、high 3.2 → 3.15**。 | **是**：In!H80:J80 `xfmr_gva` | **In!H80:J80：1.8／2.0／3.2 → 2.1／2.67／3.15**。只影響 Facility!D22:K22 的實物量顯示，沒有下游公式。A22 的標籤「判斷值；P50 例子不支持此基準」可以拿掉。 |
| P#68 台達電 | 不符 | 更正數值 | 「今年 AI 相關營收占比將超過 25%，高於去年的 18%」。**low 18 → 25**：18 是 2025 年實績，不能當 2026 年的低值。base 25 是下限；high 改標 [Assumed]，例如 28。 | 否（全活頁簿查無引用） | 無 |

## 4. 其餘各列

| 檔案:列 | API | 我的判定 | 正確值／口徑 | Excel 是否採用 | 影響 |
|---|---|---|---|---|---|
| A#7 Microsoft 12 GW | 口徑不同 | 維持 | 原文："owned-and-leased… about 12 gigawatts today"、">38 GW by 2032"、"excludes capacity rented from neoclouds"，出處為 Bloomberg reports，另有 "Sources cautioned"。「知情人士」一詞換成「Bloomberg 報導（消息來源）」即可。是否為全部負載，原文沒定義，basis unknown 是對的。 | 否：In!I60 `tgt_msft_32`＝38 只作 S4 對照，0 個引用 | 無 |
| A#56 Meta Prometheus | 口徑不同 | 更正標記 | 主來源 cryptobriefing 被擋，**無法核對**。第二來源 Register（2025-07-15）：Zuckerberg "coming online in '26" 屬實；Prometheus 1 GW 是 SemiAnalysis 估計，Zuckerberg 沒有說 1 GW。「後為 Meta 自述目標」沒有核到。 | 否（Excel 用 Epoch 的 cp_meta_pro1–4） | 無 |
| A#61 Meta–Nebius | 口徑不同 | 更正標記 | $12B 專用（Vera Rubin，2027 年初起）＋最多 $15B 額外容量（合計最多 $27B），5 年，屬實。as_of 2026-03-16 → 2026-03（頁面沒有日期）。tag 改為第三方彙整（引 Nebius 新聞稿）。 | 否 | 無 |
| B#4 OpenAI 2027 | 口徑不同 | 更正標記 | Epoch（2026-05-20）："low double-digit GW in 2027, up from 1.9 GW in 2025"。11 是 [Assumed] 中點。30 GW（2030）出自 taekim 備忘錄（2026-04-09，memo 為「Wednesday」）。as_of 改為：備忘錄 2026-04-08，Epoch 2026-05-20。tag 去掉 Bloomberg。 | 是：In!L156 `usr_oai` 2027＝11（只作 Matrix 對照） | 值不變；G156「B#4 中點，屬實」改為「low double-digit 屬實；11 為 [Assumed] 中點」。 |
| B#8 Abilene 600MW | 口徑不同 | 更正標記 | ledger 引 Bloomberg：600 MW 擴建因融資談判停滯而取消（2026-03）。「本體 1.2 GW 不受影響」出自 TechRadar（"top out at 1.2GW"），不是 ledger，也不是記者推論。600 與 TechRadar 推得的約 800 有出入，要並列。 | 否 | 無 |
| B#20 CoreWeave×OpenAI | 口徑不同 | 維持（tag 小修） | 三筆合計 22.4 屬實。8-K 原文："$6.5 billion through May 31, 2031"；$4B 至 2029；$11.9B 為 5 年。所以「至 2029–2031」**有依據**，API 這點錯了。tag：只有 6.5 是 8-K Verified，其餘來自 ledger 轉述。delay_note 的 Denton 改引 D#18（無法核對）。 | 否 | 無 |
| B#27 UAE 第一期 | 口徑不同 | 更正口徑 | delay_note 改為：Reuters 分散多站指 5 GW 園區，第一期繼續施工（有修改）；G42 稱進度如期。「Epoch 預測 2026-12-01」兩個來源都沒有，刪除或另附出處。200 MW 屬實（G42 發言人經 Gulf News；Mubadala Q3 2026）。 | In!G224（納入＝0） | 無 |
| B#30 Broadcom 3.5 GW | 口徑不同 | 更正標記 | 8-K 原文（已讀到）："beginning in 2027… approximately 3.5 gigawatts"、"dependent on Anthropic's continued commercial success"。tag：8-K 屬實，但內容是公司對前瞻承諾的自述，應與 A#30 統一標 Interested-party。$161.2B 出自 TNW（2026-09-29），不是此 url。Anthropic 頁另稱 Google＋Broadcom 5 GW。 | 否 | 無 |
| B#34 SpaceX×Anthropic | 口徑不同 | 更正標記 | Anthropic 頁：Colossus 1 全部、>300 MW、>22 萬顆 GPU。Reuters（Yahoo，第二來源）："now include both… Colossus and Colossus II"，所以「＋Colossus 2」有依據，API 這點錯了。$84.5B 兩個來源都沒有。tag 改為「Anthropic 公告＋Reuters 轉述 S-1」。 | In!G177 註（未分配） | 無 |
| B#37 Riot×Anthropic | 不符 | 更正數值 | expected_online：「未載」→「96 MW 2027-12；全部 2028-06」。191 MW 是否為 IT，原文沒說。 | 否（不在 T2） | 無。若日後納入 T2，應以 96 MW（2027-12）＋95 MW（2028-06）兩筆列入。 |
| F#31 Microsoft→neocloud 60 | 口徑不同 | 更正口徑 | >60、Nscale 約 23（匿名知情人士）、IREN 9.7、Lambda 數十億屬實。但這是 Microsoft **租用算力的採購承諾**，Microsoft 是付款方，不是 Microsoft 的資金來源。tag 改為「Bloomberg 報導（含匿名來源）」。 | **是**（O33＝1，msft、2026、vendor_financing） | **FundRaw O33：1 → 0**（或 item 改為 compute_purchase）。Funding!S7：60 → 空白。S 欄沒有下游引用。 |
| F#194 Anthropic×Fluidstack 50 | 口徑不同 | 更正口徑 | 原文："$50 billion investment… building data centers with Fluidstack in Texas and New York"、"sites coming online throughout 2026"。原文沒有提到 JV、表外或擁有權。item 改為 capex_commitment（結構未揭露）。 | 否（O196＝0） | 無 |
| F#215 Crusoe Series E | 口徑不同 | 更正標記 | 1.375 B、估值 >10、2025-10 屬實。AI Weekly 這頁是 2026-09-04 報導新一輪 >3B／30B（引 Bloomberg），Series E 只是背景敘述。tag 改為「報導（AI Weekly 引 Bloomberg）」。 | 否 | 無 |
| F#221 IREN 9.7 | 口徑不同 | 更正口徑 | 9.7 是 Microsoft 向 IREN 租用 GPU 的合約，對 IREN 是營收，不是融資。NVIDIA 認股權證最多 2.1 由 TNW（2026-05-10）證實；預付 1.9 兩個來源都沒有，要另附出處（F#220 有寫）。與 F#31 的 60 重複。 | 否 | 無 |
| F#227 四家發債 194 | 口徑不同 | 維持 | 194、108、12bp 屬實（Reuters 分析 LSEG）。Goldman 的 250／400 是五家（含 Microsoft）。tag 可改為「第三方數據（LSEG）」，Goldman 部分另標。 | 是（O229＝1，other bond_issue） | 無 |

## 5. 會改到 Excel 輸入的清單

| 輸入 | 位置 | 舊值 | 新值 | 依據列 | 計算影響 |
|---|---|---|---|---|---|
| `cp_msft_abl` MW | In!J189 | 900 | **700** | A#18 | Microsoft 2027 新增 −0.2 GW-IT（Pipeline、Players、Deploy） |
| `xfmr_gva` 低／基準／高 | In!H80:J80 | 1.8／2.0／3.2 | **2.1／2.67／3.15** | P#50 | 只影響 Facility 第 22 列的實物量顯示 |
| FundRaw 採用 F#199 | FundRaw!O201 | 1 | **0**（或 item→compute_purchase） | F#199 | Funding!S30：57→42；不進比例 |
| FundRaw 採用 F#31 | FundRaw!O33 | 1 | **0**（或 item→compute_purchase） | F#31 | Funding!S7：60→空白；不進比例 |
| （選擇性）`fac_gen_px` 低值 | In!H67 | 0.8 | 0.5 | P#27 | 只在 fac_gen_px 單獨取低值時有影響：每 1 倍 genset 少 0.3 $B／IT GW，差額由「其他專業服務」吸收。若 fac_gen_mult 同時取低值 0，則沒有影響 |
| （連帶，非本 46 列）FundRaw F#222 | FundRaw!O224 | 1 | 0 或改 item | 同 F#226 口徑 | Funding other 的 S 欄；不進比例 |

### 只改說明欄、不改數值

| 位置 | 依據列 | 修改內容 |
|---|---|---|
| In!B57、G57 | A#19、B#33 | 「$30B 確定＋額外至多 1 GW」 |
| In!G23 | D#18、D#19 | 刪除 Denton |
| In!G156 | B#4 | 11 為 [Assumed] |
| In!G172 | A#29 | 註明 >1 GW 的下限 |
| In!G189 | A#18 | 改為 700 MW 的說明 |
| In!G210–212 | B#16 | OpenAI 租約 8.0 GW-IT |
| In!G224 | D#21 | 不視為延遲 |
| FundRaw!N201 | F#199 | 撤回「刪除至 2029-05」 |
| Facility!A22 | P#50 | 標籤 |

### F.md 要同步改的地方

- 第 191、330 行：Stargate 股本 52 → 45。

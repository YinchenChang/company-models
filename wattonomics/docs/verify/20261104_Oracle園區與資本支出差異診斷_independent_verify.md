# 獨立核對：Oracle 2028 園區 GW 與資本支出推得 GW 差異診斷（2026-11-04）

- 核對者：未參與製作的代理。對象：branch `claude/wattonomics-oracle-gap`（head 08659bf）。
- 範圍：`wattonomics/docs/verify/20261104_Oracle園區與資本支出差異診斷.{md,xlsx}`、`wattonomics/evidence_raw/E_oracle_20261104.md`、`wattonomics/tools/oracle_gap_diag.py`；模型 `wattonomics/model/20261102_Wattonomics_v0.1.9.xlsx`；個股模型 `oracle/dist/20261030_Oracle個股模型_v0.4.2.xlsx`。
- 程序：先只讀 v0.1.9 活頁簿自行取數，之後才讀報告。xlsx 以 LibreOffice 重算。

## 判定：有條件通過

數字可以重現，橋接的加總恆等式成立，「年底存量看，只承擔部分（A1）是最大單項」的結論在檢驗過的區間內都成立。以下兩項須先修正，才能交給 Andy 判讀：(1) 用個股模型做的交叉驗證並不是獨立檢查；(2) 建議 1、2 的效果數字用了 Shapley 分攤值，不是單獨套用時的效果。

## 一、獨立取數（先於讀報告）

| 項目 | 儲存格 | 本核對值 |
|---|---|---|
| 園區 2028 年新增 | Players!I112（＝I106 − H106） | 4.9117 |
| 園區 2028 年底 | Players!I106 | 8.6048（2027 年底 3.6932） |
| 甲2 2028 上線年新增 | CapexGW!I136 | 1.2367 |
| 甲2 2028 年底 | CapexGW!I139 | 4.8736（2027 年底 3.637） |
| 甲2 中間值 2028 | CapexGW!I126／I127／I133／I128＋I129 | 曆年 58.169 → AI 49.444 → 資金 54.488 ÷ 50.960 → 付款年 1.069 |

2028 年的 T2 園區（Pipeline 列；P＝2028，Q＝1）：

| 列 | 園區 | MW（口徑） | IT GW | 上線（t 基準） |
|---|---|---|---|---|
| 20 | Shackelford 其餘 9 棟 | 1,400（設施；In 272 標 Analogy） | 1.1667 | 2028-11 |
| 21 | Doña Ana（Jupiter） | 2,200（設施；Analogy） | 1.8333 | 2028-12 |
| 22 | Port Washington | 1,300（設施；Analogy） | 1.0833 | 2028-12 |
| 23 | Saline 其餘 | 922（設施；Analogy） | 0.7683 | 2028-12 |
| 96 | DataBank DFW11 | 60（IT；Derived，二手） | 0.0600 | 2028-01 |
| | 合計 | | 4.9117 | 4.852 GW（98.8%）落在 11–12 月 |

讀報告前的初步觀察：Oracle 2028 年曆年資本支出只有 58.2（FY29 Y9 降到 38.0）。甲2 用全額成本 50.96，其中廠房占 23.9%，而且園區都是整站放在年底完工月。這和報告列出的主因方向相同。

## 二、逐項核對

### 檢查 1：數字重現（通過）
4.912、8.605、1.237、4.873（4.8736）都已重現。報告附件 xlsx 的 Check 頁 14 項全部通過；重算前後 1,771 個數值格完全相同。

### 檢查 2：橋接（通過，附三項說明）
- 重算 Bridge 頁各項，2028 年新增／年底：C1 −0.187／−0.267；B1 −1.666／0；D −0.848／−0.883；B2 −0.178／+0.035；C3 −0.110／−0.266；C4 −0.155／−0.477；A1 −0.532／−1.874。合計 3.675／3.731，等於總差距。
- Shapley 權重 |s|!(n−|s|−1)!/n! 正確；四因子合計＝V_ACST − V_base（0.9747），符合效率性。D 由定義（V_ACST − 分期後園區）就是殘差，所以各項加總等於總差距是恆等式。
- 重複計算：A1 與 C3、C4 之間的交互作用由 Shapley 分攤，沒有重複計算。B1（園區時點）和 B2（資本支出時差）分在橋接的兩側，也沒有重複計算。不過兩者方向略有矛盾：分期代表園區比清單早上線，B2 卻採用較長的時差 7.7 月。B2 只有 0.18，影響小。
- **未說明的順序相依：** 園區端 C1 與 B1 是依序計算，沒有用 Shapley。若先分期、再縮放 Shackelford，C1 的年新增由 −0.19 變成 −0.11（Site S9 3.173 對 S2 3.059），B1 則相應變大。報告只寫「轉換段四因子與順序無關」，字面上沒錯，但應註明園區端兩項的值取決於順序。

### 檢查 3：外部來源

| 主張 | 結果 | 網址 |
|---|---|---|
| (i) 財務長 2025-12-10「not for land, buildings, or power … covered via leases」 | 屬實（二手轉述，Kehring，Q2 FY26 法說）。E#6 Fortune 是同一段發言，並不是兩個獨立來源 | https://www.thestack.technology/oracle-defends-debt-led-ai-land-grab/ |
| (ii-a) Investor Day 新聞稿：FY27–28 資本支出高峰、FY29 下降；供應商融資、客戶自帶硬體、預付 | 屬實（公司原文） | https://investor.oracle.com/investor-news/news-details/2026/Oracle-Investor-Day-2026-Accelerating-Value-Creation/default.aspx |
| (ii-b)「about seven gigawatts」「exceed 10 GW by end FY28」 | 屬實（二手，作者轉述，不是直接引語；新聞稿本身沒有 GW 數字；口徑未說明） | https://aiweekly.co/alerts/oracle-projects-30b-fcf-by-fy2029-caps-fundraising-at-20b ；https://book.st-hakky.com/en/news/oracle-targets-30b-free-cash-flow-by-2029 |
| (iii) 未起租租賃承諾 288B（2026-08-31）、FY27Q2–FY29 起租、15–19 年 | 屬實（二手；10-Q 原文本次未開） | https://business-news-today.com/oracle-has-664bn-of-contracted-revenue-visibility-what-will-it-cost-to-turn-that-demand-into-cloud-capacity/ |
| (iv) Shackelford：Vantage 1.4 GW、10 棟、首棟 2026 下半年 | 1.4 GW 屬實；T2 的 2.0 GW（Epoch）本次無法核對 | https://www.datacenterdynamics.com/en/news/vantage-breaks-ground-on-texas-gigawatt-data-center-campus-for-openai/ |
| (v-a) Shackelford、Port Washington 由 Vantage 開發，380 億債務 | 屬實（文中寫「used by Oracle」，沒有寫「承租」） | https://datacenterdynamics.com/en/news/jpmorgan-chase-and-mitsubishi-ufj-lead-38bn-debt-package-for-oracle-linked-data-centers-report/ |
| (v-b) Jupiter：STACK 與 BorderPlex 開發、Oracle 為承租方；不可抗力通知的對象是 Blue Owl 旗下開發商 | 屬實 | https://www.datacenterdynamics.com/en/news/oracle-revealed-as-tenant-of-project-jupiter-data-center-campus-in-new-mexico/ ；https://thenextweb.com/news/oracle-force-majeure-project-jupiter-new-mexico |
| (v-c) Saline：Related Digital、Blackstone 股權、PIMCO 債務、160 億、為 Oracle 興建 | 屬實（公司原文） | https://www.blackstone.com/news/press/related-digital-announces-financing-for-16-billion-oracle-data-center-project-in-saline-township-michigan/ |

### 檢查 4：每 IT GW 隱含成本 30.7（算術正確，口徑有條件）
- (92.5＋86.398) ÷ (7 ÷ 1.2) ＝ 30.67；若 7 GW 是 IT 口徑則為 25.56。重現無誤。
- 口徑：Wattonomics 用 IT GW，觀測值 obs_orcl 也是把 Oracle 的 MW 除以 1.2，所以把 7 GW 視為設施口徑，與 Wattonomics 一致。但 Oracle 沒有說明 7 GW 是 IT 還是設施口徑。
- 三個偏誤：(a) 分子用的是全部資本支出，不是 AI 部分；乘上 0.94 後約為 28.9。(b) FY27–28 資本支出與 FY27–28 上線容量之間有時差：一部分容量用 FY26 的錢，一部分錢會在 FY29 才上線。(c) 7 GW 可能包含客戶自帶硬體的容量，所以「Oracle 只付 IT」和「客戶自帶硬體」在這項檢查中無法分開辨識。
- 與園區清單的時點衝突（未調和）：依未分期的 T2 清單，FY27–28（2026-06 至 2028-05）只新增約 2.2 IT GW，遠低於 7 ÷ 1.2＝5.8。只有接受 B1 分期，兩者才大致相容。

### 檢查 5：標記與一致性（有條件通過）
- 錯誤格 0。md 與 xlsx 的數字逐項相符，例外如下。
- 缺低／高區間：w1、w2（Assumed）；sh_m 的 FY25–26（Assumed）；itm 的 FY25–26（Assumed）；ai_inserv 隱含耐用年限 6 年（標 Derived，但含 Assumed 參數）。In 的園區表沒有標記欄。
- md 與 xlsx 不一致或敘述有誤：
  1. 結論 2 寫「其次是 AI 占比與每 GW 成本組合（合計 1.01 GW）」，但 1.01 是 C1＋C3＋C4，含 Shackelford 0.27。AI 占比加成本組合只有 0.74。
  2. §4 寫「管理層原話（E#5、E#6，兩來源）」，實際上是同一段發言的兩篇報導。
  3. §1.3 L 的來源寫「2028 五站」，§1.2 寫「九站」。

### 檢查 6：邏輯
- 「年底存量看，只承擔部分約占 50%」成立：A1 的 Shapley 值為 1.87／3.73＝50%。單獨作用到最後作用的區間是 44–56%。L＝0.8 時 A1 約 1.31（35%），仍是最大單項；第二大是 D 的 24%。前提是財務長的說法（單一管理層來源）成立。
- **個股模型的交叉驗證不是獨立檢查（須修正）。** 個股模型 Thesis!D41–G41 的已連網 MW＝MIN(cap_B, mw_q1＋pace_B×期間)，其中 cap_B 是「合約容量上限（已簽站點）」（build_oracle.py:258）。也就是說，MW 由交付速度和已簽站點決定，不是由資本支出推得。所以 md §3 的「同一筆 Y9 資本支出對應 8.78……差距只出在 Wattonomics 的換算」不成立：8.78 和園區 8.61 一致，是因為兩者都來自已簽站點，接近循環論證。此外還有兩處未調和的差距：
  - 用 Oracle 一致假設（A＋S＋C＋T）換算，2028 年底只有 7.455，比個股模型低 1.32；
  - 2027 年底個股模型為 6.32，Oracle 一致換算為 5.24，md 只列了 5.28 對 5.24。
- **建議的效果數字（須修正）。** 建議 1 的「+0.53／+1.87」和建議 2 的「+0.16／+0.48」是 Shapley 分攤值。若單獨執行，效果是：
  - 建議 1（只套 A）：+0.41／+1.66（Calc V_A：1.650／6.529）；
  - 建議 2（只套 S）：+0.12／+0.37（V_S：1.354／5.247）。
  兩項一起執行的 7.075（V_AS）正確，建議 4 的 7.08、1.53、1.26 也正確。
- 未調和：Oracle「FY28 年底（2028-05）逾 10 GW」換算約 8.33 IT GW。本核對粗估分期後的園區在 2028-05 約 6.5，仍差約 1.8，原因可能是 10 GW 含非 AI 舊區域，或清單漏了站點。md 把這項當作「支持分期」的證據，應同時註明這個缺口。

## 三、必要修正
1. md §3 第 3 列與 xlsx Bridge 交叉驗證列：改述個股模型的 MW 來自交付速度與已簽站點上限，不是由資本支出推得；刪除「差距只出在 Wattonomics 的換算」；補列 2027 年底 6.32 與 2028 年底 7.455 對 8.78 的差距。
2. md 建議 1、2：效果改列單獨執行值（+0.41／+1.66、+0.12／+0.37），並註明 Shapley 值的意義。
3. md 結論 2：「合計 1.01」改為「AI 占比＋成本組合 0.74，另 Shackelford 口徑 0.27」。
4. 把 E#5、E#6 改為單一來源（同一法說發言）。
5. 為 w1、w2、sh_m（FY25–26）、itm（FY25–26）、life_ai 補上區間。
6. Bridge 註明園區端 C1 與 B1 依序計算：C1 為 −0.11 至 −0.19，視順序而定。
7. md §3 補註 30.7 的三個偏誤，以及與未分期清單的時點衝突。

## 複核（91ad296）

判定：**通過**。七項必要修正都已完成。xlsx 經 LibreOffice 重算後錯誤格 0，Check 頁 14 項全數通過。

| 項目 | 結果 |
|---|---|
| 1 個股模型交叉驗證 | 已改述為「不是獨立檢查」（cap_B 為已簽站點上限）。重算各項未調和差距：7.455 對 8.778（−1.323）；5.244 對 6.320（+1.077）；8.333 對 6.578（1.755）。均相符 |
| 2 單獨套用效果 | Bridge 第四節：只套 A 為 +0.4136／+1.6558；只套 S 為 +0.1171／+0.3732；A＋S 同時為 +0.5633／+2.2015，年底值 7.0751。與本核對的 Calc 頁 V_A、V_S、V_AS 一致 |
| 3 結論 2 | 已改為 0.74，Shackelford 0.27 另列 |
| 4 E#5／E#6 | md、xlsx 與 E 段證據檔都已註明為單一來源 |
| 5 區間 | w1、w2 為 (1,0)／(0,1)；sh_m FY25 為 0.75–0.95、FY26 為 0.85–0.95；itm FY25–26 為 24.02–32.24；life_ai 為 5–6，推回的毛額為 48.4–58.0，且已改為公式。均已補上 |
| 6 C1／B1 順序 | 已加註；先分期再縮放時 C1＝−0.1141，相符 |
| 7 30.7 的偏誤 | 分子改用 AI 資本支出得 28.907，相符。FY27–28 新增：未分期 1.918、分期 4.743，相符。小瑕疵：未分期值用未縮放的 Shackelford 首棟（0.50），分期值用縮放後的（0.42），兩者口徑差約 0.08，不影響結論 |
| 租金已扣在營業現金流 | 屬實。路徑為 Thesis!57 租金 → Thesis!63「EBITDA（租金後）」＝收入 − 營運費用 − 租金 → 分部利潤貢獻 → PL 分部利潤 → 淨利 → CF!55 營業現金流（＝淨利＋D&A＋SBC＋預付流入 − 預付認列）。Wattonomics In 第 197 列 gocf_orcl 讀的正是 CF 第 55 列 |
| 融資租賃備註 | 合理。依美國會計準則，融資租賃本金償還列在籌資活動；個股模型把租金全部以直線方式放在營業面，融資租賃負債 9.185 另計入淨借款 |

其餘小瑕疵：In 的 mw_fy26 來源說明寫「FY26 交付 1.2 GW ÷ 1.2 等」，語意不清，且單位標 GW、名稱卻寫 MW-IT。建議下次改版時修正，不影響本次判定。

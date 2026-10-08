# WhiteFiber 收支模型 — 交接檔（v0.1，2026-10-08；評價日 2026-06-30〔FY26Q2 10-Q〕；現價 $16.83〔2026-10-07〕）

## 1. 這是什麼（命題）
WhiteFiber（WYFI）的資金與評價模型，CRWV「Backlog 不是現金」、Nebius「預付款」、Oracle「現金牛＋客戶出資」之後的第四個對照組。WhiteFiber 同時做「GPU 雲（短約、高單價）」與「託管 colocation（10 年長約、低單價、重資本）」。命題：**資產負債表最小、控股股東（Bit Digital 59.9%）撐腰、對手方本身也是 neocloud（Nscale）時，長約託管能否成為可融資的現金流底座、支應 NC-1 擴建與 GPU 資本支出並償付可轉債，而不必反覆稀釋。**

**目前答案（信心中低）**：**託管是價值與融資的底座，雲端擴張本身不創造價值。** 三情境目標價 保守 $10.04（賣出）／基準 $16.84（中立）／積極 $34.66（中立），基準與現價持平。價值幾乎全來自託管分部的 EV/EBITDA 腿（DCF 腿在三情境皆約 0）；Nscale 第一期營收歸零時基準目標價只剩 $5.1。基準情境 NC-1 專案貸款（307m）若完成，五期不需新股；未完成則需 ATM 股權 76m（新股 5.0m，約 +13%）與高息債 200m（目標價 $16.31）。雲端每 MW 年 EBITDA（GPU 租金前）基準 $8.2m，低於全額 GPU 年化資本回收 $10.9m；只有積極情境的每 MW 單價（$24.2m）略高於打平所需（約 $22.9m）。

模型同時輸出 HTML（互動、可調輸入）與 Excel（全部活公式），兩者數字必須一致（cmp31 三情境＋錨定 1 各 427 項）。**金額單位 US$m、股數 m 股**（`company.json` → `meta.unit`＝m）。

## 2. 檔案與規則
| 檔案 | 說明 |
|---|---|
| `dist/20261008_WhiteFiber收支模型_v0_1.html` | 互動版；單一檔案、離線雙擊即可開啟（無任何網路請求） |
| `dist/20261008_WhiteFiber收支模型_v0_1.xlsx` | Excel 版；全部活公式（情境區間以模擬運算表計算） |
| `company.json` | 公司原始輸入的唯一來源（HTML 與 Excel 共用）；欄位說明見 `README.md`「company.json 欄位說明」 |
| `data/whitefiber_facts_20261008.json` | 事實總帳 303 筆（每筆：數值、單位、期間、來源與網址、文件日、擷取日、標記；金額 USD M） |
| `data/permw_tokenomics_20261008.json` | 每 MW 年收入（Tokenomics v5.26 Interface I7:N15，與 v5.24 相同：11.62／17.40／24.20）＋ WhiteFiber 對照（Q2 實現、合約隱含單價、毛利率） |
| `data/consensus_wyfi_20261008.json` | 市場共識（只讀；`meta.consensusFile` 指向；資料檔 US$M，載入時依 `meta.unit` 換算） |
| `docs/reports/20261008_whitefiber_v0.1a_資料.md`、`…_v0.1b_模型.md`、`…_v0.1.md`、`…_進度.md` | 資料報告、模型改寫報告、**v0.1 最終報告**（結論、敏感度、預設、缺口、限制、verify 完整輸出）、分段建置進度 |
| `docs/workorders/`、`docs/plan/` | 工作單（共同規則、v0.1a／b／c）與命題與驅動對照表 r1 |
| `scripts/verify.sh` | 一鍵建置與核對（19 項；`--vs-dist` 22 項） |

- **命名**：`更新日_WhiteFiber收支模型_v版本號`（公司名讀 `meta.company`；版本讀 `vlog.py` 最後一列；更新日＝`dist/` 成品檔名日期，可用環境變數 `DATE` 覆寫）。
- **升版**：只有 `dist/` 成品改變時才升版；`vlog.py`（Excel）與 `tail.js` 的 VLOG（HTML）兩處都要新增一列；`dist/` 與 `docs/handoff/` 都只留最新版。
- **流程**：chat 端寫工作單 → 建置代理執行 → chat 端審查並合併；代理不得合併、不得 force push、不得改 `main`，只改 `whitefiber/`（另可改根目錄 README 的 WhiteFiber 列）。

## 3. 模型結構與因果
**因果主軸**：雲端已連網 MW → 可計費 MW（延誤平移）→ 雲端營收 × EBITDAR 率 − 機房租金 − GPU 租金 → 雲端 EBITDA；託管站點 MW × 合約租金 × 年調 → 託管營收 × EBITDA 率；營運現金＋客戶預付 − CapEx（GPU 自購＋託管建置）− 租金 − 利息 − 稅 → 融資缺口 → NC-1 專案貸款（排程）→ 瀑布（現金 → 額度與資產層新債〔總債務 ≤ 5× EBITDA〕→ 可轉債 → ATM 股權 → 高息債）→ 股數與淨負債 → DCF＋分部 EV/EBITDA → 目標價。

| 區塊 | 公式與輸入（`company.json` 位置） |
|---|---|
| 雲端已連網 MW-IT | 由合約 GPU 數 × 每顆 IT kW 換算：首期期末 14.574（冰島 5.5＋Baseten 2.858＋巴黎 1.84＋雪梨 2.5＋其他 1.876）；FY27 加 Prime VR200 1.974；保守不含雪梨；基準／積極 FY28 起 +3／+6 MW-IT／年 [Assumed]（`scenarios.mwPath`） |
| 可計費 MW | 期初＝Q2 經常性雲端營收 11.506 × 4 ÷ 各情境每 MW（保守 3.96／基準 2.65／積極 1.90）；期末＝期初＋(已連網 − 期初)× 收斂 70／85／90／90／90%（`scenarios.billableRatio`）；計費 MW 往後平移延誤月數 |
| 雲端營收 | 平均計費 MW × 每 MW 年收入（11.62／17.40／24.20 US$m／MW-IT·年，`scenarios.revMW`）× 期間長度；合約隱含單價（`texts.revMwContracts`）只作對照 |
| 雲端 EBITDA | EBITDAR 率（起點 −11.49%＋7.24% → 穩態 47%＋1.70%，`defaults.ebStart`／`ebSteady`／`ebitdarAdj`）× 營收 − 機房租金 − GPU 租金；基準 FY30 EBITDA 率 25.9%（GPU 租金前 47%） |
| 託管分部 | `defaults.colo.sites`：已簽約 NC-1 A／B 各 20 MW（1.886／MW·年、年調 3%）、MTL-3 5 MW（2.35）、MTL-1 3 MW（2.38）；基準加 MTL-2 3 MW、NC-1 下一批 33.3 MW（2028-01，建置 383.3）；積極再加 NC-2／3 44.4 MW、NC-1 擴建 A／B、Krambu 74.1 MW（多為 [Assumed]）。未簽約站點租金 1.85、建置 11.5／MW-IT；EBITDA 率 52.2% → 65%；建物折舊 20 年（期初 PP&E 556.3）；經模板「傳統事業」列進入損益與 EV/EBITDA（畫面用語 `texts.labelMap`） |
| CapEx 與 GPU | GPU 成長型投資＝新增 MW × 每 MW 37.45／37.59（Tokenomics IF_CapexIT）；新增 50% 租賃（年租金係數 26.04%，`defaults.gpuLease`）、50% 自購走 CapEx 與 5 年折舊；託管建置另計 |
| 客戶預付 | 預付＝(雲端自購成長型 CapEx＋託管建置) × 21.1%；期初合約負債 143.1；8 年認列為非現金營收（`defaults.prepay`） |
| 債務與證券 | 期後結構（`valuation.postEvents`）：2032 可轉債 310（5.0%，轉換價 33.84）、2031 剩 31.85、DDTL 80（9.5%，MOIC 加付 1.9）、冰島 18、RBC 聯貸 26.2；零履約價買權 5.91m 股自評價股數排除；評價股數 39.34m |
| NC-1 專案貸款（v0.1c） | `scenarios.projectFinance`：基準／積極 307.1（NC-1 累計建置 472.5 × 65%）、保守 0（未完成）；FY27 期初一次動用、利率 8.5%、次期起 10 年直線攤還；利息併入存量利息、攤還併入排程還本、餘額計入總債務；**不占公司層級 5× 上限**（專案層級核貸） |
| 融資瀑布 | 現金（最低 30）→ DDTL 未動用 20 → 資產層新債（總債務 ≤ 5.0× 年化 EBITDA，9.5%）→ 可轉債（每年 ≤ 現市值 10%＝75.9，票息 5%）→ ATM 股權（每年 ≤ 10%，現價 × 0.9）→ 高息債 12% |
| 評價 | WACC 13.87%（CAPM：rf 5.27%＋β 2.5 × ERP 5%＝17.77%；kd 9.5% 稅後 7.5%；權重 62/38）、g 3%；DCF 45%（0 截斷）＋分部 EV/EBITDA 55%（託管 × 20.75〔DLR／EQIX 中位〕＋雲端 × 6，錨定 FY29）；目標價時點＝評價日＋12 個月 |
| 建設延誤 | 保守 6／基準 3／積極 0 個月：雲端計費 MW 平移、GPU 資本支出照原時程（閒置資本）、雪梨未起租 delayLink 0.5、託管未簽約站點起租同步後移；罰則 0 |

## 4. 目前結果（v0.1；US$m，除另註）
| 情境 | 目標價 | 評等 | EV/EBITDA 腿 | 五期 專案貸款／新債／可轉債／股權／高息債 | 新股（m） | FY30 營收／EBITDA／總債務 |
|---|---|---|---|---|---|---|
| 保守（只有已簽約） | $10.04 | 賣出 | $18.26 | 0／48.2／8.2／0／0 | 0 | 255.0／104.1／399.1 |
| 基準（加 NC-1 下一批） | $16.84 | 中立 | $30.62 | 307.1／98.8／96.9／0／0 | 0 | 545.0／209.8／753.4 |
| 積極（NC-2／3 與 Krambu） | $34.66 | 中立（終值占比規則） | $63.01 | 307.1／1,719.8／151.8／80.6／1,156.8 | 5.32 | 1,388.1／666.1／3,586.2 |

基準 FY26／FY27／FY28 營收 129.9／278.2／417.4 對 S&P 共識 137／296／530（−5.2%／−6.0%／−21.2%）；共識平均目標價 $38.0（MarketBeat，11 家，32–50），模型情境區間 $10.0–$34.7 全在共識之下（主因：雲端每 MW 單價取 Tokenomics 正向推導而非公司目標、託管倍數只有 2 家同業、beta 2.5、Nscale 集中）。關鍵敏感度見最終報告 §5。

## 5. 資料來源與標記
- 10-K FY2025（經查核）＝[Verified]；10-Q（核閱）、新聞稿、法說、簡報＝[Interested-party]；法說只有二手摘要（MarketBeat）。由其他數字換算＝[Derived]；類比＝[Analogy]（給區間）；無實證＝[Assumed]（給區間）。
- SEC EDGAR CIK 0002042022。事實總帳 303 筆（Verified 50／Interested-party 203／Derived 44／Analogy 4／Assumed 2）；每個 `company.json` 數字可追溯到總帳 id 或 `data/` 檔。
- 每 MW 單價只用 Tokenomics 正向推導；公司「年化 > $200M」目標與合約單價不作輸入（只作對照與敏感度）；例外：Q2 經常性雲端營收校準期初、NC-1 等已簽約託管合約排程。

## 6. 已套用的預設（v0.1a–c 彙整；詳表見最終報告 §7）
| 類 | 預設（替代；影響） |
|---|---|
| 評價 | β 2.5（1.5／3.5：$21.32／$15.75）；kd 9.5%；託管倍數 20.75×（15×：$10.88）；雲端倍數 6×；權重 45/55；錨定 FY29 |
| 雲端 | 每 MW 11.62／17.40／24.20；期初以經常性雲端營收校準；EBITDA 起點經常性 −11.49%（含一次性 20.07%：+$2.8）；穩態 47%；GPU 50% 租賃（0%／100%：$18.66／$16.26） |
| 託管 | 未簽約租金 1.85（1.45／2.35：$14.73／$19.95）；建置 11.5（9.0／13.5：$18.29／$15.79）；EBITDA 率 52.2%→65%；建物 20 年；情境站點如 §3 |
| 融資 | 期後資本結構；零履約價買權排除（計入：基準 $14.64）；NC-1 專案貸款 307.1、8.5%、10 年、FY27、不占上限（未完成：$16.31）；債務上限 5×（4×／6×：$16.86／$16.79）；最低現金 30；可轉債與 ATM 各 10% 市值／年 |
| 其他 | 延誤 6／3／0 個月（0／12 個月：$17.80／$14.75）；預付覆蓋 21.1%、8 年、財務組成 0；稅率 21%、無 NOL；金額單位 US$m（v0.1c） |

## 7. 資料缺口（找不到 ≠ 不存在）
- **不存在（公司未揭露）**：預付與遞延營收的託管／雲端拆分；CapEx 依建物／電力／GPU 拆分與 NC-1 總建置成本；GPU 自購 vs 租賃比例；巴黎、安大略、加拿大、亞特蘭大 MW 與機房租金；Nscale NRC；NC-1／巴黎專案融資條件；正式財務指引；Krambu 條件。
- **找不到**：EBITDA、CapEx、淨負債共識；託管同業 NTM EV/EBITDA 只 DLR、EQIX（APLD、CORZ、IREN、CIFR 無法算出）；可靠 beta（價格歷史 API 被環境阻擋）；法說完整逐字稿；CRWV、NBIS 同業數字未能更新（沿用 2026-10-07）。

## 8. 已知限制
1. 續約價格衰退未做（雲端合約到期後以 Tokenomics 每 MW 單價續作，可能高估後段營收）。
2. Nscale 對手方信用未建模（只有違約敏感度：NC-1 第一期營收歸零 基準 $5.07）。
3. 上市時間短（2025-08 IPO）：beta 與共識樣本小（目標價 11 家、FY28 營收 3 家）；託管同業倍數只 2 家。
4. DCF 腿在三情境約為 0（股權為負、0 截斷），目標價實質由 55% × EV/EBITDA 腿決定；積極情境評等因終值占比規則為中立。
5. 託管站點逐列明細只在 Excel（HTML 只有彙總列與檢查頁）；龍捲風與敏感度只在 HTML。
6. 畫面用語以 `texts.labelMap` 由模板名稱替換（「傳統事業」→「託管」等）；程式內部仍用模板名稱。
7. GPU 租賃負債未計入 leaseAdj 槓桿（本模型上限基準為 ebitda，不受影響）；閒置資本只計 GPU，未含託管已支出未起租的建置。
8. NC-1 專案貸款只建模第一期（472.5 × 65%）；NC-1 下一批與擴建的專案融資不建模（由瀑布支應，偏保守）；專案貸款不占公司層級上限為 [Assumed]（若佔上限，積極情境完成 $32.47 < 未完成 $34.42）。
9. Krambu、NC-1 擴建、NC-2／3 的時程與經濟條件多為 [Assumed]；積極情境需高息債約 1.16bn，結果對這些假設高度敏感。
10. 單位改 US$m 時發現模板兩處以 US$bn 寫死的下限（終值占比分母、損益兩平貢獻率分母），在 WhiteFiber 規模失真；v0.1 起按實際值計算（只影響「100% EV/EBITDA 目標價」評等代碼與顯示）。
11. historicalPL 無 FY23（上市前）；季度層無季度共識與指引（「不適用」）。

## 9. 待辦與季度更新
**待辦**（見 `CLAUDE.md`「待辦（WhiteFiber）」）：(1) FY26Q3 季度更新（升 v0.2）；(2) Nscale 對手方信用；(3) 託管同業倍數與 beta；(4) 共識（EBITDA、CapEx、淨負債）；(5) v0.1b 審查列為已知限制的 4 項；(6) CRWV 未完成待辦（範圍外）。

**FY26Q3 財報後的季度更新步驟**（季末 2026-09-30，財報約 2026-11 中，10-Q 約同時）：
1. **財報當天（新聞稿）**：填 `quarterly.actuals.FY26Q3`（營收、調整後 EBITDA、調整後營業利益、CapEx、MW；來源、日期、標記）；`calendar.latestQuarterReported` 改 `FY26Q3`；焦點季改 `FY26Q4`；更新 `callFacts`（指引若有、RPO、MW、預付），檢查 NC-1 專案融資、巴黎專案融資、NC-2／3 收購、Krambu 是否有進展。
2. **10-Q 申報後滾動評價日**：`calendar.latestQuarterFiled` 改 `FY26Q3`，評價日 2026-09-30，首期縮為 Q4（0.25 年）。同時更新（金額 US$m）：
   - `ytdActual`（三季累計：營收、CapEx、CFO、預付、利息、借款、股權、還款、D&A、SBC、調整後 EBITDA 與各季 `adjEbitdaMeta`、逐項 `notes`）、`historicalPL` 最後一格（9M26）、`latestQuarter`（現金、債務本金、流通股、RPO、遞延營收、租賃負債、未起租承諾）。
   - **期初可計費雲端 MW 重新校準**：`scenarios.billableRatio.openAnnualRevenue`＝Q3 經常性雲端營收 × 4（扣一次性）；`mwPath.connectedStart`＝評價日雲端已連網 MW-IT；檢討收斂比例與雲端 EBITDA 起點（`defaults.ebStart`，以 Q3 經常性數字校準）。
   - 期後事件 `valuation.postEvents`：已在 Q3 資產負債表內的項目（2032 可轉債、2031 交換、DDTL、RBC 聯貸）移除，改由 `latestQuarter`／`debt` 反映；Q3 後的新事件另列。
   - 託管：NC-1 實際起租與計費、MTL-2 完工；`defaults.colo.ppeOpen` 改 Q3 託管 PP&E；`rpo.colo`／`colo.rpoColo` 改 Q3 託管 RPO 年度分布。
   - NC-1 專案貸款：若已簽約，`scenarios.projectFinance` 改為實際金額、利率、攤還與動用期別（必要時改為首期動用），並移出「假設」。
   - `asOf` 全部項目（清單在 `calendar_q.py` → `ROLL_FIELDS`）改完後把季度改為 `FY26Q3`，缺一項即建置失敗。
   - **EBITDAR 校準**：`node scripts/calib_ebitdar.js --write`（verify.sh 第 0c 項會擋未更新者）。
3. 共識：新共識檔 `data/consensus_wyfi_YYYYMMDD.json`，改 `meta.consensusFile`；股價、市值、同業（`peers`、`valuation.price`、`callFacts.priceLast`、`meta.priceDate`）一併更新。
4. 目標價變動拆成 (a) 時間推移、(b) 實際數更新、(c) 假設變更、(d) 方法變更（`scripts/attrib.py`），列入 VLOG 第二列。
5. `vlog.py` 與 `tail.js` 新增 v0.2；`DATE=YYYY-MM-DD scripts/verify.sh` 建置，`dist/` 換新成品；對舊成品跑 `verify.sh --vs-dist`（以 `EXPECT=` 清單列預期變動的格）；交接檔換新版。

## 10. 如何重建
1. 環境：LibreOffice Calc（缺時 `apt-get install -y --no-install-recommends libreoffice-calc-nogui`）、Python `openpyxl`、`playwright`（1.56.0）＋ Chromium（`python3 -m playwright install chromium`）、Node.js 18+。
2. 在 `whitefiber/` 執行 `DATE=2026-10-08 scripts/verify.sh`：季度層檢查 → README 欄位說明 → EBITDAR 校準 → 建 HTML → 建 Excel → LibreOffice 重算（公式錯誤 0）→ 反向 DCF（`rv_snap.json`）→ fix_outline → fix_datatable → verify_ooxml → 離線開啟 → 三情境與錨定 1 cmp31（各 427 項）→ 季度層、期間滾動、拆解工具測試；19 項全過，約 5 分鐘。產物在 `out/`（不入版控）：`out/20261008_WhiteFiber收支模型_v0_1.html`／`.xlsx`。
3. 成品：把上一步產物複製到 `dist/`，再跑 `DATE=2026-10-08 scripts/verify.sh --vs-dist`（22 項，畫面文字與 Excel 值、公式 0 差異）。
4. 改輸入：只改 `company.json`（金額 US$m；新欄位在 `scripts/fields_doc.py` 補說明並 `--write`）；改公式：Excel 在 `build_xlsx.py`、HTML 在 `segA.js`～`segE.js`，兩邊須經 cmp31 一致；以 US$bn 寫的程式常數一律乘單位係數（JS `UFQ`、Python `UF`／`_u()`）；`CLAUDE.md`「勿改」各條仍適用。

## 附錄　模板來源
本模型由 `oracle/` @ `e4540c3`（Oracle v0.2 分支＝CRWV v4.5＋Nebius v0.1b＋Oracle v0.1–v0.2）複製建立；CRWV、Nebius、Oracle 的版本紀錄與交接檔留在各自資料夾，本資料夾 VLOG 自 v0.1 起算。

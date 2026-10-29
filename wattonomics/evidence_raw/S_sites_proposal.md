# Wattonomics T2 園區列合併提案｜摘要（2026-10-28）

輸入：`M.csv`（409 列候選證據、131 個園區群）＋ `existing_T2.csv`（43 列）。輸出：`proposal.csv`（206 列增量，每列一個增量事件；與其他候選完全重複的證據列已併入保留列的 evidence 欄）。409 列證據都至少被一列引用（S#28 拆成 Mesa B1、B2 兩列使用）。

## 1. 計數

| 項目 | 數 |
|---|---|
| 證據園區群（M.csv `site`） | 131 |
| 提案園區標籤（含拆分，如 QTS Fayetteville 單列） | 135 |
| 提案增量列 | 206 |
| 納入=1 | 11 列，合計 5,898 MW |
| 納入=0 | 195 列 |
| 已確保電力=1 | 0 列（見 §5 第 6 點） |

### 納入=1 依年份（MW）

| 上線年 | MW | 列 |
|---|---|---|
| 2026 | 30 | cp_sov_ajman_p1 |
| 2027 | 900 | cp_neo_barber 207（03）、cp_oai_tomakomai 50（03）、cp_neo_bp_p1 352（09）、cp_amzn_ulsan1 41（11）、cp_neo_pf2 200（12）、cp_nbis_oxagon 50（12） |
| 2028 | 1,906 | cp_neo_monarch_p1 1,350（03）、cp_amzn_rdg_b614 556（09） |
| 2029 | 62 | cp_amzn_ulsan2（02） |
| 2030–2031 | 0 | — |
| 2032 | 3,000 | cp_meta_hyp_5g（12，只有年份） |
| **合計** | **5,898** | |

### 納入=1 依營運者（MW）

| op | MW | 說明 |
|---|---|---|
| meta | 3,000 | Hyperion 2→5GW 餘量（占 51%，單列影響大，見 §5 第 1 點） |
| neo | 2,109 | Monarch 1,350、Beacon Point 352、Barber Lake 207、Polaris Forge 2 200 |
| amzn | 659 | Ridgeland 556、蔚山 41＋62 |
| nbis | 50 | DataVolt Oxagon |
| oai | 50 | SoftBank 苫小牧（歸類待定，見 §5 第 3 點） |
| sov | 30 | Khazna Ajman |
| msft／googl／orcl／crwv／xai／ant | 0 | |

口徑：IT 1,345（556＋352＋207＋200＋30）；未說明 4,553（3,000＋1,350＋50＋50＋41＋62）；設施 0。

### 納入=0 依原因類別（195 列）

| 類別 | 列數 | 內容 |
|---|---|---|
| API 判定不足 | 63 | 判定為無法取得／原文找不到／不符／口徑不同，或小誤涉 MW、口徑、上線日 |
| oth（企業／colo 無租戶） | 30 | 依規則只列不入 |
| 重疊既有列 | 29 | 同園區同期已在 T2 |
| 觀測點前已上線或超出窗口 | 18 | 含 msft/googl/amzn/orcl/crwv/xai 早於 2026-07 者（可能已在觀測值內）及 2036 等窗外 |
| 電力／發電／需量／PPA，非 DC 容量 | 17 | |
| 無上線日 | 10 | |
| 無 MW | 8 | |
| 選擇權／MOU／暫停／已退出 | 8 | |
| cn | 6 | 走 Rystad 由上而下 |
| 遠期願景 | 3 | |
| 重疊既有 納入=0 列，建議升級既有列 | 3 | 見 §3 |

## 2. 與既有 T2 列矛盾

| 既有鍵 | 既有值 | 候選值 | 證據 | API 判定 |
|---|---|---|---|---|
| cp_msft_atl | 636 IT 全放 2025-11 | B1–2 321 IT 於 2025-10；636 至 2026-06 才達成（增量 315） | S#5、S#6 | 小誤、小誤 |
| cp_msft_wi1 | 369 IT @2026-06 | 一期 2026-04-16 已上線 | S#12 | 不符（369 屬實、日期不符） |
| cp_msft_abl | 700 未說明 @2027-06 整筆 | 2 棟 × 336 IT＝672 IT；首棟 2027 年中，第二棟時點未揭露；900 為園區／電廠口徑 | S#24、S#25、S#26 | 小誤（basis）、小誤、小誤（expected_online） |
| cp_meta_pro1 | 471 IT 至 2026-08 | 562 IT（2026-09-02，Epoch） | S#46 | 小誤（不涉 562） |
| cp_meta_pro2／pro3 | 676 @2026-11、868 @2027-07 | 854 IT @2027Q1 | S#47 | 口徑不同 |
| cp_meta_pro1–4 | 合計 1,022 至 2028-09 才滿 | Meta 稱 1GW 叢集 2026 年上線 | S#45 | 屬實 |
| cp_orcl_shk1／shk2 | 全站 2,000 設施（200@2027-08＋1,800@2028-11） | Vantage：1.4GW、10 棟、首棟 2026 下半年；Oracle 稱客戶交付 2027 上半年起 | S#56、S#57 | 小誤（basis、user）、無法取得 |
| cp_orcl_sal | 1,400 設施 @2028-12 整筆 | MPSC 合約負載 1,383；Oracle 稱 2H27 開始交付、Related 稱 2027Q1 首批 | S#65 | 無法取得 |
| cp_orcl_jup | 已知延遲 12 月 | Oracle 稱進度不變、Phase 1 2028Q3 | S#60 | 原文找不到（延遲欄） |
| cp_oai_pp1–pp3 | 4,250 IT 三年均分（1,417 起 2028-12） | OpenAI 稱 2028 年首批 800MW；租約 2028–2032 分期 | S#150、S#148 | 無法取得、無法取得 |
| cp_sov_hum1 | 200 @2026-06 | 至 2026-09-27 無公開運作確認（逾期說法不在原文） | S#163 | 小誤 |
| cp_sov_hum2 | 1,700 @2030-12 | Al Sa'ad 一期 1GW 2027 交付 | S#158 | 屬實 |
| cpx_sov_uae（納入 0） | 200 @2026-09 | 2026 Q4 | S#4 | 小誤（不涉 MW、日期） |
| cpx_xai_c2（納入 0） | 585 @2027-12，註「1,531 無出處」 | 1,531 IT @2027Q1（DCD 轉述 Epoch），946 已運轉 → 增量 585 @2027-03 | S#135、S#134 | 小誤（tag）、小誤 |
| cpx_sov_mis（納入 0） | 1,400 @2028-12 | 2028 初只有 Mistral 96MW（可擴至 200）；Phase 2 為 2030–2033 | S#153、S#154、S#155 | 屬實、屬實、不符 |
| cpx_nbis_lpr（納入 0） | 310 @2027-12 | 310 為完工容量，2027 只有首期 | S#81 | 小誤 |
| cpx_nbis_mo（納入 0） | 1,200 @2027-12 | 首棟 200MW；頁面時程 2026H2–2029+ | S#83、S#82 | 小誤、小誤（expected_online） |
| cpx_crwv_lan（納入 0） | 300 @2027-12 | 300 為擴充潛力無時程；初期 100MW | S#80、S#79 | 小誤、小誤 |
| cpx_oai_milam（納入 0） | 857 IT @2026-10 | 1、2 號樓租約 800MW（2027）；DC01 2026-10-15 完工但無 MW；全站 1.2GW | S#143、S#144、S#142 | 口徑不同、口徑不同、小誤 |
| cpx_neo_nexus（納入 0） | 500 @2026-12 | TDLR 工期至 2027-10，首棟 2027 啟用 | S#130、S#131 | 不符、小誤 |
| cpx_amzn_cip（納入 0） | 216 IT @2026-10 | 約 300 gross，2026-08 首批提前交付，全站 2026Q4 | S#42、S#43 | 無法取得、無法取得 |

一致（無矛盾）的既有列：cp_msft_wi2、cp_amzn_nc1／nc2、cp_meta_hyp1、cp_meta_elp、cp_orcl_abl、cp_orcl_pw、cp_crwv_hel2、cp_neo_iren、cp_neo_twf1／twf2、cpx_oai_lord（214 IT @2027Q1）、cpx_neo_hut8（245 IT @2027Q2）。無新證據：cp_neo_iren27、cpx_googl_bx、cpx_neo_cer。

## 3. 建議升級的既有 納入=0 列（不在本提案 inc=1 內，避免重複）

| 既有鍵 | 建議改為 | 依據 |
|---|---|---|
| cpx_sov_uae | 200 未說明 @2026-12，納入 1 | S#4 小誤（不涉 MW、日期） |
| cpx_xai_c2 | 585 IT @2027-03，納入 1 | S#135 小誤（只涉 tag） |
| cpx_sov_mis | 96 未說明 @2028-03，納入 1（其餘 1,304 不入） | S#153 屬實、S#154 屬實 |

三列合計 881 MW；若一併採用，新增納入量為 5,898＋881＝6,779 MW。

## 4. 規則套用說明

- 納入=1 的 11 列，主證據判定皆為屬實或小誤，且 API 小誤欄位不含 value／basis／expected_online／unit；上線年月皆在 2026-10…2032-12；去重皆為「新」或「補充」；無 cn／oth、無晶片合約、無發電容量。
- 只有年份者月取 12（Hyperion 2032、PF2 2027）；「2028 初」取 3（Monarch、Fouju）；Q1→3、Q3→9、Q4→12；日本 2026 年度→2027-03。
- 未換算口徑：Monarch 1,350、Hyperion 3,000、蔚山、苫小牧、Oxagon 保留「未說明」。
- 非中國境內的中資租戶（ByteDance Pecém）依 colo 規則歸 oth，不歸 cn。

## 5. 待解問題

1. **Hyperion 3,000MW 一列占新增量一半**：證據只有發言人口頭說法「約 2032 年完成 5GW」，Entergy 新電廠待 LPSC 2026-12-16 表決（S#50 附註，非核對範圍）。規則上可納入，但建議 Andy 判斷是否先只放到 2GW。
2. **Ridgeland 與 Canton 的 284MW 混淆**：S#35 註「另一聚合站將 284 MW 歸給 Canton」，S#38 Canton 也是 284 IT。本提案只納入 Ridgeland B6–14（556），Canton 至 2028Q1 的 399 未納入（S#40 口徑不同，主因 tag）；兩個 Epoch 頁面需確認確為不同園區後，Canton 可升級。
3. **SoftBank 苫小牧歸 oai**：依指示 SoftBank 歸 oai，但此站為 SoftBank 自用／大學研究（S#333），會灌入 OpenAI 桶。若 oai 只指 OpenAI 可用算力，建議改 oth 並改為納入 0。
4. **Khazna Ajman tag**：數字出自第三方 Blackridge（API 指 [Assumed] 不當），本提案標 Verified，實為第三方轉述而非一手，可改其他標籤。
5. **營運者鍵**：Galaxy Helios、Polaris Forge 1、Core Scientific Denton 依既有 cp_crwv_hel2 先例用 crwv（CoreWeave 為承租人，使用者鍵無 crwv）；Polaris Forge 2、Beacon Point、Barber Lake 承租方未證實，用 neo、使用者留空（媒體分別稱 Oracle、Anthropic、未具名 AI 實驗室）。
6. **已確保電力**：本提案沒有列可確證。可補查：Shackelford VoltaGrid 700MW 現場燃氣已獲准運轉（S#58，無法取得）→ 可能讓 cp_orcl_shk1 改為 1；Hyperion Entergy 電廠尚待 LPSC 表決（S#50）→ 既有 cp_meta_hyp1 標 1 可能過早。
7. **小誤只因 basis／tag 而未納入、可快速升級者**：Polaris Forge 1 B2 P1 75MW @2026-07（S#72，改標未說明即可）、Ridgeland B5 56 @2027-02（S#35 不符但 API 已給正確日期）、Mesa B2 93 IT @2026-08（S#28）、Hut 8 River Bend 245（S#115 MW、日期屬實）。
8. **2026-07…10 已上線但來源 403**：IREN H1 50（S#87）、Lake Mariner CB-3 42（S#111）、Nscale–Ionic 234（S#96）、Core Scientific Denton 餘 128（S#75）。重新取得 SEC／IR 原文即可判定。
9. **Humain 重疊**：DataVolt Oxagon 的 Nebius 50MW 位於與 HUMAIN 合作的 100MW 區塊（S#271），若 Humain 1.9GW 目標含合作區塊，cp_nbis_oxagon 與 cp_sov_hum2 會小幅重疊。
10. **IREN 公司層級列**：cp_neo_iren27（2027 累計 0.8GW）可能已含 Sweetwater 1 首期 300 gross（S#260），故 Sweetwater 未提新增。

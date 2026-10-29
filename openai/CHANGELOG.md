# CHANGELOG

每次同步記錄：Excel 版本、工作包、commit 與變動摘要。

## 20261029_OpenAI_v0.6.4.xlsx — v0.6.4（口徑一 Q2 GW 採 B／C 段、D2 推導式、衍生值移 Inputs、TK 快照 v21.7）

- **r1（獨立核對後，同版不升版，只改文字與來源欄，數值不變）**：PxQ X227 說明改與 INP_303 同義；SRC_OAI_146／147／148 列層級覆寫為等級 1／一手並改立場說明；Checks C175 標籤去掉『v0.6.3』（編號不變）；W6 年化欄註明『顯示用，不參與計算』。
- **W1** 口徑一 Q2 季平均 GW（INP_289–291）：2.34／2.21／3.3 → **2.26／2.17／2.35**（Tokenomics B 段 Fairwater 占比 f＝0.4／0.65／0.9；C 段 10-29 確認不改）。倍數 0.999／0.959／0.921、比值 1.237／1.188／1.142。J7 六情境列（INP_295–300）不改。PxQ 第八節 X228 的 (A)(B) 欄改取 INP_297（證據原 GW 2.34），(C) 欄與 4.1 區塊隨本模型基準；X232 (A) 核對改為『同式、同 GW』；X205、Checks C175 預期值改 0.959／0.96。
- **W2** PxQ 新增第 8.4 節 D2 推導式（Codex 份額 × ChatGPT token 量；候選因子，不進基準 token 帳）：新增 SRC_OAI_146–147（arXiv 2606.26959 份額 0.633／0.165）、INP_308–314（w 三列、時點比、k 兩列、D3 6,000）；(A) 證據重現欄對 B 段 token帳 B14:D19、B34 全 OK（5.071、7.606；token 帳 k＝1 為 1.857／1.769／1.689，k＝2 為 1.666／1.530／1.414）。基準 token 帳 X236 (A)(B) 2.0966／1.9876 不變、(C) 隨 W1 1.9365→1.8703。
- **W3** INP_303（1.397）值不變，註記改為『＝1 ÷ Tokenomics `L1_LReqFleet_Full_Lo`（0.7156）』，刪『起源待確認』。
- **W4** SRC_OAI_141–143 退役（`id_registry.json`）；改為 Inputs 公式列 INP_315–320（廣告 Q2／FY2026、每年季數、含廣告 26.8、不含廣告 26.3、FY2026 預測 35；Derived 列的低／高由廣告區間反向帶出）。口徑一／二營收與 v0.6.3 逐格相同。
- **W5** `TK_Link` 改取 Tokenomics master `8309fdb`（CURRENT＝`20261025_Tokenomics_v21.7.xlsx`）：88 名數值與 v21.4 逐名相同（check_tk_snapshot OK 88／DIFF 0／MISSING 2〔`IF_PxEffMix_Lo／_Hi`〕）。
- **W6** SRC_OAI_148（Microsoft FY27 Q1 自 OpenAI 營收 10.0，10-Q 關係人附註）；PxQ 第九節只顯示、不參與任何計算。
- builder：`preserve.merge` 新增 `retired` 參數（已退役 ID 不算 unmatched）；Inputs 公式佔位 `{INP:}`／`{SRC:}`／`{INPF:}`／`{INPG:}`（`finalize_ids` 後換成最終 ID）。Checks 新增 C178–C180，C164 標籤改為 v21.7（編號不變）。工具 `tools/diff_v064.py`（範圍外變動比較）。
- 測試：test_builder 新增 `test_w1_scope1_gw_values`、`test_w2_d2_derivation_matches_b_segment`、`test_w4_derived_revenue_inputs`；scenarios.yaml 新增 bd–bf、期望值更新。`dist/` 重產，v0.6.3 移 `dist/_待刪除/`；v0.6.3 xlsx 移 `model/archive/`。報告：`docs/reports/20261029_v0.6.4.md`；PR #235。

## 20261025_OpenAI_v0.6.3.xlsx — v0.6.3（PxQ 比值 2026 三口徑、TK 快照 v21.4、GW 路線 token 帳）

- `TK_Link` 改取 Tokenomics master `693792f`（CURRENT＝`20261024_Tokenomics_v21.4.xlsx`）：88 名有值、2 名待提供（`IF_PxEffMix_Lo／_Hi`）。check_tk_snapshot（v19.0 快照 vs v21.4）OK 53／DIFF 35／MISSING 2。除 PxQ 以外各頁計算邏輯不改。
- `PxQ` 頁改為 11 欄（D–N）：2025 指數／線性 + 2026 三口徑 × 低／基準／高。口徑一＝Q2 年化營收 26.3 ÷ Q2 季平均 GW（2.21／2.34／3.3，主口徑）；口徑二＝FY2026 營收 35 ÷ 全年平均 GW（2.5／3.0／3.8）；口徑三＝9 月底 run-rate 50 ÷ 年底 GW（3.5／4.6／5.7）。頂部三張主輸出：比值、對 2025 主值的倍數、持平門檻 GW（以 2025 線性 1.118／指數 1.239／上緣 1.397）。
- 新增第 4.6／4.7 節（管理層 36、WSJ 26.8 營收對照）、6B 證據對帳（A 段／J7／工作單驗收值）、第七節 Fairwater 六情境與 2025 站點估（J7）、第八節 GW 路線 token 帳（交接檔組合×0.85＝2.0966、×0.806＝1.9876，容差 1%）。6A 交接檔 v18.0 預覽對帳改為歷史對照、不再計數。
- 新增 SRC_OAI_141–145、INP_289–307；INP_277–279 改名為口徑二 GW（值不變）。Checks 新增 C165–C175、退役 C156（見 id_registry）。
- 測試：新增 `test_pxq_three_scope_structure`、`test_w3_gw_route_token_ledger`、`test_fairwater_six_scenarios_match_j7`、`test_pxq_recon_6b_clean`；scenarios.yaml 期望值更新並新增 az–bc 四個情境。
- `tools/build_html.py`：回本比值區塊（⑥′）顯示三口徑；`dist/` 重產，v0.6.2 移 `dist/_待刪除/`。工具 `tools/tk_attribution.py` 加細 Q 面分組。報告：`docs/reports/20261025_v0.6.3.md`；PR #104。

## 20261022_OpenAI_v0.6.2.xlsx — v0.6.2（Tokenomics 同步 v5.26 → v19.0；單一快照）

- `TK_Link` 改取 Tokenomics master `1916a06`（CURRENT＝`20261022_Tokenomics_v19.0.xlsx`，SHA-256 `d1728705…4b8d`）：88 名有值（原 63＋P×Q 取數名稱 22＋`RNG_IF_AllocQ1_Lo／_Hi`、`RNG_StaleFlag`）、2 名待提供（`IF_PxEffMix_Lo／_Hi`）。`TK_PxQ` 頁與 `TKQ_` 名稱退役，`PxQ` 公式改引用 `TK_`；`PxQ` 776 個數值格與 v0.6.1 逐格相同。builder 移除 `--tk-pxq-dir`；`tk_link.read_pxq_snapshot` 移除。
- **命題結論不變**。TK 變動的影響（對照 Excel ②，LibreOffice 重算 10 變體歸因）：η 2025 0.2795→0.1684；2030 有效推論 GW 5.60→6.84（+22.3%）、研發 GW 8.88→7.64、研發占總需求 61.4%→52.7%；供應商持有成本 2030 −7.5 $B、雲端毛利 2030 2.6（1.6%）→10.1（6.4%）；命題差額（現金）2025 −66.4→−56.9、2030 −10.05→−10.03；累計外部資金需求 442.8 與峰值年 2029 不變；截頂係數仍為 1。全部差異來自 `IF_TokGW_*`（推論／研發 GW、VR 等值係數）與 `IF_HoldEcon`（持有成本、雲端毛利）；`IF_CapexTotal` 無影響。表 2（範圍外變動）＝0。
- Checks 新增 C155–C163（單一快照版本、名稱數、SHA、研發占比對照與 `IF_AllocQ1` 區間、`RNG_StaleFlag`）；退役 C132–C138。`IF_AllocQ1` 0.636→0.184 的口徑改變與 `L1_Ans3` 1.157→1.907 寫入報告第 5 節；`L1_Ans3` 因 v19.0 無 `IF_` 等義名稱而保留。
- 測試：`test_pxq_reference_whitelist` 改為 `TK_` 且限 P×Q 取數清單；新增 `test_tk_link_single_snapshot_v19`；`scenarios.yaml` 期望值更新（具名範圍 709、tk_values 85、tk_pending 2）；`test_html` 敘述主張 6–9 成改 5–9 成。parity 113 項、`test_builder` 16 項、`test_html` 5 項全過。`dist/` 重新產生，v0.6.1 成品移 `dist/_待刪除/`。工作單 `docs/workorders/20261022_openai_TK同步v19.md`；PR #62。

## 20261022_OpenAI_v0.6.1.xlsx — v0.6.1（P×Q 與回本比值序列）

- **v0.6 原有數值一格不變**（表 2＝0）；新增 `PxQ` 頁、`TK_PxQ` 頁（Tokenomics v19.0 第二快照，master `ad33e28`，檔案 SHA-256 `d1728705…4b8d`；具名範圍 `TKQ_`）。`TK_Link` 維持 v5.26（63 名逐格相同）。v0.6 原檔移至 `model/archive/20261008_OpenAI_v0.6.xlsx`。
- 主輸出：回本比值（實際 m×L ÷ 所需 m×L＝營收 ÷ 算力成本）：2025 指數／線性 1.24／1.12；2026（2.5／3.0／3.5 GW）1.47／1.23／1.05。持平所需 2026 GW 2.97（指數基準）／3.29（線性基準）。中間量（h、C、Q、P_eff、實際／所需 m×L、隱含 m）、四個敏感度區塊、外部支出對照倍數（Epoch 160 億、Brockman 500 億、204 億）、預覽值對帳（15 項中 14 項在 0.5% 內，Q 一項因預覽值只給兩位有效數字）。
- 新增 SRC_OAI_138–140（Epoch 2025 雲端 16、Brockman 2026 的 50、Epoch 服務占比 0.50）、INP_276–288（GW 主值開關、2026 GW 三情境、組合來源開關與交接檔六列、折減讀法開關、對帳容許誤差）、Checks C131–C154（C140、C141 退役）；公式格 3,063→3,898；具名範圍 646→719。
- builder：`--tk-pxq-dir`、`--tk-date`；`tk_link.read_pxq_snapshot`；`builder/p6.py`。`engine/core.py` 檔名正則接受 `v0.6.1`。`dist/` 成品隨現行版重新產生，v0.6 版移 `dist/_待刪除/`。
- 測試：parity 新情境 aq–ay（C132、C142、C154 有牙齒）；`test_e6_no_constants_in_pxq_formulas`、`test_pxq_reference_whitelist`、`test_pxq_not_fed_back`、`test_tk_link_values_unchanged_from_v06`。工具：`tools/diff_workbooks.py`、`tools/make_pxq_report.py`。報告：`docs/reports/20261022_v0.6.1_PxQ.md`、`_對照.xlsx`。PR #59。

## 成品 v0.6（S6：HTML 一頁摘要、命題與驅動對照表、交接檔；Excel 不變）

- Excel 仍為 `20261008_OpenAI_v0.6.xlsx`（本段不改模型；parity 107 項全過）。複製到 `dist/20261008_OpenAI收支模型_v0_6.xlsx`（與 `model/CURRENT` 位元組相同）。
- 新增 HTML 一頁摘要 `dist/20261008_OpenAI收支模型_v0_6.html`（單一檔案、離線可開、內嵌 CSS 與 SVG、系統字型、無外部 URL）：①命題結論 ②每 VR 等值 GW 營收 vs 全成本 ③營收結構與管理層目標 ④算力需求 vs 供給 ⑤現金與外部資金需求（Amazon 情境並列）、或有負債 ⑥三個關鍵驅動與敏感度 ⑦最該審的 5 項預設 ⑧資料與版本。374 個數字皆帶 Excel 位置（title、圖下小字）。
- 產生器 `tools/build_html.py`（LibreOffice 重算後以具名範圍或「頁!列 ID」取值；HTML 不含計算）；敏感度情境 `tools/html_scenarios.yaml`（值取自 Inputs 低／高欄與倍數列，15 個情境）；截圖工具 `tools/screenshot_html.py`（1280／390 × 淺／深色，攔截外部請求、檢查水平捲動）。
- 新增 `tests/test_html.py`（HTML 374 個數字與 engine 逐一比對、無外部引用、dist xlsx＝現行模型、文字主張）並納入 CI governance job。
- 新增 `docs/plan/20261008_OpenAI_命題與驅動對照表.md`；交接檔更新為 v0.6 完成狀態。報告：`docs/reports/20261008_v0.6_成品.md`。PR #20。

## 20261008_OpenAI_v0.6.xlsx — v0.6（v0.6-P5；S5：融資、檢查與反向模式）

- **v0.6 完成**。TK_Link 快照改取 Tokenomics master `3dd1216`（CURRENT＝`20261007_Tokenomics_v5.26.xlsx`，SHA-256 `abb8592c…224ed0d`）；63 名數值不變；v5.26 新增 DC_Cost 名稱不引用。
- 新增 Funding 頁（F01–F38；沿用 v0.5「資金與檢查」S7／S4b）：自由現金流（營收淨額 − 算力成本 − 非算力成本；股權報酬加回）、已到位融資（2026-03 輪無條件部分 87；Nvidia 現金比例）、最低現金、外部資金需求（年底現金低於最低現金的累計差額）、缺口旗標、只靠已到位融資的現金路徑、Amazon 條件式 $35B 情境（不計入基準）、或有負債（只列示）。
- 新增 Reverse 頁（X01–X49；R1–R3，只作對照、不回饋基準）：管理層目標路徑（2027–28 擬合值 Inputs、2029＝Σ840 殘差）、三組充分條件倍數、2030 可觀測量、橋接、反向資金、與簡報對帳。
- Checks 新增 C91–C130（含 2025 實際值對帳、TK 快照狀態；排序改為數字排序）；Inputs 258→263（INP_271–275）；v0.5 葉節點 `reverse.targetRevenue.v[2]`、`[3]` 由公式改為 Inputs；具名範圍 619→646；公式格 2,629→3,063。SRC_OAI 不變（131）。
- 主要結果：自由現金流 2026–2030 −39.4／−116.0／−134.9／−137.6／−131.9；已到位融資撐到 2026 年底；2027 起現金跌破最低現金；累計外部資金需求 442.8（峰值 2029 年 137.6）；計入 Amazon 情境 407.8；或有負債 600（只列示）。反向模式 2030 倍數：只靠 API 4.01、只靠訂閱 5.80、等比例 2.85；反向累計外部資金需求 92.7。
- 測試：新增情境 af–ap（11 個；C94、C95、C96、C122 有牙齒）；`test_e6_no_constants_in_p5_formulas`、`test_reverse_not_fed_back`；範圍外變動 0。S4 版改名移至 `model/archive/20261007_OpenAI_v0.6-P4.xlsx`。P4 報告第 185 行小修。報告：`docs/reports/20261008_v0.6-P5.md`（v0.6 完成報告）、`_對照.xlsx`。PR #19。

## 20261007_OpenAI_v0.6.xlsx — v0.6-P4（S4：成本、Q3 與單位經濟）

- TK_Link 快照改取 Tokenomics master `97e7b20`（CURRENT＝`20261007_Tokenomics_v5.25.xlsx`，SHA-256 `43341b8d…c59127`）；62 名值不變，新增 `IF_CapexTotal`（63 名）。
- 新增 Cost 頁（K01–K106）：算力成本（V3：逐合約實付＋自有資本支出；現金口徑）、供應商持有成本（TK `IF_HoldEcon` 世代加權；自研＝VR200 × genFactor.custom）與雲端毛利、經濟口徑、非算力成本（V4：2025 財報；2026 起人數 × 每人年成本，股權報酬單列）、Q3（對 TK `L1_Ans3`）、命題輸出（每 VR 等值 GW；現金基準、經濟並列；TK 單位成本參考）、V16 Azure 攤入檢查。Compute 新增第十三節自建 GW（計入供給與命題分母；G72 改為合約＋自建），研發 GW 殘差 2030 6.11→8.88。
- SRC_OAI 122→131（SRC_OAI_129–137：2025 銷售／管理／營業損失、人數、股權報酬、FT 人數）；Inputs 249→258（INP_262–270）；Checks 新增 C67–C90（C50 公式含自建 GW）；具名範圍 562→619；公式格 1,989→2,629。
- 主要結果：每 VR 等值 GW 差額（含股權報酬、現金口徑）2025–2030：−66.4／−42.1／−22.3／−16.7／−12.7／−10.1（每年皆為負）；經濟口徑 2030 −7.8；2025–2030 累計差額 −650B；雲端毛利率 19%→1.6%；Q3 0.465→0.509；V16 旗標 0。
- 測試：parity 83 項全過（新增情境 x–ae；E6 掃描擴及 Cost）；範圍外變動 0。合約實付公式以 IF 防除以 0（CI 第一次失敗原因）。S3 版改名移至 `model/archive/20261007_OpenAI_v0.6-P3.xlsx`。報告：`docs/reports/20261007_v0.6-P4.md`、`_對照.xlsx`。PR #15。

## 20261007_OpenAI_v0.6.xlsx — v0.6-P3（S3：算力需求與供給）

- 新增 Compute 頁（G01–G133）：V1 世代組合（Blackwell 拆 GB200／GB300；自研＝VR200 × genFactor.custom）、付費／免費每 GW 年產能（TK `IF_TokGW_*` × `IF_Util`，Tokenomics Alloc 同式）、推論 GW（token 換算）、V2 η（2025 推論支出 ÷ 合約價校準；基準沿用，情境線性回升至 2030＝1）、有效推論 GW、三角對照、研發 GW（v0.5 殘差法）、總需求 GW、供給 GW（逐合約；v0.5 S2、N4(b)）、容量上限（V1a）、VR 等值換算、敏感度（合約價、η 路徑、GB300 占比、自研產能）。Revenue 新增第八節 R46–R52（容量上限係數、截頂後總額／分成／淨額；未截頂值保留）。
- Inputs 243→249（INP_256–261）；Checks 新增 C40–C66（C05 期望 243→249）；具名範圍 516→562（`CMP_*`、`REV_*Capped`）；公式格 1,260→1,989。SRC_OAI 不變。
- 主要結果：η（2025）0.2795；有效推論 GW 2025 0.70 → 2030 5.60；合約供給 1.23 → 13.01 GW；2026–2030 皆未截頂（2030 餘裕 46%）；2025 支出換算總需求 1.70 GW 對年均 1.25（+0.45）；三角對照 16.32 $B/GW/年（+36%）。
- 測試：parity 66 項全過（新增情境 p–w；E6 掃描擴及 Compute）；範圍外變動 0。C46 以最大絕對值寫法避開 pycel `SUMPRODUCT(ABS())` 型別差異。S2 版改名移至 `model/archive/20261007_OpenAI_v0.6-P2.xlsx`。報告：`docs/reports/20261007_v0.6-P3.md`、`_對照.xlsx`。PR #14。

## 20261007_OpenAI_v0.6.xlsx — v0.6-P2（S2：需求與營收）

- 新增 Demand（D01–D48）、Revenue（R01–R45＋第七節價格事件時點表）兩頁；機制沿用 v0.5「訂閱與廣告」「API」「營收彙總」頁。訂閱（T4 遷移、pro 依 V12／V15）、任務 × 每任務 token、API（2025 由營收倒推；2026 牌價依價格事件時點天數加權；2027 起年變動與價格彈性）、廣告、其他、總額、Microsoft 分成、淨額；token 量（層級 × 付費／免費）具名範圍 `DEM_Tok_*` 供 P3。
- SRC_OAI 106→122（價格事件數值與公告日，SRC_OAI_113–128）；Inputs 233→243（E8c 事件時點偏移 INP_246–249、E6 定義常數 INP_250–255）；Checks C27 WARN 撤除（退役），新增 C28–C39；具名範圍 426→516；公式格 652→1,260。
- 主要結果：2025 營收對 SRC 實際差距 0（校準），未校準 +1.03B；2026 總額 34.20B；2030 總額 126.38B、淨額 126.38B（分成上限 2029 觸頂）；2025 token 3,538T＝TK `IF_AllocDemand` 的 0.84 倍。
- 測試：parity 49 項全過（新增情境 h–o；E6 掃描擴及 Demand、Revenue）；範圍外變動 0。S1 版改名移至 `model/archive/20261007_OpenAI_v0.6-P1.1.xlsx`。報告：`docs/reports/20261007_v0.6-P2.md`、`_對照.xlsx`。PR #10。

## 20261007_OpenAI_v0.6.xlsx — v0.6-P1.1（S1：TK_Link 快照更新至 Tokenomics v5.24）

- 快照來源：Tokenomics master `bdb0de7`（CURRENT＝`20261007_Tokenomics_v5.24.xlsx`）。check_tk_snapshot OK 37→62、MISSING 11→0、DIFF 0。原待合併 11 名全部取得；新增 IF_TrainCost_*、IF_RevGW_*、IF_FullCostDefault_*、L1_Ans3（＋低高）、SRC_DEM_013_Lo／Hi。原 37 名中 10 名值變動（TokGW×3、HoldEcon、FullCost×3、ProgGWyr×3）；第 5 世代口徑 Rubin Ultra NVL576→MGX NVL 單架 72 GPU。
- 只動 TK_Link、Checks（C07、C08 標籤與期望值；ID 不變）、README 兩格；具名範圍 401→426；SRC_OAI、Inputs、Derived_V9、Map_v05 逐格不變。舊檔移至 `model/archive/`。
- 工具：`tools/diff_tk_snapshots.py` 新增；`builder/tk_link.py`、`build.py` 沿用舊 PR #2（`479e94b`）改動。parity 32 項全過。報告：`docs/reports/20261007_v0.6-P1.1.md`、`_對照.xlsx`。PR #8。

## S0 搬遷（2026-10-07，Excel 不變）

- 由 `YinchenChang/openai-model` `main@13c16f0` 搬入 `company-models/openai/`；CI 移至根目錄 `.github/workflows/openai-parity.yml`（working-directory `openai`）。新增共同規則、S1–S6 工作單、進度檔、交接檔（`docs/handoff/OpenAI_handoff.md`，取代 Project 內舊交接檔）。parity 32 項全過，Excel 逐位元組相同。

## 20261004_OpenAI_v0.6.xlsx — v0.6-P1（repo、Source 與 Tokenomics 連結）

- 報告：`docs/reports/20261004_v0.6-P1.md`。Commit：見本 PR 的合併提交（合併後補上雜湊）。
- 新增：repo 骨架（builder／engine／tests／tools／CI）、CLAUDE.md；Excel 8 頁：README、SRC_OAI（106 列）、TK_Link（Tokenomics v5.14 快照，37 個名稱有值、11 個待 v5.15 合併）、Inputs（233 列）、Derived_V9、Checks（26 項，編號凍結）、Map_v05（849 個 v0.5 葉節點去處）、_Defaults（隱藏）。穩定 ID registry（`builder/id_registry.json`）。
- 測試：parity 24 項全過；v0.5 每個葉節點有去處；`check_tk_snapshot`：OK 37、MISSING 11（WARN）。
- 依工作單 r2：V6（SRC 10.59、訓練支出改公式、其他雲端 1.41／0–2.91）、V8（合約價單一列 12／8–20）、V9（跨年拆開）、V10（四欄初評）、V11；E4（PR）、E5（環境安全）。詳見報告第 2 節。
- 依工作單 r3（內容依 chat 訊息摘要，r3 原文未在 main）：E6（V9 公式不含常數、換算係數進 Inputs）、E7（帳目四項）、V12（pro 改正）、V13（Cerebras／Azure 終點與總額 Assumed＋區間）。V14、E8 待 r3 原文。詳見報告。
- 依工作單 r3 補做與 V15：V14／F2（來源初評規則）、E8 a–j（待判斷 A 類全數處理）、F1（Checks 編號 registry）、V15（pro 2027–2030 改公式）。詳見報告。
- 審查補充（G1–G8、V16）：合約列立場初評（未評 11→5）、Checks C27 WARN、定義常數註記、Azure 起點區間 2025–2026。詳見 `docs/reports/20261004_v0.6-P1_supplement.md`。
- 審查補充二（H1–H4）：SRC_OAI_105–108 補出處（OpenAI 與 NVIDIA／Broadcom／AMD 聯合公告）、「同上」出處追溯、一手／二手圖例、SRC_OAI_082 立場說明。V10：立場未評 1（082）、等級未評 0、一手／二手未評 0。

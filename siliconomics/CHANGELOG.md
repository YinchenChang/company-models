# Siliconomics 變更紀錄

## v0.16（2026-11-05；`model/20261105_Siliconomics_v0.16.xlsx`；Tokenomics v21.10）

- **情境過渡完成：產業情境 A、C 移到市場層 `markets/scen/`**（決策 S41）：v0.15 In 頁的 70 列（記憶體價格路徑 cqA／cqC、nqA／nqC 各 12 季；顆數係數 um28／um29 的 A、C；HBM 單價 hbm_p28／29 的 A、C；世芯 Trainium 份額 al27／al28 的 A、C；CPO 滲透率 cpo A、C 2026–29；設備係數 eq_ck C 2028–29）逐字搬到 `markets/scen/industry_scen_v0.1.json`（數值、標記、來源不變；`markets/scen/README.md` 有對照表）。In 675 → 605 列；Scen 頁只剩基準；Units、Deploy、ABF、PGW、Rack、Equip、Path、Mkt、ShareTbl、Checks 只算基準（K1 只比對 SK 海力士成品的 B、K2 只查基準係數區間、K6-trn／pwr／lc 只查基準）。
  - 讀者：`markets/memory/market.py`（`SCEN`、`SCEN_DESC` 改由市場層提供；`merged_rows()` 把 A／C 列插回原列序，記憶體公司 In 頁重播逐列相同；`write()` 新增 `scen` 參數，本庫只寫基準）、network、ase、mpi、alchip、tuc、arm、umc、vis 改讀 `markets/scen`。`tools/r4_bridge.py` 的每 GW IT capex 改依列名定位（Deploy 頁列位改變）。
  - 唯一數值變動：Interface `SIF_TestWFE_SoC` 2028–29 由原情境 A 改為基準（0.091／0.094 → 0.078／0.083；無下游讀者）；Checks K8-2028 由 WARN 轉 OK（SGOV_Warn 5 → 4）。
- **Cap 頁分項上限**（決策 S42；算力收支 11-04 問題的後續）：新增「只受 CoWoS／只受 HBM 限制的可支撐部署 GW」兩列（基準 × 產能 ÷ 需求，不截頂）與 Interface `SIF_CapGW_CoWoS_Base`＝28.9／43.3／59.8／75.1、`SIF_CapGW_HBM_Base`＝26.2／38.5／48.8／57.8（2026–29，本庫部署口徑）；`SIF_CapOnlineGW_CoWoS_Base`＝25.7／37.6／65.9／76.3、`SIF_CapOnlineGW_HBM_Base`＝23.3／33.4／53.8／58.7（Wattonomics 可比上線口徑，毛額）。`SIF_CapGW_Base`、`SIF_CapOnlineGW_Base` 不變。
- **TK_In 改 Tokenomics v21.10**（決策 S43）：22 個名稱中只有 `IF_NonNV_Out_TPU`、`IF_NonNV_Out_Trn` 改值（本庫與下游不讀產出比）；chip_bridge 仍讀 v21.7（差異格皆為產出比）。
- R4 r6（讀 Wattonomics v0.3）：與 r5 逐值相同；Recon 快照改 `recon_snapshot_v0.16.json`（v0.14、v0.15 未重抽，只差來源檔名）。
- 下游：cascade 第一輪 63 個 CHANGED（全部為版本與來源文字）、umc 首輪建置失敗（讀 `um28_A`）已改讀市場層；第二輪只有 umc CHANGED（文字），第三輪全部 SAME。投資建議、目標價全部不變。

## v0.15（2026-11-05；`model/20261105_Siliconomics_v0.15.xlsx`；Tokenomics v21.7）

- **刪除 Interface 推導節的情境 A／B／C 名稱**（v0.12 起停用過渡；決策 S40）：66 個名稱（含 _lo／_hi），快照 2,067 → 2,001 個名稱；頁內列保留並標「名稱已刪除」，`*_Base` 仍指向原 B 列。In 頁層情境列（um28_A、cqA_8、cpo_A_26 等）名稱保留，作為個股自訂情境參考值（v0.16 移到個股）。
- 讀者改寫（數值不變）：tuc（部署 GW、CCL 市場＝基準 × 個股情境係數 ÷ 基準係數）、arm（顆數係數由 um28／um29 列組成）、ase、mpi（原始顆數改讀 `SIF_UnitsPlat_Total`；ase Recon 部署 GW 改讀基準〔展示列〕；mpi SoC 測試機三情境改讀 Advantest 成品）、smci、dell、jentech、ter（一致性檢查只核對基準）、tsmc（HBM 2027 缺口改讀基準）。Wattonomics v0.2.2 起只讀 `*_Base`（算力收支 11-04 確認）。`tools/r4_bridge.py` 改讀 `SIF_UnitsCoef_Base`。

## v0.14（2026-11-05；`model/20261105_Siliconomics_v0.14.xlsx`；Tokenomics v21.7）

- **HBM 供給基數改用廠商出貨推估**（Andy 2026-11-05 08:22 選甲；決策 S39；報告 `docs/reports/20261105_HBM供給校準.md`）：新增 In 頁 M3b `hbm_sup26`＝4.23 EB（3.8–4.6；UBS 廠商出貨推得，Derived）；2026 供給改讀此值（原＝本庫 2026 需求 3.78），2027 年增仍 +55%（TrendForce）→ 2027 供給 5.86 → 6.56 EB。markets/memory `_INPUT_SECS` 加 M3b（記憶體三家 In 頁重播）。
  - 2027：HBM 需求 ÷ 供給 121% → 108%；HBM 出貨 +12%；HBM 市場 154.1 → 172.3 $B；可支撐部署 GW 34.4 → 38.5（K10 仍 WARN）；`SIF_CapOnlineGW_Base` 29.9 → 33.4。2026 出貨不變（2026 HBM 供給成為獨立數據，需求 ÷ 供給 89%）。2028–29 不變。
  - 下游加權目標價：SK 海力士 +2.6%、三星 +1.4%、美光 +1.5%、泰瑞達 +0.9%、BESI +0.5%、Advantest +0.1%；Disco −1.3%、韓美 −2.9%、TOWA −3.2%（CY2027 出貨提高使 CY2028 增量變小）；投資建議全部不變。cascade 第二、三輪全部 SAME。

## v0.13.1（2026-11-04；`model/20261104_Siliconomics_v0.13.1.xlsx`；Tokenomics v21.7）

- **分世代每 IT GW NVIDIA 內容**（算力收支 Wattonomics Requests R7，期限 11-12；決策 S38）：In 頁新增 M28（NVIDIA 合計 ÷ IT capex：GB300 0.647、VR200 2026 0.624、VR200 2027 0.581〔高值 0.599＝HBM 漲價轉嫁〕取自晶片內容橋接表；Hopper、GB200 以 GB300、Rubin Ultra 以 VR200 2027 代入〔Analogy〕）；Interface `SIF_NVContentGW_<H100|GB200|GB300|VR200_26|VR200|RU>`（$B／IT GW）：17.80／15.54／20.85／23.47／21.85／29.11（含 _lo／_hi）。
- **Wattonomics 可比口徑的供給上限** `SIF_CapOnlineGW_Base`（可支撐新增上線 IT GW，毛額）：20.81／29.89／46.50／51.91（2026–29），供 Wattonomics v0.3「部署 GW 不得超過晶片產能」檢查。
- **逐買方 NVIDIA 份額**（算力收支 11-04 Wattonomics v0.2／v0.2.1 訊息第 3 節，期限 11-14）：In 頁新增 M29，31 列 `nvsh_<買方>_<yy>`（Google、Amazon、Microsoft、Meta、Oracle、xAI、OpenAI 直接、新雲、中國、主權＋其他、Anthropic 直接 TPU；2025–2027；低／基準／高），Interface `SIF_nvsh_*`。研究報告 `docs/reports/20261104_逐買方NVIDIA份額_研究.md`（公司未揭露金額拆分；Google、AWS、中國為 Derived，其餘 Assumed／Interested-party）。全市場檢查：由上而下 NVIDIA 約 80%（2026），各買方以資本支出加權約 67%，差 13pt 主因為資本支出 ≠ 加速器支出。
- R4 r5（讀 Wattonomics v0.2.2）：與 r4 逐值相同。
- `tools/r4_bridge.py`：以 Wattonomics 具名範圍 `WIF_MktGW_Add_*`、`WIF_MktGW_YE_*` 交叉核對讀到的列（列序改變時報錯）。
- 輸入值不變；下游只有版本文字（cascade 第二輪全部 SAME）。

## v0.13（2026-11-04；`model/20261104_Siliconomics_v0.13.xlsx`；Tokenomics v21.7）

- **三層校準原則落地**（Andy 2026-11-04 08:26、12:45；決策 S37）：
  - 新增 **Cap 頁**（供給上限）：基準顆數的 CoWoS 與 HBM 需求 vs 產能 → 可支撐部署 GW；Checks K10（本庫基準 ≤ 上限）、K10w（Wattonomics 錨定 ≤ 上限）。結果：**2027 HBM 需求 ÷ 供給 121%，可支撐 34.4 GW，低於基準 41.6（K10-2027 WARN）**；其餘年度未觸頂；錨點全部在上限內。報告 `docs/reports/20261104_v0.13_供給上限與錨定案例.md`。
  - Recon **F 節季度實績驗證**：NVIDIA 資料中心營收（隱含 ÷ 公司）2026 1.017、2027 1.044。
  - Interface 新增 `SIF_WTAnchor_Base`（R4 比值）、`SIF_CapGW_Base`、`SIF_CapHead_Base`、`SIF_CoWoSUtil_Base`、`SIF_HBMUtil_Base`；`reader.py` 新增 `WTA`、`wta_rows()`。
- **R4 r4**（讀 Wattonomics v0.2.1）：0.946／0.735／0.640／0.539（一致／時差／待診斷／待診斷）。
- **個股敏感度「Wattonomics 錨定」案例**（60 檔；世界先進不適用）：以情境 B 為底，產業量 CY2027 起 × wta，FY2026 維持錨點；Sens 頁最下方新增小表與案例區塊，In 頁加 wta_27／28／29。既有數值不變（In、Sens 以外 0 差異）。
- `tools/in_layout/check_in_layout.py`：「Z. Wattonomics 錨定係數」節不屬情境驅動對照展示區。
- 輸入值不變。

## v0.12（2026-11-03；`model/20261103_Siliconomics_v0.12.xlsx`；Tokenomics v21.7）

- **只留基準**（工作方式 v1.7「三層只給基準、情境在個股」；決策 S33 基準＝原情境 B〔Andy 11-03 開場：未另行指示就照預設〕、S35）：Interface 新增 D 節 57 組 `*_Base` 名稱（等同原 B 列；例 `SIF_DeployGW_Base`、`SIF_UnitsPlat_Total_Base`、`SIF_um28_Base`、`SIF_cpo_Base_27`、`SIF_cqBase_8`、`SIF_ASICShareSum_trnBase`）；A／B／C 名稱項目欄標「停用過渡」，v0.13 刪除。Scen、Units 頁改標（原始顆數＝舊情境 A；本庫基準＝原始顆數 × 基準係數 0.8／0.85）；`um28_*`、`um29_*` 列名改「÷ 原始顆數」。
- **Recon 頁**：B 節改讀 R4 r3（基準），新增「本庫基準 ÷ Wattonomics 基準 − 1」列與 Interface `SIF_GWvsWT_Base`：+9.1%／+40.9%／+59.7%／+87.2%（2026–29）。狀態：2026 一致、2027 時差、2028–29 待診斷（比值 0.916／0.710／0.626／0.534；r2 情境 A 為 0.553／0.442）。
- **`tools/r4_bridge.py` r3**：平台別資本支出乘基準係數、部署 GW 讀 `SIF_DeployGW_Base`；新增 NVIDIA 資料中心營收對帳點（出貨年 NVIDIA GW × chip_bridge 每 GW 內容 vs 公司營收）：2026 1.017、2027 1.044（回應算力收支 11-02 的 2026 年 1.16：該值來自上線年往返換算）。報告 `docs/reports/r4_bridge/20261103_R4對帳橋接_Siliconomics側_r3.{md,xlsx}`。
- **個股自訂情境**（決策 S36）：35 個公司模型與 tsmc 的情境係數列改標「個股自訂」，說明欄由 `reader.scen_note()` 產生（相對本庫基準倍數；產業 GW 相對 Wattonomics 基準：A +100%／+120%、B +60%／+87%、C +20%／+32%，2028／2029）。預設沿用 v0.11 值，**數值不變**；其餘「情境 A 基準」等文字改「原始顆數」。
- 下游：cascade 第一輪 55 個 CHANGED（全部為文字：版本、標籤、說明；無數值差異），第二輪 64 個全部 SAME；HTML 展示層全部重建並通過檢核。
- 未做（v0.13）：記憶體價格路徑（cq／nq）、HBM 單價、世芯份額、液冷滲透率、CPO、設備係數等情境列移到個股或市場層自訂；刪除 A／B／C 名稱；Wattonomics 改讀 `*_Base`（已通知算力收支）。

## v0.11（2026-11-02；`model/20261102_Siliconomics_v0.11.xlsx`；Tokenomics v21.7）

- **新增 Recon 頁（由下往上對帳）**：依工作方式 v1.7 第 8 節（Andy 11-02：三層只給基準、對帳方向 覆蓋個股 → Siliconomics → Wattonomics → Tokenomics）。個股值與 R4 值來自版本固定快照 `data/recon_snapshot_v0.11.json`（`tools/recon.py` 產生；`data/CURRENT_RECON` 指向），不即時連結個股成品，避免反向依賴與 CI 漂移。狀態以公式判定：|比值 − 1| ≤ `rc_tol`＝一致；超出且有已寄出的 `交流/` 證據送件（`builder/seed/recon_evidence_v0.11.json`）＝時差；其餘＝待診斷。
  - A 節 SoC 測試機：Advantest＋泰瑞達基準情境份額 2025 90.9%、2026–29 98.1%，對預期 96%（1 − 其他廠 4%）全部「一致」；Checks K6-soc。
  - B 節 R4 部署 GW（讀 Wattonomics v0.1.8 基準，毛額）：比值 0.916／0.710／0.553／0.442（2026–29）；2026「一致」、2027「時差」（11-02 已送 NVIDIA 指引證據）、2028、2029「待診斷」；Checks K9（待診斷格數，WARN）。
  - C 節列出已由本庫份額表約束的市場（結構一致）；D 節列出未建對帳的市場（GPU／ASIC 晶片金額、CoWoS、設備、PCB／CCL、光學總量，v0.12 起）。
- 新增 In 頁 M27：`rc_tol` 0.15（0.10–0.20，Assumed，判斷類預設）、`rc_soc_oth` 0.04（0.02–0.07，Analogy）。決策 S32、S33（個股基準情境取 B，待 Andy）。
- 標記與區間更正（決策 S34）：`hbm_hbp_27` 高值 0.05 → 0（2027 不分歧，與 Equip 頁 D 節、韓美、Interface 一致）；`u_rb_28`、`u_bw_28` 原文為 Assumed 卻標 Interested-party → 改標 Assumed，區間 5.0–9.0、0–1.0（基準不變）；`nd_sh_essd25` 改標 Derived；混合鍵合列名「研削工序」改「HBM 工序比例」；`u_rb_27` 第二來源欄補 NVIDIA 營收指引交叉檢驗。
- `tools/r4_bridge.py`：Wattonomics 停用情境後改讀基準列（找不到 A／D0 時）；每次升版都跑，Wattonomics v0.1.8 主路徑改為毛額時 ④ 汰換不扣（旗標 `wt_gross`）；本版報告 `docs/reports/r4_bridge/20261102_R4對帳橋接_Siliconomics側_r2.{md,xlsx}`，與算力收支 R4 歸因 r2 逐值相同。
- 下游（cascade 兩輪，第二輪全部 SAME）：**BESI** 情境 A 2027 混合鍵合由讀高值 5% 改為 0，FY2027 EPS（A）6.96 → 7.10，加權目標價 €157.28 → €157.42（+0.09%），上檔空間 +8.4% → +8.5%，中立不變；**Disco** 只有 Sens 頁（混合鍵合高值案例）；**韓美** In 頁一格（高值）；SK 海力士、美光、三星、NAND 市場層與 UMC、世界、日月光、旺矽、台燿、Arm 只有上游版本或標記文字。其餘目標逐格相同。

## v0.10（2026-10-31；`model/20261031_Siliconomics_v0.10.xlsx`；Tokenomics v21.7）

- 輸入值變更（次版）：NAND 季價 `nq_6`（2026Q3）0.15 → 0.25（Verified：Kioxia、SK 海力士實現平均售價季增 mid-20%）；`nq_7`（2026Q4）0.175 → `=(1+0.15)*(1+0.175)/(1+0.25)-1`＝0.081（Derived；低 0.032／高 0.175）。2026Q4 起價格指數水準不變。
- 原因：SanDisk、Kioxia 2026Q4 模型營收比公司指引高 9.6%／21.5%；拆解為市場層把合約價當實現價（見 `docs/reports/20261031_NAND第四季_差距拆解.md`）。
- 連動：markets/nand v0.3、SanDisk v0.4、Kioxia v0.4；SK 海力士、三星、美光的 NAND 年均價（CY2026 年均價上升）隨之改變，目標價 −0.7%～−0.9%，投資建議不變。其他下游兩輪 SAME。決策 S30。

## v0.9（2026-10-31；`model/20261031_Siliconomics_v0.9.xlsx`；Tokenomics v21.7）

- 輸入值變更（次版）：`ts_mk_soc26`（CY2026 SoC 測試機市場）11.0 → 11.5（區間 11.0–12.0，US$B）。Advantest 2026-10-28 稱約 115 億美元（105–115 上緣）；Advantest 與泰瑞達 CY2026 SoC 營收（實績＋指引推估）合計約 113 億美元，已高於舊值。
- 連動：每加權顆測試支出 450 → 484 美元；`SIF_TestSoCMkt_*` CY2027 16.23（原 15.41），CY2028 A／B／C 18.24／15.52／12.81，CY2029 20.31／17.99／14.12。讀者：Advantest（v0.6）、泰瑞達（v0.3）、markets/equipment。
- 未改：非 AI SoC 市場 `ts_nonai26` 4.3（泰瑞達 3Q26 稱車用工業 +90%、行動 +60%，可能偏低；若為 5.0，CY2027 市場 16.2 → 15.8），列下一步查核。
- 報告：`docs/reports/20261031_SoC測試機_兩家合併對帳.md`。決策 S29。

## v0.8.4（2026-10-29；`model/20261029_Siliconomics_v0.8.4.xlsx`；Tokenomics v21.7）

- 來源複核（API，2026-10-28；結果在分支 `source-audit-results`）後的修補版：只改標記、來源文字與 2029 價格路徑的區間，輸入基準值全部不變。
  - DRAM／NAND 情境季價路徑 `cqA_*`、`cqB_*`、`nqA_*`、`nqB_*`：2027–28（第 8–15 列）由 Interested-party 改標 Derived（標示來源只有年度或半年數字，逐季為本庫拆分）；2029（第 16–19 列）改標 Assumed，區間取情境 A、B 同季值（標示來源只涵蓋到 2028）。
  - `ts_mk_soc26`、`ts_mk_mem26`：來源改為 Advantest 2026-07-29 決算說明原文（SoC 105–115 億、記憶體 25–30 億美元）；原寫「4/27 簡報」為誤植（4/27 為 87–95 億）。
- 下游：見報告 `docs/reports/20261029_v0.8.4.md`。決策 S28。

## v0.8.3（2026-10-28；`model/20261028_Siliconomics_v0.8.3.xlsx`；Tokenomics v21.7）

- 衝突登錄清理（修補版：只改文字）：CF02（Ibiden 份額相對值 `abf_ibrel_28` 的來源改以欣興 2026 資本支出 537 億為準）、CF05（AMD README 已一致）、CF06（韓美 Sens 情境 C 標籤更正為「混合鍵合低值（導入慢）」）、CF07（besi/README.md 跨模型注意更正）、CF08（lrcx/README.md 25 倍）、CF16（Drive 候選母體 v0.2.4）結案；CF01 移到南電分部重建；CF03 狀態補「待 Andy」字首。Checks G5：10 → 4（CF03、CF11、CF13、CF23）。
- 下游：Ibiden In 頁 1 格來源文字、韓美 Sens 1 格標籤；數值不變。決策 S27。

## v0.8.2（2026-10-28；`model/20261028_Siliconomics_v0.8.2.xlsx`；Tokenomics v21.7）

- 新增 In 頁 M26（3 列）：共用匯率 新台幣 32.0（30.5–33.0）、人民幣 6.72（6.50–7.00；2026-10-16 中間價）、韓元 1,420（1,350–1,500）。台廠 15 個、陸廠 3 個、韓廠 3 個模型改讀；日圓不共用（Ibiden ¥152、Advantest ¥150、Kioxia ¥160 為公司財測前提）。
- 數值變動：欣興、南電、景碩 32.5 → 32.0；中際旭創、新易盛 7.1 → 6.72（Coherent、Fabrinet 讀中際旭創 Units，22 格）。FY2027 EPS 變動 0.00% 至 −0.07%，目標價 0 至 −0.36%，投資建議全部不變（份額以 FY2026 同一匯率校準，匯率變動大多被吸收）。其餘 51 個目標逐格相同。決策 S26。

## v0.8.1（2026-10-28；`model/20261028_Siliconomics_v0.8.1.xlsx`；Tokenomics v21.7）

- 新增 In 頁 M25（17 列）：廣達、鴻海、緯創本體、緯穎的 FY2026 組裝有效份額快照（Derived，公司模型校準值）、2027–29 份額相對值（原在四家模型）、緯穎 ASIC 比重；四家改讀本庫數值。
- ShareTbl 新增 AI 伺服器組裝份額表（緯穎乘 ASIC 傾斜係數 1.45／1.48／1.49）、Checks K6-asm（合計 43.0／44.6／43.6／41.4%）、Interface `SIF_AsmShareSum`、`SIF_AsmWyTilt`、`SIF_AsmShare_wy`。
- CF23 登錄：緯創模型自算的緯穎 AI 份額路徑（3.42／3.42／3.42／3.25%）與緯穎模型（3.45／4.99／5.09／4.89%，含 ASIC 傾斜）不同；試算改讀後緯創 FY2027 EPS（A）27.18 → 30.28（+11.4%）、目標價 349.4 → 395.6；預設暫不改，待 Andy。
- 下游數值不變。決策 S25。

## v0.8（2026-10-28；`model/20261028_Siliconomics_v0.8.xlsx`；Tokenomics v21.7）

- 新增 In 頁 M24（15 列）：SK 海力士、三星、美光 HBM 營收份額 2025–2029（基準與低／高、標記、來源逐字搬入；原在各 company.py），三家改讀本庫。
- ShareTbl 新增 HBM 份額表、Checks K6-hbm（合計 2025 96%、2026–29 100%；餘額為長鑫等）、Interface `SIF_HBMShareSum`。
- 台積電模型的 HBM 2027 需求缺口 ÷ 供給由手打 21% 改讀本庫推導值（21.0%；只在 CoWoS 對照列）。
- 下游數值不變。決策 S24。

## v0.7.1（2026-10-28；`model/20261028_Siliconomics_v0.7.1.xlsx`；Tokenomics v21.7）

- 新增 In 頁 M22（18 列）：HBM 核心晶粒容量、堆疊層數、混合鍵合占 HBM 工序比例（2025–2030），原在 disco；disco 改讀，hanmi、besi 由讀 Disco 成品改讀本庫。Equip 頁 D 節：混合鍵合比例三情境（2028 A 15%／B 5%／C 0%；2029 30／15／5%）。
- 新增 In 頁 M23（61 列）：markets/nand 的行業輸入（每顆加速器企業級 SSD 容量、2025 位元基期與應用結構、各應用與供給位元年增、各廠季營收、份額年變動），nand 改讀；前五大占比公式、公司指引、合資群份額情境留在 nand。
- 下游數值不變（besi、hanmi 只有 In 頁來源文字）。決策 S23。

## v0.7（2026-10-28；`model/20261028_Siliconomics_v0.7.xlsx`；Tokenomics v21.7）

- 新增 In 頁 M20（55 列）：markets/equipment 的行業輸入（SEMI 2026-07 年中預測、2029 外推 g29、其他晶圓廠設備成長與 2025 資本支出、設備情境係數 B／C、AI 相關分類、券商 WFE 對照），equipment 改讀；三大記憶體廠與台積電資本支出、設備占比仍在 equipment（公司變數）。
- 新增 In 頁 M21（18 列）：Advantest 的測試機市場錨點、非 AI SoC、每顆測試時間係數、HBM 權重，泰瑞達的測試 ÷ WFE 長期區間 7–9%。
- 新增 Equip 頁：SEMI 分項路徑（WFE 116.9／143.9／175.3／200.0／216.0 US$B，2025–29）、SoC 測試機市場三情境（A 11.0／15.4／17.3／19.2）、測試 ÷ WFE；Checks K8（CF15，SoC ÷ SEMI WFE 7.6／8.8／8.6%，全部在 7–9% 內）。
- 下游：58 個目標逐格相同。決策 S22。

## v0.6.1（2026-10-27；`model/20261027_Siliconomics_v0.6.1.xlsx`；Tokenomics v21.7）

- 新增 In 頁 M19（液冷，20 列）：機架內液冷每架內容 `lc_cool_gb`／`lc_cool_vr`（50／56 千美元，原在 markets/chip_bridge）、分歧管 `lc_mani`、內容調整 `lc_adj_*`、滲透率 `lc_pen_*`（行業情境：FY2027 起 A 高、B 基準、C 低）、TrendForce 對照（原在 avc）、奇鋐與雙鴻 FY2026 份額快照、奇鋐份額相對值（雙鴻原即讀同一組）、健策冷板元件份額（原在 jentech）。
- Rack 頁新增 D 節 AI 液冷市場：情境 A 4.5／11.6／17.6／20.1 US$B（2026–29）；ShareTbl 新增 AI 液冷份額表（CF04，奇鋐＋雙鴻＋健策合計 54.5–65.8%）與 Checks K6-lc。
- CF12：市場定義為機架內液冷＋分歧管；台達電液冷多在機架外（CDU、Liquid-to-Air），不屬本市場，原「台達電＋奇鋐＝115%」是兩個市場相加。機架外冷卻市場未建。
- 下游：橋接表、avc、auras、jentech 改讀；數值不變（avc、auras 只有 In 頁來源文字）。決策 S21。

## v0.6（2026-10-27；`model/20261027_Siliconomics_v0.6.xlsx`；Tokenomics v21.7）

- 新增 In 頁 M17（AI 電源，16 列）：每 kW 電源內容 `pwr_pkw_26–29`、台達電份額相對值 `pwr_drel_*`、光寶份額相對值 `pwr_lrel_*`（原在 delta、liteon，數值與文字逐字搬入）、FY2026 有效份額快照 `pwr_dsh_26`／`pwr_lsh_26`（Derived，公司模型校準值）、凱基兩個錨點 `pwr_kgi_0701`（70%）／`pwr_kgi_0926`（62.5%）、UBS 市場 `pwr_ubs_26／27`。
- 新增 In 頁 M18（AI 伺服器組裝，3 列）：ODM 每櫃價格 `asm_odm_gb`／`asm_odm_vr`（原在 quanta）、緯穎 2026 AI 占比 `asm_wyai_26`＝48%（CF09：以緯穎模型為準，緯創原 55%）。
- 新增 Rack 頁：AI 伺服器組裝市場（ODM 價）、商用 GPU 系統市場（CF10：AMD 2026 MI355 計 GB300 世代、2027 起 MI450 計 Rubin 世代；2026＝595.5）、AI 電源市場；ShareTbl 新增 AI 電源份額表（三情境），Checks K6-pwr、K7。
- 下游：delta、liteon、quanta、wiwynn 改讀（數值不變）；**戴爾**改用同一世代分組（2026 系統市場 638.2 → 595.5，FY2028 EPS +3.5%、目標價 +3.1%，中立不變）；**緯創**改讀緯穎 48%（FY2027 EPS −1.5%、目標價 −1.6%，強力買進不變）。
- CF11 待 Andy：只建份額表，凱基兩個值只作對照；試算見報告。液冷（CF04、CF12）移到 v0.6.1。決策 S20。

## 工程（2026-10-27）

- 新增 `tools/cascade.sh`：升版後本機下游批次重建（清單直接取自 CI 工作流程，與 CI 同順序、同比對；跑兩輪，第二輪須全部相同）。

## v0.5.2（2026-10-27；`model/20261027_Siliconomics_v0.5.2.xlsx`；Tokenomics v21.7）

- 新增 In 頁 M16（24 列）：勝宏模型的每 GW PCB、台光電模型的每 GW CCL（2026 元件金額與 2027–29 指數），兩家改讀本庫（數值與文字不變）。
- 新增 PGW 頁與 Interface `SIF_PCBperGW`、`SIF_CCLperGW`、`SIF_PCBMkt_*`、`SIF_CCLMkt_*`：每 GW PCB 450／682／792／877、CCL 350／528／609／663 US$M（2026–29）。
- 報告：公司 ABF 滿足率改讀本庫缺口的試算（模型未改）；建議改研究缺口對價格溢價的對應（決策 S19）。

## v0.5.1（2026-10-27；`model/20261027_Siliconomics_v0.5.1.xlsx`；Tokenomics v21.7）

- 新增 In 頁 M15（ABF 供需表輸入，28 列）：2025 ABF 產值（TPCA 77.2 億美元）、2026 年增、AI 占比、加速器占 AI 比、其他 AI 彈性、非 AI 年增、四家 2026 ABF 營收（供給權重）、Ibiden 與其他載板廠產能指數（新增，Assumed）；欣興、南電、景碩產能指數與味之素膜上限由公司模型移入（數值不變）。
- 新增 ABF 頁：需求指數、載板產能指數、膜上限、有效供給、缺口三情境；Interface `SIF_ABFDemIdx_*`、`SIF_ABFSubCapIdx`、`SIF_ABFFilmIdx`、`SIF_ABFSupIdx`、`SIF_ABFGap_*`。情境 A 缺口 2027／28／29＝15.6／11.0／12.3%。
- 下游：unimicron、nanya_pcb、kinsus 的 `cap_*`、ajinomoto 的 `cap_*` 改讀本庫；各公司滿足率未改讀本庫缺口（待 Andy）。
- 決策 S17（Andy 18:07：CF20 景碩份額待本表後確認，目前照預設）、S18。

## v0.5（2026-10-27；`model/20261027_Siliconomics_v0.5.xlsx`；Tokenomics v21.7）

- 新增 In 頁 M13（加速器 ABF 載板，37 列）：Ibiden 2026 份額 70% 與 2027–29 相對值、非 NVIDIA 權重、`v_ot`、每顆載板金額年增 `g_asp`、價格溢價指數 `pp_*`、欣興／南電／景碩的 NVIDIA 與非 NVIDIA 份額（原在 ibiden、unimicron、nanya_pcb、kinsus），以及單一外部 ABF 缺口路徑 `abf_gap_26–28`＝9／24／37%（CF03 部分併入）。
- 新增 In 頁 M14（CoWoS，20 列）：每片晶圓封裝數（10 個平台）、其他 CoWoS 用途、台積電年底月產能 2025–29、委外補充、MS 需求（原在 tsmc）；Assumed 列補區間、兩列補來源欄（基準不變）。
- ShareTbl 新增兩張 ABF 載板份額表（NVIDIA 側合計 99.0–99.6%、非 NVIDIA 側 88.0–89.5%）、Checks K6-abf、Interface `SIF_ABFShareSum_nv`／`_ot`（CF20 表已建，景碩份額待 Andy）。
- 留在公司模型：每顆載板金額 `v26`（以 Ibiden 營收 ÷ 70% 校準）、每顆台積電內容 `ct_*`（公司營收變數）。
- 下游：ibiden、unimicron、nanya_pcb、kinsus、tsmc 改讀本庫。

## v0.4.1（2026-10-27；`model/20261027_Siliconomics_v0.4.1.xlsx`；Tokenomics v21.7）

- 新增 In 頁 M12：CPO 光學金額份額 `cposh_*`（Lumentum、Coherent、中際旭創、新易盛、Fabrinet，2027–2029，15 列），原在五個公司模型，數值與文字逐字搬入（CF21）。
- ShareTbl 新增 CPO 份額表：四家模組與雷射廠合計 50／65／75%，餘額 50／35／25%（NVIDIA、Broadcom 交換器端自製光引擎與其他，未建模）；Fabrinet 為代工、與模組廠重疊，列示不加總（含 Fabrinet 的舊口徑為 55／75／90%）。Checks K6-cpo（合計 ≤ 100%，全部 OK）、Interface `SIF_CPOShareSum`。
- 下游：lite、cohr、innolight、eoptolink、fabrinet 的 `cpo_27`–`cpo_29` 改以 `SIL.row9("cposh_…")` 讀本庫（鍵名與文字不變）。
- builder：seed 檔改依版本號排序（字串排序會把 v0.4.1 排在 v0.4 之前）。

## v0.4（2026-10-27；`model/20261027_Siliconomics_v0.4.xlsx`；Tokenomics v21.7）

- 新增 In 頁 M8（每顆加速器網通內容 54 列）、M9（每顆 IT capex 9 列；GB300、VR200、Rubin Ultra、Hopper 改為 TK_In 公式 IF_CapexIT ÷ IF_GPUsPerGW）、M10（scale-up 光學占比、銅占比、CPO 三情境 12 列）、M11（vr26：2026 Rubin 占比 22%），原在 markets/network 與 11 個公司模型；Assumed 列補區間（非 NVIDIA 每顆 IT capex、scale-up 光學、CPO 三情境 ×0.7–×1.4）。
- 新增 Deploy 頁與 Interface `SIF_AIITCapex_*`、`SIF_ITCapexPerGW`、`SIF_DeployGW_*`（2026–2029；情境 A 23.4／41.6／52.8／60.1 GW）。
- 橋接表 v0.2.2：改讀本庫 TK v21.7 快照與 R1 三個契約名稱（CF17）；只有 VR200 低成本欄變動（機架價 7.8 → 7.5）。Checks K3 歸零。
- 下游：network 改讀 M8–M10；台達電、廣達、奇鋐、台光電、勝宏、光寶、金像電的 vr26 改讀 M11；Amkor 部署 GW 改用 vr26 混合口徑（CF22，2026 24.25 → 23.40，只影響 Recon）。重建 57 個目標：目標價與上檔空間全部不變（差異 < 1e-12）；其餘差異為橋接表檔名文字、In 頁區間、Dell 低端敏感度（隨 VR200 低成本機架價）；第二輪全部相同。

## v0.3（2026-10-27；`model/20261027_Siliconomics_v0.3.xlsx`；Tokenomics v21.7）

- 新增 In 頁 M6（加速器每顆售價 `pa_*`、ASIC 核心價 `core_*`、HBM 代購比例 `f_hbm`；原在 nvidia、avgo）與 M7（ASIC 設計服務份額：Broadcom `sh_*`、Marvell `msh_*`、聯發科 `mtk_*`、世芯 `al*`；原在 avgo、mrvl、mediatek、alchip），數值與文字逐字搬入；In 頁新增 J 欄第二來源。
- 新增 ShareTbl 頁（TPU、其他 ASIC、Trainium 三情境的設計夥伴份額合計與餘額）、Checks K6（合計 ≤ 100%，全部 OK）、Interface `SIF_ASICShareSum_*`。
- 新增 `siliconomics/reader.py`；nvidia、avgo、mrvl、mediatek、alchip 改以 `inp(*SIL.row9(鍵))` 讀本庫。下游 57 個目標重建：五家模型逐格不變；tsmc、network、nand 只有一格來源文字（拿掉版本號，之後升版不再改動下游文字）；第二輪全部相同。
- 衝突：CF14 釐清為公司營收口徑差（非衝突）；CF13 兩組價格已同處，口徑待 Andy；CF19 份額表與加總檢查已建，數值維持模型路徑（Andy 15:09 同意）。

## v0.2（2026-10-27；`model/20261027_Siliconomics_v0.2.xlsx`；Tokenomics v21.7）

- In 頁輸入改動（工作單 `docs/workorders/20261027_v0.2_inputs.json`，以 `tools/apply_inputs.py` 寫入 Excel 後重建）：Assumed 列全部補區間（C 情境 DRAM／NAND 季價 2027 ±5pt、2028–29 ±10pt；Rubin Ultra、AMD、Trainium、其他 ASIC 顆數與每顆 HBM）；Rubin Ultra 每顆 HBM 標記改 Interested-party、情境 A 係數改 Derived；2029 HBM 單價兩列補來源。**基準值全部不變。**
- 新增 M5 節 2029 平台別顆數成長（`g29_*`，Blackwell −50%、Hopper 0、其餘 +15%，區間 ±10pt）與 Units 頁；Interface 新增 `SIF_g29_*`、`SIF_UnitsPlat_*`、`SIF_UnitsGrowth29`（13.77%）。
- Checks：G2（缺來源）、G3（Assumed 缺區間）升為 ERROR（皆 0）；新增 K4（平台別合計＝HBM 頁合計）、K5（成長率在區間內）。
- 下游：`market.py` 新增 `G29`、`G_U29`，replay 只放 M1–M4；tsmc、network 的 2029 成長改讀本庫（值不變）；NAND 的 2029 顆數年增 10% → 13.77%（CF18 已併入）。重建 57 個目標：數值變動只有 NAND、SanDisk、Kioxia；skhynix、micron、samsung、tsmc、network 只有 In 頁區間與來源文字；第二輪重建全部相同。

## v0.1（2026-10-27；`model/20261027_Siliconomics_v0.1.xlsx`；Tokenomics v21.7）

- 新建 AI 半導體產業基準庫（比照 Tokenomics）：README、Scen、In、Path、Mkt、TK_In、Conflicts、Interface、Checks、Decisions、DB_Evidence。
- `markets/memory/market.py` 的全部輸入（DRAM／NAND 季價、加速器顆數與每顆 HBM、情境係數、HBM 市場、對帳參考，共 161 列）逐字搬入 In 頁；`market.py` 改為只讀 `data/` 快照，保留原 Python 名稱。驗證：新舊 `market.py` 的全部公開常數與輸入列 `repr` 相同；`ai-semis-build` 58 個目標重建後與 dist 逐格相同（錯誤格 0）。
- Tokenomics v21.7 鏡像 22 個名稱（`TK_In`、`SIF_TK_*`）；Checks K3：下游（chip_bridge v20.1 快照）與現行 TK 不同 2 格（VR200 低成本欄 `IF_CapexIT`、`IF_CapexTotal`）。
- Conflicts 頁登錄 22 項跨模型衝突（含算力收支 2026-10-25 所列 15 項；盤點見 docs/plan/20261027_行業變數盤點.md）。
- CI：`ai-semis-build` 新增本庫重建比對與快照一致檢查。

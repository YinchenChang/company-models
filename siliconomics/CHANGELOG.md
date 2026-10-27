# Siliconomics 變更紀錄

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

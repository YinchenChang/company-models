# Tokenomics 連接審視（2026-10-08）

對照對象：Tokenomics `model/CURRENT`＝`20261008_Tokenomics_v5.27.xlsx`（ca78a8f）。契約：`YinchenChang/Tokenomics` `docs/plan/Tokenomics_downstream_contract.md`。只回報，不改模型。

| 公司 | 取數方式 | 釘住版本 | 使用名稱數 | 不存在於現行 | 違規引用 | 值漂移（名稱數） | 四層瀑布／用途欄 |
|---|---|---|---|---|---|---|---|
| anthropic | 無 | —（落後） | 52 | 3 | 6 | 0 | 否／否 |
| coreweave | 官方快照（tokenomics_snapshot_v5.26.json） | 20261007_Tokenomics_v5.26.xlsx（落後） | 25 | 0 | 3 | 0 | 否／否 |
| minimax | 無 | —（落後） | 6 | 0 | 1 | 0 | 否／否 |
| nebius | 手工 JSON（permw_tokenomics_20261007.json） | model/20261007_Tokenomics_v5.24.xlsx（落後） | 10 | 0 | 3 | 0 | 否／否 |
| openai | 無 | —（落後） | 49 | 4 | 4 | 0 | 否／否 |
| oracle | 手工 JSON（permw_tokenomics_20261007.json） | model/20261007_Tokenomics_v5.24.xlsx（落後） | 19 | 0 | 3 | 0 | 否／否 |
| whitefiber | 手工 JSON（permw_tokenomics_20261008.json） | model/20261007_Tokenomics_v5.24.xlsx（落後） | 10 | 0 | 2 | 0 | 否／否 |
| zhipu | 無 | —（落後） | 52 | 5 | 2 | 0 | 否／否 |

## anthropic
- 取數方式：無；釘住 `—`（—）（落後）。
- 使用名稱（52）：IF_Alloc, IF_AllocDemand, IF_AllocImpliedNk, IF_AllocQ1, IF_AllocQ1_R2, IF_AllocQ2, IF_AllocRDGW, IF_AllocServeGW, IF_CapexTotal, IF_FrontRef_Astra, IF_FrontRef_Luna, IF_FrontRef_Sol, IF_FullCostDefault_Astra, IF_FullCostDefault_Luna, IF_FullCostDefault_Sol, IF_FullCost_Astra, IF_FullCost_Astra_Prod, IF_FullCost_Luna, IF_FullCost_Luna_Prod, IF_FullCost_Sol, IF_FullCost_Sol_Prod, IF_HoldEcon, IF_Price, IF_ProgGWyr_Astra, IF_ProgGWyr_Luna, IF_ProgGWyr_Sol, IF_RDMult, IF_RevGW_Astra, IF_RevGW_Astra_Life, IF_RevGW_Astra_Prod, IF_RevGW_Luna, IF_RevGW_Luna_Life, IF_RevGW_Luna_Prod, IF_RevGW_Sol, IF_RevGW_Sol_Life, IF_RevGW_Sol_Prod, IF_TaskLen, IF_TaskTok, IF_TaskTokCached, IF_TaskTokDec, IF_TaskTokFresh, IF_TokGW_Astra, IF_TokGW_Luna, IF_TokGW_Sol, IF_TrainCost_Astra, IF_TrainCost_Luna, IF_TrainCost_Sol, IF_TrainGenDefault, IF_Util, L1_Ans3, L1_Ans3_Hi, L1_Ans3_Lo
- **不存在於現行 Tokenomics**：IF_Alloc, IF_Price, IF_TaskTok
- **違規引用（契約第 2 條第 1 項）**：`IF_HdrTask`（builder/build.py）；`IF_HdrTask`（builder/tk_link.py）；`IF_HdrTask`（tests/parity/scenarios.yaml）；`IF_HdrTask`（tools/make_report_tables.py）；`Inputs!`（builder/a3.py）；`Inputs!`（builder/build.py）
- 四層瀑布（IFW_）：否；用途欄（IFC_）：否。

## coreweave
- 取數方式：官方快照（tokenomics_snapshot_v5.26.json）；釘住 `20261007_Tokenomics_v5.26.xlsx`（4074684）（落後）。
- 使用名稱（25）：IF_AvgDraw, IF_CapexFacility, IF_CapexIT, IF_CapexTotal, IF_DeprFac, IF_DeprIT, IF_DeprLifeIT, IF_FacilityGW, IF_GPUhrEcon, IF_GPUsPerGW, IF_HoldAcct, IF_HoldEcon, IF_MaintFac, IF_MaintIT, IF_OpexGW, IF_PowerCost, IF_PowerPrice, IF_RacksPerGW, IF_StaffSW, IF_TaxIns, IF_Util, L1_FacCapexMW, L1_GPUhr_GB200_vsCW, L1_GPUhr_GB300_vsBE, L1_RevGW_Fleet_VR200
- **違規引用（契約第 2 條第 1 項）**：`IF_Hdr`（build_xlsx.py）；`IF_Hdr`（company.json）；`IF_Hdr`（data/tokenomics_names.txt）
- 四層瀑布（IFW_）：否；用途欄（IFC_）：否。

## minimax
- 取數方式：無；釘住 `—`（—）（落後）。
- 使用名稱（6）：IF_FullCost_Luna, IF_FullCost_Sol, IF_HoldEcon, IF_TokGW_Luna, IF_TokGW_Sol, IF_Util
- **違規引用（契約第 2 條第 1 項）**：`Inputs!`（build_xlsx.py）
- 四層瀑布（IFW_）：否；用途欄（IFC_）：否。

## nebius
- 取數方式：手工 JSON（permw_tokenomics_20261007.json）；釘住 `model/20261007_Tokenomics_v5.24.xlsx`（098873a）（落後）。
- 使用名稱（10）：IF_CapexFacility, IF_CapexIT, IF_CapexTotal, IF_FacilityGW, IF_GPUhrEcon, IF_GPUsPerGW, IF_HoldAcct, IF_HoldEcon, IF_PowerCost, IF_Util
- **違規引用（契約第 2 條第 1 項）**：`DC_Cost!`（data/permw_tokenomics_20261007.json）；`Inputs!`（data/nebius_facts_20261007.json）；`Inputs!`（data/permw_tokenomics_20261007.json）
- 四層瀑布（IFW_）：否；用途欄（IFC_）：否。

## openai
- 取數方式：無；釘住 `—`（—）（落後）。
- 使用名稱（49）：IF_Alloc, IF_AllocDemand, IF_AllocImpliedNk, IF_AllocQ1, IF_AllocQ1_R2, IF_AllocQ2, IF_AllocRDGW, IF_AllocServeGW, IF_CapexTotal, IF_FrontRef_Astra, IF_FrontRef_Luna, IF_FrontRef_Sol, IF_FullCost, IF_FullCostDefault_Astra, IF_FullCostDefault_Luna, IF_FullCostDefault_Sol, IF_FullCost_Astra, IF_FullCost_Astra_Prod, IF_FullCost_Luna, IF_FullCost_Luna_Prod, IF_FullCost_Sol, IF_FullCost_Sol_Prod, IF_HoldEcon, IF_Price, IF_ProgGWyr_Astra, IF_ProgGWyr_Luna, IF_ProgGWyr_Sol, IF_RDMult, IF_RevGW_Astra, IF_RevGW_Astra_Life, IF_RevGW_Astra_Prod, IF_RevGW_Luna, IF_RevGW_Luna_Life, IF_RevGW_Luna_Prod, IF_RevGW_Sol, IF_RevGW_Sol_Life, IF_RevGW_Sol_Prod, IF_TokGW, IF_TokGW_Astra, IF_TokGW_Luna, IF_TokGW_Sol, IF_TrainCost_Astra, IF_TrainCost_Luna, IF_TrainCost_Sol, IF_TrainGenDefault, IF_Util, L1_Ans3, L1_Ans3_Hi, L1_Ans3_Lo
- **不存在於現行 Tokenomics**：IF_Alloc, IF_FullCost, IF_Price, IF_TokGW
- **違規引用（契約第 2 條第 1 項）**：`Inputs!`（builder/build.py）；`Inputs!`（builder/p2.py）；`Inputs!`（builder/p3.py）；`Inputs!`（builder/p4.py）
- 四層瀑布（IFW_）：否；用途欄（IFC_）：否。

## oracle
- 取數方式：手工 JSON（permw_tokenomics_20261007.json）；釘住 `model/20261007_Tokenomics_v5.24.xlsx`（098873a）（落後）。
- 使用名稱（19）：IF_CapexFacility, IF_CapexIT, IF_CapexTotal, IF_FacilityGW, IF_GPUhrEcon, IF_GPUsPerGW, IF_HoldAcct, IF_HoldEcon, IF_PowerCost, IF_RevGW_Astra, IF_RevGW_Astra_Life, IF_RevGW_Astra_Prod, IF_RevGW_Luna, IF_RevGW_Luna_Life, IF_RevGW_Luna_Prod, IF_RevGW_Sol, IF_RevGW_Sol_Life, IF_RevGW_Sol_Prod, IF_Util
- **違規引用（契約第 2 條第 1 項）**：`Cap_In!`（data/oracle_facts_20261007.json）；`DC_Cost!`（data/permw_tokenomics_20261007.json）；`Inputs!`（data/permw_tokenomics_20261007.json）
- 四層瀑布（IFW_）：否；用途欄（IFC_）：否。

## whitefiber
- 取數方式：手工 JSON（permw_tokenomics_20261008.json）；釘住 `model/20261007_Tokenomics_v5.24.xlsx`（098873a）（落後）。
- 使用名稱（10）：IF_CapexFacility, IF_CapexIT, IF_CapexTotal, IF_FacilityGW, IF_GPUhrEcon, IF_GPUsPerGW, IF_HoldAcct, IF_HoldEcon, IF_PowerCost, IF_Util
- **違規引用（契約第 2 條第 1 項）**：`DC_Cost!`（data/permw_tokenomics_20261008.json）；`Inputs!`（data/permw_tokenomics_20261008.json）
- 四層瀑布（IFW_）：否；用途欄（IFC_）：否。

## zhipu
- 取數方式：無；釘住 `—`（—）（落後）。
- 使用名稱（52）：IF_Alloc, IF_AllocDemand, IF_AllocImpliedNk, IF_AllocQ1, IF_AllocQ1_R2, IF_AllocQ2, IF_AllocRDGW, IF_AllocServeGW, IF_CapexTotal, IF_Cost, IF_CostPre, IF_DeprLifeIT, IF_FrontRef_Astra, IF_FrontRef_Luna, IF_FrontRef_Sol, IF_FullCostDefault_Astra, IF_FullCostDefault_Luna, IF_FullCostDefault_Sol, IF_FullCost_Astra, IF_FullCost_Astra_Prod, IF_FullCost_Luna, IF_FullCost_Luna_Prod, IF_FullCost_Sol, IF_FullCost_Sol_Prod, IF_HoldEcon, IF_OpexGW, IF_Price, IF_ProgGWyr_Astra, IF_ProgGWyr_Luna, IF_ProgGWyr_Sol, IF_RDMult, IF_RevGW_Astra, IF_RevGW_Astra_Life, IF_RevGW_Astra_Prod, IF_RevGW_Luna, IF_RevGW_Luna_Life, IF_RevGW_Luna_Prod, IF_RevGW_Sol, IF_RevGW_Sol_Life, IF_RevGW_Sol_Prod, IF_TokGW, IF_TokGW_Astra, IF_TokGW_Luna, IF_TokGW_Sol, IF_TrainCost_Astra, IF_TrainCost_Luna, IF_TrainCost_Sol, IF_TrainGenDefault, IF_Util, L1_Ans3, L1_Ans3_Hi, L1_Ans3_Lo
- **不存在於現行 Tokenomics**：IF_Alloc, IF_Cost, IF_CostPre, IF_Price, IF_TokGW
- **違規引用（契約第 2 條第 1 項）**：`Inputs!`（builder/build.py）；`Inputs!`（builder/z2.py）
- 四層瀑布（IFW_）：否；用途欄（IFC_）：否。

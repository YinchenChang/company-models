# 智譜收支模型 交接檔（權威版；放 repo）

> 新對話串：先讀本檔、規格 `docs/plan/20261008_Zhipu_v0.1_規格.md` 與進度檔 `docs/reports/20261008_zhipu_進度.md`，再以 git 確認狀態。

## 1. 命題與範圍
- 命題（與 OpenAI v0.6 同一句）：智譜每 VR 等值 GW 的年營收能否覆蓋每 GW 年全成本；若不能，缺口要多少外部資金、由誰以什麼條件提供。FY2025–FY2030，曆年制，2025 為實際校準年，1H26 為第二個實際點。
- 來由：Andy 2026-10-08「請比照 OpenAI，做智譜（智譜華章／02513.HK）」；依 2026-10-07 授權，Claude 套用預設（規格 D1–D26）不先確認，v0.1 完成時彙總審查。

## 2. 位置
- repo `YinchenChang/company-models`，資料夾 `zhipu/`；建置分支 `claude/zhipu-v0.1`（單一 PR）。
- 成品：`dist/`；Drive `02-算力收支/Zhipu/`（chat 端放置）。

## 3. 模型狀態
- Z0 完成（2026-10-08）：骨架、規格 r0、共同規則、Z1–Z5 工作單。
- Tokenomics 快照：v5.26（master `3dd1216`），與 OpenAI v0.6 相同。

# markets/network — AI 網通市場模型（AI 半導體 Project 步驟 3）

**模型命題：** AI 網通的錢由「加速器顆數 × 每顆網通內容」決定；每顆內容隨 scale-up 域擴大與每 GPU 網卡頻寬倍增而上升，但 scale-up 的大頻寬大多由便宜的銅承載，所以金額主體仍在 scale-out（交換器、光模組、網卡）。

- 建置：`python3 markets/network/build_network.py raw.xlsx`，再以 LibreOffice headless 重算（錯誤 0）。
- 顆數取自 `markets/memory/market.py`（與 HBM 市場同一份加速器數量層）；NVIDIA 平台每顆內容以 `markets/chip_bridge` v0.2 為錨（GB300 約 $11.6K（v0.1 $13.0K；Tokenomics v19.0 機架基準價 $4.3M）、VR200 約 $24.5K；Bernstein GB200 約 $10.7K 為第二來源）。
- 頁面：README、In、Units、Mkt（依元件、依產品類別、網通 ÷ AI IT capex、NVIDIA／Broadcom 對帳）、Sens、Sources。

## v0.2.3（2026-10-26）：CPO 滲透改三情境

依 Andy 2026-10-25 13:13 規則（個股 A／B／C 一律讀同字母行業情境；技術替代屬行業變數，A＝替代快、C＝慢），CPO 占 scale-out 光學金額由單一路徑改為三情境（`In` 頁 F 節 `cpo_A_*`／`cpo_B_*`／`cpo_C_*`；`Mkt` 頁末「CPO 三情境」新增 `其中 CPO（情境 X）`、`其中可插拔（情境 X）`）。既有列位置不動；`Mkt` 主表的「其中 CPO」「其中可插拔」＝情境 A（`cpo_*` 改連到 `cpo_A_*`）。

| CPO ÷ scale-out 光學 | 2026 | 2027 | 2028 | 2029 | 標記與依據 |
|---|---|---|---|---|---|
| A 替代快 | 1% | 5% | 10% | 16% | Assumed（2028 區間 8–14%）；錨：Citi 2026-03 scale-out CPO 交換器 2027 4 萬台、2028 約 10 萬台 [Interested-party]，以每埠光學 $300–400 換算約 2027 $2–3B、2028 $6–8B |
| B 平平 | 1% | 3% | 5% | 8% | ＝v0.2.2 單一路徑（2028 區間 4–7%）；LightCounting 2030 年約 10% [Interested-party]、Morgan Stanley 2026-06「快速成長在 2028 年後」[Interested-party]、TrendForce 放量 2027–28 |
| C 替代慢 | 1% | 2% | 3% | 5% | Assumed（2028 區間 2–4%）；Cignal AI 2026-05「CPO 對可插拔初期影響小」；TSMC COUPE／SoIC 良率（Morgan Stanley 轉述）與維修性拖延 |

| 十億美元（顆數基準，下游再乘同字母情境顆數係數） | 2026 | 2027 | 2028 | 2029 |
|---|---|---|---|---|
| CPO A／B／C | 0.3／0.3／0.3 | 2.8／1.7／1.1 | 6.4／3.2／1.9 | 11.7／5.9／3.7 |
| 可插拔 A／B／C | 29.8／29.8／29.8 | 53.6／54.7／55.3 | 58.0／61.2／62.5 | 72.9／78.8／81.0 |

**為何 A＝CPO 滲透快：** 行業情境字母代表 AI 部署「順利／平平／不如預期」。部署順利時，客戶有預算與急迫性採用新世代（102.4T 交換器、1.6T／3.2T 每埠）與省電方案，NVIDIA Spectrum-X／Quantum-X Photonics 與 Broadcom CPO 交換器導入快；部署不如預期時，客戶延用已驗證、可維修的可插拔光模組，CPO 延後。方向：CPO 快對 **可插拔光模組廠（中際旭創、新易盛、Coherent 模組、Fabrinet 代工）是減分**（同一顆數下可插拔金額較少），對 **CPO 外部光源與雷射（Lumentum、Coherent 雷射）、CPO 交換晶片（Broadcom、NVIDIA）是加分**。受害公司的 A 情境 EPS 不必最高。光學合計、交換器、銅、網卡三情境相同；本模型的 CPO 只是 scale-out 光學金額的拆分，未模擬 CPO 把光學價值移入交換器整機（NVIDIA、Broadcom、Arista、智邦的數值不受影響）。scale-up 光學占比（`suopt_*`，2029 20%）仍為單一路徑。

## v0.2.2（2026-10-23）

Tokenomics 快照改 v20.1（`markets/chip_bridge/data/tokenomics_snapshot_v20.1.json`）；`IF_CapexIT`、`IF_GPUsPerGW` 等五個名稱與 v19.0 逐值相同，只更新版本標籤與說明文字，數值格 0 差異。

## v0.2.1（2026-10-22）

每顆 IT capex（`In` 頁 `it_GB`、`it_VR`、`it_RU`、`it_HP`）由 Tokenomics 快照（`markets/chip_bridge/data/tokenomics_snapshot_v19.0.json`）以 `IF_CapexIT ÷ IF_GPUsPerGW` 完整精度計算，不再取整（v0.2 為取到 0.1 $K 的寫死值，例：GB300 66.2 → 66.1944）。與 v0.2 相比只有 `In` 頁 4 個 IT capex 格與 `Mkt` 頁 AI IT 資本支出、網通 ÷ AI IT capex、NVIDIA 平台 ÷ IT 三列不同；網通 ÷ AI IT capex 變動 ≤ 0.003 個百分點，市場金額不變。AMD／TPU／Trainium／ASIC 的 IT capex 仍為 Assumed 寫死值。

## v0.2 結果（十億美元；Tokenomics v19.0；v0.2.1 的顯示值相同）

| | 2026 | 2027 | 2028 | 2029 |
|---|---|---|---|---|
| AI 網通市場 | 145（v0.1 153） | 292（296） | 338 | 385 |
| scale-up 占比 | 26% | 27% | 28% | 28% |
| 光學（其中 CPO；v0.2.3 起為情境 B） | 30（0.3） | 56（1.7） | 64（3.2） | 85（5.9） |
| 銅 | 26 | 51 | 59 | 61 |
| 網通 ÷ AI IT capex | 18.6% | 18.7% | 17.1% | 17.0% |

## 對帳（2026）

- NVIDIA 網通：本模型 62 vs 實際約 74（−17%；v0.1 −12%）。
- Broadcom AI 網通：本模型 12.6 vs 實際約 17.6（−28%；v0.1 −25%）。
- 光學：本模型 30 vs LightCounting 26（+16%；v0.1 +23%）；拆分比例未因 v0.2 重新校準；光模組廠營收加總約 26–30。初版機架外拆分（交換器 35／光 50／銅 15）使光學高 61%，已改 45／35／20 並註明為校準。

## 已知限制

AI 網通總市場與 scale-up／scale-out 美元拆分沒有公開資料（Dell'Oro、650 Group 付費）；ASIC 每顆網通內容、廠商歸屬比例、scale-up 互連中銅的比例皆為 Assumed；只含 AI 後端網路。

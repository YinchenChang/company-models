# markets/network — AI 網通市場模型（AI 半導體 Project 步驟 3）

**模型命題：** AI 網通的錢由「加速器顆數 × 每顆網通內容」決定；每顆內容隨 scale-up 域擴大與每 GPU 網卡頻寬倍增而上升，但 scale-up 的大頻寬大多由便宜的銅承載，所以金額主體仍在 scale-out（交換器、光模組、網卡）。

- 建置：`python3 markets/network/build_network.py raw.xlsx`，再以 LibreOffice headless 重算（錯誤 0）。
- 顆數取自 `markets/memory/market.py`（與 HBM 市場同一份加速器數量層）；NVIDIA 平台每顆內容以 `markets/chip_bridge` v0.2 為錨（GB300 約 $11.6K（v0.1 $13.0K；Tokenomics v19.0 機架基準價 $4.3M）、VR200 約 $24.5K；Bernstein GB200 約 $10.7K 為第二來源）。
- 頁面：README、In、Units、Mkt（依元件、依產品類別、網通 ÷ AI IT capex、NVIDIA／Broadcom 對帳）、Sens、Sources。

## v0.2.2（2026-10-23）

Tokenomics 快照改 v20.1（`markets/chip_bridge/data/tokenomics_snapshot_v20.1.json`）；`IF_CapexIT`、`IF_GPUsPerGW` 等五個名稱與 v19.0 逐值相同，只更新版本標籤與說明文字，數值格 0 差異。

## v0.2.1（2026-10-22）

每顆 IT capex（`In` 頁 `it_GB`、`it_VR`、`it_RU`、`it_HP`）由 Tokenomics 快照（`markets/chip_bridge/data/tokenomics_snapshot_v19.0.json`）以 `IF_CapexIT ÷ IF_GPUsPerGW` 完整精度計算，不再取整（v0.2 為取到 0.1 $K 的寫死值，例：GB300 66.2 → 66.1944）。與 v0.2 相比只有 `In` 頁 4 個 IT capex 格與 `Mkt` 頁 AI IT 資本支出、網通 ÷ AI IT capex、NVIDIA 平台 ÷ IT 三列不同；網通 ÷ AI IT capex 變動 ≤ 0.003 個百分點，市場金額不變。AMD／TPU／Trainium／ASIC 的 IT capex 仍為 Assumed 寫死值。

## v0.2 結果（十億美元；Tokenomics v19.0；v0.2.1 的顯示值相同）

| | 2026 | 2027 | 2028 | 2029 |
|---|---|---|---|---|
| AI 網通市場 | 145（v0.1 153） | 292（296） | 338 | 385 |
| scale-up 占比 | 26% | 27% | 28% | 28% |
| 光學（其中 CPO） | 30（0.3） | 56（1.7） | 64（3.2） | 85（5.9） |
| 銅 | 26 | 51 | 59 | 61 |
| 網通 ÷ AI IT capex | 18.6% | 18.7% | 17.1% | 17.0% |

## 對帳（2026）

- NVIDIA 網通：本模型 62 vs 實際約 74（−17%；v0.1 −12%）。
- Broadcom AI 網通：本模型 12.6 vs 實際約 17.6（−28%；v0.1 −25%）。
- 光學：本模型 30 vs LightCounting 26（+16%；v0.1 +23%）；拆分比例未因 v0.2 重新校準；光模組廠營收加總約 26–30。初版機架外拆分（交換器 35／光 50／銅 15）使光學高 61%，已改 45／35／20 並註明為校準。

## 已知限制

AI 網通總市場與 scale-up／scale-out 美元拆分沒有公開資料（Dell'Oro、650 Group 付費）；ASIC 每顆網通內容、廠商歸屬比例、scale-up 互連中銅的比例皆為 Assumed；只含 AI 後端網路。

# markets/network — AI 網通市場模型（AI 半導體 Project 步驟 3）

**模型命題：** AI 網通的錢由「加速器顆數 × 每顆網通內容」決定；每顆內容隨 scale-up 域擴大與每 GPU 網卡頻寬倍增而上升，但 scale-up 的大頻寬大多由便宜的銅承載，所以金額主體仍在 scale-out（交換器、光模組、網卡）。

- 建置：`python3 markets/network/build_network.py raw.xlsx`，再以 LibreOffice headless 重算（錯誤 0）。
- 顆數取自 `markets/memory/market.py`（與 HBM 市場同一份加速器數量層）；NVIDIA 平台每顆內容以 `markets/chip_bridge` v0.1 為錨（GB300 約 $13.0K、VR200 約 $24.5K；Bernstein GB200 約 $10.7K 為第二來源）。
- 頁面：README、In、Units、Mkt（依元件、依產品類別、網通 ÷ AI IT capex、NVIDIA／Broadcom 對帳）、Sens、Sources。

## v0.1 結果（十億美元）

| | 2026 | 2027 | 2028 | 2029 |
|---|---|---|---|---|
| AI 網通市場 | 153 | 296 | 339 | 386 |
| scale-up 占比 | 27% | 27% | 28% | 28% |
| 光學（其中 CPO） | 32（0.3） | 57（1.7） | 65（3.2） | 85（5.9） |
| 銅 | 28 | 52 | 60 | 61 |
| 網通 ÷ AI IT capex | 18.3% | 18.6% | 17.0% | 17.0% |

## 對帳（2026）

- NVIDIA 網通：本模型 65 vs 實際約 74（−12%）。
- Broadcom AI 網通：本模型 13 vs 實際約 17.6（−25%）。
- 光學：本模型 32 vs LightCounting 26（+23%）；光模組廠營收加總約 26–30。初版機架外拆分（交換器 35／光 50／銅 15）使光學高 61%，已改 45／35／20 並註明為校準。

## 已知限制

AI 網通總市場與 scale-up／scale-out 美元拆分沒有公開資料（Dell'Oro、650 Group 付費）；ASIC 每顆網通內容、廠商歸屬比例、scale-up 互連中銅的比例皆為 Assumed；只含 AI 後端網路。

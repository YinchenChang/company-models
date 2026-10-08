# markets/

跨公司的市場層模型（AI 半導體）：每 GW 晶片內容橋接表、加速器數量層、HBM／網通市場模型。公司模型（例如 SK 海力士）從這裡取市場切片。

| 資料夾 | 內容 | 版本 | 取數 |
|---|---|---|---|
| `memory/` | 記憶體共用市場層：季價路徑（情境 A／B／C）、加速器數量層、HBM 市場；見 `memory/README.md` | v0.2（2026-10-08） | — |
| `chip_bridge/` | 每 GW 晶片內容橋接表（GB300、VR200；HBM／非 HBM 記憶體分列；VR200 另列 2027 HBM 價吸收／轉嫁） | v0.1（2026-10-08） | Tokenomics v5.27 |

## chip_bridge

- 建置：`python3 chip_bridge/build_bridge.py raw.xlsx`，再以 LibreOffice headless 重算（`soffice --headless --convert-to xlsx --outdir dist raw.xlsx`），重算錯誤須為 0。
- 頁面：README（命題、驅動 → 推導）、In（輸入與標記、來源）、Rack（每架拆分）、Bridge（每 GW 樹狀拆分與恆等式核對）、Sens（單變數敏感度）、Requests（給 Tokenomics 的修改申請）、Sources。
- 核對：Bridge 的 IT 合計＝`IF_CapexIT`、合計＝`IF_CapexTotal`（GB300、VR200 2026 價、VR200 2027 吸收三欄差額 0；轉嫁欄差額＝轉嫁額）。
- **取數例外（待 Tokenomics 修改申請 R1）**：機架單價（Spec_Rack）、scale-out 網路比率與每 GPU 儲存成本（Inputs）目前不是 Interface 具名範圍，In 頁以數值抄錄並標註原格位置；R1 通過後改接 `IF_` 名稱。其餘 Tokenomics 值為 `IF_RacksPerGW`、`IF_GPUsPerGW`、`IF_CapexIT`、`IF_CapexFacility`、`IF_CapexTotal`。

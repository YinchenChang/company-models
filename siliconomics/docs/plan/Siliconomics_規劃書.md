# Siliconomics 規劃書（AI 半導體產業基準庫）

**版本：** P0.1（2026-10-27）　**擁有者：** AI 半導體 Project　**依據：** Andy 2026-10-25 13:26「公司模型完成後建本專案自己的產業事實庫 Excel（比照 Tokenomics）」；2026-10-27 07:59「規格如同 Tokenomics，與 TK 相關數據要與 TK 一致……是相關個股模型的基準事實來源」。

## 1. 為什麼要建：現況的問題

約 55 家公司模型各自帶著行業數字，同一個量常有兩個以上的值：

1. **單一來源只做到一半。** 加速器顆數、HBM、DRAM／NAND 價格路徑雖集中在 `markets/memory/market.py`（38 個建檔程式讀它），但那是 Python 常數，Andy 無法直接讀或改，也沒有標記、區間、檢查的治理層。
2. **其他子產業沒有共用層。** 液冷滲透率、AI 電源份額、ABF 供需、載板份額、伺服器市場、CPO 等由各公司模型各自假設；算力收支 2026-10-25 撰寫投資地圖時已列出 15 項跨模型不一致（例：緯創 vs 緯穎 2026 AI 占比 55% vs 48%；Dell vs 美超微 2026 GPU 系統市場 638 vs 596 億美元；台達電 vs 光寶的台達電份額約 70% vs 60–65%；Trainium 每顆內容 5.9K vs 2.5K）。
3. **2029 之後的顆數路徑分裂。** 市場層只給到 2028（2029 為係數），2029–2032 由台積電、NAND、網通等模型各自外推。
4. **Tokenomics 版本落差。** 下游仍讀 Tokenomics v20.1 快照；現行 v21.7 的 VR200 機架價低端由 7.8 改 7.5，`IF_CapexIT`、`IF_CapexTotal` 低成本欄因此不同（2 格）。

## 2. 規格：比照 Tokenomics

| Tokenomics | Siliconomics |
|---|---|
| Excel 是唯一事實來源；builder 只產生公式頁 | 同（`siliconomics/model/`、`siliconomics/builder/`） |
| 藍字輸入由 Excel 擁有；重建時讀回 | 同（In 頁；`build.py --base`） |
| SRC_／DB_Evidence 來源登錄、Verified／Interested-party／Analogy／Assumed／Derived 標記；Analogy、Assumed 給區間 | 同標記；v0.1 來源寫在 In 頁 H 欄、DB_Evidence 頁建立；獨立 SRC 頁於 v0.2 起依子產業建立 |
| Interface `IF_` 具名範圍＝下游唯一連結點；L1 常用推算值 | Interface `SIF_` 具名範圍（輸入、推導、TK 鏡像） |
| Checks、GOV_Errors＝0 才可合併 | Checks、`SGOV_Errors`＝0 才可匯出快照 |
| Decisions 決策登錄 | 同 |
| 下游契約 `Tokenomics_downstream_contract.md` | `docs/Siliconomics_downstream_contract.md` |
| `import_tokenomics.py` 版本固定快照 | `tools/export_snapshot.py` 版本固定快照 |
| CI：LibreOffice 重算、錯誤格 0、逐格比較 | 併入 `ai-semis-build`：本庫重建＝現行檔；快照＝現行檔；下游 58 個目標逐格相同 |

**與 Tokenomics 一致的做法：** 本庫不重估任何 Tokenomics 權威的量（每 GW 顆數、機架數、機架價、功率、每 GW 資本支出、NonNV 比例、產業部署路徑）。`TK_In` 頁以 Tokenomics 快照鏡像（綠字、不得修改），Interface 以 `SIF_TK_*` 轉出；Checks K3 列出下游仍用舊版 TK 的格數。Tokenomics 升版 → 重抽快照 → 重建本庫 → K3 指出哪些下游要重算。

## 3. 範圍：行業變數與公司變數的界線（決策 S4，預設）

- **本庫（行業變數）：** 跨公司共用或加總受限的量——加速器顆數與每顆內容、各市場規模與價格路徑、技術替代滲透率（CPO、銅轉光、液冷、混合鍵合）、同一市場各供應商份額表（合計 ≤ 100%）、供需缺口、情境 A／B／C 定義。
- **公司模型：** 會計年度錨點、毛利率與費用、資本支出、股數、估值參數、公司特有的份額相對值。
- 依據：Andy 2026-10-25 13:13「個股 A／B／C 一律讀同字母行業情境（技術替代屬行業變數）」。

## 4. 分段計畫

每段一個 PR、可單獨通過 CI、可單獨合併；每段結束時下游「預期變動以外的差異＝0」。

| 段 | 內容 | 下游數值 | 狀態 |
|---|---|---|---|
| v0.1 | 骨架（README、Scen、In、Path、Mkt、TK_In、Conflicts、Checks、Decisions、DB_Evidence、Interface）；Tokenomics v21.7 鏡像；`markets/memory/market.py` 的全部輸入搬入、`market.py` 改讀快照；跨模型衝突登錄；CI | **不變**（58 個目標逐格相同） | 完成（PR #180） |
| v0.2 | 補齊 Assumed 列區間與缺來源（G2、G3 升為 ERROR）；2029 平台別顆數成長改為單一來源（CF18：台積電、網通改讀；NAND 2029 顆數年增 10% → 13.8%） | NAND、SanDisk、Kioxia 數值變動；其餘只有 In 頁文字與區間 | 完成（PR #181） |
| v0.3 | 運算層：加速器每顆售價與 ASIC 核心價（M6，CF13、CF14）、ASIC 設計服務份額表（M7、ShareTbl，CF19）；部署 GW（CF22）依賴網通，移到 v0.4 | 不變（只搬位置） | 完成（PR #186） |
| v0.4 | 網通市場層輸入（每顆網通內容、每顆 IT capex、技術轉換、CPO 三情境）、vr26、Deploy 頁（部署 GW 單一路徑，CF22）；橋接表改讀 TK v21.7（CF17）。CPO 份額表（CF21）延到 v0.4.1 | 目標價不變；Amkor Recon、Dell 低端敏感度 | 完成（PR #189） |
| v0.4.1 | CPO 光學金額份額表（M12、ShareTbl、K6-cpo；CF21）；五家光學廠改讀本庫 | 不變（只搬位置） | 完成（PR #191） |
| v0.5 | ABF 載板（份額表 CF20、每顆金額路徑、外部缺口單一路徑 CF03 部分）；CoWoS 供需輸入（M14） | 目標價不變；台積電 In 頁 9 格 | 完成（PR #192） |
| v0.5.1 | ABF 供需表（M15、ABF 頁；CF03 其餘）；欣興、南電、景碩產能與味之素膜上限移入 | 不變（公司滿足率改讀本庫缺口待 Andy） | 完成（PR #193） |
| v0.5.2 | 每 GW PCB／CCL（M16、PGW 頁）；ABF 滿足率改讀試算（報告） | 不變 | 完成（PR #194） |
| v0.6 | AI 電源（M17、份額表、K6-pwr；CF11 只建表、改值待 Andy）、AI 伺服器組裝（M18；CF09 緯穎 AI 占比單一值）、Rack 頁（組裝市場、商用 GPU 系統市場 CF10、AI 電源市場） | 戴爾（CF10）、緯創（CF09）；其餘不變 | 完成（PR #197） |
| v0.6.1 | 液冷（M19：滲透率、每架內容、分歧管、橋接表 cool 移入；Rack 頁 D 節；液冷份額表 CF04；CF12 市場定義為機架內，台達電屬機架外） | 不變（只有 In 頁文字） | 完成（PR #198） |
| v0.7 | 設備市場層行業輸入（M20）、測試設備（M21、Equip 頁、CF15） | 不變 | 完成（PR #200） |
| v0.7.1 | HBM 堆疊技術與混合鍵合比例（M22；disco、hanmi、besi 同讀）、NAND 市場輸入（M23） | 不變（只有 In 頁文字） | 完成（PR #202） |
| v0.8 | HBM 營收份額表（M24、ShareTbl、K6-hbm；三家記憶體廠改讀）；台積電 HBM 2027 缺口改讀本庫推導值 | 不變（台積電只有對照列 21% → 21.0%） | 完成（PR #213） |
| v0.8.1 | AI 伺服器組裝份額表（M25：廣達、鴻海、緯創本體、緯穎；緯穎 ASIC 傾斜係數）＋ K6-asm；CF23 登錄並試算（戴爾、美超微分母不同，不入表） | 不變（CF23 改值會使緯創 FY2027 EPS +11.4%，待 Andy） | 本版 |
| v0.8.2 | 共用匯率（新台幣、韓元、日圓）；現各模型不一致，改值前先試算 EPS | 會改 EPS（判斷類） | 排程（試算後待 Andy） |
| v0.9 | 衝突登錄清理：CF05–08、CF16 文件類；CF01 南電新廠資本支出、CF02 欣興資本支出對照 | 文件類不變；CF01 改南電 EPS | 排程 |
| R4 | 算力收支產業部署路徑三情境上線後，情境係數改由部署 GW 推導，並延伸到 2032 | 全部 | 待算力收支（預計 10-31 後） |

## 5. 待 Andy 判斷（預設照做、可撤回）

1. **名稱：** Siliconomics（中文：AI 半導體產業基準庫）。替代：Chipnomics；或只用中文名。
2. **界線（第 3 節）：** 份額表放本庫（合計受限），公司只保留「份額相對值」。替代：份額全留在公司模型，本庫只放市場規模。
3. **2030–2032：** 不在 v0.1 自訂，等 R4；R4 延遲超過 11-07 時，v0.2 先以「2029 年增率線性收斂到長期成長率」的共用規則填入。
4. **衝突的建議單一值：** 見活頁簿 Conflicts 頁；各項在對應子產業的版本生效，生效時逐項列出對公司 EPS 與目標價的影響。

## 6. 已知限制（v0.1）

- v0.1 搬入時 In 頁 50 列 Assumed 缺區間、2 列缺來源（`market.py` 原狀）；v0.2 已補齊。
- 只有記憶體與顆數層在本庫；其餘子產業的行業變數仍在各公司模型，衝突暫列 Conflicts 頁。

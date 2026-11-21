# ABB（ABBN）個股模型 v0.1.1

成品日 2026-11-21；股價基準日 2026-11-20（收盤）；build_id abb-v0.1.1-dca153d553。
金額：十億美元（報表幣別）；股價與目標價：瑞士法郎（SIX：ABBN）；匯率 CHF=X 0.83593（每美元）。

## 本版要點（v0.1.1，2026-11-21）
- 現價與同業報價改 2026-11-20 收盤（工作單 px1120）。現價 69.96 → 65.34 CHF；同業前瞻 P/E 中位數 19.1 → 17.5；加權目標價 52.11 → 49.52 CHF；上檔空間 -25.5% → -24.2%；評等「放空」→「中立」（越過 -25% 邊界，由模型公式重算）。
- 同業前瞻 P/E（Yahoo 2026-11-17）× 11-20 收盤 ÷ 11-17 收盤：ETN 20.50→18.93、SU.PA 19.18→18.37、NVT 15.56→14.51、VRT 19.11→16.69。共識 EPS、rf、β、ERP、營運假設、情境權重不變。
- 預測匯率 CHF=X 維持 0.83593（2026-11-13），依工作單不換。
- 未換價同業：無。ABBN.SW、SU.PA 的 11-20 值為報價資料最近成交價（Yahoo 日線尚無該列）。
- Sens 快照已重跑（sens.json）；Coord 不適用，無快照。

## 結果（Excel Val 頁為準）
- 現價 65.34 CHF；DCF 目標價 41.99 CHF；Comparable 目標價 57.05 CHF；加權目標價 49.52 CHF（59.24 美元）。
- 上檔空間 -24.2%，評等「中立」，距「放空」邊界（-25%）僅 0.8 個百分點。
- 2026-11-17 獨立核對後修正：Robotics 出售狀態更新（2026-11-05 報導中國禁止 SoftBank 收購其中國業務，交易未完成、協商中；估值仍 50 億美元 [Assumed]，Sens 下緣由 44 億擴大到 35 億，加權目標價、上檔空間、評等不變）；In 頁資本支出註記「42 億」更正為「420 億」；dps_tgt 與 rf 日期差異於待判斷註明；命題旁並列 FY2025 資料中心占比約 7%；ABBN.md 補出處標記。
- 價格帶 33～48 CHF（DCF 情境 C～A）。

## 檔案
| 檔 | 用途 |
|---|---|
| dist/20261121_ABB個股模型_v0.1.1.xlsx | 事實基礎（LibreOffice 重算，無錯誤值） |
| dist/20261121_ABB個股模型_v0.1.1.html | 展示層（由 html.json 與 xlsx 產生） |
| build_abb.py、readme_abb.py、segments.json | 建置腳本與公司原文分部資料 |
| sens_abb.py、sens.json | Sens 案例（覆寫 In 頁鍵值後重算；ABB_OVERRIDE 與 ABB_SNAP=1） |
| gen_html_json.py、html.json | HTML 設定檔（引用格位，不手打數字） |

## 來源與等級
- 公司原文：Q4 2025 financial information（FY2025／FY2024，Robotics 重編後）與 Q3 2026 新聞稿（9M 2026、資產負債表、指引）[Verified]。
- 股價：投資地圖的收盤價 CSV（closes_20261109_20261120.csv 等）；Yahoo 僅作交叉，差異記在 Recon 頁。
- 同業前瞻 P/E、共識、β：Yahoo 2026-11-17（Interested-party，單一來源）。
- 美元 rf 5.10%：FRED 2026-11-12。
- AI 資料中心變數：Wattonomics 快照 v0.3.6（sha256 前 8 碼 acb92d6f）；ABB 不讀 Siliconomics（設施供應商，機櫃外屬設施）。ABB 份額＝資料中心營收 ÷（配電＋UPS）為 Derived，區間與 Sens 見模型。

## 公司原文與 Yahoo 的差異（Recon 頁）
- FY2024 營收 30,583（公司與 Yahoo 皆同）；工作單的 32,850 為未重編數（含 Robotics 2,267）。
- 營業利益：Yahoo 不含「其他收入（淨額）」，FY2025 差 318、FY2024 差 -18。
- FY2022–FY2023 Yahoo 數未重編，FY2023A 僅供參考。
- 非控制權益：公司 2025-12 為 559，Yahoo 2026-09 為 483。
- Yahoo 的 ABB 前瞻 P/E 18.57 把 CHF 股價除以 USD EPS，不採用。

## 待判斷（區間與標記；詳見 Excel README 頁第 27–40 列）
1. 資料中心拆分：FY2026E 占電氣化 17%–27%（基準 22%）[Assumed]。
2. ABB 占 Wattonomics 配電加 UPS 份額：FY2027 6%–9%、FY2032 5%–9% [Assumed]。
3. Wattonomics 係數 0.8–1.2；FY2031–32 設施支出下降使終值年偏保守。
4. Robotics 價值 35–54 億美元（基準 50 億）[Assumed]：2026-11-05 報導中國外資安全審查禁止 SoftBank 收購其中國業務，交易未完成、雙方仍在協商（Robotics247、Supply Chain Digital；Interested-party）。基準：交割則收淨款，不交割則保留事業，價值仍約當出價；Sens 低 35 億＝交易失敗且保留事業價值打折，高 54 億＝全額企業價值；交割與否待公告。
5. Rotork 收購與三菱電機低壓配電 JV 不計入（視為價值中性）；Rotork 條款待補。
6. β 1.05 單一來源；區間 β_u ± 0.3 [Analogy]。美元 rf 5.10%（FRED 2026-11-12）與其他美元模型（2026-10-23 的 5.00%）日期與數值不一致，未對齊，不改值。
7. 非控制權益 0.559 至約 3.0 十億美元（影響約 -2%）。
8. 淨退休金資產 0.619 十億美元未計入（計入約每股 +0.3 美元）。
9. Q4 房產出售收益 2.4 億 CHF 全額計入（區間 0–全額，每股約 ±0.2 美元）。
10. 評價日後股利（2027 年股利）未扣除，dps_tgt＝0 不改值（若扣除目標價約 -1.3%、上檔空間約 −26%）。
11. 買回視為價值中性；EPS 另計每年 -1% 股數。
12. 評等邊界：距「中立」不到 1 個百分點，請以價格帶與 Sens 閱讀。
13. 共識口徑：FY2026 EPS 共識 3.244 隱含 Q4 偏高，口徑待核。

## 建議補進上游（不在本分支修改）
- Wattonomics：機櫃外設施供應商份額對照層；FY2031–32 下降段是否為定案。
- 三層協調規則：寫入「設施供應商讀 Wattonomics」的例外。
- 同業 β：建立電氣設備同業自算 β。

## 命名與範圍
- Automation 即原 Process Automation（2026 年起公司新聞稿用名）。
- Robotics 為停業單位，不在 FCFF；其價值（出售淨款或保留事業）以 In 頁非營業投資加入。出售交易目前受中國審查阻礙，未完成。

## Coord
- Coord／alt 讀數：不適用（電力與機房範圍定案 2026-11-17）：機房設施供應商只讀 Wattonomics 廠房拆分，該數字只有 Wattonomics 一個權威層、沒有 Siliconomics 側的另一讀數，三層協調 C4 兩讀數不適用；與 Wattonomics 的差異走對帳單。 CI 建檔目標已由主串加入（2026-11-17）。

## 檢查結果（v0.1.1）
- 本版跑過：check_company_html、check_layout_v13（html＋--xlsx）、check_governance（含 --html）、check_sens_coverage --require-keycol、baseline_coordination/check.py（0 項不符）、check_cached_values；LibreOffice 重算無錯誤值。layout_check.py 本環境沒有 playwright，未跑。

## 檢查結果（v0.1.0，上一版）
- check_company_html.py abb：全部通過。
- check_layout_v13.py（含 --xlsx）：通過。
- check_governance.py（含 --html）：全部通過；Sens 與 README、HTML 的 build_id 一致。
- layout_check.py：@1280 與 @390 無水平捲動，7 張圖、63 條長條皆有資料。
- LibreOffice 重算：無 #REF!／#VALUE!／#DIV/0!。

## 重建
1. `python3 abb/build_abb.py` 後以 LibreOffice 重算，輸出至 dist。
2. 模型數字變動時重跑 `python3 abb/sens_abb.py` 更新 sens.json。
3. `python3 abb/gen_html_json.py`，再 `python3 tools/company_html/build_company_html.py abb`。
4. build_abb.py、sens.json、segments.json、Wattonomics 快照任一變動，build_id 隨之改變，需重跑全部步驟。

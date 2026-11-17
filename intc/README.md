# intc/：英特爾（INTC）個股模型 v0.1.0（2026-11-17）

AI 半導體 Project 代工層（兼 CPU／AI 產品線）。樣板 `tsmc/`、`umc/`（代工），產品線結構參 `amd/`。

| 檔案 | 內容 |
|---|---|
| `build_intc.py` | 產生器（公司設定、季度輸入、情境引擎、DCF、相對倍數） |
| `v010_post.py` | 後處理：Hist 頁、PL 標準損益表與分部營收推導、RevDCF、Val 0.7、Sens 相對基準表、Coord C0、治理（build_id、模型狀態、來源誘因） |
| `show_intc.py` | README 展示層（命題、重點觀察、命題回覆） |
| `gen_html_json.py` → `html.json` | 版面 v1.3 設定（直接依 xlsx 標籤取列；勿手改 `html.json`） |
| `segments.json`、`coord_alt.json`、`gov_snapshot.json` | 分部宣告、Coord 另一側讀值、治理快照 |
| `dist/` | 成品 xlsx（LibreOffice 重算，錯誤 0 格）與 HTML |

## 結論（2026-11-13 收盤價 59.66）

DCF 14.87、Comparable（相對倍數調整後）17.89（未調整 25.94），加權目標價 16.38，上檔空間 -72.5%，級距「小於 -50%」，投資建議強力放空。主因：現價隱含的前瞻 P/E 遠高於同業與英特爾自身歷史，Foundry 仍虧損且資本支出占營收二成以上，情境 B 的自由現金流撐不起現價；市場共識目標價 90.65 與本模型的差距主要來自倍數因素。

## 資料來源與等級

財報取自公司新聞稿（3Q26、1Q26、2Q26、4Q25 及 FY2024、FY2023 稿；SEC 8-K／IR）為 Verified；4Q26 指引為 Interested-party；Foundry 利益率、外部代工營收、DCAI AI 相關、CCPG 成長為 Assumed。Yahoo 彙整數僅作交叉檢查與同業 EPS：3Q26 營業利益 Yahoo 2,576 對公司 2,522；FY2025 營業利益 Yahoo -23 對公司 -2,214（以公司為準）。股價只用 `markets/invest_map/data` 內 11-13 收盤檔。

## 待判斷（預設值與範圍）

1. 資本支出取公司口徑總額（投資活動＋融資活動的 PP&E 增加，未扣補助與夥伴挹注）。
2. 租賃負債 0（已含於債務）；存貨取 1Q26 的 8.481 十億美元（數值待複核）。
3. Escrowed Shares 評價損失為非現金，FY2026 後不預測；不計美國政府認股權證稀釋。
4. 4Q26 分部拆分為假設；FY2030 起收斂法淨利率往 FY2029 乘 0.9（A）、0.8（B／C 範圍）。
5. Foundry 利益率路徑、CCPG／DCAI／外部代工假設為 Assumed，區間見 In 頁低／高。
6. 同業取 TSM、AMD、2303.TW（有 Bloomberg 資料者）；相對倍數係數 0.69，夾在 0.5–2.0。
7. rf 日期與價格日不一致；kd 4.25%；aig_L 0.40。
8. 前瞻 EPS 為 GAAP，市場共識多為 Non-GAAP，差距不代表錯誤。
9. EPS 分子調整：3Q26 淨利除以稀釋股數得 0.626，公司公告 0.60。

## 建議補進 Siliconomics

伺服器 CPU、客戶端 PC、代工利益率、外部代工營收（目前皆為公司層級假設，已標「待併入 Siliconomics」）。讀取：um28／um29（A、C 取 markets/scen，B 取 Siliconomics）、wta_27／28／29（只進敏感度）、eq_g_fd_26／27／28（僅 Recon、Coord 對照）、RECO_FLAG、ALT_RATIO。

## 工具缺口（本資料夾不改 tools/）

- Intel 不在 `tools` 的 universe／peers CSV，建置時以 monkeypatch 補入。
- `alt_run` 需 `.github/workflows/ai-semis-build.yml` 加 intc 目標與 coord_meta，現以 /tmp 修補版產生 `coord_alt.json`。
- `tools/relative_multiple.py` 的 Val 0.5 標題寫入人名字樣，公開 repo 應移除。
- 無 `sens_intc.py`：Sens 頁在活頁簿內逐案重算（同 UMC）。

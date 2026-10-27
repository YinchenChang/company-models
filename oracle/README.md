# Oracle（ORCL）個股模型 v0.4（命題＋營收分部架構）

**命題（一句話）**：OCI 已簽約的 backlog（RPO 約 6,640 億美元，約一半來自 OpenAI／Stargate）能否以高於資金成本的報酬轉成營收與現金，同時不把資產負債表撐破。

**核心問題**：FY2029 的 OCI 營收與營業利益率，相對於資本支出與租約的擴張，以及資金缺口（自由現金流為負）要如何補。

規格：`docs/plan/個股模型原則_命題加營收分部.md`（命題模組＋官方營收分部 → 損益表）；工作單：`mag/docs/workorders/20261027_MAG_v0.3_個股模型重建_分部架構.md` 第 1、2、4、5、6 節（第 3 節以 Oracle 段取代）。報告：`docs/reports/20261027_oracle_v0.4_分部架構.md`。

## 結果摘要（2026-10-27；現價 US$126.06〔2026-10-26 收盤〕）

| | 情境 A（樂觀） | 情境 B（基準） | 情境 C（保守） |
|---|---|---|---|
| FY2029 OCI 營收（其中 AI 算力） | 101.0（86.4） | 93.6（79.0） | 56.5（41.9） |
| FY2029 OCI AI 營業利益率（含 AI 折舊、租金） | 14.8% | 6.3% | −42.8% |
| FY2029 AI 增量 ROIC（WACC 12.2%） | 4.4% | 1.7% | −7.5% |
| FY2029 EPS（GAAP） | 9.56 | 7.19 | 1.11 |
| 收支關卡：FY2027–FY2031 自由現金流扣股利 | +127.9 | +90.0 | +17.5 |
| 情境目標價 | 151.9 | 140.1 | 100.7 |

加權目標價 **US$133.2**（DCF 55.5 × 50%＋分部加總 211.0 × 50%；情境 25／50／25），上檔空間 **+5.7% → 中立**。金額單位十億美元（EPS、目標價為美元）。數字以 `dist/` 的 xlsx 為準。

## 檔案

| 檔案 | 內容 |
|---|---|
| `build_oracle.py` | 獨立 Python → Excel 產生器（openpyxl，全部活公式；藍字輸入、黑字公式） |
| `dist/20261027_Oracle個股模型_v0.4.xlsx` | LibreOffice headless 重算後的成品（公式錯誤 0） |
| `dist/20261027_Oracle個股模型_v0.4.html`、`html.json` | `tools/company_html/build_company_html.py oracle` 產生；`check_company_html.py` 通過 |
| `segments.json` | 官方分部宣告；`tools/segment_check/check_segments.py oracle` 通過 |
| `scripts/sens_v04.py` | 敏感度：一次動一個輸入，完整重算（寫 `data/sens_v04.json`，再建置時寫入 Sens 頁） |
| `scripts/make_html_json.py` | 由建置列號對照產生 `html.json` |
| `data/` | 事實總帳（v0.1a SEC 一手、v0.3.1 增補）、共識、Tokenomics v19.0 快照（沿用） |
| `legacy_v0.3/` | v0.3.1 舊引擎（CRWV 樣板：HTML＋Excel 雙引擎、company.json、驗證腳本），保留不刪 |
| `dist/_待刪除/` | v0.3.1 舊成品 |

## 建置

```
python3 oracle/scripts/sens_v04.py --jobs 6          # 敏感度（有改公式或輸入時才需重跑）
ORCL_ROWMAP=/tmp/rowmap.json python3 oracle/build_oracle.py /tmp/raw.xlsx
soffice --headless --calc --convert-to xlsx --outdir /tmp/out /tmp/raw.xlsx
cp /tmp/out/raw.xlsx oracle/dist/20261027_Oracle個股模型_v0.4.xlsx
python3 oracle/scripts/make_html_json.py /tmp/rowmap.json
python3 tools/company_html/build_company_html.py oracle && python3 tools/company_html/check_company_html.py oracle
python3 tools/segment_check/check_segments.py oracle
```

## 工作表與驅動 → 推導對照

順序：README → In → Q → Thesis → Seg → PL → CF → Val → Recon → Sens → Sources。

| 頁 | 輸入 | 推導量 | 因果方向 |
|---|---|---|---|
| In | 一列一鍵（值、單位、標記、來源、低／高） | — | 所有公式以鍵名引用 |
| Q | 事實總帳：8 季＋FY2024–FY2026 營收線、分部、分部利潤、損益、現金流 | 近四季年增率、分部利潤率、m_sw 對照 | 實績核對列（營收線＝合併、分部＝合併、分部利潤合計、自由現金流）皆 0 |
| Thesis | MW 速度、上限、延誤、計費比例、世代組合、Tokenomics 每 MW 持有成本／資本支出／營運費用、k、未起租租約 | 已連網與計費 MW → OCI AI 營收；AI 營運費用、租金、資本支出、折舊 → OCI AI 分部利潤貢獻、ROIC；Q1 校準 m_sw；RPO 檢查 | 只驅動 OCI 的「AI 算力」子列；AI 資本支出接到折舊與營業利益 |
| Seg | 近四季年增率 → 長期年增（In G 節） | SaaS、授權、支援、硬體、服務；OCI＝AI＋非 AI → 三分部 → 合併營收 | 子列＝母列、分部＝合併 |
| PL | m_sw、硬體與服務利潤率、未分攤費用、稅率 | 分部利潤 → GAAP 營業利益 → 利息（CF）→ 淨利 → EPS | 分部利潤加總到合併營業利益 |
| CF | 客戶出資覆蓋比、FY27 資本支出指引、瀑布參數 | 營運現金、資本支出、自由現金流、股利 → 收支關卡；現金 → 投資級債 → 新股 → 高息債 | 「收支關卡：自由現金流扣股利（五年合計）」 |
| Val | CAPM、g、同業倍數、權重 | DCF（UFCF）、分部加總（OCI AI × neocloud、其餘 × 軟體）→ 加權目標價、投資建議 | 「上檔空間：加權目標價」「投資建議」 |
| Recon | 共識、公司指引、OCI 路徑 | 逐列差距；v0.4 vs v0.3.1 | 不作校準 |
| Sens | 20 個輸入的低／高 | 加權目標價、FY2029 OCI AI 營業利益率、收支關卡 | 建置時完整重算 |

## 與 v0.3.1 的差異

v0.3.1 是 CoreWeave 樣板：OCI 以 MW × 每 MW 收入驅動，SaaS、軟體、硬體、服務壓成「傳統事業」四條成長率，OCI 利潤以 EBITDAR 率校準，評價 DCF 45%＋分部 EV/EBITDA 55%（OCI 6x）。v0.4 改為官方分部（Cloud and software、Hardware、Services）與六條營收線，OCI 拆 AI（命題）／非 AI；AI 營運費用由 Tokenomics 正向推導，AI 折舊與租金進分部利潤與 EPS；評價改為 DCF 50%＋分部加總 50%、情境 25／50／25。融資瀑布保留為收支模組。v0.3.1 基準目標價 84.2（賣出）→ v0.4 情境 B 140.1、加權 133.2（中立），拆解見報告第 6 節。

## 已知限制

1. OCI 的 AI／非 AI 拆分不存在公開資料：Q1 FY27 非 AI 為殘差（占 34%，v0.3.1 區間 12–30%）。
2. 季度分部利潤只取得 FY26Q1（雲端與軟體）與 Q1 FY27；其餘季度未找到（10-Q 在建置環境無法開啟）。
3. 容量上限＝已簽站點 8,778 MW-IT（沿用 v0.3.1），不假設新站點：FY2029 起 AI 資本支出降到汰換水準，自由現金流轉正。
4. OCI 同業倍數只有兩家（CoreWeave、Nebius）；硬體與服務以軟體倍數計。
5. 預付以外的營運資金變動不建模；在帳租賃負債持平；到期借款以原利率展延；續約價格衰退不建模。
6. Tokenomics 快照維持 v19.0（v0.3.1 沿用），下次季度滾動時升級。

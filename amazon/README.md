# Amazon（AMZN）個股模型 v0.3 — 命題模組＋官方營收分部 → 損益表

成品：`dist/20261027_Amazon個股模型_v0.3.xlsx`（全部活公式，LibreOffice 重算、公式錯誤 0）與同名 `.html`（`tools/company_html` 產生）。
規格：`docs/plan/個股模型原則_命題加營收分部.md`；工作單：`mag/docs/workorders/20261027_MAG_v0.3_個股模型重建_分部架構.md`；報告：`docs/reports/20261027_amazon_v0.3_分部架構.md`。

## 命題一句話
AWS 的 AI 容量（Trainium＋NVIDIA）能否以高於資金成本的報酬變現，零售與廣告的現金流能否支應資本支出。

## 核心問題
FY2028 AWS 營業利益率與自由現金流（扣股利後）是否轉正並維持。

## 結果摘要（修正一後；資料截至 2026Q2；現價 US$249.09，2026-10-26）
| 項目 | A（樂觀） | B（基準） | C（保守） | 加權 |
|---|---|---|---|---|
| 情境目標價（DCF 50%＋分部加總 50%；讀法 2 主值） | 265.28 | 199.45 | 178.98 | **210.79** |
| DCF／分部加總 | 205.64／324.91 | 139.30／259.60 | 118.73／239.22 | 150.74／270.83 |
| 讀法 1（自用 AI 只計成本）目標價 | 243.59 | 178.26 | 157.78 | 189.47 |
| FY2028 AWS 營業利益率 | 43.6% | 34.7% | 31.3% | |
| FY2028 對外 AI ROIC（WACC 11.3%） | 16.1% | 6.6% | 3.6% | |
| FY2028 自由現金流扣股利 | 19.3 | −15.5 | −27.6 | |
| 收支關卡：FY2026–FY2030 自由現金流扣股利合計 | 192.9 | 10.2 | −51.6 | 40.4 |

上檔空間 −15.4% →**投資建議：中立**（五級規則；讀法 1 為 −23.9%，同為中立）。悲觀情境價（C 情境 DCF）118.73；市場共識目標價 332.24。
修正一（獨立核對後）：自用 AI 讀法 2 為主值、SBC 口徑一致、AWS AI 以自身 DCF 進分部加總、WACC（Blume β、rf 2026-10-26、類債務）、終值前 MW 速度收斂、淨負債含融資租賃；修正前後對照見報告「修正一」節。

## 結構（工作表順序固定）
README → In → Q → Thesis → Seg → PL → CF → Val → Recon → Sens → Sources

| 頁 | 輸入 | 推導量 | 因果方向 |
|---|---|---|---|
| In | 全部輸入一列一鍵（定義名稱＝鍵；低／基準／高；標記與來源） | — | — |
| Q | `data/amazon_facts_20261008.json`（10-K、10-Q、新聞稿） | FY2023–FY2025、最近 8 季、半年合計；核對列 | 實績 → 校準 |
| Thesis | 對外 AI MW（期初 2,080、每年新增 pace）、Tokenomics v21.4 持有成本與資本支出、k、對外比例 | AWS AI 營收、AI 資本支出、AI 折舊、AWS AI 營業利益、自用 AI 成本、對外 AI ROIC、打平 k | MW × 每 MW 收入 → 營收；MW ÷ 對外比例 × 每 MW 成本 → 資本支出 → 折舊 → 分部營業利益 |
| Seg | 零售六條營收線的長期年增 | 營收線 → 北美／國際（地區比例）；AWS＝AWS AI＋AWS 其他 | 營收線加總＝分部加總＝合併營收（核對列 0） |
| PL | 廣告利潤率、零售本業利益率改善、稅率 | 分部營業利益 → 合併營業利益 → 淨利 → EPS | 廣告驅動零售分部利益；自用 AI 成本扣零售 |
| CF | SBC 比例、其他營運項目比例、最低現金 | 營運現金 − 資本支出 − 融資租賃本金 → 自由現金流；借款、淨負債；收支關卡 | 現金不足即舉債 → 利息回到 PL |
| Val | WACC（CAPM）、g、RONIC、同業倍數、持股、自用 AI 讀法 | DCF、分部加總（AWS AI 用自身 DCF）、情境目標價、上檔空間、投資建議；讀法 1 並列 | 50%／50%；情境 25／50／25 |
| Recon | 共識（MarketScreener、S&P）、公司指引、v0.2.2 | 差距、驗證列、差異拆解 | 只對照，不校準 |
| Sens | 32 個案例、WACC × g 網格、讀法 1 三情境（每案完整引擎區塊） | 情境 B 目標價變動（龍捲風） | 一次動一個輸入 |

## 與 v0.2.2 的差異
- 營收主軸由「AI 雲端 MW × 每 MW 收入＋成長率線」改為**官方分部**（North America、International、AWS）＋**10-K 營收線**；命題只驅動 AWS 之下的 AWS AI 子列。
- AI 資本支出接到**折舊與分部營業利益**（對外 → AWS、自用 → 北美／國際），不只接現金。
- 融資瀑布、站點、租賃承諾、RPO 等 CRWV 模組移除；評價改為 DCF 50%＋分部加總 50%，情境加權。
- 目標價 v0.2.2 基準 201.71 → v0.3 加權 210.79（拆解見 Recon 頁 D 節）。
- 舊引擎（segA–E.js、tail.js、build_xlsx.py、company.json、kspec.py、scripts/ 等）移至 `legacy_v0.2/`（保留不刪；v0.2.2 交接檔仍在 `docs/handoff/`）；舊成品移至 `dist/_待刪除/`。

## 重建
```
python3 amazon/build_amazon.py                 # 產生 dist/20261027_Amazon個股模型_v0.3.xlsx（未重算）與 html.json
# 重算（與 CI 相同方式）：
python3 amazon/build_amazon.py /tmp/raw/20261027_Amazon個股模型_v0.3.xlsx
soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir amazon/dist /tmp/raw/20261027_Amazon個股模型_v0.3.xlsx
python3 tools/company_html/build_company_html.py amazon
python3 tools/company_html/check_company_html.py amazon
python3 tools/segment_check/check_segments.py amazon
```
`AMZN_WRITE_HTMLJSON=1` 時以指定輸出路徑建置也會重寫 `html.json`（列號由產生器決定，不手寫）。

## 已知限制
1. 公司未揭露 AI 在役 MW、對外比例與 AI 營收：AWS AI／AWS 其他的歷史分拆為估計（評價日年化 20.1 vs 公司「>25」）。
2. 分部與營收線共識：未找到（Recon 寫明）。
3. 同業倍數沿用 2026-10-07；kd 以 rf＋AA 利差代理；Tokenomics 維持 v21.4。
4. 模型 EBITDA（PL 頁）只加回 PP&E 折舊且不加回 SBC，低於供應商口徑（差距約各半）；分部加總另用 SBC 加回口徑；自由現金流低於共識。
5. 2026Q3 財報（2026-10-29）未納入，之後以 v0.3.1 滾入。

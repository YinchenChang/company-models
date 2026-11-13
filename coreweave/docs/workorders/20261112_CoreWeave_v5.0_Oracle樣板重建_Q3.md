# 工作單：CoreWeave v5.0｜Oracle 收支型樣板重建＋2026Q3 滾入＋v4.13 驗收採用｜2026-11-12｜S1 線

- **依據：** Andy 2026-11-12 09:05「請繼續」＝確認採用規劃書（Drive `個股模型/CoreWeave/20261112_CoreWeave驗收v4.13_採用規劃書.md`），J1–J5 照預設。Andy 原話（08:55）：「Coreweave 已经公布财报了，所以要更新 Coreweave 的财报。再将 html 的版面套用 Oracle 最新的版本。」
- **樣板：** `oracle/`（v0.5.1：`build_oracle.py`、`funding_v05.py`、`make_v05.sh`、`scripts/*`、`html.json` 由 `scripts/make_html_json.py` 產生）。CoreWeave 沒有傳統業務，Seg 只有「AI 雲端」一個分部（另列「其他／服務」若公司有揭露）；其餘頁照樣板：README → In → Q → Thesis → Seg → PL → CF → Fund → RPO → Sites → Debt → Lease → Val → RevDCF → Recon → Sens → Sources。
- **規格（全部必讀）：** `docs/plan/個股模型原則_命題加營收分部.md`（§10 第 1–62 條；56–62 為本次新增）、`docs/plan/個股DCF標準規格.md`、`docs/plan/個股HTML版面規格.md`（v1.2，含可選節 `funding`、`revdcf` 與分區主題圖）、`docs/plan/個股模型驗收_不需驗收項目.md`（v1.7）、`docs/plan/自算產業β_20261105.md`。
- **前版：** `coreweave/dist/20261106_CoreWeave收支模型_v4_13.xlsx`（加權目標價 25.51、賣出；現價 43.66〔2026-11-05〕）。舊引擎（JS＋company.json）移 `coreweave/legacy_v4/`（照 Oracle 的 `legacy_v0.3/` 作法，保留不刪）。
- **可重用的資料登錄：** 分支 `origin/claude/coreweave-2026q3`（4 個 commit，11-11）的 `coreweave/company.json`（Q3 實績、Q4 隱含指引、2033 可轉債本金 4.2、期後募資 6.63）與 `coreweave/data/tokenomics_snapshot_v22.0.json`。逐筆對原文核對後使用；舊格式程式不要搬。

## 0. J 題決定（Andy 已確認，照做）
- **J1：** 直接重建為 v5.0（新格式），同版滾入 Q3。檔名 `coreweave/dist/20261112_CoreWeave個股模型_v5.0.xlsx`／`.html`（日期以成品日為準）。
- **J2：** 情境 A（樂觀）＝8 GW、B（基準）＝5.6 GW、C（保守）＝4.2 GW（沿用 v4.13 容量路徑；若 Q3 合約電力 4.6 GW 使 C 的路徑低於已簽約量，C 改為「已簽約量」並在報告寫明）；權重 25／50／25。橫幅寫明「A 情境擴張最多但目標價最低」時要依數值由公式組字（原則 57、60）。
- **J3：** DCF 照標準規格（FCFF，到 FY2032，WACC，永續 NOPAT×(1−g/RONIC)/(WACC−g)，RONIC＝WACC）；股數＝評價日 pro forma＋RSU／選擇權庫藏股法；未來瀑布新股不進 DCF 每股（X35）；稀釋在 Fund 頁呈現（原則 62）。v4.13 的「融資調整 DCF」值放 Recon，標「舊法，不進結論」。
- **J4：** 公司 Q3 法說宣稱的新約價格（4–5 年約每 MW 年收入 US$20–25m；短約 >US$40m；7 月起各 SKU 漲 25%＋10%）＝Interested-party：只進 Recon 對照、Sens 案例（每 MW 收入改用公司宣稱值的新約部分）、反解「支持現價所需每 MW 收入」；**不改基準 k**。另在 `coreweave/docs/reports/` 附對帳單草稿（主張、數值、單位、口徑〔分母是 IT 還是設施 MW 待查〕、出處、等級、立場、影響 Tokenomics 哪些名稱），chat 端會放進 Drive `對帳/`。
- **J5：** FY26 資本支出＝9M 實績（1H 16.139＋Q3 14.511＝30.650，Verified）＋Q4 依模型公式；公司指引 37–39 只進 Recon 與 Sens；不得有以指引作界值的 MAX／MIN（原則 59）。Q4 模型值與指引隱含 Q4（6.35–8.35）差距照原則 50 標示。

## 1. v4.13 驗收採用項（規劃書第三節 A1–A16、B1–B6）
- A1 DCF 50%＋Comparable 50%；兩法單獨目標價並列；Sens `w_dcf` 0.3／0.7。Comparable：新雲同業（Nebius、IREN、Oracle 等，依 Oracle 樣板的同業取法；估值日一致、NTM；同業 <8 家不迴歸）；相對倍數調整依 `tools/relative_multiple.py`（資料不足＝1 並寫明）。
- A3／A8 原則 56：終值占 EV、淨負債／EBITDA、利息保障倍數分母 ≤0 → N/M、檢查不通過；DCF 失效條件含 EV≤0。
- A4 原則 57：所有方向字樣公式組字；評等五級（無「賣出」）；邊界 >+50%、>+25%、≥−25%、≥−50%。
- A5／A13 WACC 橋：rf（美國 10 年期，標日期）、ERP 5%、自算產業 β（`docs/plan/自算產業β_20261105.md`，Blume、重加槓桿）、稅前債務成本（Q3 公司稱加權 7.9%＝Interested-party；與市場殖利率並列）、稅率、市值權重；g＝MIN(3%, 10 年期)。
- A6 原則 58：情境區間取三情境 MIN／MAX＋基準在區間內檢查。
- A2／A7 原則 59：公式值、採用值、差額、理由並列。
- A9 原則 61：IREN–Microsoft 錨標 Analogy 並寫差異。
- A10 原則 60：每增加 1 GW 的增量淨現值（A 對 B、B 對 C）。
- A11 Sens 租金高案例（v4.13 站點檢查：五期可能低估 3.29、FY30 每 MW 年租金 1.00）；風險登錄。
- A12 Sens：營運資金 0%／4%、股權上限 10%／30%、高息債 15%／18%。
- A14 情境 25／50／25；A15 市場參數日期（>2 交易日標逾期）；A16 `build_id`（`tools/model_governance`）。
- B1 舊融資調整 DCF 放 Recon；B2 原則 9、55 兩指標 WARN（4.5／3）＋原則 62 新股占市值逐年與累計；B3 長期現金稅率 21% 標 Assumed，Sens 25%；B4 NOL 80% 依法規，Section 382 列風險登錄（Unknown）。
- 也要套用原則 1–55 中適用者（條件式建議與失效條件、單一標記＋來源誘因、反解主導變數、兩個期間並列、三個日期分列、展示層一致性 51、收支關卡旁並列期間流動性 55 等）。新雲不適用自用 AI 讀法（52）。

## 2. 資料（先查證再登錄；每筆來源＋標記）
- **Q3 2026（11-10 公布）：** 營收 3.661；調整後 EBITDA 約 2.3（62%）；調整後營業利益 0.379；淨損 0.709；利息費用 0.846；資本支出 14.511；backlog 132.5、RPO 132.3（24 個月內 36%、25–48 個月 38%、48 個月後 26%）；在役 >1.7 GW、合約 4.6 GW；現金等約 6.5；FY26 指引營收 13.0–13.3、調整後營業利益 1.275–1.375、資本支出 37–39、年底 ARR 20–21；Q4 利息 1.05–1.15；加權債務成本 7.9%；未來三個月預收約 9。以公司新聞稿（investors.coreweave.com、BusinessWire）與 8-K 附件為準；法說逐字稿只作線索、全文不入 repo。
- **10-Q：** 雲端連不到 SEC，不得繞過。找不到的項目（債務到期表、租賃、股數）先沿用 Q2 10-Q 並標「待 Q3 10-Q」，在報告列資料缺口（分「找不到」與「不存在」）。
- **Tokenomics：** 快照 v22.0（`tools/tokenomics/import_tokenomics.py`；現行 master 以 Tokenomics repo `model/CURRENT` 為準）。
- **現價：** 最新收盤（11-11），寫明日期與來源；評價日＝2026-09-30（Q3 資產負債表）；目標價時點＝現價日＋12 個月。
- **Wattonomics：** Recon 並列 Wattonomics 對 CoreWeave 的上線 GW（版本＋儲存格）；差距 >15% 寫對帳單草稿。

## 3. HTML
- 依版面規格 v1.2：橫幅置頂（加權與 Comparable 並列）、命題分點＋回覆摺疊、資金缺口與融資一節放命題之後（`funding`）、反向 DCF（`revdcf`）、敏感度只顯示 ≥10%、附錄。
- 主題圖 3 張（v1.2）：(1) 每 MW 年收入：模型 vs Tokenomics 錨 vs 公司宣稱新約價格；(2) 資金缺口與融資來源 × 三情境；(3) backlog 轉營收路徑（RPO 分桶 vs 模型營收）。
- 單一自含檔、離線可開；`check_company_html.py` 通過；1280／390 寬無水平捲動。

## 4. 核對與成品
- `make_v5.sh`（照 `oracle/make_v05.sh`）：建置 → LibreOffice 重算 → 公式錯誤 0 → Sens 快照 → 收支恆等式 → `check_in_layout` → `check_sens_coverage --require-keycol` → `check_segments` → `check_governance`（含 `--display` 與 `coreweave/display_check.json`）→ HTML。
- 目標價變動拆解（`coreweave/scripts/decomp_v50.py`）：v4.13 25.51 → ① 方法改為新格式（同一 Q2 資料、TK v18.0、同現價）→ ② Q3 實績與日曆滾動 → ③ TK v22.0 → ④ 現價；四段相加＝總變動。① 若因方法差異過大無法同資料重現，改為「v4.13 → v5.0 基準」單段，並把 ② ③ ④ 在 v5.0 內逐一切換拆出，報告寫明。
- 報告 `coreweave/docs/reports/20261112_coreweave_v5.0_Oracle樣板重建與Q3.md`，非工程師看得懂：結論三句（加權目標價與評等前後、收支關卡、命題是否改變）；拆解表；J4 證據與反解；驗收 A／B 各項落實位置（工作表＋列）；資料缺口；判斷題附預設。
- README（`coreweave/README.md`）改寫為新格式說明；v4.13 成品移 `dist/_待刪除/`。

## 5. 規則
- 分段 commit＋push（每頁或每段一次），額度將盡時先 push 並在報告寫下一步。
- soffice 用獨立 `-env:UserInstallation`。
- PR 先開 draft（REST：`gh api repos/YinchenChang/company-models/pulls`），**不要合併**；chat 端安排未參與代理核對後合併。
- 不以 backlog、指引反推 k 或每 MW 價格；指引只作 Recon／Sens。
- html.json 文字不得含數字；白話文字避免四類錯誤（把假設寫成結論、把敏感度案例寫成門檻、方法描述不精確、時點或口徑混用）。

## 6. 舊守則的承接（`coreweave/CLAUDE.md` 已決定事項）
- **仍有效、要帶進 v5.0：** 第 1 條可移植（新雲後續 Nebius、WhiteFiber 照此重建）、第 2 條差異要能解釋、第 3 條精簡、第 10 條每季滾動、第 11 條 HTML 單一檔案離線開啟（`check_company_html.py`＋離線開啟檢查）、第 12 條目標價變動拆解、第 13 條營運成本以 Tokenomics 滿載穩態為主值、**第 14 條每 MW 收入以 Tokenomics `IF_HoldEcon` × 公司因素 k 為錨（不以營收、ARR、RPO 反推 k）**、**第 15 條公司實況驗證（差距 >10% 找有證據的機制，找不到則 TK 為基準、公司實際為 Sens）**、第 16(3) 條單一 k。
- **被取代：** 第 8 條雙引擎與 company.json 為唯一來源 → 改為 Oracle 樣板（Python 產生 xlsx、事實總帳 `data/*.json`、HTML 由共用產生器）；第 16(1) 條 DCF 權重 45% → 50%（X2）；第 12(1)(2) 目標價時點 → 現價日＋12 個月（X20、DCF 規格）；評等門檻「賣出」→ 五級。
- **CLAUDE.md 改寫：** v5.0 完成時把 `coreweave/CLAUDE.md` 改寫為新格式守則（參考 `oracle/README.md`），舊版移 `coreweave/legacy_v4/CLAUDE.md`。

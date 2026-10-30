# spacex/ — SpaceX（Nasdaq: SPCX，含 xAI）個股模型 v0.2.1（命題＋官方營收分部 → 損益表）

**命題（一句話）：** SI 分部的算力投入（9M26 SI 資本支出約 400 億美元、名目算力 1.7 GW；Q3 法說：Q4 資本支出 350–400 億、年底 2.5–3 GW）能否由 SI 營收（對外算力出租＋Grok／Cursor／X）回本；Connectivity（Starlink）現金流與 IPO 資金能支撐多久。

**核心問題（S2）：** FY2028 SI 每 GW 年營收 ÷ Tokenomics 每 GW 年持有成本（與 OpenAI PxQ 頁同口徑）；收支關卡＝FY2026–FY2030 自由現金流合計（扣融資租賃本金與以其他融資取得之 PP&E；Valor 失敗售後租回本金列融資、不扣；公司口徑 FCF 並列）。

**投資建議：強力放空**（v0.2.1；加權目標價 US$34.79，比現價 143.89〔2026-10-29 收盤〕低 75.8%；v0.2 為 34.83、v0.1 為 28.76）。主值＝自用 AI 讀法 2；讀法 1 目標價 22.57（−84.3%）。評等邊界：目標價 ≥ 71.95 才離開強力放空。市場共識目標價 244.97（MarketBeat 46 位）。

## v0.2.1 改了什麼（工作單 `docs/workorders/20261030_SpaceX_v0.2.1_一致性小修.md`；報告 `docs/reports/20261030_spacex_v0.2.1.md`）
- **以融資取得 PP&E 的算法統一（第 1 點）**：pace 的 SI 投資季速 `si_cx_rr`＝Q4 SI 現金資本支出 34.06 ÷ (1 − fin_sh 8.93%)＝37.40（＝Thesis Q4 26 SI 投資；融資比例隨建置規模放大），不再「＋9M 以融資取得 PP&E 季均 1.307」（v0.2 35.37）。目標價 +0.62。
- **pace 分母改 FY2027 新增世代組合（第 2 點）**：Thesis 新增 `cxgw27`＝am_27 × GB300 IT＋(1 − am_27) × VR200 IT＋自建 × 機房＝46.31（v0.2 用 GB300 43.64）；pace＝4 × 37.40 ÷ 46.31＝3.23 GW／年（v0.2 3.24）。FY2027 SI 投資＝4 × Q4 26 SI 投資＝149.61（差 0；check_v02 容差 1%）。目標價 −0.66。
- **In 頁 pace 低／高（第 3 點）**：採「build 時由同一函式產生、check 比對」——`build_spacex.py` 的 `pace_of(cx)` 同時產生基準參考值與低／高（Q4 指引 35／40 → 2.99／3.47），`check_v02.py` 以重算後模型的 q4cash、fin_sh、cxgw27 同式重算低／基準／高三欄比對（容差 5e-4）。不改成 Excel 公式的理由：`sens_v02.py` 從未重算的建置檔讀 D／F 欄數值，公式會讀成字串；且 D／F 欄公式須引用 cx_q4 的 D／F 欄，觸及原則第 8 節「主模型不讀敏感度欄」的檢查。
- **G2 記錄 Andy 決定**：2027Q1 營收指引只校準、不硬套——Andy 2026-10-30 17:15 回覆「同意，請繼續」（原為預設）。G2 校準 k0 因 FY2027 營收隨 pace 變動而重解：4.8197 → 4.8227（`data/g2cal_v021_20261030.json`）。
- 腳本沿用 v02 名稱：`make_v02.sh` 以 `VER`（預設 0.2.1）決定成品與敏感度檔名，`G2CAL=1` 先重解 G2；`waterfall_v02.py` 以 `WF`（預設 v021＝v0.2 → v0.2.1 兩步；`WF=v02`＝v0.1 → v0.2）區分；`build_spacex.py` 的 `SPCX_STAGE` 新增 8、9。v0.2 成品、敏感度與 G2 校準檔移 `dist/_待刪除/`、`data/_待刪除/`。

## v0.2 改了什麼（工作單 `docs/workorders/20261030_SpaceX_v0.2_原文換版與法說指引.md`；報告 `docs/reports/20261030_spacex_v0.2.md`）
- **原文換版（G5）**：`facts_v02.py` → `data/spacex_facts_v02.json`（424B4 最終招股書 FY2025／FY2024、2026Q2 10-Q 附註、兩次法說的指引數字與 ≤10 字短句；逐字稿不入庫）。FY2025 分部對帳差額 0；Space FY2025 資本支出 3.832 原文確認（v0.1 判為媒體混用，有誤）；2024 年底名目算力 0.3（v0.1 判斷 0.15）。Q 頁新增 2025Q4（＝FY2025 − 9M25）。
- **法說指引入基準（G1、G3）**：Q4 26 合併資本支出 37.5（SI＝37.5 − Space − Connectivity 3.437＝34.06，隱含 SI 占 90.8%）；2026 年底名目算力 2.75 GW；FY2027–28 pace＝(34.06＋以融資取得 PP&E 季均 1.307) × 4 ÷ 每 GW 投資 43.64＝3.24 GW／年（二修：與 Q4 同一拆法，不再用 9M 占比 83.2%；v0.1 1.63）。
- **只作檢查或 Sens（G2、G3、G4、G6、G7）**：2027Q1 營收指引 33–37 與 12 月 ARR 1,000 億列 Recon 第 5 節（模型 2027Q1 約 20.8，低於下緣；G2 不硬套指引：Andy 2026-10-30 同意）；Musk 2027 年底 13 GW 列 Sens「管理層路徑」；k0 以 2026Q2 校準列 Sens；Cursor 法說口徑與 backlog 列 Recon。 二修新增 Sens「G2 校準：k0 使 2027Q1 營收＝35」（k0 4.82）與「G2 校準＋k 不收斂」。
- **v0.1 遺留小修**：(a) 分部加總讀法 2 刪除無對應扣除的自用 AI 營運成本加回（−0.51／股；讀法 1 改扣）；(b) FY2027 年底淨現金只扣融資租賃年底餘額（+0.01）。
- **目標價歸因**（`waterfall_v02.py` → `data/waterfall_v02_20261030.json`）：28.76 →(a) −0.51 →(b) +0.01 → G5 −0.02 → G1 資本支出 −1.81 → G1 年底 GW +3.01 → G3 pace（9M 占比）+4.51 → pace 一致性（Q4 拆法）+0.88 → 34.83。

## 重點觀察（數字取自成品 Excel）
1. 目標價 34.79＝DCF 23.34 與分部加總 46.23 各半；情境 A／B／C 為 53.79／33.39／18.58。現價隱含 NTM EV／調整後 EBITDA 約 18.6 倍，本模型分部加總隱含 6.7 倍（情境 B）。FY2027 營收 106.2（v0.2 106.3、v0.1 86.2）。
2. S2 比值：2026Q3 校準 k0＝1.87；FY2028 A／B／C＝1.73／1.56／1.43（k 假設路徑內插，v0.2.1 未改）。名目算力 2026 年底 2.75、FY2028 年底 9.21 GW；SI 投資 FY2027 149.6（＝4 × Q4 26 SI 投資 37.40）、FY2028 158.3。
3. 收支關卡 A／B／C＝−176.6／−248.5／−300.9（v0.2 B −249.7、v0.1 B −84.1）；公司口徑 −121.0／−192.9／−245.3；五年新增舉債 B 138.2（FY2027 61.6）。指引的資本支出步調使 IPO 現金在 FY2027 用盡。

## 結構（三層）
| 層 | 內容 | 頁 |
|---|---|---|
| 命題模組 | SI 名目算力（IT 口徑 GW，評價日 1.7；2026 年底 2.75 指引中值）＋每年新增 pace＝SI 投資季速（Q4 SI 現金資本支出 34.06〔37.5 − Space − Connectivity〕÷ (1 − fin_sh)＝Q4 SI 投資 37.40）× 4 ÷ FY2027 新增世代組合每 GW 投資（GB300／VR200 IT 各半＋自建 90% × 機房 46.31）＝3.23 GW／年（FY2027–28，FY2029 起收斂）；對外 GW＝× ext_sh 72%（Customer B 衍生）；每 GW 年營收＝k × 機隊持有成本（k 由 k0 1.87 線性收斂到 FY2030 情境值）；GW → SI 投資 → AI 折舊；自用 GW → 讀法 2 影子效益 | Thesis、Seg |
| 營收分部 | **Space**（Launch Services＝客戶發射 × 每次收入；Launch & Development）、**Connectivity**（Consumer＝平均訂戶 × ARPU × uplift 1.178；Enterprise & Government）、**SI**（Advertising；SI Solutions & Infrastructure＝SI 對外雲端服務＋Grok／Cursor／X 訂閱、資料授權與 API 兩條命題子列）。**情境軸＝SI 的 k（A 1.6／B 1.25／C 1.0）與 Starlink FY2027 淨增（7.0／5.5／4.0 百萬）** | Seg |
| 損益表 | 分部調整後 EBITDA（Space 利益率 −5.9% → 20%；Connectivity 63.9%；SI＝對外出租營收 − Tokenomics 營運費用＋SI 其他 × 16.9% − 自用增量營運費用）− 折舊 − 股權報酬（含 Cursor 收購後服務期 22.2，直線到 2029-06）→ 分部營業利益 → 合併 → 利息、稅（明確期 12%）→ EPS | PL |

之後：CF（收支關卡口徑與公司口徑；Valor 失敗售後租回本金列融資；EchoStar 頻譜按成本、價值中性）、Val（DCF 50%＋分部加總 50%；情境 25／50／25；SOTP：Connectivity 電信同業 7.13×、SI 其他 EV／NTM 營收 5×、SI 對外出租自身 DCF、Space EV／營收 3×；明確期稅率 12%、終值期 21%）、Recon（v0.2 新增第 5 節：法說指引校準檢查 G2 與對照 G3、G6、G7）、Sens（92 案整本重算）、Sources。

**敏感度（92 案，目標價 16.24–247.14；v0.2.1 重跑）**：G2 校準＋k 不收斂 247.14（+610.5%，強力買進）與 IF_HoldEcon 四世代高成本 87.15（+150.5%，放空）是僅有的兩案離開強力放空；其次 k 不收斂 +95.1%、Tokenomics 一致高成本 +88.6%、IF_HoldEcon 低成本 −53.3%、G2 校準（k0 4.82）+40.2%、讀法 1 −35.1%、管理層路徑 13 GW +32.9%；k0 以 2026Q2 校準 −15.1%。設計特性：k0 隨持有成本重新校準、FY2030 k 為固定倍數，故持有成本同時推動營收。FY2028 S2 比值是 k 假設路徑的內插，不是模型產出。

## 已知限制
- **SOTP 的 NTM 含 FY2027 11–12 月**（v0.2 獨立核對非阻擋建議 3；v0.2.1 只註明、不改）：目標時點 2027-10-30，分部加總的 NTM 指標＝FY2027 × w_ntm27（62 ÷ 365，即 2027 年 11–12 月）＋FY2028 ×(1 − w_ntm27)，以年度值按天數加權近似，未拆季度路徑。
- FY2027 SI 投資＝4 × Q4 26 SI 投資只在基準成立；管理層路徑案（mgmt_gw27＝1）FY2027 新增改為 13 − 2026 年底，不適用。FY2028 新增全為 VR200 錨，FY2028 SI 投資 ÷ pace 高於 cxgw27（每 GW 投資隨世代上升，屬設計）。

## 驅動 → 推導對照表
| 頁 | 輸入（In 頁鍵） | 推導量 | 因果方向 |
|---|---|---|---|
| Thesis | gw_0、si_cx_rr → pace、ext_sh、mix0_*、selfb、Tokenomics IF_HoldEcon／OpexGW／CapexIT／Facility | SI 名目算力、對外／自用 GW、SI 投資、AI 折舊、自用增量成本、影子效益 | SI 投資 ÷ 每 GW 投資 → GW；GW → 投資 → 折舊 → SI 營業利益 |
| Seg | k30_A／B／C、conv_k、g27_nr、gl_nr、gl_ads、nl_*、g_rpl、gl_ld、nadd_A／B／C、g_arpu、gl_eg | 三分部、六條官方子列、兩條 SI 命題子列、合併營收 | 對外 GW × k × 持有成本；發射 × 每次收入；訂戶 × ARPU |
| PL | m_spl、m_cn、m_so、sbcr_*、sbc_cur_left、da_*、tax | 分部調整後 EBITDA、分部營業利益、淨利、EPS | 分部營收 → EBITDA − 折舊 − SBC → 營業利益 → EPS |
| CF | wc_k、cx_cn、cx_spl、fin_sh、mat_*、lease_*、spec_*、min_cash | 兩種口徑 FCF、新債、淨現金、收支關卡 | 淨利＋折舊＋SBC − 資本支出 → FCF → 舉債 |
| Val | rf、beta、erp、kd、g_tv、ronic、同業倍數、evs_space、tp_w*、w_dcf | DCF、分部加總、目標價、評等 | 無槓桿 FCF 折現＋分部 EBITDA × 倍數 → 情境加權 → 50／50 |
| Sens | In 頁敏感度低／高 | 加權目標價變動 | 一次改一個輸入 |

## 口徑與資料說明
- **名目算力（nameplate compute）**：10-Q 定義為安裝 GPU × all-in power，排除冷卻與設施負載 → IT 口徑，與 Tokenomics Interface 第 1 列 IT 關鍵電力同口徑，不乘 PUE；含對外出租、推定不含第三方雲（S6）。設施口徑讀法只列 Recon 對照。
- **歷史欄**：Q 頁 2025Q2–2026Q3、9M25、9M26、FY2025（424B4 原文，比較期依共同控制追溯重編）、2025Q4（衍生）。成長起點仍用 9M 年增：近四季年增需 2024Q4（9M2024 未揭露），FY2025 的 Connectivity、SI 子列 424B4 亦未揭露（G5 無法改用近四季）。
- **Cursor**（2026-08-14 併入，Q3 只含約 1.5 月）：預測以備考全季為起點（Q4 26 加 0.539）。
- **同業樣本小**：Connectivity 五家（MarketScreener 2026-10-08，FY2026 口徑）；SI 三家（CRWV、NBIS FY2027、ORCL NTM 2026-10-07，口徑不一）；Space 無直接上市同業（EV／營收判斷）。
- **敏感度欄只用於 Sens 頁**；主模型公式不讀 In 頁 D／F 欄（check_in_layout 無警告）。
- **Tesla 合併評估**（董事會特別委員會）不建模，列事件風險；EchoStar 交割發行約 261.8M 股與績效獎酬等稀釋只在 Sens（dil_x）。

## 檔案
| 檔案 | 內容 |
|---|---|
| `build_spacex.py` | Python → Excel 產生器（openpyxl，全部活公式；藍字輸入、黑字公式、黃底判斷題） |
| `facts_2026q3.py` | 10-Q 與新聞稿兩讀 → `data/spacex_facts_2026q3.json`（A 段；B 段依審查修正） |
| `facts_v02.py` | v0.2 事實增補（424B4、2026Q2 10-Q、法說指引；逐字稿不入庫，本機有逐字稿時重讀核對）→ `data/spacex_facts_v02.json` |
| `sens_v02.py` | 敏感度：一次改一個 In 頁輸入、整本 LibreOffice 重算 → `data/sens_v021_20261030.json`（v0.2.1；v0.1、v0.2 結果在 `data/_待刪除/`） |
| `waterfall_v02.py` | 目標價歸因瀑布：`WF=v021`（預設）v0.2 → v0.2.1（SPCX_STAGE 7–9）→ `data/waterfall_v021_20261030.json`；`WF=v02` v0.1 → v0.2（0–7）→ `data/waterfall_v02_20261030.json` |
| `g2cal_v02.py` | 二修：解 G2 校準 k0（情境 B 的 2027Q1 營收＝35）→ `data/g2cal_v021_20261030.json`（v0.2.1 重解 4.8227；只供 Sens 的 k0_ovr；v0.2 的 4.8197 在 `data/_待刪除/`） |
| `check_v02.py` | 成品核對：公式錯誤 0、核對列 0（分部對帳、子列、命題子列）、敏感度未過期、固定列標籤存在 |
| `gen_html_json.py` | 由列號對照產生 `html.json`（勿手改 html.json） |
| `make_v02.sh` | 一次跑完：事實總帳（v0.1＋v0.2）→（敏感度）→ 建置 → 重算 → check_v02 → check_in_layout → check_sens_coverage → check_segments → HTML 與檢核 |
| `segments.json` | 官方分部宣告（`tools/segment_check/check_segments.py spacex` 符合） |
| `dist/20261030_SpaceX個股模型_v0.2.1.xlsx`／`.html` | 成品（v0.1、v0.2 成品移 `dist/_待刪除/`） |
| `docs/reports/20261030_spacex_v0.2.1.md` | v0.2.1 報告；v0.2 報告 `docs/reports/20261030_spacex_v0.2.md`；v0.1 B 段報告 `docs/reports/20261030_spacex_v0.1.md`；A 段事實總帳報告 `docs/reports/20261030_spacex_v0.1_A段_事實總帳.md` |

## 建置
```
bash spacex/make_v02.sh            # 沿用已提交的敏感度結果（VER 預設 0.2.1）
SENS=1 bash spacex/make_v02.sh     # 改了任何輸入或公式時重跑敏感度（約 2 分鐘，JOBS＝2）
G2CAL=1 SENS=1 bash spacex/make_v02.sh   # 影響 FY2027 營收的改動（如 pace）：先重解 G2 校準 k0
python3 spacex/waterfall_v02.py /tmp/wf spacex/data/waterfall_v021_20261030.json   # 歸因瀑布 v0.2 → v0.2.1（WF=v02 為 v0.1 → v0.2）
```
**工作表**：README → In → Q → Thesis → Seg → PL → CF → Val → Recon → Sens → Sources。金額 US$B；FY2026＝9M 實績＋Q4 模型（欄 E＝9M26A、F＝Q4 26E）；明確期到 FY2032；評價日 2026-09-30；目標價時點 2027-10-30。

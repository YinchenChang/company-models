# spacex/ — SpaceX（Nasdaq: SPCX，含 xAI）個股模型 v0.1（命題＋官方營收分部 → 損益表）

**命題（一句話）：** SI 分部的算力投入（9M26 SI 資本支出約 400 億美元、名目算力 1.7 GW）能否由 SI 營收（對外算力出租＋Grok／Cursor／X）回本；Connectivity（Starlink）現金流與 IPO 資金能支撐多久。

**核心問題（S2）：** FY2028 SI 每 GW 年營收 ÷ Tokenomics 每 GW 年持有成本（與 OpenAI PxQ 頁同口徑）；收支關卡＝FY2026–FY2030 自由現金流合計（扣融資租賃本金與以其他融資取得之 PP&E；Valor 失敗售後租回本金列融資、不扣〔審查第 2 條 10-30 更正〕；公司口徑 FCF 並列）。

**投資建議：強力放空**（v0.1，10-30 二修後；加權目標價 US$28.76，比現價 143.89〔2026-10-29 收盤，MarketBeat、Equibles 兩來源一致〕低 80.0%）。主值＝自用 AI 讀法 2；讀法 1 目標價 21.88（−84.8%）。評等邊界：目標價 ≥ 71.95 才離開強力放空。市場共識目標價 244.97（MarketBeat 46 位）。

## 重點觀察（數字取自成品 Excel）
1. 目標價 28.76＝DCF 18.81 與分部加總 38.71 各半；情境 A／B／C 為 39.58／27.99／19.47。現價隱含 NTM EV／調整後 EBITDA 約 26.5 倍，本模型分部加總隱含 7.2 倍（情境 B）——差距主要在評價口徑（SI 對外出租用 Tokenomics 持有成本錨定的自身 DCF、Connectivity 用電信同業 7.13 倍），不在營收：本模型 FY2027 營收 86.2，比 6 月的過期共識 68.3 高 26%。
2. S2 比值：2026Q3 校準 k0＝1.87（Q3 年化 SI S&I 28.1 ÷ 平均在役 1.55 GW＝18.1 US$B/GW-年，÷ 機隊持有成本 9.71）；FY2028 A／B／C＝1.73／1.56／1.43（情境 B 每 GW 年營收 17.0、持有成本 10.9）。名目算力 2026 年底 2.11 GW（公司稱逾 2 GW）、FY2028 年底 5.36 GW；SI 投資 FY2027 75.4、FY2028 79.7。
3. 收支關卡（收支關卡口徑，10-30 更正後）A／B／C＝−42.7／−84.1／−114.6；公司口徑 FCF −12.2／−53.7／−84.1。IPO 後現金 89.9 撐到 FY2028，情境 B 需新債 13.1。

## 結構（三層）
| 層 | 內容 | 頁 |
|---|---|---|
| 命題模組 | SI 名目算力（IT 口徑 GW，評價日 1.7）＋每年新增 pace＝SI 投資季速（2026Q3 SI 資本支出 16.444＋以其他融資取得 PP&E 季均 1.307）× 4 ÷ Tokenomics 每 GW 投資（GB300 IT＋自建 90% × 機房 43.64）＝1.63 GW／年（FY2027–28，FY2029 起收斂）；對外 GW＝× ext_sh 72%（Customer B 衍生）；每 GW 年營收＝k × 機隊持有成本（k 由 k0 1.87 線性收斂到 FY2030 情境值）；GW → SI 投資 → AI 折舊；自用 GW → 讀法 2 影子效益 | Thesis、Seg |
| 營收分部 | **Space**（Launch Services＝客戶發射 × 每次收入；Launch & Development）、**Connectivity**（Consumer＝平均訂戶 × ARPU × uplift 1.178；Enterprise & Government）、**SI**（Advertising；SI Solutions & Infrastructure＝SI 對外雲端服務＋Grok／Cursor／X 訂閱、資料授權與 API 兩條命題子列）。**情境軸＝SI 的 k（A 1.6／B 1.25／C 1.0）與 Starlink FY2027 淨增（7.0／5.5／4.0 百萬）** | Seg |
| 損益表 | 分部調整後 EBITDA（Space 利益率 −5.9% → 20%；Connectivity 63.9%；SI＝對外出租營收 − Tokenomics 營運費用＋SI 其他 × 16.9% − 自用增量營運費用）− 折舊 − 股權報酬（含 Cursor 收購後服務期 22.2，直線到 2029-06）→ 分部營業利益 → 合併 → 利息、稅（明確期 12%）→ EPS | PL |

之後：CF（收支關卡口徑與公司口徑；Valor 失敗售後租回本金列融資；EchoStar 頻譜按成本、價值中性）、Val（DCF 50%＋分部加總 50%；情境 25／50／25；SOTP：Connectivity 電信同業 7.13×、SI 其他 EV／NTM 營收 5×、SI 對外出租自身 DCF、Space EV／營收 3×；明確期稅率 12%、終值期 21%）、Recon、Sens（88 案整本重算）、Sources。

**敏感度（88 案，目標價 17.92–57.28，全部強力放空）**：最大單一驅動是 Tokenomics 每 GW 持有成本 IF_HoldEcon（四世代同時高／低成本 +99.2%／−37.7%）；其次 k 不收斂 +63.3%、Tokenomics 一致高／低成本 +58.4%／−12.2%、Connectivity 同業上緣 16.66× +28.1%、讀法 1 −23.9%。設計特性：k0 隨持有成本重新校準、FY2030 k 為固定倍數，故持有成本同時推動營收（持有成本調高 → 每 GW 租金上升）。FY2028 S2 比值是 k 假設路徑的內插，不是模型產出。

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
- **歷史欄**：Q 頁 2025Q2–2026Q3、9M25、9M26；FY2025* 為媒體轉述 S-1（二手），Space 資本支出未取得（A 段審查第 1 條）。成長起點用 9M 年增（近四季需 2025Q4，只有二手）。
- **Cursor**（2026-08-14 併入，Q3 只含約 1.5 月）：預測以備考全季為起點（Q4 26 加 0.539）。
- **同業樣本小**：Connectivity 五家（MarketScreener 2026-10-08，FY2026 口徑）；SI 三家（CRWV、NBIS FY2027、ORCL NTM 2026-10-07，口徑不一）；Space 無直接上市同業（EV／營收判斷）。
- **敏感度欄只用於 Sens 頁**；主模型公式不讀 In 頁 D／F 欄（check_in_layout 無警告）。
- **Tesla 合併評估**（董事會特別委員會）不建模，列事件風險；EchoStar 交割發行約 261.8M 股與績效獎酬等稀釋只在 Sens（dil_x）。

## 檔案
| 檔案 | 內容 |
|---|---|
| `build_spacex.py` | Python → Excel 產生器（openpyxl，全部活公式；藍字輸入、黑字公式、黃底判斷題） |
| `facts_2026q3.py` | 10-Q 與新聞稿兩讀 → `data/spacex_facts_2026q3.json`（A 段；B 段依審查修正） |
| `sens_v01.py` | 敏感度：一次改一個 In 頁輸入、整本 LibreOffice 重算 → `data/sens_v01_20261030.json` |
| `check_v01.py` | 成品核對：公式錯誤 0、核對列 0（分部對帳、子列、命題子列）、敏感度未過期、固定列標籤存在 |
| `gen_html_json.py` | 由列號對照產生 `html.json`（勿手改 html.json） |
| `make_v01.sh` | 一次跑完：事實總帳 →（敏感度）→ 建置 → 重算 → check_v01 → check_in_layout → check_segments → HTML 與檢核 |
| `segments.json` | 官方分部宣告（`tools/segment_check/check_segments.py spacex` 符合） |
| `dist/20261030_SpaceX個股模型_v0.1.xlsx`／`.html` | 成品 |
| `docs/reports/20261030_spacex_v0.1.md` | B 段報告；A 段事實總帳報告 `docs/reports/20261030_spacex_v0.1_A段_事實總帳.md` |

## 建置
```
bash spacex/make_v01.sh            # 沿用已提交的敏感度結果
SENS=1 bash spacex/make_v01.sh     # 改了任何輸入或公式時重跑敏感度（約 2 分鐘，JOBS＝2）
```
**工作表**：README → In → Q → Thesis → Seg → PL → CF → Val → Recon → Sens → Sources。金額 US$B；FY2026＝9M 實績＋Q4 模型（欄 E＝9M26A、F＝Q4 26E）；明確期到 FY2032；評價日 2026-09-30；目標價時點 2027-10-30。

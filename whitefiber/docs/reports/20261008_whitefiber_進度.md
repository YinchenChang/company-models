# WhiteFiber v0.1 分段建置進度（接手用；三張工作單共用）

## 目前狀態（每次 push 前覆寫）
- 已完成：**v0.1a 全部完成**：步驟 0（cfdc8c0）、1（5923076）、2–3（dfbbb79）、4（dc02b7d）、5（c74f2be）、6–7（857e793）、8a（c130c77）、8b 資料報告與 PR 回報（本 commit）。產出：`data/whitefiber_facts_20261008.json`（303 筆）、`data/consensus_wyfi_20261008.json`、`data/permw_tokenomics_20261008.json`、company.json `whitefiber` 區段、`docs/reports/20261008_whitefiber_v0.1a_資料.md`。verify.sh 19 項、`--vs-dist` 22 項全過。
- 下一步（v0.1b 起點，詳見資料報告第 2 節）：(1) 步驟 1 欄位搬移：依 `company.json → whitefiber.*.mapTo` 換 meta（WhiteFiber／WYFI）、calendar（12 月財年、FY26Q2、首期 0.5 年）、ytdActual（1H26）、historicalPL、latestQuarter、debt（可轉債 cv31 剩 31.85＋cv32 310、DDTL、冰島、RBC 聯貸；建議以期後資本結構建模並寫明）、leases、rpo、valuation（16.83、45.12M 股）、peers、quarterly、meta.consensusFile＝data/consensus_wyfi_20261008.json；刪 oracle 區段與 Oracle／CRWV 資料檔。(2) 雲端 MW 主軸：期初可計費 MW 以經常性雲端年化 46.0（扣終止費）÷ 每 MW 校準。(3) 新建託管分部（NC-1、MTL-1／3 照合約；擴建用租金 1.85、建置 11.5／MW-IT、建物 20 年）。(4) 預付期初 143.1。(5) 評價：beta 2.5、kd 9.5%、託管 20.7×／雲端 6×。
- 未解問題（給 chat 端）：(a) 託管同業 NTM EV/EBITDA 只取得 DLR、EQIX 兩家（未達「至少 4 家」；APLD／CORZ／IREN／CIFR 無可用 NTM EBITDA，StockAnalysis 被環境代理阻擋）；(b) 個股 beta 0.84–4.77 不可靠，預設 2.5；(c) 2032 可轉債與 DDTL 追加屬期後事件，v0.1b 是否以期後結構建模待定（建議是）；(d) 一次性 GPU 租賃成本 4.0 只見於二手法說摘要；(e) EBITDA 共識找不到。

## 工作單總覽
| 工作單 | 分支 | PR | 狀態 |
|---|---|---|---|
| v0.1a 資料蒐集 | `claude/whitefiber-v0.1` | #23（draft） | 完成（待 chat 端審查） |
| v0.1b 模型改寫 | `claude/whitefiber-v0.1` | 同一 PR | 未開始 |
| v0.1c 驗證與成品 | `claude/whitefiber-v0.1` | 同一 PR | 未開始 |

<!-- 各工作單在下方新增自己的段落：「## v0.1x」＋步驟紀錄表（步驟｜狀態｜commit｜備註） -->

## v0.1a 資料蒐集

- 分支：`claude/whitefiber-v0.1`（自 af4af27）；PR #23。

### 步驟紀錄
| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 基線 verify、進度檔 | 完成 | cfdc8c0 | verify.sh 19 項全過（約 2 分 38 秒；cmp31 每情境 398 項）。soffice 24.2、openpyxl、playwright 1.56.0 已在環境；補裝 Playwright Chromium（`python3 -m playwright install chromium`）。複製無遺漏（company.json 引用的 data 檔齊全）。 |
| 1 SEC 一手文件 | 完成 | 5923076 | EDGAR 直接取得（CIK 0002042022，curl＋User-Agent）：FY2025 10-K、2025 Q2／Q3 與 2026 Q1／Q2 10-Q、8-K（2031 可轉債 2026-01、Nscale 2025-12、巴黎 2026-05、DDTL 2026-05、NC-2／3 2026-08、2032 可轉債 2026-08、9 月簡報）、13D/A（2026-08-25）。2026 Q1／Q2 財報新聞稿未以 8-K 提交，改用 PR Newswire 轉載（finanznachrichten）；法說用 MarketBeat 二手摘要。群組 1 共 204 筆（Verified 34／Interested-party 145／Derived 25）。重點更正：(1) 工作單「設備與過橋融資約 $83.2M」＝定期借款帳面合計（DDTL 30＋B. Riley 20＋冰島 18＋RBC 17.3，本金 85.3）；(2) 另有 2026-01 發行的 2031 可轉債 $230M（4.5%、轉換價 25.91），8 月以現金 118.5＋6.3M 股交換 198.15，剩 31.85；(3) Bit Digital 持股交換後降為 59.9%（13D/A）；(4) Q2 扣一次性後調整後 EBITDA 約 −0.6。 |
| 2 站點與 MW | 完成 | dfbbb79 | 群組 2 共 15 筆。公司簡報明載「All MW reflected on a gross basis」，IT 另列：MTL-1 4 毛／3 IT、MTL-2 5／3（2026 年底）、MTL-3 7／5（Cerebras，PUE 1.3 自述）、NC-1 99 毛（Duke 協議，2029-05 前）／第一期 54 毛已供電／40 IT（Nscale）→ IT／毛 0.741；冰島 6 毛／5.5 IT（承租）；雪梨 2.5 IT（2026 Q4）；亞特蘭大、安大略、巴黎 MW 未揭露。候選 60 MW 站點＝NC-2／NC-3（$60M，最多 198 MW 毛，2027 Q3 RFS）；Krambu 100 MW 只見於法說二手摘要（口徑不明）。 |
| 3 合約 | 完成 | dfbbb79 | 群組 3 共 24 筆。每 MW-IT 年營收（[Derived]，每顆 IT kW 取 Tokenomics v5.26 每 GW GPU 數倒數；B300／B200 以 GB300／GB200 近似）：Baseten 19.3、Prime Intellect VR200 18.3、冰島 B300 14.8、128 顆 B300 20.3、B200 384 顆 11.8、GB200 216 顆 15.8、H100 2,048 顆 8.8、Initial Customer H100 17.5；託管：NC-1 平均 2.16（第 1 年約 1.89）、MTL-3 2.35、MTL-1 約 2.38。巴黎、Hyperbolic GPU 數未揭露。最大客戶占 1H26 營收 63%。 |
| 4 託管租金與建置成本 | 完成 | dc02b7d | 群組 4 共 21 筆。託管租金（USD M／MW-IT·年，[Derived] 自各公司自述）：CORZ–CoreWeave 1.44（租戶出資改建，下緣）、CIFR–AWS 約 1.59（只揭露毛 MW，IT 以 1.3 換算）、CIFR–Fluidstack 1.79、APLD–CoreWeave 1.83、GLXY–CoreWeave ≥1.90、WhiteFiber NC-1 2.16、MTL-3 2.35 → 區間 [Analogy] 1.45／1.85／2.35。建置成本（不含 GPU，USD M／MW-IT）：WhiteFiber 改建 8–10／毛 MW（≈10.8–13.5／IT）、APLD PF2 10.0、Helios 第一期 13.2、NC-1 粗估 11.8、Tokenomics 10.2／12.7／16.2 → 區間 9.0／11.5／13.5。建物折舊：10-K 30 年 vs 10-Q 20–25 年（矛盾）。 |
| 5 每 MW（雲端） | 完成 | c74f2be | Tokenomics master 70d859e 的 `model/CURRENT` 已由 v5.24 更新為 v5.26；以 openpyxl 讀 v5.26 Interface!I7:N15，與 v5.24 逐格相同（v5.25／v5.26 只改 SRC、L1 第 36 列與新增 I 節名稱）→ 不重算，沿用 11.62／17.40／24.20。`data/permw_tokenomics_20261008.json` 新增 whitefiber 區段：(a) Q2 經常性雲端年化 ÷ 冰島 5.5 MW＝8.37（上限）；(b) 新簽 B300／VR200 長約 14.8–20.3（平均 18.2，接近基準 17.40）；(c) 毛利率 59.4% → 扣一次性 53.4%，調整後 EBITDA 率 19.2% → 扣一次性 −3.4%（IREN 35%、CRWV 59%）。三情境未用公司目標或合約單價作輸入。 |
| 6 評價輸入 | 完成 | 857e793 | 群組 6 共 24 筆。beta：eToro 4.76、FFFinstill 4.77（52 週，可能同法）、Snowball 0.84 → 個股 beta 不可靠（上市 14 個月）；預設 2.5（[Analogy]，NBIS 1.46–CIFR 3.45 之間偏保守，區間 1.5–3.5）。10 年期公債 5.27%（10/06）。CAPM 股權成本 17.8%。債務成本：2031 可轉債有效 5.37%、2032 票面 5.0%、DDTL 9.5%（→8.0%，MOIC 1.1×）、冰島 7.92%（有效 10.64%）；預設純債 9.5%。同業 NTM EV/EBITDA（[Derived]，EV ÷ 時間加權共識 EBITDA）：DLR 20.4×、EQIX 21.1×；APLD、CORZ、IREN、CIFR 無可用 NTM EBITDA（只有 EV/Sales 或 EBITDA 為負）→ 只有 2 家，列資料缺口；預設託管倍數 20.7×（下緣 15× 敏感度）、雲端 6×。CRWV、NBIS 沿用 10-07 資料（StockAnalysis 被環境阻擋）。 |
| 7 市場共識 | 完成 | 857e793 | `data/consensus_wyfi_20261008.json`：收盤 16.83（10/07，Yahoo／Benzinga）；市值 759（交換後 45.12M 股，[Derived]）；目標價 MarketBeat 38.00（32–50，11 家，9 買／1 持有／1 賣）、Investing 40.50（10 家）、Yahoo 40.45、S&P（Simply Wall St）40.89（8 家）；營收共識 S&P FY26 137／FY27 296／FY28 530（10／10／3 家）、淨利 −43／3／56；Tickflow FY26 138（7 家）；EBITDA 共識找不到。 |
| 8a company.json、fields_doc、verify | 完成 | c130c77 | company.json 新增頂層 `whitefiber`（20 子區：calendar、ytdActual、historicalPL、quarters、latestQuarter、onetime、balance、debt、shares、leases、rpo、mw、contracts、perMw、colo、prepay、valuation、consensus、guidance；每欄 value＋unit＋ref＋mapTo，ref 全部存在於總帳）；`scripts/fields_doc.py` 補 whitefiber 子區說明並 `--write`（README 欄位表）；verify.sh 19 項全過（引擎未動）。 |
| 8b 資料報告、verify --vs-dist、PR 回報 | 完成 | （本 commit） | `docs/reports/20261008_whitefiber_v0.1a_資料.md`（結論、v0.1b 起點、核心數字表、矛盾 12 項、每 MW 與託管租金對照、資料缺口、已套用預設 11 項、verify 輸出、總帳表格版 303 筆）；`verify.sh --vs-dist` 22 項全過（與 dist/ 0 差異）；PR #23 留言回報。 |

### 已套用的預設（判斷類；問題｜採用的預設｜替代選項｜影響方向）
| # | 問題 | 採用的預設 | 替代選項 | 影響方向 |
|---|---|---|---|---|
| 1 | 個股 beta 0.84–4.77 不可靠 | 2.5 [Analogy]（NBIS 1.46–CIFR 3.45 之間偏保守） | 4.77 | 4.77 → 股權成本 29.1%（+11.3pt）；1.5 → 12.8% |
| 2 | 稅前債務成本 | 9.5%（DDTL 純債） | 可轉債票面 5.0% | 5.0% 使 WACC 低約 1–2pt |
| 3 | 託管倍數只有 2 家 NTM | DLR／EQIX 中位 20.7× | 15× | 15× 使託管腿 EV −28% |
| 4 | 未簽約擴建託管租金 | 區間中點 1.85／MW-IT·年 | NC-1 平均 2.16 | 2.16 → 擴建營收 +17% |
| 5 | 建置成本 | 11.5／MW-IT（9.0–13.5） | 公司自述 8–10／毛 MW | 9.0 → 擴建資本支出 −22% |
| 6 | B300／B200 每顆 kW | GB300／GB200 近似 | HGX 約 1.8 kW | 只影響對照列 0–12% |
| 7 | 「256 H100 GPU servers」 | 2,048 顆（8 卡／台） | 256 顆 | 只影響對照列 |
| 8 | 一次性 GPU 租賃成本 | 4.0（二手） | 0 | 經常性毛利率 53.4% vs 29.2% |
| 9 | 市值股數 | 交換後 45.12M | 38.85M | 市值 759 vs 654 |
| 10 | 共識主來源 | 目標價 MarketBeat、營收 S&P | Investing／Yahoo | 目標價平均差約 6% |
| 11 | Tokenomics v5.26 | I7:N15 未變 → 沿用 | 重算 | 無 |

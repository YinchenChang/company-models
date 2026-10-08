# WhiteFiber v0.1 分段建置進度（接手用；三張工作單共用）

## 目前狀態（每次 push 前覆寫）
- 已完成：v0.1a 步驟 0（cfdc8c0）、1（5923076）、2–3（dfbbb79）、4（dc02b7d）、5 每 MW（permw_tokenomics_20261008.json＋群組 5）。
- 下一步：步驟 6 評價輸入（群組 6）→ 7 共識（consensus 檔＋群組 7）→ 8 company.json、verify、報告、PR 回報。產生器 `gen_facts.py` 的群組 5–7 草稿已寫好（在建置代理暫存區，未入庫；接手者可直接依本進度檔重做）。
- 未解問題：無。

## 工作單總覽
| 工作單 | 分支 | PR | 狀態 |
|---|---|---|---|
| v0.1a 資料蒐集 | `claude/whitefiber-v0.1` | #23（draft） | 進行中 |
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
| 5 每 MW（雲端） | 完成 | （本 commit） | Tokenomics master 70d859e 的 `model/CURRENT` 已由 v5.24 更新為 v5.26；以 openpyxl 讀 v5.26 Interface!I7:N15，與 v5.24 逐格相同（v5.25／v5.26 只改 SRC、L1 第 36 列與新增 I 節名稱）→ 不重算，沿用 11.62／17.40／24.20。`data/permw_tokenomics_20261008.json` 新增 whitefiber 區段：(a) Q2 經常性雲端年化 ÷ 冰島 5.5 MW＝8.37（上限）；(b) 新簽 B300／VR200 長約 14.8–20.3（平均 18.2，接近基準 17.40）；(c) 毛利率 59.4% → 扣一次性 53.4%，調整後 EBITDA 率 19.2% → 扣一次性 −3.4%（IREN 35%、CRWV 59%）。三情境未用公司目標或合約單價作輸入。 |

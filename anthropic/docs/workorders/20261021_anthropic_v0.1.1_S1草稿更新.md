# 工作單 Anthropic v0.1.1｜依 WSJ 轉述的 S-1 草稿與 Q3 投資人簡報更新（2026-10-21，chat 端）

- 來由：Andy 2026-10-21 轉貼 WSJ（引草擬 S-1）重點，「請依此訊息更新」。修補版（資料更新＋既有機制內的校準與對照），不拆段。
- 分支 `claude/anthropic-v0.1.1`（自 `main` `b4b04dc`；v0.1 已由 PR #22 合併）。完成後開 PR、轉 ready；**不得合併**（Project 規則：合併只在 Andy 明確指示時做）。
- 必讀：`anthropic/CLAUDE.md`、`docs/workorders/20261008_anthropic_共同規則.md`、規格 `docs/plan/20261008_Anthropic_v0.1_規格.md`（r1）、交接檔、v0.1 報告 `docs/reports/20261008_v0.1.md`。
- Tokenomics 快照維持 v5.26（下游同步依 Andy 2026-10-15 指示擱置）。
- 不得在任何對外請求（含 HTTP User-Agent）中使用使用者 email。超過 2 分鐘的命令背景執行，但必須等它跑完再回報（代理結束時背景程序會被終止）。

## 1. 新增 SRC_ANT 列（接續 SRC_ANT_380 起；每筆照共同規則第 4 節：網址必須實際開啟、note 附原文摘錄；全部 Interested-party，利害方 Anthropic（IPO 誘因），WSJ 為轉述者）
chat 端已初查的來源（請逐一開啟核對，打不開就換同一報導的其他轉載）：
- WSJ「Anthropic S-1 Shows Red-Hot Growth, Risks and Massive Computing Costs」（2026-10-20）的轉載：https://www.tradingkey.com/analysis/stocks/us-stocks/262227958-anthropic-ipo-filing-80-7-billion-h1-loss-tradingkey 、https://coincentral.com/anthropic-ipo-filing-reveals-80-billion-loss-and-600-billion-in-computing-bills/ 、https://runtimewire.com/article/anthropic-draft-s1-revenue-loss-compute-costs 、https://cryptobriefing.com/anthropic-ipo-draft-revenue-net-loss/ 、https://www.nationalheraldindia.com/business/anthropic-posts-807-billion-first-half-loss-largely-from-accounting-charges-report
- Q3 投資人簡報（2026-10-14）：https://www.implicator.ai/anthropic-tells-investors-q3-revenue-hit-16-8-billion-as-growth-slowed-to-45/ （其引 FT、CNBC、Euronews、Axios、Fortune；能開原報導就加列原報導）
需登錄的項目：
| 項目 | 值 |
|---|---|
| 2026 上半年營收 | 16.3（2025 上半年約 1.1） |
| 2026 上半年淨損 | 80.7，其中約 72 為可轉債等融資工具重估的非現金費用（2025 上半年淨損約 13） |
| 2026 上半年算力與基礎設施支出 | 近 9（2025 上半年 2.3；2025 全年約 7.3） |
| 2026 Q2 營收（實際） | 11.6（取代 v0.1 的「>11.5 初估」作為校準值；舊列保留） |
| 2026 Q3 營收（投資人簡報） | 16.8；調整後營業利益率 11%（第二個獲利季）；當季自由現金流為正；每 $1 算力支出 $1.70 營收 |
| 雲端運算付款義務 | 約 450，未來 7–10 年，多數不可取消 |
| 逐家 | Google 約 111、AWS 約 110、Microsoft 約 31、Nscale＋Lambda 最高 86（租用 NVIDIA 晶片，至 2033–2034）、Broadcom（Google 客製晶片）>160 |
| 已鎖定承諾 | >600；談判中最高約 150，全數成交可達約 750 |
| 客戶集中 | 兩家客戶約占 2026 上半年營收 20% |
| 供給限制 | 「算力供給是擴張最主要的限制」；雲端夥伴不願出售其所需的全部算力（WSJ 原文措辭，請找到後摘錄） |
| 年化營收（對照，Analogy） | SemiAnalysis 約 74（2026-10；**不可單獨引用**，須同列第二來源，找不到就只登為「單一來源，未採用」）；分析師估年底約 110 |

## 2. 模型變動（預設；報告「已套用的預設」逐項列出替代與影響）
| # | 項目 | 預設 | 替代 |
|---|---|---|---|
| U1 | 2026 半校準（D15 r2） | 2026 營收基準＝上半年實際 16.3＋Q3 實際 16.8＋Q4＝Q3（持平，保守；Decision）＝49.9 | Q4＝Q3 ×（Q3／Q2）；Q4＝最新年化 ÷ 4 |
| U2 | Nscale＋Lambda | 兩約合計錨定 WSJ「最高 86」：依兩約既有 GW 比例分攤金額，付款＝金額 ÷ 期間（至 2033–2034）；GW 依 D10 由金額回推並與既有揭露 GW 對照 | 維持各自公告值（45、35） |
| U3 | 承諾對帳 | 逐家對照改用 S-1 草稿（WSJ）數字；總額並列三個口徑：雲端付款義務約 450、已鎖定 >600、含談判中約 750；Checks WARN 門檻沿用 25% | — |
| U4 | 2026 算力支出對照（不校準） | 新增：上半年實際近 9、Q3 隱含＝Q3 營收 ÷ 1.70（Derived）；與模型 2026 合約實付（及含自建資本支出）對照，差距 >25% 立 WARN；報告說明 Fluidstack 自建資本支出 2026 年 16.7 是否與「算力與基礎設施支出」同口徑 | 校準到實際（不採：只有半年資料，且口徑未明） |
| U5 | D22 對照 | 營業利益率對照加 Q3 11%；量化模型 2026 非算力成本低估的隱含值（承 v0.1 A4-17） | — |
| U6 | 談判中 $150B | **只作情境**（基準仍只計已簽合約）：GW＝150 ÷ 期間（Assumed 7 年，5–10）÷ 合約價，2028 起爬坡（沿用 A3 爬坡預設）；報告列對命題與累計外部資金需求的影響。這是 v0.1 Andy 待審第 1 項「新增合約機制」的第一個有依據的情境 | 計入基準 |
| U7 | D21 | 2026 上半年非現金費用約 72 不進現金流；Funding 列示 | — |
| U8 | 2026 年中現金 | 若 S-1 轉述有 2026-06-30 現金，登 SRC 並在 Funding 加對照列（不改來源順序）；沒有就列資料缺口 | — |

## 3. 交付
1. builder 重建 `model/20261021_Anthropic_v0.1.1.xlsx`（舊版 `git mv` 到 `model/archive/`、更新 CURRENT）；LibreOffice 重算錯誤 0、`CHK_Errors`＝0；`python3 -m pytest tests -q` 全過（含 HTML 測試）；parity 情境補 U1、U2、U6 各一。
2. HTML 與 dist：`dist/20261021_Anthropic收支模型_v0_1_1.html`／`.xlsx`（舊 dist 檔 `git mv` 到 `dist/_待刪除/`）；HTML 加一段「v0.1.1 更新了什麼」（新數據 → 命題輸出前後值）。
3. 報告 `docs/reports/20261021_v0.1.1.md`＋`_對照.xlsx`：①新增 SRC 列；②**前後對照**：命題輸出（每 VR 等值 GW 差額、覆蓋率、累計外部資金需求、現金谷底、2026 營收）v0.1 → v0.1.1，並**逐項拆解**變動來自 U1／U2／…；③已套用的預設；④資料缺口（S-1 公開後待補）。結論先行、繁體中文、非工程語言。
4. CHANGELOG、交接檔（第 3、5 節）、進度檔；PR 留言第一行 `[Anthropic 回報] v0.1.1｜完成｜2026-10-21`；PR 轉 ready。每步 commit＋push。

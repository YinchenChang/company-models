# WhiteFiber v0.1 分段建置進度（接手用；三張工作單共用）

## 目前狀態（每次 push 前覆寫）
- 已完成：v0.1a 步驟 0（cfdc8c0）、步驟 1 SEC 一手文件（總帳群組 1，204 筆）。
- 下一步：v0.1a 步驟 2 站點與 MW、步驟 3 合約（總帳群組 2、3；產生器草稿已含，見步驟紀錄備註）。
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
| 0 基線 verify、進度檔 | 完成 | （本 commit） | verify.sh 19 項全過（約 2 分 38 秒；cmp31 每情境 398 項）。soffice 24.2、openpyxl、playwright 1.56.0 已在環境；補裝 Playwright Chromium（`python3 -m playwright install chromium`）。複製無遺漏（company.json 引用的 data 檔齊全）。 |
| 1 SEC 一手文件 | 完成 | （本 commit） | EDGAR 直接取得（CIK 0002042022，curl＋User-Agent）：FY2025 10-K、2025 Q2／Q3 與 2026 Q1／Q2 10-Q、8-K（2031 可轉債 2026-01、Nscale 2025-12、巴黎 2026-05、DDTL 2026-05、NC-2／3 2026-08、2032 可轉債 2026-08、9 月簡報）、13D/A（2026-08-25）。2026 Q1／Q2 財報新聞稿未以 8-K 提交，改用 PR Newswire 轉載（finanznachrichten）；法說用 MarketBeat 二手摘要。群組 1 共 204 筆（Verified 34／Interested-party 145／Derived 25）。重點更正：(1) 工作單「設備與過橋融資約 $83.2M」＝定期借款帳面合計（DDTL 30＋B. Riley 20＋冰島 18＋RBC 17.3，本金 85.3）；(2) 另有 2026-01 發行的 2031 可轉債 $230M（4.5%、轉換價 25.91），8 月以現金 118.5＋6.3M 股交換 198.15，剩 31.85；(3) Bit Digital 持股交換後降為 59.9%（13D/A）；(4) Q2 扣一次性後調整後 EBITDA 約 −0.6。 |

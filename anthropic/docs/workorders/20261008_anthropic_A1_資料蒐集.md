# 工作單 Anthropic A1｜資料蒐集（2026-10-08，chat 端）

- 目標：規格第 4 節 A–G 全部項目，寫成 `data/anthropic_src.yaml`（SRC_ANT 的唯一來源），並附研究筆記與資料報告。
- 分三組並行（各寫自己的檔，chat 端合併編號）：
  - **A1-R** 營收、價格、用戶（規格第 4 節 A、B、C）→ `data/a1_raw/revenue_pricing_users.yaml`＋`_notes.md`
  - **A1-C** 算力合約與算力支出（D）→ `data/a1_raw/compute.yaml`＋`_notes.md`
  - **A1-F** 成本、人數、融資、管理層目標（E、F、G）→ `data/a1_raw/finance.yaml`＋`_notes.md`
- 截止日：2026-10-08 可得的資料。先找最新，再回補 2024–2025 序列。

## 每筆格式（YAML list）
```yaml
- key: rev_runrate_2026_02        # 小寫英文 slug，唯一
  group: A                        # 規格第 4 節的分組字母
  metric: 年化營收（run-rate）     # 繁體中文
  value: 14.0                     # 數值；只有區間時 value 留 null
  lo: null
  hi: null
  unit: $B/年
  period: 2026-02                 # 數值所指期間（年、季、月或時點）
  source: Anthropic 新聞稿〈…〉    # 文件名
  url: https://...
  doc_date: 2026-02-12
  retrieved: 2026-10-08
  tag: Interested-party           # Verified／Interested-party／Analogy／Assumed／Derived
  party: Anthropic（募資誘因）      # Interested-party 必填：利害方與誘因
  basis: null                     # GW 類必填：IT／設施／未明；金額類：總額／淨額／未明
  note: 原文摘錄（≤40 字）與口徑說明
```
- 同一事實多個來源互相矛盾：全部登錄（key 加後綴 `_b`、`_c`），note 寫矛盾點。
- 找不到：另列 `not_found:` 區段（項目、試過的來源與查詢詞、判斷是「找不到」還是「可能不存在／未揭露」）。
- 不得杜撰；不得用模型記憶補數字——每筆必須有可開啟的網址與原文摘錄。SemiAnalysis 不可單獨引用。

## 判斷類預設
| 問題 | 預設 |
|---|---|
| 公司自述（新聞稿、高層訪談） | Interested-party（Anthropic：募資、招募、市場地位誘因） |
| 媒體引述「知情人士」、外流投資人簡報 | Interested-party（同上；媒體名列入 source） |
| 第三方量測（Similarweb、Sensor Tower、Ramp、OpenRouter 等） | Analogy（給區間；說明與 Anthropic 全體的差異） |
| 合作方公告（Google、Amazon、Microsoft、NVIDIA、Fluidstack） | Interested-party（合作方誘因：雲端業務宣傳） |
| 監管文件、法院文件 | Verified |

## 交付
- 三份 yaml＋notes；chat 端合併為 `data/anthropic_src.yaml`（加 `id: SRC_ANT_nnn`）並寫 `docs/reports/20261008_v0.1-A1_資料.md`＋`_對照.xlsx`。

# CHANGELOG（智譜收支模型）

## Z0（2026-10-08，chat 端）
- 建立 `zhipu/`：骨架（builder／engine／tools／tests 基礎自 OpenAI v0.6 經 Anthropic A0 複製，名稱待 Z2 改寫）、規格 r0（D1–D26）、共同規則、Z1–Z5 工作單、進度檔、交接檔。

## Z1（2026-10-08，chat 端＋三個研究代理）
- `data/zhipu_src.yaml`：SRC_ZP_001–604（財報 240、價格用量算力 251、融資市值 113），找不到 41 項；規格 r1（第 6 節：D2r、D8r、D9r、D10r、D11r、D14r、D16r、D21r、D22r、D24r）。

## v0.1-Z2（2026-10-08，建置代理）— Excel `model/20261008_Zhipu_v0.1.xlsx`
- commit：e425dce（步驟 1–3：builder 骨架、TK_Link 讀表 7 列、OAI_Link、Inputs 草稿）、6a177f6（步驟 4–6：Demand、Revenue、Checks）、569d362（步驟 7：parity 20 情境、test_builder、`zhipu-parity.yml`）、本節報告 commit。
- 新頁：README、SRC_ZP（604 列）、TK_Link（63 名＋7 讀表列）、OAI_Link（OpenAI v0.6 命題輸出 14 列）、Inputs（73 列）、Demand、Revenue、Checks（C01–C42）。
- 營收（RMB 億）：2025 7.24（校準差距 0）、2026 37.14（1H 實際 9.54＋2H 27.60）、2027 69.55、2030 226.66；雲端 2H26 年化 40.4 對 MaaS ARR 107.8 → WARN −63%。
- 驗收：pytest 59 passed；CHK_Errors＝0；TK 快照 OK 70。報告 `docs/reports/20261008_v0.1-Z2.md`＋`_對照.xlsx`。
- 工具：engine／tools 名稱改智譜；`check_tk_snapshot.py` 支援讀表列；`make_report_tables.py` 改為產生本模型對照 Excel。

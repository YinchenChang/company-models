# 2026-11-13 收盤換價（工作單 px1113，Andy 2026-11-16）：建置腳本在存檔前呼叫 apply(wb)。
# 只加版本說明（README A1、本版要點列、沿革追加列）；數值與公式不在此改（現價與同業報價在 build 腳本輸入處改）。
from openpyxl.styles import Alignment

VER, OLD, BUILT = "0.1.4", "0.1.3", "2026-11-16"
SHORT = "現價與同業報價改 2026-11-13 收盤（工作單 px1113）"
NOTE = '現價與同業報價改 2026-11-13 收盤（工作單 px1113；Yahoo Finance 日線收盤，markets/invest_map/data/closes_20261020_20261113.csv）。現價 81.71 → 108.72；同業前瞻 P/S 中位數 8.02 → 9.59（NVIDIA 8.14→7.30、AMD 7.89→10.42、Broadcom 8.84→10.53、Marvell 7.47→8.77；共識前瞻 P/S 與頁面價格不變，只把 11-05 收盤換成 11-13 收盤）；加權目標價 83.94（上檔 +2.7%，中立）→ 97.21（上檔 -10.6%，中立）。共識、rf、β、ERP、情境權重與營運假設不變；本模型無 Sens 快照與 Coord 另一側讀數快照。'


def apply(wb):
    RD = wb["README"]
    a1 = RD["A1"].value
    head = a1.split("；", 1)[0]
    assert "v" + OLD in head, head[:120]
    RD["A1"].value = head.replace("v" + OLD, "v" + VER, 1) + a1[len(head):] + f"；v{VER}（{BUILT}）：{SHORT}"
    rows = [r for r in range(1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith(f"v{OLD}（")]
    assert len(rows) == 1, rows
    ver_row = rows[0]
    old = RD.cell(ver_row, 1).value
    r = RD.max_row + 1
    RD.cell(r, 1, old); RD.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")   # 前一版本版要點移到沿革（README 末）
    RD.cell(ver_row, 1).value = f"v{VER}（{BUILT}）：{NOTE}"
    return f"px1113_post：README!A{ver_row} 改本版要點，前版移 A{r}"

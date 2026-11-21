# 2026-11-20 收盤換價（工作單 px1120，Andy 2026-11-21）：建置腳本在存檔前呼叫 apply(wb)。
# 只加版本說明（README A1、本版要點列、沿革追加列）；數值與公式不在此改（現價與同業報價在 build 腳本輸入處改）。
from openpyxl.styles import Alignment

VER, OLD, BUILT = "0.1.5", "0.1.4", "2026-11-21"
SHORT = "現價與同業報價改 2026-11-20 收盤（工作單 px1120）"
NOTE = '現價與同業報價改 2026-11-20 收盤（工作單 px1120；Yahoo Finance 日線收盤，markets/invest_map/data/closes_20261109_20261120.csv）。現價 108.72 → 82.67；同業前瞻 P/S 中位數 9.59 → 7.56（NVIDIA 7.30→5.63、AMD 10.42→8.28、Broadcom 10.53→9.24、Marvell 8.77→6.84；共識前瞻 P/S 與頁面價格不變，只把 11-13 收盤換成 11-20 收盤）；加權目標價 97.21（上檔 -10.6%，中立）→ 80.10（上檔 -3.1%，中立）。共識、rf、β、ERP、情境權重與營運假設不變；以市值計的 WACC 權重隨現價自動重算；本模型無 Sens 快照與 Coord 另一側讀數快照。未換價同業：無。'


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
    return f"px1120_post：README!A{ver_row} 改本版要點，前版移 A{r}"

# 2026-11-13 收盤換價（工作單 px1113，Andy 2026-11-16）：建置腳本在存檔前呼叫 apply(wb, …)。
# 只加版本說明（README A1、本版要點列、沿革追加列）；數值與公式不在此改（現價與同業報價在 build 腳本輸入處改）。
from openpyxl.styles import Alignment

VER, OLD_VER, BUILT = "0.5.8", "0.5.7", "2026-11-16"
SHORT = "現價與同業報價改 2026-11-13 收盤（工作單 px1113）"
NOTE = '現價與同業報價改 2026-11-13 收盤（工作單 px1113；Yahoo Finance 日線收盤，markets/invest_map/data/closes_20261020_20261113.csv）。現價 291.53 → 347.01；同業前瞻 P/E 中位數 16.93 → 18.72（NVIDIA 15.02→13.47、AMD 36.12→47.70、Marvell 23.66→27.77、Qualcomm 16.93→18.63、台積電 ADR 16.58→18.72；新倍數＝v0.5.7 倍數 × 11-13 收盤 ÷ 11-05 收盤，共識 EPS 不變）；加權目標價 163.70（上檔 -43.8%，放空）→ 173.71（上檔 -49.9%，放空）。共識 EPS、rf、β、ERP、情境權重與營運假設不變；Sens、Coord 快照已重跑。上檔空間 -49.9% 距 -50%（強力放空）邊界僅 0.1 個百分點。'


def apply(wb, ver_row):
    RD = wb["README"]
    a1 = RD["A1"].value
    head = a1.split("；", 1)[0]
    assert "v" + OLD_VER in head, head[:120]
    RD["A1"].value = head.replace("v" + OLD_VER, "v" + VER, 1) + a1[len(head):] + f"；v{VER}（{BUILT}）：{SHORT}"
    old = RD.cell(ver_row, 1).value
    assert str(old).startswith("v" + OLD_VER), str(old)[:60]
    r = RD.max_row + 1
    RD.cell(r, 1, old); RD.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")   # 前一版本版要點移到沿革（README 末）
    RD.cell(ver_row, 1).value = f"v{VER}（{BUILT}）：{NOTE}"
    return f"px1113_post：README!A{ver_row} 改本版要點，前版移 A{r}"

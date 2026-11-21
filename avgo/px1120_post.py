# 2026-11-20 收盤換價（工作單 px1120，Andy 2026-11-16）：建置腳本在存檔前呼叫 apply(wb, …)。
# 只加版本說明（README A1、本版要點列、沿革追加列）；數值與公式不在此改（現價與同業報價在 build 腳本輸入處改）。
from openpyxl.styles import Alignment

VER, OLD_VER, BUILT = "0.5.9", "0.5.8", "2026-11-21"
SHORT = "現價與同業報價改 2026-11-20 收盤（工作單 px1120）"
NOTE = '現價與同業報價改 2026-11-20 收盤（工作單 px1120；Yahoo Finance 日線收盤，markets/invest_map/data/closes_20261109_20261120.csv）。現價 347.01 → 304.41；同業前瞻 P/E 中位數 18.72 → 16.20（NVIDIA 13.47→10.38、AMD 47.70→37.90、Marvell 27.77→21.66、Qualcomm 18.63→16.20、台積電 ADR 18.72→14.96；新倍數＝v0.5.8 倍數 × 11-20 收盤 ÷ 11-13 收盤，共識 EPS 不變）；加權目標價 166.10（上檔 -52.1%，強力放空）→ 152.59（上檔 -49.9%，放空）。共識 EPS、rf、β、ERP、情境權重與營運假設不變；Sens、Coord 快照已重跑。上檔空間 -49.9% 距 -50%（強力放空）邊界僅 0.1 個百分點。'


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
    return f"px1120_post：README!A{ver_row} 改本版要點，前版移 A{r}"

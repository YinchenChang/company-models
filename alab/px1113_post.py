# 2026-11-13 收盤換價（工作單 px1113，Andy 2026-11-16）：建置腳本在存檔前呼叫 apply(wb)。
# 只加版本說明（README A1、本版要點列、沿革追加列）；數值與公式不在此改（現價與同業報價在 build 腳本輸入處改）。build_id 標籤由 v036_post.build_id 帶版號。
from openpyxl.styles import Alignment

VER, OLD, BUILT = "0.3.9", "0.3.8", "2026-11-16"
SHORT = "現價與同業報價改 2026-11-13 收盤（工作單 px1113）"
NOTE = '現價與同業報價改 2026-11-13 收盤（工作單 px1113；Yahoo Finance 日線收盤，markets/invest_map/data/closes_20261020_20261113.csv）。現價 179.47 → 221.83；同業前瞻 P/E 中位數 20.53 → 24.22（Credo 15.40→18.03、Marvell 23.63→27.74、Broadcom 16.78→19.98、Rambus 20.53→24.22、MACOM 23.04→28.08；新倍數＝v0.3.8 倍數 × 11-13 收盤 ÷ 11-05 收盤，共識 EPS 不變）；加權目標價 101.32（上檔 -43.5%，放空）→ 115.24（上檔 -48.1%，放空）。共識 EPS、rf、β、ERP、情境權重與營運假設不變；Sens、Coord 快照已重跑。build_id 標籤改 alab-v0.3.9。'


def apply(wb):
    RD = wb["README"]
    RD["A1"].value = str(RD["A1"].value) + f"；v{VER}（{BUILT}）：{SHORT}"
    rows = [r for r in range(1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith(f"v{OLD}（")]
    assert len(rows) == 1, rows
    ver_row = rows[0]
    old = RD.cell(ver_row, 1).value
    r = RD.max_row + 1
    RD.cell(r, 1, old); RD.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")   # 前一版本版要點移到沿革（README 末）
    RD.cell(ver_row, 1).value = f"v{VER}（{BUILT}）：{NOTE}"
    return f"px1113_post：README!A{ver_row} 改本版要點，前版移 A{r}"

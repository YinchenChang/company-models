# 2026-11-20 收盤換價（工作單 px1120，Andy 2026-11-21）：建置腳本在存檔前呼叫 apply(wb)。
# 只加版本說明（README A1、本版要點列、沿革追加列）；數值與公式不在此改（現價與同業報價在 build 腳本輸入處改）。build_id 標籤由 v036_post.build_id 帶版號。
from openpyxl.styles import Alignment

VER, OLD, BUILT = "0.3.10", "0.3.9", "2026-11-21"
SHORT = "現價與同業報價改 2026-11-20 收盤（工作單 px1120）"
NOTE = '現價與同業報價改 2026-11-20 收盤（工作單 px1120；Yahoo Finance 日線收盤，markets/invest_map/data/closes_20261109_20261120.csv）。現價 221.83 → 180.35；同業前瞻 P/E 中位數 24.22 → 19.54（Credo 18.03→13.69、Marvell 27.74→21.64、Broadcom 19.98→17.53、Rambus 24.22→19.54、MACOM 28.08→23.59；新倍數＝v0.3.9 倍數 × 11-20 收盤 ÷ 11-13 收盤，共識 EPS 不變）；Comparable 目標價 166.71 → 134.51；加權目標價 106.42（上檔 -52.0%，放空）→ 90.33（上檔 -49.9%，放空）。共識 EPS、rf、β、ERP、情境權重與營運假設不變；Sens、Coord 快照已重跑。build_id 標籤改 alab-v0.3.10。'


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
    return f"px1120_post：README!A{ver_row} 改本版要點，前版移 A{r}"

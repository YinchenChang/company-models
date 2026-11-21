# 2026-11-20 收盤換價（工作單 px1120，Andy 2026-11-21）：建置腳本在存檔前呼叫 apply(wb)。
# 只加版本說明（README A1、A2、本版要點列、沿革追加列）與版號字串；數值與公式不在此改（現價與同業報價在 build 腳本輸入處改）。
from openpyxl.styles import Alignment

VER, OLD, BUILT = "0.3.7", "0.3.6", "2026-11-21"
SHORT = "現價與同業報價改 2026-11-20 收盤（工作單 px1120）"
NOTE = '現價與同業報價改 2026-11-20 收盤（工作單 px1120；Yahoo Finance 日線收盤，markets/invest_map/data/closes_20261109_20261120.csv）。現價 190.53 → 152.20；同業前瞻 P/E 中位數 23.92 → 19.74（Qualcomm 18.63→16.20、Broadcom 19.95→17.50、Marvell 27.77→21.66、Rambus 22.07→17.81、Synopsys 25.77→22.43、Cadence 34.28→29.88；新倍數＝v0.3.6 倍數 × 11-20 收盤 ÷ 11-13 收盤，共識 EPS 不變）；加權目標價 42.59（上檔 -77.6%，強力放空）→ 36.00（上檔 -76.3%，強力放空）。共識 EPS、rf、β、ERP、情境權重與營運假設不變；Sens、Coord 快照已重跑。上檔空間 -76.3%，仍在強力放空級距。'


def apply(wb):
    RD, SE = wb["README"], wb["Sens"]
    a1 = RD["A1"].value
    assert isinstance(a1, str) and f"個股模型 v{OLD}（" in a1.split("；", 1)[0] and f"build_id arm-v{OLD}-" in a1, a1[:120]
    RD["A1"].value = (a1.replace(f"個股模型 v{OLD}（", f"個股模型 v{VER}（", 1).replace(f"build_id arm-v{OLD}-", f"build_id arm-v{VER}-", 1)
                      + f"；v{VER}（{BUILT}）：{SHORT}")
    a2 = str(RD["A2"].value)
    assert a2.startswith(f"版本：v{OLD}（"), a2[:40]
    RD["A2"].value = f"版本：v{VER}（{BUILT}）：{SHORT}；" + a2[len("版本："):]
    rb = next(r for r in range(1, SE.max_row + 1) if str(SE.cell(r, 1).value or "").startswith("build_id（README、Sens、HTML"))
    bid = str(SE.cell(rb, 4).value)
    assert bid.startswith(f"arm-v{OLD}-"), bid
    SE.cell(rb, 4).value = bid.replace(f"arm-v{OLD}-", f"arm-v{VER}-", 1)
    rows = [r for r in range(1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith(f"v{OLD}（")]
    assert len(rows) == 1, rows
    ver_row = rows[0]
    old = RD.cell(ver_row, 1).value
    r = RD.max_row + 1
    RD.cell(r, 1, old); RD.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")   # 前一版本版要點移到沿革（README 末）
    RD.cell(ver_row, 1).value = f"v{VER}（{BUILT}）：{NOTE}"
    return f"px1120_post：README!A{ver_row} 改本版要點，前版移 A{r}"

# 台積電 v0.5.7（2026-11-21，工作單 px1120）：現價與同業報價改 2026-11-20 收盤。
# 價格輸入與同業倍數在 build_tsmc.py（px、px_guc、px_vis、PEERS 末欄）；本檔只處理 README 版號與「本版要點」列。
PX1120_ROW = "v0.5.7（2026-11-21）：現價與同業報價改 2026-11-20 收盤（工作單 px1120）：現價 2,350 → 2,000；創意 8,610 → 5,400、采鈺 400.5 → 290（非控制權益市值）；同業前瞻 P/E＝現行倍數 × 11-20 收盤 ÷ 11-13 收盤：聯電 16.4 → 12.9、格芯 19.2 → 16.4、英特爾 34.7 → 29.9、中芯 34.4 → 34.8、三星 3.09 → 2.74，中位數 19.2 → 16.4（落在格芯）；DCF 目標價 1,416 → 1,429，Comparable（相對倍數調整後）3,145 → 2,696，加權目標價 2,280 → 2,063，上檔空間 −3.0% → +3.1%，投資建議 中立 → 中立（上檔空間仍在中立級距內）；共識 EPS、rf、預測匯率與模型假設不變；Sens、Coord 快照重跑。"
VER_OLD, VER_NEW = "v0.5.6", "v0.5.7"


def apply(wb, px1113_row):
    rd = wb["README"]
    h = str(rd["A1"].value)
    assert VER_OLD in h.split("；", 1)[0], h[:80]
    rd["A1"] = h.replace(VER_OLD, VER_NEW, 1) + "；v0.5.7（2026-11-21）：現價與同業報價改 2026-11-20 收盤（工作單 px1120）"
    rows = [r for r in range(1, rd.max_row + 1) if str(rd.cell(r, 1).value or "").startswith("v0.5.6（2026-11-16）：現價與同業報價改 2026-11-13 收盤")]
    assert len(rows) == 1, rows
    rd.cell(rows[0], 1).value = PX1120_ROW + "上一版：" + str(rd.cell(rows[0], 1).value)

# 2026-11-20 收盤換價（工作單 px1120，Andy 2026-11-21）：建置腳本在存檔前呼叫 apply(wb)。
# 只加版本說明（README A1、本版要點列、沿革追加列）與版號字串；數值與公式不在此改（現價與同業報價在 build 腳本輸入處改）。
from openpyxl.styles import Alignment

VER, OLD, BUILT = "0.4.9", "0.4.8", "2026-11-21"
SHORT = "現價與同業報價改 2026-11-20 收盤（工作單 px1120）"
NOTE = '現價與同業報價改 2026-11-20 收盤（工作單 px1120；Yahoo Finance 日線收盤，markets/invest_map/data/closes_20261109_20261120.csv）。現價 534.06 → 424.39；同業前瞻 P/E 中位數 19.95 → 17.50（NVIDIA 13.47→10.38、Broadcom 19.95→17.50、Intel 38.66→33.23、Qualcomm 18.63→16.20、Marvell 27.77→21.66；新倍數＝v0.4.8 倍數 × 11-20 收盤 ÷ 11-13 收盤，共識 EPS 不變）；加權目標價 300.30（上檔 -43.8%，放空）→ 269.49（上檔 -36.5%，放空）。共識 EPS、rf、β、ERP、情境權重與營運假設不變；Sens、Coord 快照已重跑。股價基準日與資料截止日改 2026-11-20（目標時點＝現價日＋12 個月隨之後移）；成品日 2026-11-21。'


def lit(t):
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def apply(wb):
    RD, SE = wb["README"], wb["Sens"]
    a1 = RD["A1"].value
    assert isinstance(a1, str) and a1.startswith("=") and f"模型 v{OLD}（" in a1, a1[:120]
    a1 = a1.replace(f"模型 v{OLD}（", f"模型 v{VER}（", 1).replace(f"build_id amd-v{OLD}-", f"build_id amd-v{VER}-", 1)
    RD["A1"].value = a1 + "&" + lit(f"；v{VER}（{BUILT}）：{SHORT}")
    rb = next(r for r in range(1, SE.max_row + 1) if str(SE.cell(r, 1).value or "").startswith("build_id（README、Sens、HTML"))
    bid = str(SE.cell(rb, 4).value)
    assert bid.startswith(f"amd-v{OLD}-"), bid
    SE.cell(rb, 4).value = bid.replace(f"amd-v{OLD}-", f"amd-v{VER}-", 1)
    rows = [r for r in range(1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith(f"v{OLD}（")]
    assert len(rows) == 1, rows
    ver_row = rows[0]
    old = RD.cell(ver_row, 1).value
    r = RD.max_row + 1
    RD.cell(r, 1, old); RD.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")   # 前一版本版要點移到沿革（README 末）
    RD.cell(ver_row, 1).value = f"v{VER}（{BUILT}）：{NOTE}"
    return f"px1120_post：README!A{ver_row} 改本版要點，前版移 A{r}"

# 雙鴻 v0.2.8（2026-11-21）後處理：現價與同業報價改 2026-11-20 收盤（工作單 px1120；只換價格，公式與架構不改）
# 價格本體在 build_auras.py（px、PEERS）；這裡只改 README 版本文字。
import re
VER, BUILT = "v0.2.8", "2026-11-21"
NOTE = "現價與同業報價改 2026-11-20 收盤（工作單 px1120）：現價 1,140（11-13 收盤）→ 794（Yahoo Finance 日線收盤，closes_20261109_20261120.csv）；同業前瞻 P/E 中位數 24.9 → 19.3 倍（奇鋐 3,040→2,205、健策 5,345→3,705、Vertiv 202.07→166.24、台達電 1,240→900；共識 EPS 不變）；加權目標價 1,740.9 → 1,407.7 元（DCF 1,355.0 → 1,292.5、Comparable 相對倍數調整後 2,126.8 → 2,083.1）；上檔空間 +52.7% → +77.3%；評等強力買進 → 強力買進。另一側讀數（Coord）1,270.0（+60.0%，強力買進；coord_alt.json 重跑）。未換價同業：無（四家同業皆在 CSV）。rf、β、ERP、共識 EPS、預測用匯率、營運假設與情境權重不變"


def apply(wb, P):
    RD = wb["README"]
    a1 = RD["A1"].value
    head, sep, rest = a1.partition("；")
    head2, n = re.subn(r"v\d+(?:\.\d+)+（", f"{VER}（", head, count=1)
    assert n == 1, head
    head2 = re.sub("　20\\d\\d-\\d\\d-\\d\\d", "　" + BUILT, head2, count=1)
    RD["A1"].value = head2 + sep + rest + f"；{VER}（{BUILT}）：" + NOTE
    RM = wb._v13["RM"]
    rv = RD.cell(RM["ver"], 1)
    t = rv.value
    assert isinstance(t, str) and t.startswith("v0.2.7"), t[:40]
    rv.value = f"{VER}（{BUILT}）：" + NOTE + "。前一版：" + t
    return f"px1120_post：README 版本文字 {VER}"

# 奇鋐 v0.3.8（2026-11-21）後處理：現價與同業報價改 2026-11-20 收盤（工作單 px1120；只換價格，公式與架構不改）
# 價格本體在 build_avc.py（px、peers、富世達股價）；這裡只改 README 版本文字。須在 px1113_post 之後執行。
import re
VER, BUILT = "v0.3.8", "2026-11-21"
NOTE = "現價與同業報價改 2026-11-20 收盤（工作單 px1120）：現價 3,040（11-13 收盤）→ 2,205；Comparable 同業前瞻 P/E 中位數 24.9 → 19.3 倍（雙鴻 1,140→794、健策 5,345→3,705、Vertiv 202.07→166.24、台達電 1,240→900；共識 EPS 不變）；加權目標價 2,423.6 → 2,192.2 元（DCF 2,382.6 → 2,473.2、Comparable 相對倍數調整後 2,464.6 → 1,911.1）；上檔空間 −20.3% → −0.6%；評等中立 → 中立。非控制權益用富世達（6805）股價 1,975 → 1,430 元（2026-11-20 收盤）；未換價：無。價格取自 closes_20261109_20261120.csv（Yahoo Finance 日線收盤）。rf、β、ERP、共識 EPS、預測用匯率與營運假設不變"


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
    assert isinstance(t, str) and t.startswith("v0.3.7"), t[:40]
    rv.value = f"{VER}（{BUILT}）：" + NOTE + "。前一版：" + t
    return f"px1120_post：README 版本文字 {VER}"

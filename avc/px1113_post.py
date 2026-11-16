# 奇鋐 v0.3.7（2026-11-16）後處理：現價與同業報價改 2026-11-13 收盤（工作單 px1113；只換價格，公式與架構不改）
# 價格本體在 build_*.py（px、peers 等）；這裡只改 README 版本文字。
import re
VER, BUILT = "v0.3.7", "2026-11-16"
NOTE = "現價與同業報價改 2026-11-13 收盤（工作單 px1113）：現價 2,565 → 3,040（Yahoo Finance 日線收盤，closes_20261020_20261113.csv）；同業前瞻 P/E 中位數 22.2 → 24.9 倍（雙鴻、健策、Vertiv、台達電；共識 EPS 不變）；加權目標價 2,446 → 2,586 元；上檔空間 −4.7% → −14.9%；評等中立 → 中立。未換價：富世達（6805）股價 1,875 元（非控制權益用；CSV 無此代號，維持 2026-11-05 值）。rf、β、ERP、共識 EPS、預測用匯率與營運假設不變"


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
    assert isinstance(t, str) and t.startswith("v0.3.6"), t[:40]
    rv.value = f"{VER}（{BUILT}）：" + NOTE + "。前一版：" + t
    return f"px1113_post：README 版本文字 {VER}"

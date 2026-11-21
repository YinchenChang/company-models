# 美超微 v0.3.9（2026-11-21）後處理：現價與同業報價改 2026-11-20 收盤（工作單 px1120；只換價格，公式與架構不改）
# 價格本體在 build_smci.py（px、PEERS）；這裡只改 README 版本文字。build_smci.py 在 px1113_post.apply 之後（有 Coord 頁時）呼叫。
import re
VER, BUILT = "v0.3.9", "2026-11-21"
NOTE = "現價與同業報價改 2026-11-20 收盤（工作單 px1120）：現價與同業報價取自 closes_20261109_20261120.csv（Yahoo Finance 日線收盤），公式與架構不改；同業前瞻 P/E 以 11-13 倍數 × 11-20 收盤 ÷ 11-13 收盤換算（戴爾、HPE、緯穎、廣達、緯創；共識 EPS 不變）；前後值見本版要點。未換價同業：無。rf、β、ERP、共識 EPS、預測用匯率與營運假設不變"


def apply(wb, POST):
    RD = wb["README"]
    a1 = RD["A1"].value
    head, sep, rest = a1.partition("；")
    head2, n = re.subn(r"v\d+(?:\.\d+)+（", f"{VER}（", head, count=1)
    if n != 1:
        head2, n = re.subn(r"v\d+\.\d+(?:\.\d+)?", VER, head, count=1)
    assert n == 1, head
    RD["A1"].value = head2 + sep + rest + f"；{VER}（{BUILT}）：" + NOTE
    RM = wb._v037["RM"]
    rv = RD.cell(RM["ver"], 1)
    t = rv.value
    assert isinstance(t, str) and t.startswith(('="v0.3.8', "v0.3.8")), t[:40]
    if t.startswith("="):
        rv.value = '="' + f"{VER}（{BUILT}）：" + NOTE.replace('"', '""') + '。前一版："&' + t[1:]
    else:
        rv.value = f"{VER}（{BUILT}）：" + NOTE + "。前一版：" + t
    return f"px1120_post：README 版本文字 {VER}"

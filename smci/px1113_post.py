# 美超微 v0.3.8（2026-11-16）後處理：現價與同業報價改 2026-11-13 收盤（工作單 px1113；只換價格，公式與架構不改）
# 價格本體在 build_smci.py（px、PEERS）；這裡只改 README 版本文字。build_smci.py 在 v037_post.apply 之後（有 Coord 頁時）呼叫。
import re
VER, BUILT = "v0.3.8", "2026-11-16"
NOTE = "現價與同業報價改 2026-11-13 收盤（工作單 px1113）：現價 27.08 → 26.87 美元（Yahoo Finance 日線收盤，closes_20261020_20261113.csv）；同業前瞻 P/E 中位數 10.3 → 10.2 倍（戴爾、HPE、緯穎、廣達、緯創；共識 EPS 不變）；加權目標價 65.18 → 64.94 美元；上檔空間 +140.7% → +141.7%；評等強力買進 → 強力買進。未換價同業：無（五家同業皆在 CSV）。rf、β、ERP、共識 EPS、預測用匯率與營運假設不變"


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
    assert isinstance(t, str) and t.startswith(('="v0.3.7', "v0.3.7")), t[:40]
    if t.startswith("="):
        rv.value = '="' + f"{VER}（{BUILT}）：" + NOTE.replace('"', '""') + '。前一版："&' + t[1:]
    else:
        rv.value = f"{VER}（{BUILT}）：" + NOTE + "。前一版：" + t
    return f"px1113_post：README 版本文字 {VER}"

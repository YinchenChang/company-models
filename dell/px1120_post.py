# 戴爾 v0.2.8（2026-11-21）後處理：現價與同業報價改 2026-11-20 收盤（工作單 px1120；只換價格，公式與架構不改）
# 價格本體在 build_dell.py（px、PEERS）；這裡只改 README 版本文字。接在 px1113_post.apply 之後呼叫。
import re
VER, BUILT = "v0.2.8", "2026-11-21"
NOTE = "現價與同業報價改 2026-11-20 收盤（工作單 px1120）：現價 344.53 → 278.02 美元（Yahoo Finance 日線收盤，closes_20261109_20261120.csv）；同業前瞻 P/E 中位數 9.8 → 8.8 倍（HPE、SMCI、聯想、HPQ、NetApp；共識 EPS 不變）；DCF 254.10 → 255.50、Comparable（相對倍數調整後）199.46 → 178.54 美元；加權目標價 226.78 → 217.02 美元；上檔空間 −34.2% → −21.9%；評等放空 → 中立（模型公式重算；前後值取自 v0.2.7 成品 xlsx 的重算值，與 v0.2.7 版本文字所寫 238.57／−30.8% 不同，以成品為準）。未換價同業：無（五家同業皆在 CSV）。rf、β、ERP、共識 EPS、預測用匯率與營運假設不變"


def apply(wb, POST):
    RD = wb["README"]
    a1 = RD["A1"].value
    head, sep, rest = a1.partition("；")
    head2, n = re.subn(r"v\d+(?:\.\d+)+", VER, head, count=1)
    assert n == 1, head
    RD["A1"].value = head2 + sep + rest + f"；{VER}（{BUILT}）：" + NOTE
    RM = wb._v026["RM"]
    rv = RD.cell(RM["ver"], 1)
    t = rv.value
    assert isinstance(t, str) and t.startswith(('="v0.2.7', "v0.2.7")), t[:40]
    if t.startswith("="):
        rv.value = '="' + f"{VER}（{BUILT}）：" + NOTE.replace('"', '""') + '。前一版："&' + t[1:]
    else:
        rv.value = f"{VER}（{BUILT}）：" + NOTE + "。前一版：" + t
    return f"px1120_post：README 版本文字 {VER}"

# 光寶科 v0.4.7（2026-11-21）後處理：現價與同業報價改 2026-11-20 收盤（工作單 px1120；只換價格，公式與架構不改）
# 價格本體在 build_liteon.py（px、PEERS）；這裡只改 README 版本文字。
import re
VER, BUILT = "v0.4.7", "2026-11-21"
NOTE = "現價與同業報價改 2026-11-20 收盤（工作單 px1120）：現價 231.5（11-13 收盤）→ 175.5（Yahoo Finance 日線收盤，closes_20261109_20261120.csv）；同業前瞻 P/E 中位數 23.8 → 19.1 倍（台達電 1,240→900、Vertiv 202.07→166.24、Eaton 352.97→311.33、Flex 74.98→66.24；共識 EPS 不變）；加權目標價 279.24 → 250.78 元（DCF 263.5 → 264.9、Comparable 487.5 → 391.2；前版 README 寫 294.8 是 px1113 當時值，之後 Siliconomics 重建使 11-13 現價下的成品為 279.24）；上檔空間 +20.6% → +42.9%；評等中立 → 買進。另一側讀數（Coord）258.54 中立 → 232.56 買進（coord_alt.json 重跑）。未換價同業：無（四家同業皆在 CSV）。rf、β、ERP、共識 EPS、預測用匯率與營運假設不變"


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
    assert isinstance(t, str) and t.startswith("v0.4.6"), t[:40]
    rv.value = f"{VER}（{BUILT}）：" + NOTE + "。前一版：" + t
    return f"px1120_post：README 版本文字 {VER}"

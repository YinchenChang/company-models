# 台達電 v0.5.8（2026-11-16）後處理：現價與同業報價改 2026-11-13 收盤（工作單 px1113；只換價格，公式與架構不改）
# build_delta.py 在 v057_post.apply 之後、存檔前呼叫 apply(wb, P56)。價格本體在 build_delta.py（px、peers）；這裡只改版本文字。
VER, BUILT = "0.5.8", "2026-11-16"
NOTE = ("現價與同業報價改 2026-11-13 收盤（工作單 px1113）：現價 1,130 → 1,240（Yahoo Finance 日線收盤，closes_20261020_20261113.csv）；"
        "同業前瞻 P/E 中位數 20.0 → 21.5 倍（Vertiv、Eaton、Schneider、光寶科；共識 EPS 不變）；加權目標價 931 → 980 元；上檔空間 −17.6% → −21.0%；評等中立 → 中立。"
        "非控制權益用台達泰國（DELTA.BK）股價 187 → 190 泰銖、USD/TWD 33.034 → 32.939、USD/THB 33.42 → 33.85（Yahoo Finance 日線收盤 2026-11-13）：台達泰國少數股東市值 871.5 → 871.8、非控制權益合計 898.4 → 899.2（NT$B）；晶睿（3454）維持帳面；預測用匯率不變。未換價：無。"
        "rf、β、ERP、共識 EPS、預測用匯率與營運假設不變")


def apply(wb, P56):
    RD = wb["README"]
    a1 = RD["A1"].value
    old = "個股模型 v0.5.7（"
    assert old in a1, a1[:120]
    import re
    a1 = a1.replace(old, f"個股模型 v{VER}（", 1)
    a1 = re.sub("\u300020\\d\\d-\\d\\d-\\d\\d", "\u3000" + BUILT, a1, count=1)
    RD["A1"].value = a1 + f"；v{VER}（{BUILT}）：" + NOTE
    RM = wb._v056["RM"]
    rv = RD.cell(RM["ver"], 1)
    t = rv.value
    assert isinstance(t, str) and t.startswith("v0.5.7"), t[:40]
    rv.value = f"v{VER}（{BUILT}）：" + NOTE + "。前一版：" + t
    return "px1113_post：README 版本文字 v0.5.8"

# 台達電 v0.5.9（2026-11-21）後處理：現價與同業報價改 2026-11-20 收盤（工作單 px1120；只換價格，公式與架構不改）
# build_delta.py 在 px1113_post.apply 之後、存檔前呼叫 apply(wb, P56)。價格本體在 build_delta.py（px、peers、非控制權益）；這裡只改版本文字。
VER, BUILT = "0.5.9", "2026-11-21"
NOTE = ("現價與同業報價改 2026-11-20 收盤（工作單 px1120）：現價 1,240 → 900（Yahoo Finance 日線收盤，closes_20261109_20261120.csv）；"
        "同業前瞻 P/E：Vertiv 202.07 → 166.24、Eaton 352.97 → 311.33、Schneider 243.45 → 226.55 歐元、光寶科 231.5 → 175.5（共識 EPS 不變）；"
        "非控制權益用台達泰國（DELTA.BK）股價 190 → 173.5 泰銖、USD/TWD 32.939 → 33.026、USD/THB 33.85 → 33.942（2026-11-20 收盤）：非控制權益合計 899.2 → 823.5（NT$B 口徑同前）；晶睿（3454）維持帳面；預測用匯率不變。未換價：無。"
        "加權目標價、上檔空間與評等由模型公式重算（見 README 本版要點）。rf、β、ERP、共識 EPS、預測用匯率與營運假設不變")


def apply(wb, P56):
    import re
    RD = wb["README"]
    a1 = RD["A1"].value
    old = "個股模型 v0.5.8（"
    assert old in a1, a1[:120]
    a1 = a1.replace(old, f"個股模型 v{VER}（", 1)
    a1 = re.sub("　20\\d\\d-\\d\\d-\\d\\d", "　" + BUILT, a1, count=1)
    RD["A1"].value = a1 + f"；v{VER}（{BUILT}）：" + NOTE
    RM = wb._v056["RM"]
    rv = RD.cell(RM["ver"], 1)
    t = rv.value
    assert isinstance(t, str) and t.startswith("v0.5.8"), t[:40]
    rv.value = f"v{VER}（{BUILT}）：" + NOTE + "。前一版：" + t
    return "px1120_post：README 版本文字 v0.5.9"

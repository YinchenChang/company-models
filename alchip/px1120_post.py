# 世芯-KY v0.2.7（2026-11-21）後處理：現價與同業報價改 2026-11-20 收盤（工作單 px1120；只換價）。
# 數值改動全在 build_alchip.py（In 頁 px、fpe；Val 0.2 同業前瞻 P/E）；本檔只改版號與版本說明文字，不改任何數值公式。
# 建檔程式在 v025_post.apply 之後、存檔前呼叫 apply(wb)。另一側讀數建置（COORD_RUN＝1）與 sens_alchip.py 建置（ALCHIP_SENS＝1）時不執行。
import os

VER, BUILT, PREV = "0.2.7", "2026-11-21", "0.2.6"
NOTE = '現價與同業報價改 2026-11-20 收盤（工作單 px1120；只換價，共識 EPS、rf、β、ERP、模型假設與情境權重不變）。現價 4,020 → 3,240；本公司前瞻 P/E 19.72 → 15.89；同業前瞻 P/E 中位數 39.82 → 32.38（創意 105.93 → 66.44、智原 41.29 → 32.86、聯發科 39.82 → 32.38、博通 19.95 → 17.50、邁威爾 27.77 → 21.66）；DCF 目標價 2,874.81 → 2,874.91；Comparable 目標價（相對倍數調整後）8,119 → 6,602（市場主流同業法未調整 7,770.23 → 6,318.44）；加權目標價 5,496.70 → 4,738.31（上檔 +36.7% → +46.2%），評等 買進 → 買進。另一側讀數（C4）加權目標價 5,176.97 → 4,465.99（買進 → 買進）。前後值取自 v0.2.6 與本版 dist 的重算值（v0.2.6 說明列的 5,688.32 為當時數字，之後 Siliconomics 快照更新使 dist 重算值變為 5,496.70）。未換價：近 12 月 P/E（pe_ttm 50.74，資料商值）。'


def apply(wb):
    if os.environ.get("COORD_RUN") == "1" or "Coord" not in wb.sheetnames:
        return "px1120_post：略過（另一側讀數建置或無 Coord 頁）"
    RD = wb["README"]
    a1 = RD["A1"].value
    head = a1.split("；", 1)[0]
    assert f"v{PREV}（AI 半導體" in head, head[:80]
    RD["A1"].value = head.replace(f"v{PREV}", f"v{VER}", 1) + a1[len(head):] + f"；v{VER}（{BUILT}）：現價與同業報價改 2026-11-20 收盤（工作單 px1120）"
    a2 = str(RD["A2"].value)
    assert a2.startswith(f"版本：v{PREV}（"), a2[:40]
    RD["A2"].value = a2.replace(f"版本：v{PREV}（", f"版本：v{VER}（{BUILT}：現價與同業報價改 2026-11-20 收盤；前版 v{PREV}，2026-11-16）；v{PREV}（", 1)
    # 本版要點（HTML「資料狀態」章讀 README v1.3 區最後一列）：本版說明在前，前兩版說明依序往下一列移
    r13 = next(r for r in range(1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith("v1.3 展示層"))
    rv = r13 + 9
    old = [str(RD.cell(rv + i, 1).value) for i in range(3)]
    assert old[0].startswith(f"v{PREV}（") and old[1].startswith("v0.2.5（") and old[2].startswith("v0.2.4（"), [o[:30] for o in old]
    assert RD.cell(rv + 3, 1).value in (None, ""), RD.cell(rv + 3, 1).value
    RD.cell(rv, 1, f"v{VER}（{BUILT}）：{NOTE}")
    for i, o in enumerate(old, 1):
        RD.cell(rv + i, 1, o)
    return f"px1120_post：README A1／A2 版號、本版要點 README!A{rv}（v{PREV}、v0.2.5、v0.2.4 說明下移）"

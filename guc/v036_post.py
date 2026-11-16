# 創意 v0.3.6（2026-11-16）後處理：現價與同業報價改 2026-11-13 收盤（工作單 px1113；只換價）。
# 數值改動全在 build_guc.py（In 頁 px、fpe；Val 0.2 同業前瞻 P/E）；本檔只改版號與版本說明文字，不改任何數值公式。
# 建檔程式在 v035_post.apply 之後、存檔前呼叫 apply(wb)。另一側讀數建置（COORD_RUN＝1）與 sens_guc.py 建置（GUC_SENS＝1）時不執行。
import os

VER, BUILT, PREV = "0.3.6", "2026-11-16", "0.3.5"
NOTE = '現價與同業報價改 2026-11-13 收盤（工作單 px1113；只換價，共識 EPS、rf、β、ERP、模型假設與情境權重不變）。現價 6,405 → 8,610；本公司前瞻 P/E 78.8 → 105.93；同業前瞻 P/E 中位數 20.21 → 23.86（世芯 14.52 → 19.72、聯發科 33.44 → 39.82、博通 16.76 → 19.95、邁威爾 23.66 → 27.77）；DCF 目標價 1,684.95 → 1,684.92；加權目標價 2,841.08 → 3,202.03（上檔 −55.6% → −62.8%），評等 強力放空 → 強力放空。另一側讀數（C4）加權目標價 2,591.69 → 2,919.18（強力放空）。近 12 月 P/E 為 px ÷ 近 12 月 EPS，隨現價自動更新。'


def apply(wb):
    if os.environ.get("COORD_RUN") == "1" or "Coord" not in wb.sheetnames:
        return "v036_post：略過（另一側讀數建置或無 Coord 頁）"
    RD = wb["README"]
    a1 = RD["A1"].value
    head = a1.split("；", 1)[0]
    assert f"v{PREV}（AI 半導體" in head, head[:80]
    RD["A1"].value = head.replace(f"v{PREV}", f"v{VER}", 1) + a1[len(head):] + f"；v{VER}（{BUILT}）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）"
    a2 = str(RD["A2"].value)
    assert a2.startswith(f"版本：v{PREV}（"), a2[:40]
    RD["A2"].value = a2.replace(f"版本：v{PREV}（", f"版本：v{VER}（{BUILT}：現價與同業報價改 2026-11-13 收盤；前版 v{PREV}，2026-11-15）；v{PREV}（", 1)
    r13 = next(r for r in range(1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith("v1.3 展示層"))
    rv = r13 + 9
    o5, o4 = str(RD.cell(rv, 1).value), str(RD.cell(rv + 1, 1).value)
    assert o5.startswith(f"v{PREV}（") and o4.startswith("v0.3.4（"), (o5[:30], o4[:30])
    assert RD.cell(rv + 2, 1).value in (None, ""), RD.cell(rv + 2, 1).value
    RD.cell(rv, 1, f"v{VER}（{BUILT}）：{NOTE}")
    RD.cell(rv + 1, 1, o5)
    RD.cell(rv + 2, 1, o4)
    return f"v036_post：README A1／A2 版號、本版要點 README!A{rv}（v{PREV}、v0.3.4 說明下移）"

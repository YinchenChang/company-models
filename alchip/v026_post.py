# 世芯-KY v0.2.6（2026-11-16）後處理：現價與同業報價改 2026-11-13 收盤（工作單 px1113；只換價）。
# 數值改動全在 build_alchip.py（In 頁 px、fpe；Val 0.2 同業前瞻 P/E）；本檔只改版號與版本說明文字，不改任何數值公式。
# 建檔程式在 v025_post.apply 之後、存檔前呼叫 apply(wb)。另一側讀數建置（COORD_RUN＝1）與 sens_alchip.py 建置（ALCHIP_SENS＝1）時不執行。
import os

VER, BUILT, PREV = "0.2.6", "2026-11-16", "0.2.5"
NOTE = '現價與同業報價改 2026-11-13 收盤（工作單 px1113；只換價，共識 EPS、rf、β、ERP、模型假設與情境權重不變）。現價 2,960 → 4,020；本公司前瞻 P/E 14.52 → 19.72；同業前瞻 P/E 中位數 33.44 → 39.82（創意 78.8 → 105.93、智原 36.13 → 41.29、聯發科 33.44 → 39.82、博通 16.76 → 19.95、邁威爾 23.66 → 27.77）；DCF 目標價 2,967.58 → 2,967.43；加權目標價 5,014.73 → 5,688.32（上檔 +69.4% → +41.5%），評等 強力買進 → 買進。另一側讀數（C4）加權目標價 4,505.62 → 5,105.27（強力買進 → 買進）。未換價：近 12 月 P/E（pe_ttm 50.74，資料商值）。'


def apply(wb):
    if os.environ.get("COORD_RUN") == "1" or "Coord" not in wb.sheetnames:
        return "v026_post：略過（另一側讀數建置或無 Coord 頁）"
    RD = wb["README"]
    a1 = RD["A1"].value
    head = a1.split("；", 1)[0]
    assert f"v{PREV}（AI 半導體" in head, head[:80]
    RD["A1"].value = head.replace(f"v{PREV}", f"v{VER}", 1) + a1[len(head):] + f"；v{VER}（{BUILT}）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）"
    a2 = str(RD["A2"].value)
    assert a2.startswith(f"版本：v{PREV}（"), a2[:40]
    RD["A2"].value = a2.replace(f"版本：v{PREV}（", f"版本：v{VER}（{BUILT}：現價與同業報價改 2026-11-13 收盤；前版 v{PREV}，2026-11-15）；v{PREV}（", 1)
    # 本版要點（HTML「資料狀態」章讀 README v1.3 區最後一列）：本版說明在前，前兩版說明依序往下一列移
    r13 = next(r for r in range(1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith("v1.3 展示層"))
    rv = r13 + 9
    o5, o4 = str(RD.cell(rv, 1).value), str(RD.cell(rv + 1, 1).value)
    assert o5.startswith(f"v{PREV}（") and o4.startswith("v0.2.4（"), (o5[:30], o4[:30])
    assert RD.cell(rv + 2, 1).value in (None, ""), RD.cell(rv + 2, 1).value
    RD.cell(rv, 1, f"v{VER}（{BUILT}）：{NOTE}")
    RD.cell(rv + 1, 1, o5)
    RD.cell(rv + 2, 1, o4)
    return f"v026_post：README A1／A2 版號、本版要點 README!A{rv}（v{PREV}、v0.2.4 說明下移）"

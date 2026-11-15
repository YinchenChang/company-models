# 世芯-KY v0.2.5（2026-11-15）後處理：滾入 3Q26 財報（FY2026 錨點改為 1–3Q26 財報實績＋4Q26 假設）。
# 數值改動全在 build_alchip.py In 頁 E 節與 Q 頁 E 欄（輸入與錨點）；本檔只改版號與版本說明文字，不改任何數值公式。
# 建檔程式在 v024_post.apply 之後、存檔前呼叫 apply(wb)。另一側讀數建置（COORD_RUN＝1）與 sens_alchip.py 建置（ALCHIP_SENS＝1）時不執行。
import os

VER, BUILT, PREV = "0.2.5", "2026-11-15", "0.2.4"
NOTE = ("滾入 3Q26 財報：FY2026 錨點由「3Q26 月營收實績＋指引」改為「1–3Q26 財報實績（3Q26＝前 9 月累計 − 1H26；公開資訊觀測站 2026-11-06：前 9 月營收 37,877,562、毛利 9,462,133、營業利益 6,643,602、"
        "稅前 7,900,866、歸屬母公司淨利 6,317,516 千元，EPS 75.94）＋4Q26 假設（營收＝10 月單月 102.60 億 × 3；非 Trainium 毛利率＝3Q26 實績推得；營業費用＝3Q26；營業外 0.40、稅率 20%）」；"
        "3Q26／4Q26 匯率改為 3Q26 新台幣營收 ÷ 美元營收 8.193 億。資產負債表日改 2026-09-30（公開資訊觀測站 t164sb03）：淨現金與投資 55.65 → 39.78 十億元（存貨與應收增加），stub 期 0.504 → 0.252 年（9-30 → 12-31，10 月股利仍在 stub 期扣除），stub FCFF 改為 4Q26 單季 FCFF（4Q 營業利益稅後＋4Q 營收 ×（折舊率 − 資本支出率）＋4Q 營運資金變動 0（維持 9-30 水準））；股數改 9-30 實際 86.35M（股本 863,496 千元；評價、EPS、Comparable 同一股數；FY2026 加權依 GDR 6-30 發行時點；2026 年 GDR 增資 4.00M 股，其餘約 1.07M 股來源未查明）。"
        "法說（3／2 奈米占 3Q26 營收 80%、4Q26 營收再創新高、毛利率與 3Q 相當接近）只列參考。公式與架構不變。")


def apply(wb):
    if os.environ.get("COORD_RUN") == "1" or "Coord" not in wb.sheetnames:
        return "v025_post：略過（另一側讀數建置或無 Coord 頁）"
    RD = wb["README"]
    a1 = RD["A1"].value
    head = a1.split("；", 1)[0]
    assert f"v{PREV}（AI 半導體" in head, head[:80]
    RD["A1"].value = head.replace(f"v{PREV}", f"v{VER}", 1) + a1[len(head):] + f"；v{VER}（{BUILT}）：滾入 3Q26 財報（FY2026 錨點＝1–3Q26 實績＋4Q26 假設）"
    a2 = str(RD["A2"].value)
    assert a2.startswith(f"版本：v{PREV}（"), a2[:40]
    RD["A2"].value = a2.replace(f"版本：v{PREV}（", f"版本：v{VER}（{BUILT}：滾入 3Q26 財報；前版 v{PREV}）；v{PREV}（", 1)
    # 本版要點（HTML「資料狀態」章讀 README v1.3 區最後一列）：本版說明在前，v0.2.4 說明移到下一列保留
    r13 = next(r for r in range(1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith("v1.3 展示層"))
    rv = r13 + 9
    old = str(RD.cell(rv, 1).value)
    assert old.startswith(f"v{PREV}（"), old[:40]
    assert RD.cell(rv + 1, 1).value in (None, ""), RD.cell(rv + 1, 1).value
    RD.cell(rv, 1, f"v{VER}（{BUILT}）：{NOTE}")
    RD.cell(rv + 1, 1, old)
    # stub 期 FCFF＝4Q26 單季（主 session 2026-11-15 決定；同 Cerebras v0.1.3）：說明文字
    STUB = "4Q26 單季 FCFF＝4Q26 營業利益 ×（1 − 稅率）＋4Q26 營收 ×（折舊率 − 資本支出率）＋0（4Q26 營運資金變動設為 0，維持 9-30 實績水準；Sens 另列依 FY2026 全年假設反推、釋出 22.28 的案例）"
    REPL = (("含 stub 期會計年度（FY2026＝日曆年 2026，錨點）FCFF × stub_frac", STUB),
            ("stub FCFF＝自訂（路徑 stub_fcff）", "stub FCFF＝" + STUB),
            ("stub FCFF＝各路徑第一年 FCFF × 本值（年度比例）", "v0.2.5 起 stub FCFF＝" + STUB + "；本值不進 stub FCFF（只供說明）"))
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and not v.startswith("v0."):
                    w = v
                    for a, b in REPL:
                        w = w.replace(a, b)
                    if w != v:
                        c.value = w; n += 1
    assert n >= 6, n
    return f"v025_post：stub 說明 {n} 格；README A1／A2 版號、本版要點 README!A{rv}（v{PREV} 說明移到 A{rv + 1}）"

# 創意 v0.3.5（2026-11-15）後處理：3Q26 滾入（損益已於 v0.2 滾入，只核對並把出處改公開資訊觀測站）＋資產負債表日改 2026-09-30。
# 數值改動全在 build_guc.py In 頁（debt、lease、cash、inv、mi、bs_date、q_sh）與 stub FCFF 公式（4Q26 單季）；本檔只改版號、本版要點與 stub 說明文字，不改任何數值公式。
# 建檔程式在 v034_post.apply 之後、存檔前呼叫 apply(wb)。另一側讀數建置（COORD_RUN＝1）與 sens_guc.py 建置（GUC_SENS＝1）時不執行。
import os

VER, BUILT, PREV = "0.3.5", "2026-11-15", "0.3.4"
NOTE = ("3Q26 損益已於 v0.2 滾入（3Q26＝前 9 月 − 1H26），本版只核對並把出處改公開資訊觀測站 t164sb04（115 年第 3 季：前 9 月營收 44,098,869、毛利 9,967,659、營業利益 6,051,687、稅前 6,306,235、歸屬母公司淨利 5,473,445 千元，EPS 40.84；3Q26 單季營收 18,754,500、毛利 3,860,335、營業利益 2,551,163、EPS 16.95；模型推得的 3Q 單季與財報單季差：毛利 −0.0008、營業費用 −0.0037、營業利益 +0.0029 十億元；3Q 營業外 0.074、所得稅 0.354 改財報單季值，2Q 為前 9 月 − 3Q − 1Q，前 9 月合計不變）。資產負債表日改 2026-09-30（公開資訊觀測站 t164sb03，口徑同前：現金＋短期投資）：現金 8.947＋短期投資 0.539（FVTPL 0.501＋攤銷後成本 0.038）＋非流動攤銷後成本 0.032 − 短期借款 0.150 − 租賃 0.261，無非控制權益；淨現金與投資 10.10 → 9.11 十億元（借款 0 → 0.150）。stub 期 0.504 → 0.252 年（9-30 → 12-31；2026 年股利已於 6-30 發放，stub 期不扣）；stub FCFF 改為 4Q26 單季 FCFF（4Q 營業利益稅後＋4Q 營收 ×（折舊率 − 資本支出率）＋4Q 營運資金變動 0；同世芯 v0.2.5）＝2.59（原 FY2026 FCFF × 0.504＝4.34）；淨現金滾到評價日 14.44 → 11.69。股數改 9-30 股本 1,340,119 千元 ÷ 面額 10 元＝134.0119M（原 134.01M）。DCF 目標價 1,707.49 → 1,684.95；加權目標價 2,852.38 → 2,841.08（上檔 −55.5% → −55.6%，現價 6,405），評等強力放空不變；另一側讀數（C4）2,602.99 → 2,591.69。sens.json、coord_alt.json 重跑。")


def apply(wb):
    if os.environ.get("COORD_RUN") == "1" or "Coord" not in wb.sheetnames:
        return "v035_post：略過（另一側讀數建置或無 Coord 頁）"
    RD = wb["README"]
    a1 = RD["A1"].value
    head = a1.split("；", 1)[0]
    assert f"v{PREV}（AI 半導體" in head, head[:80]
    RD["A1"].value = head.replace(f"v{PREV}", f"v{VER}", 1) + a1[len(head):] + f"；v{VER}（{BUILT}）：資產負債表日改 2026-09-30、3Q26 出處改公開資訊觀測站"
    a2 = str(RD["A2"].value)
    assert a2.startswith(f"版本：v{PREV}（"), a2[:40]
    RD["A2"].value = a2.replace(f"版本：v{PREV}（", f"版本：v{VER}（{BUILT}：資產負債表日 2026-09-30；前版 v{PREV}）；v{PREV}（", 1)
    r13 = next(r for r in range(1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith("v1.3 展示層"))
    rv = r13 + 9
    old = str(RD.cell(rv, 1).value)
    assert old.startswith(f"v{PREV}（"), old[:40]
    assert RD.cell(rv + 1, 1).value in (None, ""), RD.cell(rv + 1, 1).value
    RD.cell(rv, 1, f"v{VER}（{BUILT}）：{NOTE}")
    RD.cell(rv + 1, 1, old)
    STUB = "4Q26 單季 FCFF＝4Q26 營業利益 ×（1 − 稅率）＋4Q26 營收 ×（折舊率 − 資本支出率）＋0（4Q26 營運資金變動設為 0，維持 9-30 實績水準；同世芯 v0.2.5）"
    REPL = (("含 stub 期會計年度（FY2026＝日曆年 2026，錨點）FCFF × stub_frac", STUB),
            ("stub FCFF＝自訂（路徑 stub_fcff）", "stub FCFF＝" + STUB),
            ("stub FCFF＝各路徑第一年 FCFF × 本值（年度比例）", "v0.3.5 起 stub FCFF＝" + STUB + "；本值不進 stub FCFF（只供說明）"))
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
    assert n >= 1, n
    return f"v035_post：stub 說明 {n} 格；README A1／A2 版號、本版要點 README!A{rv}（v{PREV} 說明移到 A{rv + 1}）"

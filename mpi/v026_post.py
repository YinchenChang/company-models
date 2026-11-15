# 旺矽 v0.2.6（2026-11-15）後處理：滾入 3Q26 財報（FY2026 錨點＝1–3Q26 財報實績＋4Q26 假設）＋資產負債表日改 2026-09-30。
# 數值改動全在 build_mpi.py In 頁 E 節、資本結構列與 Q 頁 E 欄、stub FCFF 公式（4Q26 單季）；本檔只改版號、本版要點與 stub 說明文字，不改任何數值公式。
# 建檔程式在 v025_post.apply 之後、存檔前呼叫 apply(wb)。
import re

VER, BUILT, PREV = "v0.2.6", "2026-11-15", "v0.2.5"
NOTE = ("滾入 3Q26 財報（公開資訊觀測站 t164sb04／t164sb03，115 年第 3 季）：FY2026 錨點由「1H26 實績＋3Q26 營收實績＋公司說法」改為「1–3Q26 財報實績＋4Q26 假設」。3Q26 單季取財報值：營收 5.904、毛利 3.641（毛利率 61.7%，原假設 59.4%）、營業利益 2.276、歸屬淨利 1.880、EPS 19.18（稀釋 19.02）；2Q26 營收改為前 9 月 15.0705 − 3Q − 1Q。4Q26 沿用既有做法：營收季增 7.5%（2Q26 法說指引中點；10 月 22.32 億 × 3 較 3Q +13.4%，只列參考）、毛利率改取 3Q26 實績 61.7%（原 60.0%；公司稱 4Q 小升）、費用季增 6%、淨利 ÷ 營業利益 0.87。FY2026 EPS 68.70 → 69.17。資產負債表日改 2026-09-30（口徑同前）：現金 9.880 → 7.783、長期投資 0.912 → 1.282、借款 8.147 → 9.662（應付公司債 4.583〔旺矽六，面額 50 億元、0% 票面、2026-06-15～2031-06-15〕＋長期借款 5.051＋資料商差額 0.029〔stockanalysis Long-Term Debt 與 MOPS 公司債＋長借之差，性質未辨識〕，短借 1.350 → 0）、租賃 0.519 → 0.494；淨現金與投資 2.126 → −1.091。stub 期 0.504 → 0.252 年；8 月股利 2.16 已在 9-30 前發放，stub 期不再扣；stub FCFF 改 4Q26 單季（4Q 營運資金變動 0；同世芯 v0.2.5）＝1.41；淨現金滾到評價日 0.263 → 0.322。評價股數改 9-30 股本 ÷ 10＝97.98M；可轉債以帳面計入債務，轉換價格未取得，稀釋未計。DCF 目標價 1,357.53 → 1,347.75；Comparable 2,283.01 → 2,266.17；加權目標價 1,820.27 → 1,806.96（上檔 −51.1% → −51.5%，現價 3,725），評等強力放空不變；另一側讀數 1,598.85 → 1,586.53。")


def apply(wb):
    RD = wb["README"]
    a1 = RD["A1"].value
    a1, n = re.subn(rf"個股模型 {re.escape(PREV)}（", f"個股模型 {VER}（", a1, count=1)
    assert n == 1, a1[:80]
    RD["A1"].value = a1 + f"；{VER}（{BUILT}）：滾入 3Q26 財報、資產負債表日改 2026-09-30"
    r13 = next(r for r in range(1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith("v1.3 展示層"))
    rv = r13 + 9
    old = str(RD.cell(rv, 1).value)
    assert old.startswith(f"{PREV}（"), old[:40]
    assert RD.cell(rv + 1, 1).value in (None, ""), RD.cell(rv + 1, 1).value
    RD.cell(rv, 1, f"{VER}（{BUILT}）：{NOTE}")
    RD.cell(rv + 1, 1, old)
    STUB = "4Q26 單季 FCFF＝4Q26 營業利益 ×（1 − 稅率）＋4Q26 營收 ×（折舊率 − 資本支出率）＋0（4Q26 營運資金變動設為 0，維持 9-30 實績水準；同世芯 v0.2.5）"
    REPL = (("含 stub 期會計年度 FY2026（日曆年）FCFF × stub_frac", STUB),)
    k = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and not v.startswith("v0."):
                    w = v
                    for a, b in REPL:
                        w = w.replace(a, b)
                    if w != v:
                        c.value = w; k += 1
    return f"v026_post：stub 說明 {k} 格；README A1 版號、本版要點 README!A{rv}（{PREV} 說明移到 A{rv + 1}）"

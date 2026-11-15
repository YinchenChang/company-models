# 韓美半導體 v0.4.5（2026-11-15）後處理：兩讀數的另一側讀數口徑修正（tools/coord/README.md「兩讀數：另一側讀數的口徑」；同 Disco、TOWA v0.3.5）。
# 另一側建置（SIL_COORD_ALT＝1）的修正在 build_hanmi.py engine()（kfix：CY2026 購機讀未縮放的 CY2027 產能，校準係數 k 與現行建置相同）；
# 本檔只做現行建置的展示：README A1 版號與版本列、Coord 頁末新增「5. 兩讀數口徑」一節（口徑說明＋Sens 錨定案例重算 k 的口徑揭露）。
# 只新增列與改版號文字，既有數值格不變。build_hanmi.py 在 v044_post.apply 之後、存檔前呼叫 apply(wb, CR)。
import os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from c13 import T, F, REC5, BOLD, BLUE

VER, PREV, BUILT = "v0.4.5", "v0.4.4", "2026-11-15"
# v0.4.4 的另一側讀數（hanmi/coord_alt.json v0.4.4 版，tools/coord/alt_run.py；另一側建置以縮放後的 CY2026 購機重算 k）：只列口徑揭露
TP_RECAL = 26255.4096994615


def _row(ws, prefix, start=1, col=1):
    return next(r for r in range(start, ws.max_row + 1) if str(ws.cell(r, col).value or "").startswith(prefix))


def apply(wb, CR):
    PL, SE, CW, RD = wb["PL"], wb["Sens"], wb["Coord"], wb["README"]
    kB = _row(PL, "校準係數 k", _row(PL, "情境 B："))                     # 現行建置情境 B 的 k（三情境相同）
    kS = _row(SE, "校準係數 k", _row(SE, "案例區塊：Wattonomics 錨定"))   # Sens「Wattonomics 錨定」案例內重算的 k
    tp, rec = CR["tp"], CR["rec"]
    r = CW.max_row + 2; out = {}
    CW.cell(r, 1, f"5. 兩讀數口徑（{VER}；tools/coord/README.md「兩讀數：另一側讀數的口徑」）").font = BOLD; out["title"] = r; r += 1
    CW.cell(r, 1, F(T("另一側讀數口徑：不重算校準係數 k；已實現與錨點期間不受縮放影響。另一側建置的情境 B、C 於 CY2027 起 HBM 需求位元 × 另一側比值（出貨＝MIN(需求 × 比值, 供給)）；"
                      "CY2026 購機（FY2026 錨點期間，也是 k 的分母）看的 CY2027 產能讀未縮放值，k 與現行建置相同（"),
                    f'TEXT(PL!C{kB},"0.00")',
                    T(" 億韓元／百萬顆當量）；FY2025 實績與 FY2026 錨點營收不變，只有 FY2027 起的可得新機需求隨比值縮放。理由：k 的分母含 CY2026→CY2027 產能增量，縮放會使 k 虛增、方向反轉"
                      f"（{PREV} 以縮放後的 CY2026 購機重算 k，另一側目標價反而高於現行）。")))
    out["basis"] = r; r += 1
    CW.cell(r, 4, TP_RECAL).number_format = "#,##0"; CW.cell(r, 4).font = BLUE
    up = f"(D{r}/px-1)"
    CW.cell(r, 1, F(T("口徑揭露：Sens「Wattonomics 錨定」案例在案例內重算 k（案例 k "), f'TEXT(Sens!C{kS},"0.00")', T("，現行 "), f'TEXT(PL!C{kB},"0.00")',
                    T("，"), f'TEXT(Sens!C{kS}/PL!C{kB}-1,"+0%;-0%")', T("；案例縮放係數為 Wattonomics 錨定係數 wta，與另一側比值不同）。若另一側讀數照此重算 k，加權目標價為 "),
                    f'TEXT(D{r},"#,##0")', T(" 韓元（"), REC5(up), T("），本模型採用的不重算口徑為 "), f'TEXT(Coord!C{tp},"#,##0")', T(" 韓元（"), f"Coord!C{rec}",
                    T("），兩口徑差 "), f'TEXT(D{r}/Coord!C{tp}-1,"+0.0%;-0.0%")', T("、評等"), f'IF({REC5(up)}=Coord!C{rec},"相同","不同")',
                    T("。重算值只列揭露、不進兩讀數表；Sens 案例維持一次動一個輸入、案例內重算 k 的做法，不改。")))
    CW.cell(r, 8, f"D 欄：{PREV} 的另一側讀數（hanmi/coord_alt.json {PREV} 版，tools/coord/alt_run.py；另一側建置以縮放後的 CY2026 購機重算 k；Siliconomics v0.19 對帳中比值；Derived）加權目標價 26,255.41 韓元；本模型正式另一側讀數不採此口徑")
    out["recal"] = r; r += 1
    # README：A1 標題版號、末尾說明；v1.3 展示層版本列前加本版說明（HTML 版本列）
    a1 = RD["A1"].value
    head, sep, rest = a1.partition("；")
    head2, n = re.subn(re.escape(PREV) + "（", VER + "（", head, count=1)
    assert n == 1, head
    RD["A1"].value = head2 + sep + rest + (f"；{VER}（{BUILT}）：兩讀數的另一側讀數改為不重算校準係數 k（已實現與錨點期間不受縮放影響；tools/coord/README.md 兩讀數規格），"
                                           "Coord 頁新增兩讀數口徑與 Sens 錨定案例重算 k 的口徑揭露；現行建置數值、目標價與評等不變")
    r13 = _row(RD, "v1.3 展示層")
    rv = r13 + 9
    assert str(RD.cell(rv, 1).value).startswith(PREV + "（"), RD.cell(rv, 1).value
    RD.cell(rv, 1).value = (f"{VER}（{BUILT}）：兩讀數的另一側讀數改為不重算校準係數 k（CY2026 購機讀未縮放產能；FY2025 實績與 FY2026 錨點不受縮放影響），"
                            "另一側目標價由高於改為低於現行基準讀數；Coord 頁揭露 Sens 錨定案例重算 k 的差異；現行建置數值、目標價與評等不變。前一版 " + RD.cell(rv, 1).value)
    return out

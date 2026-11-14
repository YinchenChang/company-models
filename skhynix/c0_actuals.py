# SK 海力士 C0（實績）：公司實績 vs 層級隱含值（三層基準協調規則 v1.4.2 C0 第 2 點；v0.7.7，2026-11-14）。
# 由 tools/memory_model/v13_post.py 的 coord_c0 呼叫 add()；HTML 對帳章由 tools/memory_model/html_v13.py 呼叫 html_blocks()。
# 結論：公司未揭露 HBM 營收、HBM 占 DRAM 營收比、HBM 位元出貨（2026 年第一至第三季）→ HBM 列「無可比實績」；
# 已揭露的 DRAM（含 HBM）營收只做上限檢查（參考，不判定）。實績數字寫在本檔（藍字常數，出處見 G 欄），層級隱含值連到 Coord 既有層列。
import sys
from openpyxl.styles import Font, PatternFill

BOLD = Font(bold=True)
BLUE = Font(color="0000FF")
HDR = PatternFill("solid", fgColor="DDEBF7")

TITLE = "C0（實績）：公司實績 vs 層級隱含值"
LBL = dict(
    ledger="公司實績明細（2026 年已公布季度）",
    cmp="實績與層級隱含值比較（HBM 營收）",
    bound="上限檢查（參考，不判定）：層級隱含 HBM 營收 ÷ 公司揭露 DRAM（含 HBM）營收",
    sumr="C0（實績）摘要",
    ext="外部估計（參考，不判定）：IDC 市占對層級市占",
)


def _hdr(Co, r, vals):
    for i, v in enumerate(vals, 1):
        c = Co.cell(r, i, v)
        c.font = BOLD
        c.fill = HDR


def add(wb, Co, L, F, T, hdr_row=None):
    """附在 Coord 頁尾（不位移既有列）。L＝coord_c0 的列號表；F、T＝v13_post 的公式組字工具。回傳列號表。"""
    tol = L["tol"]
    rev_l = L["rev"][1]            # Coord：SK 海力士 HBM 營收，層（Siliconomics）列（US$B，C＝CY2026）
    A = {}
    r = Co.max_row + 2
    Co.cell(r, 1, TITLE + "（三層基準協調規則 v1.4.2 C0 第 2 點；v0.7.7 新增）：個股側只用公司已公布的實績或法定揭露，層級側是同口徑的隱含值；"
                  "個股自層級讀入的同源值已移到上方「讀取層級（同源，不判定）」，不再計入對帳中").font = BOLD
    A["top"] = r
    r += 1
    Co.cell(r, 1, "比較方法")
    Co.cell(r, 7, "主讀數＝2026 年已公布期間（1–9 月）累計實績 vs 層級 CY2026 × 9 ÷ 12（線性時差近似，Derived）；第二讀數＝最近一季 × 4 vs 層級 CY2026。"
                  "成長年中主讀數偏低、第二讀數偏高，兩者夾住真值。兩讀數差距都超過 ±15% 且同號＝對帳中；只有一個超過＝時差邊界（參考）；都在內＝一致。"
                  "沒有可比實績時寫「無可比實績」與缺的揭露項，不以模型預測或同源值補位。會計年度與日曆年相同，無須換月。")
    A["method"] = r
    r += 1
    Co.cell(r, 1, "容差（同上方 C0 表容差）")
    Co.cell(r, 2, f"=$B${tol}").number_format = "0%"
    A["tol"] = r
    r += 2
    # ---- 實績明細 ----
    Co.cell(r, 1, LBL["ledger"]).font = BOLD
    A["ledger"] = r
    r += 1
    _hdr(Co, r, ["項目", "單位", "2026Q1", "2026Q2", "2026Q3", "2026 年 1–9 月", "出處、口徑與證據等級", "網址（不入 HTML）"])
    A["lh"] = r
    r += 1
    # 營收（連 Q 頁既有實績輸入）
    Co.cell(r, 1, "合併營收（公司實績；Q 頁實績輸入）")
    Co.cell(r, 2, "兆韓元")
    for c, q in zip("CDE", "FGH"):
        Co[f"{c}{r}"] = f"=Q!{q}4"
        Co[f"{c}{r}"].number_format = "0.000"
    Co[f"F{r}"] = f"=C{r}+D{r}+E{r}"
    Co[f"F{r}"].number_format = "0.000"
    Co.cell(r, 7, "Verified。2026Q1 52,576,287 百萬韓元與 2026Q2 79,318,746 百萬韓元：SK hynix 6-K 半年報（2026-08-18）第 II.1.C.a 節；"
                  "2026Q3 1,046,155 億韓元：SK hynix 3Q26 財報新聞稿（2026-10-28）；Q 頁為三位小數；網址見 H 欄")
    Co.cell(r, 8, "https://www.sec.gov/Archives/edgar/data/2120882/000119312526354777/d147827d6k.htm ； https://news.skhynix.co.kr/q3-2026-business-results/")
    A["rev"] = r
    r += 1
    Co.cell(r, 1, "DRAM（含 HBM）占營收比（公司揭露）")
    Co.cell(r, 2, "%")
    for c, v in (("C", 0.773), ("D", 0.73), ("E", 0.75)):
        Co[f"{c}{r}"] = v
        Co[f"{c}{r}"].font = BLUE
        Co[f"{c}{r}"].number_format = "0.0%"
    Co.cell(r, 7, "2026Q1 77.3%：SK hynix 424B4（2026-07）第 18 頁（風險因素第一項），Verified；"
                  "2026Q2 73%：公司 2Q26 連結決算新聞稿（7-29），經 ITmedia（福田昭，2026-08-07，圖說）與 Blocks & Files（2026-07-29）引述，整數捨入、二手；"
                  "2026Q3 75%（NAND 24%）：BizWatch（2026-10-28）報導 3Q26 財報，整數捨入、二手（Q 頁的 75.7% 是推估 Derived，不當實績）。"
                  "三季口徑不完全相同（424B4 的 1Q 為 77.3%，2Q 簡報口徑下的 1Q 為 78%），故占比誤差約 ±1 個百分點；等級：Q1 Verified，Q2、Q3 為 Interested-party 二手；網址見 H 欄")
    Co.cell(r, 8, "https://www.sec.gov/Archives/edgar/data/2120882/000119312526299963/d32785d424b4.htm ； https://eetimes.itmedia.co.jp/ee/spv/2608/07/news029_2.html ； "
                  "https://www.blocksandfiles.com/flash/2026/07/29/sk-hynix-announces-extraordinarily-high-revenues-but-misses-expectations/5280499 ； "
                  "https://news.bizwatch.co.kr/article/industry/2026/10/28/0011/")
    A["dsh"] = r
    r += 1
    Co.cell(r, 1, "DRAM（含 HBM）營收＝營收 × 占比")
    Co.cell(r, 2, "兆韓元")
    Co["C%d" % r] = 40.659   # 424B4 直接揭露：1Q26 DRAM 營收 W40,659bn
    Co["C%d" % r].font = BLUE
    for c in "DE":
        Co[f"{c}{r}"] = f"={c}{A['rev']}*{c}{A['dsh']}"
    for c in "CDEF":
        Co[f"{c}{r}"].number_format = "0.0"
    Co[f"F{r}"] = f"=C{r}+D{r}+E{r}"
    Co.cell(r, 7, "2026Q1＝424B4 直接揭露的 DRAM 營收 40,659 十億韓元（Verified）；2026Q2、Q3＝營收 × 二手占比（Derived；整數捨入誤差約 ±0.5 個百分點 × 營收，Q2 約 ±0.4、Q3 約 ±0.5 兆韓元）。公司不在 DRAM 內拆 HBM：此為 DRAM 加 HBM 合計")
    A["dram"] = r
    r += 1
    miss = [("HBM 營收", "兆韓元", "424B4 與 6-K 均將 DRAM 營收不拆 HBM（Blocks & Files 2026-07-09 對 F-1 的說明相同）；3Q26 新聞稿只說 HBM4 量產出貨、第四季 HBM4 在出貨量與營收都將超過 HBM3E，沒有金額"),
            ("HBM 占 DRAM 營收比", "%", "公司未揭露；424B4、6-K、2Q26 與 3Q26 新聞稿與法說均無。二手報導對 2025 年的說法互相矛盾（有稱 HBM 超過 DRAM 營收一半，有稱 F-1 不拆分），無原文，不採用"),
            ("HBM 位元出貨", "EB", "公司只給 DRAM 位元季增（2Q26 高個位數、3Q26 約 10%、4Q26 指引中 20% 段），不分 HBM")]
    A["miss"] = []
    for nm, un, nt in miss:
        Co.cell(r, 1, f"{nm}（公司揭露）")
        Co.cell(r, 2, un)
        for c in "CDE":
            Co[f"{c}{r}"] = "未揭露"
        Co.cell(r, 7, nt)
        A["miss"].append(r)
        r += 1
    r += 1
    # ---- 比較 ----
    Co.cell(r, 1, LBL["cmp"]).font = BOLD
    A["cmp"] = r
    r += 1
    _hdr(Co, r, ["項目", "單位", "公司實績", "層級隱含", "差距（實績 ÷ 層級 − 1）", "判定", "口徑、來源與證據等級"])
    A["ch"] = r
    r += 1
    for key, nm, f, note in (
        ("m1", "主讀數：SK 海力士 HBM 營收（2026 年 1–9 月累計 vs 層級 CY2026 × 9 ÷ 12）", f"=C{rev_l}*9/12",
         "層級隱含＝Siliconomics 的 HBM 市場 × 本公司 HBM 營收市占（連上方 C0 表層列 CY2026）× 9 ÷ 12（Derived）；公司實績：無可比實績，缺 HBM 營收（見上方明細：公司不拆分揭露）"),
        ("m2", "第二讀數：SK 海力士 HBM 營收（2026Q3 × 4 vs 層級 CY2026）", f"=C{rev_l}",
         "層級隱含＝同上 CY2026 全年；公司實績：無可比實績，缺 HBM 營收")):
        Co.cell(r, 1, nm)
        Co.cell(r, 2, "US$B")
        Co[f"C{r}"] = "無可比實績"
        Co[f"D{r}"] = f
        Co[f"D{r}"].number_format = "0.0"
        Co[f"E{r}"] = f'=IF(ISNUMBER(C{r}),C{r}/D{r}-1,"無可比實績")'
        Co.cell(r, 7, note)
        A[key] = r
        r += 1
    m1, m2 = A["m1"], A["m2"]
    for k in (m1, m2):
        Co[f"F{k}"] = f'=IF(ISNUMBER(E{k}),IF(ABS(E{k})>$B${tol},"超過容差","在容差內"),"無可比實績")'
    Co.cell(r, 1, "判定（HBM 營收）")
    Co[f"C{r}"] = (f'=IF(COUNT(E{m1},E{m2})<2,"無可比實績",IF(AND(ABS(E{m1})>$B${tol},ABS(E{m2})>$B${tol},E{m1}*E{m2}>0),"對帳中",'
                   f'IF(OR(ABS(E{m1})>$B${tol},ABS(E{m2})>$B${tol}),"時差邊界（參考）","一致")))')
    Co.cell(r, 7, "規則：兩讀數都超過 ±15% 且同號＝對帳中；只一個超過＝時差邊界（參考）；都在內＝一致；公司沒有可比實績＝無可比實績（不計入對帳中）。"
                  "缺的揭露項：HBM 營收、HBM 占 DRAM 營收比、HBM 位元出貨（2026 年第一至第三季皆未揭露）")
    A["j"] = r
    r += 1
    Co.cell(r, 1, "對帳中項目數（C0 實績）")
    Co[f"C{r}"] = f'=IF(C{A["j"]}="對帳中",1,0)'
    A["n"] = r
    r += 2
    # ---- 上限檢查 ----
    Co.cell(r, 1, LBL["bound"]).font = BOLD
    Co.cell(r, 7, "公司揭露的是 DRAM（含 HBM）營收，HBM 營收必然小於它；用已揭露期間（2026 年 1–9 月；Q1 取 424B4，Q2、Q3 為二手占比）檢查層級隱含值是否超過上限。這只是合理性檢查，不是實績對帳，不判定高低")
    A["bound"] = r
    r += 1
    Co.cell(r, 1, "層級隱含 HBM 營收（CY2026 × 9 ÷ 12；韓元＝US$B × 匯率，匯率為 In 頁 fx，Derived）")
    Co.cell(r, 2, "兆韓元")
    Co[f"C{r}"] = f"=C{rev_l}*9/12*fx/1000"
    Co[f"C{r}"].number_format = "0.0"
    A["b1"] = r
    r += 1
    Co.cell(r, 1, "公司揭露 DRAM（含 HBM）營收（2026 年 1–9 月）")
    Co.cell(r, 2, "兆韓元")
    Co[f"C{r}"] = f"=F{A['dram']}"
    Co[f"C{r}"].number_format = "0.0"
    A["b2"] = r
    r += 1
    Co.cell(r, 1, "層級隱含 HBM 營收 ÷ DRAM（含 HBM）營收：1–9 月（主）")
    Co.cell(r, 2, "%")
    Co[f"C{r}"] = f"=C{A['b1']}/C{A['b2']}"
    Co[f"C{r}"].number_format = "0.0%"
    A["b3"] = r
    r += 1
    Co.cell(r, 1, "層級隱含 HBM 營收 ÷ DRAM（含 HBM）營收：2026Q3 × 4 vs 層級 CY2026（第二）")
    Co.cell(r, 2, "%")
    Co[f"C{r}"] = f"=C{rev_l}*fx/1000/(E{A['dram']}*4)"
    Co[f"C{r}"].number_format = "0.0%"
    A["b4"] = r
    r += 1
    Co.cell(r, 1, "上限檢查狀態")
    Co[f"C{r}"] = f'=IF(MAX(C{A["b3"]},C{A["b4"]})<=1,"通過：未超過 DRAM（含 HBM）營收","不通過：層級隱含 HBM 營收超過公司揭露的 DRAM 營收")'
    Co.cell(r, 7, "參考、不判定：公司未揭露 HBM 占 DRAM 的實際比例，無法判斷隱含比例偏高或偏低；若主控要驗證層級的 HBM 市場與市占，需新的揭露（例如法說給 HBM 占比或位元）")
    A["bs"] = r
    r += 2
    # ---- 外部估計（參考，不判定）----
    Co.cell(r, 1, LBL["ext"]).font = BOLD
    Co.cell(r, 7, "外部估計只列參考，不判定、不開對帳單（規則 v1.4.2）。IDC 的 1Q26 HBM 營收市占由 424B4 摘要引用（Interested-party）")
    A["ext"] = r
    r += 1
    Co.cell(r, 1, "IDC 2026Q1 HBM 營收市占：SK 海力士（424B4 引用）")
    Co.cell(r, 2, "%")
    Co[f"C{r}"] = 0.564
    Co[f"C{r}"].font = BLUE
    Co[f"C{r}"].number_format = "0.0%"
    A["e1"] = r
    r += 1
    Co.cell(r, 1, "層級：SK 海力士 HBM 營收市占 CY2026（上方 C0 表層列）")
    Co.cell(r, 2, "%")
    Co[f"C{r}"] = f"=C{L['sh'][1]}"
    Co[f"C{r}"].number_format = "0.0%"
    A["e2"] = r
    r += 1
    Co.cell(r, 1, "  差距：IDC ÷ 層級 − 1（參考，不判定）")
    Co.cell(r, 2, "%")
    Co[f"C{r}"] = f"=C{A['e1']}/C{A['e2']}-1"
    Co[f"C{r}"].number_format = "+0.0%;-0.0%"
    Co.cell(r, 7, "單季（1Q）對全年市占的時差與口徑差未調整；賣方與研調為參考")
    A["e3"] = r
    r += 2
    # ---- 摘要 ----
    Co.cell(r, 1, LBL["sumr"] + "（HTML 用；公式帶數字）")
    Co.cell(r, 2, F(T("公司未揭露 SK 海力士的 HBM 營收、HBM 占 DRAM 營收比與 HBM 位元出貨（2026 年第一至第三季），HBM 營收列判定為「"), f"C{A['j']}",
                    T("」，不計入對帳中（對帳中項目 "), f'C{A["n"]}', T(" 個）。已揭露的 DRAM（含 HBM）營收只做上限檢查：層級隱含 HBM 營收約為 DRAM（含 HBM）營收的 "),
                    f'TEXT(C{A["b3"]},"0%")', T("（1–9 月）與 "), f'TEXT(C{A["b4"]},"0%")', T("（2026Q3 × 4），"), f'C{A["bs"]}', T("（參考，不判定）。")))
    A["sum"] = r
    r += 1
    Co.cell(r, 1, "C0（實績）狀態旗標（核對：HBM 營收判定不是「對帳中」時為 0）")
    Co[f"C{r}"] = f'=IF(C{A["j"]}="對帳中",1,0)'
    A["flag"] = r
    return A


def html_blocks(CW, lab):
    """HTML 對帳章：C0（實績）區塊（數字、標籤全取自 xlsx；字串不含年份與日期以外的數字）。"""
    top = lab("Coord", TITLE)
    lh = lab("Coord", "項目", lab("Coord", LBL["ledger"]))
    ch = lab("Coord", "項目", lab("Coord", LBL["cmp"]))
    j = lab("Coord", "判定（HBM 營收）", ch)
    bd = lab("Coord", LBL["bound"])
    sm = lab("Coord", LBL["sumr"])
    rows = lambda r0, r1: [r for r in range(r0, r1) if CW.cell(r, 1).value]
    fm = lambda c, r, f: f"@Coord!{c}{r}" if isinstance(CW[f"{c}{r}"].value, str) else f"@Coord!{c}{r}|{f}"
    ledger_rows = []
    for r in range(lh + 1, ch - 1):
        if not CW.cell(r, 1).value or str(CW.cell(r, 1).value).startswith(LBL["cmp"]):
            continue
        u = CW.cell(r, 2).value
        f = {"%": "p1", "兆韓元": "n1", "EB": "n1"}.get(u, "n1")
        if str(CW.cell(r, 1).value).startswith("合併營收"):
            f = "n3"
        ledger_rows.append([f"@Coord!A{r}", f"@Coord!B{r}"] + [fm(c, r, f) for c in "CDEF"])
    cmp_rows = []
    for r in range(ch + 1, j):
        cmp_rows.append([f"@Coord!A{r}", f"@Coord!B{r}", fm("C", r, "n1"), fm("D", r, "n1"), fm("E", r, "sp1"), f"@Coord!F{r}"])
    cmp_rows.append([f"@Coord!A{j}", "—", f"@Coord!C{j}", "—", "—", "—"])
    bound_rows = []
    for r in range(bd + 1, lab("Coord", "上限檢查狀態") + 1):
        f = "p1" if CW.cell(r, 2).value == "%" else "n1"
        bound_rows.append([f"@Coord!A{r}", f"@Coord!B{r}", fm("C", r, f)])
    ex = lab("Coord", LBL["ext"])
    ext_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}|" + ("sp1" if "差距" in str(CW.cell(r, 1).value) else "p1")] for r in range(ex + 1, ex + 4)]
    notes = [r for r in range(lh + 1, lab("Coord", "上限檢查狀態") + 1) if CW.cell(r, 7).value] + [lab("Coord", LBL["ext"])]
    return [
        {"type": "h3", "text": "實績對帳：公司實績與層級隱含值"},
        {"type": "note", "parts": [f"@Coord!B{sm}"]},
        {"type": "table", "caption": f"@Coord!A{lab('Coord', LBL['ledger'])}", "head": [f"@Coord!{c}{lh}" for c in "ABCDEF"], "rows": ledger_rows},
        {"type": "table", "caption": f"@Coord!A{lab('Coord', LBL['cmp'])}", "head": [f"@Coord!{c}{ch}" for c in "ABCDEF"], "rows": cmp_rows},
        {"type": "table", "caption": f"@Coord!A{bd}", "head": ["項目", "單位", "值"], "rows": bound_rows},
        {"type": "table", "caption": f"@Coord!A{ex}", "head": ["項目", "單位", "值"], "rows": ext_rows},
        {"type": "table", "caption": "實績對帳的口徑、出處與證據等級", "collapse": True, "summary": "展開：實績出處、口徑與判定規則",
         "head": ["項目", "出處、口徑與證據等級"], "rows": [[f"@Coord!A{r}", f"@Coord!G{r}"] for r in notes]},
        {"type": "note", "parts": [f"@Coord!G{lab('Coord', '比較方法')}"], "collapse": True, "summary": "展開：比較方法"},
    ]

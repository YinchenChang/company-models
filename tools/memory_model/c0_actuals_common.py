# 記憶體樣板 C0（實績）共用實作（三星、美光 v0.5.5；SK 海力士 v0.7.7 的 skhynix/c0_actuals.py 為同結構的專屬版，未改動）。
# 三層基準協調規則 v1.4.2 C0 第 2 點：個股側只用公司已公布的實績或法定揭露，層級側為同口徑隱含值（HBM 市場 × 本公司 HBM 營收市占，連 Coord 層列）。
# 公司資料夾的 c0_actuals.py 提供 SPEC，並以 add／html_blocks 轉呼叫本模組；由 tools/memory_model/v13_post.py（xlsx）與 html_v13.py（HTML）載入。
from openpyxl.styles import Font, PatternFill

BOLD = Font(bold=True)
BLUE = Font(color="0000FF")
HDR = PatternFill("solid", fgColor="DDEBF7")
TITLE = "C0（實績）：公司實績 vs 層級隱含值"
LBL = dict(ledger="公司實績明細（2026 年已公布季度）", cmp="實績與層級隱含值比較（HBM 營收）",
           bound="上限檢查（參考，不判定）：層級隱含 HBM 營收 ÷ 公司揭露營收（含 HBM）", sumr="C0（實績）摘要")


def _hdr(Co, r, vals):
    for i, v in enumerate(vals, 1):
        c = Co.cell(r, i, v)
        c.font = BOLD
        c.fill = HDR


def add(SP, wb, Co, L, F, T):
    tol, rev_l = L["tol"], L["rev"][1]
    cur, mult = SP["cur"], SP["mult"]          # mult：US$B → 公司報表單位的乘數公式片段（例 "*fx/1000"；美元公司為空）
    A = {}
    r = Co.max_row + 2
    Co.cell(r, 1, TITLE + "（三層基準協調規則 v1.4.2 C0 第 2 點；v0.5.5 新增）：個股側只用公司已公布的實績或法定揭露，層級側是同口徑的隱含值；"
                  "個股自層級讀入的同源值已在上方「讀取層級（同源，不判定）」，不計入對帳中").font = BOLD
    A["top"] = r
    r += 1
    Co.cell(r, 1, "比較方法")
    Co.cell(r, 7, f"主讀數＝2026 年已公布期間（{SP['span']}）累計實績 vs 層級 CY2026 × 9 ÷ 12（線性時差近似，Derived）；第二讀數＝{SP['last']} × 4 vs 層級 CY2026。"
                  "最近一季落在年度後半，成長年中第二讀數偏高、主讀數偏低（偏誤方向依季中點，不得寫成夾住真值）。兩讀數差距都超過 ±15% 且同號＝對帳中；只有一個超過＝時差邊界（參考）；都在內＝一致。"
                  "沒有可比實績時寫「無可比實績」與缺的揭露項，不以模型預測或同源值補位。" + SP["fy_note"])
    A["method"] = r
    r += 1
    Co.cell(r, 1, "容差（同上方 C0 表容差）")
    Co.cell(r, 2, f"=$B${tol}").number_format = "0%"
    A["tol"] = r
    r += 1
    Co.cell(r, 1, "同源檢查（層級份額與市場是否由本公司模型校準）")
    Co.cell(r, 7, SP["same_src"])
    A["same"] = r
    r += 2
    Co.cell(r, 1, LBL["ledger"]).font = BOLD
    A["ledger"] = r
    r += 1
    _hdr(Co, r, ["項目", "單位"] + SP["qcols"] + [SP["sumcol"], "出處、口徑與證據等級", "網址（不入 HTML）"])
    A["lh"] = r
    r += 1
    for row in SP["rows"]:
        Co.cell(r, 1, row["label"])
        Co.cell(r, 2, row["unit"])
        for c, v in zip("CDE", row["cells"]):
            Co[f"{c}{r}"] = v
            if not (isinstance(v, str) and v.startswith("=")):
                Co[f"{c}{r}"].font = BLUE
            Co[f"{c}{r}"].number_format = row["fmt"]
        Co[f"F{r}"] = f"=C{r}+D{r}+E{r}"
        Co[f"F{r}"].number_format = row["fmt"]
        Co.cell(r, 7, row["note"])
        Co.cell(r, 8, row["url"])
        A[row["key"]] = r
        r += 1
    A["miss"] = []
    for mi in SP["missing"]:   # (名稱, 單位, 說明[, 網址（H 欄，不入 HTML）])
        nm, un, nt = mi[:3]
        if len(mi) > 3:
            Co.cell(r, 8, mi[3])
        Co.cell(r, 1, f"{nm}（公司揭露）")
        Co.cell(r, 2, un)
        for c in "CDE":
            Co[f"{c}{r}"] = "未揭露"
        Co.cell(r, 7, nt)
        A["miss"].append(r)
        r += 1
    r += 1
    Co.cell(r, 1, LBL["cmp"]).font = BOLD
    A["cmp"] = r
    r += 1
    _hdr(Co, r, ["項目", "單位", "公司實績", "層級隱含", "差距（實績 ÷ 層級 − 1）", "判定", "口徑、來源與證據等級"])
    A["ch"] = r
    r += 1
    nm = SP["name"]
    for key, t, f, note in (
        ("m1", f"主讀數：{nm} HBM 營收（{SP['span']}累計 vs 層級 CY2026 × 9 ÷ 12）", f"=C{rev_l}*9/12",
         "層級隱含＝Siliconomics 的 HBM 市場 × 本公司 HBM 營收市占（連上方 C0 表層列 CY2026）× 9 ÷ 12（Derived）；公司實績：無可比實績，缺 HBM 營收（見上方明細）"),
        ("m2", f"第二讀數：{nm} HBM 營收（{SP['last']} × 4 vs 層級 CY2026）", f"=C{rev_l}", "層級隱含＝同上 CY2026 全年；公司實績：無可比實績，缺 HBM 營收")):
        Co.cell(r, 1, t)
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
    Co.cell(r, 7, "規則：兩讀數都超過 ±15% 且同號＝對帳中；只一個超過＝時差邊界（參考）；都在內＝一致；公司沒有可比實績＝無可比實績（不計入對帳中）。缺的揭露項：" + SP["miss_txt"])
    A["j"] = r
    r += 1
    Co.cell(r, 1, "對帳中項目數（C0 實績）")
    Co[f"C{r}"] = f'=IF(C{A["j"]}="對帳中",1,0)'
    A["n"] = r
    r += 2
    Co.cell(r, 1, LBL["bound"]).font = BOLD
    Co.cell(r, 7, f"公司揭露的是{SP['den_name']}，HBM 營收必然小於它；用已揭露期間檢查層級隱含值是否超過上限。這只是合理性檢查，不是實績對帳，不判定高低")
    A["bound"] = r
    r += 1
    Co.cell(r, 1, f"層級隱含 HBM 營收（CY2026 × 9 ÷ 12；{SP['conv']}）")
    Co.cell(r, 2, cur)
    Co[f"C{r}"] = f"=C{rev_l}*9/12{mult}"
    Co[f"C{r}"].number_format = "0.0"
    A["b1"] = r
    r += 1
    Co.cell(r, 1, f"公司揭露：{SP['den_name']}（{SP['span']}）")
    Co.cell(r, 2, cur)
    Co[f"C{r}"] = f"=F{A[SP['den_key']]}"
    Co[f"C{r}"].number_format = "0.0"
    A["b2"] = r
    r += 1
    Co.cell(r, 1, "層級隱含 HBM 營收 ÷ 公司揭露營收：已公布期間累計（主）")
    Co.cell(r, 2, "%")
    Co[f"C{r}"] = f"=C{A['b1']}/C{A['b2']}"
    Co[f"C{r}"].number_format = "0.0%"
    A["b3"] = r
    r += 1
    Co.cell(r, 1, f"層級隱含 HBM 營收 ÷ 公司揭露營收：{SP['last']} × 4 vs 層級 CY2026（第二）")
    Co.cell(r, 2, "%")
    Co[f"C{r}"] = f"=C{rev_l}{mult}/(E{A[SP['den_key']]}*4)"
    Co[f"C{r}"].number_format = "0.0%"
    A["b4"] = r
    r += 1
    Co.cell(r, 1, "上限檢查狀態")
    Co[f"C{r}"] = f'=IF(MAX(C{A["b3"]},C{A["b4"]})<=1,"通過：未超過公司揭露營收","不通過：層級隱含 HBM 營收超過公司揭露營收")'
    Co.cell(r, 7, "參考、不判定：公司未揭露 HBM 占比，無法判斷隱含比例偏高或偏低；要驗證層級 HBM 市場與市占，需新的揭露（HBM 營收、占比或位元）")
    A["bs"] = r
    r += 2
    Co.cell(r, 1, LBL["sumr"] + "（HTML 用；公式帶數字）")
    Co.cell(r, 2, F(T(f"公司未揭露{nm}的 HBM 營收、HBM 占營收比與 HBM 位元出貨（2026 年已公布季度），HBM 營收列判定為「"), f"C{A['j']}",
                    T("」，不計入對帳中（對帳中項目 "), f'C{A["n"]}', T(f" 個）。已揭露的{SP['den_name']}只做上限檢查：層級隱含 HBM 營收約為其 "),
                    f'TEXT(C{A["b3"]},"0%")', T("（已公布期間累計）與 "), f'TEXT(C{A["b4"]},"0%")', T(f"（{SP['last']} × 4），"), f'C{A["bs"]}', T("（參考，不判定）。")))
    A["sum"] = r
    r += 1
    Co.cell(r, 1, "C0（實績）狀態旗標（核對：HBM 營收判定不是「對帳中」時為 0）")
    Co[f"C{r}"] = f'=IF(C{A["j"]}="對帳中",1,0)'
    A["flag"] = r
    return A


def html_blocks(SP, CW, lab):
    top = lab("Coord", TITLE)
    lh = lab("Coord", "項目", lab("Coord", LBL["ledger"]))
    ch = lab("Coord", "項目", lab("Coord", LBL["cmp"]))
    j = lab("Coord", "判定（HBM 營收）", ch)
    bd = lab("Coord", LBL["bound"])
    sm = lab("Coord", LBL["sumr"])
    fm = lambda c, r, f: f"@Coord!{c}{r}" if isinstance(CW[f"{c}{r}"].value, str) else f"@Coord!{c}{r}|{f}"
    ledger_rows = []
    for r in range(lh + 1, ch - 1):
        if not CW.cell(r, 1).value:
            continue
        u = CW.cell(r, 2).value
        f = SP["ledger_fmt"].get(u, "n1")
        ledger_rows.append([f"@Coord!A{r}", f"@Coord!B{r}"] + [fm(c, r, f) for c in "CDEF"])
    cmp_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", fm("C", r, "n1"), fm("D", r, "n1"), fm("E", r, "sp1"), f"@Coord!F{r}"] for r in range(ch + 1, j)]
    cmp_rows.append([f"@Coord!A{j}", "—", f"@Coord!C{j}", "—", "—", "—"])
    bound_rows = []
    for r in range(bd + 1, lab("Coord", "上限檢查狀態") + 1):
        bound_rows.append([f"@Coord!A{r}", f"@Coord!B{r}", fm("C", r, "p1" if CW.cell(r, 2).value == "%" else "n1")])
    notes = [r for r in range(lh + 1, lab("Coord", "上限檢查狀態") + 1) if CW.cell(r, 7).value] + [lab("Coord", "同源檢查")]
    return [
        {"type": "h3", "text": "實績對帳：公司實績與層級隱含值"},
        {"type": "note", "parts": [f"@Coord!B{sm}"]},
        {"type": "table", "caption": f"@Coord!A{lab('Coord', LBL['ledger'])}", "head": [f"@Coord!{c}{lh}" for c in "ABCDEF"], "rows": ledger_rows},
        {"type": "table", "caption": f"@Coord!A{lab('Coord', LBL['cmp'])}", "head": [f"@Coord!{c}{ch}" for c in "ABCDEF"], "rows": cmp_rows},
        {"type": "table", "caption": f"@Coord!A{bd}", "head": ["項目", "單位", "值"], "rows": bound_rows},
        {"type": "table", "caption": "實績對帳的口徑、出處與證據等級", "collapse": True, "summary": "展開：實績出處、口徑、同源檢查與判定規則",
         "head": ["項目", "出處、口徑與證據等級"], "rows": [[f"@Coord!A{r}", f"@Coord!G{r}"] for r in notes]},
        {"type": "note", "parts": [f"@Coord!G{lab('Coord', '比較方法')}"], "collapse": True, "summary": "展開：比較方法"},
    ]

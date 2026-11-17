# 英特爾 v0.1.0（2026-11-17）：個股 HTML 版面規格 v1.3（docs/plan/個股HTML版面規格.md；樣板 聯電 v0.2.4、台積電 v0.5.5）＋三層協調 C0。
# 用法：python3 intc/gen_html_json.py [成品 xlsx]（省略時取 intc/dist/ 內唯一的 xlsx；寫 intc/html.json；勿手改 html.json）。
# 與聯電不同：本檔不經 html_src.json，直接依 xlsx 的標籤找列號、組出 v1.3 章節（命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態）。
# 數字與白話文字一律取自 xlsx（README 頁「展示層」「命題回覆」「v1.3 展示層」、PL 頁末標準損益表與分部營收推導、RevDCF 頁、Val 0.7、Sens 頁末相對基準表、Coord 頁）。
# 本檔文字不含年份與日期以外的數字（建置時由 build_company_html 檢查）。
import copy, glob, json, os, re, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import relayout as RL  # noqa: E402
import coord as COORD  # noqa: E402

X = sys.argv[1] if len(sys.argv) > 1 else (lambda f: f[0] if len(f) == 1 else sys.exit(f"dist 內應剛好一個 xlsx：{f}"))(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))
WB = openpyxl.load_workbook(X, data_only=True)
W = RL._WB(X)


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def row_of(sh, prefix, need_num=False, exact=False, start=1):
    ws = WB[sh]
    for r in range(start, ws.max_row + 1):
        t = str(ws.cell(r, 1).value or "").strip()
        if (t == prefix if exact else t.startswith(prefix)) and (not need_num or num(ws.cell(r, 2).value)):
            return r
    raise KeyError((sh, prefix))


I = {str(WB["In"].cell(r, 1).value): r for r in range(4, WB["In"].max_row + 1) if WB["In"].cell(r, 1).value}


def in_tok(prefix):
    ws = WB["In"]
    for r in range(4, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "")
        if a.startswith(prefix) and ws.cell(r, 2).value in (None, ""):
            return re.split(r"[.\s]", a)[0]
    raise KeyError(prefix)


# ---------------- 頂層與橫幅 ----------------
cfg = {"company": "英特爾", "ticker": "INTC", "thesis": None,
       "question": "Intel Foundry 的虧損收斂與資本支出強度，能否撐起現價隱含的成長；Intel Products 的現金流與 AI 需求能撐多少；現價隱含的倍數與同業、與英特爾自己的歷史差多少",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守）"}],
       "sections": []}
lead, props, obs, ans = RL._readme(W)
cfg["thesis"] = lead
cfg["wide_text"] = True
cfg["hero"] = RL._hero(W)
cfg["layout"] = "v1.3"
tpw, dcft, comp, compadj, cons = (row_of("Val", k, need_num=True) for k in ("目標價（DCF 與 Comparable 加權）", "DCF 目標價（情境加權）", "Comparable 目標價（市場主流同業法）", "Comparable 目標價（相對倍數調整後）", "市場共識目標價"))
upw, upc, dcf12, px = (row_of("Val", k, need_num=True) for k in ("上檔空間：加權目標價", "上檔空間：Comparable", "DCF 情境價值推到 12 個月後", "現價"))
rec, grade = row_of("Val", "投資建議", exact=True), row_of("Val", "上檔空間級距")
h = cfg["hero"]
h["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{comp}|n2", ["上檔空間 ", f"@Val!B{upc}|sp1"]]
for t in h["tiles"]:
    t[0] = [t[0], "（美元）"] if isinstance(t[0], str) else list(t[0]) + ["（美元）"]
if abs(WB["Val"].cell(compadj, 2).value - WB["Val"].cell(comp, 2).value) > 1e-9:
    h["note"] = (h.get("note") or []) + ["；進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{compadj}|n2", " 美元。"]
old_note = list(h.get("note") or [])

# ---------------- 先放章節骨架，讓 COORD.inject_html 加橫幅一格與估值章兩讀數區 ----------------
IDS = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
TITLES = {"thesis": "命題與核心問題", "valuation": "估值", "sens": "敏感度", "revdcf": "反向 DCF：現價隱含什麼", "pl": "損益（標準損益表）", "drivers": "驅動因子：分部營收推導", "recon": "對帳：三層協調、公司實績與共識對照", "status": "資料狀態"}
S = {k: {"id": k, "title": TITLES[k], "blocks": []} for k in IDS}
cfg["sections"] = [S[k] for k in IDS]
COORD.inject_html(cfg, X)
coord_blocks = S["valuation"]["blocks"][:4]
assert [b["type"] for b in coord_blocks] == ["h3", "note", "table", "bars"], [b["type"] for b in coord_blocks]
S["valuation"]["blocks"] = []
note_tail = list(cfg["hero"]["note"])[-2:]   # " "、Coord 橫幅揭露
RD, PLW, RVW, CW, SEW, VW = WB["README"], WB["PL"], WB["RevDCF"], WB["Coord"], WB["Sens"], WB["Val"]
r13 = row_of("README", "v1.3 展示層")
RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
assert str(RD[f"A{RM['ver']}"].value).startswith("v0.1.0"), RD[f"A{RM['ver']}"].value
cfg["hero"]["note"] = [f"@README!A{RM['hero']}"] + note_tail
cr = lambda p, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))

# ================= 1. 命題 =================
sg = row_of("PL", "分部營收推導（情境 B")
hS = sg + 1
ex = lambda label, start=sg: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == label)
r_ccg, r_dnai, r_dai, r_oth, r_ext, r_fm, r_rs, r_sh, r_cx, r_cxs, r_fndoi, r_dcr, r_dcm, r_ccm = (ex(t) for t in (
    "CCPG 營收", "DCAI 其他（非 AI 相關）營收", "DCAI 中 AI 相關營收", "All Other 營收", "外部代工營收", "Intel Foundry 營業利益率", "營收", "AI 相關 ÷ 營收", "總資本支出", "總資本支出 ÷ 營收",
    "Intel Foundry 營業利益（虧損）", "DCAI 營收", "DCAI 營業利益率", "CCPG 營業利益率"))
chart1 = {"type": "theme", "title": "營收分部（情境 B，十億美元）與 Intel Foundry 營業利益率：FY2026E–FY2029E", "x": f"@PL!F{hS}:I{hS}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.5,
          "total": f"@PL!F{r_rs}:I{r_rs}",
          "bars": [{"name": [f"@PL!A{r_ccg}"], "y": f"@PL!F{r_ccg}:I{r_ccg}", "color": "#6fc7cb"}, {"name": [f"@PL!A{r_dnai}"], "y": f"@PL!F{r_dnai}:I{r_dnai}", "color": "#10a0a7"},
                   {"name": [f"@PL!A{r_dai}"], "y": f"@PL!F{r_dai}:I{r_dai}", "color": "#0b7a80"}, {"name": [f"@PL!A{r_oth}"], "y": f"@PL!F{r_oth}:I{r_oth}", "color": "#c9c9c9"},
                   {"name": [f"@PL!A{r_ext}"], "y": f"@PL!F{r_ext}:I{r_ext}", "color": "#c29a2b"}],
          "lines": [{"name": [f"@PL!A{r_fm}"], "y": f"@PL!F{r_fm}:I{r_fm}", "axis": "r", "color": "#b4472f", "labels": True}],
          "caption": "堆疊＝情境 B（基準）的 CCPG、DCAI 其他、DCAI 中 AI 相關、All Other 與外部代工營收；頂端數字＝營收；折線＝Intel Foundry 營業利益率（右軸，負值＝虧損）。FY2026 為前三季實績加第四季指引的錨點。"}
S["thesis"]["blocks"] = [chart1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]}, {"type": "h3", "text": "命題"}, {"type": "list", "ordered": True, "items": props},
                         {"type": "list", "ordered": True, "collapse": True, "summary": "展開：命題回覆（依本模型計算的答案；想先看結論可展開，或讀完全文再看）", "items": ans},
                         {"type": "h3", "text": "重點觀察"}, {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}]

# ================= 2. 估值 =================
g0 = row_of("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = g0 + 7
bars4 = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n1",
         "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                   ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
         "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
        "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_px}", f"@Val!B{r_px}|n2", "", f"@Val!D{r_px}"]]}
rules = {"type": "note", "parts": old_note[:-0] if False else ["投資建議依加權目標價判定：", f"@Val!E{rec}"]}
hc = row_of("README", "展示層輔助：Comparable 目標價（相對倍數調整後）的上檔空間")
c0 = row_of("Val", "0.3 結論")
rk = row_of("Val", "股權成本 ke")
r_pvx, r_dcf = row_of("Val", "明確期 FCFF 現值"), row_of("Val", "DCF 目標價（情境加權）")
assert r_pvx == rk + 6, (r_pvx, rk)
_fm = lambda lab: "n3" if lab.startswith("評價用") else "p1" if "終值現值占 EV" in lab else "x1" if "隱含終值" in lab else "n2" if lab.startswith(("DCF 情境", "DCF 目標價")) else "n1"
fb = next(r for r in range(row_of("DCF", "情境 B（基準）"), WB["DCF"].max_row + 1) if WB["DCF"].cell(r, 1).value == "營收")
H5 = ["項目", "數值", "", "", "說明"]
kv2tbl = lambda items, cap, **kw: dict({"type": "table", "caption": cap, "head": ["項目", "數值"], "rows": [[k, v] for k, v in items]}, **kw)
r_c = row_of("Val", "同業前瞻 P/E 中位數")
r_peer_h = row_of("Val", "同業與前瞻 P/E")
rm0 = row_of("Val", "0.5 相對倍數調整")
rm_items = [(f"@Val!A{r}", f"@Val!B{r}|{f}") for r, f in ((row_of("Val", "本公司 10 年中位數"), "n2"), (row_of("Val", "同業 10 年中位數的中位數"), "n2"), (row_of("Val", "原始比值"), "n2"),
                                                          (row_of("Val", "相對倍數調整係數"), "n2"), (row_of("Val", "Comparable 目標價（相對倍數調整後）"), "n2"))]
pxr = row_of("Val", "1. 價格與市場隱含倍數")
px_items = [(f"@Val!A{r}", f"@Val!B{r}|x1" if r != pxr + 2 else f"@Val!B{r}|n2") for r in range(pxr + 2, pxr + 11) if WB["Val"].cell(r, 2).value is not None and WB["Val"].cell(r, 1).value]
rule_tbls = [
    {"type": "table", "collapse": True, "summary": "展開：規則值隱含股價（FY2027 EPS × 依循環位置的倍數）", "head": [f"@Val!{c}{pxr + 19}" for c in "ABDEFG"][:0] or ["情境", "FY2027 EPS", "規則倍數"],
     "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!D{r}|x1"] for r in range(row_of("Val", "3. 規則值隱含股價") + 2, row_of("Val", "3. 規則值隱含股價") + 5)]},
    {"type": "table", "collapse": True, "summary": "展開：循環位置（本模型營收最高年）", "head": ["情境", "營收最高年", "判定", "規則倍數"],
     "rows": [[f"@Val!A{r}", f"@Val!B{r}|yr", f"@Val!C{r}", f"@Val!D{r}|x1"] for r in range(row_of("Val", "2. 循環位置") + 2, row_of("Val", "2. 循環位置") + 5)]}]
r5 = row_of("Val", "5. 十年 P/E")
hist_tbl = {"type": "table", "collapse": True, "summary": "展開：十年年底 P/E（有效年度；淨利為負的年度剔除）", "head": ["年底", "平均 P/E", "來源"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|x1", f"@Val!C{r}|x1"] for r in range(r5 + 2, r5 + 12)]}
S["valuation"]["blocks"] = [bars4, tbl4, rules] + coord_blocks + [
    {"type": "table", "caption": "目標價結論（主要項目）", "head": H5, "rows": [
        [f"@Val!A{tpw}", f"@Val!B{tpw}|n2", "", "", f"@Val!E{tpw}"], [f"@Val!A{dcft}", f"@Val!B{dcft}|n2", "", "", f"@Val!E{dcft}"],
        [f"@Val!A{compadj}", f"@Val!B{compadj}|n2", "", "", f"@Val!E{compadj}"], [f"@Val!A{comp}", f"@Val!B{comp}|n2", "", "", f"@Val!E{comp}"],
        [f"@Val!A{px}", f"@Val!B{px}|n2", "", "", ""], [f"@Val!A{upw}", f"@Val!B{upw}|sp1", "", "", f"@Val!E{upw}"], [f"@Val!A{upc}", f"@Val!B{upc}|sp1", "", "", f"@Val!E{upc}"],
        ["Comparable（相對倍數調整後）的上檔空間", f"@README!B{hc}|sp1", "", "", ""], [f"@Val!A{rec}", f"@Val!B{rec}", "", "", f"@Val!E{rec}"],
        [f"@Val!A{dcf12}", f"@Val!B{dcf12}|n1", f"@Val!C{dcf12}|n1", f"@Val!D{dcf12}|n1", ""], [f"@Val!A{cons}", f"@Val!B{cons}|n2", "", "", f"@Val!E{cons}"]]},
    {"type": "table", "collapse": True, "summary": "展開：結論的其他項目（與共識的差距拆解、價格帶、十年分位、上檔空間級距）", "head": H5,
     "rows": [[f"@Val!A{r}", f"@Val!B{r}|{f}", f"@Val!C{r}|{f}" if r == c0 + 10 else "", f"@Val!D{r}|{f}" if r == c0 + 10 else "", f"@Val!E{r}"]
              for r, f in ((c0 + 7, "sp1"), (c0 + 8, "sp1"), (c0 + 9, ""), (c0 + 10, "p0"))] + [[f"@Val!A{grade}", f"@Val!B{grade}", "", "", f"@Val!E{grade}"]]},
    {"type": "table", "caption": "DCF 折現率（WACC、股權成本、稅後債務成本、RONIC、永續成長率；直接顯示）", "head": ["項目", "數值", "說明"], "rows": [[f"@Val!A{r}", f"@Val!B{r}|p2", f"@Val!E{r}"] for r in range(rk, rk + 6)]},
    {"type": "table", "caption": "DCF 彙總（標準版：無槓桿自由現金流 FCFF＋WACC；三情境）", "head": ["項目", "A（樂觀）", "B（基準）", "C（保守）", "說明"], "rows": [
        [f"@Val!A{r}", f"@Val!B{r}|{fm}", f"@Val!C{r}|{fm}" if r != r_dcf else "", f"@Val!D{r}|{fm}" if r != r_dcf else "", f"@Val!E{r}"]
        for r, fm in ((r, _fm(str(WB["Val"].cell(r, 1).value).strip())) for r in range(r_pvx, r_dcf + 1))]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B（基準）FCFF 逐年明細（DCF 頁；十億美元）", "head": ["@DCF!A5", "@DCF!C5", "@DCF!D5", "@DCF!E5", "@DCF!F5", "@DCF!G5", "@DCF!H5"],
     "rowgen": {"sheet": "DCF", "rows": f"{fb}-{fb + 10}", "cols": ["A", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n1", "n1", "n1", "n1", "n1", "n1"],
                "row_fmt": {str(fb + 3): {c: "p1" for c in "CDEFGH"}, str(fb + 5): {c: "p1" for c in "CDEFGH"}, str(fb + 6): {c: "p1" for c in "CDEFGH"}}}},
    kv2tbl([(f"@Val!A{r_c + k}", f"@Val!B{r_c + k}|{f}") for k, f in ((0, "x1"), (1, "n2"), (2, "n1"))], "Comparable（同業前瞻 P/E 中位數 × 一年後的前瞻 EPS）"),
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻倍數", "caption": f"@Val!A{r_peer_h - 1}", "head": [f"@Val!{c}{r_peer_h}" for c in "ABCD"] + [f"@Val!F{r_peer_h}"],
     "rowgen": {"sheet": "Val", "rows": f"{r_peer_h + 1}-{r_c - 1}", "cols": ["A", "B", "C", "D", "F"], "fmts": ["", "", "x1", "", "tag"]}},
    kv2tbl(rm_items, "相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司相對同業的長期倍數比）"),
    {"type": "note", "parts": [f"@Val!A{row_of('Val', '目標價方法', exact=True)}", "：", f"@Val!B{row_of('Val', '目標價方法', exact=True)}"]},
    kv2tbl(px_items, "價格與市場隱含倍數", collapse=True, summary="展開：現價與市場隱含倍數"),
    {"type": "h3", "text": "十年 P/E 規則值（舊法參考，不進目標價）"}] + rule_tbls + [{"type": "note", "parts": ["@Val!A2"]}, hist_tbl, {"type": "note", "parts": [f"@Val!A{row_of('Val', '規則（')}"]}]

# ================= 3. 敏感度 =================
t0 = row_of("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
while SEW.cell(r2 + 1, 1).value and not str(SEW.cell(r2 + 1, 1).value).startswith("核對"):
    r2 += 1
tor = []
for col, side in (("B", "低"), ("C", "高")):
    big = [r for r in range(r1, r2 + 1) if num(SEW[f"{col}{r}"].value) and abs(SEW[f"{col}{r}"].value) >= 0.1]
    small = [r for r in range(r1, r2 + 1) if num(SEW[f"{col}{r}"].value) and 0.0005 <= abs(SEW[f"{col}{r}"].value) < 0.1]
    base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
    tor.append((dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))),
                dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True, summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例（幾乎為零者見全部案例表）")))
all_tbl = {"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案與設定）",
           "head": [f"@Sens!{c}{t0 + 1}" for c in "ABCD"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1", f"@Sens!D{r}"] for r in range(r1, r2 + 1)]}
s_chk = row_of("Sens", "檢查：閉式公式"); s_note = row_of("Sens", "讀法：")
eps_tor = {"type": "tornado", "sheet": "Sens", "rows": f"5-{s_chk - 1}", "label_col": "A", "value_col": "F", "fmt": "sp1", "collapse": True, "summary": "展開：對照：FY2028 EPS 相對基準的變動（情境 A）", "title": "FY2028 EPS 相對基準的變動（情境 A）"}
full_tbl = {"type": "table", "collapse": True, "summary": "展開：全部案例的 EPS、DCF、Comparable 與加權目標價",
            "head": [f"@Sens!{c}3" for c in "ABCDEFGHIJ"],
            "rowgen": {"sheet": "Sens", "rows": f"4-{s_chk - 1}", "cols": list("ABCDEFGHIJ"), "fmts": ["", "", "", "n2", "n2", "sp1", "n2", "n2", "n2", "sp1"]}}
S["sens"]["blocks"] = [tor[0][0], tor[1][0],
                       {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入，三情境同時改、整本逐案重算；變動達兩位數百分比者直接顯示，其餘摺疊）。業務案例取 Sens 頁逐案推導區塊，目標價參數取 Sens 頁 DCF 參數案例。下方摺疊為全部案例與 EPS 敏感度對照。"]},
                       {"type": "note", "parts": ["@Sens!A2"], "collapse": True, "summary": "展開：敏感度頁的說明"},
                       tor[0][1], tor[1][1], all_tbl, eps_tor, full_tbl, {"type": "note", "parts": [f"@Sens!A{s_note}"]}]

# ================= 4. 反向 DCF =================
rv = lambda p, start=1: next(r for r in range(start, RVW.max_row + 1) if str(RVW.cell(r, 1).value or "").startswith(p))
ra, rb_, rbb, rc, rcmp, rsum, yh = rv("A. DCF"), rv("(a)"), rv("(b)"), rv("(c)"), rv("B. Comparable"), rv("C. 摘要"), rv("項目（情境 B）")
m1, m0 = rv("現價隱含的營業利益率", rb_), rv("模型：基準的等效營業利益率", rb_)
s1, rv1, rv0, g1, g0_ = rv("現價隱含的放大倍數", rbb), rv("現價隱含的 FY2028 營收（", rbb), rv("模型：FY2028 營收（", rbb), rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
w1, w0 = rv("現價隱含的 WACC", rc), rv("模型 WACC", rc)
kvr = [rv(p, ra) for p in ("現價（", "支持現價所需的評價日", "淨現金滾到評價日", "評價用股數", "支持現價所需的企業價值", "模型企業價值", "EV 需放大", "模型：情境 B 評價日")]
kvf = ["n2", "n2", "n1", "n2", "n1", "n1", "n2", "n2"]
cm = [rv(p, rcmp) for p in ("12 個月後前瞻 EPS", "只看 Comparable：支持現價所需的前瞻", "只看 Comparable：所需的相對", "加權（DCF 不變）", "同業前瞻本益比中位數")]
yrr = [rv(p, yh) for p in ("營收（", "營業利益 EBIT", "營業利益率", "稅率", "NOPAT＋折舊", "營運資金增加", "現值（")]
yfm = ["n1", "n1", "p1", "p1", "n1", "n1", "n1"]
S["revdcf"].update(intro="@RevDCF!A2", blocks=[
    {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]}, {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
    {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
     "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
              ["現金流放大倍數（FY2028 起各年同乘；模型基準為一倍）", f"@RevDCF!B{s1}|n2", "—"],
              ["FY2028 營收（十億美元）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
              ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0_}|p1"],
              ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
    {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(kvr, kvf)]},
    {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
     "rows": [[f"@RevDCF!A{cm[0]}", "—", f"@RevDCF!C{cm[0]}|n2", "—"]]
             + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!C{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!D{r}|sp1"] for k, r in enumerate(cm[1:4])]
             + [[f"@RevDCF!A{cm[4]}", "—", f"@RevDCF!C{cm[4]}|n1", "—"]]},
    {"type": "table", "caption": f"@RevDCF!A{yh}", "collapse": True, "summary": "展開：情境 B 各年營收、營業利益與現值（反解的計算輸入；十億美元）", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"],
     "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{f}" for c in "BCDEFG"] for r, f in zip(yrr, yfm)]}])

# ================= 5. 損益（標準損益表） =================
p0 = row_of("PL", "標準損益表（v0.1.0")
ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（含重組）", "營業利益", "營業利益率", "其他收入（淨額；營業外）", "稅前利益", "所得稅",
        "非控制權益損益", "稅後淨利（歸屬 Intel）", "淨利率", "股數（十億股；稀釋加權平均）", "EPS（美元；稀釋，歷史為公司公告值）"]
fmt = lambda u: {"%": "p1", "US$": "n2", "十億股": "n2"}.get(str(u).strip(), "n1")
PC = "CDEFGHIJKL"
rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in PC]
r_ratio = plr("比率"); r_chk = next(r for r in range(r_ratio, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("核對（應為 0"))
r_ycav = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("年度與對帳中"))
head = [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC]
is_tbl = {"type": "table", "caption": "標準損益表（十億美元；EPS 為美元）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B（基準）；美國 GAAP 合併（GAAP 口徑；Intel 另公布 Non-GAAP，市場共識多為 Non-GAAP，本表不做調整）", "head": head,
          "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]}
chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": head, "collapse": True, "summary": "展開：核對列（歷史捨入差、預測期對 PL 情境 B 與 DCF 頁、FY2030E–FY2032E 的推法差額）",
           "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{'p1' if '占營收' in str(PLW.cell(r, 1).value) else 'n3'}" for c in PC] for r in range(r_chk + 1, r_ycav) if PLW.cell(r, 1).value]}
rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
is_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
            "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:L{rv_}", "color": "#6fc7cb"}],
            "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:L{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                      {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:L{opm_}", "axis": "r", "color": "#c29a2b"},
                      {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:L{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
            "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸；美國 GAAP 合併，淨利為歸屬 Intel）。FY2026E 為前三季實績加第四季指引的錨點，含 Escrowed Shares 評價損失（非現金）；FY2030E–FY2032E 為收斂法延伸期（營業利益由淨利還原）。"}
S["pl"]["blocks"] = [{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, {"type": "note", "parts": [f"@PL!A{r_ycav}"]}, chk_tbl,
                     {"type": "pl", "collapse": True, "summary": "展開：三情境損益摘要與折線", "years": "@PL!C3:J3", "metrics": [
                         {"name": ["營收（", "@PL!B5", "）"], "fmt": "n1", "rows": {"A": "@PL!C5:J5", "B": "@PL!C25:J25", "C": "@PL!C45:J45"}, "chart": True},
                         {"name": "淨利率", "fmt": "p2", "rows": {"A": "@PL!C18:J18", "B": "@PL!C38:J38", "C": "@PL!C58:J58"}, "chart": True},
                         {"name": ["EPS（", "@PL!B19", "）"], "fmt": "n2", "rows": {"A": "@PL!C19:J19", "B": "@PL!C39:J39", "C": "@PL!C59:J59"}, "chart": True}]},
                     {"type": "pl", "collapse": True, "summary": "展開：三情境分部（FY2026 錨點、FY2027–FY2029）", "years": "@PL!D3:G3", "metrics": [
                         {"name": ["DCAI 營收（", "@PL!B8", "）"], "fmt": "n1", "rows": {"A": "@PL!D8:G8", "B": "@PL!D28:G28", "C": "@PL!D48:G48"}, "chart": True},
                         {"name": ["外部代工營收（", "@PL!B11", "）"], "fmt": "n1", "rows": {"A": "@PL!D11:G11", "B": "@PL!D31:G31", "C": "@PL!D51:G51"}, "chart": True},
                         {"name": ["Intel Foundry 營業利益（", "@PL!B13", "）"], "fmt": "n1", "rows": {"A": "@PL!D13:G13", "B": "@PL!D33:G33", "C": "@PL!D53:G53"}, "chart": True},
                         {"name": ["總資本支出（", "@PL!B21", "）"], "fmt": "n1", "rows": {"A": "@PL!D21:G21", "B": "@PL!D41:G41", "C": "@PL!D61:G61"}, "chart": True}]}]

# ================= 6. 驅動因子 =================
sf = {"%": "p1", "US$B": "n2"}
r_chk2 = ex("核對：CCPG＋DCAI＋All Other＋外部代工 − 營收（應為 0）")
seg_rows = [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{'n3' if r == r_chk2 else sf.get(str(PLW.cell(r, 2).value), 'n2')}" for c in "FGHI"] for r in range(r_ccg, r_sh + 1)]
seg_tbl = {"type": "table", "caption": f"@PL!A{sg}", "head": [f"@PL!A{hS}", f"@PL!B{hS}"] + [f"@PL!{c}{hS}" for c in "FGHI"], "rows": seg_rows}
chart2 = {"type": "theme", "title": "外部代工營收與總資本支出（情境 B，十億美元）及 Foundry 營業虧損（右軸）：FY2026E–FY2029E", "x": f"@PL!F{hS}:I{hS}", "fmt": "n1", "fmt_r": "n1", "stack": False, "pad": 1.3,
          "bars": [{"name": [f"@PL!A{r_ext}"], "y": f"@PL!F{r_ext}:I{r_ext}", "color": "#c29a2b"}, {"name": [f"@PL!A{r_cx}"], "y": f"@PL!F{r_cx}:I{r_cx}", "color": "#6fc7cb"}],
          "lines": [{"name": [f"@PL!A{r_fndoi}"], "y": f"@PL!F{r_fndoi}:I{r_fndoi}", "axis": "r", "color": "#b4472f", "labels": True}],
          "caption": "長條＝外部代工營收與總資本支出（公司口徑，未扣政府補助與夥伴挹注）；折線＝Intel Foundry 營業利益（右軸，負值＝虧損）。外部代工營收、Foundry 利益率與資本支出皆為公司層級假設，Siliconomics 尚無代工利益率與外部代工營收層。"}


def scn_table():
    ws = WB["In"]
    t = next(r for r in range(4, ws.max_row + 1) if str(ws.cell(r, 1).value or "").startswith("情境驅動對照"))
    hh = t + 1; rows = []; r = hh + 1
    while ws.cell(r, 1).value not in (None, ""):
        rows.append(r); r += 1
    assert ws.cell(rows[-1], 1).value == "scn_w", ws.cell(rows[-1], 1).value

    def fm(rr):
        u = str(ws.cell(rr, 3).value or "")
        vals = [ws.cell(rr, c).value for c in (4, 5, 6) if num(ws.cell(rr, c).value)]
        if u in ("%", "pt") or "%" in u: return "p1"
        return "n0" if vals and max(abs(v) for v in vals) >= 100 else "n2"
    return {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；取自 In 頁最後一節「情境驅動對照」）", "head": [f"@In!B{hh}", f"@In!D{hh}", f"@In!E{hh}", f"@In!F{hh}", f"@In!C{hh}"],
            "rows": [[f"@In!B{rr}", f"@In!D{rr}|{fm(rr)}", f"@In!E{rr}|{fm(rr)}", f"@In!F{rr}|{fm(rr)}", f"@In!C{rr}"] for rr in rows]}


SPEC = {
 "in_skip": [in_tok("F."), in_tok("情境驅動對照")],
 "in_groups": [
  ["Intel Products：CCPG（客戶端）與 DCAI（資料中心與 AI）", [
    ["CCPG 營收年增與營業利益率（三情境）", [k for k in I if re.match(r"^cc[gm]2\d_[ABC]$", k)] + ["ccg_adj", "ccm_adj"]],
    ["DCAI 中 AI 相關營收", ["ai_sh26", "aig27_A", "aig27_B", "aig27_C", "aig_L", "ai_adj"]],
    ["AI 加速器顆數係數（Siliconomics，FY2028 起乘在 DCAI 中 AI 相關營收）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["DCAI 其他營收年增與營業利益率（三情境）", [k for k in I if re.match(r"^d(ng|cm)2\d_[ABC]$", k)] + ["dcm_adj"]],
    ["All Other（Mobileye 等）", ["oth_g", "oth_m"]]]],
  ["Intel Foundry（估值主軸）", [
    ["內部銷售與外部代工營收", ["iota"] + [k for k in I if re.match(r"^ext2\d_[ABC]$", k)] + ["ext_adj"]],
    ["Foundry 營業利益率（占 Foundry 營收）", [k for k in I if re.match(r"^fm2\d_[ABC]$", k)] + ["fm_adj"]],
    ["Corporate Unallocated", ["cu_f"]]]],
  ["全公司（費用、資本支出、FY2026 錨點、股價與目標價規則、折現率）", [
    ["營業費用、營業外、稅與非控制權益、股數", ["opx27", "opx_g", "opx_adj", "nonop_f", "tax", "nci_bs", "nci_r", "nci_g", "sh_g"]],
    ["折舊與總資本支出", ["da_g"] + [k for k in I if re.match(r"^cx2\d_[ABC]$", k)] + ["cx_adj"]],
    ["FY2026 錨點（第四季指引或假設）", [k for k in I if k.startswith("q4_")] + ["q_sh"]],
    ["股價與市場倍數", ["px", "pe_ttm", "fpe", "cons_feps", "cons_tp", "nim_5"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC"]],
    ["收斂法（FY2030–FY2032）", ["tp_g", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta_u", "beta_du", "beta_de", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["行業對照與 Wattonomics 錨定（只作對照或進敏感度）", ["eqfd26", "eqfd27", "eqfd28", "wta_27", "wta_28", "wta_29"]]]]]}
RL._check_in_coverage(W, SPEC)
blocks, nums_ = [scn_table()], "一二三四五六七八九十"
for i, (tt, gs) in enumerate(SPEC["in_groups"]):
    blocks.append({"type": "h3", "text": f"{nums_[i]}、{tt}"})
    blocks += RL._in_tables(W, gs, tt.split("（")[0])
un_tbl = {"type": "table", "caption": "@Units!A4", "head": [f"@Units!{c}3" for c in "ABCDEF"],
          "rowgen": {"sheet": "Units", "rows": "5-22", "cols": list("ABCDEF"), "fmts": ["", "", "n1", "n1", "n1", "n1"], "row_fmt": {str(r): {c: "p1" for c in "CDEF"} for r in (17, 18, 19)}}}
un_tbl2 = {"type": "table", "caption": "@Units!A24", "head": [f"@Units!{c}3" for c in "ABCDEF"],
           "rowgen": {"sheet": "Units", "rows": "25-33", "cols": list("ABCDEF"), "fmts": ["", "", "n1", "n1", "n1", "n1"], "row_fmt": {str(r): {c: "p1" for c in "CDEF"} for r in (28, 29, 30)}}}
S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, chart2, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + blocks + [un_tbl, un_tbl2]

# ================= 7. 對帳：三層協調 C0（Coord 頁）與共識對照 =================
c1, c2, c3 = cr("1. 讀哪一層"), cr("2. C0 讀取層級核對"), cr("3. 對帳中年份")
h1 = c1 + 1; h2 = next(r for r in range(c2, CW.max_row + 1) if CW.cell(r, 1).value == "項目")
cnt = cr("對帳中項目數（讀取層級核對")
ucf = lambda r: {"%": "p1", "x": "n3", "%/年": "p1"}.get(str(CW.cell(r, 2).value), "n1")
c0_tbl = {"type": "table", "caption": "讀取層級核對：個股讀入的層級值 vs 層級值（同源，不判定、不計入對帳中；有檢驗力的對帳見上方公司實績表）", "collapse": True, "summary": "展開：讀取層級核對（同源，不判定）",
          "head": [f"@Coord!{c}{h2}" for c in "ABCDEFG"],
          "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|{ucf(r)}", f"@Coord!E{r}|{ucf(r)}", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in range(h2 + 1, cnt) if CW.cell(r, 4).value is not None]}
c0_src = {"type": "table", "caption": "層級名稱、版本、口徑與證據等級（每列的來源；展開看）", "collapse": True, "summary": "展開：每列的層級名稱、版本、口徑與證據等級",
          "head": ["項目", "年度", "層級名稱與版本／口徑／證據等級／備註"], "rows": [[f"@Coord!A{r}", f"@Coord!C{r}", f"@Coord!H{r}"] for r in range(h2 + 1, cnt) if CW.cell(r, 8).value]}
rd_tbl = {"type": "table", "caption": f"@Coord!A{c1}", "head": [f"@Coord!{c}{h1}" for c in "ABCDE"], "collapse": True, "summary": "展開：本模型讀哪一層、哪一版（下游只註明版本、不自行重估）",
          "rows": [[f"@Coord!A{r}", f"@Coord!B{r}" + ("|n3" if num(CW.cell(r, 2).value) else ""), f"@Coord!C{r}", f"@Coord!D{r}", f"@Coord!E{r}"] for r in range(h1 + 1, c2 - 1) if CW.cell(r, 1).value]}
rf0 = cr("層級之間"); rfh = next(r for r in range(rf0, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); rfe = next(r for r in range(rfh, CW.max_row + 1) if not CW.cell(r + 1, 1).value)
ref_blocks = [{"type": "h3", "text": f"@Coord!A{rf0}"}, {"type": "note", "parts": [f"@Coord!A{rf0 + 1}"]},
              {"type": "table", "head": [f"@Coord!{c}{rfh}" for c in "ABCDEFG"], "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n3", f"@Coord!E{r}|n3", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in range(rfh + 1, rfe + 1)]}]
a4 = cr("4. C0（實績）"); ih = next(r for r in range(a4, CW.max_row + 1) if CW.cell(r, 1).value == "項目")
jt = cr("C0（實績）判定", ih); jr0 = cr("C0（實績）判定：合併營收"); a4e = cr("對帳中項目數（C0 實績表", jr0)
ifm = lambda r: {"%": "p1", "%/年": "p1", "US$B": "n3"}.get(str(CW.cell(r, 2).value), "n3")
in_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|{ifm(r)}", f"@Coord!E{r}", f"@Coord!H{r}"] for r in range(ih + 1, jt) if CW.cell(r, 1).value]
jr_rows = [[f"@Coord!A{r}", f"@Coord!G{r}", f"@Coord!H{r}"] for r in range(jr0, a4e)]
act_blocks = [{"type": "h3", "text": "公司實績對層級隱含值（實績對帳）"}, {"type": "note", "parts": [f"@Coord!A{a4 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 2}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 3}"]},
              {"type": "table", "caption": "公司實績（季營收、分部、資本支出）與參考值：期間、等級與出處", "head": [f"@Coord!{c}{ih}" for c in "ABCDE"] + [f"@Coord!H{ih}"], "rows": in_rows},
              {"type": "table", "caption": "公司實績對層級的可比性判定", "head": ["項目", "判定", "結論與缺漏揭露"], "rows": jr_rows},
              {"type": "note", "parts": [f"@Coord!A{a4e + 1}"]}, {"type": "note", "parts": ["對帳中項目數：", f"@Coord!G{a4e}|n0", "。"]}]
c0_note = [{"type": "h3", "text": "讀取層級核對（同源，不判定）"}, {"type": "note", "parts": [f"@Coord!A{c2}"]}, {"type": "note", "parts": [f"@Coord!A{c2 + 1}"]}, c0_tbl]
c3_notes = [{"type": "note", "parts": [f"@Coord!A{c3 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{c3 + 2}"]}]
rcn = {"type": "table", "head": [f"@Recon!{c}3" for c in "ABCDEFGHI"],
       "rowgen": {"sheet": "Recon", "rows": "5,6,8,9,10,11", "cols": list("ABCDEFGHI"), "fmts": ["", "n2", "n2", "sp1", "n2", "sp1", "n2", "sp1", ""]}}
qrows = [r for r in range(4, 39) if WB["Q"].cell(r, 1).value]
recon_main = [{"type": "h3", "text": "本模型與公司說法、共識、券商的對照"}, dict(rcn, caption="本模型與共識、行業的對照（共識為 Non-GAAP 口徑，與本模型 GAAP 不同；差距不代表錯誤）"),
              {"type": "table", "caption": "@Q!A1", "collapse": True, "summary": "展開：FY2025 與 FY2026 各季損益與分部（Q 頁）", "head": [f"@Q!{c}3" for c in "ABCDEFG"],
               "rowgen": {"sheet": "Q", "rows": f"4-{max(qrows)}", "cols": list("ABCDEFG"), "fmts": ["", "n2", "n2", "n2", "n2", "n2", "n2"],
                          "row_fmt": {str(r): {c: "p1" for c in "BCDEFG"} for r in range(4, 39) if str(WB["Q"].cell(r, 1).value or "").strip().replace("  ", "").startswith(("毛利率", "有效稅率", "內部銷售占"))}}}]
S["recon"]["blocks"] = [{"type": "h3", "text": "三層協調：由下往上對帳（公司實績對各層）"}] + act_blocks + c0_note + ref_blocks + c3_notes + [c0_src, rd_tbl] + recon_main

# ================= 8. 資料狀態 =================
dm_rows = [["@In!B" + str(I[k]), ["@In!E" + str(I[k]) + "|" + f, " ", "@In!C" + str(I[k])], "@In!G" + str(I[k]) + "|tag", [lbl, "@PL!" + cell + "|" + pf, " ", unit], note]
           for k, f, lbl, cell, pf, unit, note in (
               ("fm29_B", "p1", "Foundry 營業利益率（FY2029，情境 B）：", f"I{r_fm}", "p1", "", "Foundry 虧損收斂路徑決定合併營業利益與估值"),
               ("ext29_B", "n1", "外部代工營收（FY2029，情境 B）：", f"I{r_ext}", "n1", "十億美元", "外部客戶放量；扣除分部間沖銷後的對外營收"),
               ("cx29_B", "n1", "總資本支出（FY2029，情境 B）：", f"I{r_cx}", "n1", "十億美元", "資本支出壓低自由現金流；DCF 逐年取用"),
               ("ccg29_B", "p1", "CCPG 營收（FY2029，情境 B）：", f"I{r_ccg}", "n1", "十億美元", "客戶端接近成熟，年增低個位數"),
               ("dcm29_B", "p1", "DCAI 營業利益率（FY2029，情境 B）：", f"I{r_dcm}", "p1", "", "資料中心利潤支付 Foundry 虧損"),
               ("aig27_B", "p1", "DCAI 中 AI 相關營收（FY2029，情境 B）：", f"I{r_dai}", "n1", "十億美元", "FY2028 起乘 Siliconomics 顆數係數"))]
S["status"]["blocks"] = [
    {"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]},
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋（取代建置 commit，理由見 README）", "$sha"],
                             ["Tokenomics 快照版本", "本模型不直接讀取 Tokenomics 快照；行業變數（顆數情境係數、Wattonomics 錨定、代工與邏輯設備成長）讀 Siliconomics 快照，版本見 In 頁 A 節標題。"]]},
    {"type": "h3", "text": "上游成品（市場層）"}, {"type": "note", "parts": ["@In!A4"]},
    {"type": "h3", "text": "待複核項目"},
    {"type": "list", "items": ["公司財報數字取自公司新聞稿（IR 頁）；Yahoo 彙整數只作交叉檢查與同業 EPS，差異見 README。", "@Q!A2", "標為「未找到」與「待併入 Siliconomics」的項目（見下方清單）待補。"]},
    {"type": "h3", "text": "已知限制"}, {"type": "note", "parts": [f"@README!A{row_of('README', '已知限制')}"]},
    {"type": "h3", "text": "附錄：驅動 → 推導對照表（查核用）"},
    {"type": "table", "collapse": True, "summary": "展開：每頁的輸入、推導量與因果方向", "head": ["輸入", "基準值", "標記", "推導量", "因果方向"], "rows": dm_rows},
    {"type": "h3", "text": "附錄：版本要點"},
    {"type": "table", "collapse": True, "summary": "展開：本版與前版的修改要點", "head": ["要點"],
     "rows": [[f"@README!A{r}"] for r in range(1, RD.max_row + 1) if re.match(r"^v\d+\.\d", str(RD.cell(r, 1).value or ""))]},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的輸入與來源", "sheets": {"In": ["B", "H", "I"], "Q": ["A", "J"], "Sources": ["A"]}, "empty": "沒有標為未找到的項目。"},
    {"type": "h3", "text": "「待併入 Siliconomics」清單"},
    {"type": "scan", "needle": "待併入 Siliconomics", "collapse": True, "summary": "展開：公司層級、待併入行業層的變數", "sheets": {"In": ["A", "B", "I"]}, "empty": "沒有待併入的變數。"}]

json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

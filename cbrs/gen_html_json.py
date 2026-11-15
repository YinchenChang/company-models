# Cerebras html.json（v0.1.2 起）：把 build_cbrs.py --html 寫出的 html_src.json 轉成個股 HTML 版面規格 v1.3。
# 用法：python3 cbrs/gen_html_json.py cbrs/dist/<日期>_Cerebras個股模型_v<版本>.xlsx
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含 Coord：讀取層級、兩讀數判定、C0 實績）→ 資料狀態。
# 數字與文字取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末、Coord 頁；cbrs/v012_post.py 產生）。
# 列號依標籤找；設定文字除年份與日期外不含數字。樣板：amd/gen_html_json.py（v0.4.7 段）、nvidia/gen_html_json.py。
import json, os, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
X = sys.argv[1]
wb = openpyxl.load_workbook(X, data_only=True)
cfg = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))


def find(sheet, p, start=1, exact=False):
    ws = wb[sheet]
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and ((v.strip() == p) if exact else v.strip().startswith(p)):
            return r
    raise KeyError((sheet, p))


S = {x["id"]: x for x in cfg["sections"]}
RD, PLW, RV, CW, SW, VW = wb["README"], wb["PL"], wb["RevDCF"], wb["Coord"], wb["Sens"], wb["Val"]
r13 = find("README", "v1.3 展示層")
RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "coord", "ver"], range(r13 + 1, r13 + 11)))
assert str(RD[f"A{RM['ver']}"].value).startswith("v0.1.3"), RD[f"A{RM['ver']}"].value
cfg["layout"] = "v1.3"
old_note = list(cfg["hero"]["note"])
assert old_note[0] == "投資建議依加權目標價判定：", old_note
cfg["hero"]["note"] = [f"@README!A{RM['hero']}"]

# ---- 1. 命題：置頂對照圖、投資建議前提、重點觀察 5 點 ----
th = S["thesis"]["blocks"]
rx27, rrpo, rmw = find("Recon", "FY2027 營收 ÷ FY2026"), find("Recon", "雲端營收 4Q26–3Q28"), find("Recon", "FY2028 平均計費 MW vs 公司")
grp = lambda nm, r, fmt, ref: {"name": [nm], "fmt": fmt, "items": [{"name": [ref], "v": f"@Recon!E{r}", "ref": True},
                                                                  {"name": ["模型 情境 A"], "v": f"@Recon!B{r}"}, {"name": ["模型 情境 B"], "v": f"@Recon!C{r}"}, {"name": ["模型 情境 C"], "v": f"@Recon!D{r}"}]}
cmp_ = {"type": "compare", "title": "回答命題：模型的 FY2027 營收倍數、RPO 期間雲端營收與 FY2028 計費 MW，對公司說法與合約揭露",
        "groups": [grp("FY2027 營收 ÷ FY2026（倍）", rx27, "x1", "公司說法（三倍）"), grp("RPO 期間的雲端營收（十億美元）", rrpo, "n2", "RPO 二十四個月內認列額"),
                   grp("FY2028 平均計費 MW", rmw, "n0", "公司：2027 年底前上線＋簽約")],
        "caption": "灰色＝公司說法或合約揭露；青綠＝本模型三情境。對照只作事後檢查，校準時未使用；公司 MW 為容量口徑（含尚未計費），與計費 MW 不同。"}
i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
th.insert(0, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
th.insert(0, cmp_)

# ---- 2. 估值：四個目標價置頂（G 位置） ----
va = S["valuation"]["blocks"]
i_bar = next(i for i, b in enumerate(va) if b["type"] == "bars")
g0 = find("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = g0 + 7
four = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
        "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                  ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整）", "加權目標價（橫幅主值）", "市場共識目標價"))],
        "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
tbl4 = {"type": "table", "caption": "四個目標價並列（美元；相對現價）", "head": ["目標價", "美元", "相對現價", "說明"],
        "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_px}", f"@Val!B{r_px}|n2", "", f"@Val!D{r_px}"]]}
va[i_bar] = four
va.insert(i_bar + 1, {"type": "note", "parts": old_note})
va.insert(i_bar + 1, tbl4)
assert i_bar == 0
va_intro = S["valuation"].pop("intro", None)
if va_intro:
    va.insert(3, {"type": "note", "parts": [va_intro]})

# ---- 3. 敏感度：加權目標價相對基準（H 位置） ----
t0 = find("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
while SW.cell(r2 + 1, 1).value:
    r2 += 1
tor = []
for col, side in (("B", "低"), ("C", "高")):
    big = [r for r in range(r1, r2 + 1) if isinstance(SW[f"{col}{r}"].value, (int, float)) and abs(SW[f"{col}{r}"].value) >= 0.1]
    small = [r for r in range(r1, r2 + 1) if isinstance(SW[f"{col}{r}"].value, (int, float)) and 0.0005 <= abs(SW[f"{col}{r}"].value) < 0.1]
    base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
    tor.append((dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))),
                dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                     summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例")))
se = S["sens"]["blocks"]
for b in se:
    if b["type"] == "tornado":
        b.update(collapse=True, summary="展開：對照：情境 B 加權價值相對基準的變動（一次動一個輸入）")
    elif b["type"] == "table" and not b.get("collapse"):
        b.update(collapse=True, summary="展開：目標價參數表（WACC、β、ERP、永續成長率等）")
all_tbl = {"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
           "head": [f"@Sens!{c}{t0 + 1}" for c in "ABCDEF"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1", f"@Sens!D{r}|n2", f"@Sens!E{r}|n2", f"@Sens!F{r}"] for r in range(r1, r2 + 1)]}
S["sens"]["blocks"] = [tor[0][0], tor[1][0], {"type": "note", "parts": [f"@Sens!A{t0}"]}, tor[0][1], tor[1][1], all_tbl] + se
se_intro = S["sens"].pop("intro", None)
S["sens"]["blocks"].insert(3, {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入；變動達兩位數百分比者直接顯示，其餘摺疊）；營運輸入案例只套用情境 B。下方摺疊為全部案例、原案例表與目標價參數明細。"]})
if se_intro:
    S["sens"]["blocks"].insert(4, {"type": "note", "parts": [se_intro]})

# ---- 4. 反向 DCF ----
rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").startswith(p))
ra = rv("A. DCF"); rb_ = rv("(a)"); rbb = rv("(b)"); rc = rv("(c)"); rcmp = rv("B. Comparable"); rsum = rv("C. 摘要"); yh = rv("項目（情境 B）")
m1, m0 = rv("現價隱含的 EBIT 率", rb_), rv("模型：基準的等效 EBIT 率", rb_)
s1, rv1, rv0 = rv("現價隱含的營收倍數", rbb), rv("現價隱含的 FY2028 營收", rbb), rv("模型：FY2028 營收", rbb)
gg1, gg0 = rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
w1, w0, wok = rv("現價隱含的 WACC", rc + 1), rv("模型 WACC", rc), rv("判讀（隱含 WACC）", rc)
a_rows = [rv(p, ra) for p in ("現價（美元）", "支持現價所需的評價日每股價值", "淨現金滾到評價日", "評價用稀釋股數", "支持現價所需的企業價值", "模型企業價值", "模型：情境 B 評價日每股價值")]
afmt = ["n2", "n2", "n2", "n3", "n2", "n2", "n2"]
rev_ = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2", "blocks": [
    {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
    {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
    {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
     "rows": [["預測期 EBIT 率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
              ["營收規模（FY2028 起各年同乘的倍數）", f"@RevDCF!B{s1}|n2", "—"],
              ["FY2028 營收（十億美元）", f"@RevDCF!B{rv1}|n2", f"@RevDCF!B{rv0}|n2"],
              ["FY2028 營收年增率", f"@RevDCF!B{gg1}|p1", f"@RevDCF!B{gg0}|p1"],
              ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"],
              ["隱含 WACC 判讀", f"@RevDCF!B{wok}", ""]]},
    {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(a_rows, afmt)]},
    {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
     "rows": [[f"@RevDCF!A{rcmp + 2}", "—", f"@RevDCF!C{rcmp + 2}|n2", "—"]]
             + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|n2", f"@RevDCF!C{r}|n2", f"@RevDCF!D{r}|sp1"] for r in range(rcmp + 3, rcmp + 6)]
             + [[f"@RevDCF!A{rcmp + 6}", "—", f"@RevDCF!C{rcmp + 6}|n2", "—"]]},
    {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
     "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）",
     "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|" + ("p1" if r in (yh + 3, yh + 4) else "n3" if r in (yh + 5, yh + 6, yh + 7, yh + 8) else "n2") for c in "BCDEFG"]
              for r in range(yh + 1, rb_ - 1)]}]}

# ---- 5. 損益（標準損益表） ----
p0 = find("PL", "標準損益表（v0.1.2")
ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
COLS_ = "CDEFGHIJK"
keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（研發＋銷管，含經常性股票薪酬）", "營業利益", "營業利益率",
        "其他收入（淨額）", "稅前利益", "所得稅", "稅後淨利", "淨利率", "稀釋股數（含客戶認股權證）", "EPS（完全稀釋；模型口徑）"]
fmt = lambda u: {"%": "p1", "US$": "n2", "十億股": "n3"}.get(u, "n3")
cellf = lambda r, c: (f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" if isinstance(PLW[f"{c}{r}"].value, (int, float)) else f"@PL!{c}{r}")
rowx = lambda r: [f"@PL!A{r}"] + [cellf(r, c) for c in COLS_]
r_ratio = plr("比率"); r_chk = plr("核對（應為 0）")
is_tbl = {"type": "table", "caption": "標準損益表（GAAP 近似；十億美元，EPS 為美元，完全稀釋）：FY2024A–FY2025A GAAP 實績、FY2026E–FY2032E 情境 B（基準）",
          "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in COLS_],
          "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]}
chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in COLS_], "collapse": True, "summary": "展開：標準損益表核對列",
           "rows": [[f"@PL!A{r}"] + [(f"@PL!{c}{r}|n3" if PLW[f"{c}{r}"].value is not None else "—") for c in COLS_] for r in range(r_chk + 1, r_chk + 5)]}
rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
is_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率：FY2024A–FY2032E（情境 B）", "x": f"@PL!C{ph}:K{ph}",
            "fmt": "n1", "fmt_r": "p0", "stack": True, "pad": 1.3,
            "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:K{rv_}", "color": "#6fc7cb"}],
            "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:K{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                      {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:K{opm_}", "axis": "r", "color": "#c29a2b"}],
            "caption": "長條＝營收；兩條線＝毛利率、營業利益率（右軸）。歷史為 GAAP；預測為模型口徑（含經常性股票薪酬，不含 IPO 一次性費用）。淨利率見下表（歷史年含一次性項目）。"}
old_pl = S["pl"]["blocks"]
keep = [dict(b, collapse=True, summary="展開：三情境損益摘要（營收、平均計費 MW、毛利率、EBIT、EPS、資本支出）") if b["type"] == "pl" else b for b in old_pl if b["type"] != "note"]
pl_intro = S["pl"].pop("intro", None)
S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_tbl, is_chart, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, chk_tbl]
               + [{"type": "note", "parts": [pl_intro], "collapse": True, "summary": "展開：三情境損益的口徑說明"}] + keep)

# ---- 6. 驅動因子：分部營收推導（情境 B） ----
rbB = find("PL", "情境 B（基準）"); hdrB = rbB + 1
assert PLW.cell(hdrB, 1).value == "項目"
pr = lambda lab_: next(r for r in range(hdrB, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
r_mw, r_pmw, r_cl, r_hw, r_rev = pr("平均計費 MW"), pr("每 MW 年營收"), pr("雲端核心營收"), pr("硬體核心營收"), pr("核心營收合計")
seg_chart = {"type": "theme", "title": "分部營收（情境 B，十億美元）與平均計費 MW（右軸）：FY2026E–FY2032E", "x": f"@PL!D{hdrB}:J{hdrB}",
             "fmt": "n1", "fmt_r": "n0", "stack": True, "pad": 1.2, "total": f"@PL!D{r_rev}:J{r_rev}",
             "bars": [{"name": ["推論雲端"], "y": f"@PL!D{r_cl}:J{r_cl}", "color": "#0b7a80"}, {"name": ["系統（硬體）"], "y": f"@PL!D{r_hw}:J{r_hw}", "color": "#c29a2b"}],
             "lines": [{"name": [f"@PL!A{r_mw}"], "y": f"@PL!D{r_mw}:J{r_mw}", "axis": "r", "color": "#3b3838", "labels": True}],
             "caption": "堆疊長條＝推論雲端（平均計費 MW × 每 MW 年營收）＋系統銷售＝核心營收合計；折線＝平均計費 MW。FY2026E 為錨點（前三季實績＋第四季指引）。"}
fm_ = {r_mw: "n0", r_pmw: "n2"}
seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B）：推論雲端＝平均計費 MW × 每 MW 年營收；系統（硬體）＝FY2026 錨點 × 情境年增；合計＝核心營收",
           "head": [f"@PL!A{hdrB}", f"@PL!B{hdrB}"] + [f"@PL!{c}{hdrB}" for c in "DEFGHIJ"],
           "rows": [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{fm_.get(r, 'n3')}" for c in "DEFGHIJ"] for r in (r_mw, r_pmw, r_cl, r_hw, r_rev)]}
S["drivers"]["blocks"] = [seg_chart, {"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_tbl] + S["drivers"]["blocks"]

# ---- 8. 對帳：Coord（讀取層級、兩讀數判定、層級之間、C0 實績） ----
cr = lambda p, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
cA, cB, cC, cD = cr("A. 讀取層級"), cr("B. 兩讀數"), cr("C. 層級之間"), cr("D. C0")
d_hdr = next(r for r in range(cD, CW.max_row + 1) if CW.cell(r, 1).value == "項目")
d_in = next(r for r in range(cD, CW.max_row + 1) if CW.cell(r, 1).value == "輸入")
d_end = d_hdr + 1
while CW.cell(d_end + 1, 1).value:
    d_end += 1
r_cnt, r_cnote = cr("對帳中的 C0 項目數"), cr("說明：D 節")
coord_blocks = [
    {"type": "h3", "text": "三層協調：公司實績對層級隱含值"},
    {"type": "note", "parts": [f"@README!A{RM['coord']}"]},
    {"type": "table", "caption": f"@Coord!A{cD}", "head": [f"@Coord!{c}{d_hdr}" for c in "ABCDEFGHI"],
     "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|n3"] + [f"@Coord!{c}{r}" for c in "CDEFGHI"] for r in range(d_hdr + 1, d_end + 1)]},
    {"type": "note", "parts": [f"@Coord!A{cD + 1}"]},
    {"type": "table", "caption": "實績輸入：公司季度實績（十億美元）與出處", "head": [f"@Coord!{c}{d_in}" for c in "ABCDEF"], "collapse": True, "summary": "展開：實績輸入（季度實績、出處與等級）",
     "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|" + ("n0" if CW.cell(r, 3).value == "月" else "n3")] + [f"@Coord!{c}{r}" for c in "CDEF"] for r in range(d_in + 1, d_hdr - 1) if CW.cell(r, 1).value]},
    {"type": "kv", "items": [[f"@Coord!A{r_cnt}", f"@Coord!B{r_cnt}|n0"]]},
    {"type": "note", "parts": [f"@Coord!A{r_cnote}"]},
    {"type": "table", "caption": f"@Coord!A{cA}", "head": [f"@Coord!{c}{cA + 1}" for c in "ABCDE"], "collapse": True, "summary": "展開：讀取層級（讀哪一層、哪一版、哪些名稱）",
     "rows": [[f"@Coord!{c}{r}" for c in "ABCDE"] for r in range(cA + 2, cA + 5)]},
    {"type": "table", "caption": f"@Coord!A{cB}", "head": [f"@Coord!{c}{cB + 1}" for c in "ABCD"], "collapse": True, "summary": "展開：兩讀數判定：不適用（另一側讀數無傳導）",
     "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|" + ("sp1" if "上檔" in str(CW.cell(r, 1).value) else "n2"), f"@Coord!C{r}|" + ("sp1" if "上檔" in str(CW.cell(r, 1).value) else "n2"), f"@Coord!D{r}|sp1"] for r in range(cB + 2, cB + 6)]
             + [[f"@Coord!A{cB + 6}", f"@Coord!B{cB + 6}", f"@Coord!C{cB + 6}", ""], [f"@Coord!A{cB + 7}", f"@Coord!B{cB + 7}", "", ""]]},
    {"type": "note", "parts": [f"@Coord!B{cB + 8}"], "collapse": True, "summary": "展開：為何兩讀數不適用"},
    {"type": "table", "caption": f"@Coord!A{cC}", "head": [f"@Coord!{c}{cC + 1}" for c in "ABCDE"], "collapse": True, "summary": "展開：層級之間（參考）",
     "rows": [[f"@Coord!A{cC + 2}", f"@Coord!B{cC + 2}|n0", f"@Coord!C{cC + 2}", f"@Coord!D{cC + 2}", f"@Coord!E{cC + 2}"],
              [f"@Coord!A{cC + 3}", f"@Coord!B{cC + 3}|sp1", f"@Coord!C{cC + 3}", f"@Coord!D{cC + 3}", f"@Coord!E{cC + 3}"]]},
]
S["recon"]["blocks"] = S["recon"]["blocks"] + coord_blocks

# ---- 章節順序 ----
order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
S["revdcf"] = rev_
cfg["sections"] = [S[k] for k in order]
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("html.json：版面 v1.3（章節", " → ".join(order), "）")

# ABB html.json 產生器（個股 HTML 版面規格 v1.3；樣板 delta/gen_html_json.py 的結構，直接由 xlsx 標籤定位列號，不經 relayout）。
# 讀 abb/dist/ 內唯一的 xlsx（LibreOffice 重算後）→ 寫 abb/html.json（勿手改 html.json）。文字不含阿拉伯數字（年度與日期除外）；數字一律取自 xlsx。
# 用法：python3 abb/gen_html_json.py [成品 xlsx]
import glob, json, os, re, sys
import openpyxl
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
X = sys.argv[1] if len(sys.argv) > 1 else (lambda fs: fs[0] if len(fs) == 1 else sys.exit(f"dist 內應剛好一個 xlsx：{fs}"))(sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx"))))
wb = openpyxl.load_workbook(X, data_only=True)


def row(sh, label, start=1, exact=False, end=None):
    ws = wb[sh]
    for r in range(start, (end or ws.max_row) + 1):
        v = str(ws.cell(r, 1).value or "").strip()
        if (v == label.strip()) if exact else v.startswith(label.strip()):
            return r
    raise KeyError((sh, label))


def cell(sh, coord): return wb[sh][coord].value


V = lambda lab, c="B", f="", **kw: f"@Val!{c}{row('Val', lab, **kw)}" + (f"|{f}" if f else "")
VA = lambda lab, **kw: f"@Val!A{row('Val', lab, **kw)}"
RD = lambda prefix: row("README", prefix)
RDa = lambda prefix: f"@README!A{RD(prefix)}"

# ---- PL 情境 B 區塊 ----
PL = wb["PL"]
rbB = row("PL", "情境 B")
hB = rbB + 1
assert PL.cell(hB, 1).value == "項目"
pB = lambda lab, ex=False: row("PL", lab, rbB, ex)
YC = {y: get_column_letter(3 + i) for i, y in enumerate(range(2023, 2033))}
blkA, blkC = row("PL", "情境 A"), row("PL", "情境 C")
hA, hC = blkA + 1, blkC + 1
pX = lambda base, lab: row("PL", lab, base, True)

# ================= 命題與橫幅 =================
hero_r = RD("ABB 的 AI 資料中心只占營收約")
cfg = {"company": "ABB", "ticker": "ABBN", "thesis": RDa("命題 1："),
       "question": "ABB 的 AI 暴露度有限，現在的股價還撐得住嗎",
       "scenarios": [{"key": "A", "name": "情境 A（順利）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（不如預期）"}],
       "wide_text": True, "layout": "v1.3",
       "hero": {"rec": V("投資建議"), "tiles": [
           [[VA("目標價（DCF 與 Comparable 加權）"), "（瑞士法郎）"], V("目標價（DCF 與 Comparable 加權）", f="n2"), ["上檔空間 ", V("上檔空間：加權目標價", f="sp1")]],
           [[VA("Comparable 目標價（市場主流同業法）"), "（瑞士法郎）"], V("Comparable 目標價（市場主流同業法）", f="n2"), ["上檔空間 ", V("上檔空間：Comparable 目標價", f="sp1")]],
           [[VA("現價（"), "（瑞士法郎）"], V("現價（", f="n2"), ""],
           [[VA("市場共識目標價"), "（瑞士法郎）"], V("市場共識目標價", f="n2"), ""]],
           "note": [f"@README!A{hero_r}"]}}
sections = []

# ---- 1. 命題 ----
r_nd, r_dc, r_rev = pB("電氣化：非資料中心營收"), pB("電氣化：資料中心營收"), pB("營收", True)
r_dcg, r_el, r_mo, r_au = pB("資料中心占集團營收"), pB("電氣化營收"), pB("馬達驅動營收"), pB("自動化營收")
r_co = pB("公司與其他及沖銷（淨）")
r_oth = pB("非電氣化營收")
obs = [f"@README!A{RD('重點 ' + str(i) + '：')}" for i in range(1, 6)]
ths = [f"@README!A{RD('命題 ' + str(i) + '：')}" for i in range(1, 4)]
ans = [f"@README!A{RD('回覆 ' + str(i) + '：')}" for i in range(1, 4)]
sections.append({"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "電氣化資料中心營收、其他電氣化與其他分部營收（情境 B，十億美元）與資料中心占集團營收（右軸）", "x": f"@PL!F{hB}:L{hB}",
     "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.25, "total": f"@PL!F{r_rev}:L{r_rev}",
     "bars": [{"name": ["資料中心（電氣化的一部分；本模型拆分）"], "y": f"@PL!F{r_dc}:L{r_dc}", "color": "#0b7a80"},
              {"name": ["電氣化其餘（電網、建築、工業）"], "y": f"@PL!F{r_nd}:L{r_nd}", "color": "#6fc7cb"},
              {"name": ["非電氣化（馬達驅動、自動化與公司沖銷）"], "y": f"@PL!F{r_oth}:L{r_oth}", "color": "#c29a2b"}],
     "lines": [{"name": ["資料中心占集團營收（AI 暴露度）"], "y": f"@PL!F{r_dcg}:L{r_dcg}", "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝三塊營收，頂端數字＝合併營收；折線＝資料中心占集團營收（右軸）。回答命題：資料中心占比約一成多，且 Wattonomics 設施支出自高點回落後占比下降。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": ths},
    {"type": "note", "parts": [f"@README!A{RD('投資建議條件式：')}"]},
    {"type": "list", "ordered": True, "collapse": True, "summary": "展開：命題回覆（依本模型計算的答案；想先看結論可展開，或讀完全文再看）", "items": ans},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": obs},
    {"type": "h3", "text": "情境定義"},
    {"type": "kv", "items": [["情境 A", f"@PL!A{blkA}"], ["情境 B", f"@PL!A{rbB}"], ["情境 C", f"@PL!A{blkC}"], ["會計年度與口徑", f"@README!A{RD('用途：')}"], ["分部命名", f"@README!A{RD('分部：')}"]]}]})

# ---- 2. 估值 ----
vtp, vdc, vcp, vpx, vup, vupc = (row("Val", p) for p in ("目標價（DCF 與 Comparable 加權）", "DCF 目標價（情境加權）", "Comparable 目標價（市場主流同業法）", "現價（", "上檔空間：加權目標價", "上檔空間：Comparable 目標價"))
vcons = row("Val", "市場共識目標價")
bars4 = {"type": "bars", "title": "目標價與現價（瑞士法郎；豎線＝現價）", "fmt": "n1",
         "items": [{"label": ["DCF 目標價"], "value": f"@Val!B{vdc}", "series": 1}, {"label": ["Comparable 目標價"], "value": f"@Val!B{vcp}", "series": 2},
                   {"label": ["加權目標價（橫幅主值）"], "value": f"@Val!B{vtp}", "series": 0}, {"label": ["市場共識目標價"], "value": f"@Val!B{vcons}", "series": 1}],
         "ref": {"label": f"@Val!A{vpx}", "value": f"@Val!B{vpx}"}}
conc = {"type": "table", "caption": "目標價、現價與投資建議（瑞士法郎）", "head": ["項目", "數值", "說明"],
        "rows": [[f"@Val!A{r}", f"@Val!B{r}|{f}", f"@Val!E{r}"] for r, f in ((vtp, "n2"), (vdc, "n2"), (vcp, "n2"), (vpx, "n2"), (vup, "sp1"), (vupc, "sp1"), (row("Val", "上檔空間級距"), ""), (row("Val", "投資建議"), ""), (vcons, "n2"), (row("Val", "價格帶"), ""))]}
INr = {str(wb["In"].cell(r, 1).value): r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}
rs_ = row("Val", "DCF 情境價值推到 12 個月後（瑞士法郎）")
sc_t = {"type": "table", "caption": "三情境 DCF 價值（推到十二個月後，瑞士法郎）與權重", "head": ["項目", "情境 A", "情境 B", "情境 C"],
        "rows": [[f"@Val!A{rs_}"] + [f"@Val!{c}{rs_}|n1" for c in "BCD"], ["情境權重", f"@In!E{INr['tp_wA']}|p0", f"@In!E{INr['tp_wB']}|p0", f"@In!E{INr['tp_wC']}|p0"]]}
vr = lambda lab: row("Val", lab)
rate_t = {"type": "table", "caption": "DCF 折現率：WACC、股權成本、稅後債務成本、RONIC、永續成長率", "head": ["項目", "數值", "說明"],
          "rows": [[f"@Val!A{vr(l)}", f"@Val!B{vr(l)}|{f}", f"@Val!E{vr(l)}"] for l, f in (("股權成本 ke", "p2"), ("稅後債務成本", "p2"), ("權益占資本", "p1"), ("WACC", "p2"), ("永續成長率 g", "p2"), ("永續期 RONIC", "p2"))]}
hs = vr("股權成本 ke") - 1
summ = [("明確期 FCFF 現值", "n2"), ("終值現值", "n2"), ("企業價值 EV", "n2"), ("＋淨現金與投資", "n2"), ("＋Stub 期 FCFF", "n2"), ("淨現金滾到評價日", "n2"), ("評價用稀釋股數", "n3"),
        ("DCF 情境每股價值（評價日）", "n2"), ("DCF 情境價值推到 12 個月後（美元", "n2"), ("其中：終值現值占 EV", "p1"), ("隱含終值 EV／EBITDA", "x1")]
sum_t = {"type": "table", "caption": "DCF 彙總（美元）：明確期現值、終值現值、企業價值、淨現金與每股價值", "head": [f"@Val!{c}{hs}" for c in "ABCD"],
         "rows": [[f"@Val!A{vr(l)}"] + [f"@Val!{c}{vr(l)}|{f}" for c in "BCD"] for l, f in summ] + [[f"@Val!A{vr('DCF 目標價（情境加權；美元）')}", f"@Val!B{vr('DCF 目標價（情境加權；美元）')}|n2", "", ""]]}
D = wb["DCF"]
hd = next(r for r in range(1, 14) if D.cell(r, 1).value == "項目")
dcols = [get_column_letter(c) for c in range(3, D.max_column + 1) if D.cell(hd, c).value]
dcf_blocks = []
for r in range(1, D.max_row + 1):
    if D.cell(r, 1).value == "營收" and str(D.cell(r - 1, 1).value or "").startswith("情境"):
        pv = next(k for k in range(r, r + 20) if str(D.cell(k, 1).value or "").startswith("現值（FCFF"))
        rf = {str(k): {c: ("p1" if str(D.cell(k, 1).value).startswith(("稅率", "折舊攤銷 ÷", "資本支出 ÷")) else "n3") for c in dcols}
              for k in range(r, pv + 1) if str(D.cell(k, 1).value).startswith(("稅率", "折舊攤銷 ÷", "資本支出 ÷", "折現因子"))}
        dcf_blocks.append({"type": "table", "caption": f"@DCF!A{r - 1}", "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細",
                           "head": [f"@DCF!A{hd}"] + [f"@DCF!{c}{hd}" for c in dcols],
                           "rowgen": {"sheet": "DCF", "rows": f"{r}-{pv}", "cols": ["A"] + dcols, "fmts": [""] + ["n2"] * len(dcols), "row_fmt": rf}})
q0 = row("Val", "0.2 Comparable")
comp_t = {"type": "table", "caption": f"@Val!A{q0}", "head": [f"@Val!{c}{q0 + 1}" for c in "ABC"],
          "rows": [[f"@Val!A{r}", f"@Val!B{r}", f"@Val!C{r}|n1"] for r in range(q0 + 2, q0 + 6)] +
                  [[f"@Val!A{vr(l)}", "", f"@Val!B{vr(l)}|{f}"] for l, f in (("同業前瞻 P/E 中位數", "n1"), ("12 個月後的前瞻 EPS", "n3"), ("Comparable 目標價（美元）", "n2"), ("加權目標價（美元）", "n2"))]}
pe_t = {"type": "table", "caption": f"@Val!A{row('Val', '3. 現價隱含倍數')}", "head": ["項目", "情境 A", "情境 B", "情境 C"], "collapse": True, "summary": "展開：現價隱含倍數（對模型 EPS 與共識 EPS）",
        "rows": [[f"@Val!A{r}"] + [f"@Val!{c}{r}|n1" for c in "BCD"] for r in range(row("Val", "現價 ÷ FY2026 EPS"), row("Val", "Yahoo 前瞻 P/E（幣別錯置）") + 1)]}
sections.append({"id": "valuation", "title": "估值", "blocks": [bars4, {"type": "note", "parts": [RDa("估值：")]}, conc, sc_t, rate_t, sum_t] + dcf_blocks + [
    {"type": "note", "parts": ["@DCF!A2"], "collapse": True, "summary": "展開：DCF 的營業利益與自由現金流口徑"}, comp_t, pe_t]})

# ---- 3. 敏感度 ----
SE = wb["Sens"]
hdr_r = [r for r in range(1, 41) if SE.cell(r, 3).value == "低：加權目標價"]
case_rows = []
for r in range(hdr_r[0] + 1, SE.max_row + 1):
    a = str(SE.cell(r, 1).value or "")
    if a.startswith(("檢查", "註", "經營輸入")) or SE.cell(r, 1).value is None or a == "參數": continue
    if any(isinstance(SE.cell(r, c).value, (int, float)) for c in (3, 5)): case_rows.append(r)
big = 0.10


def side(col, is_big):
    out = []
    for r in case_rows:
        v = SE[f"{col}{r}"].value
        if isinstance(v, (int, float)) and ((abs(v) >= big) == is_big): out.append(r)
    return ",".join(map(str, sorted(out, key=lambda r: SE[f"{col}{r}"].value)))


sb = []
for col, nm in (("D", "低"), ("F", "高")):
    rr = side(col, True)
    sb.append({"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1", "title": f"加權目標價相對基準：輸入取{nm}值，變動達兩位數百分比的案例", "rows": rr})
sb.append({"type": "note", "parts": ["@Sens!A2"]})
for col, nm in (("D", "低"), ("F", "高")):
    rr = side(col, False)
    sb.append({"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1", "title": f"加權目標價相對基準：輸入取{nm}值", "rows": rr, "collapse": True,
               "summary": f"展開：加權目標價相對基準：輸入取{nm}值，變動未達兩位數百分比的案例"})
sb.append({"type": "table", "collapse": True, "summary": "展開：全部案例的加權目標價（美元）與相對基準（低案、高案）", "head": [f"@Sens!{c}{hdr_r[0]}" for c in "ABCDEF"],
           "rows": [[f"@Sens!A{r}", f"@Sens!B{r}", f"@Sens!C{r}|n2", f"@Sens!D{r}|sp1", f"@Sens!E{r}|n2", f"@Sens!F{r}|sp1"] for r in case_rows]})
sections.append({"id": "sens", "title": "敏感度", "blocks": sb})
assert sb[0]["rows"] and sb[1]["rows"], "龍捲風圖主圖沒有達兩位數的案例"

# ---- 4. 反向 DCF ----
RV = wb["RevDCF"]
rv = lambda p: next(r for r in range(1, RV.max_row + 1) if str(RV.cell(r, 1).value or "").startswith(p))
rows_rv = lambda lab: rv(lab)
sections.append({"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2", "blocks": [
    {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
     "rows": [["預測期平均營業利益率（FY2027–FY2032）", f"@RevDCF!B{rv('現價隱含的平均營業利益率')}|p1", f"@RevDCF!B{rv('模型：FY2027–FY2032 平均')}|p1"],
              ["營收規模（FY2027 起各年同乘的倍數）", f"@RevDCF!B{rv('現價隱含的營收倍數')}|n2", "—"],
              ["FY2027 營收（十億美元）", f"@RevDCF!B{rv('現價隱含的 FY2027 營收')}|n1", f"@RevDCF!B{rv('模型 FY2027 營收')}|n1"],
              ["折現率（WACC）", f"@RevDCF!B{rv('現價隱含的 WACC')}|p2", f"@RevDCF!B{rv('模型 WACC')}|p2"],
              ["前瞻 P/E（Comparable）", f"@RevDCF!B{rv('現價隱含的前瞻 P/E')}|n1", f"@RevDCF!B{rv('同業前瞻 P/E 中位數')}|n1"]]},
    {"type": "table", "caption": "反解的中間數字（情境 B）", "head": ["項目", "值"],
     "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n3" if r in (rv("淨現金"), rv("評價用稀釋股數"), rv("支持現價所需的企業價值"), rv("模型企業價值")) else "n2")] for r in range(rv("現價（美元）"), rv("模型：情境 B 評價日每股價值") + 1)]},
    {"type": "table", "caption": f"@RevDCF!A{rv('WACC 網格')}", "collapse": True, "summary": "展開：WACC 網格（情境 B 價值推到十二個月後，美元）",
     "head": [f"@RevDCF!A{rv('WACC 網格')}"] + [f"@RevDCF!{get_column_letter(c)}{rv('WACC 網格')}|p1" for c in range(2, 19)],
     "rows": [[f"@RevDCF!A{rv('WACC 網格') + 1}"] + [f"@RevDCF!{get_column_letter(c)}{rv('WACC 網格') + 1}|n1" for c in range(2, 19)]]}]})

# ---- 5. 損益 ----
cols10 = [YC[y] for y in range(2023, 2033)]
keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用", "營業利益（", "營業利益率", "其他收入", "稅前利益", "所得稅", "非控制權益與停業", "稅後淨利", "淨利率", "股數", "EPS", "淨利年增率"]
def fm(lab):
    return "p1" if ("率" in lab or "YoY" in lab) else "n3" if lab.startswith(("EPS", "股數")) else "n2"
def prow(lab, base=rbB, h=hB):
    r = row("PL", lab, base, lab in ("營收", "稅後淨利（歸屬 ABB）"))
    return [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fm(str(PL.cell(r, 1).value))}" for c in cols10]
is_tbl = {"type": "table", "caption": "標準損益表（GAAP 口徑；十億美元，EPS 為美元）：FY2023A 未重編僅供參考、FY2024A–FY2025A 實績、FY2026E 前九個月實績加第四季模型、FY2027E–FY2032E 情境 B",
          "head": [f"@PL!A{hB}"] + [f"@PL!{c}{hB}" for c in cols10],
          "rows": [prow(k) for k in keys] + [{"sub": "Operational EBITA（ABB 調整後 Non-GAAP 指標）與橋"}] + [prow(k) for k in ("集團 Operational EBITA", "集團 Operational EBITA 率", "電氣化 Operational EBITA 率", "橋：營業利益")]}
rv_, gm_, om_, nm_ = pB("營收", True), pB("毛利率"), pB("營業利益率"), pB("淨利率")
is_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率、淨利率：FY2024A–FY2032E（情境 B）", "x": f"@PL!D{hB}:L{hB}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
            "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!D{rv_}:L{rv_}", "color": "#6fc7cb"}],
            "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!D{gm_}:L{gm_}", "axis": "r", "color": "#0b7a80", "labels": True}, {"name": f"@PL!A{om_}", "y": f"@PL!D{om_}:L{om_}", "axis": "r", "color": "#c29a2b"},
                      {"name": f"@PL!A{nm_}", "y": f"@PL!D{nm_}:L{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
            "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸）。FY2026E 前九個月為公司實績、第四季為模型；FY2027E 起為情境推導。"}
ep = lambda base, lab: row("PL", lab, base, lab == '營收')
sc_chart = {"type": "theme", "title": "三情境並列：營收（長條，十億美元）與 EPS（折線，美元，右軸）", "x": f"@PL!F{hB}:L{hB}", "fmt": "n0", "fmt_r": "n1", "stack": False, "pad": 1.15,
            "bars": [{"name": ["情境 A（順利）營收"], "y": f"@PL!F{ep(blkA, '營收')}:L{ep(blkA, '營收')}", "color": "#b7e3e4"}, {"name": ["情境 B（基準）營收"], "y": f"@PL!F{ep(rbB, '營收')}:L{ep(rbB, '營收')}", "color": "#10a0a7"},
                     {"name": ["情境 C（不如預期）營收"], "y": f"@PL!F{ep(blkC, '營收')}:L{ep(blkC, '營收')}", "color": "#8a8c8c"}],
            "lines": [{"name": ["情境 A EPS"], "y": f"@PL!F{ep(blkA, 'EPS（基本')}:L{ep(blkA, 'EPS（基本')}", "axis": "r", "color": "#0b7a80", "dash": True},
                      {"name": ["情境 B EPS"], "y": f"@PL!F{ep(rbB, 'EPS（基本')}:L{ep(rbB, 'EPS（基本')}", "axis": "r", "color": "#3b3838", "labels": True},
                      {"name": ["情境 C EPS"], "y": f"@PL!F{ep(blkC, 'EPS（基本')}:L{ep(blkC, 'EPS（基本')}", "axis": "r", "color": "#b4472f", "dash": True}],
            "caption": "三情境在 FY2026E 相同（前九個月實績加第四季模型），差異自 FY2027E 起。"}
def metric(lab, fmt_, chart=False, ex=True):
    d = {"name": [lab], "fmt": fmt_, "rows": {k: f"@PL!F{row('PL', lab, b0, ex)}:L{row('PL', lab, b0, ex)}" for k, b0 in (("A", blkA), ("B", rbB), ("C", blkC))}}
    if chart: d["chart"] = True
    return d
pl_blk = {"type": "pl", "collapse": True, "summary": "展開：三情境摘要表與折線（FY2026E–FY2032E；營收、資料中心營收、營業利益率、EPS）", "years": f"@PL!F{hB}:L{hB}",
          "metrics": [metric("營收", "n1", True), metric("電氣化：資料中心營收（拆分為本模型假設）", "n2", True), metric("營業利益率", "p1"), metric("EPS（基本；美元）", "n2", True)]}
sections.append({"id": "pl", "title": "損益（標準損益表）", "blocks": [{"type": "note", "parts": ["@PL!A1"]}, is_chart, is_tbl, sc_chart, pl_blk]})

# ---- 6. 驅動因子 ----
rw = lambda base: row("PL", "Wattonomics 配電＋UPS", base)
r_w, r_sh = pB("Wattonomics 配電＋UPS", False), pB("ABB 資料中心營收占", False)
dcols8 = [YC[y] for y in range(2025, 2033)]
drv = [{"type": "note", "parts": [f"@README!A{RD('三層基準：')}"]},
       {"type": "theme", "title": "ABB 資料中心營收（長條，十億美元）與 Wattonomics 配電加 UPS 設施支出（折線，十億美元，右軸）：情境 B", "x": f"@PL!E{hB}:L{hB}", "fmt": "n1", "fmt_r": "n0", "stack": False, "pad": 1.2,
        "bars": [{"name": ["ABB 資料中心營收（本模型拆分）"], "y": f"@PL!E{r_dc}:L{r_dc}", "color": "#0b7a80"}],
        "lines": [{"name": ["Wattonomics 配電＋UPS 設施支出"], "y": f"@PL!E{r_w}:L{r_w}", "axis": "r", "color": "#3b3838", "labels": True}],
        "caption": "兩個產業輸入：Wattonomics 設施支出（機櫃外的配電與 UPS）與 ABB 份額（假設）。FY2031 起設施支出下降，ABB 資料中心營收隨之回落。"},
       {"type": "table", "caption": "分部營收推導（情境 B；十億美元）：資料中心營收＝ABB 份額 × Wattonomics 配電＋UPS 設施支出", "head": [f"@PL!A{hB}"] + [f"@PL!{c}{hB}" for c in dcols8],
        "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{'p1' if '份額' in str(PL.cell(r, 1).value) or '占' in str(PL.cell(r, 1).value) else 'n2' if r in (r_w,) else 'n2'}" for c in dcols8]
                 for r in (pB("電氣化：非資料中心營收"), r_dc, pB("電氣化營收"), pB("馬達驅動營收"), pB("自動化營收"), pB("公司與其他及沖銷"), pB("營收", True), r_w, r_sh, pB("資料中心占電氣化營收"), r_dcg)]},
       {"type": "table", "collapse": True, "summary": "展開：Wattonomics 讀數（快照名稱、各年數值；Wat 頁）", "caption": "Wattonomics 讀數（快照；Wat 頁）", "head": ["@Wat!A4"] + [f"@Wat!{YC[y]}4" for y in range(2025, 2033)],
        "rows": [[f"@Wat!A{r}"] + [f"@Wat!{YC[y]}{r}|n2" for y in range(2025, 2033)] for r in (5, 6, 7, 8, 9, 10)]},
       {"type": "h3", "text": "輸入參數（基準進目標價；低／高只用於敏感度頁）"}]
IW = wb["In"]
FMT = {"%": "p1", "pt": "p1", "x": "n2", "US$B": "n3", "十億股": "n3", "CHF": "n2", "US$": "n2", "權重": "p0", "CHF/US$": "n3", "年": "n2", "日期": ""}
groups, cur = [], None
for r in range(4, IW.max_row + 1):
    a, b_ = IW.cell(r, 1).value, IW.cell(r, 2).value
    if a and not b_ and IW.cell(r, 3).value is None:
        cur = {"title": a, "rows": []}; groups.append(cur)
    elif a and cur is not None:
        cur["rows"].append(r)
for g in groups:
    if not g["rows"]: continue
    t = {"type": "table", "caption": f"@In!A{[r for r in range(4, IW.max_row + 1) if IW.cell(r, 1).value == g['title']][0]}", "collapse": not g["title"].startswith(("C.", "D.", "E.")),
         "head": ["@In!B3", "@In!C3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
         "rows": []}
    if t["collapse"]: t["summary"] = "展開：" + re.sub(r"^[A-Z]\d?\.\s*", "", str(g["title"])).split("（")[0]
    for r in g["rows"]:
        f = FMT.get(str(IW.cell(r, 3).value), "n2")
        if str(IW.cell(r, 1).value) == "bs_date": f = ""
        t["rows"].append([f"@In!B{r}", f"@In!C{r}", f"@In!D{r}|{f}", f"@In!E{r}|{f}", f"@In!F{r}|{f}", f"@In!G{r}|tag"])
    drv.append(t)
sections.append({"id": "drivers", "title": "驅動因子：分部營收推導", "blocks": drv})

# ---- 7. 對帳 ----
RC = wb["Recon"]
sections.append({"id": "recon", "title": "對帳：公司原文、Yahoo、市場共識與三層基準", "intro": "@Recon!A1", "blocks": [
    {"type": "h3", "text": "公司原文與 Yahoo、共識、本模型"},
    {"type": "table", "caption": "對帳表（單位依列：百萬美元、十億美元、美元、倍數）", "head": [f"@Recon!{c}3" for c in "ABCDEF"],
     "rows": [[f"@Recon!A{r}", f"@Recon!B{r}|n2", f"@Recon!C{r}|n2", f"@Recon!D{r}|n2", f"@Recon!E{r}", f"@Recon!F{r}|tag"] for r in range(4, RC.max_row + 1)]},
    {"type": "h3", "text": "三層基準：本模型讀 Wattonomics"},
    {"type": "note", "parts": ["@Coord!A1"]}, {"type": "note", "parts": ["@Coord!A2"]}, {"type": "note", "parts": ["@Coord!A10"]}, {"type": "note", "parts": ["@Coord!A11"]},
    {"type": "table", "collapse": True, "summary": "展開：讀取層級的名稱、版本與用途", "caption": "讀取層級：本模型讀哪一層、哪一版、用在哪裡", "head": [f"@Coord!{c}4" for c in "ABCDEFGHI"],
     "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}", f"@Coord!E{r}|n2", f"@Coord!F{r}|n2", f"@Coord!G{r}|n2", f"@Coord!H{r}|n2", f"@Coord!I{r}"] for r in range(5, 9)]}]})

# ---- 8. 資料狀態 ----
def readme_list(start_prefix, stop_prefixes):
    r0 = RD(start_prefix) + 1; out = []; r = r0
    ws = wb["README"]
    while r <= ws.max_row and ws.cell(r, 1).value and not str(ws.cell(r, 1).value).startswith(stop_prefixes):
        out.append(f"@README!A{r}"); r += 1
    return out
HEADS = ("待判斷", "公司原文與 Yahoo", "建議補進", "已知限制", "版本要點", "命題回覆")
sections.append({"id": "status", "title": "資料狀態", "blocks": [
    {"type": "h3", "text": "本版要點"}, {"type": "list", "items": ["@README!A1"] + readme_list("版本要點", HEADS)},
    {"type": "h3", "text": "上游成品與來源"}, {"type": "upstream", "sheets": ["In", "Wat"]},
    {"type": "table", "caption": "來源與等級", "head": [f"@Sources!{c}3" for c in "ABCDE"], "collapse": True, "summary": "展開：來源與等級（公司原文、Wattonomics、價格、Yahoo、FRED）",
     "rows": [[f"@Sources!{c}{r}" for c in "ABCDE"] for r in range(4, wb["Sources"].max_row + 1)]},
    {"type": "h3", "text": "待判斷項目"}, {"type": "list", "items": readme_list("待判斷", HEADS)},
    {"type": "h3", "text": "公司原文與 Yahoo 的差異"}, {"type": "list", "items": readme_list("公司原文與 Yahoo", HEADS)},
    {"type": "h3", "text": "建議補進 Siliconomics 與 Wattonomics"}, {"type": "list", "items": readme_list("建議補進", HEADS)},
    {"type": "h3", "text": "已知限制"}, {"type": "list", "items": readme_list("已知限制", HEADS)},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的輸入與來源", "sheets": {"In": ["B", "H", "I"], "Q": ["A", "H"], "Sources": ["A", "B"]}, "empty": "沒有標為未找到的項目。"}]})
cfg["sections"] = sections
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("html.json 已寫入", len(json.dumps(cfg)))

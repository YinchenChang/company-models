# 安川電機 html.json 產生器（個股 HTML 版面規格 v1.3；結構參考 abb/gen_html_json.py：由 xlsx 標籤定位列號，不寫死）。
# 讀 physical_ai/yaskawa/dist/ 內唯一的 xlsx（LibreOffice 重算後）→ 寫 physical_ai/yaskawa/html.json（勿手改 html.json）。
# 文字不含阿拉伯數字（年度與日期除外）；數字一律取自 xlsx。
# 用法：python3 physical_ai/yaskawa/gen_html_json.py [成品 xlsx]
import glob, json, os, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
X = sys.argv[1] if len(sys.argv) > 1 else (lambda fs: fs[0] if len(fs) == 1 else sys.exit(f"dist 內應剛好一個 xlsx：{fs}"))(sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx"))))
wb = openpyxl.load_workbook(X, data_only=True)
YRS = "CDEFGHIJK"      # PL 欄：FY2024..FY2032


def row(sh, label, start=1, exact=False, end=None):
    ws = wb[sh]
    for r in range(start, (end or ws.max_row) + 1):
        v = str(ws.cell(r, 1).value or "").strip()
        if (v == label.strip()) if exact else v.startswith(label.strip()):
            return r
    raise KeyError((sh, label))


def rd(prefix): return f"@README!A{row('README', prefix)}"


V = lambda lab, f="", c="B": f"@Val!{c}{row('Val', lab)}" + (f"|{f}" if f else "")
VA = lambda lab: f"@Val!A{row('Val', lab)}"

# ---- PL 三情境區塊 ----
blk = {s: row("PL", f"情境 {s}") for s in "ABC"}
HDR = blk["B"] + 1
assert wb["PL"].cell(HDR, 1).value == "項目"


def pr(s, lab):
    return row("PL", lab if s == "B" else f"{lab}（情境 {s}）", blk[s], True)


def plr(r, f): return [f"@PL!{c}{r}|{f}" for c in YRS]


# ================= 橫幅 =================
cfg = {"company": "安川電機", "ticker": "6506.T", "layout": "v1.3", "wide_text": True,
       "thesis": rd("命題："),
       "question": "訂單創新高之下，機器人利益率下滑是一次性還是結構性？現價隱含的利益率是否合理",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守）"}],
       "hero": {"rec": V("投資建議"), "tiles": [
           [VA("加權目標價（円）"), V("加權目標價（円）", "n0"), ["上檔空間 ", V("上檔空間", "sp1")]],
           [[VA("分部加總目標價（情境加權）"), "（円）"], V("分部加總目標價（情境加權）", "n0"), ""],
           [["現價（", V("股價基準日"), " 收盤，円）"], V("現價（円）", "n0"), ""]],
           "note": [V("投資建議句（條件式）")]}}

# ================= 1. 命題 =================
TH = wb["Thesis"]
t_h = row("Thesis", "季度受注（億日圓）")
t_mc, t_rob = row("Thesis", "運動控制 受注", exact=True), row("Thesis", "機器人 受注", exact=True)
t_roby, t_mcy = row("Thesis", "機器人 受注 YoY"), row("Thesis", "運動控制 受注 YoY")
t_mh = row("Thesis", "機器人 營業利益率")
t_rm, t_mm = t_mh + 1, row("Thesis", "運動控制 利益率")
t_test = row("Thesis", "證偽測試")
thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "季度受注（億日圓）：運動控制與機器人；折線＝受注年增（右軸）", "x": f"@Thesis!B{t_h}:K{t_h}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.3,
     "bars": [{"name": [f"@Thesis!A{t_mc}"], "y": f"@Thesis!B{t_mc}:K{t_mc}", "color": "#0b7a80"},
              {"name": [f"@Thesis!A{t_rob}"], "y": f"@Thesis!B{t_rob}:K{t_rob}", "color": "#c29a2b"}],
     "lines": [{"name": [f"@Thesis!A{t_mcy}"], "y": f"@Thesis!B{t_mcy}:K{t_mcy}", "axis": "r", "color": "#3b3838", "labels": True},
               {"name": [f"@Thesis!A{t_roby}"], "y": f"@Thesis!B{t_roby}:K{t_roby}", "axis": "r", "color": "#b4472f", "dash": True}],
     "caption": "長條＝兩大分部的季度受注；折線＝對前年同季的受注年增（右軸，前一年資料起算）。運動控制受注由半導體設備與資料中心驅動器帶動。"},
    {"type": "theme", "title": "機器人與運動控制營業利益率：實績、公司指引與下半年指引隱含值", "x": f"@Thesis!B{t_mh}:I{t_mh}", "fmt": "p1", "stack": False, "pad": 1.2,
     "lines": [{"name": ["機器人 營業利益率"], "y": f"@Thesis!B{t_rm}:I{t_rm}", "color": "#c29a2b", "labels": True},
               {"name": [f"@Thesis!A{t_mm}"], "y": f"@Thesis!B{t_mm}:I{t_mm}", "color": "#0b7a80", "labels": "below"}],
     "caption": "分部營業利益 ÷ 外部營收。命題的承重變數是機器人利益率能否回到下半年指引隱含水準以上；第一季含歐洲重組與核心系統切換的一次性費用。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [rd("命題："), "@Thesis!A2"]},
    {"type": "table", "collapse": True, "summary": "展開：命題回覆（證偽測試的最新讀數與下一個讀點）", "head": [f"@Thesis!{c}{t_test}" for c in "ABCDEF"],
     "rowgen": {"sheet": "Thesis", "rows": f"{t_test + 1}-{t_test + 4}", "cols": list("ABCDEF"), "fmts": ["", "p1", "p1", "", "", ""],
                "row_fmt": {str(t_test + 1): {"B": "n2", "C": "n2"}, str(t_test + 2): {"C": "sp1"}}}},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [rd(p) for p in ("訂單強：", "機器人利益率：", "估值：", "現價隱含：", "收支：")]},
    {"type": "note", "parts": [f"@Thesis!A{row('Thesis', '人形機器人')}"]}]}

# ================= 2. 估值 =================
vtp, vdc, vcp, vpx = (row("Val", p) for p in ("加權目標價（円）", "DCF 目標價（情境加權）", "分部加總目標價（情境加權）", "現價（円）"))
rc_cons = row("Recon", "市場對照：")
bars = {"type": "bars", "title": "目標價與現價（円；豎線＝現價）", "fmt": "n0",
        "items": [{"label": ["DCF 目標價（情境加權）"], "value": f"@Val!B{vdc}", "series": 1},
                  {"label": ["Comparable 目標價（分部加總，情境加權）"], "value": f"@Val!B{vcp}", "series": 2},
                  {"label": ["加權目標價（橫幅主值）"], "value": f"@Val!B{vtp}", "series": 0}],
        "ref": {"label": f"@Val!A{vpx}", "value": f"@Val!B{vpx}"}}
conc_rows = [("DCF 目標價（情境加權）", "n0"), ("分部加總目標價（情境加權）", "n0"), ("DCF 權重", "p0"), ("加權目標價（円）", "n0"), ("現價（円）", "n0"), ("上檔空間", "sp1"),
             ("估值評等（只看上檔空間）", ""), ("距最近級距邊界（百分點）", "n1"), ("投資建議", ""), ("股價基準日", ""), ("評價日", ""), ("目標時點", "")]
conc = {"type": "table", "caption": "目標價、現價與投資建議（円）", "head": ["項目", "數值"],
        "rows": [[f"@Val!A{row('Val', l)}", V(l, f)] for l, f in conc_rows] + [["市場共識目標價", f"@Recon!A{rc_cons}"]]}
scen = {"type": "table", "caption": "三情境價值（円）與權重", "head": [f"@Val!{c}3" for c in "ABCDE"],
        "rowgen": {"sheet": "Val", "rows": "4-6", "cols": list("ABCDE"), "fmts": ["", "n0", "n0", "n0", "n0"], "row_fmt": {"4": {c: "p0" for c in "BCDE"}}}}
attr = {"type": "table", "caption": "上檔空間三段歸因（百分點）", "head": ["項目", "數值", "說明"],
        "rows": [[f"@Val!A{r}", f"@Val!B{r}|n1", f"@Val!C{r}"] for r in range(row("Val", "上檔空間三段歸因：DCF"), row("Val", "上檔空間三段歸因：相對倍數") + 1)]}
D = wb["DCF"]
rates = {"type": "table", "caption": "DCF 折現率：股權成本、稅後債務成本、WACC、永續成長率、RONIC", "head": ["項目", "數值"],
         "rows": [[f"@DCF!A{row('DCF', l, exact=True)}", f"@DCF!B{row('DCF', l, exact=True)}|p2"] for l in ("股權成本 ke", "稅後債務成本", "權益占資本", "WACC")]
         + [["永續成長率（取上限與無風險利率較小者）", f"@DCF!B{row('DCF', '永續成長率 g')}|p2"], [f"@DCF!A{row('DCF', 'RONIC', exact=True)}", f"@DCF!B{row('DCF', 'RONIC', exact=True)}|p2"]]}
rates_x = {"type": "table", "collapse": True, "summary": "展開：折現率組成（市值、負債、重新加槓桿 β）", "head": ["項目", "數值"],
           "rows": [[f"@DCF!A{row('DCF', l)}", f"@DCF!B{row('DCF', l)}|{f}"] for l, f in (("股權市值 E", "n1"), ("負債 D", "n1"), ("β（重新加槓桿）", "n3"))]}
dsc = {s: row("DCF", f"情境 {s}", exact=True) for s in "ABC"}


def drow(s, lab): return row("DCF", lab, dsc[s])


main_d = [("企業價值 EV", "n1"), ("淨現金滾到評價日", "n1"), ("每股價值（評價日，円）", "n0"), ("DCF 情境價值推到 12 個月後（円）", "n0"), ("終值占 EV", "p1")]
more_d = [("明確期現值", "n1"), ("終值現值", "n1"), ("＋淨現金與投資（資產負債表）", "n1"), ("＋Stub 期 FCFF", "n1"), ("− Stub 期已宣告股利", "n1"), ("股權價值", "n1"), ("期間每股股利", "n1"), ("檢查：終值占 EV", "")]
dsum = lambda items: [[f"@DCF!A{drow('B', l)}"] + [f"@DCF!B{drow(s, l)}|{f}" for s in "ABC"] for l, f in items]
dcf_t = {"type": "table", "caption": "DCF 彙總（十億日圓，每股為円）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rows": dsum(main_d)}
dcf_x = {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rows": dsum(more_d)}
dcf_y = []
for s in "ABC":
    h = dsc[s] + 1
    dcf_y.append({"type": "table", "collapse": True, "summary": f"展開：情境 {s} 各年 FCFF 與折現", "head": [f"@DCF!{c}{h}" for c in "ABCDEFGHI"],
                  "rowgen": {"sheet": "DCF", "rows": f"{h + 1}-{h + 4}", "cols": list("ABCDEFGHI"), "fmts": [""] + ["n1"] * 8,
                             "row_fmt": {str(h + 2): {c: "n3" for c in "BCDEFGHI"}, str(h + 3): {c: "n3" for c in "BCDEFGHI"}}}})
csc = {s: row("Comp", f"情境 {s}", exact=True) for s in "ABC"}
comp_main = {"type": "table", "caption": "分部加總（Comparable）：企業價值與每股價值（十億日圓，每股為円；倍數為假設值）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
             "rows": [[f"@Comp!A{row('Comp', l, csc['B'])}"] + [f"@Comp!D{row('Comp', l, csc[s])}|{f}" for s in "ABC"]
                      for l, f in (("企業價值（分部加總）", "n1"), ("＋淨現金滾到評價日", "n1"), ("隱含 EV ÷ NTM EBIT", "x1"), ("分部加總每股價值（円）", "n0"))]}
comp_x = []
for s in "ABC":
    h = csc[s] + 1
    comp_x.append({"type": "table", "collapse": True, "summary": f"展開：情境 {s} 分部 NTM EBIT 與倍數", "head": [f"@Comp!{c}{h}" for c in "ABCD"],
                   "rowgen": {"sheet": "Comp", "rows": f"{h + 1}-{h + 5}", "cols": list("ABCD"), "fmts": ["", "n1", "n1", "n1"]}})
valuation = {"id": "valuation", "title": "估值", "blocks": [bars, {"type": "note", "parts": ["@Val!A1"]}, {"type": "note", "parts": ["@Comp!A1"]}, conc, scen, attr, rates, rates_x, dcf_t, dcf_x] + dcf_y
             + [{"type": "note", "parts": ["@DCF!A1"], "collapse": True, "summary": "展開：DCF 口徑（評價日、剩餘期、終值式）"}, comp_main] + comp_x}

# ================= 3. 敏感度 =================
SE = wb["Sens"]
s_h = row("Sens", "驅動", exact=True)
assert SE.cell(s_h, 6).value == "低案 相對基準" and SE.cell(s_h, 7).value == "高案 相對基準"
s_end = row("Sens", "檢查：") - 1
case_rows = [r for r in range(s_h + 2, s_end + 1) if SE.cell(r, 1).value]
sens_b = []
small_b = []
for side, col in (("低", "F"), ("高", "G")):
    ci = 6 if col == "F" else 7
    rs = [r for r in case_rows if isinstance(SE.cell(r, ci).value, (int, float))]
    big = [r for r in rs if abs(SE.cell(r, ci).value) >= 0.1]
    small = [r for r in rs if 0.0005 <= abs(SE.cell(r, ci).value) < 0.1]
    base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
    if big:
        sens_b.append(dict(base, title=f"加權目標價相對基準：輸入取{side}案，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
    else:
        top = sorted(rs, key=lambda r: -abs(SE.cell(r, ci).value))[:1]
        sens_b.append(dict(base, title=f"加權目標價相對基準：輸入取{side}案，沒有變動達兩位數百分比的案例，列影響最大者", rows=",".join(map(str, top))))
    if small:
        small_b.append(dict(base, title=f"加權目標價相對基準：輸入取{side}案，其餘案例", rows=",".join(map(str, small)), collapse=True, summary=f"展開：輸入取{side}案，變動未達兩位數百分比的案例"))
sens_all = {"type": "table", "collapse": True, "summary": "展開：全部案例（加權目標價、DCF 單法與估值評等）", "head": [f"@Sens!{c}{s_h}" for c in "ABDEFGHIJK"],
            "rowgen": {"sheet": "Sens", "rows": f"{s_h + 1}-{s_end}", "cols": list("ABDEFGHIJK"), "fmts": ["", "", "n0", "n0", "sp1", "sp1", "n0", "n0", "", ""]}}
s_chk = row("Sens", "檢查：")
sens = {"id": "sens", "title": "敏感度", "blocks": sens_b + [{"type": "note", "parts": ["讀法：數值為整本重算的快照，一次只動一個參數；下方核對列比對快照基準與估值頁即時加權目標價。舊收盤價只作對照、不進模型；情境專屬參數的樂觀與保守值已由三情境加權涵蓋，不另列。"]}] + small_b + [sens_all,
        {"type": "kv", "collapse": True, "summary": "展開：快照核對與 build_id", "items": [[f"@Sens!A{s_chk}", [f"@Sens!D{s_chk}|n2", "（", f"@Sens!F{s_chk}", "）"]], ["build_id", "@Sens!A1"]]}]}

# ================= 4. 反向 DCF（反解） =================
rs_h = row("Recon", "驅動", exact=True)
rlab = {"rob_m_B": ("機器人穩態利益率（情境 B）", "p1"), "mc_m_B": ("運動控制穩態利益率（情境 B）", "p1"), "beta_u": ("無槓桿 β", "n2"), "mult_rob": ("機器人分部倍數（EV ÷ NTM EBIT）", "n1")}
rv_rows = []
for r in range(rs_h + 1, rs_h + 5):
    lab, f = rlab[wb["Recon"].cell(r, 1).value]
    rv_rows.append([lab] + [f"@Recon!{c}{r}|{f}" for c in "BCDE"])
revdcf = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "blocks": [
    {"type": "note", "parts": [rd("現價隱含：")]},
    {"type": "table", "caption": f"@Recon!A{rs_h - 1}", "head": ["驅動"] + [f"@Recon!{c}{rs_h}" for c in "BCDE"], "rows": rv_rows},
    {"type": "note", "parts": [f"@Recon!F{rs_h + 1}"]}]}

# ================= 5. 損益 =================
B = lambda lab: pr("B", lab)
std = [("營收（合併，四個分部加總）", B("合併營收"), "n1"), ("營業利益（分部營業利益＋全社調整額）", B("合併營業利益（EBIT）"), "n1"), (None, B("合併營業利益率"), "p1"),
       ("業外淨額（其他收入淨額）", B("業外淨額"), "n1"), (None, B("稅前利益"), "n1"), (None, B("所得稅"), "n1"), (None, B("稅後淨利"), "n1"), (None, B("非支配股東損益"), "n1"),
       (None, B("歸屬母公司淨利"), "n1"), ("EPS（基本，円；模型未建稀釋 EPS）", B("基本 EPS（円）"), "n2"), (None, B("每股股利（円）"), "n2")]
std_rows = [[lab if lab else f"@PL!A{r}"] + plr(r, f) for lab, r, f in std]
met = lambda name, lab, f, chart=False: dict({"name": name, "fmt": f, "rows": {s: f"@PL!C{pr(s, lab)}:K{pr(s, lab)}" for s in "ABC"}}, **({"chart": True} if chart else {}))
pl = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "合併營收（十億日圓）與營業利益率：FY2024A–FY2032E（情境 B）", "x": f"@PL!C{HDR}:K{HDR}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.4,
     "bars": [{"name": [f"@PL!A{B('合併營收')}"], "y": f"@PL!C{B('合併營收')}:K{B('合併營收')}", "color": "#6fc7cb"}],
     "lines": [{"name": [f"@PL!A{B('合併營業利益率')}"], "y": f"@PL!C{B('合併營業利益率')}:K{B('合併營業利益率')}", "axis": "r", "color": "#0b7a80", "labels": True},
               {"name": [f"@PL!A{B('機器人 利益率')}"], "y": f"@PL!C{B('機器人 利益率')}:K{B('機器人 利益率')}", "axis": "r", "color": "#c29a2b", "dash": True}]},
    {"type": "note", "parts": ["口徑：IFRS（國際財務報導準則，公司法定揭露口徑；非美國 GAAP，也不是 Non-GAAP 調整數）。安川以分部營業利益揭露，模型由分部利益直接加總到營業利益，未建營業成本、毛利與營業費用列。", "@PL!A1"]},
    {"type": "table", "caption": "標準損益表（IFRS；十億日圓，每股為円）：FY2024A–FY2025A 實績、FY2026E 起情境 B", "head": [f"@PL!A{HDR}"] + [f"@PL!{c}{HDR}" for c in YRS], "rows": std_rows},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線）", "years": f"@PL!C{HDR}:K{HDR}",
     "metrics": [met("合併營收（十億日圓）", "合併營收", "n1", True), met("合併營業利益（十億日圓）", "合併營業利益（EBIT）", "n1", True),
                 met("合併營業利益率", "合併營業利益率", "p1"), met("基本 EPS（円）", "基本 EPS（円）", "n2")]},
    {"type": "table", "collapse": True, "summary": "展開：現金流（折舊、資本支出、營運資金、FCFF、股利；情境 B）", "head": [f"@PL!A{HDR}"] + [f"@PL!{c}{HDR}" for c in YRS],
     "rowgen": {"sheet": "PL", "rows": f"{B('折舊攤銷')}-{B('FCFF − 股利')}", "cols": ["A"] + list(YRS), "fmts": [""] + ["n1"] * 9}}]}

# ================= 6. 驅動 =================
seg = [("運動控制 營收", "#0b7a80"), ("機器人 營收", "#c29a2b"), ("系統工程 營收", "#6fc7cb"), ("其他 營收", "#8a8c8c")]
r_seg0, r_segm = B("運動控制 營收"), B("合併營業利益率")
pct_rows = {B(l) for l in ("運動控制 營收年增", "機器人 營收年增", "運動控制 利益率", "機器人 利益率", "合併營業利益率")}
IN = wb["In"]
INR = {str(IN.cell(r, 1).value): r for r in range(4, row("In", "Z.")) if IN.cell(r, 1).value and IN.cell(r, 3).value}
z0 = row("In", "Z.")
UF = {"%": "p1", "¥B": "n3", "円": "n0", "倍": "n2", "百萬股": "n3", "日期": ""}


def in_row(k):
    r = INR[k]
    f = UF[IN.cell(r, 3).value]
    return [f"@In!B{r}", f"@In!D{r}|{f}", f"@In!E{r}|{f}", f"@In!F{r}|{f}", f"@In!G{r}|tag"]


def is_range(k):
    r = INR[k]
    return IN.cell(r, 4).value is not None or IN.cell(r, 6).value is not None


groups = [("一、運動控制", lambda k: k.startswith("mc_") or k == "mult_mc"),
          ("二、機器人", lambda k: k.startswith("rob_") or k == "mult_rob"),
          ("三、系統工程、其他與全社調整", lambda k: k.split("_")[0] in ("se", "oth", "corp", "g") and k != "g_cap" or k in ("mult_se", "mult_oth")),
          ("四、全公司：指引、資產負債、損益表以下、現金流、折現率與估值權重", lambda k: True)]
seen = set()
drv_blocks = []
for title, pred in groups:
    ks = [k for k in INR if k not in seen and pred(k)]
    seen.update(ks)
    rng, fix = [k for k in ks if is_range(k)], [k for k in ks if not is_range(k)]
    drv_blocks.append({"type": "h3", "text": title})
    head = ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"]
    if rng:
        drv_blocks.append({"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": head, "rows": [in_row(k) for k in rng]})
    if fix:
        drv_blocks.append({"type": "table", "collapse": True, "summary": "展開：固定輸入（實績、公司揭露、情境專屬值與換算值）", "head": head, "rows": [in_row(k) for k in fix]})
assert seen == set(INR), set(INR) - seen
drivers = {"id": "drivers", "title": "驅動因子 → 推導量：分部營收與利益", "blocks": [
    {"type": "theme", "title": "分部營收（十億日圓，堆疊）與合併營業利益率：FY2024A–FY2032E（情境 B）", "x": f"@PL!C{HDR}:K{HDR}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.25,
     "total": f"@PL!C{B('合併營收')}:K{B('合併營收')}",
     "bars": [{"name": [f"@PL!A{B(l)}"], "y": f"@PL!C{B(l)}:K{B(l)}", "color": c} for l, c in seg],
     "lines": [{"name": [f"@PL!A{r_segm}"], "y": f"@PL!C{r_segm}:K{r_segm}", "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝四個官方應報導部門的外部營收，頂端數字＝合併營收（模型合計列）；折線＝合併營業利益率（右軸）。FY2026 以公司分部預估與指引為錨，FY2027 起由情境驅動。"},
    {"type": "note", "parts": [f"@PL!B{B('系統工程 營收')}"]},
    {"type": "table", "caption": "分部營收與利益推導（情境 B；十億日圓）", "head": [f"@PL!A{HDR}"] + [f"@PL!{c}{HDR}" for c in YRS],
     "rowgen": {"sheet": "PL", "rows": f"{r_seg0}-{r_segm}", "cols": ["A"] + list(YRS), "fmts": [""] + ["n1"] * 9,
                "row_fmt": {**{str(r): {c: "p1" for c in YRS} for r in pct_rows}, str(B("檢查：分部合計 − 合併營收")): {c: "n3" for c in YRS}}}},
    {"type": "table", "caption": f"@In!A{z0}", "head": [f"@In!{c}{z0 + 1}" for c in "BDEF"],
     "rows": [[f"@In!B{r}", f"@In!D{r}|p1", f"@In!E{r}|p1", f"@In!F{r}|p1"] for r in range(z0 + 2, z0 + 8)] + [[f"@In!B{z0 + 8}", f"@In!D{z0 + 8}|p0", f"@In!E{z0 + 8}|p0", f"@In!F{z0 + 8}|p0"]]},
    {"type": "note", "parts": ["@In!A1"]}] + drv_blocks}

# ================= 8. 對帳 =================
rc_h = row("Recon", "項目", exact=True)
rc_rows = list(range(rc_h + 1, rc_cons - 1))
pct = [r for r in rc_rows if "率" in str(wb["Recon"].cell(r, 1).value)]
rb_h = row("對帳", "變數名")
rb_rows = list(range(rb_h + 1, wb["對帳"].max_row + 1))
rbf = {}
for r in rb_rows:
    v, u = wb["對帳"].cell(r, 3).value, wb["對帳"].cell(r, 4).value
    rbf[str(r)] = {"C": "p1" if u == "%" else "n0"}
recon = {"id": "recon", "title": "對帳：公司指引、中期計畫、市場與線層級變數", "blocks": [
    {"type": "note", "parts": ["@Recon!A1"]},
    {"type": "table", "caption": "模型（情境 B）對公司指引與中期計畫", "head": [f"@Recon!{c}{rc_h}" for c in "ABCD"],
     "rowgen": {"sheet": "Recon", "rows": ",".join(map(str, rc_rows)), "cols": list("ABCD"), "fmts": ["", "n1", "n1", "n1"], "row_fmt": {str(r): {c: "p1" for c in "BCD"} for r in pct}}},
    {"type": "table", "collapse": True, "summary": "展開：各對帳列的說明（含網址的出處列見 xlsx 對帳頁）", "head": [f"@Recon!{c}{rc_h}" for c in "AE"],
     "rowgen": {"sheet": "Recon", "rows": ",".join(str(r) for r in rc_rows if "http" not in str(wb["Recon"].cell(r, 5).value)), "cols": list("AE")}},
    {"type": "note", "parts": [f"@Recon!A{rc_cons}"]},
    {"type": "table", "caption": "@對帳!A1", "head": [f"@對帳!{c}{rb_h}" for c in "ABCG"],
     "rowgen": {"sheet": "對帳", "rows": f"{rb_rows[0]}-{rb_rows[-1]}", "cols": list("ABCG"), "fmts": ["", "", "", "tag"], "row_fmt": rbf}},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數的口徑與影響儲存格（出處見 xlsx 對帳頁）", "head": [f"@對帳!{c}{rb_h}" for c in "AEH"],
     "rowgen": {"sheet": "對帳", "rows": f"{rb_rows[0]}-{rb_rows[-1]}", "cols": list("AEH")}}]}

# ================= 9. 資料狀態 =================
rk_h = row("Risk", "風險", exact=True)
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與 build_id", "@README!A1"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "h3", "text": "日期、口徑與限制"},
    {"type": "list", "items": ["@README!A2", "@README!A7", "@README!A8", "@README!A9", "@README!A5", "@README!A6"]},
    {"type": "table", "caption": "@Risk!A1", "head": [f"@Risk!{c}{rk_h}" for c in "ABCD"],
     "rowgen": {"sheet": "Risk", "rows": f"{rk_h + 1}-{wb['Risk'].max_row}", "cols": list("ABCD")}},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"In": ["B"], "Thesis": ["A"]}, "empty": "沒有標為未取得的項目。"}]}

cfg["sections"] = [thesis, valuation, sens, revdcf, pl, drivers, recon, status]
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", os.path.join(HERE, "html.json"))

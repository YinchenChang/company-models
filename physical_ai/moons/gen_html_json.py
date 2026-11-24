# 產生 physical_ai/moons/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/moons/gen_html_json.py [成品 xlsx]（預設 dist/ 內最新的 xlsx）→ physical_ai/leaderdrive/html.json
import glob, json, os, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
X = sys.argv[1] if len(sys.argv) > 1 else sorted(f for f in glob.glob(os.path.join(HERE, "dist", "*.xlsx")))[-1]
wb = openpyxl.load_workbook(X, data_only=True)
P, V, I, D, SE, RC, TH = wb["PL"], wb["Val"], wb["In"], wb["DCF"], wb["Sens"], wb["Recon"], wb["Thesis"]


def find(ws, text, start=1, col=1, exact=True):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and (v == text if exact else v.startswith(text)):
            return r
    raise KeyError(f"{ws.title}: {text}")


IN = {}
for r in range(4, I.max_row + 1):
    k = I.cell(r, 1).value
    if isinstance(k, str) and k not in IN: IN[k] = r
def iv(k, f=""): return f"@In!E{IN[k]}" + (f"|{f}" if f else "")
def il(k): return f"@In!B{IN[k]}"
def nourl(sh, coord, f=""):
    v = wb[sh][coord].value
    return "（出處含網址，見 xlsx）" if isinstance(v, str) and "http" in v else f"@{sh}!{coord}" + (f"|{f}" if f else "")

# PL 三情境區塊
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm); pre = "" if s == "B" else f"（情境 {s}）"
    lab = lambda t, h=h, pre=pre: find(P, t + pre, h)
    PR[s] = dict(hdr=h + 1, ind=lab("控制電機：非機器人用（子列）"), robo=lab("控制電機：機器人應用（非人形，子列）"), hum=lab("控制電機：人形空心杯用（子列）"),
                 cm=lab("控制電機及其驅動系統 營收"), oth=lab("其他板塊 營收（電源與照明、設備狀態管理、貿易）"), rev=lab("營收（合併）"),
                 yoy=lab("營收年增率"), cogs=lab("營業成本"), gp=lab("毛利"), gm=lab("毛利率"), opex=lab("營業費用（銷售＋管理＋研發，含稅金及附加與其他營業損益）"),
                 op=lab("營業利益（核心，不含財務費用與投資收益）"), opm=lab("營業利益率"),
                 nonop=lab("其他收入淨額（業外：財務費用、投資收益、營業外）"), pbt=lab("稅前利益"), tax=lab("所得稅"), nit=lab("稅後淨利（含少數股東）"),
                 nci=lab("少數股東損益"), ni=lab("歸屬母公司淨利"), nim=lab("淨利率"), niy=lab("淨利年增率"), sh=lab("股數（百萬股）"), eps=lab("EPS（元）"),
                 da=lab("折舊攤銷"), cx=lab("資本支出"), fcff=lab("FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）"), gate=lab("FCFF 扣股利（收支）"))
B = PR["B"]
ALL = "CDEFGHIJKL"     # FY2023 … FY2032
FC = "FGHIJKL"         # FY2026 … FY2032
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

# Thesis 情境 B 區塊
TB = find(TH, "情境 B")
TR = {k: find(TH, t, TB) for k, t in dict(ind="控制電機：非機器人用 營收", robo="控制電機：機器人應用（非人形）營收", hum="控制電機：人形空心杯用 營收", oth="其他板塊（電源與照明、設備狀態管理、貿易）",
                                       cm="控制電機及其驅動系統 合計", humsh="人形空心杯用占合併營收", hsh="人形空心杯份額", px="空心杯單價（元／顆）", ship="全球人形出貨（線 4 × 出貨倍數，千台）",
                                       lmk="線模型口徑的全球人形空心杯市場（百萬元）").items()}
THDR = 3

VR = {k: find(V, t) for k, t in dict(dcf3="DCF 情境價值推到 12 個月後", eps27="FY2027 EPS（模型）", cmpu3="Comparable 未調整（同業 P/E × FY2027 EPS）",
                                       cmp3="Comparable（× 相對倍數係數）", wt="情境目標價：DCF × w ＋ Comparable ×（1 − w）", w="情境權重",
                                       ev="企業價值 EV（百萬元）", tvs="終值占 EV",
                                       dcf="DCF 目標價（情境加權）", cmp="Comparable 目標價（情境加權；同業 P/E × 相對倍數係數）", cmpu="Comparable 未調整（情境加權）",
                                       tp="加權目標價（DCF 與 Comparable 加權）", px="現價（股價基準日）",
                                       up="上檔空間：加權目標價", rt="估值評等（只看上檔空間）", gate="收支關卡：FY2026–FY2030 FCFF 扣股利合計（情境 A／B／C，百萬元）",
                                       gb="入不敷出判定用（情境 B 合計）", gj="入不敷出判定", rec="投資建議", dist="距最近級距邊界（百分點）", upb="升一級所需目標價（上一級邊界）",
                                       a1="DCF 貢獻＝w ×（DCF − 現價）÷ 現價", a2="未調整 Comparable 貢獻＝(1 − w) ×（未調整 − 現價）÷ 現價", a3="相對倍數調整貢獻",
                                       relnow="本公司現行前瞻倍數 ÷ 同業中位數（資料商 NTM 口徑）", relreq="支持現價所需的相對倍數係數（DCF 不變）",
                                       relflip="評等升一級所需的相對倍數係數", pei="隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2027 EPS",
                                       ke="股權成本 ke＝rf＋β ×（ERP＋CRP）", kd="稅後債務成本", wacc="WACC", g="永續成長率 g", ronic="RONIC",
                                       tm="模型折現期間（評價日→目標時點，年）", th="持有期間（股價基準日→目標時點，年；上檔空間不是年化報酬）").items()}
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")

DB = {s: find(D, f"情境 {s}") for s in "ABC"}
DK = {k: t for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", nc="淨現金滾到評價日（現金 − 有息負債 − 少數股東權益 ＋ Stub FCFF）",
                                 eq="股權價值", ps="每股價值（評價日，元）", v12="推到目標時點（×（1＋ke）^t − 期間股利）", tvs="終值占 EV").items()}
def dsc(k): return [f"@DCF!B{find(D, DK[k], DB[s])}" for s in "ABC"]

# Sens
s_hdr = find(SE, "驅動")
s_rows = [r for r in range(s_hdr + 1, SE.max_row + 1) if isinstance(SE.cell(r, 6).value, (int, float))]
lo = [r for r in s_rows if str(SE.cell(r, 1).value).endswith("（敏感度低）")]
hi = [r for r in s_rows if str(SE.cell(r, 1).value).endswith("（敏感度高）")]
oth = [r for r in s_rows if r not in lo and r not in hi]
TOR = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": "F", "fmt": "sp1"}
big = lambda rs: [r for r in rs if abs(SE.cell(r, 6).value) >= 0.1]
def main_tor(side, rs):
    b = big(rs)
    if b: return dict(TOR, title=f"加權目標價相對基準：輸入取{side}值（變動達兩位數百分比的案例）", rows=",".join(map(str, b)))
    return dict(TOR, title=f"加權目標價相對基準：輸入取{side}值（沒有變動達兩位數百分比的案例；列出影響最大者）", rows=str(sorted(rs, key=lambda r: -abs(SE.cell(r, 6).value))[0]))
sens_blocks = [main_tor("低", lo), main_tor("高", hi), {"type": "note", "parts": ["@Sens!A1"]},
               {"type": "kv", "items": [["@Sens!A2", "@Sens!D2"], ["@Sens!A3", "@Sens!D3"]]},
               dict(TOR, title="加權目標價相對基準：情境 B 改為情境 A／C 值、人形出貨倍數、情境權重", rows=",".join(map(str, oth)))]
for side, rs in (("低", lo), ("高", hi)):
    shown = set(map(int, main_tor(side, rs)["rows"].split(",")))
    rest = [r for r in rs if r not in shown]
    if rest:
        sens_blocks.append(dict(TOR, title=f"加權目標價相對基準：輸入取{side}值，其餘案例", rows=",".join(map(str, rest)), collapse=True, summary=f"展開：輸入取{side}值、變動未達兩位數百分比的案例"))
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：全部案例（加權目標價、DCF 目標價與估值評等）",
                    "head": [f"@Sens!{c}{s_hdr}" for c in "ACEFGHI"],
                    "rowgen": {"sheet": "Sens", "rows": f"{s_rows[0]}-{s_rows[-1]}", "cols": list("ACEFGHI"), "fmts": ["", "", "n1", "sp1", "n1", "sp1", ""]}})
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}

sol0 = find(RC, "反解目標")
SOL1 = sol0 + 1
hero = {"rec": vv("rec"),
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n1"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n1"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n1"), ""],
                  [il("px"), iv("px", "n2"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等", vv("rt"), "；收支關卡", vv("gj"), "（情境 B 自由現金流扣股利，FY2026–FY2030 合計 ", vv("gb", "B", "n1"),
                 " 百萬元）。人形空心杯用營收占比小，估值取決於本業毛利率與費用率：其他輸入不變時，支持現價需要全球人形出貨達線模型基準的 ", f"@Recon!B{SOL1}|n1",
                 " 倍；單一輸入取區間兩端都不改變評等。"]}

RD = wb["README"]
thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "合併營收依官方分部堆疊（百萬元）與毛利率：FY2024–FY2032（情境 B）", "x": f"@PL!D{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev", "B", "D"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "D"), "color": c} for k, c in (("cm", "#0b7a80"), ("oth", "#c29a2b"))],
     "lines": [{"name": [f"@PL!A{B['gm']}"], "y": pl("gm", "B", "D"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝官方分部（控制電機及驅動系統與其他板塊，合計等於合併營收；FY2023 分部未取得，圖從 FY2024 起）；折線＝毛利率（右軸）。FY2024–FY2025 為實績，FY2026 為一至九月實績加第四季假設，FY2027 起為模型。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A2",
        ["承重變數：FY2030 毛利率（情境 B ", iv("gm30_B", "p1"), "）、營業費用率降到 ", iv("opex30", "p1"), "、人形空心杯：線模型每台 ", iv("n_cup", "n0"), " 顆與單價年降 ", iv("pd", "p0"), "、情境 B 的 2030 年份額 ", iv("hs30_B", "p0"), "。"]]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A3"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n1"), " 元（DCF ", vv("dcf", "B", "n1"), "、同業本益比 ", vv("cmp", "B", "n1"), "），對現價 ", iv("px", "n2"), " 元 ", vv("up", "B", "sp1"),
         "；賣方目標價未取得。"],
        ["人形量級：情境 B 人形空心杯用營收由 FY2026 的 ", f"@Thesis!E{TR['hum']}|n0", " 百萬元增至 FY2032 的 ", f"@Thesis!K{TR['hum']}|n0", " 百萬元，占合併營收 ", f"@Thesis!K{TR['humsh']}|p1",
         "；支持現價需人形出貨達線模型基準的 ", f"@Recon!B{SOL1}|n0", " 倍。"],
        ["獲利：毛利率 FY2026 ", iv("gm26", "p1"), "（一至九月實績），營業費用率 ", iv("opex26", "p1"), " 降到 ", iv("opex30", "p1"), "；FY2026–FY2030 自由現金流扣股利合計為負，資本支出與營運資金吃掉現金。"],
        ["倍數：本公司現價對資料商前瞻盈餘的倍數是同業中位數的 ", vv("relnow", "B", "n1"), " 倍；Comparable 要支持現價，相對倍數係數須達 ", vv("relreq", "B", "n1"), "。"],
        "@README!A7"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A9"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "三個目標價與現價（元；豎線＝現價）", "fmt": "n1",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}],
     "ref": {"label": f"@Val!A{VR['px']}", "value": vv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n1"), ("cmp", "n1"), ("tp", "n1"), ("px", "n2"), ("up", "sp1"), ("rt", ""), ("gb", "n1"), ("gj", ""), ("rec", ""), ("dist", "p1"), ("upb", "n1"))]},
    {"type": "table", "caption": "情境價值（元／股，目標時點）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@Val!A{VR[k]}"] + [f"@Val!{c}{VR[k]}|{f}" for c in "BCD"] for k, f in (("dcf3", "n1"), ("eps27", "n2"), ("cmpu3", "n1"), ("cmp3", "n1"), ("wt", "n1"), ("w", "p0"))]},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@Val!A{VR[k]}", vv(k, "B", "p2")] for k in ("ke", "kd", "wacc", "ronic", "g")]},
    {"type": "table", "caption": "DCF 彙總（百萬元；每股為元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{find(D, DK[k], DB['B'])}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n0"), ("nc", "n0"), ("eq", "n0"), ("v12", "n1"), ("tvs", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{find(D, DK[k], DB['B'])}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n0"), ("tv", "n0"), ("pvtv", "n0"), ("ps", "n1"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n2"), ("erp", "p1"), ("crp", "p2"), ("sh_v", "n2"), ("cash_v", "n0"))]},
    {"type": "table", "caption": "上檔三段歸因與隱含倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("relnow", "n2"), ("relreq", "n2"), ("relflip", "n2"), ("pei", "n1"), ("tm", "n2"), ("th", "n2"))]},
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻本益比（股價皆為股價基準日收盤）", "head": [f"@Comp!{c}4" for c in "ABEFHJ"],
     "rowgen": {"sheet": "Comp", "rows": "5-11", "cols": list("ABEFHJ"), "fmts": ["", "", "", "n1", "n2", ""]}},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{IN[k]}|{f}", iv(k, f), f"@In!F{IN[k]}|{f}", f"@In!G{IN[k]}|tag"] for k, f in (("w_dcf", "p0"), ("tp_wA", "p0"), ("tp_wB", "p0"), ("tp_wC", "p0"), ("rel_k", "n2"), ("pe_peer", "n1"))]}]}

revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價、或落在升一級邊界所需的值，其他輸入不變。負一表示在搜尋區間內無解。"]},
    {"type": "table", "caption": f"@Recon!A{sol0 - 1}", "head": [f"@Recon!{c}{sol0}" for c in "ABCE"],
     "rowgen": {"sheet": "Recon", "rows": f"{sol0 + 1}-{RC.max_row}", "cols": list("ABCE"), "fmts": ["", "n2", "n2", ""]}},
    {"type": "table", "caption": "相對倍數反解", "head": ["項目", "數值"], "rows": [[f"@Val!A{VR[k]}", vv(k, "B", "n2")] for k in ("relnow", "relreq", "relflip")]}]}

STD = list(ALL)
def prow(k, f, lab=None):
    return [lab or f"@PL!A{B[k]}"] + [f"@PL!{c}{B[k]}|{f}" for c in STD]
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "合併營收（百萬元）× 情境 A／B／C 與營業利益率（情境 B）：FY2023–FY2032", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：合併營收"], "y": pl("rev", s), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['opm']}", "（情境 B）"], "y": pl("opm"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境合併營收；折線＝情境 B 營業利益率（右軸）。歷史與 FY2026 三情境相同，分化從 FY2027 開始。"},
    {"type": "note", "parts": ["會計口徑：GAAP（中國會計準則，公司法定財報口徑；A 股市場與券商皆用此口徑），未做 Non-GAAP 調整。營業利益為核心口徑：不含財務費用與投資收益，這些列在其他收入淨額（業外）；營業費用含稅金及附加與其他營業損益。2023–2025 年稅前、淨利、EPS 為公司年報摘要，毛利與核心營業利益為資料商整理值。"]},
    {"type": "table", "caption": "標準損益表（GAAP，中國會計準則；百萬元，EPS 為元）：FY2023–FY2025 實績、FY2026 一至九月實績加第四季假設、FY2027 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [prow("rev", "n1"), prow("yoy", "p1"), prow("cogs", "n1"), prow("gp", "n1"), prow("gm", "p1"), prow("opex", "n1"),
              prow("op", "n1"), prow("opm", "p1"), prow("nonop", "n1"), prow("pbt", "n1"), prow("tax", "n1"), prow("nit", "n1"), prow("nci", "n1"), prow("ni", "n1"),
              prow("sh", "n2"), prow("eps", "n2"), prow("nim", "p1"), prow("niy", "p1")]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "合併營收（百萬元）", "fmt": "n0", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "營業利益（百萬元）", "fmt": "n0", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "毛利率", "fmt": "p1", "rows": {s: pl("gm", s, "F") for s in "ABC"}},
                 {"name": "EPS（元）", "fmt": "n2", "chart": True, "rows": {s: pl("eps", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣股利（百萬元；收支關卡用）", "fmt": "n0", "rows": {s: pl("gate", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流（百萬元）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in FC],
     "rowgen": {"sheet": "PL", "rows": f"{B['da']}-{B['gate']}", "cols": ["A"] + list(FC), "fmts": [""] + ["n1"] * len(FC)}}]}

def ptab(caption, keys):
    return {"type": "table", "caption": caption, "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
            "rows": [[il(k), f"@In!D{IN[k]}|{f}", iv(k, f), f"@In!F{IN[k]}|{f}", f"@In!G{IN[k]}|tag"] for k, f in keys]}
SCN = find(I, "情境驅動對照", exact=False) + 1
drivers = {"id": "drivers", "title": "驅動因子：線模型人形需求 × 份額 × 單價；本業成長 → 官方分部 → 合併營收", "blocks": [
    {"type": "theme", "title": "控制電機依用途堆疊（百萬元）與人形空心杯用占合併營收：FY2026–FY2032（情境 B）", "x": f"@Thesis!E{THDR}:K{THDR}", "fmt": "n0", "fmt_r": "p1", "stack": True, "pad": 1.3,
     "total": f"@Thesis!E{TR['cm']}:K{TR['cm']}",
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": f"@Thesis!E{TR[k]}:K{TR[k]}", "color": c} for k, c in (("ind", "#0b7a80"), ("robo", "#6fa8dc"), ("hum", "#c29a2b"))],
     "lines": [{"name": [f"@Thesis!A{TR['humsh']}"], "y": f"@Thesis!E{TR['humsh']}:K{TR['humsh']}", "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝控制電機三條用途子列（合計＝控制電機分部營收）；折線＝人形空心杯用占合併營收（右軸）。用途拆分為假設值，公司未揭露機器人與人形別營收。"},
    {"type": "note", "parts": ["@Thesis!A1"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{r}", f"@In!D{r}|p1", f"@In!E{r}|p1", f"@In!F{r}|p1"] for r in range(SCN + 1, SCN + 5)]},
    {"type": "table", "caption": "分部營收推導（情境 B；百萬元）：用途子列 → 控制電機分部 → 合併營收；核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in FC],
     "rowgen": {"sheet": "PL", "rows": f"{B['ind']}-{B['yoy']}", "cols": ["A"] + list(FC), "fmts": [""] + ["n1"] * len(FC), "row_fmt": {str(B["yoy"]): {c: "p1" for c in FC}}}},
    {"type": "table", "caption": "命題模組（情境 B）：線模型人形出貨、單價、份額", "head": [f"@Thesis!A{THDR}"] + [f"@Thesis!{c}{THDR}" for c in "EFGHIJK"],
     "rows": [[f"@Thesis!A{TR[k]}"] + [f"@Thesis!{c}{TR[k]}|{f}" for c in "EFGHIJK"] for k, f in (("ship", "n0"), ("lmk", "n0"), ("hsh", "p1"), ("px", "n0"))]},
    {"type": "h3", "text": "一、控制電機：人形空心杯（命題模組）與機器人占比"},
    ptab("有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", (("p_cm", "n0"), ("pd", "p1"), ("n_cup", "n0"), ("hu_k", "n1"), ("hum_sh26", "p0"), ("r_sh26", "p0"), ("q4g", "p0"))),
    {"type": "h3", "text": "二、其他板塊與毛利率調整"},
    ptab("有敏感度區間的參數", (("g_o", "p0"), ("gm_adj", "p1"))),
    {"type": "h3", "text": "三、全公司（損益下半段、現金流、折現）"},
    ptab("有敏感度區間的參數", (("opex26", "p1"), ("opex30", "p1"), ("fin_r", "p1"), ("tax_r", "p1"), ("cx26", "n0"), ("cx_r", "p1"),
                              ("nwc_r", "p0"), ("erp", "p1"), ("beta", "n2"), ("wacc_adj", "p1"), ("g_tv", "p1"), ("ronic_sp", "p1"), ("tv_alt", "n0"),
                              ("pe_peer", "n1"), ("rel_k", "n2"), ("w_dcf", "p0"))),
    {"type": "table", "collapse": True, "summary": "展開：固定輸入（實績錨、換算值、市場資料）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, f), f"@In!G{IN[k]}|tag"] for k, f in (("g9m26", "p1"), ("sh_cm26", "p1"), ("gm26", "p1"),
                                                                ("nci_r", "p1"), ("payout", "p0"), ("da26", "p1"), ("da_ss", "n2"), ("rf", "p2"), ("crp", "p2"), ("b_own", "n2"),
                                                                ("b_peer", "n2"), ("kd", "p2"), ("g_cap", "p1"), ("debt_v", "n2"), ("nci_v", "n2"), ("fx", "n4"))]}]}

RE = wb["對帳"]
lr_end = find(RC, "反解（整本重算", exact=False) - 2
recon = {"id": "recon", "title": "對帳：實績、券商預測、線模型市場與線層級變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}3" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"4-{lr_end}", "cols": list("ABCDE"), "fmts": ["", "n1", "n1", "n1", ""],
                "row_fmt": {str(r): ({c: "p1" for c in "BCD"} if "%" in (RC[f"B{r}"].number_format or "") else {c: "n2" for c in "BCD"} if "0.00" == RC[f"B{r}"].number_format else {})
                            for r in range(4, lr_end + 1) if ("%" in (RC[f"B{r}"].number_format or "") or RC[f"B{r}"].number_format == "0.00")}}},
    {"type": "table", "caption": "@Q!A1", "head": [f"@Q!{c}3" for c in "ABCDEF"],
     "rowgen": {"sheet": "Q", "rows": "4-18", "cols": list("ABCDEF"), "fmts": [""] + ["n1"] * 4 + [""], "row_fmt": {"7": {c: "p1" for c in "BCDE"}}}},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（工業機器人、人形、共用零組件線模型）", "head": [f"@對帳!{c}3" for c in "ABCDGI"],
     "rowgen": {"sheet": "對帳", "rows": f"4-{RE.max_row}", "cols": list("ABCDGI"), "fmts": ["", "", "", "", "tag", ""]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["build_id", "@Sens!D2"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價與折現率的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!J3"],
     "rows": [[il(k), iv(k, f), f"@In!G{IN[k]}|tag", nourl("In", f"J{IN[k]}")] for k, f in (("px", "n2"), ("rf", "p2"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A3", "@Risk!B3", "@Risk!C3", "@Risk!D3"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(4, RK.max_row + 1)]},
    {"type": "note", "parts": ["@README!A9"]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"In": ["B", "H", "J"], "Comp": ["F", "K"], "對帳": ["A", "C"]}, "empty": "沒有標為未取得的項目。"}]}

cfg = {"company": "鳴志電器", "ticker": "603728.SS", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A2",
       "question": "股價帶有人形機器人題材，但人形空心杯用營收占比小；本業控制電機的毛利率與費用率回升，以及自由現金流，能否支持現價",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：本業與機器人成長快、毛利率較高、人形份額較大）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：本業成長慢、毛利率回到 2025 年以下、人形份額小）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

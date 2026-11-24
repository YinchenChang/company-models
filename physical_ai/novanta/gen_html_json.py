# 產生 physical_ai/novanta/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/novanta/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/novanta/html.json
import glob, json, os, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
X = sys.argv[1] if len(sys.argv) > 1 else sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))[0]
wb = openpyxl.load_workbook(X, data_only=True)
P, V, I, D, SE, RC, Q = wb["PL"], wb["Val"], wb["In"], wb["DCF"], wb["Sens"], wb["Recon"], wb["Q"]


def find(ws, text, start=1, col=1, exact=True):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and (v == text if exact else v.startswith(text)):
            return r
    raise KeyError(f"{ws.title}: {text}")


IN = {I.cell(r, 1).value: r for r in range(1, I.max_row + 1) if isinstance(I.cell(r, 1).value, str)}
def inr(k): return IN[k]
def iv(k, f=""): return f"@In!E{inr(k)}" + (f"|{f}" if f else "")
def il(k): return f"@In!B{inr(k)}"
def nourl(sh, coord, f=""):
    v = wb[sh][coord].value
    return "（出處含網址，見 xlsx）" if isinstance(v, str) and "http" in v else f"@{sh}!{coord}" + (f"|{f}" if f else "")


LABS = {"c_aet": "營收：Automation Enabling Technologies（AET）", "c_med": "營收：Medical Solutions", "hum": "人形機器人六維力感測（命題子列，增量）", "rev": "營收（合併）",
        "chk1": "檢查：分部合計 − 合併營收", "gr": "營收年增率", "cogs": "營業成本", "gp": "毛利", "gm": "毛利率", "opex": "營業費用（毛利 − 營業利益；含銷管、研發、攤銷與併購整合）",
        "op": "營業利益", "opm": "營業利益率", "ebitda": "調整後 EBITDA（非 GAAP）", "ebm": "調整後 EBITDA 率", "fin": "業外：利息淨額與其他", "pbt": "稅前利益", "tax": "所得稅",
        "ni": "稅後淨利", "nim": "淨利率", "sh": "稀釋平均股數（百萬股）", "eps": "EPS（GAAP 稀釋，美元）", "epsa": "調整後 EPS（非 GAAP，美元；Comparable 分母）", "nig": "淨利年增率",
        "da": "折舊與攤銷", "dep": "其中折舊（調整後營業利益只加回攤銷）", "sbc": "股票報酬", "oth": "其他調整項（併購、整合、重整）", "cx": "資本支出", "dwc": "營運資金增加",
        "fcff": "FCFF（NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；股票報酬當成本）", "fcfad": "自由現金流扣稅後利息（收支關卡用；無股利）",
        "mkt": "參考：六維力感測市場 2030（百萬美元，線 6）", "rp": "參考：人形營收爬坡（占 2030 水準）"}
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm)
    PR[s] = dict(hdr=h + 1, **{k: find(P, t, h) for k, t in LABS.items()})
B = PR["B"]
YC = "CDEFGHIJKL"   # PL：C＝FY2023 … L＝FY2032
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VR = {k: find(V, t) for k, t in dict(dcfr="DCF 情境價值推到 12 個月後", cu="Comparable（未調整）：同業前瞻 P/E 中位數 × 情境 FY2028 調整後 EPS",
                                       ca="Comparable（相對倍數調整後，進加權）", tps="各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", tp="加權目標價（美元）",
                                       dcf="DCF 目標價（情境加權）", cmp="Comparable 目標價（情境加權，進加權）", px="現價（In 頁 px）", up="上檔空間", rt="估值評等（只看上檔空間）",
                                       gj="入不敷出判定", rec="投資建議", dist="距最近級距邊界（百分點）",
                                       bnd="翻轉目標價：升一級／降一級", a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻",
                                       pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC", pe="隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 調整後 EPS",
                                       pe1="現價 ÷ 情境 B FY2027 調整後 EPS", hum="人形六維力營收占 FY2030 合併營收（情境 B）", gaap="現價 ÷ 情境 B FY2027 GAAP EPS").items()}
VR["gate"] = find(V, "收支關卡：", exact=False)
VR["pe2"] = find(V, "Novanta 現行前瞻 P/E", exact=False)
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")

DB = find(D, "情境 B（基準）")
DK = {k: find(D, t) for k, t in dict(ke="ke（股權成本）", e="股權市值 E（US$M）", dd="負債 D（債務；營業租賃未取得）", kd="稅後 kd", wacc="WACC", g="永續成長率 g＝MIN(g_cap, rf)", ronic="RONIC",
                                       nc="淨債務（現金 − 債務；2026-10-02）").items()}
DBR = {k: find(D, t, DB) for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", eq="股權價值", psv="每股價值（評價日）",
                                          d12="DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）", tvsh="終值占 EV", alt="終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）").items()}
DA, DC = find(D, "情境 A（樂觀）"), find(D, "情境 C（保守）")
def dsc(k):
    t = D.cell(DBR[k], 1).value
    return [f"@DCF!B{find(D, t, DA)}", f"@DCF!B{DBR[k]}", f"@DCF!B{find(D, t, DC)}"]

# ---------------- 敏感度 ----------------
SH = find(SE, "驅動")
s_rows = [r for r in range(SH + 1, SE.max_row + 1) if isinstance(SE.cell(r, 5).value, (int, float))]
lo = [r for r in s_rows if str(SE.cell(r, 1).value).endswith("（低）")]
hi = [r for r in s_rows if str(SE.cell(r, 1).value).endswith("（高）")]
oth = [r for r in s_rows if r not in lo and r not in hi]
s_end = max(s_rows)
TOR = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": "E", "fmt": "sp1"}
def big(rs): return [r for r in rs if abs(SE.cell(r, 5).value) >= 0.1]
def small(rs): return [r for r in rs if abs(SE.cell(r, 5).value) < 0.1]
def main_tor(side, rs):
    if big(rs):
        return dict(TOR, title=f"加權目標價相對基準：輸入取{side}值（變動達兩位數百分比的案例）", rows=",".join(map(str, big(rs))))
    return dict(TOR, title=f"加權目標價相對基準：輸入取{side}值（沒有變動達兩位數百分比的案例；列出影響最大者）", rows=",".join(map(str, sorted(rs, key=lambda r: -abs(SE.cell(r, 5).value))[:1])))
sens_blocks = [main_tor("低", lo), main_tor("高", hi),
               {"type": "note", "parts": ["@Sens!A1"]},
               {"type": "kv", "items": [["@Sens!A2", "@Sens!B2|n2"], ["@Sens!C2", "@Sens!D2|n2"], ["@Sens!A4", "@Sens!D4"], ["@Sens!A3", "@Sens!D3"]]}]
if oth:
    sens_blocks.append(dict(TOR, title="加權目標價相對基準：只看單一情境、終值替代法", rows=",".join(map(str, oth))))
for side, rs in (("低", lo), ("高", hi)):
    shown = set(map(int, main_tor(side, rs)["rows"].split(",")))
    rest = [r for r in small(rs) if r not in shown]
    if rest:
        sens_blocks.append(dict(TOR, title=f"加權目標價相對基準：輸入取{side}值，其餘案例", rows=",".join(map(str, rest)), collapse=True,
                                summary=f"展開：輸入取{side}值、變動未達兩位數百分比的案例"))
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：全部案例（加權目標價、DCF 目標價與估值評等）",
                    "head": [f"@Sens!{c}{SH}" for c in "ABDEFGH"],
                    "rowgen": {"sheet": "Sens", "rows": f"{SH + 1}-{s_end}", "cols": ["A", "B", "D", "E", "F", "G", "H"], "fmts": ["", "", "n2", "sp1", "n2", "sp1", ""]}})
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}


hero = {"rec": vv("rec"),
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n2"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n2"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n2"), ""],
                  [il("px"), iv("px", "n2"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等 ", vv("rec"), "；收支關卡", vv("gj"), "（情境 B 自由現金流扣稅後利息合計 ", vv("gate", "C", "n0"),
                 " 百萬美元）。現價是模型下一年度調整後每股盈餘的 ", vv("pe1", "B", "x1"), "，是 GAAP 每股盈餘的 ", vv("gaap", "B", "x1"), "；只看保守情境時情境目標價為 ", vv("tps", "D", "n2"), " 美元。"]}

QS = find(Q, "AET 營收")
thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "兩個分部營收（百萬美元）：今年第三季對去年同期", "x": f"@Q!A{QS}:A{QS + 1}", "fmt": "n0", "stack": False, "pad": 1.25,
     "bars": [{"name": [f"@Q!B2"], "y": f"@Q!B{QS}:B{QS + 1}", "color": "#a9c9cb"}, {"name": [f"@Q!C2"], "y": f"@Q!C{QS}:C{QS + 1}", "color": "#0b7a80"}],
     "caption": "兩季季末分別為去年九月底與今年十月初；Medical 的增加含 Riverpoint 約兩個月；AET 為 ATI 力感測所在分部。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["承重變數：FY2027 AET 年增（情境 B ", iv("ga27_B", "p0"), "）、Medical 年增（", iv("gmed27_B", "p0"), "）、調整後 EBITDA 率（", iv("eb27_B", "p1"),
         "）、同業前瞻本益比中位數（", iv("pe_med", "n1"), " 倍）。"]]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A8", " ", "@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n2"), " 美元（DCF ", vv("dcf", "B", "n2"), "、Comparable ", vv("cmp", "B", "n2"), "），對現價 ", iv("px", "n2"), " 美元上檔 ", vv("up", "B", "sp1"),
         "，投資建議", vv("rec"), "；距評等邊界只有 ", vv("dist", "B", "n1"), " 個百分點。"],
        ["本益比：現價是情境 B FY2027 調整後每股盈餘的 ", vv("pe1", "B", "x1"), "、GAAP 每股盈餘的 ", vv("gaap", "B", "x1"), "；同業中位數 ", iv("pe_med", "n1"), " 倍，資料商口徑 Novanta 前瞻 ", iv("pe_self", "n1"), " 倍。"],
        ["財報：第三季營收 ", f"@Q!C{QS + 3}|n1", " 百萬美元、調整後每股盈餘 ", f"@Q!C{QS + 9}|n2", " 美元；FY2026 調整後 EBITDA 指引 ", iv("g26_ebitda", "n1"), " 百萬美元，毛利率因 Stockport 暫停下修。"],
        ["人形機器人：情境 B FY2030 六維力子列只占合併營收 ", vv("hum", "B", "p1"), "，結論不靠人形；Novanta 未揭露人形營收或客戶。"],
        "@README!A6"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A7"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "三個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}],
     "ref": {"label": il("px"), "value": iv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n2"), ("cmp", "n2"), ("tp", "n2"), ("px", "n2"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("dist", "n1"), ("bnd", ""))]},
    {"type": "note", "parts": ["估值邊界：", "@README!A5", "。同業倍數只納入收盤檔有同日收盤的六家；Cognex、Omron、Hexagon 報價日未取得，不進中位數。"]},
    {"type": "table", "caption": "情境目標價（美元）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", "n2"]}},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（百萬美元；每股為美元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n0"), ("eq", "n0"), ("psv", "n2"), ("d12", "n2"), ("tvsh", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n0"), ("tv", "n0"), ("pvtv", "n0"), ("alt", "n2"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n3"), ("erp", "p1"), ("crp", "p2"), ("sh_v", "n2"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n1", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n1", "", ""], [f"@DCF!A{DK['dd']}", f"@DCF!B{DK['dd']}|n1", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因與倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("pd2", "n2"), ("pe", "x1"), ("pe1", "x1"), ("gaap", "x1"), ("pe2", "n2"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻 P/E 與中位數（含報價日未取得者）", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}")] for k in ("pe_fanuc", "pe_yask", "pe_nab", "pe_smc", "pe_kyn", "pe_rrx", "pe_med", "pe_cgnx", "pe_omron", "pe_hexa", "pe_self")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"), ("dps_tgt", "n2"))]}]}

RV0 = find(RC, "反解目標")
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價（上檔為零）及落在評等邊界所需的值，其他輸入不變。「無解」表示在搜尋區間內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{find(RC, '反解（整本重算', exact=False)}", "head": [f"@Recon!{c}{RV0}" for c in "ACDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{RV0 + 1}-{RC.max_row}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n3", "n3", ""]}}]}

STD = list(YC)
STDROWS = ["rev", "gr", "cogs", "gp", "gm", "opex", "op", "opm", "ebitda", "ebm", "fin", "pbt", "tax", "ni", "nim", "sh", "eps", "epsa", "nig"]
PCT = {"gr", "gm", "opm", "ebm", "nim", "nig"}
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "合併營收（百萬美元）× 情境 A／B／C 與營業利益率（情境 B）：FY2024–FY2032", "x": f"@PL!D{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：營收"], "y": pl("rev", s, "D"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['opm']}", "（情境 B）"], "y": pl("opm", "B", "D"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境合併營收；折線＝情境 B GAAP 營業利益率（右軸）。FY2024–FY2025 實績、FY2026 公司指引三情境相同，分化從 FY2027 開始。"},
    {"type": "note", "parts": ["會計口徑：標準損益表為美國 GAAP；調整後 EBITDA 與調整後 EPS 是公司的 Non-GAAP 定義（排除攤銷、股票報酬、併購與重整成本），Comparable 的 EPS 分母用 Non-GAAP 調整後 EPS，DCF 用 GAAP 營業利益並把股票報酬當成本。FY2023 毛利以下細項未取得。"]},
    {"type": "table", "caption": "標準損益表（GAAP；百萬美元，EPS 為美元）：FY2023–FY2025 實績、FY2026 公司指引、FY2027 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [[f"@PL!A{B[k]}"] + [f"@PL!{c}{B[k]}|" + ("p1" if k in PCT else "n2" if k in ("eps", "epsa") else "n1" if k == "sh" else "n0") for c in STD] for k in STDROWS]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "合併營收（百萬美元）", "fmt": "n0", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "營業利益（百萬美元，GAAP）", "fmt": "n0", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "調整後 EBITDA 率", "fmt": "p1", "rows": {s: pl("ebm", s, "F") for s in "ABC"}},
                 {"name": "調整後 EPS（美元，Non-GAAP）", "fmt": "n2", "chart": True, "rows": {s: pl("epsa", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣稅後利息（百萬美元；收支關卡用）", "fmt": "n0", "rows": {s: pl("fcfad", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 調整項與現金流（百萬美元）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[3:]],
     "rowgen": {"sheet": "PL", "rows": f"{B['da']}-{B['fcfad']}", "cols": ["A"] + STD[3:], "fmts": [""] + ["n0"] * len(STD[3:])}}]}

drivers = {"id": "drivers", "title": "驅動因子：AET、Medical 與人形六維力 → 合併營收與利潤", "blocks": [
    {"type": "theme", "title": "營收依分部與人形子列堆疊（百萬美元）與調整後 EBITDA 率：FY2025–FY2032（情境 B）", "x": f"@PL!E{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev", "B", "E"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "E"), "color": c} for k, c in (("c_aet", "#0b7a80"), ("c_med", "#6fa8dc"), ("hum", "#b4472f"))],
     "lines": [{"name": [f"@PL!A{B['ebm']}"], "y": pl("ebm", "B", "E"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝AET、Medical Solutions（公司揭露的兩個報告分部）與人形六維力命題子列（紅色），合計等於合併營收；折線＝調整後 EBITDA 率（右軸）。"},
    {"type": "note", "parts": ["@PL!A1"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("ga27", "p1"), ("ga28", "p1"), ("gam", "p1"), ("gal", "p1"), ("gmed27", "p1"), ("gmed28", "p1"), ("gmedm", "p1"), ("gmedl", "p1"), ("m30", "n0"), ("shr", "p1"), ("gmadj", "p1"), ("eb27", "p1"), ("ebx", "p2"))]
             + [[f"@In!B{inr('scn_w')}", f"@In!D{inr('scn_w')}|p0", f"@In!E{inr('scn_w')}|p0", f"@In!F{inr('scn_w')}|p0"]]},
    {"type": "table", "caption": "營收推導（情境 B；百萬美元）：兩個分部＋人形子列 → 合併營收；核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[2:]],
     "rowgen": {"sheet": "PL", "rows": f"{B['c_aet']}-{B['gr']}", "cols": ["A"] + STD[2:], "fmts": [""] + ["n1"] * len(STD[2:]), "row_fmt": {str(B["gr"]): {c: "p1" for c in STD[2:]}}}},
    {"type": "table", "caption": "人形六維力推導（情境 B）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[4:]],
     "rows": [[f"@PL!A{B['mkt']}"] + [f"@PL!{c}{B['mkt']}|n0" for c in STD[4:]], [f"@PL!A{B['rp']}"] + [f"@PL!{c}{B['rp']}|p0" for c in STD[4:]],
              [f"@PL!A{B['hum']}"] + [f"@PL!{c}{B['hum']}|n1" for c in STD[4:]]]},
    {"type": "h3", "text": "一、AET 與 Medical 營收"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("g_anc", "p0"), ("aet26", "p0"), ("ga27_B", "p0"), ("ga28_B", "p0"), ("gam_B", "p0"), ("gal_B", "p0"), ("gmed27_B", "p0"), ("gmed28_B", "p0"), ("gmedm_B", "p0"), ("gmedl_B", "p0"))]},
    {"type": "h3", "text": "二、人形機器人六維力感測（命題子列）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("shr_B", "p1"), ("rp27", "p0"))]},
    {"type": "h3", "text": "三、利潤結構與全公司（現金流、折現）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("gm_aet", "p1"), ("gm_med", "p1"), ("gmadj_B", "p1"), ("eb27_B", "p1"), ("ebx_B", "p2"), ("da_26", "n0"), ("sbc_26", "n0"), ("oth_26", "n0"), ("fin_26", "n0"), ("sh_26", "n1"),
               ("da_pct", "p1"), ("dep_sh", "n2"), ("sbc_pct", "p1"), ("oth_27", "n0"), ("oth_lt", "n0"), ("fin_x", "n0"), ("tax_w", "p0"), ("tax_adj", "p0"), ("sh_v", "n2"), ("fc_cx", "p1"),
               ("fc_wc", "p0"), ("gap26", "p1"), ("beta", "n3"), ("erp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("w_dcf", "p0"), ("pe_med", "n2"))]},
    {"type": "table", "collapse": True, "summary": "展開：FY2025 實績與 FY2026 公司指引（固定輸入）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag"] for k in ["a25_aet", "a25_med", "g26_rev", "g26_ebitda", "g26_epsa", "g26_gma", "cx_26", "cash_v", "debt_v"]]}]}

RE = wb["對帳"]
recon = {"id": "recon", "title": "對帳：公司指引、共識與線層級變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, 'FY2027 調整後 EPS', exact=False)}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""],
                "row_fmt": {str(find(RC, "FY2026 營收", exact=False)): {c: "n0" for c in "BCD"}, str(find(RC, "FY2026 調整後 EBITDA", exact=False)): {c: "n1" for c in "BCD"},
                            str(find(RC, "FY2027 營收年增", exact=False)): {c: "p1" for c in "BCD"}, str(find(RC, "FY2027 調整後 EBITDA 年增", exact=False)): {c: "p1" for c in "BCD"}}}},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（共用零組件、人形、工業機器人）", "head": [f"@對帳!{c}2" for c in "ABCGI"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{RE.max_row}", "cols": ["A", "B", "C", "G", "I"], "fmts": ["", "", "", "tag", ""]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", "@Sens!D3"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價與倍數的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3", "@In!I3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}"), nourl("In", f"I{inr(k)}")] for k, f in (("px", "n2"), ("rf", "p2"), ("pe_self", "n2"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"Hist": ["A"], "對帳": ["A", "C"]}, "empty": "沒有標為未取得的項目。"}]}

cfg = {"company": "Novanta", "ticker": "NOVT", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "AET 與併入 Riverpoint 後的醫療部門能否支持高於現價的目標價；人形機器人六維力感測對結論有多重要",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：AET 與醫療成長加速、人形採用較快）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：需求降溫、利潤率不改善）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

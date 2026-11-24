# 產生 physical_ai/sunny/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/sunny/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/sunny/html.json
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

SEGS = [("hs", "手機產品"), ("veh", "車載產品"), ("xr", "XR 產品"), ("oth", "其他產品（機器人、智慧影像、泛物聯網）")]
LABS = dict([(f"c_{k}", f"營收：{n}") for k, n in SEGS] + [
    ("core", "四分部營收合計"), ("hum", "機器人與 Robotaxi 相機模組（命題子列；線 6 市場 × 份額）"), ("rev", "營收（合併）"), ("chk1", "檢查：分部合計 − 四分部營收合計"),
    ("chk2", "檢查：四分部＋子列 − 合併營收"), ("gr", "營收年增率"), ("cogs", "營業成本"), ("gp", "毛利"), ("gm", "毛利率"), ("sga", "研發、銷售及管理費用"),
    ("oth", "其他營業損益淨額（其他收入＋其他損益＋減損＋聯營）"), ("opex", "營業費用（費用 − 其他營業損益淨額）"), ("op", "營業利益"), ("opm", "營業利益率"),
    ("fin", "業外：財務費用"), ("pbt", "稅前利益"), ("tax", "所得稅"), ("cont", "稅後淨利（合併）"),
    ("nci", "非支配股東損益"), ("ni", "稅後淨利（歸屬母公司）"), ("nim", "淨利率"), ("sh", "平均股數（百萬股）"), ("eps", "EPS（基本，人民幣元）"),
    ("nig", "淨利年增率"), ("da", "折舊攤銷"), ("cx", "資本支出"), ("dwc", "營運資金增加"), ("finc", "其中利息與理財收益（不進 FCFF）"),
    ("fcff", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加；NOPAT 不含利息與理財收益）"), ("div", "股利（前一年淨利 × 支付率）"),
    ("fcfad", "自由現金流扣股利（收支關卡用）"), ("mkt", "參考：機器人與 Robotaxi 相機模組市場（百萬美元，線 6）"), ("rp", "參考：子列爬坡（占穩態）")])
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm)
    PR[s] = dict(hdr=h + 1, **{k: find(P, t, h) for k, t in LABS.items()})
B = PR["B"]
YC = "CDEFGHIJKL"   # PL：C＝FY2023 … L＝FY2032
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VR = {k: find(V, t) for k, t in dict(dcfr="DCF 情境價值推到 12 個月後", cu="Comparable（未調整）：同業前瞻 P/E 中位數 × 情境 FY2028 EPS（人民幣，× 港元匯率）",
                                       ca="Comparable（相對倍數調整後，進加權）", tps="各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", tp="加權目標價（港元）",
                                       dcf="DCF 目標價（情境加權）", cmp="Comparable 目標價（情境加權，進加權）", px="現價（In 頁 px）", up="上檔空間", rt="估值評等（只看上檔空間）",
                                       gate="收支關卡：FY2026–FY2030 自由現金流扣股利合計（百萬人民幣；情境 A／B／C）", gj="入不敷出判定", rec="投資建議", dist="距最近級距邊界（百分點）",
                                       bnd="翻轉目標價：升一級／降一級", a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻",
                                       pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC", pe="隱含前瞻 P/E：加權目標價 ÷（情境 B FY2028 EPS × 匯率）",
                                       pe1="現價 ÷（情境 B FY2027 EPS × 匯率）", hum="機器人與 Robotaxi 相機模組子列占 FY2030 合併營收（情境 B）").items()}
VR["cons"] = find(V, "市場共識目標價", exact=False)
VR["pe2"] = find(V, "現行前瞻 P/E（資料商 NTM", exact=False)
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")

DB = find(D, "情境 B（基準）")
DK = {k: find(D, t) for k, t in dict(ke="ke（股權成本）", e="股權市值 E（百萬人民幣）", dd="負債 D（借款債券＋租賃負債）", kd="稅後 kd", wacc="WACC", g="永續成長率 g＝MIN(g_cap, rf)", ronic="RONIC",
                                       nc="淨現金與投資（2026-06-30）＝現金 − 借款債券 − 租賃負債 − 非支配 ＋ 理財資產 × fv_k").items()}
DBR = {k: find(D, t, DB) for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", eq="股權價值", psv="每股價值（評價日，港元）",
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
               {"type": "kv", "items": [["@Sens!A2", "@Sens!B2|n0"], ["@Sens!C2", "@Sens!D2|n0"], ["@Sens!A4", "@Sens!D4"], ["@Sens!A3", "@Sens!D3"]]}]
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
                    "rowgen": {"sheet": "Sens", "rows": f"{SH + 1}-{s_end}", "cols": ["A", "B", "D", "E", "F", "G", "H"], "fmts": ["", "", "n0", "sp1", "n0", "sp1", ""]}})
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}

hero = {"rec": vv("rec"),
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n2"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n2"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n2"), ""],
                  [il("px"), iv("px", "n2"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等 ", vv("rt"), "；收支關卡", vv("gj"), "（情境 B 自由現金流扣股利合計 ", vv("gate", "C", "n0"),
                 " 百萬人民幣）。現價只有模型下一年度每股盈餘的 ", vv("pe1", "B", "x1"), "；只看保守情境時情境目標價為 ", vv("tps", "D", "n2"), " 港元，評等取決於車載與其他產品成長能否抵銷手機毛利率壓力。"]}

# 日本單體受注
QS = find(Q, "手機產品營收"); QH = 2
thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "分部營收（百萬人民幣）與合併毛利率：兩個年度與兩個上半年", "x": f"@Q!B{QH}:E{QH}", "fmt": "n0", "fmt_r": "p1", "stack": True, "pad": 1.3,
     "total": f"@Q!B{QS+4}:E{QS+4}",
     "bars": [{"name": [f"@Q!A{QS+i}"], "y": f"@Q!B{QS+i}:E{QS+i}", "color": c} for i, c in enumerate(("#0b7a80", "#6fc7cb", "#c9b98f", "#b4472f"))],
     "lines": [{"name": ["合併毛利率"], "y": f"@Q!B{QS+18}:E{QS+18}", "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝四個分部營收（公司新分部；前期為公司重編值），折線＝合併毛利率（右軸，由 Q 頁合併毛利 ÷ 合併營收）。上半年與全年期間不同，只看組成與趨勢。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["承重變數：車載產品與其他產品營收年增（情境 B：未來兩年車載 ", iv("g1veh_B", "p0"), "、其他 ", iv("g1oth_B", "p0"), "）、手機分部毛利率（", iv("gm_hs", "p1"),
         "）、同業前瞻本益比中位數（", iv("pe_med", "n1"), " 倍，家數少）。"]]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A8", " ", "@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n2"), " 港元（DCF ", vv("dcf", "B", "n2"), "、Comparable ", vv("cmp", "B", "n2"), "），對現價 ", iv("px", "n2"), " 港元上檔 ", vv("up", "B", "sp1"),
         "，投資建議", vv("rec"), "；賣方平均目標價 ", iv("cons_tp", "n2"), " 港元（只作對照）。"],
        ["分部：最近半年手機營收 ", f"@Q!E{QS}|n0", "（毛利率 ", f"@Q!E{QS+14}|p1", "）、車載 ", f"@Q!E{QS+1}|n0", "（", f"@Q!E{QS+15}|p1", "）、其他產品 ", f"@Q!E{QS+3}|n0", "（", f"@Q!E{QS+17}|p1", "）。"],
        ["資產負債表：淨現金與投資約 ", f"@DCF!B{DK['nc']}|n0", " 百萬人民幣，其中理財資產計入比例 ", iv("fv_k", "p0"), "；現價對應市值約 ", f"@DCF!B{DK['e']}|n0", " 百萬人民幣。"],
        ["實體 AI：情境 B FY2030 機器人與 Robotaxi 相機模組子列只占合併營收 ", vv("hum", "B", "p2"), "，結論不靠它；舜宇份額未取得。"],
        "@README!A6"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A7"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（港元；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}, {"label": f"@Val!A{VR['cons']}", "value": vv("cons"), "series": 1}],
     "ref": {"label": il("px"), "value": iv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n2"), ("cmp", "n2"), ("tp", "n2"), ("cons", "n2"), ("px", "n2"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("dist", "n1"), ("bnd", ""))]},
    {"type": "note", "parts": ["估值邊界：", "@README!A5", "。只看保守情境時情境目標價為 ", vv("tps", "D", "n2"), " 港元；同業倍數已以收盤檔同一天收盤換算到同一估值日，但家數少、差距大。"]},
    {"type": "table", "caption": "情境目標價（港元）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", "n2"]}},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（百萬人民幣；每股為港元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n0"), ("eq", "n0"), ("psv", "n2"), ("d12", "n2"), ("tvsh", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n0"), ("tv", "n0"), ("pvtv", "n0"), ("alt", "n2"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n3"), ("erp", "p1"), ("crp", "p2"), ("sh_v", "n2"), ("lease_v", "n0"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n0", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n0", "", ""], [f"@DCF!A{DK['dd']}", f"@DCF!B{DK['dd']}|n0", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因與倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("pd2", "n2"), ("pe", "x1"), ("pe1", "x1"), ("pe2", "n2"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻 P/E 與中位數", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}")] for k in ("pe_lar", "pe_gyc", "pe_lgi", "pe_med", "pe_self")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"), ("payout", "p0"))]}]}

RV0 = find(RC, "反解目標")
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價（上檔為零）及落在評等邊界所需的值，其他輸入不變。「無解」表示在搜尋區間內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{find(RC, '反解（整本重算', exact=False)}", "head": [f"@Recon!{c}{RV0}" for c in "ACDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{RV0 + 1}-{RC.max_row}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n3", "n3", ""]}}]}

STD = list(YC)
STDROWS = ["rev", "gr", "cogs", "gp", "gm", "sga", "oth", "opex", "op", "opm", "fin", "pbt", "tax", "cont", "nci", "ni", "nim", "sh", "eps", "nig"]
PCT = {"gr", "gm", "opm", "nim", "nig"}
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "合併營收（百萬人民幣）× 情境 A／B／C 與營業利益率（情境 B）：FY2024–FY2032", "x": f"@PL!D{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：營收"], "y": pl("rev", s, "D"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['opm']}", "（情境 B）"], "y": pl("opm", "B", "D"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境合併營收（持續營業）；折線＝情境 B 營業利益率（右軸）。FY2024–FY2025 為實績，FY2026 起三情境分化。"},
    {"type": "note", "parts": ["會計口徑：香港財務報告準則（公司法定財報準則），即 GAAP 口徑，未做 Non-GAAP 調整。FY2025 含一次性的換股收益（見 Hist 頁）。FY2023 為舊分部，分部、費用拆分與歸屬母公司淨利未取得。"]},
    {"type": "table", "caption": "標準損益表（IFRS／GAAP 口徑；百萬人民幣，EPS 為日圓）：FY2023–FY2025 實績、FY2026 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [[f"@PL!A{B[k]}"] + [f"@PL!{c}{B[k]}|" + ("p1" if k in PCT else "n2" if k in ("eps",) else "n2" if k == "sh" else "n1") for c in STD] for k in STDROWS]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "合併營收（百萬人民幣）", "fmt": "n0", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "營業利益（百萬人民幣）", "fmt": "n0", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "營業利益率", "fmt": "p1", "rows": {s: pl("opm", s, "F") for s in "ABC"}},
                 {"name": "EPS（人民幣元）", "fmt": "n2", "chart": True, "rows": {s: pl("eps", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣股利（百萬人民幣；收支關卡用）", "fmt": "n0", "rows": {s: pl("fcfad", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流與股東回饋（百萬人民幣）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[3:]],
     "rowgen": {"sheet": "PL", "rows": f"{B['da']}-{B['fcfad']}", "cols": ["A"] + STD[3:], "fmts": [""] + ["n0"] * len(STD[3:])}}]}

FAMS = [(f"{fam}{k}", fm) for fam in ("g26", "g1", "g2") for k, _ in SEGS for fm in ("p1",)] + [("m26", "n1"), ("m30", "n1"), ("m32", "n1"), ("shr", "p1"), ("gmadj", "p1")]
def prm(keys):
    return {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
            "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in keys]}
drivers = {"id": "drivers", "title": "驅動因子：四分部成長與機器人／Robotaxi 相機模組 → 營收 → 毛利", "blocks": [
    {"type": "theme", "title": "營收依分部與命題子列堆疊（百萬人民幣）與毛利率：FY2025–FY2032（情境 B）", "x": f"@PL!E{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev", "B", "E"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "E"), "color": c} for k, c in (("c_hs", "#0b7a80"), ("c_veh", "#6fc7cb"), ("c_xr", "#c9b98f"), ("c_oth", "#8a8c8c"), ("hum", "#b4472f"))],
     "lines": [{"name": [f"@PL!A{B['gm']}"], "y": pl("gm", "B", "E"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝四個分部營收加上機器人與 Robotaxi 相機模組命題子列（紅色，規模極小），合計等於合併營收；折線＝合併毛利率（右軸），由分部毛利率與組合推導。"},
    {"type": "note", "parts": ["@PL!A1"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"] for f, fm in FAMS]
             + [[f"@In!B{inr('scn_w')}", f"@In!D{inr('scn_w')}|p0", f"@In!E{inr('scn_w')}|p0", f"@In!F{inr('scn_w')}|p0"]]},
    {"type": "table", "caption": "分部營收推導（情境 B；百萬人民幣）：四分部＋子列 → 合併營收；核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[2:]],
     "rowgen": {"sheet": "PL", "rows": f"{B['c_hs']}-{B['gr']}", "cols": ["A"] + STD[2:], "fmts": [""] + ["n0"] * len(STD[2:]), "row_fmt": {str(B["gr"]): {c: "p1" for c in STD[2:]}}}},
    {"type": "table", "caption": "機器人與 Robotaxi 相機模組推導（情境 B）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[4:]],
     "rows": [[f"@PL!A{B['mkt']}"] + [f"@PL!{c}{B['mkt']}|n1" for c in STD[4:]], [f"@PL!A{B['rp']}"] + [f"@PL!{c}{B['rp']}|p0" for c in STD[4:]],
              [f"@PL!A{B['hum']}"] + [f"@PL!{c}{B['hum']}|n1" for c in STD[4:]]]},
    {"type": "h3", "text": "一、分部成長（情境 B）"},
    prm([(f"{fam}{k}_B", "p0") for fam in ("g26", "g1", "g2") for k, _ in SEGS]),
    {"type": "h3", "text": "二、機器人與 Robotaxi 相機模組（命題子列）"},
    prm([("shr_B", "p1"), ("rp27", "p0"), ("fx_usd", "n2")]),
    {"type": "h3", "text": "三、損益結構與全公司（現金流、折現）"},
    prm([("gm_hs", "p1"), ("gm_veh", "p1"), ("gm_xr", "p1"), ("gm_oth", "p1"), ("gmadj_B", "p1"), ("sga_el", "n2"), ("oth_op", "n0"), ("nonop", "n0"), ("fin_inc", "n0"), ("tax_w", "p1"), ("nci_pct", "p1"),
         ("fc_da", "p1"), ("fc_cx", "p1"), ("fc_wc", "p0"), ("payout", "p0"), ("tr_sh", "n1"), ("fv_k", "p0"), ("beta", "n3"), ("erp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("w_dcf", "p0"), ("pe_med", "n2")]),
    {"type": "table", "collapse": True, "summary": "展開：FY2025 實績與固定輸入", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag"] for k in ([f"a25_{k}" for k, _ in SEGS] + ["gd_floor", "cash_v", "fv_v", "debt_v", "lease_v", "mi_v", "sh_out", "fx_hk"])]}]}

RE = wb["對帳"]
QE = find(Q, "檢查：分部毛利合計 − 合併毛利（應為 0）")
recon = {"id": "recon", "title": "對帳：公司指引、共識、官方分部與線層級變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, '加權目標價 vs 賣方平均目標價', exact=False)}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n0", "n0", "n0", ""],
                "row_fmt": {str(find(RC, "FY2026 毛利率 vs", exact=False)): {c: "p1" for c in "BCD"}, str(find(RC, "FY2027 EPS（模型 B", exact=False)): {c: "n2" for c in "BCD"},
                            str(find(RC, "加權目標價 vs 賣方平均目標價", exact=False)): {c: "n2" for c in "BCD"}}}},
    {"type": "table", "caption": "官方報告分部（手機、車載、XR、其他；百萬人民幣）與分部對帳", "head": [f"@Q!{c}2" for c in "ABCDEF"],
     "rowgen": {"sheet": "Q", "rows": f"{QS}-{QE}", "cols": ["A", "B", "C", "D", "E", "F"], "fmts": ["", "n0", "n0", "n0", "n0", "tag"]}},
    {"type": "note", "parts": [f"@Q!A{find(Q, '註：2025 年報起', exact=False)}"]},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（相機模組與人形）", "head": [f"@對帳!{c}2" for c in "ABCGI"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{RE.max_row}", "cols": ["A", "B", "C", "G", "I"], "fmts": ["", "", "", "tag", ""]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", "@Sens!D3"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價與倍數的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3", "@In!I3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}"), nourl("In", f"I{inr(k)}")] for k, f in (("px", "n2"), ("rf", "p2"), ("pe_self", "n2"), ("cons_tp", "n2"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"In": ["B"], "Hist": ["A"], "對帳": ["A", "C"]}, "empty": "沒有標為未取得的項目。"}]}

cfg = {"company": "舜宇光學", "ticker": "2382.HK", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "車載與其他產品的成長能否抵銷手機模組的低毛利並支持高於現價的目標價；機器人與 Robotaxi 相機模組對結論有多重要",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：車載與機器人視覺成長較快）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：手機續降、車載價格競爭）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

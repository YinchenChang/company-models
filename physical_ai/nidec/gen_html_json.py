# 產生 physical_ai/nidec/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/nidec/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/nidec/html.json
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

REG = [("sp", "精密小型馬達（SPMS）"), ("au", "車載（AMEC）"), ("mo", "家電・商業・產業用 MOEN"), ("ac", "家電・商業・產業用 ACIM"), ("ma", "機器與自動化（Nidec Machinery and Automation）"), ("gc", "集團公司事業（Group Company Business）")]
LABS = dict([(f"c_{k}", f"營收：{n}") for k, n in REG] + [
    ("core", "核心營收（六分部合計）"), ("hum", "人形機器人用無框力矩馬達（命題子列，分部未分）"), ("rev", "營收（合併）"), ("chk1", "檢查：分部合計 − 核心營收"),
    ("chk2", "檢查：核心＋人形 − 合併營收"), ("gr", "營收年增率"), ("cogs", "營業成本"), ("gp", "毛利"), ("gm", "毛利率"),
    ("opex", "營業費用（含減損與其他營業損益）"), ("op", "營業利益"), ("opm", "營業利益率"),
    ("fin", "業外：稅前 − 營業利益（金融收支與其他）"), ("pbt", "稅前利益"), ("tax", "所得稅"), ("cont", "稅後利益（含非支配）"), ("disc", "停業單位損益（無）"),
    ("nci", "非支配股東損益"), ("ni", "稅後淨利（歸屬母公司）"), ("nim", "淨利率"), ("sh", "平均股數（百萬股）"), ("eps", "EPS（基本，日圓）"), ("epsc", "EPS（持續營業，日圓）"),
    ("nig", "淨利年增率"), ("da", "折舊攤銷"), ("cx", "資本支出"), ("dwc", "營運資金增加"), ("fcff", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）"),
    ("eq", "歸屬母公司權益（年底）"), ("div", "股利"), ("fcfad", "自由現金流扣股利（收支關卡用）"), ("mkt", "參考：無框力矩馬達市場（百萬美元，線 6）"), ("rp", "參考：人形營收爬坡（占穩態）")])
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm)
    PR[s] = dict(hdr=h + 1, **{k: find(P, t, h) for k, t in LABS.items()})
B = PR["B"]
YC = "CDEFGHIJKL"   # PL：C＝FY2023 … L＝FY2032
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VR = {k: find(V, t) for k, t in dict(dcfr="DCF 情境價值推到 12 個月後", cu="Comparable（未調整）：同業前瞻 P/E 中位數 × 情境目標時點起算 NTM EPS（FY2028 ×1/3＋FY2029 ×2/3）",
                                       ca="Comparable（相對倍數調整後，進加權）", tps="各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", tp="加權目標價（日圓）",
                                       dcf="DCF 目標價（情境加權）", cmp="Comparable 目標價（情境加權，進加權）", px="現價（In 頁 px）", up="上檔空間", rt="估值評等（只看上檔空間）",
                                       gate="收支關卡：FY2027–FY2031 自由現金流扣股利合計（¥B；情境 A／B／C）", gj="入不敷出判定", rec="投資建議", dist="距最近級距邊界（百分點）",
                                       bnd="翻轉目標價：升一級／降一級", a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻",
                                       pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC", pe="隱含前瞻 P/E：加權目標價 ÷ 情境 B 目標時點起算 NTM EPS",
                                       pe1="現價 ÷ 情境 B FY2027 持續營業 EPS", hum="人形無框馬達營收占 FY2030 合併營收（情境 B）").items()}
VR["cons"] = find(V, "市場共識目標價", exact=False)
VR["pe2"] = find(V, "現行前瞻 P/E（資料商", exact=False)
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")

DB = find(D, "情境 B（基準）")
DK = {k: find(D, t) for k, t in dict(ke="ke（股權成本）", e="股權市值 E（¥B）", dd="負債 D（社債及借款＋租賃代理）", kd="稅後 kd", wacc="WACC", g="永續成長率 g＝MIN(g_cap, rf)", ronic="RONIC",
                                       nc="淨現金與投資（2026-03-31）＝現金 − 借款 − 租賃代理 − 非支配 ＋ 其他非流動金融資產 × inv_k").items()}
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
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n0"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n0"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n0"), ""],
                  [il("px"), iv("px", "n0"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等 ", vv("rt"), "；收支關卡", vv("gj"), "（情境 B 自由現金流扣股利合計 ", vv("gate", "C", "n1"),
                 " 十億日圓）。現價只有模型下一年度持續營業每股盈餘的 ", vv("pe1", "B", "x1"), "；只看保守情境時情境目標價為 ", vv("tps", "D", "n0"), " 日圓，評等取決於產業機器循環能否延續。"]}

QS = find(Q, "精密小型馬達（SPMS） 營收"); QS2 = QS + 5
thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "六大事業單位外部營收（十億日圓）：FY2024、FY2025（原公告）、前九個月 FY2026", "x": f"@Q!A{QS}:A{QS2}", "fmt": "n0", "fmt_r": "n0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["@Q!B2"], "y": f"@Q!B{QS}:B{QS2}", "color": "#a9c9cb"}, {"name": ["@Q!C2"], "y": f"@Q!C{QS}:C{QS2}", "color": "#0b7a80"}, {"name": ["@Q!D2"], "y": f"@Q!D{QS}:D{QS2}", "color": "#b4472f"}],
     "lines": [],
     "caption": "官方報告分部外部顧客營收；九個月累計不可與全年直接比較。家電・商業・產業用 MOEN 為最大單一單位，資料中心需求是 FY2027 預想的主要成長來源。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["承重變數：核心營收年增（情境 B：FY2027 ", iv("g27_B", "p0"), "、FY2028 ", iv("g28_B", "p0"), "）、FY2030 起毛利率（", iv("gmlt_B", "p1"),
         "）、FY2027 營業費用率（", iv("opx27", "p1"), "）、同業前瞻本益比中位數（", iv("pe_med", "n1"), " 倍）。"]]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A8", " ", "@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n0"), " 日圓（DCF ", vv("dcf", "B", "n0"), "、Comparable ", vv("cmp", "B", "n0"), "），對現價 ", iv("px", "n0"), " 日圓上檔 ", vv("up", "B", "sp1"),
         "，投資建議", vv("rec"), "；賣方平均目標價為 ", iv("cons_tp", "n0"), " 日圓。"],
        ["盈利重置：FY2026（最近一個會計年度，三月期）營業損失 ", f"@PL!F{B['op']}|n0", " 十億日圓（含巨額資產減損）；FY2027 公司預想營業利益 ", iv("g27_op", "n0"), "（營業利益率 ", f"@PL!G{B['opm']}|p1",
         "），市場共識 ", iv("cs27_op", "n0"), "；FY2025 重編值只有 ", f"@PL!E{B['op']}|n0", "。"],
        ["資訊缺口：FY2027 中期決算短信因調查延後公布，公司預想為最近一次發布的版本，尚未更新；會計師對有價證券報告書出具無法表示意見。"],
        ["人形機器人：情境 B FY2030 人形無框馬達只占合併營收 ", vv("hum", "B", "p1"), "，結論不靠人形；Nidec 未揭露人形營收或客戶。"],
        "@README!A6"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A7"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（日圓；豎線＝現價）", "fmt": "n0",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}, {"label": f"@Val!A{VR['cons']}", "value": vv("cons"), "series": 1}],
     "ref": {"label": il("px"), "value": iv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n0"), ("cmp", "n0"), ("tp", "n0"), ("cons", "n0"), ("px", "n0"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("dist", "n1"), ("bnd", ""))]},
    {"type": "note", "parts": ["估值邊界：", "@README!A5", "。只看保守情境時情境目標價為 ", vv("tps", "D", "n0"), " 日圓；同業倍數中有多家報價日早於十一月同業回落，未換算到同一估值日。"]},
    {"type": "table", "caption": "情境目標價（日圓）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n0", "n0", "n0", "n0"]}},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（十億日圓；每股為日圓）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n1"), ("eq", "n1"), ("psv", "n0"), ("d12", "n0"), ("tvsh", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n1"), ("tv", "n1"), ("pvtv", "n1"), ("alt", "n0"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n3"), ("erp", "p1"), ("crp", "p2"), ("sh_v", "n2"), ("lease_v", "n2"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n1", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n1", "", ""], [f"@DCF!A{DK['dd']}", f"@DCF!B{DK['dd']}|n1", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因與倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("pd2", "n2"), ("pe", "x1"), ("pe1", "x1"), ("pe2", "n2"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻 P/E 與中位數", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}")] for k in ("pe_fanuc", "pe_yask", "pe_nab", "pe_thk", "pe_smc", "pe_med", "pe_self")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"), ("dps_tgt", "n0"))]}]}

RV0 = find(RC, "反解目標")
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價（上檔為零）及落在評等邊界所需的值，其他輸入不變。「無解」表示在搜尋區間內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{find(RC, '反解（整本重算', exact=False)}", "head": [f"@Recon!{c}{RV0}" for c in "ACDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{RV0 + 1}-{RC.max_row}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n3", "n3", ""]}}]}

STD = list(YC)
STDROWS = ["rev", "gr", "cogs", "gp", "gm", "opex", "op", "opm", "fin", "pbt", "tax", "cont", "disc", "nci", "ni", "nim", "sh", "eps", "epsc", "nig"]
PCT = {"gr", "gm", "opm", "nim", "nig"}
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "合併營收（十億日圓）× 情境 A／B／C 與營業利益率（情境 B）：FY2024–FY2032", "x": f"@PL!D{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：營收"], "y": pl("rev", s, "D"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['opm']}", "（情境 B）"], "y": pl("opm", "B", "D"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境合併營收；折線＝情境 B 營業利益率（右軸）。FY2024–FY2026 實績三情境相同，分化從 FY2027 開始；FY2026 營業利益率為負（減損）。"},
    {"type": "note", "parts": ["會計口徑：IFRS（公司法定財報準則），即 GAAP 口徑，未做 Non-GAAP 調整；三月期決算（FY2026 指最近一個會計年度）。FY2025 為公司後來公布的重編值，FY2023 與 FY2024 為原公告值，三年口徑不完全可比；FY2026 含巨額資產減損。"]},
    {"type": "table", "caption": "標準損益表（IFRS／GAAP 口徑；十億日圓，EPS 為日圓）：FY2023–FY2026 實績、FY2027 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [[f"@PL!A{B[k]}"] + [f"@PL!{c}{B[k]}|" + ("p1" if k in PCT else "n0" if k in ("eps", "epsc") else "n2" if k == "sh" else "n1") for c in STD] for k in STDROWS]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "合併營收（十億日圓）", "fmt": "n1", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "營業利益（十億日圓）", "fmt": "n1", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "營業利益率", "fmt": "p1", "rows": {s: pl("opm", s, "F") for s in "ABC"}},
                 {"name": "持續營業 EPS（日圓）", "fmt": "n0", "chart": True, "rows": {s: pl("epsc", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣股利（十億日圓；收支關卡用）", "fmt": "n1", "rows": {s: pl("fcfad", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流與股東回饋（十億日圓）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[3:]],
     "rowgen": {"sheet": "PL", "rows": f"{B['da']}-{B['fcfad']}", "cols": ["A"] + STD[3:], "fmts": [""] + ["n1"] * len(STD[3:])}}]}

drivers = {"id": "drivers", "title": "驅動因子：核心成長與分部傾斜、毛利率與費用率、人形無框馬達 → 合併營收", "blocks": [
    {"type": "theme", "title": "營收依事業單位與人形子列堆疊（十億日圓）與毛利率：FY2025–FY2032（情境 B）", "x": f"@PL!E{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev", "B", "E"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "E"), "color": c} for k, c in (("c_mo", "#0b7a80"), ("c_gc", "#6fc7cb"), ("c_ac", "#6fa8dc"), ("c_sp", "#8a8c8c"), ("c_au", "#c9b98f"), ("c_ma", "#d9c7a0"), ("hum", "#b4472f"))],
     "lines": [{"name": [f"@PL!A{B['gm']}"], "y": pl("gm", "B", "E"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝六個事業單位營收（最近兩個會計年度為官方外部營收按重編合併營收縮放，最近年度全年分部以前九個月占比分配）加上人形無框馬達命題子列（紅色），合計等於合併營收；折線＝毛利率（右軸）。"},
    {"type": "note", "parts": ["@PL!A1"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("g27", "p1"), ("g28", "p1"), ("gm", "p1"), ("glt", "p1"), ("m26", "n0"), ("m30", "n0"), ("m32", "n0"), ("shr", "p1"), ("gm27", "p1"), ("gmlt", "p1"), ("opxadj", "p1"))]
             + [[f"@In!B{inr('scn_w')}", f"@In!D{inr('scn_w')}|p0", f"@In!E{inr('scn_w')}|p0", f"@In!F{inr('scn_w')}|p0"]]},
    {"type": "table", "caption": "分部營收推導（情境 B；十億日圓）：六分部核心＋人形子列 → 合併營收；核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[2:]],
     "rowgen": {"sheet": "PL", "rows": f"{B['c_sp']}-{B['gr']}", "cols": ["A"] + STD[2:], "fmts": [""] + ["n1"] * len(STD[2:]), "row_fmt": {str(B["gr"]): {c: "p1" for c in STD[2:]}}}},
    {"type": "table", "caption": "人形無框馬達推導（情境 B）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[4:]],
     "rows": [[f"@PL!A{B['mkt']}"] + [f"@PL!{c}{B['mkt']}|n0" for c in STD[4:]], [f"@PL!A{B['rp']}"] + [f"@PL!{c}{B['rp']}|p0" for c in STD[4:]],
              [f"@PL!A{B['hum']}"] + [f"@PL!{c}{B['hum']}|n2" for c in STD[4:]]]},
    {"type": "h3", "text": "一、核心營收成長與分部傾斜"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("g_anc", "p0"), ("g27_B", "p0"), ("g28_B", "p0"), ("gm_B", "p1"), ("glt_B", "p0"), ("tl_mo", "p0"), ("tl_au", "p0"))]},
    {"type": "h3", "text": "二、人形機器人用無框馬達（命題子列）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("shr_B", "p1"), ("rp27", "p0"), ("fx", "n2"))]},
    {"type": "h3", "text": "三、損益結構與全公司（現金流、折現）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("gm27_B", "p1"), ("gmlt_B", "p1"), ("opx27", "p1"), ("opxadj_B", "p1"), ("opx_step", "p1"), ("nonop", "n1"), ("tax_w", "p0"), ("fc_da", "p1"), ("fc_cx", "p1"), ("fc_wc", "p0"), ("payout", "p0"),
               ("inv_k", "p0"), ("beta", "n3"), ("erp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("w_dcf", "p0"), ("pe_med", "n2"))]},
    {"type": "table", "collapse": True, "summary": "展開：分部營收與 FY2027 公司預想（固定輸入）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag"] for k in ([f"a25_{k}" for k, _ in REG] + [f"a9_{k}" for k, _ in REG] + ["g27_rev", "g27_op", "g27_pbt", "g27_ni", "cs27_rev", "cs27_op", "cs27_ni", "eq_26"])]}]}

RE = wb["對帳"]
QS = find(Q, "精密小型馬達（SPMS） 營收"); QE = find(Q, "檢查：分部合計 − 合併營收", exact=False)
recon = {"id": "recon", "title": "對帳：公司指引、共識、官方分部與線層級變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, '加權目標價 vs 賣方平均目標價')}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1", ""],
                "row_fmt": {str(find(RC, "FY2027 EPS（模型 B）", exact=False)): {c: "n0" for c in "BCD"}, str(find(RC, "加權目標價 vs 賣方平均目標價")): {c: "n0" for c in "BCD"}}}},
    {"type": "table", "caption": "官方報告分部（外部顧客營收，十億日圓）與分部對帳", "head": [f"@Q!{c}2" for c in "ABCD"],
     "rowgen": {"sheet": "Q", "rows": f"{QS}-{QE}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n3", "n3", "n3"]}},
    {"type": "note", "parts": [f"@Q!A{find(Q, '註：官方分部名稱', exact=False)}"]},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（共用零組件、人形）", "head": [f"@對帳!{c}2" for c in "ABCGI"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{RE.max_row}", "cols": ["A", "B", "C", "G", "I"], "fmts": ["", "", "", "tag", ""]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", "@Sens!D3"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價與倍數的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3", "@In!I3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}"), nourl("In", f"I{inr(k)}")] for k, f in (("px", "n0"), ("rf", "p2"), ("pe_self", "n2"), ("cons_tp", "n0"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"In": ["B"], "Hist": ["A"], "對帳": ["A", "C"]}, "empty": "沒有標為未取得的項目。"}]}

cfg = {"company": "Nidec 日本電產", "ticker": "6594.T", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "車載減損與會計調查之後，FY2027 公司預想的獲利回升能否兌現並支持高於現價的目標價；人形機器人用馬達對結論有多重要",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：資料中心需求延續、費用率改善、人形採用較快）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：成長停滯、費用率降不下來）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

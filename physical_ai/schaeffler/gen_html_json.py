# 產生 physical_ai/schaeffler/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/schaeffler/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/schaeffler/html.json
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

REG = [("pc", "動力傳動與底盤（Powertrain & Chassis）"), ("em", "電動化（E-Mobility）"), ("vl", "車輛生命週期（Vehicle Lifetime Solutions，售後）"),
       ("bi", "軸承與工業（Bearings & Industrial Solutions）"), ("ot", "其他與調節（含人形、國防等成長領域；Derived）")]
LABS = dict([(f"c_{k}", f"營收：{n}") for k, n in REG] + [
    ("core", "核心營收（五列合計）"), ("hum", "人形機器人致動器（命題子列）"), ("rev", "營收（合併）"), ("chk1", "檢查：五列合計 − 核心營收"),
    ("chk2", "檢查：核心＋人形 − 合併營收"), ("gr", "營收年增率"), ("cogs", "營業成本"), ("gp", "毛利"), ("gm", "毛利率"), ("sga", "銷售及管理費用"),
    ("oth", "其他營業損益淨額（其他收益 − 其他費用 ＋ 持分法）"), ("opex", "營業費用（銷管費用 − 其他營業損益淨額）"), ("op", "營業利益"), ("opm", "營業利益率"),
    ("fin", "業外：金融收益 − 金融費用（其他收入淨額）"), ("pbt", "稅前利益"), ("tax", "所得稅"), ("cont", "稅後淨利（含非支配）"), ("disc", "停業單位損益"),
    ("nci", "非支配股東損益"), ("ni", "稅後淨利（歸屬母公司）"), ("nim", "淨利率"), ("sh", "平均股數（百萬股）"), ("eps", "EPS（基本，歐元）"), ("epsc", "EPS（持續營業，歐元）"),
    ("nig", "淨利年增率"), ("da", "折舊攤銷"), ("cx", "資本支出"), ("dwc", "營運資金增加＋重組現金支出"), ("fcff", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）"),
    ("eq", "歸屬母公司權益（年底）"), ("div", "股利"), ("fcfad", "自由現金流扣股利（收支關卡用）"), ("mkt", "參考：行星滾柱絲杠市場（百萬美元，線 6）"), ("rp", "參考：人形營收爬坡（占穩態）")])
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm)
    PR[s] = dict(hdr=h + 1, **{k: find(P, t, h) for k, t in LABS.items()})
B = PR["B"]
YC = "CDEFGHIJKL"   # PL：C＝FY2023 … L＝FY2032
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VR = {k: find(V, t) for k, t in dict(dcfr="DCF 情境價值推到 12 個月後", cu="Comparable（未調整）：同業前瞻 P/E 中位數 × 情境 FY2028 EPS（持續營業）",
                                       ca="Comparable（相對倍數調整後，進加權）", tps="各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", tp="加權目標價（歐元）",
                                       dcf="DCF 目標價（情境加權）", cmp="Comparable 目標價（情境加權，進加權）", px="現價（In 頁 px）", up="上檔空間", rt="估值評等（只看上檔空間）",
                                       gate="收支關卡：FY2026–FY2030 自由現金流扣股利合計（€M；情境 A／B／C）", gj="入不敷出判定", rec="投資建議", dist="距最近級距邊界（百分點）",
                                       bnd="翻轉目標價：升一級／降一級", a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻",
                                       pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC", pe="隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 持續營業 EPS",
                                       pe1="現價 ÷ 情境 B FY2027 持續營業 EPS", hum="人形線性關節營收占 FY2030 合併營收（情境 B）").items()}
VR["cons"] = find(V, "市場共識目標價", exact=False)
VR["pe2"] = find(V, "現行前瞻 P/E（資料商 NTM", exact=False)
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")

DB = find(D, "情境 B（基準）")
DK = {k: find(D, t) for k, t in dict(ke="ke（股權成本）", e="股權市值 E（€M）", dd="負債 D（社債及借款＋租賃代理）", kd="稅後 kd", wacc="WACC", g="永續成長率 g＝MIN(g_cap, rf)", ronic="RONIC",
                                       nc="淨現金與投資（2026-09-30）＝現金 − 借款 − 租賃代理 − 非支配 − 退休金 ＋ 其他非流動金融資產 × inv_k").items()}
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
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n2"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n2"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n2"), ""],
                  [il("px"), iv("px", "n2"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等 ", vv("rt"), "；收支關卡", vv("gj"), "（情境 B 自由現金流扣股利合計 ", vv("gate", "C", "n1"),
                 " 百萬歐元）。現價只有模型下一年度持續營業每股盈餘的 ", vv("pe1", "B", "x1"), "；只看保守情境時情境目標價為 ", vv("tps", "D", "n2"), " 歐元，評等取決於重組節省與電動化虧損收斂能否兌現。"]}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "事業部門營收（百萬歐元）與 EBIT 利益率（特殊項目前）：四個事業部門", "x": "@Q!A3:A6", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["@Q!B2"], "y": "@Q!B3:B6", "color": "#0b7a80"}, {"name": ["@Q!E2"], "y": "@Q!E3:E6", "color": "#a9c9cb"}],
     "lines": [{"name": ["@Q!C2"], "y": "@Q!C3:C6", "axis": "r", "color": "#b4472f", "labels": True}],
     "caption": "長條＝各事業部門去年實績與今年指引上緣營收；折線＝去年特殊項目前 EBIT 利益率（右軸）。電動化部門持續虧損，是重組與減損的來源；軸承與工業、售後的利益率較高，是本模型利益率回升的主要依靠。"},
    {"type": "table", "caption": "事業部門：FY2025 營收與 EBIT 利益率、FY2026 公司指引（百萬歐元；利益率為特殊項目前）", "head": [f"@Q!{c}2" for c in "ABCDEFG"],
     "rowgen": {"sheet": "Q", "rows": "3-9", "cols": ["A", "B", "C", "D", "E", "F", "G"], "fmts": ["", "n0", "p1", "n0", "n0", "p1", "p1"]}},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["承重變數：核心營收年增（情境 B：FY2027 ", iv("g27_B", "p1"), "、FY2028 ", iv("g28_B", "p1"), "）、邊際毛利率（", iv("mm_gp", "p0"),
         "）、重組節省實現度（", iv("sav_k", "p0"), "）、同業前瞻本益比中位數（", iv("pe_med", "n1"), " 倍）。"]]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A8", " ", "@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n2"), " 歐元（DCF ", vv("dcf", "B", "n2"), "、Comparable ", vv("cmp", "B", "n2"), "），對現價 ", iv("px", "n2"), " 歐元上檔 ", vv("up", "B", "sp1"),
         "，投資建議", vv("rec"), "；賣方平均目標價為 ", iv("cons_tp", "n2"), " 歐元。"],
        ["盈利重置：FY2026 公司指引營收 ", f"@PL!F{B['rev']}|n0", " 百萬歐元、營業利益（特殊項目後）", f"@PL!F{B['op']}|n0", "（利益率 ", f"@PL!F{B['opm']}|p1",
         "），FY2025 只有 ", f"@PL!E{B['op']}|n0", "；情境 B FY2027 營業利益 ", f"@PL!G{B['op']}|n0", "。"],
        ["前九個月：營收 ", "@Q!B13|n0", "（百萬歐元），第三季 EBIT 利益率（特殊項目前）", "@Q!B16|p1", "，自由現金流（M&A 前）", "@Q!B18|n0", "，淨財務負債 ", "@Q!B19|n0", "。"],
        ["人形機器人：情境 B FY2030 人形致動器只占合併營收 ", vv("hum", "B", "p1"), "，結論不靠人形；公司宣布與四家人形機器人廠合作，但未揭露營收或客戶金額。"],
        "@README!A6"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A7"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（歐元；豎線＝現價）", "fmt": "n0",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}, {"label": f"@Val!A{VR['cons']}", "value": vv("cons"), "series": 1}],
     "ref": {"label": il("px"), "value": iv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n2"), ("cmp", "n2"), ("tp", "n2"), ("cons", "n2"), ("px", "n2"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("dist", "n1"), ("bnd", ""))]},
    {"type": "note", "parts": ["估值邊界：", "@README!A5", "。只看保守情境時情境目標價為 ", vv("tps", "D", "n2"), " 歐元；同業倍數只有五家亞洲同業，歐洲同業在股價基準日的報價未取得、未納入；退休金金額未找到，以區間代入。"]},
    {"type": "table", "caption": "情境目標價（歐元）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", "n2"]}},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（百萬歐元；每股為歐元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n1"), ("eq", "n1"), ("psv", "n2"), ("d12", "n2"), ("tvsh", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n1"), ("tv", "n1"), ("pvtv", "n1"), ("alt", "n2"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n3"), ("erp", "p1"), ("crp", "p2"), ("sh_v", "n2"), ("lease_v", "n0"), ("pens_v", "n0"), ("mi_v", "n0"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n1", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n1", "", ""], [f"@DCF!A{DK['dd']}", f"@DCF!B{DK['dd']}|n1", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因與倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("pd2", "n2"), ("pe", "x1"), ("pe1", "x1"), ("pe2", "n2"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻 P/E 與中位數", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}")] for k in ("pe_mobis", "pe_denso", "pe_nsk", "pe_tuopu", "pe_nxt", "pe_med", "pe_self")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"), ("dps_tgt", "n2"))]}]}

RV0 = find(RC, "反解目標")
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價（上檔為零）及落在評等邊界所需的值，其他輸入不變。「無解」表示在搜尋區間內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{find(RC, '反解（整本重算', exact=False)}", "head": [f"@Recon!{c}{RV0}" for c in "ACDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{RV0 + 1}-{RC.max_row}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n3", "n3", ""]}}]}

STD = list(YC)
STDROWS = ["rev", "gr", "cogs", "gp", "gm", "sga", "oth", "opex", "op", "opm", "fin", "pbt", "tax", "cont", "disc", "nci", "ni", "nim", "sh", "eps", "epsc", "nig"]
PCT = {"gr", "gm", "opm", "nim", "nig"}
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "合併營收（百萬歐元）× 情境 A／B／C 與營業利益率（情境 B）：FY2024–FY2032", "x": f"@PL!D{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：營收"], "y": pl("rev", s, "D"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['opm']}", "（情境 B）"], "y": pl("opm", "B", "D"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境合併營收（持續營業）；折線＝情境 B 營業利益率（右軸）。FY2024–FY2025 實績、FY2026 公司指引三情境相同，分化從 FY2027 開始。"},
    {"type": "note", "parts": ["會計口徑：IFRS（公司法定財報準則），即 GAAP 口徑，營業利益為特殊項目（重組、減損）後；FY2024 只含 Vitesco 自二〇二四年十月起一季，與 FY2025 不可比；FY2023 與 FY2024 無事業部門拆分、特殊項目未拆出。"]},
    {"type": "table", "caption": "標準損益表（IFRS／GAAP 口徑；百萬歐元，EPS 為歐元）：FY2023–FY2025 實績、FY2026 公司指引、FY2027 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [[f"@PL!A{B[k]}"] + [f"@PL!{c}{B[k]}|" + ("p1" if k in PCT else "n2" if k in ("eps", "epsc") else "n1" if k == "sh" else "n0") for c in STD] for k in STDROWS]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "合併營收（百萬歐元）", "fmt": "n0", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "營業利益（百萬歐元）", "fmt": "n0", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "營業利益率", "fmt": "p1", "rows": {s: pl("opm", s, "F") for s in "ABC"}},
                 {"name": "持續營業 EPS（歐元）", "fmt": "n2", "chart": True, "rows": {s: pl("epsc", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣股利（百萬歐元；收支關卡用）", "fmt": "n0", "rows": {s: pl("fcfad", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流與股東回饋（百萬歐元）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[3:]],
     "rowgen": {"sheet": "PL", "rows": f"{B['da']}-{B['fcfad']}", "cols": ["A"] + STD[3:], "fmts": [""] + ["n0"] * len(STD[3:])}}]}

drivers = {"id": "drivers", "title": "驅動因子：汽車與工業循環＋人形致動器 → 事業部門營收 → 合併營收", "blocks": [
    {"type": "theme", "title": "營收依事業部門與人形子列堆疊（百萬歐元）與毛利率：FY2025–FY2032（情境 B）", "x": f"@PL!E{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev", "B", "E"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "E"), "color": c} for k, c in (("c_pc", "#0b7a80"), ("c_bi", "#6fc7cb"), ("c_vl", "#6fa8dc"), ("c_em", "#8a8c8c"), ("c_ot", "#c9b98f"), ("hum", "#b4472f"))],
     "lines": [{"name": [f"@PL!A{B['gm']}"], "y": pl("gm", "B", "E"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝四個事業部門營收（公司揭露）、其他與調節（倒推）與人形致動器命題子列（紅色），合計等於合併營收；折線＝毛利率（右軸）。"},
    {"type": "note", "parts": ["@PL!A1"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("g27", "p1"), ("g28", "p1"), ("gm", "p1"), ("glt", "p1"), ("m26", "n0"), ("m30", "n0"), ("m32", "n0"), ("shr", "p1"), ("gmadj", "p1"))]
             + [[f"@In!B{inr('scn_w')}", f"@In!D{inr('scn_w')}|p0", f"@In!E{inr('scn_w')}|p0", f"@In!F{inr('scn_w')}|p0"]]},
    {"type": "table", "caption": "事業部門營收推導（情境 B；百萬歐元）：五列核心＋人形子列 → 合併營收；核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[2:]],
     "rowgen": {"sheet": "PL", "rows": f"{B['c_pc']}-{B['gr']}", "cols": ["A"] + STD[2:], "fmts": [""] + ["n0"] * len(STD[2:]), "row_fmt": {str(B["gr"]): {c: "p1" for c in STD[2:]}}}},
    {"type": "table", "caption": "人形線性關節推導（情境 B）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[4:]],
     "rows": [[f"@PL!A{B['mkt']}"] + [f"@PL!{c}{B['mkt']}|n0" for c in STD[4:]], [f"@PL!A{B['rp']}"] + [f"@PL!{c}{B['rp']}|p0" for c in STD[4:]],
              [f"@PL!A{B['hum']}"] + [f"@PL!{c}{B['hum']}|n2" for c in STD[4:]]]},
    {"type": "h3", "text": "一、汽車與工業循環（核心營收）"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("g_anc", "p0"), ("g27_B", "p1"), ("g28_B", "p1"), ("gm_B", "p1"), ("glt_B", "p1"))]},
    {"type": "h3", "text": "二、人形機器人致動器（命題子列）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("shr_B", "p1"), ("rp27", "p0"), ("fx", "n3"))]},
    {"type": "h3", "text": "三、損益結構與全公司（現金流、折現）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("ebitm_26", "p1"), ("gmr_26", "p1"), ("spi_26", "n0"), ("mm_gp", "p0"), ("gmadj_B", "p1"), ("sga_el", "n2"), ("sav", "n0"), ("sav_k", "p0"), ("oth_op", "n0"), ("nonop", "n0"), ("tax_w", "p0"), ("fc_da", "p1"), ("fc_cx", "p1"), ("fc_wc", "p0"), ("rcx_27", "n0"), ("rcx_28", "n0"),
               ("lease_v", "n0"), ("pens_v", "n0"), ("mi_v", "n0"), ("beta", "n3"), ("erp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("w_dcf", "p0"), ("pe_med", "n2"))]},
    {"type": "table", "collapse": True, "summary": "展開：FY2025 實績與 FY2026 公司指引（固定輸入）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag"] for k in ([f"a25_{k}" for k, _ in REG] + [f"g26_{k}" for k, _ in REG] +
                                                                  ["g26_rev", "g26_ot", "ebitm_26", "gmr_26", "spi_26", "fin_26", "tax_26", "nci_c", "da_26", "dps_26", "eq_26", "doe"])]}]}

RE = wb["對帳"]
recon = {"id": "recon", "title": "對帳：公司指引、共識、事業部門與線層級變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, '加權目標價 vs 賣方平均目標價')}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n0", "n0", "n0", ""],
                "row_fmt": {str(find(RC, "FY2027 EPS", exact=False)): {c: "n2" for c in "BCD"}, str(find(RC, "加權目標價 vs 賣方平均目標價")): {c: "n2" for c in "BCD"}}}},
    {"type": "table", "caption": "事業部門對帳（百萬歐元）：四部門合計＋其他與調節＝合併營收", "head": [f"@Q!{c}2" for c in "ABCDEFGH"],
     "rowgen": {"sheet": "Q", "rows": "3-10", "cols": ["A", "B", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n0", "p1", "n0", "n0", "p1", "p1", "tag"]}},
    {"type": "note", "parts": ["@Q!A1"]},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（共用零組件、人形）", "head": [f"@對帳!{c}2" for c in "ABCGI"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{RE.max_row}", "cols": ["A", "B", "C", "G", "I"], "fmts": ["", "", "", "tag", ""]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", "@Sens!D3"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價與倍數的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3", "@In!I3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}"), nourl("In", f"I{inr(k)}")] for k, f in (("px", "n2"), ("rf", "p2"), ("pe_self", "n2"), ("cons_tp", "n2"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"Hist": ["A"], "對帳": ["A", "C"]}, "empty": "沒有標為未取得的項目。"}]}

cfg = {"company": "Schaeffler 舍弗勒", "ticker": "SHA0.DE", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "重組節省與電動化虧損收斂能否支持高於現價的目標價；人形機器人致動器對結論有多重要",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：歐洲車廠量回升、人形採用較快）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：車廠量減、重組節省落空）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

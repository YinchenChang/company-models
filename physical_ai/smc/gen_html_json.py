# 產生 physical_ai/smc/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；模型改版移列後重跑本檔即可。設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/smc/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/smc/html.json
import glob, json, os, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
X = sys.argv[1] if len(sys.argv) > 1 else sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))[0]
wb = openpyxl.load_workbook(X, data_only=True)
P, V, I, D, SE, RC = wb["PL"], wb["Val"], wb["In"], wb["DCF"], wb["Sens"], wb["Recon"]


def find(ws, text, start=1, col=1, exact=True):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and (v == text if exact else v.startswith(text)):
            return r
    raise KeyError(f"{ws.title}: {text}")


IN = {I.cell(r, 1).value: r for r in range(1, I.max_row + 1) if isinstance(I.cell(r, 1).value, str)}
def inr(k): return IN[k]

LABS = dict(vol="銷售量年增（命題驅動）", prc="價格年變化", jp="地區子列：日本", cn="地區子列：中國", as_="地區子列：亞洲（中國以外）", eu="地區子列：歐洲", us="地區子列：美國",
            ot="地區子列：其他地區", rev="營收（合併；自動制御機器事業，單一報告分部）", chk="檢查：地區子列合計 − 合併營收", yoy="營收年增率", cogs="營業成本", gp="毛利", gm="毛利率",
            sga="營業費用（販管費）", sgar="營業費用率", op="營業利益", om="營業利益率", nonop="業外損益（營業外＋特別損益）", pti="稅前利益", tax="所得稅", nit="稅後淨利（含非支配）",
            nci="非支配股東損益", ni="淨利（歸屬母公司）", nim="淨利率", nig="淨利年增率", sh="平均股數（百萬股）", eps="EPS（日圓）", da="折舊攤銷", cx="資本支出", wc="營運資金增加",
            fcff="FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", div="股利", fcfad="自由現金流扣股利（收支關卡用）")
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm)
    PR[s] = dict(hdr=h + 1, **{k: find(P, t, h) for k, t in LABS.items()})
B = PR["B"]
STD = list("CDEFGHIJKL")   # PL：C＝FY2023 … L＝FY2032
FC = list("FGHIJKL")       # FY2026 起
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VR = {k: find(V, t) for k, t in dict(ntm="NTM EPS（目標時點 2027-11 起算：FY2027 × 1/3＋FY2028 × 2/3；日圓）", dcfr="DCF 情境價值推到目標時點",
                                       cu="Comparable（未調整）：同業前瞻 P/E 中位數 × NTM EPS", ca="Comparable（相對倍數調整後，進加權）",
                                       tps="各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", tp="加權目標價（日圓）", dcf="DCF 目標價（情境加權）",
                                       cmp="Comparable 目標價（情境加權，進加權）", px="現價（In 頁 px）", up="上檔空間", rt="估值評等（只看上檔空間）",
                                       gate="收支關卡：FY2026–FY2030 自由現金流扣股利合計（¥B；情境 A／B／C）", gj="入不敷出判定", rec="投資建議", dist="距最近級距邊界（百分點）",
                                       bnd="翻轉目標價：升一級／降一級", a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻",
                                       cmpup="Comparable 目標價上檔空間（只作對照，評等不依此）", pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC",
                                       pe="隱含前瞻 P/E：加權目標價 ÷ 情境 B NTM EPS", pe2="現行前瞻 P/E（現價 ÷ FY2026 指引 EPS）÷ 同業中位數", tvsh="DCF 終值占 EV（情境 B）").items()}
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")
def iv(k, f=""): return f"@In!E{inr(k)}" + (f"|{f}" if f else "")
def il(k): return f"@In!B{inr(k)}"
def nourl(sh, coord, f=""):
    v = wb[sh][coord].value
    return "（出處含網址，見 xlsx）" if isinstance(v, str) and "http" in v else f"@{sh}!{coord}" + (f"|{f}" if f else "")
def prm(keys):
    return [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in keys]
PH = ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"]

DB = find(D, "情境 B（基準）")
DK = {k: find(D, t) for k, t in dict(ke="ke（股權成本）", e="股權市值 E（¥B）", dd="負債 D（借款）", kd="稅後 kd", wacc="WACC", g="永續成長率 g＝MIN(g_cap, rf)", ronic="RONIC",
                                       nc="淨現金與投資（資產負債表 2026-09-30）＝現金＋有價證券 − 借款 − 非支配 ＋ 投資有價證券 × inv_k").items()}
DBR = {k: find(D, t, DB) for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", nc2="＋淨現金與投資（2026-09-30）",
                                          stub="＋下半年 FCFF（2026-10～2027-03：指引全年 − 上半實績）", bb="− 尚待買回金額", dv="− 評價日前支付的中間股利與 FY2026 期末股利",
                                          eq="股權價值（評價日）", psv="每股價值（評價日）", d12="DCF 情境價值推到目標時點（×(1＋ke)^t − 期間股利）", tvsh="終值占 EV",
                                          alt="終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到目標時點）").items()}
DA, DC = find(D, "情境 A（樂觀）"), find(D, "情境 C（保守）")
def dsc(k):
    t = D.cell(DBR[k], 1).value
    return [f"@DCF!B{find(D, t, DA)}", f"@DCF!B{DBR[k]}", f"@DCF!B{find(D, t, DC)}"]

# Sens：表頭第 9 列；第 10 列起為案例
s_rows = [r for r in range(10, SE.max_row + 1) if isinstance(SE.cell(r, 4).value, (int, float)) and isinstance(SE.cell(r, 5).value, (int, float))]
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
sens_blocks = [main_tor("低", lo), main_tor("高", hi), {"type": "note", "parts": ["@Sens!A1"]},
               {"type": "kv", "items": [["@Sens!A6", "@Sens!D6|n0"], ["模型狀態", "@Sens!D7"], ["build_id", "@Sens!D3"]]}]
if oth:
    sens_blocks.append(dict(TOR, title="加權目標價相對基準：只看單一情境、終值替代法", rows=",".join(map(str, oth))))
for side, rs in (("低", lo), ("高", hi)):
    shown = set(map(int, main_tor(side, rs)["rows"].split(",")))
    rest = [r for r in small(rs) if r not in shown]
    if rest:
        sens_blocks.append(dict(TOR, title=f"加權目標價相對基準：輸入取{side}值，其餘案例", rows=",".join(map(str, rest)), collapse=True,
                                summary=f"展開：輸入取{side}值、變動未達兩位數百分比的案例"))
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：全部案例（加權目標價、DCF 目標價與估值評等）",
                    "head": ["@Sens!A9", "@Sens!B9", "@Sens!D9", "@Sens!E9", "@Sens!F9", "@Sens!G9", "@Sens!H9"],
                    "rowgen": {"sheet": "Sens", "rows": f"10-{s_end}", "cols": ["A", "B", "D", "E", "F", "G", "H"], "fmts": ["", "", "n0", "sp1", "n0", "sp1", ""]}})
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}

PAI0 = find(RC, "機器人拉動的氣動市場 FY2026", exact=False)
PAIX = {k: PAI0 + i for i, k in enumerate(("m26", "m30", "sh", "r26", "r30", "p26", "p30", "ir", "cagr"))}

hero = {"rec": vv("rec"),
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n0"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n0"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n0"), ["上檔空間 ", vv("cmpup", "B", "sp1")]],
                  [il("px"), iv("px", "n0"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等 ", vv("rt"), "（只看上檔空間）；收支關卡", vv("gate", "C", "n0"), " 十億日圓，", vv("gj"),
                 "。目標價建立在半導體帶動的設備投資高峰延續、毛利率小幅回落的假設上；機器人相關營收約占 ", f"@Recon!C{PAIX['p26']}|p1", "，不主導目標價。"]}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "合併營收依地區堆疊（十億日圓）與營業利益率：FY2023–FY2032（情境 B）", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k), "color": c} for k, c in (("cn", "#0b7a80"), ("jp", "#6fa8dc"), ("as_", "#6fc7cb"), ("eu", "#c29a2b"), ("us", "#8a8c8c"), ("ot", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['om']}"], "y": pl("om"), "axis": "r", "color": "#3b3838", "labels": True},
               {"name": [f"@PL!A{B['gm']}"], "y": pl("gm"), "axis": "r", "color": "#b4472f", "dash": True}],
     "caption": "堆疊＝單一報告分部之下的地區子列（合計等於合併營收）；FY2023–FY2025 實績、FY2026 公司指引、FY2027 起模型（地區結構固定為 FY2025）。實線＝營業利益率、虛線＝毛利率（右軸）。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["承重變數：FY2027 銷售量年增（情境 B ", iv("vg27_B", "p1"), "）、FY2028–FY2029 每年 ", iv("vgm_B", "p1"), "、毛利率相對指引調整（", iv("gadj_B", "p1"),
         "）、同業前瞻本益比中位數（", iv("pe_med", "n1"), " 倍）。"]]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n0"), " 日圓（DCF ", vv("dcf", "B", "n0"), "、Comparable ", vv("cmp", "B", "n0"), "），對現價 ", iv("px", "n0"), " 日圓上檔 ", vv("up", "B", "sp1"),
         "，估值評等", vv("rt"), "、投資建議", vv("rec"), "。"],
        ["本期：公司把 FY2026 營收指引上修到 ", iv("g26_rev", "n0"), " 十億日圓、營業利益 ", iv("g26_op", "n0"), "，毛利率指引 ", iv("g26_gm", "p1"), "；現價對指引每股盈餘的本益比 ",
         iv("pe_self", "n1"), " 倍，低於同業中位數 ", iv("pe_med", "n1"), " 倍。"],
        ["需求：受注由半導體與資料中心帶動，下半年指引營收低於上半，為模型基準；FY2027 量的情境 A／B／C 為 ", iv("vg27_A", "p0"), "／", iv("vg27_B", "p0"), "／", iv("vg27_C", "p0"), "。"],
        ["評等邊界：", "@README!A5"],
        "@README!A6"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A8"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（日圓；豎線＝現價）", "fmt": "n0",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}, {"label": il("cons_tp"), "value": iv("cons_tp"), "series": 1}],
     "ref": {"label": il("px"), "value": iv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n0"), ("cmp", "n0"), ("tp", "n0"), ("px", "n0"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("dist", "n1"), ("bnd", ""))]
             + [[il("cons_tp"), iv("cons_tp", "n0")]]},
    {"type": "table", "caption": "情境目標價（日圓；NTM EPS 為日圓）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['ntm']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n0", "n0", "n0", "n0"]}},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（十億日圓；每股為日圓）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n1"), ("eq", "n1"), ("psv", "n0"), ("d12", "n0"), ("tvsh", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n1"), ("tv", "n1"), ("pvtv", "n1"), ("nc2", "n1"), ("stub", "n1"), ("bb", "n1"), ("dv", "n1"), ("alt", "n0"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n3"), ("erp", "p1"), ("crp", "p2"), ("sh_v", "n2"), ("bb_rem", "n1"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n1", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n1", "", ""], [f"@DCF!A{DK['dd']}", f"@DCF!B{DK['dd']}|n1", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因與隱含倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("cmpup", "sp1"), ("pd2", "n2"), ("pe", "x1"), ("pe2", "n2"), ("tvsh", "p1"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻本益比與中位數", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!I3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag", nourl("In", f"I{inr(k)}")] for k in ("pe_keyence", "pe_fanuc", "pe_nab", "pe_okuma", "pe_dmg", "pe_yask", "pe_thk", "pe_nidec", "pe_med", "pe_self")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": PH,
     "rows": prm((("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"), ("dps_tgt", "n0"), ("dps_int", "n0"), ("dps_ye", "n0")))}]}

RVH = find(RC, "反解目標")
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價（上檔為零）及落在評等邊界所需的值，其他輸入不變。「無解」表示在搜尋區間內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{RVH - 1}", "head": [f"@Recon!{c}{RVH}" for c in "ACDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{RVH + 1}-{RC.max_row}", "cols": ["A", "C", "D", "E"], "fmts": ["", "p2", "p2", ""]}}]}

HR = wb["Hist"]
pl_rows = [("rev", "n1"), ("yoy", "p1"), ("cogs", "n1"), ("gp", "n1"), ("gm", "p1"), ("sga", "n1"), ("sgar", "p1"), ("op", "n1"), ("om", "p1"), ("nonop", "n1"), ("pti", "n1"),
           ("tax", "n1"), ("nit", "n1"), ("nci", "n2"), ("ni", "n1"), ("nim", "p1"), ("nig", "p1"), ("sh", "n2"), ("eps", "n0")]
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "合併營收（十億日圓）× 情境 A／B／C 與營業利益率（情境 B）：FY2023–FY2032", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：合併營收"], "y": pl("rev", s), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['om']}", "（情境 B）"], "y": pl("om"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境合併營收；折線＝情境 B 營業利益率（右軸）。FY2023–FY2025 實績、FY2026 公司指引三情境相同，分化從 FY2027 開始。"},
    {"type": "note", "parts": ["會計口徑：日本會計準則（J-GAAP，公司法定財報準則；本表即 GAAP 口徑，未做 Non-GAAP 調整，因公司與市場皆以法定財報比較）。年度＝起始年（三月底決算）：FY2025 為最近實績年度、FY2026 為本期公司指引。營業費用＝毛利 − 營業利益（含販管費與其他營業費用）；業外含營業外損益與特別損益。"]},
    {"type": "table", "caption": "標準損益表（日本會計準則／GAAP；十億日圓，EPS 為日圓）：FY2023–FY2025 實績、FY2026 公司指引、FY2027 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [[f"@PL!A{B[k]}"] + [f"@PL!{c}{B[k]}|{f}" for c in STD] for k, f in pl_rows]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "合併營收（十億日圓）", "fmt": "n1", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "營業利益（十億日圓）", "fmt": "n1", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "營業利益率", "fmt": "p1", "rows": {s: pl("om", s, "F") for s in "ABC"}},
                 {"name": "EPS（日圓）", "fmt": "n0", "chart": True, "rows": {s: pl("eps", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣股利（十億日圓；收支關卡用）", "fmt": "n1", "rows": {s: pl("fcfad", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流（十億日圓）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['da']}-{B['fcfad']}", "cols": ["A"] + STD, "fmts": [""] + ["n1"] * len(STD)}}]}

drivers = {"id": "drivers", "title": "驅動因子：銷售量 × 價格 → 單一分部營收（地區子列）→ 利潤槓桿", "blocks": [
    {"type": "theme", "title": "銷售量年增 × 情境 A／B／C 與毛利率（情境 B）：FY2027–FY2032", "x": f"@PL!G{B['hdr']}:L{B['hdr']}", "fmt": "p0", "fmt_r": "p0", "stack": False, "pad": 1.4,
     "bars": [{"name": ["情境 ", s, "：銷售量年增"], "y": pl("vol", s, "G"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['gm']}", "（情境 B）"], "y": pl("gm", "B", "G"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "長條＝三情境銷售量年增（價格另計）；折線＝情境 B 毛利率（右軸）。情境 C 在 FY2027 設定設備投資循環回落。"},
    {"type": "note", "parts": [f"@PL!A{find(P, 'PL — ', exact=False)}"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("vg27", "p1"), ("vgm", "p1"), ("vglt", "p1"), ("prc", "p1"), ("gadj", "p1"))]},
    {"type": "table", "caption": "分部營收推導（情境 B；十億日圓）：銷售量 × 價格 → 合併營收（單一分部）→ 地區子列；核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['vol']}-{B['yoy']}", "cols": ["A"] + STD, "fmts": [""] + ["n1"] * len(STD),
                "row_fmt": {str(B["vol"]): {c: "p1" for c in STD}, str(B["prc"]): {c: "p1" for c in STD}, str(B["yoy"]): {c: "p1" for c in STD}}}},
    {"type": "h3", "text": "一、需求與價格（命題驅動）"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": PH,
     "rows": prm((("g_anc", "p0"), ("vg27_B", "p1"), ("vgm_B", "p1"), ("vglt_B", "p1"), ("prc_B", "p1")))},
    {"type": "h3", "text": "二、利潤槓桿（毛利率、營業費用）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": PH, "rows": prm((("gadj_B", "p1"), ("e_sga", "n2"), ("nonop", "n1"), ("tax_w", "p0")))},
    {"type": "h3", "text": "三、現金流與折現"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": PH,
     "rows": prm((("fc_da", "p1"), ("cx_27", "n0"), ("fc_cx", "p1"), ("fc_wc", "p0"), ("stub_wc", "n0"), ("pay_r", "p0"), ("beta", "n3"), ("erp", "p1"), ("g_cap", "p1"),
                  ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("inv_k", "p0"), ("bb_rem", "n0"), ("w_dcf", "p0")))},
    {"type": "table", "collapse": True, "summary": "展開：FY2026 公司指引（固定輸入）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag"] for k, f in (("g26_rev", "n1"), ("g26_op", "n1"), ("g26_ord", "n1"), ("g26_ni", "n1"), ("g26_eps", "n0"), ("g26_gm", "p1"),
                                                                    ("g26_sga", "p1"), ("g26_da", "n1"), ("g26_cx", "n1"), ("g26_dps", "n0"), ("xo_26", "n2"))]}]}

pai = {"id": "pai", "title": "實體 AI 曝險（機器人用氣動元件；只作對照，不進目標價）", "blocks": [
    {"type": "compare", "title": "機器人拉動的 SMC 氣動營收（十億日圓）與占合併營收比（情境 B 營收為分母）", "groups": [
        {"name": f"@Recon!A{PAIX['r26']}", "fmt": "n1", "items": [{"name": f"@Recon!B{find(RC, '項目', PAI0 - 2)}", "v": f"@Recon!B{PAIX['r26']}"},
                                                                    {"name": f"@Recon!C{find(RC, '項目', PAI0 - 2)}", "v": f"@Recon!C{PAIX['r26']}"},
                                                                    {"name": f"@Recon!D{find(RC, '項目', PAI0 - 2)}", "v": f"@Recon!D{PAIX['r26']}"}]},
        {"name": f"@Recon!A{PAIX['r30']}", "fmt": "n1", "items": [{"name": f"@Recon!{c}{find(RC, '項目', PAI0 - 2)}", "v": f"@Recon!{c}{PAIX['r30']}"} for c in "BCD"]},
        {"name": f"@Recon!A{PAIX['p26']}", "fmt": "p1", "items": [{"name": f"@Recon!{c}{find(RC, '項目', PAI0 - 2)}", "v": f"@Recon!{c}{PAIX['p26']}"} for c in "BCD"]},
        {"name": f"@Recon!A{PAIX['p30']}", "fmt": "p1", "items": [{"name": f"@Recon!{c}{find(RC, '項目', PAI0 - 2)}", "v": f"@Recon!{c}{PAIX['p30']}"} for c in "BCD"]}],
     "caption": "低／基準／高＝共用零組件線模型的機器人氣動市場三情境 × SMC 份額低／基準／高（同向）。人形、AMR、車用的氣動內容在線模型基準為零，未計入。"},
    {"type": "note", "parts": ["@README!A7"]},
    {"type": "table", "caption": f"@Recon!A{PAI0 - 2}", "head": [f"@Recon!{c}{PAI0 - 1}" for c in "ABCDE"],
     "rows": [[f"@Recon!A{r}", f"@Recon!B{r}|{fm}", f"@Recon!C{r}|{fm}", f"@Recon!D{r}|{fm}", f"@Recon!E{r}"] for r, fm in
              ((PAIX["m26"], "n1"), (PAIX["m30"], "n1"), (PAIX["sh"], "p0"), (PAIX["r26"], "n1"), (PAIX["r30"], "n1"), (PAIX["p26"], "p1"), (PAIX["p30"], "p1"), (PAIX["ir"], "n0"), (PAIX["cagr"], "p1"))]}]}

RE = wb["對帳"]; HRg = find(HR, "地區別營收（地域ごとの情報；顧客所在地）")
recon = {"id": "recon", "title": "對帳：公司指引、共識、歷史分部與線層級變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, '加權目標價 vs 賣方平均目標價')}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1", ""],
                "row_fmt": {str(find(RC, "FY2026 EPS")): {c: "n0" for c in "BCD"}, str(find(RC, "10 月受注指數", exact=False)): {c: "n2" for c in "BCD"},
                            str(find(RC, "加權目標價 vs 賣方平均目標價")): {c: "n0" for c in "BCD"}}}},
    {"type": "table", "caption": "地區別營收實績（十億日圓；有價證券報告書轉載）與合併營收對帳", "head": [f"@Hist!{c}{HRg}" for c in "ABCD"],
     "rowgen": {"sheet": "Hist", "rows": f"{HRg + 1}-{HRg + 7}", "cols": list("ABCD"), "fmts": [""] + ["n2"] * 3}},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（共用零組件、工業機器人）", "head": [f"@對帳!{c}2" for c in "ABCDG"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{RE.max_row}", "cols": ["A", "B", "C", "D", "G"], "fmts": ["", "", "", "", "tag"]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", "@Sens!D3"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價與市場資料的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3", "@In!I3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}"), nourl("In", f"I{inr(k)}")] for k, f in (("px", "n0"), ("rf", "p2"), ("pe_self", "n2"), ("cons_tp", "n0"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"In": ["B", "H", "I"], "Hist": ["A", "F"], "對帳": ["A", "C"]}, "empty": "沒有標為未取得的項目。"}]}

cfg = {"company": "SMC（SMC株式会社）", "ticker": "6273.T", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "半導體與資料中心帶動的工廠自動化設備投資高峰，能否在三情境下支持高於現價的目標價；機器人（實體 AI）對 SMC 的直接貢獻有多大",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：受注延續、毛利率維持）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：設備投資循環回落）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, pai, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

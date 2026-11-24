# 產生 physical_ai/hesai/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；模型改版移列後重跑本檔即可。設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/hesai/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/hesai/html.json
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

LAB = dict(dem="線 6 光達需求量（百萬顆；基準）", dm="情境需求乘數", shr="禾賽涵蓋率（出貨量 ÷ 線需求）", ship="禾賽光達出貨量（百萬顆）", asp="光達 ASP（RMB／顆）",
           aspu="光達 ASP（US$／顆；按 2026-11-20 匯率）", lid="光達營收（出貨量 × ASP）", sgi="策略成長業務營收（致動模組等）", prod="產品營收", svc="服務營收", rev="營收",
           chk1="檢查：光達＋策略成長業務 − 產品營收（應為 0）", chk2="檢查：分部合計 − 合併營收（應為 0）", gr="營收年增率", cgl="營業成本：光達", cgs="營業成本：策略成長業務",
           cgv="營業成本：服務", cogs="營業成本", gp="毛利", gm="毛利率", lgm="光達毛利率（模型推算；分業務毛利未揭露）", rd="研發費用", sm="銷售費用", ga="管理費用",
           oi="其他營業收益（政府補助等，淨額）", opex="營業費用合計（淨額）", op="營業利益", opm="營業利益率", nonop="業外收支（利息收入等）", pbt="稅前利益", tax="所得稅",
           ni="淨利", nim="淨利率", nig="淨利年增率", sh="加權平均稀釋股數（百萬股；拆股後）", eps="EPS（人民幣／股；稀釋）", epsa="EPS（人民幣／ADS；＝每股 × 8）",
           nopat="NOPAT（營業利益為正時 ×(1 − 稅率)）", da="折舊攤銷（資本支出五年平均攤提；未揭露，推算）", cx="資本支出", dwc="營運資金增加",
           fcff="FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", fcfad="自由現金流扣股利（收支關卡用）", liq="期末現金準備（未計新融資）",
           gap="資金缺口（最低流動資金 − 期末現金準備，正值＝需融資）")
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm)
    PR[s] = dict(hdr=h + 1, **{k: find(P, t, h) for k, t in LAB.items()})
B = PR["B"]
STD = list("CDEFGHIJKL")    # PL：C＝FY2023 … L＝FY2032
FCS = list("FGHIJKL")       # FY2026 起
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VT = dict(dcfr="DCF 情境價值推到 12 個月後", nct="目標時點淨現金（評價日淨現金 ＋ Stub FCFF ＋ FY2027 FCFF 與稅後業外；RMB M）",
          cu="Comparable（未調整）：EV／營收中位數 × FY2027 營收 ＋ 目標時點淨現金（每 ADS；US$）", ca="Comparable（相對倍數調整後，進加權）",
          tps="各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", tp="加權目標價（美元）", dcf="DCF 目標價（情境加權）", cmp="Comparable 目標價（情境加權，進加權）",
          px="現價（In 頁 px）", up="上檔空間", rt="估值評等（只看上檔空間）", gate="收支關卡：FY2026–FY2030 自由現金流扣股利合計（RMB M；情境 A／B／C）",
          gj="入不敷出判定", rec="投資建議", dist="距最近級距邊界（百分點）", bnd="翻轉目標價：升一級／降一級",
          gapmax="資金缺口：FY2026–FY2032 最大缺口（RMB M；情境 A／B／C；未計新融資）", gapyr="資金缺口：現金準備首次低於最低流動資金的年份（情境 A／B／C）",
          liqmin="資金缺口：FY2026–FY2032 最低期末現金準備（RMB M；情境 A／B／C）", opy="營業利益轉正年份（FY2026 起首次；情境 A／B／C）",
          pen30="光達出貨量 FY2030（百萬顆；情境 A／B／C）", pen35="光達出貨量 FY2032（百萬顆；情境 A／B／C）", gpm30="營業利益率 FY2030（情境 A／B／C）",
          a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻",
          cash="現價隱含：每 ADS 淨現金（評價日淨現金 ÷ ADS 數 ÷ 匯率；US$）", mkt="現價隱含：營運業務 EV（市值 − 淨現金，RMB M）",
          pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC", evs="隱含 EV／營收：加權目標價隱含 EV ÷ 情境 B FY2027 營收",
          evs2="現行 EV／TTM 營收（禾賽）÷ 同業中位數")
VR = {k: find(V, t) for k, t in VT.items()}
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")
def iv(k, f=""): return f"@In!E{inr(k)}" + (f"|{f}" if f else "")
def il(k): return f"@In!B{inr(k)}"
def nourl(sh, coord, f=""):
    v = wb[sh][coord].value
    return "（出處含網址，見 xlsx）" if isinstance(v, str) and "http" in v else f"@{sh}!{coord}" + (f"|{f}" if f else "")

DK = {k: find(D, t) for k, t in dict(ke="ke（股權成本）", e="股權市值 E（RMB M）", dd="負債 D（RMB M）", kd="稅後 kd", wacc="WACC", g="永續成長率 g＝MIN(g_cap, rf)", ronic="RONIC",
                                       nc="淨現金（2026-09-30）＝現金準備 − 債務與租賃（RMB M）").items()}
DB = find(D, "情境 B（基準）")
DBR = {k: find(D, t, DB) for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", eq="股權價值（RMB M）", psv="每 ADS 價值（評價日；US$）",
                                          d12="DCF 情境價值推到 12 個月後（×(1＋ke)^t；US$／ADS）", tvsh="終值占 EV", alt="終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每 ADS 價值（推到 12 個月後）").items()}
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
               {"type": "note", "parts": ["讀法：結果最敏感的是情境權重（情境 A 的 DCF 是加權值的主要來源）、涵蓋率、ASP 年降幅與光達毛利率；單看情境 B 的目標價見估值章。"]},
               {"type": "kv", "items": [["@Sens!A6", "@Sens!D6|n2"], ["模型狀態", "@Sens!D7"], ["build_id", "@Sens!D3"]]}]
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
        "note": ["估值評等", vv("rt"), "；收支關卡", vv("gj"), "（情境 B 自由現金流，FY2026–FY2030 合計 ", vv("gate", "C", "n0"),
                 " 百萬人民幣）。加權值主要來自樂觀情境的 DCF；只看基準情境為 ", vv("tps", "C", "n2"), " 美元。結論建立在 Assumed 輸入上，現價附近的上檔空間不足以改變評等。"]}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "營收依產品與服務堆疊（百萬人民幣）與涵蓋率：FY2026–FY2032（情境 B）", "x": f"@PL!F{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev", "B", "F"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "F"), "color": c} for k, c in (("lid", "#0b7a80"), ("sgi", "#c29a2b"), ("svc", "#6fa8dc"))],
     "lines": [{"name": [f"@PL!A{B['shr']}"], "y": pl("shr", "B", "F"), "axis": "r", "color": "#b4472f", "labels": True}],
     "caption": "堆疊＝光達營收、策略成長業務與服務（合計等於營收；公司官方分類為產品與服務兩類）；折線＝禾賽光達出貨量占線模型 需求的比例（右軸）。FY2026 為九個月實績加第四季指引；FY2027 起為模型。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["核心命題的承重變數：情境 B 線模型 需求乘數 ", iv("dm_B", "n2"), "、FY2030 涵蓋率 ", iv("shr_B", "p0"), "、ASP 年降 ", iv("asd_B", "p0"), "、光達毛利率 FY2030 ", iv("gmt_B", "p0"), "；FY2026 涵蓋率校準 ", iv("shr26", "n2"), "。"]]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n2"), " 美元（DCF ", vv("dcf", "B", "n2"), "、Comparable ", vv("cmp", "B", "n2"), "），對現價 ", iv("px", "n2"), " 美元上檔 ", vv("up", "B", "sp1"),
         "；估值評等", vv("rt"), "、投資建議", vv("rec"), "。情境 A／B／C 目標價 ", vv("tps", "B", "n2"), "／", vv("tps", "C", "n2"), "／", vv("tps", "D", "n2"), "。"],
        ["市場定價：每 ADS 淨現金 ", vv("cash", "B", "n2"), " 美元（約占現價一半），現價隱含營運業務 EV ", vv("mkt", "B", "n0"), " 百萬人民幣；禾賽 EV／TTM 營收約為同業中位數的 ", vv("evs2", "B", "n2"), " 倍。"],
        ["出貨量（情境 B）：FY2030 ", vv("pen30", "C", "n1"), " 百萬顆、FY2032 ", vv("pen35", "C", "n1"), " 百萬顆；營業利益率 FY2030 ", vv("gpm30", "C", "p1"), "。情境 C 營業利益率 FY2030 ", vv("gpm30", "D", "p1"), "。"],
        ["現金：現金準備遠高於最低流動資金，三情境皆無資金缺口（情境 C 最低 ", vv("liqmin", "D", "n0"), " 百萬人民幣）。"],
        "@README!A7"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A8"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "三個目標價與現價（美元／ADS；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}],
     "ref": {"label": il("px"), "value": iv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n2"), ("cmp", "n2"), ("tp", "n2"), ("px", "n2"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("dist", "n1"), ("bnd", ""))]},
    {"type": "table", "caption": "情境目標價（美元／ADS；目標時點淨現金為百萬人民幣）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", "n2"],
                "row_fmt": {str(VR["nct"]): {c: "n0" for c in "BCDE"}}}},
    {"type": "table", "caption": "收支關卡、現金與營業利益轉正（情境 A／B／C）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@Val!A{VR[k]}"] + [f"@Val!{c}{VR[k]}|{f}" for c in "BCD"] for k, f in (("gate", "n0"), ("liqmin", "n0"), ("gapmax", "n0"), ("gapyr", ""), ("opy", "yr"))]},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（百萬人民幣；每 ADS 為美元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n0"), ("eq", "n0"), ("psv", "n2"), ("d12", "n2"), ("tvsh", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n0"), ("tv", "n0"), ("pvtv", "n0"), ("alt", "n2"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n3"), ("erp", "p1"), ("crp", "p2"), ("sh_v", "n1"), ("liq_v", "n1"), ("debt_v", "n0"), ("fx", "n4"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n1", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n0", "", ""], [f"@DCF!A{DK['dd']}", f"@DCF!B{DK['dd']}|n1", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因與隱含倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("cash", "n2"), ("mkt", "n0"), ("pd2", "n2"), ("evs", "x1"), ("evs2", "n2"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業 EV／營收與中位數（市值以收盤檔收盤價重算）", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}")] for k in ("ev_rs", "ev_oust", "ev_mbly", "ev_med", "ev_hzn", "ev_self")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"))]}]}

RVH = find(RC, "反解目標")
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個情境 B 的承重驅動，求使加權目標價等於現價（上檔為零）及落在評等邊界所需的值，其他輸入不變。「無解」表示在搜尋區間內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{find(RC, '反解（整本重算', exact=False)}", "head": [f"@Recon!{c}{RVH}" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{RVH + 1}-{RC.max_row}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "", "n3", "n2", ""]}}]}

pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "營收（百萬人民幣）× 情境 A／B／C 與營業利益率（情境 B）：FY2026–FY2032", "x": f"@PL!F{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：營收"], "y": pl("rev", s, "F"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['opm']}", "（情境 B）"], "y": pl("opm", "B", "F"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境營收；折線＝情境 B 營業利益率（右軸）。FY2026 三情境相同，分化從 FY2027 開始。"},
    {"type": "note", "parts": ["會計口徑：US GAAP（公司法定財報準則，報表幣別人民幣）。營業利益含其他營業收益（政府補助等），為公司報表口徑；股權報酬留在費用內。歷史股數已還原為拆股後口徑（細節見 xlsx Hist 頁）；FY2023 只取得部分科目。"]},
    {"type": "table", "caption": "標準損益表（US GAAP；百萬人民幣，EPS 為人民幣／股）：FY2023–FY2025 實績、FY2026 九個月實績加第四季指引、FY2027 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [[f"@PL!A{B[k]}"] + [f"@PL!{c}{B[k]}|{f}" for c in STD] for k, f in
              (("lid", "n1"), ("sgi", "n1"), ("prod", "n1"), ("svc", "n1"), ("rev", "n1"), ("gr", "p1"), ("cogs", "n1"), ("gp", "n1"), ("gm", "p1"), ("rd", "n1"), ("sm", "n1"), ("ga", "n1"),
               ("oi", "n1"), ("opex", "n1"), ("op", "n1"), ("opm", "p1"), ("nonop", "n1"), ("pbt", "n1"), ("tax", "n1"), ("ni", "n1"), ("nim", "p1"), ("nig", "p1"),
               ("sh", "n1"), ("eps", "n3"), ("epsa", "n2"))]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "營收（百萬人民幣）", "fmt": "n0", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "營業利益（百萬人民幣）", "fmt": "n0", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "營業利益率", "fmt": "p1", "rows": {s: pl("opm", s, "F") for s in "ABC"}},
                 {"name": "EPS（人民幣／ADS）", "fmt": "n2", "chart": True, "rows": {s: pl("epsa", s, "F") for s in "ABC"}},
                 {"name": "期末現金準備（百萬人民幣；未計新融資）", "fmt": "n0", "chart": True, "rows": {s: pl("liq", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流與現金準備（百萬人民幣）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in FCS],
     "rows": [[f"@PL!A{B[k]}"] + [f"@PL!{c}{B[k]}|n1" for c in FCS] for k in ("nopat", "da", "cx", "dwc", "fcff", "fcfad", "liq", "gap")]}]}

PARM = lambda keys: [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in keys]
PH = ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"]
drivers = {"id": "drivers", "title": "驅動因子：線模型 需求 × 涵蓋率 → 出貨量 × ASP → 光達營收", "blocks": [
    {"type": "theme", "title": "光達 ASP（人民幣／顆，左軸）與出貨量（百萬顆，右軸）：FY2026–FY2032（情境 B）", "x": f"@PL!F{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "n1", "stack": False, "pad": 1.3,
     "bars": [{"name": [f"@PL!A{B['asp']}"], "y": pl("asp", "B", "F"), "color": "#0b7a80"}],
     "lines": [{"name": [f"@PL!A{B['ship']}"], "y": pl("ship", "B", "F"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "出貨量＝線模型 基準需求 × 需求乘數 × 涵蓋率；ASP 以 FY2026 九個月實績校準後逐年下降。營收成長是否快過降價，是本命題的核心。"},
    {"type": "note", "parts": [f"@PL!A{find(P, 'PL — ', exact=False)}"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("dm", "n2"), ("shr", "p0"), ("asd", "p0"), ("gmt", "p0"), ("gsgi", "p0"), ("gsgl", "p0"), ("grd", "p0"))]},
    {"type": "table", "caption": "命題模組推導（情境 B）：線需求 → 出貨量 → ASP → 光達營收 → 產品與服務 → 營收；核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in FCS],
     "rowgen": {"sheet": "PL", "rows": f"{B['dem']}-{B['gr']}", "cols": ["A"] + FCS, "fmts": [""] + ["n1"] * len(FCS),
                "row_fmt": {str(B["gr"]): {c: "p1" for c in FCS}, str(B["shr"]): {c: "p0" for c in FCS}, str(B["dem"]): {c: "n2" for c in FCS}, str(B["ship"]): {c: "n2" for c in FCS}}}},
    {"type": "table", "caption": "成本與毛利推導（情境 B）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in FCS],
     "rowgen": {"sheet": "PL", "rows": f"{B['cgl']}-{B['lgm']}", "cols": ["A"] + FCS, "fmts": [""] + ["n1"] * len(FCS),
                "row_fmt": {str(B[k]): {c: "p1" for c in FCS} for k in ("gm", "lgm")}}},
    {"type": "h3", "text": "一、光達（命題模組：線需求、涵蓋率、ASP、毛利）"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": PH,
     "rows": PARM((("dm_B", "n2"), ("shr_B", "p0"), ("asd_B", "p0"), ("gmt_B", "p0"), ("q4rev", "n0"), ("q4gm", "p0"), ("q4ship", "n2")))},
    {"type": "table", "collapse": True, "summary": "展開：讀自線模型 的逐年需求與 FY2026 涵蓋率校準（固定輸入）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag"] for k, f in [(f"dmd{y}", "n2") for y in range(26, 33)] + [("shr26", "n2"), ("ship9m", "n3")]]},
    {"type": "h3", "text": "二、策略成長業務與服務"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": PH,
     "rows": PARM((("sgi26", "n0"), ("gsgi_B", "p0"), ("gsgl_B", "p0"), ("gm_sgi", "p0"), ("svc26", "n0"), ("svc_g", "p0"), ("gm_sv", "p0")))},
    {"type": "h3", "text": "三、全公司（費用、現金流、折現）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": PH,
     "rows": PARM((("grd_B", "p0"), ("sm_p", "p1"), ("gga", "p0"), ("goi", "p0"), ("yld", "p1"), ("tax_w", "p0"), ("dil_g", "p1"), ("fc_wc", "p0"), ("cx_p26", "p0"), ("cx_p", "p0"),
                   ("sh_x", "p0"), ("debt_v", "n0"), ("beta", "n3"), ("erp", "p1"), ("crp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("w_dcf", "p0")))},
    {"type": "table", "collapse": True, "summary": "展開：FY2026 錨點（固定輸入）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag"] for k, f in (("rd_q4", "n0"), ("sm_q4", "n0"), ("ga_q4", "n0"), ("oi_q4", "n0"), ("nonop_q4", "n0"), ("sh26", "n1"), ("min_cash", "n0"))]}]}

Q = wb["Q"]; qs = find(Q, "項目"); qe = find(Q, "營收年增（Q3；9M）", qs)
RE = wb["對帳"]
recon = {"id": "recon", "title": "對帳：公司指引、季度財務與線模型變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, 'FY2030 光達營收', exact=False)}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""]}},
    {"type": "table", "caption": "季度與前三季對照（百萬人民幣；出貨量為顆數）", "head": [f"@Q!{c}{qs}" for c in "ABCDE"],
     "rowgen": {"sheet": "Q", "rows": f"{qs + 1}-{qe}", "cols": list("ABCDE"), "fmts": ["", "n1", "n1", "n1", "n1"],
                "row_fmt": {str(r): {c: "n0" for c in "BCDE"} for r in range(find(Q, "光達出貨量（顆）", qs), find(Q, "光達出貨量（顆）", qs) + 3)} | {str(find(Q, "毛利率", qs)): {c: "p1" for c in "BCDE"}, str(qe): {c: "p1" for c in "BCDE"}}}},
    {"type": "table", "collapse": True, "summary": "展開：線變數對帳（讀入值、用法、不一致處）", "head": [f"@對帳!{c}2" for c in "ABCDEF"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{RE.max_row}", "cols": ["A", "B", "C", "D", "E", "F"], "fmts": ["", "", "g", "", "", ""]}}]}

RK = wb["Risk"]
MR = find(RK, "未找到／未取得清單", exact=False)
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", "@Sens!D3"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價與市場資料的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!I3"],
     "rows": [[il("px"), iv("px", "n2"), f"@In!G{inr('px')}|tag", nourl("In", f"I{inr('px')}")],
              [il("px_hk"), iv("px_hk", "n2"), f"@In!G{inr('px_hk')}|tag", nourl("In", f"I{inr('px_hk')}")],
              [il("fx"), iv("fx", "n4"), f"@In!G{inr('fx')}|tag", nourl("In", f"I{inr('fx')}")],
              [il("rf"), iv("rf", "p2"), f"@In!G{inr('rf')}|tag", nourl("In", f"I{inr('rf')}")]]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, MR - 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未找到」與「未取得」清單"},
    {"type": "table", "collapse": True, "summary": "展開：未找到與未取得的資料、替代做法", "head": [f"@Risk!{c}{MR + 1}" for c in "ABCD"],
     "rows": [[f"@Risk!{c}{r}" for c in "ABCD"] for r in range(MR + 2, RK.max_row + 1)]}]}

cfg = {"company": "禾賽 Hesai", "ticker": "HSAI", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "光達出貨量成長能否快過單價下降、讓毛利率守住目前水準；現價是否已把營運業務定價在接近淨現金之上",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：需求高於線基準、涵蓋率守住、降價慢）"}, {"key": "B", "name": "情境 B（基準：線模型 基準需求、涵蓋率逐年下降、ASP 逐年下降）"}, {"key": "C", "name": "情境 C（保守：需求偏低、份額流失、降價快）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

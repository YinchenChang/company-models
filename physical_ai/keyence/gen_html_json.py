# 產生 physical_ai/keyence/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/keyence/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/keyence/html.json
import glob, json, os, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
X = sys.argv[1] if len(sys.argv) > 1 else sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))[0]
wb = openpyxl.load_workbook(X, data_only=True)
P, V, I, D, SE, RC, XA = wb["PL"], wb["Val"], wb["In"], wb["DCF"], wb["Sens"], wb["Recon"], wb["實體AI"]


def find(ws, text, start=1, col=1, exact=True):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and (v == text if exact else v.startswith(text)):
            return r
    raise KeyError(f"{ws.title}: {text}")


IN = {I.cell(r, 1).value: r for r in range(1, I.max_row + 1) if isinstance(I.cell(r, 1).value, str)}
def inr(k): return IN[k]

LABS = dict(dom="國內營收（子列）", ovs="海外營收（子列）", rev="合併營收（單一分部：電子應用機器）", chk="檢查：國內＋海外 − 合併營收（應為 0）", gr="營收年增率",
            ovsr="海外營收占比", cogs="營業成本（売上原価）", gp="毛利", gm="毛利率", sga="營業費用（販売費及び一般管理費）", sgar="營業費用率", op="營業利益",
            opm="營業利益率", nonop="業外收支淨額（其他收入淨額）", pbt="稅前利益", tax="所得稅", ni="稅後淨利（歸屬母公司）", nim="淨利率", nig="淨利年增率",
            sh="平均股數（百萬股，分割後）", eps="EPS（日圓，分割後；無潛在股份）", da="折舊攤銷", cx="資本支出", dwc="營運資金增加",
            fcff="FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", dps="每股股利（日圓，分割後）", div="股利", fcfad="自由現金流扣股利（收支關卡用）")
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm)
    PR[s] = {k: find(P, t, h) for k, t in LABS.items()}; PR[s]["hdr"] = h + 1
B = PR["B"]
STD = list("CDEFGHIJKL")   # FY2023 … FY2032
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VR = {k: find(V, t) for k, t in dict(dcfr="DCF 情境價值推到目標時點", ntm="目標時點 NTM EPS（4/12 × FY2027 ＋ 8/12 × FY2028；分割後）", cu="Comparable（未調整）：同業前瞻 P/E 中位數 × NTM EPS",
                                       ca="Comparable（相對倍數調整後，進加權）", tps="各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", tp="加權目標價（日圓，分割後）",
                                       dcf="DCF 目標價（情境加權）", cmp="Comparable 目標價（情境加權，進加權）", px="現價（In 頁 px，分割調整後）", up="上檔空間", rt="估值評等（只看上檔空間）",
                                       gate="收支關卡：FY2026–FY2030 自由現金流扣股利合計（¥B；情境 A／B／C）", gj="入不敷出判定", rec="投資建議", dist="距最近級距邊界（百分點）",
                                       bnd="翻轉目標價：升一級／降一級", a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻",
                                       pd1="期間：模型折現（評價日 2027-03-20 → 目標時點，年）", pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC",
                                       pe="隱含前瞻 P/E：加權目標價 ÷ 情境 B 目標時點 NTM EPS", pe2="現行前瞻 P/E（現價 ÷ 共識 FY2026 EPS）÷ 同業中位數",
                                       dsh="DCF 在加權目標價的占比", cash="每股淨現金與投資（分割後，日圓）", cashp="每股淨現金占現價").items()}
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
DK = {k: find(D, t) for k, t in dict(ke="ke（股權成本）", e="股權市值 E（¥B）", wacc="WACC", g="永續成長率 g＝MIN(g_cap, rf)", ronic="RONIC", kd="稅後 kd（無負債，不適用）",
                                       nc="淨現金與投資（2026-09-20）＝現金及預金＋有價證券 ＋ 投資有価証券 × inv_k − 有息負債").items()}
DBR = {k: find(D, t, DB) for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", eq="股權價值", psv="每股價值（評價日，分割後）",
                                          d12="DCF 情境價值推到目標時點（×(1＋ke)^t − 期間股利）", tvsh="終值占 EV",
                                          alt="終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到目標時點）").items()}
DA, DC = find(D, "情境 A（樂觀）"), find(D, "情境 C（保守）")
def dsc(k):
    t = D.cell(DBR[k], 1).value
    return [f"@DCF!B{find(D, t, DA)}", f"@DCF!B{DBR[k]}", f"@DCF!B{find(D, t, DC)}"]

# ---------------- 敏感度 ----------------
hdr_s = find(SE, "驅動")
s_rows = [r for r in range(hdr_s + 1, SE.max_row + 1) if isinstance(SE.cell(r, 5).value, (int, float))]
lo = [r for r in s_rows if str(SE.cell(r, 1).value).endswith("（低）")]
hi = [r for r in s_rows if str(SE.cell(r, 1).value).endswith("（高）")]
oth = [r for r in s_rows if r not in lo and r not in hi]
s_end = max(s_rows)
TOR = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": "E", "fmt": "sp1"}
def big(rs): return [r for r in rs if abs(SE.cell(r, 5).value) >= 0.1]
def top(rs, n=6): return sorted(rs, key=lambda r: -abs(SE.cell(r, 5).value))[:n]
def main_tor(side, rs):
    if big(rs):
        return dict(TOR, title=f"加權目標價相對基準：輸入取{side}值（變動達兩位數百分比的案例）", rows=",".join(map(str, big(rs))))
    return dict(TOR, title=f"加權目標價相對基準：輸入取{side}值（沒有變動達兩位數百分比的案例；列出影響最大的六項）", rows=",".join(map(str, top(rs))))
sens_blocks = [main_tor("低", lo), main_tor("高", hi),
               {"type": "note", "parts": ["@Sens!A1"]},
               {"type": "note", "parts": ["單一參數一次只動一個，其餘不變、整本重算；所有單一參數案例的估值評等都維持基準評等，只有「只看情境 C」一案跨級。情境權重與終值替代法見下圖。"]},
               {"type": "kv", "items": [["@Sens!A6", "@Sens!D6|n0"], ["@Sens!A7", "@Sens!D7"], ["@Sens!A3", "@Sens!D3"]]}]
if oth:
    sens_blocks.append(dict(TOR, title="加權目標價相對基準：只看單一情境、終值替代法", rows=",".join(map(str, oth))))
for side, rs in (("低", lo), ("高", hi)):
    shown = set(map(int, main_tor(side, rs)["rows"].split(",")))
    rest = [r for r in rs if r not in shown]
    if rest:
        sens_blocks.append(dict(TOR, title=f"加權目標價相對基準：輸入取{side}值，其餘案例", rows=",".join(map(str, rest)), collapse=True,
                                summary=f"展開：輸入取{side}值的其餘案例"))
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：全部案例（加權目標價、DCF 目標價與估值評等）",
                    "head": [f"@Sens!{c}{hdr_s}" for c in "ABDEFGH"],
                    "rowgen": {"sheet": "Sens", "rows": f"{hdr_s + 1}-{s_end}", "cols": ["A", "B", "D", "E", "F", "G", "H"], "fmts": ["", "", "n0", "sp1", "n0", "sp1", ""]}})
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}

hero = {"rec": vv("rec"),
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n0"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n0"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n0"), ""],
                  [il("px"), iv("px", "n0"), ["現價日 ", iv("px_date")]]],
        "note": ["估值評等", vv("rt"), "；收支關卡", vv("gj"), "（情境 B 自由現金流扣股利 FY2026–FY2030 合計 ", vv("gate", "C", "n0"),
                 " 十億日圓）。DCF 低於現價、同業倍數法略高；實體 AI 只占營收一小部分。股價與每股數字皆為一股分五股後的口徑。"]}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "合併營收依國內／海外堆疊（十億日圓）與營業利益率：FY2023–FY2032（情境 B）", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k), "color": c} for k, c in (("dom", "#c29a2b"), ("ovs", "#0b7a80"))],
     "lines": [{"name": [f"@PL!A{B['opm']}"], "y": pl("opm"), "axis": "r", "color": "#3b3838", "labels": True},
               {"name": [f"@PL!A{B['gm']}"], "y": pl("gm"), "axis": "r", "color": "#b4472f", "dash": True}],
     "caption": "堆疊＝國內與海外營收（合計等於單一官方分部的合併營收）；實線＝營業利益率、虛線＝毛利率（右軸）。FY2023–FY2025 實績，FY2026 為上半實績加下半模型，FY2027 起為模型。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["承重變數：海外當地貨幣營收增速（情境 B：FY2027 ", iv("go27_B", "p0"), "、FY2028–FY2029 每年 ", iv("gom_B", "p0"), "）、匯率換算效果、毛利率（", iv("gm_B", "p1"),
         "），以及同業倍數是否給 Keyence 溢價。"],
        "@README!A7"]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n0"), " 日圓（DCF ", vv("dcf", "B", "n0"), "、Comparable ", vv("cmp", "B", "n0"), "），對現價 ", iv("px", "n0"), " 日圓上檔 ", vv("up", "B", "sp1"),
         "，估值評等", vv("rt"), "。"],
        ["本波成長：FY2026 上半營收年增 ", iv("h1_ovs_lc", "p1"), "（海外當地貨幣）、國內 ", iv("h1_dom_g", "p1"), "，匯率另貢獻 ", iv("h1_fx", "p1"),
         "；業種以半導體與電子設備最強，屬 AI 算力投資的間接需求。"],
        ["倍數：Keyence 現行前瞻本益比 ", iv("pe_self", "n1"), " 倍，為同業中位數 ", iv("pe_med", "n1"), " 倍的 ", vv("pe2", "B", "n2"), " 倍；每股淨現金與投資占現價 ", vv("cashp", "B", "p1"), "。"],
        ["評等邊界：", "@README!A5"],
        "@README!A6"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A8"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（日圓，分割後；豎線＝現價）", "fmt": "n0",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}, {"label": il("cons_tp"), "value": iv("cons_tp"), "series": 1}],
     "ref": {"label": il("px"), "value": iv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n0"), ("cmp", "n0"), ("tp", "n0"), ("px", "n0"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("dist", "n1"), ("bnd", ""))]
             + [[il("cons_tp"), iv("cons_tp", "n0")]]},
    {"type": "table", "caption": "情境目標價（日圓；情境權重見下方）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n0", "n0", "n0", "n0"]}},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（十億日圓；每股為日圓、分割後）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n0"), ("eq", "n0"), ("psv", "n0"), ("d12", "n0"), ("tvsh", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n0"), ("tv", "n0"), ("pvtv", "n0"), ("alt", "n0"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n3"), ("erp", "p1"), ("crp", "p2"), ("sh_v", "n2"), ("split_r", "n0"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n0", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n0", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因、期間與隱含倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("dsh", "p1"), ("pd1", "n2"), ("pd2", "n2"), ("pe", "x1"), ("pe2", "n2"), ("cash", "n0"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻 P/E 與中位數", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}")] for k in ("pe_cgnx", "pe_omron", "pe_hexa", "pe_novt", "pe_smc", "pe_fanuc", "pe_med", "pe_vm", "pe_fa", "pe_self")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": PH,
     "rows": prm((("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"), ("dps_tgt", "n0")))}]}

rv0 = find(RC, "反解目標")
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價（上檔為零）及落在評等邊界所需的值，其他輸入不變。「無解」表示在搜尋區間內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{find(RC, '反解（整本重算', exact=False)}", "head": [f"@Recon!{c}{rv0}" for c in "ACDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{rv0 + 1}-{RC.max_row}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n3", "n3", ""]}}]}

H_ = wb["Hist"]
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "EPS（日圓，分割後）× 情境 A／B／C 與淨利率（情境 B）：FY2023–FY2032", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：EPS"], "y": pl("eps", s), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['nim']}", "（情境 B）"], "y": pl("nim"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境 EPS（分割後口徑；歷史年度以原文 EPS 除以五）；折線＝情境 B 淨利率（右軸）。FY2025 以前為實績，三情境相同。"},
    {"type": "note", "parts": ["會計口徑：日本會計基準（公司法定財報準則），本表即法定財報的 GAAP 口徑，未做 Non-GAAP 調整；公司無特別損益與非支配股東損益，稅前利益等於經常利益。公司不揭露業績預測，FY2026 為上半實績加下半模型。"]},
    {"type": "table", "caption": "標準損益表（日本基準／GAAP 口徑；十億日圓，EPS 為日圓、分割後）：FY2023–FY2025 實績、FY2026 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [["營收（合併；單一分部）"] + [f"@PL!{c}{B['rev']}|n1" for c in STD],
              [f"@PL!A{B['gr']}"] + ["—"] + [f"@PL!{c}{B['gr']}|p1" for c in STD[1:]],
              *[[f"@PL!A{B[k]}"] + [f"@PL!{c}{B[k]}|{f}" for c in STD] for k, f in (("cogs", "n1"), ("gp", "n1"), ("gm", "p1"), ("sga", "n1"), ("op", "n1"), ("opm", "p1"),
                                                                                     ("nonop", "n1"), ("pbt", "n1"), ("tax", "n1"), ("ni", "n1"), ("nim", "p1"))],
              [f"@PL!A{B['nig']}"] + ["—"] + [f"@PL!{c}{B['nig']}|p1" for c in STD[1:]],
              [f"@PL!A{B['sh']}"] + [f"@PL!{c}{B['sh']}|n1" for c in STD],
              [f"@PL!A{B['eps']}"] + [f"@PL!{c}{B['eps']}|n1" for c in STD],
              [f"@PL!A{B['dps']}"] + [f"@PL!{c}{B['dps']}|n1" for c in STD]]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "合併營收（十億日圓）", "fmt": "n1", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "營業利益（十億日圓）", "fmt": "n1", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "營業利益率", "fmt": "p1", "rows": {s: pl("opm", s, "F") for s in "ABC"}},
                 {"name": "EPS（日圓，分割後）", "fmt": "n0", "chart": True, "rows": {s: pl("eps", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣股利（十億日圓；收支關卡用）", "fmt": "n1", "rows": {s: pl("fcfad", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流（十億日圓）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[1:]],
     "rowgen": {"sheet": "PL", "rows": f"{B['da']}-{B['fcfad']}", "cols": ["A"] + STD[1:], "fmts": [""] + ["n1"] * len(STD[1:])}}]}

drivers = {"id": "drivers", "title": "驅動因子：國內＋海外（當地貨幣×匯率）→ 合併營收", "blocks": [
    {"type": "theme", "title": "合併營收 × 情境 A／B／C（十億日圓）與海外營收占比（情境 B）：FY2023–FY2032", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：合併營收"], "y": pl("rev", s), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['ovsr']}", "（情境 B）"], "y": pl("ovsr"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境合併營收；折線＝情境 B 海外營收占比（右軸）。分化從 FY2027 開始；情境 C 含日圓升值與設備投資回落。"},
    {"type": "note", "parts": [f"@PL!A{find(P, 'PL — ', exact=False)}"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("gd27", "p1"), ("gdm", "p1"), ("gdlt", "p1"), ("go27", "p1"), ("gom", "p1"), ("golt", "p1"), ("fx27", "p1"), ("gm", "p1"))]},
    {"type": "table", "caption": "分部營收推導（情境 B；十億日圓）：國內＋海外 → 單一官方分部；核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['dom']}-{B['ovsr']}", "cols": ["A"] + STD, "fmts": [""] + ["n1"] * len(STD),
                "row_fmt": {str(B["gr"]): {c: "p1" for c in STD}, str(B["ovsr"]): {c: "p1" for c in STD}}}},
    {"type": "table", "caption": "毛利與營業費用推導（情境 B；十億日圓）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['cogs']}-{B['opm']}", "cols": ["A"] + STD, "fmts": [""] + ["n1"] * len(STD),
                "row_fmt": {str(B[k]): {c: "p1" for c in STD} for k in ("gm", "sgar", "opm")}}},
    {"type": "h3", "text": "一、FY2026（上半實績＋下半模型）"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": PH,
     "rows": prm((("gd26", "p0"), ("go26", "p0"), ("fx26", "p0"), ("gm26", "p1"), ("sga26_r", "p0")))},
    {"type": "table", "collapse": True, "summary": "展開：FY2026 上半實績（固定輸入）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag"] for k, f in (("h1_rev", "n1"), ("h1_cogs", "n1"), ("h1_sga", "n1"), ("h1_op", "n1"), ("h1_nonop", "n1"), ("h1_ni", "n1"),
                                                                   ("h1_dom_g", "p1"), ("h1_ovs_lc", "p1"), ("h1_fx", "p1"))]},
    {"type": "h3", "text": "二、FY2027 起營收與利潤率（情境 B 的區間）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": PH,
     "rows": prm((("gd27_B", "p0"), ("gdm_B", "p0"), ("gdlt_B", "p0"), ("go27_B", "p0"), ("gom_B", "p0"), ("golt_B", "p0"), ("fx27_B", "p0"), ("gm_B", "p1"), ("e_sga", "n2"), ("sga_fl", "p0")))},
    {"type": "h3", "text": "三、全公司（損益下半段、現金流、折現、股數）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": PH,
     "rows": prm((("nonop", "n1"), ("tax_r", "p0"), ("fc_da", "p1"), ("fc_cx", "p1"), ("fc_wc", "p0"), ("beta", "n3"), ("erp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"),
                  ("wacc_adj", "p1"), ("inv_k", "p0"), ("dps_tgt", "n0"), ("w_dcf", "p0")))},
    {"type": "table", "collapse": True, "summary": "展開：股票分割、股數與資產負債表（固定輸入）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag"] for k, f in (("split_r", "n0"), ("sh_iss", "n3"), ("sh_trs", "n3"), ("sh_v", "n2"), ("sh_26", "n2"), ("cash_v", "n1"), ("inv_v", "n1"),
                                                                   ("debt_v", "n1"), ("stub_frac", "n2"), ("stub_div", "n1"), ("dps_26", "n0"), ("pay_r", "p0"))]}]}

YX = "DEFGHIJ"
xr = {k: find(XA, t) for k, t in dict(lo="Keyence 機器人相關感測視覺營收（低：低市場 × 份額低端）", b="Keyence 機器人相關感測視覺營收（基準：基準市場 × 份額基準）",
                                      hi="Keyence 機器人相關感測視覺營收（高：高市場 × 份額高端）", sb="占合併營收（基準）", slo="占合併營收（低）", shi="占合併營收（高）").items()}
physai = {"id": "physai", "title": "實體 AI 曝險（另列；不進目標價）", "blocks": [
    {"type": "theme", "title": "機器人工作站相關感測視覺營收（十億日圓；低／基準／高）與占營收比（基準）：FY2026–FY2032", "x": "@實體AI!D2:J2", "fmt": "n1", "fmt_r": "p1", "stack": False, "pad": 1.3,
     "bars": [{"name": [f"@實體AI!A{xr[k]}"], "y": f"@實體AI!D{xr[k]}:J{xr[k]}", "color": c} for k, c in (("lo", "#d9a08f"), ("b", "#0b7a80"), ("hi", "#7fb8bb"))],
     "lines": [{"name": [f"@實體AI!A{xr['sb']}"], "y": f"@實體AI!D{xr['sb']}:J{xr['sb']}", "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "長條＝共用零組件線模型的工業感測與機器視覺市場（只計機器人工作站）乘以 Keyence 份額假設，再換成日圓；折線＝基準值占情境 B 合併營收（右軸）。份額未取得，全部為假設區間。"},
    {"type": "note", "parts": ["@實體AI!A1"]},
    {"type": "table", "caption": "曝險推導（年度欄為 FY）", "head": [f"@實體AI!{c}2" for c in "AB" + YX],
     "rowgen": {"sheet": "實體AI", "rows": "3-14", "cols": ["A", "B"] + list(YX), "fmts": ["", ""] + ["n1"] * 7,
                "row_fmt": {str(r): {c: "p1" for c in YX} for r in (12, 13, 14)}}},
    {"type": "note", "parts": ["@實體AI!A16"]}, {"type": "note", "parts": ["@實體AI!A17"]}, {"type": "note", "parts": ["@實體AI!A18"]},
    {"type": "table", "caption": "份額與匯率假設", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag"] for k, f in (("kshr_lo", "p0"), ("kshr", "p0"), ("kshr_hi", "p0"), ("usdjpy", "n2"))]}]}

Q = wb["Q"]; RE = wb["對帳"]
recon = {"id": "recon", "title": "對帳：共識、季度損益與線層級變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, '加權目標價 vs 賣方平均目標價')}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n1", "n1", "sp1", ""]}},
    {"type": "table", "caption": "@Q!A1", "head": [f"@Q!{c}2" for c in "ABCDEFGH"],
     "rowgen": {"sheet": "Q", "rows": f"3-{find(Q, '檢查：營收', exact=False)}", "cols": list("ABCDEFGH"), "fmts": [""] + ["n1"] * 6 + ["tag"],
                "row_fmt": {str(find(Q, "營業利益率")): {c: "p1" for c in "BCDEFG"}, str(find(Q, "毛利率")): {c: "p1" for c in "BCDEFG"}}}},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（共用零組件、工業機器人）", "head": [f"@對帳!{c}2" for c in "ABCDGI"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{RE.max_row}", "cols": ["A", "B", "C", "D", "G", "I"], "fmts": ["", "", "", "", "tag", ""]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", "@Sens!D3"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價、分割與市場資料的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3", "@In!I3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}"), nourl("In", f"I{inr(k)}")] for k, f in (("px", "n0"), ("split_r", "n0"), ("rf", "p2"), ("pe_self", "n2"), ("usdjpy", "n2"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"In": ["B"], "Hist": ["A"], "對帳": ["A", "C"]}, "empty": "沒有標為未取得的項目。"}]}

cfg = {"company": "Keyence（キーエンス）", "ticker": "6861.T", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "AI 帶動的設備投資上行能否讓 Keyence 的成長與利潤率支持高於現價的目標價；實體 AI（機器人工作站）在營收中占多少",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：設備投資上行延續、日圓偏弱）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：半導體與電子投資回落、日圓升值）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, physai, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

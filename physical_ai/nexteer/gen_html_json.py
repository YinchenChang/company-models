# 產生 physical_ai/nexteer/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/nexteer/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/nexteer/html.json
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
def iv(k, f=""): return f"@In!E{inr(k)}" + (f"|{f}" if f else "")
def il(k): return f"@In!B{inr(k)}"
def nourl(sh, coord, f=""):
    v = wb[sh][coord].value
    return "（出處含網址，見 xlsx）" if isinstance(v, str) and "http" in v else f"@{sh}!{coord}" + (f"|{f}" if f else "")

LAB = dict(cnm="中國 SBW 市場（線模型；億 RMB）", cns="耐世特中國 SBW 有效份額", sbwap="亞太：SBW 營收（中國；子列）", canap="亞太：SBW 替代自家 EPS 抵減（子列）",
           convap="亞太：傳統轉向與傳動軸（子列）", ap="亞太（Asia Pacific）營收", os="海外 SBW 營收合計（北美＋EMEASA；說明列）", sbwna="北美：SBW 營收（子列）",
           canna="北美：SBW 替代自家 EPS 抵減（子列）", convna="北美：傳統轉向與傳動軸（子列）", na="北美（North America）營收", sbwem="EMEASA：SBW 營收（子列）",
           canem="EMEASA：SBW 替代自家 EPS 抵減（子列）", convem="EMEASA：傳統轉向與傳動軸（子列）", em="EMEASA（歐洲、中東、非洲及南美）營收", ot="其他（Others）營收",
           rev="營收（合併）", chk1="檢查：子列合計 − 分部（三地區合計，應為 0）", chk2="檢查：分部合計 − 合併營收（應為 0）", gr="營收年增率", sbw="SBW 營收合計（三地區）",
           sbwsh="SBW 占營收", ebna="北美調整後 EBITDA", ebap="亞太調整後 EBITDA", ebem="EMEASA 調整後 EBITDA", ebot="其他與未分配調整後 EBITDA", ebitda="調整後 EBITDA",
           ebm="調整後 EBITDA 率", da="折舊攤銷", adj="減損與其他調整", cogs="營業成本", gp="毛利", gpm="毛利率", opex="營業費用（淨額：工程研發＋銷管 − 其他收益）",
           op="營業利益", opm="營業利益率", nonop="業外（財務收支淨額＋合資損益）", pbt="稅前利益", tax="所得稅", pat="稅後淨利", nci="非控股權益損益", ni="歸屬母公司淨利",
           nim="淨利率（歸屬母公司）", nig="淨利年增率", sh="平均股數（百萬股）", eps="EPS（美元，基本＝稀釋）", epshk="EPS（港元，換算）", cx="資本支出", dwc="營運資金增加",
           fcff="FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", div="股利（母公司，現金支付年）", fcfad="自由現金流扣股利（收支關卡用）")
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm)
    PR[s] = {k: find(P, t, h) for k, t in LAB.items()}; PR[s]["hdr"] = h + 1
B = PR["B"]
STD = list("CDEFGHIJKL")   # FY2023 … FY2032
FC = list("FGHIJKL")       # FY2026 …
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VT = dict(dcfr="DCF 情境價值推到 12 個月後（港元）", cu="Comparable（未調整）：同業前瞻 P/E 中位數 × FY2028 EPS × USD/HKD", ca="Comparable（相對倍數調整後，進加權）",
          tps="各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", tp="加權目標價（港元）", dcf="DCF 目標價（情境加權，港元）", cmp="Comparable 目標價（情境加權，進加權，港元）",
          cmpu="Comparable 目標價（未調整，港元）", px="現價（In 頁 px，港元）", up="上檔空間", rt="估值評等（只看上檔空間）",
          gate="收支關卡：FY2026–FY2030 自由現金流扣股利合計（US$M；情境 A／B／C）", gj="入不敷出判定", rec="投資建議", dist="距最近級距邊界（百分點）",
          bnd="翻轉目標價：升一級／降一級（港元）", a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻",
          pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC", pe="隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 EPS（港元）",
          pe2="現行前瞻 P/E（資料商 NTM，換算至 2026-11-20）÷ 同業中位數", evx="現行 EV ÷ FY2026 調整後 EBITDA（情境 B；顯示用）", sbws="SBW 營收占比 FY2030（情境 B；顯示用）")
VR = {k: find(V, t) for k, t in VT.items()}
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")

DB = find(D, "情境 B（基準）")
DK = {k: find(D, t) for k, t in dict(ke="ke（股權成本）", e="股權市值 E（US$M）", dd="負債 D（借款＋租賃）", kd="稅後 kd", wacc="WACC", g="永續成長率 g＝MIN(g_cap, rf)", ronic="RONIC",
                                       nc="淨現金（2026-06-30）＝現金 − 借款 − 租賃 − 7 月已付股息 − 非控股權益").items()}
DBR = {k: find(D, t, DB) for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", eq="股權價值（US$M）", psv="每股價值（評價日，港元）",
                                          d12="DCF 情境價值推到 12 個月後（港元；×(1＋ke)^t − 期間股利）", tvsh="終值占 EV",
                                          alt="終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（港元，推到 12 個月後）").items()}
DA, DC = find(D, "情境 A（樂觀）"), find(D, "情境 C（保守）")
def dsc(k):
    t = D.cell(DBR[k], 1).value
    return [f"@DCF!B{find(D, t, DA)}", f"@DCF!B{DBR[k]}", f"@DCF!B{find(D, t, DC)}"]

# ---------------- 敏感度 ----------------
s0 = find(SE, "驅動") + 1
s_rows = [r for r in range(s0, SE.max_row + 1) if isinstance(SE.cell(r, 4).value, (int, float)) and isinstance(SE.cell(r, 5).value, (int, float))]
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
    return dict(TOR, title=f"加權目標價相對基準：輸入取{side}值（沒有變動達兩位數百分比的案例；列出影響最大的三項）",
                rows=",".join(map(str, sorted(rs, key=lambda r: -abs(SE.cell(r, 5).value))[:3])))
sens_blocks = [main_tor("低", lo), main_tor("高", hi),
               {"type": "note", "parts": ["@Sens!A1"]},
               {"type": "kv", "items": [["@Sens!A3", "@Sens!D3"], ["@Sens!A7", "@Sens!D7"], ["@Sens!A6", "@Sens!D6|n2"]]}]
if oth:
    sens_blocks.append(dict(TOR, title="加權目標價相對基準：只看單一情境、整組拿掉線控轉向、終值替代法", rows=",".join(map(str, oth))))
for side, rs in (("低", lo), ("高", hi)):
    shown = set(map(int, main_tor(side, rs)["rows"].split(",")))
    rest = [r for r in rs if r not in shown]
    if rest:
        sens_blocks.append(dict(TOR, title=f"加權目標價相對基準：輸入取{side}值，其餘案例", rows=",".join(map(str, rest)), collapse=True,
                                summary=f"展開：輸入取{side}值、變動未達兩位數百分比的案例"))
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：全部案例（加權目標價、DCF 目標價與估值評等）",
                    "head": [f"@Sens!{c}{s0 - 1}" for c in "ABDEFGH"],
                    "rowgen": {"sheet": "Sens", "rows": f"{s0}-{s_end}", "cols": ["A", "B", "D", "E", "F", "G", "H"], "fmts": ["", "", "n2", "sp1", "n2", "sp1", ""]}})
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}

hero = {"rec": vv("rec"),
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n2"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n2"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n2"), ""],
                  [il("px"), iv("px", "n2"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等", vv("rt"), "（只看上檔空間）；收支關卡", vv("gj"), "（情境 B 自由現金流扣股利，FY2026–FY2030 合計 ", vv("gate", "C", "n0"),
                 " 百萬美元）。現價對應的企業價值只有調整後 EBITDA 的 ", vv("evx", "B", "x1"), "；上檔主要來自傳統轉向現金流與淨現金，線控轉向是加分項。"]}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "營收依官方地區分部堆疊（百萬美元）與線控轉向占營收：FY2023–FY2032（情境 B）", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k), "color": c} for k, c in (("na", "#0b7a80"), ("ap", "#6fa8dc"), ("em", "#c29a2b"), ("ot", "#8a8c8c"))],
     "lines": [{"name": [f"@PL!A{B['sbwsh']}"], "y": pl("sbwsh"), "axis": "r", "color": "#b4472f", "labels": True},
               {"name": [f"@PL!A{B['ebm']}"], "y": pl("ebm"), "axis": "r", "color": "#3b3838", "dash": True}],
     "caption": "堆疊＝三個官方地區分部與其他（合計等於合併營收；各分部含線控轉向子列）。實線＝線控轉向占營收、虛線＝調整後 EBITDA 率（右軸）。FY2023–FY2025 為實績，FY2026 為上半年實績加下半年外推，其後為模型。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["承重變數：北美與亞太傳統業務的 EBITDA 率（近四季 ", iv("m_na", "p1"), "、", iv("m_ap", "p1"), "）、線控轉向的 EBITDA 率（", iv("m_sbw", "p0"),
         "）與耐世特中國線控轉向份額（情境 B 到 FY2030 為 ", iv("shr30_B", "p0"), "）。"]]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n2"), " 港元（DCF ", vv("dcf", "B", "n2"), "、Comparable ", vv("cmp", "B", "n2"), "），對現價 ", iv("px", "n2"), " 港元上檔 ", vv("up", "B", "sp1"),
         "，投資建議", vv("rec"), "；賣方平均目標價 ", iv("cons_tp", "n2"), " 港元。"],
        ["便宜的原因：現行前瞻本益比只有同業中位數的 ", vv("pe2", "B", "p0"), "，企業價值對 EBITDA 約 ", vv("evx", "B", "x1"), "；控股股東背景、最大客戶集中與港股流動性是常見的折價理由，模型未對此另行調整（相對倍數係數維持一）。"],
        ["線控轉向：情境 B 到 FY2030 線控轉向營收 ", f"@PL!J{B['sbw']}|n0", " 百萬美元、占營收 ", vv("sbws", "B", "p1"), "；整組拿掉線控轉向，加權目標價 ",
         f"@Sens!D{find(SE, 'SBW 全無', exact=False)}|n2", " 港元（", f"@Sens!E{find(SE, 'SBW 全無', exact=False)}|sp1", "）。"],
        ["評等邊界：", "@README!A5"],
        "@README!A6"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A7"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（港元；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}, {"label": il("cons_tp"), "value": iv("cons_tp"), "series": 1}],
     "ref": {"label": il("px"), "value": iv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n2"), ("cmp", "n2"), ("cmpu", "n2"), ("tp", "n2"), ("px", "n2"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("dist", "n1"), ("bnd", ""))]
             + [[il("cons_tp"), iv("cons_tp", "n2")]]},
    {"type": "table", "caption": "情境目標價（港元）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", "n2"]}},
    {"type": "table", "caption": "收支關卡（百萬美元）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3"],
     "rows": [[f"@Val!A{VR['gate']}", vv("gate", "B", "n0"), vv("gate", "C", "n0"), vv("gate", "D", "n0")]]},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（百萬美元；每股為港元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n0"), ("eq", "n0"), ("psv", "n2"), ("d12", "n2"), ("tvsh", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n0"), ("tv", "n0"), ("pvtv", "n0"), ("alt", "n2"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n3"), ("erp", "p1"), ("crp", "p1"), ("sh_v", "n1"), ("fx_hkd", "n4"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n1", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n0", "", ""], [f"@DCF!A{DK['dd']}", f"@DCF!B{DK['dd']}|n1", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因與隱含倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("pd2", "n2"), ("pe", "x1"), ("pe2", "n2"), ("evx", "x1"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻本益比與中位數", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!I3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag", nourl("In", f"I{inr(k)}")] for k in ("pe_mobis", "pe_denso", "pe_nsk", "pe_tuopu", "pe_jtekt", "pe_med", "pe_self")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"), ("pay_r", "p0"))]}]}

rvh = find(RC, "反解目標")
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價（上檔為零）及落在評等邊界所需的值，其他輸入不變。「無解」表示在搜尋區間內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{rvh - 1}", "head": [f"@Recon!{c}{rvh}" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{rvh + 1}-{RC.max_row}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "", "p1", "p1", ""]}}]}

def prow(k, f): return [f"@PL!A{B[k]}"] + [f"@PL!{c}{B[k]}|{f}" for c in STD]
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "table", "caption": "標準損益表（IFRS，即法定財報 GAAP 口徑；百萬美元，EPS 為美元）：FY2023–FY2025 實績、FY2026 上半年實績加下半年外推、FY2027 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [prow("rev", "n0"), prow("gr", "p1"), prow("cogs", "n0"), prow("gp", "n0"), prow("gpm", "p1"), prow("opex", "n0"), prow("op", "n0"), prow("opm", "p1"),
              prow("nonop", "n1"), prow("pbt", "n0"), prow("tax", "n0"), prow("pat", "n0"), prow("nci", "n1"), prow("ni", "n0"), prow("nim", "p1"), prow("eps", "n4"),
              prow("epshk", "n3"), prow("nig", "p1")]},
    {"type": "note", "parts": ["會計口徑：IFRS（公司法定財報準則；本表為 GAAP 口徑，未做 Non-GAAP 調整）。公司另揭露調整後 EBITDA（非 IFRS 指標，見驅動因子章）。預測期營業利益由分部調整後 EBITDA 減折舊攤銷與減損推得；營業成本、毛利、營業費用以近期營業費用淨額占營收比拆出，不影響營業利益。"]},
    {"type": "theme", "title": "營收（百萬美元）× 情境 A／B／C 與營業利益率（情境 B）：FY2023–FY2032", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p1", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：營收"], "y": pl("rev", s), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['opm']}", "（情境 B）"], "y": pl("opm"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境營收；折線＝情境 B 營業利益率（右軸）。實績與 FY2026 三情境相同的部分只差線控轉向的市場讀值，分化主要從 FY2027 開始。"},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "營收（百萬美元）", "fmt": "n0", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "調整後 EBITDA（百萬美元）", "fmt": "n0", "chart": True, "rows": {s: pl("ebitda", s, "F") for s in "ABC"}},
                 {"name": "營業利益（百萬美元）", "fmt": "n0", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "EPS（美元）", "fmt": "n4", "chart": True, "rows": {s: pl("eps", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣股利（百萬美元；收支關卡用）", "fmt": "n0", "rows": {s: pl("fcfad", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流（百萬美元）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['cx']}-{B['fcfad']}", "cols": ["A"] + STD, "fmts": [""] + ["n1"] * len(STD)}}]}

drivers = {"id": "drivers", "title": "驅動因子：線控轉向 → 地區分部 → 合併營收", "blocks": [
    {"type": "theme", "title": "線控轉向營收依地區堆疊（百萬美元）與中國份額：FY2026–FY2032（情境 B）", "x": f"@PL!F{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("sbw", "B", "F"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "F"), "color": c} for k, c in (("sbwap", "#0b7a80"), ("sbwna", "#6fc7cb"), ("sbwem", "#c29a2b"))],
     "lines": [{"name": [f"@PL!A{B['cns']}"], "y": pl("cns", "B", "F"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝三地區線控轉向子列（合計＝線控轉向營收）；折線＝耐世特中國線控轉向有效份額（右軸）。中國市場讀自智慧車線模型，份額與海外營收為本模型假設。"},
    {"type": "note", "parts": ["@PL!A1"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("lmk26", "n1"), ("lmk30", "n1"), ("lmk32", "n1"), ("shr30", "p0"), ("os30", "n0"), ("gna", "p1"), ("gap", "p1"), ("gem", "p1"), ("madj", "p1"))]},
    {"type": "table", "caption": "分部營收推導（情境 B；百萬美元）：線控轉向子列＋抵減＋傳統業務 → 地區分部 → 合併營收；核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['cnm']}-{B['sbwsh']}", "cols": ["A"] + STD, "fmts": [""] + ["n0"] * len(STD),
                "row_fmt": {str(B[k]): {c: ("p1" if k != "cnm" else "n1") for c in STD} for k in ("cnm", "cns", "gr", "sbwsh")}}},
    {"type": "table", "caption": "分部調整後 EBITDA 到營業利益（情境 B；百萬美元）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['ebna']}-{B['adj']}", "cols": ["A"] + STD, "fmts": [""] + ["n0"] * len(STD), "row_fmt": {str(B["ebm"]): {c: "p1" for c in STD}}}},
    {"type": "h3", "text": "一、線控轉向（命題模組）"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("lmk30_B", "n1"), ("lmk32_B", "n1"), ("sbw26", "p0"), ("shr30_B", "p0"), ("cann", "p0"), ("os26", "n0"), ("os30_B", "n0"), ("os_na", "p0"), ("os_g", "p0"), ("m_sbw", "p0"))]},
    {"type": "h3", "text": "二、地區分部的傳統業務"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("g_anc", "p0"), ("gna_B", "p1"), ("gap_B", "p1"), ("gem_B", "p1"), ("m_na", "p1"), ("m_ap", "p1"), ("m_em", "p1"), ("madj_B", "p1"), ("eb_oth", "n0"))]},
    {"type": "h3", "text": "三、全公司（損益下半段、現金流、折現）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("da_r", "p1"), ("adj_x", "n0"), ("nonop", "n0"), ("tax_w", "p0"), ("nci_r", "p0"), ("pay_r", "p0"), ("fc_cx", "p1"), ("fc_wc", "p0"),
               ("beta", "n3"), ("erp", "p1"), ("crp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("w_dcf", "p0"), ("rel_k", "n2"))]},
    {"type": "table", "collapse": True, "summary": "展開：固定輸入（匯率、股數、資產負債表）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag"] for k, f in (("fx_cny", "n4"), ("fx_hkd", "n4"), ("sh_v", "n1"), ("cash_v", "n1"), ("debt_v", "n1"), ("lease_v", "n1"),
                                                                     ("mi_v", "n1"), ("div_jul", "n1"), ("opex_r", "p2"), ("oth_k", "n1"))]}]}

RE = wb["對帳"]
recon = {"id": "recon", "title": "對帳：市場共識、公司揭露與線層級變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, '加權目標價 vs 賣方平均目標價（港元）')}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n3", "n3", "n3", ""]}},
    {"type": "table", "caption": "@對帳!A1", "head": [f"@對帳!{c}2" for c in "ABCDEFH"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{RE.max_row}", "cols": ["A", "B", "C", "D", "E", "F", "H"], "fmts": ["", "", "n3", "n3", "n3", "", ""]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", "@Sens!D3"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價、匯率與折現率的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!I3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag", nourl("In", f"I{inr(k)}")] for k, f in (("px", "n2"), ("px_prev", "n3"), ("fx_hkd", "n4"), ("fx_cny", "n4"), ("rf", "p2"), ("pe_self", "n2"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"In": ["B"], "Hist": ["A"], "對帳": ["A"]}, "empty": "沒有標為未取得的項目。"}]}

cfg = {"company": "耐世特 Nexteer", "ticker": "1316.HK", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "以電動轉向為主的耐世特，在線控轉向放量前，傳統業務現金流與淨現金能否支持高於現價的目標價；線控轉向對目標價的貢獻有多大",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：線控轉向市場取線模型高值、份額高）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：線控轉向市場低值、份額低、傳統業務零成長）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

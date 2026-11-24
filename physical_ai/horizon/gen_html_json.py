# 產生 physical_ai/horizon/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；模型改版移列後重跑本檔即可。設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/horizon/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/horizon/html.json
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

LAB = dict(retail="中國乘用車國內零售（百萬台；線模型）", pen="中國城市 NOA 滲透率（線模型）", cityc="中國城市 NOA 新車（百萬台）", scity="地平線城區 NOA 份額",
           hzcity="地平線城區 NOA 套數（百萬套）", othc="其他新車（非城區 NOA，百萬台）", soth="地平線其他 ADAS 份額", hzoth="地平線其他 ADAS 套數（百萬套）",
           units="征程出貨合計（百萬套）", pcity="城區 NOA 每套產品方案營收（RMB／套）", poth="其他 ADAS 每套產品方案營收（RMB／套）", prod="產品方案營收",
           lic_u="城市 NOA 每車授權費（USD；線模型）", roy="授權－城區 NOA 權利金（子列）", ols="授權－其他授權及服務（子列）", lic="授權及服務營收", na="非車載營收",
           rev="營收（合併）", chk="檢查：分部合計 − 合併營收（應為 0）", gr="營收年增率", cogs="營業成本", gp="毛利", gm="毛利率", gmp="產品方案毛利率",
           rnd="研發費用", sga="銷售及管理費用", imp="金融資產減值", opex="營業費用（研發＋銷售及管理＋減值）", oi="其他收入及利得", op="營業利益（IFRS，含其他收入及利得）",
           opm="營業利益率", fin="財務收益淨額", assoc="權益法投資損益", fv="優先股及金融負債公允價值變動", nonop="業外合計（其他收入淨額：財務＋權益法＋公允價值）",
           pbt="稅前利益", tax="所得稅", disc="停業單位損益（D-Robotics）", ni="稅後淨利（歸屬母公司）", nim="淨利率", nig="淨利年增率", sh="加權平均股數（百萬股）",
           eps="EPS（基本，人民幣）", da="折舊攤銷", cx="資本支出", dwc="營運資金增加", fcff="FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", div="股利",
           fcfad="自由現金流扣股利（收支關卡用）", nc="期末淨現金（現金 − 借款 − 可轉債 − 租賃；不含金融資產）")
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm)
    PR[s] = {k: find(P, t, h) for k, t in LAB.items()}; PR[s]["hdr"] = h + 1
B = PR["B"]
YC = "CDEFGHIJKL"   # PL：C＝FY2023 … L＝FY2032
STD = list(YC)
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VT = dict(dcfr="DCF 情境價值推到 12 個月後", cu="Comparable（未調整）：同業 EV／營收中位數 × FY2028 營收 ＋ FY2027 年底淨現金與投資",
          ca="Comparable（相對倍數調整後，進加權）", tps="各情境目標價（DCF 權重 × DCF ＋（1 − DCF 權重）× Comparable）", tp="加權目標價（港元）",
          dcf="DCF 目標價（情境加權，港元）", cmp="Comparable 目標價（情境加權，進加權；EV／營收，港元）", px="現價（In 頁 px，港元）", up="上檔空間",
          rt="估值評等（只看上檔空間）", gate="收支關卡：FY2026–FY2030 自由現金流扣股利合計（RMB bn；情境 A／B／C）", gj="入不敷出判定", rec="投資建議",
          dist="距最近級距邊界（百分點）", bnd="翻轉目標價：升一級／降一級（港元）", minc="期間流動性：情境 B 最低年末淨現金（FY2026–FY2030，RMB bn）",
          be="情境 B 營業利益轉正的第一個年度", beA="情境 A 營業利益轉正的第一個年度", beC="情境 C 營業利益轉正的第一個年度", a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻",
          pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC", evs="現行 EV／FY2027 營收（情境 B）：(市值 − 淨現金與投資) ÷ 營收",
          evr="現行 EV／FY2027 營收 ÷ 同業中位數", ievs="隱含 EV／FY2028 營收（加權目標價、情境 B）")
VR = {k: find(V, t) for k, t in VT.items()}
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")
def iv(k, f=""): return f"@In!E{inr(k)}" + (f"|{f}" if f else "")
def il(k): return f"@In!B{inr(k)}"
def nourl(sh, coord, f=""):
    v = wb[sh][coord].value
    return "（出處含網址，見 xlsx）" if isinstance(v, str) and "http" in v else f"@{sh}!{coord}" + (f"|{f}" if f else "")

DB = find(D, "情境 B（基準）")
DK = {k: find(D, t, exact=False) for k, t in dict(ke="ke（股權成本）", e="股權市值 E", dd="負債 D", kd="稅後 kd", wacc="WACC", g="永續成長率 g＝", ronic="RONIC",
                                                    nc="淨現金與投資（2026-06-30 擬制）").items()}
DBR = {k: find(D, t, DB, exact=False) for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", eq="股權價值", psv="每股價值（評價日",
                                                         d12="DCF 情境價值推到 12 個月後", tvsh="終值占 EV", opm32="FY2032 營業利益率", alt="終值替代法").items()}
DA, DC = find(D, "情境 A（樂觀）"), find(D, "情境 C（保守）")
def dsc(k):
    t = D.cell(DBR[k], 1).value
    return [f"@DCF!B{find(D, t, DA)}", f"@DCF!B{DBR[k]}", f"@DCF!B{find(D, t, DC)}"]

# Sens：表頭第 5 列；A 欄以「（低）」「（高）」結尾者為單一參數低／高案
s_rows = [r for r in range(6, SE.max_row + 1) if isinstance(SE.cell(r, 5).value, (int, float))]
lo = [r for r in s_rows if str(SE.cell(r, 1).value).endswith("（低）")]
hi = [r for r in s_rows if str(SE.cell(r, 1).value).endswith("（高）")]
oth = [r for r in s_rows if r not in lo and r not in hi]
s_end = max(s_rows)
onlyA = find(SE, "只看情境 A")
TOR = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": "E", "fmt": "sp1"}
def big(rs): return [r for r in rs if abs(SE.cell(r, 5).value) >= 0.1]
def small(rs): return [r for r in rs if abs(SE.cell(r, 5).value) < 0.1]
def main_tor(side, rs):
    if big(rs):
        return dict(TOR, title=f"加權目標價相對基準：輸入取{side}值（變動達兩位數百分比的案例）", rows=",".join(map(str, sorted(big(rs), key=lambda r: -abs(SE.cell(r, 5).value)))))
    return dict(TOR, title=f"加權目標價相對基準：輸入取{side}值（沒有變動達兩位數百分比的案例；列出影響最大者）", rows=",".join(map(str, sorted(rs, key=lambda r: -abs(SE.cell(r, 5).value))[:1])))
sens_blocks = [main_tor("低", lo), main_tor("高", hi),
               {"type": "note", "parts": ["@Sens!A1"]},
               {"type": "kv", "items": [["build_id", "@Sens!D2"], ["模型狀態", "@Sens!D3"]]}]
if oth:
    sens_blocks.append(dict(TOR, title="加權目標價相對基準：只看單一情境、終值替代法", rows=",".join(map(str, oth))))
for side, rs in (("低", lo), ("高", hi)):
    shown = set(map(int, main_tor(side, rs)["rows"].split(",")))
    rest = [r for r in small(rs) if r not in shown]
    if rest:
        sens_blocks.append(dict(TOR, title=f"加權目標價相對基準：輸入取{side}值，其餘案例", rows=",".join(map(str, rest)), collapse=True,
                                summary=f"展開：輸入取{side}值、變動未達兩位數百分比的案例"))
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：全部案例（加權目標價、DCF 目標價與估值評等）",
                    "head": ["@Sens!A5", "@Sens!B5", "@Sens!D5", "@Sens!E5", "@Sens!F5", "@Sens!G5", "@Sens!H5"],
                    "rowgen": {"sheet": "Sens", "rows": f"6-{s_end}", "cols": ["A", "B", "D", "E", "F", "G", "H"], "fmts": ["", "", "n2", "sp1", "n2", "sp1", ""]}})
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}

hero = {"rec": vv("rec"),
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n2"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n2"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n2"), ""],
                  [il("px"), iv("px", "n2"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等", vv("rt"), "；收支關卡", vv("gj"), "（情境 B 自由現金流扣股利 FY2026–FY2030 合計 ", vv("gate", "C", "n1"),
                 " 十億人民幣）。營業利益轉正年度：情境 A ", vv("beA"), "、B ", vv("be"), "、C ", vv("beC"), "。只看樂觀情境時目標價 ", f"@Sens!D{onlyA}|n2",
                 " 港元（", f"@Sens!H{onlyA}", "）。成長路徑多為假設。"]}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "營收依營收線堆疊（十億人民幣）與營業利益率：FY2024–FY2032（情境 B）", "x": f"@PL!D{B['hdr']}:L{B['hdr']}", "fmt": "n1", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev", "B", "D"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "D"), "color": c} for k, c in (("prod", "#6fa8dc"), ("lic", "#0b7a80"), ("na", "#8a8c8c"))],
     "lines": [{"name": [f"@PL!A{B['opm']}"], "y": pl("opm", "B", "D"), "axis": "r", "color": "#b4472f", "labels": True}],
     "caption": "堆疊＝產品方案、授權及服務、非車載三條營收線（合計等於營收）；折線＝營業利益率（右軸）。FY2024–FY2025 實績、FY2026 為上半年實績加下半年假設、FY2027 起為模型。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["承重變數：城區 NOA 份額（2026 年 ", iv("s_city26", "p1"), "，情境 B 2030 年 ", iv("s_city30_B", "p0"), "）、其他授權及服務成長（情境 B FY2028–FY2030 每年 ",
         iv("g_ls28_B", "p0"), "）、研發費用成長（FY2028 起每年 ", iv("g_rndlt", "p0"), "）。"]]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n2"), " 港元（DCF ", vv("dcf", "B", "n2"), "、Comparable ", vv("cmp", "B", "n2"), "），對現價 ", iv("px", "n2"), " 港元上檔 ",
         vv("up", "B", "sp1"), "，估值評等", vv("rt"), "、投資建議", vv("rec"), "。"],
        ["虧損未收斂：上半年營收 ", f"@半年!E6|n2", " 十億人民幣（年增 ", f"@半年!E11|p1", "），研發費用 ", iv("h1_rnd", "n2"), "，經營虧損仍大；情境 B 營業利益 ", vv("be"),
         "，期間最低年末淨現金 ", vv("minc", "B", "n1"), " 十億人民幣（不含金融資產）。"],
        ["市場已給高倍數：現行企業價值為情境 B FY2027 營收的 ", vv("evs", "B", "n1"), " 倍，是硬體為主同業中位數的 ", vv("evr", "B", "n1"),
         " 倍；授權占營收過半、毛利率高，部分差距是商業模式差。"],
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
    {"type": "note", "parts": ["情境 B、C 的企業價值為負（明確期與終值期都是虧損），股權價值以零計，DCF 取零；DCF 目標價只來自情境 A。"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n2"), ("cmp", "n2"), ("tp", "n2"), ("px", "n2"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("dist", "n1"), ("bnd", ""), ("beA", ""), ("be", ""), ("beC", ""), ("minc", "n1"))]
             + [[il("cons_tp"), iv("cons_tp", "n2")]]},
    {"type": "table", "caption": "情境目標價（港元）與收支關卡（十億人民幣）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", "n2"]}},
    {"type": "table", "caption": "收支關卡（情境 A／B／C）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@Val!A{VR['gate']}", f"@Val!B{VR['gate']}|n1", f"@Val!C{VR['gate']}|n1", f"@Val!D{VR['gate']}|n1"]]},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（十億人民幣；每股為港元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n1"), ("eq", "n1"), ("psv", "n2"), ("d12", "n2"), ("opm32", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n1"), ("tv", "n1"), ("pvtv", "n1"), ("tvsh", "p1"), ("alt", "n2"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n3"), ("erp", "p1"), ("crp", "p2"), ("sh_v", "n1"), ("fx", "n4"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n2", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n1", "", ""], [f"@DCF!A{DK['dd']}", f"@DCF!B{DK['dd']}|n2", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因與隱含倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("pd2", "n2"), ("evs", "n2"), ("evr", "n2"), ("ievs", "n2"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業 EV／營收與中位數", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!I3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag", nourl("In", f"I{inr(k)}")] for k in ("ev_mbly", "ev_amba", "ev_hsai", "ev_rs", "ev_med")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"))]}]}

rv0 = find(RC, "反解目標")
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價（上檔為零）及落在評等邊界所需的值，其他輸入不變。「無解」表示在搜尋區間內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{rv0 - 1}", "head": [f"@Recon!{c}{rv0}" for c in "ACDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{rv0 + 1}-{RC.max_row}", "cols": ["A", "C", "D", "E"], "fmts": ["", "p1", "p1", ""]}}]}

def prow(k, f):
    return [f"@PL!A{B[k]}"] + [f"@PL!{c}{B[k]}|{f}" for c in STD]
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "營收（十億人民幣）× 情境 A／B／C 與營業利益率（情境 B）：FY2026–FY2032", "x": f"@PL!F{B['hdr']}:L{B['hdr']}", "fmt": "n1", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：營收"], "y": pl("rev", s, "F"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['opm']}", "（情境 B）"], "y": pl("opm", "B", "F"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境營收；折線＝情境 B 營業利益率（右軸）。FY2026 三情境相同（上半年實績加下半年假設），分化從 FY2027 開始。"},
    {"type": "note", "parts": ["會計口徑：IFRS（公司法定財報，即 GAAP 口徑）；市場常用的 Non-GAAP 經調整淨虧損（加回股份支付與公允價值變動）列在 xlsx 歷史頁，本表不採用。FY2023–FY2025 為申報口徑（含非車載）；FY2026 起為持續經營（車載），淨利含上半年一次性公允價值利得與停業單位利得。FY2023 取自資料商。"]},
    {"type": "table", "caption": "標準損益表（IFRS／GAAP 口徑；十億人民幣，EPS 為人民幣）：FY2023–FY2025 實績、FY2026 上半年實績加下半年假設、FY2027 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [prow("rev", "n2"), prow("gr", "p1"), prow("cogs", "n2"), prow("gp", "n2"), prow("gm", "p1"), prow("rnd", "n2"), prow("sga", "n2"), prow("opex", "n2"),
              prow("oi", "n2"), prow("op", "n2"), prow("opm", "p1"), prow("nonop", "n2"), prow("pbt", "n2"), prow("tax", "n2"), prow("disc", "n2"),
              prow("ni", "n2"), prow("nim", "p1"), prow("sh", "n0"), prow("eps", "n3"), prow("nig", "p1")]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "營收（十億人民幣）", "fmt": "n2", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "營業利益（十億人民幣）", "fmt": "n2", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "營業利益率", "fmt": "p1", "rows": {s: pl("opm", s, "F") for s in "ABC"}},
                 {"name": "EPS（人民幣）", "fmt": "n3", "chart": True, "rows": {s: pl("eps", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣股利（十億人民幣；收支關卡用）", "fmt": "n2", "rows": {s: pl("fcfad", s, "F") for s in "ABC"}},
                 {"name": "期末淨現金（十億人民幣）", "fmt": "n2", "rows": {s: pl("nc", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流（十億人民幣）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in "FGHIJKL"],
     "rowgen": {"sheet": "PL", "rows": f"{B['da']}-{B['nc']}", "cols": ["A"] + list("FGHIJKL"), "fmts": [""] + ["n2"] * 7}}]}

drivers = {"id": "drivers", "title": "驅動因子：征程套數 × 每套營收、授權 → 營收", "blocks": [
    {"type": "theme", "title": "征程出貨依用途堆疊（百萬套）與城區 NOA 每套產品方案營收：FY2026–FY2032（情境 B）", "x": f"@PL!F{B['hdr']}:L{B['hdr']}", "fmt": "n2", "fmt_r": "n0", "stack": True, "pad": 1.3,
     "total": pl("units", "B", "F"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "F"), "color": c} for k, c in (("hzcity", "#0b7a80"), ("hzoth", "#6fc7cb"))],
     "lines": [{"name": [f"@PL!A{B['pcity']}"], "y": pl("pcity", "B", "F"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝城區 NOA 套數（線模型城市 NOA 新車 × 地平線份額）與其他 ADAS 套數（合計＝征程出貨）；折線＝城區 NOA 每套產品方案營收（右軸，人民幣）。市場量讀智慧車線模型，份額與每套營收為本模型假設。"},
    {"type": "note", "parts": ["@PL!A1"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("s_city30", "p0"), ("d_soth", "p1"), ("mix", "p0"), ("g_ls27", "p0"), ("g_ls28", "p0"), ("g_lslt", "p0"), ("gm_p", "p0"))]},
    {"type": "table", "caption": "營收推導（情境 B）：市場量 → 套數 → 產品方案；權利金＋其他授權 → 授權及服務；核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in "FGHIJKL"],
     "rowgen": {"sheet": "PL", "rows": f"{B['retail']}-{B['gr']}", "cols": ["A"] + list("FGHIJKL"), "fmts": [""] + ["n2"] * 7,
                "row_fmt": {str(B[k]): {c: ("p1" if k in ("pen", "scity", "soth", "gr") else "n0") for c in "FGHIJKL"} for k in ("pen", "scity", "soth", "gr", "pcity", "poth", "lic_u")}}},
    {"type": "table", "caption": "毛利與費用推導（情境 B；十億人民幣）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in "FGHIJKL"],
     "rowgen": {"sheet": "PL", "rows": f"{B['cogs']}-{B['opm']}", "cols": ["A"] + list("FGHIJKL"), "fmts": [""] + ["n2"] * 7,
                "row_fmt": {str(B[k]): {c: "p1" for c in "FGHIJKL"} for k in ("gm", "gmp", "opm")}}},
    {"type": "h3", "text": "一、產品方案（征程晶片＋參考設計；命題模組）"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("k2_p", "n2"), ("s_city30_B", "p0"), ("d_soth_B", "p1"), ("p_oth", "n0"), ("mix_B", "p0"), ("byd_k", "p0"), ("gm_p_B", "p0"))]},
    {"type": "h3", "text": "二、授權及服務"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("k2_l", "n2"), ("g_ls27_B", "p0"), ("g_ls28_B", "p0"), ("g_lslt_B", "p0"), ("gm_l", "p1"))]},
    {"type": "h3", "text": "三、全公司（費用、業外、現金流、折現）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("k2_x", "n2"), ("g_rnd27", "p0"), ("g_rndlt", "p0"), ("g_sga", "p0"), ("oi", "n2"), ("assoc27", "n2"), ("y_cash", "p1"), ("tax_w", "p0"), ("fc_da", "p0"),
               ("fc_cx", "p0"), ("fc_wc", "p0"), ("sh_rsu", "n0"), ("inv_k", "p0"), ("ev_med", "n2"), ("beta", "n3"), ("erp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("w_dcf", "p0"))]},
    {"type": "table", "collapse": True, "summary": "展開：上半年實績與錨點（固定輸入）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, "n3"), f"@In!G{inr(k)}|tag"] for k in ("h1_p", "h1_l", "h1_cogs", "gm_p_h1", "h1_rnd", "h1_sga", "h1_imp", "h1_oi", "h1_fin", "h1_assoc", "h1_fv", "h1_tax",
                                                                  "h1_disc", "h1_ni", "h1_units", "h1_cityn", "h1_retail", "s_city26", "cont25")]}]}

RE = wb["對帳"]
recon = {"id": "recon", "title": "對帳：公司揭露、賣方、營收線與線層級變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, '加權目標價 vs 賣方平均目標價（港元）')}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""]}},
    {"type": "table", "caption": "營收線半年對帳（十億人民幣；核對列應為零）", "head": [f"@半年!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "半年", "rows": "3-8", "cols": list("ABCDE"), "fmts": [""] + ["n3"] * 4}},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數（智慧車線模型取數）", "head": [f"@對帳!{c}2" for c in "BCDEG"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{2 + sum(1 for r in range(3, RE.max_row + 1) if RE.cell(r, 6).value and str(RE.cell(r, 6).value).startswith(('下游取數', '輸入')))}", "cols": ["B", "C", "D", "E", "G"], "fmts": ["", "", "n3", "", "tag"]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", "@Sens!D2"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價、匯率與折現輸入的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}")] for k, f in (("px", "n2"), ("cny", "n4"), ("hkd", "n4"), ("rf", "p2"), ("cons_tp", "n2"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"In": ["B"], "Hist": ["A", "B"]}, "empty": "沒有標為未取得的項目。"}]}

cfg = {"company": "地平線 Horizon Robotics", "ticker": "9660.HK", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "征程晶片在中國城區 NOA 的份額與每套價值、以及高毛利授權的成長，能否在研發費用仍上升下讓虧損收斂，支持現價",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：份額上升、授權高成長）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：車廠自研、價格戰、車市下滑）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

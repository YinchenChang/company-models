# 產生 physical_ai/robosense/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/robosense/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/robosense/html.json
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

LAB = dict(mkt="命題模組：中國城市 NOA 新車數（萬台；線 2 下游取數）", pen="中國城市 NOA 滲透率（線 2 下游取數）", kf="每台 NOA 車對應速騰 ADAS 出貨（顆／台；含搭載率、每車顆數、份額）",
           au="ADAS LiDAR 出貨（萬顆）", asp_a="ADAS LiDAR 平均單價（人民幣元／顆）", ru="機器人及其他 LiDAR 出貨（萬顆）", asp_r="機器人等 LiDAR 平均單價（人民幣元／顆）",
           adas="ADAS LiDAR 營收", rob="機器人等 LiDAR 營收", oth="解決方案、服務與其他營收", rev="營業收入（合併）".replace("營業收入","營收"), chk="檢查：分部合計 − 營收（應為 0）", gr="營收年增率",
           cogs="營業成本", gp="毛利", gm="毛利率", rd="研發費用", sm="銷售及營銷費用", ga="行政費用", ox="營業費用（研發＋銷售＋行政）", op="營業利益（本模型口徑：毛利 − 營業費用）", opm="營業利益率",
           oi="其他收入淨額（業外：其他收入、其他損益、財務、聯營、減值；FY2023 含優先股公允價值變動）", pbt="稅前利益", tax="所得稅", ni="稅後淨利", nm="淨利率", nig="淨利年增率",
           sh="平均股數（百萬股）", eps="EPS（基本，人民幣元）", da="折舊攤銷", cx="資本支出", dwc="營運資金增加", fcff="FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）",
           div="股利", fcfad="自由現金流扣股利（收支關卡用）")
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm); PR[s] = {k: find(P, t, h) for k, t in LAB.items()}; PR[s]["hdr"] = h + 1
B = PR["B"]
STD = list("CDEFGHIJKL")   # FY2023 … FY2032
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VL = dict(dcfr="DCF 情境價值推到 12 個月後（港元）", cu="Comparable（未調整）：同業 EV ÷ 營收中位數 × FY2027 營收 ＋ 淨現金（港元）", ca="Comparable（相對倍數調整後，進加權）",
          tps="各情境目標價（DCF 權重 × DCF ＋ (1 − DCF 權重) × Comparable）", tp="加權目標價（港元）", dcf="DCF 目標價（情境加權，港元）", cmp="Comparable 目標價（EV ÷ 營收，情境加權，進加權，港元）",
          cmpu="Comparable 目標價（未調整，情境加權，港元）", px="現價（In 頁 px，港元）", up="上檔空間", rt="估值評等（只看上檔空間）",
          gate="收支關卡：FY2026–FY2030 自由現金流扣股利合計（百萬人民幣；情境 A／B／C）", gj="入不敷出判定", rec="投資建議",
          lq="期間流動性：情境 B 最低年末淨現金（FY2026–FY2030，百萬人民幣；不含新增舉債）", dist="距最近級距邊界（百分點）", bnd="翻轉目標價：升一級／降一級（港元）",
          a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻", pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC",
          evs="現行 EV ÷ 近四季營收：速騰聚創", evm="同業 EV ÷ 近四季營收中位數（4 家）", rel="現行相對倍數（速騰聚創 ÷ 同業中位數）",
          sg1="子群中位數（純展示）：LiDAR（禾賽、Ouster）", sg2="子群中位數（純展示）：感知晶片與系統（Mobileye、地平線）",
          ipe="隱含 EV ÷ FY2027 營收（加權目標價，情境 B 營收）", pe="隱含 P/E：加權目標價 ÷ 情境 B FY2028 EPS（港元換算）")
VR = {k: find(V, t) for k, t in VL.items()}
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")
def iv(k, f=""): return f"@In!E{inr(k)}" + (f"|{f}" if f else "")
def il(k): return f"@In!B{inr(k)}"
def nourl(sh, coord, f=""):
    v = wb[sh][coord].value
    return "（出處含網址，見 xlsx）" if isinstance(v, str) and "http" in v else f"@{sh}!{coord}" + (f"|{f}" if f else "")

DK = {k: find(D, t) for k, t in dict(ke="ke（股權成本）", e="股權市值 E（百萬人民幣）", dd="負債 D（借款＋租賃）", kd="稅後 kd", wacc="WACC", g="永續成長率 g（上限 3% 與 rf 取小）", ronic="RONIC",
                                       nc="淨現金（2026-06-30）＝現金 − 借款 − 租賃 − 6-30 後買回").items()}
DB = find(D, "情境 B（基準）")
DBR = {k: find(D, t, DB) for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", stub="＋Stub 期 FCFF（2026 下半年＝FY2026 模型 FCFF − 上半年實績）",
                                          eq="股權價值（百萬人民幣）", psv="每股價值（評價日，人民幣）", d12="DCF 情境價值推到 12 個月後（港元；×(1＋ke)^t − 期間股利）", tvsh="終值占 EV",
                                          alt="終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（港元，推到 12 個月後）").items()}
DA, DC = find(D, "情境 A（樂觀）"), find(D, "情境 C（保守）")
def dsc(k):
    t = D.cell(DBR[k], 1).value
    return [f"@DCF!B{find(D, t, DA)}", f"@DCF!B{DBR[k]}", f"@DCF!B{find(D, t, DC)}"]

# Sens：表頭第 7 列、案例第 8 列起；E 欄加權目標價、F 欄相對基準
s_rows = [r for r in range(8, SE.max_row + 1) if isinstance(SE.cell(r, 6).value, (int, float))]
lo = [r for r in s_rows if str(SE.cell(r, 1).value).endswith("（低）")]
hi = [r for r in s_rows if str(SE.cell(r, 1).value).endswith("（高）")]
oth = [r for r in s_rows if r not in lo and r not in hi]
s_end = max(s_rows)
TOR = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": "F", "fmt": "sp1"}
def big(rs): return [r for r in rs if abs(SE.cell(r, 6).value) >= 0.1]
def small(rs): return [r for r in rs if abs(SE.cell(r, 6).value) < 0.1]
def main_tor(side, rs):
    if big(rs):
        return dict(TOR, title=f"加權目標價相對基準：輸入取{side}值（變動達兩位數百分比的案例）", rows=",".join(map(str, big(rs))))
    return dict(TOR, title=f"加權目標價相對基準：輸入取{side}值（沒有變動達兩位數百分比的案例；列出影響最大者）", rows=",".join(map(str, sorted(rs, key=lambda r: -abs(SE.cell(r, 6).value))[:1])))
sens_blocks = [main_tor("低", lo), main_tor("高", hi), {"type": "note", "parts": ["@Sens!A1"]},
               {"type": "kv", "items": [["建置快照代號", "@Sens!D2"], ["@Sens!A3", "@Sens!D3"]]}]
if oth:
    sens_blocks.append(dict(TOR, title="加權目標價相對基準：只看單一情境、終值替代法", rows=",".join(map(str, oth))))
for side, rs in (("低", lo), ("高", hi)):
    shown = set(map(int, main_tor(side, rs)["rows"].split(",")))
    rest = [r for r in small(rs) if r not in shown]
    if rest:
        sens_blocks.append(dict(TOR, title=f"加權目標價相對基準：輸入取{side}值，其餘案例", rows=",".join(map(str, rest)), collapse=True, summary=f"展開：輸入取{side}值、變動未達兩位數百分比的案例"))
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：全部案例（加權目標價、DCF 目標價與估值評等）",
                    "head": ["@Sens!A7", "@Sens!B7", "@Sens!E7", "@Sens!F7", "@Sens!G7", "@Sens!H7", "@Sens!I7"],
                    "rowgen": {"sheet": "Sens", "rows": f"8-{s_end}", "cols": ["A", "B", "E", "F", "G", "H", "I"], "fmts": ["", "", "n2", "sp1", "n2", "sp1", ""]}})
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}

hero = {"rec": vv("rec"),
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n2"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n2"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n2"), ""],
                  [il("px"), iv("px", "n3"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等", vv("rt"), "（只看上檔空間）；收支關卡", vv("gj"), "（情境 B 自由現金流扣股利，FY2026–FY2030 合計 ", vv("gate", "C", "n0"),
                 " 百萬人民幣）。Comparable 用同業營收倍數 ", vv("evm", "B", "n2"), " 倍（現行 ", vv("evs", "B", "n2"), " 倍）；只看 DCF 的目標價為 ", vv("dcf", "B", "n2"), " 港元。"]}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "營收依分部堆疊（百萬人民幣）與營業利益率：FY2026–FY2032（情境 B）：FY2024–FY2032（情境 B）", "x": f"@PL!F{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev", "B", "F"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "F"), "color": c} for k, c in (("adas", "#0b7a80"), ("rob", "#6fa8dc"), ("oth", "#8a8c8c"))],
     "lines": [{"name": [f"@PL!A{B['opm']}"], "y": pl("opm", "B", "F"), "axis": "r", "color": "#3b3838", "labels": True},
               {"name": [f"@PL!A{B['gm']}"], "y": pl("gm", "B", "F"), "axis": "r", "color": "#b4472f", "dash": True}],
     "caption": "堆疊＝本模型推導的 ADAS、機器人等與其他營收（合計等於合併營收；公司只揭露產品、解決方案、服務與其他分類，ADAS 與機器人拆分為推導值）。實線＝營業利益率、虛線＝毛利率（右軸）。FY2026 起為模型。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["承重變數：每台 NOA 車對應速騰 ADAS 出貨的年變化（情境 B ", iv("shr_B", "p1"), "）、ADAS 單價年變化（", iv("dasp_B", "p0"), "）、毛利率起點（", iv("gm0", "p1"), "）、營業費用對營收成長的彈性（", iv("e_ox", "n2"), "）。"]]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n2"), " 港元（DCF ", vv("dcf", "B", "n2"), "、Comparable ", vv("cmp", "B", "n2"), "），對現價 ", iv("px", "n3"), " 港元上檔 ", vv("up", "B", "sp1"),
         "，估值評等", vv("rt"), "、投資建議", vv("rec"), "。"],
        ["上檔來源：DCF 貢獻 ", vv("a1", "B", "sp1"), "、Comparable 貢獻 ", vv("a2", "B", "sp1"), "。市場給速騰的營收倍數 ", vv("evs", "B", "n2"), " 倍，為同業中位數 ", vv("evm", "B", "n2"),
         " 倍的 ", vv("rel", "B", "n2"), "；LiDAR 同業中位數 ", vv("sg1", "B", "n2"), " 倍。結論承重在這個類比。"],
        ["獲利拐點：情境 B 的營業利益 FY2025 ", f"@PL!E{B['op']}|n0", " → FY2026 ", f"@PL!F{B['op']}|n0", " → FY2028 ", f"@PL!H{B['op']}|n0", " 百萬人民幣；FY2032 ", f"@PL!L{B['op']}|n0", "；毛利率 ", f"@PL!F{B['gm']}|p1", " → ", f"@PL!L{B['gm']}|p1", "，情境 B 到 FY2032 仍未轉正。"],
        ["收支：情境 B FY2026–FY2030 自由現金流合計 ", vv("gate", "C", "n0"), " 百萬人民幣，", vv("gj"), "；情境 C 合計 ", vv("gate", "D", "n0"), "。"],
        "@README!A6"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A7"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "三個目標價與現價（港元；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}] ,
     "ref": {"label": il("px"), "value": iv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "note", "parts": ["@README!A5"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n2"), ("cmp", "n2"), ("cmpu", "n2"), ("tp", "n2"), ("px", "n3"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("lq", "n0"), ("dist", "n1"), ("bnd", ""))]
             },
    {"type": "table", "caption": "情境目標價（港元）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", "n2"]}},
    {"type": "table", "caption": "收支關卡（百萬人民幣）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3"],
     "rows": [[f"@Val!A{VR['gate']}", vv("gate", "B", "n0"), vv("gate", "C", "n0"), vv("gate", "D", "n0")]]},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（百萬人民幣；每股為港元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n0"), ("eq", "n0"), ("d12", "n2"), ("tvsh", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n0"), ("tv", "n0"), ("pvtv", "n0"), ("stub", "n0"), ("psv", "n3"), ("alt", "n2"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n2"), ("erp", "p1"), ("crp", "p2"), ("sh_v", "n1"), ("bb_post", "n1"), ("hkcny", "n4"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n0", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n0", "", ""], [f"@DCF!A{DK['dd']}", f"@DCF!B{DK['dd']}|n1", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因與倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("pd2", "n2"), ("evs", "n2"), ("evm", "n2"), ("rel", "n2"), ("sg1", "n2"), ("sg2", "n2"), ("ipe", "n2"), ("pe", "n1"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業 EV ÷ 近四季營收", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!I3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag", nourl("In", f"I{inr(k)}")] for k in ("ev_hsai", "ev_oust", "ev_mbly", "ev_hor", "ev_med", "ev_self")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"))]}]}

rv0 = find(RC, "反解目標")
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價（上檔為零）及落在評等邊界所需的值，其他輸入不變。「無解」表示在搜尋區間內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{find(RC, '反解（整本重算', exact=False)}", "head": [f"@Recon!{c}{rv0}" for c in "ACDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{rv0 + 1}-{RC.max_row}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n3", "n3", ""]}}]}

def prow(lab, k, f): return [lab] + [f"@PL!{c}{B[k]}|{f}" for c in STD]
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "營收（百萬人民幣）× 情境 A／B／C 與營業利益率（情境 B）：FY2023–FY2032", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：營收"], "y": pl("rev", s), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['opm']}", "（情境 B）"], "y": pl("opm"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境營收；折線＝情境 B 營業利益率（右軸）。FY2023–FY2025 為實績（三情境相同），分化從 FY2026 開始。"},
    {"type": "note", "parts": ["會計口徑：IFRS（公司法定財報準則），即 GAAP 口徑；公司另揭露經調整淨虧損（Non-GAAP，加回股份支付），本表不用，因 DCF 與評價都以法定口徑計。營業利益為本模型口徑（毛利 − 研發、銷售、行政）；公司的營業虧損另含其他收入與其他損益，本表把它們列在其他收入淨額。FY2023–FY2024 為資料商彙整，銷售與行政只有合計，FY2023 業外含上市前優先股公允價值變動的巨額虧損。"]},
    {"type": "table", "caption": "標準損益表（IFRS／GAAP 口徑；百萬人民幣，EPS 為人民幣元）：FY2023–FY2025 實績、FY2026 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [prow(f"@PL!A{B['rev']}", "rev", "n1"), prow(f"@PL!A{B['gr']}", "gr", "p1"), prow(f"@PL!A{B['cogs']}", "cogs", "n1"), prow(f"@PL!A{B['gp']}", "gp", "n1"),
              prow(f"@PL!A{B['gm']}", "gm", "p1"), prow(f"@PL!A{B['ox']}", "ox", "n1"), prow(f"@PL!A{B['op']}", "op", "n1"), prow(f"@PL!A{B['opm']}", "opm", "p1"),
              prow(f"@PL!A{B['oi']}", "oi", "n1"), prow(f"@PL!A{B['pbt']}", "pbt", "n1"), prow(f"@PL!A{B['tax']}", "tax", "n1"), prow(f"@PL!A{B['ni']}", "ni", "n1"),
              prow(f"@PL!A{B['nm']}", "nm", "p1"), prow(f"@PL!A{B['eps']}", "eps", "n3"), prow(f"@PL!A{B['nig']}", "nig", "p1")]},
    {"type": "table", "caption": "比率區（情境 B）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [prow(f"@PL!A{B[k]}", k, "p1") for k in ("gr", "gm", "opm", "nm", "nig")]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "營收（百萬人民幣）", "fmt": "n0", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "營業利益（百萬人民幣）", "fmt": "n0", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "營業利益率", "fmt": "p1", "rows": {s: pl("opm", s, "F") for s in "ABC"}},
                 {"name": "EPS（人民幣元）", "fmt": "n3", "chart": True, "rows": {s: pl("eps", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣股利（百萬人民幣；收支關卡用）", "fmt": "n0", "rows": {s: pl("fcfad", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流（百萬人民幣）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[3:]],
     "rowgen": {"sheet": "PL", "rows": f"{B['da']}-{B['fcfad']}", "cols": ["A"] + STD[3:], "fmts": [""] + ["n1"] * len(STD[3:])}}]}

drivers = {"id": "drivers", "title": "驅動因子：中國城市 NOA 新車 → ADAS 出貨與單價 → 營收", "blocks": [
    {"type": "theme", "title": "中國城市 NOA 新車數（萬台，智慧車線）× 三情境與 ADAS 出貨（萬顆，情境 B）：FY2026–FY2032", "x": f"@PL!F{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "n0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：中國城市 NOA 新車"], "y": pl("mkt", s, "F"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['au']}", "（情境 B）"], "y": pl("au", "B", "F"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "長條＝智慧車線模型的中國城市 NOA 新車數（情境 A 讀高、B 讀基準、C 讀低）；折線＝情境 B 速騰 ADAS LiDAR 出貨（右軸，萬顆）。FY2026 出貨以上半年實績、第三季公告與第四季假設合成。"},
    {"type": "note", "parts": [f"@PL!A{find(P, 'PL — ', exact=False)}"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("u_a_q4", "n1"), ("u_r_q4", "n1"), ("dasp26", "p1"), ("dasp", "p1"), ("dasp_r", "p1"), ("pen26", "p1"), ("pen28", "p1"), ("pen30", "p1"), ("pen32", "p1"), ("shr", "p1"), ("gr", "p1"), ("goth", "p1"), ("gmd", "p1"))]},
    {"type": "table", "caption": "分部營收推導（情境 B；百萬人民幣）：NOA 新車 → ADAS 出貨 × 單價；機器人出貨 × 單價；加其他 → 合併營收；核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[1:]],
     "rowgen": {"sheet": "PL", "rows": f"{B['mkt']}-{B['gr']}", "cols": ["A"] + STD[1:], "fmts": [""] + ["n1"] * len(STD[1:]),
                "row_fmt": {str(B["gr"]): {c: "p1" for c in STD[1:]}, str(B["pen"]): {c: "p1" for c in STD[1:]}, str(B["kf"]): {c: "n3" for c in STD[1:]}, str(B["mkt"]): {c: "n0" for c in STD[1:]},
                            str(B["asp_a"]): {c: "n0" for c in STD[1:]}, str(B["asp_r"]): {c: "n0" for c in STD[1:]}}}},
    {"type": "table", "caption": "費用推導（情境 B；百萬人民幣）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD[1:]],
     "rowgen": {"sheet": "PL", "rows": f"{B['cogs']}-{B['opm']}", "cols": ["A"] + STD[1:], "fmts": [""] + ["n1"] * len(STD[1:]),
                "row_fmt": {str(B["gm"]): {c: "p1" for c in STD[1:]}, str(B["opm"]): {c: "p1" for c in STD[1:]}}}},
    {"type": "h3", "text": "一、出貨與單價（營收）"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("u_a_q4_B", "n1"), ("u_r_q4_B", "n1"), ("rho", "p0"), ("dasp26_B", "p1"), ("dasp_B", "p1"), ("dasp_r_B", "p1"), ("shr_B", "p1"), ("gr_B", "p0"), ("fade_r", "n2"), ("oth26", "n0"), ("goth_B", "p0"), ("pen26_B", "p1"), ("pen27_B", "p1"), ("pen30_B", "p1"), ("pen32_B", "p1"))]},
    {"type": "h3", "text": "二、毛利與費用"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("gm0", "p1"), ("gmd_B", "p1"), ("e_ox", "n2"), ("oi26", "n0"), ("oi_lt", "n0"), ("tax_w", "p0"))]},
    {"type": "h3", "text": "三、全公司（現金流、折現、評價）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("fc_da", "p1"), ("fc_cx", "p1"), ("fc_wc", "p1"), ("beta", "n2"), ("erp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("sh_v", "n1"), ("cash_v", "n0"),
               ("ev_med", "n2"), ("rel_k", "n2"), ("w_dcf", "p0"))]},
    {"type": "table", "collapse": True, "summary": "展開：固定輸入（實績、換算值）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag"] for k, f in (("u_a_h1", "n2"), ("u_r_h1", "n2"), ("u_a_q3", "n2"), ("u_r_q3", "n2"), ("prod_h1", "n1"), ("asp_a0", "n0"), ("asp_r0", "n0"), ("sh_26", "n1"), ("pay_r", "p0"),
                                                                   ("fx_hk", "n4"), ("fx_cn", "n4"), ("hkcny", "n4"), ("debt_v", "n1"), ("lease_v", "n1"), ("h1fcf", "n1"), ("rf", "p2"), ("crp", "p2"), ("kd", "p2"))]}]}

Q = wb["Q"]
RE = wb["對帳"]
recon = {"id": "recon", "title": "對帳：實績、半年度分類與線層級變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, 'FY2026 ADAS 出貨 ÷', exact=False)}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""],
                "row_fmt": {str(find(RC, "FY2026 營收年增", exact=False)): {c: "p1" for c in "BCD"}, str(find(RC, "FY2026 ADAS 出貨 ÷", exact=False)): {c: "n2" for c in "BCD"}}}},
    {"type": "table", "caption": "半年度營收分類、出貨與損益（百萬人民幣）與分類對帳", "head": [f"@Q!{c}2" for c in "ABC"],
     "rowgen": {"sheet": "Q", "rows": f"3-{find(Q, 'ADAS 營收（推導）')}", "cols": list("ABC"), "fmts": [""] + ["n1"] * 2, "row_fmt": {str(find(Q, "毛利率")): {c: "p1" for c in "BC"}}}},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（智慧車線與共用零組件線）", "head": [f"@對帳!{c}2" for c in "ABCDGI"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{RE.max_row}", "cols": ["A", "B", "C", "D", "G", "I"], "fmts": ["", "", "", "", "tag", ""]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", "@Sens!D2"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價、匯率與折現率的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3", "@In!I3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}"), nourl("In", f"I{inr(k)}")] for k, f in (("px", "n3"), ("fx_hk", "n4"), ("fx_cn", "n4"), ("rf", "p2"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"Hist": ["A"], "Q": ["A"], "對帳": ["A", "C"]}, "empty": "沒有標為未取得的項目。"}]}

cfg = {"company": "速騰聚創 RoboSense", "ticker": "2498.HK", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "車用 LiDAR 出貨成長能否跑贏單價年降、機器人 LiDAR 能否把營收規模推到攤平研發費用；市場給的營收倍數低於同業，是否合理",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：中國城市 NOA 讀智慧車線 高、份額上升、單價年降放緩、毛利率改善）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：中國城市 NOA 讀智慧車線 低、份額流失、價格戰）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

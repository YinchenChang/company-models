# 產生 physical_ai/kion/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/kion/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/kion/html.json
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

LAB = dict(yr="年度（數值，公式用）", ias="IAS（Dematic 等）營收", its="ITS（Linde、STILL 等）營收", cor="公司本部與抵銷營收", rev="營收（合併）",
           chk="檢查：分部合計 − 合併營收（應為 0）", gr="營收年增率", iasm="IAS 調整後 EBIT 率", itsm="ITS 調整後 EBIT 率", adj="調整後 EBIT（分部合計＋本部）",
           adjm="調整後 EBIT 率", cogs="營業成本", gp="毛利", gm="毛利率", opx="營業費用（銷管＋研發＋其他；毛利 − 營業利益倒推）",
           adjx="調整項（購買價格分攤與非經常項目）", op="營業利益（EBIT，報告口徑）", opm="營業利益率", fin="業外：淨財務費用", pbt="稅前利益", tax="所得稅",
           ni="淨利（含非控制權益）", nci="非控制權益損益", nia="歸屬母公司淨利", nm="淨利率", niy="淨利年增率", sh="股數（百萬股）", eps="EPS（歐元）",
           da="折舊攤銷", cx="資本支出（含租賃事業資產投資）", nwc="營運資金（年末）", dwc="營運資金增加",
           fcff="FCFF（EBIT 稅後＋折舊 − 資本支出 − 營運資金增加；歷史＝公司定義自由現金流）", intr="稅後淨利息", div="股利",
           fcfad="自由現金流扣股利（收支關卡用；FCFF＋稅後淨利息 − 股利）")
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm)
    PR[s] = dict(hdr=h + 1, **{k: find(P, t, h) for k, t in LAB.items()})
B = PR["B"]
YC = "CDEFGHIJKL"   # FY2023 … FY2032
STD = list(YC)
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VL = dict(dcfr="DCF 情境價值推到 12 個月後", cu="Comparable（未調整）：混合前瞻 P/E × FY2028 EPS", ca="Comparable（相對倍數調整後，進加權）",
          tps="各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", tp="加權目標價（歐元）", dcf="DCF 目標價（情境加權）", cmp="Comparable 目標價（情境加權，進加權）",
          px="現價（In 頁 px）", up="上檔空間", upc="Comparable 目標價上檔空間", rt="估值評等（只看上檔空間）",
          gate="收支關卡：FY2026–FY2030 自由現金流扣股利合計（EUR B；情境 A／B／C）", gj="入不敷出判定", rec="投資建議", dist="距最近級距邊界（百分點）",
          bnd="翻轉目標價：升一級／降一級", a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻",
          pd1="期間：模型折現（評價日 2026-12-31 → 目標時點，年）", pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC",
          pe="隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 EPS", pe3="現價 ÷ 情境 B FY2028 EPS（模型）", pe4="現價 ÷ 情境 B FY2026 EPS（模型）",
          wi="IAS 占 FY2028 調整後 EBIT 比例（混合倍數權重）", mx="混合前瞻 P/E＝權重 × 同業中位數 ＋ (1 − 權重) × pe_its", mc="市值（EUR B）",
          ev="企業價值（市值＋工業淨負債口徑 nd_v＋非控制權益；EUR B）", evs="EV ÷ FY2026 調整後 EBIT（倍；情境 B）")
VR = {k: find(V, t) for k, t in VL.items()}
VR["liq"] = find(V, "期間流動性", exact=False)
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")
def iv(k, f=""): return f"@In!E{inr(k)}" + (f"|{f}" if f else "")
def il(k): return f"@In!B{inr(k)}"
def nourl(sh, coord, f=""):
    v = wb[sh][coord].value
    return "（出處含網址，見 xlsx）" if isinstance(v, str) and "http" in v else f"@{sh}!{coord}" + (f"|{f}" if f else "")
def prow_in(k, f): return [il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"]

DB = find(D, "情境 B（基準）")
DK = {k: find(D, t) for k, t in dict(ke="ke（股權成本）", e="股權市值 E（股數 × 現價；EUR B）", kd="稅後 kd", wacc="WACC", g="永續成長率 g＝MIN(g_cap, rf)", ronic="RONIC",
                                       nc="評價日淨負債與非控制權益（EUR B）＝nd_v ＋ nci_v").items()}
DBR = {k: find(D, t, DB) for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", eq="股權價值", psv="每股價值（評價日）",
                                          d12="DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）", tvsh="終值占 EV", alt="終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）").items()}
DA, DC = find(D, "情境 A（樂觀）"), find(D, "情境 C（保守）")
def dsc(k):
    t = D.cell(DBR[k], 1).value
    return [f"@DCF!B{find(D, t, DA)}", f"@DCF!B{DBR[k]}", f"@DCF!B{find(D, t, DC)}"]

# Sens
s_rows = [r for r in range(5, SE.max_row + 1) if isinstance(SE.cell(r, 5).value, (int, float)) and isinstance(SE.cell(r, 3).value, str)]
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
sb = find(SE, "build_id"); ss = find(SE, "模型狀態")
sens_blocks = [main_tor("低", lo), main_tor("高", hi),
               {"type": "note", "parts": ["@Sens!A1"]},
               {"type": "kv", "items": [["@Sens!A2", "@Sens!B2|n2"], ["@Sens!C2", "@Sens!D2|n2"], ["快照狀態", "@Sens!E2"], ["模型狀態", f"@Sens!D{ss}"], ["build_id", f"@Sens!D{sb}"]]}]
if oth:
    sens_blocks.append(dict(TOR, title="加權目標價相對基準：只看單一情境、終值替代法", rows=",".join(map(str, oth))))
for side, rs in (("低", lo), ("高", hi)):
    shown = set(map(int, main_tor(side, rs)["rows"].split(",")))
    rest = [r for r in small(rs) if r not in shown]
    if rest:
        sens_blocks.append(dict(TOR, title=f"加權目標價相對基準：輸入取{side}值，其餘案例", rows=",".join(map(str, rest)), collapse=True,
                                summary=f"展開：輸入取{side}值、變動未達兩位數百分比的案例"))
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：全部案例（加權目標價、DCF 目標價與估值評等）",
                    "head": ["@Sens!A4", "@Sens!B4", "@Sens!D4", "@Sens!E4", "@Sens!F4", "@Sens!G4", "@Sens!H4"],
                    "rowgen": {"sheet": "Sens", "rows": f"5-{s_end}", "cols": ["A", "B", "D", "E", "F", "G", "H"], "fmts": ["", "", "n2", "sp1", "n2", "sp1", ""]}})
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}

hero = {"rec": vv("rec"),
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n2"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n2"), ["上檔空間 ", vv("upc", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n2"), ""],
                  [il("px"), iv("px", "n2"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等", vv("rt"), "；收支關卡", vv("gj"), "（情境 B 自由現金流扣股利 FY2026–FY2030 合計 ", vv("gate", "C", "n2"),
                 " 十億歐元）。現價對應市值 ", vv("mc", "B", "n2"), " 十億歐元、約 ", vv("pe4", "B", "x1"),
                 " 倍模型 FY2026 EPS；上檔主要來自 ITS 與 IAS 的利潤率回升假設，屬 Assumed，不是已揭露事實。"]}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "分部營收（十億歐元）與調整後 EBIT 率：FY2024–FY2032（情境 B）", "x": f"@PL!D{B['hdr']}:L{B['hdr']}", "fmt": "n1", "fmt_r": "p1", "stack": True, "pad": 1.3,
     "total": pl("rev", "B", "D"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "D"), "color": c} for k, c in (("its", "#8a8c8c"), ("ias", "#0b7a80"))],
     "lines": [{"name": [f"@PL!A{B['adjm']}"], "y": pl("adjm", "B", "D"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝兩個報告分部營收（本部抵銷為小負數，未畫；合計列為合併營收）；折線＝調整後 EBIT 率（右軸）。ITS（堆高機）約占七成，IAS（倉儲自動化）成長較快。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["承重變數：ITS 調整後 EBIT 率 FY2029 目標（", iv("its_tm_B", "p1"), "）、IAS 調整後 EBIT 率 FY2029 目標（", iv("ias_tm_B", "p1"), "）、IAS FY2027 營收年增（",
         iv("g27i_B", "p0"), "）、ITS 適用前瞻 P/E（", iv("pe_its", "n1"), " 倍）。"],
        "@README!A8"]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n2"), " 歐元（DCF ", vv("dcf", "B", "n2"), "、Comparable ", vv("cmp", "B", "n2"), "），對現價 ", iv("px", "n2"), " 歐元上檔 ", vv("up", "B", "sp1"),
         "，估值評等", vv("rt"), "、投資建議", vv("rec"), "。"],
        ["Comparable 倍數是混合：IAS 占 FY2028 調整後 EBIT ", vv("wi", "C", "p0"), "，適用同業中位數 ", iv("pe_med", "n1"), " 倍；其餘適用 ITS 倍數 ", iv("pe_its", "n1"), " 倍；混合 ",
         vv("mx", "C", "n1"), " 倍。同業皆為自動化高成長公司，ITS 倍數為自我參照的類比。"],
        ["下檔：情境 C（IAS 營收下滑、兩分部利潤率不回升）目標價 ", f"@Val!D{VR['tps']}|n2", " 歐元，情境 A ", f"@Val!B{VR['tps']}|n2", " 歐元。"],
        ["評等邊界：", "@README!A5"],
        "@README!A6"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A7"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "三個目標價與現價（歐元；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}],
     "ref": {"label": il("px"), "value": iv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n2"), ("cmp", "n2"), ("tp", "n2"), ("px", "n2"), ("up", "sp1"), ("upc", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("dist", "n1"), ("bnd", ""))]
             + [[f"@Val!A{VR['liq']}", vv("liq", "B", "n2")]]},
    {"type": "table", "caption": "情境目標價（歐元）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", "n2"]}},
    {"type": "table", "caption": "收支關卡（十億歐元；情境 A／B／C）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@Val!A{VR['gate']}", vv("gate", "B", "n2"), vv("gate", "C", "n2"), vv("gate", "D", "n2")]]},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（十億歐元；每股為歐元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n2"), ("eq", "n2"), ("psv", "n2"), ("d12", "n2"), ("tvsh", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n2"), ("tv", "n2"), ("pvtv", "n2"), ("alt", "n2"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n3"), ("erp", "p1"), ("sh_v", "n1"), ("ind_25", "n3"), ("nd_v", "n3"), ("nci_v", "n3"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n3", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n2", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因、期間與倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("pd1", "n2"), ("pd2", "n2"), ("pe", "x1"), ("pe3", "x1"), ("pe4", "x1"), ("evs", "x1"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻本益比（換算至股價基準日）", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}")] for k in ("pe_auto", "pe_daif", "pe_geek", "pe_med", "pe_its", "cons_pe")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [prow_in(k, f) for k, f in (("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"))]}]}

rv_h = find(RC, "反解目標")
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價（上檔為零）及落在評等邊界所需的值，其他輸入不變。「無解」表示在搜尋網格內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{find(RC, '反解（整本重算', exact=False)}", "head": [f"@Recon!{c}{rv_h}" for c in "ACDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{rv_h + 1}-{RC.max_row}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n3", "n3", ""]}}]}

def plrow(k, f): return [f"@PL!A{B[k]}"] + [f"@PL!{c}{B[k]}|{f}" for c in STD]
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "table", "caption": "標準損益表（IFRS；十億歐元，EPS 為歐元）：FY2023–FY2025 實績（FY2023 分部、毛利、稅等未找到）、FY2026 公司指引中點、FY2027 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [plrow("rev", "n3"), plrow("gr", "p1"), plrow("cogs", "n3"), plrow("gp", "n3"), plrow("gm", "p1"), plrow("opx", "n3"), plrow("op", "n3"), plrow("opm", "p1"),
              plrow("fin", "n3"), plrow("pbt", "n3"), plrow("tax", "n3"), plrow("ni", "n3"), plrow("nci", "n3"), plrow("nia", "n3"),
              plrow("sh", "n1"), plrow("eps", "n2"), {"sub": "比率"}, plrow("nm", "p1"), plrow("niy", "p1"), plrow("adj", "n3"), plrow("adjm", "p1")]},
    {"type": "note", "parts": ["會計口徑：IFRS 報告口徑（相當於 GAAP）；調整後 EBIT 為公司自訂的 Non-GAAP 口徑，排除購買價格分攤攤銷與非經常項目，是公司指引的口徑，調整項見 PL 頁。預測期營業費用由毛利率與營業利益倒推，不影響營業利益。FY2023 分部、毛利、財務費用與稅未找到，標「未找到」。"]},
    {"type": "theme", "title": "營收（十億歐元）× 情境 A／B／C 與營業利益率（情境 B）：FY2023–FY2032", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n1", "fmt_r": "p1", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：營收"], "y": pl("rev", s), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['opm']}", "（情境 B）"], "y": pl("opm"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境營收；折線＝情境 B 營業利益率（右軸）。FY2023–FY2026 三情境相同，分化從 FY2027 開始。"},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "營收（十億歐元）", "fmt": "n2", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "營業利益（十億歐元）", "fmt": "n3", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "營業利益率", "fmt": "p1", "rows": {s: pl("opm", s, "F") for s in "ABC"}},
                 {"name": "EPS（歐元）", "fmt": "n2", "chart": True, "rows": {s: pl("eps", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣股利（十億歐元；收支關卡用）", "fmt": "n3", "rows": {s: pl("fcfad", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流（十億歐元）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['da']}-{B['fcfad']}", "cols": ["A"] + STD, "fmts": [""] + ["n3"] * len(STD)}}]}

drivers = {"id": "drivers", "title": "驅動因子：市場成長 → IAS；循環與利潤率 → ITS → 合併營收與利潤", "blocks": [
    {"type": "theme", "title": "兩分部調整後 EBIT 率：FY2024–FY2032（情境 B）", "x": f"@PL!D{B['hdr']}:L{B['hdr']}", "fmt": "p1", "fmt_r": "p1", "stack": False, "pad": 1.2,
     "bars": [{"name": [f"@PL!A{B['gr']}"], "y": pl("gr", "B", "D"), "color": "#8a8c8c"}],
     "lines": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "D"), "axis": "r", "color": c, "labels": True} for k, c in (("iasm", "#0b7a80"), ("itsm", "#b4472f"))],
     "caption": "長條＝合併營收年增率；實線＝IAS 與 ITS 調整後 EBIT 率（右軸）。"},
    {"type": "note", "parts": ["@PL!A1"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("g27i", "p1"), ("go", "p1"), ("ias_tm", "p1"), ("g27t", "p1"), ("g28t", "p1"), ("its_tm", "p1"), ("gm", "p1"))]
             + [[f"@In!B{inr('scn_w')}", f"@In!D{inr('scn_w')}|p0", f"@In!E{inr('scn_w')}|p0", f"@In!F{inr('scn_w')}|p0"]]},
    {"type": "table", "caption": "營收線推導（情境 B；十億歐元）：兩分部 → 合併營收；核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['ias']}-{B['itsm']}", "cols": ["A"] + STD, "fmts": [""] + ["n3"] * len(STD),
                "row_fmt": {str(B["gr"]): {c: "p1" for c in STD}, str(B["iasm"]): {c: "p1" for c in STD}, str(B["itsm"]): {c: "p1" for c in STD}}}},
    {"type": "h3", "text": "一、IAS 與 ITS（命題模組）"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [prow_in(k, f) for k, f in (("g27i_B", "p1"), ("go_B", "p1"), ("ias_sp", "p1"), ("ias_tm_B", "p1"), ("g27t_B", "p1"), ("g28t_B", "p1"), ("its_tm_B", "p1"), ("gm_B", "p1"),
                                          ("pe_its", "n1"), ("ias26", "n3"), ("its26", "n3"), ("ias_e26", "n3"), ("its_e26", "n3"))]},
    {"type": "h3", "text": "二、損益下半段、現金流、折現"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [prow_in(k, f) for k, f in (("cor_rev", "n3"), ("cor_e", "n3"), ("adj26", "n3"), ("adjx", "n3"), ("fcf26", "n3"), ("fin", "n3"), ("tax_w", "p0"), ("nci", "n3"), ("pay_r", "p0"),
                                          ("fc_cx", "p1"), ("nwc_r", "p1"), ("nci_v", "n3"), ("beta", "n3"), ("erp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("w_dcf", "p0"))]},
    {"type": "table", "collapse": True, "summary": "展開：固定輸入（股數、淨負債、折舊）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag"] for k, f in (("sh_v", "n1"), ("ind_25", "n3"), ("div26", "n4"), ("nd_v", "n3"), ("fc_da", "p1"), ("rf", "p2"))]}]}

RE = wb["對帳"]
recon = {"id": "recon", "title": "對帳：公司指引、實績與線層級變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, '市值隱含', exact=False)}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n3", "n3", "n3", ""]}},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（倉儲物流自動化線）", "head": [f"@對帳!{c}2" for c in "ACDG"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{RE.max_row}", "cols": ["A", "C", "D", "G"], "fmts": ["", "", "", "tag"]}}]}

RK = wb["Risk"]
sb = find(SE, "build_id")
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", f"@Sens!D{sb}"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價與市場資料的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!I3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag", nourl("In", f"I{inr(k)}")] for k, f in (("px", "n2"), ("rf", "p2"), ("cons_pe", "n1"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"In": ["B", "I"], "Hist": ["A"]}, "empty": "沒有標為未找到的項目。"}]}

cfg = {"company": "KION", "ticker": "KGX.DE", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "倉儲自動化（IAS）的訂單減速與堆高機（ITS）的利潤率，在三情境下能否支持現價；市場是否把 KION 單純當堆高機公司定價",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：IAS 訂單回補、兩分部利潤率回升）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：IAS 營收下滑、利潤率不回升）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

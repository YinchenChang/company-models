# 產生 physical_ai/unitree/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/unitree/gen_html_json.py [成品 xlsx]（預設 dist/ 內最新的 xlsx）→ physical_ai/unitree/html.json
import glob, json, os, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
X = sys.argv[1] if len(sys.argv) > 1 else sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))[-1]
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

LAB = dict(lship="全球人形出貨（線 2b；台）", shr="宇樹人形份額（出貨 ÷ 全球）", hu="人形機器人出貨（台）", lpr="線 2b 全尺寸售價學習曲線年變化",
           hasp="人形機器人平均售價（萬元）", hrev="人形機器人營收", qu="四足機器人出貨（台）", qasp="四足機器人平均售價（萬元）", qrev="四足機器人營收",
           crev="組件及其他營收", rev="合併營收", chk="檢查：分部合計 − 合併營收", gr="營收年增率", hgm="人形機器人毛利率", qgm="四足機器人毛利率", cgm="組件及其他毛利率",
           cogs="營業成本", gp="毛利", gpm="毛利率", rd="研發費用", sm="銷售費用", ga="管理費用（不含股份支付）", sbc="股份支付費用", opex="營業費用合計",
           op="營業利益", opm="營業利益率", nonop="業外（利息收入、補助等）", resid="歷史：營業利益至淨利殘差（業外、稅金及附加、所得稅合計）", pbt="稅前利益", tax="所得稅",
           ni="淨利（歸屬母公司）", nim="淨利率", sh="平均股數（百萬股）", eps="EPS（元）", nig="淨利年增率", da="折舊攤銷", cx="資本支出", dwc="營運資金增加",
           fcff="FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", div="股利", fcfad="自由現金流扣股利（收支關卡用）")
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm)
    PR[s] = dict(hdr=h + 1, **{k: find(P, t, h) for k, t in LAB.items()})
B = PR["B"]
STD = list("CDEFGHIJKL")   # PL：C＝FY2023 … L＝FY2032
FC = list("FGHIJKL")       # 預測年 FY2026–FY2032
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VR = {k: find(V, t) for k, t in dict(dcfr="DCF 情境價值推到 12 個月後", cu="Comparable（未調整）：同業 EV／營收 TTM 中位數 × FY2027 營收 ＋ 淨現金，÷ 股數",
                                       ca="Comparable（相對倍數調整後，進加權）", tps="各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", tp="加權目標價（元）",
                                       dcf="DCF 目標價（情境加權）", cmp="Comparable 目標價（情境加權，進加權；同業 EV／營收）", px="現價（In 頁 px）", up="上檔空間", rt="估值評等（只看上檔空間）",
                                       gate="收支關卡：FY2026–FY2030 自由現金流扣股利合計（億元；情境 A／B／C）", gj="入不敷出判定", rec="投資建議", dist="距最近級距邊界（百分點）",
                                       bnd="翻轉目標價：升一級／降一級", a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻",
                                       pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC", pe="隱含本益比：加權目標價 ÷ 情境 B FY2028 EPS",
                                       pex="現價本益比：現價 ÷ 情境 B FY2027 EPS", evr="現行 EV／營收 TTM ÷ 同業中位數").items()}
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")
def iv(k, f=""): return f"@In!E{inr(k)}" + (f"|{f}" if f else "")
def il(k): return f"@In!B{inr(k)}"
def nourl(sh, coord, f=""):
    v = wb[sh][coord].value
    return "（出處含網址，見 xlsx）" if isinstance(v, str) and "http" in v else f"@{sh}!{coord}" + (f"|{f}" if f else "")
def prow3(keys):
    return [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in keys]
HEAD5 = ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"]

DB = find(D, "情境 B（基準）")
DK = {k: find(D, t) for k, t in dict(ke="ke（股權成本）", e="股權市值 E（億元）", dd="負債 D（有息負債）", kd="稅後 kd", wacc="WACC", g="永續成長率 g＝MIN(g_cap, rf)", ronic="RONIC",
                                       nc="淨現金（資產負債表 2026-09-30）＝貨幣資金 − 有息負債").items()}
DBR = {k: find(D, t, DB) for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", eq="股權價值", psv="每股價值（評價日）",
                                          d12="DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）", tvsh="終值占 EV", alt="終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）").items()}
DA, DC = find(D, "情境 A（樂觀）"), find(D, "情境 C（保守）")
def dsc(k):
    t = D.cell(DBR[k], 1).value
    return [f"@DCF!B{find(D, t, DA)}", f"@DCF!B{DBR[k]}", f"@DCF!B{find(D, t, DC)}"]

# Sens：表頭列（「驅動」）之下
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
               {"type": "note", "parts": ["讀法：所有單一參數案例與三個單一情境案例的估值評等都相同；在本模型的輸入區間內，人形出貨、份額與毛利率的變動都不足以改變評等，唯一能把目標價推近現價的是同業倍數（見反向檢驗）。"]},
               {"type": "kv", "items": [["@Sens!A4", "@Sens!B4|n1"], ["@Sens!C4", "@Sens!D4|n1"], ["@Sens!A3", "@Sens!D3"], ["@Sens!A2", "@Sens!D2"]]}]
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
                    "rowgen": {"sheet": "Sens", "rows": f"{SH + 1}-{s_end}", "cols": ["A", "B", "D", "E", "F", "G", "H"], "fmts": ["", "", "n1", "sp1", "n1", "sp1", ""]}})
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}

hero = {"rec": vv("rec"),
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n1"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n1"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n1"), ""],
                  [il("px"), iv("px", "n2"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等", vv("rt"), "；收支關卡", vv("gj"), "。現價只有單一來源，待交易所資料複核；現價隱含的同業倍數約為同業中位數的 ", vv("evr", "B", "n1"),
                 " 倍。上市僅約三個月，本年度分產品營收未揭露，結論是條件式。"]}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "合併營收依官方產品分部堆疊（億元）與毛利率：FY2023–FY2032（情境 B）", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k), "color": c} for k, c in (("hrev", "#0b7a80"), ("qrev", "#6fa8dc"), ("crev", "#c29a2b"))],
     "lines": [{"name": [f"@PL!A{B['gpm']}"], "y": pl("gpm"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝三個官方產品分部（合計等於合併營收）；折線＝毛利率（右軸；FY2023、FY2024 為主營業務毛利率）。FY2023–FY2025 為招股書實績；FY2026 為前三季實績加第四季假設，分產品占比為假設；FY2027 起為模型。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["核心命題的承重變數：全球人形出貨（讀線模型）、宇樹人形份額（情境 B 於 FY2032 收斂到 ", iv("shr32_B", "p0"), "）、人形毛利率（情境 B 於 FY2032 收斂到 ",
         iv("hgm32_B", "p0"), "）、同業企業價值對營收倍數（中位數 ", iv("ev_med", "n1"), " 倍）。"],
        "@README!A8"]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n1"), " 元（DCF ", vv("dcf", "B", "n1"), "、Comparable ", vv("cmp", "B", "n1"), "），對現價 ", iv("px", "n2"), " 元上檔 ", vv("up", "B", "sp1"),
         "，估值評等", vv("rt"), "、投資建議", vv("rec"), "。"],
        ["現價隱含：現價 ÷ 情境 B FY2027 EPS 為 ", vv("pex", "B", "n0"), " 倍；企業價值對營收倍數為同業中位數的 ", vv("evr", "B", "n1"), " 倍。要讓目標價等於現價，同業倍數須到約 ",
         f"@Recon!C{find(RC, '目標價＝現價（上檔 0%）', find(RC, '反解目標'))}|n0", " 倍。"],
        ["營收結構：FY2025 人形 ", f"@PL!E{B['hrev']}|n2", " 億元首次超過四足 ", f"@PL!E{B['qrev']}|n2", " 億元；人形銷量 ", f"@PL!E{B['hu']}|n0", " 台、平均售價 ",
         f"@PL!E{B['hasp']}|n2", " 萬元，買家以科研與教育為主。"],
        ["本年度：前三季營收 ", iv("q9_rev", "n2"), " 億元、淨利 ", iv("q9_ni", "n2"), " 億元；第三季研發費用大增，存貨較年初增加 ", iv("dwc26", "n2"), " 億元。"],
        "@README!A6"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A7"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "三個目標價與現價（元；豎線＝現價）", "fmt": "n1",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}],
     "ref": {"label": il("px"), "value": iv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n1"), ("cmp", "n1"), ("tp", "n1"), ("px", "n2"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("dist", "n1"), ("bnd", ""))]},
    {"type": "table", "caption": "情境目標價（元；情境權重見下方權重表）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1", "n1"]}},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（億元；每股為元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n1"), ("eq", "n1"), ("psv", "n1"), ("d12", "n1"), ("tvsh", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n1"), ("tv", "n1"), ("pvtv", "n1"), ("alt", "n1"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n2"), ("erp", "p1"), ("crp", "p2"), ("sh_v", "n2"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n1", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n1", "", ""], [f"@DCF!A{DK['dd']}", f"@DCF!B{DK['dd']}|n1", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因與隱含倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("pd2", "n2"), ("pe", "x1"), ("pex", "x1"), ("evr", "x1"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業企業價值對營收倍數（TTM）與中位數", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag"] for k in ("ev_ubt", "ev_dobot", "ev_geek", "ev_ldr", "ev_orb", "ev_rbs", "ev_med", "evs_self")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": HEAD5,
     "rows": prow3((("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"), ("dps_tgt", "n2")))}]}

RVH = find(RC, "反解目標")
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價（上檔為零）及落在升一級邊界所需的值，其他輸入不變。「無解」表示在搜尋區間內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{RVH - 1}", "head": [f"@Recon!{c}{RVH}" for c in "ACDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{RVH + 1}-{RC.max_row}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n2", "n2", ""]}}]}

def r_(lab, k, f, hist="C"):
    return [lab] + [f"@PL!{c}{B[k]}|{f}" for c in STD]
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "合併營收（億元）× 情境 A／B／C 與營業利益率（情境 B）：FY2026–FY2032", "x": f"@PL!F{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：合併營收"], "y": pl("rev", s, "F"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['opm']}", "（情境 B）"], "y": pl("opm", "B", "F"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境合併營收；折線＝情境 B 營業利益率（右軸）。FY2026 三情境相同（前三季實績加第四季假設），分化從 FY2027 開始。"},
    {"type": "note", "parts": ["會計口徑：中國企業會計準則（公司法定財報，即 GAAP 口徑），未做 Non-GAAP 調整；FY2025 管理費用含一次性股份支付（列為非經常性損益）。FY2023、FY2024 媒體轉述的招股書科目不全，未取得者標「未揭露」；FY2025 營業利益為毛利減四項費用的推得值（不含財務費用、稅金及附加）。"]},
    {"type": "table", "caption": "標準損益表（中國企業會計準則／GAAP 口徑；億元，EPS 為元）：FY2023–FY2025 實績、FY2026 前三季實績加第四季、FY2027 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [r_("營收（合併）", "rev", "n2"), r_("@PL!A%d" % B['gr'], "gr", "p1"), r_("@PL!A%d" % B['cogs'], "cogs", "n2"), r_("@PL!A%d" % B['gp'], "gp", "n2"),
              r_("@PL!A%d" % B['gpm'], "gpm", "p1"), r_("營業費用（合計：研發、銷售、管理、股份支付）", "opex", "n2"), r_("@PL!A%d" % B['op'], "op", "n2"),
              r_("@PL!A%d" % B['opm'], "opm", "p1"), r_("@PL!A%d" % B['nonop'], "nonop", "n2"), r_("@PL!A%d" % B['resid'], "resid", "n2"), r_("@PL!A%d" % B['pbt'], "pbt", "n2"),
              r_("@PL!A%d" % B['tax'], "tax", "n2"), r_("@PL!A%d" % B['ni'], "ni", "n2"), r_("@PL!A%d" % B['nim'], "nim", "p1"), r_("@PL!A%d" % B['sh'], "sh", "n1"),
              r_("EPS（元；基本，無稀釋性證券揭露）", "eps", "n2"), r_("@PL!A%d" % B['nig'], "nig", "p1")]},
    {"type": "table", "collapse": True, "summary": "展開：營業費用明細（情境 B；億元）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['rd']}-{B['opex']}", "cols": ["A"] + STD, "fmts": [""] + ["n2"] * len(STD)}},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "合併營收（億元）", "fmt": "n1", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "營業利益（億元）", "fmt": "n1", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "毛利率", "fmt": "p1", "rows": {s: pl("gpm", s, "F") for s in "ABC"}},
                 {"name": "EPS（元）", "fmt": "n2", "chart": True, "rows": {s: pl("eps", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣股利（億元；收支關卡用）", "fmt": "n2", "rows": {s: pl("fcfad", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流（億元）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in FC],
     "rowgen": {"sheet": "PL", "rows": f"{B['da']}-{B['fcfad']}", "cols": ["A"] + FC, "fmts": [""] + ["n2"] * len(FC)}}]}

drivers = {"id": "drivers", "title": "驅動因子：全球人形出貨 × 份額 × 售價 → 產品分部 → 合併營收", "blocks": [
    {"type": "theme", "title": "宇樹人形出貨（台）與平均售價（萬元，右軸）：FY2023–FY2032（情境 B）", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "n1", "stack": False, "pad": 1.3,
     "bars": [{"name": [f"@PL!A{B['hu']}"], "y": pl("hu"), "color": "#0b7a80"}],
     "lines": [{"name": [f"@PL!A{B['hasp']}"], "y": pl("hasp"), "axis": "r", "color": "#b4472f", "labels": True}],
     "caption": "長條＝宇樹人形出貨（FY2023–FY2025 為招股書銷量、FY2026 由營收推得、FY2027 起＝全球出貨 × 份額）；折線＝人形平均售價（右軸）。FY2027 起售價降幅借用線模型的全尺寸學習曲線。"},
    {"type": "note", "parts": [f"@PL!A{find(P, 'PL — ', exact=False)}"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("lnsel", "n0"), ("shr32", "p0"), ("hgm32", "p0"), ("qg", "p0"), ("qlt", "p0"), ("qgm", "p0"), ("cg", "p0"))]
             + [[f"@In!B{inr('scn_w')}", f"@In!D{inr('scn_w')}|p0", f"@In!E{inr('scn_w')}|p0", f"@In!F{inr('scn_w')}|p0"]]},
    {"type": "table", "caption": "分部營收推導（情境 B；億元，出貨為台、售價為萬元）：核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['lship']}-{B['gr']}", "cols": ["A"] + STD, "fmts": [""] + ["n2"] * len(STD),
                "row_fmt": {**{str(B[k]): {c: "n0" for c in STD} for k in ("lship", "hu", "qu")}, **{str(B[k]): {c: "p1" for c in STD} for k in ("shr", "lpr", "gr")}}}},
    {"type": "table", "caption": "分部毛利率（情境 B）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['hgm']}-{B['cgm']}", "cols": ["A"] + STD, "fmts": [""] + ["p1"] * len(STD)}},
    {"type": "h3", "text": "一、人形機器人（命題模組）"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": HEAD5,
     "rows": prow3((("lnsel_B", "n0"), ("shr32_B", "p0"), ("asp_k", "n2"), ("hgm26", "p0"), ("hgm32_B", "p0"), ("hs26", "p0"), ("dasp_h26", "p0")))},
    {"type": "h3", "text": "二、四足機器人、組件及其他"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": HEAD5,
     "rows": prow3((("qg_B", "p0"), ("qlt_B", "p0"), ("qasp", "p0"), ("qgm26", "p0"), ("qgm_B", "p0"), ("cs26", "p1"), ("dasp_q26", "p0"), ("cg_B", "p0"), ("cgm", "p0")))},
    {"type": "h3", "text": "三、全公司（本年度錨點、費用、現金流、折現）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": HEAD5,
     "rows": prow3((("q4k", "n2"), ("nonop26", "n2"), ("dwc26", "n2"), ("cx_26", "n2"), ("rd_r", "p0"), ("sm_r", "p0"), ("ga_r", "p0"), ("sbc_r", "p0"), ("tax_w", "p0"), ("int_r", "p1"),
                    ("fc_da", "p0"), ("fc_cx", "p0"), ("fc_wc", "p0"), ("pay_r", "p0"), ("beta", "n2"), ("erp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("w_dcf", "p0"), ("ev_med", "n2")))},
    {"type": "table", "collapse": True, "summary": "展開：本年度實績錨點與股數（固定輸入）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag"] for k, f in (("q9_rev", "n2"), ("q3_rev", "n2"), ("q9_ni", "n2"), ("q9_rd", "n2"), ("sh_out", "n2"), ("sh_pre", "n2"), ("sh_26", "n2"), ("cash_v", "n2"), ("debt_v", "n2"))]}]}

Q = wb["Q"]; RE = wb["對帳"]; LS = wb["Line"]
recon = {"id": "recon", "title": "對帳：實績、線模型與現價隱含", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, '宇樹現行 EV／營收 TTM vs 同業中位數')}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""],
                "row_fmt": {str(find(RC, "FY2026 淨利率", exact=False)): {c: "p1" for c in "BCD"}, str(find(RC, "FY2026 銷售費用率", exact=False)): {c: "p1" for c in "BCD"},
                            str(find(RC, "FY2026 宇樹人形份額", exact=False)): {c: "p1" for c in "BCD"}, str(find(RC, "FY2026 人形出貨", exact=False)): {c: "n0" for c in "BCD"}}}},
    {"type": "table", "caption": "季度合併營收與淨利（億元；累計原文，單季相減）", "head": [f"@Q!{c}2" for c in "ABCDEFGH"],
     "rowgen": {"sheet": "Q", "rows": "3-6", "cols": list("ABCDEFGH"), "fmts": [""] + ["n2"] * 6 + ["tag"]}},
    {"type": "table", "collapse": True, "summary": "展開：線模型讀入值（全球人形出貨與全尺寸出廠價）", "caption": "@Line!A1", "head": [f"@Line!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Line", "rows": f"3-{LS.max_row}", "cols": list("ABCDE"), "fmts": ["", "", "n0", "n0", "n0"]}},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（是否與線模型一致）", "head": [f"@對帳!{c}2" for c in "ABCGI"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{RE.max_row}", "cols": ["A", "B", "C", "G", "I"], "fmts": ["", "", "", "tag", ""]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", "@Sens!D2"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價與折現輸入的日期與來源", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3", "@In!I3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}"), nourl("In", f"I{inr(k)}")] for k, f in (("px", "n2"), ("rf", "p2"), ("beta", "n2"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未找到」與「未揭露」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Hist": ["A"], "Q": ["A"], "Recon": ["A"]}, "empty": "沒有標為未找到的項目。"}]}

cfg = {"company": "宇樹科技 Unitree", "ticker": "688836.SS", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "人形機器人從科研、教育與展示走到大量付費採用的速度，能否支撐現價隱含的估值；在三情境與同業倍數下，目標價離現價多遠",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：全球出貨讀線模型高欄、份額與毛利率較高）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：讀線模型低欄、份額與毛利率較低）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

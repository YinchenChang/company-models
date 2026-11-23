# 產生 physical_ai/mobileye/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；模型改版移列後重跑本檔即可。設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/mobileye/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/mobileye/html.json
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

LAB = dict(u="EyeQ 系統出貨（百萬套）", u_srd="Surround ADAS 出貨（百萬套）", u_sv="SuperVision／Chauffeur 出貨（百萬套）", u_base="基礎 ADAS 出貨（百萬套）",
           asp="基礎 ADAS 平均售價（美元）", rt="線 1 美國 Robotaxi 年新增車輛（台）", u_drv="Drive（Robotaxi）出貨（台）",
           base="基礎 ADAS（EyeQ 前視；子列）", srd="Surround ADAS（子列）", sv="SuperVision／Chauffeur（子列）", drv="Drive（Robotaxi；子列）",
           mo="Mobileye 其他（REM、NRE、記憶體轉嫁；子列）", mbly="Mobileye 分部營收", oth="Other（Moovit）分部營收", rev="營收（合併）",
           chk1="檢查：Mobileye 子列合計 − Mobileye 分部（應為 0）", chk2="檢查：分部合計 − 合併營收（應為 0）", gr="營收年增率",
           agp="調整後毛利（Non-GAAP）", agm="調整後毛利率（Non-GAAP）", aop="調整後營業費用（未扣 R&D Law 補助；Non-GAAP）", inc="R&D Law 補助（Non-GAAP 口徑）",
           aoi="調整後營業利益（Non-GAAP；分部績效合計）", aom="調整後營業利益率（Non-GAAP）", sbc="股票報酬", am="併購無形資產攤銷", oit="商譽減損與併購費用（減 R&D Law RSU 項）",
           cogs="營業成本（GAAP）", gp="毛利", gm="毛利率", opex="營業費用（GAAP，含研發淨額、行銷、管理、商譽減損）", op="營業利益（GAAP）", opm="營業利益率（GAAP）",
           fin="業外收入（財務收入淨額）", pbt="稅前利益", tax="所得稅（負值＝利益）", ni="稅後淨利", nim="淨利率", sh="平均股數（百萬股）", eps="EPS（GAAP，美元）",
           nig="淨利年增率", ani="調整後淨利（Non-GAAP）", aeps="調整後 EPS（Non-GAAP，美元）", ebit="DCF EBIT（調整後營業利益 − 股票報酬）", da="折舊", cx="資本支出",
           dwc="營運資金增加", fcff="FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", fat="稅後財務收入", div="股利", fcfad="自由現金流扣股利（收支關卡用；FCFF＋稅後財務收入 − 股利）")
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm)
    PR[s] = {k: find(P, t, h) for k, t in LAB.items()}
    PR[s]["hdr"] = h + 1
B = PR["B"]
YC = "CDEFGHIJKL"   # PL：C＝FY2023 … L＝FY2032
FC = "FGHIJKL"      # FY2026–FY2032
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VT = dict(dcfr="DCF 情境價值推到 12 個月後", cu="Comparable（未調整）：FY2028 調整後淨利 × 同業前瞻 P/E 中位數 ÷ 股數", ca="Comparable（相對倍數調整後，進加權）",
          tps="各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", tp="加權目標價（美元）", dcf="DCF 目標價（情境加權）", cmp="Comparable 目標價（情境加權，進加權）",
          px="現價（In 頁 px）", up="上檔空間", rt="估值評等（只看上檔空間）", gate="收支關卡：FY2026–FY2030 自由現金流扣股利合計（US$M；情境 A／B／C）", gj="入不敷出判定",
          rec="投資建議", dist="距最近級距邊界（百分點）", bnd="翻轉目標價：升一級／降一級", a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻",
          a3="上檔三段歸因：相對倍數調整貢獻", pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC", pe="隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 調整後 EPS")
VR = {k: find(V, t) for k, t in VT.items()}
VR["pe2"] = find(V, "現行前瞻 P/E（資料商 NTM", exact=False)
VR["ev"] = find(V, "現行 EV ÷ FY2026 調整後營業利益", exact=False)
VR["ev2"] = find(V, "現行 EV ÷ FY2026（調整後營業利益 − 股票報酬）", exact=False)
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")
def iv(k, f=""): return f"@In!E{inr(k)}" + (f"|{f}" if f else "")
def il(k): return f"@In!B{inr(k)}"
def nourl(sh, coord, f=""):
    v = wb[sh][coord].value
    return "（出處含網址，見 xlsx）" if isinstance(v, str) and "http" in v else f"@{sh}!{coord}" + (f"|{f}" if f else "")

DB = find(D, "情境 B（基準）")
DK = {k: find(D, t, exact=False) for k, t in dict(ke="ke（股權成本）", e="股權市值 E", wacc="WACC", g="永續成長率 g", ronic="RONIC", nc="淨現金與非營業資產").items()}
DK["wacc"] = find(D, "WACC")
DBR = {k: find(D, t, DB) for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", eq="股權價值", psv="每股價值（評價日）",
                                          d12="DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）", tvsh="終值占 EV", alt="終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）").items()}
DA, DC = find(D, "情境 A（樂觀）"), find(D, "情境 C（保守）")
def dsc(k):
    t = D.cell(DBR[k], 1).value
    return [f"@DCF!B{find(D, t, DA)}", f"@DCF!B{DBR[k]}", f"@DCF!B{find(D, t, DC)}"]

# ---------------- 敏感度 ----------------
s_rows = [r for r in range(5, SE.max_row + 1) if isinstance(SE.cell(r, 5).value, (int, float))]
lo = [r for r in s_rows if str(SE.cell(r, 1).value).endswith("（低）")]
hi = [r for r in s_rows if str(SE.cell(r, 1).value).endswith("（高）")]
oth = [r for r in s_rows if r not in lo and r not in hi]
s_end = max(s_rows)
r_bid, r_st = find(SE, "build_id"), find(SE, "模型狀態")
TOR = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": "E", "fmt": "sp1"}
def big(rs): return [r for r in rs if abs(SE.cell(r, 5).value) >= 0.1]
def small(rs): return [r for r in rs if abs(SE.cell(r, 5).value) < 0.1]
def main_tor(side, rs):
    if big(rs):
        return dict(TOR, title=f"加權目標價相對基準：輸入取{side}值（變動達兩位數百分比的案例）", rows=",".join(map(str, big(rs))))
    return dict(TOR, title=f"加權目標價相對基準：輸入取{side}值（沒有變動達兩位數百分比的案例；列出影響最大者）", rows=",".join(map(str, sorted(rs, key=lambda r: -abs(SE.cell(r, 5).value))[:1])))
sens_blocks = [main_tor("低", lo), main_tor("高", hi),
               {"type": "note", "parts": ["@Sens!A1"]},
               {"type": "kv", "items": [["@Sens!A2", "@Sens!B2|n2"], ["@Sens!C2", "@Sens!D2|n2"], [f"@Sens!A{r_st}", f"@Sens!D{r_st}"], [f"@Sens!A{r_bid}", f"@Sens!D{r_bid}"]]}]
if oth:
    sens_blocks.append(dict(TOR, title="加權目標價相對基準：只看單一情境、改用現行相對倍數、終值替代法", rows=",".join(map(str, oth))))
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

# ---------------- 橫幅 ----------------
hero = {"rec": vv("rec"),
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n2"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n2"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n2"), ""],
                  [il("px"), iv("px", "n2"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等", vv("rt"), "；收支關卡", vv("gj"), "。扣除股票報酬的現金流折現低於現價，同業調整後本益比高於現價；評等建立在後者的口徑上。距評等邊界 ",
                 vv("dist", "B", "n1"), " 個百分點。"]}

# ---------------- 命題 ----------------
thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "調整後營業利益、股票報酬與扣除後的營業利益（百萬美元）：FY2026–FY2032（情境 B）", "x": f"@PL!F{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "F"), "color": c} for k, c in (("aoi", "#0b7a80"), ("sbc", "#c29a2b"), ("ebit", "#3b3838"))],
     "lines": [{"name": [f"@PL!A{B['aom']}"], "y": pl("aom", "B", "F"), "axis": "r", "color": "#b4472f", "labels": True}],
     "caption": "三組長條＝調整後營業利益（公司口徑）、股票報酬、兩者相減（本模型 DCF 用的營業利益）；折線＝調整後營業利益率（右軸）。FY2026 為前三季實績加公司指引隱含的第四季，之後為模型。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["承重變數：EyeQ 出貨相對車產的年變化（情境 B ", iv("pen_B", "sp1"), "）、基礎 ADAS 單價年變化（", iv("asp_B", "sp1"), "）、Surround ADAS 出貨（FY2032 情境 B ", iv("srd_B", "n1"),
         " 百萬套）、股票報酬占營收（", iv("sbc_r", "p1"), "）、同業前瞻本益比中位數（", iv("pe_med", "n2"), " 倍）。"]]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A4", " ", "@README!A5"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n2"), " 美元（DCF ", vv("dcf", "B", "n2"), "、Comparable ", vv("cmp", "B", "n2"), "），對現價 ", iv("px", "n2"), " 美元上檔 ", vv("up", "B", "sp1"),
         "，估值評等", vv("rt"), "、投資建議", vv("rec"), "。"],
        ["兩種讀法：現行企業價值約為調整後營業利益的 ", vv("ev", "B", "n1"), " 倍，但約為扣除股票報酬後營業利益的 ", vv("ev2", "B", "n0"), " 倍。DCF 扣股票報酬，Comparable 用同業調整後本益比，兩法差距就是爭議所在。"],
        ["營收結構：FY2026 合併營收 ", f"@PL!F{B['rev']}|n0", " 百萬美元，基礎 ADAS ", f"@PL!F{B['base']}|n0", "；情境 B 到 FY2032，Surround ADAS ", f"@PL!L{B['srd']}|n0",
         "、SuperVision 與 Chauffeur ", f"@PL!L{B['sv']}|n0", "、Drive ", f"@PL!L{B['drv']}|n0", "，基礎 ADAS 降到 ", f"@PL!L{B['base']}|n0", "。"],
        ["利潤來源：FY2026 調整後營業利益 ", f"@PL!F{B['aoi']}|n0", " 百萬美元中，以色列研發補助約 ", f"@PL!F{B['inc']}|n0", "；股票報酬 ", f"@PL!F{B['sbc']}|n0", "。"],
        "@README!A6"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A7"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}, {"label": il("cons_tp"), "value": iv("cons_tp"), "series": 1}],
     "ref": {"label": il("px"), "value": iv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "note", "parts": ["@README!A5"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n2"), ("cmp", "n2"), ("tp", "n2"), ("px", "n2"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("dist", "n1"), ("bnd", ""))]
             + [[il("cons_tp"), iv("cons_tp", "n2")]]},
    {"type": "table", "caption": "情境目標價（美元；情境權重見下方）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", "n2"]}},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（百萬美元；每股為美元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n0"), ("eq", "n0"), ("psv", "n2"), ("d12", "n2"), ("tvsh", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n0"), ("tv", "n0"), ("pvtv", "n0"), ("alt", "n2"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n3"), ("erp", "p1"), ("crp", "p2"), ("sh_dcf", "n1"), ("sh_cmp", "n1"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n0", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n0", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因與現行倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("pd2", "n2"), ("pe", "x1"), ("pe2", "n2"), ("ev", "n1"), ("ev2", "n0"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻 P/E 與中位數", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}")] for k in ("pe_qcom", "pe_hsai", "pe_amba", "pe_denso", "pe_med", "pe_self")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("cmp_sbc", "n0"), ("t_tgt", "n2"))]}]}

revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價（上檔為零）及落在評等邊界所需的值，其他輸入不變。「無解」表示在搜尋區間內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{find(RC, '反解（整本重算', exact=False)}", "head": [f"@Recon!{c}{find(RC, '反解目標')}" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{find(RC, '反解目標') + 1}-{RC.max_row}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "", "n4", "n4", ""]}}]}

# ---------------- 損益 ----------------
def cellref(k, c, f):
    v = P[f"{c}{B[k]}"].value
    return f"@PL!{c}{B[k]}|{f}" if v is not None else "未揭露"
STDROWS = [("rev", "n0"), ("gr", "p1"), ("cogs", "n0"), ("gp", "n0"), ("gm", "p1"), ("opex", "n0"), ("op", "n0"), ("opm", "p1"), ("fin", "n0"), ("pbt", "n0"),
           ("tax", "n0"), ("ni", "n0"), ("nim", "p1"), ("eps", "n2"), ("nig", "p1")]
NGROWS = [("agp", "n0"), ("agm", "p1"), ("aop", "n0"), ("inc", "n0"), ("aoi", "n0"), ("aom", "p1"), ("sbc", "n0"), ("am", "n0"), ("oit", "n0"), ("ani", "n0"), ("sh", "n1"), ("aeps", "n2")]
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "合併營收（百萬美元）× 情境 A／B／C 與調整後營業利益率（情境 B）：FY2023–FY2032", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：合併營收"], "y": pl("rev", s), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['aom']}", "（情境 B）"], "y": pl("aom"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境合併營收；折線＝情境 B 調整後營業利益率（右軸）。FY2023–FY2025 實績、FY2026 為實績加指引，三情境相同，分化從 FY2027 開始。"},
    {"type": "note", "parts": ["會計口徑：標準損益表用 GAAP（美國一般公認會計原則），因為三年歷史完整、且 DCF 扣除股票報酬；市場與公司指引主要用 Non-GAAP（調整後，不含股票報酬、併購無形資產攤銷與商譽減損），列在下一張表。FY2024 與 FY2026 的 GAAP 虧損主要來自商譽減損。"]},
    {"type": "table", "caption": "標準損益表（GAAP；百萬美元，EPS 為美元）：FY2023–FY2025 實績、FY2026 實績加指引、FY2027 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in YC],
     "rows": [[f"@PL!A{B[k]}"] + [cellref(k, c, f) for c in YC] for k, f in STDROWS]},
    {"type": "table", "caption": "Non-GAAP（調整後）與 GAAP 調節項（百萬美元）：FY2023 調整後各列未揭露（本專案未取得）",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in YC],
     "rows": [[f"@PL!A{B[k]}"] + [cellref(k, c, f) for c in YC] for k, f in NGROWS]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "合併營收（百萬美元）", "fmt": "n0", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "調整後營業利益（百萬美元）", "fmt": "n0", "chart": True, "rows": {s: pl("aoi", s, "F") for s in "ABC"}},
                 {"name": "GAAP 營業利益（百萬美元）", "fmt": "n0", "rows": {s: pl("op", s, "G") for s in "ABC"}},
                 {"name": "調整後 EPS（美元）", "fmt": "n2", "chart": True, "rows": {s: pl("aeps", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣股利（百萬美元；收支關卡用）", "fmt": "n0", "rows": {s: pl("fcfad", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流（百萬美元）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in FC],
     "rowgen": {"sheet": "PL", "rows": f"{B['ebit']}-{B['fcfad']}", "cols": ["A"] + list(FC), "fmts": [""] + ["n0"] * len(FC)}}]}

drivers = {"id": "drivers", "title": "驅動因子：出貨 × 單價 → Mobileye 分部子列 → 合併營收", "blocks": [
    {"type": "theme", "title": "Mobileye 分部營收依產品堆疊（百萬美元）與 EyeQ 系統出貨（百萬套）：FY2026–FY2032（情境 B）", "x": f"@PL!F{B['hdr']}:L{B['hdr']}", "fmt": "n0", "fmt_r": "n0", "stack": True, "pad": 1.3,
     "total": pl("mbly", "B", "F"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "F"), "color": c} for k, c in (("base", "#0b7a80"), ("srd", "#6fc7cb"), ("sv", "#c29a2b"), ("drv", "#b4472f"), ("mo", "#8a8c8c"))],
     "lines": [{"name": [f"@PL!A{B['u']}"], "y": pl("u", "B", "F"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝Mobileye 分部五條子列（合計＝分部營收）；折線＝EyeQ 系統出貨（右軸）。高階產品的出貨與單價只有公司說法，屬假設值。"},
    {"type": "note", "parts": ["@PL!A1"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("pen", "sp1"), ("asp", "sp1"), ("srd", "n1"), ("sv", "n2"), ("rt27", "n0"), ("rt30", "n0"), ("rt35", "n0"), ("dshr", "p0"))]},
    {"type": "table", "caption": "分部營收推導（情境 B；出貨百萬套、單價美元、營收百萬美元）：出貨 × 單價 → 五條子列 → Mobileye 分部 ＋ Other → 合併營收；核對列應為零",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in FC],
     "rowgen": {"sheet": "PL", "rows": f"{B['u']}-{B['gr']}", "cols": ["A"] + list(FC), "fmts": [""] + ["n1"] * len(FC), "row_fmt": {str(B["gr"]): {c: "p1" for c in FC}, str(B["rt"]): {c: "n0" for c in FC}, str(B["u_drv"]): {c: "n0" for c in FC}, str(B["u_sv"]): {c: "n3" for c in FC}}}},
    {"type": "table", "caption": "調整後利潤推導（情境 B；百萬美元）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in FC],
     "rowgen": {"sheet": "PL", "rows": f"{B['agp']}-{B['am']}", "cols": ["A"] + list(FC), "fmts": [""] + ["n0"] * len(FC),
                "row_fmt": {str(B["agm"]): {c: "p1" for c in FC}, str(B["aom"]): {c: "p1" for c in FC}}}},
    {"type": "h3", "text": "一、基礎 ADAS 與 Surround ADAS（命題模組）"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("lvp", "p1"), ("pen_B", "sp1"), ("asp_B", "sp1"), ("srd_B", "n1"), ("srd_asp", "n0"), ("cust_loss", "p0"), ("g_anc", "p0"))]},
    {"type": "h3", "text": "二、SuperVision／Chauffeur、Drive（Robotaxi）與其他營收"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("sv_B", "n2"), ("sv_u26", "n3"), ("sv_asp", "n0"), ("gm_sv", "p0"), ("drv_asp", "n0"), ("dshr_B", "p0"), ("geo_m", "n1"), ("gm_drv", "p0"), ("mo_g", "p1"), ("oth_g", "p1"), ("gm_oth", "p0"))]},
    {"type": "h3", "text": "三、全公司（費用、補助、股票報酬、現金流、折現）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("opex_g", "p1"), ("inc_k", "n2"), ("sbc_r", "p1"), ("fin_y", "p1"), ("tax_w", "p1"), ("da_r", "p1"), ("cx_r", "p1"), ("wc_r", "p1"), ("recv_k", "n2"),
               ("crp", "p2"), ("beta", "n3"), ("erp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("w_dcf", "p0"), ("cmp_sbc", "n0"))]},
    {"type": "table", "collapse": True, "summary": "展開：前三季實績、公司指引與校準值（固定輸入）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, "n1"), f"@In!G{inr(k)}|tag"] for k in ("n9_rev", "n9_agp", "n9_aoi", "n9_inc", "n9_fin", "n9_tax", "n9_ani", "n9_mbly", "n9_oth", "n9_sys", "n9_u", "q3_asp",
                                                                    "n9_sbc", "n9_am", "n9_amc", "int_bal", "g26_rev", "g26_aoi", "g26_inc", "g26_sbc", "g26_am", "g26_oth", "q4_oth", "q4_rev", "q4_mo",
                                                                    "q4_sys", "q4_u", "sh_26", "core26", "oth26", "agp26", "aop26", "am_rate")]
             + [[il(k), iv(k, "p1"), f"@In!G{inr(k)}|tag"] for k in ("gm_core", "inc_r", "am_cg")]}]}

Q = wb["Q"]
RE = wb["對帳"]
recon = {"id": "recon", "title": "對帳：公司指引、共識、官方分部與線層級變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, '加權目標價 vs 賣方平均目標價')}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n0", "n0", "n0", ""],
                "row_fmt": {str(find(RC, "FY2027 調整後 EPS", exact=False)): {c: "n2" for c in "BCD"}, str(find(RC, "加權目標價 vs 賣方平均目標價")): {c: "n2" for c in "BCD"}}}},
    {"type": "table", "caption": "官方報告分部（百萬美元；季報分部附註）與分部對帳", "head": [f"@Q!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Q", "rows": "3-10", "cols": list("ABCDE"), "fmts": [""] + ["n0"] * 4}},
    {"type": "table", "caption": "系統出貨與平均售價（季）", "head": [f"@Q!{c}12" for c in "ABCDEF"],
     "rowgen": {"sheet": "Q", "rows": "13-15", "cols": list("ABCDEF"), "fmts": [""] + ["n1"] * 5}},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（Robotaxi 線、智慧車線）", "head": [f"@對帳!{c}2" for c in "ABCDGH"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{RE.max_row}", "cols": ["A", "B", "C", "D", "G", "H"], "fmts": ["", "", "", "", "tag", ""]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", f"@Sens!D{r_bid}"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價與倍數的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!I3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag", nourl("In", f"I{inr(k)}")] for k, f in (("px", "n2"), ("rf", "p2"), ("pe_self", "n2"), ("cons_tp", "n2"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"In": ["B", "I"], "Hist": ["A"], "對帳": ["A", "C"]}, "empty": "沒有標為未取得的項目。"}]}

cfg = {"company": "Mobileye", "ticker": "MBLY", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "以出貨、單價與高階產品推得的營收，在扣除股票報酬與不扣除兩種讀法下，能否支持高於現價的目標價",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：滲透提升、單價守住、高階產品放量）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：中國與自研擠壓、單價續降）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

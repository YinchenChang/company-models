# 產生 physical_ai/nabtesco/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；模型改版移列後重跑本檔即可。設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/nabtesco/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/nabtesco/html.json
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

# PL 三情境區塊：「情境 X（…）」標題列之後
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm); hdr = h + 1
    lab = lambda t, h=h: find(P, t, h)
    PR[s] = dict(hdr=hdr, rob=lab("精密減速機－機器人用（子列）"), nrob=lab("精密減速機－非機器人用（子列）"), pai=lab("精密減速機－RV mini／實體 AI 新品（子列）"),
                 prg=lab("精密減速機（CMP 子列）"), cmpo=lab("CMP 其他（子列）"), cmp=lab("コンポーネントソリューション（CMP）營收"), trs=lab("トランスポートソリューション（TRS）營收"),
                 acb=lab("アクセシビリティソリューション（ACB）營收"), mfr=lab("その他（MFR）營收"), rev=lab("合併營收"), chk1=lab("檢查：CMP 子列合計 − CMP"), chk2=lab("檢查：分部合計 − 合併營收"),
                 yoy=lab("合併營收年增"), cmpop=lab("CMP 營業利益"), trsop=lab("TRS 營業利益"), acbop=lab("ACB 營業利益"), mfrop=lab("MFR 營業利益"), corp=lab("全社費用與消去"),
                 op=lab("營業利益"), om=lab("營業利益率"), cmpm=lab("CMP 營業利益率"), nonop=lab("營業外"), pti=lab("稅前利益"), tax=lab("所得稅"), ni_all=lab("稅後淨利"),
                 nci=lab("少數股東損益"), ni=lab("歸屬母公司淨利"), sh=lab("平均股數（百萬股）"), eps=lab("EPS（日圓）"), da=lab("折舊攤銷"), cx=lab("資本支出"), wc=lab("營運資金增加"),
                 fcff=lab("FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）"), div=lab("股利"), fcfad=lab("自由現金流扣股利（收支關卡用）"))
B = PR["B"]
YC = "CDEFGHIJ"   # PL：C＝FY2025 … J＝FY2032
def pl(k, s="B", c1="C", c2="J"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VR = {k: find(V, t) for k, t in dict(dcfr="DCF 情境價值推到 12 個月後", cu=None or "Comparable（未調整）：分部本益比加總（FY2028 分部營業利益 × 稅後 × 群組前瞻 P/E）",
                                       ca="Comparable（相對倍數調整後，進加權）", tps="各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", tp="加權目標價（日圓）",
                                       dcf="DCF 目標價（情境加權）", cmp="Comparable 目標價（情境加權，進加權）", px="現價（In 頁 px）", up="上檔空間", rt="估值評等（只看上檔空間）",
                                       gate="收支關卡：FY2026–FY2030 自由現金流扣股利合計（¥B；情境 A／B／C）", gj="入不敷出判定", rec="投資建議", dist="距最近級距邊界（百分點）",
                                       bnd="翻轉目標價：升一級／降一級", a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻",
                                       pd2="期間：持有（股價日 → 目標價時點，年）", wk="WACC／ke／g／RONIC", pe="隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 EPS").items()}
VR["pe2"] = find(V, "現行前瞻 P/E（資料商 NTM", exact=False)
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")
def iv(k, f=""): return f"@In!E{inr(k)}" + (f"|{f}" if f else "")
def il(k): return f"@In!B{inr(k)}"
# 含網址的儲存格不進 HTML（自含規則：不得有外部網址）
def nourl(sh, coord, f=""):
    v = wb[sh][coord].value
    return "（出處含網址，見 xlsx）" if isinstance(v, str) and "http" in v else f"@{sh}!{coord}" + (f"|{f}" if f else "")

DB = find(D, "情境 B（基準）")
DK = {k: find(D, t) for k, t in dict(ke="ke（股權成本）", e="股權市值 E（¥B）", dd="負債 D（借款＋租賃）", kd="稅後 kd", wacc="WACC", g="永續成長率 g＝MIN(g_cap, rf)", ronic="RONIC",
                                       nc="淨現金與投資（資產負債表 2026-09-30）＝現金 − 借款 − 租賃 − 非支配 ＋ 持分法投資 × inv_k").items()}
DBR = {k: find(D, t, DB) for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", eq="股權價值", psv="每股價值（評價日）",
                                          d12="DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）", tvsh="終值占 EV", alt="終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）").items()}
DA, DC = find(D, "情境 A（樂觀）"), find(D, "情境 C（保守）")
def dsc(k):   # 三情境同一列（情境 A／B／C 區塊各自的列）
    t = D.cell(DBR[k], 1).value
    return [f"@DCF!B{find(D, t, DA)}", f"@DCF!B{DBR[k]}", f"@DCF!B{find(D, t, DC)}"]

# Sens：第 5 列起；A 欄白話標籤以「（低）」「（高）」結尾者為單一參數低／高案，其餘為情境權重與終值替代法
s_rows = [r for r in range(5, SE.max_row + 1) if isinstance(SE.cell(r, 5).value, (int, float))]
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
               {"type": "kv", "items": [["@Sens!A2", "@Sens!B2|n0"], ["@Sens!C2", "@Sens!D2|n0"], ["模型狀態", "@Sens!E2"], ["build_id", "@Sens!H2"]]}]
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
                    "rowgen": {"sheet": "Sens", "rows": f"5-{s_end}", "cols": ["A", "B", "D", "E", "F", "G", "H"], "fmts": ["", "", "n0", "sp1", "n0", "sp1", ""]}})
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}

hero = {"rec": vv("rec"),
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n0"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n0"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n0"), ""],
                  [il("px"), iv("px", "n0"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等 ", vv("rt"), "（只看上檔空間）；收支關卡", vv("gj"), "（情境 B 自由現金流扣股利，FY2026–FY2030 合計 ", vv("gate", "C", "n1"),
                 " 十億日圓）。現價為第三季決算後回落的收盤價；目標價幾乎不受現價影響，評等由股價下跌帶動。"]}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "合併營收依官方分部堆疊（十億日圓）與營業利益率：FY2025–FY2032（情境 B）", "x": f"@PL!C{B['hdr']}:J{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k), "color": c} for k, c in (("cmp", "#0b7a80"), ("trs", "#6fa8dc"), ("acb", "#c29a2b"), ("mfr", "#8a8c8c"))],
     "lines": [{"name": [f"@PL!A{B['om']}"], "y": pl("om"), "axis": "r", "color": "#3b3838", "labels": True},
               {"name": [f"@PL!A{B['cmpm']}"], "y": pl("cmpm"), "axis": "r", "color": "#b4472f", "dash": True}],
     "caption": "堆疊＝四個官方報告分部（合計等於合併營收）；精密減速機所在的 CMP 分部只是其中一塊。實線＝合併營業利益率、虛線＝CMP 分部營業利益率（右軸）。FY2025 為實績、FY2026 為公司指引、FY2027 起為模型。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["核心命題的承重變數：全球工業機器人需求年增（情境 B：FY2027 ", iv("rg27_B", "p1"), "、FY2028–FY2029 每年 ", iv("rgm_B", "p1"), "）、機器人用 RV 份額＋價格年變化（", iv("shr_B", "p1"),
         "）、CMP 邊際營業利益率（", iv("mm_cmp", "p0"), "）。"]]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n0"), " 日圓（DCF ", vv("dcf", "B", "n0"), "、分部本益比加總 ", vv("cmp", "B", "n0"), "），對現價 ", iv("px", "n0"), " 日圓上檔 ", vv("up", "B", "sp1"),
         "，估值評等", vv("rt"), "、投資建議", vv("rec"), "。"],
        ["現價變動：", il("px_prev"), " ", iv("px_prev", "n0"), " 日圓 → ", il("px"), " ", iv("px", "n0"), " 日圓；目標價只經評價日股數與資本結構權重受現價影響，幅度小，評等變化主要來自股價下跌。"],
        ["營收結構：FY2026 公司指引合併營收 ", f"@PL!D{B['rev']}|n1", " 十億日圓，其中精密減速機 ", f"@PL!D{B['prg']}|n1", "、CMP 分部 ", f"@PL!D{B['cmp']}|n1",
         "；其餘為鐵道與航空（TRS）", f"@PL!D{B['trs']}|n1", "、自動門（ACB）", f"@PL!D{B['acb']}|n1", "、包裝機等（MFR）", f"@PL!D{B['mfr']}|n1", "。機器人循環只直接作用在 CMP。"],
        ["評等邊界：", f"@README!A5"],
        "@README!A6"]},
    {"type": "h3", "text": "範圍"},
    {"type": "note", "parts": ["@README!A7"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（日圓；豎線＝現價）", "fmt": "n0",
     "items": [{"label": f"@Val!A{VR['dcf']}", "value": vv("dcf"), "series": 1}, {"label": f"@Val!A{VR['cmp']}", "value": vv("cmp"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": vv("tp"), "series": 0}, {"label": il("cons_tp"), "value": iv("cons_tp"), "series": 1}],
     "ref": {"label": il("px"), "value": iv("px")}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n0"), ("cmp", "n0"), ("tp", "n0"), ("px", "n0"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("dist", "n1"), ("bnd", ""))]
             + [[il("cons_tp"), iv("cons_tp", "n0")], [il("px_prev"), iv("px_prev", "n0")]]},
    {"type": "table", "caption": "情境目標價（日圓；情境權重見下方折現率與權重）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n0", "n0", "n0", "n0"]}},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（十億日圓；每股為日圓）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n1"), ("eq", "n1"), ("psv", "n0"), ("d12", "n0"), ("tvsh", "p1"))]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n1"), ("tv", "n1"), ("pvtv", "n1"), ("alt", "n0"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n3"), ("erp", "p1"), ("crp", "p2"), ("sh_v", "n2"), ("bb_rem", "n2"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n1", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n1", "", ""], [f"@DCF!A{DK['dd']}", f"@DCF!B{DK['dd']}|n1", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因與隱含倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("pd2", "n2"), ("pe", "x1"), ("pe2", "n2"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻 P/E 與群組中位數", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}")] for k in ("pe_fanuc", "pe_yask", "pe_hds", "pe_thk", "pe_smc", "pe_nidec", "pe_kbx", "pe_khi", "pe_assa", "pe_doka",
                                                                                     "pe_cmp", "pe_trs", "pe_acb", "pe_all", "pe_self")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"), ("dps_tgt", "n0"))]}]}

revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價（上檔為零）及落在評等邊界所需的值，其他輸入不變。「無解」表示在搜尋區間內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{find(RC, '反解（整本重算', exact=False)}", "head": [f"@Recon!{c}{find(RC, '反解目標')}" for c in "ACDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{find(RC, '反解目標') + 1}-{RC.max_row}", "cols": ["A", "C", "D", "E"], "fmts": ["", "p1", "p1", ""]}}]}

STD = list(YC)
def prow(lab, k, f):
    return [lab] + [f"@PL!{c}{B[k]}|{f}" for c in STD]
HIST = find(wb["Hist"], "營收"); HISTO = find(wb["Hist"], "營業利益")
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "合併營收（十億日圓）× 情境 A／B／C 與營業利益率（情境 B）：FY2025–FY2032", "x": f"@PL!C{B['hdr']}:J{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：合併營收"], "y": pl("rev", s), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['om']}", "（情境 B）"], "y": pl("om"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境合併營收；折線＝情境 B 營業利益率（右軸）。FY2025 實績、FY2026 公司指引三情境相同，分化從 FY2027 開始。"},
    {"type": "note", "parts": ["會計口徑：IFRS（公司法定財報準則，持續營業口徑；油壓機器事業列停業單位）。本表為法定財報口徑，即 GAAP 口徑，未做 Non-GAAP 調整。模型以分部營業利益相加推到合併營業利益，未拆營業成本、毛利與營業費用；FY2025 稅前以下只取得歸屬母公司淨利。"]},
    {"type": "table", "caption": "標準損益表（IFRS／GAAP 口徑；十億日圓，EPS 為日圓）：FY2024–FY2025 實績、FY2026 公司指引、FY2027 起情境 B",
     "head": [f"@PL!A{B['hdr']}", "@Hist!B2"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rows": [["營收（合併）", f"@Hist!B{HIST}|n1"] + [f"@PL!{c}{B['rev']}|n1" for c in STD],
              [f"@PL!A{B['yoy']}", "—"] + [f"@PL!{c}{B['yoy']}|p1" for c in STD],
              *[[lab, "未建模"] + ["未建模"] * len(STD) for lab in ("營業成本（未建模：模型以分部營業利益相加）", "毛利（未建模）", "毛利率（未建模）", "營業費用（未建模）")],
              [f"@PL!A{B['op']}", f"@Hist!B{HISTO}|n1"] + [f"@PL!{c}{B['op']}|n1" for c in STD],
              [f"@PL!A{B['om']}", "—"] + [f"@PL!{c}{B['om']}|p1" for c in STD],
              ["其他收入淨額（營業外：持分法與金融損益）", "—"] + [f"@PL!{c}{B['nonop']}|n1" for c in STD],
              [f"@PL!A{B['pti']}", "—"] + [f"@PL!{c}{B['pti']}|n1" for c in STD],
              [f"@PL!A{B['tax']}", "—"] + [f"@PL!{c}{B['tax']}|n1" for c in STD],
              [f"@PL!A{B['ni_all']}", "—"] + [f"@PL!{c}{B['ni_all']}|n1" for c in STD],
              [f"@PL!A{B['nci']}", "—"] + [f"@PL!{c}{B['nci']}|n1" for c in STD],
              [f"@PL!A{B['ni']}", "—"] + [f"@PL!{c}{B['ni']}|n1" for c in STD],
              [f"@PL!A{B['sh']}", "—"] + [f"@PL!{c}{B['sh']}|n2" for c in STD],
              ["EPS（基本，日圓；本公司無稀釋性證券揭露）", "—"] + [f"@PL!{c}{B['eps']}|n0" for c in STD],
              *[[lab, "未建模"] + ["未建模"] * len(STD) for lab in ("淨利率（未建模）", "淨利年增率（未建模）")]]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!D{B['hdr']}:J{B['hdr']}",
     "metrics": [{"name": "合併營收（十億日圓）", "fmt": "n1", "chart": True, "rows": {s: pl("rev", s, "D") for s in "ABC"}},
                 {"name": "營業利益（十億日圓）", "fmt": "n1", "chart": True, "rows": {s: pl("op", s, "D") for s in "ABC"}},
                 {"name": "營業利益率", "fmt": "p1", "rows": {s: pl("om", s, "D") for s in "ABC"}},
                 {"name": "EPS（日圓）", "fmt": "n0", "chart": True, "rows": {s: pl("eps", s, "D") for s in "ABC"}},
                 {"name": "自由現金流扣股利（十億日圓；收支關卡用）", "fmt": "n1", "rows": {s: pl("fcfad", s, "D") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流（十億日圓）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['da']}-{B['fcfad']}", "cols": ["A"] + STD, "fmts": [""] + ["n1"] * len(STD)}}]}

drivers = {"id": "drivers", "title": "驅動因子：精密減速機 → 官方分部 → 合併營收", "blocks": [
    {"type": "theme", "title": "精密減速機營收依用途堆疊（十億日圓）與 CMP 營業利益率：FY2026–FY2032（情境 B）", "x": f"@PL!D{B['hdr']}:J{B['hdr']}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("prg", "B", "D"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k, "B", "D"), "color": c} for k, c in (("rob", "#0b7a80"), ("nrob", "#6fc7cb"), ("pai", "#c29a2b"))],
     "lines": [{"name": [f"@PL!A{B['cmpm']}"], "y": pl("cmpm", "B", "D"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝精密減速機三條子列（機器人用、非機器人用、RV mini 等新品；合計＝精密減速機營收）；折線＝CMP 分部營業利益率（右軸）。機器人用比例為假設值，公司未揭露用途別營收。"},
    {"type": "note", "parts": [f"@PL!A{find(P, 'PL — ', exact=False)}"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("rg27", "p1"), ("rgm", "p1"), ("rglt", "p1"), ("shr", "p1"), ("nr", "p1"), ("pai", "n1"), ("madj", "p1"), ("gtrs", "p1"), ("gacb", "p1"), ("gmfr", "p1"))]},
    {"type": "table", "caption": "分部營收推導（情境 B；十億日圓）：精密減速機三子列 → CMP → 四分部 → 合併營收；核對列應為零", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['rob']}-{B['yoy']}", "cols": ["A"] + STD, "fmts": [""] + ["n1"] * len(STD), "row_fmt": {str(B["yoy"]): {c: "p1" for c in STD}}}},
    {"type": "table", "caption": "分部營業利益推導（情境 B；十億日圓）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in STD],
     "rowgen": {"sheet": "PL", "rows": f"{B['cmpop']}-{B['cmpm']}", "cols": ["A"] + STD, "fmts": [""] + ["n1"] * len(STD),
                "row_fmt": {str(B["om"]): {c: "p1" for c in STD}, str(B["cmpm"]): {c: "p1" for c in STD}}}},
    {"type": "h3", "text": "一、精密減速機（CMP 分部；命題模組）"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("r_rob", "p0"), ("rg27_B", "p1"), ("rgm_B", "p1"), ("rglt_B", "p1"), ("shr_B", "p1"), ("nr_B", "p1"), ("pai_B", "n1"), ("mm_cmp", "p0"), ("madj_B", "p1"), ("g_anc", "p0"))]},
    {"type": "h3", "text": "二、其他官方分部（TRS、ACB、MFR）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("gtrs_B", "p1"), ("gacb_B", "p1"), ("gmfr_B", "p1"), ("m_trs", "p1"), ("m_acb", "p1"), ("m_mfr", "p1"))]},
    {"type": "h3", "text": "三、全公司（損益下半段、現金流、折現）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in
              (("nonop", "n1"), ("tax_w", "p0"), ("fc_da", "p1"), ("fc_cx", "p1"), ("fc_wc", "p0"), ("beta", "n3"), ("erp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("inv_k", "p0"), ("bb_rem", "n2"), ("w_dcf", "p0"))]},
    {"type": "table", "collapse": True, "summary": "展開：FY2025 實績與 FY2026 公司指引（固定輸入）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, "n1"), f"@In!G{inr(k)}|tag"] for k in ("a25_rev", "a25_op", "a25_ni", "a25_cmp", "a25_trs", "a25_acb", "a25_mfr", "g26_rev", "g26_op", "g26_ni", "g26_prg", "g26_cmpo",
                                                                    "g26_cmpop", "g26_trs", "g26_trsop", "g26_acb", "g26_acbop", "g26_mfr", "g26_mfrop", "corp_26", "nonop_26", "pai_0", "cx_26")]
             + [[il("g26_eps"), iv("g26_eps", "n1"), f"@In!G{inr('g26_eps')}|tag"], [il("sh_26"), iv("sh_26", "n2"), f"@In!G{inr('sh_26')}|tag"]]}]}

Q = wb["Q"]; qs = find(Q, "單季（Derived）"); qe = find(Q, "檢查：分部合計 − 合併營收", qs)
RE = wb["對帳"]
recon = {"id": "recon", "title": "對帳：公司指引、共識、季度分部與線層級變數", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{find(RC, '加權目標價 vs 賣方平均目標價')}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1", ""],
                "row_fmt": {str(find(RC, "FY2027 EPS", exact=False)): {c: "n0" for c in "BCD"}, str(find(RC, "加權目標價 vs 賣方平均目標價")): {c: "n0" for c in "BCD"}}}},
    {"type": "table", "caption": "季度分部營收（單季，十億日圓；累計原文相減）與分部對帳", "head": [f"@Q!{c}{qs}" for c in "ABCDEFGH"],
     "rowgen": {"sheet": "Q", "rows": f"{qs + 1}-{qe}", "cols": list("ABCDEFGH"), "fmts": [""] + ["n2"] * 7}},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（工業機器人、人形機器人）", "head": [f"@對帳!{c}2" for c in "ABCDG"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{RE.max_row}", "cols": ["A", "B", "C", "D", "G"], "fmts": ["", "", "", "", "tag"]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", "@Sens!H2"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價與同業倍數的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3", "@In!I3"],
     "rows": [[il("px"), iv("px", "n0"), f"@In!G{inr('px')}|tag", nourl("In", f"H{inr('px')}"), nourl("In", f"I{inr('px')}")],
              [il("px_prev"), iv("px_prev", "n0"), f"@In!G{inr('px_prev')}|tag", nourl("In", f"H{inr('px_prev')}"), nourl("In", f"I{inr('px_prev')}")],
              [il("rf"), iv("rf", "p2"), f"@In!G{inr('rf')}|tag", nourl("In", f"H{inr('rf')}"), nourl("In", f"I{inr('rf')}")],
              [il("pe_self"), iv("pe_self", "n2"), f"@In!G{inr('pe_self')}|tag", nourl("In", f"H{inr('pe_self')}"), nourl("In", f"I{inr('pe_self')}")]]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"In": ["B"], "Hist": ["A"], "Q": ["A"], "對帳": ["A", "C"]}, "empty": "沒有標為未取得的項目。"}]}

cfg = {"company": "Nabtesco（ナブテスコ）", "ticker": "6268.T", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "工業機器人循環帶動的精密減速機（RV）成長，在三情境下能否支持高於現價的目標價；RV 對人形機器人的直接受益有多大",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：機器人需求高成長、份額不流失）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：需求回落、中國 RV 廠替代）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

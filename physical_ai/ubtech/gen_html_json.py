# 產生 physical_ai/ubtech/html.json（個股 HTML 版面規格 v1.3；tools/company_html/build_company_html.py 讀取）。
# 列號一律依 xlsx 的標籤（A 欄文字或 In 頁鍵）找，不寫死；設定檔內不寫數字（全部為儲存格參照）。
# 用法：python3 physical_ai/ubtech/gen_html_json.py [成品 xlsx]（預設 dist/ 內唯一的 xlsx）→ physical_ai/ubtech/html.json
import glob, json, os, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
X = sys.argv[1] if len(sys.argv) > 1 else sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))[0]
wb = openpyxl.load_workbook(X, data_only=True)
P, V, I, D, SE, RC, RE = wb["PL"], wb["Val"], wb["In"], wb["DCF"], wb["Sens"], wb["Recon"], wb["對帳"]


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

KEYS = dict(u="全尺寸人形：銷量（台）", glob="全尺寸人形：全球人形出貨（台；線模型）", shr="全尺寸人形：占全球出貨比例", pidx="全尺寸人形：售價指數（線模型；2026＝1）",
            asp="全尺寸人形：每台營收（千元）", fs="全尺寸具身智能人形機器人產品及服務 營收", ne="非具身智能人形機器人產品及服務 營收", oth="其他智能機器人產品及服務 營收",
            hw="其他智能硬件設備 營收", fl="園林機械、汽車及液壓零部件 營收", ot="其他 營收", rev="營收（合併）", chk="檢查：分部合計 − 合併營收", gr="營收年增率",
            gmfs="全尺寸人形毛利率", cogs="營業成本", gp="毛利", gm="毛利率", core="核心營業費用（銷售＋管理＋研發＋其他營業項目）", imp="信用及資產減值損失",
            opx="營業費用（合計）", op="營業利益", opm="營業利益率", nop="業外（財務淨額＋其他收益）", pbt="稅前利益", tax="所得稅", nia="稅後淨利", nci="少數股東損益",
            ni="歸屬母公司淨利", nim="淨利率", nig="淨利年增率", sh="平均股數（百萬股）", eps="EPS（基本＝稀釋，元）", da="折舊攤銷", cx="資本支出", dwc="營運資金增加",
            fcff="FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", div="股利", fcfad="自由現金流扣股利（收支關卡用）")
PR = {}
for s, nm in (("A", "情境 A（樂觀）"), ("B", "情境 B（基準）"), ("C", "情境 C（保守）")):
    h = find(P, nm)
    PR[s] = dict(hdr=h + 1, **{k: find(P, t, h) for k, t in KEYS.items()})
B = PR["B"]
ALL = "CDEFGHIJKL"   # PL：C＝FY2023 … L＝FY2032
FC = "FGHIJKL"       # FY2026 起
def pl(k, s="B", c1="C", c2="L"): return f"@PL!{c1}{PR[s][k]}:{c2}{PR[s][k]}"

VT = dict(dcfr="DCF 情境價值推到 12 個月後（港元）", ev="分部 EV（FY2028 營收 × 群組 EV／營收，¥B）", nct="目標時點淨現金（¥B）＝2026-06-30 淨現金 ＋ Stub FCFF ＋ FY2027 FCFF",
          cu="Comparable（未調整）：分部 EV／營收加總（港元／股）", ca="Comparable（相對倍數調整後，進加權）", tps="各情境目標價（DCF 權重 × DCF ＋ (1 − DCF 權重) × Comparable）",
          tp="加權目標價（港元）", dcf="DCF 目標價（情境加權）", cmp="Comparable 目標價（情境加權，進加權；分部 EV／營收加總）", px="現價（In 頁 px，港元）", up="上檔空間",
          rt="估值評等（只看上檔空間）", gate="收支關卡：FY2026–FY2030 自由現金流扣股利合計（¥B；情境 A／B／C）", gj="入不敷出判定", rec="投資建議",
          cov="收支缺口 ÷ 2026-06-30 淨現金（情境 B；>1＝現金不足以支應，需再融資）", dist="距最近級距邊界（百分點）", bnd="翻轉目標價：升一級／降一級（港元）",
          a1="上檔三段歸因：DCF 貢獻", a2="上檔三段歸因：未調整 Comparable 貢獻", a3="上檔三段歸因：相對倍數調整貢獻", pd2="期間：持有（股價日 → 目標價時點，年）",
          wk="WACC／ke／g／RONIC", evs="隱含 EV／營收：加權目標價隱含 EV ÷ 情境 B FY2028 營收", self="現行 EV／共識 FY2027 營收 ÷ 系統與整機群組中位數")
VR = {k: find(V, t) for k, t in VT.items()}
def vv(k, col="B", f=""): return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")

DB = find(D, "情境 B（基準）"); DA, DC = find(D, "情境 A（樂觀）"), find(D, "情境 C（保守）")
DK = {k: find(D, t) for k, t in dict(ke="ke（股權成本）", e="股權市值 E（¥B；港元市值 ÷ fx）", dd="負債 D（借款）", kd="稅後 kd", wacc="WACC", g="永續成長率 g＝MIN(上限, 無風險利率)", ronic="RONIC",
                                       nc="淨現金（2026-06-30）＝貨幣資金＋交易性金融資產 − 借款 − 少數股東權益（帳面）").items()}
DBR = {k: find(D, t, DB) for k, t in dict(pvx="明確期現值合計", tv="終值（FY2032 年底）", pvtv="終值現值", ev="企業價值 EV", eq="股權價值（¥B）", psv="每股價值（評價日，人民幣元）",
                                          d12="DCF 情境價值推到 12 個月後（港元；×(1＋ke)^t − 期間股利，× fx）", tvsh="終值占 EV").items()}
def dsc(k):
    t = D.cell(DBR[k], 1).value
    return [f"@DCF!B{find(D, t, DA)}", f"@DCF!B{DBR[k]}", f"@DCF!B{find(D, t, DC)}"]

# Sens：表頭第 7 列，第 8 列起；A 欄以「（敏感度低）」「（敏感度高）」結尾者為單一參數低／高案
s_rows = [r for r in range(8, SE.max_row + 1) if isinstance(SE.cell(r, 6).value, (int, float)) and isinstance(SE.cell(r, 1).value, str)]
lo = [r for r in s_rows if SE.cell(r, 1).value.endswith("（敏感度低）")]
hi = [r for r in s_rows if SE.cell(r, 1).value.endswith("（敏感度高）")]
oth = [r for r in s_rows if r not in lo and r not in hi and abs(SE.cell(r, 6).value) >= 0.0005]
s_end = max(s_rows)
TOR = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": "F", "fmt": "sp1"}
def big(rs): return [r for r in rs if abs(SE.cell(r, 6).value) >= 0.1]
def main_tor(side, rs):
    if big(rs):
        return dict(TOR, title=f"加權目標價相對基準：輸入取敏感度{side}值（變動達兩位數百分比的案例）", rows=",".join(map(str, big(rs))))
    return dict(TOR, title=f"加權目標價相對基準：輸入取敏感度{side}值（沒有變動達兩位數百分比的案例；列出影響最大者）", rows=",".join(map(str, sorted(rs, key=lambda r: -abs(SE.cell(r, 6).value))[:1])))
SN = find(SE, "讀法：", exact=False)
sens_blocks = [main_tor("低", lo), main_tor("高", hi), {"type": "note", "parts": ["@Sens!A1"]}, {"type": "note", "parts": [f"@Sens!A{SN}"]},
               {"type": "kv", "items": [["@Sens!A2", "@Sens!D2"], ["@Sens!A3", "@Sens!D3"]]}]
if oth:
    sens_blocks.append(dict(TOR, title="加權目標價相對基準：情境驅動改為情境 A／C 值、同業群組倍數、只看單一情境", rows=",".join(map(str, oth))))
for side, rs in (("低", lo), ("高", hi)):
    shown = set(map(int, main_tor(side, rs)["rows"].split(",")))
    rest = [r for r in rs if r not in shown and abs(SE.cell(r, 6).value) >= 0.0005]
    if rest:
        sens_blocks.append(dict(TOR, title=f"加權目標價相對基準：輸入取敏感度{side}值，其餘案例", rows=",".join(map(str, rest)), collapse=True,
                                summary=f"展開：輸入取敏感度{side}值、變動未達兩位數百分比的案例"))
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：全部案例（加權目標價、DCF 目標價與估值評等）",
                    "head": ["@Sens!A7", "@Sens!B7", "@Sens!E7", "@Sens!F7", "@Sens!G7", "@Sens!H7", "@Sens!I7"],
                    "rowgen": {"sheet": "Sens", "rows": f"8-{s_end}", "cols": ["A", "B", "E", "F", "G", "H", "I"], "fmts": ["", "", "n2", "sp1", "n2", "sp1", ""]}})
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}

hero = {"rec": vv("rec"),
        "tiles": [[f"@Val!A{VR['tp']}", vv("tp", "B", "n2"), ["上檔空間 ", vv("up", "B", "sp1")]],
                  [f"@Val!A{VR['dcf']}", vv("dcf", "B", "n2"), ""],
                  [f"@Val!A{VR['cmp']}", vv("cmp", "B", "n2"), ""],
                  [il("px"), iv("px", "n2"), ["股價基準日 ", iv("px_date")]]],
        "note": ["估值評等", vv("rt"), "；收支關卡", vv("gj"), "（情境 B 自由現金流扣股利，FY2026–FY2030 合計 ", vv("gate", "C", "n2"),
                 " 十億元人民幣）。基準情境的全尺寸人形出貨以已交付量與已公告訂單為錨，明顯低於賣方共識營收；結論建立在出貨占比與毛利率的假設上。"]}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "合併營收依產品類別堆疊（十億元人民幣）與營業利益率：FY2023–FY2032（情境 B）", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n1", "fmt_r": "p0", "stack": True, "pad": 1.3,
     "total": pl("rev"),
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": pl(k), "color": c} for k, c in (("fs", "#0b7a80"), ("ne", "#6fc7cb"), ("oth", "#6fa8dc"), ("hw", "#c29a2b"), ("fl", "#8a8c8c"), ("ot", "#b4b4b4"))],
     "lines": [{"name": [f"@PL!A{B['opm']}"], "y": pl("opm", "B", "D"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝公司財報按產品類別揭露的營收（合計等於合併營收；FY2023 未取得類別拆分，只有合計）；折線＝營業利益率（右軸）。FY2025 前為實績，FY2026 起為模型。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [
        "@README!A3",
        ["承重變數：下半年全尺寸人形銷量（情境 B ", iv("u26h2_B", "n0"), " 台）、FY2027 起占全球人形出貨（情境 B ", iv("shr27_B", "p1"), "）、全尺寸毛利率（由 ", iv("gm_fs26", "p0"),
         " 收斂到 ", iv("gm_fslt_B", "p0"), "）、核心營業費用年增（", iv("gop_B", "p0"), "）。"]]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": ["@README!A4"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", vv("tp", "B", "n2"), " 港元（DCF ", vv("dcf", "B", "n2"), "、分部 EV／營收加總 ", vv("cmp", "B", "n2"), "），對現價 ", iv("px", "n2"), " 港元上檔 ", vv("up", "B", "sp1"),
         "，估值評等", vv("rt"), "；評等邊界：", "@README!A5"],
        ["營收差距來自出貨假設：模型情境 B 的 FY2027 營收 ", f"@PL!G{B['rev']}|n2", " 十億元，賣方共識 ", iv("cons_rev27", "n2"), "；上半年全尺寸人形只交付 ", iv("u_h126", "n0"),
         " 台，下半年假設 ", iv("u26h2_B", "n0"), " 台。"],
        ["倍數不是主因：現價對共識營收的 EV／營收約為系統與整機同業中位數的 ", vv("self", "B", "n2"), " 倍；目標價低主要因為模型營收低於共識，以及非人形業務套用零組件群組倍數。"],
        ["現金：情境 B 的 FY2026–FY2030 自由現金流合計 ", vv("gate", "C", "n2"), " 十億元，為上半年末淨現金的 ", vv("cov", "B", "n1"), " 倍，需要再融資。"],
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
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("dcf", "n2"), ("cmp", "n2"), ("tp", "n2"), ("px", "n2"), ("up", "sp1"), ("rt", ""), ("gj", ""), ("rec", ""), ("cov", "n2"), ("dist", "n1"), ("bnd", ""))]
             + [[il("cons_tp"), iv("cons_tp", "n2")]]},
    {"type": "table", "caption": "情境目標價與 Comparable 組成（港元；EV 與淨現金為十億元人民幣）", "head": ["@Val!A3", "@Val!B3", "@Val!C3", "@Val!D3", "@Val!E3"],
     "rowgen": {"sheet": "Val", "rows": f"{VR['dcfr']}-{VR['tps']}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", "n2"]}},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DK['ke']}", f"@DCF!B{DK['ke']}|p2"], [f"@DCF!A{DK['kd']}", f"@DCF!B{DK['kd']}|p2"], [f"@DCF!A{DK['wacc']}", f"@DCF!B{DK['wacc']}|p2"],
                             [f"@DCF!A{DK['ronic']}", f"@DCF!B{DK['ronic']}|p2"], [f"@DCF!A{DK['g']}", f"@DCF!B{DK['g']}|p2"]]},
    {"type": "table", "caption": "DCF 彙總（十億元人民幣；每股為人民幣元，推到十二個月後為港元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("ev", "n2"), ("eq", "n2"), ("psv", "n2"), ("d12", "n2"), ("tvsh", "p1"))]},
    {"type": "note", "parts": ["DCF 每股價值以零為下限：情境 B、C 的明確期自由現金流為負、終值不足以抵銷時，股權價值為負，依有限責任取零。"]},
    {"type": "table", "collapse": True, "summary": "展開：DCF 其他項目與折現率組成", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@DCF!A{DBR[k]}"] + [c + "|" + f for c in dsc(k)] for k, f in (("pvx", "n2"), ("tv", "n2"), ("pvtv", "n2"))]
             + [[il(k), iv(k, f), "", ""] for k, f in (("rf", "p2"), ("beta", "n2"), ("erp", "p1"), ("crp", "p2"), ("sh_v", "n2"), ("fx", "n3"))]
             + [[f"@DCF!A{DK['nc']}", f"@DCF!B{DK['nc']}|n2", "", ""], [f"@DCF!A{DK['e']}", f"@DCF!B{DK['e']}|n2", "", ""], [f"@DCF!A{DK['dd']}", f"@DCF!B{DK['dd']}|n2", "", ""]]},
    {"type": "table", "caption": "上檔三段歸因與隱含倍數", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", vv(k, "B", f)] for k, f in (("a1", "sp1"), ("a2", "sp1"), ("a3", "sp1"), ("pd2", "n2"), ("evs", "x1"), ("self", "n2"), ("wk", ""))]},
    {"type": "table", "collapse": True, "summary": "展開：同業 EV／營收與群組中位數", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, "n2"), f"@In!G{inr(k)}|tag"] for k in ("evs_2432", "evs_9660", "evs_2590", "evs_2498", "ms_sys", "ms_cmp", "ms_all", "ev_self")]},
    {"type": "table", "collapse": True, "summary": "展開：權重與目標價規則輸入", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in (("w_dcf", "p0"), ("w_A", "p0"), ("w_B", "p0"), ("w_C", "p0"), ("rel_k", "n2"), ("t_tgt", "n2"))]}]}

RV0 = find(RC, "反解目標")
try:
    RU0 = find(RC, "對照：反解占比對應", exact=False)
except KeyError:
    RU0 = None
RV_END = (RU0 - 2) if RU0 else RC.max_row
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（整本重算反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以整本重算的反解代替：單獨調整一個承重驅動，求使加權目標價等於現價、或升一級所需的值，其他輸入不變。「無解」表示在掃描格點內達不到。"]},
    {"type": "table", "caption": f"@Recon!A{RV0 - 1}", "head": [f"@Recon!{c}{RV0}" for c in "ACDE"],
     "rowgen": {"sheet": "Recon", "rows": f"{RV0 + 1}-{RV_END}", "cols": ["A", "C", "D", "E"], "fmts": ["", "p1", "p1", ""]}}]
    + ([{"type": "table", "caption": f"@Recon!A{RU0}", "head": [f"@Recon!{c}{RU0 + 1}" for c in "ABCDE"],
         "rowgen": {"sheet": "Recon", "rows": f"{RU0 + 2}-{RC.max_row}", "cols": list("ABCDE"), "fmts": ["", "n0", "n0", "n2", ""]}}] if RU0 else [])}

HS = wb["Hist"]; HR = find(HS, "營收")
pl_sec = {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "theme", "title": "合併營收（十億元人民幣）× 情境 A／B／C 與營業利益率（情境 B）：FY2023–FY2032", "x": f"@PL!C{B['hdr']}:L{B['hdr']}", "fmt": "n1", "fmt_r": "p0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：合併營收"], "y": pl("rev", s), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['opm']}", "（情境 B）"], "y": pl("opm"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "三組長條＝三情境合併營收；折線＝情境 B 營業利益率（右軸）。實績期三情境相同，分化從 FY2026 下半年開始。"},
    {"type": "note", "parts": ["會計口徑：GAAP（公司法定財報採中國企業會計準則，港交所公告同一口徑），未做 Non-GAAP 調整；公司另揭露的經調整淨虧損不採用。營業利益為本模型口徑＝稅前利益 − 業外（財務淨額與其他收益），稅金及附加、投資收益與減值列在營業費用內。"]},
    {"type": "table", "caption": "標準損益表（GAAP；十億元人民幣，EPS 為人民幣元）：FY2023–FY2025 實績、FY2026 起情境 B",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in ALL],
     "rows": [["營收（合併）"] + [f"@PL!{c}{B['rev']}|n2" for c in ALL],
              ["營收年增率"] + ["—"] + [f"@PL!{c}{B['gr']}|p1" for c in ALL[1:]],
              [f"@PL!A{B['cogs']}"] + [f"@PL!{c}{B['cogs']}|n2" for c in ALL],
              [f"@PL!A{B['gp']}"] + [f"@PL!{c}{B['gp']}|n2" for c in ALL],
              [f"@PL!A{B['gm']}"] + [f"@PL!{c}{B['gm']}|p1" for c in ALL],
              [f"@PL!A{B['opx']}"] + [f"@PL!{c}{B['opx']}|n2" for c in ALL],
              [f"@PL!A{B['op']}"] + [f"@PL!{c}{B['op']}|n2" for c in ALL],
              [f"@PL!A{B['opm']}"] + [f"@PL!{c}{B['opm']}|p1" for c in ALL],
              ["其他收入淨額（業外：財務淨額＋其他收益）"] + [f"@PL!{c}{B['nop']}|n2" for c in ALL],
              [f"@PL!A{B['pbt']}"] + [f"@PL!{c}{B['pbt']}|n2" for c in ALL],
              [f"@PL!A{B['tax']}"] + [f"@PL!{c}{B['tax']}|n2" for c in ALL],
              [f"@PL!A{B['nia']}"] + [f"@PL!{c}{B['nia']}|n2" for c in ALL],
              [f"@PL!A{B['nci']}"] + [f"@PL!{c}{B['nci']}|n2" for c in ALL],
              [f"@PL!A{B['ni']}"] + [f"@PL!{c}{B['ni']}|n2" for c in ALL],
              [f"@PL!A{B['nim']}"] + [f"@PL!{c}{B['nim']}|p1" for c in ALL],
              ["EPS（基本＝稀釋，人民幣元）"] + [f"@PL!{c}{B['eps']}|n2" for c in ALL],
              ["淨利年增率（歸屬母公司）"] + ["—"] + [f"@PL!{c}{B['nig']}|p1" for c in ALL[1:]]]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線；FY2026 起）", "years": f"@PL!F{B['hdr']}:L{B['hdr']}",
     "metrics": [{"name": "合併營收（十億元人民幣）", "fmt": "n2", "chart": True, "rows": {s: pl("rev", s, "F") for s in "ABC"}},
                 {"name": "營業利益（十億元人民幣）", "fmt": "n2", "chart": True, "rows": {s: pl("op", s, "F") for s in "ABC"}},
                 {"name": "營業利益率", "fmt": "p1", "rows": {s: pl("opm", s, "F") for s in "ABC"}},
                 {"name": "EPS（人民幣元）", "fmt": "n2", "chart": True, "rows": {s: pl("eps", s, "F") for s in "ABC"}},
                 {"name": "自由現金流扣股利（十億元人民幣；收支關卡用）", "fmt": "n2", "rows": {s: pl("fcfad", s, "F") for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：情境 B 現金流（十億元人民幣）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in ALL],
     "rowgen": {"sheet": "PL", "rows": f"{B['da']}-{B['fcfad']}", "cols": ["A"] + list(ALL), "fmts": [""] + ["n2"] * len(ALL)}}]}

def ptab(keys):
    return [[il(k), f"@In!D{inr(k)}|{f}", iv(k, f), f"@In!F{inr(k)}|{f}", f"@In!G{inr(k)}|tag"] for k, f in keys]
drivers = {"id": "drivers", "title": "驅動因子：全球人形出貨 → 優必選全尺寸人形 → 產品類別 → 合併營收", "blocks": [
    {"type": "theme", "title": "全尺寸人形營收（十億元人民幣，三情境）與每台營收（千元，情境 B）：FY2025–FY2032", "x": f"@PL!E{B['hdr']}:L{B['hdr']}", "fmt": "n1", "fmt_r": "n0", "stack": False, "pad": 1.25,
     "bars": [{"name": ["情境 ", s, "：全尺寸人形營收"], "y": pl("fs", s, "E"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": [f"@PL!A{B['asp']}", "（情境 B）"], "y": pl("asp", "B", "E"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "長條＝三情境全尺寸人形營收（銷量 × 每台營收）；折線＝情境 B 每台營收（右軸），依線模型全尺寸售價學習曲線下降。FY2025 為實績。"},
    {"type": "note", "parts": ["@PL!A1"]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照；進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": [[f"@In!B{inr(f + '_scn')}", f"@In!D{inr(f + '_scn')}|{fm}", f"@In!E{inr(f + '_scn')}|{fm}", f"@In!F{inr(f + '_scn')}|{fm}"]
              for f, fm in (("u26h2", "n0"), ("shr27", "p1"), ("dshr", "p2"), ("gm_fslt", "p0"), ("ne27", "n2"), ("gne", "p0"), ("gor", "p0"), ("ghw", "p0"), ("gop", "p0"))]},
    {"type": "table", "caption": "分部營收推導（情境 B；十億元人民幣）：全球出貨 × 占比 × 每台營收 → 全尺寸人形 → 六個產品類別 → 合併營收；核對列應為零",
     "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in ALL],
     "rowgen": {"sheet": "PL", "rows": f"{B['u']}-{B['gr']}", "cols": ["A"] + list(ALL), "fmts": [""] + ["n2"] * len(ALL),
                "row_fmt": {str(B["u"]): {c: "n0" for c in ALL}, str(B["glob"]): {c: "n0" for c in ALL}, str(B["shr"]): {c: "p1" for c in ALL},
                            str(B["pidx"]): {c: "n3" for c in ALL}, str(B["asp"]): {c: "n0" for c in ALL}, str(B["gr"]): {c: "p1" for c in ALL}}}},
    {"type": "table", "caption": "毛利與營業費用推導（情境 B；十億元人民幣）", "head": [f"@PL!A{B['hdr']}"] + [f"@PL!{c}{B['hdr']}" for c in ALL],
     "rowgen": {"sheet": "PL", "rows": f"{B['gmfs']}-{B['opm']}", "cols": ["A"] + list(ALL), "fmts": [""] + ["n2"] * len(ALL),
                "row_fmt": {str(B["gmfs"]): {c: "p1" for c in ALL}, str(B["gm"]): {c: "p1" for c in ALL}, str(B["opm"]): {c: "p1" for c in ALL}}}},
    {"type": "h3", "text": "一、全尺寸人形（命題模組）"},
    {"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": ptab((("u26h2_B", "n0"), ("asp_adj", "p0"), ("shr27_B", "p1"), ("dshr_B", "p2"), ("gm_fs26", "p0"), ("gm_fslt_B", "p0")))},
    {"type": "table", "collapse": True, "summary": "展開：實績錨點（固定輸入）", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag"] for k, f in (("u_h126", "n0"), ("asp_h1", "n1"), ("bl_25", "n2"))]},
    {"type": "h3", "text": "二、其他產品類別（非具身人形、其他智能機器人、硬件、鋒龍零部件）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": ptab((("ne26", "n3"), ("ne27_B", "n2"), ("gne_B", "p0"), ("or26", "n2"), ("gor_B", "p0"), ("hw26", "n2"), ("ghw_B", "p0"), ("gfl", "p0"), ("gm_leg", "p0"), ("gm_fl", "p1")))},
    {"type": "table", "collapse": True, "summary": "展開：鋒龍並表與其他固定輸入", "head": ["@In!B3", "@In!E3", "@In!G3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag"] for k, f in (("fl26", "n3"), ("fl27", "n3"), ("ot26", "n3"))]},
    {"type": "h3", "text": "三、全公司（營業費用、損益表下半段、現金流、折現）"},
    {"type": "table", "caption": "有敏感度區間的參數", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rows": ptab((("op26", "n2"), ("gop_B", "p0"), ("imp_r", "p1"), ("nonop", "n2"), ("tax_w", "p0"), ("nci_r", "p0"), ("fc_da", "p1"), ("cx26", "n2"), ("fc_cx", "p1"),
                   ("fc_wc", "p0"), ("beta", "n2"), ("erp", "p1"), ("g_cap", "p1"), ("ronic_sp", "p1"), ("wacc_adj", "p1"), ("w_dcf", "p0")))}]}

rc_last = find(RC, "加權目標價 vs 賣方平均目標價（港元）")
LS = find(RE, "本模型情境對應（Derived；PL 引用本區）"); LC = find(RE, "一致性對照（Derived）")
recon = {"id": "recon", "title": "對帳：公司說法、共識與線模型", "blocks": [
    {"type": "table", "caption": "@Recon!A1", "head": [f"@Recon!{c}2" for c in "ABCDE"],
     "rowgen": {"sheet": "Recon", "rows": f"3-{rc_last}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2", ""],
                "row_fmt": {str(find(RC, "FY2026 全尺寸人形銷量", exact=False)): {c: "n0" for c in "BCD"}, str(find(RC, "FY2027 全尺寸人形銷量", exact=False)): {c: "n0" for c in "BCD"},
                            str(find(RC, "2026H1 每台營收", exact=False)): {c: "n0" for c in "BCD"}}}},
    {"type": "table", "caption": "@對帳!A1", "head": [f"@對帳!{c}2" for c in "BCDEF"],
     "rowgen": {"sheet": "對帳", "rows": f"3-{LS - 2}", "cols": list("BCDEF"), "fmts": ["", "n2", "n2", "n2", ""]}},
    {"type": "table", "caption": "@對帳!A" + str(LC), "head": [f"@對帳!{c}{LC + 1}" for c in "ABCDE"],
     "rowgen": {"sheet": "對帳", "rows": f"{LC + 2}-{LC + 3}", "cols": list("ABCDE"), "fmts": ["", "n2", "n2", "n2", ""]}},
    {"type": "note", "parts": [f"@對帳!A{LC + 4}"]},
    {"type": "table", "collapse": True, "summary": "展開：線模型讀值對應到本模型情境（全球出貨、售價指數）", "head": [f"@對帳!{c}{LS + 1}" for c in "ABCDEF"],
     "rowgen": {"sheet": "對帳", "rows": f"{LS + 2}-{LC - 2}", "cols": list("ABCDEF"), "fmts": ["", "", "n3", "n3", "n3", ""]}}]}

RK = wb["Risk"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["資料日期", "@README!A2"], ["build_id", "@Sens!D2"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "table", "caption": "股價、匯率與折現率的日期", "head": ["@In!B3", "@In!E3", "@In!G3", "@In!H3"],
     "rows": [[il(k), iv(k, f), f"@In!G{inr(k)}|tag", nourl("In", f"H{inr(k)}")] for k, f in (("px", "n2"), ("fx_hkd", "n3"), ("fx_cny", "n3"), ("rf", "p2"))]},
    {"type": "table", "caption": "@Risk!A1", "head": ["@Risk!A2", "@Risk!B2", "@Risk!C2", "@Risk!D2"], "rows": [[nourl("Risk", f"{c}{r}") for c in "ABCD"] for r in range(3, RK.max_row + 1)]},
    {"type": "note", "parts": ["來源清單（含網址）見 xlsx Sources 頁與 In 頁來源欄；本頁為自含網頁，不放外部網址。"]},
    {"type": "h3", "text": "「未取得」與「未找到」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"In": ["B", "I"], "Hist": ["A"], "對帳": ["A"]}, "empty": "沒有標為未取得的項目。"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"In": ["H", "I"]}, "empty": "沒有標為未找到的項目。"}]}

cfg = {"company": "優必選 UBTech", "ticker": "9880.HK", "layout": "v1.3", "wide_text": True,
       "thesis": "@README!A3",
       "question": "全尺寸人形機器人從展示與專案訂單走到規模出貨的速度，能否支持現價；在已交付量與已公告訂單的證據邊界內，現價隱含多少出貨占比",
       "scenarios": [{"key": "A", "name": "情境 A（樂觀：出貨接近公司產量目標、毛利率維持高檔）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守：維持上半年交付速度、價格戰）"}],
       "hero": hero,
       "sections": [thesis, valuation, sens, revdcf, pl_sec, drivers, recon, status]}
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", X)

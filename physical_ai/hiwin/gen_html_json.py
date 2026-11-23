# 由上銀成品 xlsx 產生 physical_ai/hiwin/html.json（版面規格 v1.3；列號依標籤找，勿手改 html.json）。
# 用法：python3 physical_ai/hiwin/gen_html_json.py physical_ai/hiwin/dist/<成品>.xlsx physical_ai/hiwin/html.json
# 台股模型對外不公開目標價（README「網站不公開本檔目標價」）：本頁不引用任何目標價儲存格
#（Val 每股價值、DCF／Comparable／情境／加權目標價、升降級價格、Sens 加權目標價與 DCF 目標價欄、PRS敏感度 第一表、
#  Recon 目標價列的模型值與差、README 含目標價的列），只列投資建議、上檔空間與相對基準百分比。建置時以 FORBID 檢查。
import json, re, sys
import openpyxl

X, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(X, data_only=True)
V, PL, S, RC, IN, TH, CP, H, RD = (wb[n] for n in ("Val", "PL", "Sens", "Recon", "In", "Thesis", "Comps", "Hist", "README"))


def find(ws, label, start=1, col=1, exact=True):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and (v == label if exact else v.startswith(label)):
            return r
    raise KeyError(f"{ws.title} 找不到「{label}」")


vr = lambda lab: find(V, lab, exact=False)
VX = {k: vr(l) for k, l in dict(px="現價（元）", pxd="股價基準日", ke="股權成本 ke", kdat="稅後債務成本", we="權益占比", wacc="WACC", g="永續成長率 g", ronic="RONIC",
                                  pv="明確期 FCFF 現值", pvtv="終值現值", ev="企業價值 EV", nc="＋淨現金", stub="＋Stub", invv="＋非營業投資", mi="− 非控制權益",
                                  tvsh="終值現值占 EV", tvm="隱含終值 EV／EBITDA", dchk="檢查：EV", w="情境權重",
                                  niLG="線性滑軌 FY2028 分部淨利", niBS="滾珠螺桿 FY2028 分部淨利", niRB="機器人 FY2028 分部淨利", niOT="其他 FY2028 分部淨利",
                                  relnow="上銀現行前瞻 P/E", up="上檔空間（加權目標價", rate="估值評等", gate="收支關卡（FY2026", gateB="入不敷出判定用", gatej="入不敷出判定",
                                  rec="投資建議", dist="距最近級距邊界", a_dcf="上檔空間歸因：DCF", a_cu="上檔空間歸因：未調整", a_rel="上檔空間歸因：相對倍數", dcfsh="DCF 在加權目標價的占比",
                                  humsh="人形命題子列在 FY2028").items()}
VX["gatej"] = find(V, "入不敷出判定")   # 精確比對（前綴會先找到「入不敷出判定用」）
# 情境區塊起點（PL、Thesis）
PB0, PA0, PC0 = (find(PL, f"情境 {s}", exact=False) for s in "BAC")
TB0, TA0, TC0 = (find(TH, f"情境 {s}", exact=False) for s in "BAC")
pl = lambda lab, s0: find(PL, lab, s0)
P = {s: {k: pl(l, s0) for k, l in dict(hdr="項目", LG="線性滑軌 營收", BS="滾珠螺桿 營收", RB="機器人 營收", OT="其他 營收", rev="合併營收", yoy="合併營收年增",
                                        hLG="線性滑軌：人形（微型滑軌與小螺桿，命題）", hBS="滾珠螺桿：人形（行星滾柱絲杠，命題）", cLG="線性滑軌：核心（工業循環）", cBS="滾珠螺桿：核心（工業循環）",
                                        cRB="機器人 核心", cOT="其他 核心", gLG="線性滑軌 核心年增", gBS="滾珠螺桿 核心年增", gRB="機器人 核心年增", gOT="其他 核心年增",
                                        hsh="人形（命題）營收 ÷ 合併營收", mLG="線性滑軌 分部利益率", mBS="滾珠螺桿 分部利益率", mRB="機器人 分部利益率", mOT="其他 分部利益率",
                                        op="營業利益（分部利益合計）", opm="營業利益率", nonop="業外淨額", pt="稅前淨利", tax="所得稅", ni="本期淨利（合併）", nci="非控制權益損益",
                                        nip="歸屬母公司淨利", eps="EPS（元）", da="折舊攤銷", cx="資本支出（含預付設備款）", fcff="FCFF（＝NOPAT＋折舊 − 資本支出 − 營運資金增加）",
                                        div="現金股利發放", fcfe="股東自由現金流扣股利（＝歸屬母公司淨利＋折舊 − 資本支出 − 營運資金增加 − 股利）",
                                        chk1="檢查：分部合計 − 合併營收", chk2="檢查：子列合計 − 分部（線性滑軌＋滾珠螺桿）").items()} for s, s0 in (("B", PB0), ("A", PA0), ("C", PC0))}
T = {s: {k: find(TH, l, s0) for k, l in dict(hdr="項目", ship="全球人形機器人出貨", bom="每台硬體 BOM", prsv="每台行星滾柱絲杠價值（＝BOM × prs_sh）", ramp="上銀份額爬坡（相對穩態）",
                                            prsu="上銀 PRS 營收（百萬美元）", gdu="上銀 微型滑軌與小螺桿 營收（百萬美元）", prsn="上銀 PRS 營收（新台幣百萬元）→ 滾珠螺桿 人形子列",
                                            gdn="上銀 微型滑軌與小螺桿 營收（新台幣百萬元）→ 線性滑軌 人形子列", mkt="全球行星滾柱絲杠市場（人形；百萬美元）").items()} for s, s0 in (("B", TB0), ("A", TA0), ("C", TC0))}
PB = P["B"]
YR = "C:J"   # FY2025A–FY2032E
rng = lambda ws, r, a="C", b="J": f"@{ws}!{a}{r}:{b}{r}"
RCx = {k: find(RC, l, exact=False) for k, l in dict(g27="FY2027 合併營收年增", q4="4Q26 營收相對 3Q26", rbsh="FY2026 機器人營收占比", cons="分析師平均目標價", fx="美系外資目標價",
                                                     prs="隱含每支 PRS 單價", hs="FY2030 人形命題營收", mkt="FY2030 全球人形 PRS 市場", rvh="主導變數").items()}
inr = {IN.cell(r, 1).value: r for r in range(4, IN.max_row + 1) if IN.cell(r, 1).value}
SE0 = find(S, "基準（本檔）")
case_rows = [r for r in range(SE0 + 1, S.max_row + 1) if isinstance(S.cell(r, 5).value, (int, float))]
SCN = [{"key": "A", "name": "情境 A（樂觀）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守）"}]
cmp_rows = [r for r in range(5, 15)]
hw_row = find(CP, "上銀（參照）")
lin_med, rob_med = find(CP, "線性傳動中位數", exact=False), find(CP, "機器人中位數", exact=False)
G2 = find(wb["PRS敏感度"], "第二表", exact=False)
g2h = G2 + 9   # 第二表表頭（prs_sh ＼ hw_prs）
assert wb["PRS敏感度"].cell(g2h, 1).value == "prs_sh ＼ hw_prs", wb["PRS敏感度"].cell(g2h, 1).value


def inrows(keys):
    return [[f"@In!B{inr[k]}", f"@In!D{inr[k]}|{fm(k)}", f"@In!E{inr[k]}|{fm(k)}", f"@In!F{inr[k]}|{fm(k)}", f"@In!G{inr[k]}|tag"] for k in keys]


def fm(k):
    u = IN.cell(inr[k], 3).value
    return "p1" if u in ("%", "pt") else ("n0" if u in ("NT$M", "台", "US$") else "n2")


def in_block(title, keys):
    rng_k = [k for k in keys if IN.cell(inr[k], 4).value is not None or IN.cell(inr[k], 6).value is not None]
    fix = [k for k in keys if k not in rng_k]
    out = [{"type": "h3", "text": title}]
    if rng_k:
        out.append({"type": "table", "head": ["參數", "敏感度低", "基準", "敏感度高", "標記"], "rows": inrows(rng_k)})
    if fix:
        out.append({"type": "table", "collapse": True, "summary": f"展開：{title}的固定輸入", "head": ["參數", "敏感度低", "基準", "敏感度高", "標記"], "rows": inrows(fix)})
    return out


GROUPS = [("一、工業循環：分部成長與利益率", ["q4g", "capex_26", "g27_LG", "g28_LG", "g27_BS", "g28_BS", "g27_RB", "g28_RB", "g27_OT", "g28_OT", "g_lt", "g_lt_RB", "g_lt_OT",
                                           "m_LG", "m_BS", "m_RB", "m_OT", "mg_d", "d27_A", "d27_B", "d27_C", "d28_A", "d28_B", "d28_C", "mg_A", "mg_B", "mg_C"]),
          ("二、人形機器人線性關節（命題模組）", ["hu_26", "bom_27", "bom_32", "prs_sh", "n_scr", "ex_cn", "hw_prs", "gd_sh", "hw_gd", "ramp_27", "ramp_28", "ramp_29", "fx",
                                         "hu30_A", "hu30_B", "hu30_C", "hu32_A", "hu32_B", "hu32_C"]),
          ("三、全公司：損益、折現率與估值規則", ["px", "px_date", "tp_date", "val_yrs", "hold_yrs", "sh_v", "nonop_r", "tax_r", "nci_r", "payout", "rf", "rf_us", "blume_a", "beta_1", "beta_2", "beta",
                                          "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "cash_v", "debt_v", "lease_v", "inv_v", "inv_k", "mi_v",
                                          "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc", "dps_tgt", "pe_lin", "pe_rob", "rel_k", "dcf_w", "w_A", "w_B", "w_C"])]
allk = [k for _, ks in GROUPS for k in ks]
missing = [k for k in inr if re.fullmatch(r"[a-z][a-z0-9_]*", str(k)) and k not in allk and inr[k] < find(IN, "情境驅動對照", exact=False, start=1)]
assert not missing and len(allk) == len(set(allk)), (missing, [k for k in allk if allk.count(k) > 1])

v = lambda k, f="", col="C": f"@Val!{col}{VX[k]}" + (f"|{f}" if f else "")
cfg = {
    "company": "上銀 HIWIN", "ticker": "2049.TW", "layout": "v1.3", "wide_text": True,
    "thesis": "@README!A7",
    "question": "工業循環的利潤率能否維持，以及人形機器人線性關節（行星滾柱絲杠、微型滑軌）能否成為第二成長曲線；兩者合起來是否支持現價",
    "scenarios": SCN,
    "hero": {"rec": v("rec"),
             "tiles": [["上檔空間（加權，持有期間）", v("up", "sp1"), ["估值評等 ", v("rate")]],
                       [["現價（", v("pxd"), " 收盤）"], v("px", "n2"), "Yahoo Finance 日線"],
                       ["收支關卡（情境 B，FY2026–FY2030）", v("gateB", "n0"), ["百萬元，", v("gatej")]]],
             "note": ["上檔空間 ", v("up", "sp1"), "，落在中立區間，距最近評等級距邊界 ", v("dist", "n1"), " 個百分點；估值主要由工業循環的分部利益率決定，人形子列在 FY2028 分部淨利只占 ",
                      v("humsh", "p1"), "。台股模型網站不公開目標價，本頁只列投資建議與上檔空間。"]},
    "sections": [],
}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "人形子列營收（新台幣百萬元，情境 B 堆疊）與人形占合併營收（右軸，三情境）", "x": rng("PL", PB["hdr"], "D"), "fmt": "n0", "fmt_r": "p1", "stack": True, "pad": 1.2,
     "bars": [{"name": [f"@PL!A{PB['hBS']}"], "y": rng("PL", PB["hBS"], "D"), "color": "#0b7a80"},
              {"name": [f"@PL!A{PB['hLG']}"], "y": rng("PL", PB["hLG"], "D"), "color": "#c29a2b"}],
     "lines": [{"name": ["情境 A：", f"@PL!A{P['A']['hsh']}"], "y": rng("PL", P["A"]["hsh"], "D"), "axis": "r", "color": "#6f6bae", "dash": True},
               {"name": ["情境 B：", f"@PL!A{PB['hsh']}"], "y": rng("PL", PB["hsh"], "D"), "axis": "r", "color": "#3b3838", "labels": True},
               {"name": ["情境 C：", f"@PL!A{P['C']['hsh']}"], "y": rng("PL", P["C"]["hsh"], "D"), "axis": "r", "color": "#b4472f", "dash": True}],
     "caption": "長條＝命題模組算出的人形子列營收（掛在滾珠螺桿與線性滑軌分部之下）；折線＝人形子列占合併營收（右軸）。這是模型輸出，非獨立證據。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": ["@README!A8", "@README!A10"]},
    {"type": "list", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "items": [
        ["工業循環上行能走多久、利潤率能否維持？答：基準情境以最近三季分部利益率持平，情境 A 回到最近一季水準、情境 C 回到 FY2025 水準；",
         f"@Sens!A{find(S, '分部利益率 ±2pt（低）', SE0)}", " 時加權目標價相對基準 ", f"@Sens!E{find(S, '分部利益率 ±2pt（低）', SE0)}|sp1", "，支持現價所需的分部利益率加減為 ",
         f"@Recon!C{RCx['rvh'] + 1}|sp1", "。"],
        ["人形線性關節能否成為第二成長曲線？答：情境 B 的人形子列 FY2030 占合併營收 ", f"@PL!H{PB['hsh']}|p1", "（情境 A ", f"@PL!H{P['A']['hsh']}|p1", "、情境 C ",
         f"@PL!H{P['C']['hsh']}|p1", "）；人形命題歸零時加權目標價相對基準 ", f"@Sens!E{find(S, '人形命題歸零', SE0)}|sp1", "，影響有限。"],
        ["兩者合起來是否支持現價？答：投資建議「", v("rec"), "」，上檔空間 ", v("up", "sp1"), "；", "@README!A6"]]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值落在中立：上檔空間 ", v("up", "sp1"), "，距最近級距邊界 ", v("dist", "n1"), " 個百分點；DCF 貢獻 ", v("a_dcf", "sp1"), "、Comparable 貢獻 ", v("a_cu", "sp1"), "。"],
        ["利潤率是最大槓桿：", f"@Sens!A{find(S, '分部利益率 ±2pt（高）', SE0)}", " 時加權目標價相對基準 ", f"@Sens!E{find(S, '分部利益率 ±2pt（高）', SE0)}|sp1",
         "；情境 A 權重全給時 ", f"@Sens!E{find(S, '情境 A 權重 100%', SE0)}|sp1", "、情境 C 全給時 ", f"@Sens!E{find(S, '情境 C 權重 100%', SE0)}|sp1", "。"],
        ["人形命題目前占比小：情境 B 人形子列在 FY2028 分部淨利占 ", v("humsh", "p1"), "；行星滾柱絲杠占 BOM 是本模型最大不確定（兩來源互相矛盾），但對估值影響有限。"],
        ["同業倍數：線性傳動前瞻本益比中位數 ", f"@Comps!F{lin_med}|n1", " 倍、機器人 ", f"@Comps!F{rob_med}|n1", " 倍；上銀現行前瞻本益比為線性傳動中位數的 ", v("relnow", "n2"),
         " 倍（缺十年比值，未作相對倍數調整）。"],
        ["收支關卡通過：情境 B FY2026–FY2030 股東自由現金流扣股利合計 ", v("gateB", "n0"), " 百萬元（情境 A ", v("gate", "n0", "C"), "、情境 C ", v("gate", "n0", "E"), "）。"]]},
    {"type": "h3", "text": "風險登錄"},
    {"type": "note", "parts": ["@README!A14"]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "compare", "title": "市場目標價與現價（元；賣方目標價為 Interested-party，灰色）與本模型上檔空間（台股模型網站不公開本檔目標價）",
     "groups": [{"name": "市場目標價與現價（元）", "fmt": "n2", "items": [{"name": "現價", "v": v("px")},
                                                                  {"name": [f"@Recon!A{RCx['cons']}"], "v": f"@Recon!C{RCx['cons']}", "ref": True},
                                                                  {"name": [f"@Recon!A{RCx['fx']}"], "v": f"@Recon!C{RCx['fx']}", "ref": True}]},
                {"name": "本模型：上檔空間與歸因", "fmt": "sp1", "items": [{"name": "加權上檔空間", "v": v("up")}, {"name": "DCF 貢獻", "v": v("a_dcf")},
                                                                    {"name": "Comparable 貢獻", "v": v("a_cu")}]}],
     "caption": "左組：現價與賣方目標價（分析師平均為財報前資料，外資為財報後）；右組：本模型加權上檔空間，及 DCF與 Comparable 兩段貢獻（相對倍數調整未作，貢獻為零；合計等於上檔空間）。本模型目標價依台股規則不公開。"},
    {"type": "note", "parts": ["DCF 與分部加總 Comparable 各占一半，三情境權重 ", v("w", "p0", "C"), "／", v("w", "p0", "D"), "／", v("w", "p0", "E"),
                               "；投資建議依加權上檔空間的五級規則判定，並經收支關卡檢查。市場共識目標價取自資料商統計頁，只作對照。"]},
    {"type": "table", "head": ["項目", "值"], "rows": [
        ["@Val!A" + str(VX["up"]), v("up", "sp1")], ["@Val!A" + str(VX["rate"]), v("rate")], ["@Val!A" + str(VX["rec"]), v("rec")],
        ["@Val!A" + str(VX["dist"]), v("dist", "n1")], ["@Val!A" + str(VX["gatej"]), v("gatej")],
        [f"@Recon!A{RCx['cons']}", f"@Recon!C{RCx['cons']}|n2"], [f"@Recon!A{RCx['fx']}", f"@Recon!C{RCx['fx']}|n2"]]},
    {"type": "h3", "text": "折現率"},
    {"type": "table", "head": ["項目", "值"], "rows": [["@Val!A" + str(VX[k]), v(k, "p2")] for k in ("wacc", "ke", "kdat", "ronic", "g")] + [["@Val!A" + str(VX["we"]), v("we", "p1")]]},
    {"type": "h3", "text": "DCF 彙總（新台幣百萬元）"},
    {"type": "table", "grid": True, "head": ["項目", "情境 A", "情境 B", "情境 C"], "rows": [
        [f"@Val!A{VX[k]}", f"@Val!C{VX[k]}:E{VX[k]}|{f}"] for k, f in (("pv", "n0"), ("pvtv", "n0"), ("ev", "n0"), ("tvsh", "p1"), ("tvm", "n1"))]},
    {"type": "table", "collapse": True, "summary": "展開：企業價值到股權的調整項與檢查", "grid": True, "head": ["項目", "情境 A", "情境 B", "情境 C"], "rows": [
        [f"@Val!A{VX[k]}", f"@Val!C{VX[k]}:E{VX[k]}|{f}"] for k, f in (("nc", "n0"), ("stub", "n0"), ("invv", "n0"), ("mi", "n0"), ("dchk", ""), ("w", "p0"))]},
    {"type": "h3", "text": "分部加總 Comparable（FY2028 分部淨利 × 同業前瞻本益比；新台幣百萬元）"},
    {"type": "table", "grid": True, "head": ["項目", "情境 A", "情境 B", "情境 C"], "rows": [
        [f"@Val!A{VX[k]}", f"@Val!C{VX[k]}:E{VX[k]}|n0"] for k in ("niLG", "niBS", "niRB", "niOT")] + [
        [f"@In!B{inr['pe_lin']}", f"@In!E{inr['pe_lin']}|n2", "", ""], [f"@In!B{inr['pe_rob']}", f"@In!E{inr['pe_rob']}|n2", "", ""],
        [f"@Val!A{VX['relnow']}", v("relnow", "n2"), "", ""]]},
    {"type": "table", "collapse": True, "summary": "展開：同業倍數（Comps 頁；部分前瞻本益比依最新收盤等比例換算）", "head": [f"@Comps!{c}4" for c in "ABCDEFGJKP"],
     "rows": [[f"@Comps!A{r}", f"@Comps!B{r}", f"@Comps!C{r}", f"@Comps!D{r}", f"@Comps!E{r}", f"@Comps!F{r}|n2", f"@Comps!G{r}|n2", f"@Comps!J{r}|n2", f"@Comps!K{r}|tag", f"@Comps!P{r}"]
              for r in cmp_rows + [hw_row]] + [[f"@Comps!A{lin_med}", "", "", "", "", f"@Comps!F{lin_med}|n2", "", "", "", ""], [f"@Comps!A{rob_med}", "", "", "", "", f"@Comps!F{rob_med}|n2", "", "", "", ""]]},
    {"type": "table", "collapse": True, "summary": "展開：上檔空間歸因與評等規則", "head": ["項目", "值"], "rows": [
        ["@Val!A" + str(VX[k]), v(k, "sp1")] for k in ("a_dcf", "a_cu", "a_rel")] + [["@Val!A" + str(VX["dcfsh"]), v("dcfsh", "p1")], [f"@Val!A{VX['rate']}", f"@Val!G{VX['rate']}"],
                                                                                         [f"@Val!A{VX['rec']}", f"@Val!G{VX['rec']}"]]},
]}

# ---- 敏感度（加權目標價相對基準 %；不引用 Sens 的目標價欄）----
lo = [r for r in case_rows if re.search(r"（(全部)?低）$", S.cell(r, 1).value)]
hi = [r for r in case_rows if re.search(r"（(全部)?高）$", S.cell(r, 1).value)]
oth = [r for r in case_rows if r not in lo + hi and abs(S.cell(r, 5).value) > 0]
TOR = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": "E", "fmt": "sp1"}
big = lambda rs: [r for r in rs if abs(S.cell(r, 5).value) >= 0.1]
small = lambda rs: [r for r in rs if 0 < abs(S.cell(r, 5).value) < 0.1]
j = lambda rs: ",".join(map(str, rs))
sens_blocks = [dict(TOR, title="加權目標價相對基準：輸入取低值，變動達兩位數百分比的案例", rows=j(big(lo))),
               dict(TOR, title="加權目標價相對基準：輸入取高值，變動達兩位數百分比的案例", rows=j(big(hi))),
               {"type": "note", "parts": ["@Sens!A1"]}]
if big(oth):
    sens_blocks.append(dict(TOR, title="加權目標價相對基準：情境權重與整組案例，變動達兩位數百分比者", rows=j(big(oth))))
for nm, rs in (("輸入取低值", lo), ("輸入取高值", hi), ("情境權重與整組案例", oth)):
    if small(rs):
        sens_blocks.append(dict(TOR, title=f"加權目標價相對基準：{nm}，其餘案例", rows=j(small(rs)), collapse=True, summary=f"展開：{nm}，變動未達兩位數百分比的案例"))
PS = wb["PRS敏感度"]
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：行星滾柱絲杠占 BOM × 上銀份額 → FY2030 人形占合併營收（情境 B）", "grid": True,
                    "caption": f"@PRS敏感度!A{G2}", "head": [f"@PRS敏感度!A{g2h}:F{g2h}"], "rows": [[f"@PRS敏感度!A{r}", f"@PRS敏感度!B{r}:F{r}|p1"] for r in range(g2h + 1, g2h + 6)]})
sens_blocks.append({"type": "table", "collapse": True, "summary": "展開：全部案例（相對基準、DCF 相對基準與估值評等）",
                    "head": ["驅動", "設定", "加權目標價相對基準", "DCF 相對基準", "估值評等"],
                    "rowgen": {"sheet": "Sens", "rows": j(case_rows), "cols": ["A", "B", "E", "G", "H"], "fmts": ["", "", "sp1", "sp1", ""]}})
sens_blocks.append({"type": "kv", "items": [["模型狀態", f"@Sens!D{find(S, '模型狀態')}"], ["build_id", f"@Sens!D{find(S, 'build_id')}"]]})
sens = {"id": "sens", "title": "敏感度", "blocks": sens_blocks}

rv0 = RCx["rvh"]
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼", "blocks": [
    {"type": "note", "parts": ["以整本重算反解：一次只動一個主導變數，求加權目標價等於現價（及升、降一級邊界）所需的值；其他假設不變。"]},
    {"type": "table", "head": [f"@Recon!A{rv0}", f"@Recon!B{rv0}", f"@Recon!C{rv0}", f"@Recon!D{rv0}", f"@Recon!E{rv0}"],
     "rows": [[f"@Recon!A{r}"] + [f"@Recon!{c}{r}|" + ("sp1" if r == rv0 + 1 else "p1") for c in "BCDE"] for r in range(rv0 + 1, rv0 + 4)]}]}

# ---- 損益（標準損益表；IFRS；模型未建營業成本、毛利、營業費用）----
HC = "C"   # Hist FY2024
hr = {k: find(H, l, col=2) for k, l in dict(rev="合併營收", op="營業利益（＝分部利益合計）", pt="稅前淨利", tax="所得稅費用", ni="本期淨利（合併）", nip="歸屬母公司淨利", eps="基本 EPS（元）").items()}
def plrow(lab, pk, f, hk=None):
    return [lab, (f"@Hist!{HC}{hr[hk]}|{f}" if hk else "—"), rng("PL", PB[pk]) + f"|{f}"]
pl_sec = {"id": "pl", "title": "損益", "intro": ["口徑：IFRS（台灣上市公司申報口徑，即 GAAP；公司與市場皆用此口徑）。FY2024A 取年報（Hist 頁），FY2025A 起取 PL 頁情境 B；FY2026E 為前三季實績加第四季模型。",
                                                 "模型以分部利益合計推營業利益，未建營業成本、毛利與營業費用列（FY2025 毛利只在 Hist 頁）。"],
          "blocks": [
              {"type": "pl", "years": rng("PL", PB["hdr"]), "metrics": [
                  {"name": "合併營收（新台幣百萬元）", "fmt": "n0", "chart": True, "rows": {s: rng("PL", P[s]["rev"]) for s in "ABC"}},
                  {"name": "營業利益率", "fmt": "p1", "chart": True, "ymin": "auto", "rows": {s: rng("PL", P[s]["opm"]) for s in "ABC"}},
                  {"name": "EPS（元）", "fmt": "n2", "chart": True, "rows": {s: rng("PL", P[s]["eps"]) for s in "ABC"}}]},
              {"type": "table", "grid": True, "caption": "標準損益表（情境 B；新台幣百萬元，EPS 元；IFRS）",
               "head": ["項目", f"@Hist!{HC}3", rng("PL", PB["hdr"])],
               "rows": [plrow("營收（合併）", "rev", "n0", "rev"), plrow(f"@PL!A{PB['yoy']}", "yoy", "p1"),
                        plrow(f"@PL!A{PB['op']}", "op", "n0", "op"), plrow(f"@PL!A{PB['opm']}", "opm", "p1"),
                        plrow(f"@PL!A{PB['nonop']}", "nonop", "n0"), plrow(f"@PL!A{PB['pt']}", "pt", "n0", "pt"), plrow(f"@PL!A{PB['tax']}", "tax", "n0", "tax"),
                        plrow(f"@PL!A{PB['ni']}", "ni", "n0", "ni"), plrow(f"@PL!A{PB['nci']}", "nci", "n0"), plrow(f"@PL!A{PB['nip']}", "nip", "n0", "nip"),
                        plrow(f"@PL!A{PB['eps']}", "eps", "n2", "eps")]},
              {"type": "table", "grid": True, "collapse": True, "summary": "展開：現金流與收支關卡（情境 B）", "head": ["項目", rng("PL", PB["hdr"])],
               "rows": [[f"@PL!A{PB[k]}", rng("PL", PB[k]) + "|n0"] for k in ("da", "cx", "fcff", "div", "fcfe")]}]}

# ---- 驅動因子：分部營收推導 ----
drivers = {"id": "drivers", "title": "驅動因子 → 推導量", "blocks": [
    {"type": "theme", "title": "分部營收（新台幣百萬元，情境 B 堆疊）與合併營收年增（右軸）", "x": rng("PL", PB["hdr"], "D"), "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.2,
     "total": rng("PL", PB["rev"], "D"),
     "bars": [{"name": [f"@PL!A{PB[k]}"], "y": rng("PL", PB[k], "D"), "color": c} for k, c in (("LG", "#0b7a80"), ("BS", "#10a0a7"), ("RB", "#c29a2b"), ("OT", "#8a8c8c"))],
     "lines": [{"name": [f"@PL!A{PB['yoy']}"], "y": rng("PL", PB["yoy"], "D"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "四個官方分部（2026 年起機器人單列）的營收由核心成長路徑推出，線性滑軌與滾珠螺桿另加人形子列；堆疊合計取 PL 頁合併營收列。"},
    {"type": "h3", "text": "分部營收推導（情境 B）"},
    {"type": "table", "grid": True, "head": ["項目", rng("PL", PB["hdr"])], "rows": [[f"@PL!A{PB[k]}", rng("PL", PB[k]) + ("|p1" if k[0] in "gm" or k == "yoy" else "|n0")]
                                                                               for k in ("gLG", "gBS", "gRB", "gOT", "cLG", "hLG", "LG", "cBS", "hBS", "BS", "RB", "OT", "rev", "chk1", "chk2", "yoy",
                                                                                         "mLG", "mBS", "mRB", "mOT", "op")]},
    {"type": "table", "grid": True, "collapse": True, "summary": "展開：情境驅動對照（三情境並排）", "head": [f"@In!B{find(IN, '情境驅動對照', exact=False) + 1}:F{find(IN, '情境驅動對照', exact=False) + 1}"],
     "rows": [[f"@In!B{r}", f"@In!C{r}", f"@In!D{r}:F{r}|" + ("p1" if IN.cell(r, 3).value in ("%", "pt") else "n0")] for r in range(find(IN, "情境驅動對照", exact=False) + 2, find(IN, "情境驅動對照", exact=False) + 8)]},
] + [b for t, ks in GROUPS for b in in_block(t, ks)]}

humanoid = {"id": "humanoid", "title": "命題模組：人形機器人線性關節", "blocks": [
    {"type": "theme", "title": "上銀行星滾柱絲杠營收（百萬美元，三情境並列）與全球人形出貨（右軸，台）", "x": rng("Thesis", T["B"]["hdr"], "E"), "fmt": "n0", "fmt_r": "n0", "stack": False,
     "bars": [{"name": [f"情境 {s}：", f"@Thesis!A{T[s]['prsu']}"], "y": rng("Thesis", T[s]["prsu"], "E"), "color": c} for s, c in (("A", "#6fc7cb"), ("B", "#0b7a80"), ("C", "#c29a2b"))],
     "lines": [{"name": ["情境 B：", f"@Thesis!A{T['B']['ship']}"], "y": rng("Thesis", T["B"]["ship"], "E"), "axis": "r", "color": "#3b3838"}],
     "caption": "上銀 PRS 營收＝全球出貨 × 每台 BOM × PRS 占 BOM × 非中國供應鏈比例 × 上銀份額 × 爬坡；三情境只差全球出貨。"},
    {"type": "note", "parts": ["@Thesis!A2"]},
    {"type": "table", "grid": True, "caption": "命題模組（情境 B）", "head": [f"@Thesis!A{T['B']['hdr']}:J{T['B']['hdr']}"],
     "rows": [[f"@Thesis!A{T['B'][k]}", f"@Thesis!B{T['B'][k]}", rng("Thesis", T["B"][k], "C") + ("|n2" if k == "ramp" else "|n1" if k in ("prsu", "gdu") else "|n0")] for k in ("ship", "bom", "prsv", "ramp", "prsu", "gdu", "prsn", "gdn", "mkt")]},
    {"type": "table", "grid": True, "collapse": True, "summary": "展開：情境 A、C 的出貨與上銀營收", "head": [f"@Thesis!A{T['A']['hdr']}:J{T['A']['hdr']}"],
     "rows": [{"sub": "情境 A（樂觀）"}] + [[f"@Thesis!A{T['A'][k]}", f"@Thesis!B{T['A'][k]}", rng("Thesis", T["A"][k], "C") + "|n0"] for k in ("ship", "prsn", "gdn", "mkt")] +
             [{"sub": "情境 C（保守）"}] + [[f"@Thesis!A{T['C'][k]}", f"@Thesis!B{T['C'][k]}", rng("Thesis", T["C"][k], "C") + "|n0"] for k in ("ship", "prsn", "gdn", "mkt")]}]}

DZ = wb["對帳"]
recon = {"id": "recon", "title": "對帳", "blocks": [
    {"type": "note", "parts": ["@Recon!A1"]},
    {"type": "table", "head": ["項目", "模型值", "對照值", "差／隱含", "標記", "說明"], "rows": [
        [f"@Recon!A{RCx[k]}", f"@Recon!B{RCx[k]}|{f}", f"@Recon!C{RCx[k]}|{f}", f"@Recon!D{RCx[k]}|{fd}", f"@Recon!E{RCx[k]}|tag", f"@Recon!G{RCx[k]}"]
        for k, f, fd in (("g27", "p1", "sp1"), ("q4", "p1", "sp1"), ("rbsh", "p1", "sp1"), ("prs", "n0", "sp1"))] + [
        [f"@Recon!A{RCx[k]}", f"@Recon!B{RCx[k]}|{f}", "", "", f"@Recon!E{RCx[k]}|tag", f"@Recon!F{RCx[k]}"] for k, f in (("hs", "p1"), ("mkt", "n0"))] + [
        [f"@Recon!A{RCx[k]}", "不公開（台股）", f"@Recon!C{RCx[k]}|n2", "", f"@Recon!E{RCx[k]}|tag", f"@Recon!G{RCx[k]}"] for k in ("cons", "fx")]},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（對帳頁）", "head": [f"@對帳!A3:E3", "@對帳!G3:H3"], "rows": [[f"@對帳!A{r}:E{r}", f"@對帳!G{r}:H{r}"] for r in range(4, DZ.max_row + 1) if DZ.cell(r, 2).value]},
    {"type": "note", "parts": [f"@對帳!A{max(r for r in range(1, DZ.max_row + 1) if DZ.cell(r, 1).value)}"]}]}

SO = wb["Sources"]
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["版本", "$version"], ["成品", "$xlsx"], ["雜湊指紋", "$sha"], ["模型標題", "@README!A1"], ["股價基準日", v("pxd")]]},
    {"type": "list", "items": ["@README!A12", "@README!A13", f"@README!A{RD.max_row}"]},
    {"type": "scan", "needle": "未取得", "sheets": {"Hist": ["B"], "Comps": ["A"]}, "collapse": True, "summary": "展開：標「未取得」的項目"},
    {"type": "note", "parts": ["來源清單（含網址）與每個輸入的出處、來源誘因見 xlsx 的 Sources 頁與 In 頁 H、I 欄；本頁為自含頁面，不放外部網址。In 頁標「未取得」的項目見 xlsx。"]},
    {"type": "note", "parts": ["@Comps!A2"]}]}

cfg["sections"] = [thesis, valuation, sens, revdcf, pl_sec, drivers, humanoid, recon, status]

# ---- 不得引用目標價儲存格 ----
FORBID = set()
for k in ("每股價值（評價日）", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）", "Comparable（未調整，每股）", "Comparable（相對倍數調整後", "Comparable 目標價", "情境目標價",
          "加權目標價（元）", "情境區間下緣", "情境區間上緣", "目標價達此值即升一級", "目標價低於此值即降一級", "分部加總股權價值", "股權價值"):
    r = vr(k)
    FORBID |= {f"Val!{c}{r}" for c in "CDEF"}
FORBID |= {f"Sens!{c}{r}" for c in "DF" for r in range(SE0, S.max_row + 1)} | {"Sens!D6"}
FORBID |= {f"PRS敏感度!{c}{r}" for c in "ABCDEF" for r in range(1, G2)} | {f"PRS敏感度!A{PS.max_row}"}
FORBID |= {f"Recon!{c}{RCx[k]}" for c in "BD" for k in ("cons", "fx")} | {f"Recon!F{r}" for r in range(rv0 + 1, rv0 + 4)}
FORBID |= {f"README!A{r}" for r in range(1, RD.max_row + 1) if re.search(r"目標價 \d|DCF \d|Comparable \d|達 \d|低於 \d|元、\d", str(RD.cell(r, 1).value))}
txt = json.dumps(cfg, ensure_ascii=False)
refs = set()
for m in re.finditer(r'"@([^"|]+)', txt):
    for part in m.group(1).split(","):
        sh, _, co = part.partition("!")
        if ":" in co:
            a, b = co.split(":")
            ca, ra = re.match(r"([A-Z]+)(\d+)", a).groups(); cb, rb = re.match(r"([A-Z]+)(\d+)", b).groups()
            from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl
            for rr in range(int(ra), int(rb) + 1):
                for cc in range(ci(ca), ci(cb) + 1):
                    refs.add(f"{sh}!{gl(cc)}{rr}")
        else:
            refs.add(f"{sh}!{co}")
bad = sorted(refs & FORBID)
assert not bad, f"html.json 引用了目標價儲存格：{bad}"
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(f"{OUT}：{len(cfg['sections'])} 章；引用 {len(refs)} 格；目標價儲存格 0 格")

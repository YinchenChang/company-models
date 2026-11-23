# -*- coding: utf-8 -*-
"""由 Tesla 成品 xlsx 產生 physical_ai/tesla/html.json（版面規格 v1.3）。
列號一律依 xlsx 欄 A 標籤尋找（不寫死），改版移列後重跑即可；勿手改 html.json。
HTML 只引用本模型自己的頁（README、In、Hist、Optimus、PL、DCF、SOTP、Val、Sens、Recon、對帳、風險登錄、改版紀錄），
不引用原檔沿用頁（Non-GAAP、長期評價分析等）。
用法：python3 physical_ai/tesla/gen_html_json.py [<成品 xlsx>] [<輸出 html.json>]
"""
import glob
import json
import os
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
X = sys.argv[1] if len(sys.argv) > 1 else sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))[-1]
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "html.json")
wb = openpyxl.load_workbook(X, data_only=True)
CH_SHEET = next(s for s in wb.sheetnames if s.startswith("改版紀錄_"))


def lab(ws, r, c=1):
    v = ws.cell(r, c).value
    return str(v).strip() if v is not None else None


def find(sh, text, start=1, col=1, prefix=False):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        v = lab(ws, r, col)
        if v is None:
            continue
        if (v.startswith(text) if prefix else v == text):
            return r
    raise SystemExit(f"{sh} 找不到標籤「{text}」（自第 {start} 列）")


# ---------------------------------------------------------------- 列號
IN = {lab(wb["In"], r): r for r in range(1, wb["In"].max_row + 1) if lab(wb["In"], r)}
Z0 = find("In", "Z. 情境驅動對照", prefix=True)
INZ = {lab(wb["In"], r): r for r in range(Z0, wb["In"].max_row + 1) if lab(wb["In"], r)}

PLK = [("seg_auto", "汽車分部營收（官方分部）"), ("veh", "車輛銷售（非 FSD／Robotaxi）"), ("fsd", "FSD"), ("rt", "Robotaxi"), ("crd", "碳權"),
       ("lease", "汽車租賃"), ("svc", "服務及其他"), ("seg_en", "能源分部營收（官方分部）"), ("seg_opt", "Optimus（尚未列報；模型新增）"),
       ("rev", "合併營收"), ("chk", "檢查：分部合計 − 合併營收"), ("gp", "毛利"), ("gm", "毛利率"), ("opex", "營業費用"), ("ebit", "營業利益（EBIT）"),
       ("nonop", "業外"), ("pti", "稅前淨利"), ("tax", "所得稅"), ("mi", "少數股權"), ("ni", "淨利（GAAP）"), ("sh", "稀釋股數（百萬）"),
       ("eps", "EPS（稀釋）"), ("da", "折舊"), ("sbc", "股票薪酬（非現金）"), ("cx", "資本支出"), ("wc", "營運資金增加"), ("fcf", "自由現金流（扣股利；Tesla 不配息）")]
PR = {}
for s in "ABC":
    r0 = find("PL", f"情境 {s}")
    PR[s] = {k: find("PL", t, r0) for k, t in PLK}
PLH = find("PL", "項目")
PLC = "CDEFGHIJKLM"      # FY2025–FY2035
FC = "EFGHIJKLM"         # FY2027–FY2035

OPK = [("u", "對外交付量"), ("cum", "累積產量"), ("bom", "每台 BOM（學習曲線）"), ("cost", "每台製造成本"), ("rev", "Optimus 營收"), ("gp", "Optimus 毛利"),
       ("gm", "毛利率"), ("sga", "Optimus 銷售與服務費用"), ("op", "Optimus 營業利益（不含研發）"), ("cx", "Optimus 產能資本支出")]
OR = {}
for s in "ABC":
    r0 = find("Optimus", f"情境 {s}")
    OR[s] = {k: find("Optimus", t, r0) for k, t in OPK}
O_ASP = find("Optimus", "平均售價")
O_TCO = find("Optimus", "TCO 與回本期", prefix=True)

DCFK = [("sumpv", "明確期 FCFF 現值合計（含 stub）"), ("tv", "終值（FY2035 年底）"), ("tvpv", "終值現值"), ("ev", "企業價值 EV"),
        ("tvsh", "終值占 EV", True), ("eq", "股權價值（EV＋淨現金與投資）"), ("eqT", "股權價值（兆美元", True), ("vfr", "CEO 獎酬歸屬比例"),
        ("vsh", "CEO 獎酬淨增股數", True), ("shv", "評價股數", True), ("vps", "每股價值（評價日）"), ("v12", "DCF 情境價值推到 12 個月後", True)]
DR = {}
for s in "ABC":
    r0 = find("DCF", f"情境 {s}")
    DR[s] = {k[0]: find("DCF", k[1], r0, prefix=len(k) > 2) for k in DCFK}
DRATE = {k: find("DCF", t, prefix=True) for k, t in (("ke", "ke＝rf"), ("kd_at", "稅後 kd"), ("E_v", "股權市值"), ("D_v", "負債（債務"), ("w_e", "權益占比"),
                                                     ("wacc", "WACC"), ("g", "永續成長率 g"), ("ronic", "RONIC"), ("nc", "淨現金與投資"))}

SOK = [("a_rev", "車輛業務（非 FSD／Robotaxi）2035 營收"), ("a_val", "價值＝營收 × npm_auto × pe_auto"), ("f_rev", "FSD 2035 營收"), ("f_ps", "P/S＝FY2035 年增率"),
       ("f_val", "價值"), ("n_rev", "非汽車業務（能源＋服務）2035 營收"), ("n_val", "價值＝營收 × npm_non × pe_non"), ("o_rev", "Optimus 2035 營收"),
       ("o_ps", "P/S＝FY2035 年增率"), ("o_val", "價值"), ("v35", "2035 年價值合計（不含 Robotaxi）"), ("pv", "折回現值"),
       ("rt", "Robotaxi 分部價值（現值", True), ("eq", "分部加總股權價值（現值）"), ("vfr", "CEO 獎酬歸屬比例", True), ("vsh", "CEO 獎酬淨增股數"),
       ("shv", "評價股數"), ("vps", "每股價值（現值）"), ("v12", "分部加總推到 12 個月後", True), ("rtps", "其中 Robotaxi 分部價值貢獻", True)]
SR, _r = {}, 4
for k in SOK:
    _r = find("SOTP", k[1], _r, prefix=True)
    SR[k[0]] = _r
    _r += 1

VK = [("dcf", "DCF 情境價值推到 12 個月後"), ("sotp", "分部加總推到 12 個月後"), ("w", "情境權重"), ("tps", "情境目標價（DCF 與分部加總加權）"),
      ("dcf_w", "DCF 目標價（情境加權）"), ("sotp_w", "分部加總目標價（情境加權", True), ("tp", "目標價（DCF 與分部加總加權）"), ("px", "股價（股價日）"),
      ("up", "上檔空間（不是年化報酬）"), ("rate", "估值評等（只看上檔空間）"), ("gate", "收支關卡（FY2026–FY2030", True), ("gate_b", "入不敷出判定用", True),
      ("rec", "投資建議"), ("sent", "條件式投資建議句"), ("dist", "距最近級距邊界", True), ("up1", "升一級所需目標價"), ("dn1", "降一級的目標價", True),
      ("at_d", "上檔空間歸因：DCF 貢獻", True), ("at_s", "上檔空間歸因：分部加總貢獻", True), ("at_m", "上檔空間歸因：相對倍數調整貢獻", True),
      ("rtsh", "Robotaxi 分部價值（讀分部", True), ("mn", "情境與方法區間：最低"), ("mx", "情境與方法區間：最高"), ("rng_chk", "檢查：基準落在區間內"),
      ("per", "期間：模型折現期間", True), ("hold", "期間：持有期間", True), ("tvsh", "DCF 終值占 EV（情境 B）"), ("minfcf", "收支關卡旁並列：情境 B", True),
      ("nc", "收支關卡旁並列：評價日淨現金與投資")]
VR = {k[0]: find("Val", k[1], prefix=len(k) > 2) for k in VK}

RC = {k: find("Recon", t, prefix=True) for k, t in (
    ("cons", "目標價 vs 市場共識"), ("old_fair", "原檔合理每股現價"), ("rev35", "FY2035 合併營收"), ("inv", "反解（原則"), ("inv_h", "目標"),
    ("req_px", "支持現價（目標價＝股價）"), ("up1", "升一級（估值評等邊界）"), ("dn1", "降一級（估值評等邊界）"), ("opt_v", "Optimus 分部價值（情境 B"),
    ("opt_bom", "Optimus FY2035 每台 BOM"), ("cx26", "FY2026 資本支出（模型）"), ("eps_c", "市場共識 FY2026"))}

SE = wb["Sens"]
SE_HDR = find("Sens", "驅動")
SE_ROWS = []
for r in range(SE_HDR + 1, SE.max_row + 1):
    if lab(SE, r) is None:
        break
    SE_ROWS.append(r)
SE_BASE = find("Sens", "基準加權目標價（建置時）")
SE_STAT = find("Sens", "模型狀態", prefix=True)
SE_BID = find("Sens", "build_id")
SE_DBASE = find("Sens", "基準 DCF 目標價（建置時）")
SE_END = SE_BID
base_tp = SE.cell(SE_BASE, 4).value

RD = {lab(wb["README"], r): r for r in range(1, wb["README"].max_row + 1) if lab(wb["README"], r)}
RD_DATE = find("README", "資料基準日")
RB = wb["對帳"]
RB_END = max(r for r in range(4, RB.max_row + 1) if lab(RB, r))
RK = wb["風險登錄"]
RK_END = max(r for r in range(4, RK.max_row + 1) if lab(RK, r))
CHW = wb[CH_SHEET]
CH_END = max(r for r in range(4, CHW.max_row + 1) if CHW.cell(r, 2).value)
HI_SEG = find("Hist", "期間")
HI_END = find("Hist", "FY2025", HI_SEG)

# ---------------------------------------------------------------- 小工具
UF = {"%": "p1", "pt": "p1", "x": "n2", "US$": "n2", "US$M": "n0", "US$B": "n1", "千美元/台": "n1", "台": "n0", "年": "n0", "0/1": "n0", "US$/小時": "n2", "日期": ""}


def v(k, col="C", f=""):
    return f"@Val!{col}{VR[k]}" + (f"|{f}" if f else "")


def inrow(k):
    r = IN[k]
    f = UF.get(lab(wb["In"], r, 3), "")
    return [f"@In!B{r}", f"@In!D{r}|{f}", f"@In!E{r}|{f}", f"@In!F{r}|{f}", f"@In!G{r}|tag"]


IN_HEAD = ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"]


def param_blocks(title, keys):
    """有敏感度區間的參數直接顯示；其餘固定輸入摺疊（規格 v1.0 驅動章分塊）。"""
    ws = wb["In"]
    rng = [k for k in keys if ws.cell(IN[k], 4).value is not None or ws.cell(IN[k], 6).value is not None]
    fix = [k for k in keys if k not in rng]
    out = [{"type": "h3", "text": title}]
    if rng:
        out.append({"type": "table", "caption": "有敏感度區間的參數（基準進主模型；低／高只用於敏感度頁）", "head": IN_HEAD, "rows": [inrow(k) for k in rng]})
    if fix:
        out.append({"type": "table", "collapse": True, "summary": "展開：本業務的固定輸入（無敏感度區間）", "head": IN_HEAD, "rows": [inrow(k) for k in fix]})
    return out


def pl_row(s, k, f, cols=PLC):
    return [f"@PL!A{PR[s][k]}"] + [f"@PL!{c}{PR[s][k]}|{f}" for c in cols]


def met(name, key, fmt, chart=False):
    m = {"name": name, "fmt": fmt, "rows": {s: f"@PL!C{PR[s][key]}:M{PR[s][key]}" for s in "ABC"}}
    if chart:
        m["chart"] = True
    return m


YEARS = f"@PL!C{PLH}:M{PLH}"
YEARS_F = f"@PL!E{PLH}:M{PLH}"
PXLAB = ["現價（", f"@In!E{IN['px_date']}", " 收盤）"]

# ---------------------------------------------------------------- 頂層
cfg = {
    "company": "Tesla", "ticker": "TSLA", "layout": "v1.3", "wide_text": True,
    "thesis": f"@README!B{RD['命題']}",
    "question": "AI 進入實體世界（FSD、Robotaxi、Optimus）的價值有多少已反映在現價，而不是汽車與能源本業？",
    "scenarios": [{"key": "A", "name": "情境 A（Optimus 與 Robotaxi 放量接近公司說法）"}, {"key": "B", "name": "情境 B（基準）"},
                  {"key": "C", "name": "情境 C（Optimus 可靠度未解決、Robotaxi 規模縮小）"}],
    "hero": {"rec": v("rec"),
             "tiles": [[f"@Val!A{VR['tp']}", v("tp", "C", "n2"), ["上檔空間 ", v("up", "C", "sp1")]],
                       [f"@Val!A{VR['sotp_w']}", v("sotp_w", "C", "n2"), ["DCF 目標價 ", v("dcf_w", "C", "n2")]],
                       [PXLAB, v("px", "C", "n2"), ""]],
             "note": ["加權目標價＝DCF 與分部加總加權（DCF 權重 ", f"@In!E{IN['w_dcf']}|p0", "）；Robotaxi 分部占加權目標價 ", v("rtsh", "C", "p1"),
                      "。收支關卡（FY2026–FY2030 情境 B 自由現金流合計）", v("gate_b", "C", "n0"), " 百萬美元，", v("gate_b", "D"),
                      "。市場共識目標價為前版對照值（擷取日見估值章），未更新。"]},
    "sections": [],
}

# ---------------------------------------------------------------- 1 命題
B = PR["B"]
thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "合併營收（百萬美元，依分部與命題子列堆疊）與毛利率：FY2025–FY2035（情境 B）", "x": YEARS, "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.25,
     "total": f"@PL!C{B['rev']}:M{B['rev']}",
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": f"@PL!C{B[k]}:M{B[k]}", "color": c} for k, c in
              (("veh", "#b7e3e4"), ("crd", "#d9b65c"), ("lease", "#8a8c8c"), ("svc", "#6fc7cb"), ("seg_en", "#10a0a7"), ("fsd", "#0b7a80"), ("rt", "#c29a2b"), ("seg_opt", "#6f6bae"))],
     "lines": [{"name": [f"@PL!A{B['gm']}"], "y": f"@PL!C{B['gm']}:M{B['gm']}", "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝汽車分部各子列（車輛銷售、碳權、租賃、服務及其他、FSD、Robotaxi）、能源分部與 Optimus（尚未列報），合計等於合併營收；折線＝毛利率（右軸）。FSD、Robotaxi、Optimus 三條是命題所指的實體 AI 業務。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [f"@README!B{RD['命題']}"]},
    {"type": "note", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "parts": [f"@README!B{RD['命題回覆']}"]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["估值：加權目標價 ", v("tp", "C", "n2"), " 美元，相對現價 ", v("px", "C", "n2"), " 上檔空間 ", v("up", "C", "sp1"), "，估值評等", v("rate"),
         "；距最近級距邊界 ", v("dist", "C", "p1"), "，目標價要到 ", v("up1", "C", "n2"), " 才升一級。"],
        ["現金流折現腿三情境（A／B／C）每股 ", v("dcf", "C", "n2"), "／", v("dcf", "D", "n2"), "／", v("dcf", "E", "n2"),
         "；情境 C 股權價值為負、依限責處理歸零。汽車本業毛利率偏低、資本支出高，明確期自由現金流撐不起現價。"],
        ["分部加總腿三情境 ", v("sotp", "C", "n2"), "／", v("sotp", "D", "n2"), "／", v("sotp", "E", "n2"), "；其中 Robotaxi 分部價值（十億美元，讀 Robotaxi 分部模型）",
         f"@In!E{IN['rt_val_A']}|n1", "／", f"@In!E{IN['rt_val_B']}|n1", "／", f"@In!E{IN['rt_val_C']}|n1", "，占加權目標價 ", v("rtsh", "C", "p1"), "；分部加總的主體是 FSD 與 Optimus。"],
        ["反解：要支持現價，情境 B 的 Robotaxi 分部價值約需 ", f"@Recon!D{RC['req_px']}|n0", " 十億美元，分部模型值為 ", f"@In!E{IN['rt_val_B']}|n1", "。"],
        ["收支關卡：FY2026–FY2030 情境 B 自由現金流合計 ", v("gate_b", "C", "n0"), " 百萬美元（", v("gate_b", "D"), "），最低單年 ", v("minfcf", "C", "n0"),
         "；評價日淨現金與投資 ", v("nc", "C", "n0"), " 百萬美元。投資建議", v("rec"), "。"]]}]}

# ---------------------------------------------------------------- 2 估值
dcf_main = ["sumpv", "tvpv", "ev", "eq", "vps", "v12"]
dcf_more = ["tv", "tvsh", "eqT", "vfr", "vsh", "shv"]
DF = {"tvsh": "p1", "eqT": "n3", "vfr": "p1", "vsh": "n1", "shv": "n1", "vps": "n2", "v12": "n2"}


def dcf_rows(keys):
    return [[f"@DCF!A{DR['B'][k]}"] + [f"@DCF!C{DR[s][k]}|{DF.get(k, 'n0')}" for s in "ABC"] for k in keys]


SF = {"f_ps": "n1", "o_ps": "n1", "vfr": "p1", "vsh": "n1", "shv": "n1", "vps": "n2", "v12": "n2", "rtps": "n2"}


def sotp_rows(keys):
    return [[f"@SOTP!A{SR[k]}"] + [f"@SOTP!{c}{SR[k]}|{SF.get(k, 'n0')}" for c in "CDE"] for k in keys]


valuation = {"id": "valuation", "title": "估值：現金流折現與分部加總兩腿", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價；市場共識為前版對照值，未更新）", "fmt": "n2",
     "items": [{"label": f"@Val!A{VR['dcf_w']}", "value": v("dcf_w"), "series": 1},
               {"label": f"@Val!A{VR['sotp_w']}", "value": v("sotp_w"), "series": 2},
               {"label": f"@Val!A{VR['tp']}", "value": v("tp"), "series": 0},
               {"label": f"@In!B{IN['cons_tp']}", "value": f"@In!E{IN['cons_tp']}", "series": 1}],
     "ref": {"label": PXLAB, "value": v("px")}},
    {"type": "note", "parts": [f"@Val!A1"]},
    {"type": "note", "parts": ["市場共識目標價：", f"@In!H{IN['cons_tp']}", "（", f"@In!I{IN['cons_tp']}", "）。"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", v(k, "C", f)] for k, f in (("dcf_w", "n2"), ("sotp_w", "n2"), ("tp", "n2"), ("px", "n2"), ("up", "sp1"), ("rate", ""), ("gate_b", "n0"), ("rec", ""))]
             + [[f"@In!B{IN['w_dcf']}", f"@In!E{IN['w_dcf']}|p0"], [f"@In!B{IN['cons_tp']}", f"@In!E{IN['cons_tp']}|n2"]]},
    {"type": "note", "parts": [v("sent")]},
    {"type": "table", "caption": "三情境：兩腿每股價值與情境目標價（美元；最後一欄為情境加權）", "head": [f"@Val!A3", f"@Val!C3", f"@Val!D3", f"@Val!E3", f"@Val!F3"],
     "rows": [[f"@Val!A{VR['dcf']}"] + [f"@Val!{c}{VR['dcf']}|n2" for c in "CDEF"], [f"@Val!A{VR['sotp']}"] + [f"@Val!{c}{VR['sotp']}|n2" for c in "CDEF"],
              [f"@Val!A{VR['w']}"] + [f"@Val!{c}{VR['w']}|p0" for c in "CDEF"], [f"@Val!A{VR['tps']}"] + [f"@Val!{c}{VR['tps']}|n2" for c in "CDE"]]},
    {"type": "h3", "text": "分部加總中 Robotaxi 的比重"},
    {"type": "table", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", v(k, "C", "sp1" if k.startswith("at_") else "p1")] for k in ("rtsh", "at_d", "at_s", "at_m")]
             + [[f"@SOTP!A{SR['rtps']}", f"@SOTP!C{SR['rtps']}|n2", f"@SOTP!D{SR['rtps']}|n2", f"@SOTP!E{SR['rtps']}|n2"]]},
    {"type": "h3", "text": "現金流折現：折現率"},
    {"type": "kv", "items": [[f"@DCF!A{DRATE[k]}", f"@DCF!C{DRATE[k]}|p2"] for k in ("ke", "kd_at", "wacc", "ronic", "g")]},
    {"type": "table", "collapse": True, "summary": "展開：折現率組成（無風險利率、β、股權風險溢酬、資本結構）",
     "head": IN_HEAD, "rows": [inrow(k) for k in ("rf", "beta_y", "beta_s", "beta", "erp", "kd", "tax_w", "wacc_adj")]
     + [[f"@DCF!A{DRATE[k]}", "", f"@DCF!C{DRATE[k]}|{'p1' if k == 'w_e' else 'n0'}", "", ""] for k in ("E_v", "D_v", "w_e", "nc")]},
    {"type": "table", "caption": "現金流折現彙總（三情境；百萬美元，每股為美元）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rows": dcf_rows(dcf_main)},
    {"type": "table", "collapse": True, "summary": "展開：現金流折現其餘項目（終值、終值占企業價值、CEO 獎酬股數）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rows": dcf_rows(dcf_more)},
    {"type": "table", "caption": "分部加總彙總（沿用原檔 2035 年分部倍數法；百萬美元，每股為美元）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": sotp_rows(["a_val", "f_val", "n_val", "o_val", "v35", "pv", "rt", "eq", "v12"])},
    {"type": "table", "collapse": True, "summary": "展開：分部加總其餘項目（各分部營收、市銷率、股數）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
     "rows": sotp_rows(["a_rev", "f_rev", "f_ps", "n_rev", "o_rev", "o_ps", "vfr", "vsh", "shv", "vps"])},
    {"type": "table", "collapse": True, "summary": "展開：分部加總倍數與權重輸入", "head": IN_HEAD,
     "rows": [inrow(k) for k in ("sotp_r", "sotp_n", "pe_auto", "npm_auto", "pe_non", "npm_non", "peg_k", "tp_wA", "tp_wB", "tp_wC", "w_dcf")]},
    {"type": "table", "collapse": True, "summary": "展開：評等邊界、區間與期間", "head": ["項目", "數值"],
     "rows": [[f"@Val!A{VR[k]}", v(k, "C", f)] for k, f in (("dist", "p1"), ("up1", "n2"), ("dn1", "n2"), ("mn", "n2"), ("mx", "n2"), ("rng_chk", ""), ("per", ""), ("hold", "n2"), ("tvsh", "p1"))]}]}

# ---------------------------------------------------------------- 3 敏感度
cases = [r for r in SE_ROWS if isinstance(SE.cell(r, 4).value, (int, float)) or isinstance(SE.cell(r, 5).value, (int, float))]


def sens_bars(side_col, side):
    rs = [r for r in cases if isinstance(SE.cell(r, side_col).value, (int, float))]
    rs.sort(key=lambda r: -abs(SE.cell(r, side_col).value / base_tp - 1))
    big = [r for r in rs if abs(SE.cell(r, side_col).value / base_tp - 1) >= 0.1]
    small = [r for r in rs if r not in big and abs(SE.cell(r, side_col).value / base_tp - 1) >= 0.0005]
    col = openpyxl.utils.get_column_letter(side_col)
    mk = lambda rows, t: {"type": "bars", "title": t, "fmt": "n2",
                          "items": [{"label": f"@Sens!A{r}", "value": f"@Sens!{col}{r}", "series": 1 if SE.cell(r, side_col).value < base_tp else 0} for r in rows],
                          "ref": {"label": f"@Sens!A{SE_BASE}", "value": f"@Sens!D{SE_BASE}"}}
    main = mk(big or rs[:1], f"加權目標價相對基準：輸入取{side}值，變動達百分之十以上的案例（美元；豎線＝基準加權目標價）")
    rest = dict(mk(small, f"加權目標價相對基準：輸入取{side}值，其餘案例"), collapse=True, summary=f"展開：輸入取{side}值，變動未達百分之十的案例") if small else None
    return main, rest


lo_main, lo_rest = sens_bars(4, "低")
hi_main, hi_rest = sens_bars(5, "高")
sens = {"id": "sens", "title": "敏感度", "blocks": [lo_main, hi_main,
    {"type": "note", "parts": ["@Sens!A1"]}, {"type": "note", "parts": ["@Sens!A2"]},
    {"type": "note", "parts": ["本模型的敏感度頁列加權目標價的絕對值；圖以基準加權目標價為豎線，長條在線左側為低於基準、右側為高於基準。"]}]
    + [b for b in (lo_rest, hi_rest) if b] + [
    {"type": "table", "collapse": True, "summary": "展開：全部案例（加權目標價、現金流折現目標價與估值評等）",
     "head": [f"@Sens!{c}{SE_HDR}" for c in "ABDEFGHI"],
     "rowgen": {"sheet": "Sens", "rows": f"{SE_ROWS[0]}-{SE_ROWS[-1]}", "cols": list("ABDEFGHI"), "fmts": ["", "", "n2", "n2", "n2", "n2", "", ""]}},
    {"type": "kv", "items": [[f"@Sens!A{SE_BASE}", f"@Sens!D{SE_BASE}|n2"], [f"@Sens!A{SE_DBASE}", f"@Sens!D{SE_DBASE}|n2"], ["模型狀態", f"@Sens!D{SE_STAT}"], ["建置編號", f"@Sens!D{SE_BID}"]]}]}

# ---------------------------------------------------------------- 4 反向檢驗
revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含的 Robotaxi 分部價值", "blocks": [
    {"type": "note", "parts": ["本模型未建整本反向現金流折現；以下用分部加總腿的線性反解回答「現價隱含什麼」：其他假設不變，三情境的 Robotaxi 分部價值同加多少，加權目標價才等於現價或評等邊界。"]},
    {"type": "table", "caption": f"@Recon!A{RC['inv']}", "head": [f"@Recon!{c}{RC['inv_h']}" for c in "ABCDE"],
     "rows": [[f"@Recon!A{r}", f"@Recon!B{r}|n2", f"@Recon!C{r}|n0", f"@Recon!D{r}|n0", f"@Recon!E{r}"] for r in (RC['req_px'], RC['up1'], RC['dn1'])]},
    {"type": "kv", "items": [[f"@In!B{IN['rt_val_B']}", f"@In!E{IN['rt_val_B']}|n1"], [f"@Val!A{VR['up1']}", v("up1", "C", "n2")]]}]}

# ---------------------------------------------------------------- 5 損益
STD = [("rev", "n0"), ("gp", "n0"), ("gm", "p1"), ("opex", "n0"), ("ebit", "n0"), ("nonop", "n0"), ("pti", "n0"), ("tax", "n0"), ("mi", "n0"), ("ni", "n0"), ("sh", "n0"), ("eps", "n2")]
CF = [("da", "n0"), ("sbc", "n0"), ("cx", "n0"), ("wc", "n0"), ("fcf", "n0")]
pl = {"id": "pl", "title": "損益（情境 B）", "blocks": [
    {"type": "note", "parts": ["損益表採美國會計準則（GAAP）口徑：Tesla 財報與本模型原檔引擎皆以 GAAP 淨利為主。FY2025 為實績、FY2026 含前三季實績；FY2027 起依輸入頁調整。本模型損益頁未列營收年增率、營業成本、營業利益率、淨利率與淨利年增率，也只有一個完整歷史年度；本版不補列。"]},
    {"type": "theme", "title": "合併營收（百萬美元）與毛利率：FY2025–FY2035（情境 B）", "x": YEARS, "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.5,
     "bars": [{"name": [f"@PL!A{B['rev']}"], "y": f"@PL!C{B['rev']}:M{B['rev']}", "color": "#6fc7cb"}],
     "lines": [{"name": [f"@PL!A{B['gm']}"], "y": f"@PL!C{B['gm']}:M{B['gm']}", "axis": "r", "color": "#0b7a80", "labels": True}],
     "caption": "長條＝合併營收；折線＝毛利率（右軸，GAAP）。"},
    {"type": "table", "caption": "損益表（GAAP；百萬美元，EPS 為美元）：FY2025 實績、FY2026 含前三季實績、FY2027 起情境 B", "head": [f"@PL!A{PLH}", YEARS],
     "rows": [["營收（合併）"] + [f"@PL!{c}{B['rev']}|n0" for c in PLC]] + [pl_row("B", k, f) for k, f in STD[1:]]
             + [{"sub": "現金流（收支關卡用）"}] + [pl_row("B", k, f) for k, f in CF]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線）", "years": YEARS,
     "metrics": [met("合併營收（百萬美元）", "rev", "n0", True), met("營業利益（百萬美元）", "ebit", "n0", True), met("毛利率", "gm", "p1"),
                 met("EPS（稀釋，美元）", "eps", "n2", True), met("自由現金流（百萬美元）", "fcf", "n0", True)]},
    {"type": "table", "caption": "檢查：分部合計 − 合併營收（三情境，應為零）", "head": [f"@PL!A{PLH}", YEARS],
     "rows": [[f"情境 {s}"] + [f"@PL!{c}{PR[s]['chk']}|n1" for c in PLC] for s in "ABC"]}]}

# ---------------------------------------------------------------- 6 驅動
drivers = {"id": "drivers", "title": "驅動因子：分部營收推導", "blocks": [
    {"type": "theme", "title": "合併營收 × 情境 A／B／C（百萬美元）：FY2027–FY2035", "x": YEARS_F, "fmt": "n0", "stack": False, "pad": 1.15,
     "bars": [{"name": ["情境 ", s, "：合併營收"], "y": f"@PL!E{PR[s]['rev']}:M{PR[s]['rev']}", "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
     "caption": "三情境只在 Robotaxi 營收（讀 Robotaxi 分部模型）與 Optimus 量倍數上分化；車輛、FSD、能源、服務三情境相同。"},
    {"type": "note", "parts": [f"@README!B{RD['分部']}"]},
    {"type": "table", "caption": "分部營收推導（情境 B；百萬美元）：汽車分部子列＋能源分部＋Optimus＝合併營收", "head": [f"@PL!A{PLH}", f"@PL!B{PLH}", YEARS],
     "rows": [[f"@PL!A{B[k]}", f"@PL!B{B[k]}"] + [f"@PL!{c}{B[k]}|{'n1' if k == 'chk' else 'n0'}" for c in PLC] for k in
              ("seg_auto", "veh", "fsd", "rt", "crd", "lease", "svc", "seg_en", "seg_opt", "rev", "chk")]},
    {"type": "table", "collapse": True, "summary": "展開：情境 A、C 的分部營收", "head": [f"@PL!A{PLH}", YEARS],
     "rows": [{"sub": f"@PL!A{find('PL', '情境 ' + s)}"} if k is None else pl_row(s, k, "n0") for s in "AC" for k in
              (None, "seg_auto", "veh", "fsd", "rt", "crd", "lease", "svc", "seg_en", "seg_opt", "rev")]},
    {"type": "table", "caption": "三情境驅動（In 頁情境驅動對照）", "head": [f"@In!B{Z0 + 1}", f"@In!C{Z0 + 1}", f"@In!D{Z0 + 1}", f"@In!E{Z0 + 1}", f"@In!F{Z0 + 1}"],
     "rows": [[f"@In!B{INZ[k]}", f"@In!C{INZ[k]}"] + [f"@In!{c}{INZ[k]}|{UF.get(lab(wb['In'], INZ[k], 3), '')}" for c in "DEF"]
              for k in ["opt_k"] + [f"rt_rev_{y}" for y in range(2027, 2036)] + ["rt_val", "scn_w"]]},
    {"type": "table", "collapse": True, "summary": "展開：歷史分部營收對帳（季度與年度；百萬美元）", "head": [f"@Hist!{c}{HI_SEG}" for c in "ABCDEFGJ"],
     "rowgen": {"sheet": "Hist", "rows": f"{HI_SEG + 1}-{HI_END}", "cols": list("ABCDEFGJ"), "fmts": ["", "n0", "n0", "n0", "n0", "n0", "n1", "n1"]}}]
    + param_blocks("一、汽車與 FSD", ["auto_gm_adj", "fsd_k", "fsd_gm", "tax_t"])
    + param_blocks("二、現金流與資本支出", ["cx_26", "cx_r27", "cx_ss", "life", "wc_r", "tail_tfab", "sbc_r"])
    + param_blocks("三、股數與 CEO 獎酬", ["ceo_ovr"])
    + param_blocks("四、全公司：折現率", ["rf", "beta_y", "beta_s", "beta", "erp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "tv_alt"])}

# ---------------------------------------------------------------- 7 個別模組：Robotaxi
robotaxi = {"id": "robotaxi", "title": "Robotaxi 分部（讀 Robotaxi 分部模型）", "blocks": [
    {"type": "theme", "title": "Robotaxi 營收 × 情境 A／B／C（百萬美元）：FY2027–FY2035", "x": YEARS_F, "fmt": "n0", "stack": False, "pad": 1.1,
     "bars": [{"name": ["情境 ", s, "：Robotaxi 營收"], "y": f"@PL!E{PR[s]['rt']}:M{PR[s]['rt']}", "color": c} for s, c in (("A", "#c29a2b"), ("B", "#3b3838"), ("C", "#d9a08f"))],
     "caption": "Robotaxi 營收逐年讀 Robotaxi 分部模型（自有車隊全額＋車主車輛抽成），乘敏感度倍數（基準為一）；FY2026 以前沿用原檔。"},
    {"type": "table", "caption": "Robotaxi 分部價值（現值，十億美元；進分部加總腿）", "head": IN_HEAD,
     "rows": [inrow(f"rt_val_{s}") for s in "ABC"]},
    {"type": "table", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rows": sotp_rows(["rt", "rtps"])},
    {"type": "note", "parts": [f"@對帳!H{find('對帳', 'Robotaxi 每英里營運成本')}"]},
    {"type": "table", "caption": "Robotaxi 相關輸入（敏感度與毛利調整）", "head": IN_HEAD, "rows": [inrow(k) for k in ("rt_k", "rt_gm")]},
    {"type": "table", "collapse": True, "summary": "展開：Robotaxi 對帳變數（車隊、票價、單量、成本、分部價值與營收）", "head": [f"@對帳!{c}3" for c in "ACDGH"],
     "rows": [[f"@對帳!{c}{r}" for c in "ACDGH"] for r in range(4, RB_END + 1) if lab(RB, r, 2) == "1a"]}]}

# ---------------------------------------------------------------- 7 個別模組：Optimus
OB = OR["B"]
optimus = {"id": "optimus", "title": "Optimus 分部（人形機器人單位經濟）", "blocks": [
    {"type": "theme", "title": "Optimus 營收（百萬美元，情境 B）與每台成本、售價（千美元，右軸）：FY2027–FY2035", "x": YEARS_F, "fmt": "n0", "fmt_r": "n0", "stack": True, "pad": 1.2,
     "bars": [{"name": [f"@Optimus!A{OB['rev']}"], "y": f"@Optimus!E{OB['rev']}:M{OB['rev']}", "color": "#6f6bae"}],
     "lines": [{"name": [f"@Optimus!A{O_ASP}"], "y": f"@Optimus!E{O_ASP}:M{O_ASP}", "axis": "r", "color": "#0b7a80", "labels": True},
               {"name": [f"@Optimus!A{OB['cost']}"], "y": f"@Optimus!E{OB['cost']}:M{OB['cost']}", "axis": "r", "color": "#b4472f", "labels": "below"},
               {"name": [f"@Optimus!A{OB['bom']}"], "y": f"@Optimus!E{OB['bom']}:M{OB['bom']}", "axis": "r", "color": "#c29a2b", "dash": True}],
     "caption": "長條＝Optimus 營收；折線＝平均售價、每台製造成本與每台 BOM（學習曲線，右軸）。Tesla 未揭露任何 Optimus 財務數字，全部為假設。"},
    {"type": "note", "parts": ["@Optimus!A1"]},
    {"type": "table", "caption": "Optimus 單位經濟（情境 B）", "head": [f"@Optimus!A3", f"@Optimus!B3", YEARS_F],
     "rows": [[f"@Optimus!A{O_ASP}", f"@Optimus!B{O_ASP}"] + [f"@Optimus!{c}{O_ASP}|n1" for c in FC]]
             + [[f"@Optimus!A{OB[k]}", f"@Optimus!B{OB[k]}"] + [f"@Optimus!{c}{OB[k]}|{'p1' if k == 'gm' else ('n1' if k in ('bom', 'cost') else 'n0')}" for c in FC] for k, _ in OPK]},
    {"type": "table", "collapse": True, "summary": "展開：情境 A、C 的 Optimus 單位經濟", "head": [f"@Optimus!A3", f"@Optimus!B3", YEARS_F],
     "rows": [{"sub": f"@Optimus!A{find('Optimus', '情境 ' + s)}"} if k is None else
              [f"@Optimus!A{OR[s][k]}", f"@Optimus!B{OR[s][k]}"] + [f"@Optimus!{c}{OR[s][k]}|{'p1' if k == 'gm' else ('n1' if k in ('bom', 'cost') else 'n0')}" for c in FC]
              for s in "AC" for k in [None] + [kk for kk, _ in OPK]]},
    {"type": "table", "caption": "對人工的 TCO 與回本期（展示，不進目標價；情境 B）", "head": [f"@Optimus!A3", f"@Optimus!B3", YEARS_F],
     "rows": [[f"@Optimus!A{r}", f"@Optimus!B{r}"] + [f"@Optimus!{c}{r}|n2" for c in FC] for r in range(O_TCO + 1, O_TCO + 4)]},
    {"type": "note", "parts": [f"@Optimus!A{O_TCO + 5}"]},
    {"type": "table", "caption": "分部加總中的 Optimus 與校準", "head": ["項目", "本模型", "對照值"],
     "rows": [[f"@Recon!A{r}", f"@Recon!B{r}|{f}", f"@Recon!C{r}|n1"] for r, f in ((RC['opt_v'], "n2"), (RC['opt_v'] + 1, "p1"), (RC['opt_v'] + 2, "p1"), (RC['opt_bom'], "n1"))]}]
    + param_blocks("Optimus 參數", ["opt_asp27", "opt_asp35", "opt_bom0", "opt_q0", "opt_lr", "opt_oh", "opt_sga", "opt_cx", "opt_k_A", "opt_k_B", "opt_k_C", "opt_wage", "opt_prod"]
                   + [f"opt_u_{y}" for y in range(2027, 2036)])}

# ---------------------------------------------------------------- 8 對帳
recon = {"id": "recon", "title": "對帳：原檔、共識與線層級變數", "blocks": [
    {"type": "note", "parts": ["@Recon!A1"]},
    {"type": "table", "caption": "本模型與原檔、市場共識對照（共識為前版對照值，未更新）", "head": [f"@Recon!{c}3" for c in "ABCDE"],
     "rows": [[f"@Recon!A{r}", f"@Recon!B{r}|n2", f"@Recon!C{r}|n2", f"@Recon!D{r}|n2", f"@Recon!E{r}"] for r in range(RC['cons'], RC['rev35'] + 1)]},
    {"type": "table", "caption": "其他對照", "head": ["項目", "本模型", "對照值", "說明"],
     "rows": [[f"@Recon!A{r}", f"@Recon!B{r}|n0", f"@Recon!C{r}|n0", f"@Recon!E{r}"] for r in (RC['cx26'], RC['eps_c'])]},
    {"type": "note", "parts": ["@對帳!A1"]},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（FSD、Robotaxi、人形機器人、訓練算力）", "head": [f"@對帳!{c}3" for c in "ABCDGH"],
     "rowgen": {"sheet": "對帳", "rows": f"4-{RB_END}", "cols": list("ABCDGH")}}]}

# ---------------------------------------------------------------- 9 資料狀態
status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["build_id", f"@Sens!D{SE_BID}"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "h3", "text": "方法、判斷題與限制"},
    {"type": "list", "items": [f"@README!B{RD[k]}" for k in ("方法", "判斷題", "限制", "主要失效條件", "可證偽指標（承重 Assumed）")]},
    {"type": "table", "caption": "資料基準日", "head": ["項目", "日期與來源"], "rowgen": {"sheet": "README", "rows": f"{RD_DATE + 1}-{RD_DATE + 5}", "cols": ["A", "B"]}},
    {"type": "table", "caption": "風險登錄", "head": [f"@風險登錄!{c}3" for c in "ABCD"], "rowgen": {"sheet": "風險登錄", "rows": f"4-{RK_END}", "cols": list("ABCD")}},
    {"type": "table", "collapse": True, "summary": "展開：改版紀錄", "head": [f"@{CH_SHEET}!{c}3" for c in "BCDEF"],
     "rowgen": {"sheet": CH_SHEET, "rows": f"4-{CH_END}", "cols": list("BCDEF")}},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"Recon": ["A"], "對帳": ["A"]}, "empty": "沒有標為未取得的項目。"}]}

cfg["sections"] = [thesis, valuation, sens, revdcf, pl, drivers, robotaxi, optimus, recon, status]
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)

# 聯電：把 html_src.json（v1.0 之前的頁面設定，內容不動）依「個股 HTML 版面規格 v1.0」（docs/plan/個股HTML版面規格.md）轉成 html.json。
# 用法：python3 umc/gen_html_json.py <成品 xlsx>（讀 umc/html_src.json，寫 umc/html.json；勿手改 html.json）
import copy, json, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
import relayout as RL  # noqa: E402

X = sys.argv[1]
WB = openpyxl.load_workbook(X, data_only=True)
cfg = copy.deepcopy(json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8")))
SEC = {s["id"]: s for s in cfg["sections"]}


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def row_of(sh, prefix, need_num=False, exact=False):
    ws = WB[sh]
    for r in range(1, ws.max_row + 1):
        t = str(ws.cell(r, 1).value or "").strip()
        if (t == prefix if exact else t.startswith(prefix)) and (not need_num or num(ws.cell(r, 2).value)):
            return r
    raise KeyError((sh, prefix))


I = {str(WB["In"].cell(r, 1).value): r for r in range(4, WB["In"].max_row + 1) if WB["In"].cell(r, 1).value}


def in_tok(prefix):
    ws = WB["In"]
    for r in range(4, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "")
        if a.startswith(prefix) and ws.cell(r, 2).value in (None, ""):
            return re.split(r"[.\s]", a)[0]
    raise KeyError(prefix)


# ─── 驅動因子：舊「輸入 → 推導量」表移到附錄（轉換器依 caption「驅動 → 推導」辨認） ───
for b in SEC["drivers"]["blocks"]:
    if b["type"] == "table" and not b.get("caption") and b.get("head", [None])[0] == "輸入":
        b["caption"] = "驅動 → 推導對照（單一參數：輸入、基準值、推導量與因果方向）"

# ─── 估值：長條圖 → 結論主要項目 → 投資建議 → DCF 折現率 → DCF 彙總 → Comparable → 其餘摺疊 ───
V = lambda k: row_of("Val", k, need_num=True)
tpw, dcft, comp, compadj, cons = V("目標價（DCF 與 Comparable 加權）"), V("DCF 目標價（情境加權）"), V("Comparable 目標價（市場主流同業法）"), V("Comparable 目標價（相對倍數調整後）"), V("市場共識目標價")
upw, upc, dcf12, px = V("上檔空間：加權目標價"), V("上檔空間：Comparable"), V("DCF 情境價值推到 12 個月後"), V("現價")
rec, grade = row_of("Val", "投資建議", exact=True), row_of("Val", "上檔空間級距", exact=True)
rk = row_of("Val", "股權成本 ke")                       # v0.2：0.1 節 ke、kd_at、w_e、WACC、g、RONIC，接著 DCF 彙總 10 列
assert str(WB["Val"].cell(rk + 15, 1).value).startswith("DCF 目標價（情境加權）")
fb = next(r for r in range(row_of("DCF", "情境 B（平平）"), WB["DCF"].max_row + 1) if WB["DCF"].cell(r, 1).value == "營收")
hc = row_of("README", "展示層輔助：Comparable 目標價（相對倍數調整後）的上檔空間")
c0 = row_of("Val", "0.3 結論")
old = SEC["valuation"]["blocks"]
pick = lambda f: next(b for b in old if f(b))
bars = pick(lambda b: b["type"] == "bars")
for _it in bars["items"]:   # v0.2：0.1 節改版後 DCF 目標價列移到 dcft
    if _it["value"] == "@Val!B24": _it.update(label=f"@Val!A{dcft}", value=f"@Val!B{dcft}")
dcf_tbl = pick(lambda b: b.get("caption") == "@Val!A5")
peer_tbl = pick(lambda b: b.get("caption") == "@Val!A26")
comp_kv = pick(lambda b: b["type"] == "kv" and b["items"][0][0] == "@Val!A33")
px_kv = pick(lambda b: b["type"] == "kv" and b["items"][0][0] == f"@Val!A{px}")
rm_kv = pick(lambda b: b["type"] == "kv" and b["items"][0][0].startswith("@Val!A14"))
rule_tbls = [b for b in old if b["type"] == "table" and not b.get("caption") and not b.get("collapse")]
hist_tbl = pick(lambda b: str(b.get("summary", "")).startswith("展開：十年年底 P/E"))
notes = {b["parts"][0]: b for b in old if b["type"] == "note"}
H5 = ["項目", "數值", "", "", "說明"]
inr = lambda k, f: [f"@In!B{I[k]}", f"@In!E{I[k]}|{f}"]
kv2tbl = lambda kv, cap, **kw: dict({"type": "table", "caption": cap, "head": ["項目", "數值"], "rows": [[k, v] for k, v in kv["items"]]}, **kw)
SEC["valuation"]["blocks"] = [
 bars,
 {"type": "table", "caption": "目標價結論（主要項目）", "head": H5, "rows": [
   [f"@Val!A{tpw}", f"@Val!B{tpw}|n2", "", "", f"@Val!E{tpw}"],
   [f"@Val!A{dcft}", f"@Val!B{dcft}|n2", "", "", f"@Val!E{dcft}"],
   [f"@Val!A{compadj}", f"@Val!B{compadj}|n2", "", "", f"@Val!E{compadj}"],
   [f"@Val!A{comp}", f"@Val!B{comp}|n2", "", "", f"@Val!E{comp}"],
   [f"@Val!A{px}", f"@Val!B{px}|n2", "", "", ""],
   [f"@Val!A{upw}", f"@Val!B{upw}|sp1", "", "", f"@Val!E{upw}"],
   [f"@Val!A{upc}", f"@Val!B{upc}|sp1", "", "", f"@Val!E{upc}"],
   ["Comparable（相對倍數調整後）的上檔空間", f"@README!B{hc}|sp1", "", "", ""],
   [f"@Val!A{rec}", f"@Val!B{rec}", "", "", f"@Val!E{rec}"],
   [f"@Val!A{dcf12}", f"@Val!B{dcf12}|n1", f"@Val!C{dcf12}|n1", f"@Val!D{dcf12}|n1", ""],
   [f"@Val!A{cons}", f"@Val!B{cons}|n2", "", "", f"@Val!E{cons}"]]},
 {"type": "table", "collapse": True, "summary": "展開：結論的其他項目（與共識的差距拆解、價格帶、十年分位、上檔空間級距）", "head": H5, "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|{f}", f"@Val!C{r}|{f}" if r == c0 + 10 else "", f"@Val!D{r}|{f}" if r == c0 + 10 else "", f"@Val!E{r}"]
   for r, f in ((c0 + 7, "sp1"), (c0 + 8, "sp1"), (c0 + 9, ""), (c0 + 10, "p0"))] + [[f"@Val!A{grade}", f"@Val!B{grade}", "", "", f"@Val!E{grade}"]]},
 {"type": "table", "caption": "DCF 折現率（WACC、股權成本、稅後債務成本、RONIC、永續成長率；直接顯示）", "head": ["項目", "數值", "說明"], "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|p2", f"@Val!E{r}"] for r in range(rk, rk + 6)]},
 {"type": "table", "caption": "DCF 彙總（標準版：無槓桿自由現金流 FCFF＋WACC；三情境）", "head": ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"], "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|{fm}", f"@Val!C{r}|{fm}" if r != rk + 15 else "", f"@Val!D{r}|{fm}" if r != rk + 15 else "", f"@Val!E{r}"]
   for r, fm in zip(range(rk + 6, rk + 16), ["n1", "n1", "n1", "n1", "n2", "n2", "n2", "p1", "x1", "n2"])]},
 {"type": "table", "collapse": True, "summary": "展開：情境 B（平平）FCFF 逐年明細（DCF 頁；十億新台幣）", "head": ["@DCF!A5", "@DCF!C5", "@DCF!D5", "@DCF!E5", "@DCF!F5", "@DCF!G5", "@DCF!H5"],
  "rowgen": {"sheet": "DCF", "rows": f"{fb}-{fb + 10}", "cols": ["A", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n1", "n1", "n1", "n1", "n1", "n1"],
             "row_fmt": {str(fb + 3): {c: "p1" for c in "CDEFGH"}, str(fb + 5): {c: "p1" for c in "CDEFGH"}, str(fb + 6): {c: "p1" for c in "CDEFGH"}}}},
 kv2tbl(comp_kv, "Comparable（同業前瞻 P/E 中位數 × 一年後的前瞻 EPS）"),
 dict(peer_tbl, collapse=True, summary="展開：同業前瞻倍數"),
 kv2tbl(rm_kv, "相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司相對同業的長期倍數比）"),
 notes["@Val!A48"],
 kv2tbl(px_kv, "價格與市場隱含倍數", collapse=True, summary="展開：現價與市場隱含倍數"),
 {"type": "h3", "text": "十年 P/E 規則值（舊法參考，不進目標價）"}] + [dict(b, collapse=True, summary=s) for b, s in zip(rule_tbls, ("展開：規則值隱含股價（FY2027 EPS × 依循環位置的倍數）", "展開：循環位置（本模型營收最高年）"))] + [
 notes["@Val!A2"], hist_tbl, notes["@Val!A101"]]

SEC["recon"]["blocks"][0]["caption"] = "本模型與公司說法、共識、券商的對照"

# 敏感度：第二張龍捲風是 FY2028 EPS（情境 A），不是加權目標價，轉換器之後還原標題
TOR = {(b["sheet"], b["value_col"]): b["title"] for b in SEC["sens"]["blocks"] if b["type"] == "tornado"}
# v0.2：舊折現率三案改為標準版 DCF 參數案例（同一張表），案例列與讀法列位置改變
_sl = next(r for r in range(5, 80) if WB["Sens"].cell(r, 1).value is None) - 1          # 最後一個案例列
_sn = row_of("Sens", "讀法：")

for b in SEC["sens"]["blocks"]:
    if b["type"] == "tornado": b["rows"] = f"5-{_sl}"
    if b["type"] == "table" and b.get("rowgen", {}).get("sheet") == "Sens": b["rowgen"]["rows"] = f"4-{_sl}"
    if b["type"] == "note" and b["parts"][0].startswith("@Sens!A") and b["parts"][0] != "@Sens!A2": b["parts"] = [f"@Sens!A{_sn}"]
pass

SPEC = {
 "in_skip": [in_tok("G."), in_tok("情境驅動對照")],
 "in_groups": [
  ["成熟製程晶圓（產能 × 稼動率 × 平均售價）", [
    ["產能（十二吋約當）", ["cap_26", "capg27", "capg28", "capg29", "cap_adj"]],
    ["平均稼動率（三情境）", [k for k in I if re.match(r"^ut2\d_[ABC]$", k)] + ["ut_adj"]],
    ["平均售價（美元，每片十二吋約當）", ["asp_26"] + [k for k in I if re.match(r"^asp2\d_[ABC]$", k)] + ["asp_adj"]]]],
  ["新業務（AI 相關：先進封裝矽中介層、矽光子；Intel 十二奈米合作）", [
    ["AI 相關營收", ["ai_26", "aig27_A", "aig27_B", "aig27_C", "aig_L", "ai_adj"]],
    ["AI 加速器顆數係數（Siliconomics，FY2028 起乘在 AI 相關營收）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["Intel 十二奈米營收", [k for k in I if re.match(r"^int(_27|2\d_[ABC])$", k)]],
    ["新業務毛利", ["m_new"]]]],
  ["全公司（匯率、成本、FY2026 錨點、股價與目標價規則）", [
    ["匯率", ["fx", "fx_26"]],
    ["成本：折舊、每片現金成本、營業費用", ["da_tot26", "da_sh", "dag_27", "dag_28", "dag_29", "dag_adj", "cc_g", "phi", "g_opx_min"]],
    ["營業外、稅與股數", ["nonop_f", "tax", "sh_g"]],
    ["FY2026 錨點（第四季指引或假設）", ["q4_ship", "q4_asp", "q4_ut", "q4_gm", "q4_opx", "q4_nonop", "q4_tax", "q_sh"]],
    ["股價與市場倍數", ["px", "pe_ttm", "fpe", "cons_feps", "cons_tp", "nim_5"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta_u", "beta_u_med", "beta_de", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["收斂法（FY2030–FY2032）", ["tp_g", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["Wattonomics 錨定（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]]],
 "blocks_to": {"1. 成熟製程晶圓": "成熟製程晶圓（產能 × 稼動率 × 平均售價）", "@Units!A34": "新業務（AI 相關：先進封裝矽中介層、矽光子；Intel 十二奈米合作）"},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "目標價結論": ["目標價（DCF", "DCF 目標價", "Comparable 目標價（相對", "Comparable 目標價（市場", "現價", "上檔空間：加權", "上檔空間：Comparable", "Comparable（相對倍數調整後）的上檔空間", "投資建議", "DCF 情境價值推到", "市場共識"],
          "本模型與公司說法、共識、券商的對照": ["FY2026 營收 vs 共識", "FY2026 EPS vs 共識", "FY2027 平均售價年增", "FY2029 AI 相關營收", "共識前瞻 12 月 EPS"]},
 "n_recon": 4,
}

out = RL.relayout(cfg, X, SPEC)

# 補充規則 A：橫幅中間格＝只看 Comparable（市場主流同業法，未經相對倍數調整）＋自己的上檔空間；附註寫進加權的調整後值
h = out["hero"]
h["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{comp}|n2", ["上檔空間 ", f"@Val!B{upc}|sp1"]]
if abs(WB["Val"].cell(compadj, 2).value - WB["Val"].cell(comp, 2).value) > 1e-9:
    h["note"] = (h.get("note") or []) + ["；進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{compadj}|n2", " 元。"]

for s in out["sections"]:
    for b in s["blocks"]:
        if b["type"] == "tornado":
            t0 = TOR.get((b["sheet"], b["value_col"]))
            if t0 and not t0.startswith("加權目標價"):
                b["title"] = t0 + ("（未達兩位數百分比的案例）" if b.get("collapse") else "（變動達兩位數百分比的案例）")
    if s["id"] == "status":   # 附錄「版本要點」只收版本說明列
        for b in s["blocks"]:
            if b.get("summary") == "展開：本版與前版的修改要點":
                ws = WB["README"]
                b["rows"] = [[f"@README!A{r}"] for r in range(1, ws.max_row + 1) if re.match(r"^v\d+\.\d", str(ws.cell(r, 1).value or ""))]

json.dump(out, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

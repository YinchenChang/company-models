# 鴻海 HTML 設定：讀 foxconn/html_src.json（v0.3.1 的 html.json，內容不動）→ 套用個股 HTML 版面規格 v1.0 → 寫 foxconn/html.json。
# 用法：python3 foxconn/gen_html_json.py <成品 xlsx> foxconn/html_src.json foxconn/html.json
# 規格：docs/plan/個股HTML版面規格.md；共用轉換器 tools/company_html/relayout.py（原樣使用）。
# 本模型沿用廣達的伺服器組裝樣板（標準版 DCF 起 DCF 為 FCFF＋WACC、Comparable 為同業前瞻本益比 × 相對倍數調整），與轉換器預設的估值頁不同，
# 所以 relayout 之後在本檔重組：投資建議橫幅（補充規則 A：中間格為市場主流同業法，調整後值寫在附註）、估值章、敏感度章（龍捲風為 FY2028 EPS，不是目標價）。
import copy
import json
import os
import re
import sys

import openpyxl

X, SRC, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "company_html"))
from relayout import relayout  # noqa: E402

cfg = json.load(open(SRC, encoding="utf-8"))
src = copy.deepcopy(cfg)
wb = openpyxl.load_workbook(X, data_only=True)


def lab(sh, prefix, start=1, col=1):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip().startswith(prefix.strip()):
            return r
    raise KeyError((sh, prefix))


def text(spec):
    if isinstance(spec, list):
        return "".join(text(p) for p in spec)
    m = re.match(r"^@([^!]+)!([A-Z]+\d+)", spec) if isinstance(spec, str) else None
    return str(wb[m.group(1)][m.group(2)].value or "") if m else str(spec or "")


def clean(t, fallback):
    t = re.sub(r"（.*?）|\(.*?\)", "", str(t or "")).strip()
    t = re.sub(r"^[\d.\s]+", "", t).strip()
    return t if t and not re.search(r"\d", re.sub(r"(FY|CY)?(19|20)\d{2}", "", t)) else fallback


INK = {str(wb["In"].cell(r, 1).value): r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}
sec = lambda c, i: next(s for s in c["sections"] if s["id"] == i)

# ---------- relayout 前：驅動 → 推導對照表移到附錄；對帳表給標題、補到最後一列 ----------
for b in sec(cfg, "drivers")["blocks"]:
    if str(b.get("caption", "")).startswith("單一參數（"):
        b["caption"] = "驅動 → 推導對照（主要輸入的基準值與推導量）"
rc = sec(cfg, "recon")["blocks"][0]
rc["caption"] = "本模型與券商估計、櫃數份額的對照"
rc["rowgen"]["rows"] = f"4-{wb['Recon'].max_row}"

# In 節的鍵：G（共識與市場資料，只作對帳）與最後的情境驅動對照不必歸組
_scn = next(re.split(r"[.\s]", str(wb["In"].cell(r, 1).value))[0] for r in range(4, wb["In"].max_row + 1)
            if str(wb["In"].cell(r, 1).value or "").startswith("情境驅動對照"))
SPEC = {
 "in_skip": ["G", _scn],
 "in_groups": [
  ["AI 伺服器組裝（部署量、整櫃價值、鴻海份額、客供與毛利率）", [
    ["部署規模（讀上游網通模型與橋接表）", ["itc_26", "itc_27", "itc_28", "itc_29", "itpg_gb", "itpg_vr", "vr26", "racks_gb", "racks_vr", "rack_gb", "rack_vr", "mf"]],
    ["情境顆數係數（2028 年起三情境不同）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["ODM 整櫃價格與每 GW 價值（讀廣達模型）", ["odm_gb", "odm_vr", "vpg_28", "vpg_29"]],
    ["鴻海有效份額（FY2026 以公司揭露比例校準）", ["cl_25", "ais_25", "cl_26", "ais_26", "rel_27", "rel_28", "rel_29"]],
    ["客供比例與保留毛利", ["cs_27", "cs_28", "cs_29", "keep"]],
    ["AI 伺服器毛利率", ["d_ai", "vrd"]],
    ["Wattonomics 錨定係數（只用於敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  ["其他事業（雲端網路非 AI、消費智能、電腦終端、元件及其他）", [
    ["營收結構錨點（公司揭露比例）", ["ce_25", "pc_25", "ce_26", "pc_26"]],
    ["雲端網路非 AI 成長率", ["g_cn27", "g_cn28", "g_cn29"]],
    ["消費智能成長率", ["g_ce27", "g_ce28", "g_ce29"]],
    ["電腦終端成長率", ["g_pc27", "g_pc28", "g_pc29"]],
    ["元件及其他成長率", ["g_co27", "g_co28", "g_co29"]]]],
  ["全公司（FY2026 錨點、費用與淨利、股價、目標價規則）", [
    ["FY2026 錨點（上半年實績＋第三季營收實績＋假設）", ["q3_rev", "q4g", "q_gm", "q_opx", "q_nio", "q_sh"]],
    ["營業費用、淨利係數與股數（FY2027 起）", ["phi", "opx_adj", "nio_f", "sh_g"]],
    ["匯率", ["fx"]],
    ["股價", ["px", "pe_ttm"]],
    ["目標價規則（情境權重、DCF 權重、延伸期收斂）", ["tp_wA", "tp_wB", "tp_wC", "tp_g", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "nim10_f"]],
    ["市場共識（只作對照）", ["cons_tp", "cons_feps"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]]]]],
 "blocks_to": {"組裝市場": "AI 伺服器組裝（部署量、整櫃價值、鴻海份額、客供與毛利率）", "@Units!A18": "AI 伺服器組裝（部署量、整櫃價值、鴻海份額、客供與毛利率）",
               "標記沿用": "AI 伺服器組裝（部署量、整櫃價值、鴻海份額、客供與毛利率）"},
 "main": {"本模型與券商估計": ["FY2027 EPS vs 法人", "FY2028 EPS vs 法人", "FY2026 有效份額 vs 三家", "12 個月後的前瞻 EPS", "目標價（DCF 與 Comparable 加權）vs"]},
}
cfg = relayout(cfg, X, SPEC)

# ---------- relayout 後（本模型樣板專用） ----------
V = lambda p, start=1: lab("Val", p, start)
r_tpw, r_dcf, r_cmp, r_px = V("目標價（DCF 與 Comparable 加權）"), V("DCF 目標價（情境加權）"), V("Comparable 目標價（相對倍數調整後）"), V("現價（")
r_up, r_upc, r_rec, r_cons = V("上檔空間：加權目標價"), V("上檔空間：Comparable 目標價"), V("投資建議"), V("市場共識目標價")
r_cmp0, r_grade = V("Comparable 目標價（市場主流同業法）"), V("上檔空間級距")
h_upc = lab("README", "上檔空間：Comparable 目標價（相對倍數調整後）")
h_val = lab("README", "估值說明：")

# 投資建議橫幅（補充規則 A）：中間格＝只看 Comparable（市場主流同業法、未經相對倍數調整）＋自己的上檔空間；
# 進加權目標價的是相對倍數調整後的 Comparable，數值寫在橫幅附註（公式帶數字）
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{r_cmp0}|n0", ["上檔空間 ", f"@Val!B{r_upc}|sp1"]]
_note = cfg["hero"].get("note") or []
cfg["hero"]["note"] = (_note if isinstance(_note, list) else [_note]) + ["；進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{r_cmp}|n0", " 元（上檔空間 ", f"@README!B{h_upc}|sp1", "）。"]
cfg["hero"]["tiles"][0][1] = f"@Val!B{r_tpw}|n0"
cfg["hero"]["tiles"][2][1] = f"@Val!B{r_px}|n0"

# 命題章：原有的情境定義與口徑摺疊在最後
th_kv = next(b for b in sec(src, "thesis")["blocks"] if b["type"] == "kv")
sec(cfg, "thesis")["blocks"].append(dict(th_kv, collapse=True, summary="展開：三情境的定義與會計口徑"))

# In 表格格式：pt（百分點）以百分比顯示；權重、每櫃價格、新台幣金額多一位小數
UFMT = {"pt": "p1", "權重": "p0", "US$M/架": "n2", "NT$": "n2"}
for b in sec(cfg, "drivers")["blocks"]:
    if b["type"] == "table" and b.get("head", [None])[0] == "@In!B3":
        for row in b["rows"]:
            if isinstance(row, dict):
                continue
            r = int(re.match(r"@In!B(\d+)", row[0]).group(1))
            f = UFMT.get(str(wb["In"].cell(r, 3).value))
            if f:
                for i in (1, 2, 3):
                    row[i] = re.sub(r"\|.*$", "|" + f, row[i])

# 估值章：說明 → 目標價長條圖 → 結論主要項目 → 折現率 → DCF → Comparable；舊法（十年本益比規則值）全部摺疊
vs = sec(src, "valuation")["blocks"]
va = sec(cfg, "valuation")
va["intro"] = f"@README!A{h_val}"
sec(cfg, "thesis").pop("intro", None)   # 核心問題已在頁首（question），README A4 原文不重複顯示


def vrow(r, f="n0"):
    return [f"@Val!A{r}", f"@Val!B{r}|{f}", f"@Val!E{r}"]


def split(blk, prefixes, summary, extra_main=(), extra_rest=()):
    rows = [r for r in blk["rows"] if not isinstance(r, dict)]
    main, used = [], set()
    for p in prefixes:
        i = next(i for i, r in enumerate(rows) if i not in used and text(r[0]).strip().startswith(p))
        used.add(i)
    main = [rows[i] for i in sorted(used)] + list(extra_main)
    rest = [r for i, r in enumerate(rows) if i not in used] + list(extra_rest)
    base = {k: v for k, v in blk.items() if k not in ("rows", "collapse", "summary")}
    out = [dict(base, rows=main)]
    if rest:
        out.append({"type": "table", "collapse": True, "summary": summary, "head": blk["head"], "rows": rest})
    return out


conc = [vrow(r_tpw), vrow(r_dcf), vrow(r_cmp), vrow(r_px), vrow(r_up, "sp1"),
        [f"@README!A{h_upc}", f"@README!B{h_upc}|sp1", "上檔空間以相對倍數調整後的 Comparable 目標價計算（進加權目標價的值）"],
        [f"@Val!A{r_rec}", f"@Val!B{r_rec}", ""], vrow(r_cons)]
r_hist = V("歷史倍數檢查")
conc_rest = [vrow(r_cmp0), vrow(r_upc, "sp1"), vrow(V("與市場共識目標價的差距：EPS 因素"), "sp1"), vrow(V("與市場共識目標價的差距：倍數因素"), "sp1"),
             vrow(V("價格帶（DCF"), ""), [f"@Val!A{r_hist}", [f"@Val!B{r_hist}|p0", "／", f"@Val!C{r_hist}|p0", "／", f"@Val!D{r_hist}|p0"], f"@Val!E{r_hist}"],
             [f"@Val!A{V('目標價方法')}", f"@Val!B{V('目標價方法')}", ""], [f"@Val!A{r_grade}", f"@Val!B{r_grade}", f"@Val!E{r_grade}"]]
ink = lambda k, f: [f"@In!B{INK[k]}", f"@In!E{INK[k]}|{f}"]
# 標準版 DCF——折現率（緊接在 DCF 彙總之前；規格 v1.0 直接顯示 WACC、ke、稅後 kd、RONIC、永續成長率）
r_d0 = V("0.1 DCF")
vr = lambda p, f: [f"@Val!A{V(p, r_d0)}", f"@Val!B{V(p, r_d0)}|{f}", f"@Val!E{V(p, r_d0)}"]
disc = {"type": "table", "caption": "DCF 折現率（WACC）、永續成長率與權重", "head": ["項目", "數值", "說明"],
        "rows": [vr("股權成本", "p2"), vr("稅後債務成本", "p2"), vr("權益占資本", "p1"), vr("WACC", "p2"), vr("永續成長率", "p2"), vr("永續期 RONIC", "p2"),
                 [f"@In!B{INK['tp_w_dcf']}", f"@In!E{INK['tp_w_dcf']}|p0", ""],
                 ["情境權重（A／B／C）", [f"@In!E{INK['tp_wA']}|p0", "／", f"@In!E{INK['tp_wB']}|p0", "／", f"@In!E{INK['tp_wC']}|p0"], ""]]}
d3 = lambda p, f: [f"@Val!A{V(p, r_d0)}"] + [f"@Val!{c}{V(p, r_d0)}|{f}" for c in "BCD"] + [f"@Val!E{V(p, r_d0)}"]
r_dh = V("項目", r_d0)
dhead = [f"@Val!A{r_dh}", "情境 A（順利）", "情境 B（平平）", "情境 C（不如預期）", "說明"]
dcf_main = {"type": "table", "caption": f"@Val!A{r_d0}", "head": dhead,
            "rows": [d3("明確期 FCFF 現值", "n1"), d3("終值現值", "n1"), d3("企業價值 EV", "n1"), d3("＋淨現金與投資", "n1"), d3("DCF 情境價值推到 12 個月後", "n0"),
                     [f"@Val!A{r_dcf}", f"@Val!B{r_dcf}|n0", "", "", f"@Val!E{r_dcf}"]]}
dcf_rest = {"type": "table", "collapse": True, "summary": "展開：DCF 的其他項目（股數、評價日每股價值、終值占比、隱含終值倍數）", "head": dhead,
            "rows": [d3("評價用稀釋股數", "n3"), d3("DCF 情境每股價值（評價日）", "n1"), d3("  其中：終值現值占 EV", "p1"), d3("  隱含終值 EV／EBITDA", "x1")]}
dcf_blk = next(b for b in vs if b.get("caption") == f"@Val!A{V('0.1 DCF')}")
cmp_blk = next(b for b in vs if b.get("caption") == f"@Val!A{V('0.2 Comparable')}")
rm_kv = next(b for b in vs if b["type"] == "kv")
kv_rows = [[k, v, ""] + [""] * (len(cmp_blk["head"]) - 3) for k, v in rm_kv["items"]]
nb = [{"type": "bars", "title": "目標價與現價", "fmt": "n0", "items": [
        {"label": f"@Val!A{r_dcf}", "value": f"@Val!B{r_dcf}", "series": 0},
        {"label": f"@Val!A{r_cmp}", "value": f"@Val!B{r_cmp}", "series": 1},
        {"label": f"@Val!A{r_tpw}", "value": f"@Val!B{r_tpw}", "series": 2},
        {"label": f"@Val!A{r_cons}", "value": f"@Val!B{r_cons}", "series": 0}],
       "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}},
      {"type": "table", "caption": "結論（主要項目）", "head": ["項目", "數值", "說明"], "rows": conc},
      {"type": "table", "collapse": True, "summary": "展開：結論的其他項目（市場主流同業法、與共識的差距拆解、價格帶、歷史倍數檢查、級距）", "head": ["項目", "數值", "說明"], "rows": conc_rest},
      disc]
nb += [dcf_main, dcf_rest]
cm = lambda p: next(r for r in cmp_blk["rows"] if text(r[0]).strip().startswith(p))[:2]
kv2 = [[k, v] for k, v in rm_kv["items"]]
kv_main = [r for r in kv2 if text(r[0]).startswith(("相對倍數調整係數", "Comparable 目標價（相對倍數調整後）"))]
nb += [{"type": "table", "caption": cmp_blk["caption"], "head": ["項目", "數值"],
        "rows": [cm("同業前瞻 P/E 中位數"), cm("12 個月後的前瞻 EPS"), cm("Comparable 目標價")] + kv_main},
       {"type": "table", "collapse": True, "summary": "展開：同業前瞻本益比", "head": cmp_blk["head"],
        "rows": [r for r in cmp_blk["rows"] if not text(r[0]).strip().startswith(("同業前瞻 P/E 中位數", "12 個月後的前瞻 EPS", "Comparable 目標價"))]},
       {"type": "table", "collapse": True, "summary": "展開：相對倍數調整的組成（本公司與同業的長期本益比）", "head": ["項目", "數值"],
        "rows": [r for r in kv2 if r not in kv_main]}]
nb.append(next(b for b in vs if b["type"] == "note" and b["parts"] == [f"@Val!E{V('同業前瞻 P/E 中位數')}"]))
nb.append({"type": "h3", "text": "參考：舊法（十年本益比規則值）與市場隱含倍數（不進目標價）"})
nb.append({"type": "note", "parts": ["@Val!A1"], "collapse": True, "summary": "展開：本頁原說明（十年本益比規則值）"})
skip = {id(dcf_blk), id(cmp_blk), id(rm_kv)}
title = "舊法"
last_cap = "舊法"
NOTE_SUM = {"@Val!A2": "循環位置的判定方式", "@Val!A33": "What must be true 表的讀法", "@Val!A35": "先驗對照", "@Val!A51": "十年倍數規則的說明"}
for b in vs[vs.index(cmp_blk) + 1:]:
    if id(b) in skip or (b["type"] == "note" and b["parts"][0].startswith(("@Val!E", "@Val!A135", "@Val!A136"))):
        continue
    if b["type"] == "h3":
        title = clean(b["text"], title)
        continue
    if b["type"] == "list":   # 原重點觀察已移到第一章（白話版）
        continue
    if b.get("collapse"):
        nb.append(b)
        continue
    if b["type"] == "note":   # 依說明所在儲存格給摘要；未列者掛在前一個表格名下
        cap = NOTE_SUM.get(b["parts"][0], f"{last_cap}的說明")
    else:
        cap = clean(text(b["caption"]), title) if b.get("caption") else clean(b.get("title", ""), title) if b["type"] == "bars" else title
        last_cap = cap
    nb.append(dict(b, collapse=True, summary=f"展開：{cap}"))
va["blocks"] = nb

# 敏感度章：龍捲風是 FY2028 EPS（情境 A），標題改回 EPS；目標價參數表改指向現行列號；另列 Wattonomics 錨定案例
se = sec(cfg, "sens")
tor_src = next(b for b in sec(src, "sens")["blocks"] if b["type"] == "tornado")
for b in se["blocks"]:
    if b["type"] == "tornado":
        b["title"] = (tor_src["title"].replace("（", "（變動達兩位數百分比的案例；", 1) if not b.get("collapse")
                      else tor_src["title"].split("（")[0] + "（未達兩位數百分比）")
r_tp = lab("Sens", "目標價參數")
r_end = r_tp + 2
while wb["Sens"].cell(r_end + 1, 1).value:
    r_end += 1
for b in se["blocks"]:
    if b.get("caption") == f"@Sens!A{r_tp}":
        b["head"] = [f"@Sens!{c}{r_tp + 1}" for c in "ABCDEFG"]
        b["rowgen"] = {"sheet": "Sens", "rows": f"{r_tp + 2}-{r_end}", "cols": list("ABCDEFG"), "fmts": ["", "", "n0", "sp1", "n0", "sp1", "n0"]}
r_w = lab("Sens", "Wattonomics 錨定（三層校準原則）")
r_wh = lab("Sens", "案例", r_w)
r_we = r_wh
while wb["Sens"].cell(r_we + 1, 1).value:
    r_we += 1
full = next(b for b in sec(src, "sens")["blocks"] if b["type"] == "table" and b.get("collapse"))
se["blocks"] += [{"type": "note", "parts": [f"@Sens!A{r_w + 1}"]},
                 {"type": "table", "collapse": True, "summary": "展開：Wattonomics 錨定案例（情境 B 乘上錨定係數；不重算目標價）",
                  "head": [f"@Sens!{c}{r_wh}" for c in full["rowgen"]["cols"]],
                  "rowgen": dict(full["rowgen"], rows=f"{r_wh + 1}-{r_we}")}]

json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")

# 緯創 HTML 設定：讀 wistron/html_src.json（v0.3 的 html.json，內容不動）→ 套用個股 HTML 版面規格 v1.0 → 寫 wistron/html.json。
# 用法：python3 wistron/gen_html_json.py <成品 xlsx> wistron/html_src.json wistron/html.json
# 規格：docs/plan/個股HTML版面規格.md；共用轉換器 tools/company_html/relayout.py（原樣使用）。
# 本模型屬伺服器組裝樣板（v0.4 起 DCF 為標準版 FCFF＋WACC、Comparable 為同業前瞻本益比 × 相對倍數調整），與轉換器預設的估值頁不同，
# 所以 relayout 之後在本檔重組：投資建議橫幅（補充規則 A：中間格為市場主流同業法，調整後值寫在附註）、估值章、敏感度章（龍捲風為 EPS，不是目標價）。
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


INK = {str(wb["In"].cell(r, 1).value): r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}
sec = lambda c, i: next(s for s in c["sections"] if s["id"] == i)

# 情境驅動表：依 In 頁「情境驅動對照」節的實際列號重寫參照（v0.4 起 In 頁新增折現率一節，html_src 的固定列號錯位；欄序同 html_src：B、D、E、F、C）
_scn_r0 = next(r for r in range(4, wb["In"].max_row + 1) if str(wb["In"].cell(r, 1).value or "").startswith("情境驅動對照")) + 1
_scn_rs, _r = [], _scn_r0 + 1
while wb["In"].cell(_r, 1).value:
    _scn_rs.append(_r); _r += 1
_n = 0
for b in sec(cfg, "drivers")["blocks"]:
    if str(b.get("caption", "")).startswith("情境驅動"):
        b["head"] = [f"@In!{c}{_scn_r0}" for c in "BDEFC"]
        b["rows"] = [[f"@In!{c}{r}" for c in "BDEFC"] for r in _scn_rs]
        _n += 1
assert _n == 1, "情境驅動表"

# ---------- relayout 前：驅動 → 推導對照表移到附錄；對帳表給標題 ----------
for b in sec(cfg, "drivers")["blocks"]:
    if str(b.get("caption", "")).startswith("單一參數（"):
        b["caption"] = "驅動 → 推導對照（主要輸入的基準值與推導量）"
rc = sec(cfg, "recon")["blocks"][0]
rc["caption"] = "本模型與共識、券商估計及櫃數調查的對照"

for b in sec(cfg, "drivers")["blocks"]:   # Q 頁 FY2026 錨點表：給標題，歸到全公司塊
    if b["type"] == "table" and b.get("rowgen", {}).get("sheet") == "Q":
        b["caption"] = "FY2026 錨點：緯創本體與緯穎（Q 頁）"
_scn = next(re.split(r"[.\s]", str(wb["In"].cell(r, 1).value))[0] for r in range(4, wb["In"].max_row + 1)
            if str(wb["In"].cell(r, 1).value or "").startswith("情境驅動對照"))
AI = "AI 伺服器（本體 AI 系統、GPU 基板與緯穎：部署量、整櫃價值、份額、客供與毛利率）"
GS = "非 AI 營收（一般伺服器與交換器、非伺服器、緯穎一般伺服器）"
CO = "全公司（FY2026 錨點、費用與淨利、股價、目標價規則、折現率與自由現金流）"
SPEC = {
 "in_skip": ["H", _scn],
 "in_groups": [
  [AI, [
    ["部署規模與 GPU 顆數（讀上游網通、橋接表與台積電模型）", ["itc_26", "itc_27", "itc_28", "itc_29", "itpg_gb", "itpg_vr", "vr26", "racks_gb", "racks_vr", "gu_26", "gu_27", "gu_28", "gu_29", "mf"]],
    ["情境顆數係數（2028 年起三情境不同）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["ODM 整櫃價格與每 GW 價值（讀廣達模型）", ["odm_gb", "odm_vr", "vpg_28", "vpg_29"]],
    ["本體 AI 系統份額與客供", ["rel_27", "rel_28", "rel_29", "cs_27", "cs_28", "cs_29", "keep"]],
    ["緯穎份額與客供（讀 Siliconomics 與緯穎模型）", ["wrel_27", "wrel_28", "wrel_29", "wcs_27", "wcs_28", "wcs_29", "wkeep", "wai_26"]],
    ["GPU 基板與運算板", ["brd_26", "bsh_26", "bv_27", "bv_28", "bv_29", "brel_27", "brel_28", "brel_29"]],
    ["毛利率與緯穎利潤率", ["d_ai", "d_brd", "d_gen", "vrd", "wgm", "wopm", "wnio"]],
    ["Wattonomics 錨定係數（只用於敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  [GS, [
    ["營收結構錨點與成長率", ["nsv_26", "gen_26", "g_gen27", "g_gen28", "g_gen29", "g_pc27", "g_pc28", "g_pc29", "g_wn27", "g_wn28", "g_wn29"]]]],
  [CO, [
    ["FY2026 錨點（上半年實績＋第三季營收實績＋假設）", ["fx", "q3_rev", "q4g", "q3_wrev", "q4gw", "q_wgm", "q_wopm", "q_wnio", "q_sgm", "q_sopx", "q_snio", "stake", "stake25", "q_sh3", "q_sh4"]],
    ["營業費用、淨利係數與股數（FY2027 起）", ["phi", "opx_adj", "snio", "sh_g"]],
    ["股價、市場倍數與共識", ["px", "pe_ttm", "fpe", "cons_feps", "cons_tp", "nim_5y"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_wA", "tp_wB", "tp_wC", "tp_w_dcf"]],
    ["收斂法（FY2030 起）", ["tp_g", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用；非控制權益以緯穎少數股權市值計）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "wy_px", "wy_sh", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]]]]],
 "blocks_to": {"組裝市場": AI, "@Units!A19": AI, "FY2026 錨點": CO, "@Q!A2": CO},
 "main": {"本模型與共識": ["FY2026 EPS vs 共識", "FY2027 EPS vs 法人下限說法", "FY2027 EPS vs 法人樂觀說法中點", "FY2027 緯穎淨利 vs 凱基", "FY2027 緯穎對歸屬母公司淨利"]},
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
cfg["hero"]["tiles"][0][1] = f"@Val!B{r_tpw}|n0"
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{r_cmp0}|n0", ["上檔空間 ", f"@Val!B{r_upc}|sp1"]]
cfg["hero"]["tiles"][2] = [f"@Val!A{r_px}", f"@Val!B{r_px}|n0", ""]
_note = cfg["hero"].get("note") or []
cfg["hero"]["note"] = (_note if isinstance(_note, list) else [_note]) + ["；進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{r_cmp}|n0", " 元（上檔空間 ", f"@README!B{h_upc}|sp1", "）。"]

# 命題章：原有的情境定義與口徑摺疊在最後
th_kv = next(b for b in sec(src, "thesis")["blocks"] if b["type"] == "kv")
sec(cfg, "thesis")["blocks"].append(dict(th_kv, collapse=True, summary="展開：三情境的定義與會計口徑"))

# 估值章：說明 → 目標價長條圖 → 結論主要項目 → 折現率 → DCF → Comparable；舊法（十年本益比規則值）全部摺疊
vs = sec(src, "valuation")["blocks"]
va = sec(cfg, "valuation")
va["intro"] = f"@README!A{h_val}"


def vrow(r, f="n0"):
    return [f"@Val!A{r}", f"@Val!B{r}|{f}", f"@Val!E{r}"]


def vrow3(r, f="n1"):
    return [f"@Val!A{r}", f"@Val!B{r}|{f}", f"@Val!C{r}|{f}", f"@Val!D{r}|{f}", f"@Val!E{r}"]


ink = lambda k, f: [f"@In!B{INK[k]}", f"@In!E{INK[k]}|{f}"]
r_hist, r_meth, r_band = V("歷史倍數檢查"), V("目標價方法"), V("價格帶（DCF")
conc = [vrow(r_tpw), vrow(r_dcf), vrow(r_cmp), vrow(r_px), vrow(r_up, "sp1"),
        [f"@README!A{h_upc}", f"@README!B{h_upc}|sp1", "上檔空間以相對倍數調整後的 Comparable 目標價計算（進加權目標價的值）"],
        [f"@Val!A{r_rec}", f"@Val!B{r_rec}", ""], vrow(r_cons)]
conc_rest = [vrow(r_cmp0), vrow(r_upc, "sp1"), vrow(V("與市場共識目標價的差距：EPS 因素"), "sp1"), vrow(V("與市場共識目標價的差距：倍數因素"), "sp1"),
             vrow(r_band, ""), [f"@Val!A{r_hist}", [f"@Val!B{r_hist}|p0", "／", f"@Val!C{r_hist}|p0", "／", f"@Val!D{r_hist}|p0"], f"@Val!E{r_hist}"],
             [f"@Val!A{r_meth}", f"@Val!B{r_meth}", ""], [f"@Val!A{r_grade}", f"@Val!B{r_grade}", f"@Val!E{r_grade}"]]
# v0.2：標準版 DCF——折現率（緊接在 DCF 彙總之前；規格 v1.0 直接顯示 WACC、ke、稅後 kd、RONIC、永續成長率）
r_d0 = V("0.1 DCF")
vr = lambda p, f: [f"@Val!A{V(p, r_d0)}", f"@Val!B{V(p, r_d0)}|{f}", f"@Val!E{V(p, r_d0)}"]
disc = {"type": "table", "caption": "DCF 折現率（WACC）、永續成長率與權重", "head": ["項目", "數值", "說明"],
        "rows": [vr("股權成本", "p2"), vr("稅後債務成本", "p2"), vr("權益占資本", "p1"), vr("WACC", "p2"), vr("永續成長率", "p2"), vr("永續期 RONIC", "p2"),
                 [f"@In!B{INK['tp_w_dcf']}", f"@In!E{INK['tp_w_dcf']}|p0", ""],
                 ["情境權重（A／B／C）", [f"@In!E{INK['tp_wA']}|p0", "／", f"@In!E{INK['tp_wB']}|p0", "／", f"@In!E{INK['tp_wC']}|p0"], ""]]}
d3 = lambda p, f: [f"@Val!A{V(p, r_d0)}"] + [f"@Val!{c}{V(p, r_d0)}|{f}" for c in "BCD"] + [f"@Val!E{V(p, r_d0)}"]
dhead = [f"@Val!A{r_d0 + 1}", "情境 A（順利）", "情境 B（平平）", "情境 C（不如預期）", "說明"]
dcf_main = [d3("明確期 FCFF 現值", "n1"), d3("終值現值", "n1"), d3("企業價值 EV", "n1"), d3("＋淨現金與投資", "n1"), d3("DCF 情境價值推到 12 個月後", "n0"),
            [f"@Val!A{r_dcf}", f"@Val!B{r_dcf}|n0", "", "", f"@Val!E{r_dcf}"]]
dcf_rest = [d3("評價用稀釋股數", "n4"), d3("DCF 情境每股價值（評價日）", "n1"), d3("  其中：終值現值占 EV", "p1"), d3("  隱含終值 EV／EBITDA", "x1")]
r_c0, r_med, r_feps, r_c = V("0.2 Comparable"), V("同業前瞻 P/E 中位數"), V("12 個月後的前瞻 EPS"), V("Comparable 目標價", V("12 個月後的前瞻 EPS"))
r_coef = V("相對倍數調整係數")
cmp_src = next(b for b in vs if b.get("caption") == f"@Val!A{r_c0}")
rm_kv = next(b for b in vs if b["type"] == "kv" and b["items"][0][0] == f"@Val!A{V('本公司 10 年中位數')}")
nb = [{"type": "bars", "title": "目標價與現價", "fmt": "n0", "items": [
        {"label": f"@Val!A{r_dcf}", "value": f"@Val!B{r_dcf}", "series": 0},
        {"label": f"@Val!A{r_cmp}", "value": f"@Val!B{r_cmp}", "series": 1},
        {"label": f"@Val!A{r_tpw}", "value": f"@Val!B{r_tpw}", "series": 2},
        {"label": f"@Val!A{r_cons}", "value": f"@Val!B{r_cons}", "series": 0}],
       "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}},
      {"type": "table", "caption": "結論（主要項目）", "head": ["項目", "數值", "說明"], "rows": conc},
      {"type": "table", "collapse": True, "summary": "展開：結論的其他項目（市場主流同業法、與共識的差距拆解、價格帶、歷史倍數檢查、目標價方法、級距）", "head": ["項目", "數值", "說明"], "rows": conc_rest},
      disc,
      {"type": "table", "caption": f"@Val!A{r_d0}", "head": dhead, "rows": dcf_main},
      {"type": "table", "collapse": True, "summary": "展開：DCF 的其他項目（股數、評價日每股價值、終值占比、隱含終值倍數）", "head": dhead, "rows": dcf_rest},
      {"type": "table", "caption": f"@Val!A{r_c0}", "head": ["項目", "數值"],
       "rows": [[f"@Val!A{r_med}", f"@Val!B{r_med}|x1"], [f"@Val!A{r_feps}", f"@Val!B{r_feps}|n1"], [f"@Val!A{r_c}", f"@Val!B{r_c}|n0"],
                [f"@Val!A{r_coef}", f"@Val!B{r_coef}|n2"], [f"@Val!A{r_cmp}", f"@Val!B{r_cmp}|n0"]]},
      dict({k: v for k, v in cmp_src.items() if k != "caption"}, collapse=True, summary="展開：同業前瞻本益比"),
      {"type": "table", "collapse": True, "summary": "展開：相對倍數調整的組成（本公司與同業的長期本益比）", "head": ["項目", "數值"],
       "rows": [[k, v] for k, v in rm_kv["items"] if not v.startswith((f"@Val!B{r_coef}", f"@Val!B{r_cmp}"))]},
      {"type": "note", "parts": [f"@Val!E{r_med}"]},
      {"type": "h3", "text": "參考：舊法（十年本益比規則值）與市場隱含倍數（不進目標價）"},
      {"type": "note", "parts": ["@Val!A1"], "collapse": True, "summary": "展開：本頁原說明"}]
# 舊法各塊：依原順序全部摺疊
old_kv = next(b for b in vs if b["type"] == "kv" and b["items"][0][0] == f"@Val!A{r_px}")
nb.append({"type": "table", "collapse": True, "summary": "展開：現價與市場隱含倍數", "head": ["項目", "數值"],
           "rows": [[k, v] for k, v in old_kv["items"]]})
i10 = next(i for i, b in enumerate(vs) if b["type"] == "h3" and str(b.get("text", "")).startswith("十年"))
NOTE_SUM = {"@Val!A2": "循環位置的判定方式", "@Val!A82": "What must be true 表的讀法"}
for b in vs[i10 + 1:]:
    if b.get("collapse"):
        nb.append(b)
    elif b["type"] == "note":
        nb.append(dict(b, collapse=True, summary="展開：" + NOTE_SUM.get(b["parts"][0], "舊法的補充說明")))
    elif b["type"] == "table":
        hd = str(wb["Val"][b["head"][0][5:]].value or "") if b["head"] else ""
        nm = "規則值隱含股價" if hd.startswith("情境") and b["rowgen"]["cols"][1] == "B" and "E" in b["rowgen"]["cols"] else "循環位置"
        nb.append(dict(b, collapse=True, summary=f"展開：{nm}"))
va["blocks"] = nb

# 敏感度章：龍捲風是 EPS（情境 A），標題改回原標題
se = sec(cfg, "sens")
tor_src = {b["value_col"]: b["title"] for b in sec(src, "sens")["blocks"] if b["type"] == "tornado"}
for b in se["blocks"]:
    if b["type"] == "tornado":
        t = tor_src[b["value_col"]]
        base, _, rest = t.partition("（")
        if b.get("collapse"):
            b["title"] = f"{base}（未達兩位數百分比）"
        else:
            b["title"] = f"{base}（變動達兩位數百分比的案例" + (f"；{rest}" if rest else "）")
# v0.2：目標價參數表（標準版 DCF 案例；表格移到案例區塊之後）
r_tp = lab("Sens", "目標價參數敏感度")
r_end = r_tp + 2
while wb["Sens"].cell(r_end + 1, 1).value:
    r_end += 1
for b in se["blocks"]:
    if b.get("caption") == "@Sens!A29":   # 原目標價參數表（v0.3 第 29 列）
        b["caption"] = f"@Sens!A{r_tp}"
        b["head"] = [f"@Sens!{c}{r_tp + 1}" for c in "ABCDEFG"]
        b["rowgen"] = {"sheet": "Sens", "rows": f"{r_tp + 2}-{r_end}", "cols": list("ABCDEFG"), "fmts": ["", "", "n0", "sp1", "n0", "sp1", "n0"]}
# Wattonomics 錨定案例（Sens 頁最後一節；原 HTML 未列）
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

# 台達電 html.json 產生器（v0.4.1：個股 HTML 版面規格 v1.0，docs/plan/個股HTML版面規格.md）
# 讀 delta/html_src.json（原版面設定，內容不動）→ 依列標籤重定位與整理 → relayout() → 寫 delta/html.json（勿手改 html.json）。
# 用法：python3 delta/gen_html_json.py [成品 xlsx]（省略時用 delta/dist/ 內唯一的 xlsx）
import copy, glob, json, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
from relayout import relayout  # noqa: E402

X = sys.argv[1] if len(sys.argv) > 1 else (lambda fs: fs[0] if len(fs) == 1 else sys.exit(f"dist 內應剛好一個 xlsx：{fs}"))(sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx"))))
wb = openpyxl.load_workbook(X, data_only=True)
cfg = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))


def row(sh, label, start=1, exact=False):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or "").strip()
        if (v == label.strip()) if exact else v.startswith(label.strip()):
            return r
    raise KeyError((sh, label))


def until_blank(sh, r0):
    ws, r = wb[sh], r0
    while ws.cell(r, 1).value not in (None, ""):
        r += 1
    return r - 1


sec = {s["id"]: s for s in cfg["sections"]}
IW = wb["In"]
INK = {str(IW.cell(r, 1).value): r for r in range(4, IW.max_row + 1) if IW.cell(r, 1).value}
SCN0 = row("In", "情境驅動對照")
SCN_TOKEN = re.split(r"[.\s]", str(IW.cell(SCN0, 1).value))[0]

# ── 驅動因子：既有「單一參數」對照表改為附錄的驅動 → 推導對照表；情境驅動表依 In 頁現行列號重建 ──
dr = sec["drivers"]["blocks"]
for b in dr:
    cap = str(b.get("caption", ""))
    if cap.startswith("單一參數"):
        b["caption"] = "驅動 → 推導對照表"
    elif cap.startswith("情境驅動"):
        h = SCN0 + 1
        b["head"] = [f"@In!{c}{h}" for c in "BDEFCG"]
        b["rows"] = [[f"@In!B{r}", f"@In!D{r}|{f}", f"@In!E{r}|{f}", f"@In!F{r}|{f}", f"@In!C{r}", f"@In!G{r}|tag"]
                     for r in range(h + 1, until_blank("In", h + 1) + 1)
                     for f in [{"%": "p0", "x": "n2", "權重": "p0"}.get(str(IW.cell(r, 3).value), "n2")]]

# ── 估值：目標價與現價長條圖、結論表（主要列）、DCF 折現率、相對倍數調整表；舊法與規則值摺疊 ──
va = sec["valuation"]["blocks"]
V = lambda lab, c="B", f="": f"@Val!{c}{row('Val', lab)}" + (f"|{f}" if f else "")
VA = lambda lab: f"@Val!A{row('Val', lab)}"
conc_labels = [("目標價（DCF 與 Comparable 加權）", "n0"), ("DCF 目標價（情境加權）", "n0"), ("Comparable 目標價（相對倍數調整後）", "n0"),
               ("Comparable 目標價（市場主流同業法）", "n0"), ("現價（", "n0"), ("上檔空間：加權目標價", "sp1"), ("上檔空間：Comparable 目標價", "sp1"),
               ("上檔空間級距", ""), ("投資建議", ""), ("市場共識目標價", "n0"), ("與市場共識目標價的差距：EPS 因素", "sp1"),
               ("與市場共識目標價的差距：倍數因素", "sp1"), ("價格帶（DCF 不如預期～順利）", ""), ("目標價方法", "")]
conc = {"type": "table", "caption": "目標價、現價與投資建議", "head": ["項目", "數值"], "rows": [[VA(l), V(l, f=f)] for l, f in conc_labels]}
hist = next(b for b in va if b["type"] == "table" and b.get("caption") == "目標價（DCF 與 Comparable 加權）")   # 原結論表：由上方 conc 取代
hr_ = row("Val", "歷史倍數檢查")
bars = {"type": "bars", "title": "目標價與現價", "fmt": "n0", "items": [
    {"label": VA("DCF 目標價（情境加權）"), "value": V("DCF 目標價（情境加權）"), "series": 0},
    {"label": VA("Comparable 目標價（相對倍數調整後）"), "value": V("Comparable 目標價（相對倍數調整後）"), "series": 1},
    {"label": VA("目標價（DCF 與 Comparable 加權）"), "value": V("目標價（DCF 與 Comparable 加權）"), "series": 2},
    {"label": VA("市場共識目標價"), "value": V("市場共識目標價"), "series": 0}],
    "ref": {"label": VA("現價（"), "value": V("現價（")}}
rate = {"type": "table", "caption": "DCF 折現率與永續成長（In 頁；直接進 DCF）", "head": ["@In!B3", "@In!E3", "@In!C3", "@In!G3"],
        "rows": [[f"@In!B{INK[k]}", f"@In!E{INK[k]}|{f}", f"@In!C{INK[k]}", f"@In!G{INK[k]}|tag"] for k, f in (("tp_r", "p1"), ("tp_g", "p1"), ("tp_roe", "p1"))]}
nb = [bars, conc,
      {"type": "table", "collapse": True, "summary": "展開：歷史倍數檢查（現價隱含市盈率的十年分位；只作檢查）", "head": ["項目", "情境 A", "情境 B", "情境 C"],
       "rows": [[f"@Val!A{hr_}", f"@Val!B{hr_}|p0", f"@Val!C{hr_}|p0", f"@Val!D{hr_}|p0"]]}]
old_ref = []   # 舊法與規則值（不進目標價）：摺疊
mode, pending_h3, pending_old = "main", None, None
for b in va:
    cap = str(b.get("caption", ""))
    if b is hist or (b["type"] == "note" and b["parts"][0] in ("@Val!A134", "@Val!A135")):
        continue
    if b["type"] == "h3" and b["text"] == "重點觀察":
        continue
    if b["type"] == "list" and b.get("items", [""])[0] == "@README!A20":
        continue
    if b["type"] == "bars":
        continue   # 舊法規則值長條圖：已由上方目標價長條圖取代，數字仍見下方規則值表
    if b["type"] == "h3" and b["text"].startswith("相對倍數調整"):
        pending_h3 = b["text"]
        continue
    if b["type"] == "kv" and pending_h3:
        rel_tbl = {"type": "table", "caption": pending_h3, "head": ["項目", "數值"], "rows": [[k, v] for k, v in b["items"]]}
        pending_h3 = None
        continue
    if cap == "@Val!A86":
        nb.append(rate)
    if cap == "舊法：規則值情境加權（參考）":
        mode = "old"
        nb.append(rel_tbl)
        nb.append({"type": "h3", "text": "參考：舊法、規則值與市場隱含倍數（不進目標價）"})
    if mode == "old":
        if b["type"] == "h3":
            pending_old = b["text"]
            continue
        if b["type"] == "table" and not b.get("collapse"):
            b = dict(b, collapse=True, summary="展開：" + (pending_old or cap))
            if not b.get("caption"):
                b.pop("caption", None)
            pending_old = None
        nb.append(b)
        continue
    nb.append(b)
sec["valuation"]["blocks"] = nb
sec["valuation"]["intro"] = f"@README!A{row('README', '估值：')}"   # 原 Val!A1 只描述舊法（規則值），改用 README 的目標價方法說明
sec["thesis"].pop("intro", None)   # 與頂層核心問題重複

# ── 對帳：加說明標題，主要列由 SPEC main 指定 ──
for b in sec["recon"]["blocks"]:
    if b["type"] == "table" and "rowgen" in b and not b.get("caption"):
        b["caption"] = "本模型與券商估計、研調市場、份額的對照"

# ── 敏感度：目標價參數表依 Sens 頁現行列號重建 ──
for b in sec["sens"]["blocks"]:
    if b["type"] == "table" and str(b.get("caption", "")).startswith("@Sens!A"):
        r0 = row("Sens", "目標價參數")
        b["caption"] = f"@Sens!A{r0}"
        b["head"] = [f"@Sens!{c}{r0 + 1}" for c in "ABCDEF"]
        b["rowgen"] = {"sheet": "Sens", "rows": f"{r0 + 2}-{until_blank('Sens', r0 + 2)}", "cols": ["A", "B", "C", "D", "E", "F"], "fmts": ["", "", "n0", "n0", "n0", "sp1"]}

old_thesis_kv = copy.deepcopy(next(b for b in sec["thesis"]["blocks"] if b["type"] == "kv"))

SPEC = {
 "in_skip": ["G", SCN_TOKEN],
 "in_groups": [
  ["AI 電源（電源及零組件事業群）", [
    ["AI 部署規模（讀網通市場模型、橋接表與台積電顆數係數）", ["itc_26", "itc_27", "itc_28", "itc_29", "itpg_gb", "itpg_vr", "vr26", "racks_vr", "um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C", "mf"]],
    ["每 kW 電源內容與台達電份額", ["pkw_26", "pkw_27", "pkw_28", "pkw_29", "rel_27", "rel_28", "rel_29", "ai_25", "ai_26"]],
    ["AI 電源毛利率", ["d_ai"]]]],
  ["液冷（基礎設施事業群）", [
    ["液冷每 GW 營收與校準", ["cool_vr", "lcr_27", "lcr_28", "lcr_29", "lc_25", "lc_26"]]]],
  ["非 AI 事業（電源零組件非 AI、基礎設施非液冷、自動化、交通）", [
    ["事業群比重（年報與月營收公告）", ["s25_pe", "s26_pe", "s25_mo", "s26_mo", "s25_au", "s26_au", "s25_inf", "s26_inf"]],
    ["非 AI 分部成長率", ["g_pe27", "g_pe28", "g_pe29", "g_inf27", "g_inf28", "g_inf29", "g_au27", "g_au28", "g_au29", "g_mo27", "g_mo28", "g_mo29"]]]],
  ["全公司：錨點、費用、股價、折現率與目標價規則", [
    ["FY2026 錨點（第三季實績與第四季假設）", ["q3_rev", "q3_gp", "q3_op", "q3_pti", "q3_ni", "q3_eps", "q4g", "q_gm", "q_opx", "q_nop", "q_nir", "q_sh"]],
    ["營業費用、業外、稅與股數（FY2027 起）", ["phi", "opx_adj", "nop_f", "nir_f", "sh_g"]],
    ["匯率", ["fx"]],
    ["股價與目標價規則", ["px", "pe_ttm", "tp_wA", "tp_wB", "tp_wC", "tp_w_dcf", "cons_tp", "cons_feps"]],
    ["折現率、永續成長與收斂法", ["tp_r", "tp_g", "tp_roe", "fade_gA", "fade_gB", "fade_gC", "nim10_f"]],
    ["自由現金流參數", ["da_r", "cx_r"]],
    ["Wattonomics 錨定係數（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]]],
 "blocks_to": {"AI 電源與液冷": "AI 電源（電源及零組件事業群）", "@Units!A15": "AI 電源（電源及零組件事業群）", "標記沿用": "AI 電源（電源及零組件事業群）"},
 "main": {"目標價、現價與投資建議": ["目標價（DCF 與 Comparable 加權）", "DCF 目標價（情境加權）", "Comparable 目標價（相對倍數調整後）", "現價（", "上檔空間：加權目標價", "投資建議", "市場共識目標價"],
          "0.1 DCF": ["DCF 情境每股價值", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "0.2 Comparable": ["同業前瞻 P/E 中位數", "12 個月後的前瞻 EPS", "Comparable 目標價"],
          "相對倍數調整": ["同業 10 年中位數的中位數", "相對倍數調整係數", "Comparable 目標價（相對倍數調整後）"],
          "本模型與券商估計": ["FY2026 EPS vs 本土券商", "FY2027 EPS vs 投顧研究團隊", "CY2027 AI 電源市場（US$B）vs UBS", "FY2026 台達電有效份額 vs 凱基"]},
}
cfg = relayout(cfg, X, SPEC)

# ── 轉換後調整 ──
# 1. 橫幅第二格：加權目標價實際使用的 Comparable（相對倍數調整後）；上檔空間取 README 頁 HTML 橫幅用列
hb = row("README", "HTML 橫幅用：Comparable 目標價（相對倍數調整後）")
cfg["hero"]["tiles"][1] = [VA("Comparable 目標價（相對倍數調整後）"), V("Comparable 目標價（相對倍數調整後）", f="n2"), ["上檔空間 ", f"@README!B{hb}|sp1"]]
# 2. 敏感度龍捲風圖：本模型的龍捲風是情境 A 的 FY2028 EPS（不是加權目標價），標題改回正確口徑
for b in sec["sens"]["blocks"]:
    if b["type"] == "tornado":
        b["title"] = "情境 A 的 FY2028 EPS 相對基準的變動（一次動一個輸入）" + ("：變動未達兩位數百分比的案例" if b.get("collapse") else "：變動達兩位數百分比的案例")
# 3. In 頁單位「pt」（比例差）與「NT$」（每股金額）的顯示格式：轉換器預設一位小數，改為百分比與兩位小數
for b in sec["drivers"]["blocks"]:
    for r in b.get("rows", []):
        if isinstance(r, list) and isinstance(r[0], str) and r[0].startswith("@In!B"):
            u = str(IW.cell(int(r[0][5:]), 3).value)
            to = {"pt": "|p1", "NT$": "|n2"}.get(u)
            if to:
                r[1:4] = [c.replace("|n1", to) if isinstance(c, str) else c for c in r[1:4]]
# 4. 第一章補情境定義與口徑（原第一章的對照表）
sec["thesis"]["blocks"] += [{"type": "h3", "text": "情境定義與口徑"}, old_thesis_kv]
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", os.path.relpath(X, HERE))

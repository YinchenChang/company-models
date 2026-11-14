# 台達電 html.json 產生器（v0.5.3：DCF 彙總主表直接顯示「＋Stub 期 FCFF」、情境驅動表百分比一位小數（第三輪修正）；v0.5.2：DCF 彙總加 stub 期兩列、In 分組加第二輪修正新參數；v0.4.1：個股 HTML 版面規格 v1.0，docs/plan/個股HTML版面規格.md；v0.5：DCF 改標準版（FCFF＋WACC），估值章改放折現率、DCF 彙總與各情境 FCFF 明細）
# 讀 delta/html_src.json（原版面設定，內容不動）→ 依列標籤重定位與整理 → relayout() → 寫 delta/html.json（勿手改 html.json）。
# 用法：python3 delta/gen_html_json.py [成品 xlsx]（省略時用 delta/dist/ 內唯一的 xlsx）
import copy, glob, json, os, re, sys
import openpyxl
from openpyxl.utils import get_column_letter
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
                     for f in [{"%": "p1", "pt": "p1", "x": "n2", "權重": "p0", "年": "n0"}.get(str(IW.cell(r, 3).value), "n2")]]   # 第三輪修正（fixes_round3 第 31 項）：百分比一位小數

# ── 估值：目標價與現價長條圖、結論表（主要列）、DCF 折現率、相對倍數調整表；舊法與規則值摺疊 ──
va = sec["valuation"]["blocks"]
V = lambda lab, c="B", f="": f"@Val!{c}{row('Val', lab)}" + (f"|{f}" if f else "")
VA = lambda lab: f"@Val!A{row('Val', lab)}"
conc_labels = [("目標價（DCF 與 Comparable 加權）", "n0"), ("DCF 目標價（情境加權）", "n0"), ("Comparable 目標價（相對倍數調整後）", "n0"),
               ("Comparable 目標價（市場主流同業法）", "n0"), ("現價（", "n0"), ("上檔空間：加權目標價", "sp1"), ("上檔空間：Comparable 目標價", "sp1"),
               ("上檔空間級距", ""), ("投資建議", ""), ("市場共識目標價", "n0"), ("與市場共識目標價的差距：EPS 因素", "sp1"),
               ("與市場共識目標價的差距：倍數因素", "sp1"), ("價格帶（DCF 保守～樂觀）", ""), ("目標價方法", "")]
conc = {"type": "table", "caption": "目標價、現價與投資建議", "head": ["項目", "數值"], "rows": [[VA(l), V(l, f=f)] for l, f in conc_labels]}
hist = next(b for b in va if b["type"] == "table" and b.get("caption") == "目標價（DCF 與 Comparable 加權）")   # 原結論表：由上方 conc 取代
hr_ = row("Val", "歷史倍數檢查")
bars = {"type": "bars", "title": "目標價與現價", "fmt": "n0", "items": [
    {"label": VA("DCF 目標價（情境加權）"), "value": V("DCF 目標價（情境加權）"), "series": 0},
    {"label": VA("Comparable 目標價（相對倍數調整後）"), "value": V("Comparable 目標價（相對倍數調整後）"), "series": 1},
    {"label": VA("目標價（DCF 與 Comparable 加權）"), "value": V("目標價（DCF 與 Comparable 加權）"), "series": 2},
    {"label": VA("市場共識目標價"), "value": V("市場共識目標價"), "series": 0}],
    "ref": {"label": VA("現價（"), "value": V("現價（")}}


def dcf_blocks():
    """v0.5 標準版 DCF：折現率表（緊接 DCF 彙總之前）、DCF 彙總、各情境 FCFF 明細（摺疊）、口徑說明。列號依標籤定位。"""
    vr = lambda lab: row("Val", lab)
    rate_t = {"type": "table", "caption": "DCF 折現率：WACC、股權成本、稅後債務成本、RONIC、永續成長率", "head": ["項目", "數值", "說明"],
              "rows": [[f"@Val!A{vr(l)}", f"@Val!B{vr(l)}|{f}", f"@Val!E{vr(l)}"] for l, f in
                       (("股權成本 ke", "p2"), ("稅後債務成本", "p2"), ("權益占資本", "p1"), ("WACC", "p2"), ("永續成長率 g", "p2"), ("永續期 RONIC", "p2"))]}
    h = vr("股權成本 ke") - 1
    summ = [("明確期 FCFF 現值", "n1"), ("終值現值", "n1"), ("企業價值 EV", "n1"), ("＋淨現金與投資", "n1"), ("＋Stub 期 FCFF", "n1"), ("淨現金滾到評價日", "n1"), ("評價用稀釋股數", "n3"), ("DCF 情境每股價值（評價日）", "n0"),
            ("DCF 情境價值推到 12 個月後", "n0"), ("其中：終值現值占 EV", "p1"), ("隱含終值 EV／EBITDA", "x1")]
    sum_t = {"type": "table", "caption": "DCF 彙總：明確期現值、終值現值、企業價值、淨現金與每股價值", "head": [f"@Val!{c}{h}" for c in "ABCD"],
             "rows": [[f"@Val!A{vr(l)}"] + [f"@Val!{c}{vr(l)}|{f}" for c in "BCD"] for l, f in summ] + [[f"@Val!A{vr('DCF 目標價（情境加權）')}", f"@Val!B{vr('DCF 目標價（情境加權）')}|n0", "", ""]]}
    D = wb["DCF"]
    hd = next(r for r in range(1, 12) if D.cell(r, 1).value == "項目")
    cols = [get_column_letter(c) for c in range(3, D.max_column + 1) if D.cell(hd, c).value]
    out = [rate_t, sum_t]
    for r in range(1, D.max_row + 1):
        if D.cell(r, 1).value == "營收" and str(D.cell(r - 1, 1).value or "").startswith("情境"):
            pv = next(k for k in range(r, r + 20) if str(D.cell(k, 1).value or "").startswith("現值（FCFF"))
            rf = {str(k): {c: ("p1" if str(D.cell(k, 1).value).startswith(("稅率", "折舊攤銷 ÷", "資本支出 ÷")) else "n3") for c in cols}
                  for k in range(r, pv + 1) if str(D.cell(k, 1).value).startswith(("稅率", "折舊攤銷 ÷", "資本支出 ÷", "折現因子"))}
            out.append({"type": "table", "caption": f"@DCF!A{r - 1}", "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細",
                        "head": [f"@DCF!A{hd}"] + [f"@DCF!{c}{hd}" for c in cols],
                        "rowgen": {"sheet": "DCF", "rows": f"{r}-{pv}", "cols": ["A"] + cols, "fmts": [""] + ["n1"] * len(cols), "row_fmt": rf}})
    out.append({"type": "note", "parts": ["@DCF!A2"]})
    return out


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
    if cap == "@Val!A86":   # 舊 0.1 DCF（股東現金流）表：v0.5 改為標準版 DCF 的折現率、彙總與明細
        nb.extend(dcf_blocks())
        continue
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
        b["head"] = [f"@Sens!{c}{r0 + 1}" for c in "ABCDEFG"]
        b["rowgen"] = {"sheet": "Sens", "rows": f"{r0 + 2}-{row('Sens', '檢查：閉式公式', r0) - 1}", "cols": ["A", "B", "C", "D", "E", "F", "G"], "fmts": ["", "", "n0", "sp1", "n0", "sp1", "n0"]}

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
    ["延伸期收斂法（FY2030–FY2032）", ["tp_g", "fade_gA", "fade_gB", "fade_gC", "nim10_f"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "rf_us", "beta_u", "beta_du", "beta_de", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["da_r", "cx_r", "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["Wattonomics 錨定係數（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]]],
 "blocks_to": {"AI 電源與液冷": "AI 電源（電源及零組件事業群）", "@Units!A15": "AI 電源（電源及零組件事業群）", "標記沿用": "AI 電源（電源及零組件事業群）"},
 "main": {"目標價、現價與投資建議": ["目標價（DCF 與 Comparable 加權）", "DCF 目標價（情境加權）", "Comparable 目標價（相對倍數調整後）", "現價（", "上檔空間：加權目標價", "投資建議", "市場共識目標價"],
          "DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "0.2 Comparable": ["同業前瞻 P/E 中位數", "12 個月後的前瞻 EPS", "Comparable 目標價"],
          "相對倍數調整": ["同業 10 年中位數的中位數", "相對倍數調整係數", "Comparable 目標價（相對倍數調整後）"],
          "本模型與券商估計": ["FY2026 EPS vs 本土券商", "FY2027 EPS vs 投顧研究團隊", "CY2027 AI 電源市場（US$B）vs UBS", "FY2026 台達電有效份額 vs 凱基"]},
}
cfg = relayout(cfg, X, SPEC)

# ── 轉換後調整 ──
# 1. 橫幅第二格（補充規則 A）：只看 Comparable 目標價（市場主流同業法，未經相對倍數調整）與其上檔空間；
#    進加權的相對倍數調整後 Comparable 寫在橫幅附註（數字取自 Val 頁與 README 頁 HTML 橫幅附註用列）
hb = row("README", "HTML 橫幅附註用：進加權的 Comparable")
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", V("Comparable 目標價（市場主流同業法）", f="n2"), ["上檔空間 ", V("上檔空間：Comparable 目標價", f="sp1")]]
cfg["hero"]["note"] = cfg["hero"]["note"] + ["；進加權的 Comparable 為相對倍數調整後 ", V("Comparable 目標價（相對倍數調整後）", f="n2"), "（上檔空間 ", f"@README!B{hb}|sp1", "）"]
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
# ================= v0.5.6（2026-11-14）：個股 HTML 版面規格 v1.3（Andy 2026-11-14 16:57 退件後全面改版；樣板 NVIDIA nvidia/gen_html_json.py） =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含 C0）→ 資料狀態；台達電無個別模組。
# 數字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 加權目標價相對基準、Coord C0；delta/v056_post.py 產生）。列號依標籤找。
def _v13():
    S = {s["id"]: s for s in cfg["sections"]}
    RD, PLW, RV = wb["README"], wb["PL"], wb["RevDCF"]
    r13 = row("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.5.6"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    old_note = list(cfg["hero"]["note"])
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"]   # 橫幅只留白話說明（≤200 字）；評等邊界與相對倍數調整移到估值章
    # ---- PL 情境 B 推導區塊（命題圖與驅動因子共用）----
    rbB = row("PL", "情境 B：CY2028")
    hB = rbB + 1
    assert PLW.cell(hB, 1).value == "項目"
    pB = lambda lab, ex=True: row("PL", lab, rbB, ex)
    r_ai, r_lc, r_na, r_rev, r_sh = pB("AI 電源營收"), pB("液冷營收"), pB("非 AI 營收（", False), pB("營收"), pB("台達電有效份額")
    # ---- 1. 命題 ----
    th = S["thesis"]["blocks"]
    units = {"type": "theme", "title": "AI 電源、液冷與非 AI 事業營收（情境 B，十億元）與台達電 AI 電源有效份額（右軸）", "x": f"@PL!D{hB}:G{hB}",
             "fmt": "n0", "fmt_r": "p1", "stack": True, "pad": 1.25, "total": f"@PL!D{r_rev}:G{r_rev}",
             "bars": [{"name": ["AI 電源（電源及零組件事業群）"], "y": f"@PL!D{r_ai}:G{r_ai}", "color": "#0b7a80"},
                      {"name": ["液冷（基礎設施事業群）"], "y": f"@PL!D{r_lc}:G{r_lc}", "color": "#6fc7cb"},
                      {"name": ["非 AI 事業（電源非 AI、基礎設施非液冷、自動化、交通）"], "y": f"@PL!D{r_na}:G{r_na}", "color": "#c29a2b"}],
             "lines": [{"name": ["台達電在 AI 電源市場的有效份額"], "y": f"@PL!D{r_sh}:G{r_sh}", "axis": "r", "color": "#3b3838", "labels": True}],
             "caption": "堆疊＝三塊營收，頂端數字＝合併營收；折線＝台達電有效份額（右軸）。回答命題：AI 部署成長帶動 AI 電源與液冷營收，但份額逐年下滑，抵銷一部分成長。FY2026 為校準年。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    i_cmd = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "命題")
    th.insert(i_cmd + 2, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})   # 投資建議條件式（白話）
    th.insert(0, units)
    for b in th:
        if b["type"] == "h3" and b.get("text") == "情境定義與口徑":
            b["text"] = "情境定義"
    # ---- 2. 估值：四個目標價（Val 0.7）----
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b["type"] == "bars" and b.get("title") == "目標價與現價")
    g0 = row("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = row("Val", "現價（2026", g0)
    bars4 = {"type": "bars", "title": "四個目標價與現價（新台幣元；豎線＝現價）", "fmt": "n0",
             "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                       ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
             "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
    t07 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
           "rows": [[f"@Val!A{r}", f"@Val!B{r}|n0", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{g0 + 7}", f"@Val!B{g0 + 7}|n0", "", f"@Val!D{g0 + 7}"]]}
    va[i_bar] = bars4
    va.insert(i_bar, t07)
    va.insert(i_bar + 2, {"type": "note", "parts": old_note})
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").startswith(p))
    ra = rv("A. DCF"); rb_ = rv("(a)"); rbb = rv("(b)"); rc = rv("(c)"); rcmp = rv("B. Comparable"); rsum = rv("C. 摘要"); yh = rv("項目（情境 B）")
    m1, m0 = rv("現價隱含的營業利益率", rb_), rv("模型：基準的等效營業利益率", rb_)
    s1, rv1, rv0 = rv("現價隱含的營收倍數", rbb), rv("現價隱含的 FY2027 營收（", rbb), rv("模型：FY2027 營收（", rbb)
    g1, g0_ = rv("現價隱含的 FY2027 營收年增率", rbb), rv("模型：FY2027 營收年增率", rbb)
    w1 = rv("現價隱含的 WACC", rc)
    rev_ = {
        "id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
        "blocks": [
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
            {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
             "rows": [["預測期營業利益率（FY2027–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                      ["營收規模（FY2027 起各年同乘的倍數）", f"@RevDCF!B{s1}|n2", "—"],
                      ["FY2027 營收（十億元）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                      ["FY2027 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0_}|p1"],
                      ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w1 + 1}|p2"]]},
            {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"],
             "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if r in (ra + 2, ra + 3, ra + 9) else "n3" if r == ra + 5 else "x1" if r == ra + 8 else "n1")]
                      for r in range(ra + 2, ra + 10)]},
            {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
             "rows": [[f"@RevDCF!A{rcmp + 2}", "—", f"@RevDCF!C{rcmp + 2}|n2", "—"]]
                     + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!C{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!D{r}|sp1"] for k, r in enumerate(range(rcmp + 3, rcmp + 6))]
                     + [[f"@RevDCF!A{rcmp + 6}", "—", f"@RevDCF!C{rcmp + 6}|n1", "—"]]},
            {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
             "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）",
             "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|" + ("p1" if r in (yh + 3, yh + 4) else "n3" if r in (yh + 5, yh + 6, yh + 7, yh + 8) else "n1") for c in "BCDEFG"]
                      for r in range(yh + 1, rb_ - 1)]
                     + [{"sub": f"@RevDCF!A{rb_}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "利益率" in str(RV.cell(r, 1).value) else "n1"), "", "", "", "", ""] for r in range(rb_ + 1, rb_ + 4)]
                     + [{"sub": f"@RevDCF!A{rbb}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "年增率" in str(RV.cell(r, 1).value) else "n2" if "倍數" in str(RV.cell(r, 1).value) else "n3" if "核對" in str(RV.cell(r, 1).value) else "n1"), "", "", "", "", ""] for r in range(rbb + 1, rc - 1)]
                     + [{"sub": f"@RevDCF!A{rc}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p2" if str(RV.cell(r, 1).value).startswith(("現價隱含的 WACC", "模型 WACC")) else "n2"), "", "", "", "", ""] for r in range(w1, w1 + 4)]},
        ]}
    # ---- 5. 損益（標準損益表；GAAP）----
    p0 = row("PL", "標準損益表（v0.5.6")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: row("PL", lab_, start, True)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（研發＋銷管）", "營業利益", "營業利益率", "其他收入（淨額，營業外收益）",
            "稅前利益", "所得稅", "非控制權益損益", "稅後淨利（歸屬母公司）", "淨利率", "股數", "EPS（基本，公司公告）"]
    fmt = lambda u: {"%": "p1", "NT$": "n2", "十億股": "n3"}.get(u, "n1")
    cols = "CDEFGHIJK"
    rowx = lambda r, f=None: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{f or fmt(PLW.cell(r, 2).value)}" for c in cols]
    r_ratio, r_chk = plr("比率"), plr("核對（應為 0）")
    r_d = row("PL", "差額（FY2030E–FY2031E）：毛利 − 營業費用", ph)
    is_tbl = {"type": "table", "caption": "標準損益表（GAAP 口徑，十億元新台幣；EPS 為新台幣元）：FY2023A–FY2025A 實績、FY2026E 錨點年、FY2027E–FY2031E 情境 B（基準）",
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in cols],
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]
                      + [{"sub": "核對：FY2030E–FY2031E 兩種推法的營業利益差額（毛利率與營業費用率沿用 FY2029）"}] + [rowx(r) for r in (r_d, r_d + 1)]}
    chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in cols], "collapse": True,
               "summary": "展開：標準損益表的核對列（恆等式與對情境 B 摘要的差額，應為零）",
               "rows": [rowx(r, "n3") for r in range(r_chk + 1, until_blank("PL", r_chk + 1) + 1) if r not in (r_d, r_d + 1)]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    xl = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
    is_chart = {"type": "theme", "title": "營收（十億元）與毛利率、營業利益率、淨利率：FY2023A–FY2031E（情境 B）", "x": xl,
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:K{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:K{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:K{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:K{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸）。FY2026E 起為模型（上半年實績＋第三季實績＋第四季假設）；FY2030E、FY2031E 為延伸期，營收與淨利率由收斂法推得，毛利率與營業費用率沿用 FY2029。FY2027E 起標「對帳中」。"}
    x0 = row("PL", "延伸：三情境到 FY2032")
    xa, xb, xc = (row("PL", f"延伸｜情境 {k}", x0) for k in "ABC")
    xh = xa + 1
    xr = lambda base, lab_: row("PL", lab_, base, True)
    sc_chart = {"type": "theme", "title": "三情境並列：營收（長條，十億元）與 EPS（折線，新台幣元，右軸）", "x": f"@PL!D{xh}:J{xh}",
                "fmt": "n0", "fmt_r": "n0", "stack": False, "pad": 1.15,
                "bars": [{"name": ["情境 A（樂觀）營收"], "y": f"@PL!D{xr(xa, '營收（延伸）')}:J{xr(xa, '營收（延伸）')}", "color": "#b7e3e4"},
                         {"name": ["情境 B（基準）營收"], "y": f"@PL!D{xr(xb, '營收（延伸）')}:J{xr(xb, '營收（延伸）')}", "color": "#10a0a7"},
                         {"name": ["情境 C（保守）營收"], "y": f"@PL!D{xr(xc, '營收（延伸）')}:J{xr(xc, '營收（延伸）')}", "color": "#8a8c8c"}],
                "lines": [{"name": ["情境 A（樂觀）EPS"], "y": f"@PL!D{xr(xa, 'EPS（延伸）')}:J{xr(xa, 'EPS（延伸）')}", "axis": "r", "color": "#0b7a80", "dash": True},
                          {"name": ["情境 B（基準）EPS"], "y": f"@PL!D{xr(xb, 'EPS（延伸）')}:J{xr(xb, 'EPS（延伸）')}", "axis": "r", "color": "#3b3838", "labels": True},
                          {"name": ["情境 C（保守）EPS"], "y": f"@PL!D{xr(xc, 'EPS（延伸）')}:J{xr(xc, 'EPS（延伸）')}", "axis": "r", "color": "#b4472f", "dash": True}],
                "caption": "三情境在 FY2027 相同（錨點），差異自 FY2028 起出現；FY2030 起為延伸期（收斂法）。"}
    old_pl = S["pl"]["blocks"]
    old_intro = S["pl"].pop("intro", None)
    pl_old = [b for b in old_pl if b["type"] == "pl"]
    assert len(pl_old) == 2, len(pl_old)
    old_blocks = [dict(pl_old[0], collapse=True, summary="展開：三情境摘要表與折線（FY2025–FY2029；營收、AI 電源營收、毛利率、營業利益率、EPS）")]
    if old_intro:
        old_blocks.append({"type": "note", "parts": [old_intro], "collapse": True, "summary": "展開：三情境摘要的會計年度說明"})
    old_blocks.append(dict(pl_old[1], collapse=True, summary="展開：三情境延伸到 FY2032（收斂法）"))
    pl_blocks = [{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, chk_tbl, sc_chart] + old_blocks
    S["pl"].update(title="損益（標準損益表）", blocks=pl_blocks)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    uf = lambda u: {"US$B": "n1", "GW": "n1", "US$/kW": "n0", "%": "p1", "NT$B": "n1", "x": "n2", "NT$B/GW": "n2", "US$M/GW": "n0"}.get(u, "n1")
    r_a, r_b, r_c, r_k = pB("A. AI 部署與電源市場", False), pB("B. 液冷、非 AI 與合計", False), pB("C. 損益", False), pB("檢查：分部合計", False)
    segc = "CDEFG"
    seg_rows = lambda rs: [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{uf(PLW.cell(r, 2).value)}" for c in segc] for r in rs]
    seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；十億元新台幣）：AI 電源＝部署 GW × 每 kW 電源內容 × 台達電有效份額；液冷＝部署 GW × 每 GW 液冷營收；加非 AI 事業得合併營收",
               "head": [f"@PL!A{hB}", f"@PL!B{hB}"] + [f"@PL!{c}{hB}" for c in segc],
               "rows": [{"sub": "一、AI 部署與 AI 電源（電源及零組件事業群的命題子列）"}] + seg_rows(range(r_a + 1, r_b))
                       + [{"sub": "二、液冷、非 AI 事業與合併營收"}] + seg_rows([r for r in range(r_b + 1, r_c) if r != r_k])}
    r_gw, r_pk = pB("部署 GW（IT 電力）"), pB("AI 電源內容")
    gw_chart = {"type": "theme", "title": "部署 GW（IT 電力，長條）與每 kW AI 電源內容（折線，美元／kW，右軸）：情境 B", "x": f"@PL!D{hB}:G{hB}",
                "fmt": "n0", "fmt_r": "n0", "stack": False, "pad": 1.3,
                "bars": [{"name": ["部署 GW（IT 電力）"], "y": f"@PL!D{r_gw}:G{r_gw}", "color": "#0b7a80"}],
                "lines": [{"name": ["每 kW AI 電源內容"], "y": f"@PL!D{r_pk}:G{r_pk}", "axis": "r", "color": "#c29a2b", "labels": True}],
                "caption": "AI 電源營收的兩個產業輸入：部署量（讀 Siliconomics 基準，再乘情境 B 顆數係數）與每 kW 電源內容（讀 Siliconomics）；乘上台達電有效份額即為 AI 電源營收。"}
    dr = S["drivers"]["blocks"]
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, gw_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + dr
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    # ---- 9. 資料狀態 ----
    S["status"]["blocks"][0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    cfg["sections"].append(rev_)
    return RM


_RM = _v13()

# 三層協調 C4（規則 v1.4；tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(cfg, X)

# ---- v1.3 後處理（COORD.inject_html 之後）：估值與敏感度主圖置頂、敏感度加權目標價龍捲風圖、對帳章加 C0 ----
def _v13_post(RM):
    S = {s["id"]: s for s in cfg["sections"]}
    SW = wb["Sens"]
    # 敏感度：加權目標價相對基準（Sens 頁末表；v056_post.sens_tor）
    t0 = row("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SW.cell(r2 + 1, 1).value:
        r2 += 1
    out = []
    for col, side in (("B", "低"), ("C", "高")):
        num = lambda r: isinstance(SW[f"{col}{r}"].value, (int, float)) and not isinstance(SW[f"{col}{r}"].value, bool)
        big = [r for r in range(r1, r2 + 1) if num(r) and abs(SW[f"{col}{r}"].value) >= 0.1]
        small = [r for r in range(r1, r2 + 1) if num(r) and 0.0005 <= abs(SW[f"{col}{r}"].value) < 0.1]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        out.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
        out.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                        summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例"))
    out.append({"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
                "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1, r2 + 1)]})
    sb = S["sens"]["blocks"]
    for b in sb:
        if b.get("type") == "tornado":
            b["collapse"] = True
            b["summary"] = "展開：" + b["title"].replace("情境 A 的 FY2028 EPS 相對基準的變動", "對照：情境 A 的 FY2028 EPS 相對基準的變動")
        if b.get("type") == "table" and not b.get("collapse") and str(b.get("caption", "")).startswith("@Sens!A"):
            b["collapse"] = True
            b["summary"] = "展開：目標價參數表（折現率、股權風險、永續成長率、DCF 權重等）"
        if b.get("type") == "note" and not b.get("collapse") and b["parts"] and str(b["parts"][0]) in ("@Sens!A22", "@Sens!A1"):
            b["collapse"] = True
            b["summary"] = "展開：敏感度頁原文說明"
    S["sens"]["blocks"] = out + sb
    S["sens"]["intro"] = "主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入；變動達兩位數百分比者直接顯示，其餘摺疊）；下方摺疊為全部案例、目標價參數與 FY2028 EPS 敏感度對照。"
    # 位置（版面規格 v1.3 G、H 節）：主圖置頂、說明放圖下
    va = S["valuation"]["blocks"]
    i_bar = next(i for i, b in enumerate(va) if b.get("type") == "bars" and str(b.get("title", "")).startswith("四個目標價"))
    i_tbl, i_nt = i_bar - 1, i_bar + 1
    assert va[i_tbl].get("type") == "table" and str(va[i_tbl].get("caption", "")).startswith("@Val!A"), va[i_tbl]
    assert va[i_nt].get("type") == "note" and "投資建議依加權目標價判定" in str(va[i_nt].get("parts")), va[i_nt]
    top = [va[i_bar]]
    if S["valuation"].get("intro"):
        top.append({"type": "note", "parts": [S["valuation"].pop("intro")]})
    top += [va[i_tbl], va[i_nt]]
    S["valuation"]["blocks"] = top + [b for i, b in enumerate(va) if i not in (i_tbl, i_bar, i_nt)]
    sb = S["sens"]["blocks"]
    main = [i for i, b in enumerate(sb) if b.get("type") == "tornado" and not b.get("collapse") and str(b.get("title", "")).startswith("加權目標價相對基準")]
    assert len(main) == 2, main
    top = [sb[i] for i in main]
    if S["sens"].get("intro"):
        top.append({"type": "note", "parts": [S["sens"].pop("intro")]})
    S["sens"]["blocks"] = top + [b for i, b in enumerate(sb) if i not in main]
    # 對帳章：三層協調 C0（Coord 頁）
    CW = wb["Coord"]
    cr = lambda p, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
    c0, s1, s2, s3 = cr("三層協調 C0"), cr("1. 本模型讀哪一層"), cr("2. C0 由下往上對帳"), cr("3. 對帳結果")
    def f2(r, c):
        lab, u = str(CW.cell(r, 1).value), str(CW.cell(r, 6).value or "")
        if "差距" in lab:
            return "sp1"
        if "判定" in lab:
            return ""
        return {"GW": "n1", "US$B": "n1", "%": "p1", "US$/kW": "n0"}.get(u, "n2")
    c0t = {"type": "table", "caption": "由下往上對帳：本模型的產業量與 Siliconomics 層級值逐年對照（差距＝本模型相對層級值的百分比；本公司自己的量超過正負一成五標「對帳中」，不改上游值；下方「層級之間」列僅供參考）",
           "head": [f"@Coord!{c}{s2 + 1}" for c in "ABCDE"],
           "rows": [({"sub": f"@Coord!A{r}"} if CW.cell(r, 2).value is None else [f"@Coord!A{r}"] + [f"@Coord!{c}{r}|{f2(r, c)}" for c in "BCDE"]) for r in range(s2 + 2, s3 - 1)]}
    c0d = {"type": "table", "caption": f"@Coord!A{s2}", "collapse": True, "summary": "展開：各列的單位、口徑、來源與證據等級", "head": [f"@Coord!{c}{s2 + 1}" for c in "AFG"],
           "rows": [[f"@Coord!A{r}", f"@Coord!F{r}|", f"@Coord!G{r}|"] for r in range(s2 + 2, s3 - 1) if CW.cell(r, 6).value or CW.cell(r, 7).value]}
    lay = {"type": "table", "caption": f"@Coord!A{s1}", "collapse": True, "summary": "展開：本模型讀哪一層、哪一版、用在哪裡", "head": [f"@Coord!{c}{s1 + 1}" for c in "ABCD"],
           "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|", f"@Coord!C{r}|", f"@Coord!D{r}|"] for r in range(s1 + 2, s2 - 1)]}
    c0b = [{"type": "h3", "text": "三層協調：本模型由下往上對帳 Siliconomics"}, {"type": "note", "parts": [f"@Coord!A{c0 + 1}"]}]
    c0b += [{"type": "note", "parts": [f"@Coord!A{r}"]} for r in range(s3 + 1, s3 + 5)]
    c0b += [c0t, c0d, lay, {"type": "h3", "text": "券商、研調市場與份額對照（原有）"}]
    S["recon"]["blocks"] = c0b + S["recon"]["blocks"]
    # 全頁可見文字（摺疊與資料狀態以外）不出現參數代號（版面規格 v1.3 B 節）：含代號的儲存格改用白話（設定檔文字，不含數字）；xlsx 原文不改
    PLAIN = {
        "Val!E91": "股權成本與稅後債務成本以市值權重加權；WACC 調整只用於敏感度（基準為零）",
        "Val!A92": "永續成長率 g（長期成長上限與無風險利率取小者）",
        "Val!A93": "永續期 RONIC（＝WACC＋超額報酬差）",
        "Val!A102": "DCF 情境價值推到目標時點（以股權成本複利，扣除期間股利）",
        "In!B114": "永續成長率 g（最後一年之後；長期成長上限與無風險利率取小者）",
        "In!B97": "情境 A（樂觀）：FY2032 營收年增收斂目標（基準收斂目標加兩個百分點）",
        "In!B152": "FY2032 營收年增收斂目標（A 為基準加兩個百分點、B 為基準、C 為基準減兩個百分點）",
    }
    FOLD = {"DCF!A2": "展開：DCF 的營業利益與自由現金流口徑"}

    def _plain(spec):
        if isinstance(spec, list):
            return [_plain(x) for x in spec]
        if isinstance(spec, str) and spec.startswith("@"):
            ref = spec[1:].split("|")[0]
            if ref in PLAIN:
                return PLAIN[ref]
        return spec
    for sc in cfg["sections"]:
        if sc["id"] == "status":
            continue
        nb = []
        for b in sc["blocks"]:
            if b.get("type") == "table":
                b["rows"] = [({"sub": _plain(r["sub"])} if isinstance(r, dict) else _plain(r)) for r in b.get("rows", [])]
                if "head" in b:
                    b["head"] = _plain(b["head"])
            if b.get("type") == "note" and not b.get("collapse") and len(b["parts"]) == 1 and str(b["parts"][0])[1:] in FOLD:
                b = dict(b, collapse=True, summary=FOLD[str(b["parts"][0])[1:]])
            nb.append(b)
        sc["blocks"] = nb
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    cfg["sections"] = [S[k] for k in order]


_v13_post(_RM)
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", os.path.relpath(X, HERE))

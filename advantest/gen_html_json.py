# v0.7：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）＋DCF 改標準版（FCFF＋WACC），估值章 DCF 區塊與目標價敏感度表由本檔依 xlsx 標籤產生
# 由 html_src.json（舊版面設定，內容不動）經共用轉換器 tools/company_html/relayout.py 產生 html.json（勿手改 html.json）。
# 用法：python3 advantest/gen_html_json.py（讀 advantest/dist 內唯一的 xlsx）
# 前處理（本檔）：把設備股模板的舊版面調成轉換器預期的形狀——單一參數表改作「驅動 → 推導量對照」（進附錄）、
#   情境驅動表與敏感度目標價參數表的列號依現況重新定位、估值章的 kv 改成具名表格、舊法規則值與同業名單摺疊、
#   補 DCF 折現率表與目標價敏感度龍捲風圖；
# 後處理：還原驅動因子章說明、EPS 龍捲風圖標題（轉換器預設寫「加權目標價」）、橫幅中間格＝只看 Comparable（市場主流同業法），調整後值進附註。
import copy, glob, json, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
from relayout import relayout, _unit_fmt  # noqa: E402

X = glob.glob(os.path.join(HERE, "dist", "*.xlsx"))
assert len(X) == 1, X
X = X[0]
cfg = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)


def text(spec):
    m = re.match(r"^@([^!]+)!([A-Z]+\d+)", spec) if isinstance(spec, str) else None
    return str(wb[m.group(1)][m.group(2)].value or "") if m else str(spec or "")


def row(sh, prefix, exact=False, start=1):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or "").strip()
        if (v == prefix) if exact else v.startswith(prefix):
            return r
    raise KeyError((sh, prefix))


IN = {wb["In"].cell(r, 1).value: r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}


def in_row(k, fmt):
    r = IN[k]
    return [f"@In!B{r}", f"@In!E{r}|{fmt}", f"@In!G{r}|tag"]


def kv_table(b, caption=None, collapse=False, summary=None):
    t = {"type": "table", "head": ["項目", "數值"], "rows": [[k, v] for k, v in b["items"]]}
    if caption:
        t["caption"] = caption
    if collapse:
        t.update(collapse=True, summary=summary)
    return t


# ---------------- v0.4：標準版 DCF（FCFF＋WACC）的估值與敏感度區塊（由 xlsx 標籤定位列號） ----------------
def dcf_summary_block(b0):
    """Val 0.1 節：DCF 彙總（明確期現值～DCF 目標價；折現率另列在前一張表）。"""
    h = row("Val", "0.1 DCF")
    r1, r2 = row("Val", "明確期 FCFF 現值", start=h), row("Val", "DCF 目標價（情境加權）", start=h)
    rf = {str(r): {"B": f, "C": f, "D": f} for r, f in ((row("Val", "評價用稀釋股數", start=h), "n3"), (row("Val", "DCF 情境每股價值", start=h), "n2"),
                                                       (row("Val", "DCF 情境價值推到 12 個月後", start=h), "n2"), (row("Val", "其中：終值現值占 EV", start=h), "p1"),
                                                       (row("Val", "隱含終值 EV／EBITDA", start=h), "x1"), (r2, "n2"))}
    return {"type": "table", "caption": f"@Val!A{h}", "head": [f"@Val!A{h + 1}", f"@Val!B{h + 1}", f"@Val!C{h + 1}", f"@Val!D{h + 1}", f"@Val!E{h + 1}"],
            "rowgen": {"sheet": "Val", "rows": f"{r1}-{r2}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1", ""], "row_fmt": rf}}


def dcf_detail_blocks():
    """DCF 頁：三情境 FCFF 明細（摺疊）＋口徑說明。"""
    ws = wb["DCF"]
    hdr_r = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "項目" and str(ws.cell(r, 3).value or "").startswith("FY"))
    ncol = max(c for c in range(3, 12) if ws.cell(hdr_r, c).value)
    cols = [chr(64 + c) for c in range(3, ncol + 1)]
    out = []
    for r in range(hdr_r, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").startswith("情境 ") and ws.cell(r + 1, 1).value == "營收":
            r1, r2 = r + 1, next(x for x in range(r + 1, ws.max_row + 1) if str(ws.cell(x, 1).value or "").startswith("現值（FCFF"))
            rf = {}
            for x in range(r1, r2 + 1):
                lab = str(ws.cell(x, 1).value)
                f = "p1" if (lab.startswith("稅率") or "÷ 營收" in lab) else ("n3" if lab.startswith("折現因子") else None)
                if f:
                    rf[str(x)] = {c: f for c in cols}
            out.append({"type": "table", "caption": f"@DCF!A{r}", "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細",
                        "head": [f"@DCF!A{hdr_r}"] + [f"@DCF!{c}{hdr_r}" for c in cols],
                        "rowgen": {"sheet": "DCF", "rows": f"{r1}-{r2}", "cols": ["A"] + cols, "fmts": [""] + ["n1"] * len(cols), "row_fmt": rf}})
    out.append({"type": "note", "parts": ["@DCF!A2"]})
    return out


def tp_sens_block():
    """Sens 頁：目標價參數敏感度（標準版 DCF 案例；低／高加權目標價與相對基準）。"""
    h = row("Sens", "目標價參數敏感度")
    last = row("Sens", "DCF 權重 0.3／0.7", start=h)
    return {"type": "table", "caption": f"@Sens!A{h}", "head": [f"@Sens!{c}{h + 1}" for c in "ABCDEFG"],
            "rowgen": {"sheet": "Sens", "rows": f"{h + 2}-{last}", "cols": list("ABCDEFG"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2"]}}


def prep(cfg, spec):
    secs = {s["id"]: s for s in cfg["sections"]}
    # 驅動因子：單一參數表 → 驅動 → 推導量對照（轉換器移到附錄）；情境驅動表不摺疊
    dr = secs["drivers"]
    for b in dr["blocks"]:
        cap = str(b.get("caption", ""))
        if cap.startswith("單一參數"):
            b["caption"] = "驅動 → 推導量對照（單一參數的輸入、基準值、推導量與因果方向）"
        elif cap.startswith("情境驅動"):
            b.pop("collapse", None)
            b.pop("summary", None)
            # 列號依 In 頁現況重新定位（舊設定的列號在 In 頁插入 Wattonomics 錨定係數後已位移）
            ws = wb["In"]
            h = row("In", "情境驅動對照") + 1
            assert ws.cell(h, 1).value == "鍵", ws.cell(h, 1).value
            fm = lambda r: "yr" if ws.cell(r, 3).value == "年" else _unit_fmt(ws.cell(r, 3).value)
            b["head"] = [f"@In!B{h}", f"@In!D{h}", f"@In!E{h}", f"@In!F{h}", f"@In!C{h}"]
            b["rows"] = [[f"@In!B{r}", f"@In!D{r}|{fm(r)}", f"@In!E{r}|{fm(r)}", f"@In!F{r}|{fm(r)}", f"@In!C{r}"]
                         for r in range(h + 1, ws.max_row + 1) if ws.cell(r, 1).value]
    intro = dr.get("intro")
    # 估值：依規格重排
    va = secs["valuation"]
    old = va["blocks"]
    nb, i = [], 0
    caps = [text(b.get("caption", "")) for b in old]
    k_dcf = next(j for j, c in enumerate(caps) if c.startswith(spec["cap_dcf"]))
    k_comp = next(j for j, c in enumerate(caps) if c.startswith(spec["cap_comp"]))
    k_old = next(j for j, b in enumerate(old) if b["type"] == "h3" and str(b.get("text", "")).startswith(spec["h3_old"]))
    conc = old[0]
    assert text(conc.get("caption", "")).endswith("結論"), conc
    head = old[:k_dcf]
    bars = next(b for b in head if b["type"] == "bars")
    ca = row("Val", "Comparable 目標價（相對倍數調整後）")
    if not any(it["value"] == f"@Val!B{ca}" for it in bars["items"]):   # 加權目標價實際使用的 Comparable（相對倍數調整後）也畫出，避免讀成加權價高於兩法
        k = next(j for j, it in enumerate(bars["items"]) if text(it["label"]).startswith("Comparable 目標價"))
        bars["items"].insert(k + 1, {"label": f"@Val!A{ca}", "value": f"@Val!B{ca}", "series": 2})
    nb.append(bars)
    nb.append(conc)
    for j, b in enumerate(head[1:], 1):
        if b["type"] == "note":
            nb.append(b)
    rm_h3 = next(b for b in head if b["type"] == "h3" and str(b.get("text", "")).startswith("相對倍數調整"))
    rm_kv = head[head.index(rm_h3) + 1]
    assert rm_kv["type"] == "kv"
    nb.append({"type": "table", "caption": "DCF 折現率：股權成本、稅後債務成本、WACC、永續成長率、RONIC（標準版 DCF；直接顯示）", "head": ["項目", "數值", "標記"],
               "rows": [in_row(k, f) for k, f in spec["dcf_rates"]] + [[f"@Val!A{r}", f"@Val!B{r}|p2", ""] for r in spec.get("dcf_rate_val", [])]})
    nb.append({"type": "h3", "text": "DCF 彙總"})
    nb.append(dcf_summary_block(old[k_dcf]))
    nb += dcf_detail_blocks()
    nb.append({"type": "h3", "text": "Comparable（同業前瞻本益比）"})
    comp_kv = old[k_comp + 1]
    assert comp_kv["type"] == "kv"
    nb.append(kv_table(comp_kv, caption="Comparable 彙總（同業前瞻本益比中位數 × 前瞻每股盈餘）"))
    nb.append(kv_table(rm_kv, caption=rm_h3["text"]))
    peers = dict(old[k_comp], collapse=True, summary="展開：同業前瞻本益比（各家）")
    nb.append(peers)
    nb.append({"type": "h3", "text": old[k_old]["text"]})
    for b in old[k_old + 1:]:
        if b["type"] == "bars":
            continue   # 舊法隱含股價長條圖不放（轉換器會把所有長條圖移到章首，與目標價長條圖混淆）
        if b["type"] == "kv":
            nb.append(kv_table(b, collapse=True, summary="展開：舊法的價格與市場隱含倍數"))
        elif b["type"] == "table":
            if not b.get("collapse"):
                c = re.sub(r"^[\d.\s]+", "", text(b.get("caption", ""))).strip()
                b = dict(b, collapse=True, summary="展開：" + (c if c and not re.search(r"\d", re.sub(r"(FY|CY)\d{4}", "", c)) else "舊法明細表"))
            nb.append(b)
        elif b["type"] == "note":
            nb.append(dict(b, collapse=True, summary="展開：舊法說明"))
        else:
            nb.append(b)
    va["blocks"] = nb
    # 對帳：給表格標題，主要列由 SPEC main 指定
    for b in secs["recon"]["blocks"]:
        if b["type"] == "table" and not b.get("caption"):
            b["caption"] = "對帳要點"
    # 敏感度：先放目標價龍捲風圖，再放 EPS 龍捲風圖
    se = secs["sens"]
    tor = {}
    for b in se["blocks"]:
        if b["type"] == "tornado":
            tor[b["value_col"]] = b["title"].split("（")[0].replace("相對基準的變動", "").strip()
    ws = wb["Sens"]
    h0 = row("Sens", "目標價參數敏感度")
    for b in se["blocks"]:   # 目標價參數案例表：列號依 Sens 頁現況重新定位（舊設定在插入 Wattonomics 案例後可能位移）
        rg = b.get("rowgen") if b["type"] == "table" else None
        if rg and isinstance(b.get("caption"), str) and b["caption"].startswith("@Sens!A") and abs(int(b["caption"][7:]) - h0) <= 6:
            hh = h0 + 1
            assert str(ws.cell(hh, 1).value).strip() in ("案例", "參數"), ws.cell(hh, 1).value
            end = hh + 1
            while ws.cell(end + 1, 1).value not in (None, ""):
                end += 1
            b["caption"] = f"@Sens!A{h0}"
            b["head"] = [f"@Sens!{c}{hh}" for c in rg["cols"]]
            rg["rows"] = f"{hh + 1}-{end}"
    # v0.4：目標價參數敏感度改為標準版 DCF 案例（低／高兩欄）；不畫龍捲風圖，表格直接顯示
    se["blocks"] = [tp_sens_block() if (b["type"] == "table" and not b.get("collapse") and text(b.get("caption", "")).startswith("目標價參數")) else b
                    for b in se["blocks"]]
    scen_kv = next((b for b in secs["thesis"]["blocks"] if b["type"] == "kv"), None)
    return intro, tor, scen_kv


def post(cfg, intro, tor, scen_kv, spec):
    secs = {s["id"]: s for s in cfg["sections"]}
    if intro:   # 轉換器把章說明當 note 放進某一塊；還原為章說明
        secs["drivers"]["intro"] = intro
        secs["drivers"]["blocks"] = [b for b in secs["drivers"]["blocks"] if not (b["type"] == "note" and b.get("parts") == [intro])]
    # 參數表格式：轉換器只認單位「%」，「%/年」等也改百分比；小於 1 的一般數值用三位小數（避免 0.0／0.1 誤讀）
    ws = wb["In"]
    for b in secs["drivers"]["blocks"]:
        for r_ in b.get("rows", []) if b["type"] == "table" else []:
            m = re.match(r"^@In!B(\d+)$", r_[0]) if isinstance(r_, list) and isinstance(r_[0], str) else None
            if not m:
                continue
            r = int(m.group(1))
            u = str(ws.cell(r, 3).value or "")
            vals = [ws.cell(r, c).value for c in (4, 5, 6)]
            for j in (1, 2, 3):
                cell, _, f = r_[j].partition("|")
                if u.startswith("%"):
                    f = "p1"
                elif f == "n1" and any(isinstance(x, (int, float)) and 0 < abs(x) < 1 for x in vals):
                    f = "n3"
                elif f == "n1" and any(isinstance(x, float) and abs(x) < 10 and x != int(x) for x in vals):
                    f = "n2"
                r_[j] = f"{cell}|{f}"
    if scen_kv:   # 舊版命題章的情境定義：放在重點觀察之後（摺疊）
        secs["thesis"]["blocks"].append(dict(scen_kv, collapse=True, summary="展開：三情境的定義"))
    for b in secs["sens"]["blocks"]:
        if b["type"] == "tornado" and tor.get(b["value_col"]):
            base = tor[b["value_col"]] + " 相對基準的變動"
            b["title"] = base + ("（未達兩位數百分比）" if b.get("collapse") else "（變動達兩位數百分比的案例；一次改一個輸入，整本重算）")
        if b["type"] == "tornado" and b.get("collapse"):
            b["summary"] = "展開：" + (tor.get(b["value_col"]) or "加權目標價") + " 變動未達兩位數百分比的案例"
    # 橫幅（補充規則 A）：中間格＝只看 Comparable 目標價（市場主流同業法，未做相對倍數調整）與其上檔空間；
    #   進加權的相對倍數調整後 Comparable 放附註（數字由 xlsx 帶出）
    v = lambda p, **kw: row("Val", p, **kw)
    tp, ca, cu, px, up, rec = v("目標價（DCF 與 Comparable 加權）", exact=True), v("Comparable 目標價（相對倍數調整後）"), v("Comparable 目標價", exact=True), v("現價（"), v("上檔空間：加權目標價"), v("投資建議", exact=True)
    rd = row("README", "HTML 橫幅附註用：Comparable 目標價（相對倍數調整後）")
    uc = v("上檔空間：Comparable")
    cfg["hero"] = {"rec": f"@Val!B{rec}",
                   "tiles": [[f"@Val!A{tp}", f"@Val!B{tp}|n2", ["上檔空間 ", f"@Val!B{up}|sp1"]],
                             ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{cu}|n2", ["上檔空間 ", f"@Val!B{uc}|sp1"]],
                             [f"@Val!A{px}", f"@Val!B{px}|n2", ""]],
                   "note": ["投資建議依加權目標價判定：", f"@Val!E{rec}", "。進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{ca}|n2",
                            "（上檔空間 ", f"@README!B{rd}|sp1", "）。"]}


def in_skip():
    ws = wb["In"]
    out = []
    for r in range(4, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "")
        if a and ws.cell(r, 2).value in (None, "") and a.startswith("情境驅動對照"):
            out.append(re.split(r"[.\s]", a)[0])
    return out


def run(SPEC, PRE):
    intro, tor, scen_kv = prep(cfg, PRE)
    SPEC = dict(SPEC, in_skip=SPEC.get("in_skip", []) + in_skip())
    out = relayout(cfg, X, SPEC)
    post(out, intro, tor, scen_kv, PRE)
    json.dump(out, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("ok")



# ---------------- Advantest 專屬：參數依業務性質分塊（SoC 測試機／記憶體測試機／其他系統與服務／全公司） ----------------
PRE = {"cap_dcf": "0.1 DCF", "cap_comp": "0.2 Comparable", "h3_old": "舊法",
       "dcf_rates": [("rf", "p2"), ("beta", "n2"), ("erp", "p1"), ("crp", "p2")],
       "dcf_rate_val": [row("Val", p) for p in ("股權成本 ke", "稅後債務成本", "權益占資本", "WACC", "永續成長率 g", "永續期 RONIC")]}
_u = lambda ks: [f"u_{k}_{y}" for k in ks for y in (26, 27, 28)] + [f"g29_{k}" for k in ks]
_mem = lambda p: [f"{p}{s}_{y}" for s in "ABC" for y in range(25, 30)]
SPEC = {
 "in_groups": [
  ["SoC 測試機（AI 加速器與非 AI SoC）", [
    ["NVIDIA 各世代加速器顆數（markets/memory）", _u(("bw", "rb", "ru", "hp"))],
    ["其他加速器顆數（AMD、TPU、Trainium、其他）", _u(("amd", "tpu", "trn", "oth"))],
    ["情境顆數係數（讀台積電模型）與 CY2030 成長", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C", "cy30g"]],
    ["SoC 測試機市場、每顆測試時間係數與校準", ["mk_soc25", "mk_soc26", "nonai26", "g_nonai", "eff_par", "asp_m", "tt_bw", "tt_rb", "tt_ru", "tt_hp", "tt_amd", "tt_tpu", "tt_trn", "tt_oth"]],
    ["SoC 市占（情境 C 泰瑞達搶回市占）", ["sh_soc", "shC_28", "shC_29", "sh_soc25"]],
    ["Wattonomics 錨定係數（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  ["記憶體測試機（HBM 與 DRAM）", [
    ["記憶體測試機市場與市占", ["mk_mem25", "mk_mem26", "sh_mem", "w_hbm"]],
    ["HBM 出貨位元指數（markets/memory，三情境）", _mem("hs")],
    ["DRAM 設備支出指數（markets/equipment，三情境）與 SEMI 測試設備對照", _mem("dr") + ["semi_test26", "semi_test27", "semi_test28"]]]],
  ["其他系統與服務", [
    ["成長率（三情境）與 FY2026 指引", ["g_oth_26", "g_svc_26", "g_oth_A", "g_svc_A", "g_oth_B", "g_svc_B", "g_oth_C", "g_svc_C", "g_oth", "g_svcg"]]]],
  ["全公司（毛利與費用、指引、股價與目標價規則）", [
    ["毛利率與營業費用（兩點推導）", ["gm_g", "gm_a", "gm_b", "gm_adj_C28", "gm_adj_C29", "opex25", "opex_g", "opex_var", "opex_fix", "phi"]],
    ["淨利、股數、匯率", ["nioi", "shares", "fx"]],
    ["FY2026 公司指引與共識", ["g_rev", "g_oi", "g_ni", "g_soc", "g_mem", "g_eps", "cons_rev26", "cons_oi26"]],
    ["股價與循環位置", ["px", "sa_fpe", "pk_A", "pk_B", "pk_C"]],
    ["目標價規則（情境權重、DCF 權重、延伸期收斂與共識）", ["tp_wA", "tp_wB", "tp_wC", "tp_g", "tp_w_dcf", "nim10", "cons_tp", "cons_feps"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc", "da_rev", "cx_rev"]]]]],
 "in_skip": [],
 "main": {"0.1 DCF": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "對帳要點": ["營收", "SoC 測試機", "營業利益", "EPS（日圓）", "SoC 測試機市場 CY2026"]},
}
if __name__ == "__main__":
    run(SPEC, PRE)

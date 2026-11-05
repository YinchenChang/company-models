# 記憶體樣板（SK 海力士、美光、三星）的 HTML 新版面：html_src.json → 個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）→ html.json。
# 用法：由各公司 <公司>/gen_html_json.py 呼叫 build(公司資料夾, 成品 xlsx, extra_blocks=...)。
# 共用轉換器 tools/company_html/relayout.py 依 Meta／SpaceX 的頁面結構寫成；記憶體樣板的頁面結構不同，本檔在轉換前後做記憶體專屬的調整：
#   轉換前：驅動對照表改名歸附錄、情境驅動表展開、估值頁的 kv 改為有標題的表、加 DCF 折現率表、移除估值章重複的重點觀察。
#   轉換時：In 參數依業務分塊（一般 DRAM／HBM／NAND／記憶體共用成本／〔非記憶體事業〕／全公司）。
#   投資建議橫幅：Val 頁沒有「現價」列（現價在 In 頁 px），所以本檔提供記憶體專用的橫幅，不用轉換器的預設橫幅。
#   中間格＝只看 Comparable（市場主流同業法、未做相對倍數調整；Andy 2026-10-25 目標價標準），上檔空間取 README 頁輔助列；
#   進加權目標價的是相對倍數調整後的 Comparable，數值寫在橫幅附註。
#   轉換後：敏感度龍捲風圖的標題還原（記憶體的龍捲風圖是營業利益的變動，不是加權目標價）、版本要點附錄只列版本說明列、
#   循環位置長條圖放回規則值表之後、情境定義放在驅動因子章開頭。
import copy
import json
import os
import re
import sys

import openpyxl

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "company_html"))
import relayout as RL  # noqa: E402

B_DRAM, B_HBM, B_NAND, B_COST, B_NM, B_CO = ("一般 DRAM", "HBM", "NAND", "記憶體成本、折舊與資本支出（三條產品線共用）",
                                           "非記憶體事業（輸入，不建模）", "全公司：股價、資產負債、股東回饋、估值與目標價規則")


def classify(k, nm_block):
    """In 頁鍵 → (塊, 小組)。未知的鍵直接失敗（避免漏列）。"""
    m = re.match(r"^(cq|nq)([ABC]?)_(\d+)$", k)
    if m:
        b = B_DRAM if m.group(1) == "cq" else B_NAND
        return b, (f"合約價季增：情境 {m.group(2)}" if m.group(2) else "合約價季增：已發生季度與基準年")
    if re.match(r"^q[34]_d[bp]$", k):
        return B_DRAM, "基準年剩餘季度（公司指引、TrendForce）"
    if re.match(r"^q[34]_n[bp]$", k):
        return B_NAND, "基準年剩餘季度（公司指引、TrendForce）"
    if re.match(r"^cb_\d$", k):
        return B_DRAM, "年度位元成長"
    if re.match(r"^nb_\d$", k):
        return B_NAND, "年度位元成長"
    if re.match(r"^sh\d\d$", k):
        return B_HBM, "HBM 營收市占"
    if re.match(r"^[ug]_[a-z]+_\d\d$", k):
        return B_HBM, "加速器顆數與每顆 HBM 容量"
    if re.match(r"^um2[89]_[ABC]$", k):
        return B_HBM, "情境顆數係數（與台積電、Advantest 同一套）"
    if k.startswith("hbm_"):
        return B_HBM, "HBM 市場：位元、供給與單價"
    if k in ("nv_dc26", "nv_content", "nv_hbm", "pub26_lo", "pub26_hi", "pub27"):
        return B_HBM, "對帳用（橋接表與已發布的 HBM 市場估計）"
    if k.startswith("wta_"):
        return B_HBM, "Wattonomics 錨定係數（只用於敏感度）"
    if k in ("costdn_q", "costdn_y", "w_c", "w_h", "w_n"):
        return B_COST, "單位成本與成本權重"
    if k in ("da_q3", "da_q4", "da_life", "da_cogs"):
        return B_COST, "折舊"
    if k in ("opex_fix_q", "opex_fix_y", "opex_var", "opex_g"):
        return B_COST, "營業費用"
    if re.match(r"^cap_\d$", k) or k == "capex_rem":
        return B_COST, "資本支出"
    if k in ("other_y", "nonop_q"):
        return B_COST, "其他營收與業外收益"
    if k.startswith("nm_"):
        return (B_NM, "營收與營業利益") if nm_block else (B_CO, "非記憶體事業（本公司沒有，皆為零）")
    if k in ("fx", "px"):
        return B_CO, "股價與匯率"
    if k in ("tax", "ncy"):
        return B_CO, "稅與淨現金收益率"
    if k in ("rf", "erp", "crp", "beta", "beta_u", "beta_u_med", "beta_u_us", "beta_u_gl", "beta_de", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"):
        return B_CO, "折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）"
    if k in ("debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "sh_v"):
        return B_CO, "資本結構、淨現金與股數（評價用）"
    if k.startswith("fc_"):
        return B_CO, "自由現金流（折舊、資本支出、營運資金的長期穩態）"
    if k in ("eq_last", "nc_last", "sh_last", "extra_eq", "extra_sh"):
        return B_CO, "資產負債錨點（最後實績季末）"
    if k in ("bb_0", "bb_sh_0", "div_0", "payout", "bb_px") or re.match(r"^dps_\d$", k):
        return B_CO, "股東回饋（股利與買回）"
    if re.match(r"^p[be]_\d$", k):
        return B_CO, "固定倍數格（對照）"
    if k == "tp_g":
        return B_CO, "收斂期（FY2030 起回到循環中段）"
    if k in ("tp_wA", "tp_wB", "tp_wC", "tp_w_dcf"):
        return B_CO, "目標價規則（情境權重、DCF 權重）"
    if k == "nim10" or k.startswith(("fade_", "cyc_")):
        return B_CO, "收斂期（FY2030 起回到循環中段）"
    if k.startswith("cons_"):
        return B_CO, "市場共識"
    raise SystemExit(f"In 鍵 {k!r} 未分組（tools/memory_model/html_layout.py classify）")


def spec_for(wb, nm_block):
    ws = wb["In"]
    skip, keys, sec = [], [], None
    for r in range(4, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if not a:
            continue
        a = str(a)
        if ws.cell(r, 2).value in (None, ""):
            sec = re.split(r"[.\s]", a)[0]
            if a.startswith("情境驅動對照"):
                skip.append(sec)
            continue
        if sec in skip:
            continue
        keys.append(a)
    order = [B_DRAM, B_HBM, B_NAND, B_COST] + ([B_NM] if nm_block else []) + [B_CO]
    groups = {b: {} for b in order}
    for k in keys:
        b, g = classify(k, nm_block)
        groups[b].setdefault(g, []).append(k)
    return {"in_skip": skip, "in_groups": [[b, [[g, ks] for g, ks in groups[b].items()]] for b in order if groups[b]]}


def _find(ws, prefix, col=1):
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, col).value or "").startswith(prefix):
            return r
    raise KeyError((ws.title, prefix))


def build(co_dir, xlsx, nm_block=False, pfmt="n0", recon_main=None):
    src = os.path.join(co_dir, "html_src.json")
    out = os.path.join(co_dir, "html.json")
    cfg = json.load(open(src, encoding="utf-8"))
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    V, IN, RD = wb["Val"], wb["In"], wb["README"]
    ink = {str(IN.cell(r, 1).value): r for r in range(4, IN.max_row + 1) if IN.cell(r, 1).value}
    txt = lambda ref: str(wb[ref.split("!")[0][1:]][ref.split("!")[1]].value or "") if isinstance(ref, str) and ref.startswith("@") else str(ref or "")
    secs = {s["id"]: s for s in cfg["sections"]}

    # ── 轉換前 ──
    for b in secs["drivers"]["blocks"]:
        cap = str(b.get("caption", ""))
        if cap.startswith("單一參數"):
            b["caption"] = "驅動 → 推導對照表（In 頁基準值與 Y 頁推導量）"
        if cap.startswith("情境驅動"):
            # 原 html.json 的列號停在加入 Wattonomics 錨定（Z 節）之前，已錯位；改由 In 頁「情境驅動對照」一節直接產生
            b.pop("collapse", None)
            b.pop("summary", None)
            h = _find(IN, "情境驅動對照") + 1
            b["head"] = [f"@In!{c}{h}" for c in "BDEFCG"]
            b["rows"] = []
            for r in range(h + 1, IN.max_row + 1):
                if not IN.cell(r, 1).value:
                    break
                f = RL._unit_fmt(IN.cell(r, 3).value)
                b["rows"].append([f"@In!B{r}", f"@In!D{r}|{f}", f"@In!E{r}|{f}", f"@In!F{r}|{f}", f"@In!C{r}", f"@In!G{r}|tag"])
    for b in secs["status"]["blocks"]:
        ref = str(b.get("parts", [""])[0]) if b["type"] == "note" else ""
        if ref.startswith("@README!") and not txt(ref).startswith(("已知簡化", "集團公司處理", "最大限制")):
            try:
                b["parts"] = [f"@README!A{_find(RD, '已知簡化')}"]   # 原列號錯位（指到版本說明），改為「已知簡化」列
            except KeyError:
                pass
    va = secs["valuation"]["blocks"]
    nb, i = [], 0
    while i < len(va):
        b = va[i]
        if b["type"] == "h3" and b.get("text") == "重點觀察" and i + 1 < len(va) and va[i + 1]["type"] == "list":
            i += 2          # 重點觀察移到第一章（README 展示層）
            continue
        if b["type"] == "h3" and str(b.get("text", "")).startswith("相對倍數調整") and i + 1 < len(va) and va[i + 1]["type"] == "kv":
            nb.append({"type": "table", "caption": b["text"], "head": ["項目", "數值"], "rows": [list(x) for x in va[i + 1]["items"]]})
            i += 2
            continue
        if b["type"] == "kv":
            items = b["items"]
            if len(items) == 1 and items[0][0] == f"@In!B{ink['px']}":
                i += 1      # 現價已在橫幅與長條圖
                continue
            nb.append({"type": "table", "caption": "Comparable 目標價的計算（市場主流同業法，未做相對倍數調整）", "head": ["項目", "數值"], "rows": [list(x) for x in items]})
            i += 1
            continue
        if b["type"] == "table" and txt(b.get("caption", "")).startswith("0.1 DCF"):
            # 樣板 v0.5：標準版 DCF。折現率（WACC 等）直接顯示、緊接 DCF 彙總之前；彙總列由 Val 頁標籤找列號（不用 html_src 的舊列號）
            h = _find(V, "0.1 DCF") + 1
            ke, ronic, pvx, dcf = _find(V, "股權成本 ke"), _find(V, "永續期 RONIC"), _find(V, "明確期 FCFF 現值"), _find(V, "DCF 目標價（情境加權）")
            nb.append({"type": "table", "caption": "DCF 折現率（標準版 DCF：WACC、股權成本、稅後債務成本、RONIC、永續成長率）", "head": ["項目", "數值", "說明"],
                       "rows": [[f"@Val!A{r}", f"@Val!B{r}|p2", f"@Val!E{r}"] for r in range(ke, ronic + 1)]})
            pct = {str(_find(V, "  其中：終值現值占 EV")): {c: "p1" for c in "BCD"}, str(_find(V, "  隱含終值 EV／EBITDA")): {c: "x1" for c in "BCD"},
                   str(_find(V, "評價用稀釋股數")): {c: "" for c in "BCD"}}
            nb.append({"type": "table", "caption": f"@Val!A{h - 1}", "head": [f"@Val!{c}{h}" for c in "ABCDE"],
                       "rowgen": {"sheet": "Val", "rows": f"{pvx}-{dcf}", "cols": list("ABCDE"), "fmts": ["", pfmt, pfmt, pfmt, ""], "row_fmt": pct}})
            D = wb["DCF"]
            b0 = _find(D, "情境 B（")
            b1 = next(r for r in range(b0, D.max_row + 1) if str(D.cell(r, 1).value or "").startswith("檢查：終值占 EV"))
            hd = _find(D, "項目")
            nb.append({"type": "table", "collapse": True, "summary": "展開：情境 B 的無槓桿自由現金流明細（DCF 頁；情境 A、C 同式）", "caption": f"@DCF!A{b0}",
                       "head": [f"@DCF!{c}{hd}" for c in "ACDEFGH"],
                       "rowgen": {"sheet": "DCF", "rows": f"{b0 + 1}-{b1}", "cols": list("ACDEFGH"), "fmts": ["", "n1", "n1", "n1", "n1", "n1", "n1"],
                                  "row_fmt": {str(r): {c: ("p1" if "率" in str(D.cell(r, 1).value) or "÷" in str(D.cell(r, 1).value) or "占" in str(D.cell(r, 1).value) else ("n4" if "因子" in str(D.cell(r, 1).value) else "n1")) for c in "CDEFGH"}
                                              for r in range(b0 + 1, b1 + 1) if any(k in str(D.cell(r, 1).value or "") for k in ("率", "÷", "占", "因子"))}}})
            i += 1
            continue
        nb.append(b)
        i += 1
    secs["valuation"]["blocks"] = va = nb
    adj0, comp2 = _find(V, "Comparable 目標價（相對倍數調整後）"), _find(V, "Comparable 目標價（市場主流同業法）")
    for b in va:      # 長條圖的 Comparable 改為進加權目標價的相對倍數調整後值（與橫幅一致）
        if b["type"] == "bars" and str(b.get("title", "")).startswith("目標價與現價"):
            dcf_r = _find(V, "DCF 目標價（情境加權）")
            for it in b["items"]:
                if str(it.get("label", "")).startswith("@Val!A") and not txt(it["label"]):   # 樣板 v0.5：舊 0.1 節的 DCF 目標價列已移位
                    it["label"], it["value"] = f"@Val!A{dcf_r}", f"@Val!B{dcf_r}"
                if it.get("label") == f"@Val!A{comp2}":
                    it["label"], it["value"] = f"@Val!A{adj0}", f"@Val!B{adj0}"
    # 樣板 v0.5：目標價參數敏感度（標準版 DCF 的標準案例）；列號由 Sens 頁標題找（html_src 的舊列號停在 v0.4 的 tp_r／tp_g 案例）
    SE = wb["Sens"]
    t0 = _find(SE, "目標価參數敏感度".replace("価", "價"))
    t1 = _find(SE, "檢查：閉式公式的基準") - 1
    for b in secs["sens"]["blocks"]:
        if b["type"] == "h3" and str(b.get("text", "")).startswith("@Sens!A") and "目標價" in txt(b["text"]) or (b["type"] == "h3" and txt(b.get("text", "")).startswith("目標價參數")):
            b["text"] = f"@Sens!A{t0}"
        if b["type"] == "table" and b.get("rowgen", {}).get("sheet") == "Sens" and txt(b.get("head", [""])[0]) in ("案例", "參數"):
            b["head"] = [f"@Sens!{c}{t0 + 1}" for c in "ABCDEFG"]
            b["rowgen"] = {"sheet": "Sens", "rows": f"{t0 + 2}-{t1}", "cols": list("ABCDEFG"), "fmts": ["", "", pfmt, "sp1", pfmt, "sp1", pfmt]}
    for b in secs["recon"]["blocks"]:
        if b["type"] == "table" and not b.get("caption"):
            b["caption"] = "與共識、公司指引的對照（本模型正向推導，只作事後對照）"

    # ── 投資建議橫幅（記憶體專用）──
    rec, tpw, upw = _find(V, "投資建議"), _find(V, "目標價（DCF 與 Comparable 加權）"), _find(V, "上檔空間：加權目標價")
    adj, hup = _find(V, "Comparable 目標價（相對倍數調整後）"), _find(RD, "上檔空間：Comparable 目標價（相對倍數調整後）")
    hup2 = _find(RD, "上檔空間：只看 Comparable 目標價（市場主流同業法")
    px = ink["px"]
    hero = {"rec": f"@Val!B{rec}",
            "tiles": [[f"@Val!A{tpw}", f"@Val!B{tpw}|{pfmt}", ["上檔空間 ", f"@Val!B{upw}|sp1"]],
                      ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{comp2}|{pfmt}", ["上檔空間 ", f"@README!B{hup2}|sp1"]],
                      [f"@In!B{px}", f"@In!E{px}|{pfmt}", ""]],
            "note": ["投資建議依加權目標價判定：", f"@Val!E{rec}", "。進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{adj}|{pfmt}",
                     "（上檔空間 ", f"@README!B{hup}|sp1", "）。"]}
    RL._hero = lambda w: copy.deepcopy(hero)

    spec = spec_for(wb, nm_block)
    first = spec["in_groups"][0][0]
    intro = secs["drivers"].get("intro")
    spec["blocks_to"] = {intro: first} if intro else {}
    spec["main"] = {"0.3 結論": ["目標價（DCF 與 Comparable 加權）", "市場共識目標價", "上檔空間：加權目標價", "DCF ÷ 現價", "價格帶"],
                    "0.1 DCF": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
                    "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"]}
    if recon_main:
        spec["main"]["與共識、公司指引的對照"] = recon_main
    old_tornado = {json.dumps(b.get("rows")): b.get("title") for b in secs["sens"]["blocks"] if b["type"] == "tornado"}
    old_title = list(old_tornado.values())[0] if old_tornado else None

    cfg = RL.relayout(cfg, xlsx, spec)
    secs = {s["id"]: s for s in cfg["sections"]}

    # ── 轉換後 ──
    for b in secs["sens"]["blocks"]:
        if b["type"] == "tornado" and old_title is not None:
            t = old_title if isinstance(old_title, list) else [old_title]
            b["title"] = t + (["（變動未達兩位數百分比的案例）"] if b.get("collapse") else ["（變動達兩位數百分比的案例）"])
    va = secs["valuation"]["blocks"]
    cyc = next((b for b in va if b["type"] == "bars" and str(b.get("title", "")).startswith("三個循環位置")), None)
    if cyc:
        va.remove(cyc)
        j = next(k for k, b in enumerate(va) if b["type"] == "table" and str(b.get("caption", "")).startswith("估值規則：每股淨值"))
        va.insert(j + 1, cyc)
    st = secs["status"]["blocks"]
    for k, b in enumerate(st):
        if b["type"] == "h3" and b.get("text") == "附錄：版本要點":
            rows = [[f"@README!A{r}"] for r in range(2, RD.max_row + 1)
                    if re.match(r"^v+\d+\.\d|^\d{4}-\d{2}-\d{2}：", str(RD.cell(r, 1).value or ""))]
            st[k + 1]["rows"] = rows
    scen_r = _find(RD, "情境：A＝")
    db = secs["drivers"]["blocks"]
    if intro:
        nt = next((b for b in db if b["type"] == "note" and b.get("parts", [None])[0] == intro), None)
        if nt:
            db.remove(nt)
            db.insert(0, nt)
    db.insert(0, {"type": "note", "parts": ["情境定義：", f"@README!A{scen_r}"]})
    json.dump(cfg, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("ok", out)

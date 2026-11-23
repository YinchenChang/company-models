# -*- coding: utf-8 -*-
"""由成品 xlsx 產生 physical_ai/fanuc/html.json（個股 HTML 版面規格 v1.3）。列號依 A 欄標籤尋找，勿手改 html.json。

用法：python3 physical_ai/fanuc/gen_html_json.py            # 讀 dist/ 內唯一 xlsx，寫 html.json
之後：python3 tools/company_html/build_company_html.py physical_ai/fanuc
"""
import glob
import json
import os

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
X = sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))
assert len(X) == 1, X
wb = openpyxl.load_workbook(X[0], data_only=True)


def find(sheet, label, col=1, start=1, prefix=False):
    ws = wb[sheet]
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and (v.startswith(label) if prefix else v == label):
            return r
    raise KeyError(f"{sheet}: {label}")


# ---- PL 情境區塊
PL = {}
for s in "ABC":
    h = find("PL", f"情境 {s}", prefix=True)
    PL[s] = {"hdr": h}
    for k, lab in (("o_fa", "FA 訂單"), ("o_rb", "Robot 訂單"), ("o_rm", "Robomachine 訂單"), ("o_sum", "訂單合計"),
                   ("bb_fa", "FA 訂單÷營收"), ("r_fa", "FA 營收"), ("r_rb", "Robot 營收"), ("r_rm", "Robomachine 營收"),
                   ("r_sv", "Service 營收"), ("rev", "合併營收"), ("chk", "檢查：分部合計"), ("g", "合併營收年增"),
                   ("oi", "營業利益"), ("opm", "營業利益率"), ("eq", "持分法投資利益"), ("nop", "其他營業外收支"),
                   ("ord", "經常利益"), ("tax", "法人稅等"), ("nci", "非支配股東利益"), ("ni", "歸屬母公司淨利"),
                   ("eps", "EPS"), ("dps", "每股股利"), ("fcf", "自由現金流（淨利"), ("fcfd", "自由現金流（扣股利）")):
        PL[s][k] = find("PL", lab, start=h, prefix=True)
B = PL["B"]
HB = B["hdr"]
YR = f"@PL!C{HB}:K{HB}"          # FY2024–FY2032
YRF = f"@PL!D{HB}:K{HB}"         # FY2025–FY2032（訂單自 FY2025 起）

# ---- README 展示層
RD = {wb["README"].cell(r, 1).value: r for r in range(1, wb["README"].max_row + 1) if wb["README"].cell(r, 1).value}
rd = lambda k: f"@README!B{RD[k]}"

# ---- Val
VA = lambda lab: find("Val", lab, prefix=True)
v_tp, v_px, v_up, v_rec, v_tier, v_gate, v_upc = VA("加權目標價"), VA("現價"), VA("上檔空間"), VA("投資建議"), VA("估值評等"), VA("入不敷出"), VA("Comparable 目標價上檔空間")
v_dist = VA("上檔空間距最近")

# ---- Sens
SE = wb["Sens"]
case_rows = [r for r in range(6, SE.max_row + 1) if isinstance(SE.cell(r, 6).value, (int, float)) and SE.cell(r, 3).value not in (None, "—")]
base_row = find("Sens", "基準", start=5)
sens_blocks = []
for side in ("低", "高"):
    rs = [r for r in case_rows if str(SE.cell(r, 1).value).endswith(f"敏感度{side}")]
    big = [r for r in rs if abs(SE.cell(r, 6).value) >= 0.1]
    small = [r for r in rs if abs(SE.cell(r, 6).value) < 0.1 and abs(SE.cell(r, 6).value) >= 0.0005]
    base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": "F", "fmt": "sp1"}
    if big:
        title = f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例"
        main_rows = big
    else:
        top = max(rs, key=lambda r: abs(SE.cell(r, 6).value))
        title = f"加權目標價相對基準：輸入取{side}值，變動最大的案例（沒有達兩位數百分比者）"
        main_rows = [top]
        small = [r for r in small if r != top]
    sens_blocks.append(dict(base, title=title, rows=",".join(map(str, main_rows))))
    if small:
        sens_blocks.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，其餘案例", rows=",".join(map(str, small)),
                                collapse=True, summary=f"展開：輸入取{side}值，變動未達兩位數百分比的案例"))
mains = [b for b in sens_blocks if not b.get("collapse")]
rest = [b for b in sens_blocks if b.get("collapse")]
sens = {"id": "sens", "title": "敏感度（一次動一個輸入，整本重算）", "blocks": mains + [
    {"type": "note", "parts": ["@Sens!A1"]}] + rest + [
    {"type": "table", "collapse": True, "summary": "展開：全部案例（DCF 目標價、加權目標價、上檔空間與評等）",
     "head": ["@Sens!A5", "@Sens!B5", "@Sens!D5", "@Sens!E5", "@Sens!F5", "@Sens!G5", "@Sens!H5"],
     "rowgen": {"sheet": "Sens", "rows": f"{base_row}-{max(case_rows)}", "cols": ["A", "B", "D", "E", "F", "G", "H"], "fmts": ["", "", "n0", "n0", "sp1", "sp1", ""]}},
    {"type": "kv", "items": [["@Sens!A3", "@Sens!D3"]]},
    {"type": "list", "collapse": True, "summary": "展開：未列入敏感度的輸入", "items": [f"@Sens!A{SE.max_row}"]}]}

cfg = {
    "company": "Fanuc（ファナック）", "ticker": "6954.T", "layout": "v1.3", "wide_text": True,
    "thesis": rd("命題"),
    "question": "這波以 FA（CNC）為主的訂單會不會在 FY2027 轉成營收與利潤？循環回落時，現價隱含了什麼？",
    "scenarios": [{"key": "A", "name": "情境 A（樂觀）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（保守）"}],
    "hero": {"rec": f"@Val!B{v_rec}",
             "tiles": [["加權目標價（DCF 與 Comparable 加權）", f"@Val!B{v_tp}|n0", ["上檔空間 ", f"@Val!B{v_up}|sp1"]],
                       ["Comparable 目標價（同業本益比，情境加權）", "@Val!E21|n0", ["上檔空間 ", f"@Val!B{v_upc}|sp1"]],
                       [f"@Val!A{v_px}", f"@Val!B{v_px}|n0", ""]],
             "note": [rd("展示層：橫幅說明")]},
    "sections": [],
}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "FA 訂單（十億日圓）× 三情境：FY2025–FY2032", "x": YRF, "fmt": "n0", "stack": False, "pad": 1.15,
     "bars": [{"name": [f"@PL!A{PL[s]['hdr']}"], "y": f"@PL!D{PL[s]['o_fa']}:K{PL[s]['o_fa']}", "color": c}
              for s, c in (("A", "#7fb8bb"), ("B", "#0b7a80"), ("C", "#d9a08f"))],
     "lines": [{"name": ["情境 B：", f"@PL!A{B['r_fa']}"], "y": f"@PL!D{B['r_fa']}:K{B['r_fa']}", "color": "#3b3838", "labels": False}],
     "caption": "長條＝三情境 FA 訂單；折線＝情境 B 的 FA 營收（同一左軸）。FY2025 為實績、FY2026 為上半實績加下半假設；之後三情境在 FA 訂單回落幅度分化，是本模型最承重的變數。"},
    {"type": "h3", "text": "命題"},
    {"type": "list", "ordered": True, "items": [rd("展示層：命題分項 1"), rd("展示層：命題分項 2"), rd("展示層：命題分項 3")]},
    {"type": "list", "collapse": True, "summary": "展開：命題回覆（依模型計算）", "items": [rd("展示層：命題回覆")]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [rd(f"展示層：重點 {i}") for i in "12345"]},
    {"type": "h3", "text": "失效條件"},
    {"type": "note", "parts": [rd("失效條件")]}]}

valuation = {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "目標價與現價（円；豎線＝現價）", "fmt": "n0",
     "items": [{"label": "DCF 目標價（情境加權）", "value": "@Val!E16", "series": 1},
               {"label": "Comparable 目標價（未調整）", "value": "@Val!E20", "series": 2},
               {"label": "Comparable 目標價（相對倍數調整後，進加權）", "value": "@Val!E21", "series": 2},
               {"label": "加權目標價（DCF 與 Comparable 加權）", "value": f"@Val!B{v_tp}", "series": 0}],
     "ref": {"label": f"@Val!A{v_px}", "value": f"@Val!B{v_px}"}},
    {"type": "note", "parts": ["@Val!A1"]},
    {"type": "note", "parts": ["市場共識目標價：不適用（本輪未取得可引用的共識來源；Recon 頁：", "@Recon!B14", "）。"]},
    {"type": "table", "caption": "結論", "head": ["項目", "數值"],
     "rows": [["DCF 目標價（情境加權）", "@Val!E16|n0"], ["Comparable 目標價（相對倍數調整後）", "@Val!E21|n0"],
              ["加權目標價（DCF 與 Comparable 加權）", f"@Val!B{v_tp}|n0"], [f"@Val!A{v_px}", f"@Val!B{v_px}|n0"], [f"@Val!A{v_up}", f"@Val!B{v_up}|sp1"],
              [f"@Val!A{v_upc}", f"@Val!B{v_upc}|sp1"], [f"@Val!A{v_tier}", f"@Val!B{v_tier}"], [f"@Val!A{v_dist}", f"@Val!B{v_dist}|p1"],
              [f"@Val!A{v_gate}", f"@Val!B{v_gate}|n1"], ["收支關卡判定", f"@Val!C{v_gate}"], [f"@Val!A{v_rec}", f"@Val!B{v_rec}"]]},
    {"type": "list", "collapse": True, "summary": "展開：上檔空間三段歸因（公式組字）", "items": [rd("三段歸因")]},
    {"type": "h3", "text": "DCF 折現率"},
    {"type": "kv", "items": [["@Val!A4", "@Val!B4|p2"], ["WACC（無有息負債，等於股權成本）", "@Val!B5|p2"], ["@Val!A6", "@Val!B6|p2"], ["@Val!A7", "@Val!B7|p2"],
                             ["@In!B70", "@In!E70|p3"], ["@In!B73", "@In!E73|n2"], ["@In!B72", "@In!E72|p1"], ["@In!B71", "@In!E71|p2"]]},
    {"type": "table", "caption": "DCF 三情境彙總（十億日圓；每股為円）", "head": ["@Val!A11", "@Val!B11", "@Val!C11", "@Val!D11", "@Val!E11"],
     "rowgen": {"sheet": "Val", "rows": "12-16", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1", "n0"],
                "row_fmt": {"13": {c: "p1" for c in "BCD"}, "15": {c: "n0" for c in "BCD"}, "16": {c: "n0" for c in "BCDE"}}}},
    {"type": "table", "collapse": True, "summary": "展開：Comparable 三情境（目標時點 NTM EPS × 同業本益比中位數）", "head": ["@Val!A18", "@Val!B18", "@Val!C18", "@Val!D18", "@Val!E18"],
     "rowgen": {"sheet": "Val", "rows": "19-21", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1", "n1"]}},
    {"type": "table", "collapse": True, "summary": "展開：同業本益比（現價日收盤；倍數按收盤比例換算）", "head": ["@Val!A39", "@Val!B39", "@Val!C39", "@Val!D39", "@Val!E39"],
     "rowgen": {"sheet": "Val", "rows": "40-47", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "", "n1", "n2", "tag"]}},
    {"type": "note", "parts": ["@Val!A38", "；換算式見 xlsx Val 頁 F 欄（含 Keyence 分割調整）。"]},
    {"type": "note", "parts": ["@Val!A48"]},
    {"type": "table", "collapse": True, "summary": "展開：上檔空間三段歸因", "head": ["項目", "數值"],
     "rowgen": {"sheet": "Val", "rows": "34-36", "cols": ["A", "B"], "fmts": ["", "sp1"]}}]}

revdcf = {"id": "revdcf", "title": "反向檢驗：現價隱含什麼（訂單年增反解）", "blocks": [
    {"type": "note", "parts": ["本模型未建反向 DCF 頁；以 Recon 頁整本重算的反解代替：三部門 FY2027–FY2029 訂單年增在三情境同時平移多少，加權目標價才等於現價或跨過評等門檻。"]},
    {"type": "table", "caption": "反解（整本重算）：三部門 FY2027–FY2029 訂單年增在三情境同加的平移", "head": ["@Recon!A17", "所需平移（三情境同加）", "@Recon!C17"],
     "rowgen": {"sheet": "Recon", "rows": "18-20", "cols": ["A", "B", "C"], "fmts": ["", "sp1", ""]}},
    {"type": "kv", "items": [["@Recon!A22", "@Recon!B22|sp0"]]},
    {"type": "note", "parts": [rd("展示層：重點 4")]}]}


def std_row(k, f):
    r = B[k]
    lab = LAB.get(k, f"@PL!A{r}")
    return [lab] + [f"@PL!{c}{r}|{f}" for c in "CDEFGHIJK"]


# 展示用列名（白話，對應 PL 頁同列；xlsx 原標籤：合併營收、法人稅等、經常利益）
LAB = {"rev": "營收（合併）", "tax": "所得稅（法人稅等）", "ord": "經常利益（模型以此為稅前基礎；未計特別損益）"}


pl = {"id": "pl", "title": "損益（日本會計準則）", "blocks": [
    {"type": "theme", "title": "合併營收（十億日圓，依部門堆疊）與營業利益率：FY2024–FY2032（情境 B）", "x": YR, "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.25,
     "total": f"@PL!C{B['rev']}:K{B['rev']}",
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": f"@PL!C{B[k]}:K{B[k]}", "color": c}
              for k, c in (("r_fa", "#0b7a80"), ("r_rb", "#6fc7cb"), ("r_rm", "#c29a2b"), ("r_sv", "#8a8c8c"))],
     "lines": [{"name": [f"@PL!A{B['opm']}"], "y": f"@PL!C{B['opm']}:K{B['opm']}", "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝公司固定揭露的四個部門營收（官方報告分部為單一分部，四部門合計等於合併營收）；折線＝營業利益率（右軸）。FY2024、FY2025 為實績。"},
    {"type": "note", "parts": ["口徑：日本會計準則（J-GAAP，非 IFRS），公司決算短信與市場都以此口徑；公司不揭露營業成本以下的部門別利益，本模型以合併營業利益起算（增量利潤率法），故無營業成本、毛利與營業費用列。歷史期只有 FY2024、FY2025 兩年。"]},
    {"type": "table", "caption": "損益表（J-GAAP；十億日圓，EPS 與每股股利為円）：FY2024–FY2025 實績、FY2026 起情境 B",
     "head": [f"@PL!A{HB}"] + [f"@PL!{c}{HB}" for c in "CDEFGHIJK"],
     "rows": [std_row("rev", "n1"), std_row("g", "sp1"), std_row("oi", "n1"), std_row("opm", "p1"), std_row("eq", "n1"), std_row("nop", "n1"),
              std_row("ord", "n1"), std_row("tax", "n1"), std_row("nci", "n1"), std_row("ni", "n1"), std_row("eps", "n1"), std_row("dps", "n1")]},
    {"type": "pl", "collapse": True, "summary": "展開：三情境損益（A／B／C 並列，含折線）", "years": f"@PL!C{HB}:K{HB}",
     "metrics": [{"name": "合併營收（十億日圓）", "fmt": "n0", "chart": True, "rows": {s: f"@PL!C{PL[s]['rev']}:K{PL[s]['rev']}" for s in "ABC"}},
                 {"name": "營業利益（十億日圓）", "fmt": "n0", "chart": True, "rows": {s: f"@PL!C{PL[s]['oi']}:K{PL[s]['oi']}" for s in "ABC"}},
                 {"name": "營業利益率", "fmt": "p1", "rows": {s: f"@PL!C{PL[s]['opm']}:K{PL[s]['opm']}" for s in "ABC"}},
                 {"name": "EPS（円）", "fmt": "n1", "chart": True, "rows": {s: f"@PL!C{PL[s]['eps']}:K{PL[s]['eps']}" for s in "ABC"}},
                 {"name": "自由現金流（扣股利，十億日圓）", "fmt": "n1", "rows": {s: f"@PL!C{PL[s]['fcfd']}:K{PL[s]['fcfd']}" for s in "ABC"}}]},
    {"type": "table", "collapse": True, "summary": "展開：現金流（情境 B）", "head": [f"@PL!A{HB}"] + [f"@PL!{c}{HB}" for c in "CDEFGHIJK"],
     "rowgen": {"sheet": "PL", "rows": f"{B['fcf'] - 3}-{B['fcfd']}", "cols": ["A"] + list("CDEFGHIJK"), "fmts": [""] + ["n1"] * 9}}]}

IN = wb["In"]
in_rows = lambda a, b: [r for r in range(a, b + 1) if IN.cell(r, 1).value and IN.cell(r, 3).value is not None]
sec_r = lambda t: find("In", t, prefix=True)
rB, rC, rD, rE = sec_r("B. "), sec_r("C. "), sec_r("D. "), sec_r("E. ")
rng = lambda a, b: [r for r in in_rows(a, b) if IN.cell(r, 4).value is not None or IN.cell(r, 6).value is not None]
fixed = lambda a, b: [r for r in in_rows(a, b) if IN.cell(r, 4).value is None and IN.cell(r, 6).value is None]


def in_table(rows, cap, collapse=False, summary=None):
    def f_of(r):
        u = IN.cell(r, 3).value
        return "p1" if u in ("%", "pt") else "n2"
    t = {"type": "table", "caption": cap, "head": ["@In!B3", "@In!C3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
         "rows": [[f"@In!B{r}", f"@In!C{r}", f"@In!D{r}|{f_of(r)}", f"@In!E{r}|{f_of(r)}", f"@In!F{r}|{f_of(r)}", f"@In!G{r}|tag"] for r in rows]}
    if collapse:
        t.update(collapse=True, summary=summary)
    return t


scn0 = find("In", "scn_o_fa27")
drivers = {"id": "drivers", "title": "驅動因子：訂單 → 部門營收 → 合併營收", "blocks": [
    {"type": "theme", "title": "三部門訂單（十億日圓，堆疊）與 FA 訂單 ÷ 營收：FY2025–FY2032（情境 B）", "x": YRF, "fmt": "n0", "fmt_r": "n2", "stack": True, "pad": 1.25,
     "total": f"@PL!D{B['o_sum']}:K{B['o_sum']}",
     "bars": [{"name": [f"@PL!A{B[k]}"], "y": f"@PL!D{B[k]}:K{B[k]}", "color": c} for k, c in (("o_fa", "#0b7a80"), ("o_rb", "#6fc7cb"), ("o_rm", "#c29a2b"))],
     "lines": [{"name": [f"@PL!A{B['bb_fa']}"], "y": f"@PL!D{B['bb_fa']}:K{B['bb_fa']}", "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "部門營收＝（一減傳導比重）× 當年訂單＋傳導比重 × 上年訂單；FA 訂單 ÷ 營收在 FY2026 高於一（積壓），之後隨訂單回落降到一附近。"},
    {"type": "note", "parts": ["@PL!A1"]},
    {"type": "table", "caption": "分部營收推導（情境 B；十億日圓）", "head": [f"@PL!A{HB}"] + [f"@PL!{c}{HB}" for c in "CDEFGHIJK"],
     "rowgen": {"sheet": "PL", "rows": f"{B['o_fa']}-{B['g']}", "cols": ["A"] + list("CDEFGHIJK"), "fmts": [""] + ["n1"] * 9,
                "row_fmt": {**{str(B[k]): {c: "n2" for c in "CDEFGHIJK"} for k in ("bb_fa",)},
                            **{str(B["bb_fa"] + i): {c: "n2" for c in "CDEFGHIJK"} for i in (1, 2)},
                            str(B["g"]): {c: "sp1" for c in "CDEFGHIJK"}, str(B["chk"]): {c: "n3" for c in "CDEFGHIJK"}}}},
    {"type": "table", "caption": "@In!A85", "head": ["@In!B86", "@In!D86", "@In!E86", "@In!F86"],
     "rows": [[f"@In!B{r}", f"@In!D{r}|p1", f"@In!E{r}|p1", f"@In!F{r}|p1"] for r in range(scn0, IN.max_row + 1) if IN.cell(r, 2).value]},
    {"type": "h3", "text": "一、FY2026 下半（部門營收與訂單）"},
    in_table(rng(rB, rC), "@In!A" + str(rB)),
    {"type": "h3", "text": "二、訂單 → 營收傳導"},
    in_table(rng(rC, rD), "@In!A" + str(rC)),
    in_table(fixed(rC, rD), "情境專屬輸入（三情境即區間）", True, "展開：三情境訂單年增與長期成長（逐項）"),
    {"type": "h3", "text": "三、損益與現金流"},
    in_table(rng(rD, rE), "@In!A" + str(rD)),
    in_table(fixed(rD, rE), "固定輸入（無敏感度區間）", True, "展開：損益固定輸入"),
    {"type": "h3", "text": "四、全公司（折現率與估值）"},
    in_table(rng(rE, IN.max_row), "@In!A" + str(rE)),
    in_table(fixed(rE, find("In", "scw_C")) + fixed(4, rB), "固定輸入（價格、股數、日期、折現率與權重）", True, "展開：全公司固定輸入")]}

recon = {"id": "recon", "title": "對帳：公司指引與線層級變數", "blocks": [
    {"type": "note", "parts": ["@Recon!A1"]},
    {"type": "table", "caption": "FY2026 公司指引對模型情境 B（公司指引＝Interested-party）", "head": ["@Recon!A3", "@Recon!B3", "@Recon!C3", "@Recon!D3", "@Recon!E3"],
     "rowgen": {"sheet": "Recon", "rows": "4-11", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1", "sp1"],
                "row_fmt": {"8": {"B": "n2", "C": "n2", "D": "n2"}, "10": {"B": "p1", "C": "p1", "D": "sp1"}}}},
    {"type": "note", "parts": ["@Recon!B13"]},
    {"type": "kv", "items": [["@Recon!A14", "@Recon!B14"]]},
    {"type": "table", "collapse": True, "summary": "展開：線層級變數對帳（工業機器人與人形零件加工）", "head": ["@對帳!A3", "@對帳!B3", "@對帳!C3", "@對帳!E3", "@對帳!G3", "@對帳!H3"],
     "rowgen": {"sheet": "對帳", "rows": f"4-{wb['對帳'].max_row}", "cols": ["A", "B", "C", "E", "G", "H"], "fmts": ["", "", "", "", "tag", ""]}}]}

status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["build_id", "@Sens!D2"],
                             ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"]]},
    {"type": "h3", "text": "限制與協調檔"},
    {"type": "list", "items": [rd("限制"), rd("協調檔"), "@In!I70", rd("可證偽指標")]},
    {"type": "table", "caption": "風險登錄", "head": ["@Risk!A3", "@Risk!B3", "@Risk!C3", "@Risk!D3"],
     "rowgen": {"sheet": "Risk", "rows": f"4-{wb['Risk'].max_row}", "cols": ["A", "B", "C", "D"]}},
    {"type": "note", "parts": ["來源全文與網址見 xlsx「Sources」頁（本頁自含、不放外部網址）。"]},
    {"type": "h3", "text": "「未取得」清單"},
    {"type": "scan", "needle": "未取得", "collapse": True, "summary": "展開：標為「未取得」的項目", "sheets": {"In": ["B"], "Val": ["A"], "DCF": ["A"]}, "empty": "沒有標為未取得的項目。"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A"]}, "empty": "沒有標為未找到的項目。"}]}

cfg["sections"] = [thesis, valuation, sens, revdcf, pl, drivers, recon, status]
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", os.path.basename(X[0]))

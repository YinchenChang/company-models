# 由 build_nri.py 的列號對照（rowmap.json）產生 html.json（列號隨模型變動，勿手改 html.json；文字不含手打數字，數字一律 @工作表!儲存格）。
# 用法：python3 gen_html_json.py <成品 xlsx> <rowmap.json> html.json
import json, sys, openpyxl
X, RM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
rm = json.load(open(RM, encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)
IR, TR, SC, PR, IS, VR, FC, RVR, RCR, CR, TKR = (rm[k] for k in ("IR", "TR", "SC", "PR", "IS", "VR", "FC", "RVR", "RCR", "CR", "TKR"))
B = PR["B"]; EB = SC["B"]; CB = CR["B"]; PH = B["hdr"]; TH_H = EB["auto"] - 1; CF_H = CB["ni"] - 1
def yrs(sh, r, c1, c2): return f"@{sh}!{c1}{r}:{c2}{r}"
tornL, tornH = [], []
for r in rm["SE_ROWS"]:
    side = wb["Sens"].cell(r, 3).value; v = wb["Sens"].cell(r, 5).value
    if isinstance(v, (int, float)) and abs(v) >= 0.001: (tornL if side == "低" else tornH).append(r)
ws_in = wb["In"]
scn_h = next(r for r in range(1, ws_in.max_row + 1) if str(ws_in.cell(r, 1).value or "").startswith("情境驅動對照")) + 1
scn_end = scn_h
while ws_in.cell(scn_end + 1, 1).value: scn_end += 1
# In 單一參數區間表：排除情境專屬列
scn_keys = [k for k in IR if k.startswith(("auto_", "ret_", "proj_", "tp_w"))]
skip = {IR[k] for k in scn_keys}
last_in = max(IR.values()); rngs = []; a = None
for r in range(4, last_in + 1):
    if r in skip:
        if a is not None: rngs.append((a, r - 1)); a = None
    elif a is None: a = r
if a is not None: rngs.append((a, last_in))
in_rows = ",".join(f"{x}-{y}" if x != y else str(x) for x, y in rngs)
pl_cols = list("CDEFGHIJKLM"); is_cols = list("CDEFGHIJK")
segrows = [B["seg"][s] for s in B["seg"]]
cfg = {
 "company": "野村總合研究所 NRI", "ticker": "4307", "thesis": "@README!A5",
 "question": "AI 進入 IT 服務後，AI 新專案與省下的人力成本能否抵過客戶要求分享生產力的價格侵蝕與推論成本",
 "scenarios": [{"key": "A", "name": "情境 A（價格保留高、AI 新專案佳）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（價格保留低、AI 新專案少）"}],
 "wide_text": True, "layout": "v1.3",
 "hero": {"rec": f"@Val!B{VR['rec']}", "tiles": [[f"@Val!A{VR['tp']}", f"@Val!B{VR['tp']}|n2", ["上檔空間 ", f"@Val!B{VR['up']}|sp1"]], [f"@Val!A{VR['tpcomp']}", f"@Val!B{VR['tpcomp']}|n2", ["上檔空間 ", f"@Val!B{VR['upc']}|sp1"]], [f"@Val!A{VR['px']}", f"@Val!B{VR['px']}|n0", ""]],
          "note": ["@README!A13"]},
 "sections": [
  {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "note", "parts": ["@README!A5"]},
    {"type": "theme", "title": "營收（十億日圓，依官方分部堆疊）：FY2027E–FY2032E（情境 B）", "x": yrs("PL", PH, "H", "M"), "fmt": "n0", "stack": True, "pad": 1.15, "total": yrs("PL", B["rev"], "H", "M"),
     "bars": [{"name": [f"@PL!A{B['seg'][s]}"], "y": yrs("PL", B["seg"][s], "H", "M"), "color": c} for s, c in (("Consulting", "#c29a2b"), ("Financial IT Solutions", "#6fa8dc"), ("Industrial IT Solutions", "#0b7a80"), ("IT Platform Services", "#8e7cc3"), ("Others／Adjustments", "#999999"))],
     "caption": "堆疊＝官方報告分部外部客戶營收（新口徑）；單位十億日圓。"},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [f"@README!A{r}" for r in range(6, 11)]},
    {"type": "h3", "text": "本版要點"},
    {"type": "list", "items": [f"@README!A{r}" for r in range(15, 19)]}]},
  {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "四個目標價與現價（日圓；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{VR[k]}", "value": f"@Val!B{VR[k]}", "series": sr} for k, sr in (("tpdcf", 1), ("tpcomp", 2), ("tp", 0), ("cons", 3))],
     "ref": {"label": f"@Val!A{VR['px']}", "value": f"@Val!B{VR['px']}"}},
    {"type": "table", "caption": "結論（情境加權）", "head": ["項目", "數值"], "rowgen": {"sheet": "Val", "rows": f"{VR['px']}-{VR['rec']}", "cols": ["A", "B"], "fmts": ["", "n2"], "row_fmt": {str(VR["om"]): {"B": "n3"}, **{str(VR[k]): {"B": "n2"} for k in ("ntmA", "ntmB", "ntmC")}, str(VR["f_ttm"]): {"B": "n3"}, **{str(VR[k]): {"B": "n1"} for k in ("ttmA", "ttmB", "ttmC", "sh_t")}, **{str(VR[k]): {"B": "sp1"} for k in ("up", "upd", "upc")}, str(VR["tier"]): {"B": ""}, str(VR["gate"]): {"B": "n1"}, str(VR["gate_t"]): {"B": ""}, str(VR["rec"]): {"B": ""}}}},
    {"type": "note", "parts": ["@README!A19"]},
    {"type": "table", "caption": "DCF（股票薪酬加回）對照（只作對照，不進加權目標價；十億日圓除另註）", "head": ["項目", "情境 A", "情境 B", "情境 C", "情境加權"], "rowgen": {"sheet": "Recon", "rows": f"{RCR['sbc_a']}-{RCR['sbc_a']+RCR['sbc_n']-1}", "cols": list("ABCDE"), "fmts": ["", "n2", "n2", "n2", "n2"], "row_fmt": {str(RCR["sbc_a"]): {c: "p1" for c in "BCD"}, str(RCR["sbc_a"] + 1): {c: "n1" for c in "BCD"}}}},
    {"type": "kv", "items": [[f"@Val!A{VR[k]}", f"@Val!B{VR[k]}|{f}"] for k, f in (("ke", "p2"), ("kdat", "p2"), ("wacc", "p2"), ("ronic", "p2"), ("g", "p2"))]},
    {"type": "table", "caption": "DCF 彙總（三情境，十億日圓除另註）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rows": [[f"@Val!A{FC['B'][k]}", f"@Val!B{FC['A'][k]}|{f}", f"@Val!B{FC['B'][k]}|{f}", f"@Val!B{FC['C'][k]}|{f}"] for k, f in (("ev", "n1"), ("tvs", "p1"), ("eq", "n1"), ("v0", "n0"), ("vt", "n0"))]},
    {"type": "table", "caption": "落差分解與同業倍數", "head": ["項目", "數值", "對照"], "rowgen": {"sheet": "Val", "rows": f"{VR['pe_now']}-{VR['wdcf_sh']}", "cols": ["A", "B", "C"], "fmts": ["", "n1", "n1"], "row_fmt": {str(VR["pe_now"]): {"B": "n2"}, str(VR["impl_pe"]): {"B": "n1"}, str(VR["ps_eff"]): {"B": "n2"}, str(VR["impl_ps"]): {"B": "n2"}, str(VR["pe_ref1"]): {"B": "n0"}, str(VR["pe_ref2"]): {"B": "n0"}, str(VR["gap_dcf_cmp"]): {"B": "n0"}, str(VR["gap_pe"]): {"B": "n0"}, str(VR["wdcf_sh"]): {"B": "n0", "C": "n0"}}}},
    {"type": "table", "collapse": True, "summary": "展開：同業 P/S（Comparable 主腿；現價日口徑）", "head": ["同業", "價格日", "P/S", "價格日（日期）", "逾期"], "rowgen": {"sheet": "Val", "rows": f"{VR['ps_a']}-{VR['ps_b']}", "cols": list("ABCEF"), "fmts": ["", "", "n2", "", ""]}},
    {"type": "table", "collapse": True, "summary": "展開：同業前瞻 P/E（對照列，不進加權）（逐家標會計準則、資料日與逾期）", "head": ["同業", "代號", "前瞻 P/E", "EV/EBITDA", "資料日", "逾期", "會計準則", "分母期間與租賃處理"], "rowgen": {"sheet": "Val", "rows": f"{rm['PE'][0]}-{rm['PE'][1]}", "cols": list("ABCDEFGH"), "fmts": ["", "", "n2", "n2", "", "", "", ""]}}]},
  {"id": "sens", "title": "敏感度", "blocks": [
    {"type": "tornado", "title": "加權目標價相對基準：輸入取低值（一次改一個，整本重算）", "sheet": "Sens", "rows": ",".join(map(str, tornL)), "label_col": "A", "value_col": "E", "fmt": "sp1"},
    {"type": "tornado", "title": "加權目標價相對基準：輸入取高值（一次改一個，整本重算）", "sheet": "Sens", "rows": ",".join(map(str, tornH)), "label_col": "A", "value_col": "E", "fmt": "sp1"},
    {"type": "note", "parts": ["@Sens!A1"]}, {"type": "note", "parts": ["@Sens!A2"]},
    {"type": "table", "collapse": True, "summary": "展開：全部案例（含 DCF 單法與投資建議）", "head": ["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "DCF 目標價", "DCF 變動（日圓）", "投資建議"],
     "rowgen": {"sheet": "Sens", "rows": f"{rm['HDR_R']+1}-{rm['SE_END']}", "cols": list("ABCDEFGHI"), "fmts": ["", "", "", "n0", "sp1", "", "n0", "n0", ""]}},
    {"type": "kv", "items": [["build_id", "@Sens!D3"], ["模型狀態", "@Sens!D4"], ["@Sens!A5", "@Sens!D5|n0"], ["@Sens!A6", "@Sens!D6|n0"]]}]},
  {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "blocks": [
    {"type": "note", "parts": ["@RevDCF!A1"]}, {"type": "note", "parts": ["@RevDCF!A2"]},
    {"type": "table", "caption": "現價隱含值與模型基準（情境 B）", "head": ["項目", "數值"], "rowgen": {"sheet": "RevDCF", "rows": f"{RVR['pn']}-{RVR['pe3']}", "cols": ["A", "B"], "fmts": ["", "n2"],
      "row_fmt": {str(RVR[k]): {"B": f} for k, f in (("pn", "n1"), ("po", "n1"), ("nc", "n1"), ("v0", "n0"), ("lam", "n3"), ("m0", "p1"), ("m1", "p1"), ("oi0", "n1"), ("oi1", "n1"), ("pe1", "n1"), ("pe2", "n1"), ("pe3", "n2"))}}},
    {"type": "note", "parts": ["@RevDCF!A20"]}, {"type": "note", "parts": ["@RevDCF!A21"]}]},
  {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "note", "parts": ["@PL!A1"]}, {"type": "note", "parts": ["@PL!A2"]},
    {"type": "theme", "title": "營收（十億日圓）與營業利益率：FY2024A–FY2032E（情境 B）", "x": yrs("PL", IS["hdr"], "C", "K"), "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.3,
     "bars": [{"name": [f"@PL!A{IS['rev']}"], "y": yrs("PL", IS["rev"], "C", "K"), "color": "#6fa8dc"}],
     "lines": [{"name": [f"@PL!A{IS['opm']}"], "y": yrs("PL", IS["opm"], "C", "K"), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "長條＝營收（十億日圓）；折線＝營業利益率（右軸）。FY2026A 營業利益含減損；FY2027E 起為情境 B。"},
    {"type": "table", "caption": "標準損益表（IFRS；十億日圓，EPS 為日圓）：FY2024A–FY2026A 實績、FY2027E–FY2032E 情境 B", "head": ["項目"] + [f"@PL!{c}{IS['hdr']}" for c in is_cols],
     "rowgen": {"sheet": "PL", "rows": f"{IS['rev']}-{IS['r_niy']}", "cols": ["A"] + is_cols, "fmts": [""] + ["n1"] * 9,
                "row_fmt": {**{str(IS[n]): {c: "p1" for c in is_cols} for n in ("yoy", "gm", "opm", "nm", "r_yoy", "r_gm", "r_opm", "r_nm", "r_niy")}, str(IS["eps"]): {c: "n2" for c in is_cols}}}},
    {"type": "theme", "title": "自由現金流：情境 A／B／C（十億日圓）", "x": yrs("CF", CF_H, "H", "M"), "fmt": "n0", "stack": False, "pad": 1.2,
     "bars": [{"name": ["情境 ", s, "：自由現金流"], "y": yrs("CF", CR[s]["fcf"], "H", "M"), "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
     "caption": "三情境的自由現金流（當期利益＋折舊攤銷−資本支出−營運資金，租賃作營業成本口徑）；情境差異來自價格保留率、AI 新專案與自動化。"},
    {"type": "table", "caption": "分部營收與產品線（情境 B；十億日圓）", "head": ["項目"] + [f"@PL!{c}{PH}" for c in pl_cols],
     "rowgen": {"sheet": "PL", "rows": f"{min(segrows)}-{B['yoy']}", "cols": ["A"] + pl_cols, "fmts": [""] + ["n1"] * 11, "row_fmt": {str(B["yoy"]): {c: "p1" for c in pl_cols}, str(B["chk"]): {c: "n3" for c in pl_cols}, str(B["chk2"]): {c: "n3" for c in pl_cols}}}},
    {"type": "table", "caption": "分部營業利益（情境 B；十億日圓；歷史為舊口徑，FY2026A Industrial IT 含減損）", "head": ["項目"] + [f"@PL!{c}{PH}" for c in pl_cols],
     "rowgen": {"sheet": "PL", "rows": f"{min(B['oi_s'].values())}-{B['opm']}", "cols": ["A"] + pl_cols, "fmts": [""] + ["n1"] * 11, "row_fmt": {str(B["opm"]): {c: "p1" for c in pl_cols}, str(B["chk3"]): {c: "n3" for c in pl_cols}}}},
    {"type": "table", "caption": "核對（應為零）", "head": ["項目"] + [f"@PL!{c}{IS['hdr']}" for c in is_cols],
     "rowgen": {"sheet": "PL", "rows": f"{IS['k1']}-{IS['k3']}", "cols": ["A"] + is_cols, "fmts": [""] + ["n3"] * 9}}]},
  {"id": "drivers", "title": "驅動因子：交付人員、價格保留與推論成本", "blocks": [
    {"type": "note", "parts": ["@Thesis!A1"]}, {"type": "note", "parts": ["@Thesis!A2"]},
    {"type": "theme", "title": "AI 淨效果的三個分項（情境 B；十億日圓）", "x": yrs("Thesis", TH_H, "I", "M"), "fmt": "n1", "stack": False, "pad": 1.2,
     "bars": [{"name": [f"@Thesis!A{EB[k]}"], "y": yrs("Thesis", EB[k], "I", "M"), "color": c} for k, c in (("prj", "#0b7a80"), ("ero", "#c29a2b"), ("inf", "#b4472f"))],
     "lines": [{"name": [f"@Thesis!A{EB['net']}"], "y": yrs("Thesis", EB["net"], "I", "M"), "color": "#3b3838", "labels": True}],
     "caption": "長條＝AI 新專案增量、價格侵蝕、推論成本（皆為正數）；折線＝AI 淨效果（營收口徑）。十億日圓；推論成本取 Tokenomics 市場價。"},
    {"type": "table", "caption": "單位經濟推導（情境 B；十億日圓除另註）", "head": ["項目"] + [f"@Thesis!{c}{TH_H}" for c in "HIJKLM"],
     "rowgen": {"sheet": "Thesis", "rows": f"{EB['auto']}-{EB['pinf']}", "cols": ["A"] + list("HIJKLM"), "fmts": [""] + ["n2"] * 6, "row_fmt": {str(EB["auto"]): {c: "p1" for c in "HIJKLM"}, str(EB["be"]): {c: "p1" for c in "HIJKLM"}, str(EB["N"]): {c: "n0" for c in "HIJKLM"}, str(EB["pe"]): {c: "n0" for c in "HIJKLM"}, str(EB["id"]): {c: "n3" for c in "HIJKLM"}}}},
    {"type": "table", "caption": "Tokenomics 快照讀值（只讀；價格層級為 Analogy）", "head": ["項目", "數值"], "rowgen": {"sheet": "Thesis", "rows": f"{min(v for k, v in TKR.items())}-{TKR['task']}", "cols": ["A", "C"], "fmts": ["", "n4"]}},
    {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；In 頁情境驅動對照）", "head": [f"@In!{c}{scn_h-1}" for c in "BDEF"],
     "rowgen": {"sheet": "In", "rows": f"{scn_h}-{scn_end}", "cols": ["B", "D", "E", "F"], "fmts": ["", "n3", "n3", "n3"]}},
    {"type": "table", "caption": "單一參數區間（In 頁；基準進主模型，敏感度低／高只用於敏感度頁、不進目標價）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rowgen": {"sheet": "In", "rows": in_rows, "cols": ["B", "D", "E", "F", "G"], "fmts": ["", "n3", "n3", "n3", ""]}}]},
  {"id": "recon", "title": "對帳：公司指引與共識", "blocks": [
    {"type": "note", "parts": ["@Recon!A1"]},
    {"type": "table", "caption": "指引與共識對照（十億日圓；EPS、股利為日圓）", "head": ["項目", "本模型", "指引", "共識", "共識口徑", "說明"], "rowgen": {"sheet": "Recon", "rows": f"{RCR['h1']+1}-{RCR['p1']}", "cols": list("ABCDEF"), "fmts": ["", "n2", "n2", "n2", "", ""]}},
    {"type": "table", "caption": "口徑與未找到", "head": ["項目"], "rowgen": {"sheet": "Recon", "rows": f"{RCR['p1']+1}-{RCR['sbc_h']-1}", "cols": ["A"], "fmts": [""]}},
    {"type": "table", "caption": "主導變數反解（使加權目標價＝現價所需的值）", "head": ["變數", "基準值", "反解值", "單位", "說明"], "rowgen": {"sheet": "Recon", "rows": f"{RCR['kreq_a']}-{RCR['kreq_a']+RCR['kreq_n']-1}", "cols": list("ABCDE"), "fmts": ["", "n3", "n3", "", ""]}}]},
  {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["build_id", "@Sens!D3"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"], ["Tokenomics 快照", "@Sources!B3"]]},
    {"type": "h3", "text": "市場資料日期與失效條件"},
    {"type": "list", "items": ["@README!A11", "@README!A21", "@README!A22"]},
    {"type": "table", "caption": "風險登錄", "head": ["風險", "是否建模", "理由", "觀察指標"], "rowgen": {"sheet": "README", "rows": "25-29", "cols": list("ABCD")}},
    {"type": "table", "collapse": True, "summary": "展開：每頁的輸入、推導量與因果方向", "head": ["頁", "輸入", "推導量", "因果方向"], "rowgen": {"sheet": "README", "rows": "34-37", "cols": list("ABCD")}},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A", "F"], "In": ["B", "H"], "Q": ["A"]}, "empty": "沒有標為未找到的項目。"}]}
 ]}
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("html.json 寫出")

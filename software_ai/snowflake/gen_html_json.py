# 由 build_snowflake.py 的列號對照（rowmap.json）產生 html.json（列號隨模型變動，勿手改 html.json）。
# 用法：python3 gen_html_json.py <成品 xlsx> <rowmap.json> html.json
import json, sys, openpyxl
X, RM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
rm = json.load(open(RM, encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)
IR, SR, PR, STR, SPR, VR, DR, RV, RCR = (rm[k] for k in ("IR", "SR", "PR", "STR", "SPR", "VR", "DR", "RV", "RCR"))
SB = SR["B"]; PHDR = 5
def yrs(sh, r, c1="C", c2="K"): return f"@{sh}!{c1}{r}:{c2}{r}"
def find(sh, text, col=1):
    ws = wb[sh]
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, col).value or "").startswith(text): return r
    raise KeyError((sh, text))
rec_r = VR["rec"]
tornL, tornH = [], []
for r in rm["SE_ROWS"]:
    side = wb["Sens"].cell(r, 3).value; v = wb["Sens"].cell(r, 5).value
    if isinstance(v, (int, float)) and abs(v) >= 0.001:
        (tornL if side == "低" else tornH).append(r)
cfg = {
 "company": "Snowflake", "ticker": "SNOW", "thesis": "@README!A3",
 "question": "AI 進入按用量計價的資料平台後，AI 工作負載與轉售營收的增量能否抵過核心消耗被替代的侵蝕與推論成本",
 "scenarios": [{"key": "A", "name": "情境 A（AI 工作負載高、被替代低、轉售加成高）"}, {"key": "B", "name": "情境 B（基準）"}, {"key": "C", "name": "情境 C（AI 工作負載低、被替代高、轉售加成低）"}],
 "wide_text": True, "layout": "v1.3",
 "hero": {"rec": f"@Val!B{rec_r}", "tiles": [[f"@Val!A{VR['tp']}", f"@Val!B{VR['tp']}|n2", ["上檔空間 ", f"@Val!B{VR['up']}|sp1"]], [f"@Val!A{VR['tp_ref']}", f"@Val!B{VR['tp_ref']}|n2", ["上檔空間 ", f"@Val!C{VR['tp_ref']}|sp1"]], [f"@Val!A{VR['px']}", f"@Val!B{VR['px']}|n2", ""]],
          "note": ["@README!A13"]},
 "sections": [
  {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "note", "parts": ["@README!A5"]},
    {"type": "theme", "title": "營收（十億美元，依來源堆疊）：FY2027E–FY2032E（情境 B）", "x": f"@Seg!F{SR['B']['n']-1}:K{SR['B']['n']-1}",
     "fmt": "n0", "stack": True, "pad": 1.15, "total": f"@Seg!F{SB['rev']}:K{SB['rev']}",
     "bars": [{"name": [f"@Seg!A{SB['core']}"], "y": f"@Seg!F{SB['core']}:K{SB['core']}", "color": "#6fa8dc"}, {"name": [f"@Seg!A{SB['air']}"], "y": f"@Seg!F{SB['air']}:K{SB['air']}", "color": "#c29a2b"},
              {"name": [f"@Seg!A{SB['ps']}"], "y": f"@Seg!F{SB['ps']}:K{SB['ps']}", "color": "#999999"}],
     "caption": "堆疊＝核心平台消耗、AI 營收增量（工作負載加轉售）、專業服務；單位為百萬美元，圖軸以十億計；官方分部只有一個，產品營收與專業服務為公司固定揭露線，核心與 AI 為本模型拆分。"},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [f"@README!A{r}" for r in range(6, 11)]},
    {"type": "h3", "text": "本版要點"},
    {"type": "list", "items": [f"@README!A{r}" for r in range(15, 19)]}]},
  {"id": "valuation", "title": "估值", "blocks": [
    {"type": "bars", "title": "目標價與現價（美元；豎線＝現價）", "fmt": "n2",
     "items": [{"label": f"@Val!A{VR[k]}", "value": f"@Val!B{VR[k]}", "series": sr} for k, sr in (("dcf_tp", 1), ("comp_tp", 2), ("tp", 0))] + [{"label": "賣方平均目標價（共識）", "value": f"@Recon!E{RCR['tp']}", "series": 3}],
     "ref": {"label": f"@Val!A{VR['px']}", "value": f"@Val!B{VR['px']}"}},
    {"type": "note", "parts": [f"@Val!A{VR['intro']}"]}, {"type": "note", "parts": [f"@Val!A{VR['hint']}"]},
    {"type": "kv", "items": [[f"@Val!A{VR[k]}", f"@Val!B{VR[k]}|{f}"] for k, f in (("ke", "p2"), ("kdat", "p2"), ("wacc", "p2"), ("ronic", "p2"), ("g", "p2"))]},
    {"type": "table", "caption": "結論（情境加權）", "head": ["項目", "數值", "對照"], "rowgen": {"sheet": "Val", "rows": f"{VR['dcf_tp']}-{VR['up']},{VR['val_rec']},{VR['gate_use']}-{VR['rec']}", "cols": ["A", "B"], "fmts": ["", "n2"], "row_fmt": {str(VR['up']): {"B": "sp1"}, str(VR['val_rec']): {"B": ""}, str(VR['gate_use']): {"B": "n0"}, str(VR['gate_res']): {"B": ""}, str(VR['rec']): {"B": ""}}}},
    {"type": "kv", "items": [[f"@Val!A{VR[k]}", f"@Val!B{VR[k]}|n2"] for k in ("dec_gap", "dec_sbc", "dec_res", "rerate")]},
    {"type": "table", "caption": "DCF 彙總（三情境）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rowgen": {"sheet": "Val", "rows": f"{VR['pvx']}-{VR['tvsh']}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1"], "row_fmt": {str(VR['tvsh']): {c: "p1" for c in "CDE"}, str(VR['vps0']): {c: "n2" for c in "CDE"}, str(VR['dcf12']): {c: "n2" for c in "CDE"}}}},
    {"type": "table", "caption": "Comparable（同業前瞻 P/S 中位數 × 目標時點 NTM 營收）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rowgen": {"sheet": "Val", "rows": f"{VR['ntm']}-{VR['comp12']},{VR['wb']}-{VR['wc']}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1"], "row_fmt": {str(VR['mult']): {c: "x1" for c in "CDE"}, str(VR['comp12']): {c: "n2" for c in "CDE"}, str(VR['wb']): {c: "n2" for c in "CDE"}, str(VR['wc']): {c: "sp1" for c in "CDE"}}}},
    {"type": "table", "collapse": True, "summary": "展開：同業倍數與淨現金", "head": ["項目", "前瞻 P/S（或數值）", "前瞻 P/E", "市值（十億美元）"], "rowgen": {"sheet": "Val", "rows": f"{VR['cashx']}-{VR['sh']},{VR['pe1']}-{VR['pe_own']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n2", "n1", "n1"]}},
    {"type": "table", "collapse": True, "summary": "展開：主腿同業 P/S（現價日；七家）", "head": ["項目", "數值", "備註"], "rowgen": {"sheet": "Val", "rows": f"{VR['cu_h']+1}-{VR['cu_pe']}", "cols": ["A", "B", "C"], "fmts": ["", "n2", ""]}},
    {"type": "table", "collapse": True, "summary": "展開：對照列（前瞻 P/S 口徑，不進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rowgen": {"sheet": "Val", "rows": f"{VR['fwd_ntm']}-{VR['fwd_cmp']}", "cols": ["A", "C", "D", "E"], "fmts": ["", "n2", "n2", "n2"]}},
    {"type": "table", "collapse": True, "summary": "展開：rel_mult（共用庫 JSON 讀數）", "head": ["項目", "數值"], "rowgen": {"sheet": "Val", "rows": f"{rm['RM']['first']}-{rm['RM']['last']}", "cols": ["A", "B"], "fmts": ["", "n3"]}},
    {"type": "table", "caption": "DCF（股票薪酬加回）對照讀數（只作對照，不進加權目標價）", "head": ["項目", "情境 A", "情境 B", "情境 C"], "rowgen": {"sheet": "Recon", "rows": f"{RCR['sb_v0']}-{RCR['sb_v12']}", "cols": ["A", "B", "C", "D"], "fmts": ["", "n2", "n2", "n2"]}},
    {"type": "kv", "items": [[f"@Recon!A{RCR['sb_w']}", f"@Recon!B{RCR['sb_w']}|n2"], [f"@Recon!A{RCR['sb_d']}", f"@Recon!B{RCR['sb_d']}|n2"]]},
    {"type": "note", "parts": ["@README!A40", "@README!A41", "@README!A42"]}]},
  {"id": "sens", "title": "敏感度", "blocks": [
    {"type": "tornado", "title": "加權目標價相對基準：輸入取低值（一次改一個，整本重算）", "sheet": "Sens", "rows": ",".join(map(str, tornL)), "label_col": "A", "value_col": "E", "fmt": "sp1"},
    {"type": "tornado", "title": "加權目標價相對基準：輸入取高值（一次改一個，整本重算）", "sheet": "Sens", "rows": ",".join(map(str, tornH)), "label_col": "A", "value_col": "E", "fmt": "sp1"},
    {"type": "note", "parts": ["@Sens!A1"]},
    {"type": "table", "collapse": True, "summary": "展開：全部案例（含 DCF 單法與投資建議）", "head": ["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "DCF 目標價", "DCF 變動（美元）", "投資建議"],
     "rowgen": {"sheet": "Sens", "rows": f"{rm['HDR_R']+1}-{rm['SE_END']}", "cols": list("ABCDEFGHI"), "fmts": ["", "", "", "n2", "sp1", "", "n2", "n2", ""]}},
    {"type": "kv", "items": [["build_id", "@Sens!D3"], ["模型狀態", "@Sens!D4"], ["@Sens!A5", "@Sens!D5|n2"], ["@Sens!A6", "@Sens!D6|n2"]]}]},
  {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "blocks": [
    {"type": "note", "parts": ["@RevDCF!A1"]}, {"type": "note", "parts": ["@RevDCF!A2"]},
    {"type": "table", "caption": "現價隱含值與模型基準（情境 B）", "head": ["項目", "數值"], "rowgen": {"sheet": "RevDCF", "rows": f"{RV['v0']}-{RV['mult_peer']}", "cols": ["A", "B"], "fmts": ["", "n2"],
      "row_fmt": {str(RV[k]): {"B": f} for k, f in (("m_fy32", "p1"), ("m_imp", "p1"), ("scale", "n2"), ("mult_imp", "x1"), ("mult_peer", "x1"), ("gap", "n0"), ("k_m", "n0"), ("ev0", "n0"), ("ev_imp", "n0"), ("d_m", "n2"))}}}]},
  {"id": "pl", "title": "損益（標準損益表）", "blocks": [
    {"type": "note", "parts": ["@PL!A1"]}, {"type": "note", "parts": ["@PL!A2"]},
    {"type": "theme", "title": "營收（十億美元）與非 GAAP 營業利益率：FY2024A–FY2032E（情境 B）", "x": yrs("PL", PHDR), "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.3,
     "bars": [{"name": [f"@PL!A{STR['營收']}"], "y": yrs("PL", STR["營收"]), "color": "#6fa8dc"}],
     "lines": [{"name": [f"@PL!A{STR['營業利益率']}"], "y": yrs("PL", STR["營業利益率"]), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "長條＝營收（百萬美元，軸以十億計）；折線＝非 GAAP 營業利益率（右軸）。FY2027E 起為情境 B。"},
    {"type": "table", "caption": "標準損益表（非 GAAP／Non-GAAP，公司與賣方共識口徑；百萬美元，EPS 為美元）：FY2024A–FY2026A 實績、FY2027E–FY2032E 情境 B", "head": ["項目"] + [f"@PL!{c}{PHDR}" for c in "CDEFGHIJK"],
     "rowgen": {"sheet": "PL", "rows": f"{STR['營收']}-{STR['淨利年增率']}", "cols": ["A"] + list("CDEFGHIJK"), "fmts": [""] + ["n0"] * 9,
                "row_fmt": {**{str(STR[n]): {c: "p1" for c in "CDEFGHIJK"} for n in ("營收年增率", "毛利率", "營業利益率", "淨利率", "淨利年增率")}, str(STR["EPS（非 GAAP 稀釋，美元）"]): {c: "n2" for c in "CDEFGHIJK"}}}},
    {"type": "note", "parts": [f"@PL!A{STR['淨利年增率']+1}"]},
    {"type": "table", "caption": "參考：GAAP 營業利益、股票薪酬與公司口徑自由現金流（FY2027E 起 GAAP 為近似，不進估值）", "head": ["項目"] + [f"@PL!{c}{PHDR}" for c in "CDEFGHIJK"],
     "rowgen": {"sheet": "PL", "rows": f"{rm['GR0']}-{rm['GR0']+4}", "cols": ["A"] + list("CDEFGHIJK"), "fmts": [""] + ["n0"] * 9, "row_fmt": {str(rm['GR0']+2): {c: "p1" for c in "CDEFGHIJK"}, str(rm['GR0']+4): {c: "p1" for c in "CDEFGHIJK"}}}},
    {"type": "theme", "title": "FCFF：情境 A／B／C（十億美元）", "x": f"@PL!F{PHDR}:K{PHDR}", "fmt": "n0", "stack": False, "pad": 1.2,
     "bars": [{"name": ["情境 ", s, "：FCFF"], "y": f"@PL!F{PR[s]['fcff']}:K{PR[s]['fcff']}", "color": c} for s, c in (("A", "#7fb8bb"), ("B", "#3b3838"), ("C", "#d9a08f"))],
     "caption": "三情境的無槓桿自由現金流（EBIT 已扣股票薪酬；百萬美元，軸以十億計）；情境差異來自 AI 工作負載、被替代、轉售加成與股票薪酬占比。"},
    {"type": "table", "caption": "分部營收（情境 B；單一營運分部下的固定揭露營收線與命題子列）", "head": ["項目"] + [f"@PL!{c}{PHDR}" for c in "CDEFGHIJK"],
     "rowgen": {"sheet": "PL", "rows": f"{min(SPR.values())}-{max(SPR.values())}", "cols": ["A"] + list("CDEFGHIJK"), "fmts": [""] + ["n0"] * 9}},
    {"type": "table", "caption": "核對（應為零）", "head": ["項目"] + [f"@PL!{c}{PHDR}" for c in "CDEFGHIJK"],
     "rowgen": {"sheet": "PL", "rows": f"{rm['CK']}-{rm['CK']+3}", "cols": ["A"] + list("CDEFGHIJK"), "fmts": [""] + ["n3"] * 9}}]},
  {"id": "drivers", "title": "驅動因子：用量、AI 工作負載與推論成本", "blocks": [
    {"type": "note", "parts": ["@Seg!A1"]}, {"type": "note", "parts": ["@Seg!A2"]},
    {"type": "theme", "title": "AI 淨效果的三個分項（情境 B；十億美元）", "x": f"@Seg!F{SB['n']-1}:K{SB['n']-1}", "fmt": "n1", "stack": False, "pad": 1.2,
     "bars": [{"name": [f"@Seg!A{SB[k]}"], "y": f"@Seg!F{SB[k]}:K{SB[k]}", "color": c} for k, c in (("air", "#0b7a80"), ("eros", "#c29a2b"), ("infc", "#b4472f"))],
     "lines": [{"name": [f"@Seg!A{SB['net']}"], "y": f"@Seg!F{SB['net']}:K{SB['net']}", "color": "#3b3838", "labels": True}],
     "caption": "長條＝AI 營收增量、核心侵蝕、推論成本（皆為正數）；折線＝AI 淨效果。百萬美元，軸以十億計。推論成本取 Tokenomics 市場價。"},
    {"type": "table", "caption": "營收推導（情境 B；百萬美元除另註）", "head": ["項目"] + [f"@Seg!{c}{SB['n']-1}" for c in "FGHIJK"],
     "rowgen": {"sheet": "Seg", "rows": f"{SB['n']}-{SB['chk2']}", "cols": ["A"] + list("FGHIJK"), "fmts": [""] + ["n1"] * 6,
                "row_fmt": {**{str(SB[k]): {c: "p1" for c in "FGHIJK"} for k in ("gb", "w", "e", "aish", "gm0n", "pgm")}, str(SB["n"]): {c: "n0" for c in "FGHIJK"}, str(SB["j"]): {c: "n2" for c in "FGHIJK"}, str(SB["pxidx"]): {c: "n2" for c in "FGHIJK"}, str(SB["cps"]): {c: "g" for c in "FGHIJK"}}}},
    {"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；In 頁情境驅動對照）", "head": [f"@In!{c}{rm['SCN_H']}" for c in "BDEF"],
     "rowgen": {"sheet": "In", "rows": f"{rm['SCN_H']+1}-{rm['SCN_H']+4}", "cols": ["B", "D", "E", "F"], "fmts": ["", "p1", "p1", "p1"]}},
    {"type": "table", "caption": "單一參數區間（In 頁；基準進主模型，敏感度低／高只用於敏感度頁、不進目標價）", "head": ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!G3"],
     "rowgen": {"sheet": "In", "rows": f"4-{rm['LAST_PARAM_ROW']}", "cols": ["B", "D", "E", "F", "G"], "fmts": ["", "g", "g", "g", ""]}}]},
  {"id": "recon", "title": "對帳：公司指引與共識", "blocks": [
    {"type": "note", "parts": ["@Recon!A1"]},
    {"type": "table", "caption": "指引與共識對照", "head": ["項目", "本模型", "指引低", "指引高", "共識", "說明"], "rowgen": {"sheet": "Recon", "rows": f"{RCR['h1']+1}-{RCR['fwdps']}" + (f",{RCR['kreq']}-{RCR['kreq2']}" if 'kreq2' in RCR else ""), "cols": list("ABCDEF"), "fmts": ["", "g", "g", "g", "g", ""]}},
    {"type": "table", "caption": "口徑與未找到", "head": ["項目"], "rowgen": {"sheet": "Recon", "rows": f"{RCR['fwdps']+(2 if 'kreq' not in RCR else 4)}-{wb['Recon'].max_row}", "cols": ["A"], "fmts": [""]}}]},
  {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["版本與日期", "@README!A1"], ["build_id", "@Sens!D3"], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋", "$sha"], ["Tokenomics 快照", "@Sources!B3"]]},
    {"type": "h3", "text": "市場資料日期與失效條件"},
    {"type": "list", "items": ["@README!A21", "@README!A22"]},
    {"type": "table", "caption": "風險登錄", "head": ["風險", "是否建模", "理由", "觀察指標"], "rowgen": {"sheet": "README", "rows": "25-29", "cols": list("ABCD")}},
    {"type": "table", "collapse": True, "summary": "展開：每頁的輸入、推導量與因果方向", "head": ["頁", "輸入", "推導量", "因果方向"], "rowgen": {"sheet": "README", "rows": "34-37", "cols": list("ABCD")}},
    {"type": "h3", "text": "「未找到」清單"},
    {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的項目", "sheets": {"Recon": ["A", "F"], "In": ["B"], "Q": ["A"]}, "empty": "沒有標為未找到的項目。"}]}
 ]}
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("html.json 寫出")

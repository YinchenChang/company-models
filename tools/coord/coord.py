# 三層基準協調規則 v1.4 C4（個股不裁決、並列兩讀數目標價）的共用工具。樣板：NVIDIA v0.3.8（nvidia/coord_post.py）；
# 本模組把其中與公司無關的部分抽出，供其他硬體個股使用（第一家：Broadcom）。
#
# 流程（每家）：
#   1. 建檔程式的主引擎在 SIL.COORD_ALT（環境變數 SIL_COORD_ALT＝1）時，情境 B、C 於 CY2027 起產業量 × wta_<yy>
#      （siliconomics/reader.py 在該模式下把 wta 列換成另一側讀數比值；情境 A＝供給側讀數不縮放）。
#   2. python3 tools/coord/alt_run.py <資料夾>：整本建置兩次（現行、另一側），LibreOffice 重算，讀 Val 頁 → <資料夾>/coord_alt.json。
#   3. 建檔程式存檔前呼叫 add_sheet(wb, <資料夾>, SIL)：新增 Coord 頁（現行讀數連到 Val 頁、另一側讀數為快照、快照狀態 CURRENT／STALE）。
#   4. gen_html_json.py 最後呼叫 inject_html(cfg, xlsx 路徑)：橫幅加一格與揭露句、估值章開頭加兩讀數表與長條圖。
# 數字一律在 xlsx；HTML 只引用 Coord 頁儲存格。
import json, os, re

YEARS = (2026, 2027, 2028, 2029)
LAB = {  # Val 頁列標籤（與 tools/target_summary/build_summary.py、markets/invest_map/extract_models.py 相同）
    "tp": r"^目標價（DCF 與 Comparable 加權）",
    "dcf": r"^DCF 目標價（情境加權）",
    "comp": r"^Comparable 目標價（市場主流同業法）",
    "up": r"^上檔空間：加權目標價",
}
KEYS = ("tp", "dcf", "comp", "up")
NAMES = {"tp": "加權目標價（DCF 與 Comparable 加權）", "dcf": "DCF 目標價（情境加權）", "comp": "Comparable 目標價（市場主流同業法；未調整）", "up": "上檔空間（加權目標價 ÷ 現價 − 1）"}
REC = lambda c: f'IF({c}>0.5,"強力買進",IF({c}>0.25,"買進",IF({c}>=-0.25,"中立",IF({c}>=-0.5,"放空","強力放空"))))'   # Val 頁五級固定規則


def rec(u):
    return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"


def val_rows(ws):
    """Val 頁：鍵 → 列號（取第一個 B 欄有值的列；公式活頁簿與重算後皆可）。"""
    out = {}
    for r in range(1, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "")
        for k, pat in LAB.items():
            if k not in out and re.search(pat, a) and ws.cell(r, 2).value is not None:
                out[k] = r
    missing = [k for k in KEYS if k not in out]
    assert not missing, f"Val 頁找不到：{missing}"
    return out


def reco_years(SIL):
    return [y for y in YEARS if SIL.RECO_FLAG.get(y) == 1]


def add_sheet(wb, folder, SIL):
    """新增 Coord 頁。folder＝公司資料夾（讀 coord_alt.json）。回傳列號 dict。無對帳中年份或無快照時不建，回傳 None。"""
    ys = reco_years(SIL)
    f = os.path.join(folder, "coord_alt.json")
    if not ys or os.environ.get("COORD_RUN") == "1" or SIL.COORD_ALT:   # alt_run.py 建置中不加 Coord 頁
        return None
    assert os.path.exists(f), f"{f} 不存在：先跑 python3 tools/coord/alt_run.py {os.path.basename(folder)}"
    J = json.load(open(f, encoding="utf-8"))
    from openpyxl.styles import Font, PatternFill
    BOLD, HDR = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
    V = val_rows(wb["Val"])
    ws = wb.create_sheet("Coord")
    ws.column_dimensions["A"].width = 46
    for c in "BCDE":
        ws.column_dimensions[c].width = 22
    ys_txt = "、".join(f"CY{y}" for y in ys)
    ws["A1"] = ("Coord — 三層協調：對帳中標示與兩讀數目標價（三層基準協調規則 v1.4 C4；共用工具 tools/coord/）。個股不裁決：基準只讀 Siliconomics 現行基準；"
                f"{ys_txt} 在 Siliconomics 標為對帳中（對帳單見 Drive 02-AI供應鏈/對帳/），結案前並列兩個讀數的目標價。")
    ws["A1"].font = BOLD
    ws["A2"] = ("另一側讀數＝只換產業部署 GW 一個輸入（Siliconomics SIF_DeployGW_AltRatio_<年>，即 Wattonomics 換算值 ÷ 現行基準）：主引擎情境 B、C 於 CY2027 起產業量 × 比值"
                "（只在對帳中年份；情境 A＝供給側讀數不縮放），其他假設不變，整本重建（tools/coord/alt_run.py）後讀 Val 頁。")
    R = {}
    r = 4
    for j, h in enumerate(["年", *[f"CY{y}" for y in YEARS]]):
        c = ws.cell(r, 1 + j, h); c.font = BOLD; c.fill = HDR
    r += 1
    R["flag"] = r; ws.cell(r, 1, "對帳中旗標（Siliconomics SIF_RecoFlag_<年>）")
    for j, y in enumerate(YEARS):
        ws.cell(r, 2 + j, SIL.RECO_FLAG[y])
    r += 1
    R["ratio"] = r; ws.cell(r, 1, "另一側讀數 ÷ 現行基準（SIF_DeployGW_AltRatio_<年>）")
    for j, y in enumerate(YEARS):
        c = ws.cell(r, 2 + j, SIL.ALT_RATIO[y]); c.number_format = "0.000"
    r += 1
    R["lab"] = r; ws.cell(r, 1, "Siliconomics 該年標示")
    for j, y in enumerate(YEARS):
        ws.cell(r, 2 + j, SIL.NAMES.get(f"SIF_RecoLabel_{y}", ""))
    r += 1
    ws.cell(r, 1, f"來源：Siliconomics {SIL.VERSION}（{SIL.FILE}）"); r += 2
    for j, h in enumerate(["項目", "現行基準讀數（Siliconomics 基準；橫幅主值）", "另一側讀數（Wattonomics 換算的產業 GW）", "另一側相對現行"]):
        c = ws.cell(r, 1 + j, h); c.font = BOLD; c.fill = HDR
    R["hdr"] = r; r += 1
    for k in KEYS:
        R[k] = r
        ws.cell(r, 1, NAMES[k])
        ws.cell(r, 2, f"=Val!B{V[k]}")
        ws.cell(r, 3, J["alt"][k])
        if k == "up":
            ws.cell(r, 4, f"=C{r}-B{r}"); ws.cell(r, 4).number_format = '+0.0%;-0.0%'
            ws.cell(r, 2).number_format = ws.cell(r, 3).number_format = '+0.0%;-0.0%'
        else:
            ws.cell(r, 4, f"=C{r}/B{r}-1"); ws.cell(r, 4).number_format = '+0.0%;-0.0%'
            ws.cell(r, 2).number_format = ws.cell(r, 3).number_format = "#,##0.00"
        r += 1
    R["rec"] = r; ws.cell(r, 1, "投資建議（五級固定規則）")
    ws.cell(r, 2, "=" + REC(f"B{R['up']}")); ws.cell(r, 3, "=" + REC(f"C{R['up']}")); r += 1
    R["state"] = r; ws.cell(r, 1, "另一側讀數快照狀態")
    ws.cell(r, 2, f'=IF(ABS(B{R["tp"]}-{J["base"]["tp"]!r})<=1E-6*ABS({J["base"]["tp"]!r}),"CURRENT","STALE")')
    ws.cell(r, 3, f"快照 {J['snapshot']}；現行讀數建置時加權目標價 {J['base']['tp']:.6g}（與 Val 頁不同即 STALE：改輸入後須重跑 tools/coord/alt_run.py）"); r += 1
    R["dep"] = r; ws.cell(r, 1, "是否有估值依賴的對帳中年份"); ws.cell(r, 2, "是"); r += 1
    R["yrs"] = r; ws.cell(r, 1, "對帳中年份"); ws.cell(r, 2, "、".join(str(y) for y in ys)); r += 1
    R["lab_cur"] = r; ws.cell(r, 1, "目標價區標籤：現行"); ws.cell(r, 2, "現行基準讀數（Siliconomics 基準；橫幅主值）"); r += 1
    R["lab_alt"] = r; ws.cell(r, 1, "目標價區標籤：另一側"); ws.cell(r, 2, "另一側讀數（Wattonomics 換算的產業 GW）"); r += 1
    rr = ["TEXT(" + f"{'BCDE'[YEARS.index(y)]}{R['ratio']}" + ',"0.00")' for y in ys]
    R["note"] = r; ws.cell(r, 1, "目標價區說明")
    ws.cell(r, 2, f'="兩讀數只差產業部署 GW 一個輸入：{ys_txt} 另一側讀數為現行基準的 "&' + '&"／"&'.join(rr) +
            '&" 倍（情境 B、C 的產業量同乘；情境 A＝供給側讀數不變），其他假設不變。個股不裁決哪一側正確，對帳在 Siliconomics 對帳單處理（期限 2027-02-28）；結案前兩個目標價並列。"'); r += 1
    R["hero"] = r; ws.cell(r, 1, "橫幅揭露")
    ws.cell(r, 2, f'="對帳中（{ys_txt}）：另一側讀數目標價 "&TEXT(C{R["tp"]},"#,##0.00")&"（"&C{R["rec"]}&"）"&IF(B{R["rec"]}<>C{R["rec"]},"，與現行基準讀數（"&B{R["rec"]}&"）評等不同","")&"。"'); r += 1
    return R


def _find(ws, prefix):
    return next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value or "").startswith(prefix))


def inject_html(cfg, xlsx):
    """html.json：橫幅加「另一側讀數目標價」一格與揭露句；估值章開頭加兩讀數表與長條圖。xlsx 沒有 Coord 頁時不動。"""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx, read_only=True)
    if "Coord" not in wb.sheetnames:
        return False
    wb = openpyxl.load_workbook(xlsx)
    CW = wb["Coord"]
    f = lambda p: _find(CW, p)
    r_tp, r_up, r_rec, r_dcf, r_comp = f("加權目標價"), f("上檔空間"), f("投資建議"), f("DCF 目標價"), f("Comparable 目標價")
    r_hero, r_cur, r_alt, r_note, r_yrs = f("橫幅揭露"), f("目標價區標籤：現行"), f("目標價區標籤：另一側"), f("目標價區說明"), f("對帳中年份")
    cfg["hero"]["tiles"].append([["另一側讀數目標價（對帳中）"], f"@Coord!C{r_tp}|n2", ["上檔空間 ", f"@Coord!C{r_up}|sp1", "；", f"@Coord!C{r_rec}"]])
    cfg["hero"]["note"] = list(cfg["hero"]["note"]) + [" ", f"@Coord!B{r_hero}"]
    rows = [r_tp, r_dcf, r_comp, r_up]
    fm = lambda r: "sp1" if r == r_up else "n2"
    tbl = {"type": "table", "caption": ["兩讀數目標價（對帳中：", f"@Coord!B{r_yrs}", "）——現行基準讀數與另一側讀數並列"],
           "head": ["項目", f"@Coord!B{r_cur}", f"@Coord!B{r_alt}", "另一側相對現行"],
           "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|{fm(r)}", f"@Coord!C{r}|{fm(r)}", f"@Coord!D{r}|sp1"] for r in rows]
                   + [[f"@Coord!A{r_rec}", f"@Coord!B{r_rec}", f"@Coord!C{r_rec}", ""]]}
    V = wb["Val"]
    px = next((r for r in range(1, V.max_row + 1) if re.match(r"^現價（", str(V.cell(r, 1).value or ""))), None)
    bars = {"type": "bars", "title": "兩讀數的加權目標價與現價（豎線＝現價）", "fmt": "n2",
            "items": [{"label": f"@Coord!B{r_cur}", "value": f"@Coord!B{r_tp}", "series": 0}, {"label": f"@Coord!B{r_alt}", "value": f"@Coord!C{r_tp}", "series": 2}]}
    if px:
        bars["ref"] = {"label": f"@Val!A{px}", "value": f"@Val!B{px}"}
    S = {x["id"]: x for x in cfg["sections"]}
    S["valuation"]["blocks"][0:0] = [{"type": "h3", "text": "對帳中：兩讀數目標價"}, {"type": "note", "parts": [f"@Coord!B{r_note}"]}, tbl, bars]
    return True


def stamp(wb, old_ver, new_ver, date):
    """README A1 版號 old_ver → new_ver，並在 A1 末加本版說明（三層協調兩讀數）。"""
    RD = wb["README"]
    a1 = RD["A1"].value
    assert isinstance(a1, str) and old_ver in a1, ("README A1 無版號", old_ver)
    head = a1.split("；", 1)[0]   # 只改標題段的版號；標題沒有現行版號時（舊檔標題未更新）只在末尾加說明，不動版本沿革
    if old_ver in head:
        a1 = head.replace(old_ver, new_ver, 1) + a1[len(head):]
    RD["A1"].value = a1 + (f"；{new_ver}（{date}）：三層協調——Coord 頁並列兩讀數目標價（三層基準協調規則 v1.4 C4；"
                                                      "Siliconomics v0.19 對帳中年份；tools/coord/），HTML 橫幅與估值章加另一側讀數；既有數值不變")

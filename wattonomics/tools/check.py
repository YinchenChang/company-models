# GW 帳核對程式：重算後活頁簿的錯誤格 0、Checks 全過（WGOV_Errors＝0）、In 頁標記與來源齊全、具名範圍可解析、快照一致（給快照時）。
# 用法：python3 wattonomics/tools/check.py <重算後的 Wattonomics xlsx> [<snapshot.json>]
# 任何 ERROR 以代碼 1 結束；WARN 只列出。
import json, os, sys
import openpyxl

ERRS = ("#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "Err:")
REQUIRED = ["README", "In", "Players", "CapexGW", "Delay", "Funding", "Gap", "FundRaw", "Facility", "SegIn", "Spend", "Matrix", "Pipeline", "Deploy", "C0", "C1", "Compare", "TK_In", "SIL_In", "Interface", "Checks",
            "Decisions", "DB_Evidence", "Sources"]
TAGS = ("Verified", "Interested-party", "Analogy", "Assumed", "Derived", "待證據")


def main():
    same = None
    if "--same-gross" in sys.argv:   # v0.2.1 W2：給前一版快照時，驗證年新增毛額、年底與 R4 用的基準名稱不變
        i_ = sys.argv.index("--same-gross"); same = sys.argv[i_ + 1]; del sys.argv[i_:i_ + 2]
    x = sys.argv[1]
    wb = openpyxl.load_workbook(x, data_only=True)
    wf = openpyxl.load_workbook(x, data_only=False)
    bad = []
    # 1) 頁面
    miss = [s for s in REQUIRED if s not in wb.sheetnames]
    if miss: bad.append(f"缺頁面：{miss}")
    # 2) 錯誤格與未重算（公式格無快取值）
    nerr = 0; nuncalc = 0
    for ws in wb.worksheets:
        wsf = wf[ws.title]
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith(ERRS):
                    nerr += 1
                    if nerr <= 10: print(f"  錯誤格 {ws.title}!{c.coordinate}：{v}")
                fv = wsf[c.coordinate].value
                if isinstance(fv, str) and fv.startswith("=") and v is None:
                    nuncalc += 1
    if nerr: bad.append(f"錯誤格 {nerr}")
    if nuncalc > 0:
        print(f"  註：{nuncalc} 個公式格值為空（空字串結果或未重算；若全部為空表示未經 LibreOffice 重算）")
    # 3) Checks
    ck = wb["Checks"]
    rows = [(ck.cell(r, 1).value, ck.cell(r, 2).value, ck.cell(r, 6).value) for r in range(4, ck.max_row + 1) if ck.cell(r, 1).value]
    ce = [x_ for x_ in rows if x_[2] == "ERROR"]; cw = [x_ for x_ in rows if x_[2] == "WARN"]
    nok = sum(1 for x_ in rows if x_[2] == "OK")
    ninfo = sum(1 for x_ in rows if x_[2] == "INFO")   # v0.1.9：INFO 類（對帳點），不計入 ERROR／WARN
    for i, l, s in ce: print(f"  Checks ERROR {i}：{l}")
    for i, l, s in cw: print(f"  Checks WARN  {i}：{l}")
    names = {}
    for nm, dn in wb.defined_names.items():
        names[nm] = dn.attr_text
    def val(nm):
        sh, ref = names[nm].split("!")
        return wb[sh.strip("'")][ref.replace("$", "")].value
    ge = val("WGOV_Errors")
    if ge != 0: bad.append(f"WGOV_Errors＝{ge}")
    if ce: bad.append(f"Checks ERROR {len(ce)}")
    # 4) In 頁標記與來源（獨立於 Excel 公式再檢一次）
    ws = wb["In"]; src_ids = {wb["Sources"].cell(r, 1).value for r in range(4, wb["Sources"].max_row + 1)}
    ntag = nsrc = nbad = 0; nrow = 0; pend = 0
    for r in range(5, ws.max_row + 1):
        if not ws.cell(r, 19).value: continue
        nrow += 1
        tag, sid = ws.cell(r, 5).value, ws.cell(r, 6).value
        if not tag: ntag += 1
        elif not any(t in str(tag) for t in TAGS): nbad += 1; print(f"  In!E{r} 標記不在清單：{tag}")
        if not sid: nsrc += 1
        elif sid not in src_ids: nbad += 1; print(f"  In!F{r} 來源 ID 不在 Sources：{sid}")
        if tag == "待證據": pend += 1
    if ntag or nsrc or nbad: bad.append(f"In 標記缺 {ntag}、來源缺 {nsrc}、無效 {nbad}")
    # 5) WIF_ 名稱
    gif = [n for n in names if n.startswith("WIF_")]
    for n in gif:
        v = val(n)
        if isinstance(v, str) and v.startswith(ERRS): bad.append(f"{n} 錯誤")
    # 5b) v0.2 V5（工作單 20261104）：停用情境名稱已刪；新別名＝舊名；Excel 文字不寫「需求 A」
    import re as _re
    stale = [n for n in gif if _re.search(r"_(B|C)_D\d|_A_D[12]|ShippedNotOnline|PlayerCumGap_A_D1", n)]
    if stale: bad.append(f"停用情境名稱仍在：{stale[:5]}（共 {len(stale)}）")
    for k_ in ("YE", "Avg", "Add"):
        for y in range(2024, 2033):
            if f"WIF_MktGW_{k_}_{y}" not in names: bad.append(f"缺別名 WIF_MktGW_{k_}_{y}"); continue
            if val(f"WIF_MktGW_{k_}_{y}") != val(f"WIF_MktGW_{k_}_A_D0_{y}"): bad.append(f"別名 WIF_MktGW_{k_}_{y} ≠ 舊名")
    nA = sum(1 for ws_ in wf.worksheets if ws_.title not in ("FundRaw", "DB_Evidence") for row in ws_.iter_rows() for c in row
             if isinstance(c.value, str) and not c.value.startswith("=") and "需求 A" in c.value and not (ws_.title == "Decisions" and c.column == 6))
    if nA: bad.append(f"Excel 文字含「需求 A」{nA} 格（Andy 11-04 07:44；Decisions 依據欄的原話除外）")
    nS = sum(1 for ws_ in wf.worksheets if ws_.title not in ("FundRaw", "DB_Evidence", "Decisions") for row in ws_.iter_rows() for c in row
             if isinstance(c.value, str) and "情景" in c.value)
    if nS: bad.append(f"Excel 標籤含「情景」{nS} 格（獨立核對 2026-11-04 建議 5）")
    # v0.2 修訂（複核 2026-11-04）：edits/v<版>_In_edits.csv 每列都須在 xlsx 生效（In 頁依鍵比對 new 值）
    import csv as _csv, glob as _glob
    ver = os.path.basename(x).split("_v")[-1].rsplit(".xlsx", 1)[0]
    ef = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "edits", f"v{ver}_In_edits.csv")
    if os.path.exists(ef):
        from openpyxl.utils import column_index_from_string as _ci
        wsI = wf["In"]; rowof = {wsI.cell(r_, 1).value: r_ for r_ in range(5, wsI.max_row + 1) if wsI.cell(r_, 1).value}
        nmiss = 0
        for e in _csv.DictReader(open(ef, encoding="utf-8-sig")):
            cur = wsI.cell(rowof[e["key"]], _ci(e["col"])).value if e["key"] in rowof else "<缺鍵>"
            ok_ = (str(cur) == e["new"]) or (cur is None and e["new"] == "")
            try: ok_ = ok_ or abs(float(cur) - float(e["new"])) < 1e-9
            except (TypeError, ValueError): pass
            if not ok_:
                nmiss += 1; print(f"  edits 未生效：{e['key']} {e['col']}：xlsx {cur!r} ≠ {e['new']!r}")
        if nmiss: bad.append(f"{os.path.basename(ef)} 有 {nmiss} 列未在 xlsx 生效")
        else: print(f"  {os.path.basename(ef)}：全部列已生效")
    # v0.2.1 獨立核對第 3 項：上市年＋壽命以前不得出現該世代的汰換（上市年＝Players 世代占比首個 > 0 的年份；H100 視為 2023 前已上市）
    if "Players" in wb.sheetnames and "Deploy" in wb.sheetnames:
        PLs, DPs = wb["Players"], wb["Deploy"]
        yrs = {c_: int(DPs.cell(3, c_).value) for c_ in range(4, 14) if str(DPs.cell(3, c_).value).isdigit()}
        life = round(val("life_ret")) if "life_ret" in names else 6
        mixrow = {str(PLs.cell(r_, 1).value).replace("世代占比：", ""): r_ for r_ in range(4, 15) if str(PLs.cell(r_, 1).value).startswith("世代占比：")}
        nbadg = 0
        for r_ in range(4, DPs.max_row + 1):
            lab = str(DPs.cell(r_, 1).value or "")
            if not lab.startswith("　其中 NVIDIA "): continue
            gname = lab.replace("　其中 NVIDIA ", "")
            if gname not in mixrow: continue
            launch = next((y_ for c_, y_ in yrs.items() if (PLs.cell(mixrow[gname], c_).value or 0) > 0), 9999)
            if gname.startswith("Hopper"): launch = min(launch, 2023) - life   # 最舊世代：早於 2023 的存量歸此
            for c_, y_ in yrs.items():
                v_ = DPs.cell(r_, c_).value or 0
                if y_ < launch + life and abs(v_) > 1e-9:
                    nbadg += 1; print(f"  世代汰換早於上市年＋壽命：{gname} {y_} 年 {v_:.3f}（上市 {launch}）")
        if nbadg: bad.append(f"世代汰換早於上市年＋{life} 年：{nbadg} 格")
        else: print(f"  世代汰換：無早於上市年＋{life} 年者")
    # v0.2.2 X1：SIL_In 頁不得讀 Siliconomics 情境名稱 A／B／C（只讀 *_Base）
    import re as _re2
    if "SIL_In" in wf.sheetnames:
        abc = [str(c.value) for row in wf["SIL_In"].iter_rows(min_col=1, max_col=2) for c in row if isinstance(c.value, str) and _re2.search(r"^(SIF|SIL)_\w*_(A|B|C)$", c.value)]
        if abc: bad.append(f"SIL_In 頁仍讀 A／B／C 名稱：{abc[:5]}")
        else: print("  SIL_In：無 A／B／C 名稱")
    # v0.3 Y5：SIL_In 須有逐買方份額與分世代內容名稱；內容低 ≤ 基準 ≤ 高
    if "SIL_In" in wf.sheetnames:
        sil = {wb["SIL_In"].cell(r_, 1).value: wb["SIL_In"].cell(r_, 4).value for r_ in range(1, wb["SIL_In"].max_row + 1) if isinstance(wb["SIL_In"].cell(r_, 1).value, str)}
        nsh = [k_ for k_ in sil if k_.startswith("SIF_nvsh_")]
        gens = ("H100", "GB200", "GB300", "VR200_26", "VR200", "RU")
        miss = [g_ for g_ in gens if f"SIF_NVContentGW_{g_}" not in sil]
        if len(nsh) < 33 or miss: bad.append(f"SIL_In 缺 nvsh（{len(nsh)} 名）或 NVContent（{miss}）")
        badr = [g_ for g_ in gens if not miss and not (sil[f"SIF_NVContentGW_{g_}_lo"] <= sil[f"SIF_NVContentGW_{g_}"] <= sil[f"SIF_NVContentGW_{g_}_hi"])]
        badr += [k_ for k_ in nsh if not k_.endswith(("_lo", "_hi")) and not (sil.get(k_ + "_lo", 0) <= sil[k_] <= sil.get(k_ + "_hi", 1))]
        if badr: bad.append(f"SIL 內容或份額 低 ≤ 基準 ≤ 高 不成立：{badr}")
        else: print(f"  SIL_In：nvsh {len(nsh)} 名、NVContent 6 世代齊全，低 ≤ 基準 ≤ 高")
    for cid in ("C-K1", "R-K1", "R-K2", "X4-K1", "X4-K2"):
        if cid not in [x_[0] for x_ in rows]: bad.append(f"Checks 缺 {cid}（v0.2）")
    if same:
        ps = json.load(open(same, encoding="utf-8"))["names"]
        keys = [k for k in ps if k.startswith(("WIF_MktGW_Add_A_D0_", "WIF_MktGW_YE_A_D0_", "WIF_MktGW_Add_2", "WIF_MktGW_YE_2")) and "SiteOnly" not in k]   # 主路徑（R4 讀主路徑毛額）；SiteOnly 隨站點證據可變
        dg = [k for k in keys if k in names and abs((val(k) or 0) - (ps[k] or 0)) > 1e-9]
        if dg: bad.append(f"年新增毛額或年底與前一版不同：{dg[:5]}")
        else: print(f"  毛額不變：{len(keys)} 名與 {os.path.basename(same)} 相同")
    # 6) 快照一致
    if len(sys.argv) > 2:
        snap = json.load(open(sys.argv[2], encoding="utf-8"))
        diff = [k for k, v in snap["names"].items() if k in names and val(k) != v]
        if diff: bad.append(f"快照與活頁簿不一致：{diff[:10]}")
        if snap["source"]["file"] != os.path.basename(x): bad.append("快照記錄的檔名不同")
    # 7) 固定規格（Andy 2026-10-28、2026-10-31）：README 須含對帳鏈、延遲情景與資金缺口規格
    rd = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "README.md")
    txt = open(rd, encoding="utf-8").read() if os.path.exists(rd) else ""
    for kw in ("### 對帳鏈", "| C0 |", "| C7 |", "IT 關鍵電力", "IF_FacilityGW", "k_chip2it", "下游取用 2027 年後數值前看覆蓋率", "Wattonomics（瓦帳）＝全市場 GW 唯一基準事實", "舊名 GW 帳／gwledger",
               # 資金缺口固定規格（Andy 2026-10-31 04:13、04:20、10:22）
               "資金缺口（Gap 頁", "新增 GW × 每 GW 建置成本", "不扣股票買回與股利", "整個市場當一間公司",
               # v0.1.8 固定規格（Andy 2026-11-02 15:58、16:22；工作單 20261102 E12）
               "### 三層只給基準", "### 地位鏈與對帳方向", "### 已確認延遲門檻", "公司本身揭露，或兩家以上獨立可信媒體報導且公司未否認", "〔停用，v0.2 刪〕",
               # v0.1.9（工作單 20261102 F4–F8）：NVIDIA 對帳點只並列、不校準
               "### NVIDIA 資料中心營收對帳點", "只並列、不校準",
               # v0.2（工作單 20261104）：承租玩家、融資租賃、刪停用情境、K4 乙、汰換
               "### 承租玩家只算 IT", "融資租賃走資本支出線", "出租方", "### 汰換", "K4 乙", "WIF_MktGW_<YE|Avg|Add>_<yyyy>", "A_D0 為舊名",
               # v0.2.1
               "殘差法", "攤到認列年",
               # v0.2.2
               "SIF_DeployGW_Base", "SIF_CapGW_Base",
               # v0.3
               "SIF_nvsh_", "SIF_NVContentGW_", "SIF_CapOnlineGW_Base"):
        if kw not in txt: bad.append(f"wattonomics/README.md 缺固定規格：{kw}")
    print(f"{os.path.basename(x)}：頁面 {len(wb.sheetnames)}、具名範圍 {len(names)}（WIF_ {len(gif)}）、Checks {len(rows)} 項（OK {nok}／WARN {len(cw)}／ERROR {len(ce)}／INFO {ninfo}）、"
          f"In 輸入列 {nrow}（待證據 {pend}）、錯誤格 {nerr}")
    if bad:
        print("不通過：", "；".join(bad)); sys.exit(1)
    print("通過")


if __name__ == "__main__":
    main()

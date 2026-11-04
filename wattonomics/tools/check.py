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
             if isinstance(c.value, str) and not c.value.startswith("=") and "需求 A" in c.value)
    if nA: bad.append(f"Excel 文字含「需求 A」{nA} 格（Andy 11-04 07:44）")
    for cid in ("C-K1", "R-K1"):
        if cid not in [x_[0] for x_ in rows]: bad.append(f"Checks 缺 {cid}（v0.2）")
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
               "### 承租玩家只算 IT", "融資租賃走資本支出線", "出租方", "### 汰換", "K4 乙", "WIF_MktGW_<YE|Avg|Add>_<yyyy>", "A_D0 為舊名"):
        if kw not in txt: bad.append(f"wattonomics/README.md 缺固定規格：{kw}")
    print(f"{os.path.basename(x)}：頁面 {len(wb.sheetnames)}、具名範圍 {len(names)}（WIF_ {len(gif)}）、Checks {len(rows)} 項（OK {nok}／WARN {len(cw)}／ERROR {len(ce)}／INFO {ninfo}）、"
          f"In 輸入列 {nrow}（待證據 {pend}）、錯誤格 {nerr}")
    if bad:
        print("不通過：", "；".join(bad)); sys.exit(1)
    print("通過")


if __name__ == "__main__":
    main()

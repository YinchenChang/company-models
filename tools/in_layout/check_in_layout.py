# 個股模型 In 頁版面核對（全 repo 共用；固定規格見 docs/plan/個股模型原則_命題加營收分部.md 第 8 節；
# Andy 2026-10-29 14:26 同意、15:03 指示套用到其他適用的個股）。
#
# 適用：In 頁有「低／基準／高」三欄、且有情境專屬輸入列（項目以「情境 A：」開頭、含「（情境 A）」或「情境權重 A」等）的模型。
# 規格：
#   1. In!D3／F3 欄名為「敏感度低」「敏感度高」（只用於 Sens 頁，不進目標價）。
#   2. In 頁最後有一節，A 欄文字含「情境驅動對照」：下一列表頭 D／E／F＝情境 A／B／C；其下每列一個情境驅動，三欄並排，
#      每格是 =E<列> 公式，引用 H 欄「引用 In 頁 k1／k2／k3；…」所列鍵的列（情境 B 無專屬輸入時可填常數 0）。
#   3. 上方所有情境專屬輸入列都必須被這一節引用（不能漏列）；情境權重列鍵為 scn_w，合計＝1。
# 用法：
#   python3 tools/in_layout/check_in_layout.py <成品 xlsx>      # 單檔；不符即失敗（結束碼 1）
#   python3 tools/in_layout/check_in_layout.py --all [--md 報告.md]  # 掃描全部 */dist/*.xlsx，只報告（結束碼 0）
import glob, os, re, sys
import openpyxl

SCN_PAT = re.compile(r"^情境 ?[ABC][：:]|（情境 ?[ABC]）|情境權重 ?[ABC]\b|情境權重 ?[ABC]$")


def check(path):
    """回傳 (適用, 失敗清單, 情境專屬列數)"""
    wbv = openpyxl.load_workbook(path, data_only=True)
    if "In" not in wbv.sheetnames:
        return False, [], 0
    ws = wbv["In"]
    wf = openpyxl.load_workbook(path, data_only=False)["In"]
    hdr = [str(ws.cell(3, c).value or "") for c in (4, 5, 6)]
    if hdr[1] != "基準":
        return False, [], 0
    sec = [r for r in range(4, ws.max_row + 1) if "情境驅動對照" in str(ws.cell(r, 1).value or "")]
    last = (sec[0] if sec else ws.max_row + 1)
    keyrow = {str(wf.cell(r, 1).value): r for r in range(4, last) if wf.cell(r, 1).value}
    scn_rows = {r for r in range(4, last) if SCN_PAT.search(str(ws.cell(r, 2).value or ""))}
    if not scn_rows:
        return False, [], 0
    fail = []
    if hdr[0] != "敏感度低" or hdr[2] != "敏感度高":
        fail.append(f"In!D3／F3 應為「敏感度低」「敏感度高」：{hdr[0]}／{hdr[2]}")
    if len(sec) != 1:
        fail.append(f"「情境驅動對照」一節應恰好一個：{len(sec)}")
        return True, fail, len(scn_rows)
    h = sec[0] + 1
    if [str(ws.cell(h, c).value or "")[:4] for c in (4, 5, 6)] != ["情境 A", "情境 B", "情境 C"]:
        fail.append(f"情境驅動對照表頭 D／E／F 應為情境 A／B／C：{[ws.cell(h, c).value for c in (4, 5, 6)]}")
    rows, refd = {}, set()
    r = h + 1
    while ws.cell(r, 1).value:
        k = ws.cell(r, 1).value
        v = [ws.cell(r, c).value for c in (4, 5, 6)]
        rows[k] = v
        if not all(isinstance(x, (int, float)) for x in v):
            fail.append(f"情境驅動 {k} 三欄應皆為數值：{v}")
        m = re.match(r"引用 In 頁 ([^；]+)", str(wf.cell(r, 8).value or ""))
        keys = m.group(1).split("／") if m else []
        fs = [wf.cell(r, c).value for c in (4, 5, 6)]
        refs = [int(f[2:]) for f in fs if isinstance(f, str) and re.fullmatch(r"=E\d+", f)]
        consts = [j for j, f in enumerate(fs) if not isinstance(f, str)]
        if [keyrow.get(x) for x in keys] != refs or any(j != 1 or fs[j] != 0 for j in consts):
            fail.append(f"情境驅動 {k} 的公式 {fs} 與 H 欄所列鍵 {keys} 不符")
        refd.update(refs)
        r += 1
    if "scn_w" not in rows:
        fail.append("缺情境權重列 scn_w")
    elif all(isinstance(x, (int, float)) for x in rows["scn_w"]) and abs(sum(rows["scn_w"]) - 1) > 1e-9:
        fail.append(f"情境權重合計應為 1：{rows['scn_w']}")
    miss = sorted(scn_rows - refd)
    if miss:
        fail.append("情境專屬輸入列未列入對照：" + "、".join(f"{wf.cell(x, 1).value}（第 {x} 列）" for x in miss))
    return True, fail, len(scn_rows)


def main(argv):
    if argv and argv[0] == "--all":
        md = argv[argv.index("--md") + 1] if "--md" in argv else None
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        res = []
        for d in sorted(glob.glob(os.path.join(root, "*", "dist"))):
            xs = sorted(x for x in glob.glob(os.path.join(d, "*.xlsx")))
            if not xs:
                continue
            x = xs[-1]
            try:
                app, fail, n = check(x)
            except Exception as e:  # 讀檔失敗也列出
                app, fail, n = True, [f"讀檔失敗：{e}"], 0
            co = os.path.basename(os.path.dirname(d))
            if app:
                res.append((co, os.path.basename(x), n, "符合" if not fail else "不符", "；".join(fail)[:200]))
        lines = ["| 資料夾 | 成品 | 情境專屬列 | 結果 | 說明 |", "|---|---|---|---|---|"] + [f"| `{a}` | {b} | {c} | {d} | {e} |" for a, b, c, d, e in res]
        print("\n".join(lines))
        print(f"\n適用 {len(res)} 檔；符合 {sum(1 for x in res if x[3] == '符合')}、不符 {sum(1 for x in res if x[3] != '符合')}")
        if md:
            open(md, "w", encoding="utf-8").write("\n".join(lines) + "\n")
        return 0
    app, fail, n = check(argv[0])
    if not app:
        print("In 頁版面核對：不適用（無 In 頁三欄或無情境專屬輸入列）")
        return 0
    if fail:
        print("In 頁版面核對失敗：")
        for f in fail:
            print("  -", f)
        return 1
    print(f"In 頁版面核對通過（情境專屬輸入 {n} 列皆已列入對照）")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))

# MAG 個股模型 In 頁版面核對（v0.3.2 起固定規格，Andy 2026-10-29 14:26 同意）：
#   1. In!D3／F3 欄名為「敏感度低」「敏感度高」（只用於 Sens 頁，不進目標價）。
#   2. 頁尾有「J. 情境驅動對照」一節：表頭 D／E／F＝情境 A／B／C，三情境並排列出所有情境驅動與情境權重（鍵 scn_w，合計＝1）。
# 用法：python3 mag/check_in_layout.py <成品 xlsx>（需已重算；make_v03.sh 會呼叫）
#   3. 情境驅動對照每格是 =E<列> 公式，且引用的列＝H 欄「引用 In 頁 k1／k2…」所列鍵的列（常數 0 只允許在 B 欄）。
import re, sys, openpyxl
ws = openpyxl.load_workbook(sys.argv[1], data_only=True)["In"]
wf = openpyxl.load_workbook(sys.argv[1], data_only=False)["In"]
KEYROW = {str(wf.cell(r, 1).value): r for r in range(4, wf.max_row + 1) if wf.cell(r, 1).value}
fail = []
if ws["D3"].value != "敏感度低" or ws["F3"].value != "敏感度高":
    fail.append(f"In!D3／F3 應為「敏感度低」「敏感度高」：{ws['D3'].value}／{ws['F3'].value}")
sec = [r for r in range(4, ws.max_row + 1) if str(ws.cell(r, 1).value or "").startswith("J. 情境驅動對照")]
if len(sec) != 1:
    fail.append(f"「J. 情境驅動對照」一節應恰好一個：{len(sec)}")
else:
    h = sec[0] + 1
    if [ws.cell(h, c).value for c in (4, 5, 6)] != ["情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）"]:
        fail.append(f"情境驅動對照表頭 D／E／F 應為情境 A／B／C：{[ws.cell(h, c).value for c in (4, 5, 6)]}")
    rows = {}
    r = h + 1
    while ws.cell(r, 1).value:
        rows[ws.cell(r, 1).value] = [ws.cell(r, c).value for c in (4, 5, 6)]
        m = re.match(r"引用 In 頁 ([^；]+)", str(wf.cell(r, 8).value or ""))
        keys = m.group(1).split("／") if m else []
        fs = [wf.cell(r, c).value for c in (4, 5, 6)]
        refs = [int(f[2:]) for f in fs if isinstance(f, str) and re.fullmatch(r"=E\d+", f)]
        consts = [j for j, f in enumerate(fs) if not isinstance(f, str)]
        if [KEYROW.get(k) for k in keys] != refs or any(j != 1 or fs[j] != 0 for j in consts):
            fail.append(f"情境驅動 {ws.cell(r, 1).value} 的公式 {fs} 與 H 欄所列鍵 {keys} 不符")
        r += 1
    if len(rows) < 2:
        fail.append("情境驅動對照至少要有一個驅動列＋情境權重列")
    for k, v in rows.items():
        if not all(isinstance(x, (int, float)) for x in v):
            fail.append(f"情境驅動 {k} 三欄應皆為數值：{v}")
    if "scn_w" not in rows:
        fail.append("缺情境權重列 scn_w")
    elif all(isinstance(x, (int, float)) for x in rows["scn_w"]) and abs(sum(rows["scn_w"]) - 1) > 1e-9:
        fail.append(f"情境權重合計應為 1：{rows['scn_w']}")
if fail:
    print("In 頁版面核對失敗："); [print("  -", f) for f in fail]; sys.exit(1)
print(f"In 頁版面核對通過（情境驅動 {len(rows) - 1} 列＋權重）")

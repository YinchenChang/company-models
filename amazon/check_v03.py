# Amazon v0.3.1 成品核對：公式錯誤 0、核對列（Q、Seg、PL、CF、Val、Recon，含三情境）皆 0、敏感度未過期、C10 對帳≈0、固定列標籤與標記齊全。
# 用法：python3 amazon/check_v03.py <重算後的 xlsx>；失敗時非零結束。
import re
import sys
import openpyxl

X = sys.argv[1]
wb = openpyxl.load_workbook(X, data_only=True)
fail = []
ERR = ("#REF", "#DIV", "#VALUE", "#NAME", "#N/A", "#NUM", "Err:")
n_err = 0
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith(ERR):
                n_err += 1
                if n_err <= 10: fail.append(f"公式錯誤 {ws.title}!{c.coordinate} {c.value}")
print(f"公式錯誤：{n_err} 格")
need_sheets = ["README", "In", "Q", "Thesis", "Seg", "PL", "CF", "Val", "Recon", "Sens", "Sources"]
if wb.sheetnames != need_sheets:
    fail.append(f"工作表順序不符：{wb.sheetnames}")


def nums(ws, r, c0=2):
    return [c.value for c in ws[r][c0 - 1:] if isinstance(c.value, (int, float)) and not isinstance(c.value, bool)]


n_chk = 0
for sh in ("Q", "Seg", "PL", "CF", "Val", "Recon", "Sens"):
    ws = wb[sh]
    for r in range(1, ws.max_row + 1):
        lab = ws.cell(r, 1).value
        if not (isinstance(lab, str) and lab.startswith("檢查")): continue
        vs = nums(ws, r, 3 if sh in ("Q", "Seg", "PL", "Sens") else 2)
        if sh == "Sens":
            vs = nums(ws, r, 3)
        n_chk += 1
        if any(abs(v) > 1e-6 for v in vs):
            fail.append(f"{sh}!A{r}「{lab}」有非零值 {[round(v, 6) for v in vs if abs(v) > 1e-6][:5]}")
print(f"核對列：{n_chk} 列（Q、Seg、PL、CF、Val、Recon、Sens；含三情境）")
V, CF, IN, SE, RC = wb["Val"], wb["CF"], wb["In"], wb["Sens"], wb["Recon"]
labs = {V.cell(r, 1).value: r for r in range(1, V.max_row + 1)}
for n in ("目標價（DCF 與 Comparable 加權）", "上檔空間：加權目標價", "投資建議", "上檔空間級距", "DCF 目標價（情境加權）", "Comparable 目標價（市場主流同業法）",
          "價格帶（DCF 不如預期～順利）", "現價（2026-10-30 收盤）", "市場共識目標價", "讀法 1：加權目標價", "FY2028 AWS 營業利益率"):
    if n not in labs: fail.append(f"Val 缺列「{n}」")
if CF["A6"].value != "收支關卡：自由現金流扣股利（五年合計）":
    fail.append("CF!A6 應為「收支關卡：自由現金流扣股利（五年合計）」")
# 敏感度未過期
st = [r for r in range(1, SE.max_row + 1) if str(SE.cell(r, 1).value or "").startswith("敏感度重算時的基準值")]
if not st or not isinstance(SE.cell(st[0], 5).value, (int, float)) or abs(SE.cell(st[0], 5).value) > 1e-6:
    fail.append(f"Sens 敏感度已過期或缺：基準差 {SE.cell(st[0], 5).value if st else None}")
# Recon 拆解快照有效
for r in range(1, RC.max_row + 1):
    if RC.cell(r, 2).value == "快照過期": fail.append(f"Recon!B{r} 拆解快照過期")
# C10 對帳（情境 B）
c10 = [r for r in range(1, RC.max_row + 1) if str(RC.cell(r, 1).value or "").startswith("C10 對帳")]
if not c10 or abs(RC.cell(c10[0], 4).value or 9) > 0.01:
    fail.append(f"C10 對帳（情境 B）落差 {RC.cell(c10[0], 4).value if c10 else None} > 0.01")
# In 頁：標記齊全、Analogy／Assumed 有區間或為開關／判斷
TAGS = {"Verified", "Interested-party", "Analogy", "Assumed", "Derived"}
sec_j = next((r for r in range(4, IN.max_row + 1) if "情境驅動對照" in str(IN.cell(r, 1).value or "")), IN.max_row + 1)
n_in = 0
for r in range(4, sec_j):
    k = IN.cell(r, 1).value
    if not k or not re.fullmatch(r"[A-Za-z0-9_]+", str(k)): continue
    n_in += 1
    tag = str(IN.cell(r, 7).value or "")
    if not any(t in tag for t in TAGS): fail.append(f"In!G{r}（{k}）缺標記")
    if not IN.cell(r, 8).value: fail.append(f"In!H{r}（{k}）缺來源")
print(f"In 頁輸入：{n_in} 鍵（標記、來源齊全檢查）")
tp = V.cell(labs["目標價（DCF 與 Comparable 加權）"], 2).value
up = V.cell(labs["上檔空間：加權目標價"], 2).value
rec = V.cell(labs["投資建議"], 2).value
print(f"加權目標價 {tp:.2f}；上檔空間 {up:+.1%}；投資建議 {rec}")
if fail:
    print("失敗：")
    for f in fail: print("  -", f)
    sys.exit(1)
print("全部通過")

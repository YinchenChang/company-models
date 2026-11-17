# AppLovin 成品核對：公式錯誤 0、各頁「檢查：」列為 0、敏感度未過期、固定列標籤存在、反解快照未過期。
# 用法：python3 software_ai/applovin/check_v02.py <重算後的 xlsx>；失敗時非零結束。
import sys
import openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
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
n_chk = {}
for sh in ("PL", "CF", "Thesis", "Val"):
    ws = wb[sh]; n_chk[sh] = 0
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if isinstance(a, str) and a.startswith("檢查："):
            vs = [c.value for c in ws[r][1:] if isinstance(c.value, (int, float)) and not isinstance(c.value, bool)]
            n_chk[sh] += 1
            if "應大於 0" in a:
                if not vs or vs[0] <= 0: fail.append(f"{sh}!A{r} 應大於 0：{vs}")
                continue
            if not vs: fail.append(f"{sh}!A{r}「{a}」沒有數值")
            elif any(abs(v) > 1e-6 for v in vs): fail.append(f"{sh}!A{r}「{a}」有非零值 {[round(v, 6) for v in vs if abs(v) > 1e-6][:5]}")
print("核對列：" + "、".join(f"{k} {v} 列" for k, v in n_chk.items()))
if n_chk["PL"] < 12 or n_chk["CF"] < 3: fail.append(f"核對列過少 {n_chk}")
V = wb["Val"]
labs = {V.cell(r, 1).value: r for r in range(1, V.max_row + 1)}
for n in ("目標價（DCF 與 Comparable 加權）", "上檔空間：加權目標價", "投資建議", "估值評等（五級）", "入不敷出判定用：收支關卡合計（情境 B）", "DCF 目標價（情境加權；12 個月後）", "Comparable 目標價（情境加權；同業 P/S 法）"):
    if n not in labs: fail.append(f"Val 缺列「{n}」")
se = wb["Sens"]
for r in range(1, se.max_row + 1):
    a = str(se.cell(r, 1).value or "")
    if a.startswith("敏感度重算時的基準值"):
        d = se.cell(r, 5).value
        if not isinstance(d, (int, float)) or abs(d) > 1e-6: fail.append(f"Sens 敏感度已過期或缺：基準差 {d}")
    if a.startswith("模型狀態") and se.cell(r, 4).value != "CURRENT": fail.append(f"Sens 模型狀態 {se.cell(r, 4).value}")
rc = wb["Recon"]
for r in range(1, rc.max_row + 1):
    if str(rc.cell(r, 1).value or "").startswith("反解快照檢查"):
        d = rc.cell(r, 2).value
        if not isinstance(d, (int, float)) or abs(d) > 0.01: fail.append(f"Recon 反解快照過期或缺：{d}")
ri = labs.get("任一關卡成立？（1＝是；須先診斷）")
flag = V.cell(ri, 2).value if ri else None
tp, up, rec = V.cell(labs["目標價（DCF 與 Comparable 加權）"], 2).value, V.cell(labs["上檔空間：加權目標價"], 2).value, V.cell(labs["投資建議"], 2).value
print(f"加權目標價 {tp:.2f}；上檔空間 {up:+.1%}；投資建議 {rec}；合理性關卡旗標 {flag}")
if flag not in (0, None):
    rd_ = wb["README"]
    if not any(str(rd_.cell(r, 1).value or "").startswith("診斷結論") for r in range(1, rd_.max_row + 1)): fail.append("合理性關卡成立：README 缺「診斷結論」")
    else: print("合理性關卡成立；README 已有診斷結論")
if fail:
    print("失敗：")
    for f in fail: print("  -", f)
    sys.exit(1)
print("全部通過")

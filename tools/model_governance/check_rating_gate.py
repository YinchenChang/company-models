# 收支關卡評等檢查（個股模型原則第 10 節第 63 條；Andy 2026-10-25、10-26、2026-11-13）：
# 收支關卡不通過（基準情境期間現金收支合計 < 0）時，個股模型自己的「投資建議」不得為買進或強力買進，須為「收支不足」。
# 用法：
#   python3 tools/model_governance/check_rating_gate.py <成品.xlsx> [--tol 50] [--strict]
#       單一模型；不符合時以非 0 結束。--strict：找不到「投資建議」或「收支關卡」列也算失敗（已套用第 63 條的模型在建置腳本內一律加）。
#   python3 tools/model_governance/check_rating_gate.py --map markets/invest_map/dist/data.json
#       盤點：列出地圖判定收支關卡不通過、而模型估值評等為買進以上的公司（這些模型下次改版必須套用第 63 條）
# 讀法：
#   投資建議＝A 欄以「投資建議」開頭的第一列（Val 頁優先），取其後第一個文字值。
#   收支合計＝A 欄以「收支關卡」開頭的列的數值，Fund 頁優先、其次 CF 頁、再其次其他頁（3 個數值視為情境 A／B／C，取中間的 B；1 個取該值）。
#   估值評等＝A 欄以「估值評等（」開頭的第一列（有此列時做一致性核對）；Val 頁「入不敷出判定用」列的值須等於收支合計。
# 判定：收支合計為負且估值評等（無此列時用投資建議）為買進以上 → 投資建議必須恰為「收支不足」；
#       其餘情況投資建議必須等於估值評等，且只能是五級之一。
import json, sys, openpyxl

BUY = ("強力買進", "買進")
TIERS = ("強力買進", "買進", "中立", "放空", "強力放空")
PREF = {"Fund": 0, "CF": 1}


def read(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    recs, gates, val, vgate = [], [], None, None
    for ws in wb:
        for row in ws.iter_rows(values_only=True):
            a = row[0]
            if not isinstance(a, str): continue
            txt = [v.strip() for v in row[1:] if isinstance(v, str) and v.strip() and v.strip() != "—"]
            num = [v for v in row[1:] if isinstance(v, (int, float)) and not isinstance(v, bool)]
            if a.startswith("投資建議") and txt: recs.append((0 if ws.title == "Val" else 1, ws.title, txt[0]))
            if a.startswith("收支關卡") and num: gates.append((PREF.get(ws.title, 2), ws.title, a, num[1] if len(num) >= 3 else num[0]))
            if val is None and a.startswith("估值評等（") and txt: val = txt[0]
            if vgate is None and a.startswith("入不敷出判定用") and num: vgate = num[0]
    rec = min(recs)[1:] if recs else None
    gate = min(gates, key=lambda g: g[0])[1:] if gates else None
    return rec, gate, val, vgate


def check(path, tol=0.0):
    rec, gate, val, vgate = read(path)
    if rec is None: return None, "找不到「投資建議」列（或檔案未重算、沒有快取值）"
    if gate is None: return None, f"不適用：沒有「收支關卡」列（投資建議 {rec[1]}）"
    r, g = rec[1], gate[2]
    fail = g < -abs(tol)
    # 尚未套用第 63 條的模型（沒有估值評等列）：投資建議可能是條件式句子，以是否含「買進」判斷
    base = val if val is not None else ("買進" if "買進" in r else r)
    if vgate is not None and abs(vgate - g) > 1e-6:
        return False, f"不符合：Val 頁入不敷出判定用的合計 {vgate:.4f} 與 {gate[0]} 頁收支關卡 {g:.4f} 不同"
    if val is not None and val not in TIERS:
        return False, f"不符合：估值評等「{val}」不是五級之一"
    if fail and base in BUY:
        if r != "收支不足":
            return False, f"不符合第 63 條：收支合計 {g:.1f}（{gate[0]} 頁）為負、估值評等 {base}，投資建議卻是「{r}」，應為「收支不足」"
    elif val is not None and r != val:
        return False, f"不符合：收支關卡{'不通過' if fail else '通過'}、估值評等 {val}，投資建議應等於估值評等，卻是「{r}」"
    elif val is None and not any(t in r for t in TIERS + ("收支不足",)):
        return False, f"不符合：投資建議「{r}」不含五級評等（收支合計 {g:.1f}）"
    return True, f"符合：投資建議 {r}" + (f"（估值評等 {val}）" if val is not None else "") + f"；收支合計 {g:.1f}（{gate[0]} 頁，{'不通過' if fail else '通過'}）"


def inventory(data_json):
    d = json.load(open(data_json, encoding="utf-8"))
    hit = [x for x in d["valuation"] if "不通過" in str(x.get("gate")) and (x.get("mtier") in BUY or x.get("vtier") in BUY)]
    print(f"地圖判定收支關卡不通過、而估值評等為買進以上：{len(hit)} 家")
    for x in hit:
        print(f"  {x['tk']} {x['name']}：估值評等 {x.get('vtier')}、期間合計 {x.get('gsum'):.1f}（{x.get('gbasis')}）、模型檔 {x.get('mfile')}")
    return 0


if __name__ == "__main__":
    a = sys.argv[1:]
    if not a: sys.exit("用法見檔頭")
    if a[0] == "--map": sys.exit(inventory(a[1]))
    tol = float(a[a.index("--tol") + 1]) if "--tol" in a else 0.0
    ok, msg = check(a[0], tol)
    print("收支關卡評等檢查：" + msg)
    sys.exit(1 if (ok is False or (ok is None and "--strict" in a)) else 0)

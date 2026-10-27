# 把工作單的輸入變更寫進現行 Excel 的 In 頁（Excel 優先：輸入改在 Excel，builder 再以 --base 讀回）。
# 用法：python3 siliconomics/tools/apply_inputs.py <現行 xlsx> <變更 json> <輸出 xlsx>
# 變更 json：[{"id": 鍵, "lo": 值, "base": 值, "hi": 值, "tag": 文字, "src": 文字}, ...]，只寫出現的欄位；
#   鍵不存在時報錯（新增輸入列請放在 builder/seed/rows_added.json，由 builder 建立）。
# 只在「仍為舊值」時寫入（"old_base" 欄若給，基準不同即報錯），避免覆寫 Excel 上已改過的值。
import json, sys
import openpyxl
from openpyxl.styles import Font

COL = {"lo": 4, "base": 5, "hi": 6, "tag": 7, "src": 8}
BLUE = Font(color="0000FF")


def main():
    src, edits, out = sys.argv[1:4]
    wb = openpyxl.load_workbook(src)
    ws = wb["In"]
    rows = {ws.cell(r, 1).value: r for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value}
    log = []
    for e in json.load(open(edits, encoding="utf-8")):
        r = rows.get(e["id"])
        if r is None:
            sys.exit(f"In 頁找不到鍵 {e['id']}")
        if "old_base" in e and ws.cell(r, 5).value != e["old_base"]:
            sys.exit(f"{e['id']} 基準已不是 {e['old_base']}（現值 {ws.cell(r, 5).value}），停止")
        for k, c in COL.items():
            if k in e:
                old = ws.cell(r, c).value
                ws.cell(r, c).value = e[k]
                if k in ("lo", "base", "hi"):
                    ws.cell(r, c).font = BLUE
                log.append(f"{e['id']} {k}: {old!r} → {e[k]!r}")
    wb.save(out)
    print("\n".join(log))
    print(f"寫入 {len(log)} 格 → {out}")


if __name__ == "__main__":
    main()

# GW 帳取數工具：把重算後活頁簿的 Interface（GIF_ 具名範圍）、GGOV_ 與 In 頁輸入列匯成版本固定的快照 JSON。
# 用法：
#   匯出：python3 gwledger/tools/export_snapshot.py <重算後的 GWLedger xlsx> <輸出 snapshot.json>
#   檢查：python3 gwledger/tools/export_snapshot.py --check <snapshot.json> [<xlsx>]
#         重新讀取快照記錄的 xlsx（預設 gwledger/model/<快照記錄的檔名>），任何值差異即以代碼 1 結束。
# 下游（Siliconomics Deploy、公司模型、AI 投資地圖）只讀快照，不直接讀 xlsx；快照記錄檔名、版本、sha256。
import hashlib, json, os, sys
import openpyxl

GL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _cells(wb, ref):
    sh, rng = ref.split("!")
    ws = wb[sh.strip("'")]
    rng = rng.replace("$", "")
    if ":" in rng:
        return [c.value for row in ws[rng] for c in row]
    return ws[rng].value


def read(xlsx):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    names = {}
    for nm, dn in wb.defined_names.items():
        if nm.startswith("GIF_") or nm.startswith("GGOV_"):
            v = _cells(wb, dn.attr_text)
            names[nm] = None if v == "" else v
    ws = wb["In"]
    hdr = [ws.cell(4, c).value for c in range(1, 18)]
    rows = []
    for r in range(5, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is None:
            continue
        if not ws.cell(r, 19).value:
            if ws.cell(r, 2).value is None:
                rows.append({"sec": a})
            continue
        rows.append(dict(id=a, label=ws.cell(r, 2).value, basis=ws.cell(r, 3).value, unit=ws.cell(r, 4).value, tag=ws.cell(r, 5).value,
                         src=ws.cell(r, 6).value, note=ws.cell(r, 7).value or "", kind=ws.cell(r, 19).value,
                         values=[ws.cell(r, c).value for c in range(8, 18)]))
    return names, rows, hdr


def main():
    if sys.argv[1] == "--check":
        snap = json.load(open(sys.argv[2], encoding="utf-8"))
        x = sys.argv[3] if len(sys.argv) > 3 else os.path.join(GL, "model", snap["source"]["file"])
        names, rows, _ = read(x)
        bad = [k for k in snap["names"] if names.get(k) != snap["names"][k]]
        bad += ["rows"] if rows != snap["rows"] else []
        if bad:
            print("快照與活頁簿不一致：", bad[:20]); sys.exit(1)
        print(f"快照與 {os.path.basename(x)} 一致（{len(names)} 個名稱）"); return
    x, out = sys.argv[1], sys.argv[2]
    names, rows, _ = read(x)
    if names.get("GGOV_Errors"):
        print("GGOV_Errors ≠ 0，拒絕匯出"); sys.exit(1)
    fn = os.path.basename(x)
    ver = fn.split("_v")[-1].rsplit(".xlsx", 1)[0] if "_v" in fn else "?"
    snap = dict(source=dict(file=fn, version="v" + ver, sha256=hashlib.sha256(open(x, "rb").read()).hexdigest(),
                            unit="TK GW（設施電力）", years=list(range(2024, 2033)),
                            note="GW 帳快照；下游只讀本檔。GGOV_Errors 須為 0；名稱格式 <前綴>_<yyyy>；空白＝未知"),
                rows=rows, names=names)
    json.dump(snap, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"快照：{out}（{fn}；{len(names)} 個名稱、{sum(1 for r in rows if 'id' in r)} 個輸入列）")


if __name__ == "__main__":
    main()

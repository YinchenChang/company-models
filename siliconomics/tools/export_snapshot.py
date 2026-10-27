# Siliconomics 取數工具：把重算後活頁簿的 Interface（SIF_ 具名範圍）與 In 頁輸入列匯成版本固定的快照 JSON。
# 用法：
#   匯出：python3 siliconomics/tools/export_snapshot.py <重算後的 Siliconomics xlsx> <輸出 snapshot.json>
#   檢查：python3 siliconomics/tools/export_snapshot.py --check <snapshot.json> [<xlsx>]
#         重新讀取快照記錄的 xlsx（預設 siliconomics/model/<快照記錄的檔名>），任何值差異即以代碼 1 結束。
# 下游（markets/memory/market.py 等）只讀快照，不直接讀 xlsx；快照記錄檔名、版本、sha256。
import hashlib, json, os, sys
import openpyxl

SIL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


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
        if nm.startswith("SIF_") or nm.startswith("SGOV_"):
            names[nm] = _cells(wb, dn.attr_text)
    # In 頁輸入列（依序，含節標題）與平台分組：供 market.py 重播 inputs()
    ws = wb["In"]
    rows, groups = [], []
    for r in range(5, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is None:
            continue
        if isinstance(a, str) and a.startswith("grp_"):
            groups.append(dict(key=a[4:], name=ws.cell(r, 2).value, tag=ws.cell(r, 7).value, src=ws.cell(r, 8).value or ""))
            continue
        if str(a).startswith("M2 分組") or a == "分組鍵":
            continue
        if ws.cell(r, 3).value is None and ws.cell(r, 5).value is None:
            rows.append({"sec": a})
            continue
        rows.append(dict(id=a, label=ws.cell(r, 2).value, unit=ws.cell(r, 3).value, lo=ws.cell(r, 4).value,
                         base=ws.cell(r, 5).value, hi=ws.cell(r, 6).value, tag=ws.cell(r, 7).value, src=ws.cell(r, 8).value or "",
                         src2=ws.cell(r, 10).value or ""))
    sc = wb["Scen"]
    scen, desc = [], {}
    for r in range(4, 7):
        s = sc.cell(r, 1).value
        scen.append(s); desc[s] = sc.cell(r, 4).value
    return names, rows, groups, scen, desc


def main():
    if sys.argv[1] == "--check":
        snap = json.load(open(sys.argv[2], encoding="utf-8"))
        x = sys.argv[3] if len(sys.argv) > 3 else os.path.join(SIL, "model", snap["source"]["file"])
        names, rows, groups, scen, desc = read(x)
        bad = [k for k in snap["names"] if names.get(k) != snap["names"][k]]
        bad += ["rows"] if rows != snap["rows"] else []
        if bad:
            print("快照與活頁簿不一致：", bad[:20]); sys.exit(1)
        print(f"快照與 {os.path.basename(x)} 一致（{len(names)} 個名稱）"); return
    x, out = sys.argv[1], sys.argv[2]
    names, rows, groups, scen, desc = read(x)
    fn = os.path.basename(x)
    ver = fn.split("_v")[-1].rsplit(".xlsx", 1)[0] if "_v" in fn else "?"
    snap = dict(source=dict(file=fn, version="v" + ver, sha256=hashlib.sha256(open(x, "rb").read()).hexdigest(),
                            note="Siliconomics 快照；下游只讀本檔。SGOV_Errors 須為 0"),
                scen=scen, scen_desc=desc, unit_groups=groups, rows=rows, names=names)
    if names.get("SGOV_Errors"):
        print("SGOV_Errors ≠ 0，拒絕匯出"); sys.exit(1)
    json.dump(snap, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"快照：{out}（{fn}；{len(names)} 個名稱、{sum(1 for r in rows if 'id' in r)} 個輸入列）")


if __name__ == "__main__":
    main()

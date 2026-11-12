# 個股模型 repo README 版本表：由建置成品（dist 內 xlsx）產生，不手寫（原則第 10 節第 51 條；首例 NVIDIA v0.3.5）。
# 設定：<公司>/display_check.json 的 "table_rows"（每列：標題、工作表、標籤前綴、欄、格式；可用 "after" 指定先找到的區塊標籤）。
# README.md 內以標記包住表格：
#   <!-- version-table: <xlsx 路徑 1>｜<xlsx 路徑 2>… -->
#   （本工具產生的表格）
#   <!-- /version-table -->
# 路徑相對於公司資料夾（例：dist/_待刪除/20261106_NVIDIA個股模型_v0.3.4.xlsx｜dist/20261111_NVIDIA個股模型_v0.3.5.xlsx），欄名取檔名中的版號。
# 用法：python3 tools/model_governance/version_table.py <公司資料夾> [--update]（不加 --update 時只檢查：表格與成品不一致即結束碼 1）
import json, os, re, sys
import openpyxl

MARK = re.compile(r"<!-- version-table: (.+?) -->\n(.*?)<!-- /version-table -->", re.S)


def fmt(v, f):
    if v is None: return "—"
    if isinstance(v, str): return v
    if f == "n2": return f"{v:,.2f}"
    if f == "n1": return f"{v:,.1f}"
    if f == "n3": return f"{v:,.3f}"
    if f == "p1": return f"{v * 100:.1f}%"
    if f == "p2": return f"{v * 100:.2f}%"
    if f == "sp1": return f"{v * 100:+.1f}%"
    return str(v)


def find_row(ws, prefix, start=1):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip().startswith(prefix): return r
    raise KeyError((ws.title, prefix))


def cell_value(wb, spec):
    ws = wb[spec["sheet"]]
    r0 = find_row(ws, spec["after"]) if spec.get("after") else 1
    r = find_row(ws, spec["find"], r0)
    vals = [ws[f"{c}{r}"].value for c in spec["col"].split(",")]
    return vals, f"{spec['sheet']}!{','.join(c + str(r) for c in spec['col'].split(','))}"


def render_cell(wb, spec):
    vals, _ = cell_value(wb, spec)
    fs = spec.get("fmt", "n2").split(",")
    parts = [fmt(v, fs[min(i, len(fs) - 1)]) for i, v in enumerate(vals)]
    s = "／".join(parts)
    if spec.get("with"):   # 括號附註另一格（例：加權目標價（上檔空間））
        w = dict(spec["with"]); v2, _ = cell_value(wb, w)
        s += f"（{fmt(v2[0], w.get('fmt', 'sp1'))}）"
    return s


def render(folder, paths, rows):
    wbs = [openpyxl.load_workbook(os.path.join(folder, p), data_only=True) for p in paths]
    vers = [re.search(r"_v(\d+(?:\.\d+)+)\.xlsx$", p).group(1) for p in paths]
    out = ["| 項目 | " + " | ".join(f"v{v}" for v in vers) + " |", "|---|" + "---|" * len(vers)]
    for spec in rows:
        out.append(f"| {spec['title']} | " + " | ".join(render_cell(wb, spec) for wb in wbs) + " |")
    return "\n".join(out) + "\n"


def main(argv):
    folder = argv[0]; update = "--update" in argv
    cfg = json.load(open(os.path.join(folder, "display_check.json"), encoding="utf-8"))
    md_path = os.path.join(folder, "README.md"); md = open(md_path, encoding="utf-8").read()
    bad = []

    def sub(m):
        paths = [p.strip() for p in m.group(1).split("｜")]
        want = render(folder, paths, cfg["table_rows"])
        if m.group(2) != want: bad.append(m.group(1))
        return f"<!-- version-table: {m.group(1)} -->\n{want}<!-- /version-table -->"
    new = MARK.sub(sub, md)
    n = len(MARK.findall(md))
    if update:
        open(md_path, "w", encoding="utf-8").write(new)
        print(f"版本表：更新 {len(bad)}／{n} 個")
        return 0
    if n == 0:
        print("版本表：README.md 沒有 version-table 標記"); return 1
    print(f"版本表：{n} 個，" + ("全部與成品一致" if not bad else f"{len(bad)} 個與成品不一致：" + "；".join(bad)))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))

# 展示層數字一致性檢查（原則第 10 節第 51 條；2026-11-11 NVIDIA v0.3.4 驗收衍生，首例 NVIDIA v0.3.5）。
# 把 README 工作表文字、HTML、repo <公司>/README.md 版本表中的關鍵數字對回來源儲存格，不一致即結束碼 1。
# 設定：<公司>/display_check.json
#   "items"：[{name, sheet, find（A 欄標籤前綴）, after（可選：先找到的區塊標籤）, col, fmt（n2／p1／sp1／x2／text）, patterns（可選：README 頁一致性正規式）}]
#   "history_re"：README 頁版本沿革列的開頭（凍結為當時文字，不檢查）
#   "table_rows"：repo README 版本表的列（tools/model_governance/version_table.py）
# 檢查：
#   1. README 頁：每個 item 的格式化值出現在非沿革列；每個 pattern 在非沿革列的所有擷取值都等於格式化值（例：手寫舊目標價即失敗）。
#   2. HTML：data-ref 指向來源儲存格的數字，data-v 與來源值相同、顯示文字等於格式化值；格式化值出現在 HTML；
#      HTML 中所有 README 頁文字（data-ref="README!A…"）與 xlsx README 頁快取值逐字相同。
#   3. repo README.md：所有版本表與成品重算一致（version_table.py），且至少一個版本表的最後一欄是目前 dist 內的 xlsx。
# 用法：python3 tools/model_governance/check_display.py <公司資料夾>（讀 dist 內唯一的 xlsx 與同名 html）
import glob, html as H, json, os, re, sys
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import version_table as VT  # noqa: E402


def fmt(v, f):
    if f == "x2": return f"{v:.2f}"
    if f == "x1": return f"{v:.1f}"
    if f == "text": return str(v)
    return VT.fmt(v, f).replace(",", "") if f.startswith("n") else VT.fmt(v, f)


def main(folder):
    cfg = json.load(open(os.path.join(folder, "display_check.json"), encoding="utf-8"))
    xs = glob.glob(os.path.join(folder, "dist", "*.xlsx")); assert len(xs) == 1, xs
    X = xs[0]; HT = X[:-5] + ".html"
    wb = openpyxl.load_workbook(X, data_only=True)
    hist = re.compile(cfg.get("history_re", r"^v\d"))
    RD = wb["README"]
    rtext = [str(RD.cell(r, c).value) for r in range(1, RD.max_row + 1) for c in range(1, 4)
             if RD.cell(r, c).value is not None and not hist.match(str(RD.cell(r, 1).value or ""))]
    blob = "\n".join(rtext)
    page = open(HT, encoding="utf-8").read() if os.path.exists(HT) else None
    fail = []; n_ok = 0
    for it in cfg["items"]:
        vals, ref = VT.cell_value(wb, it)
        v = vals[0]; s = fmt(v, it.get("fmt", "n2"))
        if s not in blob: fail.append(f"README 頁未出現 {it['name']} {s}（{ref}）")
        for pat in it.get("patterns", []):
            for m in re.finditer(pat, blob):
                if m.group(1) != s: fail.append(f"README 頁 {it['name']} 寫 {m.group(1)}，來源 {ref}＝{s}（「{m.group(0)[:40]}」）")
                else: n_ok += 1
        if page is not None:
            cell = ref.split(",")[0]
            for m in re.finditer(r'data-ref="' + re.escape(cell) + r'" data-v="([^"]*)"[^>]*>([^<]*)<', page):
                dv = float(m.group(1)) if re.match(r"^-?[\d.eE+-]+$", m.group(1)) else m.group(1)
                if isinstance(v, (int, float)) and (not isinstance(dv, float) or abs(dv - v) > 1e-9 * max(1, abs(v))):
                    fail.append(f"HTML {cell} data-v {m.group(1)} ≠ 來源 {v}")
                n_ok += 1
            if s not in H.unescape(re.sub(r"<[^>]+>", "", page)): fail.append(f"HTML 未出現 {it['name']} {s}")
    if page is not None:   # HTML 內的 README 頁文字須與 xlsx 逐字相同
        nr = 0
        for m in re.finditer(r'data-ref="README!A(\d+)"[^>]*>([^<]*)<', page):
            want = str(RD.cell(int(m.group(1)), 1).value or "")
            if H.unescape(m.group(2)).strip() != want.strip(): fail.append(f"HTML README!A{m.group(1)} 與 xlsx 不同：{H.unescape(m.group(2))[:40]}")
            nr += 1
        n_ok += nr
    else:
        fail.append(f"找不到 HTML：{HT}")
    # repo README 版本表
    md = open(os.path.join(folder, "README.md"), encoding="utf-8").read()
    blocks = VT.MARK.findall(md)
    if not blocks: fail.append("repo README.md 沒有 version-table 標記（版本表須由建置輸出產生）")
    cur = os.path.relpath(X, folder)
    if blocks and not any(p.split("｜")[-1].strip() == cur for p, _ in blocks): fail.append(f"repo README.md 沒有以目前成品 {cur} 為最後一欄的版本表")
    for p, body in blocks:
        want = VT.render(folder, [x.strip() for x in p.split("｜")], cfg["table_rows"])
        if body != want: fail.append(f"repo README.md 版本表與成品不一致：{p}")
        else: n_ok += 1
    print(f"展示層數字一致性（{os.path.basename(X)}）：比對 {n_ok} 處；" + ("全部通過" if not fail else f"失敗 {len(fail)} 項"))
    for f in fail[:40]: print("  -", f)
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))

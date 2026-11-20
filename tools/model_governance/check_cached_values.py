# 公式格快取值檢查（外部稽核優化 W2-6，2026-11-20；起因：外部稽核指出兩本地圖活頁簿的公式格都沒有快取值——
# openpyxl 直接寫出的檔沒有計算結果，不經重算就讀值的工具（data_only=True、預覽器、其他 AI）會讀到空白）。
# 只讀檔案，不改任何東西；不需要 LibreOffice。
# 用法：
#   python3 tools/model_governance/check_cached_values.py <成品.xlsx> [...]
#   python3 tools/model_governance/check_cached_values.py --all
#       掃 repo 內全部 dist／model 現行成品（清單同 check_zero_rows.py：略過 _待刪除、archive、legacy；*/model/ 只掃 CURRENT）。
# 判定：
#   直接讀 xlsx 內的工作表 XML。公式格（有 <f>）沒有 <v>，或 <v> 是空的且儲存格型別不是字串＝沒有快取值，計為失敗
#   （openpyxl 寫出的公式格是 <f>…</f><v></v>、沒有 t 屬性）。
#   <v> 是空的且 t="str"＝公式重算後回傳空字串（例：=IF(…,"",…)），不算失敗，另列計數。
#   共用公式的從屬格（<f t="shared" si="…"/> 沒有公式文字）同樣要有 <v>。
# 結束碼：0＝全部公式格都有快取值；1＝有公式格沒有快取值；2＝用法或讀檔錯誤。
import argparse, os, re, sys, zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from check_zero_rows import current_files  # noqa: E402

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def sheets(z):
    """回傳 [(工作表名稱, zip 內路徑)]。"""
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")).iter(PKG + "Relationship")}
    out = []
    for s in ET.fromstring(z.read("xl/workbook.xml")).iter(NS + "sheet"):
        t = rels[s.get(RNS + "id")]
        out.append((s.get("name"), t.lstrip("/") if t.startswith("/") else "xl/" + t))
    return out


def scan(path):
    """回傳 (公式格數, 無快取值的格 [(工作表, 儲存格)], 刻意空字串的格數)。"""
    n = empty = 0
    missing = []
    with zipfile.ZipFile(path) as z:
        for name, part in sheets(z):
            for _, c in ET.iterparse(z.open(part)):
                if c.tag != NS + "c":
                    continue
                if c.find(NS + "f") is not None:
                    n += 1
                    v = c.find(NS + "v")
                    if v is not None and (v.text or ""):
                        pass
                    elif v is not None and c.get("t") == "str":
                        empty += 1
                    else:
                        missing.append((name, c.get("r")))
                c.clear()
    return n, missing, empty


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="*")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--limit", type=int, default=5, help="每檔列出前幾個無快取值的格")
    a = ap.parse_args()
    root = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
    files = [os.path.join(root, f) for f in current_files(root)] if a.all else a.files
    if not files:
        ap.error("沒有要掃的檔案")
    bad = tot = tot_empty = 0
    for f in files:
        try:
            n, missing, empty = scan(f)
        except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError) as e:
            print(f"讀檔失敗：{f}（{e}）", file=sys.stderr)
            sys.exit(2)
        tot += n
        tot_empty += empty
        rel = os.path.relpath(f, root).replace(os.sep, "/") if a.all else f
        if missing:
            bad += 1
            head = "、".join(f"{s}!{c}" for s, c in missing[:a.limit])
            print(f"[無快取值] {rel}：公式格 {n}，無快取值 {len(missing)}（{head}{'…' if len(missing) > a.limit else ''}）")
        elif not a.all:
            print(f"[通過] {rel}：公式格 {n}，全部有快取值（其中刻意回傳空字串 {empty}）")
    print(f"掃描 {len(files)} 檔；公式格 {tot}；無快取值的檔 {bad}；刻意回傳空字串的格 {tot_empty}")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()

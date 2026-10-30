# In 頁版面套用（個股模型原則第 8 節；AI半導體 Project 2026-10-30）。
# 建置腳本在 wb.save() 之前呼叫 apply(wb, "vX.Y.Z", "YYYY-MM-DD")，只動呈現、不改任何數值：
#   1. In!D3／F3 欄名改「敏感度低」「敏感度高」，In!A1 句末補一句用途說明。
#   2. In 頁最後新增「情境驅動對照」一節：上方情境專屬列（判定與 check_in_layout.is_scn 相同）依同族兄弟分組，
#      每組一列，D／E／F＝情境 A／B／C，每格 =E<上方列>；情境 B 無專屬列時填 0；權重列（鍵 scn_w）放最後。
#      情境 A 無專屬列、但有去掉情境字母的同名列（例如 mplB_26／mplC_26 對 mpl_26）時，以該列為情境 A。
#   3. README 頁最後補一行版本說明。
# 只有單一情境的列（2026-10-30 放寬甲）：標籤「情境 X：…」者自成一列，其他情境欄填 0 並在 H 欄寫「情境 Y 無專屬輸入列」；
# 標籤含「（情境 X）」的上游參考值，在該列 H 欄加註「非本模型情境驅動」，不列入對照。
# 本節不建立任何定義名稱，也不被任何公式引用（純展示）。
import os, sys
from openpyxl.styles import Font

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import re  # noqa: E402
from check_in_layout import is_scn, KEY_OK, EXEMPT  # noqa: E402

BOLD = Font(bold=True)
TITLE = "情境驅動對照（三情境並排，進主模型的加權目標價；本節只引用上方各列，修改請改上方藍字。上方「敏感度低／高」欄只用於 Sens 頁，不進目標價）"
A1_NOTE = "【{v}】基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁（一次動一個參數）、不進目標價；情境 A／B／C 的驅動見本頁最後一節「情境驅動對照」。"
README_NOTE = ("{v}（{d}）：依個股模型原則第 8 節調整 In 頁呈現（數值不變）。In 頁 D／F 欄名改「敏感度低」「敏感度高」（只用於 Sens 頁、不進目標價）；"
               "In 頁最後新增「情境驅動對照」一節，把上方各情境專屬輸入三情境並排（只引用上方各列，純展示）；HTML 驅動因子一節分「情境驅動」與「單一參數」兩張表。")


class LayoutError(Exception):
    pass


def _letter_pos(key, keys):
    if not KEY_OK.match(key):
        return None
    for i, ch in enumerate(key):
        if ch not in "ABC":
            continue
        prev_ok = i == 0 or key[i - 1] == "_" or key[i - 1].islower() or key[i - 1].isdigit()
        next_ok = i == len(key) - 1 or key[i + 1] == "_" or key[i + 1].isdigit()
        if prev_ok and next_ok and any(key[:i] + o + key[i + 1:] in keys for o in "ABC" if o != ch):
            return i
    return None


def _strip(label):
    import re
    s = re.sub(r"^情境 ?[ABC][：:]\s*", "", label)
    s = re.sub(r"^情境 ?[ABC]（([^）]*)）[：:]\s*", r"（\1）", s)
    s = re.sub(r"（情境 ?[ABC]）", "", s)
    s = re.sub(r"情境權重 ?[ABC]", "情境權重", s)
    s = re.sub(r"[（(]\s*[ABC]\s*[）)]", "", s)
    s = re.sub(r"；\s*[ABC]）", "）", s)
    s = re.sub(r"(?<![A-Za-z])[ABC](?=（|$)", "", s).strip()  # 「目標價權重 A」之類的結尾字母
    return s or label


def groups(ws, base_map=None):
    """回傳 [(key, label, unit, kA, kB, kC, fmt)]；kB 可為 None（情境 B 填 0）。"""
    last = ws.max_row
    keys = {}
    for r in range(4, last + 1):
        v = ws.cell(r, 1).value
        if v and str(v) not in keys:
            keys[str(v)] = r
    fam, order, single = {}, [], []
    for k, r in keys.items():
        lab = str(ws.cell(r, 2).value or "")
        if not is_scn(k, lab, keys):
            continue
        i = _letter_pos(k, keys)
        if i is None:
            single.append(k)
            continue
        f = (k[:i], k[i + 1:])
        if f not in fam:
            fam[f] = {}
            order.append(f)
        fam[f][k[i]] = k
    out = []
    singles = []
    for k in single:
        r = keys[k]
        lab = str(ws.cell(r, 2).value or "")
        if re.search(r"（情境 ?[ABC]）", lab) and not re.match(r"^情境", lab):
            # 上游模型某情境的參考值（例如「…（情境 A）」）：依放寬甲標「非本模型情境驅動」，不列入對照
            h = ws.cell(r, 8)
            h.value = (str(h.value) + "；" if h.value else "") + EXEMPT + "（上游模型該情境的參考值）"
            continue
        m = re.match(r"^情境 ?([ABC])", lab)
        if not m:
            raise LayoutError(f"無法判定情境字母的單一情境列：{k}")
        singles.append((k, m.group(1)))
    for f in order:
        d = fam[f]
        if "A" not in d:
            base = (f[0].rstrip("_") + "_" + f[1].lstrip("_")) if f[0].endswith("_") or f[1].startswith("_") else f[0] + f[1]
            cand = [c for c in (f[0] + f[1], base, f[0].rstrip("_")) if c in keys]
            if not cand:
                raise LayoutError(f"{f[0]}?{f[1]} 缺情境 A 列，且找不到去掉情境字母的同名列")
            d["A"] = cand[0]
        if "C" not in d:
            raise LayoutError(f"{f[0]}?{f[1]} 缺情境 C 列")
        ra = keys[d["A"]]  # noqa
        lab = _strip(str(ws.cell(keys[d.get("C")], 2).value or ""))
        key = "scn_" + (f[0] + f[1]).strip("_")
        out.append([key, lab, ws.cell(ra, 3).value, d["A"], d.get("B"), d["C"], ws.cell(ra, 5).number_format, keys])
    for k, x in singles:
        r = keys[k]
        refs = {x: k}
        # 其他情境若用同名基準列（去掉情境字母，或字母前的前綴，例 g_tlC_28 → g_tl），就引用該列；找不到才填 0
        bases = []
        for i, ch in enumerate(k):
            if ch == x and (i + 1 == len(k) or k[i + 1] in "_0123456789"):
                bases += [k[:i] + k[i + 1:], k[:i].rstrip("_")]
        bases = ([base_map[k]] if base_map and k in base_map else []) + bases
        base = next((b for b in bases if b in keys and not is_scn(b, str(ws.cell(keys[b], 2).value or ""), keys)), None)
        if base:
            for y in "ABC":
                if y != x and not (y == "A" and x == "C" and False):
                    refs.setdefault(y, base)
        out.append(["scn_" + k, _strip(str(ws.cell(r, 2).value or "")), ws.cell(r, 3).value,
                    refs.get("A"), refs.get("B"), refs.get("C"), ws.cell(r, 5).number_format, keys])
    wt = [g for g in out if "權重" in "".join(str(ws.cell(keys[k], 2).value or "") for k in g[3:6] if k)]
    if len(wt) != 1:
        raise LayoutError(f"情境權重列應恰好一組：{[g[0] for g in wt]}")
    out.remove(wt[0])
    wt[0][0], wt[0][1] = "scn_w", "情境權重"
    return out + wt


def apply(wb, version, date, base_map=None):
    """base_map：單一情境列 → 其他情境使用的基準列鍵（名稱不同名時由建置腳本指定，例 advantest shC_28 → sh_soc）"""
    ws = wb["In"]
    if str(ws.cell(3, 5).value) != "基準":
        raise LayoutError("In!E3 不是「基準」")
    g = groups(ws, base_map)
    ws.cell(3, 4).value, ws.cell(3, 6).value = "敏感度低", "敏感度高"
    a1 = str(ws["A1"].value or "")
    ws["A1"] = (a1.rstrip("。") + "。" if a1 else "") + A1_NOTE.format(v=version)
    r = ws.max_row + 2
    ws.cell(r, 1, TITLE).font = BOLD
    r += 1
    for j, h in enumerate(["鍵", "項目", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "標記", "說明"], 1):
        ws.cell(r, j, h).font = BOLD
    r += 1
    for key, lab, unit, ka, kb, kc, fmt, keys in g:
        ws.cell(r, 1, key); ws.cell(r, 2, lab); ws.cell(r, 3, unit)
        refs = [ka, kb, kc]
        for j, k in zip((4, 5, 6), refs):
            c = ws.cell(r, j, f"=E{keys[k]}" if k else 0)
            c.number_format = fmt
        ws.cell(r, 7, "Derived")
        note = "；".join([f"情境 {x} 無專屬輸入列" for x, kk in zip("AC", (ka, kc)) if kk is None]
                         + (["情境 B 無專屬輸入列，填 0"] if kb is None else []) or ["修改請改上方藍字"])
        if ka and not is_scn(ka, str(ws.cell(keys[ka], 2).value or ""), keys):
            note += f"；情境 A 用上方基準列 {ka}"
        ws.cell(r, 8, "引用 In 頁 " + "／".join(k for k in refs if k) + "；" + note)
        r += 1
    rd = wb["README"]
    rr = rd.max_row + 2
    rd.cell(rr, 1, README_NOTE.format(v=version, d=date))
    return len(g)

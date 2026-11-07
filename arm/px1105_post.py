# 2026-11-05 收盤換價（工作單 px1105，Andy 2026-11-06 07:28）共用後處理：建置腳本在存檔前呼叫 apply(wb, …)。
# 各公司資料夾放同一份（公司自有，不屬於 tools/）。只改文字，不改任何數值公式。
#   1. 情境名稱改回「情境 A（樂觀）／B（基準）／C（保守）」（Andy 2026-11-06 03:14，取代第三輪修正第 23 點）。
#      版本沿革列（以 v0. 開頭）保留當時寫法；keep 內的字串（「平平」作一般形容詞等）不改。
#   2. README 版本沿革中以公式帶出「本版」數字的舊版列，凍結為上一版成品的顯示文字（hist_frozen.json），避免顯示新版數字。
#   3. README A1 版號、「本版要點」第一列加本版說明（固定文字）。
import json, os, re

SCN = (("不如預期", "保守"), ("順利", "樂觀"), ("平平", "基準"))
NOTE = "現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守"
_HIST = re.compile(r'^(="?)?v\d+\.\d')


def apply(wb, here, ver, old_ver, anchor, keep=()):
    """anchor：README 版本沿革中目前最新一列的開頭（本版說明插在它前面）。"""
    log = []
    RD = wb["README"]
    # 2：凍結舊版公式列
    fp = os.path.join(here, "hist_frozen.json")
    frozen = json.load(open(fp, encoding="utf-8")) if os.path.exists(fp) else {}
    m = 0
    for r in range(1, RD.max_row + 1):
        v = RD.cell(r, 1).value
        if isinstance(v, str) and v.startswith('="v0.'):
            for k, t in frozen.items():
                if v.startswith('="' + k):
                    RD.cell(r, 1).value = t; m += 1
                    break
    assert m == len(frozen), ("hist_frozen 未全部套用", m, len(frozen))
    log.append(f"凍結版本沿革 {m} 列")
    # 1：情境名稱
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str) or not any(a in v for a, _ in SCN) or _HIST.match(v):
                    continue
                w = v
                for i, k in enumerate(keep):
                    w = w.replace(k, f"\x00{i}\x00")
                for a, b in SCN:
                    w = w.replace(a, b)
                for i, k in enumerate(keep):
                    w = w.replace(f"\x00{i}\x00", k)
                if w != v:
                    c.value = w; n += 1
    log.append(f"情境名稱 {n} 格")
    # 3：版號與版本沿革
    a1 = RD["A1"].value
    assert isinstance(a1, str) and (old_ver in a1 or ver in a1), ("README A1 無版號", old_ver)
    RD["A1"].value = (a1.replace(old_ver, ver, 1) if ver not in a1 else a1) + f"；{ver}：{NOTE}"
    h = [r for r in range(1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith(anchor)]
    assert len(h) == 1, ("版本沿革錨點", anchor, h)
    RD.insert_rows(h[0])
    h = [h[0] - 1]
    RD.cell(h[0] + 1, 1, f"{ver}：{NOTE}")
    try:
        from openpyxl.styles import Alignment
        RD.cell(h[0] + 1, 1).alignment = Alignment(wrap_text=True, vertical="top")
    except Exception:
        pass
    log.append(f"版本沿革插入 README!A{h[0] + 1}")
    return "；".join(log)

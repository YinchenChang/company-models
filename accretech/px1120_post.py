# 東京精密 0.3.8（2026-11-21）後處理：工作單 px1120——現價與同業報價改 2026-11-20 收盤（只換價格；公式、共識 EPS、rf、β、ERP、預測匯率、營運假設、情境權重不變）。
# 價格只取自 markets/invest_map/data/closes_20261109_20261120.csv（Yahoo Finance 日線收盤；交易所當地日期）。
# 同業前瞻 P/E 新值＝現行倍數 × 11-20 收盤 ÷ 報價日（2026-11-13）收盤（共識 EPS 不變）；前例：px1113。疊在 px1113_post 之後執行。
import re

OLDVER, NEWVER, DATE = "0.3.7", "0.3.8", "2026-11-21"
TK = "7729.T"
PX_OLD, PX_NEW = 13150, 11185
PXF = "#,##0"
PEERS = {'6146.T': (47470, 41350), '8035.T': (9461, 8309), '6857.T': (30610, 24255), '6920.T': (36130, 30250), '6525.T': (5984, 4897)}   # 代號 -> (2026-11-13 收盤, 2026-11-20 收盤)
OLD = {'med': 25.22, 'tp': 10533.3193794466, 'up': -0.198987119433719, 'rec': '中立', 'nd': 0}      # v0.3.7 成品值
OLD_TXT = "v0.3.7（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）。現價 12,635 → 13,150；同業前瞻 P/E 中位數 23.33 → 25.22；加權目標價 10,078 → 10,533；上檔空間 -20.2%（中立）→ -19.9%（中立）。共識 EPS、rf、β、ERP、預測匯率、營運假設與情境權重不變；未換價同業：無（本表同業皆在收盤資料內）；並列參考列東京精密自身前瞻 P/E（2026-09-29，不進目標價）不換。其餘內容同 v0.3.6"
UNCH = "未換價同業：無（本表同業皆在收盤資料內）；並列參考列東京精密自身前瞻 P/E（2026-09-29，不進目標價）不換"
SRC = "工作單 px1120；Yahoo Finance 日線收盤 2026-11-20（markets/invest_map/data/closes_20261109_20261120.csv）"


def _lit(t):
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def apply(wb):
    In, V, RD = wb["In"], wb["Val"], wb["README"]
    skip = set()
    r = next(r for r in range(1, In.max_row + 1) if In.cell(r, 1).value == "px")
    assert abs(In.cell(r, 5).value - PX_OLD) < 1e-9, (In.cell(r, 5).value, PX_OLD)
    old_i = str(In.cell(r, 9).value or "")
    In.cell(r, 5).value = PX_NEW
    In.cell(r, 2).value = str(In.cell(r, 2).value).replace("2026-11-13 收盤", "2026-11-20 收盤")
    In.cell(r, 8).value = f"Yahoo Finance 日線收盤 2026-11-20：{PX_NEW:,}（{TK}）"
    In.cell(r, 9).value = f"工作單 px1120（{DATE}）：現價改 2026-11-20 收盤（原 {PX_OLD:,}，2026-11-13 收盤，v{OLDVER}）；" + old_i
    skip |= {("In", r, 8), ("In", r, 9)}
    h = next(r for r in range(1, V.max_row + 1) if str(V.cell(r, 1).value or "").startswith("0.2 Comparable"))
    done = []
    for rr in range(h + 2, h + 16):
        if str(V.cell(rr, 1).value or "").strip() == "同業前瞻 P/E 中位數":
            break
        t, pe, d = V.cell(rr, 2).value, V.cell(rr, 3).value, V.cell(rr, 4).value
        if t in PEERS and d == "2026-11-13":
            p0, p1 = PEERS[t]
            assert isinstance(pe, (int, float)), (t, pe)
            V.cell(rr, 3).value = round(pe * p1 / p0, 2)
            V.cell(rr, 4).value = "2026-11-20"
            V.cell(rr, 5).value = (str(V.cell(rr, 5).value or "") + f"；v{NEWVER}：換算到 2026-11-20＝{pe}（2026-11-13）× 11-20 收盤 {p1:,} ÷ 11-13 收盤 {p0:,}（{SRC}；共識 EPS 不變）")
            for cc in (3, 4, 5): skip.add(("Val", rr, cc))
            done.append(t)
    assert sorted(done) == sorted(PEERS), (done, sorted(PEERS))
    # README 版本列：先處理（凍結 v0.3.7 列），再做全表標籤替換
    lab = lambda ws, p: next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(p))
    try:
        r13 = lab(RD, "v1.3 展示層")
    except StopIteration:
        r13 = None   # 敏感度／協調快照建置（無 v1.3 展示層）：只換價，不寫版本列
    rv = (r13 or 0) + 9
    frozen = set()
    if r13:
        assert str(RD.cell(rv, 1).value).startswith('="v0.3.7（2026-11-16）'), RD.cell(rv, 1).value
        frozen = {rv} | {rr for rr in range(rv + 1, RD.max_row + 1) if str(RD.cell(rr, 1).value or "").startswith('="v0.3.7（2026-11-16）')}
    for rr in frozen:
        RD.cell(rr, 1).value = OLD_TXT     # 上一版本版列凍結為文字
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str) or (ws.title, c.row, c.column) in skip:
                    continue
                if ws.title == "README" and (c.coordinate == "A1" or re.match(r"v0\.\d", v.strip())):
                    continue
                w = v.replace("2026-11-13 收盤 " + f"{PX_OLD:,}", "2026-11-20 收盤 " + f"{PX_NEW:,}").replace("2026-11-13 收盤", "2026-11-20 收盤")
                if w != v:
                    c.value = w
    a1 = str(RD["A1"].value); head = a1.split("；", 1)[0]
    if f"v{OLDVER}" in head:
        a1 = head.replace(f"v{OLDVER}", f"v{NEWVER}", 1) + a1[len(head):]
    RD["A1"].value = a1 + f"；v{NEWVER}（{DATE}）：現價與同業報價改 2026-11-20 收盤（工作單 px1120）；共識 EPS、rf、模型假設不變"
    if not r13:
        return f"px1120：現價 {PX_OLD} → {PX_NEW}；同業 {len(done)} 家換價（快照建置）"
    r03 = lab(V, "0.2 Comparable")
    vr = lambda p: "Val!B%d" % next(r for r in range(r03, V.max_row + 1) if str(V.cell(r, 1).value or "").strip() == p)
    pm, tp, up, rec = vr("同業前瞻 P/E 中位數"), vr("目標價（DCF 與 Comparable 加權）"), vr("上檔空間：加權目標價"), vr("投資建議")
    T = lambda ref, f: f'TEXT({ref},"{f}")'
    o = OLD
    f = ("=" + "&".join([
        _lit(f"v{NEWVER}（{DATE}）：現價與同業報價改 2026-11-20 收盤（工作單 px1120）。現價 {PX_OLD:,} → "), T("px", PXF),
        _lit(f"；同業前瞻 P/E 中位數 {o['med']:.2f} → "), T(pm, "0.00"),
        _lit(f"；加權目標價 {o['tp']:,.{o['nd']}f} → "), T(tp, PXF),
        _lit(f"；上檔空間 {o['up']*100:+.1f}%（{o['rec']}）→ "), T(up, "+0.0%;-0.0%"), _lit("（"), rec, _lit("）。"),
        _lit(f"共識 EPS、rf、β、ERP、預測匯率、營運假設與情境權重不變；{UNCH}。其餘內容同 v{OLDVER}")]))
    RD.cell(rv, 1).value = f
    RD.cell(RD.max_row + 2, 1, f)
    return f"px1120：現價 {PX_OLD} → {PX_NEW}；同業 {len(done)} 家換價"

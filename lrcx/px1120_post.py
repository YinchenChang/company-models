# 科林研發 0.3.7（2026-11-21）後處理：工作單 px1120——現價與同業報價改 2026-11-20 收盤（只換價格；公式、共識 EPS、rf、β、ERP、預測匯率、營運假設、情境權重不變）。
# 價格只取自 markets/invest_map/data/closes_20261109_20261120.csv（Yahoo Finance 日線收盤；交易所當地日期）。
# 同業前瞻 P/E 新值＝現行倍數 × 11-20 收盤 ÷ 報價日（2026-11-13）收盤（共識 EPS 不變）；前例：px1113（2026-11-16）。接在 px1113_post 之後執行。
import re

OLDVER, NEWVER, DATE = "0.3.6", "0.3.7", "2026-11-21"
TK = "LRCX"
PX_OLD, PX_NEW = 225.92, 183.27
PXF = "#,##0.00"
PEERS = {'AMAT': (370.5, 302.63), 'KLAC': (152.32, 126.87), 'ASML': (1495.16, 1266.22), '8035.T': (9461, 8309)}   # 代號 -> (2026-11-13 收盤, 2026-11-20 收盤)
OLD = {'med': 24.38, 'tp': 157.535944751501, 'up': -0.302691462679264, 'rec': '放空', 'nd': 2}      # v0.3.6 成品值
UNCH = "未換價同業：無（本表同業皆在收盤資料內）"
SRC = "工作單 px1120；Yahoo Finance 日線收盤 2026-11-20（markets/invest_map/data/closes_20261109_20261120.csv）"
OLD_DAY, NEW_DAY = "2026-11-13", "2026-11-20"
FROZEN_V036 = ("v0.3.6（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）。現價 185.2 → 225.92；同業前瞻 P/E 中位數 22.34 → 24.38；加權目標價 148.15 → 157.54；"
               "上檔空間 -20.0%（中立）→ -30.3%（放空）。共識 EPS、rf、β、ERP、預測匯率、營運假設與情境權重不變；未換價同業：無（本表同業皆在收盤資料內）。其餘內容同 v0.3.5")

def _lit(t):   # Excel 公式字串常數上限 255 字元：切段以 & 相接
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def _fmt(v, nd):
    return f"{v:,.{nd}f}"


def apply(wb):
    In, V, RD = wb["In"], wb["Val"], wb["README"]
    skip = set()
    # (1) 現價 px
    r = next(r for r in range(1, In.max_row + 1) if In.cell(r, 1).value == "px")
    assert abs(In.cell(r, 5).value - PX_OLD) < 1e-9, (In.cell(r, 5).value, PX_OLD)
    old_i = str(In.cell(r, 9).value or "")
    In.cell(r, 5).value = PX_NEW
    In.cell(r, 2).value = str(In.cell(r, 2).value).replace(f"{OLD_DAY} 收盤", f"{NEW_DAY} 收盤")
    In.cell(r, 8).value = f"Yahoo Finance 日線收盤 {NEW_DAY}：{PX_NEW:,}（{TK}）"
    In.cell(r, 9).value = f"工作單 px1120（{DATE}）：現價改 {NEW_DAY} 收盤（原 {PX_OLD:,}，{OLD_DAY} 收盤，v{OLDVER}；{SRC}）；" + old_i
    skip |= {("In", r, 8), ("In", r, 9)}
    # (2) Comparable 同業前瞻 P/E
    h = next(r for r in range(1, V.max_row + 1) if str(V.cell(r, 1).value or "").startswith("0.2 Comparable"))
    done = []
    for rr in range(h + 2, h + 16):
        if str(V.cell(rr, 1).value or "").strip() == "同業前瞻 P/E 中位數":
            break
        t, pe, d = V.cell(rr, 2).value, V.cell(rr, 3).value, V.cell(rr, 4).value
        if t in PEERS and d == OLD_DAY:
            p0, p1 = PEERS[t]
            assert isinstance(pe, (int, float)), (t, pe)
            V.cell(rr, 3).value = round(pe * p1 / p0, 2)
            V.cell(rr, 4).value = NEW_DAY
            V.cell(rr, 5).value = (str(V.cell(rr, 5).value or "") + f"；v{NEWVER}：換算到 {NEW_DAY}＝{pe}（{OLD_DAY}）× 11-20 收盤 {p1:,} ÷ 11-13 收盤 {p0:,}（{SRC}；共識 EPS 不變）")
            for cc in (3, 4, 5): skip.add(("Val", rr, cc))
            done.append(t)
    assert sorted(done) == sorted(PEERS), (done, sorted(PEERS))
    # (3) 其他說明文字的現價日期（版本沿革列、README A1 不改）
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str) or (ws.title, c.row, c.column) in skip:
                    continue
                if ws.title == "README" and (c.coordinate == "A1" or re.match(r"v0\.\d", v.strip()) or v.startswith('="v0.3.')):
                    continue
                w = v.replace(f"{OLD_DAY} 收盤", f"{NEW_DAY} 收盤")
                if w != v:
                    c.value = w
    # (4) README：A1、版本列、本版說明
    lab = lambda ws, p: next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(p))
    RD["A1"].value = str(RD["A1"].value) + f"；v{NEWVER}（{DATE}）：現價與同業報價改 {NEW_DAY} 收盤（工作單 px1120）；共識 EPS、rf、模型假設不變"
    try:
        r13 = lab(RD, "v1.3 展示層")
    except StopIteration:
        return "px1120：無 v1.3 展示層（敏感度／協調快照建置）"
    rv = r13 + 9
    old_f = str(RD.cell(rv, 1).value)
    assert old_f.startswith('="v0.3.6'), old_f
    # 上一版（v0.3.6）版本列：README 頁尾以公式帶數字者凍結為上一版成品值的顯示文字
    frozen = [r for r in range(rv + 1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith('="v0.3.6')]
    assert len(frozen) == 1, frozen
    RD.cell(frozen[0], 1).value = FROZEN_V036
    r03 = lab(V, "0.2 Comparable")
    vr = lambda p: "Val!B%d" % next(r for r in range(r03, V.max_row + 1) if str(V.cell(r, 1).value or "").strip() == p)
    pm, tp, up, rec = vr("同業前瞻 P/E 中位數"), vr("目標價（DCF 與 Comparable 加權）"), vr("上檔空間：加權目標價"), vr("投資建議")
    T = lambda ref, f: f'TEXT({ref},"{f}")'
    o = OLD
    f = ("=" + "&".join([
        _lit(f"v{NEWVER}（{DATE}）：現價與同業報價改 {NEW_DAY} 收盤（工作單 px1120）。現價 {PX_OLD:,} → "), T("px", PXF),
        _lit(f"；同業前瞻 P/E 中位數 {o['med']:.2f} → "), T(pm, "0.00"),
        _lit(f"；加權目標價 {o['tp']:,.{o['nd']}f} → "), T(tp, PXF),
        _lit(f"；上檔空間 {o['up']*100:+.1f}%（{o['rec']}）→ "), T(up, "+0.0%;-0.0%"), _lit("（"), rec, _lit("）。"),
        _lit(f"共識 EPS、rf、β、ERP、預測匯率、營運假設與情境權重不變；{UNCH}。其餘內容同 v{OLDVER}")]))
    RD.cell(rv, 1).value = f
    RD.cell(RD.max_row + 1, 1, f)
    return f"px1120：現價 {PX_OLD} → {PX_NEW}；同業 {len(done)} 家換價"

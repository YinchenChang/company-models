# 泰瑞達 0.4.6（2026-11-16）→ 0.4.7（2026-11-21）後處理：工作單 px1120——現價與同業報價改 2026-11-20 收盤（只換價格；公式、共識 EPS、rf、β、ERP、預測匯率、營運假設、情境權重不變）。
# 價格只取自 markets/invest_map/data/closes_20261109_20261120.csv（Yahoo Finance 日線收盤；交易所當地日期；TPRO.MI 11-20 日線尚無該列，取報價資料最近成交價）。
# 同業前瞻 P/E 新值＝現行倍數 × 11-20 收盤 ÷ 報價日（2026-11-13）收盤（共識 EPS 不變）；前例：px1113（ter/px1113_post.py）。
# 本檔在 px1113_post.apply 之後執行（建置腳本末端呼叫）。
import re

OLDVER, NEWVER, DATE = "0.4.6", "0.4.7", "2026-11-21"
TK = "TER"
PX_OLD, PX_NEW = 249.55, 204.78
PXF = "#,##0.00"
PEERS = {'6857.T': (30610, 24255), 'KLAC': (152.32, 126.87), 'COHU': (43.59, 37.31), 'FORM': (87.53, 73.02), 'AMAT': (370.5, 302.63)}   # 代號 -> (2026-11-13 收盤, 2026-11-20 收盤)
OLD = {'med': 27.88, 'tp': 319.493878121244, 'up': 0.280280016514704, 'rec': '買進', 'nd': 2}      # v0.4.6 成品值
UNCH = "DCF 淨現金內 Technoprobe（TPRO.MI）市值的輸入改 11-20：股價 €21.98 → €18.41（報價資料最近成交價）、歐元兌美元 1.105 → 1.097 US$/€（EURUSD=X；預測用匯率不變）；未換價：無"
SRC = "工作單 px1120；Yahoo Finance 日線收盤 2026-11-20（markets/invest_map/data/closes_20261109_20261120.csv）"
# v0.4.6 版本列（px1113 當時顯示文字；原為公式，本版起凍結）
FROZEN_046 = ('v0.4.6（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）。現價 212.74 → 249.55；同業前瞻 P/E 中位數 23.76 → 27.88；加權目標價 300.24 → 319.49；上檔空間 +41.1%（買進）→ +28.0%（買進）。'
              '共識 EPS、rf、β、ERP、預測匯率、營運假設與情境權重不變；DCF 淨現金內 Technoprobe（TPRO.MI）市值的輸入改 11-13 收盤：股價 €20.52 → €21.98、歐元兌美元 0.8904 €/US$ → 1.105 US$/€（EURUSD=X；預測用匯率不變）；未換價：無。其餘內容同 v0.4.5')


def _lit(t):   # Excel 公式字串常數上限 255 字元：切段以 & 相接
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def apply(wb):
    In, V, RD = wb["In"], wb["Val"], wb["README"]
    skip = set()
    # (1) 現價 px
    r = next(r for r in range(1, In.max_row + 1) if In.cell(r, 1).value == "px")
    assert abs(In.cell(r, 5).value - PX_OLD) < 1e-9, (In.cell(r, 5).value, PX_OLD)
    old_i = str(In.cell(r, 9).value or "")
    In.cell(r, 5).value = PX_NEW
    In.cell(r, 2).value = str(In.cell(r, 2).value).replace("2026-11-13 收盤", "2026-11-20 收盤")
    In.cell(r, 8).value = f"Yahoo Finance 日線收盤 2026-11-20：{PX_NEW:,}（{TK}）"
    In.cell(r, 9).value = f"工作單 px1120（2026-11-21）：現價改 2026-11-20 收盤（原 {PX_OLD:,}，2026-11-13 收盤，v{OLDVER}）；" + old_i
    skip |= {("In", r, 8), ("In", r, 9)}
    # (1b) Technoprobe（TPRO.MI）股價 €21.98 → €18.41、歐元兌美元 1.105 → 1.097 US$/€；僅換市值換算，預測用匯率不變
    hits = 0
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str):
                    continue
                w = v.replace("65.326087*21.98*1.105", "65.326087*18.41*1.097")
                w = w.replace("€21.98 × 1.105 US$/€（Yahoo Finance 日線收盤 2026-11-13：TPRO.MI、EURUSD=X；工作單 px1113 K 批：",
                              "€18.41 × 1.097 US$/€（Yahoo Finance 2026-11-20：TPRO.MI 為報價資料最近成交價〔日線尚無該列〕、EURUSD=X 日線收盤；工作單 px1120：v0.4.6 為 €21.98 × 1.105 US$/€，2026-11-13；px1113 K 批：")
                w = w.replace("65.33M 股 × €21.98 × 1.105，2026-11-13）", "65.33M 股 × €18.41 × 1.097，2026-11-20）")
                if w != v:
                    c.value = w; hits += 1
                    skip.add((ws.title, c.row, c.column))
    assert hits >= 4, hits
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and "21.98*1.105" in c.value:
                    print("警告：殘留舊 Technoprobe 輸入", ws.title, c.coordinate)
    # (2) Comparable 同業前瞻 P/E
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
    # (3) 其他說明文字的現價日期（README 全頁不動：版本沿革與本版說明另處理）
    for ws in wb:
        if ws.title == "README":
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str) or (ws.title, c.row, c.column) in skip:
                    continue
                w = v.replace("2026-11-13 收盤", "2026-11-20 收盤")
                if w != v:
                    c.value = w
    # (4) README：A1、版本列、本版說明
    lab = lambda ws, p: next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(p))
    a1 = str(RD["A1"].value); head = a1.split("；", 1)[0]
    if f"v{OLDVER}" in head:
        a1 = head.replace(f"v{OLDVER}", f"v{NEWVER}", 1) + a1[len(head):]
    RD["A1"].value = a1 + f"；v{NEWVER}（{DATE}）：現價與同業報價改 2026-11-20 收盤（工作單 px1120）；共識 EPS、rf、模型假設不變"
    try:
        r13 = lab(RD, "v1.3 展示層")
    except StopIteration:
        return "px1120：無 v1.3 展示層（敏感度／協調快照建置）"
    rv = r13 + 9
    assert str(RD.cell(rv, 1).value).startswith(f'="v{OLDVER}'), RD.cell(rv, 1).value
    # px1113 在頁末另放一份同公式列：凍結為 v0.4.6 顯示文字
    frozen_rows = [r for r in range(rv + 1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith(f'="v{OLDVER}')]
    assert len(frozen_rows) == 1, frozen_rows
    RD.cell(frozen_rows[0], 1).value = FROZEN_046
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
    r = RD.max_row + 2
    RD.cell(r, 1, f)
    return f"px1120：現價 {PX_OLD} → {PX_NEW}；同業 {len(done)} 家換價"

"""px1120：WhiteFiber 現價與同業報價改 2026-11-20 收盤（一次性改寫 build_whitefiber.py 的價格輸入與標籤）。
價格取自 markets/invest_map/data/closes_20261109_20261120.csv。"""
import re
p = "whitefiber/build_whitefiber.py"
s = open(p, encoding="utf-8").read()
L = s.split("\n")
SRC = "工作單 px1120；Yahoo Finance 日線收盤 2026-11-20（`closes_20261109_20261120.csv`）"
def setline(i, f):
    L[i-1] = f(L[i-1])
def px(i, name, old, new, oldlabel):
    def f(l):
        l = l.replace("2026-11-13 收盤）", "2026-11-20 收盤）", 1)
        # 值
        l = l.replace(f", {old}, None, ", f", {new}, None, ", 1)
        return l
    setline(i, f)
# 現價
L[94-1] = ('inp("px", "現價（2026-11-20 收盤）", "US$", None, 6.35, None, "Verified", "' + SRC + '：6.35", '
           '"v0.4.3 為 2026-11-13 收盤 9.43；v0.4.2 為 2026-11-10 收盤 " + str(fv("mkt.price.20261110")) + "；v0.2.1 現價 7.80（2026-10-06 前後）")')
def peer(i, key, label, new, olds, extra=""):
    nm = label
    L[i-1] = (f'inp("{key}", "{nm} 股價（2026-11-20 收盤）", "US$", None, {new}, None, "Verified", "{SRC}：{new}", "v0.4.3 為 2026-11-13 收盤 {olds}；{extra}更早：見 v0.4.3 沿革")')
peer(225, "pc_px", "CoreWeave", 39.51, "53.13")
peer(232, "pn_px", "Nebius", 81.31, "115.88")
peer(238, "ir_px", "IREN", 11.82, "18.17", "EV＝股價 × 股數＋淨負債，依公式重算；")
peer(244, "po_px", "Oracle", 101.19, "125.35")
L[223-1] = 'PEER_D = "2026-11-20"'
setline(256, lambda l: l.replace("收盤價比 11-13 ÷ 10-07", "收盤價比 11-20 ÷ 10-07").replace("=86198*179.99/180.47", "=86198*162.46/180.47")
        .replace("11-13＝179.99（`closes_20261020_20261113.csv`，9-28～10-19 段已補）", "11-20＝162.46（`closes_20261109_20261120.csv`；10-07 收盤取 `closes_20261020_20261113.csv`）")
        .replace("工作單 px1113：", "工作單 px1120：").replace("；v0.4.2 為 EV 86,198（倍數 20.42）", "；v0.4.3（11-13）EV 85,969（倍數 20.36）；v0.4.2 為 EV 86,198（倍數 20.42）"))
setline(258, lambda l: l.replace("收盤價比 11-13 ÷ 10-02", "收盤價比 11-20 ÷ 10-02").replace("=121482*1053.79/1025.72", "=121482*968.22/1025.72")
        .replace("11-13＝1,053.79（`closes_20261020_20261113.csv`）", "11-20＝968.22（`closes_20261109_20261120.csv`；10-02 收盤取 `closes_20261020_20261113.csv`）")
        .replace("工作單 px1113：", "工作單 px1120：").replace("；v0.4.2 為 EV 121,482（倍數 21.06）", "；v0.4.3（11-13）EV 124,806（倍數 21.63）；v0.4.2 為 EV 121,482（倍數 21.06）"))
setline(938, lambda l: l.replace("2026-11-13＝現價日；IREN 11-11 收盤 20.84 → 11-13 收盤 18.17", "2026-11-20＝現價日；四家皆為 11-20 收盤（IREN 11-13 收盤 18.17 → 11-20 收盤 11.82）"))
setline(959, lambda l: l.replace("估值日 2026-11-13（＝本模型現價日；IREN 亦為 11-13 收盤 18.17，原 11-11 收盤 20.84）", "估值日 2026-11-20（＝本模型現價日；四家皆為 11-20 收盤，IREN 原 11-13 收盤 18.17）"))
for i in (944, 945, 946):
    setline(i, lambda l: l.replace("共識（擷取 2026-11-13）", "共識（擷取 2026-11-13；本版只換股價）"))
setline(966, lambda l: l.replace('"2026-11-13")', '"2026-11-20")'))
for i in (975, 976):
    setline(i, lambda l: l.replace("2026-11-13（EV 自", "2026-11-20（EV 自"))
setline(982, lambda l: l.replace("至 11-13 收盤的價格比", "至 11-20 收盤的價格比"))
setline(1078, lambda l: l.replace("同業估值日 11-13 與現價日同日（IREN 亦為 11-13 收盤；DLR、EQIX 維持 10 月初）", "同業估值日 11-20 與現價日同日（IREN 亦為 11-20 收盤；DLR、EQIX 的 EBITDA 共識與淨負債維持 10 月初）"))
L[13-1] = 'VER = "v0.4.4"'
L[14-1] = 'BUILT = "2026-11-21"'
# 沿革列：插在 v0.4.3 列後（先放佔位，數字待重建後回填）
idx = next(i for i, l in enumerate(L) if l.startswith('RD.cell(r, 1, "v0.4.3（2026-11-16）'))
L.insert(idx+1, "r += 1")
L.insert(idx+2, 'RD.cell(r, 1, "v0.4.4（2026-11-21）：現價與同業報價改 2026-11-20 收盤（工作單 px1120；Yahoo Finance 日線，markets/invest_map/data/closes_20261109_20261120.csv）。現價 9.43 → 6.35；CoreWeave 53.13 → 39.51、Nebius 115.88 → 81.31、IREN 18.17 → 11.82、Oracle 125.35 → 101.19；DLR EV 85,969 → 依 10-07 收盤 180.47 與 11-20 收盤 162.46 重算、EQIX EV 依 10-02 收盤 1,025.72 與 11-20 收盤 968.22 重算（倍數 × 價格比近似，Derived）。rf、共識、淨負債、股數、營運假設與情境權重不變；無未換價同業。結果數字見 README v0.4.4。")')
open(p, "w", encoding="utf-8").write("\n".join(L))

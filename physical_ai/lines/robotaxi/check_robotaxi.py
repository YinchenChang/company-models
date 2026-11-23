# -*- coding: utf-8 -*-
"""核對：python3 check_robotaxi.py [xlsx]（須先以 LibreOffice 重算）。全部通過才交付。v0.2 新增第 7–10 節。"""
import sys
import os
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
P = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "dist", "20261116_Robotaxi單位經濟_v0.2.xlsx")
wbv = openpyxl.load_workbook(P, data_only=True)
wbf = openpyxl.load_workbook(P)
fails = []


def ok(cond, msg):
    if not cond:
        fails.append(msg)


# 1. 無錯誤值、公式皆有值
for ws in wbv.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("#"):
                fails.append(f"錯誤值 {ws.title}!{c.coordinate} {c.value}")
for ws in wbf.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                if wbv[ws.title][c.coordinate].value is None:
                    fails.append(f"未重算 {ws.title}!{c.coordinate}")

# 2. 輸入頁：標記合法；Analogy／Assumed 三值齊全；依情境三值齊全；每列有出處欄
TAGS = {"[Verified]", "[Interested-party]", "[Analogy]", "[Assumed]", "[Derived]"}
wi = wbv["輸入"]
for r in range(2, wi.max_row + 1):
    iid, tag, use = wi.cell(r, 1).value, wi.cell(r, 7).value, wi.cell(r, 8).value
    if not tag:
        continue
    ok(tag in TAGS, f"輸入 {iid} 標記不合法 {tag}")
    vals = [wi.cell(r, j).value for j in (4, 5, 6)]
    if tag in ("[Analogy]", "[Assumed]") or use == "依情境":
        ok(all(v is not None for v in vals), f"輸入 {iid} 須有低／基準／高三值")
    ok(vals[1] is not None, f"輸入 {iid} 缺基準值")
    ok(wi.cell(r, 9).value not in (None, ""), f"輸入 {iid} 缺出處欄")

# 3. 情境頁邏輯
wn = wbv["情境"]
rows = {}
cur = None
for r in range(1, wn.max_row + 1):
    a = wn.cell(r, 1).value
    if isinstance(a, str) and a.startswith("情境："):
        cur = a
        rows[cur] = {}
    elif cur and a:
        rows[cur].setdefault(a, r)
for scn, m in rows.items():
    def vals(label):
        return [wn.cell(m[label], c).value for c in range(3, 13)]
    pen = vals("Robotaxi 滲透率（替代趟數÷叫車趟數）")
    cov = vals("商轉都會覆蓋率")
    ok(all(p <= c + 1e-12 for p, c in zip(pen, cov)), f"{scn} 滲透率超過覆蓋率")
    ok(all(b >= a - 1e-12 for a, b in zip(pen, pen[1:])), f"{scn} 滲透率非單調")
    sh = [w + t + o for w, t, o in zip(vals("Waymo 份額"), vals("Tesla 份額"), vals("其他份額"))]
    ok(all(abs(s - 1) < 1e-9 for s in sh), f"{scn} 營運商份額加總≠1")
    ok(all(x >= -1e-12 for x in vals("其他份額")), f"{scn} 其他份額為負")
    mg = vals("每付費英里毛利")
    ok(all(x > 0 for x in mg), f"{scn} 有年份每付費英里毛利≤0（價格低於成本，需檢查）")
    ind = vals("誘發需求倍數 (P/P2026)^ε")
    ok(abs(ind[0] - 1) < 1e-9, f"{scn} 2026 誘發倍數≠1")
# 三情境 2026 滲透率相同（校準點）
p26 = [wn.cell(rows[s]["Robotaxi 滲透率（替代趟數÷叫車趟數）"], 3).value for s in rows]
ok(max(p26) - min(p26) < 1e-12, "三情境 2026 滲透率不一致")
# 情境排序：2030 滲透率 低<基準<高
p30 = [wn.cell(rows[s]["Robotaxi 滲透率（替代趟數÷叫車趟數）"], 7).value for s in rows]
ok(p30[0] < p30[1] < p30[2], f"2030 滲透率情境順序錯誤 {p30}")

# 4. 計算頁合理範圍
wc = wbv["計算"]
cm = {wc.cell(r, 1).value: r for r in range(1, wc.max_row + 1) if wc.cell(r, 1).value}
g = lambda k, col=3: wc.cell(cm[k], col).value
ok(2000 < g("美國叫車年趟數 2026") < 6000, "美國叫車趟數超出 20–60 億合理區間")
ok(0.002 < g("美國 Robotaxi 滲透率（趟數）2026") < 0.03, "2026 滲透率超出 0.2–3%")
ok(g("加州占比上限（Q1 加州÷官方 50 萬）") >= wbv["輸入"].cell(
    [r for r in range(2, wi.max_row + 1) if wi.cell(r, 1).value == "R_CA_SHARE"][0], 6).value - 0.01,
   "加州占比假設高值超過上限")
for col in (2, 3, 4):
    ok(g("S 曲線斜率 k", col) > 0, "S 曲線 k ≤ 0（Smax 須大於 2026 都會區份額）")

# 5. 下游取數：每列三值皆為數字或年份文字、ID 唯一
wo = wbv["下游取數"]
ids = set()
for r in range(2, wo.max_row + 1):
    oid = wo.cell(r, 1).value
    if not oid or not str(oid).startswith("RT_"):
        continue
    ok(oid not in ids, f"下游 ID 重複 {oid}")
    ids.add(oid)
    for c in (3, 4, 5):
        v = wo.cell(r, c).value
        ok(isinstance(v, (int, float)) or (isinstance(v, str) and "未跨過" in v), f"下游 {oid} 欄 {c} 非數值：{v}")
for need in ("RT_US_TRIPS_2026", "RT_US_MILES_2026", "RT_PEN_2030", "RT_PRICE_2030", "RT_COST_2030", "RT_JEVONS_PRICE",
             "RT_CPM_WAYMO_2030", "RT_CPM_TSLA_CC_2030", "RT_CPM_TSLA_MY_2026", "RT_CPM_CN_2032", "RT_CPM_TSLA_FLEET_2031",
             "RT_REMOTE_WAYMO_2030", "RT_REMOTE_TSLA_CC_2030", "RT_YEAR_1USD_WAYMO", "RT_YEAR_1USD_TSLA_CC", "RT_YEAR_1USD_OPEX_CN",
             "RT_TSLA_MPI_2030", "RT_PMD_WAYMO_2030", "RT_OPPRICE_TSLA_MY_2026"):
    ok(need in ids, f"下游缺 {need}")

# 6. 敏感度：每個目標有排名 1
ws_ = wbv["敏感度"]
ranks1 = sum(1 for r in range(1, ws_.max_row + 1) if ws_.cell(r, 5).value == 1 and ws_.cell(r, 6).value)
ok(ranks1 == 5, f"敏感度排名 1 的數量={ranks1}（應為 5）")

# 7. v0.2：成本輸入
TAGS2 = TAGS
wk = wbv["成本輸入"]
for r in range(3, wk.max_row + 1):
    iid, tag = wk.cell(r, 1).value, wk.cell(r, 17).value
    if not tag:
        continue
    ok(tag in TAGS2, f"成本輸入 {iid} 標記不合法 {tag}")
    for c0 in (5, 8, 11, 14):
        vals = [wk.cell(r, c0 + j).value for j in range(3)]
        ok(all(isinstance(v, (int, float)) for v in vals), f"成本輸入 {iid} 欄 {c0} 起三值不全")
    ok(wk.cell(r, 18).value not in (None, ""), f"成本輸入 {iid} 缺出處欄")

# 8. v0.2：每英里成本堆疊
wk2 = wbv["每英里成本堆疊"]
blocks = {}
cur = None
for r in range(1, wk2.max_row + 1):
    a = wk2.cell(r, 1).value
    if a in ("Waymo", "Tesla Model Y（現行）", "Tesla Cybercab", "中國（百度／小馬平均）"):
        cur = a; blocks[cur] = {}
    elif a and str(a).startswith("年度路徑"):
        cur = None
    elif cur and a:
        blocks[cur][a] = r
ok(len(blocks) == 4, f"堆疊頁營運商區塊數={len(blocks)}")
for op, m in blocks.items():
    unit_r = [v for k, v in m.items() if str(k).startswith("＝單位營運成本")][0]
    full_r = [v for k, v in m.items() if str(k).startswith("＝每付費英里全成本")][0]
    for k, rr in m.items():
        if str(k)[:2].strip().rstrip(" ").isdigit() or str(k)[:1].isdigit():
            for c in range(3, 15):
                v = wk2.cell(rr, c).value
                ok(isinstance(v, (int, float)) and v >= -1e-12, f"堆疊 {op} {k} 欄 {c} 為負或非數值")
    for c in range(3, 15):
        u, f = wk2.cell(unit_r, c).value, wk2.cell(full_r, c).value
        ok(u > 0 and f >= u - 1e-9, f"堆疊 {op} 欄 {c} 全成本<單位成本")
    for c0 in (3, 6, 9, 12):  # 每個年份 低成本≤基準≤高成本
        v = [wk2.cell(full_r, c0 + j).value for j in range(3)]
        ok(v[0] <= v[1] + 1e-9 <= v[2] + 2e-9, f"堆疊 {op} 欄 {c0} 全成本情境順序錯誤 {v}")
    ok(wk2.cell(full_r, 10).value < wk2.cell(full_r, 4).value, f"堆疊 {op} 2030 基準全成本未低於 2026")

# 9. v0.2：接管轉成本（Tesla 接管里程逐年上升）、診斷核對、Obi 單位更正
wx = wbv["接管轉成本"]
for r in range(1, wx.max_row + 1):
    if wx.cell(r, 1).value == "Tesla 城市關鍵接管里程":
        for j in range(3):
            v = [wx.cell(r, 3 + yi * 3 + j).value for yi in range(4)]
            ok(all(b > a for a, b in zip(v, v[1:])), f"Tesla 接管里程非遞增 {v}")
wg = wbv["Tesla成本落差診斷"]
for r in range(1, wg.max_row + 1):
    b = wg.cell(r, 2).value
    if b and str(b).startswith("核對"):
        for c in (4, 5):
            v = wg.cell(r, c).value
            if isinstance(v, (int, float)):
                ok(abs(v) < 1e-6, f"診斷核對不為 0：{b} {v}")
    if b and str(b).startswith("主因（"):
        ok(bool(wg.cell(r, 4).value), "診斷主因空白")
wp = wbv["每英里價格"]
pv = {wp.cell(r, 1).value: wp.cell(r, 2).value for r in range(1, 30)}
obi_len = pv.get("隱含每趟英里＝每趟÷每英里")
cpuc_len = pv.get("CPUC 每趟付費英里＝Q2 總英里÷趟數×付費占比")
ok(obi_len and 3.0 < obi_len < 4.5, f"Obi 隱含每趟英里不合理 {obi_len}（5.72 應為每英里）")
ok(cpuc_len and 3.0 < cpuc_len < 4.5, f"CPUC 每趟付費英里不合理 {cpuc_len}")
ok(any(wi.cell(r, 1).value == "X_WAYMO_SF_MI" and wi.cell(r, 3).value == "US$/英里" for r in range(2, wi.max_row + 1)), "缺 X_WAYMO_SF_MI（每英里）")
ok(not any(wi.cell(r, 1).value == "X_WAYMO_SF_KM" for r in range(2, wi.max_row + 1)), "仍有 X_WAYMO_SF_KM")

# 10. 公開 repo 用語：不得出現工作情境字樣
for ws in wbv.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and "亞太基金可投資" in c.value:
                fails.append(f"出現工作情境字樣 {ws.title}!{c.coordinate}")

if fails:
    print("FAIL")
    for f in fails:
        print(" -", f)
    sys.exit(1)
print("PASS：全部核對通過", P)

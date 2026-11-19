"""Meta / Microsoft：在建工程（CIP）與 Wattonomics v0.3.9 CapexGW 上線年 GW 的對照（只量測）。
用法：python3 measure.py <repo 根目錄>
實績數字取自 mag/extract/docs/2026Q3/text/ 的 10-Q 純文字（行號見報告）；模型值讀 CapexGW 頁。"""
import sys, openpyxl
root = sys.argv[1]
wb = openpyxl.load_workbook(f"{root}/wattonomics/model/20261119_Wattonomics_v0.3.9.xlsx", data_only=True, read_only=True)
cg = {r[1]: r[3:13] for r in wb["CapexGW"].iter_rows(values_only=True) if isinstance(r[1], str)}
Y = {y: i for i, y in enumerate(range(2023, 2033))}
def m(k, y): return cg[k][Y[y]]
lag_on = 4.5  # In S7，README 第 157 列；月

# ---- Meta 10-Q（US$B）----
cip25, cip26q3 = 50.521, 89.499
hfs_out = 1.89 + 2.26            # 7 月轉待售（多為 CIP 與土地）＋投入 Texas 合資，上限讀法
gross25, gross26q3 = 233.726, 326.489
ad25, ad26q3, dep9m = 57.326, 73.267, 18.49
capex9m_cash, capex9m_incl = 83.055, 85.824
ai = m("meta_aish", 2026)
pit, pfac = m("meta_pit", 2026), m("meta_pfac", 2026)
bsh = pfac / (pit + pfac)
print(f"Meta 每 GW 合計 {pit+pfac:.2f}（IT {pit:.2f}／廠房 {pfac:.2f}），廠房占比 {bsh:.3f}，AI 占比 {ai}")

# A. 在建工程存量對照：模型隱含『已付未上線』存量（AI 口徑）
a24, a25, a26 = m("meta_ai", 2024), m("meta_ai", 2025), m("meta_ai", 2026)
L = lag_on / 12
b25, i25 = bsh * (a25 + 0.5 * a24), (1 - bsh) * a25 * L
ytd = a26 * 0.75
b26, i26 = bsh * (ytd + 0.5 * a25), (1 - bsh) * (a26 / 12) * lag_on
for lab, b, i, act, actadj in [("2025-12-31", b25, i25, cip25 * ai, cip25 * ai),
                               ("2026-09-30", b26, i26, cip26q3 * ai, (cip26q3 + hfs_out) * ai)]:
    print(f"A {lab}: 模型隱含 {b+i:.1f}（廠房 {b:.1f}＋IT {i:.1f}）｜實際 CIP×AI {act:.1f}（加回待售 {actadj:.1f}）｜超出 {act-(b+i):.1f}～{actadj-(b+i):.1f} → {(act-(b+i))/(pit+pfac):.2f}～{(actadj-(b+i))/(pit+pfac):.2f} GW")
for s in (0.75, 0.95):
    print(f"   AI 占比 {s}: 2026-09-30 超出 {(cip26q3*s-(b26+i26)):.1f}～{((cip26q3+hfs_out)*s-(b26+i26)):.1f}")

# B. 啟用金額：毛額增加 − CIP 增加 ＋ 報廢（累計折舊增加與折舊費用之差）
retire = dep9m - (ad26q3 - ad25)
placed = (gross26q3 - gross25) - (cip26q3 - cip25) + retire
model_on9m = m("meta_on", 2026) * (pit + pfac) * 0.75
print(f"B 9M 啟用 {placed:.1f}（報廢加回 {retire:.2f}）× AI = {placed*ai:.1f}｜模型 2026 上線 {m('meta_on',2026):.3f} GW × 每 GW × 0.75 = {model_on9m:.1f}｜比值 {placed*ai/model_on9m:.2f}")
print(f"   CIP 增加占 9M 現金資本支出 {(cip26q3-cip25)/capex9m_cash:.0%}（加回待售 {(cip26q3-cip25+hfs_out)/capex9m_cash:.0%}）")
print(f"   CIP ÷ 年化資本支出：2025 年底 {cip25/72.215*12:.1f} 月；2026-09-30 {cip26q3/(capex9m_incl/9*12)*12:.1f} 月；模型 lag_on {lag_on} 月")
print(f"   2026 指引 135–145 → Q4 隱含 {135-capex9m_incl:.1f}–{145-capex9m_incl:.1f}；模型曆年 {m('meta_cy',2026)}")

# ---- Microsoft 10-Q（US$B；FY27Q1 = 2026-07-01～09-30）----
g0, g1 = 431.767, 474.052; ad0, ad1, dep = 118.691, 127.208, 11.1
fl0, fl1 = 82.712, 89.779
ret = dep - (ad1 - ad0)
print(f"MSFT 自有 PP&E 原值增加 {g1-g0:.2f}＋報廢 {ret:.2f} = {g1-g0+ret:.2f}；現金購置 39.367；融資租賃原值增加 {fl1-fl0:.2f}（使用權資產取得 5.192）；公司口徑 45.2")
print(f"   模型 2026 曆年 AI 資本支出 {m('msft_ai',2026):.1f}；付款年 {m('msft_gw',2026):.2f} GW → 上線年 {m('msft_on',2026):.2f} GW；每 GW {m('msft_pit',2026)+m('msft_pfac',2026):.2f}")

# Siliconomics v0.21：產生 In 頁 2027 顆數輸入變更（apply_inputs.py 格式）。只寫 u_*_27（2026、2028、2029 不動）。
# 用法：python3 mk_in_edits.py <現行 v0.20 xlsx> <輸出 json>
# 依據：Andy 2026-11-18 21:13／21:20（決策 S56）；做法同影響分析情境 S3b（docs/reports/20261118_三層對帳/units-proposal/），
#       但 NVIDIA 合計取 rubin-price 報告最可能值的完整精度 6.1325115015625（S3b 用 6.13）。
import json, sys
import openpyxl
SRCX, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRCX)
ws = wb["In"]; wv = openpyxl.load_workbook(SRCX, data_only=True)["In"]
rows = {ws.cell(r, 1).value: r for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value}
val = lambda k: wv.cell(rows[k], 5).value
# rubin-price 報告 xlsx「互驗」頁 B15:D15（低／最可能／高；四路徑重疊區與基準中位數）
TOT = dict(lo=4.00367515151515, base=6.1325115015625, hi=8.38679039280948)
MIX = dict(bw=2.4, rb=6.7, ru=0.3)          # #662 修正後原始組合（Rubin＋Rubin Ultra＝Morgan Stanley 7.0）
RAW_NN = dict(amd=(1.2, 1.5, 1.8), tpu=(None, 5.5, None), trn=(2.0, 2.5, 3.5), oth=(3.0, 4.5, 6.0))   # v0.18 截頂前原始值（說明欄）
G = {p: val(f"g_{p}_27") for p in ("bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth")}
SUP = val("hbm_sup26") * (1 + val("hbm_sup27")); GROSS = val("hbm_gross")
s = sum(MIX.values())
nv = {k: {t: TOT[t] * MIX[k] / s for t in TOT} for k in MIX}
dem = (sum(nv[k]["base"] * G[k] for k in nv) + sum(RAW_NN[k][1] * G[k] for k in RAW_NN)) * GROSS / 1000
c = min(1.0, SUP / dem)
assert c == 1.0, c   # 截頂解除（預期）；若不為 1，本程式的非 NVIDIA 原始值寫法須改
TAGNV = "Derived（自建估計；輸入含 Interested-party 公司指引）"
SELF = ("2026-11-18 起基準改採自建估計（Andy 2026-11-18 21:13／21:20；決策 S56）：2027 NVIDIA 加速器合計 {tb:.4f}（低 {tl:.4f}／高 {th:.4f}）百萬封裝"
        "＝docs/reports/20261118_三層對帳/rubin-price/20261118_Rubin單價與顆數更新.md 結論 4（同名 xlsx 互驗!B15:D15；營收指引、買方資本支出、台積電、HBM 四路徑基準中位數，低高＝重疊區）；"
        "按 #662 修正後原始組合 Blackwell 2.4：Rubin 6.7：Rubin Ultra 0.3 等比例（世代組合不是自建，沿用券商組合）。2027 年 2 月重估（決策 S56）。"
        "HBM 截頂係數依 S48 定義式重算＝min(1, 供給 {sup:.4f} ÷ 原始需求 {dem:.4f} EB)＝1（截頂解除）。")
ed = []
for k in ("bw", "rb", "ru"):
    r = rows[f"u_{k}_27"]
    old = f"原標記「{ws.cell(r, 7).value}」；原來源／說明（逐字保留）：「{ws.cell(r, 8).value}」；原值 {ws.cell(r, 5).value}（＝券商原始 {MIX[k] if k != 'rb' else 7.0} × 0.9404）"
    extra = ("；#662：Morgan Stanley「2027年Rubin及Rubin Ultra晶片出貨量將接近700萬顆」（TechNews 2026-07-13 轉述）已含 Rubin Ultra，原始 Rubin 7.0 → 6.7（決策 S55）" if k == "rb" else "")
    ed.append({"id": f"u_{k}_27", "old_base": ws.cell(r, 5).value, "lo": round(nv[k]["lo"], 6), "base": round(nv[k]["base"], 6), "hi": round(nv[k]["hi"], 6), "tag": TAGNV,
               "src": SELF.format(tb=TOT["base"], tl=TOT["lo"], th=TOT["hi"], sup=SUP, dem=dem) + "券商值與來源：" + old + extra})
r = rows["u_hp_27"]
ed.append({"id": "u_hp_27", "old_base": 0, "base": 0, "tag": "Interested-party",
           "src": f"v0.21：Hopper 2027＝0 不動；HBM 截頂解除（決策 S56）。原標記「{ws.cell(r, 7).value}」；原說明（逐字保留）：「{ws.cell(r, 8).value}」"})
TAGNN = {"amd": "Interested-party／Assumed", "tpu": "Interested-party", "trn": "Assumed", "oth": "Derived／Assumed"}
for k, (lo, b, hi) in RAW_NN.items():
    r = rows[f"u_{k}_27"]
    e = {"id": f"u_{k}_27", "old_base": ws.cell(r, 5).value, "base": b, "tag": TAGNN[k],
         "src": (f"v0.21：HBM 截頂係數依 S48 定義式重算＝1（NVIDIA 2027 改自建 6.13 百萬封裝後原始需求 {dem:.4f} EB ＜ 供給 {SUP:.4f} EB；決策 S56），回到原始值"
                 f"（低／高同乘 1，與基準同一係數；修正 #662 只重算 Rubin Ultra 高值的不一致）。原標記「{ws.cell(r, 7).value}」；原說明（逐字保留）：「{ws.cell(r, 8).value}」")}
    if lo is not None: e.update(lo=lo, hi=hi)
    ed.append(e)
json.dump(ed, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("NV", {k: round(v["base"], 6) for k, v in nv.items()}, "sum", round(sum(v["base"] for v in nv.values()), 6), "dem EB", round(dem, 4), "sup", round(SUP, 4), "util", round(dem / SUP, 4), "coef", c)

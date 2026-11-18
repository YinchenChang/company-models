# 產生各情境 In 頁輸入變更 json（apply_inputs.py 格式）。只寫 2026／2027 顆數列。
# 用法：python3 mk_edits.py <情境> <輸出 json>
import json, sys
SUP27 = 3.782695 * 1.764          # hbm_sup26 × (1+hbm_sup27)，EB
GROSS = 1.15                       # hbm_gross
G = dict(bw=280, rb=288, ru=1024, hp=141, amd=360, tpu=250, trn=144, oth=200)
RAW_NN = dict(amd=(1.2, 1.5, 1.8), tpu=(None, 5.5, None), trn=(2.0, 2.5, 3.5), oth=(3.0, 4.5, 6.0))
MIX662 = dict(bw=2.4, rb=6.7, ru=0.3)          # #662 修正後的原始組合（Rubin＋Rubin Ultra＝7.0）
MIXSPLIT = dict(bw=1.40, rb=4.09, ru=0.10)     # 自建報告 A1 世代拆分（弱證據）
SC = {  # 情境 → (2027 NVIDIA 合計, 組合, 2026 設定)
    "S1": (9.4, MIX662, None), "S3": (6.13, MIX662, None), "S3m": (6.13, dict(bw=2.4, rb=7.0, ru=0.3), None), "S3split": (6.13, MIXSPLIT, None),
    "S5L": (4.00, MIX662, None), "S5H": (8.39, MIX662, None),
    "S4": (6.13, MIX662, "total"),     # 2026 NVIDIA 合計（含 Hopper）＝6.20，等比例
    "S4h": (6.13, MIX662, "br"),       # 2026 Blackwell＋Rubin＝6.20，Hopper 0.5 不動
}
TAG = "Derived（自建估計：NVIDIA 營收指引、買方資本支出、台積電、HBM 四路徑重疊中位數；多數輸入 Interested-party）"
def main():
    sc, out = sys.argv[1], sys.argv[2]
    tot, mix, y26 = SC[sc]
    s = sum(mix.values()); nv = {k: tot * v / s for k, v in mix.items()}
    dem = (sum(nv[k] * G[k] for k in nv) + sum(RAW_NN[k][1] * G[k] for k in RAW_NN)) * GROSS / 1000
    c = min(1.0, SUP27 / dem)
    note = (f"影響分析試算 {sc}（未提交）：2027 NVIDIA 合計 {tot} 百萬封裝（自建，rubin-price 報告 2026-11-18），按 {mix} 等比例；"
            f"HBM 截頂係數依 S48 定義式重算＝min(1, 供給 {SUP27:.4f} ÷ 原始需求 {dem:.4f} EB)＝{c:.6f}")
    ed = []
    for k in ("bw", "rb", "ru"):
        e = {"id": f"u_{k}_27", "base": round(nv[k] * c, 6), "tag": TAG, "src": note}
        if k == "ru": e.update(lo=0, hi=round(2 * nv[k] * c, 6))
        ed.append(e)
    ed.append({"id": "u_hp_27", "base": 0, "src": note})
    for k, (lo, b, hi) in RAW_NN.items():
        e = {"id": f"u_{k}_27", "base": round(b * c, 6), "src": note}
        if lo is not None: e.update(lo=round(lo * c, 6), hi=round(hi * c, 6))
        ed.append(e)
    if y26:
        base26 = dict(bw=5.4, rb=1.7, hp=0.5)
        if y26 == "total": f = 6.20 / sum(base26.values()); keys = base26
        else: f = 6.20 / (5.4 + 1.7); keys = dict(bw=5.4, rb=1.7)
        n26 = f"影響分析試算 {sc}（未提交）：2026 NVIDIA {'合計（含 Hopper）' if y26=='total' else 'Blackwell＋Rubin'}＝6.20 百萬封裝（NVIDIA「2,000 萬晶粒至 2026 年底」推得，Interested-party 二手），等比例係數 {f:.6f}"
        for k, v in keys.items():
            ed.append({"id": f"u_{k}_26", "base": round(v * f, 6), "src": n26})
    json.dump(ed, open(out, "w"), ensure_ascii=False, indent=1)
    print(sc, "NV", {k: round(v, 4) for k, v in nv.items()}, "dem EB", round(dem, 4), "coef", round(c, 6))
main()

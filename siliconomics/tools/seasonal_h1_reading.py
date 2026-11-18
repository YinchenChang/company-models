# Recon 第三個期中實績讀數「季節權重法」的試算（只顯示，不改任何模型；三層對帳工作單 A4，2026-11-18）。
# 讀數＝層級全年隱含 × 上半年權重（H1 weight），與實績（FY27 Q1＋Q2）比較；原主讀數（線性 × 6 ÷ 12）、第二讀數與判定邏輯都不動。
# 權重只用 FY2027 Q2（含）以前的歷史：NVIDIA 個股模型 Q 頁 FY26 Q1–Q4 的 Data Center 營收（B–E 欄）。
#   不使用 FY27 Q3（硬性規則 1：Q3 留作樣本外對帳來源）；也不使用 FY27 Q1、Q2 本身擬合權重（否則讀數與被檢驗的實績循環）。
# 用法：python3 siliconomics/tools/seasonal_h1_reading.py [Siliconomics xlsx] [NVIDIA 個股模型 xlsx]   預設 model/CURRENT 與 nvidia/dist 最新檔
import glob, os, sys
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def find_row(ws, startswith, col=1):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.replace("　", "").startswith(startswith):
            return r
    raise KeyError(startswith)


def weights(q):
    """q＝[FY26 Q1..Q4 Data Center 營收]。回傳 {low, base, high} H1 權重與推導說明。"""
    assert len(q) == 4
    base = (q[0] + q[1]) / sum(q)                      # FY26 上半年占全年（同公司前一年，n＝1）
    seq = [q[i + 1] / q[i] for i in range(3)]          # FY26 內部的逐季成長倍數
    # 固定季增倍數 g 的年度內：H1 占比＝(1＋g)÷(1＋g＋g²＋g³)＝1÷(1＋g²)。快的季增 → 後段重 → 上半年占比低。
    w = lambda g: 1 / (1 + g * g)
    return dict(base=base, low=w(max(seq)), high=w(min(seq)), seq=seq)


def main():
    sil = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "siliconomics", "model", open(os.path.join(ROOT, "siliconomics/model/CURRENT")).read().strip())
    nvf = sys.argv[2] if len(sys.argv) > 2 else sorted(glob.glob(os.path.join(ROOT, "nvidia", "dist", "*.xlsx")))[-1]
    Q = openpyxl.load_workbook(nvf, data_only=True)["Q"]
    hdr = [Q.cell(3, c).value for c in range(2, 9)]
    assert all(("FY26" in str(hdr[i])) for i in range(4)) and "FY27 Q1" in str(hdr[5]) and "FY27 Q2" in str(hdr[6]), hdr   # B–E＝FY26 四季、G、H＝FY27 Q1、Q2
    r = find_row(Q, "Data Center 營收")
    q26 = [Q.cell(r, c).value for c in (2, 3, 4, 5)]
    a1, a2 = Q.cell(r, 7).value, Q.cell(r, 8).value       # 實績（被檢驗的對象，不進權重）
    W = weights(q26)
    R = openpyxl.load_workbook(sil, data_only=True)["Recon"]
    f_impl = R.cell(find_row(R, "隱含 NVIDIA 資料中心營收（層級"), 4).value
    actual = R.cell(find_row(R, "公司純實績：已公布季累計"), 4).value
    assert abs(actual - (a1 + a2)) < 1e-9, (actual, a1, a2)
    d1_nv = R.cell(find_row(R, "已建模：NVIDIA"), 4).value   # D1 層級 NVIDIA 2026（營收口徑）
    bu_main = R.cell(find_row(R, "由下往上（主讀數）"), 4).value
    mkt = R.cell(find_row(R, "市場值（層級）"), 4).value
    print(f"FY26 Data Center 季營收 {q26}（合計 {sum(q26):.1f}）；FY26 逐季成長倍數 {[round(x, 4) for x in W['seq']]}")
    print(f"H1 權重 低／基準／高 ＝ {W['low']:.4f}／{W['base']:.4f}／{W['high']:.4f}（基準＝FY26 (Q1＋Q2)÷全年；低、高＝固定季增模型代入 FY26 最快、最慢季增）")
    print(f"實績 FY27 Q1＋Q2 ＝ {actual}（不進權重）")
    print()
    print("| 讀數 | 層級全年（$B） | 權重 | 層級上半年（$B） | 實績 ÷ 層級 − 1 |")
    print("|---|---|---|---|---|")
    print(f"| F 節 主讀數（線性 × 6÷12） | {f_impl:.2f} | 0.5000 | {f_impl * 0.5:.2f} | {actual / (f_impl * 0.5) - 1:+.1%} |")
    for k in ("low", "base", "high"):
        print(f"| F 節 季節權重法（{ {'low': '低', 'base': '基準', 'high': '高'}[k] }） | {f_impl:.2f} | {W[k]:.4f} | {f_impl * W[k]:.2f} | {actual / (f_impl * W[k]) - 1:+.1%} |")
    print(f"| D1 NVIDIA 主讀數（線性 × 6÷12） | {d1_nv:.2f} | 0.5000 | {d1_nv * 0.5:.2f} | {actual / (d1_nv * 0.5) - 1:+.1%} |")
    for k in ("low", "base", "high"):
        print(f"| D1 NVIDIA 季節權重法（{ {'low': '低', 'base': '基準', 'high': '高'}[k] }） | {d1_nv:.2f} | {W[k]:.4f} | {d1_nv * W[k]:.2f} | {actual / (d1_nv * W[k]) - 1:+.1%} |")
    print()
    print(f"D1 市場層（由下往上 ÷ 市場值）：主讀數 {bu_main / mkt:.3f}（NVIDIA 實績 × 12÷6 ＝ {actual * 2:.1f}）")
    for k in ("low", "base", "high"):
        bu = bu_main - actual * 2 + actual / W[k]
        print(f"　　季節權重法（{ {'low': '低', 'base': '基準', 'high': '高'}[k] }，NVIDIA 改為實績 ÷ 權重 ＝ {actual / W[k]:.1f}；Broadcom 與其他不動）：{bu / mkt:.3f}")
    # 一致性對照：讓層級隱含全年與 FY26 Q4 實績以固定季增連起來（FY27 四季＝Q4FY26 × (g＋g²＋g³＋g⁴) ＝ 隱含全年），解出 g 與 H1 權重
    lo, hi = 1.0, 3.0
    for _ in range(200):
        g = (lo + hi) / 2
        (lo, hi) = (g, hi) if q26[3] * (g + g ** 2 + g ** 3 + g ** 4) < f_impl else (lo, g)
    print(f"\n對照（Derived）：隱含全年 {f_impl:.1f} 與 FY26 Q4 實績 {q26[3]} 以固定季增連接 → 季增 {g - 1:+.1%}、H1 權重 {(g + g ** 2) / (g + g ** 2 + g ** 3 + g ** 4):.4f}")


if __name__ == "__main__":
    main()

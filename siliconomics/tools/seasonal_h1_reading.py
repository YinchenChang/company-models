# Recon 第三個期中實績讀數「季節權重法」的重算與核對（只顯示，不改任何模型；三層對帳工作單 A4、規則 v1.5 草案 C0 第 11 點）。
# 讀數＝已公布季累計 ÷（層級全年隱含 × 上半年權重）− 1；原主讀數（線性 × 6 ÷ 12）、第二讀數與判定邏輯都不動，c0_fail 不改。
#
# 2026-11-18 r1（Siliconomics v0.20 r1）：權重改用三個已完結會計年度 FY2024、FY2025、FY2026 的 NVIDIA 資料中心營收上半年占比；
#   低＝三年最小、基準＝三年中位數、高＝三年最大（chat 端在裁定材料中事前寫定的機械規則）。PR #646 的 n＝1 試算（FY2026 一年＋固定季增模型）只列對照。
#   資料界線：只用 FY2026（含）以前的已完結年度；FY2027 Q1、Q2 是被檢驗的實績，不進權重（否則循環）；FY2027 Q3 為樣本外保留（硬性規則 1），完全不讀。
#   資料來源（repo 內）：FY2026 四季＝nvidia/dist NVIDIA 個股模型 Q 頁 B–E 欄（Verified，8-K）；
#   FY2024、FY2025 的全年、Q3、Q4＝docs/reports/20261118_三層對帳/analysis-K/20261118_K1_歷史回測.xlsx 輸入頁 nv24*／nv25*（Verified，NVIDIA 新聞稿）。
#   本程式另核對 Siliconomics 活頁簿 In 頁 sw_h1_fy24／fy25／fy26 與 sw_h1 低／基準／高、Recon G 節各讀數與本程式重算相同（差 > 1e-9 即以代碼 1 結束）。
# 用法：python3 siliconomics/tools/seasonal_h1_reading.py [Siliconomics xlsx] [NVIDIA 個股模型 xlsx] [K1 回測 xlsx]   預設 model/CURRENT、nvidia/dist 最新檔、上述 K1 檔
import glob, os, statistics, sys
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
K1 = os.path.join(ROOT, "docs", "reports", "20261118_三層對帳", "analysis-K", "20261118_K1_歷史回測.xlsx")


def find_row(ws, startswith, col=1):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.replace("　", "").startswith(startswith):
            return r
    raise KeyError(startswith)


def k1_inputs(path):
    """K1 回測活頁簿輸入頁：鍵 → 值（只取 nv24*、nv25*）。"""
    ws = openpyxl.load_workbook(path, data_only=True)["輸入"]
    out = {}
    for row in ws.iter_rows(values_only=True):
        if row and isinstance(row[0], str) and row[0][:4] in ("nv24", "nv25"):
            out[row[0]] = row[2]
    return out


def weights(q26, k1):
    """三個已完結會計年度的上半年占比與 低／基準／高（最小／中位數／最大）。q26＝FY26 Q1–Q4。"""
    h = {"FY2024": (k1["nv24fy"] - k1["nv24q3"] - k1["nv24q4"]) / k1["nv24fy"],
         "FY2025": (k1["nv25fy"] - k1["nv25q3"] - k1["nv25q4"]) / k1["nv25fy"],
         "FY2026": (q26[0] + q26[1]) / sum(q26)}
    v = list(h.values())
    return h, dict(low=min(v), base=statistics.median(v), high=max(v))


def legacy_n1(q):
    """PR #646 的 n＝1 試算（只列對照）：基準＝FY26 上半年占比；低、高＝固定季增模型 1÷(1＋g²) 代入 FY26 最快、最慢季增。"""
    seq = [q[i + 1] / q[i] for i in range(3)]
    w = lambda g: 1 / (1 + g * g)
    return dict(low=w(max(seq)), base=(q[0] + q[1]) / sum(q), high=w(min(seq)))


def main():
    sil = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "siliconomics", "model", open(os.path.join(ROOT, "siliconomics/model/CURRENT")).read().strip())
    nvf = sys.argv[2] if len(sys.argv) > 2 else sorted(glob.glob(os.path.join(ROOT, "nvidia", "dist", "*.xlsx")))[-1]
    k1f = sys.argv[3] if len(sys.argv) > 3 else K1
    Q = openpyxl.load_workbook(nvf, data_only=True)["Q"]
    hdr = [Q.cell(3, c).value for c in range(2, 9)]
    assert all(("FY26" in str(hdr[i])) for i in range(4)) and "FY27 Q1" in str(hdr[5]) and "FY27 Q2" in str(hdr[6]), hdr   # B–E＝FY26 四季、G、H＝FY27 Q1、Q2
    r = find_row(Q, "Data Center 營收")
    q26 = [Q.cell(r, c).value for c in (2, 3, 4, 5)]          # 只讀 B–E（FY26）；FY27 欄只讀 G、H 作核對，I 欄以後（Q3 指引、Q4 假設）不讀
    a1, a2 = Q.cell(r, 7).value, Q.cell(r, 8).value             # 被檢驗的實績，不進權重
    k1 = k1_inputs(k1f)
    H, W = weights(q26, k1)
    wb = openpyxl.load_workbook(sil, data_only=True)
    R = wb["Recon"]
    f_impl = R.cell(find_row(R, "隱含 NVIDIA 資料中心營收（層級"), 4).value
    actual = R.cell(find_row(R, "公司純實績：已公布季累計"), 4).value
    assert abs(actual - (a1 + a2)) < 1e-9, (actual, a1, a2)
    d1_nv = R.cell(find_row(R, "已建模：NVIDIA"), 4).value
    bu_main = R.cell(find_row(R, "由下往上（主讀數）"), 4).value
    mkt = R.cell(find_row(R, "市場值（層級）"), 4).value

    lab = {"low": "低", "base": "基準", "high": "高"}
    print("NVIDIA 資料中心營收上半年占全年（[Derived]；Verified 實績相減）：" + "、".join(f"{y} {v:.1%}" for y, v in H.items()))
    print(f"權重（n＝3）低＝最小 {W['low']:.4f}、基準＝中位數 {W['base']:.4f}、高＝最大 {W['high']:.4f}；FY2027 任何季度不進權重")
    L = legacy_n1(q26)
    print(f"對照（PR #646 n＝1 試算，已不用）：低／基準／高 {L['low']:.4f}／{L['base']:.4f}／{L['high']:.4f}")
    print(f"實績 FY27 Q1＋Q2 ＝ {actual}（被檢驗對象，不進權重）\n")
    rd = {}
    print("| 讀數 | 權重 | 實績 ÷ 層級 − 1（D1 市場層為比值） |\n|---|---|---|")
    print(f"| F 節 主讀數（現行，線性） | 0.5000 | {actual / (f_impl * 0.5) - 1:+.1%} |")
    for k in ("low", "base", "high"):
        rd["f_" + k] = actual / (f_impl * W[k]) - 1
        print(f"| F 節 季節權重法（{lab[k]}） | {W[k]:.4f} | {rd['f_' + k]:+.1%} |")
    print(f"| D1 NVIDIA 主讀數（現行） | 0.5000 | {actual / (d1_nv * 0.5) - 1:+.1%} |")
    for k in ("low", "base", "high"):
        rd["d1nv_" + k] = actual / (d1_nv * W[k]) - 1
        print(f"| D1 NVIDIA 季節權重法（{lab[k]}） | {W[k]:.4f} | {rd['d1nv_' + k]:+.1%} |")
    print(f"| D1 市場層 主讀數（現行） | — | {bu_main / mkt:.3f} |")
    for k in ("low", "base", "high"):
        rd["d1m_" + k] = (bu_main - actual * 2 + actual / W[k]) / mkt
        print(f"| D1 市場層 季節權重法（{lab[k]}） | {W[k]:.4f} | {rd['d1m_' + k]:.3f} |")

    # 核對活頁簿（v0.20 r1 起有 In 頁 sw_h1* 與 Recon G 節）
    def nm(n):
        if n not in wb.defined_names:
            return None
        sh, ref = list(wb.defined_names[n].destinations)[0]
        return wb[sh][ref.replace("$", "")].value
    if nm("sw_h1") is None:
        print("\n活頁簿沒有 sw_h1（v0.20 r1 以前的版本）：只印重算結果、不核對"); return
    bad = []
    for n, v in (("sw_h1_fy24", H["FY2024"]), ("sw_h1_fy25", H["FY2025"]), ("sw_h1_fy26", H["FY2026"]),
                 ("sw_h1_lo", W["low"]), ("sw_h1", W["base"]), ("sw_h1_hi", W["high"])):
        if abs(nm(n) - v) > 1e-9:
            bad.append(f"In {n}：活頁簿 {nm(n)} ≠ 重算 {v}")
    for pre, lab0 in (("f_", "季節讀數 F 節"), ("d1nv_", "季節讀數 D1 NVIDIA"), ("d1m_", "季節讀數 D1 市場層")):
        for k in ("low", "base", "high"):
            x = R.cell(find_row(R, f"{lab0}（{lab[k]}）"), 4).value
            if abs(x - rd[pre + k]) > 1e-9:
                bad.append(f"Recon {lab0}（{lab[k]}）：活頁簿 {x} ≠ 重算 {rd[pre + k]}")
    if bad:
        print("\n核對不符：\n" + "\n".join(bad)); sys.exit(1)
    print(f"\n核對：活頁簿 In 頁 sw_h1*（6 格）與 Recon G 節 9 個讀數與本程式重算相同（{os.path.basename(sil)}）")


if __name__ == "__main__":
    main()

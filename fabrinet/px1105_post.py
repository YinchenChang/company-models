# 2026-11-05 收盤更新輪（工作單 px1105，Andy 2026-11-06）共用後處理：建置腳本在存檔前呼叫 apply(wb)。只改標籤與 README 文字，不改任何數值公式。
# 各公司資料夾放同一份（公司自有，不屬於 tools/）。
#   (1) 情境名稱改回「情境 A（樂觀）／B（基準）／C（保守）」（Andy 2026-11-06 03:14，取代第三輪修正第 23 點的順利／平平／不如預期）。
#       「平平」作一般形容詞（例：網通 CPO 情境「A＝替代快、B＝平平、C＝慢」）不改；README 版本沿革列（v0.… 開頭）是當時紀錄，不改。
#   (2) 版本沿革列不得用活公式（第三輪修正第 37 點）：仍是公式的舊版沿革列，改成當版成品顯示的固定文字（FROZEN；鍵＝公式開頭）。
import re

SCN = {"順利": "樂觀", "平平": "基準", "不如預期": "保守", "下檔": "保守"}
_P1 = re.compile(r"(情境(?:權重)? ?[ABC][^（）\"]{0,12})（(順利|平平|不如預期|下檔)）")
_P2 = re.compile(r"(?<![A-Za-z])([ABC])（(順利|平平|不如預期)）")
_P3 = re.compile(r"([ABC])＝(順利|平平|不如預期)(?!、C＝慢)")
_P4 = re.compile(r"(情境 [ABC])：(順利|平平|不如預期)")
_FIX = (("順利、平平、不如預期", "樂觀、基準、保守"), ("A 順利／B 平平／C 不如預期", "A 樂觀／B 基準／C 保守"),
        ("A 順利、B 平平、C 不如預期", "A 樂觀、B 基準、C 保守"), ("不如預期～順利", "保守～樂觀"))


def _hist(ws, v):
    t = v[2:] if v.startswith('="') else v
    return ws.title == "README" and t.startswith(("v0.", "版本："))


def rename(s):
    w = _P1.sub(lambda m: m.group(1) + "（" + SCN[m.group(2)] + "）", s)
    w = _P2.sub(lambda m: m.group(1) + "（" + SCN[m.group(2)] + "）", w)
    w = _P3.sub(lambda m: m.group(1) + "＝" + SCN[m.group(2)], w)
    w = _P4.sub(lambda m: m.group(1) + "：" + SCN[m.group(2)], w)
    for a, b in _FIX:
        w = w.replace(a, b)
    return w


def apply(wb, frozen=None, a1_append=None):
    frozen = FROZEN if frozen is None else frozen
    n_scn = n_frz = 0
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str):
                    continue
                if ws.title == "README" and frozen and v.startswith('="v0.'):
                    for k, t in frozen.items():
                        if v.startswith(k):
                            c.value = t; n_frz += 1; break
                    continue
                if _hist(ws, v):
                    continue
                w = rename(v)
                if w != v:
                    c.value = w; n_scn += 1
    if frozen:
        assert n_frz == len(frozen), (n_frz, len(frozen))
    if a1_append:
        wb["README"]["A1"] = wb["README"]["A1"].value + a1_append
    return f"px1105_post：情境名稱 {n_scn} 格、版本沿革固定文字 {n_frz} 列"


# 舊版沿革列（v0.… 開頭的公式）→ 上一版成品顯示的固定文字（2026-11-04 收盤口徑，當時紀錄）
FROZEN = {'="v0.3.3（2026-11-06）：第三輪修正（合併前）。(1) 現價與 ': 'v0.3.3（2026-11-06）：第三輪修正（合併前）。(1) 現價與 Comparable 同業報價統一為 2026-11-04 收盤（Yahoo Finance 日線收盤）：現價 314.94（10-23）→ 341.64；同業前瞻 P/E＝頁面前瞻 P/E × 11-04 收盤 ÷ '
                                             '頁面價格（共識 EPS 不變）；Lumentum 已與 Marvell 簽換股合併協議（股價跟隨換股價值），自同業組（Comparable 與相對倍數調整）剔除，中位數 16.65 → 15.01、相對倍數調整係數 1.61 → 1.65。(2) README 頁舊「重點觀察」刪除。(3) '
                                             '情境驅動收斂法列名改中性、百分比格式；情境名稱統一 A（順利）／B（平平）／C（不如預期）；Val「上檔空間級距」註明六級只供分組；HTML DCF 彙總主表直接顯示 stub 兩列，橫幅標明美元。Stub 期 FCFF 維持含 stub 期的會計年度 FY2027（2026-07～2027-06）× '
                                             'stub_frac。DCF 目標價 192.68 → 192.66、加權目標價 432.84 → 409.01（上檔 +19.7%，中立；原 +37.4%，買進）。',
 '="v0.3.1（2026-11-05）：β 改用自算產業 β（Andy 202': 'v0.3.1（2026-11-05）：β 改用自算產業 β（Andy 2026-11-05 22:55）：US 網通／光通訊同業 7 家（ANET、CSCO、CIEN、COHR、LITE、FN、AAOI）無槓桿 β 中位數 1.67（Yahoo Finance 月報酬 2021-11～2026-10 對 '
                                             'S&P 500 迴歸，Hamada 去槓桿、現金修正），Blume 調整後 1.449，依市值 D/E 重新加槓桿 β＝1.456，WACC 12.24%；取代 Damodaran 美國 Electronics (General) 0.94。β 低／高案例改為中位數 × 0.8 再 '
                                             'Blume／中位數未調整。評價日維持 2026-12-31。其他輸入不變。'}

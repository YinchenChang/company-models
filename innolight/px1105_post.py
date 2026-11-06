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
FROZEN = {'="v0.4.3（2026-11-06）：第三輪修正（合併前）。(1) 現價與 ': 'v0.4.3（2026-11-06）：第三輪修正（合併前）。(1) 現價與 Comparable 同業報價統一為 2026-11-04 收盤（Yahoo Finance 日線收盤）：現價 553.00（10-30）→ 533.38；原同業報價日期不一（天孚 09-30、光迅 08-18、華工 '
                                             '08-07），改為同業前瞻 P/E＝頁面前瞻 P/E × 11-04 收盤 ÷ 頁面價格（共識 EPS 不變），中位數 45.08 → 24.37，Comparable（相對倍數調整後）1,674.46 → 905.02。(2) Stub 期 FCFF 改用 4Q26（Q 頁 4Q26 '
                                             '假設）＝2.61（原 FY2027 全年 × stub_frac 5.34）。(3) DCF 的 EBIT＝稅前利益 − 非營業損益（投資收益＋公允價值變動 − 財務費用，FY2025 年報占營收 0.55%），避免與淨現金重複；EPS 不變。(4) README '
                                             '頁舊「重點觀察」刪除；情境驅動收斂法列名改中性、百分比格式；情境名稱統一 A（順利）／B（平平）／C（不如預期）；Val「上檔空間級距」註明六級只供分組；g 受人民幣 rf 上限，Sens 列名改「永續成長率 g −1pt／上限 rf」；HTML DCF 彙總主表直接顯示 stub '
                                             '兩列，橫幅標明人民幣。DCF 目標價 487.44 → 477.02、加權目標價 1,080.95 → 691.02（上檔 +29.6%，買進；原 +95.5%，強力買進）。',
 '="v0.4.1（2026-11-05）：β 改用自算產業 β（Andy 202': 'v0.4.1（2026-11-05）：β 改用自算產業 β（Andy 2026-11-05 22:55）：CN 光通訊同業 6 家（300308、300502、300394、002281、000988、603083）無槓桿 β 中位數 2.69（Yahoo Finance 月報酬 '
                                             '2021-11～2026-10 對上證綜指迴歸，Hamada 去槓桿、現金修正），Blume 調整後 2.132，依市值 D/E 重新加槓桿 β＝2.138，WACC 14.21%；取代 Damodaran Global Telecom. Equipment 1.30。β 低／高案例改為中位數 × '
                                             '0.8 再 Blume／中位數未調整。評價日維持 2026-12-31。其他輸入不變。'}

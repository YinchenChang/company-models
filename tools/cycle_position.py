"""設備股共用的循環位置規則（Advantest、Disco 等設備公司模型共用，確保同一把尺）。

規則（2026-10-23 擁有者核准）：每個情境的見頂年度取自設備市場模型（markets/equipment 成品 Mkt 頁的「見頂年度」列），
不在各公司模型裡另訂。會計年度 FY(t)（t 年 4 月～t+1 年 3 月，3/4 落在日曆年 t）對照日曆年 t：
  t ＜ 見頂年 → 中段（中位數）；t ＝ 見頂年 → 見頂（25 分位）；t ＞ 見頂年 → 谷底（75 分位）。
用法：from cycle_position import peak_years, excel_rule
"""
import glob
import os

import openpyxl

SC = ("A", "B", "C")


def peak_years(root):
    """回傳 ({情境: 見頂年度}, 來源字串)。讀 markets/equipment/dist 內唯一的 xlsx（快取值）。"""
    fs = glob.glob(os.path.join(root, "markets", "equipment", "dist", "*.xlsx"))
    assert len(fs) == 1, fs
    ws = openpyxl.load_workbook(fs[0], data_only=True)["Mkt"]
    out, src, sc = {}, {}, None
    for r in ws.iter_rows(min_row=4, max_row=40):
        a = r[0].value
        if isinstance(a, str) and a.startswith("情境 ") and len(a.strip()) == 4:
            sc = a.strip()[-1]
            continue
        if sc and isinstance(a, str) and a.strip() == "見頂年度" and sc not in out:
            out[sc] = int(r[2].value)
            src[sc] = f"{os.path.relpath(fs[0], root)}｜Mkt!C{r[0].row}"
    assert sorted(out) == list(SC), out
    return out, src


def classify(fy, peak):
    """Python 版（供總覽抽數使用）：回傳 '中段'／'見頂'／'谷底'。"""
    return "中段" if fy < peak else ("見頂" if fy == peak else "谷底")


def excel_rule(fy_cell, peak_name):
    """Excel 版公式字串（不含等號）：與 classify 相同。"""
    return f'IF({fy_cell}<{peak_name},"中段",IF({fy_cell}={peak_name},"見頂","谷底"))'


def rule_text(peaks):
    """規則說明文字（見頂年度由 peak_years 讀入，不寫死）。"""
    pk = "、".join(f"{s} {peaks[s]}" for s in SC)
    return ("循環位置（設備股共用規則）：各情境的見頂年度取自設備市場模型（markets/equipment 的 Mkt 頁「見頂年度」：" + pk + "）；"
            "會計年度 FY(t) 對照日曆年 t（FY(t) 有 3/4 落在 t 年）：t 早於見頂年＝中段（中位數）、等於＝見頂（25 分位）、晚於＝谷底（75 分位）")


def fy2027_text(peaks):
    """FY2027 各情境的判定（說明用）。"""
    return "、".join(f"{s} {classify(2027, peaks[s])}" for s in SC)


# ---- 會計年度與日曆年錯開的公司（例：NVIDIA，FY(t)＝t−1 年 2 月～t 年 1 月，約等於日曆年 t−1） ----
# 設備股的函式（上方）維持原狀；錯開的公司用下面的包裝函式，lag＝會計年度比對應日曆年晚幾年（NVIDIA＝1）。
# 判定改成 classify(fy − lag, peak)：FY(t) 對照日曆年 t−lag 再與見頂年度比較。
def classify_lag(fy, peak, lag=1):
    """Python 版：FY(fy) 對照日曆年 fy−lag；其餘規則同 classify。"""
    return classify(fy - lag, peak)


def excel_rule_lag(fy_cell, peak_name, lag=1):
    """Excel 版公式字串（不含等號）：與 classify_lag 相同。"""
    return excel_rule(f"({fy_cell}-{lag})", peak_name)

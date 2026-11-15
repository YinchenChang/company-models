#!/usr/bin/env python3
"""估值頁（Val）建置快照（v0.1.6）：情境 A／C 與敏感度（龍捲風）。

改 Inputs（藍字，F 欄；'lo'／'hi'＝該列低／高欄）→ LibreOffice 重算整本活頁簿 → 讀 Val 頁值。只執行與讀取 Excel（CLAUDE.md 第 1 節）。
用法（在 zhipu/ 內）：
  python3 tools/val_snapshots.py scen <xlsx>   → data/val_scenarios.yaml（情境 A／B／C 的 Val 第 3 節 B 列；B 只供過期檢查）
  python3 tools/val_snapshots.py sens <xlsx>   → data/val_sens.yaml（各案例低／高值的加權目標價；xlsx 須已含情境快照）
之後以 builder/build.py --base <xlsx> 重建，快照才寫入 Val 頁。
"""
from __future__ import annotations

import datetime as dt
import shutil
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import openpyxl
import yaml

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "builder"))
from common import lo_recalc  # noqa: E402

SCN = {"A": {"g_task_2h26": "hi"}, "B": {}, "C": {"g_task_2h26": "lo"}}
SRCK = {"rev": "營收淨額", "ebit": "EBIT", "cx": "非算力資本支出", "ma": "併購"}
# (鍵, 名稱, 白話, 低值設定, 高值設定, 低值說明, 高值說明, 低值白話, 高值白話)
CASES = [
    ("wacc", "WACC ±1pt（val_wacc_adj）", "折現率（WACC）", {"val_wacc_adj": "lo"}, {"val_wacc_adj": "hi"}, "WACC −1pt", "WACC +1pt", "折現率低 1 個百分點", "折現率高 1 個百分點"),
    ("g", "永續成長率 2%／4%（val_g）", "永續成長率", {"val_g": "lo"}, {"val_g": "hi"}, "g＝2%", "g＝4%", "永續成長 2%", "永續成長 4%"),
    ("beta", "β 1.0／1.6（val_beta）", "股價波動係數（β）", {"val_beta": "lo"}, {"val_beta": "hi"}, "β＝1.0", "β＝1.6", "波動係數 1.0", "波動係數 1.6"),
    ("erp", "ERP 4%／6%（val_erp）", "股權風險溢酬", {"val_erp": "lo"}, {"val_erp": "hi"}, "ERP＝4%", "ERP＝6%", "風險溢酬 4%", "風險溢酬 6%"),
    ("extg", "延伸期終點營收年增 5%／15%（val_ext_g33）", "2031–2033 營收成長終點", {"val_ext_g33": "lo"}, {"val_ext_g33": "hi"}, "5%", "15%", "2033 年營收年增 5%", "2033 年營收年增 15%"),
    ("extm", "延伸期營業利益率加碼 0／+50pt（val_ext_m_add）", "2031–2033 營業利益率改善", {"val_ext_m_add": "lo"}, {"val_ext_m_add": "hi"}, "維持 FY2030", "+50pt（2033 年約 +14%）",
     "利益率維持 2030 年水準", "2033 年利益率轉為約 +14%"),
    ("wdcf", "DCF 權重 25%／75%（val_w_dcf）", "DCF 與同業法的權重", {"val_w_dcf": "lo"}, {"val_w_dcf": "hi"}, "DCF 25%", "DCF 75%", "DCF 占 25%", "DCF 占 75%"),
    ("tax", "稅率 15%／25%（val_tax）", "所得稅率", {"val_tax": "lo"}, {"val_tax": "hi"}, "15%", "25%", "稅率 15%", "稅率 25%"),
    ("wc", "營運資金率 −5%／+5%（val_wc_rate）", "營運資金占營收增量", {"val_wc_rate": "lo"}, {"val_wc_rate": "hi"}, "−5%", "+5%", "預收款增加（−5%）", "營運資金占用（+5%）"),
    ("adj", "相對倍數調整：不調整／依 FY2030 利益率調整（val_comp_adj_sw）", "同業倍數是否依獲利差異打折", {"val_comp_adj_sw": "lo"}, {"val_comp_adj_sw": "hi"}, "不調整", "調整",
     "不打折", "依 2030 年虧損程度打折"),
    ("task", "2H26 API 任務數成長率低端／高端（INP_038；情境 A／C 的驅動）", "API 任務數成長", {"g_task_2h26": "lo"}, {"g_task_2h26": "hi"}, "低端", "高端", "任務數成長慢", "任務數成長快"),
    ("pxgw", "每 GW 算力價格每年 −25%／不變（px_gw_chg）", "算力採購價格年變動", {"px_gw_chg": "lo"}, {"px_gw_chg": "hi"}, "−25%／年", "不變", "算力每年降價 25%", "算力不降價"),
    ("rdfee", "研發項下算力費占研發開支 60%／80%（rd_fee_share）", "研發中的算力費占比", {"rd_fee_share": "lo"}, {"rd_fee_share": "hi"}, "60%", "80%", "研發算力占比低", "研發算力占比高"),
]


def recalc_with(model: Path, ov: dict, tag: str) -> Path:
    wb = openpyxl.load_workbook(model)
    ws = wb["Inputs"]
    rows = {ws.cell(r, 2).value: r for r in range(5, ws.max_row + 1) if ws.cell(r, 2).value}
    for k, v in ov.items():
        r = rows[k]
        ws[f"F{r}"].value = ws[f"G{r}"].value if v == "lo" else (ws[f"H{r}"].value if v == "hi" else v)
    d = Path(tempfile.mkdtemp(prefix=f"zpval_{tag}_"))
    raw = d / f"{tag}.xlsx"
    wb.save(raw)
    return lo_recalc(raw, d / "out")


def run_many(model: Path, jobs: dict, reader, workers=2):
    def one(item):
        tag, ov = item
        p = recalc_with(model, ov, tag)
        try:
            return tag, reader(p)
        finally:
            shutil.rmtree(p.parent.parent, ignore_errors=True)
    with ThreadPoolExecutor(workers) as ex:
        return dict(ex.map(one, jobs.items()))


def labels(ws):
    out = {}
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if isinstance(a, str):
            out.setdefault(a, r)
    return out


def read_src(path):
    ws = openpyxl.load_workbook(path, data_only=True)["Val"]
    lab = labels(ws)
    out = {}
    for k, pre in SRCK.items():
        r = next(r for t, r in lab.items() if t.startswith(pre) and t.endswith("（B）"))
        out[k] = [float(ws.cell(r, c).value or 0) for c in range(3, 8)]
    r = next(r for t, r in lab.items() if t.startswith("營收：FY2025") and "（B；" in t)
    out["fy"] = [float(ws.cell(r, c).value) for c in range(3, 6)]
    return out


def read_tp(path):
    ws = openpyxl.load_workbook(path, data_only=True)["Val"]
    r = next(r for t, r in labels(ws).items() if t.startswith("加權目標價＝"))
    return float(ws.cell(r, 3).value)


def scen(model):
    res = run_many(model, SCN, read_src)
    desc = {"A": "2H26 API 任務數成長率取 Inputs 高端（INP_038 高欄）", "B": "基準", "C": "2H26 API 任務數成長率取 Inputs 低端（INP_038 低欄）"}
    out = dict(note="tools/val_snapshots.py scen 產生；值＝LibreOffice 重算後 Val 第 3 節 B 列（各情境改 Inputs 後）；勿手改",
               model=model.name, date=dt.date.today().isoformat(), scenarios={})
    for s in SCN:
        v = res[s]
        out["scenarios"][s] = dict(id=f"VAL_SCN_{s}", desc=desc[s], overrides={k: str(x) for k, x in SCN[s].items()}, values=v,
                                   check=sum(v["rev"]) + sum(v["ebit"]))
        print(s, {k: [round(x, 2) for x in vv] for k, vv in v.items()})
    (ROOT / "data" / "val_scenarios.yaml").write_text(yaml.safe_dump(out, allow_unicode=True, sort_keys=False, width=200), encoding="utf-8")


def sens(model):
    jobs = {"base": {}}
    for c in CASES:
        jobs[c[0] + "_lo"], jobs[c[0] + "_hi"] = c[3], c[4]
    tp = run_many(model, jobs, read_tp)
    out = dict(note="tools/val_snapshots.py sens 產生；值＝改寫 Inputs 後整本 LibreOffice 重算的加權目標價（港元）；勿手改",
               model=model.name, date=dt.date.today().isoformat(), base_tp=tp["base"], cases=[])
    for k, name, plain, lo, hi, ll, hl, lp, hp in CASES:
        out["cases"].append(dict(key=k, name=name, plain=plain, tp_lo=tp[k + "_lo"], tp_hi=tp[k + "_hi"], lo_label=ll, hi_label=hl, lo_plain=lp, hi_plain=hp,
                                 lo_set={a: str(b) for a, b in lo.items()}, hi_set={a: str(b) for a, b in hi.items()}))
        print(f'{name}: {tp[k + "_lo"]:.2f} / {tp[k + "_hi"]:.2f}（基準 {tp["base"]:.2f}）')
    (ROOT / "data" / "val_sens.yaml").write_text(yaml.safe_dump(out, allow_unicode=True, sort_keys=False, width=200), encoding="utf-8")


if __name__ == "__main__":
    {"scen": scen, "sens": sens}[sys.argv[1]](Path(sys.argv[2]).resolve())

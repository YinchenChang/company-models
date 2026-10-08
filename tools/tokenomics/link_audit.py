#!/usr/bin/env python3
"""Tokenomics 連接審視（各公司統一執行）：盤點每家公司模型如何取用 Tokenomics，對照下游資料契約。

用法：
  python3 tools/tokenomics/link_audit.py --tokenomics <Tokenomics clone> [--out docs/reports/YYYYMMDD_tokenomics_link_audit.md]

每家公司回報：
  - 取數方式：官方快照（data/tokenomics_snapshot_*.json，由 import_tokenomics.py 產生）／手工 JSON（data/permw_tokenomics_*.json）／無。
  - 釘住的 Tokenomics 版本（檔名、commit）對 Tokenomics clone 現行 model/CURRENT 的落差。
  - 使用到的名稱：IF_／L1_（合規）；其他前綴（DRV_、CAL_、B4_、B5_、CST_、IDX_、IF_Hdr*）或工作表直接引用（Inputs!、DC_Cost!、Serving!…）列為違規（契約第 2 條第 1 項）；公司模型自己的同前綴名稱（例如行事曆 CAL_*）不計，只有在 Tokenomics 中真實存在的名稱才算。工作表直接引用需人工確認是取數還是文件註記。
  - 名稱是否仍存在於現行 Tokenomics；快照值對現行值的漂移（相對誤差 > 1e-9 者列出）。
  - 是否取四層瀑布（IFW_）與用途欄（IFC_）（v5.29 起；契約第 2 條第 2 項）。
只讀取、只回報；不改任何公司模型。
"""
import argparse, datetime, glob, json, os, re, sys
from pathlib import Path

import openpyxl

OK_PREFIX = ("IF_", "IFW_", "IFC_", "L1_")
BAD_PREFIX = ("DRV_", "CAL_", "B4_", "B5_", "CST_", "IDX_", "TRN_", "TR_", "GOV_")
SHEET_REF = re.compile(r"\b(Inputs|DC_Cost|Serving|Spec_Rack|Arch|Workload|Calib|Perf|Unit_Cost|Train_In|Training|Cap_In|Theory_Rev|Alloc|Alloc_In|Fleet_1GW|Amortize)!")
NAME_RE = re.compile(r"\b((?:IF|IFW|IFC|L1|DRV|CAL|B4|B5|CST|IDX|TRN|TR|GOV)_[A-Za-z0-9_]+)")
SCAN_EXT = (".json", ".py", ".js", ".txt", ".yaml", ".yml")
SKIP_DIRS = ("docs", "dist", "node_modules", ".git", "__pycache__", "legacy")


def load_current(tk):
    cur = (Path(tk) / "model" / "CURRENT").read_text(encoding="utf-8").strip()
    wb = openpyxl.load_workbook(Path(tk) / "model" / cur, data_only=True)
    names = {}
    for n, dn in wb.defined_names.items():
        try:
            dest = list(dn.destinations)
        except Exception:
            continue
        if not dest:
            continue
        sheet, ref = dest[0]
        ws = wb[sheet]
        ref = ref.replace("$", "")
        cells = ws[ref]
        if isinstance(cells, tuple):
            vals = [c.value for row in (cells if isinstance(cells[0], tuple) else (cells,)) for c in row]
        else:
            vals = [cells.value]
        names[n] = vals
    try:
        import subprocess
        commit = subprocess.run(["git", "-C", tk, "rev-parse", "--short", "HEAD"], capture_output=True, text=True).stdout.strip()
    except Exception:
        commit = "?"
    return cur, commit, names


def scan_company(d, tk_names):
    used, bad_names, sheet_refs = set(), set(), set()
    for root, dirs, files in os.walk(d):
        dirs[:] = [x for x in dirs if x not in SKIP_DIRS]
        for f in files:
            if not f.endswith(SCAN_EXT):
                continue
            p = Path(root) / f
            try:
                t = p.read_text(encoding="utf-8", errors="ignore")
            except Exception:
                continue
            for m in NAME_RE.finditer(t):
                n = m.group(1)
                if n.startswith("IF_Hdr") or (n.startswith(BAD_PREFIX) and n in tk_names):
                    bad_names.add((n, str(p.relative_to(d))))
                elif n.startswith(OK_PREFIX):
                    if n.endswith("_"):                      # 程式中的前綴樣式（例如 IF_TokGW_ + 層級）
                        used.update(x for x in tk_names if x.startswith(n))
                        if not any(x.startswith(n) for x in tk_names):
                            used.add(n)
                    else:
                        used.add(n)
            for m in SHEET_REF.finditer(t):
                sheet_refs.add((m.group(0), str(p.relative_to(d))))
    return used, bad_names, sheet_refs


def snapshot_info(d):
    snaps = sorted(glob.glob(str(Path(d) / "data" / "tokenomics_snapshot_*.json")))
    manual = sorted(glob.glob(str(Path(d) / "data" / "permw_tokenomics_*.json")))
    info = {"mode": "無", "pinned": "—", "commit": "—", "items": {}}
    if snaps:
        s = json.load(open(snaps[-1], encoding="utf-8"))
        src = s.get("source", {})
        info.update(mode=f"官方快照（{Path(snaps[-1]).name}）", pinned=src.get("file", "?"), commit=str(src.get("commit", "?"))[:7],
                    items={k: v.get("values") for k, v in s.get("items", {}).items()})
    elif manual:
        s = json.load(open(manual[-1], encoding="utf-8"))
        tk = s.get("tokenomics", {})
        info.update(mode=f"手工 JSON（{Path(manual[-1]).name}）", pinned=tk.get("file", "?"), commit=str(tk.get("commit", "?"))[:7])
    return info


def flat(v):
    if isinstance(v, dict):
        out = []
        for k in v.values():
            out += flat(k)
        return out
    return [v]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tokenomics", required=True)
    ap.add_argument("--out")
    a = ap.parse_args()
    cur, commit, names = load_current(a.tokenomics)
    today = datetime.date.today().isoformat()
    lines = [f"# Tokenomics 連接審視（{today}）", "",
             f"對照對象：Tokenomics `model/CURRENT`＝`{cur}`（{commit}）。契約：`YinchenChang/Tokenomics` `docs/plan/Tokenomics_downstream_contract.md`。只回報，不改模型。", ""]
    summary = ["| 公司 | 取數方式 | 釘住版本 | 使用名稱數 | 不存在於現行 | 違規引用 | 值漂移（名稱數） | 四層瀑布／用途欄 |", "|---|---|---|---|---|---|---|---|"]
    detail = []
    for c in sorted(p.name for p in Path(".").iterdir() if p.is_dir() and (p / "README.md").exists() and p.name not in ("tools",)):
        used, bad, srefs = scan_company(c, names)
        info = snapshot_info(c)
        missing = sorted(n for n in used if n not in names)
        drift = []
        for n, vals in info["items"].items():
            if n not in names or vals is None:
                continue
            a_, b_ = [x for x in flat(vals) if isinstance(x, (int, float))], [x for x in names[n] if isinstance(x, (int, float))]
            if len(a_) != len(b_):
                drift.append((n, "形狀不同"))
                continue
            for x, y in zip(a_, b_):
                if abs(x - y) > 1e-9 * max(1, abs(x), abs(y)):
                    drift.append((n, f"{x} → {y}"))
                    break
        wf = "是" if any(n.startswith("IFW_") for n in used) else "否"
        uc = "是" if any(n.startswith("IFC_") for n in used) else "否"
        pinned = info["pinned"]
        stale = "" if pinned == cur else "（落後）"
        summary.append(f"| {c} | {info['mode']} | {pinned}{stale} | {len(used)} | {len(missing)} | {len(bad) + len(srefs)} | {len(drift)} | {wf}／{uc} |")
        detail.append(f"## {c}")
        detail.append(f"- 取數方式：{info['mode']}；釘住 `{pinned}`（{info['commit']}）{stale}。")
        detail.append(f"- 使用名稱（{len(used)}）：{', '.join(sorted(used)) or '無'}")
        if missing:
            detail.append(f"- **不存在於現行 Tokenomics**：{', '.join(missing)}")
        if bad or srefs:
            detail.append("- **違規引用（契約第 2 條第 1 項）**：" + "；".join(sorted({f'`{n}`（{f}）' for n, f in bad} | {f'`{s}`（{f}）' for s, f in srefs})))
        if drift:
            detail.append("- **快照值對現行值的漂移**：" + "；".join(f"`{n}` {d}" for n, d in drift[:20]) + ("…" if len(drift) > 20 else ""))
        detail.append(f"- 四層瀑布（IFW_）：{wf}；用途欄（IFC_）：{uc}。")
        detail.append("")
    out = "\n".join(lines + summary + [""] + detail)
    if a.out:
        Path(a.out).parent.mkdir(parents=True, exist_ok=True)
        Path(a.out).write_text(out, encoding="utf-8")
        print(f"written {a.out}")
    else:
        print(out)


if __name__ == "__main__":
    main()
